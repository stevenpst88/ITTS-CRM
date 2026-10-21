'use strict';
/**
 * 簽核流程設計器 M2 — 純函式（不碰資料庫／req／時間），只依賴 node 內建 crypto、./quoteApproval、./quoteFlow。
 * 規格：scratchpad/flowdesigner_SPEC.md 第 5.4、6、7.1 章。M2 範圍＝草稿的「淨化／驗證／模擬／邊界掃描／總覽圖資料」。
 * **草稿不能生效**：本檔不寫任何資料，也不被送簽主路徑引用；生效、版本治理、雙人覆核是 M3。
 *
 * 硬性限制
 *  1. 不新增 tier；node／approver／locked 形狀沿用 lib/quoteFlow.js 的 NODE_KIND（由 validateFlowDef 強制）。
 *  2. 保護底線（規格 Q5＝B）：折後未稅金額 > 1,000 萬 → 最終關卡必含 chairman 或 board；> 5,000 萬 → 必含 board。
 *     證明方式＝「區間切分取樣」：金額軸取所有金額字面值（含 1,000 萬／5,000 萬）的 −0.01／剛好／+0.01 與相鄰斷點中點；
 *     毛利軸取決策表門檻與毛利率字面值兩側的整數分；類別、BU、指定業務、商品類別、二級主管（無／有／例外名單內）各取「字面值＋一個不在任何清單內的哨兵值」。
 *     條件只會以這些變數的「比較結果」影響路線，故各軸的斷點兩側取樣即涵蓋所有可能的真值組合（取樣數過大時標示 truncated）。
 *  3. 輸入一律當作不可信：先 sanitizeDef（大小、深度、未知鍵、型別），再 QF.validateFlowDef；金額／毛利率以十進位字串＋BigInt 精算。
 *  4. 模擬與掃描一律呼叫 QF.evaluate／QF.plan（與正式引擎同一個函式），不重寫評估邏輯。
 *  5. 回應不含帳號名單（只給人數）；不含任何真實報價單的成本或毛利。
 */
const crypto = require('crypto');
const QA = require('./quoteApproval');
const QF = require('./quoteFlow');

const LIMITS = Object.freeze({ defBytes: 65536, drafts: 50, name: 40, note: 200, routeName: 16, modName: 24, nodeLabel: 14, sampleCap: 60000, sweepRows: 120, depth: 16, maxLiterals: 64 });
const COLORS = Object.freeze(['teal', 'blue', 'pink', 'amber', 'violet', 'green', 'slate']);
const BUS = Object.freeze(['ERP', 'ITS', 'MDM', 'CRM', '資安']);
const SENTINEL_BU = '__其他BU__';
const SENTINEL_OWNER = '__其他業務__';
const RANK = Object.freeze({ mgr2: 1, mgr1: 2, gm: 3, chairman: 4, board: 5 });
const FLOOR = Object.freeze({ chairmanCents: Number(QA.AMOUNT.chairman) * 100, boardCents: Number(QA.AMOUNT.board) * 100 });
const MAXC = QA.MAX_CENTS_NUM;
const CAT_LABEL = Object.freeze(QA.ROWS.reduce((m, r) => { m[r.key] = r.label; return m; }, {}));

const has = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
const isObj = (x) => x !== null && typeof x === 'object' && !Array.isArray(x);

/** QUOTE_FLOW_DESIGNER：只有字串 '1'（前後空白不拘）才開啟；其餘一律關閉（fail-safe，預設關）。 */
function designerEnabled(raw) { return String(raw === undefined || raw === null ? '' : raw).trim() === '1'; }

// ───────────────────────── 金額／文字工具 ─────────────────────────
const comma = (s) => String(s).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
/** 分 → 白話金額：整萬顯示「N 萬」，其他顯示「N 元」（含兩位小數時補齊） */
function fmtCents(c) {
  const n = BigInt(c);
  const yuan = n / 100n, cents = n % 100n;
  if (cents === 0n && yuan % 10000n === 0n && yuan >= 10000n) return comma((yuan / 10000n).toString()) + ' 萬';
  return comma(yuan.toString()) + (cents ? '.' + String(cents).padStart(2, '0') : '') + ' 元';
}
/** 分 → 十進位元字串（無千分位，至多 2 位小數）；供 API 回傳 */
function centsToYuanStr(c) {
  const n = BigInt(c);
  const yuan = n / 100n, cents = n % 100n;
  return yuan.toString() + (cents ? '.' + String(cents).padStart(2, '0') : '');
}
const ceilDiv = (a, b) => (a >= 0n ? (a + b - 1n) / b : -((-a) / b));
const noCtl = (s) => !/[\u0000-\u001f\u007f]/.test(s);
const clipText = (s, n) => String(s == null ? '' : s).replace(/[\u0000-\u001f\u007f]/g, ' ').slice(0, n);

// ───────────────────────── 淨化輸入（不信任前端）─────────────────────────
const bad = (code, message, at) => ({ ok: false, code, message, at: at || '' });

/** JSON 值的最大深度（迭代，不遞迴，避免被深層巢狀炸掉堆疊） */
function depthOf(root, limit) {
  let max = 0;
  const stack = [[root, 1]];
  let visited = 0;
  while (stack.length) {
    const [x, d] = stack.pop();
    if (d > max) max = d;
    if (max > limit) return max;
    if (++visited > 200000) return limit + 1;
    if (x && typeof x === 'object') {
      const vals = Array.isArray(x) ? x : Object.keys(x).map((k) => x[k]);
      for (const v of vals) if (v && typeof v === 'object') stack.push([v, d + 1]);
    }
  }
  return max;
}
function onlyKeys(o, allowed, at) {
  for (const k of Object.keys(o)) if (!allowed.includes(k)) return bad('DEF_UNKNOWN_KEY', '不認得的欄位：' + clipText(k, 30), at);
  return null;
}

/**
 * 把前端送來的流程定義淨化成「一份全新的純 JSON 物件」：大小／深度上限、未知欄位一律拒絕（含 __proto__／constructor）。
 * 結構與值域的完整驗證交給 QF.validateFlowDef；這裡只負責擋住不是流程定義的東西。
 */
function sanitizeDef(raw) {
  if (!isObj(raw)) return bad('DEF_NOT_OBJECT', '流程定義必須是物件');
  if (depthOf(raw, LIMITS.depth) > LIMITS.depth) return bad('DEF_TOO_DEEP', '流程定義巢狀過深（上限 ' + LIMITS.depth + ' 層）');   // 先量深度（迭代），再序列化：避免超深物件讓 JSON.stringify 爆堆疊
  let json;
  try { json = JSON.stringify(raw); } catch (_) { return bad('DEF_NOT_JSON', '流程定義無法序列化'); }
  if (typeof json !== 'string') return bad('DEF_NOT_JSON', '流程定義無法序列化');
  if (Buffer.byteLength(json, 'utf8') > LIMITS.defBytes) return bad('DEF_TOO_LARGE', '流程定義過大（上限 ' + LIMITS.defBytes + ' bytes）');
  const o = JSON.parse(json);
  let e = onlyKeys(o, ['schema', 'id', 'no', 'name', 'origin', 'limits', 'semantics', 'tables', 'nodes', 'routes', 'modifiers'], '');
  if (e) return e;
  if (o.limits !== undefined) { if (!isObj(o.limits)) return bad('DEF_BAD_SHAPE', 'limits 必須是物件', 'limits'); e = onlyKeys(o.limits, ['maxSteps'], 'limits'); if (e) return e; }
  if (o.semantics !== undefined) { if (!isObj(o.semantics)) return bad('DEF_BAD_SHAPE', 'semantics 必須是物件', 'semantics'); e = onlyKeys(o.semantics, ['sameApprover'], 'semantics'); if (e) return e; }
  if (o.tables !== undefined) {
    if (!isObj(o.tables)) return bad('DEF_BAD_SHAPE', 'tables 必須是物件', 'tables');
    e = onlyKeys(o.tables, ['margin'], 'tables'); if (e) return e;
    if (Array.isArray(o.tables.margin)) for (let i = 0; i < o.tables.margin.length; i++) {
      const r = o.tables.margin[i];
      if (!isObj(r)) return bad('DEF_BAD_SHAPE', '決策表列必須是物件', 'tables.margin[' + i + ']');
      e = onlyKeys(r, ['cat', 'l1', 'l2', 'fixedLevel'], 'tables.margin[' + i + ']'); if (e) return e;
    }
  }
  const lists = [['nodes', ['id', 'tier', 'type', 'label', 'approver', 'reassignable', 'locked', 'color']], ['routes', ['id', 'name', 'when', 'path']], ['modifiers', ['id', 'name', 'op', 'target', 'node', 'when']]];
  for (const [key, allowed] of lists) {
    if (o[key] === undefined) continue;
    if (!Array.isArray(o[key])) return bad('DEF_BAD_SHAPE', key + ' 必須是陣列', key);
    for (let i = 0; i < o[key].length; i++) {
      const it = o[key][i];
      if (!isObj(it)) return bad('DEF_BAD_SHAPE', key + ' 的項目必須是物件', key + '[' + i + ']');
      e = onlyKeys(it, allowed, key + '[' + i + ']'); if (e) return e;
    }
  }
  return { ok: true, def: o };
}

/** 深凍結的複本（QF.evaluate 對「整棵已凍結」的定義才快取驗證結果；取樣幾萬次必須走快取） */
function frozenCopy(def) {
  const c = JSON.parse(JSON.stringify(def));
  const fr = (x) => { if (x && typeof x === 'object' && !Object.isFrozen(x)) { Object.freeze(x); Object.keys(x).forEach((k) => fr(x[k])); } return x; };
  return fr(c);
}

// ───────────────────────── 字面值蒐集（取樣軸）─────────────────────────
function newUse() {
  return {
    amounts: new Set(), pcts: [], cats: new Set(), bus: new Set(), owners: new Set(), keys: new Set(), skip: new Set(),
    uses: { amount: false, level: false, pct: false, category: false, classKeys: false, bu: false, owner: false, mgr2: false },
  };
}
function collectUse(def, use) {
  use = use || newUse();
  const walk = (x, depth) => {
    if (x === true || x === false || !isObj(x) || depth > 12) return;
    const k = Object.keys(x)[0];
    const v = x[k];
    if (k === 'all' || k === 'any') { if (Array.isArray(v)) v.forEach((c) => walk(c, depth + 1)); return; }
    if (k === 'not') { walk(v, depth + 1); return; }
    if (k === 'resolvable') { use.uses.mgr2 = true; return; }
    if (k === 'resolvedIn') { use.uses.mgr2 = true; if (Array.isArray(v) && Array.isArray(v[1])) v[1].forEach((u) => { if (typeof u === 'string') use.skip.add(u); }); return; }
    if (k !== 'cmp' || !Array.isArray(v) || v.length !== 3) return;
    const [name, , rhs] = v;
    try {
      switch (name) {
        case 'amount': use.uses.amount = true; use.amounts.add(Number(QF.parseMoneyCents(rhs))); break;
        case 'marginPct': { use.uses.pct = true; const p = QF.parsePct(rhs); use.pcts.push({ n: p.n, s: p.s }); break; }
        case 'marginLevel': use.uses.level = true; break;
        case 'category': use.uses.category = true; [].concat(rhs).forEach((c) => { if (typeof c === 'string') use.cats.add(c); }); break;
        case 'classKeys': use.uses.classKeys = true; if (typeof rhs === 'string') use.keys.add(rhs); break;
        case 'ownerBuPrimary': use.uses.bu = true; [].concat(rhs).forEach((c) => { if (typeof c === 'string') use.bus.add(c); }); break;
        case 'owner': use.uses.owner = true; [].concat(rhs).forEach((c) => { if (typeof c === 'string') use.owners.add(c); }); break;
        default: break;
      }
    } catch (_) { /* 字面值無效：validateFlowDef 會報，取樣略過這一項 */ }
  };
  (def.routes || []).forEach((r) => walk(r && r.when, 0));
  (def.modifiers || []).forEach((m) => walk(m && m.when, 0));
  return use;
}

/** 取樣：回傳 { list, truncated, axes }。list 的每一筆＝ { owner, ownerBu, category, classKeys, revenueCents, gpCents, mgr2 } */
function buildSamples(defs) {
  const use = newUse();
  defs.forEach((d) => collectUse(d, use));
  const U = use.uses;
  let axisCut = false;   // 資安審查 M5：任何軸的字面值被截去（超過上限）＝取樣不完整，視同 truncated（保護底線 fail-closed）
  // 金額軸
  const amt = new Set([1, 100, FLOOR.chairmanCents - 1, FLOOR.chairmanCents, FLOOR.chairmanCents + 1, FLOOR.boardCents - 1, FLOOR.boardCents, FLOOR.boardCents + 1]);
  let lits = [...use.amounts].filter((x) => Number.isSafeInteger(x) && x >= 0);
  if (lits.length > LIMITS.maxLiterals) { axisCut = true; lits = lits.slice(0, LIMITS.maxLiterals); }
  lits.forEach((L) => { [L - 1, L, L + 1].forEach((x) => { if (x >= 1 && x <= MAXC) amt.add(x); }); });
  let sorted = [...amt].sort((a, b) => a - b);
  const mids = [];
  for (let i = 1; i < sorted.length; i++) if (sorted[i - 1] + 1 < sorted[i]) mids.push(Math.floor((sorted[i - 1] + sorted[i]) / 2));
  const top = Math.min(MAXC, Math.max(sorted[sorted.length - 1] * 2, FLOOR.boardCents * 2));
  const amounts = [...new Set(sorted.concat(mids, [top, MAXC]))].sort((a, b) => a - b);
  // 類別軸
  const cats = (U.category || U.level) ? QA.ROWS.map((r) => r.key) : ['consult'];
  // 商品類別軸（字面值的所有子集；全部 7 個鍵最多 128 組，過大會被標為 truncated 並 fail-closed）
  let keySets = [[]];
  if (U.classKeys) {
    const ks = [...use.keys].filter((k) => QA.CLASS_KEYS.includes(k));   // 全部商品類別鍵（過去只取前 6 個，第 7 個鍵永遠取不到）
    keySets = [];
    for (let m = 0; m < (1 << ks.length); m++) keySets.push(ks.filter((_, i) => m & (1 << i)));
  }
  // BU／業務軸
  const bus = U.bu ? [...new Set([...BUS, ...use.bus, SENTINEL_BU])] : ['ERP'];
  const owners = U.owner ? [...new Set([...use.owners, SENTINEL_OWNER])] : [''];
  // 二級主管軸：none＝找不到；plain＝找得到（不在例外名單）；每個例外名單帳號各一態
  if (use.skip.size > 8) axisCut = true;
  const skipUsers = [...use.skip].slice(0, 8);
  const mgr2s = U.mgr2 ? ['none', 'plain'].concat(skipUsers) : ['plain'];
  // 毛利軸（依每一列的門檻與毛利率字面值，在兩側各取一個整數分）
  const pctPoints = [];   // 以「分子／分母」表示的百分比：p = n / 10^s
  if (U.pct && use.pcts.length > LIMITS.maxLiterals) axisCut = true;
  if (U.pct) use.pcts.slice(0, LIMITS.maxLiterals).forEach((p) => pctPoints.push({ n: p.n, d: 10n ** BigInt(p.s) }));
  const gpFor = (rev, cat) => {
    const set = new Set([0, rev]);
    const pts = pctPoints.slice();
    if (U.level) {
      const row = QA.ROW_BY_KEY[cat] || QA.ROW_BY_KEY.other;
      [row.l1, row.l2].forEach((t) => { if (t !== null && t !== undefined) pts.push({ n: BigInt(t), d: 1n }); });
    }
    pts.forEach((p) => {
      const hi = ceilDiv(BigInt(rev) * p.n, 100n * p.d);   // gp*100*d >= n*rev 的最小整數 gp
      [hi - 1n, hi].forEach((g) => { if (g >= 0n && g <= BigInt(rev) * 2n) set.add(Number(g)); });
    });
    return [...set].sort((a, b) => a - b);
  };
  const needGp = U.level || U.pct;
  // 估算筆數；過多就先丟掉中點，再截斷
  let est = amounts.length * cats.length * (needGp ? 8 : 1) * keySets.length * bus.length * owners.length * mgr2s.length;
  let amountsUse = amounts;
  let truncated = false;
  if (est > LIMITS.sampleCap) { amountsUse = sorted.concat([top]); est = amountsUse.length * cats.length * (needGp ? 8 : 1) * keySets.length * bus.length * owners.length * mgr2s.length; }
  const list = [];
  outer:
  for (const rev of amountsUse) {
    for (const cat of cats) {
      const gps = needGp ? gpFor(rev, cat) : [0];
      for (const gp of gps) for (const ks of keySets) for (const bu of bus) for (const ow of owners) for (const m2 of mgr2s) {
        if (list.length >= LIMITS.sampleCap) { truncated = true; break outer; }
        list.push({ owner: ow, ownerBu: bu, category: cat, classKeys: ks, revenueCents: rev, gpCents: gp, mgr2: m2 });
      }
    }
  }
  if (axisCut) truncated = true;
  return { list, truncated, axes: { amounts: amountsUse.length, cats: cats.length, bus: bus.length, mgr2: mgr2s.length, owners: owners.length, keySets: keySets.length } };
}

/** 單筆評估（QF.evaluate；二級主管／名冊都用佔位帳號——這裡只看路線與關卡，不看真人） */
function evalSample(frozen, s, ioOver) {
  const io = ioOver || {
    mgr2: () => (s.mgr2 === 'none' ? null : s.mgr2 === 'plain' ? '__sim_mgr2' : s.mgr2),
    mgr1: () => '__sim_mgr1',
    roster: { gm: ['__sim_gm'], chairman: ['__sim_ch'] },
    boardProxy: () => ['__sim_bp'],
  };
  try {
    const ev = QF.evaluate(frozen, { owner: s.owner, ownerBu: s.ownerBu, category: s.category, classKeys: s.classKeys, revenueCents: s.revenueCents, gpCents: s.gpCents }, io);
    return { ok: true, ev };
  } catch (e) {
    return { ok: false, err: e };
  }
}

const rankOf = (tiers) => tiers.reduce((m, t) => Math.max(m, RANK[t] || 0), 0);
/** 新版是否比舊版寬鬆：最高權限階變低，或少了舊版的某一必經關卡（董事長改成董事會代核不算少） */
function isLoose(oldT, newT) {
  if (rankOf(newT) < rankOf(oldT)) return true;
  return oldT.some((x) => !newT.includes(x) && !(x === 'chairman' && newT.includes('board')));
}
const tiersText = (def, tiers) => (tiers.length ? tiers.map((t) => labelOfTier(def, t)).join(' → ') : '（無任何關卡）');
function labelOfTier(def, tier) {
  const n = (def.nodes || []).find((x) => x && x.tier === tier);
  return (n && n.label) || QA.TIERS[tier] || tier;
}

function floorNeed(rev, tiers) {
  if (rev > FLOOR.chairmanCents && !(tiers.includes('chairman') || tiers.includes('board'))) return { need: '董事長（或董事會代核）', limit: '金額超過 1,000 萬' };
  if (rev > FLOOR.boardCents && !tiers.includes('board')) return { need: '董事會代核', limit: '金額超過 5,000 萬' };
  return null;
}

/**
 * 一次掃完所有取樣：保護底線、各路線被選中次數、調整被套用次數、評估例外（FLOW_INVALID_RESULT…）、不同的最終關卡序列、與目前生效版的放寬樣本。
 * activeDef 可省略（不做放寬偵測）。
 */
function analyze(def, activeDef) {
  const frozen = frozenCopy(def);
  const frozenActive = activeDef ? frozenCopy(activeDef) : null;
  const sm = buildSamples(frozenActive ? [frozen, frozenActive] : [frozen]);
  const routeWins = Object.create(null);
  (frozen.routes || []).forEach((r) => { routeWins[r.id] = 0; });
  const modApplied = Object.create(null);
  (frozen.modifiers || []).forEach((m) => { modApplied[m.id] = 0; });
  const seqs = new Map();
  const errors = new Map();
  const violations = [];
  let violationCount = 0;
  const loose = new Map();
  let looseCount = 0;
  for (const s of sm.list) {
    const r = evalSample(frozen, s);
    if (!r.ok) {
      const c = (r.err && r.err.code) || 'ENGINE_ERROR';
      const cur = errors.get(c) || { count: 0, first: s, message: String(r.err && r.err.message || '') };
      cur.count++;
      errors.set(c, cur);
      continue;
    }
    const ev = r.ev;
    routeWins[ev.routeId] = (routeWins[ev.routeId] || 0) + 1;
    ev.applied.forEach((id) => { modApplied[id] = (modApplied[id] || 0) + 1; });
    const key = ev.actualTiers.join(',');
    const sq = seqs.get(key) || { tiers: ev.actualTiers.slice(), count: 0 };
    sq.count++;
    seqs.set(key, sq);
    const fn = floorNeed(s.revenueCents, ev.actualTiers);
    if (fn) {
      violationCount++;
      if (violations.length < 3) violations.push({ revenueCents: s.revenueCents, gpCents: s.gpCents, category: s.category, ownerBu: s.ownerBu, classKeys: s.classKeys.slice(), owner: s.owner, tiers: ev.actualTiers.slice(), need: fn.need, limit: fn.limit, mgr2: s.mgr2 });
    }
    if (frozenActive) {
      const ra = evalSample(frozenActive, s);
      if (ra.ok && isLoose(ra.ev.actualTiers, ev.actualTiers)) {
        looseCount++;
        const k = ra.ev.actualTiers.join(',') + '|' + key;
        const g = loose.get(k) || { oldT: ra.ev.actualTiers.slice(), newT: ev.actualTiers.slice(), count: 0, samples: [] };
        g.count++;
        if (g.samples.length < 400) g.samples.push(s);
        loose.set(k, g);
      }
    }
  }
  return { frozen, samples: sm.list.length, truncated: sm.truncated, axes: sm.axes, routeWins, modApplied, seqs, errors, violations, violationCount, loose, looseCount };
}

// ───────────────────────── 11 條驗證（規格 7.1）─────────────────────────
const STATUS_ORDER = { pass: 0, warn: 1, error: 2 };
function mkCheck(n, code, title, issues, passText) {
  const errs = issues.filter((i) => i.level !== 'warn');
  const warns = issues.filter((i) => i.level === 'warn');
  const status = errs.length ? 'error' : (warns.length ? 'warn' : 'pass');
  const summary = issues.length ? (errs.length ? errs.length + ' 個錯誤' : '') + (errs.length && warns.length ? '、' : '') + (warns.length ? warns.length + ' 個警告' : '') : passText;
  return { n, code, title, status, summary, issues: issues.slice(0, 30) };
}
const issue = (text, level, loc, code) => ({ text: clipText(text, 300), level: level || 'error', loc: loc || null, code: code || '' });

function locOf(def, at) {
  if (!def || typeof at !== 'string') return null;
  let m = /^routes\[(\d+)\]/.exec(at);
  if (m && Array.isArray(def.routes) && def.routes[+m[1]] && typeof def.routes[+m[1]].id === 'string') return { lane: def.routes[+m[1]].id };
  m = /^modifiers\[(\d+)\]/.exec(at);
  if (m && Array.isArray(def.modifiers) && def.modifiers[+m[1]] && typeof def.modifiers[+m[1]].id === 'string') return { mod: def.modifiers[+m[1]].id };
  m = /^nodes\[(\d+)\]/.exec(at);
  if (m && Array.isArray(def.nodes) && def.nodes[+m[1]] && typeof def.nodes[+m[1]].tier === 'string') return { node: def.nodes[+m[1]].tier };
  return null;
}
const routeName = (r) => (r && typeof r.name === 'string' && r.name) || (r && r.id) || '';

/** 把 QF.validateFlowDef 的錯誤碼分到規格 7.1 的條目 */
const CODE_TO_CHECK = { FLOW_NO_DEFAULT: 1, FLOW_EMPTY_PATH: 1, FLOW_DUP_NODE: 2, FLOW_TOO_LONG: 2, FLOW_BAD_TERMINAL: 2, FLOW_NODE_MISSING: 2 };

/**
 * env（全部可省略，省略＝該項檢查以「無資料」處理）：
 *   signerState(un)   → 'ok'|'readonly'|'inactive'|'missing'|'unqualified'
 *   roster            → { gm:[], chairman:[], boardProxy:[] }（boardProxy 已含秘書角色帳號）
 *   chain             → { owners, noMgr1, noMgr2 }（業務主管鏈解析報告，只有人數）
 *   activeDef         → 目前生效的流程（M2＝由現行設定轉出的預設流程），做放寬偵測
 * 回傳 { ok, checks[11], blocking[], defHash, samples, truncated, floor:{ok,violations}, loosening:[…] }
 * ok＝內容類檢查（1,2,3,6,10,11）沒有 error；名冊類（4,5）的 error 只回報、不擋草稿儲存（那是環境狀態，不是草稿內容）。
 */
function runChecks(rawDef, env) {
  env = env || {};
  const def = rawDef;
  const checks = [];
  const vr = QF.validateFlowDef(def);
  const byCheck = { 1: [], 2: [], 3: [], 6: [] };
  vr.errors.forEach((e) => {
    let n = CODE_TO_CHECK[e.code] || 6;
    if (e.code === 'FLOW_BAD_COND' && /applied 只能引用排在前面/.test(e.message)) n = 3;
    byCheck[n].push(issue(e.message, 'error', locOf(def, e.at), e.code));
  });
  const structural = vr.ok;
  const used = new Set();
  let a = null;
  if (structural) {
    def.routes.forEach((r) => r.path.forEach((id) => used.add(id)));
    def.modifiers.forEach((m) => { used.add(m.node); });
    a = analyze(def, env.activeDef || null);
  }
  const tierOfId = {};
  if (structural) def.nodes.forEach((n) => { tierOfId[n.id] = n.tier; });

  // 1／2：路線完整
  if (structural) {
    def.routes.forEach((r, i) => {
      for (let k = 1; k < r.path.length; k++) {
        if (RANK[tierOfId[r.path[k]]] < RANK[tierOfId[r.path[k - 1]]]) {
          byCheck[2].push(issue(`路線「${routeName(r)}」順序不合理：「${labelOfTier(def, tierOfId[r.path[k]])}」排在較高階的「${labelOfTier(def, tierOfId[r.path[k - 1]])}」之後（通常由低階簽到高階）`, 'warn', { lane: r.id, idx: k }, 'FLOW_ORDER'));
          break;
        }
      }
    });
    a.errors.forEach((v, code) => {
      // 找出是哪一條路線（把調整規則拿掉重評一次，只看路線選擇）
      let loc = null;
      try { const bare = frozenCopy(Object.assign({}, def, { modifiers: [] })); const rr = evalSample(bare, v.first); if (rr.ok) loc = { lane: rr.ev.routeId }; } catch (_) { /* 略 */ }
      byCheck[2].push(issue(`套用調整後關卡無效（${code}）：例如金額 ${fmtCents(v.first.revenueCents)}、${v.first.ownerBu}（共 ${v.count} 組取樣）—— 送簽會被直接擋下。${v.message}`, 'error', loc, code));
    });
  }
  checks.push(mkCheck(1, 'FLOW_NO_DEFAULT／FLOW_EMPTY_PATH', '至少一關：有路線、預設路線在最後且不是空的', byCheck[1], '通過：預設路線在最後，且至少有一關'));
  checks.push(mkCheck(2, 'FLOW_DUP_NODE／FLOW_TOO_LONG／FLOW_BAD_TERMINAL', '路線完整：關卡不重複、不超過上限、順序合理', byCheck[2], '通過：每條路線關卡不重複、順序由低階到高階'));
  checks.push(mkCheck(3, 'FLOW_CYCLE（結構保證）', '無環保證', byCheck[3], '通過：路線＝「依序關卡」，沒有「跳到某節點」的指令，結構上不可能成環；調整規則只能引用排在前面的調整'));

  // 4／5：簽核人（名冊）
  const e4 = [], e5 = [], w4 = [];
  if (structural) {
    const roster = env.roster || { gm: [], chairman: [], boardProxy: [] };
    const state = typeof env.signerState === 'function' ? env.signerState : () => 'ok';
    const rosterTierOf = { gm: roster.gm, chairman: roster.chairman, board: roster.boardProxy };
    def.nodes.forEach((n) => {
      if (!used.has(n.id)) return;
      const where = def.routes.filter((r) => r.path.includes(n.id));
      const loc = where.length ? { lane: where[0].id, idx: where[0].path.indexOf(n.id) } : { node: n.tier };
      if (rosterTierOf[n.tier]) {
        const list = rosterTierOf[n.tier] || [];
        const states = list.map((un) => state(un));
        const okN = states.filter((s) => s === 'ok').length;
        const alive = states.filter((s) => s === 'ok' || s === 'readonly' || s === 'unqualified').length;
        if (!alive) e4.push(issue(`「${n.label}」名單中沒有任何在職簽核人（名冊 ${list.length} 位，皆停用或不存在）—— 走到這一關的單送簽時會被擋：NO_ELIGIBLE_SIGNER`, 'error', loc, 'FLOW_NO_SIGNER'));
        else if (!okN) e5.push(issue(`「${n.label}」名單全是唯讀者或無簽核資格者（${alive} 位），沒有人能簽`, 'error', loc, 'FLOW_ONLY_READONLY'));
      } else if (env.chain && (n.tier === 'mgr1' || n.tier === 'mgr2')) {
        const c = env.chain;
        const miss = n.tier === 'mgr1' ? c.noMgr1 : c.noMgr2;
        if (miss > 0) w4.push(issue(`「${n.label}」依業務的 supervisor 欄位解析：目前有 ${miss} 位業務（共 ${c.owners} 位）解析不到人${n.tier === 'mgr2' ? '（二級主管找不到時依現行規則靜默略過，不會擋單）' : '，這些業務送簽時會被 NO_MANAGER 擋下'}`, 'warn', loc, 'FLOW_CHAIN_UNRESOLVED'));
      }
    });
  }
  checks.push(mkCheck(4, 'FLOW_NO_SIGNER', '簽核人在職且有權限：每一關至少有人能簽', e4.concat(w4), '通過：路線用到的每一關都有候選人'));
  checks.push(mkCheck(5, 'FLOW_ONLY_READONLY', '不得只有唯讀者：候選人不能全是唯讀帳號', e5, '通過：每關都有非唯讀的候選人'));

  // 6：條件語法
  const i6 = byCheck[6];
  if (structural) def.modifiers.forEach((m, i) => {
    const bu = (m.when && Array.isArray(m.when.all) && m.when.all[0] && m.when.all[0].cmp);
    if (bu && bu[0] === 'ownerBuPrimary' && bu[1] === 'in' && Array.isArray(bu[2]) && bu[2].length === 0) i6.push(issue(`調整規則「${m.name || m.id}」沒有選任何 BU，永遠不會插入`, 'warn', { mod: m.id }, 'FLOW_BAD_COND'));
  });
  checks.push(mkCheck(6, 'FLOW_BAD_COND', '條件語法：變數在白名單、運算子與值合法', i6, '通過：所有條件只用白名單變數，值在合法範圍'));

  // 7：可達性
  const i7 = [];
  if (structural) {
    def.routes.forEach((r) => { if (!a.routeWins[r.id]) i7.push(issue(`路線「${routeName(r)}」在 ${a.samples.toLocaleString('en-US')} 組邊界取樣中從來沒被選中（可能被前面的路線吃掉）`, 'warn', { lane: r.id }, 'FLOW_UNREACHABLE_RULE')); });
    def.modifiers.forEach((m) => {
      const nobu = m.when && Array.isArray(m.when.all) && m.when.all[0] && m.when.all[0].cmp && Array.isArray(m.when.all[0].cmp[2]) && m.when.all[0].cmp[2].length === 0;
      if (!a.modApplied[m.id] && !nobu) i7.push(issue(`調整規則「${m.name || m.id}」在取樣中從來沒被套用`, 'warn', { mod: m.id }, 'FLOW_UNREACHABLE_RULE'));
    });
  }
  checks.push(mkCheck(7, 'FLOW_UNREACHABLE_RULE', '可達性：不存在永遠到不了的路線', i7, structural ? `通過：每條路線至少被選中一次（取樣 ${a.samples.toLocaleString('en-US')} 組）` : '通過'));

  // 8：同人多關可行性（用真實名冊候選人）
  const i8 = [];
  if (structural && env.roster) {
    const state = typeof env.signerState === 'function' ? env.signerState : () => 'ok';
    const okOf = (list) => (list || []).filter((un) => state(un) === 'ok');
    for (const sq of a.seqs.values()) {
      const cands = sq.tiers.map((t) => (t === 'gm' ? okOf(env.roster.gm) : t === 'chairman' ? okOf(env.roster.chairman) : t === 'board' ? okOf(env.roster.boardProxy) : ['__sim_' + t]));
      if (cands.some((c) => !c.length)) continue;   // 另由 4／5 報
      if (!QF.distinctAssignment(cands)) i8.push(issue(`關卡序列「${tiersText(def, sq.tiers)}」找不到「每關不同人」的指派方式（名冊重疊，送簽時被擋：SIGNER_OVERLAP）`, 'warn', null, 'SIGNER_OVERLAP'));
    }
  }
  checks.push(mkCheck(8, 'SIGNER_OVERLAP', '可行性：每一關能指派給不同的人', i8, '通過：每種最終關卡序列都能排出「每關不同人」'));

  // 9：放寬偵測
  const looseOut = [];
  if (structural && env.activeDef) {
    const groups = [...a.loose.values()].map((g) => {
      const nice = g.samples.filter((s) => s.revenueCents % 100000000 === 0);
      const L = nice.length ? nice : g.samples;
      return { g, s: L[Math.floor(L.length / 2)] };
    }).sort((x, y) => y.s.revenueCents - x.s.revenueCents).slice(0, 8);
    groups.forEach(({ g, s }) => {
      let lv = '';
      try { lv = '、毛利層級 ' + QF.marginLevelOf(def.tables.margin, s.category, BigInt(s.gpCents), BigInt(s.revenueCents)); } catch (_) { /* 略 */ }
      looseOut.push({
        text: `金額 ${fmtCents(s.revenueCents)}${lv}、${CAT_LABEL[s.category] || s.category}、${s.ownerBu}：原需 ${tiersText(env.activeDef, g.oldT)}；新版只需 ${tiersText(def, g.newT)}`,
        count: g.count, oldTiers: g.oldT, newTiers: g.newT,
      });
    });
  }
  checks.push(mkCheck(9, 'FLOW_LOOSENING', '放寬偵測：和目前生效版比，有沒有讓控管變鬆', looseOut.map((l) => issue(l.text + (l.count > 1 ? `（類似樣本 ${l.count} 組）` : ''), 'warn', null, 'FLOW_LOOSENING')), '通過：沒有任何取樣比目前生效版更寬鬆'));

  // 10：保護底線
  const i10 = [];
  if (structural) a.violations.forEach((v) => i10.push(issue(`${v.limit}（${fmtCents(v.revenueCents)}、${CAT_LABEL[v.category] || v.category}、${v.ownerBu}）的單最終關卡是「${tiersText(def, v.tiers)}」，沒有經過${v.need}`, 'error', null, 'FLOW_FLOOR_VIOLATION')));
  if (structural && a.violationCount > a.violations.length) i10.push(issue(`另有 ${a.violationCount - a.violations.length} 組取樣同樣違反底線`, 'error', null, 'FLOW_FLOOR_VIOLATION'));
  // 資安審查 M5：取樣被截斷（含任一軸字面值被截去）→ 不能證明底線成立 → 一律判 error（fail-closed），不再顯示「通過」
  if (structural && a.truncated) i10.push(issue('取樣空間過大或條件字面值過多，部分取樣被截斷，無法證明「金額 >1,000 萬必含董事長、>5,000 萬必含董事會」成立；請簡化條件（減少路線／調整規則內的金額、帳號、毛利率字面值）後再驗證', 'error', null, 'FLOW_FLOOR_VIOLATION'));
  const c10 = mkCheck(10, 'FLOW_FLOOR_VIOLATION', '保護底線：金額 >1,000 萬必含董事長、>5,000 萬必含董事會', i10, structural ? `通過：${a.samples.toLocaleString('en-US')} 組邊界取樣都符合底線${a.truncated ? '（取樣已達上限被截斷）' : ''}` : '通過');
  checks.push(c10);

  // 11：標籤與鎖定節點
  const i11 = [];
  const badLabel = (s) => typeof s !== 'string' || !s.trim() || s.length > LIMITS.nodeLabel || !noCtl(s) || /[<>"'&]/.test(s);
  if (Array.isArray(def.nodes)) def.nodes.forEach((n, i) => {
    if (!n || typeof n !== 'object') return;
    if (badLabel(n.label)) i11.push(issue(`「${QA.TIERS[n.tier] || n.tier}」的顯示名稱不合法（1–${LIMITS.nodeLabel} 字、不可含 < > " ' &、不可含控制字元）`, 'error', locOf(def, 'nodes[' + i + ']'), 'FLOW_BAD_LABEL'));
    if (n.tier === 'board' && n.label !== QA.TIERS.board) i11.push(issue('鎖定節點「董事會代核」的名稱被改動（代核／秘書／名冊的語意待董事長溝通，設計器不能改它）', 'error', locOf(def, 'nodes[' + i + ']'), 'FLOW_LOCKED_EDIT'));
    if (n.color !== undefined && !COLORS.includes(n.color)) i11.push(issue(`「${QA.TIERS[n.tier] || n.tier}」的節點色不在允許清單內`, 'error', locOf(def, 'nodes[' + i + ']'), 'FLOW_BAD_LABEL'));
    if (n.tier === 'board' && n.color !== undefined) i11.push(issue('鎖定節點「董事會代核」不可更改節點色', 'error', locOf(def, 'nodes[' + i + ']'), 'FLOW_LOCKED_EDIT'));
  });
  const nameBad = (s, max) => typeof s !== 'string' || !s.trim() || s.length > max || !noCtl(s);
  if (Array.isArray(def.routes)) def.routes.forEach((r, i) => { if (r && nameBad(r.name, LIMITS.routeName)) i11.push(issue(`第 ${i + 1} 條路線的名稱不合法（必填、至多 ${LIMITS.routeName} 字、不可含控制字元）`, 'error', locOf(def, 'routes[' + i + ']'), 'FLOW_BAD_LABEL')); });
  if (Array.isArray(def.modifiers)) def.modifiers.forEach((m, i) => { if (m && nameBad(m.name, LIMITS.modName)) i11.push(issue(`第 ${i + 1} 個調整規則的名稱不合法（必填、至多 ${LIMITS.modName} 字、不可含控制字元）`, 'error', locOf(def, 'modifiers[' + i + ']'), 'FLOW_BAD_LABEL')); });
  checks.push(mkCheck(11, 'FLOW_LOCKED_EDIT', '鎖定節點與標籤：代核關的屬性不可變更、名稱合法', i11, '通過：代核關未被改動；所有名稱合法（顯示一律以純文字／跳脫輸出）'));

  const BLOCKING = [1, 2, 3, 6, 10, 11];
  const blocking = checks.filter((c) => BLOCKING.includes(c.n) && c.status === 'error');
  return {
    ok: structural && blocking.length === 0,
    structural,
    checks,
    blocking: blocking.map((c) => ({ n: c.n, code: c.code, issues: c.issues.filter((i) => i.level !== 'warn').map((i) => i.text) })),
    defHash: structural ? QF.defHash(def) : null,
    samples: a ? a.samples : 0,
    routeWins: a ? a.routeWins : {},
    truncated: a ? a.truncated : false,
    floor: { ok: structural ? (a.violationCount === 0 && !a.truncated) : null, violations: structural ? a.violations : [] },
    loosening: looseOut,
  };
}

// ───────────────────────── 白話條件（給畫布／追蹤／總覽用）─────────────────────────
const OP_ZH = { '==': '等於', '!=': '不等於', '>': '大於', '>=': '大於等於', '<': '小於', '<=': '小於等於', in: '屬於', not_in: '不屬於', contains: '包含', not_contains: '不包含' };
const VAR_ZH = { amount: '折後未稅金額', marginLevel: '毛利層級', marginPct: '毛利率', category: '商品歸類', classKeys: '商品類別', ownerBuPrimary: '業務所屬 BU', owner: '指定業務' };
/** 單一比較的白話（與前端 itemText 同一套寫法；scripts/check-quote-flow-designer-ui.js 會比對兩邊） */
function describeCmp(name, op, rhs) {
  const one = (v) => {
    if (name === 'amount') { try { return fmtCents(QF.parseMoneyCents(v)); } catch (_) { return String(v); } }
    if (name === 'marginLevel') return '層級 ' + v;
    if (name === 'marginPct') return String(v) + '%';
    if (name === 'category') return CAT_LABEL[v] || String(v);
    if (name === 'classKeys') return QA.CLASS_LABELS[v] || String(v);
    return String(v);
  };
  const val = Array.isArray(rhs) ? (rhs.length ? '[' + rhs.map(one).join('、') + ']' : '[（無）]') : one(rhs);
  return (VAR_ZH[name] || String(name)) + ' ' + (OP_ZH[op] || String(op)) + ' ' + val;
}
/** 路線的「何時走這條」：預設路線(true)顯示成「其他情況（預設）」 */
const whenZh = (w, names) => (w === true ? '其他情況（預設）' : describeWhen(w, names));
function describeWhen(x, names, depth) {
  depth = depth || 0;
  names = names || {};
  const nm = (id) => (has(names, id) ? names[id] : String(id));
  if (depth > 10) return '…';
  if (x === true) return '一律';
  if (x === false) return '永不';
  if (!isObj(x)) return '（無效條件）';
  const k = Object.keys(x)[0];
  const v = x[k];
  switch (k) {
    case 'all': return v.map((c) => describeWhen(c, names, depth + 1)).join('，且');
    case 'any': return '（' + v.map((c) => describeWhen(c, names, depth + 1)).join('，或') + '）';
    case 'not': return '並非（' + describeWhen(v, names, depth + 1) + '）';
    case 'cmp': return describeCmp(v[0], v[1], v[2]);
    case 'applied': return '已套用「' + nm(v) + '」';
    case 'resolvable': return nm(v) + '找得到人';
    case 'resolvedIn': return nm(v[0]) + '的簽核人屬於例外名單（' + v[1].length + ' 人）';
    default: return '（未知條件）';
  }
}

// ───────────────────────── 白話檢視（給前端畫布用）─────────────────────────
function nameMap(def) {
  const names = Object.create(null);
  (def.nodes || []).forEach((n) => { if (n && typeof n.id === 'string') names[n.id] = n.label; });
  (def.modifiers || []).forEach((m) => { if (m && typeof m.id === 'string') names[m.id] = m.name || m.id; });
  return names;
}
function modKind(m) {
  const w = m && m.when;
  if (m && (m.op === 'insert_before' || m.op === 'insert_after') && w && Array.isArray(w.all) && w.all.length === 2 && w.all[0] && w.all[0].cmp && w.all[0].cmp[0] === 'ownerBuPrimary' && w.all[0].cmp[1] === 'in' && Array.isArray(w.all[0].cmp[2]) && w.all[1] && typeof w.all[1].resolvable === 'string') return { kind: 'bu', bus: w.all[0].cmp[2].slice() };
  if (m && m.op === 'remove' && w && Array.isArray(w.all) && w.all.length === 2 && w.all[0] && typeof w.all[0].applied === 'string' && w.all[1] && Array.isArray(w.all[1].resolvedIn)) return { kind: 'skip', userCount: Array.isArray(w.all[1].resolvedIn[1]) ? w.all[1].resolvedIn[1].length : 0 };
  return { kind: 'other' };
}
/** 流程定義 → 白話檢視資料（路線條件文字、調整規則類型）。def 須已通過 validateFlowDef。 */
function viewOf(def) {
  const names = nameMap(def);
  return {
    routes: def.routes.map((r) => ({ id: r.id, whenText: whenZh(r.when, names), isDefault: r.when === true, pathText: r.path.map((id) => names[id] || id).join(' → ') })),
    modifiers: def.modifiers.map((m) => Object.assign({ id: m.id, whenText: describeWhen(m.when, names) }, modKind(m))),
  };
}

// ───────────────────────── 模擬 ─────────────────────────
const SIM_AMOUNT_RE = /^(0|[1-9][0-9]{0,15})(?:\.([0-9]{1,2}))?$/;
/** 驗證並整理模擬輸入。回傳 { ok, input } 或 { ok:false, code, message } */
function parseSimInput(b, flowCfg) {
  if (!isObj(b)) return bad('SIM_BAD_INPUT', '模擬輸入必須是物件');
  const allowed = ['amount', 'marginLevel', 'marginPct', 'category', 'classKeys', 'ownerBu', 'noMgr2', 'mgr2Exception'];
  for (const k of Object.keys(b)) if (!allowed.includes(k)) return bad('SIM_BAD_INPUT', '不認得的輸入欄位：' + clipText(k, 30));
  if (typeof b.amount !== 'string' || !SIM_AMOUNT_RE.test(b.amount)) return bad('SIM_BAD_AMOUNT', '折後未稅金額必須是十進位字串（元，至多 2 位小數）');
  let revenue;
  try { revenue = QF.parseMoneyCents(b.amount); } catch (_) { return bad('SIM_BAD_AMOUNT', '折後未稅金額無效或超過上限'); }
  if (revenue <= 0n) return bad('SIM_BAD_AMOUNT', '折後未稅金額必須大於 0');
  let category = 'consult';
  if (b.classKeys !== undefined) {
    if (!Array.isArray(b.classKeys) || b.classKeys.length > 12 || !b.classKeys.every((k) => typeof k === 'string' && QA.CLASS_KEYS.includes(k))) return bad('SIM_BAD_CLASS', '商品類別必須是類別代碼的陣列');
    category = QA.resolveRow(b.classKeys).key;
  }
  if (b.category !== undefined) {
    if (typeof b.category !== 'string' || !QA.ROW_BY_KEY[b.category]) return bad('SIM_BAD_CATEGORY', '商品歸類無效');
    category = b.category;
  }
  const ownerBu = b.ownerBu === undefined ? 'ERP' : b.ownerBu;
  if (typeof ownerBu !== 'string' || !ownerBu || ownerBu.length > 40 || !noCtl(ownerBu)) return bad('SIM_BAD_BU', 'BU 無效');
  const row = QA.ROW_BY_KEY[category];
  let gp, effective;
  if (b.marginPct !== undefined && b.marginPct !== null && b.marginPct !== '') {
    let p;
    try { p = QF.parsePct(b.marginPct); } catch (_) { return bad('SIM_BAD_PCT', '毛利率必須是十進位字串（至多 4 位小數，可為負）'); }
    gp = (revenue * p.n) / (100n * (10n ** BigInt(p.s)));   // 向零截斷（分）
    effective = 'pct';
  } else {
    const lv = b.marginLevel === undefined ? 1 : b.marginLevel;
    if (![1, 2, 3].includes(lv)) return bad('SIM_BAD_LEVEL', '毛利層級只能是 1、2、3');
    if (row.fixedLevel) { gp = 0n; }
    else if (lv === 1) { if (row.l1 === null) return bad('SIM_LEVEL_IMPOSSIBLE', `「${row.label}」沒有層級 1（毛利再高也需總經理）`); gp = ceilDiv(revenue * BigInt(row.l1), 100n); }
    else if (lv === 2) { gp = ceilDiv(revenue * BigInt(row.l2), 100n); if (row.l1 !== null && gp >= ceilDiv(revenue * BigInt(row.l1), 100n)) return bad('SIM_LEVEL_IMPOSSIBLE', '金額太小，無法用整數分表達層級 2'); }
    else gp = 0n;
    effective = 'level';
  }
  if (gp > BigInt(MAXC) || gp < -BigInt(MAXC)) return bad('SIM_BAD_PCT', '毛利超出範圍');
  if ((b.noMgr2 !== undefined && typeof b.noMgr2 !== 'boolean') || (b.mgr2Exception !== undefined && typeof b.mgr2Exception !== 'boolean')) return bad('SIM_BAD_FLAG', 'noMgr2／mgr2Exception 必須是布林值');
  const noMgr2 = b.noMgr2 === true;
  const exc = b.mgr2Exception === true;
  const skipUsers = QF.legacyInputs(flowCfg).skipUsers;
  return {
    ok: true,
    input: { revenueCents: Number(revenue), gpCents: Number(gp), category, classKeys: Array.isArray(b.classKeys) ? b.classKeys.slice() : [], ownerBu, noMgr2, mgr2Exception: exc && !noMgr2 && skipUsers.length > 0, marginMode: effective, marginLevelAsked: b.marginLevel, marginPctAsked: b.marginPct === undefined ? null : b.marginPct, skipUser: skipUsers[0] || null },
  };
}

function simIo(input, env) {
  const roster = (env && env.roster) || { gm: [], chairman: [], boardProxy: [] };
  const state = env && typeof env.signerState === 'function' ? env.signerState : () => 'ok';
  const okOf = (l) => (l || []).filter((un) => state(un) === 'ok');
  const mgr2 = input.noMgr2 ? null : (input.mgr2Exception ? input.skipUser : '__sim_mgr2');
  return {
    owner: '__sim_owner', costBy: null,
    mgr2: () => mgr2, mgr1: () => '__sim_mgr1',
    roster: { gm: okOf(roster.gm), chairman: okOf(roster.chairman) },
    boardProxy: () => okOf(roster.boardProxy),
    canApprove: (un) => (typeof un === 'string' && un.startsWith('__sim_')) || state(un) === 'ok',
  };
}

function approverText(tier, input, env, def) {
  const roster = (env && env.roster) || { gm: [], chairman: [], boardProxy: [] };
  const state = env && typeof env.signerState === 'function' ? env.signerState : () => 'ok';
  const n = (l) => (l || []).filter((un) => state(un) === 'ok').length;
  if (tier === 'mgr2') return '業務的直屬二級主管（送簽時鎖定成具體的人）';
  if (tier === 'mgr1') return '主管鏈上第一位一級主管（送簽時鎖定成具體的人）';
  if (tier === 'gm') return n(roster.gm) ? `名冊「總經理」${n(roster.gm)} 位候選（任一位）` : '⚠ 名冊「總經理」沒有可簽核的人';
  if (tier === 'chairman') return n(roster.chairman) ? `名冊「董事長」${n(roster.chairman)} 位候選（任一位）` : '⚠ 名冊「董事長」沒有可簽核的人';
  return n(roster.boardProxy) ? `董事會決議由秘書／代核人代為登錄（${n(roster.boardProxy)} 位候選）` : '⚠ 找不到董事會代核人';
}

function factText(name, input, lvl) {
  switch (name) {
    case 'amount': return fmtCents(input.revenueCents);
    case 'marginLevel': return '層級 ' + lvl;
    case 'marginPct': return (input.revenueCents ? (Number(BigInt(input.gpCents) * 1000000n / BigInt(input.revenueCents)) / 10000).toFixed(2) : '0') + '%';
    case 'category': return CAT_LABEL[input.category] || input.category;
    case 'classKeys': return input.classKeys.length ? input.classKeys.map((k) => QA.CLASS_LABELS[k] || k).join('、') : '（無）';
    case 'ownerBuPrimary': return input.ownerBu;
    case 'owner': return '（模擬業務）';
    default: return '';
  }
}

/** 逐條件解釋（每個原子條件的真假）。回傳 { ok, text, atoms:[{text, ok}] } */
function explainCond(when, def, input, io, lvlFn, appliedSet, isMod) {
  const names = nameMap(def);
  const facts = QF.makeFacts({ owner: '', ownerBu: input.ownerBu, category: input.category, classKeys: input.classKeys, revenueCents: input.revenueCents, gpCents: input.gpCents });
  const nodeById = {};
  def.nodes.forEach((n) => { nodeById[n.id] = n; });
  const env = {
    facts, modifier: !!isMod, applied: appliedSet, level: lvlFn,
    resolveNode: (id) => {
      const n = nodeById[id];
      if (!n) return { person: null, persons: [] };
      if (n.tier === 'mgr2') return { person: io.mgr2() || null, persons: [] };
      if (n.tier === 'mgr1') return { person: io.mgr1() || null, persons: [] };
      if (n.tier === 'gm') return { person: null, persons: io.roster.gm };
      if (n.tier === 'chairman') return { person: null, persons: io.roster.chairman };
      return { person: null, persons: io.boardProxy() };
    },
  };
  const atoms = [];
  const leaf = (x) => {
    let ok = false;
    try { ok = QF.evalCond(x, env); } catch (_) { ok = false; }
    let t = describeWhen(x, names);
    if (x && x.cmp) t += `（本單：${factText(x.cmp[0], input, lvlFn())}）`;
    atoms.push({ text: t, ok });
    return ok;
  };
  const walk = (x) => {
    if (x === true) return true;
    if (!isObj(x)) return false;
    const k = Object.keys(x)[0];
    if (k === 'all') { let r = true; for (const c of x.all) { const v = walk(c); if (!v) { r = false; } } return r; }
    if (k === 'any') { let r = false; for (const c of x.any) { if (walk(c)) r = true; } return r; }
    if (k === 'not') return !walk(x.not);
    return leaf(x);
  };
  const ok = walk(when);
  return { ok, text: describeWhen(when, names), atoms };
}

/**
 * 模擬一張假想單：與正式引擎同一個 QF.plan。def 須已通過 validateFlowDef。
 * 回傳 { route, steps[], trace[], applied[], blockers[], marginLevel, text… }；compare 時另有 against（目前生效版）與差異判斷。
 */
function simulate(def, parsed, env, activeDef) {
  const input = parsed.input;
  const frozen = frozenCopy(def);
  const io = simIo(input, env);
  const facts = { owner: io.owner, ownerBu: input.ownerBu, category: input.category, classKeys: input.classKeys, revenueCents: input.revenueCents, gpCents: input.gpCents };
  const p = QF.plan(frozen, facts, io);
  const lvl = () => QF.marginLevelOf(frozen.tables.margin, input.category, BigInt(input.gpCents), BigInt(input.revenueCents));
  const names = nameMap(frozen);
  const nodeByTier = {};
  frozen.nodes.forEach((n) => { nodeByTier[n.tier] = n; });
  const routeIdx = frozen.routes.findIndex((r) => r.id === p.routeId);
  const route = frozen.routes[routeIdx];
  const basePath = route.path.map((id) => (frozen.nodes.find((n) => n.id === id) || {}).tier);
  const steps = p.actualTiers.map((t, i) => ({
    step: i + 1, tier: t, nodeId: nodeByTier[t].id, label: nodeByTier[t].label, inserted: !basePath.includes(t), locked: !!nodeByTier[t].locked,
    approver: approverText(t, input, env, frozen),
  }));
  // 追蹤：路線（命中前逐條，命中後標「不再檢查」）＋調整（依序，累計 applied）
  const trace = [];
  let hit = false;
  frozen.routes.forEach((r, i) => {
    if (hit) { trace.push({ kind: 'route', id: r.id, name: routeName(r), order: i + 1, evaluated: false, matched: false, atoms: [], text: whenZh(r.when, names) }); return; }
    const ex = explainCond(r.when, frozen, input, io, lvl, new Set(), false);
    const matched = r.id === p.routeId;
    if (matched) hit = true;
    trace.push({ kind: 'route', id: r.id, name: routeName(r), order: i + 1, evaluated: true, matched, isDefault: r.when === true, atoms: ex.atoms, text: ex.text });
  });
  const appliedSet = new Set();
  frozen.modifiers.forEach((m) => {
    const ex = explainCond(m.when, frozen, input, io, lvl, new Set(appliedSet), true);
    const did = p.applied.includes(m.id);
    if (did) appliedSet.add(m.id);
    trace.push({ kind: 'modifier', id: m.id, name: m.name || m.id, evaluated: true, matched: ex.ok, applied: did, atoms: ex.atoms, text: ex.text, op: m.op, target: names[m.target] || m.target, node: m.node ? (names[m.node] || m.node) : null });
  });
  const level = lvl();
  const out = {
    ok: true,
    input: { amount: centsToYuanStr(input.revenueCents), category: input.category, categoryLabel: CAT_LABEL[input.category] || input.category, classKeys: input.classKeys, ownerBu: input.ownerBu, noMgr2: input.noMgr2, mgr2Exception: input.mgr2Exception, marginMode: input.marginMode, marginLevel: level, marginPct: factText('marginPct', input, level) },
    route: { id: p.routeId, name: routeName(route), order: routeIdx + 1, isDefault: route.when === true },
    steps, applied: p.applied.map((id) => ({ id, name: names[id] || id })),
    trace, blockers: p.blockers, tiers: p.actualTiers.slice(), financialTiers: p.financialTiers.slice(),
    defHash: p.defHash,
  };
  if (activeDef) {
    const fa = frozenCopy(activeDef);
    const pa = QF.plan(fa, facts, io);
    const same = pa.actualTiers.join(',') === p.actualTiers.join(',');
    out.against = {
      route: { id: pa.routeId, name: routeName(fa.routes.find((r) => r.id === pa.routeId)) },
      steps: pa.actualTiers.map((t, i) => ({ step: i + 1, tier: t, label: labelOfTier(fa, t), removedInDraft: !p.actualTiers.includes(t) })),
      same, loose: !same && isLoose(pa.actualTiers, p.actualTiers),
      addedInDraft: p.actualTiers.filter((t) => !pa.actualTiers.includes(t)),
    };
    out.steps.forEach((s) => { s.addedVsActive = !pa.actualTiers.includes(s.tier); });
  }
  const fn = floorNeed(input.revenueCents, p.actualTiers);
  out.floor = fn ? { ok: false, need: fn.need, limit: fn.limit } : { ok: true };
  return out;
}

// ───────────────────────── 邊界掃描 ─────────────────────────
/**
 * 邊界掃描：以 base（模擬輸入，金額除外）固定其他欄位，對「所有金額字面值 ±0.01、1,000 萬／5,000 萬 ±0.01」逐列模擬並與生效版比較；
 * 另對完整取樣集合（金額×類別×毛利×BU×二級主管…）統計：各路線被選中次數、永遠不會生效的規則、放寬／不同／違反底線的組數、無法解析簽核人的組數。
 */
function sweep(def, baseBody, env, activeDef, flowCfg) {
  const useAmt = new Set([FLOOR.chairmanCents, FLOOR.boardCents]);
  const u = collectUse(def);
  const ua = activeDef ? collectUse(activeDef) : newUse();
  [...u.amounts, ...ua.amounts].forEach((x) => { if (Number.isSafeInteger(x)) useAmt.add(x); });
  const amounts = [];
  [...useAmt].sort((a, b) => a - b).forEach((L) => { [L - 1, L, L + 1].forEach((x) => { if (x >= 1 && x <= MAXC && !amounts.includes(x)) amounts.push(x); }); });
  const limited = amounts.slice(0, LIMITS.sweepRows);
  const rows = [];
  const frozen = frozenCopy(def);
  const fa = activeDef ? frozenCopy(activeDef) : null;
  for (const a of limited) {
    const body = Object.assign({}, baseBody, { amount: centsToYuanStr(a) });
    const pi = parseSimInput(body, flowCfg);
    if (!pi.ok) return pi;
    const io = simIo(pi.input, env);
    const facts = { owner: io.owner, ownerBu: pi.input.ownerBu, category: pi.input.category, classKeys: pi.input.classKeys, revenueCents: pi.input.revenueCents, gpCents: pi.input.gpCents };
    const p = QF.plan(frozen, facts, io);
    const pa = fa ? QF.plan(fa, facts, io) : null;
    const same = pa ? pa.actualTiers.join(',') === p.actualTiers.join(',') : null;
    const fn = floorNeed(a, p.actualTiers);
    rows.push({
      amount: centsToYuanStr(a), amountText: fmtCents(a),
      route: routeName(frozen.routes.find((r) => r.id === p.routeId)), tiers: p.actualTiers.slice(), stepsText: tiersText(frozen, p.actualTiers),
      activeStepsText: pa ? tiersText(fa, pa.actualTiers) : null, same, loose: pa ? (!same && isLoose(pa.actualTiers, p.actualTiers)) : null,
      floorBad: !!fn, blockers: p.blockers.map((b) => b.code),
    });
  }
  const an = analyze(def, activeDef || null);
  const unreachableRoutes = def.routes.filter((r) => !an.routeWins[r.id]).map((r) => ({ id: r.id, name: routeName(r) }));
  const unreachableMods = def.modifiers.filter((m) => !an.modApplied[m.id]).map((m) => ({ id: m.id, name: m.name || m.id }));
  // 全取樣集合：不同／放寬／無法解析簽核人
  const sm = buildSamples(fa ? [frozen, fa] : [frozen]);
  let differing = 0, blocked = 0;
  const blockCodes = Object.create(null);
  // 「不同／無法解析簽核人」的統計逐筆跑 QF.plan（較貴）：取樣超過 12,000 組時等距抽樣，耗時才有上限（路線被選中次數／放寬／底線仍用 analyze 的完整取樣）
  const stride = Math.max(1, Math.ceil(sm.list.length / 12000));
  for (let si = 0; si < sm.list.length; si += stride) {
    const s = sm.list[si];
    const sio = simIo({ noMgr2: s.mgr2 === 'none', mgr2Exception: s.mgr2 !== 'none' && s.mgr2 !== 'plain', skipUser: s.mgr2 }, env);
    const facts = { owner: sio.owner, ownerBu: s.ownerBu, category: s.category, classKeys: s.classKeys, revenueCents: s.revenueCents, gpCents: s.gpCents };
    let p;
    try { p = QF.plan(frozen, facts, sio); } catch (_) { continue; }
    if (p.blockers.length) { blocked++; p.blockers.forEach((b) => { blockCodes[b.code] = (blockCodes[b.code] || 0) + 1; }); }
    if (fa) {
      try { const pa = QF.plan(fa, facts, sio); if (pa.actualTiers.join(',') !== p.actualTiers.join(',')) differing++; } catch (_) { /* 略 */ }
    }
  }
  return {
    ok: true, rows, rowsTruncated: amounts.length > limited.length,
    summary: {
      samples: sm.list.length, truncated: sm.truncated, sampledEvery: stride, routeWins: an.routeWins, unreachableRoutes, unreachableMods,
      differing: fa ? differing : null, loosened: fa ? an.looseCount : null, floorViolations: an.violationCount, blockedSamples: blocked, blockCodes,
      engineErrors: [...an.errors.keys()],
    },
  };
}

// ───────────────────────── 總覽圖資料（唯讀自由畫布）─────────────────────────
const GEO = Object.freeze({ approval: { w: 176, h: 64 }, decision: { w: 250, h: 104 }, start: { w: 84, h: 40 }, end: { w: 84, h: 40 }, mod: { w: 360, h: 74 } });
const X0 = 48, COLGAP = 56, ROWGAP = 56;

/**
 * 由流程定義產生「節點／連線／自動排版座標」。排版為確定性的分層版面：
 *   最左 開始 → 每條非預設路線一個判斷菱形（上到下，「不符合」往下一條）→ 「符合」往右接該路線的關卡序列 → 最右 結束；預設路線在最下方一列；
 *   調整規則放在最下方的虛線備註卡（不連線，避免與每條路線交錯），受影響的關卡標 ⚙。
 * sim（可省略）＝ simulate() 的結果：命中路線改畫成「實際關卡序列」（含調整插入的關卡），並標記命中的節點與連線。
 */
function buildOverview(def, sim) {
  const names = nameMap(def);
  const tierOfId = {};
  def.nodes.forEach((n) => { tierOfId[n.id] = n; });
  const nodes = [];
  const edges = [];
  const hitRoute = sim && sim.route ? sim.route.id : null;
  const hitIdx = hitRoute ? def.routes.findIndex((r) => r.id === hitRoute) : -1;
  const seqFor = (r, i) => {
    if (i === hitIdx && sim) return sim.tiers.map((t) => def.nodes.find((n) => n.tier === t));
    return r.path.map((id) => tierOfId[id]);
  };
  const seqs = def.routes.map(seqFor);
  const maxLen = Math.max(1, ...seqs.map((s) => s.length));
  const decisionCount = def.routes.length - 1;
  const xDec = X0 + GEO.start.w + COLGAP + GEO.decision.w / 2;
  const xPath0 = xDec + GEO.decision.w / 2 + COLGAP + GEO.approval.w / 2;
  const stepW = GEO.approval.w + COLGAP;
  const xEnd = xPath0 + (maxLen - 1) * stepW + GEO.approval.w / 2 + COLGAP + GEO.end.w / 2;
  const rowH = GEO.decision.h + ROWGAP;
  const yOf = (i) => 40 + GEO.decision.h / 2 + i * rowH;
  const modTargets = {};
  def.modifiers.forEach((m) => { if (m.target) { (modTargets[m.target] = modTargets[m.target] || []).push(m.id); } if (m.node) { (modTargets[m.node] = modTargets[m.node] || []).push(m.id); } });
  const add = (n) => { nodes.push(n); return n; };
  const edge = (from, to, label, kind, hit) => edges.push({ id: 'e' + edges.length, from, to, label: label || '', kind: kind || 'flow', hit: !!hit });
  add({ id: 'start', kind: 'start', label: '送簽', x: X0 + GEO.start.w / 2, y: yOf(0), w: GEO.start.w, h: GEO.start.h, hit: !!sim });
  add({ id: 'end', kind: 'end', label: '核准完成', x: xEnd, y: yOf(Math.max(0, def.routes.length - 1)) , w: GEO.end.w + 16, h: GEO.end.h, hit: !!sim });
  def.routes.forEach((r, i) => {
    const y = yOf(i);
    const isDef = r.when === true;
    const hit = i === hitIdx;
    const passed = sim && hitIdx >= 0 && i <= hitIdx;   // 判斷鏈上走過的菱形
    if (!isDef) {
      add({ id: 'd' + i, kind: 'decision', label: `${i + 1}. ${routeName(r)}`, sub: whenZh(r.when, names), x: xDec, y, w: GEO.decision.w, h: GEO.decision.h, routeId: r.id, order: i + 1, hit: !!passed });
    }
    const seq = seqs[i];
    seq.forEach((n, k) => {
      const inserted = hit && sim && !r.path.includes(n.id);
      add({
        id: `r${i}n${k}`, kind: 'approval', tier: n.tier, label: n.label, sub: `STEP ${k + 1}${inserted ? '・調整插入' : ''}`, x: xPath0 + k * stepW, y, w: GEO.approval.w, h: GEO.approval.h,
        locked: !!n.locked, color: n.color || null, routeId: r.id, hit: hit && !!sim, inserted: !!inserted, mods: (modTargets[n.id] || []).slice(),
      });
    });
    // 連線
    const first = `r${i}n0`;
    if (!isDef) {
      edge('d' + i, first, '符合', 'yes', hit);
      if (i < def.routes.length - 2) edge('d' + i, 'd' + (i + 1), '不符合', 'no', sim && hitIdx > i);
      else edge('d' + i, `r${def.routes.length - 1}n0`, '其他情況', 'no', sim && hitIdx > i);
    }
    for (let k = 0; k + 1 < seq.length; k++) edge(`r${i}n${k}`, `r${i}n${k + 1}`, '核准', 'flow', hit && !!sim);
    edge(`r${i}n${seq.length - 1}`, 'end', '核准', 'flow', hit && !!sim);
  });
  edge('start', def.routes.length > 1 ? 'd0' : 'r0n0', '', 'flow', !!sim);
  // 調整規則備註卡
  const bottom = yOf(def.routes.length - 1) + GEO.decision.h / 2 + 60;
  def.modifiers.forEach((m, j) => {
    const hit = !!(sim && sim.applied && sim.applied.some((a) => a.id === m.id));
    const kind = modKind(m);
    add({
      id: 'm' + j, kind: 'mod', label: `⚙ ${m.name || m.id}`, sub: describeWhen(m.when, names) + '　→　' + ({ insert_before: '在「' + (names[m.target] || m.target) + '」前插入「' + (names[m.node] || m.node) + '」', insert_after: '在「' + (names[m.target] || m.target) + '」後插入「' + (names[m.node] || m.node) + '」', remove: '移除「' + (names[m.target] || m.target) + '」', replace: '把「' + (names[m.target] || m.target) + '」換成「' + (names[m.node] || m.node) + '」' }[m.op] || ''),
      x: X0 + GEO.mod.w / 2 + (j % 2) * (GEO.mod.w + 24), y: bottom + Math.floor(j / 2) * (GEO.mod.h + 20) + GEO.mod.h / 2, w: GEO.mod.w, h: GEO.mod.h, modId: m.id, hit, modKind: kind.kind,
    });
  });
  const minX = Math.min(...nodes.map((n) => n.x - n.w / 2)), maxX = Math.max(...nodes.map((n) => n.x + n.w / 2));
  const minY = Math.min(...nodes.map((n) => n.y - n.h / 2)), maxY = Math.max(...nodes.map((n) => n.y + n.h / 2));
  return { nodes, edges, bounds: { x0: minX - 24, y0: minY - 24, x1: maxX + 24, y1: maxY + 24 }, hitRoute, outline: def.routes.map((r, i) => ({ order: i + 1, name: routeName(r), whenText: whenZh(r.when, names), stepsText: seqs[i].map((n) => n.label).join(' → '), hit: i === hitIdx })) };
}

// ───────────────────────── 草稿記錄 ─────────────────────────
const newDraftId = () => 'fd_' + crypto.randomBytes(6).toString('hex');
const ID_RE = /^fd_[0-9a-f]{12}$/;
function cleanName(s, max) {
  if (typeof s !== 'string') return null;
  const t = s.trim();
  if (!t || t.length > max || !noCtl(t)) return null;
  return t;
}
function summaryOf(rec) {
  return { id: rec.id, name: rec.name, note: rec.note || '', baseVer: rec.baseVer || '', createdAt: rec.createdAt, createdBy: rec.createdBy, updatedAt: rec.updatedAt, updatedBy: rec.updatedBy, routes: rec.def && rec.def.routes ? rec.def.routes.length : 0, defHash: rec.def ? QF.defHash(rec.def).slice(0, 12) : '' };
}

module.exports = {
  LIMITS, COLORS, BUS, FLOOR, RANK, ID_RE,
  designerEnabled, fmtCents, centsToYuanStr, clipText,
  sanitizeDef, frozenCopy, depthOf, buildSamples, evalSample, analyze, isLoose, floorNeed,
  runChecks, viewOf, modKind, parseSimInput, simulate, sweep, buildOverview, collectUse, describeWhen, describeCmp,
  newDraftId, cleanName, summaryOf, tiersText,
};
