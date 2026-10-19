#!/usr/bin/env node
/**
 * 簽核流程引擎（M1）等價檢查。用法：node scripts/check-quote-flow-engine.js（純函式＋假 app＋記憶體資料庫；不開伺服器、不碰 data.json／auth.json／audit.log.json）
 *
 * Oracle＝現有程式本身：quoteApproval.requiredPath（財務層）＋ quoteRoutes 內的 buildActualTiers／findManager1／chainBlockers（實際關卡、簽核人、阻擋）。
 * 被測＝lib/quoteFlow.js 的「預設流程（由核決表／門檻／路徑表／BU 設定轉出）＋評估引擎」。兩邊逐項（財務層、實際關卡、mgr2、mgr1、阻擋碼與訊息）完全相等才算過。
 *
 *   A 釘選與結構    DEFAULT_FLOW_HASH 釘選、預設流程通過驗證、tier 只有現有五種、原始碼沒有 eval／new Function
 *   B 等價網格      B1 財務網格（11 個歸類 × 22 個金額 × 毛利邊界）× 代表性世界；B2 BU 開關 32 組合 × 主要 BU 5 種 × 例外；
 *                   B3 主管結構／BU 格式／例外名單的完整乘積；B4 名冊／秘書／成本填寫人的完整乘積（SIGNER_OVERLAP 等）；
 *                   B5 固定種子 100,000 組隨機（種子印在報告）；整批輸出的雜湊與 GOLDEN 比對（回歸快照）
 *   C 路由層        shadow 零差異且不改寫任何資料；on 的結果與 off 逐項相同（僅多 approval.flow）；凍結快照語意（開關事後改變、設定事後改變，在途單照 steps 走）；
 *                   故障注入：引擎例外／結果不一致 → shadow 只稽核、on 自動退回舊路徑；稽核節流；環境變數解析；flow 不外洩
 *   D 條件 AST 安全 惡意輸入、深度／寬度炸彈、循環參照、超大數字、型別錯誤、原型污染、凍結不可變
 *   E 變異測試      把引擎改壞（> ↔ >=、少一個 BU、漏 skip_mgr1、金額用浮點、略過去重…），上面的等價核心必須變紅
 *   F off 對基底    QFE_BASE_ROOT=<git archive 出來的基底目錄> 時，同一組請求在基底與本工作樹（off）逐位元比對回應、資料、稽核、通知
 * 環境變數：QFE_QUICK=1（縮小網格，除錯用；不比對 GOLDEN）、QFE_BASE_ROOT（F 區）、QFE_SEED（預設 20261020）
 * 測試資料只用通用字串。
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const Module = require('module');

const ROOT = path.join(__dirname, '..');
const QF_PATH = path.join(ROOT, 'lib/quoteFlow.js');
const QA = require(path.join(ROOT, 'lib/quoteApproval.js'));
const QF = require(QF_PATH);
const registerQuoteRoutes = require(path.join(ROOT, 'lib/quoteRoutes.js'));

const QUICK = process.env.QFE_QUICK === '1';
const SKIP = new Set((process.env.QFE_SKIP || '').split(',').filter(Boolean));   // 除錯用：QFE_SKIP=B,E 略過等價網格／變異測試
const SEED = parseInt(process.env.QFE_SEED || '20261020', 10);
const BASE_ROOT = process.env.QFE_BASE_ROOT || '';
// 回歸快照（QFE_SEED 預設值、非 QUICK 時比對）：整批「引擎輸出」的 sha256。引擎、預設流程或網格定義被改動而結果不同時才會變，變了就必須說明。
const GOLDEN = '1057d46a143449adb320e008e4cbf99b437f7bd858c13a0674563354dd4efc3c';
const PINNED_DEFAULT_FLOW_HASH = '520fc5a5d553240ea068eb23bba3ec4712b163d7db76301b342166fb99284194';

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);
const ALL_BUS = ['ERP', 'ITS', 'MDM', 'CRM', '資安'];
const TIER5 = ['mgr2', 'mgr1', 'gm', 'chairman', 'board'];

function canon(v) {
  if (Array.isArray(v)) return '[' + v.map(canon).join(',') + ']';
  if (v && typeof v === 'object') return '{' + Object.keys(v).sort().map((k) => JSON.stringify(k) + ':' + canon(v[k])).join(',') + '}';
  return JSON.stringify(v === undefined ? null : v);
}
// 固定種子亂數（mulberry32）
function rng(seed) {
  let a = seed >>> 0;
  return () => { a = (a + 0x6D2B79F5) >>> 0; let x = Math.imul(a ^ (a >>> 15), 1 | a); x = (x + Math.imul(x ^ (x >>> 7), 61 | x)) ^ x; return ((x ^ (x >>> 14)) >>> 0) / 4294967296; };
}

// 與 server.js 同一份 normalizeBu（從原始碼抽出來，不手抄）
const serverSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
const nbSrc = /function normalizeBu\(bu\) \{[\s\S]*?\n\}\n/.exec(serverSrc);
if (!nbSrc) { console.error('讀不到 server.js 的 normalizeBu'); process.exit(2); }
const normalizeBu = vm.runInNewContext(nbSrc[0] + '\n;normalizeBu');

// ───────────────────────── 取得 oracle（舊路徑函式）─────────────────────────
function mkInternal(root) {
  const reg = require(path.join(root, 'lib/quoteRoutes.js'));
  const app = new Proxy({}, { get: () => () => {} });
  const logs = [];
  const deps = {
    db: { load: () => ({}), save() {}, flush: async () => {} }, loadAuth: () => ({ users: [] }), saveAuth() {},
    requireAuth: (a, b, n) => n(), requireAdmin: (a, b, n) => n(), writeLog: (...a) => logs.push(a), pushNotification() {},
    getViewableOwners: () => [], sanitizeStr: (s) => String(s == null ? '' : s), genQuoteNo: () => 'QT', taipeiToday: () => '2026-10-20', resolveIssuer: () => ({}),
    buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: async () => Buffer.from(''), QUOTE_TEMPLATE: '', uuidv4: () => 'x', normalizeBu, getUserFeatures: () => ['quotations'],
  };
  const r = reg(app, deps);
  return Object.assign({ logs }, r._internal);
}
const I = mkInternal(ROOT);

// ───────────────────────── 世界（使用者／名冊／BU 設定）─────────────────────────
const U = (username, role, extra) => Object.assign({ username, role, active: true, displayName: username }, extra || {});
function baseUsers() {
  return {
    o: U('o', 'user', { bu: ['ERP'] }), m2: U('m2', 'manager2'), m1: U('m1', 'manager1'),
    g1: U('g1', 'executive'), g2: U('g2', 'executive'), c1: U('c1', 'executive'), c2: U('c2', 'executive'),
    b1: U('b1', 'user'), cons: U('cons', 'user'), adm: U('adm', 'admin'),
    gI: U('gI', 'executive', { active: false }), gRO: U('gRO', 'executive', { quoteMode: 'readonly' }), gV: U('gV', 'executive', { accessMode: 'view' }), cI: U('cI', 'executive', { active: false }),
    s1: U('s1', 'user'), s2: U('s2', 'user'), sRO: U('sRO', 'user', { quoteMode: 'readonly' }), sI: U('sI', 'user', { active: false }),
  };
}
const SUPK = ['m2ok', 'm2inactive', 'm2view', 'm2ro', 'm2wrongrole', 'toM1', 'none', 'dangling', 'loop', 'm2pool'];
function applySup(k, u) {
  const o = u.o, m2 = u.m2;
  switch (k) {
    case 'm2ok': o.supervisor = 'm2'; m2.supervisor = 'm1'; break;
    case 'm2inactive': o.supervisor = 'm2'; m2.supervisor = 'm1'; m2.active = false; break;
    case 'm2view': o.supervisor = 'm2'; m2.supervisor = 'm1'; m2.accessMode = 'view'; break;
    case 'm2ro': o.supervisor = 'm2'; m2.supervisor = 'm1'; m2.quoteMode = 'readonly'; break;
    case 'm2wrongrole': o.supervisor = 'm2'; m2.supervisor = 'm1'; m2.role = 'user'; break;
    case 'toM1': o.supervisor = 'm1'; break;
    case 'none': break;
    case 'dangling': o.supervisor = 'ghost'; break;
    case 'loop': o.supervisor = 'm2'; m2.supervisor = 'o'; break;
    case 'm2pool': o.supervisor = 'm2'; m2.supervisor = 'm1'; m2.role = 'pool'; break;
    default: throw new Error('sup ' + k);
  }
}
const M1K = ['ok', 'inactive', 'ro', 'absent', 'view'];
function applyM1(k, u) {
  if (k === 'inactive') u.m1.active = false;
  else if (k === 'ro') u.m1.quoteMode = 'readonly';
  else if (k === 'absent') delete u.m1;
  else if (k === 'view') u.m1.accessMode = 'view';
}
const OBU = [['ERP'], ['ITS'], ['MDM'], ['CRM'], ['資安'], 'ERP', '資安', ['資安', 'ERP'], [], undefined, 'FOO', [123]];
const EXCK = [
  [], [{ username: 'm2', action: 'skip_mgr1' }], [{ username: 'm2', action: 'other' }], [{ username: 'x', action: 'skip_mgr1' }],
  [{ username: 'm2', action: 'skip_mgr1' }, { username: 'm2', action: 'skip_mgr1' }],
  [{ username: '', action: 'skip_mgr1' }, null, { username: 'm2' }, { username: 'm2', action: 'skip_mgr1' }],
];
const GMK = [[], ['g1'], ['g1', 'g2'], ['o'], ['m1'], ['m2'], ['ghost'], ['gI'], ['gRO'], ['gI', 'g1'], ['g1', 'c1'], ['adm'], ['gV', 'g2']];
const CHK = [[], ['c1'], ['c1', 'c2'], ['o'], ['g1'], ['cons'], ['cI'], ['g1', 'c1']];
const BPK = [[], ['b1'], ['s1'], ['gI'], ['g1'], ['o']];
const SECK = [[], ['s1'], ['s1', 's2'], ['sRO'], ['sI']];
const CBK = [undefined, 'cons', 'g1', 'c1', 'm1', 'b1', 'o'];
const BASE_SPEC = { sup: 'm2ok', m1: 'ok', ownerBu: ['ERP'], bus: ['ERP'], exc: [], gm: ['g1'], ch: ['c1'], bp: ['b1'], secs: ['s1'], costBy: undefined, noFlowCfg: false };

function mkWorld(sp0) {
  const sp = Object.assign({}, BASE_SPEC, sp0 || {});
  const users = baseUsers();
  if (sp.ownerBu === undefined) delete users.o.bu; else users.o.bu = sp.ownerBu;
  applySup(sp.sup, users);
  applyM1(sp.m1, users);
  sp.secs.forEach((n) => { users[n].role = 'secretary'; });
  const cfg = { roster: { gm: sp.gm, chairman: sp.ch, boardProxy: sp.bp, costProviders: [], sealManagers: [] }, productClasses: {} };
  const buSettings = {};
  sp.bus.forEach((b) => { buSettings[b] = { hasMgr2Step: true }; });
  const data = sp.noFlowCfg ? {} : { approvalFlowConfig: { buSettings, exceptions: sp.exc } };
  return { users, cfg, data, q: { id: 'Q', quoteNo: 'QT', owner: 'o', costBy: sp.costBy }, sp };
}

// ───────────────────────── 財務網格 ─────────────────────────
const ROWKEYS = QA.ROWS.map((r) => r.key).concat(['zzz', '__proto__']);
const REVS = [1, 3, 100, 12345, 500000000, 999999999, 1000000000, 1000000001, 3000000000, 4999999999, 5000000000, 5000000001, 100000000000, 10000000000000, 999999999999999, 1000000000000000];
const REVS_BIG = [1000000000, 5000000000, 100000000000000, 999999999999999, 1000000000000000];
function gpPoints(cat, rev) {
  const row = QA.ROW_BY_KEY[cat] && Object.prototype.hasOwnProperty.call(QA.ROW_BY_KEY, cat) ? QA.ROW_BY_KEY[cat] : QA.ROW_BY_KEY.other;
  const pts = new Set([0, -1, rev, rev + 1, Math.floor(rev / 2), -rev]);
  for (const th of [row.l1, row.l2]) {
    if (th === null || th === undefined) continue;
    const g = Number((BigInt(th) * BigInt(rev) + 99n) / 100n);   // gp*100 >= th*rev 的最小整數 gp
    [g - 1, g, g + 1].forEach((x) => pts.add(x));
  }
  return [...pts].filter((x) => Number.isSafeInteger(x));
}
function finAll() {
  const out = [];
  for (const cat of ROWKEYS) for (const rev of REVS) for (const gp of gpPoints(cat, rev)) out.push({ cat, rev, gp });
  return out;
}
const FIN_PATHS = [
  { cat: 'consult', rev: 500000000, gp: 200000000 }, { cat: 'consult', rev: 500000000, gp: 100000000 }, { cat: 'consult', rev: 500000000, gp: 25000000 },
  { cat: 'consult', rev: 2000000000, gp: 800000000 }, { cat: 'consult', rev: 6000000000, gp: 2400000000 },
];

// ───────────────────────── Oracle 與被測 ─────────────────────────
function oracleOf(w, fin) {
  const rp = QA.requiredPath(fin.cat, fin.gp, fin.rev);
  const flowCfg = I.getFlowCfg(w.data);
  const { tiers: actual, mgr2 } = I.buildActualTiers(rp.tiers, w.q.owner, w.users, flowCfg);
  const mgr1 = actual.includes('mgr1') ? I.findManager1(w.users, w.q.owner) : null;
  const blockers = I.chainBlockers(rp.tiers, w.q, { cfg: w.cfg, users: w.users, data: w.data });
  return { financialTiers: rp.tiers, actualTiers: actual, mgr2: mgr2 || null, mgr1, blockers };
}
function engineOf(QFm, w, fin) {
  const flowCfg = I.getFlowCfg(w.data);
  const def = QFm.legacyDefault(flowCfg);
  const { users, cfg } = w;
  const io = {
    owner: w.q.owner, costBy: w.q.costBy,
    mgr2: () => I.findManager2(users, w.q.owner), mgr1: () => I.findManager1(users, w.q.owner),
    canApprove: (un) => I.canApproveAct(users[un]),
    roster: { gm: cfg.roster.gm, chairman: cfg.roster.chairman }, boardProxy: () => [...I.boardProxySet(cfg, users)],
  };
  const facts = { owner: w.q.owner, ownerBu: normalizeBu((users[w.q.owner] || {}).bu)[0] || 'ERP', category: fin.cat, classKeys: [], revenueCents: fin.rev, gpCents: fin.gp };
  return QFm.plan(def, facts, io);
}
const keyOf = (p) => canon({ f: p.financialTiers, a: p.actualTiers, m2: p.mgr2 || null, m1: p.mgr1 || null, b: p.blockers });

// ───────────────────────── 等價核心（供完整檢查與變異測試共用）─────────────────────────
function core(QFm, opt) {
  const st = { total: 0, fails: 0, first: [], thrown: 0, digest: crypto.createHash('sha256'), groups: {} };
  const grp = (g) => (st.groups[g] || (st.groups[g] = { total: 0, fails: 0 }));
  function cmp(g, w, fin, label) {
    st.total++; const G = grp(g); G.total++;
    let o, e, ek;
    try { o = oracleOf(w, fin); } catch (err) { o = { err: String(err && err.message) }; }
    try { e = engineOf(QFm, w, fin); ek = keyOf(e); } catch (err) { e = null; ek = 'THROW:' + String(err && err.code || err && err.message).slice(0, 60); st.thrown++; }
    const ok = !o.err && e && ek === keyOf(o);
    st.digest.update(ek + '\n');
    if (!ok) { st.fails++; G.fails++; if (st.first.length < 5) st.first.push(g + ' ' + label + ' fin=' + J(fin) + ' sp=' + J(w.sp) + ' oracle=' + (o.err || keyOf(o)) + ' engine=' + ek); }
  }
  const every = (arr, k, off) => arr.filter((_, i) => i % k === (off || 0));
  const finA = finAll();
  const sp = (x) => mkWorld(x);

  // B1 財務網格 × 代表性世界
  const reps = [
    {}, { bus: [] }, { bus: ['ERP'], exc: EXCK[1] }, { bus: ALL_BUS, exc: EXCK[1] }, { sup: 'toM1', bus: ['ERP'] }, { gm: [], ch: [], bp: [] },
    { gm: ['g1'], ch: ['g1'] }, { bp: [], secs: [] },
  ].map(sp);
  for (const w of reps) for (const f of every(finA, opt.strideF)) cmp('B1', w, f, 'rep');

  // B2 BU 開關 × 主要 BU × 例外
  for (let mask = 0; mask < 32; mask++) {
    const bus = ALL_BUS.filter((_, i) => (mask >> i) & 1);
    for (const ob of [['ERP'], ['ITS'], ['MDM'], ['CRM'], ['資安']]) for (const exc of [[], EXCK[1]]) {
      const w = sp({ bus, ownerBu: ob, exc });
      for (const f of every(finA, opt.strideG1, (mask + ob[0].length) % opt.strideG1)) cmp('B2', w, f, 'mask=' + mask);
    }
  }

  // B3 主管結構 × BU 格式 × 例外 × BU 設定（乾淨名冊）
  const bucfg = [[], ['ERP'], ALL_BUS, ['資安', 'ITS']];
  let n3 = 0;
  for (const s of SUPK) for (const m of M1K) for (const ob of OBU) for (const bs of bucfg) for (const exc of EXCK) {
    if ((n3++ % opt.strideG2) !== 0) continue;
    const w = sp({ sup: s, m1: m, ownerBu: ob, bus: bs, exc });
    for (const f of FIN_PATHS) cmp('B3', w, f, 'struct');
  }

  // B4 名冊 × 秘書 × 成本填寫人（兩種結構）
  let n4 = 0;
  for (const base of [{ sup: 'm2ok', bus: ['ERP'] }, { sup: 'toM1', bus: [] }]) {
    for (const gm of GMK) for (const ch of CHK) for (const bp of BPK) for (const secs of SECK) for (const cb of CBK) {
      if ((n4++ % opt.strideG3) !== 0) continue;
      const w = sp(Object.assign({}, base, { gm, ch, bp, secs, costBy: cb }));
      for (const f of FIN_PATHS) cmp('B4', w, f, 'roster');
    }
  }

  // B5 隨機
  const R = rng(opt.seed);
  const pick = (a) => a[Math.floor(R() * a.length)];
  for (let i = 0; i < opt.nRandom; i++) {
    const mask = Math.floor(R() * 32);
    const w = sp({
      sup: pick(SUPK), m1: pick(M1K), ownerBu: pick(OBU), bus: ALL_BUS.filter((_, k) => (mask >> k) & 1), exc: pick(EXCK),
      gm: pick(GMK), ch: pick(CHK), bp: pick(BPK), secs: pick(SECK), costBy: pick(CBK), noFlowCfg: R() < 0.03,
    });
    const cat = pick(ROWKEYS);
    const rev = R() < 0.5 ? pick(REVS) : Math.max(1, Math.floor(Math.exp(R() * Math.log(1e15))));
    const gps = gpPoints(cat, rev);
    const gp = R() < 0.6 ? pick(gps) : Math.floor((R() * 2 - 0.5) * rev);
    cmp('B5', w, { cat, rev, gp }, 'rand#' + i);
  }
  st.digestHex = st.digest.digest('hex');
  return st;
}

// ───────────────────────── 條件 AST 安全測試 ─────────────────────────
function astTests(QFm, tt) {
  const bad = (name, ast, ctx) => { let r; try { r = QFm.validateCond(ast, ctx); } catch (e) { tt(name + '（validateCond 不得丟例外）', false, String(e && e.message)); return; } tt(name, r.ok === false && r.errors.length > 0, J(r)); };
  const good = (name, ast, ctx) => { let r; try { r = QFm.validateCond(ast, ctx); } catch (e) { tt(name, false, String(e && e.message)); return; } tt(name, r.ok === true, J(r)); };
  const nested = (n, leaf) => { let x = leaf; for (let i = 0; i < n; i++) x = { not: x }; return x; };
  good('D1.1 合法條件：金額 > 10000000', { cmp: ['amount', '>', '10000000'] });
  good('D1.2 合法：all／any／not 組合', { all: [{ any: [{ cmp: ['marginLevel', '==', 3] }, { cmp: ['category', 'in', ['consult', 'crm']] }] }, { not: { cmp: ['ownerBuPrimary', '==', 'ERP'] } }] });
  good('D1.3 合法：true／false 字面值', true); good('D1.3b 合法：false', false);
  good('D1.4 合法：classKeys contains；owner in 清單；marginPct >= 25.5', { all: [{ cmp: ['classKeys', 'contains', 'software'] }, { cmp: ['owner', 'in', ['a', 'b']] }, { cmp: ['marginPct', '>=', '25.5'] }] });
  good('D1.5 modifier 內可用 applied／resolvable／resolvedIn', { all: [{ applied: 'm1' }, { resolvable: 'n1' }, { resolvedIn: ['n1', ['u']] }] }, { modifier: true, nodeIds: new Set(['n1']), nodeTier: { n1: 'mgr2' }, modIdsBefore: new Set(['m1']) });
  bad('D2.1 路線條件不得用 applied', { applied: 'm1' });
  bad('D2.2 未知運算子 eval', { eval: 'process.exit(1)' });
  bad('D2.3 未知運算子 fn', { fn: 'x' });
  bad('D2.4 兩個運算子', { all: [true], any: [true] });
  bad('D2.5 空物件', {});
  bad('D2.6 null', null); bad('D2.7 undefined', undefined); bad('D2.8 數字', 5); bad('D2.9 字串', 'true'); bad('D2.10 陣列', [true]);
  bad('D2.11 函式', () => true);
  bad('D2.12 all 空陣列', { all: [] });
  bad('D2.13 all 非陣列', { all: 'x' });
  bad('D2.14 not 帶陣列', { not: [true] });
  bad('D3.1 變數不在白名單', { cmp: ['password', '==', 'x'] });
  bad('D3.2 變數 constructor', { cmp: ['constructor', '==', 'x'] });
  bad('D3.3 變數 __proto__', { cmp: ['__proto__', '==', 'x'] });
  bad('D3.4 變數 toString', { cmp: ['toString', '==', 'x'] });
  bad('D3.5 變數 hasOwnProperty', { cmp: ['hasOwnProperty', '==', 'x'] });
  bad('D3.6 變數為數字', { cmp: [1, '==', 1] });
  bad('D3.7 cmp 長度 2', { cmp: ['amount', '>'] });
  bad('D3.8 cmp 長度 4', { cmp: ['amount', '>', '1', '2'] });
  bad('D3.9 cmp 非陣列', { cmp: 'amount > 1' });
  bad('D3.10 運算子注入', { cmp: ['amount', '> 0 || process.exit()', '1'] });
  bad('D3.11 運算子 ===', { cmp: ['amount', '===', '1'] });
  bad('D3.12 運算子為物件', { cmp: ['amount', {}, '1'] });
  bad('D3.13 金額 in 運算子', { cmp: ['amount', 'in', ['1']] });
  bad('D3.14 category 用 > ', { cmp: ['category', '>', 'consult'] });
  bad('D3.15 classKeys 用 ==', { cmp: ['classKeys', '==', 'software'] });
  bad('D3.16 classKeys 值不在值域', { cmp: ['classKeys', 'contains', 'nope'] });
  bad('D3.17 category 值不在值域', { cmp: ['category', '==', 'nope'] });
  bad('D3.18 category in 含非法值', { cmp: ['category', 'in', ['consult', 'nope']] });
  bad('D3.19 in 的值不是陣列', { cmp: ['owner', 'in', 'abc'] });
  bad('D3.20 in 陣列含非字串', { cmp: ['owner', 'in', ['a', 1]] });
  bad('D3.21 in 陣列超過上限', { cmp: ['owner', 'in', Array.from({ length: 501 }, (_, i) => 'u' + i)] });
  bad('D3.22 字串過長', { cmp: ['owner', '==', 'x'.repeat(101)] });
  bad('D3.23 空字串', { cmp: ['owner', '==', ''] });
  for (const v of ['10000000.001', '1e9', '1,000', '-5', ' 5', '5 ', '0x10', '.5', '5.', '00', '01', '1_0', 'NaN', 'Infinity', '99999999999999999999', '10000000000000000', '１２３', '5\n', '']) bad('D4 金額字串無效：' + J(v), { cmp: ['amount', '>', v] });
  for (const v of [10000000, 1e21, NaN, Infinity, -1, null, undefined, true, {}, [], 10n]) bad('D4 金額不是字串：' + String(typeof v) + ':' + String(v), { cmp: ['amount', '>', v] });
  good('D4.x 金額邊界 MAX（9999999999999.99 元＝上限內）', { cmp: ['amount', '<=', '9999999999999.99'] });
  bad('D4.y 金額超過 MAX_CENTS（10000000000000.01 元）', { cmp: ['amount', '<=', '10000000000000.01'] });
  for (const v of ['4', '0', '-1', 1.5, NaN, Infinity, '3', true, null, [3], 2 ** 53]) bad('D5 層級值無效：' + String(typeof v) + ':' + String(v), { cmp: ['marginLevel', '==', v] });
  for (const v of ['1e2', '100.00001', '1234', ' 5', 25, 'abc', '--5']) bad('D5 毛利率值無效：' + J(v), { cmp: ['marginPct', '>=', v] });
  bad('D6.1 深度炸彈 1000 層 not', nested(1000, true));
  bad('D6.2 深度 9 層（上限 8）', nested(9, true));
  good('D6.3 深度 8 層（上限內）', nested(7, true));
  let deep = true; for (let i = 0; i < 200000; i++) deep = { not: deep };
  bad('D6.4 深度炸彈 200000 層（迭代建出，驗證不得爆堆疊）', deep);
  const wide = { all: Array.from({ length: 1000000 }, () => true) };
  bad('D6.5 寬度炸彈 1,000,000 個子條件', wide);
  const bomb = { all: Array.from({ length: 64 }, () => ({ all: Array.from({ length: 64 }, () => ({ all: Array.from({ length: 64 }, () => ({ all: Array.from({ length: 64 }, () => true) })) })) })) };
  const t0 = Date.now();
  bad('D6.6 寬×深組合爆炸（64^4 個節點）：必須快速中止', bomb);
  tt('D6.6b 組合爆炸驗證耗時 < 500ms', Date.now() - t0 < 500, (Date.now() - t0) + 'ms');
  const cyc = { not: null }; cyc.not = cyc;
  bad('D6.7 循環參照', cyc);
  const cyc2 = { all: [] }; cyc2.all.push(cyc2);
  bad('D6.8 循環參照（all 陣列）', cyc2);
  bad('D6.9 JSON.parse 出的 __proto__ 鍵', JSON.parse('{"__proto__":{"all":[true]}}'));
  bad('D6.10 constructor 鍵', JSON.parse('{"constructor":{"prototype":{"polluted":true}}}'));
  const thrower = new Proxy({}, { ownKeys() { throw new Error('boom'); }, getPrototypeOf() { return Object.prototype; } });
  bad('D6.11 Proxy 在 ownKeys 丟例外', thrower);
  bad('D6.12 帶 getter 的類別實例', new (class X { get all() { return [true]; } })());
  bad('D6.13 Symbol 當 key 的物件仍視為空物件', { [Symbol('x')]: 1 });
  bad('D6.14 Date 物件', new Date());
  bad('D6.15 Map', new Map());
  bad('D6.16 Object.create(非 Object.prototype)', Object.create({ all: [true] }));
  // 求值層：未經驗證的惡意 AST 丟進 evalCond 必須丟 FlowError（不是 RangeError/TypeError）、不執行任何東西
  const env = { facts: QFm.makeFacts({ owner: 'o', ownerBu: 'ERP', category: 'consult', classKeys: [], revenueCents: 100, gpCents: 10 }), modifier: false, applied: new Set(), level: () => 1, resolveNode: () => ({ person: null, persons: [] }) };
  const ev = (name, ast) => { let thrown = null, val; try { val = QFm.evalCond(ast, env); } catch (e) { thrown = e; } tt(name, !!thrown && thrown.name === 'FlowError' && /^FLOW_/.test(thrown.code), thrown ? thrown.name + ':' + thrown.code : 'returned ' + val); };
  ev('D7.1 evalCond 深度炸彈 → FlowError', nested(5000, true));
  ev('D7.2 evalCond 未知運算子 → FlowError', { eval: 'x' });
  ev('D7.3 evalCond 變數 constructor → FlowError', { cmp: ['constructor', '==', 'x'] });
  ev('D7.4 evalCond 循環 → FlowError', cyc);
  ev('D7.5 evalCond 路線內用 applied → FlowError', { applied: 'm' });
  ev('D7.6 evalCond 金額字串惡意 → FlowError', { cmp: ['amount', '>', '1e9'] });
  ev('D7.7 evalCond 型別錯誤（in 非陣列）→ FlowError', { cmp: ['owner', 'in', 'abc'] });
  ev('D7.8 evalCond 寬度炸彈 → FlowError', wide);
  const bombOnEval = (() => { const t1 = Date.now(); try { QFm.evalCond(bomb, env); } catch (e) { /* 預期 */ } return Date.now() - t1; })();
  tt('D7.9 evalCond 組合爆炸耗時 < 500ms', bombOnEval < 500, bombOnEval + 'ms');
  // facts 型別錯誤
  const badFacts = [null, undefined, 5, {}, { revenueCents: 1.5, gpCents: 1 }, { revenueCents: 1, gpCents: '1' }, { revenueCents: 0, gpCents: 0 }, { revenueCents: -5, gpCents: 0 }, { revenueCents: 2 ** 60, gpCents: 0 }, { revenueCents: 1, gpCents: NaN }];
  badFacts.forEach((f, i) => { let thrown = null; try { QFm.makeFacts(f); } catch (e) { thrown = e; } tt('D8.' + (i + 1) + ' makeFacts 拒絕無效 facts：' + String(J(f)), !!thrown && thrown.name === 'FlowError', thrown ? thrown.code : 'no throw'); });
  // 流程定義層
  const base = () => JSON.parse(JSON.stringify(QFm.buildLegacyDefault({ buSettings: { ERP: { hasMgr2Step: true } }, exceptions: [{ username: 'm', action: 'skip_mgr1' }] })));
  const defBad = (name, mut, code) => { const d = base(); mut(d); let r; try { r = QFm.validateFlowDef(d); } catch (e) { tt(name + '（不得丟例外）', false, String(e && e.message)); return; } tt(name, r.ok === false && (!code || r.errors.some((e) => e.code === code)), J(r.errors.slice(0, 3))); };
  tt('D9.0 預設流程（含 BU 與例外）通過驗證', QFm.validateFlowDef(base()).ok === true);
  defBad('D9.1 新增 tier "accounting" 節點被拒（M1 不新增 tier）', (d) => { d.nodes.push({ id: 'n_acc', tier: 'accounting', type: 'approval', label: 'x', approver: {}, reassignable: false, locked: false }); }, 'FLOW_BAD_NODE');
  defBad('D9.2 board 節點 locked 被改成 false', (d) => { d.nodes.find((n) => n.tier === 'board').locked = false; }, 'FLOW_BAD_NODE');
  defBad('D9.3 board 節點 approver 被改（代核語意不可動）', (d) => { d.nodes.find((n) => n.tier === 'board').approver = { kind: 'roster', name: 'gm', mode: 'any_one' }; }, 'FLOW_BAD_NODE');
  defBad('D9.4 gm 節點 approver 被改成 chairman 名冊', (d) => { d.nodes.find((n) => n.tier === 'gm').approver.name = 'chairman'; }, 'FLOW_BAD_NODE');
  defBad('D9.5 最後一條路線不是預設（when≠true）', (d) => { d.routes[d.routes.length - 1].when = { cmp: ['amount', '>', '1'] }; }, 'FLOW_NO_DEFAULT');
  defBad('D9.6 路線 path 為空', (d) => { d.routes[0].path = []; }, 'FLOW_EMPTY_PATH');
  defBad('D9.7 路線引用不存在的節點', (d) => { d.routes[0].path = ['n_nope']; }, 'FLOW_NODE_MISSING');
  defBad('D9.8 路線內節點重複', (d) => { d.routes[0].path = ['n_mgr1', 'n_mgr1']; }, 'FLOW_DUP_NODE');
  defBad('D9.9 路線過長（> maxSteps）', (d) => { d.limits.maxSteps = 2; }, 'FLOW_TOO_LONG');
  defBad('D9.10 maxSteps 超過 8', (d) => { d.limits.maxSteps = 9; }, 'FLOW_BAD_SCHEMA');
  defBad('D9.11 路線條件含惡意 AST', (d) => { d.routes[0].when = { eval: 'x' }; }, 'FLOW_BAD_COND');
  defBad('D9.12 modifier 的 applied 引用排在後面的調整', (d) => { d.modifiers[0].when = { applied: 'm_skip' }; }, 'FLOW_BAD_COND');
  defBad('D9.13 modifier 的 target 不存在', (d) => { d.modifiers[0].target = 'n_nope'; }, 'FLOW_NODE_MISSING');
  defBad('D9.14 sameApprover 改成 skip_later（M1 只支援 forbid）', (d) => { d.semantics.sameApprover = 'skip_later'; }, 'FLOW_BAD_SCHEMA');
  defBad('D9.15 決策表缺 other 列', (d) => { d.tables.margin = d.tables.margin.filter((r) => r.cat !== 'other'); }, 'FLOW_BAD_TABLE');
  defBad('D9.16 決策表 l2 為 null 且無 fixedLevel', (d) => { d.tables.margin[0].l2 = null; }, 'FLOW_BAD_TABLE');
  defBad('D9.17 決策表門檻為浮點小數', (d) => { d.tables.margin[0].l1 = 25.5; }, 'FLOW_BAD_TABLE');
  defBad('D9.18 schema 版本不符', (d) => { d.schema = 2; }, 'FLOW_BAD_SCHEMA');
  defBad('D9.19 節點 tier 重複', (d) => { d.nodes.push(Object.assign({}, d.nodes[0], { id: 'n_dup' })); }, 'FLOW_DUP_NODE');
  defBad('D9.20 resolvedIn 用在名冊型節點（gm）', (d) => { d.modifiers[1].when = { resolvedIn: ['n_gm', ['a']] }; }, 'FLOW_BAD_COND');
  let thr = null; try { QFm.validateFlowDef(null); QFm.validateFlowDef('x'); QFm.validateFlowDef([]); QFm.validateFlowDef({ nodes: 1, routes: 1, modifiers: 1, tables: 1 }); } catch (e) { thr = e; }
  tt('D9.21 validateFlowDef 對 null／字串／陣列／畸形物件不丟例外', thr === null, thr && thr.message);
  // evaluate 不修改 def、可處理凍結物件；結果無效時丟 FlowError
  const d0 = QFm.buildLegacyDefault({ buSettings: { ERP: { hasMgr2Step: true } }, exceptions: [] });
  const before = canon(d0);
  const io = { mgr2: () => 'm2', mgr1: () => 'm1', roster: { gm: ['g'], chairman: ['c'] }, boardProxy: () => ['s'] };
  let okEval = true;
  try { for (let i = 0; i < 50; i++) QFm.evaluate(d0, { owner: 'o', ownerBu: 'ERP', category: 'consult', classKeys: [], revenueCents: 1000000000 + i, gpCents: 1 }, io); } catch (e) { okEval = false; }
  tt('D10.1 evaluate 不修改（已凍結的）流程定義，且可重複執行', okEval && canon(d0) === before && Object.isFrozen(d0) && Object.isFrozen(d0.routes[0]));
  tt('D10.2 原型未被污染', ({}).polluted === undefined && ({}).all === undefined && Object.prototype.hasOwnProperty.call(Object.prototype, 'polluted') === false);
  let unfrozenThrows = null; const dm = JSON.parse(JSON.stringify(d0)); dm.routes[0].when = { eval: 1 };
  try { QFm.evaluate(dm, { owner: 'o', ownerBu: 'ERP', category: 'consult', classKeys: [], revenueCents: 100, gpCents: 1 }, io); } catch (e) { unfrozenThrows = e; }
  tt('D10.3 evaluate 對未驗證／無效的定義丟 FlowError(FLOW_INVALID_DEF)', !!unfrozenThrows && unfrozenThrows.code === 'FLOW_INVALID_DEF', unfrozenThrows && unfrozenThrows.code);
  // 非預設流程的行為：插入造成重複 → FLOW_INVALID_RESULT；多餘 modifier
  const dd = JSON.parse(JSON.stringify(d0));
  dd.modifiers = [{ id: 'm_dup', op: 'insert_before', target: 'n_mgr1', node: 'n_gm', when: true }];
  let rthrow = null; try { QFm.evaluate(dd, { owner: 'o', ownerBu: 'ERP', category: 'consult', classKeys: [], revenueCents: 1000000000, gpCents: 100000000 }, io); } catch (e) { rthrow = e; }
  tt('D10.4 調整使序列出現重複節點 → FlowError(FLOW_INVALID_RESULT)', !!rthrow && rthrow.code === 'FLOW_INVALID_RESULT', rthrow && rthrow.code);
  // D12 運算子真值表（evalCond 直接求值，期望值用 JS 原生 BigInt／數字比較獨立算出）
  const mkEnv = (rev, gp, level) => ({ facts: QFm.makeFacts({ owner: 'u1', ownerBu: 'ITS', category: 'consult', classKeys: ['consult', 'software'], revenueCents: rev, gpCents: gp }), modifier: false, applied: new Set(), level: () => level, resolveNode: () => ({ person: null, persons: [] }) });
  const num = (a, op, b) => ({ '==': a === b, '!=': a !== b, '>': a > b, '>=': a >= b, '<': a < b, '<=': a <= b })[op];
  const OPS6 = ['==', '!=', '>', '>=', '<', '<='];
  let bad12 = [];
  for (const [rhs, cents] of [['99.99', 9999n], ['100', 10000n], ['100.00', 10000n], ['100.01', 10001n], ['0', 0n], ['0.01', 1n], ['9999999999999.99', 999999999999999n]]) {
    for (const op of OPS6) {
      let got; try { got = QFm.evalCond({ cmp: ['amount', op, rhs] }, mkEnv(10000, 5, 1)); } catch (e) { got = 'throw'; }
      if (got !== num(10000n, op, cents)) bad12.push('amount ' + op + ' ' + rhs);
    }
  }
  tt('D12.1 金額 6 種運算子 × 7 個邊界值（-0.01／剛好／+0.01…）真值表全對', bad12.length === 0, bad12.slice(0, 3).join(';'));
  bad12 = [];
  for (const lv of [1, 2, 3]) {
    for (const rhs of [1, 2, 3]) {
      for (const op of OPS6) {
        let got; try { got = QFm.evalCond({ cmp: ['marginLevel', op, rhs] }, mkEnv(10000, 5, lv)); } catch (e) { got = 'throw'; }
        if (got !== num(lv, op, rhs)) bad12.push('level ' + lv + op + rhs);
      }
    }
  }
  tt('D12.2 毛利層級 3×3×6 真值表全對', bad12.length === 0, bad12.slice(0, 3).join(';'));
  bad12 = [];
  // marginPct 與字面值的關係（-1／0／1＝marginPct 小於／等於／大於 rhs），全部是手算的精確有理數
  const pctCases = [[1, 3, '33.3333', 1], [1, 3, '33.3334', -1], [1, 4, '25', 0], [1, 4, '25.0000', 0], [1, 4, '24.9999', 1], [1, 4, '25.0001', -1], [-1, 4, '-25', 0], [-1, 4, '-24.9999', -1], [-1, 4, '-25.0001', 1], [0, 5, '0', 0], [0, 5, '0.0001', -1], [0, 5, '-0.0001', 1], [5, 2, '250', 0], [5, 2, '249.9999', 1], [333333, 1000000, '33.3333', 0], [999999999999999, 1000000000000000, '99.9999', 1], [999999999999999, 1000000000000000, '100', -1], [7, 100, '7', 0], [29, 100, '29', 0], [57, 100, '57', 0], [14, 100, '14', 0], [58, 100, '58', 0], [29, 100, '28.9999', 1], [29, 100, '29.0001', -1]];
  for (const [gp, rev, rhs, rel] of pctCases) {
    for (const op of OPS6) {
      let got; try { got = QFm.evalCond({ cmp: ['marginPct', op, rhs] }, mkEnv(rev, gp, 1)); } catch (e) { got = 'throw'; }
      if (got !== num(rel, op, 0)) bad12.push('pct gp=' + gp + '/rev=' + rev + ' ' + op + ' ' + rhs);
    }
  }
  tt('D12.3 毛利率（有理數精確比較，含負值、250%、貼近 100% 的極大營收、浮點會算錯的 7%／29%／57%）24 個關係 × 6 種運算子全對', bad12.length === 0, bad12.slice(0, 3).join(';'));
  bad12 = [];
  const chk = (ast, exp, nm) => { let got; try { got = QFm.evalCond(ast, mkEnv(10000, 5, 1)); } catch (e) { got = 'throw'; } if (got !== exp) bad12.push(nm); };
  for (const [v, val, other] of [['category', 'consult', 'crm'], ['ownerBuPrimary', 'ITS', 'ERP'], ['owner', 'u1', 'zz']]) {
    chk({ cmp: [v, '==', val] }, true, v + '==自己'); chk({ cmp: [v, '!=', val] }, false, v + '!=自己'); chk({ cmp: [v, '==', other] }, false, v + '==他人'); chk({ cmp: [v, '!=', other] }, true, v + '!=他人');
    chk({ cmp: [v, 'in', [val, other]] }, true, v + ' in 含自己'); chk({ cmp: [v, 'in', [other]] }, false, v + ' in 不含自己'); chk({ cmp: [v, 'in', []] }, false, v + ' in 空清單');
    chk({ cmp: [v, 'not_in', [val]] }, false, v + ' not_in 含自己'); chk({ cmp: [v, 'not_in', [other]] }, true, v + ' not_in 不含自己'); chk({ cmp: [v, 'not_in', []] }, true, v + ' not_in 空清單');
  }
  chk({ cmp: ['classKeys', 'contains', 'software'] }, true, 'contains 有'); chk({ cmp: ['classKeys', 'contains', 'crm'] }, false, 'contains 無');
  chk({ cmp: ['classKeys', 'not_contains', 'crm'] }, true, 'not_contains 無'); chk({ cmp: ['classKeys', 'not_contains', 'software'] }, false, 'not_contains 有');
  chk({ all: [true, true] }, true, 'all TT'); chk({ all: [true, false] }, false, 'all TF'); chk({ any: [false, false] }, false, 'any FF'); chk({ any: [false, true] }, true, 'any FT');
  chk({ not: true }, false, 'not T'); chk({ not: false }, true, 'not F'); chk(false, false, 'false 字面'); chk(true, true, 'true 字面');
  tt('D12.4 字串／集合／邏輯運算子真值表全對', bad12.length === 0, bad12.slice(0, 4).join(';'));
  // D13 marginLevelOf 直接測（含極大營收；期望值用獨立的 BigInt 式子，並與 QA.requiredPath 在 ≤1,000 萬時交叉驗證）
  {
    const table = QFm.buildLegacyDefault({ buSettings: {}, exceptions: [] }).tables.margin;
    const bad13 = []; let n13 = 0;
    for (const row of QA.ROWS) {
      for (const rev of [1, 7, 100, 12345, 100000000, 999999999, 1000000000, 123456789012345, 999999999999999, 1000000000000000]) {
        for (const gp of gpPoints(row.key, rev)) {
          n13++;
          const g100 = BigInt(gp) * 100n, R = BigInt(rev);
          const exp = row.fixedLevel ? row.fixedLevel : (row.l1 !== null && g100 >= BigInt(row.l1) * R) ? 1 : (g100 >= BigInt(row.l2) * R) ? 2 : 3;
          let got; try { got = QFm.marginLevelOf(table, row.key, BigInt(gp), R); } catch (e) { got = 'throw'; }
          if (got !== exp) bad13.push(row.key + ' rev=' + rev + ' gp=' + gp + ' got=' + got + ' exp=' + exp);
          if (rev <= 1000000000) { const rp = QA.requiredPath(row.key, gp, rev); if (Math.max(exp, 1) !== rp.level) bad13.push('與 requiredPath 不符 ' + row.key + ' ' + rev + ' ' + gp); }
        }
      }
    }
    tt('D13 marginLevelOf：' + n13 + ' 個點（含營收 1e15 分、每個門檻 ±1 分）與獨立 BigInt 式子相等，且 ≤1,000 萬的點與 QA.requiredPath.level 一致', bad13.length === 0, bad13.slice(0, 3).join(';'));
  }
  // D8b makeFacts 預設
  { const f = QFm.makeFacts({ owner: 'o', category: 'x', revenueCents: 5, gpCents: 1 }); tt('D8b makeFacts：沒給主要 BU → 預設 ERP（與舊程式 normalizeBu(...)[0] || ERP 一致）；金額欄位轉 BigInt', f.ownerBuPrimary === 'ERP' && typeof f.amountCents === 'bigint' && typeof f.gpCents === 'bigint' && f.classKeys.length === 0); }
  // 模式字串
  const em = QFm.engineMode;
  tt('D11.1 engineMode：未設定／空字串／off → off', em(undefined) === 'off' && em('') === 'off' && em('off') === 'off' && em(null) === 'off');
  tt('D11.2 engineMode：shadow／on（大小寫、空白不拘）', em('shadow') === 'shadow' && em(' ON ') === 'on' && em('Shadow') === 'shadow' && em('on') === 'on');
  tt('D11.3 engineMode：無法辨識的值一律 off（fail-safe）', ['1', 'true', 'yes', 'enable', 'shadow2', 'o n', 'ONN', 0, 1, true, {}].every((v) => em(v) === 'off'));
  tt('D11.4 engineModeInfo：無法辨識的值 valid=false，空值 valid=true', QFm.engineModeInfo('banana').valid === false && QFm.engineModeInfo('').valid === true && QFm.engineModeInfo('on').valid === true);
}

// ───────────────────────── 變異測試 ─────────────────────────
function loadMutant(src) {
  const m = new Module(QF_PATH, module);
  m.filename = QF_PATH;
  m.paths = Module._nodeModulePaths(path.dirname(QF_PATH));
  m._compile(src, QF_PATH);
  return m.exports;
}
const SRC = fs.readFileSync(QF_PATH, 'utf8');
const MUTATIONS = [
  ['金額比較 > 改成 >=', "case '>': return c > 0;", "case '>': return c >= 0;"],
  ['金額比較 >= 改成 >', "case '>=': return c >= 0;", "case '>=': return c > 0;"],
  ['預設流程：董事長金額門檻 > 改成 >=', "{ cmp: ['amount', '>', yuan(QA.AMOUNT.chairman)] }", "{ cmp: ['amount', '>=', yuan(QA.AMOUNT.chairman)] }"],
  ['預設流程：董事會金額門檻 > 改成 >=', "when: { cmp: ['amount', '>', yuan(QA.AMOUNT.board)] }", "when: { cmp: ['amount', '>=', yuan(QA.AMOUNT.board)] }"],
  ['預設流程：董事會門檻誤用董事長金額', "when: { cmp: ['amount', '>', yuan(QA.AMOUNT.board)] }", "when: { cmp: ['amount', '>', yuan(QA.AMOUNT.chairman)] }"],
  ['預設流程：漏掉「毛利層級==3 → 董事長」', "{ any: [{ cmp: ['amount', '>', yuan(QA.AMOUNT.chairman)] }, { cmp: ['marginLevel', '==', 3] }] }", "{ cmp: ['amount', '>', yuan(QA.AMOUNT.chairman)] }"],
  ['預設流程：毛利層級==2 改成 ==1', "when: { cmp: ['marginLevel', '==', 2] }", "when: { cmp: ['marginLevel', '==', 1] }"],
  ['預設流程：預設路線誤用總經理路徑', "when: true, path: ids(QA.PATHS[1])", "when: true, path: ids(QA.PATHS[2])"],
  ['決策表：l1 的 >= 改成 >', "if (row.l1 !== null && gp100 >= BigInt(row.l1) * revCents) return 1;", "if (row.l1 !== null && gp100 > BigInt(row.l1) * revCents) return 1;"],
  ['決策表：l2 的 >= 改成 >', "if (gp100 >= BigInt(row.l2) * revCents) return 2;", "if (gp100 > BigInt(row.l2) * revCents) return 2;"],
  ['決策表：忽略 fixedLevel', "if (row.fixedLevel) return row.fixedLevel;", "if (false) return row.fixedLevel;"],
  ['決策表：找不到列時不退回 other（改成第一列）', "row = row || other;", "row = row || (table[0]);"],
  ['EQUIV 金額用浮點（毛利率先除成浮點數再與門檻比較）', "if (row.l1 !== null && gp100 >= BigInt(row.l1) * revCents) return 1;\n  if (gp100 >= BigInt(row.l2) * revCents) return 2;", "if (row.l1 !== null && Number(gpCents) / Number(revCents) * 100 >= row.l1) return 1;\n  if (Number(gpCents) / Number(revCents) * 100 >= row.l2) return 2;"],
  ['EQUIV 金額用浮點（毛利比較兩邊改用 Number）', "if (row.l1 !== null && gp100 >= BigInt(row.l1) * revCents) return 1;\n  if (gp100 >= BigInt(row.l2) * revCents) return 2;", "if (row.l1 !== null && Number(gp100) >= row.l1 * Number(revCents)) return 1;\n  if (Number(gp100) >= row.l2 * Number(revCents)) return 2;"],
  ['少一個 BU 組合（最後一個啟用的 BU 被丟掉）', ".filter((bu) => bs[bu] && bs[bu].hasMgr2Step).sort();", ".filter((bu) => bs[bu] && bs[bu].hasMgr2Step).sort().slice(0, -1);"],
  ['少一個 BU（只認前四個 BU）', ".filter((bu) => bs[bu] && bs[bu].hasMgr2Step).sort();", ".filter((bu) => bs[bu] && bs[bu].hasMgr2Step && bu !== '資安').sort();"],
  ['漏掉 skip_mgr1（例外名單永遠空）', "e.action === 'skip_mgr1' &&", "e.action === 'skip_mgr1_x' &&"],
  ['漏掉 skip_mgr1（拿掉 m_skip 調整）', "      { id: 'm_skip', name: '二級主管例外：跳過一級主管', op: 'remove', target: 'n_mgr1', when: { all: [{ applied: 'm_mgr2' }, { resolvedIn: ['n_mgr2', inp.skipUsers] }] } },\n", ""],
  ['二級主管插在 mgr1 之後', "op: 'insert_before', target: 'n_mgr1', node: 'n_mgr2'", "op: 'insert_after', target: 'n_mgr1', node: 'n_mgr2'"],
  ['二級主管找不到人也插入（拿掉 resolvable）', "{ resolvable: 'n_mgr2' }", "true"],
  ['skip 例外不看 applied(m_mgr2)', "{ applied: 'm_mgr2' }, { resolvedIn: ['n_mgr2', inp.skipUsers] }", "{ resolvedIn: ['n_mgr2', inp.skipUsers] }"],
  ['略過去重（不檢查 SIGNER_OVERLAP）', "} else if (!distinctAssignment(cands)) {", "} else if (false) {"],
  ['去重回溯壞掉（不標記已用）', "    used.add(p);\n    if (distinctAssignment(cands, i + 1, used)) return true;", "    if (distinctAssignment(cands, i + 1, used)) return true;"],
  ['成本填寫人不排除', "if (io.costBy) excluded.add(io.costBy);", ""],
  ['擁有者不排除', "const excluded = new Set([io.owner]);", "const excluded = new Set();"],
  ['mgr1 沒人不擋', "if (actualTiers.includes('mgr1') && !mgr1) {", "if (false) {"],
  ['總經理名單不過濾失效帳號', "if (actualTiers.includes('gm') && !activeOf(io.roster.gm).length)", "if (actualTiers.includes('gm') && !io.roster.gm.length)"],
  ['董事會代核人沒人不擋', "if (actualTiers.includes('board') && !io.boardProxy.length)", "if (false)"],
  ['主要 BU 預設值改成 ITS', "String(i.ownerBu === undefined || i.ownerBu === null ? 'ERP' : i.ownerBu)", "String(i.ownerBu === undefined || i.ownerBu === null ? 'ITS' : i.ownerBu)"],
  ['關卡數上限改成 1', "limits: { maxSteps: MAX_STEPS_CAP },", "limits: { maxSteps: 1 },"],
  ['金額用浮點（條件 AST 的毛利率比較改用浮點）', "return cmpOp(op, cmpBig(f.gpCents * 100n * (10n ** BigInt(p.s)), p.n * f.amountCents));", "return cmpOp(op, Math.sign(Number(f.gpCents) / Number(f.amountCents) * 100 - Number(rhs)));"],
  ['金額用浮點（條件 AST 的金額比較改用 parseFloat）', "case 'money': return cmpOp(op, cmpBig(f.amountCents, parseMoneyCents(rhs)));", "case 'money': return cmpOp(op, Math.sign(Number(f.amountCents) - Math.round(parseFloat(rhs) * 100)));"],
  ['路線選擇改成最後一條符合者', "    if (m) { route = r; break; }", "    if (m) { route = r; }"],
];

// ───────────────────────── 路由層（假 app＋記憶體資料庫）─────────────────────────
const RUSERS = [
  { username: 'admin1', role: 'admin', active: true, displayName: 'Admin' },
  { username: 'own1', role: 'user', active: true, displayName: 'Own1', supervisor: 'm2a', bu: ['ITS'] },
  { username: 'own2', role: 'user', active: true, displayName: 'Own2', supervisor: 'm1a', bu: ['ERP'] },
  { username: 'own3', role: 'user', active: true, displayName: 'Own3', supervisor: 'ghost', bu: ['ERP'] },
  { username: 'm2a', role: 'manager2', active: true, displayName: 'M2a', supervisor: 'm1a', bu: ['ITS'] },
  { username: 'm1a', role: 'manager1', active: true, displayName: 'M1a', bu: ['ITS', 'ERP'] },
  { username: 'm1b', role: 'manager1', active: true, displayName: 'M1b', bu: ['ITS', 'ERP'] },
  { username: 'gm1', role: 'executive', active: true, displayName: 'Gm' },
  { username: 'ch1', role: 'executive', active: true, displayName: 'Chair' },
  { username: 'sec1', role: 'secretary', active: true, displayName: 'Sec', bu: ['ITS'] },
];
const RCLASSES = { PS: { cls: 'software', costBySales: true }, PC: { cls: 'consult', costBySales: true }, PX: { cls: 'other', costBySales: true } };
const RVAR = {
  V1: { buSettings: { ITS: { hasMgr2Step: true } }, exceptions: [], roster: { gm: ['gm1'], chairman: ['ch1'], boardProxy: [] } },
  V2: { buSettings: { ITS: { hasMgr2Step: true }, ERP: { hasMgr2Step: false } }, exceptions: [{ username: 'm2a', action: 'skip_mgr1' }], roster: { gm: ['gm1'], chairman: ['ch1'], boardProxy: ['sec1'] } },
  V3: { buSettings: {}, exceptions: [], roster: { gm: [], chairman: ['ch1'], boardProxy: [] } },
  V4: { buSettings: { ITS: { hasMgr2Step: true } }, exceptions: [], roster: { gm: ['gm1'], chairman: ['gm1'], boardProxy: [] } },
};
const mkQ = (id, owner, products, price, cost, extra) => Object.assign({
  id, quoteNo: 'QU-' + id, owner, company: 'TestCo', projectName: 'Proj ' + id, quoteDate: '2026-10-08', status: 'draft', createdAt: '2026-10-08T00:00:00Z', updatedAt: '2026-10-08T00:00:00Z',
  validUntil: '2026-10-30', products, approval: null, discountType: 'none', discountValue: 0,
  items: [{ lid: id + '-i1', desc: '服務', unit: '式', qty: 1, unitPrice: price, cost }],
}, extra || {});
const RQUOTES = () => [
  mkQ('a1', 'own1', ['PS'], 800000, 400000), mkQ('a2', 'own2', ['PS'], 800000, 400000), mkQ('a3', 'own1', ['PS'], 800000, 760000), mkQ('a4', 'own2', ['PS'], 800000, 770000),
  mkQ('a5', 'own1', ['PS'], 12000000, 6000000), mkQ('a6', 'own2', ['PS'], 60000000, 30000000), mkQ('a7', 'own1', ['PX'], 800000, 400000), mkQ('a9', 'own3', ['PS'], 800000, 400000),
  mkQ('b1', 'own1', ['PS'], 800000, 400000), mkQ('b2', 'own2', ['PS'], 800000, 770000), mkQ('b3', 'own1', ['PS'], 800000, 400000), mkQ('b4', 'own2', ['PC'], 60000000, 20000000),
  mkQ('c1', 'own1', ['PS'], 10000000, 5000000), mkQ('c2', 'own1', ['PS'], 10000000.01, 5000000), mkQ('c3', 'own2', ['PS'], 50000000, 25000000), mkQ('c4', 'own2', ['PS'], 50000000.01, 25000000),
];
function mkRouteEnv(root, variant, opts) {
  opts = opts || {};
  const reg = require(path.join(root, 'lib/quoteRoutes.js'));
  const PNL = require(path.join(root, 'lib/quotePnlExcel.js'));
  const routes = {};
  const app = {};
  ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
  const v = RVAR[variant];
  const env = {
    data: { quotations: RQUOTES(), quoteApproval: { roster: Object.assign({ costProviders: [], sealManagers: [] }, JSON.parse(J(v.roster))), productClasses: RCLASSES }, approvalFlowConfig: { version: 1, buSettings: JSON.parse(J(v.buSettings)), exceptions: JSON.parse(J(v.exceptions)) }, contacts: [] },
    logs: [], notes: [], saves: 0, seq: 0, nseq: 0, rec: [],
  };
  env.auth = { users: JSON.parse(J(RUSERS)) };
  const deps = {
    db: { load: () => JSON.parse(J(env.data)), save: (d) => { env.saves++; env.data = JSON.parse(J(d)); }, flush: async () => {}, auditStore: { total: async (f) => env.logs.filter((l) => l[0] === f.action).length } }, loadAuth: () => env.auth, saveAuth: () => {},
    requireAuth: (rq, rs, next) => next(), requireAdmin: function requireAdminStub(rq, rs, next) { next(); },
    writeLog: (...a) => env.logs.push([a[0], a[1], a[2], a[3]]),
    pushNotification: (to, type, title, body, refId) => { env.notes.push([to, type, title, body, refId]); const d = deps.db.load(); (d.notifications = d.notifications || []).unshift({ id: 'n' + (++env.nseq), to, type, title, body, refId, read: false, createdAt: 'TS' }); deps.db.save(d); },
    getViewableOwners: (rq) => { const rl = rq.session.user.role; if (['admin', 'executive', 'manager1', 'manager2', 'secretary'].includes(rl)) return ['own1', 'own2', 'own3']; return [rq.session.user.username]; },
    sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-NEW', taipeiToday: () => '2026-10-20', resolveIssuer: () => ({}),
    buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: PNL.buildQuotePnlExcel, QUOTE_TEMPLATE: '',
    uuidv4: () => 'u' + (++env.seq), normalizeBu, getUserFeatures: () => ['quotations'],
  };
  const reg_ = reg(app, deps);
  env.internal = reg_._internal;
  env.routeHandlers = (m, pth) => routes[m.toUpperCase() + ' ' + pth];
  env.requireAdmin = deps.requireAdmin;
  env.call = (user, method, p, params, body) => new Promise((resolve, reject) => {
    const h = [].concat(...routes[method.toUpperCase() + ' ' + p]);
    const role = (env.auth.users.find((u) => u.username === user) || {}).role;
    const req = { session: { user: { username: user, role } }, params: params || {}, query: {}, headers: {}, body: body === undefined ? {} : JSON.parse(J(body)), ip: '127.0.0.1' };
    const rs = { status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j }); return this; }, send() { resolve({ s: this._s || 200, j: null }); return this; }, setHeader() {}, set() {} };
    let i = 0;
    const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
    next();
  });
  return env;
}
const TSRE = /\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z/g;
const norm = (x) => J(x).replace(TSRE, 'TS');
const ACTOR = { mgr2: 'm2a', mgr1: 'm1a', gm: 'gm1', chairman: 'ch1', board: 'sec1' };
const GET_ONE = '/api/quotations/:id', SUBMIT = '/api/quotations/:id/submit', APPROVE = '/api/quotations/:id/approve', RETURN = '/api/quotations/:id/return', WITHDRAW = '/api/quotations/:id/withdraw', REASSIGN = '/api/quotations/:id/reassign';
const Q = (env, id) => env.data.quotations.find((x) => x.id === id);

/** 一組固定的請求腳本：預覽 → 送簽 → 逐關核准（含駁回重送、撤回重送、改派）；每個回應都記進 env.rec */
async function scenario(env) {
  const rec = (label, r) => env.rec.push([label, r.s, norm(r.j)]);
  const ids = env.data.quotations.map((q) => q.id);
  const approveAll = async (id, maxSteps) => {
    for (let k = 0; k < (maxSteps || 8); k++) {
      const q = Q(env, id);
      if (!q.approval || q.approval.state !== 'pending') break;
      const step = q.approval.steps[q.approval.cur];
      const body = { hash: q.approval.hash, comment: 'ok' };
      if (step.tier === 'board') Object.assign(body, { resolutionDate: '2026-10-10', resolutionNo: '115-BD-001' });
      rec('approve:' + id + ':' + step.tier, await env.call(ACTOR[step.tier], 'POST', APPROVE, { id }, body));
    }
  };
  for (const id of ids) {
    const owner = Q(env, id).owner;
    rec('get:' + id, await env.call(owner, 'GET', GET_ONE, { id }));
    const sub = await env.call(owner, 'POST', SUBMIT, { id }, {});
    rec('submit:' + id, sub);
    if (sub.s !== 200) continue;
    if (id === 'b1') {   // 駁回後重送
      const q = Q(env, id); const step = q.approval.steps[q.approval.cur];
      rec('return:' + id, await env.call(ACTOR[step.tier], 'POST', RETURN, { id }, { hash: q.approval.hash, comment: '請修正' }));
      rec('resubmit:' + id, await env.call(owner, 'POST', SUBMIT, { id }, {}));
    }
    if (id === 'b2') {   // 撤回後重送
      rec('withdraw:' + id, await env.call(owner, 'POST', WITHDRAW, { id }, {}));
      rec('resubmit:' + id, await env.call(owner, 'POST', SUBMIT, { id }, {}));
    }
    if (id === 'b3') {   // 管理員改派一級主管
      const q = Q(env, id);
      if (q.approval.steps[q.approval.cur].tier === 'mgr1') rec('reassign:' + id, await env.call('admin1', 'POST', REASSIGN, { id }, { username: 'm1b', reason: '輪調' }));
    }
    await approveAll(id);
    rec('final-get:' + id, await env.call(owner, 'GET', GET_ONE, { id }));
    rec('admin-get:' + id, await env.call('admin1', 'GET', GET_ONE, { id }));
  }
}
function deepDiff(x, y, p, out) {
  out = out || [];
  if (out.length >= 6 || J(x) === J(y)) return out;
  if (x && y && typeof x === 'object' && typeof y === 'object') { for (const k of new Set([...Object.keys(x), ...Object.keys(y)])) deepDiff(x[k], y[k], p + '.' + k, out); } else out.push(p + ': ' + J(x) + ' | ' + J(y));
  return out;
}
const stripFlow = (o) => JSON.parse(JSON.stringify(o, (k, v) => (k === 'flow' ? undefined : v)));

async function withEnvMode(mode, fn) {
  const prev = process.env.QUOTE_FLOW_ENGINE;
  if (mode === undefined) delete process.env.QUOTE_FLOW_ENGINE; else process.env.QUOTE_FLOW_ENGINE = mode;
  try { return await fn(); } finally { if (prev === undefined) delete process.env.QUOTE_FLOW_ENGINE; else process.env.QUOTE_FLOW_ENGINE = prev; }
}
async function runVariant(root, variant, mode) {
  return withEnvMode(mode, async () => { const env = mkRouteEnv(root, variant); await scenario(env); return env; });
}
const finalState = (env) => ({ data: norm(env.data), logs: norm(env.logs), notes: norm(env.notes), rec: env.rec });

// ───────────────────────── 主程式 ─────────────────────────
(async () => {
  const T0 = Date.now();
  // ═════ A 釘選與結構 ═════
  const emptyCfg = { buSettings: {}, exceptions: [] };
  const hNow = QF.defHash(QF.buildLegacyDefault(emptyCfg));
  t('A1 預設流程（無 BU／例外）的 defHash＝腳本釘選值＝lib 的 DEFAULT_FLOW_HASH', hNow === PINNED_DEFAULT_FLOW_HASH && QF.DEFAULT_FLOW_HASH === PINNED_DEFAULT_FLOW_HASH, hNow);
  t('A2 預設流程通過 validateFlowDef，且已深凍結', QF.validateFlowDef(QF.buildLegacyDefault(emptyCfg)).ok && Object.isFrozen(QF.legacyDefault(emptyCfg)) && Object.isFrozen(QF.legacyDefault(emptyCfg).routes[0]));
  t('A3 節點 tier 集合＝現有五種（QA.TIERS 的 key），M1 沒有新增 tier 字串', J(QF.TIER_KEYS) === J(Object.keys(QA.TIERS)) && J(QF.TIER_KEYS) === J(TIER5) && QF.buildLegacyDefault(emptyCfg).nodes.every((n) => TIER5.includes(n.tier)));
  t('A4 預設流程的路線／節點 id 只用 n_<現有 tier>', QF.buildLegacyDefault(emptyCfg).routes.every((r) => r.path.every((id) => id.startsWith('n_') && TIER5.includes(id.slice(2)))));
  const srcNoComments = SRC.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
  t('A5 lib/quoteFlow.js 沒有 eval／new Function／vm／child_process／動態 require／RegExp 建構子', !/\beval\s*\(|new\s+Function|\bFunction\s*\(|require\(\s*['"](vm|child_process)['"]|\bnew\s+RegExp\b|require\(\s*[^'"\s]/.test(srcNoComments), (srcNoComments.match(/\beval\s*\(|new\s+Function|\bFunction\s*\(|require\(\s*['"](vm|child_process)['"]|new\s+RegExp|require\(\s*[^'"\s]/) || [''])[0]);
  t('A6 lib/quoteFlow.js 只 require crypto 與 ./quoteApproval（純函式，無 I/O）', J((srcNoComments.match(/require\('[^']+'\)/g) || []).sort()) === J(["require('./quoteApproval')", "require('crypto')"]), J(srcNoComments.match(/require\('[^']+'\)/g)));
  t('A7 lib/quoteRoutes.js 以靜態 require 引入 ./quoteFlow（Vercel 追蹤得到）', /const QF = require\('\.\/quoteFlow'\);/.test(fs.readFileSync(path.join(ROOT, 'lib/quoteRoutes.js'), 'utf8')));
  const mail = fs.readFileSync(path.join(ROOT, 'lib/mail/quoteMail.js'), 'utf8');
  t('A8 信件模組未被修改的前提：quoteMail 的未知 tier 預設當 mgr1 寄（M1 不新增 tier 因此不會走到）', /mgr1/.test(mail));

  // ═════ B 等價網格（完整）═════
  const full = QUICK ? { strideF: 6, strideG1: 40, strideG2: 60, strideG3: 90, nRandom: 5000, seed: SEED } : { strideF: 1, strideG1: 6, strideG2: 1, strideG3: 1, nRandom: 100000, seed: SEED };
  const tB0 = Date.now();
  const st = SKIP.has('B') ? core(QF, { strideF: 400, strideG1: 400, strideG2: 4000, strideG3: 4000, nRandom: 100, seed: SEED }) : core(QF, full);
  const grpTxt = Object.keys(st.groups).sort().map((g) => g + ':' + st.groups[g].total).join(' ');
  t('B 等價網格：引擎與舊路徑逐項相同（財務層、實際關卡、mgr2、mgr1、阻擋碼與訊息）— ' + st.total + ' 筆（' + grpTxt + '，隨機種子 ' + SEED + '），不一致 ' + st.fails + '、引擎例外 ' + st.thrown, st.fails === 0 && st.thrown === 0 && st.total > (QUICK ? 1000 : 400000), st.first.join('\n  '));
  for (const g of Object.keys(st.groups).sort()) t('B.' + g + ' 分組：' + st.groups[g].total + ' 筆全相等', st.groups[g].fails === 0 && st.groups[g].total > 0, st.groups[g].fails);
  if (!QUICK) t('B.golden 回歸快照：整批引擎輸出的 sha256 ＝ GOLDEN（' + st.digestHex.slice(0, 16) + '…）', st.digestHex === GOLDEN, st.digestHex);
  console.log('[B] ' + st.total + ' 筆，耗時 ' + ((Date.now() - tB0) / 1000).toFixed(1) + 's，digest=' + st.digestHex);

  // 邊界點抽查（人讀得懂的）
  const w0 = mkWorld({});
  const L = (cat, rev, gp) => engineOf(QF, w0, { cat, rev, gp }).financialTiers.join('>');
  t('B.edge1 金額剛好 1,000 萬 → 不加董事長；+1 分 → 加董事長（consult 毛利 40%）', L('consult', 1000000000, 400000000) === 'mgr1' && L('consult', 1000000001, 400000001) === 'mgr1>gm>chairman');
  t('B.edge2 金額剛好 5,000 萬 → 董事長路徑；+1 分 → 董事會路徑', L('consult', 5000000000, 2000000000) === 'mgr1>gm>chairman' && L('consult', 5000000001, 2000000001) === 'mgr1>gm>board');
  t('B.edge3 consult 毛利剛好 25% → 一級主管；少 1 分 → 總經理；剛好 15% → 總經理；少 1 分 → 董事長', L('consult', 100000000, 25000000) === 'mgr1' && L('consult', 100000000, 24999999) === 'mgr1>gm' && L('consult', 100000000, 15000000) === 'mgr1>gm' && L('consult', 100000000, 14999999) === 'mgr1>gm>chairman');
  t('B.edge4 ot（l1 null）毛利再高也不能一級主管終局；other 固定總經理', L('ot', 100000000, 100000000) === 'mgr1>gm' && L('other', 100000000, 100000000) === 'mgr1>gm' && L('other', 100000000, -5) === 'mgr1>gm');
  t('B.edge5 金額剛好 1,000 萬時毛利 25% 邊界精確（少 1 分 → 總經理）；極大營收（1e15 分）直接走董事會', L('consult', 1000000000, 250000000) === 'mgr1' && L('consult', 1000000000, 249999999) === 'mgr1>gm' && L('consult', 1e15, 1) === 'mgr1>gm>board' && L('consult', 1e15, 1e15) === 'mgr1>gm>board');
  const qr = (cat, rev, gp) => QA.requiredPath(cat, gp, rev).tiers.join('>');
  t('B.edge6 邊界點與舊 requiredPath 一致', ['consult:1000000000:400000000', 'consult:1000000001:400000001', 'consult:5000000001:2000000001', 'consult:100000000:24999999', 'ot:100000000:100000000'].every((s) => { const [c, r, g] = s.split(':'); return qr(c, +r, +g) === L(c, +r, +g); }));
  // 人選規則（轉換表的人讀版）
  const wPeople = mkWorld({ ownerBu: ['ERP'], bus: ['ERP'], exc: EXCK[1] });
  const pp = engineOf(QF, wPeople, { cat: 'consult', rev: 500000000, gp: 200000000 });
  t('B.people1 啟用二級主管的 BU：mgr2 在 mgr1 前；例外名單內的二級主管 → 移除 mgr1', J(engineOf(QF, mkWorld({ bus: ['ERP'] }), { cat: 'consult', rev: 500000000, gp: 200000000 }).actualTiers) === J(['mgr2', 'mgr1']) && J(pp.actualTiers) === J(['mgr2']) && pp.mgr1 === null && pp.mgr2 === 'm2');
  t('B.people2 BU 未啟用 → 不插入 mgr2、mgr2 指派為 null；多 BU 只看第一個', J(engineOf(QF, mkWorld({ bus: ['ITS'], ownerBu: ['ERP', 'ITS'] }), { cat: 'consult', rev: 500000000, gp: 200000000 }).actualTiers) === J(['mgr1']) && J(engineOf(QF, mkWorld({ bus: ['ITS'], ownerBu: ['ITS', 'ERP'] }), { cat: 'consult', rev: 500000000, gp: 200000000 }).actualTiers) === J(['mgr2', 'mgr1']));
  t('B.people3 直屬不是二級主管（停用／唯讀／角色不符）→ 靜默不插入 mgr2（舊程式行為，忠實複製）', ['m2inactive', 'm2view', 'm2ro', 'm2wrongrole', 'm2pool'].every((s) => J(engineOf(QF, mkWorld({ sup: s, bus: ['ERP'] }), { cat: 'consult', rev: 500000000, gp: 200000000 }).actualTiers) === J(['mgr1'])));
  t('B.people4 同一人在總經理與董事長名單 → SIGNER_OVERLAP；與擁有者重疊 → NO_ELIGIBLE_SIGNER', engineOf(QF, mkWorld({ gm: ['g1'], ch: ['g1'] }), { cat: 'consult', rev: 500000000, gp: 25000000 }).blockers[0].code === 'SIGNER_OVERLAP' && engineOf(QF, mkWorld({ gm: ['o'] }), { cat: 'consult', rev: 500000000, gp: 100000000 }).blockers[0].code === 'NO_ELIGIBLE_SIGNER');

  // ═════ D 條件 AST 安全 ═════
  astTests(QF, t);

  // ═════ C 路由層 ═════
  const reg0 = process.env.QUOTE_FLOW_ENGINE;
  const VARS_ = Object.keys(RVAR);
  const offRuns = {}, shadowRuns = {}, onRuns = {};
  for (const v of VARS_) { offRuns[v] = await runVariant(ROOT, v, undefined); shadowRuns[v] = await runVariant(ROOT, v, 'shadow'); onRuns[v] = await runVariant(ROOT, v, 'on'); }
  const submitsOk = (env) => env.rec.filter((r) => /^submit:/.test(r[0]) && r[1] === 200).length;
  const blocked = (env) => env.rec.filter((r) => /^submit:/.test(r[0]) && r[1] === 400).map((r) => JSON.parse(r[2]).code + ':' + (JSON.parse(r[2]).blockers || []).map((b) => b.code).join('+')).join(',');
  t('C0 劇本有效：V1 的 16 張單中 15 張送簽成功並簽到最終核准；1 張（直屬主管不存在）被 NO_MANAGER 擋下', submitsOk(offRuns.V1) === 15 && /NO_MANAGER/.test(blocked(offRuns.V1)) && offRuns.V1.data.quotations.filter((q) => q.approval && q.approval.state === 'approved').length >= 13, 'submits=' + submitsOk(offRuns.V1) + ' blocked=' + blocked(offRuns.V1));
  const tiersSeen = new Set(); offRuns.V1.data.quotations.forEach((q) => q.approval && q.approval.steps.forEach((s) => tiersSeen.add(s.tier)));
  t('C0b V1 劇本涵蓋五種關卡', TIER5.every((x) => tiersSeen.has(x)), [...tiersSeen].join(','));
  t('C0c V3（無總經理名單）／V4（總經理與董事長同人）會被阻擋：NO_GM／SIGNER_OVERLAP', /NO_GM/.test(blocked(offRuns.V3)) && /SIGNER_OVERLAP/.test(blocked(offRuns.V4)), blocked(offRuns.V3) + ' | ' + blocked(offRuns.V4));
  for (const v of VARS_) {
    const off = finalState(offRuns[v]), sh = finalState(shadowRuns[v]), on = finalState(onRuns[v]);
    t('C1.' + v + ' shadow：所有回應／儲存資料／稽核／通知與 off 逐位元相同（不改寫任何報價單資料、不多稽核）', J(sh.rec) === J(off.rec) && sh.data === off.data && sh.logs === off.logs && sh.notes === off.notes, J(sh.rec.filter((r, i) => J(r) !== J(off.rec[i])).slice(0, 1)).slice(0, 300));
    t('C2.' + v + ' on：所有回應與 off 逐位元相同（flow 快照不外洩）', J(on.rec) === J(off.rec), J(on.rec.filter((r, i) => J(r) !== J(off.rec[i])).slice(0, 1)).slice(0, 300));
    const onData = JSON.parse(on.data);
    t('C3.' + v + ' on：去掉 approval.flow 與歷程 meta.flow 後，儲存資料／稽核／通知與 off 逐位元相同', J(stripFlow(onData)) === J(stripFlow(JSON.parse(off.data))) && on.logs === off.logs && on.notes === off.notes, deepDiff(stripFlow(onData), stripFlow(JSON.parse(off.data)), 'data').join(' ;; '));
    const flows = onRuns[v].data.quotations.filter((q) => q.approval && q.approval.flow);
    const submitted = onRuns[v].data.quotations.filter((q) => q.approval && q.approval.steps && q.approval.steps.length);
    t('C4.' + v + ' on：每張在途／已核准的單都有 flow 快照（' + flows.length + '/' + submitted.length + '），且內容自洽（tiers＝steps、defHash＝預設流程、engine 版本）', flows.length === submitted.length && flows.every((q) => {
      const f = q.approval.flow;
      const flowCfg = onRuns[v].internal.getFlowCfg(onRuns[v].data);
      return f.v === 1 && f.engine === QF.ENGINE_ID && f.origin === 'legacy' && f.defNo === 0 && f.defHash === QF.defHash(QF.legacyDefault(flowCfg)) && J(f.tiers) === J(q.approval.steps.map((s) => s.tier)) && J(f.path) === J(q.approval.derived.tiers) && typeof f.routeId === 'string' && Array.isArray(f.applied) && Array.isArray(f.trace);
    }));
    const noFlowOff = offRuns[v].data.quotations.every((q) => !q.approval || !('flow' in q.approval)) && !/"flow"/.test(off.data);
    t('C5.' + v + ' off：資料中完全沒有 flow 鍵（舊單位元級不變）', noFlowOff);
    const leak = [...off.rec, ...on.rec].some((r) => /defHash|"flow"|"trace"|qfe-1|routeId|factsPublic/.test(r[2]));
    t('C6.' + v + ' 任何 API 回應都不含 flow／defHash／trace／routeId（不洩漏）', !leak);
    const audMis = shadowRuns[v].logs.filter((l) => /^QUOTE_FLOW_/.test(l[0])).length + onRuns[v].logs.filter((l) => /^QUOTE_FLOW_/.test(l[0])).length;
    t('C7.' + v + ' shadow／on 全劇本零稽核事件（零差異、零例外、零熔斷）', audMis === 0, audMis);
    const stats = shadowRuns[v].internal.flowReporter.stats;
    t('C8.' + v + ' shadow 引擎確實有跑（runs>0）且 mismatches=0、errors=0', stats.runs > 0 && stats.mismatches === 0 && stats.errors === 0, J(stats));
  }
  // 凍結快照語意：on 送簽後，開關／設定／名冊事後改變，在途單照 steps 走，不重新評估
  {
    const env = await withEnvMode('on', async () => mkRouteEnv(ROOT, 'V1'));
    const sub = await withEnvMode('on', () => env.call('own1', 'POST', SUBMIT, { id: 'a5' }, {}));
    const q0 = JSON.parse(J(Q(env, 'a5')));
    t('C9.1 on 送簽 a5（ITS、1,200 萬）→ steps＝mgr2>mgr1>gm>chairman，flow 快照已凍結', sub.s === 200 && J(q0.approval.steps.map((s) => s.tier)) === J(['mgr2', 'mgr1', 'gm', 'chairman']) && !!q0.approval.flow && q0.approval.flow.tiers.length === 4);
    // 事後：關掉引擎、關掉 ITS 的二級主管關、加 skip 例外、清空名冊 — 在途單仍照原 steps
    env.data.approvalFlowConfig = { version: 1, buSettings: { ITS: { hasMgr2Step: false } }, exceptions: [{ username: 'm2a', action: 'skip_mgr1' }] };
    const gmBefore = env.data.quoteApproval.roster.gm.slice();
    await withEnvMode('off', async () => {
      const q = Q(env, 'a5');
      const r1 = await env.call('m2a', 'POST', APPROVE, { id: 'a5' }, { hash: q.approval.hash, comment: '' });
      const r2 = await env.call('m1a', 'POST', APPROVE, { id: 'a5' }, { hash: Q(env, 'a5').approval.hash, comment: '' });
      const r3 = await env.call('gm1', 'POST', APPROVE, { id: 'a5' }, { hash: Q(env, 'a5').approval.hash, comment: '' });
      const r4 = await env.call('ch1', 'POST', APPROVE, { id: 'a5' }, { hash: Q(env, 'a5').approval.hash, comment: '' });
      t('C9.2 開關改 off、BU 二級主管關被關掉、加了 skip 例外之後，在途單仍依原 4 關逐關核准到完成', [r1, r2, r3, r4].every((r) => r.s === 200) && Q(env, 'a5').approval.state === 'approved' && J(Q(env, 'a5').approval.steps.map((s) => s.tier)) === J(['mgr2', 'mgr1', 'gm', 'chairman']), J([r1.s, r2.s, r3.s, r4.s]));
    });
    t('C9.3 核准完成後 flow 快照仍在、內容未變（不重新評估）', J(Q(env, 'a5').approval.flow) === J(q0.approval.flow) && gmBefore.length === 1);
    // 舊單（off 時送簽、沒有快照）在 on 模式下照常核准
    const env2 = mkRouteEnv(ROOT, 'V1');
    await withEnvMode('off', () => env2.call('own2', 'POST', SUBMIT, { id: 'a2' }, {}));
    t('C9.4 off 送簽的舊單沒有 approval.flow', !('flow' in Q(env2, 'a2').approval));
    const ap2 = await withEnvMode('on', () => env2.call('m1a', 'POST', APPROVE, { id: 'a2' }, { hash: Q(env2, 'a2').approval.hash, comment: '' }));
    t('C9.5 on 模式下核准舊單（無快照）→ 照 steps 走到核准，不產生 flow、不報錯', ap2.s === 200 && Q(env2, 'a2').approval.state === 'approved' && !('flow' in Q(env2, 'a2').approval));
    // 撤回：flow 隨 derived 清掉；重送依當下開關
    const env3 = mkRouteEnv(ROOT, 'V1');
    await withEnvMode('on', () => env3.call('own2', 'POST', SUBMIT, { id: 'a2' }, {}));
    await withEnvMode('on', () => env3.call('own2', 'POST', WITHDRAW, { id: 'a2' }, {}));
    const wd = Q(env3, 'a2').approval;
    t('C9.6 撤回後 approval.flow 已清除；SUBMIT 歷程 meta 仍留有 flow 摘要（可還原）', wd.state === 'none' && !('flow' in wd) && wd.history.find((h) => h.action === 'SUBMIT').meta.flow.defHash === QF.defHash(QF.legacyDefault(env3.internal.getFlowCfg(env3.data))));
    await withEnvMode(undefined, () => env3.call('own2', 'POST', SUBMIT, { id: 'a2' }, {}));
    t('C9.7 撤回後在 off 下重送 → 新一輪沒有 flow（走舊路徑）', !('flow' in Q(env3, 'a2').approval) && Q(env3, 'a2').approval.state === 'pending');
    // 駁回後重送：新一輪重建 flow
    const env4 = mkRouteEnv(ROOT, 'V1');
    await withEnvMode('on', () => env4.call('own2', 'POST', SUBMIT, { id: 'a2' }, {}));
    await withEnvMode('on', () => env4.call('m1a', 'POST', RETURN, { id: 'a2' }, { hash: Q(env4, 'a2').approval.hash, comment: 'x' }));
    const retAp = Q(env4, 'a2').approval;
    await withEnvMode('on', () => env4.call('own2', 'POST', SUBMIT, { id: 'a2' }, {}));
    t('C9.8 駁回後重送（on）→ state=pending、新的 flow 快照、歷程累積', retAp.state === 'returned' && Q(env4, 'a2').approval.state === 'pending' && !!Q(env4, 'a2').approval.flow && Q(env4, 'a2').approval.history.filter((h) => h.action === 'SUBMIT').length === 2);
    // 核准後改內容 → 作廢（重建 approval）：不帶 flow
    const env5 = mkRouteEnv(ROOT, 'V1');
    await withEnvMode('on', () => env5.call('own2', 'POST', SUBMIT, { id: 'a2' }, {}));
    await withEnvMode('on', () => env5.call('m1a', 'POST', APPROVE, { id: 'a2' }, { hash: Q(env5, 'a2').approval.hash, comment: '' }));
    const edit = await withEnvMode('on', () => env5.call('own2', 'PUT', GET_ONE, { id: 'a2' }, { items: [{ lid: 'a2-i1', desc: '服務', unit: '式', qty: 2, unitPrice: 800000, cost: 400000 }], products: ['PS'], company: 'TestCo', projectName: 'Proj a2', confirmVoid: true }));
    t('C9.9 核准後修改報價內容 → 核准作廢，approval 重建且不留 flow（PUT 200）', edit.s === 200 && Q(env5, 'a2').approval.state === 'none' && !('flow' in Q(env5, 'a2').approval), edit.s + ' ' + J(edit.j && edit.j.code) + ' ' + Q(env5, 'a2').approval.state);
  }
  // 故障注入
  {
    const realLegacyDefault = QF.legacyDefault, realPlan = QF.plan;
    const restore = () => { QF.legacyDefault = realLegacyDefault; QF.plan = realPlan; };
    const logsOf = (env, a) => env.logs.filter((l) => l[0] === a);
    try {
      // 1) 引擎例外
      QF.plan = () => { throw new QF.FlowError('FLOW_TEST_BOOM', 'injected'); };
      const offRef = offRuns.V1, ref = finalState(offRef);
      const shEx = await runVariant(ROOT, 'V1', 'shadow'), shExS = finalState(shEx);
      t('C10.1 注入引擎例外（shadow）：送簽等所有回應／資料與 off 完全相同（例外被吞掉）', J(shExS.rec) === J(ref.rec) && shExS.data === ref.data);
      t('C10.2 shadow 例外只寫 QUOTE_FLOW_SHADOW_ERROR（節流：同單同階段 10 分鐘內一筆），並計數', logsOf(shEx, 'QUOTE_FLOW_SHADOW_ERROR').length > 0 && logsOf(shEx, 'QUOTE_FLOW_SHADOW_ERROR').length <= 40 && shEx.internal.flowReporter.stats.errors > logsOf(shEx, 'QUOTE_FLOW_SHADOW_ERROR').length - 1 && shEx.internal.flowReporter.stats.suppressed > 0, J(shEx.internal.flowReporter.stats) + ' rows=' + logsOf(shEx, 'QUOTE_FLOW_SHADOW_ERROR').length);
      const onEx = await runVariant(ROOT, 'V1', 'on'), onExS = finalState(onEx);
      t('C10.3 注入引擎例外（on）：自動退回舊路徑，回應與 off 完全相同；資料也相同（沒有 flow 快照）', J(onExS.rec) === J(ref.rec) && onExS.data === ref.data);
      t('C10.4 on 熔斷寫 QUOTE_FLOW_ENGINE_FALLBACK，detail 含「熔斷」與例外碼，不含金額／成本', logsOf(onEx, 'QUOTE_FLOW_ENGINE_FALLBACK').length > 0 && logsOf(onEx, 'QUOTE_FLOW_ENGINE_FALLBACK').every((l) => /熔斷/.test(l[3]) && /FLOW_TEST_BOOM/.test(l[3]) && !/\d{6,}/.test(l[3])), (logsOf(onEx, 'QUOTE_FLOW_ENGINE_FALLBACK')[0] || [])[3]);
      restore();
      // 2) 引擎結果不一致：注入一個「預設路線少了總經理」的錯誤預設流程
      QF.legacyDefault = (flowCfg) => {
        const d = JSON.parse(JSON.stringify(realLegacyDefault(flowCfg)));
        d.routes.find((r) => r.id === 'r_l2').path = ['n_mgr1'];   // 錯：毛利層級 2 應為 mgr1>gm
        return d;
      };
      const shMis = await runVariant(ROOT, 'V1', 'shadow'), shMisS = finalState(shMis);
      t('C10.5 注入「少一關」的錯誤預設流程（shadow）：舊路徑為準，回應／資料與 off 完全相同', J(shMisS.rec) === J(ref.rec) && shMisS.data === ref.data);
      const mm = logsOf(shMis, 'QUOTE_FLOW_SHADOW_MISMATCH');
      t('C10.6 shadow 偵測到不一致 → 寫 QUOTE_FLOW_SHADOW_MISMATCH（有 actualTiers 差異、新舊關卡、歸類／BU，不含金額／成本）並計數', mm.length > 0 && mm.every((l) => /差異欄位=.*(actualTiers|financialTiers)/.test(l[3]) && /舊：關卡 .*新：關卡 /.test(l[3]) && !/\d{5,}/.test(l[3])) && shMis.internal.flowReporter.stats.mismatches >= mm.length, mm.length + ' ' + (mm[0] || [])[3]);
      t('C10.7 稽核節流：同一張單同階段重複不一致只寫第一筆（preview 與 submit 各一筆）', (() => { const keys = mm.map((l) => l[2]); return keys.length > 0 && new Set(keys).size < keys.length + 1 && mm.length <= 2 * RQUOTES().length; })(), mm.length);
      const onMis = await runVariant(ROOT, 'V1', 'on'), onMisS = finalState(onMis);
      const lvl2 = onMis.data.quotations.filter((q) => q.approval && q.approval.derived && J(q.approval.derived.tiers) === J(['mgr1', 'gm']));
      t('C10.8 on＋預設流程不一致：不一致的單（level 2 路徑）自動退回舊路徑、沒有 flow 快照並寫 FALLBACK；其餘單照常採用；所有回應與去掉 flow 後的資料與 off 完全相同', J(onMisS.rec) === J(ref.rec) && J(stripFlow(JSON.parse(onMisS.data))) === J(stripFlow(JSON.parse(ref.data))) && lvl2.length > 0 && lvl2.every((q) => !q.approval.flow) && onMis.data.quotations.some((q) => q.approval && q.approval.flow) && logsOf(onMis, 'QUOTE_FLOW_ENGINE_FALLBACK').length > 0 && logsOf(onMis, 'QUOTE_FLOW_SHADOW_MISMATCH').length > 0, 'lvl2=' + lvl2.length + ' fb=' + logsOf(onMis, 'QUOTE_FLOW_ENGINE_FALLBACK').length + ' mm=' + logsOf(onMis, 'QUOTE_FLOW_SHADOW_MISMATCH').length);
      restore();
      // 3) 引擎結果不一致、但原因是「放寬阻擋」：注入讓 NO_GM 消失
      QF.plan = (def, facts, io) => { const p = realPlan(def, facts, io); return Object.assign({}, p, { blockers: [] }); };
      const onNoBlock = await runVariant(ROOT, 'V3', 'on');
      t('C10.9 注入「引擎吃掉阻擋」：on 仍以舊路徑為準，NO_GM 照擋（V3 與 off 的回應相同）', J(finalState(onNoBlock).rec) === J(finalState(offRuns.V3).rec));
      restore();
      // 4) 時間：引擎很慢也不影響結果（同步）— 這裡只確認 shadow 不改變 steps 指派
      const stepsOf = (env) => env.data.quotations.filter((q) => q.approval && q.approval.steps).map((q) => q.id + ':' + q.approval.steps.map((s) => s.tier + '=' + (s.assignee || '')).join('>')).join('|');
      t('C10.10 shadow／on／off 三種模式送簽出的 steps（含指派人）完全相同', stepsOf(offRuns.V2) === stepsOf(shadowRuns.V2) && stepsOf(offRuns.V2) === stepsOf(onRuns.V2));
    } finally { restore(); }
  }
  // 後台唯讀狀態端點
  {
    const ST = '/api/admin/quote-flow/engine-status';
    const env = mkRouteEnv(ROOT, 'V2');
    const hs = [].concat(...env.routeHandlers('GET', ST));
    t('C12.1 engine-status 路由掛在 requireAdmin 之後（僅管理員）；沒有 PUT／POST／DELETE 版本（唯讀）', hs[0] === env.requireAdmin && !env.routeHandlers('PUT', ST) && !env.routeHandlers('POST', ST) && !env.routeHandlers('DELETE', ST));
    env.logs.push(['QUOTE_FLOW_SHADOW_MISMATCH', 'u', 't', 'd'], ['QUOTE_FLOW_SHADOW_MISMATCH', 'u', 't', 'd'], ['QUOTE_FLOW_SHADOW_ERROR', 'u', 't', 'd'], ['OTHER', 'u', 't', 'd']);
    const rOff = await withEnvMode(undefined, () => env.call('admin1', 'GET', ST));
    const rSh = await withEnvMode('shadow', () => env.call('admin1', 'GET', ST));
    const rOn = await withEnvMode('on', () => env.call('admin1', 'GET', ST));
    const rBad = await withEnvMode('banana', () => env.call('admin1', 'GET', ST));
    t('C12.2 狀態端點回報目前模式（off／shadow／on），亂碼值視為 off 並標示 envValid=false', rOff.s === 200 && rOff.j.mode === 'off' && rSh.j.mode === 'shadow' && rOn.j.mode === 'on' && rBad.j.mode === 'off' && rBad.j.envValid === false && rBad.j.envRaw === 'banana' && rOff.j.envValid === true && rOff.j.modeSource === 'env');
    t('C12.3 近 30 天筆數來自稽核（不一致 2、例外 1、熔斷 0）', rOff.j.audit.mismatch === 2 && rOff.j.audit.error === 1 && rOff.j.audit.fallback === 0 && /^\d{4}-\d{2}-\d{2}$/.test(rOff.j.audit.from), J(rOff.j.audit));
    const df = rOff.j.defaultFlow;
    t('C12.4 預設流程摘要：結構雜湊＝釘選值；V2 設定下啟用二級主管的 BU＝ITS、跳過一級主管例外 1 人；4 條路線、2 個調整、核決表 9 列', df.structureHash === QF.DEFAULT_FLOW_HASH && df.pinnedHash === QF.DEFAULT_FLOW_HASH && J(df.mgr2Bus) === J(['ITS']) && df.skipUserCount === 1 && df.routes.length === 4 && df.modifiers.length === 2 && df.table.length === 9 && df.nodes.length === 5, J(df).slice(0, 200));
    const body = J(rOff.j);
    t('C12.5 回應不含帳號名單、金額、成本、flow 快照（例外名單的帳號 m2a 不外洩，只回人數）', !/m2a|revenue|costCents|gpCents|marginText|"flow"/.test(body.replace(/"m2a"/g, '')) && !body.includes('m2a'));
    const noStore = QF.describeDef(QF.legacyDefault(env.internal.getFlowCfg(env.data)));
    t('C12.6 describeDef 摘要可讀：路線條件含「折後未稅金額大於 5,000 萬」「毛利層級等於 3」', noStore.routes[0].when.includes('折後未稅金額 大於 5,000 萬') && noStore.routes[1].when.includes('1,000 萬') && noStore.routes[1].when.includes('毛利層級 等於 3'), J(noStore.routes.map((r) => r.when)));
  }
  // C13 不就地修改輸入（Postgres 的 db.load() 是共用快取：引擎若就地改了 q／ctx／derived，會污染其他請求）；C14 同一張單同時兩個送簽請求只有一個成功
  {
    const env = mkRouteEnv(ROOT, 'V2');
    const users = {}; env.auth.users.forEach((u) => { users[u.username] = u; });
    const cfg = env.internal.getCfg(env.data);
    const q = JSON.parse(J(env.data.quotations.find((x) => x.id === 'a1')));
    const derived = QA.buildDerived(q, cfg.productClasses);
    const ctx = { data: env.data, users, cfg, me: 'own1' };
    const snap = () => J({ ctx, q, derived });
    const before = snap();
    let okAll = true, results = [];
    for (const mode of ['shadow', 'on']) {
      const r = await withEnvMode(mode, async () => env.internal.flowCheck(derived, q, ctx, 'test', mode, null));
      results.push(!!r);
      if (snap() !== before) okAll = false;
    }
    t('C13 shadow／on 的 flowCheck 不就地修改 q／ctx（data、users、cfg）／derived（前後深度比對相同；引擎結果可用）', okAll && results.every(Boolean), results.join(','));
    const dd = JSON.parse(J(derived)); const roster = JSON.parse(J(cfg.roster));
    Object.freeze(q); Object.freeze(derived); Object.freeze(ctx.users); Object.freeze(ctx.cfg); Object.freeze(ctx.cfg.roster); Object.freeze(ctx.data);
    let thrown = null; try { await withEnvMode('on', async () => env.internal.flowCheck(derived, q, ctx, 'test', 'on', null)); } catch (e) { thrown = e; }
    t('C13b 輸入物件全部淺層凍結後 on 模式仍可運作（沒有任何寫入嘗試造成例外）', thrown === null && J(derived) === J(dd) && J(cfg.roster) === J(roster), thrown && thrown.message);
    const env2 = mkRouteEnv(ROOT, 'V1');
    const rs = await withEnvMode('on', () => Promise.all([env2.call('own2', 'POST', SUBMIT, { id: 'a2' }, {}), env2.call('own2', 'POST', SUBMIT, { id: 'a2' }, {})]));
    t('C14 on：同一張單同時兩個送簽請求 → 一個 200、一個 409 ALREADY_PENDING；只有一份 flow 快照、歷程只有一筆 SUBMIT', J(rs.map((r) => r.s).sort()) === J([200, 409]) && rs.some((r) => r.j && r.j.code === 'ALREADY_PENDING') && Q(env2, 'a2').approval.history.filter((h) => h.action === 'SUBMIT').length === 1 && !!Q(env2, 'a2').approval.flow, J(rs.map((r) => r.s)));
  }
  // 環境變數解析（路由層）
  {
    const modes = [[undefined, 'off'], ['', 'off'], ['off', 'off'], ['banana', 'off'], ['1', 'off'], [' Shadow ', 'shadow'], ['ON', 'on']];
    let allOk = true; const bad = [];
    for (const [raw, exp] of modes) {
      const env = await runVariant(ROOT, 'V1', raw);
      const hasFlow = env.data.quotations.some((q) => q.approval && q.approval.flow);
      const ran = env.internal.flowReporter.stats.runs > 0;
      const expFlow = exp === 'on', expRan = exp !== 'off';
      if (hasFlow !== expFlow || ran !== expRan) { allOk = false; bad.push(String(raw) + '→flow=' + hasFlow + ' ran=' + ran); }
    }
    t('C11 QUOTE_FLOW_ENGINE 解析：未設／空／off／亂碼／"1" ＝ off（引擎完全不執行、無 flow）；shadow（大小寫空白不拘）＝只比對；ON＝採用並凍結', allOk, bad.join(';'));
  }

  // ═════ E 變異測試 ═════
  const quick = { strideF: 60, strideG1: 90, strideG2: 400, strideG3: 500, nRandom: 4000, seed: SEED };
  if (SKIP.has('E')) MUTATIONS.length = 0;
  const sanity = core(QF, quick);
  t('E0 變異測試的縮小網格對「未改壞」的引擎全數通過（' + sanity.total + ' 筆）', sanity.fails === 0 && sanity.thrown === 0, sanity.first.join('\n  '));
  let killed = 0; const survived = []; const equivs = [];
  for (const [name, find, repl] of MUTATIONS) {
    if (SRC.split(find).length !== 2) { t('E.' + name + '：變異點在原始碼中必須唯一（找到 ' + (SRC.split(find).length - 1) + ' 處）', false); continue; }
    let mut, ast = { fails: 0 };
    try {
      mut = loadMutant(SRC.replace(find, () => repl));
      const r = core(mut, quick);
      const astRes = []; try { astTests(mut, (n, ok) => { if (!ok) astRes.push(n); }); } catch (e) { astRes.push('throw'); }
      ast = { fails: astRes.length };
      const isKilled = r.fails > 0 || r.thrown > 0 || ast.fails > 0;
      if (name.startsWith('EQUIV ')) { equivs.push(name.slice(6) + (isKilled ? '（竟被抓到）' : '（存活＝在護欄範圍內與精確算法等價）')); continue; }
      if (isKilled) { killed++; t('E.' + name + ' → 測試變紅（等價不符 ' + r.fails + '／引擎例外 ' + r.thrown + '／AST 檢查 ' + ast.fails + '）', true); }
      else { survived.push(name); t('E.' + name + ' → 測試變紅', false, '變異體存活'); }
    } catch (e) {
      // 載入／執行時整個炸掉也算被抓到（測試不可能是綠的）
      killed++; t('E.' + name + ' → 測試變紅（變異體執行即例外：' + String(e && e.message).slice(0, 60) + '）', true);
    }
  }
  const nonEquiv = MUTATIONS.filter((m) => !m[0].startsWith('EQUIV ')).length;
  t('E 變異測試：' + nonEquiv + ' 種變異全部被抓到（存活 ' + survived.length + '）；另 ' + equivs.length + ' 種為等價變異體：' + equivs.join('、'), survived.length === 0 && killed === nonEquiv, survived.join('；'));
  {
    // 等價變異體的佐證：1e15 < 2^53、門檻為整數百分比 → 浮點算法在護欄範圍內與精確算法分類相同。隨機搜尋 40 萬個貼近門檻的點，沒有找到分類不同的反例。
    const R2 = rng(SEED + 7); let diff = 0, n = 0;
    for (const th of [5, 10, 15, 25]) for (let i = 0; i < 100000; i++) {
      const rev = Math.max(1, Math.floor(R2() < 0.5 ? 1e15 - R2() * 1e12 : R2() * 1e15)); const g = Number((BigInt(th) * BigInt(rev) + 99n) / 100n);
      for (const gp of [g - 1, g]) { n++; const ex = BigInt(gp) * 100n >= BigInt(th) * BigInt(rev); if (ex !== (Number(gp) / Number(rev) * 100 >= th) || ex !== (Number(BigInt(gp) * 100n) >= th * Number(rev))) diff++; }
    }
    t('E.float 等價變異體佐證：' + n + ' 個貼近門檻的點，毛利浮點算法與精確算法的分類不一致 ' + diff + ' 個（0＝在 1e15 分上限內浮點恰好夠用；精確算法是必要的防線而非必然的差異）', diff === 0, diff);
  }
  // 路由層變異：diffPlans 失效（永遠回報一致）→ 故障注入測試必須變紅
  {
    const realDiff = QF.diffPlans, realLD = QF.legacyDefault;
    try {
      QF.diffPlans = () => [];
      QF.legacyDefault = (flowCfg) => { const d = JSON.parse(JSON.stringify(realLD(flowCfg))); d.routes.find((r) => r.id === 'r_l2').path = ['n_mgr1']; return d; };
      const env = await runVariant(ROOT, 'V1', 'shadow');
      const caught = env.logs.some((l) => l[0] === 'QUOTE_FLOW_SHADOW_MISMATCH');
      t('E.route diffPlans 永遠回報一致 → 故障注入的「不一致」稽核消失（故障注入測試 C10.6 會變紅）', caught === false);
    } finally { QF.diffPlans = realDiff; QF.legacyDefault = realLD; }
  }

  // ═════ F off 對基底 ═════
  if (BASE_ROOT) {
    const baseOk = fs.existsSync(path.join(BASE_ROOT, 'lib/quoteRoutes.js')) && !fs.existsSync(path.join(BASE_ROOT, 'lib/quoteFlow.js'));
    t('F0 基底目錄是未含引擎的程式（有 quoteRoutes.js、沒有 quoteFlow.js）', baseOk, BASE_ROOT);
    if (baseOk) {
      for (const v of VARS_) {
        const b = finalState(await runVariant(BASE_ROOT, v, undefined)), w = finalState(offRuns[v]);
        t('F1.' + v + ' off 對基底：' + b.rec.length + ' 個回應、儲存資料、稽核、通知全部逐位元相同', J(b.rec) === J(w.rec) && b.data === w.data && b.logs === w.logs && b.notes === w.notes, J(b.rec.filter((r, i) => J(r) !== J(w.rec[i])).slice(0, 1)).slice(0, 300));
        const bOn = finalState(await runVariant(BASE_ROOT, v, 'on'));   // 基底不認得這個環境變數：設了也不能有任何差別
        t('F2.' + v + ' 基底即使設了 QUOTE_FLOW_ENGINE=on 也無差別（確認基底確實沒有引擎）', J(bOn.rec) === J(b.rec) && bOn.data === b.data);
      }
    }
  } else t('F off 對基底：未指定 QFE_BASE_ROOT（略過；請以 git archive 出基底後再跑）', true, 'SKIP');

  if (reg0 === undefined) delete process.env.QUOTE_FLOW_ENGINE; else process.env.QUOTE_FLOW_ENGINE = reg0;
  let pass = 0, fail = 0;
  for (const [n, ok, x] of res) { if (ok) pass++; else fail++; if (!ok || process.env.QFE_VERBOSE) console.log((ok ? 'PASS ' : 'FAIL ') + n + (x && !ok ? '  <- ' + String(x).slice(0, 600) : '')); }
  console.log('\n簽核流程引擎 M1 檢查：PASS ' + pass + ' / FAIL ' + fail + '（共 ' + res.length + ' 項；耗時 ' + ((Date.now() - T0) / 1000).toFixed(1) + 's）' + (QUICK ? '  [QUICK 模式：網格縮小，不比對 GOLDEN]' : ''));
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error('檢查腳本例外', e && e.stack || e); process.exit(2); });
