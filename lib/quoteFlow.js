// 報價單簽核「流程引擎」M1 — 純函式（不碰資料庫／req／時間），只依賴 node 內建 crypto 與 ./quoteApproval
// 規格：scratchpad/flowdesigner_SPEC.md 第 3、4 章。M1 範圍＝流程定義 schema＋驗證、JSON AST 條件（不 eval）、
// 「預設流程」產生器（由現行核決表／金額門檻／路徑表／BU 二級主管設定轉出）、評估引擎、與舊程式等價的「簽核人可行性檢查」。
//
// 硬性限制（M1）
//  1. 不新增任何關卡 tier 字串：節點的 tier 只能是 mgr2／mgr1／gm／chairman／board（信件／通知／權限都只認這五種）。
//  2. 行為必須與舊程式（quoteApproval.requiredPath ＋ quoteRoutes.buildActualTiers ＋ chainBlockers）完全等價；
//     舊程式的怪行為（只取業主第一個 BU、mgr2 找不到人時靜默不插入、NO_MGR2 實際不可達…）在這裡忠實複製，不修。
//  3. 條件只用 JSON AST ＋ 變數白名單，沒有 eval／new Function／正規表示式輸入；金額用 BigInt（分），不用浮點。
//  4. 同一人多關：維持現狀（semantics.sameApprover='forbid'）——不可行就擋送簽（SIGNER_OVERLAP），不自動略過。
//  5. 代核／秘書／名冊的語意不動：board 節點是 locked 的 proxy 節點，候選人＝名冊 boardProxy ∪ 秘書角色（由呼叫端的 resolver 提供）。
const crypto = require('crypto');
const QA = require('./quoteApproval');

const SCHEMA = 1;
const ENGINE_ID = 'qfe-1';
const TIER_KEYS = Object.freeze(Object.keys(QA.TIERS));   // mgr2, mgr1, gm, chairman, board（M1 不新增）
const MAX_STEPS_CAP = 8;
const MAX_CENTS_BIG = BigInt(QA.MAX_CENTS_NUM);

// 條件 AST 的資源上限（防深度炸彈／超大輸入）
const LIM = Object.freeze({ depth: 8, nodes: 64, arr: 64, list: 500, str: 100, nodesInDef: 16, routes: 32, modifiers: 32 });

// 每個層級允許的節點形狀（M1 只開放現有五種；tier 決定 type／approver，不能自訂）
const NODE_KIND = Object.freeze({
  mgr2: { type: 'approval', approver: { kind: 'owner_direct_supervisor', role: 'manager2' }, reassignable: true, locked: false },
  mgr1: { type: 'approval', approver: { kind: 'supervisor_chain', role: 'manager1' }, reassignable: true, locked: false },
  gm: { type: 'approval', approver: { kind: 'roster', name: 'gm', mode: 'any_one' }, reassignable: false, locked: false },
  chairman: { type: 'approval', approver: { kind: 'roster', name: 'chairman', mode: 'any_one' }, reassignable: false, locked: false },
  board: { type: 'proxy', approver: { kind: 'roster_plus_role', roster: 'boardProxy', role: 'secretary' }, reassignable: false, locked: true },
});

class FlowError extends Error {
  constructor(code, message) { super(message || code); this.name = 'FlowError'; this.code = code; }
}

const has = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
const isPlain = (x) => {
  if (x === null || typeof x !== 'object' || Array.isArray(x)) return false;
  const p = Object.getPrototypeOf(x);
  return p === Object.prototype || p === null;
};
function canon(v) {
  if (Array.isArray(v)) return '[' + v.map(canon).join(',') + ']';
  if (v && typeof v === 'object') return '{' + Object.keys(v).sort().map((k) => JSON.stringify(k) + ':' + canon(v[k])).join(',') + '}';
  return JSON.stringify(v === undefined ? null : v);
}
function deepFreeze(o) {
  if (o && typeof o === 'object' && !Object.isFrozen(o)) { Object.freeze(o); Object.keys(o).forEach((k) => deepFreeze(o[k])); }
  return o;
}

/** 整棵樹都已凍結才算（淺層 Object.freeze 後內層仍可改）。有 getter／setter 的屬性視為不可信（每次讀可能回不同物件）。循環參照安全。 */
function isDeepFrozen(o, seen) {
  if (o === null || typeof o !== 'object') return true;
  seen = seen || new Set();
  if (seen.has(o)) return true;
  seen.add(o);
  if (!Object.isFrozen(o)) return false;
  const d = Object.getOwnPropertyDescriptors(o);
  for (const k of Reflect.ownKeys(d)) {
    const pd = d[k];
    if (pd.get || pd.set) return false;
    if (!isDeepFrozen(pd.value, seen)) return false;
  }
  return true;
}
/** 保留字 id：Object.prototype 上有的名字（__proto__、constructor、toString、hasOwnProperty…）一律不准當節點／路線／調整的 id（防原型鏈查表誤判；M2 起 id 來自儲存資料）。 */
const reservedId = (id) => typeof id === 'string' && (id in Object.prototype || id === 'prototype');

// ───────────────────────── 引擎模式（環境變數；熔斷器）─────────────────────────
/** QUOTE_FLOW_ENGINE：off（預設）｜shadow｜on。大小寫與前後空白不拘；無法辨識的值一律視為 off（fail-safe）。 */
function engineModeInfo(raw) {
  const s = (raw === undefined || raw === null) ? '' : String(raw).trim().toLowerCase();
  if (s === 'shadow' || s === 'on') return { mode: s, raw: String(raw), valid: true };
  if (s === 'off' || s === '') return { mode: 'off', raw: raw === undefined || raw === null ? '' : String(raw), valid: true };
  return { mode: 'off', raw: String(raw).slice(0, 40), valid: false };
}
const engineMode = (raw) => engineModeInfo(raw).mode;

// ───────────────────────── 十進位解析（只接受十進位字串）─────────────────────────
const MONEY_RE = /^(0|[1-9][0-9]{0,15})(?:\.([0-9]{1,2}))?$/;
const PCT_RE = /^(-?)(0|[1-9][0-9]{0,2})(?:\.([0-9]{1,4}))?$/;
/** 金額字串（元，至多 2 位小數）→ 分（BigInt）。不是十進位字串或超過上限就丟 FlowError。 */
function parseMoneyCents(s) {
  if (typeof s !== 'string') throw new FlowError('FLOW_BAD_COND', '金額必須是十進位字串');
  const m = MONEY_RE.exec(s);
  if (!m) throw new FlowError('FLOW_BAD_COND', '金額格式無效');
  const c = BigInt(m[1]) * 100n + BigInt((m[2] || '').padEnd(2, '0') || '0');
  if (c > MAX_CENTS_BIG) throw new FlowError('FLOW_BAD_COND', '金額超過上限');
  return c;
}
/** 百分比字串（至多 4 位小數）→ { n: BigInt（含正負號，已乘 10^s）, s } */
function parsePct(s) {
  if (typeof s !== 'string') throw new FlowError('FLOW_BAD_COND', '百分比必須是十進位字串');
  const m = PCT_RE.exec(s);
  if (!m) throw new FlowError('FLOW_BAD_COND', '百分比格式無效');
  const frac = m[3] || '';
  const n = BigInt(m[2] + frac);
  return { n: m[1] === '-' ? -n : n, s: frac.length };
}

// ───────────────────────── 變數白名單 ─────────────────────────
// type：money（分）｜pct（有理數）｜level（1..3）｜enum（固定值域）｜string（任意短字串）｜set（集合）
const VARS = Object.freeze({
  amount: { type: 'money', sensitive: false, label: '折後未稅金額' },
  marginLevel: { type: 'level', sensitive: true, label: '毛利層級' },
  marginPct: { type: 'pct', sensitive: true, label: '毛利率' },
  category: { type: 'enum', sensitive: false, label: '商品歸類列', values: QA.ROWS.map((r) => r.key) },
  classKeys: { type: 'set', sensitive: false, label: '商品類別', values: QA.CLASS_KEYS.slice() },
  ownerBuPrimary: { type: 'string', sensitive: false, label: '業務主要 BU' },
  owner: { type: 'string', sensitive: false, label: '指定業務' },
});
const OPS_NUM = ['==', '!=', '>', '>=', '<', '<='];
const OPS_STR = ['==', '!=', 'in', 'not_in'];
const OPS_SET = ['contains', 'not_contains'];

// ───────────────────────── 條件 AST：驗證 ─────────────────────────
const BAD = (msg) => ({ code: 'FLOW_BAD_COND', message: msg });
const ABORT = { abort: true };   // 超過資源上限時中止整個走訪（避免寬度×深度的組合爆炸）

function checkCmp(v, errs) {
  if (!Array.isArray(v) || v.length !== 3) { errs.push(BAD('cmp 必須是 [變數, 運算子, 值]')); return; }
  const [name, op, rhs] = v;
  if (typeof name !== 'string' || !has(VARS, name)) { errs.push(BAD('不在白名單的變數：' + String(name).slice(0, 30))); return; }
  const vd = VARS[name];
  if (typeof op !== 'string') { errs.push(BAD('運算子必須是字串')); return; }
  const strOk = (s) => typeof s === 'string' && s.length >= 1 && s.length <= LIM.str;
  switch (vd.type) {
    case 'money':
      if (!OPS_NUM.includes(op)) { errs.push(BAD('金額不支援運算子 ' + op)); return; }
      try { parseMoneyCents(rhs); } catch (e) { errs.push(BAD('金額值無效（需十進位字串，至多 2 位小數）')); }
      return;
    case 'pct':
      if (!OPS_NUM.includes(op)) { errs.push(BAD('毛利率不支援運算子 ' + op)); return; }
      try { parsePct(rhs); } catch (e) { errs.push(BAD('毛利率值無效（需十進位字串，至多 4 位小數）')); }
      return;
    case 'level':
      if (!OPS_NUM.includes(op)) { errs.push(BAD('層級不支援運算子 ' + op)); return; }
      if (!(typeof rhs === 'number' && Number.isInteger(rhs) && rhs >= 1 && rhs <= 3)) errs.push(BAD('毛利層級的值必須是 1、2 或 3'));
      return;
    case 'enum':
    case 'string': {
      if (!OPS_STR.includes(op)) { errs.push(BAD(vd.label + '不支援運算子 ' + op)); return; }
      const one = (s) => strOk(s) && (vd.type !== 'enum' || vd.values.includes(s));
      if (op === 'in' || op === 'not_in') {
        if (!Array.isArray(rhs) || rhs.length > LIM.list || !rhs.every(one)) errs.push(BAD(vd.label + '的清單無效'));
      } else if (!one(rhs)) errs.push(BAD(vd.label + '的值無效'));
      return;
    }
    case 'set':
      if (!OPS_SET.includes(op)) { errs.push(BAD(vd.label + '不支援運算子 ' + op)); return; }
      if (!(typeof rhs === 'string' && vd.values.includes(rhs))) errs.push(BAD(vd.label + '的值無效'));
      return;
    default:
      errs.push(BAD('未知變數型別'));
  }
}

function checkCond(x, ctx, depth, st, errs) {
  if (errs.length >= 20) return;
  if (++st.n > LIM.nodes) { errs.push(BAD('條件節點過多（上限 ' + LIM.nodes + '）')); throw ABORT; }
  if (depth > LIM.depth) { errs.push(BAD('條件巢狀過深（上限 ' + LIM.depth + ' 層）')); return; }
  if (typeof x === 'boolean') return;
  if (!isPlain(x)) { errs.push(BAD('條件必須是布林或單一運算子的物件')); return; }
  const keys = Object.keys(x);
  if (keys.length !== 1) { errs.push(BAD('條件物件必須恰有一個運算子')); return; }
  const k = keys[0];
  const v = x[k];
  switch (k) {
    case 'all':
    case 'any':
      if (!Array.isArray(v) || v.length < 1 || v.length > LIM.arr) { errs.push(BAD(k + ' 必須是 1～' + LIM.arr + ' 個條件的陣列')); return; }
      for (const c of v) checkCond(c, ctx, depth + 1, st, errs);
      return;
    case 'not':
      checkCond(v, ctx, depth + 1, st, errs);
      return;
    case 'cmp':
      checkCmp(v, errs);
      return;
    case 'applied':
      if (!ctx.modifier) { errs.push(BAD('applied 只能用在調整（modifier）的條件')); return; }
      if (typeof v !== 'string' || !ctx.modIdsBefore.has(v)) errs.push(BAD('applied 只能引用排在前面的調整'));
      return;
    case 'resolvable':
      if (!ctx.modifier) { errs.push(BAD('resolvable 只能用在調整（modifier）的條件')); return; }
      if (typeof v !== 'string' || !ctx.nodeIds.has(v)) errs.push(BAD('resolvable 引用的節點不存在'));
      return;
    case 'resolvedIn':
      if (!ctx.modifier) { errs.push(BAD('resolvedIn 只能用在調整（modifier）的條件')); return; }
      if (!Array.isArray(v) || v.length !== 2 || typeof v[0] !== 'string' || !ctx.nodeIds.has(v[0])) { errs.push(BAD('resolvedIn 必須是 [節點, 帳號清單]')); return; }
      if (!['mgr2', 'mgr1'].includes(ctx.nodeTier[v[0]])) { errs.push(BAD('resolvedIn 只能用在單一指定人的節點（mgr2／mgr1）')); return; }
      if (!Array.isArray(v[1]) || v[1].length > LIM.list || !v[1].every((s) => typeof s === 'string' && s.length >= 1 && s.length <= LIM.str)) errs.push(BAD('resolvedIn 的帳號清單無效'));
      return;
    default:
      errs.push(BAD('未知的條件運算子：' + String(k).slice(0, 30)));
  }
}

/** 驗證單一條件（供單元測試與設計器使用）。ctx 預設為「路線」環境（不允許 applied／resolvable／resolvedIn）。 */
function validateCond(ast, ctx) {
  const errs = [];
  const c = Object.assign({ modifier: false, nodeIds: new Set(), nodeTier: Object.create(null), modIdsBefore: new Set() }, ctx || {});
  try { checkCond(ast, c, 1, { n: 0 }, errs); } catch (e) { if (e !== ABORT) errs.push(BAD('條件結構無法檢查')); }
  return { ok: errs.length === 0, errors: errs };
}

// ───────────────────────── 條件 AST：求值 ─────────────────────────
function cmpOp(op, c) {
  switch (op) {
    case '==': return c === 0;
    case '!=': return c !== 0;
    case '>': return c > 0;
    case '>=': return c >= 0;
    case '<': return c < 0;
    case '<=': return c <= 0;
    default: throw new FlowError('FLOW_BAD_COND', '未知比較運算子');
  }
}
const cmpBig = (a, b) => (a < b ? -1 : a > b ? 1 : 0);

function evalCmp(v, env) {
  if (!Array.isArray(v) || v.length !== 3 || typeof v[0] !== 'string' || !has(VARS, v[0])) throw new FlowError('FLOW_BAD_COND', '條件無效');
  const [name, op, rhs] = v;
  const f = env.facts;
  switch (VARS[name].type) {
    case 'money': return cmpOp(op, cmpBig(f.amountCents, parseMoneyCents(rhs)));
    case 'pct': {
      const p = parsePct(rhs);
      // gp*100/rev  ⋚  p.n/10^s  ⇔  gp*100*10^s ⋚ p.n*rev（rev>0，makeFacts 已確保）
      return cmpOp(op, cmpBig(f.gpCents * 100n * (10n ** BigInt(p.s)), p.n * f.amountCents));
    }
    case 'level': {
      if (!(typeof rhs === 'number' && Number.isInteger(rhs))) throw new FlowError('FLOW_BAD_COND', '層級值無效');
      return cmpOp(op, cmpBig(BigInt(env.level()), BigInt(rhs)));
    }
    case 'enum':
    case 'string': {
      const val = name === 'category' ? f.category : name === 'ownerBuPrimary' ? f.ownerBuPrimary : f.owner;
      if (op === '==') return val === rhs;
      if (op === '!=') return val !== rhs;
      if (op === 'in' || op === 'not_in') {
        if (!Array.isArray(rhs)) throw new FlowError('FLOW_BAD_COND', '清單無效');
        const hit = rhs.includes(val);
        return op === 'in' ? hit : !hit;
      }
      throw new FlowError('FLOW_BAD_COND', '未知運算子');
    }
    case 'set': {
      const hit = f.classKeys.includes(rhs);
      if (op === 'contains') return hit;
      if (op === 'not_contains') return !hit;
      throw new FlowError('FLOW_BAD_COND', '未知運算子');
    }
    default: throw new FlowError('FLOW_BAD_COND', '未知變數型別');
  }
}

function evalCond(x, env, depth, st) {
  depth = depth || 1;
  st = st || { n: 0 };
  if (++st.n > LIM.nodes) throw new FlowError('FLOW_BAD_COND', '條件節點過多');
  if (depth > LIM.depth) throw new FlowError('FLOW_BAD_COND', '條件巢狀過深');
  if (x === true) return true;
  if (x === false) return false;
  if (!isPlain(x)) throw new FlowError('FLOW_BAD_COND', '條件結構無效');
  const keys = Object.keys(x);
  if (keys.length !== 1) throw new FlowError('FLOW_BAD_COND', '條件物件必須恰有一個運算子');
  const k = keys[0];
  const v = x[k];
  switch (k) {
    case 'all': {
      if (!Array.isArray(v) || v.length < 1 || v.length > LIM.arr) throw new FlowError('FLOW_BAD_COND', 'all 無效');
      for (const c of v) if (!evalCond(c, env, depth + 1, st)) return false;
      return true;
    }
    case 'any': {
      if (!Array.isArray(v) || v.length < 1 || v.length > LIM.arr) throw new FlowError('FLOW_BAD_COND', 'any 無效');
      for (const c of v) if (evalCond(c, env, depth + 1, st)) return true;
      return false;
    }
    case 'not': return !evalCond(v, env, depth + 1, st);
    case 'cmp': return evalCmp(v, env);
    case 'applied':
      if (!env.modifier || typeof v !== 'string') throw new FlowError('FLOW_BAD_COND', 'applied 位置無效');
      return env.applied.has(v);
    case 'resolvable': {
      if (!env.modifier || typeof v !== 'string') throw new FlowError('FLOW_BAD_COND', 'resolvable 位置無效');
      const r = env.resolveNode(v);
      return r.person ? true : r.persons.length > 0;
    }
    case 'resolvedIn': {
      if (!env.modifier || !Array.isArray(v) || v.length !== 2 || typeof v[0] !== 'string' || !Array.isArray(v[1])) throw new FlowError('FLOW_BAD_COND', 'resolvedIn 無效');
      const r = env.resolveNode(v[0]);
      return !!r.person && v[1].includes(r.person);
    }
    default: throw new FlowError('FLOW_BAD_COND', '未知的條件運算子');
  }
}

// ───────────────────────── 決策表（毛利層級）─────────────────────────
/**
 * 毛利層級（1=一級主管 2=總經理 3=董事長）。演算法與 quoteApproval.requiredPath 的第 330–344 行逐句相同：
 * fixedLevel 優先；l1 非 null 且 gp*100 >= l1*rev → 1；gp*100 >= l2*rev → 2；否則 3（壓線取低層）。
 * 找不到該類別的列 → 用 other 列（同 requiredPath 的 has(ROW_BY_KEY,…) ? … : other）。
 */
function marginLevelOf(table, cat, gpCents, revCents) {
  let row = null, other = null;
  for (const r of table) { if (r.cat === cat) row = r; if (r.cat === 'other') other = r; }
  row = row || other;
  if (!row) throw new FlowError('FLOW_BAD_TABLE', '決策表缺少 other 列');
  if (row.fixedLevel) return row.fixedLevel;
  const gp100 = gpCents * 100n;
  if (row.l1 !== null && gp100 >= BigInt(row.l1) * revCents) return 1;
  if (gp100 >= BigInt(row.l2) * revCents) return 2;
  return 3;
}

// ───────────────────────── 流程定義：驗證 ─────────────────────────
const DERR = (code, message, at) => ({ code, message, at: at || '' });

function validateFlowDef(def) {
  const errs = [];
  const push = (code, msg, at) => { if (errs.length < 30) errs.push(DERR(code, msg, at)); };
  try {
    if (!isPlain(def)) return { ok: false, errors: [DERR('FLOW_BAD_SCHEMA', '流程定義必須是物件')] };
    if (def.schema !== SCHEMA) push('FLOW_BAD_SCHEMA', 'schema 必須是 ' + SCHEMA, 'schema');
    if (typeof def.id !== 'string' || !def.id || def.id.length > LIM.str) push('FLOW_BAD_SCHEMA', 'id 無效', 'id');
    if (!Number.isInteger(def.no) || def.no < 0) push('FLOW_BAD_SCHEMA', 'no 必須是非負整數', 'no');
    if (def.origin !== 'legacy' && def.origin !== 'designer') push('FLOW_BAD_SCHEMA', 'origin 無效', 'origin');
    const maxSteps = def.limits && def.limits.maxSteps;
    if (!Number.isInteger(maxSteps) || maxSteps < 1 || maxSteps > MAX_STEPS_CAP) push('FLOW_BAD_SCHEMA', 'limits.maxSteps 必須是 1～' + MAX_STEPS_CAP, 'limits.maxSteps');
    if (!def.semantics || def.semantics.sameApprover !== 'forbid') push('FLOW_BAD_SCHEMA', 'semantics.sameApprover 目前只支援 forbid', 'semantics');

    // 決策表
    const tbl = def.tables && def.tables.margin;
    if (!Array.isArray(tbl) || tbl.length < 1 || tbl.length > 32) push('FLOW_BAD_TABLE', 'tables.margin 無效', 'tables.margin');
    else {
      const seen = new Set();
      tbl.forEach((r, i) => {
        const at = 'tables.margin[' + i + ']';
        if (!isPlain(r) || typeof r.cat !== 'string' || !VARS.category.values.includes(r.cat)) { push('FLOW_BAD_TABLE', '決策表列的類別無效', at); return; }
        if (seen.has(r.cat)) push('FLOW_BAD_TABLE', '決策表類別重複：' + r.cat, at);
        seen.add(r.cat);
        const pctOk = (x) => x === null || (Number.isInteger(x) && x >= 0 && x <= 100);
        if (!pctOk(r.l1) || !pctOk(r.l2)) push('FLOW_BAD_TABLE', 'l1／l2 必須是 0～100 的整數或 null', at);
        if (r.fixedLevel !== undefined && !(Number.isInteger(r.fixedLevel) && r.fixedLevel >= 1 && r.fixedLevel <= 3)) push('FLOW_BAD_TABLE', 'fixedLevel 必須是 1～3', at);
        if (!r.fixedLevel && r.l2 === null) push('FLOW_BAD_TABLE', '沒有 fixedLevel 的列 l2 不可為 null', at);
      });
      if (!seen.has('other')) push('FLOW_BAD_TABLE', '決策表必須有 other 列', 'tables.margin');
    }

    // 節點
    const nodeIds = new Set();
    const nodeTier = Object.create(null);
    const nodeById = Object.create(null);
    const tiersSeen = new Set();
    if (!Array.isArray(def.nodes) || def.nodes.length < 1 || def.nodes.length > LIM.nodesInDef) push('FLOW_NODE_MISSING', 'nodes 無效', 'nodes');
    else {
      def.nodes.forEach((n, i) => {
        const at = 'nodes[' + i + ']';
        if (!isPlain(n) || typeof n.id !== 'string' || !n.id || n.id.length > LIM.str) { push('FLOW_BAD_NODE', '節點 id 無效', at); return; }
        if (reservedId(n.id)) { push('FLOW_BAD_NODE', '節點 id 是保留字：' + n.id, at); return; }
        if (nodeIds.has(n.id)) { push('FLOW_DUP_NODE', '節點 id 重複：' + n.id, at); return; }
        nodeIds.add(n.id);
        nodeById[n.id] = n;
        if (typeof n.tier !== 'string' || !TIER_KEYS.includes(n.tier)) { push('FLOW_BAD_NODE', '節點 tier 必須是現有五種關卡之一（M1 不新增 tier）', at); return; }
        nodeTier[n.id] = n.tier;
        if (tiersSeen.has(n.tier)) push('FLOW_DUP_NODE', '同一個 tier 只能有一個節點：' + n.tier, at);
        tiersSeen.add(n.tier);
        const k = NODE_KIND[n.tier];
        if (n.type !== k.type) push('FLOW_BAD_NODE', 'tier ' + n.tier + ' 的節點 type 必須是 ' + k.type, at);
        if (canon(n.approver) !== canon(k.approver)) push('FLOW_BAD_NODE', 'tier ' + n.tier + ' 的 approver 不可自訂（M1）', at);
        if (n.locked !== k.locked) push('FLOW_BAD_NODE', 'tier ' + n.tier + ' 的 locked 必須是 ' + k.locked, at);
        if (n.reassignable !== k.reassignable) push('FLOW_BAD_NODE', 'tier ' + n.tier + ' 的 reassignable 必須是 ' + k.reassignable, at);
        if (typeof n.label !== 'string' || !n.label || n.label.length > LIM.str) push('FLOW_BAD_NODE', '節點 label 無效', at);
      });
    }

    // 路線
    const condCtxRoute = { modifier: false, nodeIds, nodeTier, modIdsBefore: new Set() };
    if (!Array.isArray(def.routes) || def.routes.length < 1 || def.routes.length > LIM.routes) push('FLOW_NO_DEFAULT', 'routes 無效', 'routes');
    else {
      const rids = new Set();
      def.routes.forEach((r, i) => {
        const at = 'routes[' + i + ']';
        if (!isPlain(r) || typeof r.id !== 'string' || !r.id || r.id.length > LIM.str) { push('FLOW_BAD_SCHEMA', '路線 id 無效', at); return; }
        if (reservedId(r.id)) { push('FLOW_BAD_SCHEMA', '路線 id 是保留字', at); return; }
        if (rids.has(r.id)) push('FLOW_BAD_SCHEMA', '路線 id 重複', at);
        rids.add(r.id);
        if (!Array.isArray(r.path) || r.path.length < 1) { push('FLOW_EMPTY_PATH', '路線的關卡序列不可為空', at); return; }
        if (r.path.length > (maxSteps || MAX_STEPS_CAP)) push('FLOW_TOO_LONG', '路線關卡數超過上限', at);
        const ps = new Set();
        r.path.forEach((id) => {
          if (typeof id !== 'string' || !nodeIds.has(id)) push('FLOW_NODE_MISSING', '路線引用的節點不存在：' + String(id).slice(0, 30), at);
          else if (ps.has(id)) push('FLOW_DUP_NODE', '路線內節點重複：' + id, at);
          ps.add(id);
        });
        const last = nodeById[r.path[r.path.length - 1]];
        if (last && last.type !== 'approval' && last.type !== 'proxy') push('FLOW_BAD_TERMINAL', '路線最後一關必須是簽核或代核關', at);
        const cv = validateCondInternal(r.when, condCtxRoute);
        cv.forEach((e) => push(e.code, e.message, at + '.when'));
        if (i === def.routes.length - 1 && r.when !== true) push('FLOW_NO_DEFAULT', '最後一條路線必須是預設路線（when: true）', at);
      });
    }

    // 調整
    if (!Array.isArray(def.modifiers) || def.modifiers.length > LIM.modifiers) push('FLOW_BAD_MODIFIER', 'modifiers 無效', 'modifiers');
    else {
      const mids = new Set();
      def.modifiers.forEach((m, i) => {
        const at = 'modifiers[' + i + ']';
        if (!isPlain(m) || typeof m.id !== 'string' || !m.id || m.id.length > LIM.str) { push('FLOW_BAD_MODIFIER', '調整 id 無效', at); return; }
        if (reservedId(m.id)) { push('FLOW_BAD_MODIFIER', '調整 id 是保留字', at); return; }
        if (mids.has(m.id)) push('FLOW_BAD_MODIFIER', '調整 id 重複', at);
        if (!['insert_before', 'insert_after', 'remove', 'replace'].includes(m.op)) push('FLOW_BAD_MODIFIER', '調整 op 無效', at);
        if (typeof m.target !== 'string' || !nodeIds.has(m.target)) push('FLOW_NODE_MISSING', '調整的 target 不存在', at);
        if (m.op !== 'remove' && (typeof m.node !== 'string' || !nodeIds.has(m.node))) push('FLOW_NODE_MISSING', '調整的 node 不存在', at);
        const cv = validateCondInternal(m.when, { modifier: true, nodeIds, nodeTier, modIdsBefore: new Set(mids) });
        cv.forEach((e) => push(e.code, e.message, at + '.when'));
        mids.add(m.id);
      });
    }
  } catch (e) {
    push('FLOW_BAD_SCHEMA', '流程定義結構無法檢查', '');
  }
  return { ok: errs.length === 0, errors: errs };
}
function validateCondInternal(ast, ctx) {
  const errs = [];
  try { checkCond(ast, ctx, 1, { n: 0 }, errs); } catch (e) { if (e !== ABORT) errs.push(BAD('條件結構無法檢查')); }
  return errs;
}

// ───────────────────────── 預設流程（由現行核決表／門檻／路徑／BU 設定轉出）─────────────────────────
/** 把 getFlowCfg() 的輸出整理成預設流程需要的兩個清單（排序＝集合語意，雜湊穩定） */
function legacyInputs(flowCfg) {
  const fc = (flowCfg && typeof flowCfg === 'object') ? flowCfg : {};
  const bs = (fc.buSettings && typeof fc.buSettings === 'object') ? fc.buSettings : {};
  const mgr2Bus = Object.keys(bs).filter((bu) => bs[bu] && bs[bu].hasMgr2Step).sort();
  const ex = Array.isArray(fc.exceptions) ? fc.exceptions : [];
  const skipUsers = [...new Set(ex.filter((e) => e && e.action === 'skip_mgr1' && typeof e.username === 'string' && e.username).map((e) => e.username))].sort();
  return { mgr2Bus, skipUsers };
}

function buildLegacyDefault(flowCfg) {
  const inp = legacyInputs(flowCfg);
  const yuan = (n) => BigInt(n).toString();
  const ids = (arr) => arr.map((t) => 'n_' + t);
  const def = {
    schema: SCHEMA, id: 'legacy-v0', no: 0, name: '預設流程（現行核決表）', origin: 'legacy',
    limits: { maxSteps: MAX_STEPS_CAP },
    semantics: { sameApprover: 'forbid' },
    tables: {
      margin: QA.ROWS.map((r) => {
        const o = { cat: r.key, l1: r.l1, l2: r.l2 };
        if (r.fixedLevel) o.fixedLevel = r.fixedLevel;
        return o;
      }),
    },
    nodes: TIER_KEYS.map((t) => ({ id: 'n_' + t, tier: t, type: NODE_KIND[t].type, label: QA.TIERS[t], approver: JSON.parse(JSON.stringify(NODE_KIND[t].approver)), reassignable: NODE_KIND[t].reassignable, locked: NODE_KIND[t].locked })),
    routes: [
      { id: 'r_board', name: '金額超過董事會門檻', when: { cmp: ['amount', '>', yuan(QA.AMOUNT.board)] }, path: ids(QA.PATHS.board) },
      { id: 'r_l3', name: '董事長層級', when: { any: [{ cmp: ['amount', '>', yuan(QA.AMOUNT.chairman)] }, { cmp: ['marginLevel', '==', 3] }] }, path: ids(QA.PATHS[3]) },
      { id: 'r_l2', name: '總經理層級', when: { cmp: ['marginLevel', '==', 2] }, path: ids(QA.PATHS[2]) },
      { id: 'r_default', name: '一級主管層級（預設）', when: true, path: ids(QA.PATHS[1]) },
    ],
    modifiers: [
      { id: 'm_mgr2', name: 'BU 啟用二級主管關', op: 'insert_before', target: 'n_mgr1', node: 'n_mgr2', when: { all: [{ cmp: ['ownerBuPrimary', 'in', inp.mgr2Bus] }, { resolvable: 'n_mgr2' }] } },
      { id: 'm_skip', name: '二級主管例外：跳過一級主管', op: 'remove', target: 'n_mgr1', when: { all: [{ applied: 'm_mgr2' }, { resolvedIn: ['n_mgr2', inp.skipUsers] }] } },
    ],
  };
  return deepFreeze(def);
}

/** 流程定義的內容雜湊（key 排序的 canonical JSON → sha256） */
function defHash(def) { return crypto.createHash('sha256').update(canon(def), 'utf8').digest('hex'); }

// 預設流程（沒有任何 BU 啟用二級主管、沒有例外時）的釘選雜湊：核決表／門檻／路徑表任何一處變動、或轉換邏輯被改，這個值就會變（scripts/check-quote-rules-pin.js 會失敗提醒）。
const DEFAULT_FLOW_HASH = '520fc5a5d553240ea068eb23bba3ec4712b163d7db76301b342166fb99284194';

const legacyCache = new Map();
/** 取得（快取的）預設流程；同樣的 BU／例外輸入重用同一個已驗證、已凍結的物件 */
function legacyDefault(flowCfg) {
  const inp = legacyInputs(flowCfg);
  const key = canon(inp);
  let d = legacyCache.get(key);
  if (!d) {
    d = buildLegacyDefault(flowCfg);
    const v = validateFlowDef(d);
    if (!v.ok) throw new FlowError('FLOW_INVALID_DEF', '預設流程驗證失敗：' + v.errors.map((e) => e.code).join(','));
    VALIDATED.add(d);
    if (legacyCache.size >= 32) legacyCache.clear();
    legacyCache.set(key, d);
  }
  return d;
}

const VALIDATED = new WeakSet();
function assertValid(def) {
  if (VALIDATED.has(def)) return;
  const v = validateFlowDef(def);
  if (!v.ok) throw new FlowError('FLOW_INVALID_DEF', '流程定義無效：' + v.errors.slice(0, 3).map((e) => e.code + (e.at ? '@' + e.at : '')).join(','));
  if (isDeepFrozen(def)) VALIDATED.add(def);   // 只有「整棵樹都已凍結」（真正不可變）的定義才快取驗證結果；淺層凍結或未凍結的每次重新驗證
}

// ───────────────────────── 評估 ─────────────────────────
/**
 * 事實（facts）：由呼叫端從 derived／使用者資料組成。revenueCents／gpCents 必須是安全整數（分），revenue > 0。
 * { owner, ownerBu（已套用「normalizeBu(...)[0] || 'ERP'」的結果）, category（核決表列 key）, classKeys[], revenueCents, gpCents }
 */
function makeFacts(i) {
  if (!i || typeof i !== 'object') throw new FlowError('FLOW_BAD_FACTS', 'facts 無效');
  if (!Number.isSafeInteger(i.revenueCents) || !Number.isSafeInteger(i.gpCents)) throw new FlowError('FLOW_BAD_FACTS', '金額必須是安全整數（分）');
  if (i.revenueCents <= 0) throw new FlowError('FLOW_BAD_FACTS', '折後未稅金額必須大於 0');
  return {
    owner: String(i.owner === undefined || i.owner === null ? '' : i.owner),
    ownerBuPrimary: String(i.ownerBu === undefined || i.ownerBu === null ? 'ERP' : i.ownerBu),   // 屬性鍵語意：舊程式以 buSettings[ownerBu] 查，鍵會被轉成字串
    category: String(i.category),
    classKeys: Array.isArray(i.classKeys) ? i.classKeys.filter((x) => typeof x === 'string') : [],
    amountCents: BigInt(i.revenueCents),
    gpCents: BigInt(i.gpCents),
  };
}

function applyModifier(seq, m) {
  const ti = seq.indexOf(m.target);
  if (ti < 0) return false;   // 目標不在序列：沒有作用（舊程式 indexOf<0 時同樣不動）
  switch (m.op) {
    case 'insert_before':
      if (seq.includes(m.node)) throw new FlowError('FLOW_INVALID_RESULT', '調整會使序列出現重複節點');
      seq.splice(ti, 0, m.node);
      return true;
    case 'insert_after':
      if (seq.includes(m.node)) throw new FlowError('FLOW_INVALID_RESULT', '調整會使序列出現重複節點');
      seq.splice(ti + 1, 0, m.node);
      return true;
    case 'remove':
      seq.splice(ti, 1);
      return true;
    case 'replace':
      if (m.node !== m.target && seq.includes(m.node)) throw new FlowError('FLOW_INVALID_RESULT', '調整會使序列出現重複節點');
      seq[ti] = m.node;
      return true;
    default: throw new FlowError('FLOW_BAD_MODIFIER', '未知 op');
  }
}

/**
 * 評估：選第一條符合的路線 → 依序套用所有符合的調整 → 驗證結果。純函式，不修改 def／facts。
 * io（解析簽核人的介面，由呼叫端以現有函式提供）：{ mgr2():string|null, mgr1():string|null, roster:{gm[],chairman[]}, boardProxy():string[] }
 */
function evaluate(def, factsInput, io) {
  assertValid(def);
  const facts = makeFacts(factsInput);
  const nodeById = Object.create(null);
  def.nodes.forEach((n) => { nodeById[n.id] = n; });
  let lv = null;
  const resolved = Object.create(null);
  const env = {
    facts, modifier: false, applied: new Set(),
    level: () => (lv === null ? (lv = marginLevelOf(def.tables.margin, facts.category, facts.gpCents, facts.amountCents)) : lv),
    resolveNode: (id) => {
      if (resolved[id]) return resolved[id];
      const n = nodeById[id];
      let r;
      if (!n) r = { person: null, persons: [] };
      else if (n.tier === 'mgr2') r = { person: io.mgr2() || null, persons: [] };
      else if (n.tier === 'mgr1') r = { person: io.mgr1() || null, persons: [] };
      else if (n.tier === 'gm') r = { person: null, persons: (io.roster && io.roster.gm) || [] };
      else if (n.tier === 'chairman') r = { person: null, persons: (io.roster && io.roster.chairman) || [] };
      else r = { person: null, persons: io.boardProxy() || [] };
      return (resolved[id] = r);
    },
  };
  const trace = [];
  let route = null;
  for (const r of def.routes) {
    const m = evalCond(r.when, env);
    trace.push({ rule: r.id, matched: m });
    if (m) { route = r; break; }
  }
  if (!route) throw new FlowError('FLOW_NO_ROUTE', '沒有符合的路線');
  const seq = route.path.slice();
  const applied = [];
  env.modifier = true;
  for (const m of def.modifiers) {
    const matched = evalCond(m.when, env);
    let did = false;
    if (matched) { did = applyModifier(seq, m); if (did) { applied.push(m.id); env.applied.add(m.id); } }
    trace.push({ rule: m.id, matched, applied: did });
  }
  if (seq.length < 1 || seq.length > def.limits.maxSteps) throw new FlowError('FLOW_INVALID_RESULT', '關卡數超出範圍');
  if (new Set(seq).size !== seq.length) throw new FlowError('FLOW_INVALID_RESULT', '關卡重複');
  const tierOf = (id) => nodeById[id].tier;
  return {
    routeId: route.id, applied, trace,
    financialTiers: route.path.map(tierOf),
    actualTiers: seq.map(tierOf),
    marginLevel: lv,   // 只有條件用到才會算；僅供測試／除錯，不寫進快照
  };
}

// ───────────────────────── 簽核人可行性檢查（複製 quoteRoutes.chainBlockers 的後半，錯誤碼與訊息逐字相同）─────────────────────────
/** 每一關各挑一位、且不重複的人是否存在（與 quoteRoutes.distinctAssignment 相同的回溯） */
function distinctAssignment(cands, i, used) {
  i = i || 0;
  used = used || new Set();
  if (i === cands.length) return true;
  for (const p of cands[i]) {
    if (used.has(p)) continue;
    used.add(p);
    if (distinctAssignment(cands, i + 1, used)) return true;
    used.delete(p);
  }
  return false;
}

/**
 * io：{ owner, costBy, mgr2, mgr1（已解析的人；沒有＝null）, canApprove(un):boolean, roster:{gm[],chairman[]}, boardProxy:string[] }
 * 回傳 blockers 陣列（空＝可送簽）。順序、code、message 與舊 chainBlockers 一致。
 */
function chainBlockersOf(actualTiers, io) {
  const out = [];
  const { mgr2, mgr1 } = io;
  if (actualTiers.includes('mgr2') && (!mgr2 || !io.canApprove(mgr2))) {
    out.push({ code: 'NO_MGR2', message: '找不到你的直屬二級主管，請管理員先設定 supervisor 欄位' });
  }
  if (actualTiers.includes('mgr1') && !mgr1) {
    out.push({ code: 'NO_MANAGER', message: '找不到你的一級主管（直屬主管鏈上沒有在職的一級主管），請管理員先設定直屬主管' });
  }
  const activeOf = (list) => list.filter((un) => io.canApprove(un));
  if (actualTiers.includes('gm') && !activeOf(io.roster.gm).length) out.push({ code: 'NO_GM', message: '尚未設定總經理名單，請管理員至「簽核設定」維護' });
  if (actualTiers.includes('chairman') && !activeOf(io.roster.chairman).length) out.push({ code: 'NO_CHAIRMAN', message: '尚未設定董事長名單，請管理員至「簽核設定」維護' });
  if (actualTiers.includes('board') && !io.boardProxy.length) out.push({ code: 'NO_BOARD_PROXY', message: '找不到董事會代核人（管理部秘書），請管理員確認' });
  if (out.length) return out;
  const excluded = new Set([io.owner]);
  if (io.costBy) excluded.add(io.costBy);
  const cands = actualTiers.map((t) => {
    let list = [];
    if (t === 'mgr2') list = mgr2 ? [mgr2] : [];
    else if (t === 'mgr1') list = mgr1 ? [mgr1] : [];
    else if (t === 'gm') list = activeOf(io.roster.gm);
    else if (t === 'chairman') list = activeOf(io.roster.chairman);
    else if (t === 'board') list = io.boardProxy.slice();
    return list.filter((un) => !excluded.has(un));
  });
  const emptyIdx = cands.findIndex((c) => !c.length);
  if (emptyIdx >= 0) {
    out.push({ code: 'NO_ELIGIBLE_SIGNER', message: `「${QA.TIERS[actualTiers[emptyIdx]]}」這一關沒有可以簽核的人（擁有者與成本填寫人本人不能簽自己的單），請管理員調整簽核名單` });
  } else if (!distinctAssignment(cands)) {
    out.push({ code: 'SIGNER_OVERLAP', message: '簽核名單有重疊：同一個人不能簽同一張單的兩個關卡，目前無法讓每一關都有不同的人簽核，請管理員調整簽核名單' });
  }
  return out;
}

/**
 * 一次做完：評估 ＋ 簽核人指派 ＋ 可行性檢查。
 * io 同 evaluate，另需 owner／costBy／canApprove／roster／boardProxy()。mgr2／mgr1 只在該關真的在序列裡才取人
 * （舊程式：mgr2＝BU 啟用且找得到才非 null；mgr1＝tiers 含 mgr1 才查）。
 */
function plan(def, factsInput, io) {
  const ev = evaluate(def, factsInput, io);
  const mgr2 = ev.actualTiers.includes('mgr2') ? (io.mgr2() || null) : null;
  const mgr1 = ev.actualTiers.includes('mgr1') ? (io.mgr1() || null) : null;
  const blockers = chainBlockersOf(ev.actualTiers, {
    owner: io.owner, costBy: io.costBy, mgr2, mgr1, canApprove: io.canApprove,
    roster: io.roster, boardProxy: io.boardProxy(),
  });
  return { routeId: ev.routeId, applied: ev.applied, trace: ev.trace, financialTiers: ev.financialTiers, actualTiers: ev.actualTiers, mgr2, mgr1, blockers, defNo: def.no, defHash: defHash(def), origin: def.origin };
}

/** 送簽時凍結在 q.approval.flow 的快照（純 JSON；不含毛利／成本的數值，不含 marginLevel） */
function buildSnapshot(def, p, facts, flowCfg) {
  const inp = legacyInputs(flowCfg);
  return {
    v: 1, engine: ENGINE_ID, defNo: p.defNo, defHash: p.defHash, origin: p.origin,
    routeId: p.routeId, applied: p.applied.slice(), path: p.financialTiers.slice(), tiers: p.actualTiers.slice(),
    trace: p.trace.map((t) => (t.applied === undefined ? { rule: t.rule, matched: t.matched } : { rule: t.rule, matched: t.matched, applied: t.applied })),
    factsPublic: { category: String(facts.category), ownerBuPrimary: String(facts.ownerBu === undefined || facts.ownerBu === null ? 'ERP' : facts.ownerBu), amountCents: facts.revenueCents },
    legacyCfg: p.origin === 'legacy' ? { mgr2Bus: inp.mgr2Bus, skipUsers: inp.skipUsers } : null,
  };
}

// ───────────────────────── 比對（影子模式）─────────────────────────
/** 比對舊路徑與引擎的結果；回傳不一致的欄位名稱陣列（空＝一致）。old/new 形狀：{ financialTiers, actualTiers, mgr2, mgr1, blockers } */
function diffPlans(a, b) {
  const diffs = [];
  const eq = (x, y) => canon(x) === canon(y);
  if (!eq(a.financialTiers, b.financialTiers)) diffs.push('financialTiers');
  if (!eq(a.actualTiers, b.actualTiers)) diffs.push('actualTiers');
  if ((a.mgr2 || null) !== (b.mgr2 || null)) diffs.push('mgr2');
  if ((a.mgr1 || null) !== (b.mgr1 || null)) diffs.push('mgr1');
  if (!eq(a.blockers, b.blockers)) diffs.push('blockers');
  return diffs;
}

// ───────────────────────── 稽核回報（節流；不洩漏金額／成本）─────────────────────────
/**
 * createReporter({ writeLog, now })：影子不一致／引擎例外／熔斷退回 的稽核寫入器。
 * 同一個 (kind, key) 在 windowMs 內只寫第一筆（key 含階段與單號）；全域預算「依階段分桶」：每個階段（preview／submit…）每 globalWindowMs 各最多 globalMax 筆（超過只計數）——
 * 清單頁的 preview 洪水只會用光 preview 桶，不會擠掉最需要留痕的 submit 階段事件。writeLog 例外一律吞掉；寫入成功（writeLog 沒丟例外）後才把 key 記入 seen（寫入失敗下次還有機會補寫，預算仍扣除避免狂打）。
 * 計數器（stats）是行程內的，Vercel 多實例各自計；權威紀錄是稽核事件。
 */
function createReporter(opts) {
  const o = Object.assign({ writeLog: null, now: Date.now, windowMs: 10 * 60 * 1000, globalMax: 60, globalWindowMs: 60 * 60 * 1000 }, opts || {});
  const seen = new Map();
  const buckets = new Map();   // phase -> { start, count }
  const stats = { since: new Date(o.now()).toISOString(), runs: 0, mismatches: 0, errors: 0, fallbacks: 0, suppressed: 0, written: 0 };
  const clip = (s, n) => String(s).replace(/[\u0000-\u001f\u007f]/g, ' ').slice(0, n);
  function emit(kind, action, phase, quoteId, operator, target, detail, req) {
    const t = o.now();
    const k = kind + '|' + phase + '|' + quoteId;
    const last = seen.get(k);
    if (last !== undefined && t - last < o.windowMs) { stats.suppressed++; return false; }
    const bk = String(phase);
    let g = buckets.get(bk);
    if (!g) { g = { start: t, count: 0 }; if (buckets.size < 16) buckets.set(bk, g); }   // 階段只有少數幾種；超過 16 種（不應發生）的不進表，等於每次都是新桶
    if (t - g.start >= o.globalWindowMs) { g.start = t; g.count = 0; }
    if (g.count >= o.globalMax) { stats.suppressed++; return false; }
    g.count++;
    let wrote = false;
    try { if (typeof o.writeLog === 'function') { o.writeLog(action, clip(operator, 80), clip(target, 80), clip(detail, 480), req); stats.written++; wrote = true; } } catch (_) { /* 稽核寫入失敗不影響簽核 */ }
    if (wrote) {
      seen.set(k, t);
      if (seen.size > 500) for (const [kk, v] of seen) if (t - v >= o.windowMs) seen.delete(kk);
    }
    return wrote;
  }
  const tiersText = (a) => (Array.isArray(a) && a.length ? a.join('→') : '∅');
  const codes = (b) => (Array.isArray(b) && b.length ? b.map((x) => x.code).join(',') : '-');
  return {
    stats,
    run() { stats.runs++; },
    mismatch(ctx, oldP, newP, diffs, req) {
      stats.mismatches++;
      const detail = `階段=${ctx.phase}｜模式=${ctx.mode}｜歸類=${ctx.category}｜BU=${ctx.ownerBu}｜差異欄位=${diffs.join(',')}｜舊：關卡 ${tiersText(oldP.actualTiers)}（財務 ${tiersText(oldP.financialTiers)}）阻擋碼 ${codes(oldP.blockers)}｜新：關卡 ${tiersText(newP.actualTiers)}（財務 ${tiersText(newP.financialTiers)}）阻擋碼 ${codes(newP.blockers)}`;
      return emit('mm', 'QUOTE_FLOW_SHADOW_MISMATCH', ctx.phase, ctx.quoteId, ctx.operator, ctx.quoteNo, detail, req);
    },
    error(ctx, err, req) {
      stats.errors++;
      const detail = `階段=${ctx.phase}｜模式=${ctx.mode}｜引擎例外=${(err && (err.code || err.name)) || 'Error'}：${clip((err && err.message) || '', 160)}｜已忽略，仍以舊路徑為準`;
      return emit('er', 'QUOTE_FLOW_SHADOW_ERROR', ctx.phase, ctx.quoteId, ctx.operator, ctx.quoteNo, detail, req);
    },
    fallback(ctx, err, req) {
      stats.fallbacks++;
      const detail = `階段=${ctx.phase}｜熔斷：引擎${err ? '例外 ' + ((err.code || err.name) + '：' + clip(err.message || '', 160)) : '結果無效'}，已自動退回舊路徑`;
      return emit('fb', 'QUOTE_FLOW_ENGINE_FALLBACK', ctx.phase, ctx.quoteId, ctx.operator, ctx.quoteNo, detail, req);
    },
  };
}

// ───────────────────────── 白話描述（後台唯讀顯示）─────────────────────────
const OP_ZH = { '==': '等於', '!=': '不等於', '>': '大於', '>=': '大於等於', '<': '小於', '<=': '小於等於', in: '屬於', not_in: '不屬於', contains: '包含', not_contains: '不包含' };
function fmtYuanStr(s) {
  const n = BigInt(s.split('.')[0]);
  if (!s.includes('.') && n % 10000n === 0n && n >= 10000n) return (n / 10000n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',') + ' 萬';
  return s.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + ' 元';
}
function describeCond(x, depth, names) {
  depth = depth || 0;
  names = names || {};
  const nm = (id) => (has(names, id) ? names[id] : String(id));
  if (depth > LIM.depth) return '…';
  if (x === true) return '一律';
  if (x === false) return '永不';
  if (!isPlain(x)) return '（無效條件）';
  const k = Object.keys(x)[0];
  const v = x[k];
  switch (k) {
    case 'all': return v.map((c) => describeCond(c, depth + 1, names)).join('，且');
    case 'any': return '（' + v.map((c) => describeCond(c, depth + 1, names)).join('，或') + '）';
    case 'not': return '並非（' + describeCond(v, depth + 1, names) + '）';
    case 'cmp': {
      const [n, op, r] = v;
      const vd = VARS[n];
      const rhs = vd.type === 'money' ? fmtYuanStr(r) : vd.type === 'pct' ? r + '%' : Array.isArray(r) ? (r.length ? r.join('、') : '（無）') : String(r);
      return vd.label + ' ' + OP_ZH[op] + ' ' + rhs;
    }
    case 'applied': return '已套用「' + nm(v) + '」';
    case 'resolvable': return nm(v) + '找得到人';
    case 'resolvedIn': return nm(v[0]) + '的簽核人屬於例外名單（' + v[1].length + ' 人）';
    default: return '（未知條件）';
  }
}
/** 給後台顯示的摘要（唯讀）。不含帳號名單內容，只有人數。 */
function describeDef(def) {
  const byId = Object.create(null);
  def.nodes.forEach((n) => { byId[n.id] = n; });
  const names = Object.create(null);
  def.nodes.forEach((n) => { names[n.id] = n.label; });
  def.modifiers.forEach((m) => { names[m.id] = m.name || m.id; });
  const pathNames = (path) => path.map((id) => byId[id].label).join(' → ');
  return {
    name: def.name, origin: def.origin, no: def.no, hash: defHash(def),
    nodes: def.nodes.map((n) => ({ tier: n.tier, label: n.label, type: n.type, locked: !!n.locked })),
    routes: def.routes.map((r, i) => ({ order: i + 1, name: r.name || r.id, when: describeCond(r.when, 0, names), path: pathNames(r.path) })),
    modifiers: def.modifiers.map((m) => ({ name: m.name || m.id, op: m.op, when: describeCond(m.when, 0, names) })),
    table: def.tables.margin.map((r) => ({ cat: r.cat, l1: r.l1, l2: r.l2, fixedLevel: r.fixedLevel || null })),
  };
}

module.exports = {
  SCHEMA, ENGINE_ID, TIER_KEYS, LIM, VARS, NODE_KIND, DEFAULT_FLOW_HASH, FlowError,
  engineMode, engineModeInfo,
  parseMoneyCents, parsePct, validateCond, validateFlowDef,
  evalCond, marginLevelOf, evaluate, plan, chainBlockersOf, distinctAssignment, makeFacts,
  buildLegacyDefault, legacyDefault, legacyInputs, defHash, buildSnapshot, diffPlans, createReporter, describeDef, describeCond,
  canon,
};
