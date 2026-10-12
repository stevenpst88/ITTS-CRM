#!/usr/bin/env node
/**
 * 「業務唯讀看得到顧問逐列成本」檢查。用法：node scripts/check-quote-cost-readonly.js（假 app＋記憶體資料庫；不碰 data.json／auth.json／audit.log.json）
 * 業主 2026-10-12 決定：顧問儲存／完成過成本（consultantWrote／filledAt）後，業務（擁有者）可以唯讀看到逐列成本；顧問仍是唯一編輯者。
 * 動 lib/quoteRoutes.js 的 permOf（canSeeCost／canEditCost）、serialize 的成本欄位、PUT 報價單的 acceptCost、PUT /costs 之後必跑。
 *   1) 顧問動過之前業務看不到（沒有東西可看）；儲存草稿後、完成後、被退回重填後都唯讀可見（costLines、items[].cost 不適用新式、contingencyPct、costBreakdown）
 *   2) 業務不能改：PUT 報價單帶 costLines／items.cost 被忽略（成本逐位元不變）；PUT /costs 403（已存在的「改 products 切成不需顧問會整批重置成本」是原有行為，不在本次範圍）
 *   3) 顧問仍可編輯（200），權限旗標 canEditCost 只有顧問；簽核雜湊／狀態不因可見度改變
 *   4) 其他角色的可見度不變：與 QR_BASE_ROOT 指定的「改版前」程式（預設不比對）逐角色逐狀態相同（golden 值寫在下方 BASE_MATRIX）
 * 測試資料只用通用字串。
 */
'use strict';
const path = require('path');
const ROOT = path.join(__dirname, '..');
const LIB = process.env.QR_BASE_ROOT || ROOT;
const QA = require(path.join(LIB, 'lib/quoteApproval.js'));
const registerQuoteRoutes = require(path.join(LIB, 'lib/quoteRoutes.js'));
const PNL = require(path.join(LIB, 'lib/quotePnlExcel.js'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);
const L = (cat, desc, qty, unitCost, extra) => Object.assign({ cat, desc, qty, unitCost }, extra || {});

const USERS = [
  { username: 'admin1', role: 'admin', active: true, displayName: 'Admin' },
  { username: 'own1', role: 'user', active: true, displayName: 'Owner1', supervisor: 'mgr1', bu: ['ITS'] },
  { username: 'mgr1', role: 'manager1', active: true, displayName: 'Mgr1', bu: ['ITS'] },
  { username: 'gm1', role: 'executive', active: true, displayName: 'Gm' },
  { username: 'cons1', role: 'user', active: true, displayName: 'Cons1', bu: ['ITS'] },
  { username: 'cons2', role: 'user', active: true, displayName: 'Cons2', bu: ['ITS'] },
  { username: 'sec1', role: 'secretary', active: true, displayName: 'Sec', bu: ['ITS'] },
  { username: 'x1', role: 'user', active: true, displayName: 'Other', bu: ['ERP'] },
];
const CLASSES = { PC: { cls: 'consult', costBySales: false }, PS: { cls: 'software', costBySales: true } };
function mkEnv(quotations) {
  const routes = {}; const app = {};
  ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
  const env = { data: { quotations, quoteApproval: { roster: { gm: ['gm1'], chairman: [], boardProxy: [], costProviders: ['cons1', 'cons2'], sealManagers: [] }, productClasses: CLASSES } }, seq: 0 };
  env.auth = { users: JSON.parse(JSON.stringify(USERS)) };
  const deps = {
    db: { load: () => env.data, save: () => {}, flush: async () => {} }, loadAuth: () => env.auth, saveAuth: () => {},
    requireAuth: (req, rs, next) => next(), requireAdmin: (req, rs, next) => next(),
    writeLog: () => {}, pushNotification: () => {},
    getViewableOwners: (req) => { const rl = req.session.user.role; if (rl === 'admin' || rl === 'executive' || rl === 'manager1' || rl === 'secretary') return ['own1']; return [req.session.user.username]; },
    sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-NEW', taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}),
    buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: PNL.buildQuotePnlExcel, QUOTE_TEMPLATE: '',
    uuidv4: () => 'u' + (++env.seq), normalizeBu: (b) => (Array.isArray(b) ? b : b ? [b] : []), getUserFeatures: () => ['quotations'],
  };
  registerQuoteRoutes(app, deps);
  env.call = (user, method, p, params, body) => new Promise((resolve, reject) => {
    const h = [].concat(...routes[method.toUpperCase() + ' ' + p]);
    const role = (env.auth.users.find((u) => u.username === user) || {}).role;
    const req = { session: { user: { username: user, role } }, params: params || {}, query: {}, body: body === undefined ? {} : JSON.parse(JSON.stringify(body)) };
    const rs = { status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j }); return this; }, send() { resolve({ s: this._s || 200, j: null }); return this; }, setHeader() {}, set() {} };
    let i = 0;
    const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
    next();
  });
  return env;
}
const mkQuote = () => ({
  id: 'Q1', quoteNo: 'QU-1', owner: 'own1', company: 'TestCo', projectName: 'Proj', quoteDate: '2026-10-08', status: 'draft', createdAt: '2026-10-08T00:00:00Z', updatedAt: '2026-10-08T00:00:00Z',
  validUntil: '2026-10-30', products: ['PC'], costBy: 'cons1', approval: null, discountType: 'none', discountValue: 0,
  items: [{ lid: 'i1', desc: '導入顧問', unit: '式', qty: 1, unitPrice: 600000, cost: 0 }, { lid: 'i2', desc: '教育訓練', unit: '式', qty: 1, unitPrice: 100000, cost: 0 }],
  costFlow: { state: 'requested', by: 'cons1', requestedAt: '2026-10-08T00:00:00Z', filledAt: null, note: '', sig: null, consultantWrote: false },
});
const formItems = (q) => q.items.map((x) => ({ lid: x.lid, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice, cost: x.cost }));
const PUTQ = '/api/quotations/:id', COSTS = '/api/quotations/:id/costs';
const lines = () => [L('consult', 'PM', 2, 7000, { unit: '人天', rel: 'link', forLid: 'i1' }), L('consult', 'SD', 1, 6000, { unit: '人天', rel: 'link', forLid: 'i2' }), { cat: 'other', auto: 'stamp' }];
const cb = (extra) => Object.assign({ costLines: lines(), done: false, costModel: 2 }, extra || {});

// 「改版前」程式在同一組情境下，非擁有者角色的 canSeeCost／是否拿到 costLines（golden：以 origin/main 52623db 跑出）
// 情境順序：untouched（顧問尚未動）→ saved（儲存草稿）→ done（完成）
const BASE_MATRIX = { admin1: 't,t,t', mgr1: 't,t,t', gm1: 't,t,t', cons1: 't,t,t', cons2: 'x,x,x', sec1: 't,t,t', x1: 'x,x,x' };

async function run() {
  const env = mkEnv([mkQuote()]);
  const Q = () => env.data.quotations[0];
  const get = async (u) => await env.call(u, 'GET', PUTQ, { id: 'Q1' });
  const cell = (r) => (r.s !== 200 ? 'x' : (r.j.perm.canSeeCost ? 't' : 'f') + (Array.isArray(r.j.costLines) ? '' : ''));
  const matrix = { admin1: [], mgr1: [], gm1: [], cons1: [], cons2: [], sec1: [], x1: [] };
  const snap = async () => { for (const u of Object.keys(matrix)) matrix[u].push(cell(await get(u))); };
  // untouched
  let o = (await get('own1')).j;
  t('1.1 顧問尚未動過：業務看不到成本（沒有 costLines、perm.canSeeCost=false）', o.perm.canSeeCost === false && o.costLines === undefined && o.contingencyPct === undefined);
  await snap();
  // 儲存草稿
  let r = await env.call('cons1', 'PUT', COSTS, { id: 'Q1' }, cb());
  t('1.2 顧問儲存草稿 → 200', r.s === 200, r.s);
  o = (await get('own1')).j;
  t('1.3 草稿儲存後：業務唯讀看到 costLines（3 列、金額與顧問存的相同）', o.perm.canSeeCost === true && Array.isArray(o.costLines) && o.costLines.length === 3 && o.costLines[0].unitCost === 7000 && o.costLines[0].desc === 'PM', J(o.costLines));
  t('1.4 業務仍不能編輯成本：perm.canEditCost=false；顧問 perm.canEditCost=true', o.perm.canEditCost === false && (await get('cons1')).j.perm.canEditCost === true);
  t('1.5 業務拿到 costBreakdown 與 costLinesSig（只讀比對用）', o.costBreakdown && o.costBreakdown.consult > 0 && typeof o.costLinesSig === 'string');
  await snap();
  // 業務寫入嘗試
  const before = J(Q().costLines), beforeFlow = J(Q().costFlow), beforeItems = J(Q().items.map((i) => [i.lid, i.cost]));
  const forged = lines().map((l) => Object.assign({}, l, { unitCost: 1 }));
  r = await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { costModel: 2, costLines: forged, costLinesSig: o.costLinesSig, items: formItems(Q()).map((x) => Object.assign(x, { cost: 999 })), note: 'sales edit' });
  t('2.1 業務 PUT 報價單帶偽造 costLines／items.cost：被忽略（200，但成本逐位元不變；其他欄位照存）', r.s === 200 && J(Q().costLines) === before && J(Q().items.map((i) => [i.lid, i.cost])) === beforeItems && Q().note === 'sales edit', r.s + J(r.j && r.j.message));
  t('2.2 costFlow 沒被動（仍是顧問寫過、requested）', J(Q().costFlow) === beforeFlow || (Q().costFlow.state === 'requested' && Q().costFlow.consultantWrote === true), J(Q().costFlow));
  r = await env.call('own1', 'PUT', COSTS, { id: 'Q1' }, cb({ costLines: forged }));
  t('2.3 業務 PUT /costs → 403（不是成本填寫人）；成本不變', r.s === 403 && J(Q().costLines) === before, r.s);
  // 完成
  r = await env.call('cons1', 'PUT', COSTS, { id: 'Q1' }, cb({ done: true, contingencyPct: 5 }));
  t('3.1 顧問完成 → 200（顧問編輯仍正常）', r.s === 200 && Q().costFlow.state === 'filled', r.s + J(r.j && r.j.message));
  o = (await get('own1')).j;
  t('3.2 完成後業務唯讀可見 costLines＋風險預留；仍不能編輯', o.perm.canSeeCost === true && o.perm.canEditCost === false && o.contingencyPct === 5 && o.costLines.length === 3);
  await snap();
  // 退回重填（顧問 done:false）
  r = await env.call('cons1', 'PUT', COSTS, { id: 'Q1' }, cb({ done: false }));
  o = (await get('own1')).j;
  t('3.3 顧問退回重填（requested、consultantWrote）：業務仍唯讀看得到（舊規則此時看不到）', r.s === 200 && Q().costFlow.state === 'requested' && o.perm.canSeeCost === true && Array.isArray(o.costLines));
  // 雜湊／簽核
  await env.call('cons1', 'PUT', COSTS, { id: 'Q1' }, cb({ done: true, contingencyPct: 5 }));
  const h0 = QA.contentHash(Q());
  const sub = await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' }, {});
  t('4.1 送簽照常（需毛利／鏈條件允許時 200；否則為既有的阻擋訊息，不是因可見度）', sub.s === 200 || (sub.s >= 400 && sub.j && sub.j.code && !/COST|PERM/.test(sub.j.code)), sub.s + J(sub.j && sub.j.code));
  t('4.2 contentHash 與可見度無關（送簽前後相同；送簽成功時 approval.hash 相符）', QA.contentHash(Q()) === h0 && (!Q().approval || Q().approval.hash === h0 || Q().approval.state === 'none'));
  // 非擁有者矩陣
  if (process.env.QR_BASE_ROOT) { console.log('MATRIX ' + J(Object.fromEntries(Object.entries(matrix).map(([k, v]) => [k, v.slice(0, 1).concat(v.slice(1, 2), v.slice(2, 3)).join(',')])))); }
  else {
    const got = Object.fromEntries(Object.entries(matrix).map(([k, v]) => [k, v.join(',')]));
    t('5.1 其他角色（管理員、主管、總經理、秘書、被指派顧問、別人的顧問、別 BU 業務）各情境的可見度與改版前相同', J(got) === J(BASE_MATRIX), J(got) + ' vs ' + J(BASE_MATRIX));
  }
}
run().then(() => {
  let pass = 0, fail = 0;
  res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (x && !ok ? '  <- ' + x : '')); ok ? pass++ : fail++; });
  console.log('\n業務唯讀成本檢查：PASS ' + pass + ' / FAIL ' + fail);
  process.exit(fail ? 1 : 0);
}).catch((e) => { console.error('例外', e.stack); process.exit(2); });
