'use strict';
// check-audit-wave2.js 的各項目區段（B 單號流水帳、C 規則快照、D 編輯事件、E 董事會證據）。由 check-audit-wave2.js 載入；不單獨執行。
// 這些項目不開伺服器：假 app＋記憶體資料庫直接呼叫路由 handler（與 check-quote-timeline.js 同一套做法）。
const path = require('path');
const fs = require('fs');

const ROOT = path.join(__dirname, '..');
const J = (x) => JSON.stringify(x);

const USERS = [
  { username: 'admin1', role: 'admin', active: true, displayName: 'Admin' },
  { username: 'own1', role: 'user', active: true, displayName: 'Owner1', supervisor: 'mgr1', bu: ['ITS'] },
  { username: 'own2', role: 'user', active: true, displayName: 'Owner2', supervisor: 'mgr1', bu: ['ITS'] },
  { username: 'mgr1', role: 'manager1', active: true, displayName: 'Mgr1', bu: ['ITS'] },
  { username: 'gm1', role: 'executive', active: true, displayName: 'Gm' },
  { username: 'gm2', role: 'executive', active: true, displayName: 'Gm2' },
  { username: 'ch1', role: 'executive', active: true, displayName: 'Chair' },
  { username: 'cons1', role: 'user', active: true, displayName: 'Cons1', bu: ['ITS'] },
  { username: 'sec1', role: 'secretary', active: true, displayName: 'Sec', bu: ['ITS'] },
];
const CLASSES = { PC: { cls: 'consult', costBySales: false }, PS: { cls: 'software', costBySales: true } };

/** 假環境：routesPath 可換成變異版 quoteRoutes；events 收 db.insertAuditEvent 的事件 */
function mkEnv(routesPath, quotations, extraData, roster) {
  const registerQuoteRoutes = require(routesPath || path.join(ROOT, 'lib/quoteRoutes.js'));
  const PNL = require(path.join(ROOT, 'lib/quotePnlExcel.js'));
  const routes = {}; const app = {};
  ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
  const env = {
    data: Object.assign({ opportunities: [{ id: 'O1', owner: 'own1', stage: 'C', company: 'TestCo', product: 'Prod', amount: '100', bu: 'ITS' }, { id: 'O2', owner: 'own1', stage: 'C', company: 'TestCo', product: 'Prod2', amount: '100', bu: 'ITS' }], quotations: quotations || [],
      quoteApproval: { roster: Object.assign({ gm: ['gm1'], chairman: ['ch1'], boardProxy: ['sec1'], costProviders: ['cons1'], sealManagers: [] }, roster || {}), productClasses: CLASSES } }, extraData || {}),
    logs: [], notes: [], events: [], saves: 0, seq: 0,
  };
  env.auth = { users: JSON.parse(JSON.stringify(USERS)) };
  const deps = {
    db: { load: () => env.data, save: () => { env.saves++; }, flush: async () => {}, insertAuditEvent: (ev) => { env.events.push(ev); return Promise.resolve(true); } },
    loadAuth: () => env.auth, saveAuth: () => {},
    requireAuth: (req, rs, next) => next(), requireAdmin: (req, rs, next) => next(),
    writeLog: (...a) => env.logs.push(a), pushNotification: (...a) => env.notes.push(a),
    getViewableOwners: (req) => { const rl = req.session.user.role; if (rl === 'admin' || rl === 'executive' || rl === 'manager1' || rl === 'secretary') return ['own1', 'own2']; return [req.session.user.username]; },
    sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200),
    genQuoteNo: (data) => {   // 與 server.js genQuoteNo 同邏輯（永不倒退的計數器）
      const ym = '202610'; const prefix = `QU-${ym}-`; let mx = 0;
      (data.quotations || []).forEach((q) => { if (q.quoteNo && q.quoteNo.startsWith(prefix)) mx = Math.max(mx, parseInt(q.quoteNo.slice(prefix.length), 10)); });
      if (!data.quoteSeq) data.quoteSeq = {};
      const seq = Math.max(data.quoteSeq[ym] || 0, mx) + 1; data.quoteSeq[ym] = seq; return prefix + String(seq).padStart(3, '0');
    },
    taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}),
    buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: PNL.buildQuotePnlExcel, QUOTE_TEMPLATE: '',
    uuidv4: () => 'u' + (++env.seq), normalizeBu: (b) => (Array.isArray(b) ? b : b ? [b] : []), getUserFeatures: () => ['quotations'],
  };
  registerQuoteRoutes(app, deps);
  env.call = (user, method, p, params, body, query) => new Promise((resolve, reject) => {
    const h = [].concat(...routes[method.toUpperCase() + ' ' + p]);
    const role = (env.auth.users.find((u) => u.username === user) || {}).role;
    const req = { session: { user: { username: user, role } }, params: params || {}, query: query || {}, body: body === undefined ? {} : JSON.parse(JSON.stringify(body)), ip: '127.0.0.1', headers: {}, get: () => '' };
    const rs = { status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j }); return this; }, send() { resolve({ s: this._s || 200, j: null }); return this; }, setHeader() {}, set() {} };
    let i = 0;
    const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
    next();
  });
  return env;
}
const mkItems = () => [{ desc: '品項A', unit: '式', qty: 1, unitPrice: 1000, cost: 100 }];
const create = (env, opp, extra) => env.call('own1', 'POST', '/api/quotations', {}, Object.assign({ oppId: opp || 'O1', company: 'TestCo', products: ['PS'], items: mkItems() }, extra || {}));

// ═════════════ B：刪除快照＋單號流水帳 ═════════════
async function runB(h, routesPath) {
  const { ok, eq } = h;
  const QL = require(path.join(ROOT, 'lib/quoteLedger.js'));
  const OL = require(path.join(ROOT, 'lib/quoteOppLink.js'));
  const env = mkEnv(routesPath, []);
  // 建 5 張
  const ids = [];
  for (let i = 0; i < 5; i++) { const r = await create(env, 'O1'); ok(r.s === 201, `[B] 建立第 ${i + 1} 張 (${r.s} ${r.j && r.j.code})`); ids.push(r.j && r.j.id); }
  const nos = env.data.quotations.map((q) => q.quoteNo);
  eq(nos, ['QU-202610-001', 'QU-202610-002', 'QU-202610-003', 'QU-202610-004', 'QU-202610-005'], '[B] 單號依序');
  const led = env.data.quoteNoLedger || [];
  eq(led.map((e) => e.no), nos, '[B] 每張建立時就進流水帳');
  ok(led.every((e) => e.owner === 'own1' && e.status === 'active' && !!e.createdAt && !e.backfilled), '[B] 流水帳項目：owner／createdAt／active，非補登');

  // 刪除第 3 張（附原因）
  const q3 = JSON.parse(JSON.stringify(env.data.quotations[2]));
  q3.items[0].note = '稽核測試品項備註';
  env.data.quotations[2].items[0].note = '稽核測試品項備註';
  const before = JSON.parse(JSON.stringify(env.data.quotations[2]));
  let r = await env.call('own1', 'DELETE', '/api/quotations/:id', { id: ids[2] }, { reason: '客戶取消，重新報價' });
  eq(r.s, 200, '[B] 刪除草稿 200');
  ok(!env.data.quotations.some((q) => q.id === ids[2]), '[B] 已從 quotations 移除');
  const dq = env.data.deletedQuotations || [];
  eq(dq.length, 1, '[B] deletedQuotations 有一筆');
  const rec = dq[0] || {};
  eq([rec.id, rec.quoteNo, rec.deletedBy, rec.reason], [ids[2], 'QU-202610-003', 'own1', '客戶取消，重新報價'], '[B] 快照：id／單號／刪除者／原因');
  ok(!isNaN(Date.parse(rec.deletedAt)), '[B] 快照有 deletedAt');
  eq(rec.quotation, before, '[B] 快照是刪除當下的完整報價單（含 timeline／costFlow／items，位元級相同）');
  ok(Array.isArray(rec.quotation.timeline) && rec.quotation.timeline[0].action === 'DRAFT_CREATE' && !!rec.quotation.costFlow && rec.quotation.items[0].cost === 100, '[B] 快照內含 timeline、costFlow、成本（內部資料不去敏）');
  eq(rec.snapshotHash, QL.snapshotHash(before), '[B] snapshotHash＝快照的 sha256（與鍵順序無關）');
  const e3 = env.data.quoteNoLedger.find((e) => e.no === 'QU-202610-003');
  eq([e3.status, e3.deletedBy, e3.snapshotHash], ['deleted', 'own1', rec.snapshotHash], '[B] 流水帳標為 deleted 並帶雜湊');
  const log = env.logs.filter((l) => l[0] === 'DELETE_QUOTATION').pop() || [];
  ok(log[2] === 'QU-202610-003' && /業務 own1/.test(log[3]) && /未稅 1000/.test(log[3]) && /建立 \d{4}-/.test(log[3]) && /簽核 none/.test(log[3]) && log[3].includes(rec.snapshotHash.slice(0, 16)) && /原因：客戶取消/.test(log[3]), '[B] 稽核紀錄含單號／業務／未稅金額／建立時間／簽核狀態／雜湊／原因：' + log[3]);

  // 對一般使用者完全不可見
  r = await env.call('own1', 'GET', '/api/quotations');
  ok(r.s === 200 && !J(r.j).includes('QU-202610-003') && !J(r.j).includes('deletedQuotations'), '[B] 一般清單看不到已刪除的單');
  r = await env.call('own1', 'GET', '/api/quotations/:id', { id: ids[2] });
  eq(r.s, 404, '[B] 業務以 id 取已刪單 → 404');
  r = await env.call('own1', 'GET', '/api/admin/deleted-quotations');
  eq(r.s, 403, '[B] 非管理員讀已刪除清單 403');
  r = await env.call('own1', 'GET', '/api/admin/quote-ledger');
  eq(r.s, 403, '[B] 非管理員讀流水帳 403');
  r = await env.call('admin1', 'GET', '/api/admin/deleted-quotations');
  ok(r.s === 200 && r.j.total === 1 && r.j.items[0].quoteNo === 'QU-202610-003' && r.j.items[0].snapshotHash === rec.snapshotHash && !('quotation' in r.j.items[0]), '[B] 管理員清單：摘要、不含整張內容');
  r = await env.call('admin1', 'GET', '/api/admin/deleted-quotations/:id', { id: ids[2] });
  ok(r.s === 200 && r.j.hashOk === true && r.j.quotation.items[0].note === '稽核測試品項備註', '[B] 管理員取單張快照：hashOk');
  rec.quotation.company = '被竄改';
  r = await env.call('admin1', 'GET', '/api/admin/deleted-quotations/:id', { id: ids[2] });
  eq(r.j.hashOk, false, '[B] 快照被竄改 → hashOk=false');
  rec.quotation.company = before.company;
  // 其他人的統計：getViewableOwners 的主管看清單也沒有
  r = await env.call('mgr1', 'GET', '/api/quotations');
  ok(!J(r.j).includes('QU-202610-003'), '[B] 主管清單同樣看不到');

  // 流水帳端點：刪除的號不是缺號
  r = await env.call('admin1', 'GET', '/api/admin/quote-ledger');
  eq(r.s, 200, '[B] 流水帳 200');
  eq(r.j.summary, { total: 5, active: 4, invalid: 0, deleted: 1, missing: 0, backfilled: 0, gapUnexplained: 0 }, '[B] 摘要：5 筆、4 有效、1 已刪、0 缺號');
  eq(r.j.months, [{ ym: '202610', max: 5, known: 5, gaps: [] }], '[B] 月份：最大號 5、無缺號');
  eq(r.j.entries.find((e) => e.no === 'QU-202610-003').status, 'deleted', '[B] 該號顯示 deleted');

  // 舊式硬刪（沒有快照、沒有流水帳）→ 缺號
  env.data.quotations = env.data.quotations.filter((q) => q.quoteNo !== 'QU-202610-002');
  env.data.quoteNoLedger = env.data.quoteNoLedger.filter((e) => e.no !== 'QU-202610-002');
  env.data.quoteSeq['202610'] = 7;   // 計數器走到 7，但 6、7 從沒出現過
  r = await env.call('admin1', 'GET', '/api/admin/quote-ledger');
  eq(r.j.months[0].gaps.map((g) => g.no), ['QU-202610-002', 'QU-202610-006', 'QU-202610-007'], '[B] 缺號：002（被硬刪且無紀錄）與計數器 6、7');
  ok(r.j.months[0].gaps.every((g) => g.status === 'GAP_UNEXPLAINED'), '[B] 缺號標 GAP_UNEXPLAINED');
  eq(r.j.summary.gapUnexplained, 3, '[B] 缺號數');

  // 舊單（沒有流水帳項目）：檢視時惰性補登
  const legacyQ = { id: 'LQ1', quoteNo: 'QU-202609-001', owner: 'own2', company: 'LegacyCo', status: 'draft', createdAt: '2026-09-01T00:00:00Z', products: [], items: mkItems(), approval: null };
  env.data.quotations.push(legacyQ);
  const saves0 = env.saves;
  r = await env.call('admin1', 'GET', '/api/admin/quote-ledger');
  const le = (env.data.quoteNoLedger || []).find((e) => e.no === 'QU-202609-001');
  ok(le && le.backfilled === true && le.owner === 'own2' && le.createdAt === '2026-09-01T00:00:00Z' && le.status === 'active', '[B] 舊單檢視時補登（backfilled=true，owner／createdAt 取自單據）');
  ok(env.saves > saves0, '[B] 補登後有存檔');
  const s1 = env.saves;
  await env.call('admin1', 'GET', '/api/admin/quote-ledger');
  eq(env.saves, s1, '[B] 沒有新補登就不重複存檔');
  // 舊單被刪（刪除時才補登）
  const legacy2 = { id: 'LQ2', quoteNo: 'QU-202609-002', owner: 'own1', company: 'Legacy2', status: 'draft', createdAt: '2026-09-02T00:00:00Z', products: [], items: mkItems(), approval: null };
  env.data.quotations.push(legacy2);
  r = await env.call('own1', 'DELETE', '/api/quotations/:id', { id: 'LQ2' }, {});
  eq(r.s, 200, '[B] 刪除舊單 200');
  const le2 = env.data.quoteNoLedger.find((e) => e.no === 'QU-202609-002');
  ok(le2 && le2.status === 'deleted' && le2.backfilled === true, '[B] 舊單刪除時補登並標 deleted');
  ok(env.data.deletedQuotations.some((d) => d.id === 'LQ2' && d.reason === ''), '[B] 沒給原因 → reason 空字串');

  // 作廢／還原同步流水帳
  const nowI = '2026-10-09T00:00:00.000Z';
  const q1 = env.data.quotations.find((q) => q.quoteNo === 'QU-202610-001'); q1.oppId = 'O1';
  OL.invalidateForOpp(env.data, { id: 'O1', company: 'TestCo' }, 'own1', nowI);
  eq(env.data.quoteNoLedger.find((e) => e.no === 'QU-202610-001').status, 'invalid', '[B] 商機作廢 → 流水帳 invalid');
  r = await env.call('admin1', 'GET', '/api/admin/quote-ledger');
  eq(r.j.entries.find((e) => e.no === 'QU-202610-001').status, 'invalid', '[B] 端點顯示 invalid');
  OL.revalidateForOpp(env.data, { id: 'O1' }, 'own1', nowI);
  eq(env.data.quoteNoLedger.find((e) => e.no === 'QU-202610-001').status, 'active', '[B] 商機還原 → 流水帳回 active');
  // 已送簽的單仍不可刪，且不動快照
  const nDel = env.data.deletedQuotations.length;
  const q4 = env.data.quotations.find((q) => q.quoteNo === 'QU-202610-004');
  q4.approval = { state: 'pending', history: [{ action: 'SUBMIT' }], steps: [], cur: 0 };
  r = await env.call('own1', 'DELETE', '/api/quotations/:id', { id: q4.id }, {});
  ok(r.s === 409 && env.data.deletedQuotations.length === nDel && env.data.quotations.includes(q4), '[B] 已送簽核 → 409，不搬快照');
  // canonical 與鍵順序無關
  eq(QL.snapshotHash({ a: 1, b: { c: 2, d: [1, { e: 3, f: 4 }] } }), QL.snapshotHash({ b: { d: [1, { f: 4, e: 3 }], c: 2 }, a: 1 }), '[B] 雜湊與鍵順序無關');
  ok(QL.snapshotHash({ a: 1 }) !== QL.snapshotHash({ a: 2 }), '[B] 內容不同雜湊不同');
}

module.exports = {
  mkEnv, create, mkItems, USERS,
  sections(c) {
    return {
      B: async () => { console.log('[B] 刪除快照＋單號流水帳'); await runB(c); },
    };
  },
  seed: {},
  noServer: ['B', 'C', 'D', 'E'],
  async mutations(h) { void h; },
};
