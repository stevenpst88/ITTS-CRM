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
function mkEnv(routesPath, quotations, extraData, roster, opts) {
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
    db: Object.assign({ load: () => env.data, save: () => { env.saves++; }, flush: async () => {} }, (opts && opts.noAudit) ? {} : { insertAuditEvent: (ev) => { if (env.auditThrow) throw new Error('audit boom'); env.events.push(ev); return Promise.resolve(true); } }),
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

// ═════════════ C：規則快照＋版本釘選 ═════════════
const mkSubmitQuote = (over) => Object.assign({
  id: 'Q1', quoteNo: 'QU-202610-901', owner: 'own1', company: 'TestCo', projectName: 'Proj', quoteDate: '2026-10-08', status: 'draft', createdAt: '2026-10-08T00:00:00Z', updatedAt: '2026-10-08T00:00:00Z',
  validUntil: '2026-10-30', products: ['PS'], costBy: null, approval: null, discountType: 'none', discountValue: 0,
  items: [{ lid: 'i1', desc: '導入顧問', unit: '式', qty: 1, unitPrice: 600000, cost: 100000 }, { lid: 'i2', desc: '教育訓練', unit: '式', qty: 1, unitPrice: 100000, cost: 20000 }],
  costFlow: { state: 'na' },
}, over || {});

/** 從 client/quote-approval.js 取出 buildRulesDetails，用最小替身執行（不需要瀏覽器） */
function loadRulesDetails() {
  const src = fs.readFileSync(path.join(ROOT, '_client/quote-approval.js'), 'utf8');
  const i = src.indexOf('function buildRulesDetails');
  if (i < 0) return null;
  const j = src.indexOf('\n}\n', i);
  const body = src.slice(i, j + 3);
  const e = (v) => String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const TL = { mgr1: '一級主管', mgr2: '二級主管複核', gm: '總經理', chairman: '董事長', board: '董事會決議（秘書代核）' };
  return new Function('e', 'tierLabel', 'fmtNum', 'fmtTime', body + '\nreturn buildRulesDetails;')(e, (k) => TL[k] || String(k || ''), (n) => Number(n).toLocaleString('en-US'), (v) => String(v));
}

async function runC(h, routesPath) {
  const { ok, eq } = h;
  const QA = require(path.join(ROOT, 'lib/quoteApproval.js'));
  const { spawnSync } = require('child_process');
  let env = mkEnv(routesPath, [mkSubmitQuote()]);
  const Q = () => env.data.quotations[0];
  const hashBefore = QA.contentHash(Q());
  let r = await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' });
  eq(r.s, 200, '[C] 送簽 200（' + (r.j && r.j.error) + '）');
  const ap = Q().approval, d = ap.derived;
  const th = d.thresholds;
  ok(th && typeof th === 'object', '[C] ap.derived.thresholds 存在');
  eq([th.rowKey, th.l1, th.l2], ['software', 10, 5], '[C] 類別門檻：software 10／5');
  eq(th.amountThresholds, { chairman: 10000000, board: 50000000 }, '[C] 金額門檻');
  eq(th.path, ['mgr1'], '[C] 財務核決路徑');
  eq(th.actualPath, ['mgr1'], '[C] 實際簽核路徑');
  eq([th.rulesVersion, th.rulesHash], [QA.RULES_VERSION, QA.RULES_HASH], '[C] 規則版本與雜湊');
  eq(th.rosterSnapshot.gm, ['gm1'], '[C] 名冊快照 gm');
  eq(th.rosterSnapshot.chairman, ['ch1'], '[C] 名冊快照 chairman');
  eq(th.rosterSnapshot.boardProxy, ['sec1'], '[C] 名冊快照 boardProxy');
  eq(th.rosterSnapshot.at, ap.submittedAt, '[C] 名冊快照時間＝送簽時間');
  eq(th.flowSnapshot, { buKey: 'ITS', hasMgr2Step: false, mgr2Assignee: null, skipMgr1Applied: false }, '[C] 流程快照（無二級關）');
  // 既有 derived 欄位完全不變
  const base = QA.buildDerived(Q(), CLASSES);
  const { thresholds, ...rest } = d;
  eq(rest, base, '[C] derived 除 thresholds 外與 QA.buildDerived 位元級相同（核決運算沒動）');
  eq(QA.contentHash(Q()), hashBefore, '[C] contentHash 不受影響（快照不進簽章）');
  eq(ap.steps.map((x) => x.tier), ['mgr1'], '[C] 簽核關卡與以往相同');
  // 名冊之後怎麼改，快照不動；SUBMIT 歷程 meta 也留一份
  env.data.quoteApproval.roster.gm = ['gm2'];
  r = await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' });
  eq(r.j.approval.derived.thresholds.rosterSnapshot.gm, ['gm1'], '[C] 名冊事後變更 → 快照仍是送簽當時的 gm1');
  ok(r.j.approval.derived.rosterNames && r.j.approval.derived.rosterNames.gm1 === 'Gm', '[C] 序列化帶 rosterNames（帳號→顯示名稱）');
  ok(!('thresholds' in r.j.approval.history[0]) && !('meta' in r.j.approval.history[0]), '[C] 歷程輸出不外洩 meta');
  const sub = Q().approval.history.find((x) => x.action === 'SUBMIT');
  eq(sub.meta.thresholds.rosterSnapshot.gm, ['gm1'], '[C] SUBMIT 歷程 meta 也留有快照');
  env.data.quoteApproval.roster.gm = ['gm1'];
  // 撤回後 derived 被清掉，但歷程快照還在
  r = await env.call('own1', 'POST', '/api/quotations/:id/withdraw', { id: 'Q1' });
  ok(r.s === 200 && Q().approval.derived === null && Q().approval.history.find((x) => x.action === 'SUBMIT').meta.thresholds.rulesHash === QA.RULES_HASH, '[C] 撤回後 derived=null，SUBMIT meta 快照仍在');

  // 長路徑＋董事會（低毛利 → chairman 路徑）
  env = mkEnv(routesPath, [mkSubmitQuote({ items: [{ lid: 'i1', desc: 'x', unit: '式', qty: 1, unitPrice: 1000000, cost: 960000 }] })]);
  r = await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' });
  const th2 = env.data.quotations[0].approval && env.data.quotations[0].approval.derived.thresholds;
  eq(r.s, 200, '[C] 低毛利送簽 200');
  eq(th2 && th2.path, ['mgr1', 'gm', 'chairman'], '[C] 低毛利 → 董事長路徑');
  // 二級主管流程與略過一級主管例外
  const mk2 = (skip) => {
    const e2 = mkEnv(routesPath, [mkSubmitQuote({ owner: 'own2' })]);
    e2.auth.users.push({ username: 'mgr2x', role: 'manager2', active: true, displayName: 'Mgr2X', bu: ['ITS'], supervisor: 'mgr1' });
    e2.auth.users.find((u) => u.username === 'own2').supervisor = 'mgr2x';
    e2.data.approvalFlowConfig = { buSettings: { ITS: { hasMgr2Step: true } }, exceptions: skip ? [{ username: 'mgr2x', action: 'skip_mgr1' }] : [] };
    return e2;
  };
  let e2 = mk2(false);
  r = await e2.call('own2', 'POST', '/api/quotations/:id/submit', { id: 'Q1' });
  let t3 = e2.data.quotations[0].approval.derived.thresholds;
  eq([r.s, t3.path, t3.actualPath], [200, ['mgr1'], ['mgr2', 'mgr1']], '[C] 二級主管流程：財務路徑 mgr1、實際 mgr2→mgr1');
  eq(t3.flowSnapshot, { buKey: 'ITS', hasMgr2Step: true, mgr2Assignee: 'mgr2x', skipMgr1Applied: false }, '[C] 流程快照：二級關啟用、指派 mgr2x');
  e2 = mk2(true);
  r = await e2.call('own2', 'POST', '/api/quotations/:id/submit', { id: 'Q1' });
  t3 = e2.data.quotations[0].approval.derived.thresholds;
  eq([r.s, t3.actualPath, t3.flowSnapshot.skipMgr1Applied], [200, ['mgr2'], true], '[C] 套用 skip_mgr1 例外 → 實際只剩 mgr2，flowSnapshot.skipMgr1Applied=true');
  // 舊單（沒有快照）照常顯示
  env = mkEnv(routesPath, [mkSubmitQuote()]);
  await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' });
  delete env.data.quotations[0].approval.derived.thresholds;
  r = await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' });
  ok(r.s === 200 && r.j.approval.derived && !('thresholds' in r.j.approval.derived) && !('rosterNames' in r.j.approval.derived) && r.j.approval.derived.rowKey === 'software', '[C] 舊單（無快照）序列化正常且不多鍵');
  r = await env.call('mgr1', 'POST', '/api/quotations/:id/approve', { id: 'Q1' }, { hash: QA.contentHash(env.data.quotations[0]) });
  ok(r.s === 200, '[C] 舊單（無快照）仍可核准 (' + r.s + ' ' + J(r.j) + ')');
  // 前端
  const fn = loadRulesDetails();
  ok(typeof fn === 'function', '[C] 前端有 buildRulesDetails');
  if (fn) {
    const dd = { rowLabel: '軟體規劃服務', rosterNames: { gm1: '王總' }, thresholds: th };
    const html = fn({ frozen: true, d: dd });
    ok(/<details/.test(html) && /適用規則/.test(html) && html.includes(QA.RULES_VERSION) && /≥ 10%/.test(html) && /≥ 5%/.test(html) && /10,000,000/.test(html) && /50,000,000/.test(html) && /王總/.test(html), '[C] 前端：可收合「適用規則」含版本／門檻／金額／名冊（顯示名稱）');
    ok(/一級主管/.test(html) && /BU ITS/.test(html), '[C] 前端：路徑與流程設定');
    const xss = fn({ frozen: true, d: { rowLabel: '<img onerror=1>', thresholds: Object.assign({}, th, { rosterSnapshot: { gm: ['<script>x</script>'], chairman: [], boardProxy: [], at: null } }) } });
    ok(!/<script>|<img onerror/.test(xss), '[C] 前端：名冊與類別文字跳脫');
    ok(/data-rules="none"/.test(fn({ frozen: true, d: { rowLabel: 'x' } })), '[C] 前端：舊單沒有快照 → 一行說明');
    eq(fn({ frozen: false, d: dd }), '', '[C] 前端：即時試算（未送簽）不顯示');
    eq(fn(null), '', '[C] 前端：沒有判定資料不顯示');
    const full = fs.readFileSync(path.join(ROOT, '_client/quote-approval.js'), 'utf8');
    ok(/html \+= buildRulesDetails\(pk\);/.test(full), '[C] 前端：已掛進判定區塊');
  }
  // 釘選腳本
  const pin = path.join(ROOT, 'scripts/check-quote-rules-pin.js');
  const run = (lib) => spawnSync(process.execPath, [pin], { env: Object.assign({}, process.env, lib ? { QRP_LIB: lib } : {}), encoding: 'utf8' });
  let pr = run();
  ok(pr.status === 0 && /規則內容與釘選一致/.test(pr.stdout), '[C] 釘選腳本：目前規則通過');
  ok(typeof QA.RULES_HASH === 'string' && QA.RULES_HASH.length === 64, '[C] lib 匯出 64 碼 RULES_HASH');
  const muts = [
    ['l1', "{ key: 'consult', label: '顧問服務', l1: 25, l2: 15 }", "{ key: 'consult', label: '顧問服務', l1: 24, l2: 15 }"],
    ['amt', 'const AMOUNT = Object.freeze({ chairman: 10000000, board: 50000000 });', 'const AMOUNT = Object.freeze({ chairman: 10000000, board: 60000000 });'],
    ['path', "2: Object.freeze(['mgr1', 'gm']),", "2: Object.freeze(['gm']),"],
    ['hash', "const RULES_HASH = '", "const RULES_HASH = '0"],
  ];
  const libFile = path.join(ROOT, 'lib/quoteApproval.js');
  for (const [tag, from, to] of muts) {
    const src = fs.readFileSync(libFile, 'utf8');
    if (!src.includes(from)) throw new Error('釘選變異樣式找不到：' + tag);
    const mf = path.join(ROOT, 'lib', 'quoteApproval.__mut_pin_' + tag + '.js');
    fs.writeFileSync(mf, src.replace(from, () => to));
    try {
      const m = run(mf);
      ok(m.status === 1 && m.stdout.includes('規則內容已變動，請更新 RULES_VERSION 與 RULES_HASH'), `[C] 釘選腳本變異 ${tag}（規則被改）→ 失敗並印出指定訊息`);
    } finally { try { fs.unlinkSync(mf); } catch { /* 已刪 */ } }
  }
}


// ═════════════ D：編輯事件完整性 ═════════════
async function runD(h, routesPath) {
  const { ok, eq } = h;
  const PUTQ = '/api/quotations/:id';
  const env = mkEnv(routesPath, []);
  const evs = () => env.events.filter((x) => x.kind === 'quote-edit');
  let r = await create(env, 'O1');
  eq(r.s, 201, '[D] 建立 201');
  const q = () => env.data.quotations[0];
  const id = q().id;
  let e0 = evs();
  eq(e0.length, 1, '[D] 建立 → 1 筆原始編輯事件');
  eq([e0[0].action, e0[0].operator, e0[0].target, e0[0].extra.count, e0[0].extra.quoteNo], ['DRAFT_CREATE', 'own1', q().quoteNo, 1, q().quoteNo], '[D] 事件欄位：動作／操作者／單號／count=1');
  ok(/建立報價單/.test(e0[0].detail) && !isNaN(Date.parse(e0[0].ts)), '[D] 事件含摘要與時間');
  // 連續 6 次儲存（60 秒內）：時間軸合併成 1 筆，原始事件仍是 6 筆
  for (let i = 0; i < 6; i++) {
    const rr = await env.call('own1', 'PUT', PUTQ, { id }, { note: 'n' + i, items: q().items.map((x) => ({ lid: x.lid, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice, cost: x.cost })) });
    ok(rr.s === 200, '[D] 儲存 ' + i + ' 200（' + rr.s + '）');
  }
  eq(q().timeline.length, 2, '[D] timeline：DRAFT_CREATE＋合併後的 1 筆 DRAFT_SAVE');
  eq(q().timeline[1].count, 6, '[D] timeline 合併 count=6');
  const saves = evs().filter((x) => x.action === 'DRAFT_SAVE');
  eq(saves.length, 6, '[D] 原始事件不合併：6 筆 DRAFT_SAVE');
  ok(saves.every((x) => x.extra.count === 1 && x.operator === 'own1' && x.target === q().quoteNo), '[D] 每筆 count=1、帶操作者與單號');
  ok(saves.every((x) => /備註/.test(x.detail)), '[D] 每筆都有各自的完整摘要');
  ok(saves.every((x, i, a) => i === 0 || Date.parse(x.ts) >= Date.parse(a[i - 1].ts)), '[D] 事件時間遞增');

  // 上限：200 筆；DRAFT_CREATE 釘住；標記累計省略筆數
  const base = Date.parse('2026-01-01T00:00:00Z');
  q().timeline = [q().timeline[0]].concat(Array.from({ length: 250 }, (_, i) => ({ at: new Date(base + (i + 1) * 120000).toISOString(), by: 'own1', action: 'DRAFT_SAVE', comment: 'e' + i })));
  const nEv = evs().length;
  await env.call('own1', 'PUT', PUTQ, { id }, { note: 'cap1', items: q().items.map((x) => ({ lid: x.lid, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice, cost: x.cost })) });
  let tl = q().timeline;
  eq(tl.length, 200, '[D] 超過上限 → 剛好 200 筆（含標記）');
  eq(tl[0].action, 'DRAFT_CREATE', '[D] DRAFT_CREATE 釘在最前面');
  eq(tl[1].action, 'TIMELINE_TRIMMED', '[D] 第二筆是省略標記');
  eq(tl[1].comment, '（已省略 53 筆較早事件，完整紀錄見稽核日誌）', '[D] 標記文字與筆數（251 筆一般事件保留 198 筆 → 省略 53）');
  eq(tl[1].omitted, 53, '[D] 標記 omitted=53');
  eq(tl[2].comment, 'e53', '[D] 保留的是最新的事件');
  eq(tl.filter((x) => x.action === 'TIMELINE_TRIMMED').length, 1, '[D] 只有一個標記');
  eq(evs().length, nEv + 1, '[D] 修剪不影響稽核事件表（照常多 1 筆）');
  // 再來一筆（避開合併）：標記累計
  q().timeline[q().timeline.length - 1].at = new Date(Date.now() - 120000).toISOString();
  await env.call('own1', 'PUT', PUTQ, { id }, { note: 'cap2', items: q().items.map((x) => ({ lid: x.lid, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice, cost: x.cost })) });
  tl = q().timeline;
  eq([tl.length, tl[0].action, tl.filter((x) => x.action === 'TIMELINE_TRIMMED').length, tl[1].omitted], [200, 'DRAFT_CREATE', 1, 54], '[D] 再超過一筆 → 標記累計 54、仍只有一個、仍 200 筆');
  eq(tl[1].comment, '（已省略 54 筆較早事件，完整紀錄見稽核日誌）', '[D] 標記文字更新');
  // 長時間大量事件：DRAFT_CREATE 永遠在
  for (let i = 0; i < 300; i++) {
    q().timeline[q().timeline.length - 1].at = new Date(Date.now() - 120000).toISOString();
    await env.call('own1', 'PUT', PUTQ, { id }, { note: 'bulk' + i, items: q().items.map((x) => ({ lid: x.lid, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice, cost: x.cost })) });
  }
  tl = q().timeline;
  ok(tl.length === 200 && tl[0].action === 'DRAFT_CREATE' && tl.filter((x) => x.action === 'DRAFT_CREATE').length === 1 && tl.filter((x) => x.action === 'TIMELINE_TRIMMED').length === 1, '[D] 再 300 次儲存後：DRAFT_CREATE 仍在、標記仍 1 個、共 200 筆');
  eq(tl[1].omitted, 54 + 300, '[D] 標記累計 354');
  ok(evs().filter((x) => x.action === 'DRAFT_SAVE').length === 6 + 1 + 1 + 300, '[D] 稽核事件表：全部 DRAFT_SAVE 一筆不缺（308）');
  // 序列化
  r = await env.call('own1', 'GET', PUTQ, { id });
  const mk = r.j.timeline.find((x) => x.action === 'TIMELINE_TRIMMED');
  ok(mk && mk.byName === '系統' && /已省略/.test(mk.comment) && mk.kind === 'edit', '[D] 序列化：標記事件 byName=系統');
  ok(r.j.timeline[0].action === 'DRAFT_CREATE', '[D] 序列化：第一筆仍是建立');
  // 不影響簽核歷程
  eq(q().approval, null, '[D] 沒有送簽；approval 不受影響');
  // 顧問成本事件也進原始事件
  const env2 = mkEnv(routesPath, [mkSubmitQuote({ products: ['PC'], costBy: 'cons1', costFlow: { state: 'requested', requestedAt: '2026-10-08T00:00:00Z' } })]);
  r = await env2.call('cons1', 'PUT', '/api/quotations/:id/costs', { id: 'Q1' }, { costLines: [{ cat: 'consult', desc: 'PM', qty: 2, unitCost: 7000, unit: '人天', rel: 'link', forLid: 'i1' }], done: false, costModel: 2 });
  ok(r.s === 200 && env2.events.some((x) => x.kind === 'quote-edit' && x.action === 'COST_SAVE' && x.operator === 'cons1'), '[D] 顧問儲存成本 → 原始事件 COST_SAVE（' + r.s + '）');
  // 事件寫入丟例外、或根本沒有 insertAuditEvent，都不影響請求
  const env3 = mkEnv(routesPath, []);
  env3.auditThrow = true;
  r = await create(env3, 'O1');
  ok(r.s === 201 && env3.data.quotations[0].timeline.length === 1, '[D] insertAuditEvent 丟例外 → 建單照樣 201、時間軸照常');
  const env4 = mkEnv(routesPath, [], null, null, { noAudit: true });
  r = await create(env4, 'O1');
  ok(r.s === 201 && env4.data.quotations[0].timeline.length === 1, '[D] db 沒有 insertAuditEvent（舊環境）→ 照常');
}


// ═════════════ E：董事會代核證據 ═════════════
/** 把一張單推進到「輪到董事會」的關卡（金額 > 5,000 萬 → mgr1→gm→board） */
async function toBoardStep(env, QA) {
  const r = await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' });
  if (r.s !== 200) throw new Error('送簽失敗 ' + r.s + ' ' + J(r.j));
  const q = () => env.data.quotations[0];
  for (const u of ['mgr1', 'gm1']) {
    const a = await env.call(u, 'POST', '/api/quotations/:id/approve', { id: 'Q1' }, { hash: QA.contentHash(q()) });
    if (a.s !== 200) throw new Error('前置核准失敗 ' + u + ' ' + a.s + ' ' + J(a.j));
  }
  return q;
}
async function runE(h, routesPath) {
  const { ok, eq } = h;
  const QA = require(path.join(ROOT, 'lib/quoteApproval.js'));
  const reg = require(routesPath || path.join(ROOT, 'lib/quoteRoutes.js'));
  const big = () => mkSubmitQuote({ items: [{ lid: 'i1', desc: '大案', unit: '式', qty: 1, unitPrice: 60000000, cost: 1000000 }] });
  const approve = (env, q, body, user) => env.call(user || 'sec1', 'POST', '/api/quotations/:id/approve', { id: 'Q1' }, Object.assign({ hash: QA.contentHash(q()), resolutionDate: '2026-10-01' }, body));
  let env = mkEnv(routesPath, [big()]);
  let q = await toBoardStep(env, QA);
  eq(q().approval.steps.map((x) => x.tier), ['mgr1', 'gm', 'board'], '[E] 大額單路徑 mgr1→gm→board');

  // 格式驗證（每個失敗都不得推進狀態）
  const bad = [
    ['', 'BAD_RESOLUTION', '空白'],
    ['T-1', 'BAD_RESOLUTION_NO_LENGTH', '太短(3)'],
    ['A'.repeat(41) + '1', 'BAD_RESOLUTION_NO_LENGTH', '太長(42)'],
    ['ABCD', 'BAD_RESOLUTION_NO_DIGIT', '沒有數字'],
    ['113 董 045', 'BAD_RESOLUTION_NO_CHARS', '含空白'],
    ['113-董-045!', 'BAD_RESOLUTION_NO_CHARS', '含驚嘆號'],
    ['<b>113</b>', 'BAD_RESOLUTION_NO_CHARS', '含 HTML'],
    ['113😀號', 'BAD_RESOLUTION_NO_CHARS', '含表情符號'],
    ['113#045', 'BAD_RESOLUTION_NO_CHARS', '含 #'],
    ['113.045', 'BAD_RESOLUTION_NO_CHARS', '含句點'],
  ];
  for (const [no, code, label] of bad) {
    const r = await approve(env, q, { resolutionNo: no });
    ok(r.s === 400 && r.j.code === code && /董事會決議文號/.test(r.j.error), `[E] 文號 ${label} → 400 ${code}（得到 ${r.s} ${r.j && r.j.code}）`);
    ok(q().approval.cur === 2 && q().approval.board === null && q().approval.state === 'pending', `[E] 文號 ${label}：狀態沒有被推進`);
  }
  // 日期錯誤優先於文號（原有行為不變）
  let r = await approve(env, q, { resolutionNo: '113-董-045', resolutionDate: '2099-01-01' });
  ok(r.s === 400 && r.j.code === 'BAD_RESOLUTION', '[E] 未來日期仍擋（原有行為）');
  r = await approve(env, q, { resolutionNo: '113-董-045', resolutionDate: '2026-02-31' });
  ok(r.s === 400 && r.j.code === 'BAD_RESOLUTION', '[E] 不存在的日期仍擋（原有行為）');
  // 議事錄欄位
  r = await approve(env, q, { resolutionNo: '113-董-045', minutesRef: 'x'.repeat(101) });
  ok(r.s === 400 && r.j.code === 'BAD_MINUTES_REF', '[E] 議事錄 101 字 → 400 BAD_MINUTES_REF');
  ok(q().approval.cur === 2, '[E] 議事錄錯誤：狀態沒推進');
  // 非有權者
  r = await approve(env, q, { resolutionNo: '113-董-045' }, 'own1');
  ok(r.s !== 200 && q().approval.cur === 2, '[E] 非代核人（業務本人）不能代核（原有權限不變）(' + r.s + ')');
  r = await approve(env, q, { resolutionNo: '113-董-045' }, 'gm1');
  ok(r.s !== 200 && q().approval.cur === 2, '[E] 總經理也不能代董事會登錄（原有權限不變）(' + r.s + ')');
  // 合格：含 minutesRef（含控制字元會被清掉）
  const logsBefore = env.logs.length;
  r = await approve(env, q, { resolutionNo: ' 113-董-045 ', minutesRef: ' 第12次董事會議事錄 第3案\t附件A ' });
  eq(r.s, 200, '[E] 合格文號＋議事錄 → 200（' + (r.j && r.j.error) + '）');
  const bd = q().approval.board;
  eq([bd.resolutionNo, bd.minutesRef, bd.resolutionDate, bd.by], ['113-董-045', '第12次董事會議事錄 第3案 附件A', '2026-10-01', 'sec1'], '[E] ap.board 內容（文號 trim、議事錄清掉控制字元）');
  const hist = q().approval.history.filter((x) => x.action === 'APPROVE').pop();
  eq([hist.meta.resolutionNo, hist.meta.minutesRef, hist.meta.resolutionDate], ['113-董-045', '第12次董事會議事錄 第3案 附件A', '2026-10-01'], '[E] 歷程 meta 含文號／議事錄／日期');
  const lg = env.logs.slice(logsBefore).find((l) => l[0] === 'APPROVE_QUOTATION');
  ok(lg && /決議 113-董-045/.test(lg[3]) && /議事錄 第12次董事會議事錄 第3案 附件A/.test(lg[3]), '[E] 稽核紀錄含議事錄：' + (lg && lg[3]));
  eq(q().approval.state, 'approved', '[E] 全部簽完 → approved');
  const g = await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' });
  eq(g.j.approval.board.minutesRef, '第12次董事會議事錄 第3案 附件A', '[E] 序列化 board.minutesRef');
  eq(g.j.approval.board.resolutionNo, '113-董-045', '[E] 序列化 board.resolutionNo');
  // 不填 minutesRef：形狀與舊資料相同（沒有這個鍵）
  env = mkEnv(routesPath, [big()]);
  q = await toBoardStep(env, QA);
  r = await approve(env, q, { resolutionNo: '(113)董/045_A' });
  ok(r.s === 200 && !('minutesRef' in q().approval.board), '[E] 沒填議事錄 → ap.board 沒有 minutesRef 鍵；允許字元 ( ) / _ 與文字 (' + r.s + ')');
  const h2 = q().approval.history.filter((x) => x.action === 'APPROVE').pop();
  ok(!('minutesRef' in h2.meta), '[E] 沒填議事錄 → 歷程 meta 也沒有該鍵');
  const lg2 = env.logs.filter((l) => l[0] === 'APPROVE_QUOTATION').pop();
  ok(!/議事錄/.test(lg2[3]), '[E] 沒填議事錄 → 稽核紀錄不多字');
  const g2 = await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' });
  ok(!('minutesRef' in g2.j.approval.board), '[E] 沒填議事錄 → 序列化沒有該鍵');
  // 非董事會關：多帶的欄位被忽略、不驗證
  env = mkEnv(routesPath, [big()]);
  await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' });
  r = await env.call('mgr1', 'POST', '/api/quotations/:id/approve', { id: 'Q1' }, { hash: QA.contentHash(env.data.quotations[0]), resolutionNo: '!!', minutesRef: 'y'.repeat(300) });
  ok(r.s === 200 && env.data.quotations[0].approval.board === null, '[E] 一般關卡核准不驗證也不存董事會欄位');
  // 舊資料（board 沒有 minutesRef）顯示不變
  const old = mkEnv(routesPath, [big()]);
  const oq = old.data.quotations[0];
  oq.approval = { state: 'approved', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(oq), submittedAt: '2026-10-01T00:00:00Z', submittedBy: 'own1', derived: null, steps: [], cur: 0, board: { resolutionDate: '2026-09-30', resolutionNo: 'T-1', by: 'sec1', at: '2026-10-01T00:00:00Z' }, history: [] };
  const og = await old.call('own1', 'GET', '/api/quotations/:id', { id: 'Q1' });
  eq(og.j.approval.board, { resolutionDate: '2026-09-30', resolutionNo: 'T-1', byName: 'Sec', at: '2026-10-01T00:00:00Z' }, '[E] 舊單（短文號、無議事錄）序列化不變');
  // 純函式與前端一致
  const cases = ['', 'T-1', 'ABCD', 'A1B2', '113-董-045', '113 董 045', '１１３號', '(1)/2_3', 'A'.repeat(40), 'A'.repeat(39) + '1', 'A'.repeat(41), '董事會決議', '決議123', '<script>1', '1234!', ' 1234 '];
  const src = fs.readFileSync(path.join(ROOT, '_client/quote-approval.js'), 'utf8');
  const i0 = src.indexOf('function resolutionNoProblem');
  const fnClient = i0 < 0 ? null : new Function(src.slice(i0, src.indexOf('\n}\n', i0) + 3) + '\nreturn resolutionNoProblem;')();
  ok(typeof fnClient === 'function', '[E] 前端有 resolutionNoProblem');
  if (fnClient) {
    const diff = cases.filter((c) => (reg.checkResolutionNo(c).ok) !== (fnClient(c) === ''));
    eq(diff, [], '[E] 前端驗證與伺服器對 ' + cases.length + ' 個樣本的結論一致');
  }
  eq(reg.checkResolutionNo('１１３號').ok, true, '[E] 全形數字視為數字（合格）');
  eq(reg.checkResolutionNo('A'.repeat(40)).ok, false, '[E] 40 個字母沒有數字 → 不合格');
  eq(reg.checkResolutionNo('A'.repeat(39) + '1').ok, true, '[E] 剛好 40 字且含數字 → 合格');
  eq(reg.checkResolutionNo('A12').code, 'BAD_RESOLUTION_NO_LENGTH', '[E] 3 字 → 長度錯誤');
  eq(reg.checkMinutesRef('x'.repeat(100)).ok, true, '[E] 議事錄 100 字合格');
  eq(reg.checkMinutesRef('😀'.repeat(100)).ok, true, '[E] 議事錄以字元計（100 個表情符號合格）');
  // 前端靜態：對話框欄位、確認文字、簽核進度顯示、列印
  ok(/data-draft="minutesRef"/.test(src) && /maxlength="100"/.test(src), '[E] 前端：董事會代核對話框有議事錄欄位（maxlength 100）');
  ok(/body\.minutesRef = mref/.test(src) && /議事錄／附件：' \+ mref/.test(src), '[E] 前端：送出與確認文字帶議事錄');
  ok(/ap\.board\.minutesRef \? '　議事錄／附件：'/.test(src), '[E] 前端：簽核進度顯示議事錄（經 e() 跳脫）');
  ok(/b\.minutesRef \?/.test(src), '[E] 前端：董事會簽呈列印含議事錄');
}


module.exports = {
  mkEnv, create, mkItems, USERS,
  sections(c) {
    return {
      B: async () => { console.log('[B] 刪除快照＋單號流水帳'); await runB(c); },
      D: async () => { console.log('[D] 編輯事件完整性'); await runD(c); },
      E: async () => { console.log('[E] 董事會代核證據'); await runE(c); },
      C: async () => { console.log('[C] 規則快照＋版本釘選'); await runC(c); },
    };
  },
  seed: {},
  noServer: ['B', 'C', 'D', 'E'],
  async mutations(h) { void h; },
};
