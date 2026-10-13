#!/usr/bin/env node
/**
 * 報價單「編輯事件時間軸」(q.timeline) 檢查。用法：node scripts/check-quote-timeline.js（不需要伺服器；假 app＋記憶體資料庫直接呼叫路由 handler；不碰 data.json／auth.json／audit.log.json）
 * 動 lib/quoteRoutes.js 的 pushEdit／editSummary／visibleTimeline／serialize、POST／PUT 報價單、PUT /costs、item-notes，或 _client/quote-approval.js 的歷程表之後必跑。
 *   1) 建立／儲存／指派顧問／顧問儲存／完成／退回／業務改品項退回 都留一筆（時間＋人員＋動作＋摘要）
 *   2) 60 秒內同人同動作合併（count＋時間更新）；超過 60 秒不合併；編輯事件上限 200（超過時放「已省略 N 筆」標記、DRAFT_CREATE 釘住，見 check-audit-wave2 的 D 項）且不影響 approval.history
 *   3) 不影響簽核：timeline 不進 contentHash／itemsSig／costLinesSig；存檔留事件後核准仍有效；送簽→核准照常；刪除草稿不被擋；approval.history 完全沒有新動作
 *   4) 舊單（沒有 timeline）序列化輸出不多任何鍵；清單 API 只給 timelineCount
 *   5) 前端靜態檢查：標籤、合併排序、舊單不變
 * 測試資料只用通用字串（公開 repo：不放客戶名、人名、真實費率）。
 */
'use strict';
const path = require('path');
const fs = require('fs');
const ROOT = path.join(__dirname, '..');
const QA = require(path.join(ROOT, 'lib/quoteApproval.js'));
const registerQuoteRoutes = require(path.join(ROOT, 'lib/quoteRoutes.js'));
const PNL = require(path.join(ROOT, 'lib/quotePnlExcel.js'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);
const L = (cat, desc, qty, unitCost, extra) => Object.assign({ cat, desc, qty, unitCost }, extra || {});
const EDIT = ['DRAFT_CREATE', 'DRAFT_SAVE', 'COST_REQUEST', 'COST_SAVE', 'COST_DONE', 'COST_REOPEN'];

const USERS = [
  { username: 'admin1', role: 'admin', active: true, displayName: 'Admin' },
  { username: 'own1', role: 'user', active: true, displayName: 'Owner1', supervisor: 'mgr1', bu: ['ITS'] },
  { username: 'mgr1', role: 'manager1', active: true, displayName: 'Mgr1', bu: ['ITS'] },
  { username: 'gm1', role: 'executive', active: true, displayName: 'Gm' },
  { username: 'ch1', role: 'executive', active: true, displayName: 'Chair' },
  { username: 'cons1', role: 'user', active: true, displayName: 'Cons1', bu: ['ITS'] },
  { username: 'sec1', role: 'secretary', active: true, displayName: 'Sec', bu: ['ITS'] },
];
const CLASSES = { PC: { cls: 'consult', costBySales: false }, PS: { cls: 'software', costBySales: true } };
function mkEnv(quotations) {
  const routes = {}; const app = {};
  ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
  const env = { data: { opportunities: [{ id: 'O1', owner: 'own1', stage: 'C', company: 'TestCo', product: 'Prod', amount: '100', bu: 'ITS' }], quotations, quoteApproval: { roster: { gm: ['gm1'], chairman: ['ch1'], boardProxy: [], costProviders: ['cons1'], sealManagers: [] }, productClasses: CLASSES } }, logs: [], notes: [], saves: 0, seq: 0 };
  env.auth = { users: JSON.parse(JSON.stringify(USERS)) };
  const deps = {
    db: { load: () => env.data, save: () => { env.saves++; }, flush: async () => {} }, loadAuth: () => env.auth, saveAuth: () => {},
    requireAuth: (req, rs, next) => next(), requireAdmin: (req, rs, next) => next(),
    writeLog: (...a) => env.logs.push(a), pushNotification: (...a) => env.notes.push(a),
    getViewableOwners: (req) => { const rl = req.session.user.role; if (rl === 'admin' || rl === 'executive' || rl === 'manager1' || rl === 'secretary') return ['own1']; return [req.session.user.username]; },
    sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-NEW', taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}),
    buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: PNL.buildQuotePnlExcel, QUOTE_TEMPLATE: '',
    uuidv4: () => 'u' + (++env.seq), normalizeBu: (b) => (Array.isArray(b) ? b : b ? [b] : []), getUserFeatures: () => ['quotations'],
  };
  registerQuoteRoutes(app, deps);
  env.call = (user, method, p, params, body, query) => new Promise((resolve, reject) => {
    const h = [].concat(...routes[method.toUpperCase() + ' ' + p]);
    const role = (env.auth.users.find((u) => u.username === user) || {}).role;
    const req = { session: { user: { username: user, role } }, params: params || {}, query: query || {}, body: body === undefined ? {} : JSON.parse(JSON.stringify(body)) };
    const rs = { status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j }); return this; }, send() { resolve({ s: this._s || 200, j: null }); return this; }, setHeader() {}, set() {} };
    let i = 0;
    const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
    next();
  });
  return env;
}
const mkQuote = (over) => Object.assign({
  id: 'Q1', quoteNo: 'QU-1', owner: 'own1', company: 'TestCo', projectName: 'Proj', quoteDate: '2026-10-08', status: 'draft', createdAt: '2026-10-08T00:00:00Z', updatedAt: '2026-10-08T00:00:00Z',
  validUntil: '2026-10-30', products: ['PS'], costBy: null, approval: null, discountType: 'none', discountValue: 0,
  items: [
    { lid: 'i1', desc: '導入顧問', unit: '式', qty: 1, unitPrice: 600000, cost: 100000 },
    { lid: 'i2', desc: '教育訓練', unit: '式', qty: 1, unitPrice: 100000, cost: 20000 },
  ],
  costFlow: { state: 'na' },
}, over || {});
const formItems = (q, fn) => { const a = q.items.map((x) => ({ lid: x.lid, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice, cost: x.cost })); if (fn) fn(a); return a; };
const PUTQ = '/api/quotations/:id', COSTS = '/api/quotations/:id/costs';

async function run() {
  // ═════════ 1) 建立／儲存 ═════════
  let env = mkEnv([]);
  let r = await env.call('own1', 'POST', '/api/quotations', {}, { oppId: 'O1', company: 'TestCo', products: ['PS'], items: [{ desc: '品項A', unit: '式', qty: 1, unitPrice: 1000, cost: 100 }] });
  const q = () => env.data.quotations[0];
  t('1.1 建立 → 201；timeline 第一筆 DRAFT_CREATE（人員＝own1、有時間、有摘要）', r.s === 201 && q().timeline.length === 1 && q().timeline[0].action === 'DRAFT_CREATE' && q().timeline[0].by === 'own1' && !isNaN(Date.parse(q().timeline[0].at)) && /建立報價單/.test(q().timeline[0].comment), J(q().timeline));
  t('1.2 回應帶 timeline（kind:edit、byName 顯示名稱）與 timelineCount', r.j.timeline && r.j.timeline[0].kind === 'edit' && r.j.timeline[0].byName === 'Owner1' && r.j.timelineCount === 1);
  const id = q().id;
  r = await env.call('own1', 'PUT', PUTQ, { id }, { projectName: 'NewName', items: formItems(q(), (a) => { a[0].qty = 3; a[0].unitPrice = 7777; }) });
  let tl = q().timeline;
  t('1.3 儲存修改 → DRAFT_SAVE，摘要只有欄位名稱（專案名稱、修改 1 列（數量、單價）），沒有金額／原文', r.s === 200 && tl.length === 2 && tl[1].action === 'DRAFT_SAVE' && /專案名稱/.test(tl[1].comment) && /修改 1 列（數量、單價）/.test(tl[1].comment) && !/7777|NewName/.test(tl[1].comment), J(tl[1]));
  r = await env.call('own1', 'PUT', PUTQ, { id }, { items: formItems(q(), (a) => { a[0].qty = 4; a.push({ desc: '新列', unit: '式', qty: 1, unitPrice: 5, cost: 1 }); }) });
  tl = q().timeline;
  t('1.4 60 秒內同人同動作再儲存 → 合併：筆數不變、count=2、摘要合併、時間更新', tl.length === 2 && tl[1].count === 2 && /新增 1 列/.test(tl[1].comment) && /專案名稱/.test(tl[1].comment), J(tl[1]));
  t('1.5 序列化：連續次數輸出 count', r.j.timeline[1].count === 2);
  tl[1].at = new Date(Date.now() - 61000).toISOString();
  await env.call('own1', 'PUT', PUTQ, { id }, { note: 'x' });
  t('1.6 超過 60 秒 → 另起一筆', q().timeline.length === 3 && !q().timeline[2].count && /備註/.test(q().timeline[2].comment));
  await env.call('own1', 'PUT', PUTQ, { id }, { note: 'x' });
  t('1.7 無實質變更也留一筆（合併進上一筆；摘要含「無內容變更」）', q().timeline.length === 3 && /無內容變更/.test(q().timeline[2].comment) && q().timeline[2].count === 2, J(q().timeline[2]));
  const a1 = JSON.parse(JSON.stringify(q())); const a2 = JSON.parse(JSON.stringify(q())); a2.timeline = [{ at: 'x', by: 'z', action: 'DRAFT_SAVE', comment: 'c' }];
  t('1.8 timeline 不進 contentHash／itemsSig／costLinesSig', QA.contentHash(a1) === QA.contentHash(a2) && QA.itemsSig(a1) === QA.itemsSig(a2) && QA.costLinesSig(a1) === QA.costLinesSig(a2));

  // ═════════ 2) 顧問流程 ═════════
  env = mkEnv([mkQuote({ products: ['PC'], costFlow: { state: 'na' } })]);
  const Q = () => env.data.quotations[0];
  r = await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { costBy: 'cons1', items: formItems(Q()) });
  let ev = Q().timeline.map((x) => x.action);
  t('2.1 指派顧問 → DRAFT_SAVE＋COST_REQUEST（含顧問名）', r.s === 200 && ev.includes('COST_REQUEST') && /Cons1/.test(Q().timeline.find((x) => x.action === 'COST_REQUEST').comment), J(Q().timeline));
  const body = (extra) => Object.assign({ costLines: [L('consult', 'PM', 2, 7000, { unit: '人天', rel: 'link', forLid: 'i1' }), L('consult', 'SD', 1, 6000, { unit: '人天', rel: 'link', forLid: 'i2' }), { cat: 'other', auto: 'stamp' }], done: false, costModel: 2 }, extra || {});
  r = await env.call('cons1', 'PUT', COSTS, { id: 'Q1' }, body());
  let last = Q().timeline[Q().timeline.length - 1];
  t('2.2 顧問儲存草稿 → COST_SAVE（人員 cons1；列數＋新增筆數）', r.s === 200 && last.action === 'COST_SAVE' && last.by === 'cons1' && /儲存成本明細（3 列）/.test(last.comment) && /新增 3/.test(last.comment), J(last));
  r = await env.call('cons1', 'PUT', COSTS, { id: 'Q1' }, body({ done: true, contingencyPct: 5 }));
  last = Q().timeline[Q().timeline.length - 1];
  t('2.3 顧問完成 → COST_DONE（舊 UI 標籤從沒被寫過，現在真的有資料）', r.s === 200 && last.action === 'COST_DONE' && /完成成本填寫/.test(last.comment) && Q().costFlow.state === 'filled', J(last));
  r = await env.call('cons1', 'PUT', COSTS, { id: 'Q1' }, body({ done: false }));
  last = Q().timeline[Q().timeline.length - 1];
  t('2.4 顧問 done:false 從 filled 退回 → COST_REOPEN', r.s === 200 && last.action === 'COST_REOPEN' && Q().costFlow.state === 'requested', J(last));
  await env.call('cons1', 'PUT', COSTS, { id: 'Q1' }, body({ done: true, contingencyPct: 5 }));
  t('2.5 再完成 → filled', Q().costFlow.state === 'filled');
  r = await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { costModel: 2, items: formItems(Q(), (a) => { a.push({ desc: '新增品項', unit: '式', qty: 1, unitPrice: 10, cost: 0 }); }) });
  ev = Q().timeline.slice(-2).map((x) => x.action);
  t('2.6 業務在顧問完成後改品項 → DRAFT_SAVE＋COST_REOPEN，costFlow 回到 requested', r.s === 200 && ev[1] === 'COST_REOPEN' && Q().costFlow.state === 'requested', J(ev) + r.s + J(Q().costFlow));
  const cv = (await env.call('cons1', 'GET', PUTQ, { id: 'Q1' })).j;
  t('2.7 顧問可讀（GET 200）並拿到 timeline；approval 仍是 null（沒有開放簽核資訊）', cv.timeline && cv.timeline.length === Q().timeline.length && cv.approval === null);
  const ov = (await env.call('own1', 'GET', PUTQ, { id: 'Q1' })).j;
  t('2.8 業務端 timeline 與資料一致、依時間順序', ov.timeline.length === Q().timeline.length && ov.timeline.every((x, i, a) => !i || a[i - 1].at <= x.at));

  // ═════════ 3) 不影響簽核 ═════════
  env = mkEnv([mkQuote()]);
  const S = () => env.data.quotations[0];
  await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { note: 'a' });
  r = await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' }, {});
  t('3.1 送簽 → 200，pending；approval.history 只有 SUBMIT', r.s === 200 && S().approval.state === 'pending' && S().approval.history.length === 1 && S().approval.history[0].action === 'SUBMIT', r.s + ' ' + J(r.j && r.j.message));
  t('3.2 送簽後 timeline 仍在（簽核重建 q.approval 不影響頂層陣列）；hash 與內容相符', S().timeline.length >= 1 && S().approval.hash === QA.contentHash(S()));
  r = await env.call('mgr1', 'POST', '/api/quotations/:id/approve', { id: 'Q1' }, { comment: 'ok', hash: QA.contentHash(S()) });
  t('3.3 主管核准路徑正常', r.s === 200, r.s + " " + J(r.j));
  for (const u of ['gm1', 'ch1', 'sec1']) { if (S().approval.state === 'pending') await env.call(u, 'POST', '/api/quotations/:id/approve', { id: 'Q1' }, { hash: QA.contentHash(S()), comment: 'ok', resolutionDate: '2026-10-10', resolutionNo: 'T-1' }); }
  const histActions = S().approval.history.map((x) => x.action);
  t('3.4 approval.history 完全沒有新動作（只有簽核類）', histActions.every((a) => !EDIT.includes(a)), J(histActions));
  if (S().approval.state === 'approved') {
    const aprH = S().approval.hash;
    r = await env.call('own1', 'PUT', '/api/quotations/:id/item-notes', { id: 'Q1' }, { items: [{ lid: 'i1', note: '備註文字' }] });
    t('3.5 核准後改說明／備註 → 核准仍有效（hash 不變）；timeline 多一筆 DRAFT_SAVE', r.s === 200 && S().approval.state === 'approved' && S().approval.hash === aprH && QA.contentHash(S()) === aprH && S().timeline[S().timeline.length - 1].action === 'DRAFT_SAVE' && r.j.approval.valid === true, r.s + ' ' + J(r.j && r.j.approval && r.j.approval.valid));
    r = await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { items: formItems(S(), (a) => { a[0].qty = 9; }) });
    t('3.6 核准後改數量沒帶 confirmVoid → 409 WILL_VOID（原有行為），hash 沒被動', r.s === 409 && S().approval.state === 'approved' && S().approval.hash === aprH);
    r = await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { items: formItems(S(), (a) => { a[0].qty = 9; }), confirmVoid: true });
    const lastTl = S().timeline[S().timeline.length - 1];
    t('3.7 確認作廢 → approval.history 有 INVALIDATE，timeline 有 DRAFT_SAVE（含「核准已作廢」），簽核事件沒被擠掉', r.s === 200 && S().approval.state === 'none' && S().approval.history.slice(-1)[0].action === 'INVALIDATE' && /核准已作廢/.test(lastTl.comment) && S().approval.history.filter((x) => x.action === 'APPROVE').length >= 1, J(lastTl));
  } else t('3.5-3.7 簽核鏈沒走到核准（測試環境）', false, S().approval.state);
  env = mkEnv([mkQuote()]);
  await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { note: 'a' });
  r = await env.call('own1', 'DELETE', PUTQ, { id: 'Q1' });
  t('3.8 只有編輯事件的草稿仍可刪除（不被 HAS_APPROVAL 擋）', r.s === 200 && env.data.quotations.length === 0, r.s);

  // ═════════ 4) 上限 200 與舊單 ═════════
  env = mkEnv([mkQuote()]);
  const T = () => env.data.quotations[0];
  T().approval = { state: 'none', history: [{ at: '2026-10-01T00:00:00.000Z', by: 'own1', action: 'SUBMIT', comment: '', tier: '' }, { at: '2026-10-01T00:00:01.000Z', by: 'mgr1', action: 'RETURN', comment: 'r', tier: 'mgr1' }] };
  T().timeline = Array.from({ length: 205 }, (_, i) => ({ at: new Date(Date.UTC(2026, 9, 1, 0, 0, 0) + i * 120000).toISOString(), by: 'own1', action: 'DRAFT_SAVE', comment: 'e' + i }));
  await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { note: 'trim' });
  // 稽核第二波（D）刻意更新：超過上限時不再無聲截斷——最前面放一筆「已省略 N 筆」標記（占上限一格），其餘保留最新 199 筆；完整紀錄在稽核事件表。
  t('4.1 超過 200 筆：標記＋最新 199 筆共 200（含剛寫的），標記寫明省略 7 筆，approval.history 兩筆簽核事件完全不動', T().timeline.length === 200 && T().timeline[0].action === 'TIMELINE_TRIMMED' && /已省略 7 筆較早事件，完整紀錄見稽核日誌/.test(T().timeline[0].comment) && T().timeline[1].comment === 'e7' && /備註/.test(T().timeline[199].comment) && T().approval.history.length === 2, T().timeline.length + ' ' + T().timeline[0].comment);
  const plain = JSON.parse(JSON.stringify(mkQuote()));
  env = mkEnv([plain]);
  const g = (await env.call('own1', 'GET', PUTQ, { id: 'Q1' })).j;
  t('4.2 舊單（沒有 timeline）序列化沒有 timeline／timelineCount 鍵', !('timeline' in g) && !('timelineCount' in g) && g.id === 'Q1');
  const lst = (await env.call('own1', 'GET', '/api/quotations', {})).j;
  t('4.3 舊單在清單也沒有新鍵', !('timeline' in lst[0]) && !('timelineCount' in lst[0]));
  env = mkEnv([mkQuote({ timeline: [{ at: '2026-10-08T00:00:00.000Z', by: 'own1', action: 'DRAFT_SAVE', comment: 'x' }] })]);
  const l2 = (await env.call('own1', 'GET', '/api/quotations', {})).j;
  t('4.4 清單只給 timelineCount、不帶整串 timeline', l2[0].timelineCount === 1 && !('timeline' in l2[0]));
  env = mkEnv([JSON.parse(JSON.stringify(plain))]);
  r = await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { note: 'first' });
  t('4.5 舊單第一次存檔：timeline 從現在開始（1 筆）', r.s === 200 && env.data.quotations[0].timeline.length === 1);

  // ═════════ 5) 前端靜態 ═════════
  const qa = fs.readFileSync(path.join(ROOT, '_client/quote-approval.js'), 'utf8');
  const qj = fs.readFileSync(path.join(ROOT, '_client/quote.js'), 'utf8');
  for (const a of ['DRAFT_CREATE', 'DRAFT_SAVE', 'COST_SAVE', 'COST_REOPEN', 'COST_REQUEST', 'COST_DONE']) t('5.1 歷程標籤有 ' + a, new RegExp(a + ": '").test(qa));
  t('5.2 歷程表合併 q.timeline 並依時間排序；舊單沒有 timeline 直接用 approval.history', /function mergedHistory/.test(qa) && /if \(!t\.length\) return h;/.test(qa));
  t('5.3 顧問成本對話框也顯示歷程；草稿有 timelineCount 時有「簽核進度」按鈕', /\$\{buildHistorySection\(q\)\}`;/.test(qa) && /q\.timelineCount > 0/.test(qj));
  t('5.4 時間格式 Asia/Taipei', /timeZone: 'Asia\/Taipei'/.test(qa));
}
run().then(() => {
  let pass = 0, fail = 0;
  res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (x && !ok ? '  <- ' + x : '')); ok ? pass++ : fail++; });
  console.log('\n編輯事件時間軸檢查：PASS ' + pass + ' / FAIL ' + fail);
  process.exit(fail ? 1 : 0);
}).catch((e) => { console.error('例外', e.stack); process.exit(2); });
