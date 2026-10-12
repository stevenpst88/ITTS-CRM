#!/usr/bin/env node
/**
 * 報價單 ↔ 商機連結（q.oppId）檢查。用法：node scripts/check-quote-opp-link.js（不需要伺服器；假 app＋記憶體資料庫直接呼叫路由 handler；不碰 data.json／auth.json／audit.log.json）
 * 動 lib/quoteOppLink.js、lib/quoteRoutes.js 的 POST／PUT 報價單／serialize／GET /api/opportunities/:id/quotes，或 server.js 的商機清單／改公司／刪除之後必跑。
 * （server.js 的商機刪除／改公司／清單 linkedQuoteCount 需要真實 HTTP，由 e2e 驗；這裡只驗純函式 countByOpp／isLiveQuote 與其靜態接線。）
 *   1) 建單必須有 oppId：缺、空白、不存在、已刪除、別人的、客戶池的、admin 連別人的、Won 或沒階段 → 對應錯誤碼；A／B／C／D → 201
 *   2) 一筆商機可連多張；company／聯絡人以商機為準（公司不符 → OPP_COMPANY_MISMATCH；商機有聯絡人 → 伺服器採用商機的）；oppSnapshot 內容
 *   3) PUT 規則：草稿／被駁回可換商機；簽核中鎖定；已核准不可換；不可取消連結；公司不可和商機分歧；不帶 oppId 維持原樣
 *   4) 不影響簽核：contentHash／itemsSig／costLinesSig 與 oppId 無關；核准後帶同一個 oppId 存檔不作廢
 *   5) 舊單（沒有 oppId）：序列化沒有任何 opp 鍵、PUT／送簽沒有擋關；草稿可補連，已核准不可
 *   6) 序列化 oppLive：商機轉 Won／被刪除後的樣子；清單 API 也有
 *   7) 時間軸：DRAFT_CREATE／DRAFT_SAVE 摘要提到連結商機；無變更時不提
 *   8) GET /api/opportunities/:id/quotes：只列看得到的報價單；其他角色不能建單（NO_CREATE_ROLE）
 * 測試資料只用通用字串（公開 repo：不放客戶名、人名、真實費率）。
 */
'use strict';
const path = require('path');
const fs = require('fs');
const ROOT = path.join(__dirname, '..');
const QA = require(path.join(ROOT, 'lib/quoteApproval.js'));
const OL = require(path.join(ROOT, 'lib/quoteOppLink.js'));
const registerQuoteRoutes = require(path.join(ROOT, 'lib/quoteRoutes.js'));
const PNL = require(path.join(ROOT, 'lib/quotePnlExcel.js'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);

const USERS = [
  { username: 'admin1', role: 'admin', active: true, displayName: 'Admin' },
  { username: 'own1', role: 'user', active: true, displayName: 'Owner1', supervisor: 'mgr1', bu: ['ITS'] },
  { username: 'own2', role: 'user', active: true, displayName: 'Owner2', supervisor: 'mgr1', bu: ['ITS'] },
  { username: 'mgr1', role: 'manager1', active: true, displayName: 'Mgr1', bu: ['ITS'] },
  { username: 'gm1', role: 'executive', active: true, displayName: 'Gm' },
  { username: 'ch1', role: 'executive', active: true, displayName: 'Chair' },
  { username: 'cons1', role: 'user', active: true, displayName: 'Cons1', bu: ['ITS'] },
  { username: 'sec1', role: 'secretary', active: true, displayName: 'Sec', bu: ['ITS'] },
];
const CLASSES = { PC: { cls: 'consult', costBySales: false }, PS: { cls: 'software', costBySales: true } };
const opp = (id, over) => Object.assign({ id, owner: 'own1', stage: 'C', company: 'TestCo', product: 'Prod', amount: '800', bu: 'ITS', contactId: '', contactName: '' }, over || {});
function mkEnv(quotations, opps) {
  const routes = {}; const app = {};
  ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
  const env = { data: { opportunities: opps || [], lostOpportunities: [], quotations, quoteApproval: { roster: { gm: ['gm1'], chairman: ['ch1'], boardProxy: [], costProviders: ['cons1'], sealManagers: [] }, productClasses: CLASSES } }, logs: [], notes: [], saves: 0, seq: 0 };
  env.auth = { users: JSON.parse(JSON.stringify(USERS)) };
  const deps = {
    db: { load: () => env.data, save: () => { env.saves++; }, flush: async () => {} }, loadAuth: () => env.auth, saveAuth: () => {},
    requireAuth: (req, rs, next) => next(), requireAdmin: (req, rs, next) => next(),
    writeLog: (...a) => env.logs.push(a), pushNotification: (...a) => env.notes.push(a),
    getViewableOwners: (req) => { const rl = req.session.user.role; if (rl === 'admin' || rl === 'executive' || rl === 'manager1' || rl === 'secretary') return ['own1', 'own2']; return [req.session.user.username]; },
    sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-' + (++env.seq), taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}),
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
const formItems = (q) => q.items.map((x) => ({ lid: x.lid, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: x.unitPrice, cost: x.cost }));
const PUTQ = '/api/quotations/:id', POSTQ = '/api/quotations';
const body = (oppId, extra) => Object.assign({ oppId, company: 'TestCo', projectName: 'P', products: ['PS'], validUntil: '2026-10-30', items: [{ desc: '品項A', unit: '式', qty: 1, unitPrice: 1000, cost: 100 }] }, extra || {});
const OPPS = () => [
  opp('OA', { stage: 'A' }), opp('OB', { stage: 'B' }), opp('OC', { stage: 'C' }), opp('OD', { stage: 'D' }), opp('OW', { stage: 'Won' }), opp('ONS', { stage: '' }),
  opp('OX', { owner: 'own2' }), opp('OP', { owner: '_pool' }), opp('OA2', { stage: 'C', owner: 'admin1' }),
  opp('OOther', { company: 'OtherCo' }), opp('OSame', { company: 'TestCo', product: 'Prod2' }),
  opp('OCon', { contactId: 'ct-1', contactName: 'Contact One' }), opp('ONoCo', { company: '' }),
];

async function run() {
  // ═════════ 1) 建單必須有 oppId ═════════
  let env = mkEnv([], OPPS());
  const post = (user, oppId, extra) => env.call(user, 'POST', POSTQ, {}, body(oppId, extra));
  let r = await env.call('own1', 'POST', POSTQ, {}, (() => { const b = body('OC'); delete b.oppId; return b; })());
  t('1.1 沒有 oppId → 400 OPP_REQUIRED，沒有存檔', r.s === 400 && r.j.code === 'OPP_REQUIRED' && env.data.quotations.length === 0, J(r.j));
  for (const bad of ['', '   ', null]) { r = await post('own1', bad); t('1.2 oppId=' + J(bad) + ' → 400 OPP_REQUIRED', r.s === 400 && r.j.code === 'OPP_REQUIRED', r.s + J(r.j)); }
  r = await post('own1', 'no-such');
  t('1.3 不存在的商機 → 400 OPP_NOT_FOUND', r.s === 400 && r.j.code === 'OPP_NOT_FOUND', J(r.j));
  env.data.lostOpportunities.push(opp('OLost'));
  r = await post('own1', 'OLost');
  t('1.4 已被刪除（在 lostOpportunities）的商機 → 400 OPP_NOT_FOUND', r.s === 400 && r.j.code === 'OPP_NOT_FOUND');
  r = await post('own1', { x: 1 });
  t('1.5 oppId 不是字串（物件）→ 4xx，不會丟例外也不會存檔', r.s >= 400 && r.s < 500 && env.data.quotations.length === 0, r.s + J(r.j));
  r = await post('own1', 'OX');
  t('1.6 別人的商機 → 403 OPP_NOT_YOURS', r.s === 403 && r.j.code === 'OPP_NOT_YOURS', J(r.j));
  r = await post('own1', 'OP');
  t('1.7 客戶池（_pool）的商機 → 403 OPP_NOT_YOURS', r.s === 403 && r.j.code === 'OPP_NOT_YOURS');
  r = await post('admin1', 'OC');
  t('1.8 admin 也不能連別人的商機（沒有 admin 例外）→ 403 OPP_NOT_YOURS', r.s === 403 && r.j.code === 'OPP_NOT_YOURS', J(r.j));
  r = await post('admin1', 'OA2');
  t('1.9 admin 連自己的商機 → 201', r.s === 201 && r.j.oppId === 'OA2', r.s + J(r.j));
  r = await post('own1', 'OW');
  t('1.10 Won 階段 → 400 OPP_STAGE_NOT_ALLOWED', r.s === 400 && r.j.code === 'OPP_STAGE_NOT_ALLOWED', J(r.j));
  r = await post('own1', 'ONS');
  t('1.11 沒有階段 → 400 OPP_STAGE_NOT_ALLOWED', r.s === 400 && r.j.code === 'OPP_STAGE_NOT_ALLOWED');
  for (const id of ['OA', 'OB', 'OC', 'OD']) { r = await post('own1', id); t('1.12 階段 ' + id.slice(1) + ' → 201，回應帶 oppId', r.s === 201 && r.j.oppId === id, r.s + J(r.j)); }
  r = await env.call('own1', 'POST', POSTQ, {}, body('OC', { oppId: 'OC' }));
  const before = env.data.quotations.length;
  for (const u of ['sec1', 'mgr1', 'gm1']) { r = await post(u, 'OC'); t('1.13 ' + u + '（不建單的角色）→ 403 NO_CREATE_ROLE（角色檢查先於商機）', r.s === 403 && r.j.code === 'NO_CREATE_ROLE', J(r.j)); }
  t('1.14 被擋的請求都沒有存檔', env.data.quotations.length === before);

  // ═════════ 2) 多張、公司／聯絡人、snapshot ═════════
  env = mkEnv([], OPPS());
  const a1 = await env.call('own1', 'POST', POSTQ, {}, body('OC'));
  const a2 = await env.call('own1', 'POST', POSTQ, {}, body('OC'));
  t('2.1 同一筆商機可連多張報價單（兩張都 201、不同 id）', a1.s === 201 && a2.s === 201 && a1.j.id !== a2.j.id && env.data.quotations.every((q) => q.oppId === 'OC'));
  r = await post('own1', 'OC', { company: 'WrongCo' });
  t('2.2 公司和商機不同 → 400 OPP_COMPANY_MISMATCH，沒存檔', r.s === 400 && r.j.code === 'OPP_COMPANY_MISMATCH' && env.data.quotations.length === 2, J(r.j));
  r = await post('own1', 'OC', { company: '  TestCo  ' });
  t('2.3 公司前後空白忽略 → 201', r.s === 201);
  const nb = body('OC'); delete nb.company;
  r = await env.call('own1', 'POST', POSTQ, {}, nb);
  t('2.4 沒送 company → 伺服器用商機的公司', r.s === 201 && r.j.company === 'TestCo');
  r = await post('own1', 'ONoCo', { company: 'FreeCo' });
  t('2.5 商機沒有公司 → 採用請求的公司', r.s === 201 && r.j.company === 'FreeCo', J(r.j));
  r = await post('own1', 'OCon', { company: 'TestCo', contactId: 'zzz', contactName: 'Someone Else' });
  t('2.6 商機有聯絡人 → 伺服器採用商機的（請求的聯絡人被忽略）', r.s === 201 && r.j.contactId === 'ct-1' && r.j.contactName === 'Contact One', J([r.j.contactId, r.j.contactName]));
  r = await post('own1', 'OC', { contactId: 'req-1', contactName: 'Req Name' });
  t('2.7 商機沒有聯絡人 → 採用請求帶的聯絡人', r.s === 201 && r.j.contactId === 'req-1' && r.j.contactName === 'Req Name');
  const sn = env.data.quotations[0].oppSnapshot;
  t('2.8 oppSnapshot：product／amountK（數字，NT$K）／bu／stage／company／linkedAt', sn && sn.product === 'Prod' && sn.amountK === 800 && sn.bu === 'ITS' && sn.stage === 'C' && sn.company === 'TestCo' && !isNaN(Date.parse(sn.linkedAt)), J(sn));
  t('2.9 回應帶 oppLive {exists:true, stage, product}', J(a1.j.oppLive) === J({ exists: true, stage: 'C', product: 'Prod' }), J(a1.j.oppLive));
  env.data.opportunities.find((o) => o.id === 'OC').amount = '';
  r = await post('own1', 'OC');
  t('2.10 商機金額空白 → amountK:null（不是 0、不丟例外）', r.s === 201 && env.data.quotations[env.data.quotations.length - 1].oppSnapshot.amountK === null);

  // ═════════ 3) PUT 規則 ═════════
  env = mkEnv([], OPPS());
  let c = await env.call('own1', 'POST', POSTQ, {}, body('OC'));
  const id = c.j.id; const Q = () => env.data.quotations.find((x) => x.id === id);
  r = await env.call('own1', 'PUT', PUTQ, { id }, { projectName: 'NewName' });
  t('3.1 PUT 不帶 oppId → 200，連結不變', r.s === 200 && Q().oppId === 'OC' && r.j.oppId === 'OC');
  r = await env.call('own1', 'PUT', PUTQ, { id }, { oppId: 'OC' });
  t('3.2 PUT 帶同一個 oppId → 200 無變更', r.s === 200 && Q().oppId === 'OC');
  const snapAt = Q().oppSnapshot.linkedAt;
  t('3.2b 同一商機不重拍 snapshot（linkedAt 不變）', Q().oppSnapshot.linkedAt === snapAt);
  r = await env.call('own1', 'PUT', PUTQ, { id }, { oppId: 'OSame' });
  t('3.3 草稿換成另一筆自己的商機（同公司）→ 200；oppId／snapshot 更新', r.s === 200 && Q().oppId === 'OSame' && Q().oppSnapshot.product === 'Prod2' && r.j.oppSnapshot.product === 'Prod2', J(r.j.oppSnapshot));
  r = await env.call('own1', 'PUT', PUTQ, { id }, { oppId: 'OOther' });
  t('3.4 換成不同公司的商機（沒送 company）→ 200，公司跟著商機換', r.s === 200 && Q().company === 'OtherCo' && Q().oppId === 'OOther', r.s + J(r.j.code));
  r = await env.call('own1', 'PUT', PUTQ, { id }, { oppId: 'OC', company: 'OtherCo' });
  t('3.5 換商機但同時送另一家公司 → 400 OPP_COMPANY_MISMATCH，原連結不變', r.s === 400 && r.j.code === 'OPP_COMPANY_MISMATCH' && Q().oppId === 'OOther', J(r.j));
  r = await env.call('own1', 'PUT', PUTQ, { id }, { oppId: 'OC', company: 'TestCo' });
  t('3.5b 換商機並送對應公司 → 200', r.s === 200 && Q().oppId === 'OC' && Q().company === 'TestCo');
  for (const [oid, code, st] of [['OX', 'OPP_NOT_YOURS', 403], ['OW', 'OPP_STAGE_NOT_ALLOWED', 400], ['nope', 'OPP_NOT_FOUND', 400], ['OP', 'OPP_NOT_YOURS', 403]]) {
    r = await env.call('own1', 'PUT', PUTQ, { id }, { oppId: oid });
    t('3.6 換成 ' + oid + ' → ' + st + ' ' + code + '，連結不變', r.s === st && r.j.code === code && Q().oppId === 'OC', r.s + J(r.j));
  }
  r = await env.call('own1', 'PUT', PUTQ, { id }, { oppId: '' });
  t('3.7 取消連結（空字串）→ 400 OPP_REQUIRED，連結不變', r.s === 400 && r.j.code === 'OPP_REQUIRED' && Q().oppId === 'OC');
  r = await env.call('own1', 'PUT', PUTQ, { id }, { oppId: null, projectName: 'p2' });
  t('3.7b oppId:null 視為沒帶（不取消）', r.s === 200 && Q().oppId === 'OC');
  r = await env.call('own1', 'PUT', PUTQ, { id }, { company: 'SomethingElse' });
  t('3.8 已連結的單把公司改成別家 → 400 OPP_COMPANY_MISMATCH', r.s === 400 && r.j.code === 'OPP_COMPANY_MISMATCH' && Q().company === 'TestCo', J(r.j));
  r = await env.call('own1', 'PUT', PUTQ, { id }, { company: 'TestCo' });
  t('3.8b 公司送一樣的 → 200', r.s === 200);
  // 商機被刪除後：公司仍以 snapshot 為準
  env.data.opportunities = env.data.opportunities.filter((o) => o.id !== 'OC');
  r = await env.call('own1', 'PUT', PUTQ, { id }, { company: 'SomethingElse' });
  t('3.9 商機已被刪除：公司仍以 snapshot 的公司為準 → 400', r.s === 400 && r.j.code === 'OPP_COMPANY_MISMATCH');
  r = await env.call('own1', 'PUT', PUTQ, { id }, { projectName: 'still-ok' });
  t('3.9b 商機已被刪除：其他欄位照常存檔（oppLive.exists=false，snapshot 保留）', r.s === 200 && r.j.oppLive.exists === false && r.j.oppSnapshot.product === 'Prod');
  env.data.opportunities.push(opp('OC'));

  // 簽核中：鎖定
  r = await env.call('own1', 'POST', '/api/quotations/:id/submit', { id }, {});
  t('3.10 送簽（新單、已連商機）→ 200 pending', r.s === 200 && Q().approval.state === 'pending', r.s + J(r.j && r.j.message));
  r = await env.call('own1', 'PUT', PUTQ, { id }, { oppId: 'OSame' });
  t('3.11 簽核中換商機 → 409 LOCKED_PENDING，連結不變', r.s === 409 && r.j.code === 'LOCKED_PENDING' && Q().oppId === 'OC');
  // 被駁回（returned）→ 可換
  const hashPending = Q().approval.hash;
  r = await env.call('mgr1', 'POST', '/api/quotations/:id/return', { id }, { comment: 'no', hash: QA.contentHash(Q()) });
  t('3.12 主管退回 → returned', Q().approval.state === 'returned', r.s + J(r.j));
  r = await env.call('own1', 'PUT', PUTQ, { id }, { oppId: 'OSame' });
  t('3.13 被退回後可換商機（同公司）→ 200；核准雜湊不受影響', r.s === 200 && Q().oppId === 'OSame' && QA.contentHash(Q()) === hashPending, r.s + J(r.j.code));

  // 已核准
  env = mkEnv([], OPPS());
  c = await env.call('own1', 'POST', POSTQ, {}, body('OC', { items: [{ desc: '導入顧問', unit: '式', qty: 1, unitPrice: 600000, cost: 100000 }, { desc: '教育訓練', unit: '式', qty: 1, unitPrice: 100000, cost: 20000 }] }));
  const id2 = c.j.id; const Q2 = () => env.data.quotations.find((x) => x.id === id2);
  await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: id2 }, {});
  for (const u of ['mgr1', 'gm1', 'ch1', 'sec1']) { if (Q2().approval.state === 'pending') await env.call(u, 'POST', '/api/quotations/:id/approve', { id: id2 }, { hash: QA.contentHash(Q2()), comment: 'ok', resolutionDate: '2026-10-10', resolutionNo: 'T-1' }); }
  t('3.14 前置：走完簽核 → approved', Q2().approval.state === 'approved', Q2().approval.state);
  const aprHash = Q2().approval.hash;
  r = await env.call('own1', 'PUT', PUTQ, { id: id2 }, { oppId: 'OSame' });
  t('3.15 已核准換商機 → 409 OPP_LOCKED，連結與核准都不變', r.s === 409 && r.j.code === 'OPP_LOCKED' && Q2().oppId === 'OC' && Q2().approval.state === 'approved');
  r = await env.call('own1', 'PUT', PUTQ, { id: id2 }, { oppId: 'OC', projectName: Q2().projectName });
  t('3.16 已核准帶同一個 oppId 存檔 → 200，核准仍有效（沒有 WILL_VOID）', r.s === 200 && Q2().approval.state === 'approved' && Q2().approval.hash === aprHash && QA.contentHash(Q2()) === aprHash, r.s + J(r.j.code));
  r = await env.call('own1', 'PUT', PUTQ, { id: id2 }, { oppId: '' });
  t('3.17 已核准取消連結 → 400', r.s === 400 && Q2().oppId === 'OC');

  // ═════════ 4) 雜湊 ═════════
  const lq = mkQuote(); const lq2 = JSON.parse(JSON.stringify(lq)); lq2.oppId = 'OC'; lq2.oppSnapshot = { product: 'x', amountK: 1, bu: 'ITS', stage: 'C', company: 'TestCo', linkedAt: 'z' };
  t('4.1 contentHash／itemsSig／costLinesSig 與 oppId、oppSnapshot 無關', QA.contentHash(lq) === QA.contentHash(lq2) && QA.itemsSig(lq) === QA.itemsSig(lq2) && QA.costLinesSig(lq) === QA.costLinesSig(lq2) && QA.structureSig(lq) === QA.structureSig(lq2));

  // ═════════ 5) 舊單 ═════════
  env = mkEnv([mkQuote()], OPPS());
  let g = await env.call('own1', 'GET', PUTQ, { id: 'Q1' });
  t('5.1 舊單 GET：沒有 oppId／oppSnapshot／oppLive 鍵', g.s === 200 && !('oppId' in g.j) && !('oppSnapshot' in g.j) && !('oppLive' in g.j), J(Object.keys(g.j).filter((k) => /opp/i.test(k))));
  const lst = (await env.call('own1', 'GET', '/api/quotations', {})).j;
  t('5.2 舊單清單項目也沒有 opp 鍵', lst.length === 1 && !('oppId' in lst[0]) && !('oppLive' in lst[0]));
  r = await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { projectName: 'edited', items: formItems(env.data.quotations[0]) });
  t('5.3 舊單 PUT（沒帶 oppId）→ 200，仍沒有 oppId 欄位', r.s === 200 && !('oppId' in env.data.quotations[0]));
  r = await env.call('own1', 'POST', '/api/quotations/:id/submit', { id: 'Q1' }, {});
  t('5.4 舊單送簽沒有商機擋關 → 200 pending', r.s === 200 && env.data.quotations[0].approval.state === 'pending', r.s + J(r.j && r.j.message));
  for (const u of ['mgr1', 'gm1', 'ch1', 'sec1']) { if (env.data.quotations[0].approval.state === 'pending') await env.call(u, 'POST', '/api/quotations/:id/approve', { id: 'Q1' }, { hash: QA.contentHash(env.data.quotations[0]), comment: 'ok', resolutionDate: '2026-10-10', resolutionNo: 'T-1' }); }
  t('5.5 舊單可走完簽核 → approved', env.data.quotations[0].approval.state === 'approved');
  r = await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { oppId: 'OC' });
  t('5.6 已核准的舊單補連商機 → 409 OPP_LOCKED（不能動）', r.s === 409 && r.j.code === 'OPP_LOCKED' && !('oppId' in env.data.quotations[0]));
  env = mkEnv([mkQuote()], OPPS());
  r = await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { oppId: 'OSame' });
  t('5.7 舊單草稿可補連自己的商機 → 200，產生 oppId／snapshot', r.s === 200 && env.data.quotations[0].oppId === 'OSame' && env.data.quotations[0].oppSnapshot.product === 'Prod2', r.s + J(r.j.code));
  env = mkEnv([mkQuote()], OPPS());
  r = await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { oppId: 'OX' });
  t('5.8 舊單草稿補連別人的商機 → 403', r.s === 403 && r.j.code === 'OPP_NOT_YOURS');
  r = await env.call('own1', 'PUT', PUTQ, { id: 'Q1' }, { oppId: '' });
  t('5.9 舊單帶空 oppId → 視為沒連結、不報錯（不強迫補連）', r.s === 200 && !('oppId' in env.data.quotations[0]));

  // ═════════ 6) oppLive ═════════
  env = mkEnv([], OPPS());
  c = await env.call('own1', 'POST', POSTQ, {}, body('OD'));
  const id3 = c.j.id;
  env.data.opportunities.find((o) => o.id === 'OD').stage = 'Won';
  g = await env.call('own1', 'GET', PUTQ, { id: id3 });
  t('6.1 商機轉 Won 後：oppLive.stage=Won、exists=true；snapshot 仍是 D', g.j.oppLive.exists === true && g.j.oppLive.stage === 'Won' && g.j.oppSnapshot.stage === 'D', J(g.j.oppLive));
  let l = (await env.call('own1', 'GET', '/api/quotations', {})).j;
  t('6.2 清單 API 同樣有 oppId／oppSnapshot／oppLive', l[0].oppId === 'OD' && l[0].oppLive.stage === 'Won' && l[0].oppSnapshot.stage === 'D');
  const gone = env.data.opportunities.splice(env.data.opportunities.findIndex((o) => o.id === 'OD'), 1)[0]; env.data.lostOpportunities.push(gone);
  l = (await env.call('own1', 'GET', '/api/quotations', {})).j;
  t('6.3 商機被刪除後：oppLive.exists=false，snapshot 保留', l[0].oppLive.exists === false && l[0].oppSnapshot.product === 'Prod');
  r = await env.call('own1', 'PUT', PUTQ, { id: id3 }, { projectName: 'x' });
  t('6.4 商機已刪除／Won 的既有報價單仍可存檔（連結只管建單與換商機）', r.s === 200);

  // ═════════ 7) 時間軸 ═════════
  env = mkEnv([], OPPS());
  c = await env.call('own1', 'POST', POSTQ, {}, body('OC'));
  const Q7 = () => env.data.quotations[0];
  t('7.1 DRAFT_CREATE 摘要有「連結商機：TestCo Prod」，且不含金額', /連結商機：TestCo Prod/.test(Q7().timeline[0].comment) && !/800/.test(Q7().timeline[0].comment), Q7().timeline[0].comment);
  await env.call('own1', 'PUT', PUTQ, { id: Q7().id }, { projectName: 'again' });
  t('7.2 沒換商機的存檔摘要不提連結商機', !/連結商機/.test(Q7().timeline[Q7().timeline.length - 1].comment), Q7().timeline[Q7().timeline.length - 1].comment);
  Q7().timeline[Q7().timeline.length - 1].at = new Date(Date.now() - 61000).toISOString();
  await env.call('own1', 'PUT', PUTQ, { id: Q7().id }, { oppId: 'OSame' });
  t('7.3 換商機的存檔摘要有「連結商機：TestCo Prod2」', /連結商機：TestCo Prod2/.test(Q7().timeline[Q7().timeline.length - 1].comment), Q7().timeline[Q7().timeline.length - 1].comment);
  t('7.4 時間軸不進 approval.history', (Q7().approval || null) === null);

  // ═════════ 8) GET /api/opportunities/:id/quotes ═════════
  env = mkEnv([], OPPS());
  await env.call('own1', 'POST', POSTQ, {}, body('OC')); await env.call('own1', 'POST', POSTQ, {}, body('OC'));
  const gq = (u, oid) => env.call(u, 'GET', '/api/opportunities/:id/quotes', { id: oid });
  r = await gq('own1', 'OC');
  t('8.1 擁有者：看到 2 張（摘要含 quoteNo／approvalState，沒有金額）', r.s === 200 && r.j.total === 2 && r.j.quotes.length === 2 && r.j.quotes[0].approvalState === 'none' && !('items' in r.j.quotes[0]) && r.j.exists === true, J(r.j));
  r = await gq('own2', 'OC');
  t('8.2 別的業務：看不到（quotes 空），只知道總數', r.s === 200 && r.j.quotes.length === 0 && r.j.total === 2);
  r = await gq('mgr1', 'OC');
  t('8.3 主管（看得到下屬的單）：看到 2 張', r.s === 200 && r.j.quotes.length === 2);
  r = await gq('own1', 'nope');
  t('8.4 不存在的商機：exists=false、空清單', r.s === 200 && r.j.exists === false && r.j.total === 0);

  // ═════════ 9) 純函式 ═════════
  const d = { quotations: [{ oppId: 'A', approval: null }, { oppId: 'A', approval: { state: 'returned' } }, { oppId: 'A', approval: { state: 'approved' } }, { oppId: 'B', approval: { state: 'returned' } }, { id: 'legacy' }] };
  const m = OL.countByOpp(d);
  t('9.1 countByOpp：total 含全部、live 不含被退回；舊單不計', m.get('A').total === 3 && m.get('A').live === 2 && m.get('B').total === 1 && m.get('B').live === 0 && !m.has(undefined) && m.size === 2);
  t('9.2 linkedQuotes／isLiveQuote', OL.linkedQuotes(d, 'A').length === 3 && OL.linkedQuotes(d, '').length === 0 && OL.linkedQuotes(d, 'A').filter(OL.isLiveQuote).length === 2);

  // ═════════ 9b) config.me.canCreate（商機卡「建立報價單」鈕只給能建單的角色）═════════
  env = mkEnv([], OPPS());
  const cfgOf = async (u) => (await env.call(u, 'GET', '/api/quote-approval/config', {})).j.me.canCreate;
  t('9b.1 業務、admin 能建單；主管／總經理／董事長／秘書不能（與 POST 的 NO_CREATE_ROLE 同一條件）', (await cfgOf('own1')) === true && (await cfgOf('admin1')) === true && (await cfgOf('mgr1')) === false && (await cfgOf('sec1')) === false && (await cfgOf('gm1')) === false);

  // ═════════ 10) 靜態接線 ═════════
  const srv = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  t('10.1 server.js：商機清單帶 linkedQuoteCount、PUT 商機改公司擋 OPP_COMPANY_LOCKED、DELETE 商機擋 OPP_HAS_QUOTES', /linkedQuoteCount: _qc\.get/.test(srv) && /OPP_COMPANY_LOCKED/.test(srv) && /OPP_HAS_QUOTES/.test(srv) && /require\('\.\/lib\/quoteOppLink'\)/.test(srv));
}

run().then(() => {
  let pass = 0, fail = 0;
  res.forEach(([n, ok, ex]) => { if (ok) { pass++; console.log('PASS ' + n); } else { fail++; console.log('FAIL ' + n + (ex ? '  <- ' + ex : '')); } });
  console.log('\n報價單連結商機檢查：PASS ' + pass + ' / FAIL ' + fail);
  process.exit(fail ? 1 : 0);
}).catch((e) => { res.forEach(([n, ok, ex]) => console.log((ok ? 'PASS ' : 'FAIL ') + n + (ok ? '' : '  <- ' + ex))); console.log('例外', e.stack); process.exit(2); });
