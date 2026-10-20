#!/usr/bin/env node
'use strict';
/**
 * 資安審查（2026-10-20）第五節「建議立即修復的低風險項目」彙整檢查 —— 分支 fix/secrev-quick。
 * 不含任何帳密：測試帳號與隨機密碼只在記憶體裡；JWT 金鑰每次執行隨機產生。
 *
 * 涵蓋（每項都有「變異驗證」：把修法改壞，對應檢查必須變紅）：
 *   A（行程內，假 app＋記憶體資料庫，不開埠、不碰任何資料檔）
 *     #1 M2  秘書跨 BU 清單彙總稽核 VIEW_QUOTE_CROSS_BU_LIST（10 分鐘節流）＋ GET 路徑稽核 flush（auditCrossBu、權限類型報告）
 *     #2 L5  名冊／例外驗證不接受 constructor、__proto__；exceptions 元素型別／長度驗證（null → 400，不再 500）
 *     #7 M5  設計器保護底線：取樣被截斷 fail-closed；商品類別鍵取全部（第 7 個鍵不再是盲點）
 *     #9 L7  看不到成本的人：時間軸「成本合計 A → B」與 preview.marginText 遮罩
 *   B（真實 server.js，本機 JSON 模式，埠 SQ_PORT 預設 3180；臨時 _s_* 帳號；絕不連任何資料庫）
 *     #6 M6  登入憑證綁密碼版本（pwa）：改密碼／管理員重設後舊 cookie → 401；舊版 token（無 pwa）放行；SESSION_PWD_BINDING=off 可關
 *     #5 M3  登入鎖定鍵含來源 IP（另一個 IP 仍可登入）；LOGIN_LOCK_SCOPE=name 還原舊行為
 *     #12 H1 建立／升格管理員、重設管理員密碼 → ADMIN_ACCOUNT_CHANGE 稽核＋站內通知其他現任管理員
 *   C（靜態）#11 L12 公開 repo 不含本機路徑／人名；#1 auditCrossBu 內有 flush
 *   另有（各自的既有檢查腳本已擴充，本檔不重複）：#3 L1、#8 L3 → scripts/check-quote-issuance.js（F 區）；
 *     #4 M4 → scripts/check-csp-report.js；#5 單元 → scripts/check-login-guard.js
 *
 *   node scripts/check-secrev-quick.js
 *   SQ_SKIP_MUT=1 node scripts/check-secrev-quick.js     # 略過變異檢查
 *   SQ_ONLY=A,B node scripts/check-secrev-quick.js       # 只跑指定區
 *
 * 守門：工作樹若已有 auth.json／data.json／audit.log.json（可能是真實資料）一律拒絕執行。
 * 禁用埠：3000／3020／3030／3101-3161。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn, execFileSync } = require('child_process');
const nodeHttp = require('http');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.SQ_PORT || '3180', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3161)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const BASE = `http://localhost:${PORT}`;
const ONLY = process.env.SQ_ONLY ? new Set(process.env.SQ_ONLY.split(',').map(s => s.trim().toUpperCase())) : null;
const want = (k) => !ONLY || ONLY.has(k);
const DATA_FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));

let pass = 0, fail = 0; const failures = [];
function mkT(counter) {
  const T = {
    n: 0, f: 0, fl: [],
    ok(c, name) { if (c) { T.n++; } else { T.f++; T.fl.push(name); } },
    eq(a, b, name) { T.ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); },
  };
  return T;
}
const G = mkT();   // 全域（正式檢查）
const ok = (c, n) => { G.ok(c, n); if (!c) console.log('  ✗ FAIL:', n); };
const eq = (a, b, n) => { G.eq(a, b, n); if (JSON.stringify(a) !== JSON.stringify(b)) console.log('  ✗ FAIL:', n, '得到', JSON.stringify(a), '預期', JSON.stringify(b)); };
const sleep = (n) => new Promise(r => setTimeout(r, n));
const tmpFiles = [];

// ═════════════════════════ A：行程內探針（可針對「變異版」模組重跑）═════════════════════════
const NOW = '2026-10-20T00:00:00Z';
const CLASSES = { PC: { cls: 'consult', costBySales: false }, PS: { cls: 'software', costBySales: true } };
function mkEnv(registerQuoteRoutes, opt) {
  const routes = {}; const app = {};
  ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
  const env = { logs: [], flushes: 0, seq: 0 };
  env.data = { quotations: JSON.parse(JSON.stringify(opt.quotations || [])), quoteApproval: { roster: { gm: ['gm1'], chairman: [], boardProxy: [], costProviders: ['cons1'], sealManagers: [] }, productClasses: CLASSES } };
  env.auth = { users: JSON.parse(JSON.stringify(opt.users)) };
  const deps = {
    db: { load: () => env.data, save: () => {}, flush: async () => {} }, loadAuth: () => env.auth, saveAuth: () => {},
    requireAuth: (req, rs, next) => next(), requireAdmin: (req, rs, next) => next(),
    writeLog: (a, op, t, d) => env.logs.push({ a, op, t, d }), pushNotification: () => {},
    flushAuditBeforeEnd: () => { env.flushes++; },
    getViewableOwners: (req) => (opt.scope || {})[req.session.user.username] || [req.session.user.username],
    sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-NEW', taipeiToday: () => '2026-10-20', resolveIssuer: () => ({}),
    buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: async () => Buffer.from(''), QUOTE_TEMPLATE: '',
    uuidv4: () => 'u' + (++env.seq), normalizeBu: (b) => (Array.isArray(b) ? b : b ? [b] : []), getUserFeatures: () => ['quotations'],
  };
  registerQuoteRoutes(app, deps);
  env.call = (user, method, p, params, body) => new Promise((resolve, reject) => {
    const h = [].concat(...routes[method.toUpperCase() + ' ' + p]);
    const role = (env.auth.users.find((u) => u.username === user) || {}).role;
    const req = { method: method.toUpperCase(), ip: '127.0.0.1', headers: {}, get: () => '', session: { user: { username: user, role } }, params: params || {}, query: {}, body: body === undefined ? {} : JSON.parse(JSON.stringify(body)) };
    const rs = { status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j }); return this; }, send() { resolve({ s: this._s || 200, j: null }); return this; }, setHeader() {}, set() {} };
    let i = 0;
    const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
    next();
  });
  return env;
}
const U = (username, role, extra) => Object.assign({ username, role, active: true, displayName: username, bu: [] }, extra || {});
const mkQuote = (id, owner, extra) => Object.assign({
  id, quoteNo: 'QN-' + id, owner, company: 'TestCo', projectName: 'P', quoteDate: '2026-10-20', status: 'draft', createdAt: NOW, updatedAt: NOW, validUntil: '2026-10-30',
  products: ['PS'], approval: null, discountType: 'none', discountValue: 0,
  items: [{ lid: 'i1-' + id, desc: '軟體授權', unit: '套', qty: 1, unitPrice: 100000, cost: 60000 }],
}, extra || {});

async function probeA1(T, mods) {   // #1 M2
  const users = [U('adm', 'admin'), U('sec1', 'secretary', { bu: ['ITS'] }), U('secAll', 'secretary', { bu: ['ITS', 'ERP'] }), U('own1', 'user', { bu: ['ITS'] }), U('x1', 'user', { bu: ['ERP'] })];
  const qs = [mkQuote('Q1', 'own1'), mkQuote('Q2', 'x1'), mkQuote('Q3', 'x1')];
  const env = mkEnv(mods.routes, { users, quotations: qs });
  const lg = (a) => env.logs.filter(l => l.a === a);
  let r = await env.call('sec1', 'GET', '/api/quotations');
  T.eq([r.s, Array.isArray(r.j) ? r.j.length : -1], [200, 3], '[#1] 秘書清單看得到全部 3 張（可視範圍不變）');
  T.eq(lg('VIEW_QUOTE_CROSS_BU_LIST').length, 1, '[#1] 秘書打清單 → 記 1 筆彙總稽核 VIEW_QUOTE_CROSS_BU_LIST');
  T.ok(lg('VIEW_QUOTE_CROSS_BU_LIST')[0] && /跨 BU 張數 2/.test(lg('VIEW_QUOTE_CROSS_BU_LIST')[0].d) && lg('VIEW_QUOTE_CROSS_BU_LIST')[0].op === 'sec1', '[#1] 彙總稽核附跨 BU 張數（2 張 ERP 的單）與操作者');
  T.ok(env.flushes >= 1, '[#1] 清單的彙總稽核有 flush（Vercel GET 路徑回應前落地）');
  T.eq(lg('VIEW_QUOTE_CROSS_BU').length, 0, '[#1] 清單不逐張記 VIEW_QUOTE_CROSS_BU（維持既有規則）');
  const f0 = env.flushes;
  await env.call('sec1', 'GET', '/api/quotations');
  T.eq(lg('VIEW_QUOTE_CROSS_BU_LIST').length, 1, '[#1] 10 分鐘內再打清單不重複記（節流）');
  T.eq(env.flushes, f0, '[#1] 被節流的那次不寫稽核，也不多餘 flush');
  await env.call('secAll', 'GET', '/api/quotations');
  T.eq(lg('VIEW_QUOTE_CROSS_BU_LIST').filter(l => l.op === 'secAll').length, 0, '[#1] 掛 ITS+ERP 的秘書：3 張都在自己 BU 內 → 不算跨 BU → 不記');
  await env.call('adm', 'GET', '/api/quotations'); await env.call('own1', 'GET', '/api/quotations');
  T.eq(lg('VIEW_QUOTE_CROSS_BU_LIST').filter(l => l.op === 'adm' || l.op === 'own1').length, 0, '[#1] 管理員／業務打清單不記');
  const fl = env.flushes;
  await env.call('sec1', 'GET', '/api/quotations/:id', { id: 'Q2' });
  T.eq(lg('VIEW_QUOTE_CROSS_BU').length, 1, '[#1] 秘書開跨 BU 單 → VIEW_QUOTE_CROSS_BU');
  T.ok(env.flushes > fl, '[#1] 單筆 GET 的 VIEW_QUOTE_CROSS_BU 有 flush（auditCrossBu 內）');
  const fm = env.flushes;
  await env.call('adm', 'GET', '/api/admin/quote-mode-report');
  T.eq(lg('VIEW_QUOTE_MODE_REPORT').length, 1, '[#1] 權限類型報告有稽核');
  T.ok(env.flushes > fm, '[#1] VIEW_QUOTE_MODE_REPORT 有 flush');
}
async function probeA2(T, mods) {   // #2 L5
  const users = [U('adm', 'admin'), U('gm1', 'executive'), U('mgr1', 'manager1', { bu: ['ITS'] }), U('mgr2', 'manager1', { bu: ['ITS'] }), U('own1', 'user', { bu: ['ITS'] })];
  const env = mkEnv(mods.routes, { users, quotations: [] });
  const cfgBody = (roster) => ({ roster: Object.assign({ gm: ['gm1'], chairman: [], boardProxy: [], costProviders: [], sealManagers: [] }, roster), productClasses: CLASSES });
  let r = await env.call('adm', 'PUT', '/api/admin/quote-approval/config', null, cfgBody({}));
  T.eq(r.s, 200, '[#2] 基準：合法名冊 → 200（' + JSON.stringify(r.j).slice(0, 120) + '）');
  for (const bad of ['constructor', '__proto__', 'toString', 'hasOwnProperty', 'valueOf']) {
    for (const key of ['gm', 'chairman', 'sealManagers']) {
      r = await env.call('adm', 'PUT', '/api/admin/quote-approval/config', null, cfgBody({ [key]: [bad] }));
      T.ok(r.s === 400 && r.j && r.j.code === 'BAD_ROSTER', `[#2] 名冊 ${key} 含「${bad}」→ 400 BAD_ROSTER（得到 ${r.s} ${r.j && r.j.code}）`);
    }
  }
  const flowPut = (exceptions) => env.call('adm', 'PUT', '/api/admin/quote-approval/flow-config', null, { buSettings: {}, exceptions });
  r = await flowPut([]); T.eq(r.s, 200, '[#2] 基準：空例外 → 200（' + JSON.stringify(r.j).slice(0, 120) + '）');
  r = await flowPut([null]); T.ok(r.s === 400 && r.j && r.j.code === 'BAD_EXCEPTION', `[#2] exceptions:[null] → 400 BAD_EXCEPTION（不再 TypeError 500）（${r.s}）`);
  for (const bad of ['str', 5, [], true]) { r = await flowPut([bad]); T.eq(r.s, 400, '[#2] 例外元素為 ' + JSON.stringify(bad) + ' → 400'); }
  r = await flowPut([{ id: 'x'.repeat(65), username: 'mgr1', action: 'skip_mgr1' }]); T.eq(r.s, 400, '[#2] id 超過 64 字 → 400');
  r = await flowPut([{ id: 5, username: 'mgr1', action: 'skip_mgr1' }]); T.eq(r.s, 400, '[#2] id 非字串 → 400');
  r = await flowPut([{ username: 'mgr1', action: 'skip_mgr1', label: 'y'.repeat(201) }]); T.eq(r.s, 400, '[#2] label 超過 200 字 → 400');
  r = await flowPut([{ username: 'mgr1', action: 'skip_mgr1', label: { a: 1 } }]); T.eq(r.s, 400, '[#2] label 為物件 → 400');
  r = await flowPut([{ username: 'constructor', action: 'skip_mgr1' }]);
  T.ok(r.s === 200 && (env.data.approvalFlowConfig.exceptions || []).length === 0, '[#2] 例外帳號「constructor」不算存在的帳號 → 不會被存進設定（' + r.s + '）');
  r = await flowPut([{ username: '__proto__', action: 'skip_mgr1' }]);
  T.ok(r.s === 200 && (env.data.approvalFlowConfig.exceptions || []).length === 0, '[#2] 例外帳號「__proto__」同樣不存');
  r = await flowPut([{ id: 'ex-1', username: 'mgr1', bu: 'ITS', role: 'manager1', action: 'skip_mgr1', label: '正常' }]);
  T.ok(r.s === 200 && (env.data.approvalFlowConfig.exceptions || []).length === 1 && env.data.approvalFlowConfig.exceptions[0].username === 'mgr1', '[#2] 合法例外照常存入（回歸）');
}
async function probeA9(T, mods) {   // #9 L7
  const users = [U('adm', 'admin'), U('own1', 'user', { bu: ['ITS'] }), U('mgr1', 'manager1', { bu: ['ITS'] }), U('viewer1', 'user', { bu: ['ITS'] })];
  const tl = [{ at: NOW, by: 'own1', action: 'COST_SAVE', comment: '業務儲存成本（1 項）｜成本合計 NT$ 0 → NT$ 600' }, { at: NOW, by: 'own1', action: 'COST_DONE', comment: '顧問完成成本填寫（1 列）（新增 1、刪除 0、修改 0）｜成本合計 NT$ 1,200 → NT$ 3,400｜同步報價品項 1 項' }, { at: NOW, by: 'own1', action: 'DRAFT_SAVE', comment: '修改 1 列（數量）' }];
  const q = mkQuote('Q9', 'own1', { timeline: tl });
  const env = mkEnv(mods.routes, { users, quotations: [q], scope: { viewer1: ['own1'], mgr1: ['own1'], own1: ['own1'] } });
  const v = await env.call('viewer1', 'GET', '/api/quotations/:id', { id: 'Q9' });
  T.ok(v.s === 200 && v.j && v.j.perm && v.j.perm.canSeeCost === false && v.j.perm.canSeePrice === true, '[#9] 前提：檢視者看得到價格、看不到成本（canSeeCost=false）（' + v.s + '）');
  const vs = JSON.stringify(v.j);
  T.ok(!vs.includes('成本合計'), '[#9] 看不到成本的人：整份回應不含「成本合計」');
  T.eq((v.j.timeline || []).length, 3, '[#9] 時間軸筆數不變（只遮成本金額，動作與筆數仍可見）');
  T.ok((v.j.timeline || [])[0] && /業務儲存成本（1 項）/.test(v.j.timeline[0].comment) && (v.j.timeline || [])[1] && /顧問完成成本填寫（1 列）/.test(v.j.timeline[1].comment) && /同步報價品項 1 項/.test(v.j.timeline[1].comment), '[#9] 遮罩後仍保留動作、筆數與其他說明文字');
  T.eq((v.j.timeline || [])[2] && v.j.timeline[2].comment, '修改 1 列（數量）', '[#9] 非成本事件的文字完全不動');
  T.ok(v.j.preview && 'marginText' in v.j.preview && v.j.preview.marginText === null, '[#9] 看不到成本的人：preview.marginText 為 null（鍵仍在）（' + (v.j.preview && v.j.preview.marginText) + '）');
  T.ok(!/NT\$ 3,400|NT\$ 600/.test(vs), '[#9] 回應不含時間軸裡的任何成本金額');
  const m = await env.call('mgr1', 'GET', '/api/quotations/:id', { id: 'Q9' });
  T.ok(m.j.perm.canSeeCost === true && JSON.stringify(m.j).includes('成本合計 NT$ 1,200 → NT$ 3,400'), '[#9] 看得到成本的一級主管：時間軸成本合計照常（行為不變）');
  T.ok(m.j.preview && typeof m.j.preview.marginText === 'string' && m.j.preview.marginText.length > 0, '[#9] 看得到成本的一級主管：preview.marginText 照常（' + (m.j.preview && m.j.preview.marginText) + '）');
  const o = await env.call('own1', 'GET', '/api/quotations/:id', { id: 'Q9' });
  T.ok(o.j.perm.canSeeCost === true && JSON.stringify(o.j).includes('成本合計') && typeof o.j.preview.marginText === 'string', '[#9] 擁有者（不需顧問的單）看自己的成本：不受影響');
  const l = await env.call('viewer1', 'GET', '/api/quotations');
  T.ok(Array.isArray(l.j) && l.j.length === 1 && l.j[0].preview && l.j[0].preview.marginText === null, '[#9] 清單回應同樣遮罩 preview.marginText');
}
function probeA7(T, mods) {   // #7 M5
  const FD = mods.fd, QF = require(path.join(ROOT, 'lib', 'quoteFlow.js')), QA = require(path.join(ROOT, 'lib', 'quoteApproval.js'));
  const base = () => JSON.parse(JSON.stringify(QF.buildLegacyDefault({ buSettings: {}, exceptions: [] })));
  const envOf = (withActive) => { const e = { roster: { gm: ['g'], chairman: ['c'], boardProxy: ['b'] }, signerState: () => 'ok', chain: { owners: 1, noMgr1: 0, noMgr2: 0 } }; if (withActive) e.activeDef = QF.legacyDefault({ buSettings: {}, exceptions: [] }); return e; };
  // 基準：預設流程底線通過、沒有被截斷
  let rc = FD.runChecks(base(), envOf(true));
  T.ok(rc.ok === true && rc.floor.ok === true && rc.truncated === false, '[#7] 基準：預設流程 ok、底線通過、未截斷（' + rc.ok + '/' + rc.floor.ok + '/' + rc.truncated + '）');
  // A. 墊料＋捷徑：取樣被灌爆截斷，過去判「通過」；現在 fail-closed
  { const d = base(); const owners = Array.from({ length: 60 }, (_, i) => 'u' + i);
    d.routes.unshift({ id: 'r_evil', name: '捷徑', when: { cmp: ['owner', 'in', owners] }, path: ['n_mgr1'] });
    d.routes.unshift({ id: 'r_pad', name: '墊', when: { all: [{ cmp: ['amount', '<', '0.01'] }, { cmp: ['marginPct', '>', '1'] }, { cmp: ['marginPct', '>', '2'] }, { cmp: ['ownerBuPrimary', 'in', ['X1', 'X2', 'X3']] }, { cmp: ['classKeys', 'contains', 'consult'] }] }, path: ['n_mgr1', 'n_gm', 'n_chairman'] });
    rc = FD.runChecks(d, envOf(true));
    T.ok(rc.truncated === true && rc.floor.ok === false && rc.ok === false, '[#7] 取樣被截斷的草稿：floor.ok=false、整體 ok=false（不可存檔）（truncated=' + rc.truncated + ' floor=' + rc.floor.ok + ' ok=' + rc.ok + '）');
    const c10 = rc.checks.find(c => c.n === 10);
    T.ok(c10 && c10.status === 'error' && c10.issues.some(i => /截斷/.test(i.text)), '[#7] 第 10 條「保護底線」判 error 並說明是因取樣被截斷');
  }
  // B. 例外名單字面值超過 8 個（過去只取前 8 個，其餘盲點）→ 視為被截去 → fail-closed
  { const d = base(); const pad = Array.from({ length: 9 }, (_, i) => 'p' + i);
    d.modifiers.push({ id: 'm_evil', name: '拿掉董事長', op: 'remove', target: 'n_chairman', when: { resolvedIn: ['n_mgr2', pad] } });
    rc = FD.runChecks(d, envOf(false));
    T.ok(rc.truncated === true && rc.floor.ok === false, '[#7] 例外帳號字面值 9 個（上限 8）→ 取樣不完整，底線不得判通過');
  }
  // C. 第 7 個商品類別鍵：過去只取前 6 個，邪惡路線永遠取不到
  { const d = base(); const ks = QA.CLASS_KEYS.slice();
    const find = (id) => d.routes.find(r => r.id === id);
    const pads = ks.slice(0, 6).map((k, i) => ({ id: 'r_pad' + i, name: '墊' + i, when: { cmp: ['classKeys', 'contains', k] }, path: ['n_mgr1', 'n_gm', 'n_chairman', 'n_board'] }));
    const evil = { id: 'r_evil', name: '捷徑', when: { cmp: ['classKeys', 'contains', ks[6]] }, path: ['n_mgr1'] };
    d.routes = pads.concat([evil], [find('r_board'), { id: 'r_l3', name: '董事長層級', when: { cmp: ['amount', '>', '10000000'] }, path: ['n_mgr1', 'n_gm', 'n_chairman'] }, find('r_default')]);
    rc = FD.runChecks(d, envOf(false));
    T.ok(ks.length >= 7 && rc.truncated === false && rc.floor.ok === false && rc.floor.violations.length > 0 && rc.ok === false, `[#7] 第 7 個商品類別鍵（${ks[6]}）的捷徑路線被抓到（未截斷、有違規）（truncated=${rc.truncated} violations=${rc.floor.violations.length}）`);
  }
}
async function partA(mods, T, only) {
  if (!only || only.includes('1')) await probeA1(T, mods);
  if (!only || only.includes('2')) await probeA2(T, mods);
  if (!only || only.includes('7')) probeA7(T, mods);
  if (!only || only.includes('9')) await probeA9(T, mods);
}

// 變異：複製 lib 檔為 __mut_ 版，改壞一處，再用變異版重跑對應探針（必須有失敗）
function mutateLib(file, edits) {
  let src = fs.readFileSync(path.join(ROOT, 'lib', file), 'utf8');
  for (const [from, to, name] of edits) { if (!src.includes(from)) throw new Error('變異樣式找不到：' + name); src = src.replace(from, () => to); }
  const out = path.join(ROOT, 'lib', file.replace(/\.js$/, '') + '.__mut_sq.js');
  fs.writeFileSync(out, src); tmpFiles.push(out); return out;
}
async function mutA(name, file, edits, only, which) {
  const f = mutateLib(file, edits);
  const T = mkT();
  try {
    const mods = { routes: require(path.join(ROOT, 'lib', 'quoteRoutes.js')), fd: require(path.join(ROOT, 'lib', 'quoteFlowDesigner.js')) };
    delete require.cache[require.resolve(f)];
    if (which === 'routes') mods.routes = require(f); else mods.fd = require(f);
    await partA(mods, T, only);
  } catch (e) { T.f++; T.fl.push('例外：' + (e && e.message)); }
  ok(T.f > 0, `變異「${name}」→ 對應檢查必須失敗，實際失敗 ${T.f} 項` + (T.f ? '' : '（變異未被抓到！）'));
}

// ═════════════════════════ B：真實伺服器 ═════════════════════════
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const jwtLib = require(path.join(ROOT, 'node_modules', 'jsonwebtoken'));
const rnd = (p) => p + crypto.randomBytes(9).toString('hex') + '!Z';
const PW = rnd('Sq'), PW2 = rnd('Sq'), PW3 = rnd('Sq');
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
let child = null, childPid = null, srvSecret = '';
function guardClean() {
  const g = DATA_FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
}
function cleanupFiles() { DATA_FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }
function seed() {
  const users = [
    mkUser('_s_adm1', 'admin'), mkUser('_s_adm2', 'admin'), mkUser('_s_adm3', 'admin', { active: false }),
    mkUser('_s_u1', 'user', { bu: ['ERP'] }), mkUser('_s_u2', 'user', { bu: ['ERP'] }), mkUser('_s_u3', 'user', { bu: ['ERP'] }),
  ];
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify({ contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], groups: [], quotations: [] }, null, 2));
}
async function startServer(serverFile, extraEnv) {
  srvSecret = crypto.randomBytes(24).toString('hex');
  const env = Object.assign({}, process.env, extraEnv || {}, { PORT: String(PORT), JWT_SECRET: srvSecret, SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  delete env.DATABASE_URL; delete env.MAIL_MODE; delete env.AUDIT_EVENTS_TABLE;
  if (!(extraEnv && 'LOGIN_LOCKOUT' in extraEnv)) delete env.LOGIN_LOCKOUT;
  if (!(extraEnv && 'SESSION_PWD_BINDING' in extraEnv)) delete env.SESSION_PWD_BINDING;
  if (!(extraEnv && 'LOGIN_LOCK_SCOPE' in extraEnv)) delete env.LOGIN_LOCK_SCOPE;
  child = spawn(process.execPath, [serverFile || path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  const onData = d => { out += d; if (process.env.SQ_DEBUG) process.stdout.write('[srv] ' + d); };
  child.stdout.on('data', onData); child.stderr.on('data', onData);
  for (let i = 0; i < 80; i++) {
    await sleep(250);
    try { const r = await fetch(BASE + '/login.html'); if (r.ok) return; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-500));
  }
  throw new Error('伺服器啟動逾時：' + out.slice(-500));
}
async function stopServer() {
  if (!child) return;
  try { process.kill(childPid); } catch { /* 已結束 */ }
  for (let i = 0; i < 20 && child.exitCode === null; i++) await sleep(100);
  child = null;
}
let jar = {};
// 指定「來源位址」發請求（伺服器開發模式不信任 X-Forwarded-For，所以用本機迴路位址 127.0.0.x 當不同的來源 IP；Windows／Linux 都可綁）
function rawFrom(srcIp, method, url, body, headers) {
  return new Promise((resolve, reject) => {
    const data = body === undefined ? null : Buffer.from(JSON.stringify(body));
    const req = nodeHttp.request({ host: '127.0.0.1', port: PORT, path: url, method, localAddress: srcIp, headers: Object.assign({}, headers, data ? { 'Content-Length': data.length } : {}) }, (res) => {
      const chunks = []; res.on('data', c => chunks.push(c));
      res.on('end', () => resolve({ status: res.statusCode, buf: Buffer.concat(chunks), headers: { getSetCookie: () => [].concat(res.headers['set-cookie'] || []), get: (k) => res.headers[String(k).toLowerCase()] } }));
    });
    req.on('error', reject); if (data) req.write(data); req.end();
  });
}
async function http(who, method, url, body, hdrs) {
  const h = Object.assign({ 'Content-Type': 'application/json' }, hdrs || {});
  const src = h.__src; delete h.__src;
  if (who && jar[who]) h.Cookie = jar[who];
  const res = src ? await rawFrom(src, method, url, body, h) : await (async () => { const r0 = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' }); r0.buf = Buffer.from(await r0.arrayBuffer()); return r0; })();
  const buf = src ? res.buf : res.buf;
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  const sc = res.headers.getSetCookie ? res.headers.getSetCookie() : [];
  const fresh = sc.find(c => /^itts_auth=[^;]+/.test(c));
  if (who && fresh) jar[who] = fresh.split(';')[0];
  return { status: res.status, json, text: buf.toString('utf8'), headers: res.headers, setCookie: sc };
}
async function login(who, pw, ip) {
  const r = await http(null, 'POST', '/api/login', { username: who, password: pw || PW }, ip ? { __src: ip } : undefined);
  if (r.status === 200) jar[who] = r.setCookie.map(c => c.split(';')[0]).filter(c => /^itts_auth=/.test(c)).join('; ');
  return r;
}
const withCookie = (cookie, url) => fetch(BASE + url, { headers: { Cookie: cookie }, redirect: 'manual' }).then(async r => ({ status: r.status, text: await r.text() }));
async function fresh(serverFile, extraEnv) { await stopServer(); cleanupFiles(); seed(); await startServer(serverFile, extraEnv); jar = {}; }
async function logsOf(action) {
  await sleep(300);
  const r = await http('_s_adm1', 'GET', '/api/admin/logs?limit=500');
  const arr = Array.isArray(r.json) ? r.json : ((r.json && (r.json.logs || r.json.rows)) || []);
  return arr.filter(l => l.action === action);
}
const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));

// ── #6 M6：登入憑證綁密碼版本 ──
async function probeB6(T, serverFile) {
  await fresh(serverFile);
  let r = await login('_s_u1'); T.eq(r.status, 200, '[#6] 登入 200');
  r = await http('_s_u1', 'GET', '/api/me'); T.ok(r.status === 200 && r.json && r.json.username === '_s_u1' && !('pwa' in r.json), '[#6] /api/me 200 且不回傳 pwa（密碼版本不外洩）');
  const oldCookie = jar._s_u1;
  r = await http('_s_u1', 'PUT', '/api/user/password', { currentPassword: PW, newPassword: PW2 }); T.eq(r.status, 200, '[#6] 自己改密碼 200');
  r = await http('_s_u1', 'GET', '/api/me'); T.eq(r.status, 200, '[#6] 改密碼後，換發的新 cookie 仍有效（自己不會被登出）');
  const selfCookie = jar._s_u1;   // 自己改密碼時換發的 cookie：它帶著「新密碼版本」，之後密碼再被改它也要失效
  r = await withCookie(oldCookie, '/api/me'); T.eq(r.status, 401, '[#6] 改密碼前簽發的舊 cookie → 401（被竊 token 無法續用）');
  r = await login('_s_u1', PW2); T.eq(r.status, 200, '[#6] 用新密碼可正常重新登入');
  await login('_s_adm1');
  r = await http('_s_adm1', 'PUT', '/api/admin/users/_s_u1/password', { password: PW3 }); T.eq(r.status, 200, '[#6] 管理員再重設 u1 的密碼 200');
  r = await withCookie(selfCookie, '/api/me'); T.eq(r.status, 401, '[#6] 自己改密碼換發的 cookie 也帶著密碼版本：之後密碼再被重設 → 401（不是從此永不檢查）');
  // 管理員重設別人的密碼
  await login('_s_adm1'); await login('_s_u2');
  const u2old = jar._s_u2;
  r = await withCookie(u2old, '/api/me'); T.eq(r.status, 200, '[#6] 前提：u2 的 cookie 目前有效');
  r = await http('_s_adm1', 'PUT', '/api/admin/users/_s_u2/password', { password: PW3 }); T.eq(r.status, 200, '[#6] 管理員重設 u2 密碼 200');
  r = await withCookie(u2old, '/api/me'); T.eq(r.status, 401, '[#6] 管理員重設密碼後，u2 的舊 cookie → 401');
  r = await login('_s_u2', PW3); T.eq(r.status, 200, '[#6] u2 用新密碼登入');
  r = await http('_s_adm1', 'GET', '/api/me'); T.eq(r.status, 200, '[#6] 重設別人密碼不影響管理員自己的登入');
  // 管理員重設自己的密碼：不被自己登出
  const a2old = (await login('_s_adm2'), jar._s_adm2);
  r = await http('_s_adm2', 'PUT', '/api/admin/users/_s_adm2/password', { password: PW3 }); T.eq(r.status, 200, '[#6] 管理員重設自己的密碼 200');
  r = await http('_s_adm2', 'GET', '/api/me'); T.eq(r.status, 200, '[#6] 管理員重設自己的密碼後，本機換發的 cookie 仍有效');
  r = await withCookie(a2old, '/api/me'); T.eq(r.status, 401, '[#6] 但重設前的舊 cookie 失效');
  // 部署前簽發的舊版 token（沒有 pwa 欄位）放行（寬限）；token 內 pwa 與帳號不符 → 401
  const nowS = Math.floor(Date.now() / 1000);
  const mkTok = (user) => 'itts_auth=' + encodeURIComponent(jwtLib.sign({ user, lgn: nowS, lat: nowS, exp: nowS + 3600 }, srvSecret));
  r = await withCookie(mkTok({ username: '_s_u3', displayName: '_s_u3', role: 'user', bu: ['ERP'], loginAt: Date.now() }), '/api/me');
  T.eq(r.status, 200, '[#6] 舊版 token（沒有 pwa）→ 放行（部署當下不會全員被登出）');
  r = await withCookie(mkTok({ username: '_s_u3', displayName: '_s_u3', role: 'user', bu: ['ERP'], loginAt: Date.now(), pwa: '1999-01-01T00:00:00.000Z' }), '/api/me');
  T.eq(r.status, 401, '[#6] token 內 pwa 與帳號現值不符 → 401');
  const acct = JSON.parse(fs.readFileSync(path.join(ROOT, 'auth.json'), 'utf8')).users.find(u => u.username === '_s_u3');
  r = await withCookie(mkTok({ username: '_s_u3', displayName: '_s_u3', role: 'user', bu: ['ERP'], loginAt: Date.now(), pwa: acct.passwordChangedAt }), '/api/me');
  T.eq(r.status, 200, '[#6] token 內 pwa 與帳號現值相符 → 200');
  const rej = await logsOf('SESSION_REJECTED_INACTIVE');
  T.ok(rej.some(l => /密碼已變更/.test(l.detail)), '[#6] 舊 cookie 被拒有稽核（SESSION_REJECTED_INACTIVE，原因含「密碼已變更」）');
}
async function probeB6off(T, serverFile) {
  await fresh(serverFile, { SESSION_PWD_BINDING: 'off' });
  await login('_s_u1'); const old = jar._s_u1;
  await login('_s_adm1');
  let r = await http('_s_adm1', 'PUT', '/api/admin/users/_s_u1/password', { password: PW3 });
  T.eq(r.status, 200, '[#6] （開關）重設密碼 200');
  r = await withCookie(old, '/api/me'); T.eq(r.status, 200, '[#6] SESSION_PWD_BINDING=off：舊 cookie 維持有效（與改版前相同）');
}
// ── #5 M3：登入鎖定鍵含來源 IP ──
async function probeB5(T, serverFile) {
  const LOCK = { LOGIN_LOCKOUT: '1', LOGIN_MAX_FAILS: '5' };
  await fresh(serverFile, LOCK);
  const A = '127.0.0.3', V = '127.0.0.2';   // 攻擊者與真實管理員的來源 IP（本機迴路位址）
  for (let i = 0; i < 5; i++) await login('_s_adm1', 'wrong-' + i, A);
  let r = await login('_s_adm1', PW, A); T.eq(r.status, 401, '[#5] 攻擊者 IP 連續失敗 5 次後，連正確密碼也被鎖（鎖的是它自己的來源）');
  r = await login('_s_adm1', PW, V); T.eq(r.status, 200, '[#5] 真正的管理員從另一個 IP 仍可登入（匿名者無法把管理員鎖死）');
  for (let i = 0; i < 5; i++) await login('_s_adm2', 'wrong-' + i, A);
  r = await login('_s_adm2', PW, V); T.eq(r.status, 200, '[#5] 第二位管理員同樣不會被鎖死');
  r = await http('_s_adm1', 'GET', '/api/admin/login-locks', undefined, { __src: V });
  T.ok(r.status === 200 && Array.isArray(r.json.locks) && r.json.locks.some(l => l.username === '_s_adm1') && r.json.locks.some(l => l.username === '_s_adm2'), '[#5] 管理員後台鎖定清單看得到被鎖的來源（形狀不變）');
  const lf = JSON.stringify(readData().loginFailures || {});
  T.ok(lf.length > 2 && !lf.includes(A) && !lf.includes(V) && !lf.includes('127.0.0'), '[#5] 主資料 blob 的鎖定表不存 IP 明文');
  r = await http('_s_adm1', 'POST', '/api/admin/users/_s_adm1/unlock', undefined, { __src: V });
  T.ok(r.status === 200 && r.json.wasLocked === true, '[#5] 管理員解鎖');
  r = await login('_s_adm1', PW, A); T.eq(r.status, 200, '[#5] 解鎖後攻擊者 IP（此時用正確密碼）可登入');
  // 開關：LOGIN_LOCK_SCOPE=name 還原舊行為
  await fresh(serverFile, Object.assign({ LOGIN_LOCK_SCOPE: 'name' }, LOCK));
  for (let i = 0; i < 5; i++) await login('_s_adm1', 'wrong-' + i, A);
  r = await login('_s_adm1', PW, V); T.eq(r.status, 401, '[#5] LOGIN_LOCK_SCOPE=name：任何來源都被鎖（舊行為，緊急還原用）');
  // 預設（LOGIN_LOCKOUT 未開）不受影響
  await fresh(serverFile, {});
  for (let i = 0; i < 8; i++) await login('_s_adm1', 'wrong-' + i, A);
  r = await login('_s_adm1', PW, A); T.eq(r.status, 200, '[#5] LOGIN_LOCKOUT 未開（預設）：連續失敗也不鎖，行為與改版前相同');
}
// ── #12 H1：管理員帳號異動通知 ──
async function probeB12(T, serverFile) {
  await fresh(serverFile);
  await login('_s_adm1'); await login('_s_adm2');
  const notifsTo = (u, type) => (readData().notifications || []).filter(n => n.to === u && n.type === (type || 'admin_account_change'));
  let r = await http('_s_adm1', 'POST', '/api/admin/users', { username: '_s_new1', password: rnd('Sq'), role: 'user', bu: ['ERP'] });
  T.eq(r.status, 200, '[#12] 建立一般使用者 200');
  T.ok(notifsTo('_s_adm2').length === 0 && (await logsOf('ADMIN_ACCOUNT_CHANGE')).length === 0, '[#12] 建立一般使用者：不通知、不記 ADMIN_ACCOUNT_CHANGE');
  const secret = rnd('Sq');
  r = await http('_s_adm1', 'POST', '/api/admin/users', { username: '_s_newadm', password: secret, role: 'admin' });
  T.eq(r.status, 200, '[#12] 建立管理員帳號 200（' + r.status + ' ' + r.text.slice(0, 100) + '）');
  let n2 = notifsTo('_s_adm2');
  T.ok(n2.length === 1 && n2[0].refId === '_s_newadm' && /新增管理員帳號/.test(n2[0].title + n2[0].body) && n2[0].read === false, '[#12] 其他現任管理員收到站內通知（動作、對象）');
  T.eq(notifsTo('_s_adm1').length, 0, '[#12] 操作者自己不收通知');
  T.eq(notifsTo('_s_adm3').length, 0, '[#12] 已停用的管理員不收通知');
  T.eq(notifsTo('_s_u1').length, 0, '[#12] 非管理員不收通知');
  let lg = await logsOf('ADMIN_ACCOUNT_CHANGE');
  T.ok(lg.length === 1 && lg[0].operator === '_s_adm1' && lg[0].target === '_s_newadm' && /新增管理員帳號/.test(lg[0].detail), '[#12] 稽核 ADMIN_ACCOUNT_CHANGE：操作者、對象、動作');
  T.ok(!JSON.stringify(lg).includes(secret) && !JSON.stringify(readData().notifications).includes(secret), '[#12] 稽核與通知都不含密碼');
  const apiN = await http('_s_adm2', 'GET', '/api/notifications');
  T.ok(apiN.status === 200 && JSON.stringify(apiN.json).includes('_s_newadm'), '[#12] 通知經 GET /api/notifications 對管理員可見（既有鈴鐺機制）');
  // 升格
  r = await http('_s_adm1', 'PUT', '/api/admin/users/_s_u1', { role: 'admin' });
  T.eq(r.status, 200, '[#12] 把一般使用者升格為管理員 200（' + r.status + ' ' + r.text.slice(0, 100) + '）');
  lg = (await logsOf('ADMIN_ACCOUNT_CHANGE')).filter(l => /升格為管理員/.test(l.detail));
  T.ok(lg.length === 1 && lg[0].target === '_s_u1', '[#12] 升格有 ADMIN_ACCOUNT_CHANGE');
  T.ok(notifsTo('_s_adm2').some(n => n.refId === '_s_u1'), '[#12] 升格會通知其他管理員');
  r = await http('_s_adm1', 'PUT', '/api/admin/users/_s_u2', { displayName: '改名不升格' });
  T.eq((await logsOf('ADMIN_ACCOUNT_CHANGE')).filter(l => l.target === '_s_u2').length, 0, '[#12] 修改一般使用者（不涉及升格）不記 ADMIN_ACCOUNT_CHANGE');
  // 重設管理員密碼
  r = await http('_s_adm1', 'PUT', '/api/admin/users/_s_adm2/password', { password: PW3 });
  T.eq(r.status, 200, '[#12] 重設另一位管理員的密碼 200');
  lg = (await logsOf('ADMIN_ACCOUNT_CHANGE')).filter(l => /重設管理員密碼/.test(l.detail));
  T.ok(lg.length === 1 && lg[0].target === '_s_adm2' && !JSON.stringify(lg).includes(PW3), '[#12] 重設管理員密碼有 ADMIN_ACCOUNT_CHANGE（不含密碼）');
  T.ok(notifsTo('_s_adm2').some(n => n.refId === '_s_adm2' && /重設管理員密碼/.test(n.title)), '[#12] 被重設密碼的那位管理員本人收到通知');
  const before = (readData().notifications || []).length;
  r = await http('_s_adm1', 'PUT', '/api/admin/users/_s_new1/password', { password: rnd('Sq') });
  T.ok(r.status === 200 && (readData().notifications || []).length === before && (await logsOf('ADMIN_ACCOUNT_CHANGE')).filter(l => l.target === '_s_new1').length === 0, '[#12] 重設一般使用者密碼：不通知、不記 ADMIN_ACCOUNT_CHANGE');
  T.ok((await logsOf('CREATE_USER')).length >= 2 && (await logsOf('RESET_PASSWORD')).length >= 2, '[#12] 既有稽核 CREATE_USER／RESET_PASSWORD 仍在（沒有被取代）');
}
async function partB(T, serverFile, only) {
  if (!only || only.includes('6')) { await probeB6(T, serverFile); }
  if (!only || only.includes('5')) { await probeB5(T, serverFile); }
  if (!only || only.includes('12')) { await probeB12(T, serverFile); }
}
function mutateServer(edits) {
  let src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  for (const [from, to, name] of edits) { if (!src.includes(from)) throw new Error('變異樣式找不到：' + name); src = src.replace(from, () => to); }
  const out = path.join(ROOT, 'server.__mut_sq.js'); fs.writeFileSync(out, src); tmpFiles.push(out); return out;
}
async function mutB(name, edits, only, extra) {
  const f = mutateServer(edits);
  const T = mkT();
  try { if (extra) await extra(T, f); else await partB(T, f, only); } catch (e) { T.f++; T.fl.push('例外：' + (e && e.message)); }
  await stopServer();
  try { fs.unlinkSync(f); } catch { /* 已刪 */ }
  ok(T.f > 0, `變異「${name}」→ 對應檢查必須失敗，實際失敗 ${T.f} 項` + (T.f ? '' : '（變異未被抓到！）'));
}

// ═════════════════════════ C：靜態 ═════════════════════════
function partC() {
  const routes = fs.readFileSync(path.join(ROOT, 'lib', 'quoteRoutes.js'), 'utf8');
  const fn = routes.slice(routes.indexOf('function auditCrossBu('));
  const body = fn.slice(0, fn.indexOf('\n  }\n') + 4);
  ok(/flushAuditFor\(req\);[^\n]*\n\s*writeLog\('VIEW_QUOTE_CROSS_BU'/.test(body), '[#1] auditCrossBu 在 writeLog 之前先 flushAuditFor(req)');
  let files = [];
  try { files = execFileSync('git', ['-c', 'core.quotepath=false', 'ls-files'], { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }).split('\n').filter(f => f && !f.startsWith('scratchpad/') && !/\.(png|jpg|jpeg|gif|ico|xlsx|pptx|woff2?|pdf|zip)$/i.test(f)); } catch (e) { ok(false, '[#11] 無法執行 git ls-files：' + e.message); return; }
  const hits = [];
  for (const f of files) {
    let s; try { s = fs.readFileSync(path.join(ROOT, f), 'utf8'); } catch { continue; }
    if (/steven[.\-]lee/i.test(s)) hits.push(f);
  }
  ok(hits.length === 0, '[#11] 公開 repo 的追蹤檔不含本機路徑／使用者名稱（本機使用者帳號名稱）：' + (hits.join(',') || '0 筆'));
  const vbs = files.find(f => f.endsWith('.vbs'));
  ok(vbs && !/[A-Za-z]:\\Users\\/.test(fs.readFileSync(path.join(ROOT, vbs), 'utf8')) && /GetParentFolderName\(WScript\.ScriptFullName\)/.test(fs.readFileSync(path.join(ROOT, vbs), 'utf8')), '[#11] 啟動伺服器.vbs 改用相對（腳本所在資料夾）');
  for (const f of ['make_ppt.js', 'make_ppt2.js', 'make_ppt3.js']) ok(!/['"][A-Za-z]:[\\/]/.test(fs.readFileSync(path.join(ROOT, f), 'utf8')) && /PPT_OUT_DIR/.test(fs.readFileSync(path.join(ROOT, f), 'utf8')), `[#11] ${f} 輸出路徑不寫死（PPT_OUT_DIR 或使用者 Documents）`);
  const cdr = fs.readFileSync(path.join(ROOT, 'scripts', 'check-config-dual-review.js'), 'utf8');
  ok(!/console\.log\([^)]*\bpw=' \+ PW\)/.test(cdr) && /CDR_SERVE_SHOW_PW/.test(cdr), '[#11] check-config-dual-review 預設不再把測試密碼印到終端機');
  const sw = fs.readFileSync(path.join(ROOT, '_client', 'sw.js'), 'utf8');
  const swm = /SW_VERSION = 'itts-crm-v(\d+)'/.exec(sw);
  ok(swm && parseInt(swm[1], 10) >= 18, '[sw] _client/sw.js 版號已 +1（>= v18）（' + (swm && swm[0]) + '）');
  const csp = fs.readFileSync(path.join(ROOT, 'lib', 'cspReport.js'), 'utf8');
  ok(/KNOWN_DOC_PATHS/.test(csp) && /LEGACY_MAX_WRITES = 4/.test(csp), '[#4] cspReport 有來源頁面白名單與舊路徑預算 4（詳細行為見 check-csp-report.js）');
}

// ═════════════════════════ 主程式 ═════════════════════════
(async () => {
  const t0 = Date.now();
  if (want('B') || want('M')) guardClean();
  try {
    const realMods = () => ({ routes: require(path.join(ROOT, 'lib', 'quoteRoutes.js')), fd: require(path.join(ROOT, 'lib', 'quoteFlowDesigner.js')) });
    if (want('A')) { console.log('[A] 行程內探針（#1 #2 #7 #9）'); await partA(realMods(), G); G.n += 0; if (G.f) { G.fl.forEach(f => console.log('  ✗ FAIL:', f)); } }
    if (want('C')) { console.log('[C] 靜態（#11 #1 sw）'); partC(); }
    if (want('B')) {
      console.log('[B] 真實伺服器（#6 #5 #12），埠 ' + PORT);
      await probeB6(G, null); await probeB6off(G, null); await probeB5(G, null); await probeB12(G, null);
      if (G.f) G.fl.forEach(f => console.log('  ✗ FAIL:', f));
    }
    if (want('M') && !process.env.SQ_SKIP_MUT) {
      console.log('[M] 變異檢查（每個變異都必須讓對應檢查失敗）');
      const R = 'routes', F = 'fd';
      await mutA('#1 auditCrossBu 不 flush', 'quoteRoutes.js', [['    flushAuditFor(req);   // 資安審查 M2：GET 路徑在 Vercel', '    // flushAuditFor(req);   // 資安審查 M2：GET 路徑在 Vercel', 'm1a']], ['1'], R);
      await mutA('#1 清單不記彙總稽核', 'quoteRoutes.js', [['    auditCrossBuList(visible, ctx, req);', '', 'm1b']], ['1'], R);
      await mutA('#1 彙總稽核不節流', 'quoteRoutes.js', [['if (crossBuListSeen.get(ctx.me) > now - CROSS_BU_AUDIT_MS) return;', '', 'm1c']], ['1'], R);
      await mutA('#1 權限類型報告不 flush', 'quoteRoutes.js', [['    flushAuditFor(req);   // 資安審查 M2：GET 路徑稽核在回應前落地\n    writeLog(\'VIEW_QUOTE_MODE_REPORT\'', '    writeLog(\'VIEW_QUOTE_MODE_REPORT\'', 'm1d']], ['1'], R);
      await mutA('#2 userMap 回到 {}', 'quoteRoutes.js', [['const m = Object.create(null); (auth.users', 'const m = {}; (auth.users', 'm2a'], ['const hasUser = (users, un) => typeof un === \'string\' && !!users && Object.prototype.hasOwnProperty.call(users, un) && !!users[un];', 'const hasUser = (users, un) => typeof un === \'string\' && !!users[un];', 'm2a2']], ['2'], R);
      await mutA('#2 例外元素不驗證型別', 'quoteRoutes.js', [['if (!e || typeof e !== \'object\' || Array.isArray(e)) return', 'if (false) return', 'm2b']], ['2'], R);
      await mutA('#2 例外欄位不驗長度', 'quoteRoutes.js', [['if (e[k] != null && (typeof e[k] !== \'string\' || e[k].length > max)) return', 'if (false) return', 'm2c']], ['2'], R);
      await mutA('#7 截斷不 fail-closed', 'quoteFlowDesigner.js', [['if (structural && a.truncated) i10.push(', 'if (false) i10.push(', 'm7a'], ['(a.violationCount === 0 && !a.truncated)', '(a.violationCount === 0)', 'm7a2']], ['7'], F);
      await mutA('#7 商品類別鍵只取前 6 個', 'quoteFlowDesigner.js', [['const ks = [...use.keys].filter((k) => QA.CLASS_KEYS.includes(k));', 'const ks = [...use.keys].filter((k) => QA.CLASS_KEYS.includes(k)).slice(0, 6);', 'm7b']], ['7'], F);
      await mutA('#7 例外帳號上限被截去不標記', 'quoteFlowDesigner.js', [['if (use.skip.size > 8) axisCut = true;', '', 'm7c']], ['7'], F);
      await mutA('#9 時間軸成本合計不遮', 'quoteRoutes.js', [['comment = comment.replace(COST_TOTAL_RE, \'\');', '', 'm9a']], ['9'], R);
      await mutA('#9 preview.marginText 不遮', 'quoteRoutes.js', [['if (!perm.canSeeCost && out.preview.marginText != null) out.preview =', 'if (false) out.preview =', 'm9b']], ['9'], R);
      await mutA('#9 看得到成本的人也被遮（過度遮罩）', 'quoteRoutes.js', [['if (!(perm && perm.canSeeCost) && (h.action', 'if (true && (h.action', 'm9c']], ['9'], R);
      await mutB('#6 不比對 pwa', [['typeof su.pwa === \'string\' && su.pwa !== (acct.passwordChangedAt || \'\')', 'false', 'm6a']], ['6']);
      await mutB('#6 登入不寫入 pwa', [['pwa: user.passwordChangedAt || \'\',   // 資安審查 M6：登入當下', '   // 資安審查 M6：登入當下', 'm6b']], ['6']);
      await mutB('#6 自己改密碼不換發 pwa', [['pwa: user.passwordChangedAt   // 自己改密碼', 'pwa: \'stale\'   // 自己改密碼', 'm6c']], ['6']);
      await mutB('#6 自己改密碼後 session 不再帶 pwa（從此永不檢查）', [['pwa: user.passwordChangedAt   // 自己改密碼', 'pwa: undefined   // 自己改密碼', 'm6c2']], ['6']);
      await mutB('#6 /api/me 外洩 pwa', [['res.json({ ...pub, bu: normalizeBu(me?.bu) });', 'res.json({ ...pub, pwa: req.session.user.pwa, bu: normalizeBu(me?.bu) });', 'm6d']], ['6']);
      await mutB('#6 管理員重設自己的密碼不換發', [['if (req.session.user && req.session.user.username === user.username) req.session.user = { ...req.session.user, pwa: user.passwordChangedAt };', '', 'm6e']], ['6']);
      await mutB('#5 鎖定檢查不帶 IP', [['loginGuard.checkLocked(username, req.ip).locked', 'loginGuard.checkLocked(username).locked', 'm5a'], ['loginGuard.recordFailure(attempted, req.ip)', 'loginGuard.recordFailure(attempted)', 'm5b']], ['5']);
      await mutB('#12 建立管理員不通知', [['if (finalRole === \'admin\') adminAccountChange(req, \'CREATE_ADMIN\'', 'if (false) adminAccountChange(req, \'CREATE_ADMIN\'', 'm12a']], ['12']);
      await mutB('#12 升格不通知', [['userAfter.role === \'admin\') adminAccountChange(req, \'PROMOTE_ADMIN\'', 'userAfter.role === \'__never__\') adminAccountChange(req, \'PROMOTE_ADMIN\'', 'm12b']], ['12']);
      await mutB('#12 重設管理員密碼不通知', [['if (user.role === \'admin\') adminAccountChange(req, \'RESET_ADMIN_PASSWORD\'', 'if (false) adminAccountChange(req, \'RESET_ADMIN_PASSWORD\'', 'm12c']], ['12']);
      await mutB('#12 通知對象包含操作者與停用帳號', [['u.role === \'admin\' && u.active !== false && u.username !== actor', 'u.role === \'admin\'', 'm12d']], ['12']);
      await mutB('#12 只通知不記稽核', [['try { writeLog(\'ADMIN_ACCOUNT_CHANGE\', actor,', 'try { void (\'ADMIN_ACCOUNT_CHANGE\', actor,', 'm12e']], ['12']);
    }
  } catch (e) {
    G.f++; G.fl.push('例外：' + (e && e.stack || e)); console.log('✗ 例外', e && e.stack || e);
  } finally {
    await stopServer();
    cleanupFiles();
    tmpFiles.forEach(f => { try { fs.unlinkSync(f); } catch { /* 已刪 */ } });
  }
  pass = G.n; fail = G.f;
  console.log(`\ncheck-secrev-quick 結果：${pass} 通過，${fail} 失敗（${((Date.now() - t0) / 1000).toFixed(1)} 秒）`);
  if (fail) { console.log(G.fl.map(f => ' - ' + f).join('\n')); process.exit(1); }
  console.log('全部通過'); process.exit(0);
})();
