#!/usr/bin/env node
'use strict';
/**
 * 內控稽核第一波（audit wave 1）端對端檢查。不含任何帳密：測試帳號與隨機密碼只在記憶體裡。
 *
 * 做法：在「本工作樹」建立臨時 auth.json / data.json（假資料 + _s_* 臨時帳號），自行以 PORT=AW1_PORT（預設 3121）
 * 啟動 server.js，跑完後只用 PID 關閉並刪除臨時檔。
 * 守門：工作樹若已有 auth.json / data.json / audit.log.json（可能是真實資料），一律拒絕執行。
 * 禁用埠：3000 / 3020 / 3030 / 3101 / 3111。
 *
 *   node scripts/check-audit-wave1.js              # 全部
 *   AW1_ONLY=1,2 node scripts/check-audit-wave1.js # 只跑指定項目
 *   AW1_SKIP_MUT=1 ...                             # 略過變異檢查
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.AW1_PORT || '3121', 10);
if ([3000, 3020, 3030, 3101, 3111].includes(PORT)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json'].map(f => path.join(ROOT, f));
const guardFiles = FILES.slice(0, 3).filter(f => fs.existsSync(f));
if (guardFiles.length) { console.log('拒絕：此工作樹已有', guardFiles.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));

// ───────── 沙盒資料 ─────────
const PW = 'Aw1' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
function mkQuote(tag, steps, cur) {
  return { id: 'q-' + tag, quoteNo: 'SB-' + tag, company: '沙盒公司甲', owner: '_s_u1', status: 'draft', createdAt: nowIso, updatedAt: nowIso, products: [], items: [],
    approval: { state: 'pending', rulesVersion: 1, hash: 'h', submittedAt: nowIso, submittedBy: '_s_u1', steps, cur, board: null, derived: null, history: [{ at: nowIso, by: '_s_u1', action: 'SUBMIT', comment: '', tier: '' }] } };
}
function seedFiles() {
  const users = [
    mkUser('_s_admin', 'admin'), mkUser('_s_admin2', 'admin'),
    mkUser('_s_m1', 'manager1', { bu: ['ERP'], canDownloadContacts: true }),
    mkUser('_s_u1', 'user', { bu: ['ERP'], supervisor: '_s_m1', canDownloadContacts: true }),
    mkUser('_s_dl', 'user', { bu: ['ERP'], canDownloadContacts: true }),
    mkUser('_s_off', 'user', { bu: ['ERP'] }),
    mkUser('_s_gone', 'user', { bu: ['ERP'] }),
    mkUser('_s_dem', 'manager1', { bu: ['ERP'] }),
    mkUser('_s_view', 'user', { bu: ['ERP'], accessMode: 'view' }),
    mkUser('_s_mc', 'user', { bu: ['ERP'], mustChangePassword: true }),
    mkUser('_s_tpm', 'tecopm', { bu: [], viewOwnerScope: '_s_u1', viewGroupId: 'g1' }),
    mkUser('_q_m1', 'manager1', { bu: ['ERP'] }), mkUser('_q_m1b', 'manager1', { bu: ['ERP'] }),
    mkUser('_q_m2', 'manager2', { bu: ['ERP'] }), mkUser('_q_m2b', 'manager2', { bu: ['ERP'] }),
    mkUser('_q_off', 'manager1', { bu: ['ERP'], active: false }),
  ];
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  const d = {
    contacts: [
      { id: 'c1', name: '測試甲', company: '沙盒公司甲', owner: '_s_u1', bu: 'ERP', email: 'a@sandbox.test', phone: '0200000001', createdAt: nowIso },
      { id: 'c2', name: '測試乙', company: '沙盒公司乙', owner: '_s_u1', bu: 'ERP', email: 'b@sandbox.test', phone: '0200000002', createdAt: nowIso },
    ],
    companies: [{ id: 'co1', name: '沙盒公司甲', taxId: '12345675', industry: '製造業' }],
    opportunities: [{ id: 'o1', title: '沙盒商機', company: '沙盒公司甲', owner: '_s_u1', category: 'ERP', stage: 'B', amount: 100, expectedDate: new Date().getFullYear() + '-12-01', createdAt: nowIso }],
    visits: [{ id: 'v1', visitDate: '2026-01-01', owner: '_s_u1', contactName: '測試甲', topic: '沙盒', createdAt: nowIso }],
    contracts: [{ id: 'k1', contractNo: 'SB-1', company: '沙盒公司甲', owner: '_s_u1', createdAt: nowIso }],
    quotations: [
      // QA：一般流程 [mgr1(進行中), gm]，cur=0
      mkQuote('QA', [{ tier: 'mgr1', label: '一級主管', assignee: '_q_m1', status: 'pending' }, { tier: 'gm', label: '總經理', assignee: null, status: 'waiting' }], 0),
      // QB：BU 啟用二級主管流程 [mgr2(已核准), mgr1(進行中), gm]，cur=1 —— 舊程式會誤改 steps[0]
      mkQuote('QB', [{ tier: 'mgr2', label: '二級主管', assignee: '_q_m2', status: 'approved', by: '_q_m2', at: nowIso, comment: '同意' }, { tier: 'mgr1', label: '一級主管', assignee: '_q_m1', status: 'pending' }, { tier: 'gm', label: '總經理', assignee: null, status: 'waiting' }], 1),
      // QC：二級主管關進行中 [mgr2(進行中), mgr1, gm]，cur=0
      mkQuote('QC', [{ tier: 'mgr2', label: '二級主管', assignee: '_q_m2', status: 'pending' }, { tier: 'mgr1', label: '一級主管', assignee: '_q_m1', status: 'waiting' }, { tier: 'gm', label: '總經理', assignee: null, status: 'waiting' }], 0),
      // QD：已核准（不可改派）
      Object.assign(mkQuote('QD', [{ tier: 'mgr1', label: '一級主管', assignee: '_q_m1', status: 'approved', by: '_q_m1', at: nowIso }], 1), { approvalState: 'approved' }),
    ],
    receivables: [], callins: [], keyAccounts: [], groups: [{ id: 'g1', name: '沙盒集團', memberCompanies: ['沙盒公司甲'] }],
  };
  d.quotations.forEach(q => { if (q.approvalState) { q.approval.state = q.approvalState; delete q.approvalState; } });
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
}
function cleanupFiles() { FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }

// ───────── 伺服器 ─────────
let child = null, childPid = null;
let extraEnv = {};
async function startServer(serverFile) {
  const env = Object.assign({}, process.env, extraEnv, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  delete env.DATABASE_URL; delete env.MAIL_MODE;
  child = spawn(process.execPath, [serverFile || path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  const onData = d => { out += d; if (process.env.AW1_DEBUG) process.stdout.write('[srv] ' + d); };
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
async function restart(serverFile, env) { await stopServer(); extraEnv = env || {}; await startServer(serverFile); jar = {}; }

// ───────── HTTP ─────────
let jar = {};
async function http(who, method, url, body, opt) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  if (opt && opt.ua) h['User-Agent'] = opt.ua;
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  const sc = res.headers.getSetCookie ? res.headers.getSetCookie() : [];
  const fresh = sc.find(c => /^itts_auth=[^;]+/.test(c));
  if (who && fresh) jar[who] = fresh.split(';')[0];   // 伺服器重簽 cookie（例如改密碼後）就跟著換
  return { status: res.status, json, buf, text: buf.toString('utf8'), headers: res.headers, setCookie: res.headers.getSetCookie ? res.headers.getSetCookie() : [] };
}
async function login(who, user, pw) {
  const r = await http(null, 'POST', '/api/login', { username: user || who, password: pw || PW });
  if (r.status === 200) jar[who] = r.setCookie.map(c => c.split(';')[0]).join('; ');
  return r;
}
async function loginAll(names) { for (const n of names) { const r = await login(n); if (r.status !== 200) throw new Error('登入失敗 ' + n + ' ' + r.status); } }
async function getLogs() { const r = await http('_s_admin', 'GET', '/api/admin/logs'); return Array.isArray(r.json) ? r.json : []; }
const byAction = (logs, a) => logs.filter(l => l.action === a);
async function adminPutUser(u, body, who) { return http(who || '_s_admin', 'PUT', '/api/admin/users/' + encodeURIComponent(u), body); }

// ═════════════ 第 1 項：停用／異動即時生效 ═════════════
async function sec1() {
  console.log('[1] 停用／異動即時生效');
  await loginAll(['_s_admin', '_s_off', '_s_gone', '_s_m1', '_s_dem', '_s_tpm', '_s_mc', '_s_view', '_s_u1']);
  let r = await http('_s_off', 'GET', '/api/me');
  eq(r.status, 200, '停用前 /api/me 200');
  ok(!/\$2[aby]\$|"password"/i.test(r.text), '/api/me 不含密碼雜湊');

  // 停用 → 立刻 401（遠小於 TTL）
  const t0 = Date.now();
  r = await adminPutUser('_s_off', { active: false });
  eq(r.status, 200, '管理員停用 _s_off');
  r = await http('_s_off', 'GET', '/api/me');
  eq(r.status, 401, '停用後下一個請求 401');
  eq(r.json && r.json.code, 'SESSION_INVALID', '401 帶 code SESSION_INVALID');
  ok(Date.now() - t0 < 4000, '停用到拒絕 < 4 秒（TTL 5 秒內）');
  ok(r.setCookie.some(c => /itts_auth=;|itts_auth=\s*;|Expires=Thu, 01 Jan 1970/i.test(c)), '回應清除 cookie');
  // 連續打其他 API 也被擋、且稽核只記一筆（節流）
  r = await http('_s_off', 'GET', '/api/contacts');
  eq(r.status, 401, '被清除後仍以舊 cookie 打其他 API → 401');
  let logs = await getLogs();
  const rej = byAction(logs, 'SESSION_REJECTED_INACTIVE').filter(l => l.operator === '_s_off');
  eq(rej.length, 1, 'SESSION_REJECTED_INACTIVE 同一人一分鐘只記一筆');
  ok(rej[0] && /已停用/.test(rej[0].detail), '稽核內容說明原因（已停用）');
  // 舊 cookie 不是未登入：直接重用舊 cookie 打
  r = await http('_s_off', 'GET', '/api/me');
  eq(r.status, 401, '重用舊 cookie 仍 401');
  // 重新啟用後可再登入
  await adminPutUser('_s_off', { active: true });
  r = await login('_s_off');
  eq(r.status, 200, '重新啟用後可登入');

  // 刪除帳號 → 401
  r = await http('_s_gone', 'GET', '/api/me');
  eq(r.status, 200, '刪除前 200');
  r = await http('_s_admin', 'DELETE', '/api/admin/users/_s_gone');
  eq(r.status, 200, '刪除帳號');
  r = await http('_s_gone', 'GET', '/api/contacts');
  eq(r.status, 401, '帳號被刪後 401');
  logs = await getLogs();
  ok(byAction(logs, 'SESSION_REJECTED_INACTIVE').some(l => l.operator === '_s_gone' && /不存在/.test(l.detail)), '刪除帳號的拒絕有稽核（帳號已不存在）');

  // 降權：manager1 本來看得到下屬 _s_u1 的聯絡人；降成 user 後即時看不到
  r = await http('_s_m1', 'GET', '/api/contacts');
  ok(Array.isArray(r.json) && r.json.some(c => c.id === 'c1'), '降權前 manager1 看得到下屬聯絡人');
  r = await http('_s_dem', 'GET', '/api/me');
  eq(r.json && r.json.role, 'manager1', '降權前 session role=manager1');
  r = await adminPutUser('_s_m1', { role: 'user' });
  eq(r.status, 200, '管理員把 _s_m1 降為 user');
  r = await http('_s_m1', 'GET', '/api/me');
  eq(r.json && r.json.role, 'user', '降權後同一 cookie 的 role 即時變 user（以 DB 為準）');
  r = await http('_s_m1', 'GET', '/api/contacts');
  ok(Array.isArray(r.json) && !r.json.some(c => c.id === 'c1'), '降權後即時看不到下屬聯絡人');
  await adminPutUser('_s_m1', { role: 'manager1' });
  r = await http('_s_m1', 'GET', '/api/contacts');
  ok(Array.isArray(r.json) && r.json.some(c => c.id === 'c1'), '還原角色後立即恢復');

  // 唯讀即時生效（accessMode 本來就即時，確認未被破壞）＋ 改成唯讀
  r = await http('_s_u1', 'POST', '/api/contacts', { name: '測試丙', company: '沙盒公司甲', taxId: '12345675', bu: 'ERP', email: 'c@sandbox.test' });
  ok([200, 201].includes(r.status), '一般帳號可新增聯絡人 (' + r.status + ')');
  await adminPutUser('_s_u1', { accessMode: 'view' });
  r = await http('_s_u1', 'POST', '/api/contacts', { name: '測試丁', company: '沙盒公司甲', taxId: '12345675', bu: 'ERP', email: 'd@sandbox.test' });
  eq(r.status, 403, '改為唯讀後寫入即時 403');
  await adminPutUser('_s_u1', { accessMode: 'edit' });

  // 集團角色白名單仍有效
  r = await http('_s_tpm', 'GET', '/api/contacts');
  eq(r.status, 200, 'tecopm 白名單內端點 200');
  r = await http('_s_tpm', 'GET', '/api/quotations');
  eq(r.status, 403, 'tecopm 白名單外端點 403');
  r = await http('_s_tpm', 'POST', '/api/contacts', { name: 'x' });
  eq(r.status, 403, 'tecopm 非 GET 403');
  r = await http('_s_tpm', 'GET', '/api/me');
  eq(r.json && r.json.viewOwnerScope, '_s_u1', 'tecopm viewOwnerScope 由 DB 帶入');
  await adminPutUser('_s_tpm', { viewOwnerScope: '_s_m1' });
  r = await http('_s_tpm', 'GET', '/api/me');
  eq(r.json && r.json.viewOwnerScope, '_s_m1', 'tecopm viewOwnerScope 異動即時生效');

  // 強制改密碼流程
  r = await http('_s_mc', 'GET', '/api/contacts');
  eq(r.status, 403, 'mustChangePassword 仍被擋');
  eq(r.json && r.json.code, 'MUST_CHANGE_PASSWORD', '帶 MUST_CHANGE_PASSWORD');
  r = await http('_s_mc', 'GET', '/api/me');
  eq(r.status, 200, 'mustChangePassword 仍可 /api/me');
  const newPw = PW + 'n1';
  r = await http('_s_mc', 'PUT', '/api/user/password', { currentPassword: PW, newPassword: newPw });
  eq(r.status, 200, '強制改密碼流程可完成');
  r = await http('_s_mc', 'GET', '/api/contacts');
  eq(r.status, 200, '改完密碼後解除限制');

  // 效能：連打 /api/me 不應慢（JSON 模式有 mtime+TTL 快取）
  const t1 = Date.now();
  for (let i = 0; i < 100; i++) await http('_s_admin', 'GET', '/api/me');
  const per = (Date.now() - t1) / 100;
  console.log(`  (100 次 /api/me 平均 ${per.toFixed(1)} ms)`);
  ok(per < 60, '每請求平均 < 60ms');
}

// 第 1 項補充：本地 express-session（SESSION_BACKEND=memory）路徑也要即時生效
async function sec1b() {
  console.log('[1b] 本地 express-session 路徑');
  await restart(null, { SESSION_BACKEND: 'memory', SESSION_SECRET: crypto.randomBytes(24).toString('hex') });
  await loginAll(['_s_admin', '_s_off', '_s_dem']);
  let r = await http('_s_off', 'GET', '/api/me');
  eq(r.status, 200, 'memory 模式：停用前 200');
  await adminPutUser('_s_off', { active: false });
  r = await http('_s_off', 'GET', '/api/me');
  eq(r.status, 401, 'memory 模式：停用後 401');
  eq(r.json && r.json.code, 'SESSION_INVALID', 'memory 模式：code');
  r = await http('_s_dem', 'GET', '/api/me');
  eq(r.json && r.json.role, 'manager1', 'memory 模式：降權前 manager1');
  await adminPutUser('_s_dem', { role: 'user' });
  r = await http('_s_dem', 'GET', '/api/me');
  eq(r.json && r.json.role, 'user', 'memory 模式：降權即時生效');
}

// ═════════════ 第 2 項：資料匯出稽核 ═════════════
async function sec2() {
  console.log('[2] 資料匯出稽核');
  await loginAll(['_s_admin', '_s_u1', '_s_off', '_s_tpm']);
  const yr = new Date().getFullYear();
  // 端點清單（新增匯出端點時要同步加進來；範本類無資料不在內）
  const EXPORTS = [
    { url: '/api/export', who: '_s_u1', action: 'EXPORT_CONTACTS' },
    { url: '/api/forecast/export?year=' + yr, who: '_s_u1', action: 'EXPORT_FORECAST' },
    { url: '/api/admin/opportunities/export', who: '_s_admin', action: 'EXPORT_OPPORTUNITIES' },
    { url: '/api/admin/contacts/export', who: '_s_admin', action: 'EXPORT_CONTACTS_ADMIN' },
    { url: '/api/admin/visits/export', who: '_s_admin', action: 'EXPORT_VISITS' },
    { url: '/api/admin/contracts/export', who: '_s_admin', action: 'EXPORT_CONTRACTS' },
    { url: '/api/companies/export', who: '_s_admin', action: 'EXPORT_COMPANY_MASTER' },
  ];
  for (const e of EXPORTS) {
    const before = byAction(await getLogs(), e.action).length;
    const r = await http(e.who, 'GET', e.url);
    eq(r.status, 200, e.url + ' 匯出成功');
    const rows = byAction(await getLogs(), e.action);
    eq(rows.length, before + 1, e.url + ' 寫入一筆 ' + e.action);
    const l = rows[0] || {};
    eq(l.operator, e.who, e.action + ' operator');
    ok(/匯出 [0-9]+ 筆/.test(l.detail || ''), e.action + ' detail 含筆數：' + l.detail);
    ok(/範圍：/.test(l.detail || '') && /格式：xlsx/.test(l.detail || ''), e.action + ' detail 含範圍與格式');
    ok(!/沙盒公司|sandbox[.]test|0200000001|測試甲/.test(l.detail || ''), e.action + ' detail 不含個資內容');
    ok(l.ip !== undefined && l.timestamp, e.action + ' 帶 ip／時間');
  }
  // 計數器：同一人同一天（跨端點累計；此人已匯出聯絡人名單與預測表各一次）
  await http('_s_u1', 'GET', '/api/export');
  const again = byAction(await getLogs(), 'EXPORT_CONTACTS')[0];
  ok(/今日第 3 次匯出/.test(again.detail), '每人每日（跨端點）匯出計數遞增：' + again.detail);
  // 被拒絕的匯出不應記成功匯出
  const n0 = byAction(await getLogs(), 'EXPORT_CONTACTS').length;
  let r = await http('_s_off', 'GET', '/api/export');
  ok(r.status === 403, '無下載權限者 403');
  eq(byAction(await getLogs(), 'EXPORT_CONTACTS').length, n0, '403 不寫匯出稽核');
  r = await http('_s_u1', 'GET', '/api/admin/opportunities/export');
  eq(r.status, 403, '非管理員打 admin 匯出 403');
  // 已有的牌價簿匯出仍有稽核（回歸）
  const pb0 = byAction(await getLogs(), 'EXPORT_QUOTE_PRICEBOOK').length;
  r = await http('_s_admin', 'GET', '/api/admin/quote-pricebook/export');
  if (r.status === 200) eq(byAction(await getLogs(), 'EXPORT_QUOTE_PRICEBOOK').length, pb0 + 1, '牌價簿匯出稽核維持');
  else console.log('  (牌價簿匯出 ' + r.status + '，略過回歸)');
  // 靜態防線：server.js 中每個 XLSX.write + 資料 Content-Disposition 的路由都必須呼叫 auditExport（範本除外）
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const routes = src.split(new RegExp(String.fromCharCode(10) + 'app[.](?=get|post)')).filter(b => b.includes('XLSX.write(wb') && b.includes('Content-Disposition'));
  const missing = routes.filter(b => !b.includes('auditExport(') && !/import-template|範本/.test(b.slice(0, 300) + b.slice(-300))).map(b => b.slice(0, 60));
  ok(routes.length >= 7, '靜態掃描有找到匯出路由（' + routes.length + '）');
  eq(missing, [], '每個產生 xlsx 下載的路由都有 auditExport（範本除外）');
}

// ═════════════ 第 3 項：管理員改派（只動目前關卡＋理由必填）═════════════
async function sec3() {
  console.log('[3] 管理員改派');
  await loginAll(['_s_admin', '_s_u1', '_q_m1']);
  const rq = (id, body, who) => http(who || '_s_admin', 'POST', '/api/quotations/' + id + '/reassign', body);
  const getQ = async (id) => (await http('_s_admin', 'GET', '/api/quotations/' + id)).json;
  const dataQ = (id) => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8')).quotations.find(q => q.id === id);
  await sleep(300);
  let r = await getQ('q-QB');
  ok(r && r.perm && r.perm.canReassign === true, 'QB（cur=1 的一級關）管理員 canReassign=true');
  eq((await getQ('q-QD')).perm.canReassign, false, '已核准的單 canReassign=false');

  // 理由必填
  r = await rq('q-QB', { username: '_q_m1b' });
  eq([r.status, r.json && r.json.code], [400, 'REASON_REQUIRED'], '沒帶理由 → 400 REASON_REQUIRED');
  r = await rq('q-QB', { username: '_q_m1b', reason: '   ' + String.fromCharCode(10, 9) + ' ' });
  eq([r.status, r.json && r.json.code], [400, 'REASON_REQUIRED'], '理由全空白 → 400');
  r = await rq('q-QB', { username: '_q_m1b', reason: '' });
  eq(r.status, 400, '空字串理由 → 400');
  eq(dataQ('q-QB').approval.steps[1].assignee, '_q_m1', '被拒絕時不改動');

  // 角色／身分驗證
  r = await rq('q-QB', { username: '_q_m2b', reason: '角色不符' });
  eq([r.status, r.json && r.json.code], [400, 'BAD_ASSIGNEE'], '一級關不能改派給二級主管');
  r = await rq('q-QB', { username: '_q_off', reason: '已停用' });
  eq([r.status, r.json && r.json.code], [400, 'BAD_ASSIGNEE'], '不能改派給停用帳號');
  r = await rq('q-QB', { username: '_s_u1', reason: '擁有者' });
  eq(r.status, 400, '不能改派給報價單擁有者');
  r = await rq('q-QB', { username: '_q_m1b', reason: '無權' }, '_s_u1');
  eq(r.status, 403, '非管理員 403');
  r = await rq('q-QD', { username: '_q_m1b', reason: '已核准' });
  eq([r.status, r.json && r.json.code], [409, 'CANNOT_REASSIGN'], '已核准的單 409');

  // 核心：QB（steps=[mgr2 已核准, mgr1 進行中], cur=1）→ 只動 steps[1]
  const before = JSON.parse(JSON.stringify(dataQ('q-QB').approval.steps[0]));
  r = await rq('q-QB', { username: '_q_m1b', reason: '原承辦人休假 <script>alert(1)</script>' });
  eq(r.status, 200, 'QB 改派成功');
  const qb = dataQ('q-QB');
  eq(qb.approval.steps[0], before, 'QB：steps[0]（已核准的二級主管關）完全沒被動到');
  eq(qb.approval.steps[1].assignee, '_q_m1b', 'QB：steps[1]（目前的一級關）換成新承辦人');
  eq(qb.approval.cur, 1, 'QB：cur 不變');
  const h = qb.approval.history[qb.approval.history.length - 1];
  eq(h.action, 'REASSIGN', 'QB：歷史多一筆 REASSIGN');
  eq(h.tier, 'mgr1', 'QB：歷史 tier 是實際關卡 mgr1');
  ok(h.comment.includes('理由：') && h.comment.includes('原承辦人休假'), 'QB：歷史 comment 含理由：' + h.comment);
  ok(!/[<>]/.test(h.comment) && !/[<>]/.test(h.meta.reason), 'QB：理由內的 < > 被清掉');
  eq([h.meta.from, h.meta.to, h.meta.step], ['_q_m1', '_q_m1b', 1], 'QB：meta {from,to,step}');
  ok(typeof h.meta.reason === 'string' && h.meta.reason.length > 0, 'QB：meta.reason');
  const logs = byAction(await getLogs(), 'REASSIGN_QUOTATION');
  ok(logs[0] && logs[0].target === 'SB-QB' && /一級主管關（第 2 關）/.test(logs[0].detail) && /_q_m1 → _q_m1b/.test(logs[0].detail) && /理由：/.test(logs[0].detail), '稽核含關卡、from→to、理由：' + (logs[0] && logs[0].detail));
  const apiH = ((await getQ('q-QB')).approval.history || []).slice(-1)[0];
  ok(apiH && apiH.action === 'REASSIGN' && !('meta' in apiH), 'API 回傳的歷程不含 meta');

  // 二級主管關進行中（QC）：只能改派給二級主管、改的是 steps[0]
  r = await rq('q-QC', { username: '_q_m1b', reason: '角色不符' });
  eq([r.status, r.json && r.json.code], [400, 'BAD_ASSIGNEE'], '二級關不能改派給一級主管');
  r = await rq('q-QC', { username: '_q_m2b', reason: '二級承辦人調動' });
  eq(r.status, 200, 'QC 改派二級主管成功');
  const qc = dataQ('q-QC');
  eq([qc.approval.steps[0].assignee, qc.approval.steps[1].assignee], ['_q_m2b', '_q_m1'], 'QC：只換目前這關（二級），一級關未動');
  eq(qc.approval.history.slice(-1)[0].tier, 'mgr2', 'QC：歷史 tier=mgr2');

  // 一般流程（QA）與同人改派去重（meta.from===to）
  r = await rq('q-QA', { username: '_q_m1b', reason: '工作分配' });
  eq(r.status, 200, 'QA 改派成功');
  eq(dataQ('q-QA').approval.steps[0].assignee, '_q_m1b', 'QA：steps[0] 換人');
  r = await rq('q-QA', { username: '_q_m1b', reason: '同一人再改派' });
  eq(r.status, 200, 'QA 改派給目前承辦人本人仍可');
  const ha = dataQ('q-QA').approval.history.slice(-1)[0];
  eq([ha.meta.from, ha.meta.to], ['_q_m1b', '_q_m1b'], '同人改派：meta.from===meta.to（信件模組靠它不重複寄）');
  r = await rq('q-QA', { username: '_q_m1', reason: 'x'.repeat(500) });
  eq(r.status, 200, '超長理由被截斷而非拒絕');
  eq(dataQ('q-QA').approval.history.slice(-1)[0].meta.reason.length, 200, '理由最長 200 字');
}

// ═════════════ 主程式 ═════════════
const SECTIONS = { 1: sec1, '1b': sec1b, 2: sec2, 3: sec3 };
(async () => {
  const only = (process.env.AW1_ONLY || '').split(',').map(s => s.trim()).filter(Boolean);
  seedFiles();
  try {
    await startServer();
    for (const k of Object.keys(SECTIONS)) {
      if (only.length && !only.includes(k)) continue;
      await restart();   // 每項一台新伺服器：登入速率限制（15 分鐘 20 次）是行程內記憶體，重啟歸零
      try { await SECTIONS[k](); } catch (e) { fail++; failures.push(`第 ${k} 項例外：${e && e.stack || e}`); console.log('  ✗ 例外', e && e.stack || e); }
    }
  } finally {
    await stopServer();
    cleanupFiles();
  }
  console.log(`\n結果：${pass} 通過，${fail} 失敗`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); process.exit(1); }
})().catch(async e => { console.error(e); await stopServer(); cleanupFiles(); process.exit(1); });
