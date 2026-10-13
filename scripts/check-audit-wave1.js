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
    mkUser('_s_slow', 'user', { bu: ['ERP'], password: bcrypt.hashSync(PW, 12) }),
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

// ═════════════ 第 4 項：管理操作的 before→after ═════════════
async function sec4() {
  console.log('[4] 管理操作 before→after');
  const AD = require(path.join(ROOT, 'lib', 'auditDiff.js'));
  // 單元：組字工具
  const many = Array.from({ length: 400 }, (_, i) => '項目' + i + '：' + 'x'.repeat(20));
  const capped = AD.joinCapped(many, { prefix: '前綴｜' });
  ok(capped.length <= AD.MAX_DETAIL + 40, 'joinCapped 長度受限（' + capped.length + '）');
  ok(/…\(截斷，另有 [0-9]+ 項未顯示\)/.test(capped), 'joinCapped 有截斷標記與未顯示數量');
  eq(AD.joinCapped(['a', 'b']), 'a；b', 'joinCapped 未超長不截斷');
  ok(!AD.scrub('x $2a$12$abcdefghijklmnopqrstuv y').includes('$2a$'), 'scrub 遮蔽 bcrypt 雜湊');
  eq(AD.fieldDiff({ a: 1, b: [1, 2] }, { a: 2, b: [1, 2] }, [{ key: 'a', label: 'A' }, { key: 'b', label: 'B' }]), ['A 1→2'], 'fieldDiff 只列有變的欄位');

  await loginAll(['_s_admin']);
  const lastLog = async (a) => (byAction(await getLogs(), a)[0] || {});
  // ── UPDATE_USER 欄位層級 ──
  let r = await adminPutUser('_s_dl', { role: 'manager2', bu: ['ERP', 'ITS'], accessMode: 'view', canDownloadContacts: false, canSetTargets: true, supervisor: '_s_m1', nickname: '沙盒暱稱', displayName: '沙盒顯示名', active: false });
  eq(r.status, 200, 'UPDATE_USER 多欄位成功：' + (r.json && r.json.error));
  let l = await lastLog('UPDATE_USER');
  const d = l.detail || '';
  ok(/角色 user→manager2/.test(d), '角色 before→after：' + d);
  ok(/BU ERP→ERP\+ITS/.test(d), 'BU before→after');
  ok(/存取模式 edit→view/.test(d), '存取模式 before→after');
  ok(/可下載客戶名單 是→否/.test(d), '下載權限 before→after');
  ok(/可設定目標 否→是/.test(d), '設定目標 before→after');
  ok(/直屬主管 ∅→_s_m1/.test(d), '直屬主管 before→after');
  ok(/暱稱 ∅→沙盒暱稱/.test(d), '暱稱 before→after');
  ok(/啟用 是→否/.test(d), '啟用 before→after');
  ok(/顯示名稱 _s_dl→沙盒顯示名/.test(d), '顯示名稱 before→after');
  eq(l.target, '_s_dl', 'UPDATE_USER target');
  // 獨立動作
  eq((await lastLog('DISABLE_USER')).target, '_s_dl', 'DISABLE_USER 獨立動作');
  eq((await lastLog('REVOKE_DOWNLOAD_PERMISSION')).target, '_s_dl', 'REVOKE_DOWNLOAD_PERMISSION 獨立動作');
  eq((await lastLog('SET_READONLY')).target, '_s_dl', 'SET_READONLY 獨立動作');
  ok(/user→manager2/.test((await lastLog('CHANGE_USER_ROLE')).detail || ''), 'CHANGE_USER_ROLE 獨立動作');
  r = await adminPutUser('_s_dl', { role: 'user', active: true, accessMode: 'edit', canDownloadContacts: true, bu: ['ERP'] });
  eq(r.status, 200, '還原');
  eq((await lastLog('ENABLE_USER')).target, '_s_dl', 'ENABLE_USER 獨立動作');
  eq((await lastLog('GRANT_DOWNLOAD_PERMISSION')).target, '_s_dl', 'GRANT_DOWNLOAD_PERMISSION 獨立動作');
  eq((await lastLog('SET_EDITABLE')).target, '_s_dl', 'SET_EDITABLE 獨立動作');
  // 沒有實際變動
  r = await adminPutUser('_s_dl', { canSetTargets: true });
  l = await lastLog('UPDATE_USER');
  ok(!/啟用|角色/.test(l.detail || '') || /可設定目標/.test(l.detail || '') === false, '無變動時不亂列欄位：' + l.detail);
  // Email 只記遮蔽訊息
  const em = 'secret.person@itts.com.tw';
  r = await adminPutUser('_s_dl', { email: em });
  if (r.status === 200) {
    l = await lastLog('UPDATE_USER');
    ok(/email 未設定→已設定/.test(l.detail) && !l.detail.includes('secret.person'), 'Email 只記遮蔽訊息，不含位址');
  } else console.log('  (測試網域被 Email 白名單擋下 ' + r.status + '，略過 Email 正向案例)');

  // ── 角色功能權限矩陣 ──
  const rp = (await http('_s_admin', 'GET', '/api/admin/role-permissions')).json;
  const matrix = JSON.parse(JSON.stringify(rp.matrix));
  const victim = matrix.user.find(k => k !== 'home') || matrix.user[0];
  const removed = matrix.user.filter(k => k === victim);
  matrix.user = matrix.user.filter(k => k !== victim);
  const adding = rp.features.map(f => f.key).find(k => !(rp.matrix.manager1 || []).includes(k));
  if (adding) matrix.manager1 = (rp.matrix.manager1 || []).concat([adding]);
  r = await http('_s_admin', 'PUT', '/api/admin/role-permissions', { matrix });
  eq(r.status, 200, '更新權限矩陣');
  l = await lastLog('UPDATE_ROLE_PERMISSIONS');
  const lbl = (k) => ((rp.features.find(f => f.key === k) || {}).label) || k;
  ok(l.detail.includes('user 移除[' + lbl(victim) + ']'), '權限矩陣：user 移除功能（中文）：' + l.detail);
  if (adding) ok(l.detail.includes('manager1 新增[' + lbl(adding) + ']'), '權限矩陣：manager1 新增功能');
  r = await http('_s_admin', 'PUT', '/api/admin/role-permissions', { matrix });
  ok(/無實際異動/.test((await lastLog('UPDATE_ROLE_PERMISSIONS')).detail), '權限矩陣：沒變動時註明無實際異動');
  r = await http('_s_admin', 'DELETE', '/api/admin/role-permissions');
  eq(r.status, 200, '重置權限矩陣');
  ok(/異動：|無實際異動/.test((await lastLog('RESET_ROLE_PERMISSIONS')).detail), '重置也記錄影響角色');

  // ── 簽核設定：商品歸類 before→after ──
  const qc = (await http('_s_admin', 'GET', '/api/quote-approval/config')).json;
  const cls = Object.keys(qc.classLabels)[0], cls2 = Object.keys(qc.classLabels)[1] || cls;
  const putCfg = (pc) => http('_s_admin', 'PUT', '/api/admin/quote-approval/config', { roster: qc.roster, productClasses: pc });
  r = await putCfg({ '沙盒商品X': { cls, costBySales: false } });
  eq(r.status, 200, '簽核設定新增商品歸類：' + (r.json && r.json.error));
  ok(/新增「沙盒商品X」/.test((await lastLog('UPDATE_QUOTE_APPROVAL_CONFIG')).detail), '商品歸類：新增');
  r = await putCfg({ '沙盒商品X': { cls, costBySales: true } });
  ok(/「沙盒商品X」業務填成本 否→是/.test((await lastLog('UPDATE_QUOTE_APPROVAL_CONFIG')).detail), '商品歸類：costBySales before→after：' + (await lastLog('UPDATE_QUOTE_APPROVAL_CONFIG')).detail);
  if (cls2 !== cls) { r = await putCfg({ '沙盒商品X': { cls: cls2, costBySales: true } }); ok(/類別 .+→.+/.test((await lastLog('UPDATE_QUOTE_APPROVAL_CONFIG')).detail), '商品歸類：類別 before→after'); }
  r = await putCfg({});
  ok(/移除「沙盒商品X」/.test((await lastLog('UPDATE_QUOTE_APPROVAL_CONFIG')).detail), '商品歸類：移除');
  const big = {}; for (let i = 0; i < 300; i++) big['沙盒大量商品' + i + '_' + 'y'.repeat(30)] = { cls, costBySales: i % 2 === 0 };
  r = await putCfg(big);
  eq(r.status, 200, '大量商品歸類可儲存');
  l = await lastLog('UPDATE_QUOTE_APPROVAL_CONFIG');
  ok(l.detail.length <= AD.MAX_DETAIL + 60 && /截斷，另有 [0-9]+ 項未顯示/.test(l.detail), '大量異動：長度受限並標示截斷（' + l.detail.length + '）');
  await putCfg({});

  // ── 流程設定 ──
  const fc = (await http('_s_admin', 'GET', '/api/admin/quote-approval/flow-config')).json;
  const fcBody = JSON.parse(JSON.stringify(fc));
  fcBody.buSettings.ERP.hasMgr2Step = !fcBody.buSettings.ERP.hasMgr2Step;
  fcBody.exceptions = (fcBody.exceptions || []).concat([{ username: '_q_m2', bu: 'ERP', role: 'manager2', action: 'skip_mgr1', label: '沙盒例外' }]);
  r = await http('_s_admin', 'PUT', '/api/admin/quote-approval/flow-config', fcBody);
  eq(r.status, 200, '流程設定儲存');
  l = await lastLog('SAVE_APPROVAL_FLOW_CONFIG');
  ok(/ERP 二級主管關 (啟用→停用|停用→啟用)/.test(l.detail), '流程設定：BU 開關 before→after：' + l.detail);
  ok(/新增例外 skip_mgr1：_q_m2/.test(l.detail), '流程設定：例外新增（帳號）');
  fcBody.exceptions = [];
  fcBody.buSettings.ERP.hasMgr2Step = fc.buSettings.ERP.hasMgr2Step;
  await http('_s_admin', 'PUT', '/api/admin/quote-approval/flow-config', fcBody);
  ok(/移除例外 skip_mgr1：_q_m2/.test((await lastLog('SAVE_APPROVAL_FLOW_CONFIG')).detail), '流程設定：例外移除');

  // ── 整合設定（不含機敏）──
  const SECRET = 'Sentinel' + crypto.randomBytes(6).toString('hex') + '!9';
  r = await http('_s_admin', 'POST', '/api/admin/integrations', { type: 'sap-sales-cloud-v2', name: '沙盒整合', baseUrl: 'https://sandbox-a.invalid', username: 'sb_user_a', password: SECRET, clientSecret: SECRET + 'c' });
  eq(r.status, 200, '建立整合：' + (r.json && r.json.error));
  r = await http('_s_admin', 'POST', '/api/admin/integrations', { type: 'sap-sales-cloud-v2', baseUrl: 'https://sandbox-b.invalid', username: 'sb_user_b', password: SECRET + '2' });
  eq(r.status, 200, '修改整合');
  l = await lastLog('SAVE_INTEGRATION');
  ok(/baseUrl https:\/\/sandbox-a\.invalid→https:\/\/sandbox-b\.invalid/.test(l.detail), '整合：baseUrl before→after：' + l.detail);
  ok(/使用者名稱 sb_user_a→sb_user_b/.test(l.detail), '整合：username before→after');
  ok(/password 已更新（值不記錄）/.test(l.detail), '整合：密碼只標示已更新');
  // ── 全域：任何稽核資料都不含密碼／雜湊／測試密碼／哨兵 ──
  await adminPutUser('_s_dl', { displayName: '沙盒顯示名2' });
  await http('_s_admin', 'PUT', '/api/admin/users/_s_dl/password', { password: 'Zz9' + crypto.randomBytes(8).toString('hex') + '!A' });
  const all = JSON.stringify(await getLogs());
  ok(!all.includes('$2a$') && !all.includes('$2b$'), '稽核全文不含 bcrypt 雜湊');
  ok(!all.includes(PW) && !all.includes(SECRET), '稽核全文不含測試密碼／整合密碼');
  ok(!/"detail":"[^"]*password[^"]*"/i.test(all.replace(/password 已更新（值不記錄）/g, '').replace(/重設密碼/g, '')), '稽核 detail 不含 password 字樣（除遮蔽說明外）');
}

// ═════════════ 第 5 項：登入失敗／登出稽核、統一失敗訊息 ═════════════
async function sec5() {
  console.log('[5] 登入失敗／登出稽核');
  const UA = 'AuditWave1-Test/1.0';
  const attempt = async (user, pw) => {
    const t = Date.now();
    const res = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json', 'User-Agent': UA }, body: JSON.stringify({ username: user, password: pw }) });
    const text = await res.text();
    return { status: res.status, text, ms: Date.now() - t };
  };
  await loginAll(['_s_admin']);
  // 統一回應：不存在／密碼錯／停用(密碼錯)／停用(密碼對)
  await adminPutUser('_s_off', { active: false });
  const rUnknown = await attempt('_s_nobody', 'wrong-pw-1');
  const rBad = await attempt('_s_slow', 'wrong-pw-2');
  const rInactiveBad = await attempt('_s_off', 'wrong-pw-3');
  const rInactiveOk = await attempt('_s_off', PW);
  for (const [n, r] of [['不存在', rUnknown], ['密碼錯', rBad], ['停用+密碼錯', rInactiveBad], ['停用+密碼對', rInactiveOk]]) {
    eq(r.status, 401, n + '：HTTP 401');
  }
  ok(new Set([rUnknown.text, rBad.text, rInactiveBad.text, rInactiveOk.text]).size === 1, '四種失敗回應內容完全相同：' + rUnknown.text);
  ok(!/UNKNOWN|INACTIVE|停用|不存在/.test(rUnknown.text + rInactiveOk.text), '回應不洩漏原因');
  // 時間：不存在帳號也做了等成本的 bcrypt（>=80ms），且與「密碼錯（成本 12）」同量級
  const t1 = []; const t2 = [];
  for (let i = 0; i < 3; i++) { t1.push((await attempt('_s_nobody' + i, 'x')).ms); t2.push((await attempt('_s_slow', 'x' + i)).ms); }
  t1.sort((a, b) => a - b); t2.sort((a, b) => a - b);
  console.log(`  (不存在 ${t1[1]}ms／密碼錯 ${t2[1]}ms)`);
  ok(t1[1] >= 60, '不存在帳號的回應時間 >= 60ms（有做 dummy bcrypt）');
  ok(Math.abs(t1[1] - t2[1]) < Math.max(t1[1], t2[1]) * 0.7, '不存在與密碼錯的回應時間同量級');
  // 稽核記錄
  let logs = byAction(await getLogs(), 'LOGIN_FAIL');
  const reasonOf = (name) => (logs.find(l => l.target === name) || {}).detail || '';
  ok(/原因=UNKNOWN_USER/.test(reasonOf('_s_nobody')), 'LOGIN_FAIL 原因 UNKNOWN_USER');
  ok(/原因=BAD_PASSWORD/.test(reasonOf('_s_slow')), 'LOGIN_FAIL 原因 BAD_PASSWORD');
  const offs = logs.filter(l => l.target === '_s_off').map(l => l.detail);
  ok(offs.some(d => /原因=INACTIVE\b/.test(d)) && offs.some(d => /原因=INACTIVE_CORRECT_PASSWORD/.test(d)), 'LOGIN_FAIL 區分 INACTIVE 與 INACTIVE_CORRECT_PASSWORD（停用帳號用對密碼）：' + offs.join(' | '));
  const sample = logs.find(l => l.target === '_s_nobody') || {};
  ok(sample.ip && sample.userAgent === UA && sample.timestamp, 'LOGIN_FAIL 帶 IP／User-Agent／時間');
  // 長帳號字串截斷 + 清掉 < >
  const long = '<img src=x onerror=alert(1)>' + 'A'.repeat(200);
  await attempt(long, 'p');
  logs = byAction(await getLogs(), 'LOGIN_FAIL');
  const longRow = logs.find(l => l.target.includes('img src=x')) || {};
  ok(longRow.target && longRow.target.length <= 40 && !/[<>]/.test(longRow.target), '帳號字串截斷 40 字且去掉 < >：' + longRow.target);
  ok(!/wrong-pw/.test(JSON.stringify(logs)) && !JSON.stringify(logs).includes(PW), 'LOGIN_FAIL 不記錄嘗試的密碼');
  // 節流：同 IP 同帳號 30 秒內只一筆
  const before = byAction(await getLogs(), 'LOGIN_FAIL').filter(l => l.target === '_s_thr').length;
  for (let i = 0; i < 5; i++) await attempt('_s_thr', 'bad' + i);
  await attempt('_s_thr2', 'bad');
  logs = byAction(await getLogs(), 'LOGIN_FAIL');
  eq(logs.filter(l => l.target === '_s_thr').length - before, 1, '同 IP 同帳號 5 次失敗只記 1 筆');
  eq(logs.filter(l => l.target === '_s_thr2').length, 1, '不同帳號各自記錄');
  // 成功登入與登出
  let r = await login('_s_u1');
  eq(r.status, 200, '成功登入回應不變');
  ok(r.json && r.json.success === true && r.json.role === 'user', '成功登入回應欄位');
  ok(byAction(await getLogs(), 'LOGIN').some(l => l.operator === '_s_u1'), '成功登入仍有 LOGIN 稽核');
  const lo0 = byAction(await getLogs(), 'LOGOUT').length;
  r = await http('_s_u1', 'POST', '/api/logout');
  eq(r.status, 200, '登出成功');
  logs = byAction(await getLogs(), 'LOGOUT');
  eq(logs.length - lo0, 1, '登出寫入 LOGOUT');
  eq(logs[0].operator, '_s_u1', 'LOGOUT operator');
  r = await http(null, 'POST', '/api/logout');
  eq(byAction(await getLogs(), 'LOGOUT').length, lo0 + 1, '未登入呼叫登出不產生 LOGOUT');
  r = await http('_s_u1', 'GET', '/api/me');
  eq(r.status === 401 || r.status === 200, true, '(登出後舊 cookie 行為記錄)');
  // 30 秒窗後的下一筆要註明合併次數
  console.log('  (等 31 秒驗證合併計數…)');
  await sleep(31000);
  await attempt('_s_thr', 'bad-after-window');
  logs = byAction(await getLogs(), 'LOGIN_FAIL').filter(l => l.target === '_s_thr');
  ok(/另有 4 次相同失敗已合併/.test(logs[0].detail), '窗後的下一筆註明合併 4 次：' + logs[0].detail);
  // 限流仍在：持續失敗最終 429，且留 LOGIN_RATE_LIMITED
  let got429 = false;
  for (let i = 0; i < 30 && !got429; i++) { const a = await attempt('_s_rl' + i, 'x'); if (a.status === 429) got429 = true; }
  ok(got429, '登入速率限制仍生效（429）');
  ok(byAction(await getLogs(), 'LOGIN_RATE_LIMITED').length >= 1, '限流觸發有 LOGIN_RATE_LIMITED 稽核');
}

// ═════════════ 第 6 項：管理員安全護欄 ═════════════
async function sec6() {
  console.log('[6] 管理員安全護欄');
  await loginAll(['_s_admin', '_s_admin2']);
  const blocked = async (who) => { const r = await http(who || '_s_admin', 'GET', '/api/admin/logs'); return (Array.isArray(r.json) ? r.json : []).filter(l => l.action === 'ADMIN_RAIL_BLOCKED'); };
  const me = async (who) => (await http(who, 'GET', '/api/me')).json;
  let r;
  // a. 自己改自己的角色／啟用 → 403
  const b0 = (await blocked()).length;
  r = await adminPutUser('_s_admin', { role: 'user' });
  eq([r.status, r.json && r.json.code], [403, 'ADMIN_SELF_CHANGE_FORBIDDEN'], '管理員降自己的角色 → 403 ADMIN_SELF_CHANGE_FORBIDDEN');
  ok(/管理員不能變更自己/.test(r.json && r.json.error), '訊息是清楚的中文');
  eq((await me('_s_admin')).role, 'admin', '角色確實沒變');
  r = await adminPutUser('_s_admin', { active: false });
  eq([r.status, r.json && r.json.code], [403, 'ADMIN_SELF_CHANGE_FORBIDDEN'], '管理員停用自己 → 403');
  r = await adminPutUser('_s_admin', { role: 'manager1', bu: ['ERP'] });
  eq(r.status, 403, '管理員把自己改成別的角色 → 403');
  let rows = await blocked();
  eq(rows.length - b0, 3, '三次被擋都有 ADMIN_RAIL_BLOCKED 稽核');
  ok(/ADMIN_SELF_CHANGE_FORBIDDEN/.test(rows[0].detail) && rows[0].operator === '_s_admin' && rows[0].target === '_s_admin', '稽核含規則代碼、操作者、目標：' + rows[0].detail);
  // c. 不改敏感欄位的自我編輯仍可（包含帶著「沒變」的完整 payload、admin 的 accessMode 恆為 edit）
  r = await adminPutUser('_s_admin', { displayName: '沙盒管理員', role: 'admin', active: true, accessMode: 'view' });
  eq(r.status, 200, '自己改顯示名稱（role/active 沒變、admin accessMode 恆 edit）仍可：' + (r.json && r.json.error));
  eq(rows.length, (await blocked()).length, '沒變的欄位不觸發護欄');
  // d. 另一位管理員可以做（需另一位管理員操作）
  r = await adminPutUser('_s_admin', { accessMode: 'edit' }, '_s_admin2');
  eq(r.status, 200, '其他管理員改 admin 帳號仍可');
  r = await adminPutUser('_s_admin', { role: 'user', bu: ['ERP'] }, '_s_admin2');
  eq(r.status, 200, '另一位管理員可降級管理員（不是最後一位）');
  r = await http('_s_admin', 'GET', '/api/admin/logs');
  eq(r.status, 403, '被降級者立即失去後台權限');
  r = await adminPutUser('_s_admin', { role: 'admin' }, '_s_admin2');
  eq(r.status, 200, '另一位管理員可復原');
  r = await adminPutUser('_s_admin', { active: false }, '_s_admin2');
  eq(r.status, 200, '另一位管理員可停用管理員（不是最後一位）');
  r = await adminPutUser('_s_admin', { active: true }, '_s_admin2');
  eq(r.status, 200, '復原啟用');
  // e. 刪除：非最後一位可刪、最後一位不可（含自刪）
  r = await http('_s_admin2', 'DELETE', '/api/admin/users/_s_admin');
  eq(r.status, 200, '刪除另一位管理員（還剩自己）');
  // f. 現在 _s_admin2 是唯一啟用中的管理員
  r = await adminPutUser('_s_admin2', { role: 'user', bu: ['ERP'] }, '_s_admin2');
  eq([r.status, r.json && r.json.code], [400, 'LAST_ADMIN_PROTECTED'], '最後一位管理員自我降級 → 400 LAST_ADMIN_PROTECTED');
  r = await adminPutUser('_s_admin2', { active: false }, '_s_admin2');
  eq([r.status, r.json && r.json.code], [400, 'LAST_ADMIN_PROTECTED'], '最後一位管理員自我停用 → 400 LAST_ADMIN_PROTECTED');
  r = await http('_s_admin2', 'DELETE', '/api/admin/users/_s_admin2');
  eq([r.status, r.json && r.json.code], [400, 'LAST_ADMIN_PROTECTED'], '刪除最後一位管理員 → 400（原本就擋，現在多了 code 與稽核）');
  eq((await me('_s_admin2')).role, 'admin', '最後一位管理員安然無恙');
  rows = await blocked('_s_admin2');
  ok(rows.filter(l => /LAST_ADMIN_PROTECTED/.test(l.detail)).length >= 3, 'LAST_ADMIN_PROTECTED 被擋三次都有稽核');
  // g. 一般使用者的編輯不受影響
  r = await adminPutUser('_s_dl', { displayName: '一般帳號可正常編輯', role: 'user' }, '_s_admin2');
  eq(r.status, 200, '編輯一般使用者不受護欄影響');
}

// ═════════════ 變異檢查：故意弄壞規則，測試必須抓得到 ═════════════
// 做法：把 server.js／lib/quoteRoutes.js 複製成臨時檔並改壞一條規則，用同一套 HTTP 斷言打變異伺服器；
// 「抓到」＝斷言失敗（與正常版相反）。臨時檔結束後一律刪除。
async function secM() {
  console.log('[M] 變異檢查');
  const tmpFiles = [];
  const mkMut = (name, serverEdit, libEdit) => {
    let srv = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
    if (serverEdit) { if (!srv.includes(serverEdit[0])) throw new Error('變異樣式找不到：' + name); srv = srv.replace(serverEdit[0], () => serverEdit[1]); }
    if (libEdit) {
      let lib = fs.readFileSync(path.join(ROOT, 'lib', 'quoteRoutes.js'), 'utf8');
      if (!lib.includes(libEdit[0])) throw new Error('變異樣式找不到：' + name);
      lib = lib.replace(libEdit[0], () => libEdit[1]);
      const libFile = path.join(ROOT, 'lib', 'quoteRoutes.__mut_' + name + '.js');
      fs.writeFileSync(libFile, lib); tmpFiles.push(libFile);
      srv = srv.replace("require('./lib/quoteRoutes')", () => "require('./lib/quoteRoutes.__mut_" + name + ".js')");
    }
    const f = path.join(ROOT, 'server.__mut_' + name + '.js');
    fs.writeFileSync(f, srv); tmpFiles.push(f);
    return f;
  };
  const fresh = async (file) => { await stopServer(); cleanupFiles(); seedFiles(); extraEnv = {}; await startServer(file); jar = {}; };
  try {
    // 變異 1：拿掉「停用帳號即時拒絕」
    {
      const f = mkMut('m1', ["const bad = !acct ? '帳號已不存在' : (acct.active === false ? '帳號已停用' : null);", 'const bad = null;']);
      await fresh(f);
      await loginAll(['_s_admin', '_s_off']);
      await adminPutUser('_s_off', { active: false });
      const r = await http('_s_off', 'GET', '/api/me');
      ok(r.status !== 401, '變異1（拿掉停用即時拒絕）→ 停用後仍 ' + r.status + '：測試 sec1 的 401 斷言會失敗（抓得到）');
    }
    // 變異 2：改派又改回「永遠改 steps[0]」
    {
      const f = mkMut('m2', null, ['const from = curStep.assignee;\n    curStep.assignee = un;', 'const from = ap.steps[0].assignee;\n    ap.steps[0].assignee = un;']);
      await fresh(f);
      await loginAll(['_s_admin']);
      const r = await http('_s_admin', 'POST', '/api/quotations/q-QB/reassign', { username: '_q_m1b', reason: '變異測試' });
      const q = JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8')).quotations.find(x => x.id === 'q-QB');
      ok(r.status === 200 && q.approval.steps[0].assignee === '_q_m1b' && q.approval.steps[1].assignee === '_q_m1', '變異2（改派永遠改 steps[0]）→ 已核准的二級關被改掉：sec3 的 steps[0] 不變斷言會失敗（抓得到）');
    }
    // 變異 3：拿掉「管理員不能改自己的角色」
    {
      const f = mkMut('m3', ["if (req.params.username === opName0 && changes.length) {", 'if (false) {']);
      await fresh(f);
      await loginAll(['_s_admin']);
      const r = await adminPutUser('_s_admin', { role: 'user', bu: ['ERP'] });
      ok(r.status === 200, '變異3（拿掉自我變更護欄）→ 自己降自己成功(' + r.status + ')：sec6 的 403 斷言會失敗（抓得到）');
    }
    // 變異 4：匯出不寫稽核
    {
      const f = mkMut('m4', ["auditExport(req, 'EXPORT_VISITS', rows.length,", "void (req, 'EXPORT_VISITS', rows.length,"]);
      await fresh(f);
      await loginAll(['_s_admin']);
      await http('_s_admin', 'GET', '/api/admin/visits/export');
      const n = byAction(await getLogs(), 'EXPORT_VISITS').length;
      eq(n, 0, '變異4（拿掉拜訪匯出稽核）→ 沒有 EXPORT_VISITS 紀錄：sec2 的「寫入一筆」斷言會失敗（抓得到）');
    }
  } finally {
    await stopServer();
    tmpFiles.forEach(f => { try { fs.unlinkSync(f); } catch { /* 已刪 */ } });
  }
}

// ═════════════ 主程式 ═════════════
const SECTIONS = { 1: sec1, '1b': sec1b, 2: sec2, 3: sec3, 4: sec4, 5: sec5, 6: sec6, M: secM };
(async () => {
  const only = (process.env.AW1_ONLY || '').split(',').map(s => s.trim()).filter(Boolean);
  seedFiles();
  try {
    await startServer();
    for (const k of ['1', '1b', '2', '3', '4', '5', '6', 'M']) {
      if (!SECTIONS[k] || (k === 'M' && process.env.AW1_SKIP_MUT)) continue;
      if (only.length && !only.includes(k)) continue;
      await stopServer(); cleanupFiles(); seedFiles(); await restart();   // 每項一台新伺服器：登入速率限制（15 分鐘 20 次）是行程內記憶體，重啟歸零
      try { await SECTIONS[k](); } catch (e) { fail++; failures.push(`第 ${k} 項例外：${e && e.stack || e}`); console.log('  ✗ 例外', e && e.stack || e); }
    }
  } finally {
    await stopServer();
    cleanupFiles();
  }
  console.log(`\n結果：${pass} 通過，${fail} 失敗`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); process.exit(1); }
})().catch(async e => { console.error(e); await stopServer(); cleanupFiles(); process.exit(1); });
