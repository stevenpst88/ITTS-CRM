#!/usr/bin/env node
'use strict';
/**
 * 報價單權限類型（quoteMode：簽核／唯讀 二擇一，職能分工）端對端檢查。不含任何帳密：測試帳號與隨機密碼只在記憶體裡。
 *
 * 做法：在「本工作樹」建立臨時 auth.json / data.json（假資料 + _s_* 臨時帳號），自行以 PORT=QM_PORT（預設 3161）
 * 啟動 server.js，跑完後只用 PID 關閉並刪除臨時檔。
 * 守門：工作樹若已有 auth.json / data.json / audit.log.json（可能是真實資料），一律拒絕執行。
 * 禁用埠：3000 / 3020 / 3030 / 3101 / 3111 / 3121 / 3131 / 3141 / 3151（其他 session 在用）。
 *
 *   node scripts/check-quote-mode.js               # 全部
 *   QM_ONLY=1,2 node scripts/check-quote-mode.js   # 只跑指定項目（1 帳號驗證／2 伺服器端執法／3 建單角色不受影響＋舊帳號＋報告／4 靜態／M 變異）
 *   QM_SKIP_MUT=1 ...                              # 略過變異檢查
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.QM_PORT || '3161', 10);
if ([3000, 3020, 3030, 3101, 3111, 3121, 3131, 3141, 3151].includes(PORT)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));
const guardFiles = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
if (guardFiles.length) { console.log('拒絕：此工作樹已有', guardFiles.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));

// ───────── 沙盒資料 ─────────
const PW = 'Qm1' + crypto.randomBytes(9).toString('hex') + '!Z';       // 只在記憶體
const NEWPW = 'Qn2' + crypto.randomBytes(9).toString('hex') + '!Y';    // 新增帳號用，只在記憶體
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
const mkOpp = (owner) => ({ id: 'o-' + owner, title: '沙盒商機 ' + owner, company: '沙盒公司甲', owner, category: 'ERP', stage: 'B', amount: 100, expectedDate: new Date().getFullYear() + '-12-01', createdAt: nowIso });
function seedFiles() {
  const users = [
    mkUser('_s_admin', 'admin'), mkUser('_s_admin2', 'admin'),
    mkUser('_s_m1', 'manager1', { bu: ['ERP'] }),                       // 一級主管（舊帳號，沒有 quoteMode）
    mkUser('_s_m1b', 'manager1', { bu: ['ERP'] }),                      // 另一位一級主管（舊帳號）
    mkUser('_s_m2', 'manager2', { bu: ['ERP'], supervisor: '_s_m1' }),  // 二級主管：建單角色
    mkUser('_s_s1', 'user', { bu: ['ERP'], supervisor: '_s_m2' }),      // 業務：直屬二級主管（ERP 啟用二級主管關）
    mkUser('_s_s2', 'user', { bu: ['ERP'], supervisor: '_s_m1' }),      // 業務：直屬一級主管 m1
    mkUser('_s_s3', 'user', { bu: ['ERP'], supervisor: '_s_m1b' }),     // 業務：直屬一級主管 m1b
    mkUser('_s_gma', 'executive', {}), mkUser('_s_gmb', 'executive', {}),   // 兩位總經理（都在 gm 名冊，舊帳號）
    mkUser('_s_ch', 'executive', {}),                                   // 董事長
    mkUser('_s_sec1', 'secretary', { bu: ['ERP', 'ITS', 'MDM', 'CRM'] }), mkUser('_s_sec2', 'secretary', { bu: ['ERP', 'ITS', 'MDM', 'CRM'] }), mkUser('_s_sec3', 'secretary', { bu: ['ERP', 'ITS', 'MDM', 'CRM'] }),
    mkUser('_s_sec4', 'secretary', { bu: ['ITS'] }),   // 只掛 ITS 的秘書：看不到 ERP 業務的單，除非輪到董事會代核關
    mkUser('_s_cm', 'consult_manager_south', { bu: ['ERP'] }),          // 顧問主管：成本填寫人
    mkUser('_s_mkt', 'marketing', { bu: ['ERP'] }),                     // 行銷：預設沒有報價單功能
    mkUser('_s_acct', 'accounting_manager', {}),
    mkUser('_s_off', 'secretary', { bu: ['ERP'], active: false }),
  ];
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  const d = {
    contacts: [], companies: [{ id: 'co1', name: '沙盒公司甲', taxId: '12345675', industry: '製造業' }],
    opportunities: ['_s_s1', '_s_s2', '_s_s3', '_s_m2'].map(mkOpp),
    visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [],
    quotations: [{ id: 'q-BD', quoteNo: 'SB-BD', company: '沙盒公司甲', owner: '_s_s2', status: 'draft', createdAt: nowIso, updatedAt: nowIso, products: ['PS'], items: [], approval: { state: 'pending', rulesVersion: 1, hash: 'h', submittedAt: nowIso, submittedBy: '_s_s2', steps: [{ tier: 'board', label: '董事會', assignee: null, status: 'pending' }], cur: 0, board: null, derived: null, history: [] }, costFlow: { state: 'na' } }],
    quoteApproval: { roster: { gm: ['_s_gma', '_s_gmb'], chairman: ['_s_ch'], boardProxy: [], costProviders: ['_s_cm'], sealManagers: [] }, productClasses: { PS: { cls: 'software', costBySales: true } } },
    approvalFlowConfig: { version: 1, buSettings: { ERP: { hasMgr2Step: true } }, exceptions: [] },
  };
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
}
function cleanupFiles() { FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }
const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
const readAuth = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'auth.json'), 'utf8'));

// ───────── 伺服器 ─────────
let child = null, childPid = null;
let extraEnv = {};
async function startServer(serverFile) {
  const env = Object.assign({}, process.env, extraEnv, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  delete env.DATABASE_URL; delete env.MAIL_MODE;
  child = spawn(process.execPath, [serverFile || path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  const onData = d => { out += d; if (process.env.QM_DEBUG) process.stdout.write('[srv] ' + d); };
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
async function http(who, method, url, body) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  const sc = res.headers.getSetCookie ? res.headers.getSetCookie() : [];
  const fresh = sc.find(c => /^itts_auth=[^;]+/.test(c));
  if (who && fresh) jar[who] = fresh.split(';')[0];
  return { status: res.status, json, text: buf.toString('utf8'), setCookie: sc };
}
async function login(who, user, pw) {
  const r = await http(null, 'POST', '/api/login', { username: user || who, password: pw || PW });
  if (r.status === 200) jar[who] = r.setCookie.map(c => c.split(';')[0]).join('; ');
  return r;
}
async function loginAll(names) { for (const n of names) { const r = await login(n); if (r.status !== 200) throw new Error('登入失敗 ' + n + ' ' + r.status); } }
const getLogs = async () => { const r = await http('_s_admin', 'GET', '/api/admin/logs'); return Array.isArray(r.json) ? r.json : []; };
const byAction = (logs, a) => logs.filter(l => l.action === a);
const putUser = (u, body, who) => http(who || '_s_admin', 'PUT', '/api/admin/users/' + encodeURIComponent(u), body);
const setMode = (u, mode) => putUser(u, { quoteMode: mode });
const adminUsers = async () => { const r = await http('_s_admin', 'GET', '/api/admin/users'); return Array.isArray(r.json) ? r.json : []; };
const userRow = async (u) => (await adminUsers()).find(x => x.username === u) || {};
const getCfg = async () => (await http('_s_admin', 'GET', '/api/quote-approval/config')).json;
const putRoster = async (patch) => { const c = await getCfg(); const roster = Object.assign({}, c.roster, patch); return http('_s_admin', 'PUT', '/api/admin/quote-approval/config', { roster, productClasses: c.productClasses }); };

// 報價單流程輔助
async function mkQuote(who, price, cost, extra) {
  const r = await http(who, 'POST', '/api/quotations', Object.assign({ oppId: 'o-' + who, company: '沙盒公司甲', products: ['PS'], items: [{ desc: '品項A', unit: '式', qty: 1, unitPrice: price, cost }] }, extra || {}));
  return r;
}
const submitQ = (who, id) => http(who, 'POST', `/api/quotations/${id}/submit`);
const getQ = (who, id) => http(who, 'GET', `/api/quotations/${id}`);
async function actQ(who, id, act, body) {
  const g = await getQ(who, id);
  const hash = g.json && g.json.contentHash;
  return http(who, 'POST', `/api/quotations/${id}/${act}`, Object.assign({ hash, comment: act === 'return' ? '沙盒駁回' : '' }, body || {}));
}
const inboxIds = async (who) => { const r = await http(who, 'GET', '/api/quotations/inbox'); return ((r.json && r.json.approvals) || []).map(x => x.id); };
const notesOf = async (who) => { const r = await http(who, 'GET', '/api/notifications'); return Array.isArray(r.json) ? r.json : []; };
const stepsOf = (j) => ((j && j.approval && j.approval.steps) || []).map(s => s.tier);
const assigneeOf = (id, i) => (((readData().quotations || []).find(x => x.id === id) || {}).approval || { steps: [{}] }).steps[i || 0].assignee;
const HI = [1000, 100], LOW = [1000, 930];   // 毛利 90%（只需一級主管）／7%（軟體列：一級主管→總經理）

// ═════════════ 第 1 項：帳號驗證（新增／編輯／角色衝突／自我護欄／稽核／通知） ═════════════
async function sec1() {
  console.log('[1] 帳號驗證');
  await loginAll(['_s_admin', '_s_admin2', '_s_sec2', '_s_m2', '_s_s1']);
  const newUser = (username, role, extra) => http('_s_admin', 'POST', '/api/admin/users', Object.assign({ username, password: NEWPW, displayName: username, role, bu: ['ERP'] }, extra || {}));
  // 必填
  let r = await newUser('_s_n1', 'secretary');
  eq([r.status, r.json && r.json.code], [400, 'QUOTE_MODE_REQUIRED'], '新增非建單角色（秘書，有報價單功能）沒給類型 → 400 QUOTE_MODE_REQUIRED');
  ok(!readAuth().users.some(u => u.username === '_s_n1'), '被拒絕的新增沒有留下帳號');
  r = await newUser('_s_n1', 'manager1', { quoteMode: 'boss' });
  eq([r.status, r.json && r.json.code], [400, 'BAD_QUOTE_MODE'], '非法值 → 400 BAD_QUOTE_MODE');
  r = await newUser('_s_n1', 'manager1', { quoteMode: ['approver'] });
  eq([r.status, r.json && r.json.code], [400, 'BAD_QUOTE_MODE'], '陣列值 → 400 BAD_QUOTE_MODE');
  r = await newUser('_s_n1', 'secretary', { quoteMode: 'readonly' });
  eq(r.status, 200, '新增秘書＋唯讀 → 200');
  eq((await userRow('_s_n1')).quoteMode, 'readonly', '帳號列表回傳 quoteMode=readonly');
  ok(readAuth().users.find(u => u.username === '_s_n1').quoteMode === 'readonly', 'auth 記錄存了 quoteMode 欄位');
  r = await newUser('_s_n2', 'executive', { quoteMode: 'approver' });
  eq(r.status, 200, '新增董事長/總經理＋簽核 → 200');
  r = await newUser('_s_n3', 'user', { quoteMode: 'approver' });
  eq([r.status, !!(r.json && r.json.notice)], [200, true], '建單角色（業務）送簽核 → 成功但忽略並回「建單角色，不適用」提示');
  ok(/建單角色/.test((r.json && r.json.notice) || ''), '提示文字：此角色為建單角色，不適用');
  ok(!('quoteMode' in readAuth().users.find(u => u.username === '_s_n3')), '建單角色帳號沒有存 quoteMode');
  r = await newUser('_s_n4', 'marketing');
  eq(r.status, 200, '行銷（預設沒有報價單功能）新增不需要類型');
  // 列表欄位
  const row = await userRow('_s_n1');
  eq([row.quoteModeApplicable, row.quoteModeRequired], [true, true], '列表：秘書 applicable／required 都為 true');
  const rowU = await userRow('_s_s1');
  eq([rowU.quoteMode, rowU.quoteModeApplicable, rowU.quoteModeRequired], ['', false, false], '列表：業務不適用');
  eq([(await userRow('_s_n4')).quoteModeApplicable, (await userRow('_s_n4')).quoteModeRequired], [true, false], '列表：行銷適用但（功能關閉）不必填');
  eq((await userRow('_s_admin')).quoteModeApplicable, false, '列表：管理員不適用');
  // 稽核與通知
  let logs = await getLogs();
  const l1 = byAction(logs, 'QUOTE_MODE_CHANGE').filter(l => l.target === '_s_n1');
  ok(l1.length === 1 && /報價單權限類型 未設定→唯讀/.test(l1[0].detail) && l1[0].operator === '_s_admin', '新增即設定也寫 QUOTE_MODE_CHANGE（未設定→唯讀）：' + (l1[0] && l1[0].detail));
  ok(byAction(logs, 'QUOTE_MODE_CHANGE').filter(l => l.target === '_s_n3').length === 0, '被忽略（建單角色）沒有 QUOTE_MODE_CHANGE');
  // 編輯：未設定→簽核，通知本人（用種子帳號 _s_sec2，已登入可讀通知）
  const nBefore = (await notesOf('_s_sec2')).length;
  r = await setMode('_s_sec2', 'approver');
  eq(r.status, 200, '舊帳號秘書改為簽核 → 200');
  let notes = await notesOf('_s_sec2');
  const nn = notes.find(n => n.type === 'quote_mode_changed');
  ok(notes.length === nBefore + 1 && nn && /已由管理員 _s_admin 設為『簽核』/.test(nn.body) && /你的報價單簽核權限/.test(nn.body), '通知本人：「你的報價單簽核權限已由管理員 X 設為『簽核』」：' + (nn && nn.body));
  logs = await getLogs();
  ok(byAction(logs, 'QUOTE_MODE_CHANGE').some(l => l.target === '_s_sec2' && /未設定→簽核/.test(l.detail)), '稽核 QUOTE_MODE_CHANGE 未設定→簽核');
  ok(byAction(logs, 'UPDATE_USER').some(l => l.target === '_s_sec2' && /報價單權限類型 未設定→簽核/.test(l.detail)), 'UPDATE_USER 的欄位 diff 也含 quoteMode');
  // 簽核→唯讀：稽核有 before→after；不通知
  const nMid = (await notesOf('_s_sec2')).length;
  await setMode('_s_sec2', 'readonly');
  eq((await notesOf('_s_sec2')).length, nMid, '改為唯讀不發「設為簽核」通知');
  logs = await getLogs();
  ok(byAction(logs, 'QUOTE_MODE_CHANGE').some(l => l.target === '_s_sec2' && /簽核→唯讀/.test(l.detail)), '稽核 簽核→唯讀');
  // 沒變不重複記
  const cnt = byAction(logs, 'QUOTE_MODE_CHANGE').length;
  await setMode('_s_sec2', 'readonly');
  eq(byAction(await getLogs(), 'QUOTE_MODE_CHANGE').length, cnt, '值沒變不重複寫 QUOTE_MODE_CHANGE');
  // 不可清回未設定（必填）
  r = await setMode('_s_sec2', '');
  eq([r.status, r.json && r.json.code], [400, 'QUOTE_MODE_REQUIRED'], '必填角色已設定後不能清回未設定 → 400');
  r = await setMode('_s_sec2', null);
  eq([r.status, r.json && r.json.code], [400, 'QUOTE_MODE_REQUIRED'], 'null 同樣不能清除');
  // 未帶欄位的其他編輯不動它
  r = await putUser('_s_sec2', { displayName: '改名不動類型' });
  eq(r.status, 200, '只改顯示名稱 200');
  eq((await userRow('_s_sec2')).quoteMode, 'readonly', '只改顯示名稱 → quoteMode 不變');
  // 角色改成「原本不適用→適用」必填
  r = await putUser('_s_s3', { role: 'secretary' });
  eq([r.status, r.json && r.json.code], [400, 'QUOTE_MODE_REQUIRED'], '業務改成秘書沒給類型 → 400 QUOTE_MODE_REQUIRED');
  eq((await userRow('_s_s3')).role, 'user', '被拒絕後角色沒變');
  r = await putUser('_s_s3', { role: 'secretary', quoteMode: 'approver' });
  eq(r.status, 200, '業務改成秘書＋簽核 → 200');
  // 角色衝突：簽核改建單角色
  r = await putUser('_s_s3', { role: 'user' });
  eq([r.status, r.json && r.json.code], [400, 'QUOTE_MODE_ROLE_CONFLICT'], '簽核者改回建單角色（沒清除）→ 400 QUOTE_MODE_ROLE_CONFLICT');
  r = await putUser('_s_s3', { role: 'user', quoteMode: 'approver' });
  eq([r.status, r.json && r.json.code], [400, 'QUOTE_MODE_ROLE_CONFLICT'], '簽核者改回建單角色且仍送簽核 → 400');
  eq((await userRow('_s_s3')).role, 'secretary', '衝突被拒絕後角色與類型都沒變');
  r = await putUser('_s_s3', { role: 'user', quoteMode: '' });
  eq(r.status, 200, '簽核者改回建單角色＋明確清除 → 200');
  let rw = await userRow('_s_s3');
  eq([rw.role, rw.quoteMode], ['user', ''], '角色回到 user、類型已清除');
  ok(byAction(await getLogs(), 'QUOTE_MODE_CHANGE').some(l => l.target === '_s_s3' && /簽核→未設定/.test(l.detail)), '清除也有稽核（簽核→未設定）');
  // 唯讀改建單角色：自動清除
  await putUser('_s_s3', { role: 'secretary', quoteMode: 'readonly' });
  r = await putUser('_s_s3', { role: 'user' });
  eq(r.status, 200, '唯讀者改回建單角色 → 200（自動清除）');
  ok(/自動清除/.test((r.json && r.json.notice) || ''), '回應提示已自動清除');
  eq((await userRow('_s_s3')).quoteMode, '', '唯讀者改建單角色後類型已清除');
  // 建單角色送值被忽略
  r = await setMode('_s_s2', 'approver');
  eq([r.status, !!(r.json && r.json.notice)], [200, true], '對建單角色（業務）設簽核 → 忽略並提示');
  eq((await userRow('_s_s2')).quoteMode, '', '業務沒有被存上 quoteMode');
  r = await setMode('_s_m2', 'readonly');
  eq((await userRow('_s_m2')).quoteMode, '', '二級主管（建單角色）也不存');
  // 自我護欄
  r = await setMode('_s_admin', 'approver');
  eq([r.status, r.json && r.json.code], [403, 'QUOTE_MODE_SELF_CHANGE_FORBIDDEN'], '管理員改自己的類型 → 403 QUOTE_MODE_SELF_CHANGE_FORBIDDEN');
  ok(byAction(await getLogs(), 'ADMIN_RAIL_BLOCKED').some(l => l.operator === '_s_admin' && /QUOTE_MODE_SELF_CHANGE_FORBIDDEN/.test(l.detail)), '自我變更被擋有 ADMIN_RAIL_BLOCKED 稽核');
  r = await putUser('_s_admin', { displayName: '沙盒管理員' });
  eq(r.status, 200, '管理員改自己的顯示名稱（沒送 quoteMode）仍可');
  r = await setMode('_s_admin', 'approver'); // 另一位管理員對管理員設定：不適用 → 忽略
  r = await putUser('_s_admin', { quoteMode: 'approver' }, '_s_admin2');
  eq([r.status, !!(r.json && r.json.notice)], [200, true], '另一位管理員對管理員帳號設類型 → 不適用被忽略');
  eq((await userRow('_s_admin')).quoteMode, '', '管理員帳號沒有 quoteMode');
  // 一般使用者不能呼叫
  r = await setMode('_s_s1', 'approver'); // 這行是 admin 呼叫（無害）
  r = await http('_s_s1', 'PUT', '/api/admin/users/_s_m1', { quoteMode: 'approver' });
  ok([401, 403].includes(r.status), '一般使用者不能改類型（' + r.status + '）');
  r = await http('_s_s1', 'GET', '/api/admin/users');
  ok([401, 403].includes(r.status), '一般使用者不能列帳號（看不到他人的類型）');
  // /api/me 只回自己的
  r = await http('_s_sec2', 'GET', '/api/me');
  eq(r.json && r.json.quoteMode, 'readonly', '/api/me 回本人的 quoteMode');
  ok(!/\$2[aby]\$|"password"/.test(r.text), '/api/me 不含密碼雜湊');
  r = await http('_s_s1', 'GET', '/api/me');
  eq(r.json && r.json.quoteMode, '', '業務的 /api/me quoteMode 為空字串');
  r = await http('_s_sec2', 'GET', '/api/me/permissions');
  eq(r.json && r.json.quoteMode, 'readonly', '/api/me/permissions 也帶本人的 quoteMode');
  // 角色功能權限矩陣：新開報價單功能 → 列出沒設定的帳號（不阻擋）
  const rp = await http('_s_admin', 'GET', '/api/admin/role-permissions');
  ok(Array.isArray(rp.json.nonCreatorQuoteRoles) && rp.json.nonCreatorQuoteRoles.includes('secretary') && !rp.json.nonCreatorQuoteRoles.includes('user') && !rp.json.nonCreatorQuoteRoles.includes('manager2') && !rp.json.nonCreatorQuoteRoles.includes('pool') && !rp.json.nonCreatorQuoteRoles.includes('admin'), 'role-permissions 回傳 nonCreatorQuoteRoles（含秘書；不含業務／二級主管／pool／admin）');
  const matrix = JSON.parse(JSON.stringify(rp.json.matrix));
  matrix.marketing = (matrix.marketing || []).concat('quotations');
  r = await http('_s_admin', 'PUT', '/api/admin/role-permissions', { matrix });
  eq(r.status, 200, '行銷新開報價單功能 → 儲存成功（不阻擋）');
  const miss = (r.json && r.json.quoteModeMissing) || [];
  ok(miss.some(m => m.role === 'marketing' && m.users.includes('_s_mkt') && m.users.includes('_s_n4')), '回應列出行銷中未設定類型的帳號：' + JSON.stringify(miss));
  ok(!miss.some(m => m.role !== 'marketing'), '只列出新開功能的角色');
  ok(byAction(await getLogs(), 'UPDATE_ROLE_PERMISSIONS').some(l => /未設定報價單權限類型的帳號：marketing\(/.test(l.detail)), '稽核也記下哪些帳號未設定');
  eq((await userRow('_s_mkt')).quoteModeRequired, true, '功能開啟後行銷帳號變成必填');
  r = await newUser('_s_n5', 'marketing');
  eq([r.status, r.json && r.json.code], [400, 'QUOTE_MODE_REQUIRED'], '功能開啟後新增行銷沒給類型 → 400');
  r = await http('_s_admin', 'DELETE', '/api/admin/role-permissions');
  eq(r.status, 200, '還原預設矩陣');
  // 沒有任何機敏字樣進稽核
  ok(!/\$2[aby]\$/.test(JSON.stringify(await getLogs())) && !JSON.stringify(await getLogs()).includes(PW) && !JSON.stringify(await getLogs()).includes(NEWPW), '稽核紀錄沒有密碼或雜湊');
}

// ═════════════ 第 2 項：伺服器端執法（唯讀不能簽核／不被指派／不進名冊） ═════════════
async function sec2() {
  console.log('[2] 伺服器端執法');
  await loginAll(['_s_admin', '_s_s2', '_s_s3', '_s_m1', '_s_m1b', '_s_gma', '_s_gmb', '_s_ch', '_s_sec1', '_s_sec2', '_s_sec3', '_s_sec4', '_s_cm']);
  // 舊帳號（沒有 quoteMode）照舊可簽
  let r = await mkQuote('_s_s2', ...HI);
  eq(r.status, 201, '業務建單 201');
  const Q1 = r.json.id;
  r = await submitQ('_s_s2', Q1);
  eq(r.status, 200, '業務送簽 200');
  eq([stepsOf(r.json), assigneeOf(Q1)], [['mgr1'], '_s_m1'], '低門檻單：只需一級主管，指派給 m1（舊帳號）');
  ok((await inboxIds('_s_m1')).includes(Q1), '舊帳號 m1：待簽清單看得到');
  r = await getQ('_s_m1', Q1);
  eq([r.json.perm.canApprove, r.json.perm.canReturn, !!r.json.perm.quoteReadonly], [true, true, false], '舊帳號 m1：perm 可核准／駁回');
  // 設為唯讀 → 即使已被指派也立刻失去簽核能力
  r = await setMode('_s_m1', 'readonly');
  eq(r.status, 200, 'm1 設為唯讀');
  ok(!(await inboxIds('_s_m1')).includes(Q1), '唯讀：待簽清單不再列出（即使是已被指派的關卡）');
  r = await getQ('_s_m1', Q1);
  eq(r.status, 200, '唯讀者仍可檢視報價單（既有可視規則不變）');
  eq([r.json.perm.canApprove, r.json.perm.canReturn, r.json.perm.canReassign, r.json.perm.quoteReadonly], [false, false, false, true], '唯讀：perm 的核准／駁回／改派旗標全為 false，quoteReadonly=true（前端據此隱藏按鈕）');
  r = await actQ('_s_m1', Q1, 'approve');
  eq([r.status, r.json && r.json.code], [403, 'NOT_YOUR_STEP'], '唯讀者直接打 approve → 403 NOT_YOUR_STEP');
  r = await actQ('_s_m1', Q1, 'return');
  eq([r.status, r.json && r.json.code], [403, 'NOT_YOUR_STEP'], '唯讀者直接打 return → 403');
  eq((await getQ('_s_s2', Q1)).json.approval.state, 'pending', '報價單仍在簽核中（沒被唯讀者動到）');
  // 改派：不可改派給唯讀者
  r = await http('_s_admin', 'POST', `/api/quotations/${Q1}/reassign`, { username: '_s_m1', reason: '沙盒' });
  eq([r.status, r.json && r.json.code], [400, 'BAD_ASSIGNEE'], '管理員不能改派給唯讀的一級主管 → 400 BAD_ASSIGNEE');
  r = await http('_s_admin', 'POST', `/api/quotations/${Q1}/reassign`, { username: '_s_m1b', reason: '沙盒' });
  eq(r.status, 200, '改派給（舊帳號）m1b → 200');
  ok((await inboxIds('_s_m1b')).includes(Q1), 'm1b 待簽清單出現');
  // 候選人排除：s2 的主管鏈只有 m1（唯讀）→ 送簽被擋
  r = await mkQuote('_s_s2', ...LOW);
  const Q2 = r.json.id;
  r = await submitQ('_s_s2', Q2);
  ok(r.status === 400 && r.json.code === 'CANNOT_SUBMIT' && (r.json.blockers || []).some(b => b.code === 'NO_MANAGER'), '主管鏈上只有唯讀的一級主管 → 送簽被擋（NO_MANAGER，唯讀者不被當成簽核人）：' + JSON.stringify(r.json && r.json.blockers && r.json.blockers.map(b => b.code)));
  const pv = (await getQ('_s_s2', Q2)).json.preview;
  ok(pv && (pv.blockers || []).some(b => b.code === 'NO_MANAGER'), '預覽（候選人檢查）同樣列出 NO_MANAGER');
  // 改回簽核 → 恢復
  await setMode('_s_m1', 'approver');
  r = await submitQ('_s_s2', Q2);
  eq([r.status, assigneeOf(Q2)], [200, '_s_m1'], 'm1 設為簽核後可被指派，送簽成功');
  // 簽核者可正常簽（與以前相同）
  r = await actQ('_s_m1', Q2, 'approve');
  eq(r.status, 200, '簽核類型 m1 核准一級關 → 200');
  eq(r.json.approval.steps[1].tier, 'gm', '下一關是總經理');
  // 總經理關：gmb 設為唯讀（舊名冊仍有他）
  const gmaNotes0 = (await notesOf('_s_gma')).length;
  ok((await inboxIds('_s_gma')).includes(Q2) && (await inboxIds('_s_gmb')).includes(Q2), 'gm 關：兩位舊名冊成員 inbox 都有');
  await setMode('_s_gmb', 'readonly');
  ok(!(await inboxIds('_s_gmb')).includes(Q2) && (await inboxIds('_s_gma')).includes(Q2), 'gmb 設唯讀後 inbox 排除，gma 仍在');
  r = await getQ('_s_gmb', Q2);
  eq([r.status, r.json && r.json.perm.canApprove, r.json && r.json.perm.quoteReadonly], [200, false, true], '唯讀總經理仍可檢視（執行長可視範圍），但 canApprove=false');
  r = await actQ('_s_gmb', Q2, 'approve');
  eq([r.status, r.json && r.json.code], [403, 'NOT_YOUR_STEP'], '名冊內的唯讀總經理打 approve → 403（舊名冊項目在動作當下被擋）');
  r = await actQ('_s_gma', Q2, 'approve');
  eq(r.status, 200, '舊帳號 gma 仍可核准');
  eq((await getQ('_s_s2', Q2)).json.approval.state, 'approved', '整張單核准');
  // 通知收件人排除：新單送簽到 gm 關，gmb（唯讀）不收通知
  r = await mkQuote('_s_s3', ...LOW); const Q3 = r.json.id;
  await submitQ('_s_s3', Q3);   // 主管 m1b
  const gmbN0 = (await notesOf('_s_gmb')).length, gmaN0 = (await notesOf('_s_gma')).length;
  r = await actQ('_s_m1b', Q3, 'approve');
  eq(r.status, 200, 'm1b 核准後輪到 gm');
  eq((await notesOf('_s_gmb')).length, gmbN0, '唯讀的 gmb 沒有收到「待簽核」通知（同一個 stepRecipients 也供信件使用）');
  ok((await notesOf('_s_gma')).length === gmaN0 + 1, '簽核類（舊帳號）gma 收到通知');
  // 沒有可簽的人：gm 名冊只剩唯讀者 → 送簽擋 NO_GM
  r = await putRoster({ gm: ['_s_gmb'] });
  eq(r.status, 200, '名冊裡原本就有的唯讀者，其他欄位儲存不受阻（只擋「新增」）');
  r = await mkQuote('_s_s2', ...LOW); const Q4 = r.json.id;
  r = await submitQ('_s_s2', Q4);
  ok(r.status === 400 && (r.json.blockers || []).some(b => b.code === 'NO_GM'), '總經理名冊只有唯讀者 → 送簽擋 NO_GM（候選人排除）：' + JSON.stringify(r.json && r.json.blockers && r.json.blockers.map(b => b.code)));
  await putRoster({ gm: ['_s_gma', '_s_gmb'] });
  // 名冊：不能新增唯讀者
  await setMode('_s_sec3', 'readonly');
  for (const [key, label] of [['gm', '總經理'], ['chairman', '董事長'], ['boardProxy', '董事會代核'], ['sealManagers', '報價章']]) {
    const cur = (await getCfg()).roster[key];
    r = await putRoster({ [key]: cur.concat('_s_sec3') });
    ok(r.status === 400 && r.json.code === 'ROSTER_READONLY_USER' && /唯讀/.test(r.json.error), `名冊「${label}」不能新增唯讀者 → 400 ROSTER_READONLY_USER：` + (r.json && r.json.error));
    ok(!(await getCfg()).roster[key].includes('_s_sec3'), `名冊「${label}」沒有被寫入`);
  }
  r = await putRoster({ costProviders: ['_s_cm', '_s_sec3'] });
  eq(r.status, 200, '成本填寫人（填寫職能、非簽核）名單不受限');
  await putRoster({ costProviders: ['_s_cm'] });
  r = await putRoster({ gm: ['_s_gma', '_s_gmb', '_s_sec1'] });
  eq(r.status, 200, '簽核類／未設定的人可以加入名冊（秘書 sec1 未設定）');
  await putRoster({ gm: ['_s_gma', '_s_gmb'] });
  // 秘書：代核人與章管理
  await setMode('_s_sec1', 'readonly'); await setMode('_s_sec2', 'approver');
  let c1 = await http('_s_sec1', 'GET', '/api/quote-approval/config');
  eq([c1.json.me.isBoardProxy, c1.json.me.isSealManager, c1.json.me.quoteMode], [false, false, 'readonly'], '唯讀秘書：不是董事會代核人、不是章管理人，me.quoteMode=readonly');
  let c2 = await http('_s_sec2', 'GET', '/api/quote-approval/config');
  eq([c2.json.me.isBoardProxy, c2.json.me.isSealManager], [true, true], '簽核秘書：仍是代核人與章管理人');
  let c3 = await http('_s_gma', 'GET', '/api/quote-approval/config');
  eq(c3.json.me.quoteMode, '', '舊帳號 me.quoteMode 為空字串');
  r = await http('_s_sec1', 'GET', '/api/quote-approval/seal');
  eq(r.status, 403, '唯讀秘書不能檢視章圖');
  r = await http('_s_sec1', 'DELETE', '/api/quote-approval/seal');
  eq(r.status, 403, '唯讀秘書不能刪除章');
  r = await http('_s_sec1', 'PUT', '/api/quote-approval/seal', {});
  eq(r.status, 403, '唯讀秘書不能上傳章');
  // 唯讀秘書仍可檢視（只擋動作）：ITS 秘書本來因為『輪到董事會代核關』才看得到 ERP 的單
  let bd = await getQ('_s_sec4', 'q-BD');
  eq([bd.status, bd.json && bd.json.perm.canApprove], [200, true], '董事會關：ITS 秘書（未設定）看得到且可代核');
  await setMode('_s_sec4', 'readonly');
  bd = await getQ('_s_sec4', 'q-BD');
  eq([bd.status, bd.json && bd.json.perm.canApprove, !!(bd.json && bd.json.perm.quoteReadonly)], [200, false, true], '唯讀秘書：同一張單仍可檢視（檢視關係不變），只是不能代核');
  r = await http('_s_sec4', 'POST', '/api/quotations/q-BD/approve', { hash: 'h' });
  eq([r.status, r.json && r.json.code], [403, 'NOT_YOUR_STEP'], '唯讀秘書不能登錄董事會決議／代核');
  // 流程例外規則
  const fc = (await http('_s_admin', 'GET', '/api/admin/quote-approval/flow-config')).json;
  r = await http('_s_admin', 'PUT', '/api/admin/quote-approval/flow-config', { buSettings: fc.buSettings, exceptions: [{ username: '_s_sec1', action: 'skip_mgr1' }] });
  ok(r.status === 400 && r.json.code === 'EXCEPTION_READONLY_USER', '流程例外不能新增唯讀者 → 400 EXCEPTION_READONLY_USER');
  // 管理員名冊內：config 的 users 帶 quoteMode（只有管理員）
  const adminCfg = await getCfg();
  eq((adminCfg.users.find(u => u.username === '_s_sec1') || {}).quoteMode, 'readonly', '管理員的簽核設定可見各人類型（供名冊頁標示）');
  eq('users' in c3.json, false, '非管理員拿不到使用者清單（他人類型不外洩）');
  ok(!JSON.stringify(c3.json).includes('readonly'), '非管理員的設定回應不含他人類型');
  // 補：董事會關的候選人與 inbox（以 board 步驟的單測更完整見變異/報告項）
}

// ═════════════ 第 3 項：建單角色不受影響＋舊帳號＋報告 ═════════════
async function sec3() {
  console.log('[3] 建單角色不受影響／舊帳號／報告');
  await loginAll(['_s_admin', '_s_s1', '_s_s2', '_s_s3', '_s_m1', '_s_m1b', '_s_m2', '_s_gma', '_s_sec1', '_s_ch']);
  // 業務 s1（直屬二級主管 m2、ERP 啟用二級主管關）：建單→送簽→二級→一級
  let r = await mkQuote('_s_s1', ...HI);
  eq(r.status, 201, 'sales 建單 201（建單角色不受 quoteMode 影響）');
  const A = r.json.id;
  r = await submitQ('_s_s1', A);
  eq([r.status, stepsOf(r.json)], [200, ['mgr2', 'mgr1']], 'sales 送簽：二級主管關→一級主管關');
  eq(assigneeOf(A), '_s_m2', '二級主管關指派給 m2');
  r = await actQ('_s_m2', A, 'approve');
  eq(r.status, 200, '二級主管 m2（建單角色，沒有 quoteMode）核准二級關');
  r = await actQ('_s_m1', A, 'approve');
  eq([r.status, r.json && r.json.approval.state], [200, 'approved'], '一級主管 m1（舊帳號）核准 → 已核准');
  // 對建單角色設 quoteMode 無效，且不影響其簽核能力
  await setMode('_s_s1', 'readonly'); await setMode('_s_m2', 'readonly');
  eq([(await userRow('_s_s1')).quoteMode, (await userRow('_s_m2')).quoteMode], ['', ''], '對 sales／manager2 設定不會存上');
  r = await mkQuote('_s_s1', ...HI);
  eq(r.status, 201, '被「設定」過後 sales 照樣可建單');
  const B = r.json.id;
  r = await submitQ('_s_s1', B);
  eq(r.status, 200, 'sales 照樣可送簽');
  ok((await inboxIds('_s_m2')).includes(B), 'manager2（建單角色）仍可作為二級主管關簽核人（quoteMode 不適用）');
  r = await actQ('_s_m2', B, 'approve');
  eq(r.status, 200, 'manager2 照樣核准二級關');
  // manager2 自己的單
  r = await mkQuote('_s_m2', ...HI);
  eq(r.status, 201, 'manager2 建立自己的報價單 201');
  const C = r.json.id;
  r = await submitQ('_s_m2', C);
  eq([r.status, stepsOf(r.json)], [200, ['mgr1']], 'manager2 自己的單：一級主管關（主管鏈上的 m1）');
  r = await actQ('_s_m1', C, 'approve');
  eq([r.status, r.json && r.json.approval.state], [200, 'approved'], 'm1 核准 manager2 的單');
  // manager1 仍是舊帳號時完整保留既有行為：B 的一級關
  r = await actQ('_s_m1', B, 'approve');
  eq(r.status, 200, '舊帳號一級主管完成 B 的一級關');
  // 非建單者仍不可建單
  r = await mkQuote('_s_m1', ...HI, { oppId: 'o-_s_s1' });
  eq([r.status, r.json && r.json.code], [403, 'NO_CREATE_ROLE'], 'manager1 仍不可建單（NO_CREATE_ROLE 不變）');
  await setMode('_s_gma', 'approver');
  r = await mkQuote('_s_gma', ...HI, { oppId: 'o-_s_s1' });
  eq([r.status, r.json && r.json.code], [403, 'NO_CREATE_ROLE'], '簽核類型的總經理也不能建單');
  // 舊帳號行為不變＋報告
  r = await putUser('_s_m1b', { displayName: '舊帳號改顯示名稱' });
  eq(r.status, 200, '舊帳號（未設定）只改顯示名稱不被強迫補類型');
  ok(!('quoteMode' in readAuth().users.find(u => u.username === '_s_m1b')), '舊帳號沒有被自動遷移（沒有 quoteMode 欄位）');
  const rowOld = await userRow('_s_m1b');
  eq([rowOld.quoteMode, rowOld.quoteModeRequired], ['', true], '列表：舊帳號 quoteMode 空字串且標示為必填（UI 顯示「未設定」）');
  r = await mkQuote('_s_s3', ...LOW); const D = r.json.id;
  await submitQ('_s_s3', D);   // 待 m1b 簽（身負職責）
  await setMode('_s_sec2', 'readonly'); await setMode('_s_acct', 'approver');
  r = await http('_s_admin', 'GET', '/api/admin/quote-mode-report');
  eq(r.status, 200, '報告端點 200');
  const rep = r.json; const row = (u) => rep.rows.find(x => x.username === u) || {};
  eq(row('_s_m1b').status, 'unset', '報告：m1b 未設定');
  ok(row('_s_m1b').issues.includes('UNSET_WITH_DUTIES') && row('_s_m1b').holdsDuties && row('_s_m1b').duties.mgr1For >= 1 && row('_s_m1b').duties.pendingCount >= 1, '報告：m1b 未設定且身負簽核職責（一級主管鏈＋待簽關卡）：' + JSON.stringify(row('_s_m1b').duties));
  ok(row('_s_gmb').issues.includes('UNSET_WITH_DUTIES') && row('_s_gmb').duties.roster.includes('總經理名冊'), '報告：gmb 在總經理名冊且未設定');
  eq(row('_s_gma').status, 'approver', '報告：gma 簽核');
  ok(row('_s_sec2').status === 'readonly' && row('_s_sec2').issues.includes('READONLY_WITH_DUTIES'), '報告：唯讀秘書仍有（秘書角色的）代核職責 → READONLY_WITH_DUTIES');
  ok(!rep.rows.some(x => x.username === '_s_s1' || x.username === '_s_m2' || x.username === '_s_admin'), '報告不含建單角色與管理員');
  ok(rep.rows.some(x => x.username === '_s_off' && x.active === false), '報告含停用帳號並標 active=false');
  ok(!rep.rows.some(x => x.username === '_s_mkt'), '報告不含沒有報價單功能的行銷');
  eq(rep.summary.unset >= 3 && rep.summary.approver >= 2 && rep.summary.readonly >= 1, true, '報告 summary 數字合理：' + JSON.stringify(rep.summary));
  ok(!/\$2[aby]\$|password|email|@/i.test(r.text), '報告沒有密碼／雜湊／Email 等機敏欄位');
  eq((await http('_s_s1', 'GET', '/api/admin/quote-mode-report')).status === 403 || (await http('_s_s1', 'GET', '/api/admin/quote-mode-report')).status === 401, true, '非管理員不能看報告');
  ok(byAction(await getLogs(), 'VIEW_QUOTE_MODE_REPORT').length >= 1, '檢視報告有稽核紀錄');
  // 沒有資料被報告端點改動（唯讀）：兩次呼叫之間資料雜湊相同（除了稽核）
  const h1 = crypto.createHash('sha256').update(JSON.stringify(readAuth()) + JSON.stringify(readData().quotations)).digest('hex');
  await http('_s_admin', 'GET', '/api/admin/quote-mode-report');
  const h2 = crypto.createHash('sha256').update(JSON.stringify(readAuth()) + JSON.stringify(readData().quotations)).digest('hex');
  eq(h1 === h2, true, '報告端點不改動帳號與報價單資料');
  // 稽核與機敏
  const logs = await getLogs();
  const all = JSON.stringify(logs);
  ok(!/\$2[aby]\$/.test(all) && !all.includes(PW) && !all.includes(NEWPW), '稽核紀錄沒有密碼或雜湊');
  ok(byAction(logs, 'QUOTE_MODE_CHANGE').length >= 3, 'QUOTE_MODE_CHANGE 筆數（' + byAction(logs, 'QUOTE_MODE_CHANGE').length + '）');
  // 即時失效（wave-1）仍有效：停用帳號下一請求 401
  r = await putUser('_s_gma', { active: false });
  eq(r.status, 200, '停用 gma');
  r = await http('_s_gma', 'GET', '/api/me');
  eq(r.status, 401, 'wave-1：停用後既有連線立即 401（未被破壞）');
}

// ═════════════ 第 4 項：靜態檢查（前端／檔案紀律） ═════════════
function sec4() {
  console.log('[4] 靜態檢查');
  const rd = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');
  const admin = rd('_client/admin.html'), qa = rd('_client/quote-approval.js'), qr = rd('lib/quoteRoutes.js'), srv = rd('server.js');
  ok(/id="fieldQuoteMode"/.test(admin) && /報價單權限類型/.test(admin) && /value="approver"/.test(admin) && /value="readonly"/.test(admin), 'admin.html 有「報價單權限類型」下拉（簽核／唯讀）');
  ok(/function qmShown\(role\)\s*\{\s*return qmApplicable\(role\) && qmFeatureOn\(role\)/.test(admin), '欄位只在「非建單角色且有報價單功能」時顯示（由角色權限矩陣資料計算）');
  ok(/nonCreatorQuoteRoles/.test(admin) && /nonCreatorQuoteRoles/.test(srv), '非建單角色清單由伺服器單一來源提供，前端不複製');
  ok(/請選擇「報價單權限類型」/.test(admin) && /aria-required="true"/.test(admin), '必填的行內驗證與無障礙屬性');
  ok(/function uvQmBadge/.test(admin) && /uvRoleBadge\(u\)\}\$\{uvQmBadge\(u\)\}/.test(admin), '使用者列表（表格與卡片）有徽章');
  ok(!/\.innerHTML\s*=\s*[^;]*\bx\.(role|username|displayName)\b(?![^;]*escH)/.test(admin.slice(admin.indexOf('async function openQuoteModeReport'), admin.indexOf("$('qmReportBtn')"))), '報告彈窗輸出的帳號／角色文字都經 escH');
  ok(/perm\.canApprove/.test(qa) && /perm\.canReturn/.test(qa) && !/quoteMode/.test(qa.replace(/u\.quoteMode/g, '').replace(/perm\.quoteReadonly/g, '')), 'quote-approval.js 的簽核按鈕只看伺服器 perm 旗標（只在名冊候選與說明橫幅用到 quoteMode）');
  ok(/canApproveAct/.test(qr) && (qr.match(/canApproveAct\(/g) || []).length >= 7, 'quoteRoutes 以 canApproveAct 集中判斷簽核資格（出現 ' + (qr.match(/canApproveAct\(/g) || []).length + ' 次）');
  ok(/require\('\.\/quoteMode'\)/.test(qr) && /const NO_CREATE_ROLES = QM\.NO_CREATE_ROLES/.test(qr), 'NO_CREATE_ROLES 單一來源在 lib/quoteMode.js');
  const QMmod = require(path.join(ROOT, 'lib', 'quoteMode.js'));
  eq(QMmod.NON_CREATOR_ROLES.slice().sort(), ['accounting_manager', 'consult_manager_north', 'consult_manager_south', 'executive', 'finance_manager', 'groupsales', 'manager1', 'marketing', 'secretary', 'tecopm'], '非建單角色清單');
  ok(QMmod.isApplicable('secretary', 'x', {}) && !QMmod.isApplicable('user', 'x', {}) && !QMmod.isApplicable('manager2', 'x', {}) && !QMmod.isApplicable('admin', 'x', {}) && QMmod.isApplicable('user', 'g', { gm: ['g'] }), 'isApplicable：非建單角色／名冊總經理適用；業務、二級主管、管理員不適用');
  eq([QMmod.parseInput(undefined), QMmod.parseInput(''), QMmod.parseInput('approver'), QMmod.parseInput('x')].map(x => x.ok ? (x.present ? x.value || 'clear' : 'absent') : 'bad'), ['absent', 'clear', 'approver', 'bad'], 'parseInput');
  ok(!/C:\\Users|steven\.lee/i.test(rd('lib/quoteMode.js') + rd('scripts/check-quote-mode.js')), '新檔案沒有本機路徑或人名');
}

// ═════════════ 變異檢查：故意弄壞規則，測試必須抓得到 ═════════════
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
  // 共同情境：m1 被指派 → 設唯讀 → 打 approve。回傳 { approveStatus, canApprove, inInbox }
  const readonlyApprover = async () => {
    await loginAll(['_s_admin', '_s_s2', '_s_m1']);
    const r = await mkQuote('_s_s2', ...HI); const id = r.json.id;
    await submitQ('_s_s2', id);
    await setMode('_s_m1', 'readonly');
    const a = await actQ('_s_m1', id, 'approve');
    const g = await getQ('_s_m1', id);
    return { id, approveStatus: a.status, canApprove: g.json && g.json.perm && g.json.perm.canApprove, inInbox: (await inboxIds('_s_m1')).includes(id) };
  };
  try {
    // 變異 1：唯讀者的簽核資格檢查（stepActorOk）退回舊的 canAct
    {
      const f = mkMut('m1', null, ["if (!step || !(forView ? canAct(users[me]) : canApproveAct(users[me]))) return false;", "if (!step || !canAct(users[me])) return false;"]);
      await fresh(f);
      const o = await readonlyApprover();
      ok(o.approveStatus === 200 || o.canApprove === true, `變異1（stepActorOk 不排除唯讀）→ 唯讀者 approve=${o.approveStatus}、canApprove=${o.canApprove}：sec2 的 403／canApprove=false 斷言會失敗（抓得到）`);
    }
    // 變異 2：指派一級主管時不排除唯讀者
    {
      const f = mkMut('m2', null, ["(includeReadonly ? canAct(up) : canApproveAct(up))", "canAct(up)"]);
      await fresh(f);
      await loginAll(['_s_admin', '_s_s2']);
      await setMode('_s_m1', 'readonly');
      const r = await mkQuote('_s_s2', ...HI);
      const s = await submitQ('_s_s2', r.json.id);
      ok(s.status === 200, `變異2（findManager1 不排除唯讀）→ 送簽成功(${s.status})、指派給唯讀者：sec2 的 NO_MANAGER 斷言會失敗（抓得到）`);
    }
    // 變異 3：名冊允許新增唯讀者
    {
      const f = mkMut('m3', null, ["if (APPROVAL_ROSTER_KEYS.includes(key) && QM.isReadonly(users[un])", "if (false && QM.isReadonly(users[un])"]);
      await fresh(f);
      await loginAll(['_s_admin']);
      await setMode('_s_sec3', 'readonly');
      const r = await putRoster({ gm: ['_s_gma', '_s_gmb', '_s_sec3'] });
      ok(r.status === 200, `變異3（名冊不擋唯讀者）→ 加入成功(${r.status})：sec2 的 ROSTER_READONLY_USER 斷言會失敗（抓得到）`);
    }
    // 變異 4：管理員可改自己的類型
    {
      const f = mkMut('m4', ["if (req.params.username === req.session.user.username && qmPi.present && qmPi.value !== QM.modeOf(auth.users[idx])) {", 'if (false) {']);
      await fresh(f);
      await loginAll(['_s_admin']);
      const r = await setMode('_s_admin', 'approver');
      ok(r.status !== 403, `變異4（拿掉自我變更護欄）→ 不再 403(${r.status})：sec1 的 QUOTE_MODE_SELF_CHANGE_FORBIDDEN 斷言會失敗（抓得到）`);
    }
    // 變異 5：新增非建單角色不必填
    {
      const f = mkMut('m5', ["if (required && !value && (!before || (roleChanged && !wasApplicable))) {", 'if (false) {']);
      await fresh(f);
      await loginAll(['_s_admin']);
      const r = await http('_s_admin', 'POST', '/api/admin/users', { username: '_s_n1', password: NEWPW, displayName: 'x', role: 'secretary', bu: ['ERP'] });
      ok(r.status === 200, `變異5（拿掉必填驗證）→ 新增成功(${r.status})：sec1 的 QUOTE_MODE_REQUIRED 斷言會失敗（抓得到）`);
    }
    // 變異 6：唯讀者仍收簽核通知（stepRecipients 不過濾）
    {
      const f = mkMut('m6', null, ["const ok = (list) => list.filter(un => !QM.isReadonly(ctx.users[un]));", "const ok = (list) => list;"]);
      await fresh(f);
      await loginAll(['_s_admin', '_s_s3', '_s_m1b', '_s_gmb']);
      await setMode('_s_gmb', 'readonly');
      const r = await mkQuote('_s_s3', ...LOW);
      await submitQ('_s_s3', r.json.id);
      const n0 = (await notesOf('_s_gmb')).length;
      await actQ('_s_m1b', r.json.id, 'approve');
      ok((await notesOf('_s_gmb')).length === n0 + 1, '變異6（通知收件人不排除唯讀者）→ 唯讀的 gmb 收到待簽通知：sec2 的「沒收到」斷言會失敗（抓得到）');
    }
    // 對照：正常版同樣情境必須「不被」變異影響
    {
      await fresh();
      const o = await readonlyApprover();
      ok(o.approveStatus === 403 && o.canApprove === false && o.inInbox === false, `對照（正常版）→ approve=${o.approveStatus}、canApprove=${o.canApprove}、inInbox=${o.inInbox}`);
    }
  } finally {
    await stopServer();
    tmpFiles.forEach(f => { try { fs.unlinkSync(f); } catch { /* 已刪 */ } });
  }
}

// ═════════════ 主程式 ═════════════
const SECTIONS = { 1: sec1, 2: sec2, 3: sec3, 4: sec4, M: secM };
(async () => {
  const only = (process.env.QM_ONLY || '').split(',').map(s => s.trim()).filter(Boolean);
  seedFiles();
  try {
    await startServer();
    for (const k of ['1', '2', '3', '4', 'M']) {
      if (!SECTIONS[k] || (k === 'M' && process.env.QM_SKIP_MUT)) continue;
      if (only.length && !only.includes(k)) continue;
      await stopServer(); cleanupFiles(); seedFiles(); await restart();   // 每項一台新伺服器：登入速率限制是行程內記憶體，重啟歸零
      try { await SECTIONS[k](); } catch (e) { fail++; failures.push(`第 ${k} 項例外：${e && e.stack || e}`); console.log('  ✗ 例外', e && e.stack || e); }
    }
  } finally {
    await stopServer();
    cleanupFiles();
  }
  console.log(`\n結果：${pass} 通過，${fail} 失敗`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); process.exit(1); }
})().catch(async e => { console.error(e); await stopServer(); cleanupFiles(); process.exit(1); });
