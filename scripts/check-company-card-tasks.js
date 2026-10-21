#!/usr/bin/env node
'use strict';
/**
 * 企業主檔「責任業務」＋「待補名片」待辦檢查：單元（lib/companyCardTasks.js）＋ API（真實 server.js）＋ 開關關閉位元級不變
 * ＋ 前端互動（headless Edge + CDP，見 _company-card-tasks-ui.js）＋ 變異測試。
 * 不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體；資料檔只存在於本次執行；結束以 PID 清掉伺服器與 Edge；不連遠端資料庫；全部是假公司、假人名。
 *
 *   node scripts/check-company-card-tasks.js                  單元＋API＋開關（不含瀏覽器）
 *   CARDTASK_ONLY=unit,api,import,off,ui node ...             只跑指定群組（ui 需要 Edge）
 *   CARDTASK_MUT=1 node ...                                   另外跑變異測試（每個變異都必須讓對應群組變紅）
 *   CARDTASK_SHOTS=<資料夾> node ...                          另外輸出截圖（ui 群組）
 *   CARDTASK_PORT（預設 3562）／CARDTASK_CDP（預設 9473）／CARDTASK_BASE_DIR（開關比對用的「修改前」程式碼資料夾；沒給就用 git archive origin/main 匯出到暫存）
 * 守門：工作樹已有 auth.json／data.json／audit.log.json（可能是真實資料）時拒絕執行。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.CARDTASK_PORT || '3562', 10);
const CDP = parseInt(process.env.CARDTASK_CDP || '9473', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3161)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map((f) => path.join(ROOT, f));
{
  const g = FILES.filter((f) => /(auth|data|audit.log)\.jsonl?$/.test(f)).filter((f) => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map((f) => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
}
const ONLY = (process.env.CARDTASK_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length ? k !== 'ui' : ONLY.includes(k);
const SHOTS = process.env.CARDTASK_SHOTS || '';
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const XLSX = require(path.join(ROOT, 'node_modules', 'xlsx'));
const CT = require(path.join(ROOT, 'lib', 'companyCardTasks.js'));

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 300)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = (n) => new Promise((r) => setTimeout(r, n));

// ───────────────── 沙盒資料（全部假的）─────────────────
const PW = 'Ct' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const NOW = Date.now();
const nowIso = new Date(NOW).toISOString();
const agoIso = (ms) => new Date(NOW - ms).toISOString();
const DAY = 86400000, HOUR = 3600000;
const XSS = '<img src=x onerror="window.__xss=1">';
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
const mkCo = (id, name, tax, extra) => Object.assign({ id, taxId: tax || '', matchKey: tax ? 'tax:' + tax : 'name:' + name, name, shortName: '', capital: null, industry: '', address: '', phone: '', website: '', region: '', customerCode: '', buImport: [], groupId: null, gcisEnriched: false, source: 'manual', createdAt: nowIso, updatedAt: nowIso }, extra || {});
const mkCard = (id, owner, co, extra) => Object.assign({ id, owner, name: '名片' + id, company: co.name, companyId: co.id, taxId: co.taxId || '', bu: 'ERP', createdAt: nowIso }, extra || {});
const C = {
  c1: mkCo('c1', '假公司甲', '90000001'), c2: mkCo('c2', '假公司乙', '90000002'), c3: mkCo('c3', '假公司丙', '90000003'),
  c4: mkCo('c4', '假公司丁', '90000004'), c5: mkCo('c5', '假公司戊', '90000005'), c6: mkCo('c6', '假公司己', '90000006'),
  c7: mkCo('c7', XSS + '假公司庚', '90000007'),
  c8: mkCo('c8', '假公司逾期一', '90000008', { responsibleSales: ['_s_u3'], responsibleAssignedAt: { _s_u3: agoIso(20 * DAY) } }),
  c9: mkCo('c9', '假公司十四天內', '90000009', { responsibleSales: ['_s_u1'], responsibleAssignedAt: { _s_u1: agoIso(14 * DAY + HOUR) } }),
  c10: mkCo('c10', '假公司十五天', '90000010', { responsibleSales: ['_s_u1'], responsibleAssignedAt: { _s_u1: agoIso(15 * DAY + HOUR) } }),
  c11: mkCo('c11', '假公司兩人', '90000011', { responsibleSales: ['_s_u2', '_s_u3'], responsibleAssignedAt: { _s_u2: agoIso(3 * DAY), _s_u3: agoIso(3 * DAY) } }),
};
function seedUsers() {
  return [
    mkUser('_s_adm', 'admin'), mkUser('_s_mkt', 'marketing', { displayName: '測試行銷' }), mkUser('_s_exec', 'executive'),
    mkUser('_s_mgr1', 'manager1', { bu: ['ERP'], displayName: '測試主管ERP' }), mkUser('_s_mgr2', 'manager1', { bu: ['ITS'], displayName: '測試主管ITS' }),
    mkUser('_s_u1', 'user', { bu: ['ERP'], supervisor: '_s_mgr1', displayName: '測試業務甲' }),
    mkUser('_s_u2', 'user', { bu: ['ERP'], supervisor: '_s_mgr1', displayName: '測試業務乙' }),
    mkUser('_s_u3', 'user', { bu: ['ITS'], supervisor: '_s_mgr2', displayName: '測試業務丙' }),
    mkUser('_s_off', 'user', { bu: ['ERP'], supervisor: '_s_mgr1', displayName: '離職業務', active: false }),
    mkUser('_s_sec', 'secretary', { bu: ['ERP'] }),
  ];
}
function seedFiles(opts) {
  opts = opts || {};
  const DIR = opts.dir || ROOT;
  const users = seedUsers();
  if (opts.users) opts.users(users);
  fs.writeFileSync(path.join(DIR, 'auth.json'), JSON.stringify({ users }, null, 2));
  const cos = Object.values(C).map((c) => JSON.parse(JSON.stringify(c)));
  if (opts.preset) opts.preset(cos);
  const d = {
    contacts: [
      mkCard('k1', '_s_u1', C.c1), mkCard('k2', '_s_u2', C.c2), mkCard('k3', '_s_u3', C.c3, { bu: 'ITS' }),
      mkCard('k5a', '_s_u2', C.c5),
      mkCard('k5b', '_s_u1', C.c5, { isPlaceholder: true, name: '（待補名片）' }),
      mkCard('k5c', '_s_u1', C.c5, { deleted: true }),
      mkCard('k6', '_s_u1', C.c6, { companyId: undefined, company: ' 假公司己 ' }),
    ],
    companies: cos, visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], quotations: [], opportunities: [],
  };
  if (opts.data) opts.data(d);
  fs.writeFileSync(path.join(DIR, 'data.json'), JSON.stringify(d, null, 2));
}
const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
function cleanupFiles(dir) { FILES.forEach((f) => { try { fs.unlinkSync(dir ? path.join(dir, path.basename(f)) : f); } catch { /* 不存在 */ } }); }

// ───────────────── 伺服器（預載腳本可在載入時改寫指定檔案＝變異測試，不動磁碟上的檔案）─────────────────
const PRELOAD = path.join(os.tmpdir(), 'cardtask_preload_' + process.pid + '.js');
fs.writeFileSync(PRELOAD, [
  "const Module=require('module'),path=require('path');const orig=Module.prototype._compile;",
  "Module.prototype._compile=function(content,filename){const raw=process.env.CARDTASK_MUT_SPEC;",
  "if(raw){const specs=JSON.parse(raw);for(const m of specs){if(path.resolve(filename)===path.resolve(process.env.CARDTASK_ROOT,m.file)){for(const e of m.edits){if(!content.includes(e.find))throw new Error('MUTATION ANCHOR MISSING: '+e.find.slice(0,60));content=content.replace(e.find,()=>e.rep);}}}}",
  "return orig.call(this,content,filename);};",
].join('\n') + '\n');
let child = null, childPid = null, curPort = PORT, curRoot = ROOT;
async function startServer(mut, extraEnv, root, port) {
  root = root || ROOT; port = port || PORT;
  const env = Object.assign({}, process.env, { PORT: String(port), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: `http://localhost:${port}` });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'QUOTE_FLOW_ENGINE', 'QUOTE_FLOW_DESIGNER', 'FEATURE_GUARD_MODE', 'CARDTASK_MUT_SPEC', 'COMPANY_CARD_TASKS']) delete env[k];
  Object.assign(env, extraEnv || {});
  if (mut && mut.length) { env.CARDTASK_MUT_SPEC = JSON.stringify(mut); env.CARDTASK_ROOT = ROOT; }
  await assertPortFree(port);
  child = spawn(process.execPath, ['-r', PRELOAD, path.join(root, 'server.js')], { cwd: root, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid; curPort = port; curRoot = root;
  let out = '';
  child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { out += d; });
  for (let i = 0; i < 100; i++) {
    await sleep(250);
    try { const r = await fetch(`http://localhost:${port}/login.html`); if (r.ok) return; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-600));
  }
  throw new Error('伺服器啟動逾時：' + out.slice(-600));
}
async function stopServer() {
  if (!child) return;
  try { process.kill(childPid); } catch { /* 已結束 */ }
  for (let i = 0; i < 30 && child.exitCode === null; i++) await sleep(100);
  child = null;
  await sleep(300);
}
async function freshServer(mut, extraEnv, seedOpts) { await stopServer(); cleanupFiles(); seedFiles(seedOpts); await startServer(mut, extraEnv); jar = {}; }

// ───────────────── HTTP ─────────────────
let jar = {};
async function httpAt(base, jr, who, method, url, body) {
  const h = {};
  if (body !== undefined && !(body instanceof FormData)) h['Content-Type'] = 'application/json';
  if (who && jr[who]) h.Cookie = jr[who];
  const res = await fetch(base + url, { method, headers: h, body: body === undefined ? undefined : (body instanceof FormData ? body : JSON.stringify(body)), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: res.status, json, buf, text: buf.toString('utf8'), headers: res.headers };
}
const http = (who, method, url, body) => httpAt(BASE, jar, who, method, url, body);
async function loginAt(base, jr, who) {
  const res = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status === 200) jr[who] = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
  return res.status;
}
const login = (who) => loginAt(BASE, jar, who);
async function loginAll() { for (const w of ['_s_adm', '_s_mkt', '_s_exec', '_s_mgr1', '_s_mgr2', '_s_u1', '_s_u2', '_s_u3', '_s_sec']) await login(w); }
async function auditLines(action) {
  await sleep(300);
  const r = await http('_s_adm', 'GET', '/api/admin/logs');
  return (Array.isArray(r.json) ? r.json : []).filter((l) => l.action === action);
}
const myTasks = async (who) => (await http(who, 'GET', '/api/my-tasks/company-cards')).json;
const taskNames = async (who) => ((await myTasks(who)).items || []).map((t) => t.companyId).sort();
const setResp = (who, id, list) => http(who, 'POST', `/api/companies/${id}/set-responsible-sales`, { responsibleSales: list });
const notifs = async (who, type) => ((await http(who, 'GET', '/api/notifications')).json || []).filter((n) => n.type === type);
function xlsxBuf(rows) {
  const ws = XLSX.utils.aoa_to_sheet([['公司名稱', '統一編號', '產業', '服務業務', '區域', '客戶代號'], ...rows]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '企業主檔');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}
const upload = (who, url, buf) => { const fd = new FormData(); fd.append('file', new Blob([buf]), 'import.xlsx'); return http(who, 'POST', url, fd); };
const newCard = (who, co, extra) => http(who, 'POST', '/api/contacts', Object.assign({ name: '新名片', company: co.name, taxId: co.taxId, bu: 'ERP' }, extra || {}));

// ═════════════ 單元：lib/companyCardTasks.js ═════════════
function groupUnit() {
  console.log('[單元] lib/companyCardTasks.js');
  eq(CT.enabled({}), true, 'U1 沒設環境變數＝開');
  eq(CT.enabled({ COMPANY_CARD_TASKS: 'off' }), false, 'U2 off＝關');
  eq(CT.enabled({ COMPANY_CARD_TASKS: ' OFF ' }), false, 'U3 大小寫與空白不影響');
  eq(CT.enabled({ COMPANY_CARD_TASKS: 'on' }), true, 'U4 on＝開');
  eq(CT.normList([' a ', 'a', '', null, 'b']), ['a', 'b'], 'U5 normList 去空白去重');
  const users = { u: { username: 'u', role: 'user' }, m: { username: 'm', role: 'marketing' }, p: { username: '_pool', role: 'pool' }, off: { username: 'off', role: 'user', active: false }, g: { username: 'g', role: 'groupsales' }, t: { username: 't', role: 'tecopm' } };
  ['u', 'g'].forEach((k) => ok(CT.eligibleOwner(users[k]), 'U6 可擁有名片的角色合格：' + k));
  ['m', 'p', 'off', 't'].forEach((k) => ok(!CT.eligibleOwner(users[k]), 'U7 不合格：' + k));
  ok(!CT.eligibleOwner(null), 'U8 不存在不合格');
  ok(CT.validateNewOwners(['u'], users).ok, 'U9 合法');
  ['m', 'p', 'off', 'zzz'].forEach((k) => ok(!CT.validateNewOwners([k], users).ok, 'U10 非法被拒：' + k));
  // 待辦推導
  const co = { id: 'x', name: ' 公司X ', responsibleSales: ['u', 'g'], responsibleAssignedAt: { u: agoIso(2 * DAY), g: agoIso(2 * DAY) } };
  const mk = (contacts, extra) => ({ companies: [Object.assign({}, co, extra || {})], contacts, companyCardDismissals: (extra && extra.dis) || [] });
  const now = new Date(NOW);
  eq(CT.deriveTasks(mk([]), users, now).map((t) => t.username), ['u', 'g'], 'U11 零名片 → 兩人都有待辦');
  eq(CT.deriveTasks(mk([{ owner: 'u', companyId: 'x' }]), users, now).map((t) => t.username), ['g'], 'U12 u 有真實名片 → 只剩 g');
  eq(CT.deriveTasks(mk([{ owner: 'u', companyId: 'x', isPlaceholder: true }]), users, now).length, 2, 'U13 placeholder 不算');
  eq(CT.deriveTasks(mk([{ owner: 'u', companyId: 'x', deleted: true }]), users, now).length, 2, 'U14 已刪除不算');
  eq(CT.deriveTasks(mk([{ owner: 'g', companyId: 'x' }]), users, now).map((t) => t.username), ['u'], 'U15 他人補名片不影響 u');
  eq(CT.deriveTasks(mk([{ owner: 'u', company: '公司X' }]), users, now).map((t) => t.username), ['g'], 'U16 沒有 companyId 的名片以公司名 trim 比對');
  eq(CT.deriveTasks(mk([{ owner: 'u', company: '公司X', companyId: 'other' }]), users, now).length, 2, 'U17 有 companyId 但不是這家 → 不以公司名比對');
  eq(CT.deriveTasks(mk([{ owner: 'u', companyId: 'x' }]), { u: users.u }, now).length, 0, 'U18 帳號不存在者不產生待辦（g 不存在、u 有名片）');
  eq(CT.deriveTasks(mk([]), { u: users.off, g: users.g }, now).map((t) => t.username), ['g'], 'U19 停用帳號不產生待辦');
  eq(CT.deriveTasks(mk([]), users, now)[0].days, 2, 'U20 days＝指定後經過天數');
  // 解除
  const data = mk([]);
  eq(CT.dismissTask(data, users, { companyId: 'x', username: 'u', reason: '', note: '' }, now).status, 400, 'U21 沒選原因 400');
  eq(CT.dismissTask(data, users, { companyId: 'x', username: 'u', reason: 'bogus', note: '' }, now).status, 400, 'U22 不在列舉內 400');
  eq(CT.dismissTask(data, users, { companyId: 'x', username: 'u', reason: 'other', note: '   ' }, now).status, 400, 'U23 other 說明空白 400');
  eq(CT.dismissTask(data, users, { companyId: 'x', username: 'u', reason: 'not_mine' }, now).ok, true, 'U24 not_mine 不需說明');
  eq(CT.dismissTask(data, users, { companyId: 'x', username: 'u', reason: 'not_mine' }, now).status, 409, 'U25 同一待辦再解除＝409（已解除過）');
  eq(CT.deriveTasks(data, users, now).find((t) => t.username === 'u').dismissal.reason, 'not_mine', 'U26 解除紀錄綁在待辦上');
  const re = data.companies[0]; re.responsibleAssignedAt.u = new Date(NOW + 1000).toISOString();
  eq(CT.deriveTasks(data, users, now).find((t) => t.username === 'u').dismissal, null, 'U27 指定時間更新 → 舊解除自動失效');
  eq(CT.cleanNote('a\u0001b\u0000c').length, 3, 'U28 備註去控制字元');
  eq(CT.cleanNote('x'.repeat(500)).length, CT.MAX_NOTE, 'U29 備註截長度');
  const ap = { responsibleSales: ['a', 'b'], responsibleAssignedAt: { a: 'T0', b: 'T0' } };
  const df = CT.applyResponsibleSales(ap, ['b', 'c'], 'T1', []);
  eq([ap.responsibleAssignedAt, df.added, df.removed], [{ b: 'T0', c: 'T1' }, ['c'], ['a']], 'U30 保留舊指定時間、新人蓋新時間、移除者消失');
  eq(CT.applyResponsibleSales(ap, ['b', 'c'], 'T2', []).changed, false, 'U31 沒變就是沒變');
  eq(CT.applyResponsibleSales(ap, ['b', 'c'], 'T3', ['b']).changed && ap.responsibleAssignedAt.b === 'T3', true, 'U32 fresh 會重蓋指定時間');
}

// ═════════════ API ═════════════
async function groupApi() {
  console.log('[API] 資料與權限、待辦推導、解除與回報、主管逾期、通知');
  await freshServer(); await loginAll();
  let r;

  // ── 設定權限 ──
  r = await setResp('_s_mkt', 'c4', ['_s_u1', '_s_u2']);
  eq([r.status, r.json && r.json.responsibleSales], [200, ['_s_u1', '_s_u2']], 'P1 行銷可指定責任業務');
  const co4 = readData().companies.find((c) => c.id === 'c4');
  eq(co4.responsibleSales, ['_s_u1', '_s_u2'], 'P2 資料檔已寫入 responsibleSales');
  ok(co4.responsibleAssignedAt._s_u1 && co4.responsibleAssignedAt._s_u2 && !isNaN(Date.parse(co4.responsibleAssignedAt._s_u1)), 'P3 responsibleAssignedAt 每人一個 ISO 時間');
  const stampU1 = co4.responsibleAssignedAt._s_u1;
  r = await setResp('_s_u1', 'c4', ['_s_u1']); eq(r.status, 403, 'P4 一般業務不能設定');
  r = await setResp('_s_sec', 'c4', ['_s_u1']); eq(r.status, 403, 'P5 秘書不能設定');
  r = await setResp('_s_adm', 'c4', ['_s_u1', '_s_u2']); eq(r.status, 200, 'P6 管理員可設定（沒變更也回 200）');
  eq(r.json.changed, false, 'P6b 名單相同＝沒有變更');
  r = await http(null, 'POST', '/api/companies/c4/set-responsible-sales', { responsibleSales: [] }); eq(r.status, 401, 'P7 未登入 401');
  r = await setResp('_s_exec', 'c1', ['_s_u1']); eq(r.status, 200, 'P8 董總可設定（公司有可見名片）');
  // 非法帳號
  const before = JSON.stringify(readData().companies.find((c) => c.id === 'c2'));
  for (const [list, label] of [[['_s_nobody'], '不存在帳號'], [['_pool'], '_pool 客戶池'], [['_s_off'], '停用帳號'], [['_s_mkt'], '行銷（不能擁有名片）'], [['_s_adm'], '管理員'], [['_s_sec'], '秘書']]) {
    r = await setResp('_s_mkt', 'c2', list); eq(r.status, 400, 'P9 非法帳號被拒（400）：' + label);
  }
  r = await http('_s_mkt', 'POST', '/api/companies/c2/set-responsible-sales', { responsibleSales: 'x' }); eq(r.status, 400, 'P9b 不是陣列 400');
  r = await setResp('_s_mkt', 'c2', Array.from({ length: 21 }, (_, i) => '_s_u' + (i % 3 + 1)).concat(['x1', 'x2'])); ok(r.status === 400, 'P9c 超過人數上限或含不存在帳號被拒');
  eq(JSON.stringify(readData().companies.find((c) => c.id === 'c2')), before, 'P9d 被拒的請求完全沒有改到資料');
  r = await setResp('_s_mkt', 'nope', ['_s_u1']); eq(r.status, 404, 'P10 找不到公司 404');
  // 主管範圍
  r = await setResp('_s_mgr1', 'c3', ['_s_u1']); eq(r.status, 403, 'P11 ERP 主管看不到 ITS 的公司（c3 只有 ITS 業務的名片）→ 403');
  r = await setResp('_s_mgr1', 'c1', ['_s_u3']); eq(r.status, 403, 'P12 ERP 主管不能指定 ITS 業務（超出可視範圍）');
  r = await setResp('_s_mgr1', 'c1', ['_s_u2']); eq(r.status, 200, 'P13a ERP 主管可指定自己轄下業務');
  // c1 現況：exec 設 [u1]；mgr1 replace 成 [u2] → u1 在 mgr1 範圍內被移除
  eq(readData().companies.find((c) => c.id === 'c1').responsibleSales, ['_s_u2'], 'P13 主管取代範圍內名單（u1 → u2）');
  await setResp('_s_mkt', 'c1', ['_s_u2', '_s_u3']);     // 行銷設成含 ITS 的 u3
  r = await setResp('_s_mgr1', 'c1', ['_s_u1']);
  eq(readData().companies.find((c) => c.id === 'c1').responsibleSales.sort(), ['_s_u1', '_s_u3'], 'P14 主管取代時，範圍外的 u3 被保留、不會被主管移除');
  // 欄位可見性
  r = await http('_s_mgr1', 'GET', '/api/companies'); const l1 = r.json.find((c) => c.id === 'c1');
  eq(l1.responsibleSales, ['_s_u1'], 'P15 ERP 主管的列表看不到 ITS 責任業務（跨 BU 不外洩）');
  r = await http('_s_mkt', 'GET', '/api/companies'); const m1 = r.json.find((c) => c.id === 'c1');
  eq(m1.responsibleSales.sort(), ['_s_u1', '_s_u3'], 'P16 行銷看得到全部責任業務');
  eq(m1.responsibleSalesDisplay.map((x) => [x.username, x.label, x.bu]).sort(), [['_s_u1', '測試業務甲', ['ERP']], ['_s_u3', '測試業務丙', ['ITS']]], 'P17 附顯示名稱與 BU');
  eq(m1.owners, ['測試業務甲'], 'P18 既有「服務業務」推導欄不變（仍是名片擁有者）');
  r = await http('_s_u2', 'GET', '/api/companies'); eq(r.json.find((c) => c.id === 'c2').responsibleSales, [], 'P19 一般業務只看得到自己（u2 對 c2：沒有責任業務）');
  r = await http('_s_adm', 'GET', '/api/admin/companies'); eq(r.json.find((c) => c.id === 'c1').responsibleSalesDisplay.length, 2, 'P20 後台列表也帶責任業務顯示');
  r = await http('_s_adm', 'GET', '/api/admin/companies/c1/detail'); eq(r.json.company.responsibleSales.length, 2, 'P21 後台明細帶責任業務');
  r = await http('_s_mgr1', 'GET', '/api/companies/c1'); eq(r.json.company.responsibleSales, ['_s_u1'], 'P22 主管看明細也只看到範圍內');
  // 狀態標記
  eq(m1.responsibleStatus._s_u1, 'ok', 'P23 u1 對 c1 有真實名片 → ok');
  eq(m1.responsibleStatus._s_u3, 'pending', 'P24 u3 對 c1 沒有名片 → 待補名片');
  // 稽核
  const lg = await auditLines('SET_COMPANY_RESPONSIBLE');
  ok(lg.length >= 5 && lg.every((l) => /→/.test(l.detail)), 'P25 每次變更都有 SET_COMPANY_RESPONSIBLE 稽核且記 before→after', lg.length);
  ok(lg.some((l) => /測試業務甲\(_s_u1\)/.test(l.detail)), 'P26 稽核內容含顯示名稱與帳號');
  // 指定時間規則
  await sleep(15);
  await setResp('_s_mkt', 'c4', ['_s_u1', '_s_u3']);
  const co4b = readData().companies.find((c) => c.id === 'c4');
  eq([co4b.responsibleAssignedAt._s_u1 === stampU1, co4b.responsibleAssignedAt._s_u3 > stampU1, '_s_u2' in co4b.responsibleAssignedAt], [true, true, false], 'P27 保留者沿用舊指定時間、新加入者蓋新時間、移除者的時間消失');

  // ── 批次 ──
  const nU2 = (await notifs('_s_u2', 'company_card_task')).length;
  r = await http('_s_mkt', 'POST', '/api/companies/responsible-sales-batch', { companyIds: ['c6', 'c7', 'nope'], usernames: ['_s_u2'] });
  eq([r.status, r.json.updated, r.json.denied.length], [200, 2, 1], 'B1 批次指定：2 家成功、1 家找不到');
  eq(r.json.notifiedUsers, 1, 'B2 批次也只通知 u2 一則');
  eq((await notifs('_s_u2', 'company_card_task')).length - nU2, 1, 'B3 u2 只多收到一則彙總通知（兩家公司）');
  r = await http('_s_mkt', 'POST', '/api/companies/responsible-sales-batch', { companyIds: ['c6'], usernames: ['_s_mkt'] }); eq(r.status, 400, 'B4 批次含非法帳號整批拒絕');
  r = await http('_s_u1', 'POST', '/api/companies/responsible-sales-batch', { companyIds: ['c6'], usernames: ['_s_u1'] }); eq(r.status, 403, 'B5 一般業務不能批次');
  r = await http('_s_mgr1', 'POST', '/api/companies/responsible-sales-batch', { companyIds: ['c3', 'c1'], usernames: ['_s_u2'] });
  eq([r.status, r.json.updated, r.json.denied.map((x) => x.id)], [200, 1, ['c3']], 'B6 主管批次：看不到的公司（c3）被拒、其餘照做');
  r = await http('_s_mkt', 'POST', '/api/companies/responsible-sales-batch', { companyIds: [], usernames: ['_s_u2'] }); eq(r.status, 400, 'B7 沒選公司 400');
  r = await http('_s_mkt', 'POST', '/api/companies/responsible-sales-batch', { companyIds: Array.from({ length: 501 }, (_, i) => 'x' + i), usernames: ['_s_u2'] }); eq(r.status, 400, 'B8 超過 500 家 400');
  r = await http('_s_mkt', 'POST', '/api/companies/responsible-sales-batch', { companyIds: ['c6'], usernames: ['_s_u1'], mode: 'replace' });
  eq(readData().companies.find((c) => c.id === 'c6').responsibleSales, ['_s_u1'], 'B9 replace 模式取代');
  // 候選清單
  r = await http('_s_mkt', 'GET', '/api/company-responsible/candidates');
  eq(r.json.users.map((u) => u.username).sort(), ['_s_mgr1', '_s_mgr2', '_s_u1', '_s_u2', '_s_u3'], 'B10 候選＝可擁有名片且啟用的帳號（不含 _pool、停用、行銷、管理員、秘書）');
  ok(r.json.users.every((u) => u.label && Array.isArray(u.bu)), 'B11 候選含顯示名稱與 BU');
  r = await http('_s_mgr1', 'GET', '/api/company-responsible/candidates'); eq(r.json.users.map((u) => u.username).sort(), ['_s_mgr1', '_s_u1', '_s_u2'], 'B12 主管只看到自己範圍內的候選');
  r = await http('_s_u1', 'GET', '/api/company-responsible/candidates'); eq(r.status, 403, 'B13 一般業務 403');

  // ── 待辦推導 ──
  // 目前責任業務：c4=[u1,u3]、c1=[u1,u3]、c6=[u1]、c7=[u2]、(c1 另被主管改過)；seed：c8=[u3]20天、c9=[u1]14d1h、c10=[u1]15d1h、c11=[u2,u3]
  console.log('  [待辦推導]');
  await setResp('_s_mkt', 'c5', ['_s_u1', '_s_u2']);
  let t = await myTasks('_s_u1');
  eq(t.enabled, true, 'T0 功能啟用');
  ok(t.items.some((x) => x.companyId === 'c4'), 'T1 責任業務且零名片 → 有待辦（c4）');
  ok(!t.items.some((x) => x.companyId === 'c1'), 'T2 已有真實名片（c1）→ 沒有待辦');
  ok(t.items.some((x) => x.companyId === 'c5'), 'T3 只有 placeholder＋已刪除名片＋別人的名片（c5）→ 仍有待辦');
  ok(!t.items.some((x) => x.companyId === 'c6'), 'T4 沒有 companyId 的名片以公司名 trim 比對（c6）→ 沒有待辦');
  eq((await myTasks('_s_u2')).items.some((x) => x.companyId === 'c5'), false, 'T5 他人（u2）自己有 c5 的名片 → 對 u2 沒有待辦');
  ok(t.items.find((x) => x.companyId === 'c4').days === 0 && t.items.find((x) => x.companyId === 'c4').overdue === false, 'T6 剛指定：0 天、未逾期');
  ok(t.items.find((x) => x.companyId === 'c10').overdue === true && t.items.find((x) => x.companyId === 'c9').overdue === false, 'T7 逾期邊界：超過 14 天才算（14天又1小時不算、15天又1小時算）');
  ok(t.items[0].days >= t.items[t.items.length - 1].days, 'T8 依天數由多到少排序');
  eq(Object.keys(t.items[0]).sort(), ['address', 'assignedAt', 'companyId', 'days', 'industry', 'name', 'overdue', 'taxId', 'website'], 'T9 待辦欄位（供前端帶入新增名片）');
  // 補一張真實名片 → 自動消失
  const beforeTasks = await taskNames('_s_u1');
  r = await newCard('_s_u1', C.c4); eq(r.status, 201, 'T10 業務補一張真實名片（c4）');
  ok((await taskNames('_s_u1')).indexOf('c4') < 0 && beforeTasks.indexOf('c4') >= 0, 'T11 補名片後待辦自動消失（沒有任何手動動作）');
  eq((await myTasks('_s_u3')).items.some((x) => x.companyId === 'c4'), true, 'T12 c4 的另一位責任業務（u3）不受 u1 補名片影響');
  ok((await taskNames('_s_u3')).includes('c8') && (await taskNames('_s_u3')).includes('c11'), 'T13 多位責任業務各自判斷（u3 在 c8、c11 有待辦）');
  // 主管首頁逾期（> 14 天）：依可視範圍，跨 BU 不外洩
  const odm1 = (await http('_s_mgr1', 'GET', '/api/manager/company-card-overdue')).json, odm2 = (await http('_s_mgr2', 'GET', '/api/manager/company-card-overdue')).json;
  eq(odm1.overdue.map((x) => x.companyId), ['c10'], 'O0a ERP 主管只看到轄下 u1 的逾期（c10）；看不到 ITS 業務 u3 的 c8（跨 BU 不外洩）；c9 剛好 14 天又 1 小時不算');
  eq(odm2.overdue.map((x) => x.companyId), ['c8'], 'O0b ITS 主管只看到 u3 的逾期 c8');
  eq((await http('_s_exec', 'GET', '/api/manager/company-card-overdue')).json.overdue.map((x) => x.companyId), ['c8', 'c10'], 'O0c 董總看到全部，依逾期天數由多到少');
  ok(odm1.overdue[0].userLabel === '測試業務甲' && odm1.overdue[0].bu.join() === 'ERP' && odm1.overdue[0].days >= 15, 'O0d 逾期列含業務名稱、BU、天數');
  await setResp('_s_mkt', 'c4', ['_s_u3']);                       // u1 被移除
  await setResp('_s_mkt', 'c10', ['_s_u2']);                      // c10：u1 被換成 u2
  ok(!(await taskNames('_s_u1')).includes('c10') && (await taskNames('_s_u2')).includes('c10'), 'T14 改人：舊人的待辦消失、新人出現');
  ok(((await myTasks('_s_u2')).items.find((x) => x.companyId === 'c10') || {}).days === 0, 'T15 新人的天數從被指定當下重算');
  // 停用責任業務不產生待辦
  const d0 = readData(); d0.companies.find((c) => c.id === 'c7').responsibleSales = ['_s_u2', '_s_off']; /* 只驗推導，不寫回 */
  eq(CT.deriveTasks(d0, Object.fromEntries(seedUsers().map((u) => [u.username, u])), new Date()).filter((x) => x.companyId === 'c7').map((x) => x.username), ['_s_u2'], 'T17 停用的責任業務不產生待辦');

  // ── 手動解除 ──
  console.log('  [手動解除]');
  const dis = (who, body) => http(who, 'POST', '/api/my-tasks/company-cards/dismiss', body);
  r = await dis('_s_u2', { companyId: 'c11' }); eq(r.status, 400, 'D1 沒選原因 400');
  r = await dis('_s_u2', { companyId: 'c11', reason: 'bogus' }); eq(r.status, 400, 'D2 原因不在列舉內 400');
  r = await dis('_s_u2', { companyId: 'c11', reason: 'other' }); eq(r.status, 400, 'D3 選「其他」沒有說明 400');
  r = await dis('_s_u2', { companyId: 'c11', reason: 'other', note: '   ' }); eq(r.status, 400, 'D4 選「其他」說明全空白 400');
  r = await dis('_s_u1', { companyId: 'c11', reason: 'not_mine' }); eq(r.status, 404, 'D5 不是自己的待辦不能解除（u1 對 c11 沒有待辦）');
  r = await dis('_s_u1', { companyId: 'c11', reason: 'not_mine', username: '_s_u2' }); eq(r.status, 404, 'D5b 在 body 夾帶別人的帳號也不能替人解除（只認登入者）');
  r = await dis('_s_u2', { reason: 'not_mine' }); eq(r.status, 400, 'D6 缺 companyId 400');
  eq((await myTasks('_s_u2')).items.some((x) => x.companyId === 'c11'), true, 'D7 被拒的解除沒有生效');
  r = await dis('_s_u2', { companyId: 'c11', reason: 'not_mine' }); eq(r.status, 200, 'D8 選原因（不是我負責）即可解除，不強迫寫說明');
  t = await myTasks('_s_u2');
  eq([t.items.some((x) => x.companyId === 'c11'), t.dismissed], [false, 1], 'D9 解除後對自己隱藏（dismissed 計數 1）');
  r = await dis('_s_u2', { companyId: 'c11', reason: 'not_mine' }); eq(r.status, 409, 'D10 重複解除 409');
  r = await dis('_s_u2', { companyId: 'c10', reason: 'other', note: XSS + '客戶說不需要' }); eq(r.status, 200, 'D11 其他＋說明（含 XSS 酬載文字）可解除');
  r = await dis('_s_u3', { companyId: 'c8', reason: 'inactive' }); eq(r.status, 200, 'D12 u3 解除 c8（客戶已不往來）');
  r = await dis('_s_u3', { companyId: 'c4', reason: 'no_contact' }); eq(r.status, 200, 'D13 u3 解除 c4（還沒有聯絡人）');
  // 行銷清單
  r = await http('_s_mkt', 'GET', '/api/marketing/responsibility-feedback');
  eq(r.json.count, 4, 'D14 行銷的「業務回報待修」出現 4 筆');
  const fb11 = r.json.items.find((x) => x.companyId === 'c11');
  eq([fb11.username, fb11.reason, fb11.reasonLabel, fb11.userLabel, fb11.bu], ['_s_u2', 'not_mine', '不是我負責', '測試業務乙', ['ERP']], 'D15 回報含原因、業務名稱與 BU');
  ok(r.json.items.find((x) => x.companyId === 'c10').note.includes('客戶說不需要'), 'D16 說明原樣保存（純文字）');
  eq((await http('_s_adm', 'GET', '/api/marketing/responsibility-feedback')).status, 200, 'D17 管理員也可看');
  for (const w of ['_s_u1', '_s_mgr1', '_s_exec', '_s_sec']) eq((await http(w, 'GET', '/api/marketing/responsibility-feedback')).status, 403, 'D18 無權看回報清單：' + w);
  // 主管
  r = await http('_s_mgr1', 'GET', '/api/manager/company-card-overdue');
  eq(r.json.dismissed.map((x) => x.companyId).sort(), ['c10', 'c11'], 'D19 ERP 主管只看到轄下（u1/u2）的回報（看不到 ITS 的 c8、c4）');
  ok(r.json.dismissed.every((x) => !('responsibleSales' in x)), 'D19b 主管看到的回報不含完整責任業務名單');
  r = await http('_s_mgr2', 'GET', '/api/manager/company-card-overdue');
  eq(r.json.dismissed.map((x) => x.companyId).sort(), ['c4', 'c8'], 'D20 ITS 主管只看到 ITS 業務的回報');
  eq((await http('_s_exec', 'GET', '/api/manager/company-card-overdue')).json.dismissed.length, 4, 'D21 董總看到全部');
  for (const w of ['_s_u1', '_s_mkt', '_s_sec']) eq((await http(w, 'GET', '/api/manager/company-card-overdue')).status, 403, 'D22 非主管 403：' + w);
  // 逾期（>14 天）
  await sleep(5);
  const od1 = (await http('_s_mgr1', 'GET', '/api/manager/company-card-overdue')).json;
  const od2 = (await http('_s_mgr2', 'GET', '/api/manager/company-card-overdue')).json;
  eq(od1.overdue.map((x) => x.companyId), [], 'O1 c10 已改派給 u2 並被解除 → ERP 主管不再看到逾期');
  eq(od2.overdue.map((x) => x.companyId), [], 'O2 c8 已被 u3 解除 → 解除後從逾期清單消失（改出現在「業務已回報」）');
  const dayLabel = (await http('_s_exec', 'GET', '/api/manager/company-card-overdue')).json;
  eq(dayLabel.overdueDays, 14, 'O3 逾期天數門檻 14');
  eq((await http('_s_u1', 'GET', '/api/my-tasks/company-cards')).json.items.find((x) => x.companyId === 'c9').overdue, false, 'O4 c9 14 天又 1 小時：不算逾期（> 14 天才算）');
  // 稽核
  const dl = await auditLines('DISMISS_COMPANY_CARD_TASK');
  ok(dl.length === 4 && dl.every((l) => /原因：/.test(l.detail) && /已回報行銷與主管/.test(l.detail)), 'D23 每次解除都有稽核，記原因並註明已回報', dl.length);
  // 行銷處理
  const fbItems = (await http('_s_mkt', 'GET', '/api/marketing/responsibility-feedback')).json.items;
  const id11 = fbItems.find((x) => x.companyId === 'c11').id, id10 = fbItems.find((x) => x.companyId === 'c10').id, id8 = fbItems.find((x) => x.companyId === 'c8').id;
  const resolve = (who, id, body) => http(who, 'POST', `/api/marketing/responsibility-feedback/${id}/resolve`, body);
  eq((await resolve('_s_u1', id11, { action: 'handled' })).status, 403, 'R1 一般業務不能處理回報');
  eq((await resolve('_s_mgr1', id11, { action: 'handled' })).status, 403, 'R2 主管不能處理回報');
  eq((await resolve('_s_mkt', id11, { action: 'bogus' })).status, 400, 'R3 action 不合法 400');
  eq((await resolve('_s_mkt', 'ccd_nope', { action: 'handled' })).status, 404, 'R4 找不到回報 404');
  r = await resolve('_s_mkt', id11, { action: 'handled', note: '已電話確認' }); eq(r.status, 200, 'R5 標示已處理');
  eq((await http('_s_mkt', 'GET', '/api/marketing/responsibility-feedback')).json.items.some((x) => x.id === id11), false, 'R6 已處理的從清單移除');
  eq((await myTasks('_s_u2')).items.some((x) => x.companyId === 'c11'), false, 'R7 標示已處理不會讓待辦復活（解除維持有效）');
  eq((await resolve('_s_mkt', id11, { action: 'handled' })).status, 409, 'R8 重複處理 409');
  // 重新指定
  r = await resolve('_s_mkt', id10, { action: 'reassign', responsibleSales: [] }); eq(r.status, 400, 'R9 重新指定至少選一位');
  r = await resolve('_s_mkt', id10, { action: 'reassign', responsibleSales: ['_s_mkt'] }); eq(r.status, 400, 'R10 重新指定給非法帳號 400');
  eq((await http('_s_mkt', 'GET', '/api/marketing/responsibility-feedback')).json.items.some((x) => x.id === id10), true, 'R11 失敗的重新指定不會結案');
  await sleep(15);
  const nR = (await notifs('_s_u2', 'company_card_task')).length;
  r = await resolve('_s_mkt', id10, { action: 'reassign', responsibleSales: ['_s_u2'] }); eq([r.status, r.json && r.json.notifiedUsers], [200, 1], 'R12 重新指定（含同一位 u2）成功並通知');
  t = await myTasks('_s_u2');
  ok(t.items.some((x) => x.companyId === 'c10') && t.items.find((x) => x.companyId === 'c10').days === 0, 'R13 重新指定會重置該業務的解除：c10 待辦復活、天數歸零');
  eq((await notifs('_s_u2', 'company_card_task')).length - nR, 1, 'R14 重新指定多發了一則新通知給 u2');
  r = await resolve('_s_mkt', id8, { action: 'handled' });
  // 解除後業務自己補名片：回報清單自動消失
  r = await newCard('_s_u3', C.c4, { bu: 'ITS' }); eq(r.status, 201, 'R15 u3 補了 c4 的名片');
  eq((await http('_s_mkt', 'GET', '/api/marketing/responsibility-feedback')).json.items.some((x) => x.companyId === 'c4'), false, 'R16 解除後業務補了名片 → 回報不再成立、自動從行銷清單消失');
  const rl = await auditLines('RESOLVE_COMPANY_CARD_FEEDBACK');
  ok(rl.length >= 3, 'R17 處理回報有稽核', rl.length);
  // 通知欄位
  const nn = (await notifs('_s_u2', 'company_card_task'))[0];
  ok(nn && /請補名片/.test(nn.title) && nn.read === false, 'N1 通知 type=company_card_task、標題含「請補名片」、未讀');
  // 稽核內容無 HTML 轉義的問題：原樣文字
  const note10 = (readData().companyCardDismissals || []).find((x) => x.companyId === 'c10');
  ok(note10 && note10.note.includes('<img'), 'X1 酬載原樣以純文字儲存（輸出端負責跳脫）');
  await stopServer(); cleanupFiles();
}

// ═════════════ 匯入 ═════════════
async function groupImport() {
  console.log('[匯入] 服務業務欄 → 責任業務');
  await freshServer(null, null, { preset: (cos) => { const c1 = cos.find((c) => c.id === 'c1'); c1.responsibleSales = ['_s_u3']; c1.responsibleAssignedAt = { _s_u3: agoIso(2 * DAY) }; } });
  await loginAll();
  const rows = [
    ['匯入新公司一', '91000001', '', '測試業務甲、不存在的人甲', '', ''],
    ['匯入新公司二', '91000002', '', '測試業務甲', '', ''],
    ['匯入新公司三', '91000003', '', '_s_u2', '', ''],
    ['假公司丁', '90000004', '', '測試業務乙', '', ''],
    ['假公司甲', '90000001', '', '測試業務乙', '', ''],
    ['匯入新公司四', '91000004', '', '測試行銷', '', ''],
    ['匯入新公司五', '91000005', '', '離職業務', '', ''],
  ];
  const buf = xlsxBuf(rows);
  const dataBefore = fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8');
  let r = await upload('_s_mkt', '/api/companies/import?dryRun=1', buf);
  eq(r.status, 200, 'I1 預覽成功');
  eq([r.json.responsible.assigned, r.json.responsible.kept, r.json.responsible.notFound], [4, 1, 1], 'I2 預覽：已指定 4（甲甲乙乙）、沿用既有 1（不覆蓋）、找不到名稱 1');
  eq(r.json.responsible.notFoundNames, ['不存在的人甲'], 'I3 預覽列出找不到的名稱');
  ok(r.json.responsible.ineligible.includes('_s_off') && r.json.responsible.ineligible.includes('_s_mkt'), 'I4 預覽列出不能擁有名片者（停用／行銷）', JSON.stringify(r.json.responsible.ineligible));
  eq(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'), dataBefore, 'I5 預覽完全不寫資料');
  r = await upload('_s_mkt', '/api/companies/import', buf);
  eq(r.status, 200, 'I6 正式匯入成功');
  eq([r.json.responsible.assigned, r.json.responsible.kept, r.json.responsible.notFound], [4, 1, 1], 'I7 結果與預覽一致');
  let d = readData();
  const byName = (n) => d.companies.find((c) => c.name === n);
  eq(byName('匯入新公司一').responsibleSales, ['_s_u1'], 'I8 名稱解析成功 → 寫入 responsibleSales（不存在的名字略過）');
  eq(byName('匯入新公司三').responsibleSales, ['_s_u2'], 'I9 帳號也能解析');
  eq(byName('假公司丁').responsibleSales, ['_s_u2'], 'I10 既有公司責任業務空白 → 填入');
  eq(byName('假公司甲').responsibleSales, ['_s_u3'], 'I11 既有公司已有責任業務 → 預設不覆蓋');
  eq(byName('匯入新公司四').responsibleSales || [], [], 'I12 行銷不能當責任業務');
  eq(byName('匯入新公司五').responsibleSales || [], [], 'I13 停用帳號不能當責任業務');
  eq(byName('匯入新公司一').buImport, ['ERP'], 'I14 既有的 BU 累加行為保留');
  ok(byName('匯入新公司一').responsibleAssignedAt._s_u1, 'I15 有指定時間');
  eq((await notifs('_s_u1', 'company_card_task')).length, 1, 'I16 一次匯入多家（u1 兩家）只發一則通知');
  const n1 = (await notifs('_s_u1', 'company_card_task'))[0];
  ok(/2 家/.test(n1.title), 'I17 通知標題帶家數（2 家）', n1.title);
  eq((await notifs('_s_u2', 'company_card_task')).length, 1, 'I18 u2 兩家（匯入新公司三、假公司丁）也只發一則');
  eq(r.json.responsible.notifiedUsers, 2, 'I19 回應帶通知人數');
  eq((await notifs('_s_u3', 'company_card_task')).length, 0, 'I20 沒被新指定的人不通知');
  const al = await auditLines('SET_COMPANY_RESPONSIBLE');
  ok(al.some((l) => /匯入指定責任業務 4 家/.test(l.detail)), 'I21 匯入指定有稽核');
  // 覆蓋
  r = await upload('_s_mkt', '/api/companies/import?overwriteResponsible=1', buf);
  eq([r.json.responsible.assigned, r.json.responsible.kept, r.json.responsible.overwrite], [1, 0, true], 'I22 勾覆蓋：只有內容不同的那一家（假公司甲 u3→u2）算指定，其餘相同＝未變更');
  eq(readData().companies.find((c) => c.name === '假公司甲').responsibleSales, ['_s_u2'], 'I23 覆蓋後以匯入為準');
  eq((await notifs('_s_u2', 'company_card_task')).length, 2, 'I24 覆蓋後 u2 新被指定到假公司甲（名下沒有名片）→ 再收到一則通知（共 2 則）');
  // 範本說明
  const tpl = await http('_s_mkt', 'GET', '/api/companies/import-template');
  const wb = XLSX.read(tpl.buf, { type: 'buffer' });
  const guide = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[1]], { header: 1 });
  const svc = guide.find((x) => x[0] === '服務業務');
  ok(svc && /責任業務/.test(svc[2]) && !/名字在系統中找不到時會被忽略，不會報錯/.test(svc[2]), 'I25 範本說明已改成新語意（含責任業務；舊的「只累加 BU、找不到靜默忽略」說法不在）', svc && svc[2].slice(0, 80));
  eq(wb.SheetNames[0], '企業主檔匯入範本', 'I26 說明頁仍是第二張（匯入只讀第一張）');
  r = await upload('_s_u1', '/api/companies/import', buf); eq(r.status, 403, 'I27 一般業務不能匯入');
  await stopServer(); cleanupFiles();
}

// ═════════════ 開關：COMPANY_CARD_TASKS=off ═════════════
function exportBase() {
  if (process.env.CARDTASK_BASE_DIR) return { dir: process.env.CARDTASK_BASE_DIR, cleanup() {} };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cardtask-base-'));
  const a = spawnSync('git', ['archive', 'origin/main'], { cwd: ROOT, maxBuffer: 256 * 1024 * 1024 });
  if (a.status !== 0) throw new Error('git archive 失敗：' + String(a.stderr));
  const t = spawnSync('tar', ['-x', '-C', dir], { input: a.stdout, maxBuffer: 256 * 1024 * 1024 });
  if (t.status !== 0) throw new Error('tar 失敗：' + String(t.stderr));
  spawnSync('cmd', ['/c', 'mklink', '/J', path.join(dir, 'node_modules'), path.join(ROOT, 'node_modules')], { stdio: 'ignore' });
  return { dir, cleanup() { spawnSync('cmd', ['/c', 'rmdir', path.join(dir, 'node_modules')], { stdio: 'ignore' }); try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ } } };
}
async function groupOff() {
  console.log('[開關] COMPANY_CARD_TASKS=off ＝ 與沒有此功能時位元級相同');
  const gets = ['/api/companies', '/api/companies/c1', '/api/admin/companies', '/api/admin/companies/c1/detail', '/api/company-master?taxId=90000001'];
  const importRows = [['匯入新公司一', '91000001', '', '測試業務甲、不存在的人甲', '北區', ''], ['假公司丁', '90000004', '', '測試業務乙', '', ''], ['假公司甲', '90000001', '半導體', '測試業務丙', '', '']];
  const norm = (x) => String(x).replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/g, '<uuid>').replace(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/g, '<ts>');
  const run = async (base, root, port, env, label) => {
    await stopServer(); cleanupFiles(root); seedFiles({ dir: root, preset: (cos) => cos.forEach((c) => { delete c.responsibleSales; delete c.responsibleAssignedAt; }) });
    await startServer(null, env, root, port);
    const jr = {};
    for (const w of ['_s_adm', '_s_mkt', '_s_mgr1', '_s_u1']) await loginAt(base, jr, w);
    const out = {};
    for (const who of ['_s_adm', '_s_mkt', '_s_mgr1', '_s_u1']) for (const g of gets) { const rr = await httpAt(base, jr, who, 'GET', g); out[who + ' ' + g] = rr.status + ' ' + norm(rr.text); }
    const fdb = new FormData(); fdb.append('file', new Blob([xlsxBuf(importRows)]), 'i.xlsx');
    const dry = await httpAt(base, jr, '_s_mkt', 'POST', '/api/companies/import?dryRun=1&toPool=1', fdb);
    out.dry = dry.status + ' ' + norm(dry.text);
    const fdc = new FormData(); fdc.append('file', new Blob([xlsxBuf(importRows)]), 'i.xlsx');
    const real = await httpAt(base, jr, '_s_mkt', 'POST', '/api/companies/import?toPool=1', fdc);
    out.real = real.status + ' ' + norm(real.text);
    for (const who of ['_s_adm', '_s_mkt']) for (const g of gets.slice(0, 3)) { const rr = await httpAt(base, jr, who, 'GET', g); out['after ' + who + ' ' + g] = rr.status + ' ' + norm(rr.text); }
    const nt = await httpAt(base, jr, '_s_u1', 'GET', '/api/notifications');
    out.notif = JSON.stringify((nt.json || []).map((n) => n.type));
    const poolMsgs = JSON.parse(fs.readFileSync(path.join(root, 'data.json'), 'utf8')).contacts.filter((c) => c.owner === '_pool').map((c) => c.company).sort();
    out.pool = JSON.stringify(poolMsgs);
    const tpl = await httpAt(base, jr, '_s_mkt', 'GET', '/api/companies/import-template');
    const wb = XLSX.read(tpl.buf, { type: 'buffer' });
    out.tpl = JSON.stringify(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[1]], { header: 1 }));
    const raw = JSON.parse(fs.readFileSync(path.join(root, 'data.json'), 'utf8'));
    out.companyKeys = JSON.stringify(raw.companies.map((c) => Object.keys(c).sort().filter((k) => !['updatedAt', 'createdAt', 'id'].includes(k))));
    out.contactCount = String(raw.contacts.length);
    // 新增的端點（只在我方版本存在）：停用時行為
    const extra = {};
    extra.mine = await httpAt(base, jr, '_s_u1', 'GET', '/api/my-tasks/company-cards');
    extra.set = await httpAt(base, jr, '_s_mkt', 'POST', '/api/companies/c4/set-responsible-sales', { responsibleSales: ['_s_u1'] });
    extra.batch = await httpAt(base, jr, '_s_mkt', 'POST', '/api/companies/responsible-sales-batch', { companyIds: ['c4'], usernames: ['_s_u1'] });
    extra.dismiss = await httpAt(base, jr, '_s_u1', 'POST', '/api/my-tasks/company-cards/dismiss', { companyId: 'c4', reason: 'not_mine' });
    extra.fb = await httpAt(base, jr, '_s_mkt', 'GET', '/api/marketing/responsibility-feedback');
    extra.mgr = await httpAt(base, jr, '_s_mgr1', 'GET', '/api/manager/company-card-overdue');
    await stopServer(); cleanupFiles(root);
    return { out, extra };
  };
  const baseX = exportBase();
  let A, B;
  try {
    A = await run(`http://localhost:${PORT}`, baseX.dir, PORT, {}, '修改前');
    B = await run(`http://localhost:${PORT}`, ROOT, PORT, { COMPANY_CARD_TASKS: 'off' }, '修改後 off');
  } finally { baseX.cleanup(); }
  let diffKeys = Object.keys(A.out).filter((k) => A.out[k] !== B.out[k]);
  ok(diffKeys.length === 0, `F1 off：所有既有端點與匯入的輸出與「修改前程式碼」逐位元相同（比對 ${Object.keys(A.out).length} 項）`, diffKeys.join(', '));
  if (diffKeys.length) console.log('    差異首項：', diffKeys[0], '\n      前：', String(A.out[diffKeys[0]]).slice(0, 300), '\n      後：', String(B.out[diffKeys[0]]).slice(0, 300));
  eq(B.extra.mine.json, { enabled: false, items: [], total: 0, dismissed: 0, overdueDays: 14 }, 'F2 off：我的待辦回 enabled:false、空清單');
  eq([B.extra.set.status, B.extra.batch.status, B.extra.dismiss.status], [404, 404, 404], 'F3 off：寫入端點全部 404');
  eq([B.extra.fb.json.enabled, B.extra.mgr.json.enabled], [false, false], 'F4 off：行銷回報與主管逾期清單 enabled:false');
  // off 且資料已有責任業務欄位（例如曾經開啟過）：輸出完全不帶這些欄位、也不產生待辦
  await stopServer(); cleanupFiles();
  seedFiles({ preset: (cos) => { cos.find((c) => c.id === 'c4').responsibleSales = ['_s_u1']; cos.find((c) => c.id === 'c4').responsibleAssignedAt = { _s_u1: agoIso(20 * DAY) }; } });
  await startServer(null, { COMPANY_CARD_TASKS: 'off' }); jar = {}; await loginAll();
  for (const [who, u] of [['_s_adm', '/api/admin/companies'], ['_s_adm', '/api/admin/companies/c4/detail'], ['_s_mkt', '/api/companies'], ['_s_mkt', '/api/companies/c4']]) {
    const rr = await http(who, 'GET', u);
    ok(!/responsible/i.test(rr.text), 'F5 off＋資料已有欄位：輸出不含任何 responsible* 欄位 ' + u);
  }
  eq((await myTasks('_s_u1')).items, [], 'F6 off＋資料已有責任業務：不產生待辦');
  await stopServer(); cleanupFiles();
  // 預設（沒設環境變數）＝開
  await freshServer();
  await login('_s_u1'); await login('_s_mkt'); await login('_s_u2');
  eq((await myTasks('_s_u1')).enabled, true, 'F7 沒設環境變數＝開');
  const nobody = (await http('_s_mkt', 'GET', '/api/companies')).json;
  ok(nobody.every((c) => Array.isArray(c.responsibleSales) && Array.isArray(c.responsibleSalesDisplay)), 'F8 開啟且沒有任何責任業務資料時，只是多了空欄位、不觸發任何待辦');
  eq((await myTasks('_s_u2')).items.map((x) => x.companyId).sort(), ['c11'], 'F9 沒有責任業務資料的人不會有待辦（u2 只有 seed 的 c11）');
  // 通知網址
  const rel = spawnSync(process.execPath, ['-e', "const s=require('./server.js');console.log(s._test.notificationUrlFor('company_card_task',null))"], { cwd: ROOT, encoding: 'utf8', env: Object.assign({}, process.env, { DB_BACKEND: 'json', JWT_SECRET: 'x'.repeat(30), SESSION_BACKEND: 'jwt' }) });
  eq(String(rel.stdout).trim().split('\n').pop(), '/index.html#cardtasks', 'F10 推播通知網址 → /index.html#cardtasks（點擊會開彙總視窗）');
  await stopServer(); cleanupFiles();
}

// ═════════════ 變異測試 ═════════════
const L = 'lib/companyCardTasks.js';
const MUTANTS = [
  { name: 'M1 觸發條件改錯（有名片也算待辦）', groups: ['api', 'unit'], mut: [{ file: L, edits: [{ find: 'if (hasRealCard(index, m, un)) return;', rep: 'if (false) return;' }] }] },
  { name: 'M2 placeholder 被算進真實名片', groups: ['api', 'unit'], mut: [{ file: L, edits: [{ find: 'if (!c || c.deleted || c.isPlaceholder || !c.owner) return;', rep: 'if (!c || c.deleted || !c.owner) return;' }] }] },
  { name: 'M3 已刪除名片被算進去', groups: ['api', 'unit'], mut: [{ file: L, edits: [{ find: 'if (!c || c.deleted || c.isPlaceholder || !c.owner) return;', rep: 'if (!c || c.isPlaceholder || !c.owner) return;' }] }] },
  { name: 'M4 解除原因未強制', groups: ['api', 'unit'], mut: [{ file: L, edits: [{ find: "if (!Object.prototype.hasOwnProperty.call(DISMISS_REASONS, reason)) return { ok: false, status: 400, error: '請選擇解除原因' };", rep: '' }] }] },
  { name: 'M5 「其他」不必填說明', groups: ['api', 'unit'], mut: [{ file: L, edits: [{ find: "if (reason === 'other' && !cleaned)", rep: 'if (false)' }] }] },
  { name: 'M6 權限放寬：一般業務也能設定責任業務', groups: ['api'], mut: [{ file: L, edits: [{ find: "const WRITE_ROLES = ['admin', 'marketing', 'manager1', 'manager2', 'executive'];", rep: "const WRITE_ROLES = ['admin', 'marketing', 'manager1', 'manager2', 'executive', 'user', 'secretary'];" }] }] },
  { name: 'M7 權限放寬：主管可指定可視範圍外的人', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: 'const outside = sub.find(u => !viewable.has(u));', rep: 'const outside = undefined;' }] }] },
  { name: 'M8 _pool／不可擁有名片的帳號被允許', groups: ['api', 'unit'], mut: [{ file: L, edits: [{ find: "const OWNER_ROLES = ['user', 'manager1', 'manager2', 'groupsales'];", rep: "const OWNER_ROLES = ['user', 'manager1', 'manager2', 'groupsales', 'pool', 'marketing', 'admin', 'secretary'];" }, { find: "if (u.username === '_pool' || u.role === 'pool') return { ok: false, error: '不能指定客戶池為責任業務' };", rep: '' }, { find: "if (user.username === '_pool' || user.role === 'pool') return false;", rep: '' }, { find: 'if (user.active === false) return false;', rep: '' }] }] },
  { name: 'M9 主管首頁跨 BU 外洩（沒有依可視範圍過濾）', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: '.filter(t => visible.has(t.username) && !t.dismissal && t.days > cardTasks.OVERDUE_DAYS)', rep: '.filter(t => !t.dismissal && t.days > cardTasks.OVERDUE_DAYS)' }, { find: ".filter(f => visible.has(f.username)), usersByName, data)", rep: ', usersByName, data)' }] }] },
  { name: 'M10 重新指定沒有重置解除', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: 'cardTasks.applyResponsibleSales(m, list, nowIso, list);', rep: 'cardTasks.applyResponsibleSales(m, list, nowIso, []);' }] }] },
  { name: 'M11 匯入預設就覆蓋既有責任業務', groups: ['import'], mut: [{ file: 'server.js', edits: [{ find: 'if (hasExisting && !RS.overwrite) RS.kept++;', rep: 'if (false) RS.kept++;' }] }] },
  { name: 'M12 一次匯入多家公司每家各發一則通知', groups: ['import', 'api'], mut: [{ file: 'server.js', edits: [{ find: "try { pushNotification(un, 'company_card_task', title, body, null); n++; } catch { /* 通知失敗不影響指定 */ }", rep: "names.forEach(() => { try { pushNotification(un, 'company_card_task', title, body, null); n++; } catch { /* x */ } });" }] }] },
  { name: 'M13 開關失效（off 仍然啟用）', groups: ['off'], mut: [{ file: L, edits: [{ find: "return v !== 'off';", rep: 'return true;' }] }] },
  { name: 'M14 解除沒有綁定指定時間（舊解除永遠有效）', groups: ['api', 'unit'], mut: [{ file: L, edits: [{ find: "(d.assignedAt || '') === (assignedAt || '')", rep: 'true' }] }] },
  { name: 'M15 他人也能解除別人的待辦', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: "{ companyId: b.companyId, username: me, reason: b.reason, note: b.note }", rep: "{ companyId: b.companyId, username: (b.username || me), reason: b.reason, note: b.note }" }] }] },
];
async function runMutants(C2) {
  console.log('[變異] 每個變異都必須讓對應群組變紅');
  for (const m of MUTANTS) {
    if (process.env.CARDTASK_MUT_PICK && !new RegExp(process.env.CARDTASK_MUT_PICK).test(m.name)) continue;
    const snap = { pass, fail, n: failures.length };
    quiet = true;
    try {
      for (const g of m.groups) {
        if (g === 'unit') continue;   // 單元測試直接 require 磁碟上的檔案，不經預載腳本；改由 api 群組（真實 server）抓
        MUT_CUR = m.mut;
        try { await ({ api: groupApi, import: groupImport, off: groupOff })[g](); } catch (e) { fail++; failures.push('變異執行例外：' + e.message); }
      }
    } finally { quiet = false; MUT_CUR = null; await stopServer(); cleanupFiles(); }
    const newFails = failures.length - snap.n;
    pass = snap.pass; fail = snap.fail; failures.length = snap.n;
    ok(newFails > 0, `${m.name} → 測試變紅（${newFails} 項失敗）`);
  }
}
let MUT_CUR = null;
const _freshServer0 = freshServer;
// 變異模式：freshServer／startServer 一律帶入目前的變異規格
freshServer = async function (mut, extraEnv, seedOpts) { return _freshServer0(mut || MUT_CUR, extraEnv, seedOpts); };
const _startServer0 = startServer;
startServer = async function (mut, extraEnv, root, port) { return _startServer0(mut || (root && root !== ROOT ? null : MUT_CUR), extraEnv, root, port); };

const exportsForUi = { ok, eq, sleep, http, login, loginAll, freshServer: (...a) => freshServer(...a), stopServer, cleanupFiles, readData, newCard, myTasks, setResp, auditLines, notifs, C, XSS, BASE, PW, ROOT, SHOTS, CDP, seedFiles, startServer: (...a) => startServer(...a), snap: () => ({ pass, fail, n: failures.length }), restore: (x) => { pass = x.pass; fail = x.fail; failures.length = x.n; }, newFails: (x) => failures.slice(x.n), setQuiet: (q) => { quiet = q; }, setMut: (m) => { MUT_CUR = m; }, agoIso, DAY, HOUR };
module.exports = exportsForUi;
if (require.main === module) {
  (async () => {
    try {
      if (want('unit')) groupUnit();
      if (want('api')) await groupApi();
      if (want('import')) await groupImport();
      if (want('off')) await groupOff();
      if (want('ui')) await require('./_company-card-tasks-ui.js').run(exportsForUi);
      if (process.env.CARDTASK_MUT === '1') {
        await runMutants(exportsForUi);
        if (want('ui') || ONLY.includes('ui')) await require('./_company-card-tasks-ui.js').mutants(exportsForUi);
      }
    } catch (e) {
      fail++; failures.push('例外：' + (e && e.stack || e)); console.log('例外：', e && e.stack || e);
    } finally {
      await stopServer(); cleanupFiles();
      try { fs.unlinkSync(PRELOAD); } catch { /* 不存在 */ }
    }
    console.log(`\n企業主檔責任業務／待補名片待辦檢查：PASS ${pass} / FAIL ${fail}`);
    if (fail) console.log('失敗項目：\n  ' + failures.join('\n  '));
    process.exit(fail ? 1 : 0);
  })();
}
