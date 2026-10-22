#!/usr/bin/env node
'use strict';
/**
 * 統編查無 → 申請行銷建檔 檢查：單元（lib/companyRequests.js）＋ API（真實 server.js）＋ 開關關閉位元級不變
 * ＋ 前端互動（headless Edge + CDP，見 _company-requests-ui.js）＋ 變異測試。
 * 不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體；資料檔只存在於本次執行；結束以 PID 清掉伺服器與 Edge；不連遠端資料庫；全部是假公司、假統編。
 *
 *   node scripts/check-company-requests.js                  單元＋API＋開關（不含瀏覽器）
 *   CREQ_ONLY=unit,api,off,ui node ...                      只跑指定群組（ui 需要 Edge）
 *   CREQ_MUT=1 node ...                                     另外跑變異測試（每個變異都必須讓對應群組變紅）
 *   CREQ_SHOTS=<資料夾> node ...                            另外輸出截圖（ui 群組）
 *   CREQ_PORT（預設 3200）／CREQ_CDP（預設 9500）／CREQ_BASE_DIR（開關比對用的「修改前」程式碼資料夾；沒給就用 git archive 匯出 CREQ_BASE_REF 到暫存；
 *     預設 425fc4f＝尚未有此功能的 main（PR #43 合併後），不是 origin/main）
 * 守門：工作樹已有 auth.json／data.json／audit.log.json（可能是真實資料）時拒絕執行。埠限制 3200–3209（CDP 9500–9509）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.CREQ_PORT || '3200', 10);
const CDP = parseInt(process.env.CREQ_CDP || '9500', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3199)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map((f) => path.join(ROOT, f));
{
  const g = FILES.filter((f) => /(auth|data|audit.log)\.jsonl?$/.test(f)).filter((f) => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map((f) => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
}
const ONLY = (process.env.CREQ_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length ? k !== 'ui' : ONLY.includes(k);
const SHOTS = process.env.CREQ_SHOTS || '';
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const CR = require(path.join(ROOT, 'lib', 'companyRequests.js'));

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 300)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = (n) => new Promise((r) => setTimeout(r, n));

// ───────────────── 沙盒資料（全部假的）─────────────────
const PW = 'Cr' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const NOW = Date.now();
const nowIso = new Date(NOW).toISOString();
const DAY = 86400000;
const agoIso = (ms) => new Date(NOW - ms).toISOString();
const XSS = '<img src=x onerror="window.__xss=1">';
const TAX = '08628407';        // 假想的政府機關統編（測試用，不連任何外部服務）
const OFFICIAL = '假衛生福利部中央健康保險署';
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false }, extra || {});
const mkCo = (id, name, tax, extra) => Object.assign({ id, taxId: tax || '', matchKey: tax ? 'tax:' + tax : 'name:' + name, name, shortName: '', capital: null, industry: '', address: '', phone: '', website: '', region: '', customerCode: '', buImport: [], groupId: null, gcisEnriched: false, source: 'manual', createdAt: nowIso, updatedAt: nowIso }, extra || {});
const mkCard = (id, owner, company, tax, extra) => Object.assign({ id, owner, name: '名片' + id, company, companyId: '', taxId: tax || '', bu: 'ERP', email: id + '@example.test', createdAt: nowIso }, extra || {});
function seedUsers() {
  return [
    mkUser('_s_adm', 'admin'), mkUser('_s_mkt', 'marketing', { bu: ['ERP'], displayName: '測試行銷' }), mkUser('_s_mkt2', 'marketing', { displayName: '測試行銷二' }),
    mkUser('_s_mktoff', 'marketing', { displayName: '離職行銷', active: false }),
    mkUser('_s_mgr1', 'manager1', { bu: ['ERP'], displayName: '測試主管' }), mkUser('_s_sec', 'secretary', { bu: ['ERP'] }),
    mkUser('_s_u1', 'user', { bu: ['ERP'], supervisor: '_s_mgr1', displayName: '測試業務甲' }),
    mkUser('_s_u2', 'user', { bu: ['ERP'], supervisor: '_s_mgr1', displayName: '測試業務乙' }),
    mkUser('_s_u3', 'user', { bu: ['ERP'], supervisor: '_s_mgr1', displayName: '測試業務丙' }),
    mkUser('_s_gs', 'groupsales', { bu: ['ERP'], displayName: '測試集團業務' }), mkUser('_s_tp', 'tecopm', { displayName: '測試集團PM' }),
  ];
}
const CO_EXIST = mkCo('cm1', '假公司甲股份有限公司', '90000001');
const CO_SAP = mkCo('cm2', '假公司SAP股份有限公司', '90000099');
const CO_SAP_OK = mkCo('cm3', '假公司正常股份有限公司', '90000098');
function seedData(opts) {
  opts = opts || {};
  const d = {
    contacts: [
      mkCard('kok', '_s_u1', CO_EXIST.name, CO_EXIST.taxId, { companyId: CO_EXIST.id }),
      mkCard('ksap', '_s_u1', CO_SAP.name, CO_SAP.taxId, { companyId: CO_SAP.id, companyVerify: 'pending', companyRequestId: 'seedreq-sap' }),
    ],
    companies: [JSON.parse(JSON.stringify(CO_EXIST)), JSON.parse(JSON.stringify(CO_SAP)), JSON.parse(JSON.stringify(CO_SAP_OK))],
    companyRequests: [{ id: 'seedreq-sap', taxId: CO_SAP.taxId, companyName: CO_SAP.name, category: 'gov', note: '種子資料用申請', requestedBy: '_s_u1', requestedAt: nowIso, contactId: 'ksap', contactIds: ['ksap'],
      requesters: [{ username: '_s_u1', contactId: 'ksap', companyName: CO_SAP.name, category: 'gov', note: '種子資料用申請', at: nowIso }], status: 'pending', resolvedBy: '', resolvedAt: '', result: null, rejectReason: '', rejectNote: '', cancelledBy: '', lastRemindedDate: nowIso.slice(0, 10) }],
    quotations: [
      { id: 'q-ok', owner: '_s_u1', quoteNo: 'QT-OK', contactId: 'kok', company: CO_EXIST.name, status: 'draft', products: [], items: [], approval: null, createdAt: nowIso, updatedAt: nowIso },
      { id: 'q-by-name', owner: '_s_u1', quoteNo: 'QT-NAME', contactId: '', company: '北區健保署', status: 'draft', products: [], items: [], approval: null, createdAt: nowIso, updatedAt: nowIso },
      { id: 'q-nolink', owner: '_s_u1', quoteNo: 'QT-NOLINK', contactId: '', company: '完全無關的公司', status: 'draft', products: [], items: [], approval: null, createdAt: nowIso, updatedAt: nowIso },
      { id: 'q-by-id', owner: '_s_u1', quoteNo: 'QT-ID', contactId: 'PENDING_ID', company: '無關名稱', status: 'draft', products: [], items: [], approval: null, createdAt: nowIso, updatedAt: nowIso },
    ],
    visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], opportunities: [], industries: ['政府機關'],
  };
  if (opts.data) opts.data(d);
  return d;
}
function seedFiles(opts) {
  opts = opts || {};
  const DIR = opts.dir || ROOT;
  const users = seedUsers();
  if (opts.users) opts.users(users);
  fs.writeFileSync(path.join(DIR, 'auth.json'), JSON.stringify({ users }, null, 2));
  fs.writeFileSync(path.join(DIR, 'data.json'), JSON.stringify(seedData(opts), null, 2));
}
const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
function cleanupFiles(dir) { FILES.forEach((f) => { try { fs.unlinkSync(dir ? path.join(dir, path.basename(f)) : f); } catch { /* 不存在 */ } }); }

// ───────────────── 伺服器（預載腳本可在載入時改寫指定檔案＝變異測試，不動磁碟上的檔案）─────────────────
const PRELOAD = path.join(os.tmpdir(), 'creq_preload_' + process.pid + '.js');
fs.writeFileSync(PRELOAD, [
  "const Module=require('module'),path=require('path');const orig=Module.prototype._compile;",
  "Module.prototype._compile=function(content,filename){const raw=process.env.CREQ_MUT_SPEC;",
  "if(raw){const specs=JSON.parse(raw);for(const m of specs){if(path.resolve(filename)===path.resolve(process.env.CREQ_ROOT,m.file)){for(const e of m.edits){if(!content.includes(e.find))throw new Error('MUTATION ANCHOR MISSING: '+e.find.slice(0,60));content=content.replace(e.find,()=>e.rep);}}}}",
  "return orig.call(this,content,filename);};",
].join('\n') + '\n');
let child = null, childPid = null;
async function startServer(mut, extraEnv, root, port) {
  root = root || ROOT; port = port || PORT;
  const env = Object.assign({}, process.env, { PORT: String(port), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: `http://localhost:${port}` });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'QUOTE_FLOW_ENGINE', 'QUOTE_FLOW_DESIGNER', 'FEATURE_GUARD_MODE', 'CREQ_MUT_SPEC', 'COMPANY_REQUESTS']) delete env[k];
  Object.assign(env, extraEnv || {});
  if (mut && mut.length) { env.CREQ_MUT_SPEC = JSON.stringify(mut); env.CREQ_ROOT = ROOT; }
  await assertPortFree(port);
  child = spawn(process.execPath, ['-r', PRELOAD, path.join(root, 'server.js')], { cwd: root, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
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
  if (body !== undefined) h['Content-Type'] = 'application/json';
  if (who && jr[who]) h.Cookie = jr[who];
  const res = await fetch(base + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: res.status, json, text: buf.toString('utf8') };
}
const http = (who, method, url, body) => httpAt(BASE, jar, who, method, url, body);
async function loginAt(base, jr, who) {
  const res = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status === 200) jr[who] = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
  return res.status;
}
const login = (who) => loginAt(BASE, jar, who);
const ALL = ['_s_adm', '_s_mkt', '_s_mkt2', '_s_mktoff', '_s_mgr1', '_s_sec', '_s_u1', '_s_u2', '_s_u3', '_s_gs', '_s_tp'];
async function loginAll() { for (const w of ALL) await login(w); }
async function auditLines(action) {
  await sleep(300);
  const r = await http('_s_adm', 'GET', '/api/admin/logs');
  return (Array.isArray(r.json) ? r.json : []).filter((l) => l.action === action);
}
const notifs = async (who, type) => { const j = (await http(who, 'GET', '/api/notifications')).json; return (Array.isArray(j) ? j : []).filter((n) => n.type === type); };
const reqList = async (who) => (await http(who, 'GET', '/api/company-requests')).json;
const newCard = (who, company, extra) => http(who, 'POST', '/api/contacts', Object.assign({ name: '新名片', company, taxId: TAX, bu: 'ERP', email: 'a@example.test' }, extra || {}));
const RQ = (category, note) => ({ category: category || 'gov', note: note || '衛福部健保署北區業務組' });

// ═════════════ 單元：lib/companyRequests.js ═════════════
function groupUnit() {
  const CR = CR_REF.mod;
  console.log('[單元] lib/companyRequests.js');
  eq(CR.enabled({}), true, 'U1 沒設環境變數＝開');
  eq(CR.enabled({ COMPANY_REQUESTS: 'off' }), false, 'U2 off＝關');
  eq(CR.enabled({ COMPANY_REQUESTS: ' OFF ' }), false, 'U3 大小寫與空白不影響');
  eq(CR.enabled({ COMPANY_REQUESTS: 'on' }), true, 'U4 on＝開');
  eq(CR.CATEGORIES.map((c) => c.label), ['政府機關', '學校', '公立醫院', '財團法人', '協會團體', '行號', '其他'], 'U5 申請類別 7 類（決定 1）');
  eq(CR.REJECT_REASONS.map((c) => c.label), ['統編疑似輸入錯誤', '不需另建檔', '資料不足'], 'U6 退回原因 3 類');
  eq(CR.validateRequest({ category: 'gov', note: '衛福部健保署北區業務組' }).ok, true, 'U7 合法申請');
  eq(CR.validateRequest({ category: 'gov', note: '1234' }).code, 'NOTE_TOO_SHORT', 'U8 說明 4 字被拒');
  eq(CR.validateRequest({ category: 'gov', note: '12345' }).ok, true, 'U9 說明 5 字通過');
  eq(CR.validateRequest({ category: 'gov', note: '字'.repeat(200) }).ok, true, 'U10 200 字通過');
  eq(CR.validateRequest({ category: 'gov', note: '字'.repeat(201) }).code, 'NOTE_TOO_LONG', 'U11 201 字被拒');
  eq(CR.validateRequest({ category: 'gov', note: '  \u0000  ab  ' }).code, 'NOTE_TOO_SHORT', 'U12 空白與控制字元不算字數');
  eq(CR.validateRequest({ category: 'xx', note: '足夠長的說明' }).code, 'BAD_CATEGORY', 'U13 不在清單的類別被拒');
  eq(CR.validateRequest({ category: '', note: '足夠長的說明' }).code, 'BAD_CATEGORY', 'U14 沒選類別被拒');
  eq(CR.validateRequest('字串').code, 'BAD_REQUEST', 'U15 非物件被拒');
  eq(CR.validateRequest(null).code, 'BAD_REQUEST', 'U16 null 被拒');
  eq(CR.validateReject({ reason: '', note: '' }).code, 'REASON_REQUIRED', 'U17 退回沒原因被拒');
  eq(CR.validateReject({ reason: 'xx' }).code, 'REASON_REQUIRED', 'U18 退回原因不在清單被拒');
  eq(CR.validateReject({ reason: 'typo', note: ' 備註 ' }).value, { reason: 'typo', note: '備註' }, 'U19 退回原因＋備註');
  // 提交／併單
  const data = { contacts: [], companies: [], companyRequests: [] };
  let gid = 0; const genId = () => 'r' + (++gid);
  const c1 = { id: 'c1', owner: 'a', company: '北區', taxId: '11111111' }, c2 = { id: 'c2', owner: 'b', company: '中區', taxId: '11111111' }, c3 = { id: 'c3', owner: 'a', company: '別家', taxId: '22222222' };
  const r1 = CR.submitRequest(data, { contact: c1, user: { username: 'a' }, request: { category: 'gov', note: '說明一二三' }, now: nowIso, genId });
  eq([r1.merged, r1.request.id, c1.companyVerify, c1.companyRequestId], [false, 'r1', 'pending', 'r1'], 'U20 首次申請：新建＋名片標待確認');
  const r2 = CR.submitRequest(data, { contact: c2, user: { username: 'b' }, request: { category: 'school', note: '另一個說明' }, now: nowIso, genId });
  eq([r2.merged, r2.request.id, data.companyRequests.length, r2.request.requesters.length, r2.request.contactIds], [true, 'r1', 1, 2, ['c1', 'c2']], 'U21 同統編併單：不重複建立、記錄多位申請人');
  CR.submitRequest(data, { contact: c2, user: { username: 'b' }, request: { category: 'school', note: '改過的說明' }, now: nowIso, genId });
  eq([data.companyRequests[0].requesters.length, data.companyRequests[0].contactIds.length], [2, 2], 'U22 同一張名片重送不會重複計入申請人');
  const r3 = CR.submitRequest(data, { contact: c3, user: { username: 'a' }, request: { category: 'gov', note: '另一家說明' }, now: nowIso, genId });
  eq([r3.merged, data.companyRequests.length], [false, 2], 'U23 不同統編不併單');
  // 取消
  eq(CR.cancelRequest(data, 'r1', 'zzz', nowIso).status, 403, 'U24 非申請人不能取消');
  eq(CR.cancelRequest(data, 'nope', 'a', nowIso).status, 404, 'U25 找不到申請');
  data.contacts = [c1, c2, c3];
  const cc = CR.cancelRequest(data, 'r1', 'b', nowIso);
  eq([cc.ok, cc.fullyCancelled, data.companyRequests.find((x) => x.id === 'r1').status, c2.companyRequestId, c2.companyVerify], [true, false, 'pending', '', 'pending'], 'U26 一位申請人取消：申請仍在、名片保留待確認');
  const cc2 = CR.cancelRequest(data, 'r1', 'a', nowIso);
  eq([cc2.fullyCancelled, data.companyRequests.find((x) => x.id === 'r1').status], [true, 'cancelled'], 'U27 全部申請人取消：申請作廢');
  eq(CR.cancelRequest(data, 'r1', 'a', nowIso).status, 403, 'U28 已作廢後申請人名單已空，不能再取消');
  eq(CR.findPendingByTax(data, '11111111'), null, 'U29 作廢後同統編可再次申請（找不到待處理）');
  // detach
  const d2 = { contacts: [], companyRequests: [] };
  const x1 = { id: 'x1', owner: 'a', company: 'A', taxId: '33333333' }, x2 = { id: 'x2', owner: 'b', company: 'B', taxId: '33333333' };
  CR.submitRequest(d2, { contact: x1, user: { username: 'a' }, request: { category: 'gov', note: '說明一二三' }, now: nowIso, genId });
  CR.submitRequest(d2, { contact: x2, user: { username: 'b' }, request: { category: 'gov', note: '說明一二三' }, now: nowIso, genId });
  eq(CR.detachContact(d2, x1, nowIso, 'a'), null, 'U30 還有別人時 detach 不作廢申請');
  ok(CR.detachContact(d2, x2, nowIso, 'b') !== null, 'U31 最後一位走光 → 申請作廢');
  // 核可／退回
  const d3 = { contacts: [], companyRequests: [] };
  const y1 = { id: 'y1', owner: 'a', company: 'A', taxId: '44444444' }, y2 = { id: 'y2', owner: 'b', company: 'B', taxId: '44444444' }, y3 = { id: 'y3', owner: 'c', company: 'C', taxId: '44444444', companyVerify: 'pending', companyRequestId: '' }, y4 = { id: 'y4', owner: 'c', company: 'D', taxId: '55555555', companyVerify: 'pending' };
  d3.contacts = [y1, y2, y3, y4];
  CR.submitRequest(d3, { contact: y1, user: { username: 'a' }, request: { category: 'gov', note: '說明一二三' }, now: nowIso, genId });
  CR.submitRequest(d3, { contact: y2, user: { username: 'b' }, request: { category: 'gov', note: '說明一二三' }, now: nowIso, genId });
  const req3 = d3.companyRequests[0];
  const ap = CR.applyApproval(d3, req3, { id: 'M1' }, 'mkt', nowIso, { name: '正式名稱' });
  eq(ap.touched.sort(), ['y1', 'y2', 'y3'], 'U32 核可：同統編全部名片（含已取消者）清掉待確認，其他統編不動');
  eq([y1.companyVerify, y1.companyId, y4.companyVerify, req3.status, req3.result.companyId], [undefined, 'M1', 'pending', 'approved', 'M1'], 'U33 核可：掛上主檔、狀態 approved');
  eq(ap.notify.map((x) => x.username).sort(), ['a', 'b'], 'U34 核可：通知所有申請人（去重）');
  const d4 = { contacts: [], companyRequests: [] };
  const z1 = { id: 'z1', owner: 'a', company: 'A', taxId: '66666666' };
  d4.contacts = [z1];
  CR.submitRequest(d4, { contact: z1, user: { username: 'a' }, request: { category: 'gov', note: '說明一二三' }, now: nowIso, genId });
  const rj = CR.applyReject(d4, d4.companyRequests[0], 'mkt', { reason: 'typo', note: '看起來少一碼' }, nowIso);
  eq([z1.companyVerify, z1.companyVerifyReason, d4.companyRequests[0].status], ['rejected', '統編疑似輸入錯誤：看起來少一碼', 'rejected'], 'U35 退回：名片 rejected＋原因文字');
  eq(rj.notify.length, 1, 'U36 退回：通知申請人');
  // 閘門
  const g = { contacts: [{ id: 'p', owner: 'a', company: 'X', taxId: '1', companyVerify: 'pending' }, { id: 'r', owner: 'a', company: 'Y', taxId: '2', companyVerify: 'rejected' }, { id: 'n', owner: 'a', company: 'Z', taxId: '3' }, { id: 'd', owner: 'a', company: 'W', deleted: true, companyVerify: 'pending' }], opportunities: [{ id: 'o1', contactId: 'p' }], companyRequests: [] };
  eq((CR.blockForQuote(g, { owner: 'a', contactId: 'p' }) || {}).code, 'COMPANY_UNVERIFIED', 'U37 報價單 contactId 指到待確認名片 → 擋');
  eq((CR.blockForQuote(g, { owner: 'a', contactId: 'r' }) || {}).code, 'COMPANY_UNVERIFIED', 'U38 被退回也擋');
  ok(/退回/.test(CR.blockForQuote(g, { owner: 'a', contactId: 'r' }).message) && /等行銷確認/.test(CR.blockForQuote(g, { owner: 'a', contactId: 'p' }).message), 'U39 訊息分「等行銷確認」與「被退回」');
  eq(CR.blockForQuote(g, { owner: 'a', contactId: 'n' }), null, 'U40 正常名片不擋');
  eq((CR.blockForQuote(g, { owner: 'a', oppId: 'o1' }) || {}).code, 'COMPANY_UNVERIFIED', 'U41 經商機聯絡人對應 → 擋');
  eq((CR.blockForQuote(g, { owner: 'a', company: 'X' }) || {}).code, 'COMPANY_UNVERIFIED', 'U42 同擁有者同公司名對應 → 擋');
  eq(CR.blockForQuote(g, { owner: 'b', company: 'X' }), null, 'U43 別人的同名公司不擋');
  eq(CR.blockForQuote(g, { owner: 'a', company: 'W' }), null, 'U44 已刪除名片不算');
  eq(CR.blockForQuote(g, { owner: 'a', company: '無關' }), null, 'U45 找不到關聯 → 不擋');
  eq(CR.blockForQuote(g, null), null, 'U46 空物件不擋');
  const sg = { contacts: [{ id: 'p', taxId: '1', companyId: 'M', companyVerify: 'pending' }], companyRequests: [{ status: 'pending', taxId: '9' }] };
  eq((CR.blockForSapCompanies(sg, [{ id: 'M', name: 'M公司', taxId: '1' }]) || {}).code, 'COMPANY_UNVERIFIED', 'U47 SAP：同統編名片待確認 → 擋');
  eq((CR.blockForSapCompanies(sg, [{ id: 'Q', name: 'Q公司', taxId: '9' }]) || {}).names, ['Q公司'], 'U48 SAP：該統編有待處理申請 → 擋');
  eq(CR.blockForSapCompanies(sg, [{ id: 'Z', name: 'Z', taxId: '7' }]), null, 'U49 SAP：無關公司不擋');
  // 逾期／提醒
  const dd = { companyRequests: [
    { id: 'a', status: 'pending', requestedAt: agoIso(2 * DAY), taxId: '1', lastRemindedDate: '' },
    { id: 'b', status: 'pending', requestedAt: agoIso(3 * DAY + 1000), taxId: '2', lastRemindedDate: '' },
    { id: 'c', status: 'pending', requestedAt: agoIso(9 * DAY), taxId: '3', lastRemindedDate: '2026-10-22' },
    { id: 'd', status: 'approved', requestedAt: agoIso(9 * DAY), taxId: '4', lastRemindedDate: '' },
  ] };
  eq(CR.dueReminders(dd, NOW, '2026-10-22').map((r) => r.id), ['b'], 'U50 提醒：滿 3 天、今天沒提醒過、仍待處理者');
  eq(CR.dueReminders(dd, NOW, '2026-10-23').map((r) => r.id), ['b', 'c'], 'U51 隔天可再提醒（每天最多一次）');
  const viewer = { username: 'x', canResolve: true };
  eq(CR.serialize({ requestedAt: agoIso(8 * DAY), status: 'pending', requesters: [] }, NOW, viewer).overdue, true, 'U52 超過 7 天標紅');
  eq(CR.serialize({ requestedAt: agoIso(7 * DAY), status: 'pending', requesters: [] }, NOW, viewer).overdue, false, 'U53 剛好 7 天不標紅');
  eq(CR.serialize({ requestedAt: agoIso(9 * DAY), status: 'approved', requesters: [] }, NOW, viewer).overdue, false, 'U54 已結案不標紅');
  const rq = { id: 'q', requestedAt: nowIso, status: 'pending', requestedBy: 'a', requesters: [{ username: 'a', note: 'A的' }, { username: 'b', note: 'B的' }] };
  eq(CR.serialize(rq, NOW, { username: 'a', canResolve: false }).requesters.map((x) => x.username), ['a'], 'U55 申請人只看得到自己那份資料');
  eq(CR.serialize(rq, NOW, { username: 'm', canResolve: true }).requesters.length, 2, 'U56 行銷看得到全部申請人');
  eq(CR.visibleList({ companyRequests: [rq] }, { username: 'zzz', canResolve: false }, NOW).length, 0, 'U57 無關的人看不到');
  eq(CR.masterExists({ companies: [{ matchKey: 'tax:12345678' }] }, '12345678'), true, 'U58 masterExists 以 matchKey 判斷');
  eq(CR.masterExists({ companies: [] }, 'abc'), false, 'U59 非 8 碼不算');
}

// ═════════════ API ═════════════
async function groupApi() {
  console.log('[API] 真實 server.js（沙盒）');
  await freshServer(MUT_CUR, null, { data: (d) => { d.contacts.push(mkCard('kqid', '_s_u1', '子單位', '08000001', { companyVerify: 'pending', companyRequestId: '' })); d.quotations.find((q) => q.id === 'q-by-id').contactId = 'kqid'; } });
  await loginAll();
  const countCards = () => readData().contacts.length;
  const m0 = (await notifs('_s_mkt', 'company_request_new')).length;

  // ── 建立 ──
  let r = await newCard('_s_u1', '北區健保署', { companyRequest: RQ() });
  eq(r.status, 201, 'A1 帶 companyRequest：名片建立成功');
  const k1 = r.json;
  eq([k1.companyVerify, !!k1.companyRequestId, !!(k1.companyRequest && k1.companyRequest.id), k1.companyRequest && k1.companyRequest.merged], ['pending', true, true, false], 'A2 名片標「待確認」並綁定申請（新建，非併單）');
  let d = readData();
  eq(d.companyRequests.filter((x) => x.taxId === TAX).length, 1, 'A3 產生一筆申請');
  const rq1 = d.companyRequests.find((x) => x.taxId === TAX);
  eq([rq1.status, rq1.category, rq1.note, rq1.requestedBy, rq1.companyName, rq1.contactId], ['pending', 'gov', '衛福部健保署北區業務組', '_s_u1', '北區健保署', k1.id], 'A4 申請內容完整');
  ok(!d.companies.some((c) => c.taxId === TAX || c.matchKey === 'tax:' + TAX), 'A5 申請期間「不建」企業主檔（不讓子單位名稱佔位）');
  ok(!d.contacts.find((c) => c.id === k1.id).companyId, 'A6 名片沒有 companyId（等核可後掛上）');
  let ns = await notifs('_s_mkt', 'company_request_new');
  eq(ns.length - m0, 1, 'A7 行銷收到 company_request_new');
  eq((await notifs('_s_mkt2', 'company_request_new')).length, 1, 'A8 每位在職行銷都收到');
  eq((await notifs('_s_mktoff', 'company_request_new')).length, 0, 'A9 停用的行銷不通知');
  eq([(await notifs('_s_adm', 'company_request_new')).length, (await notifs('_s_u1', 'company_request_new')).length], [0, 0], 'A10 管理員與申請人本人不收新申請通知');
  ok(/北區健保署/.test(ns[0].body) && /08628407/.test(ns[0].body), 'A11 通知內容含公司名稱與統編');
  let lg = await auditLines('COMPANY_REQUEST_CREATE');
  ok(lg.length === 1 && /新建申請/.test(lg[0].detail) && lg[0].target === TAX, 'A12 稽核 COMPANY_REQUEST_CREATE（新建申請）', JSON.stringify(lg[0] || {}));

  // ── 驗證 ──
  const before = countCards();
  const bads = [[{ note: '衛福部健保署北區業務組' }, 'BAD_CATEGORY'], [{ category: 'zzz', note: '衛福部健保署北區業務組' }, 'BAD_CATEGORY'], [{ category: 'gov', note: '1234' }, 'NOTE_TOO_SHORT'], [{ category: 'gov', note: '字'.repeat(201) }, 'NOTE_TOO_LONG'], ['字串', 'BAD_REQUEST'], [[], 'BAD_REQUEST']];
  for (const [body, code] of bads) { r = await newCard('_s_u3', '南區健保署', { companyRequest: body }); ok(r.status === 400 && r.json.code === code, 'A13 無效申請被拒 ' + code + '：' + JSON.stringify(body).slice(0, 30), r.status + JSON.stringify(r.json)); }
  eq(countCards(), before, 'A14 申請驗證失敗時，名片也不建立');
  // 境外／已有主檔：忽略申請
  r = await newCard('_s_u3', CO_EXIST.name, { taxId: CO_EXIST.taxId, companyRequest: RQ() });
  eq([r.status, r.json.companyVerify, r.json.companyRequest], [201, undefined, undefined], 'A15 統編已有內部主檔：companyRequest 忽略，名片正常（不需申請）');
  r = await http('_s_u3', 'POST', '/api/contacts', { name: '境外', company: 'Foreign Co', isForeign: true, taxId: '', companyRequest: RQ(), bu: 'ERP', email: 'f@example.test' });
  eq([r.status, r.json.companyVerify], [201, undefined], 'A16 境外公司忽略 companyRequest');
  // 沒帶 → 與現在相同
  const dBefore = readData().companyRequests.length;
  r = await newCard('_s_u3', '一般公司', { taxId: '09999999' });
  eq([r.status, r.json.companyVerify, r.json.companyRequest, r.json.newCompany], [201, undefined, undefined, true], 'A17 沒帶 companyRequest：行為與現在相同（建主檔、全新客戶、沒有待確認）');
  eq(readData().companyRequests.length, dBefore, 'A18 沒帶 companyRequest 不產生申請');
  ok(readData().companies.some((c) => c.matchKey === 'tax:09999999'), 'A19 沒帶時照舊建立企業主檔');

  // ── 併單 ──
  r = await newCard('_s_u2', '中區健保署', { companyRequest: RQ('school', '中區業務組也要用') });
  eq([r.status, r.json.companyRequest && r.json.companyRequest.merged, r.json.companyRequest && r.json.companyRequest.id], [201, true, rq1.id], 'A20 同統編第二位申請：併入同一筆');
  const k2 = r.json;
  d = readData();
  const m1 = d.companyRequests.find((x) => x.id === rq1.id);
  eq([d.companyRequests.filter((x) => x.taxId === TAX).length, m1.requesters.map((x) => x.username).sort(), m1.contactIds.length], [1, ['_s_u1', '_s_u2'], 2], 'A21 仍只有一筆申請、記錄兩位申請人');
  eq((await notifs('_s_mkt', 'company_request_new')).length - m0, 1, 'A22 併單不重複通知行銷');
  lg = await auditLines('COMPANY_REQUEST_CREATE');
  ok(lg.length === 2 && lg.some((l) => /併入既有申請/.test(l.detail)), 'A23 併單也寫稽核（併入既有申請）');
  ok(!readData().companies.some((c) => c.taxId === TAX), 'A24 併單後仍沒有主檔；子單位名稱（中區）沒有被擋');

  // ── 權限 ──
  eq((await http('_s_u1', 'POST', `/api/company-requests/${rq1.id}/approve`, { name: 'x公司' })).status, 403, 'A25 業務不能核可');
  eq((await http('_s_u1', 'POST', `/api/company-requests/${rq1.id}/reject`, { reason: 'typo' })).status, 403, 'A26 業務不能退回');
  eq((await http('_s_mgr1', 'POST', `/api/company-requests/${rq1.id}/approve`, { name: 'x公司' })).status, 403, 'A27 主管不能核可');
  eq((await http('_s_sec', 'POST', `/api/company-requests/${rq1.id}/approve`, { name: 'x公司' })).status, 403, 'A28 秘書不能核可');
  eq((await http('_s_u3', 'POST', `/api/company-requests/${rq1.id}/cancel`)).status, 403, 'A29 非申請人不能取消');
  eq((await http('_s_mkt', 'POST', `/api/company-requests/${rq1.id}/cancel`)).status, 403, 'A30 行銷也不能代替申請人取消（取消限申請人）');
  eq((await http(null, 'GET', '/api/company-requests')).status, 401, 'A31 沒登入 401');
  eq((await reqList('_s_u1')).items.map((x) => x.id).sort(), [rq1.id, 'seedreq-sap'].sort(), 'A32 業務只看到自己是申請人的申請');
  eq((await reqList('_s_u3')).items.map((x) => x.id), [], 'A33 無關的業務看不到');
  const lm = await reqList('_s_mkt');
  eq([lm.canResolve, lm.items.length >= 2, lm.categories.length, lm.rejectReasons.length], [true, true, 7, 3], 'A34 行銷看全部，並帶類別與退回原因清單');
  const l1 = await reqList('_s_u1');
  eq(l1.items.find((x) => x.id === rq1.id).requesters.map((x) => x.username), ['_s_u1'], 'A35 申請人只看到自己那份申請人資料（看不到別人的說明）');
  eq((await http('_s_gs', 'GET', '/api/company-requests')).status, 200, 'A36 集團業務可查自己的申請（路徑白名單已放行）');
  eq((await http('_s_tp', 'GET', '/api/company-requests')).status, 403, 'A37 集團PM（唯讀白名單）不可');

  // ── 閘門：報價單送簽／SAP ──
  const sub = (id) => http('_s_u1', 'POST', `/api/quotations/${id}/submit`);
  r = await sub('q-by-id'); eq([r.status, r.json && r.json.code], [409, 'COMPANY_UNVERIFIED'], 'A38 報價單（contactId 指到待確認名片）送簽被擋 409', r.status + r.text.slice(0, 200));
  ok(r.json && /建檔|行銷/.test(r.json.error || r.json.message || ''), 'A39 擋下訊息是清楚的中文', r.text.slice(0, 200));
  r = await sub('q-by-name'); eq([r.status, r.json && r.json.code], [409, 'COMPANY_UNVERIFIED'], 'A40 報價單沒有 contactId 但公司名對得到待確認名片 → 擋');
  r = await sub('q-ok'); ok(!(r.json && r.json.code === 'COMPANY_UNVERIFIED'), 'A41 正常客戶的報價單不被建檔閘門擋（之後的其他檢查另計）', r.text.slice(0, 160));
  r = await sub('q-nolink'); ok(!(r.json && r.json.code === 'COMPANY_UNVERIFIED'), 'A42 找不到關聯名片 → 不擋', r.text.slice(0, 160));
  r = await http('_s_adm', 'POST', '/api/admin/integrations/push/batch', { companyIds: [CO_SAP.id] });
  eq([r.status, r.json && r.json.code], [409, 'COMPANY_UNVERIFIED'], 'A43 推送 SAP：有待處理建檔申請的公司被擋', r.text.slice(0, 200));
  r = await http('_s_adm', 'POST', '/api/admin/integrations/push/batch', { companyIds: [CO_SAP_OK.id, CO_SAP.id] });
  eq(r.status, 409, 'A44 批次裡只要有一家被擋就整批擋');
  r = await http('_s_adm', 'POST', '/api/admin/integrations/push/batch', { companyIds: [CO_SAP_OK.id] });
  ok(!(r.json && r.json.code === 'COMPANY_UNVERIFIED'), 'A45 正常公司不被建檔閘門擋（會停在「連線未啟用」）', r.text.slice(0, 160));
  // 日常動作不擋
  r = await http('_s_u1', 'PUT', `/api/contacts/${k1.id}`, { note: '備註改一下', phone: '02-1234-5678' });
  eq([r.status, r.json && r.json.note, r.json && r.json.companyVerify], [200, '備註改一下', 'pending'], 'A46 待確認期間編輯名片不擋（仍待確認）');
  ok(!readData().companies.some((c) => c.taxId === TAX), 'A47 編輯待確認名片不會建立主檔');
  r = await http('_s_u1', 'GET', '/api/contacts'); ok(r.status === 200 && r.json.some((c) => c.id === k1.id && c.companyVerify === 'pending'), 'A48 名片列表照常，並帶 companyVerify');
  r = await http('_s_u1', 'POST', '/api/visits', { contactId: k1.id, date: '2026-10-20', type: '親訪', note: '拜訪測試' });
  ok(r.status === 200 || r.status === 201, 'A49 待確認期間仍可新增拜訪', r.status + r.text.slice(0, 120));

  // ── 取消／重新申請 ──
  r = await http('_s_u2', 'POST', `/api/company-requests/${rq1.id}/cancel`);
  eq([r.status, r.json && r.json.fullyCancelled], [200, false], 'A50 申請人 u2 取消：申請仍在（u1 還在）');
  d = readData();
  const c2 = d.contacts.find((c) => c.id === k2.id);
  eq([c2.companyVerify, c2.companyRequestId, d.companyRequests.find((x) => x.id === rq1.id).requesters.length], ['pending', '', 1], 'A51 取消後名片保留待確認、可再申請；申請人名單剩 u1');
  r = await http('_s_u2', 'PUT', `/api/contacts/${k2.id}`, { companyRequest: RQ('school', '取消後我又要申請了') });
  eq(r.status, 200, 'A52 取消後可再次申請（PUT 只帶 companyRequest）');
  d = readData();
  eq([d.companyRequests.find((x) => x.id === rq1.id).requesters.length, d.contacts.find((c) => c.id === k2.id).companyRequestId], [2, rq1.id], 'A53 重新併入同一筆申請');
  r = await http('_s_u2', 'POST', `/api/company-requests/${rq1.id}/cancel`); eq(r.status, 200, 'A54 u2 再取消');
  r = await http('_s_u1', 'POST', `/api/company-requests/${rq1.id}/cancel`);
  eq([r.status, r.json && r.json.fullyCancelled, readData().companyRequests.find((x) => x.id === rq1.id).status], [200, true, 'cancelled'], 'A55 最後一位取消 → 申請作廢');
  eq((await http('_s_u1', 'POST', `/api/company-requests/${rq1.id}/cancel`)).status, 403, 'A56 作廢後不能再取消');
  lg = await auditLines('COMPANY_REQUEST_CANCEL'); ok(lg.length >= 3, 'A57 取消都寫稽核 COMPANY_REQUEST_CANCEL', lg.length);
  eq(readData().contacts.find((c) => c.id === k1.id).companyVerify, 'pending', 'A58 全部取消後名片仍是待確認（仍擋送簽）');
  r = await sub('q-by-name'); eq(r.status, 409, 'A59 取消後名片仍待確認 → 仍擋報價單送簽');
  r = await http('_s_u1', 'PUT', `/api/contacts/${k1.id}`, { companyRequest: RQ('gov', '重新申請的完整說明') });
  const rq2id = readData().contacts.find((c) => c.id === k1.id).companyRequestId;
  ok(r.status === 200 && rq2id && rq2id !== rq1.id, 'A60 取消後再申請 → 新的申請（不同 id）');
  r = await http('_s_u2', 'PUT', `/api/contacts/${k2.id}`, { companyRequest: RQ('school', '中區也一起併進來') });
  eq(r.status, 200, 'A61 u2 也重新併入');

  // ── 退回 ──
  const mn = (await notifs('_s_u1', 'company_request_done')).length;
  r = await http('_s_mkt', 'POST', `/api/company-requests/${rq2id}/reject`, {});
  eq([r.status, r.json && r.json.code], [400, 'REASON_REQUIRED'], 'A62 退回沒有原因被拒');
  r = await http('_s_mkt', 'POST', `/api/company-requests/${rq2id}/reject`, { reason: 'bogus' });
  eq(r.status, 400, 'A63 退回原因不在清單被拒');
  eq(readData().companyRequests.find((x) => x.id === rq2id).status, 'pending', 'A64 被拒的退回不改變申請狀態');
  r = await http('_s_mkt', 'POST', `/api/company-requests/${rq2id}/reject`, { reason: 'typo', note: XSS });
  eq(r.status, 200, 'A65 行銷退回（原因＋備註）');
  d = readData();
  const kk1 = d.contacts.find((c) => c.id === k1.id);
  eq([kk1.companyVerify, d.companyRequests.find((x) => x.id === rq2id).status, d.companyRequests.find((x) => x.id === rq2id).rejectReason], ['rejected', 'rejected', 'typo'], 'A66 名片 rejected、申請 rejected');
  ok(/統編疑似輸入錯誤/.test(kk1.companyVerifyReason) && kk1.companyVerifyReason.includes(XSS), 'A67 名片帶退回原因文字（含備註，原樣保存、前端 textContent 顯示）');
  const nd = await notifs('_s_u1', 'company_request_done');
  eq(nd.length - mn, 1, 'A68 申請人收到 company_request_done');
  ok(/退回/.test(nd[0].title) && nd[0].refId === k1.id, 'A69 通知標題為退回、refId 是名片 id（點擊可開名片）');
  eq((await notifs('_s_u2', 'company_request_done')).length >= 1, true, 'A70 併單的另一位申請人也收到');
  lg = await auditLines('COMPANY_REQUEST_REJECT'); ok(lg.length === 1 && /統編疑似輸入錯誤/.test(lg[0].detail), 'A71 稽核 COMPANY_REQUEST_REJECT 含原因');
  eq((await http('_s_mkt', 'POST', `/api/company-requests/${rq2id}/reject`, { reason: 'typo' })).status, 409, 'A72 已處理的申請不能再退回');
  eq((await http('_s_mkt', 'POST', `/api/company-requests/${rq2id}/approve`, { name: '某公司' })).status, 409, 'A73 已處理的申請不能核可');
  r = await sub('q-by-name'); eq([r.status, r.json && r.json.code], [409, 'COMPANY_UNVERIFIED'], 'A74 被退回期間不可送簽');
  ok(/退回/.test(r.json.error || r.json.message || ''), 'A75 擋下訊息說明是被退回');
  // 修正統編後自動重新送出
  r = await http('_s_u2', 'PUT', `/api/contacts/${k2.id}`, { taxId: '08628408', companyRequest: RQ('school', '統編修正後重新送出') });
  d = readData();
  const kk2 = d.contacts.find((c) => c.id === k2.id);
  const newReq = d.companyRequests.find((x) => x.taxId === '08628408');
  eq([r.status, kk2.companyVerify, !!newReq && newReq.status, !!newReq && newReq.id === kk2.companyRequestId], [200, 'pending', 'pending', true], 'A76 被退回後改統編＋重新帶 companyRequest：自動送出新申請、rejected 清掉');
  eq(d.companyRequests.find((x) => x.id === rq2id).status, 'rejected', 'A77 舊的被退回紀錄保留（稽核軌跡）');
  r = await http('_s_u1', 'PUT', `/api/contacts/${k1.id}`, { taxId: CO_EXIST.taxId, company: CO_EXIST.name });
  d = readData(); const kk1b = d.contacts.find((c) => c.id === k1.id);
  eq([r.status, kk1b.companyVerify, kk1b.companyRequestId], [200, undefined, undefined], 'A78 被退回後改成命中內部主檔的統編（前端已驗證通過）：待確認標記清掉');
  r = await sub('q-by-name'); ok(!(r.json && r.json.code === 'COMPANY_UNVERIFIED'), 'A79 清掉後報價單不再被建檔閘門擋', r.text.slice(0, 120));
  // 待確認改統編：舊申請（唯一申請人）自動作廢
  r = await newCard('_s_u3', '孤單位', { taxId: '08700001', companyRequest: RQ('shop', '行號測試用說明') });
  const k3 = r.json; const k3req = k3.companyRequestId;
  r = await http('_s_u3', 'PUT', `/api/contacts/${k3.id}`, { taxId: '08700002', companyRequest: RQ('shop', '改了統編後的說明') });
  d = readData();
  eq([d.companyRequests.find((x) => x.id === k3req).status, d.companyRequests.find((x) => x.id === d.contacts.find((c) => c.id === k3.id).companyRequestId).taxId], ['cancelled', '08700002'], 'A80 待確認改統編：舊申請（唯一申請人）自動作廢、改開新統編的申請');
  lg = await auditLines('COMPANY_REQUEST_CANCEL'); ok(lg.some((l) => /改統編，申請自動作廢/.test(l.detail)), 'A81 自動作廢也寫稽核');

  // ── 核可 ──
  // 以 u1 重新申請（k1 現在掛的是 CO_EXIST；新建一張）
  r = await newCard('_s_u1', '北區健保署', { companyRequest: RQ('gov', '核可流程測試用說明') });
  const k4 = r.json;
  const rq3 = readData().companyRequests.find((x) => x.taxId === TAX && x.status === 'pending');
  // 讓 u2 的舊名片（kk2 已改統編）不影響；另外放一張「已取消申請但仍待確認」的同統編名片
  r = await newCard('_s_u3', '南區健保署', { companyRequest: RQ('gov', '南區也申請了一次') });
  const k5 = r.json;
  await http('_s_u3', 'POST', `/api/company-requests/${rq3.id}/cancel`);   // u3 取消 → k5 保留待確認、無申請
  eq(readData().contacts.find((c) => c.id === k5.id).companyRequestId, '', 'A82 u3 取消後 k5 無綁定申請但仍待確認');
  r = await http('_s_mkt', 'POST', `/api/company-requests/${rq3.id}/approve`, {});
  eq(r.status, 200, 'A83 核可時沒帶名稱 → 預設採用申請的公司名稱（官方名稱流程見 API-2）');
  d = readData();
  let master = d.companies.find((c) => c.taxId === TAX);
  ok(master && master.name === '北區健保署', 'A84 核可建立企業主檔（名稱預設為申請公司名）', JSON.stringify(master));
  eq([master.matchKey, master.source], ['tax:' + TAX, 'company-request'], 'A85 主檔以統編為鍵、來源標記 company-request');
  eq([d.contacts.find((c) => c.id === k4.id).companyVerify, d.contacts.find((c) => c.id === k4.id).companyId, d.contacts.find((c) => c.id === k5.id).companyVerify], [undefined, master.id, undefined], 'A86 核可：申請人名片清除待確認並掛主檔；已取消者（同統編）也一併清除');
  ok(!d.contacts.some((c) => c.taxId === TAX && c.companyVerify), 'A87 同統編沒有任何名片還是待確認');
  eq((await http('_s_mkt', 'POST', `/api/company-requests/${rq3.id}/approve`, { name: '再核可一次' })).status, 409, 'A88 重複核可 409');
  lg = await auditLines('COMPANY_REQUEST_APPROVE'); ok(lg.length === 1 && /新建企業主檔/.test(lg[0].detail), 'A89 稽核 COMPANY_REQUEST_APPROVE');
  eq((await notifs('_s_u1', 'company_request_done')).filter((n) => /核可/.test(n.title)).length >= 1, true, 'A90 申請人收到核可通知');
  r = await sub('q-by-name'); ok(!(r.json && r.json.code === 'COMPANY_UNVERIFIED'), 'A91 核可後不再擋', r.text.slice(0, 100));
}

// 第二台伺服器：核可流程用「官方名稱＋產業」並驗證子單位名稱規則
async function groupApi2() {
  console.log('[API-2] 核可（官方名稱＋產業）與子單位名稱規則');
  await freshServer(MUT_CUR);
  await loginAll();
  let r = await newCard('_s_u1', '北區健保署', { companyRequest: RQ('gov', '衛福部健保署北區業務組') });
  const kA = r.json;
  r = await newCard('_s_u2', '中區健保署', { companyRequest: RQ('gov', '中區業務組') });
  const kB = r.json;
  r = await http('_s_u1', 'PUT', `/api/contacts/${kA.id}`, { industry: '偷填產業', note: '更新' });
  eq(readData().contacts.find((c) => c.id === kA.id).industry, '', 'B1 業務填的產業被忽略（產業鎖）');
  r = await http('_s_mkt', 'POST', `/api/company-requests/${kA.companyRequestId}/approve`, { name: '12345678', industry: '政府機關' });
  eq(r.status, 400, 'B2 正式名稱純數字被拒');
  r = await http('_s_mkt', 'POST', `/api/company-requests/${kA.companyRequestId}/approve`, { name: '長'.repeat(101) });
  eq(r.status, 400, 'B3 正式名稱太長被拒');
  r = await http('_s_u1', 'POST', `/api/company-requests/${kA.companyRequestId}/approve`, { name: OFFICIAL });
  eq(r.status, 403, 'B4 業務不能核可');
  r = await http('_s_mkt', 'POST', `/api/company-requests/${kA.companyRequestId}/approve`, { name: OFFICIAL, address: '台北市中正區某路1號', industry: '政府機關' });
  eq([r.status, r.json && r.json.created, r.json && r.json.cleared], [200, true, 2], 'B5 行銷核可（正式名稱＋地址＋產業）：新建主檔、清除 2 張名片');
  const d = readData();
  const m = d.companies.find((c) => c.taxId === TAX);
  eq([m.name, m.address, m.industry, m.industrySource], [OFFICIAL, '台北市中正區某路1號', '政府機關', '手動'], 'B6 主檔內容（產業只由行銷填）');
  eq([d.contacts.find((c) => c.id === kA.id).industry, d.contacts.find((c) => c.id === kB.id).industry], ['政府機關', '政府機關'], 'B7 產業同步到同公司名片（沿用既有 cascade）');
  eq([d.contacts.find((c) => c.id === kA.id).company, d.contacts.find((c) => c.id === kB.id).company], ['北區健保署', '中區健保署'], 'B8 名片上業務填的子單位名稱不被改動');
  r = await http('_s_u3', 'GET', `/api/company-master?taxId=${TAX}`);
  eq([r.json.found, r.json.company.name], [true, OFFICIAL], 'B9 核可後任何人輸入同統編直接命中內部主檔（前端會自動帶入正式名稱）');
  // 子單位名稱 vs TaxIdGuard.checkSave（主檔已存在）
  r = await newCard('_s_u3', '南區健保署');
  eq([r.status, r.json && r.json.code], [400, 'TAXID_COMPANY_MISMATCH'], 'B10 【現行規則】核可後新名片若手動輸入子單位名稱「南區健保署」→ 400 不符（前端會自動帶入官方名稱避免此情況）');
  r = await newCard('_s_u3', OFFICIAL);
  eq(r.status, 201, 'B11 核可後輸入官方名稱 → 通過');
  r = await newCard('_s_u3', '南區健保署', { companyRequest: RQ() });
  eq([r.status, r.json && r.json.code], [400, 'TAXID_COMPANY_MISMATCH'], 'B12 主檔已存在時帶 companyRequest 不會繞過名稱一致性規則');
  r = await http('_s_u1', 'PUT', `/api/contacts/${kA.id}`, { note: '核可後編輯既有子單位名片' });
  eq(r.status, 200, 'B13 【現行規則】核可前已存在的子單位名片，只改其他欄位（公司／統編沒動）→ 不受影響');
  r = await http('_s_u1', 'PUT', `/api/contacts/${kA.id}`, { company: '北區健保署第二組' });
  eq([r.status, r.json && r.json.code], [400, 'TAXID_COMPANY_MISMATCH'], 'B14 【現行規則】核可後若改公司欄成另一個子單位名稱 → 400 不符');
  r = await http('_s_u1', 'PUT', `/api/contacts/${kA.id}`, { company: OFFICIAL });
  eq(r.status, 200, 'B15 改成官方名稱 → 通過');
  // 申請期間：主檔尚未存在 → 不誤擋
  r = await newCard('_s_u3', '甲子單位', { taxId: '08555555', companyRequest: RQ('school', '某某大學資訊處') });
  eq(r.status, 201, 'B16 申請待處理期間（主檔尚未建立），輸入任意子單位名稱不被擋');
  r = await newCard('_s_u1', '乙子單位', { taxId: '08555555', companyRequest: RQ('school', '某某大學圖書館') });
  eq([r.status, r.json.companyRequest && r.json.companyRequest.merged], [201, true], 'B17 同統編另一個子單位名稱也不被擋，且併單');
  // 行銷自己申請自己核可；admin 也可核可
  r = await http('_s_mkt', 'POST', '/api/contacts', { name: '行銷自建', company: '行銷申請單位', taxId: '08666666', bu: 'ERP', email: 'm@example.test', companyRequest: RQ('foundation', '財團法人測試說明') });
  eq(r.status, 201, 'B18 行銷也可以走申請（自己建名片）');
  const mreq = r.json.companyRequestId;
  r = await http('_s_mkt', 'POST', `/api/company-requests/${mreq}/approve`, { name: '假財團法人測試基金會' });
  eq(r.status, 200, 'B19 行銷本人可核可自己的申請（不設雙人覆核）');
  r = await newCard('_s_u3', '管理員核可用', { taxId: '08777777', companyRequest: RQ('association', '協會團體測試說明') });
  r = await http('_s_adm', 'POST', `/api/company-requests/${r.json.companyRequestId}/approve`, { name: '假協會' });
  eq(r.status, 200, 'B20 管理員可核可');
  // 說明含 XSS 酬載：原樣存、JSON 回傳（前端 textContent）
  r = await newCard('_s_u3', XSS, { taxId: '08888888', companyRequest: RQ('other', XSS + '說明') });
  eq(r.status, 201, 'B21 酬載字串可存（以純文字處理）');
  const lst = await reqList('_s_mkt');
  ok(lst.items.some((x) => x.note.includes('<img src=x')), 'B22 清單 API 原樣回傳，不轉義（顯示端負責 textContent）');
  // 通知網址
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  ok(/case 'company_request_new':\s*\n\s*return '\/index\.html#creq'/.test(src) && /case 'company_request_done':\s*\n\s*return refId \? `\/index\.html#contact:\$\{refId\}`/.test(src), 'B23 notificationUrlFor 補了兩個 type');
  // 逾期提醒（lazy）
  await stopServer(); cleanupFiles();
  seedFiles({ data: (dd) => {
    const mk = (id, tax, ageD, reminded) => ({ id, taxId: tax, companyName: '公司' + id, category: 'gov', note: '提醒測試說明', requestedBy: '_s_u1', requestedAt: agoIso(ageD * DAY + 3600000), contactId: '', contactIds: [], requesters: [{ username: '_s_u1', contactId: '', companyName: '公司' + id, category: 'gov', note: '提醒測試說明', at: nowIso }], status: 'pending', resolvedBy: '', resolvedAt: '', result: null, rejectReason: '', rejectNote: '', cancelledBy: '', lastRemindedDate: reminded || '' });
    dd.companyRequests = [mk('r2', '10000002', 2), mk('r4', '10000004', 4), mk('r8', '10000008', 8)];
  } });
  await startServer(MUT_CUR); jar = {}; await loginAll();
  const l = await reqList('_s_mkt');
  eq(l.items.map((x) => [x.id, x.overdue]).sort(), [['r2', false], ['r4', false], ['r8', true]], 'B24 逾 7 天標紅（8 天 true、4 天 false）');
  const rem = (await notifs('_s_mkt', 'company_request_new'));
  eq(rem.length, 2, 'B25 超過 3 天的 2 筆各提醒一次（2 天的不提醒）');
  ok(rem.every((n) => /已等待/.test(n.title)), 'B26 提醒標題為「已等待 N 天」');
  await reqList('_s_mkt'); await reqList('_s_mkt2');
  eq([(await notifs('_s_mkt', 'company_request_new')).length, (await notifs('_s_mkt2', 'company_request_new')).length], [2, 2], 'B27 同一筆同一天不重複提醒（再載入兩次不增加）；每位行銷各收到');
  eq((await notifs('_s_adm', 'company_request_new')).length, 0, 'B28 不通知主管／管理員');
  await reqList('_s_u1');
  eq((await notifs('_s_mkt', 'company_request_new')).length, 2, 'B29 業務載入清單不會觸發提醒');
  eq(readData().companyRequests.filter((x) => x.lastRemindedDate).length, 2, 'B30 提醒日期寫回資料（lastRemindedDate）');
  await stopServer(); cleanupFiles();
}

// ═════════════ 開關：COMPANY_REQUESTS=off ═════════════
const BASE_REF_DEFAULT = '425fc4f'; // 尚未有此功能的 main（PR #43 合併後）
function exportBase() {
  if (process.env.CREQ_BASE_DIR) return { dir: process.env.CREQ_BASE_DIR, cleanup() {} };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'creq-base-'));
  const baseRef = process.env.CREQ_BASE_REF || BASE_REF_DEFAULT;
  const a = spawnSync('git', ['archive', baseRef], { cwd: ROOT, maxBuffer: 256 * 1024 * 1024 });
  if (a.status !== 0) throw new Error('git archive ' + baseRef + ' 失敗：' + String(a.stderr));
  const t = spawnSync('tar', ['-x', '-C', dir], { input: a.stdout, maxBuffer: 256 * 1024 * 1024 });
  if (t.status !== 0) throw new Error('tar 失敗：' + String(t.stderr));
  spawnSync('cmd', ['/c', 'mklink', '/J', path.join(dir, 'node_modules'), path.join(ROOT, 'node_modules')], { stdio: 'ignore' });
  return { dir, cleanup() { spawnSync('cmd', ['/c', 'rmdir', path.join(dir, 'node_modules')], { stdio: 'ignore' }); try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ } } };
}
async function groupOff() {
  console.log('[開關] COMPANY_REQUESTS=off ＝ 與沒有此功能時位元級相同');
  const norm = (x) => String(x).replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/g, '<uuid>').replace(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/g, '<ts>');
  const SEED = { data: (d) => { d.contacts.push(mkCard('kold', '_s_u1', '已有待確認標記的舊名片', '08999999', { companyVerify: 'pending', companyRequestId: 'x' })); d.quotations.find((q) => q.id === 'q-by-id').contactId = 'kold'; } };
  const run = async (base, root, port, env, mut) => {
    await stopServer(); cleanupFiles(root); seedFiles(Object.assign({ dir: root }, SEED));
    await startServer(mut || null, env, root, port);
    const jr = {};
    for (const w of ['_s_adm', '_s_mkt', '_s_u1', '_s_u2']) await loginAt(base, jr, w);
    const out = {};
    const body = (extra) => Object.assign({ name: '離線測試', company: '北區健保署', taxId: TAX, bu: 'ERP', email: 'o@example.test' }, extra || {});
    let rr = await httpAt(base, jr, '_s_u1', 'POST', '/api/contacts', body({ companyRequest: { category: 'gov', note: '衛福部健保署北區業務組' } }));
    out.postWithReq = rr.status + ' ' + norm(rr.text);
    const id1 = rr.json && rr.json.id;
    rr = await httpAt(base, jr, '_s_u1', 'POST', '/api/contacts', body({ company: '中區健保署' }));
    out.postPlain = rr.status + ' ' + norm(rr.text);
    rr = await httpAt(base, jr, '_s_u1', 'POST', '/api/contacts', body({ company: '壞資料', companyRequest: { category: 'zz', note: '1' } }));
    out.postBadReq = rr.status + ' ' + norm(rr.text);
    rr = await httpAt(base, jr, '_s_u1', 'PUT', `/api/contacts/${id1}`, { note: '改備註', taxId: '08628408', company: '另一單位', companyRequest: { category: 'gov', note: '衛福部健保署北區業務組' } });
    out.putWithReq = rr.status + ' ' + norm(rr.text);
    rr = await httpAt(base, jr, '_s_u1', 'PUT', '/api/contacts/kold', { taxId: '08999998', company: '改過的舊名片' });
    out.putOld = rr.status + ' ' + norm(rr.text);
    for (const who of ['_s_adm', '_s_mkt', '_s_u1', '_s_u2']) for (const g of ['/api/contacts', '/api/companies', '/api/admin/companies', `/api/company-master?taxId=${TAX}`, '/api/notifications']) { rr = await httpAt(base, jr, who, 'GET', g); out[who + ' ' + g] = rr.status + ' ' + norm(rr.text); }
    out.quotePending = await (async () => { const x = await httpAt(base, jr, '_s_u1', 'POST', '/api/quotations/q-by-id/submit'); return x.status + ' ' + norm(x.text); })();
    out.sapPending = await (async () => { const x = await httpAt(base, jr, '_s_adm', 'POST', '/api/admin/integrations/push/batch', { companyIds: [CO_SAP.id] }); return x.status + ' ' + norm(x.text); })();
    const raw = JSON.parse(fs.readFileSync(path.join(root, 'data.json'), 'utf8'));
    out.contactKeys = JSON.stringify(raw.contacts.map((c) => Object.keys(c).sort().filter((k) => !['updatedAt', 'createdAt', 'id'].includes(k))));
    out.companyKeys = JSON.stringify(raw.companies.map((c) => c.taxId + ':' + c.name));
    out.reqCount = String((raw.companyRequests || []).length);
    const extra = {};
    extra.list = await httpAt(base, jr, '_s_mkt', 'GET', '/api/company-requests');
    extra.approve = await httpAt(base, jr, '_s_mkt', 'POST', '/api/company-requests/seedreq-sap/approve', { name: 'x' });
    extra.reject = await httpAt(base, jr, '_s_mkt', 'POST', '/api/company-requests/seedreq-sap/reject', { reason: 'typo' });
    extra.cancel = await httpAt(base, jr, '_s_u1', 'POST', '/api/company-requests/seedreq-sap/cancel');
    await stopServer(); cleanupFiles(root);
    return { out, extra };
  };
  const baseX = exportBase();
  let A, B;
  try {
    A = await run(BASE, baseX.dir, PORT, {});
    B = await run(BASE, ROOT, PORT, { COMPANY_REQUESTS: 'off' }, MUT_CUR);
  } finally { baseX.cleanup(); }
  const diffKeys = Object.keys(A.out).filter((k) => A.out[k] !== B.out[k]);
  ok(diffKeys.length === 0, `F1 off：名片存檔（含 companyRequest）、列表、企業主檔、通知、報價單送簽、SAP 推送的輸出與「修改前程式碼」逐位元相同（比對 ${Object.keys(A.out).length} 項）`, diffKeys.join(', '));
  if (diffKeys.length) console.log('    差異首項：', diffKeys[0], '\n      前：', String(A.out[diffKeys[0]]).slice(0, 300), '\n      後：', String(B.out[diffKeys[0]]).slice(0, 300));
  eq([B.extra.list.status, B.extra.approve.status, B.extra.reject.status, B.extra.cancel.status], [404, 404, 404, 404], 'F2 off：新端點全部 404');
  eq(B.extra.list.json && B.extra.list.json.code, 'COMPANY_REQUESTS_OFF', 'F3 off：404 帶 COMPANY_REQUESTS_OFF');
  ok(/^201 /.test(B.out.postWithReq) && !/companyVerify/.test(B.out.postWithReq), 'F4 off：帶 companyRequest 的存檔照常建立、不標待確認');
  ok(!/^409 /.test(B.out.quotePending) && !/^409 /.test(B.out.sapPending), 'F5 off：已有待確認標記的舊資料不擋送簽與 SAP 推送');
  eq(B.out.reqCount, '1', 'F6 off：不產生任何申請（只剩種子資料那 1 筆）');
  // 預設（沒設環境變數）＝開
  await freshServer();
  await login('_s_u1');
  eq((await reqList('_s_u1')).enabled, true, 'F7 沒設環境變數＝開');
  await stopServer(); cleanupFiles();
}

// ═════════════ 變異測試 ═════════════
let MUT_CUR = null;
const MUTS = [
  { name: 'M1 拿掉核可的權限檢查', groups: ['api'], spec: [{ file: 'server.js', edits: [{ find: "app.post('/api/company-requests/:id/approve', requireAuth, (req, res) => {\n  if (!CREQ.enabled()) return creqOff(res);\n  if (!isAdminOrMarketing(req)) return res.status(403).json({ error: '無權限（限行銷／管理員）' });", rep: "app.post('/api/company-requests/:id/approve', requireAuth, (req, res) => {\n  if (!CREQ.enabled()) return creqOff(res);" }] }] },
  { name: 'M1b 拿掉退回的權限檢查', groups: ['api'], spec: [{ file: 'server.js', edits: [{ find: "app.post('/api/company-requests/:id/reject', requireAuth, (req, res) => {\n  if (!CREQ.enabled()) return creqOff(res);\n  if (!isAdminOrMarketing(req)) return res.status(403).json({ error: '無權限（限行銷／管理員）' });", rep: "app.post('/api/company-requests/:id/reject', requireAuth, (req, res) => {\n  if (!CREQ.enabled()) return creqOff(res);" }] }] },
  { name: 'M2 拿掉併單', groups: ['unit', 'api'], spec: [{ file: 'lib/companyRequests.js', edits: [{ find: 'let r = findPendingByTax(data, taxId);', rep: 'let r = null;' }] }] },
  { name: 'M3 拿掉關閉開關', groups: ['unit', 'off'], spec: [{ file: 'lib/companyRequests.js', edits: [{ find: "return v !== 'off';", rep: 'return true;' }] }] },
  { name: 'M4 拿掉報價單送簽擋', groups: ['api'], spec: [{ file: 'lib/quoteRoutes.js', edits: [{ find: 'const blk = CREQ.blockForQuote(ctx.data, q);', rep: 'const blk = null;' }] }] },
  { name: 'M4b 拿掉 SAP 推送擋', groups: ['api'], spec: [{ file: 'server.js', edits: [{ find: 'const _blk = CREQ.blockForSapCompanies(', rep: 'const _blk = null && CREQ.blockForSapCompanies(' }] }] },
  { name: 'M5 拿掉退回原因必填', groups: ['unit', 'api'], spec: [{ file: 'lib/companyRequests.js', edits: [{ find: "if (!REJECT_REASONS.some((x) => x.key === reason)) return { ok: false, code: 'REASON_REQUIRED', error: '退回必須選擇原因' };", rep: '' }] }] },
  { name: 'M6 申請期間照舊建立企業主檔', groups: ['api'], spec: [{ file: 'server.js', edits: [{ find: 'const _master = _creqIn ? null : ensureCompanyMaster(data, contact);', rep: 'const _master = ensureCompanyMaster(data, contact);' }] }] },
  { name: 'M7 核可不清除名片的待確認標記', groups: ['unit', 'api'], spec: [{ file: 'lib/companyRequests.js', edits: [{ find: '    clearVerify(c);\n    if (!c.deleted) c.companyId = master.id;', rep: '    if (!c.deleted) c.companyId = master.id;' }] }] },
  { name: 'M8 拿掉說明字數下限', groups: ['unit', 'api'], spec: [{ file: 'lib/companyRequests.js', edits: [{ find: "if ([...note].length < NOTE_MIN) return", rep: "if (false) return" }] }] },
  { name: 'M9 取消不驗證申請人', groups: ['unit', 'api'], spec: [{ file: 'lib/companyRequests.js', edits: [{ find: "if (!(r.requesters || []).some((x) => x.username === username)) return { ok: false, status: 403, error: '只有申請人可以取消自己的申請' };", rep: '' }] }] },
];
async function groupMut() {
  console.log('[變異] 每個變異都必須讓對應群組變紅');
  const snap = { p: pass, f: fail, n: failures.length };
  let undetected = 0;
  for (const m of MUTS) {
    const before = { p: pass, f: fail, n: failures.length };
    quiet = true; MUT_CUR = m.spec;
    try {
      for (const g of m.groups) {
        if (g === 'unit') {
          // 單元測試直接 require 本地檔：用與預載相同的改寫方式載入一份變異後的副本
          const src = fs.readFileSync(path.join(ROOT, 'lib', 'companyRequests.js'), 'utf8');
          let mutated = src;
          m.spec.filter((x) => x.file === 'lib/companyRequests.js').forEach((x) => x.edits.forEach((e) => { if (!mutated.includes(e.find)) throw new Error('MUTATION ANCHOR MISSING ' + e.find.slice(0, 50)); mutated = mutated.replace(e.find, () => e.rep); }));
          if (mutated === src) continue;
          const tmp = path.join(os.tmpdir(), 'creq_mut_' + process.pid + '.js');
          fs.writeFileSync(tmp, mutated);
          const Real = CR_REF.mod; CR_REF.mod = require(tmp);
          try { groupUnit(); } finally { CR_REF.mod = Real; delete require.cache[tmp]; try { fs.unlinkSync(tmp); } catch { /* 忽略 */ } }
        }
        if (g === 'api') { await groupApi(); await groupApi2(); }
        if (g === 'off') await groupOff();
      }
    } catch (e) { fail++; failures.push('(變異造成例外) ' + e.message); }
    quiet = false; MUT_CUR = null;
    const newFails = fail - before.f;
    const killedBy = failures.slice(before.n, before.n + 3).map((x) => x.slice(0, 48));
    console.log('  ' + (newFails > 0 ? '✓' : '✗') + ' ' + m.name + '：' + (newFails > 0 ? '變紅（' + newFails + ' 項失敗：' + killedBy.join(' / ') + '…）' : '沒有變紅！對應測試無效'));
    pass = before.p; fail = before.f; failures.length = before.n;   // 變異造成的失敗不計入本次結果
    if (newFails === 0) { undetected++; fail++; failures.push('變異未被偵測：' + m.name); }
  }
  console.log('  變異總計 ' + MUTS.length + ' 個，未被偵測 ' + undetected + ' 個');
  await stopServer(); cleanupFiles();
}
const CR_REF = { mod: CR };

// ═════════════ 主程式 ═════════════
(async () => {
  let ui = null;
  try {
    if (want('unit')) groupUnit();
    if (want('api')) { await groupApi(); await groupApi2(); }
    if (want('off')) await groupOff();
    if (want('ui')) {
      ui = require('./_company-requests-ui.js');
      await ui.run({ ok, eq, sleep, http, httpAt, login, loginAll, freshServer, stopServer, cleanupFiles, readData, newCard, notifs, reqList, RQ, TAX, OFFICIAL, XSS, BASE, PW, ROOT, SHOTS, CDP, MUT: MUT_CUR, seedFiles, startServer, agoIso });
    }
    if (process.env.CREQ_MUT === '1') await groupMut();
  } catch (e) {
    console.log('例外：', e && e.stack || e); fail++; failures.push('例外：' + (e && e.message));
  } finally {
    await stopServer(); cleanupFiles();
    try { fs.unlinkSync(PRELOAD); } catch { /* 忽略 */ }
  }
  console.log(`\n統編查無申請建檔檢查結果：${pass} 通過，${fail} 失敗`);
  if (fail) console.log(failures.map((f) => ' - ' + f).join('\n'));
  process.exit(fail ? 1 : 0);
})();
