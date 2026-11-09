#!/usr/bin/env node
'use strict';
/**
 * 秘書唯讀客戶視圖（SECRETARY_CUSTOMERS=1）檢查：
 *   unit   lib 純函式：白名單、已成交判定、前後端「已成交」對拍（vm 載入 app.js 的 getWonContactIds，隨機資料）、靜態檢查
 *   api    真實 server.js（開關開）：範圍（同 BU／已成交／刪除／待補／其他 BU／多 BU／無 BU）、欄位白名單、PII 哨兵掃描、
 *          郵寄地址三情境、權限（只有 secretary）、功能鍵自動補上、寫入端點仍擋、稽核（只記 id＋公司名）
 *   roles  23 個帳號（13 種 role）× 唯讀端點電池：基準版(SCU_BASE_REF，預設 4ef32a9＝郵寄地址 PR 合併後、本功能加入前的 main)／開關關／開關開 三方逐位元比對，
 *          開啟時只允許登錄的差異；新端點只有 secretary 200
 *   off    開關關：新端點 404、功能鍵不登錄、featureGuard 對照表不含新路由（含 enforce 模式）
 *   ui     headless Edge + CDP（見 _secretary-customers-ui.js）
 *
 *   node scripts/check-secretary-customers.js              unit + api + roles + off
 *   SCU_ONLY=unit,api,roles,off,ui node ...                 只跑指定群組
 *   SCU_MUT=1 node ...                                      另外跑變異測試（每個變異都必須讓對應群組變紅）；SCU_MUT_PICK=<regex> 只跑符合名稱者
 *   SCU_SHOTS=<資料夾>                                      ui 截圖輸出
 *   SCU_PORT（預設 3472，限 3470–3479）／SCU_CDP（預設 9772，限 9770–9779）／SCU_BASE_REF
 * 不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體；全部假資料；結束清掉伺服器與 Edge。
 * 守門：工作樹已有 auth.json／data.json／audit.log.json 時拒絕執行。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const Module = require('module');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.SCU_PORT || '3472', 10);
const CDP = parseInt(process.env.SCU_CDP || '9772', 10);
if (PORT < 3470 || PORT > 3479 || CDP < 9770 || CDP > 9779) { console.log('拒絕：埠需在 3470–3479 / CDP 9770–9779'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map((f) => path.join(ROOT, f));
{
  const g = FILES.filter((f) => /(auth|data|audit.log)\.jsonl?$/.test(f)).filter((f) => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map((f) => path.basename(f)).join('/'), '（可能是真實資料）'); process.exit(2); }
}
const ONLY = (process.env.SCU_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length ? k !== 'ui' : ONLY.includes(k);
const SHOTS = process.env.SCU_SHOTS || '';
const BASE_REF = process.env.SCU_BASE_REF || '4ef32a9';
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const XLSX = require(path.join(ROOT, 'node_modules', 'xlsx'));
const LIBPATH = path.join(ROOT, 'lib', 'secretaryCustomers.js');

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 400)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = (n) => new Promise((r) => setTimeout(r, n));

// ───────────────── 沙盒資料（全部假的）─────────────────
const PW = 'Sc' + crypto.randomBytes(9).toString('hex') + '!Z';
const HASH = bcrypt.hashSync(PW, 6);
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
// 嚴格 PII 哨兵：Email／備註／偏好／名片圖等仍不得出現在任何秘書回應。電話／分機／手機（TEL／EXT／MOB）已開放，只准出現在被選中的主要聯絡人（PHONE_RE 另行精確斷言）
const PII_RE = /(NOTE_PII_|MEMO_PII_|HOBBY_PII_|CARD_PII_|EVIL_PII_|@pii\.example\.test)/;
const PHONE_RE = /(?:TEL|EXT|MOB)_PII_[A-Za-z0-9]+/g;
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: HASH, createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canDownloadContacts: false, canSetTargets: false, accessMode: 'edit', bu: [] }, extra || {});
// 23 個帳號、13 種 role（KNOWN_ROLES 全部）
const ACCTS = [
  ['_s_adm', 'admin', {}], ['_s_adm2', 'admin', {}], ['_s_exe', 'executive', {}], ['_s_exe2', 'executive', {}],
  ['_s_m1', 'manager1', { bu: ['ERP'] }], ['_s_m1b', 'manager1', { bu: ['ITS'] }], ['_s_m2', 'manager2', { bu: ['ERP'] }],
  ['_s_sec', 'secretary', { bu: ['ERP'], displayName: '秘書甲' }], ['_s_sec2', 'secretary', { bu: ['ERP', 'ITS'] }], ['_s_secx', 'secretary', { bu: ['ITS'] }],
  ['_s_secnb', 'secretary', {}], ['_s_secp', 'secretary', { bu: ['ERP'], pureSupervisor: true }],
  ['_s_tp', 'tecopm', { viewOwnerScope: '_s_u1', viewGroupId: 'g1' }], ['_s_gs', 'groupsales', { viewOwnerScope: '_s_u1', viewGroupId: 'g1', bu: ['ERP'] }],
  ['_s_mkt', 'marketing', {}],
  ['_s_u1', 'user', { bu: ['ERP'], displayName: '測試業務甲', nickname: '甲哥' }], ['_s_u2', 'user', { bu: ['ERP'], displayName: '測試業務乙' }],
  ['_s_u3', 'user', { bu: ['ITS'], displayName: '測試業務丙' }], ['_s_u4', 'user', { bu: ['ERP'] }],
  ['_s_acc', 'accounting_manager', {}], ['_s_fin', 'finance_manager', {}],
  ['_s_cs', 'consult_manager_south', {}], ['_s_cn', 'consult_manager_north', {}],
];
const ALL_ACCTS = ACCTS.map((a) => a[0]);
const mkCard = (id, owner, company, extra) => Object.assign({
  id, owner, name: '名片' + id, company, title: '職稱' + id, bu: 'ERP', createdAt: nowIso,
  phone: 'TEL_PII_' + id, ext: 'EXT_PII_' + id, mobile: 'MOB_PII_' + id, email: 'pii_' + id + '@pii.example.test', note: 'NOTE_PII_' + id,
  personalMemo: 'MEMO_PII_' + id, personalHobbies: 'HOBBY_PII_' + id, cardImage: 'CARD_PII_' + id, address: '', taxId: '', customerType: '',
}, extra || {});
const mkCo = (id, name, tax) => ({ id, taxId: tax, matchKey: 'tax:' + tax, name, shortName: '', capital: null, industry: '', address: '', phone: '', website: '', region: '', customerCode: '', buImport: [], groupId: null, createdAt: nowIso });
const mkOpp = (id, owner, contactId, company, stage) => ({ id, owner, contactId, company, stage, product: '測試產品', amount: 100, expectedDate: '2026-12-31', category: 'ERP', createdAt: nowIso });
function seedData() {
  const contacts = [
    mkCard('w1', '_s_u1', '甲公司', { customerType: 'customer', address: 'ADDR-A路1號', mailingAddress: 'MAIL-A路1號', mailingSameAsCompany: false }),
    mkCard('w2', '_s_u2', '乙公司', { address: 'ADDR-B路2號', mailingAddress: 'ADDR-B路2號', mailingSameAsCompany: true }),
    mkCard('w3', '_s_u1', '丙公司', { address: 'ADDR-C路3號' }),
    mkCard('w3b', '_s_u2', '丙公司', { customerType: 'customer', address: 'ADDR-C路3號' }),
    mkCard('w4', '_s_u1', '丁公司', { customerType: 'customer', phone: '', ext: '', mobile: '' }),
    mkCard('n1', '_s_u1', '戊公司', { address: 'ADDR-N1' }),
    mkCard('n2', '_s_u1', '己公司', { address: 'ADDR-N2' }),
    mkCard('n3', '_s_u3', '庚公司', { bu: 'ITS', customerType: 'customer', address: 'ADDR-N3路' }),
    mkCard('n4', '_s_u1', '辛公司', { customerType: 'customer', deleted: true, address: 'ADDR-N4' }),
    mkCard('n5', '_pool', '壬公司', { customerType: 'customer', isPlaceholder: true, address: 'ADDR-N5' }),
    mkCard('n6', '_s_mkt', '癸公司', { customerType: 'customer', address: 'ADDR-N6' }),
    mkCard('n8', '_s_u1', '子公司', { bu: 'ITS', customerType: 'customer', address: 'ADDR-N8' }),
    mkCard('x1', '_s_u1', '跨公司', { address: 'ADDR-X1' }),
    mkCard('x2', '_s_u3', '跨公司', { bu: 'ITS', customerType: 'customer', address: 'ADDR-X2' }),
    mkCard('ka1', '_s_u3', 'KA公司', { bu: 'ITS', address: 'ADDR-KA' }),
  ];
  return {
    contacts, companies: [mkCo('c1', '甲公司', '90000001'), mkCo('c2', '乙公司', '90000002')],
    opportunities: [mkOpp('o1', '_s_u2', 'w2', '乙公司', 'Won'), mkOpp('o2', '_s_u1', 'n1', '戊公司', 'A'), mkOpp('o3', '_s_u3', 'n3', '庚公司', 'Won')],
    visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [{ id: 'ka-1', owner: '_s_u3', company: 'KA公司', createdAt: nowIso }], quotations: [],
    groups: [{ id: 'g1', name: '測試集團', memberCompanies: ['甲公司'], createdAt: nowIso }],
  };
}
// 第一階段(基準)的 23 個功能鍵：模擬「管理員在新功能出現前就存過權限矩陣」
const KNOWN_BASE_FEATURES = ['home', 'managerHome', 'execDash', 'prospects', 'contacts', 'companyMaster', 'visits', 'targets', 'forecast', 'pipeline', 'pipelineReport', 'bizAnalysis', 'yoy', 'contractGroup', 'accountingGroup', 'lostOpp', 'keyAccount', 'transfer', 'campaigns', 'leads', 'quotations', 'marketingHome', 'consultHome'];
const EXP = { ERP: ['w1', 'w2', 'w3', 'w4'], BOTH: ['n3', 'n8', 'w1', 'w2', 'w3', 'w4', 'x1'], ITS: ['n3', 'n8', 'x2'] };   // 一家公司一列：丙公司 w3/w3b、跨公司 x1/x2 各留 id 最小者（建立時間相同、都沒設主要）
function seedFiles(dir, over) {
  dir = dir || ROOT;
  let users = ACCTS.map((a) => mkUser(a[0], a[1], a[2])).concat([mkUser('_s_seco', 'secretary', { bu: ['ERP'], active: false })]);
  let data = seedData();
  if (over && over.data) data = over.data(data);
  if (over && over.users) users = over.users(users);
  fs.writeFileSync(path.join(dir, 'auth.json'), JSON.stringify({ users }, null, 2));
  fs.writeFileSync(path.join(dir, 'data.json'), JSON.stringify(data, null, 2));
}
const readData = (dir) => JSON.parse(fs.readFileSync(path.join(dir || ROOT, 'data.json'), 'utf8'));
function cleanupFiles(dir) { FILES.forEach((f) => { try { fs.unlinkSync(dir ? path.join(dir, path.basename(f)) : f); } catch { /* 不存在 */ } }); }

// ───────────────── 伺服器（預載腳本在載入時改寫指定檔案＝變異測試，不動磁碟）─────────────────
const PRELOAD = path.join(os.tmpdir(), 'scu_preload_' + process.pid + '.js');
fs.writeFileSync(PRELOAD, [
  "const Module=require('module'),path=require('path');const orig=Module.prototype._compile;",
  "Module.prototype._compile=function(content,filename){const raw=process.env.SCU_MUT_SPEC;",
  "if(raw){const specs=JSON.parse(raw);for(const m of specs){if(path.resolve(filename)===path.resolve(process.env.SCU_ROOT,m.file)){for(const e of m.edits){if(!content.includes(e.find))throw new Error('MUTATION ANCHOR MISSING: '+e.find.slice(0,60));content=content.replace(e.find,()=>e.rep);}}}}",
  "return orig.call(this,content,filename);};",
].join('\n') + '\n');
let child = null, childPid = null, MUT_CUR = null;
async function startServer(opts) {
  opts = opts || {};
  const root = opts.root || ROOT;
  const mut = opts.mut || (root === ROOT ? MUT_CUR : null);
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: `http://localhost:${PORT}` });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'SCU_MUT_SPEC', 'SECRETARY_CUSTOMERS', 'FEATURE_GUARD_MODE', 'LOGIN_LOCKOUT']) delete env[k];
  Object.assign(env, opts.env || {});
  // 全域 API 速率限制（300/分鐘/IP）會讓大量比對吃 429：測試行程一律把上限調高（基準版與受測版同樣處理，只影響限流本身）
  const spec = (mut || []).slice();
  spec.push({ file: 'server.js', edits: [{ find: '  max: 300,\n  standardHeaders', rep: '  max: 1e9,\n  standardHeaders' }] });
  env.SCU_MUT_SPEC = JSON.stringify(spec); env.SCU_ROOT = root;
  await assertPortFree(PORT);
  child = spawn(process.execPath, ['-r', PRELOAD, path.join(root, 'server.js')], { cwd: root, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { out += d; });
  for (let i = 0; i < 100; i++) {
    await sleep(250);
    try { const r = await fetch(`${BASE}/login.html`); if (r.ok) return; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-600));
  }
  throw new Error('伺服器啟動逾時：' + out.slice(-600));
}
async function stopServer() {
  if (!child) return;
  try { process.kill(childPid); } catch { /* 已結束 */ }
  for (let i = 0; i < 30 && child.exitCode === null; i++) await sleep(100);
  child = null; await sleep(300);
}
let jar = {};
/** opts: { root, env, mut, seed:{data,users} } —— 重新播種並啟動；jar 歸零 */
async function freshServer(opts) { opts = opts || {}; await stopServer(); cleanupFiles(opts.root); seedFiles(opts.root, opts.seed); await startServer(opts); jar = {}; }
async function http(who, method, url, body) {
  const h = {};
  if (body !== undefined) h['Content-Type'] = 'application/json';
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  const ct = res.headers.get('content-type') || '';
  let text = buf.toString('utf8');
  if (/spreadsheet|excel|octet-stream/i.test(ct) && buf.length > 4 && buf[0] === 0x50 && buf[1] === 0x4b) {
    try { const wb = XLSX.read(buf, { type: 'buffer' }); text = wb.SheetNames.map((n) => n + '\n' + XLSX.utils.sheet_to_csv(wb.Sheets[n])).join('\n'); } catch { /* 保留原文 */ }
  }
  return { status: res.status, json, buf, text };
}
async function login(who) {
  const res = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status === 200) jar[who] = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
  return res.status;
}
async function loginAll(names) { for (const w of names) { const st = await login(w); if (st !== 200) throw new Error('登入失敗 ' + w + ' ' + st); } }
const logsOf = async () => { await sleep(300); const r = await http('_s_adm', 'GET', '/api/admin/logs'); return Array.isArray(r.json) ? r.json : (r.json && r.json.logs) || []; };
const idsOf = (r) => ((r.json && r.json.customers) || []).map((c) => c.id).sort();

// ───────────────── 載入（可變異的）lib ─────────────────
function loadLib(edits) {
  let src = fs.readFileSync(LIBPATH, 'utf8');
  for (const e of edits || []) { if (!src.includes(e.find)) throw new Error('MUTATION ANCHOR MISSING: ' + e.find.slice(0, 60)); src = src.replace(e.find, () => e.rep); }
  const m = new Module(LIBPATH, null); m.filename = LIBPATH; m.paths = Module._nodeModulePaths(path.dirname(LIBPATH));
  m._compile(src, LIBPATH);
  return m.exports;
}
const LIB_MUT = () => (MUT_CUR || []).filter((m) => path.resolve(ROOT, m.file) === LIBPATH).flatMap((m) => m.edits);

// 取出前端 getWonContactIds（app.js），在 vm 中以給定資料執行
function frontendWon(contacts, opps) {
  const app = fs.readFileSync(path.join(ROOT, '_client', 'app.js'), 'utf8');
  const m = app.match(/function getWonContactIds\(\) \{[\s\S]*?\n\}\n/);
  if (!m) throw new Error('找不到前端 getWonContactIds');
  const ctx = { allContacts: contacts, allOpportunities: opps };
  vm.createContext(ctx);
  vm.runInContext(m[0] + '\nthis.__r = Array.from(getWonContactIds());', ctx);
  return ctx.__r.filter(Boolean).sort();
}
function rng(seed) { let s = seed >>> 0; return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; }; }

// ═════════════ unit ═════════════
function groupUnit() {
  console.log('[unit] lib 純函式、前後端對拍、靜態檢查');
  const L = loadLib(LIB_MUT());
  const read = (rel) => fs.readFileSync(path.join(ROOT, rel), 'utf8');
  // 對拍：隨機資料，後端 wonContactIds 與前端 getWonContactIds 集合相同
  const R = rng(20261103);
  let diffs = 0, cases = 0, nonEmpty = 0;
  for (let t = 0; t < 400; t++) {
    const n = 1 + Math.floor(R() * 12);
    const cos = ['甲', '乙', '丙', '', '丁'];
    const contacts = Array.from({ length: n }, (_, i) => ({ id: 'c' + i, company: cos[Math.floor(R() * cos.length)], customerType: R() < 0.25 ? 'customer' : (R() < 0.2 ? 'prospect' : '') }));
    const opps = Array.from({ length: Math.floor(R() * 6) }, (_, i) => ({ id: 'o' + i, contactId: R() < 0.1 ? undefined : 'c' + Math.floor(R() * (n + 2)), stage: ['Won', 'A', 'B', 'C', 'D', 'Lost'][Math.floor(R() * 6)] }));
    const f = frontendWon(contacts, opps), b = Array.from(L.wonContactIds(contacts, opps)).sort();
    cases++; if (f.length) nonEmpty++;
    if (JSON.stringify(f) !== JSON.stringify(b)) { diffs++; if (diffs <= 2) console.log('    對拍差異', JSON.stringify({ contacts, opps, f, b })); }
  }
  ok(diffs === 0 && nonEmpty > 100, `U1 前後端「已成交」對拍：${cases} 組隨機資料集合全部相同（非空 ${nonEmpty} 組）`, diffs + ' 組不同');
  // 定點
  const c = [{ id: 'a', company: 'X', customerType: 'customer' }, { id: 'b', company: 'X' }, { id: 'c', company: 'Y' }, { id: 'd', company: '', customerType: '' }, { id: 'e', company: 'Z' }];
  eq(Array.from(L.wonContactIds(c, [{ contactId: 'c', stage: 'Won' }, { contactId: 'd', stage: 'A' }])).sort(), ['a', 'b', 'c'], 'U2 customer／同公司／Won 商機；非 Won 商機與空公司不算');
  // 白名單
  const raw = mkCard('t1', '_s_u1', '測試公司', { customerType: 'customer', address: 'A路', mailingAddress: '' });
  const v = L.viewOf(raw, '業務名');
  eq(Object.keys(v).sort(), L.VIEW_KEYS.slice().sort(), 'U3 viewOf 只輸出白名單鍵');
  eq(L.VIEW_KEYS.filter((k) => /phone|ext|mobile|tel|email|note|card|memo/i.test(k)).sort(), ['ext', 'mobile', 'phone'], 'U3b 白名單裡與聯絡方式有關的鍵只有 phone／ext／mobile（Email／備註／名片圖仍不在）');
  eq([L.viewOf(Object.assign({}, raw, { phone: ' 02-0000-0000 ', ext: 123, mobile: null }), 'x')].map((q) => [q.phone, q.ext, q.mobile])[0], ['02-0000-0000', '123', ''], 'U3c 電話欄位字串化＋trim、缺值為空字串');
  ok(!PII_RE.test(JSON.stringify(v)), 'U4 viewOf 輸出不含任何 PII 哨兵', JSON.stringify(v));
  const monster = Object.assign({}, raw, { evilKey: 'EVIL_PII_', password: 'x', cardImage: 'CARD_PII_' });
  ok(!JSON.stringify(L.viewOf(monster, 'x')).includes('EVIL_PII_'), 'U5 未知欄位不會被帶出（白名單而非黑名單）');
  eq([v.mailingAddress, v.mailingFallback], ['A路', true], 'U6 沒填郵寄地址 → 公司地址＋fallback 旗標');
  const v2 = L.viewOf(Object.assign({}, raw, { mailingAddress: ' 郵寄路 ' }), 'x');
  eq([v2.mailingAddress, v2.mailingFallback], ['郵寄路', false], 'U7 有郵寄地址 → 用它（trim），不是 fallback');
  const v3 = L.viewOf(Object.assign({}, raw, { address: '', mailingAddress: '' }), 'x');
  eq([v3.mailingAddress, v3.mailingFallback], ['', false], 'U8 兩者都空 → 空字串、不標 fallback');
  // 範圍
  const users = ACCTS.map((a) => mkUser(a[0], a[1], a[2]));
  const lf = (me) => L.listFor(seedData(), users, users.find((u) => u.username === me), (n) => n).map((x) => x.id).sort();
  eq(lf('_s_sec'), EXP.ERP, 'U9 ERP 秘書：同 BU＋已成交，排除未成交／刪除／待補／非業務擁有者／其他 BU');
  eq(lf('_s_secx'), EXP.ITS, 'U10 ITS 秘書');
  eq(lf('_s_sec2'), EXP.BOTH, 'U11 兩個 BU 的秘書：兩邊都看（且跨 BU 同公司會互相成立）');
  eq(lf('_s_secnb'), [], 'U12 沒有 BU 的秘書：空');
  eq(L.listFor(seedData(), users, null, null), [], 'U13 找不到帳號 → 空（fail-safe）');
  // 靜態
  const server = read('server.js');
  ok(!/innerHTML|insertAdjacentHTML|document\.write|outerHTML/.test(read('_client/secretary-customers.js').replace(/\/\*[\s\S]*?\*\//, '')), 'U14 前端新檔沒有 innerHTML 等危險寫法');
  const html = read('_client/index.html');
  const scripts = [...html.matchAll(/<script\s+src="([^"]+\.js)"/g)].map((m) => m[1]);
  ok(scripts.indexOf('secretary-customers.js') > scripts.indexOf('app.js') && scripts.indexOf('secretary-customers.js') > 0, 'U15 index.html：secretary-customers.js 在 app.js 之後');
  ok(/itts-crm-v42'/.test(read('_client/sw.js')), 'U16 SW_VERSION 為 v42');
  ok(/#SECRETARY_CUSTOMERS=1/.test(read('.env.example')), 'U17 .env.example 有記載 SECRETARY_CUSTOMERS');
  ok(/if \(SCU\.enabled\(\)\) FEATURE_REGISTRY\.push/.test(server) && /if \(SCU\.enabled\(\)\) DEFAULT_ROLE_PERMISSIONS\.secretary/.test(server), 'U18 功能鍵與預設權限只在開關開啟時登錄');
  // getViewableOwners 與基準逐字相同（不得改動：getEditableOwners 重用它）
  const gv = (t) => { const i = t.indexOf('function getViewableOwners('); const j = t.indexOf('\n}\n', i); return t.slice(i, j); };
  const baseSrv = spawnSync('git', ['show', BASE_REF + ':server.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024 }).stdout.toString('utf8');
  ok(gv(server).length > 500 && gv(server) === gv(baseSrv), 'U19 getViewableOwners 與基準版逐字相同');
  ok(/app\.get\('\/api\/secretary\/customers'/.test(server) && !/app\.(put|delete|patch)\('\/api\/secretary/.test(server), 'U20 只有 GET 清單與 POST view 兩條路由，沒有寫入型路由');
  // 一家公司一列（主要聯絡人）
  const pk = (cards) => { const r = L.pickPrimary(cards); return r && [r.card.id, r.flag]; };
  const cc = (id, extra) => Object.assign({ id, owner: '_s_u1', company: '某公司', createdAt: '2026-01-02T00:00:00.000Z' }, extra || {});
  eq(pk([cc('a')]), ['a', 'single'], 'U21 單張 → 視為主要（single，即使 isPrimary 未設）');
  eq(pk([cc('a', { isPrimary: false })]), ['a', 'single'], 'U21b 單張且 isPrimary=false 仍是 single');
  eq(pk([cc('a'), cc('b', { isPrimary: true }), cc('c')]), ['b', 'primary'], 'U22 多張恰一個主要 → 取它（即使不是最早）');
  eq(pk([cc('b', { createdAt: '2026-03-01T00:00:00.000Z' }), cc('a', { createdAt: '2026-02-01T00:00:00.000Z' }), cc('c', { createdAt: '2026-01-15T00:00:00.000Z' })]), ['c', 'unset'], 'U23 多張都沒主要 → 最早建立者＋unset（最早者排中間，避免剛好靠輸入順序）');
  eq(pk([cc('x', { isPrimary: true, createdAt: '2026-05-01T00:00:00.000Z' }), cc('y', { createdAt: '2026-01-01T00:00:00.000Z' }), cc('z', { isPrimary: true, createdAt: '2026-04-01T00:00:00.000Z' })]), ['z', 'multi'], 'U24 多張兩個以上主要 → 最早建立的主要＋multi（不選更早但非主要的 y）');
  eq(pk([cc('b', { createdAt: undefined }), cc('a', { createdAt: undefined })]), ['a', 'unset'], 'U25 時間全缺失 → 以 id 排序，穩定');
  eq(pk([cc('b', { createdAt: 'garbage' }), cc('c', { createdAt: '2026-09-09T00:00:00.000Z' })]), ['c', 'unset'], 'U25b 有時間者先於缺失／無法解析者');
  eq([pk([cc('b'), cc('a'), cc('c')]), pk([cc('c'), cc('b'), cc('a')])], [['a', 'unset'], ['a', 'unset']], 'U26 時間相同 → id 當 tie-break，與輸入順序無關');
  eq(pk([cc('q', { isPrimary: true }), cc('p', { isPrimary: true })]), ['p', 'multi'], 'U26b 多主要且時間相同 → id 小者');
  eq(pk([]), null, 'U26c 空集合 → null');
  {
    // 範圍：只在秘書可視範圍內選；範圍外（其他 BU／刪除／待補／非業務）名片不影響選擇
    const d = seedData();
    d.contacts = [
      mkCard('s1', '_s_u1', '丙公司', { createdAt: '2026-03-01T00:00:00.000Z', customerType: 'customer' }),
      mkCard('s2', '_s_u2', '丙公司', { createdAt: '2026-04-01T00:00:00.000Z' }),
      mkCard('o1', '_s_u3', '丙公司', { bu: 'ITS', createdAt: '2026-01-01T00:00:00.000Z', isPrimary: true }),   // 其他 BU：ERP 秘書看不到
      mkCard('o2', '_s_u1', '丙公司', { deleted: true, createdAt: '2026-01-01T00:00:00.000Z', isPrimary: true }),
      mkCard('o3', '_pool', '丙公司', { isPlaceholder: true, createdAt: '2026-01-01T00:00:00.000Z', isPrimary: true }),
      mkCard('o4', '_s_mkt', '丙公司', { createdAt: '2026-01-01T00:00:00.000Z', isPrimary: true }),
      mkCard('e1', '_s_u1', '', { customerType: 'customer' }), mkCard('e2', '_s_u2', '', { customerType: 'customer' }),
    ];
    d.opportunities = [];
    const row = (me) => L.listFor(d, users, users.find((u) => u.username === me), (n) => n).filter((x) => x.company === '丙公司').map((x) => [x.id, x.primaryFlag]);
    eq(row('_s_sec'), [['s1', 'unset']], 'U27 ERP 秘書：範圍外的主要名片（其他 BU／刪除／待補／行銷）不被選中、也不影響旗標');
    eq(row('_s_sec2'), [['o1', 'primary']], 'U28 ERP+ITS 秘書：其他 BU 名片進入範圍後才可能被選（且只有它是主要）');
    { const r1 = L.listFor(d, users, users.find((u) => u.username === '_s_sec'), (n) => n).find((x) => x.company === '丙公司'); eq([r1.phone, r1.ext, r1.mobile], ['TEL_PII_s1', 'EXT_PII_s1', 'MOB_PII_s1'], 'U27b 電話三欄是被選中的主要聯絡人 s1 自己的，不混入同公司 s2'); ok(!JSON.stringify(r1).includes('_PII_s2'), 'U27c 同公司其他人的電話不在輸出'); }
    eq(L.listFor(d, users, users.find((u) => u.username === '_s_sec'), (n) => n).filter((x) => !x.company).map((x) => x.id).sort(), ['e1', 'e2'], 'U29 company 為空的名片不歸併，各自一列');
    eq(L.listFor(d, users, users.find((u) => u.username === '_s_sec'), (n) => n).every((x) => ['single', 'primary', 'unset', 'multi'].includes(x.primaryFlag)), true, 'U30 每列都有合法 primaryFlag');
    const cos = L.listFor(seedData(), users, users.find((u) => u.username === '_s_sec2'), (n) => n).map((x) => x.company);
    eq(cos.length, new Set(cos).size, 'U31 同一份清單中公司不重複');
  }
}

// ═════════════ api ═════════════
async function groupApi() {
  console.log('[api] 開關開：範圍／白名單／PII／權限／稽核／寫入');
  await freshServer({ env: { SECRETARY_CUSTOMERS: '1' } });
  await loginAll(['_s_adm', '_s_sec', '_s_sec2', '_s_secx', '_s_secnb', '_s_secp', '_s_u1', '_s_mkt', '_s_m1', '_s_gs']);
  const L = loadLib([]);
  const r = await http('_s_sec', 'GET', '/api/secretary/customers');
  eq(r.status, 200, 'A1 秘書 GET /api/secretary/customers = 200');
  eq(idsOf(r), EXP.ERP, 'A2 ERP 秘書的清單 = 同 BU 已成交客戶（一家公司一列：w1,w2,w3,w4）');
  { const m = Object.fromEntries((r.json.customers || []).map((c) => [c.id, c.primaryFlag])); eq([m.w1, m.w3, m.w4], ['single', 'unset', 'single'], 'A2b 旗標：單張 single、丙公司兩張無主要 unset'); const cs = r.json.customers.map((c) => c.company); eq(cs.length, new Set(cs).size, 'A2c 每家公司只有一列'); }
  const text = r.text;
  ok(!PII_RE.test(text), 'A3 秘書回應整串 JSON 沒有 Email／備註／名片圖／偏好等 PII 哨兵', (text.match(PII_RE) || [])[0]);
  {
    const found = [...new Set(text.match(PHONE_RE) || [])].sort();
    eq(found, ['EXT_PII_w1', 'EXT_PII_w2', 'EXT_PII_w3', 'MOB_PII_w1', 'MOB_PII_w2', 'MOB_PII_w3', 'TEL_PII_w1', 'TEL_PII_w2', 'TEL_PII_w3'], 'A3b 電話／分機／手機只出現在被選中的主要聯絡人（w1,w2,w3）；同公司非主要（w3b）與其他人的電話不得出現');
    ok(!/_PII_w3b|_PII_x|_PII_n/.test(text), 'A3c 同公司非主要聯絡人 w3b 的電話／分機／手機不在回應中');
    const by = Object.fromEntries(r.json.customers.map((c) => [c.id, c]));
    eq([by.w3.phone, by.w3.ext, by.w3.mobile], ['TEL_PII_w3', 'EXT_PII_w3', 'MOB_PII_w3'], 'A3d 丙公司列的電話三欄＝主要聯絡人 w3 自己的值');
    eq([by.w4.phone, by.w4.ext, by.w4.mobile], ['', '', ''], 'A3e 缺值時三欄皆為空字串');
    ok(r.json.customers.every((c) => ['phone', 'ext', 'mobile'].every((k) => typeof c[k] === 'string')), 'A3f 三欄一律為字串');
  }
  ok(r.json && r.json.piiMasked === true && r.json.customers.every((c) => c.piiMasked === true), 'A4 回 piiMasked:true（外層與每筆）');
  ok(r.json.customers.every((c) => JSON.stringify(Object.keys(c).sort()) === JSON.stringify(L.VIEW_KEYS.slice().sort())), 'A5 每筆只有白名單鍵');
  eq(Object.keys(r.json).sort(), ['count', 'customers', 'piiMasked'], 'A5b 外層鍵');
  for (const bad of ['n1', 'n2', 'n3', 'n4', 'n5', 'n6', 'n8', 'x1', 'x2', 'ka1']) ok(!idsOf(r).includes(bad), `A6 ${bad}（未成交／刪除／待補／其他BU／非業務擁有者）不在清單`);
  const by = Object.fromEntries(r.json.customers.map((c) => [c.id, c]));
  eq([by.w1.mailingAddress, by.w1.mailingFallback, by.w1.mailingSameAsCompany], ['MAIL-A路1號', false, false], 'A7 郵寄地址情境①：有另填郵寄地址');
  eq([by.w2.mailingAddress, by.w2.mailingFallback, by.w2.mailingSameAsCompany], ['ADDR-B路2號', false, true], 'A8 郵寄地址情境②：同公司地址（旗標 true）');
  eq([by.w3.mailingAddress, by.w3.mailingFallback], ['ADDR-C路3號', true], 'A9 郵寄地址情境③：舊資料沒填 → 公司登記地址＋mailingFallback:true');
  eq([by.w4.mailingAddress, by.w4.mailingFallback], ['', false], 'A9b 郵寄與公司地址都沒填 → 空字串（前端顯示未填）');
  eq([by.w1.company, by.w1.name, by.w1.title, by.w1.ownerDisplayName, by.w1.bu], ['甲公司', '名片w1', '職稱w1', '甲哥', ['ERP']], 'A10 可見欄位：公司／姓名／職稱／負責業務顯示名（暱稱優先）／BU');
  // 多 BU／其他 BU／無 BU
  eq(idsOf(await http('_s_sec2', 'GET', '/api/secretary/customers')), EXP.BOTH, 'A11 兩個 BU 的秘書：ERP＋ITS 都看到');
  eq(idsOf(await http('_s_secx', 'GET', '/api/secretary/customers')), EXP.ITS, 'A12 ITS 秘書只看 ITS');
  const nb = await http('_s_secnb', 'GET', '/api/secretary/customers');
  eq([nb.status, idsOf(nb)], [200, []], 'A13 沒有 BU 的秘書：200 但空清單');
  eq(idsOf(await http('_s_secp', 'GET', '/api/secretary/customers')), EXP.ERP, 'A14 掛名（pureSupervisor）秘書：與一般秘書相同');
  eq(await login('_s_seco'), 401, 'A15 停用的秘書無法登入');
  // 其他角色
  for (const who of ['_s_adm', '_s_u1', '_s_mkt', '_s_m1', '_s_gs']) {
    const x = await http(who, 'GET', '/api/secretary/customers');
    ok([401, 403].includes(x.status) && !PII_RE.test(x.text) && !x.text.includes('甲公司'), `A16 ${who} 不是秘書：${x.status}，無資料`);
    const v = await http(who, 'POST', '/api/secretary/customers/w1/view', {});
    ok([401, 403].includes(v.status), `A16b ${who} POST view：${v.status}`);
  }
  eq((await http(null, 'GET', '/api/secretary/customers')).status, 401, 'A17 未登入 401');
  // 稽核
  const before = (await logsOf()).filter((l) => l.action === 'SECRETARY_VIEW_CUSTOMER').length;
  eq(before, 0, 'A18 只 GET 清單不寫稽核');
  const v1 = await http('_s_sec', 'POST', '/api/secretary/customers/w1/view', {});
  eq(v1.status, 200, 'A19 開啟詳情 200');
  ok(v1.json && v1.json.customer && v1.json.customer.id === 'w1' && !PII_RE.test(v1.text) && v1.json.customer.phone === 'TEL_PII_w1' && v1.json.customer.mobile === 'MOB_PII_w1', 'A19b 詳情回白名單內容：電話／手機＝本人、無其他 PII');
  await http('_s_sec', 'POST', '/api/secretary/customers/w1/view', {});
  await http('_s_sec', 'POST', '/api/secretary/customers/w2/view', {});
  const lg = (await logsOf()).filter((l) => l.action === 'SECRETARY_VIEW_CUSTOMER');
  eq(lg.length, 2, 'A20 同人同客戶 10 分鐘內只記一筆（w1 兩次 + w2 一次 = 2 筆）');
  const row = lg.find((l) => l.target === 'w1');
  ok(row && /甲公司/.test(row.detail) && row.operator === '_s_sec', 'A21 稽核記聯絡人 id 與公司名、操作者', JSON.stringify(row));
  ok(!lg.some((l) => PII_RE.test(JSON.stringify(l)) || /名片w1|TEL_|@/.test(l.detail || '')), 'A22 稽核內容不含電話／Email／備註／聯絡人姓名');
  eq((await http('_s_sec', 'POST', '/api/secretary/customers/n1/view', {})).status, 404, 'A23 不在範圍內的客戶（未成交）詳情 404');
  eq((await http('_s_sec', 'POST', '/api/secretary/customers/n3/view', {})).status, 404, 'A24 其他 BU 的客戶詳情 404（不能憑 id 猜）');
  eq((await logsOf()).filter((l) => l.action === 'SECRETARY_VIEW_CUSTOMER').length, 2, 'A25 404 不寫稽核');
  for (const [m, u] of [['PUT', '/api/secretary/customers/w1'], ['DELETE', '/api/secretary/customers/w1'], ['POST', '/api/secretary/customers'], ['GET', '/api/secretary/customers/w1']]) {
    const x = await http('_s_sec', m, u, m === 'GET' || m === 'DELETE' ? undefined : {});
    ok([404, 405].includes(x.status), `A26 ${m} ${u} 不存在（${x.status}）`);
  }
  // 寫入端點：秘書改不了別人的聯絡人
  const rawBefore = JSON.stringify(readData().contacts);
  const w = {};
  w.put = await http('_s_sec', 'PUT', '/api/contacts/w1', { note: 'HACK', mailingAddress: 'HACK' });
  w.del = await http('_s_sec', 'DELETE', '/api/contacts/w1');
  w.put2 = await http('_s_sec', 'PUT', '/api/contacts/w2', { name: 'HACK', company: '乙公司' });
  w.tr = await http('_s_sec', 'POST', '/api/transfer-contacts', { fromOwner: '_s_u1', toOwner: '_s_sec', contactIds: ['w1'] });
  w.bulk = await http('_s_sec', 'POST', '/api/admin/pool/assign', { ids: ['w1'], owner: '_s_sec' });
  w.ka = await http('_s_sec', 'POST', '/api/key-accounts', { company: '甲公司' });
  for (const k of ['put', 'del', 'put2', 'tr', 'bulk']) ok([400, 403, 404].includes(w[k].status), `A27 秘書寫入 ${k} 被擋（${w[k].status}）`);
  eq(JSON.stringify(readData().contacts), rawBefore, 'A28 資料檔中的聯絡人完全沒變');
  // 唯讀聯絡人端點的現況（與本功能無關；基準比對在 roles 群組）
  const gc = await http('_s_sec', 'GET', '/api/contacts');
  const gcIds = (gc.json || []).map((c) => c.id).sort();
  ok(!gcIds.some((i) => EXP.ERP.includes(i)), 'A29 GET /api/contacts 對秘書沒有因本功能多出任何同BU已成交客戶（現況只有 KA 公司的聯絡人：' + gcIds.join(',') + '）');
  // 功能鍵與權限矩陣
  const mp = await http('_s_sec', 'GET', '/api/me/permissions');
  ok(mp.json.features.includes('secretaryCustomers'), 'A30 秘書預設有 secretaryCustomers 功能');
  ok(!(await http('_s_u1', 'GET', '/api/me/permissions')).json.features.includes('secretaryCustomers'), 'A31 業務沒有此功能');
  const rp = await http('_s_adm', 'GET', '/api/admin/role-permissions');
  ok(rp.json.features.some((f) => f.key === 'secretaryCustomers') && rp.json.matrix.admin.includes('secretaryCustomers') && rp.json.matrix.secretary.includes('secretaryCustomers') && Object.keys(rp.json.matrix).filter((k) => rp.json.matrix[k].includes('secretaryCustomers')).sort().join() === 'admin,secretary', 'A32 管理員矩陣：只有 admin（永遠全開）與 secretary 含此功能');
  // 管理員把功能從秘書拿掉 → 秘書 403
  const m2 = JSON.parse(JSON.stringify(rp.json.matrix)); m2.secretary = m2.secretary.filter((k) => k !== 'secretaryCustomers');
  const sv = await http('_s_adm', 'PUT', '/api/admin/role-permissions', { matrix: m2 });
  eq(sv.status, 200, 'A33 管理員存權限矩陣');
  eq((await http('_s_sec', 'GET', '/api/secretary/customers')).status, 403, 'A34 秘書被拿掉功能 → 403（儲存後不會被自動補回）');
  await stopServer();

  // 舊矩陣（開關開啟前就存過、不含新功能）→ 自動補上給預設含它的角色
  await freshServer({ env: { SECRETARY_CUSTOMERS: '1' }, seed: { data: (d) => Object.assign(d, { rolePermissions: { secretary: ['home', 'quotations'], user: ['home', 'contacts'] }, rolePermissionsKnownFeatures: KNOWN_BASE_FEATURES }) } });
  await loginAll(['_s_sec', '_s_u1']);
  const o1 = await http('_s_sec', 'GET', '/api/me/permissions');
  eq(o1.json.features.slice().sort(), ['home', 'quotations', 'secretaryCustomers'], 'A35 已存過的舊權限矩陣：秘書自動補上全新功能，其餘維持管理員設定');
  eq((await http('_s_u1', 'GET', '/api/me/permissions')).json.features, ['home', 'contacts'], 'A36 業務的舊矩陣不受影響');
  eq((await http('_s_sec', 'GET', '/api/secretary/customers')).status, 200, 'A37 補上後秘書可用');
  await stopServer();
  // FEATURE_GUARD_MODE=enforce：非秘書被功能守門擋（403），秘書通過
  await freshServer({ env: { SECRETARY_CUSTOMERS: '1', FEATURE_GUARD_MODE: 'enforce' } });
  await loginAll(['_s_sec', '_s_u1']);
  const e1 = await http('_s_u1', 'GET', '/api/secretary/customers');
  ok(e1.status === 403 && e1.json && e1.json.code === 'FEATURE_DENIED', 'A38 enforce：業務被 featureGuard 擋（FEATURE_DENIED）', e1.text);
  eq((await http('_s_sec', 'GET', '/api/secretary/customers')).status, 200, 'A39 enforce：秘書通過');
  eq((await http('_s_sec', 'POST', '/api/secretary/customers/w1/view', {})).status, 200, 'A39b enforce：秘書 POST view 通過');
  await stopServer(); cleanupFiles();
}

// ═════════════ roles（23 帳號 × 唯讀端點電池，三方逐位元比對）═════════════
function exportBase() {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'scu-base-'));
  const a = spawnSync('git', ['archive', BASE_REF], { cwd: ROOT, maxBuffer: 256 * 1024 * 1024 });
  if (a.status !== 0) throw new Error('git archive ' + BASE_REF + ' 失敗：' + String(a.stderr));
  const t = spawnSync('tar', ['-x', '-C', dir], { input: a.stdout, maxBuffer: 256 * 1024 * 1024 });
  if (t.status !== 0) throw new Error('tar 失敗：' + String(t.stderr));
  spawnSync('cmd', ['/c', 'mklink', '/J', path.join(dir, 'node_modules'), path.join(ROOT, 'node_modules')], { stdio: 'ignore' });
  return { dir, cleanup() { spawnSync('cmd', ['/c', 'rmdir', path.join(dir, 'node_modules')], { stdio: 'ignore' }); try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ } } };
}
const norm = (x) => String(x).replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/g, '<uuid>').replace(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/g, '<ts>').replace(/\d{4}\/\d{1,2}\/\d{1,2}[^"]*?(?=")/g, '<tl>').replace(/"(?:iat|exp|ts|now|generatedAt|loginAt)":\s*\d+/g, '"<n>"');
const BATTERY = [
  '/api/me', '/api/me/permissions', '/api/contacts', '/api/contacts?search=NOTE_PII', '/api/contacts?search=pii_', '/api/export', '/api/groups', '/api/key-accounts',
  '/api/companies', '/api/companies/c1', '/api/companies/c1/yoy', '/api/companies/duplicates', '/api/companies/ta/options', '/api/companies/export',
  '/api/company-master', '/api/company-similar?name=%E7%94%B2', '/api/company-requests', '/api/company-responsible/candidates',
  '/api/contacts-by-owner?owner=_s_u1', '/api/contacts-by-owner?owner=_s_u2', '/api/admin/pool/contacts', '/api/leads', '/api/campaigns',
  '/api/opportunities', '/api/visits', '/api/contracts', '/api/receivables', '/api/usermap', '/api/users-bu', '/api/notifications', '/api/poll-bundle',
  '/api/birthday-reminders', '/api/manager-home', '/api/manager/achievement', '/api/manager/company-card-overdue', '/api/manager/business-type-analysis',
  '/api/pipeline-report', '/api/pipeline-date-changes', '/api/zombie-opportunities', '/api/lost-opportunities', '/api/my-tasks/company-cards', '/api/targets',
  '/api/quotations', '/api/quotations/inbox', '/api/yoy/report', '/api/yoy/links', '/api/industries', '/api/product-catalog', '/api/pure-supervisors', '/api/plan',
  '/api/exec/conversion', '/api/exec/trend', '/api/exec/product-analysis', '/api/exec/bu-comparison', '/api/forecast/export', '/api/monthly-budget',
  '/api/settings/quarter-ratios', '/api/settings/opr-range', '/api/settings/user-quarter-ratios', '/api/marketing/responsibility-feedback',
  '/api/admin/role-permissions', '/api/admin/users', '/api/admin/feature-route-map',
];
const WRITES = [
  ['PUT', '/api/contacts/w1', { note: 'HACK' }], ['DELETE', '/api/contacts/w1'], ['PUT', '/api/contacts/w2', { name: 'HACK', company: '乙公司' }],
  ['POST', '/api/transfer-contacts', { fromOwner: '_s_u1', toOwner: '_s_u2', contactIds: ['w1'] }], ['POST', '/api/groups', { name: 'G' }], ['PUT', '/api/groups/g1', { name: 'G2' }],
  ['POST', '/api/key-accounts', { company: '甲公司' }], ['DELETE', '/api/key-accounts/ka-1'], ['POST', '/api/companies/c1/set-industry', { industry: '製造' }],
  ['POST', '/api/admin/pool/assign', { ids: ['w1'], owner: '_s_u2' }], ['POST', '/api/upload', {}], ['PUT', '/api/contacts/n3', { note: 'HACK' }],
];
const NEW_RE = /^\/api\/secretary\//;
async function runBattery(mode, root, batch) {
  await freshServer({ root, env: mode === 'on' ? { SECRETARY_CUSTOMERS: '1' } : {} });
  await loginAll(batch);
  const out = {};
  for (const who of batch) {
    for (const u of BATTERY) { const r = await http(who, 'GET', u); out[who + ' GET ' + u] = { status: r.status, text: norm(r.text) }; }
    out[who + ' GET /api/secretary/customers'] = { status: (await http(who, 'GET', '/api/secretary/customers')).status };
  }
  for (const who of batch.filter((x) => /^_s_(sec|u1|mkt|m1|exe|gs|tp)/.test(x))) {
    for (const [m, u, b] of WRITES) { const r = await http(who, m, u, b); out[who + ' ' + m + ' ' + u] = { status: r.status, text: norm(r.text) }; }
  }
  const raw = readData(root);
  out.__rawContacts = { text: norm(JSON.stringify(raw.contacts, null, 1)) };
  await stopServer(); cleanupFiles(root);
  return out;
}
async function groupRoles() {
  console.log('[roles] 23 帳號 × 唯讀端點電池 × 寫入電池：基準版 vs 開關關 vs 開關開（分兩批重啟避開登入限流）');
  const bx = exportBase();
  const batches = [ALL_ACCTS.slice(0, 12), ALL_ACCTS.slice(12)];
  const res = { base: {}, off: {}, on: {} };
  try {
    for (const b of batches) {
      for (const mode of ['base', 'off', 'on']) {
        const o = await runBattery(mode, mode === 'base' ? bx.dir : ROOT, b);
        for (const k of Object.keys(o)) res[mode][k + (k.startsWith('__') ? '@' + b[0] : '')] = o[k];
      }
    }
  } finally { bx.cleanup(); }
  const keys = Object.keys(res.base);
  const allKeys = keys, sameKeys = () => keys.filter((k) => !/GET \/api\/secretary\/customers$/.test(k));
  const cmp = (A, B, ks) => ks.filter((k) => JSON.stringify(A[k]) !== JSON.stringify(B[k]));
  // 1) 基準 vs 關
  const d1 = cmp(res.base, res.off, allKeys);
  ok(d1.length === 0, `R1 開關關：23 帳號 × 唯讀／寫入電池（${allKeys.length} 項，含探測新路徑）與基準版逐位元相同`, d1.slice(0, 4).join(' | '));
  if (d1.length) { const k = d1[0]; console.log('    差異首項：', k, '\n      基準：', JSON.stringify(res.base[k]).slice(0, 300), '\n      關：', JSON.stringify(res.off[k]).slice(0, 300)); }
  // 2) 基準 vs 開：允許的差異只有功能鍵登錄
  const d2 = cmp(res.base, res.on, sameKeys());
  const allowed = (k) => / GET \/api\/(me\/permissions|admin\/role-permissions|admin\/feature-route-map)$/.test(k);
  const unexpected = d2.filter((k) => !allowed(k));
  ok(unexpected.length === 0, `R2 開關開：除功能鍵登錄外所有端點輸出與基準版逐位元相同（差異 ${d2.length} 項、非預期 ${unexpected.length} 項）`, unexpected.slice(0, 4).join(' | '));
  if (unexpected.length) { const k = unexpected[0]; console.log('    差異首項：', k, '\n      基準：', JSON.stringify(res.base[k]).slice(0, 300), '\n      開：', JSON.stringify(res.on[k]).slice(0, 300)); }
  // 允許的差異：每一項都要符合預期的形狀
  for (const k of d2.filter(allowed)) {
    const who = k.split(' ')[0];
    const role = ACCTS.find((a) => a[0] === who)[1];
    const bj = safeJson(res.base[k].text), oj = safeJson(res.on[k].text);
    if (/me\/permissions$/.test(k)) {
      ok(role === 'secretary' || role === 'admin', `R3 /api/me/permissions 有差異者只有 secretary／admin（${who}）`);
      const exp = bj ? bj.features.concat(['secretaryCustomers']) : null;
      ok(bj && oj && JSON.stringify(oj.features.slice().sort()) === JSON.stringify(exp.slice().sort()) && JSON.stringify(Object.assign({}, bj, { features: 0 })) === JSON.stringify(Object.assign({}, oj, { features: 0 })), `R3b ${who} 只多了 secretaryCustomers 一個功能、其餘欄位相同`);
    } else {
      ok(role === 'admin', `R3c ${k} 只有 admin 看得到`);
    }
  }
  const diffPerm = d2.filter((k) => / GET \/api\/me\/permissions$/.test(k)).map((k) => ACCTS.find((a) => a[0] === k.split(' ')[0])[1]);
  ok(diffPerm.length >= 6 && diffPerm.every((r) => r === 'secretary' || r === 'admin') && ACCTS.filter((a) => a[1] === 'secretary').every((a) => d2.includes(a[0] + ' GET /api/me/permissions')), 'R4 開啟時 me/permissions 只有 secretary（5 帳號）與 admin（2 帳號）有變');
  // 3) 新端點狀態
  const st = (mode, who) => (res[mode][who + ' GET /api/secretary/customers'] || {}).status;
  const wantOn = (who) => ACCTS.find((a) => a[0] === who)[1] === 'secretary' ? 200 : 403;
  const badOn = ALL_ACCTS.filter((w) => st('on', w) !== wantOn(w));
  ok(badOn.length === 0, 'R5 開啟時新端點：只有 secretary（5 帳號）200、其餘 18 個帳號 403', badOn.map((w) => w + ':' + st('on', w)).join(','));
  const badOff = ALL_ACCTS.filter((w) => st('off', w) !== st('base', w) || ![404, 403].includes(st('off', w)) || (st('off', w) === 403 && !/^(tecopm|groupsales)$/.test(ACCTS.find((a) => a[0] === w)[1])));
  ok(badOff.length === 0, 'R6 關閉時新端點對 23 個帳號與基準版相同：404（集團 PM／集團業務因角色白名單本來就一律 403）', badOff.map((w) => w + ':' + st('off', w)).join(','));
  // 4) 秘書在開啟時：電池輸出相對基準，新出現的 PII 哨兵 = 0
  let newPii = 0, prePii = [];
  for (const k of Object.keys(res.on)) {
    if (!/^_s_sec/.test(k) || !res.on[k].text) continue;
    const onHas = PII_RE.test(res.on[k].text), baseHas = res.base[k] && PII_RE.test(res.base[k].text);
    if (onHas && !baseHas) { newPii++; console.log('    新增洩漏：', k); }
    if (baseHas) prePii.push(k.replace(/^_s_\w+ /, ''));
  }
  ok(newPii === 0, 'R7 秘書帳號在所有端點：開啟功能後沒有任何端點新增 PII 哨兵');
  console.log('    [盤點] 基準版(第一階段)秘書本來就看得到 PII 的端點（與本功能無關、未改動）：' + [...new Set(prePii)].join(', '));
  global.__scu_prePii = [...new Set(prePii)];
  const secWrites = Object.keys(res.on).filter((k) => /^_s_sec \w+ \/api\//.test(k) && !/ GET /.test(k));
  ok(secWrites.length > 0 && secWrites.every((k) => [400, 403, 404].includes(res.on[k].status) || res.on[k].status === res.base[k].status), 'R8 秘書所有寫入嘗試：不是 400/403/404，就是與基準版同一個狀態碼（沒有因本功能多出可寫範圍）');
  const secRows = secWrites.map((k) => k.replace('_s_sec ', '') + '=' + res.on[k].status);
  console.log('    [盤點] 秘書寫入嘗試結果：' + secRows.join(' | '));
  global.__scu_secWrites = secRows;
  global.__scu_count = { battery: BATTERY.length, writes: WRITES.length, keys: Object.keys(res.base).length };
}
function safeJson(t) { try { return JSON.parse(t); } catch { return null; } }

// ═════════════ off ═════════════
async function groupOff() {
  console.log('[off] 開關關：新端點 404、功能鍵不登錄、featureGuard 不含新路由');
  for (const mode of ['', 'enforce']) {
    await freshServer({ env: mode ? { FEATURE_GUARD_MODE: mode } : {} });
    await loginAll(['_s_adm', '_s_sec', '_s_u1']);
    for (const who of ['_s_sec', '_s_u1', '_s_adm']) {
      eq((await http(who, 'GET', '/api/secretary/customers')).status, 404, `O1 [${mode || '預設'}] ${who} GET 新端點 = 404`);
      eq((await http(who, 'POST', '/api/secretary/customers/w1/view', {})).status, 404, `O2 [${mode || '預設'}] ${who} POST view = 404`);
    }
    const mp = await http('_s_sec', 'GET', '/api/me/permissions');
    ok(!mp.json.features.includes('secretaryCustomers'), `O3 [${mode || '預設'}] 秘書功能清單沒有新功能`);
    const rp = await http('_s_adm', 'GET', '/api/admin/role-permissions');
    ok(!JSON.stringify(rp.json).includes('secretaryCustomers'), `O4 [${mode || '預設'}] 管理員矩陣沒有新功能`);
    const fm = await http('_s_adm', 'GET', '/api/admin/feature-route-map');
    ok(!JSON.stringify(fm.json.routes).includes('secretary/customers') && !fm.json.unmappedRegistered.length, `O5 [${mode || '預設'}] featureGuard 對照表輸出沒有新路由`);
    const lg = await logsOf();
    ok(!lg.some((l) => /SECRETARY_VIEW_CUSTOMER|FEATURE_GUARD_(WOULD_)?DENY/.test(l.action || '')), `O6 [${mode || '預設'}] 探測新路徑沒有產生 SECRETARY_VIEW_CUSTOMER／FEATURE_GUARD_DENY 稽核（與基準版對未知路徑相同，最多一筆 UNCLASSIFIED）`);
    await stopServer();
  }
  cleanupFiles();
}

// ═════════════ 變異 ═════════════
const LIBF = 'lib/secretaryCustomers.js', SRV = 'server.js';
const MUTANTS = [
  { name: 'M1 拿掉 BU 過濾（scopeContacts 一律放行 BU）', groups: ['unit', 'api'], mut: [{ file: LIBF, edits: [{ find: 'return contactBus(c, ou).some((b) => myBus.includes(b));', rep: 'return true;' }] }] },
  { name: 'M2 拿掉已成交過濾（listFor 回全部範圍內客戶）', groups: ['unit', 'api'], mut: [{ file: LIBF, edits: [{ find: 'scoped.filter((c) => won.has(c.id))', rep: 'scoped' }] }] },
  { name: 'M3 欄位白名單改成黑名單（展開原物件）', groups: ['unit', 'api'], mut: [{ file: LIBF, edits: [{ find: '  return {\n    id: c.id,', rep: '  return {\n    ...c,\n    id: c.id,' }] }] },
  { name: 'M4 開關判斷拿掉（關閉時仍註冊路由）', groups: ['off'], mut: [{ file: SRV, edits: [{ find: 'if (SCU.enabled()) {\n  const _scuViewAudit', rep: 'if (true) {\n  const _scuViewAudit' }] }] },
  { name: 'M5 寫入阻擋失效（PUT /api/contacts 開放秘書編輯同 BU 名片）', groups: ['api', 'roles'], mut: [{ file: SRV, edits: [{ find: "const editOwners = role === 'groupsales' ? getEditableOwners(req) : [owner];", rep: "const editOwners = (role === 'groupsales' || role === 'secretary') ? getEditableOwners(req) : [owner];" }] }] },
  { name: 'M6 角色檢查拿掉（任何登入者可用新端點）', groups: ['api', 'roles'], mut: [{ file: SRV, edits: [{ find: "if (u.role !== 'secretary' || !getUserFeatures(u.role).includes(SCU.FEATURE_KEY))", rep: 'if (false)' }] }] },
  { name: 'M7 功能鍵不登錄到預設權限（秘書沒有新功能）', groups: ['api'], mut: [{ file: SRV, edits: [{ find: "if (SCU.enabled()) DEFAULT_ROLE_PERMISSIONS.secretary = DEFAULT_ROLE_PERMISSIONS.secretary.concat([SCU.FEATURE_KEY]);", rep: '' }] }] },
  { name: 'M8 稽核寫進個資（detail 帶姓名）', groups: ['api'], mut: [{ file: SRV, edits: [{ find: '`秘書開啟客戶詳情：${c.company}`', rep: '`秘書開啟客戶詳情：${c.company} ${c.name}`' }] }] },
  { name: 'M9 已成交判定與前端分歧（後端漏掉同公司規則）', groups: ['unit'], mut: [{ file: LIBF, edits: [{ find: "(c.company && customerCompanies.has(c.company))", rep: 'false' }] }] },
  { name: 'M10 GET 清單也寫稽核（違反 GET 不寫入）', groups: ['api'], mut: [{ file: SRV, edits: [{ find: "    const customers = _scuList(g);\n", rep: "    const customers = _scuList(g);\n    writeLog('SECRETARY_VIEW_CUSTOMER', req.session.user.username, 'list', 'x', req);\n" }] }] },
  { name: 'M11 拿掉「單張視為主要」', groups: ['unit'], mut: [{ file: LIBF, edits: [{ find: "if (arr.length === 1) return { card: arr[0], flag: 'single' };", rep: '' }] }] },
  { name: 'M12 拿掉「無主要退回最早建立」（改取最後）', groups: ['unit', 'api'], mut: [{ file: LIBF, edits: [{ find: "card: arr.slice().sort(earliestCmp)[0], flag: 'unset'", rep: "card: arr.slice().sort(earliestCmp)[arr.length - 1], flag: 'unset'" }] }] },
  { name: 'M13 拿掉「多主要取最早」（改取最後）', groups: ['unit'], mut: [{ file: LIBF, edits: [{ find: "card: prim[0], flag: 'multi'", rep: "card: prim[prim.length - 1], flag: 'multi'" }] }] },
  { name: 'M14 拿掉 id tie-break', groups: ['unit'], mut: [{ file: LIBF, edits: [{ find: 'return ia < ib ? -1 : ia > ib ? 1 : 0;', rep: 'return 0;' }] }] },
  { name: 'M15 選主要不限範圍（改用全部名片分組）', groups: ['unit'], mut: [{ file: LIBF, edits: [{ find: "scoped.filter((c) => won.has(c.id)).forEach", rep: "(data.contacts || []).filter((c) => c && !c.deleted && won.has(c.id) || (c && !c.deleted && c.isPrimary)).forEach" }] }] },
  { name: 'M16 Email 也被放進輸出', groups: ['unit', 'api'], mut: [{ file: LIBF, edits: [{ find: '    phone: str(c.phone),', rep: '    phone: str(c.phone),\n    email: str(c.email),' }] }] },
  { name: 'M17 電話取成同公司非主要聯絡人的', groups: ['unit', 'api'], mut: [{ file: LIBF, edits: [{ find: '    v.bu = contactBus(c, byName.get(c.owner));', rep: "    v.bu = contactBus(c, byName.get(c.owner));\n    { const o = scoped.find((x) => x.company && x.company === c.company && x.id !== c.id); if (o) v.phone = str(o.phone); }" }] }] },
  { name: 'M18 備註也被放進輸出', groups: ['unit', 'api'], mut: [{ file: LIBF, edits: [{ find: '    phone: str(c.phone),', rep: '    phone: str(c.phone),\n    note: str(c.note),' }] }] },
];
async function runMutants() {
  console.log('[變異] 每個變異都必須讓對應群組變紅');
  const G = { unit: async () => groupUnit(), api: groupApi, roles: groupRoles, off: groupOff };
  for (const m of MUTANTS) {
    if (process.env.SCU_MUT_PICK && !new RegExp(process.env.SCU_MUT_PICK).test(m.name)) continue;
    const snap = { pass, fail, n: failures.length };
    quiet = true;
    try {
      for (const g of m.groups) {
        MUT_CUR = m.mut;
        try { await G[g](); } catch (e) { fail++; failures.push('變異執行例外：' + e.message); }
      }
    } finally { quiet = false; MUT_CUR = null; await stopServer(); cleanupFiles(); }
    const newFails = failures.length - snap.n;
    const names = failures.slice(snap.n, snap.n + 3).join(' / ');
    pass = snap.pass; fail = snap.fail; failures.length = snap.n;
    ok(newFails > 0, `${m.name} → 測試變紅（${newFails} 項失敗：${names.slice(0, 90)}）`);
    console.log('  ' + (newFails > 0 ? '✓ 變紅' : '✗ 沒變紅') + '：' + m.name + '（' + newFails + ' 項）');
  }
}

const api = { ok, eq, sleep, http, login, loginAll, freshServer, stopServer, cleanupFiles, readData, seedFiles, PW, BASE, ROOT, SHOTS, CDP, PORT, PII_RE, EXP, idsOf, logsOf, snap: () => ({ pass, fail, n: failures.length }), restore: (x) => { pass = x.pass; fail = x.fail; failures.length = x.n; }, newFails: (x) => failures.slice(x.n), setQuiet: (v) => { quiet = v; } };
module.exports = api;
if (require.main === module) {
  (async () => {
    try {
      if (want('unit')) groupUnit();
      if (want('api')) await groupApi();
      if (want('roles')) await groupRoles();
      if (want('off')) await groupOff();
      if (want('ui')) await require('./_secretary-customers-ui.js').run(api);
      if (process.env.SCU_MUT === '1') {
        await runMutants();
        if (want('ui')) await require('./_secretary-customers-ui.js').mutants(api);
      }
    } catch (e) {
      fail++; failures.push('例外：' + (e && e.stack || e)); console.log('例外：', e && e.stack || e);
    } finally {
      await stopServer(); cleanupFiles();
      try { fs.unlinkSync(PRELOAD); } catch { /* 不存在 */ }
    }
    console.log(`\n秘書唯讀客戶視圖檢查：PASS ${pass} / FAIL ${fail}`);
    if (fail) console.log('失敗項目：\n  ' + failures.join('\n  '));
    process.exit(fail ? 1 : 0);
  })();
}
