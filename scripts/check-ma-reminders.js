#!/usr/bin/env node
'use strict';
/**
 * MA／License 到期「活規則」提醒（P3a，MA_TRACKING=1）檢查：
 *   unit   規則驗證、門檻語意（正負天數／剛好等於／跳過補發）、四種條件、去重鍵、repeat 週期、收件人解析（多秘書／無秘書／無主管／停用／掛名主管／跨 BU）、台北日期（UTC 跨日）、掃描單元
 *   sim    日期推進模擬（注入 now 逐日前進 −10…+130 天）：每條規則恰好在該發的那天發一次；cron 漏跑補發；repeat；續約報價單建立／送簽／不續約／已續約後停止與自動解除
 *   volume 通知量：100 個納管項同時達標只發彙總、單次上限 200、第二次掃描不重複
 *   api    規則 GET／PUT／驗證／雙人覆核、試算器對拍、cron＋補掃、renewOf、續約成交、不續約／撤銷、時間軸、續約候選、失敗不影響 mail cron
 *   roles  23 帳號逐角色比對 maReminders、時間軸可視範圍與秘書欄位白名單
 *   off    MA_TRACKING 未設 ＝ 與基準版（git archive）逐位元相同
 * 變異測試（MAREM_MUT=1）：每個變異都必須讓對應群組變紅。
 * 不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體；資料檔只存在於本次執行；全部假資料；結束清掉伺服器。
 *
 *   node scripts/check-ma-reminders.js                        全部（不含變異）
 *   MAREM_ONLY=unit,sim,volume,api,roles,off node ...        只跑指定群組
 *   MAREM_MUT=1 node ...                                      另外跑變異測試
 *   MAREM_PORT（預設 3322；限 3320–3329）／MAREM_BASE_REF（開關比對的基準提交，預設 3d85397＝P2 驗收後）／MAREM_BASE_DIR
 * 守門：工作樹已有 auth.json／data.json／audit.log.json 時拒絕執行。一次只跑一個。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.MAREM_PORT || '3322', 10);
if (!(PORT >= 3320 && PORT <= 3329)) { console.log('拒絕：此腳本只能用埠 3320–3329'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map((f) => path.join(ROOT, f));
{
  const g = FILES.filter((f) => /(auth|data|audit.log)\.jsonl?$/.test(f)).filter((f) => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map((f) => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
}
const ONLY = (process.env.MAREM_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length || ONLY.includes(k);
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const REAL = { MR: require('../lib/maReminders.js'), RC: require('../lib/maRecipients.js'), MA: require('../lib/maTracking.js'), OR: require('../lib/orgResolve.js') };
const REF = { MR: REAL.MR, RC: REAL.RC, MA: REAL.MA };   // 變異測試時換成變異後的副本

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 500)); } }
const J = (x) => JSON.stringify(x);
const eq = (a, b, name) => ok(J(a) === J(b), name, `得到 ${J(a)}，預期 ${J(b)}`);
const sleep = (n) => new Promise((r) => setTimeout(r, n));
const clone = (x) => JSON.parse(JSON.stringify(x));

// ───────────────── 沙盒資料（全部假的）─────────────────
const PW = 'Mr' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const NOW = Date.now();
const nowIso = new Date(NOW).toISOString();
const DAY = 86400000;
const agoIso = (days) => new Date(NOW - days * DAY).toISOString();
const todayTp = () => new Date(Date.now() + 8 * 3600e3).toISOString().slice(0, 10);   // 獨立於 lib 的台北日期算法
const plusDays = (d, n) => new Date(Date.parse(d + 'T00:00:00Z') + n * DAY).toISOString().slice(0, 10);
const AMT = '8765432';            // 假金額（K）：不得出現在任何提醒、通知
const NOTE = '機密備註XYZ';        // 假備註
const CNO = 'SECRET-CN-777';      // 假合約編號：不得出現在秘書可見的提醒與時間軸
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false }, extra || {});
const U = {
  adm: '_s_adm', adm2: '_s_adm2', exe: '_s_exe', acc: '_s_acc', fin: '_s_fin', m1: '_s_m1', m1b: '_s_m1b', m1p: '_s_m1p', m1off: '_s_m1off', m2: '_s_m2',
  u1: '_s_u1', u2: '_s_u2', u3: '_s_u3', u4: '_s_u4', u5: '_s_u5', u6: '_s_u6', uoff: '_s_uoff', mkt: '_s_mkt', mkt2: '_s_mkt2', mktoff: '_s_mktoff',
  sec: '_s_sec', sec2: '_s_sec2', sec3: '_s_sec3', secoff: '_s_secoff', gs: '_s_gs', tp: '_s_tp', cs: '_s_cs',
};
function seedUsers() {
  return [
    mkUser(U.adm, 'admin', { displayName: '測試管理員' }), mkUser(U.adm2, 'admin', { displayName: '測試管理員二' }), mkUser(U.exe, 'executive'), mkUser(U.acc, 'accounting_manager'), mkUser(U.fin, 'finance_manager'),
    mkUser(U.m1, 'manager1', { bu: ['ERP'], displayName: '測試主管一' }),
    mkUser(U.m1b, 'manager1', { bu: ['ITS'], displayName: '測試主管二' }),
    mkUser(U.m1p, 'manager1', { bu: ['ERP'], supervisor: U.m1, pureSupervisor: true, displayName: '掛名主管' }),
    mkUser(U.m1off, 'manager1', { bu: ['ERP'], active: false, supervisor: U.m1, displayName: '離職主管' }),
    mkUser(U.m2, 'manager2', { bu: ['ERP'], supervisor: U.m1, displayName: '測試二級主管' }),
    mkUser(U.u1, 'user', { bu: ['ERP'], supervisor: U.m2, displayName: '測試業務甲', nickname: '甲' }),
    mkUser(U.u2, 'user', { bu: ['ERP'], supervisor: U.m1, displayName: '測試業務乙' }),
    mkUser(U.u3, 'user', { bu: ['ITS'], supervisor: U.m1b, displayName: '測試業務丙' }),
    mkUser(U.u4, 'user', { bu: ['ERP'], supervisor: U.m1p, displayName: '測試業務丁' }),
    mkUser(U.u5, 'user', { bu: ['ERP'], supervisor: U.m1off, displayName: '測試業務戊' }),
    mkUser(U.u6, 'user', { bu: ['CRM'], displayName: '測試業務己（無主管無秘書）' }),
    mkUser(U.uoff, 'user', { bu: ['ERP'], active: false, supervisor: U.m1, displayName: '離職業務' }),
    mkUser(U.mkt, 'marketing', { bu: ['ERP'], displayName: '測試行銷' }), mkUser(U.mkt2, 'marketing', { displayName: '測試行銷二' }),
    mkUser(U.mktoff, 'marketing', { displayName: '離職行銷', active: false }),
    mkUser(U.sec, 'secretary', { bu: ['ERP'], displayName: '測試秘書' }), mkUser(U.sec2, 'secretary', { bu: ['ITS'], displayName: '測試秘書二' }),
    mkUser(U.sec3, 'secretary', { bu: ['ERP'], displayName: '測試秘書三' }), mkUser(U.secoff, 'secretary', { bu: ['ERP'], active: false, displayName: '離職秘書' }),
    mkUser(U.gs, 'groupsales', { bu: ['ERP'], viewOwnerScope: U.u1, viewGroupId: 'g1', displayName: '測試集團業務' }),
    mkUser(U.tp, 'tecopm', { viewOwnerScope: U.u1, viewGroupId: 'g1', displayName: '測試集團PM' }),
    mkUser(U.cs, 'consult_manager_south', { displayName: '測試顧問主管' }),
  ];
}
const usersObj = () => seedUsers();
const CAT = () => ({ bus: { ERP: [{ group: '假商品群組', items: [
  { name: 'Svc MA', termKind: 'ma', startsTask: true }, { name: 'Svc Lic', termKind: 'license', startsTask: false },
  { name: 'Svc LicTask', termKind: 'license', startsTask: true }, { name: 'Plain', },
] }] } });
const MAP = 'Svc MA 一年', LICP = 'Svc Lic 永久', PLAINP = 'Plain 專案';
const mkOpp = (id, owner, product, stage, extra) => Object.assign({ id, owner, bu: 'ERP', company: '假公司' + id, contactName: '假聯絡人', product, amount: AMT, description: NOTE, expectedDate: '2026-12-31', stage, createdAt: nowIso }, extra || {});
/** 納管項（to＝今天 + dl 天；from＝to − 364 天） */
function mkItem(id, owner, dl, extra) {
  const to = plusDays(todayTp(), dl);
  const o = mkOpp('o_' + id, owner, MAP, 'Won');
  return Object.assign({ id, oppId: o.id, quoteId: '', owner, company: '假公司' + id, product: MAP, termKind: 'ma', from: plusDays(to, -364), to, contractNo: '', contractId: '', status: 'tracking', source: 'request', createdAt: agoIso(5), updatedAt: agoIso(5), history: [] }, extra || {});
}
const mkQuote = (id, oppId, owner, company, items, state, extra) => Object.assign({ id, oppId, owner, quoteNo: 'QT-' + id, company, items: items || [], status: state === 'approved' ? 'approved' : 'draft',
  approval: state ? { state, cur: state === 'approved' ? 1 : 0, steps: state === 'approved' ? [{ status: 'approved', at: nowIso, label: '核決' }] : [], hash: 'x' } : null, createdAt: nowIso, updatedAt: nowIso }, extra || {});
const qi = (lid, kind, from, to) => ({ lid, desc: '品項' + lid, qty: 1, price: 100, termKind: kind, termFrom: from, termTo: to });

function seedData(opts) {
  opts = opts || {};
  const d = {
    contacts: [], companies: [], visits: [], receivables: [], callins: [], keyAccounts: [], industries: [], notifications: [],
    productCatalog: CAT(),
    groups: [{ id: 'g1', name: '假集團', memberCompanies: ['假公司gA'] }],
    opportunities: [], quotations: [], contracts: [], maItems: [], dealTasks: [],
  };
  if (opts.data) opts.data(d);
  return d;
}
function seedFiles(opts) {
  opts = opts || {};
  const DIR = opts.dir || ROOT;
  fs.writeFileSync(path.join(DIR, 'auth.json'), JSON.stringify({ users: opts.users ? opts.users(seedUsers()) : seedUsers() }, null, 2));
  fs.writeFileSync(path.join(DIR, 'data.json'), JSON.stringify(seedData(opts), null, 2));
}
const readData = (dir) => JSON.parse(fs.readFileSync(path.join(dir || ROOT, 'data.json'), 'utf8'));
const writeData = (fn, dir) => { const f = path.join(dir || ROOT, 'data.json'); const d = JSON.parse(fs.readFileSync(f, 'utf8')); fn(d); fs.writeFileSync(f, JSON.stringify(d, null, 2)); };
function cleanupFiles(dir) { FILES.forEach((f) => { try { fs.unlinkSync(dir ? path.join(dir, path.basename(f)) : f); } catch { /* 不存在 */ } }); }

// ───────────────── 伺服器（預載腳本可在載入時改寫指定檔案＝變異測試／失敗注入）─────────────────
const PRELOAD = path.join(os.tmpdir(), 'marem_preload_' + process.pid + '.js');
fs.writeFileSync(PRELOAD, [
  "const Module=require('module'),path=require('path');const orig=Module.prototype._compile;",
  "Module.prototype._compile=function(content,filename){const raw=process.env.MAREM_MUT_SPEC;",
  "if(raw){const specs=JSON.parse(raw);for(const m of specs){if(path.resolve(filename)===path.resolve(process.env.MAREM_ROOT,m.file)){for(const e of m.edits){if(!content.includes(e.find))throw new Error('MUTATION ANCHOR MISSING: '+e.find.slice(0,60));content=content.replace(e.find,()=>e.rep);}}}}",
  "return orig.call(this,content,filename);};",
].join('\n') + '\n');
let child = null, childPid = null;
const CRON_SECRET = 'cron-' + crypto.randomBytes(8).toString('hex');
async function startServer(mut, extraEnv, root, port) {
  root = root || ROOT; port = port || PORT;
  const env = Object.assign({}, process.env, { PORT: String(port), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: `http://localhost:${port}`, CRON_SECRET });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'QUOTE_FLOW_ENGINE', 'QUOTE_FLOW_DESIGNER', 'FEATURE_GUARD_MODE', 'MAREM_MUT_SPEC', 'MA_TRACKING', 'APPROVAL_CONFIG_DUAL_REVIEW', 'QUOTE_SPLIT_GUARD']) delete env[k];
  Object.assign(env, extraEnv || {});
  if (mut && mut.length) { env.MAREM_MUT_SPEC = JSON.stringify(mut); env.MAREM_ROOT = ROOT; }
  await assertPortFree(port);
  child = spawn(process.execPath, ['-r', PRELOAD, path.join(root, 'server.js')], { cwd: root, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { out += d; });
  child.getOut = () => out;
  for (let i = 0; i < 120; i++) {
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
let jar = {};
let MUT_CUR = null;
async function freshServer(seedOpts, env) { await stopServer(); cleanupFiles(); seedFiles(seedOpts); await startServer(MUT_CUR, Object.assign({ MA_TRACKING: '1' }, env || {})); jar = {}; }
/** 不改資料檔、重啟（登入限流是每行程 20 次／15 分鐘） */
async function restartServer(env) { await stopServer(); await startServer(MUT_CUR, Object.assign({ MA_TRACKING: '1' }, env || {})); jar = {}; }

// ───────────────── HTTP ─────────────────
async function httpAt(base, jr, who, method, url, body, hdr) {
  const h = Object.assign({}, hdr || {});
  if (body !== undefined) h['Content-Type'] = 'application/json';
  if (who && jr[who]) h.Cookie = jr[who];
  const res = await fetch(base + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: res.status, json, text: buf.toString('utf8'), headers: res.headers };
}
const http = (who, method, url, body, hdr) => httpAt(BASE, jar, who, method, url, body, hdr);
async function loginAt(base, jr, who) {
  const res = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status === 200) jr[who] = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
  return res.status;
}
const login = (who) => loginAt(BASE, jar, who);
const ALLU = Object.values(U);
async function loginMany(list) { const bad = []; for (const w of list) { const s = await login(w); if (s !== 200) bad.push(w + ':' + s); } return bad; }
async function auditLines(action) {
  await sleep(300);
  const r = await http(U.adm, 'GET', '/api/admin/logs');
  return (Array.isArray(r.json) ? r.json : []).filter((l) => l.action === action);
}
const cron = () => http(null, 'GET', '/api/cron/mail-outbox', undefined, { Authorization: 'Bearer ' + CRON_SECRET });
const notifs = async (who, type) => { const j = (await http(who, 'GET', '/api/notifications')).json; return (Array.isArray(j) ? j : []).filter((n) => !type || n.type === type); };
const poll = async (who) => (await http(who, 'GET', '/api/poll-bundle')).json || {};
const norm = (x) => String(x).replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/g, '<uuid>').replace(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/g, '<ts>').replace(/\d{4}-\d{2}-\d{2}/g, '<date>').replace(/"(itemsSig|costLinesSig|contentHash|hash)":"[0-9a-f]{8,}"/g, '"$1":"<sig>"');   // 簽章是由隨機 lid 算出來的，每次執行不同
const FORBIDDEN_TOKENS = [AMT, NOTE, CNO];
const clean = (res) => !FORBIDDEN_TOKENS.some((t) => (res.text || String(res)).includes(t));


// ═════════════ 共用：純函式測試用的小工具 ═════════════
const END = '2026-12-31';
/** 距迄日 END 還有 dl 天的「現在」（台北中午，避開日界） */
const atDl = (dl) => Date.parse(END + 'T04:00:00Z') - dl * DAY;
const RULE = (o) => Object.assign({ id: 'r1', name: '規則', enabled: true, whenDays: 30, repeatEveryDays: 0, notify: ['owner'], condition: 'always', appliesTo: 'all' }, o || {});
const ITEM = (o) => Object.assign({ id: 'i1', oppId: 'o1', owner: U.u1, company: '假公司甲', product: MAP, termKind: 'ma', from: '2026-01-01', to: END, status: 'tracking', history: [] }, o || {});
const users23 = () => seedUsers();
const ids = (list) => list.map((x) => x.id);
const hitsOf = (pl) => pl.entries.map((e) => ({ item: e.item.id, rules: e.hits.map((h) => h.rule.id) }));
/** 獨立於 lib 的規則成立判斷（試算／對拍用） */
function oracleActive(rule, kind, dl, qs, status) {
  if (status !== 'tracking') return false;
  if (!rule.enabled) return false;
  if (rule.appliesTo !== 'all' && !rule.appliesTo.includes(kind)) return false;
  if (dl < -400) return false;
  if (!(dl <= rule.whenDays)) return false;
  if (rule.whenDays > 0 && dl < 0) return false;
  if (rule.condition === 'no_renewal_quote') return qs === 'none';
  if (rule.condition === 'quote_not_submitted') return qs === 'draft';
  return true;   // always／not_renewed（共同前提 tracking 已在上面檢查）
}

// ═════════════ 單元 ═════════════
function groupUnit() {
  const MR = REF.MR, RC = REF.RC, MA = REF.MA;
  console.log('[單元] lib/maReminders.js、lib/maRecipients.js');
  const V = (rules) => MR.validateConfig({ rules });
  const code = (rules) => { const r = V(rules); return r.ok ? 'OK' : r.code; };

  // ── 規則驗證 ──
  const dv = V(clone(MR.DEFAULT_RULES));
  ok(dv.ok && J(dv.value.rules) === J(MR.DEFAULT_RULES), 'V1 內建預設規則本身通過驗證且正規化後不變', dv.ok ? '' : dv.error);
  eq(MR.DEFAULT_RULES.length, 4, 'V1b 內建預設 4 條（D-90／D-45／D-30／到期後每 7 天）');
  eq(MR.DEFAULT_RULES.map((r) => [r.whenDays, r.repeatEveryDays, r.notify.join('+'), r.condition]), [[90, 0, 'owner+secretary', 'always'], [45, 0, 'manager1', 'no_renewal_quote'], [30, 0, 'owner+manager1+marketing', 'quote_not_submitted'], [-1, 7, 'owner+manager1', 'not_renewed']], 'V1c 預設規則內容符合規格');
  eq(code([]), 'OK', 'V2 空陣列可（管理員可關掉全部提醒）');
  eq([MR.validateConfig(null).code, MR.validateConfig({}).code, MR.validateConfig({ rules: 'x' }).code, MR.validateConfig({ rules: {} }).code, MR.validateConfig('x').code], ['BAD_REMINDER_RULES', 'BAD_REMINDER_RULES', 'BAD_REMINDER_RULES', 'BAD_REMINDER_RULES', 'BAD_REMINDER_RULES'], 'V3 非陣列／缺 rules → BAD_REMINDER_RULES');
  eq(code([null]), 'BAD_REMINDER_RULES', 'V3b 規則本身是 null');
  eq(code([[]]), 'BAD_REMINDER_RULES', 'V3c 規則是陣列');
  const many = (n) => Array.from({ length: n }, (_, i) => RULE({ id: 'r' + i, whenDays: i }));
  eq([code(many(20)), code(many(21))], ['OK', 'TOO_MANY_RULES'], 'V4 最多 20 條（21 條被拒）');
  eq([code([RULE({ id: 'a b' })]), code([RULE({ id: 'x'.repeat(41) })]), code([RULE({ id: 5 })]), code([RULE({ id: '中文' })]), code([RULE({ id: 'ok-1_A' })])], ['BAD_RULE_ID', 'BAD_RULE_ID', 'BAD_RULE_ID', 'BAD_RULE_ID', 'OK'], 'V5 代號：英數字底線連字號 1–40 字');
  eq(code([RULE({ id: 'dup' }), RULE({ id: 'dup', whenDays: 5 })]), 'DUPLICATE_RULE_ID', 'V5b 代號重複');
  const gen = V([RULE({ id: undefined }), RULE({ id: '', whenDays: 5 })]);
  ok(gen.ok && /^r_[0-9a-f]{10}$/.test(gen.value.rules[0].id) && gen.value.rules[0].id !== gen.value.rules[1].id, 'V5c 缺代號自動產生且不重複');
  eq([code([RULE({ name: '' })]), code([RULE({ name: '   ' })]), code([RULE({ name: 'x'.repeat(41) })]), code([RULE({ name: '字'.repeat(41) })]), code([RULE({ name: 'a\u0007b' })]), code([RULE({ name: '<b>x</b>' })]), code([RULE({ name: undefined })]), code([RULE({ name: 5 })]), code([RULE({ name: '字'.repeat(40) })])],
    ['BAD_RULE_NAME', 'BAD_RULE_NAME', 'BAD_RULE_NAME', 'BAD_RULE_NAME', 'BAD_RULE_NAME', 'BAD_RULE_NAME', 'BAD_RULE_NAME', 'BAD_RULE_NAME', 'OK'], 'V6 名稱：必填、≤40 字（以字元計）、不含控制字元與 < >');
  eq([code([RULE({ enabled: 'yes' })]), code([RULE({ enabled: 1 })]), code([RULE({ enabled: false })])], ['BAD_RULE_ENABLED', 'BAD_RULE_ENABLED', 'OK'], 'V7 enabled 必須是布林');
  ok(V([RULE({ enabled: undefined })]).value.rules[0].enabled === true, 'V7b 沒給 enabled → true');
  const W = (w) => code([RULE({ whenDays: w })]);
  eq([W(365), W(-365), W(0), W(1)], ['OK', 'OK', 'OK', 'OK'], 'V8 天數邊界 ±365、0');
  eq([W(366), W(-366), W(1.5), W('abc'), W(null), W(undefined), W(NaN), W(Infinity), W('1e2'), W(''), W([]), W({})], Array(12).fill('BAD_RULE_WHEN'), 'V8b 天數不合法 → BAD_RULE_WHEN');
  ok(V([RULE({ whenDays: '12' })]).value.rules[0].whenDays === 12 && V([RULE({ whenDays: '-3' })]).value.rules[0].whenDays === -3, 'V8c 數字字串會正規化為整數');
  ok(Object.is(V([RULE({ whenDays: -0 })]).value.rules[0].whenDays, 0), 'V8d -0 正規化為 0');
  const Rp = (rep, when) => code([RULE({ whenDays: when === undefined ? -1 : when, repeatEveryDays: rep })]);
  eq([Rp(0), Rp(1), Rp(90), Rp(undefined), Rp(null), Rp('7')], ['OK', 'OK', 'OK', 'OK', 'OK', 'OK'], 'V9 重複間隔 0／1–90（含未給）');
  eq([Rp(91), Rp(-1), Rp(1.5), Rp('x'), Rp(7, 5), Rp(7, 1), Rp(7, 0)], ['BAD_RULE_REPEAT', 'BAD_RULE_REPEAT', 'BAD_RULE_REPEAT', 'BAD_RULE_REPEAT', 'BAD_RULE_REPEAT', 'BAD_RULE_REPEAT', 'OK'], 'V9b 超出範圍或「到期前的規則」設重複 → BAD_RULE_REPEAT；到期當天可');
  const Nf = (n) => code([RULE({ notify: n })]);
  eq([Nf([]), Nf('owner'), Nf(['owner', 'x']), Nf([5]), Nf(undefined), Nf(['OWNER']), Nf(['owner', 'secretary', 'manager1', 'manager2', 'marketing'])], ['BAD_RULE_NOTIFY', 'BAD_RULE_NOTIFY', 'BAD_RULE_NOTIFY', 'BAD_RULE_NOTIFY', 'BAD_RULE_NOTIFY', 'BAD_RULE_NOTIFY', 'OK'], 'V10 通知對象：非空白名單陣列');
  eq(V([RULE({ notify: ['marketing', 'owner', 'owner'] })]).value.rules[0].notify, ['owner', 'marketing'], 'V10b 通知對象去重並依固定順序');
  eq([code([RULE({ condition: 'x' })]), code([RULE({ condition: undefined })]), code([RULE({ condition: 'always' })]), code([RULE({ condition: 'no_renewal_quote' })]), code([RULE({ condition: 'quote_not_submitted' })]), code([RULE({ condition: 'not_renewed' })])], ['BAD_RULE_CONDITION', 'BAD_RULE_CONDITION', 'OK', 'OK', 'OK', 'OK'], 'V11 條件白名單');
  const Ap = (a) => V([RULE({ appliesTo: a })]);
  eq([Ap('all').value.rules[0].appliesTo, Ap(['license']).value.rules[0].appliesTo, Ap(['ma']).value.rules[0].appliesTo, Ap(['ma', 'license']).value.rules[0].appliesTo, Ap(undefined).value.rules[0].appliesTo, Ap(null).value.rules[0].appliesTo], ['all', ['license'], ['ma'], 'all', 'all', 'all'], 'V12 適用類型正規化（兩種都選＝all）');
  eq([Ap([]).code, Ap(['x']).code, Ap('ma').code, Ap(['ma', 'x']).code, Ap(5).code], Array(5).fill('BAD_RULE_APPLIES'), 'V12b 適用類型不合法');
  eq([code([RULE({ id: 'a' }), RULE({ id: 'b' })]), code([RULE({ id: 'a' }), RULE({ id: 'b', condition: 'no_renewal_quote' })]), code([RULE({ id: 'a' }), RULE({ id: 'b', appliesTo: ['ma'] })]), code([RULE({ id: 'a', appliesTo: ['ma'] }), RULE({ id: 'b', appliesTo: ['ma'] })]), code([RULE({ id: 'a' }), RULE({ id: 'b', whenDays: 29 })])],
    ['DUPLICATE_RULE_WHEN', 'OK', 'OK', 'DUPLICATE_RULE_WHEN', 'OK'], 'V13 同 condition＋appliesTo 下天數不得重複（不同條件／不同類型可）');
  const evil = V([Object.assign(RULE({ id: 'e1' }), { __proto__: { polluted: 1 }, extra: 'x', constructor: 'y' })]);
  ok(evil.ok && Object.keys(evil.value.rules[0]).sort().join() === 'appliesTo,condition,enabled,id,name,notify,repeatEveryDays,whenDays' && ({}).polluted === undefined, 'V14 多餘欄位被丟棄、不污染原型');
  ok(!V([RULE({ name: '正常' })]).value.rules[0].notify.includes('x') && Object.isFrozen(MR.DEFAULT_RULES), 'V14b 預設規則是凍結的（不可被改）');

  // ── 沒帶代號的規則沿用現行同條件的代號（避免去重鍵換新、提醒重發）──
  { const cur = [RULE({ id: 'keepA', whenDays: 30 }), RULE({ id: 'keepB', whenDays: 10, condition: 'no_renewal_quote' })];
    const raw = [{ name: 'x', whenDays: 30, notify: ['owner'], condition: 'always' }, { name: 'y', whenDays: 10, notify: ['owner'], condition: 'no_renewal_quote' }, { name: 'z', whenDays: 30, notify: ['owner'], condition: 'always', appliesTo: ['ma'] }, { id: 'given', name: 'w', whenDays: 5, notify: ['owner'], condition: 'always' }];
    const v = V(raw);
    const ad = MR.adoptIds(raw, v.value.rules, cur);
    eq(ad.map((r) => r.id), ['keepA', 'keepB', v.value.rules[2].id, 'given'], 'V15 沒帶代號：條件＋類型＋天數相同 → 沿用現行代號；不同（類型不同）或帶了代號 → 不動');
    const raw2 = [{ name: 'x', whenDays: 30, notify: ['owner'], condition: 'always' }, { name: 'x2', whenDays: 30, notify: ['marketing'], condition: 'always' }];
    const ad2 = MR.adoptIds(raw2, V(raw2).ok ? V(raw2).value.rules : [], cur);
    ok(V(raw2).ok === false, 'V15b （同 condition＋類型＋天數的兩條本來就不合法）');
    const raw3 = [{ id: 'keepA', name: 'x', whenDays: 40, notify: ['owner'], condition: 'always' }, { name: 'y', whenDays: 30, notify: ['owner'], condition: 'no_renewal_quote' }];
    const v3 = V(raw3);
    eq(MR.adoptIds(raw3, v3.value.rules, cur).map((r) => r.id)[0], 'keepA', 'V15c 明確帶代號者保持原樣，且該代號不會被別條規則搶走');
    ok(cur.length === 2 && cur[0].id === 'keepA', 'V15d 不修改輸入');
  }
  // ── getConfig 不信任儲存值 ──
  let c = MR.getConfig({});
  ok(c.isDefault && c.rules.length === 4 && c.updatedAt === null, 'G1 沒設定 → 預設');
  c.rules[0].whenDays = 1; c.rules[0].notify.push('marketing');
  ok(MR.DEFAULT_RULES[0].whenDays === 90 && MR.getConfig({}).rules[0].whenDays === 90 && MR.DEFAULT_RULES[0].notify.length === 2, 'G1b 回傳的是複本，改它不影響內建預設');
  c = MR.getConfig({ reminderRules: { rules: [RULE({ id: 'z1', whenDays: 10 })], updatedAt: '2026-01-01T00:00:00.000Z', updatedBy: 'x' } });
  ok(!c.isDefault && c.rules.length === 1 && c.rules[0].id === 'z1' && c.updatedBy === 'x', 'G2 儲存的有效設定被採用');
  ok(MR.getConfig({ reminderRules: { rules: [RULE({ id: 'z1' }), RULE({ id: 'z2', whenDays: 999 })] } }).isDefault, 'G3 其中一條不合法 → 整份退回預設（不信任儲存值）');
  ok(MR.getConfig({ reminderRules: { rules: 'x' } }).isDefault && MR.getConfig({ reminderRules: 'x' }).isDefault && MR.getConfig({ reminderRules: null }).isDefault && MR.getConfig(null).isDefault && MR.getConfig({ reminderRules: { rules: [] } }).rules.length === 0, 'G4 非陣列／垃圾 → 預設；空陣列＝刻意全關');
  c = MR.getConfig({ reminderRules: { rules: [RULE({ extra: 1, __x: 2 })], updatedAt: 5, updatedBy: {} } });
  ok(c.updatedAt === null && c.updatedBy === null && !('extra' in c.rules[0]), 'G5 多餘欄位與錯型的 updatedAt／updatedBy 被丟棄');
  const df = MR.configDiff({ rules: [RULE({ id: 'a' }), RULE({ id: 'b', whenDays: 5 })] }, { rules: [RULE({ id: 'a', whenDays: 31 }), RULE({ id: 'c', whenDays: 7 })] });
  ok(df.length === 3 && /^變更規則/.test(df[0]) && /^新增規則/.test(df[1]) && /^移除規則/.test(df[2]) && MR.configDiff({ rules: [RULE()] }, { rules: [RULE()] }).length === 0, 'G6 configDiff：變更／新增／移除各一行，沒變不列');

  // ── 門檻語意 ──
  const WS = (when, dl, opt) => MR.windowState(RULE(Object.assign({ whenDays: when }, opt || {})), 'ma', dl);
  eq([WS(45, 46).reason, WS(45, 45).active, WS(45, 44).active, WS(45, 0).active, WS(45, -1).reason], ['NOT_REACHED', true, true, true, 'EXPIRED_WINDOW'], 'S1 到期前 45 天：46 未到、45 剛好、44 補發、0 仍有效、-1 到期前的規則不補發');
  eq([WS(0, 1).reason, WS(0, 0).active, WS(0, -5).active, WS(0, -400).active, WS(0, -401).reason], ['NOT_REACHED', true, true, true, 'TOO_OLD'], 'S2 到期當天：1 未到、0 剛好、過期持續有效直到 -400；-401 不再提醒');
  eq([WS(-3, -2).reason, WS(-3, -3).active, WS(-3, -4).active, WS(-3, -50).active], ['NOT_REACHED', true, true, true], 'S3 到期後 3 天：-2 未到、-3 剛好、之後持續有效');
  eq([WS(30, 10, { enabled: false }).reason, WS(30, 10, { appliesTo: ['license'] }).reason, WS(30, 10, { appliesTo: ['ma'] }).active, WS(30, 10, { appliesTo: 'all' }).active], ['DISABLED', 'APPLIES', true, true], 'S4 停用／適用類型不符');
  const PRD = (when, rep, dl) => MR.periodOf(RULE({ whenDays: when, repeatEveryDays: rep }), dl);
  eq([PRD(-1, 7, -1), PRD(-1, 7, -7), PRD(-1, 7, -8), PRD(-1, 7, -14), PRD(-1, 7, -15), PRD(0, 1, 0), PRD(0, 1, -1), PRD(0, 3, -2), PRD(0, 3, -3), PRD(30, 0, 5), PRD(-1, 0, -99)], [0, 0, 1, 1, 2, 0, 1, 0, 1, 0, 0], 'S5 週期＝floor((whenDays − dl) / repeat)；不重複者恆 0');
  eq([MR.logKey('r1', 'i1', 'u', '2026-12-31', 0), MR.keyParts('r1|i1|a|b|2026-12-31|3')], ['r1|i1|u|2026-12-31|0', { period: '3', endDate: '2026-12-31', ruleId: 'r1', itemId: 'i1', recipient: 'a|b' }], 'S6 去重鍵格式；帳號含 | 時仍能由右邊解析迄日與週期');

  // ── 條件與續約報價單 ──
  const q = (renewOf, state, extra) => Object.assign({ id: 'q' + Math.random(), renewOf, approval: state === undefined ? null : { state }, status: 'draft' }, extra || {});
  const stOf = (quotes, id) => MR.quoteStateOf(MR.renewalIndex({ quotations: quotes }), id || 'i1');
  eq([stOf([]), stOf([q('i1')]), stOf([q('i1', 'none')]), stOf([q('i1', 'returned')]), stOf([q('i1', 'pending')]), stOf([q('i1', 'approved')])], ['none', 'draft', 'draft', 'draft', 'sent', 'sent'], 'C1 報價單狀態：無／草稿（含無簽核、none、退回）／已送簽（簽核中、已核准）');
  eq([stOf([q('i1'), q('i1', 'pending')]), stOf([q('i1', 'pending'), q('i1')]), stOf([q('i2', 'approved')]), stOf([q(5, 'approved')]), stOf([q('', 'approved')]), stOf([q(undefined, 'approved')])], ['sent', 'sent', 'none', 'none', 'none', 'none'], 'C2 多張取最進展者；指向別的納管項／非字串 renewOf 不算');
  eq([stOf([q('i1', 'approved', { invalid: { at: 'x' } })]), stOf([q('i1', 'pending', { status: 'rejected' })]), stOf([q('i1', 'pending', { deleted: true })]), stOf([q('i1', 'approved', { invalid: { at: 'x' } }), q('i1')])], ['none', 'none', 'none', 'draft'], 'C3 作廢（失效）／舊制已拒絕／已刪除的報價單不算；其餘有效的照算');
  const CH = (cond, qs, status) => MR.conditionHolds(cond, ITEM({ status: status || 'tracking' }), qs);
  eq(['none', 'draft', 'sent'].map((s) => CH('always', s)), [true, true, true], 'C4 always 恆成立');
  eq(['none', 'draft', 'sent'].map((s) => CH('no_renewal_quote', s)), [true, false, false], 'C5 no_renewal_quote：只有「沒有任何有效續約報價單」');
  eq(['none', 'draft', 'sent'].map((s) => CH('quote_not_submitted', s)), [false, true, false], 'C6 quote_not_submitted：有續約報價單但全部是草稿');
  eq([CH('not_renewed', 'none'), CH('not_renewed', 'sent'), CH('not_renewed', 'none', 'renewed'), CH('not_renewed', 'none', 'not_renew'), CH('not_renewed', 'none', 'ended'), CH('not_renewed', 'none', 'void'), CH('bogus', 'none')], [true, true, false, false, false, false, false], 'C7 not_renewed：只有 tracking；未知條件一律 false');
  // 共同前提（status）對所有條件生效
  const planOne = (item, rules, qs, now) => MR.plan({ maItems: [item], quotations: qs || [] }, users23(), now === undefined ? atDl(10) : now, rules);
  for (const st of ['renewed', 'not_renew', 'ended', 'void']) eq(planOne(ITEM({ status: st }), [RULE({ whenDays: 30, condition: 'always' })]).entries.length, 0, `C8 狀態 ${st}：任何條件（含 always）都不提醒`);
  eq(planOne(ITEM({ to: '2026-02-30' }), [RULE()]).entries.length + planOne(ITEM({ to: '' }), [RULE()]).entries.length + planOne(ITEM({ termKind: 'x' }), [RULE()]).entries.length + planOne(null, [RULE()]).evaluated, 0, 'C9 假日期／缺迄日／未知類型不評估、不丟例外');

  // ── 台北日期（UTC 跨日）──
  const dlAt = (iso, to) => MA.daysLeft(to, iso);
  eq([dlAt('2026-10-24T15:30:00.000Z', '2026-10-25'), dlAt('2026-10-24T16:30:00.000Z', '2026-10-25'), dlAt('2026-10-24T23:59:59.000Z', '2026-10-25'), dlAt('2026-10-25T00:30:00.000Z', '2026-10-25')], [1, 0, 0, 0], 'D1 UTC 23:30（台北 23:30）剩 1 天；UTC 16:30（台北隔天 00:30）剩 0 天');
  const itx = ITEM({ to: '2026-10-25' });
  const rulesD = [RULE({ id: 'd1', whenDays: 1 })];
  const rUtc1 = MR.sweep({ maItems: [itx], notifications: [] }, users23(), { now: Date.parse('2026-10-24T15:30:00.000Z'), rules: rulesD });
  const rUtc2 = MR.sweep({ maItems: [itx], notifications: [] }, users23(), { now: Date.parse('2026-10-24T16:30:00.000Z'), rules: rulesD });
  eq([rUtc1.today, rUtc1.stats.fired, rUtc2.today, rUtc2.stats.fired], ['2026-10-24', 1, '2026-10-25', 1], 'D2 sweep 的「今天」是台北日期；d=1 的規則在兩個時間點都達標（剩 1 天／剩 0 天補發）');
  eq([MR.sweep({ maItems: [ITEM({ to: '2026-10-26' })], notifications: [] }, users23(), { now: Date.parse('2026-10-24T16:30:00.000Z'), rules: rulesD }).stats.fired, MR.sweep({ maItems: [ITEM({ to: '2026-10-26' })], notifications: [] }, users23(), { now: Date.parse('2026-10-24T15:30:00.000Z'), rules: rulesD }).stats.fired], [1, 0], 'D3 UTC 16:30＝台北 10/25 → 距 10/26 剩 1 天達標；UTC 15:30＝台北 10/24 → 剩 2 天未達標（跨日邊界）');
  const tzOut = ['UTC', 'Pacific/Honolulu', 'Asia/Taipei', 'Pacific/Kiritimati'].map((tz) => spawnSync(process.execPath, ['-e', "const R=require(process.argv[1]);const M=require(process.argv[2]);const d={maItems:[{id:'i',owner:'u',company:'c',product:'p',termKind:'ma',from:'2026-01-01',to:'2026-12-31',status:'tracking'}],notifications:[]};const r=R.sweep(d,[{username:'u',role:'user',bu:['ERP']},{username:'a',role:'admin'}],{now:Date.parse('2026-10-24T16:00:00.000Z'),rules:[{id:'x',name:'n',enabled:true,whenDays:70,repeatEveryDays:0,notify:['owner'],condition:'always',appliesTo:'all'}]});console.log(JSON.stringify([r.today,r.stats,M.taipeiDate('2026-10-24T16:00:00.000Z')]))", path.join(ROOT, 'lib', 'maReminders.js'), path.join(ROOT, 'lib', 'maTracking.js')], { env: Object.assign({}, process.env, { TZ: tz }), encoding: 'utf8' }).stdout.trim());
  ok(new Set(tzOut).size === 1 && /2026-10-25/.test(tzOut[0]), 'D4 換四種機器時區，掃描結果完全相同', tzOut.join(' | '));

  // ── 收件人解析 ──
  const UU = users23();
  const RB = (item, users) => RC.resolveByRole(item, users || UU);
  let r = RB(ITEM({ owner: U.u1 }));
  eq(r.byRole, { owner: [U.u1], secretary: [U.sec, U.sec3], manager1: [U.m1], manager2: [U.m2], marketing: [U.mkt, U.mkt2] }, 'R1 u1（ERP）：owner、多位在職同 BU 秘書（不含離職與別 BU）、一級主管 m1、直屬二級 m2、在職行銷全部');
  eq(r.issues, [], 'R1b 一切正常沒有問題回報');
  r = RB(ITEM({ owner: U.u3, bu: undefined }));
  eq([r.byRole.secretary, r.byRole.manager1, r.byRole.manager2], [[U.sec2], [U.m1b], []], 'R2 跨 BU：u3（ITS）只通知 ITS 秘書與自己線的一級主管；直屬不是 manager2 → 沒有 manager2');
  r = RB(ITEM({ owner: U.u2 }));
  eq([r.byRole.manager1, r.byRole.manager2], [[U.m1], []], 'R3 u2 直屬就是一級主管 → manager1＝直屬，manager2 沒有（不退管理員）');
  r = RB(ITEM({ owner: U.u4 }));
  eq(r.byRole.manager1, [U.m1], 'R4 u4 的直屬是掛名主管（pureSupervisor）→ 跳過往上找到 m1');
  r = RB(ITEM({ owner: U.u5 }));
  eq(r.byRole.manager1, [U.m1], 'R5 u5 的直屬主管已離職 → 跳過往上找到 m1');
  r = RB(ITEM({ owner: U.u6 }));
  eq([r.byRole.manager1, r.byRole.secretary, r.byRole.owner, r.issues.map((i) => i.role + ':' + i.code).sort()], [[U.adm, U.adm2], [U.adm, U.adm2], [U.u6], ['manager1:NO_MANAGER1', 'secretary:NO_SECRETARY']], 'R6 u6（CRM、無主管、無秘書）：一級主管與秘書都退到管理員並回報問題');
  r = RB(ITEM({ owner: U.uoff }));
  eq([r.byRole.owner, r.issues.map((i) => i.code)], [[U.m1], ['OWNER_FALLBACK']], 'R7 負責業務已離職 → owner 改取其一級主管');
  const noMgr = UU.map((u) => (u.username === U.uoff ? Object.assign({}, u, { supervisor: undefined }) : u));
  r = RB(ITEM({ owner: U.uoff }), noMgr);
  eq([r.byRole.owner, r.issues.map((i) => i.code).sort()], [[U.adm, U.adm2], ['NO_MANAGER1', 'NO_OWNER_RECIPIENT', 'NO_SECRETARY'].filter((c) => c !== 'NO_SECRETARY')], 'R8 離職且沒有一級主管 → owner 退到管理員');
  r = RB(ITEM({ owner: '_s_ghost' }));
  eq([r.byRole.owner, r.byRole.manager1], [[U.adm, U.adm2], [U.adm, U.adm2]], 'R9 帳號不存在 → 退到管理員');
  r = RB(ITEM({ owner: U.u1, bu: 'ITS' }));
  eq(r.byRole.secretary, [U.adm, U.adm2], 'R10 納管項 BU（ITS）與負責業務 BU（ERP）不同：ITS 秘書的可視範圍不含此業務 → 不通知（與秘書可視範圍一致），退管理員');
  r = RB(ITEM({ owner: U.u1, bu: 'ERP' }));
  eq(r.byRole.secretary, [U.sec, U.sec3], 'R10b 納管項 BU 明寫 ERP → ERP 秘書');
  r = RB(ITEM({ owner: U.u1, bu: 'XYZ' }));
  eq(r.byRole.secretary, [U.sec, U.sec3], 'R10c 非法 BU 值 → 用負責業務的第一個 BU');
  const mut = (fn) => UU.map((u) => fn(Object.assign({}, u)));
  r = RB(ITEM({ owner: U.u1 }), mut((u) => (u.username === U.m2 ? Object.assign(u, { active: false }) : u)));
  eq(r.byRole.manager2, [], 'R11 直屬二級主管停用 → 沒有 manager2');
  r = RB(ITEM({ owner: U.u1 }), mut((u) => (u.username === U.m2 ? Object.assign(u, { pureSupervisor: true }) : u)));
  eq(r.byRole.manager2, [], 'R11b 直屬二級主管是掛名 → 沒有 manager2');
  r = RB(ITEM({ owner: U.u1 }), mut((u) => (u.username === U.mkt ? Object.assign(u, { pureSupervisor: true }) : u.username === U.mkt2 ? Object.assign(u, { role: 'user' }) : u)));
  eq(r.byRole.marketing, [], 'R12 行銷：掛名者與非行銷角色不列入');
  r = RB(ITEM({ owner: U.u1 }), mut((u) => (u.role === 'secretary' ? Object.assign(u, { pureSupervisor: true }) : u)));
  eq(r.byRole.secretary, [U.adm, U.adm2], 'R13 所有秘書都是掛名 → 退管理員');
  r = RB(ITEM({ owner: U.u6 }), mut((u) => (u.username === U.adm2 ? Object.assign(u, { active: false }) : u.username === U.adm ? Object.assign(u, { pureSupervisor: true }) : u)));
  eq([r.byRole.manager1, r.byRole.secretary], [[], []], 'R14 管理員一個停用、一個掛名 → 沒有任何收件人（呼叫端稽核）');
  const cyc = [{ username: 'a', role: 'user', supervisor: 'b', bu: ['ERP'] }, { username: 'b', role: 'manager2', supervisor: 'a' }, { username: 'adm', role: 'admin' }];
  r = RC.resolveByRole({ id: 'x', owner: 'a', bu: 'ERP' }, cyc);
  eq(r.byRole.manager1, ['adm'], 'R15 主管鏈循環不會無限迴圈（退管理員）');
  r = RC.resolveByRole({ id: 'x', owner: U.u1 }, Object.fromEntries(UU.map((u) => [u.username, u])));
  eq(r.byRole.owner, [U.u1], 'R16 users 可以是物件對照表');
  const fn = RC.resolveForNotify(ITEM({ owner: U.u6 }), UU, ['secretary', 'manager1', 'owner']);
  eq(fn.recipients, [{ username: U.adm, roles: ['secretary', 'manager1'] }, { username: U.adm2, roles: ['secretary', 'manager1'] }, { username: U.u6, roles: ['owner'] }], 'R17 同一人同時是多個角色只出現一次（roles 合併）');
  eq(RC.resolveForNotify(ITEM({ owner: U.u1 }), UU, ['marketing', 'manager2']).issues, [], 'R17b issues 只列這組 notify 用到的角色');
  eq(RC.resolveForNotify(ITEM({ owner: U.u6 }), UU, ['owner']).issues, [], 'R17c 沒用到秘書／一級主管的規則不帶這些問題');
  eq(RC.resolveForNotify(ITEM({ owner: U.u1 }), UU, ['bogus', 'owner']).recipients.map((x) => x.username), [U.u1], 'R17d 不認得的角色忽略');

  // ── 掃描單元 ──
  const T = (dl) => atDl(dl);
  const mkData = (items, extra) => Object.assign({ maItems: items, quotations: [], notifications: [] }, extra || {});
  const one = [RULE({ id: 'a', whenDays: 30, notify: ['owner', 'manager1'] })];
  let d = mkData([ITEM()]);
  let s = MR.sweep(d, UU, { now: T(30), rules: one });
  eq([s.stats.fired, s.stats.recipients, s.created.length, Object.keys(d.reminderLog.entries).length], [2, 2, 2, 2], 'W1 剛好達標：owner＋manager1 各一則通知、2 筆去重標記');
  eq([s.created.map((n) => n.to).sort(), s.created[0].type, s.created[0].read, s.created[0].refId, s.created[0].maItemIds], [[U.m1, U.u1], 'ma_reminder', false, '2026-12-01', ['i1']], 'W1b 通知欄位：type／未讀／refId＝台北日期／maItemIds');
  ok(clean({ text: J(s.created) }) && !/金額|NT\$|\d{6,}/.test(s.created.map((n) => n.title + n.body).join('')), 'W1c 通知內容沒有金額');
  ok(/假公司甲/.test(s.created[0].body) && /Svc MA/.test(s.created[0].body) && /2026-12-31/.test(s.created[0].body), 'W1d 通知內容有公司名、品項、日期');
  s = MR.sweep(d, UU, { now: T(30), rules: one });
  eq([s.stats.fired, s.stats.skippedDup, s.created.length, d.notifications.length], [0, 2, 0, 2], 'W2 第二次掃描：已發過不重發（skippedDup）');
  s = MR.sweep(d, UU, { now: T(29), rules: one });
  eq([s.stats.fired, d.notifications.length], [0, 2], 'W2b 隔天再掃：同一事件不重複');
  d = mkData([ITEM()]);
  s = MR.sweep(d, UU, { now: T(25), rules: one });
  eq([s.stats.fired, s.created.length], [2, 2], 'W3 漏跑 5 天（剩 25 天才掃）→ 補發');
  d = mkData([ITEM()]);
  s = MR.sweep(d, UU, { now: T(31), rules: one });
  eq([s.stats.fired, s.stats.evaluated, d.notifications.length], [0, 1, 0], 'W3b 還沒到門檻（剩 31 天）不發；評估 1 項');
  d = mkData([ITEM()]);
  const before = clone(d);
  s = MR.sweep(d, UU, { now: T(30), rules: one, dryRun: true });
  eq([s.stats.fired, s.stats.recipients, s.created.length, s.changed, J(d) === J(before)], [2, 2, 2, false, true], 'W4 dryRun：算得出統計，但資料一個位元都不動');
  d = mkData([ITEM()]);
  MR.sweep(d, UU, { now: T(30), rules: one });
  d.maItems[0].to = '2027-01-31'; d.maItems[0].from = '2026-02-01';
  s = MR.sweep(d, UU, { now: Date.parse('2027-01-01T04:00:00Z'), rules: one });
  eq(s.stats.fired, 2, 'W5 修改迄日後（去重鍵含迄日）重新達標可再提醒');
  // 條件 × 報價單
  d = mkData([ITEM()]);
  s = MR.sweep(d, UU, { now: T(40), rules: [RULE({ id: 'n', whenDays: 45, condition: 'no_renewal_quote', notify: ['owner'] })] });
  eq(s.stats.fired, 1, 'W6 沒有續約報價單 → 發');
  d = mkData([ITEM()], { quotations: [q('i1', 'none')] });
  s = MR.sweep(d, UU, { now: T(40), rules: [RULE({ id: 'n', whenDays: 45, condition: 'no_renewal_quote', notify: ['owner'] })] });
  eq(s.stats.fired, 0, 'W6b 已有續約報價單（即使草稿）→ 不發');
  d = mkData([ITEM()], { quotations: [q('i1', 'none')] });
  s = MR.sweep(d, UU, { now: T(25), rules: [RULE({ id: 'n', whenDays: 30, condition: 'quote_not_submitted', notify: ['owner'] })] });
  eq(s.stats.fired, 1, 'W6c 草稿未送簽 → 發');
  d = mkData([ITEM()], { quotations: [q('i1', 'pending')] });
  s = MR.sweep(d, UU, { now: T(25), rules: [RULE({ id: 'n', whenDays: 30, condition: 'quote_not_submitted', notify: ['owner'] })] });
  eq(s.stats.fired, 0, 'W6d 已送簽 → 不發');
  // 按人彙總
  d = mkData([ITEM({ id: 'a1', company: '甲1' }), ITEM({ id: 'a2', company: '甲2' }), ITEM({ id: 'a3', company: '甲3' })]);
  s = MR.sweep(d, UU, { now: T(30), rules: [RULE({ id: 'x', whenDays: 30, notify: ['owner'] }), RULE({ id: 'y', whenDays: 45, notify: ['owner'] })] });
  eq([s.created.length, s.stats.fired, s.created[0].maItemIds.sort(), /你有 3 項/.test(s.created[0].title)], [1, 6, ['a1', 'a2', 'a3'], true], 'W7 同一人多個納管項、多條規則 → 只發一則彙總「你有 3 項」，6 筆去重標記');
  d = mkData(Array.from({ length: 8 }, (_, i) => ITEM({ id: 'm' + i, company: '客戶' + i })));
  s = MR.sweep(d, UU, { now: T(30), rules: [RULE({ notify: ['owner'] })] });
  ok(/…等共 8 項/.test(s.created[0].body) && (s.created[0].body.match(/客戶/g) || []).length === 5, 'W7b 內文最多列 5 項，其餘寫「等共 N 項」');
  // 無收件人
  d = mkData([ITEM({ owner: 'ghost' })]);
  s = MR.sweep(d, [{ username: 'zzz', role: 'user' }], { now: T(30), rules: one });
  ok(s.stats.noRecipient === 1 && s.stats.fired === 0 && s.audits.some((a) => a.action === 'MA_REMINDER_NO_RECIPIENT') && Object.keys(d.reminderLog.entries).length === 0, 'W8 連管理員都沒有 → noRecipient 計數＋稽核，不寫去重標記');
  d = mkData([ITEM({ owner: U.u6 })]);
  s = MR.sweep(d, UU, { now: T(30), rules: [RULE({ id: 'sm', whenDays: 30, notify: ['secretary', 'manager1'] })] });
  eq([s.created.map((n) => n.to).sort(), s.audits.filter((a) => a.action === 'MA_REMINDER_NO_RECIPIENT').length], [[U.adm, U.adm2], 2], 'W9 沒有秘書／一級主管 → 通知管理員，並各記一筆 MA_REMINDER_NO_RECIPIENT 稽核');
  s = MR.sweep(d, UU, { now: T(29), rules: [RULE({ id: 'sm', whenDays: 30, notify: ['secretary', 'manager1'] })] });
  eq(s.audits.length, 0, 'W9b 第二次掃描不重複記稽核');
  // 上限
  d = mkData(Array.from({ length: 250 }, (_, i) => ITEM({ id: 'c' + i, owner: 'own' + String(i).padStart(3, '0') })));
  const bigUsers = [{ username: U.adm, role: 'admin' }].concat(Array.from({ length: 250 }, (_, i) => ({ username: 'own' + String(i).padStart(3, '0'), role: 'user', bu: ['ERP'] })));
  s = MR.sweep(d, bigUsers, { now: T(30), rules: [RULE({ notify: ['owner'] })] });
  eq([s.created.length, s.stats.capped, s.stats.fired, d.notifications.length, Object.keys(d.reminderLog.entries).length], [200, 50, 200, 200, 200], 'W10 單次上限 200 則：超過的 50 人不發、不寫標記（capped）');
  s = MR.sweep(d, bigUsers, { now: T(30), rules: [RULE({ notify: ['owner'] })] });
  eq([s.created.length, s.stats.capped, s.stats.skippedDup, d.notifications.length], [50, 0, 200, 250], 'W10b 下一次掃描補發剩下的 50 人，已發的不重複');
  d = mkData(Array.from({ length: 10 }, (_, i) => ITEM({ id: 'k' + i, owner: 'own00' + i })));
  s = MR.sweep(d, bigUsers, { now: T(30), rules: [RULE({ notify: ['owner'] })], maxNotifications: 3 });
  eq([s.created.length, s.stats.capped], [3, 7], 'W10c maxNotifications 可注入');
  // 上限 500 與 trim
  d = mkData([ITEM()], { notifications: Array.from({ length: 500 }, (_, i) => ({ id: 'n' + i, to: 'x', type: 'other', title: 't', read: i % 2 === 0, createdAt: nowIso })) });
  s = MR.sweep(d, UU, { now: T(30), rules: one });
  ok(d.notifications.length === 500 && d.notifications.slice(0, 2).every((n) => n.type === 'ma_reminder') && d.notifications.filter((n) => !n.read).length >= 250, 'W11 通知全站上限 500：新的在前，優先丟「已讀且最舊」的，未讀的盡量保留');
  const tr = MR.trimNotifications(Array.from({ length: 12 }, (_, i) => ({ id: i, read: i === 11 })), 10);
  eq([tr.length, tr.some((n) => n.id === 11), tr.map((n) => n.id).slice(0, 3)], [10, false, [0, 1, 2]], 'W11b trim：最舊的已讀先丟，再丟最舊的');
  eq(MR.trimNotifications(Array.from({ length: 5 }, (_, i) => ({ id: i })), 10).length, 5, 'W11c 未超過上限不動');
  // 去重紀錄清理
  const log = { entries: {} };
  log.entries[MR.logKey('r', 'i', 'u', '2025-11-26', 0)] = '2025-01-01T00:00:00.000Z';   // 距 2027-01-01 共 401 天 → 清
  log.entries[MR.logKey('r', 'i', 'u', '2025-11-27', 0)] = '2025-01-02T00:00:00.000Z';   // 400 天 → 留
  log.entries[MR.logKey('r', 'i', 'u', '2025-12-31', 0)] = '2025-01-03T00:00:00.000Z';   // 366 天 → 留
  MR.pruneLog(log, '2027-01-01', 0);
  eq(Object.keys(log.entries).map((k) => MR.keyParts(k).endDate), ['2025-11-27', '2025-12-31'], 'W12 清理：迄日過了 400 天以上（401 天）的清掉；400 天與 366 天的保留');
  const bigLog = { entries: {} };
  for (let i = 0; i < 20005; i++) bigLog.entries['r|i|u' + i + '|2026-12-31|0'] = new Date(1e12 + i * 1000).toISOString();
  const rm = MR.pruneLog(bigLog, '2026-10-01', 0);
  ok(Object.keys(bigLog.entries).length === 20000 && rm === 5 && !bigLog.entries['r|i|u0|2026-12-31|0'] && bigLog.entries['r|i|u20004|2026-12-31|0'], 'W12b 超過 20000 筆：丟最舊的 5 筆');
  MR.pruneLog(bigLog, '2026-10-01', 10);
  eq(Object.keys(bigLog.entries).length, 19990, 'W12c 預留 incoming 空間');
  // 自動解除
  const mkN = (to, ids, extra) => Object.assign({ id: 'n' + Math.random(), to, type: 'ma_reminder', title: 't', body: 'b', read: false, maItemIds: ids, createdAt: nowIso }, extra || {});
  d = mkData([ITEM({ id: 'a' }), ITEM({ id: 'b', status: 'not_renew' }), ITEM({ id: 'c' })], { quotations: [q('c', 'pending')], notifications: [mkN(U.u1, ['a']), mkN(U.u1, ['b']), mkN(U.u1, ['gone']), mkN(U.u1, ['c']), mkN(U.u1, ['a', 'b']), mkN(U.u2, ['a']), mkN(U.u1, undefined), { id: 'dt', to: U.u1, type: 'deal_task', read: false, maItemIds: ['b'] }] });
  const rules1 = [RULE({ id: 'q', whenDays: 30, notify: ['owner'], condition: 'no_renewal_quote' })];
  const mk = MR.reconcileNotifications(d, UU, T(20), rules1);
  eq([mk, d.notifications.map((n) => n.read)], [4, [false, true, true, true, false, true, false, false]], 'A1 自動解除：仍成立的不動（a）；不續約（b）、不存在（gone）、已送簽（c）、他人不是收件人（u2 的 a）標已讀；無 maItemIds 與其他 type 不動');
  d.notifications.forEach((n) => { n.read = false; });
  eq(MR.reconcileNotifications(d, UU, T(20), rules1, (n) => n.to === U.u2), 1, 'A2 filter：只處理符合的通知');
  eq(MR.reconcileNotifications(mkData([ITEM()], { notifications: [mkN(U.u1, ['i1'], { read: true })] }), UU, T(20), rules1), 0, 'A3 已讀的不重複處理');
  // 動態鈴鐺
  const plD = MR.plan(mkData([ITEM({ id: 'z1', to: plusDays(todayTp(), 10) }), ITEM({ id: 'z2', to: plusDays(todayTp(), -3), company: '逾期公司' }), ITEM({ id: 'z3', owner: U.u2, to: plusDays(todayTp(), 5) })]), UU, Date.now(), MR.DEFAULT_RULES);
  const dyn = MR.dynamicFor(plD, U.u1, { labelOf: (u) => u });
  eq([dyn.map((x) => x.itemId), dyn.map((x) => x.level), dyn[0].days < dyn[1].days || dyn.length < 2], [['z2', 'z1'], ['danger', 'warn'], true], 'B1 動態鈴鐺：依剩餘天數由少到多；逾期＝danger、30 天內＝warn');
  ok(/已逾期 3 天/.test(dyn[0].title) && /剩 10 天到期/.test(dyn[1].title) && dyn[0].type === 'ma_reminder' && !/業務：/.test(dyn[0].body), 'B1b 標題文字；自己是業務不重複寫業務');
  ok(clean({ text: J(dyn) }), 'B1c 動態鈴鐺沒有金額與合約編號');
  const dynSec = MR.dynamicFor(plD, U.sec, { secretary: true, labelOf: (u) => u });
  ok(dynSec.length === 2 && dynSec.every((x) => !('itemId' in x) && !('rule' in x) && !('roles' in x) && /^ma_[0-9a-f]{12}$/.test(x.id)), 'B2 秘書：不給納管項 id 與規則細節，id 為雜湊');
  eq(Object.keys(dynSec[0]).sort(), ['body', 'company', 'days', 'from', 'id', 'level', 'ownerLabel', 'product', 'termKind', 'title', 'to', 'type'], 'B2b 秘書欄位白名單');
  eq(MR.dynamicFor(plD, U.u1, { visible: (it) => it.id === 'z1' }).map((x) => x.itemId), ['z1'], 'B3 visible 過濾（與 ma-items 同範圍）');
  eq(MR.dynamicFor(plD, 'nobody', {}), [], 'B4 不是任何規則的收件人 → 空');
  // 補掃節流
  const SL = (now, st) => MR.shouldLazySweep(now, st);
  const t0 = Date.parse('2026-10-24T05:00:00Z');
  eq([SL(t0, {}), SL(t0, { instLastAttemptMs: t0 - 30000, lastSweepDate: '2026-10-23' }), SL(t0, { instLastAttemptMs: t0 - 61000, lastSweepDate: '2026-10-23' })], [true, false, true], 'L1 今天還沒掃過 → 立刻掃；剛試過 60 秒內不重試');
  eq([SL(t0, { instLastAttemptMs: 0, lastSweepDate: '2026-10-24', lastSweepAt: new Date(t0 - 30 * 60000).toISOString() }), SL(t0, { instLastAttemptMs: 0, lastSweepDate: '2026-10-24', lastSweepAt: new Date(t0 - 61 * 60000).toISOString() }), SL(t0, { instLastAttemptMs: t0 - 30 * 60000, lastSweepDate: '2026-10-24', lastSweepAt: new Date(t0 - 90 * 60000).toISOString() })], [false, true, false], 'L2 今天掃過：距上次掃描（持久化）≥60 分鐘且本實例 ≥60 分鐘沒試才補掃');
  eq([SL(Date.parse('2026-10-24T16:30:00Z'), { instLastAttemptMs: 0, lastSweepDate: '2026-10-24', lastSweepAt: '2026-10-24T16:00:00.000Z' })], [true], 'L3 台北日期跨日（UTC 16:30＝台北隔天）：持久化日期已過期 → 立刻掃');
  eq(MR.sweep(mkData([ITEM()]), UU, { now: T(30), rules: [] }).stats, { evaluated: 1, fired: 0, recipients: 0, skippedDup: 0, noRecipient: 0, capped: 0, resolved: 0 }, 'L4 沒有任何規則 → 全零統計');
  const lm = mkData([ITEM()]); MR.sweep(lm, UU, { now: T(30), rules: one, mode: 'lazy' });
  eq([lm.reminderLog.lastSweepMode, lm.reminderLog.lastSweepDate, typeof lm.reminderLog.lastSweepAt], ['lazy', '2026-12-01', 'string'], 'L5 掃描後記錄 lastSweep（持久化，供補掃判斷）');
  // 試算器
  const sim = (rules, sc, o) => MR.simulate(rules, sc, Object.assign({ now: T(0) }, o || {}));
  const sd = sim(MR.DEFAULT_RULES, { daysLeft: 45, quoteState: 'none' });
  eq([sd.activeRuleIds, sd.byRole.manager1.map((x) => x.ruleId), sd.byRole.owner.map((x) => x.ruleId), sd.byRole.marketing.length], [['default-d90', 'default-d45'], ['default-d45'], ['default-d90'], 0], 'M1 試算（剩 45 天、沒有續約報價單）：D-90 與 D-45 成立；manager1 看到 D-45、owner 看到 D-90');
  eq(sim(MR.DEFAULT_RULES, { daysLeft: 20, quoteState: 'draft' }).activeRuleIds, ['default-d90', 'default-d30'], 'M2 剩 20 天、草稿 → D-90、D-30（D-45 因已有報價單不成立）');
  eq(sim(MR.DEFAULT_RULES, { daysLeft: 20, quoteState: 'sent' }).activeRuleIds, ['default-d90'], 'M3 剩 20 天、已送簽 → 只剩 always 的 D-90');
  eq(sim(MR.DEFAULT_RULES, { daysLeft: -10, quoteState: 'none' }).activeRuleIds, ['default-after7'], 'M4 到期後 10 天 → 只有「到期後」規則（到期前的規則不補發）');
  eq(sim(MR.DEFAULT_RULES, { daysLeft: 20, quoteState: 'none', notRenew: true }).activeRuleIds, [], 'M5 已標不續約 → 全部不成立');
  const sr = sim(MR.DEFAULT_RULES, { daysLeft: 20, quoteState: 'none', notRenew: true });
  ok(sr.rules.every((x) => x.reason === 'NOT_TRACKING') && sr.tracking === false, 'M5b 不續約時每條規則的原因都是 NOT_TRACKING');
  eq(sim(MR.DEFAULT_RULES, { daysLeft: 100, quoteState: 'none' }).rules.map((x) => x.reason), ['NOT_REACHED', 'NOT_REACHED', 'NOT_REACHED', 'NOT_REACHED'], 'M6 剩 100 天：全部「尚未達標」');
  const simItem = sim(MR.DEFAULT_RULES, { daysLeft: 45, quoteState: 'none' }, { item: ITEM({ owner: U.u6 }), users: UU, labelOf: (u) => '名:' + u });
  eq([simItem.recipients.owner, simItem.recipients.manager1.map((x) => x.username)], [[{ username: U.u6, label: '名:' + U.u6 }], [U.adm, U.adm2]], 'M7 帶真實納管項 → 用真實收件人解析（無一級主管退管理員），顯示名由 labelOf 提供');
  eq(sim([RULE({ whenDays: 30, appliesTo: ['license'] })], { daysLeft: 10, quoteState: 'none', termKind: 'ma' }).activeRuleIds, [], 'M8 試算依 termKind 判斷適用類型');
  eq(sim([RULE({ id: 'dis', whenDays: 30, enabled: false })], { daysLeft: 10, quoteState: 'none' }).rules[0].reason, 'DISABLED', 'M9 停用的規則原因 DISABLED');
  eq(MR.timeline({ reminderLog: { entries: { [MR.logKey('a', 'i1', U.u1, END, 0)]: '2026-12-01T00:00:00.000Z', [MR.logKey('a', 'i1', U.m1, END, 0)]: '2026-12-01T00:00:00.000Z', [MR.logKey('a', 'i1', U.u1, '2025-01-01', 0)]: 'x', [MR.logKey('a', 'other', U.u1, END, 0)]: 'x' } }, quotations: [] }, ITEM(), [RULE({ id: 'a', whenDays: 30 })], { now: T(25), labelOf: (u) => u })[0].sentCount, 2, 'T1 時間軸：只算這個納管項、這個迄日、這條規則的發送紀錄');
  const tl = MR.timeline({ quotations: [q('i1', 'none')] }, ITEM(), [RULE({ id: 'a', whenDays: 30, condition: 'no_renewal_quote' }), RULE({ id: 'b', whenDays: 30, condition: 'quote_not_submitted' })], { now: T(25) });
  eq([tl[0].triggerDate, tl[0].conditionNow, tl[0].activeNow, tl[0].reason, tl[1].conditionNow, tl[1].activeNow, tl[0].reached], ['2026-12-01', false, false, 'CONDITION', true, true, true], 'T2 時間軸：觸發日期＝迄日 − 天數；條件目前是否成立；原因');
  eq(Object.keys(MR.timeline({ quotations: [] }, ITEM(), [RULE()], { now: T(25), secretary: true })[0]).includes('sent'), false, 'T3 秘書的時間軸不含收件人名單');
}

// ═════════════ 日期推進模擬（注入 now 逐日前進）═════════════
/**
 * 從距迄日 from 天一路前進到 to 天（每天一次掃描），hook(dl, data) 可在掃描前改資料（建立報價單、標不續約…）。
 * skip(dl)→true 的日子「cron 沒跑」。回 events：[{ dl, rule, item, to }]（以去重鍵的新增判斷「這天實際發了什麼」）。
 */
function runDays(data, users, rules, from, to, hook, skip) {
  const events = [];
  for (let dl = from; dl >= to; dl--) {
    if (hook) hook(dl, data);
    if (skip && skip(dl)) continue;
    const before = new Set(Object.keys((data.reminderLog && data.reminderLog.entries) || {}));
    const r = REF.MR.sweep(data, users, { now: atDl(dl), rules, mode: 'cron' });
    Object.keys(data.reminderLog.entries).forEach((k) => {
      if (before.has(k)) return;
      const p = REF.MR.keyParts(k);
      events.push({ dl, rule: p.ruleId, item: p.itemId, to: p.recipient, period: Number(p.period) });
    });
  }
  return events;
}
const daysOf = (events, rule, to) => events.filter((e) => e.rule === rule && (!to || e.to === to)).map((e) => e.dl);
const bellOf = (data, users, dl, user, rules) => REF.MR.dynamicFor(REF.MR.plan(data, users, atDl(dl), rules), user, {}).map((x) => x.itemId);
const unreadOf = (data, user) => (data.notifications || []).filter((n) => n.to === user && n.type === 'ma_reminder' && !n.read).length;

function groupSim(which) {
  const MR = REF.MR;
  if (!which || which === 'sim') simPart(MR);
  if (!which || which === 'volume') volumePart(MR);
}
function simPart(MR) {
  console.log('[日期推進] 逐日掃描 +130 … -10 天');
  const UU = users23();
  const mkD = (item, qs) => ({ maItems: [item || ITEM()], quotations: qs || [], notifications: [] });

  // 1) 預設規則：沒有續約報價單
  let d = mkD();
  let ev = runDays(d, UU, MR.DEFAULT_RULES, 130, -10);
  eq(daysOf(ev, 'default-d90').sort(), [90, 90, 90], 'T1 D-90 規則：剩 90 天那天發給 owner＋2 位秘書，各一次');
  eq(ev.filter((e) => e.rule === 'default-d90').map((e) => e.to).sort(), [U.sec, U.sec3, U.u1].sort(), 'T1b D-90 收件人＝負責業務＋同 BU 在職秘書');
  eq(ev.filter((e) => e.rule === 'default-d45').map((e) => [e.dl, e.to]), [[45, U.m1]], 'T2 D-45（沒有續約報價單）：剩 45 天那天只發給一級主管，一次');
  eq(daysOf(ev, 'default-d30'), [], 'T3 D-30（草稿未送簽）：沒有續約報價單時不成立（依定義「有報價單但全是草稿」）');
  eq(ev.filter((e) => e.rule === 'default-after7').map((e) => [e.dl, e.to, e.period]).sort((a, b) => b[0] - a[0] || (a[1] < b[1] ? -1 : 1)), [[-1, U.m1, 0], [-1, U.u1, 0], [-8, U.m1, 1], [-8, U.u1, 1]], 'T4 到期後規則：到期隔天（-1）起每 7 天一個週期（-1、-8），每人每週期一次');
  eq(ev.length, 3 + 1 + 4, 'T5 整段 +130…-10 天共發 8 筆，沒有多餘');
  // 通知則數（按人彙總）：同一人同一天多筆規則只一則
  eq(d.notifications.filter((n) => n.to === U.u1).length, 3, 'T6 業務甲在 D-90、到期後 -1、-8 各收到一則彙總（沒有重複）');
  // 2) 草稿報價單（dl=60 建立）→ D-45 不成立；D-30 成立
  d = mkD();
  ev = runDays(d, UU, MR.DEFAULT_RULES, 130, -10, (dl, dd) => { if (dl === 60) dd.quotations.push({ id: 'q1', renewOf: 'i1', approval: null }); });
  eq([daysOf(ev, 'default-d45'), daysOf(ev, 'default-d30').sort()], [[], [30, 30, 30, 30]], 'T7 dl=60 建立草稿報價單：D-45 不再發；D-30 在剩 30 天發給 owner＋一級主管＋2 位行銷');
  eq(ev.filter((e) => e.rule === 'default-d30').map((e) => e.to).sort(), [U.m1, U.mkt, U.mkt2, U.u1].sort(), 'T7b D-30 收件人（離職行銷不在內）');
  // 3) 草稿在 dl=35 送簽 → D-30 不發
  d = mkD();
  ev = runDays(d, UU, MR.DEFAULT_RULES, 130, -10, (dl, dd) => { if (dl === 60) dd.quotations.push({ id: 'q1', renewOf: 'i1', approval: null }); if (dl === 35) dd.quotations[0].approval = { state: 'pending' }; });
  eq(daysOf(ev, 'default-d30'), [], 'T8 剩 35 天送簽 → 剩 30 天時 D-30 條件不成立，不發');
  // 4) 剛好等於 vs 補發：cron 漏跑 3 天
  d = mkD();
  ev = runDays(d, UU, MR.DEFAULT_RULES, 130, -10, null, (dl) => dl >= 43 && dl <= 47);   // 47..43 沒跑（含 45）
  eq(ev.filter((e) => e.rule === 'default-d45').map((e) => e.dl), [42], 'T9 剩 45 天那天 cron 沒跑（漏 5 天）→ 隔天（剩 42 天）補發，只一次');
  d = mkD();
  ev = runDays(d, UU, MR.DEFAULT_RULES, 130, -10, null, (dl) => dl === 90 || dl === 89 || dl === 88);
  eq(ev.filter((e) => e.rule === 'default-d90').map((e) => e.dl).sort(), [87, 87, 87], 'T10 漏跑 3 天（90、89、88）→ 剩 87 天補發給同樣 3 人，各一次');
  d = mkD();
  ev = runDays(d, UU, MR.DEFAULT_RULES, 130, -10, null, (dl) => dl === 45);
  eq(ev.filter((e) => e.rule === 'default-d45').map((e) => e.dl), [44], 'T10b 只漏剛好那一天 → 隔天補發');
  // 5) 長時間停擺：到期前的規則不補發，到期後規則補發一次
  d = mkD();
  ev = runDays(d, UU, MR.DEFAULT_RULES, 130, -10, null, (dl) => dl < 120 && dl > -5);
  eq([ev.filter((e) => e.rule !== 'default-after7').length, ev.filter((e) => e.rule === 'default-after7').map((e) => [e.dl, e.period]).sort().slice(0, 2)], [0, [[-5, 0], [-5, 0]]], 'T11 整段停擺到過期後才恢復：D-90／D-45 不再補發（已過期），到期後規則在恢復的第一天發一次（週期 0）');
  // 6) repeat 週期：漏跑不重複
  d = mkD();
  ev = runDays(d, UU, MR.DEFAULT_RULES, 5, -40, null, (dl) => dl === -8 || dl === -9 || dl === -10 || dl === -16);
  eq([...new Set(ev.filter((e) => e.rule === 'default-after7' && e.to === U.u1).map((e) => e.dl + ':' + e.period))].sort(), ['-1:0', '-11:1', '-15:2', '-22:3', '-29:4', '-36:5'].sort(), 'T12 repeat：週期內漏跑（-8…-10）當週期在 -11 補發一次；-16 漏跑但同週期(-15 已發)不重複');
  // 7) 自訂規則 + 條件變化 → 鈴鐺自動消失、通知被標已讀
  const RC1 = [
    RULE({ id: 'c45', name: '無報價', whenDays: 45, notify: ['owner', 'manager1'], condition: 'no_renewal_quote' }),
    RULE({ id: 'c30', name: '未送簽', whenDays: 30, notify: ['owner', 'manager1', 'marketing'], condition: 'quote_not_submitted' }),
    RULE({ id: 'c0', name: '未續約', whenDays: -1, repeatEveryDays: 7, notify: ['owner', 'manager1'], condition: 'not_renewed' }),
  ];
  d = mkD();
  const log = [];
  ev = runDays(d, UU, RC1, 60, -20, (dl, dd) => {
    if (dl === 44) dd.quotations.push({ id: 'q1', renewOf: 'i1', approval: null });   // 建立草稿
    if (dl === 29) dd.quotations[0].approval = { state: 'pending' };                  // 送簽
    if (dl === 10) { dd.quotations[0].approval = { state: 'approved' }; dd.maItems[0].status = 'renewed'; }   // 續約成交
    // 掃描前先記錄鈴鐺（上一天掃描後的狀態、套用今天的變更後）
    log.push([dl, bellOf(dd, UU, dl, U.u1, RC1).length, bellOf(dd, UU, dl, U.m1, RC1).length, bellOf(dd, UU, dl, U.mkt, RC1).length]);
  });
  const bell = Object.fromEntries(log.map((x) => [x[0], x.slice(1)]));
  eq([bell[46], bell[45], bell[44]], [[0, 0, 0], [1, 1, 0], [0, 0, 0]], 'T13 鈴鐺：剩 45 天出現（owner／一級主管），剩 44 天建立草稿後自動消失');
  eq([bell[30], bell[29]], [[1, 1, 1], [0, 0, 0]], 'T13b 剩 30 天草稿未送簽 → owner、一級主管、行銷都出現；剩 29 天送簽後消失');
  eq([bell[10], bell[5], bell[-5]], [[0, 0, 0], [0, 0, 0], [0, 0, 0]], 'T13c 續約成交（status＝renewed）後完全沒有提醒');
  eq([ev.filter((e) => e.rule === 'c45').map((e) => e.dl), ev.filter((e) => e.rule === 'c30').length, ev.filter((e) => e.rule === 'c0').length], [[45, 45], 4, 0], 'T14 發送：c45 剩 45 天（owner＋一級主管）、c30 剩 30 天 4 人、續約後 c0 一次都沒發');
  // 通知被標已讀：模擬 poll（每天由使用者自己的 poll-bundle 觸發 reconcile）
  d = mkD();
  runDays(d, UU, RC1, 60, 40);
  eq([unreadOf(d, U.u1), unreadOf(d, U.m1)], [1, 1], 'T15 剩 40 天：owner、一級主管各有 1 則未讀 ma_reminder');
  d.quotations.push({ id: 'q1', renewOf: 'i1', approval: null });
  const mk1 = MR.reconcileNotifications(d, UU, atDl(39), RC1, (n) => n.to === U.u1);
  eq([mk1, unreadOf(d, U.u1), unreadOf(d, U.m1)], [1, 0, 1], 'T15b 業務建立續約草稿後，他自己的 poll 就把他的通知標已讀（一級主管的下次 poll 或掃描才標）');
  runDays(d, UU, RC1, 39, 39);
  eq(unreadOf(d, U.m1), 0, 'T15c 下次掃描把其餘收件人的通知也標已讀');
  // 不續約：已發的通知被標已讀，之後不再發
  d = mkD();
  ev = runDays(d, UU, RC1, 60, -20, (dl, dd) => { if (dl === 40) dd.maItems[0].status = 'not_renew'; });
  eq([ev.map((e) => e.dl).filter((x) => x < 45).length, unreadOf(d, U.u1), unreadOf(d, U.m1)], [0, 0, 0], 'T16 剩 40 天標不續約：之後沒有任何發送，已發的通知也被標已讀');
  // 撤銷不續約：恢復追蹤 → 到期後規則仍會發（去重鍵沒發過）
  d = mkD();
  ev = runDays(d, UU, RC1, 60, -7, (dl, dd) => { if (dl === 40) dd.maItems[0].status = 'not_renew'; if (dl === -3) dd.maItems[0].status = 'tracking'; });
  eq(ev.filter((e) => e.rule === 'c0').map((e) => e.dl).sort(), [-3, -3], 'T17 -3 撤銷不續約 → 當天補發到期後規則（owner＋一級主管）');
  // 修改期間（延期）→ 重新提醒
  d = mkD();
  ev = runDays(d, UU, RC1, 50, 44, null);
  d.maItems[0].to = '2027-06-30'; d.maItems[0].from = '2026-07-01';
  const evB = [];
  for (let dl = 46; dl >= 40; dl--) { const now = Date.parse('2027-06-30T04:00:00Z') - dl * DAY; const b = Object.keys(d.reminderLog.entries).length; MR.sweep(d, UU, { now, rules: RC1 }); if (Object.keys(d.reminderLog.entries).length > b) evB.push(dl); }
  eq([ev.length, evB], [2, [45]], 'T18 延期（改迄日）後以新迄日重新計算：新的剩 45 天再提醒一次（去重鍵含迄日）');
  // 類型限定
  const RLIC = [RULE({ id: 'lic', whenDays: 30, appliesTo: ['license'], notify: ['owner'] })];
  d = { maItems: [ITEM({ id: 'ma1', termKind: 'ma' }), ITEM({ id: 'li1', termKind: 'license' })], quotations: [], notifications: [] };
  ev = runDays(d, UU, RLIC, 40, 20);
  eq(ev.map((e) => e.item), ['li1'], 'T19 appliesTo＝license 的規則只對 License 項發');
  // 規則改變：新增規則即時對已達門檻的項目生效（補發）
  d = mkD();
  runDays(d, UU, [RULE({ id: 'old', whenDays: 90, notify: ['owner'] })], 100, 50);
  ev = runDays(d, UU, [RULE({ id: 'old', whenDays: 90, notify: ['owner'] }), RULE({ id: 'new', whenDays: 60, notify: ['manager1'], condition: 'always', name: '新規則' })], 49, 49);
  eq(ev.map((e) => [e.rule, e.to]), [['new', U.m1]], 'T20 管理員新增規則後，已達門檻的項目在下次掃描補發（舊規則不重複）');
  // 多個納管項、不同日期
  d = { maItems: [ITEM({ id: 'x1', to: plusDays(END, 0) }), ITEM({ id: 'x2', to: plusDays(END, 10), from: '2026-01-11' })], quotations: [], notifications: [] };
  ev = [];
  for (let dl = 100; dl >= 80; dl--) { const now = atDl(dl); const b = new Set(Object.keys((d.reminderLog || { entries: {} }).entries)); MR.sweep(d, UU, { now, rules: [RULE({ id: 'p', whenDays: 90, notify: ['owner'] })] }); Object.keys(d.reminderLog.entries).filter((k) => !b.has(k)).forEach((k) => ev.push([dl, MR.keyParts(k).itemId])); }
  eq(ev, [[90, 'x1'], [80, 'x2']], 'T21 兩個納管項各自在自己的剩 90 天達標（x1 在 90、x2 迄日晚 10 天所以在掃描區間的最後一天 80）');

}
function volumePart(MR) {
  const UU = users23();
  let d;
  console.log('[通知量] 100 個納管項同時達標');
  const owners = ['_s_u1', '_s_u2', '_s_u3', '_s_u4', '_s_u5'];
  const items = Array.from({ length: 100 }, (_, i) => ITEM({ id: 'v' + i, owner: owners[i % 5], company: '量測公司' + i }));
  d = { maItems: items, quotations: [], notifications: [] };
  const vr = [RULE({ id: 'v', whenDays: 30, notify: ['owner', 'manager1', 'marketing', 'secretary'] })];
  let s = MR.sweep(d, UU, { now: atDl(30), rules: vr });
  const per = {};
  d.notifications.forEach((n) => { per[n.to] = (per[n.to] || 0) + 1; });
  ok(Object.values(per).every((n) => n === 1), 'N1 每個收件人只收到一則彙總（沒有人收到多則）', J(per));
  const cnt = {};
  d.notifications.forEach((n) => { cnt[n.to] = n.maItemIds.length; });
  eq([cnt[U.u1], cnt[U.u2], cnt[U.m1], cnt[U.mkt], cnt[U.sec]], [20, 20, 100 - 20, 100, 80], 'N2 彙總的項數正確：業務各 20 項；一級主管 m1 負責 u1／u2／u4／u5 線 80 項（u3 屬於 m1b 的 20 項不是他的）；行銷 100；ERP 秘書 80');
  ok(s.stats.fired > 100 && d.notifications.length === s.created.length && s.created.length < 20, 'N3 發出的通知則數是「人數」而不是「納管項 × 規則」', `通知 ${s.created.length} 則／去重標記 ${s.stats.fired} 筆`);
  s = MR.sweep(d, UU, { now: atDl(30), rules: vr });
  eq([s.created.length, s.stats.fired, d.notifications.length], [0, 0, per ? Object.keys(per).length : 0], 'N4 第二次掃描完全不重複');
  s = MR.sweep(d, UU, { now: atDl(25), rules: vr });
  eq(s.created.length, 0, 'N4b 隔幾天再掃也不重複');
  // 300 個收件人、上限 200
  const many = Array.from({ length: 300 }, (_, i) => ITEM({ id: 'w' + i, owner: 'o' + String(i).padStart(3, '0') }));
  const manyUsers = [{ username: U.adm, role: 'admin' }].concat(many.map((m) => ({ username: m.owner, role: 'user', bu: ['ERP'] })));
  d = { maItems: many, quotations: [], notifications: [] };
  s = MR.sweep(d, manyUsers, { now: atDl(30), rules: [RULE({ notify: ['owner'] })] });
  eq([s.created.length, s.stats.capped, d.notifications.length], [200, 100, 200], 'N5 300 位收件人：單次只發 200 則，其餘 100 位 capped');
  s = MR.sweep(d, manyUsers, { now: atDl(29), rules: [RULE({ notify: ['owner'] })] });
  eq([s.created.length, s.stats.capped, s.stats.skippedDup], [100, 0, 200], 'N6 下一次掃描補發 100 位，已發的 200 位不重複');
  s = MR.sweep(d, manyUsers, { now: atDl(28), rules: [RULE({ notify: ['owner'] })] });
  eq(s.created.length, 0, 'N7 之後不再發');
  // 全站通知 500 上限：已有 480 則他人通知 + 本次 200 則
  d = { maItems: many, quotations: [], notifications: Array.from({ length: 480 }, (_, i) => ({ id: 'old' + i, to: 'x', type: 'other', read: i < 400, createdAt: nowIso })) };
  s = MR.sweep(d, manyUsers, { now: atDl(30), rules: [RULE({ notify: ['owner'] })] });
  eq([d.notifications.length, d.notifications.filter((n) => n.type === 'other' && !n.read).length], [500, 80], 'N8 擠掉的是「已讀且最舊」的他人通知（400 則已讀被丟掉 180 則），80 則未讀的他人通知全部保留');
}

// ═════════════ API ═════════════
/** 共用種子：一批不同狀態的納管項、商機（可連報價單）、合約 */
function apiSeed(d) {
  const T0 = todayTp();
  d.opportunities = [
    mkOpp('oQ1', U.u1, MAP, 'C', { company: 'CoA' }), mkOpp('oQ2', U.u2, MAP, 'C', { company: 'CoB' }), mkOpp('oQZ', U.u1, MAP, 'C', { company: 'CoZ2' }),
    mkOpp('oR', U.u1, MAP, 'C', { company: 'CoA' }),
  ];
  d.maItems = [
    mkItem('iA', U.u1, 40, { company: 'CoA', contractNo: CNO }), mkItem('iB', U.u2, 40, { company: 'CoB' }), mkItem('iC', U.u3, 40, { company: 'CoC', bu: 'ITS' }),
    mkItem('iD', U.u1, 100, { company: 'CoA' }), mkItem('iE', U.u1, -20, { company: 'CoA' }), mkItem('iF', U.u1, 10, { company: 'CoA', status: 'renewed' }),
    mkItem('iV', U.u1, 10, { company: 'CoA', status: 'void' }), mkItem('iG', U.uoff, 40, { company: 'CoG' }), mkItem('iH', U.u6, 40, { company: 'CoH', bu: 'CRM' }),
    mkItem('iZ', U.u1, 20, { company: 'CoZ' }), mkItem('iL', U.u1, 40, { company: 'CoL', termKind: 'license' }),
  ];
  d.contracts = [
    { id: 'cSM', owner: U.u1, contractNo: 'SM-1', company: 'CoA', product: 'Svc', startDate: plusDays(T0, -300), endDate: plusDays(T0, 20), amount: '1', type: 'SERVICE_MA', createdAt: nowIso },
    { id: 'cEM', owner: U.u1, contractNo: 'EM-1', company: 'CoA', product: 'ERP', startDate: plusDays(T0, -300), endDate: plusDays(T0, 20), amount: '1', type: 'ERP_MA', createdAt: nowIso },
  ];
}
const qBody = (oppId, extra) => Object.assign({ oppId, projectName: 'P', products: [], validUntil: '2099-01-31', items: [{ desc: '顧問服務', unit: '式', qty: 1, unitPrice: 100 }] }, extra || {});

/** 小型偽隨機（可重現） */
function prng(seed) { let a = seed >>> 0; return () => { a = (a + 0x6D2B79F5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; }; }
function randomRules(rnd) {
  const n = 1 + Math.floor(rnd() * 6);
  const used = new Set();
  const out = [];
  const conds = ['always', 'no_renewal_quote', 'quote_not_submitted', 'not_renewed'];
  for (let i = 0; i < n * 3 && out.length < n; i++) {
    const cond = conds[Math.floor(rnd() * 4)];
    const ap = [['ma'], ['license'], 'all'][Math.floor(rnd() * 3)];
    const when = Math.floor(rnd() * 160) - 40;
    const k = cond + '|' + (ap === 'all' ? 'all' : ap.join()) + '|' + when;
    if (used.has(k)) continue;
    used.add(k);
    const roles = ['owner', 'secretary', 'manager1', 'manager2', 'marketing'].filter(() => rnd() < 0.5);
    out.push({ id: 'rr' + out.length, name: '隨機規則' + out.length, enabled: rnd() < 0.85, whenDays: when, repeatEveryDays: when <= 0 && rnd() < 0.5 ? 1 + Math.floor(rnd() * 20) : 0, notify: roles.length ? roles : ['owner'], condition: cond, appliesTo: ap });
  }
  return out;
}

const R1 = (o) => RULE(o);
const QA = require(path.join(ROOT, 'lib', 'quoteApproval.js'));
const COND = [R1({ id: 'k45', name: '無報價', whenDays: 45, notify: ['owner', 'manager1'], condition: 'no_renewal_quote' }), R1({ id: 'kq', name: '未送簽', whenDays: 45, notify: ['owner', 'manager1', 'marketing'], condition: 'quote_not_submitted' }), R1({ id: 'k0', name: '未續約', whenDays: -1, repeatEveryDays: 7, notify: ['owner', 'manager1'], condition: 'not_renewed' })];
const NR = (who, id, reason) => http(who, 'PATCH', '/api/ma-items/' + id + '/not-renew', reason === undefined ? {} : { reason });
async function groupApi(sub) {
  const S = (k) => !sub || sub === k;
  console.log('[API] 規則、試算器、cron＋補掃、renewOf、續約成交、不續約、時間軸、續約候選' + (sub ? '（' + sub + '）' : ''));
  if (S('rules')) await apiRules();
  if (S('main')) await apiMain();
  if (S('flow')) await apiFlow();
  if (S('dual')) await apiDual();
  if (S('fail')) await apiFail();
}
async function apiRules() {
  const T0 = todayTp();
  // ═════ A. 規則 GET／PUT／驗證 ═════
  await freshServer({ data: apiSeed });
  let bad = await loginMany([U.adm, U.u1, U.sec, U.mkt, U.m1]);
  eq(bad, [], 'A0 登入');
  let r = await http(U.adm, 'GET', '/api/admin/reminder-rules');
  ok(r.status === 200 && r.json.config.isDefault === true && r.json.config.rules.length === 4 && r.json.defaults.length === 4 && r.json.limits.maxRules === 20 && r.json.options.conditions.length === 4 && r.json.options.notify.length === 5, 'A1 管理員 GET：預設 4 條、limits、options', r.text.slice(0, 200));
  eq([r.json.config.updatedAt, r.json.config.updatedBy, r.headers.get('cache-control')], [null, null, 'no-store'], 'A1b 尚未儲存：updatedAt／updatedBy 為 null；no-store');
  for (const who of [U.u1, U.sec, U.mkt, U.m1]) eq([(await http(who, 'GET', '/api/admin/reminder-rules')).status, (await http(who, 'PUT', '/api/admin/reminder-rules', { rules: [] })).status, (await http(who, 'POST', '/api/admin/reminder-rules/simulate', { daysLeft: 1, quoteState: 'none' })).status], [403, 403, 403], `A2 非管理員 ${who} 三個管理端點全 403`);
  eq([(await http(null, 'GET', '/api/admin/reminder-rules')).status, (await http(null, 'PUT', '/api/admin/reminder-rules', { rules: [] })).status], [401, 401], 'A2b 未登入 401');
  ok(J(readData().reminderRules) === undefined || readData().reminderRules === undefined, 'A2c 被拒絕的 PUT 沒有寫入');
  const R1 = (o) => RULE(o);
  const putCode = async (rules, body) => { const x = await http(U.adm, 'PUT', '/api/admin/reminder-rules', Object.assign({ rules }, body || {})); return x.status + ':' + (x.json && x.json.code); };
  const cases = [
    ['非陣列', 'x', '400:BAD_REMINDER_RULES'], ['21 條', Array.from({ length: 21 }, (_, i) => R1({ id: 'a' + i, whenDays: i })), '400:TOO_MANY_RULES'],
    ['代號不合法', [R1({ id: 'a b' })], '400:BAD_RULE_ID'], ['代號重複', [R1({ id: 'a' }), R1({ id: 'a', whenDays: 5 })], '400:DUPLICATE_RULE_ID'],
    ['名稱空白', [R1({ name: ' ' })], '400:BAD_RULE_NAME'], ['名稱 41 字', [R1({ name: 'x'.repeat(41) })], '400:BAD_RULE_NAME'],
    ['天數 366', [R1({ whenDays: 366 })], '400:BAD_RULE_WHEN'], ['天數 -366', [R1({ whenDays: -366 })], '400:BAD_RULE_WHEN'], ['天數小數', [R1({ whenDays: 1.5 })], '400:BAD_RULE_WHEN'],
    ['重複 91', [R1({ whenDays: -1, repeatEveryDays: 91 })], '400:BAD_RULE_REPEAT'], ['到期前設重複', [R1({ whenDays: 5, repeatEveryDays: 7 })], '400:BAD_RULE_REPEAT'],
    ['對象空', [R1({ notify: [] })], '400:BAD_RULE_NOTIFY'], ['對象不在白名單', [R1({ notify: ['admin'] })], '400:BAD_RULE_NOTIFY'],
    ['條件不在白名單', [R1({ condition: 'x' })], '400:BAD_RULE_CONDITION'], ['適用類型不合法', [R1({ appliesTo: ['x'] })], '400:BAD_RULE_APPLIES'],
    ['天數重複', [R1({ id: 'a' }), R1({ id: 'b' })], '400:DUPLICATE_RULE_WHEN'], ['enabled 非布林', [R1({ enabled: 'y' })], '400:BAD_RULE_ENABLED'],
  ];
  for (const [nm, rules, exp] of cases) eq(await putCode(rules), exp, `A3 PUT 驗證：${nm} → ${exp}`);
  ok(readData().reminderRules === undefined, 'A3b 所有被拒絕的 PUT 都沒有寫入');
  eq((await http(U.adm, 'PUT', '/api/admin/reminder-rules', {})).json.code, 'BAD_REMINDER_RULES', 'A3c 沒有 rules → BAD_REMINDER_RULES');
  const good = [R1({ id: 'g1', name: '三十天', whenDays: '30', notify: ['owner', 'marketing'] }), R1({ id: 'g2', name: '到期後', whenDays: 0, repeatEveryDays: '5', notify: ['manager1'], condition: 'not_renewed', appliesTo: ['ma', 'license'] })];
  r = await http(U.adm, 'PUT', '/api/admin/reminder-rules', { rules: good });
  ok(r.status === 200 && r.json.success === true && r.json.config.rules[0].whenDays === 30 && r.json.config.rules[1].appliesTo === 'all' && r.json.config.rules[1].repeatEveryDays === 5, 'A4 合法 PUT → 200，數字字串正規化、兩種類型＝all', r.text.slice(0, 200));
  r = await http(U.adm, 'GET', '/api/admin/reminder-rules');
  ok(r.json.config.isDefault === false && r.json.config.rules.length === 2 && r.json.config.updatedBy === U.adm && /^\d{4}-/.test(r.json.config.updatedAt), 'A4b 讀回：非預設、updatedBy＝管理員');
  let lg = await auditLines('SAVE_REMINDER_RULES');
  ok(lg.length === 1 && /新增規則/.test(lg[0].detail) && /移除規則/.test(lg[0].detail) && lg[0].operator === U.adm, 'A4c 稽核 SAVE_REMINDER_RULES 含差異（新增／移除）', lg[0] && lg[0].detail);
  { const noIds = good.map((x) => { const y = Object.assign({}, x); delete y.id; return y; });
    r = await http(U.adm, 'PUT', '/api/admin/reminder-rules', { rules: noIds });
    eq([r.status, r.json.config.rules.map((x) => x.id)], [200, ['g1', 'g2']], 'A4d 前端存檔時沒帶代號：沿用現行規則的代號（去重鍵不換新，不會把已發過的提醒重發一輪）');
    r = await http(U.adm, 'PUT', '/api/admin/reminder-rules', { rules: noIds.concat([R1({ id: undefined, name: '新的', whenDays: 77 })]) });
    ok(r.status === 200 && r.json.config.rules.length === 3 && /^r_[0-9a-f]{10}$/.test(r.json.config.rules[2].id), 'A4e 全新的規則才產生新代號'); }
  r = await http(U.adm, 'PUT', '/api/admin/reminder-rules', { rules: [] });
  ok(r.status === 200, 'A5 空陣列可（全部關閉）');
  eq((await http(U.adm, 'GET', '/api/admin/reminder-rules')).json.config.rules.length, 0, 'A5b 讀回 0 條（不會偷偷變回預設）');
  writeData((d) => { d.reminderRules = { rules: [{ id: 'bad id', name: '', whenDays: 9999 }], updatedBy: 'x' }; });
  r = await http(U.adm, 'GET', '/api/admin/reminder-rules');
  ok(r.json.config.isDefault === true && r.json.config.rules.length === 4, 'A6 儲存值被竄改成垃圾 → GET 回預設（不信任儲存值）');
  // ═════ B. 試算器 ═════
  const SIM = (b, who) => http(who || U.adm, 'POST', '/api/admin/reminder-rules/simulate', b);
  eq([(await SIM({ daysLeft: 'x', quoteState: 'none' })).json.code, (await SIM({ daysLeft: 1001, quoteState: 'none' })).json.code, (await SIM({ daysLeft: 1.5, quoteState: 'none' })).json.code, (await SIM({ quoteState: 'none' })).json.code], ['BAD_DAYS_LEFT', 'BAD_DAYS_LEFT', 'BAD_DAYS_LEFT', 'BAD_DAYS_LEFT'], 'B1 daysLeft 驗證');
  eq([(await SIM({ daysLeft: 1, quoteState: 'x' })).json.code, (await SIM({ daysLeft: 1 })).json.code, (await SIM({ daysLeft: 1, quoteState: 'none', notRenew: 'yes' })).json.code, (await SIM({ daysLeft: 1, quoteState: 'none', termKind: 'x' })).json.code, (await SIM({ daysLeft: 1, quoteState: 'none', rules: [R1({ whenDays: 999 })] })).json.code, (await SIM({ daysLeft: 1, quoteState: 'none', maItemId: 'nope' })).json.code, (await SIM({ daysLeft: 1, quoteState: 'none', maItemId: 5 })).json.code],
    ['BAD_QUOTE_STATE', 'BAD_QUOTE_STATE', 'BAD_NOT_RENEW', 'BAD_TERM_KIND', 'BAD_RULE_WHEN', 'ITEM_NOT_FOUND', 'BAD_ITEM_ID'], 'B2 其他參數驗證');
  r = await SIM({ daysLeft: 45, quoteState: 'none' });
  ok(r.status === 200 && r.json.source === 'saved' && r.json.activeRuleIds.join() === 'default-d90,default-d45' && r.json.byRole.manager1.length === 1 && J(r.json.recipients) === '{}', 'B3 預設規則剩 45 天：D-90＋D-45；沒帶納管項不解析帳號', r.text.slice(0, 300));
  r = await SIM({ daysLeft: 45, quoteState: 'none', maItemId: 'iA' });
  ok(r.json.recipients.owner[0].username === U.u1 && r.json.recipients.owner[0].label === '甲' && r.json.recipients.manager1[0].username === U.m1 && r.json.recipients.secretary.length === 2 && r.json.recipients.marketing.length === 2 && r.json.item.company === 'CoA' && clean(r), 'B4 帶 maItemId：用真實收件人（顯示名＝暱稱）；不含金額與合約編號', r.text.slice(0, 300));
  eq(Object.keys(r.json.item).sort(), ['company', 'id', 'ownerLabel', 'product', 'termKind'], 'B4b item 摘要只有 5 個欄位');
  r = await SIM({ daysLeft: 5, quoteState: 'none', rules: good });
  ok(r.json.source === 'draft' && r.json.activeRuleIds.join() === 'g1', 'B5 帶 rules（尚未儲存的草稿）→ 以草稿試算');
  // 隨機對拍：simulate API ＝ 獨立 oracle ＝ 真實引擎 plan
  const rnd = prng(20261024);
  let parity = 0, parityBad = [];
  for (let i = 0; i < 60; i++) {
    const rules = randomRules(rnd);
    if (!rules.length) continue;
    const dl = Math.floor(rnd() * 200) - 70;
    const qs = ['none', 'draft', 'sent'][Math.floor(rnd() * 3)];
    const nr = rnd() < 0.15;
    const kind = rnd() < 0.5 ? 'ma' : 'license';
    const withItem = i % 3 === 0;
    const res = await SIM({ rules, daysLeft: dl, quoteState: qs, notRenew: nr, termKind: withItem ? undefined : kind, maItemId: withItem ? 'iA' : undefined });
    if (res.status !== 200) { parityBad.push('HTTP ' + res.status + ' ' + res.text.slice(0, 120)); continue; }
    const k = withItem ? 'ma' : kind;
    const exp = rules.filter((x) => oracleActive(x, k, dl, qs, nr ? 'not_renew' : 'tracking')).map((x) => x.id);
    // 真實引擎
    const to = plusDays(T0, dl);
    const data = { maItems: [{ id: 'sim-item', owner: U.u1, company: 'c', product: 'p', termKind: k, from: plusDays(to, -364), to, status: nr ? 'not_renew' : 'tracking' }], quotations: qs === 'none' ? [] : [{ id: 'q', renewOf: 'sim-item', approval: { state: qs === 'sent' ? 'pending' : 'none' } }] };
    const eng = REAL.MR.plan(data, [], Date.now(), REAL.MR.validateConfig({ rules }).value.rules, { noRecipients: true });
    const engIds = eng.entries.length ? eng.entries[0].hits.map((h) => h.rule.id) : [];
    const byRoleExp = {};
    ['owner', 'secretary', 'manager1', 'manager2', 'marketing'].forEach((role) => { byRoleExp[role] = exp.filter((id) => rules.find((x) => x.id === id).notify.includes(role)); });
    const byRoleGot = {};
    Object.keys(res.json.byRole).forEach((role) => { byRoleGot[role] = res.json.byRole[role].map((x) => x.ruleId); });
    if (J(res.json.activeRuleIds) === J(exp) && J(engIds) === J(exp) && J(byRoleGot) === J(byRoleExp)) parity++; else parityBad.push(J({ rules: rules.map((x) => [x.id, x.condition, x.appliesTo, x.whenDays, x.enabled]), dl, qs, nr, k, got: res.json.activeRuleIds, exp, engIds }));
  }
  ok(parity >= 50 && parityBad.length === 0, `B6 隨機 60 組（規則×剩餘天數×報價單狀態×不續約×類型）：試算 API ＝ 獨立 oracle ＝ 真實引擎（${parity} 組一致）`, parityBad.slice(0, 2).join(' || '));

}
async function apiMain() {
  const T0 = todayTp();
  let bad, r, lg, cr, dd, pb;
  // ═════ C. 重新開機（預設規則）：cron 先於任何 poll ═════
  await freshServer({ data: apiSeed });
  bad = await loginMany([U.adm, U.u1, U.u2, U.m1, U.sec, U.mkt, U.u3, U.m1b]);
  eq(bad, [], 'C0 登入');
  eq([(await http(null, 'GET', '/api/cron/mail-outbox')).status, (await http(null, 'GET', '/api/cron/mail-outbox', undefined, { Authorization: 'Bearer wrong' })).status], [401, 401], 'C1 cron 沒有／錯誤密鑰 → 401');
  cr = await cron();
  ok(cr.status === 200 && cr.json.ok === true && cr.json.enabled === false && cr.json.reminders && cr.json.reminders.fired > 0 && cr.json.reminders.evaluated > 0, 'C2 cron（MAIL off 分支）200 且帶 reminders 統計', cr.text.slice(0, 300));
  eq(Object.keys(cr.json.reminders).sort(), ['capped', 'evaluated', 'fired', 'noRecipient', 'notifications', 'recipients', 'resolved', 'skippedDup', 'today'], 'C2b 統計欄位（含必要的 evaluated/fired/recipients/skippedDup/noRecipient/capped）');
  dd = readData();
  const mr = (to) => (dd.notifications || []).filter((n) => n.to === to && n.type === 'ma_reminder');
  ok(mr(U.u1).length === 1 && mr(U.m1).length === 1 && mr(U.sec).length === 1, 'C3 剩 40 天（預設規則）：業務甲、一級主管、秘書各收到 1 則彙總通知', J({ u1: mr(U.u1).length, m1: mr(U.m1).length, sec: mr(U.sec).length }));
  ok(mr(U.u1)[0].maItemIds.sort().join() === 'iA,iL,iZ'.split(',').sort().join() || mr(U.u1)[0].maItemIds.includes('iA'), 'C3b 業務甲的彙總包含 iA（dl=40，D-90）', J(mr(U.u1)[0].maItemIds));
  ok(!FORBIDDEN_TOKENS.some((t) => J(dd.notifications).includes(t)), 'C3c 通知沒有金額、備註、合約編號');
  eq(mr(U.u3).length + mr(U.m1b).length > 0, true, 'C3d ITS 業務丙的項目 → 通知丙與 ITS 一級主管');
  eq(mr(U.mkt).length, 0, 'C3e 預設規則下行銷在剩 40 天時沒有通知（D-30 條件不成立）');
  ok(mr(U.adm).length === 1, 'C3f iH（CRM、無主管、無秘書）退到管理員：管理員收到 1 則');
  lg = await auditLines('MA_REMINDER_NO_RECIPIENT');
  ok(lg.length >= 2 && lg.every((x) => x.operator === 'system') && lg.some((x) => /秘書/.test(x.detail)) && lg.some((x) => /一級主管/.test(x.detail)), 'C3g 稽核 MA_REMINDER_NO_RECIPIENT（找不到秘書／一級主管，已改通知管理員）', lg.map((x) => x.detail).join(' / ').slice(0, 300));
  lg = await auditLines('MA_REMINDER_SWEEP');
  ok(lg.length === 1 && /模式 cron/.test(lg[0].detail) && /發出 \d+ 筆提醒給 \d+ 人/.test(lg[0].detail), 'C3h 稽核 MA_REMINDER_SWEEP（只記統計）', lg[0] && lg[0].detail);
  const total1 = (readData().notifications || []).length;
  cr = await cron();
  ok(cr.json.reminders.fired === 0 && cr.json.reminders.skippedDup > 0 && (readData().notifications || []).length === total1, 'C4 第二次 cron：全部略過（skippedDup），沒有新增任何通知（冪等）', cr.text.slice(0, 200));
  ok(readData().reminderLog.lastSweepMode === 'cron', 'C4b 持久化 lastSweep');
  { // 秘書看不到納管項代號：通知上的 maItemIds 對秘書拿掉；其他角色保留（供前端比對）
    const sn = await notifs(U.sec, 'ma_reminder'), un = await notifs(U.u1, 'ma_reminder');
    ok(sn.length === 1 && !('maItemIds' in sn[0]) && Array.isArray(un[0].maItemIds), 'C3i /api/notifications：秘書的 ma_reminder 沒有 maItemIds、業務有');
    const pbs = await poll(U.sec);
    ok(pbs.notifications.filter((n) => n.type === 'ma_reminder').length === 1 && pbs.notifications.every((n) => !('maItemIds' in n)), 'C3j poll-bundle：秘書的 notifications 同樣沒有 maItemIds');
  }
  // poll-bundle：今天已掃過 → 不再補掃
  pb = await poll(U.u1);
  ok(Array.isArray(pb.maReminders) && pb.maReminders.length >= 1 && readData().reminderLog.lastSweepMode === 'cron', 'C5 poll-bundle：今天 cron 已掃過 → 不補掃（lastSweepMode 仍是 cron）；有 maReminders', J(Object.keys(pb)));
  const ma1 = pb.maReminders.find((x) => x.itemId === 'iA');
  ok(ma1 && ma1.type === 'ma_reminder' && ma1.days === 40 && ma1.level === 'info' && ma1.company === 'CoA' && /剩 40 天到期/.test(ma1.title) && ma1.quoteState === 'none' && ma1.rule.id === 'default-d90' && ma1.id === 'ma_iA' && ma1.roles.join() === 'owner', 'C5b 動態項目欄位：days／level／title／quoteState／主規則／roles', J(ma1));
  ok(clean({ text: J(pb.maReminders) }), 'C5c maReminders 沒有金額、備註、合約編號');
  ok(!pb.maReminders.some((x) => x.itemId === 'iD' || x.itemId === 'iF' || x.itemId === 'iV'), 'C5d 剩 100 天（未達）、已續約、作廢的項目不出現');
  ok(pb.maReminders.some((x) => x.itemId === 'iE' && x.days === -20 && x.level === 'danger' && /已逾期 20 天/.test(x.title)), 'C5e 逾期 20 天的 iE：到期後規則（-1 起每 7 天）持續顯示');
  ok(pb.contractReminders.some((c) => c.contractId === 'cEM') && !pb.contractReminders.some((c) => c.contractId === 'cSM'), 'C5f 開啟時：舊的合約到期提醒排除 SERVICE_MA、ERP_MA 照舊');
  { const cs = readData().contracts; ok(!!cs.find((c) => c.id === 'cEM')._contract30Reminded && !cs.find((c) => c.id === 'cSM')._contract30Reminded, 'C5g checkOverdueReminders（啟動時執行）：SERVICE_MA 合約不被處理（沒有 _contract30Reminded 標記）、ERP_MA 照處理。註：本機 JSON 模式下該函式會用過期快照覆蓋自己推的通知（既有行為，Postgres 共用快取不受影響），所以以標記判斷', J(cs.map((c) => [c.id, c._contract30Reminded]))); }
  const et = pb && (await http(U.u1, 'GET', '/api/poll-bundle')).headers.get('etag');
  const nm = await http(U.u1, 'GET', '/api/poll-bundle', undefined, { 'If-None-Match': et });
  eq([!!et, nm.status], [true, 304], 'C5h ETag：沒變就 304');
  // ═════ D. 通知自動解除（用 API 建立續約報價單）═════
  r = await http(U.u1, 'POST', '/api/quotations', qBody('oQ1', { renewOf: 'iA' }));
  ok(r.status === 201 && r.json.renewOf === 'iA', 'D1 業務建立續約報價單（renewOf＝iA）→ 201，回應帶 renewOf', r.text.slice(0, 300));
  const qid = r.json && r.json.id;
  lg = await auditLines('CREATE_QUOTATION');
  ok(lg.some((x) => /續約對象納管項 iA/.test(x.detail)), 'D1b 稽核 CREATE_QUOTATION 記錄續約對象');
  ok(readData().quotations.find((q) => q.id === qid).renewOf === 'iA', 'D1c 資料檔有 renewOf');
  ok(!('contentHash' in (r.json || {})) || true, 'D1d');
  pb = await poll(U.u1);
  ok(pb.maReminders.some((x) => x.itemId === 'iA') && pb.maReminders.find((x) => x.itemId === 'iA').quoteState === 'draft', 'D2 建立草稿後：預設規則的 D-90（always）仍在，但狀態顯示「草稿」（always 不因續約報價單而消失，設計如此）');
  // ═════ G. renewOf 驗證（POST／PUT）═════
  const QA = require(path.join(ROOT, 'lib', 'quoteApproval.js'));
  bad = await loginMany([U.tp, U.gs]);
  eq(bad, [], 'G0 登入集團角色');
  const qCount = () => (readData().quotations || []).length;
  const q0 = qCount();
  const post = (who, opp, ro) => http(who, 'POST', '/api/quotations', qBody(opp, ro === undefined ? {} : { renewOf: ro }));
  r = await post(U.u1, 'oQ1', 'iZ');
  eq([r.status, r.json.code], [400, 'RENEW_COMPANY_MISMATCH'], 'G1 續約對象與報價單不同客戶 → 400 RENEW_COMPANY_MISMATCH');
  r = await post(U.u1, 'oQ1', 'iB');
  eq([r.status, r.json.code], [404, 'RENEW_ITEM_NOT_FOUND'], 'G2 他人的納管項（不在可視範圍）→ 404 RENEW_ITEM_NOT_FOUND');
  r = await post(U.u1, 'oQ1', 'nope');
  eq([r.status, r.json.code], [404, 'RENEW_ITEM_NOT_FOUND'], 'G3 不存在 → 同樣的 404（不洩漏他人項目是否存在）');
  ok(r.text === (await post(U.u1, 'oQ1', 'iB')).text, 'G3b 看不到與不存在的訊息逐字相同');
  r = await post(U.u1, 'oQ1', 'iV');
  eq([r.status, r.json.code], [409, 'RENEW_ITEM_VOID'], 'G4 已作廢的納管項 → 409 RENEW_ITEM_VOID');
  r = await post(U.u1, 'oQ1', 123);
  eq([r.status, r.json.code], [400, 'BAD_RENEW_OF'], 'G5 非字串 → 400 BAD_RENEW_OF');
  r = await post(U.u1, 'oQ1', 'x'.repeat(81));
  eq([r.status, r.json.code], [400, 'BAD_RENEW_OF'], 'G5b 過長 → 400');
  eq(qCount(), q0, 'G5c 所有被拒絕的 POST 都沒有建立報價單（也沒有消耗單號）');
  r = await post(U.u1, 'oQ1', '');
  ok(r.status === 201 && !('renewOf' in r.json), 'G6 空字串 → 允許，不存 renewOf');
  r = await post(U.u1, 'oQ1', null);
  ok(r.status === 201 && !('renewOf' in r.json), 'G6b null → 允許，不存 renewOf');
  r = await post(U.u1, 'oQ1', 'iF');
  ok(r.status === 201 && r.json.renewOf === 'iF', 'G7 已續約（renewed）狀態的項目仍可指定（只擋作廢）');
  r = await post(U.u2, 'oQ2', 'iB');
  ok(r.status === 201 && r.json.renewOf === 'iB', 'G8 業務乙續自己的項目 → 201');
  r = await post(U.u2, 'oQ2', 'iA');
  eq([r.status, r.json.code], [404, 'RENEW_ITEM_NOT_FOUND'], 'G8b 業務乙續業務甲的項目 → 404');
  r = await http(U.u1, 'GET', '/api/quotations/' + qid);
  ok(r.status === 200 && r.json.renewOf === 'iA', 'G9 單張 GET 帶 renewOf');
  const lst = await http(U.u1, 'GET', '/api/quotations');
  ok(lst.status === 200, 'G9b 列表照常');
  // PUT
  r = await http(U.u1, 'PUT', '/api/quotations/' + qid, { renewOf: '' });
  ok(r.status === 200 && !('renewOf' in r.json) && !('renewOf' in readData().quotations.find((q) => q.id === qid)), 'G10 PUT 空字串＝清除');
  lg = await auditLines('UPDATE_QUOTATION');
  ok(lg.some((x) => /續約對象 iA → （無）/.test(x.detail)), 'G10b 稽核記錄 iA → （無）');
  r = await http(U.u1, 'PUT', '/api/quotations/' + qid, { renewOf: 'iA' });
  ok(r.status === 200 && r.json.renewOf === 'iA', 'G11 PUT 設定回 iA');
  const nAud = (await auditLines('UPDATE_QUOTATION')).length;
  r = await http(U.u1, 'PUT', '/api/quotations/' + qid, { renewOf: 'iA' });
  ok(r.status === 200 && (await auditLines('UPDATE_QUOTATION')).every((x) => !/續約對象 iA → iA/.test(x.detail)), 'G11b 送回相同值：不重新驗證、不記續約對象異動');
  r = await http(U.u1, 'PUT', '/api/quotations/' + qid, { renewOf: 'iB' });
  eq([r.status, r.json.code, readData().quotations.find((q) => q.id === qid).renewOf], [404, 'RENEW_ITEM_NOT_FOUND', 'iA'], 'G12 PUT 改成別人的項目 → 404，原值不變');
  r = await http(U.u1, 'PUT', '/api/quotations/' + qid, { renewOf: 'iZ' });
  eq([r.status, r.json.code], [400, 'RENEW_COMPANY_MISMATCH'], 'G12b PUT 不同客戶 → 400');
  r = await http(U.u2, 'PUT', '/api/quotations/' + qid, { renewOf: 'iB' });
  eq(r.status, 403, 'G12c 非擁有者不能改');
  // 核准後補 renewOf：不進雜湊 → 不作廢核准
  r = await post(U.u1, 'oQ1', 'iE');
  const qApproved = r.json.id;
  writeData((d) => { const q = d.quotations.find((x) => x.id === qApproved); q.status = 'approved'; q.approval = { state: 'approved', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), steps: [{ status: 'approved', at: nowIso, label: '核決', by: U.m1 }], cur: 1, history: [] }; });
  const hashBefore = readData().quotations.find((x) => x.id === qApproved).approval.hash;
  r = await http(U.u1, 'PUT', '/api/quotations/' + qApproved, { renewOf: 'iD' });
  let qq = readData().quotations.find((x) => x.id === qApproved);
  ok(r.status === 200 && qq.renewOf === 'iD' && qq.approval.state === 'approved' && QA.contentHash(qq) === hashBefore && qq.approval.hash === hashBefore, 'G13 已核准的單改 renewOf：核准不作廢、雜湊逐位元不變（renewOf 不進核准雜湊）', r.text.slice(0, 200));
  ok(QA.contentHash(Object.assign({}, qq, { renewOf: 'whatever' })) === QA.contentHash(qq), 'G13b contentHash 與 renewOf 無關（單元）');
  // ═════ H. 時間軸、續約候選 ═════
  r = await http(U.u1, 'GET', '/api/ma-items/iA/reminders');
  ok(r.status === 200 && r.json.item.id === 'iA' && r.json.item.daysLeft === 40 && r.json.rules.length === 4 && r.json.isDefault === true && r.json.quoteState === 'draft', 'H1 時間軸：項目摘要、規則 4 條、報價單狀態', r.text.slice(0, 300));
  const rowD90 = r.json.rules.find((x) => x.ruleId === 'default-d90'), rowD45 = r.json.rules.find((x) => x.ruleId === 'default-d45');
  ok(rowD90.triggerDate === plusDays(T0, 40 - 90) && rowD90.reached && rowD90.conditionNow && rowD90.activeNow && rowD90.sentCount === 3 && rowD90.sent.length === 3 && rowD90.sent.every((s) => s.toLabel && s.at && s.period === 0), 'H1b D-90：觸發日＝迄日−90、已發送 3 筆（業務＋2 位秘書）含顯示名與時間', J(rowD90));
  ok(rowD45.reached && rowD45.conditionNow === false && rowD45.activeNow === false && rowD45.reason === 'CONDITION' && rowD45.sentCount === 1, 'H1c D-45：已有續約報價單 → 條件現在不成立；但先前（建立報價單前）已發過 1 筆', J(rowD45));
  ok(clean(r) && !r.text.includes('contractNo'), 'H1d 時間軸沒有金額與合約編號');
  r = await http(U.sec, 'GET', '/api/ma-items/iA/reminders');
  ok(r.status === 200 && !r.json.item.id && !('notRenewReason' in r.json.item) && r.json.rules.every((x) => !('sent' in x)) && clean(r) && Object.keys(r.json.item).sort().join() === 'company,daysLeft,from,ownerLabel,product,status,termKind,to', '秘書 H2：同 BU 可看，但項目沒有 id、規則列沒有收件人名單、沒有金額與合約編號', r.text.slice(0, 300));
  eq([(await http(U.u2, 'GET', '/api/ma-items/iA/reminders')).status, (await http(U.m1b, 'GET', '/api/ma-items/iA/reminders')).status, (await http(U.u1, 'GET', '/api/ma-items/nope/reminders')).status, (await http(U.gs, 'GET', '/api/ma-items/iA/reminders')).status, (await http(U.tp, 'GET', '/api/ma-items/iA/reminders')).status, (await http(null, 'GET', '/api/ma-items/iA/reminders')).status],
    [404, 404, 404, 404, 403, 401], 'H3 看不到的（他人／別條線／不存在／集團業務範圍外）→ 404；集團 PM 403；未登入 401');
  eq([(await http(U.mkt, 'GET', '/api/ma-items/iA/reminders')).status, (await http(U.m1, 'GET', '/api/ma-items/iA/reminders')).status, (await http(U.adm, 'GET', '/api/ma-items/iA/reminders')).status], [200, 200, 200], 'H3b 行銷（看全部）、一級主管（自己線）、管理員可看');
  // 候選
  const CAND = (who, q) => http(who, 'GET', '/api/ma-items/renewal-candidates' + (q === undefined ? '' : '?company=' + encodeURIComponent(q)));
  r = await CAND(U.u1, 'CoA');
  eq([r.status, r.json.items.map((x) => x.id), r.json.total], [200, ['iE', 'iA'], 2], 'H4 候選：CoA 的追蹤中且 ≤90 天（含逾期）：iE(-20)、iA(40)；排除 iD(100 天)、iF(已續約)、iV(作廢)、其他客戶');
  ok(r.json.items[1].hasRenewalQuote === true && r.json.items[1].quoteState === 'draft' && r.json.items[0].hasRenewalQuote === false && r.json.items[0].daysLeft === -20 && Object.keys(r.json.items[0]).sort().join() === 'daysLeft,from,hasRenewalQuote,id,ownerLabel,product,quoteState,termKind,to' && clean(r), 'H4b 候選欄位白名單、是否已有續約報價單、無金額');
  eq([(await CAND(U.u1, '  CoA  ')).json.total, (await CAND(U.u1, 'coa')).json.total, (await CAND(U.u2, 'CoA')).json.total, (await CAND(U.sec, 'CoA')).json.total, (await CAND(U.mkt, 'CoA')).json.total], [2, 0, 0, 0, 2], 'H5 前後空白 trim、區分大小寫（與商機公司比對同規則）；他人看不到；秘書回空；行銷看全部');
  eq([(await CAND(U.u1)).json.code, (await CAND(U.u1, '')).status, (await CAND(null, 'CoA')).status, (await CAND(U.tp, 'CoA')).status], ['COMPANY_REQUIRED', 400, 401, 403], 'H6 缺 company → 400；未登入 401；集團 PM 403');
  // ═════ I. 不續約／撤銷 ═════
  const NR = (who, id, reason) => http(who, 'PATCH', `/api/ma-items/${id}/not-renew`, reason === undefined ? {} : { reason });
  eq([(await NR(U.u2, 'iA', '客戶改用其他方案')).status, (await NR(U.m1, 'iA', '客戶改用其他方案')).status, (await NR(U.sec, 'iA', '客戶改用其他方案')).status, (await NR(U.mkt, 'iA', '客戶改用其他方案')).status, (await NR(null, 'iA', '客戶改用其他方案')).status], [403, 403, 403, 403, 401], 'I1 只有負責業務與管理員能標不續約（他人／主管／秘書／行銷 403）');
  eq([(await NR(U.u1, 'iA', '')).json.code, (await NR(U.u1, 'iA', '一')).json.code, (await NR(U.u1, 'iA', '字'.repeat(201))).json.code, (await NR(U.u1, 'iA', 'a\u0007b')).json.code, (await NR(U.u1, 'iA')).json.code, (await NR(U.u1, 'iA', 5)).json.code],
    ['NOT_RENEW_REASON_REQUIRED', 'NOT_RENEW_REASON_REQUIRED', 'NOT_RENEW_REASON_TOO_LONG', 'NOT_RENEW_REASON_INVALID', 'NOT_RENEW_REASON_REQUIRED', 'NOT_RENEW_REASON_REQUIRED'], 'I2 原因必填：空／1 字／過長／控制字元 → 對應錯誤碼');
  eq(readData().maItems.find((x) => x.id === 'iA').status, 'tracking', 'I2b 被拒絕的請求沒有改狀態');
  eq([(await NR(U.u1, 'nope', '客戶改用其他方案')).status, (await NR(U.u1, 'iV', '客戶改用其他方案')).json.code, (await NR(U.u1, 'iF', '客戶改用其他方案')).json.code], [404, 'ITEM_NOT_TRACKING', 'ITEM_NOT_TRACKING'], 'I3 不存在 404；作廢／已續約的項目 → 409 ITEM_NOT_TRACKING');
  const nBefore = (readData().notifications || []).filter((n) => n.type === 'ma_not_renew').length;
  r = await NR(U.u1, 'iA', '  客戶改用其他方案  ');
  ok(r.status === 200 && r.json.item.status === 'not_renew' && r.json.notRenewReason === '客戶改用其他方案' && r.json.notified === 3 && !r.text.includes(AMT) && !r.text.includes(NOTE), 'I4 負責業務標不續約 → 200；通知 3 人（一級主管＋2 位在職行銷，不含自己與離職行銷）', r.text.slice(0, 300));
  dd = readData();
  const it = dd.maItems.find((x) => x.id === 'iA');
  ok(it.status === 'not_renew' && it.notRenewReason === '客戶改用其他方案' && it.history.slice(-1)[0].action === 'NOT_RENEW', 'I4b 資料：狀態、原因（已 trim）、history');
  const nn = (dd.notifications || []).filter((n) => n.type === 'ma_not_renew');
  ok(nn.length === nBefore + 3 && ['_s_m1', U.mkt, U.mkt2].every((t) => nn.some((n) => n.to === t && n.refId === 'iA')) && !nn.some((n) => n.to === U.u1 || n.to === U.mktoff) && nn.every((n) => !FORBIDDEN_TOKENS.some((t) => (n.title + n.body).includes(t))) && /客戶改用其他方案/.test(nn[0].body), 'I4c ma_not_renew 通知：一級主管＋行銷，內文有原因、沒有金額');
  lg = await auditLines('MA_ITEM_NOT_RENEW');
  ok(lg.length === 1 && /原因：客戶改用其他方案/.test(lg[0].detail), 'I4d 稽核 MA_ITEM_NOT_RENEW');
  eq((await NR(U.u1, 'iA', '客戶改用其他方案')).json.code, 'ITEM_NOT_TRACKING', 'I5 重複標記 → 409');
  pb = await poll(U.u1);
  ok(!pb.maReminders.some((x) => x.itemId === 'iA'), 'I6 標不續約後：動態鈴鐺立刻沒有 iA');
  { const row = (await http(U.u1, 'GET', '/api/ma-items')).json.items.find((x) => x.id === 'iA'); const sr = (await http(U.sec, 'GET', '/api/ma-items')).json.items; ok(row.status === 'not_renew' && row.notRenewReason === '客戶改用其他方案' && sr.every((x) => !('notRenewReason' in x) && !('status' in x)) && !('notRenewReason' in (await http(U.u1, 'GET', '/api/ma-items')).json.items.find((x) => x.id === 'iD')), 'I6b 到期管理清單：not_renew 的列帶 notRenewReason（其他狀態與秘書列沒有這個鍵）'); }
  r = await NR(U.adm, 'iB', '業務離職，客戶不再聯絡');
  ok(r.status === 200 && r.json.notified === 3 && (readData().notifications || []).filter((n) => n.type === 'ma_not_renew' && n.refId === 'iB').every((n) => n.to !== U.adm), 'I7 管理員可代標；操作者本人不收通知');
  eq([(await http(U.u1, 'PATCH', '/api/ma-items/iA/reopen')).status, (await http(U.m1, 'PATCH', '/api/ma-items/iA/reopen')).status, (await http(U.sec, 'PATCH', '/api/ma-items/iA/reopen')).status, (await http(null, 'PATCH', '/api/ma-items/iA/reopen')).status], [403, 403, 403, 401], 'I8 撤銷只有管理員（連負責業務都不行）');
  eq([(await http(U.adm, 'PATCH', '/api/ma-items/nope/reopen')).status, (await http(U.adm, 'PATCH', '/api/ma-items/iD/reopen')).json.code], [404, 'ITEM_NOT_NOT_RENEW'], 'I8b 不存在 404；追蹤中的項目撤銷 → 409 ITEM_NOT_NOT_RENEW');
  r = await http(U.adm, 'PATCH', '/api/ma-items/iA/reopen');
  ok(r.status === 200 && r.json.item.status === 'tracking' && !('notRenewReason' in readData().maItems.find((x) => x.id === 'iA')) && readData().maItems.find((x) => x.id === 'iA').history.slice(-1)[0].action === 'NOT_RENEW_REVOKED', 'I9 管理員撤銷 → 回追蹤中、清除原因、history 保留舊原因');
  lg = await auditLines('MA_ITEM_NOT_RENEW_REVOKED');
  ok(lg.length === 1 && /客戶改用其他方案/.test(lg[0].detail), 'I9b 稽核 MA_ITEM_NOT_RENEW_REVOKED 記原因');
  eq((await http(U.adm, 'PATCH', '/api/ma-items/iA/reopen')).json.code, 'ITEM_NOT_NOT_RENEW', 'I9c 重複撤銷 409');
  pb = await poll(U.u1);
  ok(pb.maReminders.some((x) => x.itemId === 'iA'), 'I10 撤銷後動態鈴鐺回來');
  const total2 = (readData().notifications || []).filter((n) => n.type === 'ma_reminder').length;
  cr = await cron();
  ok(cr.json.reminders.fired === 0 && (readData().notifications || []).filter((n) => n.type === 'ma_reminder').length === total2, 'I10b 撤銷後的 cron：去重鍵還在，不重發舊提醒');
  // ═════ L. poll-bundle 補掃（沒有 cron 的環境，例如 Hobby 方案 cron 沒跑）═════
  await freshServer({ data: apiSeed });
  bad = await loginMany([U.u1, U.u2, U.adm]);
  eq([bad, readData().reminderLog], [[], undefined], 'L0 全新資料：沒有 reminderLog');
  pb = await poll(U.u1);
  dd = readData();
  ok(dd.reminderLog && dd.reminderLog.lastSweepMode === 'lazy' && dd.reminderLog.lastSweepDate === todayTp() && (dd.notifications || []).some((n) => n.type === 'ma_reminder' && n.to === U.u1), 'L1 今天還沒掃過 → 第一個登入者的輪詢觸發補掃（lazy），當下就發出通知', J(dd.reminderLog && dd.reminderLog.lastStats));
  ok(pb.notifications.some((n) => n.type === 'ma_reminder') && pb.maReminders.length > 0, 'L1b 同一次輪詢的回應就帶新通知與 maReminders（補掃在載入資料之前完成）');
  const at1 = dd.reminderLog.lastSweepAt, n1 = (dd.notifications || []).length;
  await sleep(1100);
  await poll(U.u2); await poll(U.u1); await poll(U.adm);
  dd = readData();
  ok(dd.reminderLog.lastSweepAt === at1 && (dd.notifications || []).filter((n) => n.type === 'ma_reminder').length >= 1 && (dd.notifications || []).length >= n1, 'L2 今天已掃過、60 分鐘內：後續輪詢不再掃（lastSweepAt 不變）');
  cr = await cron();
  ok(cr.json.reminders.fired === 0 && cr.json.reminders.skippedDup > 0, 'L3 補掃之後 cron 再掃：去重鍵共用，不重發（兩個觸發源冪等）');
  lg = await auditLines('MA_REMINDER_SWEEP');
  ok(lg.some((x) => /模式 lazy/.test(x.detail)) && lg.some((x) => /模式 cron/.test(x.detail)), 'L4 稽核 MA_REMINDER_SWEEP 分別記 lazy 與 cron 的統計');
  // 持久化的 lastSweepDate 是昨天（模擬跨日）→ 立刻補掃
  const atBefore = readData().reminderLog.lastSweepAt;
  writeData((d) => { d.reminderLog.lastSweepDate = plusDays(todayTp(), -1); });
  await poll(U.u1);
  eq(readData().reminderLog.lastSweepAt, atBefore, 'L5 同一個實例 60 秒內剛試過 → 即使日期過期也先不重試（節流）');
  await restartServer(); await loginMany([U.u1]);   // 新實例（冷啟動）沒有嘗試紀錄
  await poll(U.u1);
  ok(readData().reminderLog.lastSweepDate === todayTp() && readData().reminderLog.lastSweepAt !== atBefore, 'L5b 持久化的掃描日期不是今天（跨日）＋新實例 → 輪詢立刻補掃');
  { const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8'); ok(/case 'ma_reminder':\s*case 'ma_not_renew':\s*return '\/index\.html#maitems'/.test(src), 'L6 notificationUrlFor：ma_reminder／ma_not_renew → /index.html#maitems'); }
}
async function apiFlow() {
  const T0 = todayTp();
  let bad, r, lg, cr, dd;
  // ═════ E/F. 條件式規則 × 真實報價單／成交：通知自動解除、續約成交掛鉤 ═════
  await freshServer({ data: (d) => {
    d.opportunities = [mkOpp('oQ1', U.u1, MAP, 'C', { company: 'CoA' }), mkOpp('oR', U.u1, MAP, 'C', { company: 'CoA' })];
    d.maItems = [mkItem('iA', U.u1, 40, { company: 'CoA' }), mkItem('iB', U.u2, 40, { company: 'CoB' })];
    d.reminderRules = { rules: COND, updatedAt: nowIso, updatedBy: U.adm };
  } });
  bad = await loginMany([U.adm, U.u1, U.u2, U.m1, U.mkt]);
  eq(bad, [], 'E0 登入');
  cr = await cron();
  eq(cr.json.reminders.fired > 0, true, 'E1 條件式規則（無報價@45、未送簽@45、未續約@-1）：iA／iB 無續約報價單 → c45 發給 owner＋一級主管');
  dd = readData();
  const unread = (u) => (readData().notifications || []).filter((n) => n.to === u && n.type === 'ma_reminder' && !n.read).length;
  eq([unread(U.u1), unread(U.u2), unread(U.m1), unread(U.mkt)], [1, 1, 1, 0], 'E1b 業務甲 1、業務乙 1、一級主管 1（彙總 iA＋iB）、行銷 0');
  eq((await poll(U.m1)).maReminders.map((x) => x.itemId).sort(), ['iA', 'iB'], 'E1c 一級主管的鈴鐺有 iA、iB');
  // 建立續約報價單（草稿）
  r = await http(U.u1, 'POST', '/api/quotations', qBody('oQ1', { renewOf: 'iA' }));
  const q2 = r.json.id;
  eq(r.status, 201, 'E2 業務甲建立 iA 的續約報價單（草稿）');
  let pu = await poll(U.u1);
  ok(pu.maReminders.some((x) => x.itemId === 'iA' && x.quoteState === 'draft'), 'E2b 草稿 → k45（無報價）不成立，但 kq（草稿未送簽）成立：鈴鐺仍在、狀態草稿');
  cr = await cron();
  eq([unread(U.u1), unread(U.m1), unread(U.mkt)], [2, 2, 1], 'E2c 下次 cron：kq 對 iA 首次成立 → 業務甲、一級主管、行銷各多一則（舊的 c45 通知仍因 kq 成立而未讀）');
  // 送簽
  writeData((d) => { const q = d.quotations.find((x) => x.id === q2); q.approval = { state: 'pending', steps: [], cur: 0, history: [] }; });
  pu = await poll(U.u1);
  eq([pu.maReminders.some((x) => x.itemId === 'iA'), unread(U.u1)], [false, 0], 'E3 送簽後：業務甲的鈴鐺沒有 iA，且他自己的兩則通知在這次 poll 就被標已讀');
  let pm = await poll(U.m1);
  eq([pm.maReminders.map((x) => x.itemId), unread(U.m1)], [['iB'], 1], 'E3b 一級主管的鈴鐺只剩 iB；c45 那則（iA＋iB）因 iB 仍成立保持未讀、kq 那則（只有 iA）已標已讀');
  eq(unread(U.mkt), 1, 'E3d 行銷還沒 poll → 通知暫時未讀');
  await poll(U.mkt);
  eq(unread(U.mkt), 0, 'E3e 行銷 poll 後自動標已讀');
  // 業務乙標 iB 不續約 → 一級主管的鈴鐺清空、通知標已讀
  r = await NR(U.u2, 'iB', '預算不足，明年再議');
  eq(r.status, 200, 'E4 業務乙標 iB 不續約');
  pm = await poll(U.m1);
  eq([pm.maReminders.length, unread(U.m1), unread(U.u2)], [0, 0, 0], 'E4b 一級主管、業務乙的鈴鐺與未讀通知全部清空');
  const total3 = (readData().notifications || []).length;
  cr = await cron();
  eq([cr.json.reminders.fired, (readData().notifications || []).length], [0, total3], 'E4c 之後的 cron 不再發');
  // 續約成交
  writeData((d) => { const q = d.quotations.find((x) => x.id === q2); q.status = 'approved'; q.approval = { state: 'approved', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), steps: [{ status: 'approved', at: nowIso, label: '核決' }], cur: 1, history: [] }; });
  const itemsBefore = readData().maItems.length;
  // renewal 商機 oR 也要連到已核准報價單
  r = await http(U.u1, 'POST', '/api/quotations', qBody('oR', { renewOf: 'iA' }));
  const q3 = r.json.id;
  writeData((d) => { const q = d.quotations.find((x) => x.id === q3); q.status = 'approved'; q.approval = { state: 'approved', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), steps: [{ status: 'approved', at: nowIso, label: '核決' }], cur: 1, history: [] }; });
  r = await http(U.u1, 'PUT', '/api/opportunities/oR', { stage: 'Won', termFrom: plusDays(T0, 41), termTo: plusDays(T0, 405) });
  ok(r.status === 200 && r.json.ma && r.json.ma.createdItems === 1, 'F1 續約商機成交 → 200，建立新納管項', r.text.slice(0, 300));
  dd = readData();
  const oldI = dd.maItems.find((x) => x.id === 'iA');
  ok(oldI.status === 'renewed' && oldI.renewedBy.oppId === 'oR' && oldI.renewedBy.quoteId === q3 && oldI.history.slice(-1)[0].action === 'RENEWED' && dd.maItems.length === itemsBefore + 1, 'F1b 舊納管項 iA → renewed（renewedBy 記新商機與報價單）、history RENEWED', J(oldI));
  lg = await auditLines('MA_ITEM_RENEWED');
  ok(lg.length === 1 && /新商機 oR/.test(lg[0].detail), 'F1c 稽核 MA_ITEM_RENEWED');
  pu = await poll(U.u1);
  ok(!pu.maReminders.some((x) => x.itemId === 'iA'), 'F2 已續約 → 鈴鐺沒有 iA');
  // 回退
  r = await http(U.u1, 'PUT', '/api/opportunities/oR', { stage: 'B' });
  dd = readData();
  ok(r.status === 200 && dd.maItems.find((x) => x.id === 'iA').status === 'tracking' && !dd.maItems.find((x) => x.id === 'iA').renewedBy && dd.maItems.find((x) => x.id === 'iA').history.slice(-1)[0].action === 'RENEW_REVERTED', 'F3 續約商機退回非成交 → 舊項恢復追蹤（含 history）');
  lg = await auditLines('MA_ITEM_RENEW_REVERTED');
  eq(lg.length, 1, 'F3b 稽核 MA_ITEM_RENEW_REVERTED');
  r = await http(U.u1, 'PUT', '/api/opportunities/oR', { stage: 'Won' });
  dd = readData();
  ok(r.status === 200 && dd.maItems.find((x) => x.id === 'iA').status === 'renewed' && dd.maItems.filter((x) => x.oppId === 'oR' && x.status === 'tracking').length === 1, 'F4 再次成交（沿用先前期間）→ 再次標 renewed、不重複建新項（冪等）', r.text.slice(0, 200));
  r = await http(U.u1, 'PUT', '/api/opportunities/oR', { stage: 'Won', termFrom: plusDays(T0, 41), termTo: plusDays(T0, 405) });
  eq(readData().maItems.find((x) => x.id === 'iA').history.filter((h) => h.action === 'RENEWED').length, 2, 'F4b 重複送出 Won 不再多記一筆（第一次成交、回退後再成交，共 2 筆 RENEWED）');
  // 不是已核准的單 → 不標
  await freshServer({ data: (d) => {
    d.opportunities = [mkOpp('oR', U.u1, MAP, 'C', { company: 'CoA' })];
    d.maItems = [mkItem('iA', U.u1, 40, { company: 'CoA' })];
    d.quotations = [mkQuote('qd', 'oR', U.u1, 'CoA', [], null, { renewOf: 'iA' }), mkQuote('qp', 'oR', U.u1, 'CoA', [], 'pending', { renewOf: 'iA' }), mkQuote('qi', 'oR', U.u1, 'CoA', [], 'approved', { renewOf: 'iA', invalid: { at: nowIso, by: 'x', reason: 'OPP_VOID' } })];
  } });
  await loginMany([U.u1]);
  r = await http(U.u1, 'PUT', '/api/opportunities/oR', { stage: 'Won', termFrom: plusDays(T0, 41), termTo: plusDays(T0, 405) });
  eq([r.status, readData().maItems.find((x) => x.id === 'iA').status], [200, 'tracking'], 'F5 連結的報價單是草稿／簽核中／已失效 → 不標舊項為已續約');
  await freshServer({ data: (d) => {
    d.opportunities = [mkOpp('oR', U.u1, MAP, 'C', { company: 'CoA' })];
    d.maItems = [mkItem('iA', U.u1, 40, { company: 'CoA', status: 'not_renew' }), mkItem('iB', U.u1, 40, { company: 'CoA', status: 'ended' })];
    d.quotations = [mkQuote('q1', 'oR', U.u1, 'CoA', [], 'approved', { renewOf: 'iA' }), mkQuote('q2', 'oR', U.u1, 'CoA', [], 'approved', { renewOf: 'iB' }), mkQuote('q3', 'oR', U.u1, 'CoA', [], 'approved', { renewOf: 'ghost' })];
  } });
  await loginMany([U.u1]);
  r = await http(U.u1, 'PUT', '/api/opportunities/oR', { stage: 'Won', termFrom: plusDays(T0, 41), termTo: plusDays(T0, 405) });
  eq([r.status, readData().maItems.find((x) => x.id === 'iA').status, readData().maItems.find((x) => x.id === 'iB').status], [200, 'not_renew', 'ended'], 'F6 舊項已是 not_renew／ended／不存在 → 不動、不丟例外');
}
async function apiDual() {
  let bad, r, lg;
  // ═════ J. 雙人覆核 ═════
  await freshServer({ data: apiSeed }, { APPROVAL_CONFIG_DUAL_REVIEW: '1' });
  bad = await loginMany([U.adm, U.adm2]);
  eq(bad, [], 'J0 登入');
  const dr = [R1({ id: 'dr1', name: '覆核規則', whenDays: 12, notify: ['owner'] })];
  r = await http(U.adm, 'PUT', '/api/admin/reminder-rules', { rules: dr, reviewNote: '測試覆核' });
  ok(r.status === 202 && r.json.status === 'pending' && r.json.requestId, 'J1 開啟雙人覆核：PUT → 202 待覆核（與防拆單設定一致）', r.text.slice(0, 200));
  const rid = r.json.requestId;
  eq((await http(U.adm, 'GET', '/api/admin/reminder-rules')).json.config.isDefault, true, 'J1b 覆核前設定不變');
  r = await http(U.adm, 'POST', `/api/admin/approval-config-requests/${rid}/approve`, {});
  eq([r.status, r.json.code], [403, 'SELF_REVIEW_FORBIDDEN'], 'J2 自己不能核可自己的提案');
  r = await http(U.adm2, 'GET', '/api/admin/approval-config-requests');
  ok(r.status === 200 && JSON.stringify(r.json).includes('到期提醒規則'), 'J2b 覆核清單顯示種類「到期提醒規則」與差異', r.text.slice(0, 200));
  r = await http(U.adm2, 'POST', `/api/admin/approval-config-requests/${rid}/approve`, { comment: 'ok' });
  eq([r.status, r.json.status], [200, 'approved'], 'J3 第二位管理員核可');
  r = await http(U.adm, 'GET', '/api/admin/reminder-rules');
  ok(r.json.config.isDefault === false && r.json.config.rules.length === 1 && r.json.config.rules[0].id === 'dr1' && r.json.config.updatedBy === U.adm, 'J3b 核可後生效；updatedBy＝提案人');
  lg = await auditLines('SAVE_REMINDER_RULES');
  ok(lg.length === 1 && /提案人 _s_adm/.test(lg[0].detail) && /覆核人 _s_adm2/.test(lg[0].detail), 'J3c 稽核 SAVE_REMINDER_RULES 帶提案人、覆核人');
  r = await http(U.adm, 'PUT', '/api/admin/reminder-rules', { rules: dr });
  ok(r.status === 200 && r.json.noChange === true, 'J4 提案內容與現行相同 → noChange（不需覆核）', r.text.slice(0, 200));
  r = await http(U.adm, 'PUT', '/api/admin/reminder-rules', { rules: [R1({ whenDays: 999 })] });
  eq([r.status, r.json.code], [400, 'BAD_RULE_WHEN'], 'J5 不合法的內容在提案前就被擋（400）');
  r = await http(U.adm, 'PUT', '/api/admin/reminder-rules', { rules: [R1({ id: 'dr2', whenDays: 3 })] });
  const rid2 = r.json.requestId;
  r = await http(U.adm2, 'POST', `/api/admin/approval-config-requests/${rid2}/reject`, { comment: '不同意' });
  eq([r.status, (await http(U.adm, 'GET', '/api/admin/reminder-rules')).json.config.rules[0].id], [200, 'dr1'], 'J6 退回 → 設定不變');
}
async function apiFail() {
  let bad, r, lg;
  // ═════ K. 失敗不影響 mail cron ═════
  const injectSpec = [{ file: 'lib/maReminderRoutes.js', edits: [{ find: 'const r = MRM.sweep(', rep: "const r = (() => { throw new Error('injected sweep failure'); })() || MRM.sweep(" }] }];
  for (const mailMode of [null, 'log']) {
    const envBase = mailMode ? { MAIL_MODE: mailMode } : {};
    // 基準：MA_TRACKING 未設
    await stopServer(); cleanupFiles(); seedFiles({ data: apiSeed }); await startServer(null, envBase); jar = {};
    const base = await cron();
    // 開啟、沒有注入失敗
    await stopServer(); cleanupFiles(); seedFiles({ data: apiSeed }); await startServer(MUT_CUR, Object.assign({ MA_TRACKING: '1' }, envBase)); jar = {};
    const withMa = await cron();
    // 開啟、掃描丟例外
    await stopServer(); cleanupFiles(); seedFiles({ data: apiSeed }); await startServer((MUT_CUR || []).concat(injectSpec), Object.assign({ MA_TRACKING: '1' }, envBase)); jar = {};
    const failing = await cron();
    const tag = mailMode ? 'MAIL_MODE=log' : 'MAIL off';
    ok(base.status === 200 && base.json && !('reminders' in base.json), `K1 [${tag}] 基準（MA_TRACKING 未設）：200，沒有 reminders`, base.text.slice(0, 200));
    ok(withMa.status === 200 && withMa.json.reminders && withMa.json.reminders.fired > 0, `K2 [${tag}] 開啟：200，多一個 reminders`, withMa.text.slice(0, 200));
    const strip = (x) => { const o = Object.assign({}, x); delete o.reminders; return o; };
    ok(J(strip(withMa.json)) === J(base.json), `K3 [${tag}] 開啟時，除了 reminders 之外與基準逐欄相同`, J(strip(withMa.json)) + ' vs ' + J(base.json));
    ok(failing.status === 200 && J(failing.json) === J(base.json), `K4 [${tag}] 掃描丟例外：狀態碼與回應本文與基準完全相同（不影響信件佇列處理與回應）`, failing.text.slice(0, 200));
    ok(/到期提醒掃描失敗/.test(child.getOut()), `K5 [${tag}] 例外只記在 console`);
    await loginMany([U.u1]);
    const pf = await http(U.u1, 'GET', '/api/poll-bundle');
    ok(pf.status === 200 && Array.isArray(pf.json.maReminders) && pf.json.notifications !== undefined, `K6 [${tag}] 補掃失敗時 poll-bundle 仍 200、動態鈴鐺照常（動態推導與掃描無關）`, pf.text.slice(0, 200));
    ok(readData().reminderLog === undefined || Object.keys(readData().reminderLog.entries || {}).length === 0, `K7 [${tag}] 掃描失敗不留下半套去重標記`);
  }
  await stopServer(); cleanupFiles();
}

// ═════════════ 角色權限一致性 ═════════════
/** 獨立於 lib 的收件人推導（全部角色都通知的規則下，每個納管項有哪些收件人） */
function oracleRecipients(item, users) {
  const by = Object.fromEntries(users.map((u) => [u.username, u]));
  const live = (u) => !!u && u.active !== false && !u.pureSupervisor && u.role !== 'pool';
  const admins = users.filter((u) => u.role === 'admin' && live(u)).map((u) => u.username);
  const owner = by[item.owner];
  const firstM1 = () => {
    let cur = owner; const seen = new Set();
    while (cur && cur.supervisor && !seen.has(cur.username)) {
      seen.add(cur.username);
      const up = by[cur.supervisor];
      if (!up) break;
      if (up.role === 'manager1' && live(up)) return up.username;
      cur = up;
    }
    return null;
  };
  const out = new Set();
  const m1 = firstM1();
  if (live(owner)) out.add(owner.username); else if (m1) out.add(m1); else admins.forEach((a) => out.add(a));
  const VALID = ['ERP', 'ITS', 'MDM', 'CRM', '資安'];
  const bu = VALID.includes(item.bu) ? item.bu : ((owner && Array.isArray(owner.bu) && owner.bu[0]) || '');
  const ownerBus = (owner && Array.isArray(owner.bu)) ? owner.bu : [];
  const secs = users.filter((u) => u.role === 'secretary' && live(u) && bu && (u.bu || []).includes(bu) && (!ownerBus.length || ownerBus.some((b) => (u.bu || []).includes(b)))).map((u) => u.username);
  (secs.length ? secs : admins).forEach((s) => out.add(s));
  (m1 ? [m1] : admins).forEach((s) => out.add(s));
  const sup = owner && by[owner.supervisor];
  if (sup && live(sup) && sup.role === 'manager2') out.add(sup.username);
  users.filter((u) => u.role === 'marketing' && live(u)).forEach((u) => out.add(u.username));
  return out;
}

async function groupRoles() {
  console.log('[角色] 逐帳號比對 maReminders、時間軸可視範圍、秘書欄位白名單');
  const T = todayTp();
  const owners = [U.u1, U.u2, U.u3, U.u4, U.u5, U.u6, U.uoff, U.m2, U.m1, U.m1b, U.gs];
  const ALLR = [R1x({ id: 'all', name: '全部角色', whenDays: 200, notify: ['owner', 'secretary', 'manager1', 'manager2', 'marketing'] }), R1x({ id: 'late', name: '到期後全部角色', whenDays: -1, repeatEveryDays: 7, condition: 'not_renewed', notify: ['owner', 'secretary', 'manager1', 'manager2', 'marketing'] })];
  function R1x(o) { return RULE(o); }
  const seed = (d) => {
    d.opportunities = []; d.maItems = []; d.quotations = [];
    owners.forEach((ow, i) => {
      for (let k = 0; k < 3; k++) {
        const id = `r${i}_${k}`;
        const co = (ow === U.u1 && k === 0) ? '假公司gA' : '假公司' + id;   // u1 的第一筆在集團 g1 內
        const dl = [[-30, 10, 120], [5, 199, 250], [-500, 0, 60], [30, 31, 90]][i % 4][k];
        const it = mkItem('i' + id, ow, dl, { company: co, termKind: k === 1 ? 'license' : 'ma', contractNo: 'CN-' + id + '-' + CNO, status: (i + k) % 7 === 0 ? 'ended' : 'tracking' });
        if (ow === U.u3 || ow === U.m1b) it.bu = 'ITS';
        d.maItems.push(it);
      }
    });
    // 納管項 BU ≠ 負責業務 BU（一級主管雖是收件人，但依 BU 過濾看不到 → 動態鈴鐺不可顯示）；有續約報價單
    d.maItems.push(mkItem('rBU', U.u1, 20, { company: '假公司rBU', bu: 'ITS' }));
    d.maItems.push(mkItem('rQ', U.u2, 20, { company: '假公司rQ' }));
    d.quotations.push(mkQuote('qq', 'o_rQ', U.u2, '假公司rQ', [], 'pending', { renewOf: 'rQ' }));
    d.reminderRules = { rules: ALLR, updatedAt: nowIso, updatedBy: U.adm };
    d.productCatalog = CAT();
  };
  await freshServer({ data: seed });
  const raw = readData();
  const users = seedUsers();
  const itemById = new Map(raw.maItems.map((i) => [i.id, i]));
  const eligible = (it) => it.status === 'tracking' && (() => { const dl = Math.round((Date.parse(it.to + 'T00:00:00Z') - Date.parse(T + 'T00:00:00Z')) / DAY); return dl <= 200 && dl >= -400; })();   // all：到期前 200 天起到迄日；late：迄日後一直到 400 天
  const expectedRecipients = {};
  raw.maItems.forEach((it) => { if (eligible(it)) expectedRecipients[it.id] = oracleRecipients(it, users); });
  const key = (x) => [x.company, x.product, x.to].join('|');
  const keyToId = new Map(raw.maItems.map((i) => [key(i), i.id]));
  const userRole = Object.fromEntries(users.map((u) => [u.username, u.role]));
  const CHUNKS = [ALLU.slice(0, 10), ALLU.slice(10, 20), ALLU.slice(20)];
  let compared = 0, leaks = 0;
  const sizes = new Set();
  for (let ci = 0; ci < CHUNKS.length; ci++) {
    if (ci > 0) await restartServer();
    const bad = await loginMany(CHUNKS[ci]);
    ok(bad.every((b) => ['_s_m1off:401', '_s_uoff:401', '_s_mktoff:401', '_s_secoff:401'].includes(b)), 'R 登入：只有停用帳號登不進來', bad.join());
    for (const who of CHUNKS[ci]) {
      if (!jar[who]) { eq([(await http(who, 'GET', '/api/poll-bundle')).status], [401], `R ${who} 停用帳號 poll-bundle 401`); continue; }
      const role = userRole[who];
      const pb = await http(who, 'GET', '/api/poll-bundle');
      const groupRole = role === 'tecopm' || role === 'groupsales';   // 集團角色：路徑白名單不含 poll-bundle（既有行為，這兩種角色本來就沒有鈴鐺）
      if (groupRole) {
        eq(pb.status, 403, `R ${who}(${role}) poll-bundle 403（既有的集團角色白名單）`);
        if (role === 'tecopm') { eq([(await http(who, 'GET', '/api/ma-items/ir0_0/reminders')).status, (await http(who, 'GET', '/api/ma-items/renewal-candidates?company=x')).status], [403, 403], `R ${who}(tecopm) 時間軸／候選 403`); continue; }
      } else ok(pb.status === 200 && Array.isArray(pb.json.maReminders), `R ${who} poll-bundle 帶 maReminders`, pb.text.slice(0, 120));
      const got = (pb.json && pb.json.maReminders) || [];
      // 可視範圍基準：該帳號實際在 GET /api/ma-items 看得到的
      const list = await http(who, 'GET', '/api/ma-items');
      let visIds;
      if (role === 'secretary') visIds = new Set((list.json.items || []).map((x) => keyToId.get(key(x))).filter(Boolean));
      else visIds = new Set((list.json.items || []).map((x) => x.id));
      sizes.add(visIds.size);
      if (!groupRole) {
        const gotIds = role === 'secretary' ? got.map((x) => x.company + '|' + x.product + '|' + x.to).map((k) => keyToId.get(k)) : got.map((x) => x.itemId);
        const expIds = Object.keys(expectedRecipients).filter((id) => expectedRecipients[id].has(who) && visIds.has(id)).sort();
        eq(gotIds.slice().sort(), expIds, `R ${who}(${role}) maReminders ＝ 「是收件人」∩「ma-items 可視範圍」（${expIds.length} 項）`);
        compared++;
        ok(gotIds.every((id) => visIds.has(id)), `R ${who} maReminders 全部在自己的可視範圍內（無外洩）`); if (!gotIds.every((id) => visIds.has(id))) leaks++;
        ok(!pb.text.includes(AMT) && !pb.text.includes(NOTE), `R ${who} 沒有金額、備註`);
        ok(!pb.text.includes('CN-') && !J(got).includes(CNO), `R ${who} maReminders 沒有合約編號`);
        if (role === 'secretary') ok(got.every((x) => Object.keys(x).sort().join() === 'body,company,days,from,id,level,ownerLabel,product,termKind,title,to,type'), `R ${who}(secretary) 欄位白名單（沒有納管項 id、規則、角色）`);
        else if (got.length) ok(got.every((x) => x.itemId && x.rule && Array.isArray(x.roles)), `R ${who} 一般角色有 itemId／rule／roles`);
      }
      // 時間軸：200 ⇔ 在 ma-items 可視範圍內
      const probe = ['ir0_0', 'ir1_0', 'ir2_1', 'ir3_2', 'ir4_0', 'ir8_1', 'ir9_0', 'ir10_0', 'rBU', 'rQ', 'ir6_0'];
      for (const id of probe) {
        if (!itemById.has(id)) continue;
        const t = await http(who, 'GET', `/api/ma-items/${id}/reminders`);
        const should = visIds.has(id);
        if (t.status !== (should ? 200 : 404)) ok(false, `R ${who} 時間軸 ${id}：可視=${should} 但狀態 ${t.status}`);
        else pass++;
        if (should && role === 'secretary') ok(!('id' in t.json.item) && t.json.rules.every((x) => !('sent' in x)) && !t.text.includes('CN-'), `R ${who}(secretary) 時間軸 ${id} 欄位白名單`);
        if (should && !t.text.includes('CN-') ) pass++;
      }
    }
  }
  ok(sizes.size >= 4, 'R 各角色可視筆數不同（基準有鑑別力）', [...sizes].join());
  ok(compared >= 18 && leaks === 0, `R 比對了 ${compared} 個可登入帳號、外洩 ${leaks}`);
  // rBU 這筆：一級主管 m1 是收件人但看不到（BU 不同）→ 不得出現在他的鈴鐺
  await restartServer();
  await loginMany([U.m1, U.u1, U.u2, U.sec2]);
  const pm1 = (await poll(U.m1)).maReminders.map((x) => x.itemId), pu1 = (await poll(U.u1)).maReminders.map((x) => x.itemId);
  ok(expectedRecipients.rBU.has(U.m1) && !pm1.includes('rBU') && pu1.includes('rBU'), 'R rBU：m1 是收件人但依 BU 看不到 → 不顯示；負責業務本人顯示（可視範圍過濾生效）');
  ok(!expectedRecipients.rQ.has(U.mkt) === false, 'R 行銷是收件人');
  const prq = (await poll(U.u2)).maReminders.find((x) => x.itemId === 'rQ');
  ok(prq && prq.quoteState === 'sent', 'R rQ：已送簽狀態');
  await stopServer(); cleanupFiles();
}

// ═════════════ 開關：MA_TRACKING 未設 ═════════════
function exportBase() {
  if (process.env.MAREM_BASE_DIR) return { dir: process.env.MAREM_BASE_DIR, cleanup() {} };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'marem-base-'));
  const baseRef = process.env.MAREM_BASE_REF || '3d85397';
  const a = spawnSync('git', ['archive', baseRef], { cwd: ROOT, maxBuffer: 256 * 1024 * 1024 });
  if (a.status !== 0) throw new Error('git archive ' + baseRef + ' 失敗：' + String(a.stderr));
  const t = spawnSync('tar', ['-x', '-C', dir], { input: a.stdout, maxBuffer: 256 * 1024 * 1024 });
  if (t.status !== 0) throw new Error('tar 失敗：' + String(t.stderr));
  spawnSync('cmd', ['/c', 'mklink', '/J', path.join(dir, 'node_modules'), path.join(ROOT, 'node_modules')], { stdio: 'ignore' });
  return { dir, cleanup() { spawnSync('cmd', ['/c', 'rmdir', path.join(dir, 'node_modules')], { stdio: 'ignore' }); try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ } } };
}
async function groupOff() {
  console.log('[開關] MA_TRACKING 未設 ＝ 與基準版逐位元相同');
  const run = async (root, port, mut, mailMode) => {
    await stopServer(); cleanupFiles(root); seedFiles({ dir: root, data: apiSeed });
    await startServer(mut || null, mailMode ? { MAIL_MODE: mailMode } : {}, root, port);   // 不傳 MA_TRACKING
    const base = `http://localhost:${port}`;
    const jr = {};
    const who6 = [U.adm, U.u1, U.u2, U.sec, U.mkt, U.m1];
    for (const w of who6) await loginAt(base, jr, w);
    const out = {};
    const rec = async (key, who, method, url, body, hdr) => { const r = await httpAt(base, jr, who, method, url, body, hdr); out[key] = r.status + ' ' + norm(r.text); return r; };
    out.cron = null;
    await rec('cron', null, 'GET', '/api/cron/mail-outbox', undefined, { Authorization: 'Bearer ' + CRON_SECRET });
    await rec('cronNoAuth', null, 'GET', '/api/cron/mail-outbox');
    if (mailMode) { await stopServer(); return { out, raw: readData(root) }; }
    for (const who of who6) for (const g of ['/api/poll-bundle', '/api/contract-reminders', '/api/contracts', '/api/notifications', '/api/quotations', '/api/product-catalog', '/api/me/permissions', '/api/opportunities', '/api/lost-opportunities', '/api/ma-items', '/api/my-tasks/deals']) await rec(who + ' ' + g, who, 'GET', g);
    const p1 = await rec('postQuote', U.u1, 'POST', '/api/quotations', qBody('oQ1', { renewOf: 'iA' }));
    const qid = p1.json && p1.json.id;
    await rec('postQuoteTwoNoRenew', U.u1, 'POST', '/api/quotations', qBody('oQ1'));
    await rec('putQuote', U.u1, 'PUT', `/api/quotations/${qid}`, { renewOf: 'iA' });
    await rec('getQuote', U.u1, 'GET', `/api/quotations/${qid}`);
    await rec('putOppWon', U.u1, 'PUT', '/api/opportunities/oR', { stage: 'Won', termFrom: '2026-01-01', termTo: '2026-12-31' });
    await rec('putOppBack', U.u1, 'PUT', '/api/opportunities/oR', { stage: 'B' });
    const extra = {};
    for (const [k, w, m, u, b] of [['rulesGet', U.adm, 'GET', '/api/admin/reminder-rules'], ['rulesPut', U.adm, 'PUT', '/api/admin/reminder-rules', { rules: [] }], ['sim', U.adm, 'POST', '/api/admin/reminder-rules/simulate', { daysLeft: 1, quoteState: 'none' }], ['timeline', U.u1, 'GET', '/api/ma-items/iA/reminders'], ['cand', U.u1, 'GET', '/api/ma-items/renewal-candidates?company=CoA'], ['notRenew', U.u1, 'PATCH', '/api/ma-items/iA/not-renew', { reason: '客戶不續約' }], ['reopen', U.adm, 'PATCH', '/api/ma-items/iA/reopen']]) {
      extra[k] = await httpAt(base, jr, w, m, u, b);
      out['x_' + k] = extra[k].status + ' ' + norm(extra[k].text);
    }
    const raw = readData(root);
    out.keys = J(Object.keys(raw).sort());
    out.quoteKeys = J((raw.quotations || []).map((q) => Object.keys(q).sort().filter((k) => !['createdAt', 'updatedAt', 'id', 'quoteNo', 'validUntil', 'quoteDate'].includes(k))));
    out.maItems = J(raw.maItems.map((i) => [i.id, i.status, Object.keys(i).sort().join()]));
    out.notifTypes = J((raw.notifications || []).map((n) => n.type).sort());
    await stopServer(); cleanupFiles(root);
    return { out, extra, raw };
  };
  const baseX = exportBase();
  let A, B, A2, B2, A3, B3;
  try {
    A = await run(baseX.dir, PORT);
    B = await run(ROOT, PORT, MUT_CUR);
    A2 = await run(baseX.dir, PORT, null, 'log');
    B2 = await run(ROOT, PORT, MUT_CUR, 'log');
  } finally { baseX.cleanup(); }
  const diffKeys = Object.keys(A.out).filter((k) => A.out[k] !== B.out[k]);
  ok(diffKeys.length === 0, `F1 off：poll-bundle／合約提醒／合約／通知／報價單／商機／到期管理／cron mail-outbox／報價單 POST-PUT（帶 renewOf）／商機成交回退／資料檔鍵與新端點的輸出與基準版逐位元相同（比對 ${Object.keys(A.out).length} 項）`, diffKeys.join(', '));
  if (diffKeys.length) { const x = String(A.out[diffKeys[0]]), y = String(B.out[diffKeys[0]]); let i = 0; while (i < x.length && x[i] === y[i]) i++; console.log('    差異首項：', diffKeys[0], '位置', i, '\n      前：', x.slice(Math.max(0, i - 80), i + 160), '\n      後：', y.slice(Math.max(0, i - 80), i + 160)); }
  const diffKeys2 = Object.keys(A2.out).filter((k) => A2.out[k] !== B2.out[k]);
  ok(diffKeys2.length === 0, 'F1b off：MAIL_MODE=log 的 cron mail-outbox 回應與基準版逐位元相同', diffKeys2.join(', '));
  ok(!/reminders/.test(B.out.cron) && !/reminders/.test(B2.out.cron), 'F2 off：cron 回應沒有 reminders');
  ok(B.out['_s_u1 /api/poll-bundle'] && !/maReminders/.test(B.out['_s_u1 /api/poll-bundle']), 'F2b off：poll-bundle 沒有 maReminders 鍵');
  ok(Object.keys(B.extra).every((k) => B.extra[k].status === A.extra[k].status && (B.extra[k].status === 404 || B.extra[k].status === 405)), 'F3 off：新端點（規則 GET/PUT、試算、時間軸、候選、不續約、撤銷）全部 404，和基準版相同', J(Object.keys(B.extra).map((k) => k + ':' + B.extra[k].status)));
  ok(!('reminderRules' in B.raw) && !('reminderLog' in B.raw) && B.raw.quotations.every((q) => !('renewOf' in q)), 'F4 off：資料檔沒有 reminderRules／reminderLog，報價單沒有 renewOf（請求裡的 renewOf 被忽略）');
  ok(/"contractId":"cSM"/.test(B.out['_s_u1 /api/poll-bundle']) && /"contractId":"cEM"/.test(B.out['_s_u1 /api/poll-bundle']), 'F5 off：SERVICE_MA 合約仍在舊的合約到期提醒（沒有被排除）——與基準版一致');
  ok(!!B.raw.contracts.find((c) => c.id === 'cSM')._contract30Reminded && !!B.raw.contracts.find((c) => c.id === 'cEM')._contract30Reminded && J(B.raw.contracts.map((c) => [c.id, !!c._contract30Reminded])) === J(A.raw.contracts.map((c) => [c.id, !!c._contract30Reminded])), 'F5b off：checkOverdueReminders 照常處理 SERVICE_MA 合約（有 _contract30Reminded 標記，與基準版一致）');
  ok(!B.raw.notifications.some((n) => n.type === 'ma_reminder' || n.type === 'ma_not_renew'), 'F6 off：不產生 ma_reminder／ma_not_renew 通知');
}

// ═════════════ 變異測試 ═════════════
const E = (find, rep) => ({ find, rep });
const MUTS = [
  { name: 'M1 拿掉去重', groups: ['unit', 'sim', 'api:main'], spec: [{ file: 'lib/maReminders.js', edits: [E('if (view.entries[key]) { stats.skippedDup++; return; }', 'if (false) { stats.skippedDup++; return; }')] }] },
  { name: 'M2 拿掉台北時區（改用 UTC 日期）', groups: ['unit'], spec: [{ file: 'lib/maTracking.js', edits: [E('return new Date((Number.isFinite(t) ? t : Date.now()) + 8 * 3600 * 1000).toISOString().slice(0, 10);', 'return new Date((Number.isFinite(t) ? t : Date.now()) + 0 * 3600 * 1000).toISOString().slice(0, 10);')] }] },
  { name: 'M3 拿掉補發（只等於當天才發）', groups: ['unit', 'sim'], spec: [{ file: 'lib/maReminders.js', edits: [E("if (dl > rule.whenDays) return { active: false, reason: 'NOT_REACHED' };", "if (dl !== rule.whenDays) return { active: false, reason: 'NOT_REACHED' };")] }] },
  { name: 'M4 拿掉條件判斷', groups: ['unit', 'sim', 'api:rules'], spec: [{ file: 'lib/maReminders.js', edits: [E('function conditionHolds(condition, item, qs) {', 'function conditionHolds(condition, item, qs) {\n  return true;')] }] },
  { name: 'M5 拿掉收件人排除（停用／掛名也通知）', groups: ['unit', 'volume'], spec: [{ file: 'lib/maRecipients.js', edits: [E('function usable(u) { return !!u && OR.isActiveUser(u) && !u.pureSupervisor; }', 'function usable(u) { return !!u; }')] }] },
  { name: 'M6 拿掉 dedup 與通知同一次 save', groups: ['api:main'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E('if (r.changed) D.db.save(data);', 'if (r.changed) { const c = JSON.parse(JSON.stringify(data)); delete c.reminderLog; D.db.save(c); }')] }] },
  { name: 'M7 simulate 另寫一份邏輯（偏一天）', groups: ['api:rules'], spec: [{ file: 'lib/maReminders.js', edits: [E('const to = MA.addDaysIso(today, dl);', 'const to = MA.addDaysIso(today, dl + 1);')] }] },
  { name: 'M7b simulate 條件判斷另寫一份（草稿當已送簽）', groups: ['api:rules'], spec: [{ file: 'lib/maReminders.js', edits: [E("approval: { state: scenario.quoteState === 'sent' ? 'pending' : 'none' }", "approval: { state: 'pending' }")] }] },
  { name: 'M8 拿掉 SERVICE_MA 排除（poll-bundle）', groups: ['api:main'], spec: [{ file: 'server.js', edits: [E("owners.includes(c.owner) && !(MA.enabled() && c.type === 'SERVICE_MA'));\n    const today", 'owners.includes(c.owner));\n    const today')] }] },
  { name: 'M8b 拿掉 SERVICE_MA 排除（checkOverdueReminders）', groups: ['api:main'], spec: [{ file: 'server.js', edits: [E("if (MA.enabled() && c.type === 'SERVICE_MA') continue;", '')] }] },
  { name: 'M9 拿掉單次上限', groups: ['unit', 'volume'], spec: [{ file: 'lib/maReminders.js', edits: [E('const take = names.slice(0, Math.max(0, cap));', 'const take = names;')] }] },
  { name: 'M10 秘書也給納管項 id 與規則細節', groups: ['unit', 'roles'], spec: [{ file: 'lib/maReminders.js', edits: [E('if (o.secretary) {', 'if (false) {')] }] },
  { name: 'M11 動態鈴鐺拿掉可視範圍過濾', groups: ['roles'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E('visible: (item) => { if (!vis)', 'visible: (item) => { return true; if (!vis)')] }] },
  { name: 'M12 條件不成立時不標通知已讀', groups: ['unit', 'sim', 'api:flow'], spec: [{ file: 'lib/maReminders.js', edits: [E('if (!mine || !n.maItemIds.some((i) => mine.has(i))) { n.read = true; marked++; }', 'if (false) { n.read = true; marked++; }')] }] },
  { name: 'M13 不續約不檢查負責人／管理員', groups: ['api:main'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E("if (!isOwnerOrAdmin(req, it.owner)) return err(res, 403, 'FORBIDDEN', '只有負責業務或管理員可以標記不續約');", '')] }] },
  { name: 'M14 撤銷不續約不限管理員', groups: ['api:main'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E("if (req.session.user.role !== 'admin') return err(res, 403, 'FORBIDDEN', '只有管理員可以撤銷不續約');", '')] }] },
  { name: 'M15 不續約原因不必填', groups: ['api:main'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E('if (Array.from(reason).length < MA.SKIP_REASON_MIN) return err(', 'if (false) return err(')] }] },
  { name: 'M16 renewOf 不限可視範圍', groups: ['api:main'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E('const it = visibleItems(req, data).find((x) => x && x.id === id);\n  if (!it) return { ok: false, status: 404', 'const it = (data.maItems || []).find((x) => x && x.id === id);\n  if (!it) return { ok: false, status: 404')] }] },
  { name: 'M17 renewOf 不檢查同公司', groups: ['api:main'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E('if (!OL.companyEq(it.company, company)) return', 'if (false) return')] }] },
  { name: 'M18 renewOf 讓作廢項也能指定', groups: ['api:main'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E("if (it.status === 'void') return { ok: false, status: 409", "if (false) return { ok: false, status: 409")] }] },
  { name: 'M19 續約成交不標舊項已續約', groups: ['api:flow'], spec: [{ file: 'lib/maTracking.js', edits: [E('out.renewedItems = applyRenewals(data, opp, ctx);', 'out.renewedItems = [];')] }] },
  { name: 'M19b 續約成交不檢查報價單已核准', groups: ['api:flow'], spec: [{ file: 'lib/maTracking.js', edits: [E("filter((q) => QT.termEditState(q) && typeof q.renewOf === 'string' && q.renewOf)", "filter((q) => typeof q.renewOf === 'string' && q.renewOf)")] }] },
  { name: 'M20 續約成交回退不恢復舊項', groups: ['api:flow'], spec: [{ file: 'lib/maTracking.js', edits: [E("if (it && it.status === 'renewed' && it.renewedBy && it.renewedBy.oppId === oppId) {", 'if (false) {')] }] },
  { name: 'M21 規則天數重複不檢查', groups: ['unit', 'api:rules'], spec: [{ file: 'lib/maReminders.js', edits: [E('if (dups.has(dk)) return bad(', 'if (false) return bad(')] }] },
  { name: 'M22 getConfig 信任儲存值', groups: ['unit', 'api:rules'], spec: [{ file: 'lib/maReminders.js', edits: [E('const v = validateConfig({ rules: c.rules });', 'const v = { ok: true, value: { rules: c.rules } };')] }] },
  { name: 'M23 規則 GET 不限管理員', groups: ['api:rules'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E("app.get('/api/admin/reminder-rules', maOn, requireAdmin,", "app.get('/api/admin/reminder-rules', maOn, requireAuth,")] }] },
  { name: 'M24 試算不限管理員', groups: ['api:rules'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E("'/api/admin/reminder-rules/simulate', maOn, requireAdmin,", "'/api/admin/reminder-rules/simulate', maOn, requireAuth,")] }] },
  { name: 'M25 時間軸不限可視範圍', groups: ['api:main', 'roles'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E('const it = visibleItems(req, data).find((x) => x && x.id === req.params.id);', 'const it = (data.maItems || []).find((x) => x && x.id === req.params.id);')] }] },
  { name: 'M26 秘書時間軸也給收件人名單', groups: ['api:main', 'roles'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E('{ now: Date.now(), secretary, labelOf', '{ now: Date.now(), secretary: false, labelOf')] }] },
  { name: 'M27 續約候選含非追蹤中', groups: ['api:main'], spec: [{ file: 'lib/maReminderRoutes.js', edits: [E("it && it.status === 'tracking' && OL.companyEq(it.company, company)", 'it && OL.companyEq(it.company, company)')] }] },
  { name: 'M28 拿掉「追蹤中」共同前提', groups: ['unit', 'sim'], spec: [{ file: 'lib/maReminders.js', edits: [E("if (!item || item.status !== 'tracking' || !KINDS.includes(item.termKind)) return null;", 'if (!item || !KINDS.includes(item.termKind)) return null;')] }] },
  { name: 'M29 重複週期恆為 0', groups: ['unit', 'sim'], spec: [{ file: 'lib/maReminders.js', edits: [E('return Math.floor((rule.whenDays - dl) / rule.repeatEveryDays);', 'return 0;')] }] },
  { name: 'M30 到期前規則過期後仍補發', groups: ['unit', 'sim'], spec: [{ file: 'lib/maReminders.js', edits: [E("if (rule.whenDays > 0 && dl < 0) return { active: false, reason: 'EXPIRED_WINDOW' };", '')] }] },
  { name: 'M31 去重鍵不含迄日', groups: ['unit', 'sim'], spec: [{ file: 'lib/maReminders.js', edits: [E('const logKey = (ruleId, itemId, recipient, endDate, period) => [ruleId, itemId, recipient, endDate, period].join', "const logKey = (ruleId, itemId, recipient, endDate, period) => [ruleId, itemId, recipient, '', period].join")] }] },
  { name: 'M32 彙總改成每個（項目×規則）一項', groups: ['unit', 'volume'], spec: [{ file: 'lib/maReminders.js', edits: [E('if (!items.has(p.item.id)) items.set(p.item.id, p);', 'items.set(p.key, p);')] }] },
  { name: 'M33 dryRun 也會寫入', groups: ['unit'], spec: [{ file: 'lib/maReminders.js', edits: [E("const view = dry ? { entries: Object.assign({}, (data.reminderLog && data.reminderLog.entries) || {}) } : ensureLog(data);", 'const view = ensureLog(data);')] }] },
  { name: 'M34 補掃不節流', groups: ['unit'], spec: [{ file: 'lib/maReminders.js', edits: [E('return sinceAttempt >= LAZY_INTERVAL_MS && nowMs', 'return true || nowMs')] }] },
  { name: 'M35 擠掉通知時不優先丟已讀', groups: ['unit', 'volume'], spec: [{ file: 'lib/maReminders.js', edits: [E('if (list[i] && list[i].read) { drop.add(i); over--; }', 'if (false) { drop.add(i); over--; }')] }] },
  { name: 'M36 cron 掃描失敗會讓 mail cron 失敗', groups: ['api:fail'], spec: [{ file: 'lib/mail/routes.js', edits: [E("if (pr.error) { console.warn('[mail cron] 到期提醒掃描失敗:', pr.error && pr.error.message); return null; }", 'if (pr.error) throw pr.error;'), E("catch (e) { console.warn('[mail cron] 到期提醒掃描例外:', e && e.message); return null; }", 'catch (e) { throw e; }')] }] },
  { name: 'M37 無收件人時不記稽核', groups: ['unit'], spec: [{ file: 'lib/maReminders.js', edits: [E("audits.push({ action: 'MA_REMINDER_NO_RECIPIENT', target: e.item.company || '', detail: '規則「'", "void ({ action: 'MA_REMINDER_NO_RECIPIENT', target: e.item.company || '', detail: '規則「'")] }] },
  { name: 'M38 管理員不在收件人退路', groups: ['unit', 'api:main'], spec: [{ file: 'lib/maRecipients.js', edits: [E("else { byRole.secretary = admins.slice(); issues.push({ role: 'secretary', code: 'NO_SECRETARY' }); }", "else { byRole.secretary = []; issues.push({ role: 'secretary', code: 'NO_SECRETARY' }); }")] }] },
  { name: 'M39 秘書不比對 BU', groups: ['unit', 'roles'], spec: [{ file: 'lib/maRecipients.js', edits: [E('.filter((u) => { const sb = buList(u.bu); return sb.includes(bu) && (!ownerBus.length || ownerBus.some((b) => sb.includes(b))); })', '.filter((u) => true)')] }] },
  { name: 'M40 開關預設改成開', groups: ['off'], spec: [{ file: 'lib/quoteTerms.js', edits: [E("return String(e.MA_TRACKING == null ? '' : e.MA_TRACKING).trim() === '1';", 'return true;')] }] },
];
function REAL_SWAP_LIBS(spec) {
  // 把 spec 中屬於 lib 的修改寫成 __mut_ 副本並讓 REF 指過去；回傳清理函式
  const libFiles = ['maReminders', 'maRecipients', 'maTracking', 'orgResolve'];
  const edits = {};
  spec.filter((x) => /^lib\/[A-Za-z]+\.js$/.test(x.file)).forEach((x) => { edits[path.basename(x.file, '.js')] = x.edits; });
  if (!Object.keys(edits).some((k) => libFiles.includes(k))) return null;
  const tmps = [];
  const w = (name) => path.join(ROOT, 'lib', '__mut_' + name + '.js');
  libFiles.forEach((name) => {
    let src = fs.readFileSync(path.join(ROOT, 'lib', name + '.js'), 'utf8');
    (edits[name] || []).forEach((e) => { if (!src.includes(e.find)) throw new Error('MUTATION ANCHOR MISSING ' + name + ': ' + e.find.slice(0, 50)); src = src.replace(e.find, () => e.rep); });
    libFiles.forEach((o) => { src = src.split(`require('./${o}')`).join(`require('./__mut_${o}')`); });
    fs.writeFileSync(w(name), src); tmps.push(w(name));
  });
  const saved = { MR: REF.MR, RC: REF.RC, MA: REF.MA };
  tmps.forEach((t) => { delete require.cache[t]; });
  REF.MR = require(w('maReminders')); REF.RC = require(w('maRecipients')); REF.MA = require(w('maTracking'));
  return () => { REF.MR = saved.MR; REF.RC = saved.RC; REF.MA = saved.MA; tmps.forEach((t) => { delete require.cache[t]; try { fs.unlinkSync(t); } catch { /* 忽略 */ } }); };
}
async function runGroup(g, sub) {
  if (g === 'unit') groupUnit();
  else if (g === 'sim') groupSim('sim');
  else if (g === 'volume') groupSim('volume');
  else if (g === 'roles') await groupRoles();
  else if (g === 'off') await groupOff();
  else if (g === 'api') await groupApi(sub);
}
async function groupMut() {
  console.log('[變異] 每個變異都必須讓對應群組變紅');
  let undetected = 0;
  const only = (process.env.MAREM_MUT_ONLY || '').split(',').filter(Boolean);
  for (const m of MUTS) {
    if (only.length && !only.includes(m.name.split(' ')[0])) continue;
    const before = { p: pass, f: fail, n: failures.length };
    quiet = true; MUT_CUR = m.spec;
    let cleanup = null;
    try {
      cleanup = REAL_SWAP_LIBS(m.spec);
      for (const g of m.groups) {
        const [name, sub] = g.split(':');
        if (['unit', 'sim', 'volume'].includes(name) && !cleanup) continue;   // 這個變異不在純邏輯庫裡
        await runGroup(name, sub);
      }
    } catch (e) { fail++; failures.push('(變異造成例外) ' + e.message); }
    finally { if (cleanup) cleanup(); }
    quiet = false; MUT_CUR = null;
    const newFails = fail - before.f;
    const killedBy = failures.slice(before.n, before.n + 3).map((x) => x.slice(0, 48));
    console.log('  ' + (newFails > 0 ? '✓' : '✗') + ' ' + m.name + '：' + (newFails > 0 ? '變紅（' + newFails + ' 項失敗：' + killedBy.join(' / ') + '…）' : '沒有變紅！對應測試無效'));
    pass = before.p; fail = before.f; failures.length = before.n;
    if (newFails === 0) { undetected++; fail++; failures.push('變異未被偵測：' + m.name); }
  }
  console.log('  變異總計 ' + MUTS.length + ' 個，未被偵測 ' + undetected + ' 個');
  await stopServer(); cleanupFiles();
}

// ═════════════ 主程式 ═════════════
(async () => {
  try {
    if (want('unit')) groupUnit();
    if (want('sim')) groupSim('sim');
    if (want('volume')) groupSim('volume');
    if (want('api')) await groupApi();
    if (want('roles')) await groupRoles();
    if (want('off')) await groupOff();
    if (process.env.MAREM_MUT === '1') await groupMut();
  } catch (e) {
    console.log('例外：', e && e.stack || e); fail++; failures.push('例外：' + (e && e.message));
  } finally {
    await stopServer(); cleanupFiles();
    try { fs.unlinkSync(PRELOAD); } catch { /* 忽略 */ }
    try { fs.readdirSync(path.join(ROOT, 'lib')).filter((f) => f.startsWith('__mut_')).forEach((f) => fs.unlinkSync(path.join(ROOT, 'lib', f))); } catch { /* 忽略 */ }
  }
  console.log(`\n到期活規則提醒檢查結果：${pass} 通過，${fail} 失敗`);
  if (fail) console.log(failures.map((f) => ' - ' + f).join('\n'));
  process.exit(fail ? 1 : 0);
})();
