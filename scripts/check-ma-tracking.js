#!/usr/bin/env node
'use strict';
/**
 * 成交納管＋任務卡＋合約串接（P2a，MA_TRACKING=1）檢查：單元（lib/maTracking.js、lib/orgResolve.js、productCatalog.resolveTermRule）
 * ＋ API（真實 server.js，三種成交路徑一致、缺期間 400 無副作用、冪等、簽約／略過／延後／改派、各端點權限）
 * ＋ 角色權限一致性（22 帳號沙盒，逐角色比對「由 data.json＋GET /api/opportunities 暴力算出的預期」）
 * ＋ 開關關閉（MA_TRACKING 未設）與基準版（git archive 的 P1 提交）逐位元相同 ＋ 變異測試。
 * 不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體；資料檔只存在於本次執行；全部假資料；結束清掉伺服器。
 *
 *   node scripts/check-ma-tracking.js                       單元＋API＋角色＋開關
 *   MATRK_ONLY=unit,api,roles,off node ...                  只跑指定群組
 *   MATRK_MUT=1 node ...                                    另外跑變異測試（每個變異都必須讓對應群組變紅）
 *   MATRK_PORT（預設 3292；限 3290–3299 或 3310–3319）／MATRK_BASE_REF（開關比對的基準提交，預設 7efdba5＝P2 完成）／MATRK_BASE_DIR
 * 守門：工作樹已有 auth.json／data.json／audit.log.json 時拒絕執行。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.MATRK_PORT || '3292', 10);
if (!((PORT >= 3290 && PORT <= 3299) || (PORT >= 3310 && PORT <= 3319))) { console.log('拒絕：此腳本只能用埠 3290–3299 或 3310–3319'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map((f) => path.join(ROOT, f));
{
  const g = FILES.filter((f) => /(auth|data|audit.log)\.jsonl?$/.test(f)).filter((f) => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map((f) => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
}
const ONLY = (process.env.MATRK_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length || ONLY.includes(k);
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const XLSX = require(path.join(ROOT, 'node_modules', 'xlsx'));
const REAL = { MA: require('../lib/maTracking.js'), OR: require('../lib/orgResolve.js'), PC: require('../lib/productCatalog.js') };
const REF = { MA: REAL.MA };   // 變異測試時換成變異後的副本

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 400)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = (n) => new Promise((r) => setTimeout(r, n));

// ───────────────── 沙盒資料（全部假的）─────────────────
const PW = 'Ma' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const NOW = Date.now();
const nowIso = new Date(NOW).toISOString();
const DAY = 86400000;
const agoIso = (days) => new Date(NOW - days * DAY).toISOString();
const todayTp = () => new Date(Date.now() + 8 * 3600e3).toISOString().slice(0, 10);   // 獨立於 lib 的台北日期算法
const plusDays = (d, n) => new Date(Date.parse(d + 'T00:00:00Z') + n * DAY).toISOString().slice(0, 10);
const AMT = '8765432';            // 假金額（K）：不得出現在任何 MA 端點與通知
const NOTE = '機密備註XYZ';        // 假備註：不得出現在 MA 端點
const SECRET_CO = '機密客戶丙股份有限公司';   // 只有 u2 看得到的客戶（重複合約編號訊息不得洩漏）
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false }, extra || {});
const U = {
  adm: '_s_adm', exe: '_s_exe', acc: '_s_acc', fin: '_s_fin', m1: '_s_m1', m1b: '_s_m1b', m1p: '_s_m1p', m1off: '_s_m1off', m2: '_s_m2',
  u1: '_s_u1', u2: '_s_u2', u3: '_s_u3', u4: '_s_u4', u5: '_s_u5', uoff: '_s_uoff', mkt: '_s_mkt', mkt2: '_s_mkt2', mktoff: '_s_mktoff',
  sec: '_s_sec', sec2: '_s_sec2', gs: '_s_gs', tp: '_s_tp', cs: '_s_cs',
};
function seedUsers() {
  return [
    mkUser(U.adm, 'admin'), mkUser(U.exe, 'executive'), mkUser(U.acc, 'accounting_manager'), mkUser(U.fin, 'finance_manager'),
    mkUser(U.m1, 'manager1', { bu: ['ERP'], displayName: '測試主管一' }),
    mkUser(U.m1b, 'manager1', { bu: ['ITS'], displayName: '測試主管二' }),
    mkUser(U.m1p, 'manager1', { bu: ['ERP'], supervisor: U.m1, pureSupervisor: true, displayName: '掛名主管' }),
    mkUser(U.m1off, 'manager1', { bu: ['ERP'], active: false, supervisor: U.m1, displayName: '離職主管' }),
    mkUser(U.m2, 'manager2', { bu: ['ERP'], supervisor: U.m1 }),
    mkUser(U.u1, 'user', { bu: ['ERP'], supervisor: U.m2, displayName: '測試業務甲', nickname: '甲' }),
    mkUser(U.u2, 'user', { bu: ['ERP'], supervisor: U.m1, displayName: '測試業務乙' }),
    mkUser(U.u3, 'user', { bu: ['ITS'], supervisor: U.m1b, displayName: '測試業務丙' }),
    mkUser(U.u4, 'user', { bu: ['ERP'], supervisor: U.m1p, displayName: '測試業務丁' }),
    mkUser(U.u5, 'user', { bu: ['ERP'], supervisor: U.m1off, displayName: '測試業務戊' }),
    mkUser(U.uoff, 'user', { bu: ['ERP'], active: false, supervisor: U.m1, displayName: '離職業務' }),
    mkUser(U.mkt, 'marketing', { bu: ['ERP'], displayName: '測試行銷' }), mkUser(U.mkt2, 'marketing', { displayName: '測試行銷二' }),
    mkUser(U.mktoff, 'marketing', { displayName: '離職行銷', active: false }),
    mkUser(U.sec, 'secretary', { bu: ['ERP'] }), mkUser(U.sec2, 'secretary', { bu: ['ITS'] }),
    mkUser(U.gs, 'groupsales', { bu: ['ERP'], viewOwnerScope: U.u1, viewGroupId: 'g1', displayName: '測試集團業務' }),
    mkUser(U.tp, 'tecopm', { viewOwnerScope: U.u1, viewGroupId: 'g1', displayName: '測試集團PM' }),
    mkUser(U.cs, 'consult_manager_south', { displayName: '測試顧問主管' }),
  ];
}
const CAT = () => ({ bus: { ERP: [{ group: '假商品群組', items: [
  { name: 'Svc MA', termKind: 'ma', startsTask: true }, { name: 'Svc Lic', termKind: 'license', startsTask: false },
  { name: 'Svc LicTask', termKind: 'license', startsTask: true }, { name: 'Plain', }, { name: 'Svc MA Long', termKind: 'ma', startsTask: false },
] }] } });
const MAP = 'Svc MA 一年', LICP = 'Svc Lic 永久', PLAINP = 'Plain 專案';
const mkOpp = (id, owner, product, stage, extra) => Object.assign({ id, owner, bu: 'ERP', company: '假公司' + id, contactName: '假聯絡人', product, amount: AMT, description: NOTE, expectedDate: '2026-12-31', stage, createdAt: nowIso }, extra || {});
const mkQuote = (id, oppId, owner, items, approvedAt, extra) => Object.assign({ id, oppId, owner, quoteNo: 'QT-' + id, company: '假公司' + oppId, items, status: 'approved',
  approval: approvedAt ? { state: 'approved', cur: 1, steps: [{ status: 'approved', at: approvedAt, label: '核決' }], hash: 'x' } : { state: 'pending', steps: [] },
  createdAt: nowIso, updatedAt: approvedAt || nowIso }, extra || {});
const qi = (lid, kind, from, to) => ({ lid, desc: '品項' + lid, qty: 1, price: 100, termKind: kind, termFrom: from, termTo: to });

function seedData(opts) {
  opts = opts || {};
  const d = {
    contacts: [], companies: [], visits: [], receivables: [], callins: [], keyAccounts: [], industries: [], notifications: [],
    productCatalog: CAT(),
    groups: [{ id: 'g1', name: '假集團', memberCompanies: ['假公司gA'] }],
    opportunities: [
      mkOpp('oQ1', U.u1, MAP, 'C'), mkOpp('oQ2', U.u1, MAP, 'C'), mkOpp('oQ3', U.u1, MAP, 'C'), mkOpp('oQ4', U.u1, MAP, 'C'),
      mkOpp('oBF', U.u1, MAP, 'Won', { achievedDate: '2026-09-01' }),
      mkOpp('oOld', U.uoff, MAP, 'Won', { achievedDate: '2026-08-01' }),
    ],
    quotations: [
      mkQuote('q1a', 'oQ1', U.u1, [qi('l1', 'ma', '2025-01-01', '2025-12-31')], '2026-05-01T00:00:00.000Z'),
      mkQuote('q1b', 'oQ1', U.u1, [qi('l1', 'ma', '2026-03-01', '2027-02-28'), qi('l2', 'ma', '2026-03-01', '2027-02-28'), qi('l3', 'license', '2026-03-01', '2029-02-28'), { lid: 'l4', desc: '一般品項', qty: 1, price: 5 }], '2026-06-01T00:00:00.000Z'),
      mkQuote('q1c', 'oQ1', U.u1, [qi('l1', 'ma', '2030-01-01', '2030-12-31')], null),   // 簽核中（未核准）不算
      mkQuote('q2a', 'oQ2', U.u1, [qi('l1', 'ma', '2026-04-01', '2027-03-31')], '2026-06-01T00:00:00.000Z'),
      mkQuote('q3a', 'oQ3', U.u1, [qi('l1', 'ma', '', ''), qi('l2', 'ma', '2026-05-01', '2026-04-01')], '2026-06-01T00:00:00.000Z'),   // 期間不合法 → 等於沒有
      mkQuote('q4a', 'oQ4', U.u1, [qi('l1', 'ma', '2026-04-01', '2027-03-31')], '2026-06-01T00:00:00.000Z', { invalid: { at: nowIso, by: 'x', reason: 'OPP_VOID' } }),   // 已失效不算
    ],
    contracts: [
      { id: 'cLEG', owner: U.u2, contractNo: 'LEG-001', company: SECRET_CO, product: 'ERP MA', startDate: '2026-01-01', endDate: '2026-12-31', amount: '1', type: 'ERP_MA', createdAt: nowIso },
      { id: 'cGONE', owner: U.u2, contractNo: 'GONE-001', company: '已刪除客戶', product: 'ERP MA', startDate: '2026-01-01', endDate: '2026-12-31', amount: '1', type: 'ERP_MA', createdAt: nowIso, deleted: true },
    ],
    maItems: [
      { id: 'iOld', oppId: 'oOld', quoteId: '', owner: U.uoff, company: '假公司oOld', product: MAP, termKind: 'ma', from: '2026-01-01', to: '2026-12-31', contractNo: '', contractId: '', status: 'tracking', source: 'request', createdAt: agoIso(20), updatedAt: agoIso(20), history: [] },
    ],
    dealTasks: [
      { id: 'tOld', type: 'ma_sign', oppId: 'oOld', maItemId: 'iOld', maItemIds: ['iOld'], owner: U.uoff, company: '假公司oOld', product: MAP, status: 'open', steps: [{ key: 'sign', status: 'pending', value: null, doneBy: '', doneAt: '', reason: '' }], createdAt: agoIso(20), updatedAt: agoIso(20), history: [] },
    ],
  };
  if (opts.data) opts.data(d);
  return d;
}
function seedFiles(opts) {
  opts = opts || {};
  const DIR = opts.dir || ROOT;
  fs.writeFileSync(path.join(DIR, 'auth.json'), JSON.stringify({ users: seedUsers() }, null, 2));
  fs.writeFileSync(path.join(DIR, 'data.json'), JSON.stringify(seedData(opts), null, 2));
}
const readData = (dir) => JSON.parse(fs.readFileSync(path.join(dir || ROOT, 'data.json'), 'utf8'));
function cleanupFiles(dir) { FILES.forEach((f) => { try { fs.unlinkSync(dir ? path.join(dir, path.basename(f)) : f); } catch { /* 不存在 */ } }); }

// ───────────────── 伺服器（預載腳本可在載入時改寫指定檔案＝變異測試）─────────────────
const PRELOAD = path.join(os.tmpdir(), 'matrk_preload_' + process.pid + '.js');
fs.writeFileSync(PRELOAD, [
  "const Module=require('module'),path=require('path');const orig=Module.prototype._compile;",
  "Module.prototype._compile=function(content,filename){const raw=process.env.MATRK_MUT_SPEC;",
  "if(raw){const specs=JSON.parse(raw);for(const m of specs){if(path.resolve(filename)===path.resolve(process.env.MATRK_ROOT,m.file)){for(const e of m.edits){if(!content.includes(e.find))throw new Error('MUTATION ANCHOR MISSING: '+e.find.slice(0,60));content=content.replace(e.find,()=>e.rep);}}}}",
  "return orig.call(this,content,filename);};",
].join('\n') + '\n');
let child = null, childPid = null;
async function startServer(mut, extraEnv, root, port) {
  root = root || ROOT; port = port || PORT;
  const env = Object.assign({}, process.env, { PORT: String(port), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: `http://localhost:${port}` });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'QUOTE_FLOW_ENGINE', 'QUOTE_FLOW_DESIGNER', 'FEATURE_GUARD_MODE', 'MATRK_MUT_SPEC', 'MA_TRACKING']) delete env[k];
  Object.assign(env, extraEnv || {});
  if (mut && mut.length) { env.MATRK_MUT_SPEC = JSON.stringify(mut); env.MATRK_ROOT = ROOT; }
  await assertPortFree(port);
  child = spawn(process.execPath, ['-r', PRELOAD, path.join(root, 'server.js')], { cwd: root, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { out += d; });
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

// ───────────────── HTTP ─────────────────
async function httpAt(base, jr, who, method, url, body, raw) {
  const h = {};
  if (body !== undefined && !raw) h['Content-Type'] = 'application/json';
  if (who && jr[who]) h.Cookie = jr[who];
  const res = await fetch(base + url, { method, headers: h, body: body === undefined ? undefined : (raw ? body : JSON.stringify(body)), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: res.status, json, text: buf.toString('utf8') };
}
const http = (who, method, url, body, raw) => httpAt(BASE, jar, who, method, url, body, raw);
async function loginAt(base, jr, who) {
  const res = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status === 200) jr[who] = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
  return res.status;
}
const login = (who) => loginAt(BASE, jar, who);
const ALLU = Object.values(U);
const LOGIN_STATUS = {};
async function loginAll(list) { for (const w of (list || ALLU)) LOGIN_STATUS[w] = await login(w); }
async function auditLines(action) {
  await sleep(300);
  const r = await http(U.adm, 'GET', '/api/admin/logs');
  return (Array.isArray(r.json) ? r.json : []).filter((l) => l.action === action);
}
const notifs = async (who, type) => { const j = (await http(who, 'GET', '/api/notifications')).json; return (Array.isArray(j) ? j : []).filter((n) => !type || n.type === type); };
const norm = (x) => String(x).replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/g, '<uuid>').replace(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/g, '<ts>');
const FORBIDDEN_TOKENS = [AMT, NOTE];
const clean = (res) => !FORBIDDEN_TOKENS.some((t) => res.text.includes(t));
async function newOpp(who, product, extra) {
  const r = await http(who, 'POST', '/api/opportunities', Object.assign({ company: '假公司' + crypto.randomBytes(3).toString('hex'), contactName: '假聯絡人', product, amount: AMT, description: NOTE, expectedDate: '2026-12-31', bu: 'ERP', category: 'ERP' }, extra || {}));
  if (r.status !== 201) throw new Error('建立商機失敗 ' + r.status + ' ' + r.text.slice(0, 200));
  return r.json;
}
const oppRaw = (d, id) => (d.opportunities || []).find((o) => o.id === id);
const itemsOf = (d, oppId) => (d.maItems || []).filter((i) => i.oppId === oppId);
const taskOf = (d, oppId) => (d.dealTasks || []).find((t) => t.oppId === oppId);

// ═════════════ 單元 ═════════════
function groupUnit() {
  const MA = REF.MA, OR = REAL.OR, PC = REAL.PC;
  console.log('[單元] lib/maTracking.js、lib/orgResolve.js、resolveTermRule');
  const cat = CAT();
  eq(MA.enabled({}), false, 'U1 沒設環境變數＝關');
  eq(MA.enabled({ MA_TRACKING: '1' }), true, 'U2 MA_TRACKING=1＝開');
  eq(MA.enabled({ MA_TRACKING: 'true' }), false, 'U3 只認 1');
  // resolveTermRule：與 resolveDealIdRule 同一套比對
  eq(PC.resolveTermRule(cat, 'Svc MA').termKind, 'ma', 'U4 名稱完全相符');
  eq(PC.resolveTermRule(cat, 'Svc MA 一年').startsTask, true, 'U5 名稱＋空白前綴；startsTask 取自目錄');
  eq(PC.resolveTermRule(cat, 'Svc MA Long x').name, 'Svc MA Long', 'U6 最長前綴優先（不是 Svc MA）');
  eq(PC.resolveTermRule(cat, 'Svc MA Long x').startsTask, false, 'U7 最長前綴的 startsTask');
  eq(PC.resolveTermRule(cat, 'Svc MAX').termKind, 'none', 'U8 沒有空白分隔不算前綴');
  eq(PC.resolveTermRule(cat, 'Plain 專案').termKind, 'none', 'U9 目錄內但無類型');
  eq(PC.resolveTermRule(cat, '不在目錄').termKind, 'none', 'U10 不在目錄');
  eq(PC.resolveTermRule(cat, '').termKind, 'none', 'U11 空字串');
  eq(PC.resolveTermRule(cat, 'Svc Lic 永久'), { termKind: 'license', startsTask: false, name: 'Svc Lic' }, 'U12 license 輸出');
  const dup = { bus: { ERP: [{ group: 'a', items: [{ name: 'X', termKind: 'license', startsTask: true }] }], ITS: [{ group: 'b', items: [{ name: 'X', termKind: 'ma' }] }] } };
  eq(PC.resolveTermRule(dup, 'X').termKind, 'ma', 'U13 同名多項：任一有類型即採用，ma 優先');
  eq(PC.resolveTermRule(dup, 'X').startsTask, false, 'U13b startsTask 取同類型那一項');
  // 日期（台北日曆日、與機器時區無關）
  eq(MA.taipeiDate('2026-10-24T15:59:59.000Z'), '2026-10-24', 'U14 UTC 15:59 ＝ 台北 23:59 同日');
  eq(MA.taipeiDate('2026-10-24T16:00:00.000Z'), '2026-10-25', 'U15 UTC 16:00 ＝ 台北隔天 00:00');
  eq(MA.daysLeft('2026-10-25', '2026-10-24T16:00:00.000Z'), 0, 'U16 到期日＝台北今天 → 0');
  eq(MA.daysLeft('2026-10-24', '2026-10-24T16:00:00.000Z'), -1, 'U17 台北昨天到期 → -1');
  eq(MA.daysLeft('2026-02-30', '2026-10-24T00:00:00.000Z'), null, 'U18 假日期 → null');
  eq(MA.addDaysIso('2026-12-30', 3), '2027-01-02', 'U19 跨年加天數');
  const tzOut = ['UTC', 'Pacific/Honolulu', 'Asia/Taipei', 'Pacific/Kiritimati'].map((tz) => spawnSync(process.execPath, ['-e', "const M=require(process.argv[1]);console.log(JSON.stringify([M.taipeiDate('2026-10-24T16:00:00.000Z'),M.daysLeft('2026-12-31','2026-10-24T16:00:00.000Z'),M.daysSinceIso('2026-10-01T00:00:00.000Z','2026-10-24T16:00:00.000Z')]))", path.join(ROOT, 'lib', 'maTracking.js')], { env: Object.assign({}, process.env, { TZ: tz }), encoding: 'utf8' }).stdout.trim());
  ok(new Set(tzOut).size === 1, 'U20 換四種機器時區輸出完全相同', tzOut.join(' | '));
  eq(MA.itemKey('o', 'ma', '2026-01-01', '2026-12-31'), 'o|ma|2026-01-01|2026-12-31', 'U21 冪等鍵');

  // planWon 期間來源優先序
  const D = () => seedData();
  let d = D();
  let p = MA.planWon(d, { id: 'oQ1', product: MAP }, { termFrom: '2026-07-01', termTo: '2027-06-30' });
  eq([p.ok, p.entries.map((e) => e.source), p.entries[0].from], [true, ['request'], '2026-07-01'], 'U22 請求帶入優先於報價單');
  p = MA.planWon(d, { id: 'oQ1', product: MAP }, {});
  eq([p.ok, p.entries.length, p.entries.map((e) => e.termKind + e.from + e.to).sort()], [true, 2, ['license2026-03-012029-02-28', 'ma2026-03-012027-02-28']], 'U23 報價單：最近一張已核准、重複期間併成一組、一般品項略過、簽核中的不算');
  eq(p.entries[0].quoteId, 'q1b', 'U24 引用最近核准的 q1b（不是更舊的 q1a）');
  p = MA.planWon(d, { id: 'oQ3', product: MAP }, {});
  eq([p.ok, p.code], [false, 'TERM_REQUIRED'], 'U25 報價單品項期間不合法 → 視同沒有');
  p = MA.planWon(d, { id: 'oQ4', product: MAP }, {});
  eq([p.ok, p.code], [false, 'TERM_REQUIRED'], 'U26 已失效報價單不算');
  p = MA.planWon(d, { id: 'oX', product: MAP }, { termFrom: '2026-07-01' });
  eq([p.ok, p.code], [false, 'TERM_REQUIRED'], 'U27 只填起日 → TERM_REQUIRED');
  p = MA.planWon(d, { id: 'oX', product: MAP }, { termFrom: '2026-07-01', termTo: '2026-06-30' });
  eq([p.ok, p.code], [false, 'TERM_INVALID'], 'U28 迄日早於起日 → TERM_INVALID');
  p = MA.planWon(d, { id: 'oX', product: MAP }, { termFrom: '2026-02-30', termTo: '2026-12-31' });
  eq([p.ok, p.code], [false, 'TERM_INVALID'], 'U29 非真實日期 → TERM_INVALID');
  p = MA.planWon(d, { id: 'oX', product: MAP }, { termFrom: ' 2026-07-01 ', termTo: '2026-07-01' });
  eq([p.ok, p.entries && p.entries[0].from], [true, '2026-07-01'], 'U30 前後空白 trim、起迄同日可');
  p = MA.planWon(d, { id: 'oX', product: PLAINP }, {});
  eq([p.ok, p.track], [true, false], 'U31 非 License／MA 商品不納管（即使沒期間）');
  p = MA.planWon(d, { id: 'oX', product: PLAINP }, { termFrom: 'garbage', termTo: 'x' });
  eq([p.ok, p.track], [true, false], 'U32 非 License／MA 商品忽略請求裡的期間');
  p = MA.planWon(d, { id: 'oX', product: MAP }, { termFrom: '2026-07-01', termTo: '2026-12-31', amount: 'x' });
  ok(p.ok, 'U33 多餘欄位不影響');

  // applyWon 冪等／狀態機
  d = D();
  const opp = { id: 'oZ', owner: U.u1, company: '甲', product: MAP, bu: 'ERP' };
  const plan = MA.planWon(d, opp, { termFrom: '2026-01-01', termTo: '2026-12-31' });
  const c1 = { by: U.u1, nowIso: '2026-10-24T00:00:00.000Z' };
  let r = MA.applyWon(d, opp, plan, c1);
  eq([r.createdItems.length, r.taskAction, d.maItems.filter((i) => i.oppId === 'oZ').length, d.dealTasks.filter((t) => t.oppId === 'oZ').length], [1, 'created', 1, 1], 'U34 首次成交：1 納管項＋1 任務卡');
  r = MA.applyWon(d, opp, plan, c1);
  eq([r.createdItems.length, r.revivedItems.length, r.taskAction, d.maItems.filter((i) => i.oppId === 'oZ').length, d.dealTasks.filter((t) => t.oppId === 'oZ').length], [0, 0, '', 1, 1], 'U35 重複套用：不重複建（冪等）');
  const un = MA.applyUnwon(d, 'oZ', c1, 'STAGE_REVERTED');
  eq([un.cancelledTasks.length, un.voidedItems.length, taskOf(d, 'oZ').status, itemsOf(d, 'oZ')[0].status], [1, 1, 'cancelled', 'void'], 'U36 離開 Won：任務卡 cancelled、納管項 void');
  ok(itemsOf(d, 'oZ')[0].history.length >= 2 && taskOf(d, 'oZ').history.length >= 2, 'U37 history 保留');
  const p2 = MA.planWon(d, opp, {});
  eq([p2.ok, p2.entries[0].source], [true, 'history'], 'U38 再次成交沒帶期間：沿用先前作廢項的期間');
  r = MA.applyWon(d, opp, p2, c1);
  eq([r.revivedItems.length, r.taskAction, d.maItems.filter((i) => i.oppId === 'oZ').length, taskOf(d, 'oZ').status], [1, 'reopened', 1, 'open'], 'U39 再次成交：重開不重複建');
  eq(MA.applyUnwon(d, 'oZ', c1, 'X').cancelledTasks.length, 1, 'U40 再退回再取消');
  const p3 = MA.planWon(d, opp, { termFrom: '2027-01-01', termTo: '2027-12-31' });
  r = MA.applyWon(d, opp, p3, c1);
  eq([d.maItems.filter((i) => i.oppId === 'oZ').length, d.maItems.filter((i) => i.oppId === 'oZ' && i.status === 'tracking').length, r.taskAction], [2, 1, 'reopened'], 'U41 再次成交帶新期間：新增一項，舊作廢項保留');
  // 無 startsTask 的商品：只有納管項
  d = D();
  const optL = { id: 'oL', owner: U.u1, company: '乙', product: LICP };
  r = MA.applyWon(d, optL, MA.planWon(d, optL, { termFrom: '2026-01-01', termTo: '2026-12-31' }), c1);
  eq([r.createdItems.length, r.task, (d.dealTasks || []).filter((t) => t.oppId === 'oL').length], [1, null, 0], 'U42 startsTask=false：有納管項、沒有任務卡');

  // signTask
  d = D();
  const o2 = { id: 'oS', owner: U.u1, company: '丙公司', product: MAP, amount: '500', contactName: '王', bu: 'ERP' };
  d.opportunities.push(o2);
  r = MA.applyWon(d, o2, MA.planWon(d, o2, { termFrom: '2026-01-01', termTo: '2026-12-31' }), c1);
  const tid = r.task.id;
  const bad = (inp) => MA.signTask(d, tid, inp, { by: U.u1, nowIso: c1.nowIso, salesPerson: '甲' });
  eq([bad({}).code, bad({ contractNo: '   ', signedDate: '2026-02-01' }).code, bad({ contractNo: 'A', signedDate: '' }).code, bad({ contractNo: 'A', signedDate: '2026-02-30' }).code], ['CONTRACT_NO_REQUIRED', 'CONTRACT_NO_REQUIRED', 'SIGNED_DATE_REQUIRED', 'SIGNED_DATE_INVALID'], 'U43 簽約確認必填檢查');
  eq(bad({ contractNo: 'x'.repeat(65), signedDate: '2026-02-01' }).code, 'CONTRACT_NO_INVALID', 'U44 合約編號 65 字被拒');
  eq(bad({ contractNo: '字'.repeat(64), signedDate: '2026-02-01' }).ok, true, 'U45 合約編號 64 個字可（以字元數計）');
  d = D(); d.opportunities.push(o2);
  r = MA.applyWon(d, o2, MA.planWon(d, o2, { termFrom: '2026-01-01', termTo: '2026-12-31' }), c1);
  const t2 = r.task.id;
  eq(MA.signTask(d, t2, { contractNo: 'leg-001', signedDate: '2026-02-01' }, { by: U.u1, nowIso: c1.nowIso }).code, 'DUPLICATE_CONTRACT_NO', 'U46 與既有合約同編號（不分大小寫）→ 重複');
  eq(MA.signTask(d, t2, { contractNo: '  LEG-001 ', signedDate: '2026-02-01' }, { by: U.u1, nowIso: c1.nowIso }).code, 'DUPLICATE_CONTRACT_NO', 'U47 前後空白 trim 後比對');
  ok(!JSON.stringify(MA.signTask(d, t2, { contractNo: 'LEG-001', signedDate: '2026-02-01' }, { by: U.u1, nowIso: c1.nowIso })).includes(SECRET_CO), 'U48 重複訊息不含他人客戶名');
  const sg = MA.signTask(d, t2, { contractNo: 'GONE-001', signedDate: '2026-02-01' }, { by: U.u1, nowIso: c1.nowIso, salesPerson: '甲' });
  eq(sg.ok, true, 'U49 已刪除合約的編號可重用');
  eq([sg.contract.type, sg.contract.startDate, sg.contract.endDate, sg.contract.owner, sg.contract.oppId, sg.contract.maItemId === sg.items[0].id, sg.contract.company, sg.contract.amount], ['SERVICE_MA', '2026-01-01', '2026-12-31', U.u1, 'oS', true, '丙公司', '500'], 'U50 合約欄位取自納管項與商機');
  eq([d.dealTasks.find((t) => t.id === t2).status, sg.items[0].contractNo, sg.items[0].contractId === sg.contract.id], ['done', 'GONE-001', true], 'U51 任務完成、納管項回寫合約');
  eq(MA.signTask(d, t2, { contractNo: 'ZZ', signedDate: '2026-02-01' }, { by: U.u1, nowIso: c1.nowIso }).code, 'TASK_NOT_OPEN', 'U52 已完成不能再簽');
  eq(MA.contractNoTaken(d, 'gone-001'), true, 'U53 簽完後該編號被占用');
  // skip / snooze / reassign
  d = D(); d.opportunities.push(o2);
  r = MA.applyWon(d, o2, MA.planWon(d, o2, { termFrom: '2026-01-01', termTo: '2026-12-31' }), c1);
  const t3 = r.task.id;
  const users = seedUsers();
  eq([MA.skipTask(d, t3, { reason: '' }, c1, users).code, MA.skipTask(d, t3, { reason: '一' }, c1, users).code, MA.skipTask(d, t3, { reason: '字'.repeat(201) }, c1, users).code], ['SKIP_REASON_REQUIRED', 'SKIP_REASON_REQUIRED', 'SKIP_REASON_TOO_LONG'], 'U54 略過原因 ≥2 字、≤200 字');
  const sn = MA.snoozeTask(d, t3, { by: U.u1, nowIso: '2026-10-24T16:30:00.000Z' });
  eq(sn.snoozedUntil, '2026-10-28', 'U55 延後＝台北日期 +3 天（UTC 16:30 已是台北隔天）');
  const sk = MA.skipTask(d, t3, { reason: '客戶自行處理' }, { by: U.u1, nowIso: c1.nowIso }, users);
  eq([sk.ok, sk.notify.sort()], [true, [U.m1, U.mkt, U.mkt2].sort()], 'U56 略過通知：在職行銷＋一級主管（u1→m2→m1），不含離職行銷與操作者');
  eq([d.dealTasks.find((t) => t.id === t3).status, itemsOf(d, 'oS')[0].status], ['skipped', 'tracking'], 'U57 任務 skipped、納管項保留追蹤');
  eq(MA.skipTask(d, t3, { reason: '客戶自行處理' }, c1, users).code, 'TASK_NOT_OPEN', 'U58 略過後不能再略過');
  d = D(); d.opportunities.push(o2);
  r = MA.applyWon(d, o2, MA.planWon(d, o2, { termFrom: '2026-01-01', termTo: '2026-12-31' }), c1);
  const t4 = r.task.id;
  eq([MA.reassignTask(d, t4, U.mktoff, c1, users).code, MA.reassignTask(d, t4, U.mkt, c1, users).code, MA.reassignTask(d, t4, 'nobody', c1, users).code, MA.reassignTask(d, t4, U.u1, c1, users).code, MA.reassignTask(d, t4, U.uoff, c1, users).code], ['ASSIGNEE_INACTIVE', 'ASSIGNEE_ROLE', 'ASSIGNEE_NOT_FOUND', 'ASSIGNEE_SAME', 'ASSIGNEE_INACTIVE'], 'U59 改派目標檢查（停用／角色／不存在／同人）');
  eq([MA.reassignTask(d, t4, U.u2, c1, users).ok, d.dealTasks.find((t) => t.id === t4).owner, itemsOf(d, 'oS')[0].owner], [true, U.u2, U.u2], 'U60 改派：任務與納管項負責人一起換');
  // 改派候選人與驗證（驗收後修正）：只列在職業務（user）、BU 與任務卡負責人有交集；負責人已停用 → 以商機 BU 為準
  const cand = MA.reassignCandidates(d, d.dealTasks.find((t) => t.id === t4), users).map((x) => x.username).sort();
  eq(cand, [U.u1, U.u4, U.u5].sort(), 'U60b 候選人＝ERP 在職業務（現任負責人 u2 排除）；不含主管／行銷／秘書／集團／ITS／離職');
  const d2 = D(); const o3 = { id: 'oS3', owner: U.u1, company: '己公司', product: MAP, amount: '5', bu: 'ERP' }; d2.opportunities.push(o3);
  const tt = MA.applyWon(d2, o3, MA.planWon(d2, o3, { termFrom: '2026-01-01', termTo: '2026-12-31' }), c1).task;
  eq([MA.reassignTask(d2, tt.id, U.m1, c1, users).code, MA.reassignTask(d2, tt.id, U.m2, c1, users).code, MA.reassignTask(d2, tt.id, U.gs, c1, users).code, MA.reassignTask(d2, tt.id, U.cs, c1, users).code], ['ASSIGNEE_ROLE', 'ASSIGNEE_ROLE', 'ASSIGNEE_ROLE', 'ASSIGNEE_ROLE'], 'U60c 改派給一級／二級主管、集團業務、顧問主管 → ASSIGNEE_ROLE');
  eq(MA.reassignTask(d2, tt.id, U.u3, c1, users).code, 'ASSIGNEE_BU', 'U60d 改派給別的 BU 業務 → ASSIGNEE_BU（驗證在後端）');
  eq(d2.dealTasks.find((t) => t.id === tt.id).owner, U.u1, 'U60e 被拒絕時負責人不變');
  const d3 = D(); const o4 = { id: 'oS4', owner: U.uoff, company: '庚公司', product: MAP, amount: '5', bu: 'ITS' }; d3.opportunities.push(o4);
  const t5 = MA.applyWon(d3, o4, MA.planWon(d3, o4, { termFrom: '2026-01-01', termTo: '2026-12-31' }), c1).task;
  eq(MA.reassignCandidates(d3, t5, users).map((x) => x.username), [U.u3], 'U60f 負責人已停用 → 以商機 BU（ITS）為準，只剩 u3');
  eq([MA.reassignTask(d3, t5.id, U.u2, c1, users).code, MA.reassignTask(d3, t5.id, U.u3, c1, users).ok], ['ASSIGNEE_BU', true], 'U60g 同上規則套用在改派驗證');
  // 再成交期間提示
  const hint = MA.termHint([{ termKind: 'ma', from: '2020-01-01', to: '2020-12-31', status: 'void', updatedAt: '2026-01-01T00:00:00Z' }, { termKind: 'ma', from: '2026-01-01', to: '2099-12-31', status: 'void', updatedAt: '2026-05-01T00:00:00Z' }, { termKind: 'ma', from: '2026-01-01', to: '2099-12-31', status: 'tracking', updatedAt: '2026-04-01T00:00:00Z' }], new Date());
  eq([hint.length, hint[0].to, hint[0].expired, hint[1].expired], [2, '2099-12-31', false, true], 'U60m termHint：最近異動在前、相同期間只列一次、過期者 expired');
  eq(MA.termHint([], new Date()), [], 'U60n termHint 空');
  // 合約刪除重開（驗收後修正）
  const d4 = D(); const o5 = { id: 'oS5', owner: U.u1, company: '辛公司', product: MAP, amount: '7', bu: 'ERP' }; d4.opportunities.push(o5);
  const w5 = MA.applyWon(d4, o5, MA.planWon(d4, o5, { termFrom: '2026-01-01', termTo: '2026-12-31' }), c1);
  const sg5 = MA.signTask(d4, w5.task.id, { contractNo: 'DEL-1', signedDate: '2026-02-01' }, c1);
  sg5.contract.deleted = true;
  const cdx = MA.applyContractDeleted(d4, sg5.contract, { by: U.u1, nowIso: '2026-10-25T00:00:00.000Z' });
  const t5b = d4.dealTasks.find((t) => t.id === w5.task.id), i5 = d4.maItems.find((i) => i.oppId === 'oS5');
  eq([cdx.touched, !!cdx.task, t5b.status, t5b.steps[0].status, t5b.history.slice(-1)[0].action, t5b.history.slice(-1)[0].note, i5.contractNo, i5.contractId, i5.history.slice(-1)[0].action], [true, true, 'open', 'pending', 'REOPEN', '合約已刪除，重開簽約', '', '', 'CONTRACT_DELETED'], 'U60h 刪除合約：任務卡重開（簽約步驟回 pending、history 記錄）、納管項清掉合約連結');
  eq(MA.listItems(d4.maItems, users, new Date(), { filter: 'pending_sign', tasks: d4.dealTasks }).filter((x) => x.oppId === 'oS5').length, 1, 'U60i 重開後納管項重新出現在待簽約');
  const cdy = MA.applyContractDeleted(d4, sg5.contract, { by: U.u1, nowIso: '2026-10-26T00:00:00.000Z' });
  eq([cdy.touched, cdy.task, t5b.history.filter((h) => h.action === 'REOPEN').length], [false, null, 1], 'U60j 重複套用冪等（不再重開、history 不重複）');
  eq(MA.signTask(d4, t5b.id, { contractNo: 'DEL-1', signedDate: '2026-03-01' }, c1).ok, true, 'U60k 重開後可用同編號重新簽約');
  const d5 = D(); const o6 = { id: 'oS6', owner: U.u1, company: '壬公司', product: LICP, amount: '7', bu: 'ERP' }; d5.opportunities.push(o6);
  const w6 = MA.applyWon(d5, o6, MA.planWon(d5, o6, { termFrom: '2026-01-01', termTo: '2026-12-31' }), c1);
  const fakeC = { id: 'cx6', contractNo: 'LIC-1', deleted: true };
  d5.maItems.find((i) => i.oppId === 'oS6').contractId = 'cx6'; d5.maItems.find((i) => i.oppId === 'oS6').contractNo = 'LIC-1';
  const cd6 = MA.applyContractDeleted(d5, fakeC, { by: U.u1, nowIso: c1.nowIso });
  eq([w6.task, cd6.touched, cd6.task, (d5.dealTasks || []).filter((t) => t.oppId === 'oS6').length, d5.maItems.find((i) => i.oppId === 'oS6').contractNo], [null, true, null, 0, ''], 'U60l 沒有任務卡的納管項：只清連結，不建任務卡');
  // 通知已讀
  d = D(); d.opportunities.push(o2);
  r = MA.applyWon(d, o2, MA.planWon(d, o2, { termFrom: '2026-01-01', termTo: '2026-12-31' }), c1);
  d.notifications = [{ id: 'n1', to: U.u1, type: 'deal_task', refId: r.task.id, read: false }, { id: 'n2', to: U.u1, type: 'deal_task', refId: 'other', read: false }, { id: 'n3', to: U.mkt, type: 'deal_task_skipped', refId: r.task.id, read: false }];
  MA.signTask(d, r.task.id, { contractNo: 'NEW-1', signedDate: '2026-02-01' }, { by: U.u1, nowIso: c1.nowIso });
  eq(d.notifications.map((n) => n.read), [true, false, false], 'U61 完成只標該任務的 deal_task 通知為已讀');
  // editTerm
  d = D(); d.opportunities.push(o2);
  r = MA.applyWon(d, o2, MA.planWon(d, o2, { termFrom: '2026-01-01', termTo: '2026-12-31' }), c1);
  const iid = r.items[0].id;
  eq([MA.editTerm(d, iid, { termFrom: '2026-12-31', termTo: '2026-01-01' }, c1).code, MA.editTerm(d, iid, { termFrom: '', termTo: '' }, c1).code, MA.editTerm(d, 'nope', {}, c1).code], ['TERM_INVALID', 'TERM_REQUIRED', 'ITEM_NOT_FOUND'], 'U62 修改期間驗證');
  const ed = MA.editTerm(d, iid, { termFrom: '2026-02-01', termTo: '2027-01-31', reason: '延後啟用' }, c1);
  eq([ed.ok, ed.changed, d.maItems.find((i) => i.id === iid).to, d.maItems.find((i) => i.id === iid).history.slice(-1)[0].action], [true, true, '2027-01-31', 'TERM_EDIT'], 'U63 修改期間＋history');
  eq(MA.editTerm(d, iid, { termFrom: '2026-02-01', termTo: '2027-01-31' }, c1).changed, false, 'U64 沒變不寫 history');
  // 檢視：listItems / 秘書白名單 / 無金額
  d = D();
  const its = [
    { id: 'a', oppId: 'o1', owner: U.u1, company: '甲', product: 'P', termKind: 'ma', from: '2026-01-01', to: plusDays(todayTp(), 90), contractNo: 'C1', status: 'tracking', amount: AMT, note: NOTE },
    { id: 'b', oppId: 'o2', owner: U.u2, company: '乙', product: 'P', termKind: 'license', from: '2026-01-01', to: plusDays(todayTp(), 91), contractNo: '', status: 'tracking' },
    { id: 'c', oppId: 'o3', owner: U.u1, company: '丙', product: 'P', termKind: 'ma', from: '2025-01-01', to: plusDays(todayTp(), -1), contractNo: '', status: 'tracking' },
    { id: 'd', oppId: 'o4', owner: U.u1, company: '丁', product: 'P', termKind: 'ma', from: '2025-01-01', to: todayTp(), contractNo: '', status: 'void' },
    { id: 'e', oppId: 'o5', owner: U.u1, company: '戊', product: 'P', termKind: 'ma', from: '2025-01-01', to: plusDays(todayTp(), 0), contractNo: 'C5', status: 'ended' },
  ];
  const ids = (f) => MA.listItems(its, seedUsers(), new Date(), { filter: f }).map((x) => x.id);
  eq(ids('all'), ['c', 'e', 'a', 'b'], 'U65 all：排除 void，依到期日排序');
  eq(ids('d90'), ['a'], 'U66 d90：0～90 天內（91 天不算；ended 不算）');
  eq(ids('overdue'), ['c'], 'U67 overdue：已過期');
  const tk = [{ id: 'tb', type: 'ma_sign', maItemIds: ['b'], status: 'open' }, { id: 'tc', type: 'ma_sign', maItemIds: ['c'], status: 'skipped' }];
  const idsT = (f) => MA.listItems(its, seedUsers(), new Date(), { filter: f, tasks: tk }).map((x) => x.id);
  eq(idsT('pending_sign'), ['b'], 'U68 pending_sign：追蹤中、沒有合約編號、而且有開啟中的簽約任務（略過的不算）');
  eq(ids('pending_sign'), [], 'U68b 沒有任務卡的納管項（License MA 類）不算待簽約');
  const pv = MA.listItems(its, seedUsers(), new Date(), { filter: 'all', tasks: tk });
  eq(['a', 'b', 'c', 'e'].map((k) => pv.find((x) => x.id === k).pendingSign), [false, true, false, false], 'U68c pendingSign 欄位：只有 b（有開啟任務）為真；沒任務／略過／已簽／已結束都不是');
  eq(MA.isPendingSign({ status: 'tracking', contractNo: '' }, { status: 'open' }), true, 'U68d isPendingSign 開啟任務');
  eq([MA.isPendingSign({ status: 'tracking', contractNo: '' }, null), MA.isPendingSign({ status: 'tracking', contractNo: '' }, { status: 'done' }), MA.isPendingSign({ status: 'tracking', contractNo: 'X' }, { status: 'open' })], [false, false, false], 'U68e isPendingSign 無任務／任務已完成／已有合約編號');
  eq(ids('bogus'), ids('all'), 'U69 未知 filter 視同 all');
  const sec = MA.listItems(its, seedUsers(), new Date(), { filter: 'all', secretary: true });
  ok(sec.every((x) => JSON.stringify(Object.keys(x).sort()) === JSON.stringify(['company', 'daysLeft', 'from', 'ownerLabel', 'product', 'termKind', 'to'])), 'U70 秘書只有 7 個欄位', JSON.stringify(sec[0]));
  const full = JSON.stringify(MA.listItems(its, seedUsers(), new Date(), { filter: 'all' }));
  ok(!full.includes(AMT) && !full.includes(NOTE), 'U71 一般檢視也不含金額與備註');
  eq(MA.listItems(its, seedUsers(), new Date(), {}).find((x) => x.id === 'a').contractNo, 'C1', 'U72 非秘書含合約編號');
  // 待補納管
  d = D();
  const pb = MA.pendingBackfill(d, d.opportunities, seedUsers());
  eq(pb.map((x) => x.oppId).sort(), ['oBF'], 'U73 待補納管：已成交、License／MA 商品、沒有（非作廢）納管項');
  ok(!JSON.stringify(pb).includes(AMT), 'U74 待補納管不含金額');
  // 容量
  d = D(); d.maItems = []; for (let i = 0; i < MA.MAX_ITEMS; i++) d.maItems.push({ id: 'z' + i, oppId: 'z' + i, status: 'tracking', termKind: 'ma', from: 'a', to: 'b', createdAt: 'x' });
  eq(MA.planWon(d, { id: 'oNew', product: MAP }, { termFrom: '2026-01-01', termTo: '2026-12-31' }).code, 'MA_CAPACITY', 'U75 納管項滿額且無終態項可清 → MA_CAPACITY');
  d.maItems[0].status = 'void';
  const pc = MA.planWon(d, { id: 'oNew', product: MAP }, { termFrom: '2026-01-01', termTo: '2026-12-31' });
  ok(pc.ok, 'U76 有終態項可清 → 允許');
  const ro = { id: 'oNew', owner: U.u1, company: 'N', product: MAP };
  MA.applyWon(d, ro, pc, c1);
  ok(d.maItems.length <= MA.MAX_ITEMS, 'U77 套用後不超過上限（清掉最舊的終態項）', d.maItems.length);
  // orgResolve
  const us = OR.toMap(seedUsers());
  eq([OR.findManager1(us, U.u1), OR.findManager1(us, U.u2), OR.findManager1(us, U.u3)], [U.m1, U.m1, U.m1b], 'U78 一級主管沿 supervisor 找（u1→m2→m1）');
  eq(OR.findManager1(us, U.u4), U.m1, 'U79 掛名主管（pureSupervisor）跳過、再往上找下一位 manager1');
  eq(OR.findManager1(us, U.u5), U.m1, 'U80 停用的一級主管跳過、往上找');
  eq(OR.findManager1(us, U.m1b), null, 'U81 沒有上級 → null（不退回 BU）');
  const loop = OR.toMap([{ username: 'a', role: 'user', supervisor: 'b' }, { username: 'b', role: 'user', supervisor: 'a' }]);
  eq(OR.findManager1(loop, 'a'), null, 'U82 supervisor 循環不會無窮迴圈');
  eq(OR.marketingUsers(seedUsers()), [U.mkt, U.mkt2].sort(), 'U83 在職行銷（排除離職與其他角色）');
  eq(OR.findManager1(seedUsers(), U.u1), U.m1, 'U84 也接受帳號陣列');
}

// ═════════════ API ═════════════
async function groupApi() {
  console.log('[API] 成交掛鉤、任務卡、合約串接、權限');
  await freshServer();
  await loginAll([U.adm, U.u1, U.u2, U.m1, U.m1b, U.m2, U.mkt, U.mkt2, U.sec]);   // 登入限流每行程 20 次，只登入會用到的帳號
  const snapOpp = (id) => JSON.stringify(oppRaw(readData(), id));

  // A1 缺期間：三種路徑一致、400、無副作用
  const a1 = await newOpp(U.u1, MAP), a2 = await newOpp(U.u1, MAP), a3 = await newOpp(U.u1, MAP);
  const before = [a1, a2, a3].map((o) => snapOpp(o.id));
  const r1 = await http(U.u1, 'PUT', `/api/opportunities/${a1.id}`, { stage: 'Won' });                                   // 看板拖曳
  const full = (o) => ({ company: o.company, contactName: o.contactName, product: o.product, amount: o.amount, expectedDate: o.expectedDate, description: o.description, stage: 'Won', category: o.category, bu: o.bu });
  const r2 = await http(U.u1, 'PUT', `/api/opportunities/${a2.id}`, full(a2));                                           // 編輯表單
  const r3 = await http(U.u1, 'PUT', `/api/opportunities/${a3.id}`, { stage: 'Won', product: MAP });                    // 直接打 API
  eq([r1.status, r2.status, r3.status], [400, 400, 400], 'A1 三種路徑缺期間都 400');
  eq([r1.json.code, r2.json.code, r3.json.code], ['TERM_REQUIRED', 'TERM_REQUIRED', 'TERM_REQUIRED'], 'A1b 三種路徑同一個錯誤碼');
  ok(r1.json.error === r2.json.error && r2.json.error === r3.json.error && /服務期間/.test(r1.json.error), 'A1c 三種路徑同一句中文訊息');
  const dA = readData();
  eq([a1, a2, a3].map((o) => snapOpp(o.id)), before, 'A1d 缺期間時商機一個位元都沒變（無 stageHistory、無寫入）');
  ok(!(dA.maItems || []).some((i) => [a1, a2, a3].some((o) => o.id === i.oppId)) && !(dA.dealTasks || []).some((t) => [a1, a2, a3].some((o) => o.id === t.oppId)), 'A1e 沒有建立納管項／任務卡');
  eq((await notifs(U.u1, 'deal_task')).length, 0, 'A1f 沒有通知');
  // A2 不合法期間
  for (const [label, body, code] of [['只填起日', { termFrom: '2026-01-01' }, 'TERM_REQUIRED'], ['顛倒', { termFrom: '2026-12-31', termTo: '2026-01-01' }, 'TERM_INVALID'], ['假日期', { termFrom: '2026-02-30', termTo: '2026-12-31' }, 'TERM_INVALID'], ['亂字', { termFrom: 'abc', termTo: 'def' }, 'TERM_INVALID']]) {
    const x = await http(U.u1, 'PUT', `/api/opportunities/${a1.id}`, Object.assign({ stage: 'Won' }, body));
    eq([x.status, x.json && x.json.code], [400, code], 'A2 ' + label + ' → ' + code);
  }
  eq(snapOpp(a1.id), before[0], 'A2b 不合法期間也沒有副作用');
  // A3 三種路徑成功且結果一致
  const T = { termFrom: '2026-01-01', termTo: '2026-12-31' };
  const s1 = await http(U.u1, 'PUT', `/api/opportunities/${a1.id}`, Object.assign({ stage: 'Won' }, T));
  const s2 = await http(U.u1, 'PUT', `/api/opportunities/${a2.id}`, Object.assign(full(a2), T));
  const s3 = await http(U.u1, 'PUT', `/api/opportunities/${a3.id}`, Object.assign({ stage: 'Won', product: MAP }, T));
  eq([s1.status, s2.status, s3.status], [200, 200, 200], 'A3 三種路徑成交成功');
  eq([s1.json.ma && s1.json.ma.taskAction, s2.json.ma && s2.json.ma.taskAction, s3.json.ma && s3.json.ma.taskAction], ['created', 'created', 'created'], 'A3b 回應 ma.taskAction=created');
  ok(!('termFrom' in s1.json) && s1.json.stage === 'Won', 'A3c 期間欄位不寫進商機本體');
  let d = readData();
  const shape = (o) => { const it = itemsOf(d, o.id)[0], t = taskOf(d, o.id); return JSON.stringify([itemsOf(d, o.id).length, it.termKind, it.from, it.to, it.status, it.owner, it.source, it.contractNo, t.status, t.owner, t.type, t.steps.map((s) => s.key + s.status), t.maItemIds.length, (oppRaw(d, o.id).stageHistory || []).map((h) => h.from + '>' + h.to)]); };
  ok(shape(a1) === shape(a2) && shape(a2) === shape(a3), 'A3d 三種路徑建立的納管項與任務卡內容相同', shape(a1) + ' / ' + shape(a2));
  eq(shape(a1), JSON.stringify([1, 'ma', '2026-01-01', '2026-12-31', 'tracking', U.u1, 'request', '', 'open', U.u1, 'ma_sign', ['signpending'], 1, ['C>Won']]), 'A3e 納管項／任務卡欄位');
  const nt = await notifs(U.u1, 'deal_task');
  eq(nt.length, 3, 'A3f 每張任務卡一則通知給負責業務');
  ok(nt.every((n) => !n.title.includes(AMT) && !n.body.includes(AMT) && !n.body.includes(NOTE)) && nt.some((n) => n.title.includes(a1.company)), 'A3g 通知只有公司名與品項、沒有金額');
  eq(nt.map((n) => n.refId).sort(), [taskOf(d, a1.id).id, taskOf(d, a2.id).id, taskOf(d, a3.id).id].sort(), 'A3h 通知 refId＝任務卡 id');
  eq((await auditLines('MA_TASK_CREATE')).length, 3, 'A3i 稽核 MA_TASK_CREATE');
  eq((await auditLines('MA_ITEM_CREATE')).length, 3, 'A3j 稽核 MA_ITEM_CREATE');
  // A4 非類型商品／無 startsTask
  const pl = await newOpp(U.u1, PLAINP), li = await newOpp(U.u1, LICP);
  const x1 = await http(U.u1, 'PUT', `/api/opportunities/${pl.id}`, { stage: 'Won' });
  eq([x1.status, 'ma' in x1.json], [200, false], 'A4 非 License／MA 商品成交：不要求期間、回應沒有 ma');
  const x2 = await http(U.u1, 'PUT', `/api/opportunities/${li.id}`, Object.assign({ stage: 'Won' }, T));
  d = readData();
  eq([x2.status, itemsOf(d, li.id).length, !!taskOf(d, li.id), x2.json.ma && x2.json.ma.taskId], [200, 1, false, ''], 'A4b license（startsTask=false）：有納管項、沒任務卡');
  eq([itemsOf(d, pl.id).length, !!taskOf(d, pl.id)], [0, false], 'A4c 非類型商品沒有納管項');
  {
    const lAll = (await http(U.u1, 'GET', '/api/ma-items?filter=all')).json.items, lPend = (await http(U.u1, 'GET', '/api/ma-items?filter=pending_sign')).json.items;
    const liIt = itemsOf(d, li.id)[0], a1It = itemsOf(d, a1.id)[0];
    const row = lAll.find((i) => i.id === liIt.id);
    eq([row.pendingSign, row.status, row.taskId, lPend.some((i) => i.id === liIt.id), lPend.some((i) => i.id === a1It.id), lAll.find((i) => i.id === a1It.id).pendingSign], [false, 'tracking', '', false, true, true], 'A4d License MA（無任務卡）：追蹤中、不是待簽約、不在 pending_sign 篩選；Service MA（有開啟任務）是待簽約');
  }
  // A5 報價單來源
  const q1 = await http(U.u1, 'PUT', '/api/opportunities/oQ1', { stage: 'Won' });
  d = readData();
  eq([q1.status, itemsOf(d, 'oQ1').map((i) => i.termKind + i.from + i.to).sort(), itemsOf(d, 'oQ1').every((i) => i.source === 'quote' && i.quoteId === 'q1b'), taskOf(d, 'oQ1').maItemIds.length], [200, ['license2026-03-012029-02-28', 'ma2026-03-012027-02-28'], true, 2], 'A5 報價單帶入：最近核准那張、依 distinct 期間 2 項、同一張任務卡');
  const q2 = await http(U.u1, 'PUT', '/api/opportunities/oQ2', Object.assign({ stage: 'Won' }, T));
  d = readData();
  eq([q2.status, itemsOf(d, 'oQ2').map((i) => i.from + i.to + i.source)], [200, ['2026-01-012026-12-31request']], 'A5b 請求帶入優先於報價單');
  const q3 = await http(U.u1, 'PUT', '/api/opportunities/oQ3', { stage: 'Won' });
  const q4 = await http(U.u1, 'PUT', '/api/opportunities/oQ4', { stage: 'Won' });
  eq([q3.status, q3.json.code, q4.status, q4.json.code], [400, 'TERM_REQUIRED', 400, 'TERM_REQUIRED'], 'A5c 期間不合法／已失效的報價單不算 → 缺期間 400');
  // A6 冪等與回退
  const back = await http(U.u1, 'PUT', `/api/opportunities/${a1.id}`, { stage: 'A' });
  d = readData();
  eq([back.status, back.json.ma && back.json.ma.cancelledTasks, back.json.ma && back.json.ma.voidedItems, taskOf(d, a1.id).status, itemsOf(d, a1.id)[0].status], [200, 1, 1, 'cancelled', 'void'], 'A6 Won→A：任務卡取消、納管項作廢');
  const nA1 = (await notifs(U.u1, 'deal_task')).filter((n) => n.refId === taskOf(d, a1.id).id);
  ok(nA1.length === 1 && nA1[0].read === true, 'A6b 取消後該任務通知標為已讀', JSON.stringify(nA1));
  ok(itemsOf(d, a1.id)[0].history.length >= 2, 'A6c 作廢保留 history');
  {
    const th = await http(U.u1, 'GET', '/api/ma-items/term-hint?oppId=' + a1.id);
    eq([th.status, (th.json.terms || []).map((x) => x.termKind + x.from + x.to), (th.json.terms || [])[0] && th.json.terms[0].expired], [200, ['ma2026-01-012026-12-31'], '2026-12-31' < todayTp()], 'A6c2 再成交期間提示：回傳先前（含作廢）納管項的期間，expired 依台北今天判斷');
    ok(clean(th), 'A6c3 期間提示不含金額與備註');
    eq([(await http(U.u2, 'GET', '/api/ma-items/term-hint?oppId=' + a1.id)).json.terms.length, (await http(U.sec, 'GET', '/api/ma-items/term-hint?oppId=' + a1.id)).json.terms.length, (await http(U.u1, 'GET', '/api/ma-items/term-hint')).json.terms.length], [0, 0, 0], 'A6c4 別的業務／秘書／沒給 oppId 都是空');
  }
  const again = await http(U.u1, 'PUT', `/api/opportunities/${a1.id}`, { stage: 'Won' });
  d = readData();
  eq([again.status, itemsOf(d, a1.id).length, (d.dealTasks || []).filter((t) => t.oppId === a1.id).length, itemsOf(d, a1.id)[0].status, taskOf(d, a1.id).status, again.json.ma && again.json.ma.taskAction], [200, 1, 1, 'tracking', 'open', 'reopened'], 'A6d 再次成交沒帶期間：重開、不重複建');
  const nA2 = (await notifs(U.u1, 'deal_task')).filter((n) => n.refId === taskOf(d, a1.id).id);
  eq(nA2.map((n) => n.read).sort(), [false, true], 'A6e 重開後再發一則未讀通知');
  eq((oppRaw(d, a1.id).stageHistory || []).map((h) => h.from + '>' + h.to), ['C>Won', 'Won>A', 'A>Won'], 'A6f stageHistory 照舊記錄');
  await http(U.u1, 'PUT', `/api/opportunities/${a1.id}`, { stage: 'B' });
  const newP = await http(U.u1, 'PUT', `/api/opportunities/${a1.id}`, { stage: 'Won', termFrom: '2027-01-01', termTo: '2027-12-31' });
  d = readData();
  eq([newP.status, itemsOf(d, a1.id).length, itemsOf(d, a1.id).filter((i) => i.status === 'tracking').length, (d.dealTasks || []).filter((t) => t.oppId === a1.id).length], [200, 2, 1, 1], 'A6g 再成交帶新期間：新增一項、舊作廢項保留、任務卡仍只有一張');
  const same = await http(U.u1, 'PUT', `/api/opportunities/${a1.id}`, { stage: 'Won', termFrom: '2027-01-01', termTo: '2027-12-31', description: 'x' });
  d = readData();
  eq([same.status, itemsOf(d, a1.id).length, 'ma' in same.json], [200, 2, false], 'A6h Won→Won（階段沒變）不再處理');

  // A7 Excel 匯入成交不自動建立 → 待補納管 → 補登
  const aoa = [['客戶名稱', '銷售案名', '業務帳號', '把握度', '預定簽約日', '合約金額', '成交日期'], ['匯入假公司', MAP, U.u1, 'Won', '2026-09-01', AMT, '2026-09-02']];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'S');
  const fd = new FormData(); fd.append('file', new Blob([XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })]), 'o.xlsx');
  const imp = await http(U.adm, 'POST', '/api/admin/opportunities/import', fd, true);
  d = readData();
  const impOpp = (d.opportunities || []).find((o) => o.company === '匯入假公司');
  eq([imp.status, !!impOpp && impOpp.stage, !!impOpp && itemsOf(d, impOpp.id).length, !!impOpp && !!taskOf(d, impOpp.id)], [200, 'Won', 0, false], 'A7 Excel 匯入成交：不自動建立納管項與任務卡');
  let pb = await http(U.u1, 'GET', '/api/ma-items/pending-backfill');
  eq(pb.json.items.map((x) => x.oppId).sort(), [impOpp.id, 'oBF'].sort(), 'A7b 待補納管清單列出匯入的與舊的已成交商機');
  ok(clean(pb), 'A7c 待補納管不含金額與備註');
  eq((await http(U.u2, 'POST', '/api/ma-items/backfill', Object.assign({ oppId: impOpp.id }, T))).status, 403, 'A7d 非負責業務不能補登');
  eq((await http(U.u1, 'POST', '/api/ma-items/backfill', { oppId: impOpp.id })).json.code, 'TERM_REQUIRED', 'A7e 補登缺期間 400');
  eq((await http(U.u1, 'POST', '/api/ma-items/backfill', Object.assign({ oppId: 'oQ4' }, T))).json.code, 'OPP_NOT_WON', 'A7f 補登未成交商機被拒');
  const bf = await http(U.u1, 'POST', '/api/ma-items/backfill', Object.assign({ oppId: impOpp.id }, T));
  d = readData();
  eq([bf.status, itemsOf(d, impOpp.id).length, !!taskOf(d, impOpp.id), bf.json.taskAction], [201, 1, true, 'created'], 'A7g 補登：建立納管項＋任務卡');
  eq((await http(U.u1, 'POST', '/api/ma-items/backfill', Object.assign({ oppId: impOpp.id }, T))).json.code, 'ALREADY_TRACKED', 'A7h 重複補登被拒');
  pb = await http(U.u1, 'GET', '/api/ma-items/pending-backfill');
  eq(pb.json.items.map((x) => x.oppId), ['oBF'], 'A7i 補登後從待補清單消失');

  // A8 簽約確認
  const tA1 = taskOf(readData(), a1.id), tA2 = taskOf(readData(), a2.id), tA3 = taskOf(readData(), a3.id);
  const sign = (who, id, body) => http(who, 'POST', `/api/deal-tasks/${id}/sign`, body);
  eq((await sign(U.u2, tA1.id, { contractNo: 'C-1', signedDate: '2026-02-01' })).status, 403, 'A8 非負責業務不能簽約確認');
  eq((await sign(U.mkt, tA1.id, { contractNo: 'C-1', signedDate: '2026-02-01' })).status, 403, 'A8b 行銷也不行');
  eq((await sign(U.u1, 'nope', { contractNo: 'C-1', signedDate: '2026-02-01' })).status, 404, 'A8c 找不到任務卡 404');
  for (const [lab, b, code] of [['沒編號', { signedDate: '2026-02-01' }, 'CONTRACT_NO_REQUIRED'], ['空白編號', { contractNo: '   ', signedDate: '2026-02-01' }, 'CONTRACT_NO_REQUIRED'], ['65 字', { contractNo: 'x'.repeat(65), signedDate: '2026-02-01' }, 'CONTRACT_NO_INVALID'], ['控制字元', { contractNo: 'a\u0007b', signedDate: '2026-02-01' }, 'CONTRACT_NO_INVALID'], ['沒簽約日', { contractNo: 'C-1' }, 'SIGNED_DATE_REQUIRED'], ['假簽約日', { contractNo: 'C-1', signedDate: '2026-13-01' }, 'SIGNED_DATE_INVALID']]) {
    const x = await sign(U.u1, tA1.id, b); eq([x.status, x.json && x.json.code], [400, code], 'A8d ' + lab);
  }
  eq(readData().dealTasks.find((t) => t.id === tA1.id).status, 'open', 'A8e 驗證失敗任務卡仍開啟');
  const dupL = await sign(U.u1, tA1.id, { contractNo: 'leg-001', signedDate: '2026-02-01' });
  eq([dupL.status, dupL.json.code], [409, 'DUPLICATE_CONTRACT_NO'], 'A8f 與他人合約同編號 → 409');
  ok(!dupL.text.includes(SECRET_CO) && !dupL.text.includes('LEG-001'.toLowerCase() + SECRET_CO), 'A8g 重複訊息不洩漏他人看不到的客戶名', dupL.text);
  const goneOk = await sign(U.u1, tA1.id, { contractNo: ' GONE-001 ', signedDate: '2026-02-01' });
  d = readData();
  eq(goneOk.status, 200, 'A8h 已刪除合約的編號可重用（並 trim）');
  const cA1 = (d.contracts || []).find((c) => c.maItemId && c.oppId === a1.id);
  const itemNow = itemsOf(d, a1.id).find((i) => i.status === 'tracking');
  eq([cA1.type, cA1.contractNo, cA1.startDate, cA1.endDate, cA1.owner, cA1.company, cA1.amount, cA1.signedDate, cA1.salesPerson, cA1.deleted], ['SERVICE_MA', 'GONE-001', '2027-01-01', '2027-12-31', U.u1, a1.company, AMT, '2026-02-01', '甲', undefined], 'A8i 合約欄位（起迄日取自納管項、金額取商機、業務取暱稱）');
  eq([itemNow.contractNo, itemNow.contractId === cA1.id, taskOf(d, a1.id).status, taskOf(d, a1.id).steps[0].status, taskOf(d, a1.id).steps[0].value.contractNo], ['GONE-001', true, 'done', 'done', 'GONE-001'], 'A8j 納管項回寫合約、任務完成');
  ok((await notifs(U.u1, 'deal_task')).filter((n) => n.refId === taskOf(d, a1.id).id).every((n) => n.read), 'A8k 簽完相關通知全部已讀');
  eq((await sign(U.u1, tA1.id, { contractNo: 'C-9', signedDate: '2026-02-01' })).json.code, 'TASK_NOT_OPEN', 'A8l 已完成不能再簽');
  const dupN = await sign(U.u1, tA2.id, { contractNo: 'gone-001', signedDate: '2026-02-01' });
  eq(dupN.json.code, 'DUPLICATE_CONTRACT_NO', 'A8m 與剛建立的 SERVICE_MA 合約同編號（大小寫不同）→ 重複');
  ok(dupN.status === 409 && !dupN.text.includes(a1.company), 'A8n 重複訊息不含公司名', dupN.text);
  const cl = (await http(U.u1, 'GET', '/api/contracts')).json;
  ok(cl.some((c) => c.id === cA1.id && c.type === 'SERVICE_MA'), 'A8o 合約管理列表看得到 SERVICE_MA 合約');
  eq((await http(U.sec, 'GET', '/api/contracts')).json, [], 'A8p 秘書仍看不到合約');
  eq((await http(U.u2, 'GET', '/api/contracts')).json.some((c) => c.id === cA1.id), false, 'A8q 別的業務看不到');
  eq((await auditLines('MA_TASK_SIGN')).length, 1, 'A8r 稽核 MA_TASK_SIGN');
  // 合約編號修改／刪除同步
  const cput = await http(U.u1, 'PUT', `/api/contracts/${cA1.id}`, { contractNo: 'GONE-002' });
  d = readData();
  eq([cput.status, itemsOf(d, a1.id).find((i) => i.status === 'tracking').contractNo], [200, 'GONE-002'], 'A8s 合約改編號 → 納管項同步');
  const cdel = await http(U.u1, 'DELETE', `/api/contracts/${cA1.id}`);
  d = readData();
  const ia = itemsOf(d, a1.id).find((i) => i.status === 'tracking');
  eq([cdel.status, ia.contractNo, ia.contractId], [200, '', ''], 'A8t 合約刪除 → 納管項解除連結');
  // 合約刪除重開簽約（驗收後修正）
  const tRe = taskOf(d, a1.id);
  eq([tRe.status, tRe.steps[0].status, tRe.history.slice(-1)[0].action, tRe.history.slice(-1)[0].note], ['open', 'pending', 'REOPEN', '合約已刪除，重開簽約'], 'A8t1 刪除合約 → 任務卡重開（open、簽約步驟 pending、history 記錄）');
  const reN = (await notifs(U.u1, 'deal_task')).filter((n) => n.refId === tRe.id && !n.read);
  ok(reN.length === 1 && reN[0].title.includes(a1.company) && !reN[0].title.includes(AMT) && !reN[0].body.includes(AMT) && !reN[0].body.includes(NOTE), 'A8t2 重新通知負責業務（未讀 1 則，只有公司名與品項、無金額）', JSON.stringify(reN));
  const delReopen = async () => (await auditLines('MA_TASK_REOPEN')).filter((l) => JSON.stringify(l).includes('合約已刪除')).length;
  eq([await delReopen(), (await auditLines('MA_CONTRACT_UNLINK')).length], [1, 1], 'A8t3 稽核 MA_TASK_REOPEN／MA_CONTRACT_UNLINK 各一筆');
  const lAfter = (await http(U.u1, 'GET', '/api/ma-items?filter=pending_sign')).json.items.find((i) => i.id === ia.id);
  eq([!!lAfter, lAfter && lAfter.pendingSign, lAfter && lAfter.taskStatus], [true, true, 'open'], 'A8t4 到期管理：刪除後重新出現在待簽約、任務 open（有入口可簽）');
  const cdel2 = await http(U.u1, 'DELETE', '/api/contracts/' + cA1.id);
  d = readData();
  eq([cdel2.status, taskOf(d, a1.id).history.filter((h) => h.note === '合約已刪除，重開簽約').length, (await notifs(U.u1, 'deal_task')).filter((n) => n.refId === tRe.id && !n.read).length, await delReopen(), (await auditLines('MA_CONTRACT_UNLINK')).length], [200, 1, 1, 1, 1], 'A8t5 重複刪除冪等：不重複重開、不重複通知、不重複稽核');
  const rs = await sign(U.u1, tRe.id, { contractNo: 'GONE-001', signedDate: '2026-02-05' });
  d = readData();
  eq([rs.status, taskOf(d, a1.id).status, itemsOf(d, a1.id).find((i) => i.status === 'tracking').contractNo, (d.contracts || []).filter((c) => c.oppId === a1.id && !c.deleted).length], [200, 'done', 'GONE-001', 1], 'A8t6 重開後可重新簽約（同編號可重用），只有一份有效合約');
  const cNew = (d.contracts || []).find((c) => c.oppId === a1.id && !c.deleted);
  const cdelOther = await http(U.u2, 'DELETE', '/api/contracts/' + cNew.id);
  d = readData();
  eq([cdelOther.status, (d.contracts.find((c) => c.id === cNew.id) || {}).deleted, taskOf(d, a1.id).status, itemsOf(d, a1.id).find((i) => i.status === 'tracking').contractNo], [200, undefined, 'done', 'GONE-001'], 'A8t7 非合約擁有者刪除：不生效，任務卡與納管項不動');
  // admin 可代簽
  const adminSign = await sign(U.adm, tA2.id, { contractNo: 'ADM-1', signedDate: '2026-03-01' });
  eq(adminSign.status, 200, 'A8u 管理員可代為簽約確認');
  eq((readData().contracts.find((c) => c.contractNo === 'ADM-1') || {}).owner, U.u1, 'A8v 合約 owner 仍是任務卡負責業務');

  // A9 略過
  const skip = (who, id, b) => http(who, 'POST', `/api/deal-tasks/${id}/skip`, b);
  eq((await skip(U.u2, tA3.id, { reason: '客戶自行處理' })).status, 403, 'A9 非負責業務不能略過');
  eq((await skip(U.u1, tA3.id, {})).json.code, 'SKIP_REASON_REQUIRED', 'A9b 沒原因 400');
  eq((await skip(U.u1, tA3.id, { reason: '一' })).json.code, 'SKIP_REASON_REQUIRED', 'A9c 原因 1 字 400');
  const sk = await skip(U.u1, tA3.id, { reason: '客戶改用其他方案' });
  d = readData();
  eq([sk.status, sk.json.notified, taskOf(d, a3.id).status, itemsOf(d, a3.id)[0].status], [200, 3, 'skipped', 'tracking'], 'A9d 略過：任務 skipped、納管項保留追蹤、通知 3 人（兩位行銷＋一級主管）');
  for (const [who, n] of [[U.mkt, 1], [U.mkt2, 1], [U.m1, 1], [U.mktoff, 0], [U.u1, 0], [U.m2, 0]]) eq((await notifs(who, 'deal_task_skipped')).length, n, 'A9e 略過通知 ' + who + ' = ' + n);
  ok((await notifs(U.mkt, 'deal_task_skipped')).every((n) => n.body.includes('客戶改用其他方案') && !n.body.includes(AMT)), 'A9f 通知含原因、不含金額');
  {
    const aAll = (await http(U.u1, 'GET', '/api/ma-items?filter=all')).json.items.find((i) => i.oppId === a3.id), aPend = (await http(U.u1, 'GET', '/api/ma-items?filter=pending_sign')).json.items;
    eq([aAll.pendingSign, aAll.taskStatus, aPend.some((i) => i.oppId === a3.id)], [false, 'skipped', false], 'A9d2 此案不簽合約的納管項：不再算待簽約');
  }
  eq((await skip(U.u1, tA3.id, { reason: '客戶改用其他方案' })).json.code, 'TASK_NOT_OPEN', 'A9g 不能重複略過');
  eq((await sign(U.u1, tA3.id, { contractNo: 'C-77', signedDate: '2026-02-01' })).json.code, 'TASK_NOT_OPEN', 'A9h 略過後不能簽');
  ok((await notifs(U.u1, 'deal_task')).filter((n) => n.refId === tA3.id).every((n) => n.read), 'A9i 略過後原通知已讀');
  eq((await auditLines('MA_TASK_SKIP')).length, 1, 'A9j 稽核 MA_TASK_SKIP');

  // A10 延後 / A11 改派
  const sn = await http(U.u1, 'POST', `/api/deal-tasks/${taskOf(readData(), li.id) ? 'x' : q1.json.ma.taskId}/snooze`);
  eq(sn.json.snoozedUntil, plusDays(todayTp(), 3), 'A10 延後 3 天（台北日曆日）');
  eq((await http(U.u2, 'POST', `/api/deal-tasks/${q1.json.ma.taskId}/snooze`)).status, 403, 'A10b 非負責業務不能延後');
  let mine = (await http(U.u1, 'GET', '/api/my-tasks/deals')).json;
  const mq1 = mine.items.find((t) => t.id === q1.json.ma.taskId);
  eq([mq1.snoozed, mq1.snoozedUntil, mq1.items.length, mq1.overdue, mq1.daysOpen], [true, plusDays(todayTp(), 3), 2, false, 0], 'A10c my-tasks 顯示延後狀態與納管項摘要');
  ok(clean({ text: JSON.stringify(mine) }), 'A10d my-tasks 不含金額與備註');
  const tq1 = q1.json.ma.taskId;
  eq((await http(U.u1, 'POST', `/api/deal-tasks/${tq1}/reassign`, { to: U.u2 })).status, 403, 'A11 非管理員不能改派');
  eq((await http(U.adm, 'POST', `/api/deal-tasks/${tq1}/reassign`, { to: U.uoff })).json.code, 'ASSIGNEE_INACTIVE', 'A11b 改派給停用者被拒');
  eq((await http(U.adm, 'POST', `/api/deal-tasks/${tq1}/reassign`, { to: U.mkt })).json.code, 'ASSIGNEE_ROLE', 'A11c 改派給行銷被拒');
  const rsn = (to) => http(U.adm, 'POST', '/api/deal-tasks/' + tq1 + '/reassign', { to });
  eq([(await rsn(U.m1)).json.code, (await rsn(U.m2)).json.code, (await rsn(U.gs)).json.code], ['ASSIGNEE_ROLE', 'ASSIGNEE_ROLE', 'ASSIGNEE_ROLE'], 'A11c2 改派給主管／集團業務 → 400 ASSIGNEE_ROLE');
  eq((await rsn(U.u3)).json.code, 'ASSIGNEE_BU', 'A11c3 改派給別的 BU 業務 → 400 ASSIGNEE_BU');
  eq(readData().dealTasks.find((t) => t.id === tq1).owner, U.u1, 'A11c4 被拒絕後負責人不變');
  const cands = await http(U.adm, 'GET', '/api/deal-tasks/' + tq1 + '/reassign-candidates');
  eq([cands.status, (cands.json.users || []).map((x) => x.username).sort()], [200, [U.u2, U.u4, U.u5].sort()], 'A11c5 候選人 API：只有在職、user、ERP、非目前負責人');
  eq([(await http(U.u1, 'GET', '/api/deal-tasks/' + tq1 + '/reassign-candidates')).status, (await http(U.m1, 'GET', '/api/deal-tasks/' + tq1 + '/reassign-candidates')).status, (await http(U.adm, 'GET', '/api/deal-tasks/nope/reassign-candidates')).status], [403, 403, 404], 'A11c6 候選人 API 限管理員；找不到 404');
  const ra = await http(U.adm, 'POST', '/api/deal-tasks/' + tq1 + '/reassign', { to: U.u2 });
  d = readData();
  eq([ra.status, d.dealTasks.find((t) => t.id === tq1).owner, itemsOf(d, 'oQ1').every((i) => i.owner === U.u2)], [200, U.u2, true], 'A11d 管理員改派：任務與納管項換負責人');
  eq([(await http(U.u2, 'GET', '/api/my-tasks/deals')).json.items.some((t) => t.id === tq1), (await http(U.u1, 'GET', '/api/my-tasks/deals')).json.items.some((t) => t.id === tq1)], [true, false], 'A11e 新負責人看得到、舊的看不到');
  eq((await notifs(U.u2, 'deal_task')).filter((n) => n.refId === tq1).length, 1, 'A11f 新負責人收到通知');
  eq((await auditLines('MA_TASK_REASSIGN')).length, 1, 'A11g 稽核 MA_TASK_REASSIGN');
  // A12 離職者的逾期任務
  const ov = (who) => http(who, 'GET', '/api/manager/deal-task-overdue');
  eq((await ov(U.m1)).json.items.map((t) => t.id), ['tOld'], 'A12 一級主管看得到離職團隊成員逾期任務');
  eq(((await ov(U.m1)).json.items[0] || {}).ownerActive, false, 'A12b 標示負責人已停用');
  eq([(await ov(U.u1)).status, (await ov(U.sec)).status, (await ov(U.mkt)).status], [403, 403, 403], 'A12c 業務／秘書／行銷 403');
  eq((await ov(U.m1b)).json.items.length, 0, 'A12d 別條線的主管看不到');
  eq((await ov(U.adm)).json.canReassign, true, 'A12e 管理員可改派');
  const rr = await http(U.adm, 'POST', '/api/deal-tasks/tOld/reassign', { to: U.u2 });
  eq([rr.status, (await http(U.u2, 'GET', '/api/my-tasks/deals')).json.items.some((t) => t.id === 'tOld')], [200, true], 'A12f 管理員把離職者任務改派給在職業務');
  // A13 修改期間
  const itemA2 = itemsOf(readData(), a2.id)[0];
  const patch = (who, id, b) => http(who, 'PATCH', `/api/ma-items/${id}/term`, b);
  eq((await patch(U.u2, itemA2.id, { termFrom: '2026-02-01', termTo: '2027-01-31' })).status, 403, 'A13 非負責業務不能改期間');
  eq((await patch(U.u1, itemA2.id, { termFrom: '2027-01-31', termTo: '2026-02-01' })).json.code, 'TERM_INVALID', 'A13b 顛倒 400');
  const pt = await patch(U.u1, itemA2.id, { termFrom: '2026-02-01', termTo: '2027-01-31', reason: '延後啟用' });
  d = readData();
  eq([pt.status, pt.json.from, pt.json.to, itemsOf(d, a2.id)[0].history.slice(-1)[0].action], [200, '2026-02-01', '2027-01-31', 'TERM_EDIT'], 'A13c 業務改自己的期間＋history');
  ok(clean(pt), 'A13d 回應不含金額');
  eq((await auditLines('MA_ITEM_TERM_EDIT')).length, 1, 'A13e 稽核 MA_ITEM_TERM_EDIT');
  const cA2 = readData().contracts.find((c) => c.contractNo === 'ADM-1');
  const pt2 = await patch(U.adm, itemA2.id, { termFrom: '2026-03-01', termTo: '2027-02-28' });
  eq([pt2.status, readData().contracts.find((c) => c.id === cA2.id).startDate, readData().contracts.find((c) => c.id === cA2.id).endDate], [200, '2026-03-01', '2027-02-28'], 'A13f 已簽約的項目改期間 → 合約起迄日同步');
  // A14 刪除／還原
  const delOpp = await newOpp(U.u1, MAP);
  const wd = await http(U.u1, 'PUT', `/api/opportunities/${delOpp.id}`, Object.assign({ stage: 'Won' }, T));
  const del = await http(U.u1, 'DELETE', `/api/opportunities/${delOpp.id}`, { deleteReason: '其他：測試' });
  d = readData();
  eq([wd.status, del.status, taskOf(d, delOpp.id).status, itemsOf(d, delOpp.id)[0].status], [200, 200, 'cancelled', 'void'], 'A14 商機刪除 → 任務卡取消、納管項作廢');
  const rest = await http(U.u1, 'POST', `/api/opportunities/restore/${delOpp.id}`);
  d = readData();
  eq([rest.status, taskOf(d, delOpp.id).status, itemsOf(d, delOpp.id)[0].status, (d.dealTasks || []).filter((t) => t.oppId === delOpp.id).length], [200, 'open', 'tracking', 1], 'A14b 商機還原 → 重開、不重複建');
  // A15 到期管理清單基本
  const lst = await http(U.u1, 'GET', '/api/ma-items');
  ok(lst.status === 200 && lst.json.items.length >= 5 && clean(lst), 'A15 到期管理清單可用且不含金額／備註', lst.status);
  eq(lst.json.items.map((i) => i.to), lst.json.items.map((i) => i.to).slice().sort(), 'A15b 依到期日排序');
  const sec = await http(U.sec, 'GET', '/api/ma-items');
  ok(sec.status === 200 && sec.json.items.length > 0 && sec.json.items.every((x) => Object.keys(x).sort().join() === 'company,daysLeft,from,ownerLabel,product,termKind,to') && !sec.text.includes('GONE') && !sec.text.includes('ADM-1'), 'A15c 秘書只有 7 個欄位、沒有合約編號', sec.text.slice(0, 200));
  await stopServer(); cleanupFiles();
}

// ═════════════ 角色權限一致性（23 帳號）═════════════
async function groupRoles() {
  console.log('[角色] 23 帳號逐一比對 my-tasks／ma-items／overdue／pending-backfill');
  const T0 = todayTp();
  const owners = [U.u1, U.u2, U.u3, U.u4, U.u5, U.uoff, U.m2, U.m1, U.m1b, U.gs, U.mkt];
  let n = 0;
  const seed = (d) => {
    d.opportunities = []; d.maItems = []; d.dealTasks = []; d.quotations = []; d.lostOpportunities = [];
    owners.forEach((ow, i) => {
      for (let k = 0; k < 3; k++) {
        const id = `r${i}_${k}`;
        const co = (ow === U.u1 && k === 0) ? '假公司gA' : '假公司' + id;   // u1 的第一筆在集團 g1 內
        const o = mkOpp(id, ow, k === 2 ? PLAINP : MAP, k === 1 && i % 3 === 0 ? 'B' : 'Won', { company: co });
        if (ow === U.u3 || ow === U.m1b) o.bu = 'ITS';
        d.opportunities.push(o);
        if (o.stage !== 'Won' || k === 2) continue;
        const offs = [[-30, 60], [0, 90], [10, 91], [-400, -1], [5, 400]][(i + k) % 5];
        const iid = 'i' + id;
        d.maItems.push({ id: iid, oppId: id, quoteId: '', owner: ow, company: co, product: o.product, termKind: k === 0 ? 'ma' : 'license', from: plusDays(T0, offs[0]), to: plusDays(T0, offs[1]), contractNo: (i + k) % 2 ? 'CN-' + id : '', contractId: '', status: (i + k) % 7 === 0 ? 'ended' : 'tracking', source: 'request', createdAt: agoIso(5), updatedAt: agoIso(5), history: [] });
        if (k === 0) d.dealTasks.push({ id: 't' + id, type: 'ma_sign', oppId: id, maItemId: iid, maItemIds: [iid], owner: ow, company: co, product: o.product, status: i % 4 === 3 ? 'done' : 'open', steps: [{ key: 'sign', status: 'pending', value: null }], createdAt: agoIso(i % 2 ? 20 : 3), updatedAt: agoIso(3), history: [] });
        n++;
      }
    });
    // 已作廢項（商機已被刪除）、有納管項但商機已不是 Won（作廢）、已成交缺納管（待補）
    d.lostOpportunities.push(mkOpp('rLost', U.u1, MAP, 'Won'));
    d.maItems.push({ id: 'iLost', oppId: 'rLost', owner: U.u1, company: '假公司rLost', product: MAP, termKind: 'ma', from: '2026-01-01', to: plusDays(T0, 30), contractNo: '', status: 'void', createdAt: nowIso, history: [] });
    d.opportunities.push(mkOpp('rBF1', U.u1, MAP, 'Won'), mkOpp('rBF2', U.u3, MAP, 'Won', { bu: 'ITS' }), mkOpp('rBF3', U.u2, LICP, 'Won'));
    d.maItems.push({ id: 'iVoidOnly', oppId: 'rBF3', owner: U.u2, company: '假公司rBF3', product: LICP, termKind: 'license', from: '2026-01-01', to: plusDays(T0, 30), contractNo: '', status: 'void', createdAt: nowIso, history: [] });
    d.productCatalog = CAT();
  };
  await freshServer({ data: seed });
  const raw = readData();
  const catalog = raw.productCatalog;
  const TERM_ONLY = (o) => REAL.PC.resolveTermRule(catalog, o.product).termKind !== 'none';
  const ops = new Map(raw.opportunities.map((o) => [o.id, o]));
  const FIVE = ['company', 'daysLeft', 'from', 'ownerLabel', 'product', 'termKind', 'to'];
  const dl = (to) => Math.round((Date.parse(to + 'T00:00:00Z') - Date.parse(T0 + 'T00:00:00Z')) / DAY);
  const mgrRoles = ['admin', 'executive', 'manager1', 'manager2'];
  const userRole = Object.fromEntries(seedUsers().map((u) => [u.username, u.role]));
  const label = (un) => { const u = seedUsers().find((x) => x.username === un); return (u && (u.nickname || u.displayName || u.username)) || un; };
  let total = 0;
  const szMap = {};
  const CHUNKS = [ALLU.slice(0, 11), ALLU.slice(11)];   // 登入限流＝每個行程 20 次／15 分鐘 → 分兩批、中間重啟（資料檔不動）
  for (let ci = 0; ci < CHUNKS.length; ci++) {
  if (ci > 0) { await stopServer(); await startServer(MUT_CUR, { MA_TRACKING: '1' }); jar = {}; }
  Object.keys(LOGIN_STATUS).forEach((k) => delete LOGIN_STATUS[k]);
  await loginAll(CHUNKS[ci]);
  ok(Object.keys(LOGIN_STATUS).filter((w) => LOGIN_STATUS[w] !== 200).every((w) => ['_s_m1off', '_s_uoff', '_s_mktoff'].includes(w)), 'R 登入：只有停用帳號登不進來', JSON.stringify(LOGIN_STATUS));
  for (const who of CHUNKS[ci]) {
    if (!jar[who]) { eq([(await http(who, 'GET', '/api/ma-items')).status, (await http(who, 'GET', '/api/my-tasks/deals')).status], [401, 401], `R ${who} 停用帳號無法登入、端點 401`); continue; }
    const opRes = await http(who, 'GET', '/api/opportunities');
    if (!Array.isArray(opRes.json)) console.log('    ! GET /api/opportunities', who, opRes.status, opRes.text.slice(0, 120));
    const visible = new Set((Array.isArray(opRes.json) ? opRes.json : []).map((o) => o.id));   // 暴力基準：該角色實際看得到的商機
    const role = userRole[who];
    if (role === 'tecopm') {   // 集團PM：路徑白名單只放行商機／名片等唯讀端點，MA 端點一律 403（不是漏洞，是此角色本來就不用）
      eq([(await http(who, 'GET', '/api/my-tasks/deals')).status, (await http(who, 'GET', '/api/ma-items')).status, (await http(who, 'GET', '/api/manager/deal-task-overdue')).status, (await http(who, 'GET', '/api/ma-items/pending-backfill')).status], [403, 403, 403, 403], `R ${who}(tecopm) MA 端點全部 403`);
      szMap[who] = -2; continue;
    }
    // my-tasks
    const mine = await http(who, 'GET', '/api/my-tasks/deals');
    eq(mine.json.items.map((t) => t.id).sort(), raw.dealTasks.filter((t) => t.owner === who && t.status === 'open').map((t) => t.id).sort(), `R my-tasks ${who}(${role}) 只有自己的開啟中任務卡`);
    ok(clean(mine), `R my-tasks ${who} 不含金額／備註`);
    // ma-items 各 filter
    for (const f of ['all', 'pending_sign', 'd90', 'overdue']) {
      const res = await http(who, 'GET', '/api/ma-items?filter=' + f);
      const exp = raw.maItems.filter((it) => it.status !== 'void' && visible.has(it.oppId) && (f === 'all' || (it.status === 'tracking' && (f === 'pending_sign' ? !it.contractNo && raw.dealTasks.some((t) => t.status === 'open' && (t.maItemIds || [t.maItemId]).includes(it.id)) : f === 'd90' ? dl(it.to) >= 0 && dl(it.to) <= 90 : dl(it.to) < 0))))
        .sort((a, b) => (a.to < b.to ? -1 : a.to > b.to ? 1 : (a.id < b.id ? -1 : 1)));
      if (role === 'secretary') {
        eq(res.json.items, exp.map((it) => ({ company: it.company, product: it.product, termKind: it.termKind, from: it.from, to: it.to, daysLeft: dl(it.to), ownerLabel: label(it.owner) })), `R ma-items(${f}) ${who} 秘書：欄位白名單＋內容逐項相等`);
        ok(res.json.items.every((x) => Object.keys(x).sort().join() === FIVE.join()), `R ma-items(${f}) ${who} 秘書每筆恰 7 欄`);
      } else {
        eq(res.json.items.map((x) => x.id), exp.map((x) => x.id), `R ma-items(${f}) ${who}(${role}) 清單與預期逐項相等（含排序）`);
        ok(res.json.items.every((x) => x.daysLeft === dl(x.to) && !('amount' in x) && x.contractNo === (ops.size && raw.maItems.find((i) => i.id === x.id).contractNo)), `R ma-items(${f}) ${who} 欄位值正確、無金額`);
      }
      ok(clean(res), `R ma-items(${f}) ${who} 不含金額／備註`);
      total += exp.length;
    }
    // overdue
    const ovr = await http(who, 'GET', '/api/manager/deal-task-overdue');
    if (!mgrRoles.includes(role)) eq(ovr.status, 403, `R overdue ${who}(${role}) 非主管 403`);
    else {
      // 預期：開啟中、建立超過 14 天、負責人的商機在可視範圍內
      const exp = raw.dealTasks.filter((t) => t.status === 'open' && Math.round((Date.parse(todayTp() + 'T00:00:00Z') - Date.parse(new Date(Date.parse(t.createdAt) + 8 * 3600e3).toISOString().slice(0, 10) + 'T00:00:00Z')) / DAY) > 14 && visible.has(t.oppId)).map((t) => t.id).sort();
      eq(ovr.json.items.map((t) => t.id).sort(), exp, `R overdue ${who}(${role}) 逾期清單與預期相等`);
      ok(clean(ovr), `R overdue ${who} 不含金額／備註`);
    }
    // pending-backfill
    const pbr = await http(who, 'GET', '/api/ma-items/pending-backfill');
    const trackedOpp = new Set(raw.maItems.filter((i) => i.status !== 'void').map((i) => i.oppId));
    const expPb = role === 'secretary' ? [] : raw.opportunities.filter((o) => visible.has(o.id) && o.stage === 'Won' && TERM_ONLY(o) && !trackedOpp.has(o.id)).map((o) => o.id).sort();
    eq(pbr.json.items.map((x) => x.oppId).sort(), expPb, `R pending-backfill ${who}(${role}) 與預期相等`);
    ok(clean(pbr), `R pending-backfill ${who} 不含金額／備註`);
    szMap[who] = (await http(who, 'GET', '/api/ma-items')).json.items.length;
  }
  }
  ok(total > 100, 'R 比對規模足夠（預期項目數）', total);
  // 暴力基準自檢：不同角色的預期不全相同（避免所有人都空而假過）
  const sizes = new Set(Object.values(szMap));
  eq(Object.keys(szMap).length, 20, 'R 20 個可登入帳號都比對過（另 3 個停用帳號驗 401）');
  ok(sizes.size >= 4, 'R 各角色可視筆數不同（基準有鑑別力）', [...sizes].join());
  await stopServer(); cleanupFiles();
}

// ═════════════ 開關：MA_TRACKING 未設 ═════════════
function exportBase() {
  if (process.env.MATRK_BASE_DIR) return { dir: process.env.MATRK_BASE_DIR, cleanup() {} };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'matrk-base-'));
  const baseRef = process.env.MATRK_BASE_REF || '7efdba5';
  const a = spawnSync('git', ['archive', baseRef], { cwd: ROOT, maxBuffer: 256 * 1024 * 1024 });
  if (a.status !== 0) throw new Error('git archive ' + baseRef + ' 失敗：' + String(a.stderr));
  const t = spawnSync('tar', ['-x', '-C', dir], { input: a.stdout, maxBuffer: 256 * 1024 * 1024 });
  if (t.status !== 0) throw new Error('tar 失敗：' + String(t.stderr));
  spawnSync('cmd', ['/c', 'mklink', '/J', path.join(dir, 'node_modules'), path.join(ROOT, 'node_modules')], { stdio: 'ignore' });
  return { dir, cleanup() { spawnSync('cmd', ['/c', 'rmdir', path.join(dir, 'node_modules')], { stdio: 'ignore' }); try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ } } };
}
async function groupOff() {
  console.log('[開關] MA_TRACKING 未設 ＝ 與基準版逐位元相同');
  const run = async (root, port, mut) => {
    await stopServer(); cleanupFiles(root); seedFiles({ dir: root });
    await startServer(mut || null, {}, root, port);   // 不傳 MA_TRACKING
    const base = `http://localhost:${port}`;
    const jr = {};
    for (const w of [U.adm, U.u1, U.u2, U.sec, U.mkt, U.m1]) await loginAt(base, jr, w);
    const out = {}; const extra = {};
    const rec = async (key, who, method, url, body) => { const r = await httpAt(base, jr, who, method, url, body); out[key] = r.status + ' ' + norm(r.text); return r; };
    const o1 = await rec('postOpp', U.u1, 'POST', '/api/opportunities', { company: '開關測試公司', contactName: '甲', product: MAP, amount: AMT, expectedDate: '2026-12-31', bu: 'ERP', category: 'ERP' });
    const id = o1.json.id;
    await rec('putWonNoTerm', U.u1, 'PUT', `/api/opportunities/${id}`, { stage: 'Won' });
    await rec('putWonTermBody', U.u1, 'PUT', `/api/opportunities/${id}`, { stage: 'Won', termFrom: '2026-01-01', termTo: '2026-12-31' });
    await rec('putBack', U.u1, 'PUT', `/api/opportunities/${id}`, { stage: 'A' });
    await rec('putWonQuote', U.u1, 'PUT', '/api/opportunities/oQ1', { stage: 'Won' });
    await rec('putWonPlain', U.u1, 'PUT', '/api/opportunities/oQ2', { stage: 'Won', product: PLAINP });
    await rec('putForm', U.u1, 'PUT', `/api/opportunities/${id}`, { company: '開關測試公司', product: MAP, amount: AMT, expectedDate: '2026-12-31', stage: 'Won' });
    await rec('delOpp', U.u1, 'DELETE', `/api/opportunities/${id}`, { deleteReason: '其他：測試' });
    await rec('restoreOpp', U.u1, 'POST', `/api/opportunities/restore/${id}`);
    await rec('postContract', U.u1, 'POST', '/api/contracts', { contractNo: 'OFF-1', company: '開關公司', product: 'x', startDate: '2026-01-01', endDate: '2026-12-31', type: 'SERVICE_MA', oppId: 'zzz', maItemId: 'yyy' });
    const cid = (readData(root).contracts.find((c) => c.contractNo === 'OFF-1') || {}).id;
    await rec('putContract', U.u1, 'PUT', `/api/contracts/${cid}`, { contractNo: 'OFF-2', oppId: 'zzz' });
    await rec('delContract', U.u1, 'DELETE', `/api/contracts/${cid}`);
    for (const who of [U.adm, U.u1, U.u2, U.sec, U.mkt, U.m1]) for (const g of ['/api/opportunities', '/api/contracts', '/api/notifications', '/api/poll-bundle', '/api/product-catalog', '/api/me/permissions', '/api/lost-opportunities', '/api/pipeline-report', '/api/quotations']) await rec(who + ' ' + g, who, 'GET', g);
    const raw = readData(root);
    out.keys = JSON.stringify(Object.keys(raw).sort());
    out.oppKeys = JSON.stringify((raw.opportunities || []).map((o) => Object.keys(o).sort().filter((k) => !['updatedAt', 'createdAt', 'id'].includes(k))));
    out.contractKeys = JSON.stringify((raw.contracts || []).map((c) => Object.keys(c).sort().filter((k) => !['createdAt', 'id'].includes(k))));
    for (const [k, m, u, b] of [['myTasks', 'GET', '/api/my-tasks/deals'], ['overdue', 'GET', '/api/manager/deal-task-overdue'], ['maItems', 'GET', '/api/ma-items'], ['pending', 'GET', '/api/ma-items/pending-backfill'], ['backfill', 'POST', '/api/ma-items/backfill', {}], ['patch', 'PATCH', '/api/ma-items/x/term', {}], ['sign', 'POST', '/api/deal-tasks/tOld/sign', {}], ['skip', 'POST', '/api/deal-tasks/tOld/skip', {}], ['snooze', 'POST', '/api/deal-tasks/tOld/snooze'], ['reassign', 'POST', '/api/deal-tasks/tOld/reassign', {}], ['reassignCands', 'GET', '/api/deal-tasks/tOld/reassign-candidates'], ['termHint', 'GET', '/api/ma-items/term-hint?oppId=oOld']].map((a) => a.length === 3 ? [a[0], a[1], a[2], undefined] : a)) extra[k] = await httpAt(base, jr, U.adm, m, u, b);
    await stopServer(); cleanupFiles(root);
    return { out, extra, raw };
  };
  const baseX = exportBase();
  let A, B;
  try {
    A = await run(baseX.dir, PORT);
    B = await run(ROOT, PORT, MUT_CUR);
  } finally { baseX.cleanup(); }
  const diffKeys = Object.keys(A.out).filter((k) => A.out[k] !== B.out[k]);
  ok(diffKeys.length === 0, `F1 off：成交／回退／刪除／還原／合約／列表／通知／poll-bundle 的輸出與基準版逐位元相同（比對 ${Object.keys(A.out).length} 項）`, diffKeys.join(', '));
  if (diffKeys.length) console.log('    差異首項：', diffKeys[0], '\n      前：', String(A.out[diffKeys[0]]).slice(0, 300), '\n      後：', String(B.out[diffKeys[0]]).slice(0, 300));
  eq(Object.keys(B.extra).map((k) => B.extra[k].status), Object.keys(B.extra).map((k) => A.extra[k].status), 'F2 off：新端點的回應狀態與基準版（本來就沒有）相同');
  ok(Object.keys(B.extra).every((k) => B.extra[k].status === 404 || B.extra[k].status === 405), 'F2b off：新端點全部 404（未登記）', JSON.stringify(Object.keys(B.extra).map((k) => k + ':' + B.extra[k].status)));
  ok(!B.raw.maItems || B.raw.maItems.length === 1, 'F3 off：不新增納管項（只剩種子那 1 筆）');
  ok(!('dealTasks' in B.raw) || B.raw.dealTasks.length === 1, 'F3b off：不新增任務卡');
  eq((B.raw.notifications || []).filter((n) => /^deal_task/.test(n.type)).length, 0, 'F3c off：不產生 deal_task 通知');
  ok(/^200 /.test(B.out.putWonNoTerm) && /"stage":"Won"/.test(B.out.putWonNoTerm) && !/"ma":/.test(B.out.putWonNoTerm), 'F4 off：商機直接成交不需要期間、回應沒有 ma 區塊');
}

// ═════════════ 變異測試 ═════════════
const E = (find, rep) => ({ find, rep });
const MUTS = [
  { name: 'M1 拿掉缺期間檢查', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("if (!entries || !entries.length) {", "if (!entries) entries = [];\n  if (false) {")] }] },
  { name: 'M1b 缺期間回應改成放行（server 不擋）', groups: ['api'], spec: [{ file: 'server.js', edits: [E("if (!_pl.ok) return res.status(_pl.status).json({ error: _pl.message, code: _pl.code });", "if (!_pl.ok) { /* mut */ }")] }] },
  { name: 'M2 拿掉納管項冪等', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("let it = findItem(data, opp.id, e.termKind, e.from, e.to);", "let it = null;")] }] },
  { name: 'M2b 拿掉任務卡冪等', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("let t = findTask(data, opp.id);", "let t = null;")] }] },
  { name: 'M3 拿掉合約編號唯一', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("if (contractNoTaken(data, cn.value)) return fail(409", "if (false) return fail(409")] }] },
  { name: 'M4 拿掉秘書欄位過濾', groups: ['unit', 'api', 'roles'], spec: [{ file: 'lib/maTracking.js', edits: [E("if (o.secretary) return {", "if (false) return {")] }] },
  { name: 'M5 拿掉可視範圍（納管項清單）', groups: ['roles'], spec: [{ file: 'lib/maRoutes.js', edits: [E("let items = (data.maItems || []).filter((it) => it && owners.has(it.owner));", "let items = (data.maItems || []).filter((it) => it);")] }] },
  { name: 'M5b 拿掉 BU 過濾', groups: ['roles'], spec: [{ file: 'lib/maRoutes.js', edits: [E("    items = filterByBu(req, items);\n    items = filterByViewGroup(req, items, (it) => it.company);\n    return items;", "    return items;")] }] },
  { name: 'M5c 逾期清單拿掉可視範圍', groups: ['unit', 'roles'], spec: [{ file: 'lib/maTracking.js', edits: [E("t.status === 'open' && visibleOwners.has(t.owner) &&", "t.status === 'open' &&")] }] },
  { name: 'M5d 我的任務改成看全部人的', groups: ['unit', 'roles'], spec: [{ file: 'lib/maTracking.js', edits: [E("t && t.owner === me && t.status === 'open'", "t && t.status === 'open'")] }] },
  { name: 'M5e 逾期清單拿掉主管角色限制', groups: ['roles', 'api'], spec: [{ file: 'lib/maRoutes.js', edits: [E("if (!ADMINLIKE_OVERDUE.includes(role)) return err(res, 403", "if (false) return err(res, 403")] }] },
  { name: 'M5f 待補清單秘書不再回空', groups: ['roles'], spec: [{ file: 'lib/maRoutes.js', edits: [E("req.session.user.role === 'secretary' ? [] : MA.pendingBackfill", "MA.pendingBackfill")] }] },
  { name: 'M6 拿掉 Won 回退取消（server）', groups: ['api', 'off'], spec: [{ file: 'server.js', edits: [E("else if (oldStage === 'Won' && req.body.stage !== 'Won') _maUnwon = true;", "else if (false) _maUnwon = true;")] }] },
  { name: 'M6b 拿掉 Won 回退取消（lib）', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("t.status = 'cancelled'; t.cancelReason", "t.cancelReasonX")] }] },
  { name: 'M7 簽約確認不檢查負責人', groups: ['api'], spec: [{ file: 'lib/maRoutes.js', edits: [E("const ok = adminOnly ? req.session.user.role === 'admin' : isOwnerOrAdmin(req, t.owner);", "const ok = true;")] }] },
  { name: 'M8 改派不限管理員', groups: ['api'], spec: [{ file: 'lib/maRoutes.js', edits: [E("const ok = adminOnly ? req.session.user.role === 'admin' : isOwnerOrAdmin(req, t.owner);", "const ok = isOwnerOrAdmin(req, t.owner);")] }] },
  { name: 'M9 略過不必填原因', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("if (Array.from(reason).length < SKIP_REASON_MIN) return", "if (false) return")] }] },
  { name: 'M10 完成不標通知已讀', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("x.type === 'deal_task' && x.refId === taskId && !x.read", "false")] }] },
  { name: 'M11 通知帶金額', groups: ['api'], spec: [{ file: 'lib/maRoutes.js', edits: [E("' 已成交。請確認合約是否已簽訂，並在任務卡填寫合約編號與簽約日（或回報此案不簽合約）。'", "' 已成交，金額 ' + (task.amount || '8765432') + '。請確認。'")] }] },
  { name: 'M12 不合法期間不檢查', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("if (bad) return { ok: false, status: 400, code: bad.code, message: '成交需要有效的服務期間：' + bad.reason };", "")] }] },
  { name: 'M13 報價單來源不分最近', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("(approvedAtOf(b) < approvedAtOf(a) ? -1 : approvedAtOf(b) > approvedAtOf(a) ? 1 : 0)", "(approvedAtOf(a) < approvedAtOf(b) ? -1 : approvedAtOf(a) > approvedAtOf(b) ? 1 : 0)")] }] },
  { name: 'M14 離職行銷也通知', groups: ['unit', 'api'], spec: [{ file: 'lib/orgResolve.js', edits: [E("isActiveUser(u) && u.role === 'marketing'", "u.role === 'marketing'")] }] },
  { name: 'M15 一級主管不排除掛名主管', groups: ['unit'], spec: [{ file: 'lib/orgResolve.js', edits: [E("isActiveUser(u) && u.role === 'manager1' && !u.pureSupervisor", "isActiveUser(u) && u.role === 'manager1'")] }] },
  { name: 'M16 時區改用機器本地日', groups: ['unit'], spec: [{ file: 'lib/maTracking.js', edits: [E("return new Date((Number.isFinite(t) ? t : Date.now()) + 8 * 3600 * 1000).toISOString().slice(0, 10);", "return new Date((Number.isFinite(t) ? t : Date.now()) + 0 * 3600 * 1000).toISOString().slice(0, 10);")] }] },
  { name: 'M17 開關預設改成開', groups: ['unit', 'off'], spec: [{ file: 'lib/quoteTerms.js', edits: [E("return String(e.MA_TRACKING == null ? '' : e.MA_TRACKING).trim() === '1';", "return true;")] }] },
  { name: 'M19 合約刪除不重開任務卡（lib）', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("if (t && t.status === 'done') {", "if (false) {")] }] },
  { name: 'M20 拿掉改派驗證（角色＋BU，lib）', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("  const prob = assigneeProblem(data, t, u, users);\n  if (prob) return prob;", "  const prob = null;")] }] },
  { name: 'M21 拿掉 pending_sign 的任務條件', groups: ['unit', 'api', 'roles'], spec: [{ file: 'lib/maTracking.js', edits: [E("!it.contractNo && !!task && task.status === 'open';", "!it.contractNo;")] }] },
  { name: 'M22 改派只驗角色、不驗 BU', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("if (![...mine].some((b) => ref.has(b)))", "if (false)")] }] },
  { name: 'M23 合約重複刪除不冪等（lib 沒有連結時照樣當作有異動）', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("  if (!linked.length) return out;\n  out.touched = true;", "  out.touched = true;")] }] },
  { name: 'M24 候選人 API 不限管理員', groups: ['api'], spec: [{ file: 'lib/maRoutes.js', edits: [E("(req, res) => {\n    const L = loadTaskFor(req, res, true); if (!L) return;\n    const users = loadAuth().users || [];\n    res.set('Cache-Control', 'no-store');\n    res.json({ enabled: true, users:", "(req, res) => {\n    const L = loadTaskFor(req, res, false); if (!L) return;\n    const users = loadAuth().users || [];\n    res.set('Cache-Control', 'no-store');\n    res.json({ enabled: true, users:")] }] },
  { name: 'M25 合約刪除重開不重新通知（routes）', groups: ['api'], spec: [{ file: 'lib/maRoutes.js', edits: [E("taskNotify(r.task, 'contract_deleted');", "")] }] },
  { name: 'M18 合約刪除不解除連結', groups: ['unit', 'api'], spec: [{ file: 'lib/maTracking.js', edits: [E("it.contractNo = ''; it.contractId = ''; it.updatedAt = ctx.nowIso;\n    pushHist(it.history, H(ctx.nowIso, ctx.by, 'CONTRACT_DELETED'", "it.updatedAt = ctx.nowIso;\n    pushHist(it.history, H(ctx.nowIso, ctx.by, 'CONTRACT_DELETED'")] }] },
];
async function groupMut() {
  console.log('[變異] 每個變異都必須讓對應群組變紅');
  let undetected = 0;
  for (const m of MUTS) {
    const before = { p: pass, f: fail, n: failures.length };
    quiet = true; MUT_CUR = m.spec;
    try {
      for (const g of m.groups) {
        if (g === 'unit') {
          const edits = m.spec.filter((x) => x.file.startsWith('lib/'));
          const tmps = [];
          const Real = { MA: REF.MA };
          try {
            for (const x of edits) {
              const src = fs.readFileSync(path.join(ROOT, x.file), 'utf8');
              let mutated = src;
              x.edits.forEach((e) => { if (!mutated.includes(e.find)) throw new Error('MUTATION ANCHOR MISSING ' + e.find.slice(0, 50)); mutated = mutated.replace(e.find, () => e.rep); });
              const tmp = path.join(ROOT, 'lib', '__mut_' + path.basename(x.file));
              fs.writeFileSync(tmp, mutated); tmps.push(tmp);
            }
            // maTracking 被變異 → 用 __mut_ 副本；orgResolve／quoteTerms 被變異 → 清掉快取，讓 maTracking 重新載入時取到變異版
            if (edits.some((x) => x.file === 'lib/maTracking.js')) {
              const mutMa = path.join(ROOT, 'lib', '__mut_maTracking.js');
              delete require.cache[mutMa]; REF.MA = require(mutMa);
            }
            if (edits.some((x) => x.file === 'lib/orgResolve.js' || x.file === 'lib/quoteTerms.js')) {
              // 以變異後的 orgResolve／quoteTerms 取代並重載 maTracking 副本
              const sub = edits.filter((x) => x.file !== 'lib/maTracking.js');
              const maSrc = fs.readFileSync(path.join(ROOT, 'lib', 'maTracking.js'), 'utf8');
              let ms = maSrc;
              sub.forEach((x) => { ms = ms.replace(`require('./${path.basename(x.file, '.js')}')`, `require('./__mut_${path.basename(x.file)}')`); });
              const mutMa2 = path.join(ROOT, 'lib', '__mut_maTracking2.js');
              fs.writeFileSync(mutMa2, ms); tmps.push(mutMa2);
              REF.MA = require(mutMa2);
              // 單元測試裡直接用的 OR 也要變異
              if (sub.some((x) => x.file === 'lib/orgResolve.js')) REAL_OR_SWAP(require(path.join(ROOT, 'lib', '__mut_orgResolve.js')));
            }
            groupUnit();
          } finally {
            REF.MA = Real.MA; REAL_OR_SWAP(null);
            tmps.forEach((t) => { delete require.cache[t]; try { fs.unlinkSync(t); } catch { /* 忽略 */ } });
          }
        }
        if (g === 'api') await groupApi();
        if (g === 'roles') await groupRoles();
        if (g === 'off') await groupOff();
      }
    } catch (e) { fail++; failures.push('(變異造成例外) ' + e.message); }
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
function REAL_OR_SWAP(mod) { if (mod) { REAL._OR = REAL.OR; REAL.OR = mod; } else if (REAL._OR) { REAL.OR = REAL._OR; delete REAL._OR; } }

// ═════════════ 主程式 ═════════════
(async () => {
  try {
    if (want('unit')) groupUnit();
    if (want('api')) await groupApi();
    if (want('roles')) await groupRoles();
    if (want('off')) await groupOff();
    if (process.env.MATRK_MUT === '1') await groupMut();
  } catch (e) {
    console.log('例外：', e && e.stack || e); fail++; failures.push('例外：' + (e && e.message));
  } finally {
    await stopServer(); cleanupFiles();
    try { fs.unlinkSync(PRELOAD); } catch { /* 忽略 */ }
    try { fs.readdirSync(path.join(ROOT, 'lib')).filter((f) => f.startsWith('__mut_')).forEach((f) => fs.unlinkSync(path.join(ROOT, 'lib', f))); } catch { /* 忽略 */ }
  }
  console.log(`\n成交納管檢查結果：${pass} 通過，${fail} 失敗`);
  if (fail) console.log(failures.map((f) => ' - ' + f).join('\n'));
  process.exit(fail ? 1 : 0);
})();
