#!/usr/bin/env node
'use strict';
/**
 * 商機成交自動轉客戶（WON_AUTO_CUSTOMER）＋建立商機必須綁名片（OPP_REQUIRE_CONTACT）檢查。
 *   unit  lib/wonAutoCustomer.js 純函式（開關解析、轉換、冪等、resync 上限與寬鬆比對只報告）
 *   api   真實 server.js（隔離沙盒）：成交轉客戶／稽核／非 owner／冪等／回退不降級／Lead 轉換；管理員重新同步端點（dryRun／apply／權限／無個資）；
 *         建立商機綁名片（缺／空／不存在／已刪除／看不到／可見性與 GET /api/contacts 一致／KA 跨部門）；PUT 不可清空綁定、舊資料可編輯；Excel 匯入拒絕無名片列
 *   off   兩個開關=0：逐筆回應與基準版本（WON_BASE_REF，預設 origin/main）逐位元相同（去掉 id／時間）
 *   mut   變異測試：每個變異都必須讓 api 群組變紅
 *
 *   node scripts/check-won-auto-customer.js              unit + api + off
 *   WON_ONLY=unit,api,off,mut   只跑指定群組；WON_MUT=1 等同加上 mut
 *   WON_PORT（預設 3381，限 3380–3389）；WON_BASE_REF（預設 origin/main；設成空字串略過 off 群組並標示未驗證）
 * 只在臨時沙盒跑（git ls-files 複製目前工作樹；資料／帳號檔只寫在臨時資料夾；不碰工作樹的 data.json／auth.json／audit.log.json）。
 * 不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體；全部假資料。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn, spawnSync, execSync } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.WON_PORT || '3381', 10);
if (PORT < 3380 || PORT > 3389) { console.log('拒絕：埠需在 3380–3389'); process.exit(2); }
const BASE_REF = process.env.WON_BASE_REF === undefined ? 'origin/main' : process.env.WON_BASE_REF;
const ONLY = (process.env.WON_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => (!ONLY.length ? k !== 'mut' || process.env.WON_MUT === '1' : ONLY.includes(k));
const NODE_PATH = path.join(ROOT, 'node_modules');
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const XLSX = require(path.join(ROOT, 'node_modules', 'xlsx'));
const WONC = require('../lib/wonAutoCustomer');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let pass = 0, fail = 0; const fails = [];
function ok(c, name, extra) { if (c) pass++; else { fail++; fails.push(name); console.log('  ✗ ' + name + (extra ? '  ' + extra : '')); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), name, `got=${JSON.stringify(a)} want=${JSON.stringify(b)}`); }

const TMPS = [];
function mkSandbox() {
  const dst = fs.mkdtempSync(path.join(os.tmpdir(), 'won_run_')); TMPS.push(dst);
  const list = spawnSync('git', ['ls-files', '-co', '--exclude-standard', '-z'], { cwd: ROOT, encoding: 'utf8', maxBuffer: 1 << 28 }).stdout.split('\0').filter(Boolean);
  for (const f of list) {
    if (/^(scratchpad|node_modules)\//.test(f) || /(^|\/)(auth|data|audit\.log)\.jsonl?$/.test(f)) continue;
    const src = path.join(ROOT, f);
    if (!fs.existsSync(src)) continue;
    fs.mkdirSync(path.dirname(path.join(dst, f)), { recursive: true });
    fs.copyFileSync(src, path.join(dst, f));
  }
  return dst;
}
function mkBase(ref) {
  const dst = fs.mkdtempSync(path.join(os.tmpdir(), 'won_base_')); TMPS.push(dst);
  execSync(`git -C "${ROOT}" archive --format=tar ${ref} | tar -x`, { cwd: dst, stdio: 'ignore', maxBuffer: 1 << 28 });
  return dst;
}
function cleanAll() { for (const d of TMPS) { try { fs.rmSync(d, { recursive: true, force: true }); } catch { /* 忽略 */ } } }

// ───────────────── 沙盒資料 ─────────────────
const PW = 'Wn' + crypto.randomBytes(9).toString('hex') + '!Z';
const HASH = bcrypt.hashSync(PW, 6);
const BASEURL = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: HASH, createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canDownloadContacts: false, canSetTargets: false, accessMode: 'edit', bu: [] }, extra || {});
const ACCTS = [
  ['_s_adm', 'admin', {}],
  ['_s_m1', 'manager1', { bu: ['ERP'] }],
  ['_s_sec', 'secretary', { bu: ['ERP'] }],
  ['_s_mkt', 'marketing', {}],
  ['_s_tp', 'tecopm', { viewGroupId: 'g1', bu: ['ERP'] }],
  ['_s_u1', 'user', { bu: ['ERP'], supervisor: '_s_m1' }],
  ['_s_u2', 'user', { bu: ['ERP'], supervisor: '_s_m1' }],
  ['_s_u3', 'user', { bu: ['ITS'] }],
];
const PII = { phone: '0912345678', email: 'pii-secret@example.test', name: '機密姓名甲' };
const ct = (id, owner, company, extra) => Object.assign({ id, owner, name: '聯絡人' + id, company, phone: '', email: '', customerType: '', productLine: '', bu: 'ERP', createdAt: nowIso }, extra || {});
const op = (id, owner, contactId, stage, company, extra) => Object.assign({ id, owner, contactId, contactName: '', company, category: 'ERP', product: '產品', amount: '100', expectedDate: '2027-03-15', description: '', stage, createdAt: nowIso }, extra || {});
function seedData() {
  return {
    contacts: [
      ct('c1', '_s_u1', '甲公司'), ct('c2', '_s_u1', '甲公司'),
      ct('c3', '_s_u2', '乙公司'),
      ct('c4', '_s_u1', '丙公司', { customerType: 'customer', productLine: 'ITS' }),
      ct('c5', '_s_u1', '丁公司', { deleted: true }),
      ct('c6', '_s_u3', '戊公司', { bu: 'ITS' }),
      ct('c7', '_s_u1', '己公司', { isPlaceholder: true }),
      ct('c8', '_s_u1', '庚公司', Object.assign({ name: PII.name, phone: PII.phone, email: PII.email })),
      ct('c9', '_s_u2', '甲乙丙有限公司'),
      ct('c10', '_s_u1', '集團內公司'),
      ct('c11', '_s_u1', '臺灣ＡＢＣ股份有限公司'),   // 全形英文＋「臺」
      ct('c12', '_s_u1', '甲公司'),
    ],
    companies: [],
    opportunities: [
      op('o1', '_s_u1', 'c1', 'B', '甲公司'),
      op('o2', '_s_u1', 'c3', 'B', '乙公司'),      // 綁別人（u2）的名片
      op('o3', '_s_u1', '', 'B', '舊商機公司'),    // 舊資料：沒有名片
      op('o4', '_s_u1', 'ghost', 'B', '幽靈公司'), // 舊資料：綁到不存在的名片
      op('o5', '_s_u1', 'c4', 'B', '丙公司'),
      op('o6', '_s_u1', 'c5', 'B', '丁公司'),      // 綁已刪除名片
      op('o7', '_s_u1', 'c7', 'B', '己公司'),      // 綁待補名片
      op('oE', '_s_u1', 'c1', 'B', ''),            // 舊資料：已綁名片但公司是空的
      op('oE2', '_s_u1', '', 'B', ''),             // 舊資料：沒名片也沒公司
      op('oR', '_s_u1', 'c1', 'Won', '甲公司'),    // 已成交（c1 在 W1 之後已是客戶）
      op('oN', '_s_u1', '', 'B', '', { contactName: '機密姓名乙' }),   // 無公司、無名片、有人名
      op('oW1', '_s_u1', 'c8', 'Won', '庚公司'),   // 已成交、名片仍是潛在
      op('oW2', '_s_u1', '', 'Won', '甲乙丙股份有限公司'),   // 已成交、沒綁名片（寬鬆比對會找到 c9）
      op('oW3', '_s_u1', 'c4', 'Won', '丙公司'),   // 已成交、名片已是客戶
    ],
    visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [{ company: '戊公司', owner: '_s_u3' }],
    quotations: [], leads: [{ id: 'l1', company: '潛在客戶股份有限公司', contactName: '王小明', assignedTo: '_s_u1', status: 'new', campaignName: 'C', interest: 'ERP', createdAt: nowIso }],
    groups: [{ id: 'g1', name: '測試集團', memberCompanies: ['集團內公司'] }],
  };
}
function seedFiles(dir) {
  fs.writeFileSync(path.join(dir, 'auth.json'), JSON.stringify({ users: ACCTS.map((a) => mkUser(a[0], a[1], a[2])) }, null, 2));
  fs.writeFileSync(path.join(dir, 'data.json'), JSON.stringify(seedData(), null, 2));
}
function cleanupFiles(dir) { ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].forEach((f) => { try { fs.unlinkSync(path.join(dir, f)); } catch { /* 不存在 */ } }); }

let child = null, jar = {};
async function startServer(root, extraEnv) {
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', NODE_PATH, ALLOWED_ORIGINS: `http://localhost:${PORT}`, AUDIT_EVENTS_TABLE: '1' });   // 稽核走 append-only audit.log.jsonl（無 500 筆上限，方便斷言）
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'FEATURE_GUARD_MODE', 'LOGIN_LOCKOUT', 'WON_AUTO_CUSTOMER', 'OPP_REQUIRE_CONTACT']) delete env[k];
  Object.assign(env, extraEnv || {});
  await assertPortFree(PORT);
  child = spawn(process.execPath, [path.join(root, 'server.js')], { cwd: root, env, stdio: ['ignore', 'pipe', 'pipe'] });
  let out = '';
  child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { out += d; });
  for (let i = 0; i < 100; i++) {
    await sleep(250);
    try { const r = await fetch(`${BASEURL}/login.html`); if (r.ok) return; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-600));
  }
  throw new Error('伺服器啟動逾時：' + out.slice(-600));
}
async function stopServer() {
  if (!child) return;
  const c = child; child = null;
  try { c.kill(); } catch { /* 已結束 */ }
  for (let i = 0; i < 30 && c.exitCode === null; i++) await sleep(100);
  await sleep(300);
}
async function freshServer(root, extraEnv) { await stopServer(); cleanupFiles(root); seedFiles(root); await startServer(root, extraEnv); jar = {}; }
async function login(who) {
  const res = await fetch(BASEURL + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status === 200) jar[who] = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
  return res.status;
}
async function call(who, method, url, body) {
  const headers = { Cookie: (who && jar[who]) || '' };
  if (body !== undefined) headers['Content-Type'] = 'application/json';
  const res = await fetch(BASEURL + url, { method, headers, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const text = await res.text();
  let json = null; try { json = JSON.parse(text); } catch { /* 非 JSON */ }
  return { status: res.status, json, text };
}
const contactsOf = async (who) => { const r = await call(who, 'GET', '/api/contacts'); return Array.isArray(r.json) ? r.json : []; };
const ctype = async (who, id) => { const c = (await contactsOf(who)).find((x) => x.id === id); return c ? c.customerType : undefined; };
function auditLines(root) {
  try { return fs.readFileSync(path.join(root, 'audit.log.jsonl'), 'utf8').split('\n').filter(Boolean).map((l) => { try { return JSON.parse(l); } catch { return null; } }).filter(Boolean); } catch { return []; }
}
async function auditCount(root, action, needle) {
  await sleep(350);
  return auditLines(root).filter((e) => e.action === action && (!needle || String(e.detail || '').includes(needle))).length;
}
const putOpp = (who, id, body) => call(who, 'PUT', '/api/opportunities/' + id, body);
const postOpp = (who, body) => call(who, 'POST', '/api/opportunities', Object.assign({ company: '甲公司', category: 'ERP', product: '產品', amount: '10', expectedDate: '2027-05-01' }, body));

// ───────────────── unit ─────────────────
function unitGroup() {
  console.log('  [unit] lib/wonAutoCustomer.js');
  ok(WONC.enabled({}) === true, 'U1 未設＝啟用');
  for (const v of ['0', 'off', 'OFF', 'false', ' 0 ']) ok(WONC.enabled({ WON_AUTO_CUSTOMER: v }) === false, 'U2 關閉值 ' + JSON.stringify(v));
  for (const v of ['1', 'on', '', 'yes']) ok(WONC.enabled({ WON_AUTO_CUSTOMER: v }) === true, 'U3 其他值＝啟用 ' + JSON.stringify(v));
  const d = { contacts: [{ id: 'a', customerType: '' }, { id: 'b', customerType: 'customer', productLine: 'ERP' }, { id: 'd', deleted: true }, { id: 'p', isPlaceholder: true }] };
  eq(WONC.convertForWonOpp(d, { contactId: 'a' }), { status: 'converted', contactId: 'a', changed: true, prevType: '' }, 'U4 轉換');
  ok(!!d.contacts[0].updatedAt, 'U4c 轉換時設 updatedAt');
  eq(d.contacts[0].customerType, 'customer', 'U4b 名片被寫入');
  eq(WONC.convertForWonOpp(d, { contactId: 'a' }).status, 'already', 'U5 冪等');
  const before = JSON.stringify(d.contacts[1]);
  eq(WONC.convertForWonOpp(d, { contactId: 'b' }).status, 'already', 'U6 已是客戶');
  eq(JSON.stringify(d.contacts[1]), before, 'U6b 已是客戶的名片逐位元不變（productLine 保留）');
  eq(WONC.convertForWonOpp(d, { contactId: '' }).reason, 'no-binding', 'U7 未綁定');
  eq(WONC.convertForWonOpp(d, { contactId: 'zz' }).reason, 'not-found', 'U8 不存在');
  eq(WONC.convertForWonOpp(d, { contactId: 'd' }).reason, 'deleted', 'U9 已刪除');
  eq(WONC.convertForWonOpp(d, { contactId: 'p' }).reason, 'placeholder', 'U10 待補名片');
  eq(d.contacts[2].customerType, undefined, 'U10b 已刪除名片不被改');
  // resync：上限與寬鬆比對只報告
  const big = { contacts: [], opportunities: [] };
  for (let i = 0; i < 250; i++) { big.contacts.push({ id: 'k' + i, company: '公司' + i, customerType: '' }); big.opportunities.push({ id: 'p' + i, stage: 'Won', contactId: 'k' + i, company: '公司' + i }); }
  const r = WONC.resync(big, { apply: false, cap: 200 });
  eq([r.candidates.length, r.candidateTotal, r.truncated], [200, 250, true], 'U11 dryRun 清單上限 200、總數 250、標示截斷');
  ok(big.contacts.every((c) => c.customerType === ''), 'U12 dryRun 不寫入');
  const r2 = WONC.resync(big, { apply: true, cap: 200 });
  eq(r2.converted, 250, 'U13 apply 轉換全部（含超出清單上限者）');
  eq(WONC.resync(big, { apply: true }).converted, 0, 'U14 重跑冪等');
  const lo = { contacts: [{ id: 'x', company: '甲乙丙有限公司', customerType: '' }], opportunities: [{ id: 'q', stage: 'Won', contactId: '', company: '甲乙丙股份有限公司' }] };
  const r3 = WONC.resync(lo, { apply: true, normalizeLoose: (s) => String(s).replace(/股份有限公司|有限公司/g, '') });
  eq([r3.unboundTotal, r3.unbound[0].looseMatchContactIds, r3.converted, lo.contacts[0].customerType], [1, ['x'], 0, ''], 'U15 寬鬆比對只報告、apply 也不擴散');
  for (const e of r3.unbound.concat(r.candidates)) ok(!Object.keys(e).some((k) => /phone|email|name$|mobile/i.test(k) && k !== 'company'), 'U16 報告列不含個資欄位 ' + Object.keys(e).join(','));
}

// ───────────────── api ─────────────────
async function apiGroup(root, tag) {
  console.log('  [api] ' + (tag || '成交轉客戶＋商機綁名片'));
  await freshServer(root);
  for (const w of ACCTS.map((a) => a[0])) if (await login(w) !== 200) throw new Error('登入失敗 ' + w);

  // ── 成交轉客戶 ──
  eq(await ctype('_s_u1', 'c1'), '', 'W0 初始：c1 為潛在客戶');
  let r = await putOpp('_s_u1', 'o1', { stage: 'Won' });
  eq(r.status, 200, 'W1 業務把 o1 拖到 Won → 200');
  eq(await ctype('_s_u1', 'c1'), 'customer', 'W1b 綁定名片轉為客戶');
  eq(await ctype('_s_u1', 'c2'), '', 'W1c 同公司其他名片不靜默擴散（維持前端既有規則）');
  const c1 = (await contactsOf('_s_u1')).find((c) => c.id === 'c1');
  eq(c1.productLine, '', 'W1d 不要求產品線：productLine 維持原樣');
  eq(await auditCount(root, 'WON_AUTO_CUSTOMER', 'c1'), 1, 'W1e 稽核：自動轉換 1 筆（含名片 id 與操作者）');
  ok(auditLines(root).filter((e) => e.action === 'WON_AUTO_CUSTOMER').every((e) => /操作者 _s_u1/.test(e.detail) && !/聯絡人c1|\d{9,}|@/.test(e.detail)), 'W1f 稽核內容含操作者、不含姓名／電話／Email');
  // 冪等與回退
  r = await putOpp('_s_u1', 'o1', { stage: 'Won' });
  eq(await auditCount(root, 'WON_AUTO_CUSTOMER', 'c1'), 1, 'W2 Won→Won 不重複寫、不重複稽核');
  r = await putOpp('_s_u1', 'o1', { stage: 'B' });
  eq(await ctype('_s_u1', 'c1'), 'customer', 'W3 Won 回退 B 不降級客戶');
  r = await putOpp('_s_u1', 'o1', { stage: 'Won' });
  eq(r.status, 200, 'W4 B→Won 再成交 200');
  eq(await auditCount(root, 'WON_AUTO_CUSTOMER', 'c1'), 1, 'W4b 名片已是客戶 → 不重寫、不重複稽核');
  // 非名片 owner 的操作者（名片屬 u2、商機屬 u1，由主管成交）
  r = await putOpp('_s_m1', 'o2', { stage: 'Won' });
  eq(r.status, 200, 'W5 主管（非名片 owner）把 o2 成交 → 200');
  eq(await ctype('_s_adm', 'c3'), 'customer', 'W5b 非名片 owner 操作者也生效');
  eq(await auditCount(root, 'WON_AUTO_CUSTOMER', '操作者 _s_m1'), 1, 'W5c 稽核記到操作者');
  // 已是客戶：不動
  r = await putOpp('_s_u1', 'o5', { stage: 'Won' });
  const c4 = (await contactsOf('_s_u1')).find((c) => c.id === 'c4');
  eq([c4.customerType, c4.productLine], ['customer', 'ITS'], 'W6 已是客戶：productLine 保留');
  eq(await auditCount(root, 'WON_AUTO_CUSTOMER', 'c4'), 0, 'W6b 已是客戶不產生稽核');
  // 無綁定／不存在／已刪除／待補名片：不報錯、不阻擋、有稽核
  for (const [oid, why] of [['o3', '未綁定'], ['o4', '不存在'], ['o6', '已刪除'], ['o7', '待補']]) {
    r = await putOpp('_s_u1', oid, { stage: 'Won' });
    eq(r.status, 200, `W7 ${why}名片的商機成交 → 200（不阻擋）`);
    eq(r.json && r.json.stage, 'Won', `W7b ${why}：階段確實變 Won`);
    eq(await auditCount(root, 'WON_AUTO_CUSTOMER_PENDING', oid), 1, `W7c ${why}：有「待人工」稽核`);
  }
  eq(await ctype('_s_adm', 'c5'), undefined, 'W7d 已刪除名片仍不在清單');
  // 「我的客戶」判定一致：app.js 以 customerType==='customer' 判斷，後端寫入的欄位與之相同
  const appJs = fs.readFileSync(path.join(root, '_client', 'app.js'), 'utf8');
  ok(/allContacts\.filter\(c => c\.customerType === 'customer'/.test(appJs), 'W8 前端「我的客戶」判定＝customerType==="customer"（後端寫的就是這個欄位）');
  // Lead 轉換直接成交
  r = await call('_s_adm', 'POST', '/api/leads/l1/convert', { salesPerson: '_s_u1', stage: 'Won', product: '產品', category: 'ERP' });
  eq(r.status, 200, 'W9 Lead 轉換（stage=Won）→ 200');
  const newCid = r.json && r.json.contactId;
  eq(await ctype('_s_adm', newCid), 'customer', 'W9b Lead 轉換先建名片再綁、成交即為客戶');
  ok(r.json && r.json.opportunity && r.json.opportunity.contactId === newCid, 'W9c Lead 轉出的商機綁定了該名片');

  // ── 管理員重新同步 ──
  const dry = await call('_s_adm', 'POST', '/api/admin/won-resync-customers', {});
  eq(dry.status, 200, 'S1 管理員 dryRun（空 body）→ 200');
  eq([dry.json.apply, dry.json.candidateTotal], [false, 1], 'S1b dryRun：待轉 1 張（c8）');
  eq(dry.json.candidates[0].contactId, 'c8', 'S1c 清單列出 c8');
  eq(dry.json.unbound.length >= 2, true, 'S1d 另列出「Won 但沒綁名片」的商機（含 oW2 與前面 W7 的無名片商機）');
  const u2 = dry.json.unbound.find((x) => x.oppId === 'oW2');
  eq(u2 && u2.looseMatchContactIds, ['c9'], 'S1e 寬鬆比對只報告疑似名片 id（c9）');
  eq(await ctype('_s_adm', 'c8'), '', 'S2 dryRun 不寫入');
  for (const b of [{ apply: false }, { apply: 'true' }, { apply: 1 }, { apply: 'yes' }]) {
    await call('_s_adm', 'POST', '/api/admin/won-resync-customers', b);
  }
  eq(await ctype('_s_adm', 'c8'), '', 'S2b apply 不是布林 true（false／"true"／1／"yes"）一律不寫入');
  const dtext = JSON.stringify(dry.json);
  ok(!dtext.includes(PII.phone) && !dtext.includes(PII.email) && !dtext.includes(PII.name), 'S3 dryRun 回應不含電話／Email／姓名');
  ok(dry.json.candidates.every((x) => Object.keys(x).sort().join() === 'company,contactId,oppId,reason'), 'S3b 清單欄位白名單');
  const ap = await call('_s_adm', 'POST', '/api/admin/won-resync-customers', { apply: true });
  eq([ap.status, ap.json.converted], [200, 1], 'S4 apply:true → 轉換 1 張');
  eq(await ctype('_s_adm', 'c8'), 'customer', 'S4b c8 已是客戶');
  eq(await ctype('_s_adm', 'c9'), '', 'S4c 寬鬆比對找到的 c9 不被靜默擴散');
  const ap2 = await call('_s_adm', 'POST', '/api/admin/won-resync-customers', { apply: true });
  eq(ap2.json.converted, 0, 'S4d 再跑 apply 冪等（0）');
  eq(await auditCount(root, 'WON_RESYNC_CUSTOMERS'), 2, 'S4e 兩次 apply 各有稽核');
  ok((await auditCount(root, 'WON_RESYNC_CUSTOMERS_DRYRUN')) >= 1, 'S4f dryRun 也有稽核');
  for (const w of ['_s_m1', '_s_u1', '_s_sec', '_s_mkt', '_s_tp']) {
    const x = await call(w, 'POST', '/api/admin/won-resync-customers', { apply: true });
    eq(x.status, 403, 'S5 非管理員 ' + w + ' → 403');
  }
  eq((await call(null, 'POST', '/api/admin/won-resync-customers', { apply: true })).status, 401, 'S6 未登入 → 401');

  // ── 建立商機必須綁名片 ──
  const NOCODE = (x) => x.json && x.json.code;
  for (const [label, body] of [['缺 contactId', {}], ['空字串', { contactId: '' }], ['全空白', { contactId: '   ' }], ['null', { contactId: null }]]) {
    const x = await postOpp('_s_u1', body);
    eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_REQUIRED'], 'P1 ' + label + ' → 400 OPP_CONTACT_REQUIRED');
    ok(/聯絡人/.test((x.json && x.json.error) || ''), 'P1b ' + label + ' 訊息為繁中且提到聯絡人');
  }
  for (const [label, v] of [['數字', 123], ['陣列', ['c1']], ['物件', { id: 'c1' }]]) {
    const y = await postOpp('_s_u1', { contactId: v });
    eq([y.status, y.json && y.json.code], [400, 'OPP_CONTACT_INVALID'], 'P1c POST contactId 為' + label + ' → 400 OPP_CONTACT_INVALID');
  }
  let x = await postOpp('_s_u1', { contactId: 'ghost' });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_NOT_FOUND'], 'P2 不存在 id → 400');
  x = await postOpp('_s_u1', { contactId: 'c5' });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_NOT_FOUND'], 'P2b 已刪除名片 → 400');
  x = await postOpp('_s_u1', { contactId: 'c3' });
  eq([x.status, NOCODE(x)], [403, 'OPP_CONTACT_FORBIDDEN'], 'P3 他人（u2）名片 → 403');
  x = await postOpp('_s_u1', { contactId: 'c9' });
  eq(x.status, 403, 'P3b 他人名片 c9 → 403');
  x = await postOpp('_s_u1', { contactId: 'c6', company: '戊公司' });
  eq(x.status, 201, 'P4 KA 公司的跨部門名片（c6，GET /api/contacts 看得到）→ 201（KA 流程不被破壞）');
  eq(x.json && x.json.kaNotified !== undefined, true, 'P4b KA 通知流程照常回報 kaNotified');
  x = await postOpp('_s_u1', { contactId: 'c1' });
  eq(x.status, 201, 'P5 正常綁定自己的名片 → 201');
  eq(x.json && x.json.contactId, 'c1', 'P5b 商機寫入 contactId');
  x = await postOpp('_s_m1', { contactId: 'c1' });
  eq(x.status, 201, 'P6 主管綁部屬名片 → 201（沿用可視範圍）');
  // 可見性不變式：能看到（GET /api/contacts）＝能綁；看不到＝不能 201
  const CO = { c1: '甲公司', c2: '甲公司', c3: '乙公司', c4: '丙公司', c6: '戊公司', c8: '庚公司', c9: '甲乙丙有限公司', c10: '集團內公司' };
  const IDS = Object.keys(CO);
  for (const w of ['_s_adm', '_s_m1', '_s_u1', '_s_u2', '_s_u3', '_s_mkt', '_s_sec', '_s_tp']) {
    const vis = new Set((await contactsOf(w)).map((c) => c.id));
    for (const id of IDS) {
      const y = await postOpp(w, { contactId: id, company: CO[id] });
      if (vis.has(id)) ok(!/^OPP_CONTACT/.test(NOCODE(y) || ''), `P7 ${w} 看得到 ${id} → 不被聯絡人檢查擋（${y.status} ${NOCODE(y) || ''}）`);
      else ok(y.status !== 201, `P7 ${w} 看不到 ${id} → 不可建立（${y.status}）`);
    }
  }
  ok((await contactsOf('_s_u1')).some((c) => c.id === 'c6'), 'P7z KA 公司的跨部門名片 c6 對 u1 確實可見（P4 走的是 KA 路徑）');

  // ── 公司一致性（寬鬆比對）／contactId 正規化 ──
  x = await postOpp('_s_u1', { contactId: 'c1', company: '完全不同公司' });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_COMPANY_MISMATCH'], 'Q1 名片公司與商機公司不一致 → 400 OPP_CONTACT_COMPANY_MISMATCH');
  ok(/不一致/.test((x.json && x.json.error) || ''), 'Q1b 訊息為繁中');
  x = await postOpp('_s_u1', { contactId: 'c1', company: '甲公司股份有限公司' });
  eq(x.status, 201, 'Q2 寬鬆比對：「甲公司股份有限公司」＝「甲公司」→ 201');
  x = await postOpp('_s_u1', { contactId: 'c1', company: '  甲公司 ' });
  eq(x.status, 201, 'Q2b 公司前後空白不影響比對');
  x = await postOpp('_s_u1', { contactId: 'c1', company: '' });
  eq([x.status, x.json && x.json.company], [201, '甲公司'], 'Q3 商機公司空白 → 由名片帶入（不誤擋）');
  x = await postOpp('_s_u1', { contactId: ' c1 ' });
  eq([x.status, x.json && x.json.contactId], [201, 'c1'], 'Q4 POST contactId 前後空白 → 存 trim 後的 "c1"');
  const oq4 = x.json && x.json.id;
  x = await call('_s_u1', 'POST', '/api/opportunities', { contactId: 'c1', category: 'ERP', product: '產品', expectedDate: '2027-05-01' });
  eq([x.status, x.json && x.json.company], [201, '甲公司'], 'Q4b 完全沒帶 company → 由名片帶入');
  x = await putOpp('_s_u1', oq4, { contactId: ['c1'] });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_INVALID'], 'Q5 PUT contactId 陣列 → 400（即使與現值同字串）');
  x = await putOpp('_s_u1', oq4, { contactId: { a: 1 } });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_INVALID'], 'Q5b PUT contactId 物件 → 400');
  x = await putOpp('_s_u1', oq4, { contactId: 7 });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_INVALID'], 'Q5c PUT contactId 數字 → 400');
  x = await putOpp('_s_u1', oq4, { contactId: ' c2 ' });
  eq([x.status, x.json && x.json.contactId], [200, 'c2'], 'Q6 PUT contactId 前後空白 → 存 trim 後的 "c2"');
  x = await putOpp('_s_u1', oq4, { contactId: ' c2 ', description: 'same' });
  eq([x.status, x.json && x.json.contactId], [200, 'c2'], 'Q6b PUT 帶空白的現值 → 視為未改、存 "c2"');
  x = await putOpp('_s_u1', oq4, { contactId: 'c4' });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_COMPANY_MISMATCH'], 'Q7 PUT 改綁別家公司的名片 → 400 公司不一致');
  x = await putOpp('_s_u1', oq4, { company: '別家公司' });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_COMPANY_MISMATCH'], 'Q8 PUT 把公司改成與綁定名片不同 → 400');
  x = await putOpp('_s_u1', oq4, { contactId: 'c4', company: '丙公司' });
  eq([x.status, x.json && x.json.contactId], [200, 'c4'], 'Q9 PUT 同時改名片與公司且一致 → 200');
  x = await putOpp('_s_u1', 'o4', { company: '改名的幽靈公司', description: 'z' });
  eq(x.status, 200, 'Q10 舊的幽靈綁定商機改公司名：名片不存在 → 略過比對（不誤擋）');
  x = await putOpp('_s_u1', 'o3', { company: '舊商機改名公司' });
  eq(x.status, 200, 'Q10b 舊的無名片商機改公司名 → 200');
  const me = await call('_s_u1', 'GET', '/api/me');
  ok(me.json && !('oppRequireContact' in me.json), 'Q11 開關開著：/api/me 不多帶旗標（輸出與舊版相同）');

  // ── 公司清空繞過／補公司／全形與臺台／成交後改綁 ──
  x = await postOpp('_s_u1', { contactId: 'c1', company: '甲公司' });
  const ob = x.json && x.json.id;
  x = await putOpp('_s_u1', ob, { company: '' });
  eq([x.status, NOCODE(x)], [400, 'OPP_COMPANY_REQUIRED'], 'R1 已綁名片的商機不可把公司清成空白 → 400');
  x = await putOpp('_s_u1', ob, { company: '   ' });
  eq([x.status, NOCODE(x)], [400, 'OPP_COMPANY_REQUIRED'], 'R1b 清成全空白同樣擋');
  x = await putOpp('_s_u1', ob, { contactId: 'c4' });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_COMPANY_MISMATCH'], 'R1c 繞過路徑（先清空再改綁別家名片）的第二步仍被公司檢查擋住');
  x = await putOpp('_s_u1', 'oE', { company: '', description: 'legacy' });
  eq(x.status, 200, 'R2 舊資料（已綁名片、公司本來就空）原樣送回公司空白 → 不誤擋');
  x = await putOpp('_s_u1', 'oE', { contactId: 'c2' });
  eq([x.status, x.json && x.json.company], [200, '甲公司'], 'R2b 綁定（改 contactId）時商機公司是空的 → 用名片公司補上');
  x = await putOpp('_s_u1', 'oE2', { contactId: 'c4', company: '' });
  eq([x.status, x.json && x.json.company], [200, '丙公司'], 'R2c 舊的無名片無公司商機補綁：公司空白（原樣送回）→ 由名片補上');
  x = await postOpp('_s_u1', { contactId: 'c11', company: '台灣ABC股份有限公司' });
  eq(x.status, 201, 'R3 全形英文＋臺／台＋公司後綴：「臺灣ＡＢＣ股份有限公司」＝「台灣ABC股份有限公司」');
  x = await postOpp('_s_u1', { contactId: 'c11', company: '台灣ＡＢＣ（股）' });
  eq(x.status, 201, 'R3b 全形括號「（股）」與全形英文可比對');
  x = await postOpp('_s_u1', { contactId: 'c11', company: '台湾ABC股份有限公司' });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_COMPANY_MISMATCH'], 'R3c 簡體「台湾」不是同一家 → 仍擋');
  // 成交後改綁別張名片：轉新名片、舊名片不降級、有名片稽核（無姓名）、updatedAt
  eq(await ctype('_s_u1', 'c12'), '', 'R4 前置：c12 是潛在客戶');
  x = await putOpp('_s_u1', 'oR', { description: 'won edit' });
  eq(await auditCount(root, 'WON_AUTO_CUSTOMER', 'c12'), 0, 'R4a 已是 Won、只改備註 → 不觸發');
  x = await putOpp('_s_u1', 'oR', { contactId: 'c12' });
  eq(x.status, 200, 'R4b 已 Won 的商機改綁 c12 → 200');
  eq(await ctype('_s_u1', 'c12'), 'customer', 'R4c 轉為新綁定的名片');
  eq(await ctype('_s_u1', 'c1'), 'customer', 'R4d 舊名片不降級');
  eq(await auditCount(root, 'WON_AUTO_CUSTOMER', 'c12'), 1, 'R4e 有 WON_AUTO_CUSTOMER 稽核');
  ok(!!((await contactsOf('_s_u1')).find((c) => c.id === 'c12') || {}).updatedAt, 'R4f 名片 updatedAt 已設');
  const cAud = auditLines(root).filter((e) => e.kind === 'contact' && e.target === 'c12');
  eq(cAud.length, 1, 'R4g 有名片欄位稽核（writeContactAudit）1 筆');
  ok(cAud.length && JSON.stringify(cAud[0]).indexOf('聯絡人c12') < 0 && (cAud[0].extra || {}).targetName === '' && /customerType/.test(JSON.stringify(cAud[0].extra.changes)), 'R4h 名片稽核不含姓名、記錄 customerType 變更', JSON.stringify(cAud[0] || {}).slice(0, 300));
  x = await putOpp('_s_u1', 'oR', { contactId: 'c12' });
  eq(await auditCount(root, 'WON_AUTO_CUSTOMER', 'c12'), 1, 'R4i 再送同名片 → 不重複稽核（冪等）');
  // 稽核 target 不退用人名
  x = await putOpp('_s_u1', 'oN', { stage: 'Won' });
  eq(x.status, 200, 'R5 無公司、無名片、有人名的商機成交 → 200');
  await sleep(350);
  const pend = auditLines(root).filter((e) => e.action === 'WON_AUTO_CUSTOMER_PENDING' && String(e.detail || '').includes('oN'));
  ok(pend.length === 1 && pend[0].target === '（無公司名）' && JSON.stringify(pend[0]).indexOf('機密姓名乙') < 0, 'R5b 稽核 target 為固定字樣、不含人名', JSON.stringify(pend[0] || {}).slice(0, 200));

  // ── PUT：不可清空／改成不存在；舊資料可編輯 ──
  x = await putOpp('_s_u1', 'o1', { description: 'x' });
  eq([x.status, x.json && x.json.contactId], [200, 'c1'], 'T1 一般欄位編輯不受影響');
  x = await putOpp('_s_u1', 'o1', { contactId: '' });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_REQUIRED'], 'T2 已綁的商機不可清空 contactId');
  x = await putOpp('_s_u1', 'o1', { contactId: 'ghost' });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_NOT_FOUND'], 'T3 不可改成不存在的名片');
  x = await putOpp('_s_u1', 'o1', { contactId: 'c9' });
  eq(x.status, 403, 'T4 不可改綁看不到的他人名片');
  x = await putOpp('_s_u1', 'o1', { contactId: 'c2' });
  eq([x.status, x.json && x.json.contactId], [200, 'c2'], 'T5 可改綁自己看得到的另一張名片');
  // 舊資料 o3（無名片，已在 W7 成交）、o4（幽靈）：其他欄位照常
  x = await putOpp('_s_u1', 'o3', { description: 'legacy' });
  eq(x.status, 200, 'T6 舊的無名片商機可編輯其他欄位');
  x = await putOpp('_s_u1', 'o3', { contactId: '', description: 'again' });
  eq(x.status, 200, 'T7 舊商機整份送回（contactId 與現值同為空）不報錯');
  x = await putOpp('_s_u1', 'o4', { contactId: 'ghost', description: 'g' });
  eq(x.status, 200, 'T8 舊的幽靈綁定原樣送回不報錯');
  x = await putOpp('_s_u1', 'o3', { contactId: 'c1' });
  eq([x.status, NOCODE(x)], [400, 'OPP_CONTACT_COMPANY_MISMATCH'], 'T9 舊商機補綁但公司名不一致 → 400');
  x = await putOpp('_s_u1', 'o3', { contactId: 'c1', company: '甲公司' });
  eq([x.status, x.json && x.json.contactId], [200, 'c1'], 'T9b 舊商機補綁一張有效名片並改成同公司 → 200');

  // ── Excel 匯入（管理員）：無名片列拒絕並列出原因 ──
  const aoa = [['客戶名稱', '銷售案名', 'BU', '預定簽約日', '業務帳號', '把握度', '合約金額'],
    ['甲公司', '案A', 'ERP', '2027-06-01', '_s_u1', 'B', 100],
    ['完全沒有名片的公司', '案B', 'ERP', '2027-06-01', '_s_u1', 'B', 100],
    ['甲乙丙有限公司', '案C', 'ERP', '', '_s_u2', 'Won', 100],
    ['歸客戶池的成交公司', '案D', 'ERP', '', '_pool', 'Won', 100]];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'S');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  async function imp(dry) {
    const fd = new FormData(); fd.append('file', new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'opps.xlsx'); if (dry) fd.append('dryRun', '1');
    const res = await fetch(BASEURL + '/api/admin/opportunities/import', { method: 'POST', headers: { Cookie: jar._s_adm }, body: fd });
    let j = null; try { j = await res.json(); } catch { /* 無 */ } return { status: res.status, json: j };
  }
  const id1 = await imp(true);
  eq([id1.status, id1.json && id1.json.wouldCreate, (id1.json && id1.json.noContactErrors || []).length], [200, 3, 1], 'X1 匯入預檢：3 列可匯入、1 列無名片被拒（列出原因）');
  eq([id1.json && id1.json.wonAutoConverted, id1.json && id1.json.wonPending], [1, 1], 'X1c 預檢回報：成交列會自動轉客戶 1 筆、待人工 1 筆');
  eq(await ctype('_s_adm', 'c9'), '', 'X1d 預檢不寫入（c9 仍是潛在客戶）');
  ok(/第3列.*找不到.*名片/.test(((id1.json && id1.json.noContactErrors) || [])[0] || ''), 'X1b 原因含列號與說明');
  const im = await imp(false);
  eq([im.status, im.json && im.json.created, (im.json && im.json.noContactErrors || []).length], [200, 3, 1], 'X2 實際匯入：建 3 筆、無名片列被拒並回報');
  eq([im.json && im.json.wonAutoConverted, im.json && im.json.wonPending], [1, 1], 'X2b 匯入結果回傳：自動轉客戶 1 筆、待人工 1 筆');
  eq(await ctype('_s_adm', 'c9'), 'customer', 'X2c 匯入成交列：綁定的名片（c9）當下轉為客戶');
  eq(await auditCount(root, 'WON_AUTO_CUSTOMER_IMPORT'), 1, 'X2d 匯入稽核彙總一筆');
  eq(await auditCount(root, 'WON_AUTO_CUSTOMER_PENDING', '待補名片'), 2, 'X2e 無可轉名片的成交列記「待人工」（共 2 筆：前面 W7 的 o7＋這次匯入的客戶池列）');
  const opps = (await call('_s_adm', 'GET', '/api/opportunities')).json || [];
  ok(!opps.some((o) => o.company === '完全沒有名片的公司') && opps.some((o) => o.company === '甲公司' && o.product === '案A' && o.contactId), 'X3 被拒的列沒寫入、有名片的列已綁 contactId');
}

// ───────────────── off：兩個開關=0 與基準逐位元相同 ─────────────────
const UUID = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/g;
const ISO = /\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/g;
const norm = (v) => JSON.stringify(v).replace(UUID, '<id>').replace(ISO, '<t>');
async function scenario(root, env) {
  await freshServer(root, env);
  for (const w of ['_s_adm', '_s_m1', '_s_u1', '_s_u2']) await login(w);
  const out = [];
  const rec = (label, r) => out.push([label, r.status, norm(r.json !== null ? r.json : r.text.slice(0, 200))]);
  rec('POST 無 contactId', await postOpp('_s_u1', {}));
  rec('POST 空字串', await postOpp('_s_u1', { contactId: '' }));
  rec('POST 不存在', await postOpp('_s_u1', { contactId: 'ghost' }));
  rec('POST 他人名片', await postOpp('_s_u1', { contactId: 'c3' }));
  rec('POST 正常', await postOpp('_s_u1', { contactId: 'c1' }));
  rec('PUT 清空 contactId', await putOpp('_s_u1', 'o1', { contactId: '' }));
  rec('PUT 不存在', await putOpp('_s_u1', 'o2', { contactId: 'ghost' }));
  rec('PUT o1 Won', await putOpp('_s_u1', 'o1', { stage: 'Won' }));
  rec('PUT o2 Won（主管）', await putOpp('_s_m1', 'o2', { stage: 'Won' }));
  rec('PUT o3 Won（無名片）', await putOpp('_s_u1', 'o3', { stage: 'Won' }));
  rec('PUT o1 回 B', await putOpp('_s_u1', 'o1', { stage: 'B' }));
  rec('Lead 轉換 Won', await call('_s_adm', 'POST', '/api/leads/l1/convert', { salesPerson: '_s_u1', stage: 'Won', product: '產品', category: 'ERP' }));
  for (const w of ['_s_adm', '_s_u1', '_s_u2']) rec('GET contacts ' + w, await call(w, 'GET', '/api/contacts'));
  rec('POST contactId 空白＋公司不一致', await postOpp('_s_u1', { contactId: ' c1 ', company: '完全不同公司' }));
  rec('POST 陣列 contactId', await postOpp('_s_u1', { contactId: ['c1'] }));
  rec('PUT 陣列 contactId', await putOpp('_s_u1', 'o5', { contactId: ['c4'] }));
  { const m = await call('_s_u1', 'GET', '/api/me'); if (m.json) { delete m.json.oppRequireContact; delete m.json.loginAt; } rec('GET me（旗標欄位除外：關閉時才多帶，是刻意的前端通知）', m); }
  rec('GET opportunities', await call('_s_adm', 'GET', '/api/opportunities'));
  const aoa = [['客戶名稱', '銷售案名', 'BU', '預定簽約日', '業務帳號', '把握度', '合約金額'], ['甲公司', '案A', 'ERP', '2027-06-01', '_s_u1', 'B', 100], ['沒有名片的公司', '案B', 'ERP', '2027-06-01', '_s_u1', 'B', 100], ['甲乙丙有限公司', '案C', 'ERP', '', '_s_u2', 'Won', 100]];
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'S');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  for (const dry of [true, false]) {
    const fd = new FormData(); fd.append('file', new Blob([buf]), 'o.xlsx'); if (dry) fd.append('dryRun', '1');
    const res = await fetch(BASEURL + '/api/admin/opportunities/import', { method: 'POST', headers: { Cookie: jar._s_adm }, body: fd });
    let j = null; try { j = await res.json(); } catch { /* 無 */ }
    out.push(['匯入 dry=' + dry, res.status, norm(j)]);
  }
  rec('GET opportunities 匯入後', await call('_s_adm', 'GET', '/api/opportunities'));
  return out;
}
async function offGroup(root) {
  console.log('  [off] 開關=0 與基準逐位元相同');
  if (!BASE_REF) { console.log('    （WON_BASE_REF 為空：略過，標示為未驗證）'); return; }
  let base;
  try { base = mkBase(BASE_REF); } catch (e) { ok(false, 'O0 取得基準版本失敗 ' + e.message); return; }
  const a = await scenario(base, {});
  const b = await scenario(root, { WON_AUTO_CUSTOMER: '0', OPP_REQUIRE_CONTACT: '0' });
  eq(b.length, a.length, 'O1 情境數相同');
  a.forEach((x, i) => { const y = b[i] || []; ok(x[0] === y[0] && x[1] === y[1] && x[2] === y[2], 'O2 ' + x[0] + ' 與基準相同', x[1] + ' vs ' + y[1] + (x[2] !== y[2] ? ' body 不同 ' + x[2].slice(0, 300) + ' ≠ ' + String(y[2]).slice(0, 300) : '')); });
  // 只關一個開關時，另一個仍然運作
  await freshServer(root, { WON_AUTO_CUSTOMER: '0' });
  await login('_s_u1'); await login('_s_adm');
  eq((await postOpp('_s_u1', {})).status, 400, 'O3 只關成交轉客戶：建立商機仍強制綁名片');
  await putOpp('_s_u1', 'o1', { stage: 'Won' });
  eq(await ctype('_s_u1', 'c1'), '', 'O3b WON_AUTO_CUSTOMER=0：成交不轉客戶');
  eq((await call('_s_adm', 'POST', '/api/admin/won-resync-customers', {})).status, 404, 'O3c WON_AUTO_CUSTOMER=0：重新同步端點 404');
  await freshServer(root, { OPP_REQUIRE_CONTACT: '0' });
  await login('_s_u1');
  eq(((await call('_s_u1', 'GET', '/api/me')).json || {}).oppRequireContact, false, 'O4a OPP_REQUIRE_CONTACT=0：/api/me 帶 oppRequireContact:false（前端據此不擋）');
  eq((await postOpp('_s_u1', {})).status, 201, 'O4 OPP_REQUIRE_CONTACT=0：不綁名片可建立');
  await putOpp('_s_u1', 'o1', { stage: 'Won' });
  eq(await ctype('_s_u1', 'c1'), 'customer', 'O4b 只關綁名片：成交轉客戶仍運作');
}

// ───────────────── mut：變異測試 ─────────────────
const MUTS = [
  ['M1 拿掉成交轉客戶（不寫 customerType）', 'lib/wonAutoCustomer.js', "  c.customerType = 'customer';\n  c.updatedAt", "  c.updatedAt"],
  ['M2 拿掉冪等（已是客戶也重寫並記稽核）', 'lib/wonAutoCustomer.js', "  if (c.customerType === 'customer') return { status: 'already', contactId: c.id, changed: false };\n", ''],
  ['M3 拿掉重新同步端點的管理員權限', 'server.js', "app.post('/api/admin/won-resync-customers', requireAdmin,", "app.post('/api/admin/won-resync-customers', requireAuth,"],
  ['M3b dryRun 也寫入', 'lib/wonAutoCustomer.js', "if (apply) { f.contact.customerType = 'customer'; f.contact.updatedAt", "{ f.contact.customerType = 'customer'; f.contact.updatedAt", ['server.js', 'if (apply && rep.converted > 0) db.save(data);', 'if (rep.converted > 0) db.save(data);']],
  ['M4 拿掉 contactId 必填', 'server.js', "  if (!cid) return { status: 400, code: 'OPP_CONTACT_REQUIRED'", "  if (false) return { status: 400, code: 'OPP_CONTACT_REQUIRED'"],
  ['M5 拿掉名片存在檢查', 'server.js', "  if (!c) return { status: 400, code: 'OPP_CONTACT_NOT_FOUND'", "  if (false) return { status: 400, code: 'OPP_CONTACT_NOT_FOUND'"],
  ['M6 拿掉名片權限（可視）檢查', 'server.js', "  if (!visibleContactsOf(req, data, loadAuth(), [c]).length) return", "  if (false) return"],
  ['M7 拿掉 PUT 的聯絡人檢查', 'server.js', "    if (_cidChanged) {", "    if (false) {"],
  ['M8 PUT 把「與現值相同」也檢查（舊資料會壞）', 'server.js', "    const _cidChanged = _newCid !== _curCid;", "    const _cidChanged = true;"],
  ['M9 Excel 匯入不拒絕無名片列', 'server.js', "if (!contactId && requireContactImport) {", "if (false) {"],
  ['M11 拿掉公司一致性檢查', 'server.js', "  if (oppContactCompanyMismatch(c, company)) return OPP_CONTACT_MISMATCH;", ""],
  ['M12 contactId 不 trim', 'server.js', "const cid = typeof contactId === 'string' ? contactId.trim() : '';", "const cid = typeof contactId === 'string' ? contactId : '';"],
  ['M12b PUT 不擋非字串 contactId', 'server.js', "if (_rawC !== undefined && _rawC !== null && typeof _rawC !== 'string') return", "if (false) return"],
  ['M13 匯入成交列不轉客戶', 'server.js', "if (wonOn && opp.stage === 'Won') {", "if (false) {"],
  ['M14 拿掉「只改公司」的一致性分支', 'server.js', "} else if (_companyChanged && _newCid) {", "} else if (false) {"],
  ['M15 允許把已綁名片商機的公司清空', 'server.js', "if (_companyChanged && !_effCompany && _newCid) return", "if (false) return"],
  ['M16 公司比對不做 NFKC', 'server.js', ".normalize('NFKC')", ""],
  ['M17 成交後改綁不轉新名片', 'server.js', "(oldStage !== 'Won' || (_newCidSaved && _newCidSaved !== _oldCid))", "(oldStage !== 'Won')"],
  ['M18 稽核 target 退用人名', 'server.js', "const tgt = opp.company || '（無公司名）';", "const tgt = opp.company || opp.contactName || '';"],
  ['M19 綁定時不補公司', 'server.js', "if (!_effCompany && _rc.contact && _rc.contact.company) req.body.company = _rc.contact.company;", ""],
  ['M20 不做臺→台', 'server.js', ".replace(/臺/g, '台')", ""],
  ['M10 成交時不轉（Lead 與 PUT 都關掉）', 'server.js', "const _won = (WONC.enabled() && newStage === 'Won'", "const _won = (false && newStage === 'Won'"],
];
async function mutGroup() {
  console.log('  [mut] 變異測試（每個變異 api 群組都必須變紅）');
  for (const [name, file, find, rep, extra] of MUTS) {
    const dir = mkSandbox();
    const p = path.join(dir, file);
    const src = fs.readFileSync(p, 'utf8');
    if (!src.includes(find)) { ok(false, name + '：錨點不存在（測試腐爛）'); continue; }
    fs.writeFileSync(p, src.replace(find, () => rep));
    if (extra) { const p2 = path.join(dir, extra[0]); const s2 = fs.readFileSync(p2, 'utf8'); if (!s2.includes(extra[1])) { ok(false, name + '：第二個錨點不存在'); continue; } fs.writeFileSync(p2, s2.replace(extra[1], () => extra[2])); }
    const f0 = fail, p0 = pass, fl0 = fails.length;
    const quiet = console.log; console.log = () => {};
    try { await apiGroup(dir, name); } catch (e) { fail++; } finally { console.log = quiet; }
    const red = fail > f0;
    const failedHere = fail - f0;
    fail = f0; pass = p0; fails.length = fl0;
    ok(red, name + ' → 必須變紅', '（卻全綠）');
    if (red) console.log(`    ✓ ${name}：變紅（${failedHere} 項失敗）`);
  }
}

(async () => {
  let RUN_ROOT = '';
  try {
    if (want('unit')) unitGroup();
    if (want('api') || want('off') || want('mut')) RUN_ROOT = mkSandbox();
    if (want('api')) await apiGroup(RUN_ROOT);
    if (want('off')) await offGroup(RUN_ROOT);
    if (want('mut')) await mutGroup();
  } catch (e) { fail++; fails.push('例外 ' + e.message); console.log('  ✗ 例外', e.stack || e.message); }
  finally { await stopServer(); cleanAll(); }
  console.log(`\n結果：${fail ? '失敗' : '通過'}（pass ${pass}／fail ${fail}）`);
  if (fails.length) console.log('失敗項目：' + fails.slice(0, 40).join('；'));
  process.exit(fail ? 1 : 0);
})();
