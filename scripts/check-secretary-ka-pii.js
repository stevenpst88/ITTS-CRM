#!/usr/bin/env node
'use strict';
/**
 * 秘書不可經「KA 跨部門共享」取得聯絡人／拜訪（個資）— 端對端檢查。不含任何帳密：測試帳號與隨機密碼只在記憶體裡。
 *
 * 背景：秘書刻意看不到聯絡人與拜訪（GET /api/contacts、/api/visits 的基底清單為空），
 *       但 mergeKaVisible 曾把「Key Account 公司」的聯絡人／拜訪以完整物件併回清單（phone/mobile/email/note/cardImage…），
 *       連 ?search= 也能探測。本檢查用 SENTINEL 字串確認秘書的回應完全不含這些資料。
 *
 * 做法：在 KA_ROOT（預設＝本工作樹）建立臨時 auth.json / data.json（假資料 + _s_* 臨時帳號），
 *       以 PORT=KA_PORT（預設 3480，限 3480-3489）啟動 server.js，跑完只用 PID 關閉並刪除臨時檔。
 * 守門：KA_ROOT 若已有 auth.json / data.json / audit.log.json（可能是真實資料），一律拒絕執行。
 *
 *   node scripts/check-secretary-ka-pii.js                      # 斷言（修正後應全綠；在修正前的程式上應變紅）
 *   KA_ROOT=<其他程式樹> node scripts/check-secretary-ka-pii.js  # 對別的程式樹跑（例如 origin/main 的 git archive 基準）
 *   KA_DUMP_OUT=<檔> node scripts/...                            # 另外把各角色三支 API 的原始回應存檔
 *   KA_COMPARE=<檔> node scripts/...                             # 與先前存的基準比對：非秘書角色須逐位元相同
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.resolve(process.env.KA_ROOT || path.join(__dirname, '..'));
const PORT = parseInt(process.env.KA_PORT || '3480', 10);
if (!(PORT >= 3480 && PORT <= 3489)) { console.log('拒絕：埠須在 3480-3489'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));
const guardFiles = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
if (guardFiles.length) { console.log('拒絕：', ROOT, '已有', guardFiles.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const bcrypt = require(path.join(__dirname, '..', 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));

const PW = 'Ka' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const T0 = '2026-01-01T00:00:00.000Z';                       // 固定時間戳：讓兩次執行的輸出可逐位元比對
const KA_CO = 'KA沙盒公司', PLAIN_CO = '一般沙盒公司';
// SENTINEL：若出現在秘書的回應裡就代表個資外洩
const S = { phone: 'SENT_PHONE_7731', mobile: 'SENT_MOBILE_8842', email: 'SENT_EMAIL_9953@sandbox.test', note: 'SENT_NOTE_AAA', img: 'SENT_CARDIMG_BBB', topic: 'SENT_TOPIC_CCC', visitNote: 'SENT_VISITNOTE_DDD' };
const SENT_RE = /SENT_[A-Z]+_/;

const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: T0, passwordChangedAt: new Date().toISOString(), mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
const ROLES = {   // 每個角色的臨時帳號（秘書另有 SEC_NONE：沒掛 BU）
  admin: '_s_adm', executive: '_s_exec', accounting_manager: '_s_acct', finance_manager: '_s_fin',
  manager1: '_s_m1', manager2: '_s_m2', user_erp: '_s_u_erp', user_its: '_s_u_its', user_its2: '_s_u_its2',
  marketing: '_s_mkt', tecopm: '_s_tpm', groupsales: '_s_gs', consult_manager_south: '_s_csouth', consult_manager_north: '_s_cnorth',
  secretary: '_s_sec', secretary_none: '_s_sec_none',
};
function seedFiles() {
  const users = [
    mkUser(ROLES.admin, 'admin'), mkUser(ROLES.executive, 'executive'),
    mkUser(ROLES.accounting_manager, 'accounting_manager'), mkUser(ROLES.finance_manager, 'finance_manager'),
    mkUser(ROLES.manager1, 'manager1', { bu: ['ERP'] }), mkUser(ROLES.manager2, 'manager2', { bu: ['ITS'] }),
    mkUser(ROLES.user_erp, 'user', { bu: ['ERP'], supervisor: ROLES.manager1 }),
    mkUser(ROLES.user_its, 'user', { bu: ['ITS'], supervisor: ROLES.manager2 }),
    mkUser(ROLES.user_its2, 'user', { bu: ['ITS'], supervisor: ROLES.manager2 }),
    mkUser(ROLES.marketing, 'marketing'),
    mkUser(ROLES.tecopm, 'tecopm', { viewOwnerScope: ROLES.user_its, viewGroupId: '_s_g1' }),
    mkUser(ROLES.groupsales, 'groupsales', { viewOwnerScope: ROLES.user_its, viewGroupId: '_s_g1' }),
    mkUser(ROLES.consult_manager_south, 'consult_manager_south'), mkUser(ROLES.consult_manager_north, 'consult_manager_north'),
    mkUser(ROLES.secretary, 'secretary', { bu: ['ERP'] }), mkUser(ROLES.secretary_none, 'secretary', { bu: [] }),
  ];
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  const ct = (id, owner, bu, company, name) => ({ id, name, company, owner, bu, title: '經理', phone: S.phone, mobile: S.mobile, email: S.email, note: S.note, cardImage: S.img, createdAt: T0 });
  const vt = (id, owner, bu, contactId, company) => ({ id, owner, bu, contactId, company, visitDate: '2026-01-05', topic: S.topic, notes: S.visitNote, createdAt: T0 });
  const op = (id, owner, bu, category, company) => ({ id, owner, bu, category, company, contactName: '沙盒聯絡人', product: '沙盒商品', amount: '100', stage: 'B', expectedDate: '2027-06-01', description: '沙盒商機', createdAt: T0 });
  const d = {
    contacts: [
      ct('c-ka-its', ROLES.user_its, 'ITS', KA_CO, 'KA聯絡甲'),
      ct('c-ka-erp', ROLES.user_erp, 'ERP', KA_CO, 'KA聯絡乙'),
      ct('c-plain-its', ROLES.user_its, 'ITS', PLAIN_CO, '一般聯絡丙'),
    ],
    companies: [],
    keyAccounts: [{ id: 'ka1', company: KA_CO }],
    visits: [
      vt('v-ka-its', ROLES.user_its, 'ITS', 'c-ka-its', KA_CO),
      vt('v-ka-erp', ROLES.user_erp, 'ERP', 'c-ka-erp', KA_CO),
      vt('v-plain-its', ROLES.user_its, 'ITS', 'c-plain-its', PLAIN_CO),
    ],
    opportunities: [
      op('o-ka-its', ROLES.user_its, 'ITS', 'ERP', KA_CO),     // ITS 業務建、類別 ERP → 跨部門，ERP 看得到
      op('o-ka-erp', ROLES.user_erp, 'ERP', 'ITS', KA_CO),     // ERP 業務建、類別 ITS → 跨部門，ITS 看得到
      op('o-plain-its', ROLES.user_its, 'ITS', 'ITS', PLAIN_CO),
    ],
    contracts: [], receivables: [], callins: [], quotations: [],
    groups: [{ id: '_s_g1', name: '沙盒集團', memberCompanies: [KA_CO, PLAIN_CO] }],
  };
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
}
function cleanupFiles() { FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }

let child = null, childPid = null;
async function startServer() {
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  delete env.DATABASE_URL; delete env.MAIL_MODE;
  await assertPortFree(PORT);
  child = spawn(process.execPath, [path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  const onData = d => { out += d; if (process.env.KA_DEBUG) process.stdout.write('[srv] ' + d); };
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

const jar = {};
async function http(who, url) {
  const h = {};
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { headers: h, redirect: 'manual' });
  const text = await res.text();
  let json = null; try { json = JSON.parse(text); } catch { /* 非 JSON */ }
  return { status: res.status, json, text };
}
async function login(who) {
  const res = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status !== 200) throw new Error('登入失敗 ' + who + ' ' + res.status);
  jar[who] = res.headers.getSetCookie().map(c => c.split(';')[0]).join('; ');
}
const idsOf = (r) => (Array.isArray(r.json) ? r.json.map(x => x.id).sort() : null);

async function run() {
  for (const u of Object.values(ROLES)) await login(u);
  const EP = { contacts: '/api/contacts', visits: '/api/visits', opps: '/api/opportunities' };

  console.log('[1] 秘書（ERP 與無 BU 兩種）：聯絡人／拜訪不得含任何個資');
  for (const who of [ROLES.secretary, ROLES.secretary_none]) {
    for (const [k, url] of [['contacts', EP.contacts], ['visits', EP.visits]]) {
      const r = await http(who, url);
      eq(r.status, 200, `1.1 ${who} GET ${url} 200`);
      ok(!SENT_RE.test(r.text), `1.2 ${who} GET ${url} 回應不含任何 SENTINEL 個資`);
      eq(r.json, [], `1.3 ${who} GET ${url} 為空陣列（秘書刻意看不到聯絡人／拜訪）`);
    }
    for (const q of [S.phone, S.mobile, S.email, S.note, 'KA聯絡甲', KA_CO, 'KA沙盒', '@sandbox.test']) {
      const r = await http(who, '/api/contacts?search=' + encodeURIComponent(q));
      ok(!SENT_RE.test(r.text) && Array.isArray(r.json) && r.json.length === 0, `1.4 ${who} ?search=${q} 探不到任何聯絡人`);
    }
  }

  console.log('[2] 秘書商機維持既有行為（可見 KA 跨部門商機；商機無電話／email／備註欄）');
  {
    const r = await http(ROLES.secretary, EP.opps);
    eq(r.status, 200, '2.1 秘書 GET /api/opportunities 200');
    ok(!SENT_RE.test(r.text), '2.2 秘書商機回應不含 SENTINEL');
    const ids = idsOf(r) || [];
    ok(ids.includes('o-ka-its') && ids.includes('o-ka-erp'), `2.3 秘書仍看得到 KA 跨 BU 商機（既有設計；得到 ${JSON.stringify(ids)}）`);
  }

  console.log('[3] 不是空轉：KA 共享對業務仍有效，且其他角色的可視範圍符合預期');
  {
    let r = await http(ROLES.user_erp, EP.contacts);
    ok((idsOf(r) || []).includes('c-ka-its'), '3.1 ERP 業務仍看得到 ITS 業務的 KA 聯絡人（KA 跨部門共享功能保留）');
    ok(!(idsOf(r) || []).includes('c-plain-its'), '3.2 ERP 業務看不到 ITS 的非 KA 聯絡人');
    r = await http(ROLES.user_erp, EP.visits);
    ok((idsOf(r) || []).includes('v-ka-its'), '3.3 ERP 業務仍看得到 ITS 業務的 KA 拜訪');
    r = await http(ROLES.user_its, EP.contacts);
    ok((idsOf(r) || []).includes('c-ka-erp'), '3.4 ITS 業務仍看得到 ERP 業務的 KA 聯絡人');
    r = await http(ROLES.admin, EP.contacts);
    eq(idsOf(r), ['c-ka-erp', 'c-ka-its', 'c-plain-its'], '3.5 管理員看全部聯絡人');
    r = await http(ROLES.executive, EP.contacts);
    eq(idsOf(r), ['c-ka-erp', 'c-ka-its', 'c-plain-its'], '3.6 董事長／總經理看全部聯絡人');
  }

  // 逐位元比對用：蒐集所有角色三支 API 的原始回應
  const dump = {};
  for (const [role, who] of Object.entries(ROLES)) {
    dump[role] = {};
    for (const [k, url] of Object.entries(EP)) { const r = await http(who, url); dump[role][k] = { status: r.status, body: r.text }; }
  }
  if (process.env.KA_DUMP_OUT) { fs.writeFileSync(process.env.KA_DUMP_OUT, JSON.stringify(dump)); console.log('已存檔：' + process.env.KA_DUMP_OUT); }
  if (process.env.KA_COMPARE) {
    console.log('[4] 與基準逐位元比對（非秘書角色全部相同；秘書商機相同）');
    const base = JSON.parse(fs.readFileSync(process.env.KA_COMPARE, 'utf8'));
    let same = 0;
    for (const role of Object.keys(ROLES)) {
      for (const k of Object.keys(EP)) {
        const isSec = role.startsWith('secretary');
        if (isSec && k !== 'opps') continue;   // 秘書的聯絡人／拜訪預期不同（本次修正）
        const a = base[role] && base[role][k], b = dump[role][k];
        const same1 = a && a.status === b.status && a.body === b.body;
        ok(same1, `4.x ${role} ${k} 與基準逐位元相同${same1 ? '' : '（不同）'}`);
        if (same1) same++;
      }
    }
    console.log('  逐位元相同組數：' + same);
  }
}

(async () => {
  let code = 0;
  try {
    cleanupFiles(); seedFiles();
    await startServer();
    await run();
  } catch (e) { console.log('錯誤：', e.message); fail++; failures.push('例外：' + e.message); code = 1; }
  finally { await stopServer(); cleanupFiles(); }
  console.log(`\n通過 ${pass}　失敗 ${fail}`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); code = 1; }
  process.exit(code);
})();
