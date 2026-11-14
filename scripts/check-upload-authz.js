#!/usr/bin/env node
'use strict';
/**
 * SEC-02：GET /uploads/:key（名片照片下載）必須核對「引用這個檔案的聯絡人」是否在目前使用者的可視範圍內 — 端對端檢查。
 * 不含任何帳密：測試帳號與隨機密碼只在記憶體裡；不需要瀏覽器；用本地 JSON 後端。
 *
 * 背景：原本 /uploads/:key 只要求登入。檔名雖是 UUID，但任何登入者只要拿到網址就能下載別人的名片照片（個資）。
 *       修正後：找出引用該 key 的聯絡人（contacts[].cardImage），用與 GET /api/contacts 相同的可視範圍（visibleContactsOf）判斷；
 *       看不到 → 404（不用 403，避免洩漏檔案存在）；找不到任何引用 → 404；未登入 → 401；管理員可看全部「有被引用」的名片圖。
 *
 * 做法：在 SEC02_ROOT（預設＝本工作樹）建立臨時 auth.json / data.json（假資料 + _s_* 臨時帳號）與 uploads/ 假圖檔，
 *       以 PORT=SEC02_PORT（預設 3710，限 3710-3719）啟動 server.js，跑完只用 PID 關閉並刪除自己建立的檔案。
 * 守門：SEC02_ROOT 若已有 auth.json / data.json / audit.log.json（可能是真實資料），一律拒絕執行。
 *
 *   node scripts/check-upload-authz.js                          # 斷言（修正後應全綠；在修正前的程式上應有多項變紅）
 *   SEC02_PORT=3711 node scripts/check-upload-authz.js          # 換埠
 *   SEC02_ROOT=<其他程式樹> node scripts/check-upload-authz.js   # 對別的程式樹跑（例如修正前的基準）
 *   SEC02_DEBUG=1 node scripts/...                              # 顯示伺服器輸出
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.resolve(process.env.SEC02_ROOT || path.join(__dirname, '..'));
const PORT = parseInt(process.env.SEC02_PORT || '3710', 10);
if (!(PORT >= 3710 && PORT <= 3719)) { console.log('拒絕：埠須在 3710-3719'); process.exit(2); }
const UPLOAD_DIR = path.join(ROOT, 'uploads');
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));
const guardFiles = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
if (guardFiles.length) { console.log('拒絕：', ROOT, '已有', guardFiles.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const bcrypt = require(path.join(__dirname, '..', 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));

const PW = 'Up' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const T0 = '2026-01-01T00:00:00.000Z';
const KA_CO = 'KA沙盒公司', ERP_CO = 'ERP沙盒公司', ITS_CO = 'ITS沙盒公司';
// 1x1 透明 PNG（假圖檔，內容固定；下載成功時用來確認拿到的是這個檔案）
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==', 'base64');

const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: T0, passwordChangedAt: new Date().toISOString(), mustChangePassword: false }, extra || {});
const U = {
  admin: '_s_adm', executive: '_s_exec', manager1: '_s_m1', manager2: '_s_m2',
  A: '_s_a_erp', B: '_s_b_erp', ITS: '_s_u_its', marketing: '_s_mkt', secretary: '_s_sec', tecopm: '_s_tpm', groupsales: '_s_gs',
};
// 每張假名片圖的檔名（uuid.副檔名；與 POST /api/upload 產生的格式相同）
const uuid = () => crypto.randomUUID();
const K = { A: uuid() + '.png', B: uuid() + '.png', ITS: uuid() + '.png', KA: uuid() + '.png', DEL: uuid() + '.png', POOL: uuid() + '.png', ORPHAN: uuid() + '.png', NEW: uuid() + '.png', XB: uuid() + '.png', OUT: uuid() + '.png', SEC: uuid() + '.png' };
const createdUploads = [];

function seedFiles() {
  const users = [
    mkUser(U.admin, 'admin'), mkUser(U.executive, 'executive'),
    mkUser(U.manager1, 'manager1', { bu: ['ERP'] }), mkUser(U.manager2, 'manager2', { bu: ['ITS'] }),
    mkUser(U.A, 'user', { bu: ['ERP'], supervisor: U.manager1 }),
    mkUser(U.B, 'user', { bu: ['ERP'], supervisor: U.manager1 }),     // 與 A 同部門、同主管、同層級
    mkUser(U.ITS, 'user', { bu: ['ITS'], supervisor: U.manager2 }),
    mkUser(U.marketing, 'marketing'),
    mkUser(U.secretary, 'secretary', { bu: ['ERP'] }),
    mkUser('_s_xb', 'user', { bu: ['ITS'], supervisor: U.manager1 }),   // 第 6 節：在 ERP 主管 manager1 的子樹內、但 BU 是 ITS
    mkUser(U.tecopm, 'tecopm', { viewOwnerScope: U.A, viewGroupId: '_s_g1' }),
    mkUser(U.groupsales, 'groupsales', { viewOwnerScope: U.A, viewGroupId: '_s_g1' }),
  ];
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  const ct = (id, owner, bu, company, name, key, extra) => Object.assign({ id, name, company, owner, bu, title: '經理', phone: '0912000000', email: id + '@sandbox.test', note: '沙盒', cardImage: key ? '/uploads/' + key : '', createdAt: T0 }, extra || {});
  const d = {
    contacts: [
      ct('c-a', U.A, 'ERP', ERP_CO, '甲（A 的名片）', K.A),
      ct('c-b', U.B, 'ERP', ERP_CO, '乙（B 的名片）', K.B),
      ct('c-its', U.ITS, 'ITS', ITS_CO, '丙（ITS 的名片）', K.ITS),
      ct('c-ka', U.ITS, 'ITS', KA_CO, '丁（ITS 的 KA 名片）', K.KA),
      ct('c-del', U.A, 'ERP', ERP_CO, '戊（A 已刪除的名片）', K.DEL, { deleted: true }),
      ct('c-pool', '_pool', 'ERP', ERP_CO, '己（客戶池的名片）', K.POOL),
      ct('c-xb', '_s_xb', 'ITS', ERP_CO, '辛（manager1 子樹內、BU＝ITS 的名片）', K.XB),
      ct('c-out', U.A, 'ERP', '集團外沙盒公司', '壬（A 的名片，但公司不在集團 _s_g1）', K.OUT),
      ct('c-sec', U.secretary, 'ERP', ERP_CO, '癸（秘書本人名下、與秘書同 BU 的名片）', K.SEC),
      ct('c-new', U.A, 'ERP', ERP_CO, '庚（A 的名片，第 5 節播種；原為真實上傳後存成名片）', K.NEW),
    ],
    companies: [], keyAccounts: [{ id: 'ka1', company: KA_CO }],
    visits: [], opportunities: [], contracts: [], receivables: [], callins: [], quotations: [],
    groups: [{ id: '_s_g1', name: '沙盒集團', memberCompanies: [ERP_CO, KA_CO] }],
  };
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
  if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
  for (const k of Object.values(K)) { fs.writeFileSync(path.join(UPLOAD_DIR, k), PNG); createdUploads.push(k); }
}
function cleanupFiles() {
  FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } });
  createdUploads.forEach(k => { try { fs.unlinkSync(path.join(UPLOAD_DIR, k)); } catch { /* 不存在 */ } });
}

let child = null, childPid = null;
async function startServer() {
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', STORAGE_BACKEND: 'local', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  delete env.DATABASE_URL; delete env.MAIL_MODE; delete env.SUPABASE_URL; delete env.SUPABASE_SERVICE_ROLE_KEY;
  await assertPortFree(PORT);
  child = spawn(process.execPath, [path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  const onData = d => { out += d; if (process.env.SEC02_DEBUG) process.stdout.write('[srv] ' + d); };
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
async function http(who, method, url, body) {
  const h = {};
  if (who && jar[who]) h.Cookie = jar[who];
  let payload;
  if (body instanceof FormData) payload = body;
  else if (body !== undefined) { h['Content-Type'] = 'application/json'; payload = JSON.stringify(body); }
  const res = await fetch(BASE + url, { method, headers: h, body: payload, redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: res.status, json, buf, type: res.headers.get('content-type') || '' };
}
async function login(who) {
  const res = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status !== 200) throw new Error('登入失敗 ' + who + ' ' + res.status);
  jar[who] = res.headers.getSetCookie().map(c => c.split(';')[0]).join('; ');
}
// 下載某個 key：200＝拿到的就是那張假圖；其他回 status
async function dl(who, key) { const r = await http(who, 'GET', '/uploads/' + key); return r; }
const got = async (who, key) => (await dl(who, key)).status;
const isPng = (r) => r.status === 200 && r.buf.equals(PNG);

async function run() {
  for (const u of Object.values(U)) await login(u);

  console.log('[1] A 業務：自己的名片圖可下載（內容正確）；同層級 B、他部門 ITS 的不行');
  {
    const r = await dl(U.A, K.A);
    ok(isPng(r), `1.1 A 下載自己的名片圖＝200 且內容正確（得到 ${r.status}）`);
    eq(await got(U.A, K.B), 404, '1.2 A 下載「同層級同部門 B」的名片圖＝404（不是 403）');
    eq(await got(U.A, K.ITS), 404, '1.3 A 下載 ITS 業務（非 KA 公司）的名片圖＝404');
    eq(await got(U.B, K.B) === 200 && await got(U.B, K.A), 404, '1.4 B 下載自己的＝200，下載 A 的＝404（對稱）');
    eq(await got(U.A, K.KA), 200, '1.5 A 下載 ITS 業務的「KA 公司」名片圖＝200（KA 跨部門共享是既有設計，清單看得到就下載得到）');
    eq(await got(U.A, K.DEL), 404, '1.6 已軟刪除的名片圖，連擁有者本人也 404（清單也看不到）');
  }

  console.log('[2] 主管／管理員／其他角色依原權限');
  {
    eq([await got(U.manager1, K.A), await got(U.manager1, K.B), await got(U.manager1, K.ITS)], [200, 200, 404], '2.1 ERP 主管：看得到部屬 A、B 的名片圖，看不到 ITS 的');
    eq([await got(U.manager2, K.ITS), await got(U.manager2, K.A)], [200, 404], '2.2 ITS 主管：看得到自己部屬的，看不到 ERP 的');
    eq([K.A, K.B, K.ITS, K.KA, K.DEL, K.POOL].map(() => 0).length, 6, '2.0（自檢）');
    const adm = [];
    for (const k of [K.A, K.B, K.ITS, K.KA, K.DEL, K.POOL]) adm.push(await got(U.admin, k));
    eq(adm, [200, 200, 200, 200, 200, 200], '2.3 管理員：所有「有被引用」的名片圖都可看（含已軟刪除、客戶池）');
    eq([await got(U.executive, K.A), await got(U.executive, K.ITS), await got(U.executive, K.DEL)], [200, 200, 404], '2.4 董事長／總經理：跨 BU 看全部，但已刪除的不行（與清單一致）');
    eq([await got(U.marketing, K.A), await got(U.marketing, K.ITS), await got(U.marketing, K.KA)], [200, 200, 200], '2.5 行銷：可檢視全部客戶（既有設計）');
    const sec = [];
    for (const k of [K.A, K.B, K.ITS, K.KA]) sec.push(await got(U.secretary, k));
    eq(sec, [404, 404, 404, 404], '2.6 秘書：刻意看不到聯絡人（含 KA 共享），名片圖一律 404');
    eq([await got(U.tecopm, K.A), await got(U.tecopm, K.B), await got(U.tecopm, K.KA)], [200, 404, 404], '2.7 集團 PM：只看 viewOwnerScope（A）且限集團公司；B 的不行；KA 共享不適用');
  }

  console.log('[3] 未登入、不存在、格式不對、沒人引用');
  {
    const r = await http(null, 'GET', '/uploads/' + K.A);
    eq(r.status, 401, '3.1 未登入下載＝401（不是導向登入頁）');
    eq(await got(U.A, uuid() + '.png'), 404, '3.2 不存在的 key＝404');
    eq(await got(U.admin, K.ORPHAN), 404, '3.3 檔案存在但沒有任何名片引用＝404（連管理員也是）');
    eq(await got(U.A, '..%2Fauth.json'), 404, '3.4 路徑穿越（..%2Fauth.json）＝404');
    eq(await got(U.A, 'auth.json'), 404, '3.5 非 uuid 格式的檔名＝404');
    eq(await got(U.A, K.A.replace('.png', '.PNG')), 404, '3.6 副檔名大小寫不同（沒有任何紀錄引用這個字串）＝404');
  }

  console.log('[4] 一致性：每個角色「清單看得到的名片圖」＝「下載得到的名片圖」（與 GET /api/contacts 同一套規則）');
  {
    const withImg = ['c-a', 'c-b', 'c-its', 'c-ka'];   // 未刪除、非客戶池
    const keyOf = { 'c-a': K.A, 'c-b': K.B, 'c-its': K.ITS, 'c-ka': K.KA };
    for (const [role, who] of Object.entries(U)) {
      if (role === 'admin') continue;   // 管理員另有「看全部」規則（見 2.3）
      const list = await http(who, 'GET', '/api/contacts');
      const listed = new Set((Array.isArray(list.json) ? list.json : []).map(c => c.id));
      for (const id of withImg) {
        const status = await got(who, keyOf[id]);
        ok((status === 200) === listed.has(id), `4.x ${role} 對 ${id}：清單${listed.has(id) ? '有' : '無'}，下載 ${status}（兩者必須一致）`);
      }
    }
  }

  // 名片圖不再儲存於伺服器：POST /api/upload 已停用（410），所以原本「真實上傳→存成名片」改成播種——
  // 直接在儲存目錄放檔＋資料裡的 c-new.cardImage 指向它（seedFiles）。授權斷言（5.5–5.9）維持不變，只是來源改成播種。
  // 上傳入口停用本身另見 scripts/check-no-card-storage.js。
  console.log('[5] 播種的名片圖（取代原「真實上傳」流程）：A 的名片圖 → 依角色可下載');
  {
    const url = '/uploads/' + K.NEW;
    const up = await http(U.A, 'POST', '/api/upload', (() => { const f = new FormData(); f.append('card', new Blob([PNG], { type: 'image/png' }), 'card.png'); return f; })());
    eq(up.status, 410, '5.1 A 上傳名片圖＝410（上傳入口已停用）');
    eq(await got(U.A, K.ORPHAN), 404, '5.3 檔案在儲存目錄但沒有任何名片引用：連 A 自己也是 404');
    ok(isPng(await http(U.A, 'GET', url)), '5.5 A 可下載自己的名片圖（內容正確）');
    eq(await got(U.B, K.NEW), 404, '5.6 同層級 B 下載＝404');
    eq(await got(U.manager1, K.NEW), 200, '5.7 A 的主管下載＝200');
    eq(await got(U.admin, K.NEW), 200, '5.8 管理員下載＝200');
    eq(await got(U.manager2, K.NEW), 404, '5.9 別的部門主管下載＝404');
  }
  console.log('[6] 補 filterByBu／filterByViewGroup／秘書分支的覆蓋（這三個函式與 GET /api/contacts 共用）');
  {
    await login('_s_xb');
    const listOf = async (who) => new Set(((await http(who, 'GET', '/api/contacts')).json || []).map(c => c.id));
    // ① 主管子樹內但 BU 不符：manager1 的清單看不到、也不得下載；admin／ITS 主管（manager2 不在其子樹）也要有明確預期
    const l1 = await listOf(U.manager1);
    ok(!l1.has('c-xb'), '6.1 ERP 主管 manager1：GET /api/contacts 看不到「子樹內但 BU＝ITS」的名片');
    eq(await got(U.manager1, K.XB), 404, '6.2 ERP 主管 manager1 不得下載該名片圖（BU 不符）');
    eq(await got('_s_xb', K.XB), 200, '6.3 名片擁有者本人（_s_xb）可下載自己的圖');
    eq(await got(U.admin, K.XB), 200, '6.4 管理員可下載');
    eq(await got(U.manager2, K.XB), 404, '6.5 ITS 主管 manager2（該員不在其子樹）不得下載');
    // ② tecopm／groupsales：範圍（A）內，但公司不在其集團
    for (const [role, who] of [['tecopm', U.tecopm], ['groupsales', U.groupsales]]) {
      const l = await listOf(who);
      ok(l.has('c-a') && !l.has('c-out'), '6.6 ' + role + '：清單有集團內公司的 c-a、沒有集團外公司的 c-out');
      eq(await got(who, K.OUT), 404, '6.7 ' + role + ' 不得下載「公司不在其集團」的名片圖（即使 owner 在範圍內）');
      eq(await got(who, K.A), 200, '6.8 ' + role + ' 仍可下載集團內公司的名片圖（對照組）');
    }
    eq(await got(U.A, K.OUT), 200, '6.9 擁有者 A 可下載自己的（集團外公司）名片圖');
    eq(await got(U.manager1, K.OUT), 200, '6.10 A 的主管可下載（一般角色不受集團篩選）');
    // ③ 秘書：任何圖一律 404，含「名片 owner 就是秘書本人、與秘書同 BU」
    const lsec = await listOf(U.secretary);
    ok(!lsec.has('c-sec') && !lsec.has('c-a') && !lsec.has('c-xb'), '6.11 秘書：GET /api/contacts 看不到任何名片（含自己名下）');
    const secs = [];
    for (const k of [K.SEC, K.A, K.B, K.XB, K.OUT, K.NEW, K.KA]) secs.push(await got(U.secretary, k));
    eq(secs, [404, 404, 404, 404, 404, 404, 404], '6.12 秘書下載任何名片圖都 404（含 owner＝秘書本人、同 BU）');
    eq(await got(U.admin, K.SEC), 200, '6.13 管理員可下載秘書名下名片的圖（對照組，證明 404 是授權造成而非檔案不存在）');
    // 一致性（新增名片）：各角色「清單有＝下載得到」
    const items = [['c-xb', K.XB], ['c-out', K.OUT], ['c-sec', K.SEC]];
    for (const [role, who] of Object.entries(U)) {
      if (role === 'admin') continue;
      const l = await listOf(who);
      for (const [id, key] of items) ok((await got(who, key) === 200) === l.has(id), '6.14 ' + role + ' 對 ' + id + '：清單與下載一致');
    }
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
