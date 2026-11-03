#!/usr/bin/env node
'use strict';
/**
 * 名片「郵寄地址」(mailingAddress) ＋「同公司地址」(mailingSameAsCompany) 檢查：
 *   unit   靜態檢查（欄位白名單／標籤／表單／payload／SW 版本）
 *   api    真實 server.js：POST／PUT 語意、長度、舊值沿用、審計、通知／日誌不含地址、Excel 匯出匯入
 *   compat 舊名片（無新欄位）所有既有端點輸出與 CMAIL_BASE_REF（預設 origin/main）逐位元相同
 *   ui     headless Edge + CDP（見 _contact-mailing-ui.js）
 *
 *   node scripts/check-contact-mailing.js                  unit + api + compat
 *   CMAIL_ONLY=unit,api,compat,ui node ...                 只跑指定群組
 *   CMAIL_MUT=1 node ...                                   另外跑變異測試（每個變異都必須讓對應群組變紅）
 *   CMAIL_MUT_PICK=<regex>                                 只跑名稱符合的變異
 *   CMAIL_SHOTS=<資料夾>                                   ui 群組輸出截圖
 *   CMAIL_PORT（預設 3462，限 3460–3469）／CMAIL_CDP（預設 9762，限 9760–9769）／CMAIL_BASE_REF
 * 不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體；資料檔只存在於本次執行；全部假資料；結束清掉伺服器與 Edge。
 * 守門：工作樹已有 auth.json／data.json／audit.log.json 時拒絕執行。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.CMAIL_PORT || '3462', 10);
const CDP = parseInt(process.env.CMAIL_CDP || '9762', 10);
if (PORT < 3460 || PORT > 3469 || CDP < 9760 || CDP > 9769) { console.log('拒絕：埠需在 3460–3469 / CDP 9760–9769'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map((f) => path.join(ROOT, f));
{
  const g = FILES.filter((f) => /(auth|data|audit.log)\.jsonl?$/.test(f)).filter((f) => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map((f) => path.basename(f)).join('/'), '（可能是真實資料）'); process.exit(2); }
}
const ONLY = (process.env.CMAIL_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length ? k !== 'ui' : ONLY.includes(k);
const SHOTS = process.env.CMAIL_SHOTS || '';
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const XLSX = require(path.join(ROOT, 'node_modules', 'xlsx'));

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 300)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = (n) => new Promise((r) => setTimeout(r, n));

// ───────────────── 沙盒資料（全部假的）─────────────────
const PW = 'Cm' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const XSS = '<img src=x onerror="window.__xss=1">';
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
const mkCo = (id, name, tax) => ({ id, taxId: tax, matchKey: 'tax:' + tax, name, shortName: '', capital: null, industry: '', address: '', phone: '', website: '', region: '', customerCode: '', buImport: [], groupId: null, gcisEnriched: false, source: 'manual', createdAt: nowIso, updatedAt: nowIso });
const CO = { c1: mkCo('c1', '假公司甲', '90000001'), c2: mkCo('c2', '假公司乙', '90000002') };
const mkCard = (id, owner, co, extra) => Object.assign({ id, owner, name: '名片' + id, company: co.name, companyId: co.id, taxId: co.taxId, email: id + '@example.test', bu: 'ERP', address: '', createdAt: nowIso }, extra || {});
function seedFiles(dir) {
  dir = dir || ROOT;
  const users = [mkUser('_s_adm', 'admin'), mkUser('_s_mkt', 'marketing'), mkUser('_s_sec', 'secretary', { bu: ['ERP'] }),
    mkUser('_s_u1', 'user', { bu: ['ERP'], displayName: '測試業務甲' }), mkUser('_s_u2', 'user', { bu: ['ERP'], displayName: '測試業務乙' })];
  fs.writeFileSync(path.join(dir, 'auth.json'), JSON.stringify({ users }, null, 2));
  const d = {
    contacts: [
      mkCard('k1', '_s_u1', CO.c1, { address: '舊址路1號', phone: '02-1111', note: '舊資料一' }),
      mkCard('k2', '_s_u1', CO.c1, { address: '', note: '舊資料二（地址空）' }),
      mkCard('k3', '_s_u2', CO.c2, { address: '乙路2號' }),
    ],
    companies: [JSON.parse(JSON.stringify(CO.c1)), JSON.parse(JSON.stringify(CO.c2))],
    visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], quotations: [], opportunities: [],
  };
  fs.writeFileSync(path.join(dir, 'data.json'), JSON.stringify(d, null, 2));
}
const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
function cleanupFiles(dir) { FILES.forEach((f) => { try { fs.unlinkSync(dir ? path.join(dir, path.basename(f)) : f); } catch { /* 不存在 */ } }); }

// ───────────────── 伺服器（預載腳本在載入時改寫指定檔案＝變異測試，不動磁碟）─────────────────
const PRELOAD = path.join(os.tmpdir(), 'cmail_preload_' + process.pid + '.js');
fs.writeFileSync(PRELOAD, [
  "const Module=require('module'),path=require('path');const orig=Module.prototype._compile;",
  "Module.prototype._compile=function(content,filename){const raw=process.env.CMAIL_MUT_SPEC;",
  "if(raw){const specs=JSON.parse(raw);for(const m of specs){if(path.resolve(filename)===path.resolve(process.env.CMAIL_ROOT,m.file)){for(const e of m.edits){if(!content.includes(e.find))throw new Error('MUTATION ANCHOR MISSING: '+e.find.slice(0,60));content=content.replace(e.find,()=>e.rep);}}}}",
  "return orig.call(this,content,filename);};",
].join('\n') + '\n');
let child = null, childPid = null, MUT_CUR = null;
async function startServer(mut, root) {
  root = root || ROOT; mut = mut || (root === ROOT ? MUT_CUR : null);
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: `http://localhost:${PORT}` });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'CMAIL_MUT_SPEC']) delete env[k];
  if (mut && mut.length) { env.CMAIL_MUT_SPEC = JSON.stringify(mut); env.CMAIL_ROOT = ROOT; }
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
async function freshServer(mut) { await stopServer(); cleanupFiles(); seedFiles(); await startServer(mut); jar = {}; }
async function http(who, method, url, body) {
  const h = {};
  if (body !== undefined && !(body instanceof FormData)) h['Content-Type'] = 'application/json';
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : (body instanceof FormData ? body : JSON.stringify(body)), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: res.status, json, buf, text: buf.toString('utf8') };
}
async function login(who) {
  const res = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status === 200) jar[who] = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
  return res.status;
}
async function loginAll() { for (const w of ['_s_adm', '_s_u1', '_s_u2', '_s_sec', '_s_mkt']) await login(w); }
const post = (who, body) => http(who, 'POST', '/api/contacts', Object.assign({ name: '新名片', company: CO.c1.name, taxId: CO.c1.taxId, email: 'n@example.test', bu: 'ERP' }, body || {}));
const put = (who, id, body) => http(who, 'PUT', '/api/contacts/' + id, body);
const rawContact = (id) => readData().contacts.find((c) => c.id === id);
const auditRows = async () => { await sleep(300); return (await http('_s_adm', 'GET', '/api/admin/contact-audit')).json || []; };

// ═════════════ unit（靜態）═════════════
function groupUnit() {
  console.log('[unit] 靜態檢查');
  const srv = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const html = fs.readFileSync(path.join(ROOT, '_client/index.html'), 'utf8');
  const app = fs.readFileSync(path.join(ROOT, '_client/app.js'), 'utf8');
  const sw = fs.readFileSync(path.join(ROOT, '_client/sw.js'), 'utf8');
  const cf = (srv.match(/const CONTACT_FIELDS\s*=\s*\[([^\]]*)\]/) || [])[1] || '';
  ok(/'mailingAddress'/.test(cf) && /'mailingSameAsCompany'/.test(cf), 'U1 CONTACT_FIELDS 含兩個新欄位');
  ok(/mailingAddress:'郵寄地址'/.test(srv) && /mailingSameAsCompany:'郵寄地址同公司地址'/.test(srv), 'U2 FIELD_LABELS 有中文標籤');
  ok(/id="mailingAddress"/.test(html) && /id="mailingSameAsCompany"/.test(html), 'U3 表單有郵寄地址輸入框與核取方塊');
  ok(/mailingSameAsCompany: \$\('mailingSameAsCompany'\)\.checked/.test(app) && /mailingAddress: \$\('mailingSameAsCompany'\)/.test(app), 'U4 前端儲存 payload 帶兩個欄位');
  ok(/itts-crm-v(3[89]|[4-9][0-9])'/.test(sw) && /v38：名片郵寄地址/.test(sw), 'U5 SW_VERSION 不低於 v38 且 v38 註解有說明（後續版本如秘書唯讀客戶視圖 v39 會往上加）');
  ok(!/mailing[A-Za-z]*\s*[:=][^\n]*innerHTML/.test(app), 'U6 郵寄地址沒有被塞進 innerHTML 賦值');
  const view = (app.match(/<span class="view-label">郵寄地址<\/span>[^\n]*/) || [''])[0];
  ok(/escapeHtml\(c\.mailingAddress\)/.test(view), 'U7 詳情檢視郵寄地址經 escapeHtml');
  for (const f of ['server.js', '_client/app.js', '_client/sw.js']) ok(spawnSync(process.execPath, ['--check', path.join(ROOT, f)]).status === 0, 'U8 node --check ' + f);
}

// ═════════════ api ═════════════
const LONG = (n) => '長'.repeat(n);
function xlsxBuf(headers, rows) {
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, '客戶資料');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}
const upload = (who, url, buf, fields) => { const fd = new FormData(); fd.append('file', new Blob([buf]), 'import.xlsx'); Object.entries(fields || {}).forEach(([k, v]) => fd.append(k, v)); return http(who, 'POST', url, fd); };
async function groupApi() {
  console.log('[api] 真實 server.js');
  await freshServer(); await loginAll();
  // —— POST ——
  let r = await post('_s_u1', { name: 'P1', address: '台北市甲路1號', mailingSameAsCompany: true, mailingAddress: '應被忽略' });
  eq([r.status, r.json && r.json.mailingAddress, r.json && r.json.mailingSameAsCompany], [201, '台北市甲路1號', true], 'A1 POST 勾同公司地址：mailingAddress 等於 address（忽略前端送的值）');
  const p1 = r.json.id;
  eq([rawContact(p1).mailingAddress, rawContact(p1).mailingSameAsCompany], ['台北市甲路1號', true], 'A1b 已寫入資料檔');
  r = await post('_s_u1', { name: 'P2', address: '台北市乙路2號', mailingAddress: '  新北市郵寄路9號  ' });
  eq([r.status, r.json.mailingAddress, 'mailingSameAsCompany' in r.json], [201, '新北市郵寄路9號', false], 'A2 POST 只帶 mailingAddress：存修剪後的值、不憑空多旗標鍵');
  const p2 = r.json.id;
  r = await post('_s_u1', { name: 'P3', address: '丙路3號' });
  ok(!('mailingAddress' in r.json) && !('mailingSameAsCompany' in r.json), 'A3 POST 沒帶新欄位：輸出沒有新鍵');
  r = await post('_s_u1', { name: 'P4', address: '丁路', mailingAddress: '', mailingSameAsCompany: false });
  ok(!('mailingAddress' in r.json) && !('mailingSameAsCompany' in r.json), 'A3b POST 送空值／false：不新增鍵');
  // —— 長度 ——
  r = await post('_s_u1', { name: 'L1', mailingAddress: LONG(201) });
  eq(r.status, 400, 'A4 POST 郵寄地址 201 字 → 400'); ok(/200/.test((r.json || {}).error || ''), 'A4b 錯誤訊息提到 200 字', JSON.stringify(r.json));
  r = await post('_s_u1', { name: 'L2', mailingAddress: LONG(200) });
  eq(r.status, 201, 'A5 POST 郵寄地址剛好 200 字 → 201');
  r = await put('_s_u1', p2, { mailingAddress: LONG(201) });
  eq(r.status, 400, 'A6 PUT 郵寄地址 201 字 → 400');
  eq(rawContact(p2).mailingAddress, '新北市郵寄路9號', 'A6b 被拒絕後資料未變動');
  r = await post('_s_u1', { name: 'L3', mailingAddress: { a: 1 } });
  eq(r.status, 400, 'A7 POST 非字串 → 400');
  r = await post('_s_u1', { name: 'L4', address: LONG(260), mailingSameAsCompany: true });
  eq([r.status, r.json && r.json.mailingAddress === LONG(260)], [201, true], 'A8 勾同公司地址時以 address 為準（address 本來沒有長度限制，不因此被擋）');
  // —— PUT 同步 ——
  r = await put('_s_u1', p1, { address: '台北市新路88號', name: 'P1' });
  eq([r.status, r.json.mailingAddress, r.json.mailingSameAsCompany], [200, '台北市新路88號', true], 'A9 PUT 改 address（旗標沒送、舊值 true）→ mailingAddress 同步');
  eq(rawContact(p1).mailingAddress, '台北市新路88號', 'A9b 資料檔也同步');
  r = await put('_s_u1', p1, { address: '台北市再改1號', mailingAddress: '陳舊分頁送的舊值', mailingSameAsCompany: true });
  eq(r.json.mailingAddress, '台北市再改1號', 'A10 旗標 true 時前端送的 mailingAddress 一律被 address 覆蓋');
  r = await put('_s_u1', p1, { name: '只改名字' });
  eq([r.json.mailingAddress, r.json.mailingSameAsCompany, r.json.address], ['台北市再改1號', true, '台北市再改1號'], 'A11 PUT 沒送新欄位 → 沿用舊值（舊分頁不洗掉）');
  r = await put('_s_u1', p1, { mailingSameAsCompany: false });
  eq([r.json.mailingSameAsCompany, r.json.mailingAddress], [false, '台北市再改1號'], 'A12 取消勾選：保留當前值');
  r = await put('_s_u1', p1, { address: '台北市變更後的公司地址', name: '只改名字' });
  eq(r.json.mailingAddress, '台北市再改1號', 'A13 旗標 false 時改 address 不影響 mailingAddress');
  r = await put('_s_u1', p1, { mailingAddress: '自由輸入的郵寄地址', mailingSameAsCompany: false });
  eq(r.json.mailingAddress, '自由輸入的郵寄地址', 'A14 取消勾選後可自由編輯');
  r = await put('_s_u1', p1, { mailingSameAsCompany: 'true' });
  eq([r.json.mailingSameAsCompany, r.json.mailingAddress], [true, '台北市變更後的公司地址'], 'A15 重新勾選：立刻等於目前 address（字串 "true" 也視為勾選）');
  r = await put('_s_u1', p2, { mailingAddress: '  ' });
  eq([r.status, r.json.mailingAddress], [200, ''], 'A16 清空郵寄地址（已有該鍵）→ 空字串');
  // 取消勾選且 PUT 沒送 mailingAddress 但改 address（同一請求內旗標變 false）
  r = await put('_s_u1', p1, { address: '又換一個地址', mailingSameAsCompany: false });
  eq(r.json.mailingAddress, '台北市變更後的公司地址', 'A17 同一請求取消勾選＋改 address：郵寄地址維持取消前的值');
  // —— 權限：別人改不了 ——
  r = await put('_s_u2', p1, { mailingAddress: '偷改' });
  eq(r.status, 404, 'A18 非擁有者 PUT → 404（沿用既有權限）');
  // —— 審計 ——
  const rows = await auditRows();
  const blob = JSON.stringify(rows);
  ok(/郵寄地址/.test(blob), 'A19 contact-audit 的變更明細含「郵寄地址」中文標籤');
  ok(/郵寄地址同公司地址/.test(blob), 'A19b 旗標變更也有中文標籤');
  const ch = rows.flatMap((x) => (x.changes || (x.extra && x.extra.changes) || [])).filter((c) => c.field === 'mailingAddress');
  ok(ch.length > 0 && ch.every((c) => c.fieldLabel === '郵寄地址'), 'A19c changes 項目 field=mailingAddress 且 fieldLabel=郵寄地址', blob.slice(0, 200));
  // —— 日誌／通知不含地址 ——
  const SENT = '自由輸入的郵寄地址';
  const logs = await http('_s_adm', 'GET', '/api/admin/logs');
  ok(!JSON.stringify(logs.json).includes(SENT) && !JSON.stringify(logs.json).includes('新北市郵寄路9號'), 'A20 一般稽核日誌不含郵寄地址內容');
  let allNotif = '';
  for (const w of ['_s_adm', '_s_u1', '_s_u2', '_s_sec', '_s_mkt']) allNotif += JSON.stringify((await http(w, 'GET', '/api/notifications')).json || []);
  ok(!allNotif.includes(SENT) && !allNotif.includes('新北市郵寄路9號') && !allNotif.includes('台北市變更後的公司地址'), 'A21 通知不含郵寄地址內容');
  // —— GET 帶出 ——
  const g = (await http('_s_u1', 'GET', '/api/contacts')).json;
  const gp1 = g.find((c) => c.id === p1);
  eq([gp1.mailingAddress, gp1.mailingSameAsCompany], ['台北市變更後的公司地址', false], 'A22 GET /api/contacts 整物件輸出（新欄位自然帶出）');
  ok('mailingAddress' in gp1 && 'mailingSameAsCompany' in gp1, 'A22b GET 輸出含兩個新鍵');
  const gk1 = g.find((c) => c.id === 'k1');
  ok(!('mailingAddress' in gk1) && !('mailingSameAsCompany' in gk1), 'A23 舊名片 GET 輸出沒有新鍵');
  // 舊名片 + 前端新版 payload（空值／false）：不憑空新增鍵、稽核不出現新欄位
  const before = (await auditRows()).length;
  r = await put('_s_u1', 'k2', { name: '改了名字', address: '新增的地址', mailingAddress: '', mailingSameAsCompany: false });
  eq(r.status, 200, 'A24 舊名片送新版 payload 成功');
  ok(!('mailingAddress' in rawContact('k2')) && !('mailingSameAsCompany' in rawContact('k2')), 'A24b 舊名片送空值／false：資料檔沒有新鍵');
  const after = (await auditRows());
  const newRows = after.slice(0, Math.max(0, after.length - before));
  ok(!JSON.stringify(after.filter((x) => (x.target || x.targetId) === 'k2' || JSON.stringify(x).includes('k2'))).includes('mailingAddress'), 'A24c 舊名片的稽核明細不含新欄位', JSON.stringify(newRows).slice(0, 200));
  // —— Excel 匯出 ——
  const ex = await http('_s_adm', 'GET', '/api/admin/contacts/export');
  const wb1 = XLSX.read(ex.buf, { type: 'buffer' });
  const rows1 = XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[0]], { defval: '' });
  ok(rows1.length > 0 && '郵寄地址' in rows1[0], 'A25 管理員匯出有「郵寄地址」欄');
  eq(rows1.find((x) => x['聯絡人ID'] === p2) && rows1.find((x) => x['聯絡人ID'] === p2)['郵寄地址'], '', 'A25b 值對應正確（p2 已清空）');
  eq(rows1.find((x) => x['聯絡人ID'] === p1)['郵寄地址'], '台北市變更後的公司地址', 'A25c p1 的郵寄地址值正確');
  const hdr = XLSX.utils.sheet_to_json(wb1.Sheets[wb1.SheetNames[0]], { header: 1 })[0];
  eq(hdr.indexOf('郵寄地址'), hdr.indexOf('地址') + 1, 'A25d 郵寄地址欄緊接在地址欄之後');
  const ex2 = await http('_s_adm', 'GET', '/api/export');
  const wb2 = XLSX.read(ex2.buf, { type: 'buffer' });
  const rows2 = XLSX.utils.sheet_to_json(wb2.Sheets[wb2.SheetNames[0]], { defval: '' });
  ok(rows2.length > 0 && '郵寄地址' in rows2[0], 'A26 一般匯出（/api/export）有「郵寄地址」欄');
  eq(rows2.find((x) => x['姓名'] === '只改名字')['郵寄地址'], '台北市變更後的公司地址', 'A26b 一般匯出值正確');
  const tpl = await http('_s_adm', 'GET', '/api/admin/import-template');
  const wbt = XLSX.read(tpl.buf, { type: 'buffer' });
  ok(XLSX.utils.sheet_to_json(wbt.Sheets[wbt.SheetNames[0]], { header: 1 })[0].includes('郵寄地址'), 'A27 匯入範本有「郵寄地址」欄');
  // —— Excel 匯入 ——
  const H1 = ['姓名', '公司', '地址', '郵寄地址', '統一編號', 'Email'];
  r = await upload('_s_adm', '/api/admin/contacts/import', xlsxBuf(H1, [['匯入甲', '假公司甲', '公司路1號', '郵寄路2號', '90000001', 'a@example.test'], ['匯入乙', '假公司甲', '公司路3號', '', '90000001', 'b@example.test'], ['匯入丙', '假公司甲', '公司路5號', LONG(250), '90000001', 'c@example.test']]), { defaultOwner: '_s_u1' });
  eq(r.status, 200, 'A28 管理員匯入成功', r.text.slice(0, 200));
  let d = readData();
  const im = (n) => d.contacts.find((c) => c.name === n);
  eq([im('匯入甲').address, im('匯入甲').mailingAddress], ['公司路1號', '郵寄路2號'], 'A29 匯入能讀回郵寄地址（地址欄不被「郵寄地址」欄搶走）');
  ok(!('mailingAddress' in im('匯入乙')), 'A30 郵寄地址空白的列：不產生鍵');
  eq(im('匯入丙').mailingAddress, LONG(200), 'A31 匯入超長郵寄地址截斷為 200 字');
  r = await upload('_s_adm', '/api/admin/contacts/import', xlsxBuf(['姓名', '公司', '郵寄地址', '地址', '統一編號', 'Email'], [['欄序反甲', '假公司甲', '郵寄在前', '地址在後', '90000001', 'd@example.test']]), { defaultOwner: '_s_u1' });
  d = readData();
  eq([im('欄序反甲').address, im('欄序反甲').mailingAddress], ['地址在後', '郵寄在前'], 'A32 欄位順序顛倒（郵寄地址欄在前）也對應正確');
  // 往返：匯出 → 匯入
  const ex3 = await http('_s_adm', 'GET', '/api/admin/contacts/export');
  r = await upload('_s_adm', '/api/admin/contacts/import', ex3.buf, { defaultOwner: '_s_u2' });
  d = readData();
  const rt = d.contacts.filter((c) => c.name === '匯入甲');
  eq([rt.length, rt.every((c) => c.mailingAddress === '郵寄路2號')], [2, true], 'A33 匯出檔再匯入：郵寄地址完整往返');
  // 範本匯入端點
  r = await upload('_s_adm', '/api/admin/import-contacts', xlsxBuf(['姓名', '公司', '地址', '郵寄地址', '統一編號', 'Email'], [['範本匯入', '假公司甲', '範本公司路', '範本郵寄路', '90000001', 'e@example.test']]), { targetOwner: '_s_u1' });
  d = readData();
  eq([r.status, im('範本匯入') && im('範本匯入').mailingAddress, im('範本匯入') && im('範本匯入').address], [200, '範本郵寄路', '範本公司路'], 'A34 客戶名單匯入（範本）也能讀郵寄地址');
  // —— 同公司補空白（syncCompanyFields）——
  await stopServer(); cleanupFiles();
  { seedFiles(); const dd = readData();
    dd.contacts.push(mkCard('k4', '_s_u1', CO.c1, { address: '', mailingSameAsCompany: true, mailingAddress: '' }));
    dd.contacts.push(mkCard('k5', '_s_u1', CO.c1, { address: '', mailingAddress: '自訂不動' }));
    fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(dd, null, 2)); }
  await startServer(); jar = {}; await loginAll();
  r = await put('_s_u1', 'k1', { name: '觸發補空白', address: '補空白路7號' });
  eq(r.status, 200, 'A35 k1 存檔觸發 syncCompanyFields');
  eq([rawContact('k4').address, rawContact('k4').mailingAddress, rawContact('k4').mailingSameAsCompany], ['補空白路7號', '補空白路7號', true], 'A35b 已勾旗標且 address 空的同公司名片：補空白後 mailingAddress 一併同步');
  eq([rawContact('k5').address, rawContact('k5').mailingAddress, 'mailingSameAsCompany' in rawContact('k5')], ['補空白路7號', '自訂不動', false], 'A35c 沒勾旗標的同公司名片：mailingAddress 不被動到');
  ok(!('mailingAddress' in rawContact('k2')) && !('mailingSameAsCompany' in rawContact('k2')) && rawContact('k2').address === '補空白路7號', 'A35d 舊名片（無旗標）補空白後仍沒有新鍵');
  await stopServer(); cleanupFiles();
}

// ═════════════ compat（與基準版逐位元比對）═════════════
function exportBase() {
  const ref = process.env.CMAIL_BASE_REF || 'origin/main';
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cmail-base-'));
  const a = spawnSync('git', ['archive', ref], { cwd: ROOT, maxBuffer: 256 * 1024 * 1024 });
  if (a.status !== 0) throw new Error('git archive ' + ref + ' 失敗：' + String(a.stderr));
  const t = spawnSync('tar', ['-x', '-C', dir], { input: a.stdout, maxBuffer: 256 * 1024 * 1024 });
  if (t.status !== 0) throw new Error('tar 失敗：' + String(t.stderr));
  spawnSync('cmd', ['/c', 'mklink', '/J', path.join(dir, 'node_modules'), path.join(ROOT, 'node_modules')], { stdio: 'ignore' });
  return { dir, cleanup() { spawnSync('cmd', ['/c', 'rmdir', path.join(dir, 'node_modules')], { stdio: 'ignore' }); try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ } } };
}
async function groupCompat() {
  console.log('[compat] 舊名片（無新欄位）輸出與基準版逐位元相同');
  const norm = (x) => String(x).replace(/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/g, '<uuid>').replace(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z/g, '<ts>').replace(/\d{4}\/\d{1,2}\/\d{1,2}[^"]*?(?=")/g, '<tl>');
  const run = async (root) => {
    await stopServer(); cleanupFiles(root); seedFiles(root); await startServer(null, root); jar = {};
    await loginAll();
    const out = {};
    const rec = (k, r) => { out[k] = r.status + ' ' + norm(r.text); };
    rec('get1 u1', await http('_s_u1', 'GET', '/api/contacts'));
    rec('get1 u2', await http('_s_u2', 'GET', '/api/contacts'));
    rec('get1 sec', await http('_s_sec', 'GET', '/api/contacts'));
    rec('put k1 舊版payload', await put('_s_u1', 'k1', { name: '改名', company: '假公司甲', title: '經理', address: '新址路2號', taxId: '90000001', isPrimary: true, employmentStatus: 'active', email: 'k1@example.test' }));
    rec('put k1 只改備註', await put('_s_u1', 'k1', { note: '只改備註' }));
    rec('put k2 清地址', await put('_s_u1', 'k2', { address: '補上地址' }));
    rec('put k2 非擁有者', await put('_s_u2', 'k2', { note: 'x' }));
    rec('post 無新欄位', await post('_s_u1', { name: '新增一', address: '新增路1號', phone: '02-2222' }));
    rec('post 超長地址', await post('_s_u1', { name: '新增二', address: LONG(260) }));
    rec('post 缺公司', await http('_s_u1', 'POST', '/api/contacts', { name: 'x' }));
    rec('get2 u1', await http('_s_u1', 'GET', '/api/contacts'));
    const au = await http('_s_adm', 'GET', '/api/admin/contact-audit'); await sleep(300);
    const au2 = await http('_s_adm', 'GET', '/api/admin/contact-audit'); rec('audit', au2); void au;
    const raw = JSON.parse(fs.readFileSync(path.join(root, 'data.json'), 'utf8'));
    out.rawContacts = norm(JSON.stringify(raw.contacts.map((c) => { const o = Object.assign({}, c); delete o.updatedAt; return o; }), null, 1));
    out.rawCompanies = norm(JSON.stringify(raw.companies.map((c) => Object.keys(c).sort())));
    const imp = await upload('_s_adm', '/api/admin/contacts/import', xlsxBuf(['姓名', '公司', '地址', '統一編號', 'Email'], [['舊匯入', '假公司乙', '匯入路', '90000002', 'z@example.test']]), { defaultOwner: '_s_u2' });
    out.imp = imp.status + ' ' + norm(imp.text);
    const raw2 = JSON.parse(fs.readFileSync(path.join(root, 'data.json'), 'utf8'));
    out.impRaw = norm(JSON.stringify(raw2.contacts.filter((c) => c.name === '舊匯入').map((c) => { const o = Object.assign({}, c); delete o.updatedAt; return o; })));
    const imp2 = await upload('_s_adm', '/api/admin/import-contacts', xlsxBuf(['姓名', '公司', '地址', '統一編號', 'Email'], [['舊匯入二', '假公司乙', '匯入路二', '90000002', 'y@example.test']]), { targetOwner: '_s_u2' });
    out.imp2 = imp2.status + ' ' + norm(imp2.text);
    const raw3 = JSON.parse(fs.readFileSync(path.join(root, 'data.json'), 'utf8'));
    out.imp2Raw = norm(JSON.stringify(raw3.contacts.filter((c) => c.name === '舊匯入二').map((c) => { const o = Object.assign({}, c); delete o.updatedAt; return o; })));
    // 匯出：舊欄位的值必須相同（新增的郵寄地址欄是設計上的差異，比對時排除該欄）
    for (const [k, u] of [['exp admin', '/api/admin/contacts/export'], ['exp user', '/api/export']]) {
      const e = await http('_s_adm', 'GET', u);
      const wb = XLSX.read(e.buf, { type: 'buffer' });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' }).map((x) => { delete x['郵寄地址']; delete x['建立時間']; return x; });
      out[k] = norm(JSON.stringify(rows));
    }
    await stopServer(); cleanupFiles(root);
    return out;
  };
  const bx = exportBase();
  let A, B;
  try { A = await run(bx.dir); B = await run(ROOT); } finally { bx.cleanup(); }
  const diff = Object.keys(A).filter((k) => A[k] !== B[k]);
  ok(diff.length === 0, `C1 舊名片所有既有端點輸出與基準版逐位元相同（比對 ${Object.keys(A).length} 項）`, diff.join(', '));
  if (diff.length) console.log('    差異首項：', diff[0], '\n      前：', String(A[diff[0]]).slice(0, 400), '\n      後：', String(B[diff[0]]).slice(0, 400));
  ok(!JSON.stringify(B).includes('mailing'), 'C2 修改後的輸出（含資料檔、稽核）完全沒有出現 mailing 字樣');
}

// ═════════════ 變異 ═════════════
const MUTANTS = [
  { name: 'M1 PUT 拿掉同步（旗標 true 時不再跟著 address）', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: "if (updated.mailingSameAsCompany === true) updated.mailingAddress = updated.address || '';", rep: '' }] }] },
  { name: 'M2 旗標沒送時不沿用舊值（只在請求帶旗標時才同步）', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: 'if (updated.mailingSameAsCompany === true) updated.mailingAddress', rep: 'if (safeBody.mailingSameAsCompany !== undefined && updated.mailingSameAsCompany === true) updated.mailingAddress' }] }] },
  { name: 'M3 拿掉長度限制', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: "if ([...v.trim()].length > MAILING_MAX_LEN) return '郵寄地址不可超過 ' + MAILING_MAX_LEN + ' 字';", rep: '' }] }] },
  { name: 'M4 匯入欄位對應拿掉（郵寄地址欄讀不到）', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: "mailingAddress: col('郵寄地址'),", rep: 'mailingAddress: -1,' }] }] },
  { name: 'M5 匯入地址欄被「郵寄地址」欄搶走', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: "address:        header.findIndex(h => h.includes('地址') && !h.includes('郵寄')),", rep: "address:        header.findIndex(h => h.includes('地址'))," }] }] },
  { name: 'M6 舊名片送空值／false 也憑空寫入鍵', groups: ['api', 'compat'], mut: [{ file: 'server.js', edits: [{ find: "if (f || 'mailingSameAsCompany' in old) updated.mailingSameAsCompany = f; else delete updated.mailingSameAsCompany;", rep: 'updated.mailingSameAsCompany = f;' }, { find: "if (m || 'mailingAddress' in old) updated.mailingAddress = m; else delete updated.mailingAddress;", rep: 'updated.mailingAddress = m;' }] }] },
  { name: 'M7 POST 勾同公司地址卻存空值', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: 'contact.mailingSameAsCompany = true; contact.mailingAddress = contact.address; }', rep: "contact.mailingSameAsCompany = true; contact.mailingAddress = ''; }" }] }] },
  { name: 'M8 拿掉 FIELD_LABELS 中文標籤', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: "mailingAddress:'郵寄地址', mailingSameAsCompany:'郵寄地址同公司地址'", rep: '' }] }] },
  { name: 'M9 CONTACT_FIELDS 漏掉新欄位（PUT 全被丟棄）', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: ",'mailingAddress','mailingSameAsCompany'];", rep: '];' }] }] },
  { name: 'M10 對舊名片也外洩出新鍵（POST 一律帶 mailingAddress）', groups: ['api', 'compat'], mut: [{ file: 'server.js', edits: [{ find: "    createdAt: new Date().toISOString()\n  };\n  // 郵寄地址：", rep: "    createdAt: new Date().toISOString(), mailingAddress: ''\n  };\n  // 郵寄地址：" }] }] },
  { name: 'M12 補空白同步拿掉（已勾旗標名片 mailingAddress 不跟著補上的地址）', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: "if (field === 'address' && c.mailingSameAsCompany === true) c.mailingAddress = c.address;", rep: '' }] }] },
  { name: 'M11 匯出少了郵寄地址欄', groups: ['api'], mut: [{ file: 'server.js', edits: [{ find: "    c.mailingAddress || '',\n", rep: '' }, { find: "'地址', '郵寄地址', '網站'", rep: "'地址', '網站'" }] }] },
];
async function runMutants() {
  console.log('[變異] 每個變異都必須讓對應群組變紅');
  for (const m of MUTANTS) {
    if (process.env.CMAIL_MUT_PICK && !new RegExp(process.env.CMAIL_MUT_PICK).test(m.name)) continue;
    const snap = { pass, fail, n: failures.length };
    quiet = true;
    try {
      for (const g of m.groups) {
        MUT_CUR = m.mut;
        try { await ({ api: groupApi, compat: groupCompat })[g](); } catch (e) { fail++; failures.push('變異執行例外：' + e.message); }
      }
    } finally { quiet = false; MUT_CUR = null; await stopServer(); cleanupFiles(); }
    const newFails = failures.length - snap.n;
    const names = failures.slice(snap.n, snap.n + 3).join(' / ');
    pass = snap.pass; fail = snap.fail; failures.length = snap.n;
    ok(newFails > 0, `${m.name} → 測試變紅（${newFails} 項失敗：${names.slice(0, 90)}）`);
    console.log('  ' + (newFails > 0 ? '✓ 變紅' : '✗ 沒變紅') + '：' + m.name + '（' + newFails + ' 項）');
  }
}

const api = { ok, eq, sleep, http, login, loginAll, freshServer: (...a) => freshServer(...a), stopServer, cleanupFiles, readData, post, put, rawContact, seedFiles, PW, BASE, ROOT, SHOTS, CDP, XSS, snap: () => ({ pass, fail, n: failures.length }), restore: (x) => { pass = x.pass; fail = x.fail; failures.length = x.n; }, newFails: (x) => failures.slice(x.n), setQuiet: (q) => { quiet = q; } };
module.exports = api;
if (require.main === module) {
  (async () => {
    try {
      if (want('unit')) groupUnit();
      if (want('api')) await groupApi();
      if (want('compat')) await groupCompat();
      if (want('ui')) await require('./_contact-mailing-ui.js').run(api);
      if (process.env.CMAIL_MUT === '1') {
        await runMutants();
        if (want('ui')) await require('./_contact-mailing-ui.js').mutants(api);
      }
    } catch (e) {
      fail++; failures.push('例外：' + (e && e.stack || e)); console.log('例外：', e && e.stack || e);
    } finally {
      await stopServer(); cleanupFiles();
      try { fs.unlinkSync(PRELOAD); } catch { /* 不存在 */ }
    }
    console.log(`\n名片郵寄地址檢查：PASS ${pass} / FAIL ${fail}`);
    if (fail) console.log('失敗項目：\n  ' + failures.join('\n  '));
    process.exit(fail ? 1 : 0);
  })();
}
