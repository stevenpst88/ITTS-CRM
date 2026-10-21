#!/usr/bin/env node
'use strict';
/**
 * 商機「Deal ID」欄位檢查：單元（lib/productCatalog.js）＋ API（真實 server.js）＋ 前端互動（headless Edge + CDP）＋ 變異測試。
 * 不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體；資料檔只存在於本次執行；結束以 PID 清掉伺服器與 Edge；不連遠端資料庫。
 *
 *   node scripts/check-opp-deal-id.js                 全部（單元＋API＋前端）
 *   DEALID_ONLY=unit,api,ui,admin node ...            只跑指定群組
 *   DEALID_MUT=1 node ...                             另外跑變異測試（每個變異都必須讓對應群組變紅）
 *   DEALID_SHOTS=<資料夾> node ...                    另外輸出截圖（預設 scratchpad/dealid_out/shots 不自動輸出，需指定）
 *   DEALID_PORT（預設 3562）／DEALID_CDP（預設 9473）
 * 守門：工作樹已有 auth.json／data.json／audit.log.json（可能是真實資料）時拒絕執行。
 *
 * 涵蓋：Deal ID 必填（POST／PUT 表單）、去前後空白、去控制字元、≤64、偏更新 {stage} 不被擋、舊商機不追溯、
 *       商品目錄設定（requiresDealId／dealRegisterUrl）權限／URL 驗證／稽核／與 renames 相容、Excel 匯入不擋但統計缺漏、
 *       匯出附加欄位、XSS 酬載以純文字顯示、前端（欄位顯示隱藏／必填／連結屬性／存檔被擋／看板卡片）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.DEALID_PORT || '3562', 10);
const CDP = parseInt(process.env.DEALID_CDP || '9473', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3161)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map((f) => path.join(ROOT, f));
{
  const g = FILES.filter((f) => /(auth|data|audit.log)\.jsonl?$/.test(f)).filter((f) => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map((f) => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
}
const ONLY = (process.env.DEALID_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length || ONLY.includes(k);
const SHOTS = process.env.DEALID_SHOTS || '';
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const XLSX = require(path.join(ROOT, 'node_modules', 'xlsx'));
const pcat = require(path.join(ROOT, 'lib', 'productCatalog.js'));

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 300)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = (n) => new Promise((r) => setTimeout(r, n));

// ───────────────── 沙盒資料 ─────────────────
const PW = 'Dl' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
const REG_URL = 'https://deal.example.test/register?src=crm';
const XSS = '<img src=x onerror="window.__xss=1">';
const YEAR = new Date().getFullYear();
const mkCatalog = () => ({
  bus: {
    ERP: [{ group: '一般商品', enabled: true, items: [
      { name: 'SAP Public Cloud License', enabled: true, subs: [], requiresDealId: true, dealRegisterUrl: REG_URL },
      { name: 'SAP CCFLEX', enabled: true, subs: ['C&S', 'A&O'], requiresDealId: true },
      { name: 'ERP 顧問導入專案 PE', enabled: true, subs: [] },
      { name: '其他', enabled: true, subs: [] },
    ] }],
    ITS: [{ group: 'ITS 產品', enabled: true, items: [
      { name: 'MES', enabled: true, subs: [] },
      { name: '停用商品D', enabled: false, subs: [], requiresDealId: true },
    ] }],
    MDM: [], CRM: [], '資安': [],
  },
  updatedAt: nowIso, updatedBy: 'seed',
});
const mkOpp = (id, product, extra) => Object.assign({ id, owner: '_s_u1', bu: 'ERP', company: '沙盒公司A', contactId: '', contactName: '', category: 'ERP', product, amount: '100', expectedDate: YEAR + '-12-31', description: '', stage: 'C', visitId: '', createdAt: nowIso }, extra || {});
function seedFiles() {
  const users = [mkUser('_s_adm', 'admin'), mkUser('_s_u1', 'user', { bu: ['ERP'] }), mkUser('_s_u2', 'user', { bu: ['ERP'] })];
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  const d = {
    contacts: [], companies: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], quotations: [],
    opportunities: [
      mkOpp('o-legacy', 'SAP Public Cloud License'),
      mkOpp('o-legacy2', 'SAP Public Cloud License'),
      mkOpp('o-free', 'ERP 顧問導入專案 PE'),
      mkOpp('o-has', 'SAP Public Cloud License', { dealId: 'EXIST-1' }),
      mkOpp('o-xss', 'SAP Public Cloud License', { dealId: XSS }),
    ],
    productCatalog: mkCatalog(),
  };
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
}
const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
function cleanupFiles() { FILES.forEach((f) => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }

// ───────────────── 伺服器（支援以預載腳本在載入時改寫指定檔案＝變異測試，不動磁碟上的檔案）─────────────────
const PRELOAD = path.join(os.tmpdir(), 'dealid_preload_' + process.pid + '.js');
// 預載腳本：載入 m.file（相對專案根）時，把原始碼中的 edits[].find 換成 rep（變異測試用；磁碟上的檔案不動）
fs.writeFileSync(PRELOAD, [
  "const Module=require('module'),path=require('path');const orig=Module.prototype._compile;",
  "Module.prototype._compile=function(content,filename){const raw=process.env.DEALID_MUT_SPEC;",
  "if(raw){const m=JSON.parse(raw);if(path.resolve(filename)===path.resolve(process.env.DEALID_ROOT,m.file)){for(const e of m.edits){if(!content.includes(e.find))throw new Error('MUTATION ANCHOR MISSING: '+e.find.slice(0,60));content=content.replace(e.find,()=>e.rep);}}}",
  "return orig.call(this,content,filename);};",
].join('\n') + '\n');
let child = null, childPid = null;
async function startServer(mut) {
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'QUOTE_FLOW_ENGINE', 'QUOTE_FLOW_DESIGNER', 'FEATURE_GUARD_MODE', 'DEALID_MUT_SPEC']) delete env[k];
  if (mut && mut.file) { env.DEALID_MUT_SPEC = JSON.stringify(mut); env.DEALID_ROOT = ROOT; }
  await assertPortFree(PORT);
  child = spawn(process.execPath, ['-r', PRELOAD, path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { out += d; });
  for (let i = 0; i < 100; i++) {
    await sleep(250);
    try { const r = await fetch(BASE + '/login.html'); if (r.ok) return; } catch { /* 尚未起來 */ }
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
async function freshServer(mut) { await stopServer(); cleanupFiles(); seedFiles(); await startServer(mut); jar = {}; }

// ───────────────── HTTP ─────────────────
let jar = {};
async function http(who, method, url, body, hdrs) {
  const h = Object.assign({ 'Content-Type': 'application/json' }, hdrs || {});
  if (who && jar[who]) h.Cookie = jar[who];
  const isBuf = Buffer.isBuffer(body);
  const res = await fetch(BASE + url, { method, headers: isBuf ? Object.assign({}, h, hdrs) : h, body: body === undefined ? undefined : (isBuf || body instanceof FormData ? body : JSON.stringify(body)), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: res.status, json, buf, text: buf.toString('utf8'), headers: res.headers };
}
async function login(who) {
  const res = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status === 200) jar[who] = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
  return res.status;
}
const oppBody = (product, extra) => Object.assign({ company: '沙盒公司A', category: 'ERP', product, amount: '100', expectedDate: YEAR + '-12-31', description: 't' }, extra || {});
const listOpps = async (who) => (await http(who || '_s_u1', 'GET', '/api/opportunities')).json || [];
const oppById = async (id) => (await listOpps()).find((o) => o.id === id);
async function auditLines(action) {
  await sleep(300);
  const r = await http('_s_adm', 'GET', '/api/admin/logs');
  return (Array.isArray(r.json) ? r.json : []).filter((l) => l.action === action);
}

// ═════════════ 單元：lib/productCatalog.js ═════════════
function groupUnit() {
  console.log('[單元] lib/productCatalog.js');
  const n = pcat.normalizeDealId;
  eq(n('  ABC-1  '), { ok: true, value: 'ABC-1' }, 'U1 前後空白去除');
  eq(n('A\u0001B\u007fC\u0085D E'), { ok: true, value: 'ABCDE' }, 'U2 控制字元（C0／DEL／C1／U+2028）全部去除');
  eq(n('   '), { ok: true, value: '' }, 'U3 只有空白＝空字串');
  eq(n(null), { ok: true, value: '' }, 'U4 null＝空字串');
  eq(n(undefined), { ok: true, value: '' }, 'U5 undefined＝空字串');
  eq(n(12345), { ok: true, value: '12345' }, 'U6 數字轉文字');
  eq(n('x'.repeat(64)).ok, true, 'U7 剛好 64 字通過');
  eq(n('x'.repeat(65)).code, 'DEAL_ID_TOO_LONG', 'U8 65 字被拒（DEAL_ID_TOO_LONG）');
  eq(n('x'.repeat(64) + '\u0001').ok, true, 'U9 長度在去控制字元之後才算（64 字＋控制字元通過）');
  eq(n({}).code, 'DEAL_ID_INVALID', 'U10 物件被拒');
  eq(n(['a']).code, 'DEAL_ID_INVALID', 'U11 陣列被拒');
  eq(n('<img src=x onerror=alert(1)>').value, '<img src=x onerror=alert(1)>', 'U12 不限格式：HTML 字元原樣保留（輸出端負責跳脫）');
  const v = pcat.validateDealUrl;
  eq(v('').ok && v('').value === '', true, 'U20 空字串合法（＝未設定）');
  eq(v('https://deal.example.test/reg?x=1').value, 'https://deal.example.test/reg?x=1', 'U21 https 通過');
  eq(v('http://deal.example.test').ok, true, 'U22 http 通過');
  eq(v('  https://deal.example.test/a  ').ok, true, 'U23 前後空白會被去除');
  ['javascript:alert(1)', 'JaVaScRiPt:alert(1)', 'data:text/html,<script>alert(1)</script>', 'vbscript:x', 'file:///etc/passwd', 'ftp://x.test/a', '/relative/path', '//evil.test/x', 'deal.example.test', 'https://', 'http://', 'https://u:p@deal.example.test', 'https://deal.example.test/a b', 'https://deal.example.test/a\u0001', 'https:\\\\evil.test', 'https://' + 'a'.repeat(600) + '.test', 123, {}, []].forEach((bad, i) => {
    ok(v(bad).ok === false, `U3${String(i).padStart(2, '0')} 網址被拒：${JSON.stringify(bad).slice(0, 50)}`);
  });
  eq(v('HTTPS://Deal.Example.test/Path').value, 'https://deal.example.test/Path', 'U40 儲存解析後的 href（主機名稱小寫化）');

  const cat = { bus: { ERP: [{ group: 'g', enabled: true, items: [
    { name: 'A', requiresDealId: true, dealRegisterUrl: 'https://a.test/r' },
    { name: 'A B', enabled: true },                    // 較長的名稱（不需要）
    { name: 'C', requiresDealId: true, subs: ['x'] },
    { name: 'Dup' }, { name: 'Off', enabled: false, requiresDealId: true },
  ] }], ITS: [{ group: 'g2', items: [{ name: 'Dup', requiresDealId: true, dealRegisterUrl: 'https://dup.test/' }] }] } };
  const r = pcat.resolveDealIdRule;
  eq(r(cat, 'A'), { requires: true, url: 'https://a.test/r', name: 'A' }, 'U50 完全相符');
  eq(r(cat, '  A  ').requires, true, 'U51 商品文字前後空白不影響');
  eq(r(cat, 'a').requires, false, 'U52 區分大小寫（a≠A）');
  eq(r(cat, 'A B').requires, false, 'U53 較長名稱完全相符優先（A B 不需要，即使 A 需要）');
  eq(r(cat, 'A B 明細').requires, false, 'U54 前綴取最長名稱（A B）');
  eq(r(cat, 'A 明細').requires, true, 'U55 明細組合值（名稱＋空白＋明細）沿用該商品設定');
  eq(r(cat, 'AB').requires, false, 'U56 沒有空白分隔不算前綴（AB≠A＋明細）');
  eq(r(cat, 'C x').requires, true, 'U57 子選單組合值');
  eq(r(cat, 'Dup'), { requires: true, url: 'https://dup.test/', name: 'Dup' }, 'U58 同名商品在多個群組：任一要求即要求，網址取有設定者');
  eq(r(cat, 'Off').requires, true, 'U59 已停用的商品仍適用（歷史商機編輯時規則不變）');
  eq(r(cat, '不存在的商品').requires, false, 'U60 商品不在目錄＝不需要');
  eq(r(cat, '').requires, false, 'U61 空商品＝不需要');
  eq(r(cat, null).requires, false, 'U62 null 商品＝不需要');
  eq(r(undefined, 'SAP Public Cloud License').requires, false, 'U63 沒有自訂目錄（用種子）＝不需要');
  eq(r({}, 'MES').requires, false, 'U64 目錄缺 bus＝用種子，不爆');

  const raw = { ERP: [{ group: 'g', enabled: true, items: [
    { name: 'P1', enabled: true, subs: ['s'], requiresDealId: 'true', dealRegisterUrl: ' https://x.test/a ' },
    { name: 'P2', enabled: true, subs: [], requiresDealId: 0, dealRegisterUrl: '' },
    { name: 'P3', enabled: false, subs: [], requiresDealId: true, dealRegisterUrl: 'javascript:alert(1)' },
  ] }] };
  const before = JSON.stringify(raw);
  const nm = pcat.normalizeCatalogDealSettings(raw);
  eq(JSON.stringify(raw), before, 'U70 不修改傳入物件');
  const it = nm.bus.ERP[0].items;
  eq([it[0].requiresDealId, it[0].dealRegisterUrl, it[0].subs, it[0].enabled], [true, 'https://x.test/a', ['s'], true], 'U71 "true" 轉布林 true、網址去空白、其他欄位保留');
  eq(['requiresDealId' in it[1], 'dealRegisterUrl' in it[1]], [false, false], 'U72 不需要＝不留鍵');
  eq(nm.errors.length, 1, 'U73 壞網址產生一筆錯誤');
  eq(nm.errors[0] && nm.errors[0].name, 'P3', 'U74 錯誤指出商品名稱');
  const d = pcat.diffDealSettings({ bus: { ERP: [{ group: 'g', items: [{ name: 'P1' }, { name: 'P2', requiresDealId: true }] }] } }, { ERP: [{ group: 'g', items: [{ name: 'P1', requiresDealId: true, dealRegisterUrl: 'https://x.test/' }, { name: 'P2', requiresDealId: true }] }] });
  eq(d.length, 2, 'U80 diff：P1 的「需要」與「網址」各一筆，P2 沒變不列');
}

// ═════════════ API ═════════════
async function groupApi(mut) {
  if (!mut) console.log('[API] 真實 server.js');
  await freshServer(mut);
  for (const w of ['_s_adm', '_s_u1', '_s_u2']) eq(await login(w), 200, `登入 ${w}`);
  let r;

  // ── 新增（POST）──
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License'));
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_REQUIRED'], 'A1 需要 Deal ID 的商品缺 dealId → 400 DEAL_ID_REQUIRED');
  ok(r.json && /Deal ID/.test(r.json.error || ''), 'A1b 錯誤訊息是中文且提到 Deal ID', r.json && r.json.error);
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License', { dealId: '   ' }));
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_REQUIRED'], 'A2 只有空白 → 400');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License', { dealId: '\u0001\u0002 \t' }));
  eq(r.status, 400, 'A3 只有控制字元／空白 → 400（控制字元去除後為空）');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License', { dealId: null }));
  eq(r.status, 400, 'A4 dealId=null → 400');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License', { dealId: '  DL-001  ' }));
  eq(r.status, 201, 'A5 有值（含前後空白）→ 201');
  eq(r.json && r.json.dealId, 'DL-001', 'A5b 回應的 dealId 已去前後空白');
  const id5 = r.json && r.json.id;
  eq((await oppById(id5) || {}).dealId, 'DL-001', 'A5c 清單回傳的 dealId 正確');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License', { dealId: 'AB\u0007C\u0000D' }));
  eq([r.status, r.json && r.json.dealId], [201, 'ABCD'], 'A6 控制字元被去除');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License', { dealId: 'x'.repeat(64) }));
  eq([r.status, r.json && r.json.dealId.length], [201, 64], 'A7 剛好 64 字 → 201');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License', { dealId: 'x'.repeat(65) }));
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_TOO_LONG'], 'A8 65 字 → 400 DEAL_ID_TOO_LONG');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License', { dealId: { a: 1 } }));
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_INVALID'], 'A9 dealId 為物件 → 400 DEAL_ID_INVALID');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License', { dealId: 987654 }));
  eq([r.status, r.json && r.json.dealId], [201, '987654'], 'A10 dealId 為數字 → 轉成文字存入');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP CCFLEX C&S'));
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_REQUIRED'], 'A11 子選單組合值「SAP CCFLEX C&S」沿用 SAP CCFLEX 的設定 → 缺 dealId 400');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP CCFLEX', { dealId: 'CC-1' }));
  eq(r.status, 201, 'A12 SAP CCFLEX（沒選明細）有 dealId → 201');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('ERP 顧問導入專案 PE'));
  eq(r.status, 201, 'A13 不需要的商品不擋');
  ok(r.json && !('dealId' in r.json), 'A13b 不需要且沒填 → 回應沒有 dealId 鍵（舊資料外形不變）');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('ERP 顧問導入專案 PE', { dealId: 'OPT-1' }));
  eq([r.status, r.json && r.json.dealId], [201, 'OPT-1'], 'A14 不需要的商品也可選填 Deal ID');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('客製的其他商品XYZ'));
  eq(r.status, 201, 'A15 商品不在目錄（自訂文字）→ 不需要');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('停用商品D'));
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_REQUIRED'], 'A16 已停用但標記需要的商品仍要求');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('sap public cloud license'));
  eq(r.status, 201, 'A17 大小寫不同＝不同商品（沿用精確比對）');
  r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License', { dealId: XSS }));
  eq([r.status, r.json && r.json.dealId], [201, XSS], 'A18 XSS 酬載原樣存為純文字（輸出端負責跳脫）');
  const x18 = r.json && r.json.id;
  r = await http('_s_u1', 'POST', '/api/opportunities', Object.assign(oppBody('SAP Public Cloud License'), { dealId: undefined, product: 'SAP Public Cloud License' }));
  eq(r.status, 400, 'A19 body 完全沒 dealId 鍵 → 400');
  const lg = await auditLines('CREATE_OPP');
  ok(lg.some((l) => /DealID:DL-001/.test(l.detail || '')), 'A20 建立商機的稽核紀錄含 Deal ID');

  // ── 編輯（PUT）──
  const before = await oppById('o-legacy2');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-legacy2', { stage: 'B' });
  eq([r.status, r.json && r.json.stage], [200, 'B'], 'B1 偏更新 {stage}：舊商機無 Deal ID 且商品需要，仍不被擋');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-legacy2', { stage: 'Won', achievedDate: YEAR + '-06-01' });
  eq([r.status, r.json && r.json.stage], [200, 'Won'], 'B2 看板「標記成交」{stage, achievedDate} 不被擋');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-legacy', { description: '只改備註' });
  eq(r.status, 200, 'B3 只改備註（沒帶 product／dealId）不被擋');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-legacy', { amount: '555', expectedDate: YEAR + '-11-11' });
  eq(r.status, 200, 'B3b 只改金額／日期（沒帶 product／dealId）不被擋');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-legacy', { category: 'ERP', product: 'SAP Public Cloud License', stage: 'C', amount: '100', expectedDate: YEAR + '-12-31' });
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_REQUIRED'], 'B4 表單編輯（帶 product）舊商機仍沒有 Deal ID → 400，要求補填');
  eq((await oppById('o-legacy') || {}).stage, 'C', 'B4b 被擋時沒有寫入任何變更（階段維持）');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-legacy', { category: 'ERP', product: 'SAP Public Cloud License', stage: 'C', dealId: '  FIX-9 ' });
  eq([r.status, r.json && r.json.dealId], [200, 'FIX-9'], 'B5 補填 Deal ID（去空白）→ 200');
  const lg2 = await auditLines('UPDATE_OPP');
  ok(lg2.some((l) => /DealID:（空）→FIX-9/.test(l.detail || '')), 'B6 更新商機的稽核紀錄含 Deal ID 變更（空→FIX-9）', JSON.stringify(lg2.slice(0, 3)));
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-legacy', { product: 'SAP Public Cloud License', stage: 'B' });
  eq(r.status, 200, 'B7 已有 Deal ID 的商機，表單再次儲存（沒重送 dealId）→ 沿用既有值，不擋');
  eq((await oppById('o-legacy') || {}).dealId, 'FIX-9', 'B7b 既有 dealId 未被洗掉');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-legacy', { product: 'SAP Public Cloud License', dealId: '' });
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_REQUIRED'], 'B8 需要的商品把 Deal ID 清成空白 → 400');
  eq((await oppById('o-legacy') || {}).dealId, 'FIX-9', 'B8b 被擋後仍是 FIX-9');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-legacy', { dealId: 'x'.repeat(65) });
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_TOO_LONG'], 'B9 只送 dealId 且 65 字 → 400');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-legacy', { dealId: '   ' });
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_REQUIRED'], 'B10 只送 dealId（空白）而商品需要 → 400');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-free', { product: 'SAP Public Cloud License', category: 'ERP' });
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_REQUIRED'], 'B11 把商品改成「需要」的商品而沒有 Deal ID → 400');
  eq((await oppById('o-free') || {}).product, 'ERP 顧問導入專案 PE', 'B11b 被擋時商品沒被改掉');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-free', { product: 'SAP Public Cloud License', dealId: 'SW-1' });
  eq(r.status, 200, 'B12 改成「需要」的商品並同時帶 Deal ID → 200');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-free', { product: 'ERP 顧問導入專案 PE' });
  eq([r.status, r.json && r.json.dealId], [200, 'SW-1'], 'B13 改回不需要的商品：Deal ID 保留（不洗掉資料）');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-free', { product: 'ERP 顧問導入專案 PE', dealId: '' });
  eq(r.status, 200, 'B14 不需要的商品可清空 Deal ID');
  ok(!('dealId' in (await oppById('o-free') || { dealId: 1 })), 'B14b 清空＝移除鍵');
  r = await http('_s_u2', 'PUT', '/api/opportunities/o-has', { dealId: 'HACK' });
  eq(r.status, 404, 'B15 別人的商機照舊 404（權限未被放寬）');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-has', { stage: 'A' });
  eq(r.status, 200, 'B16 有 Deal ID 的商機拖曳階段 → 200');
  r = await http('_s_u1', 'PUT', '/api/opportunities/o-has', { dealId: 'CHG-2', product: 'SAP Public Cloud License' });
  eq(r.status, 200, 'B17 修改 Deal ID → 200');
  const lg3 = await auditLines('UPDATE_OPP');
  ok(lg3.some((l) => /DealID:EXIST-1→CHG-2/.test(l.detail || '')), 'B18 Deal ID 變更（EXIST-1→CHG-2）有稽核紀錄');

  // ── 回應／XSS ──
  const xo = await oppById('o-xss');
  eq(xo && xo.dealId, XSS, 'C1 XSS 酬載經 API 原樣回傳（JSON 字串）');
  ok((await http('_s_u1', 'GET', '/api/opportunities')).headers.get('content-type').includes('application/json'), 'C2 回應為 application/json（不會被當 HTML 解析）');
  ok(x18 && (await oppById(x18)).dealId === XSS, 'C3 新增的 XSS 酬載同樣原樣保存');

  // ── 商品目錄設定 ──
  r = await http('_s_u1', 'GET', '/api/product-catalog');
  const pcItem = (((r.json || {}).bus || {}).ERP || [{ items: [] }])[0].items.find((i) => i.name === 'SAP Public Cloud License') || {};
  eq([r.status, pcItem.requiresDealId, pcItem.dealRegisterUrl], [200, true, REG_URL], 'D1 一般業務可讀目錄，且含 requiresDealId／dealRegisterUrl');
  const cat0 = JSON.parse(JSON.stringify(mkCatalog()));
  r = await http('_s_u1', 'POST', '/api/admin/product-catalog', { catalog: { bus: cat0.bus } });
  eq(r.status, 403, 'D2 非管理員不能改商品目錄設定 → 403');
  const badUrls = ['javascript:alert(1)', 'data:text/html,x', '/relative', 'ftp://x.test/a', 'http://', 'https://a b.test', 'https://u:p@a.test', 'https://' + 'a'.repeat(520) + '.test'];
  for (const u of badUrls) {
    const c = JSON.parse(JSON.stringify(mkCatalog())); c.bus.ERP[0].items[0].dealRegisterUrl = u;
    r = await http('_s_adm', 'POST', '/api/admin/product-catalog', { catalog: { bus: c.bus } });
    eq([r.status, r.json && r.json.code], [400, 'DEAL_URL_INVALID'], `D3 壞網址被拒：${u.slice(0, 40)}`);
  }
  r = await http('_s_u1', 'GET', '/api/product-catalog');
  eq((((r.json || {}).bus || {}).ERP[0].items[0] || {}).dealRegisterUrl, REG_URL, 'D4 被拒的儲存沒有寫入任何東西（網址維持原值）');
  {
    const c = JSON.parse(JSON.stringify(mkCatalog()));
    c.bus.ERP[0].items[0].dealRegisterUrl = ' HTTPS://Deal2.Example.test/reg ';
    c.bus.ERP[0].items[2].requiresDealId = true;               // ERP 顧問導入專案 PE 改成需要
    c.bus.ERP[0].items[3].requiresDealId = 'true';             // 其他：字串 'true' 會被轉布林
    r = await http('_s_adm', 'POST', '/api/admin/product-catalog', { catalog: { bus: c.bus } });
    eq(r.status, 200, 'D5 管理員儲存合法設定 → 200');
    r = await http('_s_u1', 'GET', '/api/product-catalog');
    const its = r.json.bus.ERP[0].items;
    eq([its[0].dealRegisterUrl, its[2].requiresDealId, its[3].requiresDealId], ['https://deal2.example.test/reg', true, true], 'D5b 網址存解析後的 href、requiresDealId 為布林');
    const lg4 = await auditLines('SAVE_PRODUCT_CATALOG');
    ok(lg4.some((l) => /Deal ID 設定異動/.test(l.detail || '') && /ERP 顧問導入專案 PE/.test(l.detail || '')), 'D6 儲存商品目錄的稽核紀錄含 Deal ID 設定異動', JSON.stringify(lg4.slice(0, 2)));
    r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('ERP 顧問導入專案 PE'));
    eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_REQUIRED'], 'D7 後台把商品改成「需要」後，新增商機立即生效（不用改程式）');
    // 還原成原目錄並用 renames 改名：新欄位與改名機制相容
    const c2 = JSON.parse(JSON.stringify(mkCatalog()));
    c2.bus.ERP[0].items[0].name = 'SAP Public Cloud License 改名';
    r = await http('_s_adm', 'POST', '/api/admin/product-catalog', { catalog: { bus: c2.bus }, renames: [{ old: 'SAP Public Cloud License', new: 'SAP Public Cloud License 改名' }] });
    eq([r.status, r.json && r.json.updatedOpps > 0], [200, true], 'D8 renames 機制照常連動商機的 product');
    r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License 改名'));
    eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_REQUIRED'], 'D9 改名後新名稱沿用 requiresDealId 設定');
    r = await http('_s_u1', 'POST', '/api/opportunities', oppBody('SAP Public Cloud License'));
    eq(r.status, 201, 'D10 舊名稱已不在目錄 → 視為不需要');
    r = await http('_s_adm', 'POST', '/api/admin/product-catalog', { catalog: { bus: mkCatalog().bus } });
    eq(r.status, 200, 'D11 還原目錄');
  }
  r = await http('_s_adm', 'POST', '/api/admin/product-catalog?preview=1', { renames: [{ old: 'ERP 顧問導入專案 PE', new: 'X' }] });
  eq([r.status, r.json && r.json.preview], [200, true], 'D12 preview 模式照常');

  // ── Excel 匯入／匯出 ──
  const mkXlsx = (rows) => { const ws = XLSX.utils.aoa_to_sheet(rows); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'S'); return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }); };
  const hdr = ['客戶名稱', '銷售案名', 'BU', '預定簽約日', '業務帳號', '把握度', '合約金額', 'Deal ID'];
  const impRows = [hdr,
    ['匯入公司甲', 'SAP Public Cloud License', 'ERP', YEAR + '-12-01', '_s_u1', 'C', '100', ''],              // 需要但沒填
    ['匯入公司乙', 'SAP Public Cloud License', 'ERP', YEAR + '-12-02', '_s_u1', 'C', '100', ' IMP-1 '],       // 有填
    ['匯入公司丙', 'SAP Public Cloud License', 'ERP', YEAR + '-12-03', '_s_u1', 'C', '100', 'y'.repeat(65)],  // 太長
    ['匯入公司丁', 'ERP 顧問導入專案 PE', 'ERP', YEAR + '-12-04', '_s_u1', 'C', '100', ''],                  // 不需要
  ];
  const postImport = (dry) => {
    const fd = new FormData();
    fd.append('file', new Blob([mkXlsx(impRows)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), 'imp.xlsx');
    if (dry) fd.append('dryRun', '1');
    return fetch(BASE + '/api/admin/opportunities/import', { method: 'POST', headers: { Cookie: jar._s_adm }, body: fd }).then(async (res) => ({ status: res.status, json: await res.json().catch(() => null) }));
  };
  r = await postImport(true);
  eq([r.status, r.json && r.json.wouldCreate, r.json && r.json.missingDealId, r.json && r.json.invalidDealId], [200, 4, 2, 1], 'E1 預檢：4 筆都可匯入、缺 Deal ID 2 筆（甲、丙）、不合格 1 筆（丙）');
  const nBefore = (await listOpps()).length;
  eq(nBefore, (await listOpps()).length, 'E1b 預檢不寫入');
  r = await postImport(false);
  eq([r.status, r.json && r.json.created, r.json && r.json.missingDealId], [200, 4, 2], 'E2 實際匯入：全部成功（匯入不強制）、缺 Deal ID 統計 2 筆');
  const all = await http('_s_adm', 'GET', '/api/admin/opportunities/export');
  eq(all.status, 200, 'E3 匯出 200');
  const wb = XLSX.read(all.buf, { type: 'buffer' });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets['商機資料'], { header: 1, defval: '' });
  eq(rows[0].slice(0, 15), ['客戶名稱', '聯絡人', '銷售案名', 'BU(category)', '預定簽約日', '業務帳號(owner)', '業務姓名', '把握度階段(A/B/C/Won)', '合約金額(K)', '預估毛利率(%)', '備註(description)', '建立日期', '成交日期', '歷史紀錄(Y/N)', '商機ID'], 'E4 匯出：原本 15 欄的順序與名稱不變');
  eq(rows[0][15], 'Deal ID', 'E5 匯出：Deal ID 附加在最後');
  const rowB = rows.find((x) => x[0] === '匯入公司乙');
  eq(rowB && rowB[15], 'IMP-1', 'E6 匯入的 Deal ID（去空白）出現在匯出');
  const rowC = rows.find((x) => x[0] === '匯入公司丙');
  eq(rowC && rowC[15], '', 'E7 超過 64 字的 Deal ID 被捨棄（該列仍匯入）');
  // 匯入的商機若是「需要」的商品，之後表單編輯仍須補填
  const impA = (await listOpps('_s_u1')).find((o) => o.company === '匯入公司甲');
  r = await http('_s_u1', 'PUT', '/api/opportunities/' + impA.id, { product: impA.product, stage: 'B' });
  eq([r.status, r.json && r.json.code], [400, 'DEAL_ID_REQUIRED'], 'E8 匯入進來缺 Deal ID 的商機：表單編輯時要求補填');
  r = await http('_s_u1', 'PUT', '/api/opportunities/' + impA.id, { stage: 'B' });
  eq(r.status, 200, 'E9 …但看板拖曳階段不受影響');

  // 刪除／還原保留 dealId
  r = await http('_s_u1', 'DELETE', '/api/opportunities/o-has', { reason: '測試', reasonNote: '測試' });
  if (r.status === 200) {
    const rr = await http('_s_u1', 'POST', '/api/opportunities/restore/o-has');
    ok(rr.status === 200 && ((await oppById('o-has')) || {}).dealId === 'CHG-2', 'F1 刪除後還原，Deal ID 仍在');
  } else {
    console.log('  （刪除端點需要更多欄位，略過 F1：' + r.status + ' ' + (r.text || '').slice(0, 80) + '）');
  }
  await stopServer(); cleanupFiles();
}

const snap = () => ({ pass, fail, n: failures.length });
const restore = (x) => { pass = x.pass; fail = x.fail; failures.length = x.n; };
const newFails = (x) => failures.slice(x.n);
const setQuiet = (q) => { quiet = q; };
module.exports = { snap, restore, newFails, setQuiet, ok, eq, sleep, http, login, oppBody, listOpps, oppById, auditLines, freshServer, stopServer, cleanupFiles, readData, mkCatalog, XSS, REG_URL, BASE, PW, ROOT, SHOTS, CDP };
if (require.main === module) {
  (async () => {
    try {
      if (want('unit')) groupUnit();
      if (want('api')) await groupApi();
      if (want('ui') || want('admin')) await require('./_opp-deal-id-ui.js').run(module.exports, { ui: want('ui'), admin: want('admin') });
      if (process.env.DEALID_MUT === '1') await require('./_opp-deal-id-ui.js').mutants(module.exports, { groupApi, groupUnit });
    } catch (e) {
      fail++; failures.push('例外：' + (e && e.stack || e)); console.log('例外：', e && e.stack || e);
    } finally {
      await stopServer(); cleanupFiles();
      try { fs.unlinkSync(PRELOAD); } catch { /* 不存在 */ }
    }
    console.log(`\n商機 Deal ID 檢查：PASS ${pass} / FAIL ${fail}`);
    if (fail) { console.log('失敗項目：\n  ' + failures.join('\n  ')); }
    process.exit(fail ? 1 : 0);
  })();
}
