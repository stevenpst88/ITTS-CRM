#!/usr/bin/env node
'use strict';
/**
 * 銷售預測：可搜尋人員下拉（_client/searchable-select.js）＋匯出 Excel 帶畫面篩選（GET /api/forecast/export?owner&stages&bu）。
 *   unit  元件純邏輯（vm 載入）：比對（名稱／暱稱／帳號、不分大小寫、多詞）、排序（全部置頂）、鍵盤移動、靜態檢查（無 innerHTML、載入順序、sw 版號）
 *   api   真實 server.js（隔離沙盒）：多位 owner（重複參數／混入越權／畸形／超量）＋稽核限長；無參數＝與基準（FCS_BASE 指向 origin/main 的另一份程式樹）逐位元相同；owner／stages／bu 篩選；
 *         越權 owner 不會多出資料（多角色對拍：篩選結果必為「無篩選結果」的子集）；tecopm 403；usermap 舊格式不變／detail 格式
 *   ui    headless Edge + CDP：真實點擊／鍵盤操作下拉；選取後畫面列數＝匯出列數；深淺色／手機寬度截圖（FCS_SHOTS=資料夾）
 *
 *   node scripts/check-forecast-search-export.js            unit + api（+ ui）
 *   FCS_ONLY=unit,api,ui    只跑指定群組
 *   FCS_MUT=1               另外跑變異測試（每個變異都必須讓 api 群組變紅）
 *   FCS_BASE_REF=origin/main  基準版本（以 git archive 解到臨時資料夾）；沒設就略過「逐位元相同」比對並標示未驗證
 *   整個腳本只在「臨時沙盒」跑：以 git ls-files 把目前工作樹的程式（含未 commit 的修改與新檔）複製到系統臨時資料夾，資料／帳號檔只寫在那裡；
 *   工作樹本身的 data.json／auth.json／audit.log.json（可能正被別的伺服器使用）完全不碰、不刪、不讀。node_modules 以 NODE_PATH 指回工作樹的（不建 junction）。
 *   FCS_PORT（預設 3381，限 3380–3389）／FCS_CDP（預設 9381，限 9380–9389）
 * 不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體；全部假資料；守門：工作樹已有 auth.json／data.json／audit.log.json 時拒絕執行。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');
const { spawn } = require('child_process');
const { assertPortFree } = require('./_port-free');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.FCS_PORT || '3381', 10);
const CDP = parseInt(process.env.FCS_CDP || '9381', 10);
if (PORT < 3380 || PORT > 3389 || CDP < 9380 || CDP > 9389) { console.log('拒絕：埠需在 3380–3389 / CDP 9380–9389'); process.exit(2); }
const BASE_REF = process.env.FCS_BASE_REF || '';
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'];
const { spawnSync, execSync } = require('child_process');
const TMPS = [];
// 臨時沙盒：目前工作樹「追蹤＋未追蹤但未被忽略」的檔案（不含 scratchpad／node_modules／任何資料檔）
function mkSandbox() {
  const dst = fs.mkdtempSync(path.join(os.tmpdir(), 'fcs_run_')); TMPS.push(dst);
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
  const dst = fs.mkdtempSync(path.join(os.tmpdir(), 'fcs_base_')); TMPS.push(dst);
  execSync(`git -C "${ROOT}" archive --format=tar ${ref} | tar -x`, { cwd: dst, stdio: 'ignore', maxBuffer: 1 << 28 });   // cwd＝目的地：GNU tar 會把 "C:\…" 的冒號當遠端主機
  return dst;
}
let RUN_ROOT = '', BASE_DIR = '';
const ONLY = (process.env.FCS_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length ? true : ONLY.includes(k);
const SHOTS = process.env.FCS_SHOTS || '';
const NODE_PATH = path.join(ROOT, 'node_modules');
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
const XLSX = require(path.join(ROOT, 'node_modules', 'xlsx'));
const SS_PATH = path.join(ROOT, '_client', 'searchable-select.js');

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 400)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = (n) => new Promise((r) => setTimeout(r, n));

// ───────────────── unit ─────────────────
function loadSS(edits) {
  let src = fs.readFileSync(SS_PATH, 'utf8');
  for (const e of edits || []) { if (!src.includes(e.find)) throw new Error('MUTATION ANCHOR MISSING: ' + e.find.slice(0, 60)); src = src.replace(e.find, () => e.rep); }
  const sandbox = { module: { exports: {} }, console };
  vm.runInNewContext(src, sandbox, { filename: SS_PATH });
  return sandbox.module.exports;
}
function unitGroup(edits) {
  console.log('  [unit] 元件純邏輯');
  const SS = loadSS(edits);
  const items = [
    { value: '', label: '全部業務' },
    { value: '_s_u2', label: '測試業務乙', keywords: ['測試業務乙', ''] },
    { value: '_s_u1', label: '甲哥', keywords: ['測試業務甲', '甲哥'] },
    { value: '_s_u3', label: '丙', keywords: [] },
    { value: 'Zed_s_u9', label: 'Alice', keywords: ['Alice Chen'] },
  ];
  const vals = (q) => SS.filterItems(items, q).map((i) => i.value);
  eq(vals(''), ['', '_s_u2', '_s_u1', '_s_u3', 'Zed_s_u9'], 'U1 空查詢＝全部');
  eq(vals('甲哥'), ['', '_s_u1'], 'U2 比對暱稱（label）');
  eq(vals('測試業務甲'), ['', '_s_u1'], 'U3 比對原顯示名稱（別名，與顯示的暱稱不同）');
  eq(vals('_S_U2'), ['', '_s_u2'], 'U4 比對帳號且不分大小寫');
  eq(vals('alice'), ['', 'Zed_s_u9'], 'U5 英文不分大小寫');
  eq(vals('ALICE chen'), ['', 'Zed_s_u9'], 'U6 多詞皆須命中（AND）');
  eq(vals('alice 乙'), [''], 'U7 多詞其一未命中＝排除');
  eq(vals('找不到的人'), [''], 'U8 無結果只剩置頂「全部」');
  eq(vals('  測試  '), ['', '_s_u2', '_s_u1'], 'U9 前後空白忽略、部分字串命中');
  eq(vals('ＡＬＩＣＥ'), ['', 'Zed_s_u9'], 'U10 全形英數（NFKC）也能比對');
  const sorted = SS.sortItems([items[3], items[1], items[0], items[4], items[2]]);
  eq(sorted[0].value, '', 'U11 「全部」固定置頂');
  eq(sorted.slice(1).map((i) => i.label), sorted.slice(1).map((i) => i.label).slice().sort((a, b) => new Intl.Collator('zh-Hant', { sensitivity: 'base', numeric: true }).compare(a, b)), 'U12 其餘依顯示名稱排序');
  ok(sorted.length === items.length, 'U13 排序不增不減');
  eq([SS.moveIndex(-1, 'ArrowDown', 3), SS.moveIndex(0, 'ArrowDown', 3), SS.moveIndex(2, 'ArrowDown', 3)], [0, 1, 0], 'U14 ↓ 逐項並循環');
  eq([SS.moveIndex(0, 'ArrowUp', 3), SS.moveIndex(2, 'ArrowUp', 3), SS.moveIndex(-1, 'ArrowUp', 3)], [2, 1, 2], 'U15 ↑ 逐項並循環');
  eq([SS.moveIndex(1, 'Home', 5), SS.moveIndex(1, 'End', 5), SS.moveIndex(1, 'x', 5), SS.moveIndex(0, 'ArrowDown', 0)], [0, 4, 1, -1], 'U16 Home/End／其他鍵不動／空清單 -1');
  // 多選純邏輯
  eq(SS.toggleValue(['a'], 'b'), ['a', 'b'], 'U21 toggleValue 加入');
  eq(SS.toggleValue(['a', 'b'], 'a'), ['b'], 'U22 toggleValue 取消');
  const orig = ['x', 'y'];
  SS.toggleValue(orig, 'z'); SS.applyVisible(orig, ['y', 'q'], 'clear');
  eq(orig, ['x', 'y'], 'U23 不修改傳入陣列');
  eq(SS.applyVisible(['x'], ['a', 'b', 'x'], 'all'), ['x', 'a', 'b'], 'U24 全選目前結果＝併入可見項目、不重複');
  eq(SS.applyVisible(['x', 'a', 'b'], ['a', 'q'], 'clear'), ['x', 'b'], 'U25 清除只作用在可見結果（看不到的 x、b 保留）');
  eq(SS.applyVisible(['x'], [], 'all'), ['x'], 'U26 沒有可見結果時全選不動');
  eq(SS.applyVisible(['x', 'y'], ['x', 'y'], 'clear'), [], 'U27 清除全部可見＝清空');
  eq([SS.buttonText([], '全部業務'), SS.buttonText(['甲哥'], '全部業務'), SS.buttonText(['甲哥', '乙乙'], '全部業務')], ['全部業務', '甲哥', '甲哥、乙乙'], 'U28 按鈕文字：0／1／2 人（短名）');
  eq([SS.buttonText(['測試業務甲', '測試業務乙'], 'X'), SS.buttonText(['a', 'b', 'c'], 'X')], ['已選 2 人', '已選 3 人'], 'U29 兩個長名字或 3 人以上 → 已選 N 人（窄寬度不爆版）');
  // 面板夾邊（純函式）：視窗 390、邊距 10
  eq([SS.clampShift(116, 416, 390, 10), SS.clampShift(8, 308, 390, 10), SS.clampShift(60, 360, 390, 10), SS.clampShift(250, 550, 390, 10)], [-36, 2, 0, -170], 'U30 clampShift：右緣超出→左移；左緣不足→右移；已在內不動');
  eq(SS.clampShift(5, 500, 390, 10), 5, 'U31 面板比視窗寬時以左緣為準（不再往左推出視窗）');
  // 靜態
  const src = fs.readFileSync(SS_PATH, 'utf8');
  ok(!/innerHTML|outerHTML|insertAdjacentHTML|document\.write/.test(src), 'U17 元件不使用 innerHTML 類 sink');
  const html = fs.readFileSync(path.join(ROOT, '_client', 'index.html'), 'utf8');
  const iSS = html.indexOf('<script src="searchable-select.js">'), iApp = html.indexOf('<script src="app.js">');
  ok(iSS > 0 && iApp > 0 && iSS < iApp, 'U18 index.html 在 app.js 之前載入 searchable-select.js');
  const app = fs.readFileSync(path.join(ROOT, '_client', 'app.js'), 'utf8');
  ok(/buildForecastExportUrl/.test(app) && /params|URLSearchParams/.test(app) && /ITTSSearchSelect\.enhance/.test(app) && /multiple: true/.test(app) && /forecastSelectedOwners/.test(app), 'U19 app.js 接線存在（多選＋選取集合）');
  ok(!/\$\('forecastSalesFilter'\)\.value|\$\('forecastSalesFilter'\) \? \$\('forecastSalesFilter'\)\.value/.test(app) && !/salesEl\.value/.test(app), 'U19b app.js 不再讀 forecastSalesFilter.value（多選下語意不可信）');
  ok(/'itts-crm-v(\d+)'/.test(fs.readFileSync(path.join(ROOT, '_client', 'sw.js'), 'utf8')), 'U20 sw.js 版號格式');
}

// ───────────────── 沙盒 ─────────────────
const PW = 'Fc' + crypto.randomBytes(9).toString('hex') + '!Z';
const HASH = bcrypt.hashSync(PW, 6);
const BASEURL = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: HASH, createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canDownloadContacts: false, canSetTargets: false, accessMode: 'edit', bu: [] }, extra || {});
const ACCTS = [
  ['_s_adm', 'admin', {}],
  ['_s_m1', 'manager1', { bu: ['ERP'] }],
  ['_s_m2', 'manager2', { bu: ['ERP'] }],
  ['_s_sec', 'secretary', { bu: ['ERP'] }],
  ['_s_mkt', 'marketing', {}],
  ['_s_tp', 'tecopm', { viewOwnerScope: '_s_u1', viewGroupId: 'g1' }],
  ['_s_u1', 'user', { bu: ['ERP'], displayName: '測試業務甲', nickname: '甲哥', supervisor: '_s_m1' }],
  ['_s_u2', 'user', { bu: ['ERP'], displayName: '測試業務乙', supervisor: '_s_m1' }],
  ['_s_u3', 'user', { bu: ['ITS'], displayName: '測試業務丙' }],
  ['_s_u4', 'user', { bu: ['ERP'], displayName: '測試業務丁', supervisor: '_s_m2' }],
];
const ROLE_ACCTS = ['_s_adm', '_s_m1', '_s_m2', '_s_sec', '_s_mkt', '_s_u1', '_s_u2', '_s_u3', '_s_u4'];
const ALL_OWNERS = ['_s_u1', '_s_u2', '_s_u3', '_s_u4', '_s_m1'];
let seq = 0;
const mkOpp = (owner, stage, cat, date, amt, gm) => ({ id: 'o' + (++seq), owner, stage, category: cat, expectedDate: date, amount: amt, grossMarginRate: gm, company: '公司' + seq, product: '產品' + seq, contactId: 'c' + seq, createdAt: nowIso });
function seedData() {
  seq = 0;   // 每次播種都從 0 開始，基準樹與受測樹的假資料才會逐位元相同
  const opps = [];
  const stages = ['A', 'B', 'C', 'Won', 'D'];
  const cats = ['ERP', 'ITS', 'CRM'];
  ALL_OWNERS.forEach((ow, oi) => stages.forEach((st, si) => cats.forEach((ct, ci) => {
    opps.push(mkOpp(ow, st, ct, `2026-${String(1 + ((oi + si + ci) % 12)).padStart(2, '0')}-15`, 100 + oi * 10 + si, 20 + ci));
  })));
  opps.push(mkOpp('_s_u1', 'A', 'ERP', '2025-06-01', 50, 10));   // 他年度：不應出現
  opps.push(mkOpp('_s_u1', 'A', 'ERP', '', 50, 10));               // 無日期：不應出現
  return { contacts: [], companies: [], opportunities: opps, visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], quotations: [], groups: [] };
}
function seedFiles(dir) {
  fs.writeFileSync(path.join(dir, 'auth.json'), JSON.stringify({ users: ACCTS.map((a) => mkUser(a[0], a[1], a[2])) }, null, 2));
  fs.writeFileSync(path.join(dir, 'data.json'), JSON.stringify(seedData(), null, 2));
}
function cleanupFiles(dir) { FILES.forEach((f) => { try { fs.unlinkSync(path.join(dir, f)); } catch { /* 不存在 */ } }); }

const PRELOAD = path.join(os.tmpdir(), 'fcs_preload_' + process.pid + '.js');
fs.writeFileSync(PRELOAD, [
  "const Module=require('module'),path=require('path');const orig=Module.prototype._compile;",
  "Module.prototype._compile=function(content,filename){const raw=process.env.FCS_MUT_SPEC;",
  "if(raw){const specs=JSON.parse(raw);for(const m of specs){if(path.resolve(filename)===path.resolve(process.env.FCS_ROOT,m.file)){for(const e of m.edits){if(!content.includes(e.find))throw new Error('MUTATION ANCHOR MISSING: '+e.find.slice(0,60));content=content.replace(e.find,()=>e.rep);}}}}",
  "return orig.call(this,content,filename);};",
].join('\n') + '\n');
let child = null, MUT_CUR = null, UI_MUT = null, jar = {};
async function startServer(root, mut) {
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', NODE_PATH, ALLOWED_ORIGINS: `http://localhost:${PORT}` });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'FCS_MUT_SPEC', 'FEATURE_GUARD_MODE', 'LOGIN_LOCKOUT']) delete env[k];
  const spec = (mut || []).slice();
  spec.push({ file: 'server.js', edits: [{ find: '  max: 300,\n  standardHeaders', rep: '  max: 1e9,\n  standardHeaders' }] });
  env.FCS_MUT_SPEC = JSON.stringify(spec); env.FCS_ROOT = root;
  await assertPortFree(PORT);
  child = spawn(process.execPath, ['-r', PRELOAD, path.join(root, 'server.js')], { cwd: root, env, stdio: ['ignore', 'pipe', 'pipe'] });
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
async function freshServer(root, mut) { await stopServer(); cleanupFiles(root); seedFiles(root); await startServer(root, mut); jar = {}; }
async function login(who) {
  const res = await fetch(BASEURL + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: who, password: PW }) });
  if (res.status === 200) jar[who] = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
  return res.status;
}
async function get(who, url) {
  const res = await fetch(BASEURL + url, { headers: { Cookie: jar[who] || '' }, redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: res.status, buf, json };
}
// 匯出 → { rows: [資料列...], total: 合計列, n }
function parseXlsx(buf) {
  const wb = XLSX.read(buf, { type: 'buffer' });
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
  const rows = aoa.slice(2, -1), total = aoa[aoa.length - 1];
  return { rows, total, n: rows.length, aoa };
}
const exp = async (who, qs) => { const r = await get(who, '/api/forecast/export?year=2026' + (qs ? '&' + qs : '')); return Object.assign({ status: r.status, buf: r.buf }, r.status === 200 ? parseXlsx(r.buf) : {}); };
// 沙盒資料中某人某條件的預期列數
const label = (u) => { const a = ACCTS.find((x) => x[0] === u); return a ? ((a[2] && (a[2].nickname || a[2].displayName)) || u) : u; };
const D = seedData().opportunities.filter((o) => o.expectedDate && o.expectedDate.startsWith('2026') && o.stage !== 'D');
const count = (f) => D.filter(f).length;

async function apiGroup(tag) {
  console.log('  [api] ' + (tag || '匯出篩選'));
  await freshServer(RUN_ROOT, MUT_CUR);
  for (const w of ROLE_ACCTS.concat(['_s_tp'])) if (await login(w) !== 200) throw new Error('登入失敗 ' + w);

  // A. admin 基本篩選
  const all = await exp('_s_adm', '');
  eq(all.n, D.length, 'A1 admin 無參數：列數＝當年度非 D 且有日期的全部商機');
  const o1 = await exp('_s_adm', 'owner=_s_u1');
  eq(o1.n, count((o) => o.owner === '_s_u1'), 'A2 owner=u1 列數正確');
  ok(o1.rows.every((r) => r[4] === label('_s_u1')), 'A3 owner=u1 每列的業務人員欄都是該業務', JSON.stringify(o1.rows.map((r) => r[4])));
  const st = await exp('_s_adm', 'stages=A,Won');
  eq(st.n, count((o) => ['A', 'Won'].includes(o.stage)), 'A4 stages=A,Won 列數正確');
  ok(st.rows.every((r) => ['Commit', 'Won'].includes(r[5])), 'A5 階段欄只有 Commit／Won');
  const bu = await exp('_s_adm', 'bu=ITS');
  eq(bu.n, count((o) => o.category === 'ITS'), 'A6 bu=ITS 列數正確');
  ok(bu.rows.every((r) => r[2] === 'ITS'), 'A7 BU 欄全為 ITS');
  const combo = await exp('_s_adm', 'owner=_s_u2&stages=B&bu=CRM');
  eq(combo.n, count((o) => o.owner === '_s_u2' && o.stage === 'B' && o.category === 'CRM'), 'A8 三條件交集');
  const combo2 = await exp('_s_adm', 'owner=_s_u2&stages=B&bu=ITS');
  eq(combo2.n, count((o) => o.owner === '_s_u2' && o.stage === 'B' && o.category === 'ITS'), 'A8b 三條件交集（ITS）');
  const tot = (x) => x.total.slice(7);
  const sum = combo.rows.reduce((s, r) => s + (parseFloat(r[7]) || 0), 0);
  eq(parseFloat(combo.total[7]) || 0, sum, 'A9 合計列＝篩選後各列加總');
  ok(/合計（\d+筆）/.test(combo.total[0]) && combo.total[0] === `合計（${combo.n}筆）`, 'A10 合計列筆數與列數一致', combo.total[0]);
  const dOnly = await exp('_s_adm', 'stages=D');
  eq(dOnly.n, D.length, 'A11 stages 只含無效值（D）→ 視為未帶，仍不輸出 D');
  ok(all.rows.every((r) => r[5] !== ''), 'A11b 無參數時任何列都不是 D');
  const junk = await exp('_s_adm', 'owner=%27%20OR%201%3D1&stages=A,,%20,X&bu=%25');
  eq(junk.n, 0, 'A12 亂值 owner／bu 無匹配 → 0 列（不當萬用字元、不 500）');
  const arr = await exp('_s_adm', 'owner=_s_u1&owner=_s_u2');
  eq(arr.n, count((o) => o.owner === '_s_u1' || o.owner === '_s_u2'), 'A13 重複 owner 參數＝多位業務（聯集）');
  ok(arr.rows.every((r) => [label('_s_u1'), label('_s_u2')].includes(r[4])), 'A13b 多位 owner 每列都屬於所選業務');
  const dup = await exp('_s_adm', 'owner=_s_u1&owner=_s_u1');
  eq(dup.n, o1.n, 'A13c 重複相同值＝單一');
  const arr2 = await exp('_s_adm', 'owner[]=_s_u1&owner[]=_s_u3&stages=A');
  eq(arr2.n, count((o) => ['_s_u1', '_s_u3'].includes(o.owner) && o.stage === 'A'), 'A13d owner[]= 陣列寫法也支援，可與階段並用');

  // A2. 多位 owner：越權／畸形／超量
  const m1mix = await exp('_s_m1', 'owner=_s_u1&owner=_s_u3&owner=_s_u2');
  eq(m1mix.n, count((o) => ['_s_u1', '_s_u2'].includes(o.owner)), 'F1 manager1 混入線外 u3：只剩線內的 u1＋u2');
  const u1mix = await exp('_s_u1', 'owner=_s_u2&owner=_s_u3&owner=_s_u4');
  eq(u1mix.n, 0, 'F2 一般業務指定的全是別人 → 0 列');
  const u1self = await exp('_s_u1', 'owner=_s_u2&owner=_s_u1');
  eq(u1self.n, count((o) => o.owner === '_s_u1'), 'F3 一般業務指定自己＋別人 → 只有自己');
  const secMix = await exp('_s_sec', 'owner=_s_u3&owner=_s_u1');
  ok(secMix.status === 200 && secMix.rows.every((r) => r[4] === label('_s_u1')), 'F4 秘書（同 BU）混入他 BU 的 u3：只剩 u1');
  const bad1 = await exp('_s_adm', 'owner[a]=_s_u1');
  ok(bad1.status === 200 && bad1.n === D.length, 'F5 owner[a]=（畸形鍵）→ 當作沒帶、不 500、不超出可視範圍', bad1.status + '/' + bad1.n);
  const bad2 = await exp('_s_adm', 'owner[0][x]=1&owner=_s_u2');
  ok(bad2.status === 200 && bad2.n > 0 && bad2.rows.every((r) => r[4] === label('_s_u2')), 'F6 混入畸形鍵（owner[0][x]=）→ 畸形鍵忽略、其餘照用', bad2.status + '/' + bad2.n);
  const longName = 'x'.repeat(300);
  const lg1 = await exp('_s_adm', 'owner=' + longName);
  eq(lg1.n, 0, 'F7 超長單一值 → 0 列（不會變成不篩）');
  const lg2 = await exp('_s_adm', 'owner=' + longName + '&owner=_s_u1');
  eq(lg2.n, o1.n, 'F8 超長值混入有效值 → 有效值照用');
  const many = Array.from({ length: 200 }, (_, i) => 'zz' + i).map((v) => 'owner=' + v).join('&');
  const over = await exp('_s_adm', many + '&owner=_s_u1');
  ok(over.status === 200, 'F9 201 個 owner → 200');
  eq(over.n, 0, 'F10 第 201 個起的元素被忽略（前 200 個都是無效值 → 0 列）', over.n);
  const ok200 = await exp('_s_adm', many.split('&').slice(0, 199).join('&') + '&owner=_s_u1');
  eq(ok200.n, o1.n, 'F11 第 200 個仍有效');
  // 稽核限長
  const bigOwners = Array.from({ length: 100 }, (_, i) => 'owner=' + 'q'.repeat(100) + i).join('&');
  const big = await exp('_s_adm', bigOwners);
  ok(big.status === 200, 'F12 大量長 owner 不報錯');
  const lgs = await get('_s_adm', '/api/admin/logs');
  const arrL2 = (Array.isArray(lgs.json) ? lgs.json : (lgs.json && lgs.json.logs) || []).map((l) => JSON.stringify(l)).filter((t) => /EXPORT_FORECAST/.test(t) && /業務=q/.test(t));
  ok(arrL2.length >= 1, 'F13 有大量 owner 的匯出留下稽核紀錄');
  const seg = arrL2.length ? (/業務=([^"\\]*?)；格式/.exec(arrL2[0]) || [])[1] || '' : '';
  ok(seg.length > 0 && seg.length <= 301 && seg.endsWith('…'), 'F14 稽核 owner 描述限長（≤300 字＋「…」）', seg.length);
  const oneLong = await exp('_s_adm', 'owner=' + 'w'.repeat(100));
  const lgs2 = await get('_s_adm', '/api/admin/logs');
  const t2 = (Array.isArray(lgs2.json) ? lgs2.json : (lgs2.json && lgs2.json.logs) || []).map((l) => JSON.stringify(l)).find((t) => /EXPORT_FORECAST/.test(t) && /業務=w/.test(t)) || '';
  const seg2 = (/業務=([^"\\]*?)；格式/.exec(t2) || [])[1] || '';
  ok(seg2.length === 65 && seg2.endsWith('…'), 'F15 單一值稽核限 64 字＋「…」', seg2.length);

  // B. 越權：任何角色任何 owner，結果必為無篩選結果的子集
  const key = (r) => r.join('|');
  for (const who of ROLE_ACCTS) {
    const base = await exp(who, '');
    ok(base.status === 200, `B0 ${who} 無參數 200`, base.status);
    const baseSet = new Set(base.rows.map(key));
    const visOwners = new Set(base.rows.map((r) => r[4]));
    let leak = 0, mism = 0;
    for (const ow of ALL_OWNERS.concat(['_s_nobody'])) {
      const r = await exp(who, 'owner=' + ow);
      if (r.status !== 200) { mism++; continue; }
      if (r.rows.some((x) => !baseSet.has(key(x)))) leak++;
      if (!visOwners.has(label(ow)) && r.n !== 0) leak++;                 // 看不到的人 → 必須 0 列
      if (visOwners.has(label(ow)) && r.n !== base.rows.filter((x) => x[4] === label(ow)).length) mism++;   // 看得到的人 → 列數＝無篩選時該人的列數
    }
    ok(leak === 0, `B1 ${who}：owner 篩選不會多出無篩選看不到的資料`, 'leak=' + leak);
    ok(mism === 0, `B2 ${who}：看得到的 owner 列數與無篩選時一致`, 'mism=' + mism);
    const allBu = await exp(who, 'owner=_s_u3&stages=A,B,C,Won&bu=ITS');
    ok(allBu.rows.every((x) => baseSet.has(key(x))), `B3 ${who}：三條件並用也不超出可視範圍`);
    const multiAll = await exp(who, ALL_OWNERS.concat(['_s_nobody']).map((u) => 'owner=' + u).join('&'));
    eq(multiAll.rows.map(key).sort(), base.rows.map(key).sort(), `B3b ${who}：指定所有人＝與無篩選完全相同（不多不少）`);
  }
  const u1x = await exp('_s_u1', 'owner=_s_u2');
  eq(u1x.n, 0, 'B4 一般業務 u1 指定 u2 → 0 列');
  const m1x = await exp('_s_m1', 'owner=_s_u3');
  eq(m1x.n, 0, 'B5 manager1 指定線外的 u3 → 0 列');
  const m1own = await exp('_s_m1', 'owner=_s_u1');
  eq(m1own.n, count((o) => o.owner === '_s_u1'), 'B6 manager1 指定線內的 u1 → 該業務全部列');
  const tp = await exp('_s_tp', '');
  eq(tp.status, 403, 'B7 tecopm 無參數 403');
  const tp2 = await exp('_s_tp', 'owner=_s_u1');
  eq(tp2.status, 403, 'B8 tecopm 帶參數也 403');
  const noauth = await fetch(BASEURL + '/api/forecast/export?year=2026&owner=_s_u1', { redirect: 'manual' });
  ok(noauth.status === 401 || noauth.status === 302, 'B9 未登入不可匯出', noauth.status);

  // C. usermap
  const um = await get('_s_adm', '/api/usermap');
  ok(um.json && typeof um.json._s_u1 === 'string' && um.json._s_u1 === '甲哥', 'C1 usermap 無參數仍是 {帳號: 字串}（舊格式不變）', JSON.stringify(um.json && um.json._s_u1));
  const umd = await get('_s_adm', '/api/usermap?detail=1');
  eq(umd.json && umd.json._s_u1, { label: '甲哥', displayName: '測試業務甲', nickname: '甲哥' }, 'C2 detail=1 回 {label,displayName,nickname}');
  const um1 = await get('_s_u1', '/api/usermap?detail=1');
  eq(Object.keys(um1.json || {}), ['_s_u1'], 'C3 一般業務 detail=1 仍只有自己（可視範圍不變）');

  // D. 稽核
  const lg = await get('_s_adm', '/api/admin/logs');
  const arrL = Array.isArray(lg.json) ? lg.json : (lg.json && lg.json.logs) || [];
  const fl = arrL.filter((l) => l.action === 'EXPORT_FORECAST' || /EXPORT_FORECAST/.test(JSON.stringify(l)));
  ok(fl.some((l) => /篩選：業務=_s_u1/.test(JSON.stringify(l))), 'D1 稽核紀錄寫入篩選條件');
  ok(fl.some((l) => !/篩選：/.test(JSON.stringify(l))), 'D2 無篩選的匯出稽核文字不變');
  await stopServer();
}

// 基準逐位元：同一份假資料、同一個使用者，基準樹與修改後各跑一次，無參數的 xlsx 位元組須完全相同
async function baselineGroup() {
  console.log('  [api] 無參數輸出與基準逐位元相同');
  if (!BASE_REF) { console.log('  （未設 FCS_BASE_REF：略過，標為未驗證）'); return; }
  BASE_DIR = mkBase(BASE_REF);
  const grab = async (root) => {
    await freshServer(root);
    const out = {};
    for (const w of ['_s_adm', '_s_m1', '_s_sec', '_s_u1']) { await login(w); const r = await get(w, '/api/forecast/export?year=2026'); out[w] = r; }
    const r2 = await get('_s_adm', '/api/forecast/export?year=2026');
    out.admAgain = r2;
    if (process.env.FCS_DUMP) { fs.writeFileSync(path.join(process.env.FCS_DUMP, path.basename(root) + '_a.xlsx'), out._s_adm.buf); fs.writeFileSync(path.join(process.env.FCS_DUMP, path.basename(root) + '_b.xlsx'), r2.buf); }
    const um = await get('_s_adm', '/api/usermap'); out.um = um;
    await stopServer(); cleanupFiles(root);
    return out;
  };
  const base = await grab(BASE_DIR);
  const cur = await grab(RUN_ROOT);
  for (const w of ['_s_adm', '_s_m1', '_s_sec', '_s_u1']) {
    ok(base[w].status === 200 && cur[w].status === 200, `E0 ${w} 兩邊 200`);
    ok(Buffer.compare(base[w].buf, cur[w].buf) === 0, `E1 ${w} 無參數 xlsx 與基準逐位元相同`, `${base[w].buf.length} vs ${cur[w].buf.length}`);
  }
  ok(Buffer.compare(base.um.buf, cur.um.buf) === 0, 'E2 usermap 無參數回應與基準逐位元相同');
  ok(cur.admAgain.status === 200, 'E3 重複請求正常');
}

// ───────────────── ui ─────────────────
async function uiGroup() {
  console.log('  [ui] 真實瀏覽器操作（多選）');
  const { withBrowser } = require('./_fd-cdp');
  await freshServer(RUN_ROOT);
  ok(await login('_s_adm') === 200, 'UI0 取得匯出用工作階段');
  const J = JSON.stringify;
  const coll = new Intl.Collator('zh-Hant', { sensitivity: 'base', numeric: true });
  const shot = async (b, name) => { if (SHOTS) { fs.mkdirSync(SHOTS, { recursive: true }); await b.shot(path.join(SHOTS, name)); } };
  try {
    await withBrowser({ cdpPort: CDP, maxMs: 270000 }, async (b) => {
      if (UI_MUT) {   // 變異：頁面載入被改壞的元件（不動磁碟上的檔案）
        let src = fs.readFileSync(SS_PATH, 'utf8');
        for (const e of UI_MUT) { if (!src.includes(e.find)) throw new Error('MUTATION ANCHOR MISSING: ' + e.find.slice(0, 40)); src = src.replace(e.find, () => e.rep); }
        b.replace['/searchable-select.js'] = src; await b.interceptOn();
      }
      const open = async (who, w, h, mobile, dark) => {
        await b.viewport(w, h, mobile);
        await b.nav(BASEURL + '/login.html');
        const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(who)},password:${J(PW)}})}).then(r=>r.status)`);
        ok(st === 200, 'UI 瀏覽器登入 ' + who);
        await b.ev(`sessionStorage.setItem('justLoggedIn','1'); localStorage.setItem('darkMode',${J(dark ? '1' : '0')}); localStorage.setItem('lastSection','forecast'); 0`);
        await b.nav(BASEURL + '/');
        ok(await b.waitFor(`typeof showSection==='function' && typeof userPermissions==='object' && !!userPermissions.features && !!document.querySelector('#forecastSalesFilter')`, 15000), who + ' 載入完成');
        await b.ev(`showSection('forecast')`);
        ok(await b.waitFor(`!!document.querySelector('.ss-btn') && document.querySelectorAll('#forecastTbody tr').length>0`, 10000), who + ' 預測頁出現可搜尋下拉');
        await b.sleep(500);
      };
      const rowsShown = () => b.ev(`(()=>{const t=Array.from(document.querySelectorAll('#forecastTbody tr')).filter(x=>!x.classList.contains('ft-total-row'));return t.length===1&&t[0].children.length===1?0:t.length})()`);
      const panelOpen = () => b.ev(`!document.querySelector('.ss-panel').hidden`);
      const optsText = () => b.ev(`Array.from(document.querySelectorAll('.ss-opt')).map(e=>e.textContent)`);
      const checked = () => b.ev(`Array.from(document.querySelectorAll('.ss-opt')).filter(e=>e.getAttribute('aria-selected')==='true').map(e=>e.textContent)`);
      const btnText = () => b.ev(`document.querySelector('.ss-btn').textContent`);
      const selOwners = () => b.ev(`forecastSelectedOwners.slice().sort()`);
      const setSearch = async (t) => {
        await b.ev(`(()=>{const i=document.querySelector('.ss-input');i.focus();i.value='';i.dispatchEvent(new Event('input',{bubbles:true}));0})()`);
        if (t) await b.type(t);
        await b.sleep(120);
      };
      const clickOptText = async (txt) => {   // 真實滑鼠點擊含該文字的選項
        const idx = (await optsText()).indexOf(txt);
        ok(idx >= 0, 'UI 找得到選項 ' + txt);
        await b.click(`.ss-opt:nth-child(${idx + 1})`); await b.sleep(250);
      };
      const exportRows = async () => {
        const url = await b.ev(`buildForecastExportUrl()`);
        const r = await get('_s_adm', url.replace(/^.*\/api/, '/api'));
        return { url, n: parseXlsx(r.buf).n, status: r.status };
      };
      const quarterVisible = () => b.ev(`(()=>{const p=document.getElementById('forecastQuarterPanel');return !!p && p.style.display!=='none' && p.innerText.trim().length>0})()`);
      const quarterHidden = () => b.ev(`(()=>{const p=document.getElementById('forecastQuarterPanel');return !p || p.style.display==='none'})()`);
      const U1 = '_s_u1', U2 = '_s_u2', U3 = '_s_u3', U4 = '_s_u4';

      await open('_s_adm', 1440, 900, false, false);
      ok(await b.ev(`document.getElementById('forecastSalesFilter').style.display==='none'`), 'UI1 原 select 已隱藏（只當選項來源）');
      eq(await btnText(), '全部業務 ▾', 'UI2 未選＝「全部業務」');
      eq(await rowsShown(), D.length, 'UI3 初始畫面列數＝全部');
      ok(/\/forecast\/export\?year=\d+$/.test(await b.ev(`buildForecastExportUrl()`)), 'UI4 無篩選的匯出網址只有 year（與舊版相同）');
      ok(await quarterHidden(), 'UI4b 沒選人時季度業績面板隱藏');

      await b.click('.ss-btn'); await b.sleep(200);
      ok(await panelOpen(), 'UI5 點按鈕開啟面板');
      ok(await b.ev(`document.activeElement===document.querySelector('.ss-input')`), 'UI6 輸入框取得焦點');
      const t0 = await optsText();
      ok(!t0.includes('全部業務'), 'UI7 多選沒有「全部業務」列（不勾＝全部）', JSON.stringify(t0));
      eq(t0, t0.slice().sort((x, y) => coll.compare(x, y)), 'UI8 清單依名稱排序');
      ok(await b.ev(`document.querySelectorAll('.ss-opt .ss-chk').length===document.querySelectorAll('.ss-opt').length && document.querySelectorAll('.ss-opt').length>=10`), 'UI9 每列一個 checkbox');
      ok(await b.ev(`!!document.querySelector('.ss-all') && !!document.querySelector('.ss-clear')`), 'UI10 面板有「全選」「清除」');
      await shot(b, 'fcs_multi_empty_1440_light.png');

      // 搜尋＋全選目前結果（只作用在可見）
      await setSearch('測試業務');
      eq((await optsText()).length, 4, 'UI11 搜「測試業務」→ 4 人（含以原顯示名稱命中的暱稱者）');
      eq(await b.ev(`document.querySelector('.ss-all').textContent`), '全選目前結果', 'UI12 搜尋中按鈕字樣改為「全選目前結果」');
      await b.click('.ss-all'); await b.sleep(300);
      eq(await selOwners(), [U1, U2, U3, U4].sort(), 'UI13 全選只勾可見的 4 人');
      eq(await btnText(), '已選 4 人 ▾', 'UI14 4 人 → 「已選 N 人」');
      eq(await rowsShown(), count((o) => [U1, U2, U3, U4].includes(o.owner)), 'UI15 勾選即時重算畫面（4 位的列數）');
      ok(await quarterHidden(), 'UI16 選 4 人：季度面板隱藏');
      // 清除只作用在可見結果
      await setSearch('乙');
      await b.click('.ss-clear'); await b.sleep(300);
      eq(await selOwners(), [U1, U3, U4].sort(), 'UI17 搜「乙」後清除 → 只取消乙，其他 3 人保留');
      // 搜尋為空時清除＝全部清掉
      await setSearch('');
      await b.click('.ss-clear'); await b.sleep(300);
      eq(await selOwners(), [], 'UI18 搜尋為空時「清除」＝全部清掉');
      eq(await btnText(), '全部業務 ▾', 'UI19 清空後回到「全部業務」');
      eq(await rowsShown(), D.length, 'UI20 清空後畫面回到全部');

      // 鍵盤：↑↓、Space（搜尋框空）、Enter、Esc
      await setSearch('');
      await b.key('ArrowDown'); await b.key('ArrowDown'); await b.key(' '); await b.sleep(250);
      eq((await selOwners()).length, 1, 'UI21 搜尋框為空時 ↓↓ ＋ Space 勾選一人');
      ok(await panelOpen(), 'UI22 勾選後面板保持開啟');
      await b.key('Enter'); await b.sleep(250);
      eq((await selOwners()).length, 0, 'UI23 Enter 對同一項取消勾選（切換）');
      await b.type('甲'); await b.sleep(150);
      await b.key(' '); await b.sleep(150);
      eq(await b.ev(`document.querySelector('.ss-input').value`), '甲 ', 'UI24 搜尋框有字時 Space 照常輸入空白（不勾選）');
      eq((await selOwners()).length, 0, 'UI25 上述 Space 沒有勾選');
      await b.key('Backspace'); await b.sleep(100);
      await b.key('Enter'); await b.sleep(250);
      eq(await selOwners(), [U1], 'UI26 搜「甲」＋ Enter 勾選甲哥');
      eq(await btnText(), '甲哥 ▾', 'UI27 1 人 → 顯示名字');
      eq(await rowsShown(), count((o) => o.owner === U1), 'UI28 1 人的列數');
      ok(await quarterVisible(), 'UI29 恰好選 1 人：季度業績面板顯示且有內容');
      await b.key('Escape'); await b.sleep(150);
      ok(!(await panelOpen()), 'UI30 Esc 關閉');
      ok(await b.ev(`document.activeElement===document.querySelector('.ss-btn')`), 'UI31 Esc 後焦點回到按鈕');
      eq(await selOwners(), [U1], 'UI32 Esc 不改變已勾選');

      // 兩人（滑鼠點選）
      await b.click('.ss-btn'); await b.sleep(150);
      await setSearch('乙'); await clickOptText('測試業務乙');
      await setSearch('');
      eq(await selOwners(), [U1, U2].sort(), 'UI33 滑鼠再點一人 → 共 2 人');
      eq(await btnText(), '甲哥、測試業務乙 ▾', 'UI34 2 人且名字不長 → 列出兩個名字');
      eq(await rowsShown(), count((o) => [U1, U2].includes(o.owner)), 'UI35 2 人的列數');
      ok(await quarterHidden(), 'UI36 選 2 人：季度面板隱藏（無殘留）');
      const hint = await b.ev(`document.getElementById('forecastFilterHint').innerText`);
      ok(/甲哥/.test(hint) && /測試業務乙/.test(hint), 'UI37 篩選標籤同時顯示兩人', hint);
      ok(!(await b.ev(`document.getElementById('forecastFilterHint').innerHTML`)).includes('<script'), 'UI37b 標籤沒有未跳脫內容');
      // 加階段 B → 畫面 vs 匯出
      await b.click('#forecastStageBtn'); await b.sleep(150);
      await b.ev(`(()=>{const c=document.querySelector('.fc-stage-chk[value="B"]');c.checked=true;c.dispatchEvent(new Event('change',{bubbles:true}));0})()`);
      await b.sleep(300);
      const scr = await rowsShown();
      const ex = await exportRows();
      ok(/owner=_s_u1&owner=_s_u2|owner=_s_u2&owner=_s_u1/.test(ex.url) && /stages=B/.test(ex.url), 'UI38 匯出網址以重複參數 owner=a&owner=b 帶多位＋階段', ex.url);
      ok(!/owner=[^&]*,/.test(ex.url), 'UI38b 不使用逗號分隔 owner');
      eq(ex.n, scr, 'UI39 畫面列數＝匯出列數（2 位業務＋階段 B）');
      eq(ex.n, count((o) => [U1, U2].includes(o.owner) && o.stage === 'B'), 'UI40 列數＝預期');
      await shot(b, 'fcs_multi_picked_1440_light.png');
      // 再加 BU
      await b.ev(`(()=>{const s=document.getElementById('forecastBuFilter');s.value='ITS';s.dispatchEvent(new Event('change',{bubbles:true}));0})()`);
      await b.sleep(300);
      const ex2 = await exportRows(); const scr2 = await rowsShown();
      eq(ex2.n, scr2, 'UI41 再加 BU 後仍然一致');
      await b.ev(`(()=>{const s=document.getElementById('forecastBuFilter');s.value='';s.dispatchEvent(new Event('change',{bubbles:true}));0})()`);
      // 取消一人 → 回到恰好 1 人 → 面板再出現
      await b.click('.ss-btn'); await b.sleep(150);
      await setSearch('乙'); await clickOptText('測試業務乙');
      ok(await quarterVisible(), 'UI42 從 2 人取消 1 人 → 季度面板再次顯示');
      // 點外部／Tab 關閉
      await b.click('#forecastSummaryCards'); await b.sleep(150);
      ok(!(await panelOpen()), 'UI43 點外部關閉');
      await b.click('.ss-btn'); await b.sleep(150);
      await b.key('Tab'); await b.sleep(150);
      ok(!(await panelOpen()), 'UI44 Tab 離開並關閉');
      // 無結果
      await b.click('.ss-btn'); await b.sleep(150);
      await setSearch('zzzz不存在');
      ok(await b.ev(`!document.querySelector('.ss-empty').hidden`), 'UI45 無符合時顯示提示');
      ok(await b.ev(`document.querySelector('.ss-all').disabled && document.querySelector('.ss-clear').disabled`), 'UI46 無結果時全選／清除停用');
      const beforeEnter = await selOwners();
      await b.key('Enter'); await b.sleep(150);
      eq(await selOwners(), beforeEnter, 'UI47 無結果時 Enter 不改變選取');
      await b.key('Escape');
      // 清空回全部
      await b.click('.ss-btn'); await b.sleep(150);
      await setSearch(''); await b.click('.ss-clear'); await b.sleep(250);
      eq(await rowsShown(), count((o) => o.stage === 'B'), 'UI48 清空業務後列數恢復（階段 B 仍套用）');
      await b.key('Escape');

      // 單選模式（日後其他頁重用）：在頁面注入一個 select 實測
      await b.ev(`(()=>{const s=document.createElement('select');s.id='ssTest';[['','全部'],['a','Alpha'],['b','Beta']].forEach(([v,t])=>{const o=document.createElement('option');o.value=v;o.textContent=t;s.appendChild(o)});document.getElementById('forecastView').appendChild(s);window.__ssT=ITTSSearchSelect.enhance(s);window.__chg=0;s.addEventListener('change',()=>window.__chg++);0})()`);
      await b.ev(`window.__ssT.open(); 0`); await b.sleep(150);
      await b.type('bet'); await b.key('Enter'); await b.sleep(200);
      eq(await b.ev(`[document.getElementById('ssTest').value, window.__chg, window.__ssT.getSelected()]`), ['b', 1, ['b']], 'UI49 單選模式仍可用：搜尋＋Enter 選取、觸發 1 次 change');
      await b.ev(`document.getElementById('ssTest').parentNode.querySelectorAll('.ss-wrap').forEach(e=>e.remove()); document.getElementById('ssTest').remove(); 0`);
      ok(b.errors.length === 0, 'UI50 頁面無 JS 例外', b.errors.join(' | '));

      // 截圖：深色、淺色、手機（面板打開、已勾兩人）
      const twoPicked = async (name) => {
        await b.click('.ss-btn'); await b.sleep(200);
        await setSearch('甲'); await clickOptText('甲哥');
        await setSearch('乙'); await clickOptText('測試業務乙');
        await setSearch('測試');
        await shot(b, name);
      };
      await open('_s_adm', 1440, 900, false, false); await twoPicked('fcs_multi_1440_light.png');
      await open('_s_adm', 1440, 900, false, true); await twoPicked('fcs_multi_1440_dark.png');
      await open('_s_adm', 390, 800, true, false); await twoPicked('fcs_multi_390_light.png');
      ok(await b.ev(`(()=>{const r=document.querySelector('.ss-panel').getBoundingClientRect();return r.left>=0&&r.right<=window.innerWidth})()`), 'UI51 手機寬度面板不超出視窗');
      ok(await b.ev(`(()=>{const q=s=>document.querySelector(s).getBoundingClientRect().height;return q('.ss-opt')>=40 && q('.ss-tool')>=36})()`), 'UI52 手機觸控目標夠大（選項 ≥40px、按鈕 ≥36px）');
      ok(await b.ev(`parseFloat(getComputedStyle(document.querySelector('.ss-input')).fontSize)>=16`), 'UI53 手機輸入框字級 ≥16px');
      ok(await b.ev(`document.documentElement.scrollWidth<=window.innerWidth+1`), 'UI54 手機寬度頁面沒有橫向捲動');
      // 面板夾邊：按鈕靠左／中間／靠右，面板完整落在視窗內（多選）；再轉成 360 寬重新開啟重算
      const rectIn = () => b.ev(`(()=>{const r=document.querySelector('.ss-panel').getBoundingClientRect();const w=document.documentElement.clientWidth;return {l:Math.round(r.left),r:Math.round(r.right),w}})()`);
      for (const [nm, ml] of [['靠左', 0], ['中間', 90], ['靠右', 200]]) {
        await b.ev(`document.querySelector('.ss-wrap').style.marginLeft='${ml}px'; 0`);
        await b.click('.ss-btn'); await b.sleep(200);
        const rr = await rectIn(), bx = await b.ev(`Math.round(document.querySelector('.ss-btn').getBoundingClientRect().left)`);
        ok(rr.l >= 0 && rr.r <= rr.w, `UI57 390 寬、按鈕${nm}（left≈${bx}）：面板完全在視窗內`, JSON.stringify(rr));
        await b.key('Escape'); await b.sleep(120);
      }
      await b.ev(`document.querySelector('.ss-wrap').style.marginLeft='200px'; 0`);
      await b.viewport(360, 800, true); await b.sleep(300);
      await b.click('.ss-btn'); await b.sleep(200);
      { const rr = await rectIn(); ok(rr.l >= 0 && rr.r <= rr.w, 'UI58 視窗縮成 360 後重新開啟：重算且仍在視窗內', JSON.stringify(rr)); }
      await b.key('Escape'); await b.sleep(120);
      await b.viewport(390, 800, true); await b.sleep(300);
      // 單選模式也受益（頁面注入 select，緊貼右側）
      await b.ev(`(()=>{const s=document.createElement('select');s.id='ssTest2';[['','全部'],['a','Alpha']].forEach(([v,t])=>{const o=document.createElement('option');o.value=v;o.textContent=t;s.appendChild(o)});const h=document.createElement('div');h.id='ssHost';h.style.cssText='display:flex;justify-content:flex-end;padding:0 4px';h.appendChild(s);document.getElementById('forecastView').insertBefore(h,document.getElementById('forecastView').firstChild);window.__ss2=ITTSSearchSelect.enhance(s);0})()`);
      await b.ev(`window.__ss2.open(); 0`); await b.sleep(200);
      { const rr = await b.ev(`(()=>{const r=window.__ss2.panel.getBoundingClientRect();return {l:Math.round(r.left),r:Math.round(r.right),w:document.documentElement.clientWidth}})()`); ok(rr.l >= 0 && rr.r <= rr.w, 'UI59 單選模式、按鈕貼右：面板完全在視窗內', JSON.stringify(rr)); }
      await b.ev(`document.getElementById('ssHost').remove(); 0`);
      // 截圖：按鈕靠右＋面板打開（淺／深）
      await open('_s_adm', 390, 800, true, false);
      await b.ev(`document.querySelector('.ss-wrap').style.marginLeft='200px'; 0`);
      await b.click('.ss-btn'); await b.sleep(250); await setSearch('測試'); await shot(b, 'fcs_clamp_390_right_light.png');
      await open('_s_adm', 390, 800, true, true);
      await b.ev(`document.querySelector('.ss-wrap').style.marginLeft='200px'; 0`);
      await b.click('.ss-btn'); await b.sleep(250); await setSearch('測試'); await shot(b, 'fcs_clamp_390_right_dark.png');
      await open('_s_adm', 390, 800, true, true); await twoPicked('fcs_multi_390_dark.png');
      // 一般業務：不建立下拉；匯出網址無 owner
      await b.viewport(1440, 900, false);
      await b.nav(BASEURL + '/login.html');
      await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:'_s_u1',password:${J(PW)}})}).then(r=>r.status)`);
      await b.ev(`sessionStorage.setItem('justLoggedIn','1'); localStorage.setItem('lastSection','forecast'); 0`);
      await b.nav(BASEURL + '/');
      await b.waitFor(`typeof showSection==='function' && typeof userPermissions==='object' && !!userPermissions.features`, 15000);
      await b.ev(`showSection('forecast')`); await b.sleep(1200);
      eq(await b.ev(`document.querySelectorAll('.ss-wrap').length`), 0, 'UI55 一般業務不建立可搜尋下拉（canFilter 規則不變）');
      ok(/\/forecast\/export\?year=\d+$/.test(await b.ev(`buildForecastExportUrl()`)), 'UI56 一般業務匯出網址不帶 owner');
    });
  } finally { await stopServer(); cleanupFiles(RUN_ROOT); }
}

// ───────────────── 變異測試 ─────────────────
const MUTS = [
  { name: 'M1 拿掉可視範圍縮小（owner 有值時略過 forecastOwners）', file: 'server.js', find: "forecastOwners.includes(o.owner) && o.expectedDate && new Date(o.expectedDate).getFullYear() === yr && o.stage !== 'D')\n    .filter(", rep: "(fOwners.size ? true : forecastOwners.includes(o.owner)) && o.expectedDate && new Date(o.expectedDate).getFullYear() === yr && o.stage !== 'D')\n    .filter(" },
  { name: 'M2 拿掉 owner 篩選', file: 'server.js', find: '(!fOwners.size || fOwners.has(o.owner)) && ', rep: '' },
  { name: 'M6 拿掉 owner 200 個上限', file: 'server.js', find: '_ownerRaw.slice(0, 200)', rep: '_ownerRaw' },
  { name: 'M7 不支援 owner[] 寫法', file: 'server.js', find: ".concat(sp.getAll('owner[]'))", rep: '' },
  { name: 'M8 拿掉稽核 owner 總長限制', file: 'server.js', find: "if (_ownDesc.length > 300) _ownDesc = _ownDesc.slice(0, 300) + '…';", rep: '' },
  { name: 'M9 owner 多值只取第一個', file: 'server.js', find: '_ownerRaw.slice(0, 200)', rep: '_ownerRaw.slice(0, 1)' },
  { name: 'M10 owner 單值不截斷至 128 且超長值被丟棄（變成不篩）', file: 'server.js', find: '.map(v => v.trim().slice(0, 128)).filter(Boolean)', rep: '.map(v => v.trim()).filter(v => v && v.length <= 128)' },
  { name: 'M3 拿掉階段篩選', file: 'server.js', find: '(!fStages.length || fStages.includes(o.stage)) && ', rep: '' },
  { name: 'M4 拿掉 BU 篩選', file: 'server.js', find: ' && (!fBu || o.category === fBu))', rep: ')' },
  { name: 'M5 階段白名單放寬成含 D', file: 'server.js', find: "const FORECAST_FILTER_STAGES = ['A', 'B', 'C', 'Won'];", rep: "const FORECAST_FILTER_STAGES = ['A', 'B', 'C', 'Won', 'D'];" },
];
async function mutGroup() {
  console.log('  [mut] 變異測試（每個變異 api 群組都必須變紅）');
  // 元件變異
  {
    const p0 = pass, f0 = fail; quiet = true;
    unitGroup([{ find: "return terms.every(t => hay.some(h => h.indexOf(t) !== -1));", rep: 'return true;' }]);
    quiet = false;
    const red = fail - f0 > 0; pass = p0; fail = f0;
    ok(red, '變異 MU1 比對永遠為真 → unit 變紅');
  }
  {
    const p0 = pass, f0 = fail; quiet = true;
    unitGroup([{ find: "if (mode === 'clear') return selected.filter(v => !vis.has(v));", rep: "if (mode === 'clear') return [];" }]);
    quiet = false;
    const red = fail - f0 > 0; pass = p0; fail = f0;
    ok(red, '變異 MU2 「清除」清掉看不到的項目 → unit 變紅');
  }
  {
    const p0 = pass, f0 = fail; quiet = true;
    unitGroup([{ find: 'if (right > vw - margin) dx = (vw - margin) - right;', rep: '' }]);
    quiet = false;
    const red = fail - f0 > 0; pass = p0; fail = f0;
    ok(red, '變異 MU3 拿掉夾邊（clampShift 右緣不處理）→ unit 變紅');
  }
  // UI 變異：把頁面載入的元件換成「open() 不呼叫 position()」的版本，三種位置的 UI 測試必須變紅
  {
    const p0 = pass, f0 = fail, n0 = failures.length; quiet = true; let err = '';
    try { UI_MUT = [{ find: '      render();\n      position();\n      input.focus();', rep: '      render();\n      input.focus();' }]; await uiGroup(); } catch (e) { err = e.message; await stopServer(); }
    UI_MUT = null; quiet = false;
    const red = fail - f0 > 0 || err; pass = p0; fail = f0; failures.length = n0;
    ok(red && !/ANCHOR MISSING/.test(err), '變異 MU4 開啟時不夾邊 → ui（夾邊測試）變紅', err || '沒有任何測試變紅');
  }
  for (const m of MUTS) {
    const p0 = pass, f0 = fail, n0 = failures.length; quiet = true;
    let err = '';
    try { MUT_CUR = [{ file: m.file, edits: [{ find: m.find, rep: m.rep }] }]; await apiGroup('變異 ' + m.name); } catch (e) { err = e.message; await stopServer(); }
    MUT_CUR = null; quiet = false;
    const red = fail - f0 > 0 || err;
    pass = p0; fail = f0; failures.length = n0;
    ok(red && !/ANCHOR MISSING/.test(err), m.name + ' → 變紅', err || '沒有任何測試變紅');
  }
}

(async () => {
  const t0 = Date.now();
  try {
    RUN_ROOT = mkSandbox();
    if (want('unit')) unitGroup();
    if (want('api')) { await apiGroup(); await baselineGroup(); }
    if (want('ui')) await uiGroup();
    if (process.env.FCS_MUT === '1') await mutGroup();
  } catch (e) { fail++; failures.push('例外：' + e.message); console.log('例外：', e.stack); }
  finally {
    await stopServer();
    try { fs.unlinkSync(PRELOAD); } catch { /* 略 */ }
    for (const d of TMPS) { try { fs.rmSync(d, { recursive: true, force: true }); } catch { /* 略 */ } }   // 臨時資料夾內沒有 junction／連結
  }
  console.log(`\n結果：${pass} 通過、${fail} 失敗（${Math.round((Date.now() - t0) / 1000)} 秒）`);
  if (fail) { console.log('失敗項：', failures.join(' ／ ')); process.exit(1); }
})();
