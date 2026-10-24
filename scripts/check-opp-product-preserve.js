#!/usr/bin/env node
'use strict';
/**
 * 商機編輯表單「儲存不可把舊商機的商品清成空字串」檢查：API（真實 server.js）＋前端互動（headless Edge + CDP，真實滑鼠／鍵盤）＋變異測試。
 *
 * 背景：商機編輯視窗的商品下拉選不到原值時（類別是空的、類別在目錄裡沒有商品、商品已改名／停用且清單沒有「其他」），
 * 舊版儲存會把 product 送成 ''，後端 PUT 照單全收 → 舊商機的商品被清空。修法在前端（_client/app.js 的 _oppEditOrig）：
 * 使用者沒動過商品欄、目前值為空 → 不送 product；後端語意不變（仍接受空字串，這支腳本也驗這點）。
 *
 * 不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體；資料檔只存在於本次執行；結束以 PID 清掉伺服器與 Edge；不連遠端資料庫。
 *
 *   node scripts/check-opp-product-preserve.js           API＋前端
 *   OPPPRES_MUT=1 node scripts/check-opp-product-preserve.js   另外跑變異測試（每個變異都必須讓前端群組變紅）
 *   OPPPRES_ONLY=api,ui                                  只跑指定群組
 *   OPPPRES_PORT（預設 3340）／OPPPRES_CDP（預設 9640）
 * 守門：工作樹已有 auth.json／data.json／audit.log.json（可能是真實資料）時拒絕執行。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');
const { assertPortFree } = require('./_port-free');
const { withBrowser } = require('./_fd-cdp');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.OPPPRES_PORT || '3340', 10);
const CDP = parseInt(process.env.OPPPRES_CDP || '9640', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3161)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map((f) => path.join(ROOT, f));
{
  const g = FILES.filter((f) => /(auth|data|audit.log)\.jsonl?$/.test(f)).filter((f) => fs.existsSync(f));
  if (g.length) { console.log('拒絕：此工作樹已有', g.map((f) => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
}
const ONLY = (process.env.OPPPRES_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length || ONLY.includes(k);
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0, quiet = false;
const failures = [];
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name); if (!quiet) console.log('  ✗ FAIL:', name, extra === undefined ? '' : '— ' + String(extra).slice(0, 300)); } }
const eq = (a, b, name) => ok(JSON.stringify(a) === JSON.stringify(b), name, `得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}`);
const sleep = (n) => new Promise((r) => setTimeout(r, n));
const J = JSON.stringify;

// ───────────────── 沙盒資料 ─────────────────
const PW = 'Pp' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const YEAR = new Date().getFullYear();
const REG_URL = 'https://deal.example.test/register?src=crm';
const XSS = '<img src=x onerror="window.__xss=1">';
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
const mkCatalog = () => ({
  bus: {
    ERP: [{ group: '一般商品', enabled: true, items: [
      { name: 'SAP Public Cloud License', enabled: true, subs: [], requiresDealId: true, dealRegisterUrl: REG_URL },
      { name: 'SAP CCFLEX', enabled: true, subs: ['C&S', 'A&O'], requiresDealId: true },
      { name: 'ERP 顧問導入專案 PE', enabled: true, subs: [] },
      { name: '其他', enabled: true, subs: [] },
    ] }],
    ITS: [{ group: 'ITS 產品', enabled: true, items: [        // 注意：ITS 沒有「其他」
      { name: 'MES', enabled: true, subs: [] },
      { name: 'WMS', enabled: true, subs: [] },
    ] }],
    MDM: [], CRM: [], '資安': [],                                // 沒有商品清單 → 商品群組整個隱藏
  },
  updatedAt: nowIso, updatedBy: 'seed',
});
const mkOpp = (id, category, product, extra) => Object.assign({ id, owner: '_s_u1', bu: 'ERP', company: '沙盒公司A', contactId: '', contactName: '', category, product, amount: '100', expectedDate: YEAR + '-12-31', description: '', stage: 'C', visitId: '', createdAt: nowIso }, extra || {});
// id → 情境
const SEED_OPPS = [
  mkOpp('o-nocat', '', 'Legacy Product X'),                                   // 無類別（API／匯入／Lead 轉換建的舊商機）
  mkOpp('o-nocat2', '', 'Legacy Product Y'),                                  // 無類別：使用者補類別 ERP
  mkOpp('o-nocat3', '', 'ERP 顧問導入專案 PE'),                              // 無類別但商品剛好在 ERP 清單
  mkOpp('o-nocat-req', '', 'SAP Public Cloud License'),                       // 無類別、商品需要 Deal ID、尚無 Deal ID
  mkOpp('o-nocat-drag', '', 'Legacy Product D'),                              // 看板拖曳
  mkOpp('o-mdm', 'MDM', 'MDM 專案Z'),                                         // 類別沒有商品清單
  mkOpp('o-its-gone', 'ITS', '已下架商品'),                                   // 類別有清單，但商品不在且沒有「其他」
  mkOpp('o-its-gone2', 'ITS', '已下架商品2'),                                 // 同上：使用者改選 MES
  mkOpp('o-its-gone3', 'ITS', '已下架商品3'),                                 // 同上：使用者明確選回「-- 請選擇 --」
  mkOpp('o-erp-custom', 'ERP', '舊自訂案名'),                                 // ERP 有「其他」→ 走自訂欄預填
  mkOpp('o-normal', 'ERP', 'ERP 顧問導入專案 PE'),                           // 一般商機
  mkOpp('o-normal2', 'ERP', 'ERP 顧問導入專案 PE'),                          // 使用者明確清空
  mkOpp('o-normal3', 'ERP', 'ERP 顧問導入專案 PE'),                          // 改類別 ERP→ITS 再改回
  mkOpp('o-normal4', 'ERP', 'ERP 顧問導入專案 PE'),                          // 改類別 ERP→ITS 後存檔
  mkOpp('o-dealreq', 'ERP', 'SAP Public Cloud License', { dealId: 'D-EXIST' }),   // Deal ID 規則商品
  mkOpp('o-xss-its', 'ITS', XSS),                                             // XSS：有類別、商品不在清單
  mkOpp('o-xss-nocat', '', XSS),                                              // XSS：無類別
];
function seedFiles() {
  const users = [mkUser('_s_adm', 'admin'), mkUser('_s_u1', 'user', { bu: ['ERP'] })];
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  const d = {
    contacts: [], companies: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], quotations: [],
    opportunities: JSON.parse(JSON.stringify(SEED_OPPS)),
    productCatalog: mkCatalog(),
  };
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
}
function cleanupFiles() { FILES.forEach((f) => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } }); }

// ───────────────── 伺服器 ─────────────────
let child = null, childPid = null;
async function startServer() {
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'VERCEL', 'AUDIT_EVENTS_TABLE', 'QUOTE_FLOW_ENGINE', 'QUOTE_FLOW_DESIGNER', 'FEATURE_GUARD_MODE']) delete env[k];
  await assertPortFree(PORT);
  child = spawn(process.execPath, [path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
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
async function freshServer() { await stopServer(); cleanupFiles(); seedFiles(); await startServer(); jar = {}; }

// ───────────────── HTTP ─────────────────
let jar = {};
async function http(who, method, url, body) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const text = await res.text();
  let json = null; try { json = JSON.parse(text); } catch { /* 非 JSON */ }
  return { status: res.status, json, text };
}
async function login(who) {
  const res = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: J({ username: who, password: PW }) });
  if (res.status === 200) jar[who] = (res.headers.getSetCookie ? res.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
  return res.status;
}
const listOpps = async () => (await http('_s_u1', 'GET', '/api/opportunities')).json || [];
const oppById = async (id) => (await listOpps()).find((o) => o.id === id);

// ───────────────── API 群組 ─────────────────
async function groupApi() {
  console.log('  [API] 後端語意不變＋API 建立無類別商機');
  await freshServer();
  eq(await login('_s_u1'), 200, 'A0 登入沙盒帳號');
  // 用 API 直接 POST 建立「沒有 category」的商機（舊商機的真實來源之一）
  let r = await http('_s_u1', 'POST', '/api/opportunities', { company: '沙盒公司A', product: 'API 建的商品', amount: '50', expectedDate: YEAR + '-11-30', description: 'api' });
  eq([r.status, r.json && r.json.category, r.json && r.json.product], [201, '', 'API 建的商品'], 'A1 API 可建立沒有 category 的商機（category 存成空字串）');
  const apiId = r.json && r.json.id;
  // 後端語意不變：PUT 帶 product:'' 仍會清空（刻意不改 API 行為；修正在前端不送）
  r = await http('_s_u1', 'PUT', '/api/opportunities/' + apiId, { product: '', stage: 'B' });
  eq([r.status, r.json && r.json.product], [200, ''], 'A2 後端 PUT 帶空字串 product 仍照單全收清空（API 語意未改）');
  // 偏更新（payload 不含 product）不動 product——前端修正後送出的形狀
  r = await http('_s_u1', 'POST', '/api/opportunities', { company: '沙盒公司A', product: 'API 建的商品2', amount: '50', expectedDate: YEAR + '-11-30' });
  const id2 = r.json.id;
  r = await http('_s_u1', 'PUT', '/api/opportunities/' + id2, { stage: 'B', amount: '77', expectedDate: YEAR + '-12-01', description: 'x', bu: 'ERP', businessType: '' });
  eq([r.status, r.json && r.json.product, r.json && r.json.stage], [200, 'API 建的商品2', 'B'], 'A3 PUT 不含 product 鍵：product 保持原值（前端修正後的 payload 形狀）');
  await stopServer(); cleanupFiles();
}

// ───────────────── 前端群組 ─────────────────
const DISPLAY = (id) => `(()=>{const e=document.getElementById(${J(id)});return !!e && !e.hidden && getComputedStyle(e).display!=='none'})()`;
const selSet = (b, id, v) => b.ev(`(()=>{const e=document.getElementById(${J(id)}); if(!e) return false; e.value=${J(v)}; e.dispatchEvent(new Event('change',{bubbles:true})); return e.value===${J(v)}})()`);
const reqs = (b, re) => b.ev(`window.__reqs.filter(r=>${re}.test(r.u)).map(r=>({m:r.m,b:r.b}))`);
const clearReqs = (b) => b.ev(`window.__reqs.length=0; 0`);
const OVERLAY_CLOSED = `!document.getElementById('oppEditOverlay').classList.contains('open')`;

async function loginBrowser(b, who) {
  await b.nav(BASE + '/login.html');
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${J(who)},password:${J(PW)}})}).then(r=>r.status)`);
  ok(st === 200, '瀏覽器登入 ' + who);
}
async function openApp(b) {
  await loginBrowser(b, '_s_u1');
  await b.nav(BASE + '/');
  const ready = await b.waitFor('typeof openOppEdit==="function" && typeof loadOpportunities==="function" && typeof OPP_PRODUCTS!=="undefined" && !!OPP_PRODUCTS.ITS && Object.keys(DEAL_ID_RULES).length>0', 25000);
  ok(ready, '使用者端載入完成且商品目錄已就緒（含沙盒目錄）');
  await b.ev(`(async()=>{await loadOpportunities(); return allOpportunities.length})()`);
  // 記錄所有非 GET 請求
  await b.ev(`(()=>{ if(window.__reqs) return; window.__reqs=[]; const of=window.fetch; window.fetch=function(u,o){ if(o&&o.method&&o.method!=='GET') window.__reqs.push({u:String(u),m:o.method,b:typeof o.body==='string'?o.body:''}); return of.apply(this,arguments); }; })()`);
}
// 開啟編輯 → 回傳目前下拉狀態
async function openEdit(b, id) {
  await clearReqs(b);
  await b.ev(`openOppEdit(${J(id)}); 0`);
  await b.sleep(450);   // 等開啟動畫結束再操作
  return b.ev(`(()=>{const g=id=>document.getElementById(id); const ps=g('oppEditProduct'); return {open:g('oppEditOverlay').classList.contains('open'), groupShown:getComputedStyle(g('oppEditProductGroup')).display!=='none', value:ps.value, selText:ps.selectedOptions[0]?ps.selectedOptions[0].textContent:null, optCount:ps.options.length, custom:g('oppEditProductCustom').value, customShown:getComputedStyle(g('oppEditProductCustom')).display!=='none', category:g('oppEditCategory').value}})()`);
}
// 真實點擊儲存；回傳這次 PUT 的 body（物件）
async function saveClick(b, id) {
  await clearReqs(b);
  await b.click('#oppEditSave');
  const closed = await b.waitFor(OVERLAY_CLOSED, 8000);
  const put = (await reqs(b, new RegExp('\\/api\\/opportunities\\/' + id + '$'))).find((x) => x.m === 'PUT');
  return { closed, body: put ? JSON.parse(put.b) : null };
}
const server = async (id) => (await oppById(id)) || {};

async function uiOnce() {
  await freshServer();
  eq(await login('_s_u1'), 200, 'U0 登入沙盒帳號');
  const mutNote = process.env.OPPPRES_MUTNOTE || '';
  void mutNote;
  await withBrowser({ cdpPort: CDP, maxMs: 420000 }, async (b) => {
    if (uiOnce.mutate) { await b.interceptOn(); b.replace['/app.js'] = uiOnce.mutate; }
    await openApp(b);

    // ── 1. 無類別舊商機：改階段儲存，商品不可被清掉 ──
    console.log('  [UI] 1 無類別舊商機');
    let st = await openEdit(b, 'o-nocat');
    eq([st.open, st.category, st.value, st.groupShown], [true, '', '', false], 'P1 無類別：下拉只剩「請選擇」且商品群組隱藏（這就是原本會清空的情境）');
    await selSet(b, 'oppEditStage', 'B');
    let sv = await saveClick(b, 'o-nocat');
    ok(sv.closed, 'P2 儲存成功並關閉視窗');
    ok(sv.body && !('product' in sv.body), 'P3 PUT payload 不含 product 鍵', J(sv.body));
    eq(await server('o-nocat').then((o) => [o.product, o.stage, o.category]), ['Legacy Product X', 'B', ''], 'P4 伺服器：商品保持原值、階段已更新');

    // ── 2. 無類別，使用者補上類別 ERP：商品依清單帶出（不在清單 → 走「其他」自訂欄預填）──
    console.log('  [UI] 2 補類別');
    st = await openEdit(b, 'o-nocat2');
    await selSet(b, 'oppEditCategory', 'ERP');
    st = await b.ev(`(()=>{const g=id=>document.getElementById(id); return {value:g('oppEditProduct').value, custom:g('oppEditProductCustom').value, shown:getComputedStyle(g('oppEditProductCustom')).display!=='none', group:getComputedStyle(g('oppEditProductGroup')).display!=='none'}})()`);
    eq([st.value, st.custom, st.shown, st.group], ['其他', 'Legacy Product Y', true, true], 'P5 無類別商機補選 ERP：原商品以「其他」＋自訂欄預填顯示（看得到、不會悄悄丟掉）');
    sv = await saveClick(b, 'o-nocat2');
    eq(await server('o-nocat2').then((o) => [o.product, o.category]), ['Legacy Product Y', 'ERP'], 'P6 儲存後：類別 ERP、商品保持原值');

    st = await openEdit(b, 'o-nocat3');
    await selSet(b, 'oppEditCategory', 'ERP');
    eq(await b.ev(`document.getElementById('oppEditProduct').value`), 'ERP 顧問導入專案 PE', 'P7 無類別商機補選 ERP：商品剛好在清單 → 直接選中');
    await b.ev(`closeOppEdit(); 0`);

    // ── 3. 類別沒有商品清單（MDM）──
    console.log('  [UI] 3 類別沒有商品清單');
    st = await openEdit(b, 'o-mdm');
    eq([st.category, st.groupShown, st.value], ['MDM', false, ''], 'P8 MDM：商品群組隱藏（沒有清單可選）');
    await selSet(b, 'oppEditStage', 'B');
    sv = await saveClick(b, 'o-mdm');
    ok(sv.body && !('product' in sv.body), 'P9 PUT 不含 product', J(sv.body));
    eq(await server('o-mdm').then((o) => [o.product, o.category, o.stage]), ['MDM 專案Z', 'MDM', 'B'], 'P10 伺服器：商品與類別保持原值');

    // ── 4. 類別有清單但商品不在且沒有「其他」（ITS）──
    console.log('  [UI] 4 商品不在清單、無「其他」');
    st = await openEdit(b, 'o-its-gone');
    eq([st.groupShown, st.value, st.selText], [true, '已下架商品', '（原值）已下架商品'], 'P11 ITS：補「（原值）」選項並選中，使用者看得到原商品');
    await selSet(b, 'oppEditStage', 'A');
    sv = await saveClick(b, 'o-its-gone');
    eq(sv.body && sv.body.product, '已下架商品', 'P12 payload 帶回原商品');
    eq(await server('o-its-gone').then((o) => [o.product, o.stage]), ['已下架商品', 'A'], 'P13 伺服器：商品保持原值');

    st = await openEdit(b, 'o-its-gone2');
    await selSet(b, 'oppEditProduct', 'MES');
    sv = await saveClick(b, 'o-its-gone2');
    eq(await server('o-its-gone2').then((o) => o.product), 'MES', 'P14 使用者主動改選 MES：商品確實被改掉（改商品功能不受影響）');

    st = await openEdit(b, 'o-its-gone3');
    await selSet(b, 'oppEditProduct', '');
    sv = await saveClick(b, 'o-its-gone3');
    eq([sv.body && sv.body.product, await server('o-its-gone3').then((o) => o.product)], ['', ''], 'P15 使用者明確選回「-- 請選擇 --」：照原行為清空（那是他的操作）');

    // ── 5. ERP 有「其他」→ 自訂欄預填（原行為）──
    console.log('  [UI] 5 ERP 自訂案名');
    st = await openEdit(b, 'o-erp-custom');
    eq([st.value, st.custom, st.customShown], ['其他', '舊自訂案名', true], 'P16 ERP 商品不在清單：選「其他」並預填自訂欄（原行為不變）');
    await selSet(b, 'oppEditStage', 'B');
    sv = await saveClick(b, 'o-erp-custom');
    eq(await server('o-erp-custom').then((o) => o.product), '舊自訂案名', 'P17 儲存後商品保持原值');

    // ── 6. 一般商機行為不變 ──
    console.log('  [UI] 6 一般商機');
    st = await openEdit(b, 'o-normal');
    eq([st.value, st.groupShown], ['ERP 顧問導入專案 PE', true], 'P18 一般商機：下拉選中原商品');
    await selSet(b, 'oppEditProduct', 'SAP CCFLEX');
    await selSet(b, 'oppEditProductSub', 'C&S');
    eq(await b.ev(DISPLAY('oppEditDealIdGroup')), true, 'P19 改選 SAP CCFLEX C&S：Deal ID 欄位出現（連動正常）');
    await b.click('#oppEditDealId');
    await b.type('NEW-DEAL-1');
    sv = await saveClick(b, 'o-normal');
    eq(await server('o-normal').then((o) => [o.product, o.dealId]), ['SAP CCFLEX C&S', 'NEW-DEAL-1'], 'P20 一般商機改商品＋補 Deal ID：存成「SAP CCFLEX C&S」');

    st = await openEdit(b, 'o-normal2');
    await selSet(b, 'oppEditProduct', '');
    sv = await saveClick(b, 'o-normal2');
    eq([sv.body && sv.body.product, await server('o-normal2').then((o) => o.product)], ['', ''], 'P21 一般商機明確選回「請選擇」：照舊清空（原行為不變）');

    // 改類別 ERP→ITS 再改回 → 還原原商品
    st = await openEdit(b, 'o-normal3');
    await selSet(b, 'oppEditCategory', 'ITS');
    eq(await b.ev(`document.getElementById('oppEditProduct').value`), '', 'P22 改類別 ERP→ITS：商品重置為「請選擇」（原行為）');
    await selSet(b, 'oppEditCategory', 'ERP');
    eq(await b.ev(`document.getElementById('oppEditProduct').value`), 'ERP 顧問導入專案 PE', 'P23 改回原類別 ERP：還原原商品');
    sv = await saveClick(b, 'o-normal3');
    eq(await server('o-normal3').then((o) => [o.category, o.product]), ['ERP', 'ERP 顧問導入專案 PE'], 'P24 儲存：類別與商品都保持原值');

    // 改類別 ERP→ITS 後存檔：商品依新類別重選（空）→ 清空是使用者改類別的結果（原行為）
    st = await openEdit(b, 'o-normal4');
    await selSet(b, 'oppEditCategory', 'ITS');
    sv = await saveClick(b, 'o-normal4');
    eq(await server('o-normal4').then((o) => [o.category, o.product]), ['ITS', ''], 'P25 使用者把 ERP 改成 ITS 且沒選商品：商品為空（原行為不變，不偷偷保留跨類別的舊商品）');

    // ── 7. Deal ID 規則 ──
    console.log('  [UI] 7 Deal ID 規則');
    st = await openEdit(b, 'o-dealreq');
    eq([st.value, await b.ev(`document.getElementById('oppEditDealId').value`)], ['SAP Public Cloud License', 'D-EXIST'], 'P26 Deal ID 商品：下拉與 Deal ID 帶出既有值');
    await selSet(b, 'oppEditStage', 'B');
    sv = await saveClick(b, 'o-dealreq');
    eq(await server('o-dealreq').then((o) => [o.product, o.dealId, o.stage]), ['SAP Public Cloud License', 'D-EXIST', 'B'], 'P27 儲存後 product／dealId 不變');

    st = await openEdit(b, 'o-nocat-req');
    eq(await b.ev(DISPLAY('oppEditDealIdGroup')), true, 'P28 無類別但商品（原值）需要 Deal ID：欄位照樣顯示（規則看原商品，不因下拉空白而漏掉）');
    ok(/補填/.test(await b.ev(`document.getElementById('oppEditDealIdErr').textContent`)), 'P29 立即提示請補填');
    await clearReqs(b);
    await b.click('#oppEditSave'); await b.sleep(400);
    eq((await reqs(b, /\/api\/opportunities\//)).filter((x) => x.m === 'PUT').length, 0, 'P30 沒補 Deal ID 按儲存：前端擋下，沒送 PUT');
    await b.click('#oppEditDealId');
    await b.type('FIX-REQ-1');
    sv = await saveClick(b, 'o-nocat-req');
    ok(sv.body && !('product' in sv.body) && sv.body.dealId === 'FIX-REQ-1', 'P31 補填後送出：帶 dealId、不帶 product', J(sv.body));
    eq(await server('o-nocat-req').then((o) => [o.product, o.dealId]), ['SAP Public Cloud License', 'FIX-REQ-1'], 'P32 伺服器：商品保持、Deal ID 已補');

    // ── 8. 看板拖曳 ──
    console.log('  [UI] 8 看板拖曳');
    await b.ev(`(async()=>{await loadOpportunities(); renderKanban(); return 1})()`);
    await clearReqs(b);
    const dr = await b.ev(`(()=>{const card=document.querySelector('.kanban-card[data-id="o-nocat-drag"]'); if(!card) return 'nocard'; card.dispatchEvent(new DragEvent('dragstart',{bubbles:true})); const col=document.getElementById('kanbanCardsB'); col.dispatchEvent(new DragEvent('drop',{bubbles:true,cancelable:true})); return 'ok'})()`);
    eq(dr, 'ok', 'P33 找得到看板卡片並觸發拖曳');
    await b.waitFor(`window.__reqs.some(r=>/o-nocat-drag$/.test(r.u))`, 5000);
    await b.sleep(500);
    const dput = (await reqs(b, /o-nocat-drag$/)).find((x) => x.m === 'PUT');
    eq(dput && JSON.parse(dput.b), { stage: 'B' }, 'P34 看板拖曳 PUT 只有 {stage}');
    eq(await server('o-nocat-drag').then((o) => [o.product, o.stage]), ['Legacy Product D', 'B'], 'P35 拖曳後商品保持原值');

    // ── 9. XSS ──
    console.log('  [UI] 9 XSS');
    st = await openEdit(b, 'o-xss-its');
    eq([st.value, st.selText], [XSS, '（原值）' + XSS], 'P36 酬載商品：「（原值）」選項以純文字顯示（textContent）');
    eq(await b.ev(`document.querySelectorAll('#oppEditProduct img').length + document.querySelectorAll('#oppEditOverlay img[src="x"]').length`), 0, 'P37 沒有任何 <img> 被注入');
    await selSet(b, 'oppEditStage', 'B');
    sv = await saveClick(b, 'o-xss-its');
    eq(await server('o-xss-its').then((o) => o.product), XSS, 'P38 儲存後商品字串逐字保留（不被跳脫、不被清空）');
    st = await openEdit(b, 'o-xss-nocat');
    eq(st.groupShown, false, 'P39 無類別酬載商機：群組隱藏');
    await selSet(b, 'oppEditStage', 'B');
    sv = await saveClick(b, 'o-xss-nocat');
    ok(sv.body && !('product' in sv.body), 'P40 PUT 不含 product', J(sv.body));
    eq(await server('o-xss-nocat').then((o) => o.product), XSS, 'P41 伺服器保持原字串');
    eq(await b.ev(`window.__xss === undefined`), true, 'P42 全程沒有觸發 onerror（window.__xss 未設定）');

    // ── 10. 新增模式不受影響 ──
    console.log('  [UI] 10 新增模式');
    await clearReqs(b);
    await b.ev(`openOppCreate('沙盒公司A'); 0`);
    await b.sleep(450);
    await selSet(b, 'oppEditCategory', 'ITS');
    await selSet(b, 'oppEditProduct', 'WMS');
    await b.ev(`(()=>{document.getElementById('oppEditAmount').value='10'; document.getElementById('oppEditExpectedDate').value='${YEAR}-12-30'; return 0})()`);
    await b.click('#oppEditSave');
    await b.waitFor(OVERLAY_CLOSED, 8000);
    const post = (await reqs(b, /\/api\/opportunities$/)).find((x) => x.m === 'POST');
    eq(post && JSON.parse(post.b).product, 'WMS', 'P43 新增模式：POST 照常帶 product');

    const errs = b.errors.filter((e) => !/favicon|Failed to load resource/.test(e));
    ok(errs.length === 0, 'P44 全程無頁面例外／console.error', errs.slice(0, 3).join(' | '));
  });
  ok((withBrowser.leftover || 0) === 0, '測試結束沒有殘留的 Edge 程序');
  await stopServer(); cleanupFiles();
}

// ───────────────── 變異測試 ─────────────────
async function mutants() {
  console.log('\n[變異] 拿掉保留邏輯，前端群組必須變紅');
  const src = fs.readFileSync(path.join(ROOT, '_client', 'app.js'), 'utf8');
  const MUTS = [
    { name: 'M1 儲存時不再略過空 product', find: `if (!payload.product && oppEditProductPreserved()) delete payload.product;`, rep: '' },
    { name: 'M2 拿掉「（原值）」選項（改回舊的 ps.value=其他）', find: `if (allValues.includes('其他')) {`, rep: `if (true) {` },
    { name: 'M3 oppEditProductPreserved 恆為 false', find: `function oppEditProductPreserved() { return !!(`, rep: `function oppEditProductPreserved() { return false && !!(` },
    { name: 'M4 Deal ID 規則不看原商品', find: `return (!v && oppEditProductPreserved()) ? _oppEditOrig.product : v;`, rep: `return v;` },
    { name: 'M5 改回原類別時不還原原商品', find: `st.productTouched = false; populateOppEditProduct(this.value, st.product);`, rep: `st.productTouched = false; populateOppEditProduct(this.value);` },
  ];
  const results = [];
  // M0：整份換回修正前的 app.js（origin/main 上沒有 _oppEditOrig 的版本）→ 這支腳本必須抓得到原本的 bug
  try {
    const old = require('child_process').execFileSync('git', ['show', 'origin/main:_client/app.js'], { cwd: ROOT, maxBuffer: 64 * 1024 * 1024 }).toString('utf8');
    if (!old.includes('_oppEditOrig')) MUTS.unshift({ name: 'M0 修正前的 app.js（origin/main）', whole: old });
  } catch (e) { console.log('  （略過 M0：讀不到 origin/main 的 app.js：' + e.message.slice(0, 80) + '）'); }
  for (const m of MUTS) {
    if (!m.whole && !src.includes(m.find)) { ok(false, m.name + '：找不到變異錨點（程式已變動，請更新腳本）'); continue; }
    const snap = { pass, fail, n: failures.length };
    quiet = true;
    uiOnce.mutate = m.whole || src.replace(m.find, () => m.rep);
    try { await uiOnce(); } catch (e) { fail++; failures.push('變異執行例外：' + e.message); }
    uiOnce.mutate = null;
    quiet = false;
    const newFails = failures.slice(snap.n);
    // 還原計數：變異跑出來的失敗是「預期的紅」，不算進本腳本的失敗
    pass = snap.pass; fail = snap.fail; failures.length = snap.n;
    const red = newFails.length > 0;
    results.push({ name: m.name, red, n: newFails.length, first: newFails[0] });
    ok(red, m.name + ' → 前端群組必須變紅', '沒有任何檢查失敗＝這個保護是裝飾品');
    console.log(`  ${red ? '紅（符合預期）' : '仍然全綠（有問題）'}：${m.name}（${newFails.length} 項失敗，例：${(newFails[0] || '').slice(0, 60)}）`);
  }
  return results;
}

(async () => {
  try {
    if (want('api')) await groupApi();
    if (want('ui')) { console.log('  [UI] headless Edge + CDP'); await uiOnce(); }
    if (process.env.OPPPRES_MUT === '1') await mutants();
  } catch (e) {
    fail++; failures.push('例外：' + (e && e.stack || e)); console.log('例外：', e && e.stack || e);
  } finally {
    await stopServer(); cleanupFiles();
  }
  console.log(`\n商機商品保留檢查：PASS ${pass} / FAIL ${fail}`);
  if (fail) console.log('失敗項目：\n  ' + failures.join('\n  '));
  process.exit(fail ? 1 : 0);
})();
