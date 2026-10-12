#!/usr/bin/env node
/**
 * 統編 ↔ 公司名稱一致性守門 檢查。用法：node scripts/check-taxid-guard.js
 * 動 _client/taxid-guard.js、server.js 的 POST/PUT /api/contacts 守門、app.js 的 tg* 區塊、
 * /api/company-lookup 的 lookupStatus 之後必跑。
 *
 *   A) 單元：正規化 / 純數字 / 檢查碼（對照獨立參考實作，隨機 2 萬組）/ evaluate / checkSave
 *   B) 接線：index.html 載入順序、server.js 版本注入與 require、vercel includeFiles、sw 版號
 *   C) 後端 e2e（隔離沙盒：獨立埠、本 worktree 內全新 data.json/auth.json、臨時 _s_ 帳號，結束全部移除）
 *   D) 前端 DOM（Edge headless + CDP；/api/company-master 與 /api/company-lookup 由頁面內 fetch stub 提供，不碰網際網路）
 *   E) 變異：把 sameCompany 比對拆掉，A 的單元必須失敗
 *
 * 環境變數：TG_PORT（預設 3081）、TG_CDP_PORT（預設 9381）、TG_SHOTS（截圖輸出資料夾，未設則不截圖）、
 *           TG_SKIP_DOM=1（略過 D，例如無 Edge）
 */
'use strict';
const fs = require('fs');
const path = require('path');
const os = require('os');
const vm = require('vm');
const { spawn, execFileSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.TG_PORT || '3081', 10);
const CDP_PORT = parseInt(process.env.TG_CDP_PORT || '9381', 10);
if ([3000, 3020, 3031, 3041, 3051, 3061, 3071].includes(PORT)) { console.error('此埠保留給其他 session，請換埠'); process.exit(2); }
const SHOTS = process.env.TG_SHOTS || '';
const TGSRC = fs.readFileSync(path.join(ROOT, '_client/taxid-guard.js'), 'utf8');
const TG = require(path.join(ROOT, '_client/taxid-guard.js'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);

// ═══ A) 單元 ═══════════════════════════════════════════════
function refChecksum(tid) {   // 獨立參考實作（刻意寫法不同）
  if (!/^\d{8}$/.test(tid)) return false;
  const d = tid.split('').map(Number), w = [1, 2, 1, 2, 1, 2, 4, 1];
  const s = d.map((x, i) => x * w[i]).map(p => (p >= 10 ? Math.floor(p / 10) + p % 10 : p)).reduce((a, b) => a + b, 0);
  return s % 5 === 0 || (d[6] === 7 && (s + 1) % 5 === 0);
}
function runUnit(G, label) {
  const out = [];
  const u = (n, ok, e) => out.push([label + n, !!ok, e === undefined ? '' : String(e)]);
  const same = G.sameCompany;
  u('正規化：(股) ≡ 股份有限公司', same('測試科技(股)', '測試科技股份有限公司'));
  u('正規化：全形（股）≡ 股份有限公司', same('測試科技（股）', '測試科技股份有限公司'));
  u('正規化：(有) ≡ 有限公司', same('測試(有)', '測試有限公司'));
  u('正規化：台 ≡ 臺', same('台灣測試股份有限公司', '臺灣測試股份有限公司'));
  u('正規化：全形英數與全形空白', same('ＡＢＣ　科技股份有限公司', 'abc科技股份有限公司'));
  u('正規化：空白/大小寫', same(' Foo  Bar Co., Ltd. ', 'foo bar co ltd'));
  u('保守：股份有限公司 ≠ 有限公司', !same('測試股份有限公司', '測試有限公司'));
  u('保守：省略後綴 ≠ 完整名稱', !same('測試科技', '測試科技股份有限公司'));
  u('保守：不同公司', !same('甲公司', '乙公司'));
  u('保守：空字串不相等', !same('', ''));
  u('純數字：半形', G.isDigitsOnly('70624227'));
  u('純數字：全形', G.isDigitsOnly('７０６２４２２７'));
  u('純數字：含空白', G.isDigitsOnly(' 123 456 '));
  u('純數字：7-11 不是', !G.isDigitsOnly('7-11'));
  u('純數字：空字串不是', !G.isDigitsOnly(''));
  u('檢查碼：22099131 有效', G.isValidTaxIdChecksum('22099131'));
  u('檢查碼：70624227 有效', G.isValidTaxIdChecksum('70624227'));
  u('檢查碼：末碼錯無效', !G.isValidTaxIdChecksum('22099132'));
  u('檢查碼：非 8 碼無效', !G.isValidTaxIdChecksum('2209913') && !G.isValidTaxIdChecksum('abcdefgh'));
  let diff = 0, n7 = 0;
  for (let i = 0; i < 20000; i++) {
    const s = String(Math.floor(Math.random() * 1e8)).padStart(8, '0');
    if (G.isValidTaxIdChecksum(s) !== refChecksum(s)) diff++;
    if (s[6] === '7') n7++;
  }
  u('檢查碼：隨機 2 萬組與參考實作 0 差異（含第 7 碼=7 的 ' + n7 + ' 組）', diff === 0, 'diff=' + diff);
  const ev = G.evaluate;
  u('evaluate：不符 → TAXID_COMPANY_MISMATCH', (ev({ company: '乙公司', taxId: '22099131', officialName: '甲公司' }) || {}).code === 'TAXID_COMPANY_MISMATCH');
  u('evaluate：訊息含統編與官方名稱', /統編 22099131 對應公司為「甲公司」，與輸入的公司名稱不符，請修正/.test((ev({ company: '乙公司', taxId: '22099131', officialName: '甲公司' }) || {}).error || ''));
  u('evaluate：一致（寫法不同）通過', ev({ company: '甲(股)', taxId: '22099131', officialName: '甲股份有限公司' }) === null);
  u('evaluate：官方名稱未知不擋', ev({ company: '乙公司', taxId: '22099131', officialName: '' }) === null);
  u('evaluate：官方名稱為純數字（主檔被污染）不當基準', ev({ company: '乙公司', taxId: '22099131', officialName: '22099131' }) === null);
  u('evaluate：境外不比對', ev({ company: '乙公司', taxId: '22099131', isForeign: true, officialName: '甲公司' }) === null);
  u('evaluate：純數字=統編 → COMPANY_IS_TAXID', (ev({ company: '22099131', taxId: '22099131' }) || {}).code === 'COMPANY_IS_TAXID');
  u('evaluate：純數字 8 碼（不等於統編）也擋', (ev({ company: '12345678', taxId: '22099131' }) || {}).code === 'COMPANY_IS_TAXID');
  u('evaluate：短數字後端不擋', ev({ company: '104', taxId: '22099131' }) === null);
  u('evaluate：短數字前端 strictDigits 擋', (ev({ company: '104', taxId: '22099131', strictDigits: true }) || {}).code === 'COMPANY_IS_TAXID');
  u('evaluate：空公司交給必填規則', ev({ company: '', taxId: '22099131', officialName: '甲公司' }) === null);
  const comps = [{ matchKey: 'tax:22099131', taxId: '22099131', name: '甲股份有限公司' }, { matchKey: 'tax:70624227', taxId: '70624227', name: '70624227' }];
  u('masterOfficialName：命中', G.masterOfficialName(comps, '22099131') === '甲股份有限公司');
  u('masterOfficialName：被污染主檔回空', G.masterOfficialName(comps, '70624227') === '');
  u('masterOfficialName：未命中回空', G.masterOfficialName(comps, '11111111') === '');
  const cs = G.checkSave;
  u('checkSave：新增不符 → 擋', (cs({ company: '乙', taxId: '22099131', companies: comps }) || {}).code === 'TAXID_COMPANY_MISMATCH');
  u('checkSave：更新、公司/統編沒動 → 不驗（舊資料）', cs({ company: '乙', taxId: '22099131', companies: comps, isUpdate: true, oldCompany: '乙', oldTaxId: '22099131' }) === null);
  u('checkSave：更新、公司改成不符 → 擋', (cs({ company: '丙', taxId: '22099131', companies: comps, isUpdate: true, oldCompany: '乙', oldTaxId: '22099131' }) || {}).code === 'TAXID_COMPANY_MISMATCH');
  u('checkSave：更新、統編改到不符的主檔 → 擋', (cs({ company: '乙', taxId: '22099131', companies: comps, isUpdate: true, oldCompany: '乙', oldTaxId: '11111111' }) || {}).code === 'TAXID_COMPANY_MISMATCH');
  u('checkSave：更新、舊資料純數字但沒動 → 不擋', cs({ company: '70624227', taxId: '70624227', companies: comps, isUpdate: true, oldCompany: '70624227', oldTaxId: '70624227' }) === null);
  u('checkSave：更新、改成純數字 → 擋', (cs({ company: '70624227', taxId: '70624227', companies: comps, isUpdate: true, oldCompany: '乙', oldTaxId: '70624227' }) || {}).code === 'COMPANY_IS_TAXID');
  u('checkSave：統編不在內部主檔 → 無法驗證、放行', cs({ company: '任意', taxId: '33333333', companies: comps }) === null);
  return out;
}

// ═══ B) 接線 ═══════════════════════════════════════════════
function runWiring() {
  const idx = fs.readFileSync(path.join(ROOT, '_client/index.html'), 'utf8');
  const srv = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const vj = JSON.parse(fs.readFileSync(path.join(ROOT, 'vercel.json'), 'utf8'));
  const app = fs.readFileSync(path.join(ROOT, '_client/app.js'), 'utf8');
  t('接線：index.html 先載 taxid-guard.js 再載 app.js', idx.indexOf('src="taxid-guard.js"') > 0 && idx.indexOf('src="taxid-guard.js"') < idx.indexOf('src="app.js"'));
  t('接線：server.js 有 require 與版本注入', /require\('\.\/_client\/taxid-guard\.js'\)/.test(srv) && /src="taxid-guard\\\.js"/.test(srv));
  t('接線：vercel includeFiles 含 _client', /_client/.test(vj.functions['api/index.js'].includeFiles));
  t('接線：POST 與 PUT 都呼叫 TaxIdGuard.checkSave', (srv.match(/TaxIdGuard\.checkSave\(/g) || []).length === 2);
  t('接線：app.js 存檔前有守門且按鈕狀態由 tgRender 控制', /const tg = tgRender\(\);\s*\n\s*if \(tg\.block\)/.test(app) && /sb\.disabled = !!g\.block/.test(app));
  t('接線：admin.html/其他檔不受影響（app.js 語法）', (() => { try { new vm.Script(app); return true; } catch (e) { return false; } })());
}

// ═══ 沙盒工具 ═══════════════════════════════════════════════
const created = [];      // 本次建立的檔案（結束刪除）
const pids = [];         // 啟動的 PID（結束只殺這些）
function killPid(pid) { try { execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' }); } catch (e) { /* 已結束 */ } }
const sleep = (ms) => new Promise(r => setTimeout(r, ms));
function validId(prefix7) { for (let d = 0; d < 10; d++) { const s = prefix7 + d; if (refChecksum(s)) return s; } throw new Error('no valid id'); }

const ID_A = validId('1234567');   // 主檔：測試甲
const ID_B = validId('2345678');   // 主檔被污染（名稱=統編）
const ID_C = validId('3456789');   // 不在主檔
const ID_D = validId('4567890');   // 主檔：測試乙（舊資料不符用）
const NAME_A = '測試甲股份有限公司';
const NAME_D = '測試乙股份有限公司';

function seedFiles() {
  const bcrypt = require(path.join(ROOT, 'node_modules/bcryptjs'));
  const dataFile = path.join(ROOT, 'data.json'), authFile = path.join(ROOT, 'auth.json');
  if (fs.existsSync(dataFile) || fs.existsSync(authFile)) throw new Error('worktree 內已有 data.json/auth.json，拒絕覆蓋（沙盒檢查只在乾淨 worktree 跑）');
  const now = new Date().toISOString();
  const master = (id, tax, name) => ({ id, taxId: tax, matchKey: 'tax:' + tax, name, industry: '', address: '測試路 1 號', createdAt: now, updatedAt: now, buImport: [] });
  const contact = (id, company, taxId, extra) => Object.assign({ id, owner: '_s_sales', bu: 'ERP', name: '測試' + id, company, taxId, email: id + '@example.test', createdAt: now, deleted: false }, extra || {});
  const data = {
    contacts: [
      contact('L1', ID_C, ID_C),                    // 舊資料：公司=統編（純數字）
      contact('L2', '測試乙', ID_D),                  // 舊資料：公司 ≠ 主檔官方名稱
      contact('L3', '正常公司', ID_C),                // 舊資料：統編不在主檔
    ],
    companies: [master('M1', ID_A, NAME_A), master('M2', ID_B, ID_B), master('M3', ID_D, NAME_D)],
  };
  fs.writeFileSync(dataFile, JSON.stringify(data, null, 2)); created.push(dataFile);
  const auth = { users: [{ username: '_s_sales', password: bcrypt.hashSync('Tg-Sandbox-9x!zQ', 10), displayName: '沙盒業務', role: 'user', bu: ['ERP'], active: true, passwordChangedAt: now }] };
  fs.writeFileSync(authFile, JSON.stringify(auth, null, 2)); created.push(authFile);
  return 'Tg-Sandbox-9x!zQ';
}
async function startServer() {
  const logFile = path.join(os.tmpdir(), 'taxid-guard-server.log');
  const out = fs.openSync(logFile, 'w');
  const p = spawn(process.execPath, ['server.js'], { cwd: ROOT, stdio: ['ignore', out, out], env: Object.assign({}, process.env, { PORT: String(PORT), DB_BACKEND: 'json', SESSION_SECRET: 'taxid-guard-sandbox-secret', NODE_ENV: 'development', ALLOWED_ORIGINS: 'http://localhost:' + PORT }) });
  pids.push(p.pid);
  for (let i = 0; i < 60; i++) {
    try { const r = await fetch(`http://127.0.0.1:${PORT}/login.html`); if (r.ok) return p; } catch (e) { /* 等待 */ }
    await sleep(500);
  }
  throw new Error('server 沒起來：' + fs.readFileSync(logFile, 'utf8').slice(-500));
}

// ═══ C) 後端 e2e ═══════════════════════════════════════════
async function runBackend(pw) {
  const base = `http://127.0.0.1:${PORT}`;
  let cookie = '';
  const api = async (method, url, body) => {
    const r = await fetch(base + url, { method, headers: { 'Content-Type': 'application/json', Cookie: cookie }, body: body ? JSON.stringify(body) : undefined });
    const sc = r.headers.get('set-cookie'); if (sc) cookie = sc.split(';')[0];
    let j = null; try { j = await r.json(); } catch (e) { /* 非 JSON */ }
    return { s: r.status, j };
  };
  const lg = await api('POST', '/api/login', { username: '_s_sales', password: pw });
  t('e2e：沙盒帳號登入', lg.s === 200 && lg.j && lg.j.success, lg.s);
  const base0 = { name: 'x', email: 'x@example.test', bu: 'ERP' };
  const post = (company, taxId, extra) => api('POST', '/api/contacts', Object.assign({}, base0, { company, taxId }, extra || {}));
  const count = async () => (await api('GET', '/api/contacts')).j.length;
  const n0 = await count();
  let r;
  r = await post(NAME_A, ID_A); t('e2e POST：與主檔一致 → 201', r.s === 201, r.s + ' ' + JSON.stringify(r.j && r.j.error));
  r = await post('測試甲(股)', ID_A); t('e2e POST：寫法不同（(股)）→ 201', r.s === 201, r.s);
  r = await post('測試甲 股份有限公司', ID_A); t('e2e POST：含空白 → 201', r.s === 201, r.s);
  const n1 = await count();
  r = await post('別家股份有限公司', ID_A); t('e2e POST：不符 → 400 TAXID_COMPANY_MISMATCH', r.s === 400 && r.j.code === 'TAXID_COMPANY_MISMATCH', r.s + ' ' + (r.j && r.j.code));
  t('e2e POST：不符訊息含官方名稱', r.j && r.j.error && r.j.error.includes(NAME_A) && r.j.error.includes(ID_A));
  r = await post(ID_C, ID_C); t('e2e POST：公司=統編(未知統編) → 400 COMPANY_IS_TAXID', r.s === 400 && r.j.code === 'COMPANY_IS_TAXID', r.s + ' ' + (r.j && r.j.code));
  r = await post(ID_A, ID_A); t('e2e POST：公司=統編(主檔統編) → 400 COMPANY_IS_TAXID', r.s === 400 && r.j.code === 'COMPANY_IS_TAXID', r.s);
  r = await post('12345678', '', { isForeign: true }); t('e2e POST：境外 + 8 碼純數字公司 → 400', r.s === 400 && r.j.code === 'COMPANY_IS_TAXID', r.s);
  t('e2e POST：被拒絕的請求沒有寫入', (await count()) === n1);
  r = await post('任意新公司', ID_C); t('e2e POST：統編不在內部主檔 → 無法驗證、放行 201（限制）', r.s === 201, r.s);
  r = await post('某公司', ID_B); t('e2e POST：主檔被污染(名稱=統編) → 不當基準、放行 201', r.s === 201, r.s);
  r = await post(ID_B, ID_B); t('e2e POST：主檔被污染但公司填數字 → 400', r.s === 400, r.s);
  r = await post('境外公司 Inc.', '', { isForeign: true }); t('e2e POST：境外一般公司 → 201（既有流程不變）', r.s === 201, r.s);
  // PUT（舊資料）
  const put = (id, body) => api('PUT', '/api/contacts/' + id, body);
  r = await put('L1', { phone: '02-1234' }); t('e2e PUT：舊資料(公司=統編)只改電話 → 200（不 hard-fail）', r.s === 200, r.s);
  r = await put('L1', { phone: '02-5678', company: ID_C, taxId: ID_C }); t('e2e PUT：前端整包送回但公司/統編沒動 → 200', r.s === 200, r.s);
  r = await put('L1', { company: ID_C, taxId: ID_C, isForeign: false, title: '經理' }); t('e2e PUT：同上帶 isForeign → 200', r.s === 200, r.s);
  r = await put('L1', { company: '53' + ID_C.slice(2), taxId: ID_C }); t('e2e PUT：舊資料改成另一組純數字 → 400', r.s === 400 && r.j.code === 'COMPANY_IS_TAXID', r.s);
  r = await put('L2', { company: '測試乙', taxId: ID_D, phone: '1' }); t('e2e PUT：舊資料(公司≠主檔)整包送回未變 → 200', r.s === 200, r.s);
  r = await put('L2', { title: '協理' }); t('e2e PUT：舊資料只改職稱(不帶公司/統編) → 200', r.s === 200, r.s);
  r = await put('L2', { company: '亂改公司', taxId: ID_D }); t('e2e PUT：舊資料把公司改成與主檔不符 → 400', r.s === 400 && r.j.code === 'TAXID_COMPANY_MISMATCH', r.s);
  r = await put('L2', { company: NAME_D, taxId: ID_D }); t('e2e PUT：改成主檔官方名稱 → 200', r.s === 200, r.s);
  r = await put('L1', { company: NAME_A, taxId: ID_A }); t('e2e PUT：修好舊的純數字(公司=官方名稱、統編改到主檔) → 200', r.s === 200, r.s);
  r = await put('L3', { company: '正常公司', taxId: ID_A }); t('e2e PUT：只改統編到不符的主檔 → 400', r.s === 400 && r.j.code === 'TAXID_COMPANY_MISMATCH', r.s);
  const cs = (await api('GET', '/api/contacts')).j;
  t('e2e：L1 修正後公司為官方名稱', (cs.find(c => c.id === 'L1') || {}).company === NAME_A);
  t('e2e：L3 被拒絕後維持原值', (cs.find(c => c.id === 'L3') || {}).taxId === ID_C);
  // lookupStatus 欄位存在（不打外網：無統編格式錯誤時 400；格式正確會連 GCIS，沙盒內不呼叫）
  r = await api('GET', '/api/company-lookup?taxId=abc'); t('e2e：/api/company-lookup 格式錯誤仍 400（既有行為）', r.s === 400, r.s);
  t('e2e：GCIS 區塊含 lookupStatus 欄位（靜態）', /lookupStatus: 'error'/.test(fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8')) && /result\.lookupStatus = data\?\.\[0\] \? 'found' : 'notfound'/.test(fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8')));
  return cookie;
}

// ═══ D) 前端 DOM（Edge headless + CDP）═════════════════════
function findEdge() {
  const c = ['C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', 'C:/Program Files/Microsoft/Edge/Application/msedge.exe'];
  return c.find(p => fs.existsSync(p));
}
async function runDom(pw) {
  const edge = findEdge();
  if (!edge || process.env.TG_SKIP_DOM) { t('DOM：略過（無 Edge 或 TG_SKIP_DOM）', true, '未驗證'); return; }
  const prof = fs.mkdtempSync(path.join(os.tmpdir(), 'taxid-edge-'));
  const ep = spawn(edge, ['--headless=new', '--disable-gpu', `--remote-debugging-port=${CDP_PORT}`, `--user-data-dir=${prof}`, '--window-size=1280,1100', 'about:blank'], { stdio: 'ignore' });
  pids.push(ep.pid);
  let targets;
  for (let i = 0; i < 40; i++) { try { targets = await (await fetch(`http://127.0.0.1:${CDP_PORT}/json`)).json(); if (targets.some(x => x.type === 'page')) break; } catch (e) { /* 等待 */ } await sleep(300); }
  const page = targets.find(x => x.type === 'page');
  const ws = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((ok, no) => { ws.onopen = ok; ws.onerror = no; });
  let seq = 0; const waiting = {};
  ws.onmessage = (m) => { const d = JSON.parse(m.data); if (d.id && waiting[d.id]) { waiting[d.id](d); delete waiting[d.id]; } };
  const send = (method, params) => new Promise(r => { const id = ++seq; waiting[id] = r; ws.send(JSON.stringify({ id, method, params: params || {} })); });
  const ev = async (expr) => { const r = await send('Runtime.evaluate', { expression: expr, awaitPromise: true, returnByValue: true }); if (r.result.exceptionDetails) throw new Error('eval: ' + JSON.stringify(r.result.exceptionDetails).slice(0, 400)); return r.result.result.value; };
  const waitFor = async (expr, ms) => { const end = Date.now() + (ms || 6000); while (Date.now() < end) { if (await ev(expr)) return true; await sleep(100); } return false; };
  const shot = async (name) => { if (!SHOTS) return; fs.mkdirSync(SHOTS, { recursive: true }); await ev(`(function(){var e=document.getElementById('company');if(e)e.scrollIntoView({block:'start'});window.scrollBy(0,-0);})()`); await sleep(200); const r = await send('Page.captureScreenshot', { format: 'png' }); fs.writeFileSync(path.join(SHOTS, name + '.png'), Buffer.from(r.result.data, 'base64')); };
  const nav = async (url) => { await send('Page.enable'); await send('Page.navigate', { url }); await sleep(800); await waitFor('document.readyState==="complete"', 10000); };

  await send('Emulation.setDeviceMetricsOverride', { width: 1100, height: 900, deviceScaleFactor: 1, mobile: false });
  await nav(`http://localhost:${PORT}/login.html`);
  const lg = await ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:'_s_sales',password:${JSON.stringify(pw)}})}).then(r=>r.status)`);
  t('DOM：登入', lg === 200, lg);
  await nav(`http://localhost:${PORT}/`);
  t('DOM：主頁載入且 TaxIdGuard 可用', await waitFor('typeof openModal==="function" && !!window.TaxIdGuard', 15000));

  // 頁面內 fetch stub：mock.master / mock.gcis 決定回應；calls 計次
  await ev(`(function(){
    window.__calls={master:0,gcis:0};
    window.__mock={master:{found:false}, gcis:{companyName:'',lookupStatus:'notfound'}, gcisFail:false};
    const real=window.fetch.bind(window);
    window.fetch=function(u,o){
      const s=String(u);
      const j=(b,st)=>Promise.resolve(new Response(JSON.stringify(b),{status:st||200,headers:{'Content-Type':'application/json'}}));
      if(s.indexOf('/api/company-master')===0){window.__calls.master++;return j(window.__mock.master);}
      if(s.indexOf('/api/company-lookup')===0){window.__calls.gcis++;return window.__mock.gcisFail?j({error:'x'},502):j(window.__mock.gcis);}
      return real(u,o);
    };
  })()`);
  const setTax = (v) => ev(`(function(){const e=document.getElementById('taxId');e.value=${JSON.stringify(v)};e.dispatchEvent(new Event('input'));})()`);
  const setCo = (v) => ev(`(function(){const e=document.getElementById('company');e.value=${JSON.stringify(v)};e.dispatchEvent(new Event('input'));})()`);
  const st = () => ev(`({co:document.getElementById('company').value,tax:document.getElementById('taxId').value,save:document.getElementById('saveBtn').disabled,status:document.getElementById('taxGuardStatus').style.display==='none'?'':document.getElementById('taxGuardStatus').innerText,err:document.getElementById('companyGuardErr').style.display==='none'?'':document.getElementById('companyGuardErr').innerText,hasUse:!!document.getElementById('useOfficialNameBtn'),hasRetry:!!document.getElementById('taxGuardRetry'),calls:window.__calls})`);
  const mock = (o) => ev(`Object.assign(window.__mock,${JSON.stringify(o)})`);
  const reopen = async () => { await ev(`openModal()`); await sleep(100); await ev(`window.__calls.master=0;window.__calls.gcis=0`); };
  const OFF = '測試丙股份有限公司', TAX = validId('5678901');
  const GC = { companyName: OFF, address: '測試市測試路 9 號', lookupStatus: 'found' };
  let s;

  // 1) GCIS 命中 → 自動帶入公司
  await reopen(); await mock({ master: { found: false }, gcis: GC, gcisFail: false });
  await setTax(TAX);
  t('DOM 1：輸入統編 → 公司欄自動帶入官方名稱', await waitFor(`document.getElementById('company').value===${JSON.stringify(OFF)}`));
  s = await st();
  t('DOM 1：顯示「已比對」且可存檔', /已比對/.test(s.status) && s.save === false && !s.err, JSON.stringify(s));
  await shot('1-auto-filled-success');
  // 快取：同統編再輸入不再打 API
  await ev(`window.__calls.master=0;window.__calls.gcis=0`);
  await setCo(''); await setTax(''); await setTax(TAX);
  await waitFor(`document.getElementById('company').value===${JSON.stringify(OFF)}`);
  s = await st(); t('DOM 1b：同統編第二次走快取，不再呼叫 API', s.calls.master === 0 && s.calls.gcis === 0, JSON.stringify(s.calls));

  // 2) 公司與官方不符 → 紅色錯誤、擋存檔、使用官方名稱
  await reopen();
  await setCo('別家股份有限公司'); await setTax(TAX);
  t('DOM 2：不符 → 出現錯誤訊息', await waitFor(`document.getElementById('companyGuardErr').style.display==='block'`));
  s = await st();
  t('DOM 2：訊息含統編與官方名稱、有「使用官方名稱」、存檔被停用', s.err.includes(TAX) && s.err.includes(OFF) && s.hasUse && s.save === true, JSON.stringify(s));
  t('DOM 2：不符時不覆蓋使用者輸入的公司', s.co === '別家股份有限公司');
  await shot('2-mismatch-error');
  await ev(`document.getElementById('useOfficialNameBtn').click()`);
  await sleep(150); s = await st();
  t('DOM 2：按「使用官方名稱」→ 公司改為官方、錯誤消失、可存檔', s.co === OFF && !s.err && s.save === false, JSON.stringify(s));
  // 寫法不同（(股)）視為一致
  await setCo(OFF.replace('股份有限公司', '(股)')); await sleep(100); s = await st();
  t('DOM 2b：(股) 寫法視為一致、不報錯', !s.err && s.save === false, JSON.stringify(s));

  // 3) 純數字公司
  await reopen(); await mock({ gcis: GC });
  await setCo(TAX); await sleep(100); s = await st();
  t('DOM 3：公司欄純數字 → 錯誤並擋存檔', /不可只有數字/.test(s.err) && s.save === true, JSON.stringify(s));

  // 4) 無法驗證（GCIS 失敗）→ 擋存檔、可重試；恢復後通過
  await reopen(); await mock({ master: { found: false }, gcisFail: true });
  await setCo('任意公司'); await setTax(validId('9012345'));   // 另一組統編（TAX 已在本頁快取，不能重用）
  t('DOM 4：GCIS 失敗 → 顯示「無法驗證」與重新查詢', await waitFor(`!!document.getElementById('taxGuardRetry')`, 8000));
  s = await st();
  t('DOM 4：自動重試 1 次（共 2 次 GCIS 呼叫）', s.calls.gcis === 2, JSON.stringify(s.calls));
  t('DOM 4：無法驗證時存檔被擋', s.save === true && /無法驗證/.test(s.status), JSON.stringify(s));
  await shot('4-lookup-unavailable-blocked');
  await mock({ gcisFail: false, gcis: GC });
  await ev(`document.getElementById('taxGuardRetry').click()`);
  t('DOM 4：按「重新查詢」成功後 → 比對不符錯誤出現（verified）', await waitFor(`document.getElementById('companyGuardErr').style.display==='block'`));
  await ev(`document.getElementById('useOfficialNameBtn').click()`); await sleep(150); s = await st();
  t('DOM 4：修正後可存檔', s.save === false && s.co === OFF, JSON.stringify(s));

  // 5) 內部主檔命中 = 已驗證，GCIS 掛了也不影響且不打 GCIS
  await reopen(); await mock({ master: { found: true, company: { name: '測試丁股份有限公司', address: 'A', industry: '' } }, gcisFail: true });
  await setTax(validId('6789012'));
  t('DOM 5：主檔命中 → 帶入公司', await waitFor(`document.getElementById('company').value==='測試丁股份有限公司'`));
  s = await st(); t('DOM 5：不呼叫 GCIS、可存檔', s.calls.gcis === 0 && s.save === false && /企業主檔/.test(s.status), JSON.stringify(s));
  // 主檔名稱被污染（=統編）→ 不當基準，改走 GCIS
  const TAX6 = validId('7890123');
  await reopen(); await mock({ master: { found: true, company: { name: TAX6, address: '' } }, gcisFail: false, gcis: { companyName: '測試戊股份有限公司', lookupStatus: 'found' } });
  await setTax(TAX6);
  t('DOM 5b：主檔名稱為純數字 → 改用 GCIS 官方名稱帶入', await waitFor(`document.getElementById('company').value==='測試戊股份有限公司'`));

  // 6) 查無此統編 → 擋存檔（使用者決定：查無也擋）
  await reopen(); await mock({ master: { found: false }, gcis: { companyName: '', lookupStatus: 'notfound' }, gcisFail: false });
  await setCo('某商號'); await setTax(validId('8901234'));
  t('DOM 6：GCIS 確定查無 → 警告', await waitFor(`/查無此統編/.test(document.getElementById('taxGuardStatus').innerText)`));
  s = await st(); t('DOM 6：查無 → 存檔被擋', s.save === true && s.calls.gcis === 1, JSON.stringify(s));
  await shot('6-notfound-warning');

  // 7) 沒填統編 / 境外：行為不變
  await reopen(); await setCo('無統編公司'); await sleep(100); s = await st();
  t('DOM 7：沒填統編 → 不擋（既有必填規則由原本檢查負責）', s.save === false && !s.err && !s.status, JSON.stringify(s));

  // 8) 編輯舊資料
  const mk = (company, taxId) => JSON.stringify({ id: 'DOM-L', name: '舊', company, taxId, email: 'a@example.test', owner: '_s_sales', bu: 'ERP' });
  await ev(`openModal(${mk('舊公司名', validId('1357913'))})`);   // 未快取的統編 await sleep(200);
  await ev(`window.__calls.master=0;window.__calls.gcis=0`);
  s = await st(); t('DOM 8：編輯舊資料(公司/統編沒動) → 不查詢、不擋', s.save === false && s.calls.gcis === 0 && s.calls.master === 0, JSON.stringify(s));
  await setCo('舊公司名改一下'); await sleep(200);
  t('DOM 8：改了公司才開始要求驗證（觸發查詢）', await waitFor(`window.__calls.gcis>0 || window.__calls.master>0`, 3000));
  await mock({ master: { found: false }, gcis: GC, gcisFail: false });
  await ev(`openModal(${mk(TAX, TAX)})`);
  t('DOM 8b：舊資料公司=統編 → 自動查詢並帶入官方名稱', await waitFor(`document.getElementById('company').value===${JSON.stringify(OFF)}`));
  await ev(`openModal(${mk(TAX, TAX)})`); await mock({ gcisFail: true }); await ev(`taxGuard.cache={}`); await sleep(100);
  await ev(`openModal(${mk(TAX, TAX)})`); await sleep(300);
  s = await st(); t('DOM 8c：舊資料公司=統編且尚未修好 → 擋存檔、提示修正', s.save === true && /不可只有數字/.test(s.err), JSON.stringify(s));

  // 9) 真的按存檔（經 UI 流程 → 真實後端）
  await mock({ master: { found: false }, gcis: GC, gcisFail: false }); await ev(`taxGuard.cache={}`);
  await reopen();
  await ev(`(function(){for(const [id,v] of [['name','沙盒測試人'],['email','dom@example.test']]){document.getElementById(id).value=v;}})()`);
  await setTax(TAX);
  await waitFor(`document.getElementById('company').value===${JSON.stringify(OFF)}`);
  await sleep(200);
  await ev(`document.getElementById('saveBtn').click()`);
  await sleep(1200);
  const list = await ev(`fetch('/api/contacts').then(r=>r.json()).then(a=>a.filter(c=>c.taxId===${JSON.stringify(TAX)}).map(c=>c.company))`);
  t('DOM 9：UI 存檔成功，公司為官方名稱', list.length === 1 && list[0] === OFF, JSON.stringify(list));
  ws.close();
  killPid(ep.pid);
  try { fs.rmSync(prof, { recursive: true, force: true }); } catch (e) { /* Edge 可能還鎖著，下次清 */ }
}

// ═══ E) 變異 ═══════════════════════════════════════════════
function runMutation() {
  const mutated = TGSRC.replace('return !!x && x === y;', 'return true;');   // 拆掉「名稱是否相同」比對
  if (mutated === TGSRC) { t('變異：變異點存在', false); return; }
  const m = { exports: {} };
  new vm.Script('(function(module){' + mutated + '\n})').runInThisContext()(m);
  const failed = runUnit(m.exports, 'M:').filter(x => !x[1]).length;
  t('變異：拆掉 sameCompany 比對後單元測試必須失敗', failed > 0, '失敗 ' + failed + ' 項');
}

// ═══ 主流程 ═══════════════════════════════════════════════
(async () => {
  let server = null;
  try {
    runUnit(TG, '').forEach(x => res.push(x));
    runWiring();
    runMutation();
    const pw = seedFiles();
    server = await startServer();
    await runBackend(pw);
    await runDom(pw);
  } catch (e) {
    t('執行例外', false, e && e.stack || e);
  } finally {
    pids.forEach(killPid);
    await sleep(500);
    created.forEach(f => { try { fs.unlinkSync(f); } catch (e) { /* 已不存在 */ } });
    for (const extra of ['audit.log.json']) { const p = path.join(ROOT, extra); if (fs.existsSync(p)) { try { fs.unlinkSync(p); } catch (e) { /* ignore */ } } }
    let bad = 0;
    for (const [n, ok, e] of res) { console.log((ok ? 'PASS ' : 'FAIL ') + n + (e && !ok ? '  → ' + e : (e ? '  (' + e + ')' : ''))); if (!ok) bad++; }
    console.log(`\n${res.length - bad}/${res.length} 通過` + (bad ? `，${bad} 失敗` : ''));
    process.exit(bad ? 1 : 0);
  }
})();
