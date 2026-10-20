#!/usr/bin/env node
'use strict';
/**
 * CSP Report-Only 實瀏覽器煙霧測試：用無頭 Edge＋CDP 載入 login／admin（逐一點開側欄各區塊）／index（逐一點開導覽）／help，
 * 收集瀏覽器端真實的 securitypolicyviolation，並與伺服器端稽核（CSP_VIOLATION）對照，驗證回報管線完整、
 * 同時比較 CSP_REPORT_ONLY=off 與 on 兩輪的頁面功能指標（JS 例外數、DOM 元素數）與「強制 CSP 違規 = 0」。
 *
 *   node scripts/csp-smoke.js            # 兩輪（off、on），輸出 scratchpad 外的 JSON 由 CSP_SMOKE_OUT 指定（預設 os.tmpdir）
 *   CSP_SMOKE_STYLE=1 ...                # 另跑一輪 CSP_REPORT_ONLY_STYLE=on（style 嚴格觀察，回報量很大）
 *   EDGE_PATH=...                        # Edge 位置（預設 Program Files (x86)\Microsoft\Edge）
 *
 * 守門與埠：沙盒伺服器埠 3080（B2_PORT）、Edge 除錯埠 9333；工作樹已有 auth/data/audit 檔時拒絕執行。
 * Edge 以獨立的暫存使用者設定檔啟動，結束時只對自己的 PID 做 taskkill /T（不碰使用者的其他 Edge）。
 * 整體時間上限 9 分鐘。
 */
process.env.B2_PORT = process.env.B2_PORT || '3080';
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn, spawnSync } = require('child_process');
const H = require('./_b2-harness');
const { sleep, mkUser, fresh, finish, http } = H;

const DEBUG_PORT = parseInt(process.env.CSP_EDGE_PORT || '9333', 10);
const EDGE = process.env.EDGE_PATH || ['C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe', 'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe'].find(p => fs.existsSync(p));
const OUT = process.env.CSP_SMOKE_OUT || path.join(os.tmpdir(), 'csp_smoke_out.json');
const HARD_LIMIT_MS = 9 * 60 * 1000;

let edgePid = null, edgeDir = null;
function killEdge() {
  if (edgePid) { try { spawnSync('taskkill', ['/PID', String(edgePid), '/T', '/F'], { stdio: 'ignore' }); } catch { /* 已結束 */ } edgePid = null; }
}
async function rmDirSoon(d) {
  for (let i = 0; i < 10; i++) { try { fs.rmSync(d, { recursive: true, force: true }); return; } catch { await sleep(500); } }
}

// ── 極小 CDP 用戶端 ──
function cdpConnect(wsUrl) {
  return new Promise((resolve, reject) => {
    const ws = new WebSocket(wsUrl);
    let id = 0; const pending = new Map(); const handlers = [];
    ws.onmessage = (m) => {
      const msg = JSON.parse(m.data);
      if (msg.id && pending.has(msg.id)) { const p = pending.get(msg.id); pending.delete(msg.id); msg.error ? p.rej(new Error(msg.error.message)) : p.res(msg.result); }
      else if (msg.method) handlers.forEach(h => h(msg.method, msg.params));
    };
    ws.onopen = () => resolve({
      send: (method, params) => new Promise((res, rej) => { const i = ++id; pending.set(i, { res, rej }); ws.send(JSON.stringify({ id: i, method, params: params || {} })); setTimeout(() => { if (pending.has(i)) { pending.delete(i); rej(new Error('CDP 逾時 ' + method)); } }, 20000); }),
      on: (h) => handlers.push(h), close: () => ws.close(),
    });
    ws.onerror = (e) => reject(new Error('CDP 連線失敗'));
  });
}

const HOOK = `(() => { window.__csp = window.__csp || [];
  document.addEventListener('securitypolicyviolation', (e) => { window.__csp.push({ d: e.effectiveDirective, b: e.blockedURI, disp: e.disposition, src: e.sourceFile, line: e.lineNumber, sample: (e.sample || '').length }); });
})();`;

async function runPass(label, env) {
  const users = [mkUser('_s_admin', 'admin'), mkUser('_s_user', 'user', { bu: ['ERP'] })];
  await fresh(users, {}, { env });
  edgeDir = fs.mkdtempSync(path.join(os.tmpdir(), 'csp-edge-'));
  const child = spawn(EDGE, ['--headless=new', '--disable-gpu', `--remote-debugging-port=${DEBUG_PORT}`, `--user-data-dir=${edgeDir}`, '--no-first-run', '--no-default-browser-check', '--disable-extensions', '--window-size=1400,900'].concat((process.env.CSP_EDGE_FLAGS || '').split(' ').filter(Boolean), ['about:blank']), { stdio: 'ignore' });
  edgePid = child.pid;
  let target = null;
  for (let i = 0; i < 60 && !target; i++) {
    await sleep(500);
    try { const l = await (await fetch(`http://127.0.0.1:${DEBUG_PORT}/json/list`)).json(); target = l.find(t => t.type === 'page'); } catch { /* 尚未起來 */ }
  }
  if (!target) throw new Error('Edge 未啟動');
  const c = await cdpConnect(target.webSocketDebuggerUrl);
  const exceptions = [], logErrors = [], reportPosts = [];
  c.on((method, p) => {
    if (method === 'Runtime.exceptionThrown') exceptions.push(String((p.exceptionDetails.exception && p.exceptionDetails.exception.description) || p.exceptionDetails.text).slice(0, 160));
    if (method === 'Log.entryAdded' && p.entry.level === 'error') logErrors.push((p.entry.source + ': ' + p.entry.text).slice(0, 160));
    if (method === 'Network.requestWillBeSent' && /\/api\/csp-report/.test(p.request.url)) reportPosts.push({ ct: p.request.headers['Content-Type'] || p.request.headers['content-type'] || '', size: (p.request.postData || '').length });
  });
  await c.send('Page.enable'); await c.send('Runtime.enable'); await c.send('Log.enable'); await c.send('Network.enable');
  await c.send('Page.addScriptToEvaluateOnNewDocument', { source: HOOK });
  const evalJs = async (expr) => { const r = await c.send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: true }); return r.result ? r.result.value : undefined; };
  const events = [];   // 瀏覽器端違規 { page, ... }
  const metrics = {};
  const collect = async (page) => {
    const raw = await evalJs('JSON.stringify(window.__csp||[])');
    let arr = []; try { arr = JSON.parse(raw || '[]'); } catch { /* 空 */ }
    arr.forEach(e => events.push(Object.assign({ page }, e)));
    await evalJs('window.__csp = []');
    return arr.length;
  };
  const nav = async (url, waitMs) => {
    const loaded = new Promise((res) => { const t = setTimeout(res, 15000); c.on((m) => { if (m === 'Page.loadEventFired') { clearTimeout(t); res(); } }); });
    await c.send('Page.navigate', { url: H.BASE + url });
    await loaded; await sleep(waitMs || 2500);
  };
  const login = async (user) => {
    await nav('/login.html', 2500);
    await collect('login');
    await evalJs(`(() => { const u = document.getElementById('username'), p = document.getElementById('password');
      u.value = ${JSON.stringify(user)}; p.value = ${JSON.stringify(H.PW)}; u.dispatchEvent(new Event('input', { bubbles: true })); p.dispatchEvent(new Event('input', { bubbles: true }));
      document.getElementById('loginForm').requestSubmit(); })()`);
    await sleep(7000);
    return evalJs('location.pathname');
  };
  // script-src-attr 違規是「事件執行時」才檢查（行內 on* 屬性在第一次觸發時才編譯），所以只載入頁面看不到；
  // 這裡對「無副作用」的 handler（close*/toggle*/clear*/switch*/load*）各觸發一次，並對 oninput 元素發 input 事件。
  const fireSafe = async (page) => {
    const names = JSON.parse((await evalJs(`(() => { const seen = {}, out = [];
      document.querySelectorAll('[onclick]').forEach((e) => { const m = /^\\s*([A-Za-z_$][\\w$]*)\\(/.exec(e.getAttribute('onclick') || '');
        if (m && /^(close|toggle|clear|switch|loadLogs|loadContactAudit|loadDeletedContacts)/.test(m[1]) && !seen[m[1]]) { seen[m[1]] = 1; out.push(m[1]); try { e.click(); } catch (x) { /* 例外由 Runtime 事件記錄 */ } } });
      document.querySelectorAll('[oninput]').forEach((e) => { try { e.dispatchEvent(new Event('input', { bubbles: true })); } catch (x) { /* 同上 */ } });
      return JSON.stringify(out); })()`)) || '[]');
    await sleep(800);
    return names;
  };
  const dom = async (page) => {
    metrics[page] = { elements: await evalJs('document.querySelectorAll("*").length'), title: await evalJs('document.title'), path: await evalJs('location.pathname') };
  };

  // 1) login 頁
  await nav('/login.html', 3000);
  await dom('login'); await collect('login');
  // 2) 以管理員實際登入 → admin 後台逐區塊
  const p1 = await login('_s_admin');
  metrics.afterAdminLogin = { path: p1 };
  await collect('login(submit)');
  await nav('/admin.html', 3500);
  await dom('admin'); await collect('admin');
  const secs = await evalJs('JSON.stringify(Array.from(document.querySelectorAll("[data-sec]")).map(e => e.dataset.sec))');
  const secList = JSON.parse(secs || '[]');
  metrics.adminSections = secList.length;
  for (const s of secList) {
    await evalJs(`(() => { const e = document.querySelector('[data-sec="${s}"]'); if (e) e.click(); })()`);
    await sleep(700);
    await collect('admin:' + s);
  }
  metrics.adminFired = await fireSafe('admin'); await collect('admin:handlers');
  await dom('admin(after clicks)');
  // 3) 一般使用者 → index 逐導覽 → help
  await evalJs('fetch("/api/logout", { method: "POST" }).then(() => 1)');
  await sleep(500);
  const p2 = await login('_s_user');
  metrics.afterUserLogin = { path: p2 };
  await collect('login(submit,user)');
  await nav('/index.html', 4000);
  await dom('index'); await collect('index');
  const navs = JSON.parse((await evalJs('JSON.stringify(Array.from(document.querySelectorAll(".sidebar-nav-item[data-section]")).filter(e => e.offsetParent !== null).map(e => e.dataset.section))')) || '[]');
  metrics.indexNavs = navs.length;
  for (const s of navs) {
    await evalJs(`(() => { const e = document.querySelector('.sidebar-nav-item[data-section="${s}"]'); if (e) e.click(); })()`);
    await sleep(700);
    await collect('index:' + s);
  }
  metrics.indexFired = await fireSafe('index'); await collect('index:handlers');
  await dom('index(after clicks)');
  await nav('/help.html', 2500);
  await dom('help'); await collect('help');

  // 4) 等報告送達（Reporting API 會批次延遲送出）；輪詢直到伺服器端數字穩定
  let lastN = -1, stable = 0, waited = 0;
  await http(null, 'POST', '/api/login', { username: '_s_admin', password: H.PW }).then(r => { H.jar._s_admin = r.setCookie.map(x => x.split(';')[0]).join('; '); });
  const serverRows = async () => (await http('_s_admin', 'GET', '/api/admin/logs?limit=1000&action=CSP_VIOLATION&paged=1')).json;
  let rowsJson = null;
  let firstRowAt = null;
  while (waited < 150000 && stable < 6) {
    await sleep(5000); waited += 5000;
    rowsJson = await serverRows();
    const n = (rowsJson && rowsJson.rows ? rowsJson.rows.length : 0) + reportPosts.length;
    if (n > 0 && firstRowAt == null) firstRowAt = waited;
    if (n === lastN && n > 0) stable++; else if (n !== lastN) { stable = 0; lastN = n; }   // 還沒收到第一筆時不算「穩定」，最久等 150 秒
    if (!env.CSP_REPORT_ONLY || env.CSP_REPORT_ONLY === 'off') break;
  }
  const summary = (await http('_s_admin', 'GET', '/api/admin/csp-summary')).json;
  c.close(); killEdge(); await rmDirSoon(edgeDir); edgeDir = null;
  return { label, env, events, metrics, exceptions, logErrors: logErrors.slice(0, 20), logErrorCount: logErrors.length, reportPosts, server: rowsJson && rowsJson.rows || [], summary, waitedMs: waited, firstRowAtMs: firstRowAt };
}

(async () => {
  H.guardClean();
  if (!EDGE) { console.log('找不到 Edge（可設 EDGE_PATH）'); process.exit(2); }
  const t0 = Date.now();
  const timer = setTimeout(() => { console.log('超過時間上限，結束'); killEdge(); H.finish().finally(() => process.exit(3)); }, HARD_LIMIT_MS);
  const results = [];
  try {
    if (!process.env.CSP_SMOKE_ONLY || process.env.CSP_SMOKE_ONLY === 'off') results.push(await runPass('off', { CSP_REPORT_ONLY: 'off' }));
    if (!process.env.CSP_SMOKE_ONLY || process.env.CSP_SMOKE_ONLY === 'on') results.push(await runPass('on', { CSP_REPORT_ONLY: 'on', CSP_REPORT_TO: process.env.CSP_SMOKE_REPORT_TO || 'off', CSP_REPORT_IP_PER_MIN: '100000', CSP_REPORT_GLOBAL_PER_MIN: '100000', CSP_REPORT_MAX_WRITES_PER_HOUR: '100000' }));
    if (process.env.CSP_SMOKE_STYLE) results.push(await runPass('on+style', { CSP_REPORT_ONLY: 'on', CSP_REPORT_TO: process.env.CSP_SMOKE_REPORT_TO || 'off', CSP_REPORT_ONLY_STYLE: 'on', CSP_REPORT_IP_PER_MIN: '100000', CSP_REPORT_GLOBAL_PER_MIN: '100000', CSP_REPORT_MAX_WRITES_PER_HOUR: '100000' }));
  } finally {
    clearTimeout(timer);
    killEdge(); if (edgeDir) await rmDirSoon(edgeDir);
    await finish();
  }
  fs.writeFileSync(OUT, JSON.stringify(results, null, 2));
  console.log('完成，耗時', Math.round((Date.now() - t0) / 1000), '秒；結果已寫入', OUT);
})().catch(async (e) => { console.error('煙霧測試例外：', e); killEdge(); try { await finish(); } catch { /* 忽略 */ } process.exit(1); });
