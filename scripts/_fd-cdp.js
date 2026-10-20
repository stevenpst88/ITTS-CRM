'use strict';
/**
 * headless Edge ＋ CDP 小工具（簽核流程設計器 UI 檢查用）。不含任何帳密。
 *  - 每次啟動用臨時 profile（--user-data-dir）；結束以「profile 路徑」找出屬於自己的 msedge.exe 逐一用 PID 結束（不按名稱殺別人的 Edge）
 *  - --disable-gpu；整體時間上限（watchdog），逾時就強制清掉
 *  - 真實的滑鼠／鍵盤輸入（Input.dispatchMouseEvent／dispatchKeyEvent）：拖拉、雙擊、鍵盤操作都走瀏覽器的真實事件流程
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn, spawnSync } = require('child_process');

const sleep = (n) => new Promise((r) => setTimeout(r, n));
function findEdge() {
  return ['C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', 'C:/Program Files/Microsoft/Edge/Application/msedge.exe'].find((p) => fs.existsSync(p));
}
function killByProfile(prof) {
  for (let round = 0; round < 6; round++) {
    const r = spawnSync('powershell', ['-NoProfile', '-Command', `(Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like '*${path.basename(prof)}*' }).ProcessId`], { encoding: 'utf8' });
    const pids = String(r.stdout || '').split(/\s+/).filter((x) => /^\d+$/.test(x));
    if (!pids.length) return 0;
    pids.forEach((pid) => { try { spawnSync('taskkill', ['/PID', pid, '/F'], { stdio: 'ignore' }); } catch { /* 已結束 */ } });
    spawnSync('powershell', ['-NoProfile', '-Command', 'Start-Sleep -Milliseconds 600']);
  }
  const r = spawnSync('powershell', ['-NoProfile', '-Command', `(Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like '*${path.basename(prof)}*' }).ProcessId`], { encoding: 'utf8' });
  return String(r.stdout || '').split(/\s+/).filter((x) => /^\d+$/.test(x)).length;
}

async function withBrowser(opts, fn) {
  opts = Object.assign({ cdpPort: 9461, width: 1440, height: 1000, maxMs: 600000 }, opts || {});
  const edge = findEdge();
  if (!edge) throw new Error('找不到 Edge');
  const prof = fs.mkdtempSync(path.join(os.tmpdir(), 'fdedge-'));
  const ep = spawn(edge, ['--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-features=msEdgeSidebar,msUndersideButton', `--remote-debugging-port=${opts.cdpPort}`, `--user-data-dir=${prof}`, `--window-size=${opts.width},${opts.height}`, 'about:blank'], { stdio: 'ignore' });
  let dog = setTimeout(() => { console.log('逾時看門狗：強制清掉 Edge'); killByProfile(prof); process.exit(3); }, opts.maxMs);
  let leftover = 0;
  try {
    let targets;
    for (let i = 0; i < 60; i++) { try { targets = await (await fetch(`http://127.0.0.1:${opts.cdpPort}/json`)).json(); if (targets.some((x) => x.type === 'page')) break; } catch { /* 等待 */ } await sleep(300); }
    const page = targets.find((x) => x.type === 'page' && /^about:blank/.test(x.url)) || targets.find((x) => x.type === 'page');
    const ws = new WebSocket(page.webSocketDebuggerUrl);
    await new Promise((res, rej) => { ws.onopen = res; ws.onerror = rej; });
    let seq = 0; const waiting = {}; const handlers = [];
    const errors = [], logs = [];
    ws.onmessage = (m) => {
      const d = JSON.parse(m.data);
      if (d.id && waiting[d.id]) { waiting[d.id](d); delete waiting[d.id]; }
      else if (d.method) {
        if (d.method === 'Runtime.exceptionThrown') errors.push(((d.params.exceptionDetails.exception && d.params.exceptionDetails.exception.description) || d.params.exceptionDetails.text || '').slice(0, 400));
        if (d.method === 'Runtime.consoleAPICalled') { const t = d.params.type; const txt = (d.params.args || []).map((a) => a.value !== undefined ? String(a.value) : (a.description || '')).join(' ').slice(0, 300); logs.push(t + ': ' + txt); if (t === 'error') errors.push('console.error: ' + txt); }
        handlers.forEach((hh) => hh(d));
      }
    };
    const send = (method, params) => new Promise((r) => { const id = ++seq; waiting[id] = r; ws.send(JSON.stringify({ id, method, params: params || {} })); });
    const ev = async (expr) => {
      const r = await send('Runtime.evaluate', { expression: expr, awaitPromise: true, returnByValue: true });
      if (!r.result) throw new Error('CDP 錯誤：' + JSON.stringify(r.error || r).slice(0, 300));
      if (r.result.exceptionDetails) throw new Error('頁面例外：' + ((r.result.exceptionDetails.exception && r.result.exceptionDetails.exception.description) || r.result.exceptionDetails.text).slice(0, 400));
      return r.result.result.value;
    };
    const waitFor = async (expr, ms) => { const end = Date.now() + (ms || 8000); while (Date.now() < end) { try { if (await ev(expr)) return true; } catch { /* 頁面切換中 */ } await sleep(120); } return false; };
    const nav = async (url) => { await send('Page.enable'); await send('Page.navigate', { url }); await sleep(600); await waitFor('document.readyState==="complete"', 20000); };
    await send('Runtime.enable');
    await send('Page.addScriptToEvaluateOnNewDocument', { source: 'try { delete Navigator.prototype.serviceWorker; } catch (e) {}' });
    await send('Network.enable'); await send('Network.setCacheDisabled', { cacheDisabled: true }); await send('Network.setBypassServiceWorker', { bypass: true });
    await send('Emulation.setDeviceMetricsOverride', { width: opts.width, height: opts.height, deviceScaleFactor: 1, mobile: false });
    const viewport = async (w, h, mobile) => { await send('Emulation.setDeviceMetricsOverride', { width: w, height: h, deviceScaleFactor: 1, mobile: !!mobile }); await sleep(300); };
    const shot = async (file) => { const r = await send('Page.captureScreenshot', { format: 'png' }); fs.mkdirSync(path.dirname(file), { recursive: true }); fs.writeFileSync(file, Buffer.from(r.result.data, 'base64')); return file; };
    // ── 真實輸入 ──
    let mx = 0, my = 0;
    const mouse = async (type, x, y, extra) => { mx = x; my = y; await send('Input.dispatchMouseEvent', Object.assign({ type, x, y, button: type === 'mouseMoved' ? 'none' : 'left', buttons: type === 'mouseReleased' || type === 'mouseMoved' ? 0 : 1, clickCount: 1, pointerType: 'mouse' }, extra || {})); };
    const move = (x, y) => mouse('mouseMoved', x, y);
    const down = (x, y) => mouse('mousePressed', x, y);
    const up = (x, y) => mouse('mouseReleased', x, y);
    const rectOf = async (sel) => ev(`(()=>{const e=${typeof sel === 'string' && /^\(\)/.test(sel) ? `(${sel})()` : `document.querySelector(${JSON.stringify(sel)})`};if(!e)return null;e.scrollIntoView({block:'center',inline:'center'});const r=e.getBoundingClientRect();return {x:r.left,y:r.top,w:r.width,h:r.height,cx:r.left+r.width/2,cy:r.top+r.height/2}})()`);
    const click = async (sel, off) => { const r = await rectOf(sel); if (!r) throw new Error('找不到元素：' + sel); const x = r.cx + ((off && off.dx) || 0), y = r.cy + ((off && off.dy) || 0); await move(x, y); await down(x, y); await up(x, y); await sleep(60); return r; };
    const dblclick = async (sel) => { const r = await rectOf(sel); if (!r) throw new Error('找不到元素：' + sel); await move(r.cx, r.cy); await send('Input.dispatchMouseEvent', { type: 'mousePressed', x: r.cx, y: r.cy, button: 'left', buttons: 1, clickCount: 1, pointerType: 'mouse' }); await send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: r.cx, y: r.cy, button: 'left', buttons: 0, clickCount: 1, pointerType: 'mouse' }); await send('Input.dispatchMouseEvent', { type: 'mousePressed', x: r.cx, y: r.cy, button: 'left', buttons: 1, clickCount: 2, pointerType: 'mouse' }); await send('Input.dispatchMouseEvent', { type: 'mouseReleased', x: r.cx, y: r.cy, button: 'left', buttons: 0, clickCount: 2, pointerType: 'mouse' }); await sleep(80); return r; };
    const dragTo = async (x0, y0, x1, y1, steps) => { await move(x0, y0); await down(x0, y0); const n = steps || 8; for (let i = 1; i <= n; i++) { await move(x0 + (x1 - x0) * i / n, y0 + (y1 - y0) * i / n); await sleep(15); } await up(x1, y1); await sleep(150); };
    // 拖拉：兩端都要在可視範圍內才下手（必要時先捲動到兩者中間），之後重新量位置，避免捲動造成座標過期
    const rectRaw = (sel) => ev(`(()=>{const e=document.querySelector(${JSON.stringify(sel)});if(!e)return null;const r=e.getBoundingClientRect();return {x:r.left,y:r.top,w:r.width,h:r.height,cx:r.left+r.width/2,cy:r.top+r.height/2,vh:window.innerHeight}})()`);
    const drag = async (fromSel, toSel, off) => {
      let ra = await rectRaw(fromSel), rb = await rectRaw(toSel);
      if (!ra || !rb) throw new Error('拖拉找不到元素：' + fromSel + ' → ' + toSel);
      const inView = (r) => r.cy > 70 && r.cy < r.vh - 30;
      if (!inView(ra) || !inView(rb)) { const mid = (ra.cy + rb.cy) / 2; await ev(`window.scrollBy(0, ${Math.round(mid - ra.vh / 2)})`); await sleep(250); ra = await rectRaw(fromSel); rb = await rectRaw(toSel); }
      await dragTo(ra.cx, ra.cy, rb.cx + ((off && off.dx) || 0), rb.cy + ((off && off.dy) || 0));
    };
    const KEYS = { Enter: { key: 'Enter', code: 'Enter', windowsVirtualKeyCode: 13, text: '\r' }, Escape: { key: 'Escape', code: 'Escape', windowsVirtualKeyCode: 27 }, Delete: { key: 'Delete', code: 'Delete', windowsVirtualKeyCode: 46 }, Tab: { key: 'Tab', code: 'Tab', windowsVirtualKeyCode: 9 }, ArrowLeft: { key: 'ArrowLeft', code: 'ArrowLeft', windowsVirtualKeyCode: 37 }, ArrowUp: { key: 'ArrowUp', code: 'ArrowUp', windowsVirtualKeyCode: 38 }, ArrowRight: { key: 'ArrowRight', code: 'ArrowRight', windowsVirtualKeyCode: 39 }, ArrowDown: { key: 'ArrowDown', code: 'ArrowDown', windowsVirtualKeyCode: 40 }, ' ': { key: ' ', code: 'Space', windowsVirtualKeyCode: 32, text: ' ' }, '+': { key: '+', code: 'Equal', windowsVirtualKeyCode: 187, text: '+' }, '-': { key: '-', code: 'Minus', windowsVirtualKeyCode: 189, text: '-' }, 0: { key: '0', code: 'Digit0', windowsVirtualKeyCode: 48, text: '0' } };
    const key = async (k, mods) => { const d = KEYS[k] || { key: k, code: 'Key' + k.toUpperCase(), windowsVirtualKeyCode: k.toUpperCase().charCodeAt(0), text: k }; await send('Input.dispatchKeyEvent', Object.assign({ type: d.text ? 'keyDown' : 'rawKeyDown', modifiers: mods || 0 }, d)); await send('Input.dispatchKeyEvent', Object.assign({ type: 'keyUp', modifiers: mods || 0 }, d, { text: undefined })); await sleep(60); };
    const type = async (text) => { for (const ch of text) { await send('Input.dispatchKeyEvent', { type: 'keyDown', key: ch, text: ch }); await send('Input.dispatchKeyEvent', { type: 'keyUp', key: ch }); } await sleep(40); };
    const focus = (sel) => ev(`(()=>{const e=document.querySelector(${JSON.stringify(sel)});if(!e)return false;e.focus();return document.activeElement===e})()`);
    // 以「網址片段 → 內容」攔截並替換頁面資源（變異測試：把前端檔換成改壞的版本，不動磁碟上的檔案）
    const replace = {};
    handlers.push(async (d) => {
      if (d.method !== 'Fetch.requestPaused') return;
      const url = d.params.request.url.split('?')[0];
      const k = Object.keys(replace).find((x) => url.endsWith(x));
      if (k) await send('Fetch.fulfillRequest', { requestId: d.params.requestId, responseCode: 200, responseHeaders: [{ name: 'Content-Type', value: k.endsWith('.css') ? 'text/css; charset=utf-8' : 'application/javascript; charset=utf-8' }, { name: 'Cache-Control', value: 'no-store' }], body: Buffer.from(replace[k], 'utf8').toString('base64') });
      else await send('Fetch.continueRequest', { requestId: d.params.requestId });
    });
    const interceptOn = async () => { await send('Fetch.enable', { patterns: [{ urlPattern: '*' }] }); };
    const media = async (features) => send('Emulation.setEmulatedMedia', { features });
    // beforeunload／alert 之類的對話框一律接受（頁面有「未儲存變更」離開提示）
    handlers.push(async (d) => { if (d.method === 'Page.javascriptDialogOpening') await send('Page.handleJavaScriptDialog', { accept: true }); });
    const ctx = { replace, interceptOn, media, send, ev, waitFor, nav, sleep, shot, viewport, move, down, up, click, dblclick, drag, dragTo, rectOf, key, type, focus, errors, logs, port: opts.cdpPort };
    return await fn(ctx);
  } finally {
    clearTimeout(dog);
    leftover = killByProfile(prof);
    await sleep(500);
    try { fs.rmSync(prof, { recursive: true, force: true }); } catch { /* Edge 尚未放開檔案 */ }
    withBrowser.leftover = leftover;
  }
}
module.exports = { withBrowser, findEdge, killByProfile, sleep };
