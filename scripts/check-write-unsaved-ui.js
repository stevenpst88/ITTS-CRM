#!/usr/bin/env node
/**
 * 前端「尚未儲存」(503) 處理檢查。用法：node scripts/check-write-unsaved-ui.js（不需要伺服器、不開瀏覽器）
 * 背景：DB_WRITE_SERIALIZE=1 時，鎖等太久／存檔失敗會回 503（DB_BUSY／DB_WRITE_FAILED／QUOTE_NO_UNAVAILABLE，body 帶 retryable:true）。
 * 前端若不看狀態碼就顯示「已儲存」，就把「靜默掉資料」從伺服器搬到畫面上。
 *   1) app.js 的全域 fetch 包裝（在 vm 沙盒裡實跑）：寫入型 503＋retryable → 標記 __unsaved、顯示一次 toast、Response 內容仍可讀；其他情況原樣放行
 *   2) admin.html 的 adminFetch 同樣處理
 *   3) 掃描所有前端檔案的寫入型 fetch：每一處都必須「檢查回應」（isUnsaved／.ok／.success／.status）或在白名單內（登出、通知已讀…），
 *      新增寫入呼叫卻沒處理回應時這條會失敗——提醒你補上 if (isUnsaved(r)) return;
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ROOT = path.join(__dirname, '..');
const CLIENT = path.join(ROOT, '_client');

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
let summaryPrinted = false;
process.on('exit', (code) => {
  if (summaryPrinted) return;
  res.forEach(([n, p, e]) => console.log((p ? 'PASS ' : 'FAIL ') + n + (e && !p ? '  ← ' + e : '')));
  console.log('\n【未完成】測試在印出摘要前就結束了，視為失敗。');
  process.exitCode = code || 3;
});

const mkResponse = (status, body) => new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } });

/** 取出 app.js 中 patchFetch 區段（常數＋IIFE＋isUnsaved），在沙盒裡執行 */
function loadAppPatch() {
  const src = fs.readFileSync(path.join(CLIENT, 'app.js'), 'utf8');
  const a = src.indexOf('const _ACCESS_VIEW_EXEMPT_CLIENT');
  const b = src.indexOf('function applyPermissions()');
  if (a < 0 || b < 0) throw new Error('找不到 patchFetch 區段');
  const toasts = [];
  let next = () => mkResponse(200, { ok: 1 });
  const win = { fetch: (url, opts) => Promise.resolve(next(url, opts)) };
  const ctx = vm.createContext({ window: win, Response, userPermissions: { accessMode: 'edit' }, showToast: (m) => toasts.push(m), Date, JSON, Promise, String });
  vm.runInContext(src.slice(a, b), ctx);
  return { win, toasts, setNext: (f) => { next = f; }, isUnsaved: vm.runInContext('isUnsaved', ctx), ctx };
}

(async () => {
  // ═══ 1) app.js 全域 fetch 包裝 ═══
  {
    const app = loadAppPatch();
    app.setNext(() => mkResponse(503, { code: 'DB_BUSY', retryable: true, error: '系統目前較忙，您的操作尚未儲存。' }));
    const r = await app.win.fetch('/api/opportunities', { method: 'POST', body: '{}' });
    t('1.1 寫入型請求收到 503＋retryable → Response 標記為尚未儲存', app.isUnsaved(r) === true && r.status === 503);
    t('1.2 顯示 toast（含伺服器的友善訊息）', app.toasts.length === 1 && app.toasts[0].includes('尚未儲存'), app.toasts.join('|'));
    const body = await r.json();
    t('1.3 Response 內容仍可讀（原本檢查 r.ok／讀 j.error 的呼叫點不受影響）', body.code === 'DB_BUSY');
    await app.win.fetch('/api/visits/1', { method: 'PUT', body: '{}' });
    t('1.4 短時間連續兩次 503 只跳一次 toast（不洗版）', app.toasts.length === 1, app.toasts.length);
    const g = await app.win.fetch('/api/opportunities');
    t('1.5 GET 收到 503 不標記、不跳 toast（只處理寫入）', app.isUnsaved(g) === false && app.toasts.length === 1);
    app.setNext(() => mkResponse(503, { error: '資料庫初始化失敗，請稍後再試' }));
    const p2 = await app.win.fetch('/api/contacts', { method: 'POST', body: '{}' });
    t('1.6 503 但伺服器沒標 retryable → 不當作「尚未儲存」', app.isUnsaved(p2) === false && app.toasts.length === 1);
    const okResp = mkResponse(200, { ok: 1 });
    app.setNext(() => okResp);
    const p3 = await app.win.fetch('/api/contacts', { method: 'POST', body: '{}' });
    t('1.7 成功回應原樣放行（同一個物件、沒被改動）', p3 === okResp && app.isUnsaved(p3) === false);
    const p4 = await app.win.fetch('/api/contacts/1', { method: 'DELETE' });
    t('1.8 DELETE 成功（200）不標記', app.isUnsaved(p4) === false);   // 200 不標記
    app.setNext(() => mkResponse(503, { code: 'DB_WRITE_FAILED', retryable: true, error: '資料暫時無法儲存' }));
    await new Promise((r2) => setTimeout(r2, 1600));   // 超過 toast 去重時間
    const p5 = await app.win.fetch('/api/contacts/1', { method: 'DELETE' });
    t('1.9 DELETE 的 DB_WRITE_FAILED 同樣標記並再次提示（超過去重時間）', app.isUnsaved(p5) === true && app.toasts.length === 2, app.toasts.length);
    const ctx2 = loadAppPatch();
    ctx2.ctx.userPermissions.accessMode = 'view';
    const ro = await ctx2.win.fetch('/api/contacts', { method: 'POST', body: '{}' });
    t('1.10 唯讀模式的原有行為不變（403 攔截）', ro.status === 403);
  }

  // ═══ 2) admin.html 的 adminFetch／markUnsaved ═══
  {
    const html = fs.readFileSync(path.join(CLIENT, 'admin.html'), 'utf8');
    const a = html.indexOf('let _unsavedToastAt = 0;');
    const b = html.indexOf('// ── 載入使用者列表 ──');
    if (a < 0 || b < 0) throw new Error('找不到 adminFetch 區段');
    const toasts = [];
    let next = () => mkResponse(200, {});
    const ctx = vm.createContext({ fetch: (u, o) => Promise.resolve(next(u, o)), Response, showToast: (m) => toasts.push(m), showLoginScreen() {}, $: () => ({ textContent: '', classList: { add() {} } }), Date, Promise, JSON, String, Error });
    vm.runInContext(html.slice(a, b), ctx);
    next = () => mkResponse(503, { code: 'DB_BUSY', retryable: true, error: '尚未儲存，請稍候' });
    const r = await vm.runInContext("adminFetch('/api/admin/users', { method: 'POST', body: '{}' })", ctx);
    const un = vm.runInContext('isUnsaved', ctx);
    t('2.1 adminFetch：寫入型 503＋retryable → 標記並 toast', un(r) === true && toasts.length === 1);
    const g = await vm.runInContext("adminFetch('/api/admin/users')", ctx);
    t('2.2 adminFetch：GET 不處理', un(g) === false);
  }

  // ═══ 3) 掃描所有前端寫入型 fetch：每一處都要檢查回應或在白名單 ═══
  {
    // 白名單：明確不需要處理回應的呼叫（fire-and-forget／登出／通知已讀／上傳（豁免寫入鎖）／已用其他方式處理）
    const ALLOW = [
      '/api/logout', '/logout', '/notifications/', '/announcements/', '/api/push/', '/upload`', '/api/upload', '/export-log',   // export-log：只寫稽核（豁免寫入鎖）、fire-and-forget
    ];
    const files = fs.readdirSync(CLIENT).filter((f) => /\.(js|html)$/.test(f));
    const offenders = [];
    let total = 0;
    for (const f of files) {
      const L = fs.readFileSync(path.join(CLIENT, f), 'utf8').split('\n');
      for (let i = 0; i < L.length; i++) {
        if (!/method:\s*['"`](POST|PUT|PATCH|DELETE)['"`]|method:\s*[a-z]+\s*\?\s*['"`](PUT|POST)/.test(L[i])) continue;
        let s = i; while (s > i - 5 && !/[fF]etch\(/.test(L[s])) s--;
        if (!/[fF]etch\(/.test(L[s])) continue;
        total++;
        const head = L.slice(s, s + 3).join(' ');
        if (ALLOW.some((a) => head.includes(a))) continue;
        // 回應有沒有被拿去用：fetch 的結果指派給變數、且該變數在後續 20 行內被引用（檢查 .ok／isUnsaved／讀 json…）；包在 Promise.all／return／.then 裡也算；
        // 整個丟棄（await fetch(...); 然後直接顯示成功）＝這次要抓的假成功，列為違規
        let v = null; let mm;
        if ((mm = /(?:const|let|var)\s+(\w+)\s*=\s*(?:await\s+)?(?:markUnsaved\(\s*(?:await\s+)?)?(?:\w+\.)?(?:adminFetch|fetch)\(/.exec(L[s]))) v = mm[1];
        else if ((mm = /^\s*(\w+)\s*=\s*await\s+(?:adminFetch|fetch)\(/.exec(L[s]))) v = mm[1];
        const prev = L.slice(Math.max(0, s - 2), s + 1).join(' ');
        const consumed = v ? new RegExp('\\b' + v + '\\b').test(L.slice(s + 1, s + 21).join('\n')) : /Promise\.all|return\s+(?:await\s+)?fetch|\.then\(/.test(prev);
        if (consumed) continue;
        offenders.push(f + ':' + (s + 1) + ' ' + L[s].trim().slice(0, 80));
      }
    }
    t('3.1 前端所有寫入型 fetch（共 ' + total + ' 處）都檢查了回應或在白名單內', offenders.length === 0, offenders.join(' ; '));
    const app = fs.readFileSync(path.join(CLIENT, 'app.js'), 'utf8');
    const names = ['_rNewOpp', '_rVisitPut', '_rVisitOpp', '_rVisitDel', '_rOppPut1', '_rOppPut2', '_rTarget', '_rContractDel', '_rRecvDel'];
    t('3.2 先前直接丟棄回應的 9 個呼叫點都已加 isUnsaved 檢查', names.every((n) => app.includes('isUnsaved(' + n + ')')), names.filter((n) => !app.includes('isUnsaved(' + n + ')')).join(','));
    const adm = fs.readFileSync(path.join(CLIENT, 'admin.html'), 'utf8');
    t('3.3 admin.html 直接丟棄回應的呼叫點（還原／永久刪除／清除個人季比）已加檢查', ['_rRestore', '_rPerm', '_rClr'].every((n) => adm.includes('isUnsaved(' + n + ')')));
  }

  // ═══ 4) 第二輪審查（R2）的前端修正 ═══
  {
    const app = fs.readFileSync(path.join(CLIENT, 'app.js'), 'utf8');
    // 4.1~4.2 showToast：新訊息要取消前一個計時器（R2-#13）
    {
      const a = app.indexOf('var _toastTimer = null;');
      const b = app.indexOf('function getInitial(');
      if (a < 0 || b < 0) throw new Error('找不到 showToast 區段');
      const timers = []; let nextId = 1;
      const el = { textContent: '', classList: { shown: false, add() { this.shown = true; }, remove() { this.shown = false; } } };
      const ctx = vm.createContext({
        $: () => el,
        setTimeout: (fn, ms) => { const id = nextId++; timers.push({ id, fn, ms, cleared: false }); return id; },
        clearTimeout: (id) => { const t = timers.find((x) => x.id === id); if (t) t.cleared = true; },
      });
      vm.runInContext(app.slice(a, b), ctx);
      const show = vm.runInContext('showToast', ctx);
      show('通用提示', 6500);
      show('重點提示', 8000);
      t('4.1 連續兩則 toast：前一則的計時器被取消（重點提示不會被前一則的 6.5 秒計時提早收掉）', timers.length === 2 && timers[0].cleared === true && timers[1].cleared === false && el.textContent === '重點提示');
      // 被取消的計時器不會觸發；第二個計時器到期才收掉
      el.classList.shown = true;
      timers[1].fn();
      t('4.2 第二則自己的計時器到期才收起 toast', el.classList.shown === false);
    }
    // 4.3~4.4 商機階段下拉：503 後要重畫表格讓下拉回到舊值（R2-#11）
    {
      const a = app.indexOf('async function updateOppStage(');
      const b = app.indexOf('\n}\n', a);
      if (a < 0 || b < 0) throw new Error('找不到 updateOppStage');
      const calls = [];
      const mkCtx = (unsaved) => vm.createContext({
        API: '/api', allOpportunities: [], currentSection: 'pipeline',
        confirm: () => true,
        fetch: async () => ({ __unsaved: unsaved, ok: !unsaved }),
        isUnsaved: (r) => !!(r && r.__unsaved),
        loadOpportunities: async () => { calls.push('loadOpportunities'); },
        renderOppTable: () => { calls.push('renderOppTable'); },
        renderTargetHistory() {}, updateTargetCard() {}, updateStatCards() {}, renderDashboardCharts() {}, loadContacts() {},
        showToast: (m) => { calls.push('toast'); }, celebrateWon() {}, promptConvertToCustomer() {}, setTimeout, JSON, Date,
      });
      const c1 = mkCtx(true);
      vm.runInContext(app.slice(a, b + 3), c1);
      await vm.runInContext("updateOppStage('o1', 'A')", c1);
      t('4.3 階段下拉收到 503（尚未儲存）→ 立刻重畫表格（下拉回到實際的舊值），不重新載入、不跳成功訊息', calls.join() === 'renderOppTable', calls.join());
      calls.length = 0;
      const c2 = mkCtx(false);
      vm.runInContext(app.slice(a, b + 3), c2);
      await vm.runInContext("updateOppStage('o1', 'A')", c2);
      t('4.4 成功時流程不變（loadOpportunities → renderOppTable → 成功訊息）', calls.join() === 'loadOpportunities,renderOppTable,toast', calls.join());
    }
    // 4.6 SAP 批次推送：伺服器回 saved:false → 後台明確提示「請勿重推」（R2-#1）
    {
      const adm = fs.readFileSync(path.join(CLIENT, 'admin.html'), 'utf8');
      t('4.6 admin.html 的 SAP 推送：回應 saved:false 時顯示「SAP 已推送，但本機對照寫入失敗，請勿重推，請聯絡 IT」（不是顯示推送完成）', adm.includes('d.saved === false') && adm.includes('本機對照寫入失敗') && adm.includes('請勿重推，請聯絡 IT') && adm.indexOf('if (linkUnsaved) {') < adm.indexOf('✅ 推送完成'));
    }
    // 4.5 通知已讀：503 不假裝已讀（R2-#12）
    {
      t('4.5 通知「已讀」「全部已讀」的 PUT 檢查 isUnsaved：503 時不移除未讀樣式、不重抓（避免閃一下又跳回未讀）', app.includes('isUnsaved(_rRead)') && app.includes('isUnsaved(_rReadAll)'));
    }
  }

  let bad = 0;
  for (const [n, p, e] of res) { if (!p) bad++; console.log((p ? 'PASS ' : 'FAIL ') + n + (e && !p ? '  ← ' + e : '')); }
  console.log(`\n${res.length - bad}/${res.length} 通過` + (bad ? `，${bad} 失敗` : ''));
  summaryPrinted = true;
  process.exit(bad ? 1 : 0);
})().catch((e) => { console.error('測試腳本例外：', e); process.exit(2); });
