/* ════════════════════════════════════════════════════════════
 *  閒置自動登出（前端）— window.ITTSIdle
 *
 *  伺服器端（middleware/jwtSession.js）才是權威：最後一次「使用者操作」後超過 30 分鐘沒有操作，
 *  登入 cookie 就失效。這支檔案做三件事：
 *   1. 偵測使用者操作（滑鼠、鍵盤、觸控、捲動），最多每 TOUCH_EVERY_MS 呼叫一次 /api/session/touch
 *      讓伺服器延長閒置視窗（背景輪詢 poll-bundle 不算操作，所以光開著畫面不會一直續命）。
 *   2. 閒置快到時（距離逾時 WARN_BEFORE_MS）跳提醒（沿用 index.html 的 #idleTimeoutOverlay，倒數＋「繼續使用／立即登出」；
 *      沒有這個元素的頁面如 admin.html 改用簡易提示），按「繼續使用」就續期——避免正在填的表單
 *      （例如報價單）到存檔時才發現已登出而遺失內容。逾時時會呼叫 /api/logout（留下 LOGOUT 稽核並清 cookie）再導回登入頁。
 *   3. 攔截 fetch：只要回應帶 X-Session-Expired 標頭（登入失效的 401），就導回登入頁；
 *      其他 401（例如舊密碼不正確）沒有這個標頭，不受影響。
 *
 *  全部 try/catch 包起來：這支檔案壞了不能影響原本功能（最壞情況＝回到只有伺服器端逾時）。
 * ════════════════════════════════════════════════════════════ */
(function (global) {
  'use strict';

  var IDLE_MS = 30 * 60 * 1000;        // 與伺服器預設一致；實際值以 /api/session/touch 回的 idleMs 為準
  var TOUCH_EVERY_MS = 60 * 1000;      // 有操作時，最多每分鐘通知伺服器一次（伺服器端的誤差因此不超過 1 分鐘）
  var WARN_BEFORE_MS = 2 * 60 * 1000;  // 逾時前 2 分鐘提醒
  var TICK_MS = 1000;                  // 提醒顯示時要每秒更新倒數；平常只是比較兩個時間，成本可忽略

  // data-ignore-expired="1"（admin.html 用）：一般的「沒登入」401（X-Session-Expired: 1）不導向，交給頁面自己的登入畫面；
  // 只有「曾經登入、因閒置逾時」（idle）才導回 /login.html。
  var curScript = global.document && global.document.currentScript;
  var opts = { redirectOn401: true, onExpired: null, ignoreExpired: !!(curScript && curScript.getAttribute('data-ignore-expired') === '1') };
  function shouldGo(h) { return !!h && (h === 'idle' || !opts.ignoreExpired); }
  var lastActivity = Date.now();       // 最後一次使用者操作（本頁面的時鐘）
  var lastTouch = Date.now();          // 最後一次成功讓伺服器知道的時間
  var touching = false;
  var started = false;
  var sessionKnown = false;            // 第一次 /api/session/touch 成功（確定有登入）之後才開始計時；沒登入的頁面（admin 的登入畫面等）不倒數、不提醒
  var touchGuard = 0;
  var warnEl = null;
  var expiredFired = false;
  var origFetch = null;

  function now() { return Date.now(); }

  function goLogin(reason) {
    if (expiredFired) return;
    expiredFired = true;
    hideWarn();
    try {
      // 閒置登出（共用電腦的情境）：不暫存目前單據連結，避免下一位登入的人被帶進上一位的單據；
      // 一般的登入失效（401）才暫存，登入後導回原本的單據（與 app.js initUser 的做法一致）
      if (reason === 'idle') { if (typeof global._clearDeepLink === 'function') global._clearDeepLink(); }
      else if (typeof global._rememberDeepLink === 'function') global._rememberDeepLink(true);
    } catch (e) { /* 忽略 */ }
    if (typeof opts.onExpired === 'function') {
      try { opts.onExpired(reason); return; } catch (e) { /* 退回預設 */ }
    }
    var dest = reason === 'idle' ? '/login.html?reason=idle' : '/login.html';   // 只有閒置逾時才帶原因；其餘與 app.js 原本的導向一致
    var go = function () { global.location.href = dest; };
    if (reason === 'idle' && origFetch) {
      // 伺服器的登入可能還有效（用戶端時鐘比伺服器早到期、或 cookie 還沒過期）→ 明確登出：寫稽核、清 cookie
      var done = false;
      var once = function () { if (!done) { done = true; go(); } };
      try { origFetch('/api/logout', { method: 'POST', credentials: 'same-origin', keepalive: true }).then(once, once); } catch (e) { once(); }
      global.setTimeout(once, 1500);
    } else {
      go();
    }
  }

  function installFetchHook() {
    if (origFetch || typeof global.fetch !== 'function') return;
    origFetch = global.fetch.bind(global);
    global.fetch = function () {
      var p = origFetch.apply(global, arguments);
      return p.then(function (res) {
        try {
          if (opts.redirectOn401 && res && res.status === 401) {
            var h = res.headers && res.headers.get && res.headers.get('X-Session-Expired');
            if (shouldGo(h)) goLogin(h === 'idle' ? 'idle' : 'expired');
          }
        } catch (e) { /* 忽略 */ }
        return res;
      });
    };
  }

  function touchServer() {
    if ((touching && now() - touchGuard < 15000) || !origFetch) return;   // 前一次請求卡住超過 15 秒就放行下一次
    touching = true;
    touchGuard = now();
    origFetch('/api/session/touch', { cache: 'no-store', credentials: 'same-origin' })
      .then(function (res) {
        if (res.status === 401) {
          var h = res.headers.get('X-Session-Expired');
          if (opts.redirectOn401 && shouldGo(h)) goLogin(h === 'idle' ? 'idle' : 'expired');
          return null;
        }
        if (!res.ok) return null;   // 403 等（例如強制改密碼期間）：伺服器端的 cookie 仍已延長，不處理
        lastTouch = now();
        sessionKnown = true;
        return res.json().catch(function () { return null; });
      })
      .then(function (j) {
        if (j && Number(j.idleMs) >= 5000) IDLE_MS = Number(j.idleMs);
      })
      .catch(function () { /* 網路暫時不通：下一次操作再試 */ })
      .then(function () { touching = false; });
  }

  // 多個分頁共用「最後操作時間」：只在 A 分頁操作時，B 分頁不能因為自己沒人碰就把使用者登出
  var SHARED_KEY = 'itts_idle_last';
  var lastShared = 0;
  function publishActivity(t) {
    if (t - lastShared < 5000) return;
    lastShared = t;
    try { global.localStorage.setItem(SHARED_KEY, String(t)); } catch (e) { /* 無痕模式等：只用本頁面的時間 */ }
  }
  function sharedActivity() {
    try { var n = Number(global.localStorage.getItem(SHARED_KEY)); return n > 0 && n <= now() + 60000 ? n : 0; } catch (e) { return 0; }
  }

  function inWarn(target) {
    // 事件目標在警告視窗（#idleTimeoutOverlay 或簡易提示）內：這是在操作警告本身，不算「使用者回來了」——
    // 否則 pointerdown 會先把視窗收掉，「立即登出」的 click 永遠發生不了（也不該因此悄悄續命）。
    try {
      var ov = overlayEl();
      return !!(target && ((ov && ov.contains(target)) || (warnEl && warnEl.contains(target))));
    } catch (e) { return false; }
  }

  function onActivity(ev) {
    if (warnShown) {
      if (ev && ev.type === 'keydown' && ev.key === 'Escape') { stay(); return; }   // Esc＝繼續使用
      if (inWarn(ev && ev.target)) return;   // 在警告視窗內操作（Tab、捲動、按鈕）：不續命、不收視窗；「繼續使用」靠自己的 click
    }
    var t = now();
    lastActivity = t;
    publishActivity(t);
    if (warnShown) hideWarn();   // 在視窗外的操作（滑鼠、鍵盤、捲動）＝使用者回來了：收起提醒，操作本身就會續期
    if (t - lastTouch >= TOUCH_EVERY_MS) touchServer();
  }

  var overlayWired = false;
  function overlayEl() { return global.document && global.document.getElementById('idleTimeoutOverlay'); }
  function fmtLeft(sec) { var t = Math.max(0, Math.ceil(sec)); var m = Math.floor(t / 60), r = t % 60; return m + ':' + (r < 10 ? '0' : '') + r; }

  function stay() { lastActivity = now(); publishActivity(lastActivity); lastTouch = 0; hideWarn(); touchServer(); }
  function logoutNow() {
    try { if (typeof global._clearDeepLink === 'function') global._clearDeepLink(); } catch (e) { /* 忽略 */ }
    expiredFired = true;
    var go = function () { global.location.href = '/login.html'; };
    var done = false; var once = function () { if (!done) { done = true; go(); } };
    try { (origFetch || global.fetch)('/api/logout', { method: 'POST', credentials: 'same-origin', keepalive: true }).then(once, once); } catch (e) { once(); }
    global.setTimeout(once, 1500);
  }

  function hideWarn() {
    var ov = overlayEl();
    if (ov) ov.style.display = 'none';
    if (warnEl && warnEl.parentNode) warnEl.parentNode.removeChild(warnEl);
    warnEl = null;
    var wasShown = warnShown;
    warnShown = false;
    // 還原顯示警告前的焦點（正在填的表單欄位）
    if (wasShown && prevFocus) {
      try { if (prevFocus.isConnected && typeof prevFocus.focus === 'function') prevFocus.focus(); } catch (e) { /* 忽略 */ }
    }
    prevFocus = null;
  }
  var prevFocus = null;

  // 焦點陷阱：警告視窗只有兩顆按鈕，Tab／Shift+Tab 在兩者之間循環
  function trapTab(ev) {
    if (!warnShown || ev.key !== 'Tab') return;
    var ov = overlayEl(), box = ov || warnEl;
    if (!box) return;
    var btns = Array.prototype.slice.call(box.querySelectorAll('button')).filter(function (b) { return !b.disabled; });
    if (!btns.length) return;
    var first = btns[0], last = btns[btns.length - 1], a = global.document.activeElement;
    if (!box.contains(a)) { ev.preventDefault(); first.focus(); }
    else if (ev.shiftKey && a === first) { ev.preventDefault(); last.focus(); }
    else if (!ev.shiftKey && a === last) { ev.preventDefault(); first.focus(); }
  }
  var warnShown = false;

  function showWarn(secondsLeft) {
    if (!global.document || !global.document.body) return;
    var firstShow = !warnShown;
    warnShown = true;
    if (firstShow) { try { prevFocus = global.document.activeElement; } catch (e) { prevFocus = null; } }
    var ov = overlayEl();
    if (ov) {
      // 沿用既有的警告視窗（倒數＋繼續使用／立即登出）
      if (!overlayWired) {
        overlayWired = true;
        var sb = global.document.getElementById('idleStayBtn'), lb = global.document.getElementById('idleLogoutBtn');
        if (sb) sb.addEventListener('click', stay);
        if (lb) lb.addEventListener('click', logoutNow);
      }
      var cd = global.document.getElementById('idleCountdown');
      if (cd) cd.textContent = fmtLeft(secondsLeft);
      if (ov.style.display !== 'flex') ov.style.display = 'flex';
      if (firstShow) { var sb2 = global.document.getElementById('idleStayBtn'); if (sb2) { try { sb2.focus(); } catch (e) { /* 忽略 */ } } }   // 焦點進對話框（預設＝繼續使用）
      return;
    }
    // 沒有警告視窗的頁面（admin.html）：簡易提示
    if (!warnEl) {
      warnEl = global.document.createElement('div');
      warnEl.setAttribute('role', 'alertdialog');
      warnEl.setAttribute('aria-label', '閒置逾時警告');
      warnEl.setAttribute('aria-live', 'assertive');
      warnEl.style.cssText = 'position:fixed;left:50%;bottom:24px;transform:translateX(-50%);z-index:2147483000;' +
        'background:#1f2937;color:#fff;padding:12px 16px;border-radius:10px;box-shadow:0 8px 28px rgba(0,0,0,.35);' +
        'font:14px/1.5 system-ui,-apple-system,"Segoe UI","Noto Sans TC",sans-serif;max-width:92vw;display:flex;gap:12px;align-items:center;flex-wrap:wrap';
      var msg = global.document.createElement('span');
      msg.id = 'ittsIdleMsg';
      var btn = global.document.createElement('button');
      btn.type = 'button';
      btn.textContent = '繼續使用';
      btn.style.cssText = 'background:#2563eb;color:#fff;border:0;border-radius:6px;padding:6px 14px;cursor:pointer;font:inherit';
      btn.addEventListener('click', stay);
      warnEl.appendChild(msg);
      warnEl.appendChild(btn);
      global.document.body.appendChild(warnEl);
      try { btn.focus(); } catch (e) { /* 忽略 */ }
    }
    var m = warnEl.querySelector('#ittsIdleMsg');
    if (m) m.textContent = '您已閒置一段時間，系統將在 ' + fmtLeft(secondsLeft) + ' 後自動登出。';
  }

  var sharedCache = 0, sharedAt = 0;
  function tick() {
    if (!sessionKnown || expiredFired) return;
    var t = now();
    if (t - sharedAt >= 5000) { sharedCache = sharedActivity(); sharedAt = t; }   // 別的分頁的操作時間（localStorage），每 5 秒讀一次
    var last = Math.max(lastActivity, sharedCache);
    var left = IDLE_MS - (t - last);
    if (warnShown && left > WARN_BEFORE_MS) hideWarn();   // 別的分頁已續期
    if (left <= 0) { goLogin('idle'); return; }
    if (left <= WARN_BEFORE_MS) showWarn(left / 1000);
  }

  function start(o) {
    if (started) return;
    started = true;
    try {
      if (o) { if (o.redirectOn401 === false) opts.redirectOn401 = false; if (typeof o.onExpired === 'function') opts.onExpired = o.onExpired; }
      installFetchHook();
      ['mousedown', 'keydown', 'touchstart', 'wheel', 'scroll', 'pointerdown'].forEach(function (ev) {
        global.addEventListener(ev, onActivity, { passive: true, capture: true });
      });
      global.addEventListener('keydown', trapTab, true);
      // 分頁從背景切回來：只重新判斷時間（可能已逾時 → 登出）。不替「沒有操作的人」續期——
      // 否則螢幕喚醒、切換分頁都會讓伺服器的最後活動時間往後挪，無人看管的電腦會多開一段時間
      global.document.addEventListener('visibilitychange', function () {
        if (global.document.visibilityState === 'visible') { sharedAt = 0; tick(); }
      });
      global.setInterval(tick, TICK_MS);
      lastTouch = 0; touchServer();   // 載入時問一次：取得伺服器實際的閒置視窗，也順便確認登入還有效
    } catch (e) { /* 忽略：退回只有伺服器端逾時 */ }
  }

  // 預設載入即啟動；<script data-manual> 可改成手動呼叫 ITTSIdle.start()
  global.ITTSIdle = { start: start, _state: function () { return { lastActivity: lastActivity, lastTouch: lastTouch, IDLE_MS: IDLE_MS }; } };
  try {
    var cs = global.document && global.document.currentScript;
    if (global.document && !(cs && cs.hasAttribute('data-manual'))) start();
  } catch (e) { /* 忽略 */ }
})(typeof window !== 'undefined' ? window : this);
