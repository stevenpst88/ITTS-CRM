/**
 * _client/deep-link.js — 信件深層連結的暫存小工具（瀏覽器端；P3 純邏輯，尚未接線）
 *
 * 用途：信內連結 → /q/<id> 跳板頁 → /index.html#quote:<id>。若使用者尚未登入，/index.html 會被導向 /login.html，
 * 而 JS 導向（app.js 的 401 分支）與登入成功後的固定導向（login.html 登入後 location.href = '/'）都會丟掉 # 片段。
 * 這支檔案把「合法的 #quote:… 片段」暫存在 sessionStorage，登入後再取出帶回去。
 *
 * 掛載：window.ITTSDeepLink；同時支援 module.exports（Node 測試用）。沒有任何相依。
 *   remember(hashOrLocation?) → boolean   片段符合格式才存（存成功回 true）。參數可以是字串（'#quote:…'）、帶 .hash 的物件（window.location），
 *                                         省略時讀 window.location.hash。格式不符或 sessionStorage 不可用 → false，不 throw，也不會清掉先前已記住的值。
 *   rememberForLogin(h)       → boolean   同 remember，另外「上膛」（itts.deepLink.armed）：給 app.js 在「為了登入而導向 /login.html」之前用（401 分支）。
 *                                         上膛的暫存只會被「下一次」載入的登入頁保留，一次性。
 *   settleOnLoginPage(h?)     → 'hash'|'kept'|'cleared'|'none'   給 login.html 載入時呼叫，決定要不要沿用暫存：
 *                                         網址上有合法 #quote:… → 記住它（'hash'，覆蓋舊值）；沒有，但暫存是上膛的（app.js 為了登入而導過來）→ 保留（'kept'）；
 *                                         都不是（使用者主動打開登入頁、登出後回到登入頁、強制改密碼畫面登出…）→ 清掉舊暫存（'cleared'），
 *                                         所以下一位登入的人不會被帶去上一位的單據。上膛旗標每次載入都會被取走（一次性）。
 *   clear()                   → boolean   清掉暫存（含上膛旗標）。所有登出路徑（一般登出、強制改密碼畫面的登出、閒置自動登出）都呼叫它。
 *   consume()                 → string    取出並清除（只能取一次，上膛旗標一併清掉）；沒有或內容不合格 → ''。儲存內容被竄改（不符格式）也回 '' 並清掉。
 *   hashForRedirect()         → string    給「登入成功後的導向」用：回傳可直接接在網址後面的片段（'' 或 '#quote:…'），語意同 consume（取一次即清），
 *                                         避免同一分頁下次登入又被帶去舊單據。
 *   isValidHash(h)            → boolean   格式檢查（給 app.js 等處共用，避免各寫一份正規式）。
 *   STORAGE_KEY               = 'itts.deepLink'；ARMED_KEY = 'itts.deepLink.armed'
 *
 * 格式：^#quote:[A-Za-z0-9_-]{1,64}(:cost)?$  ——與 lib/mail/link.js 的 deepHash／events.QUOTE_ID_RE 一致。
 *   不接受其他格式（防 open redirect／注入）：結果只會是 '#quote:' 開頭的短字串，不可能含 / : . ? # 空白或換行。
 *   注意：_client/app.js:8540 現行的 _handleQuoteDeepLink 只接受 ^#quote:([0-9a-fA-F-]{8,64})(:cost)?$（uuid）。
 *   單據 id 目前全是 uuid 所以兩者一致；整合時若要支援非 uuid 的 id，app.js 的正規式需同步放寬（或改呼叫 isValidHash）。
 *
 * 失敗模式（全部 try/catch，不影響頁面）：sessionStorage 在隱私模式／封鎖 cookie／沙箱 iframe 可能在「存取」時就丟 SecurityError，
 *   setItem 可能丟 QuotaExceededError。remember 失敗回 false；consume 讀不到回 ''；
 *   若讀到值但 removeItem 失敗，仍回傳該值（盡力而為，下次可能再取到同一個值）。
 *
 * 已接進 login.html（載入時 settleOnLoginPage、登入成功 hashForRedirect）與 app.js（401 分支 rememberForLogin、強制改密碼分支 remember、各登出路徑 clear）。
 * 殘留防護（FIX-4）：沒有上膛、也沒有網址片段的登入頁載入一律清掉暫存；各登出路徑也明確 clear()，兩層都擋。
 */
(function () {
  'use strict';

  var STORAGE_KEY = 'itts.deepLink';
  var ARMED_KEY = 'itts.deepLink.armed';
  var MAX_LEN = 80;                                   // 合法片段最長 7+64+1+4+... 遠小於此；先擋長度再跑正規式
  var HASH_RE = /^#quote:[A-Za-z0-9_-]{1,64}(?::cost)?$/;

  function isValidHash(h) {
    return typeof h === 'string' && h.length <= MAX_LEN && HASH_RE.test(h);
  }

  function pickHash(arg) {
    try {
      if (typeof arg === 'string') return arg;
      if (arg && typeof arg === 'object') return typeof arg.hash === 'string' ? arg.hash : '';
      if (arg === undefined && typeof window !== 'undefined' && window.location && typeof window.location.hash === 'string') {
        return window.location.hash;
      }
    } catch (e) { /* 取不到就當沒有 */ }
    return '';
  }

  function getStorage() {
    try {
      if (typeof window === 'undefined') return null;
      return window.sessionStorage || null;           // 存取 sessionStorage 本身就可能丟 SecurityError
    } catch (e) {
      return null;
    }
  }

  function remember(hashOrLocation) {
    var h = pickHash(hashOrLocation);
    if (!isValidHash(h)) return false;
    var s = getStorage();
    if (!s) return false;
    try {
      s.setItem(STORAGE_KEY, h);
      return true;
    } catch (e) {
      return false;
    }
  }

  function rememberForLogin(hashOrLocation) {
    if (!remember(hashOrLocation)) return false;
    var s = getStorage();
    try { if (s) s.setItem(ARMED_KEY, '1'); } catch (e) { /* 上膛失敗：登入頁會把暫存當成「非 app 導過來的」而清掉，只是少了一次帶回 */ }
    return true;
  }

  function clear() {
    var s = getStorage();
    if (!s) return false;
    var ok = true;
    try { s.removeItem(STORAGE_KEY); } catch (e) { ok = false; }
    try { s.removeItem(ARMED_KEY); } catch (e) { ok = false; }
    return ok;
  }

  function consume() {
    var s = getStorage();
    if (!s) return '';
    var v = '';
    try { v = s.getItem(STORAGE_KEY); } catch (e) { return ''; }
    try { s.removeItem(STORAGE_KEY); } catch (e) { /* 清不掉就算了 */ }
    try { s.removeItem(ARMED_KEY); } catch (e) { /* 同上 */ }
    return isValidHash(v) ? v : '';
  }

  function settleOnLoginPage(hashOrLocation) {
    var h = pickHash(hashOrLocation);
    var s = getStorage();
    var armed = false;
    if (s) {
      try { armed = s.getItem(ARMED_KEY) === '1'; } catch (e) { armed = false; }
      try { s.removeItem(ARMED_KEY); } catch (e) { /* 一次性旗標清不掉就算了 */ }
    }
    if (isValidHash(h)) return remember(h) ? 'hash' : 'none';
    if (armed) {
      var v = '';
      try { v = s.getItem(STORAGE_KEY); } catch (e) { v = ''; }
      if (isValidHash(v)) return 'kept';
    }
    return s ? (clear(), 'cleared') : 'none';
  }

  function hashForRedirect() {
    return consume();
  }

  var api = {
    remember: remember,
    rememberForLogin: rememberForLogin,
    settleOnLoginPage: settleOnLoginPage,
    clear: clear,
    consume: consume,
    hashForRedirect: hashForRedirect,
    isValidHash: isValidHash,
    STORAGE_KEY: STORAGE_KEY,
    ARMED_KEY: ARMED_KEY,
  };

  if (typeof window !== 'undefined') window.ITTSDeepLink = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})();
