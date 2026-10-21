'use strict';
/**
 * 閒置登出的「單一來源」（M-01，2026-10 程式碼審查）：
 * 過去有兩套互相獨立的閒置機制——middleware/jwtSession.js（固定 30 分鐘、背景輪詢只認 /api/poll-bundle）與
 * lib/loginGuard.js（SESSION_IDLE_MINUTES、輪詢路徑三個）——實測會互相打架（有效值＝兩者取小、重複簽發 cookie、
 * /api/notifications 在一邊續命另一邊不續命）。現在：
 *   · 閒置長度只有一個解析函式 resolveIdleMs()：SESSION_IDLE_MS（測試用，下限 5 秒）‖ SESSION_IDLE_MINUTES*60000 ‖ 30 分鐘
 *   · 背景輪詢路徑只有一份 IDLE_POLL_PATHS／isPollingRequest()（GET／HEAD 才算輪詢；其餘方法一律算使用者操作）
 *   · JWT 模式（正式站與預設本機）由 jwtSession 一手負責閒置與 8 小時上限；loginGuard.idleMiddleware 只服務 SESSION_BACKEND=memory
 */
const DEFAULT_IDLE_MS = 30 * 60 * 1000;
const MIN_TEST_IDLE_MS = 5000;
const ABSOLUTE_SESSION_MS = 8 * 60 * 60 * 1000;   // 登入後的絕對上限（與 JWT exp 一致；不因續期而延長）

/** 背景自動輪詢端點：不代表使用者還在操作，不可刷新閒置計時 */
const IDLE_POLL_PATHS = ['/api/poll-bundle', '/api/notifications', '/api/announcements/active'];

/** Express 預設不分大小寫、容許結尾斜線：/API/Poll-Bundle/ 一樣會進同一支路由，所以比對前先正規化 */
function isPollingRequest(method, path) {
  const m = String(method == null ? 'GET' : method).toUpperCase();
  if (m !== 'GET' && m !== 'HEAD') return false;
  const p = String(path || '').toLowerCase().replace(/\/+$/, '');
  return IDLE_POLL_PATHS.includes(p);
}

function resolveIdleMs(env) {
  env = env || process.env;
  const t = Number(env.SESSION_IDLE_MS);
  if (Number.isFinite(t) && t >= MIN_TEST_IDLE_MS) return Math.floor(t);
  const m = parseFloat(env.SESSION_IDLE_MINUTES);
  if (Number.isFinite(m) && m > 0) return Math.max(1000, Math.floor(m * 60000));
  return DEFAULT_IDLE_MS;
}

module.exports = { DEFAULT_IDLE_MS, MIN_TEST_IDLE_MS, ABSOLUTE_SESSION_MS, IDLE_POLL_PATHS, isPollingRequest, resolveIdleMs };
