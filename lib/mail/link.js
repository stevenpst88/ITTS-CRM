'use strict';
/**
 * lib/mail/link.js — 信內連結與「跳板頁」（P3 純邏輯，不接線；本檔不碰伺服器、不查資料庫）
 *
 * 匯出與簽名：
 *   buildQuoteLink(config, quoteId, {cost}?) → string   `${appBaseUrl}/q/${id}` ，填成本再加 `?cost=1`；單據 id 或網站根網址不合法 → throw MailLinkError
 *   deepHash(quoteId, {cost}?)               → string   '#quote:<id>' 或 '#quote:<id>:cost'（與 _client/deep-link.js 同格式）；id 不合法 → throw
 *   renderJumpPage({quoteId, cost})          → {status, headers, body}   永不 throw
 *   isValidQuoteId(id)                       → boolean   直接引用 events.js（同一份正規式，避免兩份漂移）
 *   isCostFlag(v)                            → boolean   只有 true／1／'1' 算「填成本」
 *   MailLinkError                            extends Error，帶 .code（BAD_QUOTE_ID／BAD_BASE_URL）
 *   JUMP_PATH_PREFIX                         = '/q/'
 *
 * 為什麼需要跳板頁（背景：exchange_codemap §4.2）：
 *   登入 cookie 是 SameSite=Strict（middleware/jwtSession.js:50-58）。使用者在 Outlook／Teams 點信內連結時，
 *   第一個導覽是「跨站發起」，瀏覽器不會帶 cookie，就算 8 小時內登入過也會被當成沒登入。
 *   做法：信內連結指向「不需登入」的 /q/<id>，回一張極小的 200 HTML 頁；這張頁面由「我們自己的網站」再導向 /index.html#quote:<id>，
 *   第二次導覽的發起者是本站，cookie 才會帶上。注意：伺服器端的 302 沒有用（整條重新導向鏈仍算跨站發起），所以必須是 200 頁面＋頁面自己導向。
 *   本頁不查資料庫、不分辨單據存不存在（存在與否、有沒有權限，都留給登入後 GET /api/quotations/:id 判斷），所以不洩漏任何單據資訊。
 *
 * 與專案現有 CSP 的相容性（server.js:87-104，helmet useDefaults:false：default-src 'self'；script-src 'self' 'unsafe-inline' cdn.jsdelivr.net；
 *   style-src 'self' 'unsafe-inline'；img-src 'self' data: blob:；frame-ancestors 'none'）：
 *   - 跳板頁完全不使用 JavaScript：用 `<meta http-equiv="refresh" content="0;url=...">` ＋可見的備援連結。CSP 只管資源載入與腳本，
 *     不限制 meta refresh（navigate-to 指令沒有被瀏覽器實作），所以不管 CSP 日後是否移除 'unsafe-inline' 都不受影響，也不需要 nonce。
 *   - 唯一的行內資源是一小段 <style>（現行 style-src 允許 'unsafe-inline'）；就算日後被擋，頁面只是沒有樣式，仍可使用。
 *   - 回應自帶更嚴格的 CSP（default-src 'none'；img-src 'self'；style-src 'unsafe-inline'；base-uri/form-action 'none'；frame-ancestors 'none'）。
 *     路由用 res.set(headers) 設定時會取代 helmet 的同名標頭（不是疊加），這份只會比現行更嚴，不會比較鬆。
 *   - Service Worker（_client/sw.js）對導覽請求是 network-first，且會 cache.put 成功的 HTML（Cache API 不理會 no-store）。
 *     快取內容只是這頁導向，無機敏資料，無害；404 不會被快取。若要完全避開，需在 sw.js 的 isNetworkOnly 加 '/q/'（那是別人的檔案，這裡不改）。
 *
 * 路由建議（供整合者參考，本檔不做）：必須放在 requireAuth 與靜態檔之前，只接 GET：
 *   app.get('/q/:id', (req, res) => {
 *     const r = renderJumpPage({ quoteId: req.params.id, cost: req.query.cost });
 *     res.status(r.status).set(r.headers).send(r.body);
 *   });
 *   req.query.cost 若被重複帶（?cost=1&cost=1）會是陣列，isCostFlag 視為 false。
 *
 * 限制：未實測「從 Outlook／Teams 實際點擊後 cookie 是否帶上」；未實測登入頁是否保留 #quote 片段（見 _client/deep-link.js）。
 */

const { getMailConfig } = require('./config');
const { isValidQuoteId } = require('./events');
const { escHtml } = require('./safety');

const JUMP_PATH_PREFIX = '/q/';

class MailLinkError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'MailLinkError';
    this.code = code;
  }
}

/** 只有 true、1、'1' 算「填成本」；'true'、'yes'、陣列、物件一律不算。 */
function isCostFlag(v) {
  return v === true || v === 1 || v === '1';
}

function assertQuoteId(quoteId) {
  if (!isValidQuoteId(quoteId)) throw new MailLinkError('BAD_QUOTE_ID', '單據 id 格式不合法（僅允許英數、底線、連字號，1–64 字元）');
}

function deepHash(quoteId, opts) {
  assertQuoteId(quoteId);
  return '#quote:' + quoteId + (opts && isCostFlag(opts.cost) ? ':cost' : '');
}

/**
 * 重新驗證網站根網址：交給 config.js 的同一套規則（https；localhost／127.0.0.1 可 http；不可帶帳密／路徑／查詢；主機名稱嚴格字元）。
 * 手工拼出來的 config（測試、其他呼叫端）可能夾帶 javascript:、http://evil.example 之類，這裡 fail closed。
 */
function checkedBase(config) {
  let raw;
  try { raw = (config && typeof config === 'object') ? config.appBaseUrl : undefined; } catch (e) { raw = undefined; }
  if (typeof raw !== 'string' || raw.trim() === '') throw new MailLinkError('BAD_BASE_URL', '網站根網址未設定');
  const probe = getMailConfig({ APP_BASE_URL: raw });
  if (probe.warnings.length) throw new MailLinkError('BAD_BASE_URL', '網站根網址不合法');
  return probe.appBaseUrl;
}

function buildQuoteLink(config, quoteId, opts) {
  assertQuoteId(quoteId);
  const base = checkedBase(config);
  return base + JUMP_PATH_PREFIX + encodeURIComponent(quoteId) + (opts && isCostFlag(opts.cost) ? '?cost=1' : '');
}

// ── 跳板頁 ──────────────────────────────────────────────────────────────────
const PAGE_STYLE =
  ':root{color-scheme:light dark}' +
  'body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;' +
  'font:16px/1.7 -apple-system,"Microsoft JhengHei","PingFang TC",sans-serif;background:#f6f7f9;color:#1f2933}' +
  'main{max-width:420px;padding:24px;text-align:center}' +
  'a{color:#1a3c7a}' +
  '@media (prefers-color-scheme:dark){body{background:#14181f;color:#e4e7eb}a{color:#8fb4ff}}';

const JUMP_CSP = "default-src 'none'; img-src 'self'; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'";

/** 每次回傳新物件，呼叫端可以放心修改或交給 res.set()。 */
function pageHeaders() {
  return {
    'Content-Type': 'text/html; charset=utf-8',
    'Cache-Control': 'no-store',
    'X-Robots-Tag': 'noindex',
    'Referrer-Policy': 'no-referrer',
    'X-Content-Type-Options': 'nosniff',
    'Content-Security-Policy': JUMP_CSP,
  };
}

function pageShell(title, head, bodyInner) {
  return '<!doctype html>\n<html lang="zh-Hant">\n<head>\n<meta charset="utf-8">\n' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">\n' +
    '<meta name="robots" content="noindex, nofollow">\n' +
    head +
    '<title>' + escHtml(title) + '</title>\n<style>' + PAGE_STYLE + '</style>\n</head>\n<body>\n<main>\n' +
    bodyInner +
    '</main>\n</body>\n</html>\n';
}

// 所有無效輸入共用這一份固定內容：不回顯輸入、不因輸入而不同
const NOT_FOUND_BODY = pageShell(
  '連結無法使用',
  '',
  '<p>這個連結無法使用。</p>\n<p>請回到信件重新點選連結，或直接<a href="/">登入 ITTS-CRM</a>。</p>\n'
);

function notFound() {
  return { status: 404, headers: pageHeaders(), body: NOT_FOUND_BODY };
}

function renderJumpPage(input) {
  try {
    const quoteId = (input && typeof input === 'object') ? input.quoteId : undefined;
    if (!isValidQuoteId(quoteId)) return notFound();
    const target = '/index.html' + deepHash(quoteId, { cost: input.cost });
    const t = escHtml(target);
    const body = pageShell(
      'ITTS-CRM',
      '<meta http-equiv="refresh" content="0;url=' + t + '">\n',
      '<p>正在前往 ITTS-CRM 報價單…</p>\n<p><a href="' + t + '">若沒有自動跳轉，請按這裡</a></p>\n'
    );
    return { status: 200, headers: pageHeaders(), body };
  } catch (e) {
    return notFound();      // 任何意外（例如惡意物件的 getter 丟例外）都等同無效連結
  }
}

module.exports = {
  buildQuoteLink,
  deepHash,
  renderJumpPage,
  isValidQuoteId,
  isCostFlag,
  MailLinkError,
  JUMP_PATH_PREFIX,
};
