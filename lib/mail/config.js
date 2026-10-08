'use strict';
/**
 * lib/mail/config.js — 信件派送設定（全部來自環境變數；不讀檔、不連網、不修改 process.env）
 *
 * 匯出與簽名：
 *   getMailConfig(env = process.env) → Config     每次呼叫回傳「新物件」（陣列／巢狀物件不共用，呼叫端改了不會污染下一次）
 *   publicConfig(config)             → object     後台顯示用；不含 redirectTo 原值、不含任何密鑰／租戶資訊
 *   parseMode(raw)                   → 'off'|'log'|'redirect'|'live'|null   無法辨識回 null
 *   normalizeMode(raw)               → 同上，但無法辨識一律回 'off'（transport／dispatcher 做第二層護欄時用）
 *   MODES, DEFAULTS
 *
 * 環境變數（值絕不可寫進 repo；公開 repo）：
 *   MAIL_MODE            off|log|redirect|live。未設或非法 ＝ off（安全預設，絕不因缺設定而寄出）。
 *                        前後空白與大小寫會先正規化（' LIVE ' 等同 live）；全形、同形字、'true'/'1'/'yes' 一律不是 live。
 *   MAIL_REDIRECT_TO     redirect 模式的單一測試收件信箱（需通過 normalizeEmail；不受網域白名單限制）。
 *                        mode=redirect 但缺／非法 → 有效模式降為 log 並記一筆 warning（不會退回寄給原收件人）。
 *   MAIL_FROM_NAME       寄件顯示名稱，預設「ITTS-CRM 簽核通知」（去換行與 <>"\、≤60 字）
 *   APP_BASE_URL         信內連結的網站根網址。次選 'https://'+VERCEL_PROJECT_PRODUCTION_URL，最後預設正式站。
 *                        必須是 https（localhost／127.0.0.1 可用 http）、不可帶帳密／路徑／查詢；不合法 → 預設值＋warning。
 *   MAIL_ALLOWED_DOMAINS 收件人網域白名單（逗號分隔），預設 itts.com.tw。精確相等比對。
 *   MAIL_GRAPH_TENANT_ID / MAIL_GRAPH_CLIENT_ID / MAIL_GRAPH_CLIENT_SECRET / MAIL_GRAPH_SENDER
 *                        P4（Graph sendMail）才使用；本階段只讀取。四者皆合法才 configured=true。
 *   MAIL_PREVIEW_DIR     log／redirect 模式把信件寫成檔案的目錄，預設 '.mail-preview'。
 *                        Vercel（env.VERCEL 有值）一律 null：唯讀檔案系統、也不把信件內容印到 console。
 *   MAIL_OUTBOX_FILE     本機 JSON outbox 檔名，預設 'mail-outbox.json'。
 *
 * 密鑰保護（graph.clientSecret）：
 *   - graph.clientSecret（連同 tenantId／clientId／sender）以 non-enumerable 屬性存放：直接讀取得到，
 *     但 JSON.stringify、Object.keys、展開運算子（...）、Object.assign、console.log 預設都帶不走。只有 configured 可列舉。
 *   - graph 物件自帶 toJSON（只輸出 {configured}）與 util.inspect.custom（只輸出一行摘要，連 showHidden 也一樣）。
 *   - warnings 與 publicConfig 只講「哪個變數有問題」，不回顯任何變數的值。
 *   P4 的傳輸層直接讀 config.graph.clientSecret。
 *
 * 限制：config.mode 是「有效模式」（已套用降級規則）；原始 MAIL_MODE 設定值不另外保留。
 */

const util = require('util');
const { normalizeEmail, normalizeDomain, headerSafe } = require('./safety');

const MODES = Object.freeze(['off', 'log', 'redirect', 'live']);

const DEFAULTS = Object.freeze({
  fromName: 'ITTS-CRM 簽核通知',
  appBaseUrl: 'https://itts-crm.vercel.app',
  allowedDomains: Object.freeze(['itts.com.tw']),
  previewDir: '.mail-preview',
  outboxFile: 'mail-outbox.json',
  retentionDays: 90,
});

// ── 小工具 ──────────────────────────────────────────────────────────────────
function str(env, key) {
  const v = env[key];
  return typeof v === 'string' ? v.trim() : '';
}

/** 只裁切 ASCII 空白（不裁 NBSP 等 Unicode 空白——環境變數裡出現那種字元本身就可疑）。 */
function asciiTrim(s) {
  return s.replace(/^[ \t\r\n]+/, '').replace(/[ \t\r\n]+$/, '');
}

function parseMode(raw) {
  if (typeof raw !== 'string' || raw.length > 64) return null;
  const s = asciiTrim(raw);
  if (!/^[A-Za-z]{2,10}$/.test(s)) return null;          // 先要求純 ASCII 字母，再轉小寫（避免 Kelvin 符號之類洗成 ASCII）
  const m = s.toLowerCase();
  return MODES.indexOf(m) >= 0 ? m : null;
}

function normalizeMode(raw) {
  return parseMode(raw) || 'off';
}

/** 路徑類設定（預覽目錄、outbox 檔名）：長度、控制字元、`..` 片段檢查。 */
function isSafePathSetting(p) {
  if (typeof p !== 'string' || !p || p.length > 200) return false;
  if (/[\x00-\x1f\x7f]/.test(p)) return false;
  if (p.split(/[\\/]/).some((seg) => seg === '..')) return false;
  return true;
}

/** 回傳正規化後的網站根（origin），不合法回 ''。 */
function validateBaseUrl(candidate) {
  if (typeof candidate !== 'string' || candidate.length > 200) return '';
  const trimmed = candidate.replace(/\/+$/, '');
  let u;
  try { u = new URL(trimmed); } catch (e) { return ''; }
  if (u.username || u.password || u.search || u.hash) return '';
  if (u.pathname !== '/' && u.pathname !== '') return '';
  // URL 解析器不會拒絕主機名稱裡的引號、等號等字元（例如 https://a.test"onmouseover=），所以主機名稱另外嚴格檢查
  if (!/^[a-z0-9]([a-z0-9.-]*[a-z0-9])?$/.test(u.hostname)) return '';
  const isLocal = u.hostname === 'localhost' || u.hostname === '127.0.0.1';
  if (u.protocol === 'https:') return u.origin;
  if (u.protocol === 'http:' && isLocal) return u.origin;
  return '';
}

function resolveBaseUrl(env, warnings) {
  const explicit = str(env, 'APP_BASE_URL');
  const vercelHost = str(env, 'VERCEL_PROJECT_PRODUCTION_URL');
  let candidate = '';
  let source = '';
  if (explicit) { candidate = explicit; source = 'APP_BASE_URL'; }
  else if (vercelHost) { candidate = 'https://' + vercelHost; source = 'VERCEL_PROJECT_PRODUCTION_URL'; }
  else return DEFAULTS.appBaseUrl;
  const ok = validateBaseUrl(candidate);
  if (ok) return ok;
  warnings.push(source + ' 不是合法的網站根網址（需 https；localhost 可用 http；不可含帳密、路徑或查詢），已改用預設正式站網址');
  return DEFAULTS.appBaseUrl;
}

function resolveAllowedDomains(env, warnings) {
  const raw = str(env, 'MAIL_ALLOWED_DOMAINS');
  if (!raw) return DEFAULTS.allowedDomains.slice();
  const out = [];
  let bad = 0;
  const parts = raw.slice(0, 4000).split(',');
  for (let i = 0; i < parts.length; i++) {
    if (!parts[i].trim()) continue;
    const d = normalizeDomain(parts[i]);
    if (!d) { bad++; continue; }
    if (out.indexOf(d) < 0) out.push(d);
  }
  if (bad) warnings.push('MAIL_ALLOWED_DOMAINS 有 ' + bad + ' 個項目格式不合法，已忽略');
  if (!out.length) {
    warnings.push('MAIL_ALLOWED_DOMAINS 沒有任何合法網域，已使用預設白名單');
    return DEFAULTS.allowedDomains.slice();
  }
  return out;
}

const RE_GRAPH_ID = /^[A-Za-z0-9.-]{1,128}$/;       // tenant／client id 會進 URL 路徑與表單，限制字元避免路徑／參數注入

function buildGraph(env, warnings) {
  let tenantId = str(env, 'MAIL_GRAPH_TENANT_ID');
  let clientId = str(env, 'MAIL_GRAPH_CLIENT_ID');
  let clientSecret = str(env, 'MAIL_GRAPH_CLIENT_SECRET');
  let sender = str(env, 'MAIL_GRAPH_SENDER');
  const setCount = [tenantId, clientId, clientSecret, sender].filter(Boolean).length;

  if (tenantId && !RE_GRAPH_ID.test(tenantId)) { warnings.push('MAIL_GRAPH_TENANT_ID 格式不合法，已忽略'); tenantId = ''; }
  if (clientId && !RE_GRAPH_ID.test(clientId)) { warnings.push('MAIL_GRAPH_CLIENT_ID 格式不合法，已忽略'); clientId = ''; }
  if (clientSecret && (clientSecret.length > 512 || /[\x00-\x1f\x7f]/.test(clientSecret))) {
    warnings.push('MAIL_GRAPH_CLIENT_SECRET 含控制字元或過長，已忽略'); clientSecret = '';
  }
  if (sender) {
    const r = normalizeEmail(sender);
    if (r.ok) sender = r.value; else { warnings.push('MAIL_GRAPH_SENDER 格式不合法，已忽略'); sender = ''; }
  }
  const configured = !!(tenantId && clientId && clientSecret && sender);
  if (setCount > 0 && !configured) warnings.push('Graph 寄信設定不完整（MAIL_GRAPH_TENANT_ID / CLIENT_ID / CLIENT_SECRET / SENDER 需四項齊全且合法）');

  // 四個識別／密鑰欄位都是 non-enumerable（可直接讀取，但展開、Object.assign、Object.keys、JSON 都帶不走）；只有 configured 可列舉
  const g = {};
  Object.defineProperty(g, 'tenantId', { value: tenantId, enumerable: false, writable: false, configurable: false });
  Object.defineProperty(g, 'clientId', { value: clientId, enumerable: false, writable: false, configurable: false });
  Object.defineProperty(g, 'sender', { value: sender, enumerable: false, writable: false, configurable: false });
  Object.defineProperty(g, 'clientSecret', { value: clientSecret, enumerable: false, writable: false, configurable: false });
  Object.defineProperty(g, 'configured', { value: configured, enumerable: true, writable: false, configurable: false });
  // 序列化與 inspect 一律不含任何值（租戶／用戶端 ID 與寄件者也一併遮蔽）
  Object.defineProperty(g, 'toJSON', { value: function () { return { configured: configured }; }, enumerable: false });
  Object.defineProperty(g, util.inspect.custom, {
    value: function () { return '[GraphConfig configured=' + configured + ']'; },
    enumerable: false,
  });
  return g;
}

/**
 * 組出 Config。env 預設 process.env；傳 null／非物件視為空環境。
 * @returns {{mode:string, redirectTo:string, fromName:string, appBaseUrl:string, allowedDomains:string[],
 *   graph:{tenantId:string,clientId:string,sender:string,configured:boolean},
 *   timeouts:{connectMs:number,totalMs:number}, retry:{delaysSec:number[],maxAttempts:number},
 *   breaker:{failures:number,windowSec:number,openSec:number}, previewDir:(string|null), outboxFile:string,
 *   retentionDays:number, warnings:string[]}}
 */
function getMailConfig(env = process.env) {
  const e = (env && typeof env === 'object') ? env : {};
  const warnings = [];

  // 1) 模式
  const rawMode = typeof e.MAIL_MODE === 'string' ? e.MAIL_MODE : (e.MAIL_MODE == null ? '' : String(e.MAIL_MODE));
  let mode = parseMode(typeof e.MAIL_MODE === 'string' ? e.MAIL_MODE : null);
  if (mode === null) {
    if (asciiTrim(rawMode) !== '') warnings.push('MAIL_MODE 設定值不合法（僅接受 off / log / redirect / live），已視為 off');
    mode = 'off';
  }

  // 2) redirect 目標
  let redirectTo = '';
  const rawRedirect = str(e, 'MAIL_REDIRECT_TO');
  if (rawRedirect) {
    const r = normalizeEmail(rawRedirect);
    if (r.ok) redirectTo = r.value;
    else warnings.push('MAIL_REDIRECT_TO 格式不合法（需單一個合法 Email），已忽略');
  }
  if (mode === 'redirect' && !redirectTo) {
    mode = 'log';
    warnings.push('MAIL_MODE=redirect 但沒有有效的 MAIL_REDIRECT_TO，已降為 log（不會寄出任何信）');
  }

  // 3) 寄件顯示名稱
  let fromName = headerSafe(str(e, 'MAIL_FROM_NAME'), 60).replace(/[<>"\\]/g, '').trim();
  if (!fromName) fromName = DEFAULTS.fromName;

  // 4) 其他
  const appBaseUrl = resolveBaseUrl(e, warnings);
  const allowedDomains = resolveAllowedDomains(e, warnings);
  const graph = buildGraph(e, warnings);
  if (mode === 'live' && !graph.configured) {
    warnings.push('MAIL_MODE=live 但 Graph 寄信設定不完整，所有寄送都會以 NOT_CONFIGURED 失敗');
  }

  // 5) 預覽目錄與 outbox 檔
  const onVercel = str(e, 'VERCEL') !== '';
  let previewDir = null;
  if (!onVercel) {
    const rawDir = str(e, 'MAIL_PREVIEW_DIR');
    if (!rawDir) previewDir = DEFAULTS.previewDir;
    else if (isSafePathSetting(rawDir)) previewDir = rawDir;
    else { previewDir = DEFAULTS.previewDir; warnings.push('MAIL_PREVIEW_DIR 不合法（不可含 .. 或控制字元），已改用預設目錄'); }
  }
  let outboxFile = DEFAULTS.outboxFile;
  const rawOutbox = str(e, 'MAIL_OUTBOX_FILE');
  if (rawOutbox) {
    if (isSafePathSetting(rawOutbox)) outboxFile = rawOutbox;
    else warnings.push('MAIL_OUTBOX_FILE 不合法（不可含 .. 或控制字元），已改用預設檔名');
  }

  return {
    mode,
    redirectTo,
    fromName,
    appBaseUrl,
    allowedDomains,
    graph,
    timeouts: { connectMs: 3000, totalMs: 8000 },
    retry: { delaysSec: [60, 300, 900], maxAttempts: 3 },
    breaker: { failures: 5, windowSec: 600, openSec: 600 },
    previewDir,
    outboxFile,
    retentionDays: DEFAULTS.retentionDays,
    warnings,
  };
}

/** 後台顯示用的公開版設定：只含非機敏欄位；對殘缺的 config 也不 throw。 */
function publicConfig(config) {
  const c = (config && typeof config === 'object') ? config : {};
  const arr = (a) => (Array.isArray(a) ? a.slice() : []);
  return {
    mode: normalizeMode(c.mode),
    redirectConfigured: !!c.redirectTo,
    fromName: typeof c.fromName === 'string' ? c.fromName : DEFAULTS.fromName,
    appBaseUrl: typeof c.appBaseUrl === 'string' ? c.appBaseUrl : DEFAULTS.appBaseUrl,
    allowedDomains: arr(c.allowedDomains),
    graphConfigured: !!(c.graph && c.graph.configured),
    timeouts: Object.assign({}, c.timeouts),
    retry: { delaysSec: arr(c.retry && c.retry.delaysSec), maxAttempts: c.retry && c.retry.maxAttempts },
    warnings: arr(c.warnings),
  };
}

module.exports = {
  getMailConfig,
  publicConfig,
  parseMode,
  normalizeMode,
  MODES,
  DEFAULTS,
};
