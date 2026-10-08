'use strict';
/**
 * lib/mail/transports.js — 信件傳輸層（真正「送出」的那一層）。本階段只有假傳輸；Graph 為 P4。
 *
 * 匯出與簽名：
 *   createNullTransport()                        永遠不寄（off 模式或模式不明時使用）
 *   createLogTransport({config, rootDir?, fs?, now?})       不寄；previewDir 有值時把信件寫成檔案供人工檢視
 *   createRedirectTransport({config, inner})     把所有收件人改寫成 config.redirectTo 再交給 inner
 *   createGraphTransport({config, fetchImpl?})   P4 才實作；本階段永遠回 NOT_CONFIGURED（API 形狀見下）
 *   selectTransport(config, deps?)               依「有效模式」挑傳輸，是整個管線唯一的模式護欄
 *   normalizeResult(r) / failure(code, msg, extra?) / isPermanentCode(code) / CODES / PERMANENT_CODES / validateMessage(msg)
 *
 * 介面：transport.send(msg, {timeouts, signal}) → Promise<Result>；傳輸物件另有 .kind（'null'|'log'|'redirect'|'graph'|…）。
 *   msg    = {to:[email], subject, html, text, tag}（只認這五個欄位；cc／bcc／headers 之類多餘欄位一律不會被傳下去）
 *   Result = {ok:true, providerId?} | {ok:false, code, permanent, message, retryAfterSec?}
 *   code   = TIMEOUT|NETWORK|AUTH|THROTTLED|REJECTED|SERVER|NOT_CONFIGURED|BAD_MESSAGE
 *            可重試：TIMEOUT／NETWORK／THROTTLED／SERVER；視為 permanent：AUTH／REJECTED／BAD_MESSAGE／NOT_CONFIGURED
 *            （AUTH 雖然 permanent，仍會觸發熔斷，見 outbox.js 的 breaker）
 *   所有 transport 的 send 都「永不 throw、永不 reject」，任何例外都包成 Result。逾時由呼叫端（dispatcher）用 AbortController＋Promise.race 控制。
 *
 * 模式護欄（selectTransport；對應 config.mode 的「有效模式」，並再做一次 normalizeMode 當第二層）：
 *   off      → null 傳輸（dispatcher 在 off 時根本不會呼叫它）
 *   log      → log 傳輸。不管有沒有傳入 realTransport，都不會碰它
 *   redirect → redirect(inner = deps.realTransport，沒有就用 log)。收件人一定被改寫成 config.redirectTo；redirectTo 缺失→NOT_CONFIGURED，
 *              絕不退回寄給原收件人
 *   live     → deps.realTransport，沒有就用 graph 傳輸（本階段是 NOT_CONFIGURED）
 *   其他任何值（缺設定、未知字串、非字串）→ 一律當 off。只有 normalizeMode 後恰好是 'live' 才可能走到 live 分支。
 *   注意：deps.realTransport 傳入的是「真的會寄信的那個傳輸」，不是 selectTransport 的輸出。
 *
 * P4（Graph sendMail）的 API 形狀，供之後實作：
 *   1) 取權杖  POST https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token
 *              form：client_id、client_secret、scope=https://graph.microsoft.com/.default、grant_type=client_credentials
 *              權杖只放記憶體，到期前 5 分鐘更新；不寫 log、不進 outbox（錯誤訊息一律先過 scrubMessage）
 *   2) 寄信    POST https://graph.microsoft.com/v1.0/users/{sender}/sendMail
 *              Authorization: Bearer <token>；JSON：{message:{subject, body:{contentType:'HTML', content}, toRecipients:[{emailAddress:{address}}]}, saveToSentItems:true}
 *              成功＝202 Accepted（只代表受理，不代表送達）。純文字備援需改用 MIME 形式的 sendMail，P4 再決定。
 *   3) 只接受 https、用 Node 內建 fetch、不提供關閉憑證驗證的任何選項；逾時用 AbortSignal（signal 由 dispatcher 傳入）。
 *   4) 錯誤分類：連線失敗→NETWORK、逾時→TIMEOUT、401／403→AUTH、429→THROTTLED（讀 Retry-After 秒數放進 retryAfterSec）、
 *      500／502／503／504→SERVER、其他 4xx（400／404 等）→REJECTED、內容不合法→BAD_MESSAGE。
 *
 * 限制：log 傳輸寫出的 .json 含「真實收件位址」與主旨（規格如此，方便本機對照），所以預覽目錄必須在 .gitignore 內（.mail-preview/ 已加）。
 */

const fs = require('fs');
const path = require('path');
const { normalizeEmail, maskEmail, headerSafe, escHtml } = require('./safety');
const { normalizeMode } = require('./config');
const { scrubMessage } = require('./scrub');

const CODES = Object.freeze(['TIMEOUT', 'NETWORK', 'AUTH', 'THROTTLED', 'REJECTED', 'SERVER', 'NOT_CONFIGURED', 'BAD_MESSAGE']);
const PERMANENT_CODES = Object.freeze(['AUTH', 'REJECTED', 'BAD_MESSAGE', 'NOT_CONFIGURED']);
const MAX_RECIPIENTS = 50;
const MAX_SUBJECT = 255;
const MAX_BODY = 1024 * 1024;
const MAX_RETRY_AFTER_SEC = 24 * 3600;
const REDIRECT_SUBJECT_PREFIX = '[測試轉送] ';

const RE_HEADER_BAD = /[\x00-\x1f\x7f-\x9f\u{2028}\u{2029}]/u;

function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }
function isPermanentCode(code) { return PERMANENT_CODES.indexOf(code) >= 0; }

function failure(code, message, extra) {
  const c = CODES.indexOf(code) >= 0 ? code : 'SERVER';
  const r = { ok: false, code: c, permanent: isPermanentCode(c) || !!(extra && extra.permanent === true), message: scrubMessage(message) };
  if (extra && typeof extra.retryAfterSec === 'number' && isFinite(extra.retryAfterSec) && extra.retryAfterSec > 0) {
    r.retryAfterSec = Math.min(extra.retryAfterSec, MAX_RETRY_AFTER_SEC);
  }
  return r;
}

/** 把任何 transport 回傳值整理成標準 Result；格式不對一律當 SERVER（可重試，次數受 outbox 限制）。 */
function normalizeResult(r) {
  if (isObj(r) && r.ok === true) {
    const out = { ok: true };
    if (typeof r.providerId === 'string' && r.providerId) out.providerId = r.providerId.slice(0, 100);
    if (typeof r.warning === 'string' && r.warning) out.warning = r.warning.slice(0, 60);
    return out;
  }
  if (isObj(r) && r.ok === false) {
    return failure(typeof r.code === 'string' ? r.code : 'SERVER', typeof r.message === 'string' ? r.message : '', { retryAfterSec: r.retryAfterSec, permanent: r.permanent === true });
  }
  return failure('SERVER', '傳輸層回傳格式不合法');
}

function errCode(e) { return e && typeof e.code === 'string' ? e.code : (e && e.name) || 'ERR'; }

/** 檢查並複製出標準訊息（只留 to／subject／html／text／tag）。 */
function validateMessage(msg) {
  if (!isObj(msg)) return { ok: false, error: '訊息必須是物件' };
  if (!Array.isArray(msg.to) || msg.to.length < 1) return { ok: false, error: '收件人不可為空' };
  if (msg.to.length > MAX_RECIPIENTS) return { ok: false, error: '收件人過多' };
  const to = [];
  for (let i = 0; i < msg.to.length; i++) {
    const n = normalizeEmail(msg.to[i]);
    if (!n.ok) return { ok: false, error: '收件人位址不合法' };
    if (to.indexOf(n.value) < 0) to.push(n.value);
  }
  if (typeof msg.subject !== 'string' || msg.subject.trim() === '') return { ok: false, error: '主旨不可為空' };
  if (msg.subject.length > MAX_SUBJECT) return { ok: false, error: '主旨過長' };
  if (RE_HEADER_BAD.test(msg.subject)) return { ok: false, error: '主旨含換行或控制字元' };
  const html = msg.html === undefined || msg.html === null ? '' : msg.html;
  const text = msg.text === undefined || msg.text === null ? '' : msg.text;
  if (typeof html !== 'string' || typeof text !== 'string') return { ok: false, error: '內文必須是文字' };
  if (html.length > MAX_BODY || text.length > MAX_BODY) return { ok: false, error: '內文過大' };
  if (html === '' && text === '') return { ok: false, error: '內文不可為空' };
  const tag = typeof msg.tag === 'string' ? msg.tag.slice(0, 80) : '';
  return { ok: true, msg: { to, subject: msg.subject, html, text, tag } };
}

// ═════════════════════════════════════════════════════════════════════════
function createNullTransport() {
  return {
    kind: 'null',
    async send() {
      return failure('NOT_CONFIGURED', '寄信模式為 off，不會寄出任何信件');
    },
  };
}

// ═════════════════════════════════════════════════════════════════════════
function pad(n, w) { return String(n).padStart(w, '0'); }
function stampOf(ms) {
  const d = new Date(ms);
  if (isNaN(d.getTime())) return '00000000T000000000';
  return d.getUTCFullYear() + pad(d.getUTCMonth() + 1, 2) + pad(d.getUTCDate(), 2) + 'T' + pad(d.getUTCHours(), 2) + pad(d.getUTCMinutes(), 2) + pad(d.getUTCSeconds(), 2) + pad(d.getUTCMilliseconds(), 3);
}
function safeFilePart(s, max) {
  const t = String(s || '').replace(/[^A-Za-z0-9_-]/g, '_').slice(0, max);
  return t === '' ? 'mail' : t;
}

/**
 * log 傳輸：不寄。config.previewDir 有值（本機）→ 寫 <ts>-<tag>-<n>.json／.html／.txt；沒有值（Vercel）→ 只回 ok，不輸出任何信件內容。
 * 檔名只含 [A-Za-z0-9_-] 與固定副檔名，且寫入前再驗證解析後路徑仍在預覽目錄內（不可路徑穿越）。
 * 預覽檔寫入失敗不算寄信失敗（log 模式本來就沒有寄信）：回 ok 並帶 warning:'PREVIEW_WRITE_FAILED'。
 */
function createLogTransport(opts) {
  const o = isObj(opts) ? opts : {};
  const config = isObj(o.config) ? o.config : {};
  const fsx = o.fs || fs;
  const rootDir = typeof o.rootDir === 'string' && o.rootDir ? o.rootDir : path.join(__dirname, '..', '..');
  const nowFn = typeof o.now === 'function' ? o.now : Date.now;
  const setting = typeof config.previewDir === 'string' && config.previewDir ? config.previewDir : null;
  const dir = setting ? (path.isAbsolute(setting) ? setting : path.join(rootDir, setting)) : null;
  let counter = 0;

  function writePreview(v, providerId) {
    const resolvedDir = path.resolve(dir);
    fsx.mkdirSync(resolvedDir, { recursive: true });
    const tag = safeFilePart(v.tag, 40);
    const ts = stampOf(Number(nowFn()));
    for (let attempt = 0; attempt < 20; attempt++) {
      const base = ts + '-' + tag + '-' + (counter + attempt);
      const jsonPath = path.resolve(resolvedDir, base + '.json');
      if (path.dirname(jsonPath) !== resolvedDir) throw new Error('PATH_ESCAPE');
      try {
        fsx.writeFileSync(jsonPath, JSON.stringify({ providerId, to: v.to, subject: v.subject, tag: v.tag }, null, 2), { encoding: 'utf8', flag: 'wx' });
      } catch (e) {
        if (e && e.code === 'EEXIST') continue;
        throw e;
      }
      if (v.html) fsx.writeFileSync(path.resolve(resolvedDir, base + '.html'), v.html, { encoding: 'utf8', flag: 'wx' });
      if (v.text) fsx.writeFileSync(path.resolve(resolvedDir, base + '.txt'), v.text, { encoding: 'utf8', flag: 'wx' });
      return;
    }
    throw new Error('NAME_EXHAUSTED');
  }

  return {
    kind: 'log',
    async send(msg) {
      try {
        const v = validateMessage(msg);
        if (!v.ok) return failure('BAD_MESSAGE', v.error);
        counter += 1;
        const providerId = 'log:' + counter;
        if (!dir) return { ok: true, providerId };
        try { writePreview(v.msg, providerId); } catch (e) { return { ok: true, providerId, warning: 'PREVIEW_WRITE_FAILED' }; }
        return { ok: true, providerId };
      } catch (e) {
        return failure('NETWORK', 'log 傳輸發生未預期的錯誤：' + errCode(e));
      }
    },
  };
}

// ═════════════════════════════════════════════════════════════════════════
function bannerHtml(masked) {
  return `<div style="margin:0 0 12px 0;padding:8px 12px;background:#fff3cd;color:#664d03;border:1px solid #ffecb5;font-family:'Microsoft JhengHei',Arial,sans-serif;font-size:13px;line-height:1.5;">【測試轉送】原收件人：${escHtml(masked)}（redirect 模式，僅供測試）</div>`;
}
function injectAfterBody(html, banner) {
  const lower = html.toLowerCase();
  const i = lower.indexOf('<body');
  if (i >= 0) {
    const j = lower.indexOf('>', i);
    if (j >= 0) return html.slice(0, j + 1) + banner + html.slice(j + 1);
  }
  return banner + html;
}

/**
 * redirect 傳輸：收件人一律改寫成 config.redirectTo（單一位址），主旨加 [測試轉送]，內文最上方加註原收件人（遮罩）。
 * redirectTo 缺失或不合法 → NOT_CONFIGURED，並且「不呼叫 inner」。傳給 inner 的是全新物件，只含 to／subject／html／text／tag。
 */
function createRedirectTransport(opts) {
  const o = isObj(opts) ? opts : {};
  const config = isObj(o.config) ? o.config : {};
  const inner = o.inner;
  return {
    kind: 'redirect',
    async send(msg, sendOpts) {
      try {
        const target = normalizeEmail(config.redirectTo);
        if (!target.ok) return failure('NOT_CONFIGURED', 'redirect 模式需要有效的 MAIL_REDIRECT_TO');
        if (!inner || typeof inner.send !== 'function') return failure('NOT_CONFIGURED', 'redirect 模式缺少內層傳輸');
        const v = validateMessage(msg);
        if (!v.ok) return failure('BAD_MESSAGE', v.error);
        const masked = v.msg.to.map(maskEmail).join('、');
        const out = {
          to: [target.value],
          subject: headerSafe(REDIRECT_SUBJECT_PREFIX + v.msg.subject, MAX_SUBJECT),
          html: v.msg.html ? injectAfterBody(v.msg.html, bannerHtml(masked)) : '',
          text: v.msg.text ? '【測試轉送】原收件人：' + masked + '（redirect 模式，僅供測試）\n\n' + v.msg.text : '',
          tag: v.msg.tag,
        };
        let r;
        try { r = await inner.send(out, sendOpts); } catch (e) { return failure('NETWORK', '內層傳輸拋出例外：' + errCode(e)); }
        return normalizeResult(r);
      } catch (e) {
        return failure('NETWORK', 'redirect 傳輸發生未預期的錯誤：' + errCode(e));
      }
    },
  };
}

// ═════════════════════════════════════════════════════════════════════════
/** Graph 傳輸（P4 才實作）：本階段永遠回 NOT_CONFIGURED。API 形狀見檔頭。 */
function createGraphTransport(opts) {
  return {
    kind: 'graph',
    async send() {
      return failure('NOT_CONFIGURED', 'Graph 傳輸尚未實作（P4）');
    },
  };
}

// ═════════════════════════════════════════════════════════════════════════
/**
 * 模式護欄。回傳的傳輸物件帶 kind。
 * @param {object} config
 * @param {{realTransport?: object, logTransport?: object, fetchImpl?: function, rootDir?: string, fs?: object, now?: function}} deps
 */
function selectTransport(config, deps) {
  const d = isObj(deps) ? deps : {};
  const c = isObj(config) ? config : {};
  const mode = normalizeMode(c.mode);          // 只有恰好等於 'live'（正規化後）才可能走到 live 分支
  const real = d.realTransport && typeof d.realTransport.send === 'function' ? d.realTransport : null;
  const makeLog = () => (d.logTransport && typeof d.logTransport.send === 'function'
    ? d.logTransport
    : createLogTransport({ config: c, rootDir: d.rootDir, fs: d.fs, now: d.now }));
  if (mode === 'log') return makeLog();
  if (mode === 'redirect') return createRedirectTransport({ config: c, inner: real || makeLog() });
  if (mode === 'live') return real || createGraphTransport({ config: c, fetchImpl: d.fetchImpl });
  return createNullTransport();
}

module.exports = {
  createNullTransport,
  createLogTransport,
  createRedirectTransport,
  createGraphTransport,
  selectTransport,
  normalizeResult,
  failure,
  isPermanentCode,
  validateMessage,
  CODES,
  PERMANENT_CODES,
  REDIRECT_SUBJECT_PREFIX,
};
