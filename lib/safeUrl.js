'use strict';
/**
 * 外連 URL 的 SSRF 防護（SEC-08）。從 lib/auditChain.js 的 webhook 檢查抽出來共用：
 *   · 稽核錨點 webhook（AUDIT_ANCHOR_WEBHOOK_URL，lib/auditChain.js pushAnchor）
 *   · 管理員設定的整合網址（SAP Base URL／Token Endpoint；server.js 的 /api/admin/integrations/*）
 *
 * 規則（任何一條不過就拒絕）：
 *   1. 只允許 https；2. 不可帶帳密（https://user:pass@host）；3. 連接埠白名單（預設只有 443）；
 *   4. 主機是 localhost／*.localhost 或 IP 字面值落在私有／保留網段 → 拒絕
 *      （127/8、10/8、172.16/12、192.168/16、169.254/16〔含雲端 metadata 169.254.169.254〕、0.0.0.0/8、CGNAT、多播…；
 *        IPv6 的 ::1、::、fc00::/7、fe80::/10、ff00::/8、IPv4-mapped 等）；
 *   5. 網域「DNS 解析後的每一個位址」都必須是公開位址（避免網域指向內網）；
 *   6. 實際連線時連到「剛驗證過的那個 IP」（SNI／Host 仍用原主機名），避免驗證後 DNS 被換掉（DNS rebinding）；
 *   7. 轉址（redirect）每一跳都重新走 1–5 的完整檢查；跨來源轉址會拿掉 Authorization／Cookie；
 *   8. 整體逾時（預設 10 秒）、回應大小上限（預設 5MB）。
 *
 * 全部純 Node 內建模組，沒有新依賴。測試：scripts/check-safe-url.js（不需網路；DNS 與連線都用注入）。
 */
const dns = require('dns');
const https = require('https');
const net = require('net');
const zlib = require('zlib');

class SafeUrlError extends Error {
  constructor(code, message) { super(message || REASON_ZH[code] || code); this.name = 'SafeUrlError'; this.code = code; }
}
// 給使用者看的中文訊息（不含任何網址內容，避免把帶 token 的網址回顯出去）
const REASON_ZH = {
  BAD_URL: '網址格式不正確',
  SCHEME_NOT_HTTPS: '只允許 https 網址（不接受 http）',
  CREDENTIALS_IN_URL: '網址不可包含帳號密碼（不可把帳號與密碼寫在網址裡）',
  PORT_NOT_443: '連接埠只允許 443',
  PORT_NOT_ALLOWED: '不允許的連接埠',
  NO_HOST: '網址缺少主機名稱',
  BLOCKED_ADDRESS: '不允許連到本機或內部網路位址（含 localhost、私有網段、雲端 metadata 位址）',
  DNS_FAILED: '無法解析網域名稱',
  TOO_MANY_REDIRECTS: '轉址次數過多',
  RESPONSE_TOO_LARGE: '回應內容過大',
  REQUEST_FAILED: '連線失敗',
};

// ───────── IP 位址判斷 ─────────
function parseV4(s) {
  const m = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/.exec(s);
  if (!m) return null;
  const o = m.slice(1).map(Number);
  return o.every((n) => n >= 0 && n <= 255) ? o : null;
}
function blockedV4(o) {
  const [a, b, c] = o;
  if (a === 0 || a === 10 || a === 127) return true;
  if (a === 100 && b >= 64 && b <= 127) return true;               // CGNAT
  if (a === 169 && b === 254) return true;                          // link-local（含雲端 metadata 169.254.169.254）
  if (a === 172 && b >= 16 && b <= 31) return true;
  if (a === 192 && b === 168) return true;
  if (a === 192 && b === 0 && c === 0) return true;
  if (a === 192 && b === 0 && c === 2) return true;
  if (a === 198 && (b === 18 || b === 19)) return true;
  if (a === 198 && b === 51 && c === 100) return true;
  if (a === 203 && b === 0 && c === 113) return true;
  if (a >= 224) return true;                                        // 多播／保留／廣播
  return false;
}
function expandV6(s) {
  s = s.toLowerCase().split('%')[0];
  let tail = null;
  const lastColon = s.lastIndexOf(':');
  const v4 = parseV4(s.slice(lastColon + 1));
  if (v4) { tail = [(v4[0] << 8) | v4[1], (v4[2] << 8) | v4[3]]; s = s.slice(0, lastColon + 1) + '0:0'; }
  const parts = s.split('::');
  if (parts.length > 2) return null;
  const head = parts[0] ? parts[0].split(':') : [];
  const rest = parts.length === 2 && parts[1] ? parts[1].split(':') : [];
  const fill = parts.length === 2 ? 8 - head.length - rest.length : 0;
  if (fill < 0 || (parts.length === 1 && head.length !== 8)) return null;
  const g = head.concat(new Array(fill).fill('0'), rest).map((x) => parseInt(x || '0', 16));
  if (g.length !== 8 || g.some((n) => !Number.isFinite(n) || n < 0 || n > 0xffff)) return null;
  if (tail) { g[6] = tail[0]; g[7] = tail[1]; }
  return g;
}
/** 這個 IP（字串）是不是私有／保留／本機位址？不是合法 IP 一律視為「要擋」。 */
function isBlockedIp(ip) {
  const k = net.isIP(ip);
  if (k === 4) { const o = parseV4(ip); return !o || blockedV4(o); }
  if (k === 6) {
    const g = expandV6(ip);
    if (!g) return true;
    if (g.every((n) => n === 0)) return true;                                       // ::
    if (g.slice(0, 7).every((n) => n === 0) && g[7] === 1) return true;            // ::1
    if (g.slice(0, 5).every((n) => n === 0) && (g[5] === 0xffff || g[5] === 0)) return blockedV4([g[6] >> 8, g[6] & 255, g[7] >> 8, g[7] & 255]);   // ::ffff:a.b.c.d、::a.b.c.d
    if ((g[0] & 0xfe00) === 0xfc00) return true;                                    // fc00::/7
    if ((g[0] & 0xffc0) === 0xfe80) return true;                                    // fe80::/10
    if ((g[0] & 0xff00) === 0xff00) return true;                                    // ff00::/8
    if (g[0] === 0x2002) return blockedV4([g[1] >> 8, g[1] & 255, g[2] >> 8, g[2] & 255]);   // 6to4（2002::/16）：內嵌 IPv4 是內網就擋
    if (g[0] === 0x2001 && g[1] === 0) {                                            // Teredo（2001:0000::/32）：伺服器位址（g2,g3）與客戶端位址（g6,g7 取反）任一是內網就擋
      return blockedV4([g[2] >> 8, g[2] & 255, g[3] >> 8, g[3] & 255]) || blockedV4([(~g[6] >> 8) & 255, ~g[6] & 255, (~g[7] >> 8) & 255, ~g[7] & 255]);
    }
    if (g[4] === 0 && g[5] === 0x5efe || g[4] === 0x0200 && g[5] === 0x5efe) return blockedV4([g[6] >> 8, g[6] & 255, g[7] >> 8, g[7] & 255]);   // ISATAP（::5efe:a.b.c.d）
    if (g[0] === 0 && g[1] === 0 && g[2] === 0 && g[3] === 0 && g[4] === 0xffff && g[5] === 0) return blockedV4([g[6] >> 8, g[6] & 255, g[7] >> 8, g[7] & 255]);   // SIIT（::ffff:0:a.b.c.d）
    if (g[0] === 0x64 && g[1] === 0xff9b) return true;                              // 64:ff9b::/32 NAT64（含 64:ff9b:1::/48 本地用）
    if (g[0] === 0x2001 && g[1] === 0x0db8) return true;                            // 文件用
    if (g[0] === 0x0100 && g[1] === 0 && g[2] === 0 && g[3] === 0) return true;     // 100::/64
    return false;
  }
  return true;
}

/** 解析連接埠白名單（環境變數字串，例如 '443,8443'）。無效值丟掉；全部無效或空 → ['443']。 */
function parsePorts(str) {
  const out = String(str == null ? '' : str).split(',').map((x) => x.trim()).filter((x) => /^\d{1,5}$/.test(x) && Number(x) >= 1 && Number(x) <= 65535).map((x) => String(Number(x)));
  return out.length ? Array.from(new Set(out)) : ['443'];
}

// ───────── 網址形狀（不做 DNS）─────────
/**
 * 只驗證網址形狀。回傳 { ok, reason, host, url }。
 * opts.ports：允許的連接埠（字串或數字陣列，預設 ['443']）；網址沒寫埠＝443 一律放行。傳 null 代表不限制。
 */
function checkUrlShape(urlStr, opts) {
  const ports = opts && 'ports' in opts ? opts.ports : ['443'];
  let u;
  try { u = new URL(String(urlStr)); } catch (e) { return { ok: false, reason: 'BAD_URL' }; }
  if (u.protocol !== 'https:') return { ok: false, reason: 'SCHEME_NOT_HTTPS' };
  if (u.username || u.password) return { ok: false, reason: 'CREDENTIALS_IN_URL' };
  if (u.port && ports) {
    const allowed = ports.map(String);
    if (!allowed.includes(u.port)) return { ok: false, reason: allowed.length === 1 && allowed[0] === '443' ? 'PORT_NOT_443' : 'PORT_NOT_ALLOWED' };
  }
  if (!u.hostname) return { ok: false, reason: 'NO_HOST' };
  return { ok: true, host: u.hostname, url: u };
}

// ───────── DNS 解析＋位址檢查 ─────────
function isLocalName(host) {
  const h = host.toLowerCase().replace(/\.$/, '');
  return h === 'localhost' || h.endsWith('.localhost') || h === 'localhost.localdomain';
}
/**
 * 解析主機名稱並檢查「所有」解析結果都是公開位址。
 * 回傳 { ok:true, addrs:[{address}] } 或 { ok:false, reason:'DNS_FAILED'|'BLOCKED_ADDRESS' }。
 * o.lookup(host) -> [{address}]（測試注入；預設 dns.promises.lookup all:true）
 */
async function resolvePublic(hostname, o) {
  o = o || {};
  const host = String(hostname).replace(/^\[|\]$/g, '');
  if (isLocalName(host)) return { ok: false, reason: 'BLOCKED_ADDRESS' };
  let addrs;
  try {
    if (net.isIP(host)) addrs = [{ address: host }];
    else addrs = o.lookup ? await o.lookup(host) : await dns.promises.lookup(host, { all: true });
  } catch (e) { return { ok: false, reason: 'DNS_FAILED' }; }
  if (!Array.isArray(addrs) || !addrs.length) return { ok: false, reason: 'DNS_FAILED' };
  if (addrs.some((a) => isBlockedIp(a.address))) return { ok: false, reason: 'BLOCKED_ADDRESS' };
  return { ok: true, addrs };
}

/** 形狀＋DNS＋位址的完整檢查。回傳 { ok, reason?, message?, url?, host?, addrs? }（不丟例外）。 */
async function validateUrl(urlStr, o) {
  const shape = checkUrlShape(urlStr, o);
  if (!shape.ok) return { ok: false, reason: shape.reason, message: REASON_ZH[shape.reason] };
  const r = await resolvePublic(shape.url.hostname, o);
  if (!r.ok) return { ok: false, reason: r.reason, message: REASON_ZH[r.reason] };
  return { ok: true, url: shape.url, host: shape.host, addrs: r.addrs };
}

// ───────── 安全的 fetch ─────────
const REDIRECT_CODES = new Set([301, 302, 303, 307, 308]);

function normHeaders(h) {
  const out = {};
  if (!h) return out;
  if (typeof h.forEach === 'function' && typeof h.get === 'function') { h.forEach((v, k) => { out[String(k).toLowerCase()] = String(v); }); return out; }
  for (const [k, v] of Object.entries(h)) if (v !== undefined && v !== null) out[k.toLowerCase()] = String(v);
  return out;
}

function timeoutError(ms) {
  const e = new Error('連線逾時（' + Math.round(ms / 1000) + ' 秒）');
  e.name = 'TimeoutError';
  return e;
}

// 對「已驗證的位址」發一次請求（依序嘗試各位址，只在連線階段失敗時換下一個）
function requestOnce(u, addrs, method, headers, body, deadline, o) {
  const request = o.request || https.request;
  const maxBytes = o.maxBytes || 5 * 1024 * 1024;
  return new Promise((resolve, reject) => {
    let idx = 0, settled = false, timer = null, req = null;
    const fail = (e) => { if (!settled) { settled = true; clearTimeout(timer); try { req && req.destroy(); } catch (x) { /* 已關閉 */ } reject(e); } };
    const done = (r) => { if (!settled) { settled = true; clearTimeout(timer); resolve(r); } };
    const remaining = () => deadline - Date.now();
    timer = setTimeout(() => fail(timeoutError(o.timeoutMs || 10000)), Math.max(1, remaining()));
    if (o.signal) {
      if (o.signal.aborted) return fail(o.signal.reason || timeoutError(0));
      o.signal.addEventListener('abort', () => fail(o.signal.reason || timeoutError(0)), { once: true });
    }
    const attempt = () => {
      if (settled) return;
      const addr = addrs[idx].address;
      const hostHeader = u.host;
      const isIpHost = net.isIP(u.hostname.replace(/^\[|\]$/g, ''));
      let gotResponse = false;
      try {
        req = request({
          host: addr, port: u.port ? Number(u.port) : 443, method, agent: false, path: (u.pathname || '/') + (u.search || ''),
          servername: isIpHost ? undefined : u.hostname,
          headers: Object.assign({ 'accept-encoding': 'identity' }, headers, { host: hostHeader, ...(body != null ? { 'content-length': Buffer.byteLength(body) } : {}) }),
        }, (res) => {
          gotResponse = true;
          const chunks = []; let size = 0;
          res.on('data', (c) => { size += c.length; if (size > maxBytes) return fail(new SafeUrlError('RESPONSE_TOO_LARGE')); chunks.push(c); });
          res.on('end', () => done({ status: res.statusCode || 0, headers: res.headers || {}, body: Buffer.concat(chunks), maxBytes }));
          res.on('error', () => fail(new SafeUrlError('REQUEST_FAILED')));
        });
        req.on('error', (e) => {
          if (!gotResponse && idx + 1 < addrs.length && ['ECONNREFUSED', 'ETIMEDOUT', 'EHOSTUNREACH', 'ENETUNREACH'].includes(e && e.code)) { idx++; return attempt(); }
          fail(new SafeUrlError('REQUEST_FAILED', '連線失敗' + (e && e.code ? '（' + e.code + '）' : '')));
        });
        req.end(body != null ? body : undefined);
      } catch (e) { fail(new SafeUrlError('REQUEST_FAILED')); }
    };
    attempt();
  });
}

function wrapResponse(r, finalUrl) {
  const h = r.headers || {};
  return {
    ok: r.status >= 200 && r.status < 300,
    status: r.status,
    statusText: '',
    url: finalUrl,
    headers: { get: (n) => { const v = h[String(n).toLowerCase()]; return Array.isArray(v) ? v.join(', ') : (v === undefined ? null : String(v)); }, has: (n) => h[String(n).toLowerCase()] !== undefined, forEach: (fn) => Object.keys(h).forEach((k) => fn(Array.isArray(h[k]) ? h[k].join(', ') : String(h[k]), k)) },
    text: async () => decodeBody(r).toString('utf8'),
    json: async () => JSON.parse(decodeBody(r).toString('utf8')),
  };
}
// 我們送 accept-encoding: identity；萬一對方仍壓縮，就在上限內解壓（maxOutputLength 防解壓炸彈）
function decodeBody(r) {
  const enc = String((r.headers || {})['content-encoding'] || '').toLowerCase().trim();
  const max = r.maxBytes || 5 * 1024 * 1024;
  try {
    if (enc === 'gzip' || enc === 'x-gzip') return zlib.gunzipSync(r.body, { maxOutputLength: max });
    if (enc === 'deflate') return zlib.inflateSync(r.body, { maxOutputLength: max });
    if (enc === 'br') return zlib.brotliDecompressSync(r.body, { maxOutputLength: max });
  } catch (e) { throw new SafeUrlError('RESPONSE_TOO_LARGE', '回應內容無法解壓或過大'); }
  return r.body;
}

/**
 * fetch 的安全版（回傳物件有 ok／status／headers.get／text()／json()）。
 *   init：{ method, headers, body(string|Buffer), signal }
 *   o：{ ports, lookup, request, timeoutMs(預設 10000，含轉址的總時間), maxBytes(預設 5MB), maxRedirects(預設 3) }（lookup／request 供測試注入）
 * 網址不合規定時丟 SafeUrlError（.code 是 REASON 代碼、.message 是中文）；逾時丟 name='TimeoutError'。
 */
async function safeFetch(urlStr, init, o) {
  init = init || {}; o = o || {};
  const timeoutMs = o.timeoutMs || 10000;
  const deadline = Date.now() + timeoutMs;
  const maxRedirects = o.maxRedirects === undefined ? 3 : o.maxRedirects;
  let method = String(init.method || 'GET').toUpperCase();
  let headers = normHeaders(init.headers);
  let body = init.body != null ? (Buffer.isBuffer(init.body) ? init.body : String(init.body)) : undefined;
  let current = String(urlStr);
  for (let hop = 0; ; hop++) {
    const v = await validateUrl(current, o);
    if (!v.ok) throw new SafeUrlError(v.reason);
    const r = await requestOnce(v.url, v.addrs, method, headers, body, deadline, Object.assign({}, o, { timeoutMs, signal: init.signal }));
    const loc = r.headers && r.headers.location;
    if (REDIRECT_CODES.has(r.status) && loc) {
      if (hop >= maxRedirects) throw new SafeUrlError('TOO_MANY_REDIRECTS');
      let next;
      try { next = new URL(String(loc), v.url); } catch (e) { throw new SafeUrlError('BAD_URL'); }
      if (next.origin !== v.url.origin) { delete headers.authorization; delete headers.cookie; }   // 跨來源轉址不帶憑證
      if (r.status === 303 || ((r.status === 301 || r.status === 302) && method === 'POST')) { method = 'GET'; body = undefined; delete headers['content-type']; delete headers['content-length']; }
      current = next.href;
      continue;
    }
    return wrapResponse(r, v.url.href);
  }
}

module.exports = {
  SafeUrlError, REASON_ZH, isBlockedIp, parsePorts, checkUrlShape, resolvePublic, validateUrl, safeFetch, isLocalName,
};
