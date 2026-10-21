'use strict';
/**
 * CSP 強化第一期：Report-Only 政策＋違規回報端點＋稽核摘要。
 *
 * 目標：先「只觀察」。新增 Content-Security-Policy-Report-Only 標頭（目標嚴格政策），不影響現有強制 CSP（helmet 那條）。
 * 瀏覽器發現違規時 POST 到 /api/csp-report，本模組清洗後去重彙總寫入稽核（CSP_VIOLATION），供後台摘要與後續分期盤點。
 *
 * 開關（環境變數；都是「重新部署／重啟即生效」）：
 *   CSP_REPORT_ONLY        on（預設）｜off。off＝完全不加標頭、不註冊回報端點（與導入前位元級相同）
 *   CSP_REPORT_TO          off（預設）｜on。on＝另加 report-to／Reporting-Endpoints（Chromium 會因此忽略 report-uri，上線前先小範圍確認有送達）
 *   CSP_REPORT_ONLY_STYLE  on｜off（預設）。on＝多加一條只管 style-src 'self' 的 Report-Only 政策
 *                          （現有頁面約 2,000 個 style= 屬性，會產生大量回報，只在要盤點樣式時短暫開啟）
 *   CSP_REPORT_MAX_WRITES_PER_HOUR  本實例每小時最多寫入稽核幾筆（預設：稽核新表 120、舊路徑 4）
 *
 * 回報端點的安全設計（POST /api/csp-report，公開：登入頁也要能回報）：
 *   - 只收 application/csp-report 與 application/reports+json；拒絕壓縮編碼；本文上限 32 KB；每批最多處理 25 筆
 *   - 每 IP 每分鐘 120 次、整個實例每分鐘 600 次；超過回 429（不讀本文）
 *   - 不信任任何欄位：指令名稱走白名單；blocked-uri 只留 origin 或關鍵字；網址只留 path（去查詢字串／片段）；
 *     script-sample 只留長度；所有字串過濾控制字元與 < > " '，並限長
 *   - 瀏覽器擴充套件（chrome-extension:// 等）造成的違規直接忽略（不是我們的程式碼）
 *   - 同一「指令＋來源＋頁面＋來源檔」10 分鐘只寫一筆，期間再發生只計數，下一筆會註明「前一窗合併 N 次」
 *   - 稽核寫入有每小時預算；超過就只記憶體計數並寫一筆 CSP_REPORT_THROTTLED（每小時最多一筆），不會洗版、不會拖累寫入佇列
 *   - 不回傳任何內容（204／4xx 空本文）
 */
const ENDPOINT = '/api/csp-report';
const REPORT_GROUP = 'csp-endpoint';
const MAX_BODY_BYTES = 32 * 1024;
const MAX_REPORTS_PER_REQUEST = 25;
const DEDUP_WINDOW_MS = 10 * 60 * 1000;
const DEDUP_MAX_KEYS = 1000;
const IP_LIMIT_PER_MIN = 120;
const GLOBAL_LIMIT_PER_MIN = 600;
const HOUR_MS = 60 * 60 * 1000;
// 資安審查 M4：只有我方實際提供的頁面才會入稽核；回報裡的 document-uri 是攻擊者可任意變造的字串，
// 不在這份清單內的只在記憶體計數（stats.unknownDoc），不寫稽核、不佔去重表。新增 HTML 頁面時要同步加進來。
const KNOWN_DOC_PATHS = new Set(['/', '/index.html', '/login.html', '/admin.html', '/help.html']);
// 舊路徑（整包 JSONB 稽核、FIFO 5000 筆）每小時預算：過去 24 筆會讓匿名者在數天內擠掉真實稽核事件，降為 4
const LEGACY_MAX_WRITES = 4;

const ACTION = 'CSP_VIOLATION';
const THROTTLE_ACTION = 'CSP_REPORT_THROTTLED';
const OPERATOR = '(瀏覽器回報)';

const TYPE_LEGACY = 'application/csp-report';
const TYPE_API = 'application/reports+json';

// ───────────────────────── 政策 ─────────────────────────
function envOn(v, dflt) {
  if (v == null || String(v).trim() === '') return dflt;
  return !/^(off|0|false|no)$/i.test(String(v).trim());
}
function reportOnlyEnabled(env) { return envOn(env && env.CSP_REPORT_ONLY, true); }
function styleReportEnabled(env) { return reportOnlyEnabled(env) && /^(on|1|true|yes)$/i.test(String((env && env.CSP_REPORT_ONLY_STYLE) || '').trim()); }

function reportToEnabled(env) { return envOn(env && env.CSP_REPORT_TO, false); }
/**
 * 預設只送舊式 report-uri：Chromium／Firefox／Safari 都認，違規當下立即 POST。
 * report-to（新式 Reporting API）設 CSP_REPORT_TO=on 才加：Chromium 只要看到 report-to 就會忽略 report-uri，
 * 而本期在無頭 Edge（http://localhost）實測 report-to 150 秒內零送達，所以預設不讓整條回報管線押在它身上；Firefox／Safari 本來就不支援 report-to。
 */
function reportTail(env) { return reportToEnabled(env) ? `report-uri ${ENDPOINT}; report-to ${REPORT_GROUP}` : `report-uri ${ENDPOINT}`; }
/**
 * 目標嚴格政策（第一期只觀察）：
 *  - script-src 'self' 加 jsdelivr：index.html（Chart.js）、admin.html（xlsx）仍從 CDN 載入，盤點確認後才列入；第二期自架後移除
 *  - script-src-attr 'none'：不允許 onclick= 等行內事件屬性（目標）
 *  - style-src 暫留 'unsafe-inline'：現有頁面約 2,000 個 style= 與 4 個 <style> 區塊，短期內不可能清掉；另有可選的嚴格 style 政策
 *  - 沒有 upgrade-insecure-requests：Report-Only 中會被瀏覽器忽略並噴警告
 */
function buildPolicy(env) {
  return [
    "default-src 'self'",
    "script-src 'self' https://cdn.jsdelivr.net",
    "script-src-attr 'none'",
    "style-src 'self' 'unsafe-inline'",
    "img-src 'self' data: blob:",
    "font-src 'self'",
    "connect-src 'self'",
    "frame-src 'none'",
    "object-src 'none'",
    "frame-ancestors 'none'",
    "base-uri 'self'",
    "form-action 'self'",
    reportTail(env),
  ].join('; ');
}
function buildStylePolicy(env) { return `style-src 'self'; ${reportTail(env)}`; }

/** 回傳要加的標頭（物件）；off 時回 null */
function headersFor(env) {
  if (!reportOnlyEnabled(env)) return null;
  const pols = [buildPolicy(env)];
  if (styleReportEnabled(env)) pols.push(buildStylePolicy(env));
  const out = { 'Content-Security-Policy-Report-Only': pols.join(', ') };
  if (reportToEnabled(env)) out['Reporting-Endpoints'] = `${REPORT_GROUP}="${ENDPOINT}"`;
  return out;
}
/** Express 中介層：HTML／靜態檔加 Report-Only 標頭；/api/ 的 JSON 與 /q/ 跳板頁（自帶更嚴的 CSP）不加 */
function reportOnlyMiddleware(env) {
  const h = headersFor(env);
  if (!h) return null;
  const entries = Object.entries(h);
  return (req, res, next) => {
    const p = req.path || '';
    if (!p.startsWith('/api/') && !p.startsWith('/q/')) for (const [k, v] of entries) res.setHeader(k, v);
    next();
  };
}

// ───────────────────────── 清洗 ─────────────────────────
const DIRECTIVES = new Set(['default-src', 'script-src', 'script-src-elem', 'script-src-attr', 'style-src', 'style-src-elem', 'style-src-attr',
  'img-src', 'font-src', 'connect-src', 'media-src', 'object-src', 'frame-src', 'child-src', 'worker-src', 'manifest-src', 'prefetch-src',
  'base-uri', 'form-action', 'frame-ancestors', 'navigate-to', 'require-trusted-types-for', 'trusted-types', 'sandbox']);
const BLOCKED_KEYWORDS = new Set(['inline', 'eval', 'data', 'blob', 'self', 'about', 'wasm-eval', 'trusted-types-policy', 'trusted-types-sink', 'inline-speculation-rules', 'filesystem']);
const EXT_SCHEMES = new Set(['chrome-extension', 'moz-extension', 'safari-extension', 'safari-web-extension', 'ms-browser-extension', 'edge-extension', 'webkit-masked-url', 'chrome']);

function str(v, max) { return typeof v === 'string' ? v.slice(0, max) : ''; }
function scheme(s) { const m = /^([a-z][a-z0-9+.-]{0,30}):/i.exec(s); return m ? m[1].toLowerCase() : ''; }

function directiveOf(v) {
  const t = str(v, 200).trim().split(/\s+/)[0].toLowerCase();
  return DIRECTIVES.has(t) ? t : '(未知指令)';
}
function cleanPath(p) {
  // eslint-disable-next-line no-control-regex
  return String(p).replace(/[\u0000-\u001f\u007f]/g, '').replace(/[^A-Za-z0-9/_\-.~%@+:,;=]/g, '_').slice(0, 120) || '/';
}
/** 網址 → 只留 path（去查詢字串與片段）；非 http(s) 網址只留 scheme 括號 */
function pathOf(v) {
  const s = str(v, 2048).trim();
  if (!s) return '';
  const sc = scheme(s);
  if (sc) {
    if (sc !== 'http' && sc !== 'https') return '(' + sc + ')';
    try { return cleanPath(new URL(s).pathname); } catch (e) { return '(無法解析)'; }
  }
  if (s.startsWith('/')) return cleanPath(s.split(/[?#]/)[0]);
  return '(無法解析)';
}
/** blocked-uri → { v: 字串, ext: 是否來自瀏覽器擴充套件 } 。URL 只留 origin；data:/blob: 只留關鍵字 */
function blockedOf(v) {
  const s = str(v, 2048).trim();
  if (!s) return { v: '(空)', ext: false };
  const low = s.toLowerCase();
  if (BLOCKED_KEYWORDS.has(low)) return { v: low, ext: false };
  const sc = scheme(s);
  if (!sc) return { v: '(其他)', ext: false };
  if (EXT_SCHEMES.has(sc)) return { v: '(擴充套件)', ext: true };
  if (sc === 'http' || sc === 'https' || sc === 'ws' || sc === 'wss') {
    try {
      const u = new URL(s);
      if (!/^[a-z0-9.-]{1,100}$/i.test(u.hostname)) return { v: '(其他)', ext: false };
      return { v: `${sc}://${u.hostname.toLowerCase()}${u.port ? ':' + /^\d{1,5}$/.exec(u.port)[0] : ''}`, ext: false };
    } catch (e) { return { v: '(其他)', ext: false }; }
  }
  return { v: BLOCKED_KEYWORDS.has(sc) ? sc : '(' + sc + ')', ext: false };
}
function intOf(v) { const n = Number(v); return Number.isFinite(n) && n >= 0 && n <= 1e7 ? Math.floor(n) : 0; }
function dispOf(v) { const s = str(v, 20).toLowerCase(); return s === 'enforce' ? 'enforce' : 'report'; }

function normOne(o) {
  const b = blockedOf(o.blocked);
  return {
    directive: directiveOf(o.directive),
    blocked: b.v, ext: b.ext,
    docPath: pathOf(o.doc) || '(未知)',
    srcPath: pathOf(o.src),
    line: intOf(o.line),
    sampleLen: Math.min(typeof o.sample === 'string' ? o.sample.length : 0, 9999),
    disposition: dispOf(o.disposition),
  };
}
function fromLegacy(r) {
  if (!r || typeof r !== 'object' || Array.isArray(r)) return null;
  return normOne({ directive: r['effective-directive'] || r['violated-directive'], blocked: r['blocked-uri'], doc: r['document-uri'], src: r['source-file'],
    line: r['line-number'], sample: r['script-sample'], disposition: r.disposition });
}
function fromApi(item) {
  if (!item || typeof item !== 'object' || item.type !== 'csp-violation') return null;
  const b = item.body;
  if (!b || typeof b !== 'object' || Array.isArray(b)) return null;
  return normOne({ directive: b.effectiveDirective || b.violatedDirective, blocked: b.blockedURL, doc: b.documentURL || item.url, src: b.sourceFile,
    line: b.lineNumber, sample: b.sample, disposition: b.disposition });
}
/** 依內容類型取出已清洗的回報；形狀不符回 null（呼叫端回 400） */
function extractReports(kind, parsed) {
  if (kind === 'legacy') {
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed) || !parsed['csp-report']) return null;
    const r = fromLegacy(parsed['csp-report']);
    return r ? [r] : null;
  }
  if (!Array.isArray(parsed)) return null;
  return parsed.slice(0, MAX_REPORTS_PER_REQUEST).map(fromApi).filter(Boolean);
}
function contentKind(ct) {
  const t = String(ct || '').split(';')[0].trim().toLowerCase();
  return t === TYPE_LEGACY ? 'legacy' : t === TYPE_API ? 'api' : '';
}

// ───────────────────────── 稽核文字格式（寫入與摘要共用） ─────────────────────────
function targetOf(r) { return `${r.directive} ← ${r.blocked}`; }
function detailOf(r, merged) {
  const parts = [`指令=${r.directive}`, `來源=${r.blocked}`, `頁面=${r.docPath}`];
  if (r.srcPath) parts.push(`檔案=${r.srcPath}`);
  parts.push(`行=${r.line}`, `樣本長=${r.sampleLen}`, `前一窗合併=${merged}`, `處置=${r.disposition}`);
  return parts.join('；');
}
function parseDetail(d) {
  const o = {};
  String(d || '').split('；').forEach((kv) => { const i = kv.indexOf('='); if (i > 0) o[kv.slice(0, i)] = kv.slice(i + 1); });
  return o;
}

// ───────────────────────── 摘要（純函式） ─────────────────────────
function taipeiDate(ms) { return new Date(ms + 8 * 3600e3).toISOString().slice(0, 10); }
function topN(map, n) { return Array.from(map.entries()).sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0]))).slice(0, n).map(([name, count]) => ({ name, count })); }
/** rows＝稽核內部列（action=CSP_VIOLATION 與 CSP_REPORT_THROTTLED 皆可混在內）。權重＝1＋前一窗合併次數 */
function summarize(rows, opts) {
  const o = opts || {};
  const dirs = new Map(), srcs = new Map(), pages = new Map();
  let events = 0, merged = 0, throttled = 0, lastAt = '';
  for (const r of rows || []) {
    if (r.action === THROTTLE_ACTION) { throttled++; continue; }
    if (r.action !== ACTION) continue;
    const d = parseDetail(r.detail);
    const m = Math.max(0, Math.min(1e6, parseInt(d['前一窗合併'], 10) || 0));
    const w = 1 + m;
    events++; merged += m;
    const dir = d['指令'] || '(未知指令)', src = d['來源'] || '(空)', page = d['頁面'] || '(未知)';
    dirs.set(dir, (dirs.get(dir) || 0) + w);
    srcs.set(src, (srcs.get(src) || 0) + w);
    pages.set(page, (pages.get(page) || 0) + w);
    if (r.ts && r.ts > lastAt) lastAt = r.ts;
  }
  return { events, merged, total: events + merged, throttled, lastAt, topDirectives: topN(dirs, 5), topSources: topN(srcs, 5), topPages: topN(pages, 5) };
}

// ───────────────────────── 端點 ─────────────────────────
function create(opts) {
  const o = opts || {};
  const env = o.env || process.env;
  const now = o.now || Date.now;
  const writeLog = o.writeLog;
  const holdForFlush = o.holdForFlush || (() => {});
  const auditMode = o.auditMode || (() => 'legacy');
  const stats = { received: 0, accepted: 0, written: 0, deduped: 0, ignoredExtension: 0, unknownDoc: 0, invalid: 0, rateLimited: 0, budgetDropped: 0, mapFullDropped: 0, throttleMarkers: 0, lastWriteAt: null };

  const dedup = new Map();           // key -> { t: 視窗起點, n: 視窗內尚未寫入的合併次數 }
  const ipHits = new Map();          // ip -> { w: 視窗起點, n }
  const glob = { w: 0, n: 0 };
  const budget = { w: 0, used: 0, markedAt: 0 };

  function maxWrites() {
    const v = parseInt(env.CSP_REPORT_MAX_WRITES_PER_HOUR, 10);
    if (Number.isFinite(v) && v >= 0) return v;
    return auditMode() === 'table' ? 120 : LEGACY_MAX_WRITES;
  }
  function rateOk(ip) {
    const t = now();
    if (t - glob.w >= 60000) { glob.w = t; glob.n = 0; }
    if (glob.n >= (parseInt(env.CSP_REPORT_GLOBAL_PER_MIN, 10) || GLOBAL_LIMIT_PER_MIN)) return false;
    let e = ipHits.get(ip);
    if (!e || t - e.w >= 60000) {
      e = { w: t, n: 0 };
      if (ipHits.size >= 2000) { for (const [k, v] of ipHits) if (t - v.w >= 60000) ipHits.delete(k); if (ipHits.size >= 2000) return false; }
      ipHits.set(ip, e);
    }
    if (e.n >= (parseInt(env.CSP_REPORT_IP_PER_MIN, 10) || IP_LIMIT_PER_MIN)) return false;
    e.n++; glob.n++;
    return true;
  }
  function budgetTake() {
    const t = now();
    if (t - budget.w >= HOUR_MS) { budget.w = t; budget.used = 0; }
    if (budget.used >= maxWrites()) return false;
    budget.used++;
    return true;
  }
  function writeRow(req, r, merged) {
    holdForFlush(req);
    writeLog(ACTION, OPERATOR, targetOf(r), detailOf(r, merged), req);
    stats.written++; stats.lastWriteAt = new Date(now()).toISOString();
  }
  function markThrottled(req) {
    const t = now();
    if (budget.markedAt && t - budget.markedAt < HOUR_MS) return;
    budget.markedAt = t;
    stats.throttleMarkers++;
    holdForFlush(req);
    writeLog(THROTTLE_ACTION, OPERATOR, '本實例', `CSP 違規回報的稽核寫入已達每小時上限 ${maxWrites()} 筆；其餘回報只在記憶體計數（不入稽核）。累計丟棄 ${stats.budgetDropped} 筆`, req);
  }

  /** 處理一筆已清洗的回報。回傳處理結果字串（測試用） */
  function process_(req, r) {
    if (r.ext) { stats.ignoredExtension++; return 'extension'; }
    const t = now();
    if (!KNOWN_DOC_PATHS.has(r.docPath)) { stats.unknownDoc++; return 'unknown-doc'; }   // 資安審查 M4：來源頁面不在已知清單 → 只計數
    // 去重鍵：指令＋來源＋頁面（頁面已限於已知清單）；不含 srcPath——它是攻擊者可任意變造的字串，用來灌出無限多個「不同」事件
    const key = [r.directive, r.blocked, r.docPath].join('|');
    const e = dedup.get(key);
    if (e && t - e.t < DEDUP_WINDOW_MS) { e.n++; stats.deduped++; return 'deduped'; }
    const carry = e ? e.n : 0;
    if (!e && dedup.size >= DEDUP_MAX_KEYS) {
      for (const [k, v] of dedup) if (t - v.t >= DEDUP_WINDOW_MS) dedup.delete(k);
      if (dedup.size >= DEDUP_MAX_KEYS) { stats.mapFullDropped++; return 'map-full'; }
    }
    if (!budgetTake()) {
      dedup.set(key, { t, n: carry + 1 });   // 不寫稽核，但把次數帶到下一個視窗
      stats.budgetDropped++;
      markThrottled(req);
      return 'budget';
    }
    dedup.set(key, { t, n: 0 });
    writeRow(req, r, carry);
    return 'written';
  }

  // ── Express 處理鏈 ──
  function precheck(req, res, next) {
    const cl = req.headers['content-length'];
    if (cl !== undefined && Number(cl) > MAX_BODY_BYTES) { stats.invalid++; return res.status(413).end(); }
    const enc = String(req.headers['content-encoding'] || '').trim().toLowerCase();
    if (enc && enc !== 'identity') { stats.invalid++; return res.status(415).end(); }
    const kind = contentKind(req.headers['content-type']);
    if (!kind) { stats.invalid++; return res.status(415).end(); }
    req._cspKind = kind;
    next();
  }
  function limiter(req, res, next) {
    stats.received++;
    if (!rateOk(req.ip || 'unknown')) { stats.rateLimited++; res.set('Retry-After', '60'); return res.status(429).end(); }
    next();
  }
  function handle(req, res) {
    const buf = req.body;
    if (!Buffer.isBuffer(buf) || buf.length === 0) { stats.invalid++; return res.status(400).end(); }
    let parsed;
    try { parsed = JSON.parse(buf.toString('utf8')); } catch (e) { stats.invalid++; return res.status(400).end(); }
    const list = extractReports(req._cspKind, parsed);
    if (!list) { stats.invalid++; return res.status(400).end(); }
    stats.accepted += list.length;
    for (const r of list) {
      try { process_(req, r); } catch (e) { console.error('[csp-report]', e && e.message); }
    }
    res.status(204).end();
  }
  function onError(err, req, res, next) {   // 路由專屬錯誤處理：本文過大／被中斷等一律空本文
    stats.invalid++;
    if (res.headersSent) return;
    const code = err && (err.status || err.statusCode);
    res.status(code === 413 ? 413 : 400).end();
  }
  function register(app, express) {
    app.post('/api/csp-report', limiter, precheck, express.raw({ type: () => true, limit: MAX_BODY_BYTES, inflate: false }), handle, onError);
    // 其他方法：405（用 use 而非 all，避免被路由表列舉器當成 _ALL 路由）
    app.use(ENDPOINT, (req, res) => { res.set('Allow', 'POST'); res.status(405).end(); });
  }

  return {
    register, stats: () => Object.assign({}, stats, { dedupKeys: dedup.size, writesThisHour: budget.used, maxWritesPerHour: maxWrites() }),
    _process: process_, _handle: handle, _limiter: limiter, _precheck: precheck,
    _reset() { dedup.clear(); ipHits.clear(); glob.w = glob.n = 0; budget.w = budget.used = budget.markedAt = 0; },
  };
}

module.exports = {
  ENDPOINT, REPORT_GROUP, MAX_BODY_BYTES, MAX_REPORTS_PER_REQUEST, DEDUP_WINDOW_MS, ACTION, THROTTLE_ACTION, OPERATOR,
  reportOnlyEnabled, styleReportEnabled, reportToEnabled, buildPolicy, buildStylePolicy, headersFor, reportOnlyMiddleware,
  directiveOf, pathOf, blockedOf, normOne, fromLegacy, fromApi, extractReports, contentKind, targetOf, detailOf, parseDetail, summarize, taipeiDate, create,
};
