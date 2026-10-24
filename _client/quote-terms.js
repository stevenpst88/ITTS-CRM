/*!
 * 報價單品項「服務期間」（License／MA 起訖日）共用規則 —— 前後端單一來源（無相依、不碰 DOM、時區無關）
 *   瀏覽器：<script src="quote-terms.js"> → window.QTERMS
 *   Node  ：require('../_client/quote-terms.js')（lib/quoteTerms.js 包一層伺服器專用函式；Vercel 的 includeFiles 已含 _client/**）
 *
 * 規則（MA_TRACKING=1 才啟用，開關由伺服器端 lib/quoteTerms.js 判斷；本檔只有純函式）：
 *   · 類型 kind：'' 一般｜'license'｜'ma'。報價品項只存 license／ma（一般品項不存欄位）。商品目錄另有 'none'（＝一般）。
 *   · 起迄日：YYYY-MM-DD，必須是真實存在的日期（2026-02-30 不行）；迄日不得早於起日（同一天可以，期間 1 天）。
 *   · 期間長度（「共 N 個月」）：起日到「迄日隔天」的整月數；零頭天數另列。2026-01-01～2026-12-31＝共 12 個月；
 *     2026-01-15～2027-01-14＝共 12 個月；2026-01-01～2026-01-20＝共 20 天；2026-01-01～2026-02-10＝共 1 個月 10 天。
 *     月底：起日＋N 個月落在不存在的日子時取該月最後一天（1/31＋1 個月＝2/28），所以一年快捷往回算必得整數個月（見 quickEnd）。
 *   · 快捷：一年＝起日＋12 個月－1 天；三年＝起日＋36 個月－1 天（2 月 29 日起算 +12 個月落在 2 月 28 日，再 －1 天）。
 *   · 顯示：「YYYY/MM/DD – YYYY/MM/DD」（中間是 en dash，前後各一個半形空白）；客戶文件品名下一行：「服務期間：…」。
 *   全部用 Date.UTC 與整數運算，不讀本機時區；不接受 Date 物件，只接受字串。
 * scripts/check-quote-terms.js 對本檔做單元＋隨機測試（含用不同 TZ 環境變數各跑一次）。
 */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module && module.exports) module.exports = api;
  else if (root) root.QTERMS = api;
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this), function () {
  'use strict';

  var KINDS = ['license', 'ma'];
  var KIND_LABEL = { '': '一般', license: 'License', ma: 'MA' };
  var YEAR_MIN = 1900, YEAR_MAX = 2999;
  var DATE_RE = /^(\d{4})-(\d{2})-(\d{2})$/;
  var EN_DASH = '–';

  /** 報價品項的類型：'license'|'ma'|''（''／null／undefined／'none'＝一般）；其他任何值回 null（不認得） */
  function parseKind(v) {
    if (v === undefined || v === null) return '';
    if (typeof v !== 'string') return null;
    var s = v.trim().toLowerCase();
    if (s === '' || s === 'none') return '';
    return KINDS.indexOf(s) >= 0 ? s : null;
  }
  /** 商品目錄的類型：'none'|'license'|'ma'；不認得回 null */
  function parseCatalogKind(v) {
    var k = parseKind(v);
    return k === null ? null : (k === '' ? 'none' : k);
  }

  function daysInMonth(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }   // m: 1..12
  /** 是否為真實存在的 YYYY-MM-DD（只收字串；年限 1900～2999） */
  function isRealDate(s) {
    if (typeof s !== 'string') return false;
    var m = DATE_RE.exec(s);
    if (!m) return false;
    var y = +m[1], mo = +m[2], d = +m[3];
    if (y < YEAR_MIN || y > YEAR_MAX || mo < 1 || mo > 12 || d < 1) return false;
    return d <= daysInMonth(y, mo);
  }
  function pad(n, w) { var s = String(n); while (s.length < w) s = '0' + s; return s; }
  function iso(y, m, d) { return pad(y, 4) + '-' + pad(m, 2) + '-' + pad(d, 2); }
  function parts(s) { var m = DATE_RE.exec(s); return { y: +m[1], m: +m[2], d: +m[3] }; }
  function dayNumber(s) { var p = parts(s); return Math.round(Date.UTC(p.y, p.m - 1, p.d) / 86400000); }

  /** 加 n 天（可負）。輸入不合法回 ''；結果超出年限回 '' */
  function addDays(s, n) {
    if (!isRealDate(s) || typeof n !== 'number' || !isFinite(n)) return '';
    var p = parts(s), t = new Date(Date.UTC(p.y, p.m - 1, p.d + Math.trunc(n)));
    var out = iso(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate());
    return isRealDate(out) ? out : '';
  }
  /** 加 n 個月（可負）；日期超出該月天數時落在該月最後一天（1/31＋1 個月＝2/28 或 2/29）。不合法回 '' */
  function addMonths(s, n) {
    if (!isRealDate(s) || typeof n !== 'number' || !isFinite(n)) return '';
    var p = parts(s), total = p.y * 12 + (p.m - 1) + Math.trunc(n);
    var ny = Math.floor(total / 12), nm = total - ny * 12 + 1;
    var out = iso(ny, nm, Math.min(p.d, daysInMonth(ny, nm)));
    return isRealDate(out) ? out : '';
  }
  /** 快捷：起日＋years 年－1 天（years 預設 1；一年／三年按鈕分別傳 1、3）。起日不合法回 '' */
  function quickEnd(from, years) {
    var y = years === undefined ? 1 : years;
    var e = addMonths(from, 12 * y);
    return e ? addDays(e, -1) : '';
  }

  /**
   * 期間長度：{ months, days }（整月＋零頭天）。起迄不合法或迄日早於起日回 null。
   * 以「迄日隔天」為右端（含頭含尾計）：from=2026-01-01、to=2026-12-31 → 右端 2027-01-01 → 12 個月 0 天。
   */
  function span(from, to) {
    if (!isRealDate(from) || !isRealDate(to) || to < from) return null;
    var e = addDays(to, 1);
    if (!e) return null;
    var f = parts(from), x = parts(e);
    var m = (x.y - f.y) * 12 + (x.m - f.m);
    // 往回找到「起日＋m 個月」不晚於右端的最大 m（addMonths 對月底取該月最後一天，所以 1/31＋1 個月＝2/28）
    while (m > 0 && addMonths(from, m) > e) m -= 1;
    if (m < 0) m = 0;
    var anchor = addMonths(from, m);
    var days = dayNumber(e) - dayNumber(anchor);
    if (days < 0) days = 0;
    return { months: m, days: days };
  }
  /** 「共 N 個月」「共 N 個月 D 天」「共 D 天」；不合法回 '' */
  function spanText(from, to) {
    var s = span(from, to);
    if (!s) return '';
    if (!s.months) return '共 ' + s.days + ' 天';
    return '共 ' + s.months + ' 個月' + (s.days ? ' ' + s.days + ' 天' : '');
  }

  function slash(s) { return s.slice(0, 4) + '/' + s.slice(5, 7) + '/' + s.slice(8, 10); }
  /** 「YYYY/MM/DD – YYYY/MM/DD」；任一端不合法回 '' */
  function formatRange(from, to) {
    if (!isRealDate(from) || !isRealDate(to)) return '';
    return slash(from) + ' ' + EN_DASH + ' ' + slash(to);
  }
  /** 客戶文件品名下方那一行：「服務期間：YYYY/MM/DD – YYYY/MM/DD」；沒有類型或期間不合法回 '' */
  function servicePeriodLine(kind, from, to) {
    if (parseKind(kind) === '' || parseKind(kind) === null) return '';
    var r = formatRange(from, to);
    return r ? '服務期間：' + r : '';
  }

  /**
   * 驗證一列的期間。kind 為 '' 時不需要期間（回 null）。
   * 回 null＝通過；否則 { code:'TERM_REQUIRED'|'TERM_INVALID', reason }（reason 是不含列號的中文原因，呼叫端自行加「第 N 列」）
   */
  function validateTerm(kind, from, to) {
    var k = parseKind(kind);
    if (k === null) return { code: 'TERM_INVALID', reason: '類型不正確（只能是一般、License 或 MA）' };
    if (!k) return null;
    var f = typeof from === 'string' ? from.trim() : '', t = typeof to === 'string' ? to.trim() : '';
    if (!f || !t) return { code: 'TERM_REQUIRED', reason: '是 ' + KIND_LABEL[k] + ' 品項，請填寫服務期間的起日與迄日' };
    if (!isRealDate(f) || !isRealDate(t)) return { code: 'TERM_INVALID', reason: '服務期間的起日或迄日不是有效日期（格式 YYYY-MM-DD，且必須真實存在）' };
    if (t < f) return { code: 'TERM_INVALID', reason: '服務期間的迄日不得早於起日' };
    return null;
  }

  /**
   * 建議的類型（只是預選，使用者可改）。name＝品名／商品名；groupTitle＝商品目錄群組標題（可省略）。
   *   群組標題含「License MA」或「Service MA」→ ma；名稱含獨立的 MA（前後不是英文字母）或「維護」→ ma；
   *   名稱含 License 或「授權」→ license；其餘 ''。MA 優先於 License（「SAP PCE License MA」是 License 維護＝ma）。
   */
  function suggestKind(name, groupTitle) {
    var g = typeof groupTitle === 'string' ? groupTitle : '';
    if (/license\s*ma|service\s*ma/i.test(g)) return 'ma';
    var n = typeof name === 'string' ? name : '';
    if (/(^|[^A-Za-z])MA([^A-Za-z]|$)/i.test(n) || n.indexOf('維護') >= 0) return 'ma';
    if (/license/i.test(n) || n.indexOf('授權') >= 0) return 'license';
    return '';
  }

  return {
    KINDS: KINDS, KIND_LABEL: KIND_LABEL, YEAR_MIN: YEAR_MIN, YEAR_MAX: YEAR_MAX,
    parseKind: parseKind, parseCatalogKind: parseCatalogKind, isRealDate: isRealDate,
    addDays: addDays, addMonths: addMonths, quickEnd: quickEnd, span: span, spanText: spanText,
    formatRange: formatRange, servicePeriodLine: servicePeriodLine, validateTerm: validateTerm, suggestKind: suggestKind,
  };
});
