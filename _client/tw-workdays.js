/*!
 * 台灣工作天與「報價期限預設值」規則 —— 前後端單一來源（無相依、不碰 DOM）
 *   瀏覽器：<script src="tw-workdays.js"> → window.TWWD
 *   Node  ：require('../_client/tw-workdays.js')（lib/quoteExcel.js 用它；Vercel 的 includeFiles 已含 _client/**）
 *
 * ── 報價期限預設規則 ─────────────────────────────────────────────────────────
 *   報價日期為 D。L ＝ D 所在月份的「最後一個工作天」。
 *   remaining ＝ 「D 之後的隔天起、到 L（含）」之間的工作天數（**不含 D 當天**）。
 *   remaining < MIN_VALID_WORKING_DAYS（7）→ 預設期限＝「下個月」的最後一個工作天；否則＝ L。
 *   D 已在 L 當天或之後（月底週末／假日／最後一個工作天當天建單）時 remaining 為 0，自然落在順延。
 *   要調整門檻（7）或「是否算當天」，只改本檔的 MIN_VALID_WORKING_DAYS / workingDaysAfter，前後端自動一致。
 *   這只是「預設值」：報價期限是業務自行輸入的欄位，可以改成任何不早於報價日期的日期。
 *
 * ── 工作天定義 ────────────────────────────────────────────────────────────────
 *   週一至週五且不是國定假日／補假；週六、週日不是工作天，除非列在 TW_GOV_MAKEUP_SATURDAYS（週六補班日）。
 *
 * ── 假日資料與來源 ────────────────────────────────────────────────────────────
 *   1) 2026（民國115年）、2027（民國116年）：人事行政總處公告的辦公日曆表逐日資料（週一至週五放假日）。
 *      官方 CSV 逐列解析＋官方日曆 PDF 目視核對＋新聞稿三方吻合；查證日 2026-10-08。
 *      來源：https://www.dgpa.gov.tw/information?uid=30&pid=12573（115年）、…pid=12982（116年，115.05.21 公告）、
 *            https://data.gov.tw/dataset/14718（官方 CSV）。
 *      兩年都沒有週六補班日（人事行政總處自 114.06.13 起刪除「調整放假並補行上班」）。
 *      有逐日表的年度，**表就是全部真相**（含春節、清明、端午、中秋與各補假），不再另加規則推算。
 *   2) 其他年份（含 2024、2025、2028 年以後）：只用 8 個「固定日期」國定假日加補假規則推算——
 *      1/1 開國紀念日、2/28 和平紀念日、4/4 兒童節、5/1 勞動節、9/28 教師節（孔子誕辰紀念日）、10/10 國慶日、
 *      10/25 光復節、12/25 行憲紀念日（依 114.05.28 公布的《紀念日及節日實施條例》；全部適用於所有年份，不分舊制）。
 *      補假規則（條例第 8 條、政府機關配合紀念日與節日補假及調整放假處理要點）：放假日逢週六 → 前一個上班日補假、
 *      逢週日 → 次一個上班日補假；補假可跨年（2028-01-01 逢週六 → 2027-12-31，已在 2027 逐日表內）。
 *      固定日期之間不會互撞，所以「前一個上班日／次一個上班日」對它們等價於「前一個週五／後一個週一」。
 *
 * ── 已知限制（保守做法：未知的一律當工作天，不猜）──────────────────────────────
 *   - 春節（小年夜～初三）、清明（4/4 或 4/5）、端午、中秋日期逐年變動，無法推算。沒有逐日表的年度不含它們，
 *     這些年度若報價日期後的月底剛好有這些連假，預設日會偏（偏晚），業務要自行修改報價期限。
 *   - 2028 年辦公日曆尚未公布（人事行政總處須於 2027-06-30 前公告）→ 公布後請把 2028 年加進 TW_GOV_WEEKDAY_HOLIDAYS，
 *     並跑 scripts/check-valid-until.js（它會檢查：筆數、全是週一至週五、不重複、固定日期推算結果都在表內）。
 *   - 這是政府行政機關的辦公日曆；公司實際的休假日若與政府不同（例如補假），預設日會有出入，業務可自行修改。
 *   - 原住民族歲時祭儀放假日未納入。
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module && module.exports) module.exports = api;
  else if (root) root.TWWD = api;
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this), function () {
  'use strict';

  /** 預設報價期限至少要保留的工作天數（報價日期之後、不含當天）。不足就順延到下個月的最後一個工作天 */
  const MIN_VALID_WORKING_DAYS = 7;

  /** 固定日期國定假日 [月, 日] */
  const FIXED_HOLIDAYS = [[1, 1], [2, 28], [4, 4], [5, 1], [9, 28], [10, 10], [10, 25], [12, 25]];

  /** 有官方逐日表的年度：週一至週五不上班的日期（含補假） */
  const TW_GOV_WEEKDAY_HOLIDAYS = {
    2026: [
      '2026-01-01', '2026-02-16', '2026-02-17', '2026-02-18', '2026-02-19', '2026-02-20',
      '2026-02-27', '2026-04-03', '2026-04-06', '2026-05-01', '2026-06-19', '2026-09-25',
      '2026-09-28', '2026-10-09', '2026-10-26', '2026-12-25'
    ],
    2027: [
      '2027-01-01', '2027-02-04', '2027-02-05', '2027-02-08', '2027-02-09', '2027-02-10',
      '2027-03-01', '2027-04-05', '2027-04-06', '2027-04-30', '2027-06-09', '2027-09-15',
      '2027-09-28', '2027-10-11', '2027-10-25', '2027-12-24', '2027-12-31'
    ]
  };
  /** 週六補班日（週六要上班的日期）。2026、2027 皆無 */
  const TW_GOV_MAKEUP_SATURDAYS = [];

  const DAY_MS = 86400000;
  const pad2 = function (n) { return (n < 10 ? '0' : '') + n; };
  const toIso = function (ms) {
    const d = new Date(ms);
    return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
  };
  /** 'YYYY-MM-DD'（必須是真實存在的日期）→ UTC 毫秒；不合法回 NaN。全程用 UTC，與使用者／伺服器所在時區無關 */
  const parseIso = function (s) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(typeof s === 'string' ? s : '');
    if (!m) return NaN;
    const y = +m[1], mo = +m[2], d = +m[3];
    const ms = Date.UTC(y, mo - 1, d);
    const dt = new Date(ms);
    return (dt.getUTCFullYear() === y && dt.getUTCMonth() === mo - 1 && dt.getUTCDate() === d) ? ms : NaN;
  };
  const dowOf = function (ms) { return new Date(ms).getUTCDay(); };

  const MAKEUP_SET = {};
  TW_GOV_MAKEUP_SATURDAYS.forEach(function (d) { MAKEUP_SET[d] = true; });

  /** year 年 8 個固定日期的實際放假日（逢週六 → 前一個週五、逢週日 → 後一個週一），回傳 ISO 陣列（可能落在前／後一年） */
  function fixedObserved(year) {
    return FIXED_HOLIDAYS.map(function (md) {
      let ms = Date.UTC(year, md[0] - 1, md[1]);
      const w = dowOf(ms);
      if (w === 6) ms -= DAY_MS; else if (w === 0) ms += DAY_MS;
      return toIso(ms);
    });
  }

  const _holCache = {};
  /** year 年的平日放假日集合：有逐日表就用表；沒有就用固定日期推算（含前後一年補假落到本年的） */
  function holidaysOfYear(year) {
    if (_holCache[year]) return _holCache[year];
    const set = {};
    const table = TW_GOV_WEEKDAY_HOLIDAYS[year];
    if (table) table.forEach(function (d) { set[d] = true; });
    else [year - 1, year, year + 1].forEach(function (yy) {
      fixedObserved(yy).forEach(function (iso) { if (+iso.slice(0, 4) === year) set[iso] = true; });
    });
    _holCache[year] = set;
    return set;
  }

  /** 這天是不是工作天（週一至週五且非放假日；週六補班日算工作天）。日期不合法回 false */
  function isWorkingDay(isoDate) {
    const ms = parseIso(isoDate);
    if (Number.isNaN(ms)) return false;
    const w = dowOf(ms);
    if (w === 0 || w === 6) return MAKEUP_SET[isoDate] === true;
    return holidaysOfYear(+isoDate.slice(0, 4))[isoDate] !== true;
  }

  /**
   * 「startIso 之後隔天起、到 endIso（含）」之間的工作天數——**不含 startIso 當天**。
   * endIso 不晚於 startIso、或任一日期不合法 → 0。
   */
  function workingDaysAfter(startIso, endIso) {
    const s = parseIso(startIso), e = parseIso(endIso);
    if (Number.isNaN(s) || Number.isNaN(e) || e <= s) return 0;
    let n = 0;
    for (let ms = s + DAY_MS; ms <= e; ms += DAY_MS) if (isWorkingDay(toIso(ms))) n++;
    return n;
  }

  /** year 年 month 月（1~12）的最後一個工作天，回傳 'YYYY-MM-DD'；參數不合法回 '' */
  function lastWorkingDay(year, month) {
    if (!Number.isInteger(year) || !Number.isInteger(month) || year < 1000 || year > 9999 || month < 1 || month > 12) return '';
    let ms = Date.UTC(year, month, 0);                       // 當月最後一天
    for (let i = 0; i < 40; i++, ms -= DAY_MS) { const iso = toIso(ms); if (isWorkingDay(iso)) return iso; }
    return '';
  }

  /**
   * 報價期限預設值（新規則，見檔頭）：dateStr（報價日期 'YYYY-MM-DD'）當月最後一個工作天 L；
   * dateStr 之後（不含當天）到 L 的工作天數不足 opts.minWorkingDays（預設 MIN_VALID_WORKING_DAYS＝7）→ 下個月最後一個工作天。
   * 不論門檻設多少，結果都不會早於 dateStr。dateStr 不合法回傳 ''。
   */
  function defaultValidUntil(dateStr, opts) {
    const ms = parseIso(dateStr);
    if (Number.isNaN(ms)) return '';
    const min = (opts && typeof opts.minWorkingDays === 'number' && Number.isFinite(opts.minWorkingDays)) ? opts.minWorkingDays : MIN_VALID_WORKING_DAYS;
    let y = +dateStr.slice(0, 4), m = +dateStr.slice(5, 7);
    let r = lastWorkingDay(y, m);
    if (r < dateStr || workingDaysAfter(dateStr, r) < min) {
      m += 1; if (m === 13) { m = 1; y += 1; }
      r = lastWorkingDay(y, m);
    }
    return r;
  }

  return {
    MIN_VALID_WORKING_DAYS: MIN_VALID_WORKING_DAYS,
    COVERED_YEARS: Object.keys(TW_GOV_WEEKDAY_HOLIDAYS).map(Number),
    isWorkingDay: isWorkingDay,
    workingDaysAfter: workingDaysAfter,
    lastWorkingDay: lastWorkingDay,
    defaultValidUntil: defaultValidUntil,
    // 供測試與核對資料用（唯讀使用）
    _data: { FIXED_HOLIDAYS: FIXED_HOLIDAYS, TW_GOV_WEEKDAY_HOLIDAYS: TW_GOV_WEEKDAY_HOLIDAYS, TW_GOV_MAKEUP_SATURDAYS: TW_GOV_MAKEUP_SATURDAYS, fixedObserved: fixedObserved }
  };
});
