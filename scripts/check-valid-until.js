#!/usr/bin/env node
/**
 * 報價期限預設值規則檢查。用法：node scripts/check-valid-until.js
 * 動 _client/tw-workdays.js（假日表／門檻）、lib/quoteExcel.js 的 defaultValidUntil／effectiveValidUntil、
 * lib/quoteRoutes.js 的 resolveValidUntil、_client/quote.js 的 quoteLastWorkingDay 之後必跑。
 *
 * 規則（單一來源 _client/tw-workdays.js）：報價日期 D 當月最後工作天 L；D 之後（不含當天）到 L 的工作天 < 7 → 下個月最後工作天。
 *   1) 假日表資料完整性（筆數、全是平日、不重複、與本檔獨立抄錄的一份相同、固定日期推算結果都在表內、無補班日）
 *   2) 暴力比對：2024-01-01～2032-12-31 每一天，新 defaultValidUntil 與「逐日遞增計數」的獨立參考實作完全相同
 *   3) 邊界（手算期望值）：剩餘恰為 6／7／8、月底週末／假日、12→1 月、閏年／平年 2 月、2026／2027 連假附近、門檻參數、補班日支援
 *   4) 舊單不變：effectiveValidUntil 對凍結的舊版（git 歷史上的 lib/quoteExcel.js）逐筆比對 3 萬組隨機資料 0 差異
 *   5) 前後端單一來源：vm 載入 _client/tw-workdays.js（模擬瀏覽器）與伺服器 require 的版本逐日相同；quote.js 的 quoteLastWorkingDay 等於它
 *   6) 接線：index.html 載入順序、server.js 版本注入、欄位說明數字、vercel.json includeFiles 含 _client
 *   7) 路由層（不啟動伺服器，直接呼叫 handler）：POST 空白期限 → 新規則；舊單 PUT 原樣送回 → 維持沒存；真的改日期 → 才存
 *
 * 環境變數：
 *   VU_ROOT        要檢查的專案根目錄（預設＝本檔上一層）。變異測試用，指向被破壞的副本
 *   VU_FROZEN_QE   凍結的舊版 lib/quoteExcel.js 路徑（預設從 git 取 FROZEN_COMMIT 那個 commit 的版本）
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const Module = require('module');
const { execFileSync } = require('child_process');
const ROOT = process.env.VU_ROOT || path.join(__dirname, '..');
// 新規則上線前的最後一個 commit：lib/quoteExcel.js 在這個版本還是舊規則（4 個固定假日、沒有 7 天門檻）
const FROZEN_COMMIT = '5cc1487';

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const note = [];

const TWWD = require(path.join(ROOT, '_client/tw-workdays.js'));
const QE = require(path.join(ROOT, 'lib/quoteExcel.js'));
const read = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8').replace(/\r\n/g, '\n');

// ── 獨立的參考實作（刻意不共用被測程式的任何邏輯，也不 import 它的假日表）────────────────────
const REF_TABLE = {
  2026: new Set(['2026-01-01', '2026-02-16', '2026-02-17', '2026-02-18', '2026-02-19', '2026-02-20', '2026-02-27', '2026-04-03',
    '2026-04-06', '2026-05-01', '2026-06-19', '2026-09-25', '2026-09-28', '2026-10-09', '2026-10-26', '2026-12-25']),
  2027: new Set(['2027-01-01', '2027-02-04', '2027-02-05', '2027-02-08', '2027-02-09', '2027-02-10', '2027-03-01', '2027-04-05',
    '2027-04-06', '2027-04-30', '2027-06-09', '2027-09-15', '2027-09-28', '2027-10-11', '2027-10-25', '2027-12-24', '2027-12-31']),
};
const REF_FIXED_MD = ['01-01', '02-28', '04-04', '05-01', '09-28', '10-10', '10-25', '12-25'];
const DAY = 86400000;
const msOf = (iso) => Date.parse(iso + 'T00:00:00Z');
const isoOf = (ms) => new Date(ms).toISOString().slice(0, 10);
const dowOfIso = (iso) => new Date(msOf(iso)).getUTCDay();
function refIsHoliday(iso) {
  const y = +iso.slice(0, 4);
  if (REF_TABLE[y]) return REF_TABLE[y].has(iso);                      // 有逐日表的年度：表就是全部
  const fixed = (x) => REF_FIXED_MD.includes(x.slice(5));
  const w = dowOfIso(iso);
  if (w >= 1 && w <= 5 && fixed(iso)) return true;                     // 固定日剛好是平日
  if (w === 5 && fixed(isoOf(msOf(iso) + DAY))) return true;           // 隔天（週六）是固定日 → 週五補假
  if (w === 1 && fixed(isoOf(msOf(iso) - DAY))) return true;           // 前一天（週日）是固定日 → 週一補假
  return false;
}
const refIsWork = (iso) => { const w = dowOfIso(iso); return !(w === 0 || w === 6) && !refIsHoliday(iso); };
function refLast(y, m) {                                                // 當月最後一個工作天：從 28 號往後走到月底，再往回找
  let ms = msOf(`${y}-${String(m).padStart(2, '0')}-28`);
  while (new Date(ms + DAY).getUTCMonth() === m - 1) ms += DAY;
  while (!refIsWork(isoOf(ms))) ms -= DAY;
  return isoOf(ms);
}
function refDefault(d, min) {
  min = min === undefined ? 7 : min;
  const y = +d.slice(0, 4), m = +d.slice(5, 7), L = refLast(y, m);
  let remaining = 0;                                                    // D 之後的隔天起到 L（含）逐日數工作天
  for (let ms = msOf(d) + DAY; isoOf(ms) <= L; ms += DAY) if (refIsWork(isoOf(ms))) remaining++;
  if (remaining >= min && L >= d) return L;
  return m === 12 ? refLast(y + 1, 1) : refLast(y, m + 1);
}
const allDates = (from, to) => { const out = []; for (let ms = msOf(from); ms <= msOf(to); ms += DAY) out.push(isoOf(ms)); return out; };

function rng(seed) { let a = seed >>> 0; return () => { a = (a + 0x6D2B79F5) | 0; let x = Math.imul(a ^ (a >>> 15), 1 | a); x = (x + Math.imul(x ^ (x >>> 7), 61 | x)) ^ x; return ((x ^ (x >>> 14)) >>> 0) / 4294967296; }; }

(async () => {
  const D = TWWD._data;
  // ── 1) 假日表資料完整性 ─────────────────────────────────────────────
  t('1a. 逐日表涵蓋 2026、2027；筆數 16／17', JSON.stringify(TWWD.COVERED_YEARS) === '[2026,2027]' && D.TW_GOV_WEEKDAY_HOLIDAYS[2026].length === 16 && D.TW_GOV_WEEKDAY_HOLIDAYS[2027].length === 17,
    JSON.stringify(TWWD.COVERED_YEARS) + ' ' + Object.values(D.TW_GOV_WEEKDAY_HOLIDAYS).map((a) => a.length));
  const badRows = [];
  Object.entries(D.TW_GOV_WEEKDAY_HOLIDAYS).forEach(([y, arr]) => {
    arr.forEach((d, i) => {
      const real = /^\d{4}-\d{2}-\d{2}$/.test(d) && isoOf(msOf(d)) === d;
      if (!real) badRows.push(`${d} 不是真實日期`);
      else if (d.slice(0, 4) !== y) badRows.push(`${d} 不在 ${y} 年`);
      else if ([0, 6].includes(dowOfIso(d))) badRows.push(`${d} 是週末（表內只放週一至週五）`);
      if (i > 0 && !(d > arr[i - 1])) badRows.push(`${d} 未遞增或重複`);
    });
  });
  t('1b. 表內每筆都是真實日期、在該年、週一至週五、遞增不重複', badRows.length === 0, badRows.slice(0, 3).join('；'));
  const sameAsRef = Object.entries(D.TW_GOV_WEEKDAY_HOLIDAYS).every(([y, arr]) => REF_TABLE[y] && arr.length === REF_TABLE[y].size && arr.every((d) => REF_TABLE[y].has(d)));
  t('1c. 表內容與本檔獨立抄錄的一份（人事行政總處官方 CSV 驗證結果）完全相同', sameAsRef);
  const missFixed = [];
  [2025, 2026, 2027, 2028].forEach((y) => D.fixedObserved(y).forEach((d) => {
    const tb = D.TW_GOV_WEEKDAY_HOLIDAYS[+d.slice(0, 4)];                // 落在有逐日表的年度才檢查（2028 年的固定日只取跨年補到 2027-12-31 的）
    if (tb && !tb.includes(d)) missFixed.push(d);
  }));
  t('1d. 8 個固定日期（含補假、含 2028-01-01 補在 2027-12-31）推算結果全在逐日表內', missFixed.length === 0, missFixed.join(','));
  t('1e. 無週六補班日（2026、2027 官方都沒有）；若日後有，必須是週六', D.TW_GOV_MAKEUP_SATURDAYS.every((d) => dowOfIso(d) === 6));
  t('1f. 門檻常數＝7', TWWD.MIN_VALID_WORKING_DAYS === 7);

  // ── 2) 暴力比對 2024-01-01 ~ 2032-12-31 ─────────────────────────────
  const dates = allDates('2024-01-01', '2032-12-31');
  const diffs = []; let bumped = 0, diffLegacy = 0;
  dates.forEach((d) => {
    const got = TWWD.defaultValidUntil(d), want = refDefault(d);
    if (got !== want) diffs.push(`${d}: ${got} vs ${want}`);
    if (want.slice(0, 7) !== d.slice(0, 7)) bumped++;
    if (QE.legacyDefaultValidUntil(d) !== want) diffLegacy++;
  });
  t(`2a. 暴力比對 ${dates.length} 個日期（2024~2032）新 defaultValidUntil 與參考實作 0 差異`, dates.length === 9 * 365 + 3 && diffs.length === 0, `${dates.length} 天；差異 ${diffs.length}：${diffs.slice(0, 3).join('；')}`);
  note.push(`2a 暴力比對：${dates.length} 個日期，差異 ${diffs.length}；其中順延到下個月 ${bumped} 天；與舊規則結果不同 ${diffLegacy} 天`);
  t('2b. 比對有鑑別力：順延與不順延的日期都很多，且新舊規則確有差異', bumped > 500 && bumped < dates.length - 500 && diffLegacy > 300, `順延 ${bumped}／不同於舊規則 ${diffLegacy}`);
  const ladd = [];
  for (const d of dates) { if (QE.defaultValidUntil(d) !== TWWD.defaultValidUntil(d)) ladd.push(d); }
  t('2c. 伺服器 QE.defaultValidUntil 與 TWWD.defaultValidUntil 同一個函式（單一來源）', QE.defaultValidUntil === TWWD.defaultValidUntil && ladd.length === 0);
  const minDiffs = [];
  [0, 1, 3, 5, 7, 10, 15].forEach((min) => dates.filter((_, i) => i % 7 === 0).forEach((d) => { if (TWWD.defaultValidUntil(d, { minWorkingDays: min }) !== refDefault(d, min)) minDiffs.push(`${d}@${min}`); }));
  t('2d. 門檻參數 minWorkingDays（0/1/3/5/7/10/15）對 470 個抽樣日期與參考實作相同（含 0：結果仍不早於報價日期）', minDiffs.length === 0, minDiffs.slice(0, 3).join(','));
  const lwdDiffs = [];
  for (let y = 2024; y <= 2032; y++) for (let m = 1; m <= 12; m++) if (TWWD.lastWorkingDay(y, m) !== refLast(y, m)) lwdDiffs.push(`${y}-${m}`);
  t('2e. lastWorkingDay 108 個月（2024~2032）與參考實作相同', lwdDiffs.length === 0, lwdDiffs.slice(0, 3).join(','));
  const lwdVsLegacy = [];
  for (let y = 2000; y <= 2060; y++) for (let m = 1; m <= 12; m++) if (TWWD.lastWorkingDay(y, m) !== QE.legacyLastWorkingDay(y, m)) lwdVsLegacy.push(`${y}-${m}`);
  note.push(`2f 月底工作天新舊規則對照（2000~2060 共 ${61 * 12} 個月）：不同 ${lwdVsLegacy.length} 個${lwdVsLegacy.length ? '（' + lwdVsLegacy.slice(0, 5).join(',') + '…）' : ''}`);

  // ── 3) 邊界（期望值全是手算；Oct/Nov/Feb/Sep 2026、Dec 2027 的 6/7/8 剛好跨在連假兩側）────────────────
  const EDGE = [
    // [報價日期, 期望預設期限, 說明]
    ['2026-10-19', '2026-10-30', '10 月：剩餘 8（10/26 光復節補假不算）→ 當月'],
    ['2026-10-20', '2026-10-30', '10 月：剩餘 7 → 當月（恰好夠）'],
    ['2026-10-21', '2026-11-30', '10 月：剩餘 6 → 順延（沒有 10/26 補假的話會是 7，這格會抓到漏掉補假）'],
    ['2026-10-26', '2026-11-30', '10/26 補假當天建單：剩餘 4 → 順延'],
    ['2026-10-30', '2026-11-30', '報價日期＝當月最後工作天當天 → 剩餘 0 → 順延'],
    ['2026-10-31', '2026-11-30', '月底週六 → 順延'],
    ['2026-11-18', '2026-11-30', '11 月：剩餘 8'],
    ['2026-11-19', '2026-11-30', '11 月：剩餘 7'],
    ['2026-11-20', '2026-12-31', '11 月：剩餘 6 → 12/31'],
    ['2026-09-16', '2026-09-30', '9 月：剩餘 8（中秋 9/25、教師節 9/28 都放假）'],
    ['2026-09-17', '2026-09-30', '9 月：剩餘 7'],
    ['2026-09-18', '2026-10-30', '9 月：剩餘 6 → 順延'],
    ['2026-02-09', '2026-02-26', '2 月：春節後剩餘 8（2/16~2/20 春節＋補假、2/27 補假都放假；最後工作天是 2/26）'],
    ['2026-02-10', '2026-02-26', '2 月：剩餘 7'],
    ['2026-02-11', '2026-03-31', '2 月：剩餘 6 → 順延（沒內建春節的話會是 11，這格會抓到漏掉春節）'],
    ['2026-02-27', '2026-03-31', '2/27 和平紀念日補假當天（晚於當月最後工作天 2/26）→ 順延'],
    ['2026-02-28', '2026-03-31', '2/28 週六 → 順延'],
    ['2026-05-31', '2026-06-30', '5 月最後工作天是 5/29（週五）；5/31 週日 → 6/30'],
    ['2026-06-18', '2026-06-30', '6 月：端午 6/19 放假，剩餘 7'],
    ['2026-06-22', '2026-07-31', '6 月：剩餘 6 → 順延'],
    ['2026-12-25', '2027-01-29', '12 月→隔年 1 月：行憲紀念日當天，剩餘 4 → 1/29（1/31 是週日）'],
    ['2026-12-31', '2027-01-29', '12/31 是 12 月最後工作天 → 順延到隔年 1/29'],
    ['2027-01-31', '2027-02-26', '1/31 週日（晚於 1/29）→ 2 月最後工作天 2/26（2/28 週日、補假落在 3/1）'],
    ['2027-12-17', '2027-12-30', '2027-12：12/24、12/31 放假，最後工作天 12/30；剩餘 8'],
    ['2027-12-20', '2027-12-30', '2027-12：剩餘 7'],
    ['2027-12-21', '2028-01-31', '2027-12：剩餘 6 → 順延到隔年 1/31（沒內建 12/24 的話會是 7，這格會抓到漏掉）'],
    ['2027-12-30', '2028-01-31', '2027-12-30 是最後工作天當天 → 順延'],
    ['2027-12-31', '2028-01-31', '2027-12-31（2028 元旦補假）→ 順延'],
    ['2028-02-16', '2028-02-29', '閏年 2 月：2/28 週一和平紀念日放假、最後工作天是 2/29；剩餘 8'],
    ['2028-02-17', '2028-02-29', '閏年 2 月：剩餘 7'],
    ['2028-02-18', '2028-03-31', '閏年 2 月：剩餘 6 → 順延'],
    ['2028-02-28', '2028-03-31', '2/28 和平紀念日當天，剩餘 1 → 順延'],
    ['2028-02-29', '2028-03-31', '閏日＝最後工作天當天 → 順延'],
    ['2029-02-15', '2029-02-27', '平年 2 月：2/28 週三放假、最後工作天 2/27；剩餘 8'],
    ['2029-02-16', '2029-02-27', '平年 2 月：剩餘 7'],
    ['2029-02-19', '2029-03-30', '平年 2 月：剩餘 6 → 順延；3/31 週六 → 3/30'],
    ['2029-02-28', '2029-03-30', '2/28 放假當天（晚於 2/27）→ 順延'],
    ['2028-12-31', '2029-01-31', '12 月→隔年 1 月（2028-12-31 週日，12 月最後工作天是 12/29）'],
    ['2031-12-31', '2032-01-30', '2031-12-31 週三＝最後工作天當天；2032-01-31 週六 → 1/30'],
  ];
  const edgeBad = EDGE.filter(([d, want]) => TWWD.defaultValidUntil(d) !== want || QE.defaultValidUntil(d) !== want).map(([d, want]) => `${d} 得 ${TWWD.defaultValidUntil(d)} 應為 ${want}`);
  t(`3a. 邊界 ${EDGE.length} 例（手算期望值；剩餘恰為 6／7／8、月底週末與假日、12→1 月、閏年／平年 2 月、2026／2027 連假附近）`, edgeBad.length === 0, edgeBad.slice(0, 3).join('；'));
  // 剩餘工作天數本身
  const wd = [
    [['2026-10-08', '2026-10-12'], 1, '10/9 補假、週末，下一個工作天 10/12'],
    [['2026-10-12', '2026-10-08'], 0, '結束早於開始'], [['2026-10-08', '2026-10-08'], 0, '同一天（不含當天）'],
    [['2026-10-20', '2026-10-30'], 7, '10/21~10/30 扣 10/26'], [['2026-02-10', '2026-02-26'], 7, '2/11~2/26 扣春節'],
    [['abc', '2026-10-30'], 0, '不合法'], [['2026-02-31', '2026-03-31'], 0, '不存在的日期'], [[undefined, undefined], 0, 'undefined'],
    [['2026-12-31', '2027-01-04'], 1, '1/1 元旦、週末，只剩 1/4（週一）'],
  ];
  const wdBad = wd.filter(([a, want]) => TWWD.workingDaysAfter(a[0], a[1]) !== want).map(([a, want, m]) => `${m}: ${TWWD.workingDaysAfter(a[0], a[1])}≠${want}`);
  t(`3b. workingDaysAfter（不含起始當天）${wd.length} 例`, wdBad.length === 0, wdBad.join('；'));
  const iw = [['2026-02-16', false], ['2026-02-23', true], ['2026-10-10', false], ['2026-10-09', false], ['2027-12-31', false], ['2027-12-30', true], ['2028-01-03', true],
    ['2025-12-31', true], ['2026-02-31', false], ['abc', false], ['', false], [null, false], ['2025-10-10', false], ['2025-10-13', true], ['2030-10-10', false], ['2030-10-11', true]];
  const iwBad = iw.filter(([d, want]) => TWWD.isWorkingDay(d) !== want).map(([d, want]) => `${d}≠${want}`);
  t(`3c. isWorkingDay ${iw.length} 例（連假、週末、不合法輸入、未收錄年度只用固定日）`, iwBad.length === 0, iwBad.join('；'));
  t('3d. 門檻參數：10/21 剩餘 6 → 門檻 6 留在當月、門檻 7 順延；門檻 0 時 10/31（晚於月底）仍不早於報價日期',
    TWWD.defaultValidUntil('2026-10-21', { minWorkingDays: 6 }) === '2026-10-30' && TWWD.defaultValidUntil('2026-10-21', { minWorkingDays: 7 }) === '2026-11-30'
    && TWWD.defaultValidUntil('2026-10-31', { minWorkingDays: 0 }) === '2026-11-30' && TWWD.defaultValidUntil('2026-10-21', {}) === '2026-11-30' && TWWD.defaultValidUntil('2026-10-21', { minWorkingDays: 'x' }) === '2026-11-30');
  t('3e. 不合法輸入回 \'\'（不丟例外）', ['', 'abc', '2026-02-31', '2026-1-5', null, undefined, 20261001, ['2026-10-01'], '2026-10-01x', ' 2026-10-01'].every((x) => TWWD.defaultValidUntil(x) === ''));
  // 補班日支援：用 vm 載入「補班日＝2026-10-31（週六）」的副本。10/31 變成當月最後工作天 → 10/21 剩餘 7（22,23,27,28,29,30,31）→ 留在當月 10/31
  const srcTw = fs.readFileSync(path.join(ROOT, '_client/tw-workdays.js'), 'utf8');
  const MK_LINE = 'const TW_GOV_MAKEUP_SATURDAYS = [];';
  if (!srcTw.includes(MK_LINE)) t('3f. 補班日支援（找不到補班日宣告行，無法建立測試副本）', false);
  else {
    const ctx = { }; vm.createContext(ctx); vm.runInContext(srcTw.replace(MK_LINE, "const TW_GOV_MAKEUP_SATURDAYS = ['2026-10-31'];"), ctx);
    const M = ctx.TWWD;
    t('3f. 補班日（週六上班）納入工作天：補班日＝10/31 → isWorkingDay 為真、最後工作天變 10/31、10/21 剩餘 7 留在當月', M.isWorkingDay('2026-10-31') === true && M.isWorkingDay('2026-11-07') === false
      && M.lastWorkingDay(2026, 10) === '2026-10-31' && M.defaultValidUntil('2026-10-21') === '2026-10-31' && M.defaultValidUntil('2026-10-22') === '2026-11-30');
  }

  // ── 4) 舊單不變：對凍結的舊版逐筆比對 ────────────────────────────────
  let frozenSrc, frozenFrom;
  try {
    if (process.env.VU_FROZEN_QE) { frozenSrc = fs.readFileSync(process.env.VU_FROZEN_QE, 'utf8'); frozenFrom = process.env.VU_FROZEN_QE; }
    else { frozenSrc = execFileSync('git', ['show', `${FROZEN_COMMIT}:lib/quoteExcel.js`], { cwd: path.join(__dirname, '..'), encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'], maxBuffer: 20 * 1024 * 1024 }); frozenFrom = `git ${FROZEN_COMMIT}`; }
  } catch (e) { frozenSrc = null; }
  if (!frozenSrc || !frozenSrc.includes('function defaultValidUntil(dateStr)')) {
    t('4. 取得凍結舊版（git show ' + FROZEN_COMMIT + ':lib/quoteExcel.js；或設 VU_FROZEN_QE）', false, '取不到舊版或內容不是預期的舊規則，無法證明舊單不變');
  } else {
    // 把舊版放在虛擬路徑 lib/__frozen__.js 載入，讓它的 require('./quoteRemarks') 等照常解析到目前的 lib
    const vp = path.join(ROOT, 'lib', '__frozen_quoteExcel__.js');
    const mod = new Module(vp, module); mod.filename = vp; mod.paths = Module._nodeModulePaths(path.dirname(vp)); mod._compile(frozenSrc, vp);
    const F = mod.exports;
    t('4a. 凍結版確實是舊規則（沒有 TWWD、defaultValidUntil 對 2026-10-30 回當天）', !frozenSrc.includes('tw-workdays') && F.defaultValidUntil('2026-10-30') === '2026-10-30' && QE.defaultValidUntil('2026-10-30') === '2026-11-30', frozenFrom);
    const rnd = rng(20261008);
    const pick = (a) => a[Math.floor(rnd() * a.length)];
    const randDate = () => isoOf(Date.UTC(2018, 0, 1) + Math.floor(rnd() * 23 * 365) * DAY);
    const randCreated = () => {
      const r = rnd();
      if (r < 0.62) return new Date(Date.UTC(2018, 0, 1) + Math.floor(rnd() * 23 * 365 * DAY)).toISOString();
      if (r < 0.72) return randDate();                                       // 只有日期
      if (r < 0.78) return `${randDate()}T${pick(['15:59:59', '16:00:00', '23:30:00', '00:00:01'])}+08:00`;   // 台灣時區邊界
      return pick(['', null, undefined, 'garbage', '2026-13-45', 12345, 1.79e12, {}, [], '0000-00-00']);
    };
    const randValid = () => {
      const r = rnd();
      if (r < 0.3) return randDate();
      if (r < 0.6) return undefined;
      return pick(['', null, '2026-02-31', 'abc', '2026-10-30xyz', 20261030, ['2026-10-30'], ' 2026-10-30', '2026/10/30', true]);
    };
    const N = 30000; const bad = []; let storedHit = 0, legacyUsed = 0, differsFromNew = 0;
    for (let i = 0; i < N; i++) {
      const q = {}; if (rnd() < 0.9) q.createdAt = randCreated(); const v = randValid(); if (v !== undefined || rnd() < 0.2) q.validUntil = v;
      if (rnd() < 0.7) q.quoteDate = rnd() < 0.85 ? randDate() : pick(['', 'abc', '2026-02-31', null]);
      const a = JSON.stringify(F.effectiveValidUntil(q)), b = JSON.stringify(QE.effectiveValidUntil(q));
      if (a !== b) bad.push(`${JSON.stringify(q)} → 舊 ${a} 新 ${b}`);
      if (QE.isRealIsoDate(q.validUntil)) storedHit++;
      else {
        legacyUsed++;
        let baseDate = ''; if (q.createdAt) { const tt = new Date(q.createdAt); if (!Number.isNaN(tt.getTime())) baseDate = QE.taipeiToday(tt); }
        if (baseDate && TWWD.defaultValidUntil(baseDate) !== QE.effectiveValidUntil(q)) differsFromNew++;   // 用新規則算同一個建立日期會不同的筆數（鑑別力指標）
      }
    }
    t(`4b. effectiveValidUntil：${N} 組隨機資料（createdAt 含台灣時區邊界／非法值，validUntil 有值／無值／非法）與凍結舊版 0 差異`, bad.length === 0, `${N} 組；差異 ${bad.length}：${bad.slice(0, 2).join('；')}`);
    t('4c. 這批資料有鑑別力：走到「推算」分支的很多，且其中有不少新舊規則會給不同答案（新規則若誤用於舊單會被抓到）', legacyUsed > 10000 && differsFromNew > 500, `推算 ${legacyUsed} 組、新舊規則不同 ${differsFromNew} 組、直接用已存值 ${storedHit} 組`);
    note.push(`4b 舊單比對：${N} 組隨機資料，差異 ${bad.length}（凍結版來源：${frozenFrom}）；走推算分支 ${legacyUsed} 組，其中若誤用新規則會不同的 ${differsFromNew} 組`);
    const dd = allDates('1990-01-01', '2060-12-31'); const dBad = [];
    dd.forEach((d) => { if (F.defaultValidUntil(d) !== QE.legacyDefaultValidUntil(d)) dBad.push(d); });
    ['', 'abc', '2026-02-31', null, undefined, 20261001, '2026-1-5'].forEach((x) => { if (F.defaultValidUntil(x) !== QE.legacyDefaultValidUntil(x)) dBad.push(String(x)); });
    t(`4d. legacyDefaultValidUntil 對凍結版的 defaultValidUntil：${dd.length} 個日期（1990~2060）＋ 7 個不合法輸入 0 差異`, dBad.length === 0, dBad.slice(0, 3).join(','));
    const lBad = [];
    for (let y = 1990; y <= 2060; y++) for (let m = 1; m <= 12; m++) if (F.lastWorkingDay(y, m) !== QE.legacyLastWorkingDay(y, m) || F.lastWorkingDay(y, m) !== QE.lastWorkingDay(y, m)) lBad.push(`${y}-${m}`);
    t('4e. legacyLastWorkingDay／匯出的 lastWorkingDay 對凍結版 852 個月 0 差異', lBad.length === 0, lBad.slice(0, 3).join(','));
    t('4f. 具體例子：2026-10-26 建立、沒存期限的舊單 → 仍是 2026-10-30（新規則對 10/26 會給 2026-11-30）', QE.effectiveValidUntil({ createdAt: '2026-10-26T02:00:00.000Z' }) === '2026-10-30' && TWWD.defaultValidUntil('2026-10-26') === '2026-11-30');
  }

  // ── 5) 前後端單一來源 ──────────────────────────────────────────────
  const browserCtx = {}; vm.createContext(browserCtx);                         // 沒有 window／document／module，只有 globalThis：模擬瀏覽器全域
  vm.runInContext(srcTw, browserCtx);
  const B = browserCtx.TWWD;
  t('5a. _client/tw-workdays.js 在沒有 module 的全域（瀏覽器）載入 → 掛上 TWWD，且不需要 DOM', !!B && typeof B.defaultValidUntil === 'function' && typeof B.isWorkingDay === 'function' && B.MIN_VALID_WORKING_DAYS === 7);
  const sBad = [];
  dates.forEach((d, i) => {
    if (B.defaultValidUntil(d) !== TWWD.defaultValidUntil(d)) sBad.push('default ' + d);
    if (B.isWorkingDay(d) !== TWWD.isWorkingDay(d)) sBad.push('work ' + d);
    const e = dates[Math.min(dates.length - 1, i + 9)];
    if (B.workingDaysAfter(d, e) !== TWWD.workingDaysAfter(d, e)) sBad.push('after ' + d);
  });
  for (let y = 2024; y <= 2032; y++) for (let m = 1; m <= 12; m++) if (B.lastWorkingDay(y, m) !== TWWD.lastWorkingDay(y, m)) sBad.push(`lwd ${y}-${m}`);
  t(`5b. 瀏覽器版（vm）與伺服器 require 版：${dates.length} 天的 defaultValidUntil／isWorkingDay／workingDaysAfter 與 108 個月 lastWorkingDay 0 差異`, sBad.length === 0, sBad.slice(0, 3).join(','));
  t('5c. 兩邊的假日資料相同（逐日表、固定日、補班日）', JSON.stringify([B._data.TW_GOV_WEEKDAY_HOLIDAYS, B._data.FIXED_HOLIDAYS, B._data.TW_GOV_MAKEUP_SATURDAYS]) === JSON.stringify([D.TW_GOV_WEEKDAY_HOLIDAYS, D.FIXED_HOLIDAYS, D.TW_GOV_MAKEUP_SATURDAYS]));
  const qj = read('_client/quote.js');
  const f0 = qj.indexOf('function quoteLastWorkingDay(');
  let f1 = -1;
  if (f0 >= 0) { let depth = 0, started = false; for (let i = qj.indexOf('{', f0); i < qj.length; i++) { if (qj[i] === '{') { depth++; started = true; } else if (qj[i] === '}') { depth--; } if (started && depth === 0) { f1 = i + 1; break; } } }
  if (f0 < 0 || f1 < 0) t('5d. 找不到 quote.js 的 quoteLastWorkingDay', false);
  else {
    const fnSrc = qj.slice(f0, f1);
    const withTw = {}; vm.createContext(withTw); vm.runInContext(srcTw, withTw); vm.runInContext(fnSrc, withTw);
    const noTw = {}; vm.createContext(noTw); vm.runInContext(fnSrc, noTw);
    const qBad = [];
    dates.forEach((d) => { if (withTw.quoteLastWorkingDay(d) !== TWWD.defaultValidUntil(d)) qBad.push(d); });
    const weird = ['', 'abc', '2026-02-31', '2026-1-5', null, undefined, 20261001];
    weird.forEach((x) => { if (withTw.quoteLastWorkingDay(x) !== '') qBad.push(String(x)); });
    t(`5d. quote.js 的 quoteLastWorkingDay：${dates.length} 個日期＋ ${weird.length} 個不合法輸入都等於 TWWD.defaultValidUntil`, qBad.length === 0, qBad.slice(0, 3).join(','));
    t('5e. 沒載入 TWWD 時 quoteLastWorkingDay 回 \'\'（不丟例外，業務自己輸入）', noTw.quoteLastWorkingDay('2026-10-08') === '' && noTw.quoteLastWorkingDay('') === '');
    t('5f. quote.js 裡不再有第二份假日規則（函式本身沒有 Date.UTC／固定日陣列，全檔沒有 [[1, 1], [2, 28] 之類的表）', !/Date\.UTC|\[\[1, 1\]/.test(fnSrc) && !/\[\[1, 1\], \[2, 28\]/.test(qj) && /TWWD\.defaultValidUntil/.test(fnSrc));
  }
  const qe = read('lib/quoteExcel.js');
  t('5g. lib/quoteExcel.js 用 require 取同一支（不是複製一份），舊單走 legacyDefaultValidUntil', /require\('\.\.\/_client\/tw-workdays\.js'\)/.test(qe) && /const defaultValidUntil = TWWD\.defaultValidUntil;/.test(qe) && /return legacyDefaultValidUntil\(base \|\| q\.quoteDate \|\| taipeiToday\(\)\);/.test(qe));
  const qr = read('lib/quoteRoutes.js');
  t('5h. quoteRoutes 空白期限用新 defaultValidUntil；舊單比對仍用 effectiveValidUntil', /const v = s \|\| quoteExcel\.defaultValidUntil\(base\);/.test(qr) && /vuRes\.value === quoteExcel\.effectiveValidUntil\(draft\)/.test(qr));

  // 5i~5l 表單行為：用假 DOM 跑 quote.js 真正的 bindQuoteValidUntil（沒有 jsdom 可用，元素只實作用到的 value／textContent／addEventListener）
  {
    const winCtx = { window: {} }; vm.createContext(winCtx); vm.runInContext(srcTw, winCtx);
    t('5i. 有 window 的環境（瀏覽器）→ 掛在 window.TWWD', !!winCtx.window.TWWD && typeof winCtx.window.TWWD.defaultValidUntil === 'function');
    const b0 = qj.indexOf('let _qValidUntilTouched = false;'), b1s = qj.indexOf('function bindQuoteValidUntil()');
    let b1 = -1;
    if (b1s >= 0) { let depth = 0, started = false; for (let i = qj.indexOf('{', b1s); i < qj.length; i++) { if (qj[i] === '{') { depth++; started = true; } else if (qj[i] === '}') { depth--; } if (started && depth === 0) { b1 = i + 1; break; } } }
    if (b0 < 0 || b1 < 0) t('5j. 找不到 bindQuoteValidUntil', false);
    else {
      const mkEl = (v) => { const el = { value: v || '', textContent: '7', _l: {}, addEventListener(type, fn) { (this._l[type] = this._l[type] || []).push(fn); }, fire(type) { (this._l[type] || []).forEach((fn) => fn.call(this, {})); } }; return el; };
      const runForm = (twSrc) => {
        const els = { qValidUntil: mkEl(''), qDate: mkEl(''), qValidUntilMinDays: mkEl('') };
        const ctx = { $: (id) => els[id] || null, renderCalls: 0, renderQuoteFixedClauses() { ctx.renderCalls++; } };
        vm.createContext(ctx); if (twSrc) vm.runInContext(twSrc, ctx); vm.runInContext(qj.slice(b0, b1) + '\n' + qj.slice(f0, f1), ctx);   // quoteLastWorkingDay 與 bind 一起載入，跟瀏覽器一樣是同一個全域
        return { els, ctx };
      };
      const A = runForm(srcTw);
      vm.runInContext('bindQuoteValidUntil()', A.ctx);
      A.els.qDate.value = '2026-10-21'; A.els.qDate.fire('change');
      const afterFollow = A.els.qValidUntil.value;
      A.els.qDate.value = '2026-10-19'; A.els.qDate.fire('change');
      const afterFollow2 = A.els.qValidUntil.value;
      A.els.qValidUntil.value = '2026-10-23'; A.els.qValidUntil.fire('input');           // 業務手動改過期限
      A.els.qDate.value = '2026-10-21'; A.els.qDate.fire('change');
      t('5j. 報價日期改動 → 沒手動改過的期限自動跟隨新規則（10/21→11-30、10/19→10-30）；手動改過就不再跟隨', afterFollow === '2026-11-30' && afterFollow2 === '2026-10-30' && A.els.qValidUntil.value === '2026-10-23', `${afterFollow} ${afterFollow2} ${A.els.qValidUntil.value}`);
      t('5k. 欄位說明的天數由 TWWD.MIN_VALID_WORKING_DAYS 帶入（常數改成 5 → 說明顯示 5；沒載入 TWWD → 不動、不丟例外）',
        A.els.qValidUntilMinDays.textContent === '7' && (() => { const B2 = runForm(srcTw.replace('const MIN_VALID_WORKING_DAYS = 7;', 'const MIN_VALID_WORKING_DAYS = 5;')); vm.runInContext('bindQuoteValidUntil()', B2.ctx); return B2.els.qValidUntilMinDays.textContent === '5'; })()
        && (() => { const C2 = runForm(null); vm.runInContext('bindQuoteValidUntil()', C2.ctx); C2.els.qDate.value = '2026-10-21'; C2.els.qDate.fire('change'); return C2.els.qValidUntilMinDays.textContent === '7' && C2.els.qValidUntil.value === ''; })());
    }
    t('5l. 開窗時：新單用規則預設、編輯舊單用伺服器給的 q.validUntil（舊單推算值＝舊規則）', /\$\('qValidUntil'\)\.value\s*=\s*q \? \(q\.validUntil \|\| quoteLastWorkingDay\(\$\('qDate'\)\.value\)\) : quoteLastWorkingDay\(\$\('qDate'\)\.value\);/.test(qj) && /_qValidUntilTouched\s*=\s*!!q;/.test(qj));
  }

  // ── 6) 接線 ────────────────────────────────────────────────────────
  const html = read('_client/index.html');
  const iTw = html.indexOf('<script src="tw-workdays.js"></script>'), iQ = html.indexOf('<script src="quote.js"></script>');
  t('6a. index.html 載入 tw-workdays.js 一次，且在 quote.js 之前', iTw > 0 && iQ > 0 && iTw < iQ && html.split('tw-workdays.js').length === 2);
  const sv = read('server.js');
  t('6b. server.js 對 tw-workdays.js 注入 ?v= 版本（比照 quote-costlines.js）', /\.replace\(\/src="tw-workdays\\\.js"\/g,\s*`src="tw-workdays\.js\?v=\$\{BUILD_VERSION\}"`\)/.test(sv));
  const hint = /<span id="qValidUntilMinDays">(\d+)<\/span>/.exec(html);
  t('6c. 欄位說明的後備數字（HTML 內）＝規則常數；說明文字講到「不含當天」與順延', !!hint && +hint[1] === TWWD.MIN_VALID_WORKING_DAYS && /id="qValidUntilHint"[^>]*>[^<]*不含當天[^<]*<span id="qValidUntilMinDays">/.test(html) && /順延到下個月的最後一個工作天/.test(html));
  t('6d. quote.js 開窗時把說明裡的數字改成 TWWD.MIN_VALID_WORKING_DAYS', /qValidUntilMinDays/.test(qj) && /TWWD\.MIN_VALID_WORKING_DAYS/.test(qj));
  let incl = '';
  try { incl = JSON.parse(fs.readFileSync(path.join(ROOT, 'vercel.json'), 'utf8')).functions['api/index.js'].includeFiles; } catch (e) { /* 下面會 fail */ }
  const inclDirs = ((/\{([^}]*)\}\/\*\*/.exec(incl) || [])[1] || '').split(',').map((s) => s.trim());
  t('6e. vercel.json includeFiles 含 _client/**（伺服器 require 的 _client/tw-workdays.js 部署後才會在函式包內）', inclDirs.includes('_client'), incl);

  // ── 7) 路由層（不啟動伺服器：用假的 app 收 handler，直接呼叫）─────────────
  try {
    const store = { quotations: [], quoteApproval: {} };
    const users = [{ username: 'sales1', role: 'user', active: true, displayName: 'S1' }];
    const routes = []; const app = {};
    ['get', 'post', 'put', 'delete', 'patch', 'use'].forEach((m) => { app[m] = (...a) => { routes.push({ method: m, path: a[0], handlers: a.slice(1).flat() }); }; });
    const pass = (req, res, next) => next();
    let seq = 0;
    const deps = {
      db: { load: () => store, save: () => {}, flush: async () => {} }, loadAuth: () => ({ users }), saveAuth: () => {}, requireAuth: pass, requireAdmin: pass,
      writeLog: () => {}, pushNotification: () => {}, getViewableOwners: () => ['sales1'],
      sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 500), genQuoteNo: () => 'QU-VU-' + (++seq),
      taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}), buildQuoteWorkbook: null, buildQuotePnlExcel: null, QUOTE_TEMPLATE: '',
      uuidv4: () => 'id-' + (++seq), normalizeBu: (b) => (Array.isArray(b) ? b : []), getUserFeatures: undefined,
    };
    require(path.join(ROOT, 'lib/quoteRoutes.js'))(app, deps);
    const call = async (method, p, body, params) => {
      const r = routes.find((x) => x.method === method && x.path === p);
      if (!r) throw new Error('找不到路由 ' + method + ' ' + p);
      const req = { body: body || {}, params: params || {}, query: {}, headers: {}, ip: '127.0.0.1', session: { user: { username: 'sales1', role: 'user' } } };
      let status = 200, payload;
      const rs = { status(c) { status = c; return rs; }, json(j) { payload = j; return rs; }, setHeader() { return rs; }, send(j) { payload = j; return rs; } };
      let i = 0; const next = async () => { const h = r.handlers[i++]; if (h) await h(req, rs, next); };
      await next();
      return { s: status, j: payload };
    };
    const base = { company: '路由測試公司', projectName: 'P', products: [], items: [{ desc: 'x', unit: '式', qty: 1, unitPrice: 1000, cost: 0 }], discountType: 'none', discountValue: 0 };
    const POST = (extra) => call('post', '/api/quotations', Object.assign({}, base, extra));
    const cases = [
      [{ quoteDate: '2026-10-21' }, '2026-11-30', '空白（未提供）：10/21 剩餘 6 → 順延'], [{ quoteDate: '2026-10-20', validUntil: '' }, '2026-10-30', '空字串：剩餘 7 → 當月'],
      [{ quoteDate: '2026-10-30', validUntil: null }, '2026-11-30', 'null：最後工作天當天 → 順延'], [{ quoteDate: '2026-10-31' }, '2026-11-30', '月底週六 → 順延'],
      [{ quoteDate: '2027-12-21' }, '2028-01-31', '跨年：2027-12-21 → 2028-01-31'], [{ quoteDate: '2026-02-11' }, '2026-03-31', '春節後：2/11 剩餘 6 → 3/31'],
      [{ quoteDate: '2026-10-21', validUntil: '2026-10-23' }, '2026-10-23', '業務自己輸入的值不受規則影響'], [{ quoteDate: '2026-10-21', validUntil: '2026-10-21' }, '2026-10-21', '期限＝報價日期當天仍可'],
    ];
    const rBad = [];
    for (const [extra, want, label] of cases) { const r = await POST(extra); if (!(r.s === 201 && r.j.validUntil === want)) rBad.push(`${label}: HTTP ${r.s} ${r.j && (r.j.validUntil || r.j.error)} 應為 ${want}`); }
    t(`7a. POST /api/quotations ${cases.length} 例：空白期限 → 新規則預設；明確輸入的值照存`, rBad.length === 0, rBad.slice(0, 2).join('；'));
    const bad2 = await POST({ quoteDate: '2026-10-21', validUntil: '2026-10-20' });
    t('7b. 期限早於報價日期仍 400 BAD_VALID_UNTIL（驗證沒被預設值邏輯影響）', bad2.s === 400 && bad2.j.code === 'BAD_VALID_UNTIL');
    // 已存期限的單：PUT 清空 → 依該單報價日期套新規則
    const mk = await POST({ quoteDate: '2026-10-20', validUntil: '2026-10-23' });
    const put1 = await call('put', '/api/quotations/:id', { validUntil: '' }, { id: mk.j.id });
    t('7c. 已存期限的單 PUT 清空 → 回到新規則預設（10/20 剩餘 7 → 10/30）', put1.s === 200 && put1.j.validUntil === '2026-10-30', `${put1.s} ${put1.j && put1.j.validUntil}`);
    const put2 = await call('put', '/api/quotations/:id', { quoteDate: '2026-10-21', validUntil: '' }, { id: mk.j.id });
    t('7d. PUT 同時改報價日期並清空期限 → 依新日期套新規則（10/21 剩餘 6 → 11/30）', put2.s === 200 && put2.j.validUntil === '2026-11-30', `${put2.s} ${put2.j && put2.j.validUntil}`);
    // 舊單：沒存 validUntil，createdAt 在新舊規則會給不同答案的日子
    store.quotations.push({ id: 'legacy-1', owner: 'sales1', quoteNo: 'QU-OLD-1', company: '舊單公司', quoteDate: '2026-10-26', createdAt: '2026-10-26T02:00:00.000Z', updatedAt: '2026-10-26T02:00:00.000Z',
      products: [], items: [{ lid: 'l1', desc: 'x', unit: '式', qty: 1, unitPrice: 1000, cost: 0 }], status: 'draft', approval: null, discountType: 'none', discountValue: 0, costFlow: { note: '' } });
    const g = await call('get', '/api/quotations/:id', null, { id: 'legacy-1' });
    t('7e. 舊單（沒存期限）GET → 回舊規則推算值 2026-10-30（新規則會是 11-30）', g.s === 200 && g.j.validUntil === '2026-10-30', `${g.s} ${g.j && g.j.validUntil}`);
    const p3 = await call('put', '/api/quotations/:id', { projectName: '改專案名', validUntil: g.j.validUntil }, { id: 'legacy-1' });
    const stored = store.quotations.find((x) => x.id === 'legacy-1');
    t('7f. 舊單 PUT 把推算值原樣送回 → 200，伺服器沒存（維持「沒存」，列印日期與雜湊不變）', p3.s === 200 && stored.validUntil === undefined && p3.j.validUntil === '2026-10-30', `${p3.s} stored=${stored.validUntil} resp=${p3.j && p3.j.validUntil}`);
    const ii = await call('get', '/api/quotations/:id/issue-info', null, { id: 'legacy-1' });
    t('7g. 舊單 issue-info 的 Remarks 第 4 條印的是 2026 年 10 月 30 日（舊規則）', ii.s === 200 && JSON.stringify(ii.j.remarks).includes('2026年10月30日'), `${ii.s} ${JSON.stringify(ii.j && ii.j.remarks).slice(0, 120)}`);
    const p4 = await call('put', '/api/quotations/:id', { validUntil: '2026-11-30' }, { id: 'legacy-1' });
    t('7h. 舊單業務真的改成別的日期（11-30）→ 才存', p4.s === 200 && stored.validUntil === '2026-11-30', `${p4.s} stored=${stored.validUntil}`);
    // 沒存期限的舊單收到空白（只有直接打 API 才會）：維持改版前行為（舊規則推算），不可因新規則被存成不同日期
    const mkLegacy = (id) => store.quotations.push({ id, owner: 'sales1', quoteNo: 'QU-OLD-' + id, company: '舊單公司', quoteDate: '2026-10-26', createdAt: '2026-10-26T02:00:00.000Z', updatedAt: '2026-10-26T02:00:00.000Z',
      products: [], items: [{ lid: 'l1', desc: 'x', unit: '式', qty: 1, unitPrice: 1000, cost: 0 }], status: 'draft', approval: null, discountType: 'none', discountValue: 0, costFlow: { note: '' } });
    mkLegacy('legacy-2'); mkLegacy('legacy-3'); mkLegacy('legacy-4');
    const b1 = await call('put', '/api/quotations/:id', { validUntil: '' }, { id: 'legacy-2' });
    const b2 = await call('put', '/api/quotations/:id', { validUntil: null }, { id: 'legacy-3' });
    const sb1 = store.quotations.find((x) => x.id === 'legacy-2'), sb2 = store.quotations.find((x) => x.id === 'legacy-3');
    t('7i. 舊單（沒存期限）PUT 空白／null → 200、仍沒存、列印日期維持舊規則 2026-10-30', b1.s === 200 && b2.s === 200 && sb1.validUntil === undefined && sb2.validUntil === undefined && b1.j.validUntil === '2026-10-30' && b2.j.validUntil === '2026-10-30', `${b1.s}/${b2.s} stored=${sb1.validUntil}/${sb2.validUntil} resp=${b1.j && b1.j.validUntil}/${b2.j && b2.j.validUntil}`);
    const b3 = await call('put', '/api/quotations/:id', { quoteDate: '2026-11-23', validUntil: '' }, { id: 'legacy-4' });
    const sb3 = store.quotations.find((x) => x.id === 'legacy-4');
    t('7j. 舊單 PUT 改報價日期為 11/23 並清空期限 → 維持改版前行為：舊規則推算 11-30（新規則會是 12-31）', b3.s === 200 && sb3.validUntil === '2026-11-30' && b3.j.validUntil === '2026-11-30', `${b3.s} stored=${sb3.validUntil} resp=${b3.j && b3.j.validUntil}`);
  } catch (e) {
    t('7. 路由層測試', false, e && e.stack ? e.stack.split('\n').slice(0, 3).join(' | ') : e);
  }

  let pass = 0, fail = 0;
  res.forEach(([n, o, x]) => { console.log((o ? 'PASS ' : 'FAIL ') + n + (x && !o ? '  ← ' + x : '')); o ? pass++ : fail++; });
  console.log('');
  note.forEach((n) => console.log('數字  ' + n));
  console.log(`提醒  假日逐日表涵蓋 ${TWWD.COVERED_YEARS.join('、')} 年；2028 年辦公日曆人事行政總處須於 2027-06-30 前公告，公布後請補進 _client/tw-workdays.js 並重跑本檢查。`);
  console.log(`\n報價期限預設值檢查：PASS ${pass} / FAIL ${fail}`);
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error('檢查腳本錯誤', e.stack); process.exit(2); });
