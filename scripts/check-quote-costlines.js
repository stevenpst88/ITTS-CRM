#!/usr/bin/env node
/**
 * 報價單「成本明細」(costLines) 單元測試。用法：node scripts/check-quote-costlines.js
 * 動 lib/quoteCostLines.js、lib/quoteApproval.js（computeFinancials／contentHash／validateForSubmit）、
 * lib/quoteRoutes.js 的 costComplete／serialize 之後必跑。
 *   1) normalizeCostLines：各種非法輸入、上限 60、印花稅規則、lid 保留；200 萬字元的惡意長字串（ReDoS／CPU 耗盡）每個欄位都要在 50ms 內處理完（1.40–1.44）
 *   2) stampDollars 進位、lineCents 與 computeFinancials 對 items 的取整一致、各分類合計
 *   3) 舊單（沒有 costLines）contentHash 位元級不變：300 張決定性隨機舊單的雜湊串接後比對「改版前」程式產生的 golden 摘要
 *   4) 新式單 hash：任一欄位改變 → hash 變；印花稅金額、列順序不進 hash
 *   5) computeFinancials 新式／舊式、costComplete（QA 與路由兩處定義一致）、MISSING_COST 訊息、簽核層級隨毛利改變
 *   6) 稽核文字無 undefined；describeLinesFull（需顧問切換時的完整留痕）列出全部列、截斷品名（6.8–6.12）
 *   7) 過期分頁保護簽章 QA.itemsSig／QA.costLinesSig：對價格變動不變、對結構／成本明細內容變動會變、回應裡只有 16 碼摘要而沒有明文金額
 *   8) forLids 驗證與輸出、涵蓋判定 CL.unmatchedItems／警告 CL.costWarnings、buildDerived／buildPreview 只在「新式且有未涵蓋品項」時多鍵（其餘位元級相同）、CL.backfillForLids
 * 測試資料只用通用字串（公開 repo：不放客戶名、人名、廠商名、真實費率）。
 */
'use strict';
const path = require('path');
const crypto = require('crypto');
const ROOT = path.join(__dirname, '..');
const QA = require(path.join(ROOT, 'lib/quoteApproval.js'));
const CL = require(path.join(ROOT, 'lib/quoteCostLines.js'));
const registerQuoteRoutes = require(path.join(ROOT, 'lib/quoteRoutes.js'));

/** 「改版前」程式（quoteApproval 未含 costLines 版本）對下面 genLegacyQuotes(300) 的雜湊，依序以換行串接後的 sha256 */
const LEGACY_HASH_GOLDEN = '8ce55afb46fd08c9d81cb73436db34e0e88e2e30925cae05c15c2bce97a16b8b';

// 決定性亂數
function lcg(seed) { let s = seed >>> 0; return () => ((s = (Math.imul(s, 1103515245) + 12345) & 0x7fffffff) / 0x7fffffff); }

/** 隨機「舊式」報價單（不含 costLines）：涵蓋 kind 列、remarks／付款／追加條款、折扣、各種髒數字 */
function genLegacyQuotes(n, seed) {
  const rnd = lcg(seed === undefined ? 20261007 : seed);
  const pick = (a) => a[Math.floor(rnd() * a.length)];
  const num = () => pick([0, 1, 2, 3, 10, 99, 100, 250, 1000, 12345, 0.5, 2.5, 33.333, 1e6, '7', ' 8 ', '1,200', '', null, undefined, '0', 'abc', -5, 1e12]);
  const out = [];
  for (let k = 0; k < n; k++) {
    const q = {};
    const withField = (key, val, p) => { if (rnd() < p) q[key] = val; };
    withField('company', 'TestCo' + k, 0.9); withField('projectName', pick(['P1', ' P2 ', 'proj\r\nx', '']), 0.8);
    withField('projectNo', pick(['', 'N1']), 0.5); withField('contactName', 'C' + k, 0.6);
    withField('address', pick(['', 'addr']), 0.5); withField('phone', pick(['', '02-0000']), 0.5);
    withField('note', pick(['', 'n1', 'line1\r\nline2 ']), 0.5);
    withField('validUntil', pick(['', '2026-12-31', ' 2027-01-31 ']), 0.5);
    withField('payment', { net: pick([30, 45, '60', null]), items: [{ label: pick(['簽約', 'a']), pct: pick([100, '100', 50]) }] }, 0.4);
    withField('extraClauses', pick([['c1'], ['c1', ' ', 'c2'], [], ['']]), 0.4);
    withField('products', pick([['B', 'A'], ['A'], [], ['A', 'A', ' B ']]), 0.8);
    withField('discountType', pick(['none', 'percent', 'amount', '', null, 'weird']), 0.7);
    withField('discountValue', pick([0, 90, '85', 1000, '', null, -1, 'x']), 0.7);
    if (rnd() < 0.8) {
      const cnt = Math.floor(rnd() * 9);
      q.items = [];
      for (let i = 0; i < cnt; i++) {
        const r = rnd();
        if (r < 0.1) q.items.push({ lid: 't' + i, kind: 'title', desc: 'Part ' + i });
        else if (r < 0.2) q.items.push({ lid: 's' + i, kind: 'subtotal', desc: pick(['', '小計']) });
        else if (r < 0.95) {
          const it = { lid: rnd() < 0.9 ? 'x' + i : undefined, desc: 'd' + i, unit: pick(['式', '台', '人天', '']), qty: num(), unitPrice: num(), cost: num() };
          if (rnd() < 0.2) it.cat = pick(['consult', 'software', 'hardware', 'other']);
          if (rnd() < 0.1) it.extra = 'zzz';
          q.items.push(it);
        } else q.items.push(pick([null, 5, 'str']));
      }
    }
    withField('costFlow', { state: pick(['na', 'needed', 'requested', 'filled']), by: 'u', note: 'n' }, 0.3);
    withField('contingencyPct', pick([0, 5, 10]), 0.2);
    out.push(q);
  }
  return out;
}

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);

const L = (cat, desc, qty, unitCost, extra) => Object.assign({ cat, desc, qty, unitCost }, extra || {});
let _n = 0;
const genLid = () => 'g' + (++_n);
const norm = (raw, prev) => CL.normalizeCostLines(raw, { genLid, prev });
const errCode = (r) => (r.ok ? 'OK' : r.error.code);

function run() {
  // ───────────────────────── 1) normalizeCostLines ─────────────────────────
  let r;
  for (const bad of [null, undefined, 'x', 5, {}, true]) t('1.1 非陣列 ' + JSON.stringify(bad) + ' → BAD_COST_LINE', errCode(norm(bad)) === 'BAD_COST_LINE');
  for (const bad of [null, 5, 'x', [], true]) t('1.2 列不是物件 ' + JSON.stringify(bad) + ' → BAD_COST_LINE', errCode(norm([bad])) === 'BAD_COST_LINE');
  r = norm([]); t('1.3 空陣列合法（新式、尚未填）', r.ok && r.lines.length === 0);
  for (const cat of ['hardware', 'CONSULT', '', null, undefined, 5, ['consult'], 'x']) t('1.4 未知分類 ' + JSON.stringify(cat) + ' → BAD_COST_LINE', errCode(norm([L(cat, 'a', 1, 1)])) === 'BAD_COST_LINE');
  for (const cat of CL.CATS) t('1.5 合法分類 ' + cat, norm([L(cat, 'a', 1, 1)]).ok);
  for (const d of [undefined, null, '', '   ', 5, ['a'], {}, true]) t('1.6 desc ' + JSON.stringify(d) + ' → BAD_COST_LINE', errCode(norm([L('consult', d, 1, 1)])) === 'BAD_COST_LINE');
  // 數字欄位
  const badNums = [-1, -0.01, NaN, Infinity, -Infinity, '12abc', '', '  ', null, undefined, true, false, [], {}, '1e3', '0x10', '1,000', 'abc', '--1', '1 2'];
  for (const v of badNums) {
    t('1.7 qty ' + String(v) + ' → BAD_COST_LINE', errCode(norm([L('consult', 'a', v, 1)])) === 'BAD_COST_LINE');
    t('1.8 unitCost ' + String(v) + ' → BAD_COST_LINE', errCode(norm([L('consult', 'a', 1, v)])) === 'BAD_COST_LINE');
  }
  t('1.9 qty 1e9+1 → 拒絕；1e9 → 合法', errCode(norm([L('other', 'a', 1e9 + 1, 0)])) === 'BAD_COST_LINE' && norm([L('other', 'a', 1e9, 0)]).ok);
  t('1.10 unitCost 1e12+1 → 拒絕；1e12 合法（單列金額過大另擋）', errCode(norm([L('other', 'a', 1, 1e12 + 1)])) === 'BAD_COST_LINE' && norm([L('other', 'a', 1, 1e12)]).ok);
  t('1.11 qty 1e9 × unitCost 1e12（單列金額過大）→ BAD_COST_LINE', errCode(norm([L('other', 'a', 1e9, 1e12)])) === 'BAD_COST_LINE');
  t('1.12 數字字串：" 5 "、"0.5"、"5."、".5"、"0" 合法並轉成 number', eq(norm([L('hw', 'a', ' 5 ', '0.5'), L('hw', 'b', '5.', '.5'), L('hw', 'c', '0', '0')]).lines.map((l) => [l.qty, l.unitCost]), [[5, 0.5], [5, 0.5], [0, 0]]));
  t('1.13 qty=0、unitCost=0 允許（done 才要求 qty>0）', norm([L('other', 'a', 0, 0)]).ok);
  t('1.14 -0 正規化為 0', Object.is(norm([L('other', 'a', -0, -0)]).lines[0].qty, 0));
  // 文字欄位
  r = norm([L('consult', ' PM ', 1, 1, { vendor: ' V ', note: ' n ', unit: ' 人天 ' })]).lines[0];
  t('1.15 文字欄位去頭尾空白', r.desc === 'PM' && r.vendor === 'V' && r.note === 'n' && r.unit === '人天');
  r = norm([L('consult', 'a'.repeat(200), 1, 1, { vendor: 'v'.repeat(100), note: 'n'.repeat(300), unit: 'u'.repeat(30) })]).lines[0];
  t('1.16 超過長度截斷（desc120／vendor60／note200／unit10）', r.desc.length === 120 && r.vendor.length === 60 && r.note.length === 200 && r.unit.length === 10);
  t('1.17 unit 空白 → 預設「式」', norm([L('other', 'a', 1, 1)]).lines[0].unit === '式' && norm([L('other', 'a', 1, 1, { unit: '  ' })]).lines[0].unit === '式');
  r = norm([L('other', 'a\r\nb\tc', 1, 1, { note: 'x\ny' })]).lines[0];
  t('1.18 控制字元／換行換成空白', r.desc === 'a b c' && r.note === 'x y', JSON.stringify([r.desc, r.note]));
  for (const k of ['vendor', 'note', 'unit']) t('1.19 ' + k + ' 型別錯（數字）→ BAD_COST_LINE', errCode(norm([L('consult', 'a', 1, 1, { [k]: 5 })])) === 'BAD_COST_LINE');
  t('1.20 forLid 型別錯 → BAD_COST_LINE；合法則保留；超長截斷到 64', errCode(norm([L('consult', 'a', 1, 1, { forLid: 5 })])) === 'BAD_COST_LINE'
    && norm([L('consult', 'a', 1, 1, { forLid: ' it-1 ' })]).lines[0].forLid === 'it-1' && norm([L('consult', 'a', 1, 1, { forLid: 'x'.repeat(99) })]).lines[0].forLid.length === 64);
  t('1.21 廠商只有顧問／軟體／硬體保留；差旅、其他一律丟掉（vendor 欄位存在但為空字串）',
    eq(['consult', 'software', 'hw', 'travel', 'other'].map((c) => norm([L(c, 'a', 1, 1, { vendor: 'V' })]).lines[0].vendor), ['V', 'V', 'V', '', '']));
  r = norm([L('consult', 'a', 1, 1, { evil: 'x', lid: 'nope', cost: 5 })]).lines[0];
  t('1.22 未知欄位不保存；輸出鍵順序固定', eq(Object.keys(r), ['lid', 'cat', 'desc', 'vendor', 'note', 'unit', 'qty', 'unitCost']), Object.keys(r).join(','));
  // 上限
  const many = (n) => Array.from({ length: n }, (_, i) => L('other', 'x' + i, 1, 1));
  t('1.23 剛好 60 列合法、61 列 → TOO_MANY_COST_LINES（400）', norm(many(60)).ok && errCode(norm(many(61))) === 'TOO_MANY_COST_LINES' && norm(many(61)).error.status === 400);
  // 印花稅
  const stampIn = { cat: 'travel', desc: '亂寫', qty: 7, unit: 'kg', unitCost: 99999, vendor: 'V', auto: 'stamp', forLid: 'x', note: ' 備註 ' };
  r = norm([stampIn]).lines[0];
  t('1.24 印花稅強制 cat other／qty 1／unit 式／desc 固定／unitCost 存 0／vendor 空／不保留 forLid', r.cat === 'other' && r.qty === 1 && r.unit === '式' && r.desc === CL.STAMP_DESC && r.unitCost === 0 && r.vendor === '' && r.auto === 'stamp' && r.forLid === undefined && r.note === '備註', JSON.stringify(r));
  t('1.25 印花稅可以省略 qty／unitCost／desc', norm([{ cat: 'other', auto: 'stamp' }]).ok && norm([{ auto: 'stamp' }]).ok);
  r = norm([stampIn, { cat: 'other', auto: 'stamp' }]); t('1.26 第 2 列印花稅 → BAD_COST_LINE', errCode(r) === 'BAD_COST_LINE');
  for (const a of ['Stamp', 'x', 1, true, []]) t('1.27 auto ' + JSON.stringify(a) + ' → BAD_COST_LINE', errCode(norm([L('other', 'a', 1, 1, { auto: a })])) === 'BAD_COST_LINE');
  t('1.28 auto 為 null／空字串視為一般列', norm([L('other', 'a', 1, 1, { auto: null }), L('other', 'b', 1, 1, { auto: '' })]).lines.every((l) => l.auto === undefined));
  // lid
  r = norm([L('other', 'a', 1, 1), L('other', 'b', 1, 1)]);
  t('1.29 新列由 genLid 指派且互不相同', r.ok && r.lines[0].lid !== r.lines[1].lid && /^g\d+$/.test(r.lines[0].lid), r.lines && r.lines.map((l) => l.lid).join(','));
  const prev = [{ lid: 'keep-1', cat: 'other', desc: 'a' }, { lid: 'keep-2', cat: 'other', desc: 'b' }];
  r = norm([L('other', 'a', 1, 1, { lid: 'keep-2' }), L('other', 'b', 1, 1, { lid: 'unknown' }), L('other', 'c', 1, 1)], prev);
  t('1.30 lid 在 prev 裡保留；未知 lid 與沒 lid 都指派新的（不重用 prev 的 lid）', r.ok && r.lines[0].lid === 'keep-2' && r.lines[1].lid !== 'unknown' && r.lines[2].lid !== 'keep-1' && r.lines[2].lid !== 'keep-2', r.lines && r.lines.map((l) => l.lid).join(','));
  r = norm([L('other', 'a', 1, 1, { lid: 'keep-1' }), L('other', 'b', 1, 1, { lid: 'keep-1' })], prev);
  t('1.31 請求內 lid 重複 → DUP_LID（409）', errCode(r) === 'DUP_LID' && r.error.status === 409);
  r = norm([L('other', 'a', 1, 1, { lid: 'dup' }), L('other', 'b', 1, 1, { lid: 'dup' })]);
  t('1.32 即使 prev 沒有也算重複 lid', errCode(r) === 'DUP_LID');
  t('1.33 沒給 genLid 也能指派（內建隨機字串）', CL.normalizeCostLines([L('other', 'a', 1, 1)]).lines[0].lid.length > 4);
  // 冪等與不改輸入
  const rawIn = [L('consult', 'PM', 3, 2000.5, { vendor: 'V', note: 'n', unit: '人天' }), { cat: 'other', auto: 'stamp' }, L('travel', '差旅', 1, 3000)];
  const snap = JSON.stringify(rawIn);
  const a1 = norm(rawIn); const a2 = norm(a1.lines, a1.lines);
  t('1.34 不修改傳入的 raw', JSON.stringify(rawIn) === snap);
  t('1.35 正規化冪等：lines 再丟回去（prev＝lines）完全相同', a1.ok && a2.ok && eq(a1.lines, a2.lines));
  t('1.36 錯誤物件含 code／status／message', (() => { const e = norm('x').error; return typeof e.code === 'string' && e.status === 400 && typeof e.message === 'string' && e.message.length > 0; })());

  // checkDone
  t('1.37 checkDone：空 → MISSING_COST；只有印花稅 → MISSING_COST；總成本 0 → MISSING_COST',
    CL.checkDone([]).code === 'MISSING_COST' && CL.checkDone(norm([{ auto: 'stamp' }]).lines).code === 'MISSING_COST' && CL.checkDone(norm([L('other', 'a', 1, 0)]).lines).code === 'MISSING_COST');
  t('1.38 checkDone：非印花稅列 qty=0 → BAD_COST_LINE；正常 → ok', CL.checkDone(norm([L('other', 'a', 0, 100)]).lines).code === 'BAD_COST_LINE' && CL.checkDone(norm([L('other', 'a', 1, 100)]).lines).ok);
  t('1.39 checkDone：單價 0 的列允許（只要總成本>0）', CL.checkDone(norm([L('other', 'a', 1, 0), L('other', 'b', 1, 5)]).lines).ok);

  // 惡意長字串（ReDoS／CPU 耗盡）：body 上限 2MB，單一欄位可放約 200 萬字元。每個欄位的處理都必須在 50ms 內結束。
  // 舊的 strictNum 正規式是二次方回溯（160k 位數就 8 秒），這組測試在舊寫法下會逾時。
  {
    const HUGE = 2000000;
    const timed = (lines) => { const t0 = process.hrtime.bigint(); const out = norm(lines); return { out, ms: Number(process.hrtime.bigint() - t0) / 1e6 }; };
    const evilNums = [
      ['數字＋非法字元', '1'.repeat(HUGE) + 'x'],
      ['負號＋數字＋非法字元', '-' + '1'.repeat(HUGE) + 'x'],
      ['小數點後長串＋非法字元', '0.' + '1'.repeat(HUGE) + 'x'],
      ['長串小數點', '1' + '.'.repeat(HUGE)],
      ['全是數字（合法格式但過長）', '9'.repeat(HUGE)],
      ['前後空白包住數字', ' '.repeat(HUGE) + '1' + ' '.repeat(10)],
      ['長串零＋非法字元', '0'.repeat(HUGE) + '!'],
    ];
    for (const field of ['qty', 'unitCost']) {
      evilNums.forEach(([nm, s]) => {
        const line = L('other', 'a', 1, 1); line[field] = s;
        const { out, ms } = timed([line]);
        t('1.40 ' + field + ' ' + nm + '（200 萬字元）→ 400 BAD_COST_LINE 且 <50ms', !out.ok && out.error.code === 'BAD_COST_LINE' && out.error.status === 400 && ms < 50, ms.toFixed(2) + 'ms ' + errCode(out));
      });
    }
    // 長度剛好在門檻上下：40 字元內的合法數字照常通過，41 字元一律拒絕
    t('1.41 數字字串長度門檻：40 字元的合法數字通過；41 字元（格式合法但太長）拒絕',
      norm([L('other', 'a', '0.' + '1'.repeat(38), 1)]).ok && norm([L('other', 'a', 1, '0.' + '1'.repeat(38))]).ok
      && errCode(norm([L('other', 'a', '0.' + '1'.repeat(39), 1)])) === 'BAD_COST_LINE' && errCode(norm([L('other', 'a', 1, '0.' + '1'.repeat(39))])) === 'BAD_COST_LINE');
    // 文字欄位與其他欄位：不經過數字正規式，但仍是 200 萬字元的輸入，必須線性處理並截斷
    const textCases = [
      ['desc', { desc: 'a'.repeat(HUGE) }, (l) => l.desc.length === 120],
      ['desc 全是控制字元＋一個字', { desc: '\n'.repeat(HUGE) + 'a' }, (l) => l.desc === 'a'],
      ['desc 全是空白＋一個字', { desc: ' '.repeat(HUGE) + 'a' }, (l) => l.desc === 'a'],
      ['vendor', { cat: 'consult', vendor: 'v'.repeat(HUGE) }, (l) => l.vendor.length === 60],
      ['note', { note: 'n'.repeat(HUGE) }, (l) => l.note.length === 200],
      ['unit', { unit: 'u'.repeat(HUGE) }, (l) => l.unit.length === 10],
      ['forLid', { forLid: 'f'.repeat(HUGE) }, (l) => l.forLid.length === 64],
      ['lid', { lid: 'f'.repeat(HUGE) }, (l) => typeof l.lid === 'string' && l.lid.length > 0],
    ];
    textCases.forEach(([nm, patch, ok]) => {
      const { out, ms } = timed([Object.assign(L('other', 'a', 1, 1), patch)]);
      t('1.42 文字欄位 ' + nm + '（200 萬字元）→ 截斷／線性處理且 <50ms', out.ok && ok(out.lines[0]) && ms < 50, ms.toFixed(2) + 'ms ' + errCode(out));
    });
    const badKey = [['cat', { cat: 'c'.repeat(HUGE) }], ['auto', { auto: 'a'.repeat(HUGE) }]];
    badKey.forEach(([nm, patch]) => {
      const { out, ms } = timed([Object.assign(L('other', 'a', 1, 1), patch)]);
      t('1.43 ' + nm + ' 200 萬字元 → 400 且 <50ms', !out.ok && out.error.status === 400 && ms < 50, ms.toFixed(2) + 'ms ' + errCode(out));
    });
    t('1.44 60 列、每列 desc、note、vendor 各 3 萬字元（共約 5MB）也在 1 秒內完成（總量線性）', (() => { const t0 = Date.now(); const big = 'z'.repeat(30000); const r2 = norm(Array.from({ length: 60 }, (_, i) => L('other', big + i, 1, 1, { note: big, vendor: big }))); return r2.ok && Date.now() - t0 < 1000; })());
  }

  // ───────────────────────── 2) 印花稅進位、取整、合計 ─────────────────────────
  const sd = CL.stampDollars;
  t('2.1 stampDollars 進位：0／負／NaN／undefined → 0', sd(0) === 0 && sd(-5) === 0 && sd(NaN) === 0 && sd(undefined) === 0 && sd('abc') === 0 && sd(null) === 0);
  t('2.2 49999 分（499.99 元）→ 0；50000 分（500 元）→ 1（half-up）', sd(49999) === 0 && sd(50000) === 1);
  t('2.3 99999→1、100000→1、149999→1、150000→2、249999→2、250000→3', sd(99999) === 1 && sd(100000) === 1 && sd(149999) === 1 && sd(150000) === 2 && sd(249999) === 2 && sd(250000) === 3);
  t('2.4 營收 1,000,000 元（1e8 分）→ 1000；990,000 元 → 990；1,234,567.89 元 → 1235', sd(1e8) === 1000 && sd(99000000) === 990 && sd(123456789) === 1235);
  t('2.5 bigint 與極大值：1e15 分 → 1e10 元；bigint 150000n → 2', sd(1e15) === 1e10 && sd(150000n) === 2 && sd(-1n) === 0);
  // lineCents 與 computeFinancials 取整一致
  const rnd = lcg(42);
  let parityOk = 0, parityTotal = 0;
  const rq = () => { const k = Math.floor(rnd() * 6); return k === 0 ? Math.floor(rnd() * 1000) / 1000 + 0.001 : k === 1 ? Math.round(rnd() * 1e6) / 100 + 0.01 : k === 2 ? Math.floor(rnd() * 50) + 1 : k === 3 ? Math.round(rnd() * 1e5) / 1e4 + 0.0001 : k === 4 ? 0.5 : 1 / 3; };
  const rc = () => { const k = Math.floor(rnd() * 6); return k === 0 ? Math.round(rnd() * 1e6) / 1e3 : k === 1 ? Math.floor(rnd() * 99999) : k === 2 ? 0.005 : k === 3 ? 1234.567 : k === 4 ? Math.round(rnd() * 1e9) / 1e5 : 0.015; };
  for (let i = 0; i < 600; i++) {
    const qty = rq(), cost = rc();
    const fin = QA.computeFinancials({ items: [{ lid: 'a', desc: 'a', qty, unitPrice: 1000000, cost }] });
    parityTotal++;
    if (fin.ok && BigInt(fin.costCents) === CL.lineCentsBig({ qty, unitCost: cost }) && fin.costCents === CL.lineCents({ qty, unitCost: cost })) parityOk++;
  }
  t('2.6 lineCents 與 computeFinancials 對 items 的取整完全一致（600 組隨機小數，含 half-up 邊界）', parityOk === parityTotal, parityOk + '/' + parityTotal);
  t('2.7 lineCents：0.005×1 → 1 分（half-up）；1/3×3 → 100 分；壞資料 → NaN／拋 RangeError', CL.lineCents({ qty: 1, unitCost: 0.005 }) === 1 && CL.lineCents({ qty: 3, unitCost: 1 / 3 }) === 100
    && Number.isNaN(CL.lineCents({ qty: 'x', unitCost: 1 })) && (() => { try { CL.lineCentsBig({ qty: -1, unitCost: 1 }); return false; } catch (e) { return e instanceof RangeError; } })());
  // 合計
  const sample = norm([L('consult', 'PM', 2, 1000), L('consult', 'SD', 1.5, 800.5), L('software', 'lic', 1, 5000), L('hw', 'srv', 2, 3333.33),
    L('travel', 'trip', 4, 250), L('other', 'gift', 1, 400), { cat: 'other', auto: 'stamp' }]).lines;
  const qs = { costLines: sample };
  let tot = CL.totalsByCat(qs, 1e8);   // 營收 1,000,000 元 → 印花稅 1000
  t('2.8 各分類合計（分）：consult 320075、software 500000、hw 666666、travel 100000、other 50000（含印花稅 100000）', tot.ok && tot.consult === 320075 + 0 && tot.software === 500000 && tot.hw === 666666 && tot.travel === 100000 && tot.other === 40000 + 100000, JSON.stringify(tot));
  t('2.9 total＝各類加總；stamp＝100000；nonStamp＝total−stamp；count 7／nonStampCount 6', tot.total === tot.consult + tot.software + tot.hw + tot.travel + tot.other && tot.stamp === 100000 && tot.nonStamp === tot.total - 100000 && tot.count === 7 && tot.nonStampCount === 6);
  tot = CL.totalsByCat(qs, 0); t('2.10 營收未知（0）→ 印花稅 0，其他不變', tot.stamp === 0 && tot.other === 40000);
  t('2.11 沒有 costLines → 全 0、ok；空陣列 → 全 0', CL.totalsByCat({}, 1e8).total === 0 && CL.totalsByCat({}, 1e8).ok && CL.totalsByCat({ costLines: [] }, 1e8).count === 0);
  t('2.12 壞資料（qty 為字串垃圾）→ ok:false、code BAD_COST_LINE、index', (() => { const x = CL.totalsByCat({ costLines: [{ cat: 'other', qty: 'x', unitCost: 1 }] }, 0); return !x.ok && x.code === 'BAD_COST_LINE' && x.index === 0; })());
  t('2.13 壞分類計入「其他」（寧可多算成本）', CL.totalsByCat({ costLines: [{ cat: 'zzz', qty: 1, unitCost: 10 }] }, 0).other === 1000);
  t('2.14 effectiveLines：印花稅帶入計算後 unitCost、qty 1；一般列不動；不改原物件', (() => {
    const e = CL.effectiveLines(qs, 1e8); const st = e.find((l) => l.auto === 'stamp');
    return st.unitCost === 1000 && st.qty === 1 && qs.costLines.find((l) => l.auto === 'stamp').unitCost === 0 && e.find((l) => l.desc === 'PM').unitCost === 1000 && CL.effectiveLines({}, 1e8).length === 0;
  })());
  // publicLines（v1.1：印花稅金額對所有 canSeeCost 的人都給，不看 canSeePrice——ITTS 顧問與業務可互看成本與售價）
  const pl = (p) => CL.publicLines(qs, Object.assign({ revenueCents: 1e8 }, p));
  t('2.15 publicLines：!canSeeCost → undefined（無 costLines 欄位的單也是 undefined）', pl({ canSeeCost: false, canSeePrice: true }) === undefined && CL.publicLines({}, { canSeeCost: true, canSeePrice: true }) === undefined);
  t('2.16 publicLines：canSeeCost＋canSeePrice → 印花稅列帶金額 1000；v1.1 只有 canSeeCost（顧問，無 canSeePrice）也帶金額 1000', pl({ canSeeCost: true, canSeePrice: true }).find((l) => l.auto === 'stamp').unitCost === 1000 && pl({ canSeeCost: true, canSeePrice: false }).find((l) => l.auto === 'stamp').unitCost === 1000);
  t('2.17 publicLines：只輸出已知欄位；一般列 unitCost 保留', (() => { const o = pl({ canSeeCost: true, canSeePrice: true }); return o.length === 7 && o.every((l) => Object.keys(l).every((k) => ['lid', 'cat', 'desc', 'vendor', 'note', 'unit', 'qty', 'unitCost', 'auto', 'forLid'].includes(k))) && o.find((l) => l.desc === 'PM').unitCost === 1000; })());

  // ───────────────────────── 3) 舊單雜湊位元級不變 ─────────────────────────
  const legacy = genLegacyQuotes(300);
  const digest = crypto.createHash('sha256').update(legacy.map((q) => QA.contentHash(q)).join('\n')).digest('hex');
  t('3.1 300 張決定性隨機舊單（含 kind 列／付款／追加條款／髒數字）雜湊串接 sha256 ＝ 改版前 golden', digest === LEGACY_HASH_GOLDEN, digest);
  t('3.2 隨機舊單全部不含 costLines', legacy.every((q) => !('costLines' in q)));
  t('3.3 舊單的 computeFinancials 輸出沒有 costBreakdownCents 欄位', legacy.every((q) => { const f = QA.computeFinancials(q); return !('costBreakdownCents' in f); }));

  // ───────────────────────── 4) 新式單 hash ─────────────────────────
  const base = { company: 'TestCo', projectName: 'P', products: ['A'], discountType: 'none', discountValue: 0, items: [{ lid: 'i1', desc: 'svc', unit: '式', qty: 1, unitPrice: 1000000, cost: 0 }] };
  const withLines = (lines) => Object.assign({}, base, { costLines: lines });
  const lines0 = norm([L('consult', 'PM', 2, 1000, { vendor: 'V', note: 'n' }), L('travel', 'trip', 1, 500), { cat: 'other', auto: 'stamp' }]).lines;
  const h0 = QA.contentHash(withLines(lines0));
  t('4.1 加上 costLines（即使是 []）→ hash 與沒有該欄位的舊單不同', QA.contentHash(withLines([])) !== QA.contentHash(base) && h0 !== QA.contentHash(base));
  t('4.2 同內容往返 JSON 後 hash 穩定', QA.contentHash(JSON.parse(JSON.stringify(withLines(lines0)))) === h0);
  const mutate = (i, patch) => withLines(lines0.map((l, j) => (j === i ? Object.assign({}, l, patch) : l)));
  const fields = [['cat', 'software'], ['desc', 'PM2'], ['vendor', 'V2'], ['note', 'n2'], ['unit', '天'], ['qty', 3], ['unitCost', 1001]];
  fields.forEach(([k, v]) => t('4.3 一般列 ' + k + ' 改變 → hash 改變', QA.contentHash(mutate(0, { [k]: v })) !== h0));
  t('4.4 lid 改變 → hash 改變', QA.contentHash(mutate(0, { lid: 'other-lid' })) !== h0);
  t('4.5 新增／刪除一列 → hash 改變', QA.contentHash(withLines(lines0.concat([Object.assign({}, lines0[1], { lid: 'zz' })]))) !== h0 && QA.contentHash(withLines(lines0.slice(1))) !== h0);
  t('4.6 印花稅列 unitCost 不進 hash（改成任何值 hash 都相同）', QA.contentHash(mutate(2, { unitCost: 99999 })) === h0);
  t('4.7 印花稅列的 qty／desc／note／auto 仍進 hash（auto 拿掉 → 不同）', QA.contentHash(mutate(2, { note: 'x' })) !== h0 && QA.contentHash(mutate(2, { auto: '' })) !== h0);
  t('4.8 列順序不進 hash（與 items 一樣依 lid 排序）', QA.contentHash(withLines(lines0.slice().reverse())) === h0);
  t('4.9 forLid 不進 hash', QA.contentHash(mutate(0, { forLid: 'abc' })) === h0);
  t('4.10 新式單 items[].cost 仍照舊進 hash（改 items.cost → hash 變）', QA.contentHash(Object.assign({}, withLines(lines0), { items: [Object.assign({}, base.items[0], { cost: 5 })] })) !== h0);
  t('4.11 數字表示法不同但值相同（3 與 "3"）→ hash 相同（normNum）', QA.contentHash(mutate(0, { qty: '2' })) === h0);

  // ───────────────────────── 5) computeFinancials／簽核 ─────────────────────────
  const quote = (lines, extra) => Object.assign({ company: 'TestCo', products: ['P_CONSULT'], items: [{ lid: 'i1', desc: 'svc', unit: '式', qty: 1, unitPrice: 1000000, cost: 123456 }], costLines: lines }, extra || {});
  // 營收 1,000,000；顧問 2×200,000＝400,000；差旅 20,000；交際 10,000；印花稅 1,000
  const goodLines = norm([L('consult', 'PM', 2, 200000), L('travel', 'trip', 1, 20000), L('other', 'gift', 1, 10000), { cat: 'other', auto: 'stamp' }]).lines;
  let fin = QA.computeFinancials(quote(goodLines));
  t('5.1 新式：成本＝顧問＋差旅＋交際＋印花稅＝431,000 元；items[].cost（123456）不參與', fin.ok && fin.costCents === 43100000 && fin.revenueCents === 100000000 && fin.gpCents === 56900000, JSON.stringify(fin));
  t('5.2 新式：毛利率 56.90%、costComplete、missingCostRows 空、costBreakdownCents 分類正確（other 含印花稅）', QA.marginText(fin.gpCents, fin.revenueCents) === '56.90' && fin.costComplete && fin.missingCostRows.length === 0
    && eq(fin.costBreakdownCents, { consult: 40000000, software: 0, hw: 0, travel: 2000000, other: 1100000 }), JSON.stringify(fin.costBreakdownCents));
  fin = QA.computeFinancials(quote(goodLines, { discountType: 'percent', discountValue: 90 }));
  t('5.3 印花稅依「折扣後」未稅營收：九折 900,000 → 印花稅 900；成本 430,900', fin.ok && fin.revenueCents === 90000000 && fin.costCents === 43090000, JSON.stringify(fin));
  t('5.4 成本不含 contingency（q.contingencyPct 不影響簽核毛利）', QA.computeFinancials(quote(goodLines, { contingencyPct: 20 })).costCents === 43100000);
  const oldQ = { company: 'TestCo', products: ['P_CONSULT'], items: [{ lid: 'i1', desc: 'svc', unit: '式', qty: 1, unitPrice: 1000000, cost: 600000 }] };
  fin = QA.computeFinancials(oldQ);
  t('5.5 舊式：成本取 items[].cost（600,000），沒有 costBreakdownCents', fin.ok && fin.costCents === 60000000 && fin.costComplete && !('costBreakdownCents' in fin));
  t('5.6 新式 costLines 是 [] → costComplete false、成本 0', (() => { const f = QA.computeFinancials(quote([])); return f.ok && !f.costComplete && f.costCents === 0; })());
  t('5.7 只有印花稅 → costComplete false；總成本 0 的一般列 → false；qty 0 → false；有成本 → true', ['stamp', 'zero', 'qty0', 'ok'].every((k) => {
    const lines = k === 'stamp' ? [{ cat: 'other', auto: 'stamp' }] : k === 'zero' ? [L('other', 'a', 1, 0)] : k === 'qty0' ? [L('other', 'a', 0, 100)] : [L('other', 'a', 1, 100)];
    return QA.computeFinancials(quote(norm(lines).lines)).costComplete === (k === 'ok');
  }));
  t('5.8 新式壞資料（qty 非數字）→ computeFinancials 回 BAD_COST_LINE，不丟例外', QA.computeFinancials(quote([{ lid: 'z', cat: 'other', desc: 'a', qty: 'x', unitCost: 1 }])).code === 'BAD_COST_LINE');
  t('5.9 新式：成本明細金額過大 → TOO_LARGE', QA.computeFinancials(quote([{ lid: 'z', cat: 'other', desc: 'a', qty: 1e9, unitCost: 1e12 }])).code === 'TOO_LARGE');
  // validateForSubmit 訊息
  const classes = { P_CONSULT: { cls: 'consult', costBySales: true } };
  let v = QA.validateForSubmit(quote([]), classes);
  const mc = v.errors.find((e) => e.code === 'MISSING_COST');
  t('5.10 新式 validateForSubmit：空成本明細 → MISSING_COST「尚未填寫成本明細」（不列第 N 列）', !v.ok && mc && mc.message === '尚未填寫成本明細' && eq(mc.rows, []), JSON.stringify(v.errors));
  v = QA.validateForSubmit({ company: 'TestCo', products: ['P_CONSULT'], items: [{ lid: 'i1', desc: 'svc', qty: 1, unitPrice: 100, cost: 0 }] }, classes);
  t('5.11 舊式 validateForSubmit 訊息不變（第 1 列有單價但成本未填…）', v.errors.some((e) => e.code === 'MISSING_COST' && e.message.indexOf('第 1 列有單價但成本未填或為 0') === 0));
  v = QA.validateForSubmit(quote(goodLines), classes);
  t('5.12 新式完整 → 可送簽、derived 有毛利率 56.90、顧問服務列一級主管即可（≥25%）', v.ok && v.derived && v.derived.marginText === '56.90' && v.derived.level === 1 && eq(v.derived.tiers, ['mgr1']), JSON.stringify(v.errors) + (v.derived && v.derived.marginText));
  // 簽核層級隨毛利改變（顧問服務 l1=25、l2=15）：營收 1,000,000，成本明細調整
  const lvl = (costYuan) => { const d = QA.buildDerived(quote(norm([L('consult', 'PM', 1, costYuan)]).lines), classes); return d && [d.level, d.marginText]; };
  // 印花稅列不在 → 成本就是 costYuan
  t('5.13 毛利 25.00% → 一級主管（level 1）；毛利 24.99% → 總經理（2）；毛利 15.00% → 2；14.99% → 董事長（3）', eq(lvl(750000), [1, '25.00']) && eq(lvl(750100), [2, '24.99']) && eq(lvl(850000), [2, '15.00']) && eq(lvl(850100), [3, '14.99']), JSON.stringify([lvl(750000), lvl(750100), lvl(850000), lvl(850100)]));
  // 印花稅／差旅讓毛利壓過門檻
  const edge = (extraTravel) => QA.buildDerived(quote(norm([L('consult', 'PM', 1, 750000 - 1000)].concat(extraTravel ? [L('travel', 'trip', 1, extraTravel)] : []).concat([{ cat: 'other', auto: 'stamp' }])).lines), classes);
  t('5.14 印花稅（1,000）計入：顧問 749,000＋印花稅 1,000＝750,000 → 剛好 25.00% 一級主管；再加差旅 1 元 → 掉到總經理', edge(0).level === 1 && edge(1).level === 2 && edge(1).marginText === '24.99', edge(0).marginText + '/' + edge(1).marginText);
  // buildDerived 在成本不完整時為 null
  t('5.15 新式成本不完整 → buildDerived 回 null', QA.buildDerived(quote([]), classes) === null);

  // costComplete 兩處定義一致（QA.computeFinancials.costComplete vs 路由的 costComplete）
  const noop = () => {};
  const internal = registerQuoteRoutes({ get: noop, put: noop, post: noop, delete: noop, use: noop }, {
    db: {}, loadAuth: noop, saveAuth: noop, requireAuth: noop, requireAdmin: noop, writeLog: noop, pushNotification: noop, getViewableOwners: noop,
    sanitizeStr: (x, n) => String(x == null ? '' : x).trim().slice(0, n), genQuoteNo: noop, taipeiToday: () => '2026-10-07', resolveIssuer: noop,
    buildQuoteWorkbook: noop, buildQuotePnlExcel: noop, QUOTE_TEMPLATE: '', uuidv4: () => 'x', normalizeBu: noop, getUserFeatures: noop,
  })._internal;
  const cfg = { productClasses: {}, roster: {} };
  const rr = lcg(99); const pk = (a) => a[Math.floor(rr() * a.length)];
  let agree = 0, cases = 0, trues = 0;
  for (let i = 0; i < 300; i++) {
    const n = Math.floor(rr() * 5);
    const ls = []; let stamped = false;
    for (let k = 0; k < n; k++) {
      if (!stamped && rr() < 0.3) { ls.push({ cat: 'other', auto: 'stamp' }); stamped = true; } else ls.push(L(pk(CL.CATS), 'l' + k, pk([0, 1, 2.5]), pk([0, 0, 100, 0.004, 5000])));
    }
    const nl = norm(ls);
    if (!nl.ok) continue;
    const q = { products: [], items: [{ lid: 'i1', desc: 'a', qty: 1, unitPrice: 100000, cost: pk([0, 5]) }], costLines: nl.lines };
    const a = QA.computeFinancials(q).costComplete, b = internal.costComplete(q, cfg);
    cases++; if (a === b) agree++; if (a) trues++;
  }
  t('5.16 costComplete 兩處定義（QA／路由）在 300 組隨機新式單完全一致', cases > 250 && agree === cases && trues > 20 && trues < cases, agree + '/' + cases + ' true=' + trues);
  t('5.17 路由 costComplete：舊式單行為不變；新式單 items 為空 → false', internal.costComplete({ products: [], items: [{ lid: 'a', unitPrice: 5, cost: 3 }] }, cfg) === true
    && internal.costComplete({ products: [], items: [{ lid: 'a', unitPrice: 5, cost: 0 }] }, cfg) === false && internal.costComplete({ products: [], items: [], costLines: goodLines }, cfg) === false);

  // ───────────────────────── 6) 稽核文字 ─────────────────────────
  const A = norm([L('consult', 'PM', 2, 1000, { vendor: 'V' }), L('travel', 'trip', 1, 500), { cat: 'other', auto: 'stamp' }]).lines;
  const B = A.map((l) => Object.assign({}, l)); B[0].qty = 3; B[0].vendor = 'V2'; B.splice(1, 1);
  B.push(Object.assign({}, norm([L('other', '新列', 1, 10)]).lines[0]));
  const sm = CL.summarizeChanges(A, B, 10);
  t('6.1 summarizeChanges：有新增／刪除／修改、不含 undefined／NaN／null', sm.length === 3 && sm.some((s) => s.indexOf('新增') === 0) && sm.some((s) => s.indexOf('刪除') === 0) && sm.some((s) => s.indexOf('數量 2→3') >= 0 && s.indexOf('廠商 V→V2') >= 0)
    && !/undefined|NaN|null/.test(sm.join('')), sm.join(' / '));
  t('6.2 summarizeChanges 超過上限：只留前 N 筆並加「…另 M 筆」', (() => { const many2 = norm(Array.from({ length: 30 }, (_, i) => L('other', 'x' + i, 1, 1))).lines; const s = CL.summarizeChanges([], many2, 10); return s.length === 11 && s[10] === '…另 20 筆'; })());
  t('6.3 summarizeChanges：無異動 → []；只換順序 → 「列順序調整」', CL.summarizeChanges(A, A.map((l) => Object.assign({}, l))).length === 0 && eq(CL.summarizeChanges(A, A.slice().reverse()), ['列順序調整']));
  t('6.4 diffCounts：新增 1、刪除 1、修改 1', eq(CL.diffCounts(A, B), { added: 1, removed: 1, changed: 1 }));
  t('6.5 summarizeChanges 對壞資料（缺欄位、null 列）不丟例外也不出現 undefined', (() => { const s = CL.summarizeChanges([{ lid: 'a' }, null], [{ lid: 'b', cat: 'zzz' }], 10).join(''); return !/undefined/.test(s); })());
  t('6.6 describeLines：「品名」=金額；印花稅不寫金額；超過上限有「…另」', CL.describeLines(A) === '「PM」=2000、「trip」=500、「印花稅」' && /…另 5 筆$/.test(CL.describeLines(norm(Array.from({ length: 15 }, (_, i) => L('other', 'x' + i, 1, 1))).lines)) && CL.describeLines(undefined) === '');
  // describeLinesFull：flip 稽核的完整留痕（舊的 describeLines 只留前 10 筆，其餘變成「…另 N 筆」，無法追回）
  t('6.8 describeLinesFull：逐列「品名」=數量×單價=金額（廠商）；印花稅只寫「印花稅」；無廠商不加括號', CL.describeLinesFull(A) === '「PM」=2×1000=2000（V）、「trip」=1×500=500、「印花稅」', CL.describeLinesFull(A));
  {
    const n30 = norm(Array.from({ length: 30 }, (_, i) => L(CL.CATS[i % 5], '工項' + i, 1 + (i % 4), 1000 + i, { vendor: 'W' + i }))).lines;
    const full = CL.describeLinesFull(n30);
    const miss = n30.filter((l) => full.indexOf('「' + l.desc + '」=' + l.qty + '×' + l.unitCost + '=' + (l.qty * l.unitCost)) < 0);
    t('6.9 describeLinesFull：30 列全部出現（不像 describeLines 只有 10 列）、沒有「…另」', miss.length === 0 && full.indexOf('…另') < 0 && CL.describeLines(n30).indexOf('…另 20 筆') >= 0, 'missing=' + miss.length);
    const n60 = norm(Array.from({ length: 60 }, (_, i) => L('other', 'x' + i, 1, 1))).lines;
    t('6.10 describeLinesFull：資料上限 60 列也全部列出（60 段、不截斷）；超過 maxLines 才加「…另 N 筆」', CL.describeLinesFull(n60).split('、').length === 60 && CL.describeLinesFull(n60).indexOf('…另') < 0
      && CL.describeLinesFull(n60, { maxLines: 50 }).split('、').length === 51 && /…另 10 筆$/.test(CL.describeLinesFull(n60, { maxLines: 50 })));
    const longName = norm([L('consult', '品'.repeat(120), 1, 5, { vendor: '商'.repeat(60) })]).lines;
    const one = CL.describeLinesFull(longName);
    t('6.11 describeLinesFull：品名／廠商各截斷 30 字，單列長度合理', one === '「' + '品'.repeat(30) + '」=1×5=5（' + '商'.repeat(30) + '）' && one.length < 80, one.length);
    t('6.12 describeLinesFull：壞資料（null 列、缺欄位、非數字）不丟例外、不出現 undefined；金額寫「?」；非陣列 → 空字串', (() => {
      const s = CL.describeLinesFull([null, {}, { desc: 'x', qty: 'z', unitCost: 1 }, { desc: '換\n行', qty: 1, unitCost: 2, vendor: 'v\tx' }]);
      return s.indexOf('undefined') < 0 && s.indexOf('=?') >= 0 && s.indexOf('\n') < 0 && s.indexOf('\t') < 0 && CL.describeLinesFull(undefined) === '' && CL.describeLinesFull('x') === '';
    })());
  }
  t('6.7 hasCostLines／isStamp', CL.hasCostLines({ costLines: [] }) && !CL.hasCostLines({}) && !CL.hasCostLines(null) && !CL.hasCostLines({ costLines: null }) && CL.isStamp({ auto: 'stamp' }) && !CL.isStamp({}) && !CL.isStamp(null));

  // ───────────────────────── 7) 過期分頁保護簽章（itemsSig／costLinesSig） ─────────────────────────
  {
    const mkItems = () => [
      { lid: 'a', desc: '品項甲', unit: '式', qty: 2, unitPrice: 123456, cost: 11 },
      { lid: 'b', desc: '品項乙', unit: '台', qty: 1, unitPrice: 987654, cost: 22 },
    ];
    const LINES0 = norm([L('consult', 'PM', 2, 1000, { vendor: 'V', note: 'n', forLid: 'a' }), L('travel', 'trip', 1, 500), { cat: 'other', auto: 'stamp' }], []).lines;   // lid 只指派一次（每次重新 norm 會換 lid）
    const mkLines = () => JSON.parse(JSON.stringify(LINES0));
    const mkQ = (extra) => Object.assign({ id: 'q-sig-1', items: mkItems(), costLines: mkLines(), discountType: 'none', discountValue: 0 }, extra || {});
    const HEX16 = /^[0-9a-f]{16}$/;
    const base = mkQ();
    const iS = QA.itemsSig(base), cS = QA.costLinesSig(base);
    t('7.1 itemsSig／costLinesSig 是 16 碼十六進位摘要；同資料同值；JSON 往返後相同', HEX16.test(iS) && HEX16.test(cS) && iS === QA.itemsSig(mkQ()) && cS === QA.costLinesSig(mkQ())
      && iS === QA.itemsSig(JSON.parse(JSON.stringify(base))) && cS === QA.costLinesSig(JSON.parse(JSON.stringify(base))), iS + ' ' + cS);
    t('7.2 摘要不含明文金額：輸出只有 16 碼 hex（不含單價 123456／987654 或成本 1000 的字串）', !/123456|987654/.test(iS + cS) && iS.length === 16 && cS.length === 16);
    // itemsSig：對價格／成本變動不變
    const withItems = (fn) => mkQ({ items: fn(mkItems()) });
    t('7.3 itemsSig 對「價格金額」變動不變（單價 123456→1、維持有價；cost 改值；折扣；costLines 改動）', QA.itemsSig(withItems((x) => { x[0].unitPrice = 1; return x; })) === iS
      && QA.itemsSig(withItems((x) => { x[0].cost = 99999; x[1].cost = 0; return x; })) === iS && QA.itemsSig(mkQ({ discountType: 'percent', discountValue: 90 })) === iS
      && QA.itemsSig(mkQ({ costLines: [] })) === iS && QA.itemsSig(mkQ({ costLines: undefined })) === iS);
    // itemsSig：結構變動會變
    const chg = (fn) => QA.itemsSig(withItems(fn)) !== iS;
    t('7.4 itemsSig 對結構變動會變：新增品項、刪除品項、改說明、改數量、改單位、改 lid', chg((x) => x.concat([{ lid: 'c', desc: '新', unit: '式', qty: 1, unitPrice: 5 }])) && chg((x) => x.slice(1)) && chg((x) => { x[0].desc = '品項甲2'; return x; })
      && chg((x) => { x[0].qty = 3; return x; }) && chg((x) => { x[0].unit = '套'; return x; }) && chg((x) => { x[0].lid = 'a2'; return x; }));
    t('7.5 itemsSig：有價↔無價（單價 0）會變（贈品列改成有價必須讓顧問重看），與 lineStructureSig 同一條規則', chg((x) => { x[0].unitPrice = 0; return x; }) && chg((x) => { x[1].unitPrice = '0'; return x; }));
    t('7.6 itemsSig：分組標題／小計列增減不影響（與 lineStructureSig 一致）；列順序不影響', QA.itemsSig(withItems((x) => [{ lid: 't', kind: 'title', desc: 'Part A' }].concat(x, [{ lid: 's', kind: 'subtotal', desc: '小計' }]))) === iS
      && QA.itemsSig(withItems((x) => x.slice().reverse())) === iS);
    // costLinesSig
    t('7.7 costLinesSig：沒有 costLines（舊式單）→ 空字串；空陣列（新式、尚未填）→ 16 碼摘要，兩者不同', QA.costLinesSig({ id: 'q', items: [] }) === '' && QA.costLinesSig({ id: 'q', costLines: undefined }) === '' && QA.costLinesSig({ id: 'q', costLines: null }) === ''
      && HEX16.test(QA.costLinesSig({ id: 'q', costLines: [] })));
    const withLines = (fn) => mkQ({ costLines: fn(mkLines()) });
    const cchg = (fn) => QA.costLinesSig(withLines(fn)) !== cS;
    t('7.8 costLinesSig 任一欄位改變都會變：lid、cat、desc、vendor、note、unit、qty、unitCost、forLid', cchg((x) => { x[0].lid = 'zz'; return x; }) && cchg((x) => { x[0].cat = 'software'; return x; }) && cchg((x) => { x[0].desc = 'PM2'; return x; })
      && cchg((x) => { x[0].vendor = 'V2'; return x; }) && cchg((x) => { x[0].note = 'n2'; return x; }) && cchg((x) => { x[0].unit = '人天'; return x; }) && cchg((x) => { x[0].qty = 3; return x; })
      && cchg((x) => { x[0].unitCost = 1001; return x; }) && cchg((x) => { x[0].forLid = 'b'; return x; }) && cchg((x) => { delete x[0].forLid; return x; }));
    t('7.9 costLinesSig：新增／刪除列、改列順序、印花稅列的 auto 拿掉都會變', cchg((x) => x.concat(norm([L('hw', 'srv', 1, 1)]).lines)) && cchg((x) => x.slice(1)) && cchg((x) => x.slice().reverse()) && cchg((x) => { delete x[2].auto; return x; }));
    t('7.10 costLinesSig 不隨品項／折扣／風險預留變動（只看成本明細內容）', QA.costLinesSig(withItems((x) => { x[0].qty = 9; x[0].unitPrice = 5; return x; })) === cS && QA.costLinesSig(mkQ({ discountType: 'amount', discountValue: 1000, contingencyPct: 10 })) === cS);
    t('7.11 不同單（id 不同）的摘要不同；缺 id 也不丟錯', QA.itemsSig(mkQ({ id: 'q-sig-2' })) !== iS && QA.costLinesSig(mkQ({ id: 'q-sig-2' })) !== cS && HEX16.test(QA.itemsSig({ items: [] })) && HEX16.test(QA.itemsSig(null)) && QA.costLinesSig(null) === '');
    t('7.12 壞資料（costLines 含 null／非物件、欄位缺漏）不丟例外', (() => { try { return HEX16.test(QA.costLinesSig({ id: 'q', costLines: [null, 5, 'x', {}] })); } catch (e) { return false; } })());
    t('7.13 不改動傳入的物件', (() => { const q = mkQ(); const snap = JSON.stringify(q); QA.itemsSig(q); QA.costLinesSig(q); return JSON.stringify(q) === snap; })());
  }

  // ───────────────────────── 8) forLids／涵蓋判定（unmatchedItems）／警告／backfillForLids ─────────────────────────
  {
    const OKN = (raw, prev) => norm(raw, prev);
    // 8.1~8.9 forLids 驗證與輸出
    let x = OKN([L('hw', 'a', 1, 1, { forLid: 'f1', forLids: [' f2 ', 'f3', 'f2', '', '   ', 'f1', 'x'.repeat(99)] })]);
    t('8.1 forLids 清洗：去頭尾空白、去空白元素與重複、剔除與 forLid 相同者、每個截 64 字', x.ok && eq(x.lines[0].forLids, ['f2', 'f3', 'x'.repeat(64)]) && x.lines[0].forLid === 'f1', JSON.stringify(x.lines && x.lines[0].forLids));
    for (const bad of ['abc', 5, {}, true, { 0: 'a', length: 1 }]) t('8.2 forLids 非陣列 ' + JSON.stringify(bad) + ' → 400 BAD_COST_LINE', errCode(OKN([L('hw', 'a', 1, 1, { forLids: bad })])) === 'BAD_COST_LINE' && OKN([L('hw', 'a', 1, 1, { forLids: bad })]).error.status === 400);
    for (const bad of [5, null, {}, ['a'], true]) t('8.3 forLids 含非字串元素 ' + JSON.stringify(bad) + ' → 400', errCode(OKN([L('hw', 'a', 1, 1, { forLids: ['ok', bad] })])) === 'BAD_COST_LINE');
    const ids = (n) => Array.from({ length: n }, (_, i) => 'id' + i);
    t('8.4 forLids 恰 60 個合法、61 個 → 400（超量不默默截斷）', OKN([L('hw', 'a', 1, 1, { forLids: ids(60) })]).ok && OKN([L('hw', 'a', 1, 1, { forLids: ids(60) })]).lines[0].forLids.length === 60
      && errCode(OKN([L('hw', 'a', 1, 1, { forLids: ids(61) })])) === 'BAD_COST_LINE');
    t('8.5 forLids 為 null／undefined／[]／全空白 → 合法且不輸出該鍵', [null, undefined, [], ['', '  ']].every((v) => { const o = OKN([L('hw', 'a', 1, 1, { forLids: v })]); return o.ok && !('forLids' in o.lines[0]); }));
    x = OKN([{ cat: 'other', auto: 'stamp', forLid: 'f', forLids: ['a', 'b'] }]);
    t('8.6 印花稅列的 forLid／forLids 一律丟掉；但 forLids 型別錯誤仍 400（與 forLid 一致）', x.ok && !('forLid' in x.lines[0]) && !('forLids' in x.lines[0]) && errCode(OKN([{ cat: 'other', auto: 'stamp', forLids: 'x' }])) === 'BAD_COST_LINE');
    x = OKN([L('consult', 'a', 1, 1, { forLid: 'f', forLids: ['g'] })]);
    t('8.7 輸出鍵順序固定：…unitCost,forLid,forLids；只有 forLids 沒有 forLid 也行', eq(Object.keys(x.lines[0]), ['lid', 'cat', 'desc', 'vendor', 'note', 'unit', 'qty', 'unitCost', 'forLid', 'forLids'])
      && eq(Object.keys(OKN([L('consult', 'a', 1, 1, { forLids: ['g'] })]).lines[0]), ['lid', 'cat', 'desc', 'vendor', 'note', 'unit', 'qty', 'unitCost', 'forLids']));
    t('8.8 冪等：forLids 再丟回 normalize 不變；不改傳入的 raw', (() => { const raw = [L('hw', 'a', 1, 1, { forLid: 'f1', forLids: ['g1', 'g2'] })]; const s0 = JSON.stringify(raw); const a1 = OKN(raw); const a2 = OKN(a1.lines, a1.lines); return JSON.stringify(raw) === s0 && eq(a1.lines, a2.lines); })());
    {
      const ln1 = OKN([L('consult', 'PM', 2, 1000, { forLid: 'f1', forLids: ['g1', 'g2'] }), L('hw', 'srv', 1, 5)]).lines;
      const ln2 = ln1.map((l, i) => (i === 0 ? Object.assign({}, l, { forLids: ['zzz'] }) : l));
      const ln3 = ln1.map((l) => { const o = Object.assign({}, l); delete o.forLids; return o; });
      const pq = (lines) => Object.assign({}, base, { costLines: lines });
      t('8.9 forLids 不影響金額／contentHash／摘要與異動明細（只供涵蓋判定）；costLinesSig 也不含 forLids（併列內容本來就會一起變）',
        QA.contentHash(pq(ln1)) === QA.contentHash(pq(ln2)) && QA.contentHash(pq(ln1)) === QA.contentHash(pq(ln3)) && CL.totalsByCat(pq(ln1), 1e8).total === CL.totalsByCat(pq(ln3), 1e8).total
        && CL.summarizeChanges(ln1, ln2).length === 0 && CL.summarizeChanges(ln1, ln3).length === 0 && eq(CL.diffCounts(ln1, ln2), { added: 0, removed: 0, changed: 0 })
        && QA.costLinesSig(Object.assign({ id: 'z' }, pq(ln1))) === QA.costLinesSig(Object.assign({ id: 'z' }, pq(ln2))));
      const pub = CL.publicLines(pq(ln1), { canSeeCost: true, canSeePrice: false, revenueCents: 1e8 });
      t('8.10 publicLines 輸出 forLid／forLids（只需 canSeeCost）；沒有 forLids 的列沒有該鍵；壞值被濾掉', eq(pub[0].forLids, ['g1', 'g2']) && pub[0].forLid === 'f1' && !('forLids' in pub[1]) && !('forLid' in pub[1])
        && eq(CL.publicLines({ costLines: [{ lid: 'z', cat: 'hw', desc: 'a', qty: 1, unitCost: 1, forLids: ['ok', 5, null, '', 'ok2'] }] }, { canSeeCost: true })[0].forLids, ['ok', 'ok2'])
        && CL.publicLines(pq(ln1), { canSeeCost: false }) === undefined);
    }

    // 8.20~ 涵蓋判定 unmatchedItems
    const it8 = (lid, desc, price, extra) => Object.assign({ lid, desc, unit: '式', qty: 1, unitPrice: price }, extra || {});
    const cl8 = (desc, extra) => Object.assign({ lid: 'c-' + desc, cat: 'other', desc, vendor: '', note: '', unit: '式', qty: 1, unitCost: 10 }, extra || {});
    const un = (items, lines) => CL.unmatchedItems({ items, costLines: lines }).map((u) => u.name || '(空白)');
    t('8.20 舊式單（沒有 costLines 欄位）→ []；新式 costLines=[] 且品項有價 → 全部未涵蓋', CL.unmatchedItems({ items: [it8('a', 'A', 100)] }).length === 0 && CL.unmatchedItems(null).length === 0
      && eq(un([it8('a', 'A', 100), it8('b', 'B', 5)], []), ['A', 'B']));
    t('8.21 forLid 命中品項 lid → 涵蓋（改名也算）；forLids 命中也算；兩者可併用', eq(un([it8('a', 'A', 100), it8('b', 'B', 5), it8('c', 'C', 5)], [cl8('舊名', { forLid: 'a' }), cl8('整包', { forLid: 'b', forLids: ['c'] })]), [])
      && eq(un([it8('a', 'A', 100), it8('b', 'B', 5)], [cl8('整包', { forLids: ['a', 'b'] })]), []) && eq(un([it8('a', 'A', 100), it8('b', 'B', 5)], [cl8('整包', { forLids: ['a'] })]), ['B']));
    t('8.22 沒有 forLid 的列以品名（去頭尾空白）涵蓋；同名 N 個品項要 N 列同名成本列', eq(un([it8('a', ' A ', 100)], [cl8('A')]), []) && eq(un([it8('a', 'A', 100), it8('b', 'A', 100), it8('c', 'A', 100)], [cl8('A'), cl8('A', { lid: 'c2' })]), ['A'])
      && eq(un([it8('a', 'A', 100), it8('b', 'A', 100)], [cl8('A')]), ['A']));
    t('8.23 無價（0／空／負／亂字串／缺欄位）品項不算；title／subtotal／null／非物件不算', eq(un([it8('a', 'A', 0), it8('b', 'B', ''), it8('c', 'C', -5), it8('d', 'D', 'abc'), { lid: 'e', desc: 'E' }, { lid: 't', kind: 'title', desc: 'T', unitPrice: 9 }, { lid: 's', kind: 'subtotal', desc: 'S', unitPrice: 9 }, null, 5, 'x'], []), [])
      && eq(un([it8('a', 'A', '100'), it8('b', 'B', 0.01)], []), ['A', 'B']));
    t('8.24 印花稅列不參與：不涵蓋同名品項、其 forLid 也不涵蓋', eq(un([it8('a', CL.STAMP_DESC, 100), it8('b', 'B', 5)], [{ lid: 's', cat: 'other', desc: CL.STAMP_DESC, qty: 1, unitCost: 0, auto: 'stamp', forLid: 'b' }]), [CL.STAMP_DESC, 'B']));
    t('8.25 指向無價（現有）品項的列算「指向品項」，不再進品名池：不會順便涵蓋同名的有價品項', eq(un([it8('g', '贈品', 0), it8('p', '贈品', 100)], [cl8('贈品', { forLid: 'g' })]), ['贈品']));
    t('8.26 forLid 指向已刪除品項、forLids 全是已刪除 → 退回以品名涵蓋', eq(un([it8('a', 'A', 100)], [cl8('A', { forLid: 'gone' })]), []) && eq(un([it8('a', 'A', 100)], [cl8('A', { forLid: 'gone', forLids: ['gone2'] })]), [])
      && eq(un([it8('a', 'A', 100)], [cl8('別的', { forLid: 'gone' })]), ['A']));
    t('8.27 品名空白的有價品項：只能靠 forLid／forLids 涵蓋，列出時 name 為空', eq(un([it8('a', '', 900000)], [cl8('x')]), ['(空白)']) && eq(un([it8('a', '  ', 900000)], [cl8('x', { forLid: 'a' })]), []) && eq(un([it8('a', '', 5)], [cl8('x', { forLids: ['a'] })]), []));
    t('8.28 沒有 lid 的品項用 legacy-<索引> 代替（與 serialize 一致）', eq(un([it8(undefined, 'A', 5), it8(undefined, 'B', 5)], [cl8('x', { forLid: 'legacy-1' })]), ['A']) && CL.unmatchedItems({ items: [it8('', 'A', 5)], costLines: [] })[0].lid === 'legacy-0');
    t('8.29 品名清洗：控制字元（換行、tab）當空白，與成本列品名同一套', eq(un([it8('a', 'A\tB\nC', 5)], [cl8('A B C')]), []));
    t('8.30 回傳 index＝品項在 q.items 的位置（含 kind 列）、依品項順序', (() => { const r8 = CL.unmatchedItems({ items: [{ lid: 't', kind: 'title', desc: 'T' }, it8('a', 'A', 5), it8('b', 'B', 5)], costLines: [cl8('A')] }); return r8.length === 1 && r8[0].name === 'B' && r8[0].index === 2 && r8[0].lid === 'b'; })());
    t('8.31 不改動傳入的 q', (() => { const q8 = { items: [it8('a', 'A', 5)], costLines: [cl8('x', { forLids: ['a'] })] }; const s8 = JSON.stringify(q8); CL.unmatchedItems(q8); CL.costWarnings(q8); return JSON.stringify(q8) === s8; })());

    // 警告文字
    const w1 = CL.costWarnings({ items: [it8('a', '維護服務', 123456789), it8('b', '教育訓練', 5)], costLines: [] });
    t('8.32 costWarnings：一則 ITEMS_WITHOUT_COST_LINE；count＝2；文字列出品名與總數、說明「整包可忽略」，不含任何金額', w1.length === 1 && w1[0].code === 'ITEMS_WITHOUT_COST_LINE' && w1[0].count === 2
      && w1[0].message.indexOf('有 2 個有價品項') === 0 && w1[0].message.indexOf('維護服務') > 0 && w1[0].message.indexOf('教育訓練') > 0 && w1[0].message.indexOf('整包') > 0 && !/123456789|\d{4,}/.test(w1[0].message), w1[0] && w1[0].message);
    const many8 = Array.from({ length: 8 }, (_, i) => it8('m' + i, '品項' + i, 5));
    const w2 = CL.costWarnings({ items: many8, costLines: [] })[0];
    t('8.33 超過 5 個品名：只列前 5 個＋「…另 3 個」，總數寫 8', w2.count === 8 && w2.message.indexOf('有 8 個') === 0 && w2.message.indexOf('品項4') > 0 && w2.message.indexOf('品項5') < 0 && w2.message.indexOf('…另 3 個') > 0);
    const w3 = CL.costWarnings({ items: [it8('a', '長'.repeat(80), 5), it8('b', '', 5)], costLines: [] })[0];
    t('8.34 品名超過 30 字截成 30 字加「…」；空白品名顯示「（未命名品項）」', w3.message.indexOf('長'.repeat(30) + '…') > 0 && w3.message.indexOf('長'.repeat(31)) < 0 && w3.message.indexOf('（未命名品項）') > 0);
    t('8.35 沒有未涵蓋品項／舊式單 → costWarnings 回 []', CL.costWarnings({ items: [it8('a', 'A', 5)], costLines: [cl8('A')] }).length === 0 && CL.costWarnings({ items: [it8('a', 'A', 5)] }).length === 0);

    // buildDerived／buildPreview：只在「新式且有未涵蓋品項」時多鍵；其餘輸出與以前完全相同
    {
      const KEYS_OLD = ['rowKey', 'rowLabel', 'level', 'board', 'marginText', 'revenueCents', 'costCents', 'gpCents', 'tiers', 'reasons', 'warnings'];
      const clsx = { P_CONSULT: { cls: 'consult', costBySales: true } };
      const goodNew = { company: 'T', products: ['P_CONSULT'], items: [{ lid: 'i1', desc: '顧問服務', unit: '式', qty: 1, unitPrice: 1000000, cost: 0 }], costLines: norm([L('consult', '顧問服務', 1, 400000)]).lines };
      const d0 = QA.buildDerived(goodNew, clsx);
      t('8.40 buildDerived：新式、品項都有對應列 → 鍵與以前完全相同（沒有 costWarnings）', d0 && eq(Object.keys(d0), KEYS_OLD), d0 && Object.keys(d0).join(','));
      const dL = QA.buildDerived({ company: 'T', products: ['P_CONSULT'], items: [{ lid: 'i1', desc: '顧問服務', unit: '式', qty: 1, unitPrice: 1000000, cost: 400000 }] }, clsx);
      t('8.41 buildDerived：舊式單 → 鍵與以前完全相同', dL && eq(Object.keys(dL), KEYS_OLD));
      const extra = Object.assign({}, goodNew, { items: goodNew.items.concat([{ lid: 'i2', desc: '加購授權', unit: '式', qty: 1, unitPrice: 1500000, cost: 0 }]) });
      const dX = QA.buildDerived(extra, clsx);
      t('8.42 buildDerived：新式且事後多一個有價品項沒成本列 → 多一個 costWarnings（最後一個鍵），其餘欄位照舊計算（毛利率偏高，這正是要提醒的）',
        dX && eq(Object.keys(dX), KEYS_OLD.concat(['costWarnings'])) && dX.costWarnings.length === 1 && dX.costWarnings[0].code === 'ITEMS_WITHOUT_COST_LINE' && dX.costWarnings[0].message.indexOf('加購授權') > 0 && dX.marginText === '84.00', dX && dX.marginText);
      t('8.43 derived 去掉 costWarnings 後與「品項都有對應列」的算法一致（不擋、不改金額）', dX && (() => { const { costWarnings, ...rest } = dX; return eq(Object.keys(rest), KEYS_OLD) && QA.validateForSubmit(extra, clsx).ok === true; })());
      const ctx8 = { cfg: { productClasses: clsx, roster: { gm: [], chairman: [] } }, users: {}, data: {} };
      const pv0 = internal.buildPreview(goodNew, ctx8), pvX = internal.buildPreview(extra, ctx8), pvL = internal.buildPreview({ company: 'T', products: ['P_CONSULT'], items: [{ lid: 'i1', desc: '顧問服務', unit: '式', qty: 1, unitPrice: 1000000, cost: 400000 }] }, ctx8);
      const PV_KEYS = ['rowKey', 'rowLabel', 'level', 'board', 'marginText', 'tiers', 'reasons', 'warnings', 'needsConsultantCost', 'unclassified', 'blockers'];
      t('8.44 buildPreview：沒有未涵蓋品項（新式／舊式）→ 鍵與 warnings 與以前完全相同（warnings 恆為 []）', eq(Object.keys(pv0), PV_KEYS) && eq(pv0.warnings, []) && eq(Object.keys(pvL), PV_KEYS) && eq(pvL.warnings, []));
      t('8.45 buildPreview：有未涵蓋品項 → warnings 多一則（文字＝costWarnings[0].message）、多一個 costWarnings 鍵（最後）；blockers 不變（不擋送簽）',
        eq(Object.keys(pvX), PV_KEYS.concat(['costWarnings'])) && pvX.warnings.length === 1 && pvX.warnings[0] === pvX.costWarnings[0].message && pvX.costWarnings[0].code === 'ITEMS_WITHOUT_COST_LINE'
        && pvX.blockers.every((b) => b.code === 'NO_MANAGER') && !pvX.blockers.some((b) => /成本|COST/.test(b.code)), JSON.stringify(pvX.blockers));
    }

    // 8.50~ backfillForLids
    const bi = [it8('l1', 'A', 5), it8('l2', 'A', 5), it8('l3', 'B', 5), { lid: 'lt', kind: 'title', desc: 'T' }, { lid: 'ls', kind: 'subtotal', desc: 'S' }, it8('l4', '', 5), it8('l5', 'C', 0)];
    const bl = (desc, extra) => Object.assign({ lid: 'x-' + desc + Math.random().toString(36).slice(2, 6), cat: 'hw', desc, vendor: '', note: '', unit: '式', qty: 1, unitCost: 1 }, extra || {});
    let out = CL.backfillForLids(bi, [bl('A'), bl('A'), bl('B'), bl('A'), bl('Z'), bl('C')]);
    t('8.50 同名多個依順序對應：A、A → l1、l2；第 3 個 A 沒有可對應的品項 → 不填；B → l3；對不到的 Z 不填；無價品項 C 也會對應（種子本來就含贈品）',
      out[0].forLid === 'l1' && out[1].forLid === 'l2' && out[2].forLid === 'l3' && !('forLid' in out[3]) && !('forLid' in out[4]) && out[5].forLid === 'l5', JSON.stringify(out.map((l) => l.forLid)));
    out = CL.backfillForLids(bi, [bl('A', { forLid: 'gone' }), bl('A'), bl('B', { forLids: ['l1', 'l2'] }), bl('A')]);
    t('8.51 已有 forLid（即使指向不存在的品項）或 forLids 的列不動；forLids 涵蓋了 l1／l2 之後，其餘同名 A 的列沒有可對應的品項 → 不填',
      out[0].forLid === 'gone' && !('forLid' in out[1]) && !('forLid' in out[3]) && eq(out[2].forLids, ['l1', 'l2']), JSON.stringify(out.map((l) => [l.forLid, l.forLids])));
    out = CL.backfillForLids(bi, [bl('A', { forLid: 'l1' }), bl('A'), bl('A')]);
    t('8.52 已被其他列 forLid 涵蓋的品項（l1）不會再被回填：第 2 個 A → l2、第 3 個 A 不填', !('forLid' in out[2]) && out[1].forLid === 'l2' && out[0].forLid === 'l1');
    out = CL.backfillForLids([{ lid: 'T1', kind: 'title', desc: 'Part' }, { lid: 'S1', kind: 'subtotal', desc: '小計' }], [bl('Part'), bl('小計')]);
    t('8.53 title／subtotal 不參與（同名成本列不回填）', out.every((l) => !('forLid' in l)));
    out = CL.backfillForLids([it8('s', CL.STAMP_DESC, 5), it8('b', 'B', 5)], [{ lid: 'st', cat: 'other', desc: CL.STAMP_DESC, vendor: '', note: '', unit: '式', qty: 1, unitCost: 0, auto: 'stamp' }, bl(CL.STAMP_DESC), bl('B')]);
    t('8.54 印花稅列不參與（自己不回填，也不佔用與它同名的品項：後面同名的一般列仍可對應）', !('forLid' in out[0]) && out[1].forLid === 's' && out[2].forLid === 'b', JSON.stringify(out.map((l) => l.forLid)));
    out = CL.backfillForLids([it8('a', 'A\tB', 5), it8('b', ' C ', 5)], [bl('A B'), bl('C')]);
    t('8.55 品名比對用 trim＋控制字元換空白（與成本列同一套）；空白品名的列／品項不回填', out[0].forLid === 'a' && out[1].forLid === 'b' && !('forLid' in CL.backfillForLids([it8('a', '', 5)], [bl('', { desc: '' })])[0]));
    {
      const lines0 = [bl('A'), bl('Q', { forLid: 'keep' }), bl('B')];
      const snapB = JSON.stringify(lines0), snapI = JSON.stringify(bi);
      const o2 = CL.backfillForLids(bi, lines0);
      t('8.56 回傳新陣列、不改輸入；沒被回填的列維持同一個物件；回填的列是新物件且鍵順序 …unitCost,forLid', o2 !== lines0 && JSON.stringify(lines0) === snapB && JSON.stringify(bi) === snapI && o2[1] === lines0[1] && o2[0] !== lines0[0]
        && eq(Object.keys(o2[0]).slice(-2), ['unitCost', 'forLid']));
      const o3 = CL.backfillForLids(bi, o2);
      t('8.57 冪等：再跑一次結果不變', eq(o3, o2));
    }
    t('8.58 非陣列 lines／items、含 null 的列 → 不丟例外（lines 非陣列回 []）', eq(CL.backfillForLids(bi, null), []) && eq(CL.backfillForLids(undefined, [bl('A')]).map((l) => l.forLid), [undefined]) && (() => { try { CL.backfillForLids([null, 5], [null, 5, 'x', bl('A')]); return true; } catch (e) { return false; } })());
    out = CL.backfillForLids([it8(undefined, 'A', 5), it8('', 'A', 5)], [bl('A'), bl('A')]);
    t('8.59 品項沒有 lid → 以 legacy-<索引> 回填（與 serialize、涵蓋判定一致）', out[0].forLid === 'legacy-0' && out[1].forLid === 'legacy-1');
    // 8.70~ resolveItemRefs：畫面給新品項的暫時代號（nid）→ 真正的 lid
    {
      const mp = new Map([['nid-a', 'LID-A'], ['nid-b', 'LID-B'], ['nid-c', 'LID-C']]);
      const rl = (extra) => Object.assign({ lid: 'z' + Math.random().toString(36).slice(2, 6), cat: 'hw', desc: 'x', vendor: '', note: '', unit: '式', qty: 1, unitCost: 1 }, extra || {});
      const base0 = [rl({ forLid: 'nid-a' }), rl({ forLid: 'nid-a', forLids: ['nid-b', 'nid-c', 'nid-b', 'REAL-1'] }), rl({ forLids: ['nid-c'] }), rl({ forLid: 'nid-zzz', forLids: ['nid-yyy', 'REAL-2'] }), rl(), { lid: 's', cat: 'other', desc: CL.STAMP_DESC, qty: 1, unitCost: 0, auto: 'stamp' }];
      const snapR = JSON.stringify(base0);
      const o = CL.resolveItemRefs(base0, mp);
      t('8.70 forLid／forLids 裡的暫時代號換成 lid：forLid nid-a → LID-A；forLids [nid-b,nid-c,nid-b,REAL-1] → [LID-B,LID-C,REAL-1]（去重、真正的 lid 原樣保留）', o[0].forLid === 'LID-A' && !('forLids' in o[0]) && o[1].forLid === 'LID-A' && eq(o[1].forLids, ['LID-B', 'LID-C', 'REAL-1']), JSON.stringify(o.slice(0, 2)));
      t('8.71 只有 forLids 的列也換；換不到的暫時代號丟掉（不留垃圾）：nid-zzz／nid-yyy 都不見，REAL-2 留下', eq(o[2].forLids, ['LID-C']) && !('forLid' in o[2]) && !('forLid' in o[3]) && eq(o[3].forLids, ['REAL-2']), JSON.stringify(o.slice(2, 4)));
      t('8.72 沒有 forLid／forLids 的列與印花稅列維持同一個物件；不改輸入；鍵順序維持 …unitCost,forLid,forLids', o[4] === base0[4] && o[5] === base0[5] && JSON.stringify(base0) === snapR && eq(Object.keys(o[1]).slice(-2), ['forLid', 'forLids']));
      t('8.73 換成的 lid 若與 forLid 相同，從 forLids 剔除', eq(CL.resolveItemRefs([rl({ forLid: 'nid-a', forLids: ['nid-a', 'LID-A', 'x'] })], mp)[0].forLids, ['x']));
      t('8.74 map 不是 Map／缺：只清掉暫時代號，真正的 lid 保留；lines 不是陣列 → []；列含 null 不丟例外', eq(CL.resolveItemRefs([rl({ forLid: 'nid-a', forLids: ['K'] })], null)[0].forLids, ['K']) && !('forLid' in CL.resolveItemRefs([rl({ forLid: 'nid-a' })], undefined)[0])
        && eq(CL.resolveItemRefs('x', mp), []) && CL.resolveItemRefs([null, 5], mp).length === 2);
      t('8.75 暫時代號前綴是 "nid-"（匯出 TMP_REF_PREFIX），伺服器產生的 lid（uuid）不會以它開頭', CL.TMP_REF_PREFIX === 'nid-' && !/^nid-/.test(require('crypto').randomUUID()));
      // 先換暫時代號、再用品名補：新單「合併為一列」(整包) 與「改名」在第一次儲存後都算有對應
      const itemsNew = [it8('LID-A', '顧問人天', 5), it8('LID-B', '軟體授權', 5), it8('LID-C', '硬體設備', 5)];
      const mergedNew = [rl({ desc: '專案整包成本', forLid: 'nid-a', forLids: ['nid-b', 'nid-c'] })];
      const afterSave = CL.backfillForLids(itemsNew, CL.resolveItemRefs(mergedNew, mp));
      t('8.76 新單整包情境：合併後的列（名稱對不上任何品項）帶 nid → 存檔換成 lid → 3 個有價品項都有對應；不換的話 3 個都被提醒', CL.unmatchedItems({ items: itemsNew, costLines: afterSave }).length === 0
        && CL.unmatchedItems({ items: itemsNew, costLines: CL.backfillForLids(itemsNew, mergedNew) }).length === 3);
      t('8.77 新單改名情境：成本列用 nid 指向品項，儲存前把品項改名也算有對應（只靠品名回填做不到）', (() => {
        const renamed = [it8('LID-A', '改過的名字', 5)];
        const seeded = [rl({ desc: '原本的名字', forLid: 'nid-a' })];
        return CL.unmatchedItems({ items: renamed, costLines: CL.backfillForLids(renamed, CL.resolveItemRefs(seeded, mp)) }).length === 0
          && CL.unmatchedItems({ items: renamed, costLines: CL.backfillForLids(renamed, [rl({ desc: '原本的名字' })]) }).length === 1;
      })());
    }
    // backfill 之後 unmatchedItems 認得：新單第一次儲存後品項改名仍算有對應
    {
      const itemsAfter = [it8('n1', 'A', 100), it8('n2', 'B', 50)];
      const filled = CL.backfillForLids(itemsAfter, [bl('A'), bl('B')]);
      const renamed = [it8('n1', 'A改名', 100), it8('n2', 'B', 50)];
      t('8.60 回填後品項改名仍算有對應（沒回填則會被當成新品項）', CL.unmatchedItems({ items: renamed, costLines: filled }).length === 0 && CL.unmatchedItems({ items: renamed, costLines: [bl('A'), bl('B')] }).length === 1);
    }
  }


  let pass = 0, failN = 0;
  res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (x && !ok ? '  ← ' + x : '')); ok ? pass++ : failN++; });
  console.log('\n成本明細單元測試：PASS ' + pass + ' / FAIL ' + failN);
  process.exit(failN ? 1 : 0);
}

if (require.main === module) run();
module.exports = { genLegacyQuotes, lcg };
