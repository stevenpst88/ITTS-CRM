#!/usr/bin/env node
/**
 * 報價單 Remarks（付款方式／追加條款）一致性檢查。用法：node scripts/check-quote-remarks.js
 * 動 lib/quoteRemarks.js、_client/quote.js 的 quotePaymentSentence、templates/quotation_template.xlsx 之後必跑。
 *   1) 前端 quotePaymentSentence 與伺服器 paymentSentence 逐例輸出相同（兩邊是鏡像，沒有任何機制強制一致）
 *   2) 範本第 84/85/87/88/89 列的固定條文與 REM.FIXED 相同（改範本沒改程式、或反過來都會抓到）
 *   3) 驗證規則（嚴格型別、比例合計 100、零寬字元、舊單行備註免長度限制）
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ROOT = path.join(__dirname, '..');
const R = require(path.join(ROOT, 'lib/quoteRemarks.js'));
const QE = require(path.join(ROOT, 'lib/quoteExcel.js'));
const JSZip = require('jszip');

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);

(async () => {
  // 1) 前後端句子一致
  // 換行一律先正規化成 LF：Windows 以 autocrlf=true clone 下來的檔案是 CRLF，下面用 '\n}\n' 切函式會找不到
  const src = fs.readFileSync(path.join(ROOT, '_client/quote.js'), 'utf8').replace(/\r\n/g, '\n');
  const i0 = src.indexOf('function quotePaymentSentence(');
  const i1 = src.indexOf('\n}\n', i0) + 3;
  if (i0 < 0 || i1 < 3) throw new Error('找不到 quotePaymentSentence');
  const ctx = {}; vm.createContext(ctx); vm.runInContext(src.slice(i0, i1), ctx);
  const cases = [
    undefined, null, {}, { items: [] }, { net: 30, items: [null] }, { net: 30, items: [{ label: 'x' }] },
    R.DEFAULT_PAYMENT, ...R.PAYMENT_PRESETS,
    { net: 0, items: R.PAYMENT_PRESETS[3].items },
    { net: 90, items: [{ label: 'a', pct: 33.3 }, { label: 'b', pct: 33.3 }, { label: 'c', pct: 33.4 }] },
    { net: 45, items: [{ label: '簽約後', pct: 99.9 }, { label: '驗收後', pct: 0.1 }] },
    { net: 60, items: [{ label: '<b>&"x"', pct: 100 }] },
  ];
  const bad = cases.filter((c) => ctx.quotePaymentSentence(c) !== R.paymentSentence(c));
  t(`1. 前端 quotePaymentSentence 與伺服器 paymentSentence 輸出相同（${cases.length} 例）`, bad.length === 0, bad.length ? JSON.stringify(bad[0]) : '');
  t('1b. 資料被改壞（items 含 null／缺欄位）→ 退回預設句、不丟例外', R.paymentSentence({ net: 30, items: [null] }) === R.paymentSentence(undefined) && R.paymentSentence({ net: 30, items: [{ label: 'x' }] }) === R.paymentSentence(undefined));

  // 2) 範本固定條文
  const z = await JSZip.loadAsync(fs.readFileSync(path.join(ROOT, 'templates/quotation_template.xlsx')));
  const sst = QE._internal.parseSharedStrings(await z.file('xl/sharedStrings.xml').async('string'));
  const sh = new QE._internal.SheetXml(await z.file('xl/worksheets/sheet1.xml').async('string'), sst);
  const cell = (r) => sh.cellText('B' + r).trim();
  t('2. 範本 B84/B85/B88/B89＝FIXED 第 1/2/5/6 條', cell(84) === R.FIXED[1] && cell(85) === R.FIXED[2] && cell(88) === R.FIXED[5] && cell(89) === R.FIXED[6]);
  t('2b. 範本 B86＝預設付款句、B87＝第 4 條佔位句', cell(86) === R.paymentSentence(undefined) && cell(87) === R.FIXED[4]);

  // 3) 驗證規則
  const ok = (p) => !!R.normalizePayment(p).value;
  t('3a. 字串數字可、十六進位／科學記號不可（0x1E、1e2）', ok({ net: '45', items: [{ label: 'a', pct: '100' }] }) && !ok({ net: '0x1E', items: [{ label: 'a', pct: 100 }] }) && !ok({ net: 30, items: [{ label: 'a', pct: '1e2' }] }) && !ok({ net: 30, items: [{ label: 'a', pct: '0x64' }] }));
  t('3b. 比例合計必須剛好 100（0.1+99.9 可、33.3×3 不可）', ok({ net: 30, items: [{ label: 'a', pct: 0.1 }, { label: 'b', pct: 99.9 }] }) && !ok({ net: 30, items: [{ label: 'a', pct: 33.3 }, { label: 'b', pct: 33.3 }, { label: 'c', pct: 33.3 }] }));
  const ZW = String.fromCharCode(0x200B);
  t('3c. 只有零寬字元的付款時點／條款視為空白', !ok({ net: 30, items: [{ label: ZW, pct: 100 }] }) && R.normalizeClauses([ZW, 'a']).value.length === 1);
  const long250 = '字'.repeat(250);
  t('3d. 追加條款單條 >200 字拒絕；lenientLength 僅供比對舊單行備註時使用', !!R.normalizeClauses([long250]).error && R.normalizeClauses([long250], { lenientLength: true }).value[0].length === 250);
  t('3e. 型別嚴格（陣列／布林／物件當比例或時點）', [{ net: 30, items: [{ label: 'a', pct: [] }] }, { net: 30, items: [{ label: 'a', pct: true }] }, { net: 30, items: [{ label: ['x'], pct: 100 }] }, null, 5].every((x) => !ok(x)));

  // 4) 品項說明／備註（spec／note）不影響 Remarks 區與品項區以下的版面（完整檢查見 scripts/check-quote-item-notes.js）
  {
    const base = { quoteNo: 'QU-R', company: 'C', projectName: 'P', validUntil: '2026-10-30', discountType: 'none', discountValue: 0, payment: R.DEFAULT_PAYMENT, extraClauses: ['追加條款一', '追加條款二'], items: [{ lid: 'a', desc: '品項', unit: '式', qty: 1, unitPrice: 100 }, { lid: 'b', desc: '品項2', unit: '式', qty: 1, unitPrice: 200 }] };
    const withN = JSON.parse(JSON.stringify(base)); withN.items[0].spec = '說明'; withN.items[0].note = '備註'; withN.items[1].note = '=1+1';
    const sheetOf = async (q) => (await JSZip.loadAsync(await QE.buildQuoteWorkbook(q, path.join(ROOT, 'templates/quotation_template.xlsx'), { issueDate: '2026-10-09', issuer: {} }))).file('xl/worksheets/sheet1.xml').async('string');
    const [a, b] = [await sheetOf(base), await sheetOf(withN)];
    const tail = (x) => x.slice(x.indexOf('<row r="' + (QE.LAYOUT.itemLast + 1) + '"'));
    t('4. 品項區以下（總額、專案資料框、Remarks 第 1~6 條與追加條款、簽名區）的 XML 與沒有說明／備註時逐位元相同', tail(a).length > 5000 && tail(a) === tail(b));
  }

  let pass = 0, fail = 0;
  res.forEach(([n, o, x]) => { console.log((o ? 'PASS ' : 'FAIL ') + n + (x && !o ? '  ← ' + x : '')); o ? pass++ : fail++; });
  console.log(`\nRemarks 一致性檢查：PASS ${pass} / FAIL ${fail}`);
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error('檢查腳本錯誤', e.stack); process.exit(2); });
