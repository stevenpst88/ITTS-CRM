#!/usr/bin/env node
/**
 * 報價單「分組標題／小計列」一致性檢查。用法：node scripts/check-quote-items.js
 * 動 lib/quoteItems.js、_client/quote-preview.js 的 _qpvSubtotals、lib/quoteExcel.js 的品項迴圈、templates/quotation_template.xlsx 之後必跑。
 *   1) lib/quoteItems.js 的語意（種類判斷、小計分段、標籤）
 *   2) 前端預覽的 _qpvSubtotals 與伺服器 subtotalValues 逐例輸出相同（兩邊是鏡像）
 *   3) 範本 styles.xml 有分組標題樣式（LAYOUT.titleStyle），且 Excel 輸出：標題列合併 B:J、小計列合併 C:I、
 *      小計公式與總額公式不會重複加總（有小計列才用 SUBTOTAL，舊單維持 SUM）、項目編號只算一般品項
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ROOT = path.join(__dirname, '..');
const QI = require(path.join(ROOT, 'lib/quoteItems.js'));
const QE = require(path.join(ROOT, 'lib/quoteExcel.js'));
const JSZip = require('jszip');

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const it = (d, q, p) => ({ lid: 'i' + d, desc: d, unit: '式', qty: q, unitPrice: p });
const ti = (d) => ({ lid: 't' + d, kind: 'title', desc: d });
const su = (d) => ({ lid: 's' + d, kind: 'subtotal', desc: d });

(async () => {
  // 1) helper 語意
  t('1a. 種類判斷：無 kind／未知 kind／null 都當一般品項', QI.isItemRow({}) && QI.isItemRow({ kind: 'item' }) && QI.isItemRow({ kind: 'weird' }) && QI.isItemRow(null) && QI.isNonItemRow({ kind: 'title' }) && QI.isNonItemRow({ kind: 'subtotal' }) && QI.normalizeRowKind('x') === 'item');
  const rows = [ti('A'), it('a1', 2, 100), it('a2', 1, 50), su('A 小計'), ti('B'), it('b1', 3, 10), su(''), it('c1', 1, 7), su('')];
  const sv = QI.subtotalValues(rows);
  t('1b. 小計＝上一個標題／小計之後的品項合計；其餘為 null', JSON.stringify(sv) === JSON.stringify([null, null, null, 250, null, null, 30, null, 7]), JSON.stringify(sv));
  t('1c. 連續小計、空段小計＝0；標題之後立刻小計＝0', JSON.stringify(QI.subtotalValues([su('x'), su('y'), ti('t'), su('z')])) === JSON.stringify([0, 0, null, 0]));
  t('1d. 數量空／0 當 1（與 Excel、預覽一致）、單價空當 0', QI.lineAmount({ qty: '', unitPrice: 100 }) === 100 && QI.lineAmount({ qty: 0, unitPrice: 5 }) === 5 && QI.lineAmount({ qty: 3 }) === 0);
  t('1e. 小計標籤空白 → 「小計」', QI.subtotalLabel({ desc: '  ' }) === '小計' && QI.subtotalLabel({ desc: ' Part A 小計 ' }) === 'Part A 小計');
  t('1f. itemRows 只留一般品項、不修改原陣列', QI.itemRows(rows).length === 4 && rows.length === 9);

  // 2) 前端預覽鏡像
  const src = fs.readFileSync(path.join(ROOT, '_client/quote-preview.js'), 'utf8').replace(/\r\n/g, '\n');
  const a = src.indexOf('function _qpvIsKindRow('), b = src.indexOf('/** 金額與優惠');
  if (a < 0 || b < 0) throw new Error('找不到 _qpvSubtotals');
  const ctx = {}; vm.createContext(ctx); vm.runInContext(src.slice(a, b), ctx);
  const cases = [[], rows, [su('x'), su('y'), ti('t'), su('z')], [it('a', '', 100), su(''), it('b', 0, 5), su('')], [ti('only')], [it('x', 2.5, 100.5), it('y', '3', '7.25'), su('s')], [it('n', 1, 1), null, su('')].filter(Boolean)];
  const bad = cases.filter(c => JSON.stringify(ctx._qpvSubtotals(c)) !== JSON.stringify(QI.subtotalValues(c)));
  t(`2. 前端 _qpvSubtotals 與伺服器 subtotalValues 輸出相同（${cases.length} 例）`, bad.length === 0, bad.length ? JSON.stringify(bad[0]) : '');

  // 3) 範本樣式與 Excel 輸出
  const TEMPLATE = path.join(ROOT, 'templates/quotation_template.xlsx');
  const tz = await JSZip.loadAsync(fs.readFileSync(TEMPLATE));
  const st = await tz.file('xl/styles.xml').async('string');
  const nXf = +(/<cellXfs count="(\d+)"/.exec(st) || [])[1], nFill = +(/<fills count="(\d+)"/.exec(st) || [])[1];
  t('3a. 範本 styles.xml 有分組標題樣式（cellXfs ≥ 104、fills ≥ 7，LAYOUT.titleStyle 在範圍內）', nXf >= 104 && nFill >= 7 && QE.LAYOUT.titleStyle < nXf, `xfs=${nXf} fills=${nFill} titleStyle=${QE.LAYOUT.titleStyle}`);
  const base = { quoteNo: 'QU-X', company: 'C', projectName: 'P', validUntil: '2026-10-30', discountType: 'none', discountValue: 0 };
  const build = async (items) => {
    const buf = await QE.buildQuoteWorkbook({ ...base, items }, TEMPLATE, { issueDate: '2026-10-07', issuer: {}, approved: false, seal: null });
    const z = await JSZip.loadAsync(buf); return z.file('xl/worksheets/sheet1.xml').async('string');
  };
  const L = QE.LAYOUT;
  const x1 = await build([ti('Part A'), it('a1', 2, 100), it('a2', 1, 50), su('Part A 小計'), ti('Part B'), it('b1', 3, 10), su('')]);
  const cell = (xml, ref) => { const m = new RegExp('<c r="' + ref + '"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)').exec(xml); return m ? { attrs: m[1], inner: m[2] || '' } : null; };
  const r = (i) => L.itemFirst + i;
  t('3b. 標題列：B:J 合併、B 用 titleStyle', x1.includes(`<mergeCell ref="B${r(0)}:J${r(0)}"/>`) && new RegExp(`<c r="B${r(0)}"[^>]*s="${L.titleStyle}"`).test(x1) && !x1.includes(`<mergeCell ref="C${r(0)}:E${r(0)}"/>`));
  t('3c. 小計列：C:I 合併、J 為 SUBTOTAL 公式且快取值正確（250／30）', x1.includes(`<mergeCell ref="C${r(3)}:I${r(3)}"/>`) && cell(x1, 'J' + r(3)).inner.includes(`SUBTOTAL(9,J${r(1)}:J${r(2)})`) && cell(x1, 'J' + r(3)).inner.includes('<v>250</v>') && cell(x1, 'J' + r(6)).inner.includes(`SUBTOTAL(9,J${r(5)}:J${r(5)})`) && cell(x1, 'J' + r(6)).inner.includes('<v>30</v>'));
  t('3d. 有小計列：總額用 SUBTOTAL(9,…)（避免把小計重複加總）；快取值＝品項合計 280', cell(x1, 'J' + L.sumList).inner.includes(`SUBTOTAL(9,J${L.itemFirst}:J${L.itemLast})`) && cell(x1, 'J' + L.sumList).inner.includes('<v>280</v>'), cell(x1, 'J' + L.sumList).inner);
  t('3e. 項目編號只算一般品項（1、2、3），標題／小計列不編號', cell(x1, 'B' + r(1)).inner.includes('<v>1</v>') && cell(x1, 'B' + r(2)).inner.includes('<v>2</v>') && cell(x1, 'B' + r(5)).inner.includes('<v>3</v>') && !cell(x1, 'B' + r(3)).inner.includes('<v>') && !cell(x1, 'B' + r(4)).inner.includes('<v>'));
  const x0 = await build([it('a1', 2, 100), it('a2', 1, 50)]);
  t('3f. 沒有小計列（含所有舊單）：總額維持 SUM、合併儲存格維持範本原樣', cell(x0, 'J' + L.sumList).inner.includes(`<f>SUM(J${L.itemFirst}:J${L.itemLast})</f>`) && x0.includes(`<mergeCell ref="C${r(0)}:E${r(0)}"/>`) && !/SUBTOTAL/.test(x0));
  const xt = await build([ti('只有標題'), it('a', 1, 1)]);
  t('3g. 只有標題沒有小計列：總額仍用 SUM（沒有小計值會重複加總）', cell(xt, 'J' + L.sumList).inner.includes('<f>SUM('));
  const xe = await build([su(''), ti('T'), su('空段')]);
  t('3h. 空段小計寫 0、不產生反向範圍', cell(xe, 'J' + r(0)).inner.includes('<v>0</v>') && !cell(xe, 'J' + r(0)).inner.includes('<f>') && cell(xe, 'J' + r(2)).inner.includes('<v>0</v>') && !/SUBTOTAL\(9,J\d+:J\d+\)/.test(cell(xe, 'J' + r(2)).inner));

  // 4) 品項說明／備註（spec／note，印在客戶單上）不影響分組標題／小計列的版面與計算（完整檢查見 scripts/check-quote-item-notes.js）
  const noted = [ti('Part A'), Object.assign(it('a1', 2, 100), { spec: '說明 a1', note: '備註 a1' }), Object.assign(it('a2', 1, 50), { spec: '說明 a2' }), su('Part A 小計'), ti('Part B'), Object.assign(it('b1', 3, 10), { note: '=1+1' }), su('')];
  const x4 = await build(noted);
  t('4a. 品項帶說明／備註：標題列仍合併 B:J、小計列仍合併 C:I、小計公式與快取值（250／30）、總額 SUBTOTAL 與快取值（280）、項目編號（1、2、3）都不變', x4.includes(`<mergeCell ref="B${r(0)}:J${r(0)}"/>`) && x4.includes(`<mergeCell ref="C${r(3)}:I${r(3)}"/>`)
    && cell(x4, 'J' + r(3)).inner.includes(`SUBTOTAL(9,J${r(1)}:J${r(2)})`) && cell(x4, 'J' + r(3)).inner.includes('<v>250</v>') && cell(x4, 'J' + r(6)).inner.includes('<v>30</v>')
    && cell(x4, 'J' + L.sumList).inner.includes(`SUBTOTAL(9,J${L.itemFirst}:J${L.itemLast})`) && cell(x4, 'J' + L.sumList).inner.includes('<v>280</v>')
    && cell(x4, 'B' + r(1)).inner.includes('<v>1</v>') && cell(x4, 'B' + r(2)).inner.includes('<v>2</v>') && cell(x4, 'B' + r(5)).inner.includes('<v>3</v>'));
  t('4b. 標題／小計列帶說明／備註（不該有）→ 與沒帶時的 sheet1.xml 完全相同；前端 _qpvSubtotals 不看說明／備註', await (async () => { const a = [Object.assign(ti('T'), { spec: 'x', note: 'y' }), it('a', 1, 5), Object.assign(su('S'), { spec: 'x', note: 'y' })]; const b = [ti('T'), it('a', 1, 5), su('S')]; return (await build(a)) === (await build(b)); })()
    && JSON.stringify(ctx._qpvSubtotals(noted)) === JSON.stringify(ctx._qpvSubtotals(noted.map((x) => { const y = Object.assign({}, x); delete y.spec; delete y.note; return y; }))) && JSON.stringify(QI.subtotalValues(noted)) === JSON.stringify([null, null, null, 250, null, null, 30]));

  let pass = 0, fail = 0;
  res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (!ok && x ? '  <- ' + x : '')); ok ? pass++ : fail++; });
  console.log(`\n分組標題／小計列檢查：PASS ${pass} / FAIL ${fail}`);
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error('檢查腳本錯誤', e.stack); process.exit(2); });
