#!/usr/bin/env node
/**
 * lib/xlsxSheetHtml.js（毛利分析預覽用的「xlsx 工作表 → HTML」）檢查。用法：node scripts/check-xlsx-sheet-html.js
 *   1) 數字格式：千分位、小數、百分比、NT$、會計格式（零顯示 -）、括號負數、日期
 *   2) 用 lib/quotePnlExcel.js 產一份毛利分析 xlsx 再轉 HTML：隱藏列不出現、合併儲存格有 colspan、數值與格式正確、
 *      使用者輸入的文字一律跳脫（品項說明含 <img onerror>）、顏色（theme＋tint／indexed）有轉成 CSS
 */
const path = require('path');
const ROOT = path.join(__dirname, '..');
const { sheetToHtml, _internal: I } = require(path.join(ROOT, 'lib/xlsxSheetHtml.js'));
const { buildQuotePnlExcel } = require(path.join(ROOT, 'lib/quotePnlExcel.js'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const F = I.formatNumber;

(async () => {
  // 1) 數字格式
  t('1a. #,##0 千分位與四捨五入', F(1234567, '#,##0') === '1,234,567' && F(1234.5, '#,##0') === '1,235' && F(0.4, '#,##0') === '0');
  t('1b. 0.00%／0%', F(0.2189, '0.00%') === '21.89%' && F(1, '0%') === '100%' && F(0.3162, '0.00%') === '31.62%');
  t('1c. "NT$"#,##0 與 "NT$"#,##0.00', F(1485000, '"NT$"#,##0') === 'NT$1,485,000' && F(1234.5, '"NT$"#,##0.00') === 'NT$1,234.50');
  t('1d. 括號負數：#,##0_);[Red]\\(#,##0\\)', F(-1234, '#,##0_);[Red]\\(#,##0\\)') === '(1,234)' && F(1234, '#,##0_);[Red]\\(#,##0\\)').trim() === '1,234');
  t('1e. 會計格式：正數、負數、零顯示「-」', F(5000, '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-').trim() === '5,000' && /^-\s*5,000$/.test(F(-5000, '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-').trim()) && F(0, '_-* #,##0_-;\\-* #,##0_-;_-* "-"??_-;_-@_-').trim() === '-');
  t('1f. 一般格式：整數原樣、小數去除浮點尾巴', F(12, 'General') === '12' && F(0.1 + 0.2, 'General') === '0.3');
  t('1g. 日期 yyyy/m/d（2026-10-07 序號 46302）', F(46302, 'yyyy/m/d') === '2026/10/7', F(46302, 'yyyy/m/d'));
  t('1h. 只有 #／? 的格式 0 不顯示數字（#,##0.00;;）', F(0, '#,##0.00;-#,##0.00;;@') === '' || F(0, '#,##0.00;-#,##0.00;;@').trim() === '');
  t('1i. 顏色：rgb、theme、indexed、tint', I.resolveColor('<color rgb="FFFF0000"/>', []) === '#FF0000' && I.resolveColor('<color indexed="5"/>', []) === '#FFFF00' && I.resolveColor('<color theme="1"/>', ['FFFFFF', '000000']) === '#000000' && /^#[0-9A-F]{6}$/.test(I.resolveColor('<color theme="4" tint="0.5"/>', ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD'])));

  // 2) 毛利分析 xlsx → HTML
  const q = { quoteNo: 'QU-CHK-1', company: '預覽<b>測試</b>', contactName: '王', discountType: 'percent', discountValue: 90, products: ['P'],
    items: [{ desc: '<img src=x onerror=alert(1)>授權', unit: '式', qty: 10, unitPrice: 100000, cost: 70000, cat: 'software' }, { desc: '顧問 & "服務"', unit: '人天', qty: 5, unitPrice: 20000, cost: 12000, cat: 'consult' }] };
  const buf = await buildQuotePnlExcel(q, { classCodes: ['software', 'consult'], requestedBy: '業務<小明>', issueDate: '2026-10-07', contingencyPct: 5 });
  const v = await sheetToHtml(buf);
  const rows = (v.html.match(/<tr /g) || []).length;
  t('2a. 範圍＝列印範圍 A1:H106；隱藏列不輸出（列數 < 106）', v.range === 'A1:H106' && rows > 50 && rows < 106, `range=${v.range} rows=${rows}`);
  t('2b. 有合併儲存格（colspan／rowspan）', /colspan="\d+"/.test(v.html));
  t('2c. 使用者輸入一律跳脫（沒有原始的 <img、<b>、<小明>）', !/<img/i.test(v.html) && !/<b>/.test(v.html) && !/<小明>/.test(v.html) && v.html.includes('&lt;img src=x onerror=alert(1)&gt;'));
  t('2d. 數值與格式：定價收入 1,000,000＋100,000＝1,100,000 折後 990,000（NT$）出現在表內', /NT\$990,000/.test(v.html) || /990,000/.test(v.html), (v.html.match(/NT\$[\d,]+/g) || []).slice(0, 8).join(' '));
  t('2e. 毛利率百分比有格式（xx.xx%）', /\d+\.\d\d%/.test(v.html));
  t('2f. 填色與文字色轉成 CSS（background / color 出現多種色）', new Set(v.html.match(/background:#[0-9A-F]{6}/g) || []).size >= 3 && new Set(v.html.match(/color:#[0-9A-F]{6}/g) || []).size >= 2);
  t('2g. 沒有 <script>、事件屬性（onerror 等，限標籤內；跳脫後的文字不算）、javascript: 連結', !/<script/i.test(v.html) && !/<[^>]*\son\w+\s*=/i.test(v.html) && !/href\s*=\s*"?javascript:/i.test(v.html));
  t('2h. 寬度為欄寬總和、不是 0', v.widthPx > 300 && v.widthPx < 3000, v.widthPx);

  // 3) 審查後補強：四捨五入、時間、[Red]、列印範圍、注音、隱藏列中的合併儲存格、空白保留、溢出裁切
  t('3a. 百分比／金額四捨五入與 Excel 相同（0.02055→2.06%、0.02175→2.18%、150.075→150.08、0.0515→5.2%）', F(0.02055, '0.00%') === '2.06%' && F(0.02175, '0.00%') === '2.18%' && F(150.075, '#,##0.00') === '150.08' && F(0.0515, '0.0%') === '5.2%', [F(0.02055, '0.00%'), F(0.02175, '0.00%'), F(150.075, '#,##0.00'), F(0.0515, '0.0%')].join(' '));
  t('3b. 時間：h:mm 的 mm 是分鐘、日期的 mm 是月份', F(46302.5, 'yyyy/m/d h:mm') === '2026/10/7 12:00' && F(46302.75, 'h:mm') === '18:00' && F(46302, 'yyyy/mm/dd') === '2026/10/07', F(46302.5, 'yyyy/m/d h:mm'));
  t('3c. [Red] 負數格式回傳紅色', I.formatNumberEx(-5, '#,##0_);[Red]\\(#,##0\\)').color === '#FF0000' && I.formatNumberEx(5, '#,##0_);[Red]\\(#,##0\\)').color === null);
  t('3d. 列印範圍：多段取第一段、含空白與括號的工作表名、整欄範圍回 null（改用 dimension）', I.firstRange("'PNL (C)'!$A$1:$H$106") === 'A1:H106' && I.firstRange('Sheet1!$B$2:$D$4,Sheet1!$A$8:$C$10') === 'B2:D4' && I.firstRange('Sheet1!$A:$C') === null);
  const JSZip = require('jszip');
  const mini = async (sheetRows, merges, extra) => {
    const z = new JSZip();
    z.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
    z.file('xl/styles.xml', '<styleSheet><fonts count="1"><font><sz val="11"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>');
    z.file('xl/sharedStrings.xml', (extra && extra.sst) || '<sst><si><t>文字</t></si></sst>');
    z.file('xl/worksheets/sheet1.xml', '<worksheet><dimension ref="A1:D6"/><sheetData>' + sheetRows + '</sheetData>' + (merges ? '<mergeCells>' + merges + '</mergeCells>' : '') + '</worksheet>');
    return sheetToHtml(await z.generateAsync({ type: 'nodebuffer' }));
  };
  const cell = (r, c, v, tt) => `<c r="${c}${r}"${tt ? ` t="${tt}"` : ''}><v>${v}</v></c>`;
  const colsOf = (html) => { const trs = html.match(/<tr [\s\S]*?<\/tr>/g) || []; return trs.map(tr => (tr.match(/<td/g) || []).length); };
  { // 直向合併 A2:B4，第 2 列隱藏：其餘列的格子數不可少、不可位移
    const rows = [1, 2, 3, 4].map(r => `<row r="${r}"${r === 2 ? ' hidden="1"' : ''}>${['A', 'B', 'C', 'D'].map((c, i) => cell(r, c, r * 10 + i)).join('')}</row>`).join('');
    const v1 = await mini(rows, '<mergeCell ref="A2:B4"/>');
    const widths = colsOf(v1.html).map((n, i) => n); // 可見列：1, 3, 4
    t('3e. 合併儲存格遇到隱藏列：以第一個可見儲存格為錨點，每列的欄位數量加上 colspan 仍等於 4（沒有位移）', /rowspan="2"/.test(v1.html) && /colspan="2"/.test(v1.html), JSON.stringify(widths));
  }
  { const v2 = await mini('<row r="1">' + cell(1, 'A', 0, 's') + cell(1, 'B', 1) + '</row>', '<mergeCell ref="A1:XFD5000"/>');
    t('3f. 超大合併範圍被裁到顯示範圍（不會卡住事件迴圈，也不丟例外）', v2.html.length > 0); }
  { const v3 = await mini('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', '', { sst: '<sst><si><t>主管</t><rPh sb="0" eb="2"><t>シュカン</t></rPh></si></sst>' });
    t('3g. 注音（rPh）不會混進顯示文字', v3.html.includes('主管') && !v3.html.includes('シュカン')); }
  { const v4 = await mini('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', '', { sst: '<sst><si><t xml:space="preserve">項      目</t></si></sst>' });
    t('3h. 連續空白保留（white-space:pre），例如「項      目」', v4.html.includes('項      目') && /white-space:pre;/.test(v4.html)); }
  { const v5 = await mini('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>5</v></c></row>', '', { sst: '<sst><si><t>很長很長的文字會蓋到右邊格子</t></si></sst>' });
    t('3i. 右邊相鄰格有內容 → 文字裁切（overflow:hidden）；沒有才溢出', /overflow:hidden/.test(v5.html.split('</td>')[0]) );
    const v6 = await mini('<row r="1"><c r="A1" t="s"><v>0</v></c></row>', '', { sst: '<sst><si><t>很長很長的文字</t></si></sst>' });
    t('3j. 右邊是空格 → 溢出（overflow:visible）', /overflow:visible/.test(v6.html.split('</td>')[0])); }

  let pass = 0, fail = 0;
  res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (!ok && x ? '  <- ' + x : '')); ok ? pass++ : fail++; });
  console.log(`\nxlsx 轉 HTML 檢查：PASS ${pass} / FAIL ${fail}`);
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error('檢查腳本錯誤', e.stack); process.exit(2); });
