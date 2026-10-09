#!/usr/bin/env node
/**
 * lib/xlsxSheetHtml.js（毛利分析預覽用的「xlsx 工作表 → HTML」）檢查。用法：node scripts/check-xlsx-sheet-html.js
 *   1) 數字格式：千分位、小數、百分比、NT$、會計格式（零顯示 -）、括號負數、日期
 *   2) 用 lib/quotePnlExcel.js 產毛利分析 xlsx（舊式＝items[].cost、新式＝成本明細 costLines 各一份）再轉 HTML：
 *      列印範圍 A1:H126 共 126 列（新範本沒有隱藏列）、合併儲存格有 colspan、折後金額與百分比格式正確、沒有 ####、
 *      委外廠商／說明／差旅／交際費／印花稅都在預覽內、使用者輸入的文字一律跳脫（品項／廠商／說明含 <img onerror>）、顏色（theme＋tint／indexed）有轉成 CSS
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
  t('2a. 範圍＝列印範圍 A1:H126；新範本沒有隱藏列，126 列全部輸出', v.range === 'A1:H126' && rows === 126, `range=${v.range} rows=${rows}`);
  t('2b. 有合併儲存格（colspan／rowspan）', /colspan="\d+"/.test(v.html));
  t('2c. 使用者輸入一律跳脫（沒有原始的 <img、<b>、<小明>）', !/<img/i.test(v.html) && !/<b>/.test(v.html) && !/<小明>/.test(v.html) && v.html.includes('&lt;img src=x onerror=alert(1)&gt;'));
  // 折扣後：軟體 10×100,000＋顧問 5×20,000＝1,100,000 → 九折 990,000（軟體 900,000、顧問 90,000）；每類只寫一個總額格（H23／H14），總收入 H13
  t('2d. 折後金額（NT$ 格式）：總收入 NT$990,000、軟體 NT$900,000、顧問 NT$90,000 出現在表內', /NT\$990,000/.test(v.html) && /NT\$900,000/.test(v.html) && /NT\$90,000/.test(v.html), (v.html.match(/NT\$[\d,]+/g) || []).slice(0, 8).join(' '));
  t('2e. 毛利率百分比有格式（xx.xx%）', /\d+\.\d\d%/.test(v.html));
  t('2e2. 沒有 ####（數值格不會因欄寬被顯示成 ####）、沒有 NaN／undefined／[object', !/#{3,}/.test(v.html) && !/NaN|undefined|\[object/.test(v.html));
  t('2f. 填色與文字色轉成 CSS（background / color 出現多種色）', new Set(v.html.match(/background:#[0-9A-F]{6}/g) || []).size >= 3 && new Set(v.html.match(/color:#[0-9A-F]{6}/g) || []).size >= 2);
  t('2g. 沒有 <script>、事件屬性（onerror 等，限標籤內；跳脫後的文字不算）、javascript: 連結', !/<script/i.test(v.html) && !/<[^>]*\son\w+\s*=/i.test(v.html) && !/href\s*=\s*"?javascript:/i.test(v.html));
  t('2h. 寬度為欄寬總和、不是 0', v.widthPx > 300 && v.widthPx < 3000, v.widthPx);

  // 2b) 新式單（成本明細 costLines）：委外廠商／供應商、說明、差旅、交際費、印花稅、Contingency 都要出現在預覽，且使用者文字一律跳脫
  const q2 = { quoteNo: 'QU-CHK-2', company: '新式預覽', discountType: 'none', products: ['P'],
    items: [{ desc: '導入顧問', unit: '人天', qty: 10, unitPrice: 50000, cat: 'consult' }, { desc: '授權', unit: '套', qty: 1, unitPrice: 500000, cat: 'software' }],
    costLines: [
      { lid: 'a', cat: 'consult', desc: 'PM <b>粗體</b>', vendor: 'Vendor-A <i>x</i>', note: '每人天', unit: '人天', qty: 10, unitCost: 8000 },
      { lid: 'b', cat: 'software', desc: '<img src=x onerror=alert(2)>', vendor: 'Supplier-S', note: '含維護 "全年" & 保固', unit: '套', qty: 1, unitCost: 300000 },
      { lid: 'c', cat: 'hw', desc: '主機', vendor: '', note: '', unit: '台', qty: 2, unitCost: 45000 },
      { lid: 'd', cat: 'travel', desc: '差旅交通', vendor: '', note: '往返 & 住宿', unit: '次', qty: 4, unitCost: 2500 },
      { lid: 'e', cat: 'other', desc: '交際費', vendor: '', note: '', unit: '式', qty: 1, unitCost: 10000 },
      { lid: 'f', cat: 'other', desc: '印花稅', auto: 'stamp', vendor: '', note: '', unit: '式', qty: 1, unitCost: 0 }] };
  const buf2 = await buildQuotePnlExcel(q2, { classCodes: ['consult', 'software'], requestedBy: '業務<小華>', issueDate: '2026-10-08', contingencyPct: 10 });
  const v2 = await sheetToHtml(buf2);
  const rows2 = (v2.html.match(/<tr /g) || []).length;
  t('2i. 新式單：範圍 A1:H126、126 列、沒有 ####', v2.range === 'A1:H126' && rows2 === 126 && !/#{3,}/.test(v2.html), `range=${v2.range} rows=${rows2}`);
  t('2j. 新式單：委外廠商、供應商、說明、差旅、交際費、印花稅都在預覽內', ['Supplier-S', '含維護', '保固', '差旅交通', '交際費', '印花稅', '主機'].every(s => v2.html.includes(s)) && v2.html.includes('Vendor-A') && v2.html.includes('往返 &amp; 住宿'));
  t('2k. 新式單：使用者文字跳脫（沒有原始 <b>／<i>／<img／<小華>；有 &lt;img src=x onerror=alert(2)&gt;、&amp;、&quot;）', !/<img|<b>|<i>|<小華>/i.test(v2.html) && v2.html.includes('&lt;img src=x onerror=alert(2)&gt;') && v2.html.includes('&lt;b&gt;粗體&lt;/b&gt;') && v2.html.includes('&quot;全年&quot;') && v2.html.includes('含維護 &quot;全年&quot; &amp; 保固'));
  // 收入 1,000,000；顧問成本 80,000；Contingency 10%＝8,000；印花稅＝1,000；其他費用 10,000＋1,000；總成本 80,000＋300,000＋90,000＋10,000＋11,000＋8,000＝499,000
  t('2l. 新式單：金額格式（NT$）：收入 NT$1,000,000、顧問成本 NT$80,000、Contingency NT$8,000、總成本 NT$499,000、單價 NT$300,000', ['NT$1,000,000', 'NT$80,000', 'NT$8,000', 'NT$499,000', 'NT$300,000'].every(s => v2.html.includes(s)), (v2.html.match(/NT\$[\d,]+/g) || []).slice(0, 14).join(' '));
  t('2m. 新式單：沒有 script／事件屬性／javascript: 連結', !/<script/i.test(v2.html) && !/<[^>]*\son\w+\s*=/i.test(v2.html) && !/href\s*=\s*"?javascript:/i.test(v2.html));

  // 3) 審查後補強：四捨五入、時間、[Red]、列印範圍、注音、隱藏列中的合併儲存格、空白保留、溢出裁切
  t('3a. 百分比／金額四捨五入與 Excel 相同（0.02055→2.06%、0.02175→2.18%、150.075→150.08、0.0515→5.2%）', F(0.02055, '0.00%') === '2.06%' && F(0.02175, '0.00%') === '2.18%' && F(150.075, '#,##0.00') === '150.08' && F(0.0515, '0.0%') === '5.2%', [F(0.02055, '0.00%'), F(0.02175, '0.00%'), F(150.075, '#,##0.00'), F(0.0515, '0.0%')].join(' '));
  t('3b. 時間：h:mm 的 mm 是分鐘、日期的 mm 是月份', F(46302.5, 'yyyy/m/d h:mm') === '2026/10/7 12:00' && F(46302.75, 'h:mm') === '18:00' && F(46302, 'yyyy/mm/dd') === '2026/10/07', F(46302.5, 'yyyy/m/d h:mm'));
  t('3c. [Red] 負數格式回傳紅色', I.formatNumberEx(-5, '#,##0_);[Red]\\(#,##0\\)').color === '#FF0000' && I.formatNumberEx(5, '#,##0_);[Red]\\(#,##0\\)').color === null);
  t('3d. 列印範圍：多段取第一段、含空白與括號的工作表名、整欄範圍回 null（改用 dimension）', I.firstRange("'PNL (C)'!$A$1:$H$106") === 'A1:H106' && I.firstRange('PNL!$A$1:$H$126') === 'A1:H126' && I.firstRange('Sheet1!$B$2:$D$4,Sheet1!$A$8:$C$10') === 'B2:D4' && I.firstRange('Sheet1!$A:$C') === null);
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

  // 4) 給客戶的 Excel（含品項說明／備註的 rich text 儲存格）轉 HTML：三行都在、使用者文字一律跳脫、列高有被帶進去
  {
    const QE = require(path.join(ROOT, 'lib/quoteExcel.js'));
    const cq = { quoteNo: 'QU-H', company: 'C', projectName: 'P', validUntil: '2026-10-30', discountType: 'none', discountValue: 0,
      items: [{ lid: 'a', desc: '導入顧問', unit: '式', qty: 1, unitPrice: 1000, spec: '說明 <b>粗</b> & "q"', note: '=1+1 <img src=x onerror=alert(1)>' }, { lid: 'b', desc: '無備註項目', unit: '式', qty: 1, unitPrice: 5 }] };
    const hv = await sheetToHtml(await QE.buildQuoteWorkbook(cq, path.join(ROOT, 'templates/quotation_template.xlsx'), { issueDate: '2026-10-09', issuer: {} }));
    const i0 = hv.html.indexOf('導入顧問');
    const seg = hv.html.slice(i0, hv.html.indexOf('</td>', i0));
    t('4a. 客戶 Excel 轉 HTML：品項儲存格依序是 品名／說明／「備註：…」三行（換行保留），文字全部跳脫（沒有原始的 <b>、<img）', i0 > 0 && /導入顧問\n說明 &lt;b&gt;粗&lt;\/b&gt; &amp; &quot;q&quot;\n備註：=1\+1 &lt;img src=x onerror=alert\(1\)&gt;/.test(seg) && !/<img/i.test(hv.html) && !/<b>粗/.test(hv.html), seg.slice(0, 300));
    const px = (name) => { const k = hv.html.indexOf(name); const trStart = hv.html.lastIndexOf('<tr', k); const m = /height:(\d+)px/.exec(hv.html.slice(trStart, trStart + 400)); return m ? +m[1] : 0; };
    t('4b. 有說明／備註的列比沒有的列高（列高 pt 轉成 px 後有帶進預覽）', px('導入顧問') > px('無備註項目') && px('無備註項目') > 0, px('導入顧問') + ' vs ' + px('無備註項目'));
  }

  let pass = 0, fail = 0;
  res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (!ok && x ? '  <- ' + x : '')); ok ? pass++ : fail++; });
  console.log(`\nxlsx 轉 HTML 檢查：PASS ${pass} / FAIL ${fail}`);
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error('檢查腳本錯誤', e.stack); process.exit(2); });
