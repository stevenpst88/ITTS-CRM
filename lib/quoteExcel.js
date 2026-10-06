'use strict';
/**
 * 報價單 Excel 產生器 —— 直接把資料填進範本，不重寫整份檔案。
 *
 * 為什麼不用 SheetJS（xlsx）：社群版讀進範本再寫出時，會把框線、字型、自動換行、欄寬、
 * 圖片（logo）、列印設定通通丟掉，輸出就變成「沒有格線、文字被擋住」。
 * 這裡改用 JSZip 直接修改範本內 sheet1.xml 的儲存格，其餘部件（樣式、圖片、列印設定）原封不動。
 *
 * 範本（templates/quotation_template.xlsx）是用 Excel 整理成「單一工作表、無外部連結」的乾淨版，
 * 品項區預留 50 列（17~66），用不到的列在輸出時隱藏。**若重做範本，請同步下方 LAYOUT。**
 *   重做方式：用 Excel 開啟，保留「報價單」一張工作表 → 在「以下空白」列上方插入品項列並複製格式
 *   → 另存 xlsx → 移除 xl/calcChain.xml（含 [Content_Types].xml 與 workbook.xml.rels 內的引用）。
 */
const fs = require('fs');
const JSZip = require('jszip');

const LAYOUT = {
  sheetPath: 'xl/worksheets/sheet1.xml',
  itemFirst: 17, itemLast: 66,                          // 品項區（共 50 列）
  sumList: 68, sumDisc: 69, sumTax: 71, sumTotal: 72,   // 專案定價 / 優惠價 / 稅金5% / 優惠價(含稅)
  projName: 77, projNo: 80,                             // 專案名稱 / 專案號碼
  noteRow: 90,                                          // Remarks 第 6 條與簽名區之間的空白列：放「備註」
  descWidthUnits: 72,                                   // 「內容」欄（C:E 合併）一行可容納的半形字元寬
};
const ITEM_CAP = LAYOUT.itemLast - LAYOUT.itemFirst + 1;

// ── XML 小工具 ─────────────────────────────────────────────
const escXml = (s) => String(s == null ? '' : s)
  .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const unescXml = (s) => s
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
  .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&amp;/g, '&');
const colIndex = (letters) => letters.split('').reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
const splitRef = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return { col: m[1], row: +m[2] }; };
const num = (v) => { const n = Number(v); return Number.isFinite(n) ? n : 0; };

function parseSharedStrings(xml) {
  return [...xml.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m =>
    [...m[1].replace(/<rPh[\s\S]*?<\/rPh>/g, '').matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)]
      .map(t => unescXml(t[1])).join(''));
}

/** 內容欄的列高估算：全形字算 2 個半形寬，依欄寬折行；Excel 對「合併儲存格」不會自動調列高，必須自己算 */
function itemRowHeight(desc) {
  const width = LAYOUT.descWidthUnits;
  let lines = 0;
  for (const ln of String(desc || '').split(/\r?\n/)) {
    let u = 0;
    for (const ch of ln) u += ch.charCodeAt(0) > 0x2E7F ? 2 : 1;
    lines += Math.max(1, Math.ceil(u / width));
  }
  return Math.min(409, Math.max(21, lines * 18 + 4));   // 409 為 Excel 單列高度上限
}

class SheetXml {
  constructor(xml, sst) { this.xml = xml; this.sst = sst || []; }

  _rowRe(r) { return new RegExp(`<row r="${r}"(?=[\\s/>])([^>]*?)(?:/>|>([\\s\\S]*?)</row>)`); }
  _getRow(r) {
    const m = this._rowRe(r).exec(this.xml);
    if (!m) throw new Error(`範本缺少第 ${r} 列，請確認範本版面與 lib/quoteExcel.js 的 LAYOUT 一致`);
    return { attrs: m[1], inner: m[2] || '' };
  }
  _putRow(r, row) { this.xml = this.xml.replace(this._rowRe(r), () => `<row r="${r}"${row.attrs}>${row.inner}</row>`); }

  /** 讀取儲存格目前的文字（含共用字串），用來保留範本上的標籤，例如「公    司：」 */
  cellText(ref) {
    const { row } = splitRef(ref);
    const { inner } = this._getRow(row);
    const m = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>([\\s\\S]*?)</c>)`).exec(inner);
    if (!m || !m[2]) return '';
    if (/\bt="s"/.test(m[1])) { const v = /<v>(\d+)<\/v>/.exec(m[2]); return v ? (this.sst[+v[1]] || '') : ''; }
    if (/\bt="inlineStr"/.test(m[1])) return unescXml([...m[2].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(t => t[1]).join(''));
    return '';
  }

  _setCell(ref, build) {
    const { col, row } = splitRef(ref);
    const R = this._getRow(row);
    const re = new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>[\\s\\S]*?</c>)`);
    let found = false;
    R.inner = R.inner.replace(re, (m, attrs) => {
      found = true;
      const s = /\ss="(\d+)"/.exec(attrs);          // 保留範本原本的樣式（框線／字型／對齊／數字格式）
      return build(s ? ` s="${s[1]}"` : '');
    });
    if (!found) {                                    // 範本沒有這格：依欄位順序插入
      let at = R.inner.length;
      for (const c of R.inner.matchAll(/<c r="([A-Z]+)\d+"/g)) { if (colIndex(c[1]) > colIndex(col)) { at = c.index; break; } }
      R.inner = R.inner.slice(0, at) + build('') + R.inner.slice(at);
    }
    this._putRow(row, R);
  }

  setText(ref, text) {
    this._setCell(ref, (s) => `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${escXml(text)}</t></is></c>`);
  }
  setNumber(ref, n) { this._setCell(ref, (s) => `<c r="${ref}"${s}><v>${num(n)}</v></c>`); }
  /** 公式一併寫入快取值，這樣不重新計算的檢視器（預覽、手機、雲端預覽）也看得到數字 */
  setFormula(ref, formula, cached) {
    this._setCell(ref, (s) => `<c r="${ref}"${s}><f>${escXml(formula)}</f><v>${num(cached)}</v></c>`);
  }
  setRowHeight(r, pt) {
    const R = this._getRow(r);
    R.attrs = R.attrs.replace(/\sht="[^"]*"/, '').replace(/\scustomHeight="[^"]*"/, '') + ` ht="${pt}" customHeight="1"`;
    this._putRow(r, R);
  }
  hideRow(r) {
    const R = this._getRow(r);
    R.attrs = R.attrs.replace(/\shidden="[^"]*"/, '') + ' hidden="1"';
    this._putRow(r, R);
  }
  /** 範本上「標籤：」格 → 在標籤後接值（沒有值就維持原樣，不破壞標籤） */
  fillLabeled(ref, value) {
    const v = String(value == null ? '' : value).trim();
    if (!v) return;
    this.setText(ref, this.cellText(ref) + v);
  }
}

/**
 * 產生給客戶的報價單 Excel（只有「報價單」一張工作表）。
 * @param {object} q            報價單資料
 * @param {string} templatePath 範本路徑
 * @returns {Promise<Buffer>}
 */
async function buildQuoteWorkbook(q, templatePath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(templatePath));
  const sheetFile = zip.file(LAYOUT.sheetPath);
  if (!sheetFile) throw new Error('範本格式不符：找不到 ' + LAYOUT.sheetPath);
  const sstFile = zip.file('xl/sharedStrings.xml');
  const sst = sstFile ? parseSharedStrings(await sstFile.async('string')) : [];
  const sh = new SheetXml(await sheetFile.async('string'), sst);
  const L = LAYOUT;

  // ── 表頭 ──
  sh.setText('G6', `表單編號：${q.quoteNo || ''}`);
  const dateStr = String(q.quoteDate || new Date().toISOString().slice(0, 10)).replace(/-/g, '/');
  sh.fillLabeled('B9',  q.company);        // 客戶資料
  sh.fillLabeled('B10', q.contactName);
  sh.fillLabeled('B11', q.address);
  sh.fillLabeled('B12', q.phone);
  sh.fillLabeled('F9',  dateStr);          // 右側資料框（沿用原本的欄位對應）
  sh.fillLabeled('F10', q.contactName);
  sh.fillLabeled('F11', q.mobile);

  // ── 品項（最多 ITEM_CAP 列；不夠的列隱藏）──
  const items = (Array.isArray(q.items) ? q.items : []).slice(0, ITEM_CAP);
  let listSum = 0;
  items.forEach((it, i) => {
    const r = L.itemFirst + i;
    const qty = num(it.qty) || 1, price = num(it.unitPrice);
    const sub = qty * price;
    listSum += sub;
    sh.setNumber(`B${r}`, i + 1);
    sh.setText(`C${r}`, it.desc || '');
    sh.setNumber(`F${r}`, qty);                       // 範本欄位：F=數量、G=單位
    sh.setText(`G${r}`, it.unit || '式');
    sh.setNumber(`H${r}`, price);
    sh.setFormula(`J${r}`, `H${r}*F${r}`, sub);
    sh.setRowHeight(r, itemRowHeight(it.desc));
  });
  for (let r = L.itemFirst + Math.max(items.length, 1); r <= L.itemLast; r++) sh.hideRow(r);

  // ── 小計／優惠／稅／合計 ──
  const dType = q.discountType || 'none', dVal = num(q.discountValue);
  let discNote = '';
  sh.setFormula(`J${L.sumList}`, `SUM(J${L.itemFirst}:J${L.itemLast})`, listSum);
  let disc;
  if (dType === 'percent' && dVal > 0 && dVal < 100) {
    disc = listSum * dVal / 100;
    sh.setFormula(`J${L.sumDisc}`, `J${L.sumList}*${dVal}/100`, disc);
    discNote = `專案優惠 ${dVal}%（${+(dVal / 10).toFixed(1)} 折）`;
  } else if (dType === 'amount' && dVal > 0) {
    disc = dVal;
    sh.setNumber(`J${L.sumDisc}`, dVal);
    discNote = '專案議價金額';
  } else {
    disc = listSum;
    sh.setFormula(`J${L.sumDisc}`, `J${L.sumList}`, disc);
  }
  const tax = disc * 0.05;
  sh.setFormula(`J${L.sumTax}`, `J${L.sumDisc}*0.05`, tax);
  sh.setFormula(`J${L.sumTotal}`, `J${L.sumTax}+J${L.sumDisc}`, disc + tax);
  if (discNote) sh.setText(`C${L.sumDisc}`, discNote);   // 標籤「專案優惠價(未稅)」已佔滿合併格，折扣說明放在左側空白處

  // ── 專案資料／備註 ──
  sh.fillLabeled(`B${L.projName}`, q.projectName);
  sh.fillLabeled(`B${L.projNo}`, q.projectNo);
  const note = String(q.note || '').replace(/\s*[\r\n]+\s*/g, ' ').trim();
  if (note) sh.setText(`B${L.noteRow}`, `7.${note}`);

  zip.file(L.sheetPath, sh.xml);

  // 開檔時強制重算（保險；快取值已寫入）
  const wbFile = zip.file('xl/workbook.xml');
  if (wbFile) {
    let wbx = await wbFile.async('string');
    wbx = wbx.replace(/<calcPr([^>]*?)(\/?)>/, (m, a, s) => /fullCalcOnLoad/.test(a) ? m : `<calcPr${a} fullCalcOnLoad="1"${s}>`);
    zip.file('xl/workbook.xml', wbx);
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

module.exports = { buildQuoteWorkbook, LAYOUT, ITEM_CAP, _internal: { SheetXml, parseSharedStrings, itemRowHeight } };
