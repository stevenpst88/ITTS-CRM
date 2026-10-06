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
  signCell: 'G92',                                      // 右下簽名欄（Prepared by／ITTS Corp. 那塊，合併 G92:J94）；範本寫死的示意簽名一律清空
  // 報價專用章位置（oneCellAnchor，座標 0 起算）：簽名欄 G92:J94 約 389px × 93px，章放在欄內略偏右，
  // 錨點在 I 欄（index 8）左偏 22px、第 92 列（index 91）下偏 2px，最大 82px 見方（等比縮放）；不碰「Prepared by :」（第 91 列）與下方「ITTS Corp.」（第 95 列）。
  sealAnchor: { col: 8, colOffPx: 30, row: 91, rowOffPx: 2, maxPx: 82 },
};
const EMU_PER_PX = 9525;
const MAX_SEAL_BYTES = 1024 * 1024;                     // 章圖上傳端限制 300KB；這裡放寬到 1MB 當最後防線，避免異常大圖拖垮匯出
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

/**
 * 台灣（Asia/Taipei）的日期，格式 'YYYY-MM-DD'。
 * 伺服器（Vercel）是 UTC：不能用 toISOString() 或伺服器本地時間，否則台灣凌晨 0~8 點會差一天（月初還會差一個月）。
 */
function taipeiToday(now) {
  return new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit' }).format(now || new Date());
}

/**
 * 範本「電話：　　　　Ext.」是同一格文字，用空白把 Ext. 推到右邊。
 * 有電話/分機時重組成「電話：<電話>　　Ext.<分機>」，空白依電話長度扣減讓 Ext. 大致維持在原位。兩者都沒有則回傳 null（維持範本原樣）。
 */
function composePhoneExt(orig, phone, ext) {
  phone = String(phone || '').trim(); ext = String(ext || '').trim();
  if (!phone && !ext) return null;
  const m = /^([\s\S]*?：)\s*(Ext\.?)\s*$/.exec(orig || '');
  const label = m ? m[1] : '電    話：', extLabel = m ? m[2] : 'Ext.';
  return label + phone + ' '.repeat(Math.max(3, 26 - Math.round(phone.length * 1.8))) + extLabel + ext;
}

/**
 * 驗證章圖是否為「完整且可解碼結構」的 PNG / JPEG，並讀出像素寬高（用來等比縮放）。
 * 不信任呼叫端宣稱的 ext，以檔頭魔術數字為準。非法／截斷／過大 → 回傳 null（呼叫端略過章，匯出照常）。
 * @returns {{ext:'png'|'jpeg', w:number, h:number}|null}
 */
function sniffImage(buf) {
  if (!buf || typeof buf.length !== 'number' || buf.length < 24 || buf.length > MAX_SEAL_BYTES) return null;
  buf = Buffer.isBuffer(buf) ? buf : Buffer.from(buf);
  const LIM = 20000;
  // PNG：8 位元組簽章 + IHDR + 結尾須有 IEND
  if (buf.length >= 33 && buf.readUInt32BE(0) === 0x89504E47 && buf.readUInt32BE(4) === 0x0D0A1A0A
      && buf.toString('latin1', 12, 16) === 'IHDR') {
    const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
    if (!(w > 0 && h > 0 && w <= LIM && h <= LIM)) return null;
    if (buf.toString('latin1', Math.max(0, buf.length - 16)).indexOf('IEND') < 0) return null;
    return { ext: 'png', w, h };
  }
  // JPEG：FFD8FF 開頭，掃描 marker 找 SOF 取得寬高，結尾須有 EOI（FFD9）
  if (buf[0] === 0xFF && buf[1] === 0xD8 && buf[2] === 0xFF) {
    let i = 2, w = 0, h = 0;
    while (i + 3 < buf.length) {
      if (buf[i] !== 0xFF) return null;
      let m = buf[i + 1];
      while (m === 0xFF && i + 2 < buf.length) { i++; m = buf[i + 1]; }     // 填充用的 FF
      if (m === 0xD8 || m === 0x01 || (m >= 0xD0 && m <= 0xD7)) { i += 2; continue; }   // 無長度的 marker
      if (m === 0xD9 || m === 0xDA) break;                                  // EOI／SOS 之前仍沒看到 SOF → 不合法
      if (i + 4 > buf.length) return null;
      const len = buf.readUInt16BE(i + 2);
      if (len < 2) return null;
      if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) {   // SOFn
        if (i + 9 > buf.length) return null;
        h = buf.readUInt16BE(i + 5); w = buf.readUInt16BE(i + 7);
        break;
      }
      i += 2 + len;
    }
    if (!(w > 0 && h > 0 && w <= LIM && h <= LIM)) return null;
    const tail = buf.subarray(Math.max(0, buf.length - 64));
    let hasEoi = false;
    for (let k = 0; k + 1 < tail.length; k++) if (tail[k] === 0xFF && tail[k + 1] === 0xD9) { hasEoi = true; break; }
    if (!hasEoi) return null;
    return { ext: 'jpeg', w, h };
  }
  return null;
}

/**
 * 把報價專用章圖放進右下簽名欄：新增 xl/media/seal.<ext>、drawing 內追加 oneCellAnchor pic、補 rels 與 Content_Types。
 * 全部組好才寫回 zip；任何一步出錯 → 不動 zip、回傳 false（章被略過，匯出不受影響）。
 */
async function applySeal(zip, seal) {
  try {
    if (!seal || !seal.buffer) return false;
    const buf = Buffer.isBuffer(seal.buffer) ? seal.buffer : Buffer.from(seal.buffer);
    const info = sniffImage(buf);
    if (!info) return false;
    const drawFile = zip.file('xl/drawings/drawing1.xml');
    const relsFile = zip.file('xl/drawings/_rels/drawing1.xml.rels');
    const ctFile = zip.file('[Content_Types].xml');
    if (!drawFile || !relsFile || !ctFile) return false;
    let draw = await drawFile.async('string');
    let rels = await relsFile.async('string');
    let ct = await ctFile.async('string');
    if (draw.indexOf('</xdr:wsDr>') < 0 || rels.indexOf('</Relationships>') < 0) return false;

    const rIdNum = Math.max(0, ...[...rels.matchAll(/\bId="rId(\d+)"/g)].map(m => +m[1])) + 1;
    const rId = 'rId' + rIdNum;
    const picId = Math.max(1, ...[...draw.matchAll(/<xdr:cNvPr[^>]*?\bid="(\d+)"/g)].map(m => +m[1])) + 1;
    const mediaPath = `xl/media/seal.${info.ext}`;
    if (zip.file(mediaPath)) return false;

    const A = LAYOUT.sealAnchor;
    const scale = Math.min(A.maxPx / info.w, A.maxPx / info.h);   // 等比縮放進 maxPx 見方
    const cx = Math.max(1, Math.round(info.w * scale * EMU_PER_PX)), cy = Math.max(1, Math.round(info.h * scale * EMU_PER_PX));
    const pic =
      '<xdr:oneCellAnchor>' +
      `<xdr:from><xdr:col>${A.col}</xdr:col><xdr:colOff>${A.colOffPx * EMU_PER_PX}</xdr:colOff><xdr:row>${A.row}</xdr:row><xdr:rowOff>${A.rowOffPx * EMU_PER_PX}</xdr:rowOff></xdr:from>` +
      `<xdr:ext cx="${cx}" cy="${cy}"/>` +
      `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${picId}" name="報價專用章" descr="報價專用章"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
      `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
      `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>` +
      '<xdr:clientData/></xdr:oneCellAnchor>';
    draw = draw.replace('</xdr:wsDr>', () => pic + '</xdr:wsDr>');
    rels = rels.replace('</Relationships>', () => `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/seal.${info.ext}"/></Relationships>`);
    if (!new RegExp(`<Default[^>]*\\bExtension="${info.ext}"`, 'i').test(ct)) {
      ct = ct.replace(/(<Types\b[^>]*>)/, (m) => `${m}<Default Extension="${info.ext}" ContentType="image/${info.ext}"/>`);
    }
    zip.file(mediaPath, buf);
    zip.file('xl/drawings/drawing1.xml', draw);
    zip.file('xl/drawings/_rels/drawing1.xml.rels', rels);
    zip.file('[Content_Types].xml', ct);
    return true;
  } catch (e) {
    console.warn('[quoteExcel] 報價專用章略過：', e && e.message);
    return false;
  }
}

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
  /** 清空儲存格內容但保留樣式（框線／對齊） */
  clearCell(ref) { this._setCell(ref, (s) => `<c r="${ref}"${s}/>`); }
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
 * @param {object} [opts]
 * @param {string} [opts.issueDate] 出單日期 'YYYY-MM-DD'（未給＝台灣今天）。每次匯出都蓋當天，不使用報價單上存的日期
 * @param {object} [opts.issuer]    我方業務聯絡資訊 {name, phone, ext, mobile}，填在右上「廠商資料」框
 * @param {boolean} [opts.approved] 報價單已核准且內容未變（approval.state==='approved' && valid）。只有 true 才會蓋章
 * @param {object} [opts.report] 傳入一個空物件，完成後會被填入 {sealApplied:boolean}（實際是否蓋上章）
 * @param {{buffer:Buffer, ext:'png'|'jpeg'}|null} [opts.seal] 報價專用章圖。實際格式以檔頭判定（ext 僅供參考）；非合法 PNG/JPEG 或超過 1MB 會被略過（匯出照常、不蓋章）
 * @returns {Promise<Buffer>}
 */
async function buildQuoteWorkbook(q, templatePath, opts = {}) {
  const zip = await JSZip.loadAsync(fs.readFileSync(templatePath));
  const sheetFile = zip.file(LAYOUT.sheetPath);
  if (!sheetFile) throw new Error('範本格式不符：找不到 ' + LAYOUT.sheetPath);
  const sstFile = zip.file('xl/sharedStrings.xml');
  const sst = sstFile ? parseSharedStrings(await sstFile.async('string')) : [];
  const sh = new SheetXml(await sheetFile.async('string'), sst);
  const L = LAYOUT;

  // ── 表頭 ──
  sh.setText('G6', `表單編號：${q.quoteNo || ''}`);
  const dateStr = String(opts.issueDate || taipeiToday()).replace(/-/g, '/');
  const iss = opts.issuer || {};
  sh.fillLabeled('B9',  q.company);        // 左框「客戶資料」
  sh.fillLabeled('B10', q.contactName);
  sh.fillLabeled('B11', q.address);
  sh.fillLabeled('B12', q.phone);
  sh.fillLabeled('F9',  dateStr);          // 右框「廠商資料」＝我方業務（由業務在系統內自行維護聯絡資訊）
  sh.fillLabeled('F10', iss.name);
  sh.fillLabeled('F11', iss.mobile);
  const f12 = composePhoneExt(sh.cellText('F12'), iss.phone, iss.ext);
  if (f12) sh.setText('F12', f12);

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

  // ── 右下簽名欄：範本寫死的示意簽名一律清空；已核准且有章才蓋「報價專用章」 ──
  sh.clearCell(L.signCell);

  zip.file(L.sheetPath, sh.xml);
  let sealApplied = false;
  if (opts.approved === true && opts.seal) sealApplied = await applySeal(zip, opts.seal);
  // 讓呼叫端知道「實際有沒有蓋章」（章圖缺漏／損壞時匯出照常但不蓋章，呼叫端不可再宣稱已蓋章）
  if (opts.report && typeof opts.report === 'object') opts.report.sealApplied = sealApplied;

  // 開檔時強制重算（保險；快取值已寫入）
  const wbFile = zip.file('xl/workbook.xml');
  if (wbFile) {
    let wbx = await wbFile.async('string');
    wbx = wbx.replace(/<calcPr([^>]*?)(\/?)>/, (m, a, s) => /fullCalcOnLoad/.test(a) ? m : `<calcPr${a} fullCalcOnLoad="1"${s}>`);
    zip.file('xl/workbook.xml', wbx);
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

module.exports = { buildQuoteWorkbook, taipeiToday, LAYOUT, ITEM_CAP, _internal: { SheetXml, parseSharedStrings, itemRowHeight, composePhoneExt, sniffImage, applySeal } };
