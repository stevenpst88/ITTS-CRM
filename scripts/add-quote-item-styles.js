#!/usr/bin/env node
/**
 * 一次性（可重跑）：替 templates/quotation_template.xlsx 的 xl/styles.xml 加入「分組標題列」用的儲存格樣式。
 *   新增 fill（淺灰 theme 0 tint -0.05）→ fills 6→7；新增 cellXfs 第 103 號（微軟正黑體 12 粗體、靠左、自動換行、淺灰底、全框）。
 * lib/quoteExcel.js 的 LAYOUT.titleStyle=103 對應這個索引。報價單的「分組標題」列（Part A／Part B…）用它；
 * 「小計」列直接用範本摘要區既有的樣式（標籤 94、金額 42），不需要新增。
 * 已經加過（fills count 已是 7 且 cellXfs 已是 104）就什麼都不做。用法：node scripts/add-quote-item-styles.js
 * 注意：若之後用 scripts/build-quote-template-remarks.ps1 之類的流程重建範本，要再跑一次本腳本。
 */
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const TEMPLATE = path.join(__dirname, '..', 'templates', 'quotation_template.xlsx');

(async () => {
  const zip = await JSZip.loadAsync(fs.readFileSync(TEMPLATE));
  let st = await zip.file('xl/styles.xml').async('string');
  const fills = /<fills count="(\d+)">/.exec(st), xfs = /<cellXfs count="(\d+)">/.exec(st);
  if (!fills || !xfs) throw new Error('styles.xml 結構不符預期');
  if (+fills[1] === 7 && +xfs[1] === 104) { console.log('樣式已存在，略過'); return; }
  if (+fills[1] !== 6 || +xfs[1] !== 103) throw new Error(`styles.xml 的 fills=${fills[1]}、cellXfs=${xfs[1]} 與預期（6／103）不同，請人工確認後再改本腳本`);
  st = st.replace('<fills count="6">', '<fills count="7">')
    .replace('</fills>', '<fill><patternFill patternType="solid"><fgColor theme="0" tint="-4.9989318521683403E-2"/><bgColor indexed="64"/></patternFill></fill></fills>')
    .replace('<cellXfs count="103">', '<cellXfs count="104">')
    .replace('</cellXfs>', '<xf numFmtId="0" fontId="19" fillId="6" borderId="6" xfId="13" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1" indent="1"/></xf></cellXfs>');
  zip.file('xl/styles.xml', st);
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 9 } });
  fs.writeFileSync(TEMPLATE, buf);
  console.log('已加入分組標題樣式（cellXfs 103），範本大小', buf.length);
})().catch((e) => { console.error(e.message); process.exit(1); });
