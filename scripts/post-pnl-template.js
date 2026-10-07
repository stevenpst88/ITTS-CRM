// 範本後處理（由 scripts/build-pnl-template.ps1 呼叫）：node post-pnl-template.js <Excel 另存的 xlsx> <輸出 xlsx>
// 移除 Excel 另存時帶進來的「內部網路印表機設定」與本機路徑（absPath），並把作者改成中性的 ITTS。
const fs = require('fs');
const path = require('path');
const JSZip = require(path.join(__dirname, '..', 'node_modules', 'jszip'));
const [, , inFile, outFile] = process.argv;
if (!inFile || !outFile) { console.error('用法：node post-pnl-template.js <in.xlsx> <out.xlsx>'); process.exit(2); }
(async () => {
  const z = await JSZip.loadAsync(fs.readFileSync(inFile));
  const log = [];
  // 1) 印表機設定部件、其關聯、sheet 的 <pageSetup r:id>
  const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  let rels = await z.file(relsPath).async('string');
  const m = /<Relationship\b[^>]*Type="[^"]*\/printerSettings"[^>]*\/>/.exec(rels);
  if (m) {
    const id = (/Id="([^"]+)"/.exec(m[0]) || [])[1], target = (/Target="([^"]+)"/.exec(m[0]) || [])[1];
    z.file(relsPath, rels.replace(m[0], ''));
    const part = 'xl/' + target.replace(/^\.\.\//, '');
    if (z.file(part)) { z.remove(part); log.push('移除 ' + part); }
    let sheet = await z.file('xl/worksheets/sheet1.xml').async('string');
    sheet = sheet.replace(new RegExp(`(<pageSetup\\b[^>]*?)\\s+r:id="${id}"`), '$1');
    z.file('xl/worksheets/sheet1.xml', sheet);
    log.push('pageSetup 去掉 r:id=' + id);
  } else log.push('沒有 printerSettings 關聯');
  // 2) 活頁簿的本機路徑 absPath
  let wb = await z.file('xl/workbook.xml').async('string');
  wb = wb.replace(/<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*<x15ac:absPath\b[^>]*\/>\s*<\/mc:Choice>\s*<\/mc:AlternateContent>/g, '');
  z.file('xl/workbook.xml', wb);
  // 3) 作者
  let core = await z.file('docProps/core.xml').async('string');
  core = core.replace(/<dc:creator>[^<]*<\/dc:creator>/, '<dc:creator>ITTS</dc:creator>').replace(/<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/, '<cp:lastModifiedBy>ITTS</cp:lastModifiedBy>');
  z.file('docProps/core.xml', core);
  fs.writeFileSync(outFile, await z.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log(log.join('；'));
})().catch(e => { console.error(e.stack); process.exit(1); });
