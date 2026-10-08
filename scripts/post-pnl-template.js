// 範本後處理（由 scripts/build-pnl-template.ps1 呼叫）：node post-pnl-template.js <Excel 另存的 xlsx> <輸出 xlsx>
// ★ 輸入檔（Excel 另存的中間檔）與 build-pnl-template.ps1 的來源範本一樣需自備、不可放進 repo（可能含客戶資料）；只有清乾淨的輸出 templates/pnl_template.xlsx 可以入 repo。
// 1) 移除 Excel 另存時帶進來的「內部網路印表機設定」與本機路徑（absPath），作者改成中性的 ITTS
// 2) 把共用公式（<f t="shared" ...>）展開成每格獨立公式（相對參照依位移平移）：
//    lib/quotePnlExcel.js 用字串補丁改儲存格，共用公式的子格或母格被覆寫就會壞
// 3) 移除 calcChain.xml（系統覆寫公式儲存格後，calcChain 指到沒有公式的格子會讓 Excel 報「修復」；Excel 開檔時自行重建）
// 4) 設 calcPr fullCalcOnLoad="1"（開檔時重算），活頁簿檢視回到第一張工作表
const fs = require('fs');
const path = require('path');
const JSZip = require(path.join(__dirname, '..', 'node_modules', 'jszip'));
const [, , inFile, outFile] = process.argv;

const colNum = (c) => c.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colName = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

/** 平移公式中的相對參照（$ 鎖定的欄/列不動）。只處理「儲存格參照」，函式名（SUM、IF）後面沒有數字所以不會誤判 */
function shiftFormula(f, dRow, dCol) {
  return f.replace(/(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])/g, (m, ca, c, ra, r) => {
    const nc = ca ? c : colName(colNum(c) + dCol);
    const nr = ra ? +r : +r + dRow;
    if (nr < 1 || colNum(nc) < 1) throw new Error('平移後參照超出範圍：' + f);
    return ca + nc + ra + nr;
  });
}
const attrOf = (s, name) => { const m = new RegExp('\\b' + name + '="([^"]*)"').exec(s); return m ? m[1] : null; };

/** 回傳 { xml, expanded } */
function expandSharedFormulas(xml) {
  const masters = new Map();                       // si → { row, col, text }
  const cellRe = /<c r="([A-Z]+)(\d+)"([^>]*?)>([\s\S]*?)<\/c>/g;
  for (const m of xml.matchAll(cellRe)) {
    const f = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(m[4]);
    if (!f || attrOf(f[1], 't') !== 'shared' || !f[2]) continue;
    masters.set(attrOf(f[1], 'si'), { row: +m[2], col: colNum(m[1]), text: f[2] });
  }
  let expanded = 0;
  const out = xml.replace(cellRe, (whole, col, row, attrs, inner) => {
    const f = /<f\b([^>]*?)(\/>|>([\s\S]*?)<\/f>)/.exec(inner);
    if (!f || attrOf(f[1], 't') !== 'shared') return whole;
    const si = attrOf(f[1], 'si');
    const mst = masters.get(si);
    if (!mst) throw new Error('共用公式 si=' + si + ' 找不到母格（' + col + row + '）');
    const text = f[3] ? f[3] : shiftFormula(mst.text, +row - mst.row, colNum(col) - mst.col);
    expanded++;
    return `<c r="${col}${row}"${attrs}>` + inner.replace(f[0], `<f>${text}</f>`) + '</c>';
  });
  return { xml: out, expanded };
}

async function main() {
  if (!inFile || !outFile) { console.error('用法：node post-pnl-template.js <in.xlsx> <out.xlsx>'); process.exit(2); }
  const z = await JSZip.loadAsync(fs.readFileSync(inFile));
  const log = [];
  const read = (p) => z.file(p).async('string');

  // 工作表檔：活頁簿只剩一張
  const wb0 = await read('xl/workbook.xml');
  const sheetRids = [...wb0.matchAll(/<sheet\b[^>]*r:id="([^"]+)"/g)].map(m => m[1]);
  if (sheetRids.length !== 1) throw new Error('活頁簿應該只有 1 張工作表，實際 ' + sheetRids.length);
  const wbRelsPath = 'xl/_rels/workbook.xml.rels';
  let wbRels = await read(wbRelsPath);
  const relTag = new RegExp('<Relationship\\b[^>]*Id="' + sheetRids[0] + '"[^>]*/>').exec(wbRels)[0];
  const sheetPath = 'xl/' + /Target="([^"]+)"/.exec(relTag)[1].replace(/^\/?xl\//, '');
  const sheetRelsPath = sheetPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';
  log.push('工作表=' + sheetPath);

  // 1) 印表機設定部件、其關聯、sheet 的 <pageSetup r:id>
  let sheet = await read(sheetPath);
  if (z.file(sheetRelsPath)) {
    let rels = await read(sheetRelsPath);
    const m = /<Relationship\b[^>]*Type="[^"]*\/printerSettings"[^>]*\/>/.exec(rels);
    if (m) {
      const id = /Id="([^"]+)"/.exec(m[0])[1], target = /Target="([^"]+)"/.exec(m[0])[1];
      rels = rels.replace(m[0], '');
      z.file(sheetRelsPath, rels);
      const part = 'xl/' + target.replace(/^\.\.\//, '');
      if (z.file(part)) { z.remove(part); log.push('移除 ' + part); }
      sheet = sheet.replace(new RegExp(`(<pageSetup\\b[^>]*?)\\s+r:id="${id}"`), '$1');
      log.push('pageSetup 去掉 r:id=' + id);
    }
  }

  // 2) 展開共用公式
  const ex = expandSharedFormulas(sheet);
  sheet = ex.xml;
  log.push('展開共用公式 ' + ex.expanded + ' 格');
  if (/t="shared"/.test(sheet)) throw new Error('仍有共用公式未展開');
  z.file(sheetPath, sheet);

  // 3) 移除 calcChain（部件、關聯、Content_Types）
  if (z.file('xl/calcChain.xml')) {
    z.remove('xl/calcChain.xml');
    wbRels = wbRels.replace(/<Relationship\b[^>]*Type="[^"]*\/calcChain"[^>]*\/>/, '');
    z.file(wbRelsPath, wbRels);
    let ct = await read('[Content_Types].xml');
    ct = ct.replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/, '');
    z.file('[Content_Types].xml', ct);
    log.push('移除 calcChain');
  }

  // 4) 活頁簿：本機路徑 absPath、fullCalcOnLoad、檢視回到第一張
  let wb = await read('xl/workbook.xml');
  wb = wb.replace(/<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*<x15ac:absPath\b[^>]*\/>\s*<\/mc:Choice>\s*<\/mc:AlternateContent>/g, '');
  wb = wb.replace(/<calcPr([^>]*?)(\/?)>/, (m, a, s) => /fullCalcOnLoad/.test(a) ? m : `<calcPr${a} fullCalcOnLoad="1"${s}>`);
  wb = wb.replace(/(<workbookView\b[^>]*?)\s+firstSheet="\d+"/, '$1').replace(/(<workbookView\b[^>]*?)\s+activeTab="\d+"/, '$1');
  z.file('xl/workbook.xml', wb);

  // 5) 作者／公司：中性的 ITTS
  let core = await read('docProps/core.xml');
  core = core.replace(/<dc:creator>[^<]*<\/dc:creator>/, '<dc:creator>ITTS</dc:creator>').replace(/<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/, '<cp:lastModifiedBy>ITTS</cp:lastModifiedBy>');
  z.file('docProps/core.xml', core);
  if (z.file('docProps/app.xml')) {
    let app = await read('docProps/app.xml');
    app = app.replace(/<Company>[^<]*<\/Company>/, '<Company>ITTS</Company>').replace(/<Manager>[^<]*<\/Manager>/, '');
    z.file('docProps/app.xml', app);
  }

  fs.writeFileSync(outFile, await z.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log(log.join('；'));
}
if (require.main === module) main().catch(e => { console.error(e.stack); process.exit(1); });

module.exports = { shiftFormula, expandSharedFormulas };
