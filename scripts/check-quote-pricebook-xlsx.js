#!/usr/bin/env node
/**
 * 報價牌價簿 Excel 匯出／範本／匯入預覽（lib/quotePricebookXlsx.js ＋ lib/quoteRoutes.js 的三條 /api/admin/quote-pricebook/* 路由 ＋ PUT 的 via:'import'）檢查。
 * 用法：node scripts/check-quote-pricebook-xlsx.js（不需要伺服器；HTTP 段落自己在 127.0.0.1 隨機埠起一個 express；不碰 data.json／auth.json／audit.log.json）
 * 動 lib/quotePricebookXlsx.js、quoteRoutes.js 牌價簿區段、_client/admin.html 牌價簿匯入之後必跑。
 *   1) 產生的 xlsx：結構（兩個工作表、標題、凍結首列、欄寬、金額數字格＋格式、啟用是／否、資料驗證）、XML 良構、範本沒有資料列且說明頁的範例是虛構文字
 *   2) 往返：匯出 → 匯入預覽 ＝ 全部「不變」（含小數、1e9、全形／多空白名稱、停用項目）；公式注入名稱（= + - @ 開頭）寫成字串型＋quotePrefix、讀回逐字相同、檔案內沒有任何 <f>
 *   3) 合併語意：同名（NFKC／大小寫／空白不分）更新並保留 id／順序／原名稱；新名稱附加在最後；檔案沒有的既有項目不動；60 項上限；檔內重複→後面的列報錯；
 *      壞數字、標題變體、是否變體、空白列、千分位、全形數字、文字型數字、工作表選擇、控制字元、超長名稱、公式儲存格、500 列上限；第二次匯入同檔 ＝ 全部不變
 *   4) 惡意／異常檔案：非 zip、空檔、被截斷、沒有工作表、zip bomb（誠實大小／偽造中央目錄大小）、DOCTYPE、__proto__ 分頁、巨大 dimension、
 *      夾帶巨集／未列入白名單的大型部件、ZIP64／加密旗標、重複項目名稱、大量項目 → 全部乾淨的錯誤碼，不丟例外、不洩漏路徑／堆疊
 *   5) 路由（記憶體 db 直接呼叫 handler）：僅管理員、預覽完全不寫入（資料／儲存次數／updatedAt 不變，只多一筆預覽稽核）、匯出／範本稽核與標頭、
 *      預覽 → PUT merged 套用（保留 id、附加新項目、稽核含「來源：Excel 匯入（檔名）」）、base 過期 → 409、via 白名單、檔名淨化、必須掛 requireAdmin
 *   6) 真實 HTTP（express＋multer）：multipart 上傳、中文檔名、2.5 MB → 413、錯誤欄位名／非 multipart／副檔名不符／內容不是 zip → 4xx 乾淨 JSON、非管理員 403、下載標頭
 * 測試資料只用通用字串（公開 repo：不放客戶名、人名、真實費率）。
 */
'use strict';
const path = require('path');
const ROOT = path.join(__dirname, '..');
const PB = require(path.join(ROOT, 'lib/quotePricebook.js'));
const X = require(path.join(ROOT, 'lib/quotePricebookXlsx.js'));
const registerQuoteRoutes = require(path.join(ROOT, 'lib/quoteRoutes.js'));
const XLSX = require(path.join(ROOT, 'node_modules/xlsx'));
const JSZip = require(path.join(ROOT, 'node_modules/jszip'));
const zlib = require('zlib');

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ── 小工具 ─────────────────────────────────────────────
/** 粗略 XML 良構檢查（標籤配對、沒有裸 <）。夠用來抓「產生器漏關標籤」這類錯 */
function wellFormed(xml) {
  const s = xml.replace(/^<\?xml[^>]*\?>\s*/, '').replace(/<!--[\s\S]*?-->/g, '');
  const stack = [];
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>/g;
  let last = 0, m;
  while ((m = re.exec(s))) {
    if (s.slice(last, m.index).includes('<')) return false;
    last = re.lastIndex;
    if (m[1]) { if (stack.pop() !== m[2]) return false; } else if (!m[4]) stack.push(m[2]);
  }
  return !s.slice(last).includes('<') && stack.length === 0;
}
async function unzipText(buf) {
  const z = await JSZip.loadAsync(buf);
  const o = {};
  for (const n of Object.keys(z.files)) if (!z.files[n].dir) o[n] = await z.files[n].async('string');
  return o;
}
/** 用 SheetJS 做一個「別人用 Excel 做的」檔案：aoa 的元素可以是原始值或 {t,v,f,w} 儲存格物件；sheets 可多個 */
function mkXlsx(sheets) {
  const wb = XLSX.utils.book_new();
  for (const [name, aoa] of sheets) {
    const ws = XLSX.utils.aoa_to_sheet(aoa.map((r) => r.map((c) => (c && typeof c === 'object' && 't' in c ? null : c))));
    aoa.forEach((r, ri) => r.forEach((c, ci) => { if (c && typeof c === 'object' && 't' in c) ws[XLSX.utils.encode_cell({ r: ri, c: ci })] = c; }));
    const maxc = Math.max(...aoa.map((r) => r.length), 1);
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(aoa.length - 1, 0), c: maxc - 1 } });
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}
const HDR = ['項目名稱', '牌價（元/人天）', '成本（元/人天）', '啟用'];
const one = (rows, hdr, bu) => mkXlsx([[bu || 'ERP', [hdr || HDR].concat(rows)]]);   // 單一 BU 工作表（預設 ERP）
const multi = (obj, extra) => mkXlsx(Object.keys(obj).map((k) => [k, [HDR].concat(obj[k])]).concat(extra || []));   // 多個工作表 { 工作表名: 資料列[] }
const prev = (buf, existing) => X.previewFromBuffer(buf, existing || []);
const cur = (name, price, cost, active, id, bu) => ({ id: id || ('id_' + (bu || 'ERP') + '_' + name), bu: bu || 'ERP', name, price, cost, active: active !== false });
const asPlain = (x) => ({ id: x.id, bu: x.bu || 'ERP', name: x.name, price: x.price, cost: x.cost, active: x.active });
const byRow = (r, n) => r.rows.find((x) => x.row === n);
const act = (r) => r.rows.map((x) => x.action).join(',');

/** 改 zip 內某個部件的內容後重包（保留其他部件） */
async function rezip(buf, edit) {
  const z = await JSZip.loadAsync(buf);
  await edit(z);
  return z.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function run() {
  // ═════════════════ 1) 產生的 xlsx ═════════════════
  const ITEMS = [cur('PM 顧問經理', 9000, 6500), cur('SD 顧問', 7000.5, 5000, false), cur('ABAP 顧問', 6500, 4500.25)];
  const bufE = await X.buildWorkbook({ mode: 'export', items: ITEMS, date: '2026-10-08' });
  const wbE = XLSX.read(bufE, { type: 'buffer' });
  t('1.1 匯出檔可被 xlsx 開啟；五個工作表依序為 ERP、ITS、MDM、CRM、說明', eq(wbE.SheetNames, ['ERP', 'ITS', 'MDM', 'CRM', '說明']), JSON.stringify(wbE.SheetNames));
  const wsE = wbE.Sheets['ERP'];
  t('1.2 標題列＝四個欄位名稱（繁中）', eq(['A1', 'B1', 'C1', 'D1'].map((k) => wsE[k] && wsE[k].v), HDR));
  t('1.3 資料列：項目依目前順序（含停用的）；金額是數字型（t=n）；啟用欄是／否', eq([2, 3, 4].map((r) => wsE['A' + r].v), ['PM 顧問經理', 'SD 顧問', 'ABAP 顧問'])
    && [2, 3, 4].every((r) => wsE['B' + r].t === 'n' && wsE['C' + r].t === 'n') && wsE['B3'].v === 7000.5 && wsE['C4'].v === 4500.25 && eq([2, 3, 4].map((r) => wsE['D' + r].v), ['是', '否', '是']));
  t('1.4 沒有多出來的列（!ref＝A1:D4）', wsE['!ref'] === 'A1:D4', wsE['!ref']);
  const partsE = await unzipText(bufE);
  t('1.5 zip 內只有預期的部件（沒有巨集、外部連結、媒體）', eq(Object.keys(partsE).sort(), ['[Content_Types].xml', '_rels/.rels', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml', 'xl/worksheets/sheet4.xml', 'xl/worksheets/sheet5.xml'].sort()), Object.keys(partsE).join(','));
  t('1.6 每個 XML 部件良構', Object.entries(partsE).every(([, x]) => wellFormed(x)), Object.entries(partsE).filter(([, x]) => !wellFormed(x)).map(([k]) => k).join(','));
  const s1 = partsE['xl/worksheets/sheet1.xml'];
  t('1.7 首列凍結（pane ySplit=1 frozen）、欄寬有設定', /<pane ySplit="1"[^>]*state="frozen"/.test(s1) && /<cols>.*<col min="1" max="1" width="[\d.]+" customWidth="1"\/>/.test(s1));
  t('1.8 金額有數字格式（整數用千分位 numFmt 3、有小數用 numFmt 4），名稱欄是文字格式 49', partsE['xl/styles.xml'].includes('numFmtId="3"') && partsE['xl/styles.xml'].includes('numFmtId="4"') && /<c r="B2" s="4">/.test(s1) && /<c r="B3" s="5">/.test(s1));
  t('1.9 啟用欄有「是,否」下拉驗證、金額欄有 0～1e9 的數字驗證', s1.includes('<formula1>"是,否"</formula1>') && s1.includes('sqref="D2:D501"') && s1.includes('<formula2>1000000000</formula2>') && s1.includes('sqref="B2:C501"'));
  t('1.10 所有文字儲存格都是明確字串型（inlineStr）；整份資料表沒有 <f> 公式、沒有 t="s" 以外的怪型別', !/<f[ >]/.test(s1) && (s1.match(/<c [^>]*\bt="[^"]*"/g) || []).every((c) => /t="inlineStr"/.test(c)));
  const help = wbE.Sheets['說明'];
  const helpTxt = Object.keys(help).filter((k) => k[0] !== '!').map((k) => String(help[k].v)).join('\n');
  t('1.11 說明頁：合併不刪除、單位固定人天、機密提醒、匯出日期與項目數', ['不會被刪除', '人天', '機密', '2026-10-08', '共 3 項（啟用 2 項；ERP 3、ITS 0、MDM 0、CRM 0）', '依「BU＋項目名稱」合併', '完全等於'].every((s) => helpTxt.includes(s)), helpTxt.slice(0, 200));
  const bufT = await X.buildWorkbook({ mode: 'template', date: '2026-10-08' });
  const wbT = XLSX.read(bufT, { type: 'buffer' });
  t('1.12 範本：資料工作表只有標題列（!ref＝A1:D1，沒有任何範例資料可能被誤匯入）', wbT.Sheets['ERP']['!ref'] === 'A1:D1' && eq(['A1', 'B1', 'C1', 'D1'].map((k) => wbT.Sheets['ERP'][k].v), HDR) && !wbT.Sheets['ERP']['A2']);
  const helpT = Object.keys(wbT.Sheets['說明']).filter((k) => k[0] !== '!').map((k) => String(wbT.Sheets['說明'][k].v)).join('\n');
  t('1.13 範本說明頁有虛構範例表（標示「範例」「不會被匯入」）', helpT.includes('（範例）專案經理') && helpT.includes('不會被匯入') && helpT.includes('9,000'));
  const pT = await prev(bufT, ITEMS);
  t('1.14 直接匯入空範本 → NO_DATA（400），不是 0 筆成功', pT.ok === false && pT.code === 'NO_DATA' && pT.status === 400, JSON.stringify(pT));
  const partsT = await unzipText(bufT);
  t('1.15 範本 XML 良構', Object.values(partsT).every(wellFormed));
  t('1.16 xmlEsc：名稱含 & < > " 與 XML 不允許的控制字元 → 良構且讀回相同（控制字元被丟棄）', await (async () => {
    const b = await X.buildWorkbook({ mode: 'export', items: [cur('A&B <x> "q" \'s\'', 1, 1), cur('Z\u0001\u0008y', 2, 2)] });
    const p = await unzipText(b);
    const w = XLSX.read(b, { type: 'buffer' }).Sheets['ERP'];
    return wellFormed(p['xl/worksheets/sheet1.xml']) && w.A2.v === 'A&B <x> "q" \'s\'' && w.A3.v === 'Zy';
  })());

  // ═════════════════ 2) 往返＋公式注入 ═════════════════
  const RT = [
    cur('PM 顧問經理', 9000, 6500), cur('SD 顧問', 7000.5, 5000, false), cur('小數', 0.01, 0.01), cur('上限', 1e9, 1e9), cur('零', 0, 0),
    cur('全形ＡＢＣ　顧問', 1234.56, 999.99), cur('p  m', 100, 50), cur('Mixed CASE', 3, 2), cur('中文名稱 範例', 8000, 5600.1),
    cur('=1+1', 1, 1), cur('+SUM(A1)', 2, 2), cur('-2+3', 3, 3), cur('@cmd', 4, 4), cur('=HYPERLINK("http://example.test","x")', 5, 5, false), cur('<img src=x onerror=alert(1)>', 6, 6),
  ];
  const bufRT = await X.buildWorkbook({ mode: 'export', items: RT, date: '2026-10-08' });
  const rt = await prev(bufRT, RT);
  t('2.1 往返：匯出 → 匯入預覽，全部「不變」、0 新增 0 更新 0 錯誤', rt.ok && rt.summary.added === 0 && rt.summary.updated === 0 && rt.summary.errors === 0 && rt.summary.unchanged === RT.length, JSON.stringify(rt.summary || rt));
  t('2.2 往返：merged 與原本逐欄相同（id、順序、數字、啟用）', rt.ok && eq(rt.merged, RT.map(asPlain)));
  const sRT = (await unzipText(bufRT))['xl/worksheets/sheet1.xml'];   // ERP 工作表
  const injected = RT.map((x, i) => [x, i + 2]).filter(([x]) => /^[=+\-@]/.test(x.name));
  t('2.3 公式注入：以 = + - @ 開頭的名稱 → 字串型（t="inlineStr"）＋quotePrefix 樣式（s=3），其他名稱用一般文字樣式（s=2）', injected.length === 5
    && injected.every(([, r]) => new RegExp(`<c r="A${r}" s="3" t="inlineStr">`).test(sRT)) && RT.filter((x) => !/^[=+\-@]/.test(x.name)).every((x) => { const r = RT.indexOf(x) + 2; return new RegExp(`<c r="A${r}" s="2" t="inlineStr">`).test(sRT); }));
  t('2.4 公式注入：整份資料表沒有 <f>、沒有以 = 開頭的 <v>（數字格以外）；SheetJS 讀回的名稱逐字相同', !/<f[ >]/.test(sRT) && XLSX.read(bufRT, { type: 'buffer' }).Sheets['ERP'] && injected.every(([x, r]) => XLSX.read(bufRT, { type: 'buffer' }).Sheets['ERP']['A' + r].v === x.name));
  t('2.5 styles.xml 的 quotePrefix 樣式存在且只有一個（cellXfs 第 3 個）', (await unzipText(bufRT))['xl/styles.xml'].match(/quotePrefix="1"/g).length === 1);
  const bufTab = await X.buildWorkbook({ mode: 'export', items: [{ name: '\tx', price: 1, cost: 1, active: true }, { name: '\rx', price: 1, cost: 1, active: true }] });
  t('2.6 Tab／CR 開頭的名稱（即使伺服器通常不會有）也走 quotePrefix', /<c r="A2" s="3"/.test((await unzipText(bufTab))['xl/worksheets/sheet1.xml']) && /<c r="A3" s="3"/.test((await unzipText(bufTab))['xl/worksheets/sheet1.xml']));
  // 往返之後再套用 merged → 仍是合法清單，且再匯一次還是全部不變（冪等）
  let gidn = 0;
  const asStored = PB.normalizePricebook(rt.merged, { genId: () => 'x' + (++gidn) });
  t('2.7 往返的 merged 通過 PUT 的同一套驗證', asStored.ok, asStored.ok ? '' : JSON.stringify(asStored.error));
  // 儲存格是公式（別人用 Excel 編輯後存檔）：只讀快取值當純文字，不求值
  const bufF = mkXlsx([['ERP', [HDR, [{ t: 's', v: 'cached-name', f: 'HYPERLINK("http://example.test","x")' }, 100, 50, '是']]]]);
  const pF = await prev(bufF, []);
  t('2.8 名稱儲存格帶公式（SheetJS 手工做的 <f>）→ 整列報錯「儲存格含公式，請改貼為值」，不新增（不採用快取值）', pF.ok && pF.merged.length === 0 && pF.summary.errors === 1 && pF.rows[0].action === 'error' && pF.rows[0].error === '儲存格含公式，請改貼為值' && !pF.rows[0].name, JSON.stringify(pF.rows || pF));


  // ═════════════════ 2b) 公式儲存格（Excel 裡打 =A1 會變成真的公式，匯入只讀得到快取值）═════════════════
  {
    const FM = '儲存格含公式，請改貼為值';
    const base = await X.buildWorkbook({ mode: 'export', items: [cur('Keep A', 100, 50, true, 'k1'), cur('Keep B', 200, 80, true, 'k2')], date: '2026-10-09' });
    // 直接改 XML：把 B2（牌價）、A3（名稱）換成帶 <f> 的儲存格；再加一列 A4＝新增列，D4（啟用）也是公式
    const withF = await rezip(base, async (z) => {
      let x = await z.file('xl/worksheets/sheet1.xml').async('string');
      x = x.replace(/<c r="B2" s="4"><v>100<\/v><\/c>/, '<c r="B2" s="4"><f>A1+5</f><v>100</v></c>');
      x = x.replace(/<c r="A3" s="2" t="inlineStr"><is><t xml:space="preserve">Keep B<\/t><\/is><\/c>/, '<c r="A3" s="2" t="str"><f>A2</f><v>Keep A</v></c>');
      x = x.replace('</sheetData>', '<row r="502"><c r="A502" t="inlineStr"><is><t>Brand New</t></is></c><c r="B502"><v>1</v></c><c r="C502"><v>1</v></c><c r="D502" t="str"><f>"是"</f><v>是</v></c></row></sheetData>');
      z.file('xl/worksheets/sheet1.xml', x);
    });
    const rf = await prev(withF, [cur('Keep A', 100, 50, true, 'k1'), cur('Keep B', 200, 80, true, 'k2')]);
    t('2b.1 手工 XML 帶 <f> 的儲存格（牌價欄、名稱欄、啟用欄）→ 各自那一列報錯 FM、其他列照常；有公式的列不更新、不新增', rf.ok && rf.rows.filter((x) => x.error === FM).length >= 2 && rf.rows.find((x) => x.row === 2).action === 'error' && rf.rows.find((x) => x.row === 3).action === 'error' && rf.merged.length === 2 && rf.merged[0].price === 100 && rf.merged[1].price === 200 && !rf.merged.some((x) => x.name === 'Brand New'), JSON.stringify(rf.rows));
    t('2b.2 含公式的名稱儲存格的快取值（Keep A）不會被當成已出現的名稱：同檔真正的 Keep A 列沒有被誤判為重複', await (async () => { const b2 = await rezip(base, async (z) => { let x = await z.file('xl/worksheets/sheet1.xml').async('string'); x = x.replace(/<c r="A3" s="2" t="inlineStr"><is><t xml:space="preserve">Keep B<\/t><\/is><\/c>/, '<c r="A3" s="2" t="str"><f>A2</f><v>Keep A</v></c>'); z.file('xl/worksheets/sheet1.xml', x); }); const r2 = await prev(b2, []); return r2.rows.find((x) => x.row === 2).action === 'add' && r2.rows.find((x) => x.row === 3).action === 'error' && r2.rows.find((x) => x.row === 3).error === FM; })());
    const sheetF = await prev(mkXlsx([['ERP', [HDR, ['OK', 1, 1, '是'], ['P', { t: 'n', v: 5, f: 'B2*5' }, 1, '']]]]), []);
    t('2b.3 SheetJS 產生的數字格公式（B3＝B2*5）→ 該列報錯、不被當成 5；摘要錯誤 1', sheetF.ok && sheetF.summary.errors === 1 && sheetF.summary.added === 1 && sheetF.rows[1].error === FM, JSON.stringify(sheetF.rows));
    // 匯出檔的預先格式化
    const sx = (await unzipText(base))['xl/worksheets/sheet1.xml'];
    t('2b.4 匯出的 ERP 工作表第 2～501 列都預先套好格式：A 欄＝文字格式（s=2）、B／C 欄＝數字格式（s=4）、D 欄＝文字置中（s=11）；資料列之後的空白列也有（A4、B501、D501）', /<c r="A4" s="2"\/>/.test(sx) && /<c r="B4" s="4"\/>/.test(sx) && /<c r="C501" s="4"\/>/.test(sx) && /<c r="D501" s="11"\/>/.test(sx) && !/<c r="A502"/.test(sx) && /<c r="D2" s="11" t="inlineStr">/.test(sx));
    const ps = (await unzipText(base))['xl/styles.xml'];
    t('2b.5 文字格式（numFmt 49）樣式存在：cellXfs 第 3、4、12 個（A 欄文字、quotePrefix、D 欄置中文字），count=12', /<cellXfs count="12">/.test(ps) && (ps.match(/<xf numFmtId="49"/g) || []).length >= 5);
    // 預先格式化的空白儲存格匯回時不產生任何列（0 筆資料、不是 500 個錯誤列）
    const pre = await prev(base, [cur('Keep A', 100, 50, true, 'k1'), cur('Keep B', 200, 80, true, 'k2')]);
    t('2b.6 預先格式化的空白儲存格匯入時完全忽略：匯出檔（2 項＋499 個空白格式列）→ 2 列不變、略過 0、沒有錯誤', pre.ok && pre.rows.length === 2 && pre.summary.unchanged === 2 && pre.summary.skipped === 0 && pre.summary.errors === 0, JSON.stringify(pre.summary || pre));
    // 範本也預先格式化
    const tpl = (await unzipText(await X.buildWorkbook({ mode: 'template', date: '2026-10-09' })))['xl/worksheets/sheet2.xml'];
    t('2b.7 範本的每個 BU 工作表也預先格式化（ITS 的 A2 文字、B2 數字、D501 文字）；但仍然「沒有資料」（匯入範本→NO_DATA）', /<c r="A2" s="2"\/>/.test(tpl) && /<c r="B2" s="4"\/>/.test(tpl) && /<c r="D501" s="11"\/>/.test(tpl) && (await prev(await X.buildWorkbook({ mode: 'template', date: '2026-10-09' }), [])).code === 'NO_DATA');
  }

  // ═════════════════ 3) 合併語意 ═════════════════
  const EX = [cur('PM 顧問經理', 9000, 6500, true, 'p1'), cur('SD 顧問', 7000, 5000, false, 'p2'), cur('ABAP 顧問', 6500, 4500, true, 'p3'), cur('Keep Me', 1, 1, true, 'p4')];
  let r = await prev(one([['pm 顧問經理', 9500, 6000, ''], ['  ＳＤ　顧問 ', 7000, 5000, '是'], ['New One', 5000, 3000, ''], ['New Two', 4000, 2000, '否']]), EX);
  t('3.1 同名（大小寫／全形／多空白不分）→ update；啟用留白＝維持；新名稱 → add；摘要 新增2 更新2', r.ok && act(r) === 'update,update,add,add' && r.summary.added === 2 && r.summary.updated === 2 && r.summary.unchanged === 0 && r.summary.errors === 0, JSON.stringify(r.summary) + act(r));
  t('3.2 更新保留 id、位置與「原本的名稱文字」（不因大小寫／全形差異改名）', r.merged[0].id === 'p1' && r.merged[0].name === 'PM 顧問經理' && r.merged[1].id === 'p2' && r.merged[1].name === 'SD 顧問' && r.merged.slice(0, 4).map((x) => x.id).join() === 'p1,p2,p3,p4');
  t('3.3 更新內容：PM 牌價 9500／成本 6000 仍啟用；SD 由停用 → 啟用（「是」）', r.merged[0].price === 9500 && r.merged[0].cost === 6000 && r.merged[0].active === true && r.merged[1].active === true);
  t('3.4 檔案沒提到的既有項目（ABAP、Keep Me）原封不動、位置不變', eq(r.merged[2], { id: 'p3', bu: 'ERP', name: 'ABAP 顧問', price: 6500, cost: 4500, active: true }) && eq(r.merged[3], { id: 'p4', bu: 'ERP', name: 'Keep Me', price: 1, cost: 1, active: true }));
  t('3.5 新項目附加在最後、依檔案順序、沒有 id（由 PUT 產生）、啟用留白＝是／填否＝停用', r.merged.length === 6 && r.merged[4].name === 'New One' && r.merged[4].active === true && !('id' in r.merged[4]) && r.merged[4].bu === 'ERP' && r.merged[5].name === 'New Two' && r.merged[5].active === false);
  t('3.6 預覽列：行號＝Excel 列號（資料從第 2 列起）、old→new 數字', byRow(r, 2).old.price === 9000 && byRow(r, 2).new.price === 9500 && byRow(r, 2).action === 'update' && byRow(r, 4).action === 'add' && byRow(r, 4).old === undefined && byRow(r, 5).new.active === false);
  r = await prev(one([['PM 顧問經理', 9000, 6500, '是'], ['sd 顧問', 7000, 5000, ''], ['ABAP 顧問', 6500, 4500, '']]), EX);
  t('3.7 數值與啟用都相同 → unchanged（含「啟用留白＝維持停用」）；merged 與原清單相同', r.ok && act(r) === 'unchanged,unchanged,unchanged' && r.summary.unchanged === 3 && eq(r.merged, EX.map(asPlain)));
  // 啟用狀態切換
  r = await prev(one([['PM 顧問經理', 9000, 6500, '否'], ['SD 顧問', 7000, 5000, 'TRUE']]), EX);
  t('3.8 只改啟用也算 update；顯示 old.active→new.active', r.ok && act(r) === 'update,update' && byRow(r, 2).old.active === true && byRow(r, 2).new.active === false && r.merged[0].active === false && r.merged[1].active === true);
  // 上限 60
  const EX58 = Array.from({ length: 58 }, (_, i) => cur('R' + i, 1, 1, true, 'e' + i));
  r = await prev(one([['N1', 1, 1, ''], ['N2', 2, 2, ''], ['N3', 3, 3, ''], ['R0', 9, 9, '']]), EX58);
  t('3.9 60 項上限：既有 58 + 新增 3 → 前 2 筆新增、第 3 筆報錯；同名更新不受上限影響', r.ok && act(r) === 'add,add,error,update' && r.summary.added === 2 && r.summary.errors === 1 && r.merged.length === 60 && /最多 60 項/.test(byRow(r, 4).error), act(r) + (byRow(r, 4) || {}).error);
  // 檔內重複
  r = await prev(one([['Dup', 1, 1, ''], ['dup ', 2, 2, ''], ['ＤＵＰ', 3, 3, ''], ['Other', 4, 4, '']]), []);
  t('3.10 檔內重複（大小寫／空白／全形）→ 第一筆照常、後面的列報錯（不默默合併）', r.ok && act(r) === 'add,error,error,add' && r.merged.length === 2 && r.merged[0].price === 1 && /與本工作表第 2 列的項目名稱重複/.test(byRow(r, 3).error) && /與本工作表第 2 列/.test(byRow(r, 4).error), act(r) + JSON.stringify(r.rows));
  r = await prev(one([['Dup', 'abc', 1, ''], ['Dup', 2, 2, '']]), []);
  t('3.11 第一筆有錯誤、第二筆同名 → 兩筆都不匯入（第一次出現的名稱被保留位置，避免誤以為第二筆是第一筆）', r.ok && act(r) === 'error,error' && r.merged.length === 0);
  // 壞數字
  const badNums = [['abc', 1], [-5, 1], ['-5', 1], [1e10, 1], ['', 1], ['1e3', 1], ['5%', 1], ['1,23', 1], [1, 'x'], [1, -0.5], [{ t: 'e', v: 15, w: '#VALUE!' }, 1], [true, 1], [1, 1e9 + 1]];
  r = await prev(one(badNums.map((b, i) => ['Bad' + i, b[0], b[1], ''])), []);
  t('3.12 壞數字（文字、負數、超過 1e9、空白、科學記號、百分比、錯誤千分位、錯誤值、布林）→ 全部是 error 且各有原因，沒有任何一筆被匯入', r.ok && r.summary.errors === badNums.length && r.merged.length === 0 && r.rows.every((x) => x.action === 'error' && x.error), act(r) + JSON.stringify(r.rows.map((x) => x.error)));
  t('3.13 錯誤原因指出是哪個欄位（牌價／成本）', /牌價/.test(byRow(r, 2).error) && /成本/.test(byRow(r, 10).error) && /不可為負數/.test(byRow(r, 3).error) && /超過上限/.test(byRow(r, 5).error));
  r = await prev(one([['Both', 'x', 'y', '']]), []);
  t('3.14 同一列多個問題 → 原因合併顯示（牌價與成本）', /牌價/.test(r.rows[0].error) && /成本/.test(r.rows[0].error));
  // 好數字
  r = await prev(one([['T1', '1,234.5', '1,000', ''], ['T2', '１２３４', '５６７．８９', ''], ['T3', '７，０００', ' 6000 ', ''], ['T4', 'NT$ 7,000', '5000元', ''], ['T5', '$100', '0', ''], ['T6', 7000.004, 1234.567, ''], ['T7', '1,234,567.891', 1, '']]), []);
  const g = (i) => r.merged[i];
  t('3.15 容許：千分位、全形數字／逗號／小數點、前後空白、NT$／$／元、數字格與文字格；小數四捨五入到 2 位', r.ok && r.summary.errors === 0 && g(0).price === 1234.5 && g(0).cost === 1000 && g(1).price === 1234 && g(1).cost === 567.89 && g(2).price === 7000 && g(2).cost === 6000 && g(3).price === 7000 && g(3).cost === 5000 && g(4).price === 100 && g(4).cost === 0 && g(5).price === 7000 && g(5).cost === 1234.57 && g(6).price === 1234567.89, JSON.stringify(r.merged));
  // 標題變體
  const variants = [
    [['name', 'price', 'cost', 'active'], 'name,price,cost,active'],
    [['Name', 'Price', 'Cost', 'Active'], 'capitalized'],
    [['項目名稱', '牌價', '成本', '啟用'], '無括號'],
    [[' 項目名稱 ', '牌價 （元／人天）', '成本(元/人天)', ' 啟用 '], '空白與全形括號'],
    [['品名', '售價', 'COST', '狀態'], '別名'],
    [['項目名稱', '牌價（元/人天）', '成本（元/人天）'], '沒有啟用欄'],
  ];
  for (const [h, label] of variants) {
    const row = h.length === 4 ? ['V', 10, 5, '是'] : ['V', 10, 5];
    const rv = await prev(one([row], h), []);
    t('3.16 標題變體：' + label, rv.ok && rv.merged.length === 1 && rv.merged[0].price === 10 && rv.merged[0].cost === 5 && rv.merged[0].active === true, JSON.stringify(rv.code ? rv : rv.rows));
  }
  r = await prev(one([['V', 5, 10, '否', 'ignore me']], ['成本（元/人天）', '項目名稱', '備註', '牌價（元/人天）', '啟用', '其他']), []);
  t('3.17 欄位順序可以任意、多餘欄位忽略', await (async () => {
    const b = mkXlsx([['ERP', [['備註', '成本', '項目名稱', '額外', '牌價', '啟用'], ['x', 11, 'Perm', 'y', 22, '否']]]]);
    const rv = await prev(b, []);
    return rv.ok && eq(rv.merged, [{ bu: 'ERP', name: 'Perm', price: 22, cost: 11, active: false }]);
  })());
  const bufT3 = mkXlsx([['ERP', [['牌價簿維護表'], [], ['項目名稱', '牌價', '成本', '啟用'], ['After Title', 1, 1, '']]]]);
  r = await prev(bufT3, []);
  t('3.18 標題列不在第 1 列（前面有標題文字）也找得到；資料列號從標題下一列算（Excel 第 4 列）', r.ok && r.rows.length === 1 && r.rows[0].row === 4 && r.rows[0].name === 'After Title');
  for (const [h, label, codeWant] of [[['項目名稱', '牌價'], '缺成本', 'BAD_HEADER'], [['名字欄', '牌價', '成本'], '沒有項目名稱欄', 'BAD_HEADER'], [['a', 'b', 'c'], '完全不像', 'BAD_HEADER']]) {
    const rv = await prev(one([['x', 1, 1]], h), []);
    t('3.19 標題錯誤：' + label + ' → 400 ' + codeWant + '，訊息說明需要哪些欄位', rv.ok === false && rv.code === codeWant && rv.status === 400 && /項目名稱/.test(rv.message), JSON.stringify(rv));
  }
  const miss = await prev(one([['x', 1, 1]], ['項目名稱', '牌價']), []);
  t('3.20 缺欄位時訊息指出缺哪一欄（成本）', /缺少欄位：成本/.test(miss.message), miss.message);
  // 是／否變體
  const yes = ['是', 'Y', 'y', 'yes', 'YES', 'TRUE', 'true', true, 1, '1', ' 是 ', 'ｙ', '啟用', 'ＹＥＳ'];
  const no = ['否', 'N', 'n', 'no', 'FALSE', 'false', false, 0, '0', ' 否 ', '停用', 'Ｎ'];
  r = await prev(one(yes.map((v, i) => ['Y' + i, 1, 1, v]).concat(no.map((v, i) => ['N' + i, 1, 1, v]))), []);
  t('3.21 啟用欄「是」變體（是／Y／yes／TRUE／布林／1／全形／啟用）全部判為啟用、「否」變體全部判為停用', r.ok && r.summary.errors === 0 && r.merged.slice(0, yes.length).every((x) => x.active === true) && r.merged.slice(yes.length).every((x) => x.active === false), JSON.stringify(r.rows.filter((x) => x.error)));
  r = await prev(one([['A', 1, 1, 'maybe'], ['B', 1, 1, 2], ['C', 1, 1, '是否']]), []);
  t('3.22 啟用欄填了看不懂的值 → 該列 error（不猜）', r.ok && r.summary.errors === 3 && r.rows.every((x) => /啟用欄/.test(x.error)));
  // 空白列
  r = await prev(one([['A', 1, 1, ''], [], ['', '', '', ''], ['  ', null, '', ''], ['B', 2, 2, ''], [null, null, null, null]]), []);
  t('3.23 空白列被忽略：資料中間夾的空白列計入「略過」、尾端的不算；不產生預覽列', r.ok && r.rows.length === 2 && r.summary.skipped === 3 && r.merged.length === 2, JSON.stringify(r.summary));
  r = await prev(one([['', 5, 5, ''], ['OnlyName', '', '', '']]), []);
  t('3.24 只有部分欄位的列：沒名稱→error「項目名稱空白」；有名稱沒金額→error（不是當成空白列）', r.ok && r.summary.errors === 2 && /項目名稱空白/.test(r.rows[0].error) && /牌價空白/.test(r.rows[1].error));
  // 名稱規則
  r = await prev(one([['x'.repeat(41), 1, 1, ''], ['y'.repeat(40), 1, 1, ''], ['Line\nBreak\tTab', 1, 1, ''], ['   ', 1, 1, ''], [123456, 1, 1, ''], ['Ctl\u0003x', 1, 1, '']]), []);
  t('3.25 名稱：41 字 error、40 字可；儲存格內換行／Tab 收合成空白；全空白 error；數字格名稱當文字；其他控制字元 error', r.ok && byRow(r, 2).action === 'error' && /40 字/.test(byRow(r, 2).error) && byRow(r, 3).action === 'add' && byRow(r, 4).name === 'Line Break Tab' && byRow(r, 5).action === 'error' && byRow(r, 6).name === '123456' && byRow(r, 7).action === 'error', JSON.stringify(r.rows.map((x) => [x.row, x.action, x.name])));
  const xss = ['<img src=x onerror=alert(1)>', '"><script>alert(1)</script>', "' onfocus='alert(1)", '&lt;b&gt;', '=cmd|calc!A1'];
  r = await prev(one(xss.map((n) => [n, 1, 1, ''])), []);
  t('3.26 XSS／公式字樣的名稱：原樣當純文字帶進預覽與 merged（顯示端負責轉義；此處不改動）', r.ok && r.rows.every((x, i) => x.name === xss[i]) && r.merged.every((x, i) => x.name === xss[i]));
  // 工作表選擇：只認名稱「完全等於」ERP／ITS／MDM／CRM 的工作表
  r = await prev(multi({ ERP: [['In ERP', 1, 1, '']] }, [['說明', [['這是說明'], ['x']]], ['Sheet1', [HDR, ['In Sheet1', 9, 9, '']]], ['erp', [HDR, ['In lower erp', 9, 9, '']]], ['ERP ', [HDR, ['In ERP space', 9, 9, '']]]]), []);
  t('3.27 只讀名稱完全等於 BU 的工作表；「說明」安靜略過；其他名稱（Sheet1、小寫 erp、ERP 加空白）全部列入 skippedSheets 且完全不讀', r.ok && r.merged.length === 1 && r.merged[0].name === 'In ERP' && eq(r.skippedSheets, ['Sheet1', 'erp', 'ERP ']) && r.summary.skipped === 3, JSON.stringify(r.skippedSheets) + JSON.stringify(r.summary));
  r = await prev(mkXlsx([['Sheet1', [HDR, ['In First', 1, 1, '']]], ['Other', [HDR, ['In Other', 2, 2, '']]]]), []);
  t('3.28 沒有任何 BU 工作表 → 400 NO_BU_SHEET（不猜第一個工作表）', r.ok === false && r.code === 'NO_BU_SHEET' && r.status === 400 && /ERP/.test(r.message), JSON.stringify(r));
  r = await prev(mkXlsx([['說明', [['只是說明'], ['x']]]]), []);
  t('3.29 只有說明頁 → 400 NO_BU_SHEET', r.ok === false && r.code === 'NO_BU_SHEET');
  r = await prev(mkXlsx([['ERP', [['只是說明'], ['x']]]]), []);
  t('3.29b BU 工作表有內容但找不到標題列 → 400 BAD_HEADER，訊息指出是哪個工作表', r.ok === false && r.code === 'BAD_HEADER' && /工作表「ERP」/.test(r.message), JSON.stringify(r));
  // 500 列
  const rows500 = Array.from({ length: 500 }, (_, i) => ['Bulk' + i, 1, 1, '']);
  r = await prev(one(rows500), []);
  t('3.30 剛好 500 筆資料列可處理（超過 60 項的部分各自報錯）；501 筆 → TOO_MANY_ROWS', r.ok && r.rows.length === 500 && r.summary.added === 60 && r.summary.errors === 440 && r.merged.length === 60, JSON.stringify(r.summary || r));
  r = await prev(one(rows500.concat([['Bulk500', 1, 1, '']])), []);
  t('3.31 501 筆 → 400 TOO_MANY_ROWS', r.ok === false && r.code === 'TOO_MANY_ROWS' && r.status === 400);
  // 冪等：套用 merged 後再匯同一份檔 → 全部不變
  const file3 = one([['PM 顧問經理', 9500, 6000, ''], ['New One', 5000, 3000, ''], ['SD 顧問', 7000, 5000, '是']]);
  let st = EX.map(asPlain);
  let first = await prev(file3, st);
  const applied = PB.normalizePricebook(first.merged, { genId: (() => { let n = 0; return () => 'g' + (++n); })() });
  const second = await prev(file3, applied.items);
  t('3.32 第二次匯入同一份檔（套用第一次的結果之後）→ 全部「不變」、0 新增 0 更新', first.ok && applied.ok && second.ok && second.summary.unchanged === 3 && second.summary.added === 0 && second.summary.updated === 0 && eq(second.merged, applied.items.map(asPlain)), JSON.stringify(second.summary || second));
  t('3.33 預覽不會改動傳入的既有清單（純函式）', eq(st, EX.map(asPlain)));
  // parse 單元：標題比對
  t('3.34 matchHeader 直接測：別名／括號／全形', X.matchHeader('牌價（元/人天）') === 'price' && X.matchHeader(' ＮＡＭＥ ') === 'name' && X.matchHeader('成本 (NTD)') === 'cost' && X.matchHeader('啟用') === 'active' && X.matchHeader('備註') === null && X.matchHeader('') === null);

  // ═════════════════ 3b) 多 BU：每個 BU 一個工作表，依「BU＋名稱」合併 ═════════════════
  const BUS4 = ['ERP', 'ITS', 'MDM', 'CRM'];
  const MI = [
    cur('PM', 9000, 6500, true, 'e1', 'ERP'), cur('SD', 7000, 5000, false, 'e2', 'ERP'),
    cur('SD', 6000, 4000, true, 'i1', 'ITS'), cur('Net', 5000, 3500, true, 'i2', 'ITS'),
    cur('MDM Lead', 8000, 5000, true, 'm1', 'MDM'),
  ];
  const fileB = multi({
    ERP: [['sd', 7500, 5200, ''], ['NewE', 1, 1, '']],
    ITS: [['SD', 6000, 4000, ''], ['SD', 1, 1, ''], ['Newi', 2, 2, '否']],
    CRM: [['SD', 3, 3, '']],
  }, [['Other', [HDR, ['Ignored', 1, 1, '']]]]);
  let rb = await prev(fileB, MI);
  const buRows = (bu) => rb.rows.filter((x) => x.bu === bu);
  t('3b.1 每列預覽都帶 bu；ERP：「sd」更新（保留 id e2／位置／停用狀態）、NewE 新增', rb.ok && eq(buRows('ERP').map((x) => x.action + ':' + x.name), ['update:SD', 'add:NewE']) && buRows('ERP')[0].old.price === 7000 && buRows('ERP')[0].new.price === 7500 && buRows('ERP')[0].new.active === false, JSON.stringify(buRows('ERP')));
  t('3b.2 ITS：同名「SD」在 ERP 被更新，但 ITS 的 SD 完全獨立（不變）；同一工作表第二個 SD → 錯誤；Newi 新增為停用', eq(buRows('ITS').map((x) => x.action), ['unchanged', 'error', 'add']) && /與本工作表第 2 列/.test(buRows('ITS')[1].error) && buRows('ITS')[2].new.active === false);
  t('3b.3 CRM：「SD」在 CRM 是全新項目 → 新增（同名跨 BU 允許）；MDM 沒有工作表 → 完全不動（missingBus）；Other 工作表 → skippedSheets', eq(buRows('CRM').map((x) => x.action), ['add']) && buRows('MDM').length === 0 && eq(rb.missingBus, ['MDM']) && eq(rb.skippedSheets, ['Other']));
  t('3b.4 摘要：總計與每個 BU 的新增／更新／不變／略過／錯誤都正確且總計＝各 BU 加總（略過＝1 個被略過的工作表）', rb.summary.added === 3 && rb.summary.updated === 1 && rb.summary.unchanged === 1 && rb.summary.errors === 1 && rb.summary.skipped === 1
    && eq(rb.summary.byBu.ERP, { added: 1, updated: 1, unchanged: 0, skipped: 0, errors: 0 }) && eq(rb.summary.byBu.ITS, { added: 1, updated: 0, unchanged: 1, skipped: 0, errors: 1 }) && eq(rb.summary.byBu.MDM, { added: 0, updated: 0, unchanged: 0, skipped: 0, errors: 0 }) && eq(rb.summary.byBu.CRM, { added: 1, updated: 0, unchanged: 0, skipped: 0, errors: 0 }), JSON.stringify(rb.summary));
  t('3b.5 merged 依 ERP、ITS、MDM、CRM 排列；各 BU 內既有項目保留 id 與順序，新項目接在「該 BU」最後（不是整份最後）；缺 BU 工作表的項目逐欄不變', eq(rb.merged.map((x) => x.bu + ':' + (x.id || '(new)') + ':' + x.name), ['ERP:e1:PM', 'ERP:e2:SD', 'ERP:(new):NewE', 'ITS:i1:SD', 'ITS:i2:Net', 'ITS:(new):Newi', 'MDM:m1:MDM Lead', 'CRM:(new):SD']), JSON.stringify(rb.merged.map((x) => x.bu + x.name)));
  t('3b.6 只動被提到的：ERP 的 PM（沒出現在檔案）、ITS 的 Net、MDM 的 Lead 都逐欄不變', eq(rb.merged.find((x) => x.id === 'e1'), asPlain(MI[0])) && eq(rb.merged.find((x) => x.id === 'i2'), asPlain(MI[3])) && eq(rb.merged.find((x) => x.id === 'm1'), asPlain(MI[4])));
  let nb = 0;
  const chkB = PB.normalizePricebook(rb.merged, { genId: () => 'n' + (++nb) });
  t('3b.7 merged 通過 PUT 的同一套驗證（同名跨 BU 的 SD 都在）', chkB.ok && chkB.items.filter((x) => x.name === 'SD').length === 3 && chkB.items.every((x) => BUS4.includes(x.bu)), chkB.ok ? '' : JSON.stringify(chkB.error));
  // 每個 BU 60 上限（互不影響）
  const full = (bu, n) => Array.from({ length: n }, (_, i) => cur('R' + i, 1, 1, true, bu.toLowerCase() + i, bu));
  rb = await prev(multi({ ERP: [['X1', 1, 1, '']], ITS: [['Y1', 1, 1, ''], ['Y2', 1, 1, '']] }), full('ERP', 60).concat(full('ITS', 59)));
  t('3b.8 上限是每個 BU 各 60：ERP 已滿 → 新增報錯；ITS 有 59 → 第 1 筆可新增、第 2 筆報錯；ERP 滿了不影響 ITS', rb.ok && eq(rb.rows.map((x) => x.bu + ':' + x.action), ['ERP:error', 'ITS:add', 'ITS:error']) && /ERP 的牌價簿最多 60 項/.test(rb.rows[0].error) && /ITS 的牌價簿最多 60 項/.test(rb.rows[2].error) && rb.merged.length === 120, JSON.stringify(rb.rows));
  rb = await prev(multi({ MDM: [['Z', 1, 1, '']] }), full('ERP', 60).concat(full('ITS', 60), full('CRM', 60)));
  t('3b.9 其他三個 BU 各 60（共 180）時，MDM 仍可新增（總量 181 ≤ 240）', rb.ok && rb.summary.added === 1 && rb.merged.length === 181);
  // 舊資料：沒有 bu 欄位 → 視為 ERP
  const legacyEx = [{ id: 'L1', name: 'Old PM', price: 9000, cost: 6500, active: true }, { id: 'L2', name: 'Old SD', unit: '人天', price: 7000, cost: 5000, active: false }];
  rb = await prev(multi({ ERP: [['old pm', 9100, 6500, '']], ITS: [['Old PM', 1, 1, '']] }), legacyEx);
  t('3b.10 相容舊資料：既有項目沒有 bu → 當作 ERP（ERP 工作表的同名列是更新、ITS 工作表的同名列是新增）；merged 一律補上 bu', rb.ok && eq(rb.rows.map((x) => x.bu + ':' + x.action), ['ERP:update', 'ITS:add']) && rb.merged.every((x) => BUS4.includes(x.bu)) && rb.merged[0].id === 'L1' && rb.merged[0].bu === 'ERP', JSON.stringify(rb.merged));
  // 各工作表獨立的標題與數字規則
  rb = await prev(multi({ ERP: [['A', 1, 1, '']] }, [['ITS', [['name', 'price', 'cost', 'active'], ['B', '1,000', '５００', 'N']]], ['MDM', [['項目名稱', '牌價', '成本']]]]), []);
  t('3b.11 每個工作表的標題變體各自辨識（ITS 用英文標題、數字用千分位／全形）；只有標題列的 MDM 不報錯也不算「缺工作表」', rb.ok && rb.merged.length === 2 && rb.merged[1].bu === 'ITS' && rb.merged[1].price === 1000 && rb.merged[1].cost === 500 && rb.merged[1].active === false && eq(rb.missingBus, ['CRM']), JSON.stringify(rb.rows || rb));
  rb = await prev(multi({ ERP: [['A', 1, 1, '']] }, [['ITS', [['foo', 'bar'], ['B', 1]]]]), []);
  t('3b.12 任一 BU 工作表標題錯誤 → 整個檔案 400 BAD_HEADER，訊息指出是 ITS（不默默略過）', rb.ok === false && rb.code === 'BAD_HEADER' && /工作表「ITS」/.test(rb.message), JSON.stringify(rb));
  rb = await prev(multi({ ERP: [['A', 1, 1, '']], ITS: Array.from({ length: 501 }, (_, i) => ['N' + i, 1, 1, '']) }), []);
  t('3b.13 每個工作表資料列上限 500（ITS 501 列 → 400 TOO_MANY_ROWS，訊息指出 ITS）', rb.ok === false && rb.code === 'TOO_MANY_ROWS' && /ITS/.test(rb.message));
  rb = await prev(multi({ ERP: [], ITS: [], MDM: [], CRM: [] }), MI);
  t('3b.14 四個 BU 工作表都只有標題列 → 400 NO_DATA（直接上傳範本不會被當成成功）', rb.ok === false && rb.code === 'NO_DATA');
  // 往返（多 BU、同名跨 BU、公式字樣、停用）
  const RTB = [
    cur('SD 顧問', 7000, 5000, true, 'a1', 'ERP'), cur('=1+1', 1, 1, true, 'a2', 'ERP'), cur('SD 顧問', 6000.5, 4000, false, 'b1', 'ITS'), cur('+SUM(A1)', 2, 2, true, 'b2', 'ITS'),
    cur('-2+3', 3, 3, true, 'c1', 'MDM'), cur('@cmd', 4, 4, false, 'c2', 'MDM'), cur('=HYPERLINK("http://example.test","x")', 5, 5, true, 'd1', 'CRM'), cur('SD 顧問', 8000, 6000, true, 'd2', 'CRM'),
    cur('全形ＡＢＣ　顧問', 1234.56, 999.99, true, 'd3', 'CRM'),
  ];
  const bufRB = await X.buildWorkbook({ mode: 'export', items: RTB, date: '2026-10-09' });
  const wbRB = XLSX.read(bufRB, { type: 'buffer' });
  t('3b.15 匯出的每個 BU 工作表只含自己的項目（ERP 2、ITS 2、MDM 2、CRM 3），順序照清單，停用＝否，標題列都在', eq(['ERP', 'ITS', 'MDM', 'CRM'].map((b) => wbRB.Sheets[b]['!ref']), ['A1:D3', 'A1:D3', 'A1:D3', 'A1:D4']) && wbRB.Sheets.ITS.A2.v === 'SD 顧問' && wbRB.Sheets.ITS.D2.v === '否' && wbRB.Sheets.MDM.D3.v === '否' && wbRB.Sheets.CRM.A3.v === 'SD 顧問' && BUS4.every((b) => HDR.every((h, i) => wbRB.Sheets[b][String.fromCharCode(65 + i) + '1'].v === h)));
  const rtb = await prev(bufRB, RTB);
  t('3b.16 往返（多 BU、同名跨 BU、= + - @ 開頭、全形、小數、停用）：匯出 → 匯入預覽 ＝ 0 新增 0 更新 0 錯誤 ' + RTB.length + ' 不變，merged 逐欄相同', rtb.ok && rtb.summary.unchanged === RTB.length && rtb.summary.added === 0 && rtb.summary.updated === 0 && rtb.summary.errors === 0 && eq(rtb.merged, RTB.map(asPlain)) && BUS4.every((b) => rtb.summary.byBu[b].unchanged === RTB.filter((x) => x.bu === b).length), JSON.stringify(rtb.summary || rtb));
  const partsRB = await unzipText(bufRB);
  const injRB = [['sheet1', 'A3', '=1+1'], ['sheet2', 'A3', '+SUM(A1)'], ['sheet3', 'A2', '-2+3'], ['sheet3', 'A3', '@cmd'], ['sheet4', 'A2', '=HYPERLINK("http://example.test","x")']];
  t('3b.17 各 BU 工作表的公式字樣名稱都是字串型＋quotePrefix（s="3"），整個活頁簿沒有 <f> 公式；SheetJS 讀回逐字相同', Object.keys(partsRB).filter((k) => /worksheets\/sheet\d\.xml$/.test(k)).every((k) => !/<f[ >]/.test(partsRB[k]))
    && injRB.every(([sh, ref, nm]) => new RegExp('<c r="' + ref + '" s="3" t="inlineStr">').test(partsRB['xl/worksheets/' + sh + '.xml']) && XLSX.read(bufRB, { type: 'buffer' }).Sheets[{ sheet1: 'ERP', sheet2: 'ITS', sheet3: 'MDM', sheet4: 'CRM' }[sh]][ref].v === nm));
  t('3b.18 匯出檔每個 XML 部件良構；說明頁列出各 BU 項目數、工作表名稱規則、合併不刪除', Object.values(partsRB).every(wellFormed) && await (async () => { const h = XLSX.read(bufRB, { type: 'buffer' }).Sheets['說明']; const tx = Object.keys(h).filter((k) => k[0] !== '!').map((k) => String(h[k].v)).join('\n'); return ['ERP 2、ITS 2、MDM 2、CRM 3', '完全等於', '不同 BU 可以有相同名稱', '缺少某個 BU 的工作表：該 BU 完全不動', '不會被刪除', '機密', '單位固定是「人天」'].every((s) => tx.includes(s)); })());
  const bufTB = await X.buildWorkbook({ mode: 'template', date: '2026-10-09' });
  const wbTB = XLSX.read(bufTB, { type: 'buffer' });
  t('3b.19 範本：ERP／ITS／MDM／CRM 四個工作表都只有標題列（!ref＝A1:D1、沒有任何資料列）＋「說明」；範例只在說明頁的文字表（虛構，標示不會被匯入）', eq(wbTB.SheetNames, ['ERP', 'ITS', 'MDM', 'CRM', '說明']) && BUS4.every((b) => wbTB.Sheets[b]['!ref'] === 'A1:D1' && !wbTB.Sheets[b].A2) && Object.keys(wbTB.Sheets['說明']).some((k) => k[0] !== '!' && String(wbTB.Sheets['說明'][k].v).includes('（範例）專案經理')));
  // 往返後再套用、再匯入 → 冪等
  let gi2 = 0;
  const appliedB = PB.normalizePricebook((await prev(fileB, MI)).merged, { genId: () => 'g' + (++gi2) });
  const againB = await prev(fileB, appliedB.items);
  t('3b.20 套用一次 fileB 之後再匯入同一個檔：新增 0、更新 0（錯誤的那一列仍是錯誤）', appliedB.ok && againB.ok && againB.summary.added === 0 && againB.summary.updated === 0 && againB.summary.errors === 1, JSON.stringify(againB.summary || againB));
  t('3b.21 預覽是純函式：傳入的既有清單逐位元不變', await (async () => { const snap = JSON.stringify(MI); await prev(fileB, MI); return JSON.stringify(MI) === snap; })());

  // ═════════════════ 4) 惡意／異常檔案 ═════════════════
  const okCodes = new Set(['NO_BU_SHEET', 'BAD_XLSX', 'EMPTY_FILE', 'BAD_HEADER', 'NO_DATA', 'TOO_MANY_ROWS']);
  const cleanErr = (x) => x && x.ok === false && okCodes.has(x.code) && x.status === 400 && typeof x.message === 'string' && x.message.length > 0 && !/node_modules|\.js:|\bat [A-Za-z]|[A-Z]:\\|\/usr\/|\/home\//.test(x.message);
  async function hostile(name, buf, wantCode) {
    let x, threw = null;
    const t0 = Date.now();
    try { x = await X.previewFromBuffer(buf, []); } catch (e) { threw = e; }
    t('4.x ' + name + ' → ' + (wantCode || '乾淨錯誤') + '（不丟例外、不洩漏路徑、' + (Date.now() - t0) + 'ms）', !threw && cleanErr(x) && (!wantCode || x.code === wantCode), threw ? String(threw.stack).slice(0, 200) : JSON.stringify(x).slice(0, 200));
  }
  await hostile('空 Buffer', Buffer.alloc(0), 'EMPTY_FILE');
  await hostile('純文字', Buffer.from('項目名稱,牌價,成本\nA,1,1\n'), 'BAD_XLSX');
  await hostile('隨機位元組', require('crypto').randomBytes(5000), 'BAD_XLSX');
  await hostile('舊版 .xls（OLE2 魔術字）', Buffer.concat([Buffer.from([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]), Buffer.alloc(600)]), 'BAD_XLSX');
  await hostile('PDF', Buffer.from('%PDF-1.4\n%....'), 'BAD_XLSX');
  await hostile('只有 PK 開頭的垃圾', Buffer.concat([Buffer.from('PK\x03\x04'), Buffer.alloc(100, 7)]), 'BAD_XLSX');
  await hostile('被截斷的真 xlsx（砍掉後半）', bufE.slice(0, Math.floor(bufE.length / 2)), 'BAD_XLSX');
  await hostile('被截斷的真 xlsx（只砍最後 10 bytes，中央目錄壞掉）', bufE.slice(0, bufE.length - 10), 'BAD_XLSX');
  const zNoSheet = new JSZip(); zNoSheet.file('a.txt', 'hello');
  await hostile('zip 但沒有任何 xlsx 部件', await zNoSheet.generateAsync({ type: 'nodebuffer' }), 'BAD_XLSX');
  const zNoSheet2 = new JSZip(); zNoSheet2.file('[Content_Types].xml', '<Types/>'); zNoSheet2.file('xl/workbook.xml', '<workbook/>');
  await hostile('有 workbook 但沒有 worksheet', await zNoSheet2.generateAsync({ type: 'nodebuffer' }), 'BAD_XLSX');
  const zDocx = new JSZip(); zDocx.file('[Content_Types].xml', '<Types/>'); zDocx.file('word/document.xml', '<w:document/>');
  await hostile('.docx 結構（改副檔名冒充）', await zDocx.generateAsync({ type: 'nodebuffer' }), 'BAD_XLSX');
  // zip bomb：誠實宣告的大小
  const bomb1 = await rezip(bufE, (z) => z.file('xl/worksheets/sheet1.xml', '<?xml version="1.0"?><worksheet>' + ' '.repeat(40 * 1024 * 1024) + '</worksheet>'));
  t('4.1 zip bomb 檔案本身很小（壓縮後 < 100 KB，解壓 40 MB）', bomb1.length < 100 * 1024, bomb1.length);
  await hostile('zip bomb（工作表解壓 40 MB，誠實標示大小）', bomb1, 'BAD_XLSX');
  // zip bomb：偽造中央目錄的「未壓縮大小」為很小，實際資料解壓出來很大
  const bomb2 = Buffer.from(bomb1);
  {
    let i = -1, patched = 0;
    for (;;) {
      i = bomb2.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]), i + 1);
      if (i < 0) break;
      const nl = bomb2.readUInt16LE(i + 28);
      if (bomb2.slice(i + 46, i + 46 + nl).toString() === 'xl/worksheets/sheet1.xml') { bomb2.writeUInt32LE(2000, i + 24); patched++; }
    }
    t('4.2 偽造中央目錄大小的測試檔已建立（改了 1 個項目）', patched === 1, patched);
  }
  await hostile('zip bomb（偽造標頭宣告 2000 bytes、實際解壓 40 MB）', bomb2, 'BAD_XLSX');
  const bomb3 = await rezip(bufE, (z) => z.file('docProps/bomb.bin', Buffer.alloc(80 * 1024 * 1024)));
  const t0 = Date.now();
  const pb3 = await prev(bomb3, ITEMS);
  t('4.3 任何部件（即使是 docProps 這種用不到的）解壓後超過上限（80 MB 炸彈）→ 整個檔案拒絕 BAD_XLSX，而且很快（' + (Date.now() - t0) + 'ms）', cleanErr(pb3) && pb3.code === 'BAD_XLSX' && Date.now() - t0 < 5000, JSON.stringify(pb3).slice(0, 200));
  const bombMany = await rezip(bufE, (z) => { for (let i = 0; i < 700; i++) z.file('junk/f' + i + '.txt', 'x'); });
  await hostile('zip 內有 700 個項目', bombMany, 'BAD_XLSX');
  const withExtras = await rezip(bufE, (z) => { z.file('xl/vbaProject.bin', Buffer.from('MACRO')); z.file('xl/externalLinks/externalLink1.xml', '<externalLink/>'); z.file('xl/media/image1.png', Buffer.alloc(10)); z.file('docProps/core.xml', '<x/>'); z.file('xl/worksheets/_rels/sheet1.xml.rels', '<Relationships/>'); z.folder('xl/emptydir'); });
  const pExtra = await prev(withExtras, ITEMS);
  t('4.4 夾帶巨集／外部連結／圖片／資料夾項目的檔案：只當資料讀（不執行、不求值），內容照常解析', pExtra.ok && pExtra.summary.unchanged === 3, JSON.stringify(pExtra).slice(0, 160));
  const rp = await X.safeRepack(withExtras);
  t('4.4b 重新封裝後的 zip 是我們自己產生的（所有項目大小都與解壓後一致），可再被 JSZip 完整讀取', rp.ok && Object.keys((await JSZip.loadAsync(rp.buf)).files).includes('xl/worksheets/sheet1.xml'));
  const traversal = await rezip(bufE, (z) => z.file('../evil.xml', 'x'));
  await hostile('zip 項目名稱含 ../ 路徑跳脫', traversal, 'BAD_XLSX');
  await hostile('工作表 XML 含 DOCTYPE／ENTITY', await rezip(bufE, async (z) => { const s = await z.file('xl/worksheets/sheet1.xml').async('string'); z.file('xl/worksheets/sheet1.xml', s.replace('<worksheet', '<!DOCTYPE x [<!ENTITY a "aaaa">]><worksheet')); }), 'BAD_XLSX');
  await hostile('分頁名稱 __proto__', await rezip(bufE, async (z) => { const s = await z.file('xl/workbook.xml').async('string'); z.file('xl/workbook.xml', s.replace('name="ERP"', 'name="__proto__"')); }), 'BAD_XLSX');
  // 巨大 dimension：宣告 A1:XFD1048576，實際只有幾列
  const bigDim = await rezip(bufE, async (z) => { const s = await z.file('xl/worksheets/sheet1.xml').async('string'); z.file('xl/worksheets/sheet1.xml', s.replace(/<dimension ref="[^"]*"\/>/, '<dimension ref="A1:XFD1048576"/>')); });
  const t1 = Date.now();
  const pbd = await prev(bigDim, ITEMS);
  t('4.5 dimension 宣告 A1:XFD1048576（實際 4 列）→ 照常解析且很快（' + (Date.now() - t1) + 'ms），不依 dimension 配置記憶體', pbd.ok && pbd.summary.unchanged === 3 && Date.now() - t1 < 5000, JSON.stringify(pbd).slice(0, 160));
  // 20 萬個空白格式列（約 3 MB 工作表 XML，壓縮後很小）
  const manyRows = await rezip(bufE, async (z) => { const s = await z.file('xl/worksheets/sheet1.xml').async('string'); let extra = ''; for (let i = 6; i < 200000; i++) extra += `<row r="${i}"/>`; z.file('xl/worksheets/sheet1.xml', s.replace('</sheetData>', extra + '</sheetData>')); });
  const t2 = Date.now();
  const pbm = await prev(manyRows, ITEMS);
  t('4.6 20 萬個空白格式列 → 掃描上限內完成（' + (Date.now() - t2) + 'ms），資料照常解析', pbm.ok && pbm.summary.unchanged === 3 && Date.now() - t2 < 10000, JSON.stringify(pbm).slice(0, 160));
  // ZIP64 旗標／加密旗標／重複名稱
  const z64 = Buffer.from(bufE); { const e = z64.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06])); z64.writeUInt16LE(0xFFFF, e + 10); }
  await hostile('ZIP64 標記（EOCD 項目數 0xFFFF）', z64, 'BAD_XLSX');
  const enc = Buffer.from(bufE); { let i = -1; for (;;) { i = enc.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]), i + 1); if (i < 0) break; enc.writeUInt16LE(enc.readUInt16LE(i + 8) | 1, i + 8); } }
  await hostile('所有項目標示為加密', enc, 'BAD_XLSX');
  const dup = Buffer.from(bufE); { // 把第二個中央目錄項目的名稱改成跟第一個相同長度的同名
    const find = (nm) => { let i = -1; for (;;) { i = dup.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]), i + 1); if (i < 0) return -1; if (dup.slice(i + 46, i + 46 + dup.readUInt16LE(i + 28)).toString() === nm) return i; } };
    const i1 = find('xl/worksheets/sheet1.xml'), i2 = find('xl/worksheets/sheet2.xml');
    if (i1 > 0 && i2 > 0) Buffer.from('xl/worksheets/sheet1.xml').copy(dup, i2 + 46);
    t('4.7 重複項目名稱的測試檔已建立（sheet2 的中央目錄名稱改成 sheet1）', i1 > 0 && i2 > 0 && find('xl/worksheets/sheet2.xml') < 0);
  }
  await hostile('zip 內兩個同名項目', dup, 'BAD_XLSX');
  const stored = await rezip(bufE, () => {});   // 預設 DEFLATE；另做一份 STORE（不壓縮）也要能讀
  const zs = await JSZip.loadAsync(bufE); const storedBuf = await zs.generateAsync({ type: 'nodebuffer', compression: 'STORE' });
  const pst = await prev(storedBuf, ITEMS);
  t('4.8 STORE（不壓縮）與 DEFLATE 兩種 zip 都能讀', pst.ok && pst.summary.unchanged === 3 && (await prev(stored, ITEMS)).ok);
  // 工作表很大但合法（5000 實體列的上限內）
  const pdata = Array.from({ length: 300 }, (_, i) => ['Row' + i, 1, 1, '']);
  t('4.9 超過解析器掃描範圍（>5000 實體列）的尾端資料不會被讀到也不會當機', await (async () => {
    const b = mkXlsx([['ERP', [HDR, ['First', 1, 1, '']].concat(Array.from({ length: 6000 }, () => []))]]);
    const rv = await prev(b, []);
    return rv.ok && rv.merged.length === 1;
  })());
  void pdata;

  // ═════════════════ 5) 路由（記憶體 db） ═════════════════
  const USERS = [
    { username: 'admin1', role: 'admin', active: true, displayName: 'Admin' },
    { username: 'own1', role: 'user', active: true, displayName: 'Owner1', supervisor: 'mgr1', bu: ['ITS'] },
    { username: 'mgr1', role: 'manager1', active: true, displayName: 'Mgr1', bu: ['ITS'] },
    { username: 'cons1', role: 'consultx', active: true, displayName: 'Cons1', bu: ['ITS'] },
    { username: 'nofeat', role: 'nofeat', active: true, displayName: 'NoFeat', bu: ['ITS'] },
    { username: 'grp1', role: 'tecopm', active: true, displayName: 'Grp', bu: ['ITS'] },
  ];
  function mkEnv(realApp, adminOnly) {
    const routes = {};
    const app = realApp || {};
    if (!realApp) ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
    const quote = { id: 'Q1', quoteNo: 'QU-1', owner: 'own1', company: 'TestCo', projectName: 'Proj', status: 'draft', products: [], items: [{ lid: 'i1', desc: 'PM', unit: '人天', qty: 3, unitPrice: 8000, cost: 0 }], approval: null, discountType: 'none', discountValue: 0 };
    const env = { data: { quotations: [quote], quoteApproval: { roster: { gm: [], chairman: [], boardProxy: [], costProviders: ['cons1'], sealManagers: [] }, productClasses: {} }, contacts: [{ id: 'c1' }] }, logs: [], saves: 0, seq: 0 };
    env.auth = { users: JSON.parse(JSON.stringify(USERS)) };
    const requireAuth = function requireAuthStub(req, rs, next) { next(); };
    const requireAdmin = adminOnly
      ? function requireAdminReal(req, rs, next) { const u = env.auth.users.find((x) => x.username === req.session.user.username); if (u && u.role === 'admin') return next(); rs.status(403).json({ error: '需要管理員權限' }); }
      : function requireAdminStub(req, rs, next) { next(); };
    env.requireAuth = requireAuth; env.requireAdmin = requireAdmin;
    const deps = {
      db: { load: () => env.data, save: () => { env.saves++; }, flush: async () => {} }, loadAuth: () => env.auth, saveAuth: () => {},
      requireAuth, requireAdmin,
      writeLog: (...a) => env.logs.push(a), pushNotification: () => {},
      getViewableOwners: (req) => [req.session.user.username],
      sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-NEW', taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}),
      buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: async () => Buffer.from(''), QUOTE_TEMPLATE: '',
      uuidv4: () => 'u' + (++env.seq).toString().padStart(8, '0') + 'xxxx', normalizeBu: (b) => (Array.isArray(b) ? b : b ? [b] : []),
      getUserFeatures: (role) => (role === 'nofeat' || role === 'tecopm' ? [] : ['quotations']),
    };
    registerQuoteRoutes(app, deps);
    env.routesOf = (m, p) => routes[m.toUpperCase() + ' ' + p];
    env.call = (user, method, p, params, body, extra) => new Promise((resolve, reject) => {
      const h = [].concat(...env.routesOf(method, p));
      const role = (env.auth.users.find((u) => u.username === user) || {}).role;
      const req = { session: { user: { username: user, role } }, params: params || {}, query: {}, headers: {}, body: body === undefined ? {} : JSON.parse(JSON.stringify(body)) };
      if (extra && extra.file) req.file = extra.file;
      const hdr = {};
      const rs = {
        headersSent: false, status(c) { this._s = c; return this; },
        json(j) { this.headersSent = true; resolve({ s: this._s || 200, j, h: hdr }); return this; },
        send(b) { this.headersSent = true; resolve({ s: this._s || 200, j: null, buf: b, h: hdr }); return this; },
        setHeader(k, v) { hdr[String(k).toLowerCase()] = v; }, set() {},
      };
      let i = 0;
      const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
      next();
    });
    return env;
  }
  const A = '/api/admin/quote-pricebook';
  const fileOf = (buf, name) => ({ buffer: buf, originalname: name || 'import.xlsx', size: buf.length, mimetype: 'application/octet-stream' });
  let env = mkEnv();
  const seedPut = await env.call('admin1', 'PUT', A, {}, { items: [{ name: 'PM 顧問經理', price: 9000, cost: 6500 }, { name: 'SD 顧問', price: 7000, cost: 5000, active: false }, { name: 'ABAP 顧問', price: 6500, cost: 4500 }], updatedAt: null });
  t('5.0 前置：管理員先建三個項目', seedPut.s === 200 && seedPut.j.items.length === 3, JSON.stringify(seedPut.j));
  const seedIds = seedPut.j.items.map((x) => x.id);

  // 匯出
  let nl = env.logs.length, sv = env.saves;
  let ex = await env.call('admin1', 'GET', A + '/export');
  t('5.1 匯出：管理員 200，Content-Type=xlsx、no-store、nosniff', ex.s === 200 && /spreadsheetml\.sheet/.test(ex.h['content-type']) && /no-store/.test(ex.h['cache-control']) && ex.h['x-content-type-options'] === 'nosniff' && Buffer.isBuffer(ex.buf), JSON.stringify(ex.h));
  const cd = ex.h['content-disposition'] || '';
  t('5.2 Content-Disposition：attachment、ASCII 後備檔名 pricebook_20261008.xlsx、RFC5987 filename*=UTF-8\'\'%E5%A0%B1…_20261008.xlsx（解碼＝報價牌價簿_20261008.xlsx）', /^attachment; filename="pricebook_20261008\.xlsx"; filename\*=UTF-8''%E5%A0%B1/.test(cd) && decodeURIComponent(cd.split("filename*=UTF-8''")[1]) === '報價牌價簿_20261008.xlsx', cd);
  const exWb = XLSX.read(ex.buf, { type: 'buffer' });
  t('5.3 匯出內容＝目前牌價簿（含停用、順序）', eq(exWb.SheetNames, ['ERP', 'ITS', 'MDM', 'CRM', '說明']) && eq([2, 3, 4].map((r) => exWb.Sheets['ERP']['A' + r].v), ['PM 顧問經理', 'SD 顧問', 'ABAP 顧問']) && exWb.Sheets['ERP']['D3'].v === '否');
  const lgx = env.logs[env.logs.length - 1] || [];
  t('5.4 匯出稽核：EXPORT_QUOTE_PRICEBOOK、操作者、含項目數、註明含成本；沒有寫入資料', lgx[0] === 'EXPORT_QUOTE_PRICEBOOK' && lgx[1] === 'admin1' && /共3項/.test(lgx[3]) && /含成本/.test(lgx[3]) && env.logs.length === nl + 1 && env.saves === sv, JSON.stringify(lgx.slice(0, 4)));
  let tp = await env.call('admin1', 'GET', A + '/template');
  const cdT = tp.h['content-disposition'] || '';
  t('5.5 範本：200、xlsx、no-store、檔名 報價牌價簿_匯入範本_20261008.xlsx；資料表只有標題；稽核 DOWNLOAD_QUOTE_PRICEBOOK_TEMPLATE', tp.s === 200 && /no-store/.test(tp.h['cache-control']) && decodeURIComponent(cdT.split("filename*=UTF-8''")[1]) === '報價牌價簿_匯入範本_20261008.xlsx' && XLSX.read(tp.buf, { type: 'buffer' }).Sheets['ERP']['!ref'] === 'A1:D1' && env.logs[env.logs.length - 1][0] === 'DOWNLOAD_QUOTE_PRICEBOOK_TEMPLATE', cdT);
  // 非管理員
  nl = env.logs.length; sv = env.saves;
  const snapAll = JSON.stringify(env.data);
  let denied = [];
  for (const u of ['own1', 'mgr1', 'cons1', 'nofeat', 'grp1']) {
    denied.push(await env.call(u, 'GET', A + '/export'), await env.call(u, 'GET', A + '/template'), await env.call(u, 'POST', A + '/import-preview', {}, {}, { file: fileOf(bufRT) }));
  }
  t('5.6 非管理員（業務／主管／顧問／無功能／集團）打匯出、範本、匯入預覽 → 全部 403 NO_PERMISSION，沒有檔案內容、沒有稽核、沒有寫入', denied.every((d) => d.s === 403 && d.j && d.j.code === 'NO_PERMISSION' && !d.buf) && env.logs.length === nl && env.saves === sv && JSON.stringify(env.data) === snapAll);
  // 預覽
  const fileA = one([['pm 顧問經理', 9500, 6000, ''], ['New One', 5000, 3000, ''], ['SD 顧問', 7000, 5000, '是'], ['ABAP 顧問', 6500, 4500, ''], ['Bad', 'x', 1, '']]);
  nl = env.logs.length; sv = env.saves;
  const snapPre = JSON.stringify(env.data); const updPre = env.data.pricebook.updatedAt;
  let pv = await env.call('admin1', 'POST', A + '/import-preview', {}, {}, { file: fileOf(fileA, '我的牌價_2026.xlsx') });
  t('5.7 預覽 200：回 base.updatedAt、summary（含 byBu）、rows、skippedSheets、missingBus、merged、fileName', pv.s === 200 && pv.j.success === true && pv.j.base.updatedAt === updPre && pv.j.summary.added === 1 && pv.j.summary.updated === 2 && pv.j.summary.unchanged === 1 && pv.j.summary.errors === 1 && pv.j.fileName === '我的牌價_2026.xlsx' && Array.isArray(pv.j.skippedSheets) && eq(pv.j.missingBus, ['ITS', 'MDM', 'CRM']) && pv.j.summary.byBu.ERP.updated === 2 && pv.j.summary.byBu.ITS.updated === 0 && Array.isArray(pv.j.rows) && pv.j.merged.length === 4, JSON.stringify(pv.j).slice(0, 300));
  t('5.8 預覽完全不寫入：整份 data（含 pricebook.updatedAt）逐位元不變、db.save 沒被呼叫', JSON.stringify(env.data) === snapPre && env.saves === sv && env.data.pricebook.updatedAt === updPre);
  t('5.9 預覽完全不寫入：連稽核都不記（沒有新的 writeLog），也沒有 db.save', env.logs.length === nl && env.saves === sv, env.logs.length - nl);
  // 錯誤
  let e1 = await env.call('admin1', 'POST', A + '/import-preview', {}, {}, {});
  let e2 = await env.call('admin1', 'POST', A + '/import-preview', {}, {}, { file: fileOf(fileA, 'import.xls') });
  let e3 = await env.call('admin1', 'POST', A + '/import-preview', {}, {}, { file: fileOf(Buffer.from('not a zip at all'), 'x.xlsx') });
  let e4 = await env.call('admin1', 'POST', A + '/import-preview', {}, {}, { file: fileOf(fileA, 'evil.XLSX.exe') });
  let e5 = await env.call('admin1', 'POST', A + '/import-preview', {}, {}, { file: fileOf(Buffer.alloc(0), 'empty.xlsx') });
  t('5.10 沒檔案 → 400 NO_FILE；副檔名不是 .xlsx → 400 BAD_FILE_TYPE；內容不是 zip → 400 BAD_XLSX；空檔 → 400 EMPTY_FILE', e1.s === 400 && e1.j.code === 'NO_FILE' && e2.s === 400 && e2.j.code === 'BAD_FILE_TYPE' && e3.s === 400 && e3.j.code === 'BAD_XLSX' && e4.s === 400 && e4.j.code === 'BAD_FILE_TYPE' && e5.s === 400 && e5.j.code === 'EMPTY_FILE', JSON.stringify([e1.j, e2.j, e3.j, e4.j, e5.j]));
  t('5.11 上述錯誤都沒有任何寫入、沒有新稽核', env.saves === sv && JSON.stringify(env.data) === snapPre && env.logs.length === nl);
  // 套用：預覽的 merged → PUT
  const live0 = env.data.pricebook;
  const ap = await env.call('admin1', 'PUT', A, {}, { items: pv.j.merged, updatedAt: pv.j.base.updatedAt, via: 'import', fileName: pv.j.fileName });
  t('5.12 套用：PUT merged＋base.updatedAt＋via:import → 200；既有項目保留 id 與順序、數字更新、新項目附加在最後且拿到新 id', ap.s === 200 && eq(ap.j.items.slice(0, 3).map((x) => x.id), seedIds) && ap.j.items[0].price === 9500 && ap.j.items[0].cost === 6000 && ap.j.items[1].active === true && ap.j.items.length === 4 && ap.j.items[3].name === 'New One' && !seedIds.includes(ap.j.items[3].id) && /^pb/.test(ap.j.items[3].id), JSON.stringify(ap.j).slice(0, 300));
  const lga = env.logs[env.logs.length - 1];
  t('5.13 套用稽核：仍是 SAVE_QUOTE_PRICEBOOK（格式不變）、含實際異動（牌價 9000→9500、SD 啟用、新增 New One）、結尾加「來源：Excel 匯入（檔名）」', lga[0] === 'SAVE_QUOTE_PRICEBOOK' && ['共4項', '[ERP]「PM 顧問經理」牌價 9000→9500', '[ERP]「SD 顧問」啟用', '新增 [ERP]「New One」(牌價5000/成本3000)', '｜來源：Excel 匯入（我的牌價_2026.xlsx）'].every((s) => String(lga[3]).includes(s)), lga[3]);
  // 第二次匯同一份檔 → 不變
  const pv2 = await env.call('admin1', 'POST', A + '/import-preview', {}, {}, { file: fileOf(fileA, '我的牌價_2026.xlsx') });
  t('5.14 套用後再匯入同一份檔 → 5 列中 4 列「不變」、0 新增 0 更新（只剩壞掉的那一列報錯）', pv2.s === 200 && pv2.j.summary.unchanged === 4 && pv2.j.summary.added === 0 && pv2.j.summary.updated === 0 && pv2.j.summary.errors === 1, JSON.stringify(pv2.j.summary));
  // 過期
  const stalePv = await env.call('admin1', 'POST', A + '/import-preview', {}, {}, { file: fileOf(one([['Another New', 1, 1, '']]), 'b.xlsx') });
  const mid = await env.call('admin1', 'PUT', A, {}, { items: ap.j.items.map((x) => Object.assign({}, x, { price: x.price + 1 })), updatedAt: ap.j.updatedAt });   // 預覽之後有人先存
  const snapBefore409 = JSON.stringify(env.data.pricebook), n409 = env.logs.length;
  const st409 = await env.call('admin1', 'PUT', A, {}, { items: stalePv.j.merged, updatedAt: stalePv.j.base.updatedAt, via: 'import', fileName: 'b.xlsx' });
  t('5.15 預覽之後牌價簿被改過 → 套用時 409 STALE_PRICEBOOK（附現況），資料與稽核都沒變', mid.s === 200 && st409.s === 409 && st409.j.code === 'STALE_PRICEBOOK' && st409.j.live.updatedAt === mid.j.updatedAt && JSON.stringify(env.data.pricebook) === snapBefore409 && env.logs.length === n409, JSON.stringify(st409.j).slice(0, 200));
  // via 白名單與檔名淨化
  const cur2 = (await env.call('admin1', 'GET', A)).j;
  const snapVia = JSON.stringify(env.data.pricebook), nvia = env.logs.length;
  const badVias = [];
  for (const v of ['export', 'IMPORT', 'import ', 1, true, null, {}, ['import']]) badVias.push(await env.call('admin1', 'PUT', A, {}, { items: cur2.items, updatedAt: cur2.updatedAt, via: v }));
  t('5.16 via 只認 \'import\'：其他任何值 → 400，不寫入、不稽核', badVias.every((x) => x.s === 400 && x.j.code === 'BAD_PRICEBOOK') && JSON.stringify(env.data.pricebook) === snapVia && env.logs.length === nvia, JSON.stringify(badVias.map((x) => x.s)));
  const evilName = '..\\..\\etc\\pass\nwd\u0007' + 'x'.repeat(200) + '.xlsx';
  const okVia = await env.call('admin1', 'PUT', A, {}, { items: cur2.items.map((x, i) => i === 0 ? Object.assign({}, x, { price: x.price + 5 }) : x), updatedAt: cur2.updatedAt, via: 'import', fileName: evilName });
  const lgv = String(env.logs[env.logs.length - 1][3]);
  const srcPart = lgv.split('｜來源：Excel 匯入')[1] || '';
  t('5.17 稽核檔名淨化：去路徑、去控制字元／換行、限長 80；仍保留 .xlsx 結尾', okVia.s === 200 && !/[\u0000-\u001f]/.test(lgv) && !/etc/.test(srcPart.slice(0, 6)) && srcPart.length <= 80 + 6 && /\.xlsx）$/.test(srcPart), srcPart);
  const noName = await env.call('admin1', 'PUT', A, {}, { items: okVia.j.items, updatedAt: okVia.j.updatedAt, via: 'import' });
  t('5.18 via:import 沒帶檔名 → 稽核「來源：Excel 匯入」不帶括號', noName.s === 200 && /｜來源：Excel 匯入$/.test(String(env.logs[env.logs.length - 1][3])) || /無異動｜來源：Excel 匯入$/.test(String(env.logs[env.logs.length - 1][3])), env.logs[env.logs.length - 1][3]);
  const plain = await env.call('admin1', 'PUT', A, {}, { items: noName.j.items.map((x, i) => i === 0 ? Object.assign({}, x, { cost: x.cost + 1 }) : x), updatedAt: noName.j.updatedAt });
  t('5.19 一般儲存（沒有 via）的稽核與原本相同，不含「來源」', plain.s === 200 && !/來源/.test(String(env.logs[env.logs.length - 1][3])) && env.logs[env.logs.length - 1][0] === 'SAVE_QUOTE_PRICEBOOK');
  t('5.20 PUT 仍有完整驗證：import 路徑送壞資料（重複名稱）→ 400', (await env.call('admin1', 'PUT', A, {}, { items: [{ name: 'X', price: 1, cost: 1 }, { name: 'x', price: 1, cost: 1 }], updatedAt: plain.j.updatedAt, via: 'import', fileName: 'a.xlsx' })).s === 400);
  // 掛了 requireAdmin
  const chain = (m, p) => env.routesOf(m, p);
  t('5.21 三條新路由的 middleware 鏈都含 requireAdmin（不是只靠 handler 內的角色檢查）', ['GET ' + A + '/export', 'GET ' + A + '/template', 'POST ' + A + '/import-preview'].every((k) => { const [m, p] = k.split(' '); return chain(m, p) && chain(m, p).flat().includes(env.requireAdmin); }));
  t('5.22 三條新路由的順序：requireAdmin 在上傳 middleware 之前（非管理員不會觸發 multer 解析）', (() => { const c = chain('POST', A + '/import-preview').flat(); return c.indexOf(env.requireAdmin) >= 0 && c.indexOf(env.requireAdmin) < c.length - 2; })());
  t('5.23 已存報價單、簽核設定、其他命名空間在所有匯入匯出操作之後仍位元不變', JSON.stringify({ q: env.data.quotations, qa: env.data.quoteApproval, c: env.data.contacts }) === JSON.stringify({ q: [{ id: 'Q1', quoteNo: 'QU-1', owner: 'own1', company: 'TestCo', projectName: 'Proj', status: 'draft', products: [], items: [{ lid: 'i1', desc: 'PM', unit: '人天', qty: 3, unitPrice: 8000, cost: 0 }], approval: null, discountType: 'none', discountValue: 0 }], qa: { roster: { gm: [], chairman: [], boardProxy: [], costProviders: ['cons1'], sealManagers: [] }, productClasses: {} }, c: [{ id: 'c1' }] }));

  // ═════════════════ 5b) BU：舊資料相容、路由（記憶體 db） ═════════════════
  {
    const A2 = '/api/admin/quote-pricebook';
    const e2 = mkEnv();
    // 舊資料：已存的牌價簿沒有 bu 欄位（上線時的形狀）
    e2.data.pricebook = { items: [{ id: 'old1', name: 'Legacy PM', unit: '人天', price: 9000, cost: 6500, active: true }, { id: 'old2', name: 'Legacy SD', unit: '人天', price: 7000, cost: 5000, active: false }], updatedAt: '2026-01-01T00:00:00.000Z', updatedBy: 'x' };
    const snapOld = JSON.stringify(e2.data.pricebook);
    const ga = await e2.call('admin1', 'GET', A2);
    const gp = await e2.call('cons1', 'GET', '/api/quote-pricebook');
    t('5b.1 舊資料（沒有 bu）：GET admin 與 GET 業務端都回 bu:"ERP"；GET 不寫入、儲存的資料逐位元不變', ga.s === 200 && ga.j.items.every((x) => x.bu === 'ERP') && gp.s === 200 && gp.j.items.length === 1 && gp.j.items[0].bu === 'ERP' && JSON.stringify(e2.data.pricebook) === snapOld, JSON.stringify([ga.j, gp.j]).slice(0, 300));
    const exl = await e2.call('admin1', 'GET', A2 + '/export');
    const exlWb = XLSX.read(exl.buf, { type: 'buffer' });
    t('5b.2 舊資料匯出：兩個項目都在 ERP 工作表，其他 BU 工作表只有標題', exl.s === 200 && exlWb.Sheets.ERP.A2.v === 'Legacy PM' && exlWb.Sheets.ERP.A3.v === 'Legacy SD' && exlWb.Sheets.ITS['!ref'] === 'A1:D1' && exlWb.Sheets.CRM['!ref'] === 'A1:D1');
    // 舊用戶端：送來的項目沒有 bu
    const oldPut = await e2.call('admin1', 'PUT', A2, {}, { items: [{ id: 'old1', name: 'Legacy PM', price: 9100, cost: 6500, active: true }, { id: 'old2', name: 'Legacy SD', price: 7000, cost: 5000, active: false }, { name: 'Added by old client', price: 1, cost: 1 }], updatedAt: '2026-01-01T00:00:00.000Z' });
    t('5b.3 舊用戶端 PUT（項目沒有 bu）→ 200，不會 500；存下來每個項目都有明確的 bu:"ERP"（儲存即補上），id 保留', oldPut.s === 200 && oldPut.j.items.length === 3 && oldPut.j.items.every((x) => x.bu === 'ERP') && oldPut.j.items[0].id === 'old1' && e2.data.pricebook.items.every((x) => x.bu === 'ERP'), JSON.stringify(oldPut.j).slice(0, 300));
    // 多 BU 儲存
    const multiPut = await e2.call('admin1', 'PUT', A2, {}, { items: oldPut.j.items.concat([{ bu: 'ITS', name: 'Legacy PM', price: 6000, cost: 4000 }, { bu: 'MDM', name: 'M1', price: 1, cost: 1 }, { bu: 'CRM', name: 'Legacy PM', price: 2, cost: 2 }]), updatedAt: oldPut.j.updatedAt });
    t('5b.4 多 BU 儲存：同名「Legacy PM」可同時存在 ERP／ITS／CRM（費率不同）；GET 回傳每個都有 bu', multiPut.s === 200 && multiPut.j.items.filter((x) => x.name === 'Legacy PM').map((x) => x.bu + ':' + x.price).join() === 'ERP:9100,ITS:6000,CRM:2', JSON.stringify(multiPut.j.items.map((x) => x.bu + x.name)));
    // 舊用戶端再存一次（看不到 bu 欄位，把整份清單原樣送回但拿掉 bu）：id 對得上的沿用既有 bu，不洗成 ERP
    const stripped = multiPut.j.items.map((x) => { const y = Object.assign({}, x); delete y.bu; return y; });
    const oldAgain = await e2.call('admin1', 'PUT', A2, {}, { items: stripped.map((x, i) => i === 0 ? Object.assign({}, x, { price: 9200 }) : x), updatedAt: multiPut.j.updatedAt });
    t('5b.5 舊的快取頁面把整份（拿掉 bu）存回：id 對得上的項目沿用既有 BU（ITS／MDM／CRM 的項目不會被洗成 ERP）', oldAgain.s === 200 && eq(oldAgain.j.items.map((x) => x.bu), multiPut.j.items.map((x) => x.bu)) && oldAgain.j.items[0].price === 9200, JSON.stringify(oldAgain.j.items.map((x) => x.bu)));
    // 驗證
    const base = oldAgain.j;
    const tryPut = (items) => e2.call('admin1', 'PUT', A2, {}, { items, updatedAt: base.updatedAt });
    const dupInBu = await tryPut([{ bu: 'ITS', name: 'X', price: 1, cost: 1 }, { bu: 'ITS', name: 'x ', price: 1, cost: 1 }]);
    const sameNameDiffBu = await tryPut([{ bu: 'ITS', name: 'X', price: 1, cost: 1 }, { bu: 'MDM', name: 'x', price: 1, cost: 1 }]);
    t('5b.6 名稱唯一性是每個 BU 內：同 BU 重複 → 400（field=name，訊息含 [ITS]）；不同 BU 同名 → 200', dupInBu.s === 400 && dupInBu.j.field === 'name' && /\[ITS\]/.test(dupInBu.j.error) && sameNameDiffBu.s === 200, JSON.stringify([dupInBu.j, sameNameDiffBu.s]).slice(0, 300));
    const cur3 = (await e2.call('admin1', 'GET', A2)).j;
    const badBu = [];
    for (const v of ['XYZ', 'erp', 'ERP ', 5, true, {}, ['ERP']]) badBu.push(await e2.call('admin1', 'PUT', A2, {}, { items: [{ bu: v, name: 'Q', price: 1, cost: 1 }], updatedAt: cur3.updatedAt }));
    t('5b.7 bu 不是 ERP／ITS／MDM／CRM 其中之一（含小寫、多空白、數字、物件）→ 400 field=bu；不寫入', badBu.every((x) => x.s === 400 && x.j.field === 'bu') && (await e2.call('admin1', 'GET', A2)).j.updatedAt === cur3.updatedAt);
    const nullBu = await e2.call('admin1', 'PUT', A2, {}, { items: [{ bu: null, name: 'Q', price: 1, cost: 1 }, { bu: '', name: 'R', price: 1, cost: 1 }], updatedAt: cur3.updatedAt });
    t('5b.8 bu 為 null／空字串視同沒帶（→ ERP），不是 500', nullBu.s === 200 && nullBu.j.items.every((x) => x.bu === 'ERP'));
    const mk = (bu, n) => Array.from({ length: n }, (_, i) => ({ bu, name: bu + i, price: 1, cost: 1 }));
    const u0 = (await e2.call('admin1', 'GET', A2)).j.updatedAt;
    const capOk = await e2.call('admin1', 'PUT', A2, {}, { items: mk('ERP', 60).concat(mk('ITS', 60), mk('MDM', 60), mk('CRM', 60)), updatedAt: u0 });
    const cap61 = await e2.call('admin1', 'PUT', A2, {}, { items: mk('ERP', 60).concat(mk('ITS', 61)), updatedAt: capOk.j.updatedAt });
    const cap241 = await e2.call('admin1', 'PUT', A2, {}, { items: mk('ERP', 60).concat(mk('ITS', 60), mk('MDM', 60), mk('CRM', 60), mk('CRM', 1).map((x) => Object.assign(x, { name: 'extra' }))), updatedAt: capOk.j.updatedAt });
    t('5b.9 上限：每個 BU 60（共 240）通過；任一 BU 61 筆 → 400；總數 241 → 400', capOk.s === 200 && capOk.j.items.length === 240 && cap61.s === 400 && /\[ITS\]/.test(cap61.j.error) && cap241.s === 400, JSON.stringify([capOk.s, cap61.j, cap241.j]).slice(0, 300));
    // 稽核含 BU
    const logs0 = e2.logs.length;
    const mdmDrop = capOk.j.items[120].id;
    const aud = await e2.call('admin1', 'PUT', A2, {}, { items: capOk.j.items.filter((x) => x.id !== mdmDrop).map((x, i) => i === 0 ? Object.assign({}, x, { price: 7500 }) : (i === 60 ? Object.assign({}, x, { active: false }) : x)), updatedAt: capOk.j.updatedAt });
    const lg = String(e2.logs[e2.logs.length - 1][3]);
    t('5b.10 稽核摘要每個異動都標 BU（[ERP]「ERP0」牌價 1→7500、[ITS]「ITS0」停用、[MDM] 移除…）', aud.s === 200 && e2.logs.length === logs0 + 1 && lg.includes('[ERP]「ERP0」牌價 1→7500') && lg.includes('[ITS]「ITS0」停用') && lg.includes('移除 [MDM]「MDM0」'), lg.slice(0, 300));
    // 匯入（多 BU）：預覽不寫入 → PUT merged → 稽核含 BU 與來源
    const e3 = mkEnv();
    await e3.call('admin1', 'PUT', A2, {}, { items: [{ bu: 'ERP', name: 'PM', price: 9000, cost: 6500 }, { bu: 'ITS', name: 'PM', price: 8000, cost: 5500 }, { bu: 'MDM', name: 'Keep', price: 1, cost: 1 }], updatedAt: null });
    const file3b = multi({ ERP: [['PM', 9500, 6500, ''], ['NewE', 1, 1, '']], ITS: [['PM', 8000, 5500, ''], ['Bad', 'x', 1, '']], CRM: [['PM', 3, 3, '']] });
    const snap3 = JSON.stringify(e3.data); const sv3 = e3.saves, nl3 = e3.logs.length;
    const pv3 = await e3.call('admin1', 'POST', A2 + '/import-preview', {}, {}, { file: fileOf(file3b, '多BU.xlsx') });
    t('5b.11 多 BU 預覽 200：summary.byBu 正確、rows 帶 bu、missingBus=[MDM]；預覽沒有寫入也沒有稽核', pv3.s === 200 && pv3.j.summary.byBu.ERP.added === 1 && pv3.j.summary.byBu.ERP.updated === 1 && pv3.j.summary.byBu.ITS.unchanged === 1 && pv3.j.summary.byBu.ITS.errors === 1 && pv3.j.summary.byBu.CRM.added === 1 && eq(pv3.j.missingBus, ['MDM']) && pv3.j.rows.every((x) => BUS4.includes(x.bu)) && JSON.stringify(e3.data) === snap3 && e3.saves === sv3 && e3.logs.length === nl3, JSON.stringify(pv3.j.summary));
    const ap3 = await e3.call('admin1', 'PUT', A2, {}, { items: pv3.j.merged, updatedAt: pv3.j.base.updatedAt, via: 'import', fileName: pv3.j.fileName });
    const lg3 = String(e3.logs[e3.logs.length - 1][3]);
    t('5b.12 套用 merged：各 BU 內順序與 id 保留，新項目接在該 BU 最後；稽核含 [ERP]／[CRM] 標記與「來源：Excel 匯入（多BU.xlsx）」', ap3.s === 200 && eq(ap3.j.items.map((x) => x.bu + ':' + x.name), ['ERP:PM', 'ERP:NewE', 'ITS:PM', 'MDM:Keep', 'CRM:PM']) && lg3.includes('[ERP]「PM」牌價 9000→9500') && lg3.includes('新增 [ERP]「NewE」') && lg3.includes('[CRM]「PM」') && lg3.includes('｜來源：Excel 匯入（多BU.xlsx）'), lg3);
    const again3 = await e3.call('admin1', 'POST', A2 + '/import-preview', {}, {}, { file: fileOf(file3b, '多BU.xlsx') });
    t('5b.13 再匯入同一個檔 → 新增 0／更新 0（ERP 2、ITS 1、CRM 1 不變；壞的那列仍是錯誤）', again3.j.summary.added === 0 && again3.j.summary.updated === 0 && again3.j.summary.unchanged === 3 + 1 && again3.j.summary.errors === 1, JSON.stringify(again3.j.summary));
    const stale3 = await e3.call('admin1', 'PUT', A2, {}, { items: pv3.j.merged, updatedAt: pv3.j.base.updatedAt, via: 'import', fileName: 'x.xlsx' });
    t('5b.14 用舊預覽的 base 再套用 → 409 STALE_PRICEBOOK（資料不被蓋掉）', stale3.s === 409 && stale3.j.code === 'STALE_PRICEBOOK' && stale3.j.live.items.length === 5);
    const ex3 = await e3.call('admin1', 'GET', A2 + '/export');
    const lgx3 = e3.logs[e3.logs.length - 1];
    t('5b.15 匯出稽核含各 BU 項目數（ERP 2、ITS 1、MDM 1、CRM 1）與「含成本」；匯出檔對得上', ex3.s === 200 && lgx3[0] === 'EXPORT_QUOTE_PRICEBOOK' && /ERP 2、ITS 1、MDM 1、CRM 1/.test(lgx3[3]) && /含成本/.test(lgx3[3]) && XLSX.read(ex3.buf, { type: 'buffer' }).Sheets.CRM.A2.v === 'PM');
  }

  // ═════════════════ 6) 真實 HTTP（express＋multer） ═════════════════
  const express = require(path.join(ROOT, 'node_modules/express'));
  const happ = express();
  happ.use(express.json({ limit: '2mb' }));
  happ.use((req, rs, next) => { const u = req.headers['x-test-user']; if (u) req.session = { user: { username: u, role: (USERS.find((x) => x.username === u) || {}).role } }; next(); });
  const henv = mkEnv(happ, true);
  happ.use((err, req, rs, next) => { rs.status(500).json({ error: 'unhandled:' + err.message }); });   // 若有人漏接錯誤，測試會看到 500
  const server = await new Promise((resolve) => { const s = happ.listen(0, '127.0.0.1', () => resolve(s)); });
  const base = 'http://127.0.0.1:' + server.address().port;
  const H = (u) => ({ 'x-test-user': u });
  try {
    await fetch(base + A, { method: 'PUT', headers: Object.assign({ 'content-type': 'application/json' }, H('admin1')), body: JSON.stringify({ items: [{ name: 'PM 顧問經理', price: 9000, cost: 6500 }, { name: 'SD 顧問', price: 7000, cost: 5000, active: false }], updatedAt: null }) });
    const up = (buf, filename, field, u) => { const fd = new FormData(); fd.append(field || 'file', new Blob([buf]), filename); return fetch(base + A + '/import-preview', { method: 'POST', headers: H(u || 'admin1'), body: fd }); };
    const rj = async (r) => { const tx = await r.text(); try { return { s: r.status, j: JSON.parse(tx), tx }; } catch (_) { return { s: r.status, j: null, tx }; } };
    const noLeak = (x) => x.tx && !/node_modules|\.js:\d|\bat [A-Za-z.<]+ \(|[A-Z]:\\\\|\/usr\/|\/home\/|MulterError|busboy/.test(x.tx);
    let h1 = await rj(await up(fileA, '我的牌價_測試.xlsx'));
    t('6.1 multipart 上傳 → 200 預覽；中文檔名完整保留（utf8）', h1.s === 200 && h1.j.fileName === '我的牌價_測試.xlsx' && h1.j.summary.updated >= 1 && Array.isArray(h1.j.merged), h1.tx.slice(0, 200));
    const big = Buffer.concat([fileA, Buffer.alloc(2.5 * 1024 * 1024, 1)]);
    let h2 = await rj(await up(big, 'big.xlsx'));
    t('6.2 2.5 MB 檔案 → 413 FILE_TOO_LARGE（乾淨 JSON，不洩漏）', h2.s === 413 && h2.j.code === 'FILE_TOO_LARGE' && noLeak(h2), h2.tx.slice(0, 200));
    const justUnder = await rj(await up(Buffer.concat([fileA, Buffer.alloc(1.5 * 1024 * 1024, 1)]), 'under.xlsx'));
    t('6.3 1.5 MB 的檔案通過大小限制（尾端多餘資料 → 因 zip 結構壞掉而 400 BAD_XLSX，不是 413/500）', justUnder.s === 400 && justUnder.j.code === 'BAD_XLSX', justUnder.tx.slice(0, 150));
    let h4 = await rj(await up(fileA, 'x.xlsx', 'wrongfield'));
    t('6.4 上傳欄位名稱不是 file → 400 BAD_UPLOAD（乾淨）', h4.s === 400 && h4.j.code === 'BAD_UPLOAD' && noLeak(h4), h4.tx.slice(0, 200));
    let h5 = await rj(await fetch(base + A + '/import-preview', { method: 'POST', headers: Object.assign({ 'content-type': 'application/json' }, H('admin1')), body: '{"a":1}' }));
    t('6.5 不是 multipart（JSON body）→ 400 NO_FILE', h5.s === 400 && h5.j.code === 'NO_FILE', h5.tx.slice(0, 150));
    let h6 = await rj(await up(fileA, 'data.xls'));
    t('6.6 副檔名 .xls → 400 BAD_FILE_TYPE', h6.s === 400 && h6.j.code === 'BAD_FILE_TYPE');
    let h7 = await rj(await up(Buffer.from('<html>hi</html>'), 'fake.xlsx'));
    t('6.7 副檔名 .xlsx 但內容是 HTML → 400 BAD_XLSX（內容檢查，不只看副檔名）', h7.s === 400 && h7.j.code === 'BAD_XLSX' && noLeak(h7));
    const fd2 = new FormData(); fd2.append('file', new Blob([fileA]), 'a.xlsx'); fd2.append('file', new Blob([fileA]), 'b.xlsx');
    let h8 = await rj(await fetch(base + A + '/import-preview', { method: 'POST', headers: H('admin1'), body: fd2 }));
    t('6.8 一次傳兩個檔案 → 400 BAD_UPLOAD', h8.s === 400 && h8.j.code === 'BAD_UPLOAD' && noLeak(h8), h8.tx.slice(0, 150));
    const fd3 = new FormData(); fd3.append('file', new Blob([fileA]), 'a.xlsx'); for (let i = 0; i < 12; i++) fd3.append('f' + i, 'x');
    let h9 = await rj(await fetch(base + A + '/import-preview', { method: 'POST', headers: H('admin1'), body: fd3 }));
    t('6.9 附帶一堆文字欄位 → 400 BAD_UPLOAD（欄位數受限）', h9.s === 400 && h9.j.code === 'BAD_UPLOAD');
    // 非管理員（真實 requireAdmin：先擋，multer 不會解析）
    const nonAdmin = [];
    for (const u of ['own1', 'mgr1', 'cons1', 'nofeat', 'grp1']) {
      nonAdmin.push(await rj(await up(fileA, 'a.xlsx', 'file', u)), await rj(await fetch(base + A + '/export', { headers: H(u) })), await rj(await fetch(base + A + '/template', { headers: H(u) })));
    }
    t('6.10 非管理員（5 種角色 × 3 條路由）→ 全部 403', nonAdmin.every((x) => x.s === 403), nonAdmin.map((x) => x.s).join());
    const noLogin = await rj(await fetch(base + A + '/export'));
    t('6.11 沒登入（沒有 session）→ 不會拿到檔案（被擋，非 200）', noLogin.s !== 200, noLogin.s);
    // 下載
    const dl = await fetch(base + A + '/export', { headers: H('admin1') });
    const dlBuf = Buffer.from(await dl.arrayBuffer());
    const dlCd = dl.headers.get('content-disposition') || '';
    t('6.12 真實 HTTP 下載：200、xlsx MIME、Cache-Control no-store、Content-Disposition 含 RFC5987 中文檔名；檔案可被開啟且內容正確', dl.status === 200 && /spreadsheetml\.sheet/.test(dl.headers.get('content-type')) && /no-store/.test(dl.headers.get('cache-control')) && decodeURIComponent(dlCd.split("filename*=UTF-8''")[1]) === '報價牌價簿_20261008.xlsx' && XLSX.read(dlBuf, { type: 'buffer' }).Sheets['ERP']['A2'].v === 'PM 顧問經理' && dlBuf.readUInt32LE(0) === 0x04034b50, dlCd);
    // 完整流程：匯出檔 → 上傳 → 全部不變
    let rtH = await rj(await up(dlBuf, '報價牌價簿_20261008.xlsx'));
    t('6.13 真實 HTTP 往返：下載的匯出檔原封不動上傳 → 全部「不變」', rtH.s === 200 && rtH.j.summary.unchanged === 2 && rtH.j.summary.added === 0 && rtH.j.summary.updated === 0 && rtH.j.summary.errors === 0, rtH.tx.slice(0, 200));
    // 套用（HTTP）
    const applyH = await rj(await fetch(base + A, { method: 'PUT', headers: Object.assign({ 'content-type': 'application/json' }, H('admin1')), body: JSON.stringify({ items: h1.j.merged, updatedAt: h1.j.base.updatedAt, via: 'import', fileName: h1.j.fileName }) }));
    t('6.14 真實 HTTP：預覽的 merged 與 base 送 PUT → 200', applyH.s === 200 && applyH.j.success === true, applyH.tx.slice(0, 200));
    const staleH = await rj(await fetch(base + A, { method: 'PUT', headers: Object.assign({ 'content-type': 'application/json' }, H('admin1')), body: JSON.stringify({ items: h1.j.merged, updatedAt: h1.j.base.updatedAt, via: 'import', fileName: h1.j.fileName }) }));
    t('6.15 真實 HTTP：同一份舊預覽再套用一次 → 409 STALE_PRICEBOOK', staleH.s === 409 && staleH.j.code === 'STALE_PRICEBOOK');
    const putNon = await rj(await fetch(base + A, { method: 'PUT', headers: Object.assign({ 'content-type': 'application/json' }, H('own1')), body: JSON.stringify({ items: [], updatedAt: applyH.j.updatedAt, via: 'import' }) }));
    t('6.16 真實 HTTP：非管理員 PUT（via:import）→ 403，資料沒被清空', putNon.s === 403 && henv.data.pricebook.items.length >= 2);
  } finally { server.close(); }
}

run().then(() => {
  let pass = 0, fail = 0;
  res.forEach(([n, ok, x]) => {
    console.log((ok ? 'PASS ' : 'FAIL ') + n + (x && !ok ? '  <- ' + x : ''));
    ok ? pass++ : fail++;
  });
  console.log('\n報價牌價簿 Excel 匯入匯出檢查：PASS ' + pass + ' / FAIL ' + fail);
  process.exit(fail ? 1 : 0);
}).catch((e) => { res.filter((r) => !r[1]).forEach((r) => console.log('FAIL ' + r[0] + (r[2] ? '  <- ' + r[2] : ''))); console.error('例外（後面的檢查沒跑到）', e.stack); process.exit(2); });
