'use strict';
/**
 * 毛利分析（PNL）Excel —— **內部用，不得給客戶**（含成本單價、成本小計、毛利、毛利率）。
 * 與「給客戶的報價單」（lib/quoteExcel.js）分開匯出；權限由 GET /api/quotations/:id/export-pnl 把關。
 *
 * 版面：公司原有的「Profitability Calculation Worksheet for Implementation」（報價單範本.xlsx 的「PNL (C)」頁籤）。
 * 範本 templates/pnl_template.xlsx 是用 Excel 從原檔整理出的「單頁、無外部連結、無垃圾已定義名稱」乾淨版，
 * 並修正了原範本的公式缺陷（見檔尾「範本與原檔的差異」）。這裡只用 JSZip 直接改 sheet1.xml 的儲存格（沿用範本樣式），
 * 不經 SheetJS（社群版寫檔會丟掉所有樣式）。公式保持為公式，使用者在 Excel 內改數字會連動；
 * 同時把「公式快取值」重算寫進檔案（預覽器／手機／雲端檢視器只認快取值），並設 fullCalcOnLoad 讓 Excel 開檔時再算一次。
 *
 * 資料對應（PNL 把收入／成本分成 顧問服務／軟體／硬體／其他 四區，每區可填的列數有限）：
 *   - 每個品項有分類 cat（consult／software／hardware／other；空白＝自動：依報價單勾選的商品類別，單一類別就全歸該類，
 *     多類別再依單位推測，仍無法判斷歸「其他」）。
 *   - 收入用「折扣後金額」：折扣依各品項小計比例攤到品項（以「分」做最大餘數法，各品項加總＝專案優惠價，不差一分）。
 *   - 成本不打折：成本單價 × 數量。
 *   - 某一區的品項比可填列數多時，最後一列放「其餘 N 項合計」。
 * 範本沒有、系統也沒有的資料（PM、PCM、專案起迄日、請款階段、顧問角色/人數、供應商、差旅…）一律留白，給人工補。
 */
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { _internal: { SheetXml } } = require('./quoteExcel');
const QI = require('./quoteItems');

const TEMPLATE = path.join(__dirname, '..', 'templates', 'pnl_template.xlsx');
const SHEET_PATH = 'xl/worksheets/sheet1.xml';

/** 風險預留（%）可選值與範本 B89「Project Risk level」的英文等級字樣；與 lib/quoteRoutes.js 的 CONTINGENCY_PCTS 一致 */
const CONTINGENCY_PCTS = [0, 5, 10, 15, 20];
const RISK_LABELS = { 0: 'None', 5: 'Low', 10: 'Medium', 15: 'High', 20: 'Very High' };
const CATS = ['consult', 'software', 'hardware', 'other'];
const CAT_LABELS = { consult: '顧問服務', software: '軟體', hardware: '硬體', other: '其他' };
/** 商品類別（簽核設定的 productClasses.cls）→ PNL 四區 */
const CLASS_TO_CAT = { consult: 'consult', software: 'software', hardware: 'hardware', crm: 'other', mdm: 'other', ot: 'other', other: 'other' };

// 各區在範本上「可填的列」。收入：金額寫 H；成本：顧問寫 D(單位成本)×F(人天)，其餘寫 F(數量)×G(單價)，H 為公式
const REV_ROWS = { consult: [16, 17, 18], software: [23], hardware: [27, 28], other: [32, 33] };
const COST_ROWS = { consult: [38, 39], software: [59, 60], hardware: [64, 65], other: [83, 84] };

const num = (v) => { const n = parseFloat(v); return Number.isFinite(n) ? n : 0; };
const bigCents = (x) => BigInt(Math.round(num(x) * 100));

// ── 分類 ────────────────────────────────────────────────────
function normCat(c) { return CATS.includes(c) ? c : ''; }

/** 品項沒有明確分類時的推測：單位（人天→顧問、台/組→硬體、授權/套→軟體），都不像就歸「其他」 */
function catByUnit(unit) {
  const u = String(unit || '');
  if (/人天|人日|人月|man.?day|\bMD\b|^天$|^日$/i.test(u)) return 'consult';
  if (/^(台|組|部|臺|pcs|set|unit)$/i.test(u)) return 'hardware';
  if (/授權|license|套|user|seat|帳號|訂閱|subscription/i.test(u)) return 'software';
  return 'other';
}

/** classCodes：報價單勾選商品的類別代碼（consult/software/hardware/crm/mdm/ot/other）；單一 PNL 區就全歸該區 */
function resolveCat(item, classCodes) {
  const own = normCat(item && item.cat);
  if (own) return own;
  const set = new Set((classCodes || []).map(c => CLASS_TO_CAT[c]).filter(Boolean));
  if (set.size === 1) return [...set][0];
  return catByUnit(item && item.unit);
}

// ── 折扣攤提（以「分」做最大餘數法）──────────────────────────
/**
 * 回傳 { listCents: [..], revCents: [..] }（BigInt 陣列）。revCents 各品項加總＝折扣後總額。
 * percent：discountValue 是「實收百分比」（90＝九折，0<v<100）；amount：議價後的未稅總價；其他：不折扣。
 */
function allocateRevenue(items, discountType, discountValue) {
  const list = items.map(it => bigCents(num(it.qty) * num(it.unitPrice)));
  const L = list.reduce((s, x) => s + x, 0n);
  let A = L;
  const v = num(discountValue);
  if (discountType === 'percent' && v > 0 && v < 100) A = (L * BigInt(Math.round(v * 1000)) + 50000n) / 100000n;   // 四捨五入到分
  else if (discountType === 'amount' && v > 0) A = bigCents(v);
  if (L === 0n) return { listCents: list, revCents: list.map(() => 0n), totalCents: A };
  const base = list.map(x => (A * x) / L);
  const rem = list.map((x, i) => (A * x) % L);
  let left = A - base.reduce((s, x) => s + x, 0n);
  const order = list.map((_, i) => i).sort((a, b) => (rem[b] > rem[a] ? 1 : rem[b] < rem[a] ? -1 : a - b));
  for (const i of order) { if (left <= 0n) break; base[i] += 1n; left -= 1n; }
  return { listCents: list, revCents: base, totalCents: A };
}
const toTwd = (cents) => Number(cents) / 100;

// ── 公式重算（只支援範本用到的語法：+ - * / 括號、儲存格、SUM(範圍/儲存格…)）──────────
const colNum = (c) => c.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colName = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const splitAddr = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); return { col: m[1], row: +m[2] }; };
function shiftFormula(f, dRow, dCol) {
  return f.replace(/\$?([A-Z]{1,3})\$?(\d+)/g, (m, c, r) => colName(colNum(c) + dCol) + (+r + dRow));
}

/** 解析 sheet XML 的所有儲存格 → Map(addr → { attrs, inner, formula, value, type }) */
function parseCells(xml) {
  const cells = new Map();
  const re = /<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
  let m;
  while ((m = re.exec(xml))) {
    const [, addr, attrs, inner = ''] = m;
    const fm = /<f([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(inner);
    const vm = /<v>([\s\S]*?)<\/v>/.exec(inner);
    const t = (/\bt="(\w+)"/.exec(attrs) || [])[1] || '';
    cells.set(addr, { attrs, inner, fAttrs: fm ? fm[1] : null, fText: fm ? (fm[2] || '') : null, v: vm ? vm[1] : null, t });
  }
  return cells;
}

function evalWorkbookSheet(xml) {
  const cells = parseCells(xml);
  // 共用公式：master 的 ref 範圍內其他儲存格沿用 master 公式（相對位移）
  const shared = new Map();
  for (const [addr, c] of cells) {
    if (c.fAttrs != null && /t="shared"/.test(c.fAttrs) && c.fText) shared.set((/si="(\d+)"/.exec(c.fAttrs) || [])[1], { addr, f: c.fText });
  }
  const formulaOf = (addr) => {
    const c = cells.get(addr);
    if (!c || c.fAttrs == null) return null;
    if (c.fText) return c.fText;
    const si = (/si="(\d+)"/.exec(c.fAttrs) || [])[1];
    const mst = shared.get(si);
    if (!mst) return null;
    const a = splitAddr(addr), b = splitAddr(mst.addr);
    return shiftFormula(mst.f, a.row - b.row, colNum(a.col) - colNum(b.col));
  };
  const memo = new Map(), busy = new Set();
  const ERR = (e) => ({ err: e });
  const valueOf = (addr) => {
    if (memo.has(addr)) return memo.get(addr);
    const f = formulaOf(addr);
    let out;
    if (f == null) {
      const c = cells.get(addr);
      out = (c && c.v != null && c.t !== 's' && c.t !== 'str' && c.t !== 'inlineStr' && c.t !== 'e' && Number.isFinite(parseFloat(c.v))) ? parseFloat(c.v) : 0;
    } else {
      if (busy.has(addr)) return ERR('#REF!');
      busy.add(addr);
      try { out = evalExpr(f); } finally { busy.delete(addr); }
    }
    memo.set(addr, out);
    return out;
  };
  const rangeAddrs = (a, b) => {
    const p = splitAddr(a), q = splitAddr(b), out = [];
    for (let r = Math.min(p.row, q.row); r <= Math.max(p.row, q.row); r++) for (let c = Math.min(colNum(p.col), colNum(q.col)); c <= Math.max(colNum(p.col), colNum(q.col)); c++) out.push(colName(c) + r);
    return out;
  };
  function evalExpr(src) {
    const toks = []; const tre = /\s*(?:(\d+\.?\d*)|(SUM|IF)(?=\s*\()|(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)|(<=|>=|<>|[=<>])|([+\-*/(),]))/y;
    let pos = 0, m;
    while (pos < src.length) {
      tre.lastIndex = pos; m = tre.exec(src);
      if (!m) throw new Error('不支援的公式：' + src);
      toks.push(m[1] ? { n: parseFloat(m[1]) } : m[2] ? { f: m[2] } : m[3] ? { ref: m[3].replace(/\$/g, '') } : m[4] ? { cmp: m[4] } : { op: m[5] });
      pos = tre.lastIndex;
    }
    let i = 0;
    const isErr = (x) => x && typeof x === 'object' && x.err;
    const sumArgs = () => {
      let total = 0;
      do {
        i++;                                              // 跳過 '(' 或 ','
        const t = toks[i];
        if (t && t.ref && /:/.test(t.ref)) { const [a, b] = t.ref.split(':'); for (const ad of rangeAddrs(a, b)) { const v = valueOf(ad); if (isErr(v)) return v; total += v; } i++; }
        else { const v = expr(); if (isErr(v)) return v; total += v; }
      } while (toks[i] && toks[i].op === ',');
      i++;                                                // 跳過 ')'
      return total;
    };
    const factor = () => {
      const t = toks[i];
      if (!t) throw new Error('公式不完整：' + src);
      if (t.n !== undefined) { i++; return t.n; }
      if (t.ref) { i++; return valueOf(t.ref); }
      if (t.f === 'SUM') { i++; return sumArgs(); }
      if (t.f === 'IF') {
        // IF(條件, 值1, 值2)：兩個分支都會解析（錯誤是回傳值、不是例外），再依條件挑一個——與 Excel 一樣，沒被選中的分支即使出錯也不影響結果
        i += 2;                                           // 跳過 'IF' 與 '('
        const cond = cmp(); i++;                          // 跳過 ','
        const a = cmp(); i++;
        const b = cmp(); i++;                             // 跳過 ')'
        if (isErr(cond)) return cond;
        return cond ? a : b;
      }
      if (t.op === '(') { i++; const v = cmp(); i++; return v; }
      if (t.op === '-') { i++; const v = factor(); return isErr(v) ? v : -v; }
      throw new Error('公式解析失敗：' + src);
    };
    const term = () => {
      let a = factor();
      while (toks[i] && (toks[i].op === '*' || toks[i].op === '/')) {
        const op = toks[i++].op, b = factor();
        if (isErr(a)) continue; if (isErr(b)) { a = b; continue; }
        a = op === '*' ? a * b : (b === 0 ? ERR('#DIV/0!') : a / b);
      }
      return a;
    };
    const expr = () => {
      let a = term();
      while (toks[i] && (toks[i].op === '+' || toks[i].op === '-')) {
        const op = toks[i++].op, b = term();
        if (isErr(a)) continue; if (isErr(b)) { a = b; continue; }
        a = op === '+' ? a + b : a - b;
      }
      return a;
    };
    /** 比較運算（= <> < > <= >=）：回傳布林，只給 IF 的條件用 */
    const cmp = () => {
      const a = expr();
      const t = toks[i];
      if (!(t && t.cmp)) return a;
      i++;
      const b = expr();
      if (isErr(a)) return a; if (isErr(b)) return b;
      switch (t.cmp) { case '=': return a === b; case '<>': return a !== b; case '<': return a < b; case '>': return a > b; case '<=': return a <= b; default: return a >= b; }
    };
    const r = cmp();
    if (i < toks.length) throw new Error('公式解析未完成：' + src);   // 有沒處理到的 token＝不支援的語法，與其算錯不如直接失敗
    return r;
  }
  const result = new Map();
  for (const [addr, c] of cells) { if (c.fAttrs != null) result.set(addr, valueOf(addr)); }
  return result;
}

/** 把重算結果寫回每個公式儲存格的快取值（數值：移除 t="e"；錯誤：t="e" ＋ 錯誤字串） */
function writeCachedValues(xml, results) {
  // 與 parseCells 相同的儲存格樣式：自閉合的 <c .../> 沒有 inner（不可讓比對往後吃到別格的 </c>）
  return xml.replace(/<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (whole, addr, attrs, inner) => {
    if (inner === undefined || !results.has(addr) || !/<f[\s>]/.test(inner)) return whole;
    const r = results.get(addr);
    const f = /<f[^>]*?(?:\/>|>[\s\S]*?<\/f>)/.exec(inner)[0];
    const a2 = attrs.replace(/\st="\w+"/, '');
    if (r && typeof r === 'object' && r.err) return `<c r="${addr}"${a2} t="e">${f}<v>${r.err}</v></c>`;
    const v = Number.isFinite(r) ? (Object.is(r, -0) ? 0 : r) : 0;
    return `<c r="${addr}"${a2}>${f}<v>${v}</v></c>`;
  });
}

// ── 填表 ────────────────────────────────────────────────────
/** Excel 日期序號（1900 系統）；isoDate＝YYYY-MM-DD */
function excelSerial(isoDate) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(isoDate || ''));
  if (!m) return null;
  return Math.round((Date.UTC(+m[1], +m[2] - 1, +m[3]) - Date.UTC(1899, 11, 30)) / 86400000);
}

/** 說明文字塞進範本的固定欄寬：全形字算 2、其餘算 1，超過就截斷加「…」（範本儲存格不會自動換行，鄰格有內容時會被硬切掉） */
function clipText(s, units) {
  const str = String(s == null ? '' : s);
  let w = 0, out = '';
  for (const ch of str) {
    const cw = ch.codePointAt(0) > 0x2E7F ? 2 : 1;
    if (w + cw > units) return out.replace(/\s+$/, '') + '…';
    w += cw; out += ch;
  }
  return str;
}

/** 把同一區的品項分配到可用的列：列數夠就一列一項；不夠，最後一列放「其餘 N 項合計」 */
function packLines(entries, rows) {
  if (entries.length <= rows.length) return entries.map((e, i) => ({ row: rows[i], ...e }));
  const keep = rows.length - 1;
  const rest = entries.slice(keep);
  const lines = entries.slice(0, keep).map((e, i) => ({ row: rows[i], ...e }));
  lines.push({
    row: rows[keep], desc: `其餘 ${rest.length} 項合計`, qty: null, unitCost: null, price: null,
    revCents: rest.reduce((s, e) => s + e.revCents, 0n), costCents: rest.reduce((s, e) => s + e.costCents, 0n), merged: true,
  });
  return lines;
}

/**
 * 回傳 Promise<Buffer>。
 * q：報價單（items 含 cat/qty/unitPrice/cost；discountType/discountValue；company/projectName/contactName/quoteNo）
 * opts：{ classCodes: [勾選商品的類別代碼], requestedBy: 申請人顯示名稱, issueDate: 'YYYY-MM-DD'（匯出當天，台灣時間）,
 *         contingencyPct: 風險預留 0/5/10/15/20（顧問主管設定；沒設定＝0）}
 */
async function buildQuotePnlExcel(q, opts) {
  q = q || {}; opts = opts || {};
  // 分組標題／小計列不是品項：不佔 PNL 的列、不進分類，也不影響「其餘 N 項合計」的項數
  const src = Array.isArray(q.items) ? q.items.filter(x => x && typeof x === 'object' && QI.isItemRow(x)) : [];
  const items = src.length ? src : [{ desc: '', unit: '式', qty: 1, unitPrice: 0 }];
  const alloc = allocateRevenue(items, q.discountType || 'none', q.discountValue);

  // 依分類分組
  const groups = { consult: [], software: [], hardware: [], other: [] };
  items.forEach((it, i) => {
    const qty = num(it.qty), cost = num(it.cost);
    groups[resolveCat(it, opts.classCodes)].push({
      desc: String(it.desc || '').trim() || '（未填品項說明）', qty, unitCost: cost,
      price: qty > 0 ? toTwd(alloc.revCents[i]) / qty : 0,
      revCents: alloc.revCents[i], costCents: bigCents(qty * cost),
    });
  });

  const zip = await JSZip.loadAsync(fs.readFileSync(TEMPLATE));
  const sh = new SheetXml(await zip.file(SHEET_PATH).async('string'), []);

  // 抬頭／客戶／專案資料（範本的黃底必填格 PM、專案起迄日留白給人工補）
  if (opts.requestedBy) sh.setText('B5', opts.requestedBy);
  const serial = excelSerial(opts.issueDate);
  if (serial) sh.setNumber('H5', serial);
  // B6:C6 是合併格（約 28 個半形寬、不會溢位）；G10 往右溢位到 H10（約 37 寬）→ 超長就截斷，避免被硬切或溢出列印範圍
  if (q.company) sh.setText('B6', clipText(q.company, 24));
  if (q.contactName) sh.setText('B8', clipText(q.contactName, 30));
  if (q.quoteNo) sh.setText('B9', clipText(q.quoteNo, 30));
  if (q.projectName) sh.setText('G10', clipText(q.projectName, 34));

  // Contingency（風險預留）：顧問主管依專案風險預估 0/5/10/15/20（%），以「顧問服務成本」為基準（範本 H89=F89×G89，G89=顧問成本合計）。
  // 沒設定＝0%。範本原本預設 Low 5%，但現在改由顧問主管決定，所以一律明確寫入。
  const risk = CONTINGENCY_PCTS.includes(opts.contingencyPct) ? opts.contingencyPct : 0;
  sh.setNumber('F89', risk / 100);
  sh.setText('B89', RISK_LABELS[risk]);

  // 收入：每區金額寫 H（折扣後）；顧問區另外寫人天、單價（顯示用）
  for (const cat of CATS) {
    const rows = REV_ROWS[cat];
    if (!groups[cat].length) continue;
    if (rows.length === 1) {
      // 軟體收入只有一格可加總：品項彙總成一列，說明寫「首項 等 N 項」
      const g = groups[cat];
      const total = g.reduce((s, e) => s + e.revCents, 0n);
      // 這一格右邊緊鄰「100%」（D23），B~C 約 24 個半形寬
      sh.setText('B' + rows[0], g.length === 1 ? clipText(g[0].desc, 24) : `${clipText(g[0].desc, 14)} 等 ${g.length} 項`);
      sh.setNumber('H' + rows[0], toTwd(total));
      continue;
    }
    for (const ln of packLines(groups[cat], rows)) {
      sh.setText('B' + ln.row, clipText(ln.desc, cat === 'consult' ? 50 : 60));
      if (cat === 'consult' && !ln.merged) { sh.setNumber('F' + ln.row, ln.qty); sh.setNumber('G' + ln.row, Math.round(ln.price * 100) / 100); }
      sh.setNumber('H' + ln.row, toTwd(ln.revCents));
    }
  }

  // 成本：顧問寫 D=單位成本、F=人天（H=D×F 為公式）；其餘寫 F=數量、G=單價（H=F×G）；超出列數的合計進最後一列（數量 1）
  for (const cat of CATS) {
    const rows = COST_ROWS[cat];
    for (const ln of packLines(groups[cat], rows)) {
      const unit = ln.merged ? toTwd(ln.costCents) : ln.unitCost, qty = ln.merged ? 1 : ln.qty;
      // A 欄靠左（範本已把顧問成本列的置中改為靠左），文字可溢位到右邊空白的 B、C 欄（顧問區到 D「Rates」之前約 45 寬）
      sh.setText('A' + ln.row, clipText(ln.desc, cat === 'consult' ? 40 : 60));
      if (cat === 'consult') { sh.setNumber('D' + ln.row, unit); sh.setNumber('F' + ln.row, qty); }
      else { sh.setNumber('F' + ln.row, qty); sh.setNumber('G' + ln.row, unit); }
    }
  }
  // 沒有軟體成本品項時，範本預設的「數量 1」維持不動（H59＝1×空白＝0）

  let sheetXml = sh.xml;
  // 重算失敗（例如日後有人在範本加了這裡不支援的函式）不能讓匯出整個失敗：改成不更新快取值，靠 fullCalcOnLoad 讓 Excel 開檔時重算
  try { sheetXml = writeCachedValues(sheetXml, evalWorkbookSheet(sheetXml)); }
  catch (e) { console.warn('[quotePnlExcel] 公式快取值重算失敗，改由 Excel 開檔時重算：', e && e.message); }
  zip.file(SHEET_PATH, sheetXml);

  // 活頁簿：開檔時 Excel 全面重算；移除 Excel 另存時帶出的本機路徑（absPath）
  let wbx = await zip.file('xl/workbook.xml').async('string');
  wbx = wbx.replace(/<calcPr([^>]*?)(\/?)>/, (m, a, s) => /fullCalcOnLoad/.test(a) ? m : `<calcPr${a} fullCalcOnLoad="1"${s}>`);
  wbx = wbx.replace(/<mc:AlternateContent\b[^>]*>\s*<mc:Choice\b[^>]*>\s*<x15ac:absPath\b[^>]*\/>\s*<\/mc:Choice>\s*<\/mc:AlternateContent>/g, () => '');
  zip.file('xl/workbook.xml', wbx);
  const coreFile = zip.file('docProps/core.xml');
  if (coreFile) {
    const core = (await coreFile.async('string'))
      .replace(/<dc:creator>[^<]*<\/dc:creator>/, () => '<dc:creator>ITTS</dc:creator>')
      .replace(/<cp:lastModifiedBy>[^<]*<\/cp:lastModifiedBy>/, () => '<cp:lastModifiedBy>ITTS</cp:lastModifiedBy>');
    zip.file('docProps/core.xml', core);
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/*
 * 範本與原檔（報價單範本.xlsx「PNL (C)」）的差異（Steven 選「修正公式」）：
 *   - 只留 PNL 一頁；移除對「報價單」頁的跨表引用（F17、H17）、7 個外部連結、473 個已定義名稱（只留列印範圍）
 *   - H19 顧問收入合計：SUM(H16:H17) → SUM(H16:H18)（16~18 都是輸入列）
 *   - 軟體成本：H61=G59（單價欄、不乘數量）→ H59=F59*G59、H60=F60*G60、H61=SUM(H59:H60)
 *   - 硬體成本：H66=SUM(G64:G65)（加單價欄）→ H64=F64*G64、H65=F65*G65、H66=SUM(H64:H65)
 *   - 其他費用：H83、H84 由手填改為 F×G
 *   - 隱藏列 H42:H48 的算法不一（有的乘 Months/Year、有的不乘）→ 統一 D×F
 *   - 底部彙總表「其他」成本 F103：H90+H85 → H90+H85+H79（原本漏掉差旅費）
 *   - H8、H9 的殘留空白字串清空
 * 重建範本：scripts/build-pnl-template.ps1（Excel COM）＋ scripts/post-pnl-template.js；改版面時要同步這裡的 REV_ROWS / COST_ROWS。
 *   後續補強（審查後）：彙總表與總毛利率加 IF 除零保護（某類沒有收入顯示 0.00%）、F18 人天欄改一般格式、A38:A39 改靠左、
 *   移除兩則常駐註解與分頁預覽、移除網路印表機設定。
 *   注意：範本預設 Contingency「Low 5%」（顧問成本×5%）會計入總成本，所以 PNL 毛利率會比 CRM 簽核用的毛利率低——這是公司原範本的設計。
 */
module.exports = {
  buildQuotePnlExcel, CATS, CAT_LABELS, CONTINGENCY_PCTS,
  _internal: { resolveCat, catByUnit, allocateRevenue, evalWorkbookSheet, writeCachedValues, excelSerial, packLines, clipText, parseCells, REV_ROWS, COST_ROWS },
};
