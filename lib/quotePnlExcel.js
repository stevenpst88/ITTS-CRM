'use strict';
/**
 * 毛利分析（PNL）Excel —— **內部用，不得給客戶**（含成本單價、成本小計、毛利、毛利率）。
 * 與「給客戶的報價單」（lib/quoteExcel.js）分開匯出；權限由 GET /api/quotations/:id/export-pnl 把關。
 *
 * 版面：公司原有的「Profitability Calculation Worksheet for Implementation」（2026-10 新版 PNL 範本，工作表名 PNL）。
 * 範本 templates/pnl_template.xlsx 是用 Excel 從原檔整理出的「單頁、無外部連結、無垃圾已定義名稱、無隱藏列」乾淨版，
 * 並修正了原範本的公式缺陷、擴充了各區資料列容量（重建方式見 scripts/build-pnl-template.ps1；列號對照見 LAYOUT）。
 * 這裡只用 JSZip 直接改 sheet1.xml 的儲存格（沿用範本樣式），不經 SheetJS（社群版寫檔會丟掉所有樣式）。
 * 公式保持為公式，使用者在 Excel 內改數字會連動；同時把「公式快取值」重算寫進檔案
 * （預覽器／手機／雲端檢視器只認快取值），並設 fullCalcOnLoad 讓 Excel 開檔時再算一次。
 * 範本公式只能用 + - * / ( ) SUM IF 與儲存格參照（見 evalWorkbookSheet），日後在範本加 ROUND／MIN／MAX／IFERROR 等函式要先補評估器。
 *
 * 資料對應：
 *   收入（PNL 分 顧問服務／軟體／硬體／其他 四區，每區一個「總額」輸入格 H14／H23／H27／H31，請款階段表留白給人工補）：
 *   - 每個品項有分類 cat（consult／software／hardware／other；空白＝自動：依報價單勾選的商品類別，單一類別就全歸該類，
 *     多類別再依單位推測，仍無法判斷歸「其他」）。
 *   - 收入用「折扣後金額」：總額直接採用簽核的折扣後營收（computeFinancials(q).revenueCents；算不出來才退回舊算法），
 *     依各品項小計比例攤到品項（以「分」做最大餘數法，各品項加總＝專案優惠價，不差一分），再依分類加總寫進該區總額格。
 *     品項小計與簽核同一套規則（十進位精算、half-up 到分、數量空白或 0 當 1），所以 PNL 總收入與印花稅基數和簽核逐分一致。
 *   成本（兩種來源，由 q.costLines 欄位是否存在決定，與 lib/quoteApproval.js computeFinancials 同一條規則）：
 *   - 新式（q.costLines 是陣列，含 []）：成本明細獨立於客戶品項。顧問／軟體／硬體／差旅／其他費用五區各自逐列填入
 *     （項目、委外廠商／供應商、說明、數量、單價），印花稅固定在範本的印花稅列（整數元＝折扣後營收×0.1% 四捨五入，
 *     與伺服器 CL.stampDollars 同一個函式），沒有印花稅列＝不計入。金額以「分」精算，與簽核毛利的成本逐分一致
 *     （PNL 總成本 − Contingency ＝ 簽核成本；PNL 毛利另外含 Contingency）。
 *   - 舊式（沒有 costLines）：成本在 items[].cost（不打折：成本單價 × 數量；computeFinancials 成功時數量空白／0 與簽核一樣當 1），依品項分類填入對應區域；
 *     沒有廠商、說明、差旅、交際費、印花稅這些資料，行為等同改版前，只是列位置換成新範本的位置。
 *   - 某一區的項目比可填列數多時，最後一列放「其餘 N 項合計」（數量 1、單價＝其餘項目的成本合計，金額與伺服器一致）。
 * 範本沒有、系統也沒有的資料（PM、PCM、專案起迄日、請款階段、顧問姓名…）一律留白，給人工補。
 */
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { _internal: { SheetXml } } = require('./quoteExcel');
const QI = require('./quoteItems');
const CL = require('./quoteCostLines');
const QA = require('./quoteApproval');   // computeFinancials：PNL 總收入（折扣後）直接採用簽核營收

const TEMPLATE = path.join(__dirname, '..', 'templates', 'pnl_template.xlsx');
const SHEET_PATH = 'xl/worksheets/sheet1.xml';

/** 風險預留（%）可選值與範本 B109「Project Risk level」的英文等級字樣；與 lib/quoteRoutes.js 的 CONTINGENCY_PCTS 一致 */
const CONTINGENCY_PCTS = [0, 5, 10, 15, 20];
const RISK_LABELS = { 0: 'None', 5: 'Low', 10: 'Medium', 15: 'High', 20: 'Very High' };
const CATS = ['consult', 'software', 'hardware', 'other'];
const CAT_LABELS = { consult: '顧問服務', software: '軟體', hardware: '硬體', other: '其他' };
/** 商品類別（簽核設定的 productClasses.cls）→ PNL 四區 */
const CLASS_TO_CAT = { consult: 'consult', software: 'software', hardware: 'hardware', crm: 'other', mdm: 'other', ot: 'other', other: 'other' };
/** 品項分類（hardware）→ 成本明細分類（hw）；其餘同名 */
const ITEM_CAT_TO_COST_CAT = { consult: 'consult', software: 'software', hardware: 'hw', other: 'other' };

const range = (a, b) => { const out = []; for (let r = a; r <= b; r++) out.push(r); return out; };
/**
 * 新版 PNL 範本（templates/pnl_template.xlsx）的儲存格位置。改範本版面時要同步這裡（舊版面對照見 scripts/build-pnl-template.ps1 檔頭）。
 *   revenue：四區收入「總額」輸入格（H22 是 =H25 的連動公式，軟體收入寫 H23，不可覆寫 H22）
 *   costRows：各區可填的資料列。顧問區 A=項目 C=委外廠商 D=單位成本(Rates) F=數量(Base Days) G=1(Months/Year)；
 *             軟體／硬體 A=項目 C=供應商 D=說明 F=數量 G=單價；差旅／其他 A=項目 D=說明 F=數量 G=單價；H 是範本公式（=D*F*G 或 F*G）
 *   other：交際費用標籤列 99、印花稅列 100（固定列，F=1、G=整數元印花稅）、自由列 101–104 → 一般其他費用列用 99、101–104
 *   顧問區 61–63（第 3 組）沒有資料格式也沒有逐列公式，不寫
 */
const LAYOUT = {
  header: { requestedBy: 'B5', issueDate: 'H5', company: 'B6', contact: 'B8', quoteNo: 'B9', projectName: 'G10' },
  revenue: { consult: 'H14', software: 'H23', hardware: 'H27', other: 'H31' },
  costRows: {
    consult: [...range(37, 49), ...range(51, 57)],
    software: range(68, 73),
    hw: range(77, 82),
    travel: range(87, 94),
    other: [99, 101, 102, 103, 104],
  },
  stampRow: 100,
  contingency: { level: 'B109', pct: 'F109' },
};
/** 各區未使用的資料列要清空的欄（C＝廠商；D＝說明，顧問區的 D 是單位成本；G 在顧問區是 Months/Year＝1，不清） */
const CLEAR_COLS = {
  consult: ['A', 'B', 'C', 'D', 'F'],
  software: ['A', 'C', 'D', 'F', 'G'],
  hw: ['A', 'C', 'D', 'F', 'G'],
  travel: ['A', 'D', 'F', 'G'],
  other: ['A', 'D', 'F', 'G'],
};

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
/** 數量空白或 0 → true（伺服器 computeFinancials 把這種列當 1；存檔時 qty≤0 也會被轉成 1）。只在 computeFinancials 成功（數量是合法非負數）時使用 */
function qtyIsEmptyOrZero(x) {
  if (x === null || x === undefined) return true;
  if (typeof x === 'number') return x === 0;
  const t = String(x).trim().replace(/,/g, '');
  return t === '' || !/[1-9]/.test(t.replace(/[eE].*$/, ''));   // 尾數全是 0（0、0.00、0e5…）
}

/**
 * 品項小計（分，BigInt）：與 lib/quoteApproval.js computeFinancials 的每列金額完全相同 ——
 * 數量空白／0 當 1、數量×單價用十進位字串精算（不經浮點乘法）、half-up 到分（CL.lineCentsBig 與 computeFinancials 是同一套 parseDec／divRound 規則）。
 */
function itemListCents(it) {
  const o = (it && typeof it === 'object') ? it : {};
  return CL.lineCentsBig({ qty: qtyIsEmptyOrZero(o.qty) ? 1 : o.qty, unitCost: o.unitPrice });
}

/** computeFinancials 成功時回傳每個品項的小計（分）；小計加總對不上 fin.subtotalCents（理論上不會發生）或任何一列算不出來 → null（呼叫端退回舊算法） */
function exactListCents(items, fin) {
  if (!fin || !fin.ok) return null;
  try {
    const list = items.map(itemListCents);
    return list.reduce((s, x) => s + x, 0n) === BigInt(fin.subtotalCents) ? list : null;
  } catch (e) { return null; }
}

/**
 * 回傳 { listCents: [..], revCents: [..], totalCents } （BigInt）。revCents 各品項加總＝折扣後總額 totalCents。
 * percent：discountValue 是「實收百分比」（90＝九折，0<v<100）；amount：議價後的未稅總價；其他：不折扣。
 * fin＝computeFinancials(q) 的結果：成功（ok）時，總額直接採用簽核營收 fin.revenueCents、品項小計用與它相同的每列金額，
 *   所以 PNL 的總收入、各類收入、印花稅基數都與簽核逐分一致；fin 沒給或不成功（輸入有問題、營收為 0 等）時退回舊算法（浮點乘法、折扣百分比取 3 位小數）。
 */
function allocateRevenue(items, discountType, discountValue, fin) {
  const exact = exactListCents(items, fin);
  let list, L, A;
  if (exact) {
    list = exact;
    L = list.reduce((s, x) => s + x, 0n);
    A = BigInt(fin.revenueCents);
  } else {
    list = items.map(it => bigCents(num(it.qty) * num(it.unitPrice)));
    L = list.reduce((s, x) => s + x, 0n);
    A = L;
    const v = num(discountValue);
    if (discountType === 'percent' && v > 0 && v < 100) A = (L * BigInt(Math.round(v * 1000)) + 50000n) / 100000n;   // 四捨五入到分
    else if (discountType === 'amount' && v > 0) A = bigCents(v);
  }
  const isExact = !!exact;   // true＝總額與每列金額都和簽核（computeFinancials）一致；collectCosts 的舊式成本列也依此套用同一套數量規則
  if (L === 0n) return { listCents: list, revCents: list.map(() => 0n), totalCents: A, exact: isExact };
  const base = list.map(x => (A * x) / L);
  const rem = list.map((x, i) => (A * x) % L);
  let left = A - base.reduce((s, x) => s + x, 0n);
  const order = list.map((_, i) => i).sort((a, b) => (rem[b] > rem[a] ? 1 : rem[b] < rem[a] ? -1 : a - b));
  for (const i of order) { if (left <= 0n) break; base[i] += 1n; left -= 1n; }
  return { listCents: list, revCents: base, totalCents: A, exact: isExact };
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

/** 單行文字：換行／tab 換成空白（儲存格不自動換行，換行字元在 Excel 裡會讓一格顯示成怪樣子）、去頭尾空白 */
const oneLine = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim();

/**
 * 把同一區的成本項目（entry：{desc, vendor, note, unit, qty, unitCost, cents(BigInt)}）分配到可用的列：
 * 列數夠就一列一項；不夠，最後一列放「其餘 N 項合計」（數量 1、單價＝其餘各項成本（分）的合計，與伺服器逐分一致；沒有廠商與說明）。
 */
function packLines(entries, rows) {
  if (entries.length <= rows.length) return entries.map((e, i) => ({ row: rows[i], ...e }));
  const keep = rows.length - 1;
  const rest = entries.slice(keep);
  const lines = entries.slice(0, keep).map((e, i) => ({ row: rows[i], ...e }));
  const cents = rest.reduce((s, e) => s + e.cents, 0n);
  lines.push({ row: rows[keep], desc: `其餘 ${rest.length} 項合計`, vendor: '', consultant: '', note: '', unit: '式', qty: 1, unitCost: toTwd(cents), cents, merged: true });
  return lines;
}

/** 成本項目的「說明」欄文字：note 優先；note 空白且單位不是「式」才把單位當說明（範本沒有單位欄位） */
function noteText(e) {
  if (e.note) return oneLine(e.note);
  const u = oneLine(e.unit);
  return u && u !== '式' ? u : '';
}
/**
 * 顧問區沒有「說明」欄（D 是單位成本 Rates）：說明併進項目欄 A，格式「項目（說明）」；沒有說明時，單位若不是「式」也不是「人天」
 * （F 欄表頭 Base Days 本來就是人天）就併成「項目（單位）」。A 欄只能溢出到空白的 B（有委外廠商時 C 也被佔住），所以依是否有廠商限制寬度。
 */
function consultLabel(e) {
  const d = oneLine(e.desc), n = oneLine(e.note), u = oneLine(e.unit);
  const tail = n ? n : (u && u !== '式' && u !== '人天' ? u : '');
  // A 欄寬約 23、B（顧問姓名）約 10、C（委外廠商）約 33：A 只能溢出到「空白的」B／C。有顧問姓名＝A 只剩自己的寬度；沒有姓名但有廠商＝A＋B；兩個都沒有＝A＋B＋C
  return clipText(tail ? `${d}（${tail}）` : d, e.consultant ? 21 : (e.vendor ? 30 : 58));
}
/** 資料列 A 欄（項目）寬度預算：右邊的 B 空白，C 有廠商就只能用 A＋B，否則可以溢出到 C（D 一定有內容） */
const labelA = (e) => clipText(oneLine(e.desc), e.vendor ? 30 : 58);

/** 把一個項目寫進它所在的資料列（顧問區：A 項目＋說明、C 委外廠商、D 單位成本、F 數量、G=1；其餘區：A、C 供應商、D 說明、F 數量、G 單價） */
function writeCostRow(sh, cat, ln) {
  const r = ln.row;
  const setOrClear = (ref, text) => { if (text) sh.setText(ref, text); else sh.clearCell(ref); };
  if (cat === 'consult') {
    sh.setText('A' + r, ln.merged ? clipText(ln.desc, 58) : consultLabel(ln));
    // B＝Consultant（自家顧問姓名）、C＝Free Lancer(委外廠商)；B 寬約 10，C 沒有廠商時 B 可溢出到 C。「其餘 N 項合計」的 B／C 留白
    setOrClear('B' + r, clipText(oneLine(ln.consultant), ln.vendor ? 9 : 40));
    setOrClear('C' + r, clipText(oneLine(ln.vendor), 30));
    sh.setNumber('D' + r, ln.unitCost);
    sh.setNumber('F' + r, ln.qty);
    sh.setNumber('G' + r, 1);                                   // Months/Year：範本公式 H=D*F*G，不可留空
    return;
  }
  sh.setText('A' + r, ln.merged ? clipText(ln.desc, 58) : labelA(ln));
  if (cat === 'software' || cat === 'hw') setOrClear('C' + r, clipText(oneLine(ln.vendor), 30));
  setOrClear('D' + r, clipText(noteText(ln), 64));              // D 右邊的 E 欄是空白，說明可溢出到 E（D＋E 約 68 個半形寬）
  sh.setNumber('F' + r, ln.qty);
  sh.setNumber('G' + r, ln.unitCost);
}

/**
 * 依 q 的種類產出各成本區的項目清單：{ areas:{consult,software,hw,travel,other:[entry]}, stamp:null|{dollars} }。
 * 新式（q.costLines 是陣列）：逐列取成本明細，金額以 CL.lineCentsBig 精算（與 computeFinancials 同一套取整），印花稅列另外處理；
 * 舊式：成本取自 items[].cost（依品項分類），行為等同改版前。
 */
function collectCosts(q, items, alloc, classCodes) {
  const areas = { consult: [], software: [], hw: [], travel: [], other: [] };
  let stamp = null;
  if (CL.hasCostLines(q)) {
    for (const l of CL.effectiveLines(q, alloc.totalCents)) {
      if (!l || typeof l !== 'object') continue;
      if (CL.isStamp(l)) { stamp = { dollars: CL.stampDollars(alloc.totalCents) }; continue; }
      const cat = CL.CATS.includes(l.cat) ? l.cat : 'other';      // 與 CL.totalsByCat 相同：壞資料的分類一律當「其他」
      let cents; try { cents = CL.lineCentsBig(l); } catch (e) { cents = 0n; }   // 資料壞掉（伺服器存檔時就會擋）：不讓匯出整個失敗
      areas[cat].push({ desc: oneLine(l.desc) || '（未填項目）', vendor: oneLine(l.vendor), consultant: cat === 'consult' ? oneLine(l.consultant) : '', note: l.note, unit: l.unit, qty: num(l.qty), unitCost: num(l.unitCost), cents });
    }
    return { areas, stamp };
  }
  // 舊式成本：alloc.exact（computeFinancials 成功）時，數量空白／0 與簽核一樣當 1、列金額用同一套十進位精算；否則維持舊行為（浮點乘法、空白＝0）
  const serverRule = alloc.exact === true;
  items.forEach((it) => {
    const qRaw = serverRule && qtyIsEmptyOrZero(it.qty) ? 1 : it.qty;
    const qty = num(qRaw), cost = num(it.cost);
    let cents = bigCents(qty * cost);
    if (serverRule) { try { cents = CL.lineCentsBig({ qty: qRaw, unitCost: it.cost }); } catch (e) { /* 保留浮點結果 */ } }
    areas[ITEM_CAT_TO_COST_CAT[resolveCat(it, classCodes)]].push({
      desc: oneLine(it.desc) || '（未填品項說明）', vendor: '', note: '', unit: '', qty, unitCost: cost, cents,
    });
  });
  return { areas, stamp };
}

/**
 * 回傳 Promise<Buffer>。
 * q：報價單（items 含 cat/qty/unitPrice/cost；discountType/discountValue；company/projectName/contactName/quoteNo；
 *     costLines＝新式成本明細，沒有這個欄位＝舊式，成本取 items[].cost）
 * opts：{ classCodes: [勾選商品的類別代碼], requestedBy: 申請人顯示名稱, issueDate: 'YYYY-MM-DD'（匯出當天，台灣時間）,
 *         contingencyPct: 風險預留 0/5/10/15/20（顧問主管設定；沒設定＝0）}
 */
async function buildQuotePnlExcel(q, opts) {
  q = q || {}; opts = opts || {};
  // 分組標題／小計列不是品項：不佔 PNL 的列、不進分類，也不影響「其餘 N 項合計」的項數
  const src = Array.isArray(q.items) ? q.items.filter(x => x && typeof x === 'object' && QI.isItemRow(x)) : [];
  const items = src.length ? src : [{ desc: '', unit: '式', qty: 1, unitPrice: 0 }];
  // 簽核用的財務計算（BigInt 分）：成功時，總收入＝fin.revenueCents、品項小計＝與它同一套的每列金額（數量空白／0 當 1）→ 與簽核營收、印花稅基數逐分一致；不成功就退回舊算法
  let fin = null;
  try { fin = QA.computeFinancials(q); } catch (e) { fin = null; }
  const alloc = allocateRevenue(items, q.discountType || 'none', q.discountValue, fin);

  // 收入：依品項分類加總折扣後金額（以「分」計，各類加總＝專案優惠價）
  const revByCat = {}, present = {};
  items.forEach((it, i) => {
    const cat = resolveCat(it, opts.classCodes);
    revByCat[cat] = (revByCat[cat] || 0n) + alloc.revCents[i];
    present[cat] = true;
  });
  const { areas, stamp } = collectCosts(q, items, alloc, opts.classCodes);

  const zip = await JSZip.loadAsync(fs.readFileSync(TEMPLATE));
  const sh = new SheetXml(await zip.file(SHEET_PATH).async('string'), []);

  // 抬頭／客戶／專案資料（範本的黃底必填格 PM、專案起迄日留白給人工補）
  const H = LAYOUT.header;
  if (opts.requestedBy) sh.setText(H.requestedBy, opts.requestedBy);
  const serial = excelSerial(opts.issueDate);
  if (serial) sh.setNumber(H.issueDate, serial);
  // B6:C6 是合併格（約 28 個半形寬、不會溢位）；G10 往右溢位到 H10（約 37 寬）→ 超長就截斷，避免被硬切或溢出列印範圍
  if (q.company) sh.setText(H.company, clipText(q.company, 24));
  if (q.contactName) sh.setText(H.contact, clipText(q.contactName, 30));
  if (q.quoteNo) sh.setText(H.quoteNo, clipText(q.quoteNo, 30));
  if (q.projectName) sh.setText(H.projectName, clipText(q.projectName, 34));

  // Contingency（風險預留）：顧問主管依專案風險預估 0/5/10/15/20（%），以「顧問服務成本」為基準（範本 H109=F109×G109，G109=H60 顧問成本合計）。
  // 沒設定＝0%。範本原本預設 Low 5%，但現在改由顧問主管決定，所以一律明確寫入。
  const risk = CONTINGENCY_PCTS.includes(opts.contingencyPct) ? opts.contingencyPct : 0;
  sh.setNumber(LAYOUT.contingency.pct, risk / 100);
  sh.setText(LAYOUT.contingency.level, RISK_LABELS[risk]);

  // 收入：每區只寫「總額」格（折扣後）；請款階段表留白給人工補。H22 是 =H25 的連動公式，軟體收入寫 H23
  for (const cat of CATS) {
    if (present[cat]) sh.setNumber(LAYOUT.revenue[cat], toTwd(revByCat[cat]));
  }

  // 成本：逐區把項目放進資料列；超過列數的項目合計進最後一列；沒用到的列清空（避免殘留範本預設文字）
  for (const cat of Object.keys(areas)) {
    const rows = LAYOUT.costRows[cat];
    const lines = packLines(areas[cat], rows);
    for (const ln of lines) writeCostRow(sh, cat, ln);
    for (const r of rows.slice(lines.length)) for (const col of CLEAR_COLS[cat]) sh.clearCell(col + r);
  }
  // 印花稅列（固定在範本的印花稅列，F=1、G＝整數元；範本的 G 是 =H13*0.001 未取整的公式，所以用算好的整數元覆寫）。沒有印花稅列＝不計入：清空
  if (stamp) { sh.setNumber('F' + LAYOUT.stampRow, 1); sh.setNumber('G' + LAYOUT.stampRow, stamp.dollars); }
  else for (const col of ['A', 'F', 'G']) sh.clearCell(col + LAYOUT.stampRow);

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
 * 範本與原檔（2026-10 新版 PNL 範本「PNL 」工作表）的差異（採「修正公式」）：
 *   - 只留 PNL 一頁；移除外部連結、已定義名稱（只留列印範圍 A1:H126）、隱藏的報價單工作表、註解、印表機設定、本機路徑
 *   - 清空範例資料（人名、廠商、費率、基本天數、差旅範例列、範例金額）；範本不含任何客戶資料
 *   - 擴充資料列容量：軟體 2→6、硬體 1→6、其他費用 3→6（顧問 20 列、差旅 8 列不變）
 *   - 公式修正：H13 總收入補上其他收入 H31；顧問小計補第 49 列；顧問各列統一 =D*F*G（G 預設 1）；軟體／硬體／差旅／其他各列補 =F*G；
 *     毛利率與彙總表毛利% 加 IF 除零保護；F 欄加寬（避免彙總表「其他」成本顯示 #######）；
 *     差旅 F87:F94、其他費用 F99／F101:F104 的「數量/次數」格式由整數 #,##0 改 General（否則 qty 1.5 顯示 2、0.1 顯示 0；印花稅列 F100 維持 0.000）
 *   - 顧問成本（含差旅）併在彙總表的顧問欄 C123（=H60+H64+H95）；其他欄 F123＝Contingency＋其他費用（含印花稅）
 * 重建範本：scripts/build-pnl-template.ps1（Excel COM，來源檔自備、不入 repo）＋ scripts/post-pnl-template.js；改版面時要同步上面的 LAYOUT。
 *   注意：Contingency（顧問成本×風險係數）會計入 PNL 總成本，所以 PNL 毛利率會比 CRM 簽核用的毛利率低——這是公司原範本的設計；
 *   除此之外兩邊成本逐分一致（新式單：PNL 總成本 H34 − Contingency H109 ＝ computeFinancials 的成本，含印花稅）。
 */
module.exports = {
  buildQuotePnlExcel, CATS, CAT_LABELS, CONTINGENCY_PCTS,
  _internal: { resolveCat, catByUnit, allocateRevenue, itemListCents, qtyIsEmptyOrZero, evalWorkbookSheet, writeCachedValues, excelSerial, packLines, clipText, parseCells, collectCosts, consultLabel, noteText, LAYOUT, CLEAR_COLS },
};
