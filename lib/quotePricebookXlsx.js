/**
 * 報價牌價簿的 Excel 匯出／範本／匯入預覽（純函式模組，不碰 db／express；路由在 lib/quoteRoutes.js）。
 *
 * 檔案結構（業主決定）：與後台分頁相同，每個 BU 一個工作表，名稱必須「完全等於」ERP／ITS／MDM／CRM，外加一個「說明」工作表（匯入時不讀）。
 * 匯入規則（業主決定）：依「BU＋項目名稱」合併、永不刪除。
 *   · 名稱比對用與伺服器重複檢查相同的 nameKey（NFKC＋空白收合＋trim＋小寫，lib/quotePricebook.js），範圍是「同一個 BU 內」；
 *   · 同 BU 同名 → 更新牌價／成本（「啟用」欄有填才更新），保留原項目的 id 與順序（名稱維持原樣，不因大小寫／全半形差異改名）；
 *   · 該 BU 的新名稱 → 附加在該 BU 最後（遵守每個 BU 60 項上限，超過的列報錯）；檔案沒提到的既有項目原封不動；檔案缺少某個 BU 的工作表 → 該 BU 不動；
 *   · 其他名稱的工作表一律略過（預覽列出）。
 *   · 預覽（buildPreview）只算結果、不寫入任何東西；真正寫入走既有的 PUT /api/admin/quote-pricebook（同一套驗證＋樂觀並行＋稽核）。
 *
 * 檔案安全（不信任上傳的 xlsx）：
 *   · SheetJS 0.18.5 的 zip 讀取器「信任本地檔頭」並一次解壓全部項目（zip bomb 風險），所以原始檔案絕不直接交給它：先自己讀 zip 中央目錄，
 *     逐一用 zlib maxOutputLength 做「有界解壓」（單項 6 MB、總量 12 MB、最多 500 項；不信任標頭宣告的大小；加密／ZIP64／同名項目／路徑跳脫一律拒絕），
 *     通過後用 JSZip 重新封裝成乾淨的 zip（所有大小都是我們自己產生的）才交給 SheetJS 解析；解析選項關閉公式／樣式／HTML，sheetRows 限制列數；
 *   · 拒絕含 DOCTYPE／ENTITY 的 XML、分頁名稱為 __proto__／constructor／prototype 的活頁簿；
 *   · 所有字串只當純文字（不求值、不進 HTML；顯示端負責轉義）。
 *
 * 匯出的公式注入防護：名稱儲存格一律寫成明確的字串型（t="inlineStr"），開頭是 = + - @ Tab CR 的再加 quotePrefix 樣式
 * （Excel 的「文字前綴」：顯示原文、編輯時不會被轉成公式）。內容本身不加任何字元，所以匯入端不需要（也不會）剝除任何東西，往返逐位元相同。
 */
'use strict';
const zlib = require('zlib');
const JSZip = require('jszip');
const XLSX = require('xlsx');
const PB = require('./quotePricebook');

const SHEET_HELP = '說明';
const HEADERS = ['項目名稱', '牌價（元/人天）', '成本（元/人天）', '啟用'];

const MAX_FILE_BYTES = 2 * 1024 * 1024;       // 上傳檔案大小上限
const MAX_DATA_ROWS = 500;                    // 非空白資料列上限
const MAX_SCAN_ROWS = 5000;                   // 實體列掃描上限（空白格式列不算資料列，但不無限掃）
const MAX_HEADER_SCAN = 10;                   // 標題列最多出現在前 10 列
const MAX_COLS = 50;                          // 找標題時最多看前 50 欄
const MAX_ENTRY_BYTES = 6 * 1024 * 1024;      // 單一部件解壓後上限
const MAX_TOTAL_BYTES = 12 * 1024 * 1024;     // 白名單部件解壓後總量上限
const MAX_ZIP_ENTRIES = 500;                  // zip 內項目數上限

const bad = (code, message, status, extra) => ({ ok: false, status: status || 400, code, message, extra });

// ═════════════════ 安全解壓（有界）═════════════════
const REQUIRED = ['[Content_Types].xml', 'xl/workbook.xml'];

/** 讀 zip 中央目錄（不解壓）→ 項目清單；ZIP64、結構不合理一律回 null */
function readCentralDirectory(buf) {
  if (!Buffer.isBuffer(buf) || buf.length < 22) return null;
  let eocd = -1;
  const lo = Math.max(0, buf.length - 22 - 65535);
  for (let i = buf.length - 22; i >= lo; i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) return null;
  const total = buf.readUInt16LE(eocd + 10);
  const cdSize = buf.readUInt32LE(eocd + 12);
  const cdOff = buf.readUInt32LE(eocd + 16);
  if (total === 0xFFFF || cdSize === 0xFFFFFFFF || cdOff === 0xFFFFFFFF) return null;   // ZIP64
  if (total > MAX_ZIP_ENTRIES || cdOff + cdSize > eocd) return null;
  const entries = [];
  let p = cdOff;
  for (let n = 0; n < total; n++) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== 0x02014b50) return null;
    const flags = buf.readUInt16LE(p + 8), method = buf.readUInt16LE(p + 10);
    const comp = buf.readUInt32LE(p + 20), uncomp = buf.readUInt32LE(p + 24);
    const nl = buf.readUInt16LE(p + 28), el = buf.readUInt16LE(p + 30), cl = buf.readUInt16LE(p + 32);
    const lho = buf.readUInt32LE(p + 42);
    if (p + 46 + nl + el + cl > buf.length) return null;
    const name = buf.slice(p + 46, p + 46 + nl).toString('utf8');
    entries.push({ name, flags, method, comp, uncomp, lho });
    p += 46 + nl + el + cl;
  }
  return entries;
}

/** 有界解壓一個項目。失敗（加密、方法不支援、超過上限、大小與宣告不符、資料不完整）回 null */
function inflateEntry(buf, e) {
  if (e.flags & 1) return null;                                   // 加密
  if (e.comp === 0xFFFFFFFF || e.uncomp === 0xFFFFFFFF) return null;
  if (e.uncomp > MAX_ENTRY_BYTES) return null;
  const lho = e.lho;
  if (lho + 30 > buf.length || buf.readUInt32LE(lho) !== 0x04034b50) return null;
  const start = lho + 30 + buf.readUInt16LE(lho + 26) + buf.readUInt16LE(lho + 28);
  const end = start + e.comp;
  if (end > buf.length) return null;
  const raw = buf.slice(start, end);
  let out;
  try {
    if (e.method === 0) out = raw;
    else if (e.method === 8) out = zlib.inflateRawSync(raw, { maxOutputLength: MAX_ENTRY_BYTES });
    else return null;
  } catch (_) { return null; }
  if (out.length !== e.uncomp || out.length > MAX_ENTRY_BYTES) return null;
  return out;
}

const BAD_NAME_RE = /(^|\/)\.\.(\/|$)|^\/|\\/;   // 路徑跳脫（..）、絕對路徑、反斜線
const SHEET_PART_RE = /^xl\/worksheets\/[^/]+\.xml$/;

/** 檢查＋有界解壓全部項目＋重新封裝成乾淨的 xlsx Buffer。回 {ok:true, buf} 或 {ok:false,...} */
async function safeRepack(buf) {
  const E = bad('BAD_XLSX', '檔案不是有效的 Excel 活頁簿（.xlsx），請改用本頁的範本或匯出檔');
  if (!Buffer.isBuffer(buf) || buf.length < 4 || buf.readUInt32LE(0) !== 0x04034b50) return E;   // PK 魔術字
  const entries = readCentralDirectory(buf);
  if (!entries || !entries.length) return E;
  const seen = new Set();
  for (const e of entries) {
    if (seen.has(e.name) || BAD_NAME_RE.test(e.name)) return E;   // 同名項目（解析器之間可能取不同的那個）、路徑跳脫
    seen.add(e.name);
  }
  if (!REQUIRED.every((n) => seen.has(n)) || !entries.some((e) => SHEET_PART_RE.test(e.name))) return E;
  let total = 0;
  const zip = new JSZip();
  for (const e of entries) {
    if (e.name.endsWith('/')) continue;   // 資料夾項目
    const data = inflateEntry(buf, e);
    if (!data) return E;
    total += data.length;
    if (total > MAX_TOTAL_BYTES) return bad('BAD_XLSX', '檔案內容過大，無法匯入');
    if (/\.(?:xml|rels)$/i.test(e.name)) {
      const txt = data.toString('utf8');
      if (/<!DOCTYPE|<!ENTITY/i.test(txt)) return E;
      if (e.name === 'xl/workbook.xml') {
        const re = /<sheet\b[^>]*?\bname="([^"]*)"/g;
        let m;
        while ((m = re.exec(txt))) { if (/^(?:__proto__|constructor|prototype)$/i.test(m[1].trim())) return E; }
      }
    }
    zip.file(e.name, data);
  }
  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  return { ok: true, buf: out };
}

// ═════════════════ 儲存格解析 ═════════════════
const nfkc = (s) => String(s).normalize('NFKC');
const isBlankCell = (c) => !c || c.v === undefined || c.v === null || (typeof c.v === 'string' && c.v.trim() === '');
/** 這個儲存格是公式（Excel 裡打了 =…）：只會有快取值，不能當資料——匯入時整列報錯，要求改貼為值 */
const hasFormula = (c) => !!c && c.f !== undefined && c.f !== null;
const FORMULA_MSG = '儲存格含公式，請改貼為值';

/** 欄位標題比對：NFKC、去空白、去掉括號註記（全／半形）、小寫；別名完全相符或以別名開頭 */
function headerKey(s) {
  return nfkc(s).replace(/\s+/g, '').replace(/\([^)]*\)?/g, '').toLowerCase();
}
const ALIASES = {
  name: ['項目名稱', '項目', '名稱', '品名', '角色', '職稱', 'name', 'item'],
  price: ['牌價', '售價', '單價', '定價', 'price', 'listprice'],
  cost: ['成本', 'cost'],
  active: ['啟用', '狀態', 'active', 'enabled', 'enable'],
};
function matchHeader(s) {
  const k = headerKey(s);
  if (!k) return null;
  for (const f of ['name', 'price', 'cost', 'active']) {
    if (ALIASES[f].some((a) => k === a || (a.length >= 2 && k.startsWith(a)))) return f;
  }
  return null;
}

/** 金額儲存格 → {ok,value} | {ok:false,reason}。數字格或文字格；容許千分位、全形數字、前後空白、$／NT$／元 */
function parseMoneyCell(c, label) {
  if (isBlankCell(c)) return { ok: false, reason: `${label}空白` };
  if (c.t === 'e') return { ok: false, reason: `${label}儲存格是錯誤值` };
  let v;
  if (c.t === 'n') v = c.v;
  else if (c.t === 'b') return { ok: false, reason: `${label}不是數字` };
  else {
    let s = nfkc(c.v).trim().replace(/^(?:NT\$|NTD|\$)\s*/i, '').replace(/\s*元$/, '').trim();
    if (/^-/.test(s)) return { ok: false, reason: `${label}不可為負數` };
    if (s.includes(',')) {
      if (!/^\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.test(s)) return { ok: false, reason: `${label}格式不正確（千分位逗號位置有誤）` };
      s = s.replace(/,/g, '');
    }
    if (!/^\d+(?:\.\d+)?$/.test(s)) return { ok: false, reason: `${label}不是有效數字（「${String(c.v).slice(0, 20)}」）` };
    v = Number(s);
  }
  if (typeof v !== 'number' || !Number.isFinite(v)) return { ok: false, reason: `${label}不是有效數字` };
  if (v < 0) return { ok: false, reason: `${label}不可為負數` };
  if (v > PB.MAX_MONEY) return { ok: false, reason: `${label}超過上限 ${PB.MAX_MONEY}` };
  const m = PB.parseMoney(v);
  if (m === null) return { ok: false, reason: `${label}不是 0～${PB.MAX_MONEY} 的數字` };
  return { ok: true, value: m };
}

const TRUE_WORDS = new Set(['是', 'y', 'yes', 'true', '1', '啟用', '開', 'on', 't']);
const FALSE_WORDS = new Set(['否', 'n', 'no', 'false', '0', '停用', '關', 'off', 'f']);
/** 啟用欄 → {ok,value:true|false|undefined(空白=保持)} | {ok:false,reason} */
function parseActiveCell(c) {
  if (isBlankCell(c)) return { ok: true, value: undefined };
  if (c.t === 'b') return { ok: true, value: !!c.v };
  if (c.t === 'e') return { ok: false, reason: '啟用欄是錯誤值，請填「是」或「否」' };
  let key;
  if (c.t === 'n') key = String(c.v);
  else key = nfkc(c.v).trim().toLowerCase();
  if (TRUE_WORDS.has(key)) return { ok: true, value: true };
  if (FALSE_WORDS.has(key)) return { ok: true, value: false };
  return { ok: false, reason: `啟用欄請填「是」或「否」（目前是「${String(c.v).slice(0, 12)}」）` };
}

/** 名稱儲存格 → {ok,value} | {ok:false,reason}（Tab／換行收合成空白；其他控制字元、空白、過長拒絕） */
function parseNameCell(c) {
  if (isBlankCell(c)) return { ok: false, reason: '項目名稱空白' };
  if (c.t === 'e') return { ok: false, reason: '項目名稱儲存格是錯誤值' };
  let s = c.t === 'n' ? String(c.w !== undefined ? c.w : c.v) : String(c.v);
  s = s.replace(/[\t\r\n\u000b\u000c]+/g, ' ').trim();
  if (!s) return { ok: false, reason: '項目名稱空白' };
  if (/[\u0000-\u001f\u007f]/.test(s)) return { ok: false, reason: '項目名稱含有不允許的控制字元' };
  if (s.length > PB.MAX_NAME) return { ok: false, reason: `項目名稱超過 ${PB.MAX_NAME} 字` };
  return { ok: true, value: s };
}

// ═════════════════ 解析上傳檔（每個 BU 一個工作表）═════════════════
const MAX_SHEETS = 30;   // 活頁簿工作表數量上限（防呆；一般只有 5 個）

/**
 * 解析單一 BU 工作表 → {ok:true, rows:[{row,name?,key?,price?,cost?,active?,error?}], blankRows} | {ok:false,...}
 * rows[].row 是 Excel 的列號（1 起算）。名稱無效的列沒有 name／key。完全空白的工作表／只有標題列 → rows 為空（視為這個 BU 沒有資料）。
 */
function parseSheet(ws, bu) {
  if (!ws || !ws['!ref']) return { ok: true, rows: [], blankRows: 0, empty: true };
  const rng = XLSX.utils.decode_range(ws['!ref']);
  const lastRow = Math.min(rng.e.r, MAX_SCAN_ROWS - 1);
  const lastCol = Math.min(rng.e.c, MAX_COLS - 1);
  const cell = (r, c) => { const k = XLSX.utils.encode_cell({ r, c }); return Object.prototype.hasOwnProperty.call(ws, k) ? ws[k] : undefined; };

  // 標題列：前 10 列內第一個有「項目名稱」欄的列
  let hdrRow = -1, col = null, anyContent = false;
  for (let r = rng.s.r; r <= Math.min(lastRow, rng.s.r + MAX_HEADER_SCAN - 1); r++) {
    const cm = {};
    for (let c = rng.s.c; c <= lastCol; c++) {
      const x = cell(r, c);
      if (isBlankCell(x)) continue;
      anyContent = true;
      const f = matchHeader(String(x.v));
      if (f && cm[f] === undefined) cm[f] = c;
    }
    if (cm.name !== undefined) { hdrRow = r; col = cm; break; }
  }
  if (hdrRow < 0) {
    if (!anyContent) return { ok: true, rows: [], blankRows: 0, empty: true };
    return bad('BAD_HEADER', `工作表「${bu}」找不到欄位標題。第一列需要有：${HEADERS.slice(0, 3).join('、')}（可選：${HEADERS[3]}）。建議從本頁「下載範本」開始。`);
  }
  const miss = [];
  if (col.price === undefined) miss.push('牌價');
  if (col.cost === undefined) miss.push('成本');
  if (miss.length) return bad('BAD_HEADER', `工作表「${bu}」缺少欄位：${miss.join('、')}。需要有：${HEADERS.slice(0, 3).join('、')}（可選：${HEADERS[3]}）。`);

  const rows = [];
  let dataRows = 0, blankRows = 0, pendingBlank = 0;
  for (let r = hdrRow + 1; r <= lastRow; r++) {
    const cn = cell(r, col.name), cp = cell(r, col.price), cc = cell(r, col.cost), ca = col.active === undefined ? undefined : cell(r, col.active);
    const fRow = hasFormula(cn) || hasFormula(cp) || hasFormula(cc) || hasFormula(ca);
    if (!fRow && isBlankCell(cn) && isBlankCell(cp) && isBlankCell(cc) && isBlankCell(ca)) { pendingBlank++; continue; }
    blankRows += pendingBlank; pendingBlank = 0;   // 只算資料中間夾著的空白列（尾端的空白格式列不算）
    if (++dataRows > MAX_DATA_ROWS) return bad('TOO_MANY_ROWS', `工作表「${bu}」的資料列超過 ${MAX_DATA_ROWS} 列，請拆成多個檔案或移除多餘的列`);
    const o = { row: r + 1 };
    if (fRow) {   // 任何一格是公式 → 這一列整列不匯入（不更新、不新增；名稱若是純文字仍顯示，但不登記為「已出現」）
      const nf = parseNameCell(cn);
      if (nf.ok && !hasFormula(cn)) o.name = nf.value;
      o.error = FORMULA_MSG;
      rows.push(o);
      continue;
    }
    const nm = parseNameCell(cn);
    if (nm.ok) { o.name = nm.value; o.key = PB.nameKey(nm.value); }
    const pr = parseMoneyCell(cp, '牌價'), co = parseMoneyCell(cc, '成本'), ac = parseActiveCell(ca);
    const errs = [];
    if (!nm.ok) errs.push(nm.reason);
    if (!pr.ok) errs.push(pr.reason); else o.price = pr.value;
    if (!co.ok) errs.push(co.reason); else o.cost = co.value;
    if (!ac.ok) errs.push(ac.reason); else o.active = ac.value;
    if (errs.length) o.error = errs.join('；');
    rows.push(o);
  }
  return { ok: true, rows, blankRows };
}

/**
 * buf → {ok:true, sheets:{ERP?:{rows,blankRows}, …}, present:[有的 BU], missing:[檔案裡沒有的 BU], skippedSheets:[其他名稱的工作表]}
 *     | {ok:false,status,code,message}
 * 只認工作表名稱「完全等於」ERP／ITS／MDM／CRM 的；「說明」是本系統自己產生的說明頁，安靜略過；其他名稱列入 skippedSheets（顯示為略過）。
 */
async function parseWorkbook(buf) {
  try {
    if (!Buffer.isBuffer(buf) || buf.length === 0) return bad('EMPTY_FILE', '檔案是空的');
    const rp = await safeRepack(buf);
    if (!rp.ok) return rp;
    let wb;
    try {
      wb = XLSX.read(rp.buf, { type: 'buffer', cellFormula: true, cellHTML: false, cellStyles: false, cellNF: false, cellText: true, cellDates: false, bookVBA: false, sheetRows: MAX_SCAN_ROWS + 1 });
    } catch (_) { return bad('BAD_XLSX', '檔案不是有效的 Excel 活頁簿（.xlsx），請改用本頁的範本或匯出檔'); }
    const names = Array.isArray(wb && wb.SheetNames) ? wb.SheetNames : [];
    if (!names.length) return bad('BAD_XLSX', '活頁簿裡沒有工作表');
    if (names.length > MAX_SHEETS) return bad('BAD_XLSX', `活頁簿的工作表過多（超過 ${MAX_SHEETS} 個），請改用本頁的範本或匯出檔`);
    const sheets = {}, present = [], skippedSheets = [];
    for (const nm of names) {
      if (PB.BUS.includes(nm)) {
        if (Object.prototype.hasOwnProperty.call(sheets, nm)) continue;   // 理論上不會（SheetJS 不允許同名）
        const ws = Object.prototype.hasOwnProperty.call(wb.Sheets, nm) ? wb.Sheets[nm] : null;
        const r = parseSheet(ws, nm);
        if (!r.ok) return r;
        sheets[nm] = { rows: r.rows, blankRows: r.blankRows };
        present.push(nm);
      } else if (nm !== SHEET_HELP) skippedSheets.push(String(nm).slice(0, 40));
    }
    if (!present.length) return bad('NO_BU_SHEET', `找不到名稱為 ${PB.BUS.join('、')} 的工作表（工作表名稱必須完全相同）。請從本頁「下載範本」開始。`);
    const ordered = PB.BUS.filter((b) => present.includes(b));
    const missing = PB.BUS.filter((b) => !present.includes(b));
    if (ordered.every((b) => sheets[b].rows.length === 0)) return bad('NO_DATA', '各 BU 工作表都只有標題列，沒有任何資料列');
    return { ok: true, sheets, present: ordered, missing, skippedSheets };
  } catch (_) {
    return bad('BAD_XLSX', '檔案無法解析，請確認是由 Excel 存成的 .xlsx，或改用本頁的範本');
  }
}

// ═════════════════ 合併預覽（純函式，不寫入）═════════════════
const emptySum = () => ({ added: 0, updated: 0, unchanged: 0, skipped: 0, errors: 0 });

/**
 * existing：目前牌價簿的項目（含 id／bu／active；缺 bu 視為 ERP）。parsed：parseWorkbook 的結果（要用 .sheets／.skippedSheets／.missing）。
 * 回 { ok:true, summary:{added,updated,unchanged,skipped,errors, byBu:{ERP:{…},…}}, rows, skippedSheets, missingBus, merged } | { ok:false, ... }
 * summary.skipped＝資料中間夾著的空白列＋被略過的工作表數；byBu[bu].skipped 只算該 BU 的空白列。
 * merged：可直接 PUT 的完整清單，依 ERP、ITS、MDM、CRM 排列，每個 BU 內既有項目保留 id 與順序、新項目接在該 BU 最後（沒有 id，由 PUT 產生）。
 */
function buildPreview(existing, parsed, opts) {
  const o = opts || {};
  const maxItems = o.maxItems || PB.MAX_ITEMS;
  const cur = {};   // bu → 項目陣列（複本）
  for (const b of PB.BUS) cur[b] = [];
  for (const it of (Array.isArray(existing) ? existing : [])) {
    if (!it) continue;
    cur[PB.buOf(it)].push({ id: it.id, bu: PB.buOf(it), name: it.name, price: it.price, cost: it.cost, active: it.active !== false });
  }
  const sum = Object.assign(emptySum(), { byBu: {} });
  for (const b of PB.BUS) sum.byBu[b] = emptySum();
  const out = [];
  for (const bu of PB.BUS) {
    const sh = parsed.sheets && parsed.sheets[bu];
    if (!sh) continue;
    const list = cur[bu];
    const sb = sum.byBu[bu];
    const byKey = new Map();
    list.forEach((it, i) => { const k = PB.nameKey(it.name); if (!byKey.has(k)) byKey.set(k, i); });
    const seenRow = new Map();
    sb.skipped += sh.blankRows || 0;
    const errRow = (p, msg) => { sb.errors++; out.push({ bu, row: p.row, name: p.name || '', action: 'error', error: msg }); };
    for (const p of sh.rows) {
      if (p.key !== undefined) {
        if (seenRow.has(p.key)) { errRow(p, `與本工作表第 ${seenRow.get(p.key)} 列的項目名稱重複（不分大小寫、全半形），此列未匯入`); continue; }
        seenRow.set(p.key, p.row);
      }
      if (p.error) { errRow(p, p.error); continue; }
      const idx = byKey.get(p.key);
      if (idx !== undefined) {
        const c = list[idx];
        const nextActive = p.active === undefined ? c.active : p.active;
        const before = { price: c.price, cost: c.cost, active: c.active };
        const after = { price: p.price, cost: p.cost, active: nextActive };
        if (before.price === after.price && before.cost === after.cost && before.active === after.active) {
          sb.unchanged++;
          out.push({ bu, row: p.row, name: c.name, action: 'unchanged', old: before, new: after });
        } else {
          sb.updated++;
          list[idx] = Object.assign({}, c, after);
          out.push({ bu, row: p.row, name: c.name, action: 'update', old: before, new: after });
        }
      } else {
        if (list.length >= maxItems) { errRow(p, `${bu} 的牌價簿最多 ${maxItems} 項，已滿，此列未匯入`); continue; }
        const item = { bu, name: p.name, price: p.price, cost: p.cost, active: p.active !== false };
        list.push(item);
        byKey.set(p.key, list.length - 1);
        sb.added++;
        out.push({ bu, row: p.row, name: p.name, action: 'add', new: { price: item.price, cost: item.cost, active: item.active } });
      }
    }
  }
  for (const b of PB.BUS) for (const k of ['added', 'updated', 'unchanged', 'skipped', 'errors']) sum[k] += sum.byBu[b][k];
  const skippedSheets = Array.isArray(parsed.skippedSheets) ? parsed.skippedSheets : [];
  sum.skipped += skippedSheets.length;
  const merged = [];
  for (const b of PB.BUS) for (const it of cur[b]) merged.push(it);
  // 保險：合併結果必須能通過 PUT 的同一套驗證（理論上逐列檢查已涵蓋；這裡是最後一道）
  let n = 0;
  const chk = PB.normalizePricebook(merged, { genId: () => 'chk' + (++n) });
  if (!chk.ok) return bad('IMPORT_INVALID', '合併結果未通過驗證：' + chk.error.message);
  return { ok: true, summary: sum, rows: out, skippedSheets, missingBus: Array.isArray(parsed.missing) ? parsed.missing : [], merged };
}

/** 上傳檔 → 預覽。existingItems 是目前牌價簿（PB.adminItems 形式即可） */
async function previewFromBuffer(buf, existingItems, opts) {
  const p = await parseWorkbook(buf);
  if (!p.ok) return p;
  return buildPreview(existingItems, p, opts);
}

// ═════════════════ 產生 xlsx（匯出／範本）═════════════════
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const xmlEsc = (s) => String(s)
  .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\uFFFE\uFFFF]/g, '')   // XML 1.0 不允許的字元
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

// 樣式索引（對應 STYLES_XML 的 cellXfs 順序）
const S = { DEFAULT: 0, HEAD: 1, TEXT: 2, TEXT_QP: 3, INT: 4, DEC: 5, CENTER: 6, TITLE: 7, NOTE: 8, THEAD: 9, TCELL: 10, TEXT_C: 11 };

const STYLES_XML = XML_HEAD + `<styleSheet xmlns="${NS}">`
  + '<fonts count="4">'
  + '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
  + '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>'
  + '<font><b/><sz val="13"/><name val="Calibri"/><family val="2"/></font>'
  + '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
  + '</fonts>'
  + '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
  + '<fill><patternFill patternType="solid"><fgColor rgb="FF1A73E8"/><bgColor indexed="64"/></patternFill></fill>'
  + '<fill><patternFill patternType="solid"><fgColor rgb="FFF1F3F4"/><bgColor indexed="64"/></patternFill></fill></fills>'
  + '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
  + '<border><left style="thin"><color rgb="FFD1D5DB"/></left><right style="thin"><color rgb="FFD1D5DB"/></right><top style="thin"><color rgb="FFD1D5DB"/></top><bottom style="thin"><color rgb="FFD1D5DB"/></bottom><diagonal/></border></borders>'
  + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
  + '<cellXfs count="12">'
  + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
  + '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
  + '<xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
  + '<xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" quotePrefix="1"/>'
  + '<xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
  + '<xf numFmtId="4" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
  + '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
  + '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
  + '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
  + '<xf numFmtId="49" fontId="3" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>'
  + '<xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
  + '<xf numFmtId="49" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
  + '</cellXfs>'
  + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
  + '</styleSheet>';

const colName = (i) => String.fromCharCode(65 + i);
/** 明確的字串儲存格（inlineStr）；需要時加 quotePrefix 樣式 */
function strCell(ref, text, style) {
  const t = String(text);
  return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(t)}</t></is></c>`;
}
/** 使用者資料的文字：開頭是 = + - @ Tab CR 時用 quotePrefix 樣式（顯示原文，不當公式） */
function userTextCell(ref, text) {
  return strCell(ref, text, /^[=+\-@\t\r]/.test(String(text)) ? S.TEXT_QP : S.TEXT);
}
function moneyCell(ref, n) {
  return `<c r="${ref}" s="${Number.isInteger(n) ? S.INT : S.DEC}"><v>${String(n)}</v></c>`;
}

function dataSheetXml(items, selected) {
  const last = items.length + 1;
  const rows = [`<row r="1" ht="24" customHeight="1">${HEADERS.map((h, i) => strCell(colName(i) + '1', h, S.HEAD)).join('')}</row>`];
  items.forEach((it, i) => {
    const r = i + 2;
    rows.push(`<row r="${r}">${userTextCell('A' + r, it.name)}${moneyCell('B' + r, it.price)}${moneyCell('C' + r, it.cost)}${strCell('D' + r, it.active === false ? '否' : '是', S.TEXT_C)}</row>`);
  });
  // 預先格式化的空白列（到第 MAX_DATA_ROWS+1 列）：A、D 欄是文字格式（打「=A1」會保持文字，不會變成公式），B、C 欄是數字格式
  for (let r = last + 1; r <= MAX_DATA_ROWS + 1; r++) rows.push(`<row r="${r}"><c r="A${r}" s="${S.TEXT}"/><c r="B${r}" s="${S.INT}"/><c r="C${r}" s="${S.INT}"/><c r="D${r}" s="${S.TEXT_C}"/></row>`);
  const vEnd = MAX_DATA_ROWS + 1;
  return XML_HEAD + `<worksheet xmlns="${NS}">`
    + `<dimension ref="A1:D${last}"/>`
    + `<sheetViews><sheetView${selected ? ' tabSelected="1"' : ''} workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>`
    + '<sheetFormatPr defaultRowHeight="16.5"/>'
    + '<cols><col min="1" max="1" width="36" customWidth="1"/><col min="2" max="3" width="20" customWidth="1"/><col min="4" max="4" width="10" customWidth="1"/></cols>'
    + `<sheetData>${rows.join('')}</sheetData>`
    + '<dataValidations count="2">'
    + `<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="啟用" error="請選擇 是 或 否（留白＝維持原狀）" sqref="D2:D${vEnd}"><formula1>"是,否"</formula1></dataValidation>`
    + `<dataValidation type="decimal" operator="between" allowBlank="1" showErrorMessage="1" errorTitle="金額" error="請輸入 0 到 1,000,000,000 的數字" sqref="B2:C${vEnd}"><formula1>0</formula1><formula2>${PB.MAX_MONEY}</formula2></dataValidation>`
    + '</dataValidations>'
    + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    + '</worksheet>';
}

function helpLines(mode, info) {
  const L = [];
  const BU4 = PB.BUS.join('、');
  L.push(['title', '報價牌價簿：匯入／匯出說明']);
  if (mode === 'export') L.push(['note', `匯出日期：${info.date}　共 ${info.total} 項（啟用 ${info.active} 項；${PB.BUS.map((b) => b + ' ' + info.perBu[b]).join('、')}）`]);
  L.push(['note', '【機密】成本欄是公司內部資料，請妥善保管本檔，勿轉寄給客戶或外部人員。']);
  L.push(['blank', '']);
  L.push(['head', '檔案格式']);
  L.push(['note', `1. 牌價簿依事業單位（BU）分開維護：每個 BU 一個工作表，名稱必須「完全等於」${BU4}（不可改名、不可加空白）。本「${SHEET_HELP}」工作表只是說明，匯入時不會被讀取。`]);
  L.push(['note', '2. 每個 BU 工作表的欄位：項目名稱、牌價（元/人天）、成本（元/人天）、啟用（是／否）。單位固定是「人天」，不需要欄位。']);
  L.push(['note', `3. 項目名稱最長 ${PB.MAX_NAME} 字；金額 0～1,000,000,000，最多兩位小數（可含千分位逗號）；每個 BU 最多 ${PB.MAX_ITEMS} 項（共 ${PB.MAX_TOTAL} 項）；每個工作表一次最多 ${MAX_DATA_ROWS} 列；檔案 2 MB 以內，只接受 .xlsx。`]);
  L.push(['blank', '']);
  L.push(['head', '匯入規則（依「BU＋項目名稱」合併，不刪除）']);
  L.push(['note', '1. 在同一個 BU 內，依「項目名稱」比對（不分大小寫、全形半形與多餘空白視為相同）。不同 BU 可以有相同名稱，彼此獨立、費率可以不同。']);
  L.push(['note', '2. 同一 BU 內名稱相同：更新牌價與成本，保留原本的項目與順序（「啟用」欄有填才更新，留白＝維持原狀）。']);
  L.push(['note', '3. 該 BU 內名稱不存在：新增在該 BU 清單最後面（「啟用」留白＝啟用）。']);
  L.push(['note', '4. 檔案裡沒有出現的既有項目：完全不動，不會被刪除。要停用請把「啟用」改成「否」；要刪除請回後台畫面操作。']);
  L.push(['note', `5. 檔案裡缺少某個 BU 的工作表：該 BU 完全不動。名稱不是 ${BU4} 的其他工作表會被略過（預覽會列出）。`]);
  L.push(['note', '6. 同一個工作表內名稱重複：後面的列會被標示為錯誤，不會匯入。空白列會被略過。']);
  L.push(['note', '7. 按下匯入後會先顯示預覽（新增／更新／不變／略過／有問題的列，含各 BU 統計），確認之後才會寫入；有問題的列不會匯入，其餘照常。']);
  L.push(['note', '8. 匯入只影響牌價簿；已存的報價單不受影響（牌價簿只是建議值）。儲存格內容一律當純文字，不會被當成公式執行。']);
  if (mode === 'template') {
    L.push(['blank', '']);
    L.push(['head', '範例（虛構資料，僅供參考；不在 BU 工作表內，不會被匯入）']);
    L.push(['thead', HEADERS]);
    L.push(['trow', ['（範例）專案經理', '9,000', '6,500', '是']]);
    L.push(['trow', ['（範例）系統顧問', '7,000', '5,000', '是']]);
    L.push(['trow', ['（範例）暫停服務的角色', '6,000', '4,200', '否']]);
  }
  return L;
}

function helpSheetXml(mode, info) {
  const rows = [];
  helpLines(mode, info).forEach(([kind, v], i) => {
    const r = i + 1;
    if (kind === 'blank') return;
    if (kind === 'title') rows.push(`<row r="${r}" ht="22" customHeight="1">${strCell('A' + r, v, S.TITLE)}</row>`);
    else if (kind === 'head') rows.push(`<row r="${r}">${strCell('A' + r, v, S.TITLE)}</row>`);
    else if (kind === 'note') rows.push(`<row r="${r}">${strCell('A' + r, v, S.NOTE)}</row>`);
    else if (kind === 'thead') rows.push(`<row r="${r}">${v.map((x, c) => strCell(colName(c) + r, x, S.THEAD)).join('')}</row>`);
    else if (kind === 'trow') rows.push(`<row r="${r}">${v.map((x, c) => strCell(colName(c) + r, x, S.TCELL)).join('')}</row>`);
  });
  const n = helpLines(mode, info).length;
  return XML_HEAD + `<worksheet xmlns="${NS}">`
    + `<dimension ref="A1:D${n}"/>`
    + '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
    + '<sheetFormatPr defaultRowHeight="16.5"/>'
    + '<cols><col min="1" max="1" width="36" customWidth="1"/><col min="2" max="3" width="20" customWidth="1"/><col min="4" max="4" width="10" customWidth="1"/></cols>'
    + `<sheetData>${rows.join('')}</sheetData>`
    + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
    + '</worksheet>';
}

/**
 * opts.mode：'export'（資料列＝目前牌價簿）｜'template'（四個 BU 工作表都只有標題列，說明頁附虛構範例）
 * opts.items：export 用，{bu,name,price,cost,active}[]（缺 bu＝ERP；各 BU 內依陣列順序）；opts.date：'YYYY-MM-DD'（說明頁顯示）
 * 工作表順序：ERP、ITS、MDM、CRM、說明。
 */
async function buildWorkbook(opts) {
  const mode = opts && opts.mode === 'template' ? 'template' : 'export';
  const items = mode === 'export' ? (opts.items || []) : [];
  const groups = PB.groupByBu(items);
  const perBu = {};
  for (const b of PB.BUS) perBu[b] = groups[b].length;
  const info = { date: (opts && opts.date) || '', total: items.length, active: items.filter((x) => x.active !== false).length, perBu };
  const sheetNames = PB.BUS.concat([SHEET_HELP]);
  const zip = new JSZip();
  zip.file('[Content_Types].xml', XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + sheetNames.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')
    + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', XML_HEAD + `<workbook xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
    + '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="12000"/></bookViews>'
    + '<sheets>' + sheetNames.map((n, i) => `<sheet name="${xmlEsc(n)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + sheetNames.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')
    + `<Relationship Id="rId${sheetNames.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.file('xl/styles.xml', STYLES_XML);
  PB.BUS.forEach((b, i) => zip.file(`xl/worksheets/sheet${i + 1}.xml`, dataSheetXml(groups[b], i === 0)));
  zip.file(`xl/worksheets/sheet${PB.BUS.length + 1}.xml`, helpSheetXml(mode, info));
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

// ═════════════════ 檔名工具 ═════════════════
/** 稽核／訊息用的檔名：去路徑、去控制字元、限長（只當純文字，顯示端仍須轉義） */
function cleanFileName(s) {
  let n = String(s == null ? '' : s).split(/[\\/]/).pop().replace(/[\u0000-\u001f\u007f\u2028\u2029]/g, ' ').replace(/\s+/g, ' ').trim();
  if (n.length > 80) { const m = /(\.[A-Za-z0-9]{1,6})$/.exec(n); const ext = m ? m[1] : ''; n = n.slice(0, 80 - ext.length) + ext; }
  return n;
}

/** attachment 的 Content-Disposition：ASCII 後備檔名＋RFC 5987 的 UTF-8 檔名 */
function contentDisposition(utf8Name, asciiName) {
  const enc = encodeURIComponent(utf8Name).replace(/['()*]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase());
  return `attachment; filename="${String(asciiName).replace(/[^A-Za-z0-9._-]/g, '_')}"; filename*=UTF-8''${enc}`;
}

module.exports = {
  SHEET_HELP, HEADERS, MAX_FILE_BYTES, MAX_DATA_ROWS,
  buildWorkbook, parseWorkbook, parseSheet, buildPreview, previewFromBuffer,
  parseMoneyCell, parseActiveCell, parseNameCell, matchHeader, safeRepack,
  cleanFileName, contentDisposition,
};
