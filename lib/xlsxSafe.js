/**
 * 上傳活頁簿的共用安全閘門（SEC-05）。原本只有牌價簿匯入有的 safeRepack 抽到這裡，
 * 再套到所有 Excel／CSV 匯入口（server.js 的 uploadGuard 中介層，接在 multer 之後）。
 *
 * 原則：行為不變——正常檔案匯入結果必須與加閘門之前逐項相同。
 *   · .xlsx（zip，開頭 PK）：先自己讀中央目錄，逐項有界解壓（zlib maxOutputLength，不信任標頭宣告的大小；
 *     加密／ZIP64／同名項目／路徑跳脫／缺 [Content_Types].xml 或 xl/workbook.xml／無工作表 一律拒絕），拒絕含 DOCTYPE／ENTITY 的 XML
 *     與工作表名稱為 __proto__／constructor／prototype 的活頁簿，通過後用 JSZip 重新封裝成乾淨的 zip 再交給 SheetJS。
 *   · .xls（CFB，D0CF11E0）：原樣放行（SheetJS 0.20.3 解析）。
 *   · 其餘（.csv、HTML 偽裝成 .xls 的舊系統匯出等）：原樣放行，只拒絕「看起來是二進位垃圾」的檔案（含 NUL 位元組、
 *     又不是 UTF-16 BOM／舊版 BIFF 開頭）——這是改副檔名的執行檔／壓縮檔最常見的樣子。
 *   · 非 zip 的 zip 家族檔（docx／jar／ods…）缺少 xlsx 必要部件，會在 xlsx 路徑被拒絕。
 */
'use strict';
const zlib = require('zlib');
const JSZip = require('jszip');

const bad = (code, message, status, extra) => ({ ok: false, status: status || 400, code, message, extra });

const DEFAULTS = Object.freeze({
  maxEntryBytes: 6 * 1024 * 1024,     // 單一部件解壓後上限
  maxTotalBytes: 12 * 1024 * 1024,    // 解壓後總量上限
  maxEntries: 500,                    // zip 內項目數上限
  badMessage: '檔案不是有效的 Excel 活頁簿（.xlsx）',
  tooBigMessage: '檔案內容過大，無法匯入',
});

const REQUIRED = ['[Content_Types].xml', 'xl/workbook.xml'];
const BAD_NAME_RE = /(^|\/)\.\.(\/|$)|^\/|\\/;   // 路徑跳脫（..）、絕對路徑、反斜線
const SHEET_PART_RE = /^xl\/worksheets\/[^/]+\.xml$/;

/** 讀 zip 中央目錄（不解壓）→ 項目清單；ZIP64、結構不合理一律回 null */
function readCentralDirectory(buf, maxEntries) {
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
  if (total > maxEntries || cdOff + cdSize > eocd) return null;
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
function inflateEntry(buf, e, maxEntryBytes) {
  if (e.flags & 1) return null;                                   // 加密
  if (e.comp === 0xFFFFFFFF || e.uncomp === 0xFFFFFFFF) return null;
  if (e.uncomp > maxEntryBytes) return null;
  const lho = e.lho;
  if (lho + 30 > buf.length || buf.readUInt32LE(lho) !== 0x04034b50) return null;
  const start = lho + 30 + buf.readUInt16LE(lho + 26) + buf.readUInt16LE(lho + 28);
  const end = start + e.comp;
  if (end > buf.length) return null;
  const raw = buf.slice(start, end);
  let out;
  try {
    if (e.method === 0) out = raw;
    else if (e.method === 8) out = zlib.inflateRawSync(raw, { maxOutputLength: maxEntryBytes });
    else return null;
  } catch (_) { return null; }
  if (out.length !== e.uncomp || out.length > maxEntryBytes) return null;
  return out;
}

/** XML 部件 → 文字。支援 UTF-8（預設）與 UTF-16（有 BOM 或以 '<' 開頭的無 BOM 檔），否則編碼會讓後面的字串檢查掃不到 */
function xmlText(b) {
  if (b.length >= 2) {
    const swap = (x) => { const c = Buffer.from(x.slice(0, x.length & ~1)); c.swap16(); return c.toString('utf16le'); };
    if (b[0] === 0xFF && b[1] === 0xFE) return b.slice(2).toString('utf16le');
    if (b[0] === 0xFE && b[1] === 0xFF) return swap(b.slice(2));
    if (b.length >= 4 && b[0] === 0x3C && b[1] === 0) return b.toString('utf16le');
    if (b.length >= 4 && b[0] === 0 && b[1] === 0x3C) return swap(b);
  }
  return b.toString('utf8');
}
const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
/** 解 XML 字元參照與預設實體（重複解到穩定，防 &amp;#95; 這類疊加編碼） */
function decodeEntities(str) {
  let cur = str;
  for (let i = 0; i < 4; i++) {
    const nxt = cur.replace(/&#x([0-9a-f]+);|&#(\d+);|&(amp|lt|gt|quot|apos);/gi, (m, h, d, n) => {
      if (n) return ENT[n.toLowerCase()];
      const c = parseInt(h || d, h ? 16 : 10);
      return c >= 0 && c <= 0x10FFFF ? String.fromCodePoint(c) : '';
    });
    if (nxt === cur) break;
    cur = nxt;
  }
  return cur;
}
/** 從 [Content_Types].xml 移除指向「不保留的部件」的 <Override> */
function dropOverrides(txt, keepPart) {
  return txt.replace(/<(?:[\w.-]+:)?Override\b(?:"[^"]*"|'[^']*'|[^>"'])*>(?:\s*<\/(?:[\w.-]+:)?Override>)?/gi, (tag) => {
    const a = /\bPartName\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(tag);
    if (!a) return tag;
    const part = decodeEntities(a[1] !== undefined ? a[1] : a[2]).replace(/^\//, '');
    return keepPart(part) ? tag : '';
  });
}
const BAD_SHEET_NAME_RE = /^(?:__proto__|constructor|prototype)$/i;
/** workbook.xml 內任何 <sheet> 的 name 屬性（單／雙引號、實體、換行、命名空間前綴皆處理）是危險名稱 → true */
function hasBadSheetName(txt) {
  const tagRe = /<(?:[\w.-]+:)?sheet\b(?:"[^"]*"|'[^']*'|[^>"'])*>/gi;
  let m;
  while ((m = tagRe.exec(txt))) {
    const a = /\bname\s*=\s*(?:"([^"]*)"|'([^']*)')/i.exec(m[0]);
    if (a && BAD_SHEET_NAME_RE.test(decodeEntities(a[1] !== undefined ? a[1] : a[2]).trim())) return true;
  }
  return false;
}

/** 檢查＋有界解壓＋重新封裝成乾淨的 xlsx Buffer。回 {ok:true, buf} 或 {ok:false,...}。
 *  opts.keepPart(name)：有給時只解壓並保留符合的部件（其餘一律不解壓、不放進重新封裝的 zip——省 CPU，也縮小解析器攻擊面）；
 *  opts.compression：重新封裝用的壓縮（'STORE' 省 CPU）。 */
async function safeRepack(buf, opts) {
  const o = Object.assign({}, DEFAULTS, opts || {});
  const E = bad('BAD_XLSX', o.badMessage);
  if (!Buffer.isBuffer(buf) || buf.length < 4 || buf.readUInt32LE(0) !== 0x04034b50) return E;   // PK 魔術字
  const entries = readCentralDirectory(buf, o.maxEntries);
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
    if (o.keepPart && !o.keepPart(e.name)) continue;   // 白名單以外的部件：不解壓、不保留
    const data = inflateEntry(buf, e, o.maxEntryBytes);
    if (!data) return E;
    total += data.length;
    if (total > o.maxTotalBytes) return bad('BAD_XLSX', o.tooBigMessage);
    if (/\.(?:xml|rels)$/i.test(e.name)) {
      const txt = xmlText(data);
      if (/<!DOCTYPE|<!ENTITY/i.test(txt)) return E;
      if (e.name.toLowerCase() === 'xl/workbook.xml' && hasBadSheetName(txt)) return E;
    }
    // 丟掉部件後，[Content_Types].xml 裡指向它們的 Override 也要拿掉，否則 SheetJS 會因找不到部件（如 xl/metadata.xml）而丟錯
    zip.file(e.name, o.keepPart && e.name === '[Content_Types].xml' ? Buffer.from(dropOverrides(xmlText(data), o.keepPart), 'utf8') : data);
  }
  const out = await zip.generateAsync({ type: 'nodebuffer', compression: o.compression || 'DEFLATE' });
  return { ok: true, buf: out };
}

// ═════════════════ 一般匯入口（xlsx／xls／csv）═════════════════
const UPLOAD_LIMITS = Object.freeze({
  // 上傳上限 10MB。資料量大的合法檔（5 萬列×10 欄）工作表 XML 約 30MB；再大的檔 SheetJS 解析本身就會撞 Vercel 30 秒。
  // 上限越高，惡意的全零 zip（幾百 KB 解壓出上百 MB）越能放大 CPU，所以取「合法大檔仍過」的最小值。
  maxEntryBytes: 32 * 1024 * 1024,
  maxTotalBytes: 64 * 1024 * 1024,
  maxEntries: 2000,
  // 只保留試算表資料所需的部件；docProps、media、drawings、comments、vbaProject 等一律不解壓不保留（匯入只讀儲存格值）
  keepPart: (n) => /^(?:\[Content_Types\]\.xml|_rels\/\.rels|xl\/(?:workbook\.xml|_rels\/workbook\.xml\.rels|sharedStrings\.xml|styles\.xml|theme\/[^/]+\.xml|worksheets\/[^/]+\.xml))$/i.test(n),
  compression: 'STORE',
  badMessage: '檔案不是有效的 Excel／CSV 檔案，請確認格式（建議使用系統提供的範本）',
  tooBigMessage: '檔案內容過大，無法匯入',
});
const CFB_MAGIC = Buffer.from([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);

/** 非 zip、非 CFB 的內容：像文字（CSV／HTML）就放行；含 NUL 的二進位垃圾拒絕（UTF-16 BOM、舊版 BIFF 開頭例外） */
function looksLikeText(buf) {
  if (buf.length >= 2 && ((buf[0] === 0xFF && buf[1] === 0xFE) || (buf[0] === 0xFE && buf[1] === 0xFF))) return true;   // UTF-16 BOM
  if (buf.length >= 2 && buf[0] === 0x09 && [0x00, 0x02, 0x04, 0x08].includes(buf[1])) return true;                    // BIFF2–5 .xls
  const n = Math.min(buf.length, 4096);
  for (let i = 0; i < n; i++) if (buf[i] === 0) return false;
  return true;
}

/** 上傳內容 → {ok:true, buf, kind:'xlsx'|'xls'|'text'} | {ok:false,status,code,message} */
async function prepareUpload(buf) {
  if (!Buffer.isBuffer(buf)) return bad('BAD_UPLOAD', UPLOAD_LIMITS.badMessage);
  if (buf.length === 0) return bad('EMPTY_FILE', '檔案是空的');
  if (buf.length >= 2 && buf[0] === 0x50 && buf[1] === 0x4B) {            // 'PK'：zip 家族一律走 xlsx 嚴格路徑
    const rp = await safeRepack(buf, UPLOAD_LIMITS);
    return rp.ok ? { ok: true, buf: rp.buf, kind: 'xlsx' } : rp;
  }
  if (buf.length >= 8 && buf.slice(0, 8).equals(CFB_MAGIC)) return { ok: true, buf, kind: 'xls' };
  if (!looksLikeText(buf)) return bad('BAD_UPLOAD', UPLOAD_LIMITS.badMessage);
  return { ok: true, buf, kind: 'text' };
}

/**
 * Express 中介層：接在 multer（memoryStorage）之後。沒有檔案 → 放行（交給路由自己回「請上傳檔案」）；
 * 有檔案 → 驗證；xlsx 會把 req.file.buffer 換成重新封裝的乾淨版本，其餘原樣；不合格 → 400，不進路由。
 */
function uploadGuard() {
  return async function xlsxUploadGuard(req, res, next) {
    try {
      if (!req.file || !req.file.buffer) return next();
      const r = await prepareUpload(req.file.buffer);
      if (!r.ok) return res.status(r.status || 400).json({ error: r.message, code: r.code });
      if (r.buf !== req.file.buffer) { req.file.buffer = r.buf; req.file.size = r.buf.length; }
      return next();
    } catch (_) {
      return res.status(400).json({ error: UPLOAD_LIMITS.badMessage, code: 'BAD_UPLOAD' });
    }
  };
}

module.exports = { hasBadSheetName, decodeEntities, xmlText, safeRepack, prepareUpload, uploadGuard, readCentralDirectory, looksLikeText, UPLOAD_LIMITS, DEFAULTS };
