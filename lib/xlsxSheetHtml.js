/**
 * 把一個 xlsx 的工作表轉成 HTML 表格（給「毛利分析」網頁預覽用）。
 *
 * 為什麼不自己手刻一份版面：毛利分析是公司原本的 PNL 工作表，版面、公式、合併、框線都在範本裡；預覽若另外刻一份，
 * 範本一改就會兩邊不一致。這裡直接讀「實際要下載的那個 xlsx」（值已由 lib/quotePnlExcel.js 連快取值一起寫好），
 * 所以預覽＝下載檔案的樣子。
 *
 * 支援：列印範圍、隱藏列／欄、欄寬列高、合併儲存格（含部分隱藏的合併）、字型（粗體／斜體／底線／大小／顏色）、
 *       填色（rgb／theme＋tint／indexed）、框線（粗細／虛線／顏色）、對齊（水平／垂直／自動換行／縮排）、
 *       常見數字格式（千分位、小數、百分比、NT$、會計格式、[Red]、日期時間）。
 * 不支援（略過）：圖片／圖形（範本上的簽名線）／ActiveX（選項按鈕）、條件式格式、公式重算（直接用檔案裡的快取值）。
 * 安全：儲存格文字一律 esc()；來自 styles.xml 的值（顏色、對齊、框線…）都經白名單／格式驗證才會進 style 屬性。
 */
'use strict';

const JSZip = require('jszip');

const esc = (s) => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
const cp = (n) => { try { return String.fromCodePoint(n); } catch (e) { return '\uFFFD'; } };
const unesc = (s) => String(s || '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, d) => cp(+d)).replace(/&#x([0-9a-f]+);/gi, (m, h) => cp(parseInt(h, 16))).replace(/&amp;/g, '&');

const colIndex = (letters) => { let n = 0; for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64); return n; };
const colLetters = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const splitRef = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return m ? { col: colIndex(m[1]), row: +m[2] } : null; };
const attr = (tagText, name) => { const m = new RegExp('\\s' + name + '="([^"]*)"').exec(tagText); return m ? m[1] : undefined; };
/** 移除注音（<rPh>…</rPh>、<phoneticPr/>）：它們的 <t> 不是顯示文字 */
const stripPhonetic = (xml) => String(xml || '').replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').replace(/<phoneticPr\b[^>]*\/>/g, '');
const collectText = (xml) => unesc([...stripPhonetic(xml).matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map(t => t[1]).join(''));

// ── 顏色 ──────────────────────────────────────────────────────
const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333', '000000', 'FFFFFF'];
const HEX6 = /^[0-9A-Fa-f]{6}$/;
function hexToRgb(h) { return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
function rgbToHex(rgb) { return rgb.map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase(); }
function rgbToHsl([r, g, b]) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h = 0, s = 0; const l = (mx + mn) / 2;
  if (mx !== mn) { const d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
  return [h, s, l];
}
function hslToRgb([h, s, l]) {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  const f = (t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}
function applyTint(hex, tint) {
  if (!tint) return hex;
  const hsl = rgbToHsl(hexToRgb(hex));
  hsl[2] = tint < 0 ? hsl[2] * (1 + tint) : hsl[2] * (1 - tint) + tint;
  return rgbToHex(hslToRgb(hsl));
}
/** <color rgb|theme|indexed|auto ... tint> → '#RRGGBB'；auto／缺漏／格式不合 → null（輸出前一律驗證為 6 位十六進位） */
function resolveColor(tag, themeColors) {
  if (!tag) return null;
  const rgb = attr(tag, 'rgb'), theme = attr(tag, 'theme'), indexed = attr(tag, 'indexed'), tint = parseFloat(attr(tag, 'tint') || '0') || 0;
  let hex = null;
  if (rgb) hex = rgb.slice(-6);
  else if (theme !== undefined) hex = themeColors[+theme] || null;
  else if (indexed !== undefined) hex = INDEXED[+indexed] || null;
  if (!hex || !HEX6.test(hex)) return null;
  return '#' + applyTint(hex.toUpperCase(), Math.max(-1, Math.min(1, tint)));
}

// ── 數字格式 ──────────────────────────────────────────────────
const BUILTIN_FMT = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 5: '"$"#,##0_);("$"#,##0)', 6: '"$"#,##0_);[Red]("$"#,##0)', 7: '"$"#,##0.00_);("$"#,##0.00)', 8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 37: '#,##0_);(#,##0)', 38: '#,##0_);[Red](#,##0)', 39: '#,##0.00_);(#,##0.00)', 40: '#,##0.00_);[Red](#,##0.00)',
  41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)', 42: '_("$"* #,##0_);_("$"* (#,##0);_("$"* "-"_);_(@_)', 43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)', 44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy/m/d h:mm', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 49: '@' };
const DATE_IDS = new Set([14, 15, 16, 17, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
const stripFmt = (code) => code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
const isDateFmt = (code) => /[ymdhs]/i.test(stripFmt(code)) && !/[#0?]/.test(stripFmt(code).replace(/General/gi, ''));

const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DOW = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
function fmtDate(serial, code) {
  // 1900 年閏年錯誤：序號 61 以前（Excel 把 1900-02-29 當成存在）基準要晚一天
  const base = serial >= 61 ? Date.UTC(1899, 11, 30) : Date.UTC(1899, 11, 31);
  const d = new Date(base + Math.round(serial * 86400) * 1000);
  const Y = d.getUTCFullYear(), M = d.getUTCMonth() + 1, D = d.getUTCDate(), wd = d.getUTCDay(), h = d.getUTCHours(), mi = d.getUTCMinutes(), se = d.getUTCSeconds();
  const p2 = (n) => String(n).padStart(2, '0');
  const ampm = /AM\/PM|A\/P/i.test(code);
  const toks = [...code.matchAll(/"[^"]*"|\[[^\]]*\]|\\.|AM\/PM|A\/P|yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|./gi)].map(m => m[0]);
  const lower = toks.map(t => t.toLowerCase());
  return toks.map((t, i) => {
    const k = lower[i];
    if (t[0] === '"') return t.slice(1, -1);
    if (t[0] === '[') return /^\[h+\]$/i.test(t) ? String(Math.floor(serial * 24)) : '';
    if (t[0] === '\\') return t[1];
    // m／mm 在「小時之後」或「秒之前」是分鐘，其餘是月份
    const prevH = (() => { for (let j = i - 1; j >= 0; j--) { if (/^(h|hh|\[h+\])$/.test(lower[j])) return true; if (/^(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|s|ss)$/.test(lower[j])) return false; } return false; })();
    const nextS = (() => { for (let j = i + 1; j < toks.length; j++) { if (/^(s|ss)$/.test(lower[j])) return true; if (/^(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|h|hh)$/.test(lower[j])) return false; } return false; })();
    const hh = ampm ? (h % 12 === 0 ? 12 : h % 12) : h;
    switch (k) {
      case 'yyyy': return String(Y);
      case 'yy': return p2(Y % 100);
      case 'mmmm': return MON[M - 1];
      case 'mmm': return MON[M - 1].slice(0, 3);
      case 'mm': return (prevH || nextS) ? p2(mi) : p2(M);
      case 'm': return (prevH || nextS) ? String(mi) : String(M);
      case 'dddd': return DOW[wd];
      case 'ddd': return DOW[wd].slice(0, 3);
      case 'dd': return p2(D);
      case 'd': return String(D);
      case 'hh': return p2(hh);
      case 'h': return String(hh);
      case 'ss': return p2(se);
      case 's': return String(se);
      case 'am/pm': return h < 12 ? 'AM' : 'PM';
      case 'a/p': return h < 12 ? 'A' : 'P';
      default: return t;
    }
  }).join('');
}

/** 以十進位字串位移做四捨五入（half away from zero），避免 0.02055*100 這類浮點誤差 */
function shift10(v, n) {
  const s = String(Number(v.toPrecision(15)));
  if (/e/i.test(s)) return v * Math.pow(10, n);
  return Number(s + 'e' + n);
}
function roundHalfUp(v, dec) { return shift10(Math.round(shift10(v, dec)), -dec); }

/** 依 Excel 格式碼把數字轉成顯示文字與顏色（[Red] 等）。涵蓋這份報價／毛利檔用到的格式；其餘退回一般格式 */
function formatNumberEx(val, code) {
  if (!Number.isFinite(val)) return { text: String(val), color: null };
  if (!code || /^general$/i.test(code)) return { text: Number.isInteger(val) ? String(val) : String(Math.round(val * 1e10) / 1e10), color: null };
  if (isDateFmt(code)) return { text: fmtDate(val, code.split(';')[0]), color: null };
  const sections = code.split(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/);
  let sec = sections[0], v = val, neg = false;
  if (val < 0 && sections.length > 1) { sec = sections[1]; v = -val; }
  else if (val === 0 && sections.length > 2) { sec = sections[2]; }
  else if (val < 0) { v = -val; neg = true; }
  const color = /\[Red\]/i.test(sec) ? '#FF0000' : null;
  sec = sec.replace(/\[[^\]]*\]/g, '').replace(/_./g, ' ').replace(/\*./g, '');
  const parts = []; let i = 0, pct = 0;
  while (i < sec.length) {
    const ch = sec[i];
    if (ch === '"') { const j = sec.indexOf('"', i + 1); parts.push({ lit: sec.slice(i + 1, j < 0 ? sec.length : j) }); i = j < 0 ? sec.length : j + 1; }
    else if (ch === '\\') { parts.push({ lit: sec[i + 1] || '' }); i += 2; }
    else if (/[#0?,.]/.test(ch)) { let j = i; while (j < sec.length && /[#0?,.]/.test(sec[j])) j++; parts.push({ num: sec.slice(i, j) }); i = j; }
    else if (ch === '%') { pct++; parts.push({ lit: '%' }); i++; }
    else if (ch === '@') { parts.push({ lit: String(val) }); i++; }
    else { parts.push({ lit: ch }); i++; }
  }
  if (pct) v = shift10(v, 2 * pct);
  const out = []; let placed = false;
  for (const p of parts) {
    if (p.lit !== undefined) { out.push(p.lit); continue; }
    if (placed) continue;
    placed = true;
    const pat = p.num, dot = pat.indexOf('.');
    const intPat = dot < 0 ? pat : pat.slice(0, dot), decPat = dot < 0 ? '' : pat.slice(dot + 1);
    const dec = (decPat.match(/[0#?]/g) || []).length;
    const useComma = /,/.test(intPat);
    const rounded = roundHalfUp(v, dec);
    let [ip, dp = ''] = rounded.toFixed(dec).split('.');
    if (ip === '0' && !/0/.test(intPat.replace(/,/g, ''))) ip = '';
    else if (useComma) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    let dpOut = dp;
    for (let k = decPat.length - 1; k >= 0 && dpOut.length; k--) { if (decPat[k] !== '0' && dpOut.endsWith('0') && dpOut.length > (decPat.slice(0, k).match(/0/g) || []).length) dpOut = dpOut.slice(0, -1); else break; }
    out.push(ip + (dpOut ? '.' + dpOut : ''));
  }
  let text = out.join('').replace(/^-\s+/, '-');
  if (neg) text = '-' + text;
  return { text, color };
}
const formatNumber = (val, code) => formatNumberEx(val, code).text;

// ── styles.xml ────────────────────────────────────────────────
const H_OK = new Set(['left', 'center', 'right', 'justify', 'centerContinuous', 'fill', 'distributed', 'general']);
function parseStyles(xml, themeColors) {
  const grab = (tag) => { const m = new RegExp('<' + tag + '(?:\\s[^>]*)?>([\\s\\S]*?)</' + tag + '>').exec(xml); return m ? m[1] : ''; };
  const numFmts = {};
  for (const m of xml.matchAll(/<numFmt\s[^>]*?numFmtId="(\d+)"[^>]*?formatCode="([^"]*)"/g)) numFmts[+m[1]] = unesc(m[2]);
  const fonts = [...grab('fonts').matchAll(/<font\b[^>]*?(?:\/>|>([\s\S]*?)<\/font>)/g)].map(m => {
    const b = m[1] || '';
    const sz = /<sz val="([\d.]+)"/.exec(b), col = /<color\b[^>]*\/>/.exec(b);
    const size = sz ? parseFloat(sz[1]) : 11;
    return { b: /<b\b/.test(b) && !/<b val="0"/.test(b), i: /<i\b/.test(b) && !/<i val="0"/.test(b), u: /<u\b/.test(b) && !/<u val="none"/.test(b), strike: /<strike\b/.test(b), sz: size > 0 && size < 100 ? size : 11, color: col ? resolveColor(col[0], themeColors) : null };
  });
  const fills = [...grab('fills').matchAll(/<fill>([\s\S]*?)<\/fill>/g)].map(m => {
    const pf = /<patternFill\b([^>]*)>([\s\S]*?)<\/patternFill>|<patternFill\b([^>]*)\/>/.exec(m[1]);
    if (!pf) return null;
    const type = attr(' ' + (pf[1] || pf[3] || ''), 'patternType');
    if (type !== 'solid') return null;
    const fg = /<fgColor\b[^>]*\/>/.exec(pf[2] || '');
    return fg ? resolveColor(fg[0], themeColors) : null;
  });
  const borders = [...grab('borders').matchAll(/<border\b[^>]*?(?:\/>|>([\s\S]*?)<\/border>)/g)].map(m => {
    const b = m[1] || '', o = {};
    for (const side of ['left', 'right', 'top', 'bottom']) {
      const sm = new RegExp('<' + side + '\\b([^>]*?)(?:/>|>([\\s\\S]*?)</' + side + '>)').exec(b);
      if (!sm) continue;
      const style = attr(' ' + sm[1], 'style'); if (!style) continue;
      const col = /<color\b[^>]*\/>/.exec(sm[2] || '');
      o[side] = { style, color: (col && resolveColor(col[0], themeColors)) || '#000000' };
    }
    return o;
  });
  const xfs = [...grab('cellXfs').matchAll(/<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g)].map(m => {
    const a = ' ' + m[1], al = /<alignment\b[^>]*\/>/.exec(m[2] || '');
    const h = al ? attr(al[0], 'horizontal') : undefined, v = al ? attr(al[0], 'vertical') : undefined;
    return { numFmtId: +(attr(a, 'numFmtId') || 0) || 0, fontId: +(attr(a, 'fontId') || 0) || 0, fillId: +(attr(a, 'fillId') || 0) || 0, borderId: +(attr(a, 'borderId') || 0) || 0,
      h: H_OK.has(h) ? h : undefined, v: ['top', 'center', 'bottom'].includes(v) ? v : undefined, wrap: al ? attr(al[0], 'wrapText') === '1' : false, indent: Math.max(0, Math.min(15, al ? (+(attr(al[0], 'indent') || 0) || 0) : 0)) };
  });
  return { numFmts, fonts, fills, borders, xfs };
}

function parseTheme(xml) {
  const names = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const def = ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646', '0000FF', '800080'];
  return names.map((n, i) => {
    const m = new RegExp('<a:' + n + '>([\\s\\S]*?)</a:' + n + '>').exec(xml || '');
    if (!m) return def[i];
    const rgb = /srgbClr val="([0-9A-Fa-f]{6})"/.exec(m[1]) || /lastClr="([0-9A-Fa-f]{6})"/.exec(m[1]);
    return rgb ? rgb[1].toUpperCase() : def[i];
  });
}

function parseSst(xml) {
  const out = [];
  for (const m of (xml || '').matchAll(/<si>([\s\S]*?)<\/si>|<si\s*\/>/g)) out.push(m[1] === undefined ? '' : collectText(m[1]));
  return out;
}

const BORDER_CSS = { thin: '1px solid', hair: '1px solid', medium: '2px solid', thick: '3px solid', dashed: '1px dashed', mediumDashed: '2px dashed', dotted: '1px dotted', dashDot: '1px dashed', mediumDashDot: '2px dashed', dashDotDot: '1px dotted', slantDashDot: '2px dashed', double: '3px double' };

/** 列印範圍（可能有多段、含工作表名稱）→ 第一段純儲存格範圍 'A1:H106'；無法解析（整欄整列、具名範圍）回 null */
function firstRange(printAreaText) {
  const raw = unesc(printAreaText || '');
  // 以不在引號內的逗號切段，取第一段，再去掉工作表名稱（可能含空白、括號、逗號、驚嘆號）
  const pieces = []; let cur = '', q = false;
  for (const ch of raw) { if (ch === "'") q = !q; if (ch === ',' && !q) { pieces.push(cur); cur = ''; } else cur += ch; }
  pieces.push(cur);
  let first = pieces[0] || '';
  const bang = first.lastIndexOf('!');
  if (bang >= 0) first = first.slice(bang + 1);
  first = first.replace(/\$/g, '');
  return /^[A-Z]+\d+(:[A-Z]+\d+)?$/.test(first) ? first : null;
}

/**
 * @param {Buffer} buffer xlsx
 * @param {{sheetPath?:string, range?:string, maxCols?:number, maxRows?:number}} [opts]
 * @returns {Promise<{html:string, widthPx:number, heightPx:number, range:string}>}
 */
async function sheetToHtml(buffer, opts = {}) {
  const zip = await JSZip.loadAsync(buffer);
  const read = async (p) => { const f = zip.file(p); return f ? f.async('string') : ''; };
  const sheetXml = await read(opts.sheetPath || 'xl/worksheets/sheet1.xml');
  if (!sheetXml) throw new Error('找不到工作表');
  const themeColors = parseTheme(await read('xl/theme/theme1.xml'));
  const S = parseStyles(await read('xl/styles.xml'), themeColors);
  const sst = parseSst(await read('xl/sharedStrings.xml'));
  const wbXml = await read('xl/workbook.xml');

  // 顯示範圍：優先用列印範圍（第一段），解析不了就用 dimension
  let range = opts.range;
  if (!range) { const pa = /<definedName[^>]*Print_Area[^>]*>([^<]+)</.exec(wbXml); if (pa) range = firstRange(pa[1]); }
  if (!range) range = (/<dimension ref="([^"]+)"/.exec(sheetXml) || [])[1] || 'A1:A1';
  const rm = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(range);
  if (!rm) throw new Error('範圍格式不正確：' + range);
  const c1 = colIndex(rm[1]), r1 = +rm[2], c2 = rm[3] ? colIndex(rm[3]) : c1, r2 = rm[4] ? +rm[4] : r1;
  if (c2 - c1 > (opts.maxCols || 40) || r2 - r1 > (opts.maxRows || 400)) throw new Error('範圍過大，不產生預覽');

  const fmt0 = /<sheetFormatPr\b[^>]*>/.exec(sheetXml);
  const defCol = parseFloat(attr(fmt0 ? fmt0[0] : '', 'defaultColWidth') || '8.43') || 8.43;
  const defRowPt = parseFloat(attr(fmt0 ? fmt0[0] : '', 'defaultRowHeight') || '15') || 15;
  const colInfo = {};
  for (const m of sheetXml.matchAll(/<col\b[^>]*\/>/g)) {
    const mn = +attr(m[0], 'min'), mx = +attr(m[0], 'max');
    for (let c = Math.max(mn, c1); c <= Math.min(mx, c2); c++) colInfo[c] = { w: parseFloat(attr(m[0], 'width') || defCol) || defCol, hidden: attr(m[0], 'hidden') === '1' };
  }
  // <col width> 已含 5px 內距（ECMA-376 18.3.1.13），所以像素 ≈ 字元數 × 7
  const colW = (c) => { const ci = colInfo[c] || { w: defCol, hidden: false }; return ci.hidden ? 0 : Math.max(1, Math.round(ci.w * 7)); };

  // 列與儲存格
  const rows = {}; const cells = {};
  for (const rm2 of sheetXml.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const a = ' ' + rm2[1], rn = +attr(a, 'r');
    if (!(rn >= r1 && rn <= r2)) continue;
    rows[rn] = { ht: attr(a, 'ht') !== undefined ? (parseFloat(attr(a, 'ht')) || defRowPt) : defRowPt, hidden: attr(a, 'hidden') === '1', defStyle: attr(a, 's') !== undefined && attr(a, 'customFormat') === '1' ? +attr(a, 's') : undefined };
    for (const cm of (rm2[2] || '').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const ca = ' ' + cm[1], ref = splitRef(attr(ca, 'r') || ''); if (!ref) continue;
      const inner = cm[2] || '', t = attr(ca, 't'), s = +(attr(ca, 's') || 0) || 0;
      let val = null, kind = 'empty';
      const v = /<v>([\s\S]*?)<\/v>/.exec(inner);
      if (t === 's' && v) { val = sst[+v[1]] ?? ''; kind = 'text'; }
      else if (t === 'inlineStr') { val = collectText(inner); kind = 'text'; }
      else if (t === 'str' && v) { val = unesc(v[1]); kind = 'text'; }
      else if (t === 'b' && v) { val = v[1] === '1' ? 'TRUE' : 'FALSE'; kind = 'bool'; }
      else if (t === 'e' && v) { val = unesc(v[1]); kind = 'error'; }
      else if (v && v[1] !== '') { const n = parseFloat(v[1]); if (Number.isFinite(n)) { val = n; kind = 'num'; } }
      cells[ref.row + ',' + ref.col] = { s, val, kind };
    }
  }

  const visibleCols = []; for (let c = c1; c <= c2; c++) if (colW(c) > 0 && !(colInfo[c] && colInfo[c].hidden)) visibleCols.push(c);
  const visibleRows = []; for (let r = r1; r <= r2; r++) if (!(rows[r] && rows[r].hidden)) visibleRows.push(r);

  // 合併：先裁到顯示範圍，再裁到「可見」的列欄；以第一個可見儲存格當錨點（隱藏列／欄、列印範圍裁切都不會讓後面的格子位移）
  const mergeAt = new Map(); const covered = new Set();
  for (const m of sheetXml.matchAll(/<mergeCell\s+ref="([^"]+)"/g)) {
    const mm = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(m[1]); if (!mm) continue;
    const g = { c1: Math.max(colIndex(mm[1]), c1), r1: Math.max(+mm[2], r1), c2: Math.min(colIndex(mm[3]), c2), r2: Math.min(+mm[4], r2) };
    if (g.c1 > g.c2 || g.r1 > g.r2) continue;
    const vr = visibleRows.filter(r => r >= g.r1 && r <= g.r2), vc = visibleCols.filter(c => c >= g.c1 && c <= g.c2);
    if (!vr.length || !vc.length || (vr.length === 1 && vc.length === 1)) continue;
    const info = { ar: vr[0], ac: vc[0], lastR: vr[vr.length - 1], lastC: vc[vc.length - 1], rowspan: vr.length, colspan: vc.length };
    if (mergeAt.has(info.ar + ',' + info.ac)) continue;   // 重疊的合併：只認第一個
    let overlap = false;
    for (const rr of vr) for (const cc of vc) if (covered.has(rr + ',' + cc) || (mergeAt.has(rr + ',' + cc))) overlap = true;
    if (overlap) continue;
    mergeAt.set(info.ar + ',' + info.ac, info);
    for (const rr of vr) for (const cc of vc) if (!(rr === info.ar && cc === info.ac)) covered.add(rr + ',' + cc);
  }

  const totalW = visibleCols.reduce((s, c) => s + colW(c), 0);
  const fmtOf = (xf) => S.numFmts[xf.numFmtId] || BUILTIN_FMT[xf.numFmtId] || (DATE_IDS.has(xf.numFmtId) ? 'yyyy/m/d' : 'General');
  const sideCss = (side, bd) => { const b = bd && bd[side]; return b ? `border-${side}:${BORDER_CSS[b.style] || '1px solid'} ${b.color};` : `border-${side}:0;`; };
  const hasContent = (r, c) => { const cell = cells[r + ',' + c]; return !!cell && cell.kind !== 'empty' && String(cell.val) !== ''; };
  const nextVisibleCol = (c) => { const i = visibleCols.indexOf(c); return i >= 0 && i + 1 < visibleCols.length ? visibleCols[i + 1] : null; };
  const bdOf = (rr, cc) => { const cx = cells[rr + ',' + cc]; const x = S.xfs[(cx && cx.s) || (rows[rr] && rows[rr].defStyle) || 0]; return (x && S.borders[x.borderId]) || {}; };
  let totalH = 0;
  const body = [];
  for (const r of visibleRows) {
    const ri = rows[r] || { ht: defRowPt, hidden: false };
    const hPx = Math.max(1, Math.round(ri.ht * 4 / 3)); totalH += hPx;
    const tds = [];
    for (const c of visibleCols) {
      const key = r + ',' + c;
      if (covered.has(key)) continue;
      const cell = cells[key] || { s: ri.defStyle || 0, val: null, kind: 'empty' };
      const xf = S.xfs[cell.s] || S.xfs[0] || { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 };
      const font = S.fonts[xf.fontId] || { sz: 11 };
      const g = mergeAt.get(key);
      const own = S.borders[xf.borderId] || {};
      // 合併儲存格的框線：左上取自錨點，右取自右上角、下取自左下角
      const merged = g ? { left: own.left, top: own.top, right: bdOf(g.ar, g.lastC).right, bottom: bdOf(g.lastR, g.ac).bottom } : own;
      const fill = S.fills[xf.fillId];
      let text = '', numColor = null;
      if (cell.kind === 'num') { const f = formatNumberEx(cell.val, fmtOf(xf)); text = f.text; numColor = f.color; }
      else if (cell.kind !== 'empty') text = String(cell.val);
      const h = xf.h && xf.h !== 'general' ? xf.h : (cell.kind === 'num' ? 'right' : (cell.kind === 'bool' || cell.kind === 'error' ? 'center' : 'left'));
      const va = xf.v === 'top' ? 'top' : xf.v === 'center' ? 'middle' : 'bottom';
      // 不換行的文字：右邊相鄰格有內容時裁切（Excel 的行為），沒有才讓它溢出到空白格；連續空白保留（pre）
      const lastC = g ? g.lastC : c;
      const nc = nextVisibleCol(lastC);
      const clip = h !== 'right' && nc !== null && hasContent(r, nc);
      const css = [
        'box-sizing:border-box', 'padding:0 3px', `height:${hPx}px`, `vertical-align:${va}`, `text-align:${h === 'centerContinuous' ? 'center' : (['left', 'center', 'right', 'justify'].includes(h) ? h : 'left')}`,
        `font-size:${Math.round(font.sz * 4 / 3 * 10) / 10}px`, font.b ? 'font-weight:700' : '', font.i ? 'font-style:italic' : '',
        (font.u || font.strike) ? `text-decoration:${[font.u ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ')}` : '',
        `color:${numColor || font.color || '#000'}`, fill ? `background:${fill}` : '',
        sideCss('left', merged), sideCss('right', merged), sideCss('top', merged), sideCss('bottom', merged),
        xf.wrap ? 'white-space:pre-wrap;word-break:break-word;overflow:hidden' : `white-space:pre;overflow:${clip ? 'hidden' : 'visible'}`,
        xf.indent ? `padding-left:${3 + xf.indent * 9}px` : '', 'line-height:1.25',
      ].filter(Boolean).join(';');
      tds.push(`<td${g && g.colspan > 1 ? ` colspan="${g.colspan}"` : ''}${g && g.rowspan > 1 ? ` rowspan="${g.rowspan}"` : ''} style="${css}">${esc(text)}</td>`);
    }
    body.push(`<tr style="height:${hPx}px">${tds.join('')}</tr>`);
  }
  const colgroup = '<colgroup>' + visibleCols.map(c => `<col style="width:${colW(c)}px">`).join('') + '</colgroup>';
  const html = `<table class="xsh-table" style="border-collapse:collapse;table-layout:fixed;width:${totalW}px;background:#fff;color:#000;font-family:'Microsoft JhengHei','微軟正黑體','PingFang TC','Noto Sans TC',Calibri,Arial,sans-serif">${colgroup}<tbody>${body.join('')}</tbody></table>`;
  return { html, widthPx: totalW, heightPx: totalH, range };
}

module.exports = { sheetToHtml, _internal: { formatNumber, formatNumberEx, resolveColor, applyTint, parseTheme, parseStyles, colIndex, colLetters, firstRange, fmtDate, roundHalfUp } };
