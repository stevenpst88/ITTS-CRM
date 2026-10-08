'use strict';
/**
 * lib/mail/templates.js — 信件「版面」：把 render.js 組好的 model 排成 HTML 與純文字（純函式，無 I/O）
 *
 * 分工：render.js 決定「寄什麼內容給誰、看得到什麼」（查 visibility 表、驗證、組 model）；
 *       本檔只負責「怎麼排版」，拿到的 model 已經是過濾後的資料，這裡不做任何角色／可見性判斷。
 *
 * 匯出與簽名：
 *   layoutHtml(model)      → string   完整 HTML 文件（email 安全寫法，見下）
 *   layoutText(model)      → string   純文字版（與 HTML 同資訊、同順序；LF 換行）
 *   wrapText(str, width, firstPrefix, restPrefix) → string[]   依顯示寬度折行（全形字算 2 欄）
 *   displayWidth(str)      → number   顯示欄寬（全形／CJK 算 2，其餘 1）
 *   PALETTE, TONES         色票（凍結）；測試會逐組驗證文字對比 >= 4.5
 *   FONT_STACK, MAX_TEXT_COLS
 *
 * model 形狀（所有字串都是「純文字」，本檔負責跳脫；url 由 render.js 驗證過）：
 *   {
 *     title, preheader, brand, confidential, headline, headlineTone,
 *     result:   null | { caption, label, tone, reasonLabel, reason },
 *     decision: null | { cells: [{ kind, caption, value, tag, warn, tone }] },   // tone: green|amber|red|grey|null；warn（可無）：獨立一行的警示，例如「毛利為負」
 *     greeting, lead: string[],
 *     rows:     [{ label, value }],
 *     items:    null | { head: [desc, qty, unit], rows: [[desc, qty, unit]], moreText },
 *     button:   { label, url },
 *     notes:    string[], footer: string[]
 *   }
 *
 * HTML 寫法（為了 Outlook 桌面版 Word 引擎、網頁版、手機、深色模式都能讀）：
 *  - 只用 table 版面＋inline CSS；沒有 flex／grid、沒有 JS、沒有外部資源或圖片；border-radius 一概不用
 *    （就算被忽略也不影響資訊）。重要色塊同時寫 bgcolor 屬性與 background-color，且色塊內一定有文字標籤。
 *  - 寬度：外層 width=600 屬性（Word 引擎認得），<=480px 時用 media query 改成 100% 並把決策條三格上下堆疊。
 *  - 按鈕：用 table＋td bgcolor 做（padding 放在 td，不放在 a），另附一行純文字網址當備援。
 *  - 深色模式：<meta color-scheme="light dark"> ＋ prefers-color-scheme 媒體查詢（含 !important 覆蓋 inline）
 *    ＋ Outlook.com 的 [data-ogsc]／[data-ogsb] 選擇器；不支援的用戶端仍是淺色底、深色字，可讀。
 *    色塊（綠／琥珀／紅／灰）本身就是深色底白字，淺色與深色模式下都維持原樣。
 *    色塊之間的分隔線用卡片底色（class="stack" 的格子，深色模式覆蓋成深色卡片底）、品項表外框用 class="bd"（深色模式覆蓋成深色邊線）：
 *    不要再寫死 #ffffff／#e5e7eb 的邊線，否則深色模式下會出現突兀的白線。
 *  - 隱藏的 preheader 預覽文字：class="preheader"（display:none＋mso-hide），測試與純文字一致性檢查會排除它。
 *  - <title> 放主旨；沒有任何地方把使用者輸入放進 href／屬性（href 只有 model.button.url）。
 *
 * 未驗證：真實 Outlook 桌面版（Word 引擎）、Outlook.com／Gmail／Apple Mail 的實際渲染——本機只用 Chromium 系
 * 瀏覽器看過（見 scratchpad/mail_render_report.md）。
 *
 * 原始碼備註：不寫四位數的 \uXXXX（寫檔工具會轉成真字元）；空白一律用 &nbsp; 實體或 ASCII 空白。
 */

const { escHtml } = require('./safety');

const FONT_STACK = "'Microsoft JhengHei','PingFang TC','Noto Sans TC','Heiti TC',Arial,sans-serif";
const MAX_TEXT_COLS = 76;

// ── 色票 ────────────────────────────────────────────────────────────────────
const PALETTE = Object.freeze({
  pageBg: '#f3f4f6',
  cardBg: '#ffffff',
  soft: '#f3f4f6',
  border: '#e5e7eb',
  text: '#111827',
  muted: '#4b5563',
  link: '#1d4ed8',
  btnBg: '#1d4ed8',
  btnText: '#ffffff',
  brandBg: '#1e3a8a',
  brandText: '#ffffff',
  brandAccent: '#fde68a',
  dark: Object.freeze({
    pageBg: '#0f172a',
    cardBg: '#1e293b',
    soft: '#273449',
    border: '#334155',
    text: '#f1f5f9',
    muted: '#a8b3c5',
    link: '#93c5fd',
  }),
});

// 色塊：深色底白字，淺／深色模式下都不需要換色。tone 名稱是 render.js 與這裡的共同詞彙。
const TONES = Object.freeze({
  green: Object.freeze({ bg: '#15803d', fg: '#ffffff' }),
  amber: Object.freeze({ bg: '#b45309', fg: '#ffffff' }),
  red: Object.freeze({ bg: '#b91c1c', fg: '#ffffff' }),
  grey: Object.freeze({ bg: '#4b5563', fg: '#ffffff' }),
  blue: Object.freeze({ bg: '#1d4ed8', fg: '#ffffff' }),
});

function toneOf(name) {
  return (typeof name === 'string' && Object.prototype.hasOwnProperty.call(TONES, name)) ? TONES[name] : TONES.grey;
}

// ── HTML 小工具 ─────────────────────────────────────────────────────────────
const esc = escHtml;

/** 屬性值跳脫：外層一律雙引號，所以只需處理 & < > "（單引號不跳脫，免得 font-family 的引號被改寫成實體） */
function escAttr(v) {
  return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function attrs(o) {
  let s = '';
  Object.keys(o).forEach((k) => {
    const v = o[k];
    if (v === null || v === undefined || v === false) return;
    s += ' ' + k + '="' + escAttr(v) + '"';
  });
  return s;
}
function el(name, a, inner) {
  return '<' + name + attrs(a || {}) + '>' + (inner === undefined ? '' : inner) + '</' + name + '>';
}
const TABLE_BASE = { role: 'presentation', cellpadding: '0', cellspacing: '0', border: '0' };
function table(a, inner) {
  return el('table', Object.assign({}, TABLE_BASE, a), inner);
}
function font(size, color, extra) {
  return 'font-family:' + FONT_STACK + ';font-size:' + size + 'px;color:' + color + ';' + (extra || '');
}

// ── CSS（放在 <style>；inline 樣式是淺色預設，這裡只負責手機與深色模式覆蓋）──────────
function styleBlock() {
  const D = PALETTE.dark;
  return [
    'body,table,td,a{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;}',
    'table,td{mso-table-lspace:0pt;mso-table-rspace:0pt;}',
    'table{border-collapse:collapse;}',
    '@media only screen and (max-width:480px){',
    '  .container{width:100% !important;max-width:100% !important;}',
    '  .px{padding-left:16px !important;padding-right:16px !important;}',
    '  .stack{display:block !important;width:100% !important;box-sizing:border-box;border-right:0 !important;border-bottom:2px solid ' + PALETTE.cardBg + ' !important;}',
    '  .amt{font-size:24px !important;line-height:30px !important;}',
    '  .lbl{width:32% !important;}',
    '  .btn-tbl{width:100% !important;}',
    '}',
    '@media (prefers-color-scheme:dark){',
    '  .bg-page{background-color:' + D.pageBg + ' !important;}',
    '  .bg-card{background-color:' + D.cardBg + ' !important;border-color:' + D.border + ' !important;}',
    '  .bg-soft{background-color:' + D.soft + ' !important;}',
    '  .tx{color:' + D.text + ' !important;}',
    '  .tx2{color:' + D.muted + ' !important;}',
    '  .bd{border-color:' + D.border + ' !important;}',
    // 決策條色塊之間的分隔線＝「卡片底色」的間隔，不是白線：深色模式要跟著換成深色卡片底（.stack 就是決策條的格子；
    // 手機堆疊時的 border-bottom 也是它，且這條規則在手機規則之後，同為 !important 時後者勝出）
    '  .stack{border-color:' + D.cardBg + ' !important;}',
    '  .lnk{color:' + D.link + ' !important;}',
    '}',
    '[data-ogsc] .tx{color:' + D.text + ' !important;}',
    '[data-ogsc] .tx2{color:' + D.muted + ' !important;}',
    '[data-ogsc] .lnk{color:' + D.link + ' !important;}',
    '[data-ogsb] .bg-page{background-color:' + D.pageBg + ' !important;}',
    '[data-ogsb] .bg-card{background-color:' + D.cardBg + ' !important;}',
    '[data-ogsb] .bg-soft{background-color:' + D.soft + ' !important;}',
  ].join('\n');
}

// ── 各區塊 ──────────────────────────────────────────────────────────────────
function brandRowHtml(m) {
  const P = PALETTE;
  const inner = table({ width: '100%' }, el('tr', {}, [
    el('td', { style: font(14, P.brandText, 'font-weight:bold;line-height:20px;') }, esc(m.brand)),
    el('td', { align: 'right', style: font(12, P.brandAccent, 'font-weight:bold;line-height:20px;') }, esc(m.confidential)),
  ].join('')));
  return el('tr', {}, el('td', { class: 'px', bgcolor: P.brandBg, style: 'background-color:' + P.brandBg + ';padding:10px 24px;' }, inner));
}

function headlineRowHtml(m) {
  const P = PALETTE;
  const tone = toneOf(m.headlineTone || 'blue');
  const inner = el('div', { class: 'tx', style: font(21, P.text, 'font-weight:bold;line-height:28px;') }, esc(m.headline));
  return el('tr', {}, el('td', { class: 'px', style: 'padding:20px 24px 4px 24px;' },
    table({ width: '100%' }, el('tr', {}, el('td', {
      style: 'border-left:5px solid ' + tone.bg + ';padding:2px 0 2px 12px;',
    }, inner)))));
}

function valueSize(value, base, widthPct) {
  const n = String(value).length;
  // 40% 以上寬的格子放得下 12 字的大字；三格時的窄格子（30%）只放得下 8 字——例如極端虧損單的「<-999999%」（9 字）
  // 在 26px 會折成「<-99999／9%」，所以窄格子超過 8 字就先縮 4px
  const fits = widthPct >= 40 ? 12 : 8;
  if (n <= fits) return base;
  if (n <= 12) return Math.max(base - 4, 16);
  if (n <= 16) return Math.max(base - 6, 16);
  return Math.max(base - 10, 14);
}

function decisionHtml(d) {
  const P = PALETTE;
  const cells = d.cells;
  const n = cells.length;
  const widths = n === 3 ? [40, 30, 30] : (n === 2 ? [50, 50] : [100]);
  const tds = cells.map((c, i) => {
    const toned = !!c.tone;
    const tone = toned ? toneOf(c.tone) : null;
    const bg = toned ? tone.bg : P.soft;
    const capColor = toned ? tone.fg : P.muted;
    const valColor = toned ? tone.fg : P.text;
    const vsize = valueSize(c.value, c.kind === 'tier' ? 22 : 26, widths[i]);
    const border = i < n - 1 ? 'border-right:2px solid ' + P.cardBg + ';' : '';
    const parts = [];
    parts.push(el('div', toned ? { style: font(12, capColor, 'line-height:16px;') } : { class: 'tx2', style: font(12, capColor, 'line-height:16px;') }, esc(c.caption)));
    parts.push(el('div', toned
      ? { class: 'amt', style: font(vsize, valColor, 'font-weight:bold;line-height:' + (vsize + 8) + 'px;word-break:break-all;') }
      : { class: 'tx amt', style: font(vsize, valColor, 'font-weight:bold;line-height:' + (vsize + 8) + 'px;word-break:break-all;') }, esc(c.value)));
    if (c.tag) {
      parts.push(el('div', toned
        ? { style: font(13, capColor, 'font-weight:bold;line-height:18px;') }
        : { class: 'tx2', style: font(12, capColor, 'line-height:18px;') }, esc(c.tag)));
    }
    if (c.warn) {
      parts.push(el('div', toned
        ? { style: font(13, capColor, 'font-weight:bold;line-height:18px;') }
        : { class: 'tx', style: font(13, valColor, 'font-weight:bold;line-height:18px;') }, esc(c.warn)));
    }
    return el('td', {
      class: toned ? 'stack' : 'stack bg-soft',
      width: widths[i] + '%',
      valign: 'top',
      bgcolor: bg,
      style: 'background-color:' + bg + ';padding:12px 14px;' + border,
    }, parts.join(''));
  }).join('');
  return table({ width: '100%' }, el('tr', {}, tds));
}

function resultHtml(r) {
  const P = PALETTE;
  const tone = toneOf(r.tone);
  const head = el('tr', {}, el('td', { bgcolor: tone.bg, style: 'background-color:' + tone.bg + ';padding:12px 16px;' }, [
    el('div', { style: font(12, tone.fg, 'line-height:16px;') }, esc(r.caption)),
    el('div', { class: 'amt', style: font(24, tone.fg, 'font-weight:bold;line-height:32px;') }, esc(r.label)),
  ].join('')));
  let body = '';
  if (r.reason) {
    body = el('tr', {}, el('td', { class: 'bg-soft', bgcolor: P.soft, style: 'background-color:' + P.soft + ';padding:10px 16px;' }, [
      el('div', { class: 'tx2', style: font(12, P.muted, 'line-height:16px;') }, esc(r.reasonLabel)),
      el('div', { class: 'tx', style: font(15, P.text, 'line-height:22px;word-break:break-word;') }, esc(r.reason)),
    ].join('')));
  }
  return table({ width: '100%' }, head + body);
}

function rowsHtml(rows) {
  const P = PALETTE;
  const trs = rows.map((r) => el('tr', {}, [
    el('td', { class: 'tx2 bd lbl', width: '28%', valign: 'top', style: 'padding:9px 12px 9px 0;border-bottom:1px solid ' + P.border + ';' + font(13, P.muted, 'line-height:20px;') }, esc(r.label)),
    el('td', { class: 'tx bd', valign: 'top', style: 'padding:9px 0;border-bottom:1px solid ' + P.border + ';' + font(15, P.text, 'line-height:22px;word-break:break-word;') }, esc(r.value)),
  ].join(''))).join('');
  return table({ width: '100%' }, trs);
}

function itemsHtml(it) {
  const P = PALETTE;
  const cellStyle = 'padding:7px 8px;border-bottom:1px solid ' + P.border + ';';
  const headRow = el('tr', {}, [
    el('td', { class: 'tx2 bg-soft bd', bgcolor: P.soft, width: '62%', style: cellStyle + 'background-color:' + P.soft + ';' + font(12, P.muted, 'font-weight:bold;line-height:16px;') }, esc(it.head[0])),
    el('td', { class: 'tx2 bg-soft bd', bgcolor: P.soft, width: '19%', align: 'right', style: cellStyle + 'background-color:' + P.soft + ';' + font(12, P.muted, 'font-weight:bold;line-height:16px;') }, esc(it.head[1])),
    el('td', { class: 'tx2 bg-soft bd', bgcolor: P.soft, width: '19%', style: cellStyle + 'background-color:' + P.soft + ';' + font(12, P.muted, 'font-weight:bold;line-height:16px;') }, esc(it.head[2])),
  ].join(''));
  const bodyRows = it.rows.map((r) => el('tr', {}, [
    el('td', { class: 'tx bd', valign: 'top', style: cellStyle + font(14, P.text, 'line-height:20px;word-break:break-word;') }, esc(r[0])),
    el('td', { class: 'tx bd', valign: 'top', align: 'right', style: cellStyle + font(14, P.text, 'line-height:20px;') }, esc(r[1])),
    el('td', { class: 'tx bd', valign: 'top', style: cellStyle + font(14, P.text, 'line-height:20px;') }, esc(r[2])),
  ].join(''))).join('');
  const more = it.moreText
    ? el('tr', {}, el('td', { class: 'tx2', colspan: '3', style: 'padding:7px 8px;' + font(13, P.muted, 'line-height:18px;') }, esc(it.moreText)))
    : '';
  return table({ class: 'bd', width: '100%', style: 'border:1px solid ' + P.border + ';' }, headRow + bodyRows + more);
}

function buttonHtml(b) {
  const P = PALETTE;
  const a = el('a', {
    href: b.url,
    target: '_blank',
    style: font(16, P.btnText, 'font-weight:bold;text-decoration:none;line-height:22px;display:inline-block;'),
  }, esc(b.label));
  const btn = table({ class: 'btn-tbl', align: 'left' }, el('tr', {}, el('td', {
    align: 'center',
    bgcolor: P.btnBg,
    style: 'background-color:' + P.btnBg + ';padding:13px 32px;',
  }, a)));
  const urlLine = el('div', { style: 'padding-top:10px;' }, el('a', {
    class: 'lnk',
    href: b.url,
    target: '_blank',
    style: font(12, P.link, 'line-height:18px;word-break:break-all;text-decoration:underline;'),
  }, esc(b.url)));
  return btn + '<div style="clear:both;line-height:0;font-size:0;">&nbsp;</div>' + urlLine;
}

/** 版面：回完整 HTML 文件 */
function layoutHtml(m) {
  const P = PALETTE;
  const o = [];
  o.push('<!DOCTYPE html>');
  o.push('<html lang="zh-Hant" xmlns:o="urn:schemas-microsoft-com:office:office">');
  o.push('<head>');
  o.push('<meta charset="utf-8">');
  o.push('<meta http-equiv="Content-Type" content="text/html; charset=utf-8">');
  o.push('<meta name="viewport" content="width=device-width, initial-scale=1">');
  o.push('<meta name="color-scheme" content="light dark">');
  o.push('<meta name="supported-color-schemes" content="light dark">');
  o.push('<meta name="x-apple-disable-message-reformatting">');
  o.push('<meta name="format-detection" content="telephone=no, date=no, address=no, email=no">');
  o.push('<title>' + esc(m.title) + '</title>');
  o.push('<!--[if mso]><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml><style type="text/css">body,table,td,div,p,a,span{font-family:\'Microsoft JhengHei\',Arial,sans-serif !important;}</style><![endif]-->');
  o.push('<style type="text/css">');
  o.push(styleBlock());
  o.push('</style>');
  o.push('</head>');
  o.push('<body class="bg-page" bgcolor="' + P.pageBg + '" style="margin:0;padding:0;background-color:' + P.pageBg + ';">');
  // 預覽文字：隱藏；後面塞一串零寬字元把信件正文擠出預覽列
  o.push('<div class="preheader" style="display:none;font-size:1px;line-height:1px;max-height:0;max-width:0;opacity:0;overflow:hidden;mso-hide:all;color:' + P.pageBg + ';">' + esc(m.preheader) + '&nbsp;&zwnj;'.repeat(40) + '</div>');

  const rows = [];
  rows.push(brandRowHtml(m));
  rows.push(headlineRowHtml(m));
  if (m.result) rows.push(el('tr', {}, el('td', { class: 'px', style: 'padding:14px 24px 2px 24px;' }, resultHtml(m.result))));
  if (m.decision) rows.push(el('tr', {}, el('td', { class: 'px', style: 'padding:14px 24px 2px 24px;' }, decisionHtml(m.decision))));
  rows.push(el('tr', {}, el('td', { class: 'px tx', style: 'padding:18px 24px 6px 24px;' + font(15, P.text, 'line-height:24px;') },
    el('div', { style: 'font-weight:bold;' }, esc(m.greeting)) + m.lead.map((p) => el('div', { style: 'padding-top:6px;' }, esc(p))).join(''))));
  rows.push(el('tr', {}, el('td', { class: 'px', style: 'padding:6px 24px 6px 24px;' }, rowsHtml(m.rows))));
  if (m.items) rows.push(el('tr', {}, el('td', { class: 'px', style: 'padding:12px 24px 6px 24px;' }, itemsHtml(m.items))));
  rows.push(el('tr', {}, el('td', { class: 'px', style: 'padding:18px 24px 8px 24px;' }, buttonHtml(m.button))));
  if (m.notes.length) {
    rows.push(el('tr', {}, el('td', { class: 'px tx2', style: 'padding:8px 24px 6px 24px;' + font(13, P.muted, 'line-height:20px;') },
      m.notes.map((n) => el('div', { style: 'padding-top:4px;' }, esc(n))).join(''))));
  }
  rows.push(el('tr', {}, el('td', { class: 'px bg-soft tx2 bd', bgcolor: P.soft, style: 'background-color:' + P.soft + ';padding:14px 24px;border-top:1px solid ' + P.border + ';' + font(12, P.muted, 'line-height:18px;') },
    m.footer.map((f) => el('div', {}, esc(f))).join(''))));

  o.push(table({ class: 'bg-page', width: '100%', bgcolor: P.pageBg, style: 'background-color:' + P.pageBg + ';' }, el('tr', {}, el('td', { align: 'center', style: 'padding:16px 8px;' },
    table({ class: 'container bg-card', width: '600', bgcolor: P.cardBg, style: 'width:600px;max-width:600px;background-color:' + P.cardBg + ';border:1px solid ' + P.border + ';' }, rows.join('\n'))))));
  o.push('</body>');
  o.push('</html>');
  return o.join('\n');
}

// ── 純文字 ──────────────────────────────────────────────────────────────────
function cpWidth(cp) {
  if (cp < 0x1100) return 1;
  if ((cp >= 0x1100 && cp <= 0x115f) || (cp >= 0x2e80 && cp <= 0xa4cf) || (cp >= 0xac00 && cp <= 0xd7a3) ||
      (cp >= 0xf900 && cp <= 0xfaff) || (cp >= 0xfe30 && cp <= 0xfe6f) || (cp >= 0xff00 && cp <= 0xff60) ||
      (cp >= 0xffe0 && cp <= 0xffe6) || (cp >= 0x20000 && cp <= 0x3fffd)) return 2;
  return 1;
}

/** 顯示欄寬：全形／CJK 算 2，其餘 1（逐 code point） */
function displayWidth(str) {
  let w = 0;
  for (const ch of String(str)) w += cpWidth(ch.codePointAt(0));
  return w;
}

// 不可出現在行首的收尾標點（禁則）：寧可讓上一行多 1 個字
const NO_LINE_START = new Set(Array.from('，。、；：！？）」』》】〉,.;:!?)]}%'));

/**
 * 依顯示欄寬折行。ASCII 連續字串（英數、網址片段）不從中間切，CJK 字元可逐字換行；
 * 超長的單一 ASCII 片段才會被硬切。第一行前綴 firstPrefix，續行前綴 restPrefix。回傳行陣列（不含換行字元）。
 */
function wrapText(str, width, firstPrefix, restPrefix) {
  const first = firstPrefix || '';
  const rest = restPrefix === undefined ? first.replace(/[^ ]/g, ' ') : restPrefix;
  const max = Math.max(20, width | 0);
  // 切成單位：空白、單一寬字元、連續的窄字元（單字）
  const units = [];
  let run = '';
  for (const ch of String(str)) {
    const cp = ch.codePointAt(0);
    if (ch === ' ') {
      if (run) { units.push(run); run = ''; }
      units.push(' ');
    } else if (cpWidth(cp) === 2) {
      if (run) { units.push(run); run = ''; }
      units.push(ch);
    } else {
      run += ch;
    }
  }
  if (run) units.push(run);

  const lines = [];
  let line = first;
  let w = displayWidth(first);
  let empty = true;
  const flush = () => { lines.push(line.replace(/ +$/, '')); line = rest; w = displayWidth(rest); empty = true; };
  units.forEach((u) => {
    if (u === ' ') {
      if (!empty && w + 1 <= max) { line += ' '; w += 1; }
      return;
    }
    let uw = displayWidth(u);
    if (w + uw > max && !empty) {
      const startsClosing = NO_LINE_START.has(Array.from(u)[0]);
      if (!(startsClosing && w + uw <= max + 2)) flush();
    }
    // 單一片段比一整行還長：硬切
    if (uw > max - displayWidth(rest)) {
      Array.from(u).forEach((ch) => {
        const cw = cpWidth(ch.codePointAt(0));
        if (w + cw > max && !empty) flush();
        line += ch; w += cw; empty = false;
      });
      return;
    }
    line += u; w += uw; empty = false;
  });
  if (!empty || lines.length === 0) lines.push(line.replace(/ +$/, ''));
  return lines;
}

/** 純文字版：與 HTML 同資訊、同順序。網址獨立一行不折行。 */
function layoutText(m) {
  const W = MAX_TEXT_COLS;
  const L = [];
  const sep = '='.repeat(56);
  const kv = (k, v) => wrapText(k + '：' + v, W, '', '  ').forEach((x) => L.push(x));
  const para = (s) => wrapText(s, W, '', '').forEach((x) => L.push(x));

  L.push(m.brand + '  ' + m.confidential);
  L.push(sep);
  para(m.headline);
  L.push(sep);
  if (m.result) {
    kv(m.result.caption, m.result.label);
    if (m.result.reason) kv(m.result.reasonLabel, m.result.reason);
    L.push('');
  }
  if (m.decision) {
    m.decision.cells.forEach((c) => {
      const note = [c.tag, c.warn].filter(Boolean).join('；');
      kv(c.caption, c.value + (note ? '（' + note + '）' : ''));
    });
    L.push('');
  }
  para(m.greeting);
  m.lead.forEach((p) => para(p));
  L.push('');
  m.rows.forEach((r) => kv(r.label, r.value));
  if (m.items) {
    L.push('');
    para(m.items.head.join('  '));
    m.items.rows.forEach((r) => wrapText(r.join('  '), W, '', '  ').forEach((x) => L.push(x)));
    if (m.items.moreText) para(m.items.moreText);
  }
  L.push('');
  L.push(m.button.label + '：');
  L.push(m.button.url);
  if (m.notes.length) {
    L.push('');
    m.notes.forEach((n) => para(n));
  }
  L.push('');
  L.push(sep);
  m.footer.forEach((f) => para(f));
  return L.join('\n') + '\n';
}

module.exports = {
  layoutHtml,
  layoutText,
  wrapText,
  displayWidth,
  PALETTE,
  TONES,
  FONT_STACK,
  MAX_TEXT_COLS,
};
