#!/usr/bin/env node
/* CSP 盤點工具：掃 _client 底下所有 HTML／JS／CSS，統計「要把 CSP 改成嚴格政策必須處理的東西」。
 * 用法：node scripts/csp-inventory.js [--json | --md]
 *   預設印簡表；--json 印完整 JSON；--md 印每頁 Markdown 摘要。
 * 只讀檔，不連網路、不連資料庫。可在每一期改完後重跑，數字應逐期下降（驗收用）。
 *
 * 分類：
 *  - 「靜態標記」＝HTML 檔裡 <script>/<style> 區塊以外的部分（頁面一載入就存在）。
 *  - 「JS 字串內」＝<script> 區塊或 .js 檔裡，被拼成 innerHTML 的 `onclick="..."`、`style="..."`
 *    （執行期才產生，同樣會被 script-src-attr / style-src-attr 擋；`el.onclick = fn`、`el.style.x = ...` 不會被擋，不計）。
 */
'use strict';
const fs = require('fs');
const path = require('path');

const CLIENT = path.join(__dirname, '..', '_client');

function lineOf(text, idx) { let n = 1; for (let i = 0; i < idx; i++) if (text.charCodeAt(i) === 10) n++; return n; }
function blank(s) { return s.replace(/[^\n]/g, ' '); }
function bump(map, k, n = 1) { map[k] = (map[k] || 0) + n; }
function top(map, n = 8) { return Object.entries(map).sort((a, b) => b[1] - a[1]).slice(0, n); }
function count(re, text) { const m = text.match(re); return m ? m.length : 0; }

const ON_ATTR_RE = /(?<![\w.$-])on([a-z]{3,20})\s*=\s*\\?["']/gi;
const KNOWN_EVENTS = new Set(('click dblclick change input submit keydown keyup keypress focus blur load error mouseover mouseout mouseenter mouseleave ' +
  'mousedown mouseup contextmenu dragstart dragover dragend drop dragenter dragleave paste copy cut scroll resize select touchstart touchend ' +
  'toggle wheel pointerdown pointerup pointermove invalid reset message').split(' '));
function handlerName(text, endIdx) {
  const rest = text.slice(endIdx, endIdx + 120).replace(/^\s*(?:event\.stopPropagation\(\);?\s*|return\s+|this\.)/, '');
  const m = /^\s*([A-Za-z_$][\w$.]*)\s*\(/.exec(rest);
  if (m) return m[1];
  return '(行內運算式)';
}
function scanOnAttrs(text) {
  const byEvent = {}, byHandler = {}; let total = 0; const lines = [];
  let m; ON_ATTR_RE.lastIndex = 0;
  while ((m = ON_ATTR_RE.exec(text))) {
    if (!KNOWN_EVENTS.has(m[1].toLowerCase())) continue;
    total++; bump(byEvent, 'on' + m[1].toLowerCase());
    bump(byHandler, handlerName(text, m.index + m[0].length));
    lines.push(lineOf(text, m.index));
  }
  return { total, byEvent, byHandler, lines };
}
function scanStyleAttrs(text) {
  let total = 0; const lines = [];
  const re = /(?<![\w.$-])style\s*=\s*\\?["']/g; let m;
  while ((m = re.exec(text))) { total++; lines.push(lineOf(text, m.index)); }
  return { total, lines };
}

const BLOCK_RE = /<(script|style)\b([^>]*)>([\s\S]*?)<\/\1\s*>/gi;
const NON_EXEC_TYPE = /json|template|text\/plain/i;

function scanHtml(file) {
  const raw = fs.readFileSync(file, 'utf8');
  const out = { file: path.basename(file), lines: raw.split('\n').length, inlineScripts: [], inlineStyles: [], extScripts: [], extStylesheets: [] };
  let staticMarkup = raw;       // script/style 區塊整段換成空白後的「靜態標記」（行號不變）
  let scriptOnly = blank(raw);  // 只留 inline 可執行 script 內文（行號不變）
  let m; BLOCK_RE.lastIndex = 0;
  while ((m = BLOCK_RE.exec(raw))) {
    const tag = m[1].toLowerCase(), attrs = m[2], body = m[3];
    const startLine = lineOf(raw, m.index), bodyLines = body.split('\n').length;
    staticMarkup = staticMarkup.slice(0, m.index) + blank(m[0]) + staticMarkup.slice(m.index + m[0].length);
    if (tag === 'style') { out.inlineStyles.push({ startLine, bodyLines, bytes: Buffer.byteLength(body) }); continue; }
    const src = /\bsrc\s*=\s*["']([^"']+)["']/i.exec(attrs);
    if (src) { out.extScripts.push({ src: src[1], line: startLine, sri: /\bintegrity\s*=/.test(attrs) }); continue; }
    const type = (/\btype\s*=\s*["']([^"']+)["']/i.exec(attrs) || [])[1] || 'text/javascript';
    const executes = !NON_EXEC_TYPE.test(type);
    out.inlineScripts.push({ startLine, endLine: startLine + bodyLines - 1, bodyLines, bytes: Buffer.byteLength(body), type, executes });
    if (executes) {
      const bodyStart = m.index + m[0].indexOf('>') + 1;
      scriptOnly = scriptOnly.slice(0, bodyStart) + raw.slice(bodyStart, bodyStart + body.length) + scriptOnly.slice(bodyStart + body.length);
    }
  }
  const linkRe = /<link\b[^>]*>/gi;
  while ((m = linkRe.exec(staticMarkup))) {
    if (/rel\s*=\s*["']stylesheet["']/i.test(m[0])) out.extStylesheets.push({ href: (/href\s*=\s*["']([^"']+)["']/i.exec(m[0]) || [])[1], line: lineOf(raw, m.index) });
  }
  out.staticOn = scanOnAttrs(staticMarkup);
  out.staticStyle = scanStyleAttrs(staticMarkup);
  out.javascriptUrls = count(/(?:href|src|action|formaction)\s*=\s*["']\s*javascript:/gi, staticMarkup);
  out.iframes = count(/<iframe\b/gi, staticMarkup);
  out.forms = (staticMarkup.match(/<form\b[^>]*>/gi) || []).map(s => (/action\s*=\s*["']([^"']*)["']/i.exec(s) || [])[1] || '(無 action)');
  out.imgExternal = count(/<img\b[^>]*\bsrc\s*=\s*["']https?:\/\/[^"']+/gi, staticMarkup);
  out.jsStringOn = scanOnAttrs(scriptOnly);
  out.jsStringStyle = scanStyleAttrs(scriptOnly);
  out.jsStringJavascriptUrl = count(/javascript:/g, scriptOnly);
  out.setAttrStyle = count(/setAttribute\(\s*['"]style['"]/g, scriptOnly);
  out.evalLike = count(/(?<![\w.$])eval\s*\(|new\s+Function\s*\(/g, scriptOnly);
  out.docWrite = count(/document\.write\s*\(/g, scriptOnly);
  return out;
}

function scanJs(file) {
  const raw = fs.readFileSync(file, 'utf8');
  return {
    file: path.basename(file), lines: raw.split('\n').length,
    on: scanOnAttrs(raw), style: scanStyleAttrs(raw),
    jsUrl: count(/javascript:/g, raw),
    setAttrStyle: count(/setAttribute\(\s*['"]style['"]/g, raw),
    evalLike: count(/(?<![\w.$])eval\s*\(|new\s+Function\s*\(/g, raw),
    docWrite: count(/document\.write\s*\(/g, raw),
    fetchExternal: raw.match(/fetch\(\s*['"`]https?:\/\/[^'"`]+/g) || [],
  };
}

function main() {
  const ls = fs.readdirSync(CLIENT).sort();
  const htmls = ls.filter(f => f.endsWith('.html')).map(f => scanHtml(path.join(CLIENT, f)));
  const jss = ls.filter(f => f.endsWith('.js')).map(f => scanJs(path.join(CLIENT, f)));
  const csss = ls.filter(f => f.endsWith('.css')).map(f => {
    const t = fs.readFileSync(path.join(CLIENT, f), 'utf8');
    return { file: f, importUrl: count(/@import\s+url\(/g, t), externalUrl: Array.from(new Set(t.match(/url\(\s*['"]?https?:\/\/[^)'"]+/g) || [])), fontFace: count(/@font-face/g, t), dataUrl: count(/url\(\s*['"]?data:/g, t) };
  });
  return { htmls, jss, csss };
}

function fmtTop(map, n) { return top(map, n).map(([k, v]) => `${k} ×${v}`).join('、') || '—'; }
function renderMd(r) {
  const L = [];
  for (const h of r.htmls) {
    const inl = h.inlineScripts.filter(s => s.executes);
    L.push(`### ${h.file}（${h.lines} 行）`);
    L.push(`- inline \`<script>\` 區塊：${inl.length} 個，內文共 ${inl.reduce((a, s) => a + s.bodyLines, 0)} 行 / ${inl.reduce((a, s) => a + s.bytes, 0)} bytes` + (inl.length ? '；位置（起—迄行，內文行數）：' + inl.map(s => `${s.startLine}—${s.endLine}（${s.bodyLines}）`).join('、') : ''));
    L.push(`- 外部 \`<script src>\`：${h.extScripts.length} 個；非同源：` + (h.extScripts.filter(s => /^https?:/.test(s.src)).map(s => `${s.src}（行 ${s.line}，${s.sri ? '有' : '無'} SRI）`).join('、') || '無'));
    L.push(`- 靜態標記的 \`on*\` 屬性：${h.staticOn.total} 個（${fmtTop(h.staticOn.byEvent, 6)}）；常見 handler：${fmtTop(h.staticOn.byHandler, 8)}`);
    L.push(`- inline script 內字串拼出的 \`on*=\`（執行期生成）：${h.jsStringOn.total} 個（${fmtTop(h.jsStringOn.byEvent, 6)}）；常見 handler：${fmtTop(h.jsStringOn.byHandler, 8)}`);
    L.push(`- 靜態標記 \`style=\`：${h.staticStyle.total}；inline script 字串內 \`style=\`：${h.jsStringStyle.total}；\`setAttribute('style')\`：${h.setAttrStyle}；inline \`<style>\` 區塊：${h.inlineStyles.length} 個（${h.inlineStyles.reduce((a, s) => a + s.bodyLines, 0)} 行）`);
    L.push(`- \`javascript:\` URL（靜態／字串）：${h.javascriptUrls}／${h.jsStringJavascriptUrl}；eval／new Function：${h.evalLike}；document.write：${h.docWrite}；iframe：${h.iframes}；form action：${h.forms.join('、') || '無'}`);
    L.push('');
  }
  return L.join('\n');
}

if (require.main === module) {
  const r = main();
  if (process.argv.includes('--json')) console.log(JSON.stringify(r, null, 2));
  else if (process.argv.includes('--md')) console.log(renderMd(r));
  else {
    for (const h of r.htmls) {
      const inl = h.inlineScripts.filter(s => s.executes);
      console.log(`${h.file.padEnd(12)} inlineScript=${inl.length}(${inl.reduce((a, s) => a + s.bodyLines, 0)}行) staticOn=${h.staticOn.total} jsStrOn=${h.jsStringOn.total} staticStyle=${h.staticStyle.total} jsStrStyle=${h.jsStringStyle.total} styleBlocks=${h.inlineStyles.length}`);
    }
    for (const j of r.jss) {
      if (j.on.total || j.style.total || j.setAttrStyle || j.evalLike || j.docWrite || j.fetchExternal.length || j.jsUrl) {
        console.log(`  ${j.file.padEnd(26)} on*=${j.on.total} style=${j.style.total} setAttr(style)=${j.setAttrStyle} eval=${j.evalLike} docWrite=${j.docWrite} jsUrl=${j.jsUrl} fetchExt=${j.fetchExternal.length}`);
      }
    }
  }
}
module.exports = { main, scanHtml, scanJs, renderMd };
