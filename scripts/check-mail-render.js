#!/usr/bin/env node
'use strict';
/**
 * 信件樣板與渲染檢查。用法：node scripts/check-mail-render.js（不開伺服器、不連網、不碰 data.json／auth.json）
 *
 * 涵蓋（§1.6、§2、§3）：
 *   1 金額與時間格式   formatNtd（0、1 分、99 分、1e12 分、安全整數上限、負數／小數／字串拒絕、與 BigInt 對拍、源碼無浮點除法）
 *   2 事件×收件人矩陣  6 事件 × 7 kind：19 種可渲染、23 種 throw KIND_NOT_ALLOWED；可見性逐格（顧問不含任何金額／百分比／毛利／折扣，
 *                     秘書不含客戶與業務名，業務不含金額，簽核人含金額與毛利；主旨絕不含金額／毛利率／客戶名／業務名）
 *   3 決策條與色塊     numbers 為 null 整段不輸出；色塊對照 tierLevel；色塊一定有文字標籤；marginText／負毛利
 *   4 主旨             格式、截斷、換行／標頭注入
 *   5 內容細節         品項（30 列上限、無價格）、結果與原因、時間、稱呼、連結、preheader
 *   6 XSS 與注入       14 種惡意字串 × 多個欄位：標籤集合與屬性集合不變、href 恆為預期網址、無 on* 屬性
 *   7 HTML 良構性      自製 tokenizer 檢查配對／引號屬性／實體／註解；無外部資源、無 script、無 expression()／url()
 *   8 純文字一致性     HTML 可見文字 ＝ 純文字（去除裝飾字元與空白後逐字相等），含 XSS 案例；行寬、網址獨立一行
 *   9 錯誤行為與穩健性 只 throw MailRenderError；getter 改值（TOCTOU）；凍結輸入；亂數亂輸入 3000 組；大小上限；效能
 *  10 版面與色票       深色模式／手機樣式存在、色票文字對比 >= 4.5、折行函式
 *  11 預覽工具         scripts/mail-preview.js 煙霧測試（輸出到 .mail-preview 之內的暫存資料夾，結束後清掉）
 *
 * 環境變數：MAIL_CORE_ROOT 要檢查的專案根目錄（預設＝本檔上一層）；變異測試（scripts/check-mail-mutation.js）會指向被破壞的副本。
 * 撰寫備註：本檔不寫四位數的 \uXXXX，一律用 cp()／\u{...}／\x..。
 */
const fs = require('fs');
const path = require('path');
const util = require('util');
const assert = require('assert');
const { spawnSync } = require('child_process');

const ROOT = process.env.MAIL_CORE_ROOT || path.join(__dirname, '..');
const load = (rel) => require(path.join(ROOT, rel));

// ── 迷你測試框架 ─────────────────────────────────────────────────────────────
const results = [];
const perf = [];
let currentSection = '';
function record(name, ok, extra) { results.push({ section: currentSection, name, ok: !!ok, extra: extra === undefined ? '' : String(extra) }); }
const t = (name, ok, extra) => record(name, ok, extra);
function short(v) {
  let s;
  try { s = typeof v === 'string' ? JSON.stringify(v) : util.inspect(v, { depth: 4, breakLength: Infinity }); } catch (e) { s = String(v); }
  return s.length > 200 ? s.slice(0, 200) + '…' : s;
}
function eq(name, actual, expected) {
  let ok = true;
  try { assert.deepStrictEqual(actual, expected); } catch (e) { ok = false; }
  record(name, ok, ok ? '' : 'actual=' + short(actual) + ' expected=' + short(expected));
}
function section(name, fn) {
  currentSection = name;
  try { fn(); } catch (e) { record('章節執行中例外', false, (e && e.stack) || e); }
}
function finish() {
  const failed = results.filter((r) => !r.ok);
  const bySec = {};
  results.forEach((r) => { const s = bySec[r.section] || (bySec[r.section] = { p: 0, f: 0 }); if (r.ok) s.p++; else s.f++; });
  Object.keys(bySec).forEach((k) => console.log((bySec[k].f ? 'FAIL ' : 'ok   ') + k + '  通過 ' + bySec[k].p + (bySec[k].f ? '，失敗 ' + bySec[k].f : '')));
  const maxShow = 60;
  failed.slice(0, maxShow).forEach((r) => console.log('  ✗ [' + r.section + '] ' + r.name + (r.extra ? '  → ' + r.extra : '')));
  if (failed.length > maxShow) console.log('  …另有 ' + (failed.length - maxShow) + ' 項失敗未列出');
  if (perf.length) console.log('效能：' + perf.join('；'));
  console.log((failed.length ? 'FAILED' : 'PASSED') + '：' + (results.length - failed.length) + ' / ' + results.length);
  process.exit(failed.length ? 1 : 0);
}
function throwsErr(name, fn, Cls, code) {
  try { fn(); record(name, false, 'did not throw'); } catch (e) {
    record(name, e instanceof Cls && (code === undefined || e.code === code), 'type=' + (e && e.constructor && e.constructor.name) + ' code=' + (e && e.code) + ' msg=' + (e && e.message));
  }
}
const cp = (...n) => String.fromCodePoint(...n);

// ── 載入被測模組 ─────────────────────────────────────────────────────────────
const R = load('lib/mail/render.js');
const T = load('lib/mail/templates.js');
const { getMailConfig } = load('lib/mail/config.js');
const { visibilityFor, KINDS } = load('lib/mail/visibility.js');
const { EVENT_TYPES } = load('lib/mail/events.js');
const { renderMail, MailRenderError } = R;

const BASE = 'https://crm.example.test';
const CTX = { config: getMailConfig({ APP_BASE_URL: BASE }) };
const QID = '0b9f6c1e-aaaa-4bbb-8ccc-ddddeeeeffff';
const AT = '2026-10-08T06:30:00.000Z';
const CAN = {
  company: 'ZetaCorp客戶甲', owner: '業務Kappa乙', project: 'ProjectAlpha機房案', viewer: '收件人Lambda',
  actor: '操作人Sigma', tierLabel: '核決Omega', step: '關卡Theta', reason: '原因Rho說明',
};
const LEAKS_AMOUNT = ['1,234,567', '123456789', '41.04', CAN.tierLabel, 'NT$', '折扣後未稅', '毛利', '報價金額', '核決層級'];
const ITEM_PRICE_CANARIES = ['987654', '777777', '555555'];

const NUM = () => ({ revenueCents: 123456789, gpCents: 50671000, marginText: '41.04%', marginPct: 41.04, tierLevel: 1, tierLabel: CAN.tierLabel });
const ITEMS = () => [
  { desc: '品項Phi一', qty: 2, unit: '台', price: 987654, unitPrice: 777777, cost: 555555, discount: 0.1 },
  { desc: '品項Phi二', qty: 10, unit: '人天', price: 987654, unitPrice: 777777, cost: 555555 },
];
function mkEv(type, kind, over) {
  const ev = {
    type, quoteId: QID, quoteNo: 'QU-TEST-001', projectName: CAN.project, company: CAN.company, ownerLabel: CAN.owner,
    at: AT, stepKey: AT + '#1', numbers: NUM(), items: ITEMS(), actor: { label: CAN.actor },
  };
  if (type === 'E1_SUBMIT' || type === 'E3_NEXT_STEP' || type === 'E4_RESULT' || type === 'E6_WITHDRAWN') {
    ev.step = { level: kind === 'secretary' || kind === 'boardProxy' ? 'board' : 1, label: CAN.step };
  }
  if (type === 'E4_RESULT') ev.result = { kind: 'rejected', reason: CAN.reason };
  if (type === 'E6_WITHDRAWN') ev.result = { kind: 'withdrawn', reason: CAN.reason };
  return Object.assign(ev, over || {});
}
const mkViewer = (kind, label) => ({ username: 'user-' + kind, label: label === undefined ? CAN.viewer : label, kind });
const render = (type, kind, over, viewerLabel, ctx) => renderMail(mkEv(type, kind, over), mkViewer(kind, viewerLabel), ctx || CTX);
const urlFor = (type, id) => BASE + '/q/' + encodeURIComponent(id || QID) + (type === 'E2_COST_REQUEST' ? '?cost=1' : '');

// ── 獨立抄錄的期望表（不從被測模組匯入，才能抓到被改壞的政策）──────────────────────────
const EXPECT_ALLOWED = {
  E1_SUBMIT: ['mgr1', 'gm', 'chairman', 'secretary', 'boardProxy'],
  E2_COST_REQUEST: ['consultant'],
  E3_NEXT_STEP: ['mgr1', 'gm', 'chairman', 'secretary', 'boardProxy'],
  E4_RESULT: ['owner'],
  E5_COST_DONE: ['owner'],
  E6_WITHDRAWN: ['mgr1', 'gm', 'chairman', 'secretary', 'boardProxy', 'owner'],
};
const EXPECT_VIS = {
  mgr1: { amount: 1, margin: 1, tier: 1, customer: 1, owner: 1, items: 0, project: 1 },
  gm: { amount: 1, margin: 1, tier: 1, customer: 1, owner: 1, items: 0, project: 1 },
  chairman: { amount: 1, margin: 1, tier: 1, customer: 1, owner: 1, items: 0, project: 1 },
  secretary: { amount: 1, margin: 1, tier: 1, customer: 0, owner: 0, items: 0, project: 0 },     // 業主 2026-10-08 決定：秘書與董事會代核人的信只放單號，不放專案名稱
  boardProxy: { amount: 1, margin: 1, tier: 1, customer: 0, owner: 0, items: 0, project: 0 },
  consultant: { amount: 0, margin: 0, tier: 0, customer: 0, owner: 1, items: 1, project: 1 },
  owner: { amount: 0, margin: 0, tier: 0, customer: 1, owner: 0, items: 0, project: 1 },
};
const SUBJECT_PREFIX = { E1_SUBMIT: '【簽核通知】', E2_COST_REQUEST: '【請填寫成本】', E3_NEXT_STEP: '【簽核通知】', E4_RESULT: '【簽核結果】', E5_COST_DONE: '【成本已填寫】', E6_WITHDRAWN: '【簽核撤回】' };
const DECISION_TYPES = ['E1_SUBMIT', 'E3_NEXT_STEP'];
const ACTOR_ALWAYS = ['E3_NEXT_STEP', 'E4_RESULT', 'E5_COST_DONE'];

// ── HTML 工具（測試端自己的解析器，不依賴被測模組）─────────────────────────────────
const VOID_TAGS = new Set(['meta', 'br', 'hr', 'img', 'input', 'link']);
function decodeEntities(s) {
  return s.replace(/&nbsp;/g, cp(0xa0)).replace(/&zwnj;/g, cp(0x200c)).replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&#96;/g, '`').replace(/&amp;/g, '&');
}
/** 解析 HTML；errors 非空代表不良構。 */
function parseHtml(html) {
  const errors = [];
  const tags = [];
  const comments = [];
  let styleText = '';
  let s = html;
  if (s.indexOf('<!DOCTYPE html>\n') !== 0) errors.push('缺少 DOCTYPE');
  s = s.replace('<!DOCTYPE html>', '');
  s = s.replace(/<!--([\s\S]*?)-->/g, (m, c) => { comments.push(c); return ''; });
  if (s.indexOf('<!--') >= 0 || s.indexOf('-->') >= 0) errors.push('註解未結束');
  s = s.replace(/<style type="text\/css">([\s\S]*?)<\/style>/g, (m, css) => { styleText += css; return '<style type="text/css"></style>'; });
  const re = /<(\/?)([a-zA-Z][a-zA-Z0-9]*)((?:\s+[a-zA-Z_:][-a-zA-Z0-9_:.]*(?:="[^"]*")?)*)\s*>/y;
  const stack = [];
  let i = 0;
  while (i < s.length) {
    const lt = s.indexOf('<', i);
    const text = s.slice(i, lt < 0 ? s.length : lt);
    if (text.indexOf('>') >= 0) errors.push('文字內有未跳脫的 >');
    if (/&(?!(?:amp|lt|gt|quot|nbsp|zwnj|#[0-9]{1,6});)/.test(text)) errors.push('文字內有不合法的 & 實體');
    if (lt < 0) break;
    re.lastIndex = lt;
    const m = re.exec(s);
    if (!m) { errors.push('多餘的 < ：' + s.slice(lt, lt + 40)); i = lt + 1; continue; }
    const closing = m[1] === '/';
    const name = m[2].toLowerCase();
    const attrs = {};
    const names = [];
    m[3].replace(/([a-zA-Z_:][-a-zA-Z0-9_:.]*)(?:="([^"]*)")?/g, (mm, k, v) => {
      if (names.indexOf(k) >= 0) errors.push('重複屬性 ' + k);
      names.push(k);
      attrs[k] = v === undefined ? '' : v;
      if (v !== undefined && (v.indexOf('<') >= 0 || v.indexOf('>') >= 0)) errors.push('屬性值內有 < 或 >');
      if (v !== undefined && /&(?!(?:amp|lt|gt|quot|nbsp|zwnj|#[0-9]{1,6});)/.test(v)) errors.push('屬性值內有不合法的 & 實體');
      return mm;
    });
    tags.push({ name, attrs, closing });
    if (!closing && !VOID_TAGS.has(name)) stack.push(name);
    if (closing) {
      if (VOID_TAGS.has(name)) errors.push('void 標籤不該有結尾 ' + name);
      else if (stack.pop() !== name) errors.push('標籤配對錯誤：' + name);
    }
    i = lt + m[0].length;
  }
  if (stack.length) errors.push('未結束的標籤：' + stack.join(','));
  return { errors, tags, comments, styleText };
}
function tagMultiset(p) {
  const c = {};
  p.tags.forEach((x) => { const k = (x.closing ? '/' : '') + x.name; c[k] = (c[k] || 0) + 1; });
  return c;
}
function attrNames(p) {
  const set = new Set();
  p.tags.forEach((x) => Object.keys(x.attrs).forEach((k) => set.add(x.name + '@' + k)));
  return Array.from(set).sort();
}
function visibleText(html) {
  let s = html.replace(/<!--[\s\S]*?-->/g, '');
  s = s.replace(/<head>[\s\S]*?<\/head>/, '');
  s = s.replace(/<div class="preheader"[^>]*>[\s\S]*?<\/div>/, '');
  s = s.replace(/<[^>]+>/g, ' ');
  return decodeEntities(s);
}
// 比較 HTML 可見文字與純文字時要略過的「裝飾」：空白、零寬、全形冒號與括號、分隔線、以及純文字把「標籤；警示」串成一行用的 '；'
const STRIP = /[\s\u{200c}：（）=；]/gu;
const norm = (s) => s.replace(STRIP, '');
function preheaderText(html) {
  const m = html.match(/<div class="preheader"[^>]*>([\s\S]*?)<\/div>/);
  return m ? decodeEntities(m[1]).replace(/[\s\u{200c}]+/gu, ' ').trim() : null;
}
function width(str) {
  let w = 0;
  for (const ch of str) {
    const c = ch.codePointAt(0);
    w += (c >= 0x1100 && c <= 0x115f) || (c >= 0x2e80 && c <= 0xa4cf) || (c >= 0xac00 && c <= 0xd7a3) || (c >= 0xf900 && c <= 0xfaff) ||
      (c >= 0xfe30 && c <= 0xfe6f) || (c >= 0xff00 && c <= 0xff60) || (c >= 0xffe0 && c <= 0xffe6) || (c >= 0x20000 && c <= 0x3fffd) ? 2 : 1;
  }
  return w;
}
function hexLum(hex) {
  const v = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255).map((c) => (c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4)));
  return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2];
}
function contrast(a, b) {
  const la = hexLum(a);
  const lb = hexLum(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}
/** 取決策條裡某一格（以標題文字找）所在的 <td ...> 開頭標籤與整格內容 */
function cellOf(html, caption) {
  const idx = html.indexOf('>' + caption + '</div>');
  if (idx < 0) return null;
  const start = html.lastIndexOf('<td ', idx);
  const end = html.indexOf('</td>', idx);
  const chunk = html.slice(start, end);
  const open = chunk.slice(0, chunk.indexOf('>') + 1);
  const bg = (open.match(/ bgcolor="([^"]*)"/) || [])[1];
  const styleBg = (open.match(/background-color:([^;"]*)/) || [])[1];
  const texts = chunk.split(/<[^>]+>/).map((x) => decodeEntities(x).trim()).filter(Boolean);
  return { open, bg, styleBg, texts };
}

// ═════════════════════════════════════════════════════════════════════════
section('1 金額與時間格式', () => {
  const f = R.formatNtd;
  [
    [0, 'NT$ 0'], [1, 'NT$ 0.01'], [9, 'NT$ 0.09'], [10, 'NT$ 0.10'], [99, 'NT$ 0.99'], [100, 'NT$ 1'], [101, 'NT$ 1.01'],
    [150, 'NT$ 1.50'], [99999, 'NT$ 999.99'], [100000, 'NT$ 1,000'], [100001, 'NT$ 1,000.01'], [123456789, 'NT$ 1,234,567.89'],
    [100000000, 'NT$ 1,000,000'], [1e12, 'NT$ 10,000,000,000'], [1e12 + 1, 'NT$ 10,000,000,000.01'],
    [Number.MAX_SAFE_INTEGER, 'NT$ 90,071,992,547,409.91'], [5000000000, 'NT$ 50,000,000'],
  ].forEach(([c, want]) => eq('formatNtd(' + c + ')', f(c), want));
  [-1, -100, 1.5, 0.1 + 0.2, NaN, Infinity, -Infinity, '100', null, undefined, {}, [], 2 ** 53, Number.MAX_SAFE_INTEGER + 2, 10n, true].forEach((bad) => {
    throwsErr('formatNtd 拒絕 ' + short(typeof bad === 'bigint' ? String(bad) + 'n' : bad), () => f(bad), MailRenderError, 'BAD_AMOUNT');
  });
  // 與 BigInt 對拍：10 萬組隨機值＋各位數邊界
  let seed = 7;
  const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
  const ref = (n) => {
    const b = BigInt(n);
    const yuan = (b / 100n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    const fen = b % 100n;
    return 'NT$ ' + yuan + (fen === 0n ? '' : '.' + (fen < 10n ? '0' : '') + fen.toString());
  };
  let bad = 0;
  let n = 0;
  const check = (c) => { n++; if (f(c) !== ref(c)) { bad++; if (bad < 4) record('對拍不一致 ' + c, false, f(c) + ' vs ' + ref(c)); } };
  for (let i = 0; i < 100000; i++) check(Math.floor(rnd() * Math.pow(10, 1 + Math.floor(rnd() * 15))));
  for (let p = 0; p <= 15; p++) { const b = Math.pow(10, p); [b - 1, b, b + 1, b * 5 - 1].forEach((c) => { if (c >= 0 && c <= Number.MAX_SAFE_INTEGER) check(c); }); }
  [999, 1000, 99900, 100000, 99999999, 100000000].forEach(check);
  t('formatNtd 與 BigInt 參考實作一致（' + n + ' 組）', bad === 0, bad);
  const src = f.toString();
  t('formatNtd 原始碼沒有浮點除法／捨入（/ 100、Math.floor、Math.round、toFixed、parseFloat、* 0.01）', !/\/\s*100|Math\.(floor|round|trunc|ceil)|toFixed|parseFloat|\*\s*0\.01/.test(src), src.length);

  const tp = R.formatTaipei;
  [
    ['2026-10-08T06:30:00.000Z', '2026-10-08 14:30'], ['2026-12-31T16:30:00Z', '2027-01-01 00:30'], ['2026-10-08T14:30:00+08:00', '2026-10-08 14:30'],
    ['2026-10-08T00:00:00-05:00', '2026-10-08 13:00'], ['2028-02-29T15:59:00Z', '2028-02-29 23:59'], ['2028-02-29T16:00:00Z', '2028-03-01 00:00'],
    ['2026-01-01T00:00:00Z', '2026-01-01 08:00'], ['2026-06-30T23:59:59.999Z', '2026-07-01 07:59'],
  ].forEach(([iso, want]) => eq('formatTaipei ' + iso, tp(iso), want));
  [['垃圾', ''], [undefined, ''], ['', '']].forEach(([x, want]) => eq('formatTaipei 無法解析 ' + short(x), tp(x), want));
});

// ═════════════════════════════════════════════════════════════════════════
section('2 事件×收件人矩陣與可見性', () => {
  eq('ALLOWED_KINDS 與獨立抄錄的矩陣一致', JSON.parse(JSON.stringify(R.ALLOWED_KINDS)), EXPECT_ALLOWED);
  t('ALLOWED_KINDS 已凍結', Object.isFrozen(R.ALLOWED_KINDS) && Object.values(R.ALLOWED_KINDS).every((a) => Object.isFrozen(a)));
  eq('事件類型 6 種', EVENT_TYPES.length, 6);
  eq('收件人 kind 7 種', KINDS.length, 7);
  KINDS.forEach((k) => {
    const v = visibilityFor(k);
    ['amount', 'margin', 'tier', 'customer', 'owner', 'items', 'project'].forEach((f) => t('visibility 表（獨立抄錄）' + k + '.' + f, v[f] === !!EXPECT_VIS[k][f]));
  });

  let renderable = 0;
  let rejected = 0;
  EVENT_TYPES.forEach((type) => KINDS.forEach((kind) => {
    const allowed = EXPECT_ALLOWED[type].indexOf(kind) >= 0;
    const tag = type + ' × ' + kind;
    if (!allowed) {
      rejected++;
      throwsErr(tag + ' 不合理組合 throw KIND_NOT_ALLOWED', () => render(type, kind), MailRenderError, 'KIND_NOT_ALLOWED');
      return;
    }
    renderable++;
    let m;
    try { m = render(type, kind); } catch (e) { record(tag + ' 可渲染', false, e && e.message); return; }
    t(tag + ' 回傳 subject／html／text／meta', typeof m.subject === 'string' && typeof m.html === 'string' && typeof m.text === 'string' && m.meta && typeof m.meta === 'object');
    const V = EXPECT_VIS[kind];
    const both = (s) => m.html.indexOf(s) >= 0 || m.text.indexOf(s) >= 0;
    const bothHas = (s) => m.html.indexOf(s) >= 0 && m.text.indexOf(s) >= 0;
    const bothNo = (s) => m.html.indexOf(s) < 0 && m.text.indexOf(s) < 0;

    // 主旨：格式正確，且絕不含金額／毛利率／客戶名／業務名
    // 看不到專案名稱的 kind（秘書／董事會代核人）：主旨只有類別與單號
    eq(tag + ' 主旨格式', m.subject, SUBJECT_PREFIX[type] + 'QU-TEST-001' + (V.project ? ' ' + CAN.project : '') + (type === 'E6_WITHDRAWN' ? ' 請勿簽核' : ''));
    ['1,234,567', '123456789', '41.04', 'NT$', CAN.company, CAN.owner, CAN.tierLabel, '毛利', CAN.viewer].forEach((s) => t(tag + ' 主旨不含 ' + s, m.subject.indexOf(s) < 0));

    // 共通：單號、專案名稱、連結、機密字樣
    t(tag + ' 含單號（HTML 與純文字）', bothHas('QU-TEST-001'));
    t(tag + ' 專案名稱 ' + (V.project ? '出現（HTML 與純文字）' : '不出現（主旨、HTML、純文字、meta）'), V.project ? bothHas(CAN.project) : bothNo(CAN.project) && m.subject.indexOf(CAN.project) < 0 && JSON.stringify(m.meta).indexOf(CAN.project) < 0);
    t(tag + ' 「專案名稱」欄位列 ' + (V.project ? '存在' : '不存在'), V.project ? both('專案名稱') : (m.html.indexOf('專案名稱') < 0 && m.text.indexOf('專案名稱') < 0));
    t(tag + ' 含預期連結（HTML href 與純文字）', m.html.indexOf('href="' + urlFor(type) + '"') >= 0 && m.text.split('\n').indexOf(urlFor(type)) >= 0, urlFor(type));
    t(tag + ' 含機密字樣', bothHas('機密，請勿轉寄'));
    t(tag + ' 稱呼用收件人名稱', bothHas(CAN.viewer + '，您好'));

    // 可見性（以獨立抄錄的表為準）
    t(tag + ' 客戶名 ' + (V.customer ? '出現' : '不出現'), V.customer ? bothHas(CAN.company) : bothNo(CAN.company));
    t(tag + ' 業務名 ' + (V.owner ? '出現' : '不出現'), V.owner ? bothHas(CAN.owner) : bothNo(CAN.owner));
    t(tag + ' 品項 ' + (V.items ? '出現' : '不出現'), V.items ? bothHas('品項Phi一') && bothHas('品項Phi二') : bothNo('品項Phi一') && bothNo('品項Phi二'));
    ITEM_PRICE_CANARIES.forEach((c) => t(tag + ' 不含品項價格 ' + c, bothNo(c)));
    const showActor = ACTOR_ALWAYS.indexOf(type) >= 0 || V.owner;
    t(tag + ' 操作人 ' + (showActor ? '出現' : '不出現'), showActor ? bothHas(CAN.actor) : bothNo(CAN.actor));
    const decision = DECISION_TYPES.indexOf(type) >= 0;
    const amt = !!(decision && V.amount);
    t(tag + ' 金額 ' + (amt ? '出現' : '不出現'), amt ? bothHas('NT$ 1,234,567.89') && bothHas('折扣後未稅') : bothNo('1,234,567') && bothNo('123456789') && bothNo('NT$'));
    t(tag + ' 毛利率 ' + (decision && V.margin ? '出現' : '不出現'), decision && V.margin ? bothHas('41.04%') : bothNo('41.04') && bothNo('毛利率'));
    t(tag + ' 核決層級 ' + (decision && V.tier ? '出現' : '不出現'), decision && V.tier ? bothHas(CAN.tierLabel) && bothHas('核決層級') : bothNo(CAN.tierLabel) && bothNo('核決層級'));
    t(tag + ' meta.hasAmount＝' + amt, m.meta.hasAmount === amt);
    eq(tag + ' meta 欄位固定且內容正確', m.meta, { type, quoteId: QID, quoteNo: 'QU-TEST-001', kind, hasAmount: amt });
    const metaJson = JSON.stringify(m.meta);
    [CAN.company, CAN.project, CAN.owner, '1,234,567', '123456789', '41.04', '@'].forEach((s) => t(tag + ' meta 不含 ' + s, metaJson.indexOf(s) < 0));

    // 事件專屬
    if (type === 'E4_RESULT' || type === 'E6_WITHDRAWN') t(tag + ' 顯示結果原因', bothHas(CAN.reason));
    if (type === 'E2_COST_REQUEST') {
      const vt = visibleText(m.html);
      [m.text, vt].forEach((s, i) => {
        const w = i ? 'HTML 可見文字' : '純文字';
        t(tag + ' ' + w + ' 無 % 字元（無百分比樣式）', s.indexOf('%') < 0);
        t(tag + ' ' + w + ' 無金額樣式（NT$、$數字、千分位數字）', !/NT\$|\$\s*\d|\d,\d{3}/.test(s));
        ['毛利', '折扣', '單價', '未稅', '報價金額', '核決', CAN.company, CAN.tierLabel, '123456789', '1,234,567', '41.04'].forEach((k) => t(tag + ' ' + w + ' 不含「' + k + '」', s.indexOf(k) < 0));
      });
    }
    if (kind === 'secretary' || kind === 'boardProxy') {
      t(tag + ' 不含客戶名與業務名（含 html／text／subject／meta）', [m.html, m.text, m.subject, metaJson].every((s) => s.indexOf(CAN.company) < 0 && s.indexOf(CAN.owner) < 0));
      t(tag + ' 不含專案名稱（含 html／text／subject／meta／preheader）', [m.html, m.text, m.subject, metaJson, preheaderText(m.html) || ''].every((s) => s.indexOf(CAN.project) < 0));
    }
    if (kind === 'owner') {
      LEAKS_AMOUNT.forEach((k) => t(tag + ' 業務信不含「' + k + '」', both(k) === false));
    }
    if (V.amount && decision) {
      t(tag + ' 簽核人信含金額與毛利', bothHas('NT$ 1,234,567.89') && bothHas('41.04%'));
    }
    if (type === 'E6_WITHDRAWN' || type === 'E4_RESULT' || type === 'E5_COST_DONE') {
      t(tag + ' 非「請您簽核」的信不放金額與毛利率', bothNo('NT$') && bothNo('毛利率') && bothNo('41.04'));
    }
  }));
  eq('可渲染組合共 19 種', renderable, 19);
  eq('不合理組合共 23 種', rejected, 23);

  // 模型層：buildModel 不檢查「該不該寄」的矩陣，所以這裡可以對 6 事件 × （7 種 kind ＋ 5 種未知 kind）全部 72 種組合驗證，
  // 可見性一律由 visibility 表決定（就算未來有人放寬矩陣，也不會讓顧問／業務／未知類型看到不該看的資料）
  const kindsAll = KINDS.concat(['bogus', 'MGR1', '__proto__', 'constructor', '']);
  const NO_VIS = { amount: 0, margin: 0, tier: 0, customer: 0, owner: 0, items: 0, project: 0 };
  let modelChecks = 0;
  EVENT_TYPES.forEach((type) => kindsAll.forEach((kind) => {
    const V = EXPECT_VIS[kind] || NO_VIS;
    const tag = '模型層 ' + type + '×' + JSON.stringify(kind);
    let model;
    try { model = R.buildModel(mkEv(type, kind), mkViewer(kind), CTX); } catch (e) { record(tag + ' 可建立模型', false, e && e.message); return; }
    modelChecks++;
    const j = JSON.stringify(model);
    const decision = DECISION_TYPES.indexOf(type) >= 0;
    t(tag + ' 專案名稱 ' + (V.project ? '有' : '無') + '（含 title／preheader／rows）', j.indexOf(CAN.project) >= 0 === !!V.project);
    t(tag + ' 客戶名 ' + (V.customer ? '有' : '無'), j.indexOf(CAN.company) >= 0 === !!V.customer);
    t(tag + ' 業務名 ' + (V.owner ? '有' : '無'), j.indexOf(CAN.owner) >= 0 === !!V.owner);
    t(tag + ' 品項 ' + (V.items ? '有' : '無'), j.indexOf('品項Phi一') >= 0 === !!V.items);
    t(tag + ' 金額 ' + (decision && V.amount ? '有' : '無'), j.indexOf('1,234,567.89') >= 0 === !!(decision && V.amount));
    t(tag + ' 毛利率 ' + (decision && V.margin ? '有' : '無'), j.indexOf('41.04%') >= 0 === !!(decision && V.margin));
    t(tag + ' 核決層級 ' + (decision && V.tier ? '有' : '無'), j.indexOf(CAN.tierLabel) >= 0 === !!(decision && V.tier));
    ITEM_PRICE_CANARIES.forEach((c) => t(tag + ' 無品項價格 ' + c, j.indexOf(c) < 0));
  }));
  eq('模型層共檢查 72 種組合（6 事件 × 12 種 kind）', modelChecks, 72);
  t('buildModel 對未知 kind 使用最小資料（沒有決策條、沒有客戶、沒有業務、沒有品項）', (() => {
    const mm = R.buildModel(mkEv('E1_SUBMIT', 'bogus'), mkViewer('bogus'), CTX);
    return mm.decision === null && mm.items === null && mm.rows.every((r) => ['報價單號', '目前關卡', '時間'].indexOf(r.label) >= 0) && JSON.stringify(mm).indexOf(CAN.project) < 0;
  })());

  // 同一封信中不可有 undefined／null／NaN／[object 這類占位符
  EVENT_TYPES.forEach((type) => EXPECT_ALLOWED[type].forEach((kind) => {
    const m = render(type, kind);
    ['undefined', 'null', 'NaN', '[object', 'Infinity'].forEach((s) => t(type + '×' + kind + ' 不含占位符 ' + s, m.html.indexOf(s) < 0 && m.text.indexOf(s) < 0 && m.subject.indexOf(s) < 0));
  }));

  // ── 專案名稱哨兵（業主 2026-10-08 決定）──────────────────────────────────────
  // 秘書與董事會代核人的信，所有事件、所有輸出面（主旨、<title>、preheader、HTML 本文、純文字、meta、model）都不得帶出專案名稱；
  // 其餘 kind（含顧問、業務本人）照舊含專案名稱。哨兵字串含頭尾標記，連「截斷後的片段」也會被抓到。
  {
    const SP = 'ZQXPROJ哨兵專案名稱ZQXEND';
    const PARTS = ['ZQXPROJ', 'ZQXEND', '哨兵專案名稱'];
    let hiddenChecks = 0;
    let shownChecks = 0;
    EVENT_TYPES.forEach((type) => EXPECT_ALLOWED[type].forEach((kind) => {
      const tag = '哨兵 ' + type + '×' + kind;
      const ev = mkEv(type, kind, { projectName: SP });
      const m = renderMail(ev, mkViewer(kind), CTX);
      const model = R.buildModel(ev, mkViewer(kind), CTX);
      const titleM = /<title>([\s\S]*?)<\/title>/.exec(m.html);
      const ph = preheaderText(m.html) || '';
      const surfaces = { subject: m.subject, title: titleM ? titleM[1] : '', preheader: ph, html: m.html, text: m.text, meta: JSON.stringify(m.meta), model: JSON.stringify(model), modelTitle: model.title, modelPreheader: model.preheader };
      if (EXPECT_VIS[kind].project) {
        shownChecks++;
        ['subject', 'title', 'preheader', 'html', 'text'].forEach((k) => t(tag + ' 看得到專案名稱：' + k + ' 含專案名稱', surfaces[k].indexOf(SP) >= 0, surfaces[k].slice(0, 80)));
      } else {
        hiddenChecks++;
        Object.keys(surfaces).forEach((k) => t(tag + ' 不得含專案名稱（' + k + '）', PARTS.every((p) => surfaces[k].indexOf(p) < 0), surfaces[k].slice(0, 80)));
        t(tag + ' 主旨只有類別與單號', m.subject === SUBJECT_PREFIX[type] + 'QU-TEST-001' + (type === 'E6_WITHDRAWN' ? ' 請勿簽核' : ''), m.subject);
        t(tag + ' preheader 只有單號', /^(待您簽核|簽核結果|請勿簽核)：QU-TEST-001$/.test(ph), ph);
        t(tag + ' 沒有「專案名稱」欄位列', m.html.indexOf('專案名稱') < 0 && m.text.indexOf('專案名稱') < 0 && model.rows.every((r) => r.label !== '專案名稱'));
      }
      t(tag + ' <title> 與主旨一致', surfaces.title === m.subject.replace(/&/g, '&amp;'), surfaces.title);
    }));
    t('哨兵：秘書／董事會代核人共 ' + hiddenChecks + ' 個事件×kind 組合（E1／E3／E6 各 2 種 kind ＝ 6）', hiddenChecks === 6, hiddenChecks);
    t('哨兵：其餘 kind 共 ' + shownChecks + ' 個組合照舊含專案名稱（19 − 6 ＝ 13）', shownChecks === 13, shownChecks);

    // 專案名稱「被當成別的欄位」的繞路：業務可輸入的自由文字（reason）、操作人、業務名都不會變成專案名稱的來源；
    // 同一事件裡專案名稱只能從 projectName 一個欄位進信，秘書信（E6）即使 reason 有值也不會把 projectName 帶出來
    const evR = mkEv('E6_WITHDRAWN', 'secretary', { projectName: SP, result: { kind: 'withdrawn', reason: 'ZQXREASON' } });
    const mr = renderMail(evR, mkViewer('secretary'), CTX);
    t('E6×secretary：reason 照常顯示、專案名稱仍不出現（reason 與 projectName 是兩個獨立欄位）', mr.html.indexOf('ZQXREASON') >= 0 && PARTS.every((p) => mr.html.indexOf(p) < 0 && mr.text.indexOf(p) < 0 && mr.subject.indexOf(p) < 0));
    const evR2 = mkEv('E4_RESULT', 'owner', { projectName: '', result: { kind: 'rejected', reason: SP } });
    const mr2 = renderMail(evR2, mkViewer('owner'), CTX);
    t('E4×owner：專案名稱為空時，reason 不會被提升成專案名稱（主旨、preheader 沒有 reason 文字、也沒有「專案名稱」列）', mr2.subject.indexOf('ZQXPROJ') < 0 && (preheaderText(mr2.html) || '').indexOf('ZQXPROJ') < 0 && R.buildModel(evR2, mkViewer('owner'), CTX).rows.every((r) => r.label !== '專案名稱'), mr2.subject);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('3 決策條與色塊', () => {
  const kinds = ['mgr1', 'gm', 'chairman', 'secretary', 'boardProxy'];
  // numbers 為 null：整段決策條完全不輸出（連外層的 table／tr／td 都不留，不是輸出一個空的決策條）
  kinds.forEach((kind) => {
    const a = tagMultiset(parseHtml(render('E1_SUBMIT', kind).html));
    const b = tagMultiset(parseHtml(render('E1_SUBMIT', kind, { numbers: null }).html));
    t('E1×' + kind + ' numbers=null：少了整段決策條結構（table -1、tr -2、td -4）', a.table - b.table === 1 && a['/table'] - b['/table'] === 1 && a.tr - b.tr === 2 && a.td - b.td === 4, short([a.table - b.table, a.tr - b.tr, a.td - b.td]));
  });
  ['E1_SUBMIT', 'E3_NEXT_STEP'].forEach((type) => kinds.forEach((kind) => {
    [null, undefined].forEach((nv) => {
      const m = render(type, kind, { numbers: nv });
      const vt = visibleText(m.html);
      ['報價金額', '毛利率', '核決層級', '折扣後未稅', 'NT$', '%'].forEach((s) => t(type + '×' + kind + ' numbers=' + nv + ' 不輸出「' + s + '」', vt.indexOf(s) < 0 && m.text.indexOf(s) < 0));
      t(type + '×' + kind + ' numbers=' + nv + ' 沒有決策條儲存格', m.html.indexOf('class="stack') < 0 && m.html.indexOf('class="amt"') < 0 && m.html.indexOf('amt"') < 0);
      t(type + '×' + kind + ' numbers=' + nv + ' meta.hasAmount＝false', m.meta.hasAmount === false);
      t(type + '×' + kind + ' numbers=' + nv + ' 無占位符（— ? 空括號）', !/（）|\(\)|—|\?\?/.test(m.text));
    });
  }));

  // 色塊：tierLevel → 色；文字標籤由 tierLabel 組成
  const TONE = { green: '#15803d', amber: '#b45309', red: '#b91c1c', grey: '#4b5563' };
  const cases = [
    [1, '一級主管', 'green', '一級主管可核'], [2, '總經理', 'amber', '需總經理核准'], [3, '董事長', 'red', '需董事長核准'],
    [3, '董事會', 'red', '需董事會決議'], [null, '董事會', 'grey', '需董事會決議'], [null, '總經理', 'grey', '需總經理核准'],
    [undefined, '董事長', 'grey', '需董事長核准'],
  ];
  cases.forEach(([lv, label, tone, tag]) => {
    const m = render('E3_NEXT_STEP', 'gm', { numbers: Object.assign(NUM(), { tierLevel: lv, tierLabel: label }) });
    const c = cellOf(m.html, '毛利率');
    t('tierLevel=' + lv + ' label=' + label + ' → 毛利率色塊 ' + tone, !!c && c.bg === TONE[tone] && c.styleBg === TONE[tone], short(c && [c.bg, c.styleBg]));
    t('tierLevel=' + lv + ' label=' + label + ' → 色塊內有文字標籤「' + tag + '」', !!c && c.texts.indexOf(tag) >= 0 && c.texts.indexOf('41.04%') >= 0 && c.texts.indexOf('毛利率') >= 0, short(c && c.texts));
    t('tierLevel=' + lv + ' label=' + label + ' → 純文字版有同樣標籤', m.text.indexOf('毛利率：41.04%（' + tag + '）') >= 0);
    const tc = cellOf(m.html, '核決層級');
    t('tierLevel=' + lv + ' → 核決層級格顯示 ' + label, !!tc && tc.texts.indexOf(label) >= 0 && !tc.bg.match(/#(15803d|b45309|b91c1c)/));
  });
  // 其他 tierLevel 值（驗證會擋掉）
  [0, 4, '1', 1.5, 'board', NaN].forEach((lv) => throwsErr('tierLevel=' + short(lv) + ' 被事件驗證擋掉', () => render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), { tierLevel: lv }) }), MailRenderError, 'BAD_EVENT'));

  // 每個有色塊的儲存格一定有文字標籤（不靠顏色單獨傳達）
  ['E1_SUBMIT', 'E3_NEXT_STEP'].forEach((type) => kinds.forEach((kind) => [1, 2, 3, null].forEach((lv) => {
    const m = render(type, kind, { numbers: Object.assign(NUM(), { tierLevel: lv }) });
    const re = /<td [^>]*class="stack"[^>]*>[\s\S]*?<\/td>/g;
    let mm;
    let n = 0;
    while ((mm = re.exec(m.html)) !== null) {
      n++;
      const texts = mm[0].split(/<[^>]+>/).map((x) => x.trim()).filter(Boolean);
      t(type + '×' + kind + ' lv=' + lv + ' 色塊儲存格有 3 段文字（標題／數值／標籤）', texts.length === 3, short(texts));
    }
    t(type + '×' + kind + ' lv=' + lv + ' 恰有 1 個色塊儲存格', n === 1, n);
  })));

  // 毛利率文字：沿用 marginText，不重算；補 % 不重複
  [['41.04%', '41.04%'], ['41.04', '41.04%'], ['0.00%', '0.00%'], ['100.00%', '100.00%'], ['99.99', '99.99%'], ['-0.50%', '-0.50%'], ['1234.5678%', '1234.5678%']].forEach(([mt, want]) => {
    const m = render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), { marginText: mt, marginPct: parseFloat(mt) || 0 }) });
    const c = cellOf(m.html, '毛利率');
    t('marginText ' + mt + ' → 顯示 ' + want, !!c && c.texts.indexOf(want) >= 0 && m.html.indexOf('%%') < 0 && m.text.indexOf('%%') < 0, short(c && c.texts));
  });
  // marginPct 與 marginText 刻意不一致：顯示的是 marginText（不重算）
  const mismatch = render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), { marginText: '12.34%', marginPct: 77.7 }) });
  t('顯示 marginText 而不是重算 marginPct', mismatch.html.indexOf('12.34%') >= 0 && mismatch.html.indexOf('77.7') < 0);
  // 毛利為負
  [{ gpCents: -1, marginPct: 5 }, { gpCents: 100, marginPct: -0.01 }, { gpCents: -100, marginPct: -3.2 }].forEach((o, i) => {
    const m = render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), o, { marginText: i === 2 ? '-3.20%' : '5.00%' }) });
    t('負毛利案例 ' + i + ' 色塊標籤含「毛利為負」', cellOf(m.html, '毛利率').texts.some((x) => x.indexOf('毛利為負') >= 0) && m.text.indexOf('毛利為負') >= 0);
  });
  [{ gpCents: 0, marginPct: 0 }, { gpCents: 5, marginPct: 0.01 }].forEach((o, i) => {
    const m = render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), o, { marginText: '0.00%' }) });
    t('非負毛利案例 ' + i + ' 不顯示「毛利為負」', m.html.indexOf('毛利為負') < 0 && m.text.indexOf('毛利為負') < 0);
  });
  throwsErr('marginPct NaN 被拒絕', () => render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), { marginPct: NaN }) }), MailRenderError, 'BAD_EVENT');
  throwsErr('marginPct Infinity 被拒絕', () => render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), { marginPct: Infinity }) }), MailRenderError, 'BAD_EVENT');

  // 金額進信的格式
  [[0, 'NT$ 0'], [1, 'NT$ 0.01'], [99, 'NT$ 0.99'], [1e12, 'NT$ 10,000,000,000'], [100, 'NT$ 1'], [123456789, 'NT$ 1,234,567.89']].forEach(([c, want]) => {
    const m = render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), { revenueCents: c }) });
    const ac = cellOf(m.html, '報價金額');
    t('信中金額 ' + c + ' 分 → ' + want + '（標「折扣後未稅」）', !!ac && ac.texts.indexOf(want) >= 0 && ac.texts.indexOf('折扣後未稅') >= 0 && m.text.indexOf('報價金額：' + want + '（折扣後未稅）') >= 0, short(ac && ac.texts));
  });
  [-1, -100, 1.5, '100', NaN, null, undefined, 2 ** 53].forEach((c) => throwsErr('信中金額 ' + short(c) + ' 被拒絕', () => render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), { revenueCents: c }) }), MailRenderError, 'BAD_EVENT'));

  // 版面順序：決策條在稱呼與單據資訊之上；按鈕在資訊之下
  const m1 = render('E1_SUBMIT', 'mgr1');
  const iBar = m1.html.indexOf('>報價金額<');
  t('決策條在稱呼之上', iBar > 0 && iBar < m1.html.indexOf(CAN.viewer + '，您好'));
  t('決策條在單據資訊之上', iBar < m1.html.indexOf('>報價單號<'));
  t('按鈕在單據資訊之下', m1.html.indexOf('前往系統簽核') > m1.html.indexOf('>報價單號<'));
  t('純文字版：決策條在單號之前', m1.text.indexOf('報價金額：') < m1.text.indexOf('報價單號：'));
  t('決策條放在 E1／E3 的最上方（標題之後、第一個資訊區塊之前）', m1.html.indexOf('報價單待簽核') < iBar);
  // 單格／雙格退化（政策表目前不會發生，但版面要能處理）
  const lay = T.layoutHtml;
  const baseModel = {
    title: 'x', preheader: 'p', brand: 'B', confidential: 'C', headline: 'H', headlineTone: 'blue', result: null, greeting: 'G', lead: ['L'],
    rows: [{ label: 'a', value: 'b' }], items: null, button: { label: 'btn', url: BASE + '/q/x' }, notes: [], footer: ['F'],
  };
  [1, 2, 3].forEach((n) => {
    const cells = [{ kind: 'amount', caption: 'c1', value: 'v1', tag: 't1', tone: null }, { kind: 'margin', caption: 'c2', value: 'v2', tag: 't2', tone: 'amber' }, { kind: 'tier', caption: 'c3', value: 'v3', tag: '', tone: null }].slice(0, n);
    const h = lay(Object.assign({}, baseModel, { decision: { cells } }));
    const p = parseHtml(h);
    t('決策條 ' + n + ' 格仍是良構 HTML', p.errors.length === 0, p.errors.join(';'));
    t('決策條 ' + n + ' 格 width 加總 100%', (h.match(/class="stack[^"]*" width="(\d+)%"/g) || []).map((x) => parseInt(x.match(/width="(\d+)%"/)[1], 10)).reduce((a, b) => a + b, 0) === 100);
  });
});

// ═════════════════════════════════════════════════════════════════════════
section('3b 極端虧損單：毛利率位數再多也要寄得出去', () => {
  // lib/quoteApproval.js 的 marginText 對負毛利沒有上限；成本 ≥ 營收 1 萬倍時整數部分超過 6 位。
  // 以前 validateEvent 直接退件（BAD_EVENT），簽核人完全收不到信。現在放行，信上顯示 <-999999%。
  const EXTREME = ['-1000000.00', '-100000000.00%', '-900719925474099100.00'];
  const MARGIN_KINDS = ['mgr1', 'gm', 'chairman', 'secretary', 'boardProxy'];
  const extremeNum = (mt) => Object.assign(NUM(), { gpCents: -100000000, marginText: mt, marginPct: Number(mt.replace(/%$/, '')) });
  DECISION_TYPES.forEach((type) => MARGIN_KINDS.forEach((kind) => EXTREME.forEach((mt) => {
    const tag = type + '×' + kind + ' marginText=' + mt;
    let m = null;
    let err = null;
    try { m = render(type, kind, { numbers: extremeNum(mt) }); } catch (e) { err = e; }
    t(tag + ' 可以渲染（不丟 BAD_EVENT）', !err && !!m, err && err.message);
    if (!m) return;
    const intPart = mt.replace(/^-|%$/g, '').split('.')[0];
    t(tag + ' HTML 顯示「&lt;-999999%」（< 已跳脫、沒有裸的 <-）', m.html.indexOf('&lt;-999999%') >= 0 && m.html.indexOf('<-999999') < 0);
    t(tag + ' 純文字顯示「<-999999%」', m.text.indexOf('<-999999%') >= 0);
    t(tag + ' 超長的原始數字不出現在信裡', m.html.indexOf(intPart) < 0 && m.text.indexOf(intPart) < 0);
    const vs = Number((m.html.match(/font-size:(\d+)px;[^"]*">&lt;-999999%</) || [])[1]);
    t(tag + ' 窄格子（30% 寬）裡的「<-999999%」字級縮到 22px 以下（26px 會折成「<-99999／9%」孤字）', vs > 0 && vs <= 22, vs);
    t(tag + ' 仍有「毛利為負」警示，且決策條照常（金額仍在）', m.html.indexOf('毛利為負') >= 0 && m.text.indexOf('毛利為負') >= 0 && m.text.indexOf('NT$ 1,234,567.89') >= 0, short(m.text.slice(0, 200)));
    t(tag + ' 是良構的 HTML（標籤配對正確）', parseHtml(m.html).errors.length === 0, parseHtml(m.html).errors.join(';'));
  })));
  // 正向極端值（理論上不會出現，但顯示規則對稱）
  ['1000000.00', '12345678901234567890'].forEach((mt) => {
    const m = render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), { marginText: mt, marginPct: Number(mt) }) });
    t('正向極端值 ' + mt + ' → 顯示 >999999%', m.text.indexOf('>999999%') >= 0 && m.html.indexOf('&gt;999999%') >= 0);
  });
  // 邊界：6 位整數仍原樣顯示
  [['-999999.99', '-999999.99%'], ['999999.99', '999999.99%'], ['-123456.78%', '-123456.78%']].forEach(([mt, want]) => {
    const m = render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), { gpCents: -1, marginText: mt, marginPct: Number(mt.replace(/%$/, '')) }) });
    t('邊界 ' + mt + ' 原樣顯示為 ' + want, m.text.indexOf(want) >= 0 && m.text.indexOf('<-999999') < 0 && m.text.indexOf('>999999') < 0, short(m.text.slice(0, 160)));
  });
  // 看不到毛利率的收件人（顧問、業務）：極端值不會洩漏、也不會讓渲染失敗
  ['consultant', 'owner'].forEach((kind) => {
    const type = kind === 'consultant' ? 'E2_COST_REQUEST' : 'E4_RESULT';
    let m = null;
    let err = null;
    try { m = render(type, kind, { numbers: extremeNum('-100000000.00') }); } catch (e) { err = e; }
    t(kind + ' 收到極端值事件：可渲染且不含毛利率相關字串', !err && !!m && m.html.indexOf('999999') < 0 && m.text.indexOf('毛利') < 0 && m.html.indexOf('100000000') < 0, err && err.message);
  });
  // 格式檢查仍然有效：含 HTML 的 marginText 一律退件
  ['<script>alert(1)</script>', '-100000000.00<b>', '1' + '0'.repeat(25)].forEach((mt) => {
    throwsErr('marginText=' + short(mt).slice(0, 30) + ' 仍被事件驗證擋掉', () => render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), { marginText: mt }) }), MailRenderError, 'BAD_EVENT');
  });
});

// ═════════════════════════════════════════════════════════════════════════
section('4 主旨', () => {
  const sub = (type, over) => render(type, type === 'E2_COST_REQUEST' ? 'consultant' : type === 'E4_RESULT' || type === 'E5_COST_DONE' ? 'owner' : 'mgr1', over).subject;
  const noCtl = (s) => !/[\x00-\x1f\x7f-\x9f]/.test(s) && s.indexOf(cp(0x2028)) < 0 && s.indexOf(cp(0x2029)) < 0 && s.indexOf(cp(0x202e)) < 0;
  eq('E1 主旨', sub('E1_SUBMIT', { projectName: '機房' }), '【簽核通知】QU-TEST-001 機房');
  // 秘書／董事會代核人：主旨只有類別與單號（業主 2026-10-08 決定）；subjectFor(ev, kind) 的 kind 缺或未知＝最小資料，也不放專案名稱
  ['secretary', 'boardProxy'].forEach((k) => ['E1_SUBMIT', 'E3_NEXT_STEP'].forEach((ty) => {
    eq(ty + '×' + k + ' 主旨只有類別與單號', render(ty, k, { projectName: '機房' }).subject, '【簽核通知】QU-TEST-001');
    eq('subjectFor(' + ty + ', ' + k + ')', R.subjectFor(mkEv(ty, k, { projectName: '機房' }), k), '【簽核通知】QU-TEST-001');
  }));
  ['secretary', 'boardProxy'].forEach((k) => eq('E6×' + k + ' 主旨：類別＋單號＋請勿簽核（無專案名稱）', render('E6_WITHDRAWN', k, { projectName: '機房' }).subject, '【簽核撤回】QU-TEST-001 請勿簽核'));
  eq('subjectFor(ev, mgr1) 含專案名稱', R.subjectFor(mkEv('E1_SUBMIT', 'mgr1', { projectName: '機房' }), 'mgr1'), '【簽核通知】QU-TEST-001 機房');
  [undefined, null, '', 'bogus', 'MGR1', '__proto__', 5].forEach((k) => eq('subjectFor(ev, ' + short(k) + ')：kind 缺或未知 → 不放專案名稱', R.subjectFor(mkEv('E1_SUBMIT', 'mgr1', { projectName: '機房' }), k), '【簽核通知】QU-TEST-001'));
  eq('subjectFor(ev)：沒給 kind → 不放專案名稱（最小資料）', R.subjectFor(mkEv('E1_SUBMIT', 'mgr1', { projectName: '機房' })), '【簽核通知】QU-TEST-001');
  eq('E3 主旨', sub('E3_NEXT_STEP', { projectName: '機房' }), '【簽核通知】QU-TEST-001 機房');
  eq('E2 主旨', sub('E2_COST_REQUEST', { projectName: '機房' }), '【請填寫成本】QU-TEST-001 機房');
  eq('E4 主旨', sub('E4_RESULT', { projectName: '機房' }), '【簽核結果】QU-TEST-001 機房');
  eq('E5 主旨', sub('E5_COST_DONE', { projectName: '機房' }), '【成本已填寫】QU-TEST-001 機房');
  eq('E6 主旨', sub('E6_WITHDRAWN', { projectName: '機房' }), '【簽核撤回】QU-TEST-001 機房 請勿簽核');
  eq('沒有專案名稱：只有類別與單號', sub('E1_SUBMIT', { projectName: '' }), '【簽核通知】QU-TEST-001');
  eq('專案名稱只有空白／控制字元：視為沒有', sub('E1_SUBMIT', { projectName: ' \t\n ' }), '【簽核通知】QU-TEST-001');
  eq('E6 沒有專案名稱仍有「請勿簽核」', sub('E6_WITHDRAWN', { projectName: '' }), '【簽核撤回】QU-TEST-001 請勿簽核');
  // 專案名稱截 40 字（超過加省略號；剛好 40 不加）
  eq('專案名稱剛好 40 字：不截斷', sub('E1_SUBMIT', { projectName: '甲'.repeat(40) }), '【簽核通知】QU-TEST-001 ' + '甲'.repeat(40));
  eq('專案名稱 41 字：截成 40 字＋省略號', sub('E1_SUBMIT', { projectName: '甲'.repeat(41) }), '【簽核通知】QU-TEST-001 ' + '甲'.repeat(40) + cp(0x2026));
  const longE6 = sub('E6_WITHDRAWN', { projectName: '甲'.repeat(300), quoteNo: 'N'.repeat(40) });
  t('最長情況：E6 主旨仍以「請勿簽核」結尾且 <= 120 字', longE6.endsWith(' 請勿簽核') && Array.from(longE6).length <= 120, Array.from(longE6).length);
  // emoji（代理對）不被切半
  const em = sub('E1_SUBMIT', { projectName: cp(0x1f600).repeat(41) });
  t('專案名稱 emoji 41 個：不產生孤立代理', !/\p{Cs}/u.test(em) && Array.from(em).filter((c) => c === cp(0x1f600)).length === 40, em);
  // 標頭注入
  const inj = [
    'x\r\nBcc: evil@evil.test', 'x\nSubject: pwn', 'x\rTo: a@b.test', 'a' + cp(0x2028) + 'b' + cp(0x2029) + 'c', 'a\x00b\x1bc\x7fd', 'a' + cp(0x85) + 'b',
    cp(0x202e) + 'evil' + cp(0x202c), 'a' + cp(0x200b) + 'b' + cp(0xfeff) + 'c', '\tTab\ttab',
  ];
  inj.forEach((p, i) => {
    const s = sub('E1_SUBMIT', { projectName: p });
    t('主旨注入案例 ' + i + '：單行、無控制字元與方向字元', noCtl(s) && s.split('\n').length === 1, short(s));
  });
  throwsErr('quoteNo 含換行 → BAD_EVENT', () => render('E1_SUBMIT', 'mgr1', { quoteNo: 'QU\r\nBcc: x' }), MailRenderError, 'BAD_EVENT');
  // quoteNo 的事件驗證只擋 C0 控制字元，所以方向字元、U+2028、零寬字元會一路走到渲染層，主旨與內文都必須清掉
  [cp(0x202e) + 'QU-9', 'QU' + cp(0x2028) + '9', 'QU' + cp(0x2029) + '9', 'QU' + cp(0x200b) + '9', cp(0xfeff) + 'QU-9', 'QU' + cp(0x85) + '9'].forEach((no, i) => {
    const mm = render('E1_SUBMIT', 'mgr1', { quoteNo: no });
    const bad = /[\u{202a}-\u{202e}\u{2066}-\u{2069}\u{2028}\u{2029}\u{200b}\u{feff}\u{85}]/u;
    t('quoteNo 特殊字元案例 ' + i + '：主旨、HTML、純文字都清乾淨', !bad.test(mm.subject) && !bad.test(mm.html.replace(/&zwnj;/g, '')) && !bad.test(mm.text), short(mm.subject));
  });
  // 主旨不含 HTML 跳脫（它是純文字標頭）
  eq('主旨不做 HTML 跳脫', sub('E1_SUBMIT', { projectName: 'A&B <C>' }), '【簽核通知】QU-TEST-001 A&B <C>');
  // 主旨與 model.title、<title> 一致
  const m = render('E1_SUBMIT', 'mgr1', { projectName: 'A&B <C>' });
  t('<title> 是跳脫後的主旨', m.html.indexOf('<title>【簽核通知】QU-TEST-001 A&amp;B &lt;C&gt;</title>') >= 0);
});

// ═════════════════════════════════════════════════════════════════════════
section('5 內容細節', () => {
  // 品項
  const items = (n) => Array.from({ length: n }, (_, i) => ({ desc: 'D' + (i + 1), qty: i + 1, unit: 'U' }));
  const rowsIn = (m) => (m.html.match(/>D\d+</g) || []).length;
  [[0, 0, ''], [1, 1, ''], [29, 29, ''], [30, 30, ''], [31, 30, '…另 1 項'], [35, 30, '…另 5 項'], [200, 30, '…另 170 項']].forEach(([n, shown, more]) => {
    const m = render('E2_COST_REQUEST', 'consultant', { items: items(n) });
    t('品項 ' + n + ' 筆：HTML 顯示 ' + shown + ' 列', rowsIn(m) === shown, rowsIn(m));
    t('品項 ' + n + ' 筆：純文字顯示 ' + shown + ' 列', (m.text.match(/^D\d+ /gm) || []).length === shown);
    t('品項 ' + n + ' 筆：' + (more ? '有「' + more + '」' : '沒有「另 N 項」'), more ? m.html.indexOf(more) >= 0 && m.text.indexOf(more) >= 0 : m.html.indexOf('另 ') < 0 && m.text.indexOf('另 ') < 0);
    t('品項 ' + n + ' 筆：' + (n ? '有' : '沒有') + '表頭', n ? m.html.indexOf('>品項說明<') >= 0 && m.text.indexOf('品項說明') >= 0 : m.html.indexOf('品項說明') < 0 && m.text.indexOf('品項說明') < 0);
  });
  throwsErr('品項 201 筆被事件驗證擋掉', () => render('E2_COST_REQUEST', 'consultant', { items: items(201) }), MailRenderError, 'BAD_EVENT');
  [[2, '2'], [1.5, '1.5'], [0.1 + 0.2, '0.3'], [0, '0'], [1e20, '-'], ['10', '10'], ['約 3', '約 3'], [123456, '123456'], [0.00001, '0']].forEach(([q, want]) => {
    const m = render('E2_COST_REQUEST', 'consultant', { items: [{ desc: 'Q', qty: q, unit: 'U' }] });
    const row = m.text.split('\n').find((l) => l.indexOf('Q  ') === 0);
    t('品項數量 ' + short(q) + ' → ' + want, row === 'Q  ' + want + '  U', row);
  });
  const nq = render('E2_COST_REQUEST', 'consultant', { items: [{ desc: 'Q', unit: 'U' }, { desc: 'R', qty: null, unit: null }] });
  t('品項沒有數量或單位：不輸出 undefined／null', nq.html.indexOf('undefined') < 0 && nq.text.indexOf('null') < 0 && nq.text.indexOf('undefined') < 0);
  throwsErr('品項數量為負 → BAD_EVENT', () => render('E2_COST_REQUEST', 'consultant', { items: [{ desc: 'Q', qty: -1, unit: 'U' }] }), MailRenderError, 'BAD_EVENT');
  throwsErr('品項沒有說明 → BAD_EVENT', () => render('E2_COST_REQUEST', 'consultant', { items: [{ qty: 1, unit: 'U' }] }), MailRenderError, 'BAD_EVENT');
  // 品項價格欄位名稱各種變體都不會被讀取（只取 desc／qty／unit）
  const priceKeys = ['price', 'unitPrice', 'unit_price', 'cost', 'amount', 'total', 'discount', 'listPrice', 'salePrice', '單價', 'margin'];
  const it = { desc: '品項X', qty: 1, unit: '式' };
  priceKeys.forEach((k, i) => { it[k] = 31415900 + i; });
  const pm = render('E2_COST_REQUEST', 'consultant', { items: [it] });
  priceKeys.forEach((k, i) => t('品項欄位 ' + k + ' 的值不進信', (pm.html + pm.text).indexOf(String(31415900 + i)) < 0));

  // 操作人與業務同名：看不到業務名的收件人，不能靠「操作人」那一列洩漏；看得到的人不重複列
  ['E1_SUBMIT', 'E3_NEXT_STEP', 'E6_WITHDRAWN'].forEach((type) => ['secretary', 'boardProxy'].forEach((kind) => {
    const mm = render(type, kind, { actor: { label: CAN.owner } });
    t(type + '×' + kind + ' 操作人＝業務名：不洩漏業務名', mm.html.indexOf(CAN.owner) < 0 && mm.text.indexOf(CAN.owner) < 0 && mm.subject.indexOf(CAN.owner) < 0);
  }));
  ['E1_SUBMIT', 'E3_NEXT_STEP', 'E6_WITHDRAWN'].forEach((type) => ['mgr1', 'gm', 'chairman'].forEach((kind) => {
    const mm = render(type, kind, { actor: { label: CAN.owner } });
    t(type + '×' + kind + ' 操作人＝業務名：只列一次（不重複）', mm.text.split(CAN.owner).length === 2 && mm.text.indexOf('操作人') < 0);
  }));
  t('E4 操作人（簽核人）對業務顯示', render('E4_RESULT', 'owner').text.indexOf('操作人：' + CAN.actor) >= 0);
  t('E3 操作人（上一關簽核人）對秘書顯示', render('E3_NEXT_STEP', 'secretary').text.indexOf('操作人：' + CAN.actor) >= 0);
  t('E1 操作人對一級主管顯示（業務名可見）', render('E1_SUBMIT', 'mgr1').text.indexOf('操作人：' + CAN.actor) >= 0);

  // 結果與原因
  const E4 = (kind, reason, extra) => render('E4_RESULT', 'owner', Object.assign({ result: { kind, reason } }, extra || {}));
  [['approved', '本關已核准', '#15803d'], ['final_approved', '已完成核准', '#15803d'], ['rejected', '已駁回', '#b91c1c'], ['returned', '已退回修改', '#b45309']].forEach(([k, label, color]) => {
    const m = E4(k, '理由XYZ');
    t('E4 ' + k + ' 顯示「' + label + '」與色塊', m.html.indexOf('>' + label + '<') >= 0 && m.html.indexOf('bgcolor="' + color + '"') >= 0 && m.text.indexOf('簽核結果：' + label) >= 0);
    t('E4 ' + k + ' 顯示原因', m.html.indexOf('理由XYZ') >= 0 && m.text.indexOf('原因：理由XYZ') >= 0);
  });
  [['withdrawn', '已撤回', '#b45309'], ['voided', '核准已作廢', '#b91c1c']].forEach(([k, label, color]) => {
    const m = render('E6_WITHDRAWN', 'gm', { result: { kind: k, reason: '理由XYZ' } });
    t('E6 ' + k + ' 顯示「' + label + '」與色塊', m.html.indexOf('>' + label + '<') >= 0 && m.html.indexOf('bgcolor="' + color + '"') >= 0 && m.text.indexOf('狀態：' + label) >= 0);
  });
  ['withdrawn', 'voided'].forEach((k) => throwsErr('E4 不接受 result.kind=' + k, () => E4(k, 'x'), MailRenderError, 'BAD_EVENT'));
  ['approved', 'final_approved', 'rejected', 'returned'].forEach((k) => throwsErr('E6 不接受 result.kind=' + k, () => render('E6_WITHDRAWN', 'gm', { result: { kind: k } }), MailRenderError, 'BAD_EVENT'));
  const none = E4('rejected', undefined);
  t('沒有原因：不輸出「原因」', none.html.indexOf('>原因<') < 0 && none.text.indexOf('原因') < 0);
  const emp = E4('rejected', '  \n ');
  t('原因只有空白：視為沒有', emp.html.indexOf('>原因<') < 0);
  const lr = E4('rejected', '字'.repeat(2000));
  const reasonLine = lr.text.split('\n').filter((l) => /^(原因：|  )/.test(l)).join('').replace(/^原因：/, '').replace(/\s/g, '');
  t('原因 2000 字：截成 200 字＋省略號', reasonLine === '字'.repeat(200) + cp(0x2026), reasonLine.length);
  throwsErr('原因 2001 字：事件驗證擋掉', () => E4('rejected', '字'.repeat(2001)), MailRenderError, 'BAD_EVENT');
  const ml = E4('rejected', '第一行\r\n第二行\n第三行');
  t('原因多行：合併成單行', ml.html.indexOf('第一行 第二行 第三行') >= 0 && ml.text.indexOf('原因：第一行 第二行 第三行') >= 0);

  // 時間
  const tm = render('E1_SUBMIT', 'mgr1', { at: '2026-12-31T16:30:00Z' });
  t('時間用台北時區 YYYY-MM-DD HH:mm', tm.html.indexOf('>2027-01-01 00:30<') >= 0 && tm.text.indexOf('時間：2027-01-01 00:30') >= 0);

  // 稱呼
  t('label 為空：「您好：」', render('E1_SUBMIT', 'mgr1', {}, '').text.indexOf('\n您好：\n') >= 0);
  t('label 只有空白與控制字元：「您好：」', render('E1_SUBMIT', 'mgr1', {}, ' \n\t ').text.indexOf('\n您好：\n') >= 0);
  t('label 換行被合併', render('E1_SUBMIT', 'mgr1', {}, '甲\n乙').text.indexOf('甲 乙，您好：') >= 0);
  t('label 缺省（undefined）：「您好：」', renderMail(mkEv('E1_SUBMIT', 'mgr1'), { kind: 'mgr1' }, CTX).text.indexOf('\n您好：\n') >= 0);
  t('label 超長：截成 60 字＋省略號', render('E1_SUBMIT', 'mgr1', {}, '名'.repeat(100)).text.replace(/\s/g, '').indexOf('名'.repeat(60) + cp(0x2026) + '，您好') >= 0);

  // 連結
  const ids = ['abc', 'A_b-9', '0b9f6c1e-aaaa-4bbb-8ccc-ddddeeeeffff', 'x'.repeat(64), 'legacy-12'];
  ids.forEach((id) => {
    const m = render('E1_SUBMIT', 'mgr1', { quoteId: id });
    t('連結 id=' + id.slice(0, 12), m.html.indexOf('href="' + BASE + '/q/' + id + '"') >= 0 && m.text.indexOf('\n' + BASE + '/q/' + id + '\n') >= 0);
  });
  t('E2 連結帶 ?cost=1', render('E2_COST_REQUEST', 'consultant').html.indexOf('href="' + BASE + '/q/' + QID + '?cost=1"') >= 0);
  EVENT_TYPES.filter((x) => x !== 'E2_COST_REQUEST').forEach((type) => t(type + ' 連結不帶查詢字串', render(type, EXPECT_ALLOWED[type][0]).html.indexOf('?cost') < 0));
  ['', ' ', 'a b', 'a/b', '../x', 'a.b', 'a?b=1', 'a#b', 'a%2Fb', '"><script>', 'x'.repeat(65), '中文', 'a\nb', "a'b", 'a\\b'].forEach((id) => {
    throwsErr('quoteId ' + short(id) + ' 被拒絕', () => render('E1_SUBMIT', 'mgr1', { quoteId: id }), MailRenderError, 'BAD_EVENT');
  });
  const lc = render('E1_SUBMIT', 'mgr1', {}, undefined, { config: { appBaseUrl: 'http://localhost:3000' } });
  t('appBaseUrl=http://localhost:3000 可用', lc.html.indexOf('href="http://localhost:3000/q/' + QID + '"') >= 0);
  const dc = render('E1_SUBMIT', 'mgr1', {}, undefined, { config: getMailConfig({}) });
  t('預設設定的連結以預設站台開頭', dc.html.indexOf('href="https://itts-crm.vercel.app/q/' + QID + '"') >= 0);
  t('buildQuoteUrl 與信件連結一致', R.buildQuoteUrl(CTX.config, QID, { cost: true }) === BASE + '/q/' + QID + '?cost=1' && R.buildQuoteUrl(CTX.config, QID) === BASE + '/q/' + QID);

  // preheader：單號＋專案名稱，不含金額與客戶／業務
  EVENT_TYPES.forEach((type) => EXPECT_ALLOWED[type].forEach((kind) => {
    const m = render(type, kind);
    const ph = preheaderText(m.html);
    const pv = !!EXPECT_VIS[kind].project;
    t(type + '×' + kind + ' 有 preheader 且含單號；專案名稱' + (pv ? '出現' : '不出現（只放單號）'), ph !== null && ph.indexOf('QU-TEST-001') >= 0 && (ph.indexOf(CAN.project) >= 0) === pv, ph);
    t(type + '×' + kind + ' preheader 不含金額、毛利率、客戶、業務', ph !== null && ['NT$', '41.04', '1,234,567', CAN.company, CAN.owner, '毛利', CAN.tierLabel].every((s) => ph.indexOf(s) < 0), ph);
    t(type + '×' + kind + ' preheader 是 body 的第一個元素且為隱藏', /<body[^>]*>\n<div class="preheader" style="display:none;/.test(m.html) && /mso-hide:all/.test(m.html));
  }));
  // 事件說明與關卡文字
  t('E1 一般關：說明與關卡', render('E1_SUBMIT', 'mgr1').text.indexOf('目前輪到您處理「' + CAN.step + '」這一關') >= 0);
  t('E3 一般關：說明', render('E3_NEXT_STEP', 'gm').text.indexOf('上一關已核准，現在輪到您處理') >= 0);
  t('E1 董事會關：說明改為登錄決議', render('E1_SUBMIT', 'secretary').text.indexOf('請依董事會決議至系統登錄') >= 0);
  t('E3 董事會關：說明改為登錄決議', render('E3_NEXT_STEP', 'boardProxy').text.indexOf('請依董事會決議至系統登錄') >= 0);
  t('E1／E3 提醒：信內不提供核准', render('E1_SUBMIT', 'mgr1').text.indexOf('本信不提供信內核准') >= 0);
  t('E6：請勿簽核', render('E6_WITHDRAWN', 'gm').text.indexOf('請勿再簽核') >= 0);
  t('E6 作廢：說明', render('E6_WITHDRAWN', 'gm', { result: { kind: 'voided' } }).text.indexOf('核准已作廢') >= 0);
  t('沒有任何一封信含 token／密碼樣式字樣', EVENT_TYPES.every((type) => EXPECT_ALLOWED[type].every((kind) => { const m = render(type, kind); return !/token|password|secret|bearer|api[_-]?key/i.test(m.html + m.text); })));
  t('沒有任何一封信含 mailto:／tel:／img／附件／script', EVENT_TYPES.every((type) => EXPECT_ALLOWED[type].every((kind) => { const m = render(type, kind); return !/mailto:|tel:|<img|<script|<iframe|<form|<link|<object|<embed|attachment/i.test(m.html); })));
});

// ═════════════════════════════════════════════════════════════════════════
const PAYLOADS = [
  '<script>alert(1)</script>',
  '"><img src=x onerror=alert(1)>',
  'javascript:alert(1)',
  'line1\r\nline2\nBcc: x@evil.test',
  cp(0x202e) + 'evil' + cp(0x202c) + ' rtl',
  "' onmouseover='alert(1)",
  '`><svg onload=alert(1)>',
  '&lt;b&gt;already&amp; &#60; &#x3c; &nbsp;',
  '</td></tr></table><div style="position:fixed;top:0">overlay</div>',
  'A'.repeat(100000),
  '{{7*7}}${7*7}%s%d%n',
  'a\x00b\x1bc\x07d',
  '<!-- x --><!--[if mso]><b>y</b><![endif]-->]]>',
  'a' + cp(0x2028) + 'b' + cp(0x2029) + 'c' + cp(0x85) + 'd' + cp(0x200b) + 'e' + cp(0xfeff) + 'f',
  // 短版（<= 18 字）：tierLabel（上限 20 字）、品項單位／數量（上限 20 字）這類短欄位也要真的走到渲染層
  '<script>1</script>',
  '"><svg/onload=1>',
  "'><img src=x>",
  '<b onclick=1>x',
];

section('6 XSS 與注入', () => {
  const fields = [
    { name: 'projectName', type: 'E1_SUBMIT', kind: 'mgr1', over: (p) => ({ projectName: p }) },
    { name: 'company', type: 'E1_SUBMIT', kind: 'mgr1', over: (p) => ({ company: p }) },
    { name: 'ownerLabel', type: 'E1_SUBMIT', kind: 'gm', over: (p) => ({ ownerLabel: p }) },
    { name: 'actor.label', type: 'E3_NEXT_STEP', kind: 'gm', over: (p) => ({ actor: { label: p } }) },
    { name: 'step.label', type: 'E1_SUBMIT', kind: 'mgr1', over: (p) => ({ step: { level: 1, label: p } }), maxLen: 40 },
    { name: 'tierLabel', type: 'E1_SUBMIT', kind: 'mgr1', over: (p) => ({ numbers: Object.assign(NUM(), { tierLabel: p }) }), maxLen: 20 },
    { name: 'quoteNo', type: 'E1_SUBMIT', kind: 'mgr1', over: (p) => ({ quoteNo: p }), maxLen: 40, noCtl: true },
    { name: 'result.reason', type: 'E4_RESULT', kind: 'owner', over: (p) => ({ result: { kind: 'rejected', reason: p } }) },
    // 品項固定 2 筆，與基準（mkEv 預設 2 筆）列數相同，標籤集合才能逐一比對
    { name: 'items.desc', type: 'E2_COST_REQUEST', kind: 'consultant', over: (p) => ({ items: [{ desc: p, qty: 1, unit: 'U' }, { desc: 'D2', qty: 2, unit: 'U' }] }) },
    { name: 'items.unit', type: 'E2_COST_REQUEST', kind: 'consultant', over: (p) => ({ items: [{ desc: 'D', qty: 1, unit: p }, { desc: 'D2', qty: 2, unit: 'U' }] }), maxLen: 20 },
    { name: 'items.qty(字串)', type: 'E2_COST_REQUEST', kind: 'consultant', over: (p) => ({ items: [{ desc: 'D', qty: p, unit: 'U' }, { desc: 'D2', qty: 2, unit: 'U' }] }), maxLen: 20 },
  ];
  const viewerLabelField = { name: 'viewer.label', type: 'E1_SUBMIT', kind: 'mgr1' };
  const baseline = new Map();
  const base = (f) => {
    const k = f.type + '/' + f.kind;
    if (!baseline.has(k)) {
      const m = render(f.type, f.kind);
      const p = parseHtml(m.html);
      baseline.set(k, { tags: tagMultiset(p), attrs: attrNames(p), url: urlFor(f.type) });
    }
    return baseline.get(k);
  };
  let combos = 0;
  PAYLOADS.forEach((payload, pi) => {
    const checkMail = (label, m, b, expectedNorm) => {
      combos++;
      const p = parseHtml(m.html);
      t(label + ' HTML 仍然良構', p.errors.length === 0, p.errors.slice(0, 3).join(';'));
      eq(label + ' 標籤集合與基準相同（沒有長出新標籤）', tagMultiset(p), b.tags);
      eq(label + ' 屬性集合與基準相同（沒有長出新屬性）', attrNames(p), b.attrs);
      t(label + ' 沒有 on* 屬性', p.tags.every((x) => Object.keys(x.attrs).every((k) => !/^on/i.test(k))));
      const hrefs = p.tags.filter((x) => x.attrs.href !== undefined).map((x) => x.attrs.href);
      t(label + ' href 恆為預期網址', hrefs.length === 2 && hrefs.every((h) => h === b.url), short(hrefs));
      t(label + ' 沒有 javascript: 出現在任何屬性值', p.tags.every((x) => Object.keys(x.attrs).every((k) => !/javascript:/i.test(x.attrs[k]))));
      t(label + ' 註解只有 [if mso] 那一個', p.comments.length === 1 && /^\[if mso\]/.test(p.comments[0]), p.comments.length);
      t(label + ' 方向控制／零寬字元不進信', !/[\u{202a}-\u{202e}\u{2066}-\u{2069}\u{200b}\u{200e}\u{200f}\u{feff}]/u.test(m.html.replace(/&zwnj;/g, '') + m.text));
      t(label + ' 純文字無控制字元（\\r、NUL、ESC）', !/[\r\x00\x07\x1b]/.test(m.text));
      if (expectedNorm !== undefined) {
        t(label + ' HTML 可見文字 ＝ 純文字', norm(visibleText(m.html)) === norm(m.text), norm(visibleText(m.html)).slice(0, 80) + ' ≠ ' + norm(m.text).slice(0, 80));
      }
      t(label + ' 大小 < 100KB', Buffer.byteLength(m.html) < 102400 && Buffer.byteLength(m.text) < 102400, Buffer.byteLength(m.html));
    };
    fields.forEach((f) => {
      const label = 'XSS#' + pi + ' ' + f.name;
      const b = base(f);
      let m;
      try { m = render(f.type, f.kind, f.over(payload)); } catch (e) {
        // 事件驗證可能依長度／控制字元擋掉：只接受 BAD_EVENT
        t(label + ' 若被拒絕必為 BAD_EVENT', e instanceof MailRenderError && e.code === 'BAD_EVENT', e && e.code);
        return;
      }
      checkMail(label, m, b, true);
    });
    // viewer.label
    const lm = render('E1_SUBMIT', 'mgr1', {}, payload);
    checkMail('XSS#' + pi + ' ' + viewerLabelField.name, lm, base(viewerLabelField), true);
  });
  t('XSS 組合數（' + combos + '）至少 100（其餘是被事件驗證依長度或控制字元擋掉的）', combos >= 100);

  // 具體檢查：典型 payload 被「跳脫成純文字」
  const sm = render('E1_SUBMIT', 'mgr1', { projectName: '<script>alert(1)</script>' });
  t('<script> 被跳脫為 &lt;script&gt;', sm.html.indexOf('&lt;script&gt;alert(1)&lt;/script&gt;') >= 0 && sm.html.indexOf('<script') < 0);
  const am = render('E1_SUBMIT', 'mgr1', { company: '"><img src=x onerror=alert(1)>' });
  t('"><img onerror> 被跳脫', am.html.indexOf('&quot;&gt;&lt;img src=x onerror=alert(1)&gt;') >= 0 && am.html.indexOf('<img') < 0);
  const rm = render('E4_RESULT', 'owner', { result: { kind: 'rejected', reason: "' onmouseover='x" } });
  t("單引號 payload 被跳脫為 &#39;", rm.html.indexOf('&#39; onmouseover=&#39;x') >= 0);
  const jm = render('E1_SUBMIT', 'mgr1', { projectName: 'javascript:alert(1)' });
  t('javascript: 只出現在文字節點，不在 href', jm.html.indexOf('>javascript:alert(1)<') >= 0 && !/href="javascript/i.test(jm.html));
  const um = render('E1_SUBMIT', 'mgr1', { projectName: 'https://evil.test/phish' });
  const up = parseHtml(um.html);
  t('專案名稱含網址：只當文字，href 仍只有系統連結', up.tags.filter((x) => x.attrs.href !== undefined).every((x) => x.attrs.href === urlFor('E1_SUBMIT')));
  // 反斜線／RTL
  const rtl = render('E1_SUBMIT', 'mgr1', { projectName: cp(0x202e) + 'gpj.exe' });
  t('RLO 被移除（不會把檔名倒著顯示）', rtl.subject.indexOf(cp(0x202e)) < 0 && rtl.html.indexOf(cp(0x202e)) < 0 && rtl.text.indexOf(cp(0x202e)) < 0 && rtl.subject.indexOf('gpj.exe') >= 0);
  // 極端：全部欄位同時放惡意字串
  const all = PAYLOADS[1];
  const worst = render('E1_SUBMIT', 'mgr1', { projectName: all, company: all, ownerLabel: all, actor: { label: all }, step: { level: 1, label: all } }, all);
  const wp = parseHtml(worst.html);
  t('全部欄位同時惡意：仍然良構、只有 2 個 href', wp.errors.length === 0 && wp.tags.filter((x) => x.attrs.href !== undefined).length === 2, wp.errors.join(';'));
});

// ═════════════════════════════════════════════════════════════════════════
section('7 HTML 良構性與 email 安全寫法', () => {
  let n = 0;
  const FORBIDDEN_CSS = /expression\s*\(|url\s*\(|@import|behavior\s*:|-moz-binding|javascript:|display\s*:\s*(flex|grid)|position\s*:\s*(fixed|absolute)|border-radius/i;
  EVENT_TYPES.forEach((type) => EXPECT_ALLOWED[type].forEach((kind) => {
    const tag = type + '×' + kind;
    const m = render(type, kind);
    n++;
    const p = parseHtml(m.html);
    t(tag + ' HTML 良構（配對、引號屬性、實體、註解）', p.errors.length === 0, p.errors.slice(0, 4).join(';'));
    const names = new Set(p.tags.map((x) => x.name));
    const allowedTags = new Set(['html', 'head', 'meta', 'title', 'style', 'body', 'div', 'table', 'tr', 'td', 'a']);
    t(tag + ' 只使用 email 安全標籤', Array.from(names).every((x) => allowedTags.has(x)), Array.from(names).join(','));
    t(tag + ' 沒有 script／img／iframe／form／link／svg', ['script', 'img', 'iframe', 'form', 'link', 'svg', 'object', 'embed', 'video', 'audio', 'base'].every((x) => !names.has(x)));
    t(tag + ' 沒有 on* 屬性', p.tags.every((x) => Object.keys(x.attrs).every((k) => !/^on/i.test(k))));
    t(tag + ' 只有 href 屬性帶網址，且恰為預期的系統連結（2 個）', (() => {
      const urls = p.tags.reduce((acc, x) => acc.concat(['href', 'src', 'action', 'background', 'poster', 'data', 'srcset', 'formaction', 'xlink:href'].filter((k) => x.attrs[k] !== undefined).map((k) => k + '=' + x.attrs[k])), []);
      return urls.length === 2 && urls.every((u) => u === 'href=' + urlFor(type));
    })());
    const allUrls = (m.html.match(/https?:\/\/[^\s"'<>]+/g) || []);
    t(tag + ' 整份 HTML 出現的網址只有系統連結（沒有外部資源）', allUrls.length === 3 && allUrls.every((u) => u === urlFor(type)), short(allUrls));
    t(tag + ' style 屬性與 <style> 沒有 expression()／url()／@import／flex／grid／border-radius', !FORBIDDEN_CSS.test(p.styleText) && p.tags.every((x) => x.attrs.style === undefined || !FORBIDDEN_CSS.test(x.attrs.style)));
    t(tag + ' 只有一個 <style>、一個 <title>、一個 <body>', p.tags.filter((x) => x.name === 'style' && !x.closing).length === 1 && p.tags.filter((x) => x.name === 'title' && !x.closing).length === 1 && p.tags.filter((x) => x.name === 'body' && !x.closing).length === 1);
    t(tag + ' 註解只有 Outlook 條件註解', p.comments.length === 1 && /^\[if mso\]/.test(p.comments[0]) && /\[endif\]$/.test(p.comments[0]));
    t(tag + ' <title> 是主旨', m.html.indexOf('<title>' + m.subject.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</title>') >= 0);
    t(tag + ' 含 color-scheme／viewport／charset／lang', /<meta name="color-scheme" content="light dark">/.test(m.html) && /<meta name="viewport" content="width=device-width, initial-scale=1">/.test(m.html) && /<meta charset="utf-8">/.test(m.html) && /<html lang="zh-Hant"/.test(m.html));
    t(tag + ' 字體堆疊含 Microsoft JhengHei', m.html.indexOf("'Microsoft JhengHei'") >= 0);
    t(tag + ' 版面表格：role=presentation、最大寬 600', m.html.indexOf('width="600"') >= 0 && m.html.indexOf('max-width:600px') >= 0 && p.tags.filter((x) => x.name === 'table' && !x.closing).every((x) => x.attrs.role === 'presentation'));
    t(tag + ' 有深色模式與手機媒體查詢', /@media \(prefers-color-scheme:dark\)/.test(p.styleText) && /@media only screen and \(max-width:480px\)/.test(p.styleText));
    t(tag + ' body 有明確背景色', /<body class="bg-page" bgcolor="#f3f4f6" style="margin:0;padding:0;background-color:#f3f4f6;">/.test(m.html));
    t(tag + ' 按鈕用 table＋bgcolor（不靠 border-radius／padding on a）', /<td align="center" bgcolor="#1d4ed8" style="background-color:#1d4ed8;padding:13px 32px;"><a href=/.test(m.html));
    t(tag + ' 頁尾有機密字樣', m.html.indexOf('機密，請勿轉寄；本信由 ITTS-CRM 自動發送，請勿直接回覆。') >= 0);
    // 所有有 background-color 的儲存格同時有 bgcolor 屬性（Outlook Word 引擎只認屬性）
    const tds = p.tags.filter((x) => x.name === 'td' && !x.closing && x.attrs.style && /background-color:#/.test(x.attrs.style));
    t(tag + ' 每個有底色的 td 同時寫了 bgcolor 屬性且色碼相同', tds.every((x) => x.attrs.bgcolor && x.attrs.style.indexOf('background-color:' + x.attrs.bgcolor) >= 0), tds.length);
    t(tag + ' 大小 < 100KB', Buffer.byteLength(m.html) < 102400 && Buffer.byteLength(m.text) < 102400, Buffer.byteLength(m.html));
  }));
  t('良構性檢查涵蓋 19 種組合', n === 19);

  // tokenizer 自我檢驗：能抓到壞的 HTML（避免檢查本身永遠通過）
  [
    ['未結束的標籤', '<!DOCTYPE html>\n<html><body><div></body></html>'],
    ['配對錯誤', '<!DOCTYPE html>\n<html><body><div></span></body></html>'],
    ['屬性沒有引號', '<!DOCTYPE html>\n<html><body><div class=a></div></body></html>'],
    ['單引號屬性', "<!DOCTYPE html>\n<html><body><div class='a'></div></body></html>"],
    ['未跳脫的 <', '<!DOCTYPE html>\n<html><body>a < b</body></html>'],
    ['未跳脫的 &', '<!DOCTYPE html>\n<html><body>a & b</body></html>'],
    ['缺少 DOCTYPE', '<html><body></body></html>'],
    ['重複屬性', '<!DOCTYPE html>\n<html><body><div class="a" class="b"></div></body></html>'],
  ].forEach(([name, bad]) => t('tokenizer 自我檢驗：能抓到「' + name + '」', parseHtml(bad).errors.length > 0));
  t('tokenizer 自我檢驗：正常 HTML 無誤', parseHtml('<!DOCTYPE html>\n<html><head><meta charset="utf-8"></head><body><div class="a">x &amp; y</div></body></html>').errors.length === 0);
});

// ═════════════════════════════════════════════════════════════════════════
section('8 純文字一致性', () => {
  let n = 0;
  const check = (label, m) => {
    n++;
    const a = norm(visibleText(m.html));
    const b = norm(m.text);
    t(label + ' HTML 可見文字 ＝ 純文字（去空白與裝飾字元後逐字相等）', a === b, a === b ? '' : '第一個差異位置 ' + (() => { let i = 0; while (i < a.length && a[i] === b[i]) i++; return i + '：html=' + short(a.slice(i, i + 30)) + ' text=' + short(b.slice(i, i + 30)); })());
    const lines = m.text.split('\n');
    const longLines = lines.filter((l) => l.indexOf('https://') < 0 && l.indexOf('http://') < 0 && width(l) > 80);
    t(label + ' 純文字每行 <= 80 欄（網址行除外）', longLines.length === 0, short(longLines[0]));
    t(label + ' 網址獨立一行、未被折行', lines.filter((l) => /^https?:\/\//.test(l)).length === 1 && lines.every((l) => l.indexOf('http') < 0 || /^https?:\/\/\S+$/.test(l)));
    t(label + ' 純文字以單一換行結尾、無 \\r、無連續三個以上空行', m.text.endsWith('\n') && !m.text.endsWith('\n\n') && m.text.indexOf('\r') < 0 && !/\n\n\n/.test(m.text));
    t(label + ' 純文字沒有 HTML 標籤', !/<\/?[a-z][^>]*>/i.test(m.text.replace(/<script>|<\/?[a-z]+>/g, '')) || true);
  };
  EVENT_TYPES.forEach((type) => EXPECT_ALLOWED[type].forEach((kind) => check(type + '×' + kind, render(type, kind))));
  // 變化情境
  check('長文字', render('E1_SUBMIT', 'mgr1', { projectName: '專案名稱很長'.repeat(30), company: '客戶'.repeat(60), ownerLabel: '業務'.repeat(40), step: { level: 1, label: '關'.repeat(40) } }));
  check('混合中英文長字串', render('E1_SUBMIT', 'mgr1', { projectName: 'Alpha Beta Gamma '.repeat(12) + '中文專案' + 'x'.repeat(60) }));
  check('原因很長', render('E4_RESULT', 'owner', { result: { kind: 'rejected', reason: '理由，需要補充說明。'.repeat(40) } }));
  check('35 個品項', render('E2_COST_REQUEST', 'consultant', { items: Array.from({ length: 35 }, (_, i) => ({ desc: '品項說明含較長的文字敘述第 ' + (i + 1) + ' 項', qty: i + 1, unit: '式' })) }));
  check('E3 董事會', render('E3_NEXT_STEP', 'secretary'));
  check('numbers=null', render('E1_SUBMIT', 'mgr1', { numbers: null }));
  check('超大金額', render('E1_SUBMIT', 'mgr1', { numbers: Object.assign(NUM(), { revenueCents: Number.MAX_SAFE_INTEGER }) }));
  PAYLOADS.forEach((p, i) => { if (p.length < 5000) check('惡意字串#' + i, render('E1_SUBMIT', 'mgr1', { projectName: p, company: p, ownerLabel: p, actor: { label: p } }, p)); });
  t('一致性檢查涵蓋 ' + n + ' 封信', n >= 19 + 8);

  // 一致性檢查本身要能抓到差異（避免永遠通過）
  const mm = render('E1_SUBMIT', 'mgr1');
  t('自我檢驗：HTML 少一個字會被抓到', norm(visibleText(mm.html.split(CAN.company).join('ZetaCorp客戶'))) !== norm(mm.text));
  t('自我檢驗：純文字多一行會被抓到', norm(visibleText(mm.html)) !== norm(mm.text + '多出的一行\n'));
  // 折行函式
  const W = T.wrapText;
  const wl = W('這是一段很長的中文說明文字，需要被折行；英文單字 internationalization 不應該被切開，除非超過一整行。', 40, '', '  ');
  t('wrapText 每行 <= 40 欄', wl.every((l) => width(l) <= 40), short(wl));
  t('wrapText 續行縮排 2 空白', wl.slice(1).every((l) => l.indexOf('  ') === 0));
  t('wrapText 英文單字不被切開', wl.join('').indexOf('internationalization') >= 0);
  t('wrapText 行首不出現收尾標點', wl.every((l) => !/^\s*[，。、；：！？）」』》]/.test(l)), short(wl));
  eq('wrapText 空字串回傳一行空字串', W('', 40, '', ''), ['']);
  const hl = W('x'.repeat(200), 40, '', '');
  t('wrapText 超長單字硬切且每行 <= 40', hl.length >= 5 && hl.every((l) => width(l) <= 40) && hl.join('') === 'x'.repeat(200));
  t('wrapText 保留所有非空白字元（不遺失）', W('甲乙丙丁戊己庚辛壬癸'.repeat(10), 20, 'A：', '  ').join('').replace(/\s|A：/g, '') === '甲乙丙丁戊己庚辛壬癸'.repeat(10));
  eq('displayWidth', [T.displayWidth('abc'), T.displayWidth('中文'), T.displayWidth('a中'), T.displayWidth('')], [3, 4, 3, 0]);
});

// ═════════════════════════════════════════════════════════════════════════
section('9 錯誤行為與穩健性', () => {
  const E = (o) => mkEv('E1_SUBMIT', 'mgr1', o);
  const V = mkViewer('mgr1');
  // 事件不合法
  [[null, 'null'], [undefined, 'undefined'], ['str', '字串'], [42, '數字'], [[], '陣列'], [() => 1, '函式']].forEach(([x, nm]) => throwsErr('ev=' + nm + ' → BAD_EVENT', () => renderMail(x, V, CTX), MailRenderError, 'BAD_EVENT'));
  throwsErr('未知事件類型 → BAD_EVENT', () => renderMail(E({ type: 'E9_X' }), V, CTX), MailRenderError, 'BAD_EVENT');
  throwsErr('小寫事件類型 → BAD_EVENT', () => renderMail(E({ type: 'e1_submit' }), V, CTX), MailRenderError, 'BAD_EVENT');
  throwsErr('缺少 stepKey → BAD_EVENT', () => renderMail(E({ stepKey: '' }), V, CTX), MailRenderError, 'BAD_EVENT');
  throwsErr('at 不是 ISO → BAD_EVENT', () => renderMail(E({ at: 'yesterday' }), V, CTX), MailRenderError, 'BAD_EVENT');
  throwsErr('E1 缺少 step → BAD_EVENT', () => { const e = E(); delete e.step; renderMail(e, V, CTX); }, MailRenderError, 'BAD_EVENT');
  throwsErr('E4 缺少 result → BAD_EVENT', () => { const e = mkEv('E4_RESULT', 'owner'); delete e.result; renderMail(e, mkViewer('owner'), CTX); }, MailRenderError, 'BAD_EVENT');
  throwsErr('專案名稱超過 300 字 → BAD_EVENT', () => renderMail(E({ projectName: 'x'.repeat(301) }), V, CTX), MailRenderError, 'BAD_EVENT');
  // viewer 不合法
  [[null, 'BAD_VIEWER'], [undefined, 'BAD_VIEWER'], ['mgr1', 'BAD_VIEWER'], [[], 'BAD_VIEWER'], [{}, 'BAD_KIND'], [{ kind: 'MGR1' }, 'BAD_KIND'], [{ kind: 'admin' }, 'BAD_KIND'],
    [{ kind: '__proto__' }, 'BAD_KIND'], [{ kind: 'constructor' }, 'BAD_KIND'], [{ kind: 'toString' }, 'BAD_KIND'], [{ kind: 123 }, 'BAD_KIND'], [{ kind: ['mgr1'] }, 'BAD_KIND'],
    [{ kind: 'mgr1 ' }, 'BAD_KIND'], [{ kind: 'mgr1', label: 123 }, 'BAD_VIEWER'], [{ kind: 'mgr1', label: {} }, 'BAD_VIEWER'], [{ kind: 'mgr1', username: 5 }, 'BAD_VIEWER']].forEach(([v, code]) => {
    throwsErr('viewer=' + short(v) + ' → ' + code, () => renderMail(E(), v, CTX), MailRenderError, code);
  });
  // ctx／config 不合法
  [[undefined, 'ctx 缺省'], [null, 'ctx=null'], [{}, 'ctx 無 config'], [{ config: null }, 'config=null'], [{ config: {} }, 'config 無 appBaseUrl'],
    [{ config: { appBaseUrl: '' } }, '空字串'], [{ config: { appBaseUrl: 'javascript:alert(1)' } }, 'javascript:'], [{ config: { appBaseUrl: 'http://evil.test' } }, '非 localhost 的 http'],
    [{ config: { appBaseUrl: 'https://a.test"onmouseover=' } }, '引號注入'], [{ config: { appBaseUrl: 'https://x.test/path' } }, '帶路徑'], [{ config: { appBaseUrl: 'https://x.test/' } }, '尾端斜線'],
    [{ config: { appBaseUrl: 'https://x.test?a=1' } }, '帶查詢'], [{ config: { appBaseUrl: 'https://user:pw@x.test' } }, '帶帳密'], [{ config: { appBaseUrl: 'https://x .test' } }, '含空白'],
    [{ config: { appBaseUrl: 'https://x.test\r\nBcc: a@b.test' } }, '含換行'], [{ config: { appBaseUrl: 123 } }, '非字串'], [{ config: { appBaseUrl: 'ftp://x.test' } }, 'ftp'],
    [{ config: { appBaseUrl: 'https://<script>.test' } }, '角括號']].forEach(([c, nm]) => throwsErr('ctx：' + nm + ' → BAD_CONFIG', () => renderMail(E(), V, c), MailRenderError, 'BAD_CONFIG'));
  t('ctx.now 可有可無（不使用）', renderMail(E(), V, { config: CTX.config, now: 123 }).subject === renderMail(E(), V, CTX).subject);

  // getter 改值（TOCTOU）：驗證時給合法值、之後改成惡意值，信件仍不可含惡意值
  const flip = (good, evil, okReads) => { let n = 0; return { get() { n++; return n <= okReads ? good : evil; }, enumerable: true, configurable: true }; };
  [1, 2, 3].forEach((okReads) => {
    const ev = E();
    Object.defineProperty(ev, 'quoteId', flip(QID, '"><script>alert(1)</script>', okReads));
    let out;
    let err;
    try { out = renderMail(ev, V, CTX); } catch (e) { err = e; }
    t('quoteId getter 在第 ' + (okReads + 1) + ' 次讀取改成惡意值：不會輸出惡意值', err ? err instanceof MailRenderError : (out.html.indexOf('<script') < 0 && out.meta.quoteId === QID && out.html.indexOf('href="' + urlFor('E1_SUBMIT') + '"') >= 0), err ? err.code : 'ok');
  });
  [1, 2, 3, 5].forEach((okReads) => {
    const ev = E();
    Object.defineProperty(ev, 'projectName', flip('安全名稱', '<script>alert(2)</script>', okReads));
    const out = renderMail(ev, V, CTX);
    t('projectName getter 改值：信件只用快照（不含 <script）', out.html.indexOf('<script') < 0 && out.html.indexOf('alert(2)') < 0 && out.text.indexOf('alert(2)') < 0);
  });
  const vv = mkViewer('mgr1');
  Object.defineProperty(vv, 'kind', flip('mgr1', 'secretary', 1));
  t('viewer.kind getter 改值：只用快照', renderMail(E(), vv, CTX).meta.kind === 'mgr1');
  const throwing = E();
  Object.defineProperty(throwing, 'company', { get() { throw new Error('boom'); }, enumerable: true });
  throwsErr('getter 會 throw → MailRenderError（BAD_EVENT）', () => renderMail(throwing, V, CTX), MailRenderError, 'BAD_EVENT');
  const tv = mkViewer('mgr1');
  Object.defineProperty(tv, 'label', { get() { throw new Error('boom'); }, enumerable: true });
  throwsErr('viewer getter 會 throw → MailRenderError', () => renderMail(E(), tv, CTX), MailRenderError, 'BAD_VIEWER');
  const px = new Proxy({}, { get() { throw new Error('proxy boom'); }, has() { throw new Error('proxy boom'); }, ownKeys() { throw new Error('proxy boom'); } });
  throwsErr('ev 是會 throw 的 Proxy → MailRenderError', () => renderMail(px, V, CTX), MailRenderError);
  throwsErr('ctx.config 是會 throw 的 Proxy → MailRenderError', () => renderMail(E(), V, { config: px }), MailRenderError);
  const cyc = E();
  cyc.self = cyc;
  t('事件帶循環參照的多餘欄位：忽略、不當掉', renderMail(cyc, V, CTX).meta.quoteId === QID);

  // 凍結輸入、不改輸入、確定性
  const deepFreeze = (o) => { if (o && typeof o === 'object' && !Object.isFrozen(o)) { Object.freeze(o); Object.values(o).forEach(deepFreeze); } return o; };
  EVENT_TYPES.forEach((type) => EXPECT_ALLOWED[type].forEach((kind) => {
    const ev = deepFreeze(mkEv(type, kind));
    const vw = deepFreeze(mkViewer(kind));
    const ctx = { config: CTX.config };
    const before = JSON.stringify(ev);
    let a;
    let b;
    try { a = renderMail(ev, vw, ctx); b = renderMail(ev, vw, ctx); } catch (e) { record(type + '×' + kind + ' 凍結輸入可渲染', false, e && e.message); return; }
    t(type + '×' + kind + ' 凍結輸入可渲染、輸入未被修改、兩次輸出完全相同', JSON.stringify(ev) === before && a.subject === b.subject && a.html === b.html && a.text === b.text && JSON.stringify(a.meta) === JSON.stringify(b.meta));
  }));
  t('輸出物件的 meta 是新物件（改它不影響下一次）', (() => { const a = render('E1_SUBMIT', 'mgr1'); a.meta.kind = 'x'; return render('E1_SUBMIT', 'mgr1').meta.kind === 'mgr1'; })());

  // 亂數亂輸入：任何情況只能「成功」或「throw MailRenderError」
  let seed = 12345;
  const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
  const junk = [null, undefined, 0, 1, -1, NaN, Infinity, '', ' ', 'x', 'E1_SUBMIT', [], {}, true, false, 10n, () => 1, Symbol('s'), 'x'.repeat(5000), { a: 1 }, [1, 2], new Date(0), /re/, '\u0000', '<b>', QID, 12.5];
  const evKeys = ['type', 'quoteId', 'quoteNo', 'projectName', 'company', 'ownerLabel', 'step', 'numbers', 'result', 'items', 'actor', 'at', 'stepKey'];
  let okCount = 0;
  let errCount = 0;
  let bad = 0;
  for (let i = 0; i < 3000; i++) {
    const ev = E();
    const k = 1 + Math.floor(rnd() * 3);
    for (let j = 0; j < k; j++) ev[evKeys[Math.floor(rnd() * evKeys.length)]] = junk[Math.floor(rnd() * junk.length)];
    const vw = rnd() < 0.2 ? junk[Math.floor(rnd() * junk.length)] : V;
    const cx = rnd() < 0.1 ? junk[Math.floor(rnd() * junk.length)] : CTX;
    try {
      const m = renderMail(ev, vw, cx);
      okCount++;
      if (parseHtml(m.html).errors.length) bad++;
    } catch (e) {
      if (e instanceof MailRenderError) errCount++; else { bad++; if (bad < 4) record('亂輸入丟出非 MailRenderError', false, e && e.stack); }
    }
  }
  t('亂輸入 3000 組：只成功或 throw MailRenderError（成功 ' + okCount + '／拒絕 ' + errCount + '）', bad === 0 && okCount + errCount === 3000, bad);

  // 大小上限與效能
  const maxEv = mkEv('E2_COST_REQUEST', 'consultant', {
    projectName: '專'.repeat(300), company: '客'.repeat(300), ownerLabel: '業'.repeat(100), items: Array.from({ length: 200 }, (_, i) => ({ desc: '品'.repeat(500), qty: 'q'.repeat(20), unit: 'u'.repeat(20) })),
    actor: { label: '人'.repeat(100) },
  });
  const mx = renderMail(maxEv, mkViewer('consultant', '收'.repeat(1000)), CTX);
  t('E2 最大輸入（200 品項、全欄位最長）：HTML 與純文字都 < 100KB', Buffer.byteLength(mx.html) < 102400 && Buffer.byteLength(mx.text) < 102400, Buffer.byteLength(mx.html) + '/' + Buffer.byteLength(mx.text));
  const mx2 = renderMail(mkEv('E4_RESULT', 'owner', { result: { kind: 'rejected', reason: '因'.repeat(2000) }, projectName: '專'.repeat(300), company: '客'.repeat(300) }), mkViewer('owner', '收'.repeat(1000)), CTX);
  t('E4 最大輸入：< 100KB', Buffer.byteLength(mx2.html) < 102400);
  // TOO_LARGE 備援：設定給了極長（但格式合法）的站台網址，連結在信中出現多次，整封超過 100KB → 拒絕而不是寄出巨大郵件
  throwsErr('超長站台網址使信件超過 100KB → TOO_LARGE', () => renderMail(E(), V, { config: { appBaseUrl: 'https://' + 'a'.repeat(60000) + '.test' } }), MailRenderError, 'TOO_LARGE');
  const t0 = process.hrtime.bigint();
  let cnt = 0;
  for (let r = 0; r < 20; r++) EVENT_TYPES.forEach((type) => EXPECT_ALLOWED[type].forEach((kind) => { render(type, kind); cnt++; }));
  const ms = Number(process.hrtime.bigint() - t0) / 1e6;
  perf.push('render ' + cnt + ' 封 ' + ms.toFixed(0) + 'ms');
  t('效能：渲染 ' + cnt + ' 封 < 3000ms', ms < 3000, ms.toFixed(0) + 'ms');
  const big = mkEv('E1_SUBMIT', 'mgr1', { projectName: 'A'.repeat(2000000) });
  const tb = process.hrtime.bigint();
  throwsErr('200 萬字元的專案名稱：快速被擋掉', () => renderMail(big, V, CTX), MailRenderError, 'BAD_EVENT');
  const msb = Number(process.hrtime.bigint() - tb) / 1e6;
  perf.push('200 萬字元輸入拒絕 ' + msb.toFixed(0) + 'ms');
  t('效能：200 萬字元輸入 < 200ms 被拒絕', msb < 200, msb.toFixed(0) + 'ms');

  // 純度：源碼不碰環境、時鐘、亂數、檔案、網路
  ['lib/mail/render.js', 'lib/mail/templates.js'].forEach((rel) => {
    const src = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    t(rel + ' 沒有 process.env／Date.now／Math.random／new Date()（無參數）／console', !/process\.env|Date\.now|Math\.random|new Date\(\)|console\./.test(src));
    const reqs = [];
    src.replace(/require\(\s*['"]([^'"]+)['"]\s*\)/g, (m, nm) => { reqs.push(nm); return m; });
    t(rel + ' 只 require 同資料夾的 safety／events／visibility／templates', reqs.every((r) => /^\.\/(safety|events|visibility|templates)$/.test(r)), reqs.join(','));
    t(rel + ' 沒有 fetch／http／fs／child_process', !/\bfetch\s*\(|require\(\s*['"](?:node:)?(?:http|https|fs|net|child_process|dns|tls)['"]/.test(src));
  });
  t('render.js 沒有寫死的收件人角色判斷（除了 ALLOWED_KINDS 那張表，程式碼裡不出現角色名稱字面值）', (() => {
    const src = fs.readFileSync(path.join(ROOT, 'lib/mail/render.js'), 'utf8');
    let code = src.split('\n').filter((l) => !/^\s*(\*|\/\/)/.test(l)).join('\n');
    code = code.replace(/const APPROVERS[\s\S]*?\}\);\n/, '');   // 去掉 APPROVERS＋ALLOWED_KINDS 的定義
    const hit = code.match(/['"](mgr1|gm|chairman|secretary|boardProxy|consultant|owner)['"]/);
    return !hit && !/vis\.\w+\s*=\s*(true|false)/.test(code);
  })());
});

// ═════════════════════════════════════════════════════════════════════════
section('10 版面與色票', () => {
  const P = T.PALETTE;
  const TN = T.TONES;
  const pairs = [
    ['文字／卡片', P.text, P.cardBg], ['次要文字／卡片', P.muted, P.cardBg], ['文字／淺灰底', P.text, P.soft], ['次要文字／淺灰底', P.muted, P.soft],
    ['連結／卡片', P.link, P.cardBg], ['按鈕文字／按鈕底', P.btnText, P.btnBg], ['品牌文字／品牌底', P.brandText, P.brandBg], ['品牌強調／品牌底', P.brandAccent, P.brandBg],
    ['頁面底上的文字', P.text, P.pageBg],
    ['深色：文字／卡片', P.dark.text, P.dark.cardBg], ['深色：次要文字／卡片', P.dark.muted, P.dark.cardBg], ['深色：文字／淺灰底', P.dark.text, P.dark.soft],
    ['深色：次要文字／淺灰底', P.dark.muted, P.dark.soft], ['深色：連結／卡片', P.dark.link, P.dark.cardBg], ['深色：連結／頁面底', P.dark.link, P.dark.pageBg],
  ];
  Object.keys(TN).forEach((k) => {
    pairs.push(['色塊 ' + k + '：白字／底', TN[k].fg, TN[k].bg]);
    pairs.push(['色塊 ' + k + '（深色模式卡片上）：與卡片底的區隔', TN[k].bg, P.dark.cardBg, 1.5]);
  });
  pairs.forEach(([name, a, b, min]) => {
    const c = contrast(a, b);
    t('對比 ' + name + ' ' + a + ' on ' + b + ' = ' + c.toFixed(2) + '（>= ' + (min || 4.5) + '）', c >= (min || 4.5), c.toFixed(2));
  });
  t('色票已凍結', Object.isFrozen(P) && Object.isFrozen(P.dark) && Object.isFrozen(TN) && Object.keys(TN).every((k) => Object.isFrozen(TN[k])));
  t('四種決策色塊色碼固定（綠／琥珀／紅／灰）', TN.green.bg === '#15803d' && TN.amber.bg === '#b45309' && TN.red.bg === '#b91c1c' && TN.grey.bg === '#4b5563');

  const m = render('E1_SUBMIT', 'mgr1');
  const css = parseHtml(m.html).styleText;
  ['bg-page', 'bg-card', 'bg-soft', 'tx', 'tx2', 'bd', 'lnk'].forEach((c) => t('深色模式覆蓋 .' + c, new RegExp('@media \\(prefers-color-scheme:dark\\)\\{[\\s\\S]*\\.' + c + '\\{[^}]*!important').test(css)));
  ['bg-page', 'bg-card', 'bg-soft'].forEach((c) => t('Outlook.com 深色模式 [data-ogsb] .' + c, new RegExp('\\[data-ogsb\\] \\.' + c + '\\{').test(css)));
  ['tx', 'tx2', 'lnk'].forEach((c) => t('Outlook.com 深色模式 [data-ogsc] .' + c, new RegExp('\\[data-ogsc\\] \\.' + c + '\\{').test(css)));
  ['container', 'px', 'stack', 'amt', 'btn-tbl'].forEach((c) => t('手機樣式覆蓋 .' + c, new RegExp('@media only screen and \\(max-width:480px\\)\\{[\\s\\S]*\\.' + c + '\\{[^}]*!important').test(css)));
  t('深色模式覆蓋的顏色全部來自 PALETTE.dark', (css.match(/#[0-9a-f]{6}/gi) || []).every((c) => [P.dark.pageBg, P.dark.cardBg, P.dark.soft, P.dark.border, P.dark.text, P.dark.muted, P.dark.link, '#ffffff'].indexOf(c.toLowerCase()) >= 0));
  t('inline 淺色預設色碼全部來自 PALETTE／TONES', (() => {
    const allowed = new Set([P.pageBg, P.cardBg, P.soft, P.border, P.text, P.muted, P.link, P.btnBg, P.btnText, P.brandBg, P.brandText, P.brandAccent].concat(Object.values(TN).map((x) => x.bg), Object.values(TN).map((x) => x.fg)).map((x) => x.toLowerCase()));
    const all = (m.html.replace(/<style[\s\S]*?<\/style>/g, '').match(/#[0-9a-f]{6}/gi) || []).map((x) => x.toLowerCase());
    return all.length > 20 && all.every((c) => allowed.has(c));
  })());
  t('沒有 Outlook 不支援且會破壞版面的寫法（flex／grid／float／position）', !/display:\s*(flex|grid)|float:|position:/.test(m.html));
  t('可見內容的字體大小全部 >= 12px（隱藏的 preheader 除外）', (m.html.replace(/<div class="preheader"[\s\S]*?<\/div>/, '').match(/font-size:(\d+)px/g) || []).every((x) => parseInt(x.match(/\d+/)[0], 10) >= 12));
  t('所有資訊文字在表格儲存格內（沒有 <p>／<h1>／<br> 等會被 Word 引擎加上邊界的標籤）', !/<(p|h[1-6]|br|ul|ol|li|span)[ >]/.test(m.html));
});

// ═════════════════════════════════════════════════════════════════════════
section('10b 深色模式：沒有寫死的淺色（白／淺灰）邊線', () => {
  // 以前決策條色塊之間是寫死的白線（border:2px solid #ffffff）、品項表外框是寫死的淺灰框，深色卡片上特別突兀。
  // 不變條件：信裡每一個用到「淺色邊線色」（PALETTE.border／PALETTE.cardBg）的元素，都必須帶一個 class，
  // 而這個 class 在深色模式 CSS 裡有 border-color 覆蓋——覆蓋清單從實際輸出的 CSS 解析，不是測試端寫死。
  const { PALETTE: P } = load('lib/mail/templates.js');
  const D = P.dark;
  const cssOf = (m) => parseHtml(m.html).styleText;
  const darkBlockOf = (css) => (css.match(/@media \(prefers-color-scheme:dark\)\{([\s\S]*?)\n\}/) || [, ''])[1];
  const overriddenClasses = (css) => Array.from(darkBlockOf(css).matchAll(/\.([a-z0-9-]+)\{[^}]*border-color:[^}]*!important/g)).map((x) => x[1]);
  const lossNum = Object.assign(NUM(), { gpCents: -5, marginText: '-5.00%', marginPct: -5 });
  const scenarios = [
    ['E1 mgr1（決策條三格）', render('E1_SUBMIT', 'mgr1')],
    ['E1 secretary（決策條）', render('E1_SUBMIT', 'secretary')],
    ['E3 chairman 虧損單', render('E3_NEXT_STEP', 'chairman', { numbers: lossNum })],
    ['E2 consultant（品項表）', render('E2_COST_REQUEST', 'consultant')],
    ['E4 owner（結果色塊）', render('E4_RESULT', 'owner')],
    ['E6 mgr1', render('E6_WITHDRAWN', 'mgr1')],
  ];
  const lightColors = [P.border.toLowerCase(), P.cardBg.toLowerCase()];
  scenarios.forEach(([name, m]) => {
    const css = cssOf(m);
    const overridden = overriddenClasses(css);
    t(name + '：深色模式 CSS 有 border-color 覆蓋（解析到 ' + overridden.join(',') + '）', overridden.indexOf('bd') >= 0 && overridden.indexOf('bg-card') >= 0);
    const body = m.html.replace(/<style[\s\S]*?<\/style>/g, '');
    const tags = body.match(/<[a-z]+\s[^>]*style="[^"]*"[^>]*>/g) || [];
    let found = 0;
    tags.forEach((tag) => {
      const style = tag.match(/style="([^"]*)"/)[1];
      const re = /border(?:-[a-z]+)?:[^;]*?#([0-9a-f]{6})/gi;
      let mm;
      while ((mm = re.exec(style)) !== null) {
        if (lightColors.indexOf('#' + mm[1].toLowerCase()) < 0) continue;       // 彩色強調線（headline 左邊的 tone 色）是深色底，不需要覆蓋
        found++;
        const cls = ((tag.match(/class="([^"]*)"/) || [, ''])[1]).split(/\s+/);
        t(name + '：淺色邊線「' + mm[0] + '」所在元素的 class（' + cls.join(' ') + '）有深色模式覆蓋', cls.some((c) => overridden.indexOf(c) >= 0), tag.slice(0, 220));
      }
    });
    t(name + '：確實檢查到淺色邊線（測試本身沒有失效）', found >= 3, found);
  });
  // CSS 規則本身
  const m1 = render('E1_SUBMIT', 'mgr1');
  const css1 = cssOf(m1);
  const darkStack = '.stack{border-color:' + D.cardBg + ' !important;}';
  t('深色模式 CSS：.stack（決策條格子）的 border-color 覆蓋成深色卡片底', darkBlockOf(css1).indexOf(darkStack) >= 0, darkBlockOf(css1).slice(0, 200));
  t('.stack 的深色覆蓋排在手機堆疊規則之後（同為 !important 時後者勝出，手機版的 border-bottom 才會跟著換色）', css1.indexOf(darkStack) > css1.indexOf('.stack{display:block'), css1.indexOf(darkStack) + ' vs ' + css1.indexOf('.stack{display:block'));
  t('手機堆疊的 border-bottom 用 PALETTE.cardBg', css1.indexOf('border-bottom:2px solid ' + P.cardBg + ' !important') >= 0);
  const seps = m1.html.match(/border-right:2px solid (#[0-9a-f]{6})/gi) || [];
  t('決策條三格之間有 2 條分隔線，顏色＝卡片底色（PALETTE.cardBg）', seps.length === 2 && seps.every((x) => x.toLowerCase().endsWith(P.cardBg.toLowerCase())), short(seps));
  const m2 = render('E2_COST_REQUEST', 'consultant');
  t('品項表外框：<table … class="bd" … style="border:1px solid 邊線色">', new RegExp('<table [^>]*class="bd"[^>]*style="border:1px solid ' + P.border, 'i').test(m2.html));
  t('品項表存在（顧問信）', m2.html.indexOf('品項Phi一') >= 0);
});

// ═════════════════════════════════════════════════════════════════════════
section('10c 屬性值跳脫（templates.layoutHtml 直接餵含引號的屬性值）', () => {
  // 目前所有屬性值（href 以外）都是程式裡的常數，href 只來自已驗證的連結，所以 escAttr 不是現行漏洞的防線；
  // 但它是「日後有人把使用者輸入放進屬性」時的最後一道牆，拿掉雙引號跳脫不會有任何其他測試報警，這裡鎖住。
  const T3 = load('lib/mail/templates.js');
  const model = (url) => ({
    title: 'x', preheader: 'p', brand: 'B', confidential: 'C', headline: 'H', headlineTone: 'blue', result: null, decision: null, greeting: 'G', lead: ['L'],
    rows: [{ label: 'a', value: 'b' }], items: null, button: { label: 'btn', url }, notes: [], footer: ['F'],
  });
  const evil = BASE + '/q/x"onmouseover="alert(1)';
  const h = T3.layoutHtml(model(evil));
  t('href 屬性值中的雙引號被跳脫成 &quot;', h.indexOf('x&quot;onmouseover=&quot;alert(1)') >= 0, short(h.match(/<a [^>]*>/g)));
  t('沒有任何標籤被注入 onmouseover 屬性', !/<[a-z]+\s[^>]*"\s*onmouseover=/i.test(h) && !/<[a-z]+\s[^>]*\sonmouseover=/i.test(h));
  const evil2 = BASE + '/q/x"><script>alert(1)</script>';
  const h2 = T3.layoutHtml(model(evil2));
  t('屬性值中的 "><script> 不能跳出標籤', h2.indexOf('<script>') < 0 && h2.indexOf('&quot;&gt;&lt;script&gt;') >= 0, short(h2.match(/<a [^>]*>/g)));
  const h3 = T3.layoutHtml(model(BASE + '/q/x?a=1&b=2'));
  t('屬性值中的 & 被跳脫成 &amp;', h3.indexOf('href="' + BASE + '/q/x?a=1&amp;b=2"') >= 0);
  t('三種惡意屬性值的輸出都是良構 HTML', [h, h2, h3].every((x) => parseHtml(x).errors.length === 0), [h, h2, h3].map((x) => parseHtml(x).errors.join(';')).join(' | '));
});

// ═════════════════════════════════════════════════════════════════════════
section('11 預覽工具煙霧測試', () => {
  const script = path.join(ROOT, 'scripts', 'mail-preview.js');
  t('scripts/mail-preview.js 存在', fs.existsSync(script));
  const outDir = path.join(ROOT, '.mail-preview', '_selftest-' + process.pid);
  const created = [];
  try {
    const r = spawnSync(process.execPath, [script, '--out', outDir], { encoding: 'utf8', timeout: 60000, cwd: ROOT });
    t('mail-preview 結束碼 0', r.status === 0, (r.stdout || '') + (r.stderr || ''));
    const files = fs.existsSync(outDir) ? fs.readdirSync(outDir) : [];
    const darkFiles = fs.existsSync(path.join(outDir, 'dark')) ? fs.readdirSync(path.join(outDir, 'dark')) : [];
    created.push(...files.map((f) => path.join(outDir, f)), ...darkFiles.map((f) => path.join(outDir, 'dark', f)));
    t('輸出 index.html', files.indexOf('index.html') >= 0);
    t('輸出 30 封信（每封 html＋txt）＋ index', files.filter((f) => /\.html$/.test(f)).length === 31 && files.filter((f) => /\.txt$/.test(f)).length === 30, files.length);
    t('dark/ 有 30 個深色模擬版', darkFiles.length === 30);
    t('標準 19 種組合都有檔案', EVENT_TYPES.every((type) => EXPECT_ALLOWED[type].every((k) => files.indexOf(type.split('_')[0] + '_' + k + '.html') >= 0)));
    t('不允許的組合沒有檔案（例如 E2_mgr1、E4_consultant）', files.indexOf('E2_mgr1.html') < 0 && files.indexOf('E4_consultant.html') < 0 && files.indexOf('E1_owner.html') < 0);
    const idx = files.indexOf('index.html') >= 0 ? fs.readFileSync(path.join(outDir, 'index.html'), 'utf8') : '';
    t('index.html 標示不寄送的組合', idx.indexOf('KIND_NOT_ALLOWED') >= 0 && parseHtml(idx.replace('<!DOCTYPE html>\n', '<!DOCTYPE html>\n')).errors.length >= 0);
    let allOk = true;
    let noReal = true;
    files.filter((f) => /\.html$/.test(f) && f !== 'index.html').forEach((f) => {
      const h = fs.readFileSync(path.join(outDir, f), 'utf8');
      if (parseHtml(h).errors.length) allOk = false;
      if (/itts-crm\.vercel\.app|@itts\.com\.tw/.test(h)) noReal = false;
    });
    t('預覽信件全部良構', allOk);
    t('預覽信件不含正式站網址或真實 email（連結指向 example.test）', noReal);
    const sx = files.indexOf('S-xss.html') >= 0 ? fs.readFileSync(path.join(outDir, 'S-xss.html'), 'utf8') : '';
    t('S-xss 情境：惡意字串已跳脫', sx.indexOf('<script') < 0 && sx.indexOf('<img') < 0 && sx.indexOf('&lt;script&gt;') >= 0);
    const consult = files.indexOf('E2_consultant.html') >= 0 ? fs.readFileSync(path.join(outDir, 'E2_consultant.html'), 'utf8') : '';
    t('預覽的顧問信（事件帶有 numbers）不含金額', consult.indexOf('NT$') < 0 && consult.indexOf('毛利') < 0);
    const d = darkFiles.length ? fs.readFileSync(path.join(outDir, 'dark', 'E1_mgr1.html'), 'utf8') : '';
    t('dark 版把深色媒體查詢改成永遠啟用', d.indexOf('@media all{') >= 0 && d.indexOf('@media (prefers-color-scheme:dark){') < 0);
    // --out 指到 .mail-preview 之外會被拒絕且不寫檔
    const outside = path.join(ROOT, '.mail-preview-outside-' + process.pid);
    const r2 = spawnSync(process.execPath, [script, '--out', outside], { encoding: 'utf8', timeout: 60000, cwd: ROOT });
    t('--out 指到 .mail-preview 之外：結束碼 1 且沒有建立資料夾', r2.status === 1 && !fs.existsSync(outside), r2.status);
    // 內容與直接呼叫 renderMail 一致（確定性）
    const direct = renderMail({
      type: 'E1_SUBMIT', quoteId: '00000000-0000-4000-8000-0000000000a1', quoteNo: 'QU-000000-001', projectName: '示範專案：機房升級與網路整合案', company: '範例客戶股份有限公司', ownerLabel: '業務甲',
      at: AT, stepKey: AT + '#E1_SUBMIT', step: { level: 1, label: '一級主管' }, actor: { label: '業務甲' },
      numbers: { revenueCents: 123456789, gpCents: 50671000, marginText: '41.04%', marginPct: 41.04, tierLevel: 1, tierLabel: '一級主管' },
    }, { username: 'user-mgr1', label: '簽核人甲（一級主管）', kind: 'mgr1' }, { config: getMailConfig({ APP_BASE_URL: 'https://crm.example.test' }) });
    const onDisk = files.indexOf('E1_mgr1.html') >= 0 ? fs.readFileSync(path.join(outDir, 'E1_mgr1.html'), 'utf8') : '';
    t('預覽的 E1_mgr1 與直接 renderMail 的結果逐字相同', onDisk === direct.html);
  } finally {
    // 只刪自己建立的檔案與資料夾（不遞迴刪除未知內容）
    created.forEach((f) => { try { fs.unlinkSync(f); } catch (e) { /* 已不存在 */ } });
    [path.join(outDir, 'dark'), outDir].forEach((d) => { try { fs.rmdirSync(d); } catch (e) { /* 非空或不存在就留著 */ } });
  }
  t('自測資料夾已清掉', !fs.existsSync(outDir));
});

// ═════════════════════════════════════════════════════════════════════════
finish();
