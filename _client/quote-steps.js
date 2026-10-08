// ════════════════════════════════════════════════════════════
// ── 新增報價單「逐步填寫」(quote-steps.js) ───────────────────
// 只有「新增」報價單（openQuoteModal(null)）才會分步；編輯既有單（有 id）完全不碰：不分步、不收合、不隱藏分頁。
// 這一層只管「呈現」：未到的步驟區塊先隱藏、已完成的收成一行摘要，所有欄位仍在 DOM 裡，
// quote.js 的 saveQuote／readQuoteItems 等照常運作（資料收集與送出邏輯沒有改）。
// 步驟：客戶 → 專案 → 商品 → 指派顧問（只有勾選的商品需要顧問填成本時才出現）→ 報價項目 → 優惠 → 付款方式（含追加條款）→ 確認並儲存。
// 規則：
//   · 步驟「完成」＝使用者按過「下一步」（或開啟時預填的資料已滿足必填）且目前仍有效；第一個沒完成的步驟＝目前步驟，後面的不顯示。
//   · 已完成的步驟收成一行摘要＋「修改」；修改到變成不完整，後面的步驟自動收回；補好之後回到先前確認過的進度。
//   · 儲存按鈕在全部步驟完成前停用；「毛利與簽核路徑」分頁在走完之前隱藏。
// 掛鉤（quote.js 內各一行）：openQuoteModal → onOpen(q)、switchQuoteTab → onTab(name)、updateQuoteTotals／onQuoteProductsChanged → refresh()。
// 測試用關閉開關：只有 location.hostname 為 localhost 且 window.__qNoSteps === true 時才關閉分步（一般使用者不會碰到）。
// 所有使用者輸入的文字（公司名、專案名…）只用 textContent 顯示，不經 innerHTML。
// ════════════════════════════════════════════════════════════
var QSteps = (function () {
  'use strict';

  var CSS = `
/* 逐步填寫（新增報價單）。規則全部掛在 .qs-on／.qs-mode 底下：編輯既有單沒有這兩個 class，版面與改版前完全相同。 */
.qs-rail { display:none; }
.qs-rail-mini { display:none; }
#quoteTabContentInfo.qs-on { display:flex; align-items:flex-start; gap:22px; }
#quoteTabContentInfo.qs-on > .qs-main { flex:1 1 0; min-width:0; }
#quoteTabContentInfo.qs-on > .qs-rail { display:block; flex:0 0 148px; position:sticky; top:0; align-self:flex-start; }
#quoteModalOverlay > .modal.qs-mode.qs-info { max-width:1000px !important; }
#quoteModalOverlay > .modal.qs-mode #quoteModalSave:disabled { opacity:.45; cursor:not-allowed; }

.qs-rail-h { font-size:12px; color:#8a94a3; margin:2px 0 6px 8px; letter-spacing:.06em; }
.qs-rail-list { list-style:none; margin:0; padding:0; }
.qs-ri { display:flex; align-items:center; gap:8px; width:100%; text-align:left; border:none; background:none; padding:7px 8px; border-radius:8px;
  font-size:13px; line-height:1.4; color:#8a94a3; cursor:default; font-family:inherit; }
.qs-ri .m { flex:0 0 18px; text-align:center; font-size:13px; }
.qs-ri.done { color:#188038; cursor:pointer; }
.qs-ri.done:hover { background:#e6f4ea; }
.qs-ri.cur { color:#1a73e8; font-weight:700; background:#e8f0fe; cursor:pointer; }
.qs-ri.todo.go { cursor:pointer; }
.qs-ri.todo.go:hover { background:#f1f3f9; }
.qs-ri.rev { box-shadow:inset 0 0 0 1.5px #9ec5f4; }
.qs-ri:focus-visible { outline:2px solid #1a73e8; outline-offset:1px; }

.qs-on .qs-step { margin:0 0 12px; border:1px solid #e0e4ee; border-radius:10px; background:#fff; scroll-margin-top:6px; }
.qs-on .qs-step[data-st="future"], .qs-on .qs-step[data-st="na"] { display:none; }
.qs-on .qs-final > .qs-step[data-st="done"] { display:none; }   /* 全部完成、停在確認並儲存：清單已逐步驟列出，上面的摘要卡不重複 */
.qs-on .qs-step[data-st="done"] > :not(.qs-head), .qs-on .qs-step[data-st="pending"] > :not(.qs-head) { display:none !important; }
.qs-on .qs-step[data-st="done"] { background:#f8faf9; }
.qs-on .qs-step[data-st="pending"] { background:#fafbfe; border-style:dashed; }
.qs-on .qs-step[data-st="active"], .qs-on .qs-step[data-st="edit"], .qs-on .qs-step[data-st="review"] {
  border-color:#9ec5f4; box-shadow:0 0 0 3px rgba(26,115,232,.08); animation:qsIn .16s ease-out; }
.qs-on .qs-step > :not(.qs-head):not(.qs-nav) { margin-left:16px; margin-right:16px; }
.qs-on .quote-section-title { display:none; }
.qs-on .qs-hide { display:none; }
@keyframes qsIn { from { opacity:0; transform:translateY(4px); } to { opacity:1; transform:none; } }
@media (prefers-reduced-motion: reduce) { .qs-on .qs-step { animation:none !important; } }

.qs-head { display:flex; align-items:center; gap:10px; padding:11px 16px; box-sizing:border-box; }
.qs-badge { flex:0 0 24px; width:24px; height:24px; border-radius:50%; display:flex; align-items:center; justify-content:center;
  font-size:12.5px; font-weight:700; background:#e8f0fe; color:#1a73e8; }
.qs-step[data-st="done"] .qs-badge { background:#e6f4ea; color:#188038; }
.qs-step[data-st="pending"] .qs-badge { background:#eef1f6; color:#8a94a3; }
.qs-ht { flex:1 1 auto; min-width:0; }
.qs-l1 { display:flex; align-items:baseline; gap:10px; min-width:0; }
.qs-title { flex:0 0 auto; font-size:14px; font-weight:700; color:#1a2d52; white-space:nowrap; }
.qs-txt { min-width:0; font-size:13px; color:#444; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.qs-step[data-st="active"] .qs-txt, .qs-step[data-st="edit"] .qs-txt, .qs-step[data-st="review"] .qs-txt { color:#6b7686; white-space:normal; }
.qs-step[data-st="pending"] .qs-txt { color:#8a94a3; }
.qs-sub { margin-top:2px; font-size:12px; line-height:1.5; color:#8a94a3; }
.qs-edit { flex:0 0 auto; border:none; background:none; color:#1a73e8; font-size:13px; cursor:pointer; padding:4px 10px; border-radius:6px; font-family:inherit; }
.qs-edit:hover { background:#e8f0fe; }

.qs-nav { display:none; align-items:center; justify-content:space-between; gap:12px; margin:14px 16px; flex-wrap:wrap; }
.qs-step[data-st="active"] > .qs-nav, .qs-step[data-st="edit"] > .qs-nav { display:flex; }
.qs-hint { flex:1 1 200px; min-width:0; font-size:12.5px; line-height:1.5; color:#6b7686; }
.qs-hint.need { color:#b25e00; font-weight:600; }
.qs-next:disabled { opacity:.45; cursor:not-allowed; }

.qs-conf { padding:2px 0 14px; outline:none; }
.qs-conf-list { list-style:none; margin:0; padding:0; }
.qs-cr { display:flex; align-items:flex-start; gap:10px; padding:8px 0; border-bottom:1px solid #eef0f6; font-size:13px; line-height:1.5; }
.qs-cr .m { flex:0 0 20px; text-align:center; font-weight:700; }
.qs-cr.ok .m { color:#188038; }
.qs-cr.bad .m { color:#c62828; }
.qs-cr.todo .m { color:#8a94a3; }
.qs-cr .n { flex:0 0 92px; font-weight:600; color:#1a2d52; }
.qs-cr .s { flex:1 1 auto; min-width:0; color:#444; word-break:break-word; }
.qs-cr.bad .s { color:#c62828; }
.qs-cr.todo .s { color:#6b7686; }
.qs-cr .g { flex:0 0 auto; border:none; background:none; color:#1a73e8; font-size:12.5px; cursor:pointer; padding:2px 8px; border-radius:6px; font-family:inherit; }
.qs-cr .g:hover { background:#e8f0fe; }
.qs-conf-msg { margin-top:10px; padding:8px 12px; border-radius:8px; font-size:13px; line-height:1.6; }
.qs-conf-msg.ok { background:#e6f4ea; color:#188038; border:1px solid #b7dfc2; }
.qs-conf-msg.bad { background:#fff4e5; color:#8a4b00; border:1px solid #f5c98b; }
.qs-conf-note { margin-top:8px; font-size:12.5px; line-height:1.6; color:#6b7686; }

@media (max-width: 639px) {
  #quoteTabContentInfo.qs-on { flex-direction:column; align-items:stretch; gap:10px; }
  #quoteTabContentInfo.qs-on > .qs-rail { flex:none; align-self:stretch; position:sticky; top:0; z-index:3; background:#fff; padding:2px 0 6px; }
  .qs-rail-h, .qs-rail-list { display:none; }
  .qs-rail-mini { display:block; font-size:13px; font-weight:700; color:#1a73e8; background:#e8f0fe; border-radius:8px; padding:7px 12px; }
  .qs-head { padding:10px 12px; }
  .qs-on .qs-step > :not(.qs-head):not(.qs-nav) { margin-left:12px; margin-right:12px; }
  .qs-nav { margin:12px; }
  .qs-cr { flex-wrap:wrap; }
  .qs-cr .n { flex:0 0 auto; }
}

body.dark .qs-rail-h { color:#8b949e; }
body.dark .qs-ri { color:#8b949e; }
body.dark .qs-ri.done { color:#4ade80; }
body.dark .qs-ri.done:hover { background:#0a1f10; }
body.dark .qs-ri.cur { color:#8ecfff; background:#0d2040; }
body.dark .qs-ri.todo.go:hover { background:#21262d; }
body.dark .qs-ri.rev { box-shadow:inset 0 0 0 1.5px #1c3a5f; }
body.dark .qs-on .qs-step { background:#161b22; border-color:#30363d; }
body.dark .qs-on .qs-step[data-st="done"] { background:#12171e; }
body.dark .qs-on .qs-step[data-st="pending"] { background:#12171e; }
body.dark .qs-on .qs-step[data-st="active"], body.dark .qs-on .qs-step[data-st="edit"], body.dark .qs-on .qs-step[data-st="review"] {
  border-color:#1c3a5f; box-shadow:0 0 0 3px rgba(88,166,255,.12); }
body.dark .qs-badge { background:#0d2040; color:#8ecfff; }
body.dark .qs-step[data-st="done"] .qs-badge { background:#0a1f10; color:#4ade80; }
body.dark .qs-step[data-st="pending"] .qs-badge { background:#21262d; color:#8b949e; }
body.dark .qs-title { color:#e6edf3; }
body.dark .qs-txt { color:#c9d1d9; }
body.dark .qs-step[data-st="active"] .qs-txt, body.dark .qs-step[data-st="edit"] .qs-txt, body.dark .qs-step[data-st="review"] .qs-txt { color:#8b949e; }
body.dark .qs-step[data-st="pending"] .qs-txt, body.dark .qs-sub { color:#8b949e; }
body.dark .qs-edit, body.dark .qs-cr .g { color:#58a6ff; }
body.dark .qs-edit:hover, body.dark .qs-cr .g:hover { background:#0d2040; }
body.dark .qs-hint { color:#8b949e; }
body.dark .qs-hint.need { color:#d4a84e; }
body.dark .qs-cr { border-bottom-color:#21262d; }
body.dark .qs-cr .n { color:#e6edf3; }
body.dark .qs-cr .s { color:#c9d1d9; }
body.dark .qs-cr.ok .m { color:#4ade80; }
body.dark .qs-cr.bad .m, body.dark .qs-cr.bad .s { color:#ff8080; }
body.dark .qs-cr.todo .m, body.dark .qs-cr.todo .s { color:#8b949e; }
body.dark .qs-conf-msg.ok { background:#0a1f10; color:#4ade80; border-color:#1b5230; }
body.dark .qs-conf-msg.bad { background:#2a2000; color:#d4a84e; border-color:#5a4000; }
body.dark .qs-conf-note { color:#8b949e; }
body.dark .qs-rail-mini { background:#0d2040; color:#8ecfff; }
body.dark #quoteTabContentInfo.qs-on > .qs-rail { background:#161b22; }
`;

  // ── 小工具 ────────────────────────────────────────────────
  function el(id) { return document.getElementById(id); }
  function val(id) { var e = el(id); return e ? String(e.value == null ? '' : e.value).trim() : ''; }
  function mk(tag, cls, text) { var n = document.createElement(tag); if (cls) n.className = cls; if (text != null) n.textContent = text; return n; }
  function money(n) { return typeof fmtMoney === 'function' ? fmtMoney(n) : 'NT$ ' + Math.round(n).toLocaleString(); }
  function trunc(s, n) { var a = Array.from(String(s == null ? '' : s)); return a.length > n ? a.slice(0, n).join('') + '…' : a.join(''); }
  function modalEl() { return document.querySelector('#quoteModalOverlay > .modal'); }
  function flagOff() { try { return location.hostname === 'localhost' && window.__qNoSteps === true; } catch (e) { return false; } }
  function selectedProducts() { return typeof readSelectedProducts === 'function' ? readSelectedProducts() : []; }
  function needsConsultant() { return typeof quoteNeedsConsultant === 'function' && quoteNeedsConsultant(selectedProducts()); }
  function infoTabVisible() { var b = el('quoteTabContentInfo'); return !!b && b.style.display !== 'none'; }
  function pnlTabVisible() { var b = el('quoteTabContentPnl'); return !!b && b.style.display !== 'none'; }

  // ── 狀態 ──────────────────────────────────────────────────
  var active = false;        // 目前這次開啟是否在分步模式
  var confirmed = {};        // 步驟 id → 使用者按過「下一步」（或開啟時預填已滿足）
  var editing = null;        // 正在「修改」的步驟 id
  var review = false;        // 還沒走完時，從進度欄點「確認並儲存」檢視清單
  var hadNeed = false;       // 上一次算出來的「需要顧問」，用來偵測勾選變化
  var touched = new Set();   // 使用者動過單價欄的品項列（nid／lid）：單價 0 要使用者明確輸入過才算「有填」
  var els = {};              // id → { wrap, head, badge, title, txt, sub, edit, nav, hint, btn }
  var moves = [];            // 搬動過的節點（teardown 時還原）
  var railSig = '';          // 進度欄目前的步驟清單簽章（變了才重建）
  var railBtns = {};
  var lastWork = null;       // 上一次的「工作中步驟」（要不要捲動／聚焦用）
  var pending = false;       // refresh 合併旗標
  var styleDone = false;
  var snap = null;           // 最近一次 render 的結果（測試與按鈕處理用）

  // ── 各步驟的檢查（回傳「還需要」的字串陣列；空陣列＝有效）──────────
  function vCustomer() { return val('qCompany') ? [] : ['客戶公司']; }

  function vProject() {
    var m = [];
    if (!val('qProjectName')) m.push('專案名稱');
    var d = val('qDate'), v = val('qValidUntil');
    if (!d) m.push('建立日期');
    if (!v) m.push('報價期限');
    else if (d && v < d) m.push('報價期限（不能早於建立日期）');
    return m;
  }

  function vProducts() { return selectedProducts().length ? [] : ['至少勾選 1 項商品']; }

  function vConsultant() { return val('qCostBy') ? [] : ['支援顧問']; }

  function itemRows() {
    var body = el('quoteItemsBody');
    return body ? Array.prototype.slice.call(body.querySelectorAll('tr')).filter(function (tr) { return !tr.dataset.kind; }) : [];
  }
  function rowKey(tr) { return tr.dataset.lid || tr.dataset.nid || ''; }
  /** 逐列檢查：品名、數量 > 0、單價欄有填（單價 0 視為贈品允許；但要使用者實際輸入過 0，預設帶出的 0 不算「有填」） */
  function itemsCheck() {
    if (typeof readQuoteItems === 'function') readQuoteItems();   // 讓還沒存檔的新列都有暫時代號（nid），單價「動過」的記錄才認得出是哪一列
    var rows = itemRows(), r = { n: rows.length, desc: [], qty: [], price: [], zero: 0, first: null };
    rows.forEach(function (tr, i) {
      var n = i + 1, d = tr.querySelector('.qi-desc'), q = tr.querySelector('.qi-qty'), p = tr.querySelector('.qi-price');
      var bad = null;
      if (!d || !d.value.trim()) { r.desc.push(n); bad = bad || d; }
      if (!q || !(parseFloat(q.value) > 0)) { r.qty.push(n); bad = bad || q; }
      var raw = p ? String(p.value).trim() : '', num = Number(raw);
      var filled = !!p && raw !== '' && isFinite(num) && num >= 0 && (num > 0 || touched.has(rowKey(tr)));
      if (!filled) { r.price.push(n); bad = bad || p; } else if (num === 0) r.zero++;
      if (bad && !r.first) r.first = bad;
    });
    return r;
  }
  function vItems() {
    var c = itemsCheck(), m = [];
    if (!c.n) return ['至少 1 個報價項目'];
    if (c.desc.length) m.push('第 ' + c.desc.join('、') + ' 列的品名');
    if (c.qty.length) m.push('第 ' + c.qty.join('、') + ' 列的數量（需大於 0）');
    if (c.price.length) m.push('第 ' + c.price.join('、') + ' 列的單價（贈品請填 0）');
    return m;
  }

  function itemsSubtotal() {
    if (typeof readQuoteItems !== 'function' || typeof quoteTotal !== 'function') return 0;
    return quoteTotal(readQuoteItems(), 'none', 0).sub;
  }
  function discountState() {
    var d = typeof readQuoteDiscount === 'function' ? readQuoteDiscount() : { discountType: 'none', discountValue: 0 };
    return d;
  }
  function vDiscount() {
    var d = discountState(), v = d.discountValue;
    if (d.discountType === 'percent' && !(v > 0 && v < 100)) return ['折扣百分比（需大於 0 且小於 100）'];
    if (d.discountType === 'amount' && !(v > 0 && v < itemsSubtotal())) return ['議價總額（需大於 0 且低於小計）'];
    return [];
  }

  function vPayment() {
    if (typeof readQuotePayment !== 'function' || typeof validateQuotePayment !== 'function') return [];
    var err = validateQuotePayment(readQuotePayment());
    return err ? [err] : [];
  }

  // ── 摘要（已完成步驟的一行文字；一律純文字）──────────────────────
  function sCustomer() {
    var parts = [val('qCompany')];
    var sel = el('qContactId'), o = sel && sel.selectedOptions && sel.selectedOptions[0];
    var cn = o && o.value ? (o.getAttribute('data-name') || '') : '';
    if (cn) parts.push('聯絡人：' + cn);
    return parts.join(' ｜ ');
  }
  function sProject() { return [val('qProjectName'), '報價日 ' + val('qDate'), '有效至 ' + val('qValidUntil')].join(' ｜ '); }
  function sProducts() {
    var n = selectedProducts();
    return n.length + ' 項：' + n.slice(0, 3).join('、') + (n.length > 3 ? '…等' : '');
  }
  function sConsultant() {
    var sel = el('qCostBy'), o = sel && sel.selectedOptions && sel.selectedOptions[0];
    var t = [o && o.value ? o.textContent : ''];
    var note = val('qCostNote');
    if (note) t.push('備註：' + trunc(note, 40));
    return t.join(' ｜ ');
  }
  function sItems() {
    var c = itemsCheck(), parts = [c.n + ' 項', '小計 ' + money(itemsSubtotal()) + '（未稅）'];
    if (c.zero) parts.push(c.zero + ' 項單價為 0（贈品）');
    return parts.join(' ｜ ');
  }
  function sDiscount() {
    var d = discountState(), t;
    if (d.discountType === 'percent') t = '折扣 ' + d.discountValue + '%';
    else if (d.discountType === 'amount') t = '議價總額 ' + money(d.discountValue) + '（未稅）';
    else t = '無折扣';
    var tot = el('qTotal');
    return t + (tot ? ' ｜ 含稅合計 ' + tot.textContent : '');
  }
  function sPayment() {
    var t = '';
    if (typeof quotePaymentSentence === 'function' && typeof readQuotePayment === 'function') t = quotePaymentSentence(readQuotePayment()).replace(/^3[.．]/, '');
    var cl = typeof readQuoteClauses === 'function' ? readQuoteClauses().length : 0;
    return t + (cl ? ' ｜ 追加條款 ' + cl + ' 條' : '');
  }

  // ── 步驟定義 ──────────────────────────────────────────────
  var DEFS = [
    { id: 'customer', name: '客戶', ask: '要報價給哪一位客戶？', pre: true, focus: '#qCompany', validate: vCustomer, summary: sCustomer, okHint: '聯絡人、電話、手機、地址都可以不填。' },
    { id: 'project', name: '專案', ask: '這是什麼專案？', pre: true, focus: '#qProjectName', validate: vProject, summary: sProject, okHint: '建立日期與報價期限已帶入預設值，需要時可以修改。' },
    { id: 'products', name: '商品', ask: '這次報價包含哪些商品？', sub: '勾選本次報價包含的商品（至少 1 項），系統依商品歸類決定簽核層級。', pre: true, focus: '#qProdSearch', validate: vProducts, summary: sProducts, okHint: '' },
    { id: 'consultant', name: '指派顧問', ask: '請指派支援顧問（成本填寫人）', cond: true, pre: true, focus: '#qCostBy', validate: vConsultant, summary: sConsultant, okHint: '「給顧問的備註」可以不填。' },
    { id: 'items', name: '報價項目', ask: '報價內容有哪些項目？', pre: true, focus: null, validate: vItems, summary: sItems, okHint: function () { return '目前小計 ' + money(itemsSubtotal()) + '（未稅）。要再加項目請按「＋ 新增項目」。'; } },
    { id: 'discount', name: '優惠', ask: '有要給優惠嗎？', focus: 'input[name="qDiscountType"]:checked', validate: vDiscount, summary: sDiscount, okHint: '不需要折扣就直接略過。' },
    { id: 'payment', name: '付款方式', ask: '付款方式與追加條款', sub: '付款方式已預設好，確認後按下一步；追加條款可以不填。', focus: '#qPayPreset', validate: vPayment, summary: sPayment, okHint: '' },
    { id: 'confirm', name: '確認並儲存', ask: '確認內容並儲存', confirm: true }
  ];
  function defOf(id) { for (var i = 0; i < DEFS.length; i++) if (DEFS[i].id === id) return DEFS[i]; return null; }

  // ── 搬動節點（只有「專案名稱」從日期列下面提到日期列上方；teardown 還原）──
  function moveBefore(node, ref) {
    if (!node || !ref || !ref.parentNode) return;
    moves.push({ node: node, parent: node.parentNode, next: node.nextSibling });
    ref.parentNode.insertBefore(node, ref);
  }
  function restoreMoves() {
    while (moves.length) { var m = moves.pop(); if (m.parent) m.parent.insertBefore(m.node, m.next && m.next.parentNode === m.parent ? m.next : null); }
  }

  // ── 計算每個步驟的狀態 ────────────────────────────────────
  function syncNeed() {
    var need = needsConsultant();
    if (need !== hadNeed) {
      if (!need) { var cb = el('qCostBy'); if (cb) cb.value = ''; }   // 取消需顧問的商品：清空顧問選擇（步驟也跟著消失）
      delete confirmed.consultant;                                     // 再勾回來時要重新指派
      hadNeed = need;
    }
    return need;
  }
  function evalSteps(need) {
    return DEFS.filter(function (d) { return !d.cond || need; }).map(function (d) {
      var miss = d.validate ? d.validate() : [];
      // 變成不完整就收回「已確認」：之後補好要再按一次下一步（不會在打字打到一半、剛好有效的瞬間就被收成摘要）
      if (miss.length && confirmed[d.id]) delete confirmed[d.id];
      return { d: d, id: d.id, name: d.name, miss: miss, valid: !miss.length, done: !d.confirm && !miss.length && !!confirmed[d.id], st: 'future' };
    });
  }

  function render(opts) {
    if (!active) return;
    opts = opts || {};
    var need = syncNeed();
    var steps = evalSteps(need);
    var ci = steps.length - 1, fi = ci, i;
    for (i = 0; i < ci; i++) { if (!steps[i].done) { fi = i; break; } }
    var allDone = fi === ci;
    var ei = -1;
    if (editing) {
      for (i = 0; i < steps.length; i++) if (steps[i].id === editing) ei = i;
      if (ei < 0 || ei > fi || ei === ci) { editing = null; ei = -1; }
    }
    if (ei >= 0 || allDone) review = false;
    steps.forEach(function (s, k) {
      if (k === ci) s.st = ei >= 0 ? 'future' : (allDone ? 'active' : (review ? 'review' : 'future'));
      else if (k === ei) s.st = k < fi ? 'edit' : 'active';
      else if (k < fi) s.st = 'done';
      else if (k === fi) s.st = ei >= 0 ? 'pending' : 'active';
      else s.st = 'future';
    });
    var cur = ei >= 0 ? ei : fi;
    var mainEl = el('qsMain');
    if (mainEl) mainEl.classList.toggle('qs-final', allDone && ei < 0);
    var missingN = 0;
    steps.forEach(function (s, k) { if (k < ci && !s.done) missingN++; });
    snap = { steps: steps, ci: ci, fi: fi, ei: ei, cur: cur, allDone: allDone, missingN: missingN };

    // 各步驟區塊
    DEFS.forEach(function (d) {
      var e = els[d.id];
      if (!e) return;
      var idx = -1;
      steps.forEach(function (s, k) { if (s.id === d.id) idx = k; });
      if (idx < 0) { e.wrap.setAttribute('data-st', 'na'); return; }
      paintStep(e, steps[idx], idx, steps);
    });
    paintRail(steps, cur, fi, ci);
    paintConfirm(steps, allDone, missingN);

    // 儲存按鈕與「毛利與簽核路徑」分頁
    var sb = el('quoteModalSave');
    if (sb && !(typeof _qSaving !== 'undefined' && _qSaving)) {
      sb.disabled = !allDone;
      sb.title = allDone ? '' : '還有 ' + missingN + ' 個步驟沒完成，完成後才能儲存';
    }
    var pb = el('quoteTabBtnPnl');
    if (pb) pb.style.display = allDone ? '' : 'none';
    if (!allDone && pnlTabVisible() && typeof switchQuoteTab === 'function') switchQuoteTab('info');

    // 工作中步驟換了（使用者按了下一步／修改…）：捲到可視區並聚焦第一個欄位
    var workId = steps[cur].id + (review ? '+r' : '');
    if (opts.focus && (opts.force || workId !== lastWork)) focusStep(steps[cur].id, true);
    else if (opts.initial) focusStep(steps[cur].id, false);
    lastWork = workId;
  }

  function paintStep(e, s, idx, steps) {
    var st = s.st, d = s.d;
    e.wrap.setAttribute('data-st', st);
    e.badge.textContent = st === 'done' ? '✓' : String(idx + 1);
    e.title.textContent = d.name;
    if (st === 'done') e.txt.textContent = d.summary ? d.summary() : '';
    else if (st === 'pending') e.txt.textContent = '尚未填寫（先完成上面的修改）';
    else e.txt.textContent = d.ask;
    e.txt.title = (st === 'done') ? e.txt.textContent : '';
    var sub = (st === 'active' || st === 'edit') ? (d.sub || '') : '';
    e.sub.textContent = sub; e.sub.hidden = !sub;
    e.edit.hidden = st !== 'done';
    e.edit.setAttribute('aria-label', '修改' + d.name);
    if (e.nav) {
      var isEdit = editing === s.id;
      if (s.miss.length) { e.hint.textContent = '還需要：' + s.miss.join('、'); e.hint.className = 'qs-hint need'; }
      else { e.hint.textContent = typeof d.okHint === 'function' ? d.okHint() : (d.okHint || ''); e.hint.className = 'qs-hint'; }
      e.btn.disabled = s.miss.length > 0;
      var label = isEdit ? '完成修改' : (d.id === 'discount' && discountState().discountType === 'none' ? '下一步（略過也可以）' : '下一步');
      e.btn.textContent = label;
    }
  }

  function paintRail(steps, cur, fi, ci) {
    var rail = el('qsRail');
    if (!rail) return;
    var sig = steps.map(function (s) { return s.id; }).join(',');
    if (sig !== railSig) {
      railSig = sig; railBtns = {};
      rail.textContent = '';
      rail.appendChild(mk('div', 'qs-rail-h', '填寫進度'));
      var mini = mk('div', 'qs-rail-mini'); mini.id = 'qsRailMini'; mini.setAttribute('aria-live', 'polite');
      rail.appendChild(mini);
      var ol = mk('ol', 'qs-rail-list');
      steps.forEach(function (s) {
        var li = mk('li');
        var b = mk('button', 'qs-ri'); b.type = 'button'; b.setAttribute('data-id', s.id);
        b.appendChild(mk('span', 'm')); b.appendChild(mk('span', 'n', s.name));
        b.addEventListener('click', function () { onRail(s.id); });
        li.appendChild(b); ol.appendChild(li); railBtns[s.id] = b;
      });
      rail.appendChild(ol);
    }
    steps.forEach(function (s, k) {
      var b = railBtns[s.id];
      if (!b) return;
      var isCur = k === cur, isDone = s.done && !isCur;
      var cls = 'qs-ri ' + (isCur ? 'cur' : (isDone ? 'done' : 'todo'));
      var clickable = isCur || isDone || k === fi || k === ci;
      if (!isCur && !isDone && clickable) cls += ' go';
      if (k === ci && review) cls += ' rev';
      b.className = cls;
      b.querySelector('.m').textContent = isCur ? '●' : (isDone ? '✓' : '○');
      b.disabled = !clickable;
      if (isCur) b.setAttribute('aria-current', 'step'); else b.removeAttribute('aria-current');
      b.setAttribute('aria-label', s.name + '：' + (isCur ? '目前步驟' : (isDone ? '已完成，按下可修改' : (clickable ? '尚未完成' : '尚未到'))));
    });
    var mini2 = el('qsRailMini');
    if (mini2) mini2.textContent = '步驟 ' + (cur + 1) + '／' + steps.length + '：' + steps[cur].name;
  }

  function paintConfirm(steps, allDone, missingN) {
    var e = els.confirm;
    if (!e || !e.list) return;
    var st = e.wrap.getAttribute('data-st');
    if (st !== 'active' && st !== 'review') return;
    e.list.textContent = '';
    var fi = snap.fi;
    steps.forEach(function (s, k) {
      if (s.id === 'confirm') return;
      var row, mark, txt, btnText = '';
      if (s.done) { row = 'ok'; mark = '✓'; txt = s.d.summary ? s.d.summary() : ''; btnText = '修改'; }
      else if (s.miss.length) { row = 'bad'; mark = '✗'; txt = '還缺：' + s.miss.join('、'); if (k === fi) btnText = '前往填寫'; }
      else { row = 'todo'; mark = '○'; txt = '尚未按「下一步」確認'; if (k === fi) btnText = '前往確認'; }
      if (!s.done && k !== fi) txt += '（前面的步驟完成後才能填）';
      var li = mk('li', 'qs-cr ' + row);
      li.setAttribute('data-id', s.id);
      li.appendChild(mk('span', 'm', mark));
      li.appendChild(mk('span', 'n', s.name));
      li.appendChild(mk('span', 's', txt));
      if (btnText) {
        var gb = mk('button', 'g', btnText); gb.type = 'button';
        gb.addEventListener('click', function () { if (s.done) startEdit(s.id); else { editing = null; review = false; render({ focus: true, force: true }); } });
        li.appendChild(gb);
      }
      e.list.appendChild(li);
    });
    e.msg.className = 'qs-conf-msg ' + (allDone ? 'ok' : 'bad');
    e.msg.textContent = allDone ? '全部完成，可以按下方的「儲存」建立報價單。' : '還有 ' + missingN + ' 個步驟沒完成，完成後才能儲存（下方的「儲存」目前停用）。';
    var notes = [];
    var ap = el('qApprovalState');
    if (ap) notes.push('簽核狀態：' + ap.textContent.trim());
    var c = itemsCheck();
    if (c.zero) notes.push('提醒：有 ' + c.zero + ' 個品項單價為 0，會當成贈品；若不是，請按清單中「報價項目」那一列的「修改」補上單價。');
    e.note.textContent = notes.join('　');
  }

  // ── 動作 ──────────────────────────────────────────────────
  function stepById(id) { if (!snap) return null; for (var i = 0; i < snap.steps.length; i++) if (snap.steps[i].id === id) return snap.steps[i]; return null; }

  function next(id) {
    var s = stepById(id);
    if (!active || !s || s.d.confirm || s.d.validate().length) return;   // 按下的當下再驗一次（不依賴上次畫面的快照）
    confirmed[id] = true;
    if (editing === id) editing = null;
    review = false;
    render({ focus: true, force: true });
  }
  function startEdit(id) {
    var s = stepById(id);
    if (!active || !s || !s.done) return;
    editing = id; review = false;
    render({ focus: true, force: true });
  }
  function onRail(id) {
    var s = stepById(id);
    if (!active || !s || !snap) return;
    var k = snap.steps.indexOf(s);
    if (s.d.confirm) {
      editing = null;   // 看清單時不處在「修改」狀態（目前步驟＝第一個沒完成的步驟，或全部完成就是確認並儲存）
      if (snap.allDone) { render({ focus: true, force: true }); return; }
      review = !review;
      render({ focus: review, force: true });
      return;
    }
    if (s.done && k !== snap.cur) { startEdit(id); return; }
    if (k === snap.fi && snap.ei >= 0 && snap.ei !== snap.fi) { editing = null; render({ focus: true, force: true }); return; }   // 從修改中回到目前進度
    if (k === snap.cur) focusStep(id, true);
  }

  function visibleIn(node, box) {
    if (!node || !box) return false;
    var r = node.getBoundingClientRect(), b = box.getBoundingClientRect();
    return r.width > 0 && r.height > 0 && r.top >= b.top && r.bottom <= b.bottom;
  }
  function focusStep(id, scroll) {
    var e = els[id], d = defOf(id);
    if (!e || !d) return;
    var target = null;
    if (id === 'items') { var c = itemsCheck(); target = c.first || document.querySelector('#quoteItemsBody .qi-desc'); }
    else if (id === 'confirm') target = e.conf;
    else if (d.focus) target = document.querySelector(d.focus);
    if (scroll && e.wrap.scrollIntoView) e.wrap.scrollIntoView({ block: 'start' });
    if (!target || typeof target.focus !== 'function') return;
    // 初次開啟（scroll=false）只在欄位已經在畫面內時才聚焦，不能把畫面捲走
    if (!scroll && !visibleIn(target, el('quoteTabContentInfo'))) return;
    target.focus({ preventScroll: true });
  }

  // ── 建立／拆除 DOM ────────────────────────────────────────
  function buildStepDom() {
    var main = el('qsMain');
    els = {};
    DEFS.forEach(function (d) {
      var wrap = d.confirm ? null : main.querySelector('.qs-step[data-qs="' + d.id + '"]');
      if (d.confirm) { wrap = mk('div', 'qs-step qs-gen'); wrap.setAttribute('data-qs', 'confirm'); main.appendChild(wrap); }
      if (!wrap) throw new Error('找不到步驟區塊 ' + d.id);
      var head = mk('div', 'qs-head qs-gen');
      var badge = mk('span', 'qs-badge');
      var ht = mk('div', 'qs-ht'), l1 = mk('div', 'qs-l1'), title = mk('span', 'qs-title'), txt = mk('span', 'qs-txt'), sub = mk('div', 'qs-sub');
      l1.appendChild(title); l1.appendChild(txt); ht.appendChild(l1); ht.appendChild(sub);
      var edit = mk('button', 'qs-edit', '修改'); edit.type = 'button';
      edit.addEventListener('click', function () { startEdit(d.id); });
      head.appendChild(badge); head.appendChild(ht); head.appendChild(edit);
      wrap.insertBefore(head, wrap.firstChild);
      var e = { wrap: wrap, head: head, badge: badge, title: title, txt: txt, sub: sub, edit: edit };
      if (d.confirm) {
        var conf = mk('div', 'qs-conf qs-gen'); conf.tabIndex = -1; conf.id = 'qsConfirm';
        var list = mk('ul', 'qs-conf-list'), msg = mk('div', 'qs-conf-msg'), note = mk('div', 'qs-conf-note');
        conf.appendChild(list); conf.appendChild(msg); conf.appendChild(note);
        wrap.appendChild(conf);
        e.conf = conf; e.list = list; e.msg = msg; e.note = note;
      } else {
        var nav = mk('div', 'qs-nav qs-gen'), hint = mk('span', 'qs-hint'), btn = mk('button', 'btn btn-primary qs-next', '下一步');
        hint.setAttribute('aria-live', 'polite'); btn.type = 'button';
        btn.addEventListener('click', function () { next(d.id); });
        nav.appendChild(hint); nav.appendChild(btn); wrap.appendChild(nav);
        e.nav = nav; e.hint = hint; e.btn = btn;
      }
      els[d.id] = e;
    });
  }

  function teardown() {
    var was = active;
    active = false; editing = null; review = false; snap = null; lastWork = null; railSig = ''; railBtns = {};
    var body = el('quoteTabContentInfo'), modal = modalEl();
    Array.prototype.slice.call(document.querySelectorAll('#quoteTabContentInfo .qs-gen')).forEach(function (n) { n.remove(); });
    Array.prototype.slice.call(document.querySelectorAll('#quoteTabContentInfo .qs-step')).forEach(function (n) { n.removeAttribute('data-st'); });
    restoreMoves();
    if (body) body.classList.remove('qs-on');
    if (modal) { modal.classList.remove('qs-mode'); modal.classList.remove('qs-info'); }
    var rail = el('qsRail'); if (rail) rail.textContent = '';
    var mn = el('qsMain'); if (mn) mn.classList.remove('qs-final');
    var pb = el('quoteTabBtnPnl'); if (pb) pb.style.display = '';
    if (was) {
      var sb = el('quoteModalSave');
      if (sb) { sb.removeAttribute('title'); if (!(typeof _qSaving !== 'undefined' && _qSaving)) sb.disabled = false; }
    }
    els = {};
  }

  function setup() {
    var body = el('quoteTabContentInfo'), main = el('qsMain'), rail = el('qsRail'), modal = modalEl();
    if (!body || !main || !rail || !modal) return;   // 標記不齊就不分步（維持完整表單）
    if (!styleDone) { var st = document.createElement('style'); st.id = 'quoteStepsStyle'; st.textContent = CSS; document.head.appendChild(st); styleDone = true; }
    bindOnce(body);
    confirmed = {}; editing = null; review = false; touched = new Set(); lastWork = null; railSig = '';
    hadNeed = needsConsultant();
    // 「專案名稱」是這一步要回答的事：提到建立日期／報價期限（預設值）上方（編輯舊單不搬，teardown 還原）
    var pg = el('qProjectName') && el('qProjectName').closest('.form-group');
    var dr = el('qDate') && el('qDate').closest('.form-row');
    moveBefore(pg, dr);
    buildStepDom();
    body.classList.add('qs-on');
    modal.classList.add('qs-mode');
    modal.classList.toggle('qs-info', infoTabVisible());
    active = true;
    // 預填的資料已經滿足必填的步驟（例如從別處帶入公司）視為已完成（仍可「修改」）；付款方式與優惠一定要使用者按過下一步
    DEFS.forEach(function (d) { if (d.pre && (!d.cond || hadNeed) && d.validate().length === 0) confirmed[d.id] = true; });
    render({ initial: true });
  }

  // ── 事件（只綁一次；active=false 時什麼都不做）──────────────────
  function schedule() {
    if (pending || !active) return;
    pending = true;
    Promise.resolve().then(function () {
      pending = false;
      try { render(); } catch (e) { fail(e); }
    });
  }
  function fail(e) {
    try { if (window.console) console.error('quote-steps 發生錯誤，改用完整表單', e); } catch (x) { /* ignore */ }
    try { teardown(); } catch (x2) { /* ignore */ }
  }

  function bindOnce(body) {
    if (body._qsBound) return;
    body._qsBound = true;
    var onAny = function (ev) {
      if (!active) return;
      var t = ev.target;
      if (ev.type === 'input' && t && t.classList && t.classList.contains('qi-price')) {
        var tr = t.closest('tr');
        if (tr) { if (!tr.dataset.lid && !tr.dataset.nid && typeof readQuoteItems === 'function') readQuoteItems(); touched.add(rowKey(tr)); }
      }
      schedule();
    };
    body.addEventListener('input', onAny);
    body.addEventListener('change', onAny);
    body.addEventListener('click', onAny);
    body.addEventListener('keydown', onKey);
  }

  var TEXTY = { text: 1, number: 1, date: 1, search: 1, tel: 1, email: 1, url: 1 };
  function onKey(ev) {
    if (!active || ev.key !== 'Enter' || ev.isComposing || ev.keyCode === 229 || ev.ctrlKey || ev.metaKey || ev.altKey || ev.shiftKey) return;
    var t = ev.target;
    if (!t || t.tagName !== 'INPUT' || !TEXTY[(t.type || 'text').toLowerCase()]) return;   // textarea、下拉選單、按鈕、勾選框維持原本的 Enter 行為
    if (t.id === 'qProdSearch') { ev.preventDefault(); return; }                            // 搜尋框：Enter 不跳步（避免搜尋到一半就離開商品步驟）
    var stepEl = t.closest('.qs-step');
    if (!stepEl || !snap) return;
    var id = stepEl.getAttribute('data-qs'), cur = snap.steps[snap.cur];
    if (!cur || cur.id !== id) return;
    ev.preventDefault();   // 不送出、不做瀏覽器預設動作
    if (id === 'items') {
      // 報價項目表：Enter 先在同一列往右一格；最後一格（單價）才等同「下一步」
      var order = ['qi-desc', 'qi-unit', 'qi-qty', 'qi-price'], row = t.closest('tr');
      var at = -1; order.forEach(function (c, k) { if (t.classList.contains(c)) at = k; });
      if (row && at >= 0 && at < order.length - 1) { var nx = row.querySelector('.' + order[at + 1]); if (nx) { nx.focus(); if (nx.select) nx.select(); } return; }
      if (at < 0) return;   // 分組標題／小計列的文字欄：不動作
    }
    if (!cur.miss.length) next(id);
  }

  // ── 對外介面 ──────────────────────────────────────────────
  return {
    /** openQuoteModal 開好表單之後呼叫：q 有值＝編輯既有單 → 完全還原成完整表單；q 為 null＝新增 → 進入逐步填寫 */
    onOpen: function (q) {
      try {
        teardown();
        if (q || flagOff()) return;
        setup();
      } catch (e) { fail(e); }
    },
    /** switchQuoteTab 之後呼叫 */
    onTab: function (name) {
      if (!active) return;
      var modal = modalEl();
      if (modal) modal.classList.toggle('qs-info', name === 'info');
      if (name === 'info') schedule();
    },
    /** 品項／優惠／商品勾選變動後呼叫（合併成一次重算） */
    refresh: function () { schedule(); },
    /** 測試與除錯用：目前的步驟狀態（唯讀快照） */
    state: function () {
      if (!active || !snap) return { active: false };
      return {
        active: true, editing: editing, review: review, allDone: snap.allDone, missingN: snap.missingN, cur: snap.steps[snap.cur].id,
        steps: snap.steps.map(function (s) { return { id: s.id, st: s.st, done: s.done, miss: s.miss.slice() }; })
      };
    }
  };
})();
