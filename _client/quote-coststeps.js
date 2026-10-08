// ════════════════════════════════════════════════════════════
// ── 顧問填成本「逐步填寫」(quote-coststeps.js) ─────────────────
// 顧問填成本對話框（quote-approval.js 的 openQuoteCostFill／renderCostFill）在「完整連動畫面」（可編輯、看得到價格）時，改成一次只問一件事。
// 這一層只管「呈現」：成本明細編輯器（QCL）、即時試算卡片、儲存／完成的資料收集與送出邏輯都沒有改；未到的步驟先隱藏、已完成的收成一行摘要。
// 固定顯示（不算步驟）：案件資訊、業務備註、頂端「即時試算」卡片列（五張卡＋簽核層級預覽＋兩個預覽按鈕）。
// 步驟：1 確認報價品項 → 2 顧問服務成本 → 3 軟體成本 → 4 硬體成本 → 5 差旅費用 → 6 其他費用（含印花稅勾選）→ 7 風險預留 → 8 確認並完成。
// 規則（沿用新增報價單逐步填寫 quote-steps.js 的做法）：
//   · 步驟「完成」＝使用者按過「下一步」（或明確按「這一類沒有成本，略過」）且目前仍有效；第一個沒完成的步驟＝目前步驟，後面的不顯示。
//   · 沒有任何列的成本分區，一定要按「這一類沒有成本，略過」才算完成（避免默默漏填）；有列的分區要先通過該區的欄位檢查
//     （項目名稱非空、數量大於 0、數字合法、連動／拆項要有有效的報價品項、連動列單位不衝突）。
//   · 已完成的步驟收成一行摘要＋「修改」；修改到變成不完整，後面的步驟自動收回；補好之後回到先前確認過的進度。
//   · 重新打開：這張單已經有儲存過的成本明細（草稿或已完成）→ 通過檢查的步驟都視為已完成，直接停在「確認並完成」（有步驟不通過就停在那一步）；
//     第一次填（沒有成本明細）→ 從步驟 1 依序進行。
//   · 「完成並通知業務」在全部步驟完成前停用；「儲存」草稿隨時可按（不要求走完步驟）。
//   · 工具列（由報價品項重新帶入／補入新品項）只放在「顧問服務成本」這一步（影響所有分區，所以集中在最前面）；
//     這兩個動作若改到別的分區的內容，那些分區的「已確認」會被取消、要重新確認。各分區自己的「＋新增／合併為一列／拖曳排序」留在各分區。
// 同一個成本明細編輯器只有一份：切換步驟時把它（#qapCfFs）搬到該步驟的內容區，並用 data-cs-only 讓編輯器只顯示該步驟的分區。
// 測試用關閉開關：只有 location.hostname 為 localhost 且 window.__qNoSteps === true 時才關閉（與新增報價單的逐步填寫共用同一個開關）。
// 失敗保護：任何例外 → 關掉逐步、重畫回完整畫面（輸入內容由 renderCostFill 的草稿保留機制帶回）。
// 所有使用者輸入的文字只用 textContent 顯示，不經 innerHTML。
// 掛鉤（quote-approval.js，整個檔案包在 IIFE 內，所以用參數把需要的函式交進來）：renderCostFill 尾端 → attach(s, api)；cfRecalc 尾端 → refresh(s)；setCostBusy 尾端 → syncFooter(s)；submitCostFill → canFinish(s)。
// ════════════════════════════════════════════════════════════
var CostSteps = (function () {
  'use strict';

  var NARROW_MAX = 639;   // 與 quote-steps.js 相同的窄螢幕斷點（≤ 此寬度進度欄換成「步驟 n／N」一行）
  var CAT_IDS = ['consult', 'software', 'hw', 'travel', 'other'];
  var CAT_NAME = { consult: '顧問服務成本', software: '軟體成本', hw: '硬體成本', travel: '差旅費用', other: '其他費用' };
  var STEP_IDS = ['items', 'consult', 'software', 'hw', 'travel', 'other', 'risk', 'confirm'];
  var STEP_NAME = { items: '確認報價品項', consult: CAT_NAME.consult, software: CAT_NAME.software, hw: CAT_NAME.hw, travel: CAT_NAME.travel, other: CAT_NAME.other, risk: '風險預留', confirm: '確認並完成' };
  var ROW_LIST_MAX = 6;
  var SKIP_TEXT = '這一類沒有成本，略過';
  var MAX_ITEM_ROWS = 50;   // 報價單總列數上限（含分組標題與小計列），與 submitCostFill／伺服器相同

  var CSS = `
/* 顧問填成本的逐步畫面。共用樣式（qs-step／qs-head／qs-badge／qs-nav／qs-conf…）由 quote-steps.js 的 QSteps.ensureStyle() 注入；這裡只放這個對話框獨有的部分 */
.qap-cs { margin-top: 4px; }
.qap-cs > .qs-rail { display: block; margin: 0 0 10px; }
.qap-cs .qs-rail-mini, .qap-body.cs-mode > .qs-rail .qs-rail-mini { cursor: pointer; }
.qap-cs .qs-rail-h { display: none; }
.qap-cs .qs-rail-list { display: flex; flex-wrap: wrap; gap: 6px; }
.qap-cs .qs-rail-list li { flex: 0 0 auto; }
.qap-cs .qs-ri { width: auto; padding: 6px 12px; border-radius: 999px; background: #fff; border: 1px solid #e0e4ee; }
.qap-cs .qs-ri.cur { background: #e8f0fe; border-color: #9ec5f4; }
.qap-cs .qs-ri.done { background: #f1f8f3; border-color: #cfe6d6; }
.qap-cs .qs-ri.done:hover { background: #e6f4ea; }
.qap-cs .qs-ri:disabled { opacity: .75; }
/* 分步時對話框內容區的上緣留白改成第一個區塊的 margin：即時試算卡片固定（sticky）時才不會在它上方露出捲動中的內容 */
.qap-body.cs-mode { padding-top: 0; scroll-padding-top: var(--cs-top, 0px); }
.qap-body.cs-mode > :first-child { margin-top: 16px; }
.qap-body.cs-mode > .qs-rail { display: block; margin: 0; }   /* 窄螢幕時進度欄搬到對話框內容的最上面（見 placeRail） */
.qap-cs.qs-on .qs-step { scroll-margin-top: var(--cs-top, 6px); }
.qap-cs .qs-step .qap-sec { border: 0; background: transparent; padding: 0; margin-bottom: 0; }
.qap-cs #qapCfRef > summary { display: none; }
.qap-cs #qapCfRisk > h3 { display: none; }
.qap-cs .qap-cf-head > h3 { display: none; }
.qap-cs .qap-cf-head { margin: 4px 0 8px; }
.qap-cs .qap-cf-newbar { margin-bottom: 4px; }
.qap-cs .qcl-sec { border: 0; padding: 0; margin-bottom: 4px; }
.qap-cs .qcl-grand { display: none; }
.qap-cs .qcl-root[data-cs-only] > .qcl-sec { display: none; }
.qap-cs .qcl-root[data-cs-only="consult"] > .qcl-sec[data-sec="consult"],
.qap-cs .qcl-root[data-cs-only="software"] > .qcl-sec[data-sec="software"],
.qap-cs .qcl-root[data-cs-only="hw"] > .qcl-sec[data-sec="hw"],
.qap-cs .qcl-root[data-cs-only="travel"] > .qcl-sec[data-sec="travel"],
.qap-cs .qcl-root[data-cs-only="other"] > .qcl-sec[data-sec="other"] { display: block; }
.qap-cs .qcl-root[data-cs-only]:not([data-cs-only="consult"]) > .qcl-tools,
.qap-cs .qcl-root[data-cs-only]:not([data-cs-only="consult"]) > .qcl-help { display: none; }
.qap-cs.qs-on .qs-step > .cs-body.qs-gen { margin-left: 8px; margin-right: 8px; }
.qap-cs .cs-btns { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-left: auto; }
.qap-cs .qs-next.cs-off { opacity: .45; cursor: not-allowed; }
.qap-cs .qs-hint.flash { animation: csFlash .5s ease-out; }
@keyframes csFlash { 0% { background: #fde8e8; } 100% { background: transparent; } }
.qap-cs .cs-tot { display: flex; flex-wrap: wrap; gap: 6px 22px; margin: 10px 0 0; font-size: 13px; }
.qap-cs .cs-tot i { font-style: normal; color: #6b7686; margin-right: 6px; }
.qap-cs .cs-tot b { color: #1a2d52; font-variant-numeric: tabular-nums; }
.qap-cs .cs-note { margin-top: 10px; padding: 8px 12px; background: #f6f8fc; border: 1px solid #e0e4ee; border-radius: 8px; font-size: 12.5px; line-height: 1.7; color: #444; }
.qap-cs .cs-note-h { font-weight: 600; margin-bottom: 4px; color: #1a2d52; }
.qap-cs .cs-note-b { white-space: pre-wrap; word-break: break-word; }
.qap-cs .qs-cr.skip .m { color: #8a94a3; }
.qap-cs .qs-cr.skip .s { color: #6b7686; }
.qap-cs .qs-conf-list .qs-cr:last-child { border-bottom: 0; }
@media (min-width: 900px) and (min-height: 900px) {
  .qap-cf-live.cs-sticky { position: sticky; top: 0; z-index: 4; background: #f6f7f9; padding: 0 0 8px; box-shadow: 0 8px 8px -8px rgba(0,0,0,.22); }
}
@media (max-width: ${NARROW_MAX}px) {
  .qap-body.cs-mode > :first-child { margin-top: 12px; }
  .qap-body.cs-mode > .qs-rail { position: sticky; top: 0; z-index: 5; background: #f6f7f9; padding: 8px 0 6px; margin: 0 0 6px; }
  .qap-body.cs-mode > .qs-rail + * { margin-top: 4px; }
  .qap-cs > .qs-rail { position: sticky; top: 0; z-index: 3; background: #f6f7f9; padding: 4px 0 6px; margin-bottom: 6px; }
  .qap-cs .qs-rail-list { display: none; }
  .qap-cs .qs-nav { flex-direction: column; align-items: stretch; }
  .qap-cs .cs-btns { margin-left: 0; }
  .qap-cs .cs-btns .btn { flex: 1 1 auto; }
}
body.dark .qap-cs .qs-ri { background: #161b22; border-color: #30363d; }
body.dark .qap-cs .qs-ri.cur { background: #0d2040; border-color: #1c3a5f; }
body.dark .qap-cs .qs-ri.done { background: #0f1a13; border-color: #1b5230; }
body.dark .qap-cs .qs-ri.done:hover { background: #0a1f10; }
body.dark .qap-cs .qs-hint.flash { animation-name: csFlashDark; }
@keyframes csFlashDark { 0% { background: #3a1313; } 100% { background: transparent; } }
body.dark .qap-cs .cs-tot i { color: #8b949e; }
body.dark .qap-cs .cs-tot b { color: #e6edf3; }
body.dark .qap-cs .cs-note { background: #12171e; border-color: #30363d; color: #c9d1d9; }
body.dark .qap-cs .cs-note-h { color: #e6edf3; }
body.dark .qap-cs .qs-cr.skip .m { color: #8b949e; }
body.dark .qap-cs .qs-cr.skip .s { color: #8b949e; }
body.dark .qap-cf-live.cs-sticky { background: #0d1117; box-shadow: 0 8px 8px -8px rgba(0,0,0,.6); }
@media (max-width: ${NARROW_MAX}px) {
  body.dark .qap-cs > .qs-rail, body.dark .qap-body.cs-mode > .qs-rail { background: #0d1117; }
}
@media (prefers-reduced-motion: reduce) { .qap-cs .qs-hint.flash { animation: none !important; } }
`;

  // ══════════════════════════════════════════════════════════
  // 純函式（不碰 DOM；scripts/check-quote-coststeps.js 以 vm 載入後直接測）
  // ══════════════════════════════════════════════════════════

  /** 窄螢幕判定（與 CSS 斷點同一個數字） */
  function isNarrow(width) { return Number(width) <= NARROW_MAX; }

  /** 列號清單文字：最多列 ROW_LIST_MAX 個，超過接「…等 N 列」 */
  function rowList(nums) {
    var a = Array.isArray(nums) ? nums : [];
    var shown = a.slice(0, ROW_LIST_MAX).join('、');
    return a.length > ROW_LIST_MAX ? shown + '…等 ' + a.length + ' 列' : shown;
  }

  /**
   * 成本分區（顧問服務／軟體／硬體／差旅／其他）的檢查：回傳「還需要」的字串陣列（空陣列＝有效）。
   * sec＝{ n（列數）, invalid, blankDesc, zeroQty, badTarget, conflict（列號陣列）}；skipped＝使用者明確按過「這一類沒有成本，略過」。
   * 沒有任何列：一定要按過略過才有效；有列：逐項檢查（略過旗標不再有作用）。
   */
  function catMiss(sec, skipped) {
    var s = sec || {}, m = [];
    if (!(s.n > 0)) return skipped ? [] : ['成本明細（沒有這一類成本請按「' + SKIP_TEXT + '」）'];
    if (s.invalid && s.invalid.length) m.push('第 ' + rowList(s.invalid) + ' 列的數量或成本單價（需為 0 以上的數字）');
    if (s.blankDesc && s.blankDesc.length) m.push('第 ' + rowList(s.blankDesc) + ' 列的項目名稱');
    if (s.zeroQty && s.zeroQty.length) m.push('第 ' + rowList(s.zeroQty) + ' 列的數量（需大於 0）');
    if (s.badTarget && s.badTarget.length) m.push('第 ' + rowList(s.badTarget) + ' 列的「對應」報價品項（連動／拆項要選一個報價品項，或改成「不對應」）');
    if (s.conflict && s.conflict.length) m.push('第 ' + rowList(s.conflict) + ' 列的連動單位（同一個報價品項的連動列單位需一致，或把其中幾列改成「拆項」）');
    return m;
  }

  /** 「確認報價品項」步驟的檢查：o＝{ problem（新增報價項目的欄位問題文字，沒有為 ''）, zeroNew（數量不大於 0 的新增項目品名陣列）, total（報價單總列數＋新增項目數）} */
  function itemsMiss(o) {
    var x = o || {}, m = [];
    if (x.problem) m.push(String(x.problem));
    if (x.zeroNew && x.zeroNew.length) m.push('新增的報價項目數量需大於 0：' + x.zeroNew.slice(0, 5).map(function (n) { return '「' + String(n || '（未命名）').slice(0, 12) + '」'; }).join('、'));
    if (x.total > MAX_ITEM_ROWS) m.push('新增報價項目後總列數會超過 ' + MAX_ITEM_ROWS + ' 列（含分組標題與小計列），請減少新增的項目');
    return m;
  }

  /** 風險預留步驟的檢查：沒選（'' 或 null）就缺 */
  function riskMiss(v) { return v === '' || v === null || v === undefined ? ['風險預留（請選擇；沒有風險請選 0%）'] : []; }

  /**
   * 各步驟的狀態（純函式版的 quote-steps.js render）：
   *   ids＝步驟 id 依序（最後一個是確認步驟）；missBy＝{ id: 還需要的字串陣列 }；confirmed＝{ id: true }（按過下一步／略過）；editing＝正在「修改」的步驟 id；review＝沒走完時看確認清單。
   *   變成不完整（有 miss）就收回「已確認」（回傳的 confirmed 已清掉）。
   * 回傳 { steps:[{id, miss, valid, done, st}], confirmed, editing, review, ci（確認步驟索引）, fi（第一個沒完成的索引；全部完成＝ci）, ei（修改中的索引或 -1）, cur（工作中步驟）, allDone, missingN }
   * st：done｜active｜edit｜review｜pending（修改別的步驟時，目前進度步驟的佔位）｜future。
   */
  function evaluate(ids, missBy, confirmedIn, editingIn, reviewIn) {
    var confirmed = {};
    Object.keys(confirmedIn || {}).forEach(function (k) { if (confirmedIn[k]) confirmed[k] = true; });
    var ci = ids.length - 1, i;
    var steps = ids.map(function (id, k) {
      var miss = missBy && missBy[id] ? missBy[id].slice() : [];
      if (miss.length && confirmed[id]) delete confirmed[id];
      return { id: id, miss: miss, valid: !miss.length, done: k !== ci && !miss.length && !!confirmed[id], st: 'future' };
    });
    var fi = ci;
    for (i = 0; i < ci; i++) { if (!steps[i].done) { fi = i; break; } }
    var allDone = fi === ci;
    var editing = editingIn || null, review = !!reviewIn, ei = -1;
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
    var missingN = 0;
    steps.forEach(function (s, k) { if (k < ci && !s.done) missingN++; });
    return { steps: steps, confirmed: confirmed, editing: editing, review: review, ci: ci, fi: fi, ei: ei, cur: ei >= 0 ? ei : fi, allDone: allDone, missingN: missingN };
  }

  /** 重新打開的規則：這張單已有儲存過的成本明細（草稿或已完成，costLines 是陣列）→ true＝所有步驟先視為已完成（仍要通過檢查）；否則從步驟 1 開始 */
  function isReopen(q) { return !!q && Array.isArray(q.costLines); }

  /** 重新打開時的初始狀態：confirmed＝除確認步驟外全部 true；skipped＝目前沒有任何列的成本分區（視為上次就是沒有這類成本）；一般（第一次填）兩者皆空 */
  function initialFlags(ids, reopen, catCounts) {
    var confirmed = {}, skipped = {};
    if (reopen) {
      ids.forEach(function (id, k) { if (k !== ids.length - 1) confirmed[id] = true; });
      CAT_IDS.forEach(function (c) { if (catCounts && catCounts[c] === 0) skipped[c] = true; });
    }
    return { confirmed: confirmed, skipped: skipped };
  }

  /** 有列的分區不再需要「略過」：回傳清掉這類旗標之後的 skipped（counts＝各分區目前的列數）；之後把列刪光，要重新按略過 */
  function pruneSkipped(skipped, counts) {
    var o = {};
    Object.keys(skipped || {}).forEach(function (k) { if (skipped[k] && !(counts && counts[k] > 0)) o[k] = true; });
    return o;
  }

  /** 一個分區目前內容的簽章（列的內容全部串起來）：用來偵測「不在畫面上的分區被整批動過」（重新帶入／補入新品項） */
  function sigOf(lines) {
    return (Array.isArray(lines) ? lines : []).map(function (l) {
      var o = l || {};
      return [o.desc, o.qty, o.unit, o.unitCost, o.rel, o.forLid, (o.forLids || []).join('|'), o.vendor, o.consultant, o.note].join('\u0001');
    }).join('\u0002');
  }

  /**
   * 整批操作偵測：prev／now＝{ cat: 簽章 }；shown＝目前畫面上顯示的分區（它自己的改動是使用者正在編輯，不算）。
   * 回傳「不在畫面上、先前已確認（含略過）、但內容變了」的分區 id——這些要取消確認、重新走一次。
   */
  function bulkChanged(prev, now, shown, confirmed) {
    var out = [];
    CAT_IDS.forEach(function (id) {
      if (id === shown) return;
      if (prev && prev[id] !== undefined && now && prev[id] !== now[id] && confirmed && confirmed[id]) out.push(id);
    });
    return out;
  }

  var ntd = function (cents) { return 'NT$ ' + Math.round((Number(cents) || 0) / 100).toLocaleString('en-US'); };

  /** 成本分區的一行摘要：沒有列（略過）→「沒有這一類成本（已略過）」；有列 → 列數＋小計（顧問服務另列委外成本；其他費用的小計含印花稅，印花稅另註） */
  function catSummary(id, sec, cents, extra) {
    var s = sec || {}, x = extra || {};
    if (!(s.n > 0)) {
      if (id === 'other' && x.stampCents > 0) return '沒有其他費用（已略過）｜印花稅 ' + ntd(x.stampCents) + '（依合約金額自動計算）';
      return '沒有這一類成本（已略過）';
    }
    var parts = [s.n + ' 列', '小計 ' + ntd(cents)];
    if (id === 'consult' && x.outsourcedCents > 0) parts.push('委外 ' + ntd(x.outsourcedCents));
    if (id === 'other' && x.stampCents > 0) parts.push('含印花稅 ' + ntd(x.stampCents));
    return parts.join('｜');
  }

  /**
   * 確認清單的一列：step＝evaluate 的 steps[k]；k＝索引；fi＝第一個沒完成的索引；text＝已完成時的摘要；skipped＝該步驟是明確略過的空分區。
   * 回傳 { row: ok|skip|bad|todo, mark, text, btn }（btn＝按鈕文字，'' 表示沒有按鈕）。
   */
  function confirmRow(step, k, fi, text, skipped) {
    var row, mark, t, btn = '';
    if (step.done) { row = skipped ? 'skip' : 'ok'; mark = skipped ? '–' : '✓'; t = text; btn = '修改'; }
    else if (step.miss.length) { row = 'bad'; mark = '✗'; t = '還缺：' + step.miss.join('、'); if (k === fi) btn = '前往填寫'; }
    else { row = 'todo'; mark = '○'; t = '尚未按「下一步」確認'; if (k === fi) btn = '前往確認'; }
    if (!step.done && k !== fi) t += '（前面的步驟完成後才能填）';
    return { row: row, mark: mark, text: t, btn: btn };
  }

  // ══════════════════════════════════════════════════════════
  // DOM 部分
  // ══════════════════════════════════════════════════════════
  function mk(tag, cls, text) { var n = document.createElement(tag); if (cls) n.className = cls; if (text != null) n.textContent = text; return n; }
  function q1(s, sel) { return s.ov.querySelector(sel); }
  function flagOff() { try { return location.hostname === 'localhost' && window.__qNoSteps === true; } catch (e) { return false; } }
  var styleDone = false;
  function ensureStyle() {
    if (styleDone || document.getElementById('quoteCostStepsStyle')) { styleDone = true; return; }
    var st = document.createElement('style'); st.id = 'quoteCostStepsStyle'; st.textContent = CSS; document.head.appendChild(st); styleDone = true;
  }

  /** 這個對話框這次要不要分步（renderCostFill 已確認是可編輯＋完整連動畫面）：測試開關、共用樣式與編輯器的新方法都在才開 */
  function allowed(s) {
    if (!s || s.cstOff || flagOff()) return false;
    if (typeof QSteps === 'undefined' || !QSteps || typeof QSteps.ensureStyle !== 'function') return false;
    return !!(s.ed && !s.ed.dead && typeof s.ed.peekCat === 'function');
  }

  function money(c) { return ntd(c); }
  // quote-approval.js 整個包在 IIFE 內，所以它的函式不是全域的：attach(s, api) 時由它交進來（s.cstApi）
  //   api = { riskValue(s)：目前畫面上的風險預留值（字串，'' ＝尚未選）, newItemsProblem(s, mark)：新增報價項目的欄位問題文字（mark＝順便標紅）,
  //           doneMessage(s, c, live, sync)：按「完成並通知業務」前的確認視窗文字, rerender(s)：重畫整個對話框 }
  function riskNow(s) { return s.cstApi ? s.cstApi.riskValue(s) : ''; }
  function itemsProblem(s, mark) { return s.cstApi ? s.cstApi.newItemsProblem(s, mark) : ''; }

  // ── 蒐集每個步驟目前的事實（唯讀；不改狀態）──────────────────
  function conflictKeys(s) {
    var o = {}, al = s.live && s.live.al;
    ((al && al.conflicts) || []).forEach(function (c) { o[c.lid || c.nid] = true; });
    return o;
  }

  function gather(s) {
    var C = s.cst, ed = s.ed;
    if (!ed || ed.dead || typeof ed.peekCat !== 'function') throw new Error('逐步填寫：成本明細編輯器不存在');
    var g = { secs: {}, miss: {}, zeroNew: [] };
    var conf = conflictKeys(s);
    CAT_IDS.forEach(function (cat) {
      var p = ed.peekCat(cat);
      if (!p) throw new Error('逐步填寫：讀不到分區 ' + cat);
      p.conflict = p.rows.filter(function (r) { return r.rel === 'link' && r.ids.some(function (k) { return conf[k]; }); }).map(function (r) { return r.no; });
      g.secs[cat] = p;
      g.miss[cat] = catMiss(p, !!C.skipped[cat]);
    });
    var live = s.live;
    ((live && live.zero) || []).forEach(function (z) { if (z && z.nid) g.zeroNew.push(z.desc); });
    g.miss.items = itemsMiss({
      problem: itemsProblem(s, false),
      zeroNew: g.zeroNew,
      total: (Array.isArray(s.q.items) ? s.q.items.length : 0) + (s.newItems || []).length,
    });
    g.miss.risk = riskMiss(riskNow(s));
    g.miss.confirm = [];
    return g;
  }

  function sigsOf(g) {
    var o = {};
    CAT_IDS.forEach(function (c) { o[c] = sigOf(g.secs[c].lines); });
    return o;
  }

  // ── 摘要文字（已完成步驟的一行；一律純文字）──────────────────
  function summaryOf(s, id, g) {
    var live = s.live || {};
    if (id === 'items') {
      var base = (Array.isArray(s.q.items) ? s.q.items : []).filter(function (it) { return it && it.kind !== 'title' && it.kind !== 'subtotal'; }).length;
      var nNew = (s.newItems || []).length;
      var chg = ((live.changes) || []).filter(function (c) { return c.field !== 'new'; }).length;
      var parts = [(base + nNew) + ' 項', '報價合計 ' + money(live.revenueCents) + '（未稅，折扣後）'];
      if (chg) parts.push('連動調整 ' + chg + ' 處');
      if (nNew) parts.push('新增 ' + nNew + ' 項（單價由業務補填）');
      return parts.join('｜');
    }
    if (id === 'risk') { var v = riskNow(s); return v === '' ? '尚未選擇' : '風險預留 ' + v + '%（顧問服務成本小計 × ' + v + '% 另計為成本）'; }
    if (CAT_IDS.indexOf(id) >= 0) {
      var p = g.secs[id], by = live.byCat || {};
      var cents = (by[id] || 0) + (id === 'other' ? (live.stampCents || 0) : 0);
      return catSummary(id, p, cents, { outsourcedCents: live.outsourced && live.outsourced.outsourcedCents, stampCents: id === 'other' ? live.stampCents : 0 });
    }
    return '';
  }

  var ASK = {
    items: '請確認報價品項',
    consult: '顧問服務有哪些成本？',
    software: '軟體（授權、訂閱…）有哪些成本？',
    hw: '硬體（設備、周邊…）有哪些成本？',
    travel: '差旅費用（交通、住宿…）有哪些？',
    other: '其他費用（交際費、印花稅…）有哪些？',
    risk: '這個專案的風險預留是多少？',
    confirm: '確認內容並完成',
  };
  var SUB = {
    items: '這是客戶看到的報價品項（業務原本的單價與數量）。確認無誤後按「確認，下一步」；需要多一個報價項目可按「＋新增報價項目」，單價由業務補填。',
    consult: '自家顧問填「顧問姓名」，委外請填「委外廠商」。這一類沒有成本請按「' + SKIP_TEXT + '」。',
    software: '每列的「對應」決定它和業務報價的關係。這一類沒有成本請按「' + SKIP_TEXT + '」。',
    hw: '每列的「對應」決定它和業務報價的關係。這一類沒有成本請按「' + SKIP_TEXT + '」。',
    travel: '差旅交通、住宿等費用，通常選「不對應（純成本）」。這一類沒有成本請按「' + SKIP_TEXT + '」。',
    other: '交際費等其他支出可新增列；印花稅勾選後依合約金額 ×0.1% 自動計算。沒有其他費用請按「' + SKIP_TEXT + '」（印花稅仍依勾選計算）。',
    risk: '沒有風險請選 0%。（計算方式見下方說明）',
    confirm: '',
  };

  function hintOk(s, id, g) {
    var live = s.live || {};
    if (id === 'items') return '目前報價合計 ' + money(live.revenueCents) + '（未稅，折扣後）。';
    if (id === 'risk') { var v = riskNow(s); return v === '' ? '' : '風險預留 ' + v + '%。'; }
    var p = g.secs[id];
    if (!p || !(p.n > 0)) return '已略過：這一類沒有成本。';
    var t = p.n + ' 列，小計 ' + money(p.cents) + '。';
    if (p.zeroCost > 0) t += ' 提醒：有 ' + p.zeroCost + ' 列成本單價是 0（贈品或不計成本的列可以是 0；若不是，請補上）。';
    return t;
  }

  // ── 建立／拆除 DOM ────────────────────────────────────────
  function buildStep(main, id) {
    var wrap = mk('div', 'qs-step qs-gen'); wrap.setAttribute('data-cs', id); wrap.setAttribute('data-st', 'future');
    var head = mk('div', 'qs-head qs-gen'), badge = mk('span', 'qs-badge');
    var ht = mk('div', 'qs-ht'), l1 = mk('div', 'qs-l1'), title = mk('span', 'qs-title'), txt = mk('span', 'qs-txt'), sub = mk('div', 'qs-sub');
    l1.appendChild(title); l1.appendChild(txt); ht.appendChild(l1); ht.appendChild(sub);
    var edit = mk('button', 'qs-edit', '修改'); edit.type = 'button';
    head.appendChild(badge); head.appendChild(ht); head.appendChild(edit);
    wrap.appendChild(head);
    var e = { id: id, wrap: wrap, head: head, badge: badge, title: title, txt: txt, sub: sub, edit: edit };
    if (id === 'confirm') {
      var conf = mk('div', 'qs-conf qs-gen'); conf.tabIndex = -1; conf.id = 'qapCsConfirm';
      var list = mk('ul', 'qs-conf-list'), tot = mk('div', 'cs-tot'), msg = mk('div', 'qs-conf-msg'), note = mk('div', 'cs-note');
      var nh = mk('div', 'cs-note-h', '按下「完成並通知業務」時，系統會再請你確認以下內容：'), nb = mk('div', 'cs-note-b');
      note.appendChild(nh); note.appendChild(nb);
      conf.appendChild(list); conf.appendChild(tot); conf.appendChild(msg); conf.appendChild(note);
      wrap.appendChild(conf);
      e.conf = conf; e.list = list; e.tot = tot; e.msg = msg; e.note = note; e.noteBody = nb;
    } else {
      var body = mk('div', 'cs-body qs-gen');
      wrap.appendChild(body);
      var nav = mk('div', 'qs-nav qs-gen'), hint = mk('span', 'qs-hint'), btns = mk('span', 'cs-btns');
      hint.setAttribute('aria-live', 'polite');
      var btn = mk('button', 'btn btn-primary qs-next', '下一步'); btn.type = 'button';
      if (CAT_IDS.indexOf(id) >= 0) {
        var skip = mk('button', 'btn btn-secondary cs-skip', SKIP_TEXT); skip.type = 'button'; skip.hidden = true;
        btns.appendChild(skip); e.skip = skip;
      }
      btns.appendChild(btn);
      nav.appendChild(hint); nav.appendChild(btns); wrap.appendChild(nav);
      e.body = body; e.nav = nav; e.hint = hint; e.btn = btn;
    }
    main.appendChild(wrap);
    return e;
  }

  function freshState() { return { confirmed: {}, skipped: {}, editing: null, review: false, sigs: {}, shown: null, lastWork: null, inited: false }; }

  function setup(s) {
    var ref = q1(s, '#qapCfRef'), fs = q1(s, '#qapCfFs'), head = q1(s, '.qap-cf-head'), risk = q1(s, '#qapCfRisk'), live = q1(s, '#qapCfLive');
    if (!ref || !fs || !risk || !live || !s.ed || typeof s.ed.peekCat !== 'function') throw new Error('逐步填寫：找不到需要的區塊');
    QSteps.ensureStyle();
    ensureStyle();
    var C = s.cst;
    if (!C) { C = s.cst = freshState(); }
    var root = mk('div', 'qap-cs qs-on'); root.id = 'qapCs';
    var rail = mk('nav', 'qs-rail'); rail.id = 'qapCsRail'; rail.setAttribute('aria-label', '填寫進度');
    var main = mk('div', 'qs-main'); main.id = 'qapCsMain';
    var park = mk('div', 'cs-park'); park.id = 'qapCsPark'; park.hidden = true;
    root.appendChild(rail); root.appendChild(main); root.appendChild(park);
    ref.parentNode.insertBefore(root, ref);
    var els = {};
    STEP_IDS.forEach(function (id) { els[id] = buildStep(main, id); });
    els.items.body.appendChild(ref); ref.open = true;
    if (head) els.consult.body.appendChild(head);
    els.risk.body.appendChild(risk);
    park.appendChild(fs);
    live.classList.add('cs-sticky');
    s.body.classList.add('cs-mode');
    var D = s.cstDom = { root: root, rail: rail, main: main, park: park, fs: fs, els: els, snap: null, railSig: '', railBtns: {}, pending: false };
    bind(s, D);
    if (s.cstMq && s.cstMq.mq.removeEventListener) s.cstMq.mq.removeEventListener('change', s.cstMq.fn);   // 重畫時換掉上一次的監聽（不累積）
    s.cstMq = null;
    if (window.matchMedia) {
      var mq = window.matchMedia('(max-width: ' + NARROW_MAX + 'px)');
      var onMq = function () { if (!D.root.isConnected || s.cstDom !== D) { if (mq.removeEventListener) mq.removeEventListener('change', onMq); return; } placeRail(s); };
      if (mq.addEventListener) { mq.addEventListener('change', onMq); s.cstMq = { mq: mq, fn: onMq }; }
    }
    var firstTime = !C.inited;
    if (firstTime) {
      var g0 = gather(s);
      var counts = {}; CAT_IDS.forEach(function (c) { counts[c] = g0.secs[c].n; });
      var f = initialFlags(STEP_IDS, isReopen(s.q), counts);
      C.confirmed = f.confirmed; C.skipped = f.skipped; C.sigs = sigsOf(g0); C.inited = true;
    }
    render(s, { initial: firstTime });
  }

  var lastS = null;   // 最近一次分步的對話框（測試用：state() 不帶參數時讀它）
  function attach(s, api) {
    s.cstApi = api; lastS = s;
    try { setup(s); } catch (e) { fail(s, e); }
  }

  function fail(s, e) {
    try { if (window.console) console.error('quote-coststeps 發生錯誤，改用完整畫面', e); } catch (x) { /* ignore */ }
    s.cstOff = true; s.cst = null; s.cstDom = null;
    try { s.body.classList.remove('cs-mode'); s.body.style.removeProperty('--cs-top'); } catch (x1) { /* ignore */ }
    try { if (s.cstApi && !s.closed) s.cstApi.rerender(s); } catch (x2) { /* ignore */ }
  }

  // ── 事件 ──────────────────────────────────────────────────
  function schedule(s) {
    var D = s.cstDom;
    if (!D || D.pending) return;
    D.pending = true;
    // 用 setTimeout（不是 microtask）：使用者真的操作時，瀏覽器在「每個監聽器」之間都會跑 microtask，
    // 我們的監聽器（在步驟區上）比 quote-approval.js 綁在整個對話框上的監聽器（例如風險預留下拉把值記進 s.risk）先執行，
    // microtask 會搶在後者之前重算而讀到舊值；等整個事件派送完（下一個 task）再算就對了。
    setTimeout(function () {
      D.pending = false;
      if (s.closed || s.cstDom !== D) return;
      try { render(s); } catch (e) { fail(s, e); }
    }, 0);
  }

  function bind(s, D) {
    var onAny = function () { schedule(s); };
    D.root.addEventListener('input', onAny);
    D.root.addEventListener('change', onAny);
    D.root.addEventListener('click', onAny);
    D.root.addEventListener('keydown', function (ev) { onKey(s, ev); });
    STEP_IDS.forEach(function (id) {
      var e = D.els[id];
      e.edit.addEventListener('click', function () { startEdit(s, id); });
      if (e.btn) e.btn.addEventListener('click', function () { next(s, id); });
      if (e.skip) e.skip.addEventListener('click', function () { skip(s, id); });
    });
  }

  var TEXTY = { text: 1, number: 1, date: 1, search: 1, tel: 1, email: 1, url: 1 };
  function visibleEl(n) { return !!n && !n.disabled && n.getClientRects().length > 0; }
  function onKey(s, ev) {
    var D = s.cstDom;
    if (!D || !D.snap || ev.key !== 'Enter' || ev.isComposing || ev.keyCode === 229 || ev.ctrlKey || ev.metaKey || ev.altKey || ev.shiftKey) return;
    var t = ev.target;
    if (!t || t.tagName !== 'INPUT' || !TEXTY[(t.type || 'text').toLowerCase()]) return;   // 下拉、按鈕、勾選框、多行文字維持原本的 Enter 行為
    var stepEl = t.closest('.qs-step');
    if (!stepEl) return;
    var cur = D.snap.steps[D.snap.cur];
    if (!cur || stepEl.getAttribute('data-cs') !== cur.id || cur.id === 'confirm') return;
    // 只有該步驟「最後一個文字／數字輸入框」按 Enter 才等於下一步；其餘照原本（成本明細往下一欄、新增報價項目往下一欄）
    var ins = Array.prototype.filter.call(stepEl.querySelectorAll('input'), function (n) { return TEXTY[(n.type || 'text').toLowerCase()] && visibleEl(n); });
    if (!ins.length || ins[ins.length - 1] !== t) return;
    ev.preventDefault();
    next(s, cur.id);
  }

  // ── 計算並畫出所有步驟 ────────────────────────────────────
  function placeEditor(s, cat) {
    var D = s.cstDom, C = s.cst, fs = D.fs;
    var target = cat ? D.els[cat].body : D.park;
    if (fs.parentNode !== target) target.appendChild(fs);
    var root = s.ov.querySelector('.qcl-root');
    if (root) { if (cat) root.setAttribute('data-cs-only', cat); else root.removeAttribute('data-cs-only'); }
    C.shown = cat || null;
  }

  function render(s, opts) {
    var C = s.cst, D = s.cstDom;
    if (!C || !D || s.closed || !D.root.isConnected) return;
    opts = opts || {};
    var g = gather(s);
    // 沒有列的分區又有列了：略過旗標失效（之後刪光要重新按略過）
    var counts = {}; CAT_IDS.forEach(function (c) { counts[c] = g.secs[c].n; });
    C.skipped = pruneSkipped(C.skipped, counts);
    // 不在畫面上的分區被整批動過（重新帶入／補入新品項）：取消確認
    var now = sigsOf(g);
    bulkChanged(C.sigs, now, C.shown, C.confirmed).forEach(function (id) { delete C.confirmed[id]; delete C.skipped[id]; g.miss[id] = catMiss(g.secs[id], false); });
    C.sigs = now;
    var ev = evaluate(STEP_IDS, g.miss, C.confirmed, C.editing, C.review);
    C.confirmed = ev.confirmed; C.editing = ev.editing; C.review = ev.review;
    var steps = ev.steps;
    var cur = steps[ev.cur];
    D.snap = { steps: steps, ci: ev.ci, fi: ev.fi, ei: ev.ei, cur: ev.cur, allDone: ev.allDone, missingN: ev.missingN, g: g };
    placeEditor(s, CAT_IDS.indexOf(cur.id) >= 0 ? cur.id : null);
    D.main.classList.toggle('qs-final', ev.allDone && ev.ei < 0);
    D.root.setAttribute('data-all-done', ev.allDone ? '1' : '0');
    steps.forEach(function (st, k) { paintStep(s, D.els[st.id], st, k, g); });
    paintRail(s, steps, ev);
    paintConfirm(s, g, ev);
    syncFooter(s);
    // 捲到某一步時要避開固定在上方的東西：寬螢幕＝固定的即時試算卡片；窄螢幕＝固定的「步驟 n／N」那一行
    var live = q1(s, '#qapCfLive'), top = 6;
    if (live && getComputedStyle(live).position === 'sticky') top = live.offsetHeight + 10;
    else if (getComputedStyle(D.rail).position === 'sticky') top = D.rail.offsetHeight + 8;
    s.body.style.setProperty('--cs-top', top + 'px');
    var workId = cur.id + (C.review ? '+r' : '');
    if (opts.focus && (opts.force || workId !== C.lastWork)) focusStep(s, C.review ? 'confirm' : cur.id, true);   // 沒走完時按進度欄檢視「確認並完成」：捲到清單
    else if (opts.initial) focusStep(s, cur.id, false);
    C.lastWork = workId;
  }

  function paintStep(s, e, st, idx, g) {
    var C = s.cst, id = st.id;
    e.wrap.setAttribute('data-st', st.st);
    e.badge.textContent = st.st === 'done' ? '✓' : String(idx + 1);
    e.title.textContent = STEP_NAME[id];
    var active = st.st === 'active' || st.st === 'edit' || st.st === 'review';
    if (st.st === 'done') e.txt.textContent = summaryOf(s, id, g);
    else if (st.st === 'pending') e.txt.textContent = '尚未填寫（先完成上面的修改）';
    else e.txt.textContent = ASK[id];
    e.txt.title = st.st === 'done' ? e.txt.textContent : '';
    var sub = (st.st === 'active' || st.st === 'edit') ? (SUB[id] || '') : '';
    e.sub.textContent = sub; e.sub.hidden = !sub;
    e.edit.hidden = st.st !== 'done';
    e.edit.setAttribute('aria-label', '修改' + STEP_NAME[id]);
    if (!e.nav) return;
    var isEdit = C.editing === id;
    if (st.miss.length) { e.hint.textContent = '還需要：' + st.miss.join('、'); e.hint.className = 'qs-hint need'; }
    else { e.hint.textContent = active ? hintOk(s, id, g) : ''; e.hint.className = 'qs-hint'; }
    var off = st.miss.length > 0;
    e.btn.classList.toggle('cs-off', off);
    if (off) e.btn.setAttribute('aria-disabled', 'true'); else e.btn.removeAttribute('aria-disabled');
    e.btn.textContent = isEdit ? '完成修改' : (id === 'items' ? '確認，下一步' : '下一步');
    if (e.skip) e.skip.hidden = !(active && g.secs[id].n === 0 && !C.skipped[id]);
  }

  /** 窄螢幕：進度欄那一行搬到對話框內容的最上面（不然要先捲過案件資訊與即時試算才看得到「步驟 n／N」）；寬螢幕放回步驟區上方 */
  function placeRail(s) {
    var D = s.cstDom, narrow = !!(window.matchMedia && window.matchMedia('(max-width: ' + NARROW_MAX + 'px)').matches);
    if (narrow) { if (D.rail.parentNode !== s.body || s.body.firstChild !== D.rail) s.body.insertBefore(D.rail, s.body.firstChild); }
    else if (D.rail.parentNode !== D.root || D.root.firstChild !== D.rail) D.root.insertBefore(D.rail, D.root.firstChild);
  }

  function paintRail(s, steps, ev) {
    var D = s.cstDom, rail = D.rail;
    placeRail(s);
    var sig = steps.map(function (x) { return x.id; }).join(',');
    if (sig !== D.railSig) {
      D.railSig = sig; D.railBtns = {};
      rail.textContent = '';
      rail.appendChild(mk('div', 'qs-rail-h', '填寫進度'));
      var mini = mk('div', 'qs-rail-mini'); mini.id = 'qapCsRailMini'; mini.setAttribute('aria-live', 'polite'); mini.title = '按一下捲到目前的步驟';
      mini.addEventListener('click', function () { var sn = D.snap; if (sn) focusStep(s, sn.steps[sn.cur].id, true); });
      rail.appendChild(mini);
      var ol = mk('ol', 'qs-rail-list');
      steps.forEach(function (x) {
        var li = mk('li');
        var b = mk('button', 'qs-ri'); b.type = 'button'; b.setAttribute('data-id', x.id);
        b.appendChild(mk('span', 'm')); b.appendChild(mk('span', 'n', STEP_NAME[x.id]));
        b.addEventListener('click', function () { onRail(s, x.id); });
        li.appendChild(b); ol.appendChild(li); D.railBtns[x.id] = b;
      });
      rail.appendChild(ol);
    }
    steps.forEach(function (x, k) {
      var b = D.railBtns[x.id];
      if (!b) return;
      var isCur = k === ev.cur, isDone = x.done && !isCur;
      var clickable = isCur || isDone || k === ev.fi || k === ev.ci;
      var cls = 'qs-ri ' + (isCur ? 'cur' : (isDone ? 'done' : 'todo'));
      if (!isCur && !isDone && clickable) cls += ' go';
      if (k === ev.ci && s.cst.review) cls += ' rev';
      b.className = cls;
      b.querySelector('.m').textContent = isCur ? '●' : (isDone ? '✓' : '○');
      b.disabled = !clickable;
      if (isCur) b.setAttribute('aria-current', 'step'); else b.removeAttribute('aria-current');
      b.setAttribute('aria-label', STEP_NAME[x.id] + '：' + (isCur ? '目前步驟' : (isDone ? '已完成，按下可修改' : (clickable ? '尚未完成' : '尚未到'))));
    });
    var mini2 = q1(s, '#qapCsRailMini');
    if (mini2) mini2.textContent = '步驟 ' + (ev.cur + 1) + '／' + steps.length + '：' + STEP_NAME[steps[ev.cur].id];
  }

  function paintConfirm(s, g, ev) {
    var e = s.cstDom.els.confirm, st = e.wrap.getAttribute('data-st');
    if (st !== 'active' && st !== 'review') return;
    var C = s.cst, live = s.live || {};
    e.list.textContent = '';
    ev.steps.forEach(function (x, k) {
      if (x.id === 'confirm') return;
      var isSkip = CAT_IDS.indexOf(x.id) >= 0 && g.secs[x.id].n === 0;
      var r = confirmRow(x, k, ev.fi, x.done ? summaryOf(s, x.id, g) : '', x.done && isSkip);
      var li = mk('li', 'qs-cr ' + r.row); li.setAttribute('data-id', x.id);
      li.appendChild(mk('span', 'm', r.mark)); li.appendChild(mk('span', 'n', STEP_NAME[x.id])); li.appendChild(mk('span', 's', r.text));
      if (r.btn) {
        var gb = mk('button', 'g', r.btn); gb.type = 'button';
        gb.addEventListener('click', function () { if (x.done) startEdit(s, x.id); else { C.editing = null; C.review = false; render(s, { focus: true, force: true }); } });
        li.appendChild(gb);
      }
      e.list.appendChild(li);
    });
    // 合計（數字來源與頂端卡片相同）
    e.tot.textContent = '';
    var pair = function (k, v) { var sp = mk('span'); sp.appendChild(mk('i', null, k)); sp.appendChild(mk('b', null, v)); e.tot.appendChild(sp); };
    pair('報價合計（未稅，折扣後）', money(live.revenueCents));
    pair('成本合計（含印花稅）', money(live.costCents));
    pair('毛利', money(live.gpCents));
    pair('毛利率', live.marginText === null || live.marginText === undefined ? '—' : live.marginText + '%');
    pair('委外佔比', ((live.outsourced && live.outsourced.outsourcedPctOfCost) || '0.00') + '%');
    var rv = riskNow(s);
    pair('風險預留', rv === '' ? '尚未選擇' : rv + '%');
    e.msg.className = 'qs-conf-msg ' + (ev.allDone ? 'ok' : 'bad');
    e.msg.textContent = ev.allDone
      ? '全部完成。按下方的「完成並通知業務」會再確認一次並通知業務；只想先存起來請按「儲存」。'
      : '還有 ' + ev.missingN + ' 個步驟沒完成，完成後才能按「完成並通知業務」（下方按鈕目前停用）。「儲存」草稿隨時可以按。';
    // 完成前的提醒：與按下「完成並通知業務」時的確認視窗同一份文字（cfDoneMessage）
    var note = '';
    try {
      var zc = 0, lines = [];
      CAT_IDS.forEach(function (c) { zc += g.secs[c].zeroCost; lines = lines.concat(g.secs[c].lines); });
      var lns = s.ed && !s.ed.dead ? s.ed.getLines() : lines;
      if (s.cstApi) note = s.cstApi.doneMessage(s, { zeroCost: zc, lines: lns }, s.live, true);
    } catch (x) { note = ''; }
    e.note.hidden = !note;
    e.noteBody.textContent = note;
  }

  function syncFooter(s) {
    var D = s.cstDom;
    if (!s.cst || !D || !D.snap || !D.root.isConnected) return;
    var d = s.ov.querySelector('.qap-footer [data-act="done"]');
    if (!d) return;
    var ok = D.snap.allDone;
    d.disabled = !!s.busy || !ok;
    d.title = ok ? '' : '還有 ' + D.snap.missingN + ' 個步驟沒完成，走完所有步驟後才能完成並通知業務';
  }

  /** 按「完成並通知業務」前再驗一次（不依賴上次畫面的快照）；引擎本身出錯時不擋（伺服器仍會驗證） */
  function canFinish(s) {
    try {
      if (!s.cst || !s.cstDom) return true;
      var g = gather(s);
      return evaluate(STEP_IDS, g.miss, s.cst.confirmed, null, false).allDone;
    } catch (e) { return true; }
  }

  // ── 動作 ──────────────────────────────────────────────────
  function stepOf(s, id) {
    var D = s.cstDom; if (!D || !D.snap) return null;
    for (var i = 0; i < D.snap.steps.length; i++) if (D.snap.steps[i].id === id) return D.snap.steps[i];
    return null;
  }

  function markInvalid(s, id) {
    var D = s.cstDom, e = D.els[id], t = null;
    if (CAT_IDS.indexOf(id) >= 0) {
      var mount = q1(s, '#qapCfMount');
      if (mount && typeof QCL !== 'undefined') QCL.collect(mount);   // 標紅：空白項目、數量 0、數字不合法、對應目標無效
      t = e.wrap.querySelector('.qcl-bad, .qcl-warn, .qcl-conf');
    } else if (id === 'items') {
      itemsProblem(s, true);
      t = e.wrap.querySelector('.qci-bad');
    } else if (id === 'risk') t = e.wrap.querySelector('select[data-risk]');
    if (e.hint) { e.hint.classList.remove('flash'); void e.hint.offsetWidth; e.hint.classList.add('flash'); }
    if (t && typeof t.focus === 'function') { if (t.scrollIntoView) t.scrollIntoView({ block: 'center' }); t.focus({ preventScroll: true }); }
  }

  function next(s, id) {
    var C = s.cst, D = s.cstDom;
    if (!C || !D || !D.snap || s.closed) return;
    var st = stepOf(s, id);
    if (!st || id === 'confirm') return;
    if (st.st !== 'active' && st.st !== 'edit') return;   // 只有工作中的步驟能按下一步
    var g = gather(s);
    if (g.miss[id].length) { render(s); markInvalid(s, id); return; }   // 按下的當下再驗一次（不依賴上次畫面的快照）
    C.confirmed[id] = true;
    if (C.editing === id) C.editing = null;
    C.review = false;
    render(s, { focus: true, force: true });
  }

  function skip(s, id) {
    var C = s.cst, D = s.cstDom;
    if (!C || !D || s.closed || CAT_IDS.indexOf(id) < 0) return;
    var g = gather(s);
    if (g.secs[id].n !== 0) return;   // 有列就不能略過（要略過請先刪光）
    C.skipped[id] = true; C.confirmed[id] = true;
    if (C.editing === id) C.editing = null;
    C.review = false;
    render(s, { focus: true, force: true });
  }

  function startEdit(s, id) {
    var C = s.cst, st = stepOf(s, id);
    if (!C || !st || !st.done) return;
    C.editing = id; C.review = false;
    render(s, { focus: true, force: true });
  }

  function onRail(s, id) {
    var C = s.cst, D = s.cstDom, st = stepOf(s, id);
    if (!C || !D || !D.snap || !st) return;
    var snap = D.snap, k = snap.steps.indexOf(st);
    if (id === 'confirm') {
      C.editing = null;   // 看清單時不處在「修改」狀態
      if (snap.allDone) { render(s, { focus: true, force: true }); return; }
      C.review = !C.review;
      render(s, { focus: C.review, force: true });
      return;
    }
    if (st.done && k !== snap.cur) { startEdit(s, id); return; }
    if (k === snap.fi && snap.ei >= 0 && snap.ei !== snap.fi) { C.editing = null; render(s, { focus: true, force: true }); return; }   // 從修改中回到目前進度
    if (k === snap.cur) focusStep(s, id, true);
  }

  function visibleIn(node, box) {
    if (!node || !box) return false;
    var r = node.getBoundingClientRect(), b = box.getBoundingClientRect();
    return r.width > 0 && r.height > 0 && r.top >= b.top && r.bottom <= b.bottom;
  }

  function focusStep(s, id, scroll) {
    var D = s.cstDom, e = D && D.els[id];
    if (!e) return;
    var target = null;
    if (id === 'items') target = e.wrap.querySelector('.qci-bad') || e.btn;
    else if (id === 'confirm') target = e.conf;
    else if (id === 'risk') target = e.wrap.querySelector('select[data-risk]');
    else {
      target = e.wrap.querySelector('.qcl-bad, .qcl-warn');
      if (!target) {
        var p = s.ed && !s.ed.dead && typeof s.ed.peekCat === 'function' ? s.ed.peekCat(id) : null;
        target = p && p.n > 0 ? e.wrap.querySelector('tbody[data-cat="' + id + '"] tr.qcl-row .qcl-desc') : e.wrap.querySelector('[data-act="add"][data-cat="' + id + '"]');
      }
    }
    if (scroll && e.wrap.scrollIntoView) e.wrap.scrollIntoView({ block: 'start' });
    if (!target || typeof target.focus !== 'function') return;
    if (!scroll && !visibleIn(target, s.body)) return;   // 初次開啟只在欄位已經在畫面內時才聚焦，不能把畫面捲走
    target.focus({ preventScroll: true });
  }

  // ── 對外介面 ──────────────────────────────────────────────
  return {
    allowed: allowed,
    attach: attach,
    /** 成本明細／新增報價項目有變動（QCL onChange → cfRecalc）：合併成一次重算 */
    refresh: function (s) { if (s && s.cstDom && s.cstDom.root && s.cstDom.root.isConnected) schedule(s); },
    syncFooter: syncFooter,
    canFinish: canFinish,
    /** 測試與除錯用：目前的步驟狀態（唯讀快照） */
    state: function (s0) {
      var s = s0 || lastS;
      var C = s && s.cst, D = s && s.cstDom;
      if (!C || !D || !D.snap) return { active: false };
      var sn = D.snap;
      return {
        active: true, editing: C.editing, review: C.review, allDone: sn.allDone, missingN: sn.missingN, cur: sn.steps[sn.cur].id, shown: C.shown,
        skipped: JSON.parse(JSON.stringify(C.skipped)), steps: sn.steps.map(function (x) { return { id: x.id, st: x.st, done: x.done, miss: x.miss.slice() }; })
      };
    },
    /** 純函式（單元測試用） */
    core: {
      NARROW_MAX: NARROW_MAX, CAT_IDS: CAT_IDS, STEP_IDS: STEP_IDS, STEP_NAME: STEP_NAME, SKIP_TEXT: SKIP_TEXT,
      isNarrow: isNarrow, rowList: rowList, catMiss: catMiss, itemsMiss: itemsMiss, riskMiss: riskMiss, evaluate: evaluate,
      isReopen: isReopen, initialFlags: initialFlags, pruneSkipped: pruneSkipped, sigOf: sigOf, bulkChanged: bulkChanged, catSummary: catSummary, confirmRow: confirmRow,
      cssText: function () { return CSS; },
    },
  };
})();
