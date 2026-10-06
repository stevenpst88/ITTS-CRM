// ════════════════════════════════════════════════════════════
// ── 報價單管理 (quote.js) ────────────────────────────────────
// 建單/編輯表單（商品必勾、支援顧問）、列表（簽核/成本狀態、依 perm 顯示操作）、
// 送簽/撤回、下載警示、毛利與簽核路徑摘要。
// 簽核面板/顧問填成本/簽核設定在 quote-approval.js（W4），這裡只呼叫並容錯。
// 資料形狀與端點見規格第 5 節（伺服器依「檢視者」過濾欄位，前端不可假設欄位一定存在）。
// ════════════════════════════════════════════════════════════

let allQuotations = [];
let _qCfg = null;          // GET /quote-approval/config 的結果（每次進入列表重抓）
let _qInboxIds = null;     // Set<報價單 id>：我待處理（簽核＋填成本）；null＝尚未取得
let _qInboxOnly = false;   // 「待我處理」篩選
let _qEditing = null;      // 編輯中的報價單（伺服器物件）；null＝新增
let _qSaving = false;

// 簽核關卡顯示名稱（與 lib/quoteApproval.js 的 TIERS 一致；伺服器有給 label 時以伺服器為準）
const QUOTE_TIER_LABEL = { mgr1: '一級主管', gm: '總經理', chairman: '董事長', board: '董事會決議（秘書代核）' };
// 舊版（簽核功能上線前）的狀態，只做唯讀顯示
const QUOTE_LEGACY_LABEL = { sent: '已寄出', accepted: '已接受', rejected: '已拒絕' };

// ── 樣式（自帶，不動 style.css；暗色比照既有 body.dark 色系）──────────
const QUOTE_CSS = `
.q-inbox-badge { display:inline-block; min-width:18px; padding:0 6px; border-radius:9px; background:#d93025; color:#fff;
  font-size:11px; font-weight:700; line-height:18px; text-align:center; margin-left:4px; vertical-align:middle; }
#quoteInboxBtn.q-on { background:#1a73e8; color:#fff; border-color:#1a73e8; }
.q-ro-state { padding:7px 10px; border:1px dashed #d5d9e0; border-radius:6px; background:#f8f9fc; font-size:13px; line-height:1.7; min-height:34px; box-sizing:border-box; }
.q-ro-state .q-sub { display:block; font-size:12px; color:#6b7686; }
.quote-section-title .required { color:#ea4335; }
.q-sec-sub { font-size:12px; font-weight:400; color:#8a94a3; margin-left:8px; }
.q-hint { font-size:12px; color:#8a94a3; margin-top:4px; line-height:1.5; }
.q-infobar { margin-top:10px; background:#e8f0fe; border:1px solid #c6dafc; color:#174ea6; border-radius:8px; padding:8px 12px; font-size:12.5px; line-height:1.6; }
.q-warnbar { margin-top:8px; background:#fff4e5; border:1px solid #f5c98b; color:#8a4b00; border-radius:8px; padding:8px 12px; font-size:12.5px; line-height:1.6; }
.q-errbar { margin-top:8px; background:#fce8e6; border:1px solid #f5c6cb; color:#a31515; border-radius:8px; padding:8px 12px; font-size:12.5px; line-height:1.6; }
.q-costby { margin-top:12px; padding:10px 12px 0; border:1px solid #e0e4ee; border-radius:8px; background:#fafbfe; }

/* 商品勾選 */
.qp-top { display:flex; gap:10px; align-items:center; flex-wrap:wrap; margin-bottom:8px; }
.qp-search { flex:1; min-width:160px; border:1px solid #ddd; border-radius:6px; padding:6px 10px; font-size:13px; box-sizing:border-box; background:#fff; color:#333; }
.qp-count { font-size:12.5px; color:#5f6b7a; white-space:nowrap; }
.qp-selected { display:flex; flex-wrap:wrap; gap:6px; margin-bottom:8px; }
.qp-chip { display:inline-flex; align-items:center; gap:4px; background:#e8f0fe; color:#174ea6; border-radius:12px; padding:2px 6px 2px 10px; font-size:12px; max-width:100%; }
.qp-chip button { border:none; background:none; color:inherit; cursor:pointer; font-size:13px; line-height:1; padding:0 2px; }
.qp-list { max-height:240px; overflow:auto; border:1px solid #e0e4ee; border-radius:8px; background:#fff; }
.qp-grp-h { position:sticky; top:0; background:#f1f4fb; color:#1a2d52; font-size:12px; font-weight:700; padding:5px 10px; border-bottom:1px solid #e6e9f2; }
.qp-item { display:flex; align-items:center; gap:8px; padding:6px 10px; font-size:13px; cursor:pointer; border-bottom:1px solid #f3f4f8; flex-wrap:wrap; }
.qp-item:hover { background:#f8f9ff; }
.qp-item input { width:16px; height:16px; flex-shrink:0; cursor:pointer; }
.qp-nm { flex:1; min-width:120px; word-break:break-all; }
.qp-tag { font-size:11px; padding:1px 8px; border-radius:10px; background:#eef1f6; color:#5f6b7a; white-space:nowrap; }
.qp-tag.qp-un { background:#fff3e0; color:#b25e00; font-weight:600; }
.qp-tag.qp-self { background:#e6f4ea; color:#188038; }
.qp-empty { padding:16px; text-align:center; color:#8a94a3; font-size:13px; }

/* 列表：簽核/成本狀態與操作 */
.quote-st-ready   { background:#e8f0fe; color:#1a73e8; }
.quote-st-wait    { background:#fff3e0; color:#e65100; }
.quote-st-pending { background:#e3f2fd; color:#0b57d0; }
.quote-st-returned{ background:#fce8e6; color:#c62828; }
.quote-st-approved{ background:#e6f4ea; color:#188038; }
.quote-st-void    { background:#fff1e0; color:#b25e00; border:1px solid #f5c98b; }
.quote-status.q-wrap { white-space:normal; line-height:1.5; display:inline-block; max-width:230px; }
.q-legacy { display:block; font-size:11px; color:#8a94a3; margin-top:2px; }
.q-cost { font-size:12px; white-space:nowrap; color:#5f6b7a; }
.q-cost.warn { color:#e65100; font-weight:600; }
.q-cost.ok { color:#188038; font-weight:600; }
.q-actions { display:flex; flex-wrap:wrap; gap:4px; min-width:210px; max-width:330px; }

/* 毛利與簽核路徑 */
.q-pv { margin-top:14px; border:1px solid #c5cae9; background:#f6f8ff; border-radius:10px; padding:12px 14px; }
.q-pv-h { font-size:13px; font-weight:700; color:#1a2d52; margin-bottom:8px; }
.q-pv-grid { display:grid; grid-template-columns:repeat(auto-fit, minmax(150px, 1fr)); gap:10px; }
.q-pv-cell { background:#fff; border-radius:8px; padding:8px 12px; box-shadow:0 1px 3px rgba(0,0,0,.06); }
.q-pv-cell .k { font-size:12px; color:#8a94a3; margin-bottom:2px; }
.q-pv-cell .v { font-size:15px; font-weight:700; color:#1a2d52; }
.q-path { display:flex; flex-wrap:wrap; align-items:center; gap:6px; margin-top:10px; }
.q-chip { font-size:12px; padding:3px 10px; border-radius:12px; background:#e8f0fe; color:#174ea6; font-weight:600; }
.q-chip.dim { background:#eef1f6; color:#5f6b7a; font-weight:400; }
.q-arrow { color:#8a94a3; font-size:12px; }
.q-pv ul { margin:8px 0 0 18px; padding:0; font-size:12.5px; line-height:1.7; color:#444; }

/* 自訂對話框（取代 alert/confirm） */
.q-dlg { width:480px; max-width:96vw; }
.q-dlg.q-dlg-wide { width:620px; }
.q-dlg-body { font-size:13.5px; line-height:1.7; color:#333; }
.q-dlg-body ul { margin:6px 0 0 18px; padding:0; }
.q-dlg-body .q-kv { display:flex; gap:8px; margin:3px 0; }
.q-dlg-body .q-kv .k { color:#8a94a3; min-width:5em; flex-shrink:0; }

body.dark .q-ro-state { background:#161b22; border-color:#30363d; color:#c9d1d9; }
body.dark .q-ro-state .q-sub { color:#8b949e; }
body.dark .q-hint, body.dark .q-sec-sub, body.dark .q-legacy { color:#8b949e; }
body.dark .q-infobar { background:#0d2040; border-color:#1c3a5f; color:#8ecfff; }
body.dark .q-warnbar { background:#2a2000; border-color:#5a4000; color:#d4a84e; }
body.dark .q-errbar { background:#2a0d0d; border-color:#5a1010; color:#ff8080; }
body.dark .q-costby { background:#161b22; border-color:#30363d; }
body.dark .qp-search { background:#0d1117; border-color:#30363d; color:#c9d1d9; }
body.dark .qp-count { color:#8b949e; }
body.dark .qp-chip { background:#0d2040; color:#8ecfff; }
body.dark .qp-list { background:#0d1117; border-color:#30363d; }
body.dark .qp-grp-h { background:#21262d; color:#c9d1d9; border-bottom-color:#30363d; }
body.dark .qp-item { border-bottom-color:#21262d; color:#c9d1d9; }
body.dark .qp-item:hover { background:#161b22; }
body.dark .qp-tag { background:#21262d; color:#8b949e; }
body.dark .qp-tag.qp-un { background:#2a2000; color:#d4a84e; }
body.dark .qp-tag.qp-self { background:#0a1f10; color:#4ade80; }
body.dark .qp-empty { color:#8b949e; }
body.dark .quote-st-ready { background:#0d2040; color:#58a6ff; }
body.dark .quote-st-wait { background:#2a2000; color:#d4a84e; }
body.dark .quote-st-pending { background:#0d2040; color:#8ecfff; }
body.dark .quote-st-returned { background:#2a0d0d; color:#ff8080; }
body.dark .quote-st-approved { background:#0a1f10; color:#4ade80; }
body.dark .quote-st-void { background:#2a1500; color:#f0a050; border-color:#5a3000; }
body.dark .q-cost { color:#8b949e; }
body.dark .q-cost.warn { color:#d4a84e; }
body.dark .q-cost.ok { color:#4ade80; }
body.dark .q-pv { background:#161b22; border-color:#30363d; }
body.dark .q-pv-h { color:#e6edf3; }
body.dark .q-pv-cell { background:#0d1117; box-shadow:none; }
body.dark .q-pv-cell .v { color:#e6edf3; }
body.dark .q-chip { background:#0d2040; color:#8ecfff; }
body.dark .q-chip.dim { background:#21262d; color:#8b949e; }
body.dark .q-pv ul { color:#c9d1d9; }
body.dark .q-dlg-body { color:#c9d1d9; }

@media (max-width: 620px) {
  .q-actions { min-width:0; max-width:none; }
  .qp-list { max-height:200px; }
}
`;

function _qEnsureStyle() {
  if (document.getElementById('quoteExtraStyle')) return;
  const st = document.createElement('style');
  st.id = 'quoteExtraStyle';
  st.textContent = QUOTE_CSS;
  document.head.appendChild(st);
}
_qEnsureStyle();

/**
 * 計算報價合計（僅供畫面即時顯示；簽核用的毛利/金額一律以伺服器試算為準）
 * discountType: 'none' | 'percent' | 'amount'
 * discountValue: percent 時為百分比（90 = 九折）；amount 時為議價後未稅金額
 */
function quoteTotal(items, discountType, discountValue) {
  const sub = (items || []).reduce(
    (s, it) => s + (parseFloat(it.qty) || 0) * (parseFloat(it.unitPrice) || 0), 0
  );
  let discounted = sub;
  let discountAmt = 0;
  if (discountType === 'percent') {
    const pct = parseFloat(discountValue) || 100; // 百分比，如 90 = 九折
    discounted  = sub * pct / 100;
    discountAmt = sub - discounted;
  } else if (discountType === 'amount') {
    discounted  = parseFloat(discountValue) || sub;
    discountAmt = sub - discounted;
  }
  const tax   = discounted * 0.05;
  const total = discounted + tax;
  return { sub, discounted, discountAmt, tax, total };
}

/** 台灣（Asia/Taipei）今天，格式 YYYY-MM-DD。不可用 toISOString()：那是 UTC，台灣凌晨 0~8 點會差一天 */
function taipeiTodayClient() {
  return new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
}

/**
 * 報價期限預設值：dateStr（YYYY-MM-DD）當月的最後一個工作天；dateStr 已在當月最後工作天之後（例如月底週末或假日建單）
 * 就順延到下個月的最後一個工作天，避免預設期限早於報價日期。
 * 規則與 lib/quoteExcel.js 的 defaultValidUntil 相同（改一邊要改另一邊；單元測試會逐月比對兩邊）：週一至週五，且不是固定日期國定假日
 * 元旦 1/1、和平紀念日 2/28、勞動節 5/1、孔子誕辰紀念日／教師節 9/28（遇週六補前一個週五、遇週日補後一個週一）。
 * 春節、清明、端午、中秋日期逐年變動，未內建（2026~2030 年這些連假不影響任何月底）→ 之後月底剛好落在這些連假時請自行修改。
 */
function quoteLastWorkingDay(dateStr) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(dateStr || ''));
  if (!m) return '';
  const y0 = +m[1], mo0 = +m[2], dd0 = +m[3];
  const chk = new Date(Date.UTC(y0, mo0 - 1, dd0));
  if (chk.getUTCFullYear() !== y0 || chk.getUTCMonth() !== mo0 - 1 || chk.getUTCDate() !== dd0) return '';   // 必須是真實存在的日期
  const iso = function (d) { return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0'); };
  const lwd = function (y, mo) {
    const hol = new Set();
    [y, y + 1].forEach(function (yy) {          // 含隔年：隔年 1/1 遇週六會補到今年 12/31
      [[1, 1], [2, 28], [5, 1], [9, 28]].forEach(function (md) {
        const dt = new Date(Date.UTC(yy, md[0] - 1, md[1])), dow = dt.getUTCDay();
        if (dow === 6) dt.setUTCDate(dt.getUTCDate() - 1); else if (dow === 0) dt.setUTCDate(dt.getUTCDate() + 1);
        hol.add(iso(dt));
      });
    });
    const d = new Date(Date.UTC(y, mo, 0));   // 當月最後一天
    while (d.getUTCDay() === 0 || d.getUTCDay() === 6 || hol.has(iso(d))) d.setUTCDate(d.getUTCDate() - 1);
    return iso(d);
  };
  let r = lwd(y0, mo0);
  if (r < dateStr) { let y = y0, mo = mo0 + 1; if (mo === 13) { mo = 1; y += 1; } r = lwd(y, mo); }
  return r;
}

function fmtMoney(n) {
  return 'NT$ ' + Math.round(n).toLocaleString();
}

function _qFmtTime(t) {
  if (!t) return '';
  try {
    const d = new Date(t);
    if (isNaN(d.getTime())) return String(t);
    return d.toLocaleString('zh-TW', { hour12: false, timeZone: 'Asia/Taipei' });
  } catch (e) { return String(t); }
}

/** 毛利率文字：兩位小數、向 0 截斷（與伺服器 marginText 同規則，避免畫面 15.00 卻簽到董事長） */
function _qTruncPct(gp, rev) {
  const revC = Math.round(rev * 100), gpC = Math.round(gp * 100);
  if (!(revC > 0)) return null;
  return (Math.trunc(gpC * 10000 / revC) / 100).toFixed(2);
}

// ════════════════════════════════════════════════════════════
// ── 設定 / 待處理 / 單張報價單 ─────────────────────────────────
// ════════════════════════════════════════════════════════════
async function loadQuoteConfig() {
  try {
    const r = await fetch(`${API}/quote-approval/config`);
    if (r.ok) _qCfg = await r.json();
  } catch (e) { /* 失敗就沿用舊設定；沒有設定時表單仍可開啟，只是商品清單為空 */ }
  return _qCfg;
}
function quoteCfg() { return _qCfg || {}; }

function _qSetInboxBadge(n) {
  const b = document.getElementById('quoteInboxBadge');
  if (!b) return;
  b.textContent = String(n);
  b.style.display = n > 0 ? '' : 'none';
}

/** 取得待我處理清單，同步更新徽章；失敗則 _qInboxIds 設 null（篩選退回用 perm 推算） */
async function _qLoadInbox() {
  try {
    const r = await fetch(`${API}/quotations/inbox`);
    if (!r.ok) throw new Error('inbox');
    const j = await r.json();
    const ids = new Set();
    (j.approvals || []).forEach(x => { if (x && x.id) ids.add(x.id); });
    (j.costs || []).forEach(x => { if (x && x.id) ids.add(x.id); });
    _qInboxIds = ids;
    const c = j.counts || {};
    const n = (typeof c.approvals === 'number' || typeof c.costs === 'number')
      ? ((c.approvals || 0) + (c.costs || 0)) : ids.size;
    _qSetInboxBadge(n);
  } catch (e) {
    _qInboxIds = null;
  }
}

async function _qFetchQuote(id) {
  try {
    const r = await fetch(`${API}/quotations/${encodeURIComponent(id)}`);
    if (!r.ok) return null;
    const q = await r.json();
    const i = allQuotations.findIndex(x => x.id === id);
    if (i >= 0) allQuotations[i] = q;
    return q;
  } catch (e) { return null; }
}

function _qFind(id) { return (allQuotations || []).find(x => x.id === id) || null; }

/** 呼叫 W4（quote-approval.js）提供的全域函式；未載入時提示 */
function _qCallApproval(fnName, arg) {
  const fn = window[fnName];
  if (typeof fn !== 'function') { showToast('簽核模組尚未載入，請重新整理頁面後再試'); return; }
  return fn(arg);
}
function qOpenApproval(id) { return _qCallApproval('openQuoteApproval', id); }
function qOpenCostFill(id) { return _qCallApproval('openQuoteCostFill', id); }
function openQuoteApprovalSettingsSafe() { return _qCallApproval('openQuoteApprovalSettings'); }

// ════════════════════════════════════════════════════════════
// ── 自訂對話框（取代 alert/confirm/prompt）─────────────────────
// opts: { title, message(純文字), bodyHtml(已轉義的 HTML), buttons:[{text,value,cls}], wide, dismissValue }
// 回傳 Promise：按下按鈕的 value；點背景/X/Esc 回 dismissValue（預設 false）
// ════════════════════════════════════════════════════════════
let _qDialogCancel = null;   // 目前開著的對話框的「以 dismiss 結束」函式：開新的之前先結束舊的，避免舊 Promise 永不 resolve、keydown 監聽殘留
function qDialog(opts) {
  return new Promise(function (resolve) {
    if (_qDialogCancel) { try { _qDialogCancel(); } catch (e) { /* 忽略 */ } }
    const old = document.getElementById('qDialogOverlay');
    if (old) old.remove();
    const dismiss = opts.dismissValue === undefined ? false : opts.dismissValue;
    const buttons = opts.buttons || [{ text: '確定', value: true, cls: 'btn-primary' }];
    const msgHtml = opts.message
      ? '<div>' + String(opts.message).split('\n').map(escapeHtml).join('<br>') + '</div>'
      : '';
    const ov = document.createElement('div');
    ov.className = 'modal-overlay open';
    ov.id = 'qDialogOverlay';
    ov.style.zIndex = '120';
    ov.innerHTML = `
      <div class="modal q-dlg${opts.wide ? ' q-dlg-wide' : ''}" role="dialog" aria-modal="true">
        <div class="modal-header">
          <h2>${escapeHtml(opts.title || '')}</h2>
          <button type="button" class="modal-close" data-qd="x">&#10005;</button>
        </div>
        <div class="modal-body q-dlg-body">${msgHtml}${opts.bodyHtml || ''}</div>
        <div class="modal-footer">
          ${buttons.map((b, i) => `<button type="button" class="btn ${escapeHtml(b.cls || 'btn-secondary')}" data-qd="${i}">${escapeHtml(b.text)}</button>`).join('')}
        </div>
      </div>`;
    document.body.appendChild(ov);
    let done = false;
    function finish(v) {
      if (done) return;
      done = true;
      document.removeEventListener('keydown', onKey, true);
      if (_qDialogCancel === cancelMe) _qDialogCancel = null;
      ov.remove();
      resolve(v);
    }
    const cancelMe = function () { finish(dismiss); };
    _qDialogCancel = cancelMe;
    function onKey(e) {
      if (e.key === 'Escape') { e.stopPropagation(); finish(dismiss); }
    }
    document.addEventListener('keydown', onKey, true);
    ov.addEventListener('mousedown', function (e) { if (e.target === ov) finish(dismiss); });
    ov.addEventListener('click', function (e) {
      const b = e.target.closest('[data-qd]');
      if (!b) return;
      const k = b.getAttribute('data-qd');
      finish(k === 'x' ? dismiss : buttons[parseInt(k, 10)].value);
    });
    const first = ov.querySelector('.modal-footer .btn-primary, .modal-footer .btn-danger');
    if (first) first.focus();
  });
}

function qConfirm(title, message, okText, danger) {
  return qDialog({
    title: title, message: message,
    buttons: [
      { text: okText || '確定', value: true, cls: danger ? 'btn-danger' : 'btn-primary' },
      { text: '取消', value: false, cls: 'btn-secondary' },
    ],
  });
}

// ════════════════════════════════════════════════════════════
// ── 簽核／成本狀態標籤 ─────────────────────────────────────────
// ════════════════════════════════════════════════════════════
function _qApprovalValid(q) {
  const a = q && q.approval;
  return !!(a && a.state === 'approved' && a.valid === true);
}

function _qLastHistory(q) {
  const h = q && q.approval && q.approval.history;
  return (h && h.length) ? h[h.length - 1] : null;
}

/** 這張單是否送過簽（送過簽的單不可刪除） */
function _qEverSubmitted(q) {
  const a = q && q.approval;
  if (!a) return false;
  if (a.state && a.state !== 'none') return true;
  if (a.submittedAt) return true;
  return (a.history || []).some(h => h && (h.action === 'SUBMIT' || h.action === 'APPROVE'));
}

/** 簽核階段：{ key, cls, label, title } */
function quoteStageInfo(q) {
  const a = q.approval;
  const cf = q.costFlow || {};
  if (a && a.state === 'approved') {
    return a.valid === false
      ? { key: 'void', cls: 'quote-st-void', label: '核准已失效', title: '核准後內容被修改，需重新送簽' }
      : { key: 'approved', cls: 'quote-st-approved', label: '已核准', title: '' };
  }
  if (a && a.state === 'pending') {
    const steps = a.steps || [];
    const cur = Math.max(0, parseInt(a.cur, 10) || 0);
    const st = steps[cur];
    const who = st ? (st.assigneeName || st.label || QUOTE_TIER_LABEL[st.tier] || '') : '';
    const title = steps.map((s, i) => `${i + 1}. ${s.label || QUOTE_TIER_LABEL[s.tier] || ''}${s.assigneeName ? '：' + s.assigneeName : ''}`).join('\n');
    return { key: 'pending', cls: 'quote-st-pending', label: `簽核中（第 ${cur + 1} 關：${who}）`, title: title };
  }
  if (a && a.state === 'returned') {
    return { key: 'returned', cls: 'quote-st-returned', label: '已駁回', title: '請修改後重新送簽' };
  }
  const last = _qLastHistory(q);
  if (last && last.action === 'INVALIDATE') {
    return { key: 'void', cls: 'quote-st-void', label: '核准已作廢', title: '已核准的內容被修改，需重新送簽' };
  }
  if (cf.state === 'requested') return { key: 'cost', cls: 'quote-st-wait', label: '待顧問填成本', title: '' };
  if (cf.state === 'filled') return { key: 'ready', cls: 'quote-st-ready', label: '成本已填，待送簽', title: '' };
  return { key: 'draft', cls: 'quote-st-draft', label: '草稿', title: '' };
}

function _qStageMatches(q, filter) {
  const k = quoteStageInfo(q).key;
  if (filter === 'draft') return k === 'draft' || k === 'ready';
  return k === filter;
}

function quoteCostCell(q) {
  const cf = q.costFlow;
  if (!cf || !q.products || !q.products.length) return '<span class="q-cost">—</span>';
  const who = escapeHtml(cf.byName || q.costByName || '');
  switch (cf.state) {
    case 'na':        return '<span class="q-cost">業務自填</span>';
    case 'needed':    return '<span class="q-cost warn">待指派顧問</span>';
    case 'requested': return `<span class="q-cost warn">待 ${who || '顧問'} 填寫</span>`;
    case 'filled':    return `<span class="q-cost ok">已填${who ? '（' + who + '）' : ''}</span>`;
    default:          return '<span class="q-cost">—</span>';
  }
}

/** 需要我處理？優先用 inbox；inbox 取不到時用 perm 推算 */
function _qIsMine(q) {
  if (_qInboxIds) return _qInboxIds.has(q.id);
  const p = q.perm || {};
  const cf = q.costFlow || {};
  return !!(p.canApprove || (p.isCostProvider && p.canEditCost && cf.state === 'requested'));
}

// ════════════════════════════════════════════════════════════
// ── 報價單清單 ────────────────────────────────────────────────
// ════════════════════════════════════════════════════════════
const _qBusy = new Set();   // 進行中的列表動作（key = id:action），防雙擊重複觸發
let _qLoadError = '';   // 列表載入失敗的原因（空字串＝成功）；失敗時保留舊列表並顯示可重試的提示，不能當成「沒有報價單」
async function loadQuotationsView() {
  _qEnsureStyle();
  const listP = (async function () {
    try {
      const r = await fetch(`${API}/quotations`);
      if (r.status === 401) { window.location.href = '/login.html'; return { error: '登入已逾時，請重新登入' }; }
      if (!r.ok) {
        const e = await r.json().catch(function () { return {}; });
        return { error: (e && e.error) || ('載入失敗（HTTP ' + r.status + '）') };
      }
      return await r.json();
    } catch (e) { return { error: '網路連線失敗，請檢查連線後重試' }; }
  })();
  const [list] = await Promise.all([listP, loadQuoteConfig(), _qLoadInbox()]);
  if (Array.isArray(list)) { allQuotations = list; _qLoadError = ''; }
  else { _qLoadError = (list && list.error) || '載入失敗'; }
  // 確保 KA 資料已載入（公司名前綴 ⭐ 用）
  if (typeof loadKeyAccounts === 'function' && typeof allKeyAccounts !== 'undefined' && Array.isArray(allKeyAccounts) && allKeyAccounts.length === 0) {
    try { await loadKeyAccounts(); } catch (e) { /* KA 標記失敗不影響列表 */ }
  }
  applyQuoteToolbar();
  renderQuoteList();
  bindQuoteListHandlers();
}

/** 依登入者身分顯示「簽核設定」按鈕；待我處理按鈕狀態 */
function applyQuoteToolbar() {
  const me = quoteCfg().me || {};
  const sb = document.getElementById('quoteSettingsBtn');
  if (sb) {
    const show = !!(me.isAdmin || me.isSealManager);
    sb.style.display = show ? '' : 'none';
    sb.innerHTML = me.isAdmin ? '&#9881; 簽核設定' : '&#128278; 報價章設定';
  }
  const ib = document.getElementById('quoteInboxBtn');
  if (ib) ib.classList.toggle('q-on', _qInboxOnly);
}

function _qActionButtons(q) {
  const p = q.perm || {};
  const cfg = quoteCfg();
  const me = cfg.me || {};
  const id = escapeHtml(q.id);
  const b = (act, label, cls, title) =>
    `<button type="button" class="btn btn-sm ${cls || ''}" data-qact="${act}" data-id="${id}"${title ? ` title="${escapeHtml(title)}"` : ''}>${label}</button>`;
  const out = [];
  const a = q.approval;
  if (p.canEdit)     out.push(b('edit', '✏️ 編輯'));
  if (p.canSubmit)   out.push(b('submit', '📤 送簽', 'btn-primary', '送交主管簽核；送簽後整張單鎖定'));
  if (p.canWithdraw) out.push(b('withdraw', '↩ 撤回', '', '撤回後可修改，已簽的關卡作廢'));
  if (p.canApprove || p.canReturn) {
    out.push(b('approve', '✍ 簽核', 'btn-primary', '開啟簽核面板'));
  } else if (a && (a.state !== 'none' || (a.history || []).length || p.canReassign)) {
    out.push(b('approve', '📜 簽核進度', '', '檢視簽核進度與歷程'));
  }
  if (p.isCostProvider && p.canEditCost) out.push(b('cost', '💲 填成本', 'btn-primary', '填寫各品項成本'));
  if (p.canSeePrice !== false) {
    out.push(b('preview', '👁 預覽', '', '下載前先預覽報價單長相'));
    out.push(b('export', '&#11015; Excel', 'btn-export', _qApprovalValid(q) ? '下載給客戶的報價單（已核准，會蓋報價專用章）' : '下載報價單（尚未核准，不會有報價專用章）'));
  }
  // 毛利分析含報價單價與毛利：只負責填成本的顧問（看不到價格）不顯示，伺服器端也會擋
  if (p.canSeeCost && p.canSeePrice !== false) out.push(b('pnl', '&#11015; 毛利分析(內部)', '', '含成本與毛利率，僅限內部使用，請勿提供客戶'));
  if ((p.isOwner || me.isAdmin) && !_qEverSubmitted(q) && p.isCostProvider !== true) out.push(b('delete', '🗑️', 'btn-soft-danger', '刪除（送過簽的單不可刪除）'));
  return `<div class="q-actions">${out.join('')}</div>`;
}

function renderQuoteList() {
  const search   = ($('quoteSearchInput')  ? $('quoteSearchInput').value   : '').toLowerCase();
  const stFilter = ($('quoteStatusFilter') ? $('quoteStatusFilter').value  : '');
  let list = allQuotations;
  if (search)   list = list.filter(q =>
    (q.quoteNo     || '').toLowerCase().includes(search) ||
    (q.company     || '').toLowerCase().includes(search) ||
    (q.contactName || '').toLowerCase().includes(search) ||
    (q.projectName || '').toLowerCase().includes(search) ||
    (q.ownerName   || '').toLowerCase().includes(search)
  );
  if (_qInboxOnly) list = list.filter(_qIsMine);
  if (stFilter)    list = list.filter(q => _qStageMatches(q, stFilter));

  const tbody = $('quoteTbody');
  if (!tbody) return;
  const errRow = _qLoadError
    ? `<tr><td colspan="10" style="text-align:center;padding:12px;background:#fff4f4;color:#b3261e">⚠ ${escapeHtml(_qLoadError)}${allQuotations.length ? '（以下為上次載入的資料，可能不是最新）' : ''}　<button type="button" class="btn btn-sm btn-primary" data-qact="reload">重新載入</button></td></tr>`
    : '';
  if (!list.length) {
    const msg = _qLoadError ? '' : (_qInboxOnly ? '目前沒有需要你處理的報價單' : '尚無報價單資料，點擊「新增報價單」開始建立');
    tbody.innerHTML = errRow + (msg ? `<tr><td colspan="10" class="empty-msg" style="text-align:center;padding:32px">${escapeHtml(msg)}</td></tr>` : '');
    return;
  }
  tbody.innerHTML = errRow + list.map(q => {
    const p = q.perm || {};
    const st = quoteStageInfo(q);
    const totalCell = (p.canSeePrice === false)
      ? '<span style="color:#8a94a3">—</span>'
      : fmtMoney(quoteTotal(q.items, q.discountType, q.discountValue).total);
    const legacy = QUOTE_LEGACY_LABEL[q.status]
      ? `<span class="q-legacy">舊狀態：${escapeHtml(QUOTE_LEGACY_LABEL[q.status])}</span>` : '';
    return `<tr>
      <td><span class="quote-no">${escapeHtml(q.quoteNo || '')}</span></td>
      <td>${typeof kaCompanyMark === 'function' ? kaCompanyMark(q.company) : ''}${escapeHtml(q.company || '')}</td>
      <td>${escapeHtml(q.contactName || '')}</td>
      <td>${escapeHtml(q.projectName || '')}</td>
      <td>${escapeHtml(q.ownerName || q.owner || '')}</td>
      <td>${escapeHtml(q.quoteDate || '')}</td>
      <td style="text-align:right;font-weight:600;font-size:13px">${totalCell}</td>
      <td><span class="quote-status q-wrap ${escapeHtml(st.cls)}"${st.title ? ` title="${escapeHtml(st.title)}"` : ''}>${escapeHtml(st.label)}</span>${legacy}</td>
      <td>${quoteCostCell(q)}</td>
      <td>${_qActionButtons(q)}</td>
    </tr>`;
  }).join('');
}

function bindQuoteListHandlers() {
  const btn = $('addQuoteBtn');
  if (btn && !btn._qListBound) {
    btn._qListBound = true;
    btn.addEventListener('click', function() { openQuoteModal(null); });
  }
  const si = $('quoteSearchInput');
  if (si && !si._qListBound) {
    si._qListBound = true;
    let _qSearchTimer = null;   // 筆數多時每個按鍵都重畫整張表會卡，稍微延遲
    si.addEventListener('input', function () { clearTimeout(_qSearchTimer); _qSearchTimer = setTimeout(renderQuoteList, 150); });
  }
  const sf = $('quoteStatusFilter');
  if (sf && !sf._qListBound) {
    sf._qListBound = true;
    sf.addEventListener('change', renderQuoteList);
  }
  const ib = $('quoteInboxBtn');
  if (ib && !ib._qListBound) {
    ib._qListBound = true;
    ib.addEventListener('click', function() {
      _qInboxOnly = !_qInboxOnly;
      ib.classList.toggle('q-on', _qInboxOnly);
      renderQuoteList();
    });
  }
  // 列表操作：事件委派（不把 id 拼進 inline JS，避免跳脫問題）
  const tb = $('quoteTbody');
  if (tb && !tb._qListBound) {
    tb._qListBound = true;
    tb.addEventListener('click', function(e) {
      const el = e.target.closest('[data-qact]');
      if (!el) return;
      const id = el.getAttribute('data-id');
      const act = el.getAttribute('data-qact');
      if (act === 'reload') return loadQuotationsView();
      const q = _qFind(id);
      // 同一張單的同一個動作進行中就忽略重複點擊（雙擊「送簽」不該送兩次、開兩個確認框）
      const key = id + ':' + act;
      if (_qBusy.has(key)) return;
      _qBusy.add(key);
      const done = function () { _qBusy.delete(key); };
      let p;
      switch (act) {
        case 'edit':     p = openQuoteModal(id); break;
        case 'submit':   p = submitQuote(id); break;
        case 'withdraw': p = withdrawQuote(id); break;
        case 'approve':  p = qOpenApproval(id); break;
        case 'cost':     p = qOpenCostFill(id); break;
        case 'preview':  p = previewQuote(id); break;
        case 'export':   p = exportQuote(id, q ? q.quoteNo : ''); break;
        case 'pnl':      p = exportQuote(id, q ? q.quoteNo : '', 'pnl'); break;
        case 'delete':   p = deleteQuote(id); break;
      }
      Promise.resolve(p).then(done, done);
    });
  }
}

// ── 匯出 Excel ──────────────────────────────────────────────
async function exportQuote(id, quoteNo, kind) {
  const isPnl = kind === 'pnl';
  if (!isPnl) {
    // 下載前確認最新簽核狀態；未核准或核准已失效 → 警示（自訂 modal，不用 confirm）
    const q = (await _qFetchQuote(id)) || _qFind(id);
    if (q && !_qApprovalValid(q)) {
      const voided = q.approval && q.approval.state === 'approved' && q.approval.valid === false;
      const ok = await qDialog({
        title: '尚未簽核完成',
        message: (voided ? '（此報價單核准後內容已被修改，原核准已失效。）\n' : '') +
          '此報價單主管尚未簽核完成，下載的檔案不會有報價專用章，不可視為正式報價單。仍要下載？',
        buttons: [
          { text: '仍要下載', value: true, cls: 'btn-primary' },
          { text: '取消', value: false, cls: 'btn-secondary' },
        ],
      });
      if (!ok) return;
    }
  }
  try {
    showToast(isPnl ? '正在產生毛利分析…' : '正在產生報價單…');
    const r = await fetch(`${API}/quotations/${encodeURIComponent(id)}/${isPnl ? 'export-pnl' : 'export'}`);
    if (!r.ok) {
      const e = await r.json().catch(() => ({}));
      return showToast(e.error || '匯出失敗');
    }
    const approvedFile = r.headers.get('X-Quote-Approval') === 'approved';
    const sealState = r.headers.get('X-Quote-Seal') || (approvedFile ? 'applied' : 'none');   // applied 已蓋章／missing 已核准但章缺漏／none 未核准
    const stamped = sealState === 'applied';
    const blob = await r.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `${quoteNo || 'quotation'}${isPnl ? '_毛利分析-內部' : ''}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    showToast(isPnl ? '毛利分析（內部用）已下載，請勿提供客戶'
      : (stamped ? '報價單 Excel 已下載（已蓋報價專用章）'
        : (approvedFile && sealState === 'missing'
          ? '⚠ 此單已核准，但系統尚未上傳有效的報價專用章，下載的檔案沒有章，請聯絡管理部秘書'
          : '報價單 Excel 已下載（未蓋報價專用章）')));
  } catch(e) {
    showToast('匯出失敗，請重試');
  }
}

// ── 刪除報價單 ──────────────────────────────────────────────
async function deleteQuote(id) {
  const ok = await qConfirm('刪除報價單', '確定要刪除此報價單嗎？此動作無法復原。', '確認刪除', true);
  if (!ok) return;
  try {
    const r = await fetch(`${API}/quotations/${encodeURIComponent(id)}`, { method: 'DELETE' });
    if (!r.ok) {
      const e = await r.json().catch(() => ({}));
      return showToast(e.error || '刪除失敗');
    }
    showToast('已刪除報價單');
    loadQuotationsView();
  } catch(e) {
    showToast('刪除失敗，請重試');
  }
}

// ════════════════════════════════════════════════════════════
// ── 送簽 / 撤回 ───────────────────────────────────────────────
// ════════════════════════════════════════════════════════════
/** preview.tiers 的元素可能是 'mgr1' 這類 key，也可能是 {tier,label}；統一轉成顯示名稱 */
function _qTierName(t) {
  if (!t) return '';
  if (typeof t === 'string') return QUOTE_TIER_LABEL[t] || t;
  return t.label || QUOTE_TIER_LABEL[t.tier || t.key] || String(t.tier || t.key || '');
}

function _qPathList(pv) {
  if (!pv) return [];
  if (Array.isArray(pv.tiers) && pv.tiers.length) return pv.tiers.map(_qTierName).filter(Boolean);
  if (pv.board) return [QUOTE_TIER_LABEL.mgr1, QUOTE_TIER_LABEL.gm, QUOTE_TIER_LABEL.board];
  const lv = parseInt(pv.level, 10);
  if (lv === 1) return [QUOTE_TIER_LABEL.mgr1];
  if (lv === 2) return [QUOTE_TIER_LABEL.mgr1, QUOTE_TIER_LABEL.gm];
  if (lv === 3) return [QUOTE_TIER_LABEL.mgr1, QUOTE_TIER_LABEL.gm, QUOTE_TIER_LABEL.chairman];
  return [];
}

function _qPathChips(pv) {
  const names = _qPathList(pv);
  if (!names.length) return '';
  return '<div class="q-path"><span class="q-chip dim">業務送簽</span>' +
    names.map(n => `<span class="q-arrow">→</span><span class="q-chip">${escapeHtml(n)}</span>`).join('') + '</div>';
}

function _qListHtml(arr) {
  const a = (arr || []).filter(Boolean);
  return a.length ? '<ul>' + a.map(t => `<li>${escapeHtml(t)}</li>`).join('') + '</ul>' : '';
}

async function _qShowBlockers(blockers) {
  const msgs = (blockers || []).map(b => (b && (b.message || b.code)) || String(b));
  await qDialog({
    title: '目前無法送簽',
    bodyHtml: '<div>請先處理以下事項：</div>' + _qListHtml(msgs),
    buttons: [{ text: '知道了', value: true, cls: 'btn-primary' }],
  });
}

async function submitQuote(id) {
  const q = (await _qFetchQuote(id)) || _qFind(id);
  if (!q) return showToast('找不到此報價單');
  const pv = q.preview || {};
  if (pv.blockers && pv.blockers.length) { await _qShowBlockers(pv.blockers); return; }

  const margin = pv.marginText != null ? `${escapeHtml(pv.marginText)}%` : '—';
  const body =
    `<div class="q-kv"><span class="k">報價單</span><span>${escapeHtml(q.quoteNo || '')}　${escapeHtml(q.company || '')}</span></div>` +
    (q.projectName ? `<div class="q-kv"><span class="k">專案</span><span>${escapeHtml(q.projectName)}</span></div>` : '') +
    `<div class="q-kv"><span class="k">類別</span><span>${escapeHtml(pv.rowLabel || '')}</span></div>` +
    `<div class="q-kv"><span class="k">毛利率</span><span><b>${margin}</b>（折扣後未稅）</span></div>` +
    `<div class="q-kv"><span class="k">簽核路徑</span><span></span></div>` + _qPathChips(pv) +
    _qListHtml(pv.reasons) +
    ((pv.warnings || []).length ? `<div class="q-warnbar">${(pv.warnings || []).map(escapeHtml).join('<br>')}</div>` : '') +
    '<div class="q-hint" style="margin-top:10px">送簽後整張報價單會鎖定；要修改請先「撤回」。</div>';
  const ok = await qDialog({
    title: '確認送簽', bodyHtml: body, wide: true,
    buttons: [
      { text: '確認送簽', value: true, cls: 'btn-primary' },
      { text: '取消', value: false, cls: 'btn-secondary' },
    ],
  });
  if (!ok) return;
  try {
    const r = await fetch(`${API}/quotations/${encodeURIComponent(id)}/submit`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}',
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) {
      if (j.code === 'CANNOT_SUBMIT' && j.blockers) { await _qShowBlockers(j.blockers); }
      else showToast(j.error || '送簽失敗');
      loadQuotationsView();
      return;
    }
    const steps = (j.approval && j.approval.steps) || [];
    const first = steps[0];
    showToast('已送簽' + (first ? `，等待${first.assigneeName || first.label || ''}簽核` : '') + ' ✅');
    loadQuotationsView();
  } catch (e) {
    showToast('送簽失敗，請重試');
  }
}

async function withdrawQuote(id) {
  const ok = await qConfirm('撤回報價單', '撤回後報價單回到未送簽狀態，已簽的關卡會作廢（歷程保留）。\n要修改內容後再重新送簽，確定撤回？', '確認撤回', true);
  if (!ok) return;
  try {
    const r = await fetch(`${API}/quotations/${encodeURIComponent(id)}/withdraw`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}',
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) { showToast(j.error || '撤回失敗'); loadQuotationsView(); return; }
    showToast('已撤回，可修改後重新送簽');
    loadQuotationsView();
  } catch (e) {
    showToast('撤回失敗，請重試');
  }
}

// ════════════════════════════════════════════════════════════
// ── 商品（必勾）與支援顧問 ─────────────────────────────────────
// ════════════════════════════════════════════════════════════
function readSelectedProducts() {
  return Array.from(document.querySelectorAll('#qProductsList .qp-cb:checked')).map(cb => cb.value);
}

/** 勾選商品中是否有任一項需要顧問填成本（未歸類視為需要） */
function quoteNeedsConsultant(names) {
  const pc = quoteCfg().productClasses || {};
  return (names || []).length > 0 && names.some(n => !(pc[n] && pc[n].costBySales === true));
}

function _qUnclassified(names) {
  const pc = quoteCfg().productClasses || {};
  return (names || []).filter(n => !(pc[n] && pc[n].cls));
}

function _qProdTag(name) {
  const cfg = quoteCfg();
  const c = (cfg.productClasses || {})[name];
  if (!c || !c.cls) return '<span class="qp-tag qp-un" title="尚未歸類，簽核時視為「其他」並需顧問填成本">未歸類</span>';
  const lab = (cfg.classLabels || {})[c.cls] || c.cls;
  return `<span class="qp-tag">${escapeHtml(lab)}</span>` + (c.costBySales ? '<span class="qp-tag qp-self">業務填成本</span>' : '');
}

function renderQuoteProducts(selected) {
  const cfg = quoteCfg();
  const catalog = Array.isArray(cfg.catalog) ? cfg.catalog : [];
  const sel = new Set(selected || []);
  const known = new Set(catalog.map(c => c.name));
  const extra = (selected || []).filter(n => !known.has(n));
  const groups = new Map();
  catalog.forEach(c => {
    if (!c || !c.name) return;
    const key = [c.bu, c.group].filter(Boolean).join(' › ') || '商品';
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(c.name);
  });
  if (extra.length) groups.set('目錄外商品（可能已停用或改名）', extra);

  const box = document.getElementById('qProductsList');
  if (!groups.size) {
    box.innerHTML = `<div class="qp-empty">${_qCfg ? '商品目錄是空的，請管理員先到後台「商品目錄」設定' : '無法載入商品目錄，請重新整理頁面後再試'}</div>`;
  } else {
    box.innerHTML = Array.from(groups.entries()).map(([label, names]) =>
      `<div class="qp-grp"><div class="qp-grp-h">${escapeHtml(label)}</div>` +
      names.map(n => `<label class="qp-item" data-s="${escapeHtml(n.toLowerCase())}">` +
        `<input type="checkbox" class="qp-cb" value="${escapeHtml(n)}"${sel.has(n) ? ' checked' : ''}>` +
        `<span class="qp-nm">${escapeHtml(n)}</span>${_qProdTag(n)}</label>`).join('') +
      '</div>').join('');
  }
  const s = document.getElementById('qProdSearch');
  if (s) s.value = '';
}

function filterQuoteProducts() {
  const kw = (document.getElementById('qProdSearch').value || '').trim().toLowerCase();
  document.querySelectorAll('#qProductsList .qp-grp').forEach(function (g) {
    let any = false;
    g.querySelectorAll('.qp-item').forEach(function (it) {
      const show = !kw || (it.getAttribute('data-s') || '').includes(kw);
      it.style.display = show ? '' : 'none';
      if (show) any = true;
    });
    g.style.display = any ? '' : 'none';
  });
}

function renderCostBySelect(selected, q) {
  const providers = quoteCfg().costProviders || [];
  const sel = document.getElementById('qCostBy');
  let opts = '<option value="">-- 請選擇支援顧問 --</option>' + providers.map(p =>
    `<option value="${escapeHtml(p.username)}"${p.username === selected ? ' selected' : ''}>${escapeHtml(p.displayName || p.username)}</option>`).join('');
  if (selected && !providers.some(p => p.username === selected)) {
    opts += `<option value="${escapeHtml(selected)}" selected>${escapeHtml((q && q.costByName) || selected)}（已不在成本填寫人名單）</option>`;
  }
  sel.innerHTML = opts;
}

/** 勾選變動：更新已選標籤、顧問區塊顯示、未歸類警示、品項提示 */
function onQuoteProductsChanged() {
  const names = readSelectedProducts();
  document.getElementById('qProdCount').textContent = `已勾選 ${names.length} 項`;
  document.getElementById('qProdSelected').innerHTML = names.map(n =>
    `<span class="qp-chip">${escapeHtml(n)}<button type="button" data-unsel="${escapeHtml(n)}" title="取消勾選">&#10005;</button></span>`).join('');

  const need = quoteNeedsConsultant(names);
  document.getElementById('qCostByWrap').style.display = need ? '' : 'none';
  document.getElementById('qSelfCostHint').style.display = (names.length && !need) ? '' : 'none';

  const un = _qUnclassified(names);
  const warn = document.getElementById('qProdWarn');
  if (un.length) {
    warn.style.display = '';
    warn.textContent = un.map(n => `商品「${n}」尚未歸類，簽核時會視為「其他」（主管→總經理）並需顧問填成本。`).join(' ') +
      '請通知管理員到「簽核設定 → 商品歸類表」歸類。';
  } else {
    warn.style.display = 'none';
    warn.textContent = '';
  }

  const hint = document.getElementById('qItemsHint');
  const filled = _qEditing && _qEditing.costFlow && _qEditing.costFlow.state === 'filled';
  if (hint) {
    hint.style.display = (need && filled) ? '' : 'none';
    hint.textContent = (need && filled) ? '顧問已填過成本；若修改品項說明、單位、數量或新增/刪除品項，顧問需要重新填寫成本。' : '';
  }

  const pnl = document.getElementById('quoteTabContentPnl');
  if (pnl && pnl.style.display !== 'none') renderPnlTab();
}

// ════════════════════════════════════════════════════════════
// ── 毛利與簽核路徑 TAB（唯讀摘要；業務自填成本時才可編輯成本）─────────
// ════════════════════════════════════════════════════════════
function switchQuoteTab(tabName) {
  ['info', 'pnl'].forEach(function(t) {
    var content = $('quoteTabContent' + (t === 'info' ? 'Info' : 'Pnl'));
    var btn     = $('quoteTabBtn'     + (t === 'info' ? 'Info' : 'Pnl'));
    if (!content || !btn) return;
    content.style.display = (t === tabName) ? '' : 'none';
    btn.classList.toggle('active', t === tabName);
  });
  if (tabName === 'pnl') renderPnlTab();
}

/** 毛利率顏色分級（僅畫面提示，不代表簽核層級） */
function marginClass(pct) {
  if (pct >= 30) return 'good';
  if (pct >= 15) return 'warn';
  return 'danger';
}

function _qPreviewBlock(q) {
  const pv = q && q.preview;
  if (!pv) {
    return '<div class="q-infobar">儲存後，系統會依商品歸類與毛利率試算簽核類別與需要簽核的關卡。</div>';
  }
  const margin = pv.marginText != null ? `${escapeHtml(pv.marginText)}%` : '尚無法試算';
  const blockers = (pv.blockers || []).map(b => (b && (b.message || b.code)) || '').filter(Boolean);
  return `<div class="q-pv">
    <div class="q-pv-h">簽核試算（伺服器依上次儲存的內容計算；修改後請先儲存）</div>
    <div class="q-pv-grid">
      <div class="q-pv-cell"><div class="k">類別</div><div class="v">${escapeHtml(pv.rowLabel || '—')}</div></div>
      <div class="q-pv-cell"><div class="k">毛利率（折扣後未稅）</div><div class="v">${margin}</div></div>
    </div>
    ${_qPathChips(pv)}
    ${_qListHtml(pv.reasons)}
    ${(pv.warnings || []).length ? `<div class="q-warnbar">${(pv.warnings || []).map(escapeHtml).join('<br>')}</div>` : ''}
    ${blockers.length ? `<div class="q-errbar"><b>目前無法送簽：</b>${_qListHtml(blockers)}</div>` : ''}
  </div>`;
}

function _qCostFlowBlock(q) {
  const cf = q && q.costFlow;
  if (!cf) return '';
  const who = escapeHtml(cf.byName || q.costByName || '顧問');
  let txt = '';
  if (cf.state === 'requested') txt = `已通知 ${who} 填寫成本${cf.requestedAt ? '（' + escapeHtml(_qFmtTime(cf.requestedAt)) + '）' : ''}，尚未完成。`;
  else if (cf.state === 'filled') txt = `${who} 已填寫完成成本${cf.filledAt ? '（' + escapeHtml(_qFmtTime(cf.filledAt)) + '）' : ''}。`;
  else if (cf.state === 'needed') txt = '此單需要顧問填成本，請先選擇支援顧問並儲存。';
  else return '';
  return `<div class="q-infobar">${txt}</div>`;
}

function renderPnlTab() {
  const box = $('pnlBody');
  if (!box) return;
  const e = escapeHtml;
  const q = _qEditing;
  const perm = (q && q.perm) || null;
  const products = readSelectedProducts();
  const need = quoteNeedsConsultant(products);
  const items = readQuoteItems();
  const { discountType, discountValue } = readQuoteDiscount();
  const totals = quoteTotal(items, discountType, discountValue);

  let mode;   // edit＝業務自填成本；view＝有權者唯讀；hidden＝成本由顧問填
  if (products.length && !need && (!q || (perm && (perm.canEditCost === true || perm.isOwner === true)))) mode = 'edit';
  else if (perm && perm.canSeeCost === true && items.some(it => it.cost !== undefined)) mode = 'view';
  else mode = 'hidden';

  let html = '';
  if (!products.length) {
    html += '<div class="q-warnbar" style="margin-top:0">請先在「報價資訊」頁籤勾選商品，系統才能判斷由誰填成本與簽核層級。</div>';
  }

  if (mode === 'hidden') {
    html += '<div class="q-infobar" style="margin-top:' + (products.length ? '0' : '10px') + '">' +
      (need ? '此單的成本由顧問填寫，業務不顯示逐列成本；毛利率與簽核路徑由系統試算（見下方）。'
            : '成本資料目前不顯示。') + '</div>';
    html += _qCostFlowBlock(q);
    html += `<div class="pnl-summary-bar"><div class="pnl-sum-grid" style="grid-template-columns:repeat(2,1fr)">
      <div class="pnl-sum-card"><div class="pnl-sum-label">報價合計（未稅，折扣後）</div><div class="pnl-sum-value">${e(fmtMoney(totals.discounted))}</div></div>
      <div class="pnl-sum-card pnl-margin-card"><div class="pnl-sum-label">毛利率（伺服器試算）</div>
        <div class="pnl-sum-value pnl-margin-val">${q && q.preview && q.preview.marginText != null ? e(q.preview.marginText) + '%' : '—'}</div></div>
    </div></div>`;
  } else {
    html += `<div class="pnl-note">📌 ${mode === 'edit'
      ? '此單由業務自行填成本：請填各品項<strong>未稅單價成本</strong>；毛利以<strong>優惠後未稅報價</strong>為基準。'
      : '以下為各品項成本與毛利（僅有權限者可見，請勿提供客戶）；毛利以<strong>優惠後未稅報價</strong>為基準。'}</div>`;
    html += `<div class="quote-items-wrap" style="margin-top:12px;overflow-x:auto"><table class="quote-items-table pnl-table">
      <thead><tr>
        <th style="width:34px">#</th><th>品項說明</th>
        <th style="width:60px;text-align:right">數量</th>
        <th style="width:110px;text-align:right">報價單價</th>
        <th style="width:110px;text-align:right">報價小計</th>
        <th style="width:120px;text-align:right">成本單價</th>
        <th style="width:110px;text-align:right">成本小計</th>
        <th style="width:100px;text-align:right">毛利</th>
        <th style="width:80px;text-align:right">毛利率</th>
      </tr></thead><tbody id="pnlItemsBody">` +
      items.map(function (it, i) {
        const costVal = it.cost === undefined ? '' : it.cost;
        return `<tr data-idx="${i}">
          <td style="text-align:center;color:#999;font-size:12px">${i + 1}</td>
          <td style="font-size:13px">${e(it.desc || '（未填）')}</td>
          <td style="text-align:right;font-size:13px">${e(String(it.qty))} ${e(it.unit || '')}</td>
          <td style="text-align:right;font-size:13px">${e(fmtMoney(it.unitPrice))}</td>
          <td style="text-align:right;font-size:13px;font-weight:500">${e(fmtMoney(it.qty * it.unitPrice))}</td>
          <td style="text-align:right">${mode === 'edit'
            ? `<input type="number" class="pnl-cost-input" data-idx="${i}" value="${e(String(costVal))}" min="0" step="1" placeholder="輸入成本">`
            : `<span style="font-size:13px">${e(fmtMoney(it.cost || 0))}</span>`}</td>
          <td class="pnl-cst-sub" style="text-align:right;font-size:13px"></td>
          <td class="pnl-gp-cell" style="text-align:right;font-size:13px;font-weight:600"></td>
          <td class="pnl-pct-cell" style="text-align:right;font-size:13px;font-weight:700"></td>
        </tr>`;
      }).join('') + `</tbody></table></div>
      <div id="pnlDiscountNote" class="pnl-discount-note" style="display:none"></div>
      <div class="pnl-summary-bar"><div class="pnl-sum-grid">
        <div class="pnl-sum-card"><div class="pnl-sum-label">報價合計（未稅，折扣後）</div><div class="pnl-sum-value" id="pnlRevenue">NT$ 0</div></div>
        <div class="pnl-sum-card"><div class="pnl-sum-label">成本合計</div><div class="pnl-sum-value" id="pnlCostTotal">NT$ 0</div></div>
        <div class="pnl-sum-card"><div class="pnl-sum-label">毛利</div><div class="pnl-sum-value pnl-gp-val" id="pnlGrossProfit">NT$ 0</div></div>
        <div class="pnl-sum-card pnl-margin-card"><div class="pnl-sum-label">整體毛利率（畫面試算）</div><div class="pnl-sum-value pnl-margin-val" id="pnlMarginPct">—</div></div>
      </div></div>`;
  }

  html += _qPreviewBlock(q);
  box.innerHTML = html;

  if (mode !== 'hidden') {
    box.querySelectorAll('.pnl-cost-input').forEach(function (inp) {
      inp.addEventListener('input', function () {
        const idx = parseInt(this.dataset.idx, 10);
        const row = $('quoteItemsBody').querySelectorAll('tr')[idx];
        if (row) {
          if (this.value === '') row.removeAttribute('data-cost');
          else row.dataset.cost = String(parseFloat(this.value) || 0);
        }
        updatePnlNumbers();
      });
    });
    updatePnlNumbers();
  }
}

/** 重算 PNL 表格各列與匯總（成本取自品項列的 data-cost） */
function updatePnlNumbers() {
  const items = readQuoteItems();
  const { discountType, discountValue } = readQuoteDiscount();
  const totals = quoteTotal(items, discountType, discountValue);
  const revenue = totals.discounted;
  let totalCost = 0;

  document.querySelectorAll('#pnlItemsBody tr').forEach(function (tr) {
    const i = parseInt(tr.dataset.idx, 10);
    const it = items[i];
    if (!it) return;
    const cost = parseFloat(it.cost) || 0;
    const revSub = it.qty * it.unitPrice;
    const costSub = it.qty * cost;
    const gp = revSub - costSub;
    const pct = revSub > 0 ? (gp / revSub * 100) : 0;
    totalCost += costSub;
    tr.querySelector('.pnl-cst-sub').textContent = fmtMoney(costSub);
    const g = tr.querySelector('.pnl-gp-cell');
    g.textContent = fmtMoney(gp);
    g.style.color = gp >= 0 ? '#2e7d32' : '#c62828';
    const p = tr.querySelector('.pnl-pct-cell');
    p.textContent = revSub > 0 ? pct.toFixed(1) + '%' : '—';
    p.style.color = pct >= 30 ? '#2e7d32' : pct >= 15 ? '#e65100' : '#c62828';
  });

  const gpAll = revenue - totalCost;
  $('pnlRevenue').textContent   = fmtMoney(revenue);
  $('pnlCostTotal').textContent = fmtMoney(totalCost);
  const gpEl = $('pnlGrossProfit');
  gpEl.textContent = fmtMoney(gpAll);
  gpEl.className = 'pnl-sum-value pnl-gp-val ' + (gpAll >= 0 ? 'positive' : 'negative');
  const t = _qTruncPct(gpAll, revenue);
  const pctEl = $('pnlMarginPct');
  pctEl.textContent = t === null ? '—' : t + '%';
  pctEl.className = 'pnl-sum-value pnl-margin-val ' + (t === null ? '' : marginClass(parseFloat(t)));

  const discNote = $('pnlDiscountNote');
  if (discNote) {
    if (discountType !== 'none' && discountValue) {
      discNote.style.display = '';
      discNote.textContent = discountType === 'percent'
        ? `⚡ 已套用 ${discountValue}% 折扣，毛利以折扣後金額為基準計算。`
        : `⚡ 已套用議價總額 ${fmtMoney(discountValue)}，毛利以議價金額為基準計算。`;
    } else {
      discNote.style.display = 'none';
    }
  }
}

// ── 關閉 Modal ───────────────────────────────────────────────
function closeQuoteModal() {
  $('quoteModalOverlay').style.display = 'none';
  _qEditing = null;
}

/** 表單上方「簽核狀態」唯讀區 */
function renderQuoteApprovalState(q) {
  const box = $('qApprovalState');
  if (!box) return;
  if (!q) { box.innerHTML = '新建立（尚未送簽）'; return; }
  const st = quoteStageInfo(q);
  let html = `<span class="quote-status ${escapeHtml(st.cls)}">${escapeHtml(st.label)}</span>`;
  const a = q.approval;
  if (a && a.state === 'approved' && a.valid !== false) {
    html += '<span class="q-sub">修改客戶、聯絡人、地址、電話、專案、備註、報價期限、商品、品項、單價、成本或折扣會使核准作廢，需重新送簽（儲存前會再次確認）。</span>';
  } else if (a && a.state === 'returned') {
    const h = (a.history || []).filter(x => x && x.action === 'RETURN');
    const last = h.length ? h[h.length - 1] : null;
    if (last && last.comment) html += `<span class="q-sub">駁回原因：${escapeHtml(last.comment)}${last.byName ? '（' + escapeHtml(last.byName) + '）' : ''}</span>`;
  }
  if (QUOTE_LEGACY_LABEL[q.status]) html += `<span class="q-sub">舊狀態：${escapeHtml(QUOTE_LEGACY_LABEL[q.status])}（簽核功能上線前）</span>`;
  box.innerHTML = html;
}

// ── 報價期限欄位：業務自行輸入；沒動過就隨「建立日期」重算預設值 ─────────────────────
let _qValidUntilTouched = false;
function bindQuoteValidUntil() {
  const vu = $('qValidUntil'), qd = $('qDate');
  if (vu && !vu._qBound) {
    vu._qBound = true;
    vu.addEventListener('input', function () { _qValidUntilTouched = true; });
  }
  if (qd && !qd._qBoundVu) {
    qd._qBoundVu = true;
    qd.addEventListener('change', function () {
      if (_qValidUntilTouched) return;
      const d = quoteLastWorkingDay(qd.value);
      if (d && $('qValidUntil')) $('qValidUntil').value = d;
    });
  }
}

// ── 開啟新增 / 編輯 Modal ────────────────────────────────────
async function openQuoteModal(idOrNull) {
  try {
    // 確保聯絡人資料已載入
    if (typeof allContacts === 'undefined' || !allContacts || allContacts.length === 0) {
      try {
        const r = await fetch(API + '/contacts');
        // app.js 以 `let allContacts` 宣告：寫 window.allContacts 不會更新那個變數，必須直接賦值
        if (r.ok) { allContacts = await r.json(); }
      } catch (fetchErr) {
        console.warn('載入聯絡人失敗', fetchErr);
      }
    }
    if (!_qCfg) await loadQuoteConfig();

    const contacts = (typeof allContacts !== 'undefined' && allContacts) ? allContacts : [];
    let q = null;
    if (idOrNull) {
      q = (await _qFetchQuote(idOrNull)) || _qFind(idOrNull);
      if (!q) { showToast('找不到此報價單'); return; }
      if (q.perm && q.perm.canEdit === false) {
        showToast(q.approval && q.approval.state === 'pending'
          ? '簽核中的報價單已鎖定，請先「撤回」再修改' : '你沒有修改此報價單的權限');
        return;
      }
    }
    _qEditing = q;
    const today = taipeiTodayClient();

    $('quoteModalTitle').textContent = q ? ('編輯報價單 ' + (q.quoteNo || '')) : '新增報價單';
    $('quoteId').value      = q ? (q.id           || '') : '';
    $('qCompany').value     = q ? (q.company       || '') : '';
    $('qPhone').value       = q ? (q.phone         || '') : '';
    $('qMobile').value      = q ? (q.mobile        || '') : '';
    $('qAddress').value     = q ? (q.address       || '') : '';
    $('qDate').value        = q ? (q.quoteDate     || today) : today;
    $('qProjectName').value = q ? (q.projectName   || '') : '';
    $('qNote').value        = q ? (q.note          || '') : '';
    // 報價期限：新單預設「報價日期當月最後一個工作天」，業務沒改過就跟著報價日期走；編輯舊單則尊重已存的值
    $('qValidUntil').value  = q ? (q.validUntil || quoteLastWorkingDay($('qDate').value)) : quoteLastWorkingDay($('qDate').value);
    _qValidUntilTouched     = !!q;
    bindQuoteValidUntil();
    renderQuoteApprovalState(q);

    // 公司 datalist
    var companySet = new Set(contacts.map(function(c){ return c.company; }).filter(Boolean));
    $('quoteCompanyList').innerHTML = Array.from(companySet).sort()
      .map(function(c){ return '<option value="' + escapeHtml(c) + '">'; }).join('');

    // 聯絡人下拉
    buildQuoteContactSelect(q ? (q.contactId || '') : '', q ? (q.company || '') : '');

    // 商品與支援顧問
    renderQuoteProducts(q ? (q.products || []) : []);
    renderCostBySelect(q ? (q.costBy || '') : '', q);
    $('qCostNote').value = q && q.costFlow ? (q.costFlow.note || '') : '';
    const providers = quoteCfg().costProviders || [];
    $('qCostByHint').textContent = providers.length ? '' : '尚未設定成本填寫人，請聯絡管理員到「簽核設定 → 簽核名冊」新增。';

    // 項目列表（保留伺服器給的 lid；成本只存在列的 data-cost，不顯示在業務表單）
    var items = (q && q.items && q.items.length)
      ? q.items
      : [{ desc: '', unit: '式', qty: 1, unitPrice: 0 }];
    renderQuoteItems(items);

    // ── 優惠設定還原 ──
    var discType  = (q && q.discountType)  ? q.discountType  : 'none';
    var discValue = (q && q.discountValue != null && q.discountValue !== 0) ? q.discountValue : '';
    document.querySelectorAll('input[name="qDiscountType"]').forEach(function(radio) {
      radio.checked = (radio.value === discType);
    });
    $('qDiscountValue').value = discValue;
    applyDiscountMode(discType);

    // ── 事件綁定（用 overlay 旗標確保只綁一次）──
    var overlay = $('quoteModalOverlay');

    if (!overlay._qBound) {
      overlay._qBound = true;

      // 優惠模式切換
      document.querySelectorAll('input[name="qDiscountType"]').forEach(function(radio) {
        radio.addEventListener('change', function() {
          applyDiscountMode(this.value);
          updateQuoteTotals();
        });
      });
      $('qDiscountValue').addEventListener('input', updateQuoteTotals);

      // 新增項目
      $('addQuoteItemBtn').addEventListener('click', function() {
        var current = readQuoteItems();
        current.push({ desc: '', unit: '式', qty: 1, unitPrice: 0 });
        renderQuoteItems(current);
        updateQuoteTotals();
      });

      // 公司輸入時篩聯絡人
      $('qCompany').addEventListener('input', function() {
        buildQuoteContactSelect('', this.value);
      });

      // 聯絡人選擇後自動帶入
      $('qContactId').addEventListener('change', autoFillFromContact);

      // 商品勾選／搜尋／移除已選標籤
      $('qProductsList').addEventListener('change', function(ev) {
        if (ev.target && ev.target.classList.contains('qp-cb')) onQuoteProductsChanged();
      });
      $('qProdSearch').addEventListener('input', filterQuoteProducts);
      $('qProdSelected').addEventListener('click', function(ev) {
        const b = ev.target.closest('[data-unsel]');
        if (!b) return;
        const name = b.getAttribute('data-unsel');
        document.querySelectorAll('#qProductsList .qp-cb').forEach(function(cb) { if (cb.value === name) cb.checked = false; });
        onQuoteProductsChanged();
      });

      // 關閉事件
      $('quoteModalClose').addEventListener('click', closeQuoteModal);
      $('quoteModalCancel').addEventListener('click', closeQuoteModal);
    }

    // 儲存按鈕 — 每次重設 onclick 避免舊 context 殘留
    $('quoteModalSave').onclick = saveQuote;
    $('quoteModalSave').disabled = false;

    // 每次開啟都從第一頁開始
    switchQuoteTab('info');
    onQuoteProductsChanged();

    overlay.style.display = 'flex';
    updateQuoteTotals();

  } catch (err) {
    console.error('openQuoteModal 錯誤:', err);
    showToast('開啟報價單失敗：' + (err.message || err));
  }
}

// ── 優惠模式切換 UI ───────────────────────────────────────────
function applyDiscountMode(type) {
  const wrap   = $('qDiscountInputWrap');
  const prefix = $('qDiscountPrefix');
  const suffix = $('qDiscountSuffix');
  const input  = $('qDiscountValue');

  if (type === 'none') {
    wrap.style.display = 'none';
  } else {
    wrap.style.display = '';
    if (type === 'percent') {
      prefix.textContent    = '折扣百分比';
      suffix.textContent    = '（例：90 = 九折，85 = 八五折）';
      input.placeholder     = '請輸入大於 0、小於 100 的數字';
      input.min = '0'; input.max = '100'; input.step = '0.1';
    } else if (type === 'amount') {
      prefix.textContent    = '議價總額（未稅）';
      suffix.textContent    = '直接輸入業務談好的未稅金額（需低於小計）';
      input.placeholder     = '請輸入金額（NT$）';
      input.min = '0'; input.max = ''; input.step = '1';
    }
  }
}

// ── 讀取目前優惠設定 ──────────────────────────────────────────
function readQuoteDiscount() {
  const type  = (document.querySelector('input[name="qDiscountType"]:checked') || {}).value || 'none';
  const value = parseFloat($('qDiscountValue').value) || 0;
  return { discountType: type, discountValue: value };
}

// ── 聯絡人下拉選單建構 ────────────────────────────────────────
function buildQuoteContactSelect(selectedId, filterCompany) {
  const sel = $('qContactId');
  let contacts = allContacts || [];
  if (filterCompany) {
    contacts = contacts.filter(c => c.company === filterCompany);
  }
  let html = '<option value="">-- 選擇聯絡人（自動帶入資料）--</option>' +
    contacts.map(c =>
      '<option value="' + escapeHtml(c.id) + '" data-name="' + escapeHtml(c.name || '') + '"' + (c.id === selectedId ? ' selected' : '') + '>' +
      escapeHtml(c.name || '') + (c.company ? (' - ' + escapeHtml(c.company)) : '') +
      '</option>'
    ).join('');
  // 編輯時原聯絡人不在清單內（公司名不完全相同、或已不是我的聯絡人）：補一個「原聯絡人」選項，
  // 否則下拉是空的，儲存時印在客戶報價單上的聯絡人會被靜默清空
  if (selectedId && !contacts.some(c => c.id === selectedId) && _qEditing && _qEditing.contactName && _qEditing.contactId === selectedId) {
    html += '<option value="' + escapeHtml(selectedId) + '" data-name="' + escapeHtml(_qEditing.contactName) + '" selected>' +
      escapeHtml(_qEditing.contactName) + '（原聯絡人）</option>';
  }
  sel.innerHTML = html;
}

// ── 聯絡人自動帶入 ────────────────────────────────────────────
function autoFillFromContact() {
  const id = $('qContactId').value;
  if (!id) return;
  const c = (allContacts || []).find(x => x.id === id);
  if (!c) return;
  if (c.company)  { $('qCompany').value  = c.company; }
  if (c.phone)    { $('qPhone').value    = c.phone;   }
  if (c.mobile)   { $('qMobile').value   = c.mobile;  }
  if (c.address)  { $('qAddress').value  = c.address; }
  buildQuoteContactSelect(id, c.company);
}

// ── 項目列表渲染 ─────────────────────────────────────────────
// 每列帶 data-lid（伺服器指派的穩定列代號）與 data-cost（成本；只在有權者/業務自填時存在）。
// 成本不在這張表單顯示或編輯：顧問填、或業務自填時到「毛利與簽核路徑」頁籤。
function renderQuoteItems(items) {
  const tbody = $('quoteItemsBody');
  tbody.innerHTML = items.map(function (it, i) {
    const qty   = parseFloat(it.qty)       || 1;
    const price = parseFloat(it.unitPrice) || 0;
    const attrs = (it.lid ? ' data-lid="' + escapeHtml(it.lid) + '"' : '') +
      (it.cost !== undefined && it.cost !== null && it.cost !== '' ? ' data-cost="' + escapeHtml(String(it.cost)) + '"' : '');
    return '<tr data-idx="' + i + '"' + attrs + '>' +
      '<td style="text-align:center;color:#999;font-size:12px">' + (i + 1) + '</td>' +
      '<td><input type="text" class="qi-desc" value="' + escapeHtml(it.desc || '') + '" ' +
        'placeholder="品項說明" style="width:100%;border:1px solid #ddd;border-radius:4px;padding:5px 8px;font-size:13px;box-sizing:border-box"></td>' +
      '<td><input type="text" class="qi-unit" value="' + escapeHtml(it.unit || '式') + '" ' +
        'style="width:54px;border:1px solid #ddd;border-radius:4px;padding:5px 6px;font-size:13px;text-align:center"></td>' +
      '<td><input type="number" class="qi-qty" value="' + qty + '" min="0.001" step="1" ' +
        'style="width:64px;border:1px solid #ddd;border-radius:4px;padding:5px 6px;font-size:13px;text-align:right"></td>' +
      '<td><input type="number" class="qi-price" value="' + price + '" min="0" step="1" ' +
        'style="width:104px;border:1px solid #ddd;border-radius:4px;padding:5px 6px;font-size:13px;text-align:right"></td>' +
      '<td class="qi-subtotal" style="text-align:right;font-size:13px;padding-right:6px;white-space:nowrap">' +
        fmtMoney(qty * price) + '</td>' +
      '<td style="text-align:center"><button type="button" class="qi-remove" ' +
        'title="移除此項目" style="background:none;border:none;color:#e53935;cursor:pointer;font-size:16px;line-height:1;padding:2px 4px">&#10005;</button></td>' +
      '</tr>';
  }).join('');

  // 輸入事件：即時更新小計
  tbody.querySelectorAll('input').forEach(function (inp) {
    inp.addEventListener('input', function () {
      const row   = inp.closest('tr');
      const qty   = parseFloat(row.querySelector('.qi-qty').value)   || 0;
      const price = parseFloat(row.querySelector('.qi-price').value) || 0;
      row.querySelector('.qi-subtotal').textContent = fmtMoney(qty * price);
      updateQuoteTotals();
    });
  });

  // 移除按鈕
  tbody.querySelectorAll('.qi-remove').forEach(function (btn) {
    btn.addEventListener('click', function () {
      var current = readQuoteItems();
      if (current.length <= 1) { showToast('至少需保留一個報價項目'); return; }
      var idx = parseInt(btn.closest('tr').dataset.idx, 10);
      current.splice(idx, 1);
      renderQuoteItems(current);
      updateQuoteTotals();
    });
  });
}

// ── 讀取目前項目列表 ──────────────────────────────────────────
// 回傳 { lid?, desc, unit, qty, unitPrice, cost? }；lid/cost 只在列上有值時才出現
function readQuoteItems() {
  return Array.from($('quoteItemsBody').querySelectorAll('tr')).map(function (row) {
    const it = {
      desc:      row.querySelector('.qi-desc').value.trim(),
      unit:      row.querySelector('.qi-unit').value.trim() || '式',
      qty:       parseFloat(row.querySelector('.qi-qty').value)   || 1,
      unitPrice: parseFloat(row.querySelector('.qi-price').value) || 0,
    };
    if (row.dataset.lid) it.lid = row.dataset.lid;
    if (row.dataset.cost !== undefined) it.cost = parseFloat(row.dataset.cost) || 0;
    return it;
  });
}

// ── 更新合計顯示 ─────────────────────────────────────────────
function updateQuoteTotals() {
  const items = readQuoteItems();
  const { discountType, discountValue } = readQuoteDiscount();
  const { sub, discounted, discountAmt, tax, total } = quoteTotal(items, discountType, discountValue);

  $('qSubtotal').textContent = fmtMoney(sub);
  $('qTax').textContent      = fmtMoney(tax);
  $('qTotal').textContent    = fmtMoney(total);

  // 優惠折扣列 & 優惠價列
  const hasDiscount = discountType !== 'none' && discountAmt !== 0;
  $('qDiscountRow').style.display    = hasDiscount ? '' : 'none';
  $('qDiscountedRow').style.display  = hasDiscount ? '' : 'none';

  if (hasDiscount) {
    if (discountType === 'percent') {
      $('qDiscountRowLabel').textContent = '優惠折扣（' + discountValue + '%）';
    } else {
      $('qDiscountRowLabel').textContent = '優惠折扣（議價）';
    }
    $('qDiscountAmt').textContent      = '- ' + fmtMoney(discountAmt);
    $('qDiscountedPrice').textContent  = fmtMoney(discounted);
  }

  // 毛利與簽核路徑頁籤開著時同步
  const pnl = document.getElementById('quoteTabContentPnl');
  if (pnl && pnl.style.display !== 'none' && document.getElementById('pnlItemsBody')) updatePnlNumbers();
}

// ── 儲存報價單 ──────────────────────────────────────────────
async function saveQuote() {
  if (_qSaving) return;
  const company   = $('qCompany').value.trim();
  const quoteDate = $('qDate').value;
  if (!company)   { showToast('請輸入客戶公司名稱'); return; }
  if (!quoteDate) { showToast('請選擇報價日期');     return; }
  // 報價期限（印在報價單 Remarks 第 4 條）：必填、不可早於建立日期（伺服器端也會驗證）
  const validUntil = ($('qValidUntil').value || '').trim();
  if (!validUntil) { switchQuoteTab('info'); $('qValidUntil').focus(); showToast('請輸入報價期限'); return; }
  if (validUntil < quoteDate) { switchQuoteTab('info'); $('qValidUntil').focus(); showToast('報價期限不能早於建立日期'); return; }

  const products = readSelectedProducts();
  if (!products.length) { switchQuoteTab('info'); showToast('請至少勾選一項商品'); return; }
  const needC  = quoteNeedsConsultant(products);
  const costBy = needC ? $('qCostBy').value : '';
  if (needC && !costBy) { switchQuoteTab('info'); showToast('請選擇支援顧問（成本填寫人）'); return; }

  const id             = $('quoteId').value;
  const contactSel     = $('qContactId');
  const selOpt         = contactSel.selectedOptions[0];
  // 聯絡人姓名取自選項的 data-name（不從顯示文字切字串）；沒選聯絡人時，編輯中的單保留原值，不可靜默清空
  let autoName         = selOpt && selOpt.value ? (selOpt.getAttribute('data-name') || '') : '';
  let contactIdToSend  = contactSel.value;
  if (!contactIdToSend && _qEditing && company === (_qEditing.company || '')) { autoName = _qEditing.contactName || ''; contactIdToSend = _qEditing.contactId || ''; }

  const { discountType, discountValue } = readQuoteDiscount();
  // 數量空白或 0：畫面小計顯示 0，但存檔會被當成 1 → 不偷偷改，直接擋下請業務填正確數量
  const badQtyInput = Array.from($('quoteItemsBody').querySelectorAll('.qi-qty')).find(function (i) { return !(parseFloat(i.value) > 0); });
  if (badQtyInput) { switchQuoteTab('info'); badQtyInput.focus(); showToast('品項數量必須大於 0'); return; }
  const items = readQuoteItems();
  const sub = quoteTotal(items, 'none', 0).sub;
  if (discountType === 'percent' && !(discountValue > 0 && discountValue < 100)) {
    showToast('折扣百分比需大於 0 且小於 100'); return;
  }
  if (discountType === 'amount' && !(discountValue > 0 && discountValue < sub)) {
    showToast('議價總額需大於 0 且低於小計（未稅）'); return;
  }

  // 成本只在「業務自填」時才送（伺服器端也只接受有權者的成本）；沒動過成本的列不送，避免誤蓋成 0
  const selfCost = !needC;
  const payloadItems = items.map(function (it) {
    const o = { desc: it.desc, unit: it.unit, qty: it.qty, unitPrice: it.unitPrice };
    if (it.lid) o.lid = it.lid;
    if (selfCost && it.cost !== undefined) o.cost = it.cost;
    return o;
  });

  const payload = {
    contactId:     contactIdToSend,
    company:       company,
    contactName:   autoName,
    phone:         $('qPhone').value.trim(),
    mobile:        $('qMobile').value.trim(),
    address:       $('qAddress').value.trim(),
    quoteDate:     quoteDate,
    projectName:   $('qProjectName').value.trim(),
    validUntil:    validUntil,
    items:         payloadItems,
    discountType:  discountType,
    discountValue: discountValue,
    note:          $('qNote').value.trim(),
    products:      products,
    costBy:        needC ? costBy : null,
    costNote:      needC ? $('qCostNote').value.trim() : '',
  };

  const saveBtn = $('quoteModalSave');
  _qSaving = true;
  saveBtn.disabled = true;
  try {
    const method = id ? 'PUT'  : 'POST';
    const url    = id ? (API + '/quotations/' + encodeURIComponent(id)) : (API + '/quotations');
    const send = async function (body) {
      const r = await fetch(url, { method: method, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
      const j = await r.json().catch(function () { return {}; });
      return { r: r, j: j };
    };
    let res = await send(payload);
    // 已核准的單：修改簽核涵蓋欄位會使核准作廢，伺服器要求明確確認
    if (!res.r.ok && res.j.code === 'WILL_VOID') {
      const ok = await qDialog({
        title: '修改會使核准作廢',
        message: '此報價單已核准。你修改了簽核涵蓋的內容（客戶、聯絡人、地址、電話、專案、備註、報價期限、商品、品項、單價、成本或折扣），儲存後原核准會作廢，需要重新送簽。\n仍要儲存？',
        buttons: [
          { text: '儲存並作廢核准', value: true, cls: 'btn-danger' },
          { text: '取消', value: false, cls: 'btn-secondary' },
        ],
      });
      if (!ok) return;
      res = await send(Object.assign({}, payload, { confirmVoid: true }));
    }
    if (!res.r.ok) {
      if (res.j.code === 'LOCKED_PENDING') {
        showToast(res.j.error || '簽核中的報價單已鎖定，請先撤回再修改');
        closeQuoteModal();
        loadQuotationsView();
        return;
      }
      showToast(res.j.error || '儲存失敗');
      return;
    }
    const saved = (res.j && res.j.quote) ? res.j.quote : res.j;
    $('quoteModalOverlay').style.display = 'none';
    _qEditing = null;
    let msg = id ? '報價單已更新' : '報價單已建立 ✅';
    if (saved && saved.costFlow && saved.costFlow.state === 'requested') msg += '，已通知顧問填寫成本';
    else if (saved && saved.approval && saved.approval.state === 'none' && _qLastHistory(saved) && _qLastHistory(saved).action === 'INVALIDATE') msg += '，原核准已作廢，請重新送簽';
    showToast(msg);
    loadQuotationsView();
  } catch(e) {
    showToast('儲存失敗，請重試');
  } finally {
    _qSaving = false;
    saveBtn.disabled = false;
  }
}

// ── 我的聯絡資訊（印在報價單右上「廠商資料」框；業務自行維護、之後自動套用）─────────────────
async function openMyContactModal() {
  let info = { name: '', phone: '', ext: '', mobile: '', displayName: '' };
  try { const r = await fetch(`${API}/me/contact`); if (r.ok) info = await r.json(); } catch (e) { /* 失敗就用空白表單 */ }
  closeMyContactModal();
  const ov = document.createElement('div');
  ov.className = 'modal-overlay open';
  ov.id = 'myContactOverlay';
  ov.innerHTML = `
    <div class="modal" style="max-width:460px;width:96%">
      <div class="modal-header">
        <h2>我的聯絡資訊</h2>
        <button class="modal-close" onclick="closeMyContactModal()">&#10005;</button>
      </div>
      <div class="modal-body">
        <p style="font-size:12.5px;color:#5f6b7a;line-height:1.6;margin:0 0 14px">這些資料會印在報價單右上角的「廠商資料」框。維護一次，之後匯出與預覽的報價單都會自動套用。</p>
        <div class="form-group"><label>聯絡人姓名</label>
          <input type="text" id="mcName" maxlength="50" placeholder="${escapeHtml(info.displayName || '')}">
          <div style="font-size:12px;color:#8a94a3;margin-top:4px">未填時會使用帳號顯示名稱</div></div>
        <div style="display:flex;gap:12px">
          <div class="form-group" style="flex:2"><label>電話</label><input type="text" id="mcPhone" maxlength="40" placeholder="02-2655-2525"></div>
          <div class="form-group" style="flex:1"><label>分機</label><input type="text" id="mcExt" maxlength="10" placeholder="123"></div>
        </div>
        <div class="form-group"><label>手機</label><input type="text" id="mcMobile" maxlength="30" placeholder="0912-345-678"></div>
        <div id="mcMsg" style="font-size:12.5px;color:#c62828;min-height:18px"></div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-secondary" onclick="closeMyContactModal()">取消</button>
        <button class="btn btn-primary" id="mcSaveBtn" onclick="saveMyContact()">儲存</button>
      </div>
    </div>`;
  document.body.appendChild(ov);
  $('mcName').value = info.name || '';
  $('mcPhone').value = info.phone || '';
  $('mcExt').value = info.ext || '';
  $('mcMobile').value = info.mobile || '';
}

function closeMyContactModal() {
  const ov = document.getElementById('myContactOverlay');
  if (ov) ov.remove();
}

async function saveMyContact() {
  const btn = $('mcSaveBtn'), msg = $('mcMsg');
  msg.textContent = '';
  btn.disabled = true;
  try {
    const r = await fetch(`${API}/me/contact`, {
      method: 'PUT', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: $('mcName').value, phone: $('mcPhone').value, ext: $('mcExt').value, mobile: $('mcMobile').value }),
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) { msg.textContent = j.error || '儲存失敗'; return; }
    closeMyContactModal();
    showToast('聯絡資訊已儲存，之後的報價單會自動套用');
    if (typeof refreshQuotePreview === 'function') refreshQuotePreview();   // 預覽開著的話，立刻看到新資料
  } catch (e) {
    msg.textContent = '儲存失敗，請重試';
  } finally {
    btn.disabled = false;
  }
}
