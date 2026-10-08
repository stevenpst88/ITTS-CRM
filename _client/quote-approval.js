// ═════════════════════════════════════════════════
// ── 報價單簽核面板與設定 (quote-approval.js) ───────────────
// 對外全域函式（由 quote.js / index.html 呼叫）：
//   openQuoteApproval(id)        簽核面板（核准／駁回／送簽／撤回／改派／董事會列印與決議登錄）
//   openQuoteCostFill(id)        顧問填成本（成本明細編輯器 QCL）。cost-sync 起，被指派的顧問在填成本期間看得到單價、折扣與毛利（公司內顧問與業務本來就能互看成本與售價），
//                                對話框因此有：頂端五張卡（報價合計／成本合計／毛利／毛利率／委外佔比，即時連動）＋簽核層級預覽（伺服器 cost-draft/summary）、報價品項區（含連動結果與
//                                「新增報價項目」）、每列「對應」下拉（連動報價數量／拆項／不對應）、顧問姓名欄、報價單預覽與毛利預覽（用畫面上未存檔的內容）、完成前的同步確認清單。
//                                簽核流程本身（approval）對只負責填成本的顧問仍不顯示。看不到單價的情況（舊伺服器）退回改版前的畫面
//   openQuoteApprovalSettings()  簽核設定（名冊／商品歸類表／核決門檻與報價專用章）
//   refreshQuoteInbox()          更新工具列「待我處理」徽章（#quoteInboxBadge）
//
// 設計原則（金錢相關）：
//   · 按鈕顯示與否只看伺服器回傳的 q.perm.*，前端不自行推算「誰能簽」。
//   · 核准／駁回的 body.hash 一律取面板載入時 GET /quotations/:id 回傳的 contentHash。
//   · 任何動作送出後都會重新載入該單，確保畫面與伺服器狀態一致（含 409 衝突）。
//   · 動作期間 busy 旗標鎖住所有按鈕，防止雙擊重送。
//   · 所有事件都綁在各自 overlay 元素上（事件委派）；唯一綁在 document 的 keydown 監聽
//     會在關閉時移除。overlay 被外部移除時，下一次按鍵也會自我清除。
// API 合約見 scratchpad/quote_approval_spec.md 第 5 節。
// 整個檔案包在 IIFE 內，避免與其他 script 的頂層 const/let 重名。
// ═════════════════════════════════════════════════
(function () {
'use strict';

const e = (v) => escapeHtml(v);
const apiBase = () => (typeof API !== 'undefined' ? API : '/api');

const TIER_LABEL = { mgr1: '一級主管', gm: '總經理', chairman: '董事長', board: '董事會決議（秘書代核）' };
const STEP_STATUS = { waiting: '等待中', pending: '簽核中', approved: '已核准', returned: '已駁回' };
const ACTION_LABEL = {
  SUBMIT: '送簽', APPROVE: '核准', RETURN: '駁回', WITHDRAW: '撤回', REASSIGN: '改派',
  INVALIDATE: '核准作廢', COST_REQUEST: '通知顧問填成本', COST_DONE: '顧問完成成本',
};
const COST_STATE_LABEL = { na: '無需顧問填寫', needed: '需顧問填寫（尚未通知）', requested: '等待顧問填寫', filled: '顧問已填寫完成' };
const CLASS_ORDER = ['consult', 'software', 'hardware', 'crm', 'mdm', 'ot', 'other'];
const CLASS_LABEL_DEFAULT = {
  consult: '顧問服務', software: '軟體規劃', hardware: '硬體規劃', crm: 'CRM客服',
  mdm: 'MDM帳單列印', ot: 'OT策略性專案', other: '其他',
};
const ROSTER_GROUPS = [
  { key: 'gm', label: '總經理', desc: '任一人皆可簽總經理關。' },
  { key: 'chairman', label: '董事長', desc: '任一人皆可簽董事長關。' },
  { key: 'boardProxy', label: '董事會代核人', desc: '代理登錄董事會決議並核准。所有在職的 secretary 角色帳號自動具有此權限，這裡只需補充額外的人。' },
  { key: 'costProviders', label: '成本填寫人（支援顧問）', desc: '業務建單時可選的支援顧問名單。' },
  { key: 'sealManagers', label: '報價專用章管理人', desc: '可上傳／刪除報價專用章。所有在職的 secretary 角色帳號與管理員自動具有此權限。' },
];

// ── 樣式 ───────────────────────────────────────────
const QAP_CSS = `
.qap-modal { width: 940px; max-width: 96vw; }
.qap-modal.qap-sm { width: 640px; }
.qap-modal.qap-wide { width: 1180px; max-width: 96vw; }
.qap-body { background: #f6f7f9; padding: 16px 20px; }
.qap-sec { background: #fff; border: 1px solid #e3e6ea; border-radius: 10px; padding: 12px 14px; margin-bottom: 12px; }
.qap-sec h3 { font-size: 14px; font-weight: 700; margin: 0 0 8px; color: #111; }
.qap-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 6px 18px; font-size: 13px; }
.qap-grid .k { color: #6b7684; margin-right: 6px; }
.qap-badge { display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; background: #eef0f3; color: #4a5560; }
.qap-badge.ok { background: #e6f4ea; color: #1e7a3a; }
.qap-badge.warn { background: #fff4e5; color: #8a4b00; }
.qap-badge.bad { background: #fde8e8; color: #b3261e; }
.qap-badge.info { background: #e8f0fe; color: #1558b0; }
.qap-tablewrap { overflow-x: auto; }
.qap-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.qap-table th, .qap-table td { border-bottom: 1px solid #eceef1; padding: 6px 8px; text-align: left; vertical-align: top; }
.qap-table th { background: #f6f7f9; font-weight: 600; white-space: nowrap; }
.qap-table td.r, .qap-table th.r { text-align: right; white-space: nowrap; }
.qap-table td.c, .qap-table th.c { text-align: center; }
.qap-table td.desc { white-space: pre-wrap; word-break: break-word; }
.qap-table tr.qap-uncls td { background: #fff4e5; }
.qap-sum { display: flex; flex-wrap: wrap; gap: 6px 22px; margin-top: 8px; font-size: 13px; }
.qap-sum b { font-weight: 700; }
.qap-path { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; margin: 6px 0; }
.qap-chip { display: inline-flex; align-items: center; gap: 4px; padding: 3px 10px; border-radius: 999px; background: #eef0f3; font-size: 12.5px; }
.qap-chip.sel { background: #e8f0fe; color: #1558b0; }
.qap-chip button { border: none; background: transparent; cursor: pointer; color: inherit; font-size: 12px; padding: 0 0 0 2px; }
.qap-arrow { color: #8a94a0; font-size: 12px; }
.qap-list { margin: 4px 0 0; padding-left: 18px; font-size: 13px; line-height: 1.6; }
.qap-alert { border-radius: 8px; padding: 8px 12px; font-size: 13px; margin-bottom: 10px; line-height: 1.6; }
.qap-alert.warn { background: #fff4e5; color: #8a4b00; border: 1px solid #f5c98b; }
.qap-alert.bad { background: #fde8e8; color: #b3261e; border: 1px solid #f2b8b5; }
.qap-alert.info { background: #e8f0fe; color: #1558b0; border: 1px solid #b6cdf5; }
.qap-steps { list-style: none; margin: 0; padding: 0; }
.qap-steps li { display: flex; gap: 10px; padding: 8px 0; border-bottom: 1px dashed #e3e6ea; font-size: 13px; }
.qap-steps li:last-child { border-bottom: none; }
.qap-steps .ic { width: 22px; height: 22px; border-radius: 50%; display: flex; align-items: center; justify-content: center; flex-shrink: 0; font-size: 12px; background: #eef0f3; color: #6b7684; }
.qap-steps li.approved .ic { background: #1e7a3a; color: #fff; }
.qap-steps li.pending .ic { background: #1a73e8; color: #fff; }
.qap-steps li.returned .ic { background: #b3261e; color: #fff; }
.qap-steps .cm { color: #4a5560; white-space: pre-wrap; word-break: break-word; margin-top: 2px; }
.qap-muted { color: #6b7684; font-size: 12.5px; }
.qap-form { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px 14px; }
.qap-form .full { grid-column: 1 / -1; }
.qap-form label { display: block; font-size: 12.5px; font-weight: 600; color: #555; margin-bottom: 4px; }
.qap-form input, .qap-form textarea, .qap-form select, .qap-input {
  width: 100%; box-sizing: border-box; padding: 8px 10px; border: 1.5px solid #d9dde2; border-radius: 8px;
  font-size: 14px; font-family: inherit; background: #fff; color: #222; outline: none; }
.qap-form input:focus, .qap-form textarea:focus, .qap-form select:focus, .qap-input:focus { border-color: #1a73e8; }
.qap-form textarea { resize: vertical; min-height: 64px; }
.qap-cin { width: 130px !important; text-align: right; }
.qap-cin.bad { border-color: #ea4335 !important; background: #fde8e8; }
.qap-footer { flex-wrap: wrap; }
.qap-footer .sp { flex: 1; }
.qap-ref > summary { cursor: pointer; font-size: 14px; font-weight: 700; color: #111; }
.qap-ref[open] > summary { margin-bottom: 8px; }
.qap-cf-head { margin: 0 0 8px; }
.qap-cf-help { display: inline; }
.qap-cf-help > summary { display: inline; cursor: pointer; color: #1a73e8; margin-left: 4px; }
.qap-cf-help > div { margin-top: 6px; line-height: 1.6; }
body.dark .qap-cf-help > summary { color: #58a6ff; }
.qap-cf-head h3 { font-size: 14px; font-weight: 700; margin: 0 0 4px; color: #111; }
.qap-fs { border: 0; margin: 0; padding: 0; min-width: 0; }
.qap-cf-tot { font-size: 13px; margin-left: 12px; color: #4a5560; }
.qap-cf-tot b { font-size: 15px; color: #1a73e8; }
.qap-cf-note { font-size: 12px; color: #6b7684; margin-left: 6px; }
.qap-cdwrap { margin-bottom: 12px; }
.qap-cdh { font-size: 14px; font-weight: 700; margin: 0 0 8px; color: #111; }
.qap-tabs { display: flex; gap: 4px; border-bottom: 1px solid #e3e6ea; padding: 0 20px; background: #fff; flex-wrap: wrap; }
.qap-tab { border: none; background: transparent; padding: 10px 14px; cursor: pointer; font-size: 13.5px; color: #6b7684; border-bottom: 2px solid transparent; }
.qap-tab.on { color: #1a73e8; border-bottom-color: #1a73e8; font-weight: 600; }
.qap-tools { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; margin-bottom: 10px; font-size: 13px; }
.qap-tools input[type=text] { width: 220px; }
.qap-seal-box { display: flex; gap: 16px; align-items: center; flex-wrap: wrap; }
.qap-seal-img { width: 140px; height: 140px; border: 1px dashed #c4c9d0; border-radius: 8px; display: flex; align-items: center; justify-content: center;
  background-color: #fff; background-image: linear-gradient(45deg, #eee 25%, transparent 25%, transparent 75%, #eee 75%), linear-gradient(45deg, #eee 25%, transparent 25%, transparent 75%, #eee 75%);
  background-size: 16px 16px; background-position: 0 0, 8px 8px; overflow: hidden; }
.qap-seal-img img { max-width: 100%; max-height: 100%; }
.qap-confirm-ov { z-index: 110 !important; }
.qap-confirm-msg { font-size: 14px; line-height: 1.7; white-space: pre-wrap; word-break: break-word; }
@media (max-width: 640px) {
  .qap-grid, .qap-form { grid-template-columns: 1fr; }
  .qap-body { padding: 12px; }
  .qap-tools input[type=text] { width: 100%; }
  .qap-tabs { padding: 0 8px; }
}
body.dark .qap-body { background: #0d1117; }
body.dark .qap-sec { background: #161b22; border-color: #30363d; }
body.dark .qap-sec h3 { color: #e6edf3; }
body.dark .qap-grid .k, body.dark .qap-muted { color: #8b949e; }
body.dark .qap-badge { background: #21262d; color: #c9d1d9; }
body.dark .qap-badge.ok { background: #12351f; color: #6fdc8c; }
body.dark .qap-badge.warn { background: #3d2b0a; color: #f0b866; }
body.dark .qap-badge.bad { background: #3f1717; color: #ff8a80; }
body.dark .qap-badge.info { background: #14283f; color: #79b8ff; }
body.dark .qap-table th { background: #21262d; color: #c9d1d9; }
body.dark .qap-table th, body.dark .qap-table td { border-bottom-color: #30363d; }
body.dark .qap-table tr.qap-uncls td { background: #3d2b0a; }
body.dark .qap-chip { background: #21262d; color: #c9d1d9; }
body.dark .qap-chip.sel { background: #14283f; color: #79b8ff; }
body.dark .qap-alert.warn { background: #3d2b0a; color: #f0b866; border-color: #6b4a14; }
body.dark .qap-alert.bad { background: #3f1717; color: #ff8a80; border-color: #7a2b2b; }
body.dark .qap-alert.info { background: #14283f; color: #79b8ff; border-color: #25476e; }
body.dark .qap-steps li { border-bottom-color: #30363d; }
body.dark .qap-steps .ic { background: #21262d; color: #8b949e; }
body.dark .qap-steps .cm { color: #b1bac4; }
body.dark .qap-form label { color: #8b949e; }
body.dark .qap-form input, body.dark .qap-form textarea, body.dark .qap-form select, body.dark .qap-input { background: #0d1117; color: #e6edf3; border-color: #30363d; }
body.dark .qap-cin.bad { background: #3f1717; }
body.dark .qap-tabs { background: #161b22; border-bottom-color: #30363d; }
body.dark .qap-tab { color: #8b949e; }
body.dark .qap-tab.on { color: #58a6ff; border-bottom-color: #58a6ff; }
body.dark .qap-seal-img { background-color: #fff; border-color: #30363d; }
body.dark .qap-ref > summary, body.dark .qap-cf-head h3, body.dark .qap-cdh { color: #e6edf3; }
body.dark .qap-cf-tot { color: #8b949e; }
body.dark .qap-cf-tot b { color: #58a6ff; }
body.dark .qap-cf-note { color: #8b949e; }
/* cost-sync：顧問對話框的即時試算（卡片列＋簽核層級預覽＋預覽按鈕）、報價品項區、對應 */
.qap-cf-live { margin-bottom: 12px; }
.qap-cf-bar { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 10px; margin: 0 0 8px; }
.qap-cf-bar h3 { font-size: 14px; font-weight: 700; margin: 0; color: #111; }
.qap-cf-bar .sp { flex: 1; }
.qap-cf-live .pnl-summary-bar { margin-top: 0; padding: 12px 14px; }
.qap-cf-pv { margin-top: 10px; padding-top: 10px; border-top: 1px dashed #c5cae9; font-size: 13px; line-height: 1.7; }
.qap-cf-pv[data-state="pending"] .qap-cf-pvbody { opacity: .55; }
.qap-cf-pv .k { color: #6b7684; margin-right: 4px; }
.qap-cf-pvrow { display: flex; flex-wrap: wrap; align-items: center; gap: 4px 14px; }
.qap-cf-pv .qap-alert { margin: 6px 0 0; }
.qap-cf-mis { margin-top: 6px; font-size: 12px; color: #b3261e; }
.qap-ci th, .qap-ci td { vertical-align: middle; }
.qap-ci .qci-title td { font-weight: 700; background: rgba(127,127,127,.12); }
.qap-ci .qci-segrow td { font-weight: 700; }
.qap-ci .qci-old { color: #8a94a0; text-decoration: line-through; margin-right: 4px; }
.qap-ci .qci-chip { display: inline-block; margin-left: 6px; padding: 0 8px; font-size: 11.5px; line-height: 1.7; border-radius: 9px; font-weight: 600; white-space: nowrap; background: #eef0f3; color: #4a5560; }
.qap-ci .qci-chip.chg { background: #e8f0fe; color: #1558b0; }
.qap-ci .qci-chip.bad { background: #fde8e8; color: #b3261e; }
.qap-ci .qci-chip.new { background: #e6f4ea; color: #1e7a3a; }
.qap-ci .qci-chip.need { background: #fff4e5; color: #8a4b00; }
.qap-ci tr.qci-hit td { background: rgba(26,115,232,.06); }
.qap-ci tr.qci-bad td { background: rgba(234,67,53,.08); }
.qap-ci .qci-in { padding: 5px 8px; font-size: 13px; }
.qap-ci input.qci-nq { width: 90px; text-align: right; }
.qap-ci input.qci-nu { width: 80px; text-align: center; }
.qap-ci input.qci-linked { background: rgba(26,115,232,.08); cursor: not-allowed; }
.qap-ci input.qci-bad { border-color: #ea4335; background: #fde8e8; }
.qap-ci .qci-rm { background: none; border: none; cursor: pointer; font-size: 16px; color: #e53935; line-height: 1; padding: 2px 6px; }
.qap-cf-newbar { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 12px; margin-top: 8px; }
.qap-cf-extra { font-size: 13px; margin-left: 14px; color: #4a5560; }
.qap-cf-extra b { color: #111; }
body.dark .qap-cf-bar h3 { color: #e6edf3; }
body.dark .qap-cf-pv { border-top-color: #30363d; }
body.dark .qap-cf-pv .k { color: #8b949e; }
body.dark .qap-cf-mis { color: #ff8a80; }
body.dark .qap-ci .qci-old { color: #8b949e; }
body.dark .qap-ci .qci-chip { background: #21262d; color: #c9d1d9; }
body.dark .qap-ci .qci-chip.chg { background: #14283f; color: #79b8ff; }
body.dark .qap-ci .qci-chip.bad { background: #3f1717; color: #ff8a80; }
body.dark .qap-ci .qci-chip.new { background: #12351f; color: #6fdc8c; }
body.dark .qap-ci .qci-chip.need { background: #3d2b0a; color: #f0b866; }
body.dark .qap-ci tr.qci-hit td { background: rgba(88,166,255,.08); }
body.dark .qap-ci tr.qci-bad td { background: rgba(255,138,128,.1); }
body.dark .qap-ci input.qci-bad { background: #3f1717; }
body.dark .qap-ci input.qci-linked { background: rgba(88,166,255,.10); }
body.dark .qap-cf-extra { color: #8b949e; }
body.dark .qap-cf-extra b { color: #e6edf3; }
@media (max-width: 640px) {
  .qap-cf-bar .sp { display: none; }
  .qap-cf-extra { display: block; margin: 4px 0 0; }
}
`;

function ensureStyle() {
  if (document.getElementById('qapStyle')) return;
  const st = document.createElement('style');
  st.id = 'qapStyle';
  st.textContent = QAP_CSS;
  document.head.appendChild(st);
}

// ── 共用工具 ───────────────────────────────────────
function toast(msg, ms) {
  if (typeof showToast === 'function') showToast(msg, ms);
}

/** 呼叫 API；回傳 {ok,status,data,net}，永不 throw */
async function apiCall(method, path, body) {
  const opt = { method, headers: {}, credentials: 'same-origin' };
  if (body !== undefined) {
    opt.headers['Content-Type'] = 'application/json';
    opt.body = JSON.stringify(body);
  }
  let r;
  try {
    r = await fetch(apiBase() + path, opt);
  } catch (err) {
    return { ok: false, status: 0, net: true, data: { error: '網路連線失敗，請稍後再試' } };
  }
  let data = null;
  try { data = await r.json(); } catch (err) { data = null; }
  return { ok: r.ok, status: r.status, data: data || {} };
}

/** 錯誤訊息：伺服器 error ＋（若有）blockers 逐項列出 */
function errMsg(res) {
  const d = res.data || {};
  let m = d.error || '';
  if (Array.isArray(d.blockers) && d.blockers.length) {
    const lines = d.blockers.map((b) => (b && (b.message || b.code)) || '').filter(Boolean);
    if (lines.length) m = (m ? m + '：' : '') + lines.join('；');
  }
  if (!m) {
    if (res.status === 401) m = '登入已逾時，請重新登入';
    else if (res.status === 403) m = '你沒有權限執行此操作';
    else if (res.status === 404) m = '找不到資料';
    else m = '操作失敗（' + res.status + '）';
  }
  return m;
}

function afterMutation() {
  try { if (typeof loadQuotationsView === 'function') loadQuotationsView(); } catch (err) { /* 列表重整失敗不影響面板 */ }
  refreshQuoteInbox();
}

function fmtTime(v) {
  if (!v) return '';
  const d = new Date(v);
  if (isNaN(d.getTime())) return String(v);
  try {
    return d.toLocaleString('zh-TW', { timeZone: 'Asia/Taipei', hour12: false, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
  } catch (err) {
    return d.toISOString();
  }
}

function fmtCents(c) {
  const n = Number(c);
  if (!isFinite(n)) return '';
  return (n / 100).toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
}

function fmtNum(n) {
  const v = Number(n);
  if (!isFinite(v)) return '';
  return v.toLocaleString('en-US', { maximumFractionDigits: 2 });
}

/** 逐列毛利率顯示：兩位小數、向 0 截斷（僅供顯示，簽核門檻一律由伺服器判定） */
function lineMarginText(it) {
  const qty = parseFloat(it.qty) || 1;
  const rev = Math.round(qty * (parseFloat(it.unitPrice) || 0) * 100);
  if (rev <= 0 || it.cost === undefined || it.cost === null || it.cost === '') return '—';
  const cost = Math.round(qty * (parseFloat(it.cost) || 0) * 100);
  const t = Math.trunc(((rev - cost) * 10000) / rev);
  return (t / 100).toFixed(2) + '%';
}

function tierLabel(k) { return TIER_LABEL[k] || String(k || ''); }

function approvalBadge(q) {
  const ap = q.approval;
  const st = ap ? ap.state : 'none';
  if (st === 'approved') return ap.valid ? '<span class="qap-badge ok">已核准</span>' : '<span class="qap-badge bad">核准已失效</span>';
  if (st === 'pending') {
    const steps = ap.steps || [];
    const cur = steps[ap.cur];
    return '<span class="qap-badge info">簽核中' + (cur ? '（第 ' + (ap.cur + 1) + ' 關：' + e(cur.label || tierLabel(cur.tier)) + (cur.assigneeName ? '　' + e(cur.assigneeName) : '') + '）' : '') + '</span>';
  }
  if (st === 'returned') return '<span class="qap-badge bad">已駁回</span>';
  return '<span class="qap-badge">未送簽</span>';
}

function validDate(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(s || ''))) return false;
  const d = new Date(s + 'T00:00:00Z');
  return !isNaN(d.getTime()) && d.toISOString().slice(0, 10) === s;
}

/** 建立標準 modal overlay；回傳 {ov, body, foot} */
function mountModal(id, title, modalClass, withFooter) {
  const ov = document.createElement('div');
  ov.className = 'modal-overlay open';
  ov.id = id;
  ov.innerHTML = `
    <div class="modal qap-modal ${modalClass || ''}" role="dialog" aria-modal="true">
      <div class="modal-header">
        <h2>${title}</h2>
        <button class="modal-close" type="button" data-act="close" aria-label="關閉">&#10005;</button>
      </div>
      <div class="qap-tabs-slot"></div>
      <div class="modal-body qap-body"></div>
      ${withFooter ? '<div class="modal-footer qap-footer"></div>' : ''}
    </div>`;
  document.body.appendChild(ov);
  return { ov, tabs: ov.querySelector('.qap-tabs-slot'), body: ov.querySelector('.qap-body'), foot: ov.querySelector('.qap-footer') };
}

// ── 自訂確認視窗（不使用 confirm()） ───────────────────────
let _confirmOpen = false;
function qapConfirm(opts) {
  if (_confirmOpen) return Promise.resolve(false);
  _confirmOpen = true;
  ensureStyle();
  return new Promise((resolve) => {
    const ov = document.createElement('div');
    ov.className = 'modal-overlay open qap-confirm-ov';
    ov.id = 'qapConfirmOv';
    const okCls = opts.danger ? 'btn btn-danger' : 'btn btn-primary';
    ov.innerHTML = `
      <div class="modal qap-modal qap-sm" style="width:460px" role="alertdialog" aria-modal="true">
        <div class="modal-header"><h2>${e(opts.title || '請確認')}</h2></div>
        <div class="modal-body"><div class="qap-confirm-msg">${e(opts.message || '')}</div></div>
        <div class="modal-footer">
          <button class="btn btn-secondary" type="button" data-r="0">${e(opts.cancelText || '取消')}</button>
          <button class="${okCls}" type="button" data-r="1">${e(opts.okText || '確定')}</button>
        </div>
      </div>`;
    let done = false;
    const finish = (v) => {
      if (done) return;
      done = true;
      document.removeEventListener('keydown', onKey, true);
      ov.remove();
      _confirmOpen = false;
      resolve(v);
    };
    const onKey = (ev) => {
      if (ev.key === 'Escape') { ev.stopPropagation(); finish(false); }
    };
    ov.addEventListener('click', (ev) => {
      const b = ev.target.closest('button[data-r]');
      if (b) finish(b.dataset.r === '1');
    });
    document.addEventListener('keydown', onKey, true);
    document.body.appendChild(ov);
    const first = ov.querySelector('button[data-r="0"]');
    if (first) first.focus();
  });
}

// ═════════════════════════════════════════════════
// 待我處理徽章
// ═════════════════════════════════════════════════
let _inboxPromise = null;
function refreshQuoteInbox() {
  if (_inboxPromise) return _inboxPromise;
  _inboxPromise = (async () => {
    try {
      const r = await apiCall('GET', '/quotations/inbox');
      if (!r.ok) return null;
      const d = r.data || {};
      const c = d.counts || {};
      const n = (Number(c.approvals) || 0) + (Number(c.costs) || 0);
      window._quoteInbox = d;   // 供 quote.js 的「待我處理」篩選使用（選用）
      const b = document.getElementById('quoteInboxBadge');
      if (b) {
        b.textContent = n > 99 ? '99+' : String(n);
        b.style.display = n > 0 ? '' : 'none';
        b.title = '待簽核 ' + (Number(c.approvals) || 0) + ' 張、待填成本 ' + (Number(c.costs) || 0) + ' 張';
      }
      return d;
    } catch (err) {
      return null;
    } finally {
      _inboxPromise = null;
    }
  })();
  return _inboxPromise;
}

// ═════════════════════════════════════════════════
// 簽核面板
// ═════════════════════════════════════════════════
let _ap = null;

function closeApproval() {
  const s = _ap;
  if (!s) return;
  _ap = null;
  s.closed = true;
  document.removeEventListener('keydown', s.onKey);
  s.ov.remove();
}

/** 取得「這張單目前要顯示的判定結果」：簽核中／已核准用送簽當時凍結的 derived，其餘用即時試算 preview */
function pickDerived(q) {
  const ap = q.approval;
  if (ap && ap.derived && (ap.state === 'pending' || ap.state === 'approved')) return { d: ap.derived, frozen: true };
  if (q.preview) return { d: q.preview, frozen: false };
  if (ap && ap.derived) return { d: ap.derived, frozen: true };
  return null;
}

function tierNames(d) {
  return (d.tiers || []).map((t) => {
    if (typeof t === 'string') return tierLabel(t);
    return (t && (t.label || tierLabel(t.tier))) || '';
  }).filter(Boolean);
}

function sec(title, inner) {
  return `<div class="qap-sec"><h3>${title}</h3>${inner}</div>`;
}

function buildItemsSection(q) {
  const perm = q.perm || {};
  const items = Array.isArray(q.items) ? q.items : [];
  const showPrice = !!perm.canSeePrice && items.some((it) => it.unitPrice !== undefined && it.unitPrice !== null);
  // 新式單（有成本明細）：items[].cost 已歸 0、沒有逐列成本與逐列毛利的意義 → 隱藏這兩欄，改看下方「專案成本明細」；舊式單完全維持原樣
  const showCost = !!perm.canSeeCost && !Array.isArray(q.costLines) && items.some((it) => it.cost !== undefined && it.cost !== null);
  let head = '<th class="c">#</th><th>品項說明</th><th class="r">數量</th><th>單位</th>';
  if (showPrice) head += '<th class="r">單價</th><th class="r">金額</th>';
  if (showCost) head += '<th class="r">成本</th>';
  if (showPrice && showCost) head += '<th class="r">逐列毛利率</th>';
  let subtotal = 0, seq = 0, seg = 0;   // seq：項目編號只算一般品項；seg：目前這一段的小計（分的整數）
  const colCount = 4 + (showPrice ? 2 : 0) + (showCost ? 1 : 0) + (showPrice && showCost ? 1 : 0);
  const rows = items.map((it) => {
    // 分組標題／小計列（kind）：標題是一列標題；小計列只在看得到價格時顯示該段合計。都不是品項（不計入總價、不編號）
    if (it.kind === 'title') { seg = 0; return `<tr><td colspan="${colCount}" class="desc" style="font-weight:700;background:rgba(127,127,127,.12)">${e(it.desc || '')}</td></tr>`; }
    if (it.kind === 'subtotal') {
      const label = String(it.desc || '').trim() || '小計';
      const tr = showPrice ? `<tr><td colspan="${colCount - (showCost ? 3 : 1)}" class="r" style="font-weight:700">${e(label)}</td><td class="r" style="font-weight:700">${e(fmtCents(seg))}</td>${showCost ? '<td></td>' + (showPrice ? '<td></td>' : '') : ''}</tr>` : '';
      seg = 0; return tr;
    }
    const qty = parseFloat(it.qty) || 1;
    const price = parseFloat(it.unitPrice) || 0;
    subtotal += Math.round(qty * price * 100);
    seg += Math.round(qty * price * 100);
    let tr = `<tr><td class="c">${++seq}</td><td class="desc">${e(it.desc || '')}</td><td class="r">${e(fmtNum(qty))}</td><td>${e(it.unit || '式')}</td>`;
    if (showPrice) tr += `<td class="r">${e(fmtNum(price))}</td><td class="r">${e(fmtNum(qty * price))}</td>`;
    if (showCost) tr += `<td class="r">${it.cost === undefined || it.cost === null ? '—' : e(fmtNum(it.cost))}</td>`;
    if (showPrice && showCost) tr += `<td class="r">${e(lineMarginText(it))}</td>`;
    return tr + '</tr>';
  }).join('');
  let extra = '';
  if (showPrice) {
    let disc = '';
    if (q.discountType === 'percent' && Number(q.discountValue) > 0) disc = `折扣：${e(q.discountValue)}%`;
    else if (q.discountType === 'amount' && Number(q.discountValue) > 0) disc = `議價金額：${e(fmtNum(q.discountValue))}`;
    extra = `<div class="qap-sum"><span>折扣前合計（未稅）<b>${e(fmtCents(subtotal))}</b></span>${disc ? `<span>${disc}</span>` : ''}</div>`;
  }
  return sec('報價品項' + (showCost ? '（含成本與毛利，僅有權者可見）' : ''),
    `<div class="qap-tablewrap"><table class="qap-table"><thead><tr>${head}</tr></thead><tbody>${rows || '<tr><td colspan="8" class="qap-muted">（沒有品項）</td></tr>'}</tbody></table></div>${extra}`);
}

/** 專案成本明細（新式單且看得到成本的人）：QCL 唯讀（分區小計、合計與印花稅、廠商欄）＋伺服器算的分區彙總；內容在 renderApproval 掛載 */
function buildCostDetailSection(q) {
  const perm = q.perm || {};
  if (!perm.canSeeCost || !Array.isArray(q.costLines)) return '';
  let extra = '';
  const cb = q.costBreakdown;
  if (cb && typeof cb === 'object') {
    const one = (k, label) => `<span>${label}<b>${e(fmtCents(cb[k]))}</b></span>`;
    extra = '<div class="qap-sum">' + one('consult', '顧問服務') + one('software', '軟體') + one('hw', '硬體') + one('travel', '差旅') + one('other', '其他（含印花稅）') + '</div>';
  }
  const risk = typeof q.contingencyPct === 'number'
    ? `<div class="qap-muted" style="margin-top:6px">風險預留 ${e(q.contingencyPct)}%：毛利分析（內部）另計，不計入簽核用的毛利率。</div>` : '';
  // 委外佔比卡（cost-sync §7.2）：數字直接取伺服器序列化的 costBreakdown（分類彙總，other 含印花稅），放在「專案成本明細」區塊上方
  let osCard = '';
  if (cb && typeof cb === 'object' && cb.outsourced !== undefined && typeof QCL !== 'undefined' && typeof QCL.outsourcedFromBreakdown === 'function') {
    osCard = '<div class="qcl-os-solo" id="qapOsWrap">' + QCL.outsourcedCardHtml(QCL.outsourcedCardModel(QCL.outsourcedFromBreakdown(cb)), 'qapOsCard') + '</div>';
  }
  return `${osCard}<div class="qap-cdwrap"><div class="qap-cdh">專案成本明細（僅有權者可見，請勿提供客戶）</div><div id="qapCostDetail"></div>${extra}${risk}</div>`;
}

function mountCostDetail(body, q) {
  const host = body.querySelector('#qapCostDetail');
  if (!host) return;
  if (typeof QCL === 'undefined') { host.innerHTML = '<div class="qap-alert bad">成本明細元件沒有載入成功，請重新整理頁面（Ctrl+F5）。</div>'; return; }
  QCL.mount(host, { mode: 'view', lines: q.costLines, items: q.items });   // items：「原報價品項已刪除」徽章的依據
}

function buildJudgeSection(q) {
  const pk = pickDerived(q);
  const perm = q.perm || {};
  if (!pk) {
    return sec('類別與簽核關卡', '<div class="qap-muted">目前沒有可顯示的判定資料（可能尚無法試算，或你無權查看）。</div>');
  }
  const d = pk.d;
  const names = tierNames(d);
  const margin = d.marginText !== null && d.marginText !== undefined ? e(d.marginText) + '%' : '—';
  let html = `<div class="qap-grid">
      <div><span class="k">類別</span><b>${e(d.rowLabel || d.rowKey || '')}</b></div>
      <div><span class="k">整單毛利率（折扣後未稅）</span><b>${margin}</b></div>
    </div>`;
  html += `<div class="qap-muted" style="margin-top:6px">${pk.frozen ? '以下為送簽當時的判定結果。' : '以下為目前即時試算結果（送簽後才會固定）。'}</div>`;
  if (names.length) {
    html += '<div class="qap-path"><span class="k qap-muted">需簽核：</span>' +
      names.map((n) => `<span class="qap-chip sel">${e(n)}</span>`).join('<span class="qap-arrow">&#10140;</span>') + '</div>';
  }
  if (d.revenueCents !== undefined && d.revenueCents !== null && (perm.canSeePrice || perm.canSeeCost)) {
    let sum = `<span>折扣後未稅<b>${e(fmtCents(d.revenueCents))}</b></span>`;
    if (perm.canSeeCost && d.costCents !== undefined && d.gpCents !== undefined) {
      sum += `<span>成本<b>${e(fmtCents(d.costCents))}</b></span><span>毛利<b>${e(fmtCents(d.gpCents))}</b></span>`;
    }
    html += `<div class="qap-sum">${sum}</div>`;
  }
  if (Array.isArray(d.reasons) && d.reasons.length) {
    html += '<ul class="qap-list">' + d.reasons.map((x) => `<li>${e(x)}</li>`).join('') + '</ul>';
  }
  const warns = [].concat(d.warnings || []);
  if (!warns.length && Array.isArray(d.unclassified) && d.unclassified.length) {   // 伺服器 warnings 通常已含同樣訊息，避免重複
    warns.push('尚未歸類的商品：' + d.unclassified.join('、') + '（已視為「其他」）');
  }
  // 建議性警告：有價品項在成本明細沒有對應的成本列（送簽當時凍結的 derived.costWarnings；即時試算時 preview.warnings 已含同一句，依文字去重）。純提醒、不擋
  (Array.isArray(d.costWarnings) ? d.costWarnings : []).forEach((w) => {
    const m = w && typeof w.message === 'string' ? w.message : '';
    if (m && warns.indexOf(m) < 0) warns.push(m);
  });
  if (warns.length) {
    html += '<div class="qap-alert warn" style="margin-top:8px">' + warns.map((w) => e(w)).join('<br>') + '</div>';
  }
  return sec('類別與簽核關卡', html);
}

function buildBlockersAlert(q) {
  const ap = q.approval;
  const canTry = (q.perm || {}).canSubmit;
  const bl = q.preview && Array.isArray(q.preview.blockers) ? q.preview.blockers : [];
  if (!canTry || !bl.length || (ap && ap.state === 'pending')) return '';
  return '<div class="qap-alert bad"><b>目前還不能送簽：</b><ul class="qap-list">' +
    bl.map((b) => `<li>${e((b && b.message) || (b && b.code) || '')}</li>`).join('') + '</ul></div>';
}

function buildCostSection(q) {
  const cf = q.costFlow || {};
  const st = cf.state || 'na';
  let line = `<span class="qap-badge ${st === 'filled' ? 'ok' : st === 'requested' ? 'warn' : ''}">${e(COST_STATE_LABEL[st] || st)}</span>`;
  if (q.costByName || cf.byName) line += `　<span class="qap-muted">支援顧問：${e(cf.byName || q.costByName)}</span>`;
  if (cf.filledAt) line += `　<span class="qap-muted">完成時間：${e(fmtTime(cf.filledAt))}</span>`;
  else if (cf.requestedAt) line += `　<span class="qap-muted">通知時間：${e(fmtTime(cf.requestedAt))}</span>`;
  const note = cf.note ? `<div class="qap-muted" style="margin-top:6px">給顧問的備註：<span style="white-space:pre-wrap">${e(cf.note)}</span></div>` : '';
  return sec('成本填寫', line + note);
}

function buildStepsSection(q) {
  const ap = q.approval;
  const steps = (ap && Array.isArray(ap.steps)) ? ap.steps : [];
  if (!steps.length) return '';
  const icon = { approved: '&#10003;', pending: '&#9679;', returned: '&#10005;', waiting: '&#9675;' };
  const li = steps.map((st) => {
    const who = st.status === 'approved' || st.status === 'returned' ? (st.byName || st.assigneeName || '') : (st.assigneeName || '');
    const when = st.at ? '　' + fmtTime(st.at) : '';
    return `<li class="${e(st.status)}"><span class="ic">${icon[st.status] || '&#9675;'}</span>
      <div><div><b>${e(st.label || tierLabel(st.tier))}</b>　<span class="qap-muted">${e(STEP_STATUS[st.status] || st.status || '')}${who ? '　' + e(who) : ''}${e(when)}</span></div>
      ${st.comment ? `<div class="cm">${e(st.comment)}</div>` : ''}</div></li>`;
  }).join('');
  let board = '';
  if (ap.board && (ap.board.resolutionDate || ap.board.resolutionNo)) {
    board = `<div class="qap-alert info" style="margin-top:8px">董事會決議：日期 ${e(ap.board.resolutionDate || '')}　文號 ${e(ap.board.resolutionNo || '')}${ap.board.at ? '　（登錄時間 ' + e(fmtTime(ap.board.at)) + '）' : ''}</div>`;
  }
  return sec('簽核進度', `<ul class="qap-steps">${li}</ul>${board}`);
}

function buildHistorySection(q) {
  const h = q.approval && Array.isArray(q.approval.history) ? q.approval.history : [];
  if (!h.length) return '';
  const rows = h.map((x) => `<tr><td class="r">${e(fmtTime(x.at))}</td><td>${e(x.byName || '')}</td><td>${e(ACTION_LABEL[x.action] || x.action || '')}</td><td class="desc">${e(x.comment || '')}</td></tr>`).join('');
  return sec('簽核歷程', `<div class="qap-tablewrap"><table class="qap-table"><thead><tr><th class="r">時間</th><th>人員</th><th>動作</th><th>簽核建議／駁回原因</th></tr></thead><tbody>${rows}</tbody></table></div>`);
}

/** 目前輪到的關卡是否為董事會（僅決定「要不要顯示決議欄位」，能不能簽仍看 perm） */
function currentIsBoard(q) {
  const ap = q.approval;
  if (!ap || ap.state !== 'pending' || !Array.isArray(ap.steps)) return false;
  const st = ap.steps[ap.cur];
  return !!st && st.tier === 'board';
}

function findBoardStep(q) {
  const steps = q.approval && Array.isArray(q.approval.steps) ? q.approval.steps : [];
  return steps.find((x) => x.tier === 'board') || null;
}

function canPrintBoard(q) {
  const perm = q.perm || {};
  const bs = findBoardStep(q);
  const pk = pickDerived(q);
  return !!(bs && (bs.status === 'pending' || bs.status === 'approved') && pk && pk.frozen && (perm.canApprove || perm.canSeeCost));
}

function buildActionSection(s) {
  const q = s.q;
  const perm = q.perm || {};
  const dr = s.draft;
  const parts = [];
  const boardNow = currentIsBoard(q);
  if (perm.canApprove || perm.canReturn) {
    let f = '';
    if (perm.canApprove && boardNow) {
      f += `<div><label>董事會決議日期（必填）</label><input type="date" data-draft="resDate" value="${e(dr.resDate)}"></div>
            <div><label>董事會決議文號（必填）</label><input type="text" maxlength="60" data-draft="resNo" value="${e(dr.resNo)}" placeholder="例：113-董-045"></div>`;
    }
    f += `<div class="full"><label>簽核建議${perm.canReturn ? '（核准時選填；駁回時必填）' : '（選填）'}</label>
          <textarea data-draft="comment" maxlength="1000" placeholder="請輸入簽核建議（核准時可不填；駁回時必填原因）">${e(dr.comment)}</textarea></div>`;
    parts.push(`<div class="qap-form">${f}</div>`);
  }
  if (perm.canReassign) {
    const opts = (s.mgrs || []).map((u) => `<option value="${e(u.username)}"${dr.reassign === u.username ? ' selected' : ''}>${e(u.displayName || u.username)}（${e(u.username)}）</option>`).join('');
    parts.push(`<div class="qap-form" style="margin-top:10px"><div class="full"><label>改派一級主管關（管理員只能改派，不能代簽）</label>
      <div style="display:flex;gap:8px;flex-wrap:wrap"><select data-draft="reassign" style="flex:1;min-width:200px"><option value="">請選擇一級主管…</option>${opts}</select>
      <button class="btn btn-secondary" type="button" data-act="reassign"${s.busy ? ' disabled' : ''}>改派</button></div>
      ${(s.mgrs || []).length ? '' : '<div class="qap-muted" style="margin-top:4px">找不到可改派的一級主管帳號。</div>'}</div></div>`);
  }
  if (!parts.length) return '';
  return sec('你的簽核動作', parts.join(''));
}

function renderApproval(s) {
  if (s.closed) return;
  const q = s.q;
  const keep = s.body.scrollTop;
  if (!q) {
    s.body.innerHTML = '<div class="qap-muted" style="padding:30px;text-align:center">載入中…</div>';
    s.foot.innerHTML = '<button class="btn btn-secondary" type="button" data-act="close">關閉</button>';
    return;
  }
  const ap = q.approval;
  const perm = q.perm || {};
  let html = '';
  if (ap && ap.state === 'approved' && !ap.valid) {
    html += '<div class="qap-alert bad">此報價單的核准已失效（核准後內容被修改）。下載的 PDF 不會有報價專用章，需重新送簽。</div>';
  }
  if (ap && ap.state === 'returned') {
    const h = Array.isArray(ap.history) ? ap.history.filter((x) => x.action === 'RETURN') : [];
    const last = h.length ? h[h.length - 1] : null;
    html += `<div class="qap-alert warn">此報價單被駁回${last && last.byName ? '（' + e(last.byName) + '）' : ''}${last && last.comment ? '：' + e(last.comment) : ''}。修改後可重新送簽。</div>`;
  }
  html += buildBlockersAlert(q);
  html += sec('報價摘要', `<div class="qap-grid">
      <div><span class="k">報價單號</span><b>${e(q.quoteNo || '')}</b></div>
      <div><span class="k">狀態</span>${approvalBadge(q)}</div>
      <div><span class="k">客戶</span>${e(q.company || '')}</div>
      <div><span class="k">專案名稱</span>${e(q.projectName || '')}</div>
      ${q.validUntil ? `<div><span class="k">報價期限</span>${e(q.validUntil)}</div>` : ''}
      <div><span class="k">業務</span>${e(q.ownerName || q.owner || '')}</div>
      <div><span class="k">商品</span>${(q.products || []).map((p) => `<span class="qap-chip">${e(p)}</span>`).join(' ') || '<span class="qap-muted">（未勾選）</span>'}</div>
    </div>`);
  html += buildItemsSection(q);
  html += buildCostDetailSection(q);
  html += buildJudgeSection(q);
  html += buildCostSection(q);
  html += buildStepsSection(q);
  html += buildActionSection(s);
  html += buildHistorySection(q);
  s.body.innerHTML = html;
  mountCostDetail(s.body, q);
  s.body.scrollTop = keep;

  // 頁尾按鈕：顯示與否完全由 perm 決定
  const dis = s.busy ? ' disabled' : '';
  const boardNow = currentIsBoard(q);
  let f = '<button class="btn btn-secondary" type="button" data-act="close"' + dis + '>關閉</button><span class="sp"></span>';
  if (canPrintBoard(q)) f += `<button class="btn btn-secondary" type="button" data-act="print"${dis}>&#128438; 列印送董事會</button>`;
  if (perm.canWithdraw) f += `<button class="btn btn-secondary" type="button" data-act="withdraw"${dis}>撤回</button>`;
  if (perm.canSubmit) f += `<button class="btn btn-primary" type="button" data-act="submit"${dis}>送簽</button>`;
  if (perm.canReturn) f += `<button class="btn btn-danger" type="button" data-act="return"${dis}>駁回</button>`;
  if (perm.canApprove) f += `<button class="btn btn-export" type="button" data-act="approve"${dis}>${boardNow ? '登錄董事會決議並代為核准' : '核准'}</button>`;
  s.foot.innerHTML = f;
}

/** 重新載入這張單；失敗回傳 false（已 toast）。不負責 render */
async function fetchApprovalQuote(s) {
  const r = await apiCall('GET', '/quotations/' + encodeURIComponent(s.id));
  if (s.closed) return false;
  if (!r.ok || !r.data || !r.data.id) {
    toast(errMsg(r));
    return false;
  }
  s.q = r.data;
  if (s.q.perm && s.q.perm.canReassign && !s.mgrs) {
    const c = await apiCall('GET', '/quote-approval/config');
    if (s.closed) return false;
    const users = c.ok && Array.isArray(c.data.users) ? c.data.users : [];
    s.mgrs = users.filter((u) => u.role === 'manager1' && u.active !== false);
  }
  return true;
}

async function openQuoteApproval(id) {
  ensureStyle();
  closeApproval();
  const m = mountModal('quoteApprovalOverlay', '報價單簽核', '', true);
  const s = {
    id, q: null, busy: true, closed: false, ov: m.ov, body: m.body, foot: m.foot,
    draft: { comment: '', resDate: '', resNo: '', reassign: '' }, mgrs: null,
  };
  _ap = s;
  s.onKey = (ev) => {
    if (ev.key !== 'Escape') return;
    if (!s.ov.isConnected) { document.removeEventListener('keydown', s.onKey); return; }
    if (_confirmOpen) return;
    closeApproval();
  };
  document.addEventListener('keydown', s.onKey);

  s.ov.addEventListener('click', (ev) => {
    const t = ev.target.closest('[data-act]');
    if (!t || !s.ov.contains(t) || t.disabled) return;
    handleApprovalAction(s, t.dataset.act);
  });
  const onDraft = (ev) => {
    const t = ev.target;
    if (t && t.dataset && t.dataset.draft) s.draft[t.dataset.draft] = t.value;
  };
  s.ov.addEventListener('input', onDraft);
  s.ov.addEventListener('change', onDraft);

  renderApproval(s);
  const ok = await fetchApprovalQuote(s);
  if (s.closed) return;
  if (!ok) { closeApproval(); return; }
  s.busy = false;
  renderApproval(s);
}

/** 送出動作的共用流程：busy 鎖 → 呼叫 → 無論成敗都重載 → 解鎖並重繪 */
async function runApprovalAction(s, method, path, body, okMsg) {
  if (s.busy) return;
  s.busy = true;
  renderApproval(s);
  const res = await apiCall(method, '/quotations/' + encodeURIComponent(s.id) + path, body);
  if (s.closed) { if (res.ok) afterMutation(); return; }
  let success = false;
  if (res.ok) {
    success = true;
    toast(okMsg);
  } else if (res.status === 409 && res.data && res.data.code === 'CONTENT_CHANGED') {
    toast('內容已被修改，請重新檢視後再簽核');
  } else if (res.status === 409 && res.data && res.data.code === 'LOCKED_PENDING') {
    toast(res.data.error || '此單簽核中，已被鎖定');
  } else {
    toast(errMsg(res), 4500);
  }
  // 成敗都重新載入：409 代表伺服器狀態已變；網路錯誤時也無從確定是否已生效
  await fetchApprovalQuote(s);
  if (s.closed) { if (success) afterMutation(); return; }
  if (success) {
    s.draft.comment = ''; s.draft.resDate = ''; s.draft.resNo = ''; s.draft.reassign = '';
  }
  s.busy = false;
  renderApproval(s);
  if (success || res.status === 409) afterMutation();
}

async function handleApprovalAction(s, act) {
  if (act === 'close') { closeApproval(); return; }   // 動作進行中也允許關閉；回應回來時以 s.closed 判斷
  if (s.busy || !s.q) return;
  const q = s.q;
  const perm = q.perm || {};
  const dr = s.draft;

  if (act === 'print') {
    printBoardMemo(q);
    return;
  }
  if (act === 'submit') {
    if (!perm.canSubmit) return;
    const ok = await qapConfirm({
      title: '確認送簽',
      message: '送簽後整張報價單會被鎖定，需先撤回才能修改。\n簽核路徑與毛利判定由系統依目前內容計算。\n\n確定要送簽嗎？',
      okText: '送簽',
    });
    if (!ok || s.closed) return;
    runApprovalAction(s, 'POST', '/submit', {}, '已送簽');
    return;
  }
  if (act === 'withdraw') {
    if (!perm.canWithdraw) return;
    const ok = await qapConfirm({ title: '確認撤回', message: '撤回後已簽的關卡會作廢（歷程保留），需修改後重新送簽。\n\n確定要撤回嗎？', okText: '撤回', danger: true });
    if (!ok || s.closed) return;
    runApprovalAction(s, 'POST', '/withdraw', {}, '已撤回');
    return;
  }
  if (act === 'reassign') {
    if (!perm.canReassign) return;
    if (!dr.reassign) { toast('請先選擇要改派的一級主管'); return; }
    const u = (s.mgrs || []).find((x) => x.username === dr.reassign);
    const ok = await qapConfirm({ title: '確認改派', message: '將目前的一級主管關改派給「' + (u ? (u.displayName || u.username) : dr.reassign) + '」？\n（改派不等於代簽）', okText: '改派' });
    if (!ok || s.closed) return;
    runApprovalAction(s, 'POST', '/reassign', { username: dr.reassign }, '已改派');
    return;
  }
  if (act === 'approve' || act === 'return') {
    if (act === 'approve' && !perm.canApprove) return;
    if (act === 'return' && !perm.canReturn) return;
    const hash = q.contentHash;
    if (!hash) {
      toast('無法取得內容雜湊，已重新載入，請再檢視一次');
      s.busy = true; renderApproval(s);
      await fetchApprovalQuote(s);
      if (s.closed) return;
      s.busy = false; renderApproval(s);
      return;
    }
    const comment = String(dr.comment || '').trim();
    if (act === 'return') {
      if (!comment) {
        toast('駁回時必須填寫原因');
        const ta = s.ov.querySelector('[data-draft="comment"]');
        if (ta) ta.focus();
        return;
      }
      const ok = await qapConfirm({ title: '確認駁回', message: '駁回後業務需修改並重新送簽。\n\n駁回原因：\n' + comment, okText: '駁回', danger: true });
      if (!ok || s.closed) return;
      runApprovalAction(s, 'POST', '/return', { hash, comment }, '已駁回');
      return;
    }
    // approve
    const body = { hash };
    if (comment) body.comment = comment;
    let msg = '確認核准此報價單？\n核准後依序進入下一關；若為最後一關，報價單即核准完成。';
    if (currentIsBoard(q)) {
      const d = String(dr.resDate || '').trim();
      const n = String(dr.resNo || '').trim();
      if (!validDate(d)) { toast('請填寫正確的董事會決議日期'); const i = s.ov.querySelector('[data-draft="resDate"]'); if (i) i.focus(); return; }
      if (!n) { toast('請填寫董事會決議文號'); const i = s.ov.querySelector('[data-draft="resNo"]'); if (i) i.focus(); return; }
      body.resolutionDate = d;
      body.resolutionNo = n;
      msg = '確認代為登錄董事會決議並核准？\n決議日期：' + d + '\n決議文號：' + n + '\n\n請確認與實際董事會紀錄一致。';
    }
    const ok = await qapConfirm({ title: '確認核准', message: msg, okText: '核准' });
    if (!ok || s.closed) return;
    runApprovalAction(s, 'POST', '/approve', body, '已核准');
  }
}

// ── 董事會簽呈列印 ─────────────────────────────────
function buildBoardPrintHtml(q) {
  const ap = q.approval || {};
  const d = ap.derived || {};
  const perm = q.perm || {};
  const steps = Array.isArray(ap.steps) ? ap.steps : [];
  const today = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
  const money = (c) => 'NT$ ' + fmtCents(c);
  const stepRows = steps.map((st) => {
    const isBoard = st.tier === 'board';
    const who = st.status === 'approved' ? (st.byName || st.assigneeName || '') : (isBoard ? '管理部秘書（代理登錄）' : (st.assigneeName || ''));
    return `<tr><td>${e(st.label || tierLabel(st.tier))}</td><td>${e(who)}</td><td>${e(STEP_STATUS[st.status] || '')}</td><td>${e(st.at ? fmtTime(st.at) : '')}</td><td>${e(st.comment || '')}</td></tr>`;
  }).join('');
  const reasons = Array.isArray(d.reasons) ? d.reasons.map((x) => `<li>${e(x)}</li>`).join('') : '';
  const items = Array.isArray(q.items) ? q.items : [];
  const showPrice = !!perm.canSeePrice;
  let seq = 0, seg = 0;   // 項目編號只算一般品項；seg＝目前這一段的小計（分）
  const memoCols = 4 + (showPrice ? 2 : 0);
  const itemRows = items.map((it) => {
    if (it.kind === 'title') { seg = 0; return `<tr><td colspan="${memoCols}" style="font-weight:700;background:#f2f2f2">${e(it.desc || '')}</td></tr>`; }
    if (it.kind === 'subtotal') {
      const tr = showPrice ? `<tr><td colspan="${memoCols - 1}" class="r" style="font-weight:700">${e(String(it.desc || '').trim() || '小計')}</td><td class="r" style="font-weight:700">${e(fmtCents(seg))}</td></tr>` : '';
      seg = 0; return tr;
    }
    const qty = parseFloat(it.qty) || 1;
    const price = parseFloat(it.unitPrice) || 0;
    seg += Math.round(qty * price * 100);
    return `<tr><td class="c">${++seq}</td><td>${e(it.desc || '')}</td><td class="r">${e(fmtNum(qty))}</td><td>${e(it.unit || '式')}</td>${showPrice ? `<td class="r">${e(fmtNum(price))}</td><td class="r">${e(fmtNum(qty * price))}</td>` : ''}</tr>`;
  }).join('');
  const b = ap.board || {};
  const resDate = b.resolutionDate ? e(b.resolutionDate) : '________ 年 ____ 月 ____ 日';
  const resNo = b.resolutionNo ? e(b.resolutionNo) : '________________';
  const ref = q.contentHash ? e(String(q.contentHash).slice(0, 12)) : '';
  return `<!doctype html><html lang="zh-Hant"><head><meta charset="utf-8"><title>董事會簽呈 ${e(q.quoteNo || '')}</title>
<style>
@page { size: A4; margin: 16mm; }
* { box-sizing: border-box; }
body { margin: 0; color: #111; background: #fff; font-family: "Microsoft JhengHei","微軟正黑體","PingFang TC","Noto Sans TC",sans-serif; font-size: 13px; line-height: 1.6; }
h1 { font-size: 22px; text-align: center; margin: 0 0 4px; letter-spacing: 3px; }
.sub { text-align: center; font-size: 13px; margin-bottom: 14px; }
.tag { text-align: center; margin: 0 0 14px; font-size: 12px; color: #444; }
h2 { font-size: 14px; margin: 16px 0 6px; padding-bottom: 3px; border-bottom: 1.5px solid #111; }
table { width: 100%; border-collapse: collapse; }
th, td { border: 1px solid #111; padding: 4px 7px; text-align: left; vertical-align: top; font-size: 12.5px; }
th { background: #f0f0f0; }
td.r { text-align: right; } td.c { text-align: center; }
.kv td:first-child { width: 22%; background: #f7f7f7; font-weight: 700; }
ul { margin: 4px 0 0; padding-left: 20px; }
.sign td { height: 70px; }
.note { margin-top: 18px; font-size: 11.5px; color: #333; border-top: 1px dashed #666; padding-top: 6px; }
</style></head><body>
<h1>董事會簽呈</h1>
<div class="sub">東捷資訊服務股份有限公司　報價單核決</div>
<div class="tag">列印日期：${e(today)}${ref ? '　內容識別碼：' + ref : ''}</div>
<h2>案件資料</h2>
<table class="kv">
<tr><td>報價單號</td><td>${e(q.quoteNo || '')}</td><td style="width:16%;background:#f7f7f7;font-weight:700">承辦業務</td><td>${e(q.ownerName || q.owner || '')}</td></tr>
<tr><td>客戶</td><td colspan="3">${e(q.company || '')}</td></tr>
<tr><td>專案名稱</td><td colspan="3">${e(q.projectName || '')}</td></tr>
${q.validUntil ? `<tr><td>報價期限</td><td colspan="3">${e(q.validUntil)}</td></tr>` : ''}
<tr><td>案件類別</td><td colspan="3">${e(d.rowLabel || '')}${(q.products || []).length ? '（商品：' + e((q.products || []).join('、')) + '）' : ''}</td></tr>
</table>
<h2>金額與毛利</h2>
<table class="kv">
<tr><td>折扣後未稅金額</td><td>${d.revenueCents !== undefined ? e(money(d.revenueCents)) : ''}</td><td style="width:16%;background:#f7f7f7;font-weight:700">毛利率</td><td>${d.marginText !== undefined && d.marginText !== null ? e(d.marginText) + '%' : ''}</td></tr>
<tr><td>總成本</td><td>${d.costCents !== undefined ? e(money(d.costCents)) : ''}</td><td style="background:#f7f7f7;font-weight:700">毛利</td><td>${d.gpCents !== undefined ? e(money(d.gpCents)) : ''}</td></tr>
</table>
${reasons ? `<div style="margin-top:6px"><b>需提董事會的原因 / 判定說明：</b><ul>${reasons}</ul></div>` : ''}
${itemRows ? `<h2>報價品項</h2><table><thead><tr><th class="c" style="width:6%">#</th><th>品項說明</th><th class="r" style="width:9%">數量</th><th style="width:8%">單位</th>${showPrice ? '<th class="r" style="width:13%">單價</th><th class="r" style="width:14%">金額</th>' : ''}</tr></thead><tbody>${itemRows}</tbody></table>` : ''}
<h2>簽核經過</h2>
<table><thead><tr><th>關卡</th><th>簽核人</th><th>狀態</th><th>時間</th><th>簽核建議／駁回原因</th></tr></thead><tbody>${stepRows}</tbody></table>
<h2>董事會決議</h2>
<table class="kv">
<tr><td>決議日期</td><td>${resDate}</td><td style="width:16%;background:#f7f7f7;font-weight:700">決議文號</td><td>${resNo}</td></tr>
<tr><td>決議結果</td><td colspan="3">&#9744; 同意核准　　&#9744; 不同意　　&#9744; 附條件同意：______________________</td></tr>
</table>
<table class="sign" style="margin-top:10px"><tr><th style="width:25%">董事長</th><th style="width:25%">董事會秘書／管理部</th><th>備註</th></tr><tr><td></td><td></td><td></td></tr></table>
<div class="note">本文件由 ITTS-CRM 產生供列印使用，系統不會自動寄出，請自行列印或另存 PDF 後送交董事會。董事會決議須由管理部秘書於系統登錄決議日期與文號後，本報價單才算核准完成。</div>
</body></html>`;
}

function printBoardMemo(q) {
  const ap = q.approval || {};
  if (!ap.derived) { toast('尚無簽呈資料可列印'); return; }
  // 注意：站台 CSP 設 frame-src 'none'，不能用 iframe 列印，只能開新視窗
  const w = window.open('', '_blank');
  if (!w) { toast('瀏覽器擋住了彈出視窗，請允許本站彈出視窗後再按一次', 4500); return; }
  w.document.open();
  w.document.write(buildBoardPrintHtml(q));
  w.document.close();
  setTimeout(() => { try { w.focus(); w.print(); } catch (err) { /* 使用者可自行 Ctrl+P */ } }, 600);
}

// ═════════════════════════════════════════════════
// 顧問填成本（成本明細編輯器 QCL：與業務自填成本共用同一套編輯器與資料模型 q.costLines）
// ═════════════════════════════════════════════════
let _cf = null;

function closeCostFill() {
  const s = _cf;
  if (!s) return;
  _cf = null;
  s.closed = true;
  document.removeEventListener('keydown', s.onKey);
  if (s.ed) { try { s.ed.destroy(); } catch (err) { /* 容器隨 overlay 一起移除 */ } s.ed = null; }
  s.ov.remove();
}

async function requestCloseCostFill(s) {
  if (s.closed) return;
  if (s.dirty) {
    const ok = await qapConfirm({ title: '尚未儲存', message: '你輸入的成本明細還沒有儲存，確定要關閉嗎？', okText: '放棄並關閉', danger: true });
    if (!ok || s.closed) return;
  }
  closeCostFill();
}

/** 勾選商品的類別代碼（鏡像伺服器「單一類別就全歸該區」的規則，給 QCL 種子用）；取不到商品歸類設定就回空陣列 */
function costClassCodes(q) {
  let pc = null;
  try { pc = (typeof quoteCfg === 'function' ? quoteCfg().productClasses : null); } catch (err) { pc = null; }
  pc = pc || {};
  return (Array.isArray(q.products) ? q.products : []).map((n) => pc[n] && pc[n].cls).filter(Boolean);
}

/**
 * 把伺服器給的印花稅金額補回列。QCL 的 getLines／collect 輸出的印花稅列不帶金額（伺服器會自己算），
 * 但重畫時要補回去，編輯器才不會退回顯示「系統依合約金額自動計算」。
 */
function withServerStamp(lines, serverLines) {
  const st = (Array.isArray(serverLines) ? serverLines : []).find((l) => l && l.auto === 'stamp' && typeof l.unitCost === 'number');
  return (lines || []).map((l) => (l && l.auto === 'stamp' && st && typeof l.unitCost !== 'number' ? Object.assign({}, l, { unitCost: st.unitCost }) : l));
}

/** 上方「報價品項（參考，無價格）」：只列說明／數量／單位；分組標題當段落標題、小計列不顯示（顧問看不到品項單價） */
function buildCostRefItems(q, open) {
  const items = Array.isArray(q.items) ? q.items : [];
  let seq = 0;
  const rows = items.map((it) => {
    if (it.kind === 'title') return `<tr><td colspan="4" class="desc" style="font-weight:700;background:rgba(127,127,127,.12)">${e(it.desc || '')}</td></tr>`;
    if (it.kind === 'subtotal') return '';
    return `<tr><td class="c">${++seq}</td><td class="desc">${e(it.desc || '')}</td><td class="r">${e(fmtNum(parseFloat(it.qty) || 1))}</td><td>${e(it.unit || '式')}</td></tr>`;
  }).join('');
  return `<details class="qap-sec qap-ref"${open ? ' open' : ''}>
    <summary>報價品項（參考，無價格）<span class="qap-muted">　共 ${seq} 項。這是客戶看到的品項；下方成本明細不必和它一一對應</span></summary>
    <div class="qap-tablewrap"><table class="qap-table"><thead><tr><th class="c">#</th><th>品項說明</th><th class="r">數量</th><th>單位</th></tr></thead><tbody>${rows || '<tr><td colspan="4" class="qap-muted">（沒有品項）</td></tr>'}</tbody></table></div>
  </details>`;
}

/** 舊式單（沒有成本明細）而且不能編輯時的唯讀成本表（維持改版前的樣子） */
function buildCostLegacyTable(q) {
  let seq = 0;
  const rows = (Array.isArray(q.items) ? q.items : []).map((it) => {
    if (it.kind === 'title') return `<tr><td colspan="5" class="desc" style="font-weight:700;background:rgba(127,127,127,.12)">${e(it.desc || '')}</td></tr>`;
    if (it.kind === 'subtotal') return '';
    return `<tr><td class="c">${++seq}</td><td class="desc">${e(it.desc || '')}</td><td class="r">${e(fmtNum(parseFloat(it.qty) || 1))}</td><td>${e(it.unit || '式')}</td>
      <td class="r">${it.cost === undefined || it.cost === null ? '—' : e(fmtNum(it.cost))}</td></tr>`;
  }).join('');
  return `<div class="qap-tablewrap"><table class="qap-table"><thead><tr><th class="c">#</th><th>品項說明</th><th class="r">數量</th><th>單位</th><th class="r">成本</th></tr></thead><tbody>${rows || '<tr><td colspan="5" class="qap-muted">（沒有品項）</td></tr>'}</tbody></table></div>`;
}

/** 重畫前先收起編輯器裡尚未儲存的明細（儲存失敗後重畫，使用者輸入要保留），再銷毀舊編輯器 */
function stashCostDraft(s) {
  if (s.ed && !s.ed.dead) {
    if (s.dirty) s.draftLines = withServerStamp(s.ed.getLines(), s.q && s.q.costLines);
    try { s.ed.destroy(); } catch (err) { /* ignore */ }
  }
  s.ed = null;
}

// ═════════════════════════════════════════════════
// cost-sync：顧問對話框的即時試算、報價品項區、「對應」連動與預覽（規格 §4、§7.2、§7.4）
// 版面（可編輯且看得到單價時）：案件 → 即時試算（5 張卡＋簽核層級預覽＋兩個預覽按鈕）→ 報價品項（含新增報價項目）→ 成本明細（每列「對應」、顧問姓名、委外標籤）→ 風險預留。
// 數字來源：卡片／品項區是前端即時算（QCL.liveSummary，與伺服器逐分相同）；簽核層級預覽與預覽內容來自伺服器 POST cost-draft/*（debounce、同時一個請求、過期回應丟棄）。
// ═════════════════════════════════════════════════
const CF_SUM_DEBOUNCE = 650;   // 草稿試算端點的最短間隔（ms）。全站 /api 限流是每分鐘 300 次，這裡 debounce＋同時只留一個進行中的請求，最壞約 90 次/分
let _cfNidSeq = 0;
/** 顧問新增的報價項目的暫時代號（伺服器只收 A-Za-z0-9_.- ≤64；按「完成」時換成真正的 lid） */
function cfNewNid() { return 'nid-' + Date.now().toString(36) + '-' + (++_cfNidSeq).toString(36) + Math.random().toString(36).slice(2, 6); }
const cfMoney = (n) => 'NT$ ' + Math.round(Number(n) || 0).toLocaleString('en-US');

/** 這張單的對話框是否走完整的連動畫面：可編輯成本、看得到單價（伺服器放寬後的被指派顧問）、QCL 有連動函式 */
function cfCanSync(s) {
  const q = s && s.q;
  const perm = (q && q.perm) || {};
  if (!q || !perm.canEditCost || !perm.canSeePrice || typeof QCL === 'undefined' || typeof QCL.liveSummary !== 'function') return false;
  return (Array.isArray(q.items) ? q.items : []).some((it) => it && it.kind !== 'title' && it.kind !== 'subtotal' && it.unitPrice !== undefined);
}

/** 草稿新增的報價項目（畫面狀態 → 計算／送出用的乾淨物件）：qty 非有限數當 0（完成時會被擋） */
function cfNewItemsClean(s) {
  return (s.newItems || []).map((n) => ({ nid: n.nid, desc: String(n.desc || '').trim(), unit: String(n.unit || '').trim() || '式', qty: Number.isFinite(Number(n.qty)) ? Number(n.qty) : 0 }));
}

/** 交給成本明細編輯器的「報價品項」：業務的品項＋顧問草稿新增的（isDraftNew，用暫時代號 nid 當對應目標） */
function cfEditorItems(s) {
  const base = Array.isArray(s.q.items) ? s.q.items : [];
  return base.concat(cfNewItemsClean(s).map((n) => ({ nid: n.nid, desc: n.desc, unit: n.unit, qty: n.qty, unitPrice: 0, needPrice: true, isDraftNew: true })));
}

/** 「顧問姓名」欄的建議清單：簽核設定的顧問名單顯示名＋這張單的成本填寫人＋已填過的姓名（取不到名單就只有後兩者，不影響填寫） */
function cfConsultantNames(s) {
  const out = [];
  const add = (x) => { const t = String(x || '').trim(); if (t && out.indexOf(t) < 0) out.push(t); };
  try { ((typeof quoteCfg === 'function' ? quoteCfg().costProviders : null) || []).forEach((p) => add(p && p.displayName)); } catch (err) { /* 沒有名單就沒有建議 */ }
  if (s.q) { add(s.q.costByName); add(s.q.costFlow && s.q.costFlow.byName); (s.q.costLines || []).forEach((l) => add(l && l.consultant)); }
  return out;
}

/** 即時試算區（5 張卡＋簽核層級預覽＋兩個預覽按鈕）；內容由 cfPaint／cfPaintPv 填 */
function cfLiveHtml() {
  const os = QCL.outsourcedCardHtml(QCL.outsourcedCardModel(null), 'qapCfOs');
  return `<div class="qap-cf-live" id="qapCfLive">
    <div class="qap-cf-bar"><h3>即時試算</h3><span class="qap-muted">隨下方調整即時更新（含連動後的報價）；簽核層級由伺服器試算</span><span class="sp"></span>
      <button class="btn btn-secondary btn-sm" type="button" data-act="pvQuote" id="qapCfPvQuote" title="用目前畫面上的調整（含尚未儲存的）預覽客戶看到的報價單">報價單預覽</button>
      <button class="btn btn-secondary btn-sm" type="button" data-act="pvPnl" id="qapCfPvPnl" title="用目前畫面上的調整（含尚未儲存的）預覽毛利分析（內部）">毛利預覽</button></div>
    <div class="pnl-summary-bar"><div class="pnl-sum-grid qcl-g5" id="qapCfCards">
      <div class="pnl-sum-card"><div class="pnl-sum-label">報價合計（未稅，折扣後）</div><div class="pnl-sum-value" id="qapCfRev">NT$ 0</div></div>
      <div class="pnl-sum-card"><div class="pnl-sum-label">成本合計</div><div class="pnl-sum-value" id="qapCfCost">NT$ 0</div></div>
      <div class="pnl-sum-card"><div class="pnl-sum-label">毛利</div><div class="pnl-sum-value pnl-gp-val" id="qapCfGp">NT$ 0</div></div>
      <div class="pnl-sum-card pnl-margin-card"><div class="pnl-sum-label">整體毛利率（畫面試算）</div><div class="pnl-sum-value pnl-margin-val" id="qapCfMargin">—</div></div>
      ${os}
    </div>
    <div class="qap-cf-pv" id="qapCfPv" data-state="pending"><div class="qap-cf-pvbody"><span class="qap-muted">簽核層級試算中…</span></div></div>
    </div></div>`;
}

/** 報價品項區的列：依業務原本的順序（分組標題／小計列照列），最後接顧問新增的報價項目（可編輯）；動態欄位由 cfUpdateItems 填 */
function cfItemRowsHtml(s) {
  const items = Array.isArray(s.q.items) ? s.q.items : [];
  let seq = 0;
  let h = items.map((it, i) => {
    if (!it) return '';
    if (it.kind === 'title') return `<tr class="qci-title"><td colspan="9">${e(it.desc || '')}</td></tr>`;
    if (it.kind === 'subtotal') return `<tr class="qci-segrow" data-ci-sub="${i}"><td colspan="5" class="r">${e(String(it.desc || '').trim() || '小計')}</td><td class="r qci-seg">0</td><td colspan="3"></td></tr>`;
    return `<tr class="qci-row" data-ci="${e(it.lid || '')}" data-idx="${i}"><td class="c">${++seq}</td>
      <td class="desc">${e(it.desc || '')}<span class="qci-flags"></span></td><td class="qci-u"></td><td class="r qci-q"></td><td class="r qci-p"></td><td class="r qci-s"></td><td class="qci-n"></td><td class="r qci-c"></td><td></td></tr>`;
  }).join('');
  h += (s.newItems || []).map((n) => `<tr class="qci-row qci-new" data-ci="${e(n.nid)}" data-nid="${e(n.nid)}"><td class="c">${++seq}</td>
      <td class="desc"><input type="text" class="qap-input qci-in qci-nd${n.descBad ? ' qci-bad' : ''}" maxlength="120" value="${e(n.desc)}" placeholder="新增的報價項目品名" aria-label="新增的報價項目：品名" style="min-width:160px"><span class="qci-flags"></span></td>
      <td><input type="text" class="qap-input qci-in qci-nu" maxlength="10" value="${e(n.unit)}" aria-label="新增的報價項目：單位"></td>
      <td class="r"><input type="number" class="qap-input qci-in qci-nq${n.qtyBad ? ' qci-bad' : ''}" min="0" step="any" inputmode="decimal" value="${e(String(n.qtyRaw !== undefined ? n.qtyRaw : n.qty))}" aria-label="新增的報價項目：數量"></td>
      <td class="r qci-p"></td><td class="r qci-s"></td><td class="qci-n"></td><td class="r qci-c"></td>
      <td><button type="button" class="qci-rm" data-act="rmNewItem" data-nid="${e(n.nid)}" title="移除此新增的報價項目" aria-label="移除此新增的報價項目">&#10005;</button></td></tr>`).join('');
  return h || '<tr><td colspan="9" class="qap-muted">（沒有品項）</td></tr>';
}

function cfItemsSectionHtml(s) {
  const head = '<th class="c">#</th><th>品項說明</th><th>單位</th><th class="r">數量</th><th class="r">單價</th><th class="r">小計</th><th>對應成本列</th><th class="r">成本小計</th><th></th>';
  return `<details class="qap-sec qap-ref" id="qapCfRef"${s.refOpen !== false ? ' open' : ''}>
    <summary>報價品項<span class="qap-muted">　客戶看到的品項；成本明細不必和它一一對應。數量被「連動」改動時會標示「業務原值 → 新值」</span></summary>
    <div class="qap-tablewrap"><table class="qap-table qap-ci"><thead><tr>${head}</tr></thead><tbody id="qapCfItems">${cfItemRowsHtml(s)}</tbody></table></div>
    <div class="qap-cf-newbar"><button class="btn btn-secondary btn-sm" type="button" data-act="addNewItem" id="qapCfAddItem">＋新增報價項目</button>
      <span class="qap-muted">顧問新增的品項單價由業務補填，補完單價前業務不能送簽；按「完成並通知業務」才會寫回業務的報價。</span></div>
  </details>`;
}

/** 重算（成本明細或新增報價項目有變動時呼叫）：連動結果、5 張卡、報價品項區、頁底合計；並排程伺服器試算 */
function cfRecalc(s, lines) {
  if (s.closed || !cfCanSync(s) || !s.ed || s.ed.dead) return null;
  lines = lines || s.ed.getLines();
  const live = QCL.liveSummary({ items: s.q.items, lines, newItems: cfNewItemsClean(s), discountType: s.q.discountType, discountValue: s.q.discountValue });
  s.live = live;
  if (s.ed.revenue !== live.revenue) s.ed.setRevenue(live.revenue);   // 印花稅用連動後的營收即時重算
  s.ed.setLinkInfo(live.al);
  cfPaint(s, live);
  cfUpdateItems(s, live, lines);
  cfMarkChanged(s);
  if (s.cst && typeof CostSteps !== 'undefined') CostSteps.refresh(s);   // 逐步填寫：成本明細或新增報價項目變了，重算各步驟的完成狀態（合併成一次）
  return live;
}

function cfSetText(root, sel, v) { const el = root.querySelector(sel); if (el && el.textContent !== v) el.textContent = v; return el; }

/** 5 張卡＋頁底合計。srv（選填）＝伺服器試算的數字（與畫面不同時以伺服器為準） */
function cfPaint(s, live, srv) {
  const root = s.ov;
  const revC = srv ? srv.revenueCents : live.revenueCents;
  const costC = srv ? srv.costCents : live.costCents;
  const gpC = srv ? srv.gpCents : live.gpCents;
  const mt = srv ? srv.marginText : live.marginText;
  cfSetText(root, '#qapCfRev', cfMoney(revC / 100));
  cfSetText(root, '#qapCfCost', cfMoney(costC / 100));
  const gp = cfSetText(root, '#qapCfGp', cfMoney(gpC / 100));
  if (gp) gp.className = 'pnl-sum-value pnl-gp-val ' + (gpC >= 0 ? 'positive' : 'negative');
  const mg = cfSetText(root, '#qapCfMargin', mt === null || mt === undefined ? '—' : mt + '%');
  if (mg) { const p = parseFloat(mt); mg.className = 'pnl-sum-value pnl-margin-val ' + (mt === null || mt === undefined || !isFinite(p) ? '' : (p >= 30 ? 'good' : p >= 15 ? 'warn' : 'danger')); }
  QCL.paintOutsourcedCard(root.querySelector('#qapCfOs'), srv ? QCL.outsourcedCardModel(srv.outsourced) : live.card);
  cfSetText(root, '#qapCfTotal', fmtNum(costC / 100));
  const stampC = live.stampCents;
  cfSetText(root, '#qapCfNote', stampC > 0 ? '（含印花稅 ' + fmtNum(stampC / 100) + '）' : '');
  cfSetText(root, '#qapCfFootMargin', mt === null || mt === undefined ? '—' : mt + '%');
  cfSetText(root, '#qapCfFootOs', (srv ? srv.outsourced.outsourcedPctOfCost : live.outsourced.outsourcedPctOfCost) + '%');
}

/** 報價品項區的動態欄位：連動後的單位／數量（業務原值 → 新值）、單價、小計、對應的成本列數與成本小計、狀態晶片、各小計列 */
function cfUpdateItems(s, live, lines) {
  const tb = s.ov.querySelector('#qapCfItems');
  if (!tb) return;
  const q = s.q;
  const base = Array.isArray(q.items) ? q.items : [];
  const byKey = new Map();
  live.al.items.forEach((x) => byKey.set(x.lid || x.nid, x));
  // 每個報價品項被幾列成本對應、成本小計（連動＋拆項；舊資料沒有 rel 的列有 forLid 視為拆項；不對應的列不算）
  const cost = new Map();
  (lines || []).forEach((l) => {
    if (!l || l.auto === 'stamp') return;
    const ids = [];
    [l.forLid].concat(Array.isArray(l.forLids) ? l.forLids : []).forEach((k) => { if (typeof k === 'string' && k && ids.indexOf(k) < 0) ids.push(k); });
    const rel = l.rel || (ids.length ? 'split' : 'none');
    if (rel === 'none') return;
    const c = QCL.lineCents(l);
    ids.forEach((k) => { const o = cost.get(k) || { n: 0, cents: 0 }; o.n += 1; o.cents += c; cost.set(k, o); });
  });
  const rows = new Map();
  Array.prototype.forEach.call(tb.querySelectorAll('tr[data-ci]'), (tr) => rows.set(tr.getAttribute('data-ci'), tr));
  const effOf = (key) => {
    const x = byKey.get(key);
    if (!x) return null;
    return x.isNew ? live.items[base.length + (s.newItems || []).findIndex((n) => n.nid === key)] : live.items[x.index];
  };
  rows.forEach((tr, key) => {
    const x = byKey.get(key);
    if (!x) return;
    const src = x.isNew ? null : base[x.index];
    const eff = effOf(key) || {};
    const price = src ? (parseFloat(src.unitPrice) || 0) : 0;
    const needPrice = x.isNew || !!(src && src.needPrice === true && !(price > 0));
    const shownQty = x.conflict || x.zero ? x.qtyFrom : x.qty;
    const shownUnit = x.conflict || x.zero ? x.unitFrom : x.unit;
    const qCell = tr.querySelector('.qci-q'), uCell = tr.querySelector('.qci-u');
    if (!x.isNew) {
      if (qCell) qCell.innerHTML = x.qtyChanged && !x.conflict && !x.zero ? `<span class="qci-old">${e(fmtNum(Number(x.qtyFrom)))}</span>→ <b>${e(fmtNum(Number(shownQty)))}</b>` : e(fmtNum(Number(shownQty)));
      if (uCell) uCell.innerHTML = x.unitChanged && !x.conflict && !x.zero ? `<span class="qci-old">${e(String(x.unitFrom || ''))}</span>→ <b>${e(String(shownUnit || ''))}</b>` : e(String(shownUnit || ''));
    }
    const pCell = tr.querySelector('.qci-p');
    if (pCell) pCell.innerHTML = needPrice ? '<span class="qci-chip need">由業務補填</span>' : e(fmtNum(price));
    cfSetText(tr, '.qci-s', fmtNum(QCL.itemCents(eff) / 100));
    const co = cost.get(key);
    cfSetText(tr, '.qci-n', co ? co.n + ' 列' + (x.linkCount ? '（連動 ' + x.linkCount + '）' : '') : '—');
    cfSetText(tr, '.qci-c', co ? fmtNum(co.cents / 100) : '—');
    const flags = tr.querySelector('.qci-flags');
    if (flags) {
      let f = '';
      if (x.isNew) f += '<span class="qci-chip new">＋新增</span><span class="qci-chip need">待業務補單價</span>';
      else if (needPrice) f += '<span class="qci-chip need">待業務補單價</span>';
      if (x.conflict) f += '<span class="qci-chip bad">單位衝突</span>';
      else if (x.zero) f += '<span class="qci-chip bad">連動後數量為 0</span>';
      else if (!x.isNew && x.changed) f += '<span class="qci-chip chg">已連動</span>';
      else if (x.isNew && x.linkCount > 0) f += '<span class="qci-chip chg">已連動</span>';
      if (flags.innerHTML !== f) flags.innerHTML = f;
    }
    if (x.isNew) {
      // 顧問新增的報價項目被「連動」的成本列命中：數量／單位由那些列決定（applyLinks 以連動結果為準），輸入框改顯示連動結果並唯讀；解除連動就還原顧問自己輸入的值
      const n = (s.newItems || []).find((z) => z.nid === key);
      const linked = x.linkCount > 0 && !x.conflict && !x.zero;
      const nq = tr.querySelector('.qci-nq'), nu = tr.querySelector('.qci-nu');
      if (n && nq && nu) {
        const wantQ = linked ? String(x.qty) : String(n.qtyRaw !== undefined ? n.qtyRaw : n.qty);
        const wantU = linked ? String(x.unit) : String(n.unit);
        if (nq.value !== wantQ) nq.value = wantQ;
        if (nu.value !== wantU) nu.value = wantU;
        [nq, nu].forEach((inp) => {
          inp.readOnly = linked;
          inp.classList.toggle('qci-linked', linked);
          if (linked) inp.title = '數量與單位由連動的成本列決定；要自己輸入請先把那些列改成「拆項」或「不對應」'; else inp.removeAttribute('title');
        });
      }
    }
    tr.classList.toggle('qci-hit', !x.isNew && !!x.changed && !x.conflict && !x.zero);
    tr.classList.toggle('qci-bad', !!(x.conflict || x.zero));
  });
  // 各小計列：從最近的分組標題或小計列之後，累加連動後的品項金額
  let acc = 0;
  base.forEach((it, i) => {
    if (!it) return;
    if (it.kind === 'title') { acc = 0; return; }
    if (it.kind === 'subtotal') {
      const tr = tb.querySelector('tr[data-ci-sub="' + i + '"]');
      if (tr) cfSetText(tr, '.qci-seg', fmtNum(acc / 100));
      acc = 0; return;
    }
    acc += QCL.itemCents(live.items[i]);
  });
}

// ── 伺服器試算（簽核層級預覽）：debounce、同時只留一個進行中的請求、過期回應丟棄 ──
/** 送給 cost-draft/* 端點的 body（風險預留沒選就不帶） */
function cfDraftBody(s, lines) {
  const body = { costLines: lines, newItems: cfNewItemsClean(s).map((n) => ({ nid: n.nid, desc: n.desc, unit: n.unit, qty: n.qty })) };
  const risk = costRiskValue(s);
  if (risk !== '') body.contingencyPct = Number(risk);
  return body;
}

/** 新增報價項目的欄位檢查：品名必填、數量需是 0 以上的數字；mark＝順便把有問題的欄位標紅。回傳錯誤訊息或 '' */
function cfNewItemsProblem(s, mark) {
  let msg = '';
  (s.newItems || []).forEach((n) => {
    n.descBad = !String(n.desc || '').trim();
    n.qtyBad = !!n.qtyBad;
    if (mark) {
      const tr = s.ov.querySelector('tr[data-nid="' + n.nid.replace(/"/g, '') + '"]');
      if (tr) {
        tr.querySelector('.qci-nd').classList.toggle('qci-bad', n.descBad);
        tr.querySelector('.qci-nq').classList.toggle('qci-bad', n.qtyBad);
      }
    }
    if (!msg && n.descBad) msg = '新增的報價項目有一列沒有填品名，請填寫或移除該列';
    if (!msg && n.qtyBad) msg = '新增的報價項目數量不是有效的數字（需為 0 以上）';
  });
  return msg;
}

/** 目前畫面能不能送去試算（成本明細與新增項目都填得完整）；回傳 { ok, c } */
function cfPeek(s) {
  const c = s.ed && !s.ed.dead && typeof s.ed.peek === 'function' ? s.ed.peek() : null;
  if (!c) return { ok: false, c: null };
  const ok = c.invalid === 0 && c.blankDesc === 0 && c.badTarget === 0 && !cfNewItemsProblem(s, false);
  return { ok, c };
}

function cfMarkChanged(s) {
  s.sumSeq = (s.sumSeq || 0) + 1;
  s.sumState = 'pending';
  cfPaintPv(s);
  cfScheduleRun(s, CF_SUM_DEBOUNCE);
}

function cfScheduleRun(s, delay) {
  clearTimeout(s.sumTimer);
  if (s.closed) return;
  s.sumTimer = setTimeout(() => { cfRunSummary(s); }, Math.max(0, delay));
}

async function cfRunSummary(s) {
  if (s.closed || !cfCanSync(s)) return;
  if (s.sumBusy) return;   // 同時只留一個進行中的請求：完成時若資料又變了會再排一次
  const gap = (s.sumLastAt || 0) + CF_SUM_DEBOUNCE - Date.now();
  if (gap > 0) { cfScheduleRun(s, gap); return; }
  const pk = cfPeek(s);
  if (!pk.ok) { s.sumState = 'blocked'; cfPaintPv(s); return; }
  s.sumBusy = true;
  s.sumLastAt = Date.now();
  const seq = s.sumSeq;
  const res = await apiCall('POST', '/quotations/' + encodeURIComponent(s.id) + '/cost-draft/summary', cfDraftBody(s, pk.c.lines));
  s.sumBusy = false;
  if (s.closed) return;
  if (seq !== s.sumSeq) { cfScheduleRun(s, CF_SUM_DEBOUNCE); return; }   // 這份回應已過期（期間又改了）：丟棄，重新試算
  if (res.ok && res.data && res.data.items) {
    s.sum = res.data;
    s.sumState = 'ok';
    s.sumErr = '';
    cfApplySummary(s);
  } else {
    s.sumState = 'err';
    s.sumErr = res.status === 429 ? '請求太頻繁，稍後會自動重試' : errMsg(res);
    cfPaintPv(s);
    if (res.status === 429) cfScheduleRun(s, 8000);
  }
}

/** 伺服器試算回來：對照畫面數字（應逐分相同）；不同就以伺服器為準並標示；更新簽核層級預覽 */
function cfApplySummary(s) {
  const sum = s.sum, live = s.live;
  if (!sum || !live) { cfPaintPv(s); return; }
  const os = { outsourcedCents: sum.outsourcedCents, outsourcedPctOfCost: sum.outsourcedPctOfCost, outsourcedPctOfConsult: sum.outsourcedPctOfConsult };
  // 伺服器算不出營收／毛利（financeError：例如折扣後營收為 0）時沒有可對照的數字，不比較也不覆蓋畫面數字
  if (sum.financeError) { s.ov.setAttribute('data-cf-match', '1'); cfPaintPv(s); return; }
  const same = sum.revenueCents === live.revenueCents && sum.costCents === live.costCents && sum.gpCents === live.gpCents && sum.marginText === live.marginText
    && sum.outsourcedCents === live.outsourced.outsourcedCents && sum.outsourcedPctOfCost === live.outsourced.outsourcedPctOfCost && sum.outsourcedPctOfConsult === live.outsourced.outsourcedPctOfConsult;
  s.ov.setAttribute('data-cf-match', same ? '1' : '0');
  if (!same) cfPaint(s, live, { revenueCents: sum.revenueCents, costCents: sum.costCents, gpCents: sum.gpCents, marginText: sum.marginText, outsourced: os });
  cfPaintPv(s);
}

/** 簽核層級預覽區：類別、需要簽核的關卡、伺服器試算的毛利率、警告；不能送簽的原因只當參考（顧問不送簽） */
function cfPaintPv(s) {
  const box = s.ov && s.ov.querySelector('#qapCfPv');
  if (!box) return;
  const st = s.sumState || 'pending';
  box.setAttribute('data-state', st);
  let body;
  if (st === 'blocked') {
    body = '<span class="qap-muted">成本明細或新增的報價項目還沒填完整（項目名稱、對應的報價品項、品名…），填完後會自動試算簽核層級。</span>';
  } else if (st === 'err' && !s.sum) {
    body = '<span class="qap-muted">簽核層級試算暫時無法取得（' + e(s.sumErr || '請稍後再試') + '）；上方數字仍是畫面即時試算。</span>';
  } else if (!s.sum) {
    body = '<span class="qap-muted">簽核層級試算中…</span>';
  } else {
    const pv = s.sum.preview || {};
    const names = tierNames({ tiers: pv.tiers });
    let h = '<div class="qap-cf-pvrow"><span><span class="k">類別</span><b>' + e(pv.rowLabel || pv.rowKey || '—') + '</b></span>';
    h += '<span><span class="k">簽核層級</span>' + (names.length ? names.map((n) => `<span class="qap-chip sel">${e(n)}</span>`).join('<span class="qap-arrow">&#10140;</span>') : '<span class="qap-muted">成本尚未填妥，暫時無法判定</span>') + '</span>';
    h += '<span><span class="k">毛利率（伺服器試算）</span><b>' + (pv.marginText !== null && pv.marginText !== undefined ? e(pv.marginText) + '%' : '—') + '</b></span></div>';
    const warns = [].concat(pv.warnings || []);
    if (warns.length) h += '<div class="qap-alert warn">' + warns.map((w) => e(w)).join('<br>') + '</div>';
    const bl = (Array.isArray(pv.blockers) ? pv.blockers : []).map((b) => (b && (b.message || b.code)) || '').filter(Boolean);
    if (bl.length) h += '<div class="qap-alert info"><b>業務送簽前還需處理（供你參考）：</b>' + bl.map((w) => e(w)).join('；') + '</div>';
    if (st === 'err') h += '<div class="qap-cf-mis">最新一次試算失敗（' + e(s.sumErr || '') + '），以上是上一次的結果。</div>';
    if (s.ov.getAttribute('data-cf-match') === '0') h += '<div class="qap-cf-mis">畫面即時試算與伺服器試算不同，已改顯示伺服器的數字。</div>';
    body = h;
  }
  box.innerHTML = '<div class="qap-cf-pvbody">' + body + '</div>';
}

// ── 預覽 ──
/**
 * 把伺服器 summary 回傳的 items 套回載入的報價單（保留分組標題／小計列與順序、單價、折扣）：連動後的數量／單位覆蓋、顧問新增的項目接在最後（單價 0、待補）。
 * 預覽走既有的 buildQuotePreviewHtml，所以這裡只組出報價單物件。
 */
function cfDraftQuote(q, sumItems) {
  const base = Array.isArray(q.items) ? q.items : [];
  const byIdx = new Map();
  const fresh = [];
  (sumItems || []).forEach((x) => { if (x && x.isNew) fresh.push(x); else if (x && typeof x.index === 'number') byIdx.set(x.index, x); });
  const items = base.map((it, i) => {
    const x = byIdx.get(i);
    return x && it && it.kind !== 'title' && it.kind !== 'subtotal' ? Object.assign({}, it, { qty: x.qty, unit: x.unit }) : it;
  });
  fresh.forEach((x) => items.push({ lid: x.nid, desc: x.desc, unit: x.unit, qty: x.qty, unitPrice: 0, needPrice: true }));
  return Object.assign({}, q, { items });
}

function cfSetPvBusy(s, on) {
  s.pvBusy = !!on;
  ['#qapCfPvQuote', '#qapCfPvPnl'].forEach((sel) => { const b = s.ov.querySelector(sel); if (b) b.disabled = !!on; });
}

/** 預覽前的檢查（同「儲存」）：通過回傳該次的成本列，否則提示並回 null */
function cfPreviewReady(s) {
  if (s.busy || s.pvBusy || !s.ed || s.ed.dead) return null;
  const c = QCL.collect(s.ov.querySelector('#qapCfMount'));
  const focusMark = (sel) => { const b = s.ov.querySelector(sel); if (b) { b.scrollIntoView({ block: 'center' }); b.focus(); } };
  if (c.invalid > 0) { toast('成本明細有 ' + c.invalid + ' 個欄位不是有效的數字（需為 0 以上），請修正標紅的欄位'); focusMark('#qapCfMount .qcl-bad'); return null; }
  if (c.blankDesc > 0) { toast('有 ' + c.blankDesc + ' 列沒有填「項目」，請填寫或刪除該列'); focusMark('#qapCfMount .qcl-bad'); return null; }
  if (c.badTarget > 0) { toast('有 ' + c.badTarget + ' 列的「對應」沒有指定有效的報價品項，請選擇報價品項或改成「不對應」'); focusMark('#qapCfMount .qcl-target.qcl-bad'); return null; }
  const np = cfNewItemsProblem(s, true);
  if (np) { toast(np); focusMark('#qapCfItems .qci-bad'); return null; }
  return c;
}

async function cfPreviewQuote(s) {
  const c = cfPreviewReady(s);
  if (!c) return;
  cfSetPvBusy(s, true);
  try {
    const res = await apiCall('POST', '/quotations/' + encodeURIComponent(s.id) + '/cost-draft/summary', cfDraftBody(s, c.lines));
    if (s.closed) return;
    if (!res.ok || !res.data || !res.data.items) { toast(costErrMsg(res), 6000); return; }
    let info = null;
    try { info = typeof _qpvLoadInfo === 'function' ? await _qpvLoadInfo(s.id) : null; } catch (err) { info = null; }
    if (s.closed) return;
    if (typeof showQuoteDraftPreview !== 'function') { toast('預覽模組尚未載入，請重新整理頁面後再試'); return; }
    showQuoteDraftPreview(cfDraftQuote(s.q, res.data.items), info);
  } finally { if (!s.closed) cfSetPvBusy(s, false); }
}

async function cfPreviewPnl(s) {
  const c = cfPreviewReady(s);
  if (!c) return;
  if (typeof previewQuotePnlDraft !== 'function') { toast('預覽模組尚未載入，請重新整理頁面後再試'); return; }
  cfSetPvBusy(s, true);
  try { await previewQuotePnlDraft(s.id, cfDraftBody(s, c.lines)); } finally { if (!s.closed) cfSetPvBusy(s, false); }
}

// ── 新增報價項目 ──
function cfAddNewItem(s) {
  if (!cfCanSync(s) || s.busy) return;
  if ((s.newItems || []).length >= QCL.MAX_NEW_ITEMS) { toast('顧問新增的報價項目最多 ' + QCL.MAX_NEW_ITEMS + ' 筆'); return; }
  if ((s.q.items || []).length + (s.newItems || []).length >= 50) { toast('報價單最多 50 列（含分組標題與小計列），無法再新增報價項目'); return; }
  s.newItems = (s.newItems || []).concat([{ nid: cfNewNid(), desc: '', unit: '式', qty: 1 }]);
  s.dirty = true;
  const tb = s.ov.querySelector('#qapCfItems');
  if (tb) tb.innerHTML = cfItemRowsHtml(s);
  s.ed.setItems(cfEditorItems(s));
  cfRecalc(s);
  const rows = s.ov.querySelectorAll('#qapCfItems tr.qci-new .qci-nd');
  if (rows.length) { rows[rows.length - 1].focus(); if (rows[rows.length - 1].scrollIntoView) rows[rows.length - 1].scrollIntoView({ block: 'nearest' }); }
}

async function cfRemoveNewItem(s, nid) {
  if (s.busy) return;
  const n = (s.newItems || []).find((x) => x.nid === nid);
  if (!n) return;
  const used = (s.ed && !s.ed.dead ? s.ed.getLines() : []).filter((l) => l && (l.forLid === nid || (Array.isArray(l.forLids) && l.forLids.indexOf(nid) >= 0))).length;
  if (used > 0) {
    const ok = await qapConfirm({ title: '移除新增的報價項目', message: '有 ' + used + ' 列成本明細對應這個新增的報價項目，移除後這些列需要重新選擇對應（或改成「不對應」）。確定移除？', okText: '移除', danger: true });
    if (!ok || s.closed) return;
  }
  s.newItems = s.newItems.filter((x) => x.nid !== nid);
  s.dirty = true;
  const tb = s.ov.querySelector('#qapCfItems');
  if (tb) tb.innerHTML = cfItemRowsHtml(s);
  s.ed.setItems(cfEditorItems(s));
  cfRecalc(s);
}

/** 新增的報價項目的輸入（品名／單位／數量）：只更新狀態與計算，不重畫列（輸入框不失焦） */
function cfOnNewItemInput(s, inp) {
  const tr = inp.closest('tr[data-nid]');
  if (!tr || s.busy) return;
  const n = (s.newItems || []).find((x) => x.nid === tr.getAttribute('data-nid'));
  if (!n) return;
  if (inp.classList.contains('qci-nd')) { n.desc = inp.value; n.descBad = false; inp.classList.remove('qci-bad'); }
  else if (inp.classList.contains('qci-nu')) n.unit = inp.value;
  else if (inp.classList.contains('qci-nq')) {
    const v = inp.value.trim();
    const num = v === '' ? 0 : Number(v);
    n.qtyRaw = inp.value;
    n.qtyBad = (inp.validity && inp.validity.badInput) || !(isFinite(num) && num >= 0 && num <= 1e9);
    n.qty = n.qtyBad ? 0 : num;
    inp.classList.toggle('qci-bad', !!n.qtyBad);
  } else return;
  s.dirty = true;
  s.ed.setItems(cfEditorItems(s));
  cfRecalc(s);
}

function renderCostFill(s) {
  if (s.closed) return;
  const q = s.q;
  if (!q) {
    s.body.innerHTML = '<div class="qap-muted" style="padding:30px;text-align:center">載入中…</div>';
    s.foot.innerHTML = '<button class="btn btn-secondary" type="button" data-act="close">關閉</button>';
    return;
  }
  const keep = s.body.scrollTop;
  const perm = q.perm || {};
  const can = !!perm.canEditCost;
  if (!can) { s.dirty = false; s.draftLines = null; }   // 已不能編輯（被鎖定或沒權限）：沒有東西可存，關閉不必再確認「尚未儲存」
  stashCostDraft(s);
  clearTimeout(s.sumTimer);   // 重畫後會重新排程伺服器試算
  const sync = cfCanSync(s);   // 完整的連動畫面（即時試算、報價品項區、對應、預覽）
  const dis = s.busy ? ' disabled' : '';
  const ap = q.approval;
  const hasLines = Array.isArray(q.costLines);
  const qcl = typeof QCL !== 'undefined';
  const useEditor = qcl && (can || hasLines);   // 可編輯：一律走新編輯器（舊式單第一次儲存即轉為新式）；不可編輯：有成本明細就唯讀顯示
  let notice = '';
  if (!can) {
    notice = (ap && (ap.state === 'pending' || ap.state === 'approved'))
      ? '<div class="qap-alert warn">此報價單已送簽或核准，成本已鎖定，目前只能檢視。</div>'
      : '<div class="qap-alert warn">你目前沒有填寫此單成本的權限（若這張單已送簽或核准，成本會被鎖定），只能檢視。</div>';
  } else if (q.costFlow && q.costFlow.state === 'filled') {
    notice = '<div class="qap-alert info">你已完成過這張單的成本；如需修改，改完請再按「完成並通知業務」。只按「儲存」會讓這張單回到「未完成」，業務就無法送簽。</div>';
  }
  // 風險預留：負責填成本的顧問主管依專案風險選 0/5/10/15/20（%）；尚未設定過顯示「請選擇…」，按「完成並通知業務」前必須選
  const riskNow = costRiskValue(s);
  const riskOpts = [['', '請選擇…']].concat(QAP_CONTINGENCY_PCTS.map((n) => [String(n), n + '%']))
    .map((o) => `<option value="${o[0]}"${o[0] === riskNow ? ' selected' : ''}>${o[1]}</option>`).join('');
  const note = q.costFlow && q.costFlow.note
    ? `<div class="qap-sec"><h3>業務備註</h3><div style="white-space:pre-wrap;word-break:break-word;font-size:13px">${e(q.costFlow.note)}</div></div>` : '';
  let editorBlock;
  if (!qcl) {
    editorBlock = '<div class="qap-alert bad">成本明細編輯器沒有載入成功，請重新整理頁面（Ctrl+F5）後再試。</div>';
  } else if (useEditor) {
    const hint = !can ? '以下為這張單的成本明細，目前只能檢視。'
      : (sync
        ? '請填寫專案實際成本；每列的「對應」決定它和業務報價的關係。<details class="qap-cf-help"><summary>怎麼填？</summary><div>「連動報價數量」＝你改這列的單位／數量，按「完成」時會同步改業務報價品項；「拆項（報價不動）」＝業務維持一式、你拆很多細項；「不對應（純成本）」＝差旅、交際費等。自家顧問填「顧問姓名」，委外請在「委外廠商」填廠商名稱。贈品或不計價的列成本可填 0。</div></details>'
        : '請填寫專案實際成本，輸入成本單價會自動換算小計。<details class="qap-cf-help"><summary>怎麼填？</summary><div>可以新增／刪除列、改單位與數量；自家顧問填「顧問姓名」，委外的顧問請填「委外廠商」；差旅交通、交際費與印花稅也計入專案成本。贈品或不計價的列成本可填 0。</div></details>');
    editorBlock = `<div class="qap-cf-head"><h3>成本明細</h3><div class="qap-muted">${hint}</div></div>
      <fieldset class="qap-fs" id="qapCfFs"${s.busy ? ' disabled' : ''}><div id="qapCfMount"></div></fieldset>`;
  } else {
    editorBlock = `<div class="qap-sec"><h3>成本（舊格式）</h3>${buildCostLegacyTable(q)}</div>`;
  }
  if (s.stale) notice = '<div class="qap-alert bad" id="qapCfStale" role="alert">' + e(s.stale) + '</div>' + notice;
  s.body.innerHTML = notice + `
    <div class="qap-sec"><h3>案件</h3><div class="qap-grid">
      <div><span class="k">報價單號</span><b>${e(q.quoteNo || '')}</b></div>
      <div><span class="k">業務</span>${e(q.ownerName || q.owner || '')}</div>
      <div><span class="k">公司</span>${e(q.company || '')}</div>
      <div><span class="k">專案名稱</span>${e(q.projectName || '')}</div>
      ${q.validUntil ? `<div><span class="k">報價期限</span>${e(q.validUntil)}</div>` : ''}
    </div></div>
    ${note}
    ${sync ? cfLiveHtml() : ''}
    ${sync ? cfItemsSectionHtml(s) : buildCostRefItems(q, s.refOpen !== false)}
    ${editorBlock}
    <div class="qap-sec" id="qapCfRisk"><h3>風險預留（Contingency）</h3>
      <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
        <select class="qap-input" data-risk="1" style="width:auto;min-width:110px"${can && !s.busy ? '' : ' disabled'}>${riskOpts}</select>
        <span class="qap-muted">依這個專案的風險預估；毛利分析（內部）會把成本明細「顧問服務成本」區的小計 × 此比例另計為成本（風險預留不計入簽核用的毛利）。沒有風險請選 0%。</span>
      </div>
    </div>`;
  let f = `<button class="btn btn-secondary" type="button" data-act="close"${dis}>關閉</button><span class="qap-cf-tot">成本合計 <b id="qapCfTotal"></b><span class="qap-cf-note" id="qapCfNote"></span></span>`;
  if (sync) f += '<span class="qap-cf-extra">毛利率 <b id="qapCfFootMargin">—</b>　委外佔比 <b id="qapCfFootOs">0.00%</b></span>';
  f += '<span class="sp"></span>';
  if (can && qcl) {
    f += `<button class="btn btn-secondary" type="button" data-act="save"${dis}>儲存</button>`;
    f += `<button class="btn btn-primary" type="button" data-act="done"${dis}>完成並通知業務</button>`;
  }
  s.foot.innerHTML = f;
  if (useEditor) {
    let lines = can && Array.isArray(s.draftLines) ? s.draftLines : (hasLines ? q.costLines : null);
    if (!can) s.draftLines = null;
    if (!lines) lines = QCL.seedFromItems(sync ? cfEditorItems(s) : q.items, { classCodes: costClassCodes(q), link: sync });   // 舊式單：舊 items[].cost 帶入成為成本單價
    s.ed = QCL.mount(s.ov.querySelector('#qapCfMount'), {
      lines, items: sync ? cfEditorItems(s) : q.items, mode: can ? (sync ? 'consultant' : 'edit') : 'view', classCodes: costClassCodes(q),
      consultantNames: cfConsultantNames(s),
      confirm: (msg) => qapConfirm({ title: '請確認', message: msg, okText: '確定' }),
      onChange: (ls, t) => { s.dirty = true; if (sync) cfRecalc(s, ls); else showCostTotal(s, t); },
    });
    if (sync) cfRecalc(s);
    else showCostTotal(s, QCL.totals(lines));
    // 逐步填寫：只有「可編輯＋看得到價格的完整連動畫面」才分步（舊格式、唯讀、沒有編輯器維持原樣）；測試開關 window.__qNoSteps 見 quote-coststeps.js
    if (sync && can && typeof CostSteps !== 'undefined' && CostSteps.allowed(s)) CostSteps.attach(s, CF_STEPS_API);
  } else {
    const tot = s.ov.querySelector('.qap-cf-tot');
    if (tot) tot.style.display = 'none';
  }
  s.body.scrollTop = keep;
}

/** 交給逐步填寫（quote-coststeps.js）的函式：這個檔案整個包在 IIFE 內，外面看不到這些函式，所以 attach 時當參數交出去 */
const CF_STEPS_API = {
  riskValue: (s) => costRiskValue(s),
  newItemsProblem: (s, mark) => cfNewItemsProblem(s, mark),
  doneMessage: (s, c, live, sync) => cfDoneMessage(s, c, live, sync),
  rerender: (s) => renderCostFill(s),
};

/** 風險預留可選值（%）；與 lib/quoteRoutes.js 的 CONTINGENCY_PCTS 一致 */
const QAP_CONTINGENCY_PCTS = [0, 5, 10, 15, 20];
/** 目前畫面上的風險預留值（字串，'' ＝尚未選）：使用者改過就用草稿，否則用伺服器上已存的值 */
function costRiskValue(s) {
  if (Object.prototype.hasOwnProperty.call(s, 'risk')) return s.risk;
  const v = s.q && s.q.contingencyPct;
  return typeof v === 'number' ? String(v) : '';
}

/** 頁底合計（含印花稅）。t 是 QCL.totals() 的結果（元）。完整連動畫面的頁底由 cfPaint 以伺服器同規則的分為單位數字填 */
function showCostTotal(s, t) {
  const out = s.ov.querySelector('#qapCfTotal');
  if (!out || !t) return;
  out.textContent = fmtNum(t.total);
  const note = s.ov.querySelector('#qapCfNote');
  if (note) note.textContent = t.stampUnknown ? '（不含印花稅：系統依合約金額自動計算）' : (t.stamp > 0 ? '（含印花稅 ' + fmtNum(t.stamp) + '）' : '');
}

/** 送出期間鎖住按鈕、風險預留與整個編輯器（fieldset disabled）、新增報價項目與預覽按鈕，不重畫、輸入內容原封不動 */
function setCostBusy(s) {
  const dis = !!s.busy;
  const can = !!(s.q && s.q.perm && s.q.perm.canEditCost);
  s.ov.querySelectorAll('.qap-footer button[data-act]').forEach((b) => { b.disabled = dis; });
  const fs = s.ov.querySelector('#qapCfFs');
  if (fs) fs.disabled = dis;
  s.ov.querySelectorAll('#qapCfLive button, #qapCfItems input, #qapCfItems button, #qapCfAddItem').forEach((b) => { b.disabled = dis; });
  const sel = s.ov.querySelector('select[data-risk]');
  if (sel) sel.disabled = dis || !can;
  if (s.cst && typeof CostSteps !== 'undefined') CostSteps.syncFooter(s);   // 逐步填寫：「完成並通知業務」在步驟沒走完前維持停用（上面的迴圈會把它一起放開）
}

/**
 * 過期分頁保護：載入（或成功儲存後重新載入）時記下伺服器給的 itemsSig（品項結構摘要）、costLinesSig（成本明細內容摘要），
 * 存檔時帶回；伺服器發現不是最新的就回 409 STALE_ITEMS／STALE_COSTS。伺服器沒給（沒有成本權限）就不帶。
 */
function costSigsOf(q) {
  const o = {};
  if (q && typeof q.itemsSig === 'string') o.itemsSig = q.itemsSig;
  if (q && typeof q.costLinesSig === 'string') o.costLinesSig = q.costLinesSig;
  return o;
}

const COST_STALE_CODES = ['STALE_ITEMS', 'STALE_COSTS'];

/** 過期提示條（不重畫編輯器，輸入內容原封不動）：已存在就更新文字，沒有就加在最上面 */
function showCostStale(s, msg) {
  s.stale = msg;
  let el = s.body.querySelector('#qapCfStale');
  if (!el) {
    s.body.insertAdjacentHTML('afterbegin', '<div class="qap-alert bad" id="qapCfStale" role="alert"></div>');
    el = s.body.querySelector('#qapCfStale');
  }
  el.textContent = msg;
  s.body.scrollTop = 0;
}

/** 儲存失敗的人話提示（伺服器的 code 對照） */
function costErrMsg(res) {
  const d = res.data || {};
  const m = errMsg(res);
  switch (d.code) {
    case 'STALE_ITEMS': return '報價品項已被業務修改，這次沒有儲存（你輸入的內容仍保留在畫面上）。請關閉後重新整理，確認成本明細後再送出。';
    case 'STALE_COSTS': return '成本明細已在其他視窗被更新，這次沒有儲存（你輸入的內容仍保留在畫面上）。請關閉後重新整理。';
    case 'CLIENT_OUTDATED': return '畫面版本過舊，請重新整理頁面（Ctrl+F5）後再填寫；這次輸入的內容沒有儲存。';
    case 'LOCKED_PENDING': return '這張報價單已送簽或核准，成本已鎖定，無法儲存。';
    case 'MISSING_COST': return '尚未填寫成本明細：請至少填一列成本（不含印花稅的總成本需大於 0）。';
    case 'BAD_COST_LINE': return '成本明細有誤：' + m;
    case 'LINK_UNIT_CONFLICT': return '連動的成本列單位不一致，沒有完成：' + m;
    case 'LINK_QTY_ZERO': return '連動後的報價品項數量需大於 0，沒有完成：' + m;
    case 'TOO_MANY_ITEMS': return m;
    default: return m;
  }
}

/** 伺服器存的顧問草稿新增的報價項目 → 畫面狀態 */
function cfDraftFromQuote(q) {
  const arr = q && q.costDraft && Array.isArray(q.costDraft.newItems) ? q.costDraft.newItems : [];
  return arr.map((n) => ({ nid: String(n.nid), desc: String(n.desc || ''), unit: String(n.unit || '式'), qty: Number(n.qty) || 0 }));
}

/**
 * 按「完成並通知業務」前的確認視窗文字。submitCostFill 與逐步畫面的「確認並完成」（quote-coststeps.js）共用同一份，內容完全相同。
 * c＝{ zeroCost（成本單價為 0 的列數）, lines（目前的成本列）}；live＝連動試算結果（沒有連動畫面時 null）；sync＝是否走完整連動畫面。
 */
function cfDoneMessage(s, c, live, sync) {
  let msg = '完成後會通知業務，你仍可在送簽前修改。';
  if (c.zeroCost > 0) msg = '有 ' + c.zeroCost + ' 列成本單價填的是 0（贈品或不計成本的列可以是 0；若不是，請先回去填寫）。\n\n' + msg;
  const syncTxt = live ? QCL.syncConfirmText(live.changes) : '';
  if (syncTxt) msg = syncTxt + '\n\n' + msg;
  // 連動寫回後報價合計會變多少：顧問當下就看到前後金額（連動選錯會讓金額成倍變動）；與議價折扣衝突時一併警告
  if (live && live.changes && live.changes.length && typeof QCL.revenueOf === 'function') {
    const before = QCL.revenueOf(s.q.items, s.q.discountType, s.q.discountValue), after = live.revenue;
    if (isFinite(before) && isFinite(after) && Math.round(before * 100) !== Math.round(after * 100)) {
      const ntd = (n) => 'NT$ ' + Math.round(n).toLocaleString('en-US');
      const pct = before > 0 ? '（' + (after >= before ? '+' : '') + ((after - before) / before * 100).toFixed(1) + '%）' : '';
      msg = '報價合計將由 ' + ntd(before) + ' 變為 ' + ntd(after) + pct + '。\n\n' + msg;
    }
    const dc = ((s.sum && s.sum.preview && s.sum.preview.blockers) || []).filter((b) => b && /DISCOUNT/.test(String(b.code || ''))).map((b) => b.message || b.code);
    if (dc.length) msg = '⚠ 連動後報價合計與議價折扣衝突（' + dc.join('；') + '），業務需先調整折扣才能送簽。\n\n' + msg;
  }
  // 成本明細與報價品項不必一一對應（整包、一式拆多列）：找不到對應列的品項只提醒、不擋（可能是已包含在其他列，也可能是漏填，例如業務事後新增品項）
  const un = typeof QCL.unmatchedItemsNote === 'function' ? QCL.unmatchedItemsNote(c.lines, sync ? cfEditorItems(s) : s.q.items, { classCodes: costClassCodes(s.q), link: sync }) : '';
  if (un) msg = un + '\n\n' + msg;
  return msg;
}

async function submitCostFill(s, done) {
  if (s.busy || !s.q || !(s.q.perm || {}).canEditCost || !s.ed || typeof QCL === 'undefined') return;
  // 逐步填寫（quote-coststeps.js）：步驟沒走完不能「完成並通知業務」（按鈕本來就停用，這裡是第二道防線）；「儲存」草稿不受影響
  if (done && s.cst && typeof CostSteps !== 'undefined' && !CostSteps.canFinish(s)) { toast('還有步驟沒完成，請先走完所有步驟再按「完成並通知業務」（可以先按「儲存」存成草稿）'); return; }
  const sync = cfCanSync(s);
  const mountEl = s.ov.querySelector('#qapCfMount');
  const c = QCL.collect(mountEl);
  const focusMark = (sel) => { const b = s.ov.querySelector(sel); if (b) { b.scrollIntoView({ block: 'center' }); b.focus(); } };
  if (c.invalid > 0) { toast('成本明細有 ' + c.invalid + ' 個欄位不是有效的數字（需為 0 以上），請修正標紅的欄位'); focusMark('#qapCfMount .qcl-bad'); return; }
  if (c.blankDesc > 0) { toast('有 ' + c.blankDesc + ' 列沒有填「項目」，請填寫或刪除該列'); focusMark('#qapCfMount .qcl-bad'); return; }
  if (sync) {
    if (c.badTarget > 0) { toast('有 ' + c.badTarget + ' 列的「對應」沒有指定有效的報價品項（連動／拆項要選一個報價品項；原品項若已被業務刪除，請重新選擇或改成「不對應」）', 7000); focusMark('#qapCfMount .qcl-target.qcl-bad'); return; }
    const np = cfNewItemsProblem(s, true);
    if (np) { toast(np); focusMark('#qapCfItems .qci-bad'); return; }
  }
  const risk = costRiskValue(s);
  if (done) {
    if (c.zeroQty > 0) { toast('有 ' + c.zeroQty + ' 列數量是 0，請填寫數量或刪除該列'); focusMark('.qcl-warn'); return; }
    const nonStamp = c.lines.filter((l) => l.auto !== 'stamp');
    if (!nonStamp.length || !(QCL.totals(c.lines).subtotalExStamp > 0)) { toast('尚未填寫成本明細：請至少填一列成本（不含印花稅的總成本需大於 0）'); return; }
    if (risk === '') { toast('請選擇風險預留（依專案風險預估；沒有風險請選 0%）'); return; }
    let live = null;
    if (sync) {
      // 連動寫回報價前先在畫面擋下伺服器也會擋的情況：單位衝突、連動後數量為 0、列數超過上限
      live = cfRecalc(s, c.lines) || s.live;
      const nm = (arr) => arr.slice(0, 5).map((x) => '「' + String(x.desc || '').slice(0, 12) + '」').join('、') + (arr.length > 5 ? '…另 ' + (arr.length - 5) + ' 項' : '');
      if (live.conflicts.length) { toast('連動的成本列單位不一致：' + nm(live.conflicts) + '。同一個報價品項的連動列請用相同單位（或把其中幾列改成「拆項」）', 8000); focusMark('#qapCfMount .qcl-conf'); return; }
      if (live.zero.length) { toast('連動後的報價品項數量需大於 0：' + nm(live.zero) + '。請填數量，或把連動的列改成「拆項」', 8000); focusMark('#qapCfMount .qcl-warn, #qapCfMount .qcl-linknote.bad'); return; }
      if ((s.q.items || []).length + (s.newItems || []).length > 50) { toast('新增報價項目後總列數會超過 50 列（含分組標題與小計列），請減少新增的項目'); return; }
    }
    const msg = cfDoneMessage(s, c, live, sync);
    const ok = await qapConfirm({ title: '完成並通知業務', message: msg, okText: '完成' });
    if (!ok || s.closed) return;
  }
  s.busy = true;
  setCostBusy(s);
  const body = { costLines: c.lines, done: !!done, contingencyPct: risk === '' ? null : Number(risk), costModel: 2 };
  if (sync) body.newItems = cfNewItemsClean(s).map((n) => ({ nid: n.nid, desc: n.desc, unit: n.unit, qty: n.qty }));
  Object.assign(body, s.sigs || {});   // itemsSig／costLinesSig：載入時的版本，伺服器發現已過期就擋下（409 STALE_*）
  const res = await apiCall('PUT', '/quotations/' + encodeURIComponent(s.id) + '/costs', body);
  if (s.closed) { if (res.ok) afterMutation(); return; }
  if (!res.ok && res.data && COST_STALE_CODES.indexOf(res.data.code) >= 0) {
    // 過期分頁：不重新載入、不重畫（輸入內容保留，簽章也維持舊值，再按一次儲存仍會被擋，使用者必須重新開啟）
    const m = costErrMsg(res);
    toast(m, 8000);
    showCostStale(s, m);
    s.busy = false;
    setCostBusy(s);
    return;
  }
  if (res.ok) {
    s.dirty = false;
    s.draftLines = null;
    delete s.risk;
    if (done) {
      toast('已完成，已通知業務');
      closeCostFill();
      afterMutation();
      return;
    }
    toast('已儲存');
    afterMutation();
  } else {
    toast(costErrMsg(res), 6000);
  }
  // 成敗都重載最新狀態（409 代表鎖定或版本已變）並重畫；使用者已輸入但未成功儲存的明細由 stashCostDraft 保留（s.draftLines）
  const r = await apiCall('GET', '/quotations/' + encodeURIComponent(s.id));
  if (s.closed) return;
  if (r.ok && r.data && r.data.id) {
    s.q = r.data;
    if (!s.dirty) { s.sigs = costSigsOf(r.data); s.newItems = cfDraftFromQuote(r.data); }   // 沒有保留中的草稿＝畫面會用最新資料重畫，簽章與草稿新增的報價項目一併更新；有草稿（儲存失敗）就維持載入時的簽章與畫面上的項目
  }
  s.busy = false;
  renderCostFill(s);
}

async function openQuoteCostFill(id) {
  ensureStyle();
  closeCostFill();
  const m = mountModal('quoteCostFillOverlay', '填寫成本', 'qap-wide', true);
  const s = { id, q: null, busy: true, dirty: false, closed: false, ov: m.ov, body: m.body, foot: m.foot, draftLines: null, ed: null, refOpen: true, newItems: [] };
  _cf = s;
  s.onKey = (ev) => {
    if (ev.key !== 'Escape') return;
    if (!s.ov.isConnected) { document.removeEventListener('keydown', s.onKey); return; }
    if (_confirmOpen) return;
    const pv = document.getElementById('quotePreviewOverlay');   // 預覽視窗開著時，Esc 先關預覽
    if (pv) { if (typeof closeQuotePreview === 'function') closeQuotePreview(); else pv.remove(); return; }
    requestCloseCostFill(s);
  };
  document.addEventListener('keydown', s.onKey);
  s.ov.addEventListener('click', (ev) => {
    const t = ev.target.closest('[data-act]');
    if (!t || !s.ov.contains(t) || t.disabled) return;
    if (t.closest('.qcl-root')) return;   // 編輯器自己的按鈕（新增／移動／刪除…）由 QCL 處理
    const act = t.dataset.act;
    if (act === 'close') requestCloseCostFill(s);
    else if (act === 'save') submitCostFill(s, false);
    else if (act === 'done') submitCostFill(s, true);
    else if (act === 'pvQuote') cfPreviewQuote(s);
    else if (act === 'pvPnl') cfPreviewPnl(s);
    else if (act === 'addNewItem') cfAddNewItem(s);
    else if (act === 'rmNewItem') cfRemoveNewItem(s, t.getAttribute('data-nid'));
  });
  s.ov.addEventListener('change', (ev) => {
    const t = ev.target;
    if (t && t.dataset && t.dataset.risk !== undefined) { s.risk = t.value; s.dirty = true; }
  });
  s.ov.addEventListener('input', (ev) => {
    const t = ev.target;
    if (t && t.classList && t.classList.contains('qci-in')) cfOnNewItemInput(s, t);   // 新增的報價項目（品名／單位／數量）
  });
  // 參考品項區收合狀態：重畫時維持使用者的選擇
  s.ov.addEventListener('toggle', (ev) => {
    const t = ev.target;
    if (t && t.classList && t.classList.contains('qap-ref')) s.refOpen = t.open;
  }, true);
  renderCostFill(s);
  const r = await apiCall('GET', '/quotations/' + encodeURIComponent(id));
  if (s.closed) return;
  if (!r.ok || !r.data || !r.data.id) { toast(errMsg(r)); closeCostFill(); return; }
  s.q = r.data;
  s.sigs = costSigsOf(r.data);
  s.newItems = cfDraftFromQuote(r.data);   // 顧問先前儲存草稿時新增的報價項目（還沒完成，報價 items 沒有它們）
  s.busy = false;
  renderCostFill(s);
}

// ═════════════════════════════════════════════════
// 簽核設定（名冊／商品歸類表／核決門檻與報價專用章）
// ═════════════════════════════════════════════════
let _st = null;

function closeSettings() {
  const s = _st;
  if (!s) return;
  _st = null;
  s.closed = true;
  document.removeEventListener('keydown', s.onKey);
  s.ov.remove();
}

async function requestCloseSettings(s) {
  if (s.closed) return;
  if (s.dirty) {
    const ok = await qapConfirm({ title: '尚未儲存', message: '名冊或商品歸類有尚未儲存的變更，確定要關閉嗎？', okText: '放棄並關閉', danger: true });
    if (!ok || s.closed) return;
  }
  closeSettings();
}

function clsLabel(s, k) {
  return (s.cfg && s.cfg.classLabels && s.cfg.classLabels[k]) || CLASS_LABEL_DEFAULT[k] || k;
}

function userText(s, username) {
  const u = (s.cfg.users || []).find((x) => x.username === username);
  if (!u) return username + '（帳號不存在）';
  return (u.displayName || u.username) + (u.displayName && u.displayName !== u.username ? '（' + u.username + '）' : '') + (u.active === false ? ' [停用]' : '');
}

function buildRosterTab(s) {
  const users = (s.cfg.users || []).filter((u) => u.active !== false).slice()
    .sort((a, b) => String(a.displayName || a.username).localeCompare(String(b.displayName || b.username), 'zh-Hant'));
  const secs = users.filter((u) => u.role === 'secretary');
  let html = `<div class="qap-alert info">所有「secretary」角色的在職帳號自動具有「董事會代核」與「報價章管理」權限，不需要在此重複加入。目前有：${secs.length ? secs.map((u) => e(u.displayName || u.username)).join('、') : '（無）'}</div>`;
  ROSTER_GROUPS.forEach((g) => {
    const sel = s.roster[g.key] || [];
    const chips = sel.map((un) => `<span class="qap-chip sel">${e(userText(s, un))}<button type="button" data-act="rmUser" data-g="${e(g.key)}" data-u="${e(un)}" title="移除" aria-label="移除">&#10005;</button></span>`).join('') || '<span class="qap-muted">（尚未指定）</span>';
    const opts = users.filter((u) => sel.indexOf(u.username) < 0)
      .map((u) => `<option value="${e(u.username)}">${e(u.displayName || u.username)}（${e(u.username)}）${u.role ? ' · ' + e(u.role) : ''}</option>`).join('');
    html += sec(e(g.label), `<div class="qap-muted" style="margin-bottom:6px">${e(g.desc)}</div>
      <div class="qap-path">${chips}</div>
      <select class="qap-input" data-act="addUser" data-g="${e(g.key)}" style="max-width:340px"><option value="">＋ 加入帳號…</option>${opts}</select>`);
  });
  return html;
}

function catalogRows(s) {
  const map = new Map();
  (s.cfg.catalog || []).forEach((c) => {
    if (!c || !c.name) return;
    if (!map.has(c.name)) map.set(c.name, []);
    const w = [c.bu, c.group].filter(Boolean).join(' / ');
    if (w && map.get(c.name).indexOf(w) < 0) map.get(c.name).push(w);
  });
  return Array.from(map.entries()).map(([name, where]) => ({ name, where: where.join('；') }));
}

function buildClassesTab(s) {
  s.catRows = catalogRows(s);
  return `<div class="qap-alert info">商品名稱＝商機使用的商品目錄名稱。未歸類的商品會被視為「其他」（主管 → 總經理兩關，且需要顧問填成本）。勾選「成本由業務填」代表此商品的成本由業務自己填（例如 SAP 軟體授權），不需要顧問。</div>
    <div class="qap-tools">
      <input type="text" class="qap-input" id="qapClsSearch" placeholder="搜尋商品名稱…" value="${e(s.clsSearch)}">
      <label style="display:flex;gap:5px;align-items:center;cursor:pointer"><input type="checkbox" data-act="onlyUncls"${s.onlyUncls ? ' checked' : ''}> 只看未歸類</label>
      <span class="qap-muted" id="qapClsCount"></span>
    </div>
    <div class="qap-tablewrap"><table class="qap-table"><thead><tr><th>商品名稱</th><th>BU / 群組</th><th>類別</th><th class="c">成本由業務填</th></tr></thead><tbody id="qapClsBody"></tbody></table></div>
    <div id="qapOrphans"></div>`;
}

function renderClassRows(s) {
  const body = s.ov.querySelector('#qapClsBody');
  if (!body) return;
  const kw = String(s.clsSearch || '').trim().toLowerCase();
  const rows = s.catRows.filter((r) => {
    if (kw && r.name.toLowerCase().indexOf(kw) < 0) return false;
    if (s.onlyUncls && s.classes[r.name]) return false;
    return true;
  });
  body.innerHTML = rows.map((r) => {
    const cur = s.classes[r.name];
    const opts = ['<option value="">（未歸類）</option>'].concat(CLASS_ORDER.map((k) => `<option value="${k}"${cur && cur.cls === k ? ' selected' : ''}>${e(clsLabel(s, k))}</option>`)).join('');
    return `<tr data-name="${e(r.name)}" class="${cur ? '' : 'qap-uncls'}"><td>${e(r.name)}${cur ? '' : ' <span class="qap-badge warn">未歸類</span>'}</td><td class="qap-muted">${e(r.where)}</td>
      <td><select class="qap-input" data-act="setCls" data-name="${e(r.name)}" style="min-width:150px">${opts}</select></td>
      <td class="c"><input type="checkbox" data-act="setCbs" data-name="${e(r.name)}"${cur && cur.costBySales ? ' checked' : ''}${cur ? '' : ' disabled'}></td></tr>`;
  }).join('') || '<tr><td colspan="4" class="qap-muted">（沒有符合的商品）</td></tr>';
  updateClassCount(s, rows.length);
  const orph = Object.keys(s.classes).filter((n) => !s.catRows.some((r) => r.name === n));
  const ob = s.ov.querySelector('#qapOrphans');
  if (ob) {
    ob.innerHTML = orph.length
      ? `<details style="margin-top:12px"><summary class="qap-muted" style="cursor:pointer">已不在商品目錄的歸類（${orph.length} 筆，儲存時會保留，可在此移除）</summary>
         <div class="qap-path" style="margin-top:6px">${orph.map((n) => `<span class="qap-chip">${e(n)}：${e(clsLabel(s, s.classes[n].cls))}<button type="button" data-act="rmOrphan" data-name="${e(n)}" title="移除" aria-label="移除">&#10005;</button></span>`).join('')}</div></details>`
      : '';
  }
}

function updateClassCount(s, shown) {
  const c = s.ov.querySelector('#qapClsCount');
  if (!c) return;
  const uncls = s.catRows.filter((r) => !s.classes[r.name]).length;
  c.innerHTML = `共 ${s.catRows.length} 項商品，顯示 ${shown} 項；` + (uncls ? `<b style="color:#d93025">${uncls} 項未歸類</b>` : '全部已歸類');
}

function buildRulesTab(s) {
  const cfg = s.cfg;
  const fmtPct = (v) => (v === null || v === undefined ? null : e(v) + '%');
  const rows = (cfg.rows || []).map((r) => {
    let c1, c2, c3;
    if (r.key === 'other') {
      c1 = '—'; c2 = '固定：一級主管 &#10140; 總經理（不看毛利）'; c3 = '—';
    } else {
      c1 = r.l1 === null || r.l1 === undefined ? '主管不可終局' : '毛利率 &ge; ' + fmtPct(r.l1);
      c2 = (r.l1 === null || r.l1 === undefined ? '' : fmtPct(r.l2) + ' &le; 毛利率 &lt; ' + fmtPct(r.l1)) || '毛利率 &ge; ' + fmtPct(r.l2);
      c3 = '毛利率 &lt; ' + fmtPct(r.l2);
    }
    return `<tr><td>${e(r.label || r.key)}</td><td>${c1}</td><td>${c2}</td><td>${c3}</td></tr>`;
  }).join('');
  const amt = cfg.amount || {};
  const am = (n) => (Number(n) || 0).toLocaleString('en-US');
  let html = sec('核決門檻表（唯讀）', `<div class="qap-muted" style="margin-bottom:6px">毛利率＝整張報價單「折扣後未稅」毛利率；壓線取較低層（剛好等於門檻算較低層級，例如顧問服務 25.00% 一級主管可簽）。規則版本：<b>${e(cfg.rulesVersion || '')}</b></div>
    <div class="qap-tablewrap"><table class="qap-table"><thead><tr><th>類別</th><th>一級主管即可</th><th>需總經理</th><th>需董事長</th></tr></thead><tbody>${rows}</tbody></table></div>
    <ul class="qap-list">
      <li>折扣後未稅金額 &gt; ${e(am(amt.chairman))} 元：至少需董事長（第 3 關）。</li>
      <li>折扣後未稅金額 &gt; ${e(am(amt.board))} 元：需董事會決議（一級主管 &#10140; 總經理 &#10140; 董事會，由管理部秘書代核准並登錄決議日期與文號）。</li>
      <li>簽核一律逐級、從一級主管開始，不跳級；金額與毛利各算一個層級取較高者。</li>
    </ul>`);
  const me = cfg.me || {};
  if (me.isAdmin || me.isSealManager) {
    html += sec('報價專用章', `<div class="qap-muted" style="margin-bottom:8px">已核准的報價單，PDF 與預覽右下「Prepared by」簽名欄會蓋上此章；Excel 一律不蓋章（可編輯檔，正式有章的報價單只有 PDF）。僅限 PNG／JPEG，檔案不得超過 300KB。</div>
      <div class="qap-seal-box"><div class="qap-seal-img" id="qapSealImg"></div>
      <div style="display:flex;flex-direction:column;gap:8px;align-items:flex-start">
        <input type="file" id="qapSealFile" accept="image/png,image/jpeg" style="display:none">
        <button class="btn btn-primary" type="button" data-act="pickSeal"${s.sealBusy ? ' disabled' : ''}>上傳新章</button>
        <button class="btn btn-secondary" type="button" data-act="delSeal"${s.sealBusy || !s.hasSeal ? ' disabled' : ''}>刪除目前的章</button>
        <span class="qap-muted" id="qapSealMsg"></span>
      </div></div>`);
  }
  return html;
}

function renderSealPreview(s) {
  const box = s.ov.querySelector('#qapSealImg');
  if (!box) return;
  box.innerHTML = s.hasSeal
    ? `<img src="${e(apiBase())}/quote-approval/seal?t=${e(s.sealTs)}" alt="報價專用章" onerror="this.style.display='none'">`
    : '<span class="qap-muted" style="background:#fff;padding:2px 6px;border-radius:4px">尚未上傳</span>';
  const msg = s.ov.querySelector('#qapSealMsg');
  if (msg) msg.textContent = s.hasSeal ? '目前已有報價專用章' : '目前沒有報價專用章';
}

function renderSettings(s) {
  if (s.closed) return;
  const me = (s.cfg && s.cfg.me) || {};
  const tabs = [];
  if (me.isAdmin) { tabs.push(['roster', '簽核名冊']); tabs.push(['classes', '商品歸類表']); }
  tabs.push(['rules', me.isAdmin ? '核決門檻與報價專用章' : '報價專用章']);
  if (!tabs.some((t) => t[0] === s.tab)) s.tab = tabs[0][0];
  s.tabs.innerHTML = '<div class="qap-tabs">' + tabs.map((t) => `<button class="qap-tab${s.tab === t[0] ? ' on' : ''}" type="button" data-act="tab" data-tab="${t[0]}">${e(t[1])}</button>`).join('') + '</div>';
  const keep = s.body.scrollTop;
  if (s.tab === 'roster') s.body.innerHTML = buildRosterTab(s);
  else if (s.tab === 'classes') { s.body.innerHTML = buildClassesTab(s); renderClassRows(s); }
  else { s.body.innerHTML = buildRulesTab(s); renderSealPreview(s); }
  s.body.scrollTop = keep;
  const dis = s.busy ? ' disabled' : '';
  let f = `<button class="btn btn-secondary" type="button" data-act="close"${dis}>關閉</button><span class="sp"></span>`;
  if (me.isAdmin) f += `<span class="qap-muted">${s.dirty ? '有未儲存的變更' : ''}</span><button class="btn btn-primary" type="button" data-act="save"${dis}>儲存名冊與歸類</button>`;
  s.foot.innerHTML = f;
}

function copyRoster(r) {
  const out = {};
  ROSTER_GROUPS.forEach((g) => { out[g.key] = Array.isArray(r && r[g.key]) ? r[g.key].slice() : []; });
  return out;
}

function copyClasses(pc) {
  const out = {};
  Object.keys(pc || {}).forEach((n) => { out[n] = { cls: pc[n].cls, costBySales: !!pc[n].costBySales }; });
  return out;
}

async function loadSettingsConfig(s) {
  const r = await apiCall('GET', '/quote-approval/config');
  if (s.closed) return false;
  if (!r.ok) { toast(errMsg(r)); return false; }
  s.cfg = r.data || {};
  s.roster = copyRoster(s.cfg.roster);
  s.classes = copyClasses(s.cfg.productClasses);
  s.hasSeal = !!s.cfg.hasSeal;
  s.sealTs = Date.now();
  return true;
}

async function saveSettings(s) {
  if (s.busy || !(s.cfg.me || {}).isAdmin) return;
  const warns = [];
  if (!s.roster.gm.length) warns.push('・總經理名冊是空的，送到總經理關的報價單會無法送簽');
  if (!s.roster.chairman.length) warns.push('・董事長名冊是空的，需要董事長的報價單會無法送簽');
  if (!s.roster.costProviders.length) warns.push('・成本填寫人名冊是空的，業務將無法選擇支援顧問');
  if (warns.length) {
    const ok = await qapConfirm({ title: '名冊不完整', message: warns.join('\n') + '\n\n仍要儲存嗎？', okText: '仍要儲存' });
    if (!ok || s.closed) return;
  }
  s.busy = true;
  renderSettings(s);
  const res = await apiCall('PUT', '/admin/quote-approval/config', { roster: s.roster, productClasses: s.classes });
  if (s.closed) return;
  if (res.ok) {
    s.dirty = false;
    toast('簽核設定已儲存');
    await loadSettingsConfig(s);   // 以伺服器實際儲存的內容為準
    if (s.closed) return;
  } else {
    toast(errMsg(res), 4500);
  }
  s.busy = false;
  renderSettings(s);
}

async function uploadSeal(s, file) {
  if (s.sealBusy) return;
  if (!file) return;
  if (file.type !== 'image/png' && file.type !== 'image/jpeg') { toast('只接受 PNG 或 JPEG 圖檔'); return; }
  if (file.size > 300 * 1024) { toast('章圖不得超過 300KB（目前 ' + Math.ceil(file.size / 1024) + 'KB）'); return; }
  s.sealBusy = true;
  renderSettings(s);
  let dataUrl = '';
  try {
    dataUrl = await new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => resolve(String(fr.result || ''));
      fr.onerror = () => reject(new Error('read'));
      fr.readAsDataURL(file);
    });
  } catch (err) {
    s.sealBusy = false;
    if (!s.closed) { toast('讀取圖檔失敗'); renderSettings(s); }
    return;
  }
  const res = await apiCall('PUT', '/quote-approval/seal', { dataUrl });
  if (s.closed) return;
  if (res.ok) { s.hasSeal = true; s.sealTs = Date.now(); toast('報價專用章已更新'); }
  else toast(errMsg(res), 4500);
  s.sealBusy = false;
  renderSettings(s);
}

async function deleteSeal(s) {
  if (s.sealBusy || !s.hasSeal) return;
  const ok = await qapConfirm({ title: '刪除報價專用章', message: '刪除後，已核准報價單的 PDF 與預覽將不再蓋章（需重新上傳）。\n\n確定要刪除嗎？', okText: '刪除', danger: true });
  if (!ok || s.closed) return;
  s.sealBusy = true;
  renderSettings(s);
  const res = await apiCall('DELETE', '/quote-approval/seal');
  if (s.closed) return;
  if (res.ok) { s.hasSeal = false; s.sealTs = Date.now(); toast('已刪除報價專用章'); }
  else toast(errMsg(res), 4500);
  s.sealBusy = false;
  renderSettings(s);
}

async function openQuoteApprovalSettings() {
  ensureStyle();
  closeSettings();
  const m = mountModal('quoteApprovalSettingsOverlay', '報價單簽核設定', '', true);
  const s = {
    closed: false, busy: true, sealBusy: false, dirty: false, tab: 'roster',
    ov: m.ov, tabs: m.tabs, body: m.body, foot: m.foot,
    cfg: null, roster: copyRoster(null), classes: {}, catRows: [],
    clsSearch: '', onlyUncls: false, hasSeal: false, sealTs: Date.now(),
  };
  _st = s;
  s.onKey = (ev) => {
    if (ev.key !== 'Escape') return;
    if (!s.ov.isConnected) { document.removeEventListener('keydown', s.onKey); return; }
    if (_confirmOpen) return;
    requestCloseSettings(s);
  };
  document.addEventListener('keydown', s.onKey);

  s.ov.addEventListener('click', (ev) => {
    const t = ev.target.closest('[data-act]');
    if (!t || !s.ov.contains(t) || t.disabled) return;
    const act = t.dataset.act;
    if (act === 'close') requestCloseSettings(s);
    else if (act === 'tab') { s.tab = t.dataset.tab; renderSettings(s); }
    else if (act === 'save') saveSettings(s);
    else if (act === 'rmUser') {
      if (s.busy) return;
      const arr = s.roster[t.dataset.g] || [];
      s.roster[t.dataset.g] = arr.filter((u) => u !== t.dataset.u);
      s.dirty = true;
      renderSettings(s);
    } else if (act === 'rmOrphan') {
      delete s.classes[t.dataset.name];
      s.dirty = true;
      renderClassRows(s);
      renderSettingsFooter(s);
    } else if (act === 'pickSeal') {
      const inp = s.ov.querySelector('#qapSealFile');
      if (inp) inp.click();
    } else if (act === 'delSeal') deleteSeal(s);
  });
  s.ov.addEventListener('change', (ev) => {
    const t = ev.target;
    if (!t) return;
    if (t.id === 'qapSealFile') {
      const f = t.files && t.files[0];
      t.value = '';
      uploadSeal(s, f);
      return;
    }
    const act = t.dataset && t.dataset.act;
    if (act === 'addUser') {
      const un = t.value;
      if (!un || s.busy) return;
      const arr = s.roster[t.dataset.g] || (s.roster[t.dataset.g] = []);
      if (arr.indexOf(un) < 0) arr.push(un);
      s.dirty = true;
      renderSettings(s);
    } else if (act === 'setCls') {
      const name = t.dataset.name;
      if (!t.value) delete s.classes[name];
      else s.classes[name] = { cls: t.value, costBySales: !!(s.classes[name] && s.classes[name].costBySales) };
      s.dirty = true;
      refreshClassRow(s, name);
      renderSettingsFooter(s);
    } else if (act === 'setCbs') {
      const name = t.dataset.name;
      if (!s.classes[name]) { t.checked = false; return; }
      s.classes[name].costBySales = !!t.checked;
      s.dirty = true;
      renderSettingsFooter(s);
    } else if (act === 'onlyUncls') {
      s.onlyUncls = !!t.checked;
      renderClassRows(s);
    }
  });
  s.ov.addEventListener('input', (ev) => {
    if (ev.target && ev.target.id === 'qapClsSearch') {
      s.clsSearch = ev.target.value;
      renderClassRows(s);
    }
  });

  renderSettingsLoading(s);
  const ok = await loadSettingsConfig(s);
  if (s.closed) return;
  if (!ok) { closeSettings(); return; }
  const me = s.cfg.me || {};
  if (!me.isAdmin && !me.isSealManager) {
    toast('你沒有簽核設定的權限');
    closeSettings();
    return;
  }
  s.busy = false;
  renderSettings(s);
}

function renderSettingsLoading(s) {
  s.body.innerHTML = '<div class="qap-muted" style="padding:30px;text-align:center">載入中…</div>';
  s.foot.innerHTML = '<button class="btn btn-secondary" type="button" data-act="close">關閉</button>';
}

/** 只更新頁尾的「有未儲存變更」提示，不重畫內容（避免搜尋框／捲動位置被重置） */
function renderSettingsFooter(s) {
  if (s.closed || !s.cfg) return;
  const hint = s.foot.querySelector('.qap-muted');
  if (hint) hint.textContent = s.dirty ? '有未儲存的變更' : '';
}

/** 單一商品列的歸類改變後，更新該列樣式與未歸類計數（不重畫整張表，避免篩選時列突然消失） */
function refreshClassRow(s, name) {
  const cur = s.classes[name];
  s.ov.querySelectorAll('#qapClsBody tr').forEach((tr) => {
    if (tr.dataset.name !== name) return;
    tr.classList.toggle('qap-uncls', !cur);
    const cb = tr.querySelector('input[data-act="setCbs"]');
    if (cb) { cb.disabled = !cur; if (!cur) cb.checked = false; }
    const first = tr.firstElementChild;
    const bd = first.querySelector('.qap-badge');
    if (!cur && !bd) first.insertAdjacentHTML('beforeend', ' <span class="qap-badge warn">未歸類</span>');
    if (cur && bd) bd.remove();
  });
  updateClassCount(s, s.ov.querySelectorAll('#qapClsBody tr[data-name]').length);
}

// ── 對外全域函式 ───────────────────────────────────
window.openQuoteApproval = openQuoteApproval;
window.openQuoteCostFill = openQuoteCostFill;
window.openQuoteApprovalSettings = openQuoteApprovalSettings;
window.refreshQuoteInbox = refreshQuoteInbox;
})();
