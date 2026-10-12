'use strict';
/**
 * 報價單 ↔ 商機連結（q.oppId）的純函式。server.js（商機刪除／改公司的防呆、商機清單的報價數）與 lib/quoteRoutes.js（建單／存檔驗證、序列化）共用。
 *
 * 規則（Steven 2026-10-12 定案）：
 *  1. 新報價單必須連一筆商機；可連階段 A／B／C／D（不含 Won）。
 *  2. 只能連自己的商機（opp.owner === 建單者）。沒有跨 owner、KA、admin 例外。客戶池（_pool）商機天然排除（建單者不會是 _pool）。
 *  3. 舊單沒有 oppId：欄位完全不存在，所有功能照舊，沒有任何擋關。
 *  4. 一筆商機可連多張報價單。
 *  5. oppId／oppSnapshot 不進 contentHash／itemsSig／costLinesSig，不影響核准。
 *  6. 不自動改商機階段、不回寫金額。商機金額單位是 NT$K（千元），報價是元。
 */
const LINKABLE_STAGES = ['A', 'B', 'C', 'D'];
const POOL_OWNER = '_pool';
const txt = (s) => String(s === undefined || s === null ? '' : s).trim();

function findOpp(data, oppId) {
  const id = txt(oppId);
  if (!id) return null;
  return ((data && data.opportunities) || []).find((o) => o && o.id === id) || null;
}

/** 驗證「me 能不能把報價單連到 oppId」。回 {ok:true, opp} 或 {ok:false, status, code, message} */
function validateLink(data, oppId, me) {
  const id = txt(oppId);
  if (!id) return { ok: false, status: 400, code: 'OPP_REQUIRED', message: '新報價單必須連結一筆商機，請先選擇商機（找不到請先建立商機）' };
  const opp = findOpp(data, id);
  if (!opp) return { ok: false, status: 400, code: 'OPP_NOT_FOUND', message: '找不到這筆商機（可能已被刪除），請重新選擇' };
  if (opp.owner === POOL_OWNER || opp.owner !== me) return { ok: false, status: 403, code: 'OPP_NOT_YOURS', message: '只能連結自己的商機' };
  if (!LINKABLE_STAGES.includes(opp.stage)) return { ok: false, status: 400, code: 'OPP_STAGE_NOT_ALLOWED', message: '只有 A／B／C／D 階段的商機可以建立報價單（此商機階段：' + (opp.stage || '未設定') + '）' };
  return { ok: true, opp };
}

/** 連結當下商機的摘要（商機之後被刪除或改動，報價單仍可顯示當時的樣子）。amountK 單位＝NT$K（千元） */
function snapshotOf(opp, nowIso) {
  const amt = parseFloat(opp.amount);
  return {
    product: txt(opp.product), amountK: Number.isFinite(amt) ? amt : null, bu: opp.bu === undefined || opp.bu === null ? '' : opp.bu,
    stage: txt(opp.stage), company: txt(opp.company), linkedAt: nowIso,
  };
}

/** 讀取當下的商機狀態（序列化用）。商機被刪除（搬到 lostOpportunities）→ exists:false */
function liveOf(data, oppId) { return liveFromOpp(findOpp(data, oppId)); }
function liveFromOpp(opp) {
  if (!opp) return { exists: false, stage: null, product: '' };
  return { exists: true, stage: txt(opp.stage), product: txt(opp.product) };
}

const companyEq = (a, b) => txt(a) === txt(b);

/** 該報價單是否「算數」：沒有被標記失效（q.invalid）。商機作廢時所連結的報價單一律失效（草稿、被退回、簽核中、已核准都一樣） */
function isLiveQuote(q) { return !!q && !q.invalid; }

/** 連到某商機的報價單（全部，含失效的） */
function linkedQuotes(data, oppId) {
  const id = txt(oppId);
  if (!id) return [];
  return ((data && data.quotations) || []).filter((q) => q && q.oppId === id);
}

/** 一次算出 oppId → {total, live, invalid}（商機清單用，避免每筆商機各掃一次所有報價單）。live＝沒失效的 */
function countByOpp(data) {
  const m = new Map();
  ((data && data.quotations) || []).forEach((q) => {
    if (!q || !q.oppId) return;
    const c = m.get(q.oppId) || { total: 0, live: 0, invalid: 0 };
    c.total++; if (isLiveQuote(q)) c.live++; else c.invalid++;
    m.set(q.oppId, c);
  });
  return m;
}

// ── 商機作廢 → 連結的報價單失效；商機還原 → 因這筆商機失效的報價單恢復 ──────────────────────────────
// 失效＝q.invalid（{at, by, reason:'OPP_VOID', oppId, company, prevState}）。不動 contentHash 涵蓋的任何欄位；唯讀由 quoteRoutes 的 QUOTE_INVALID 守衛執行。
// 簽核中的單：關掉待簽關卡（approval 回到 none、steps 清空、hash 清掉，歷程保留），這樣簽核人的待辦、信件逾期檢查、收件匣都不會再看到它。
// 已核准的單：approval 維持原樣（當作紀錄），但不可下載正式（蓋章）檔、不可改。
// 草稿（approval 為 null）只留時間軸，不建立 approval 物件（草稿仍可刪除）。
const MAX_TL = 200;
function pushTl(q, by, action, comment, at) {
  if (!Array.isArray(q.timeline)) q.timeline = [];
  q.timeline.push({ at, by, action, comment });
  if (q.timeline.length > MAX_TL) q.timeline = q.timeline.slice(-MAX_TL);
}
function pushHist(q, by, action, comment, at) {
  const ap = q.approval;
  if (!ap || typeof ap !== 'object') return false;
  if (!Array.isArray(ap.history)) ap.history = [];
  ap.history.push({ at, by, action, comment, tier: '' });
  return true;
}
/** 作廢商機：回傳受影響的報價單陣列（[{ q, prevState }]）。必須和商機搬到 lostOpportunities 在同一次 db.save 內呼叫 */
function invalidateForOpp(data, opp, by, nowIso) {
  const out = [];
  linkedQuotes(data, opp.id).forEach((q) => {
    if (q.invalid) return;
    const prevState = q.approval && q.approval.state ? q.approval.state : 'none';
    q.invalid = { at: nowIso, by, reason: 'OPP_VOID', oppId: opp.id, company: txt(opp.company), prevState };
    const ap = q.approval;
    if (ap && ap.state === 'pending') {
      ap.state = 'none'; ap.steps = []; ap.cur = 0; ap.hash = null; ap.derived = null; ap.board = null; ap.submittedAt = null; ap.submittedBy = null;
    }
    // 有簽核紀錄的單寫進簽核歷程；草稿（沒有 approval）寫進編輯時間軸。兩處在畫面上會合併顯示，只寫一邊以免出現兩行
    if (!pushHist(q, by, 'OPP_VOID_INVALIDATE', '連結的商機已作廢，報價單失效' + (prevState === 'pending' ? '（簽核中的關卡已關閉）' : ''), nowIso)) pushTl(q, by, 'OPP_VOID', '連結的商機已作廢，報價單失效', nowIso);
    out.push({ q, prevState });
  });
  return out;
}
/** 還原商機：因這筆商機作廢而失效的報價單恢復（簽核中被關閉的單回到草稿，需重新送簽）。回傳恢復的報價單陣列 */
function revalidateForOpp(data, opp, by, nowIso) {
  const out = [];
  linkedQuotes(data, opp.id).forEach((q) => {
    if (!q.invalid || q.invalid.reason !== 'OPP_VOID' || q.invalid.oppId !== opp.id) return;
    const prev = q.invalid.prevState;
    delete q.invalid;
    if (!pushHist(q, by, 'OPP_RESTORE', '商機已還原，報價單恢復' + (prev === 'pending' ? '（原簽核中，需重新送簽）' : ''), nowIso)) pushTl(q, by, 'OPP_RESTORE', '商機已還原，報價單恢復', nowIso);
    out.push(q);
  });
  return out;
}

module.exports = { LINKABLE_STAGES, POOL_OWNER, findOpp, validateLink, snapshotOf, liveOf, liveFromOpp, companyEq, isLiveQuote, linkedQuotes, countByOpp, invalidateForOpp, revalidateForOpp, txt };
