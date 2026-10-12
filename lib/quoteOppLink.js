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

/** 該報價單是否「還算數」：被駁回（returned）的不擋商機刪除。其餘（草稿 none／簽核中／已核准）都算 */
function isLiveQuote(q) {
  const st = q && q.approval ? q.approval.state : 'none';
  return st !== 'returned';
}

/** 連到某商機的報價單（全部／還算數的） */
function linkedQuotes(data, oppId) {
  const id = txt(oppId);
  if (!id) return [];
  return ((data && data.quotations) || []).filter((q) => q && q.oppId === id);
}

/** 一次算出 oppId → {total, live}（商機清單用，避免每筆商機各掃一次所有報價單） */
function countByOpp(data) {
  const m = new Map();
  ((data && data.quotations) || []).forEach((q) => {
    if (!q || !q.oppId) return;
    const c = m.get(q.oppId) || { total: 0, live: 0 };
    c.total++; if (isLiveQuote(q)) c.live++;
    m.set(q.oppId, c);
  });
  return m;
}

module.exports = { LINKABLE_STAGES, POOL_OWNER, findOpp, validateLink, snapshotOf, liveOf, liveFromOpp, companyEq, isLiveQuote, linkedQuotes, countByOpp, txt };
