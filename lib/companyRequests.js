'use strict';
/**
 * 統編查無（內部主檔與 GCIS 都查不到）→ 業務「申請行銷建檔」。純邏輯，可單元測試；server.js 負責 HTTP、通知、稽核與建主檔。
 *
 * 資料（全部放在既有主資料 blob）：
 *   data.companyRequests[]  { id, taxId, companyName, category, note, requestedBy, requestedAt, contactId, contactIds[],
 *                             requesters[{username,contactId,companyName,category,note,at}], status(pending|approved|rejected|cancelled),
 *                             resolvedBy, resolvedAt, result{name,address,industry,companyId}, rejectReason, rejectNote,
 *                             cancelledBy, lastRemindedDate }
 *   contact.companyVerify   未設＝一般；'pending'（待行銷確認公司）；'rejected'（被退回）。核可後清除
 *   contact.companyRequestId / companyVerifyReason   目前綁定的申請 id／被退回原因（顯示用）
 *
 * 規則（已與 Steven 確認）：同統編已有待處理申請 → 併單（多位申請人）；待確認期間只擋「報價單送簽」與「推送 SAP」；
 *   核可＝行銷／管理員；退回必填原因；申請人可取消（名片保留待確認狀態）。
 * 開關：COMPANY_REQUESTS=off 整個關閉（預設開）。本檔只提供純函式，是否啟用由呼叫端用 enabled() 判斷。
 */

// 申請類別（集中一處，之後可改）
const CATEGORIES = [
  { key: 'gov', label: '政府機關' },
  { key: 'school', label: '學校' },
  { key: 'hospital', label: '公立醫院' },
  { key: 'foundation', label: '財團法人' },
  { key: 'association', label: '協會團體' },
  { key: 'shop', label: '行號' },
  { key: 'other', label: '其他' },
];
const REJECT_REASONS = [
  { key: 'typo', label: '統編疑似輸入錯誤' },
  { key: 'no_need', label: '不需另建檔' },
  { key: 'insufficient', label: '資料不足' },
];
const NOTE_MIN = 5;
const NOTE_MAX = 200;
const REMIND_DAYS = 3;       // 待處理超過這個天數 → 提醒行銷（同一筆每天最多一次）
const OVERDUE_DAYS = 7;      // 超過這個天數 → 首頁待辦卡標紅
const MAX_REQUESTS = 5000;   // 總量上限（超過時丟最舊的已結案紀錄）
const DAY_MS = 86400000;

function enabled(env) {
  const e = env || process.env;
  const v = String(e.COMPANY_REQUESTS == null ? '' : e.COMPANY_REQUESTS).trim().toLowerCase();
  return v !== 'off';
}

const trim = (s) => String(s == null ? '' : s).trim();
// 去控制字元（保留換行以外都清掉）；說明是純文字，前端一律 textContent
const clean = (s) => trim(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '');

const categoryLabel = (k) => (CATEGORIES.find((c) => c.key === k) || {}).label || '';
const reasonLabel = (k) => (REJECT_REASONS.find((c) => c.key === k) || {}).label || '';

/** 驗證業務送來的 companyRequest {category,note}。回傳 { ok, value:{category,note} } 或 { ok:false, code, error } */
function validateRequest(r) {
  if (!r || typeof r !== 'object' || Array.isArray(r)) return { ok: false, code: 'BAD_REQUEST', error: '請填寫建檔申請的類別與說明' };
  const category = trim(r.category);
  if (!CATEGORIES.some((c) => c.key === category)) return { ok: false, code: 'BAD_CATEGORY', error: '請選擇申請類別' };
  const note = clean(r.note);
  if ([...note].length < NOTE_MIN) return { ok: false, code: 'NOTE_TOO_SHORT', error: `說明至少 ${NOTE_MIN} 個字（例如「衛福部健保署北區業務組」）` };
  if ([...note].length > NOTE_MAX) return { ok: false, code: 'NOTE_TOO_LONG', error: `說明不可超過 ${NOTE_MAX} 個字` };
  return { ok: true, value: { category, note } };
}

const isBlockingVerify = (contact) => !!contact && (contact.companyVerify === 'pending' || contact.companyVerify === 'rejected');

function list(data) { if (!Array.isArray(data.companyRequests)) data.companyRequests = []; return data.companyRequests; }

function findPendingByTax(data, taxId) {
  const t = trim(taxId);
  if (!t) return null;
  return list(data).find((r) => r.status === 'pending' && r.taxId === t) || null;
}

/** 內部主檔是否已有這個統編（有＝不需要申請）。 */
function masterExists(data, taxId) {
  const t = trim(taxId);
  if (!/^\d{8}$/.test(t)) return false;
  return (data.companies || []).some((c) => c && (c.matchKey === 'tax:' + t || c.taxId === t));
}

function trimStore(data) {
  const l = list(data);
  if (l.length <= MAX_REQUESTS) return;
  // 先丟最舊的已結案紀錄（pending 一律保留）
  const keep = l.filter((r) => r.status === 'pending');
  const done = l.filter((r) => r.status !== 'pending').sort((a, b) => String(b.requestedAt).localeCompare(String(a.requestedAt)));
  data.companyRequests = keep.concat(done.slice(0, Math.max(0, MAX_REQUESTS - keep.length)));
}

/**
 * 送出（或併入）建檔申請，並把名片標成待確認。
 * @param p { contact, user:{username}, request:{category,note}, now(ISO), genId }
 * @returns { request, merged:boolean }
 */
function submitRequest(data, p) {
  const { contact, user, request, now, genId } = p;
  const taxId = trim(contact.taxId);
  const who = trim(user && user.username);
  let r = findPendingByTax(data, taxId);
  let merged = false;
  const entry = { username: who, contactId: contact.id, companyName: trim(contact.company), category: request.category, note: request.note, at: now };
  if (r) {
    merged = true;
    if (!Array.isArray(r.requesters)) r.requesters = [];
    if (!Array.isArray(r.contactIds)) r.contactIds = r.contactId ? [r.contactId] : [];
    const i = r.requesters.findIndex((x) => x.contactId === contact.id);
    if (i >= 0) r.requesters[i] = entry; else r.requesters.push(entry);
    if (!r.contactIds.includes(contact.id)) r.contactIds.push(contact.id);
  } else {
    r = {
      id: genId(), taxId, companyName: trim(contact.company), category: request.category, note: request.note,
      requestedBy: who, requestedAt: now, contactId: contact.id, contactIds: [contact.id], requesters: [entry],
      status: 'pending', resolvedBy: '', resolvedAt: '', result: null, rejectReason: '', rejectNote: '', cancelledBy: '', lastRemindedDate: '',
    };
    list(data).unshift(r);
    trimStore(data);
  }
  contact.companyVerify = 'pending';
  contact.companyRequestId = r.id;
  delete contact.companyVerifyReason;
  return { request: r, merged };
}

/** 把名片從它目前綁定的待處理申請移除（改統編、已命中主檔等）。申請人都走光了 → 申請自動作廢。回傳被作廢的申請或 null */
function detachContact(data, contact, now, by) {
  const r = list(data).find((x) => x.id === contact.companyRequestId && x.status === 'pending');
  if (!r) return null;
  r.requesters = (r.requesters || []).filter((x) => x.contactId !== contact.id);
  r.contactIds = (r.contactIds || []).filter((x) => x !== contact.id);
  if (!r.requesters.length) { r.status = 'cancelled'; r.cancelledBy = by || 'system'; r.resolvedAt = now; return r; }
  r.contactId = r.contactIds[0] || r.contactId;
  return null;
}

function clearVerify(contact) {
  delete contact.companyVerify;
  delete contact.companyRequestId;
  delete contact.companyVerifyReason;
}

/** 申請人取消：只移除「他自己」的申請人紀錄；名片保留待確認狀態（companyRequestId 清空）。 */
function cancelRequest(data, id, username, now) {
  const r = list(data).find((x) => x.id === id);
  if (!r) return { ok: false, status: 404, error: '找不到此建檔申請' };
  if (!(r.requesters || []).some((x) => x.username === username)) return { ok: false, status: 403, error: '只有申請人可以取消自己的申請' };
  if (r.status !== 'pending') return { ok: false, status: 409, error: '這筆申請已處理完畢，無法取消' };
  const mine = (r.requesters || []).filter((x) => x.username === username);
  const myCids = new Set(mine.map((x) => x.contactId));
  r.requesters = r.requesters.filter((x) => x.username !== username);
  r.contactIds = (r.contactIds || []).filter((c) => !myCids.has(c));
  (data.contacts || []).forEach((c) => { if (myCids.has(c.id) && c.companyRequestId === r.id) { c.companyRequestId = ''; } });
  let fully = false;
  if (!r.requesters.length) { r.status = 'cancelled'; r.cancelledBy = username; r.resolvedAt = now; fully = true; }
  else r.contactId = r.contactIds[0] || r.contactId;
  return { ok: true, request: r, fullyCancelled: fully };
}

/**
 * 核可（主檔由呼叫端用 ensureCompanyMaster 建好後傳入 master）：
 * 申請 → approved；所有同統編名片清掉 companyVerify 並掛上主檔。回傳受影響的名片與要通知的申請人
 */
function applyApproval(data, r, master, actor, now, result) {
  r.status = 'approved'; r.resolvedBy = actor; r.resolvedAt = now;
  r.result = { name: result.name, address: result.address || '', industry: result.industry || '', companyId: master.id };
  const touched = [];
  (data.contacts || []).forEach((c) => {
    if (!c || trim(c.taxId) !== r.taxId) return;
    if (!c.companyVerify && !(r.contactIds || []).includes(c.id)) return;
    clearVerify(c);
    if (!c.deleted) c.companyId = master.id;
    touched.push(c.id);
  });
  const notify = [];
  const seen = new Set();
  (r.requesters || []).forEach((x) => { if (!seen.has(x.username)) { seen.add(x.username); notify.push({ username: x.username, contactId: x.contactId }); } });
  return { touched, notify };
}

/** 退回：必填原因 key（＋備註）。申請 → rejected；綁定的名片 → companyVerify='rejected'＋原因文字 */
function validateReject(b) {
  const reason = trim(b && b.reason);
  if (!REJECT_REASONS.some((x) => x.key === reason)) return { ok: false, code: 'REASON_REQUIRED', error: '退回必須選擇原因' };
  const note = clean(b && b.note);
  if ([...note].length > NOTE_MAX) return { ok: false, code: 'NOTE_TOO_LONG', error: `備註不可超過 ${NOTE_MAX} 個字` };
  return { ok: true, value: { reason, note } };
}
function applyReject(data, r, actor, v, now) {
  r.status = 'rejected'; r.resolvedBy = actor; r.resolvedAt = now; r.rejectReason = v.reason; r.rejectNote = v.note;
  const text = reasonLabel(v.reason) + (v.note ? '：' + v.note : '');
  const notify = [];
  const seen = new Set();
  (data.contacts || []).forEach((c) => {
    if (!c || c.companyRequestId !== r.id) return;
    c.companyVerify = 'rejected'; c.companyVerifyReason = text;
  });
  (r.requesters || []).forEach((x) => { if (!seen.has(x.username)) { seen.add(x.username); notify.push({ username: x.username, contactId: x.contactId }); } });
  return { notify, text };
}

const ageDays = (r, nowMs) => Math.max(0, Math.floor((nowMs - new Date(r.requestedAt).getTime()) / DAY_MS));

/** 序列化給前端：行銷／管理員看全部；其他人只看「自己是申請人」的（且只帶自己的那份申請人資料） */
function serialize(r, nowMs, viewer, usersByName, userLabel) {
  const label = (un) => (userLabel ? userLabel(usersByName && usersByName[un], un) : un);
  const all = viewer.canResolve;
  const reqs = (r.requesters || []).filter((x) => all || x.username === viewer.username);
  const d = ageDays(r, nowMs);
  return {
    id: r.id, taxId: r.taxId, companyName: r.companyName, category: r.category, categoryLabel: categoryLabel(r.category), note: r.note,
    status: r.status, requestedAt: r.requestedAt, days: d, overdue: r.status === 'pending' && d > OVERDUE_DAYS,
    requestedBy: r.requestedBy, requestedByLabel: label(r.requestedBy),
    requesters: reqs.map((x) => ({ username: x.username, label: label(x.username), contactId: x.contactId, companyName: x.companyName, category: x.category, categoryLabel: categoryLabel(x.category), note: x.note, at: x.at })),
    requesterCount: (r.requesters || []).length,
    resolvedBy: r.resolvedBy || '', resolvedAt: r.resolvedAt || '', result: r.result || null,
    rejectReason: r.rejectReason || '', rejectReasonLabel: reasonLabel(r.rejectReason), rejectNote: r.rejectNote || '',
  };
}

function visibleList(data, viewer, nowMs, usersByName, userLabel) {
  const l = list(data).filter((r) => viewer.canResolve || (r.requesters || []).some((x) => x.username === viewer.username) || r.requestedBy === viewer.username);
  return l.map((r) => serialize(r, nowMs, viewer, usersByName, userLabel));
}

/** 待處理超過 REMIND_DAYS 天、且今天還沒提醒過的申請（純挑選，不改資料） */
function dueReminders(data, nowMs, ymd) {
  return list(data).filter((r) => r.status === 'pending' && ageDays(r, nowMs) >= REMIND_DAYS && r.lastRemindedDate !== ymd);
}

/** 報價單送簽閘門：回傳 null＝放行；否則 { code, message }。對應名片：q.contactId → 商機聯絡人 → 同擁有者同公司名 */
function blockForQuote(data, q) {
  if (!q) return null;
  const contacts = (data.contacts || []).filter((c) => c && !c.deleted);
  let hit = null;
  if (q.contactId) hit = contacts.find((c) => c.id === q.contactId && isBlockingVerify(c)) || null;
  if (!hit && q.oppId) {
    const opp = (data.opportunities || []).find((o) => o && o.id === q.oppId);
    if (opp && opp.contactId) hit = contacts.find((c) => c.id === opp.contactId && isBlockingVerify(c)) || null;
  }
  if (!hit && q.company) {
    const nm = trim(q.company);
    hit = contacts.find((c) => c.owner === q.owner && trim(c.company) === nm && isBlockingVerify(c)) || null;
  }
  if (!hit) return null;
  return { code: 'COMPANY_UNVERIFIED', message: unverifiedMessage(hit) };
}

function unverifiedMessage(contact) {
  return contact && contact.companyVerify === 'rejected'
    ? `這張報價單的客戶公司（統編 ${contact.taxId}）建檔申請被退回，請先修正統編後才能送簽`
    : `這張報價單的客戶公司（統編 ${contact && contact.taxId}）還在等行銷確認建檔，確認後才能送簽`;
}

/** SAP 推送閘門：companies 為主檔物件陣列；有待處理申請、或同統編名片仍待確認／被退回者擋下。回傳 null 或 { code, message, names } */
function blockForSapCompanies(data, companies) {
  const bad = [];
  (companies || []).forEach((m) => {
    if (!m) return;
    const t = trim(m.taxId);
    const pending = t && list(data).some((r) => r.status === 'pending' && r.taxId === t);
    const cards = (data.contacts || []).some((c) => c && !c.deleted && isBlockingVerify(c) && (c.companyId === m.id || (t && trim(c.taxId) === t)));
    if (pending || cards) bad.push(m.name || t || m.id);
  });
  if (!bad.length) return null;
  return { code: 'COMPANY_UNVERIFIED', message: `以下公司的建檔申請尚未通過，不能推送 SAP：${bad.slice(0, 5).join('、')}${bad.length > 5 ? `等 ${bad.length} 家` : ''}`, names: bad };
}

module.exports = {
  CATEGORIES, REJECT_REASONS, NOTE_MIN, NOTE_MAX, REMIND_DAYS, OVERDUE_DAYS,
  enabled, validateRequest, validateReject, isBlockingVerify, findPendingByTax, masterExists,
  submitRequest, detachContact, clearVerify, cancelRequest, applyApproval, applyReject,
  ageDays, serialize, visibleList, dueReminders, blockForQuote, blockForSapCompanies, unverifiedMessage,
  categoryLabel, reasonLabel,
};
