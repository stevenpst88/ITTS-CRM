'use strict';
/**
 * 企業主檔「責任業務」＋「待補名片」待辦（推導狀態）。
 *
 * 資料模型（全部放在既有主資料 blob，不另開檔）：
 *   company.responsibleSales        帳號陣列（行銷／管理員／主管指定）
 *   company.responsibleAssignedAt   { 帳號: ISO }  每位責任業務「被指定」的時間（逾期天數、解除紀錄比對都靠它）
 *   data.companyCardDismissals      手動解除紀錄（陣列）；待辦本身「不另存」，永遠由名片資料即時推導
 *
 * 待辦 ＝ 對「每一位」責任業務各自判斷：該業務名下對這家公司的「真實名片」為零。
 *   真實名片 ＝ 未刪除、非客戶池 placeholder、owner 為該業務；
 *   公司比對以 contact.companyId === company.id 為主；contact 沒有 companyId 時才以公司名（trim）比對。
 *   別人補了名片不會讓他免責；他自己補了真實名片 → 待辦自動消失（不會殘留）。
 *
 * 手動解除：必須選原因；紀錄會綁定「當時的指定時間（assignedAt）」。
 *   行銷重新指定（assignedAt 變新）→ 舊紀錄自動失效 ＝ 「重新指定會重置該業務的解除」，不需要另外清資料。
 *
 * 開關：COMPANY_CARD_TASKS=off 整個關閉（預設開）。本檔只提供純函式，是否啟用由呼叫端用 enabled() 判斷。
 */

// 可被指定為責任業務、也就是「可擁有名片」的角色（排除 pool、唯讀角色 tecopm、秘書、顧問主管、會計／財務、行銷、董事長）
const OWNER_ROLES = ['user', 'manager1', 'manager2', 'groupsales'];
// 可寫入責任業務的角色（再依各自的可視範圍收斂，見 server.js 的 respCanTouchCompany）
const WRITE_ROLES = ['admin', 'marketing', 'manager1', 'manager2', 'executive'];
// 手動解除原因（key → 顯示文字）
const DISMISS_REASONS = {
  not_mine: '不是我負責',
  inactive: '客戶已不往來',
  no_contact: '還沒有聯絡人',
  other: '其他',
};
const OVERDUE_DAYS = 14;          // 被指定超過這個天數仍未處理 → 主管首頁
const MAX_NOTE = 300;
const MAX_RESPONSIBLE = 20;       // 單一公司責任業務人數上限
const MAX_DISMISSALS = 3000;      // 解除紀錄總量上限（超過時先丟最舊的已處理紀錄）
const DAY_MS = 86400000;

function enabled(env) {
  const v = String((env || process.env).COMPANY_CARD_TASKS == null ? '' : (env || process.env).COMPANY_CARD_TASKS).trim().toLowerCase();
  return v !== 'off';
}

function trimName(s) { return String(s == null ? '' : s).trim(); }

function normList(v) {
  const out = [];
  const seen = new Set();
  (Array.isArray(v) ? v : []).forEach((x) => {
    const s = trimName(x);
    if (s && !seen.has(s)) { seen.add(s); out.push(s); }
  });
  return out;
}

/** 這個帳號能不能被指定為責任業務（存在、啟用、可擁有名片、不是 _pool／pool）。 */
function eligibleOwner(user) {
  if (!user) return false;
  if (user.username === '_pool' || user.role === 'pool') return false;
  if (user.active === false) return false;
  return OWNER_ROLES.includes(user.role);
}

/** 回傳 { ok, error }：檢查「新加入」的帳號是否合法（保留的舊帳號不重驗，避免離職者擋住其他人的編輯）。 */
function validateNewOwners(newUsernames, usersByName) {
  for (const un of newUsernames) {
    const u = usersByName[un];
    if (!u) return { ok: false, error: `找不到帳號「${un}」` };
    if (u.username === '_pool' || u.role === 'pool') return { ok: false, error: '不能指定客戶池為責任業務' };
    if (u.active === false) return { ok: false, error: `帳號「${un}」已停用，不能指定為責任業務` };
    if (!OWNER_ROLES.includes(u.role)) return { ok: false, error: `帳號「${un}」的角色不能擁有名片，不能指定為責任業務` };
  }
  return { ok: true };
}

/** 名片索引：companyId → Set(owner)；沒有 companyId 的名片 → 公司名 → Set(owner)。只收真實名片。 */
function buildRealCardIndex(contacts) {
  const byId = new Map();
  const byName = new Map();
  (contacts || []).forEach((c) => {
    if (!c || c.deleted || c.isPlaceholder || !c.owner) return;
    if (c.companyId) {
      if (!byId.has(c.companyId)) byId.set(c.companyId, new Set());
      byId.get(c.companyId).add(c.owner);
    } else {
      const n = trimName(c.company);
      if (!n) return;
      if (!byName.has(n)) byName.set(n, new Set());
      byName.get(n).add(c.owner);
    }
  });
  return { byId, byName };
}

function hasRealCard(index, company, username) {
  const a = index.byId.get(company.id);
  if (a && a.has(username)) return true;
  const n = trimName(company.name);
  const b = n ? index.byName.get(n) : null;
  return !!(b && b.has(username));
}

function assignedAtOf(company, username) {
  const m = company && company.responsibleAssignedAt;
  const v = m && typeof m === 'object' ? m[username] : '';
  return typeof v === 'string' ? v : '';
}

function daysSince(iso, now) {
  const t = iso ? Date.parse(iso) : NaN;
  if (isNaN(t)) return 0;
  return Math.max(0, Math.floor(((now instanceof Date ? now.getTime() : Number(now)) - t) / DAY_MS));
}

/** 找出生效中的解除紀錄（綁定當時的指定時間；已處理 resolved 也仍然有效，只是不再出現在行銷清單）。 */
function activeDismissal(dismissals, companyId, username, assignedAt) {
  let best = null;
  (dismissals || []).forEach((d) => {
    if (d && d.companyId === companyId && d.username === username && (d.assignedAt || '') === (assignedAt || '')) {
      if (!best || String(d.dismissedAt) > String(best.dismissedAt)) best = d;
    }
  });
  return best;
}

/**
 * 推導所有待辦。回傳陣列（每筆＝一家公司 × 一位責任業務、且其名下真實名片為零）：
 *   { companyId, name, taxId, industry, address, website, username, assignedAt, days, dismissal }
 * dismissal 為生效中的解除紀錄或 null。呼叫端自行決定是否排除已解除者。
 * 責任業務帳號不存在／已停用／不可擁有名片 → 不產生待辦（沒有人能補）。
 */
function deriveTasks(data, usersByName, now) {
  const out = [];
  const index = buildRealCardIndex((data && data.contacts) || []);
  const dismissals = (data && data.companyCardDismissals) || [];
  ((data && data.companies) || []).forEach((m) => {
    const rs = Array.isArray(m.responsibleSales) ? m.responsibleSales : [];
    rs.forEach((un) => {
      if (!eligibleOwner(usersByName[un])) return;
      if (hasRealCard(index, m, un)) return;
      const at = assignedAtOf(m, un);
      out.push({
        companyId: m.id, name: m.name || '', taxId: m.taxId || '', industry: m.industry || '',
        address: m.address || '', website: m.website || '',
        username: un, assignedAt: at, days: daysSince(at, now),
        dismissal: activeDismissal(dismissals, m.id, un, at),
      });
    });
  });
  return out;
}

/**
 * 套用責任業務變更（就地修改 company）。
 *   newList：最終名單（已含保留的舊帳號）；fresh：這些帳號即使已在名單內，也重新蓋指定時間（行銷「重新指定」）
 * 回傳 { added, removed, kept, changed }。
 */
function applyResponsibleSales(company, newList, nowIso, fresh) {
  const old = Array.isArray(company.responsibleSales) ? company.responsibleSales.slice() : [];
  const freshSet = new Set(fresh || []);
  const oldStamp = (company.responsibleAssignedAt && typeof company.responsibleAssignedAt === 'object') ? company.responsibleAssignedAt : {};
  const stamp = {};
  newList.forEach((un) => {
    stamp[un] = (old.includes(un) && oldStamp[un] && !freshSet.has(un)) ? oldStamp[un] : nowIso;
  });
  const added = newList.filter((u) => !old.includes(u));
  const removed = old.filter((u) => !newList.includes(u));
  const kept = newList.filter((u) => old.includes(u));
  const refreshed = newList.filter((u) => freshSet.has(u) && old.includes(u));
  company.responsibleSales = newList.slice();
  company.responsibleAssignedAt = stamp;
  return { added, removed, kept, changed: !!(added.length || removed.length || refreshed.length), before: old, after: newList.slice() };
}

/** 清掉文字裡的控制字元、截長度（備註用）。 */
function cleanNote(s) {
  return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, '').trim().slice(0, MAX_NOTE);
}

/**
 * 業務手動解除。task 必須「目前確實存在且尚未解除」才能解除（防止亂寫紀錄）。
 * 回傳 { ok, error, status, dismissal }。
 */
function dismissTask(data, usersByName, args, now) {
  const { companyId, username, reason, note } = args || {};
  if (!Object.prototype.hasOwnProperty.call(DISMISS_REASONS, reason)) return { ok: false, status: 400, error: '請選擇解除原因' };
  const cleaned = cleanNote(note);
  if (reason === 'other' && !cleaned) return { ok: false, status: 400, error: '選「其他」時必須填寫說明' };
  const task = deriveTasks(data, usersByName, now).find((t) => t.companyId === companyId && t.username === username);
  if (!task) return { ok: false, status: 404, error: '找不到這項待辦（可能已補名片或已被取消）' };
  if (task.dismissal) return { ok: false, status: 409, error: '這項待辦已經解除過了' };
  if (!Array.isArray(data.companyCardDismissals)) data.companyCardDismissals = [];
  const nowIso = (now instanceof Date ? now : new Date(now)).toISOString();
  const d = {
    id: 'ccd_' + nowIso.replace(/\D/g, '').slice(0, 14) + '_' + Math.random().toString(36).slice(2, 8),
    companyId, companyName: task.name, username,
    reason, reasonLabel: DISMISS_REASONS[reason], note: cleaned,
    assignedAt: task.assignedAt, dismissedAt: nowIso,
    status: 'open', resolvedBy: '', resolvedAt: '', resolution: '', resolveNote: '',
  };
  data.companyCardDismissals.unshift(d);
  pruneDismissals(data);
  return { ok: true, dismissal: d, task };
}

function pruneDismissals(data) {
  const arr = data.companyCardDismissals;
  if (!Array.isArray(arr) || arr.length <= MAX_DISMISSALS) return;
  // 先丟最舊的「已處理」，再不夠才丟最舊的任何一筆
  const keep = arr.slice(0, MAX_DISMISSALS);
  data.companyCardDismissals = keep;
}

/**
 * 「業務回報待修」清單：尚未被行銷處理（status=open）、且該回報目前仍成立
 *   （責任業務仍在名單內、指定時間沒被重蓋、其名下真實名片仍為零）。
 */
function feedbackItems(data, usersByName, now, onlyStatus) {
  const want = onlyStatus || 'open';
  const tasks = deriveTasks(data, usersByName, now);
  const out = [];
  tasks.forEach((t) => {
    const d = t.dismissal;
    if (!d || d.status !== want) return;
    out.push({
      id: d.id, companyId: t.companyId, companyName: t.name, taxId: t.taxId,
      username: t.username, reason: d.reason, reasonLabel: d.reasonLabel, note: d.note || '',
      dismissedAt: d.dismissedAt, assignedAt: t.assignedAt, days: daysSince(d.dismissedAt, now), status: d.status,
    });
  });
  out.sort((a, b) => String(b.dismissedAt).localeCompare(String(a.dismissedAt)));
  return out;
}

/** 行銷標示已處理 / 重新指定後結案。回傳 { ok, error, status, dismissal }。 */
function resolveFeedback(data, id, actor, resolution, note, nowIso) {
  const d = (data.companyCardDismissals || []).find((x) => x && x.id === id);
  if (!d) return { ok: false, status: 404, error: '找不到這筆回報' };
  if (d.status !== 'open') return { ok: false, status: 409, error: '這筆回報已經處理過了' };
  d.status = 'resolved';
  d.resolvedBy = actor;
  d.resolvedAt = nowIso;
  d.resolution = resolution;
  d.resolveNote = cleanNote(note);
  return { ok: true, dismissal: d };
}

module.exports = {
  OWNER_ROLES, WRITE_ROLES, DISMISS_REASONS, OVERDUE_DAYS, MAX_NOTE, MAX_RESPONSIBLE, DAY_MS,
  enabled, trimName, normList, eligibleOwner, validateNewOwners,
  buildRealCardIndex, hasRealCard, assignedAtOf, daysSince, activeDismissal,
  deriveTasks, applyResponsibleSales, cleanNote, dismissTask, feedbackItems, resolveFeedback,
};
