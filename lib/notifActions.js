'use strict';
/**
 * 鈴鐺通知的「清除」與「合約到期提醒已處理」純邏輯（無 I/O、不讀時鐘；時間一律由呼叫端傳入，方便測試）。
 *
 * 兩種完全不同的資料，不要混在一起：
 *   1) 儲存型通知  data.notifications[]（pushNotification 寫入，每人各一份，已讀只是取消未讀樣式）
 *        removeNotification  刪除「自己的」一則；找不到或不是自己的 → false（路由回 404，不洩漏他人通知存在與否）
 *        removeRead          刪除「自己所有已讀」；回刪除筆數
 *   2) 動態合約到期提醒  不存在於 data 內，每次由 contracts 即時推導（poll-bundle／GET /api/contract-reminders）
 *        「已處理」只能記一筆旁路：data.contractReminderDismissals = { `${username}|${contractId}|${endDate}`: ISO 時間 }
 *        鍵含迄日 → 合約續約改了迄日，鍵就對不上，提醒自動重新出現；沒改迄日（例如只改備註）則維持已處理。
 *        過濾必須發生在序列化（與 ETag 計算）之前，304 才不會用到舊內容。GET 內絕不寫入（清理只在 POST 時做）。
 *
 * contractReminderStatus 是 server.js 原本在 poll-bundle 與 /api/contract-reminders 各寫一份的 calcStatus，原樣搬來（行為不變）。
 */

const DISMISS_CAP = 5000;            // 已處理紀錄上限（超過就丟最舊的）
const DISMISS_STALE_DAYS = 400;      // 迄日早於今天這麼多天的紀錄沒有意義了（合約早已過期很久），寫入時順便清掉
const DAY_MS = 86400000;
const ID_RE = /^[A-Za-z0-9_-]{1,64}$/;

/** 合約的提醒狀態：{ key:'expired'|'urgent'|'expiring', days, isRenewed } 或 null（不需要提醒）。today＝當天 00:00 的 Date */
function contractReminderStatus(c, today) {
  if (!c.endDate) return null;
  const end = new Date(c.endDate);
  const endDiff = Math.ceil((end - today) / DAY_MS);
  let effectiveEnd, isRenewed = false;
  if (endDiff < 0 && c.renewDate) {
    effectiveEnd = new Date(c.renewDate);
    isRenewed = true;
  } else if (endDiff < 0) {
    return { key: 'expired', days: Math.abs(endDiff), isRenewed: false };
  } else {
    effectiveEnd = end;
  }
  const diff = Math.ceil((effectiveEnd - today) / DAY_MS);
  if (diff <= 25) return { key: 'urgent', days: diff, isRenewed };
  if (diff <= 90) return { key: 'expiring', days: diff, isRenewed };
  return null;   // 有效，不需要提醒
}

function isValidId(id) { return typeof id === 'string' && ID_RE.test(id); }

/** 已處理紀錄的鍵。迄日必須是合約當下的 endDate（字串原樣），改了迄日鍵就不同 */
function dismissKey(username, contractId, endDate) {
  return String(username) + '|' + String(contractId) + '|' + String(endDate == null ? '' : endDate);
}

function isDismissed(map, username, contractId, endDate) {
  if (!map || typeof map !== 'object') return false;
  return Object.prototype.hasOwnProperty.call(map, dismissKey(username, contractId, endDate));
}

/**
 * 從動態合約提醒中濾掉「本人已處理」的項目。沒有任何紀錄時原陣列原樣回傳（與修改前逐位元相同）。
 * reminders 的每一項需有 contractId 與 endDate。
 */
function filterDismissed(map, username, reminders) {
  if (!map || typeof map !== 'object' || !Object.keys(map).length) return reminders;
  return reminders.filter((r) => !isDismissed(map, username, r.contractId, r.endDate));
}

/** 寫入一筆已處理紀錄並順手清理（只准在 POST 呼叫）。data 會被修改。回傳目前紀錄筆數 */
function recordDismissal(data, username, contractId, endDate, nowIso) {
  if (!data.contractReminderDismissals || typeof data.contractReminderDismissals !== 'object' || Array.isArray(data.contractReminderDismissals)) data.contractReminderDismissals = {};
  const m = data.contractReminderDismissals;
  m[dismissKey(username, contractId, endDate)] = nowIso;
  pruneDismissals(m, nowIso);
  return Object.keys(m).length;
}

/** 清理：迄日早於 nowIso 400 天以上者刪除；超過 5000 筆時由最舊的開始丟。就地修改 */
function pruneDismissals(m, nowIso) {
  const nowMs = Date.parse(nowIso);
  if (Number.isFinite(nowMs)) {
    const cutoff = nowMs - DISMISS_STALE_DAYS * DAY_MS;
    Object.keys(m).forEach((k) => {
      const end = Date.parse(k.split('|').slice(2).join('|'));
      if (Number.isFinite(end) && end < cutoff) delete m[k];
    });
  }
  const keys = Object.keys(m);
  if (keys.length > DISMISS_CAP) {
    keys.sort((a, b) => String(m[a]).localeCompare(String(m[b])));
    keys.slice(0, keys.length - DISMISS_CAP).forEach((k) => { delete m[k]; });
  }
}

/** 刪除自己的一則儲存型通知。回 true＝已刪；false＝找不到，或那則不是自己的（呼叫端一律回 404，不區分） */
function removeNotification(data, id, username) {
  const list = Array.isArray(data.notifications) ? data.notifications : null;
  if (!list) return false;
  const i = list.findIndex((n) => n && n.id === id && n.to === username);
  if (i < 0) return false;
  data.notifications = list.slice(0, i).concat(list.slice(i + 1));
  return true;
}

/** 刪除自己所有「已讀」的儲存型通知。回刪除筆數；沒有要刪的就不動陣列 */
function removeRead(data, username) {
  const list = Array.isArray(data.notifications) ? data.notifications : null;
  if (!list) return 0;
  const keep = list.filter((n) => !(n && n.to === username && n.read === true));
  const removed = list.length - keep.length;
  if (removed) data.notifications = keep;
  return removed;
}

module.exports = {
  DISMISS_CAP, DISMISS_STALE_DAYS, ID_RE,
  contractReminderStatus, isValidId, dismissKey, isDismissed, filterDismissed, recordDismissal, pruneDismissals, removeNotification, removeRead,
};
