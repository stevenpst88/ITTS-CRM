'use strict';
/**
 * 組織關係解析（收件人／上級）共用純函式。MA_TRACKING（成交納管、任務卡、之後的提醒規則）使用；
 * 既有的兩份實作（lib/quoteRoutes.js 的 findManager1、server.js 合約提醒內會退回 BU 的版本）刻意不動、不改行為。
 *
 * users 可以是帳號陣列或 {username: user} 對照表。
 * 排除：停用帳號（active===false）、客戶池（role==='pool'／_pool）、唯讀掛名主管（pureSupervisor）。
 */

function toMap(users) {
  if (Array.isArray(users)) {
    const m = {};
    users.forEach((u) => { if (u && u.username) m[u.username] = u; });
    return m;
  }
  return users && typeof users === 'object' ? users : {};
}

/** 在職且可當收件人／處理人：存在、未停用、非客戶池 */
function isActiveUser(u) { return !!u && u.active !== false && u.role !== 'pool' && u.username !== '_pool'; }

/** 可當「一級主管」收件人：在職 manager1 且不是唯讀掛名（pureSupervisor） */
function isNotifiableManager1(u) { return isActiveUser(u) && u.role === 'manager1' && !u.pureSupervisor; }

/**
 * 業務的一級主管：沿 supervisor 鏈往上找「第一位」在職、非 pureSupervisor 的 manager1；
 * 中間遇到停用或 pureSupervisor 的 manager1 就跳過繼續往上。找不到 → null（不退回 BU 推斷，避免通知到不相干的人）。
 * 有循環保護。
 */
function findManager1(users, ownerUsername) {
  const map = toMap(users);
  const seen = new Set();
  let cur = map[ownerUsername];
  while (cur && cur.supervisor && !seen.has(cur.username)) {
    seen.add(cur.username);
    const up = map[cur.supervisor];
    if (!up) break;
    if (isNotifiableManager1(up)) return up.username;
    cur = up;
  }
  return null;
}

/** 在職行銷（role==='marketing' 且 active!==false），依帳號排序以求輸出穩定 */
function marketingUsers(users) {
  return Object.values(toMap(users)).filter((u) => isActiveUser(u) && u.role === 'marketing').map((u) => u.username).sort();
}

module.exports = { toMap, isActiveUser, isNotifiableManager1, findManager1, marketingUsers };
