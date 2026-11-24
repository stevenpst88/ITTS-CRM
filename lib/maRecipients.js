'use strict';
/**
 * MA／License 到期提醒的收件人解析（P3a，MA_TRACKING=1）。純函式、不碰資料庫與 req，方便單元測試。
 *
 * 角色（規則的 notify 白名單）：
 *   owner      納管項的負責業務。帳號已停用（或是唯讀掛名主管）→ 改取其一級主管；仍找不到 → 管理員。
 *   secretary  role==='secretary'、在職、BU 與該納管項 BU 有交集（多位全部通知）。沒有 → 退到管理員（回報 issue NO_SECRETARY）。
 *   manager1   沿 supervisor 找第一位在職且非 pureSupervisor 的 manager1（lib/orgResolve.js findManager1，與成交納管共用）。找不到 → 退到管理員（issue NO_MANAGER1）。
 *   manager2   負責業務的直屬主管，且其 role 為 manager2（在職、非 pureSupervisor）。找不到 → 沒有人（不退管理員）。
 *   marketing  所有 role==='marketing' 且在職者。
 *
 * 一律排除：停用帳號（active===false）、客戶池、唯讀掛名主管（pureSupervisor）。同一個人在同一組 notify 內只會出現一次（roles 合併）。
 * 收件人在「發送當下」重新解析，不綁死舊帳號（業務離職或換手後自動改取主管）。
 *
 * 納管項 BU：item.bu 是合法 BU 就用它，否則取負責業務的第一個 BU（與 server.js inferItemBu 同規則）；
 * 秘書另外要求「負責業務的任一 BU 也在秘書的 BU 內」，與秘書的可視範圍（getViewableOwners：同 BU 的業務／主管）一致。
 */
const OR = require('./orgResolve');

const ROLES = Object.freeze(['owner', 'secretary', 'manager1', 'manager2', 'marketing']);
const VALID_BUS = Object.freeze(['ERP', 'ITS', 'MDM', 'CRM', '資安']);

const buList = (v) => {
  const arr = Array.isArray(v) ? v : (typeof v === 'string' ? [v] : []);
  return arr.map((x) => String(x == null ? '' : x).trim()).filter(Boolean);
};

/** 可當提醒收件人：在職、非客戶池、非唯讀掛名主管 */
function usable(u) { return !!u && OR.isActiveUser(u) && !u.pureSupervisor; }

/** 納管項的 BU（單一值，可能為 ''） */
function itemBu(item, users) {
  const map = OR.toMap(users);
  if (item && typeof item.bu === 'string' && VALID_BUS.includes(item.bu)) return item.bu;
  const ow = map[item && item.owner];
  return buList(ow && ow.bu)[0] || '';
}

/** 管理員（在職、非掛名）帳號，依帳號排序 */
function adminUsers(users) {
  return Object.values(OR.toMap(users)).filter((u) => usable(u) && u.role === 'admin').map((u) => u.username).sort();
}

function secretariesFor(item, users) {
  const map = OR.toMap(users);
  const bu = itemBu(item, users);
  if (!bu) return [];
  const ow = map[item && item.owner];
  const ownerBus = buList(ow && ow.bu);
  // 負責業務完全查無 BU 資料時只看納管項 BU；有資料時要求有交集（與秘書可視範圍一致）
  return Object.values(map).filter((u) => usable(u) && u.role === 'secretary')
    .filter((u) => { const sb = buList(u.bu); return sb.includes(bu) && (!ownerBus.length || ownerBus.some((b) => sb.includes(b))); })
    .map((u) => u.username).sort();
}

/**
 * 解析一個納管項的各角色收件人。回
 *   { byRole:{ owner:[], secretary:[], manager1:[], manager2:[], marketing:[] }, issues:[{ role, code }], fallbackAdmin:boolean }
 * issues.code：OWNER_FALLBACK（改取一級主管）｜NO_OWNER_RECIPIENT｜NO_SECRETARY｜NO_MANAGER1（後三者代表已退到管理員；若連管理員都沒有，code 後面加 _NO_ADMIN 的對應項 role 為空）。
 */
function resolveByRole(item, users) {
  const map = OR.toMap(users);
  const admins = adminUsers(map);
  const issues = [];
  const byRole = { owner: [], secretary: [], manager1: [], manager2: [], marketing: [] };
  const ownerName = item && item.owner;
  const ow = map[ownerName];

  // owner
  if (usable(ow)) byRole.owner = [ow.username];
  else {
    const m1 = OR.findManager1(map, ownerName);
    if (m1) { byRole.owner = [m1]; issues.push({ role: 'owner', code: 'OWNER_FALLBACK' }); }
    else { byRole.owner = admins.slice(); issues.push({ role: 'owner', code: 'NO_OWNER_RECIPIENT' }); }
  }
  // secretary
  const secs = secretariesFor(item, map);
  if (secs.length) byRole.secretary = secs;
  else { byRole.secretary = admins.slice(); issues.push({ role: 'secretary', code: 'NO_SECRETARY' }); }
  // manager1
  const m1 = OR.findManager1(map, ownerName);
  if (m1) byRole.manager1 = [m1];
  else { byRole.manager1 = admins.slice(); issues.push({ role: 'manager1', code: 'NO_MANAGER1' }); }
  // manager2：直屬主管且 role 為 manager2
  const sup = ow && map[ow.supervisor];
  if (sup && usable(sup) && sup.role === 'manager2') byRole.manager2 = [sup.username];
  // marketing
  byRole.marketing = OR.marketingUsers(map).filter((n) => usable(map[n]));
  return { byRole, issues, fallbackAdmin: issues.length > 0 };
}

/**
 * 依規則的 notify 名單展開成「人」：回 { recipients:[{ username, roles:[…] }], issues }。
 * 同一個人（例如同時是 owner 與 manager1）只出現一次，roles 合併；依帳號排序以求輸出穩定。
 * issues 只列「這組 notify 內有用到的角色」的問題。
 */
function resolveForNotify(item, users, notify) {
  return expandNotify(resolveByRole(item, users), notify);
}

/** resolveByRole 的結果（可重複用於同一個納管項的多條規則）→ 依 notify 展開成人 */
function expandNotify(r, notify) {
  const want = (Array.isArray(notify) ? notify : []).filter((x) => ROLES.includes(x));
  const acc = new Map();
  want.forEach((role) => {
    r.byRole[role].forEach((u) => {
      if (!acc.has(u)) acc.set(u, new Set());
      acc.get(u).add(role);
    });
  });
  const recipients = [...acc.keys()].sort().map((username) => ({ username, roles: ROLES.filter((x) => acc.get(username).has(x)) }));
  return { recipients, issues: r.issues.filter((i) => want.includes(i.role)), byRole: r.byRole };
}

module.exports = { ROLES, VALID_BUS, usable, itemBu, adminUsers, secretariesFor, resolveByRole, resolveForNotify, expandNotify };
