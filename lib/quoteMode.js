'use strict';
/**
 * 報價單權限類型（quoteMode）：簽核／唯讀 二擇一（職能分工，內控 2026-10）。
 *
 * 純函式、無 I/O。伺服器端是唯一的執法點（lib/quoteRoutes.js 的 canApproveAct 等）；畫面只是輔助。
 *
 * 適用對象＝「不能建立報價單」的帳號：
 *   NO_CREATE_ROLES 的角色（pool 系統帳號除外），以及列在簽核名冊 gm／chairman 的人（他們一樣不能建單）。
 *   建單角色（user／manager2／自訂顧問角色…）與 admin（後台維運，不建單也不簽核）不適用，也不存這個欄位。
 * 值：
 *   'approver' 簽核：依既有規則可被指派為簽核人、進名冊、收簽核通知／信件。
 *   'readonly' 唯讀：只能在既有可視規則內「看」；任何簽核動作（核准／駁回／代核／章管理）一律擋，
 *              也不會被當成簽核人（關卡指派、名冊、候選人、待辦、通知、信件都排除）。
 *   未設定（欄位不存在）＝舊帳號：行為與以前完全相同（不自動遷移），只在後台標「未設定」並列入報表。
 */

/** 不建立報價單的角色（規格：業務發起；主管、總經理、董事長、秘書等不建單）。lib/quoteRoutes.js 引用同一份。 */
const NO_CREATE_ROLES = ['executive', 'manager1', 'secretary', 'marketing', 'tecopm', 'groupsales', 'pool', 'accounting_manager', 'finance_manager', 'consult_manager_south', 'consult_manager_north'];
/** 適用 quoteMode 的角色（NO_CREATE_ROLES 去掉 pool 系統帳號） */
const NON_CREATOR_ROLES = NO_CREATE_ROLES.filter((r) => r !== 'pool');
const MODES = ['approver', 'readonly'];
const MODE_LABEL = { approver: '簽核', readonly: '唯讀' };

/** 取帳號目前有效的 quoteMode；沒設定或不合法一律回 ''（＝未設定） */
function modeOf(u) {
  return u && MODES.includes(u.quoteMode) ? u.quoteMode : '';
}
const isReadonly = (u) => modeOf(u) === 'readonly';

/**
 * 解析送入的 quoteMode。回 { ok:true, present:boolean, value:''|'approver'|'readonly' }
 *   欄位沒出現（undefined）→ present:false；null／空字串 → present:true,value:''（＝清除）；
 *   其他非 'approver'|'readonly' 的值（含非字串）→ { ok:false }。
 */
function parseInput(raw) {
  if (raw === undefined) return { ok: true, present: false, value: '' };
  if (raw === null || raw === '') return { ok: true, present: true, value: '' };
  if (typeof raw === 'string' && MODES.includes(raw)) return { ok: true, present: true, value: raw };
  return { ok: false };
}

/** 角色是否為「非建單」角色 */
const isNonCreatorRole = (role) => NON_CREATOR_ROLES.includes(role);

/** 是否列在簽核名冊的總經理／董事長（roster 可為 undefined） */
function onApprovalRoster(username, roster) {
  const r = roster || {};
  return (Array.isArray(r.gm) && r.gm.includes(username)) || (Array.isArray(r.chairman) && r.chairman.includes(username));
}

/** 這個帳號是否適用 quoteMode（非建單角色，或列在總經理／董事長名冊；admin 與 pool 不適用） */
function isApplicable(role, username, roster) {
  if (!role || role === 'admin' || role === 'pool') return false;
  return isNonCreatorRole(role) || onApprovalRoster(username, roster);
}

module.exports = { NO_CREATE_ROLES, NON_CREATOR_ROLES, MODES, MODE_LABEL, modeOf, isReadonly, parseInput, isNonCreatorRole, onApprovalRoster, isApplicable };
