'use strict';
/**
 * lib/mail/visibility.js — 「哪一種收件人在信裡看得到什麼」的唯一事實來源
 *
 * 純資料＋純函式。renderer／樣板一律查 visibilityFor(kind)，不得在樣板內寫死任何角色判斷
 * （寫死的判斷會與這張表漂移，正是資料外洩的來源）。
 *
 * 匯出與簽名：
 *   KINDS                          = ['mgr1','gm','chairman','secretary','boardProxy','consultant','owner']   （凍結）
 *   visibilityFor(kind)            → {amount, margin, tier, customer, owner, items, project, itemPrices:false, reason}  （每次回傳新的凍結物件）
 *   isKnownKind(kind)              → boolean
 *   FIELDS                         = ['amount','margin','tier','customer','owner','items','project']
 *
 * 政策表（2026-10-08 業主決定；true＝信裡可以出現）：
 *   kind                     amount margin tier customer owner items project
 *   mgr1／gm／chairman       true   true   true true     true  false true    一級主管、總經理、董事長：金額＋毛利率＋層級＋客戶名＋業務＋專案名稱；信裡不放品項明細
 *   secretary／boardProxy    true   true   true false    false false false   秘書（董事會關）與董事會代核人：沿用站內通知，不帶客戶名、業務名；專案名稱常含客戶名，所以也不帶，主旨／preheader／本文只放報價單號
 *   consultant               false  false  false false   true  true  true    顧問：不含任何金額／毛利率／折扣／單價；可看業務是誰（找誰）、專案名稱與品項說明／數量／單位
 *   owner                    false  false  false true    false false true    業務本人（E4/E5/E6 結果通知）：只放結果／原因／連結／專案名稱，不放金額毛利
 *   未知 kind                false  false  false false   false false false   最小資料原則（大小寫不同、非字串、'__proto__' 之類一律視為未知）
 *   project＝專案名稱（主旨、preheader、本文的「專案名稱」列）。業主 2026-10-08 決定：寄給秘書與董事會代核人的信不顯示專案名稱，只放報價單號。
 *   itemPrices 恆為 false：任何收件人的信都不帶品項單價。
 *
 * 注意：這裡的 owner 欄位是「是否顯示業務姓名」，kind 'owner' 是「收件人就是業務本人」，兩者不同。
 */

const FIELDS = Object.freeze(['amount', 'margin', 'tier', 'customer', 'owner', 'items', 'project']);
const KINDS = Object.freeze(['mgr1', 'gm', 'chairman', 'secretary', 'boardProxy', 'consultant', 'owner']);

// 逐列明寫（不共用物件），避免改一處牽動多處；測試以獨立抄錄的矩陣逐格比對。
const POLICY = new Map([
  ['mgr1',       { amount: true,  margin: true,  tier: true,  customer: true,  owner: true,  items: false, project: true,  reason: '一級主管：簽核人，可見金額、毛利率、核決層級、客戶、業務與專案名稱' }],
  ['gm',         { amount: true,  margin: true,  tier: true,  customer: true,  owner: true,  items: false, project: true,  reason: '總經理：簽核人，可見金額、毛利率、核決層級、客戶、業務與專案名稱' }],
  ['chairman',   { amount: true,  margin: true,  tier: true,  customer: true,  owner: true,  items: false, project: true,  reason: '董事長：簽核人，可見金額、毛利率、核決層級、客戶、業務與專案名稱' }],
  ['secretary',  { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: false, reason: '秘書（董事會關）：可見金額與毛利率，不帶客戶名、業務名與專案名稱（只放單號）' }],
  ['boardProxy', { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: false, reason: '董事會代核人：可見金額與毛利率，不帶客戶名、業務名與專案名稱（只放單號）' }],
  ['consultant', { amount: false, margin: false, tier: false, customer: false, owner: true,  items: true,  project: true,  reason: '顧問：不含任何金額、毛利率、折扣、單價；可見業務、專案名稱與品項說明、數量、單位' }],
  ['owner',      { amount: false, margin: false, tier: false, customer: true,  owner: false, items: false, project: true,  reason: '業務本人：只看結果、原因、專案名稱與連結，不放金額與毛利' }],
]);

const UNKNOWN_REASON = '未知收件人類型：採最小資料原則，全部不顯示';

function isKnownKind(kind) {
  return typeof kind === 'string' && POLICY.has(kind);
}

/** 查政策。回傳新的凍結物件（呼叫端改不到政策表）。未知 kind → 全 false。itemPrices 恆為 false。 */
function visibilityFor(kind) {
  const row = typeof kind === 'string' ? POLICY.get(kind) : undefined;
  if (!row) {
    return Object.freeze({
      amount: false, margin: false, tier: false, customer: false, owner: false, items: false, project: false,
      itemPrices: false, reason: UNKNOWN_REASON,
    });
  }
  return Object.freeze({
    amount: row.amount,
    margin: row.margin,
    tier: row.tier,
    customer: row.customer,
    owner: row.owner,
    items: row.items,
    project: row.project,
    itemPrices: false,
    reason: row.reason,
  });
}

module.exports = {
  KINDS,
  FIELDS,
  visibilityFor,
  isKnownKind,
};
