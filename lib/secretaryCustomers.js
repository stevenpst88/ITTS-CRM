'use strict';
/**
 * 秘書唯讀客戶視圖（SECRETARY_CUSTOMERS=1 才啟用，預設關閉）。
 *
 * 動機：秘書要協助業務辦理寄信，需要看到「同 BU、已成交」客戶的郵寄地址；但秘書原本刻意看不到任何聯絡人／潛客。
 * 做法：不碰 getViewableOwners／GET /api/contacts，另開專用端點；回應一律用「白名單」組新物件（不是黑名單刪欄位），
 *       電話／手機／Email／名片圖／個人偏好／備註等其他欄位一律不輸出，回 piiMasked:true，前端顯示「（個資已遮蔽）」。
 *
 * 本模組是純函式（不讀檔、不碰 Express），server.js 負責取資料與寫稽核；方便單元測試與前後端對拍。
 */

const VALID_BUS = ['ERP', 'ITS', 'MDM', 'CRM', '資安'];
const SALES_ROLES = ['user', 'manager1', 'manager2'];   // 與秘書可看的商機 owner 範圍一致（getViewableOwners 的 secretary/opportunities 分支）
const POOL_USERNAME = '_pool';
const FEATURE_KEY = 'secretaryCustomers';

/** 白名單：只有這些鍵會出現在回應裡（測試用它檢查「多出任何鍵 = 失敗」） */
const VIEW_KEYS = ['id', 'company', 'name', 'title', 'bu', 'ownerDisplayName', 'mailingAddress', 'mailingSameAsCompany', 'mailingFallback', 'piiMasked'];

function enabled(env) { return String((env || process.env).SECRETARY_CUSTOMERS || '').trim() === '1'; }

function normalizeBu(bu) {
  if (Array.isArray(bu)) return bu.filter(Boolean);
  if (typeof bu === 'string' && bu.trim()) return [bu.trim()];
  return [];
}

/**
 * 規則 3 的單一來源：有任一張名片 customerType==='customer' 的「公司名稱字串」集合（空字串不算）。
 * norm 預設原樣比對（與前端 getWonContactIds 完全相同；秘書視圖用這個）；企業主檔 TA 傳 trim 版本拿來對應主檔名稱。
 */
function customerCompanyNames(contacts, norm) {
  const f = norm || ((v) => v);
  const names = new Set();
  (contacts || []).forEach((c) => { if (c && c.customerType === 'customer' && c.company) { const k = f(c.company); if (k) names.add(k); } });
  return names;
}

/**
 * 「已成交」客戶 id 集合——與前端 app.js 的 getWonContactIds() 同一規則（scripts/check-secretary-customers.js 對拍）：
 *   1. 有 stage==='Won' 的商機 → 其 contactId
 *   2. customerType==='customer' 的名片
 *   3. 以公司為單位：同公司（company 字串完全相同、非空）任一人是 customer，整間公司都算
 * contacts／opportunities 是「呼叫端已收斂範圍」的清單（前端是使用者看得到的清單；後端是秘書同 BU 範圍）。
 * 與前端唯一差異：後端不放 falsy 的 id（前端的 Set 會含 undefined，沒有意義）。
 */
function wonContactIds(contacts, opportunities) {
  const ids = new Set();
  (opportunities || []).forEach((o) => { if (o && o.stage === 'Won' && o.contactId) ids.add(o.contactId); });
  const customerCompanies = customerCompanyNames(contacts);
  (contacts || []).forEach((c) => {
    if (c && (c.customerType === 'customer' || (c.company && customerCompanies.has(c.company)))) ids.add(c.id);
  });
  return ids;
}

/** 名片的「資料 BU」：優先 contact.bu 內的有效值；沒有就回退擁有者的第一個 BU（與 server.js inferItemBu 同概念） */
function contactBus(c, ownerUser) {
  const own = normalizeBu(c && c.bu).filter((b) => VALID_BUS.includes(b));
  if (own.length) return own;
  const ob = normalizeBu(ownerUser && ownerUser.bu)[0];
  return ob ? [ob] : [];
}

/**
 * 秘書範圍內的名片：未刪除、非客戶池／待補名片、擁有者是同 BU 的業務／一二級主管、名片資料 BU 與秘書 BU 有交集。
 * 秘書沒有 BU → 空（fail-safe）。
 */
function scopeContacts(data, users, me) {
  const myBus = normalizeBu(me && me.bu);
  if (!myBus.length) return [];
  const byName = new Map((users || []).map((u) => [u.username, u]));
  return (data.contacts || []).filter((c) => {
    if (!c || c.deleted || c.isPlaceholder) return false;
    if (!c.owner || c.owner === POOL_USERNAME) return false;
    const ou = byName.get(c.owner);
    if (!ou || !SALES_ROLES.includes(ou.role)) return false;
    return contactBus(c, ou).some((b) => myBus.includes(b));
  });
}

/** 單筆白名單輸出。ownerLabel 由呼叫端用 userLabel() 算好 */
function viewOf(c, ownerLabel) {
  const mail = typeof c.mailingAddress === 'string' ? c.mailingAddress.trim() : '';
  const addr = typeof c.address === 'string' ? c.address.trim() : '';
  const fallback = !mail;
  return {
    id: c.id,
    company: String(c.company || ''),
    name: String(c.name || ''),
    title: String(c.title || ''),
    bu: contactBus(c, null),
    ownerDisplayName: String(ownerLabel || ''),
    mailingAddress: fallback ? addr : mail,
    mailingSameAsCompany: c.mailingSameAsCompany === true,
    mailingFallback: fallback && !!addr,
    piiMasked: true,
  };
}

/**
 * 清單：範圍內 ∩ 已成交。opps 只取「指向範圍內名片」的商機（不會因其他 BU 的資料把本 BU 名片判成成交）。
 * labelOf(username) 回擁有者顯示名。
 */
function listFor(data, users, me, labelOf) {
  const scoped = scopeContacts(data, users, me);
  const ids = new Set(scoped.map((c) => c.id));
  const opps = (data.opportunities || []).filter((o) => o && ids.has(o.contactId));
  const won = wonContactIds(scoped, opps);
  const byName = new Map((users || []).map((u) => [u.username, u]));
  return scoped.filter((c) => won.has(c.id)).map((c) => {
    const v = viewOf(c, labelOf ? labelOf(c.owner) : c.owner);
    v.bu = contactBus(c, byName.get(c.owner));
    return v;
  }).sort((a, b) => (a.company || '').localeCompare(b.company || '', 'zh-TW') || (a.name || '').localeCompare(b.name || '', 'zh-TW') || String(a.id).localeCompare(String(b.id)));
}

module.exports = { FEATURE_KEY, VIEW_KEYS, VALID_BUS, SALES_ROLES, enabled, normalizeBu, wonContactIds, customerCompanyNames, contactBus, scopeContacts, viewOf, listFor };
