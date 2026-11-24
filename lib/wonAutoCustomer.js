'use strict';
/**
 * 商機成交（Won）→ 綁定名片自動轉為「我的客戶」。純邏輯，可單元測試；server.js 負責 HTTP、權限、儲存與稽核。
 *
 * 為什麼：過去成交轉客戶完全靠前端彈窗（被略過、公司名與名片不一字不差、Excel 匯入、Lead 轉換都不會寫 customerType）。
 * 現在由後端在「階段由非 Won 變成 Won」時，依商機綁定的 contactId 把那張名片標為 customer。
 *
 * 規則：
 *   - 只動「商機綁定的那一張名片」（contactId）；同公司其他名片的擴散仍走前端既有規則（不在這裡靜默擴散）。
 *   - 已是 customer 不重寫（冪等）；不覆寫既有 productLine（沒有就維持原樣，不要求產品線）。
 *   - 名片不存在／未綁定／已刪除／待補名片（placeholder）→ 不報錯、不阻擋成交，回報原因讓呼叫端記稽核。
 *   - 不看名片 owner（系統自動行為）；Won 回退不降級。
 * 開關：WON_AUTO_CUSTOMER=0（或 off／false）整個關閉，預設開。
 */

function enabled(env) {
  const v = String((env || process.env).WON_AUTO_CUSTOMER == null ? '' : (env || process.env).WON_AUTO_CUSTOMER).trim().toLowerCase();
  return !(v === '0' || v === 'off' || v === 'false');
}

// 找到商機綁定的有效名片。回傳 { contact } 或 { reason }：no-binding／not-found／deleted／placeholder
function findBoundContact(data, opp) {
  const cid = opp && opp.contactId ? String(opp.contactId) : '';
  if (!cid) return { reason: 'no-binding' };
  const c = (data.contacts || []).find((x) => x && x.id === cid);
  if (!c) return { reason: 'not-found' };
  if (c.deleted) return { reason: 'deleted' };
  if (c.isPlaceholder) return { reason: 'placeholder' };
  return { contact: c };
}

/**
 * 把一筆成交商機對應的名片轉為客戶（就地修改 data）。
 * 回傳 { status: 'converted'|'already'|'skipped', contactId?, reason?, changed: boolean }
 */
function convertForWonOpp(data, opp) {
  const f = findBoundContact(data, opp);
  if (!f.contact) return { status: 'skipped', reason: f.reason, changed: false };
  const c = f.contact;
  if (c.customerType === 'customer') return { status: 'already', contactId: c.id, changed: false };
  const prevType = c.customerType || '';
  c.customerType = 'customer';
  c.updatedAt = new Date().toISOString();
  return { status: 'converted', contactId: c.id, changed: true, prevType };
}

const REASON_TEXT = {
  'no-binding': '商機未綁定名片',
  'not-found': '綁定的名片不存在',
  'deleted': '綁定的名片已刪除',
  'placeholder': '綁定的是待補名片',
};

/**
 * 管理員重新同步（報告用）。掃描所有 Won 商機：
 *   candidates  有綁定有效名片、且名片非 customer → 會（或已）被轉換
 *   unbound     Won 但沒有可用的綁定名片 → 只報告，不改；附「公司名寬鬆比對」找到的疑似名片 id（只報 id，不擴散）
 * apply=true 才寫入 candidates。回傳報告（不含電話／Email 等個資；各清單有筆數上限，total 為真實總數）。
 */
function resync(data, opts) {
  const o = opts || {};
  const apply = o.apply === true;
  const cap = o.cap || 200;
  const looseFn = typeof o.normalizeLoose === 'function' ? o.normalizeLoose : (s) => String(s || '').trim().toLowerCase();
  const contacts = (data.contacts || []).filter((c) => c && !c.deleted && !c.isPlaceholder);
  const looseIdx = new Map();   // 寬鬆公司名 → contact[]
  for (const c of contacts) {
    const k = looseFn(c.company);
    if (!k) continue;
    if (!looseIdx.has(k)) looseIdx.set(k, []);
    looseIdx.get(k).push(c);
  }
  const seen = new Set();
  const candidates = [], unbound = [];
  let candidateTotal = 0, unboundTotal = 0, converted = 0;
  for (const opp of (data.opportunities || [])) {
    if (!opp || opp.stage !== 'Won') continue;
    const f = findBoundContact(data, opp);
    if (f.contact) {
      if (f.contact.customerType === 'customer') continue;
      if (seen.has(f.contact.id)) continue;   // 同一張名片多筆成交只報一次
      seen.add(f.contact.id);
      candidateTotal++;
      if (candidates.length < cap) candidates.push({ company: f.contact.company || opp.company || '', contactId: f.contact.id, oppId: opp.id, reason: '成交商機的綁定名片尚未是客戶' });
      if (apply) { f.contact.customerType = 'customer'; f.contact.updatedAt = new Date().toISOString(); converted++; }
    } else {
      unboundTotal++;
      if (unbound.length < cap) {
        const near = (looseIdx.get(looseFn(opp.company)) || []).filter((c) => c.customerType !== 'customer').slice(0, 5).map((c) => c.id);
        unbound.push({ company: opp.company || '', oppId: opp.id, reason: REASON_TEXT[f.reason] || f.reason, looseMatchContactIds: near });
      }
    }
  }
  return { apply, cap, candidates, candidateTotal, unbound, unboundTotal, converted, truncated: candidateTotal > candidates.length || unboundTotal > unbound.length };
}

module.exports = { enabled, findBoundContact, convertForWonOpp, resync, REASON_TEXT };
