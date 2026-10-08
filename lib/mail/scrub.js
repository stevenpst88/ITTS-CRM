'use strict';
/**
 * lib/mail/scrub.js — 錯誤訊息清理（純函式、無 I/O、永不 throw）
 *
 * 匯出與簽名：
 *   scrubMessage(v, known?) → string
 *     v      任何值（轉成字串）；超過 2000 字先硬切
 *     known  額外要遮蔽的字串陣列（例如主旨、專案名稱、完整 email；長度 ≥3 才處理，最多 20 個）
 *   結果：單行、最多 200 字，且已移除密鑰樣式字串：
 *     Bearer 權杖、JWT、client_secret／password／token／authorization 之類的「名稱=值」、
 *     email（可能是收件人）、GUID（租戶／應用程式 ID）、32 字元以上的不透明長字串（金鑰、權杖）、
 *     獨立的 7 位數以上數字與千分位數字（可能是金額；AADSTS7000215 這類夾在字母間的錯誤代碼會保留）。
 *
 * 為什麼需要：寄信失敗的錯誤訊息會寫進 outbox（後台可看）與稽核，而 Graph／OAuth 的錯誤文字常會回顯租戶 ID、
 * 收件人、甚至部分憑證內容。公開 repo + 多人可見的後台，寧可多遮不要漏。所有正規式都是線性（沒有巢狀量詞），
 * 輸入先限制在 2000 字，所以 200 萬字元的輸入也不會拖垮。
 */

const { safeText } = require('./safety');

const RE_BEARER = /\bBearer\s+[A-Za-z0-9._~+\/=-]{8,}/gi;
const RE_JWT = /\beyJ[A-Za-z0-9_-]{6,}\.[A-Za-z0-9_-]{6,}(?:\.[A-Za-z0-9_-]*)?/g;
const RE_KV_SECRET = /\b(client[_-]?secret|secret|password|passwd|pwd|access[_-]?token|refresh[_-]?token|id[_-]?token|token|api[_-]?key|authorization)\b["']?\s*[:=]\s*(?:(?:Basic|Bearer|Digest|NTLM|Negotiate)\s+)?(?:"[^"]*"|'[^']*'|\S+)/gi;
const RE_EMAIL_LIKE = /[A-Za-z0-9._%+-]{1,64}@[A-Za-z0-9.-]{1,253}\.[A-Za-z]{2,}/g;
const RE_GUID = /\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b/g;
const RE_OPAQUE = /[A-Za-z0-9+\/_=~.-]{32,}/g;
// 金額樣式的數字（錯誤訊息若回顯信件內容，金額不該跟著進 outbox／稽核）：獨立的 7 位數以上數字、或有千分位的數字。
// 夾在字母數字之間的不處理（保留 AADSTS7000215 這類錯誤代碼，它們對除錯很重要）
const RE_BIGNUM = /(?<![A-Za-z0-9])(?:\d{1,3}(?:,\d{3}){2,}|\d{7,})(?![A-Za-z0-9])/g;

function scrubMessage(v, known) {
  let s;
  try { s = typeof v === 'string' ? v : (v === null || v === undefined ? '' : String(v)); } catch (e) { s = ''; }
  if (s.length > 2000) s = s.slice(0, 2000);
  if (Array.isArray(known)) {
    for (let i = 0; i < known.length && i < 20; i++) {
      const k = known[i];
      if (typeof k === 'string' && k.length >= 3 && k.length <= 2000) s = s.split(k).join('[已遮蔽]');
    }
  }
  s = s.replace(RE_BEARER, 'Bearer [已遮蔽]')
    .replace(RE_JWT, '[已遮蔽]')
    .replace(RE_KV_SECRET, '$1=[已遮蔽]')
    .replace(RE_EMAIL_LIKE, '[email]')
    .replace(RE_GUID, '[id]')
    .replace(RE_OPAQUE, '[已遮蔽]')
    .replace(RE_BIGNUM, '[數字]');
  return safeText(s, 200);
}

module.exports = { scrubMessage };
