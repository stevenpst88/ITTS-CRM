'use strict';
/**
 * lib/mail/recipients.js — 把一串 username 解析成「可以寄信的真實收件人」
 *
 * 匯出與簽名：
 *   resolveRecipients(usernames, {users, actorUsername, config})
 *     → {deliver:[{username, email, label}], skipped:[{username, reason}]}
 *   SKIP_REASONS = DUP | ACTOR | UNKNOWN_USER | INACTIVE | NO_EMAIL | BAD_EMAIL | DOMAIN_NOT_ALLOWED
 *
 * 規則：
 *  - username 比對「精確、區分大小寫」。依據：登入與全站帳號查詢都是 `u.username === username`（server.js 登入路由與各處 find），
 *    而且實際帳號資料裡存在「只差大小寫、卻是不同人」的帳號，所以這裡絕不能 case-fold（會把兩個人當成同一個）。
 *    查表用 Map（不是物件），'__proto__'、'constructor' 之類的帳號名稱不會誤中原型屬性。
 *  - users 可以是陣列（每個 {username, email?, active?, disabled?, displayName?, nickname?, role}），
 *    也可以是以 username 為鍵的物件（lib/quoteRoutes.js 的 ctx.users 就是這種 userMap）。
 *  - 停用：active===false、disabled===true，或系統帳號 role==='pool'（客戶池，非真人；對照 quoteRoutes.js 的 isActive）。
 *  - 順序保留；同一 username 只處理第一次，之後記 DUP；操作者本人記 ACTOR；之後依序檢查
 *    UNKNOWN_USER → INACTIVE → NO_EMAIL → BAD_EMAIL → DOMAIN_NOT_ALLOWED。
 *  - email 一律經 normalizeEmail（小寫、單一位址、嚴格字元）＋isAllowedDomain（網域精確相等）。
 *    deliver[].email 是正規化後的「真實」位址；log／redirect／off 模式的改寫由 transport 層負責，本層不改。
 *  - label＝信內稱呼：暱稱 > 顯示名稱 > 帳號（與 server.js userLabel、quoteRoutes.js dispName 同序），經 safeText（≤60 字）。
 *  - skipped[].username 已經 safeText（會寫進稽核與畫面，可能來自不受信任的輸入）。
 *  - 永不 throw：users／usernames 型別不對就當成空。
 */

const { normalizeEmail, isAllowedDomain, safeText } = require('./safety');
const { DEFAULTS } = require('./config');

const SKIP_REASONS = Object.freeze({
  DUP: 'DUP',
  ACTOR: 'ACTOR',
  UNKNOWN_USER: 'UNKNOWN_USER',
  INACTIVE: 'INACTIVE',
  NO_EMAIL: 'NO_EMAIL',
  BAD_EMAIL: 'BAD_EMAIL',
  DOMAIN_NOT_ALLOWED: 'DOMAIN_NOT_ALLOWED',
});

const MAX_LIST = 1000;
const LABEL_MAX = 60;
const FALLBACK_LABEL = '同仁';

function isObj(v) { return v !== null && typeof v === 'object'; }

function indexUsers(users) {
  const map = new Map();
  if (Array.isArray(users)) {
    for (let i = 0; i < users.length; i++) {
      const u = users[i];
      if (isObj(u) && typeof u.username === 'string' && u.username !== '' && !map.has(u.username)) map.set(u.username, u);
    }
  } else if (isObj(users)) {
    const keys = Object.keys(users);
    for (let i = 0; i < keys.length; i++) {
      const u = users[keys[i]];
      if (!isObj(u)) continue;
      const un = (typeof u.username === 'string' && u.username !== '') ? u.username : keys[i];
      if (!map.has(un)) map.set(un, u);
    }
  }
  return map;
}

function isInactive(u) {
  return u.active === false || u.disabled === true || u.role === 'pool';
}

function pickLabel(u, username) {
  const cands = [u.nickname, u.displayName, username];
  for (let i = 0; i < cands.length; i++) {
    const c = cands[i];
    if (typeof c === 'string' && c.trim() !== '') {
      const t = safeText(c, LABEL_MAX);
      if (t) return t;
    }
  }
  return FALLBACK_LABEL;
}

function resolveRecipients(usernames, opts) {
  const o = isObj(opts) ? opts : {};
  const index = indexUsers(o.users);
  const actor = typeof o.actorUsername === 'string' ? o.actorUsername : '';
  const allowed = (isObj(o.config) && Array.isArray(o.config.allowedDomains)) ? o.config.allowedDomains : DEFAULTS.allowedDomains;
  const list = Array.isArray(usernames) ? usernames : [];
  const deliver = [];
  const skipped = [];
  const seen = new Set();
  const skip = (name, reason) => { skipped.push({ username: safeText(name, 64), reason }); };

  const n = Math.min(list.length, MAX_LIST);
  for (let i = 0; i < n; i++) {
    const un = list[i];
    if (typeof un !== 'string' || un === '') { skip(un, SKIP_REASONS.UNKNOWN_USER); continue; }
    if (seen.has(un)) { skip(un, SKIP_REASONS.DUP); continue; }
    seen.add(un);
    if (actor !== '' && un === actor) { skip(un, SKIP_REASONS.ACTOR); continue; }
    const u = index.get(un);
    if (!u) { skip(un, SKIP_REASONS.UNKNOWN_USER); continue; }
    if (isInactive(u)) { skip(un, SKIP_REASONS.INACTIVE); continue; }
    const rawEmail = u.email;
    if (rawEmail === undefined || rawEmail === null || (typeof rawEmail === 'string' && rawEmail.trim() === '')) {
      skip(un, SKIP_REASONS.NO_EMAIL);
      continue;
    }
    const norm = normalizeEmail(rawEmail);
    if (!norm.ok) { skip(un, SKIP_REASONS.BAD_EMAIL); continue; }
    if (!isAllowedDomain(norm.value, allowed)) { skip(un, SKIP_REASONS.DOMAIN_NOT_ALLOWED); continue; }
    deliver.push({ username: un, email: norm.value, label: pickLabel(u, un) });
  }
  return { deliver, skipped };
}

module.exports = {
  resolveRecipients,
  SKIP_REASONS,
};
