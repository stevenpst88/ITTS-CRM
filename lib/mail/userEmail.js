'use strict';
/**
 * lib/mail/userEmail.js — 使用者 Email 欄位的純邏輯（P1：無 UI、無路由、無 I/O，不讀不寫任何資料檔）
 *
 * 匯出與簽名：
 *   validateUserEmail(raw, {config, users, selfUsername})
 *       → {ok:true, value, code:'OK'|'CLEARED'} | {ok:false, error, code, conflictUsername?}
 *   parseBulkEmailText(text, {config, users})
 *       → {rows:[{line, username, email, status:'ok'|'unchanged'|'error', code?, error?}], summary:{ok, unchanged, error}, fatal:boolean}
 *   applyBulkPlan(rows, users, {config}?)
 *       → {updated:[{username, previous, email, detail}], skipped:[{username, code}], users:[...新陣列], error?}
 *   missingEmailReport({users, roster, config})
 *       → [{username, label, roles:[roleKey], roleLabels:[中文], reason:'NO_EMAIL'|'BAD_EMAIL'|'DOMAIN_NOT_ALLOWED', readOnly}]
 *   maskedAuditDetail(oldEmail, newEmail) → 'email 未設定→已設定' | 'email 已變更' | 'email 已清除' | 'email 無變更'
 *   ROLE_KEYS, ROLE_LABELS, BULK_LIMITS
 *
 * 規則（每一條都有 scripts/check-mail-core.js「6 userEmail」章節的測試）：
 *  1. Email 格式與網域白名單一律走 safety.normalizeEmail／isAllowedDomain（不另寫一套）。
 *  2. 全站唯一，比對「不分大小寫」（用正規化後的小寫位址比）；停用帳號持有的位址也算占用；
 *     已存的值若根本不是合法位址（髒資料）不參與比對。排除 selfUsername 本人。
 *     注意：username 比對仍是「精確、區分大小寫」（server.js 登入與全站查詢都是 ===，而且實際資料存在只差大小寫的不同帳號），
 *     所以 selfUsername、批次匯入的帳號欄位都不可轉小寫。
 *  3. 清除：null、空字串、全空白 → {ok:true, value:'', code:'CLEARED'}。undefined 視為「沒傳這個欄位」→ 錯誤 MISSING，
 *     避免路由漏帶欄位時無聲把所有人的 Email 清掉。
 *  4. 錯誤訊息一律是固定文字，不回顯輸入內容（避免反射式 XSS）；衝突帳號放在 conflictUsername 欄位（已 safeText），由呼叫端決定要不要顯示。
 *  5. 批次匯入（只預覽，不寫入）：
 *     - 每行「帳號 分隔 Email」，分隔可為半形／全形逗號、分號、Tab、空白（含全形空白）。Email 一律是「最後一欄」，
 *       因此帳號本身含空白也能解析（例如「Mary Lee<Tab>mary@…」）。欄位外層的雙引號（含彎引號）會被去掉（Excel 轉出的 CSV）。
 *     - 忽略空行與 # 開頭的註解行（因此以 # 開頭的帳號無法用批次匯入）；去 BOM；CRLF／LF／CR 皆可；行號是原檔的實體行號（從 1 起算）。
 *     - 上限：資料行 500 行、全文 200000 字元、單行 1000 字元。超過 500 行或全文過大 → 整批拒絕（fatal，rows 只有一筆 line=0 的錯誤），
 *       不做「只處理前 500 行」的半套匯入。單行過長只讓該行出錯。
 *     - 空的 Email 欄位視為格式錯誤（批次匯入不支援清除，避免一份表格不小心清掉一堆人；要清除請個別編輯）。
 *     - 帳號不存在 → 錯誤 UNKNOWN_USER（若只差大小寫會提示「帳號區分大小寫」，不洩漏是哪個帳號）。
 *     - 同一批內：同一個 Email 出現在多個帳號、或同一帳號出現多次 → 涉及的每一行都標錯誤（不猜哪一行才對）。
 *       與「目前其他帳號已用的 Email」衝突也是錯誤（所以兩人互換 Email 要分兩次做）。
 *     - rows[].username：帳號存在時是「原樣的帳號名」（applyBulkPlan 要靠它精確比對）；不存在時才是 safeText 過的輸入。
 *       兩者都可能含 < > 等字元，UI 端一律要跳脫（textContent／escHtml）。
 *  6. applyBulkPlan 不信任傳進來的 rows（預覽與確認可能是兩個請求）：只處理 status==='ok' 的行，並對每一行重新驗證
 *     （帳號存在、格式、網域、與「演進中的名冊」唯一）；不合格的進 skipped。回傳新陣列且每個 user 都是淺拷貝，不就地修改輸入。
 *     只接受陣列形態的 users（auth.json 的形態）；其他形態回傳原物件並附 error，不做任何變更（寧可不動，也不要回傳會被誤存的空值）。
 *     只改 user.email，不動其他欄位。
 *  7. missingEmailReport 的角色對照（出處：lib/quoteRoutes.js getCfg／boardProxySet、lib/quoteApproval 的 tier 定義、codemap §2）：
 *       manager1      ← 帳號 role === 'manager1'                  （一級主管；實際簽核人是沿 supervisor 鏈找到的那一位，這裡不推導，整批列出）
 *       gm            ← roster.gm                                  （總經理名冊，名單內每個人都會收到通知）
 *       chairman      ← roster.chairman                            （董事長名冊）
 *       boardProxy    ← roster.boardProxy                          （董事會代核名冊）
 *       secretary     ← 帳號 role === 'secretary'                  （董事會代核人 = boardProxy ∪ 所有秘書）
 *       costProvider  ← roster.costProviders                       （成本填寫人／支援顧問，E2 的收件人）
 *     不含 roster.sealManagers（報價章管理人不會收到簽核信）。停用帳號（active===false、disabled===true、role==='pool'）不列；
 *     名冊裡不存在的帳號不列（本函式回報的是「帳號缺 Email」，名冊錯字是另一個問題）。
 *     判定缺什麼的規則與 recipients.resolveRecipients 完全相同（NO_EMAIL／BAD_EMAIL／DOMAIN_NOT_ALLOWED），測試有對拍。
 *     readOnly＝accessMode==='view' 的唯讀帳號（quoteRoutes canAct 會排除他們；仍列出以免漏掉，UI 可自行淡化）。
 *     結果依帳號（code unit 順序）排序、每個帳號只出現一次。
 *  8. maskedAuditDetail 的輸出只會是上面四句固定文字之一，不可能含位址。
 *
 * 限制：label 與停用判斷各有一份與 recipients.js 同規則的小型複本（不改動基礎模組），由測試對拍以免漂移。
 */

const { normalizeEmail, isAllowedDomain, safeText } = require('./safety');
const { DEFAULTS } = require('./config');

const MAX_USERS = 20000;
const BULK_LIMITS = Object.freeze({ maxRows: 500, maxTextChars: 200000, maxLineChars: 1000 });

const ROLE_KEYS = Object.freeze(['manager1', 'gm', 'chairman', 'boardProxy', 'secretary', 'costProvider']);
const ROLE_LABELS = Object.freeze({
  manager1: '一級主管',
  gm: '總經理',
  chairman: '董事長',
  boardProxy: '董事會代核人',
  secretary: '秘書',
  costProvider: '成本填寫人',
});
// roster 欄位 → 角色鍵
const ROSTER_FIELDS = Object.freeze({ gm: 'gm', chairman: 'chairman', boardProxy: 'boardProxy', costProviders: 'costProvider' });

const AUDIT_TEXT = Object.freeze({
  SET: 'email 未設定→已設定',
  CHANGED: 'email 已變更',
  CLEARED: 'email 已清除',
  SAME: 'email 無變更',
});

const FW_COMMA = String.fromCodePoint(0xff0c);   // ，
const FW_SEMI = String.fromCodePoint(0xff1b);    // ；
const CURLY_OPEN = String.fromCodePoint(0x201c);
const CURLY_CLOSE = String.fromCodePoint(0x201d);
const RE_WS = /\s/;                               // 含 NBSP、全形空白(U+3000)、BOM(U+FEFF)

// ── 小工具 ──────────────────────────────────────────────────────────────────
function isObj(v) { return v !== null && typeof v === 'object'; }

function allowedOf(config) {
  return (isObj(config) && Array.isArray(config.allowedDomains)) ? config.allowedDomains : DEFAULTS.allowedDomains;
}

/** users（陣列，或以 username 為鍵的物件）→ [{username, user}]；保留順序，重複帳號只收第一個（同 recipients.indexUsers）。 */
function userEntries(users) {
  const out = [];
  const seen = new Set();
  const add = (username, u) => { if (!seen.has(username)) { seen.add(username); out.push({ username, user: u }); } };
  if (Array.isArray(users)) {
    const n = Math.min(users.length, MAX_USERS);
    for (let i = 0; i < n; i++) {
      const u = users[i];
      if (isObj(u) && typeof u.username === 'string' && u.username !== '') add(u.username, u);
    }
  } else if (isObj(users)) {
    const keys = Object.keys(users);
    const n = Math.min(keys.length, MAX_USERS);
    for (let i = 0; i < n; i++) {
      const u = users[keys[i]];
      if (!isObj(u)) continue;
      add((typeof u.username === 'string' && u.username !== '') ? u.username : keys[i], u);
    }
  }
  return out;
}

/** 現有 email 的正規化小寫位址；沒有或根本不合法 → ''（髒資料不參與唯一性比對）。 */
function storedEmail(u) {
  if (!isObj(u) || typeof u.email !== 'string') return '';
  const n = normalizeEmail(u.email);
  return n.ok ? n.value : '';
}

/** 正規化位址 → 持有它的帳號們 */
function buildEmailIndex(entries) {
  const idx = new Map();
  for (let i = 0; i < entries.length; i++) {
    const e = storedEmail(entries[i].user);
    if (!e) continue;
    const list = idx.get(e);
    if (list) list.push(entries[i].username); else idx.set(e, [entries[i].username]);
  }
  return idx;
}

function fail(code, error, extra) {
  return Object.assign({ ok: false, error, code }, extra);
}

/** 單一位址的完整檢查。index 可為 null（不做唯一性檢查）。 */
function checkEmail(raw, allowed, index, selfUsername) {
  if (raw === undefined) return fail('MISSING', '未提供 Email 欄位（要清除請明確傳入空字串）');
  if (raw === null || (typeof raw === 'string' && raw.trim() === '')) return { ok: true, value: '', code: 'CLEARED' };
  const n = normalizeEmail(raw);
  if (!n.ok) return fail(n.code, n.error);
  if (!isAllowedDomain(n.value, allowed)) {
    return fail('DOMAIN_NOT_ALLOWED', 'Email 網域必須是：' + allowed.map((d) => safeText(d, 80)).join('、'));
  }
  if (index) {
    const holders = index.get(n.value);
    if (holders) {
      for (let i = 0; i < holders.length; i++) {
        if (holders[i] !== selfUsername) {
          return fail('DUPLICATE', '這個 Email 已經被另一個帳號使用，每個 Email 只能對應一個帳號', { conflictUsername: safeText(holders[i], 64) });
        }
      }
    }
  }
  return { ok: true, value: n.value, code: 'OK' };
}

// ── 單筆驗證 ────────────────────────────────────────────────────────────────
function validateUserEmail(raw, opts) {
  const o = isObj(opts) ? opts : {};
  const selfUsername = typeof o.selfUsername === 'string' ? o.selfUsername : '';
  const needIndex = typeof raw === 'string' && raw.trim() !== '';
  const index = needIndex ? buildEmailIndex(userEntries(o.users)) : null;
  return checkEmail(raw, allowedOf(o.config), index, selfUsername);
}

// ── 稽核文字 ────────────────────────────────────────────────────────────────
function present(v) { return typeof v === 'string' && v.trim() !== ''; }

function sameAddress(a, b) {
  const na = normalizeEmail(a);
  const nb = normalizeEmail(b);
  if (na.ok && nb.ok) return na.value === nb.value;
  return a.trim() === b.trim();
}

function maskedAuditDetail(oldEmail, newEmail) {
  const hadOld = present(oldEmail);
  const hasNew = present(newEmail);
  if (!hadOld && !hasNew) return AUDIT_TEXT.SAME;
  if (!hadOld) return AUDIT_TEXT.SET;
  if (!hasNew) return AUDIT_TEXT.CLEARED;
  return sameAddress(oldEmail, newEmail) ? AUDIT_TEXT.SAME : AUDIT_TEXT.CHANGED;
}

// ── 批次匯入：解析 ──────────────────────────────────────────────────────────
function isSepChar(ch) {
  return ch === ',' || ch === ';' || ch === FW_COMMA || ch === FW_SEMI || RE_WS.test(ch);
}

function unquote(field) {
  const s = field.trim();
  if (s.length >= 2) {
    const a = s.charAt(0);
    const b = s.charAt(s.length - 1);
    if ((a === '"' && b === '"') || (a === CURLY_OPEN && b === CURLY_CLOSE)) return s.slice(1, -1).trim();
  }
  return s;
}

/** 一行 → {userTok, emailTok}。Email 是最後一欄；其餘（去掉尾端分隔符）是帳號。 */
function splitLine(line) {
  let s = line.trim();
  let end = s.length;
  while (end > 0 && isSepChar(s.charAt(end - 1))) end--;       // 行尾多餘的逗號等
  s = s.slice(0, end);
  let j = s.length;
  while (j > 0 && !isSepChar(s.charAt(j - 1))) j--;             // j＝最後一欄的起點
  let k = j;
  while (k > 0 && isSepChar(s.charAt(k - 1))) k--;              // k＝帳號欄的終點
  return { userTok: unquote(s.slice(0, k)), emailTok: unquote(s.slice(j)) };
}

function fatalResult(code, error) {
  return {
    rows: [{ line: 0, username: '', email: '', status: 'error', code, error }],
    summary: { ok: 0, unchanged: 0, error: 1 },
    fatal: true,
  };
}

function errRow(line, username, email, code, error, extra) {
  return Object.assign({ line, username, email, status: 'error', code, error }, extra);
}

function parseBulkEmailText(text, opts) {
  const o = isObj(opts) ? opts : {};
  if (typeof text !== 'string') return fatalResult('NOT_STRING', '匯入內容必須是文字');
  if (text.length > BULK_LIMITS.maxTextChars) {
    return fatalResult('TOO_LARGE', '匯入內容過大（上限 ' + BULK_LIMITS.maxTextChars + ' 字元），請分批匯入');
  }
  const allowed = allowedOf(o.config);
  const entries = userEntries(o.users);
  const byName = new Map();
  for (let i = 0; i < entries.length; i++) byName.set(entries[i].username, entries[i].user);
  const emailIndex = buildEmailIndex(entries);
  let lowerNames = null;                                         // 延遲建立（只有「帳號不存在」時才需要）

  // 先挑出資料行（trim 會一併去掉 BOM），超過上限就整批拒絕
  const lines = text.split(/\r\n|\n|\r/);
  const data = [];
  for (let i = 0; i < lines.length; i++) {
    const t = lines[i].trim();
    if (t === '' || t.charAt(0) === '#') continue;
    data.push({ line: i + 1, text: t });
    if (data.length > BULK_LIMITS.maxRows) {
      return fatalResult('TOO_MANY_ROWS', '資料超過 ' + BULK_LIMITS.maxRows + ' 行的上限，請分批匯入（整批未處理）');
    }
  }

  const rows = data.map((d) => {
    if (d.text.length > BULK_LIMITS.maxLineChars) return errRow(d.line, '', '', 'LINE_TOO_LONG', '這一行太長，無法處理');
    const f = splitLine(d.text);
    if (!f.userTok || !f.emailTok) {
      return errRow(d.line, safeText(d.text, 100), '', 'FORMAT', '格式不正確，每行需為「帳號, Email」（以逗號、Tab 或空白分隔）');
    }
    if (f.userTok.indexOf('@') >= 0 && f.emailTok.indexOf('@') < 0) {
      return errRow(d.line, safeText(f.userTok, 100), safeText(f.emailTok, 120), 'SWAPPED', '欄位順序應為「帳號, Email」');
    }
    const user = byName.get(f.userTok);
    if (!user) {
      if (!lowerNames) {
        lowerNames = new Set();
        byName.forEach((_u, name) => lowerNames.add(name.toLowerCase()));
      }
      const hint = lowerNames.has(f.userTok.toLowerCase()) ? '（帳號區分大小寫，請確認大小寫）' : '';
      return errRow(d.line, safeText(f.userTok, 100), safeText(f.emailTok, 120), 'UNKNOWN_USER', '帳號不存在' + hint);
    }
    const v = checkEmail(f.emailTok, allowed, emailIndex, f.userTok);
    if (!v.ok) {
      const extra = v.conflictUsername ? { conflictUsername: v.conflictUsername } : undefined;
      return errRow(d.line, f.userTok, safeText(f.emailTok, 120), v.code, v.error, extra);
    }
    if (v.code === 'CLEARED') return errRow(d.line, f.userTok, '', 'FORMAT', 'Email 欄位不可空白');
    return { line: d.line, username: f.userTok, email: v.value, status: storedEmail(user) === v.value ? 'unchanged' : 'ok' };
  });

  // 同一批內的重複：涉及的每一行都標錯誤（只看第一輪判定為 ok／unchanged 的行，彼此獨立）
  const byEmail = new Map();
  const byUser = new Map();
  rows.forEach((r, i) => {
    if (r.status === 'error') return;
    (byEmail.get(r.email) || byEmail.set(r.email, []).get(r.email)).push(i);
    (byUser.get(r.username) || byUser.set(r.username, []).get(r.username)).push(i);
  });
  const mark = (i, code, error) => { rows[i].status = 'error'; rows[i].code = code; rows[i].error = error; };
  byEmail.forEach((list) => {
    if (list.length > 1) list.forEach((i) => mark(i, 'DUP_EMAIL_IN_BATCH', '同一批資料中有多個帳號使用相同的 Email'));
  });
  byUser.forEach((list) => {
    if (list.length > 1) list.forEach((i) => mark(i, 'DUP_USER_IN_BATCH', '同一個帳號在這批資料中出現多次'));
  });

  const summary = { ok: 0, unchanged: 0, error: 0 };
  rows.forEach((r) => { summary[r.status]++; });
  return { rows, summary, fatal: false };
}

// ── 批次匯入：套用 ──────────────────────────────────────────────────────────
function applyBulkPlan(rows, users, opts) {
  const o = isObj(opts) ? opts : {};
  if (!Array.isArray(users)) {
    return { updated: [], skipped: [], users, error: 'users 必須是陣列，未做任何變更' };
  }
  const out = users.map((u) => (isObj(u) ? { ...u } : u));
  const updated = [];
  const skipped = [];
  if (!Array.isArray(rows)) return { updated, skipped, users: out, error: 'rows 必須是陣列，未做任何變更' };
  if (rows.length > BULK_LIMITS.maxRows) {
    return { updated, skipped, users: out, error: 'rows 超過 ' + BULK_LIMITS.maxRows + ' 行的上限，未做任何變更' };
  }

  const allowed = allowedOf(o.config);
  const pos = new Map();                                         // username → out 的索引（第一個同名者）
  for (let i = 0; i < out.length; i++) {
    const u = out[i];
    if (isObj(u) && typeof u.username === 'string' && u.username !== '' && !pos.has(u.username)) pos.set(u.username, i);
  }
  const entries = [];
  pos.forEach((i, username) => entries.push({ username, user: out[i] }));
  const emailIndex = buildEmailIndex(entries);                   // 隨套用演進，後面的行會看到前面行的結果

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (!isObj(row) || row.status !== 'ok') continue;
    const un = typeof row.username === 'string' ? row.username : '';
    const at = pos.get(un);
    if (at === undefined) { skipped.push({ username: safeText(un, 64), code: 'UNKNOWN_USER' }); continue; }
    const v = checkEmail(row.email, allowed, emailIndex, un);
    if (!v.ok) { skipped.push({ username: safeText(un, 64), code: v.code }); continue; }
    if (v.code === 'CLEARED') { skipped.push({ username: safeText(un, 64), code: 'EMPTY' }); continue; }
    const cur = out[at];
    const before = storedEmail(cur);
    if (before === v.value) continue;                            // 已經是這個位址，不算變更
    const previous = typeof cur.email === 'string' ? cur.email : '';
    if (before) {
      const list = emailIndex.get(before);
      if (list) {
        const k = list.indexOf(un);
        if (k >= 0) list.splice(k, 1);
        if (!list.length) emailIndex.delete(before);
      }
    }
    const holders = emailIndex.get(v.value);
    if (holders) holders.push(un); else emailIndex.set(v.value, [un]);
    cur.email = v.value;
    updated.push({ username: un, previous, email: v.value, detail: maskedAuditDetail(previous, v.value) });
  }
  return { updated, skipped, users: out };
}

// ── 缺 Email 報表 ───────────────────────────────────────────────────────────
// 與 recipients.js 同規則的小型複本（見檔頭「限制」）
function isInactive(u) { return u.active === false || u.disabled === true || u.role === 'pool'; }

function labelOf(u, username) {
  const cands = [u.nickname, u.displayName, username];
  for (let i = 0; i < cands.length; i++) {
    const c = cands[i];
    if (typeof c === 'string' && c.trim() !== '') {
      const t = safeText(c, 60);
      if (t) return t;
    }
  }
  return '同仁';
}

/** 回傳 ''（Email 可用）或 NO_EMAIL／BAD_EMAIL／DOMAIN_NOT_ALLOWED */
function emailProblem(u, allowed) {
  const raw = u.email;
  if (raw === undefined || raw === null || (typeof raw === 'string' && raw.trim() === '')) return 'NO_EMAIL';
  const n = normalizeEmail(raw);
  if (!n.ok) return 'BAD_EMAIL';
  if (!isAllowedDomain(n.value, allowed)) return 'DOMAIN_NOT_ALLOWED';
  return '';
}

function missingEmailReport(opts) {
  const o = isObj(opts) ? opts : {};
  const allowed = allowedOf(o.config);
  const entries = userEntries(o.users);
  const byName = new Map();
  for (let i = 0; i < entries.length; i++) byName.set(entries[i].username, entries[i].user);

  const roles = new Map();                                       // username → Set(roleKey)
  const addRole = (username, key) => {
    let s = roles.get(username);
    if (!s) { s = new Set(); roles.set(username, s); }
    s.add(key);
  };
  entries.forEach((e) => {
    if (e.user.role === 'manager1') addRole(e.username, 'manager1');
    if (e.user.role === 'secretary') addRole(e.username, 'secretary');
  });
  const roster = isObj(o.roster) ? o.roster : {};
  Object.keys(ROSTER_FIELDS).forEach((field) => {
    const list = roster[field];
    if (!Array.isArray(list)) return;
    const n = Math.min(list.length, 1000);
    for (let i = 0; i < n; i++) {
      const un = list[i];
      if (typeof un === 'string' && byName.has(un)) addRole(un, ROSTER_FIELDS[field]);
    }
  });

  const out = [];
  roles.forEach((set, username) => {
    const u = byName.get(username);
    if (isInactive(u)) return;
    const reason = emailProblem(u, allowed);
    if (!reason) return;
    const keys = ROLE_KEYS.filter((k) => set.has(k));
    out.push({
      username,
      label: labelOf(u, username),
      roles: keys,
      roleLabels: keys.map((k) => ROLE_LABELS[k]),
      reason,
      readOnly: u.accessMode === 'view',
    });
  });
  out.sort((a, b) => (a.username < b.username ? -1 : (a.username > b.username ? 1 : 0)));
  return out;
}

module.exports = {
  validateUserEmail,
  parseBulkEmailText,
  applyBulkPlan,
  missingEmailReport,
  maskedAuditDetail,
  ROLE_KEYS,
  ROLE_LABELS,
  BULK_LIMITS,
};
