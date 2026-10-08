'use strict';
/**
 * lib/mail/safety.js — 信件管線的輸入安全小工具
 *
 * 純函式、無 I/O、無外部相依；所有函式「永不 throw」（輸入型別不對就回安全的空值／錯誤物件）。
 * 信件內容會離開存取控制（轉寄、手機、信箱代理），所以任何要進「收件人、主旨、標頭、HTML」的字串
 * 都必須先經過本檔的函式。其他 lib/mail/* 模組不得自己寫一套驗證。
 *
 * 匯出與簽名：
 *   normalizeEmail(raw)                  → {ok:true, value} | {ok:false, error, code}   單一位址嚴格驗證（value 為小寫）
 *   normalizeDomain(raw)                 → string        合法網域回小寫，否則 ''
 *   isAllowedDomain(email, allowed[])    → boolean       網域「完全相等」白名單比對（不做子網域／後綴比對）
 *   maskEmail(email)                     → 's***@itts.com.tw' | '(invalid)'
 *   maskedFromList(list)                 → string[]
 *   headerSafe(str, maxLen = 255)        → string        進標頭／主旨用：單行、無控制字元、截斷（不加省略號）
 *   escHtml(s)                           → string        跳脫 & < > " ' `
 *   clip(s, n = 200)                     → string        依 code point 截斷並加省略號（輸出最多 n+1 個 code point；
 *                                                        n 缺省／非數字／負數 → 200；n=0 → ''）
 *   safeText(s, n = 200)                 → string        去控制字元／單行化後再 clip（用於信件內文裡的使用者輸入）
 *   MAX_EMAIL_LENGTH                     = 254
 *
 * 設計重點（每一條都有對應測試與變異測試，見 scripts/check-mail-core.js、scripts/check-mail-mutation.js）：
 *  1. Email 的「非 ASCII」檢查必須在 toLowerCase 之前。Kelvin 符號（U+212A）之類的字元小寫化後會變成
 *     ASCII 的 k，若先轉小寫再檢查，就等於放行了全形／同形字。
 *  2. 白名單是 domain === 清單項目（精確相等）。evil-itts.com.tw、itts.com.tw.evil.com、sub.itts.com.tw
 *     都不會通過；清單項目不解讀萬用字元。
 *  3. Local part 採比 RFC 更嚴的白名單字元 [a-z0-9._+-]（不可首尾或連續的點），不收 % ! # $ & ' = ? ^ { | } ~ 等
 *     少見字元（% 是古老的 percent-hack 路由語法）。公司內部信箱實務上都落在這個範圍；
 *     真的有例外位址時，管理員會看到明確的錯誤訊息。
 *  4. 結尾的頂級網域（TLD）至少 2 字元且含字母（擋掉 a@1.2.3.4 這類 IP 形式）。
 *  5. 巨大輸入（200 萬字元）先硬切再處理：全部函式都是線性時間、無災難性回溯。
 *  6. headerSafe／safeText 會把 CR/LF/NUL/C0/C1 控制字元（含 U+0085、U+2028、U+2029）換成空白，
 *     並移除 bidi 方向控制字元（RLO 等）、零寬字元、BOM、Unicode Tag 字元；孤立的代理對換成 U+FFFD。
 *     截斷以 code point 計，不會把一個代理對（emoji）切成兩半。
 *
 * 原始碼撰寫備註：本檔的正規式一律用 \x.. 或 \u{...} 形式，不寫四位數的 \uXXXX。
 *   原因：部分編輯／寫檔工具會把 \uXXXX 轉成真正的字元，U+2028 一旦落在正規式字面值裡就是語法錯誤。
 *   scripts/check-mail-core.js 內有掃描，原始碼出現不可見字元會直接失敗。
 */

const MAX_EMAIL_LENGTH = 254;
const MAX_RAW_INPUT = 4096;          // normalizeEmail 的原始輸入硬上限（遠大於 254，只為擋巨大字串）
const ELLIPSIS = String.fromCharCode(0x2026);
const REPLACEMENT_CHAR = String.fromCharCode(0xfffd);
const DEFAULT_HEADER_MAX = 255;
const DEFAULT_CLIP = 200;

// ── 小工具 ──────────────────────────────────────────────────────────────────
function toStr(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  try { return String(v); } catch (e) { return ''; }
}

function fail(code, error) { return { ok: false, error, code }; }

function posInt(n, dflt) {
  const x = typeof n === 'number' && isFinite(n) ? Math.floor(n) : NaN;
  return x >= 0 ? x : dflt;
}

// 依 code point 取前 n 個；字串可能含孤立代理（clip 直接吃未清理的輸入），所以逐字檢查下一個是不是低代理
function takeCodePoints(s, n) {
  if (s.length <= n) return { text: s, truncated: false };   // UTF-16 單元數 <= n，code point 數必然 <= n
  let i = 0;
  let count = 0;
  while (i < s.length && count < n) {
    const c = s.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length) {
      const d = s.charCodeAt(i + 1);
      i += (d >= 0xdc00 && d <= 0xdfff) ? 2 : 1;
    } else {
      i += 1;
    }
    count++;
  }
  return { text: s.slice(0, i), truncated: i < s.length };
}

// ── Email 驗證 ──────────────────────────────────────────────────────────────
const RE_CONTROL = /[\x00-\x1f\x7f]/;
const RE_NON_ASCII = /[^\x00-\x7f]/;
const RE_LOCAL = /^[a-z0-9_+-]+(?:\.[a-z0-9_+-]+)*$/;
const RE_LABEL = /^[a-z0-9-]+$/;

/** 網域語法檢查（輸入須已是小寫 ASCII）：每段 1–63、英數與連字號、不以連字號起訖、至少一個點、TLD ≥2 且含字母、總長 ≤253。 */
function isValidDomainLower(d) {
  if (typeof d !== 'string' || d.length < 3 || d.length > 253) return false;
  const labels = d.split('.');
  if (labels.length < 2) return false;
  for (let i = 0; i < labels.length; i++) {
    const l = labels[i];
    if (l.length < 1 || l.length > 63) return false;
    if (!RE_LABEL.test(l)) return false;
    if (l.charAt(0) === '-' || l.charAt(l.length - 1) === '-') return false;
  }
  const tld = labels[labels.length - 1];
  if (tld.length < 2 || !/[a-z]/.test(tld)) return false;
  return true;
}

/**
 * 嚴格驗證單一 Email。
 * @returns {{ok:true,value:string}|{ok:false,error:string,code:string}}
 *   code：NOT_STRING / EMPTY / TOO_LONG / CONTROL_CHAR / NON_ASCII / SPACE / BAD_CHAR / BAD_FORMAT / BAD_LOCAL / BAD_DOMAIN
 */
function normalizeEmail(raw) {
  if (typeof raw !== 'string') return fail('NOT_STRING', 'Email 必須是文字');
  if (raw.length > MAX_RAW_INPUT) return fail('TOO_LONG', 'Email 過長');
  const s = raw.trim();
  if (!s) return fail('EMPTY', '尚未輸入 Email');
  if (s.length > MAX_EMAIL_LENGTH) return fail('TOO_LONG', 'Email 過長（上限 254 字元）');
  // 以下三道檢查一定要在 toLowerCase 之前（見檔頭設計重點 1）
  if (RE_CONTROL.test(s)) return fail('CONTROL_CHAR', 'Email 不可包含換行、Tab 或控制字元');
  if (RE_NON_ASCII.test(s)) return fail('NON_ASCII', 'Email 只能使用半形英數字與 . _ + - （不可含全形、中文或其他非 ASCII 字元）');
  if (s.indexOf(' ') >= 0) return fail('SPACE', 'Email 不可包含空白');
  if (/[,;<>"'()\[\]\\:]/.test(s)) return fail('BAD_CHAR', 'Email 不可包含逗號、分號、角括號、引號、括號、冒號或反斜線；一欄只能填一個位址');
  const v = s.toLowerCase();
  const at = v.indexOf('@');
  if (at < 0 || at !== v.lastIndexOf('@')) return fail('BAD_FORMAT', 'Email 必須恰好包含一個 @');
  const local = v.slice(0, at);
  const domain = v.slice(at + 1);
  if (local.length < 1 || local.length > 64) return fail('BAD_LOCAL', '@ 前面的帳號部分長度必須在 1–64 字元之間');
  if (!RE_LOCAL.test(local)) return fail('BAD_LOCAL', '@ 前面只能使用英數字與 . _ + -，且不可以句點開頭、結尾或連續出現');
  if (domain.length < 1 || domain.length > 253) return fail('BAD_DOMAIN', '@ 後面的網域長度不合法');
  if (!isValidDomainLower(domain)) return fail('BAD_DOMAIN', '@ 後面的網域格式不合法');
  return { ok: true, value: v };
}

/** 網域正規化：回小寫網域；不合法回 ''（供 config 解析白名單用）。非 ASCII 先拒絕、再轉小寫。 */
function normalizeDomain(raw) {
  if (typeof raw !== 'string' || raw.length > 400) return '';
  const s = raw.trim();
  if (!s || /[^\x21-\x7e]/.test(s)) return '';
  const d = s.toLowerCase();
  return isValidDomainLower(d) ? d : '';
}

/** 網域必須「完全等於」白名單其中一項。email 不合法、白名單不是陣列或是空的，一律 false。 */
function isAllowedDomain(email, allowedDomains) {
  const n = normalizeEmail(email);
  if (!n.ok) return false;
  if (!Array.isArray(allowedDomains)) return false;
  const domain = n.value.slice(n.value.indexOf('@') + 1);
  for (let i = 0; i < allowedDomains.length; i++) {
    const d = allowedDomains[i];
    if (typeof d === 'string' && d.trim().toLowerCase() === domain) return true;
  }
  return false;
}

/** 遮罩：只留 local 第一字與完整網域。不合法回 '(invalid)'（不回顯原值）。 */
function maskEmail(email) {
  const n = normalizeEmail(email);
  if (!n.ok) return '(invalid)';
  const at = n.value.indexOf('@');
  return n.value.charAt(0) + '***' + n.value.slice(at);
}

function maskedFromList(list) {
  if (!Array.isArray(list)) return [];
  const out = [];
  const max = Math.min(list.length, 1000);
  for (let i = 0; i < max; i++) out.push(maskEmail(list[i]));
  return out;
}

// ── 文字清理 ────────────────────────────────────────────────────────────────
// 孤立的高／低代理（u 旗標下，成對的代理會被視為一個天文字元，只有落單的才屬於 Cs 類別）→ U+FFFD
const RE_LONE_SURROGATE = /\p{Cs}/gu;
// C0、DEL、C1（含 U+0085）、U+2028/2029 → 空白（避免兩個字被黏在一起）
const RE_CONTROL_TO_SPACE = /[\x00-\x1f\x7f-\x9f\u{2028}\u{2029}]/gu;
// 不可見／方向控制字元 → 直接移除：ALM(061C)、零寬空白(200B)、LRM/RLM(200E/200F)、LRE/RLE/PDF/LRO/RLO(202A-202E)、
// WORD JOINER(2060)、LRI/RLI/FSI/PDI(2066-2069)、BOM(FEFF)、Tag 字元(E0000-E007F)
const RE_INVISIBLE = /[\u{61c}\u{200b}\u{200e}\u{200f}\u{202a}-\u{202e}\u{2060}\u{2066}-\u{2069}\u{feff}\u{e0000}-\u{e007f}]/gu;

/** 單行化：控制字元換空白、移除不可見字元、合併連續空白、trim。先硬切過長輸入（precap 為 UTF-16 單元數）。 */
function sanitizeLine(v, precap) {
  let s = toStr(v);
  if (s.length > precap) s = s.slice(0, precap);
  s = s.replace(RE_LONE_SURROGATE, REPLACEMENT_CHAR);
  s = s.replace(RE_CONTROL_TO_SPACE, ' ');
  s = s.replace(RE_INVISIBLE, '');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
}

/**
 * 進信件標頭（主旨等）的字串清理：單行、無控制字元、截斷到 maxLen 個 code point（不加省略號、不切斷代理對）。
 * 注意：這是「標頭安全」而不是 HTML 安全；放進 HTML 前仍要 escHtml。
 */
function headerSafe(str, maxLen) {
  const max = posInt(maxLen, DEFAULT_HEADER_MAX);
  const s = sanitizeLine(str, Math.max(max * 8, 1024));
  return takeCodePoints(s, max).text.trimEnd();
}

const RE_HTML_SPECIAL = /[&<>"'`]/;

/**
 * HTML 跳脫（& < > " ' `）。null/undefined → ''；非字串先轉字串。
 * 實作用「逐字元的字串取代」而不是 callback 版 replace：200 萬個 '<' 的最壞情況約 35ms（callback 版約 70ms）。
 * & 必須第一個取代，否則後面產生的實體會被二次跳脫。
 */
function escHtml(s) {
  const str = toStr(s);
  if (!RE_HTML_SPECIAL.test(str)) return str;
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/`/g, '&#96;');
}

/** 依 code point 截斷；超過 n 才加省略號（輸出最多 n+1 個 code point）。n 缺省 200，n=0 回 ''。 */
function clip(s, n) {
  const max = posInt(n, DEFAULT_CLIP);
  const str = toStr(s);
  if (max === 0) return '';
  const r = takeCodePoints(str, max);
  return r.truncated ? r.text.trimEnd() + ELLIPSIS : r.text;
}

/** 信件內文用的使用者輸入：單行化＋去不可見字元，再 clip 到 n（缺省 200）。 */
function safeText(s, n) {
  const max = posInt(n, DEFAULT_CLIP);
  return clip(sanitizeLine(s, Math.max(max * 8, 1024)), max);
}

module.exports = {
  normalizeEmail,
  normalizeDomain,
  isAllowedDomain,
  maskEmail,
  maskedFromList,
  headerSafe,
  escHtml,
  clip,
  safeText,
  MAX_EMAIL_LENGTH,
};
