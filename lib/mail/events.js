'use strict';
/**
 * lib/mail/events.js — 信件事件的資料定義與驗證（純資料＋純函式，無 I/O）
 *
 * 匯出與簽名：
 *   EVENT_TYPES    = ['E1_SUBMIT','E2_COST_REQUEST','E3_NEXT_STEP','E4_RESULT','E5_COST_DONE','E6_WITHDRAWN']
 *   RESULT_KINDS   = ['approved','final_approved','rejected','returned','withdrawn','voided']
 *   STEP_LEVELS    = [1, 2, 3, 'board', null]
 *   LIMITS         各欄位長度上限（見下）
 *   QUOTE_ID_RE    = /^[A-Za-z0-9_-]{1,64}$/     單據 id 格式（同時用於連結與 dedupeKey）
 *   isValidQuoteId(id)               → boolean
 *   validateEvent(ev)                → {ok:true} | {ok:false, error, field}
 *   dedupeKey(ev, username)          → string      `${type}:${quoteId}:${username}:${stepKey}`；不合法 throw MailEventError
 *   MailEventError                   extends Error，帶 .code（BAD_EVENT / NO_STEP_KEY / BAD_USERNAME）
 *
 * Event 形狀（金額單位一律是「分」的整數；來源是簽核 approval.derived，由呼叫端換算好再傳進來）：
 *   {
 *     type, quoteId, quoteNo, projectName, company, ownerLabel,
 *     step:    { level: 1|2|3|'board'|null, label: '一級主管'|'總經理'|'董事長'|'董事會'|... },   // 這封信要這位收件人處理的關卡
 *     numbers: null | { revenueCents, gpCents, marginText: '41.04%', marginPct: 41.04, tierLevel: 1|2|3|null, tierLabel },
 *     result?: { kind, reason? },                    // E4／E6
 *     items?:  [{ desc, qty, unit }],                // E2 顧問信用；只取這三個欄位，永遠不含價格
 *     actor?:  { label }, at: ISO 時間字串, stepKey: string
 *   }
 *
 * 驗證規則（除「型別／長度／列舉」外，本檔補充的規則——若與後續階段的實際資料衝突，請回報而不要默默放寬）：
 *   - 必填：type、quoteId、quoteNo、stepKey、at。
 *   - E1／E3 必須有 step（物件且 label 非空）；E4 必須有 result 且 kind ∈ approved/final_approved/rejected/returned；
 *     E6 必須有 result 且 kind ∈ withdrawn/voided。E2／E5 不要求 step／result。
 *   - revenueCents：非負安全整數（負數、小數、NaN、字串一律拒絕）。gpCents：安全整數（可為負：虧損單正是簽核人最需要看到的）。
 *     marginPct：有限數字。marginText：≤32 字、數字格式（可含負號、整數最多 20 位、小數點、結尾 %）。
 *     極端虧損單（營收近乎 0、成本巨大）的 marginText 可以有 7 位以上整數：驗證放行，信上由 displayMargin 顯示成「<-999999%」。
 *   - stepKey 不可空白（不給預設值，避免「誤去重」或「完全不去重」）。
 *   - 自由文字欄位（projectName／company／ownerLabel／reason／desc／label）只檢查型別與長度；
 *     控制字元、換行、HTML 由 renderer 的 safeText／headerSafe／escHtml 處理。
 *
 * dedupeKey 的 username 只跳脫 '%' 與 ':'（一般帳號字串原樣不變）。理由：系統對帳號格式沒有限制，
 * 若帳號含 ':'，兩個不同的 (username, stepKey) 組合可能組出同一把 key，造成漏寄。其餘三段
 * （type／quoteId 不含 ':'，stepKey 在最後）本來就不會有歧義。
 */

const EVENT_TYPES = Object.freeze(['E1_SUBMIT', 'E2_COST_REQUEST', 'E3_NEXT_STEP', 'E4_RESULT', 'E5_COST_DONE', 'E6_WITHDRAWN']);
const RESULT_KINDS = Object.freeze(['approved', 'final_approved', 'rejected', 'returned', 'withdrawn', 'voided']);
const STEP_LEVELS = Object.freeze([1, 2, 3, 'board', null]);
const QUOTE_ID_RE = /^[A-Za-z0-9_-]{1,64}$/;

const RESULT_KINDS_BY_TYPE = Object.freeze({
  E4_RESULT: Object.freeze(['approved', 'final_approved', 'rejected', 'returned']),
  E6_WITHDRAWN: Object.freeze(['withdrawn', 'voided']),
});

const LIMITS = Object.freeze({
  quoteNo: 40,
  projectName: 300,
  company: 300,
  ownerLabel: 100,
  stepLabel: 40,
  tierLabel: 20,
  marginText: 32,
  reason: 2000,
  itemDesc: 500,
  itemUnit: 20,
  itemQty: 20,
  items: 200,
  actorLabel: 100,
  stepKey: 200,
  username: 200,
  at: 40,
});

const RE_CTRL = /[\x00-\x1f\x7f]/;
// 毛利率文字的「格式」檢查（防注入）：整數部分最多 20 位。lib/quoteApproval.js 的 marginText 對負毛利沒有上限（BigInt 運算，
// 安全整數的虧損 ÷ 1 分營收可到 18 位整數），所以這裡不能因為「數字太大」而退件——簽核信因極端數字而寄不出去，
// 正是最需要簽核人注意的虧損單收不到信。數字太大的處理是「顯示時」截成 <-999999%（見 displayMargin），不是退件。
const RE_MARGIN_TEXT = /^-?[0-9]{1,20}(?:\.[0-9]{1,4})?%?$/;
const MARGIN_DISPLAY_MAX_INT_DIGITS = 6;                  // 信上顯示的毛利率整數部分最多 6 位（±999999%）
const RE_ISO = /^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,9})?)?(?:Z|[+-][0-9]{2}:[0-9]{2})$/;

class MailEventError extends Error {
  constructor(code, message) {
    super(message);
    this.name = 'MailEventError';
    this.code = code;
  }
}

function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }

// Date.parse 對 2026-02-30 這類日期會自動進位成 3 月 2 日而不是拒絕，所以另外用 Date.UTC 回推檢查年月日
function isRealCalendarDate(iso) {
  const y = Number(iso.slice(0, 4));
  const m = Number(iso.slice(5, 7));
  const d = Number(iso.slice(8, 10));
  const t = new Date(Date.UTC(y, m - 1, d));
  return t.getUTCFullYear() === y && t.getUTCMonth() === m - 1 && t.getUTCDate() === d;
}
function isValidQuoteId(id) { return typeof id === 'string' && QUOTE_ID_RE.test(id); }
function bad(field, error) { return { ok: false, error, field }; }

/** 檢查字串欄位。回 '' 表示通過，否則回錯誤訊息。 */
function strErr(v, max, opts) {
  const o = opts || {};
  if (v === undefined || v === null) return o.required ? '必填' : '';
  if (typeof v !== 'string') return '必須是文字';
  if (o.required && v.trim() === '') return '不可為空白';
  if (v.length > max) return '長度不可超過 ' + max + ' 字';
  if (o.noControl && RE_CTRL.test(v)) return '不可包含控制字元或換行';
  return '';
}

function validateStep(step, required) {
  if (step === undefined || step === null) return required ? bad('step', 'step 必填（此事件類型需要顯示目前關卡）') : null;
  if (!isObj(step)) return bad('step', 'step 必須是物件');
  const lv = step.level === undefined ? null : step.level;
  if (STEP_LEVELS.indexOf(lv) < 0) return bad('step.level', 'step.level 必須是 1、2、3、\'board\' 或 null');
  const e = strErr(step.label, LIMITS.stepLabel, { required: required });
  if (e) return bad('step.label', 'step.label ' + e);
  return null;
}

function validateNumbers(n) {
  if (n === undefined || n === null) return null;
  if (!isObj(n)) return bad('numbers', 'numbers 必須是物件或 null');
  if (!Number.isSafeInteger(n.revenueCents) || n.revenueCents < 0) return bad('numbers.revenueCents', 'revenueCents 必須是非負的整數（單位：分）');
  if (!Number.isSafeInteger(n.gpCents)) return bad('numbers.gpCents', 'gpCents 必須是整數（單位：分）');
  if (typeof n.marginPct !== 'number' || !isFinite(n.marginPct)) return bad('numbers.marginPct', 'marginPct 必須是有限數字');
  if (typeof n.marginText !== 'string' || n.marginText.length > LIMITS.marginText || !RE_MARGIN_TEXT.test(n.marginText)) {
    return bad('numbers.marginText', 'marginText 必須是數字格式字串（例如 41.04%）');
  }
  const tl = n.tierLevel === undefined ? null : n.tierLevel;
  if (tl !== null && tl !== 1 && tl !== 2 && tl !== 3) return bad('numbers.tierLevel', 'tierLevel 必須是 1、2、3 或 null');
  const e = strErr(n.tierLabel, LIMITS.tierLabel, { required: true });
  if (e) return bad('numbers.tierLabel', 'tierLabel ' + e);
  return null;
}

function validateResult(type, result) {
  const allowedKinds = RESULT_KINDS_BY_TYPE[type];
  const required = !!allowedKinds;
  if (result === undefined || result === null) return required ? bad('result', 'result 必填（此事件類型需要結果）') : null;
  if (!isObj(result)) return bad('result', 'result 必須是物件');
  if (RESULT_KINDS.indexOf(result.kind) < 0) return bad('result.kind', 'result.kind 不在列舉內');
  if (allowedKinds && allowedKinds.indexOf(result.kind) < 0) return bad('result.kind', 'result.kind 與事件類型不符');
  const e = strErr(result.reason, LIMITS.reason);
  if (e) return bad('result.reason', 'result.reason ' + e);
  return null;
}

function validateItems(items) {
  if (items === undefined || items === null) return null;
  if (!Array.isArray(items)) return bad('items', 'items 必須是陣列');
  if (items.length > LIMITS.items) return bad('items', 'items 最多 ' + LIMITS.items + ' 筆');
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (!isObj(it)) return bad('items[' + i + ']', '品項必須是物件');
    let e = strErr(it.desc, LIMITS.itemDesc, { required: true });
    if (e) return bad('items[' + i + '].desc', 'desc ' + e);
    e = strErr(it.unit, LIMITS.itemUnit);
    if (e) return bad('items[' + i + '].unit', 'unit ' + e);
    if (it.qty !== undefined && it.qty !== null) {
      if (typeof it.qty === 'number') {
        if (!isFinite(it.qty) || it.qty < 0) return bad('items[' + i + '].qty', 'qty 必須是非負的有限數字');
      } else {
        e = strErr(it.qty, LIMITS.itemQty);
        if (e) return bad('items[' + i + '].qty', 'qty ' + e);
      }
    }
  }
  return null;
}

/**
 * 驗證事件。回 {ok:true} 或 {ok:false, error, field}（error 為繁中描述）。永不 throw
 * （事件物件若帶有會丟例外的 getter／Proxy，視為不合法）。
 */
function validateEvent(ev) {
  try {
    return validateEventInner(ev);
  } catch (e) {
    return bad('', '事件內容無法讀取');
  }
}

function validateEventInner(ev) {
  if (!isObj(ev)) return bad('', '事件必須是物件');
  if (typeof ev.type !== 'string' || EVENT_TYPES.indexOf(ev.type) < 0) return bad('type', '事件類型不在列舉內');
  if (!isValidQuoteId(ev.quoteId)) return bad('quoteId', '單據 id 格式不合法（僅允許英數、底線、連字號，1–64 字元）');
  let e = strErr(ev.quoteNo, LIMITS.quoteNo, { required: true, noControl: true });
  if (e) return bad('quoteNo', 'quoteNo ' + e);
  e = strErr(ev.stepKey, LIMITS.stepKey, { required: true, noControl: true });
  if (e) return bad('stepKey', 'stepKey ' + e + '（去重用，不提供預設值）');
  if (typeof ev.at !== 'string' || ev.at.length > LIMITS.at || !RE_ISO.test(ev.at) || !isFinite(Date.parse(ev.at)) || !isRealCalendarDate(ev.at)) {
    return bad('at', 'at 必須是 ISO 8601 時間字串');
  }
  e = strErr(ev.projectName, LIMITS.projectName);
  if (e) return bad('projectName', 'projectName ' + e);
  e = strErr(ev.company, LIMITS.company);
  if (e) return bad('company', 'company ' + e);
  e = strErr(ev.ownerLabel, LIMITS.ownerLabel);
  if (e) return bad('ownerLabel', 'ownerLabel ' + e);

  let r = validateStep(ev.step, ev.type === 'E1_SUBMIT' || ev.type === 'E3_NEXT_STEP');
  if (r) return r;
  r = validateNumbers(ev.numbers);
  if (r) return r;
  r = validateResult(ev.type, ev.result);
  if (r) return r;
  r = validateItems(ev.items);
  if (r) return r;
  if (ev.actor !== undefined && ev.actor !== null) {
    if (!isObj(ev.actor)) return bad('actor', 'actor 必須是物件');
    e = strErr(ev.actor.label, LIMITS.actorLabel, { required: true });
    if (e) return bad('actor.label', 'actor.label ' + e);
  }
  return { ok: true };
}

/**
 * 毛利率文字 → 信上顯示的字串（一律帶一個 %）。已通過 validateEvent 的 marginText 才應傳入。
 * 整數部分超過 6 位（|毛利率| ≥ 1,000,000%）時不顯示一長串數字，改顯示「<-999999%」或「>999999%」：
 * 版面（窄色塊）放得下，且簽核信不會因為極端虧損單而寄不出去。6 位以內原樣顯示（含小數）。
 */
function displayMargin(marginText) {
  const s = String(marginText).replace(/%$/, '');
  const neg = s.charAt(0) === '-';
  const intPart = (neg ? s.slice(1) : s).split('.')[0].replace(/^0+(?=[0-9])/, '');
  if (/^[0-9]+$/.test(intPart) && intPart.length > MARGIN_DISPLAY_MAX_INT_DIGITS) {
    return (neg ? '<-' : '>') + '9'.repeat(MARGIN_DISPLAY_MAX_INT_DIGITS) + '%';
  }
  return s + '%';
}

/**
 * 去重鍵：`${type}:${quoteId}:${username}:${stepKey}`。
 * 任何一段不合法（含 stepKey 空白）都 throw MailEventError——呼叫端（dispatcher）必須 catch，
 * 絕不可以用預設值硬湊一把鍵。
 */
function dedupeKey(ev, username) {
  if (!isObj(ev) || typeof ev.type !== 'string' || EVENT_TYPES.indexOf(ev.type) < 0) {
    throw new MailEventError('BAD_EVENT', '事件類型不合法，無法產生去重鍵');
  }
  if (!isValidQuoteId(ev.quoteId)) throw new MailEventError('BAD_EVENT', '單據 id 格式不合法，無法產生去重鍵');
  if (strErr(ev.stepKey, LIMITS.stepKey, { required: true, noControl: true })) {
    throw new MailEventError('NO_STEP_KEY', 'stepKey 不可空白，無法產生去重鍵');
  }
  if (typeof username !== 'string' || username === '' || username.length > LIMITS.username || RE_CTRL.test(username)) {
    throw new MailEventError('BAD_USERNAME', '收件人帳號不合法，無法產生去重鍵');
  }
  const u = username.replace(/%/g, '%25').replace(/:/g, '%3A');
  return ev.type + ':' + ev.quoteId + ':' + u + ':' + ev.stepKey;
}

module.exports = {
  EVENT_TYPES,
  RESULT_KINDS,
  STEP_LEVELS,
  LIMITS,
  QUOTE_ID_RE,
  isValidQuoteId,
  validateEvent,
  displayMargin,
  dedupeKey,
  MailEventError,
};
