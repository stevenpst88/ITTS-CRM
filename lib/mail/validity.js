'use strict';
/**
 * lib/mail/validity.js — 「這封待寄／待重寄的信，在單據現況下還該不該寄」的純函式（README §7 的 isStillValid 契約）
 *
 * 為什麼獨立成檔：重試（路由內清理、poll-bundle、每日 Cron、後台重送）可能在事件發生後很久才寄出，
 * 這中間單據可能被重新送簽、撤回、作廢、改派、再核准。信件是「事件發生當下的通知」，狀態已經改變的舊信不能寄
 * （例如業務已重新送簽，卻收到「您的報價單已被駁回」；簽核人已輪到下一關，卻收到「請勿簽核」）。
 * 純函式（不讀時鐘、不碰 I/O、不 require 任何東西），所以每個事件類型 × 每種狀態變化都能用假單據窮舉測試。
 *
 * 匯出：
 *   checkValidity(job, q, helpers) → { valid: boolean, code: string }
 *       job      { type, toUser, dedupeKey }   stepKey 取自 dedupeKey（`type:quoteId:user:stepKey`，stepKey 內可含 ':'）
 *       q        單據現況（approval／costFlow／owner／costBy）；null／非物件 → { valid:false, code:'NO_QUOTE' }
 *       helpers  { stepRecipients(step) → string[] }   只有 E1／E3 用（收件人仍在該關才寄）
 *       valid 只有明確為 true 才可寄；未知事件類型、壞掉的 stepKey 一律 false（寧可不寄，不可誤寄）
 *   stepKeyOf(job)                 → string    dedupeKey 第三個 ':' 之後的部分
 *   stepIdxOf(stepKey)             → number    E1／E3 的關卡序號（<submittedAt>#<idx>[@<改派時間>]）；格式不符回 -1
 *   reassignEpoch(ap)              → string    目前這一輪送簽之後「最近一次承辦人真的改變的改派」的時間（ISO），沒有則 ''；改派給同一人（歷史 meta.from===meta.to）不算
 *   e1StepKey(ap, idx)             → string    E1 的 stepKey：`<submittedAt>#<idx>`；改派過則多 `@<改派時間>`
 *   VALIDITY_CODES                 所有 code 的清單（凍結）
 *
 * 各事件的有效條件（stepKey 內的 S＝事件發生當下的 approval.submittedAt，也就是「哪一次送簽」）：
 *   E1 送簽／改派   state=pending、S 仍是目前這次送簽、cur 仍是該關（0）、改派標記仍是最近一次改派、收件人仍在該關
 *                   （同一次送簽內 A→B→A：A 的第二封 E1 stepKey 帶改派時間，是新的去重鍵；A→B 之後 A 先前未寄出的 E1 因改派標記不符而過期；
 *                   A→A〔管理員改派給目前的承辦人本人〕不改變標記：不產生新的去重鍵、不多寄，A 還在重試中的 E1 也不會被取消）
 *   E3 下一關       state=pending、S 相同、cur 仍是該關、收件人仍在該關
 *   E2 請填成本     costFlow.state=requested、costBy 仍是此人、stepKey 仍等於 `requestedAt#costBy`
 *   E4 本關通過     state=pending、S 相同、該關 status=approved 且 cur 已越過該關（已經走完全部關卡、被駁回、撤回、作廢、重新送簽 → 過期）
 *   E4 最終核准     state=approved、S 相同、該關 status=approved（作廢／重新送簽後 S 已變或 state 已變 → 過期）
 *   E4 駁回         state=returned、S 相同、該關 status=returned（業務重新送簽後 S 已變、state 變成 pending → 過期）
 *                   駁回原因是業務可輸入的自由文字，只在信件內容裡顯示，不參與有效性判斷
 *   E5 成本完成     costFlow.state=filled 且 filledAt 仍等於 stepKey 內的時間（顧問改回「未完成」、改品項使成本退回 → 過期）
 *   E6 撤回         state=none 且 submittedAt 仍是被撤回的那一次（撤回不會清掉 submittedAt；業務重新送簽後 S 變了 → 過期）
 *   E6 作廢         state=none 且 submittedAt 為空（作廢會把整個 approval 重置、submittedAt 設為 null；業務重新送簽後 submittedAt 有值 → 過期）
 *                   限制：同一張單據被作廢兩次時，較早那次作廢通知的重試仍會通過（單據目前確實處於「核准已作廢」，內容為真；較晚那次另有自己的去重鍵）
 *   E4／E5 另外要求收件人（toUser）仍是單據的業務（q.owner）。
 *
 * 語意提醒（README §7）：呼叫端（quoteMail.isStillValid）只在 valid===true 時回傳 true；false → 取消（STALE）；丟例外／逾時 → 不寄也不取消，稍後重試。
 */

const VALIDITY_CODES = Object.freeze([
  'OK', 'NO_QUOTE', 'BAD_KEY', 'UNKNOWN_TYPE', 'NOT_PENDING', 'ROUND_CHANGED', 'STEP_MOVED', 'REASSIGNED', 'RECIPIENT_CHANGED',
  'COST_STATE', 'NOT_OWNER', 'NOT_APPROVED', 'NOT_RETURNED', 'NOT_WITHDRAWN', 'NOT_VOIDED',
]);

const RE_STEP = /^([^#@]+)#([0-9]{1,6})(?:@([^#@]+))?$/;              // E1／E3：<submittedAt>#<idx>[@<改派時間>]
const RE_RESULT = /^([^#@]+)#r:(approved|final_approved|rejected):([0-9]{1,6})$/;   // E4：<submittedAt>#r:<kind>:<idx>
const RE_DONE = /^([^#@]+)#done$/;                                     // E5：<filledAt>#done
const RE_WITHDRAW = /^([^#@]+)#(withdrawn|voided)$/;                   // E6：<submittedAt>#withdrawn|voided

function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }
const yes = () => ({ valid: true, code: 'OK' });
const no = (code) => ({ valid: false, code });

/** dedupeKey = `type:quoteId:user:stepKey`（user 的 ':' 已跳脫成 %3A；stepKey 內的 ISO 時間含 ':'，所以取第三個 ':' 之後的全部） */
function stepKeyOf(job) {
  return String((job && job.dedupeKey) || '').split(':').slice(3).join(':');
}

/** E1／E3 的 stepKey → 關卡序號；格式不符 -1 */
function stepIdxOf(sk) {
  const m = RE_STEP.exec(String(sk));
  return m ? Number(m[2]) : -1;
}

/**
 * 這筆 REASSIGN 是不是「改派給目前的承辦人本人」（承辦人沒有改變）。
 * lib/quoteRoutes.js 的 POST /reassign 把改派前後的承辦人記在歷史的 meta：{ from, to }（帳號）；兩者是同一個非空字串＝沒換人。
 * 沒有 meta／欄位缺漏／不是字串（上線前就存在的舊紀錄）一律回 false＝當成真的改派，行為與加入這個判斷之前完全相同。
 */
function isNoopReassign(h) {
  const m = isObj(h) && isObj(h.meta) ? h.meta : null;
  return !!m && typeof m.from === 'string' && m.from !== '' && m.from === m.to;
}

/**
 * 這一輪送簽（最近一次 SUBMIT）之後，最近一次「承辦人真的改變」的改派時間。
 * 從歷史的尾端往回找：先碰到 REASSIGN → 回傳它的時間（改派給同一人的 REASSIGN 略過，繼續往前找）；先碰到 SUBMIT（更早的改派屬於上一輪）→ ''。
 * 為什麼略過同人改派：E1 的 stepKey 帶這個時間，管理員把一級主管「改派」給目前的承辦人不該產生新的去重鍵（否則多寄一封相同的 E1），
 * 也不該讓還在重試中的舊 E1 因改派標記不符而被取消（否則承辦人一封都收不到）。
 */
function reassignEpoch(ap) {
  const hist = isObj(ap) && Array.isArray(ap.history) ? ap.history : [];
  for (let i = hist.length - 1; i >= 0; i--) {
    const h = hist[i];
    if (!isObj(h)) continue;
    if (h.action === 'SUBMIT') return '';
    if (h.action === 'REASSIGN') {
      if (isNoopReassign(h)) continue;
      return typeof h.at === 'string' && h.at !== '' && !/[#@]/.test(h.at) ? h.at : '';
    }
  }
  return '';
}

/** E1 的 stepKey：一般送簽＝`<submittedAt>#<idx>`（與改派功能加入前完全相同）；改派過＝再加 `@<最近一次改派時間>`，讓 A→B→A 的 A 是新的去重鍵 */
function e1StepKey(ap, idx) {
  const epoch = reassignEpoch(ap);
  return String(ap.submittedAt) + '#' + idx + (epoch ? '@' + epoch : '');
}

function checkValidity(job, q, helpers) {
  if (!isObj(job)) return no('BAD_KEY');
  if (!isObj(q)) return no('NO_QUOTE');
  const h = isObj(helpers) ? helpers : {};
  const ap = isObj(q.approval) ? q.approval : {};
  const cf = isObj(q.costFlow) ? q.costFlow : {};
  const steps = Array.isArray(ap.steps) ? ap.steps : [];
  const sk = stepKeyOf(job);
  const to = job.toUser;

  switch (job.type) {
    case 'E1_SUBMIT':
    case 'E3_NEXT_STEP': {
      const m = RE_STEP.exec(sk);
      if (!m) return no('BAD_KEY');
      if (job.type === 'E3_NEXT_STEP' && m[3] !== undefined) return no('BAD_KEY');      // 改派標記只屬於 E1
      const idx = Number(m[2]);
      if (ap.state !== 'pending') return no('NOT_PENDING');
      if (typeof ap.submittedAt !== 'string' || m[1] !== ap.submittedAt) return no('ROUND_CHANGED');
      if (ap.cur !== idx) return no('STEP_MOVED');
      const step = steps[idx];
      if (!isObj(step)) return no('STEP_MOVED');
      if (job.type === 'E1_SUBMIT' && (m[3] || '') !== reassignEpoch(ap)) return no('REASSIGNED');
      const list = typeof h.stepRecipients === 'function' ? h.stepRecipients(step) : [];
      return Array.isArray(list) && list.indexOf(to) >= 0 ? yes() : no('RECIPIENT_CHANGED');
    }
    case 'E2_COST_REQUEST': {
      if (cf.state !== 'requested' || typeof cf.requestedAt !== 'string' || !cf.requestedAt) return no('COST_STATE');
      if (!q.costBy || q.costBy !== to) return no('RECIPIENT_CHANGED');
      return sk === cf.requestedAt + '#' + q.costBy ? yes() : no('COST_STATE');
    }
    case 'E4_RESULT': {
      const m = RE_RESULT.exec(sk);
      if (!m) return no('BAD_KEY');
      if (!q.owner || q.owner !== to) return no('NOT_OWNER');
      const kind = m[2];
      const idx = Number(m[3]);
      const step = steps[idx];
      if (typeof ap.submittedAt !== 'string' || m[1] !== ap.submittedAt) return no('ROUND_CHANGED');
      if (kind === 'approved') {
        if (ap.state !== 'pending') return no('NOT_PENDING');
        if (!isObj(step) || step.status !== 'approved' || !(typeof ap.cur === 'number' && ap.cur > idx)) return no('STEP_MOVED');
        return yes();
      }
      if (kind === 'final_approved') {
        if (ap.state !== 'approved') return no('NOT_APPROVED');
        return isObj(step) && step.status === 'approved' ? yes() : no('NOT_APPROVED');
      }
      if (ap.state !== 'returned') return no('NOT_RETURNED');          // rejected
      return isObj(step) && step.status === 'returned' ? yes() : no('NOT_RETURNED');
    }
    case 'E5_COST_DONE': {
      const m = RE_DONE.exec(sk);
      if (!m) return no('BAD_KEY');
      if (!q.owner || q.owner !== to) return no('NOT_OWNER');
      return cf.state === 'filled' && typeof cf.filledAt === 'string' && cf.filledAt === m[1] ? yes() : no('COST_STATE');
    }
    case 'E6_WITHDRAWN': {
      const m = RE_WITHDRAW.exec(sk);
      if (!m) return no('BAD_KEY');
      if (ap.state !== 'none') return no(m[2] === 'withdrawn' ? 'NOT_WITHDRAWN' : 'NOT_VOIDED');
      if (m[2] === 'withdrawn') return typeof ap.submittedAt === 'string' && ap.submittedAt === m[1] ? yes() : no('ROUND_CHANGED');
      return ap.submittedAt === null || ap.submittedAt === undefined || ap.submittedAt === '' ? yes() : no('ROUND_CHANGED');
    }
    default:
      return no('UNKNOWN_TYPE');
  }
}

module.exports = { checkValidity, stepKeyOf, stepIdxOf, reassignEpoch, e1StepKey, VALIDITY_CODES };
