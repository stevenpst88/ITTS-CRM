'use strict';
/**
 * 報價單號流水帳（quote-number ledger）與刪除快照。純函式，操作傳入的 data（主資料 blob），不碰資料庫。
 *
 *  data.quoteNoLedger[]     { no, owner, createdAt, status: 'active'|'deleted'|'invalid', [deletedAt, deletedBy, snapshotHash, backfilled] }
 *  data.deletedQuotations[] { id, quoteNo, deletedAt, deletedBy, reason, snapshotHash, quotation: <刪除當下的完整報價單> }
 *
 * 目的：單號不得無聲消失。刪除草稿會把整張單搬進 deletedQuotations（含 timeline／costFlow／items，內部資料不去敏）；
 * 流水帳逐號記錄；管理員可查「1..最大號之間既不在報價單也不在流水帳」的缺號（GAP_UNEXPLAINED）。
 * 舊資料（沒有流水帳）採惰性補登：刪除／作廢／管理員檢視時，缺的單號補一筆 backfilled:true。
 * deletedQuotations／quoteNoLedger 是獨立的頂層鍵，不在 data.quotations 內 → 一般清單、統計、業務可見範圍都不會碰到它們。
 */
const crypto = require('crypto');

const NO_RE = /^QU-(\d{6})-(\d+)$/;
const ledgerOf = (data) => { if (!Array.isArray(data.quoteNoLedger)) data.quoteNoLedger = []; return data.quoteNoLedger; };
const deletedOf = (data) => { if (!Array.isArray(data.deletedQuotations)) data.deletedQuotations = []; return data.deletedQuotations; };
const findEntry = (data, no) => ledgerOf(data).find((e) => e.no === no) || null;

/** 穩定序列化（鍵排序）；用於快照雜湊，與物件鍵順序無關 */
function canonical(v) {
  if (v === null || typeof v !== 'object') return JSON.stringify(v === undefined ? null : v);
  if (Array.isArray(v)) return '[' + v.map(canonical).join(',') + ']';
  return '{' + Object.keys(v).sort().filter((k) => v[k] !== undefined).map((k) => JSON.stringify(k) + ':' + canonical(v[k])).join(',') + '}';
}
const snapshotHash = (q) => crypto.createHash('sha256').update(canonical(q)).digest('hex');

/** 確保該單號有流水帳項目；沒有就補登（backfilled=true 表示建立當時沒有流水帳，資訊來自單據現況） */
function ensureEntry(data, q, nowIso, backfilled) {
  if (!q || !q.quoteNo) return null;
  let e = findEntry(data, q.quoteNo);
  if (!e) {
    e = { no: q.quoteNo, owner: q.owner || '', createdAt: q.createdAt || nowIso, status: q.invalid ? 'invalid' : 'active' };
    if (backfilled) e.backfilled = true;
    ledgerOf(data).push(e);
  }
  return e;
}
const markCreated = (data, q, nowIso) => ensureEntry(data, q, nowIso, false);
function markInvalid(data, q, nowIso) { const e = ensureEntry(data, q, nowIso, true); if (e && e.status !== 'deleted') e.status = 'invalid'; return e; }
function markActive(data, q, nowIso) { const e = ensureEntry(data, q, nowIso, true); if (e && e.status !== 'deleted') e.status = 'active'; return e; }

/** 刪除：整張搬進 deletedQuotations，流水帳標 deleted。呼叫端負責把 q 從 data.quotations 移除與存檔。回傳 { rec, hash } */
function recordDeletion(data, q, by, nowIso, reason) {
  const hash = snapshotHash(q);
  const snap = JSON.parse(JSON.stringify(q));
  const rec = { id: q.id, quoteNo: q.quoteNo, deletedAt: nowIso, deletedBy: by, reason: reason || '', snapshotHash: hash, quotation: snap };
  deletedOf(data).push(rec);
  const e = ensureEntry(data, q, nowIso, true);
  if (e) { e.status = 'deleted'; e.deletedAt = nowIso; e.deletedBy = by; e.snapshotHash = hash; }
  return { rec, hash };
}

/**
 * 管理員檢視：補登缺的項目（在報價單裡但不在流水帳）並彙整缺號。會修改 data（補登），回傳 { changed, entries, months, summary }。
 * qa 為 lib/quoteApproval（算金額用，可省略）。
 */
function review(data, nowIso) {
  const quotes = Array.isArray(data.quotations) ? data.quotations : [];
  let changed = false;
  for (const q of quotes) {
    if (!q || !q.quoteNo) continue;
    if (!findEntry(data, q.quoteNo)) { ensureEntry(data, q, nowIso, true); changed = true; }
  }
  const live = new Map(quotes.filter((q) => q && q.quoteNo).map((q) => [q.quoteNo, q]));
  const entries = ledgerOf(data).map((e) => {
    const q = live.get(e.no);
    // 單據還在 → 以單據現況為準（作廢旗標可能被商機還原清掉）；不在 → 看流水帳（deleted）
    const status = q ? (q.invalid ? 'invalid' : 'active') : (e.status === 'deleted' ? 'deleted' : 'missing');
    return Object.assign({}, e, { status, present: !!q });
  });
  // 缺號：每個月份 1..max（max＝該月已知最大號與 quoteSeq 計數器取大）中，既不在報價單也不在流水帳的號碼
  const known = new Map();   // ym -> Set(serial)
  const add = (no) => { const m = NO_RE.exec(no || ''); if (!m) return; if (!known.has(m[1])) known.set(m[1], new Set()); known.get(m[1]).add(parseInt(m[2], 10)); };
  quotes.forEach((q) => q && add(q.quoteNo));
  ledgerOf(data).forEach((e) => add(e.no));
  const seqMap = (data.quoteSeq && typeof data.quoteSeq === 'object') ? data.quoteSeq : {};
  Object.keys(seqMap).forEach((ym) => { if (!known.has(ym)) known.set(ym, new Set()); });
  const months = [];
  let gapCount = 0;
  [...known.keys()].sort().forEach((ym) => {
    const set = known.get(ym);
    const maxKnown = set.size ? Math.max(...set) : 0;
    const max = Math.max(maxKnown, parseInt(seqMap[ym], 10) || 0);
    const gaps = [];
    for (let n = 1; n <= max; n++) if (!set.has(n)) gaps.push({ no: `QU-${ym}-${String(n).padStart(3, '0')}`, status: 'GAP_UNEXPLAINED' });
    gapCount += gaps.length;
    months.push({ ym, max, known: set.size, gaps });
  });
  const count = (s) => entries.filter((e) => e.status === s).length;
  return {
    changed, entries, months,
    summary: { total: entries.length, active: count('active'), invalid: count('invalid'), deleted: count('deleted'), missing: count('missing'), backfilled: entries.filter((e) => e.backfilled).length, gapUnexplained: gapCount },
  };
}

module.exports = { NO_RE, ledgerOf, deletedOf, findEntry, canonical, snapshotHash, ensureEntry, markCreated, markInvalid, markActive, recordDeletion, review };
