'use strict';
/**
 * 報價單簽核歷程（q.approval.history）的防竄改鏈。AUDIT_CHAIN=1 且有金鑰時，每筆新歷程多一個欄位 h：
 *    h = "<kid>:" + HMAC-SHA256(鏈金鑰, 前一筆的 h + "\n" + canonicalJSON(本筆除了 h 以外的欄位))
 * 沒有 h 的舊筆（鏈啟用前、或旗標關閉期間寫入的）一律容忍：驗證時 h 出現之前的筆算「舊資料」，之後缺 h 的筆計入 unchained（警示，不算斷鏈）。
 * 旗標關閉（預設）時 append() 的行為與原本的 push＋slice 位元級相同；h 與 historyBase 不會出現在前端序列化（serialize 是逐欄挑選）。
 *
 * 能抓：改任一筆內容、刪中間筆、調換順序、換 kid、用錯金鑰。
 * 抓不到：刪最後幾筆（沒有外部錨點）、把「所有」筆的 h 一起拿掉（等同回到舊資料；要靠稽核日誌 AUDIT 事件交叉比對）、同時拿到資料庫與伺服器金鑰的人。
 * 超過 MAX_HISTORY 被截掉最舊的幾筆時，把被截掉的最後一筆的 h 記在 ap.historyBase，驗證從它接續（不留驗證漏洞）。
 */
const AC = require('./auditChain');

const bodyOf = (e) => { const o = {}; for (const k of Object.keys(e)) if (k !== 'h') o[k] = e[k]; return o; };
const calcH = (key, kid, prevH, entry) => kid + ':' + AC.hmacHex(key, String(prevH) + '\n' + AC.canonical(bodyOf(entry)));

/** 為 history 最後一筆簽名（旗標關閉或沒有金鑰 → 什麼都不做）。失敗只記 log，不影響簽核流程。 */
function sign(ap, entry, cfg) {
  cfg = cfg || AC.getConfig(process.env);
  if (!cfg.enabled) return false;
  try {
    const hist = ap.history;
    let prevH = '';
    const idx = hist.lastIndexOf(entry);
    for (let i = (idx < 0 ? hist.length : idx) - 1; i >= 0; i--) if (hist[i] && typeof hist[i].h === 'string') { prevH = hist[i].h; break; }
    if (!prevH && ap.historyBase && !hist.slice(0, idx < 0 ? hist.length : idx).some((e) => e && e.h)) prevH = ap.historyBase;
    entry.h = calcH(cfg.keyFor(cfg.kid), cfg.kid, prevH, entry);
    return true;
  } catch (e) { console.error(AC.CHAIN_MARK.risk, 'history sign failed:', e && e.message); return false; }
}

/**
 * 加一筆歷程（取代原本的 ap.history.push(...) ＋ 超過上限就 slice）。
 * max 省略＝不截斷（quoteOppLink 原本就不截）。
 */
function append(ap, entry, max) {
  if (!Array.isArray(ap.history)) ap.history = [];
  ap.history.push(entry);
  sign(ap, entry);
  if (max && ap.history.length > max) {
    const dropped = ap.history.slice(0, ap.history.length - max);
    for (let i = dropped.length - 1; i >= 0; i--) if (dropped[i] && typeof dropped[i].h === 'string') { ap.historyBase = dropped[i].h; break; }
    ap.history = ap.history.slice(-max);
  }
  return entry;
}

/** 驗證一份歷程陣列 → { ok, entries, chained, legacy, unchained, firstBreak } */
function verifyHistory(history, cfg, base) {
  const res = { ok: true, entries: 0, chained: 0, legacy: 0, unchained: 0, firstBreak: null };
  if (!Array.isArray(history)) return res;
  res.entries = history.length;
  let prev = typeof base === 'string' ? base : '';
  let started = false;
  for (let i = 0; i < history.length; i++) {
    const e = history[i];
    if (!e || typeof e !== 'object') continue;
    if (typeof e.h !== 'string' || !e.h) { if (started) res.unchained++; else res.legacy++; continue; }
    started = true;
    const c = e.h.indexOf(':');
    const kid = c > 0 ? e.h.slice(0, c) : '';
    const key = kid ? cfg.keyFor(kid) : null;
    const brk = (reason) => { res.ok = false; res.firstBreak = { index: i, at: typeof e.at === 'string' ? e.at : '', action: typeof e.action === 'string' ? e.action.slice(0, 40) : '', reason }; };
    if (!key) { brk('UNKNOWN_KID'); break; }
    if (!AC.safeEq(calcH(key, kid, prev, e), e.h)) { brk('HASH_MISMATCH'); break; }
    prev = e.h; res.chained++;
  }
  return res;
}

function checkQuote(q, cfg) {
  cfg = cfg || AC.getConfig(process.env);
  if (!cfg.canVerify) return { ok: null, enabled: false, reason: cfg.disabledReason || 'NO_KEY', entries: 0, firstBreak: null };
  const ap = q && q.approval;
  const r = verifyHistory(ap && ap.history, cfg, ap && ap.historyBase);
  return Object.assign({ enabled: true, chainWriting: cfg.enabled }, r);
}

/** 全部報價單，每 200 張讓出一次事件迴圈（chunked）；最多列出 100 筆異常 */
async function checkAll(data, cfg, o) {
  cfg = cfg || AC.getConfig(process.env);
  o = o || {};
  if (!cfg.canVerify) return { ok: null, enabled: false, reason: cfg.disabledReason || 'NO_KEY', quotes: 0, entries: 0, breaks: [], firstBreak: null };
  const list = Array.isArray(data && data.quotations) ? data.quotations : [];
  const out = { ok: true, enabled: true, chainWriting: cfg.enabled, quotes: list.length, entries: 0, chained: 0, legacy: 0, unchained: 0, quotesWithUnchained: 0, breaks: [], firstBreak: null, ts: new Date().toISOString() };
  const every = o.yieldEvery || 200;
  for (let i = 0; i < list.length; i++) {
    const q = list[i];
    const ap = q && q.approval;
    if (ap && Array.isArray(ap.history)) {
      const r = verifyHistory(ap.history, cfg, ap.historyBase);
      out.entries += r.entries; out.chained += r.chained; out.legacy += r.legacy; out.unchained += r.unchained;
      if (r.unchained) out.quotesWithUnchained++;
      if (!r.ok) {
        out.ok = false;
        const b = Object.assign({ quoteNo: String(q.quoteNo || ''), id: String(q.id || '') }, r.firstBreak);
        if (!out.firstBreak) out.firstBreak = b;
        if (out.breaks.length < 100) out.breaks.push(b);
      }
    }
    if ((i + 1) % every === 0) await new Promise((res) => setImmediate(res));
  }
  return out;
}

module.exports = { sign, append, verifyHistory, checkQuote, checkAll, calcH };
