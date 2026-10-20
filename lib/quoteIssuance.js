'use strict';
/**
 * 報價單「正式文件發出機制」（內控稽核第三波 ①）。
 *
 * 問題：未核准的報價單可以匯出 Excel／PDF，只有前端 confirm 警告；蓋章 PDF 完全由瀏覽器產生，
 * 章圖可由任何看得到核准單的人直接取得；下載稽核由前端回報、少報或漏報不留痕；發出的文件沒有序號、沒有雜湊、
 * 也沒有和「核准的那個版本」連起來。
 *
 * 做法：每一份離開系統的報價文件（PDF／Excel）都由【伺服器】核發一筆發出紀錄（data.quoteIssuances[]，只增不改），
 * 內含發出序號、當下的內容雜湊、是否為核准有效版本、發出人與伺服器時間；報價專用章只給持有「蓋章憑證（stampToken）」的 PDF 流程，
 * 憑證只在 kind='official'（核准有效當下）時核發，綁定 單 + 發出序號 + 內容雜湊 + 使用者，5 分鐘有效、只能用一次。
 *
 * 資料：
 *   data.quoteIssuances[]  { issueNo, quoteId, quoteNo, kind:'official'|'draft'|'internal', format:'pdf'|'xlsx'|'pnl-xlsx',
 *                            contentHash, approvalHash|null, approvedValid, issuedBy, issuedByRole, issuedAt, ip, tokenJti|null, ver }
 *   data.quoteStampUses[]  { jti, issueNo, quoteId, by, at } —— 蓋章憑證的使用紀錄（單次使用判斷依據；超過憑證有效期的項目會被修剪，因為過期憑證本來就不會通過驗證）
 *   q.issuanceCount / q.issuances（最近 20 筆摘要）—— 只是快取，權威資料是 data.quoteIssuances；serialize 只多一個 issuanceCount（>0 才出現）
 *
 * kind='internal' 是毛利分析（內部 PNL）匯出：不是對外文件，不核發蓋章憑證，只留下「誰在何時帶走了含成本的檔案」。
 *
 * 設定：ALLOW_UNAPPROVED_ISSUE（預設 true＝與現況一致：未核准可以匯出，但一律記為「草稿」發出，文件上有草稿標示、PDF 不蓋章）。
 *   公司之後若要「未核准一律禁止匯出」，設環境變數 QUOTE_ALLOW_UNAPPROVED_ISSUE=0（不需改程式），未核准的發出要求會回 409 NOT_APPROVED。
 * 蓋章憑證簽章金鑰：環境變數 QUOTE_ISSUE_SECRET，否則由 JWT_SECRET／SESSION_SECRET 以獨立的網域標籤派生（不與 session、secretBox 共用同一把金鑰）。
 * 憑證絕不寫進日誌／稽核。
 */
const crypto = require('crypto');
const QA = require('./quoteApproval');

const RECORD_VER = 1;
const TOKEN_TTL_MS = 5 * 60 * 1000;
const SUMMARY_LAST = 20;
const KEY_LABEL = 'quote-issuance-stamp-token:v1:';
const FORMATS = ['pdf', 'xlsx', 'pnl-xlsx'];

/** 草稿浮水印開關（伺服器決定、經 POST /issue 的 watermark 欄位交給前端）：QUOTE_DRAFT_WATERMARK=1 才啟用（預設關）。控制 PDF 斜向浮水印與 Excel G4「草稿／未經核准」；關閉時草稿照樣記為 draft、照樣不蓋章，頁尾／G5 仍印序號＋雜湊＋「（草稿）／（正式）」 */
function draftWatermarkEnabled() {
  return /^(1|true|yes|on)$/i.test(String(process.env.QUOTE_DRAFT_WATERMARK || '').trim());
}

/** 未核准是否仍可發出（當作草稿）。getter 讀環境變數，方便不改程式就能收緊；預設 true（與現況一致） */
function allowUnapprovedIssue() {
  return !/^(0|false|no|off)$/i.test(String(process.env.QUOTE_ALLOW_UNAPPROVED_ISSUE == null ? 'true' : process.env.QUOTE_ALLOW_UNAPPROVED_ISSUE).trim());
}

// ── 工具 ──────────────────────────────────────────────
const b64u = (buf) => Buffer.from(buf).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const unb64u = (s) => Buffer.from(String(s).replace(/-/g, '+').replace(/_/g, '/'), 'base64');

function secretKey() {
  const dedicated = process.env.QUOTE_ISSUE_SECRET;
  const base = dedicated || process.env.JWT_SECRET || process.env.SESSION_SECRET;
  if (!base) throw new Error('[quoteIssuance] 缺 QUOTE_ISSUE_SECRET／JWT_SECRET／SESSION_SECRET，無法簽發蓋章憑證');
  return crypto.createHash('sha256').update(KEY_LABEL + base).digest();
}
function sign(payloadB64) { return crypto.createHmac('sha256', secretKey()).update(payloadB64).digest(); }

/** 目前這張單是否為「核准有效」：核准狀態且內容雜湊與核准當時相同（與 seal／export／export-log 的判斷同一式） */
function approvedValidOf(q) {
  const ap = q && q.approval;
  return !!ap && ap.state === 'approved' && QA.contentHash(q) === ap.hash;
}

/** 文件上印的短雜湊：sha256(發出序號 + ':' + 內容雜湊) 的前 12 個十六進位字元 */
function shortHashOf(issueNo, contentHash) {
  return crypto.createHash('sha256').update(String(issueNo) + ':' + String(contentHash)).digest('hex').slice(0, 12);
}

function ensureArrays(data) {
  if (!Array.isArray(data.quoteIssuances)) data.quoteIssuances = [];
  if (!Array.isArray(data.quoteStampUses)) data.quoteStampUses = [];
}

/**
 * 資安審查 L3：同人同單短時間內重複要求「同樣內容、同樣格式」的非正式 PDF 文件（xlsx／毛利分析／草稿 PDF）→ 重用既有那一筆，不再新增紀錄
 * （每筆約 350 bytes 進主資料 blob，迴圈打 /issue 可無限制灌大）。正式 PDF（kind=official 且 format=pdf）每次都要新的憑證 jti，所以不重用（改由限流把關）。
 * QUOTE_ISSUE_REUSE_SEC：重用視窗秒數，預設 60；0＝關閉重用（與改版前行為相同）。
 */
function reuseWindowMs() {
  const n = parseFloat(process.env.QUOTE_ISSUE_REUSE_SEC);
  return Number.isFinite(n) && n >= 0 ? n * 1000 : 60 * 1000;
}
/** 這個要求若新發出，會是哪一種 kind（與 issue() 同一套判斷；不核發、不改資料） */
function predictKind(q, format) {
  if (format === 'pnl-xlsx') return 'internal';
  if (approvedValidOf(q)) return 'official';
  return allowUnapprovedIssue() ? 'draft' : null;
}
/** 找可重用的紀錄（從尾端往前掃，超出視窗就停；紀錄是依時間附加的）。沒有就回 null */
function findReusable(data, p) {
  const win = reuseWindowMs();
  if (!win) return null;
  const list = data.quoteIssuances || [];
  const now = p.nowMs || Date.now();
  const kind = predictKind(p.q, p.format);
  if (!kind || (kind === 'official' && p.format === 'pdf')) return null;
  const hash = QA.contentHash(p.q);
  for (let i = list.length - 1; i >= 0; i--) {
    const r = list[i];
    if (!r) continue;
    const at = Date.parse(r.issuedAt);
    if (!(at > now - win)) break;
    if (r.quoteId === p.q.id && r.issuedBy === p.me && r.format === p.format && r.kind === kind && r.contentHash === hash) return r;
  }
  return null;
}

/** 這張單下一個序號（由既有紀錄推得，單調遞增；同一個同步流程內不會撞號） */
function nextSeq(data, quoteNo) {
  const prefix = 'ISS-' + quoteNo + '-';
  let mx = 0;
  for (const r of (data.quoteIssuances || [])) {
    if (r && typeof r.issueNo === 'string' && r.issueNo.startsWith(prefix)) {
      const n = parseInt(r.issueNo.slice(prefix.length), 10);
      if (n > mx) mx = n;
    }
  }
  return mx + 1;
}

// ── 蓋章憑證 ──────────────────────────────────────────
function signStampToken(rec, me, nowMs) {
  const payload = { q: rec.quoteId, i: rec.issueNo, h: rec.contentHash, u: me, j: rec.tokenJti, e: nowMs + TOKEN_TTL_MS };
  const p = b64u(JSON.stringify(payload));
  return p + '.' + b64u(sign(p));
}
/** 解析並驗簽（常數時間比對）。回 { ok:true, payload } 或 { ok:false, reason } */
function parseStampToken(token) {
  if (typeof token !== 'string' || token.length > 2048) return { ok: false, reason: 'malformed' };
  const parts = token.split('.');
  if (parts.length !== 2 || !parts[0] || !parts[1]) return { ok: false, reason: 'malformed' };
  let expect, got;
  try { expect = sign(parts[0]); got = unb64u(parts[1]); } catch (_) { return { ok: false, reason: 'malformed' }; }
  if (got.length !== expect.length || !crypto.timingSafeEqual(got, expect)) return { ok: false, reason: 'bad_signature' };
  let payload;
  try { payload = JSON.parse(unb64u(parts[0]).toString('utf8')); } catch (_) { return { ok: false, reason: 'malformed' }; }
  if (!payload || typeof payload !== 'object') return { ok: false, reason: 'malformed' };
  return { ok: true, payload };
}

/**
 * 核發一筆發出紀錄（push 進 data.quoteIssuances，同時更新該單的摘要快取）。呼叫端負責 db.save 與稽核。
 * @param {object} data   主資料 blob
 * @param {object} p      { q, format, me, role, ip, nowMs?, internal? }
 * @returns {{ok:true, rec, stampToken|null}|{ok:false, status, code, message}}
 */
function issue(data, p) {
  const q = p.q;
  if (!FORMATS.includes(p.format)) return { ok: false, status: 400, code: 'BAD_FORMAT', message: '不支援的格式' };
  ensureArrays(data);
  const nowMs = p.nowMs || Date.now();
  const contentHash = QA.contentHash(q);
  const valid = approvedValidOf(q);
  let kind;
  if (p.format === 'pnl-xlsx') kind = 'internal';
  else if (valid) kind = 'official';
  else {
    if (!allowUnapprovedIssue()) return { ok: false, status: 409, code: 'NOT_APPROVED', message: '這張報價單尚未核准（或核准已失效），目前不允許發出文件；請先完成簽核' };
    kind = 'draft';
  }
  const seq = nextSeq(data, q.quoteNo);
  const issueNo = 'ISS-' + q.quoteNo + '-' + String(seq).padStart(3, '0');
  const rec = {
    issueNo, quoteId: q.id, quoteNo: q.quoteNo, kind, format: p.format,
    contentHash, approvalHash: kind === 'official' ? q.approval.hash : null, approvedValid: valid,
    issuedBy: p.me, issuedByRole: p.role || '', issuedAt: new Date(nowMs).toISOString(), ip: String(p.ip || '').slice(0, 64),
    tokenJti: kind === 'official' && p.format === 'pdf' ? crypto.randomBytes(12).toString('hex') : null,
    ver: RECORD_VER,
  };
  data.quoteIssuances.push(rec);
  // 摘要快取（最近 20 筆）：不進 contentHash、不動 updatedAt
  const live = (data.quotations || []).find(x => x.id === q.id);
  if (live) {
    live.issuanceCount = (live.issuanceCount || 0) + 1;
    live.issuances = (live.issuances || []).concat([{ issueNo, kind, format: p.format, issuedAt: rec.issuedAt, issuedBy: p.me }]).slice(-SUMMARY_LAST);
  }
  const stampToken = rec.tokenJti ? signStampToken(rec, p.me, nowMs) : null;
  return { ok: true, rec, stampToken, shortHash: shortHashOf(issueNo, contentHash) };
}

/**
 * 用掉一張蓋章憑證（蓋章圖要送出前呼叫）。檢查：簽章、效期、單 id、使用者、發出紀錄存在且為 official/pdf、
 * 內容雜湊＝憑證＝目前單、核准仍有效、尚未使用過。通過就記一筆使用紀錄（單次使用）；呼叫端負責 db.save 與稽核。
 * @returns {{ok:true, rec}|{ok:false, status:403, code, reason}}
 */
function consumeStamp(data, q, token, me, nowMs) {
  const now = nowMs || Date.now();
  const bad = (reason) => ({ ok: false, status: 403, code: 'STAMP_TOKEN_INVALID', reason });
  const t = parseStampToken(token);
  if (!t.ok) return bad(t.reason);
  const pl = t.payload;
  if (!(typeof pl.e === 'number') || now > pl.e) return bad('expired');
  if (pl.q !== q.id) return bad('wrong_quote');
  if (pl.u !== me) return bad('wrong_user');
  ensureArrays(data);
  const rec = data.quoteIssuances.find(r => r.issueNo === pl.i);
  if (!rec || rec.quoteId !== q.id || rec.kind !== 'official' || rec.format !== 'pdf' || rec.tokenJti !== pl.j) return bad('no_issuance');
  const cur = QA.contentHash(q);
  if (pl.h !== cur || rec.contentHash !== cur || !approvedValidOf(q)) return bad('hash_mismatch');
  // 單次使用：修剪過期的使用紀錄（憑證過期後本來就驗不過，不需要再留）
  data.quoteStampUses = data.quoteStampUses.filter(u => u && now - Date.parse(u.at) < TOKEN_TTL_MS * 2);
  if (data.quoteStampUses.some(u => u.jti === pl.j)) return bad('used');
  data.quoteStampUses.push({ jti: pl.j, issueNo: rec.issueNo, quoteId: q.id, by: me, at: new Date(now).toISOString() });
  return { ok: true, rec };
}

// ── 查詢／驗證 ─────────────────────────────────────────
/** 這筆紀錄相對「現在的單」是否仍有效：內容雜湊相同，且（official）核准仍有效 */
function stillValidOf(rec, q) {
  if (!q) return false;
  if (rec.contentHash !== QA.contentHash(q)) return false;
  if (rec.kind === 'official') return approvedValidOf(q);
  return true;
}
function listForQuote(data, quoteId) {
  return (data.quoteIssuances || []).filter(r => r && r.quoteId === quoteId);
}

/**
 * 管理員驗證：一筆紀錄 vs 現在的單。回傳紀錄、各項比對與「不一致」清單。
 * opts.shortHash：文件上印的短雜湊，若給就一併核對是否與紀錄算出的相同（防偽造／抄錯序號）。
 */
function verify(data, issueNo, q, opts) {
  const rec = (data.quoteIssuances || []).find(r => r && r.issueNo === issueNo);
  if (!rec) return { found: false };
  const mismatches = [];
  const expectShort = shortHashOf(rec.issueNo, rec.contentHash);
  const out = {
    found: true, record: pub(rec), shortHash: expectShort, quoteFound: !!q,
    contentMatches: false, approvalValidNow: false, stillValid: false, mismatches,
  };
  if (opts && opts.shortHash != null && String(opts.shortHash).toLowerCase() !== expectShort) { out.hashMatches = false; mismatches.push('SHORT_HASH_DIFFERENT'); }
  else if (opts && opts.shortHash != null) out.hashMatches = true;
  // 紀錄本身的內部一致性：official 必須核准有效且 approvalHash＝contentHash；draft 不得帶 approvalHash／憑證
  if (rec.kind === 'official' && (!rec.approvedValid || rec.approvalHash !== rec.contentHash)) mismatches.push('OFFICIAL_RECORD_NOT_APPROVED');
  if (rec.kind === 'draft' && (rec.approvalHash || rec.tokenJti || rec.approvedValid)) mismatches.push('DRAFT_RECORD_HAS_APPROVAL');
  if (!q) { mismatches.push('QUOTE_NOT_FOUND'); return out; }
  out.contentMatches = rec.contentHash === QA.contentHash(q);
  out.approvalValidNow = approvedValidOf(q);
  out.stillValid = stillValidOf(rec, q);
  if (!out.contentMatches) mismatches.push('CONTENT_CHANGED');
  if (rec.kind === 'official' && out.contentMatches && !out.approvalValidNow) mismatches.push('OFFICIAL_NOW_NOT_APPROVED');
  if (rec.kind === 'official' && !out.contentMatches) mismatches.push('OFFICIAL_BUT_CONTENT_CHANGED');
  if (rec.kind === 'draft' && out.contentMatches && out.approvalValidNow) mismatches.push('DRAFT_NOW_APPROVED');   // 紀錄是草稿：這份文件當時未經核准，即使現在核准了也不是正式版
  return out;
}

/** 對外欄位（不含 ip、憑證 jti） */
function pub(rec) {
  return {
    issueNo: rec.issueNo, quoteId: rec.quoteId, quoteNo: rec.quoteNo, kind: rec.kind, format: rec.format,
    contentHash: rec.contentHash, approvalHash: rec.approvalHash, approvedValid: rec.approvedValid,
    issuedBy: rec.issuedBy, issuedByRole: rec.issuedByRole, issuedAt: rec.issuedAt, shortHash: shortHashOf(rec.issueNo, rec.contentHash), ver: rec.ver,
  };
}

module.exports = {
  RECORD_VER, TOKEN_TTL_MS, FORMATS, allowUnapprovedIssue, draftWatermarkEnabled,
  approvedValidOf, shortHashOf, nextSeq, reuseWindowMs, predictKind, findReusable, issue, consumeStamp, parseStampToken, stillValidOf, listForQuote, verify, pub,
};
