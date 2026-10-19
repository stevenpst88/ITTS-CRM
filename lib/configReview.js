'use strict';
/**
 * 簽核設定變更「兩人覆核」（dual review）。內控稽核第三波第 2 項。
 *
 * 開關：環境變數 APPROVAL_CONFIG_DUAL_REVIEW=1（或 true）才啟用，預設關閉；只能由環境變數控制（畫面上無法切換，
 * 所以單一管理員無法自己關掉）。關閉時 PUT 立即生效、回應與以前逐位元相同（本模組完全不介入）。
 *
 * 啟用時：管理員的 PUT 不生效，改成建立一筆「變更申請」（data.approvalChangeRequests[]，只增不刪、狀態轉換記在 history）。
 * 另一位管理員核可時，伺服器重新驗證（同一套驗證器）＋比對 baseHash，然後在「同一次 db.save」內套用變更並把申請標成 approved。
 *   · 提案人永遠不能核可／退回自己的申請（403 SELF_REVIEW_FORBIDDEN）；只能撤回自己的 pending 申請。
 *   · baseHash 與現況不同 → 申請標成 stale，409 STALE_BASE（請提案人重新送出）。
 *   · 3 天未覆核 → expired（惰性：列表／讀取／核可時才結算，沒有排程）。
 *   · 在職管理員少於 2 人 → PUT 回 409 DUAL_REVIEW_NEEDS_TWO_ADMINS（本版刻意不做「緊急單人繞過」）。
 *
 * 可擴充：registerKind(kind, spec) 是「可覆核的設定種類」登錄表。目前登錄 'config'（名冊＋商品歸類）與 'flow'（簽核流程設定），
 * 登錄在 lib/quoteRoutes.js。日後要把 quoteMode 等也納入覆核：在該處多登錄一個 kind，並在對應的 PUT 裡走同一個 propose() 即可。
 *   spec = {
 *     label,                       // 中文名稱（通知／畫面用）
 *     rawKeys: [...],              // apply 會整個取代的 data.* 鍵（失敗時用來還原）
 *     read(data) → 子樹            // 目前儲存的「相關設定」（正規化；hash 與 before 用）
 *     validate(afterInput, {users, data}) → { ok:true, value } | { ok:false, status, code, error }
 *     diff(before, after, {users}) → string[]   // 中文可讀差異
 *     apply(data, value, {by, reviewer, requestId, req}) → { action, target, detail }   // 只改記憶體，不 save、不 writeLog
 *   }
 */
const crypto = require('crypto');
const AD = require('./auditDiff');

const ENV_NAME = 'APPROVAL_CONFIG_DUAL_REVIEW';
const EXPIRY_MS = 3 * 24 * 60 * 60 * 1000;
const MAX_SUMMARY_LINES = 200;
const MAX_LIST = 200;

function isEnabled(env) {
  const v = String(((env || process.env)[ENV_NAME]) || '').trim().toLowerCase();
  return v === '1' || v === 'true';
}

/** 穩定序列化（鍵排序）→ sha256。用來比對「提案當時的設定」與「核可當下的設定」是否相同 */
function canonical(v) {
  if (Array.isArray(v)) return '[' + v.map(canonical).join(',') + ']';
  if (v && typeof v === 'object') return '{' + Object.keys(v).sort().map((k) => JSON.stringify(k) + ':' + canonical(v[k])).join(',') + '}';
  return JSON.stringify(v === undefined ? null : v);
}
const hashOf = (v) => crypto.createHash('sha256').update(canonical(v)).digest('hex');

function create(deps) {
  const { db, writeLog, pushNotification, loadAuth } = deps;
  const nowMs = deps.now || (() => Date.now());
  const envOf = () => deps.env || process.env;
  const kinds = {};

  const enabled = () => isEnabled(envOf());
  function registerKind(kind, spec) { kinds[kind] = spec; }
  const kindLabel = (k) => (kinds[k] && kinds[k].label) || k;
  const fail = (res, status, code, error, extra) => res.status(status).json(Object.assign({ error, code }, extra || {}));
  const isoNow = () => new Date(nowMs()).toISOString();

  function activeAdmins() {
    const auth = loadAuth();
    return ((auth && auth.users) || []).filter((u) => u && u.role === 'admin' && u.active !== false).map((u) => u.username);
  }
  const labelOf = (users, un) => (users && users[un] && (users[un].nickname || users[un].displayName || users[un].username)) || un || '';

  function store(data) {
    if (!Array.isArray(data.approvalChangeRequests)) data.approvalChangeRequests = [];
    return data.approvalChangeRequests;
  }

  /** 逾期結算（惰性）：pending 且 expiresAt 已過 → expired。回傳這次新結算的申請（呼叫端負責 save／稽核／通知） */
  function sweep(data) {
    const out = [];
    const t = nowMs();
    (Array.isArray(data.approvalChangeRequests) ? data.approvalChangeRequests : []).forEach((r) => {
      if (r.status === 'pending' && Date.parse(r.expiresAt) <= t) {
        r.status = 'expired';
        r.history.push({ at: isoNow(), by: 'system', action: 'expire', comment: '逾期未覆核' });
        out.push(r);
      }
    });
    return out;
  }

  function logStep(action, operator, rec, extra, req) {
    const lines = AD.joinCapped(rec.diffSummary || [], { prefix: `單號 ${rec.id}｜${kindLabel(rec.kind)}｜提案人 ${rec.proposedBy}${extra ? '｜' + extra : ''}｜異動：`, max: AD.MAX_DETAIL });
    writeLog(action, operator, 'approvalConfig:' + rec.kind, lines, req);
  }
  function notifyTo(list, type, title, body, refId) {
    new Set(list).forEach((un) => { try { pushNotification(un, type, title, body, refId); } catch (e) { console.warn('[configReview notify]', e.message); } });
  }
  /** 提案人以外的在職管理員 */
  const otherAdmins = (proposer) => activeAdmins().filter((u) => u !== proposer);

  /** 結算逾期並（若有）存檔＋稽核＋通知提案人。非同步 flush：GET 不會經過全域寫入 flush middleware */
  async function settleExpired(data, req) {
    const expired = sweep(data);
    if (!expired.length) return 0;
    db.save(data);
    expired.forEach((r) => {
      logStep('CONFIG_CHANGE_EXPIRED', 'system', r, '逾期（3 天內無人覆核）', req);
      notifyTo([r.proposedBy], 'config_review_expired', '簽核設定變更申請已逾期', `你提出的「${kindLabel(r.kind)}」變更逾期未覆核，已失效，需要的話請重新提出。`, r.id);
    });
    try { if (typeof db.flush === 'function') await db.flush(); } catch (_) { /* 盡力而為 */ }
    return expired.length;
  }

  // ── 提案（由 PUT handler 在「驗證通過、開關為 ON」時呼叫）──────────────────────
  function propose(res, kind, { data, value, proposer, users, note, req }) {
    const spec = kinds[kind];
    if (!spec) return fail(res, 500, 'UNKNOWN_KIND', '未知的設定種類');
    if (activeAdmins().length < 2) {
      return fail(res, 409, 'DUAL_REVIEW_NEEDS_TWO_ADMINS',
        '「兩人覆核」已啟用，但目前在職的管理員不足 2 人，變更無法送審。請先建立第二位管理員帳號，或由 IT 暫時關閉環境變數 APPROVAL_CONFIG_DUAL_REVIEW。');
    }
    const before = spec.read(data);
    if (hashOf(before) === hashOf(value)) {
      return res.status(200).json({ success: true, noChange: true, message: '設定沒有變更，不需覆核' });
    }
    const t = nowMs();
    const lines = spec.diff(before, value, { users }).slice();
    if (lines.length > MAX_SUMMARY_LINES) { const n = lines.length - MAX_SUMMARY_LINES; lines.length = MAX_SUMMARY_LINES; lines.push(`…(另有 ${n} 項異動未列出，請看完整內容)`); }
    const cleanNote = String(note == null ? '' : note).replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, 200);
    const rec = {
      id: crypto.randomUUID(), kind, proposedBy: proposer, proposedAt: new Date(t).toISOString(),
      baseHash: hashOf(before), before, after: value, diffSummary: lines.length ? lines : ['（內容有變，但沒有可列出的差異）'],
      status: 'pending', history: [{ at: new Date(t).toISOString(), by: proposer, action: 'propose', comment: cleanNote }],
      expiresAt: new Date(t + EXPIRY_MS).toISOString(),
    };
    store(data).push(rec);
    db.save(data);
    logStep('CONFIG_CHANGE_PROPOSED', proposer, rec, '', req);
    notifyTo(otherAdmins(proposer), 'config_review_requested', '有簽核設定變更待覆核', `${labelOf(users, proposer)} 提出「${kindLabel(kind)}」變更，請至後台覆核。`, rec.id);
    return res.status(202).json({ requestId: rec.id, status: 'pending', message: '已提出，等待另一位管理員覆核', expiresAt: rec.expiresAt });
  }

  // ── 序列化 ────────────────────────────────────────────────
  function summarize(rec, me, data, withBody) {
    const last = rec.history[rec.history.length - 1] || {};
    const spec = kinds[rec.kind];
    let baseChanged = false;
    try { baseChanged = rec.status === 'pending' && !!spec && hashOf(spec.read(data)) !== rec.baseHash; } catch (_) { baseChanged = false; }
    const out = {
      id: rec.id, kind: rec.kind, kindLabel: kindLabel(rec.kind), status: rec.status, proposedBy: rec.proposedBy, proposedAt: rec.proposedAt,
      expiresAt: rec.expiresAt, diffSummary: rec.diffSummary, history: rec.history,
      decidedBy: rec.status === 'pending' ? null : (last.by || null), decidedAt: rec.status === 'pending' ? null : (last.at || null),
      isMine: rec.proposedBy === me, baseChanged,
    };
    if (withBody) { out.before = rec.before; out.after = rec.after; }
    return out;
  }

  // ── 路由 ─────────────────────────────────────────────────
  function registerRoutes(app, guards, loadCtx) {
    const BASE = '/api/admin/approval-config-requests';

    app.get(BASE, guards, async (req, res) => {
      const ctx = loadCtx(req);
      const data = db.load();
      await settleExpired(data, req);
      const st = String(req.query.status || '').trim();
      const kd = String(req.query.kind || '').trim();
      let list = store(data).slice();
      if (st) list = list.filter((r) => r.status === st);
      if (kd) list = list.filter((r) => r.kind === kd);
      list.sort((a, b) => (a.proposedAt < b.proposedAt ? 1 : -1));
      const pendingAll = store(data).filter((r) => r.status === 'pending').length;
      res.json({
        enabled: enabled(), activeAdmins: activeAdmins().length, me: ctx.me, pendingCount: pendingAll,
        requests: list.slice(0, MAX_LIST).map((r) => summarize(r, ctx.me, data, false)),
      });
    });

    app.get(BASE + '/:id', guards, async (req, res) => {
      const ctx = loadCtx(req);
      const data = db.load();
      await settleExpired(data, req);
      const rec = store(data).find((r) => r.id === req.params.id);
      if (!rec) return fail(res, 404, 'NOT_FOUND', '找不到這筆變更申請');
      res.json(summarize(rec, ctx.me, data, true));
    });

    // 取出申請＋共通檢查；回 null 代表已回應
    async function loadForAction(req, res) {
      const ctx = loadCtx(req);
      if (ctx.role !== 'admin') { fail(res, 403, 'NO_PERMISSION', '只有管理員可以覆核'); return null; }
      const data = db.load();
      const rec = store(data).find((r) => r.id === req.params.id);
      if (!rec) { fail(res, 404, 'NOT_FOUND', '找不到這筆變更申請'); return null; }
      return { ctx, data, rec };
    }
    const notPending = async (res, data, rec, req) => {
      await settleExpired(data, req);
      if (rec.status === 'expired') return fail(res, 409, 'EXPIRED', '這筆變更申請已逾期失效，請提案人重新提出');
      return fail(res, 409, 'NOT_PENDING', `這筆變更申請目前狀態為 ${rec.status}，無法再處理`, { status: rec.status });
    };

    app.post(BASE + '/:id/approve', guards, async (req, res) => {
      const x = await loadForAction(req, res); if (!x) return;
      const { ctx, data, rec } = x;
      if (rec.proposedBy === ctx.me) return fail(res, 403, 'SELF_REVIEW_FORBIDDEN', '不能覆核自己提出的變更，請由另一位管理員核可');
      await settleExpired(data, req);
      if (rec.status !== 'pending') return notPending(res, data, rec, req);
      const spec = kinds[rec.kind];
      if (!spec) return fail(res, 500, 'UNKNOWN_KIND', '未知的設定種類');
      const comment = String((req.body && req.body.comment) || '').replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, 200);
      // 1) 基準比對：提案之後設定若已被改過，這份「前後差異」就不準了
      if (hashOf(spec.read(data)) !== rec.baseHash) {
        rec.status = 'stale';
        rec.history.push({ at: isoNow(), by: ctx.me, action: 'stale', comment: '提案之後設定已被其他變更改動' });
        db.save(data);
        logStep('CONFIG_CHANGE_STALE', ctx.me, rec, '基準已變動', req);
        notifyTo([rec.proposedBy], 'config_review_stale', '簽核設定變更申請已過期（基準變動）', `你提出的「${kindLabel(rec.kind)}」變更因設定已被改動而無法套用，請重新提出。`, rec.id);
        return fail(res, 409, 'STALE_BASE', '提案之後簽核設定已被其他變更改動，這份申請已作廢（stale）。請提案人依最新設定重新提出。', { status: 'stale', requestId: rec.id });
      }
      // 2) 重新驗證（同一套驗證器；帳號可能在提案後被停用／刪除／改成管理員）
      const users = {}; ((loadAuth().users) || []).forEach((u) => { users[u.username] = u; });
      const v = spec.validate(rec.after, { users, data });
      if (!v.ok) return fail(res, v.status || 400, v.code || 'BAD_CONFIG', v.error || '內容驗證未通過', { requestId: rec.id });
      // 3) 同一次 save：套用＋標記 approved；失敗就還原記憶體（Postgres 後端 db.load() 是共用快取）
      const backups = (spec.rawKeys || []).map((k) => [k, data[k]]);
      const histLen = rec.history.length;
      let info;
      try {
        rec.status = 'approved';
        rec.history.push({ at: isoNow(), by: ctx.me, action: 'approve', comment });
        info = spec.apply(data, v.value, { by: rec.proposedBy, reviewer: ctx.me, requestId: rec.id, req });
        db.save(data);
      } catch (e) {
        backups.forEach(([k, val]) => { data[k] = val; });
        rec.status = 'pending'; rec.history.length = histLen;
        console.error('[configReview apply]', e && e.message);
        return fail(res, 500, 'APPLY_FAILED', '套用失敗，變更未生效，申請維持待覆核', { requestId: rec.id });
      }
      writeLog(info.action, ctx.me, info.target, `${info.detail}｜提案人 ${rec.proposedBy}｜覆核人 ${ctx.me}｜單號 ${rec.id}`, req);
      logStep('CONFIG_CHANGE_APPROVED', ctx.me, rec, `覆核人 ${ctx.me}`, req);
      notifyTo([rec.proposedBy], 'config_review_approved', '簽核設定變更已核可並生效', `你提出的「${kindLabel(rec.kind)}」變更已由 ${labelOf(users, ctx.me)} 核可並生效。`, rec.id);
      res.json({ success: true, requestId: rec.id, status: 'approved' });
    });

    app.post(BASE + '/:id/reject', guards, async (req, res) => {
      const x = await loadForAction(req, res); if (!x) return;
      const { ctx, data, rec } = x;
      if (rec.proposedBy === ctx.me) return fail(res, 403, 'SELF_REVIEW_FORBIDDEN', '不能覆核自己提出的變更，請由另一位管理員處理');
      const comment = String((req.body && req.body.comment) || '').replace(/[\u0000-\u001f\u007f]/g, ' ').trim().slice(0, 200);
      if (Array.from(comment).length < 2) return fail(res, 400, 'COMMENT_REQUIRED', '退回必須填寫原因（至少 2 個字）');
      await settleExpired(data, req);
      if (rec.status !== 'pending') return notPending(res, data, rec, req);
      rec.status = 'rejected';
      rec.history.push({ at: isoNow(), by: ctx.me, action: 'reject', comment });
      db.save(data);
      logStep('CONFIG_CHANGE_REJECTED', ctx.me, rec, `退回原因：${comment}`, req);
      const users = {}; ((loadAuth().users) || []).forEach((u) => { users[u.username] = u; });
      notifyTo([rec.proposedBy], 'config_review_rejected', '簽核設定變更被退回', `你提出的「${kindLabel(rec.kind)}」變更被 ${labelOf(users, ctx.me)} 退回，請至後台查看原因。`, rec.id);
      res.json({ success: true, requestId: rec.id, status: 'rejected' });
    });

    app.post(BASE + '/:id/cancel', guards, async (req, res) => {
      const x = await loadForAction(req, res); if (!x) return;
      const { ctx, data, rec } = x;
      if (rec.proposedBy !== ctx.me) return fail(res, 403, 'NOT_PROPOSER', '只有提案人可以撤回自己的申請');
      await settleExpired(data, req);
      if (rec.status !== 'pending') return notPending(res, data, rec, req);
      rec.status = 'cancelled';
      rec.history.push({ at: isoNow(), by: ctx.me, action: 'cancel', comment: '' });
      db.save(data);
      logStep('CONFIG_CHANGE_CANCELLED', ctx.me, rec, '', req);
      res.json({ success: true, requestId: rec.id, status: 'cancelled' });
    });
  }

  return { enabled, registerKind, propose, registerRoutes, sweep, settleExpired, activeAdmins, hashOf, summarize };
}

module.exports = { ENV_NAME, EXPIRY_MS, isEnabled, canonical, hashOf, create };
