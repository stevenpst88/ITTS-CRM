'use strict';
/**
 * lib/mail/routes.js — 信件派送相關的 HTTP 路由（Cron 端點、寄件匣後台、使用者 Email 批次匯入）
 *
 * 匯出：registerMailRoutes(app, deps)
 *   deps.requireAdmin   server.js 的 requireAdmin（未登入 401、非管理員 403）
 *   deps.quoteMail      createQuoteMail() 的結果（含 config、outbox、dispatcher）
 *   deps.loadAuth / deps.saveAuth / deps.writeLog / deps.db   server.js 的同名函式（雲端 _auth 與本機 auth.json 由它們處理）
 *   deps.backend        'json' | 'postgres'（只用於後台設定頁顯示）
 *   deps.env            環境變數（預設 process.env；Cron 端點在每個請求讀 CRON_SECRET，不快取）
 *
 * 路由（server.js 必須在「強制改密碼」「集團角色路徑白名單」兩個 middleware 之後註冊）：
 *   GET  /api/cron/mail-outbox                       Vercel Cron（Hobby 每天一次）。requireAuth 之外，自己驗 Authorization: Bearer <CRON_SECRET>
 *                                                    （timingSafeEqual；CRON_SECRET 未設一律 401）。drainDue({limit:20, budgetMs:16000}) + purge()；回應只有計數，不含收件人。
 *                                                    全程有界：drainDue 整體時限（quoteMail.drainDue）、purge 與 db.flush 各 3 秒；卡住時放手並在回應標 timedOut:true
 *   GET  /api/admin/mail/outbox                      寄件匣列表（?status&type&quoteNo&toUser&since&limit&offset）＋ stats ＋ 熔斷狀態；只讀、無副作用
 *   POST /api/admin/mail/outbox/:id/requeue          立即重送：requeue → 清理一次 → 回傳該筆最新狀態；稽核 REQUEUE_MAIL
 *   GET  /api/admin/mail/config                      唯讀設定頁：publicConfig（不含 redirectTo 原值、密鑰）＋ warnings ＋ 缺 Email 清單 ＋ 執行期計數
 *   POST /api/admin/users/email-import/preview       批次匯入預覽（唯一回傳完整位址的地方；位址是管理員自己貼上的）
 *   POST /api/admin/users/email-import/apply         批次匯入套用（重新解析 text 再套用；稽核只寫摘要）
 *
 * 限制：
 *  - 本檔不 require npm 套件（lib/mail 規則）；express.static 之類的路由（/q/:id、/deep-link.js）在 server.js。
 *  - 回應內容一律是資料（不含 HTML）；顯示時的跳脫由前端負責（帳號名稱、錯誤列的 username／email 欄位可能含 < >）。
 *  - 寄件匣列表在 Postgres 上第一次呼叫會建表（outbox 的 CREATE TABLE IF NOT EXISTS）；MAIL_MODE=off 且從未啟用時，
 *    只要沒有人打開後台寄信頁，就不會碰儲存體。
 */

const crypto = require('crypto');
const { publicConfig } = require('./config');
const { parseBulkEmailText, applyBulkPlan, BULK_LIMITS } = require('./userEmail');
const { safeText } = require('./safety');
const { deadline } = require('./quoteMail');

const ID_RE = /^[A-Za-z0-9_-]{1,64}$/;
const CRON_DRAIN = Object.freeze({ limit: 20, budgetMs: 16000 });
const REQUEUE_DRAIN = Object.freeze({ limit: 5, budgetMs: 10000 });
const CRON_STEP_MS = 3000;       // Cron 裡 purge、db.flush 各自的時限（drainDue 另有整體時限）

function isObj(v) { return v !== null && typeof v === 'object' && !Array.isArray(v); }

function tally(list, key) {
  const out = {};
  (Array.isArray(list) ? list : []).forEach((x) => {
    const k = x && typeof x[key] === 'string' ? x[key] : 'UNKNOWN';
    out[k] = (out[k] || 0) + 1;
  });
  return out;
}

/** Drain 摘要 → 只含計數與代碼的版本（絕不含收件人帳號） */
function summaryCounts(s) {
  const r = isObj(s) ? s : {};
  const out = {
    sent: r.sent || 0,
    queued: r.queued || 0,
    cancelled: r.cancelled || 0,
    skipped: Array.isArray(r.skipped) ? r.skipped.length : 0,
    failed: Array.isArray(r.failed) ? r.failed.length : 0,
    errors: r.errors || 0,
    failedCodes: tally(r.failed, 'code'),
    skippedReasons: tally(r.skipped, 'reason'),
  };
  if (r.breakerOpen) out.breakerOpen = true;
  if (r.budgetExhausted) out.budgetExhausted = true;
  if (r.timedOut) out.timedOut = true;
  return out;
}

function cronAuthorized(req, env) {
  const secret = env && env.CRON_SECRET;
  if (typeof secret !== 'string' || secret === '') return false;          // 沒設定 CRON_SECRET：一律拒絕
  const h = req.headers && req.headers.authorization;
  if (typeof h !== 'string') return false;
  const m = /^Bearer\s+(\S+)\s*$/i.exec(h);
  if (!m) return false;
  // 先雜湊成等長再 timingSafeEqual（長度不同時 timingSafeEqual 會 throw，也不能因為長度洩漏資訊）
  const a = crypto.createHash('sha256').update(m[1]).digest();
  const b = crypto.createHash('sha256').update(secret).digest();
  return crypto.timingSafeEqual(a, b);
}

function operatorOf(req) { return (req.session && req.session.user && req.session.user.username) || 'unknown'; }

function registerMailRoutes(app, deps) {
  const { requireAdmin, quoteMail, loadAuth, saveAuth, writeLog, db } = deps;
  const env = deps.env || process.env;
  const config = quoteMail.config;
  const outbox = quoteMail.outbox;
  const flush = async () => { try { if (db && typeof db.flush === 'function') await deadline(Promise.resolve().then(() => db.flush()), CRON_STEP_MS); } catch (_) { /* 稽核寫入失敗不影響回應 */ } };
  const unavailable = (res) => res.status(503).json({ error: '寄信功能目前無法使用（寄件匣未初始化）', code: 'MAIL_UNAVAILABLE' });

  // ── Cron ──────────────────────────────────────────────────────────────
  app.get('/api/cron/mail-outbox', async (req, res) => {
    res.set('Cache-Control', 'no-store');
    if (!cronAuthorized(req, env)) return res.status(401).json({ error: '未授權' });
    try {
      if (!quoteMail.enabled()) {
        return res.json({ ok: true, enabled: false, mode: config.mode, sent: 0, queued: 0, cancelled: 0, skipped: 0, failed: 0, errors: 0, purged: 0 });
      }
      const drained = await quoteMail.drainDue(CRON_DRAIN);
      let purged = 0;
      const pr = await deadline(Promise.resolve().then(() => outbox.purge()), CRON_STEP_MS);
      if (pr.timedOut) console.warn('[mail cron] purge 超過 ' + CRON_STEP_MS + 'ms，先放手');
      else if (pr.error) console.warn('[mail cron] purge failed:', pr.error && pr.error.message);
      else purged = pr.value;
      await flush();
      return res.json(Object.assign({ ok: true, enabled: true, mode: config.mode, purged: Number.isFinite(purged) ? purged : 0 }, summaryCounts(drained)));
    } catch (e) {
      console.error('[mail cron]', e && e.message);
      return res.status(500).json({ ok: false, error: '寄件匣清理失敗' });
    }
  });

  // ── 後台：寄件匣 ──────────────────────────────────────────────────────
  app.get('/api/admin/mail/outbox', requireAdmin, async (req, res) => {
    if (!outbox) return unavailable(res);
    try {
      const q = req.query || {};
      const str = (v) => (typeof v === 'string' ? v : undefined);
      const num = (v) => { const n = Number(v); return typeof v === 'string' && v.trim() !== '' && Number.isFinite(n) ? n : undefined; };
      const since = str(q.since) && Number.isFinite(Date.parse(q.since)) ? q.since : undefined;
      const filter = { status: str(q.status), type: str(q.type), quoteNo: str(q.quoteNo), toUser: str(q.toUser), since, limit: num(q.limit), offset: num(q.offset) };
      const [list, stats, breaker] = await Promise.all([outbox.list(filter), outbox.stats(), outbox.breaker.state()]);
      res.set('Cache-Control', 'no-store');
      res.json({ available: true, mode: config.mode, rows: list.rows, total: list.total, limit: list.limit, offset: list.offset, stats, breaker });
    } catch (e) {
      console.error('[mail outbox list]', e && e.code, e && e.message);
      res.status(500).json({ error: '讀取寄件匣失敗', code: (e && e.code) || 'ERR' });
    }
  });

  app.post('/api/admin/mail/outbox/:id/requeue', requireAdmin, async (req, res) => {
    if (!outbox) return unavailable(res);
    const id = req.params.id;
    if (!ID_RE.test(id)) return res.status(400).json({ error: '寄信紀錄代碼格式不正確', code: 'BAD_ID' });
    try {
      const before = await outbox.get(id);
      if (!before) return res.status(404).json({ error: '找不到這筆寄信紀錄', code: 'NOT_FOUND' });
      const r = await outbox.requeue(id);
      if (!r.ok) {
        const status = r.reason === 'NOT_FOUND' ? 404 : 409;
        return res.status(status).json({ error: r.reason === 'NOT_FOUND' ? '找不到這筆寄信紀錄' : '這筆紀錄目前的狀態不能重送（只有失敗、略過、已取消的紀錄可以重送）', code: r.reason || 'BAD_STATE', status: r.status || before.status });
      }
      writeLog('REQUEUE_MAIL', operatorOf(req), safeText(before.quoteNo, 40), 'id=' + id + ' type=' + safeText(before.type, 40) + ' to=' + safeText(before.toUser, 64) + ' prev=' + before.status, req);
      // 立即重送：requeue 只是把工作放回 pending（nextAttemptAt＝現在），這裡馬上清理一次，管理員可以立刻看到結果。
      // 重送時會用單據現況重建內容並重新檢查：已過期的舊關卡信會被取消（cancelled／STALE），而不是誤寄。MAIL_MODE=off 時不清理（工作留在 pending）。
      let drain = null;
      if (quoteMail.enabled()) drain = summaryCounts(await quoteMail.drainDue(REQUEUE_DRAIN));
      const after = await outbox.get(id);
      res.json({ success: true, id, previousStatus: before.status, record: after, drain });
    } catch (e) {
      console.error('[mail requeue]', e && e.code, e && e.message);
      res.status(500).json({ error: '重送失敗', code: (e && e.code) || 'ERR' });
    }
  });

  // ── 後台：唯讀設定頁 ──────────────────────────────────────────────────
  app.get('/api/admin/mail/config', requireAdmin, (req, res) => {
    try {
      const pub = publicConfig(config);
      res.set('Cache-Control', 'no-store');
      res.json({
        config: pub,
        warnings: pub.warnings,
        enabled: quoteMail.enabled(),
        backend: deps.backend || 'json',
        cronSecretConfigured: typeof env.CRON_SECRET === 'string' && env.CRON_SECRET !== '',
        runtime: quoteMail.diagnostics(),
        missingEmails: quoteMail.missingEmails(),
      });
    } catch (e) {
      console.error('[mail config]', e && e.message);
      res.status(500).json({ error: '讀取寄信設定失敗' });
    }
  });

  // ── 使用者 Email 批次匯入 ─────────────────────────────────────────────
  app.post('/api/admin/users/email-import/preview', requireAdmin, (req, res) => {
    const text = req.body && req.body.text;
    if (typeof text !== 'string') return res.status(400).json({ error: '請提供要匯入的文字（text）', code: 'BAD_REQUEST' });
    const r = parseBulkEmailText(text, { config, users: loadAuth().users });
    res.json({ fatal: !!r.fatal, summary: r.summary, rows: r.rows, limits: BULK_LIMITS });
  });

  app.post('/api/admin/users/email-import/apply', requireAdmin, (req, res) => {
    const text = req.body && req.body.text;
    if (typeof text !== 'string') return res.status(400).json({ error: '請提供要匯入的文字（text）', code: 'BAD_REQUEST' });
    const auth = loadAuth();
    // 預覽與確認是兩個請求，帳號資料可能在中間被改過 → 這裡重新解析，再由 applyBulkPlan 對每一行重新驗證
    const parsed = parseBulkEmailText(text, { config, users: auth.users });
    if (parsed.fatal) {
      const first = parsed.rows[0] || {};
      return res.status(400).json({ error: first.error || '匯入內容不合法', code: first.code || 'BAD_REQUEST', fatal: true });
    }
    const plan = applyBulkPlan(parsed.rows, auth.users, { config });
    if (plan.error) return res.status(400).json({ error: plan.error, code: 'BAD_REQUEST' });
    if (plan.updated.length) {
      // 只改 email，且改在「原本的 user 物件」上（改密碼路由在 await 期間持有舊參照；雲端 loadAuth 是共用快取）
      const byName = new Map();
      auth.users.forEach((u) => { if (u && typeof u.username === 'string' && !byName.has(u.username)) byName.set(u.username, u); });
      plan.updated.forEach((u) => { const target = byName.get(u.username); if (target) target.email = u.email; });
      saveAuth(auth);
    }
    // 稽核只寫摘要與帳號名稱：完整位址（previous／email）只存在使用者自己貼上的文字與這次回應之外的資料庫帳號欄位
    const names = plan.updated.slice(0, 20).map((u) => safeText(u.username, 40)).join('、');
    const detail = 'email 批次匯入：更新 ' + plan.updated.length + ' 人' + (names ? '（' + names + (plan.updated.length > 20 ? '…' : '') + '）' : '') +
      '；無變更 ' + parsed.summary.unchanged + '；有錯誤未匯入 ' + parsed.summary.error + '；套用時略過 ' + plan.skipped.length;
    writeLog('IMPORT_USER_EMAILS', operatorOf(req), 'bulk', detail, req);
    res.json({
      success: true,
      updated: plan.updated.length,
      unchanged: parsed.summary.unchanged,
      errors: parsed.summary.error,
      skipped: plan.skipped,
      updatedUsers: plan.updated.map((u) => ({ username: u.username, detail: u.detail })),
    });
  });
}

module.exports = { registerMailRoutes, cronAuthorized, summaryCounts, CRON_DRAIN, REQUEUE_DRAIN };
