'use strict';
/**
 * MA／License 到期「活規則」提醒的 HTTP 端點、掃描包裝與 poll-bundle 動態鈴鐺（P3a）。純邏輯在 lib/maReminders.js、收件人解析在 lib/maRecipients.js。
 * MA_TRACKING=1 才啟用；關閉時所有端點 next('route')（與未登記相同＝404），poll／sweep 直接回 null。
 *
 * 觸發（不新增 vercel.json cron——Hobby 方案 cron 數量有限）：
 *   ① 既有的 GET /api/cron/mail-outbox（每日一次，UTC 00:00＝台北 08:00）處理完信件後呼叫 sweep({mode:'cron'})（lib/mail/routes.js 的 deps.maSweep）。
 *   ② GET /api/poll-bundle 的補掃：任何登入者的輪詢都可能觸發 sweep({mode:'lazy'})（每個實例至少間隔 60 分鐘；今天還沒掃過則盡快補）。
 *   兩者共用同一個函式與去重鍵，冪等；掃描失敗不影響原本的回應。
 *
 * server.js：  const MRR = require('./lib/maReminderRoutes');
 *              poll-bundle：  const maReminders = MRR.poll(req);   → 陣列或 null（null＝功能關閉，回應不多任何鍵）
 *              檔尾：  MRR.register(app, deps)
 */
const crypto = require('crypto');
const MA = require('./maTracking');
const MRM = require('./maReminders');
const OR = require('./orgResolve');
const OL = require('./quoteOppLink');

let D = null;
let _running = false;
let _instLastAttemptMs = 0;

const maOn = (req, res, next) => (MA.enabled() ? next() : next('route'));
const err = (res, status, code, message, extra) => res.status(status).json(Object.assign({ error: message, code }, extra || {}));
const nowIso = () => new Date().toISOString();
const usersMap = () => { const m = Object.create(null); ((D.loadAuth().users) || []).forEach((u) => { if (u && u.username) m[u.username] = u; }); return m; };
const visibleItems = (req, data) => require('./maRoutes').visibleItems(req, data);

function setCfg(data, value, by, extra) {
  data.reminderRules = Object.assign({ rules: value.rules, updatedAt: nowIso(), updatedBy: by }, extra || {});
}
function cfgLogDetail(prev, value) {
  const lines = MRM.configDiff(prev, value);
  return '提醒規則｜共 ' + value.rules.length + ' 條（啟用 ' + value.rules.filter((r) => r.enabled).length + '）' + (lines.length ? '｜異動：' + lines.join('；') : '｜無異動');
}

// ───────────────── 掃描（cron／補掃共用）─────────────────
/** 同步核心：載入 → 掃描 → 一次 save（去重標記與通知同一次寫入）→ 稽核 → Web Push。不 await 任何事。 */
function sweepNow(opts) {
  const o = opts || {};
  const data = D.db.load();
  const users = (D.loadAuth().users) || [];
  const r = MRM.sweep(data, users, { now: o.now === undefined ? Date.now() : o.now, mode: o.mode || 'cron', dryRun: !!o.dryRun, rules: MRM.getConfig(data).rules });
  if (o.dryRun) return { stats: r.stats, today: r.today, dryRun: true };
  if (r.changed) D.db.save(data);
  const noisy = r.stats.fired || r.stats.resolved || r.stats.noRecipient || r.stats.capped;
  try {
    r.audits.forEach((a) => D.writeLog(a.action, 'system', a.target, a.detail, null));
    if (o.mode === 'cron' || noisy) D.writeLog('MA_REMINDER_SWEEP', 'system', 'maReminders', `模式 ${o.mode || 'cron'}｜評估 ${r.stats.evaluated} 項｜發出 ${r.stats.fired} 筆提醒給 ${r.stats.recipients} 人｜已發過略過 ${r.stats.skippedDup}｜無收件人 ${r.stats.noRecipient}｜超過單次上限未發 ${r.stats.capped}｜自動標為已讀 ${r.stats.resolved || 0}`, null);
  } catch (e) { console.warn('[ma reminder audit]', e && e.message); }
  // Web Push（fire-and-forget，已 save 之後）：失敗不影響
  r.created.forEach((n) => { try { if (D.webPush) D.webPush(n.to, n.type, n.title, n.body, n.refId); } catch (e) { console.warn('[ma reminder push]', e && e.message); } });
  return { stats: r.stats, today: r.today, notifications: r.created.length };
}

/**
 * 對外的掃描：回傳統計 { evaluated, fired, recipients, skippedDup, noRecipient, capped }（功能關閉回 null；正在掃描中回 { skipped:true }）。
 * mode 'cron'：先 await db.ready()（Postgres 取最新快照）；掃描後 await db.flush()（等背景寫入落地）。mode 'lazy'：不 await 任何事直到掃描完成（同步段落先跑完，flush 在背景）。
 */
async function sweepReminders(opts) {
  if (!MA.enabled()) return null;
  const o = opts || {};
  if (_running) return { skipped: true };
  _running = true;
  try {
    if ((o.mode || 'cron') === 'cron' && D.db.ready) await D.db.ready();
    const out = sweepNow(o);
    if (!o.dryRun && D.db.flush) await D.db.flush();
    return Object.assign({}, out.stats, { today: out.today, notifications: out.notifications || 0 });
  } finally { _running = false; }
}

function lazyDue(data) {
  const log = (data && data.reminderLog) || {};
  return MRM.shouldLazySweep(Date.now(), { instLastAttemptMs: _instLastAttemptMs, lastSweepAt: log.lastSweepAt, lastSweepDate: log.lastSweepDate });
}

/**
 * poll-bundle 的補掃 middleware（放在 requireAuth 之後、處理器之前，與 mailPollDrain 同一個位置）：需要補掃時先 await 掃描（含 db.flush，最多等 4 秒）
 * 再放行，所以回應送出前資料已落地（Vercel 回應送出後行程可能被凍結，背景寫入會被切斷）。不需要時（含功能關閉）立刻 next()，零成本。任何失敗都放行。
 */
function lazyMiddleware(req, res, next) {
  if (!MA.enabled() || !D || _running) return next();
  let due = false;
  try { due = lazyDue(D.db.load()); } catch (e) { due = false; }
  if (!due) return next();
  _instLastAttemptMs = Date.now();
  let passed = false;
  const go = () => { if (!passed) { passed = true; next(); } };
  const timer = setTimeout(go, 4000);
  sweepReminders({ mode: 'lazy' }).catch((e) => console.warn('[ma lazy sweep]', e && e.message)).then(() => { clearTimeout(timer); go(); });
}

// ───────────────── poll-bundle 的動態鈴鐺 ─────────────────
/**
 * 目前登入者此刻的動態提醒（陣列）。功能關閉回 null（poll-bundle 就不多 maReminders 這個鍵）。
 * 順便：把「條件已不成立」的自己的 ma_reminder 通知標已讀（僅此一種寫入，無變化不 save）。補掃由 lazyMiddleware 在它之前處理。
 * 任何例外都吞掉並回 []：鈴鐺不能因為這個功能壞掉。
 */
function poll(req) {
  if (!MA.enabled() || !D) return null;
  try {
    const me = req.session.user.username;
    const role = req.session.user.role;
    const data = D.db.load();
    const users = (D.loadAuth().users) || [];
    const now = Date.now();
    const rules = MRM.getConfig(data).rules;
    const marked = MRM.reconcileNotifications(data, users, now, rules, (n) => n.to === me);
    if (marked) D.db.save(data);
    const pl = MRM.plan(data, users, now, rules);
    let vis = null;
    const rows = MRM.dynamicFor(pl, me, {
      secretary: role === 'secretary', labelOf: (u) => MA.labelOf(users, u), limit: 200,
      visible: (item) => { if (!vis) vis = new Set(visibleItems(req, data).map((x) => x.id)); return vis.has(item.id); },
    });
    return rows;
  } catch (e) {
    console.warn('[ma poll]', e && e.message);
    return [];
  }
}

// ───────────────── 報價單 renewOf 驗證（由 lib/quoteRoutes.js 的 POST／PUT 呼叫）─────────────────
/**
 * id 空 → { ok:true, value:'' }。必須是呼叫者可見範圍內（與 ma-items 同一套）的納管項、未作廢、與報價單同公司。
 * 回 { ok:true, value, item } 或 { ok:false, status, code, message }。找不到與看不到一律同一個 404 訊息（不洩漏他人納管項是否存在）。
 */
function checkRenewOf(req, data, id, company) {
  if (id === undefined || id === null || id === '') return { ok: true, value: '' };
  if (typeof id !== 'string' || id.length > 80) return { ok: false, status: 400, code: 'BAD_RENEW_OF', message: 'renewOf 必須是納管項代號（字串）' };
  const it = visibleItems(req, data).find((x) => x && x.id === id);
  if (!it) return { ok: false, status: 404, code: 'RENEW_ITEM_NOT_FOUND', message: '找不到可以續約的納管項（不存在，或不在你的可視範圍內）' };
  if (it.status === 'void') return { ok: false, status: 409, code: 'RENEW_ITEM_VOID', message: '這個納管項已作廢，不能指定為續約對象' };
  if (!OL.companyEq(it.company, company)) return { ok: false, status: 400, code: 'RENEW_COMPANY_MISMATCH', message: '續約的納管項與報價單不是同一個客戶' };
  return { ok: true, value: id, item: it };
}

/**
 * 通知清單輸出前處理：秘書看不到納管項代號（到期管理對秘書沒有 id 欄位），所以 ma_reminder 通知上的 maItemIds 對秘書拿掉。
 * 功能關閉、或不是秘書 → 原陣列原樣回傳（位元級不變）。
 */
function scrub(role, list) {
  if (!MA.enabled() || role !== 'secretary' || !Array.isArray(list)) return list;
  return list.map((n) => (n && n.type === MRM.NOTIF_TYPE && 'maItemIds' in n ? Object.fromEntries(Object.entries(n).filter(([k]) => k !== 'maItemIds')) : n));
}

// ───────────────── 路由 ─────────────────
function register(app, deps) {
  D = deps;
  const { requireAuth, requireAdmin, db, loadAuth, writeLog, pushNotification } = deps;
  const CR = deps.configReview || null;
  const isOwnerOrAdmin = (req, owner) => req.session.user.role === 'admin' || req.session.user.username === owner;
  const labelOf = (users, u) => MA.labelOf(users, u);

  // 兩人覆核登錄（APPROVAL_CONFIG_DUAL_REVIEW=1 時 PUT 走覆核；預設關閉＝立即生效，與防拆單設定一致）
  if (CR && typeof CR.registerKind === 'function') {
    CR.registerKind('reminder_rules', {
      label: '到期提醒規則', rawKeys: ['reminderRules'],
      read: (data) => { const c = MRM.getConfig(data); return { rules: c.rules }; },
      validate: (after) => MRM.validateConfig(after),
      diff: (before, after) => MRM.configDiff(before, after),
      apply: (data, value, m) => {
        const prev = MRM.getConfig(data);
        setCfg(data, value, m.by, { reviewedBy: m.reviewer, changeRequestId: m.requestId });
        return { action: 'SAVE_REMINDER_RULES', target: 'reminderRules', detail: cfgLogDetail(prev, value) };
      },
    });
  }

  app.get('/api/admin/reminder-rules', maOn, requireAdmin, (req, res) => {
    const data = db.load();
    const c = MRM.getConfig(data);
    res.set('Cache-Control', 'no-store');
    res.json({
      enabled: true,
      config: { rules: c.rules, updatedAt: c.updatedAt, updatedBy: c.updatedBy, isDefault: c.isDefault },
      defaults: MRM.DEFAULT_RULES, limits: MRM.LIMITS,
      options: { conditions: MRM.CONDITIONS, notify: MRM.NOTIFY_ROLES, kinds: MRM.KINDS },
    });
  });

  app.put('/api/admin/reminder-rules', maOn, requireAdmin, (req, res) => {
    const me = req.session.user.username;
    const b = req.body || {};
    const v = MRM.validateConfig(b);
    if (!v.ok) return err(res, v.status, v.code, v.error);
    const data = db.load();
    v.value.rules = MRM.adoptIds(b.rules, v.value.rules, MRM.getConfig(data).rules);   // 沒帶代號的規則沿用現行同條件的代號（避免去重鍵換新、提醒重發）
    if (CR && CR.enabled()) return CR.propose(res, 'reminder_rules', { data, value: v.value, proposer: me, users: usersMap(), note: b.reviewNote, req });
    const prev = MRM.getConfig(data);
    setCfg(data, v.value, me);
    db.save(data);
    writeLog('SAVE_REMINDER_RULES', me, 'reminderRules', cfgLogDetail(prev, v.value), req);
    res.json({ success: true, config: { rules: v.value.rules, updatedAt: data.reminderRules.updatedAt, updatedBy: me, isDefault: false } });
  });

  // 試算器：與實際引擎共用 lib/maReminders.js 的 plan／windowState／conditionHolds（只組一份合成資料），唯讀、不寫入
  app.post('/api/admin/reminder-rules/simulate', maOn, requireAdmin, (req, res) => {
    const b = req.body || {};
    const dl = typeof b.daysLeft === 'number' ? b.daysLeft : (typeof b.daysLeft === 'string' && /^-?\d{1,5}$/.test(b.daysLeft.trim()) ? Number(b.daysLeft) : NaN);
    if (!Number.isInteger(dl) || dl < -1000 || dl > 1000) return err(res, 400, 'BAD_DAYS_LEFT', 'daysLeft 必須是 -1000 ～ 1000 的整數（距迄日還有幾天，過期為負）');
    if (!['none', 'draft', 'sent'].includes(b.quoteState)) return err(res, 400, 'BAD_QUOTE_STATE', 'quoteState 只能是 none（沒有續約報價單）、draft（草稿未送簽）、sent（已送簽）');
    if (b.notRenew !== undefined && typeof b.notRenew !== 'boolean') return err(res, 400, 'BAD_NOT_RENEW', 'notRenew 必須是 true／false');
    if (b.termKind !== undefined && !MRM.KINDS.includes(b.termKind)) return err(res, 400, 'BAD_TERM_KIND', 'termKind 只能是 license 或 ma');
    const data = db.load();
    let rules;
    if (b.rules !== undefined) {
      const v = MRM.validateConfig({ rules: b.rules });
      if (!v.ok) return err(res, v.status, v.code, v.error);
      rules = v.value.rules;
    } else rules = MRM.getConfig(data).rules;
    let item = null;
    if (b.maItemId !== undefined && b.maItemId !== null && b.maItemId !== '') {
      if (typeof b.maItemId !== 'string') return err(res, 400, 'BAD_ITEM_ID', 'maItemId 必須是字串');
      item = (data.maItems || []).find((x) => x && x.id === b.maItemId) || null;
      if (!item) return err(res, 404, 'ITEM_NOT_FOUND', '找不到這個納管項');
    }
    const users = (loadAuth().users) || [];
    const out = MRM.simulate(rules, { daysLeft: dl, quoteState: b.quoteState, notRenew: b.notRenew === true, termKind: b.termKind }, { now: Date.now(), item, users, labelOf: (u) => labelOf(users, u) });
    if (item) out.item = { id: item.id, company: item.company, product: item.product, termKind: item.termKind, ownerLabel: labelOf(users, item.owner) };
    out.source = b.rules !== undefined ? 'draft' : 'saved';
    res.set('Cache-Control', 'no-store');
    res.json(out);
  });

  // 單一納管項的規則時間軸（權限同 ma-items 的可視範圍；秘書不給收件人名單、不含金額與合約編號）
  app.get('/api/ma-items/:id/reminders', maOn, requireAuth, (req, res) => {
    const data = db.load();
    const it = visibleItems(req, data).find((x) => x && x.id === req.params.id);
    if (!it) return err(res, 404, 'ITEM_NOT_FOUND', '找不到這個納管項');
    const users = (loadAuth().users) || [];
    const secretary = req.session.user.role === 'secretary';
    const cfg = MRM.getConfig(data);
    const idx = MRM.renewalIndex(data);
    const rows = MRM.timeline(data, it, cfg.rules, { now: Date.now(), secretary, labelOf: (u) => labelOf(users, u) });
    const out = {
      enabled: true,
      item: { company: it.company, product: it.product, termKind: it.termKind, from: it.from, to: it.to, daysLeft: MA.daysLeft(it.to, Date.now()), status: it.status, ownerLabel: labelOf(users, it.owner) },
      quoteState: MRM.quoteStateOf(idx, it.id), rules: rows, rulesUpdatedAt: cfg.updatedAt, isDefault: cfg.isDefault,
    };
    if (!secretary) { out.item.id = it.id; out.item.notRenewReason = it.status === 'not_renew' ? (it.notRenewReason || '') : ''; }
    res.set('Cache-Control', 'no-store');
    res.json(out);
  });

  // 新增報價時「這是續約嗎？」：該客戶 90 天內到期（含已逾期未處理）的追蹤中納管項（範圍同 ma-items；秘書回空）
  app.get('/api/ma-items/renewal-candidates', maOn, requireAuth, (req, res) => {
    const company = typeof req.query.company === 'string' ? req.query.company.trim() : '';
    if (!company) return err(res, 400, 'COMPANY_REQUIRED', '需要 company 參數（客戶名稱）');
    res.set('Cache-Control', 'no-store');
    if (req.session.user.role === 'secretary') return res.json({ enabled: true, company, total: 0, items: [] });
    const data = db.load();
    const users = (loadAuth().users) || [];
    const idx = MRM.renewalIndex(data);
    const now = Date.now();
    const items = visibleItems(req, data).filter((it) => it && it.status === 'tracking' && OL.companyEq(it.company, company))
      .map((it) => ({ it, dl: MA.daysLeft(it.to, now) })).filter((x) => x.dl !== null && x.dl <= 90 && x.dl >= -MRM.MAX_OVERDUE_DAYS)
      .sort((a, b) => a.dl - b.dl || (a.it.id < b.it.id ? -1 : 1))
      .map(({ it, dl }) => ({ id: it.id, product: it.product, termKind: it.termKind, from: it.from, to: it.to, daysLeft: dl, ownerLabel: labelOf(users, it.owner), hasRenewalQuote: MRM.quoteStateOf(idx, it.id) !== 'none', quoteState: MRM.quoteStateOf(idx, it.id) }));
    res.json({ enabled: true, company, total: items.length, items });
  });

  // 標記不續約（負責業務或管理員；原因必填 ≥2 字）：納管項 → not_renew，停止所有提醒；通知一級主管與行銷
  app.patch('/api/ma-items/:id/not-renew', maOn, requireAuth, (req, res) => {
    const me = req.session.user.username;
    const data = db.load();
    const it = (data.maItems || []).find((x) => x && x.id === req.params.id);
    if (!it) return err(res, 404, 'ITEM_NOT_FOUND', '找不到這個納管項');
    if (!isOwnerOrAdmin(req, it.owner)) return err(res, 403, 'FORBIDDEN', '只有負責業務或管理員可以標記不續約');
    const reason = typeof (req.body || {}).reason === 'string' ? req.body.reason.trim() : '';
    if (Array.from(reason).length < MA.SKIP_REASON_MIN) return err(res, 400, 'NOT_RENEW_REASON_REQUIRED', '請填寫不續約的原因（至少 ' + MA.SKIP_REASON_MIN + ' 個字）');
    if (Array.from(reason).length > 200) return err(res, 400, 'NOT_RENEW_REASON_TOO_LONG', '原因最長 200 個字');
    if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f-\u009f]/.test(reason)) return err(res, 400, 'NOT_RENEW_REASON_INVALID', '原因含有不允許的控制字元');
    if (it.status !== 'tracking') return err(res, 409, 'ITEM_NOT_TRACKING', '只有「追蹤中」的納管項可以標記不續約（目前狀態：' + it.status + '）');
    const t = nowIso();
    it.status = 'not_renew'; it.notRenewReason = reason; it.updatedAt = t;
    if (!Array.isArray(it.history)) it.history = [];
    it.history.push({ at: t, by: me, action: 'NOT_RENEW', note: reason });
    if (it.history.length > 100) it.history.splice(0, it.history.length - 100);
    const users = (loadAuth().users) || [];
    MRM.reconcileNotifications(data, users, Date.now(), MRM.getConfig(data).rules, (n) => Array.isArray(n.maItemIds) && n.maItemIds.includes(it.id));
    db.save(data);
    writeLog('MA_ITEM_NOT_RENEW', me, it.company || '', `標記不續約｜${it.product}｜${it.from}～${it.to}｜原因：${reason}`, req);
    const m1 = OR.findManager1(users, it.owner);
    const to = Array.from(new Set([].concat(m1 ? [m1] : [], OR.marketingUsers(users)))).filter((u) => u !== me);
    const co = it.company || '（未命名公司）';
    to.forEach((u) => { try { pushNotification(u, 'ma_not_renew', '⚠️ 客戶不續約：' + co, labelOf(users, me) + ' 標記「' + co + '」' + (it.product || '') + ' 不續約。原因：' + reason, it.id); } catch (e) { console.warn('[ma not-renew notify]', e && e.message); } });
    const fresh = db.load();
    const row = MA.listItems([(fresh.maItems || []).find((x) => x.id === it.id) || it], users, new Date(), {})[0];
    res.json({ item: row, notRenewReason: reason, notified: to.length });
  });

  // 撤銷不續約（僅管理員）：回到追蹤中；已發過的提醒不會重發（去重鍵還在）
  app.patch('/api/ma-items/:id/reopen', maOn, requireAuth, (req, res) => {
    const me = req.session.user.username;
    if (req.session.user.role !== 'admin') return err(res, 403, 'FORBIDDEN', '只有管理員可以撤銷不續約');
    const data = db.load();
    const it = (data.maItems || []).find((x) => x && x.id === req.params.id);
    if (!it) return err(res, 404, 'ITEM_NOT_FOUND', '找不到這個納管項');
    if (it.status !== 'not_renew') return err(res, 409, 'ITEM_NOT_NOT_RENEW', '只有「不續約」狀態的納管項可以撤銷（目前狀態：' + it.status + '）');
    const t = nowIso();
    const oldReason = it.notRenewReason || '';
    it.status = 'tracking'; delete it.notRenewReason; it.updatedAt = t;
    if (!Array.isArray(it.history)) it.history = [];
    it.history.push({ at: t, by: me, action: 'NOT_RENEW_REVOKED', note: oldReason });
    if (it.history.length > 100) it.history.splice(0, it.history.length - 100);
    db.save(data);
    writeLog('MA_ITEM_NOT_RENEW_REVOKED', me, it.company || '', `撤銷不續約，恢復追蹤｜${it.product}｜${it.from}～${it.to}｜原不續約原因：${oldReason}`, req);
    const users = (loadAuth().users) || [];
    res.json({ item: MA.listItems([it], users, new Date(), {})[0] });
  });
}

module.exports = { register, poll, lazyMiddleware, scrub, sweepReminders, sweepNow, checkRenewOf, lazyDue };
