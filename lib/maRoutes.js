'use strict';
/**
 * 成交納管／任務卡／到期管理 的 HTTP 端點與商機成交掛鉤的通知＋稽核（P2a）。純邏輯在 lib/maTracking.js。
 * MA_TRACKING=1 才啟用；關閉時所有端點 next('route')（與未登記相同＝404），掛鉤函式不被呼叫。
 *
 * server.js：  const MAR = require('./lib/maRoutes');  …  MAR.register(app, deps)   （在所有路由之後）
 *              商機 PUT／DELETE／restore 內呼叫 MAR.afterWon／afterUnwon／afterRestore（已 db.save 之後）
 */
const MA = require('./maTracking');
const OR = require('./orgResolve');

let D = null;   // 注入的相依（register 時設定）

const maOn = (req, res, next) => (MA.enabled() ? next() : next('route'));
const err = (res, status, code, message, extra) => res.status(status).json(Object.assign({ error: message, code }, extra || {}));
const nowIso = () => new Date().toISOString();
const usersMap = () => OR.toMap(D.loadAuth().users || []);

/** 發站內通知（失敗不影響主流程）。pushNotification 自帶 load/save，必須在 db.save 之後呼叫 */
function notify(to, type, title, body, refId) {
  try { D.pushNotification(to, type, title, body, refId); } catch (e) { console.warn('[ma notify]', e && e.message); }
}

function taskNotify(task, why) {
  const co = task.company || '（未命名公司）';
  if (why === 'reassigned') notify(task.owner, 'deal_task', '📝 你接手了一張簽約確認任務：' + co, '「' + co + '」' + (task.product || '') + ' 已成交，請確認合約是否已簽訂，並在任務卡填寫合約編號與簽約日。', task.id);
  else notify(task.owner, 'deal_task', '📝 請確認簽約：' + co, '「' + co + '」' + (task.product || '') + ' 已成交。請確認合約是否已簽訂，並在任務卡填寫合約編號與簽約日（或回報此案不簽合約）。', task.id);
}

/** 商機轉成 Won 之後（資料已存檔）：稽核＋通知。r＝MA.applyWon 的回傳 */
function afterWon(req, opp, r, by) {
  const who = by || (req && req.session && req.session.user && req.session.user.username) || '';
  const co = opp.company || opp.contactName || '';
  r.createdItems.forEach((it) => D.writeLog('MA_ITEM_CREATE', who, co, `成交納管｜${it.product}｜${it.termKind}｜${it.from}～${it.to}｜來源 ${it.source}`, req));
  r.revivedItems.forEach((it) => D.writeLog('MA_ITEM_REOPEN', who, co, `商機再次成交，重新啟用｜${it.termKind}｜${it.from}～${it.to}`, req));
  if (r.task && r.taskAction) {
    D.writeLog('MA_TASK_' + (r.taskAction === 'created' ? 'CREATE' : 'REOPEN'), who, co, `簽約確認任務卡｜負責 ${r.task.owner}｜${r.task.product}`, req);
    taskNotify(r.task, '');
  }
}

/** 商機離開 Won／被刪除之後（資料已存檔）：稽核 */
function afterUnwon(req, opp, u, reason, by) {
  const who = by || (req && req.session && req.session.user && req.session.user.username) || '';
  const co = (opp && (opp.company || opp.contactName)) || '';
  u.cancelledTasks.forEach((t) => D.writeLog('MA_TASK_CANCEL', who, co, `任務卡取消（${reason === 'OPP_DELETED' ? '商機刪除' : '商機離開成交階段'}）｜${t.id}`, req));
  u.voidedItems.forEach((it) => D.writeLog('MA_ITEM_VOID', who, co, `納管項作廢（${reason === 'OPP_DELETED' ? '商機刪除' : '商機離開成交階段'}）｜${it.termKind}｜${it.from}～${it.to}`, req));
}

function register(app, deps) {
  D = deps;
  const { requireAuth, getViewableOwners, filterByBu, filterByViewGroup, db, loadAuth, writeLog, userLabel } = deps;
  const ADMINLIKE_OVERDUE = ['admin', 'executive', 'manager1', 'manager2'];

  /** 登入者可視的納管項（owner 範圍＋BU＋集團範圍，與 GET /api/opportunities 同一套） */
  function visibleItems(req, data) {
    const owners = new Set(getViewableOwners(req, 'opportunities'));
    let items = (data.maItems || []).filter((it) => it && owners.has(it.owner));
    items = filterByBu(req, items);
    items = filterByViewGroup(req, items, (it) => it.company);
    return items;
  }
  function visibleOpps(req, data) {
    const owners = new Set(getViewableOwners(req, 'opportunities'));
    let opps = (data.opportunities || []).filter((o) => o && owners.has(o.owner));
    opps = filterByBu(req, opps);
    opps = filterByViewGroup(req, opps, (o) => o.company);
    return opps;
  }
  const isOwnerOrAdmin = (req, owner) => req.session.user.role === 'admin' || req.session.user.username === owner;

  // ── 我的開啟中任務卡 ──
  app.get('/api/my-tasks/deals', maOn, requireAuth, (req, res) => {
    const data = db.load();
    const users = loadAuth().users || [];
    const items = MA.myTasks(data, req.session.user.username, users, new Date());
    res.set('Cache-Control', 'no-store');
    res.json({ enabled: true, items, total: items.length, overdueDays: MA.OVERDUE_DAYS, snoozeDays: MA.SNOOZE_DAYS });
  });

  // ── 主管：團隊未簽超過 14 天 ──
  app.get('/api/manager/deal-task-overdue', maOn, requireAuth, (req, res) => {
    const role = req.session.user.role;
    if (!ADMINLIKE_OVERDUE.includes(role)) return err(res, 403, 'FORBIDDEN', '無權限（限主管）');
    const data = db.load();
    const users = loadAuth().users || [];
    const owners = new Set(getViewableOwners(req, 'opportunities'));
    const items = MA.overdueTasks(data, owners, users, new Date());
    res.set('Cache-Control', 'no-store');
    res.json({ enabled: true, overdueDays: MA.OVERDUE_DAYS, items, canReassign: role === 'admin' });
  });

  // ── 到期管理清單 ──
  app.get('/api/ma-items', maOn, requireAuth, (req, res) => {
    const data = db.load();
    const users = loadAuth().users || [];
    const secretary = req.session.user.role === 'secretary';
    const filter = MA.FILTERS.includes(String(req.query.filter || '')) ? String(req.query.filter) : 'all';
    const items = MA.listItems(visibleItems(req, data), users, new Date(), { filter, secretary, tasks: secretary ? [] : data.dealTasks });
    res.set('Cache-Control', 'no-store');
    res.json({ enabled: true, filter, total: items.length, items });
  });

  // ── 待補納管（已成交但缺納管：Excel 匯入、開關開啟前就成交的舊單）──
  app.get('/api/ma-items/pending-backfill', maOn, requireAuth, (req, res) => {
    const data = db.load();
    const users = loadAuth().users || [];
    const items = req.session.user.role === 'secretary' ? [] : MA.pendingBackfill(data, visibleOpps(req, data), users);
    res.set('Cache-Control', 'no-store');
    res.json({ enabled: true, total: items.length, items });
  });

  // ── 補登納管（商機已是 Won）：期間來源與成交時相同（body 或已核准報價單）──
  app.post('/api/ma-items/backfill', maOn, requireAuth, (req, res) => {
    const me = req.session.user.username;
    const b = req.body || {};
    const oppId = typeof b.oppId === 'string' ? b.oppId : '';
    const data = db.load();
    const opp = (data.opportunities || []).find((o) => o && o.id === oppId);
    if (!opp) return err(res, 404, 'OPP_NOT_FOUND', '找不到這筆商機');
    if (!isOwnerOrAdmin(req, opp.owner)) return err(res, 403, 'FORBIDDEN', '只有負責業務或管理員可以補登');
    if (opp.stage !== 'Won') return err(res, 409, 'OPP_NOT_WON', '這筆商機還沒有成交');
    if ((data.maItems || []).some((it) => it && it.oppId === opp.id && it.status !== 'void')) return err(res, 409, 'ALREADY_TRACKED', '這筆商機已經在納管中');
    const plan = MA.planWon(data, { id: opp.id, product: opp.product }, b);
    if (!plan.ok) return err(res, plan.status, plan.code, plan.message);
    if (!plan.track) return err(res, 409, 'NOT_TERM_PRODUCT', '這個商品不是 License／MA 類，不需要納管');
    const r = MA.applyWon(data, opp, plan, { by: me, nowIso: nowIso() });
    db.save(data);
    writeLog('MA_ITEM_BACKFILL', me, opp.company || '', `補登納管｜${opp.product}｜${r.items.length} 筆`, req);
    afterWon(req, opp, r, me);
    res.status(201).json({ items: r.items.map((it) => MA.listItems([it], loadAuth().users || [], new Date(), {})[0]), taskId: r.task ? r.task.id : '', taskAction: r.taskAction });
  });

  // ── 修改納管項期間 ──
  app.patch('/api/ma-items/:id/term', maOn, requireAuth, (req, res) => {
    const me = req.session.user.username;
    const data = db.load();
    const it = (data.maItems || []).find((x) => x && x.id === req.params.id);
    if (!it) return err(res, 404, 'ITEM_NOT_FOUND', '找不到這個納管項');
    if (!isOwnerOrAdmin(req, it.owner)) return err(res, 403, 'FORBIDDEN', '只有負責業務或管理員可以修改');
    const r = MA.editTerm(data, req.params.id, req.body || {}, { by: me, nowIso: nowIso() });
    if (!r.ok) return err(res, r.status, r.code, r.message);
    if (r.changed) {
      db.save(data);
      writeLog('MA_ITEM_TERM_EDIT', me, it.company || '', `修改服務期間｜${it.termKind}｜${r.old.from}～${r.old.to} → ${it.from}～${it.to}${r.contract ? '｜同步合約起迄日' : ''}${typeof (req.body || {}).reason === 'string' && req.body.reason.trim() ? '｜原因：' + req.body.reason.trim().slice(0, 200) : ''}`, req);
    }
    res.json(MA.listItems([it], loadAuth().users || [], new Date(), {})[0]);
  });

  // ── 任務卡動作 ──
  function loadTaskFor(req, res, adminOnly) {
    const data = db.load();
    const t = (data.dealTasks || []).find((x) => x && x.id === req.params.id);
    if (!t) { err(res, 404, 'TASK_NOT_FOUND', '找不到這張任務卡'); return null; }
    const ok = adminOnly ? req.session.user.role === 'admin' : isOwnerOrAdmin(req, t.owner);
    if (!ok) { err(res, 403, 'FORBIDDEN', adminOnly ? '限管理員' : '只有負責業務或管理員可以處理這張任務卡'); return null; }
    return { data, t };
  }
  const viewOf = (data, t) => MA.taskView(data, t, loadAuth().users || [], new Date());

  app.post('/api/deal-tasks/:id/sign', maOn, requireAuth, (req, res) => {
    const L = loadTaskFor(req, res, false); if (!L) return;
    const me = req.session.user.username;
    const users = loadAuth().users || [];
    const ou = OR.toMap(users)[L.t.owner];
    const r = MA.signTask(L.data, req.params.id, req.body || {}, { by: me, nowIso: nowIso(), salesPerson: userLabel(ou, L.t.owner) });
    if (!r.ok) return err(res, r.status, r.code, r.message);
    db.save(L.data);
    writeLog('MA_TASK_SIGN', me, r.contract.company || '', `簽約確認｜合約 ${r.contract.contractNo}｜簽約日 ${r.contract.signedDate}｜${r.contract.startDate}～${r.contract.endDate}｜任務 ${r.task.id}`, req);
    writeLog('CREATE_CONTRACT', me, r.contract.company || '', `合約No:${r.contract.contractNo} 產品:${r.contract.product}（由成交任務卡建立 SERVICE_MA）`, req);
    res.json({ task: viewOf(L.data, r.task), contract: { id: r.contract.id, contractNo: r.contract.contractNo, type: r.contract.type, startDate: r.contract.startDate, endDate: r.contract.endDate, signedDate: r.contract.signedDate } });
  });

  app.post('/api/deal-tasks/:id/skip', maOn, requireAuth, (req, res) => {
    const L = loadTaskFor(req, res, false); if (!L) return;
    const me = req.session.user.username;
    const users = loadAuth().users || [];
    const r = MA.skipTask(L.data, req.params.id, req.body || {}, { by: me, nowIso: nowIso() }, users);
    if (!r.ok) return err(res, r.status, r.code, r.message);
    db.save(L.data);
    writeLog('MA_TASK_SKIP', me, L.t.company || '', `此案不簽合約｜原因：${r.reason}｜任務 ${L.t.id}｜通知 ${r.notify.length} 人`, req);
    const co = L.t.company || '（未命名公司）';
    const who = userLabel(OR.toMap(users)[L.t.owner], L.t.owner);
    r.notify.forEach((u) => notify(u, 'deal_task_skipped', '⚠️ 成交案不簽合約：' + co, who + ' 回報「' + co + '」' + (L.t.product || '') + ' 此案不簽合約。原因：' + r.reason, L.t.id));
    res.json({ task: viewOf(L.data, r.task), notified: r.notify.length });
  });

  app.post('/api/deal-tasks/:id/snooze', maOn, requireAuth, (req, res) => {
    const L = loadTaskFor(req, res, false); if (!L) return;
    const me = req.session.user.username;
    const r = MA.snoozeTask(L.data, req.params.id, { by: me, nowIso: nowIso() });
    if (!r.ok) return err(res, r.status, r.code, r.message);
    db.save(L.data);
    writeLog('MA_TASK_SNOOZE', me, L.t.company || '', `稍後提醒至 ${r.snoozedUntil}｜任務 ${L.t.id}`, req);
    res.json({ task: viewOf(L.data, r.task), snoozedUntil: r.snoozedUntil });
  });

  app.post('/api/deal-tasks/:id/reassign', maOn, requireAuth, (req, res) => {
    const L = loadTaskFor(req, res, true); if (!L) return;
    const me = req.session.user.username;
    const to = req.body && typeof req.body.to === 'string' ? req.body.to : '';
    const r = MA.reassignTask(L.data, req.params.id, to, { by: me, nowIso: nowIso() }, loadAuth().users || []);
    if (!r.ok) return err(res, r.status, r.code, r.message);
    db.save(L.data);
    writeLog('MA_TASK_REASSIGN', me, L.t.company || '', `改派任務卡｜${r.from} → ${r.to}｜任務 ${L.t.id}`, req);
    taskNotify(L.t, 'reassigned');
    res.json({ task: viewOf(L.data, r.task) });
  });
}

module.exports = { register, afterWon, afterUnwon, afterRestore: afterWon };
