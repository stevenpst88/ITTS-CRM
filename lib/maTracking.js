'use strict';
/**
 * 成交納管（Service MA 管理）＋任務卡＋合約串接 —— 純邏輯（P2a）。環境變數 MA_TRACKING=1 才啟用（lib/quoteTerms.js 的 enabled()，與 P1 同一個開關）。
 *
 * 資料（主資料 blob，只在功能開啟且真的有成交時才建立陣列，不在讀取時憑空新增）：
 *   data.maItems[]   納管項：{ id, oppId, quoteId, owner, bu?, company, product, termKind('license'|'ma'), from, to, contractNo, contractId,
 *                              status('tracking'|'ended'|'renewed'|'not_renew'|'void'), prevStatus?, notRenewReason?, source('request'|'quote'|'history'|'backfill'),
 *                              createdAt, updatedAt, history[] }
 *   data.dealTasks[] 任務卡：{ id, type:'ma_sign', oppId, maItemId, maItemIds[], owner, company, product, status('open'|'done'|'skipped'|'cancelled'),
 *                              steps:[{ key:'sign', status('pending'|'done'|'skipped'), value, doneBy, doneAt, reason }], snoozedUntil?, createdAt, updatedAt, history[] }
 *   data.contracts[] 新增選填欄位 oppId／quoteId／maItemId／signedDate，type 可為 'SERVICE_MA'（舊合約沒有這些鍵，輸出不變）
 *
 * 規則重點：
 *   · 納管項以 (oppId, termKind, from, to) 冪等；任務卡每個商機一張（type ma_sign）。
 *   · 期間來源優先序：請求帶入（termFrom／termTo）> 該商機連結的「最近一張已核准」且有期間品項的報價單 > 該商機先前作廢（void）的納管項（Won 回退後再成交）。
 *   · 日期一律是 YYYY-MM-DD 字串，倒數用台北日曆日相減（與伺服器時區無關）。
 *   · 本檔不做權限判斷（owner／管理員／可視範圍由 server.js 路由處理），不發通知、不寫稽核（呼叫端依回傳結果做）。
 */
const crypto = require('crypto');
const QT = require('./quoteTerms');
const QI = require('./quoteItems');
const QOL = require('./quoteOppLink');
const PC = require('./productCatalog');
const OR = require('./orgResolve');

const MAX_ITEMS = 20000;          // data.maItems 上限
const MAX_TASKS = 20000;          // data.dealTasks 上限
const MAX_HISTORY = 100;          // 單一納管項／任務卡的 history 上限（保留最新）
const OVERDUE_DAYS = 14;          // 任務卡建立超過幾天未處理 → 一級主管逾期清單
const SNOOZE_DAYS = 3;
const CONTRACT_NO_MAX = 64;
const SKIP_REASON_MIN = 2;
const SKIP_REASON_MAX = 200;
const REASON_MAX = 200;
const TERMINAL_ITEM = ['void', 'ended', 'renewed', 'not_renew'];
const TASK_TYPE = 'ma_sign';
const OWNER_ROLES = ['user'];   // 可被改派為任務卡負責人的角色：只有業務（驗收後修正：不再含主管與集團業務）

const enabled = (env) => QT.enabled(env);
const txt = (v) => String(v === undefined || v === null ? '' : v).trim();
const newId = () => crypto.randomUUID();

// ───────────────── 日期（台北日曆日）─────────────────
function taipeiDate(now) {
  const t = now instanceof Date ? now.getTime() : (typeof now === 'number' ? now : Date.parse(now));
  return new Date((Number.isFinite(t) ? t : Date.now()) + 8 * 3600 * 1000).toISOString().slice(0, 10);
}
const dayNum = (s) => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s)); return m ? Math.round(Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000) : NaN; };
/** to 日期距今天還有幾天（今天迄日＝0、昨天迄日＝-1）；日期不合法 → null */
function daysLeft(to, now) {
  if (!QT.isRealDate(to)) return null;
  return dayNum(to) - dayNum(taipeiDate(now));
}
/** ISO 時間（UTC）換成台北日曆日後，距今天幾天 */
function daysSinceIso(iso, now) {
  const t = iso ? Date.parse(iso) : NaN;
  if (!Number.isFinite(t)) return 0;
  return Math.max(0, dayNum(taipeiDate(now)) - dayNum(taipeiDate(t)));
}
function addDaysIso(dateStr, n) { return new Date((dayNum(dateStr) + n) * 86400000).toISOString().slice(0, 10); }

const itemKey = (oppId, kind, from, to) => [oppId, kind, from, to].join('|');
function pushHist(arr, entry) { arr.push(entry); if (arr.length > MAX_HISTORY) arr.splice(0, arr.length - MAX_HISTORY); }
const H = (nowIso, by, action, note, extra) => Object.assign({ at: nowIso, by: by || '', action, note: note || '' }, extra || {});

// ───────────────── 成交：期間來源與計畫 ─────────────────
/** 報價單最終核准時間（無法判斷時退回 updatedAt／createdAt）—— 用來挑「最近已核准」 */
function approvedAtOf(q) {
  const ap = q && q.approval;
  let best = '';
  ((ap && ap.steps) || []).forEach((s) => { if (s && s.status === 'approved' && typeof s.at === 'string' && s.at > best) best = s.at; });
  return best || txt(q && q.updatedAt) || txt(q && q.createdAt);
}

/**
 * 該商機連結的最近已核准報價單中的期間品項：回 { quoteId, entries:[{termKind, from, to}] }（依 distinct 期間分組）；沒有 → null。
 * 「已核准」＝ QT.termEditState（核准成立且未失效）；期間不合法的品項略過；最近的單沒有期間品項就往更舊的找。
 */
function termsFromQuotes(data, oppId) {
  const qs = QOL.linkedQuotes(data, oppId).filter((q) => QT.termEditState(q))
    .sort((a, b) => (approvedAtOf(b) < approvedAtOf(a) ? -1 : approvedAtOf(b) > approvedAtOf(a) ? 1 : 0));
  for (const q of qs) {
    const seen = new Set();
    const entries = [];
    (q.items || []).forEach((it) => {
      if (!QI.isItemRow(it)) return;
      const k = QT.parseKind(it.termKind);
      if (!k) return;
      const from = txt(it.termFrom), to = txt(it.termTo);
      if (QT.validateTerm(k, from, to)) return;
      const key = k + '|' + from + '|' + to;
      if (seen.has(key)) return;
      seen.add(key);
      entries.push({ termKind: k, from, to });
    });
    if (entries.length) return { quoteId: q.id, entries };
  }
  return null;
}

/** 該商機先前作廢（Won 回退）的納管項期間，供再次成交時沿用 */
function voidedEntries(data, oppId) {
  const out = [];
  ((data && data.maItems) || []).forEach((it) => { if (it && it.oppId === oppId && it.status === 'void') out.push({ termKind: it.termKind, from: it.from, to: it.to, quoteId: it.quoteId || '' }); });
  return out;
}

/** 目錄可清掉的納管項（終態且超過一段時間不用）數量：容量預檢用，不修改資料 */
function prunableCount(data) { return ((data && data.maItems) || []).filter((it) => it && TERMINAL_ITEM.includes(it.status)).length; }

/**
 * 成交計畫（純檢查，不修改 data）。oppLike＝{ id, product }（用「這次寫入後」的商品文字）；body＝請求 body（只看 termFrom／termTo）。
 * 回：
 *   { ok:true, track:false }                                  商品不是 License／MA 類 → 不用納管
 *   { ok:true, track:true, rule, entries:[{termKind,from,to,source,quoteId?}], newCount }
 *   { ok:false, status, code, message }                       TERM_REQUIRED／TERM_INVALID／MA_CAPACITY
 */
function planWon(data, oppLike, body) {
  const rule = PC.resolveTermRule(data && data.productCatalog, oppLike && oppLike.product);
  if (rule.termKind === 'none') return { ok: true, track: false, rule };
  const b = body && typeof body === 'object' ? body : {};
  const from = txt(b.termFrom), to = txt(b.termTo);
  let entries = null;
  if (from || to) {
    const bad = QT.validateTerm(rule.termKind, from, to);
    if (bad) return { ok: false, status: 400, code: bad.code, message: '成交需要有效的服務期間：' + bad.reason };
    entries = [{ termKind: rule.termKind, from, to, source: 'request' }];
  } else {
    const fq = termsFromQuotes(data, oppLike.id);
    if (fq) entries = fq.entries.map((e) => Object.assign({}, e, { source: 'quote', quoteId: fq.quoteId }));
    else {
      const ve = voidedEntries(data, oppLike.id);
      if (ve.length) entries = ve.map((e) => Object.assign({}, e, { source: 'history' }));
    }
  }
  if (!entries || !entries.length) {
    return { ok: false, status: 400, code: 'TERM_REQUIRED', message: '這個商品是 ' + QT.KIND_LABEL[rule.termKind] + ' 類，成交時需要填寫服務期間（起日與迄日）；也可以先讓連結的報價單核准並填妥品項的服務期間' };
  }
  const have = new Set(((data && data.maItems) || []).filter((it) => it && it.oppId === oppLike.id).map((it) => itemKey(it.oppId, it.termKind, it.from, it.to)));
  const newCount = entries.filter((e) => !have.has(itemKey(oppLike.id, e.termKind, e.from, e.to))).length;
  const cur = ((data && data.maItems) || []).length;
  if (newCount && cur + newCount > MAX_ITEMS + prunableCount(data)) return { ok: false, status: 409, code: 'MA_CAPACITY', message: '納管項數量已達上限，請聯絡管理員清理' };
  return { ok: true, track: true, rule, entries, newCount };
}

// ───────────────── 通知已讀 ─────────────────
/** 任務卡完成／略過／取消／延後時，把「請簽約確認」那則通知標為已讀（只動 type==='deal_task' 且 refId 是該任務卡的）。回標記數 */
function markTaskNotificationsRead(data, taskId) {
  let n = 0;
  ((data && data.notifications) || []).forEach((x) => { if (x && x.type === 'deal_task' && x.refId === taskId && !x.read) { x.read = true; n++; } });
  return n;
}

function pruneItems(data) {
  if (!Array.isArray(data.maItems) || data.maItems.length < MAX_ITEMS) return;
  const open = new Set((data.dealTasks || []).filter((t) => t && t.status === 'open').flatMap((t) => t.maItemIds || [t.maItemId]));
  const keep = [];
  const drop = data.maItems.length - MAX_ITEMS + 1;
  let dropped = 0;
  // 依建立時間由舊到新丟終態項（有開啟中任務卡引用的不丟）
  const terminal = data.maItems.filter((it) => it && TERMINAL_ITEM.includes(it.status) && !open.has(it.id)).sort((a, b) => (a.createdAt < b.createdAt ? -1 : 1));
  const kill = new Set(terminal.slice(0, drop).map((it) => it.id));
  data.maItems.forEach((it) => { if (kill.has(it.id)) dropped++; else keep.push(it); });
  data.maItems = keep;
  return dropped;
}
function pruneTasks(data) {
  if (!Array.isArray(data.dealTasks) || data.dealTasks.length < MAX_TASKS) return;
  const drop = data.dealTasks.length - MAX_TASKS + 1;
  const done = data.dealTasks.filter((t) => t && t.status !== 'open').sort((a, b) => (a.createdAt < b.createdAt ? -1 : 1));
  const kill = new Set(done.slice(0, drop).map((t) => t.id));
  data.dealTasks = data.dealTasks.filter((t) => !kill.has(t.id));
}

// ───────────────── 成交套用 ─────────────────
/**
 * 續約成交（P3）：該商機連結的「已核准且未失效」報價單若帶 renewOf（指向舊納管項），把舊納管項標為 renewed（含 history）。
 * 只動「追蹤中」的舊項（已 renewed／not_renew／ended／void 不動，所以冪等）；舊項若屬於同一個商機（自己續自己）略過。
 * 在 applyWon 末尾呼叫（成交且商品為 License／MA 類時才會走到）。回 [被標為 renewed 的舊項]。
 */
function applyRenewals(data, opp, ctx) {
  const out = [];
  QOL.linkedQuotes(data, opp.id).filter((q) => QT.termEditState(q) && typeof q.renewOf === 'string' && q.renewOf).forEach((q) => {
    const old = (data.maItems || []).find((x) => x && x.id === q.renewOf);
    if (!old || old.status !== 'tracking' || old.oppId === opp.id) return;
    if (q.deleted || (typeof q.company === 'string' && !QOL.companyEq(old.company, q.company))) return;   // 審查修正 L3：renewOf 只在設定當下驗過同公司；成交時再驗一次（報價單後來被換到別家公司的商機時不誤標舊項為已續約）
    old.status = 'renewed'; old.renewedBy = { oppId: opp.id, quoteId: q.id }; old.updatedAt = ctx.nowIso;
    pushHist(old.history, H(ctx.nowIso, ctx.by, 'RENEWED', '續約成交（新商機成交，續約報價單已核准）', { oppId: opp.id, quoteId: q.id }));
    out.push(old);
  });
  return out;
}

function findItem(data, oppId, kind, from, to) { return (data.maItems || []).find((it) => it && it.oppId === oppId && it.termKind === kind && it.from === from && it.to === to) || null; }
function findTask(data, oppId) { return (data.dealTasks || []).find((t) => t && t.oppId === oppId && t.type === TASK_TYPE) || null; }

/**
 * 套用成交計畫（planWon 回 ok 且 track 才呼叫）。會修改 data（建立／重開納管項、建立／重開任務卡），不存檔、不發通知。
 * opp＝商機（寫入後的物件）；ctx＝{ by, nowIso }。
 * 回 { items:[納管項], createdItems:[], revivedItems:[], task:任務卡|null, taskAction:'created'|'reopened'|'' }
 */
function applyWon(data, opp, plan, ctx) {
  const nowIso = ctx.nowIso, by = ctx.by;
  if (!Array.isArray(data.maItems)) data.maItems = [];
  if (!Array.isArray(data.dealTasks)) data.dealTasks = [];
  if (plan.newCount) pruneItems(data);
  const out = { items: [], createdItems: [], revivedItems: [], task: null, taskAction: '' };
  plan.entries.forEach((e) => {
    let it = findItem(data, opp.id, e.termKind, e.from, e.to);
    if (it) {
      if (it.status === 'void') {
        it.status = it.prevStatus && it.prevStatus !== 'void' ? it.prevStatus : 'tracking';
        delete it.prevStatus;
        it.owner = opp.owner; it.company = txt(opp.company); it.product = txt(opp.product);
        it.updatedAt = nowIso;
        pushHist(it.history, H(nowIso, by, 'REOPEN', '商機再次成交，重新啟用'));
        out.revivedItems.push(it);
      }
    } else {
      it = {
        id: newId(), oppId: opp.id, quoteId: e.quoteId || '', owner: opp.owner,
        company: txt(opp.company), product: txt(opp.product),
        termKind: e.termKind, from: e.from, to: e.to, contractNo: '', contractId: '',
        status: 'tracking', source: e.source || 'request', createdAt: nowIso, updatedAt: nowIso,
        history: [H(nowIso, by, 'CREATE', '商機成交，納入管理', { source: e.source || 'request', from: e.from, to: e.to })],
      };
      if (opp.bu !== undefined && opp.bu !== null && opp.bu !== '') it.bu = opp.bu;   // 沒有就不多一個鍵（可視範圍的 BU 推斷改由 owner 推）
      data.maItems.push(it);
      out.createdItems.push(it);
    }
    out.items.push(it);
  });
  if (plan.rule && plan.rule.startsTask) {
    const liveIds = data.maItems.filter((x) => x && x.oppId === opp.id && x.status !== 'void').map((x) => x.id);
    let t = findTask(data, opp.id);
    if (!t) {
      pruneTasks(data);
      t = {
        id: newId(), type: TASK_TYPE, oppId: opp.id, maItemId: liveIds[0] || '', maItemIds: liveIds, owner: opp.owner,
        company: txt(opp.company), product: txt(opp.product), status: 'open',
        steps: [{ key: 'sign', status: 'pending', value: null, doneBy: '', doneAt: '', reason: '' }],
        createdAt: nowIso, updatedAt: nowIso, history: [H(nowIso, by, 'CREATE', '商機成交，建立簽約確認任務')],
      };
      data.dealTasks.push(t);
      out.taskAction = 'created';
    } else if (t.status === 'cancelled') {
      t.status = 'open'; t.owner = opp.owner; t.maItemIds = liveIds; t.maItemId = liveIds[0] || '';
      t.company = txt(opp.company); t.product = txt(opp.product);
      if (t.steps && t.steps[0] && t.steps[0].status !== 'done') t.steps[0] = { key: 'sign', status: 'pending', value: null, doneBy: '', doneAt: '', reason: '' };
      t.createdAt = nowIso;   // 重開後逾期天數重新計算
      delete t.snoozedUntil;
      t.updatedAt = nowIso;
      pushHist(t.history, H(nowIso, by, 'REOPEN', '商機再次成交，任務卡重開'));
      out.taskAction = 'reopened';
    } else if (t.status === 'open') {
      // 已存在開啟中的任務卡：只補上新增的納管項，不重複建
      const before = (t.maItemIds || []).length;
      t.maItemIds = Array.from(new Set([].concat(t.maItemIds || [], liveIds)));
      if (!t.maItemId) t.maItemId = liveIds[0] || '';
      if (t.maItemIds.length !== before) { t.updatedAt = nowIso; pushHist(t.history, H(nowIso, by, 'ITEMS', '補入新的納管項')); }
    }
    out.task = t;
  }
  out.renewedItems = applyRenewals(data, opp, ctx);
  return out;
}

/**
 * 商機離開 Won（或被刪除）：開啟中任務卡 → cancelled、追蹤中的納管項 → void（history 保留；已完成的任務卡與已簽的合約不動）。不存檔、不發通知。
 * reason：'STAGE_REVERTED'｜'OPP_DELETED'。回 { cancelledTasks:[], voidedItems:[] }
 */
function applyUnwon(data, oppId, ctx, reason) {
  const nowIso = ctx.nowIso, by = ctx.by;
  const out = { cancelledTasks: [], voidedItems: [] };
  (data.dealTasks || []).forEach((t) => {
    if (t && t.oppId === oppId && t.type === TASK_TYPE && t.status === 'open') {
      t.status = 'cancelled'; t.cancelReason = reason || 'STAGE_REVERTED'; t.updatedAt = nowIso;
      pushHist(t.history, H(nowIso, by, 'CANCEL', reason === 'OPP_DELETED' ? '商機已刪除，任務卡取消' : '商機離開成交階段，任務卡取消'));
      markTaskNotificationsRead(data, t.id);
      out.cancelledTasks.push(t);
    }
  });
  (data.maItems || []).forEach((it) => {
    if (it && it.oppId === oppId && it.status === 'tracking') {
      it.prevStatus = it.status; it.status = 'void'; it.updatedAt = nowIso;
      pushHist(it.history, H(nowIso, by, 'VOID', reason === 'OPP_DELETED' ? '商機已刪除，停止追蹤' : '商機離開成交階段，停止追蹤'));
      out.voidedItems.push(it);
    }
  });
  // 續約成交被回退（商機離開 Won 或被刪除）：被這筆商機標為 renewed 的舊項恢復追蹤
  out.revertedItems = [];
  (data.maItems || []).forEach((it) => {
    if (it && it.status === 'renewed' && it.renewedBy && it.renewedBy.oppId === oppId) {
      it.status = 'tracking'; delete it.renewedBy; it.updatedAt = nowIso;
      pushHist(it.history, H(nowIso, by, 'RENEW_REVERTED', reason === 'OPP_DELETED' ? '續約商機已刪除，恢復追蹤' : '續約商機離開成交階段，恢復追蹤'));
      out.revertedItems.push(it);
    }
  });
  return out;
}

/** 商機被還原（從流失清單回來）且階段為 Won：重開作廢的納管項與取消的任務卡（不需要重新輸入期間）。回 { items, task } */
function applyRestore(data, opp, ctx) {
  const ve = voidedEntries(data, opp.id);
  if (!ve.length) return { items: [], createdItems: [], revivedItems: [], task: null, taskAction: '', renewedItems: [] };
  const rule = PC.resolveTermRule(data.productCatalog, opp.product);
  const plan = { entries: ve.map((e) => Object.assign({}, e, { source: 'history' })), rule, newCount: 0 };
  return applyWon(data, opp, plan, ctx);
}

// ───────────────── 任務卡動作 ─────────────────
const fail = (status, code, message) => ({ ok: false, status, code, message });

/** 合約編號是否已被「非刪除」合約使用（trim、不分大小寫）。不洩漏是誰用的。 */
function contractNoTaken(data, contractNo) {
  const k = txt(contractNo).toLowerCase();
  return ((data && data.contracts) || []).some((c) => c && !c.deleted && txt(c.contractNo).toLowerCase() === k);
}

function checkContractNo(raw) {
  const s = typeof raw === 'string' ? raw.trim() : '';
  if (!s) return fail(400, 'CONTRACT_NO_REQUIRED', '請填寫合約編號');
  if (Array.from(s).length > CONTRACT_NO_MAX) return fail(400, 'CONTRACT_NO_INVALID', '合約編號最長 ' + CONTRACT_NO_MAX + ' 個字');
  if (/[\u0000-\u001f\u007f-\u009f]/.test(s)) return fail(400, 'CONTRACT_NO_INVALID', '合約編號含有不允許的控制字元');
  return { ok: true, value: s };
}

/**
 * 簽約確認：建立合約、回寫納管項、任務卡完成。ctx＝{ by, nowIso, salesPerson }。成功回 { ok:true, task, contract, items }。
 * 權限（owner／管理員）由呼叫端檢查。失敗不修改 data。
 */
function signTask(data, taskId, input, ctx) {
  const t = (data.dealTasks || []).find((x) => x && x.id === taskId);
  if (!t) return fail(404, 'TASK_NOT_FOUND', '找不到這張任務卡');
  if (t.status !== 'open') return fail(409, 'TASK_NOT_OPEN', '這張任務卡已經' + ({ done: '完成', skipped: '略過', cancelled: '取消' }[t.status] || '處理') + '，不能再簽約確認');
  const cn = checkContractNo(input && input.contractNo);
  if (!cn.ok) return cn;
  const sd = txt(input && input.signedDate);
  if (!sd) return fail(400, 'SIGNED_DATE_REQUIRED', '請填寫簽約日');
  if (!QT.isRealDate(sd)) return fail(400, 'SIGNED_DATE_INVALID', '簽約日不是有效日期（格式 YYYY-MM-DD）');
  if (contractNoTaken(data, cn.value)) return fail(409, 'DUPLICATE_CONTRACT_NO', '這個合約編號已經被使用，請確認編號是否正確');
  const ids = t.maItemIds && t.maItemIds.length ? t.maItemIds : [t.maItemId];
  const items = (data.maItems || []).filter((it) => it && ids.includes(it.id) && it.status !== 'void');
  if (!items.length) return fail(409, 'NO_MA_ITEM', '這張任務卡沒有可用的納管項（商機可能已離開成交階段）');
  const opp = (data.opportunities || []).find((o) => o && o.id === t.oppId) || (data.lostOpportunities || []).find((o) => o && o.id === t.oppId) || {};
  const primary = items[0];
  const startDate = items.map((i) => i.from).sort()[0];
  const endDate = items.map((i) => i.to).sort().slice(-1)[0];
  if (!Array.isArray(data.contracts)) data.contracts = [];
  const contract = {
    id: newId(), owner: t.owner, contractNo: cn.value, company: txt(opp.company || t.company), contactName: txt(opp.contactName),
    product: txt(opp.product || t.product), startDate, endDate, amount: opp.amount === undefined || opp.amount === null ? '' : opp.amount,
    salesPerson: ctx.salesPerson || '', note: '由成交任務卡建立', type: 'SERVICE_MA',
    oppId: t.oppId, quoteId: primary.quoteId || '', maItemId: primary.id, signedDate: sd, createdAt: ctx.nowIso,
  };
  data.contracts.push(contract);
  items.forEach((it) => {
    it.contractNo = cn.value; it.contractId = contract.id; it.updatedAt = ctx.nowIso;
    pushHist(it.history, H(ctx.nowIso, ctx.by, 'SIGNED', '簽約確認', { contractNo: cn.value, signedDate: sd }));
  });
  t.status = 'done'; t.updatedAt = ctx.nowIso;
  t.steps = [{ key: 'sign', status: 'done', value: { contractNo: cn.value, signedDate: sd }, doneBy: ctx.by, doneAt: ctx.nowIso, reason: '' }];
  delete t.snoozedUntil;
  pushHist(t.history, H(ctx.nowIso, ctx.by, 'SIGN', '簽約確認完成', { contractNo: cn.value, signedDate: sd, contractId: contract.id }));
  markTaskNotificationsRead(data, t.id);
  return { ok: true, task: t, contract, items };
}

/** 此案不簽合約：原因必填。成功回 { ok, task, notify:[帳號...] }（行銷＋一級主管，排除操作者；呼叫端發通知）。納管項保留追蹤。 */
function skipTask(data, taskId, input, ctx, users) {
  const t = (data.dealTasks || []).find((x) => x && x.id === taskId);
  if (!t) return fail(404, 'TASK_NOT_FOUND', '找不到這張任務卡');
  if (t.status !== 'open') return fail(409, 'TASK_NOT_OPEN', '這張任務卡已經不是待處理狀態');
  const reason = typeof (input && input.reason) === 'string' ? input.reason.trim() : '';
  if (Array.from(reason).length < SKIP_REASON_MIN) return fail(400, 'SKIP_REASON_REQUIRED', '請填寫不簽合約的原因（至少 ' + SKIP_REASON_MIN + ' 個字）');
  if (Array.from(reason).length > SKIP_REASON_MAX) return fail(400, 'SKIP_REASON_TOO_LONG', '原因最長 ' + SKIP_REASON_MAX + ' 個字');
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f-\u009f]/.test(reason)) return fail(400, 'SKIP_REASON_INVALID', '原因含有不允許的控制字元');
  t.status = 'skipped'; t.updatedAt = ctx.nowIso;
  t.steps = [{ key: 'sign', status: 'skipped', value: null, doneBy: ctx.by, doneAt: ctx.nowIso, reason }];
  delete t.snoozedUntil;
  pushHist(t.history, H(ctx.nowIso, ctx.by, 'SKIP', reason));
  markTaskNotificationsRead(data, t.id);
  const m1 = OR.findManager1(users, t.owner);
  const to = Array.from(new Set([].concat(OR.marketingUsers(users), m1 ? [m1] : []))).filter((u) => u !== ctx.by);
  return { ok: true, task: t, notify: to, reason };
}

/** 稍後提醒：3 天後（台北日曆日）再提醒；只記錄。成功回 { ok, task, snoozedUntil } */
function snoozeTask(data, taskId, ctx) {
  const t = (data.dealTasks || []).find((x) => x && x.id === taskId);
  if (!t) return fail(404, 'TASK_NOT_FOUND', '找不到這張任務卡');
  if (t.status !== 'open') return fail(409, 'TASK_NOT_OPEN', '這張任務卡已經不是待處理狀態');
  const until = addDaysIso(taipeiDate(ctx.nowIso), SNOOZE_DAYS);
  t.snoozedUntil = until; t.updatedAt = ctx.nowIso;
  pushHist(t.history, H(ctx.nowIso, ctx.by, 'SNOOZE', '稍後提醒至 ' + until));
  markTaskNotificationsRead(data, t.id);
  return { ok: true, task: t, snoozedUntil: until };
}

/** BU 欄位（陣列、逗號字串或單一字串）→ 去空白的 Set */
function buSet(v) {
  const arr = Array.isArray(v) ? v : (typeof v === 'string' ? v.split(',') : []);
  return new Set(arr.map((x) => txt(x)).filter(Boolean));
}
/** 任務卡的 BU 參照：負責人在職且有 BU → 用負責人的 BU；否則（已停用或沒設）用商機（或納管項）的 BU */
function taskBuRef(data, t, users) {
  const o = OR.toMap(users)[t.owner];
  const own = buSet(o && o.bu);
  if (o && OR.isActiveUser(o) && own.size) return own;
  const opp = (data.opportunities || []).find((x) => x && x.id === t.oppId) || (data.lostOpportunities || []).find((x) => x && x.id === t.oppId);
  const fromOpp = buSet(opp && opp.bu);
  if (fromOpp.size) return fromOpp;
  const ids = t.maItemIds && t.maItemIds.length ? t.maItemIds : [t.maItemId];
  const fromItem = new Set();
  (data.maItems || []).forEach((it) => { if (it && ids.includes(it.id)) buSet(it.bu).forEach((b) => fromItem.add(b)); });
  return fromItem.size ? fromItem : own;
}
/** 目標是否可接手此任務卡：在職、角色為業務（user）、BU 與參照 BU 有交集（參照 BU 完全無法判斷時不限 BU）。回 null＝可；否則回失敗物件 */
function assigneeProblem(data, t, u, users) {
  if (!OR.isActiveUser(u)) return fail(400, 'ASSIGNEE_INACTIVE', '指定的接手人已停用');
  if (!OWNER_ROLES.includes(u.role)) return fail(400, 'ASSIGNEE_ROLE', '指定的接手人必須是業務（不能是主管或其他角色）');
  const ref = taskBuRef(data, t, users);
  if (ref.size) {
    const mine = buSet(u.bu);
    if (![...mine].some((b) => ref.has(b))) return fail(400, 'ASSIGNEE_BU', '指定的接手人與這張任務卡不同事業單位（BU）');
  }
  return null;
}
/** 改派候選人（管理員的下拉用）：在職業務、BU 與任務卡有交集、不含目前負責人。依顯示名稱排序 */
function reassignCandidates(data, t, users) {
  return Object.values(OR.toMap(users)).filter((u) => u && u.username !== t.owner && !assigneeProblem(data, t, u, users))
    .map((u) => ({ username: u.username, label: labelOf(users, u.username), bu: Array.from(buSet(u.bu)) }))
    .sort((a, b) => (a.label < b.label ? -1 : a.label > b.label ? 1 : (a.username < b.username ? -1 : 1)));
}

/** 管理員改派：目標須在職、是業務、BU 與任務卡有交集。成功回 { ok, task, from, to } */
function reassignTask(data, taskId, toUsername, ctx, users) {
  const t = (data.dealTasks || []).find((x) => x && x.id === taskId);
  if (!t) return fail(404, 'TASK_NOT_FOUND', '找不到這張任務卡');
  if (t.status !== 'open') return fail(409, 'TASK_NOT_OPEN', '只有待處理的任務卡可以改派');
  const map = OR.toMap(users);
  const u = map[txt(toUsername)];
  if (!u) return fail(400, 'ASSIGNEE_NOT_FOUND', '找不到指定的接手人');
  const prob = assigneeProblem(data, t, u, users);
  if (prob) return prob;
  if (u.username === t.owner) return fail(400, 'ASSIGNEE_SAME', '已經是這位負責人了');
  const from = t.owner;
  t.owner = u.username; t.updatedAt = ctx.nowIso;
  const ids = t.maItemIds && t.maItemIds.length ? t.maItemIds : [t.maItemId];
  (data.maItems || []).forEach((it) => {
    if (it && ids.includes(it.id)) { it.owner = u.username; it.updatedAt = ctx.nowIso; pushHist(it.history, H(ctx.nowIso, ctx.by, 'REASSIGN', from + ' → ' + u.username)); }
  });
  delete t.snoozedUntil;
  pushHist(t.history, H(ctx.nowIso, ctx.by, 'REASSIGN', from + ' → ' + u.username, { from, to: u.username }));
  markTaskNotificationsRead(data, t.id);   // 舊負責人的提醒清掉；新負責人由呼叫端另發
  return { ok: true, task: t, from, to: u.username };
}

/**
 * 刪除由任務卡建立的合約（呼叫端已把合約標 deleted）：納管項解除合約連結；若對應任務卡已完成（done）→ 重開（open、簽約步驟回 pending），讓業務可重新簽約。
 * 冪等：沒有仍連著這份合約的納管項就什麼都不做（回 touched:false）；任務卡已是 open 不會重複重開。License MA 類（沒有任務卡）只清連結。
 * 回 { touched, items:[被解除連結的納管項], task:重開的任務卡|null }。不存檔、不發通知、不寫稽核（呼叫端做）。
 */
function applyContractDeleted(data, contract, ctx) {
  const out = { touched: false, items: [], task: null };
  if (!contract || !contract.id) return out;
  const linked = (data.maItems || []).filter((it) => it && it.contractId === contract.id);
  if (!linked.length) return out;
  out.touched = true;
  linked.forEach((it) => {
    it.contractNo = ''; it.contractId = ''; it.updatedAt = ctx.nowIso;
    pushHist(it.history, H(ctx.nowIso, ctx.by, 'CONTRACT_DELETED', '合約已刪除，解除連結'));
    out.items.push(it);
  });
  const live = linked.filter((it) => it.status !== 'void');
  const t = live.length ? (data.dealTasks || []).find((x) => x && x.type === TASK_TYPE && x.oppId === live[0].oppId) : null;
  if (t && t.status === 'done') {
    t.status = 'open'; t.updatedAt = ctx.nowIso; t.createdAt = ctx.nowIso;   // 重開後逾期天數重新計算
    t.steps = [{ key: 'sign', status: 'pending', value: null, doneBy: '', doneAt: '', reason: '' }];
    const ids = live.map((it) => it.id);
    t.maItemIds = Array.from(new Set([].concat(t.maItemIds || [], ids))); if (!t.maItemId) t.maItemId = ids[0];
    delete t.snoozedUntil;
    pushHist(t.history, H(ctx.nowIso, ctx.by, 'REOPEN', '合約已刪除，重開簽約', { contractId: contract.id }));
    out.task = t;
  }
  return out;
}

/** 修改納管項期間（owner／管理員由呼叫端檢查）。成功回 { ok, item, changed, contract? } */
function editTerm(data, itemId, input, ctx) {
  const it = (data.maItems || []).find((x) => x && x.id === itemId);
  if (!it) return fail(404, 'ITEM_NOT_FOUND', '找不到這個納管項');
  if (it.status === 'void') return fail(409, 'ITEM_VOID', '這個納管項已作廢（商機不在成交階段），不能修改期間');
  const from = txt(input && input.termFrom), to = txt(input && input.termTo);
  const bad = QT.validateTerm(it.termKind, from, to);
  if (bad) return fail(400, bad.code, '服務期間：' + bad.reason);
  const reason = typeof (input && input.reason) === 'string' ? input.reason.trim() : '';
  if (Array.from(reason).length > REASON_MAX) return fail(400, 'REASON_TOO_LONG', '說明最長 ' + REASON_MAX + ' 個字');
  if (from === it.from && to === it.to) return { ok: true, item: it, changed: false };
  if (findItem(data, it.oppId, it.termKind, from, to)) return fail(409, 'DUPLICATE_TERM', '這個商機已經有相同期間的納管項');
  const old = { from: it.from, to: it.to };
  it.from = from; it.to = to; it.updatedAt = ctx.nowIso;
  pushHist(it.history, H(ctx.nowIso, ctx.by, ctx.action || 'TERM_EDIT', reason, { oldFrom: old.from, oldTo: old.to, from, to }));
  let contract = null;
  if (it.contractId) {
    contract = (data.contracts || []).find((c) => c && c.id === it.contractId && !c.deleted) || null;
    if (contract && contract.maItemId) {
      const sib = (data.maItems || []).filter((x) => x && x.contractId === contract.id && x.status !== 'void');
      if (sib.length) { contract.startDate = sib.map((x) => x.from).sort()[0]; contract.endDate = sib.map((x) => x.to).sort().slice(-1)[0]; }
    } else contract = null;
  }
  return { ok: true, item: it, changed: true, old, contract };
}

/**
 * 商機負責人被移轉（名單移轉／客戶池分配）：把這些商機底下的納管項與任務卡的 owner 一併改成新負責人（審查修正 M1）。
 * 只動 owner 剛好等於 fromOwner 的（管理員另行改派過的不覆蓋）；任何狀態的納管項／任務卡都跟著走（可視範圍以 owner 判斷）；寫 history 保留舊負責人。
 * 開啟中的任務卡：清掉延後、把舊負責人的「請簽約確認」通知標已讀，呼叫端再通知新負責人。不存檔、不發通知、不寫稽核。
 * 回 { items:[納管項], tasks:[任務卡], openTasks:[開啟中任務卡] }
 */
function transferOpps(data, oppIds, fromOwner, toOwner, ctx) {
  const out = { items: [], tasks: [], openTasks: [] };
  if (!oppIds || !oppIds.size || !fromOwner || !toOwner || fromOwner === toOwner) return out;
  (data.maItems || []).forEach((it) => {
    if (!it || !oppIds.has(it.oppId) || it.owner !== fromOwner) return;
    it.owner = toOwner; it.updatedAt = ctx.nowIso;
    if (!Array.isArray(it.history)) it.history = [];
    pushHist(it.history, H(ctx.nowIso, ctx.by, 'OWNER_TRANSFER', fromOwner + ' → ' + toOwner, { from: fromOwner, to: toOwner }));
    out.items.push(it);
  });
  (data.dealTasks || []).forEach((t) => {
    if (!t || t.type !== TASK_TYPE || !oppIds.has(t.oppId) || t.owner !== fromOwner) return;
    t.owner = toOwner; t.updatedAt = ctx.nowIso;
    if (!Array.isArray(t.history)) t.history = [];
    pushHist(t.history, H(ctx.nowIso, ctx.by, 'OWNER_TRANSFER', fromOwner + ' → ' + toOwner, { from: fromOwner, to: toOwner }));
    if (t.status === 'open') { delete t.snoozedUntil; markTaskNotificationsRead(data, t.id); out.openTasks.push(t); }
    out.tasks.push(t);
  });
  return out;
}

/**
 * 管理員改納管項的負責人（審查修正 M1：沒有任務卡的 License 類納管項原本完全沒有入口）。驗證同任務卡改派：目標在職、role 為 user、BU 與該案有交集。
 * 同一個商機下、目前負責人相同的納管項（非作廢）與該商機的任務卡一起改（不會出現納管項與任務卡負責人不同的半套）。
 * 成功回 { ok, item, items, tasks, openTasks, from, to }；失敗不修改 data。
 */
function reassignItem(data, itemId, toUsername, ctx, users) {
  const it = (data.maItems || []).find((x) => x && x.id === itemId);
  if (!it) return fail(404, 'ITEM_NOT_FOUND', '找不到這個納管項');
  if (it.status === 'void') return fail(409, 'ITEM_VOID', '這個納管項已作廢（商機不在成交階段），不能改負責人');
  const u = OR.toMap(users)[txt(toUsername)];
  if (!u) return fail(400, 'ASSIGNEE_NOT_FOUND', '找不到指定的接手人');
  const prob = assigneeProblem(data, { owner: it.owner, oppId: it.oppId, maItemId: it.id, maItemIds: [it.id] }, u, users);
  if (prob) return prob;
  if (u.username === it.owner) return fail(400, 'ASSIGNEE_SAME', '已經是這位負責人了');
  const from = it.owner;
  const group = (data.maItems || []).filter((x) => x && x.oppId === it.oppId && x.owner === from && x.status !== 'void');
  const oppIds = new Set([it.oppId]);
  const items = [], tasks = [], openTasks = [];
  group.forEach((g) => {
    g.owner = u.username; g.updatedAt = ctx.nowIso;
    if (!Array.isArray(g.history)) g.history = [];
    pushHist(g.history, H(ctx.nowIso, ctx.by, 'REASSIGN', from + ' → ' + u.username, { from, to: u.username }));
    items.push(g);
  });
  (data.dealTasks || []).forEach((t) => {
    if (!t || t.type !== TASK_TYPE || !oppIds.has(t.oppId) || t.owner !== from) return;
    t.owner = u.username; t.updatedAt = ctx.nowIso;
    if (!Array.isArray(t.history)) t.history = [];
    pushHist(t.history, H(ctx.nowIso, ctx.by, 'REASSIGN', from + ' → ' + u.username, { from, to: u.username }));
    if (t.status === 'open') { delete t.snoozedUntil; markTaskNotificationsRead(data, t.id); openTasks.push(t); }
    tasks.push(t);
  });
  return { ok: true, item: it, items, tasks, openTasks, from, to: u.username };
}

/**
 * 報價單核准後用 item-terms 改了品項服務期間 → 同步已成交的納管項（審查修正 M3）。呼叫端在 q.items 已換成新內容之後、db.save 之前呼叫。
 * changes＝[{ oldKind, oldFrom, oldTo, newKind, newFrom, newTo }]（只含「原本就有類型」的列）。
 * 對應規則：oppId（q.oppId）＋termKind＋原期間完全相符的納管項（不看 quoteId）；改期間走 editTerm（驗證、重複期間檢查、history 動作 TERM_SYNC、已簽約合約日期同步），所以與 PATCH ma-items term 同一套。
 * 不動的情況（回在 notes 供稽核註記）：報價單仍有其他品項使用原期間、同一個原期間被改成多個不同新期間、類型也被改了、對應不到納管項、editTerm 拒絕（作廢、重複期間）。
 * 商機沒有任何納管項（尚未成交）＝沒有東西要同步，回 considered:false、不產生註記。
 * 回 { considered, synced:[{ item, old:{from,to}, contract }], notes:[字串] }
 */
function syncQuoteTerms(data, q, changes, ctx) {
  const out = { considered: false, synced: [], notes: [] };
  const oppId = q && typeof q.oppId === 'string' ? q.oppId : '';
  if (!oppId || !Array.isArray(changes) || !changes.length) return out;
  const own = (data.maItems || []).filter((it) => it && it.oppId === oppId);
  if (!own.length) return out;
  out.considered = true;
  const keyOf = (k, f, t) => k + '|' + f + '|' + t;
  const rng = (f, t) => f + '～' + t;
  const still = new Set();
  (q.items || []).forEach((it) => { if (QI.isItemRow(it)) { const k = QT.parseKind(it.termKind); if (k) still.add(keyOf(k, txt(it.termFrom), txt(it.termTo))); } });
  const groups = new Map();
  changes.forEach((c) => {
    if (!c || !c.oldKind) return;
    const k = keyOf(c.oldKind, c.oldFrom, c.oldTo);
    if (!groups.has(k)) groups.set(k, { old: c, news: new Map() });
    groups.get(k).news.set(keyOf(c.newKind, c.newFrom, c.newTo), c);
  });
  groups.forEach((g, k) => {
    const o = g.old;
    const label = (QT.KIND_LABEL[o.oldKind] || o.oldKind) + ' ' + rng(o.oldFrom, o.oldTo);
    if (still.has(k)) { out.notes.push(label + '：報價單仍有其他品項使用這個原期間，納管項未動'); return; }
    if (g.news.size !== 1) { out.notes.push(label + '：同一個原期間被改成多個不同的新期間，無法判斷要同步哪一個，納管項未動'); return; }
    const nw = [...g.news.values()][0];
    if (nw.newKind !== o.oldKind) { out.notes.push(label + '：類型也被改了（' + (QT.KIND_LABEL[o.oldKind] || o.oldKind) + ' → ' + (QT.KIND_LABEL[nw.newKind] || nw.newKind) + '），納管項未動'); return; }
    const it = own.find((x) => x.termKind === o.oldKind && x.from === o.oldFrom && x.to === o.oldTo);
    if (!it) { out.notes.push(label + '：對應不到納管項（商機底下沒有這個類型與期間的納管項），未同步'); return; }
    const r = editTerm(data, it.id, { termFrom: nw.newFrom, termTo: nw.newTo, reason: '報價單 ' + txt(q.quoteNo) + ' 修改服務期間，同步納管項' }, { by: ctx.by, nowIso: ctx.nowIso, action: 'TERM_SYNC' });
    if (!r.ok) { out.notes.push(label + '：納管項未同步（' + r.message + '）'); return; }
    if (r.changed) out.synced.push({ item: it, old: r.old, contract: r.contract });
  });
  return out;
}

// ───────────────── 檢視 ─────────────────
function labelOf(users, username) {
  const u = OR.toMap(users)[username];
  return (u && (u.nickname || u.displayName || u.username)) || username || '';
}

/** 納管項摘要（任務卡內嵌）：不含金額、不含合約內容 */
function itemBrief(it, now) {
  return { id: it.id, termKind: it.termKind, from: it.from, to: it.to, daysLeft: daysLeft(it.to, now), status: it.status };
}

/** 任務卡檢視（依呼叫者，company／product 以商機現況為準，商機不在了就用任務卡上的快照） */
function taskView(data, t, users, now) {
  const opp = (data.opportunities || []).find((o) => o && o.id === t.oppId);
  const ids = t.maItemIds && t.maItemIds.length ? t.maItemIds : [t.maItemId];
  const items = (data.maItems || []).filter((it) => it && ids.includes(it.id)).map((it) => itemBrief(it, now));
  const days = daysSinceIso(t.createdAt, now);
  const owner = OR.toMap(users)[t.owner];
  const snoozed = !!(t.snoozedUntil && t.snoozedUntil > taipeiDate(now));
  return {
    id: t.id, type: t.type, oppId: t.oppId, status: t.status,
    company: txt(opp ? opp.company : t.company), product: txt(opp ? opp.product : t.product),
    owner: t.owner, ownerLabel: labelOf(users, t.owner), ownerActive: !!owner && owner.active !== false,
    createdAt: t.createdAt, daysOpen: days, overdue: days > OVERDUE_DAYS,
    snoozedUntil: t.snoozedUntil || '', snoozed,
    items,
  };
}

/** 我的開啟中任務卡（逾期的排前面，再依建立時間） */
function myTasks(data, me, users, now) {
  return (data.dealTasks || []).filter((t) => t && t.owner === me && t.status === 'open')
    .map((t) => taskView(data, t, users, now))
    .sort((a, b) => (b.daysOpen - a.daysOpen) || (a.createdAt < b.createdAt ? -1 : 1));
}

/** 逾期未簽（> OVERDUE_DAYS 天）的開啟中任務卡，限 visibleOwners（Set）內的負責人 */
function overdueTasks(data, visibleOwners, users, now) {
  return (data.dealTasks || []).filter((t) => t && t.status === 'open' && visibleOwners.has(t.owner) && daysSinceIso(t.createdAt, now) > OVERDUE_DAYS)
    .map((t) => taskView(data, t, users, now))
    .sort((a, b) => (b.daysOpen - a.daysOpen) || (a.createdAt < b.createdAt ? -1 : 1));
}

const FILTERS = ['all', 'pending_sign', 'd90', 'overdue'];
/** 待簽約＝追蹤中、尚無合約編號，而且「有開啟中的簽約任務」（License MA 類沒有任務卡的納管項不算待簽約，顯示為追蹤中） */
function isPendingSign(it, task) { return !!it && it.status === 'tracking' && !it.contractNo && !!task && task.status === 'open'; }
/**
 * 到期管理清單（已套過可視範圍的 items）。secretary＝true 時每筆只含 7 個欄位（company、product、termKind、from、to、daysLeft、ownerLabel）。
 * hideContract＝true（行銷、集團業務：他們看不到合約）時一般列不含 contractNo／contractId（審查修正 L1）。
 * 任何角色都不含金額、備註。依到期日（迄日）由近到遠。
 */
function listItems(items, users, now, opts) {
  const o = opts || {};
  const filter = FILTERS.includes(o.filter) ? o.filter : 'all';
  const taskByItem = new Map();
  (o.tasks || []).forEach((t) => { if (t && t.type === TASK_TYPE) (t.maItemIds || [t.maItemId]).forEach((id) => { const p = taskByItem.get(id); if (!p || p.status !== 'open') taskByItem.set(id, t); }); });
  const rows = items.filter((it) => it && it.status !== 'void').map((it) => ({ it, dl: daysLeft(it.to, now) })).filter(({ it, dl }) => {
    if (filter === 'all') return true;
    if (it.status !== 'tracking') return false;
    if (filter === 'pending_sign') return isPendingSign(it, taskByItem.get(it.id));
    if (filter === 'd90') return dl !== null && dl >= 0 && dl <= 90;
    return dl !== null && dl < 0;   // overdue
  });
  rows.sort((a, b) => (a.it.to < b.it.to ? -1 : a.it.to > b.it.to ? 1 : (a.it.id < b.it.id ? -1 : 1)));
  return rows.map(({ it, dl }) => {
    if (o.secretary) return { company: it.company, product: it.product, termKind: it.termKind, from: it.from, to: it.to, daysLeft: dl, ownerLabel: labelOf(users, it.owner) };
    const t = taskByItem.get(it.id);
    const ow = OR.toMap(users)[it.owner];
    const row = {
      id: it.id, oppId: it.oppId, quoteId: it.quoteId || '', company: it.company, product: it.product, termKind: it.termKind, from: it.from, to: it.to, daysLeft: dl,
      status: it.status, contractNo: it.contractNo || '', contractId: it.contractId || '', pendingSign: isPendingSign(it, t),
      taskId: t ? t.id : '', taskStatus: t ? t.status : '', owner: it.owner, ownerLabel: labelOf(users, it.owner), ownerActive: !!ow && ow.active !== false,
    };
    if (it.status === 'not_renew') row.notRenewReason = it.notRenewReason || '';   // P3：標記不續約的原因（只有 not_renew 狀態才有這個鍵；秘書列不含）
    if (o.hideContract) { delete row.contractNo; delete row.contractId; }
    return row;
  });
}

/**
 * 再成交時的期間提示（驗收後修正）：同一商機先前（含作廢）的納管項期間，最近異動的在前、相同期間只列一次。
 * items＝已套過可視範圍且屬於該商機的納管項。不含金額與備註。
 */
function termHint(items, now) {
  const list = (items || []).filter((it) => it && it.from && it.to).slice()
    .sort((a, b) => ((b.updatedAt || '') < (a.updatedAt || '') ? -1 : (b.updatedAt || '') > (a.updatedAt || '') ? 1 : 0));
  const seen = new Set(), out = [];
  list.forEach((it) => {
    const k = it.termKind + '|' + it.from + '|' + it.to;
    if (seen.has(k)) return;
    seen.add(k);
    const dl = daysLeft(it.to, now);
    out.push({ termKind: it.termKind, from: it.from, to: it.to, daysLeft: dl, expired: dl !== null && dl < 0, status: it.status });
  });
  return out;
}

/** 已成交（Won）但沒有任何追蹤中納管項、商品又是 License／MA 類的商機（Excel 匯入成交、開關開啟前就成交的舊單）。opps＝已套過可視範圍的商機 */
function pendingBackfill(data, opps, users) {
  const has = new Set((data.maItems || []).filter((it) => it && it.status !== 'void').map((it) => it.oppId));
  const out = [];
  (opps || []).forEach((o) => {
    if (!o || o.stage !== 'Won' || has.has(o.id)) return;
    const rule = PC.resolveTermRule(data.productCatalog, o.product);
    if (rule.termKind === 'none') return;
    out.push({ oppId: o.id, company: txt(o.company), product: txt(o.product), termKind: rule.termKind, startsTask: rule.startsTask, owner: o.owner, ownerLabel: labelOf(users, o.owner), achievedDate: o.achievedDate || '' });
  });
  return out.sort((a, b) => (a.company < b.company ? -1 : a.company > b.company ? 1 : 0));
}

module.exports = {
  MAX_ITEMS, MAX_TASKS, OVERDUE_DAYS, SNOOZE_DAYS, CONTRACT_NO_MAX, SKIP_REASON_MIN, TASK_TYPE, FILTERS, OWNER_ROLES,
  enabled, taipeiDate, daysLeft, daysSinceIso, addDaysIso, itemKey,
  termsFromQuotes, planWon, applyWon, applyUnwon, applyRestore,
  checkContractNo, contractNoTaken, signTask, skipTask, snoozeTask, reassignTask, editTerm,
  transferOpps, reassignItem, syncQuoteTerms,
  applyRenewals, applyContractDeleted, reassignCandidates, assigneeProblem, isPendingSign, termHint, markTaskNotificationsRead, taskView, myTasks, overdueTasks, listItems, pendingBackfill, labelOf, findTask, findItem,
};
