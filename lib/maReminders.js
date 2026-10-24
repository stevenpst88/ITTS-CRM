'use strict';
/**
 * MA／License 到期「活規則」提醒引擎（P3a，MA_TRACKING=1 才啟用）。純邏輯：不碰 req、不呼叫 pushNotification、不寫稽核，
 * 資料與「現在時間」都由呼叫端注入，方便單元測試與日期推進模擬。HTTP／通知／稽核在 lib/maReminderRoutes.js。
 *
 * 資料（主資料 blob，各自獨立命名空間，掃描只改這幾個）：
 *   data.reminderRules = { rules:[{ id, name, enabled, whenDays, repeatEveryDays, notify[], condition, appliesTo }], updatedAt, updatedBy }
 *   data.reminderLog   = { v:1, entries:{ '<ruleId>|<maItemId>|<recipient>|<endDate>|<period>': '<ISO 發送時間>' }, lastSweepAt, lastSweepDate, lastSweepMode, lastStats }
 *   data.notifications 新增 type 'ma_reminder'（儲存型，按人彙總；帶 maItemIds 供自動解除比對）
 *
 * 規則語意：
 *   whenDays  整數 −365～365。正＝到期前 N 天、0＝到期當天、負＝到期後 N 天。dl＝距迄日還有幾天（台北日曆日，到期當天＝0、過期為負）。
 *   「已達門檻」＝ dl <= whenDays（不是剛好等於），所以 cron 漏跑也會補發。到期前的規則（whenDays>0）過了迄日就不再補發／顯示；
 *   到期當天或到期後的規則（whenDays<=0）一路有效，直到納管項離開追蹤中（續約、不續約、結束、作廢）或過期超過 MAX_OVERDUE_DAYS。
 *   repeatEveryDays 0＝只發一次；1–90＝達標後每 N 天一個新週期（只有 whenDays<=0 的規則能設）。period＝floor((whenDays − dl) / repeat)。
 *   condition：always｜no_renewal_quote（沒有任何有效的續約報價單）｜quote_not_submitted（有續約報價單但全部仍是草稿）｜not_renewed（納管項仍在追蹤中）。
 *   任何規則的共同前提：納管項 status==='tracking' 且迄日是真實日期。標記不續約、已續約、結束、作廢 → 全部提醒停止。
 *   去重鍵：ruleId|maItemId|recipient|endDate|period（迄日在鍵裡，所以修改期間後會重新提醒）。
 */
const crypto = require('crypto');
const MA = require('./maTracking');
const RC = require('./maRecipients');
const OR = require('./orgResolve');

const CONDITIONS = Object.freeze(['always', 'no_renewal_quote', 'quote_not_submitted', 'not_renewed']);
const NOTIFY_ROLES = RC.ROLES;
const KINDS = Object.freeze(['license', 'ma']);
const LIMITS = Object.freeze({ maxRules: 20, nameMax: 40, whenMin: -365, whenMax: 365, repeatMax: 90 });
const MAX_LOG = 20000;                // data.reminderLog.entries 上限
const LOG_KEEP_DAYS = 400;            // 迄日過了這麼多天的去重紀錄先清
const MAX_OVERDUE_DAYS = 400;         // 過期超過這麼多天仍在追蹤中的納管項不再提醒（避免永遠殘留）
const MAX_NOTIFS_PER_SWEEP = 200;     // 單次掃描最多發幾則儲存型通知（按人彙總，一人一則）
const NOTIFICATION_CAP = 500;         // data.notifications 全站上限（與 server.js pushNotification 相同）
const MAX_IDS_PER_NOTIFICATION = 200;
const LAZY_INTERVAL_MS = 60 * 60 * 1000;   // poll-bundle 補掃：每個實例至少間隔 60 分鐘
const LAZY_RETRY_MS = 60 * 1000;           // 今天還沒掃過時，同一實例最快每 60 秒再試一次
const NOTIF_TYPE = 'ma_reminder';
const ID_RE = /^[A-Za-z0-9_-]{1,40}$/;

const DEFAULT_RULES = Object.freeze([
  Object.freeze({ id: 'default-d90', name: '到期前 90 天：提醒業務並同步秘書', enabled: true, whenDays: 90, repeatEveryDays: 0, notify: Object.freeze(['owner', 'secretary']), condition: 'always', appliesTo: 'all' }),
  Object.freeze({ id: 'default-d45', name: '到期前 45 天仍無續約報價單：通知一級主管', enabled: true, whenDays: 45, repeatEveryDays: 0, notify: Object.freeze(['manager1']), condition: 'no_renewal_quote', appliesTo: 'all' }),
  Object.freeze({ id: 'default-d30', name: '到期前 30 天續約報價單未送簽：業務、一級主管、行銷', enabled: true, whenDays: 30, repeatEveryDays: 0, notify: Object.freeze(['owner', 'manager1', 'marketing']), condition: 'quote_not_submitted', appliesTo: 'all' }),
  Object.freeze({ id: 'default-after7', name: '到期後每 7 天追蹤未續約：業務、一級主管', enabled: true, whenDays: -1, repeatEveryDays: 7, notify: Object.freeze(['owner', 'manager1']), condition: 'not_renewed', appliesTo: 'all' }),
]);

const txt = (v) => String(v === undefined || v === null ? '' : v).trim();
const clone = (x) => JSON.parse(JSON.stringify(x));

// ───────────────── 規則：驗證／讀取／差異 ─────────────────
function toInt(v) {
  if (typeof v === 'number') return Number.isInteger(v) ? (v === 0 ? 0 : v) : NaN;
  if (typeof v === 'string' && /^-?\d{1,5}$/.test(v.trim())) { const n = Number(v.trim()); return n === 0 ? 0 : n; }
  return NaN;
}
const appliesKey = (a) => (a === 'all' ? 'all' : a.join(','));

/** 規則陣列驗證。回 { ok:true, value:{ rules:[…正規化後…] } } 或 { ok:false, status:400, code, error } */
function validateConfig(body) {
  const bad = (code, error) => ({ ok: false, status: 400, code, error });
  if (!body || typeof body !== 'object' || !Array.isArray(body.rules)) return bad('BAD_REMINDER_RULES', '設定內容無效（需要 rules 陣列）');
  if (body.rules.length > LIMITS.maxRules) return bad('TOO_MANY_RULES', '提醒規則最多 ' + LIMITS.maxRules + ' 條');
  const out = [];
  const ids = new Set();
  const dups = new Set();
  for (let i = 0; i < body.rules.length; i++) {
    const r = body.rules[i];
    const tag = '第 ' + (i + 1) + ' 條規則';
    if (!r || typeof r !== 'object' || Array.isArray(r)) return bad('BAD_REMINDER_RULES', tag + '格式無效');
    let id = r.id;
    if (id === undefined || id === null || id === '') id = 'r_' + crypto.randomBytes(5).toString('hex');
    else if (typeof id !== 'string' || !ID_RE.test(id)) return bad('BAD_RULE_ID', tag + '的代號只能是英數字、底線、連字號（1–40 字）');
    if (ids.has(id)) return bad('DUPLICATE_RULE_ID', tag + '的代號與其他規則重複');
    ids.add(id);
    const name = typeof r.name === 'string' ? r.name.trim() : '';
    if (!name) return bad('BAD_RULE_NAME', tag + '需要名稱');
    if (Array.from(name).length > LIMITS.nameMax) return bad('BAD_RULE_NAME', tag + '的名稱最長 ' + LIMITS.nameMax + ' 個字');
    if (/[\u0000-\u001f\u007f-\u009f<>]/.test(name)) return bad('BAD_RULE_NAME', tag + '的名稱含有不允許的字元');
    if (r.enabled !== undefined && typeof r.enabled !== 'boolean') return bad('BAD_RULE_ENABLED', tag + '的啟用狀態必須是 true／false');
    const when = toInt(r.whenDays);
    if (!Number.isInteger(when) || when < LIMITS.whenMin || when > LIMITS.whenMax) return bad('BAD_RULE_WHEN', tag + '的天數必須是 ' + LIMITS.whenMin + ' ～ ' + LIMITS.whenMax + ' 的整數（正數＝到期前、負數＝到期後）');
    const rep = (r.repeatEveryDays === undefined || r.repeatEveryDays === null || r.repeatEveryDays === '') ? 0 : toInt(r.repeatEveryDays);
    if (!Number.isInteger(rep) || rep < 0 || rep > LIMITS.repeatMax) return bad('BAD_RULE_REPEAT', tag + '的重複間隔必須是 0（只發一次）或 1 ～ ' + LIMITS.repeatMax + ' 天');
    if (rep > 0 && when > 0) return bad('BAD_RULE_REPEAT', tag + '：只有「到期當天或到期後」的規則可以設定重複提醒');
    if (!Array.isArray(r.notify) || !r.notify.length) return bad('BAD_RULE_NOTIFY', tag + '至少要選一個通知對象');
    if (!r.notify.every((x) => typeof x === 'string' && NOTIFY_ROLES.includes(x))) return bad('BAD_RULE_NOTIFY', tag + '的通知對象只能是 ' + NOTIFY_ROLES.join('、'));
    const notify = NOTIFY_ROLES.filter((x) => r.notify.includes(x));
    if (!CONDITIONS.includes(r.condition)) return bad('BAD_RULE_CONDITION', tag + '的條件只能是 ' + CONDITIONS.join('、'));
    let applies;
    if (r.appliesTo === undefined || r.appliesTo === null || r.appliesTo === 'all') applies = 'all';
    else if (Array.isArray(r.appliesTo) && r.appliesTo.length && r.appliesTo.every((x) => KINDS.includes(x))) {
      const set = KINDS.filter((k) => r.appliesTo.includes(k));
      applies = set.length === KINDS.length ? 'all' : set;
    } else return bad('BAD_RULE_APPLIES', tag + '的適用類型只能是 all 或 license／ma 的組合');
    const dk = r.condition + '|' + appliesKey(applies) + '|' + when;
    if (dups.has(dk)) return bad('DUPLICATE_RULE_WHEN', tag + '：相同條件與適用類型下，已有另一條規則使用同一個天數（' + when + '）');
    dups.add(dk);
    out.push({ id, name, enabled: r.enabled !== false, whenDays: when, repeatEveryDays: rep, notify, condition: r.condition, appliesTo: applies });
  }
  return { ok: true, value: { rules: out } };
}

/**
 * 生效的規則：沒設定（或儲存值不合法）→ 內建預設。不信任儲存值：整份重新驗證，任何一條不合法就整份退回預設。
 * 回 { rules, updatedAt, updatedBy, isDefault }
 */
function getConfig(data) {
  const c = data && data.reminderRules;
  if (c && typeof c === 'object' && Array.isArray(c.rules)) {
    const v = validateConfig({ rules: c.rules });
    if (v.ok) return { rules: v.value.rules, updatedAt: typeof c.updatedAt === 'string' ? c.updatedAt : null, updatedBy: typeof c.updatedBy === 'string' ? c.updatedBy : null, isDefault: false };
  }
  return { rules: clone(DEFAULT_RULES), updatedAt: null, updatedBy: null, isDefault: true };
}

function ruleLine(r) {
  return '「' + r.name + '」(' + r.id + ')｜' + (r.enabled ? '啟用' : '停用') + '｜' + (r.whenDays > 0 ? '到期前 ' + r.whenDays + ' 天' : r.whenDays === 0 ? '到期當天' : '到期後 ' + (-r.whenDays) + ' 天')
    + (r.repeatEveryDays ? '，每 ' + r.repeatEveryDays + ' 天重複' : '') + '｜對象 ' + r.notify.join('+') + '｜條件 ' + r.condition + '｜類型 ' + (r.appliesTo === 'all' ? 'all' : r.appliesTo.join('+'));
}
/** 稽核／覆核用的差異文字（新增、移除、變更各列一行；只列有變的規則） */
function configDiff(prev, next) {
  const a = new Map((prev.rules || []).map((r) => [r.id, r]));
  const b = new Map((next.rules || []).map((r) => [r.id, r]));
  const lines = [];
  b.forEach((r, id) => {
    if (!a.has(id)) lines.push('新增規則 ' + ruleLine(r));
    else if (JSON.stringify(a.get(id)) !== JSON.stringify(r)) lines.push('變更規則 ' + ruleLine(a.get(id)) + ' → ' + ruleLine(r));
  });
  a.forEach((r, id) => { if (!b.has(id)) lines.push('移除規則 ' + ruleLine(r)); });
  return lines;
}

// ───────────────── 續約報價單狀態 ─────────────────
/** 報價單是否「有效」（算作續約報價單）：未作廢（失效）、非舊制已拒絕、未刪除 */
function quoteCounts(q) { return !!q && !q.invalid && q.status !== 'rejected' && !q.deleted; }
/** 已送簽＝簽核中或已核准（草稿、退回、撤回都算未送簽） */
function quoteSubmitted(q) { return !!(q && q.approval && (q.approval.state === 'pending' || q.approval.state === 'approved')); }
/** Map<maItemId, { n, sent }>：每個納管項有效的續約報價單張數，以及是否有任何一張已送簽 */
function renewalIndex(data) {
  const m = new Map();
  ((data && data.quotations) || []).forEach((q) => {
    if (!q || typeof q.renewOf !== 'string' || !q.renewOf || !quoteCounts(q)) return;
    const e = m.get(q.renewOf) || { n: 0, sent: false };
    e.n++; if (quoteSubmitted(q)) e.sent = true;
    m.set(q.renewOf, e);
  });
  return m;
}
/** 'none'（沒有有效的續約報價單）｜'draft'（有但全部未送簽）｜'sent'（至少一張已送簽） */
function quoteStateOf(idx, itemId) {
  const e = idx.get(itemId);
  return !e || !e.n ? 'none' : (e.sent ? 'sent' : 'draft');
}

// ───────────────── 規則評估（單一來源：掃描、動態鈴鐺、試算器、時間軸都走這裡）─────────────────
function conditionHolds(condition, item, qs) {
  if (condition === 'always') return true;
  if (condition === 'no_renewal_quote') return qs === 'none';
  if (condition === 'quote_not_submitted') return qs === 'draft';
  if (condition === 'not_renewed') return !!item && item.status === 'tracking';
  return false;
}
const appliesTo = (rule, kind) => rule.appliesTo === 'all' || (Array.isArray(rule.appliesTo) && rule.appliesTo.includes(kind));

/** 此刻（距迄日 dl 天）這條規則是否在有效區間內（還不看條件）。reason 供試算器顯示 */
function windowState(rule, kind, dl) {
  if (!rule.enabled) return { active: false, reason: 'DISABLED' };
  if (!appliesTo(rule, kind)) return { active: false, reason: 'APPLIES' };
  if (dl < -MAX_OVERDUE_DAYS) return { active: false, reason: 'TOO_OLD' };
  if (dl > rule.whenDays) return { active: false, reason: 'NOT_REACHED' };
  if (rule.whenDays > 0 && dl < 0) return { active: false, reason: 'EXPIRED_WINDOW' };
  return { active: true, reason: '' };
}
/** 第幾個週期（只發一次的規則恆為 0） */
function periodOf(rule, dl) {
  if (!(rule.repeatEveryDays > 0)) return 0;
  return Math.floor((rule.whenDays - dl) / rule.repeatEveryDays);
}
const logKey = (ruleId, itemId, recipient, endDate, period) => [ruleId, itemId, recipient, endDate, period].join('|');

/** 納管項是否在提醒範圍內（共同前提）：追蹤中、類型正確、迄日是真實日期 */
function itemInScope(item, now) {
  if (!item || item.status !== 'tracking' || !KINDS.includes(item.termKind)) return null;
  const dl = MA.daysLeft(item.to, now);
  return dl === null ? null : dl;
}

/**
 * 計算「此刻」有哪些（納管項 × 規則）成立。回
 *   { today, entries:[{ item, dl, qs, hits:[{ rule, period, recipients:[{username, roles}], issues }] }], evaluated }
 * opts：onlyIds（Set，只算這些納管項）、noRecipients（略過收件人解析，時間軸／試算器用）。
 * 收件人在這裡「當下」解析（不綁死舊帳號）。
 */
function plan(data, users, now, rules, opts) {
  const o = opts || {};
  const map = OR.toMap(users);
  const idx = renewalIndex(data);
  const enabledRules = (rules || []).filter((r) => r && r.enabled);
  const maxWhen = enabledRules.reduce((m, r) => Math.max(m, r.whenDays), -Infinity);
  const entries = [];
  let evaluated = 0;
  ((data && data.maItems) || []).forEach((item) => {
    if (o.onlyIds && !(item && o.onlyIds.has(item.id))) return;
    const dl = itemInScope(item, now);
    if (dl === null) return;
    evaluated++;
    if (dl > maxWhen) return;
    const qs = quoteStateOf(idx, item.id);
    const hits = [];
    let resolved = null;
    enabledRules.forEach((rule) => {
      if (!windowState(rule, item.termKind, dl).active) return;
      if (!conditionHolds(rule.condition, item, qs)) return;
      let recipients = [], issues = [];
      if (!o.noRecipients) {
        if (!resolved) resolved = RC.resolveByRole(item, map);
        const x = RC.expandNotify(resolved, rule.notify);
        recipients = x.recipients; issues = x.issues;
      }
      hits.push({ rule, period: periodOf(rule, dl), recipients, issues });
    });
    if (hits.length) entries.push({ item, dl, qs, hits });
  });
  return { today: MA.taipeiDate(now), entries, evaluated };
}

/** plan → Map<username, Map<itemId, { item, dl, qs, hits:[…] }>>：每個人此刻有哪些納管項需要處理 */
function activeByUser(pl) {
  const out = new Map();
  pl.entries.forEach((e) => {
    e.hits.forEach((h) => {
      h.recipients.forEach((r) => {
        if (!out.has(r.username)) out.set(r.username, new Map());
        const m = out.get(r.username);
        if (!m.has(e.item.id)) m.set(e.item.id, { item: e.item, dl: e.dl, qs: e.qs, hits: [], roles: new Set() });
        const cell = m.get(e.item.id);
        cell.hits.push(h);
        r.roles.forEach((x) => cell.roles.add(x));
      });
    });
  });
  return out;
}

// ───────────────── 動態鈴鐺（poll-bundle 的 maReminders）─────────────────
const levelOf = (dl) => (dl < 0 ? 'danger' : dl <= 30 ? 'warn' : 'info');
function dayText(dl) { return dl > 0 ? '剩 ' + dl + ' 天到期' : dl === 0 ? '今天到期' : '已逾期 ' + (-dl) + ' 天'; }
const QS_TEXT = { none: '尚未建立續約報價單', draft: '續約報價單尚未送簽', sent: '續約報價單已送簽' };

/**
 * 指定使用者此刻的動態提醒清單（依剩餘天數由少到多，同天以公司名排序）。
 * 一個納管項只出一列（命中多條規則時取「最後到達的門檻」＝whenDays 最小者當主規則）。
 * 內容只有公司名、品項、日期與狀態，沒有金額、合約編號。秘書的欄位更少（不給納管項 id 與規則細節）。
 * opts：secretary（bool）、labelOf(username)、visible(item)→bool（可視範圍過濾，與 ma-items 同一套）、limit。
 */
function dynamicFor(pl, username, opts) {
  const o = opts || {};
  const mine = activeByUser(pl).get(username);
  if (!mine) return [];
  const rows = [];
  mine.forEach((cell) => {
    if (o.visible && !o.visible(cell.item)) return;
    const it = cell.item;
    const lead = cell.hits.slice().sort((a, b) => a.rule.whenDays - b.rule.whenDays || (a.rule.id < b.rule.id ? -1 : 1))[0];
    const co = it.company || '（未命名公司）';
    const ownerLabel = (o.labelOf && o.labelOf(it.owner)) || it.owner || '';
    const title = '⏰ ' + co + '：' + (it.product || (it.termKind === 'license' ? 'License' : 'MA')) + ' ' + dayText(cell.dl);
    const body = '服務期間 ' + it.from + ' ～ ' + it.to + '｜' + (QS_TEXT[cell.qs] || '') + (username !== it.owner ? '｜業務：' + ownerLabel : '');
    const base = { type: NOTIF_TYPE, title, body, days: cell.dl, level: levelOf(cell.dl), company: co, product: it.product || '', termKind: it.termKind, from: it.from, to: it.to, ownerLabel };
    if (o.secretary) {
      base.id = 'ma_' + crypto.createHash('sha1').update(String(it.id)).digest('hex').slice(0, 12);
    } else {
      base.id = 'ma_' + it.id;
      base.itemId = it.id;
      base.quoteState = cell.qs;
      base.rule = { id: lead.rule.id, name: lead.rule.name };
      base.ruleCount = cell.hits.length;
      base.roles = RC.ROLES.filter((x) => cell.roles.has(x));
    }
    rows.push(base);
  });
  rows.sort((a, b) => a.days - b.days || (a.company < b.company ? -1 : a.company > b.company ? 1 : 0) || (a.id < b.id ? -1 : 1));
  return o.limit ? rows.slice(0, o.limit) : rows;
}

// ───────────────── 通知自動解除 ─────────────────
/**
 * 把「條件已不成立」的未讀 ma_reminder 通知標為已讀。n.maItemIds 內沒有任何一項對該收件人仍成立 → 標已讀。
 * filter(n)：只處理符合的通知（poll-bundle 只處理自己的；單一事件只處理含該納管項的）。回標記數。
 * 注意：只動 type==='ma_reminder'；沒有 maItemIds 的（無法比對）不動。
 */
function reconcileNotifications(data, users, now, rules, filter) {
  const list = (data && data.notifications) || [];
  const cand = list.filter((n) => n && n.type === NOTIF_TYPE && !n.read && Array.isArray(n.maItemIds) && n.maItemIds.length && (!filter || filter(n)));
  if (!cand.length) return 0;
  const ids = new Set();
  cand.forEach((n) => n.maItemIds.forEach((i) => ids.add(i)));
  const act = activeByUser(plan(data, users, now, rules, { onlyIds: ids }));
  let marked = 0;
  cand.forEach((n) => {
    const mine = act.get(n.to);
    if (!mine || !n.maItemIds.some((i) => mine.has(i))) { n.read = true; marked++; }
  });
  return marked;
}

/** 超過 NOTIFICATION_CAP 時優先丟「已讀且最舊」的，再丟最舊的（陣列約定：新的在前） */
function trimNotifications(list, cap) {
  const max = cap || NOTIFICATION_CAP;
  if (list.length <= max) return list;
  let over = list.length - max;
  const drop = new Set();
  for (let i = list.length - 1; i >= 0 && over > 0; i--) if (list[i] && list[i].read) { drop.add(i); over--; }
  for (let i = list.length - 1; i >= 0 && over > 0; i--) if (!drop.has(i)) { drop.add(i); over--; }
  return list.filter((_, i) => !drop.has(i));
}

// ───────────────── 去重紀錄 ─────────────────
function ensureLog(data) {
  if (!data.reminderLog || typeof data.reminderLog !== 'object' || Array.isArray(data.reminderLog)) data.reminderLog = { v: 1, entries: {} };
  if (!data.reminderLog.entries || typeof data.reminderLog.entries !== 'object' || Array.isArray(data.reminderLog.entries)) data.reminderLog.entries = {};
  return data.reminderLog;
}
function keyParts(key) {
  const p = String(key).split('|');
  return { period: p[p.length - 1], endDate: p[p.length - 2], ruleId: p[0], itemId: p[1], recipient: p.slice(2, p.length - 2).join('|') };
}
/** 清理：迄日 + LOG_KEEP_DAYS 天前的先清；仍超過上限（含即將新增的 incoming 筆）再丟最舊的。回清掉幾筆 */
function pruneLog(log, today, incoming) {
  const keys = Object.keys(log.entries);
  let removed = 0;
  for (const k of keys) {
    const ed = keyParts(k).endDate;
    const d = MA.daysLeft(ed, today + 'T12:00:00+08:00');
    if (d !== null && d < -LOG_KEEP_DAYS) { delete log.entries[k]; removed++; }
  }
  const left = Object.keys(log.entries);
  const room = MAX_LOG - (incoming || 0);
  if (left.length > room) {
    left.sort((a, b) => (log.entries[a] < log.entries[b] ? -1 : log.entries[a] > log.entries[b] ? 1 : (a < b ? -1 : 1)));
    left.slice(0, left.length - room).forEach((k) => { delete log.entries[k]; removed++; });
  }
  return removed;
}

// ───────────────── 掃描 ─────────────────
function emptyStats() { return { evaluated: 0, fired: 0, recipients: 0, skippedDup: 0, noRecipient: 0, capped: 0 }; }

/**
 * 掃描：找出「已達門檻、條件成立、尚未發過」的（規則 × 納管項 × 收件人），按人彙總發一則儲存型通知，同一次寫入去重標記。
 * opts：{ now, mode:'cron'|'lazy'|…, dryRun, rules?, newId?(→string), maxNotifications? }
 * 修改 data（dryRun 時不修改）：data.reminderLog、data.notifications。呼叫端負責 db.save(data)（同一次）與稽核／Web Push。
 * 回 { stats, created:[通知], audits:[{ action, target, detail }], changed, today }
 */
function sweep(data, users, opts) {
  const o = opts || {};
  const now = o.now === undefined ? Date.now() : o.now;
  const dry = !!o.dryRun;
  const cap = Number.isInteger(o.maxNotifications) ? o.maxNotifications : MAX_NOTIFS_PER_SWEEP;
  const newId = o.newId || (() => crypto.randomUUID());
  const rules = o.rules || getConfig(data).rules;
  const nowIso = new Date(typeof now === 'number' ? now : (now instanceof Date ? now.getTime() : Date.parse(now))).toISOString();
  const stats = emptyStats();
  const audits = [];
  const created = [];
  const view = dry ? { entries: Object.assign({}, (data.reminderLog && data.reminderLog.entries) || {}) } : ensureLog(data);

  const pl = plan(data, users, now, rules);
  stats.evaluated = pl.evaluated;
  const today = pl.today;
  const byRecipient = new Map();   // username -> [{ key, item, rule, dl, period }]
  const noRecipAudited = new Set();
  pl.entries.forEach((e) => {
    e.hits.forEach((h) => {
      if (!h.recipients.length) {
        stats.noRecipient++;
        const k = h.rule.id + '|' + e.item.id + '|' + h.period;
        if (!noRecipAudited.has(k)) { noRecipAudited.add(k); audits.push({ action: 'MA_REMINDER_NO_RECIPIENT', target: e.item.company || '', detail: '規則「' + h.rule.name + '」找不到任何收件人（連管理員都沒有）｜納管項 ' + e.item.id + '｜' + e.item.product }); }
        return;
      }
      h.recipients.forEach((r) => {
        const key = logKey(h.rule.id, e.item.id, r.username, e.item.to, h.period);
        if (view.entries[key]) { stats.skippedDup++; return; }
        if (!byRecipient.has(r.username)) byRecipient.set(r.username, []);
        byRecipient.get(r.username).push({ key, item: e.item, rule: h.rule, dl: e.dl, period: h.period, issues: h.issues });
      });
    });
  });

  // 自動解除：條件已不成立的未讀 ma_reminder 通知標已讀（在新增本次通知之前做）
  const act = activeByUser(pl);
  const stale = ((data && data.notifications) || []).filter((n) => n && n.type === NOTIF_TYPE && !n.read && Array.isArray(n.maItemIds) && n.maItemIds.length
    && !(act.get(n.to) && n.maItemIds.some((i) => act.get(n.to).has(i))));
  stats.resolved = stale.length;
  if (!dry) stale.forEach((n) => { n.read = true; });

  const names = [...byRecipient.keys()].sort();
  const take = names.slice(0, Math.max(0, cap));
  stats.capped = names.length - take.length;
  const noIssueAudited = new Set();
  take.forEach((username) => {
    const pend = byRecipient.get(username);
    const items = new Map();
    pend.forEach((p) => { if (!items.has(p.item.id)) items.set(p.item.id, p); });
    const list = [...items.values()].sort((a, b) => a.dl - b.dl || (a.item.company < b.item.company ? -1 : 1));
    const n = list.length;
    const lines = list.slice(0, 5).map((p) => (p.item.company || '（未命名公司）') + '｜' + (p.item.product || '') + '｜到期 ' + p.item.to + '（' + dayText(p.dl) + '）');
    const body = lines.join('；') + (n > 5 ? '…等共 ' + n + ' 項' : '') + '。請開啟「MA／License 到期管理」處理。';
    const note = { id: newId(), to: username, type: NOTIF_TYPE, title: '📌 你有 ' + n + ' 項 MA／License 需要處理', body, refId: today, read: false, createdAt: nowIso, maItemIds: list.slice(0, MAX_IDS_PER_NOTIFICATION).map((p) => p.item.id) };
    created.push(note);
    stats.recipients++;
    pend.forEach((p) => {
      stats.fired++;
      if (!dry) view.entries[p.key] = nowIso;
      (p.issues || []).forEach((is) => {
        const ak = p.rule.id + '|' + p.item.id + '|' + p.period + '|' + is.code;
        if (noIssueAudited.has(ak)) return;
        noIssueAudited.add(ak);
        if (is.code === 'NO_SECRETARY' || is.code === 'NO_MANAGER1' || is.code === 'NO_OWNER_RECIPIENT') {
          audits.push({ action: 'MA_REMINDER_NO_RECIPIENT', target: p.item.company || '', detail: '找不到' + ({ NO_SECRETARY: '同 BU 的在職秘書', NO_MANAGER1: '在職的一級主管', NO_OWNER_RECIPIENT: '可通知的負責業務或一級主管' }[is.code]) + '，已改通知管理員｜規則「' + p.rule.name + '」｜納管項 ' + p.item.id + '｜' + p.item.product });
        }
      });
    });
  });

  let changed = false;
  if (!dry) {
    if (!Array.isArray(data.notifications)) data.notifications = [];
    if (created.length) {
      data.notifications = trimNotifications(created.slice().reverse().concat(data.notifications));
      changed = true;
    }
    if (stale.length) changed = true;
    pruneLog(view, today, 0);
    view.lastSweepAt = nowIso; view.lastSweepDate = today; view.lastSweepMode = o.mode || 'cron';
    view.lastStats = Object.assign({}, stats);
    changed = true;
  }
  return { stats, created, audits, changed, today };
}

/** poll-bundle 補掃要不要執行（純函式，方便測）。state：{ instLastAttemptMs, lastSweepAt, lastSweepDate }（後兩者是持久化的） */
function shouldLazySweep(nowMs, state) {
  const s = state || {};
  const today = MA.taipeiDate(nowMs);
  const sinceAttempt = nowMs - (s.instLastAttemptMs || 0);
  if (s.lastSweepDate !== today) return sinceAttempt >= LAZY_RETRY_MS;   // 今天還沒掃過：盡快補（失敗時最快每 60 秒再試）
  const lastMs = s.lastSweepAt ? Date.parse(s.lastSweepAt) : 0;
  return sinceAttempt >= LAZY_INTERVAL_MS && nowMs - (Number.isFinite(lastMs) ? lastMs : 0) >= LAZY_INTERVAL_MS;
}

// ───────────────── 試算器／時間軸 ─────────────────
/**
 * 試算：給一組規則（不傳＝目前生效的）與情境（距迄日幾天、續約報價單狀態、是否已標不續約），回「此刻每個角色會看到哪些提醒」。
 * 與實際引擎共用同一組函式（plan／windowState／conditionHolds）：這裡只組一份合成資料。
 * opts.item：真實納管項（有就用它的收件人解析，顯示帳號顯示名）；opts.users；opts.now；opts.labelOf。
 */
function simulate(rules, scenario, opts) {
  const o = opts || {};
  const now = o.now === undefined ? Date.now() : o.now;
  const dl = scenario.daysLeft;
  const today = MA.taipeiDate(now);
  const to = MA.addDaysIso(today, dl);
  const real = o.item || null;
  const kind = (scenario.termKind && KINDS.includes(scenario.termKind)) ? scenario.termKind : (real && KINDS.includes(real.termKind) ? real.termKind : 'ma');
  const item = {
    id: real ? real.id : 'sim-item', oppId: real ? real.oppId : 'sim-opp', owner: real ? real.owner : '', bu: real ? real.bu : undefined,
    company: real ? real.company : '（示範客戶）', product: real ? real.product : '（示範品項）', termKind: kind,
    from: MA.addDaysIso(to, -364), to, status: scenario.notRenew ? 'not_renew' : 'tracking',
  };
  const quotes = scenario.quoteState === 'none' ? [] : [{ id: 'sim-q', renewOf: item.id, approval: { state: scenario.quoteState === 'sent' ? 'pending' : 'none' } }];
  const data = { maItems: [item], quotations: quotes };
  // 試算用的規則一律當成啟用中的設定本身（不改 enabled）；收件人用真實使用者（有帶納管項時）
  const users = real ? (o.users || []) : [];
  // 沒帶真實納管項時不解析帳號：用一個「只有角色」的占位解析
  const pl = real ? plan(data, users, now, rules) : plan(data, [], now, rules, { noRecipients: true });
  const labelOf = o.labelOf || ((u) => u);
  const hits = pl.entries.length ? pl.entries[0].hits : [];
  const byRole = {};
  NOTIFY_ROLES.forEach((role) => { byRole[role] = hits.filter((h) => h.rule.notify.includes(role)).map((h) => ({ ruleId: h.rule.id, name: h.rule.name, whenDays: h.rule.whenDays, repeatEveryDays: h.rule.repeatEveryDays, condition: h.rule.condition })); });
  const recipients = {};
  if (real) {
    const r = RC.resolveByRole(item, OR.toMap(users));
    NOTIFY_ROLES.forEach((role) => { recipients[role] = r.byRole[role].map((u) => ({ username: u, label: labelOf(u) })); });
  }
  const qs = scenario.quoteState === 'sent' || scenario.quoteState === 'draft' ? scenario.quoteState : 'none';
  const inScope = itemInScope(item, now) !== null;
  const rulesState = (rules || []).map((r) => {
    const w = windowState(r, kind, dl);
    const cond = conditionHolds(r.condition, item, qs);
    let reason = '';
    if (!inScope) reason = 'NOT_TRACKING';
    else if (!w.active) reason = w.reason;
    else if (!cond) reason = 'CONDITION';
    return { ruleId: r.id, name: r.name, whenDays: r.whenDays, condition: r.condition, active: inScope && w.active && cond, reason, period: periodOf(r, dl) };
  });
  return { daysLeft: dl, quoteState: qs, notRenew: !!scenario.notRenew, termKind: kind, tracking: inScope, byRole, recipients, rules: rulesState, activeRuleIds: rulesState.filter((x) => x.active).map((x) => x.ruleId) };
}

/** 觸發日期：到期日 − whenDays 天（whenDays 為負＝到期後） */
function triggerDate(item, rule) { return MA.addDaysIso(item.to, -rule.whenDays); }

/**
 * 單一納管項的規則時間軸：每條規則的觸發日期、條件目前是否成立、是否在有效區間、已發送紀錄。
 * opts：now、users、labelOf、secretary（不給收件人名單）。
 */
function timeline(data, item, rules, opts) {
  const o = opts || {};
  const now = o.now === undefined ? Date.now() : o.now;
  const dl = MA.daysLeft(item.to, now);
  const idx = renewalIndex(data);
  const qs = quoteStateOf(idx, item.id);
  const entries = (data.reminderLog && data.reminderLog.entries) || {};
  const prefixMatch = (key, ruleId) => { const p = keyParts(key); return p.ruleId === ruleId && p.itemId === item.id && p.endDate === item.to; };
  const logKeys = Object.keys(entries);
  const scope = itemInScope(item, now) !== null;
  return rules.map((rule) => {
    const applies = appliesTo(rule, item.termKind);
    const w = dl === null ? { active: false, reason: 'BAD_DATE' } : windowState(rule, item.termKind, dl);
    const cond = conditionHolds(rule.condition, item, qs);
    const sent = logKeys.filter((k) => prefixMatch(k, rule.id)).map((k) => { const p = keyParts(k); return { to: p.recipient, at: entries[k], period: Number(p.period) }; })
      .sort((a, b) => (a.at < b.at ? -1 : a.at > b.at ? 1 : 0));
    const row = {
      ruleId: rule.id, name: rule.name, enabled: rule.enabled, appliesToThis: applies, whenDays: rule.whenDays, repeatEveryDays: rule.repeatEveryDays, condition: rule.condition, notify: rule.notify,
      triggerDate: triggerDate(item, rule), reached: dl !== null && dl <= rule.whenDays, conditionNow: cond, activeNow: scope && w.active && cond, reason: !scope ? 'NOT_TRACKING' : (!w.active ? w.reason : (!cond ? 'CONDITION' : '')),
      sentCount: sent.length, lastSentAt: sent.length ? sent[sent.length - 1].at : '',
    };
    if (!o.secretary) row.sent = sent.slice(-20).map((s) => ({ to: s.to, toLabel: o.labelOf ? o.labelOf(s.to) : s.to, at: s.at, period: s.period }));
    return row;
  });
}

module.exports = {
  CONDITIONS, NOTIFY_ROLES, KINDS, LIMITS, DEFAULT_RULES, NOTIF_TYPE, MAX_LOG, LOG_KEEP_DAYS, MAX_OVERDUE_DAYS, MAX_NOTIFS_PER_SWEEP, NOTIFICATION_CAP, LAZY_INTERVAL_MS, LAZY_RETRY_MS,
  validateConfig, getConfig, configDiff,
  quoteCounts, quoteSubmitted, renewalIndex, quoteStateOf, conditionHolds, appliesTo, windowState, periodOf, logKey, keyParts, itemInScope,
  plan, activeByUser, dynamicFor, reconcileNotifications, trimNotifications, ensureLog, pruneLog,
  sweep, shouldLazySweep, simulate, timeline, triggerDate, dayText,
};
