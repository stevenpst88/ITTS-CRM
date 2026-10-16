'use strict';
/**
 * 季度帳號與權限覆核報表（內控稽核第三波第二批 ③）。
 *
 * 純函式模組（不碰資料庫、不讀環境）：伺服器把帳號、主資料、角色功能矩陣、現在時間傳進來，回傳報表。
 *   buildReport()     列出每個帳號的角色／BU／權限／簽核職責／風險旗標；絕不輸出密碼雜湊、token、email 位址（只有「有無」布林）
 *   toCsv()           CSV（UTF-8 BOM；公式注入防護：= + - @ Tab CR 開頭的儲存格前面補單引號）
 *   snapshotHash()    報表「穩定部分」的 sha256（不含隨時間變動的數字，所以只要權限組態沒變就不變）
 *   簽核紀錄          validateSignoff()／makeSignoffRecord()／verifyChain()；紀錄只增不改，每筆帶 prevHash／recordHash 鏈
 *
 * 旗標（flags）：
 *   ADMIN                       管理員
 *   VIEW_ALL_ROLE               跨 BU 看全部資料的角色（executive／marketing／accounting_manager／finance_manager）
 *   DOWNLOAD_PERMISSION         可下載客戶名單
 *   INACTIVE_90D                啟用中但最後登入已超過 90 天（需有 lastLoginAt；追蹤開始前的登入不會記錄）
 *   NEVER_LOGGED_IN             啟用中但沒有任何登入紀錄（lastLoginAt 欄位上線前的老帳號在首次登入前都會出現，見 meta.lastLoginTrackingSince）
 *   PASSWORD_OLD                密碼超過 90 天沒換，或從未換過
 *   ROSTER_BUT_INACTIVE         在簽核名冊（總經理／董事長／董事會代核／報價章管理／成本填寫人）上，但帳號已停用
 *   SHARED_NAME_SUSPECT         帳號或顯示名稱像共用帳號（test／demo／shared／共用／測試…）
 *   NO_EMAIL                    啟用中但沒有 email（收不到簽核信）
 *   QUOTE_FEATURE_WITHOUT_CLEAR_DUTY  有「報價單」功能，但不能建單、也不是簽核人／名冊成員（盡力判斷，僅供人工複核）
 * 時間相關旗標（INACTIVE_90D／NEVER_LOGGED_IN／PASSWORD_OLD）不列入 snapshotHash，否則每天都會變。
 *
 * 雙人覆核備註：目前簽核只需一位管理員（reviewer 必須是 admin）。若日後要比照「核決設定兩人覆核」（APPROVAL_CONFIG_DUAL_REVIEW），
 * 可在 makeSignoffRecord 之前加「第二位管理員確認」的暫存狀態；紀錄結構已預留 co-reviewer 的位置（目前不使用）。
 */
const crypto = require('crypto');

const DAY = 86400000;
const VIEW_ALL_ROLES = ['executive', 'marketing', 'accounting_manager', 'finance_manager'];
const NO_CREATE_QUOTE_ROLES = ['executive', 'manager1', 'secretary', 'marketing', 'tecopm', 'groupsales', 'pool', 'accounting_manager', 'finance_manager', 'consult_manager_south', 'consult_manager_north'];
const QUOTE_DUTY_ROLES = ['admin', 'manager1', 'manager2', 'secretary'];   // 簽核人／代核／報價章
const ROSTER_KEYS = ['gm', 'chairman', 'boardProxy', 'sealManagers', 'costProviders'];
const ROSTER_LABEL = { gm: '總經理', chairman: '董事長', boardProxy: '董事會代核', sealManagers: '報價章管理', costProviders: '成本填寫人' };
const SHARED_RE = /(^|[^a-z])(test|tester|demo|shared|share|common|temp|tmp|guest|public|general|dummy|sample|training|trial)([^a-z]|$)|共用|公用|通用|測試|臨時|示範|訓練|試用|訪客/i;
const FLAG_ORDER = ['ADMIN', 'VIEW_ALL_ROLE', 'DOWNLOAD_PERMISSION', 'INACTIVE_90D', 'NEVER_LOGGED_IN', 'PASSWORD_OLD', 'ROSTER_BUT_INACTIVE', 'SHARED_NAME_SUSPECT', 'NO_EMAIL', 'QUOTE_FEATURE_WITHOUT_CLEAR_DUTY'];
const TIME_FLAGS = ['INACTIVE_90D', 'NEVER_LOGGED_IN', 'PASSWORD_OLD'];
const FLAG_LABEL = {
  ADMIN: '管理員', VIEW_ALL_ROLE: '跨 BU 全部可視', DOWNLOAD_PERMISSION: '可下載客戶名單', INACTIVE_90D: '90 天未登入', NEVER_LOGGED_IN: '無登入紀錄',
  PASSWORD_OLD: '密碼逾 90 天', ROSTER_BUT_INACTIVE: '名冊成員但帳號停用', SHARED_NAME_SUSPECT: '疑似共用帳號', NO_EMAIL: '無 email', QUOTE_FEATURE_WITHOUT_CLEAR_DUTY: '有報價單功能但無明確職責',
};

const str = (v) => (v == null ? '' : String(v));
const toMs = (v) => { const t = v ? new Date(v).getTime() : NaN; return Number.isFinite(t) ? t : null; };
const ageDays = (v, now) => { const t = toMs(v); return t == null ? null : Math.floor((now - t) / DAY); };
const buList = (bu) => (Array.isArray(bu) ? bu : (bu ? [bu] : [])).map(String);

/** 最後登入是否需要寫入（降低寫入頻率：同一人 6 小時內不重複寫） */
function shouldStampLogin(user, now, minGapMs) {
  const last = toMs(user && user.lastLoginAt);
  return last == null || now - last >= (minGapMs || 6 * 3600e3);
}

/**
 * opts: { users, data, featuresFor(role), userLabel(u), featureRouteMap?, now(ms), trackingSince? }
 */
function buildReport(opts) {
  const now = opts.now || Date.now();
  const users = (opts.users || []).filter(u => u && u.role !== 'pool');
  const data = opts.data || {};
  const roster = {};
  const qa = data.quoteApproval || {};
  const r = qa.roster || {};
  ROSTER_KEYS.forEach(k => { roster[k] = Array.isArray(r[k]) ? r[k].map(String) : []; });
  const label = opts.userLabel || ((u) => u.nickname || u.displayName || u.username);
  const hasQuoteMode = users.some(u => Object.prototype.hasOwnProperty.call(u, 'quoteMode'));
  const byName = new Map(users.map(u => [u.username, u]));
  const subCount = {};
  users.forEach(u => { if (u.supervisor && u.active !== false) subCount[u.supervisor] = (subCount[u.supervisor] || 0) + 1; });
  const pending = {};
  (data.quotations || []).forEach(q => {
    const ap = q && q.approval; if (!ap || ap.state !== 'pending' || q.deleted) return;
    const st = (ap.steps || [])[ap.cur]; if (!st || st.status !== 'pending') return;
    const names = new Set();
    if (st.assignee) names.add(st.assignee);
    if (st.tier === 'gm') roster.gm.forEach(n => names.add(n));
    if (st.tier === 'chairman') roster.chairman.forEach(n => names.add(n));
    if (st.tier === 'board') { roster.boardProxy.forEach(n => names.add(n)); users.forEach(u => { if (u.role === 'secretary') names.add(u.username); }); }
    names.forEach(n => { pending[n] = (pending[n] || 0) + 1; });
  });

  const accounts = users.map(u => {
    const role = u.role || 'user';
    const feats = (opts.featuresFor ? opts.featuresFor(role) : []) || [];
    const active = u.active !== false;
    const duties = ROSTER_KEYS.filter(k => roster[k].includes(u.username));
    const lastLoginAt = u.lastLoginAt || null;
    const pwAge = ageDays(u.passwordChangedAt, now);
    const emailPresent = !!(u.email && String(u.email).trim());
    const canCreateQuote = !NO_CREATE_QUOTE_ROLES.includes(role);
    const hasQuoteDuty = QUOTE_DUTY_ROLES.includes(role) || duties.length > 0 || canCreateQuote || (pending[u.username] || 0) > 0;
    const flags = [];
    if (role === 'admin') flags.push('ADMIN');
    if (VIEW_ALL_ROLES.includes(role)) flags.push('VIEW_ALL_ROLE');
    if (u.canDownloadContacts) flags.push('DOWNLOAD_PERMISSION');
    const llDays = ageDays(lastLoginAt, now);
    if (active && llDays != null && llDays > 90) flags.push('INACTIVE_90D');
    if (active && !lastLoginAt) flags.push('NEVER_LOGGED_IN');
    if (pwAge == null || pwAge > 90) flags.push('PASSWORD_OLD');
    if (!active && duties.length) flags.push('ROSTER_BUT_INACTIVE');
    if (SHARED_RE.test(str(u.username)) || SHARED_RE.test(str(u.displayName))) flags.push('SHARED_NAME_SUSPECT');
    if (active && !emailPresent) flags.push('NO_EMAIL');
    if (feats.includes('quotations') && !hasQuoteDuty) flags.push('QUOTE_FEATURE_WITHOUT_CLEAR_DUTY');
    flags.sort((a, b) => FLAG_ORDER.indexOf(a) - FLAG_ORDER.indexOf(b));
    const acc = {
      username: str(u.username),
      displayName: str(label(u)),
      role,
      bu: buList(u.bu),
      active,
      accessMode: role === 'admin' ? 'edit' : (u.accessMode === 'view' ? 'view' : 'edit'),
      canDownloadContacts: !!u.canDownloadContacts,
      canSetTargets: !!u.canSetTargets,
      supervisor: str(u.supervisor),
      features: feats.slice(),
      createdAt: u.createdAt || null,
      passwordChangedAt: u.passwordChangedAt || null,
      passwordAgeDays: pwAge,
      lastLoginAt,
      lastLoginDaysAgo: llDays,
      duties: { roster: duties, rosterLabels: duties.map(k => ROSTER_LABEL[k]), pendingApprovalSteps: pending[u.username] || 0, supervisesOthers: subCount[u.username] || 0 },
      emailPresent,
      flags,
    };
    if (hasQuoteMode) acc.quoteMode = str(u.quoteMode);
    return acc;
  }).sort((a, b) => (a.username < b.username ? -1 : a.username > b.username ? 1 : 0));

  const flagCounts = {};
  FLAG_ORDER.forEach(f => { flagCounts[f] = 0; });
  accounts.forEach(a => a.flags.forEach(f => { flagCounts[f]++; }));
  const rosterUnknown = [];
  ROSTER_KEYS.forEach(k => roster[k].forEach(n => { if (!byName.has(n)) rosterUnknown.push({ roster: k, username: n }); }));

  return {
    generatedAt: new Date(now).toISOString(),
    meta: {
      accountCount: accounts.length,
      activeCount: accounts.filter(a => a.active).length,
      hasQuoteMode,
      lastLoginTrackingSince: opts.trackingSince || null,
      rosterUnknown,   // 名冊上有、但帳號不存在的帳號名
      hashBasis: 'accounts（不含 passwordAgeDays／lastLoginAt／lastLoginDaysAgo／pendingApprovalSteps 及時間型旗標）',
    },
    summary: { accounts: accounts.length, active: accounts.filter(a => a.active).length, flagCounts, flaggedAccounts: accounts.filter(a => a.flags.length).length },
    flagLabels: FLAG_LABEL,
    accounts,
    featureRouteMap: opts.featureRouteMap || undefined,
  };
}

/** 報表穩定部分的正規化字串（鍵排序、陣列排序），給 snapshotHash 用 */
function canonical(report) {
  const sortObj = (v) => {
    if (Array.isArray(v)) return v.map(sortObj);
    if (v && typeof v === 'object') { const o = {}; Object.keys(v).sort().forEach(k => { o[k] = sortObj(v[k]); }); return o; }
    return v;
  };
  const accounts = report.accounts.map(a => {
    const c = Object.assign({}, a);
    delete c.passwordAgeDays; delete c.lastLoginAt; delete c.lastLoginDaysAgo;
    c.duties = { roster: a.duties.roster.slice().sort(), supervisesOthers: a.duties.supervisesOthers };
    c.features = a.features.slice().sort();
    c.flags = a.flags.filter(f => !TIME_FLAGS.includes(f));
    return c;
  });
  return JSON.stringify(sortObj({ accounts, rosterUnknown: report.meta.rosterUnknown.slice().sort((x, y) => (x.roster + x.username < y.roster + y.username ? -1 : 1)) }));
}
const sha256 = (s) => crypto.createHash('sha256').update(s).digest('hex');
const snapshotHash = (report) => sha256(canonical(report));

// ───────── CSV ─────────
const CSV_COLUMNS = [
  ['username', '帳號'], ['displayName', '顯示名稱'], ['role', '角色'], ['bu', 'BU'], ['quoteMode', '報價模式'], ['active', '啟用'], ['accessMode', '存取模式'],
  ['canDownloadContacts', '可下載名單'], ['canSetTargets', '可設定目標'], ['supervisor', '直屬主管'], ['features', '有效功能'], ['createdAt', '建立時間'],
  ['passwordChangedAt', '密碼更換時間'], ['passwordAgeDays', '密碼天數'], ['lastLoginAt', '最後登入'], ['lastLoginDaysAgo', '距今天數'],
  ['rosterDuties', '簽核名冊職責'], ['pendingApprovalSteps', '待簽關卡數'], ['supervisesOthers', '直屬部屬數'], ['emailPresent', '有 email'], ['flags', '風險旗標'],
];
/** 公式注入防護：以 = + - @ Tab CR 開頭的字串，在前面補一個單引號，Excel 就不會當公式執行 */
function csvCell(v) {
  let s = v == null ? '' : (typeof v === 'boolean' ? (v ? 'Y' : 'N') : String(v));
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}
function toCsv(report) {
  const cols = CSV_COLUMNS.filter(([k]) => k !== 'quoteMode' || report.meta.hasQuoteMode);
  const rowOf = (a) => cols.map(([k]) => {
    switch (k) {
      case 'bu': return a.bu.join('|');
      case 'features': return a.features.join('|');
      case 'rosterDuties': return a.duties.rosterLabels.join('|');
      case 'pendingApprovalSteps': return a.duties.pendingApprovalSteps;
      case 'supervisesOthers': return a.duties.supervisesOthers;
      case 'flags': return a.flags.join('|');
      default: return a[k];
    }
  }).map(csvCell).join(',');
  return '\uFEFF' + [cols.map(([, h]) => csvCell(h)).join(',')].concat(report.accounts.map(rowOf)).join('\r\n') + '\r\n';
}

// ───────── 簽核紀錄（只增不改）─────────
const QUARTER_RE = /^(20\d{2})Q([1-4])$/;
function validateSignoff(body, knownUsernames) {
  const b = body || {};
  const quarter = str(b.quarter).trim().toUpperCase();
  if (!QUARTER_RE.test(quarter)) return { error: '季別格式需為 YYYYQn，例如 2026Q4' };
  const comment = str(b.comment).trim();
  if (comment.length > 1000) return { error: '覆核意見最多 1000 字' };
  const exceptions = [];
  if (b.exceptions !== undefined && !Array.isArray(b.exceptions)) return { error: 'exceptions 必須是陣列' };
  if ((b.exceptions || []).length > 200) return { error: '例外清單最多 200 筆' };
  const known = new Set(knownUsernames || []);
  for (const e of (b.exceptions || [])) {
    const username = str(e && e.username).trim();
    const note = str(e && e.note).trim();
    if (!username || !known.has(username)) return { error: `例外清單中的帳號不存在：${username.slice(0, 40)}` };
    if (!note) return { error: `例外帳號 ${username} 需填寫說明` };
    if (note.length > 300) return { error: `例外帳號 ${username} 的說明最多 300 字` };
    exceptions.push({ username, note });
  }
  return { quarter, comment, exceptions };
}
function recordPayload(rec) {
  return JSON.stringify([rec.seq, rec.quarter, rec.by, rec.at, rec.accountCount, rec.flagCounts, rec.snapshotHash, rec.comment, rec.exceptions, rec.prevHash]);
}
function makeSignoffRecord({ list, v, by, at, report }) {
  const prev = (list || [])[(list || []).length - 1];
  const rec = {
    seq: (list || []).length + 1,
    quarter: v.quarter, by, at,
    accountCount: report.summary.accounts,
    flagCounts: report.summary.flagCounts,
    snapshotHash: snapshotHash(report),
    comment: v.comment,
    exceptions: v.exceptions,
    prevHash: prev ? prev.recordHash : '',
  };
  rec.recordHash = sha256(recordPayload(rec));
  return rec;
}
/** 驗證紀錄鏈：回傳 { ok, badAt }（badAt＝第一筆被竄改的 seq） */
function verifyChain(list) {
  let prev = '';
  for (let i = 0; i < (list || []).length; i++) {
    const r = list[i];
    if (r.seq !== i + 1 || r.prevHash !== prev || sha256(recordPayload(r)) !== r.recordHash) return { ok: false, badAt: r.seq || i + 1 };
    prev = r.recordHash;
  }
  return { ok: true, badAt: null };
}
/** 歷史列表（附「現況是否仍與簽核當時一致」）；currentHash＝目前報表的 snapshotHash */
function historyView(list, currentHash) {
  const chain = verifyChain(list);
  return { chain, records: (list || []).map(r => Object.assign({}, r, { matchesCurrent: r.snapshotHash === currentHash })) };
}

module.exports = {
  buildReport, canonical, snapshotHash, toCsv, csvCell, shouldStampLogin, validateSignoff, makeSignoffRecord, verifyChain, historyView,
  FLAG_ORDER, FLAG_LABEL, TIME_FLAGS, CSV_COLUMNS, SHARED_RE, VIEW_ALL_ROLES, ROSTER_KEYS, QUARTER_RE,
};
