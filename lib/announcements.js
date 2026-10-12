'use strict';
// 公告（announcements）純邏輯：驗證、時區換算、適用判斷、統計。
// 不碰 I/O（server.js 負責讀寫 data.announcements），方便單元測試。
//
// 資料形狀（data.announcements 陣列的一筆）：
//   { id, version, title, body, mode:'once'|'everyLogin', severity:'normal'|'important',
//     roles:[]（空＝全部角色）, startAt:ISO|null, endAt:ISO|null, enabled, priority,
//     readBy:{ [username]: { version, at } }, createdAt, createdBy, updatedAt, updatedBy }

const MAX_ANNOUNCEMENTS = 200;
const MAX_TITLE = 100;
const MAX_BODY = 4000;
const MODES = ['once', 'everyLogin'];
const SEVERITIES = ['normal', 'important'];
const READ_MIN_MS = 4000; // 強制閱讀 5 秒（前端 READ_SECONDS）；伺服器容忍 1 秒誤差，早於此的 ack 一律拒絕
const TAIPEI_OFFSET_MS = 8 * 3600 * 1000; // Asia/Taipei 全年 UTC+8、無日光節約

/** 去掉控制字元（保留 \n、\t），統一換行。純文字，不做 HTML 跳脫（由顯示端跳脫）。 */
function cleanText(val) {
  if (val === null || val === undefined) return '';
  return String(val)
    .replace(/\r\n?/g, '\n')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .trim();
}

/**
 * 解析管理員輸入的時間。
 *  - 空字串 / null → null（不限制）
 *  - 'YYYY-MM-DDTHH:mm'（datetime-local，無時區）→ 以 Asia/Taipei 解讀
 *  - 其他可被 Date 解析的字串（含 Z 或 +08:00）→ 照字面
 * 回傳 { ok:true, value: ISO|null } 或 { ok:false }
 */
function parseTaipeiInput(v) {
  if (v === null || v === undefined || v === '') return { ok: true, value: null };
  const s = String(v).trim();
  if (!s) return { ok: true, value: null };
  let t;
  const m = /^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})(?::(\d{2}))?$/.exec(s);
  if (m) {
    t = Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +(m[6] || 0)) - TAIPEI_OFFSET_MS;
    // 防 2026-02-31 這類被 Date.UTC 自動進位
    const back = new Date(t + TAIPEI_OFFSET_MS);
    if (back.getUTCMonth() !== +m[2] - 1 || back.getUTCDate() !== +m[3]) return { ok: false };
  } else {
    t = Date.parse(s);
  }
  if (!Number.isFinite(t)) return { ok: false };
  return { ok: true, value: new Date(t).toISOString() };
}

/** ISO → 'YYYY-MM-DDTHH:mm'（台北時間，給 datetime-local 用） */
function toTaipeiLocal(iso) {
  if (!iso) return '';
  const t = Date.parse(iso);
  if (!Number.isFinite(t)) return '';
  return new Date(t + TAIPEI_OFFSET_MS).toISOString().slice(0, 16);
}

/**
 * 驗證並正規化管理員送來的欄位。partial=true（PUT）只處理有帶的欄位。
 * 回傳 { error } 或 { fields }（fields 只含已驗證欄位）。
 */
function validateInput(b, knownRoles, partial) {
  b = b || {};
  const f = {};
  const has = k => Object.prototype.hasOwnProperty.call(b, k);

  if (!partial || has('title')) {
    const title = cleanText(b.title).replace(/\n/g, ' ');
    if (!title) return { error: '請輸入公告標題' };
    if (title.length > MAX_TITLE) return { error: `標題不可超過 ${MAX_TITLE} 字` };
    f.title = title;
  }
  if (!partial || has('body')) {
    const body = cleanText(b.body);
    if (!body) return { error: '請輸入公告內容' };
    if (body.length > MAX_BODY) return { error: `內容不可超過 ${MAX_BODY} 字` };
    f.body = body;
  }
  if (!partial || has('mode')) {
    const mode = b.mode === undefined ? 'once' : b.mode;
    if (!MODES.includes(mode)) return { error: '公告模式不正確（once／everyLogin）' };
    f.mode = mode;
  }
  if (!partial || has('severity')) {
    const sev = b.severity === undefined ? 'normal' : b.severity;
    if (!SEVERITIES.includes(sev)) return { error: '重要程度不正確（normal／important）' };
    f.severity = sev;
  }
  if (!partial || has('roles')) {
    let roles = b.roles === undefined || b.roles === null ? [] : b.roles;
    if (!Array.isArray(roles)) return { error: '目標角色格式不正確' };
    roles = [...new Set(roles.map(r => String(r)))];
    const bad = roles.find(r => !knownRoles.includes(r));
    if (bad) return { error: `未知的角色代碼：${String(bad).slice(0, 40)}` };
    // 全選＝全部（空陣列），之後新增角色也會自動涵蓋
    if (roles.length >= knownRoles.length) roles = [];
    f.roles = roles;
  }
  if (!partial || has('enabled')) {
    f.enabled = b.enabled === undefined ? true : !!b.enabled;
  }
  if (!partial || has('priority')) {
    const p = b.priority === undefined || b.priority === '' ? 100 : parseInt(b.priority, 10);
    if (!Number.isFinite(p) || p < 0 || p > 9999) return { error: '顯示順序需為 0–9999 的整數' };
    f.priority = p;
  }
  if (!partial || has('startAt')) {
    const r = parseTaipeiInput(b.startAt);
    if (!r.ok) return { error: '開始時間格式不正確' };
    f.startAt = r.value;
  }
  if (!partial || has('endAt')) {
    const r = parseTaipeiInput(b.endAt);
    if (!r.ok) return { error: '結束時間格式不正確' };
    f.endAt = r.value;
  }
  return { fields: f };
}

/** 開始／結束時間互相檢查（合併後的最終值） */
function checkWindow(a) {
  if (a.startAt && a.endAt && Date.parse(a.endAt) <= Date.parse(a.startAt)) {
    return '結束時間必須晚於開始時間';
  }
  return null;
}

function targetsRole(a, role) {
  return !Array.isArray(a.roles) || a.roles.length === 0 || a.roles.includes(role);
}

function inWindow(a, nowMs) {
  if (a.startAt && Date.parse(a.startAt) > nowMs) return false;
  if (a.endAt && Date.parse(a.endAt) <= nowMs) return false;
  return true;
}

function isLive(a, nowMs) {
  return a.enabled !== false && inWindow(a, nowMs);
}

/** 這位使用者「現在應看到」的公告（once 已確認的排除；everyLogin 的每登入一次由前端控制） */
function activeFor(list, user, nowMs) {
  const out = (list || []).filter(a => {
    if (!isLive(a, nowMs)) return false;
    if (!targetsRole(a, user.role)) return false;
    if (a.mode === 'once') {
      const r = a.readBy && a.readBy[user.username];
      if (r && r.version >= a.version) return false;
    }
    return true;
  });
  out.sort((x, y) => (x.priority - y.priority) || String(x.createdAt).localeCompare(String(y.createdAt)));
  // 只回前端需要的欄位：絕不含 readBy
  return out.map(a => ({
    id: a.id, version: a.version, title: a.title, body: a.body,
    mode: a.mode, severity: a.severity,
  }));
}

/** 統計：已讀（目前版本）/ 目標人數（啟用中帳號，排除客戶池） */
function statsFor(a, users) {
  const targets = users.filter(u => u.active !== false && u.role !== 'pool' && targetsRole(a, u.role));
  let read = 0;
  for (const u of targets) {
    const r = a.readBy && a.readBy[u.username];
    if (r && r.version >= a.version) read++;
  }
  return { readCount: read, targetCount: targets.length };
}

function status(a, nowMs) {
  if (a.enabled === false) return 'disabled';
  if (a.startAt && Date.parse(a.startAt) > nowMs) return 'scheduled';
  if (a.endAt && Date.parse(a.endAt) <= nowMs) return 'expired';
  return 'live';
}

/**
 * 記錄「這位使用者這則公告何時被送到前端」（GET /active 時呼叫）。回傳是否有變動（需存檔）。
 * 距上次記錄不足 READ_MIN_MS 且同版本就不更新：重複呼叫不會把倒數起點往後推，也減少寫入。
 */
function markShown(a, username, nowMs) {
  if (!a.shown || typeof a.shown !== 'object') a.shown = {};
  const prev = a.shown[username];
  if (prev && prev.version === a.version && nowMs - prev.at < READ_MIN_MS) return false;
  a.shown[username] = { version: a.version, at: nowMs };
  return true;
}
/** ack 是否太早：從未被送達（沒 shownAt）或距 shownAt 不足 READ_MIN_MS */
function ackTooEarly(a, username, nowMs) {
  const s = a.shown && a.shown[username];
  return !s || (nowMs - s.at) < READ_MIN_MS;
}

/** 管理員列表用的一筆（不含 readBy／shown 內容） */
function adminView(a, users, nowMs) {
  const { readBy, shown, ...rest } = a; // eslint-disable-line no-unused-vars
  return {
    ...rest,
    startLocal: toTaipeiLocal(a.startAt),
    endLocal: toTaipeiLocal(a.endAt),
    status: status(a, nowMs),
    ...statsFor(a, users),
  };
}

module.exports = {
  READ_MIN_MS, markShown, ackTooEarly,
  MAX_ANNOUNCEMENTS, MAX_TITLE, MAX_BODY, MODES, SEVERITIES,
  cleanText, parseTaipeiInput, toTaipeiLocal, validateInput, checkWindow,
  targetsRole, inWindow, isLive, activeFor, statsFor, status, adminView,
};
