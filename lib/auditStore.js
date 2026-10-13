'use strict';
/**
 * 稽核事件儲存（只增不改）。
 *
 * 目標：稽核日誌「完整、不被覆蓋、不被截斷」。
 *  - Postgres：表 audit_events，每筆事件單列 INSERT（不做讀-改-寫、不依賴快取）；應用程式碼永不對此表 UPDATE／DELETE／TRUNCATE。
 *  - 本機 JSON 模式：append-only JSON Lines（audit.log.jsonl），無筆數上限。
 *  - 舊資料（app_data 的 audit_log 整包 JSONB、audit.log.json、contact_audit.json）保留為「唯讀歷史」，讀取時與新資料依時間合併。
 *
 * 寫入失敗：重試一次；仍失敗 → console.error 'AUDIT_WRITE_FAILED' + 寫本機備援檔 + 計數（不讓請求失敗）。
 * 本模組不含任何連線資訊；pg 以注入的 query(sql, params) 函式存取，方便以假 client 測試。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');

const MAX_DETAIL = 20000;
const MAX_SHORT = 500;
const FAIL_MARK = 'AUDIT_WRITE_FAILED';

const TABLE_SQL = `
CREATE TABLE IF NOT EXISTS audit_events (
  id        BIGSERIAL PRIMARY KEY,
  ts        TIMESTAMPTZ NOT NULL DEFAULT now(),
  kind      TEXT NOT NULL,
  action    TEXT NOT NULL DEFAULT '',
  operator  TEXT NOT NULL DEFAULT '',
  target    TEXT NOT NULL DEFAULT '',
  detail    TEXT NOT NULL DEFAULT '',
  ip        TEXT NOT NULL DEFAULT '',
  ua        TEXT NOT NULL DEFAULT '',
  extra     JSONB NOT NULL DEFAULT '{}'::jsonb
)`;
const INDEX_SQL = [
  `CREATE INDEX IF NOT EXISTS audit_events_kind_ts_idx ON audit_events (kind, ts DESC)`,
  `CREATE INDEX IF NOT EXISTS audit_events_action_idx ON audit_events (action)`,
  `CREATE INDEX IF NOT EXISTS audit_events_operator_idx ON audit_events (operator)`,
];
const INSERT_SQL = `INSERT INTO audit_events (ts, kind, action, operator, target, detail, ip, ua, extra) VALUES ($1::timestamptz, $2, $3, $4, $5, $6, $7, $8, $9::jsonb)`;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function clean(s, max) {
  s = s == null ? '' : String(s);
  s = s.replace(/\u0000/g, '');           // Postgres text／jsonb 不接受 NUL
  return s.length > max ? s.slice(0, max) : s;
}
function cleanJson(o) {
  try { return JSON.stringify(o == null ? {} : o).replace(/\\u0000/g, ''); } catch (e) { return '{}'; }
}

/** 事件 → 內部列（ts 為 ISO 字串） */
function normalizeEvent(ev) {
  ev = ev || {};
  let ts = ev.ts || ev.timestamp;
  if (!ts || isNaN(Date.parse(ts))) ts = new Date().toISOString();
  return {
    id: ev.id || null,
    ts: new Date(ts).toISOString(),
    kind: clean(ev.kind || 'log', 40),
    action: clean(ev.action, MAX_SHORT),
    operator: clean(ev.operator, MAX_SHORT),
    target: clean(ev.target, MAX_SHORT * 2),
    detail: clean(ev.detail, MAX_DETAIL),
    ip: clean(ev.ip, 100),
    ua: clean(ev.ua != null ? ev.ua : ev.userAgent, 500),
    extra: ev.extra && typeof ev.extra === 'object' ? ev.extra : {},
  };
}

// ───────── 篩選 ─────────
const DATE_ONLY = /^\d{4}-\d{2}-\d{2}$/;
/** 'YYYY-MM-DD' 視為台北日（UTC+8）；其他格式交給 Date.parse；無法解析 → null */
function parseBound(s, isEnd) {
  if (s == null || s === '') return null;
  s = String(s).trim();
  if (DATE_ONLY.test(s)) {
    const t = Date.parse(s + 'T00:00:00+08:00');
    if (isNaN(t)) return null;
    return isEnd ? t + 86400000 : t;          // to 為「含當日」→ 轉成隔日 00:00 的排他上界
  }
  const t = Date.parse(s);
  return isNaN(t) ? null : t;
}
function normFilters(f) {
  f = f || {};
  const kind = f.kind ? String(f.kind) : 'log';
  return {
    kind,
    action: f.action ? clean(f.action, 200) : '',
    operator: f.operator ? clean(f.operator, 200) : '',
    q: f.q ? clean(f.q, 200).trim() : '',
    fromMs: parseBound(f.from, false),
    toMs: parseBound(f.to, true),
  };
}
function rowMatches(r, f) {
  if (f.kind !== 'all' && r.kind !== f.kind) return false;
  if (f.action && r.action !== f.action) return false;
  if (f.operator && r.operator !== f.operator) return false;
  const t = Date.parse(r.ts);
  if (f.fromMs != null && !(t >= f.fromMs)) return false;
  if (f.toMs != null && !(t < f.toMs)) return false;
  if (f.q) {
    const q = f.q.toLowerCase();
    const hay = [r.action, r.operator, r.target, r.detail].join('\n').toLowerCase();
    if (!hay.includes(q)) return false;
  }
  return true;
}
function likeEscape(s) { return s.replace(/[\\%_]/g, (m) => '\\' + m); }

// ───────── 舊資料正規化（唯讀歷史）─────────
function legacyLogToRow(l) {
  return { id: l.id || null, ts: new Date(l.timestamp && !isNaN(Date.parse(l.timestamp)) ? l.timestamp : 0).toISOString(), kind: 'log', action: l.action || '', operator: l.operator || '', target: l.target || '', detail: l.detail || '', ip: l.ip || '', ua: l.userAgent || '', extra: {}, legacy: true };
}
function legacyContactToRow(l) {
  const base = { userName: l.userName || '', targetId: l.targetId || '', targetName: l.targetName || '', targetCompany: l.targetCompany || '', changes: Array.isArray(l.changes) ? l.changes : [] };
  return { id: l.id || null, ts: new Date(l.timestamp && !isNaN(Date.parse(l.timestamp)) ? l.timestamp : 0).toISOString(), kind: 'contact', action: l.action || '', operator: l.userId || '', target: l.targetId || '', detail: contactDetail(base), ip: '', ua: '', extra: base, legacy: true };
}
function contactDetail(x) { return [x.targetName, x.targetCompany].filter(Boolean).join(' / '); }

// ───────── API 輸出形狀 ─────────
function toApiLog(r) {
  return { id: r.id, action: r.action, operator: r.operator, target: r.target, detail: r.detail, ip: r.ip, userAgent: r.ua, timestamp: r.ts };
}
function toApiContact(r) {
  const x = r.extra || {};
  return { id: r.id, timestamp: r.ts, userId: r.operator, userName: x.userName || r.operator, action: r.action, targetId: x.targetId || r.target, targetName: x.targetName || '', targetCompany: x.targetCompany || '', changes: Array.isArray(x.changes) ? x.changes : [] };
}
function toApi(r) { return r.kind === 'contact' ? toApiContact(r) : toApiLog(r); }

// ───────── 轉接器：Postgres ─────────
function pgAdapter(opts) {
  const query = opts.query;
  let ready = null;
  let legacyCache = null, legacyAt = 0;
  const LEGACY_TTL = opts.legacyTtlMs != null ? opts.legacyTtlMs : 30000;

  function ensure() {
    if (!ready) {
      ready = (async () => {
        await query(TABLE_SQL);
        for (const s of INDEX_SQL) await query(s);
      })().catch((e) => { ready = null; throw e; });
    }
    return ready;
  }
  function where(f, params, extra) {
    const c = [];
    const add = (sql, v) => { params.push(v); c.push(sql.replace('$#', '$' + params.length)); };
    if (f.kind !== 'all') add('kind = $#', f.kind);
    if (f.action) add('action = $#', f.action);
    if (f.operator) add('operator = $#', f.operator);
    if (f.fromMs != null) add('ts >= $#::timestamptz', new Date(f.fromMs).toISOString());
    if (f.toMs != null) add('ts < $#::timestamptz', new Date(f.toMs).toISOString());
    if (f.q) { params.push('%' + likeEscape(f.q) + '%'); const n = '$' + params.length; c.push(`(action ILIKE ${n} OR operator ILIKE ${n} OR target ILIKE ${n} OR detail ILIKE ${n})`); }
    if (extra) extra(add);
    return c.length ? ' WHERE ' + c.join(' AND ') : '';
  }
  const rowOf = (r) => ({ id: 'ae-' + r.id, seq: Number(r.id), ts: new Date(r.ts).toISOString(), kind: r.kind, action: r.action, operator: r.operator, target: r.target, detail: r.detail, ip: r.ip, ua: r.ua, extra: r.extra || {} });

  return {
    kind: 'pg',
    ensure,
    reset() { ready = null; },
    /** 啟動自我檢查：建表後做一次無害查詢；失敗代表這個資料庫帳號用不了 audit_events */
    async ping() { await ensure(); await query('SELECT 1 FROM audit_events LIMIT 1'); },
    async insert(row) {
      await ensure();
      await query(INSERT_SQL, [row.ts, row.kind, row.action, row.operator, row.target, row.detail, row.ip, row.ua, cleanJson(row.extra)]);
    },
    async count(f) {
      await ensure();
      const params = [];
      const r = await query('SELECT count(*)::int AS n FROM audit_events' + where(f, params), params);
      return Number(r.rows[0].n);
    },
    async *iterate(f, chunk) {
      await ensure();
      chunk = chunk || 500;
      let last = null;
      for (;;) {
        const params = [];
        const w = where(f, params, last == null ? null : (add) => add('id < $#', last));
        params.push(chunk);
        const r = await query(`SELECT id, ts, kind, action, operator, target, detail, ip, ua, extra FROM audit_events${w} ORDER BY id DESC LIMIT $${params.length}`, params);
        for (const x of r.rows) yield rowOf(x);
        if (r.rows.length < chunk) return;
        last = r.rows[r.rows.length - 1].id;
      }
    },
    /** 舊的整包 JSONB（app_data.id='audit_log'）：唯讀歷史 */
    async legacy() {
      if (legacyCache && Date.now() - legacyAt < LEGACY_TTL) return legacyCache;
      let rows = [];
      try {
        const r = await query(`SELECT content FROM app_data WHERE id = 'audit_log'`);
        if (r.rows.length) {
          const c = r.rows[0].content;
          const arr = Array.isArray(c) ? c : (c && Array.isArray(c.logs) ? c.logs : []);
          rows = arr.map(legacyLogToRow);
        }
      } catch (e) { console.error('[auditStore] legacy blob read failed:', e && e.message); }
      legacyCache = rows; legacyAt = Date.now();
      return rows;
    },
  };
}

// ───────── 轉接器：本機 JSON Lines ─────────
function jsonlAdapter(opts) {
  const file = opts.file;
  const legacyLogFile = opts.legacyLogFile;
  const legacyContactFile = opts.legacyContactFile;
  function readAll() {
    let txt = '';
    try { txt = fs.readFileSync(file, 'utf8'); } catch (e) { return []; }
    const out = [];
    let seq = 0;
    for (const line of txt.split('\n')) {
      if (!line.trim()) continue;
      seq++;
      try { const o = JSON.parse(line); o.seq = seq; out.push(o); } catch (e) { /* 壞行略過（保留在檔案內，不改寫） */ }
    }
    return out;
  }
  return {
    kind: 'jsonl',
    ensure: async () => {},
    reset() {},
    async insert(row) {
      const line = JSON.stringify({ id: row.id || crypto.randomUUID(), ts: row.ts, kind: row.kind, action: row.action, operator: row.operator, target: row.target, detail: row.detail, ip: row.ip, ua: row.ua, extra: row.extra }) + '\n';
      fs.appendFileSync(file, line, 'utf8');
    },
    async count(f) { return readAll().filter((r) => rowMatches(r, f)).length; },
    async *iterate(f) {
      const all = readAll();
      for (let i = all.length - 1; i >= 0; i--) if (rowMatches(all[i], f)) yield all[i];
    },
    async legacy() {
      const rows = [];
      if (legacyLogFile) { try { const a = JSON.parse(fs.readFileSync(legacyLogFile, 'utf8')); if (Array.isArray(a)) a.forEach((l) => rows.push(legacyLogToRow(l))); } catch (e) { /* 無舊檔 */ } }
      if (legacyContactFile) { try { const a = JSON.parse(fs.readFileSync(legacyContactFile, 'utf8')); if (Array.isArray(a)) a.forEach((l) => rows.push(legacyContactToRow(l))); } catch (e) { /* 無舊檔 */ } }
      return rows;
    },
  };
}

// ───────── 轉接器：只讀舊資料（AUDIT_EVENTS_TABLE 關閉時的讀取）─────────
// 不碰 audit_events、也不寫任何東西：只把舊的整包 JSONB／audit.log.json／contact_audit.json 當成唯讀來源，讓分頁、篩選、匯出在記憶體上照樣運作。
function legacyReadAdapter(opts) {
  return {
    kind: 'legacy',
    ensure: async () => {},
    reset() {},
    async insert() { throw new Error('legacy 唯讀轉接器不接受寫入'); },
    async count() { return 0; },
    async *iterate() { /* 沒有新表 */ },
    async legacy() {
      const rows = [];
      try { const a = await opts.readLogs(); if (Array.isArray(a)) a.forEach((l) => rows.push(legacyLogToRow(l))); } catch (e) { console.error('[auditStore] legacy log read failed:', e && e.message); }
      if (opts.legacyContactFile) { try { const a = JSON.parse(fs.readFileSync(opts.legacyContactFile, 'utf8')); if (Array.isArray(a)) a.forEach((l) => rows.push(legacyContactToRow(l))); } catch (e) { /* 無舊檔 */ } }
      return rows;
    },
  };
}

// ───────── 儲存體 ─────────
function createAuditStore(adapter, opts) {
  opts = opts || {};
  const retryDelayMs = opts.retryDelayMs != null ? opts.retryDelayMs : 60;
  const fallbackFile = opts.fallbackFile || path.join(os.tmpdir(), 'itts-audit-write-failed.jsonl');
  const st = { inserted: 0, failed: 0, retried: 0, lastError: '', lastFailedAt: null };
  const pending = new Set();

  async function doInsert(ev) {
    const row = normalizeEvent(ev);
    let lastErr = null;
    for (let attempt = 0; attempt < 2; attempt++) {
      try { await adapter.insert(row); st.inserted++; return true; }
      catch (e) {
        lastErr = e;
        if (attempt === 0) { st.retried++; if (adapter.reset) adapter.reset(); await sleep(retryDelayMs); }
      }
    }
    st.failed++; st.lastError = String(lastErr && lastErr.message || lastErr).slice(0, 300); st.lastFailedAt = new Date().toISOString();
    try { console.error(FAIL_MARK, JSON.stringify({ at: st.lastFailedAt, error: st.lastError, event: row })); } catch (e) { console.error(FAIL_MARK, st.lastError); }
    try { fs.appendFileSync(fallbackFile, JSON.stringify({ mark: FAIL_MARK, at: st.lastFailedAt, error: st.lastError, event: row }) + '\n', 'utf8'); } catch (e) { /* 唯讀環境：console 已留 */ }
    return false;
  }
  function insert(ev) {
    const p = doInsert(ev).catch((e) => { st.failed++; console.error(FAIL_MARK, e && e.message); return false; });
    pending.add(p);
    p.then(() => pending.delete(p), () => pending.delete(p));
    return p;
  }
  async function flush() { while (pending.size) await Promise.all([...pending]); }

  /** 依時間新→舊合併「新表」與「舊唯讀歷史」 */
  async function* merged(f) {
    const legacy = (await adapter.legacy()).filter((r) => rowMatches(r, f)).sort((a, b) => Date.parse(b.ts) - Date.parse(a.ts));
    const it = adapter.iterate(f);
    let li = 0;
    let nx = await it.next();
    for (;;) {
      const t = nx.done ? null : nx.value;
      const l = li < legacy.length ? legacy[li] : null;
      if (!t && !l) return;
      if (t && (!l || Date.parse(t.ts) >= Date.parse(l.ts))) { yield t; nx = await it.next(); }
      else { yield l; li++; }
    }
  }
  async function total(f) {
    const legacy = (await adapter.legacy()).filter((r) => rowMatches(r, f)).length;
    return (await adapter.count(f)) + legacy;
  }
  async function list(filters, page) {
    const f = normFilters(filters);
    const cap = (page && page.cap) || 1000;
    const limit = Math.max(1, Math.min(cap, parseInt(page && page.limit, 10) || 200));
    const offset = Math.max(0, Math.min(100000, parseInt(page && page.offset, 10) || 0));
    const rows = [];
    let i = 0;
    for await (const r of merged(f)) {
      if (i++ < offset) continue;
      rows.push(r);
      if (rows.length >= limit) break;
    }
    return { total: await total(f), limit, offset, rows };
  }

  return {
    insert, flush, list, total: (filters) => total(normFilters(filters)),
    iterate: (filters) => merged(normFilters(filters)),
    stats: () => Object.assign({}, st),
    ensure: () => adapter.ensure(),
    adapter,
  };
}

// ───────── CSV ─────────
/** CSV 儲存格：公式注入防護（= + - @ Tab CR 開頭加單引號）＋ RFC4180 引號 */
function csvCell(v) {
  let s = v == null ? '' : String(v);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  if (/[",\r\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
  return s;
}
const CSV_HEAD = ['時間', '類型', '動作', '操作者', '對象', '說明', 'IP', '裝置', '附加資料'];
function csvLine(r) {
  const extra = r.extra && Object.keys(r.extra).length ? JSON.stringify(r.extra) : '';
  return [r.ts, r.kind, r.action, r.operator, r.target, r.detail, r.ip, r.ua, extra].map(csvCell).join(',') + '\r\n';
}

/**
 * 稽核門面（安全開關）：enabled=false（預設）→ 讀寫都走舊路徑（legacyWrite／legacyStore），完全不碰 audit_events；
 * enabled=true → 走新表，但先做一次自我檢查（ensure＋SELECT 1），失敗就記 AUDIT_TABLE_UNAVAILABLE 並在本行程內自動退回舊路徑；
 * 新表單筆寫入失敗（重試後）時，事件同時寫進舊路徑，不丟。永不 reject、永不讓請求失敗。
 */
function createAuditFacade(o) {
  let mode = o.enabled ? 'table' : 'legacy';
  let fellBack = false, unavailable = '', initP = null, legacyWrites = 0;
  const pending = new Set();
  function init() {
    if (!o.enabled) return Promise.resolve(mode);
    if (!initP) {
      initP = (async () => {
        try { if (o.tableStore.adapter.ping) await o.tableStore.adapter.ping(); else await o.tableStore.ensure(); }
        catch (e) {
          mode = 'legacy'; fellBack = true; unavailable = String((e && e.message) || e).slice(0, 300);
          console.error('AUDIT_TABLE_UNAVAILABLE', unavailable, '— 稽核寫入自動退回舊路徑（整包 JSONB／audit.log.json），本行程內不再使用 audit_events');
        }
        return mode;
      })();
    }
    return initP;
  }
  async function legacyInsert(ev) {
    try { await o.legacyWrite(ev); legacyWrites++; return true; }
    catch (e) { console.error('AUDIT_WRITE_FAILED', 'legacy path:', e && e.message); return false; }
  }
  function insert(ev) {
    const p = (async () => {
      if (o.enabled) await init();
      if (mode === 'table') {
        const ok = await o.tableStore.insert(ev);
        if (!ok) await legacyInsert(ev);   // 已記 AUDIT_WRITE_FAILED；同時留一份在舊路徑
        return ok;
      }
      return legacyInsert(ev);
    })().catch((e) => { console.error('AUDIT_WRITE_FAILED', e && e.message); return false; });
    pending.add(p);
    p.then(() => pending.delete(p), () => pending.delete(p));
    return p;
  }
  const active = async () => { if (o.enabled) await init(); return mode === 'table' ? o.tableStore : o.legacyStore; };
  return {
    init, insert,
    mode: () => mode,
    async flush() { while (pending.size) await Promise.all([...pending]); await o.tableStore.flush(); },
    async list(f, page) { return (await active()).list(f, page); },
    async total(f) { return (await active()).total(f); },
    async *iterate(f) { yield* (await active()).iterate(f); },
    stats() {
      const t = o.tableStore.stats();
      return Object.assign({}, t, { inserted: t.inserted + legacyWrites, legacyWrites, auditStore: mode, fellBack, unavailable });
    },
    tableStore: o.tableStore, legacyStore: o.legacyStore,
  };
}

module.exports = { createAuditFacade, legacyReadAdapter, createAuditStore, pgAdapter, jsonlAdapter, normalizeEvent, normFilters, rowMatches, toApi, toApiLog, toApiContact, csvCell, csvLine, CSV_HEAD, contactDetail, TABLE_SQL, INSERT_SQL, FAIL_MARK, parseBound };
