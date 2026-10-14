'use strict';
/**
 * 稽核日誌防竄改（hash chain）＋外部錨點（anchor）。第三波第 3 項。預設關閉（AUDIT_CHAIN 沒設成 '1' → 行為與以前位元級相同）。
 *
 * ═══ 操作手冊（給管理員／IT）═══
 *
 * 【這是什麼】
 *  每一筆新的稽核事件（audit_events 表／本機 audit.log.jsonl）多帶四個欄位：
 *    seq        流水序號（1,2,3… 無缺號）
 *    prev_hash  前一筆的 row_hash（第一筆是 64 個 0）
 *    row_hash   HMAC-SHA256(鏈金鑰, prev_hash + "\n" + canonicalJSON({kind,action,operator,target,detail,ip,ua,ts,extra,seq,kid}))
 *    kid        金鑰代號（換金鑰用，見下方「換金鑰」）
 *  鏈金鑰只存在伺服器環境變數，**絕不存進資料庫、絕不進 repo、絕不出現在任何回應或日誌**。
 *  所以只拿到資料庫的人（含能 UPDATE/DELETE 的人）改了任何一筆，驗證就會在那一筆斷掉；也無法重算出一條「看起來合法」的新鏈。
 *
 * 【啟用步驟】（先在 Test/Demo 專案跑過 scripts/verify-audit-pg.js --chain 再上正式站）
 *  1. 在 Vercel 專案的環境變數設定（只放 Vercel，不要貼進對話、檔案、repo）：
 *       AUDIT_CHAIN_SECRET = <很長的隨機字串，至少 32 字元，例如 openssl rand -base64 48>
 *       AUDIT_EVENTS_TABLE = 1        （稽核事件新表；鏈只在新表模式有意義）
 *       AUDIT_CHAIN        = 1
 *       CRON_SECRET        = <隨機字串>（每日錨點排程用；與寄信排程共用同一個）
 *       （選用）AUDIT_ANCHOR_WEBHOOK_URL = https://…（把每日錨點自動推到外部；只接受 https:443，會擋內網位址）
 *  2. 重新部署。第一筆事件寫入時自動補欄位（ALTER TABLE ADD COLUMN IF NOT EXISTS）與唯一索引；舊資料不動。
 *  3. 後台「版本資訊」會多一張「稽核防竄改」卡：看鏈是否啟用、鏈頭序號、上次驗證結果、最近錨點；按「立即驗證」即時驗證。
 *  4. 在 vercel.json 已有每日 cron：/api/cron/audit-anchor（需 CRON_SECRET）會產生錨點、保留最近 90 份在主資料
 *     （data.auditAnchors）並（若設了 webhook）推送出去。**另外**請管理員定期從後台「下載錨點」存一份到離線／另一套系統——
 *     錨點含簽章，日後拿來證明「日誌沒有被截斷或回溯」。
 *
 * 【停用／回復】把 AUDIT_CHAIN 拿掉（或設成 0）再部署：新事件不再上鏈，已在表內的資料與欄位原封不動；
 *   之後再開啟，鏈接在最後一筆之後；中間關閉期間寫入的事件是「未納入鏈」（驗證報告會列為 unchainedAfterStart，屬警示不算斷鏈）。
 *   沒有金鑰（AUDIT_CHAIN_SECRET、JWT_SECRET、SESSION_SECRET 都沒有）時即使 AUDIT_CHAIN=1 也不會啟用，記 AUDIT_CHAIN_DISABLED_NO_KEY。
 *
 * 【換金鑰（rotate）】
 *   a. 把舊值放進 AUDIT_CHAIN_SECRETS_JSON，例如 {"k1":"<舊 secret>"}（kid → secret，僅供驗證舊的列）。
 *   b. AUDIT_CHAIN_KID 設新代號（例如 k2），AUDIT_CHAIN_SECRET 換成新的 secret。
 *   c. 重新部署。序號與 prev_hash 接續舊鏈（不重開新鏈、不會出現缺口），新列的 kid=k2、用新金鑰簽；舊列仍用 k1 驗。
 *   d. 舊 secret 要留著（放 AUDIT_CHAIN_SECRETS_JSON）直到舊列不需再驗證為止；遺失舊 secret → 那些列驗證時回報 UNKNOWN_KID。
 *   金鑰來源優先序：AUDIT_CHAIN_SECRET > JWT_SECRET > SESSION_SECRET（後兩者只是備援，會被登入機制輪替時一併失效，不建議）。
 *
 * 【效能】Postgres 每筆事件多一個交易：BEGIN → pg_advisory_xact_lock(固定鍵) → 讀鏈頭 → INSERT → COMMIT。
 *   全站稽核寫入因此被序列化（一次一筆，約 2–4 次往返；Supabase 在東京區約每筆 20–60ms）。這個系統的稽核量（人工操作為主）遠低於此上限；
 *   若鎖等待超過 4 秒（lock_timeout）→ 該筆事件改為「不上鏈」直接寫入，記 AUDIT_CHAIN_BREAK_RISK，**絕不遺失事件**。
 *
 * 【保護範圍與限制】見 scripts/check-audit-chain.js 檔頭與交付報告：
 *   能抓：改任一欄位、刪中間列、插入偽造列、換順序、換 kid、用錯金鑰；有錨點時能抓「截尾」與「整段回溯」。
 *   抓不到：同時拿到資料庫與 Vercel 環境變數金鑰的人（可重算整條鏈）；沒有錨點時的「刪尾端幾筆」；鏈啟用之前的舊資料；
 *     在鏈啟用後、不帶鏈欄位直接 INSERT 的偽造列（只會被計入 unchainedAfterStart 警示）；本機 JSONL 模式只保證單一行程。
 *
 * 本模組不含任何連線資訊、不 require 專案其他模組；Postgres 交易由 lib/auditStore.js 的 pgAdapter 呼叫 seal()。
 */
const crypto = require('crypto');
const dns = require('dns');
const https = require('https');
const net = require('net');

const GENESIS = '0'.repeat(64);
const KEY_LABEL = 'itts-crm/audit-chain/v1';      // 金鑰衍生的網域分隔標籤（和 JWT 簽章用途隔開）
const LOCK_KEY = '7340132059';                    // pg_advisory_xact_lock 的固定 bigint（以字串傳入）
const MIN_SECRET_LEN = 16;
const KID_RE = /^[A-Za-z0-9_.-]{1,24}$/;
const HEX64 = /^[0-9a-f]{64}$/;
const CHAIN_MARK = { risk: 'AUDIT_CHAIN_BREAK_RISK', noKey: 'AUDIT_CHAIN_DISABLED_NO_KEY' };

// ───────── 正規化 JSON（鍵排序、略過 undefined）─────────
// Postgres jsonb 會重排鍵與空白，所以雜湊用「鍵排序」的字串，與儲存形式無關。
function canonical(v) {
  if (v === null || v === undefined) return 'null';
  const t = typeof v;
  if (t === 'string') return JSON.stringify(v);
  if (t === 'number') return Number.isFinite(v) ? JSON.stringify(v) : 'null';
  if (t === 'boolean') return v ? 'true' : 'false';
  if (Array.isArray(v)) return '[' + v.map((x) => canonical(x)).join(',') + ']';
  if (t === 'object') {
    return '{' + Object.keys(v).filter((k) => v[k] !== undefined).sort().map((k) => JSON.stringify(k) + ':' + canonical(v[k])).join(',') + '}';
  }
  return 'null';
}

function deriveKey(secret) { return crypto.createHmac('sha256', String(secret)).update(KEY_LABEL).digest(); }
function hmacHex(key, text) { return crypto.createHmac('sha256', key).update(text, 'utf8').digest('hex'); }
function safeEq(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  const x = crypto.createHash('sha256').update(a).digest(), y = crypto.createHash('sha256').update(b).digest();
  return crypto.timingSafeEqual(x, y);
}

// ───────── 設定（金鑰只在記憶體）─────────
const _cfgCache = new Map();
function fingerprint(env) {
  return crypto.createHash('sha256').update(JSON.stringify([env.AUDIT_CHAIN, env.AUDIT_CHAIN_SECRET, env.JWT_SECRET, env.SESSION_SECRET, env.AUDIT_CHAIN_KID, env.AUDIT_CHAIN_SECRETS_JSON])).digest('hex');
}
/**
 * 讀環境變數 → { flag, enabled, canVerify, kid, keyFor(kid), kids, secretSource, disabledReason }
 *  enabled   ＝ 新事件要不要上鏈（AUDIT_CHAIN=1 且有可用金鑰）
 *  canVerify ＝ 有任何金鑰可驗證（即使 AUDIT_CHAIN 已關，仍可驗證既有的鏈）
 * 結果以環境變數指紋記憶化（衍生金鑰要做 HMAC；不重複算也不重複印警告）。金鑰本身不會出現在回傳的任何可序列化欄位。
 */
function getConfig(env) {
  env = env || process.env;
  const fp = fingerprint(env);
  const hit = _cfgCache.get(fp);
  if (hit) return hit;
  const flag = env.AUDIT_CHAIN === '1';
  const str = (v) => (typeof v === 'string' ? v : '');
  let secret = '', secretSource = null, disabledReason = '';
  const explicit = str(env.AUDIT_CHAIN_SECRET);
  if (explicit) {
    if (explicit.length >= MIN_SECRET_LEN) { secret = explicit; secretSource = 'AUDIT_CHAIN_SECRET'; }
    else disabledReason = 'KEY_TOO_SHORT';       // 設了但太短：不悄悄改用備援金鑰，直接停用
  } else if (str(env.JWT_SECRET).length >= MIN_SECRET_LEN) { secret = env.JWT_SECRET; secretSource = 'JWT_SECRET'; }
  else if (str(env.SESSION_SECRET).length >= MIN_SECRET_LEN) { secret = env.SESSION_SECRET; secretSource = 'SESSION_SECRET'; }
  const kidRaw = str(env.AUDIT_CHAIN_KID);
  const kid = KID_RE.test(kidRaw) ? kidRaw : 'k1';
  const keys = new Map();
  if (env.AUDIT_CHAIN_SECRETS_JSON) {
    try {
      const o = JSON.parse(env.AUDIT_CHAIN_SECRETS_JSON);
      if (o && typeof o === 'object' && !Array.isArray(o)) {
        for (const k of Object.keys(o)) if (KID_RE.test(k) && typeof o[k] === 'string' && o[k].length >= MIN_SECRET_LEN) keys.set(k, deriveKey(o[k]));
      }
    } catch (e) { console.error('AUDIT_CHAIN_SECRETS_JSON 無法解析（已忽略）'); }
  }
  if (secret) keys.set(kid, deriveKey(secret));
  const enabled = flag && !!secret;
  if (flag && !secret) {
    if (!disabledReason) disabledReason = 'NO_KEY';
    console.error(CHAIN_MARK.noKey, '— AUDIT_CHAIN=1 但沒有可用金鑰（' + disabledReason + '）；稽核事件維持不上鏈。請設定 AUDIT_CHAIN_SECRET（至少 ' + MIN_SECRET_LEN + ' 字元）');
  } else if (!flag) disabledReason = 'OFF';
  const cfg = {
    flag, enabled, canVerify: keys.size > 0, kid, secretSource, disabledReason: enabled ? '' : disabledReason,
    kids: Array.from(keys.keys()),
    keyFor(k) { return keys.get(k) || null; },
  };
  Object.defineProperty(cfg, 'keyFor', { enumerable: false });      // 保險：JSON.stringify(cfg) 不會碰到金鑰
  if (_cfgCache.size > 8) _cfgCache.clear();
  _cfgCache.set(fp, cfg);
  return cfg;
}
function resetConfigCache() { _cfgCache.clear(); }

// ───────── 事件列雜湊 ─────────
function hashFields(r, seq, kid) {
  return { kind: r.kind, action: r.action, operator: r.operator, target: r.target, detail: r.detail, ip: r.ip, ua: r.ua, ts: r.ts, extra: r.extra == null ? {} : r.extra, seq: Number(seq), kid };
}
function computeRowHash(key, prevHash, fields) { return hmacHex(key, String(prevHash) + '\n' + canonical(fields)); }
/**
 * 封存一筆：prev = 目前鏈頭 {seq,row_hash} 或 null；row = 已正規化的事件（extra 必須是「實際會被儲存」的物件）
 * 回傳 { seq, prev_hash, row_hash, kid }。
 */
function seal(cfg, prev, row) {
  const key = cfg.keyFor(cfg.kid);
  if (!key) throw new Error('CHAIN_KEY_MISSING');
  const seq = prev ? Number(prev.seq) + 1 : 1;
  const prev_hash = prev ? String(prev.row_hash) : GENESIS;
  if (!Number.isSafeInteger(seq) || seq < 1) throw new Error('CHAIN_SEQ_INVALID');
  if (prev && !HEX64.test(prev_hash)) throw new Error('CHAIN_HEAD_CORRUPT');
  return { seq, prev_hash, row_hash: computeRowHash(key, prev_hash, hashFields(row, seq, cfg.kid)), kid: cfg.kid };
}

// ───────── 驗證 ─────────
const toInt = (v) => { const n = parseInt(v, 10); return Number.isFinite(n) && n >= 0 ? n : null; };
/**
 * src（轉接器）需提供：chainRows({fromSeq,toSeq,chunk}) 非同步產生器（依 seq 遞增）、chainRowAt(seq)、chainHead()、chainUncovered()
 * opts: { from, to, anchor, chunk(≤1000), maxRows(預設 200000), budgetMs(預設 20000), now }
 * 回傳 { ok, checked, firstBreak:{seq,id,reason,detail}|null, legacyUncovered, unchainedAfterStart, headSeq, headHash, complete, nextFrom, anchor, ts, from, to }
 *  ok＝沒有任何斷鏈（complete=false 時代表「已驗證的前段沒問題」）。
 */
async function verifyChain(src, cfg, opts) {
  opts = opts || {};
  const chunk = Math.max(1, Math.min(1000, toInt(opts.chunk) || 1000));
  const maxRows = toInt(opts.maxRows) || 200000;
  const deadline = Date.now() + (toInt(opts.budgetMs) || 20000);
  const from = Math.max(1, toInt(opts.from) || 1);
  const toRaw = toInt(opts.to);
  const to = toRaw && toRaw >= from ? toRaw : null;
  const res = { ok: true, checked: 0, firstBreak: null, legacyUncovered: 0, unchainedAfterStart: 0, headSeq: 0, headHash: GENESIS, complete: true, nextFrom: null, anchor: null, ts: new Date().toISOString(), from, to };
  const brk = (seq, id, reason, detail) => { res.ok = false; res.firstBreak = { seq, id: id == null ? null : id, reason, detail: detail || '' }; };

  const head = await src.chainHead();
  if (head) { res.headSeq = Number(head.seq); res.headHash = head.row_hash; }
  try { const u = await src.chainUncovered(); res.legacyUncovered = u.legacyUncovered || 0; res.unchainedAfterStart = u.unchainedAfterStart || 0; } catch (e) { /* 統計失敗不影響驗證 */ }

  let prevHash = GENESIS, expected = from;
  if (from > 1) {
    const p = await src.chainRowAt(from - 1);
    if (!p) brk(from - 1, null, 'SEQ_GAP', '起點前一筆不存在（無法銜接）');
    else prevHash = p.row_hash;
  }
  if (res.ok) {
    for await (const r of src.chainRows({ fromSeq: from, toSeq: to, chunk })) {
      const seq = Number(r.seq);
      if (seq !== expected) { brk(expected, r.id, 'SEQ_GAP', `預期序號 ${expected}，實際 ${seq}（中間缺列、重複或順序被改）`); break; }
      if (r.prev_hash !== prevHash) { brk(seq, r.id, 'PREV_MISMATCH', 'prev_hash 與前一筆的 row_hash 不符'); break; }
      const key = cfg.keyFor(r.kid);
      if (!key) { brk(seq, r.id, 'UNKNOWN_KID', `找不到金鑰代號 ${String(r.kid).slice(0, 24)} 的金鑰（換金鑰後舊 secret 未保留？）`); break; }
      const calc = computeRowHash(key, r.prev_hash, hashFields(r, seq, r.kid));
      if (!safeEq(calc, r.row_hash)) { brk(seq, r.id, 'HASH_MISMATCH', '這一筆的內容與簽章不符（欄位被改過，或使用了錯誤的金鑰）'); break; }
      prevHash = r.row_hash; expected++; res.checked++;
      if (res.checked >= maxRows || Date.now() > deadline) { res.complete = false; res.nextFrom = expected; break; }
    }
  }

  if (opts.anchor) {
    const a = opts.anchor;
    const ar = { checked: true, ok: true, reason: '', seq: Number(a && a.seq) || 0 };
    if (!verifyAnchor(cfg, a)) { ar.ok = false; ar.reason = 'ANCHOR_HMAC_INVALID'; }
    else if (Number(a.seq) > 0) {
      const row = await src.chainRowAt(Number(a.seq));
      if (!row) { ar.ok = false; ar.reason = 'TRUNCATED'; }
      else if (!safeEq(String(row.row_hash), String(a.headHash))) { ar.ok = false; ar.reason = 'ANCHOR_MISMATCH'; }
    }
    if (!ar.ok) { res.ok = false; if (!res.firstBreak) res.firstBreak = { seq: ar.seq, id: null, reason: ar.reason, detail: ar.reason === 'TRUNCATED' ? '錨點記錄的那一筆已不在資料庫（尾端被截掉或整段被刪）' : ar.reason === 'ANCHOR_MISMATCH' ? '錨點那一筆的雜湊與當時記錄的不同（被改寫或回溯）' : '錨點簽章不正確（錨點本身被改過或金鑰不同）' }; }
    res.anchor = ar;
  }
  return res;
}

// ───────── 錨點 ─────────
function anchorText(seq, headHash, ts) { return 'anchor|' + seq + '|' + headHash + '|' + ts; }
function makeAnchor(cfg, head, count, now) {
  const key = cfg.keyFor(cfg.kid);
  if (!key) throw new Error('CHAIN_KEY_MISSING');
  const seq = head ? Number(head.seq) : 0;
  const headHash = head ? String(head.row_hash) : GENESIS;
  const ts = (now instanceof Date ? now : new Date(now || Date.now())).toISOString();
  return { v: 1, seq, headHash, count: Number(count) || 0, ts, kid: cfg.kid, anchorHmac: hmacHex(key, anchorText(seq, headHash, ts)) };
}
function verifyAnchor(cfg, a) {
  try {
    if (!a || typeof a !== 'object' || !Number.isSafeInteger(a.seq) || a.seq < 0 || !HEX64.test(String(a.headHash)) || typeof a.ts !== 'string' || !HEX64.test(String(a.anchorHmac))) return false;
    const key = cfg.keyFor(a.kid || cfg.kid);
    if (!key) return false;
    return safeEq(hmacHex(key, anchorText(a.seq, a.headHash, a.ts)), a.anchorHmac);
  } catch (e) { return false; }
}

// ───────── Cron 驗證（與 lib/mail/routes.js 相同規則：沒設 CRON_SECRET 一律拒絕；Bearer；等長雜湊後 timingSafeEqual）─────────
function cronAuthorized(req, env) {
  const secret = env && env.CRON_SECRET;
  if (typeof secret !== 'string' || secret === '') return false;
  const h = req && req.headers && req.headers.authorization;
  if (typeof h !== 'string') return false;
  const m = /^Bearer\s+(\S+)\s*$/i.exec(h);
  if (!m) return false;
  return safeEq(m[1], secret);
}

// ───────── Webhook（SSRF 防護）─────────
function parseV4(s) {
  const m = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/.exec(s);
  if (!m) return null;
  const o = m.slice(1).map(Number);
  return o.every((n) => n >= 0 && n <= 255) ? o : null;
}
function blockedV4(o) {
  const [a, b, c] = o;
  if (a === 0 || a === 10 || a === 127) return true;
  if (a === 100 && b >= 64 && b <= 127) return true;               // CGNAT
  if (a === 169 && b === 254) return true;                          // link-local（含雲端 metadata 169.254.169.254）
  if (a === 172 && b >= 16 && b <= 31) return true;
  if (a === 192 && b === 168) return true;
  if (a === 192 && b === 0 && c === 0) return true;
  if (a === 192 && b === 0 && c === 2) return true;
  if (a === 198 && (b === 18 || b === 19)) return true;
  if (a === 198 && b === 51 && c === 100) return true;
  if (a === 203 && b === 0 && c === 113) return true;
  if (a >= 224) return true;                                        // 多播／保留／廣播
  return false;
}
function expandV6(s) {
  s = s.toLowerCase().split('%')[0];
  let tail = null;
  const lastColon = s.lastIndexOf(':');
  const v4 = parseV4(s.slice(lastColon + 1));
  if (v4) { tail = [(v4[0] << 8) | v4[1], (v4[2] << 8) | v4[3]]; s = s.slice(0, lastColon + 1) + '0:0'; }
  const parts = s.split('::');
  if (parts.length > 2) return null;
  const head = parts[0] ? parts[0].split(':') : [];
  const rest = parts.length === 2 && parts[1] ? parts[1].split(':') : [];
  const fill = parts.length === 2 ? 8 - head.length - rest.length : 0;
  if (fill < 0 || (parts.length === 1 && head.length !== 8)) return null;
  const g = head.concat(new Array(fill).fill('0'), rest).map((x) => parseInt(x || '0', 16));
  if (g.length !== 8 || g.some((n) => !Number.isFinite(n) || n < 0 || n > 0xffff)) return null;
  if (tail) { g[6] = tail[0]; g[7] = tail[1]; }
  return g;
}
function isBlockedIp(ip) {
  const k = net.isIP(ip);
  if (k === 4) { const o = parseV4(ip); return !o || blockedV4(o); }
  if (k === 6) {
    const g = expandV6(ip);
    if (!g) return true;
    if (g.every((n) => n === 0)) return true;                                       // ::
    if (g.slice(0, 7).every((n) => n === 0) && g[7] === 1) return true;            // ::1
    if (g.slice(0, 5).every((n) => n === 0) && (g[5] === 0xffff || g[5] === 0)) return blockedV4([g[6] >> 8, g[6] & 255, g[7] >> 8, g[7] & 255]);   // ::ffff:a.b.c.d、::a.b.c.d
    if ((g[0] & 0xfe00) === 0xfc00) return true;                                    // fc00::/7
    if ((g[0] & 0xffc0) === 0xfe80) return true;                                    // fe80::/10
    if ((g[0] & 0xff00) === 0xff00) return true;                                    // ff00::/8
    if (g[0] === 0x64 && g[1] === 0xff9b) return true;                              // 64:ff9b::/96 NAT64
    if (g[0] === 0x2001 && g[1] === 0x0db8) return true;                            // 文件用
    if (g[0] === 0x0100 && g[1] === 0 && g[2] === 0 && g[3] === 0) return true;     // 100::/64
    return false;
  }
  return true;
}
/** 只驗證網址形狀（不做 DNS）。回傳 { ok, reason, host } */
function checkWebhookUrlShape(urlStr) {
  let u;
  try { u = new URL(String(urlStr)); } catch (e) { return { ok: false, reason: 'BAD_URL' }; }
  if (u.protocol !== 'https:') return { ok: false, reason: 'SCHEME_NOT_HTTPS' };
  if (u.username || u.password) return { ok: false, reason: 'CREDENTIALS_IN_URL' };
  if (u.port && u.port !== '443') return { ok: false, reason: 'PORT_NOT_443' };
  if (!u.hostname) return { ok: false, reason: 'NO_HOST' };
  return { ok: true, host: u.hostname, url: u };
}
/**
 * 把錨點 POST 到外部 webhook。永不 throw、不跟隨轉址、5 秒逾時；DNS 解析後所有位址都必須是公開位址，
 * 並直接連到「已驗證的那個 IP」（SNI／Host 仍用原主機名），避免驗證後 DNS 被換掉（rebinding）。
 * o: { lookup(host)->[{address}], request(opts)->ClientRequest, timeoutMs }（測試注入）
 */
async function pushAnchor(urlStr, anchor, o) {
  o = o || {};
  const shape = checkWebhookUrlShape(urlStr);
  if (!shape.ok) return { ok: false, reason: shape.reason };
  const u = shape.url;
  const host = u.hostname.replace(/^\[|\]$/g, '');
  let addrs;
  try {
    if (net.isIP(host)) addrs = [{ address: host }];
    else addrs = o.lookup ? await o.lookup(host) : await dns.promises.lookup(host, { all: true });
  } catch (e) { return { ok: false, reason: 'DNS_FAILED' }; }
  if (!Array.isArray(addrs) || !addrs.length) return { ok: false, reason: 'DNS_FAILED' };
  if (addrs.some((a) => isBlockedIp(a.address))) return { ok: false, reason: 'BLOCKED_ADDRESS' };
  const body = Buffer.from(JSON.stringify(anchor), 'utf8');
  const request = o.request || https.request;
  const timeoutMs = o.timeoutMs || 5000;
  return new Promise((resolve) => {
    let done = false;
    const fin = (r) => { if (!done) { done = true; resolve(r); } };
    try {
      const req = request({
        host: addrs[0].address, port: 443, method: 'POST', path: (u.pathname || '/') + (u.search || ''), timeout: timeoutMs,
        servername: net.isIP(host) ? undefined : host,
        headers: { Host: u.host, 'Content-Type': 'application/json', 'Content-Length': body.length, 'User-Agent': 'itts-crm-audit-anchor/1', 'X-Anchor-Seq': String(anchor.seq) },
      }, (res) => {
        const code = res.statusCode || 0;
        res.resume();
        if (code >= 200 && code < 300) fin({ ok: true, status: code });
        else if (code >= 300 && code < 400) fin({ ok: false, status: code, reason: 'REDIRECT_NOT_FOLLOWED' });
        else fin({ ok: false, status: code, reason: 'HTTP_' + code });
      });
      req.on('timeout', () => { try { req.destroy(); } catch (e) { /* 已關閉 */ } fin({ ok: false, reason: 'TIMEOUT' }); });
      req.on('error', () => fin({ ok: false, reason: 'REQUEST_FAILED' }));
      req.end(body);
    } catch (e) { fin({ ok: false, reason: 'REQUEST_FAILED' }); }
  });
}

module.exports = {
  GENESIS, LOCK_KEY, MIN_SECRET_LEN, CHAIN_MARK,
  canonical, deriveKey, getConfig, resetConfigCache, hashFields, computeRowHash, seal, verifyChain,
  makeAnchor, verifyAnchor, cronAuthorized, isBlockedIp, checkWebhookUrlShape, pushAnchor, safeEq, hmacHex,
};
