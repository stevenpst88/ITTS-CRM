'use strict';
/**
 * 防拆單提醒（quote split guard）的純函式：不碰資料庫、req、writeLog，方便單元測試。路由與稽核寫入在 lib/quoteRoutes.js。
 *
 * 「拆單」＝把本該一張的大額報價拆成多張小額，讓每張都落在較低的簽核金額級距之下。
 * 現行簽核的「金額相關門檻」在 lib/quoteApproval.js 的 AMOUNT（折扣後未稅金額，嚴格大於）：
 *   > 10,000,000 元 → 須董事長核准；> 50,000,000 元 → 須董事會決議。
 * 毛利率只決定一級主管／總經理／董事長，與金額無關；所以「疑似拆單」只看金額級距：
 *   同一商機（oppId）在回看天數內、未作廢的報價單，把「本張」加上「其他張」的折後未稅金額累計，
 *   若累計落入的金額級距「高於本張單獨的級距」，就是疑似拆單（本張靠拆開避開了較高一級的簽核）。
 *   例：本張 600 萬（單獨不需董事長）＋同商機另一張 700 萬＝累計 1,300 萬（會落入須董事長核准的級距）→ 疑似。
 * 金額基準與簽核引擎完全相同：QA.computeFinancials(q).revenueCents（折扣後、未稅、分）。
 *
 * 只提醒、記稽核；永遠不改簽核層級、不擋建單／送簽（簽核層級仍只依該張單本身決定）。
 *
 * 模式（環境變數 QUOTE_SPLIT_GUARD＝off|log|warn，預設 off）：
 *   off  完全不介入（回應、資料與基底位元級相同）
 *   log  只寫稽核 QUOTE_SPLIT_SUSPECT（回應不變）
 *   warn log ＋ 回應帶 split 資訊（建單人橫幅、送簽確認、簽核人面板橫幅與「同商機其他報價單」清單）
 * 環境變數有設定（非空）就以它為準（後台畫面唯讀顯示「由環境變數指定」）；沒設定才採用後台設定 data.quoteSplitGuard.mode（預設 off）。
 */
const QA = require('./quoteApproval');
const OL = require('./quoteOppLink');

const ENV_NAME = 'QUOTE_SPLIT_GUARD';
const MODES = Object.freeze(['off', 'log', 'warn']);
const LIMITS = Object.freeze({ maxYuan: 1e12, minDays: 1, maxDays: 730 });
// 預設門檻＝簽核引擎的金額級距（單一來源：QA.AMOUNT；後台沒設定就跟著引擎走，不另存一份寫死的數字）
const DEFAULTS = Object.freeze({ mode: 'off', tier1Yuan: QA.AMOUNT.chairman, tier2Yuan: QA.AMOUNT.board, lookbackDays: 90 });
const TIER_LABEL = Object.freeze({ 1: '須董事長核准', 2: '須董事會決議' });
const MAX_PUBLIC_SIBLINGS = 50;
const MAX_SNAPSHOT_SIBLINGS = 20;
const DAY_MS = 24 * 60 * 60 * 1000;

const isInt = (n) => typeof n === 'number' && Number.isSafeInteger(n);

/** 環境變數：{ set, mode, raw }。set＝有設定（非空）；mode＝合法值或 null（設了但不合法，視同 off，後台會標示） */
function envMode(env) {
  const raw = String(((env || process.env)[ENV_NAME]) == null ? '' : (env || process.env)[ENV_NAME]).trim().toLowerCase();
  if (!raw) return { set: false, mode: null, raw: '' };
  return { set: true, mode: MODES.includes(raw) ? raw : null, raw };
}

/** 後台設定（data.quoteSplitGuard）正規化：缺少或不合法的欄位一律回預設值（不信任儲存值） */
function getConfig(data) {
  const c = (data && data.quoteSplitGuard && typeof data.quoteSplitGuard === 'object') ? data.quoteSplitGuard : {};
  const t1 = isInt(c.tier1Yuan) && c.tier1Yuan >= 1 && c.tier1Yuan <= LIMITS.maxYuan ? c.tier1Yuan : DEFAULTS.tier1Yuan;
  let t2 = isInt(c.tier2Yuan) && c.tier2Yuan >= 1 && c.tier2Yuan <= LIMITS.maxYuan ? c.tier2Yuan : DEFAULTS.tier2Yuan;
  if (!(t2 > t1)) t2 = Math.max(DEFAULTS.tier2Yuan, t1 + 1);   // 儲存值不合理（兩級顛倒）→ 第二級退回預設或比第一級大 1
  return {
    mode: MODES.includes(c.mode) ? c.mode : DEFAULTS.mode,
    tier1Yuan: t1, tier2Yuan: t2,
    lookbackDays: isInt(c.lookbackDays) && c.lookbackDays >= LIMITS.minDays && c.lookbackDays <= LIMITS.maxDays ? c.lookbackDays : DEFAULTS.lookbackDays,
    updatedAt: typeof c.updatedAt === 'string' ? c.updatedAt : null,
    updatedBy: typeof c.updatedBy === 'string' ? c.updatedBy : null,
  };
}

/** 實際生效的模式：環境變數有設定就以它為準；否則用後台設定。設了不合法的值 → off */
function effectiveMode(cfg, env) {
  const e = envMode(env);
  if (e.set) return { mode: e.mode || 'off', source: 'env', envValue: e.raw, envInvalid: e.mode === null };
  return { mode: (cfg && cfg.mode) || 'off', source: 'config', envValue: '', envInvalid: false };
}

const numLike = (v) => (typeof v === 'number' ? v : (typeof v === 'string' && /^\d{1,13}$/.test(v.trim()) ? Number(v.trim()) : NaN));

/** 後台儲存的驗證。回 { ok:true, value:{mode,tier1Yuan,tier2Yuan,lookbackDays} } 或 { ok:false, status:400, code, error } */
function validateConfig(b) {
  const bad = (code, error) => ({ ok: false, status: 400, code, error });
  if (!b || typeof b !== 'object') return bad('BAD_SPLIT_CONFIG', '設定內容無效');
  if (!MODES.includes(b.mode)) return bad('BAD_SPLIT_MODE', '模式只能是 off（關閉）、log（只記稽核）或 warn（提醒＋稽核）');
  const t1 = numLike(b.tier1Yuan), t2 = numLike(b.tier2Yuan), days = numLike(b.lookbackDays);
  if (!isInt(t1) || t1 < 1 || t1 > LIMITS.maxYuan) return bad('BAD_SPLIT_TIER1', '第一級金額門檻必須是 1 ～ 1,000,000,000,000 的整數（元）');
  if (!isInt(t2) || t2 < 1 || t2 > LIMITS.maxYuan) return bad('BAD_SPLIT_TIER2', '第二級金額門檻必須是 1 ～ 1,000,000,000,000 的整數（元）');
  if (!(t2 > t1)) return bad('BAD_SPLIT_TIER_ORDER', '第二級金額門檻必須大於第一級');
  if (!isInt(days) || days < LIMITS.minDays || days > LIMITS.maxDays) return bad('BAD_SPLIT_DAYS', '回看天數必須是 ' + LIMITS.minDays + ' ～ ' + LIMITS.maxDays + ' 的整數');
  return { ok: true, value: { mode: b.mode, tier1Yuan: t1, tier2Yuan: t2, lookbackDays: days } };
}

/** 金額（分）→ 「1,234,567」或「1,234,567.5」（元；與簽核引擎的 fmtYuan 同格式） */
function fmtYuan(cents) {
  const c = BigInt(Math.round(Number(cents) || 0));
  const neg = c < 0n; const a = neg ? -c : c;
  const whole = (a / 100n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const frac = a % 100n;
  return (neg ? '-' : '') + (frac === 0n ? whole : whole + '.' + (frac < 10n ? '0' : '') + frac.toString());
}

/** 金額（分）落在哪一級：0＝不需額外金額簽核；1＝> tier1（董事長）；2＝> tier2（董事會）。嚴格大於，與引擎一致 */
function levelOf(cents, cfg) {
  const c = BigInt(Math.round(Number(cents) || 0));
  if (c > BigInt(cfg.tier2Yuan) * 100n) return 2;
  if (c > BigInt(cfg.tier1Yuan) * 100n) return 1;
  return 0;
}

const MAX_SAFE = BigInt(Number.MAX_SAFE_INTEGER);
const toNum = (b) => (b > MAX_SAFE ? Number.MAX_SAFE_INTEGER : Number(b));
const approvalStateOf = (q) => (q && q.approval && q.approval.state) || 'none';

/**
 * 評估一張報價單是否疑似拆單。
 *   data：整份資料（讀 data.quotations）；q：被評估的報價單（可已在 data.quotations 內；自己不計入「其他張」）
 *   opts.nowMs：評估時間（毫秒）；opts.cfg：getConfig() 的結果
 * 回傳：
 *   { applicable:false, reason:'NO_OPP'|'NO_AMOUNT' }  沒有商機連結（舊單）或本張金額算不出來（沒有品項／折扣不合法）→ 不評估
 *   { applicable:true, suspect, oppId, ownCents, cumCents, ownLevel, cumLevel, thresholdYuan, windowDays, countedCount, siblings:[…] }
 * siblings＝同商機所有「其他」報價單（含失效／超出回看天數／算不出金額的，各自標 counted:false 與 why），每筆帶 q 參照（toPublic／snapshot 會拿掉）。
 */
function evaluate(data, q, opts) {
  const cfg = (opts && opts.cfg) || getConfig(data);
  const nowMs = opts && Number.isFinite(opts.nowMs) ? opts.nowMs : Date.now();
  if (!q || !q.oppId) return { applicable: false, reason: 'NO_OPP' };
  const own = QA.computeFinancials(q);
  if (!own || !own.ok || !(own.revenueCents > 0)) return { applicable: false, reason: 'NO_AMOUNT' };
  const start = nowMs - cfg.lookbackDays * DAY_MS;
  let sum = BigInt(own.revenueCents);
  let counted = 0;
  const siblings = [];
  OL.linkedQuotes(data, q.oppId).forEach((x) => {
    if (x === q || (x.id && x.id === q.id)) return;
    const s = { q: x, id: x.id, quoteNo: x.quoteNo || '', owner: x.owner || '', createdAt: x.createdAt || null, approvalState: approvalStateOf(x), revenueCents: null, counted: false, why: '' };
    if (x.invalid) s.why = 'INVALID';                                  // 連結的商機已作廢 → 報價單失效，不計
    else if (x.status === 'rejected') s.why = 'LEGACY_REJECTED';       // 簽核上線前的舊狀態「已拒絕」，不計
    else {
      const t = Date.parse(x.createdAt);
      if (Number.isFinite(t) && t < start) s.why = 'OUT_OF_WINDOW';   // 建立時間早於回看起點（剛好等於起點算在內）；沒有可解析的建立時間＝保守計入
      else {
        const f = QA.computeFinancials(x);
        if (!f || !f.ok || !(f.revenueCents > 0)) s.why = 'NO_AMOUNT';
        else { s.revenueCents = f.revenueCents; s.counted = true; sum += BigInt(f.revenueCents); counted++; }
      }
    }
    siblings.push(s);
  });
  siblings.sort((a, b) => String(b.createdAt || '').localeCompare(String(a.createdAt || '')));
  const ownLevel = levelOf(own.revenueCents, cfg);
  const cumCents = toNum(sum);
  const cumLevel = levelOf(cumCents, cfg);
  const suspect = cumLevel > ownLevel;
  return {
    applicable: true, suspect, oppId: q.oppId,
    ownCents: own.revenueCents, cumCents, ownLevel, cumLevel,
    thresholdYuan: cumLevel > 0 ? (cumLevel === 2 ? cfg.tier2Yuan : cfg.tier1Yuan) : null,
    windowDays: cfg.lookbackDays, countedCount: counted, siblings,
  };
}

/**
 * 依檢視者過濾後的公開形狀（只含金額與狀態，沒有成本／毛利）。
 *   canView(q)：檢視者看不看得到那張單（路由用 relationOf(...).canView）；看不到的只計入 hiddenCount，不給單號與金額。
 *   decorate(s)：可選，路由用來補 ownerName。
 * 累計金額與張數是彙總數字：只給「看得到本張單」的人（呼叫端已保證），也是簽核人判斷所需。
 */
function toPublic(res, canView, decorate) {
  if (!res || !res.applicable) return null;
  const vis = [];
  let hidden = 0;
  res.siblings.forEach((s) => {
    if (!canView(s.q)) { hidden++; return; }
    const o = { id: s.id, quoteNo: s.quoteNo, createdAt: s.createdAt, approvalState: s.approvalState, revenueCents: s.revenueCents, counted: s.counted };
    if (s.why) o.why = s.why;
    vis.push(decorate ? decorate(o, s) : o);
  });
  return {
    suspect: res.suspect, ownCents: res.ownCents, cumCents: res.cumCents, ownLevel: res.ownLevel, cumLevel: res.cumLevel,
    thresholdYuan: res.thresholdYuan, windowDays: res.windowDays, countedCount: res.countedCount,
    siblings: vis.slice(0, MAX_PUBLIC_SIBLINGS), siblingsMore: Math.max(0, vis.length - MAX_PUBLIC_SIBLINGS), hiddenCount: hidden,
  };
}

/** 送簽時凍結在 q.splitGuard 的快照（「送簽時為準」；簽核人事後看到的就是這一份）。只存計入累計的其他張 */
function snapshot(res, meta) {
  if (!res || !res.applicable || !res.suspect) return null;
  return {
    v: 1, at: meta.at, by: meta.by, trigger: meta.trigger, mode: meta.mode,
    ownCents: res.ownCents, cumCents: res.cumCents, ownLevel: res.ownLevel, cumLevel: res.cumLevel,
    thresholdYuan: res.thresholdYuan, windowDays: res.windowDays,
    siblings: res.siblings.filter((s) => s.counted).slice(0, MAX_SNAPSHOT_SIBLINGS).map((s) => ({ id: s.id, quoteNo: s.quoteNo, revenueCents: s.revenueCents, approvalState: s.approvalState })),
  };
}

/** 快照 → 公開形狀（依檢視者過濾：看不到的其他張只計入 hiddenCount） */
function snapshotToPublic(snap, canViewId) {
  if (!snap || typeof snap !== 'object') return null;
  const sibs = Array.isArray(snap.siblings) ? snap.siblings : [];
  let hidden = 0;
  const vis = [];
  sibs.forEach((s) => { if (canViewId(s.id)) vis.push({ id: s.id, quoteNo: s.quoteNo, revenueCents: s.revenueCents, approvalState: s.approvalState }); else hidden++; });
  return {
    at: snap.at, trigger: snap.trigger, ownCents: snap.ownCents, cumCents: snap.cumCents, thresholdYuan: snap.thresholdYuan, windowDays: snap.windowDays,
    siblings: vis, hiddenCount: hidden,
  };
}

const TRIGGER_LABEL = { CREATE: '建立報價單', SUBMIT: '送簽' };

/** 稽核 detail：人讀的一行文字＋ meta=JSON（管理員查詢 API 會解析）。不含成本／毛利 */
function auditDetail(res, q, meta) {
  const sibs = res.siblings.filter((s) => s.counted);
  const sibTotal = res.cumCents - res.ownCents;
  const m = {
    trigger: meta.trigger, mode: meta.mode, oppId: res.oppId, company: q.company || '',
    ownCents: res.ownCents, cumCents: res.cumCents, thresholdYuan: res.thresholdYuan, ownLevel: res.ownLevel, cumLevel: res.cumLevel, windowDays: res.windowDays,
    siblings: sibs.slice(0, 10).map((s) => ({ no: s.quoteNo, c: s.revenueCents, st: s.approvalState })), siblingsMore: Math.max(0, sibs.length - 10),
  };
  const human = '疑似拆單（' + (TRIGGER_LABEL[meta.trigger] || meta.trigger) + '）｜商機 ' + res.oppId + '（' + (q.company || '') + '）'
    + '｜本張 ' + fmtYuan(res.ownCents) + ' 元（折扣後未稅）＋同商機 ' + sibs.length + ' 張 ' + fmtYuan(sibTotal) + ' 元＝累計 ' + fmtYuan(res.cumCents) + ' 元'
    + '，超過 ' + fmtYuan(res.thresholdYuan * 100) + ' 元門檻（' + TIER_LABEL[res.cumLevel] + '；本張單獨為' + (res.ownLevel ? TIER_LABEL[res.ownLevel] : '未達門檻') + '）'
    + '｜回看 ' + res.windowDays + ' 天｜其他單號：' + (sibs.slice(0, 10).map((s) => s.quoteNo + '(' + fmtYuan(s.revenueCents) + ')').join('、') || '無');
  return human + '｜meta=' + JSON.stringify(m);
}

/** 從稽核 detail 取回 meta（解析失敗回 null） */
function parseAuditMeta(detail) {
  const i = String(detail || '').lastIndexOf('｜meta=');
  if (i < 0) return null;
  try { return JSON.parse(String(detail).slice(i + '｜meta='.length)); } catch (e) { return null; }
}

/** 後台設定儲存的稽核文字（before→after；只列有變的欄位） */
function configDiff(prev, next) {
  const parts = [];
  if (prev.mode !== next.mode) parts.push('模式 ' + prev.mode + '→' + next.mode);
  if (prev.tier1Yuan !== next.tier1Yuan) parts.push('第一級門檻 ' + fmtYuan(prev.tier1Yuan * 100) + '→' + fmtYuan(next.tier1Yuan * 100) + ' 元');
  if (prev.tier2Yuan !== next.tier2Yuan) parts.push('第二級門檻 ' + fmtYuan(prev.tier2Yuan * 100) + '→' + fmtYuan(next.tier2Yuan * 100) + ' 元');
  if (prev.lookbackDays !== next.lookbackDays) parts.push('回看天數 ' + prev.lookbackDays + '→' + next.lookbackDays);
  return parts;
}

module.exports = {
  ENV_NAME, MODES, LIMITS, DEFAULTS, TIER_LABEL,
  envMode, getConfig, effectiveMode, validateConfig, levelOf, fmtYuan,
  evaluate, toPublic, snapshot, snapshotToPublic, auditDetail, parseAuditMeta, configDiff,
};
