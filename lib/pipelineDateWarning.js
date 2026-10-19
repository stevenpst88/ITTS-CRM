'use strict';
/**
 * 商機動態報表「預計簽約日調整」的反覆延後警示（純函式，無 I/O，方便單元測試）。
 *
 * 規則（常數可由環境變數覆寫，每次呼叫時才讀，測試可動態改）：
 *   標記 repeated = 階段為 B 或 C
 *                  且（累計延後次數 >= PIPELINE_REPEAT_MIN_COUNT(2) 或 累計延後天數 >= PIPELINE_REPEAT_MIN_DAYS(30)）
 *                  且 目前階段停留天數 > PIPELINE_REPEAT_STALL_DAYS(45)
 *   Won／D 與其他階段一律不標。
 *
 * 累計口徑（用 opportunityDateChanges 的「全歷史」，不侷限本期）：
 *   - 紀錄依 changedAt 升冪；oldDate 為空的（建立商機時的首次設定）與 newDate 為空的紀錄忽略。
 *   - 累計延後次數 = 每筆 newDate > oldDate 的紀錄數（只計往後延；提前不扣次數）。
 *   - 累計延後天數 = max(0, 最後有效 newDate - 第一筆有效 oldDate)（以第一筆 oldDate 為基準的淨延後；來回抵銷後為 0）。
 *   - 停留天數 = now - 最後一次 stageHistory.to === 目前階段 的 date；沒有就用 createdAt；都沒有 → null（不標）。
 */
const DAY = 86400000;

function envNum(name, def) {
  const v = Number(process.env[name]);
  return Number.isFinite(v) && process.env[name] !== '' && process.env[name] != null ? v : def;
}
function getRules() {
  return {
    minCount: envNum('PIPELINE_REPEAT_MIN_COUNT', 2),
    minDays: envNum('PIPELINE_REPEAT_MIN_DAYS', 30),
    stallDays: envNum('PIPELINE_REPEAT_STALL_DAYS', 45),
    stages: ['B', 'C'],
  };
}

function dayNum(s) {            // 'YYYY-MM-DD'(或 ISO) → 以 UTC 日為單位的整數；無法解析 → NaN
  const t = Date.parse(String(s || '').slice(0, 10));
  return Number.isNaN(t) ? NaN : Math.round(t / DAY);
}

function cumulativePostpone(records) {
  const list = (records || [])
    .filter(c => c && c.oldDate && c.newDate && !Number.isNaN(dayNum(c.oldDate)) && !Number.isNaN(dayNum(c.newDate)))
    .slice()
    .sort((a, b) => new Date(a.changedAt) - new Date(b.changedAt));
  if (!list.length) return { count: 0, days: 0 };
  let count = 0;
  list.forEach(c => { if (dayNum(c.newDate) > dayNum(c.oldDate)) count++; });
  const net = dayNum(list[list.length - 1].newDate) - dayNum(list[0].oldDate);
  return { count, days: Math.max(0, net) };
}

function stageStayDays(opp, now) {
  let since = null;
  const hist = Array.isArray(opp && opp.stageHistory) ? opp.stageHistory : [];
  for (const h of hist) {
    if (h && h.to === opp.stage && h.date && (!since || new Date(h.date) > new Date(since))) since = h.date;
  }
  if (!since) since = opp && opp.createdAt;
  const t = since ? new Date(since).getTime() : NaN;
  if (Number.isNaN(t)) return null;
  return Math.max(0, Math.floor(((now ? now.getTime() : Date.now()) - t) / DAY));
}

function evaluate(opp, records, now) {
  const rules = getRules();
  const cum = cumulativePostpone(records);
  const stageDays = stageStayDays(opp, now);
  const repeated = !!(
    rules.stages.includes(opp && opp.stage) &&
    (cum.count >= rules.minCount || cum.days >= rules.minDays) &&
    stageDays != null && stageDays > rules.stallDays
  );
  return { postponeCount: cum.count, postponeDays: cum.days, stageDays, repeated };
}

/** 排序：repeated 優先 → 累計延後天數大 → 本期延後天數大 → 最近異動在前 */
function compareEntries(a, b) {
  if (!!b.repeated !== !!a.repeated) return (b.repeated ? 1 : 0) - (a.repeated ? 1 : 0);
  if ((b.postponeDays || 0) !== (a.postponeDays || 0)) return (b.postponeDays || 0) - (a.postponeDays || 0);
  const da = a.daysDiff ?? -1, db = b.daysDiff ?? -1;
  if (db !== da) return db - da;
  return (b.changedAt || '').localeCompare(a.changedAt || '');
}

/** 台北（UTC+8，無日光節約）當月 [月初 00:00, 月底 23:59:59] 的絕對時間；不依賴 server 時區（Vercel=UTC） */
function taipeiMonthRange(now) {
  const t = new Date((now || new Date()).getTime() + 8 * 3600000);
  const y = t.getUTCFullYear(), m = t.getUTCMonth();
  return {
    from: new Date(Date.UTC(y, m, 1) - 8 * 3600000),
    to:   new Date(Date.UTC(y, m + 1, 0, 23, 59, 59) - 8 * 3600000),
  };
}

module.exports = { getRules, cumulativePostpone, stageStayDays, evaluate, compareEntries, taipeiMonthRange };
