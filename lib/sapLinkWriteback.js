'use strict';
/**
 * SAP 批次推送「結果寫回」：把這次推送新增／更新的 CRM↔SAP 對照（integrationLinks）與執行紀錄（integrationLogs）寫進最新的主資料。
 *
 * 為什麼不能像 AI 快取那樣「失敗就算了」：SAP 端已經建立了帳號／聯絡人，本機卻沒記下 sapId，
 * 下一次推送找不到對照就會當成未同步，改走 POST，在 SAP 重複建立一份（DB_WRITE_SERIALIZE=1 時，寫回改成短鎖＋重讀，
 * 取不到鎖就有可能發生；關閉開關時寫進快取並背景存檔，不會因鎖忙而丟失）。所以：
 *  1. 寫回用較長的等待（8 秒）並重試 3 次，且要確認存檔真的成功（見 writeGuard.applyOnLatest）。
 *  2. 仍失敗：把待補寫的對照與紀錄留在「行程記憶體」，下一次 push/batch 一開始先併進它使用的資料（避免重複建立），
 *     結尾一併補寫；回應帶 saved:false＋warning，前端提示管理員「請勿重推，請聯絡 IT」。
 *  限制：行程記憶體隨實例回收而消失（Vercel 換實例／冷啟動），補寫只在「同一個實例收到下一次 push/batch」時發生。
 *  所以 saved:false 的警告是主要防線，記憶體補寫是輔助。
 */

const PENDING_LINK_CAP = 5000;
const PENDING_LOG_CAP = 20;
const WRITEBACK_OPTS = { waitMs: 8000, holdMs: 10000, attempts: 3, retryDelayMs: 500 };
const WARNING = 'SAP 已推送，但本機的 SAP 對照沒有寫入成功。請勿重新推送（會在 SAP 重複建立帳號），請聯絡 IT 處理。';

const keyOf = (l) => l.crmType + '|' + l.crmId + '|' + l.integrationId;

/**
 * @param {object} o
 * @param {function} o.applyOnLatest (label, legacyData, mutate, opts) => Promise<boolean>
 * @param {object} [o.opts] 覆寫等待／重試設定（測試用）
 * @param {object} [o.log] console 相容
 */
function createSapLinkWriteback(o) {
  const applyOnLatest = o.applyOnLatest;
  const opts = Object.assign({}, WRITEBACK_OPTS, o.opts || {});
  const log = o.log || console;
  const pendingLinks = new Map();   // key -> link
  const pendingLogs = new Map();    // logId -> logEntry

  /** 把「上次沒寫成功」的對照併進這次要用的資料（記憶體內），讓這次推送找得到既有的 sapId、不會重複建立 */
  function mergePendingInto(data) {
    if (!pendingLinks.size) return 0;
    if (!Array.isArray(data.integrationLinks)) data.integrationLinks = [];
    let n = 0;
    for (const l of pendingLinks.values()) {
      const e = data.integrationLinks.find((x) => keyOf(x) === keyOf(l));
      if (!e) { data.integrationLinks.push(Object.assign({}, l)); n++; }
      else if (!e.sapId && l.sapId) { e.sapId = l.sapId; e.syncedAt = l.syncedAt; n++; }
    }
    if (n) log.warn('[sap writeback] 已把上次沒寫成功的 ' + n + ' 筆 SAP 對照併入這次推送使用的資料（避免重複建立）');
    return n;
  }

  /**
   * 推送結束後呼叫。回傳 true＝對照與紀錄已存入；false＝沒存成功（已留在記憶體等下次補寫，呼叫端要在回應帶警告）
   * @param {object} data      推送期間使用的資料物件（快取）
   * @param {object} logEntry  這次的完整執行紀錄（endedAt 已填）
   */
  async function commit(data, logEntry) {
    const fresh = (data.integrationLinks || []).filter((l) => l.syncedAt && l.syncedAt >= logEntry.startedAt).map((l) => Object.assign({}, l));
    const links = new Map(pendingLinks);
    for (const l of fresh) links.set(keyOf(l), l);
    const logs = new Map(pendingLogs);
    logs.set(logEntry.id, logEntry);
    const extra = pendingLinks.size > 0 || pendingLogs.size > 0;

    const mutate = (d) => {
      if (d === data && !extra) return true;   // 舊行為（或最新資料就是推送時用的同一個物件）：對照與紀錄本來就在裡面，直接存
      d.integrationLinks = d.integrationLinks || [];
      d.integrationLogs = d.integrationLogs || [];
      for (const l of links.values()) {
        const e = d.integrationLinks.find((x) => keyOf(x) === keyOf(l));
        if (e) { e.sapId = l.sapId; e.syncedAt = l.syncedAt; } else d.integrationLinks.push(Object.assign({}, l));
      }
      // 執行紀錄：最新資料裡若已有同 id 的半成品（推送期間別的請求把「endedAt:null、records 不完整」的版本存進 DB），整筆取代為完整版本；沒有才加到最前面
      for (const lg of logs.values()) {
        const i = d.integrationLogs.findIndex((x) => x.id === lg.id);
        if (i >= 0) d.integrationLogs[i] = lg; else d.integrationLogs.unshift(lg);
      }
      if (d.integrationLogs.length > 200) d.integrationLogs.length = 200;
    };

    const ok = await applyOnLatest('sap push/batch result', data, mutate, opts);
    if (ok) {
      pendingLinks.clear(); pendingLogs.clear();
      return true;
    }
    // 沒存成功：留在記憶體
    for (const [k, l] of links) pendingLinks.set(k, l);
    for (const [k, lg] of logs) pendingLogs.set(k, lg);
    while (pendingLinks.size > PENDING_LINK_CAP) pendingLinks.delete(pendingLinks.keys().next().value);
    while (pendingLogs.size > PENDING_LOG_CAP) pendingLogs.delete(pendingLogs.keys().next().value);
    log.error('[sap writeback] SAP 對照／紀錄沒有寫入成功，已留在行程記憶體等下次補寫（對照 ' + pendingLinks.size + ' 筆）。請勿重推這批公司。');
    return false;
  }

  return { mergePendingInto, commit, pendingCount: () => pendingLinks.size };
}

module.exports = { createSapLinkWriteback, WARNING, WRITEBACK_OPTS };
