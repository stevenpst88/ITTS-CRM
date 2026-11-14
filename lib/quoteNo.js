'use strict';
/**
 * 報價單號純函式（不碰資料庫）。單號＝QU-<台灣年月>-<流水號(至少3位)>。
 * floor：blob 現況下的「已用過最大序號」，定義與 server.js 的 genQuoteNo 完全相同（quoteSeq 計數器與現有報價單最大號取大）。
 * 原子計數器模式（QUOTE_NO_ATOMIC=1）以它當下界，確保舊資料與關掉開關後都能銜接。
 */
function format(ym, serial) { return `QU-${ym}-${String(serial).padStart(3, '0')}`; }

function floorOf(data, ym) {
  const prefix = `QU-${ym}-`;
  let maxSerial = 0;
  ((data && data.quotations) || []).forEach((q) => {
    if (q && q.quoteNo && q.quoteNo.startsWith(prefix)) {
      const n = parseInt(q.quoteNo.slice(prefix.length), 10);
      if (n > maxSerial) maxSerial = n;
    }
  });
  const seq = (data && data.quoteSeq && data.quoteSeq[ym]) || 0;
  return Math.max(seq, maxSerial);
}

/** 把已發出的序號記回 blob 的計數器（只增不減），這樣關掉原子模式後舊邏輯不會重用號碼 */
function recordSeq(data, ym, serial) {
  if (!data.quoteSeq) data.quoteSeq = {};
  if ((data.quoteSeq[ym] || 0) < serial) data.quoteSeq[ym] = serial;
}

module.exports = { format, floorOf, recordSeq };
