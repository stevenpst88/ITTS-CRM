// ════════════════════════════════════════════════════════════
//  歷史商機預計簽約日補建腳本
//
//  為所有現有商機補寫初始 opportunityDateChanges 記錄，
//  讓 Pipeline 月度變動報表能顯示歷史數據。
//
//  每筆商機只補一筆（oldDate: null → newDate: expectedDate），
//  changedAt 使用商機的 createdAt，重複執行安全（冪等）。
//
//  用法：
//    node scripts/migrate-opp-date-changes.js
//    DB_BACKEND=postgres node scripts/migrate-opp-date-changes.js
// ════════════════════════════════════════════════════════════
require('dotenv').config();
const db = require('../db');

(async () => {
  try {
    if (typeof db.ready === 'function') await db.ready();

    const data = db.load();
    const opps = data.opportunities || [];

    if (!data.opportunityDateChanges) data.opportunityDateChanges = [];
    const existing = data.opportunityDateChanges;

    // 已有初始記錄的 dealId（oldDate === null）
    const alreadySeeded = new Set(
      existing.filter(c => c.oldDate === null).map(c => c.dealId)
    );

    let added = 0, skipped = 0, noDate = 0;

    for (const opp of opps) {
      if (!opp.expectedDate) { noDate++; continue; }
      if (alreadySeeded.has(opp.id)) { skipped++; continue; }

      existing.push({
        dealId:    opp.id,
        dealValue: parseFloat(opp.amount) || 0,
        oldDate:   null,
        newDate:   opp.expectedDate,
        changedAt: opp.createdAt || new Date().toISOString(),
        owner:     opp.owner || '',
      });
      added++;
    }

    if (added > 0) {
      await db.save(data);
      if (typeof db.flush === 'function') await db.flush();
    }

    console.log('════════════════════════════════════');
    console.log('  Pipeline 月度變動 — 歷史補建完成');
    console.log('════════════════════════════════════');
    console.log(`  ✅ 新增記錄：${added} 筆`);
    console.log(`  ⏭  已存在（略過）：${skipped} 筆`);
    console.log(`  ⚠️  無預計簽約日（略過）：${noDate} 筆`);
    console.log(`  📦 合計商機數：${opps.length} 筆`);
    console.log('════════════════════════════════════');
    process.exit(0);
  } catch (err) {
    console.error('❌ 失敗:', err);
    process.exit(1);
  }
})();
