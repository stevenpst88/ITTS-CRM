// ════════════════════════════════════════════════════════════
//  GCIS 企業主檔補全 — 本機（台灣 IP）執行版
//
//  為什麼：Vercel 機房 IP 打 GCIS 會被限流/擋掉（逾時），但本機台灣 IP
//  打 GCIS 很順。本腳本「直接連正式 Supabase」，從你本機查 GCIS、把結果
//  寫回正式企業主檔。
//
//  安全保證：
//   · 只讀寫正式資料庫的「主資料 blob」，且【只修改 companies 裡的 GCIS 欄位】
//     （name / capital / address / representative / gcisEnriched / gcisTaxIdError）
//   · 不碰名片 / 商機 / 合約 / 帳款 / 帳號…任何其他資料
//   · 不會讀取、也不會上傳你本機的 data.json（測試資料）
//   · 每批寫入前「重讀最新 blob」再只套用公司變更 → 不蓋掉他人即時編輯
//
//  用法：
//    node scripts/gcis-backfill-local.js              # 試跑（dry-run，只看不寫）
//    node scripts/gcis-backfill-local.js --apply      # 正式執行（寫回正式庫）
//    node scripts/gcis-backfill-local.js --apply --limit 50   # 只補前 50 家
// ════════════════════════════════════════════════════════════
require('dotenv').config();

if (!process.env.DATABASE_URL) {
  console.error('✗ 找不到 DATABASE_URL（.env）。無法連正式庫，中止。');
  process.exit(1);
}

const pg = require('../db/postgres');   // 直接用 _loadAsync / _saveAsync（不走快取）

const APPLY = process.argv.includes('--apply');
const argLimit = (() => { const i = process.argv.indexOf('--limit'); return i >= 0 ? parseInt(process.argv[i + 1]) || 0 : 0; })();
const BATCH = 6;            // 每批並行查幾家（本機 IP 對 GCIS 友善）
const GAP_MS = 400;         // 批次間隔
const DRY_SAMPLE = (() => { const i = process.argv.indexOf('--sample'); return i >= 0 ? (parseInt(process.argv[i + 1]) || 8) : 8; })();  // 試跑樣本數

const sleep = ms => new Promise(r => setTimeout(r, ms));

// ── 與 server.js fetchGcisCompany 同一支 API / 同邏輯 ──
async function fetchGcisCompany(taxId) {
  const t = String(taxId || '').trim();
  if (!/^\d{8}$/.test(t)) return { notFound: true };
  try {
    const r = await fetch(
      `https://data.gcis.nat.gov.tw/od/data/api/5F64D864-61CB-4D0D-8AD9-492047CC1EA6?$format=json&$filter=Business_Accounting_NO eq ${t}&$skip=0&$top=1`,
      { headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' }, signal: AbortSignal.timeout(8000) }
    );
    if (!r.ok) return null;                  // 5xx → 暫時性，可重試
    const text = await r.text();
    if (!text || !text.trim()) return { notFound: true };  // 200 + 空 body → GCIS 無此公司登記 → 查無（非逾時！）
    let data;
    try { data = JSON.parse(text); } catch { return { notFound: true }; }  // 200 但非 JSON → 視為查無
    const row = data && data[0];
    if (!row) return { notFound: true };     // 200 + 空陣列 → 統編查無
    const capRaw = row.Capital_Stock_Amount;
    return {
      name: row.Company_Name || '',
      capital: (capRaw != null && capRaw !== '' && !isNaN(parseInt(capRaw))) ? parseInt(capRaw) : null,
      address: row.Company_Location || '',
      representative: row.Responsible_Name || '',
    };
  } catch { return null; }                   // 逾時/網路錯 → 可重試
}

function isTarget(m) {
  return /^\d{8}$/.test(String(m.taxId || '').trim()) && !m.gcisEnriched && !m.gcisNoData && !m.gcisTaxIdError;
}

// 只把 GCIS 結果套到「公司物件」的 GCIS 欄位（其餘欄位完全不動）
function applyInfo(m, info) {
  if (info === null) return 'failed';        // 逾時 → 不改，下次重試
  if (info.notFound) {
    m.gcisTaxIdError = true;
    m.gcisTaxIdErrorAt = new Date().toISOString();
    m.updatedAt = new Date().toISOString();
    return 'taxIdError';
  }
  if (info.name) m.name = info.name;
  if (info.capital != null) m.capital = info.capital;
  if (info.address && !m.address) m.address = info.address;
  if (info.representative && !m.representative) m.representative = info.representative;
  m.gcisEnriched = true;
  m.gcisNoData = false;
  m.gcisTaxIdError = false;
  m.updatedAt = new Date().toISOString();
  return 'enriched';
}

function maskDbHost() {
  try { const u = new URL(process.env.DATABASE_URL); return u.host; } catch { return '(無法解析)'; }
}

(async () => {
  console.log('連線目標（正式庫 host）：', maskDbHost());
  console.log('模式：', APPLY ? '🔴 正式執行（會寫回正式庫）' : '🟢 試跑 dry-run（只看不寫）');

  const data = await pg._loadAsync();
  const companies = data.companies || [];
  const targets = companies.filter(isTarget);
  console.log(`\n企業主檔總數：${companies.length} 家`);
  console.log(`待 GCIS 補全（有 8 碼統編、未補全/未標記）：${targets.length} 家`);
  const noTaxId = companies.filter(m => !/^\d{8}$/.test(String(m.taxId || '').trim())).length;
  console.log(`（沒統編/格式錯，自動跳過：${noTaxId} 家）`);

  const work = argLimit > 0 ? targets.slice(0, argLimit) : targets;

  if (!APPLY) {
    const sample = work.slice(0, DRY_SAMPLE);
    console.log(`\n── 試跑樣本（前 ${sample.length} 家實際查 GCIS，不寫入）──`);
    for (const m of sample) {
      const info = await fetchGcisCompany(m.taxId);
      const verdict = info === null ? '逾時/錯誤（會留待重試）'
        : info.notFound ? '⚠️ GCIS查無（非公司登記/統編誤 → 跳過）'
        : `✓ 會補：「${info.name}」 資本額 ${info.capital ?? '—'}`;
      console.log(`  ${m.taxId}  ${(m.name || '').slice(0, 16).padEnd(16)} → ${verdict}`);
      await sleep(200);
    }
    console.log(`\n🟢 這是試跑、未寫入任何資料。確認無誤後執行：`);
    console.log(`    node scripts/gcis-backfill-local.js --apply${argLimit ? ' --limit ' + argLimit : ''}`);
    await pg._poolEnd?.();
    process.exit(0);
  }

  // ── 正式執行 ──
  let enriched = 0, taxIdError = 0, failed = 0, done = 0;
  const total = work.length;
  const ids = work.map(m => m.id);

  for (let i = 0; i < ids.length; i += BATCH) {
    const batchIds = ids.slice(i, i + BATCH);
    // 先（用目前 blob 的對應公司）查 GCIS
    const byId = new Map(companies.map(c => [c.id, c]));
    const infos = await Promise.allSettled(
      batchIds.map(async id => ({ id, info: await fetchGcisCompany(byId.get(id)?.taxId) }))
    );

    // 重讀最新 blob，只把本批公司的 GCIS 欄位套上去再寫回（避免蓋掉他人即時編輯）
    const fresh = await pg._loadAsync();
    const freshById = new Map((fresh.companies || []).map(c => [c.id, c]));
    let batchTouched = false;
    for (const r of infos) {
      if (r.status !== 'fulfilled') { failed++; done++; continue; }
      const { id, info } = r.value;
      const fm = freshById.get(id);
      if (!fm) { done++; continue; }          // 公司可能已被刪 → 略過
      const res = applyInfo(fm, info);
      if (res === 'enriched') enriched++;
      else if (res === 'taxIdError') taxIdError++;
      else failed++;
      if (res !== 'failed') batchTouched = true;
      done++;
    }
    if (batchTouched) await pg._saveAsync(fresh);

    const pct = Math.round(done / total * 100);
    process.stdout.write(`\r進度 ${done}/${total} (${pct}%)  ·  已補 ${enriched}  ·  GCIS查無 ${taxIdError}  ·  逾時 ${failed}   `);
    await sleep(GAP_MS);
  }

  console.log(`\n\n✅ 完成：已補 ${enriched} 家、GCIS查無(跳過) ${taxIdError} 家、逾時待重試 ${failed} 家。`);
  if (failed) console.log('   逾時的可再執行一次本腳本繼續補（已補/已標記的不會重複）。');
  await pg._poolEnd?.();
  process.exit(0);
})().catch(async e => {
  console.error('\n✗ 發生錯誤：', e.message);
  process.exit(1);
});
