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
const RECAP = process.argv.includes('--recap');   // 只重補「已補但資本額 0/空」的公司（修登記資本額=0 漏實收的舊資料）
const STATUS = process.argv.includes('--status'); // 回填「已補但無公司狀態」的公司（補 GCIS 公司狀態：合併解散/解散…）
const RETRY_ERROR = process.argv.includes('--retry-error'); // 重試「GCIS查無(gcisTaxIdError)」的公司（搭配證交所 fallback 補回被漏掉的上市公司）
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
    // 資本額：登記資本額為 0/空 → 改用實收資本額
    const pickCapital = (...vals) => { for (const v of vals) { const n = parseInt(v); if (!isNaN(n) && n > 0) return n; } return null; };
    return {
      name: row.Company_Name || '',
      capital: pickCapital(row.Capital_Stock_Amount, row.Paid_In_Capital_Amount),
      address: row.Company_Location || '',
      representative: row.Responsible_Name || '',
      status: row.Company_Status_Desc || '',
    };
  } catch { return null; }                   // 逾時/網路錯 → 可重試
}

// ── 證交所上市清單：統編 → 名稱/實收資本額（補 GCIS 漏掉的大型上市公司）──
let twMap = {};
async function loadTwseMap() {
  try {
    const r = await fetch('https://openapi.twse.com.tw/v1/opendata/t187ap03_L', { headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' }, signal: AbortSignal.timeout(30000) });
    const arr = JSON.parse(await r.text());
    (arr || []).forEach(row => {
      const id = String(row['營利事業統一編號'] || '').trim();
      const cap = parseInt(row['實收資本額']);
      if (/^\d{8}$/.test(id)) twMap[id] = { name: row['公司名稱'] || '', capital: (!isNaN(cap) && cap > 0) ? cap : null };
    });
  } catch (e) { console.warn('證交所清單載入失敗（fallback 停用）:', e.message); }
}

// GCIS 查無時 → 改查證交所上市清單
async function resolveCompany(taxId) {
  const g = await fetchGcisCompany(taxId);
  if (g && g.notFound) {
    const t = twMap[String(taxId || '').trim()];
    if (t && (t.capital || t.name)) return { name: t.name, capital: t.capital, address: '', representative: '', status: '上市（證交所）' };
  }
  return g;
}

function has8(m) { return /^\d{8}$/.test(String(m.taxId || '').trim()); }
function isTarget(m) {
  return has8(m) && !m.gcisEnriched && !m.gcisNoData && !m.gcisTaxIdError;
}
// --recap：已補全但資本額 0/空 → 重查以補回實收資本額
function isRecapTarget(m) { return has8(m) && m.gcisEnriched && !m.capital; }

// --recap 用：只在 GCIS 有資料時更新（保守，不把已補的降級成查無）
function applyRecap(m, info) {
  if (info === null || info.notFound) return 'skipped';   // 逾時或查無 → 完全不動
  let changed = false;
  if (info.capital != null && info.capital !== m.capital) { m.capital = info.capital; changed = true; }
  if (info.name && info.name !== m.name) { m.name = info.name; changed = true; }
  if (info.address && !m.address) { m.address = info.address; changed = true; }
  if (info.representative && !m.representative) { m.representative = info.representative; changed = true; }
  if (info.status && info.status !== m.gcisStatus) { m.gcisStatus = info.status; changed = true; }
  if (changed) m.updatedAt = new Date().toISOString();
  return changed ? 'updated' : 'nochange';
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
  if (info.status) m.gcisStatus = info.status;
  m.gcisEnriched = true;
  m.gcisNoData = false;
  m.gcisTaxIdError = false;
  m.updatedAt = new Date().toISOString();
  return 'enriched';
}

// --retry-error：重試「GCIS查無」的公司（搭配證交所 fallback，補回被漏掉的上市公司）
function isRetryErrorTarget(m) { return has8(m) && m.gcisTaxIdError; }
// --status：已補全但無公司狀態 → 重查補上 gcisStatus（保守，GCIS 無資料時不動）
function isStatusTarget(m) { return has8(m) && m.gcisEnriched && !m.gcisStatus; }
function applyStatus(m, info) {
  if (info === null || info.notFound) return 'skipped';
  if (info.status && info.status !== m.gcisStatus) { m.gcisStatus = info.status; m.updatedAt = new Date().toISOString(); return 'updated'; }
  return 'nochange';
}

function maskDbHost() {
  try { const u = new URL(process.env.DATABASE_URL); return u.host; } catch { return '(無法解析)'; }
}

(async () => {
  console.log('連線目標（正式庫 host）：', maskDbHost());
  console.log('模式：', APPLY ? '🔴 正式執行（會寫回正式庫）' : '🟢 試跑 dry-run（只看不寫）');

  if (RECAP) console.log('（--recap：只重補「已補全但資本額 0/空」的公司）');
  if (STATUS) console.log('（--status：回填「已補全但無公司狀態」的公司，補 GCIS 公司狀態）');
  if (RETRY_ERROR) console.log('（--retry-error：重試「GCIS查無」的公司，搭配證交所 fallback 補回上市公司）');
  await loadTwseMap();
  console.log('證交所上市對照：', Object.keys(twMap).length, '家');
  const data = await pg._loadAsync();
  const companies = data.companies || [];
  const selector = RECAP ? isRecapTarget : STATUS ? isStatusTarget : RETRY_ERROR ? isRetryErrorTarget : isTarget;
  const targets = companies.filter(selector);
  console.log(`\n企業主檔總數：${companies.length} 家`);
  if (RECAP) {
    console.log(`待重補資本額（已補全但資本額 0/空）：${targets.length} 家`);
  } else if (RETRY_ERROR) {
    console.log(`待重試（GCIS查無，將以證交所 fallback 補回上市公司）：${targets.length} 家`);
  } else if (STATUS) {
    console.log(`待回填公司狀態（已補全但無 gcisStatus）：${targets.length} 家`);
  } else {
    console.log(`待 GCIS 補全（有 8 碼統編、未補全/未標記）：${targets.length} 家`);
    console.log(`（沒統編/格式錯，自動跳過：${companies.filter(m => !has8(m)).length} 家）`);
  }

  const work = argLimit > 0 ? targets.slice(0, argLimit) : targets;

  if (!APPLY) {
    const sample = work.slice(0, DRY_SAMPLE);
    console.log(`\n── 試跑樣本（前 ${sample.length} 家實際查 GCIS，不寫入）──`);
    for (const m of sample) {
      const info = await resolveCompany(m.taxId);
      let verdict;
      if (RECAP) {
        verdict = (info === null) ? '逾時/錯誤（略過）'
          : info.notFound ? '查無（不動）'
          : (info.capital != null && info.capital !== m.capital) ? `✓ 資本額 ${m.capital ?? '—'} → ${info.capital}`
          : '無新資本額可補（GCIS 也是 0/空）';
      } else if (STATUS) {
        verdict = (info === null) ? '逾時/錯誤（略過）'
          : info.notFound ? '查無（不動）'
          : `✓ 狀態：${info.status || '（空）'}`;
      } else {
        verdict = (info === null) ? '逾時/錯誤（會留待重試）'
          : info.notFound ? '⚠️ GCIS查無（非公司登記/統編誤 → 跳過）'
          : `✓ 會補：「${info.name}」 資本額 ${info.capital ?? '—'}`;
      }
      console.log(`  ${m.taxId}  ${(m.name || '').slice(0, 16).padEnd(16)} → ${verdict}`);
      await sleep(200);
    }
    console.log(`\n🟢 這是試跑、未寫入任何資料。確認無誤後執行：`);
    console.log(`    node scripts/gcis-backfill-local.js --apply${RECAP ? ' --recap' : ''}${STATUS ? ' --status' : ''}${RETRY_ERROR ? ' --retry-error' : ''}${argLimit ? ' --limit ' + argLimit : ''}`);
    await pg._poolEnd?.();
    process.exit(0);
  }

  // ── 正式執行 ──
  const tally = {};
  let done = 0;
  const total = work.length;
  const ids = work.map(m => m.id);

  for (let i = 0; i < ids.length; i += BATCH) {
    const batchIds = ids.slice(i, i + BATCH);
    // 先（用目前 blob 的對應公司）查 GCIS
    const byId = new Map(companies.map(c => [c.id, c]));
    const infos = await Promise.allSettled(
      batchIds.map(async id => ({ id, info: await resolveCompany(byId.get(id)?.taxId) }))
    );

    // 重讀最新 blob，只把本批公司的 GCIS 欄位套上去再寫回（避免蓋掉他人即時編輯）
    const fresh = await pg._loadAsync();
    const freshById = new Map((fresh.companies || []).map(c => [c.id, c]));
    let batchTouched = false;
    for (const r of infos) {
      if (r.status !== 'fulfilled') { tally.failed = (tally.failed || 0) + 1; done++; continue; }
      const { id, info } = r.value;
      const fm = freshById.get(id);
      if (!fm) { done++; continue; }          // 公司可能已被刪 → 略過
      const res = RECAP ? applyRecap(fm, info) : STATUS ? applyStatus(fm, info) : applyInfo(fm, info);
      tally[res] = (tally[res] || 0) + 1;
      if (res === 'enriched' || res === 'taxIdError' || res === 'updated') batchTouched = true;
      done++;
    }
    if (batchTouched) await pg._saveAsync(fresh);

    const pct = Math.round(done / total * 100);
    if (RECAP) {
      process.stdout.write(`\r進度 ${done}/${total} (${pct}%)  ·  補回資本額 ${tally.updated || 0}  ·  無新值 ${tally.nochange || 0}  ·  查無/逾時 ${(tally.skipped || 0) + (tally.failed || 0)}   `);
    } else if (STATUS) {
      process.stdout.write(`\r進度 ${done}/${total} (${pct}%)  ·  補狀態 ${tally.updated || 0}  ·  無變更 ${tally.nochange || 0}  ·  查無/逾時 ${(tally.skipped || 0) + (tally.failed || 0)}   `);
    } else {
      process.stdout.write(`\r進度 ${done}/${total} (${pct}%)  ·  已補 ${tally.enriched || 0}  ·  GCIS查無 ${tally.taxIdError || 0}  ·  逾時 ${tally.failed || 0}   `);
    }
    await sleep(GAP_MS);
  }

  if (RECAP) {
    console.log(`\n\n✅ 完成（recap）：補回資本額 ${tally.updated || 0} 家、無新值可補 ${tally.nochange || 0} 家、查無/逾時 ${(tally.skipped || 0) + (tally.failed || 0)} 家。`);
  } else if (STATUS) {
    console.log(`\n\n✅ 完成（status）：補上公司狀態 ${tally.updated || 0} 家、無變更 ${tally.nochange || 0} 家、查無/逾時 ${(tally.skipped || 0) + (tally.failed || 0)} 家。`);
  } else {
    console.log(`\n\n✅ 完成：已補 ${tally.enriched || 0} 家、GCIS查無(跳過) ${tally.taxIdError || 0} 家、逾時待重試 ${tally.failed || 0} 家。`);
    if (tally.failed) console.log('   逾時的可再執行一次本腳本繼續補（已補/已標記的不會重複）。');
  }
  await pg._poolEnd?.();
  process.exit(0);
})().catch(async e => {
  console.error('\n✗ 發生錯誤：', e.message);
  process.exit(1);
});
