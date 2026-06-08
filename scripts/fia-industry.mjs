#!/usr/bin/env node
// ─────────────────────────────────────────────────────────────
// 財政部稅籍「官方行業」批次補全（本地執行）
//
// 流程：登入正式站(admin) → 取企業主檔統編清單 → 本地串流讀取財政部
//       320MB 稅籍 CSV → 比對統編取「主行業代號」→ 回寫正式站（後端
//       換算官方 19 大類）。整包 CSV 在本機處理，不經 Vercel。
//
// 用法（PowerShell）：
//   $env:CRM_URL="https://你的網址"; $env:CRM_USER="Admin"; $env:CRM_PASS="密碼"
//   node scripts/fia-industry.mjs
//
// 用法（bash）：
//   CRM_URL=https://你的網址 CRM_USER=Admin CRM_PASS=密碼 node scripts/fia-industry.mjs
//
// 選項：CSV_URL 可覆寫來源；CSV_FILE 可指定已下載好的本地 CSV（省下載）。
// ─────────────────────────────────────────────────────────────
import https from 'node:https';
import http from 'node:http';
import fs from 'node:fs';
import readline from 'node:readline';

const CRM_URL = (process.env.CRM_URL || '').replace(/\/$/, '');
const CRM_USER = process.env.CRM_USER || '';
const CRM_PASS = process.env.CRM_PASS || '';
const CSV_URL = process.env.CSV_URL || 'https://eip.fia.gov.tw/data/BGMOPEN1.csv';
const CSV_FILE = process.env.CSV_FILE || '';

if (!CRM_URL || !CRM_USER || !CRM_PASS) {
  console.error('❌ 請設定環境變數 CRM_URL / CRM_USER / CRM_PASS');
  process.exit(1);
}

// ── 簡易 fetch（保留 cookie）──
function req(method, url, { cookie, json } = {}) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const lib = u.protocol === 'http:' ? http : https;
    const body = json ? JSON.stringify(json) : null;
    const r = lib.request(u, {
      method,
      headers: {
        'Content-Type': 'application/json',
        ...(cookie ? { Cookie: cookie } : {}),
        ...(body ? { 'Content-Length': Buffer.byteLength(body) } : {}),
      },
    }, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => resolve({
        status: res.statusCode,
        cookie: (res.headers['set-cookie'] || []).map(c => c.split(';')[0]).join('; '),
        json: (() => { try { return JSON.parse(data); } catch { return null; } })(),
        raw: data,
      }));
    });
    r.on('error', reject);
    r.setTimeout(30000, () => { r.destroy(); reject(new Error('timeout')); });
    if (body) r.write(body);
    r.end();
  });
}

// ── 解析一列 CSV（處理雙引號含逗號）──
function parseCsvLine(line) {
  const out = []; let cur = ''; let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (inQ) {
      if (ch === '"') { if (line[i + 1] === '"') { cur += '"'; i++; } else inQ = false; }
      else cur += ch;
    } else {
      if (ch === '"') inQ = true;
      else if (ch === ',') { out.push(cur); cur = ''; }
      else cur += ch;
    }
  }
  out.push(cur);
  return out;
}

// ── CSV 來源：本地檔 or 串流下載 ──
function openCsvStream() {
  if (CSV_FILE) {
    console.log('📂 讀取本地 CSV：', CSV_FILE);
    return Promise.resolve(fs.createReadStream(CSV_FILE));
  }
  console.log('⬇  串流下載財政部 CSV（320MB，約 1-3 分鐘）：', CSV_URL);
  return new Promise((resolve, reject) => {
    https.get(CSV_URL, { headers: { 'User-Agent': 'Mozilla/5.0' } }, res => {
      if (res.statusCode !== 200) { reject(new Error('CSV HTTP ' + res.statusCode)); return; }
      resolve(res);
    }).on('error', reject);
  });
}

(async () => {
  // 1. 登入
  console.log('🔑 登入', CRM_URL, '...');
  const login = await req('POST', `${CRM_URL}/api/login`, { json: { username: CRM_USER, password: CRM_PASS } });
  if (login.status !== 200 || !login.json?.success) {
    console.error('❌ 登入失敗：', login.status, login.raw?.slice(0, 200)); process.exit(1);
  }
  const cookie = login.cookie;

  // 2. 取主檔統編清單
  const comp = await req('GET', `${CRM_URL}/api/admin/companies`, { cookie });
  if (comp.status !== 200 || !Array.isArray(comp.json)) {
    console.error('❌ 取企業主檔失敗：', comp.status, comp.raw?.slice(0, 200)); process.exit(1);
  }
  const wanted = new Set();
  comp.json.forEach(c => { const t = String(c.taxId || '').trim(); if (/^\d{8}$/.test(t)) wanted.add(t); });
  console.log(`🏢 主檔共 ${comp.json.length} 家，其中有 8 碼統編可比對：${wanted.size} 家`);
  if (!wanted.size) { console.log('沒有可比對的統編，結束。'); process.exit(0); }

  // 3. 串流 CSV、比對
  const stream = await openCsvStream();
  const rl = readline.createInterface({ input: stream, crlfDelay: Infinity });
  const found = new Map(); // taxId -> 主行業代號
  let lineNo = 0;
  for await (const line of rl) {
    lineNo++;
    if (lineNo === 1) continue;            // 標頭
    if (found.size === wanted.size) { rl.close(); break; } // 全找到，提早結束
    // 快篩：沒有任何 wanted 統編在這行就跳過 parse（加速）
    const fields = parseCsvLine(line);
    const tax = (fields[1] || '').trim();
    if (!tax || !wanted.has(tax) || found.has(tax)) continue;
    const code = (fields[8] || '').trim();  // 主行業代號
    if (code) found.set(tax, code);
    if (found.size % 10 === 0) process.stdout.write(`\r   已比對到 ${found.size}/${wanted.size} ...`);
  }
  console.log(`\n✅ CSV 掃描完成（${lineNo.toLocaleString()} 列），比對到 ${found.size} 家`);

  // 4. 回寫
  const items = [...found.entries()].map(([taxId, industryCode]) => ({ taxId, industryCode }));
  if (!items.length) { console.log('沒有可回寫的資料。'); process.exit(0); }
  const post = await req('POST', `${CRM_URL}/api/admin/companies/set-industry-batch`, { cookie, json: { items } });
  if (post.status !== 200) { console.error('❌ 回寫失敗：', post.status, post.raw?.slice(0, 200)); process.exit(1); }
  console.log('🎉 完成：', JSON.stringify(post.json));
  const missing = [...wanted].filter(t => !found.has(t));
  if (missing.length) console.log(`（${missing.length} 家統編在財政部稅籍查無，可能為外商/特殊機構：${missing.slice(0, 10).join(', ')}${missing.length > 10 ? '…' : ''}）`);
})().catch(e => { console.error('❌ 錯誤：', e.message); process.exit(1); });
