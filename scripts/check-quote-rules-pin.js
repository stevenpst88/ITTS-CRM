#!/usr/bin/env node
'use strict';
/**
 * 報價單簽核「規則內容釘選」檢查。用法：node scripts/check-quote-rules-pin.js（純函式，不開伺服器、不碰任何資料檔）
 *
 * 計算 sha256(canonical JSON of { ROWS, AMOUNT, PATHS })（核決表、金額門檻、核決路徑），同時比對：
 *   1) 本腳本內的釘選值 PINNED_RULES_HASH
 *   2) lib/quoteApproval.js 匯出的 RULES_HASH
 * 任一不符 → 失敗並提示「規則內容已變動，請更新 RULES_VERSION 與 RULES_HASH」。
 * 目的：改了門檻／路徑卻忘了升規則版本，會讓舊單 ap.rulesVersion 與送簽快照失真；這道檢查讓「規則變動」一定是有意識的、成對更新的。
 *
 * 規則真的要改時：1) 改規則 2) 更新 lib/quoteApproval.js 的 RULES_VERSION 3) 執行本腳本取得新雜湊（失敗訊息會印出）
 *               4) 同時更新 lib/quoteApproval.js 的 RULES_HASH 與本腳本的 PINNED_RULES_HASH 5) 重新確認 scripts/check-quote-*.js 全過。
 * 測試用：QRP_LIB=<路徑> 可指定要檢查的 quoteApproval 檔（變異檢查用）。
 */
const path = require('path');
const crypto = require('crypto');

const PINNED_RULES_HASH = '76efc01de2f3449f0371ed3217bc1e1469ad1282bfae8066bfacef605de76c16';

function canon(v) {
  if (v === null || typeof v !== 'object') return JSON.stringify(v);
  if (Array.isArray(v)) return '[' + v.map(canon).join(',') + ']';
  return '{' + Object.keys(v).sort().map((k) => JSON.stringify(k) + ':' + canon(v[k])).join(',') + '}';
}

const lib = process.env.QRP_LIB || path.join(__dirname, '..', 'lib', 'quoteApproval.js');
const QA = require(lib);
const actual = crypto.createHash('sha256').update(canon({ ROWS: QA.ROWS, AMOUNT: QA.AMOUNT, PATHS: QA.PATHS })).digest('hex');

const problems = [];
if (typeof QA.RULES_HASH !== 'string' || !QA.RULES_HASH) problems.push('lib/quoteApproval.js 沒有匯出 RULES_HASH');
if (actual !== PINNED_RULES_HASH) problems.push(`目前規則雜湊 ${actual} ≠ 腳本釘選值 ${PINNED_RULES_HASH}`);
if (QA.RULES_HASH !== actual) problems.push(`目前規則雜湊 ${actual} ≠ lib 匯出的 RULES_HASH ${QA.RULES_HASH}`);
if (typeof QA.RULES_VERSION !== 'string' || !QA.RULES_VERSION) problems.push('lib/quoteApproval.js 沒有 RULES_VERSION');

if (problems.length) {
  console.log('✗ 規則內容已變動，請更新 RULES_VERSION 與 RULES_HASH');
  problems.forEach((p) => console.log('  - ' + p));
  console.log(`  （目前 RULES_VERSION=${QA.RULES_VERSION}；新的雜湊應為 ${actual}）`);
  process.exit(1);
}
console.log(`✓ 規則內容與釘選一致（RULES_VERSION=${QA.RULES_VERSION}，RULES_HASH=${actual.slice(0, 12)}…）`);
