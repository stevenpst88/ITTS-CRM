#!/usr/bin/env node
/**
 * 環境變數文件檢查（M-06）。用法：node scripts/check-env-docs.js
 * 掃描 server.js、api、lib、middleware、db、db.js、storage、ai 實際讀取的環境變數，
 * 與 .env.example（以及 DEPLOY.md 內出現的名稱）比對：程式有讀、文件沒寫 → 失敗並列出。
 * 反向（.env.example 有寫、程式沒讀）只警告，不算失敗。
 *
 * 讀取寫法（程式裡環境變數不只有 process.env.X）：
 *   A process.env.X            B env.X／e.X／deps.env.X／opts.env.X（env 由參數傳入的 lib）
 *   C (env || process.env).X   D str(env,'X')／envNum('X',預設) 等小工具
 *   E ENV_NAME = 'X' 常數      F process.env['X']
 * 新增的寫法請補進 PATTERNS，並確認下方「掃描器健全性」的必抓名單仍成立。
 *
 * 環境變數：ENVDOC_EXAMPLE 指定要比對的範本檔（預設專案根 .env.example；變異驗證用）。
 */
const fs = require('fs');
const path = require('path');
const ROOT = path.join(__dirname, '..');
const EXAMPLE = process.env.ENVDOC_EXAMPLE || path.join(ROOT, '.env.example');

// 平台注入／行程內建，不需要在範本中設定（範本最後一段以註解說明）
const PLATFORM = new Set(['VERCEL', 'VERCEL_ENV', 'VERCEL_URL', 'VERCEL_REGION', 'VERCEL_PROJECT_PRODUCTION_URL']);
// 行程慣用名稱，不是本專案設定
const BUILTIN = new Set(['NODE_ENV']);   // NODE_ENV 範本也有寫，這裡只是說明「不寫也不算缺」

const SCAN = ['server.js', 'db.js', 'api', 'lib', 'middleware', 'db', 'storage', 'ai'];
const NAME = '[A-Z][A-Z0-9]*(?:_[A-Z0-9]+)*';
const PATTERNS = [
  new RegExp('process\\.env\\.(' + NAME + ')\\b', 'g'),                                           // A
  new RegExp('\\b(?:env|e|deps\\.env|opts\\.env)\\.(' + NAME + ')\\b', 'g'),                       // B
  new RegExp('\\|\\|\\s*process\\.env\\)\\.(' + NAME + ')\\b', 'g'),                               // C
  new RegExp('\\b(?:str|envNum|envInt|envBool)\\(\\s*(?:env|e)?\\s*,?\\s*\'(' + NAME + ')\'', 'g'), // D
  new RegExp('\\bENV_NAME\\s*=\\s*\'(' + NAME + ')\'', 'g'),                                     // E
  new RegExp('process\\.env\\[\\s*[\'"](' + NAME + ')[\'"]\\s*\\]', 'g'),                         // F
];
// B 會誤抓到「變數名剛好叫 env/e 的非環境變數大寫屬性」；目前專案內沒有，若日後誤報請加到這裡並寫明理由
const NOT_ENV = new Set([]);
// 掃描器健全性：這些一定要被掃到，否則代表 PATTERNS 腐爛（寫法改了、掃描器變瞎）
const MUST_FIND = ['FEATURE_GUARD_MODE', 'QUOTE_SPLIT_GUARD', 'PIPELINE_REPEAT_MIN_COUNT', 'MAIL_MODE', 'APPROVAL_CONFIG_DUAL_REVIEW',
  'LOGIN_LOCKOUT', 'AUDIT_CHAIN', 'CSP_REPORT_ONLY', 'QUOTE_FLOW_ENGINE', 'CRON_SECRET', 'SESSION_IDLE_MS', 'MAIL_GRAPH_TENANT_ID', 'DATABASE_URL'];

function walk(p, out) {
  const abs = path.join(ROOT, p);
  if (!fs.existsSync(abs)) return out;
  const st = fs.statSync(abs);
  if (st.isDirectory()) { for (const f of fs.readdirSync(abs)) { if (f === 'node_modules') continue; walk(path.join(p, f), out); } }
  else if (/\.(js|mjs|cjs)$/.test(p)) out.push(p);
  return out;
}
// 去掉註解再掃（註解裡的 process.env.X 說明文字不算「讀取」）；字串內容保留
function stripComments(src) {
  return src.replace(/\/\*[\s\S]*?\*\//g, (m) => m.replace(/[^\n]/g, ' ')).replace(/(^|[^:'"`\\])\/\/[^\n]*/g, '$1');
}

const used = new Map();   // name → [file…]
for (const f of SCAN.reduce((a, p) => walk(p, a), [])) {
  const src = stripComments(fs.readFileSync(path.join(ROOT, f), 'utf8'));
  for (const re of PATTERNS) {
    re.lastIndex = 0; let m;
    while ((m = re.exec(src))) { if (NOT_ENV.has(m[1])) continue; if (!used.has(m[1])) used.set(m[1], new Set()); used.get(m[1]).add(f.replace(/\\/g, '/')); }
  }
}

const exampleText = fs.existsSync(EXAMPLE) ? fs.readFileSync(EXAMPLE, 'utf8') : '';
const documented = new Set();
for (const line of exampleText.split(/\r?\n/)) { const m = /^\s*#?\s*([A-Z][A-Z0-9]*(?:_[A-Z0-9]+)*)\s*=/.exec(line); if (m) documented.add(m[1]); }
const deployPath = path.join(ROOT, 'DEPLOY.md');
const deployNames = new Set(fs.existsSync(deployPath) ? (fs.readFileSync(deployPath, 'utf8').match(new RegExp('\\b' + NAME + '\\b', 'g')) || []) : []);

let fail = 0;
const out = (ok, msg) => { if (!ok) fail++; console.log((ok ? '  ✓ ' : '  ✗ FAIL: ') + msg); };

out(fs.existsSync(EXAMPLE), '找得到範本檔 ' + path.relative(ROOT, EXAMPLE));
const missingScan = MUST_FIND.filter(n => !used.has(n));
out(missingScan.length === 0, '掃描器健全性：必抓的 ' + MUST_FIND.length + ' 個名稱都掃得到' + (missingScan.length ? '（漏抓：' + missingScan.join(', ') + '，請修 PATTERNS）' : ''));
out(used.size >= 40, '掃到的環境變數數量合理（' + used.size + ' 個，預期 ≥ 40）');

const missing = [...used.keys()].filter(n => !PLATFORM.has(n) && !BUILTIN.has(n) && !documented.has(n) && !deployNames.has(n)).sort();
out(missing.length === 0, '程式讀取的環境變數都有文件（.env.example／DEPLOY.md）' + (missing.length ? '\n' + missing.map(n => `      缺：${n}  （${[...used.get(n)].join(', ')}）`).join('\n') : ''));

const stale = [...documented].filter(n => !used.has(n) && !PLATFORM.has(n)).sort();
if (stale.length) console.log('  ⚠ 警告（不影響結果）：.env.example 有寫但程式沒讀取：' + stale.join(', '));

console.log(`\n環境變數文件檢查結果：${fail ? '失敗 ' + fail + ' 項' : '通過'}（程式讀取 ${used.size} 個，範本記載 ${documented.size} 個）`);
process.exit(fail ? 1 : 0);
