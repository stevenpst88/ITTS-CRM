#!/usr/bin/env node
'use strict';
/**
 * 儲存型 XSS 弱點（sink）防回歸檢查。不含任何帳密：測試帳號與隨機密碼只在記憶體裡。
 *
 *   node scripts/check-xss-sinks.js              # 全部
 *   XSS_ONLY=static,server,dom,red,mut node ...  # 只跑指定區段
 *   XSS_PRINT_ALLOW=1 node ...                   # 印出靜態掃描的候選項目（維護白名單用）
 *   XSS_SHOTS=<資料夾> node ...                  # 另外輸出 before/after 截圖
 *   XSS_SKIP_DOM=1                               # 無 Edge 時略過 dom/red/mut 的瀏覽器部分（會標「未驗證」）
 *
 * 區段：
 *   [A] static  靜態掃描：app.js／admin.html 的 HTML 模板裡，使用者可控欄位的 ${…} 必須經 escapeHtml/escH/escapeAttr/jsArg…
 *               已稽核過的常數／內部標記放 scripts/xss-sinks-allowlist.json（每筆要寫 why；新增前請先確認真的不是使用者資料）
 *   [B] server  cardImage 驗證矩陣、帳號／顯示名稱規則、違規清單端點（自行啟動 server.js，埠 XSS_PORT 預設 3221）
 *   [C] dom     以 headless Edge（CDP）載入真實頁面，餵 XSS payload 給各個渲染函式，斷言「沒有執行、沒有跳出屬性、顯示成文字」
 *   [D] red     紅燈證明：把頁面的 app.js／admin.html 換成修補前版本（釘死 commit 397abc5），同一組斷言必須失敗
 *   [E] mut     變異檢查（三條規則）：拿掉規則 → 對應測試必須失敗
 * 守門：工作樹若已有 auth.json / data.json / audit.log.json*（可能是真實資料）一律拒絕執行。
 * 禁用埠：3000 / 3020 / 3030 / 3101–3211（其他 session 使用）。只用 PID 關閉自己啟動的程序。
 */
const fs = require('fs');
const { assertPortFree } = require('./_port-free');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const PORT = parseInt(process.env.XSS_PORT || '3221', 10);
const CDP_PORT = parseInt(process.env.XSS_CDP || '9441', 10);
if ([3000, 3020, 3030].includes(PORT) || (PORT >= 3101 && PORT <= 3211)) { console.log('拒絕：此埠保留給其他 session'); process.exit(2); }
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));
const guardFiles = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
if (guardFiles.length) { console.log('拒絕：此工作樹已有', guardFiles.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
const ONLY = (process.env.XSS_ONLY || '').split(',').map(s => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length || ONLY.includes(k);
const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));

let pass = 0, fail = 0;
const failures = [], notes = [];
let quiet = false;
function ok(cond, name, extra) { if (cond) pass++; else { fail++; failures.push(name + (extra ? '（' + extra + '）' : '')); if (!quiet) console.log('  ✗ FAIL:', name, extra || ''); } }
const sleep = (n) => new Promise(r => setTimeout(r, n));
// 變異檢查用：執行 fn，回傳其中「失敗」的斷言數，並把計數還原（這些失敗是預期中的）
async function captureFailures(fn) { const s = { pass, fail, failures: failures.slice() }; quiet = true; try { await fn(); } finally { quiet = false; } const d = fail - s.fail; pass = s.pass; fail = s.fail; failures.length = 0; failures.push(...s.failures); return d; }
// 「修補前」基準釘死在 XSS 修補（08cb196）之前的 commit（其父提交，397abc5）。
// 原本用會移動的 origin/main：修補併入 main 之後「修補前」版本就不再有漏洞，13 項紅燈斷言全部失敗、npm test 永遠紅。
// 可用 XSS_BASE_REF 覆寫；取不到該 commit（淺層 clone）會直接報錯，不會靜默略過。
const XSS_BASE_COMMIT = '397abc5fbb4299976cb8889be3054f8dfa5c741f';
const XSS_BASE = process.env.XSS_BASE_REF || XSS_BASE_COMMIT;
const git = (args) => { const r = spawnSync('git', args, { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }); if (r.status !== 0) throw new Error('git ' + args.join(' ') + ' 失敗：' + r.stderr); return r.stdout; };

// ═════════════ [A] 靜態掃描 ═════════════
const RISKY = /\b(name|nameEn|company|companyName|product|topic|note|title|address|stage|category|description|contactName|visitType|email|phone|mobile|displayName|username|owner|ownerName|content|reason|deleteReason|invoiceNo|contractNo|salesPerson|interest|cardImage|opportunityStage|industry|taxId|representative|website|nextAction|currency|campaignName|expectedDate|startDate|endDate|dueDate|invoiceDate|visitDate|renewDate|budget)\b/;
const SAFE = /escapeHtml|escH\(|escapeAttr|jsArg|pcEsc|\besc\(|escapeXml|encodeURIComponent|toLocaleString|toFixed|Number\(|parseInt|parseFloat|Math\.|\.length\b|\.size\b|\.count\b/;
// 這些行的插值不是 HTML（toast／confirm／textContent／console），不算 sink
const NON_HTML_LINE = /showToast\(|confirm\(|\.textContent\s*=|alert\(|console\./;
function scanText(text) {
  const out = [];
  text.split('\n').forEach((l, i) => {
    if (!/[<>]/.test(l) || NON_HTML_LINE.test(l)) return;
    const re = /\$\{((?:[^{}]|\{(?:[^{}]|\{[^{}]*\})*\})*)\}/g;
    let m;
    while ((m = re.exec(l))) {
      const e = m[1].trim();
      if (SAFE.test(e)) continue;
      if (!/^[A-Za-z_$][\w$]*$/.test(e) && !/[A-Za-z_$][\w$]*(\.|\?\.)[A-Za-z_$]/.test(e)) continue;
      const props = (e.match(/(?:\.|\?\.)([A-Za-z_$][\w$]*)/g) || []).map(x => x.replace(/^\??\./, ''));
      if (!props.some(p => RISKY.test(p))) continue;
      out.push({ line: i + 1, expr: e, text: l.trim() });
    }
  });
  return out;
}
// 修補前出現過、不得再出現的具體寫法（含屬性情境：title／value／src）
const FORBID = [
  ['app.js', /src="\$\{c\.cardImage\}"/, '名片圖 src 未跳脫'],
  ['app.js', /<option value="\$\{name\}">/, '公司 datalist option 未跳脫'],
  ['app.js', /title="\$\{z\.product\|\|''\}"/, '殭屍商機 title 未跳脫'],
  ['app.js', /title="\$\{o\.product\|\|''\}"/, '商機表 title 未跳脫'],
  ['app.js', /title="\$\{o\.deleteReason\}"/, '已刪除商機原因 title 未跳脫'],
  ['app.js', /title="\$\{c\.note \|\| ''\}"/, '合約備註 title 未跳脫'],
  ['app.js', /<span class="cg-recent-text">\$\{recentVisit\.topic/, '最近拜訪主題未跳脫'],
  ['app.js', /\$\{recentOpp\.stage\}<\/span>/, '最近商機階段未跳脫'],
  ['app.js', /\$\{v\.topic\|\|''\}<\/span>/, '當日拜訪列表主題未跳脫'],
  ['app.js', /openMbBudgetModal\('\$\{/, 'inline onclick 以單引號包字串（應用 jsArg）'],
  ['app.js', /openMbActualModal\('\$\{/, 'inline onclick 以單引號包字串（應用 jsArg）'],
  ['app.js', /<option value="\$\{u\.username\}">/, '代理指派下拉 value 未跳脫'],
  ['admin.html', /<td style="padding:7px 12px[^"]*">\$\{r\.(name|company|title|phone|mobile|email|taxId) \|\| ''\}<\/td>/, 'OCR 預覽表未跳脫'],
  ['admin.html', /<option value="\$\{u\.username\}">/, '後台下拉 value 未跳脫'],
];
const ALLOW_FILE = path.join(__dirname, 'xss-sinks-allowlist.json');
function secStatic() {
  console.log('[A] 靜態掃描');
  const allow = JSON.parse(fs.readFileSync(ALLOW_FILE, 'utf8')).entries || [];
  allow.forEach((a, i) => ok(a.file && (a.expr || a.exprRe) && a.why && String(a.why).length >= 4, `白名單第 ${i + 1} 筆欄位完整（file／expr 或 exprRe／why）`));
  const used = new Set();
  const targets = [['app.js', '_client/app.js'], ['admin.html', '_client/admin.html']];
  for (const [label, rel] of targets) {
    const text = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const found = scanText(text);
    const bad = [];
    for (const f of found) {
      const idx = allow.findIndex(a => a.file === label && (a.expr === f.expr || (a.exprRe && new RegExp(a.exprRe).test(f.expr))) && f.text.includes(a.lineHas || ''));
      if (idx >= 0) used.add(idx); else bad.push(f);
    }
    if (process.env.XSS_PRINT_ALLOW) bad.forEach(f => console.log(JSON.stringify({ file: label, expr: f.expr, lineHas: f.text.slice(0, 60), why: '' })));
    ok(bad.length === 0, `${label}：HTML 模板中的使用者欄位插值都已跳脫或在白名單（未處理 ${bad.length} 筆）`, bad.slice(0, 5).map(b => `L${b.line} \${${b.expr}}`).join('；'));
    for (const [file, re, why] of FORBID) if (file === label) ok(!re.test(text), `${label}：不得再出現舊寫法 — ${why}`);
  }
  allow.forEach((a, i) => { if (!used.has(i) && !process.env.XSS_PRINT_ALLOW) ok(false, `白名單第 ${i + 1} 筆已無對應程式碼，請移除（${a.file} \${${a.expr}}）`); });
  // 掃描器自測：拿掉 escapeHtml 的樣本必須被抓到、跳脫過的不報
  ok(scanText('x = `<td title="${z.product||\'\'}">${z.stage}</td>`;').length === 2, '掃描器自測：未跳脫的 title／內文都被抓到');
  ok(scanText('x = `<td title="${escapeHtml(z.product)}">${escH(z.stage)}</td>`;').length === 0, '掃描器自測：已跳脫的不報');
  // server 端：使用者資料進 HTML 的寫入規則存在
  const srv = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  ok(/isValidCardImage\(req\.body\.cardImage\)/.test(srv) && /isValidCardImage\(safeBody\.cardImage\)/.test(srv), 'server.js：POST／PUT 聯絡人都驗證 cardImage');
  ok(/USERNAME_RE\.test\(String\(username\)\.trim\(\)\)/.test(srv), 'server.js：建立帳號驗證帳號字元');
}

// ═════════════ 沙盒資料（server／dom 共用）═════════════
const PW = 'Xs1' + crypto.randomBytes(9).toString('hex') + '!Z';
const BASE = `http://localhost:${PORT}`;
const nowIso = new Date().toISOString();
const mkUser = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: nowIso, passwordChangedAt: nowIso, mustChangePassword: false, canSetTargets: false, canDownloadContacts: false, bu: [] }, extra || {});
const LEGACY_USER = '_s_bad name';   // 舊帳號：帳號含空白、顯示名稱含 < >（新規則之前建立）
const LEGACY_IMG = 'legacy" onerror="window.__x=1" a="';
function seedFiles() {
  const users = [
    mkUser('_s_admin', 'admin'),
    mkUser('_s_u1', 'user', { bu: ['ERP'] }),
    mkUser('_s_u2', 'user', { bu: ['ERP'] }),
    mkUser(LEGACY_USER, 'user', { bu: ['ERP'], displayName: '<b>legacy</b>' }),
  ];
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users }, null, 2));
  const d = {
    contacts: [
      { id: 'L1', name: '舊資料', company: '沙盒公司甲', isForeign: true, taxId: '', owner: '_s_u1', bu: 'ERP', email: 'a@sandbox.test', phone: '0200000001', cardImage: LEGACY_IMG, createdAt: nowIso },
    ],
    companies: [], opportunities: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [], quotations: [], groups: [],
  };
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify(d, null, 2));
}
const uploadsExisted = fs.existsSync(path.join(ROOT, 'uploads'));
function cleanupFiles() {
  FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } });
  if (!uploadsExisted) { try { fs.rmdirSync(path.join(ROOT, 'uploads')); } catch { /* 非空或不存在 */ } }
}

// ═════════════ 伺服器 ═════════════
let child = null, childPid = null;
const tmpFiles = [];
async function startServer(serverFile) {
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  delete env.DATABASE_URL; delete env.MAIL_MODE; delete env.SUPABASE_URL; delete env.SUPABASE_SERVICE_ROLE_KEY; env.STORAGE_BACKEND = 'local';
  await assertPortFree(PORT);   // B-03：埠被別人佔用時明確報錯，不誤連別人的 server
  child = spawn(process.execPath, [serverFile || path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
  childPid = child.pid;
  let out = '';
  const onData = d => { out += d; if (process.env.XSS_DEBUG) process.stdout.write('[srv] ' + d); };
  child.stdout.on('data', onData); child.stderr.on('data', onData);
  for (let i = 0; i < 240; i++) {   // 本機冷啟動可能要 10~20 秒
    await sleep(250);
    try { const r = await fetch(BASE + '/login.html'); if (r.ok) return; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-500));
  }
  throw new Error('伺服器啟動逾時：' + out.slice(-500));
}
async function stopServer() {
  if (!child) return;
  try { process.kill(childPid); } catch { /* 已結束 */ }
  for (let i = 0; i < 20 && child.exitCode === null; i++) await sleep(100);
  child = null;
}
async function freshServer(serverFile) { await stopServer(); cleanupFiles(); seedFiles(); await startServer(serverFile); jar = {}; }

let jar = {};
async function http(who, method, url, body) {
  const h = { 'Content-Type': 'application/json' };
  if (who && jar[who]) h.Cookie = jar[who];
  const res = await fetch(BASE + url, { method, headers: h, body: body === undefined ? undefined : JSON.stringify(body), redirect: 'manual' });
  const buf = Buffer.from(await res.arrayBuffer());
  let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch { /* 非 JSON */ }
  return { status: res.status, json, text: buf.toString('utf8'), setCookie: res.headers.getSetCookie ? res.headers.getSetCookie() : [] };
}
async function login(who, user) {
  const r = await http(null, 'POST', '/api/login', { username: user || who, password: PW });
  if (r.status === 200) jar[who] = r.setCookie.map(c => c.split(';')[0]).join('; ');
  return r;
}

// ═════════════ [B] 伺服器端規則 ═════════════
const uuidUrl = (ext) => `/uploads/${crypto.randomUUID()}.${ext || 'png'}`;
async function secServer(tag) {
  const T = (tag || '') ? `[${tag}] ` : '';
  console.log(T + '[B] 伺服器端規則');
  await login('_s_u1'); await login('_s_admin'); await login('_s_u2');
  let n = 0;
  const base = (extra) => Object.assign({ name: '測試' + (++n), company: '沙盒公司乙', isForeign: true, bu: 'ERP' }, extra);
  const post = (body) => http('_s_u1', 'POST', '/api/contacts', body);
  // ── cardImage 驗證矩陣（新增）──
  let r = await post(base({ cardImage: uuidUrl('png') }));
  ok(r.status === 200 || r.status === 201, T + 'cardImage：本系統上傳網址（uuid.png）→ 接受', r.status + ' ' + (r.json && r.json.error));
  const okId = r.json && r.json.id;
  r = await post(base({ cardImage: uuidUrl('JPG') })); ok(r.status < 300, T + 'cardImage：uuid.JPG（本地儲存保留原副檔名大小寫）→ 接受', r.status);
  r = await post(base({ cardImage: '' })); ok(r.status < 300, T + 'cardImage：空字串 → 接受', r.status);
  r = await post(base({ cardImage: null })); ok(r.status < 300, T + 'cardImage：null → 接受', r.status);
  r = await post(base({})); ok(r.status < 300, T + 'cardImage：未帶 → 接受', r.status);
  const bad = {
    'javascript: 協定': 'javascript:alert(1)',
    'data: URL（程式庫不產生，不放行）': 'data:image/png;base64,iVBORw0KGgo=',
    '外部網址 http://evil': 'http://evil.test/x.png',
    '外部網址 https://evil': 'https://evil.test/uploads/' + crypto.randomUUID() + '.png',
    '跳出屬性 " onerror=': 'x" onerror="window.__x=1',
    '路徑穿越': '/uploads/../../etc/passwd',
    '上傳路徑但 uuid 長度不對': '/uploads/abc.png',
    '副檔名含符號': `/uploads/${crypto.randomUUID()}.p"ng`,
    '上傳網址後面夾帶屬性': `/uploads/${crypto.randomUUID()}.png" onerror="1`,
    '超長字串': '/uploads/' + 'a'.repeat(5000) + '.png',
    '非字串（物件）': { a: 1 },
  };
  for (const [name, v] of Object.entries(bad)) {
    r = await post(base({ cardImage: v }));
    ok(r.status === 400 && /名片圖片/.test((r.json && r.json.error) || ''), T + `cardImage 新增：${name} → 400 中文訊息`, r.status + ' ' + JSON.stringify(r.json).slice(0, 60));
  }
  const list = await http('_s_u1', 'GET', '/api/contacts');
  const stored = (list.json || []).filter(c => c.cardImage && c.cardImage !== LEGACY_IMG);
  ok(stored.every(c => /^\/uploads\/[0-9a-f-]{36}\.[A-Za-z0-9]+$/.test(c.cardImage)), T + '被拒絕的值一個都沒有寫進資料', stored.map(c => c.cardImage).join('|').slice(0, 80));
  // ── 更新 ──
  const put = (id, body) => http('_s_u1', 'PUT', '/api/contacts/' + id, body);
  r = await put('L1', { title: '經理' }); ok(r.status === 200, T + '舊資料（cardImage 不合規）只改職稱、不帶 cardImage → 200', r.status + ' ' + JSON.stringify(r.json).slice(0, 80));
  r = await put('L1', { title: '協理', cardImage: LEGACY_IMG }); ok(r.status === 200, T + '舊資料整包送回、cardImage 未變 → 200（不擋）', r.status + ' ' + JSON.stringify(r.json).slice(0, 80));
  r = await put('L1', { cardImage: 'javascript:alert(1)' }); ok(r.status === 400, T + '舊資料把 cardImage 改成另一個不合規值 → 400', r.status);
  r = await put('L1', { cardImage: 'data:image/png;base64,AAAA' }); ok(r.status === 400, T + '更新為 data: URL → 400', r.status);
  r = await put('L1', { cardImage: uuidUrl('webp') }); ok(r.status === 200, T + '舊資料改成合法上傳網址 → 200', r.status);
  r = await put('L1', { cardImage: '' }); ok(r.status === 200, T + '清空 cardImage → 200', r.status);
  r = await put(okId, { cardImage: 'http://evil.test/a.png' }); ok(r.status === 400, T + '合法資料改成外部網址 → 400', r.status);
  // ── 匯入：匯入流程本來就不吃 cardImage（一律寫 ''），確認帶了也不會寫入 ──
  r = await http('_s_admin', 'POST', '/api/admin/import-contacts-json', { targetOwner: '_s_u2', contacts: [{ name: '匯入甲', company: '匯入公司', taxId: '', cardImage: 'javascript:alert(1)' }] });
  const listU2 = await http('_s_u2', 'GET', '/api/contacts');
  const imp = (listU2.json || []).filter(c => c.name === '匯入甲');
  ok(imp.every(c => !c.cardImage), T + '匯入 JSON 帶不合規 cardImage → 不會寫入（匯入路徑一律寫 \'\'）', r.status + ' ' + JSON.stringify(imp.map(c => c.cardImage)));

  // ── 帳號字元規則（只限新建）──
  const mk = (username, extra) => http('_s_admin', 'POST', '/api/admin/users', Object.assign({ username, password: PW, displayName: username, role: 'user', bu: ['ERP'] }, extra || {}));
  for (const good of ['_s_ok1', 'Admin2', 'a.b_c-d', '_pool2', '_s_' + 'x'.repeat(46)]) {
    r = await mk(good); ok(r.status === 200 || r.status === 201, T + `新帳號 ${good.slice(0, 16)} → 接受`, r.status + ' ' + JSON.stringify(r.json).slice(0, 80));
  }
  const badNames = { '含空白': 'a b', '含 <': 'a<b', '含引號': 'a"b', '含單引號': "a'b", '中文': '帳號甲', '含 @': 'x@y.com', '51 字': 'a'.repeat(51), '含分號': 'a;b', '含斜線': 'a/b' };
  for (const [name, v] of Object.entries(badNames)) {
    r = await mk(v); ok(r.status === 400 && /帳號/.test((r.json && r.json.error) || ''), T + `新帳號 ${name} → 400`, r.status + ' ' + JSON.stringify(r.json).slice(0, 80));
  }
  r = await mk('_s_dn1', { displayName: '<img src=x onerror=1>王小明 (ERP) 😀\u0007' });
  ok(r.status < 300, T + '新帳號顯示名稱含 < > 與控制字元 → 建立成功（去除而非拒絕）', r.status);
  let users = (await http('_s_admin', 'GET', '/api/admin/users')).json || [];
  const dn1 = users.find(u => u.username === '_s_dn1');
  ok(dn1 && !/[<>\u0000-\u001f]/.test(dn1.displayName) && /王小明 \(ERP\) 😀/.test(dn1.displayName), T + '顯示名稱已去除 < > 與控制字元，保留中文／空白／括號／emoji', dn1 && JSON.stringify(dn1.displayName));
  r = await mk('_s_dn2', { displayName: '<>' }); users = (await http('_s_admin', 'GET', '/api/admin/users')).json || [];
  ok(r.status < 300 && (users.find(u => u.username === '_s_dn2') || {}).displayName === '_s_dn2', T + '顯示名稱全被去除 → 退回帳號名稱', r.status);
  r = await http('_s_admin', 'PUT', '/api/admin/users/_s_u2', { displayName: '改名<script>x</script>' });
  users = (await http('_s_admin', 'GET', '/api/admin/users')).json || [];
  ok(r.status === 200 && !/[<>]/.test((users.find(u => u.username === '_s_u2') || {}).displayName || '<'), T + '修改顯示名稱含 < > → 去除', r.status + ' ' + JSON.stringify((users.find(u => u.username === '_s_u2') || {}).displayName));
  r = await http('_s_admin', 'PUT', '/api/admin/users/' + encodeURIComponent(LEGACY_USER), { displayName: '<b>legacy</b>', canSetTargets: false });
  users = (await http('_s_admin', 'GET', '/api/admin/users')).json || [];
  ok(r.status === 200 && (users.find(u => u.username === LEGACY_USER) || {}).displayName === '<b>legacy</b>', T + '舊帳號顯示名稱未變動 → 原樣保留、不被改寫（只對新值套規則）', r.status);
  // ── 既有帳號不受影響 ──
  r = await login('legacy', LEGACY_USER); ok(r.status === 200, T + '舊帳號（空白／< >）仍可登入（規則不套用在登入或讀取）', r.status);
  // ── 違規清單端點 ──
  r = await http('_s_admin', 'GET', '/api/admin/users/name-violations');
  ok(r.status === 200 && r.json && r.json.usernameViolations.includes(LEGACY_USER) && r.json.displayNameViolations.includes(LEGACY_USER), T + '違規清單列出舊帳號（帳號與顯示名稱兩類）', r.status + ' ' + JSON.stringify(r.json).slice(0, 160));
  ok(r.json && !r.json.usernameViolations.includes('_s_admin') && !r.json.displayNameViolations.includes('_s_dn1') && !r.json.usernameViolations.includes('_s_ok1'), T + '違規清單不含合規帳號');
  ok(r.json && r.json.count === new Set([...r.json.usernameViolations, ...r.json.displayNameViolations, ...r.json.nicknameViolations]).size && !JSON.stringify(r.json).includes('legacy</b>'), T + '違規清單只有帳號代碼與計數、不回顯名稱內容');
  const before = JSON.stringify((await http('_s_admin', 'GET', '/api/admin/users')).json);
  await http('_s_admin', 'GET', '/api/admin/users/name-violations');
  ok(before === JSON.stringify((await http('_s_admin', 'GET', '/api/admin/users')).json), T + '違規清單端點是唯讀（不改任何資料）');
  r = await http('_s_u1', 'GET', '/api/admin/users/name-violations'); ok(r.status === 403 || r.status === 401, T + '非管理員 → 403', r.status);
  r = await http(null, 'GET', '/api/admin/users/name-violations'); ok(r.status === 401 || r.status === 403 || r.status === 302, T + '未登入 → 拒絕', r.status);
}

// ═════════════ [C]/[D]/[E] 瀏覽器 ═════════════
function findEdge() {
  return ['C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', 'C:/Program Files/Microsoft/Edge/Application/msedge.exe'].find(p => fs.existsSync(p));
}
const PAY = {
  img: '<img src=x onerror=window.__x=1>',
  brk: '"><svg/onload=window.__x=1>',
  js: "';window.__x=1;//",
  card: 'x" onerror="window.__x=1" a="',
};
const MARK = "<img src=x onerror=\"window.__x=1;this.replaceWith(Object.assign(document.createElement('b'),{textContent:'⚠ XSS 已執行',style:'color:#fff;background:#d32f2f;padding:2px 6px;border-radius:4px'}))\">";

/* ────────── 在頁面內執行的測試案例（以 toString() 注入；只能用頁面全域）────────── */
async function pageAppCases(P) {
  const out = [];
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const $ = (id) => document.getElementById(id);
  const host = document.createElement('div'); host.id = '__xssHost'; host.style.cssText = 'position:fixed;left:0;top:0;width:1px;height:1px;overflow:hidden';
  document.body.appendChild(host);
  const DANGER = '[onerror],[onload],[onfocus],[onmouseover],[autofocus],img[src="x"],svg[onload],script';
  // 記錄是誰把 __x 設成 1（事件目標與當時的案例），失敗時附在訊息裡，方便判斷是本案例還是晚到的舊事件
  let curCase = '', xVal = 0; const xWho = [];
  Object.defineProperty(window, '__x', { configurable: true, get() { return xVal; }, set(v) { xVal = v; if (v) xWho.push((curCase || '?') + ' <- ' + String((window.event && window.event.target && window.event.target.outerHTML) || 'no-event').slice(0, 90)); } });
  // root：渲染結果所在元素。expect：必須以「文字」出現的字串。attrs：[元素選擇器, 屬性, 預期值]（屬性值必須原樣保存且沒有被破壞）
  async function check(name, root, expect, attrs) {
    window.__x = 0; xWho.length = 0; curCase = name;
    let why = '';
    try {
      if (!root.isConnected) host.appendChild(root);
      await sleep(250);
      // 等所有圖片載入／失敗事件跑完（onerror 在請求失敗後才觸發），避免晚到的事件汙染下一個案例
      await Promise.all([...root.querySelectorAll('img')].map(i => i.complete ? 0 : new Promise(r => { i.addEventListener('load', r); i.addEventListener('error', r); setTimeout(r, 3000); })));
      await sleep(150);
      if (window.__x) why += '[腳本已執行 ' + xWho.join(' ; ') + '] ';
      const bad = root.matches && root.matches(DANGER) ? [root] : [];
      root.querySelectorAll(DANGER).forEach(e => bad.push(e));
      if (bad.length) why += '[出現危險元素/屬性 ' + bad.slice(0, 3).map(e => e.outerHTML.slice(0, 50)).join(' | ') + '] ';
      const txt = root.textContent || '';
      for (const s of expect || []) if (!txt.includes(s)) why += '[文字未原樣顯示: ' + s.slice(0, 24) + '] ';
      for (const [sel, attr, val] of attrs || []) {
        const el = root.matches(sel) ? root : root.querySelector(sel);
        if (!el) why += '[找不到 ' + sel + '] '; else if (el.getAttribute(attr) !== val) why += '[屬性 ' + sel + '@' + attr + ' 被改變: ' + String(el.getAttribute(attr)).slice(0, 40) + '] ';
      }
    } catch (e) { why += '[例外 ' + (e && e.message) + ']'; }
    out.push({ name, ok: !why, why });
  }
  const run = async (name, fn) => { try { try { sessionStorage.setItem('__xlast', name); } catch (e) { /* 無 storage */ } await fn(); } catch (e) { out.push({ name, ok: false, why: '[渲染函式例外 ' + (e && e.message) + ']' }); } };
  const today = new Date(); const ds = (d) => d.toISOString().slice(0, 10);
  const soon = ds(new Date(Date.now() + 30 * 86400000));
  const created = new Date().toISOString();

  await run('app:buildContactCard', async () => {
    const c = { id: 'p1', name: P.img, company: P.brk, title: P.img, phone: P.img, email: P.brk, customerType: 'customer', opportunityStage: P.img, cardImage: P.card, owner: 'x', ownerDisplayName: P.img, createdAt: created };
    const el = buildContactCard(c);
    await check('app:buildContactCard', el, [P.img, P.brk], [['img.card-image-thumb', 'src', P.card]]);
  });
  await run('app:openView', async () => {
    allContacts = [{ id: 'p2', name: P.img, company: P.brk, taxId: P.img, title: P.img, industry: P.img, opportunityStage: P.img, cardImage: P.card, phone: P.img, mobile: '', email: P.brk, address: P.img, note: P.img, owner: 'x', createdAt: created }];
    openView('p2');
    await check('app:openView', $('viewBody'), [P.img], [['img.view-card-image', 'src', P.card]]);
  });
  await run('app:updateCompanyDatalist', async () => {
    updateCompanyDatalist([{ company: P.brk, createdAt: created }, { company: P.img, createdAt: created }]);
    const dl = $('companyList');
    await check('app:updateCompanyDatalist', dl, [], []);
    const vals = [...dl.querySelectorAll('option')].map(o => o.getAttribute('value'));
    out.push({ name: 'app:updateCompanyDatalist(option value 原樣保存)', ok: vals.length === 2 && vals.includes(P.brk) && vals.includes(P.img), why: JSON.stringify(vals).slice(0, 120) });
  });
  await run('app:makeCell+openAllVisitsOfDay', async () => {
    const d = new Date(); const dateStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    allContacts = [{ id: 'p3', name: P.img, company: P.brk, owner: 'x', createdAt: created }];
    const v = { id: 'v1', contactId: 'p3', contactName: P.img, visitDate: dateStr, visitType: P.img, topic: P.brk };
    const cell = makeCell(d, { [dateStr]: [v] }, d, false);
    await check('app:makeCell', cell, [P.img, P.brk], []);
    openAllVisitsOfDay(dateStr, [v]);
    const ov = [...document.querySelectorAll('.modal-overlay')].pop();
    await check('app:openAllVisitsOfDay', ov, [P.img, P.brk], []);
    ov.remove();
  });
  await run('app:buildCompanyGroupEl', async () => {
    allContacts = [{ id: 'p4', name: P.img, company: P.brk, title: P.img, opportunityStage: P.img, industry: P.img, owner: 'x', createdAt: created }];
    allVisits = [{ id: 'v2', contactId: 'p4', visitDate: P.img, topic: P.img, visitType: '親訪' }];
    allOpportunities = [{ id: 'o9', contactId: 'p4', company: P.brk, stage: P.img, product: P.img, category: P.img, amount: 1, createdAt: created }];
    const el = buildCompanyGroupEl(P.brk, allContacts);
    await check('app:buildCompanyGroupEl', el, [P.img], []);
  });
  await run('app:renderOppTable', async () => {
    $('oppFilterYear').value = ''; $('oppFilterStage').value = '';
    allOpportunities = [{ id: 'o1', company: P.img, contactName: P.brk, category: P.img, product: P.brk, stage: 'B', amount: 5, expectedDate: P.img, createdAt: created }];
    allLostOppsForTargets = [{ id: 'o2', company: P.img, contactName: P.brk, category: P.img, product: P.brk, deleteReason: P.brk, stage: 'B', amount: 5, expectedDate: P.img, createdAt: created, deletedAt: created }];
    allZombieOpps = [];
    renderOppTable();
    const tb = $('oppTableBody');
    await check('app:renderOppTable', tb, [P.img, P.brk], [['td[data-label="商品項目"]', 'title', P.brk]]);
    out.push({ name: 'app:renderOppTable(已刪除列 title/原因原樣保存)', ok: [...tb.querySelectorAll('.opp-lost-reason')].every(e => e.getAttribute('title') === P.brk) && tb.querySelectorAll('.opp-lost-reason').length === 1, why: '' });
  });
  await run('app:renderZombieSection', async () => {
    allZombieOpps = [{ id: 'z1', severity: 'danger', company: P.img, contactName: P.brk, product: P.brk, stage: P.img, amount: 1, lastVisit: P.img, daysSinceAny: 3, reasons: [P.img], ownerName: P.img, owner: 'x' }];
    renderZombieSection();
    await check('app:renderZombieSection', $('zombieTableBody'), [P.img, P.brk], [['td[title]', 'title', P.brk]]);
  });
  await run('app:renderSapNoticeBar', async () => {
    if (!$('sapNoticeBar')) { const b = document.createElement('div'); b.id = 'sapNoticeBar'; document.body.appendChild(b); }
    renderSapNoticeBar([{ id: 'k1', company: P.img, product: 'PCOE License MA ' + P.img, endDate: soon, contractNo: P.brk }]);
    await check('app:renderSapNoticeBar', $('sapNoticeBar'), [P.img, P.brk], []);
  });
  await run('app:renderContractTable', async () => {
    const tb = document.createElement('tbody'); tb.id = '__xssCt'; const t = document.createElement('table'); t.appendChild(tb); host.appendChild(t);
    renderContractTable([{ id: 'k2', company: P.img, contactName: P.brk, contractNo: P.img, product: 'PCOE License MA ' + P.img, startDate: '2026-01-01', endDate: soon, salesPerson: P.brk, note: P.brk, amount: 1 }], '__xssCt');
    await check('app:renderContractTable', tb, [P.img, P.brk], [['td[title]', 'title', P.brk]]);
  });
  await run('app:renderReceivables', async () => {
    allReceivables = [{ id: 'r1', company: P.img, contactName: P.brk, invoiceNo: P.img, invoiceDate: P.img, dueDate: P.brk, amount: 100, paidAmount: 0, currency: P.img }];
    currentRtab = 'all';
    renderReceivables();
    await check('app:renderReceivables', $('receivablesTbody'), [P.img, P.brk], []);
  });
  await run('app:renderCompanyMasterDetail', async () => {
    const d = { company: { name: P.img, taxId: P.img, industry: P.brk, representative: P.img, address: P.brk, website: '', capital: 1 }, contacts: [], opps: [{ product: P.img, stage: P.img, amount: 1 }], contracts: [], receivables: [], yoy: null };
    const el = document.createElement('div'); el.innerHTML = renderCompanyMasterDetail(d);
    await check('app:renderCompanyMasterDetail', el, [P.img, P.brk], []);
  });
  await run('app:renderCompanyInfo', async () => {
    if (!$('companyInfoResult')) { const b = document.createElement('div'); b.id = 'companyInfoResult'; document.body.appendChild(b); }
    renderCompanyInfo({ companyName: P.img, representative: P.brk, companyStatus: P.img, capital: P.img, address: P.brk, stockCode: '1234', listedType: P.img, exchange: P.brk, grossMargin2025: 'x', grossMargin2024: 'x' });
    await check('app:renderCompanyInfo', $('companyInfoResult'), [P.img, P.brk], []);
  });
  await run('app:renderOpportunityBadge', async () => {
    const el = document.createElement('div'); el.innerHTML = renderOpportunityBadge(P.img);
    await check('app:renderOpportunityBadge', el, [P.img], []);
  });
  await run('app:renderCampaignCards', async () => {
    if (!$('campaignCards')) { const b = document.createElement('div'); b.id = 'campaignCards'; document.body.appendChild(b); }
    allCampaigns = [{ id: 'cp1', name: 'n', type: P.img, status: P.brk, startDate: P.img, endDate: P.brk, budget: P.img, leadCount: 1, targetCount: 2, convertedCount: 0 }];
    renderCampaignCards();
    await check('app:renderCampaignCards', $('campaignCards'), [P.img], []);
  });
  await run('app:renderLeadsTable', async () => {
    if (!$('leadsTbody')) { const t = document.createElement('table'); t.innerHTML = '<tbody id="leadsTbody"></tbody>'; document.body.appendChild(t); }
    allLeads = [{ id: 'ld1', company: 'c', contactName: 'n', status: P.img, createdAt: created }];
    renderLeadsTable();
    await check('app:renderLeadsTable', $('leadsTbody'), [P.img], []);
  });
  await run('app:renderMgrCommit', async () => {
    if (!$('mgrCommitContent')) { const b = document.createElement('div'); b.id = 'mgrCommitContent'; document.body.appendChild(b); }
    const g = { items: [{ stage: P.img, company: 'c', product: 'p', amount: 1 }], count: 1, amount: 1 };
    renderMgrCommit({ confirmed: g, atRisk: g, uncertain: g });
    await check('app:renderMgrCommit', $('mgrCommitContent'), [P.img], []);
  });
  await run('app:checkCompanySimilar', async () => {
    const real = window.fetch;
    window.fetch = (u, o) => String(u).indexOf('/api/company-similar') === 0 ? Promise.resolve(new Response(JSON.stringify({ matches: [{ name: P.img, taxId: P.brk }] }), { status: 200, headers: { 'Content-Type': 'application/json' } })) : real(u, o);
    if (!$('companySimDup')) { const b = document.createElement('div'); b.id = 'companySimDup'; document.body.appendChild(b); }
    const cEl = $('company'), tEl = $('taxId'); const oc = cEl.value, ot = tEl.value; cEl.value = 'x'; tEl.value = '';
    try { await checkCompanySimilar(); } finally { window.fetch = real; cEl.value = oc; tEl.value = ot; }
    await check('app:checkCompanySimilar', $('companySimDup'), [P.img], []);
  });
  await run('app:loadManagerAchievement(inline onclick)', async () => {
    window._myRole = 'admin';
    const real = window.fetch;
    const row = { rowKey: 'k1', parentRowKey: null, username: P.js, displayName: P.img, role: 'user', rate: 50, target: 100, achieved: 10, pipeline: 5, wonCount: 1, qVals: [1, 2, 3, 4], viewMode: 'self', children: [] };
    window.fetch = (u, o) => String(u).indexOf('/api/manager/achievement') === 0 ? Promise.resolve(new Response(JSON.stringify({ rows: [row] }), { status: 200, headers: { 'Content-Type': 'application/json' } })) : real(u, o);
    const sec = $('managerAchSection'); if (sec && !$('managerAchTable')) { const b = document.createElement('div'); b.id = 'managerAchTable'; sec.appendChild(b); }
    let got = null; const realOpen = window.openMbBudgetModal; window.openMbBudgetModal = function (u, y) { got = [u, y]; };
    try { await loadManagerAchievement(2026); } finally { window.fetch = real; }
    const tb = $('managerAchTable');
    await check('app:loadManagerAchievement', tb, [P.img], [['.qr-row', 'data-user', P.js]]);
    const btn = tb.querySelector('button[onclick*="openMbBudgetModal"]');
    window.__x = 0; if (btn) btn.click(); await sleep(100);
    window.openMbBudgetModal = realOpen;
    out.push({ name: 'app:inline onclick 以 JSON 編碼傳字串（payload `\';…//` 原樣到達函式、不執行）', ok: !!btn && got && got[0] === P.js && !window.__x, why: JSON.stringify(got) + ' x=' + window.__x });
  });
  host.remove();
  return out;
}

async function pageAdminCases(P) {
  const out = [];
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const $ = (id) => document.getElementById(id);
  const DANGER = '[onerror],[onload],[onfocus],[onmouseover],[autofocus],img[src="x"],svg[onload]';
  window.__x = 0;
  let why = '';
  try {
    const rows = [{ name: P.img, company: P.brk, title: P.img, phone: P.img, mobile: P.brk, email: P.img, taxId: P.brk }];
    $('ocrJsonInput').value = JSON.stringify(rows);
    $('ocrPreviewBtn').click();
    await sleep(300);
    const tb = $('ocrPreviewBody');
    if (window.__x) why += '[腳本已執行] ';
    const bad = tb.querySelectorAll(DANGER);
    if (bad.length) why += '[出現危險元素 ' + [...bad].slice(0, 2).map(e => e.outerHTML.slice(0, 40)).join('|') + '] ';
    if (!(tb.textContent || '').includes(P.img)) why += '[文字未原樣顯示] ';
  } catch (e) { why += '[例外 ' + (e && e.message) + ']'; }
  out.push({ name: 'admin:OCR 預覽表', ok: !why, why });
  // 錯誤訊息：adminFetch 失敗訊息放進 innerHTML 的位置（以 escH 實作；直接驗 escH／pcEsc 行為）
  const esc = [['escH', typeof escH === 'function' ? escH(P.img + P.brk) : null], ['pcEsc', typeof pcEsc === 'function' ? pcEsc("'" + P.js) : null]];
  out.push({ name: 'admin:escH 跳脫 < > "', ok: esc[0][1] === '&lt;img src=x onerror=window.__x=1&gt;&quot;&gt;&lt;svg/onload=window.__x=1&gt;', why: String(esc[0][1]) });
  return out;
}

/* ────────── Edge / CDP 驅動 ────────── */
const edgePids = [];
// Edge 會再 fork／重新啟動，ep.pid 殺掉不等於整棵樹都停：以「本次臨時 profile 路徑」找出屬於我們的程序 PID，逐一用 PID 結束（不按名稱殺）
function killEdgeByProfile(prof) {
  for (let round = 0; round < 5; round++) {
    const r = spawnSync('powershell', ['-NoProfile', '-Command', `(Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like '*${path.basename(prof)}*' }).ProcessId`], { encoding: 'utf8' });
    const pids = String(r.stdout || '').split(/\s+/).filter(x => /^\d+$/.test(x));
    if (!pids.length) return;
    pids.forEach(pid => { try { spawnSync('taskkill', ['/PID', pid, '/F'], { stdio: 'ignore' }); } catch { /* 已結束 */ } });
    spawnSync('powershell', ['-NoProfile', '-Command', 'Start-Sleep -Milliseconds 700']);
  }
}
async function withBrowser(fn) {
  const edge = findEdge();
  if (process.env.XSS_SKIP_DOM) { notes.push('瀏覽器測試略過（XSS_SKIP_DOM 明確略過）：未驗證'); return null; }
  if (!edge) { ok(false, '瀏覽器測試需要 Edge：找不到 msedge.exe（沒有 Edge 的環境請設 XSS_SKIP_DOM=1 明確略過；不可靜默當成通過）'); return null; }
  const prof = fs.mkdtempSync(path.join(os.tmpdir(), 'xss-edge-'));
  const ep = spawn(edge, ['--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-features=msEdgeSidebar,msUndersideButton', `--remote-debugging-port=${CDP_PORT}`, `--user-data-dir=${prof}`, '--window-size=1280,1000', 'about:blank'], { stdio: 'ignore' });
  edgePids.push(ep.pid);
  try {
    let targets;
    for (let i = 0; i < 50; i++) { try { targets = await (await fetch(`http://127.0.0.1:${CDP_PORT}/json`)).json(); if (targets.some(x => x.type === 'page')) break; } catch { /* 等待 */ } await sleep(300); }
    const page = targets.find(x => x.type === 'page' && /^about:blank/.test(x.url)) || targets.find(x => x.type === 'page');
    const ws = new WebSocket(page.webSocketDebuggerUrl);
    await new Promise((res, rej) => { ws.onopen = res; ws.onerror = rej; });
    let seq = 0; const waiting = {}; const handlers = [];
    ws.onmessage = (m) => { const d = JSON.parse(m.data); if (d.id && waiting[d.id]) { waiting[d.id](d); delete waiting[d.id]; } else if (d.method) { if (process.env.XSS_DEBUG === '2') console.log('[ev]', d.method); handlers.forEach(h => h(d)); } };
    const send = (method, params) => new Promise(r => { const id = ++seq; waiting[id] = r; ws.send(JSON.stringify({ id, method, params: params || {} })); });
    const ev = async (expr) => { const r = await send('Runtime.evaluate', { expression: expr, awaitPromise: true, returnByValue: true }); if (!r.result) throw new Error('CDP 錯誤：' + JSON.stringify(r.error || r).slice(0, 300)); if (r.result.exceptionDetails) throw new Error('eval: ' + JSON.stringify(r.result.exceptionDetails).slice(0, 500)); return r.result.result.value; };
    const waitFor = async (expr, ms) => { const end = Date.now() + (ms || 8000); while (Date.now() < end) { try { if (await ev(expr)) return true; } catch { /* 頁面切換中 */ } await sleep(150); } return false; };
    const nav = async (url) => { await send('Page.enable'); await send('Page.navigate', { url }); await sleep(800); await waitFor('document.readyState==="complete"', 20000); };
    // 以 replace[urlRegex] = 內容 攔截並替換頁面資源（紅燈：換成修補前版本；變異：換成改壞的版本）
    const replace = {};
    handlers.push((d) => { if (process.env.XSS_DEBUG && d.method === 'Page.frameNavigated') console.log('[nav]', d.params.frame.url); });
    handlers.push(async (d) => {
      if (d.method !== 'Fetch.requestPaused') return;
      const url = d.params.request.url.split('?')[0];
      const key = Object.keys(replace).find(k => new RegExp(k).test(url));
      if (process.env.XSS_DEBUG) console.log('[fetch]', url, key ? 'REPLACED' : '');
      if (key) await send('Fetch.fulfillRequest', { requestId: d.params.requestId, responseCode: 200, responseHeaders: [{ name: 'Content-Type', value: key.includes('html') ? 'text/html; charset=utf-8' : 'application/javascript; charset=utf-8' }, { name: 'Cache-Control', value: 'no-store' }], body: Buffer.from(replace[key], 'utf8').toString('base64') });
      else await send('Fetch.continueRequest', { requestId: d.params.requestId });
    });
    // 測試瀏覽器不要有 service worker：index.html 在 SW 首次接管（controllerchange）時會 location.reload()，會把進行中的案例切斷
    await send('Page.addScriptToEvaluateOnNewDocument', { source: 'try { delete Navigator.prototype.serviceWorker; } catch (e) {}' });
    await send('Network.enable'); await send('Network.setCacheDisabled', { cacheDisabled: true }); await send('Network.setBypassServiceWorker', { bypass: true });   // 網站有 service worker：不繞過的話 app.js 由 SW 供應，攔截不到
    const fe = await send('Fetch.enable', { patterns: [{ urlPattern: '*' }] });
    if (process.env.XSS_DEBUG) console.log('[Fetch.enable]', JSON.stringify(fe).slice(0, 200));
    await send('Emulation.setDeviceMetricsOverride', { width: 1280, height: 1000, deviceScaleFactor: 1, mobile: false });
    return await fn({ send, ev, waitFor, nav, replace, sleep });
  } finally {
    edgePids.length = 0;
    killEdgeByProfile(prof);
    await sleep(500);
    try { fs.rmSync(prof, { recursive: true, force: true }); } catch { /* Edge 尚未放開檔案 */ }
  }
}
async function loginPage(b, who) {
  await b.nav(`${BASE}/login.html`);
  const st = await b.ev(`fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username:${JSON.stringify(who)},password:${JSON.stringify(PW)}})}).then(r=>r.status)`);
  return st === 200;
}
// 回傳 { app:[{name,ok,why}], admin:[…] }
async function runDomCases(b, label) {
  const res = { app: [], admin: [] };
  if (!await loginPage(b, '_s_u1')) throw new Error('登入失敗 _s_u1');
  await b.nav(`${BASE}/`);
  if (!await b.waitFor('typeof buildContactCard==="function" && typeof openView==="function" && !!document.getElementById("viewBody")', 25000)) throw new Error(label + ' 主頁沒有載入完成');
  await b.sleep(1200);
  res.jsArgDefined = await b.ev('typeof jsArg');
  if (process.env.XSS_EVAL) { console.log('[eval]', JSON.stringify(await b.ev(fs.readFileSync(process.env.XSS_EVAL, 'utf8')), null, 1)); }   // 除錯用：登入後在頁面執行指定腳本
  try { res.app = await b.ev(`(${pageAppCases.toString()})(${JSON.stringify(PAY)})`); }
  catch (e) { let last = '?'; try { last = await b.ev("sessionStorage.getItem('__xlast')"); } catch (e2) { /* */ } throw new Error(label + ' 案例執行中頁面被導走／關閉，最後一個案例：' + last + '（' + e.message + '）'); }
  if (!await loginPage(b, '_s_admin')) throw new Error('登入失敗 _s_admin');
  await b.nav(`${BASE}/admin.html`);
  if (!await b.waitFor('!!document.getElementById("ocrPreviewBtn") && typeof escH==="function"', 25000)) throw new Error(label + ' 後台沒有載入完成');
  await b.sleep(800);
  res.admin = await b.ev(`(${pageAdminCases.toString()})(${JSON.stringify(PAY)})`);
  return res;
}
let DOM_NAMES = [];
async function secDom() {
  console.log('[C] 瀏覽器 DOM（修補後程式碼）');
  await freshServer();
  const r = await withBrowser((b) => runDomCases(b, '修補後'));
  if (!r) return null;
  const all = r.app.concat(r.admin);
  DOM_NAMES = all.map(x => x.name);
  all.forEach(x => ok(x.ok, 'DOM ' + x.name, x.why));
  ok(all.length >= 24, `DOM 案例數量足夠（${all.length} 筆）`);
  return r;
}
async function secRed() {
  console.log('[D] 紅燈：同一組斷言對修補前程式碼（' + XSS_BASE.slice(0, 10) + '）');
  const base = XSS_BASE;
  const oldApp = git(['show', `${base}:_client/app.js`]), oldAdmin = git(['show', `${base}:_client/admin.html`]);
  await freshServer();
  const r = await withBrowser((b) => { b.replace['app\\.js$'] = oldApp; b.replace['admin\\.html$'] = oldAdmin; return runDomCases(b, '修補前'); });
  if (!r) return null;
  const all = r.app.concat(r.admin);
  ok(r.jsArgDefined === 'undefined', '紅燈：頁面確實載入了修補前的 app.js（jsArg 不存在）', r.jsArgDefined);
  const failed = all.filter(x => !x.ok);
  console.log(`  修補前：${failed.length}/${all.length} 筆斷言失敗（預期大多失敗）`);
  // 這幾筆在修補前本來就是安全的（已跳脫或只檢查工具函式），不要求紅燈
  const mustRed = all.filter(x => !/^admin:escH/.test(x.name) && x.name !== 'app:renderOpportunityBadge' || false);
  const stillGreen = mustRed.filter(x => x.ok).map(x => x.name);
  notes.push(`紅燈：修補前 ${failed.length}/${all.length} 筆失敗；修補前仍通過（代表該路徑修補前已安全或 payload 未觸發）：${stillGreen.join(', ') || '無'}`);
  ok(failed.length >= 15, `紅燈：修補前程式碼至少 15 筆 DOM 斷言失敗（實際 ${failed.length}）— 測試對 XSS 敏感`);
  for (const must of ['app:buildContactCard', 'app:openView', 'app:updateCompanyDatalist', 'app:renderOppTable', 'app:renderZombieSection', 'app:renderContractTable', 'app:renderReceivables', 'app:renderSapNoticeBar', 'app:openAllVisitsOfDay', 'app:renderCompanyMasterDetail', 'admin:OCR 預覽表']) {
    const x = all.find(y => y.name === must);
    ok(x && !x.ok, `紅燈：修補前 ${must} 會失敗（payload 真的能執行／破壞屬性）`, x ? x.why.slice(0, 80) : '無此案例');
  }
  return r;
}
async function secMut() {
  console.log('[E] 變異檢查（拿掉規則 → 對應測試必須失敗）');
  const srvSrc = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const mutate = (src, from, to, id) => { if (src.split(from).length !== 2) throw new Error('變異 ' + id + ' 找不到唯一錨點：' + from.slice(0, 50)); return src.replace(from, () => to); };
  const mkSrv = (id, src) => { const f = path.join(ROOT, `_xss_mut_${id}.js`); fs.writeFileSync(f, src); tmpFiles.push(f); return f; };
  // 變異 1：POST 聯絡人不驗證 cardImage
  {
    const f = mkSrv('m1', mutate(srvSrc, 'if (!isValidCardImage(req.body.cardImage)) return res.status(400).json({ error: CARD_IMAGE_ERR });', 'void 0;', 'm1'));
    await freshServer(f);
    const delta = await captureFailures(() => secServer('變異1:不驗證 cardImage'));
    ok(delta >= 6, '變異1（POST 不驗證 cardImage）→ cardImage 矩陣有 ≥6 筆斷言失敗（抓得到）', '失敗 ' + delta);
  }
  // 變異 2：帳號字元規則放行任何字元
  {
    const f = mkSrv('m2', mutate(srvSrc, 'const USERNAME_RE = /^[A-Za-z0-9._-]{1,50}$/;', 'const USERNAME_RE = /^[\\s\\S]*$/;', 'm2'));
    await freshServer(f);
    const delta = await captureFailures(() => secServer('變異2:不限帳號字元'));
    ok(delta >= 5, '變異2（帳號字元規則放行任何字元）→ 帳號矩陣有 ≥5 筆斷言失敗（抓得到）', '失敗 ' + delta);
  }
  // 變異 3：前端名片圖 src 不跳脫 → DOM 測試必須失敗
  {
    const appSrc = fs.readFileSync(path.join(ROOT, '_client/app.js'), 'utf8');
    const m = mutate(appSrc, 'src="${escapeHtml(c.cardImage)}" alt="名片">', 'src="${c.cardImage}" alt="名片">', 'm3');
    await freshServer();
    const r = await withBrowser((b) => { b.replace['app\\.js$'] = m; return runDomCases(b, '變異3'); });
    if (r) {
      const x = r.app.find(y => y.name === 'app:buildContactCard');
      ok(x && !x.ok, '變異3（名片卡 <img src> 不跳脫）→ app:buildContactCard 斷言失敗（抓得到）', x && x.why.slice(0, 60));
      const others = r.app.filter(y => y.name !== 'app:buildContactCard' && !y.ok);
      ok(others.length === 0, '變異3 只影響被拿掉規則的那一個渲染函式（其餘仍通過）', '其他失敗 ' + others.map(y => y.name + ' ' + y.why.slice(0, 80)).join('；'));
    }
  }
}
async function secShots() {
  const dir = process.env.XSS_SHOTS; if (!dir) return;
  console.log('[S] 截圖 →', dir);
  fs.mkdirSync(dir, { recursive: true });
  const oldApp = git(['show', `${XSS_BASE}:_client/app.js`]);
  const MP = { img: MARK, brk: MARK, js: MARK, card: 'x" onerror="window.__x=1;this.style.outline=\'3px solid #d32f2f\'" a="' };
  for (const mode of ['before', 'after']) {
    await freshServer();
    await withBrowser(async (b) => {
      if (mode === 'before') b.replace['app\\.js$'] = oldApp;
      await loginPage(b, '_s_u1'); await b.nav(`${BASE}/`);
      await b.waitFor('typeof buildContactCard==="function"', 25000); await b.sleep(1500);
      const shoot = async (file, setup) => {
        await b.ev(`(async function(P,mode){
          window.__x=0; document.querySelectorAll('#__xssShot').forEach(e=>e.remove());
          const ov=document.createElement('div'); ov.id='__xssShot'; ov.style.cssText='position:fixed;inset:0;z-index:99999;background:#fff;padding:18px;overflow:auto;font:14px sans-serif';
          const h=document.createElement('div'); h.style.cssText='font-weight:700;font-size:16px;margin-bottom:10px;padding:6px 10px;color:#fff;background:'+(mode==='before'?'#b71c1c':'#1b5e20'); h.textContent=(mode==='before'?'修補前（釘死 commit）':'修補後')+'：同一份含 XSS payload 的資料'; ov.appendChild(h);
          document.body.appendChild(ov); window.__ov=ov;
        })(${JSON.stringify(MP)}, ${JSON.stringify(mode)})`);
        await b.ev(`(async function(P){${setup}})(${JSON.stringify(MP)})`);
        // 等圖片載入／失敗事件跑完（payload 的 onerror 在請求失敗後才觸發）；用 addEventListener 不覆蓋 payload 自己的 onerror
        await b.ev("Promise.all([...document.images].map(i => i.complete ? 0 : new Promise(r => { i.addEventListener('load', r); i.addEventListener('error', r); setTimeout(r, 4000); })))");
        await b.sleep(500);
        const rect = await b.ev(`(function(){const r=document.getElementById('__xssShot').getBoundingClientRect();return {x:0,y:0,width:r.width,height:Math.min(r.height,520)};})()`);
        const shot = await b.send('Page.captureScreenshot', { format: 'png', clip: { x: 0, y: 0, width: rect.width, height: rect.height, scale: 1 } });
        fs.writeFileSync(path.join(dir, file), Buffer.from(shot.result.data, 'base64'));
        const x = await b.ev('window.__x'); notes.push(`截圖 ${file}：頁面腳本是否執行 window.__x=${x}（before 預期 1／after 預期 0）`);
        await b.ev(`document.getElementById('__xssShot').remove()`);
      };
      await shoot(`${mode}-contact-card.png`, `
        const c={id:'s1',name:'王小明',company:'沙盒公司甲',title:'經理',phone:'02-0000-0000',customerType:'customer',opportunityStage:'A',cardImage:P.card,owner:'x',createdAt:new Date().toISOString()};
        c.name=P.img; c.company=P.img; c.title=P.img;
        const wrap=document.createElement('div'); wrap.style.cssText='width:360px'; wrap.appendChild(buildContactCard(c));
        const note=document.createElement('div'); note.style.cssText='margin:8px 0;color:#555'; note.textContent='名片圖 cardImage 欄位內容：'+P.card;
        window.__ov.appendChild(note); window.__ov.appendChild(wrap);`);
      await shoot(`${mode}-lost-opportunities.png`, `
        $('oppFilterYear').value=''; $('oppFilterStage').value='';
        allOpportunities=[]; allZombieOpps=[];
        allLostOppsForTargets=[{id:'o2',company:P.img,contactName:'林小華',category:'ERP',product:P.img,deleteReason:P.img,stage:'B',amount:5,expectedDate:'2026-12-01',createdAt:new Date().toISOString(),deletedAt:new Date().toISOString()}];
        renderOppTable();
        const t=document.createElement('table'); t.style.cssText='border-collapse:collapse;width:100%'; t.innerHTML='<thead><tr><th>建立日期</th><th>客戶公司</th><th>聯絡人</th><th>商機類別</th><th>商品項目</th><th>金額</th><th>預計成交日</th><th>狀態</th></tr></thead>';
        const tb=document.createElement('tbody'); tb.innerHTML=document.getElementById('oppTableBody').innerHTML; t.appendChild(tb);
        t.querySelectorAll('td,th').forEach(e=>e.style.cssText='border:1px solid #ccc;padding:6px;text-align:left');
        window.__ov.appendChild(t);`.replace(/\$\('/g, "document.getElementById('"));
    });
  }
}

// ═════════════ 主程式 ═════════════
(async () => {
  let exitCode = 0;
  try {
    if (want('static')) secStatic();
    if (want('server')) { await freshServer(); await secServer(); }
    let domRes = null;
    if (want('dom')) domRes = await secDom();
    if (want('red')) await secRed();
    if (want('mut')) await secMut();
    if (want('shots') || process.env.XSS_SHOTS) await secShots();
  } catch (e) { fail++; failures.push('例外：' + (e && e.stack || e)); console.log('  ✗ 例外', e && e.stack || e); }
  finally {
    await stopServer();
    tmpFiles.forEach(f => { try { fs.unlinkSync(f); } catch { /* 已刪 */ } });
    cleanupFiles();
  }
  notes.forEach(n => console.log('  註：' + n));
  console.log(`\n結果：${pass} 通過，${fail} 失敗`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); exitCode = 1; }
  process.exit(exitCode);
})();
