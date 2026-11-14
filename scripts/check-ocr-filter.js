#!/usr/bin/env node
'use strict';
/**
 * SEC-09：名片 OCR 上傳的圖片檢查（lib/ocrImage.js＋server.js 的 ocrUploadMw／ocrImageGate）。不需要瀏覽器，不呼叫 Gemini。
 *   [U] 純函式：magic bytes 偵測、MIME＋副檔名白名單、fileFilter 的錯誤碼；server.js 原始碼的中介層順序
 *       （圖片檢查必須在 aiCreditsGuard〔扣 AI 額度〕與 Gemini 之前）
 *   [R] 路由端對端：在臨時 auth.json／data.json 啟動 server.js（本地 JSON 後端、埠 3420–3429、**不設 GEMINI_API_KEY**），
 *       業務版 POST /api/ai/ocr-card 與管理員版 POST /api/admin/ai-ocr-card：
 *         · PDF、文字檔、.txt 改名成 .jpg、假 MIME、SVG、GIF、.exe、超過大小 → 400，且 AI 額度（auth.json 的 _subscription.aiCredits）不變
 *         · 真正的 JPEG／PNG／WEBP／HEIC 檔頭 → 通過圖片檢查（因為沒設金鑰，會在 Gemini 之前被 requireAi 回 503 AI_NOT_CONFIGURED），額度同樣不變
 *           （requireAi 放在扣額度之前，沒設金鑰不扣點）
 *
 *   node scripts/check-ocr-filter.js                    全跑
 *   OCRF_ONLY=unit node scripts/check-ocr-filter.js     只跑 [U]
 *   OCRF_PORT=3421 …（3420–3429）；OCRF_ROOT 可指向別的程式樹
 * 守門：[R] 在 OCRF_ROOT 已有 auth.json／data.json／audit.log.json（可能是真實資料）時拒絕執行；不含任何帳密（密碼只在記憶體）。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');
const OI = require('../lib/ocrImage');

let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));
const ONLY = (process.env.OCRF_ONLY || 'unit,route').split(',');

// 各格式「檔頭」樣本（後面補 0 到夠長；只需要通過檔頭檢查，不需要是完整圖片）
const pad = (arr, n = 64) => Buffer.concat([Buffer.from(arr), Buffer.alloc(Math.max(0, n - arr.length))]);
const JPEG = pad([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00]);
const PNG = pad([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0x0d, 0x49, 0x48, 0x44, 0x52]);
const WEBP = Buffer.concat([Buffer.from('RIFF'), Buffer.from([0x24, 0, 0, 0]), Buffer.from('WEBPVP8 '), Buffer.alloc(40)]);
const HEIC = Buffer.concat([Buffer.from([0, 0, 0, 0x18]), Buffer.from('ftypheic'), Buffer.alloc(40)]);
const PDF = Buffer.from('%PDF-1.7\n%âãÏÓ\n1 0 obj\n<<>>\nendobj\n' + 'x'.repeat(100));
const TXT = Buffer.from('這不是圖片，只是一段文字。'.repeat(10));
const GIF = Buffer.concat([Buffer.from('GIF89a'), Buffer.alloc(60)]);
const EXE = Buffer.concat([Buffer.from('MZ'), Buffer.alloc(100, 1)]);
const SVG = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>');

function unit() {
  console.log('[U1] sniffImageType：只認 JPEG／PNG／WEBP／HEIC 檔頭');
  eq(OI.sniffImageType(JPEG), 'image/jpeg', 'JPEG（FF D8 FF）');
  eq(OI.sniffImageType(PNG), 'image/png', 'PNG（89 50 4E 47 0D 0A 1A 0A）');
  eq(OI.sniffImageType(WEBP), 'image/webp', "WEBP（'RIFF'????'WEBP'）");
  eq(OI.sniffImageType(HEIC), 'image/heic', 'HEIC（ftyp＋brand）');
  for (const [name, buf] of [['PDF', PDF], ['純文字', TXT], ['GIF', GIF], ['EXE', EXE], ['SVG', SVG], ['空 buffer', Buffer.alloc(0)], ['太短', Buffer.from([0xff, 0xd8, 0xff])],
    ['RIFF 但不是 WEBP（WAV）', Buffer.concat([Buffer.from('RIFF'), Buffer.alloc(4), Buffer.from('WAVEfmt '), Buffer.alloc(20)])],
    ['ftyp 但 brand 是 mp4（影片）', Buffer.concat([Buffer.from([0, 0, 0, 0x18]), Buffer.from('ftypmp42'), Buffer.alloc(20)])],
    ['PNG 檔頭少一個位元組', pad([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0b])]])
    eq(OI.sniffImageType(buf), null, `${name} → 不是圖片`);
  eq(OI.sniffImageType('FFD8FF'), null, '不是 Buffer → null（不丟例外）');
  eq(OI.sniffImageType(null), null, 'null → null');

  console.log('[U2] isAllowedUpload：MIME 與副檔名都要在白名單');
  const f = (mimetype, originalname) => ({ mimetype, originalname });
  for (const [m, n] of [['image/jpeg', 'card.jpg'], ['image/jpeg', 'CARD.JPEG'], ['image/png', 'a.png'], ['image/webp', 'a.webp'], ['image/heic', 'IMG_0001.HEIC'], ['image/heif', 'a.heif'], ['image/jpeg', 'IMG_1.heic'], ['image/jpeg; charset=x', 'a.jpg']])
    ok(OI.isAllowedUpload(f(m, n)), `放行：${m} ${n}`);
  for (const [m, n] of [['application/pdf', 'a.pdf'], ['text/plain', 'a.txt'], ['image/jpeg', 'a.txt'], ['text/plain', 'a.jpg'], ['image/svg+xml', 'a.svg'], ['image/gif', 'a.gif'],
    ['image/jpeg', 'a.exe'], ['image/jpeg', 'noext'], ['', 'a.jpg'], ['application/octet-stream', 'a.jpg'], ['image/jpeg', 'a.jpg.exe'], ['image/png', 'a.svg']])
    ok(!OI.isAllowedUpload(f(m, n)), `拒絕：${m || '(空 MIME)'} ${n}`);
  ok(!OI.isAllowedUpload(null) && !OI.isAllowedUpload({}), '缺欄位一律拒絕');

  console.log('[U3] fileFilter：不合格丟帶 code 的錯誤，合格放行');
  let got;
  OI.fileFilter({}, f('application/pdf', 'a.pdf'), (e, accept) => { got = [e && e.code, e && e.message, accept]; });
  eq(got, ['OCR_BAD_TYPE', OI.BAD_IMAGE_MSG, undefined], 'PDF → OCR_BAD_TYPE＋中文訊息');
  OI.fileFilter({}, f('image/jpeg', 'card.jpg'), (e, accept) => { got = [e, accept]; });
  eq(got, [null, true], 'JPEG → 放行');
  ok(/JPG|PNG|WEBP|HEIC/.test(OI.BAD_IMAGE_MSG) && /[一-鿿]/.test(OI.BAD_IMAGE_MSG), '錯誤訊息是中文並說明支援的格式');

  console.log('[U4] server.js 的中介層順序：圖片檢查一定在「扣 AI 額度」與 Gemini 之前');
  const src = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');
  for (const [name, re, guard] of [
    ['業務版 /api/ai/ocr-card', /app\.post\('\/api\/ai\/ocr-card',([^\n]*)\n/, 'aiCreditsGuard'],
    ['管理員版 /api/admin/ai-ocr-card', /app\.post\('\/api\/admin\/ai-ocr-card',([^\n]*)\n/, 'aiCreditsGuard'],
  ]) {
    const m = re.exec(src);
    ok(!!m, `${name}：找到路由宣告`);
    if (!m) continue;
    const chain = m[1].split(',').map(x => x.trim().replace(/\(.*$/, '')).filter(Boolean);
    const idx = (k) => chain.indexOf(k);
    const up = chain.findIndex(x => /^uploadOcr(User|Admin)$/.test(x));
    ok(up >= 0 && idx('ocrImageGate') > up, `${name}：uploadOcr* → ocrImageGate（${chain.join(' → ')}）`);
    ok(idx('requireAi') > idx('ocrImageGate'), `${name}：requireAi 在圖片檢查之後`);
    ok(idx(guard) > idx('requireAi') && idx(guard) > idx('ocrImageGate'), `${name}：aiCreditsGuard（扣額度）在圖片檢查與 requireAi 之後`);
  }
  ok(/req\.file\.detectedMime/.test(src) && !/const mimeType = req\.file\.mimetype/.test(src.slice(src.indexOf("'/api/admin/ai-ocr-card'"), src.indexOf("'/api/ai/visit-suggest'"))), '送進 Gemini 的 MIME 用檔頭偵測的類型（detectedMime），不用用戶端宣稱的 file.mimetype');
}

// ───────────────────────── [R] 路由端對端 ─────────────────────────
const ROOT = path.resolve(process.env.OCRF_ROOT || path.join(__dirname, '..'));
const PORT = parseInt(process.env.OCRF_PORT || '3420', 10);
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));
const BASE = `http://localhost:${PORT}`;
const PW = 'Oc' + crypto.randomBytes(9).toString('hex') + '!Z';
let child = null;

async function route() {
  if (!(PORT >= 3420 && PORT <= 3429)) { console.log('拒絕：埠須在 3420-3429'); process.exit(2); }
  const guard = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
  if (guard.length) { console.log('拒絕：', ROOT, '已有', guard.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
  const bcrypt = require(path.join(__dirname, '..', 'node_modules', 'bcryptjs'));
  const { assertPortFree } = require('./_port-free');
  const now = new Date().toISOString();
  const mk = (u, role) => ({ username: u, displayName: u, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: now, passwordChangedAt: now, mustChangePassword: false });
  const USERS = ['_s_u1', '_s_u2', '_s_u3', '_s_u4', '_s_u5', '_s_u6'];   // 業務版有「每人每分鐘 5 次」的 AI 限流，所以輪流用多個帳號
  const cleanup = () => FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } });
  cleanup();
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({
    users: [mk('_s_adm', 'admin'), ...USERS.map(u => mk(u, 'user'))],
    _subscription: { plan: 'enterprise', aiCredits: 100 },   // 有設額度才會扣點；測試用它確認「額度不變」
  }, null, 2));
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify({ contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], quotations: [], callins: [], keyAccounts: [] }, null, 2));
  const credits = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'auth.json'), 'utf8'))._subscription.aiCredits;
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
  delete env.DATABASE_URL; delete env.MAIL_MODE; delete env.GEMINI_API_KEY; delete env.GOOGLE_API_KEY;   // 沒有金鑰：永遠不會真的呼叫 Gemini
  try {
    await assertPortFree(PORT);
    child = spawn(process.execPath, [path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: ['ignore', 'pipe', 'pipe'] });
    let out = ''; child.stdout.on('data', d => { out += d; }); child.stderr.on('data', d => { out += d; });
    let up = false;
    for (let i = 0; i < 80 && !up; i++) { await sleep(250); try { up = (await fetch(BASE + '/login.html')).ok; } catch { /* 尚未起來 */ } if (child.exitCode !== null) throw new Error('伺服器提前結束：' + out.slice(-400)); }
    if (!up) throw new Error('伺服器啟動逾時');

    const jar = {};
    const login = async (u) => {
      const r = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: u, password: PW }) });
      if (r.status !== 200) throw new Error('登入失敗 ' + u + ' ' + r.status);
      jar[u] = r.headers.getSetCookie().map(c => c.split(';')[0]).join('; ');
    };
    for (const u of ['_s_adm', ...USERS]) await login(u);
    let bizCalls = 0;
    const nextBizUser = () => USERS[Math.floor(bizCalls++ / 4) % USERS.length];   // 每個帳號最多 4 次（< 5 次／分鐘）
    const post = async (who, url, file) => {
      const fd = new FormData();
      if (file) fd.append('card', new Blob([file.buf], { type: file.type }), file.name);
      const r = await fetch(BASE + url, { method: 'POST', headers: who ? { Cookie: jar[who] } : {}, body: fd });
      const t = await r.text(); let j = null; try { j = JSON.parse(t); } catch { /* 非 JSON */ }
      return { status: r.status, json: j, text: t };
    };
    const BIZ = '/api/ai/ocr-card', ADM = '/api/admin/ai-ocr-card';
    const call = (kind, file) => kind === 'biz' ? post(nextBizUser(), BIZ, file) : post('_s_adm', ADM, file);

    const BAD = [
      ['PDF（application/pdf）', { buf: PDF, type: 'application/pdf', name: 'contract.pdf' }],
      ['純文字檔（text/plain）', { buf: TXT, type: 'text/plain', name: 'notes.txt' }],
      ['.txt 改名成 .jpg，MIME 也改成 image/jpeg（檔頭不是圖片）', { buf: TXT, type: 'image/jpeg', name: 'fake.jpg' }],
      ['PDF 改名成 .png，MIME 改成 image/png', { buf: PDF, type: 'image/png', name: 'fake.png' }],
      ['SVG（image/svg+xml）', { buf: SVG, type: 'image/svg+xml', name: 'x.svg' }],
      ['GIF（不在白名單）', { buf: GIF, type: 'image/gif', name: 'x.gif' }],
      ['.exe 偽裝成 image/jpeg', { buf: EXE, type: 'image/jpeg', name: 'x.exe' }],
      ['真 JPEG 內容但副檔名是 .txt', { buf: JPEG, type: 'image/jpeg', name: 'photo.txt' }],
      ['真 JPEG 內容但 MIME 是 application/octet-stream', { buf: JPEG, type: 'application/octet-stream', name: 'photo.jpg' }],
    ];
    const GOOD = [
      ['JPEG', { buf: JPEG, type: 'image/jpeg', name: 'card.jpg' }],
      ['PNG', { buf: PNG, type: 'image/png', name: 'card.png' }],
      ['WEBP', { buf: WEBP, type: 'image/webp', name: 'card.webp' }],
      ['HEIC', { buf: HEIC, type: 'image/heic', name: 'IMG_0001.HEIC' }],
    ];
    const start = credits();
    eq(start, 100, 'R0 起始額度 100');

    for (const [kind, label] of [['biz', '業務版 /api/ai/ocr-card'], ['adm', '管理員版 /api/admin/ai-ocr-card']]) {
      console.log(`[R] ${label}：不合格 → 400，額度不變`);
      for (const [why, file] of BAD) {
        const r = await call(kind, file);
        eq(r.status, 400, `${label}：${why} → 400`);
        ok(r.json && /圖片|JPG|PNG/.test(String(r.json.error)) && /[一-鿿]/.test(String(r.json.error)), `${label}：${why}：中文錯誤訊息（${r.json && r.json.error}）`);
        ok(r.json && r.json.error !== 'AI_NOT_CONFIGURED', `${label}：${why}：是在 Gemini 之前就被圖片檢查擋下（不是被 requireAi 的 503 擋）`);
        eq(credits(), start, `${label}：${why}：AI 額度不變`);
      }
      console.log(`[R] ${label}：合法圖片通過圖片檢查（沒設金鑰 → 在 Gemini 前被 requireAi 回 503），額度不變`);
      for (const [name, file] of GOOD) {
        const r = await call(kind, file);
        ok(r.status === 503 && r.json && r.json.error === 'AI_NOT_CONFIGURED', `${label}：${name} 通過圖片檢查（得到 ${r.status} ${r.json && r.json.error}）`);
        eq(credits(), start, `${label}：${name}：沒設金鑰不扣點`);
      }
      const noFile = await call(kind, null);
      ok(noFile.status === 400 && /未收到圖片/.test(String(noFile.json && noFile.json.error)), `${label}：沒帶檔案 → 400「未收到圖片」（${noFile.status}）`);
      eq(credits(), start, `${label}：沒帶檔案：額度不變`);
    }
    console.log('[R] 大小上限（維持原值：業務 8MB、管理員 10MB）→ 400，額度不變');
    let r = await call('biz', { buf: Buffer.concat([JPEG, Buffer.alloc(8 * 1024 * 1024)]), type: 'image/jpeg', name: 'big.jpg' });
    ok(r.status === 400 && r.json && /太大/.test(r.json.error), `業務版超過 8MB → 400（${r.status} ${r.json && r.json.error}）`);
    r = await call('adm', { buf: Buffer.concat([JPEG, Buffer.alloc(10 * 1024 * 1024)]), type: 'image/jpeg', name: 'big.jpg' });
    ok(r.status === 400 && r.json && /太大/.test(r.json.error), `管理員版超過 10MB → 400（${r.status} ${r.json && r.json.error}）`);
    eq(credits(), start, '超過大小：額度不變');
    console.log('[R] 既有的存取控制不變');
    r = await post(null, BIZ, GOOD[0][1]);
    ok(r.status === 401 || r.status === 302, `未登入打業務版 → 擋下（${r.status}）`);
    r = await post('_s_u6', ADM, GOOD[0][1]);
    ok(r.status === 403 || r.status === 401, `非管理員打管理員版 → 擋下（${r.status}）`);
    eq(credits(), start, '存取被擋時額度不變');
  } finally {
    if (child) { try { process.kill(child.pid); } catch { /* 已結束 */ } for (let i = 0; i < 20 && child.exitCode === null; i++) await sleep(100); child = null; }
    cleanup();
  }
}

(async () => {
  let code = 0;
  try {
    if (ONLY.includes('unit')) unit();
    if (ONLY.includes('route')) await route();
  } catch (e) { console.log('錯誤：', e.stack || e.message); fail++; failures.push('例外：' + e.message); code = 1; }
  console.log(`\nOCR 圖片檢查：通過 ${pass}　失敗 ${fail}`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); code = 1; }
  process.exit(code);
})();
