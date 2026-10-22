#!/usr/bin/env node
'use strict';
/**
 * 版本資訊頁「版本號＋Git 提交」檢查：lib/gitInfo.js 單元、admin.html 渲染（連結／跳脫）、真實伺服器 API（依序三輪，一次只開一個）。
 * 不含帳密：測試帳號 _s_vd_admin 與隨機密碼只在記憶體；臨時 auth.json/data.json 於結束時刪除。
 * 守門：工作樹已有 auth.json/data.json/audit.log.json 時拒絕執行。預設埠 3170（VD_PORT 可改；限 3170–3179）。
 *   node scripts/check-version-deploy.js
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');
const { assertPortFree } = require('./_port-free');
const ROOT = path.join(__dirname, '..');
const { deriveGitInfo } = require(path.join(ROOT, 'lib', 'gitInfo.js'));

let pass = 0, fail = 0;
function ok(c, n) { if (c) pass++; else { fail++; console.log('  FAIL:', n); } }
function eq(a, b, n) { ok(JSON.stringify(a) === JSON.stringify(b), `${n}（得到 ${JSON.stringify(a)}）`); }

// ── 1. package.json 版本 ──
{
  const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'));
  const lock = JSON.parse(fs.readFileSync(path.join(ROOT, 'package-lock.json'), 'utf8'));
  ok(/^\d+\.\d+\.\d+$/.test(pkg.version) && pkg.version !== '1.0.0', 'package.json version 已不是 1.0.0（' + pkg.version + '）');
  eq([lock.version, lock.packages[''].version], [pkg.version, pkg.version], 'package-lock 最上層版本與 package.json 一致');
}

// ── 2. deriveGitInfo 純函式 ──
const SHA = 'abcdef1234567890abcdef1234567890abcdef12';
{
  eq(deriveGitInfo({}), null, '無變數＝null');
  eq(deriveGitInfo(undefined), null, 'undefined env 不丟錯');
  eq(deriveGitInfo({ VERCEL_GIT_COMMIT_SHA: 'not-a-sha<img>' }), null, 'sha 非十六進位＝null');
  const g = deriveGitInfo({ VERCEL_GIT_COMMIT_SHA: SHA, VERCEL_GIT_COMMIT_REF: 'main', VERCEL_GIT_COMMIT_MESSAGE: '第一行\n第二行', VERCEL_GIT_REPO_OWNER: 'own-er', VERCEL_GIT_REPO_SLUG: 'my.repo_1' });
  eq(g, { sha7: 'abcdef1', sha: SHA, ref: 'main', message: '第一行', repoUrl: 'https://github.com/own-er/my.repo_1/commit/' + SHA }, '完整輸出');
  eq(deriveGitInfo({ VERCEL_GIT_COMMIT_SHA: SHA, VERCEL_GIT_COMMIT_MESSAGE: 'x'.repeat(300) }).message.length, 100, '說明截 100 字');
  eq(deriveGitInfo({ VERCEL_GIT_COMMIT_SHA: SHA, VERCEL_GIT_REPO_OWNER: 'a/b', VERCEL_GIT_REPO_SLUG: 'ok' }).repoUrl, null, 'owner 含 / 不組連結');
  eq(deriveGitInfo({ VERCEL_GIT_COMMIT_SHA: SHA, VERCEL_GIT_REPO_OWNER: 'ok', VERCEL_GIT_REPO_SLUG: 'a"><x' }).repoUrl, null, 'slug 含非法字元不組連結');
  eq(deriveGitInfo({ VERCEL_GIT_COMMIT_SHA: SHA, VERCEL_GIT_REPO_OWNER: '..', VERCEL_GIT_REPO_SLUG: 'ok' }).repoUrl, null, 'owner=.. 不組連結');
  eq(deriveGitInfo({ VERCEL_GIT_COMMIT_SHA: SHA }).repoUrl, null, '缺 owner/slug 不組連結');
  const leak = JSON.stringify(deriveGitInfo({ VERCEL_GIT_COMMIT_SHA: SHA, JWT_SECRET: 'SECRETJWT', DATABASE_URL: 'postgres://u:SECRETPW@h/x' }));
  ok(!/SECRET|postgres/.test(leak), '輸出不含其他環境變數');
}

// ── 3. admin.html 渲染 ──
const html = fs.readFileSync(path.join(ROOT, '_client', 'admin.html'), 'utf8');
function extractFn(name) {
  const i = html.indexOf('function ' + name + '(');
  if (i < 0) throw new Error('找不到 ' + name);
  let j = html.indexOf('{', i), depth = 0, k = j;
  for (; k < html.length; k++) { if (html[k] === '{') depth++; else if (html[k] === '}') { depth--; if (!depth) break; } }
  return html.slice(i, k + 1);
}
// 極簡 DOM 替身：驗證「全用 textContent／屬性，沒有 innerHTML」
function mkEl() {
  const el = { children: [], attrs: {}, style: {}, _t: '', set textContent(v) { this._t = String(v); this.children = []; }, get textContent() { return this._t + this.children.map(c => c.textContent).join(''); },
    appendChild(c) { this.children.push(c); return c; }, set href(v) { this.attrs.href = v; }, set target(v) { this.attrs.target = v; }, set rel(v) { this.attrs.rel = v; } };
  Object.defineProperty(el, 'innerHTML', { set() { throw new Error('不得使用 innerHTML'); } });
  return el;
}
const R = new Function('document', [extractFn('escH'), extractFn('verGitInfo'), extractFn('verRenderCommitLine'), 'return { escH, verGitInfo, verRenderCommitLine };'].join('\n'))(
  { createElement: () => mkEl(), createTextNode: (t) => ({ textContent: String(t) }) });
{
  eq(R.verGitInfo({}), null, '舊 API 無 deploy＝null');
  eq(R.verGitInfo({ deploy: { git: null } }), null, 'git=null＝null');
  const g = deriveGitInfo({ VERCEL_GIT_COMMIT_SHA: SHA, VERCEL_GIT_COMMIT_REF: 'main', VERCEL_GIT_REPO_OWNER: 'o', VERCEL_GIT_REPO_SLUG: 's' });
  const el = mkEl(); R.verRenderCommitLine(el, g);
  eq(el.textContent, 'abcdef1 · main', '橫幅小字文字');
  const a = el.children[0];
  eq([a.attrs.href, a.attrs.target, a.attrs.rel], ['https://github.com/o/s/commit/' + SHA, '_blank', 'noopener noreferrer'], '短碼連結屬性');
  const el2 = mkEl(); R.verRenderCommitLine(el2, null);
  eq(el2.textContent, '本機開發', '無資訊＝本機開發');
  const el3 = mkEl(); R.verRenderCommitLine(el3, { sha7: 'abcdef1', ref: '<img src=x onerror=1>', repoUrl: 'javascript:alert(1)' });
  ok(el3.children.length === 2 && el3.children[0].attrs === undefined, 'repoUrl 非 github https 時不產生連結');
  ok(el3.textContent.includes('<img src=x onerror=1>'), 'ref 以純文字呈現（textContent，非 HTML）');
  ok(/escH\('v' \+ d\.app\.version/.test(html) && /vRow\('部署提交',\s+escH\(/.test(html), '系統版本／部署提交列經 escH');
  ok(!/id="vVersion">v1\.0\.0/.test(html), 'admin.html 不再寫死 v1.0.0');
}

// ── 4. 真實伺服器 API：三輪（完整變數／無變數／惡意字串），一次只開一個 ──
const PORT = parseInt(process.env.VD_PORT || '3170', 10);
const BASE = `http://localhost:${PORT}`;
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json', 'mail-outbox.json'].map(f => path.join(ROOT, f));
async function round(extraEnv, PW) {
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE }, extraEnv);
  ['DATABASE_URL', 'VERCEL', 'VERCEL_ENV', 'STORAGE_BACKEND', 'AUDIT_EVENTS_TABLE', 'MAIL_MODE'].forEach(k => delete env[k]);
  if (!extraEnv.VERCEL_GIT_COMMIT_SHA) ['VERCEL_GIT_COMMIT_SHA', 'VERCEL_GIT_COMMIT_REF', 'VERCEL_GIT_COMMIT_MESSAGE', 'VERCEL_GIT_REPO_OWNER', 'VERCEL_GIT_REPO_SLUG'].forEach(k => delete env[k]);
  await assertPortFree(PORT);
  const child = spawn(process.execPath, [path.join(ROOT, 'server.js')], { cwd: ROOT, env, stdio: 'ignore' });
  try {
    let up = false;
    for (let i = 0; i < 80 && !up; i++) { await new Promise(r => setTimeout(r, 250)); try { up = (await fetch(BASE + '/login.html')).ok; } catch { /* 尚未起來 */ } }
    ok(up, '沙盒伺服器啟動');
    if (!up) return null;
    const lr = await fetch(BASE + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ username: '_s_vd_admin', password: PW }) });
    const cookie = lr.headers.getSetCookie().map(c => c.split(';')[0]).join('; ');
    const r = await fetch(BASE + '/api/admin/version', { headers: { Cookie: cookie } });
    eq(r.status, 200, 'version 200');
    return await r.json();
  } finally {
    try { process.kill(child.pid); } catch { /* 已結束 */ }
    await new Promise(r => setTimeout(r, 700));
  }
}
async function serverPart() {
  if (PORT < 3170 || PORT > 3179) { console.log('拒絕：只允許 3170–3179'); process.exit(2); }
  const g = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
  if (g.length) { console.log('拒絕：工作樹已有', g.map(f => path.basename(f)).join('/')); process.exit(2); }
  const bcrypt = require(path.join(ROOT, 'node_modules', 'bcryptjs'));
  const PW = 'Vd' + crypto.randomBytes(9).toString('hex') + '!Z', now = new Date().toISOString();
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users: [{ username: '_s_vd_admin', displayName: '_s_vd_admin', role: 'admin', active: true, password: bcrypt.hashSync(PW, 6), createdAt: now, passwordChangedAt: now, mustChangePassword: false, bu: [] }] }));
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify({ contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], callins: [], keyAccounts: [] }));
  try {
    // 一：完整 Vercel 變數
    const A = await round({ VERCEL_GIT_COMMIT_SHA: SHA, VERCEL_GIT_COMMIT_REF: 'main', VERCEL_GIT_COMMIT_MESSAGE: 'feat: 版本資訊\n第二行不應出現', VERCEL_GIT_REPO_OWNER: 'own', VERCEL_GIT_REPO_SLUG: 'repo', JWT_SECRET_LEAKCHECK: 'LEAKME' }, PW);
    if (A) {
      eq(A.deploy.git, { sha7: 'abcdef1', sha: SHA, ref: 'main', message: 'feat: 版本資訊', repoUrl: 'https://github.com/own/repo/commit/' + SHA }, 'A：deploy.git');
      eq(A.deploy.kind, 'local', 'A：原 deploy.kind 保留');
      ok(A.app.version === JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).version, 'A：app.version＝package.json');
      ok(!JSON.stringify(A).includes('LEAKME') && !JSON.stringify(A).includes('第二行'), 'A：不洩漏其他環境變數／說明第二行');
    }
    // 二：無變數
    const B = await round({}, PW);
    if (B) eq(B.deploy.git, null, 'B：無變數 deploy.git＝null（前端顯示本機開發）');
    // 三：惡意字串與非法 owner/slug
    const evil = '<img src=x onerror=alert(1)>';
    const C = await round({ VERCEL_GIT_COMMIT_SHA: SHA, VERCEL_GIT_COMMIT_REF: evil, VERCEL_GIT_COMMIT_MESSAGE: evil + 'x'.repeat(200), VERCEL_GIT_REPO_OWNER: 'ow"ner', VERCEL_GIT_REPO_SLUG: 'a/b' }, PW);
    if (C) {
      eq(C.deploy.git.repoUrl, null, 'C：非法 owner/slug 不組連結');
      ok(C.deploy.git.message.length <= 100, 'C：說明截 100 字');
      ok(C.deploy.git.message.includes('<img'), 'C：API 保留原文（由前端 textContent／escH 處理）');
      // 以前端函式渲染：不得變成 HTML
      const el = mkEl(); R.verRenderCommitLine(el, R.verGitInfo(C));
      ok(el.children.every(c => c.attrs === undefined || !c.attrs.href), 'C：前端無連結');
      ok(!R.escH(C.deploy.git.message).includes('<img'), 'C：escH 後無 <img');
    }
  } finally {
    FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } });
  }
}
serverPart().catch(e => { fail++; console.log('  FAIL: 伺服器段例外', e && e.message); }).then(() => {
  console.log(`\n結果：${pass} 通過，${fail} 失敗`);
  process.exit(fail ? 1 : 0);
});
