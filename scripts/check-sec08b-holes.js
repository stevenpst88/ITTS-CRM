#!/usr/bin/env node
'use strict';
/**
 * SEC-08b：SEC-08（整合區 SSRF）之後補的三個外送請求缺口。
 *   洞 1  lib/pushNotify.js：Web Push 訂閱 endpoint 由使用者提交，伺服器之後會對它送請求（盲 SSRF）
 *   洞 2  POST /api/ai/company-insight：使用者自填網址，伺服器裸 fetch
 *   洞 3  整合設定被改 Base URL／Token Endpoint 後，已存的密碼／clientSecret 會被送到新網址
 *
 *   [U] pushNotify 純函式／注入測試（不連網路；DNS 與 web-push 都用注入或假的）
 *   [S] 靜態守門：server.js／pushNotify.js 的關鍵檢查還在
 *   [R] 路由端對端：在本程式樹以臨時 auth.json／data.json 啟動 server.js（本地 JSON 後端），完全沒有對外連線
 *   [M] 變異（SEC08B_MUT=1）：逐一拿掉每一處檢查，對應測試必須變紅
 *
 *   node scripts/check-sec08b-holes.js                          全跑（U、S、R；不含 M）
 *   SEC08B_ONLY=unit,static node scripts/check-sec08b-holes.js  只跑不開伺服器的部分
 *   SEC08B_MUT=1 node scripts/check-sec08b-holes.js             只跑變異
 *   SEC08B_PORT=3411 …（預設 3410；只接受 3410–3419）；SEC08B_PUSH／SEC08B_SERVER 指向被測檔（變異用）
 * 守門：[R] 在程式樹已有 auth.json／data.json／audit.log.json 時拒絕執行；不含任何帳密（密碼只在記憶體隨機產生）。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const https = require('https');
const { spawn } = require('child_process');

// pushNotify 在載入時讀 VAPID 環境變數；測試用臨時產生的金鑰（只在記憶體）
{
  const wp = require('web-push');
  const k = wp.generateVAPIDKeys();
  process.env.VAPID_PUBLIC_KEY = k.publicKey; process.env.VAPID_PRIVATE_KEY = k.privateKey;
}
const PUSH_PATH = path.resolve(process.env.SEC08B_PUSH || path.join(__dirname, '..', 'lib', 'pushNotify.js'));
const SERVER_PATH = path.resolve(process.env.SEC08B_SERVER || path.join(__dirname, '..', 'server.js'));
const PN = require(PUSH_PATH);

let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));
const ONLY = (process.env.SEC08B_ONLY || 'unit,static,route').split(',');

const PUB = async () => [{ address: '93.184.216.34', family: 4 }];
const PRIV = async () => [{ address: '10.0.0.5', family: 4 }];

// ───────────────────────── [U] pushNotify ─────────────────────────
function mkSub(endpoint) {
  const ecdh = crypto.createECDH('prime256v1'); ecdh.generateKeys();
  return { endpoint, keys: { p256dh: ecdh.getPublicKey().toString('base64url'), auth: crypto.randomBytes(16).toString('base64url') } };
}
async function unit() {
  console.log('[U1] validateEndpoint：正常的推播服務網域都要放行（不寫死白名單，所以新網域也過）');
  for (const e of [
    'https://fcm.googleapis.com/fcm/send/abc123', 'https://updates.push.services.mozilla.com/wpush/v2/abc',
    'https://wns2-par02p.notify.windows.com/w/?token=abc', 'https://web.push.apple.com/QGxyz', 'https://some-new-push-service.example.net/push/1',
  ]) ok((await PN.validateEndpoint(e, { lookup: PUB })).ok === true, `U1 放行：${e.split('/')[2]}`);
  console.log('[U2] validateEndpoint：內網／不合格一律擋，且回中文原因');
  const bad = [
    ['http://fcm.googleapis.com/x', 'http', /https/], ['https://169.254.169.254/latest/meta-data', 'metadata IP', /內部網路|本機/],
    ['https://10.0.0.5/x', '10/8', /內部網路|本機/], ['https://127.0.0.1/x', '127.0.0.1', /內部網路|本機/], ['https://[::1]/x', 'IPv6 ::1', /內部網路|本機/],
    ['https://localhost/x', 'localhost', /內部網路|本機/], ['https://user:pw@fcm.googleapis.com/x', '帶帳密', /帳號密碼/],
    ['https://fcm.googleapis.com:8443/x', '非 443 連接埠', /連接埠/], ['not a url', '亂字串', /格式/],
  ];
  for (const [e, why, re] of bad) {
    const r = await PN.validateEndpoint(e, { lookup: PUB });
    ok(r.ok === false && re.test(String(r.message)), `U2 擋：${why}（${r.message}）`);
  }
  for (const [e, why] of [[undefined, '未提供'], ['', '空字串'], [123, '非字串'], [{}, '物件'], ['https://fcm.googleapis.com/' + 'a'.repeat(2100), '過長']]) {
    ok((await PN.validateEndpoint(e, { lookup: PUB })).ok === false, `U2 擋：${why}`);
  }
  ok((await PN.validateEndpoint('https://evil.example.com/x', { lookup: PRIV })).ok === false, 'U2 擋：網域解析到內網位址');
  ok((await PN.validateEndpoint('https://evil.example.com/x', { lookup: async () => [{ address: '93.184.216.34' }, { address: '127.0.0.1' }] })).ok === false, 'U2 擋：解析結果有一個是內網（混合）');
  eq((await PN.validateEndpoint('https://nx.example.com/x', { lookup: async () => { throw new Error('ENOTFOUND'); } })).reason, 'DNS_FAILED', 'U2 DNS 失敗回 DNS_FAILED');

  console.log('[U3] makeSafeLookup：連線時才解析並檢查（防驗證後換位址）');
  const call = (lk, host, opts) => new Promise((res) => lk(host, opts, (err, a, f) => res({ err, a, f })));
  let r = await call(PN.makeSafeLookup(PRIV), 'x.example.com', {});
  ok(r.err && r.err.code === 'ESSRF', 'U3 解析到私有位址 → ESSRF');
  r = await call(PN.makeSafeLookup(async () => [{ address: '93.184.216.34', family: 4 }, { address: '169.254.169.254', family: 4 }]), 'x.example.com', { all: true });
  ok(r.err && r.err.code === 'ESSRF', 'U3 混合（含 metadata）→ ESSRF');
  r = await call(PN.makeSafeLookup(PUB), 'x.example.com', {});
  ok(!r.err && r.a === '93.184.216.34' && r.f === 4, 'U3 公開位址（單一）通過並回傳位址與 family');
  r = await call(PN.makeSafeLookup(PUB), 'x.example.com', { all: true });
  ok(!r.err && Array.isArray(r.a) && r.a[0].address === '93.184.216.34', 'U3 公開位址（all:true）回傳陣列');
  r = await call(PN.makeSafeLookup(async () => [{ address: '93.184.216.34', family: 4 }]), 'localhost', {});
  ok(r.err && r.err.code === 'ESSRF', 'U3 localhost 名稱一律擋');
  r = await call(PN.makeSafeLookup(async () => { throw Object.assign(new Error('x'), { code: 'ENOTFOUND' }); }), 'x.example.com', {});
  ok(r.err && r.err.code === 'ENOTFOUND', 'U3 解析失敗原樣回傳錯誤');

  console.log('[U4] sendToSubscription（假的 web-push）：不合格的端點根本不送；合格的帶 agent 與逾時');
  const calls = [];
  const fakeWp = { sendNotification: async (sub, body, opts) => { calls.push({ sub, opts }); return {}; } };
  let res = await PN.sendToSubscription({ endpoint: 'https://10.0.0.5/x', keys: {} }, { t: 1 }, { webpush: fakeWp, lookup: PUB });
  ok(calls.length === 0 && res.ok === false && res.gone === true, `U4 私有 IP 端點：沒呼叫 web-push，標記 gone（會被移除）`);
  res = await PN.sendToSubscription({ endpoint: 'https://evil.example.com/x', keys: {} }, { t: 1 }, { webpush: fakeWp, lookup: PRIV });
  ok(calls.length === 0 && res.gone === true, 'U4 網域解析到內網：沒呼叫 web-push，gone');
  res = await PN.sendToSubscription({ endpoint: 'http://fcm.googleapis.com/x', keys: {} }, { t: 1 }, { webpush: fakeWp, lookup: PUB });
  ok(calls.length === 0 && res.gone === true, 'U4 http 端點：沒呼叫 web-push，gone');
  res = await PN.sendToSubscription({ endpoint: 'https://nx.example.com/x', keys: {} }, { t: 1 }, { webpush: fakeWp, lookup: async () => { throw new Error('ENOTFOUND'); } });
  ok(calls.length === 0 && res.ok === false && res.gone === false, 'U4 DNS 暫時失敗：不送，但不當成失效（保留訂閱）');
  let rebindN = 0;
  const rebind = async () => (++rebindN === 1 ? [{ address: '93.184.216.34', family: 4 }] : [{ address: '10.0.0.5', family: 4 }]);   // 第一次公開（驗證）、之後私有（DNS rebinding）
  res = await PN.sendToSubscription({ endpoint: 'https://fcm.googleapis.com/fcm/send/abc', keys: {} }, { t: 1 }, { webpush: fakeWp, lookup: rebind });
  ok(res.ok === true && calls.length === 1, 'U4 正常端點：送出');
  const o = calls[0] && calls[0].opts;
  ok(o && o.agent instanceof https.Agent, 'U4 送出時帶了 https.Agent（連線釘定）');
  ok(o && typeof o.timeout === 'number' && o.timeout > 0 && o.timeout <= 15000, 'U4 送出時帶逾時');
  const lk = o && o.agent && o.agent.options && o.agent.options.lookup;
  ok(typeof lk === 'function', 'U4 agent 帶有自訂 lookup');
  if (typeof lk === 'function') { const rr = await call(lk, 'fcm.googleapis.com', {}); ok(rr.err && rr.err.code === 'ESSRF', 'U4 驗證後 DNS 被換成內網 → agent 的連線階段拒絕（rebinding）'); }
  res = await PN.sendToSubscription({ endpoint: 'https://fcm.googleapis.com/fcm/send/abc', keys: {} }, { t: 1 }, { webpush: { sendNotification: async () => { throw Object.assign(new Error('boom'), { code: 'ESSRF' }); } }, lookup: PUB });
  ok(res.ok === false && res.gone === true, 'U4 連線階段 ESSRF → gone');
  res = await PN.sendToSubscription({ endpoint: 'https://fcm.googleapis.com/fcm/send/abc', keys: {} }, { t: 1 }, { webpush: { sendNotification: async () => { throw Object.assign(new Error('gone'), { statusCode: 410 }); } }, lookup: PUB });
  ok(res.ok === false && res.gone === true, 'U4 410 仍視為失效（既有行為不變）');
  res = await PN.sendToSubscription({ endpoint: 'https://fcm.googleapis.com/fcm/send/abc', keys: {} }, { t: 1 }, { webpush: { sendNotification: async () => { throw Object.assign(new Error('srv'), { statusCode: 503 }); } }, lookup: PUB });
  ok(res.ok === false && res.gone === false, 'U4 503 仍是暫時性錯誤（既有行為不變）');

  console.log('[U5] 真的 web-push＋真的 agent（不連網）：驗證通過後 DNS 被換成內網，連線階段被擋');
  let n2 = 0;
  const rebind2 = async () => (++n2 === 1 ? [{ address: '93.184.216.34', family: 4 }] : [{ address: '127.0.0.1', family: 4 }]);
  const t0 = Date.now();
  res = await PN.sendToSubscription(mkSub('https://rebind.example.test/push/1'), { title: 't' }, { lookup: rebind2 });
  ok(res.ok === false && res.gone === true && /不允許的位址/.test(String(res.error)), `U5 真實 web-push 在連線階段被 agent 擋下（${res.error}）`);
  ok(Date.now() - t0 < 5000, 'U5 立即失敗（沒有真的去連線）');
}

// ───────────────────────── [S] 靜態守門 ─────────────────────────
function routeBody(src, head) {
  const lines = src.split('\n');
  const i = lines.findIndex((l) => l.startsWith(head));
  if (i < 0) return '';
  let j = i + 1; while (j < lines.length && !/^app\.(get|post|put|delete|patch)\(/.test(lines[j])) j++;
  return lines.slice(i, j).join('\n');
}
function staticGuard(srcRaw, pushRaw, label) {
  const sub = routeBody(srcRaw, "app.post('/api/push/subscribe'");
  ok(sub.includes('pushNotify.validateEndpoint(') && sub.indexOf('validateEndpoint(') < sub.indexOf('saveAuth('), `${label}：push/subscribe 在存檔前先 validateEndpoint`);
  const ci = routeBody(srcRaw, "app.post('/api/ai/company-insight'");
  ok(ci.includes('SU.validateUrl(url, PUBLIC_FETCH_OPTS)'), `${label}：company-insight 先 validateUrl`);
  ok(/await publicFetch\(url,/.test(ci) && !/(?<![\w.])fetch\s*\(/.test(ci), `${label}：company-insight 走 publicFetch、沒有裸 fetch`);
  ok(/PUBLIC_FETCH_OPTS = \{ ports: \['443'\]/.test(srcRaw), `${label}：publicFetch 只允許 443`);
  ok(/function publicFetch\(url, init\) \{ return SU\.safeFetch\(url, init, PUBLIC_FETCH_OPTS\); \}/.test(srcRaw), `${label}：publicFetch 就是 SU.safeFetch`);
  const save = routeBody(srcRaw, "app.post('/api/admin/integrations',");
  ok(save.includes('integSecretTainted(prev, f, item.baseUrl, item.tokenEndpoint)') && save.includes('SECRET_REENTRY_REQUIRED'), `${label}：整合存檔有機敏值失效檢查`);
  const test = routeBody(srcRaw, "app.post('/api/admin/integrations/test'");
  ok((test.match(/integSecretTainted\(/g) || []).length >= 2, `${label}：整合測試連線有機敏值失效檢查（resolveSecret＋提早回 400）`);
  ok(/function sendToSubscription\(/.test(pushRaw) || /async function sendToSubscription\(/.test(pushRaw), `${label}：pushNotify.sendToSubscription 存在`);
  ok(/agent: makeSafeAgent\(o\.lookup\)/.test(pushRaw), `${label}：web-push 送出帶 makeSafeAgent`);
  ok(/validateEndpoint\(subscription && subscription\.endpoint/.test(pushRaw), `${label}：送出前再驗一次 endpoint`);
  // 全站：整合區以外「使用者可控網址」的裸 fetch 不可再有（company-insight 之外的既有固定主機請求由 check-safe-url 守）
  ok(!/(?<![\w.])fetch\s*\(\s*(url|req\.body\.url)\b/.test(srcRaw.replace(/\/\/.*$/gm, '')), `${label}：server.js 沒有 fetch(url) 這種直接拿使用者網址的裸 fetch`);
}
function staticRun() {
  console.log('[S] 靜態守門');
  staticGuard(fs.readFileSync(SERVER_PATH, 'utf8'), fs.readFileSync(PUSH_PATH, 'utf8'), 'server.js／pushNotify.js');
}

// ───────────────────────── [R] 路由端對端 ─────────────────────────
const ROOT = path.resolve(process.env.SEC08B_ROOT || path.join(__dirname, '..'));
const PORT = parseInt(process.env.SEC08B_PORT || '3410', 10);
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));
const BASE = `http://localhost:${PORT}`;
const PW = 'Sb' + crypto.randomBytes(9).toString('hex') + '!Z';
const SENT = 'SENT' + crypto.randomBytes(6).toString('hex');   // 機敏值哨兵（只在記憶體）
let child = null;

async function route() {
  if (!(PORT >= 3410 && PORT <= 3419)) { console.log('拒絕：埠須在 3410-3419'); process.exit(2); }
  const guard = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
  if (guard.length) { console.log('拒絕：', ROOT, '已有', guard.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
  const bcrypt = require(path.join(__dirname, '..', 'node_modules', 'bcryptjs'));
  const { assertPortFree } = require('./_port-free');
  const now = new Date().toISOString();
  const mk = (u, role) => ({ username: u, displayName: u, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: now, passwordChangedAt: now, mustChangePassword: false });
  const cleanup = () => FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } });
  cleanup();
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users: [mk('_b_adm', 'admin'), mk('_b_u1', 'user'), mk('_b_u2', 'user'), mk('_b_u3', 'user')] }, null, 2));
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify({ contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], quotations: [], callins: [], keyAccounts: [], integrations: [] }, null, 2));
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE, GEMINI_API_KEY: 'dummy-not-used' });
  delete env.DATABASE_URL; delete env.MAIL_MODE; delete env.INTEGRATION_URL_PORTS;
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
    const call = async (who, method, url, body) => {
      const r = await fetch(BASE + url, { method, headers: { Cookie: jar[who], 'Content-Type': 'application/json' }, body: body ? JSON.stringify(body) : undefined });
      const t = await r.text(); let j = null; try { j = JSON.parse(t); } catch { /* 非 JSON */ }
      return { status: r.status, json: j, text: t };
    };
    for (const u of ['_b_adm', '_b_u1', '_b_u2', '_b_u3']) await login(u);
    const authSubs = (u) => { const a = JSON.parse(fs.readFileSync(path.join(ROOT, 'auth.json'), 'utf8')); const x = a.users.find(y => y.username === u); return (x && x.pushSubscriptions) || []; };
    const k = { p256dh: 'BPlaceholderKeyForTestOnly', auth: 'placeholderAuth' };

    console.log('[R1] POST /api/push/subscribe：不合格的 endpoint → 400＋中文訊息、且不會被存下來');
    for (const [endpoint, why, re] of [
      ['http://fcm.googleapis.com/x', 'http', /https/], ['https://169.254.169.254/latest', 'metadata IP', /內部網路|本機/], ['https://10.0.0.5/x', '10/8', /內部網路|本機/],
      ['https://127.0.0.1/x', '127.0.0.1', /內部網路|本機/], ['https://localhost/x', 'localhost', /內部網路|本機/], ['https://[::1]/x', 'IPv6 ::1', /內部網路|本機/],
      ['https://user:pw@93.184.216.34/x', '帶帳密', /帳號密碼/], ['https://93.184.216.34:8443/x', '非 443', /連接埠/],
    ]) {
      const r = await call('_b_u1', 'POST', '/api/push/subscribe', { endpoint, keys: k });
      ok(r.status === 400 && r.json && r.json.success === false && re.test(String(r.json.message)), `R1 ${why} → 400＋中文原因（${r.status} ${r.json && r.json.message}）`);
    }
    eq(authSubs('_b_u1').length, 0, 'R1 被拒的端點都沒有寫進 auth.json');
    let r = await call('_b_u1', 'POST', '/api/push/subscribe', { endpoint: 'https://93.184.216.34/push/abc', keys: k });
    ok(r.status === 200 && r.json && r.json.success === true && r.json.count === 1, `R1 合格端點（公開 IP 字面值）照常訂閱成功（${r.status}）`);
    eq(authSubs('_b_u1').length, 1, 'R1 合格端點有存下');
    r = await call('_b_u1', 'POST', '/api/push/subscribe', { endpoint: 'https://93.184.216.34/push/abc', keys: k });
    ok(r.status === 200 && r.json.count === 1, 'R1 同端點重複訂閱仍是覆蓋（既有行為不變）');
    r = await call('_b_u1', 'POST', '/api/push/subscribe', { endpoint: 'https://93.184.216.34/x' });
    ok(r.status === 400, 'R1 缺 keys 仍是原本的 400');

    console.log('[R2] POST /api/ai/company-insight：使用者自填網址不合格 → 400＋中文訊息（沒有任何對外連線）');
    const ciCases = [
      ['http://example.com', 'http', /https/], ['https://169.254.169.254/latest', 'metadata IP', /內部網路|本機/], ['https://10.0.0.5/', '10/8', /內部網路|本機/],
      ['https://localhost/', 'localhost', /內部網路|本機/], ['https://user:pw@93.184.216.34/', '帶帳密', /帳號密碼/], ['https://93.184.216.34:8443/', '非 443', /連接埠/],
    ];
    const who = ['_b_u1', '_b_u2', '_b_u3'];   // aiLimiter 每人每分鐘 5 次，分散到三個帳號
    for (let i = 0; i < ciCases.length; i++) {
      const [url, why, re] = ciCases[i];
      r = await call(who[i % 3], 'POST', '/api/ai/company-insight', { url });
      ok(r.status === 400 && r.json && /不符合安全規定/.test(String(r.json.error)) && re.test(String(r.json.error)), `R2 ${why} → 400＋安全規定訊息（${r.status} ${r.json && r.json.error}）`);
    }
    r = await call('_b_u1', 'POST', '/api/ai/company-insight', { url: 'ftp://example.com' });
    ok(r.status === 400 && /有效的網址/.test(String(r.json && r.json.error)), 'R2 非 http(s) 仍是原本的「請提供有效的網址」');
    r = await call('_b_u2', 'POST', '/api/ai/company-insight', { url: { a: 1 } });
    ok(r.status === 400, 'R2 url 不是字串 → 400（不會當機）');

    console.log('[R3] 整合設定：目標來源被改掉時，已存機敏值不能被帶去新網址');
    const I = '/api/admin/integrations';
    const base1 = 'https://93.184.216.35', base2 = 'https://93.184.216.36';
    const list = async () => (await call('_b_adm', 'GET', I)).json || [];
    const noLeak = (res, name) => ok(!res.text.includes(SENT), `${name}：回應不含機敏值`);
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', name: 'SAP', enabled: true, baseUrl: base1, authType: 'basic', username: 'u', password: SENT });
    ok(r.status === 200 && r.json.hasPassword === true, `R3 建立：存入密碼（${r.status}）`); noLeak(r, 'R3 建立');
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', baseUrl: base1 + '/', username: 'u2' });
    ok(r.status === 200 && r.json.hasPassword === true, 'R3 同主機（只多尾端斜線）＋不帶密碼 → 沿用既有行為，密碼保留');
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', baseUrl: base1 + '/other/path' });
    ok(r.status === 200 && r.json.hasPassword === true, 'R3 同主機改路徑＋不帶密碼 → 仍可（密碼送到同一主機）');
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', baseUrl: base2 });
    ok(r.status === 400 && r.json && r.json.code === 'SECRET_REENTRY_REQUIRED' && /重新輸入/.test(String(r.json.error)), `R3 改主機＋不帶密碼 → 400 要求重新輸入（${r.status}）`); noLeak(r, 'R3 改主機被拒');
    let cur = (await list())[0];
    ok(cur && cur.baseUrl === base1 + '/other/path'.replace(/\/+$/, '') && cur.hasPassword === true, `R3 被拒後設定完全沒變（baseUrl=${cur && cur.baseUrl}、密碼仍在）`);
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', baseUrl: base1 + ':8443' });
    ok(r.status === 400, `R3 只改連接埠＋不帶密碼 → 400（來源含埠）（${r.status}）`);
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', baseUrl: 'http://' + base1.slice(8) });
    ok(r.status === 400, `R3 只改協定（https→http）＋不帶密碼 → 400（${r.status}）`);
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', baseUrl: base2, password: '' });
    ok(r.status === 400, 'R3 密碼傳空字串視同沒給 → 400');
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', baseUrl: base2, password: SENT + 'new' });
    ok(r.status === 200 && r.json.baseUrl === base2 && r.json.hasPassword === true, `R3 改主機＋帶新密碼 → 允許（${r.status}）`); noLeak(r, 'R3 改主機重填');
    // 測試連線：改主機但沒重填 → 400（在任何對外連線之前）
    r = await call('_b_adm', 'POST', I + '/test', { type: 'sap-sales-cloud-v2', baseUrl: base1, authType: 'basic', username: 'u' });
    ok(r.status === 400 && r.json && r.json.ok === false && r.json.code === 'SECRET_REENTRY_REQUIRED', `R3 測試連線：改主機沒重填密碼 → 400，不會拿舊密碼去連（${r.status}）`); noLeak(r, 'R3 測試連線');
    // OAuth：clientSecret 隨 Token Endpoint；token 隨 baseUrl
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', baseUrl: base2, authType: 'oauth', clientId: 'c', tokenEndpoint: base2 + '/token', clientSecret: SENT + 'cs' });
    ok(r.status === 200 && r.json.hasClientSecret === true, `R3 OAuth：存入 clientSecret（${r.status}）`);
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', tokenEndpoint: base1 + '/token' });
    ok(r.status === 400 && r.json.code === 'SECRET_REENTRY_REQUIRED', `R3 OAuth：只改 Token Endpoint 主機＋不帶 clientSecret → 400（${r.status}）`);
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', baseUrl: base1 });
    ok(r.status === 400 && r.json.code === 'SECRET_REENTRY_REQUIRED', `R3 OAuth：只改 Base URL 主機＋不帶 clientSecret → 400（token 會被送過去）（${r.status}）`);
    r = await call('_b_adm', 'POST', I + '/test', { type: 'sap-sales-cloud-v2', authType: 'oauth', baseUrl: base2, tokenEndpoint: base1 + '/token' });
    ok(r.status === 400 && r.json.code === 'SECRET_REENTRY_REQUIRED', `R3 OAuth 測試連線：改 Token Endpoint 沒重填 → 400（${r.status}）`);
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', tokenEndpoint: base2 + '/token2' });
    ok(r.status === 200 && r.json.hasClientSecret === true, 'R3 OAuth：Token Endpoint 同主機改路徑 → 可');
    // 用不到的機敏值：改主機時被清掉（不會留著以後被切換認證方式後送出去）
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', authType: 'basic', baseUrl: base2, password: SENT + 'pw3', username: 'u' });
    ok(r.status === 200 && r.json.hasPassword === true && r.json.hasClientSecret === true, 'R3 切回 Basic（同主機）＋帶密碼 → clientSecret 暫時還在（主機沒變）');
    r = await call('_b_adm', 'POST', I, { type: 'sap-sales-cloud-v2', baseUrl: base1, password: SENT + 'pw4' });
    ok(r.status === 200 && r.json.hasPassword === true && r.json.hasClientSecret === false, `R3 Basic 改主機＋重填密碼 → 用不到的 clientSecret 被清掉（hasClientSecret=${r.json && r.json.hasClientSecret}）`); noLeak(r, 'R3 清除未用機敏值');
    cur = (await list())[0];
    ok(cur && cur.password === undefined && cur.clientSecret === undefined, 'R3 列表端點仍不回傳任何機敏欄位');
    ok(!JSON.stringify(await list()).includes(SENT), 'R3 列表端點不含機敏值');
  } finally {
    if (child) { try { process.kill(child.pid); } catch { /* 已結束 */ } for (let i = 0; i < 20 && child.exitCode === null; i++) await sleep(100); child = null; }
    cleanup();
  }
}

// ───────────────────────── [M] 變異：拿掉一處檢查，對應測試必須變紅 ─────────────────────────
const MUT_PUSH = [
  ['P1 訂閱端點不限連接埠', "Object.assign({}, o, { ports: PUSH_PORTS })", "Object.assign({}, o, { ports: null })"],
  ['P2 validateEndpoint 不看驗證結果', "return v.ok ? { ok: true } :", "return true ? { ok: true } :"],
  ['P3 送出前不再驗端點', "if (!chk.ok) {\n      if (chk.reason", "if (false) {\n      if (chk.reason"],
  ['P4 送出時不帶連線釘定的 agent', ", agent: makeSafeAgent(o.lookup) }", " }"],
  ['P5 safeLookup 不擋內網位址', "addrs.some((a) => SU.isBlockedIp(a.address))", "false"],
  ['P6 不合格端點不標記失效', "return { ok: false, gone: true, error: chk.message };", "return { ok: false, gone: false, error: chk.message };"],
  ['P7 safeLookup 不擋 localhost 名稱', "|| SU.isLocalName(hostname)", ""],
];
const MUT_SERVER = [
  ['S1 push/subscribe 不驗 endpoint', "  { const chk = await pushNotify.validateEndpoint(sub.endpoint); if (!chk.ok) return res.status(400).json({ success: false, message: '推播端點不符合安全規定：' + chk.message }); }\n", '', 'both'],
  ['S2 company-insight 不驗網址', "    { const chk = await SU.validateUrl(url, PUBLIC_FETCH_OPTS); if (!chk.ok) return res.status(400).json({ error: '此網址不符合安全規定：' + chk.message }); }\n", '', 'both'],
  ['S3 company-insight 改回裸 fetch', "const resp = await publicFetch(url, {", "const resp = await fetch(url, {", 'static'],
  ['S4 publicFetch 不限連接埠', "const PUBLIC_FETCH_OPTS = { ports: ['443'],", "const PUBLIC_FETCH_OPTS = { ports: null,", 'static'],
  ['S5 publicFetch 退回全域 fetch', "function publicFetch(url, init) { return SU.safeFetch(url, init, PUBLIC_FETCH_OPTS); }", "function publicFetch(url, init) { return fetch(url, init); }", 'static'],
  ['S6 整合存檔：需要的機敏值不重填也放行', "if (INTEG_SECRET_NEEDED[item.authType] === f) return res.status(400).json({ error: INTEG_REENTER_MSG, code: 'SECRET_REENTRY_REQUIRED' });\n", '', 'both'],
  ['S7 整合存檔：用不到的機敏值不清除', "      item[f] = '';\n    }\n  }\n", "    }\n  }\n", 'route'],
  ['S8 整合測試：不提早回 400', "  if (stored && _needF && !(typeof b[_needF]", "  if (false && stored && _needF && !(typeof b[_needF]", 'route'],
  ['S9 失效判斷永遠回 false', "function integSecretTainted(prev, field, baseUrl, tokenEndpoint) {\n  if (!prev || !prev[field]) return false;", "function integSecretTainted(prev, field, baseUrl, tokenEndpoint) {\n  return false;\n  if (!prev || !prev[field]) return false;", 'route'],
  ['S10 來源比對忽略連接埠', "return (x.protocol + '//' + x.host).toLowerCase();", "return (x.protocol + '//' + x.hostname).toLowerCase();", 'route'],
  ['S11 來源比對忽略協定', "return (x.protocol + '//' + x.host).toLowerCase();", "return ('//' + x.host).toLowerCase();", 'route'],
  ['S12 OAuth 不看 Token Endpoint 變更', "return field === 'password' ? baseChanged : (baseChanged || tokChanged);", "return field === 'password' ? baseChanged : baseChanged;", 'route'],
];
function runChild(env) {
  return new Promise((resolve) => {
    const c = spawn(process.execPath, [__filename], { env: Object.assign({}, process.env, env), stdio: ['ignore', 'pipe', 'pipe'] });
    let out = ''; c.stdout.on('data', (d) => { out += d; }); c.stderr.on('data', (d) => { out += d; });
    c.on('close', (code) => resolve({ code, out }));
  });
}
const tally = (out) => (out.match(/通過 \d+　失敗 \d+/) || ['?'])[0];
async function mutations() {
  const tmp = fs.mkdtempSync(path.join(require('os').tmpdir(), 'sec08b-mut-'));
  const libDir = path.join(__dirname, '..', 'lib');
  const pushSrc = fs.readFileSync(path.join(libDir, 'pushNotify.js'), 'utf8');
  const srvPath = path.join(__dirname, '..', 'server.js');
  const srvSrc = fs.readFileSync(srvPath, 'utf8');
  const absRewrite = (src) => src.replace("require('./safeUrl')", () => `require(${JSON.stringify(path.join(libDir, 'safeUrl.js'))})`).replace("require('web-push')", () => `require(${JSON.stringify(require.resolve('web-push'))})`);
  console.log('[M0] 未變異的基準必須全綠');
  const b = await runChild({ SEC08B_ONLY: 'unit,static', SEC08B_MUT: '' });
  ok(b.code === 0, 'M0 基準（unit＋static）全綠 ' + tally(b.out));
  console.log('[M1] lib/pushNotify.js 變異（unit 必須變紅）');
  for (const [name, from, to] of MUT_PUSH) {
    if (pushSrc.split(from).length !== 2) { ok(false, `${name}：變異目標在 lib/pushNotify.js 找不到或不唯一（測試與程式不同步）`); continue; }
    const p = path.join(tmp, 'pushNotify.js'); fs.writeFileSync(p, absRewrite(pushSrc.replace(from, () => to)));
    const rc = await runChild({ SEC08B_ONLY: 'unit', SEC08B_PUSH: p, SEC08B_MUT: '' });
    ok(rc.code !== 0 && /失敗 [1-9]/.test(rc.out), `${name} → 變紅（${tally(rc.out)}）`);
  }
  console.log('[M2] server.js 變異（static／route 依標註必須變紅）');
  for (const [name, from, to, scope] of MUT_SERVER) {
    if (srvSrc.split(from).length !== 2) { ok(false, `${name}：變異目標在 server.js 找不到或不唯一`); continue; }
    const mutated = srvSrc.replace(from, () => to);
    if (scope === 'static' || scope === 'both') {
      const p = path.join(tmp, 'server.js'); fs.writeFileSync(p, mutated);
      const rs = await runChild({ SEC08B_ONLY: 'static', SEC08B_SERVER: p, SEC08B_MUT: '' });
      ok(rs.code !== 0 && /失敗 [1-9]/.test(rs.out), `${name} → 靜態守門變紅（${tally(rs.out)}）`);
    }
    if (scope === 'route' || scope === 'both') {
      // 路由端對端：原地換上變異版 server.js，結束一定還原（finally＋行程結束保險）
      const bak = path.join(tmp, 'server.orig.js'); fs.copyFileSync(srvPath, bak);
      const restore = () => { try { fs.copyFileSync(bak, srvPath); } catch { /* 已還原 */ } };
      process.once('exit', restore);
      try {
        fs.writeFileSync(srvPath, mutated);
        const rr = await runChild({ SEC08B_ONLY: 'route', SEC08B_PORT: String(PORT + 1), SEC08B_MUT: '' });
        ok(rr.code !== 0 && /失敗 [1-9]/.test(rr.out), `${name} → 路由端對端變紅（${tally(rr.out)}）`);
      } finally { restore(); }
      ok(fs.readFileSync(srvPath, 'utf8') === srvSrc, `${name}：server.js 已還原`);
    }
  }
  try { fs.rmSync(tmp, { recursive: true, force: true }); } catch { /* 略 */ }
}

(async () => {
  let code = 0;
  try {
    if (process.env.SEC08B_MUT === '1') await mutations();
    else {
      if (ONLY.includes('unit')) await unit();
      if (ONLY.includes('static')) staticRun();
      if (ONLY.includes('route')) await route();
    }
  } catch (e) { console.log('錯誤：', e.stack || e.message); fail++; failures.push('例外：' + e.message); code = 1; }
  console.log(`\nSEC-08b 缺口檢查：通過 ${pass}　失敗 ${fail}`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); code = 1; }
  process.exit(code);
})();
