#!/usr/bin/env node
'use strict';
/**
 * SEC-08：外連 URL 的 SSRF 防護（lib/safeUrl.js）＋整合區（SAP）所有外送請求都經過檢查。
 *   [U] 純函式／注入測試：不連任何網路（DNS 與 https.request 都用假的），不需要伺服器
 *   [S] 靜態守門：掃 server.js 所有外送呼叫點（fetch／https.get／https.request），整合區一律 integFetch、
 *       三個會讀取已存網址的路由在發出請求前先 checkIntegrationUrl；新增未登記的裸 fetch 會變紅
 *   [R] 路由端對端：在本程式樹以臨時 auth.json／data.json 啟動 server.js（本地 JSON 後端），
 *       確認 integrations/test、sap-inspect、push/batch 對不合格網址回 400＋中文訊息，且完全沒有對外連線
 *   [M] 變異（SAFEURL_MUT=1）：逐一弄壞 lib／server.js，確認上面的檢查會變紅
 *
 *   node scripts/check-safe-url.js                      全跑（U、S、R；不含 M）
 *   SAFEURL_ONLY=unit,static node scripts/check-safe-url.js   只跑純函式與靜態守門
 *   SAFEURL_MUT=1 node scripts/check-safe-url.js        只跑變異
 *   SAFEURL_PORT=3381 …（預設 3380；只接受 3380–3389）；SAFEURL_LIB／SAFEURL_SERVER 指向被測的 safeUrl.js／server.js（變異用）
 * 守門：[R] 在程式樹已有 auth.json／data.json／audit.log.json 時拒絕執行；不含任何帳密（密碼只在記憶體）。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const EventEmitter = require('events');
const { spawn } = require('child_process');
const SU = require(process.env.SAFEURL_LIB || '../lib/safeUrl');
const LIB_PATH = path.resolve(process.env.SAFEURL_LIB || path.join(__dirname, '..', 'lib', 'safeUrl.js'));
const SERVER_PATH = path.resolve(process.env.SAFEURL_SERVER || path.join(__dirname, '..', 'server.js'));
const AC = require('../lib/auditChain');

let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }
const sleep = (n) => new Promise(r => setTimeout(r, n));
const ONLY = (process.env.SAFEURL_ONLY || 'unit,static,route').split(',');

// ───────────────────────── [U] 純函式與注入 ─────────────────────────
const PUB = async () => [{ address: '93.184.216.34' }];
const reasonOf = async (fn) => { try { await fn(); return 'NO_ERROR'; } catch (e) { return e.code || e.name; } };

// 假的 https.request：記錄每次呼叫的選項；script 決定回應。不碰網路。
function mkRequest(script, calls) {
  return (opts, cb) => {
    calls.push(opts);
    const req = new EventEmitter();
    req.destroy = () => {};
    req.end = (body) => {
      calls[calls.length - 1].__body = body;
      const r = script(opts, calls.length);
      if (r === 'hang') return;
      if (r && r.error) { setImmediate(() => req.emit('error', Object.assign(new Error(r.msg || 'x'), { code: r.error }))); return; }
      setImmediate(() => {
        const res = new EventEmitter();
        res.statusCode = r.status; res.headers = r.headers || {};
        cb(res);
        if (r.body) res.emit('data', Buffer.from(r.body));
        res.emit('end');
      });
    };
    return req;
  };
}

async function unit() {
  console.log('[U1] isBlockedIp：私有／本機／metadata 都擋，公開位址放行');
  const blocked = ['127.0.0.1', '127.255.255.254', '10.0.0.1', '10.255.255.255', '172.16.0.1', '172.31.255.255', '192.168.0.1', '192.168.255.255',
    '169.254.169.254', '169.254.0.1', '0.0.0.0', '0.1.2.3', '100.64.0.1', '224.0.0.1', '255.255.255.255',
    '::1', '::', '0:0:0:0:0:0:0:1', 'fc00::1', 'fd12:3456::1', 'fe80::1', 'fe80::abcd', 'ff02::1', '::ffff:127.0.0.1', '::ffff:10.0.0.1', '::ffff:169.254.169.254', 'not-an-ip',
    // 6to4／Teredo／ISATAP／SIIT／NAT64 內嵌內網 IPv4
    '2002:7f00:1::', '2002:a9fe:a9fe::1', '2002:0a00:0001::', '2002:c0a8:0101::', '2001:0:a00:1::1', '2001:0:4136:e378:8000:63bf:f5ff:fffa', '2001:0:4136:e378:8000:63bf:3fff:fdd2',
    '2606:4700::5efe:10.0.0.1', '::ffff:0:10.0.0.1', '64:ff9b::7f00:1', '64:ff9b:1::1'];
  const allowed = ['93.184.216.34', '8.8.8.8', '1.1.1.1', '172.15.255.255', '172.32.0.1', '11.0.0.1', '2606:4700:4700::1111', '::ffff:8.8.8.8',
    '2002:0808:0808::1', '2001:0:4136:e378:8000:63bf:f7f7:f7f7', '2606:4700::5efe:8.8.8.8', '::ffff:0:8.8.8.8'];
  for (const ip of blocked) ok(SU.isBlockedIp(ip), `擋下 ${ip}`);
  for (const ip of allowed) ok(!SU.isBlockedIp(ip), `放行公開位址 ${ip}`);

  console.log('[U2] checkUrlShape：只允許 https、不可帶帳密、連接埠白名單');
  for (const [u, why] of [
    ['http://example.com/x', 'http'], ['ftp://example.com', 'ftp'], ['javascript:alert(1)', 'javascript'], ['example.com/x', '沒有 scheme'],
    ['', '空字串'], ['https://user:pass@example.com/x', '帶帳密'], ['https://user@example.com/x', '只有使用者名稱'], ['https://:pw@example.com/x', '只有密碼'],
    ['https://example.com:8443/x', '非 443 連接埠'], ['https://example.com:80/x', '80 連接埠'],
  ]) ok(!SU.checkUrlShape(u).ok, `拒絕：${why}（${u}）`);
  ok(SU.checkUrlShape('https://example.com/x?token=abc').ok, '正常 https 公網網址通過');
  ok(SU.checkUrlShape('https://example.com:443/x').ok, '明寫 :443 通過');
  ok(SU.checkUrlShape('https://example.com:8443/x', { ports: ['443', '8443'] }).ok, 'ports 白名單含 8443 時通過');
  ok(SU.checkUrlShape('https://example.com:9999/x', { ports: null }).ok, 'ports:null 不限制連接埠');
  eq(SU.checkUrlShape('http://example.com').reason, 'SCHEME_NOT_HTTPS', 'http 的原因代碼');
  eq(SU.checkUrlShape('https://u:p@example.com').reason, 'CREDENTIALS_IN_URL', '帳密的原因代碼');
  eq(SU.checkUrlShape('https://example.com:8443').reason, 'PORT_NOT_443', '預設只允許 443 的原因代碼（與稽核錨點原本一致）');

  console.log('[U3] validateUrl：DNS 解析後再檢查位址、IP 字面值與各種繞法');
  const v = (u, lookup) => SU.validateUrl(u, { lookup: lookup || PUB });
  ok((await v('https://example.com/x')).ok, '正常公網網域通過');
  eq((await v('https://169.254.169.254/latest/meta-data/')).reason, 'BLOCKED_ADDRESS', 'metadata IP 字面值');
  eq((await v('https://127.0.0.1/')).reason, 'BLOCKED_ADDRESS', '127.0.0.1');
  eq((await v('https://10.1.2.3/')).reason, 'BLOCKED_ADDRESS', '10/8');
  eq((await v('https://172.20.0.1/')).reason, 'BLOCKED_ADDRESS', '172.16/12');
  eq((await v('https://192.168.1.1/')).reason, 'BLOCKED_ADDRESS', '192.168/16');
  eq((await v('https://0.0.0.0/')).reason, 'BLOCKED_ADDRESS', '0.0.0.0');
  eq((await v('https://[::1]/')).reason, 'BLOCKED_ADDRESS', 'IPv6 ::1');
  eq((await v('https://[fd00::1]/')).reason, 'BLOCKED_ADDRESS', 'IPv6 fc00::/7');
  eq((await v('https://[fe80::1]/')).reason, 'BLOCKED_ADDRESS', 'IPv6 fe80::/10');
  eq((await v('https://[::ffff:127.0.0.1]/')).reason, 'BLOCKED_ADDRESS', 'IPv4-mapped IPv6 的 127.0.0.1');
  eq((await v('https://2130706433/')).reason, 'BLOCKED_ADDRESS', '十進位整數寫法的 127.0.0.1（URL 解析後正規化）');
  eq((await v('https://0x7f.1/')).reason, 'BLOCKED_ADDRESS', '十六進位寫法的 127.0.0.1');
  eq((await v('https://localhost/')).reason, 'BLOCKED_ADDRESS', 'localhost');
  eq((await v('https://LOCALHOST./')).reason, 'BLOCKED_ADDRESS', 'LOCALHOST.（大寫＋尾端句點）');
  eq((await v('https://app.localhost/')).reason, 'BLOCKED_ADDRESS', '*.localhost');
  eq((await v('https://evil.example.com/', async () => [{ address: '10.0.0.5' }])).reason, 'BLOCKED_ADDRESS', '網域解析到私有 IP（指向內網）');
  eq((await v('https://meta.example.com/', async () => [{ address: '169.254.169.254' }])).reason, 'BLOCKED_ADDRESS', '網域解析到 metadata IP');
  eq((await v('https://mixed.example.com/', async () => [{ address: '93.184.216.34' }, { address: '192.168.0.9' }])).reason, 'BLOCKED_ADDRESS', '多筆解析結果只要有一筆是私有就拒絕');
  eq((await v('https://v6.example.com/', async () => [{ address: 'fc00::5' }])).reason, 'BLOCKED_ADDRESS', '網域解析到 IPv6 ULA');
  eq((await v('https://nx.example.com/', async () => { throw new Error('ENOTFOUND'); })).reason, 'DNS_FAILED', 'DNS 失敗');
  eq((await v('https://empty.example.com/', async () => [])).reason, 'DNS_FAILED', 'DNS 沒有結果');
  eq((await v('http://example.com/')).reason, 'SCHEME_NOT_HTTPS', 'http 先被擋（不做 DNS）');
  let looked = 0;
  await v('http://example.com/', async () => { looked++; return [{ address: '93.184.216.34' }]; });
  eq(looked, 0, 'http 在 DNS 之前就被擋，不會發 DNS 查詢');
  ok((await v('https://93.184.216.34/')).ok, '公開 IP 字面值通過');
  ok(!String((await v('http://secret-host.example.com/x?token=ABC')).message).match(/secret-host|ABC/), '拒絕訊息不回顯網址內容');
  const bad = await v('https://u:SECRETPW@example.com/');
  ok(!bad.ok && !String(bad.message).includes('SECRETPW'), '拒絕訊息不回顯帳密');

  console.log('[U4] safeFetch：連線釘在驗證過的 IP、轉址每跳重檢、逾時、大小上限');
  let calls = [];
  const o = (script, extra) => Object.assign({ lookup: PUB, request: mkRequest(script, calls) }, extra || {});
  const reset = () => { calls = []; };
  // 不合格 → 完全沒有發出請求
  reset();
  eq(await reasonOf(() => SU.safeFetch('http://example.com/x', {}, o(() => ({ status: 200 })))), 'SCHEME_NOT_HTTPS', 'http 被拒');
  eq(await reasonOf(() => SU.safeFetch('https://169.254.169.254/x', {}, o(() => ({ status: 200 })))), 'BLOCKED_ADDRESS', 'metadata 被拒');
  eq(await reasonOf(() => SU.safeFetch('https://u:p@example.com/x', {}, o(() => ({ status: 200 })))), 'CREDENTIALS_IN_URL', '帶帳密被拒');
  eq(await reasonOf(() => SU.safeFetch('https://rebind.example.com/x', {}, { lookup: async () => [{ address: '10.0.0.5' }], request: mkRequest(() => ({ status: 200 }), calls) })), 'BLOCKED_ADDRESS', '網域指向內網被拒');
  eq(calls.length, 0, '被拒的網址一個請求都沒發出');
  // 成功
  reset();
  let r = await SU.safeFetch('https://api.example.com/v1/x?$top=1', { method: 'GET', headers: { Authorization: 'Basic abc', Accept: 'application/json' } },
    o(() => ({ status: 200, headers: { 'content-type': 'application/json', location: '/a/b' }, body: '{"value":[1]}' })));
  ok(r.ok && r.status === 200, '成功回應 ok／status');
  eq(await r.json(), { value: [1] }, 'json()');
  eq(r.headers.get('Location'), '/a/b', 'headers.get 不分大小寫');
  eq(await r.text(), '{"value":[1]}', 'text()');
  eq(calls[0].host, '93.184.216.34', '連線釘在 DNS 解析、驗證過的那個 IP');
  eq(calls[0].servername, 'api.example.com', 'SNI 用原主機名（憑證驗證仍對網域）');
  eq(calls[0].headers.host, 'api.example.com', 'Host 標頭用原主機名');
  eq(calls[0].path, '/v1/x?$top=1', '路徑與查詢字串原樣');
  eq(calls[0].headers.authorization, 'Basic abc', '憑證標頭照送');
  // POST body
  reset();
  await SU.safeFetch('https://api.example.com/p', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{"a":1}' }, o(() => ({ status: 201, headers: { location: '/p/9' } })));
  eq([calls[0].method, calls[0].__body, calls[0].headers['content-length']], ['POST', '{"a":1}', 7], 'POST 送出 body 與 content-length');
  // 轉址
  reset();
  r = await SU.safeFetch('https://a.example.com/x', { headers: { Authorization: 'Bearer T' } },
    o((opts, n) => n === 1 ? { status: 302, headers: { location: 'https://b.example.org/y' } } : { status: 200, body: 'ok' }));
  eq([calls.length, r.status, calls[1].headers.host], [2, 200, 'b.example.org'], '轉址到另一個公網網域會跟隨（每一跳重新檢查）');
  ok(calls[1].headers.authorization === undefined, '跨來源轉址不帶 Authorization');
  reset();
  await SU.safeFetch('https://a.example.com/x', { headers: { Authorization: 'Bearer T' } },
    o((opts, n) => n === 1 ? { status: 301, headers: { location: '/moved' } } : { status: 200 }));
  eq([calls[1].path, calls[1].headers.authorization], ['/moved', 'Bearer T'], '同來源轉址（相對路徑）保留 Authorization');
  reset();
  eq(await reasonOf(() => SU.safeFetch('https://a.example.com/x', {}, { lookup: async (h) => [{ address: h === 'a.example.com' ? '93.184.216.34' : '10.0.0.9' }], request: mkRequest(() => ({ status: 302, headers: { location: 'https://internal.example.net/admin' } }), calls) })),
    'BLOCKED_ADDRESS', '轉址到「解析後是私有 IP」的網域 → 被拒');
  eq(calls.length, 1, '被拒的轉址目標沒有發出請求');
  reset();
  eq(await reasonOf(() => SU.safeFetch('https://a.example.com/x', {}, o(() => ({ status: 302, headers: { location: 'https://169.254.169.254/latest/meta-data/' } })))), 'BLOCKED_ADDRESS', '轉址到 metadata IP → 被拒');
  reset();
  eq(await reasonOf(() => SU.safeFetch('https://a.example.com/x', {}, o(() => ({ status: 302, headers: { location: 'http://b.example.com/' } })))), 'SCHEME_NOT_HTTPS', '轉址降級成 http → 被拒');
  reset();
  eq(await reasonOf(() => SU.safeFetch('https://a.example.com/x', {}, o(() => ({ status: 302, headers: { location: '/loop' } })))), 'TOO_MANY_REDIRECTS', '轉址超過 3 次 → 被拒');
  reset();
  await SU.safeFetch('https://a.example.com/x', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{"a":1}' },
    o((opts, n) => n === 1 ? { status: 303, headers: { location: '/done' } } : { status: 200 }));
  eq([calls[1].method, calls[1].__body, calls[1].headers['content-type']], ['GET', undefined, undefined], '303 之後改用 GET、不再帶 body');
  reset();
  r = await SU.safeFetch('https://a.example.com/x', {}, o(() => ({ status: 302, headers: { location: '/y' } }), { maxRedirects: 0 })).catch((e) => e);
  eq(r.code, 'TOO_MANY_REDIRECTS', 'maxRedirects:0 → 不跟隨轉址');
  // 連線失敗換下一個位址
  reset();
  r = await SU.safeFetch('https://a.example.com/x', {}, { lookup: async () => [{ address: '93.184.216.34' }, { address: '93.184.216.35' }],
    request: mkRequest((opts) => opts.host === '93.184.216.34' ? { error: 'ECONNREFUSED' } : { status: 200 }, calls) });
  eq([r.status, calls.map(c => c.host)], [200, ['93.184.216.34', '93.184.216.35']], '第一個位址連線被拒時改連下一個（兩個都是驗證過的公開位址）');
  // 逾時
  reset();
  const t0 = Date.now();
  const te = await SU.safeFetch('https://a.example.com/x', {}, o(() => 'hang', { timeoutMs: 80 })).catch((e) => e);
  ok(te.name === 'TimeoutError' && Date.now() - t0 < 1500, `逾時丟 TimeoutError（${Date.now() - t0}ms）`);
  const sigE = await SU.safeFetch('https://a.example.com/x', { signal: AbortSignal.timeout(60) }, o(() => 'hang', { timeoutMs: 5000 })).catch((e) => e);
  eq(sigE.name, 'TimeoutError', '呼叫端傳的 AbortSignal.timeout 也會中止（保留原本 8 秒逾時寫法）');
  // 大小上限
  reset();
  eq(await reasonOf(() => SU.safeFetch('https://a.example.com/x', {}, o(() => ({ status: 200, body: 'x'.repeat(50) }), { maxBytes: 10 }))), 'RESPONSE_TOO_LARGE', '回應超過上限 → 被拒');
  // 一般 HTTP 錯誤碼不是例外
  reset();
  r = await SU.safeFetch('https://a.example.com/x', {}, o(() => ({ status: 401, body: 'no' })));
  eq([r.ok, r.status], [false, 401], '401 以 ok:false 回傳（與原本 fetch 相同，不丟例外）');

  console.log('[U5] 稽核錨點的檢查仍是同一套（行為不變；完整比對在 scripts/check-audit-chain.js [5.x]）');
  if (!process.env.SAFEURL_LIB) ok(AC.isBlockedIp === SU.isBlockedIp, 'AC.isBlockedIp 就是共用函式');
  eq(AC.checkWebhookUrlShape('https://hooks.example.com:8443/x').reason, 'PORT_NOT_443', '稽核錨點仍只收 443');
  eq(AC.checkWebhookUrlShape('http://hooks.example.com/x').reason, 'SCHEME_NOT_HTTPS', '稽核錨點仍拒 http');
  ok(AC.checkWebhookUrlShape('https://hooks.example.com/path?token=abc').ok, '稽核錨點的正常網址仍通過');
  const rr = await AC.pushAnchor('https://meta.example.com/x', { seq: 1 }, { lookup: async () => [{ address: '169.254.169.254' }], request: () => { throw new Error('不該連線'); } });
  eq(rr, { ok: false, reason: 'BLOCKED_ADDRESS' }, '稽核錨點 pushAnchor 對 metadata 位址仍回 BLOCKED_ADDRESS');

  console.log('[U6] 各種繞過寫法（URL 解析會把十進位／八進位／十六進位／全形等正規化成 IPv4）一律不可通過');
  const bypass = [
    'https://2130706433/', 'https://0x7f000001/', 'https://0x7f.1/', 'https://127.1/', 'https://0177.0.0.1/', 'https://017700000001/', 'https://0/', 'https://0.0.0.0/',
    'https://[::ffff:7f00:1]/', 'https://[::ffff:127.0.0.1]/', 'https://[0:0:0:0:0:ffff:7f00:1]/', 'https://[::ffff:a9fe:a9fe]/', 'https://[::127.0.0.1]/',
    'https://[::1]:443/', 'https://[0000:0000:0000:0000:0000:0000:0000:0001]/', 'https://[fd00::1]/', 'https://[fe80::1]/', 'https://[64:ff9b::7f00:1]/',
    'https://[2002:7f00:1::]/', 'https://[2002:a9fe:a9fe::]/', 'https://[2001:0:a00:1::1]/', 'https://[2001:0:4136:e378:8000:63bf:f5ff:fffa]/',
    'https://localhost/', 'https://LOCALHOST/', 'https://localhost./', 'https://a.localhost/', 'https://localhost.localdomain/',
    'https://169.254.169.254/latest/meta-data/', 'https://0xa9fe.0xa9fe/', 'https://2852039166/', 'https://100.64.0.1/', 'https://198.18.0.1/', 'https://192.0.0.192/',
    'https://１２７.０.０.１/', 'https://user@127.0.0.1/', 'https://example.com@127.0.0.1/',
    'https://127.0.0.1:443/', 'https://127.0.0.1.:443/',
  ];
  let bypassed = [];
  for (const u of bypass) { const r = await SU.validateUrl(u, { lookup: PUB }); if (r.ok) bypassed.push(u); }
  eq(bypassed, [], `${bypass.length} 種繞過寫法全部被擋`);
  // 公開位址的對應寫法仍通過（避免誤擋）
  for (const u of ['https://[2002:0808:0808::1]/', 'https://[2001:0:4136:e378:8000:63bf:f7f7:f7f7]/', 'https://8.8.8.8/', 'https://[2606:4700:4700::1111]/', 'https://sap.example.com/x'])
    ok((await SU.validateUrl(u, { lookup: PUB })).ok, `合法公開位址照常通過 ${u}`);
  // DNS 回傳各種內嵌私網的 IPv6／mapped 寫法
  for (const a of ['2002:a9fe:a9fe::1', '2001:0:a00:1::1', '::ffff:10.0.0.1', '::ffff:a00:1', '64:ff9b::a00:1', '::1', 'fe80::1'])
    eq((await SU.validateUrl('https://x.example.com/', { lookup: async () => [{ address: a }] })).reason, 'BLOCKED_ADDRESS', `DNS 解析到 ${a} → 擋`);

  console.log('[U7] 連接埠白名單、機敏值不進錯誤訊息、壓縮回應');
  eq(SU.parsePorts(undefined), ['443'], 'parsePorts 未設 → 443');
  eq(SU.parsePorts('443, 8443 ,abc,0,70000,8443'), ['443', '8443'], 'parsePorts 濾掉無效值並去重');
  eq(SU.parsePorts('abc'), ['443'], 'parsePorts 全無效 → 443');
  const P = { ports: SU.parsePorts('443,8443'), lookup: PUB };
  ok((await SU.validateUrl('https://sap.example.com:8443/x', P)).ok, '白名單含 8443 → SAP 非預設埠的合法設定照常通過');
  eq((await SU.validateUrl('https://sap.example.com:22/x', P)).reason, 'PORT_NOT_ALLOWED', '不在白名單的 22 → 擋');
  eq((await SU.validateUrl('https://sap.example.com:8443/x', { lookup: PUB, ports: ['443'] })).reason, 'PORT_NOT_443', '沒設環境變數時 8443 → 擋');
  reset();
  const leak = await SU.safeFetch('https://a.example.com/x', { headers: { Authorization: 'Basic SECRETTOKEN' } },
    o(() => ({ error: 'ECONNRESET', msg: 'boom SECRETTOKEN https://u:SECRETPW@a.example.com' }))).catch((e) => e);
  ok(leak instanceof Error && !/SECRET/.test(leak.message + String(leak.stack || '').split('\n')[0]), `連線錯誤訊息不含機敏值（${leak.message}）`);
  reset();
  const zlib = require('zlib');
  r = await SU.safeFetch('https://a.example.com/x', {}, o(() => ({ status: 200, headers: { 'content-encoding': 'gzip' }, body: zlib.gzipSync('{"v":1}') })));
  eq(await r.json(), { v: 1 }, '對方仍回 gzip 時可解壓');
  reset();
  r = await SU.safeFetch('https://a.example.com/x', {}, o(() => ({ status: 200, headers: { 'content-encoding': 'gzip' }, body: zlib.gzipSync('x'.repeat(5000)) }), { maxBytes: 1000 }));
  eq(await reasonOf(() => r.text()), 'RESPONSE_TOO_LARGE', '解壓後超過上限（解壓炸彈）→ 擋');
  reset();
  await SU.safeFetch('https://a.example.com/x', {}, o(() => ({ status: 200 })));
  eq(calls[0].agent, false, '不使用共用連線池（agent:false）');
}

// ───────────────────────── [S] 靜態守門：server.js 所有外送呼叫點 ─────────────────────────
// 登記制：
//   · 整合區（GET /api/admin/integrations 到 integrations/logs 之間）不可有任何裸 fetch／https.get／https.request
//   · 其餘的裸 fetch 只能是「固定官方主機」（FIXED_HOSTS），或已知缺口 company-insight（要修時搬進 safeFetch 並從這裡移除）
//   · 新增任何未登記的外送呼叫 → 變紅，逼作者決定要不要走 safeFetch
const FIXED_HOSTS = ['openapi.twse.com.tw', 'www.tpex.org.tw', 'data.gcis.nat.gov.tw', 'api.duckduckgo.com'];
function stripComments(src) {
  return src.split('\n').map((l) => l.replace(/^\s*\/\/.*$/, '').replace(/(?<![:'"`])\/\/(?!\/).*$/, '')).join('\n');
}
function staticScan(srcRaw, label) {
  const src = stripComments(srcRaw);
  const lines = src.split('\n');
  const startIdx = lines.findIndex((l) => /^app\.get\('\/api\/admin\/integrations',/.test(l));
  const endIdx = lines.findIndex((l) => /^app\.get\('\/api\/admin\/integrations\/logs'/.test(l));
  ok(startIdx > 0 && endIdx > startIdx, `${label}：找得到整合區起訖行`);
  const sites = [];
  let route = '';
  lines.forEach((l, i) => {
    const mr = /^app\.(get|post|put|delete|patch)\('([^']+)'/.exec(l); if (mr) route = mr[1] + ' ' + mr[2];
    const re = /(?<![\w.])(integFetch|fetch)\s*\(|\b(https|http|lib)\.(get|request)\s*\(/g; let m;
    while ((m = re.exec(l))) sites.push({ line: i + 1, kind: m[1] || (m[2] + '.' + m[3]), ctx: lines.slice(i, i + 4).join(' '), route, inInteg: i >= startIdx && i <= endIdx });
  });
  const integ = sites.filter((s) => s.inInteg);
  const rawInInteg = integ.filter((s) => s.kind !== 'integFetch');
  eq(rawInInteg.map((s) => s.line + ':' + s.kind), [], `${label}：整合區沒有任何裸外送呼叫`);
  const nInteg = integ.filter((s) => s.kind === 'integFetch').length;
  ok(nInteg >= 13, `${label}：整合區的 integFetch 呼叫點 ≥ 13（實際 ${nInteg}）`);
  const rawOutside = sites.filter((s) => !s.inInteg && s.kind !== 'integFetch');
  const unregistered = rawOutside.filter((s) => {
    if (s.kind === 'fetch') {
      if (FIXED_HOSTS.some((h) => s.ctx.includes('https://' + h))) return false;
      return true;
    }
    // https.get／lib.get：只允許 Yahoo Finance 輔助函式（固定主機，由呼叫端傳入常數）
    return !/yahoo|fetchJsonWithHttps|fetchHtmlAccumCookies|fetchYahooApi/i.test(lines.slice(Math.max(0, s.line - 40), s.line).join('\n') + s.ctx);
  });
  eq(unregistered.map((s) => s.line + ':' + s.kind), [], `${label}：整合區以外沒有未登記的外送呼叫`);
  // company-insight 的缺口已在 SEC-08b 補上（改走 publicFetch；專屬檢查在 scripts/check-sec08b-holes.js）
  // 三個路由：第一次外送之前必須先 checkIntegrationUrl
  for (const rt of [['post', '/api/admin/integrations/test'], ['get', '/api/admin/integrations/sap-inspect'], ['post', '/api/admin/integrations/push/batch']]) {
    const idx = lines.findIndex((l) => l.startsWith('app.' + rt[0] + "('" + rt[1] + "'"));
    let j = idx + 1; while (j < lines.length && !/^app\.(get|post|put|delete|patch)\(/.test(lines[j])) j++;
    const body = idx >= 0 ? lines.slice(idx, j).join('\n') : '';
    const iCheck = body.indexOf('checkIntegrationUrl(');
    const iFetch = Math.min(...['integFetch(', 'ensureAccountAddress('].map((k) => { const p = body.indexOf(k); return p < 0 ? Infinity : p; }));
    ok(idx >= 0 && iCheck >= 0 && iCheck < iFetch, `${label}：${rt[0]} ${rt[1]} 在第一次外送前先 checkIntegrationUrl`);
  }
  ok((srcRaw.match(/checkIntegrationUrl\(tokenEndpoint,/g) || []).length === 1, `${label}：OAuth Token Endpoint 也先檢查`);
  ok(/function integFetch\(url, init\) \{ return SU\.safeFetch\(url, init, INTEGRATION_FETCH_OPTS\); \}/.test(srcRaw), `${label}：integFetch 就是 SU.safeFetch`);
  ok(/INTEGRATION_FETCH_OPTS = \{ ports: SU\.parsePorts\(process\.env\.INTEGRATION_URL_PORTS\)/.test(srcRaw), `${label}：連接埠白名單來自 INTEGRATION_URL_PORTS`);
  ok(/async function checkIntegrationUrl\([^)]*\) \{\s*const v = await SU\.validateUrl\(/.test(srcRaw), `${label}：checkIntegrationUrl 走 SU.validateUrl（含 DNS）`);
}
async function staticGuard() {
  console.log('[S] 靜態守門：server.js 的外送呼叫點');
  staticScan(fs.readFileSync(SERVER_PATH, 'utf8'), 'server.js');
  const envx = fs.readFileSync(path.join(__dirname, '..', '.env.example'), 'utf8');
  ok(/^#?INTEGRATION_URL_PORTS=/m.test(envx), '.env.example 有記載 INTEGRATION_URL_PORTS');
  ok(/lib\/safeUrl/.test(fs.readFileSync(path.join(__dirname, '..', 'lib', 'auditChain.js'), 'utf8')), 'auditChain 共用 lib/safeUrl（單一來源）');
}

// ───────────────────────── [R] 路由端對端 ─────────────────────────
const ROOT = path.resolve(process.env.SAFEURL_ROOT || path.join(__dirname, '..'));
const PORT = parseInt(process.env.SAFEURL_PORT || '3380', 10);
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));
const BASE = `http://localhost:${PORT}`;
const PW = 'Su' + crypto.randomBytes(9).toString('hex') + '!Z';
let child = null;

async function route() {
  if (!(PORT >= 3380 && PORT <= 3389)) { console.log('拒絕：埠須在 3380-3389'); process.exit(2); }
  const guard = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
  if (guard.length) { console.log('拒絕：', ROOT, '已有', guard.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }
  const bcrypt = require(path.join(__dirname, '..', 'node_modules', 'bcryptjs'));
  const { assertPortFree } = require('./_port-free');
  const now = new Date().toISOString();
  const mk = (u, role) => ({ username: u, displayName: u, role, active: true, password: bcrypt.hashSync(PW, 6), createdAt: now, passwordChangedAt: now, mustChangePassword: false });
  const cleanup = () => FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } });
  cleanup();
  fs.writeFileSync(path.join(ROOT, 'auth.json'), JSON.stringify({ users: [mk('_s_adm', 'admin'), mk('_s_u', 'user')] }, null, 2));
  fs.writeFileSync(path.join(ROOT, 'data.json'), JSON.stringify({
    contacts: [], companies: [], opportunities: [], visits: [], contracts: [], receivables: [], quotations: [], callins: [], keyAccounts: [],
    // 已存的整合設定：Base URL 指向 metadata IP（模擬被竄改的設定）；啟用＋Basic，只差 URL 檢查
    integrations: [{ id: 'int-1', type: 'sap-sales-cloud-v2', name: 'SAP', enabled: true, authType: 'basic', baseUrl: 'https://169.254.169.254', username: 'u', password: '' }],
  }, null, 2));
  const env = Object.assign({}, process.env, { PORT: String(PORT), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: BASE });
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
    await login('_s_adm'); await login('_s_u');
    const T = '/api/admin/integrations/test';

    console.log('[R1] POST /api/admin/integrations/test：不合格的 Base URL → 400＋中文訊息（沒有任何對外連線）');
    const cases = [
      ['http://example.com', 'http（只允許 https）', /https/],
      ['https://169.254.169.254', 'metadata IP', /內部網路|本機/],
      ['https://127.0.0.1', '127.0.0.1', /內部網路|本機/],
      ['https://10.0.0.5', '10/8', /內部網路|本機/],
      ['https://192.168.1.10', '192.168/16', /內部網路|本機/],
      ['https://172.16.5.5', '172.16/12', /內部網路|本機/],
      ['https://[::1]', 'IPv6 ::1', /內部網路|本機/],
      ['https://localhost', 'localhost', /內部網路|本機/],
      ['https://user:pass@93.184.216.34', '帶帳密', /帳號密碼/],
      ['https://93.184.216.34:8443', '非 443 連接埠', /連接埠/],
    ];
    for (const [baseUrl, why, re] of cases) {
      const r = await call('_s_adm', 'POST', T, { type: 'sap-sales-cloud-v2', baseUrl, authType: 'basic', username: 'u', password: 'PWSENTINEL9' });
      eq(r.status, 400, `R1 ${why} → 400`);
      ok(r.json && r.json.ok === false && re.test(String(r.json.reason)), `R1 ${why}：訊息是中文原因（${r.json && r.json.reason}）`);
      ok(r.json && r.json.error === r.json.reason, `R1 ${why}：error 與 reason 同文（前端相容）`);
      ok(!/pass@|user:pass|PWSENTINEL9/.test(r.text), `R1 ${why}：回應不回顯帳密或機敏值`);
    }
    console.log('[R2] OAuth Token Endpoint 同樣檢查（Base URL 用公開 IP 字面值，不發 DNS）');
    for (const [tokenEndpoint, why] of [['https://10.0.0.5/oauth/token', '私有 IP'], ['http://93.184.216.34/token', 'http'], ['https://169.254.169.254/token', 'metadata'], ['https://user:pw@93.184.216.34/token', '帶帳密']]) {
      const r = await call('_s_adm', 'POST', T, { type: 'sap-sales-cloud-v2', baseUrl: 'https://93.184.216.34', authType: 'oauth', tokenEndpoint, clientId: 'c', clientSecret: 's' });
      eq(r.status, 400, `R2 Token Endpoint ${why} → 400`);
      ok(r.json && /Token Endpoint/.test(String(r.json.reason)), `R2 ${why}：訊息指出是 Token Endpoint（${r.json && r.json.reason}）`);
    }
    console.log('[R3] 沿用既有行為：缺欄位仍是原本的回應；非管理員仍被擋');
    let r = await call('_s_adm', 'POST', T, { type: 'sap-sales-cloud-v2', baseUrl: '', authType: 'basic' });
    ok(r.status === 200 && r.json && r.json.ok === false && /缺少 Base URL/.test(r.json.reason), 'R3 空 Base URL 仍回原本的 200＋「缺少 Base URL」');
    r = await call('_s_u', 'POST', T, { baseUrl: 'https://169.254.169.254' });
    ok(r.status === 403 || r.status === 401, `R3 非管理員仍被擋（${r.status}）`);
    console.log('[R4] 已存的整合設定被竄改成內網位址：sap-inspect 與 push/batch 也不會連出去');
    r = await call('_s_adm', 'GET', '/api/admin/integrations/sap-inspect');
    ok(r.status === 400 && r.json && /Base URL 不符合安全規定/.test(r.json.error), `R4 sap-inspect → 400（${r.status} ${r.json && r.json.error}）`);
    r = await call('_s_adm', 'POST', '/api/admin/integrations/push/batch', { companyIds: ['x'] });
    ok(r.status === 400 && r.json && /Base URL 不符合安全規定/.test(r.json.error), `R4 push/batch → 400（${r.status} ${r.json && r.json.error}）`);
  } finally {
    if (child) { try { process.kill(child.pid); } catch { /* 已結束 */ } for (let i = 0; i < 20 && child.exitCode === null; i++) await sleep(100); child = null; }
    cleanup();
  }
}

// ───────────────────────── [M] 變異：弄壞一處，檢查必須變紅 ─────────────────────────
const MUT_LIB = [
  ['M1 拿掉 6to4 檢查', "if (g[0] === 0x2002) return blockedV4(", "if (false) return blockedV4("],
  ['M2 拿掉 Teredo 檢查', "if (g[0] === 0x2001 && g[1] === 0) {", "if (false) {"],
  ['M3 不再要求 https', "if (u.protocol !== 'https:') return", "if (false) return"],
  ['M4 DNS 只在「全部」都私有時才擋', "addrs.some((a) => isBlockedIp(a.address))", "addrs.every((a) => isBlockedIp(a.address))"],
  ['M5 轉址的後續各跳不檢查', "if (!v.ok) throw new SafeUrlError(v.reason);", "if (!v.ok && hop === 0) throw new SafeUrlError(v.reason); if (!v.ok) { v.url = new URL(current); v.addrs = [{ address: '93.184.216.34' }]; }"],
  ['M6 連線不釘 IP（改連主機名）', "host: addr, port:", "host: u.hostname, port:"],
  ['M7 不檢查連接埠', "if (u.port && ports) {", "if (false) {"],
  ['M8 不擋網址內的帳密', "if (u.username || u.password) return", "if (false) return"],
  ['M9 localhost 名稱不擋', "return h === 'localhost' || h.endsWith('.localhost') || h === 'localhost.localdomain';", "return false;"],
  ['M10 跨來源轉址仍帶 Authorization', "delete headers.authorization; delete headers.cookie;", ""],
  ['M11 169.254/16 不擋（metadata）', "if (a === 169 && b === 254) return true;", ""],
  ['M12 回應大小不設限', "if (size > maxBytes) return", "if (false) return"],
  ['M13 錯誤訊息洩漏底層訊息', "'連線失敗' + (e && e.code ? '（' + e.code + '）' : '')", "'連線失敗：' + (e && e.message)"],
];
const MUT_SERVER = [
  ['S1 整合區一處改回裸 fetch', "const rAcc = await integFetch(", "const rAcc = await fetch(", 'static'],
  ['S2 sap-inspect 不先檢查網址', "  { const bad = await checkIntegrationUrl(baseUrl, 'Base URL'); if (bad) return res.status(400).json({ error: bad }); }\n  if (config.authType !== 'basic') return res.status(400).json({ error: '僅支援 Basic 認證' });", "  if (config.authType !== 'basic') return res.status(400).json({ error: '僅支援 Basic 認證' });", 'both'],
  ['S3 integFetch 退回全域 fetch', "function integFetch(url, init) { return SU.safeFetch(url, init, INTEGRATION_FETCH_OPTS); }", "function integFetch(url, init) { return fetch(url, init); }", 'static'],
  ['S4 新增一處未登記的裸 fetch', "app.get('/api/admin/integrations', requireAdmin, (req, res) => {", "app.get('/api/ssrf-probe', requireAuth, async (req, res) => { const r = await fetch(req.query.u); res.send(await r.text()); });\napp.get('/api/admin/integrations', requireAdmin, (req, res) => {", 'static'],
  ['S5 連接埠白名單改成不限制', "ports: SU.parsePorts(process.env.INTEGRATION_URL_PORTS)", "ports: null", 'static'],
  ['S6 push/batch 不先檢查網址', "  { const bad = await checkIntegrationUrl(baseUrl, 'Base URL'); if (bad) return res.status(400).json({ error: bad }); }\n  if (config.authType !== 'basic') return res.status(400).json({ error: 'OAuth 批次推送尚未支援，請改用 Basic 認證' });", "  if (config.authType !== 'basic') return res.status(400).json({ error: 'OAuth 批次推送尚未支援，請改用 Basic 認證' });", 'both'],
  ['S7 OAuth Token Endpoint 不檢查', "    { const bad = await checkIntegrationUrl(tokenEndpoint, 'Token Endpoint'); if (bad) return res.status(400).json({ ok: false, error: bad, reason: bad }); }\n", "", 'static'],
  ['S8 checkIntegrationUrl 不做 DNS（只看形狀）', "const v = await SU.validateUrl(String(url || ''), INTEGRATION_FETCH_OPTS);", "const v = SU.checkUrlShape(String(url || ''), INTEGRATION_FETCH_OPTS);", 'static'],
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
  const tmp = fs.mkdtempSync(path.join(require('os').tmpdir(), 'safeurl-mut-'));
  const libSrc = fs.readFileSync(path.join(__dirname, '..', 'lib', 'safeUrl.js'), 'utf8');
  const srvPath = path.join(__dirname, '..', 'server.js');
  const srvSrc = fs.readFileSync(srvPath, 'utf8');
  console.log('[M0] 未變異的基準必須全綠');
  const b = await runChild({ SAFEURL_ONLY: 'unit,static', SAFEURL_MUT: '' });
  ok(b.code === 0, 'M0 基準（unit＋static）全綠 ' + tally(b.out));
  console.log('[M1] lib/safeUrl.js 變異（unit 必須變紅）');
  for (const [name, from, to] of MUT_LIB) {
    if (libSrc.split(from).length !== 2) { ok(false, `${name}：變異目標在 lib/safeUrl.js 找不到或不唯一（測試與程式不同步）`); continue; }
    const p = path.join(tmp, 'safeUrl.js'); fs.writeFileSync(p, libSrc.replace(from, () => to));
    const rc = await runChild({ SAFEURL_ONLY: 'unit', SAFEURL_LIB: p, SAFEURL_MUT: '' });
    ok(rc.code !== 0 && /失敗 [1-9]/.test(rc.out), `${name} → 變紅（${tally(rc.out)}）`);
  }
  console.log('[M2] server.js 變異（靜態守門必須變紅；標 both 的路由端對端也要變紅）');
  for (const [name, from, to, scope] of MUT_SERVER) {
    if (srvSrc.split(from).length !== 2) { ok(false, `${name}：變異目標在 server.js 找不到或不唯一`); continue; }
    const mutated = srvSrc.replace(from, () => to);
    const p = path.join(tmp, 'server.js'); fs.writeFileSync(p, mutated);
    const rs = await runChild({ SAFEURL_ONLY: 'static', SAFEURL_SERVER: p, SAFEURL_MUT: '' });
    ok(rs.code !== 0 && /失敗 [1-9]/.test(rs.out), `${name} → 靜態守門變紅（${tally(rs.out)}）`);
    if (scope === 'both') {
      // 路由端對端：原地換上變異版 server.js，結束一定還原（finally＋行程結束保險）
      const bak = path.join(tmp, 'server.orig.js'); fs.copyFileSync(srvPath, bak);
      const restore = () => { try { fs.copyFileSync(bak, srvPath); } catch { /* 已還原 */ } };
      process.once('exit', restore);
      try {
        fs.writeFileSync(srvPath, mutated);
        const rr = await runChild({ SAFEURL_ONLY: 'route', SAFEURL_PORT: String(PORT + 1), SAFEURL_MUT: '' });
        ok(rr.code !== 0 && /失敗 [1-9]/.test(rr.out), `${name} → 路由端對端也變紅（${tally(rr.out)}）`);
      } finally { restore(); }
      ok(fs.readFileSync(srvPath, 'utf8') === srvSrc, `${name}：server.js 已還原`);
    }
  }
  try { fs.rmSync(tmp, { recursive: true, force: true }); } catch { /* 略 */ }
}

(async () => {
  let code = 0;
  try {
    if (process.env.SAFEURL_MUT === '1') await mutations();
    else {
      if (ONLY.includes('unit')) await unit();
      if (ONLY.includes('static')) await staticGuard();
      if (ONLY.includes('route')) await route();
    }
  } catch (e) { console.log('錯誤：', e.stack || e.message); fail++; failures.push('例外：' + e.message); code = 1; }
  console.log(`\nSSRF 防護檢查：通過 ${pass}　失敗 ${fail}`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); code = 1; }
  process.exit(code);
})();
