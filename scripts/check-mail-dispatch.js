#!/usr/bin/env node
'use strict';
/**
 * 傳輸層與派送器檢查。用法：node scripts/check-mail-dispatch.js（不開伺服器、不連網、不碰 data.json／auth.json；只寫系統暫存資料夾）
 *
 * 涵蓋：
 *   1) transports      訊息驗證、null／log／redirect／graph(P4 stub)、結果正規化、log 傳輸的路徑穿越與 console 靜默
 *   2) 模式護欄         selectTransport 窮舉：只有正規化後恰好是 live 才可能選到真傳輸；log 不碰真傳輸；redirect 一定改寫收件人
 *   3) dispatcher      成功／重複／各種收件人略過／各種傳輸失敗／例外／永不 resolve／熔斷／稽核內容／outbox 內容不含金額與全址
 *   4) drainDue        重寄、租約、STALE／GONE、熔斷、並行不重複、重試用盡
 *   5) 永不 throw      各種亂輸入與會爆炸的相依（隨機 300 組）
 *   6) 整合            JSON 檔案 outbox 全流程；真的 lib/mail/render.js（若尚未就緒則略過並標示）
 *
 * 環境變數 MAIL_CORE_ROOT：要檢查的專案根目錄（變異測試用）。撰寫備註：本檔不寫四位數的 \uXXXX，一律用 cp()／\x..。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const util = require('util');
const assert = require('assert');

const ROOT = process.env.MAIL_CORE_ROOT || path.join(__dirname, '..');
const load = (rel) => require(path.join(ROOT, rel));

// ── 迷你測試框架（章節可為 async）────────────────────────────────────────────
const results = [];
const perf = [];
const sections = [];
let currentSection = '';
function record(name, ok, extra) { results.push({ section: currentSection, name, ok: !!ok, extra: extra === undefined ? '' : String(extra) }); }
const t = (name, ok, extra) => record(name, ok, extra);
function short(v) {
  let s;
  try { s = typeof v === 'string' ? JSON.stringify(v) : util.inspect(v, { depth: 4, breakLength: Infinity }); } catch (e) { s = String(v); }
  return s.length > 220 ? s.slice(0, 220) + '…' : s;
}
function eq(name, actual, expected) {
  let ok = true;
  try { assert.deepStrictEqual(actual, expected); } catch (e) { ok = false; }
  record(name, ok, ok ? '' : 'actual=' + short(actual) + ' expected=' + short(expected));
}
function section(name, fn) { sections.push({ name, fn }); }
function finish() {
  const failed = results.filter((r) => !r.ok);
  const bySec = {};
  results.forEach((r) => { const s = bySec[r.section] || (bySec[r.section] = { p: 0, f: 0 }); if (r.ok) s.p++; else s.f++; });
  Object.keys(bySec).forEach((k) => console.log((bySec[k].f ? 'FAIL ' : 'ok   ') + k + '  通過 ' + bySec[k].p + (bySec[k].f ? '，失敗 ' + bySec[k].f : '')));
  failed.forEach((r) => console.log('  ✗ [' + r.section + '] ' + r.name + (r.extra ? '  → ' + r.extra : '')));
  if (perf.length) console.log('效能：' + perf.join('；'));
  console.log((failed.length ? 'FAILED' : 'PASSED') + '：' + (results.length - failed.length) + ' / ' + results.length);
  process.exit(failed.length ? 1 : 0);
}
async function main() {
  const watchdog = setTimeout(() => { console.log('WATCHDOG：測試超過 120 秒未結束'); process.exit(2); }, 120000);
  for (const s of sections) {
    currentSection = s.name;
    try { await s.fn(); } catch (e) { record('章節執行中例外', false, (e && e.stack) || e); }
  }
  clearTimeout(watchdog);
  finish();
}

// ── 載入被測模組 ────────────────────────────────────────────────────────────
const TR = load('lib/mail/transports.js');
const { createDispatcher } = load('lib/mail/dispatcher.js');
const { createOutbox, memoryAdapter, jsonFileAdapter } = load('lib/mail/outbox.js');
const { getMailConfig } = load('lib/mail/config.js');
const EVN = load('lib/mail/events.js');
const cp = (...n) => String.fromCodePoint(...n);
const T0 = Date.UTC(2026, 9, 8, 3, 0, 0);
const SEC = 1000;
const iso = (ms) => new Date(ms).toISOString();
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
/** 等 promise 但最多 ms；回 {value}|{error}|{hung:true}（避免被測程式卡住時整個測試跟著卡住） */
async function within(p, ms) {
  let timer;
  const guard = new Promise((res) => { timer = setTimeout(() => res({ hung: true }), ms); });
  const r = await Promise.race([Promise.resolve(p).then((value) => ({ value }), (error) => ({ error })), guard]);
  clearTimeout(timer);
  return r;
}

const tmpDirs = [];
function tmpDir() {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), 'mail-dispatch-test-'));
  tmpDirs.push(d);
  return d;
}
process.on('exit', () => { tmpDirs.forEach((d) => { try { fs.rmSync(d, { recursive: true, force: true }); } catch (e) { /* ignore */ } }); });

const FULL_EMAIL_RE = /[A-Za-z0-9._+-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+/;
const VALID_MSG = () => ({ to: ['user1@example.test'], subject: '主旨 Subject', html: '<html><body><p>內文</p></body></html>', text: '內文', tag: 'E1_SUBMIT' });

// ═════════════════════════════════════════════════════════════════════════
section('1 transports', async () => {
  // ── validateMessage ──
  {
    const v = TR.validateMessage(Object.assign(VALID_MSG(), { cc: ['x@example.test'], bcc: 'y@example.test', headers: { 'X-A': '1' }, replyTo: 'z@example.test' }));
    t('validateMessage：合法訊息通過', v.ok === true);
    eq('validateMessage：只留 to／subject／html／text／tag（cc／bcc／headers 全丟）', Object.keys(v.msg).sort(), ['html', 'subject', 'tag', 'text', 'to']);
    eq('validateMessage：收件人正規化＋去重', TR.validateMessage(Object.assign(VALID_MSG(), { to: [' User1@Example.TEST ', 'user1@example.test'] })).msg.to, ['user1@example.test']);
    const bads = {
      '非物件 null': null, '非物件 陣列': [], '非物件 字串': 'x', '非物件 數字': 5,
      'to 不是陣列': { to: 'a@example.test', subject: 's', html: 'h' }, 'to 空陣列': Object.assign(VALID_MSG(), { to: [] }),
      'to 含非法位址': Object.assign(VALID_MSG(), { to: ['a@example.test', 'not-an-email'] }), 'to 含多位址字串': Object.assign(VALID_MSG(), { to: ['a@example.test,b@example.test'] }),
      'to 含換行注入': Object.assign(VALID_MSG(), { to: ['a@example.test\r\nBcc: x@example.test'] }), 'to 超過 50 個': Object.assign(VALID_MSG(), { to: Array.from({ length: 51 }, (_, i) => 'u' + i + '@example.test') }),
      '主旨空': Object.assign(VALID_MSG(), { subject: '' }), '主旨全空白': Object.assign(VALID_MSG(), { subject: '   ' }), '主旨非字串': Object.assign(VALID_MSG(), { subject: 5 }),
      '主旨含 CRLF（標頭注入）': Object.assign(VALID_MSG(), { subject: 'ok\r\nBcc: x@example.test' }), '主旨含 LF': Object.assign(VALID_MSG(), { subject: 'a\nb' }),
      '主旨含 NUL': Object.assign(VALID_MSG(), { subject: 'a\x00b' }), '主旨含 U+2028': Object.assign(VALID_MSG(), { subject: 'a' + cp(0x2028) + 'b' }),
      '主旨含 U+0085': Object.assign(VALID_MSG(), { subject: 'a' + cp(0x85) + 'b' }), '主旨過長': Object.assign(VALID_MSG(), { subject: 's'.repeat(256) }),
      'html 非字串': Object.assign(VALID_MSG(), { html: 5 }), 'text 非字串': Object.assign(VALID_MSG(), { text: {} }), 'html 與 text 都空': Object.assign(VALID_MSG(), { html: '', text: '' }),
      'html 超過 1MB': Object.assign(VALID_MSG(), { html: 'h'.repeat(1024 * 1024 + 1) }),
    };
    Object.keys(bads).forEach((k) => t('validateMessage 拒絕：' + k, TR.validateMessage(bads[k]).ok === false));
    t('validateMessage：只有 html 或只有 text 都可以', TR.validateMessage(Object.assign(VALID_MSG(), { text: undefined })).ok && TR.validateMessage(Object.assign(VALID_MSG(), { html: undefined })).ok);
    t('validateMessage：主旨剛好 255 字可以', TR.validateMessage(Object.assign(VALID_MSG(), { subject: 's'.repeat(255) })).ok);
  }

  // ── normalizeResult／failure ──
  {
    eq('normalizeResult：成功', TR.normalizeResult({ ok: true, providerId: 'p1', junk: 1 }), { ok: true, providerId: 'p1' });
    eq('normalizeResult：成功但無 providerId', TR.normalizeResult({ ok: true }), { ok: true });
    TR.CODES.forEach((c) => {
      const r = TR.normalizeResult({ ok: false, code: c, message: 'm' });
      t('normalizeResult：' + c + ' 的 permanent 依規格（' + (TR.PERMANENT_CODES.indexOf(c) >= 0) + '）', r.ok === false && r.code === c && r.permanent === (TR.PERMANENT_CODES.indexOf(c) >= 0), short(r));
    });
    ['AUTH', 'REJECTED', 'BAD_MESSAGE', 'NOT_CONFIGURED'].forEach((c) => t('normalizeResult：' + c + ' 即使傳輸宣稱 permanent:false 仍視為 permanent', TR.normalizeResult({ ok: false, code: c, permanent: false }).permanent === true));
    t('normalizeResult：可重試碼 permanent:true 時尊重傳輸的宣告', TR.normalizeResult({ ok: false, code: 'SERVER', permanent: true }).permanent === true);
    ['TIMEOUT', 'NETWORK', 'THROTTLED', 'SERVER'].forEach((c) => t('normalizeResult：' + c + ' 預設可重試', TR.normalizeResult({ ok: false, code: c }).permanent === false));
    eq('normalizeResult：retryAfterSec 保留、過大夾在 24 小時、負數與 NaN 丟棄', [TR.normalizeResult({ ok: false, code: 'THROTTLED', retryAfterSec: 30 }).retryAfterSec, TR.normalizeResult({ ok: false, code: 'THROTTLED', retryAfterSec: 1e9 }).retryAfterSec, TR.normalizeResult({ ok: false, code: 'THROTTLED', retryAfterSec: -1 }).retryAfterSec, TR.normalizeResult({ ok: false, code: 'THROTTLED', retryAfterSec: NaN }).retryAfterSec], [30, 86400, undefined, undefined]);
    [undefined, null, 'ok', 5, [], {}, { ok: 'yes' }, { ok: 1 }, { code: 'AUTH' }].forEach((g) => {
      const r = TR.normalizeResult(g);
      t('normalizeResult：亂格式 ' + short(g) + ' → SERVER（可重試）', r.ok === false && r.code === 'SERVER' && r.permanent === false, short(r));
    });
    eq('normalizeResult：未知錯誤碼 → SERVER', TR.normalizeResult({ ok: false, code: 'WAT' }).code, 'SERVER');
    t('normalizeResult：錯誤訊息會經過 scrub（email 與 GUID 被遮）', !/@|0000/.test(TR.normalizeResult({ ok: false, code: 'REJECTED', message: 'no such user a@itts.com.tw tenant 00000000-0000-4000-8000-0000000000aa' }).message));
  }

  // ── null／graph ──
  {
    const n = TR.createNullTransport();
    const r = await n.send(VALID_MSG());
    t('null 傳輸：永遠不寄（NOT_CONFIGURED／permanent）', n.kind === 'null' && r.ok === false && r.code === 'NOT_CONFIGURED' && r.permanent === true, short(r));
    const g = TR.createGraphTransport({ config: getMailConfig({ MAIL_GRAPH_TENANT_ID: 'tenant-x', MAIL_GRAPH_CLIENT_ID: 'client-x', MAIL_GRAPH_CLIENT_SECRET: 'secret-x', MAIL_GRAPH_SENDER: 'send@itts.com.tw' }), fetchImpl: () => { throw new Error('不應該被呼叫'); } });
    const gr = await g.send(VALID_MSG());
    t('graph 傳輸（P4 stub）：即使設定齊全也只回 NOT_CONFIGURED，且沒有呼叫 fetch', g.kind === 'graph' && gr.ok === false && gr.code === 'NOT_CONFIGURED' && gr.permanent === true, short(gr));
  }

  // ── log 傳輸 ──
  {
    const dir = tmpDir();
    let clockT = T0;
    const lt = TR.createLogTransport({ config: { previewDir: path.join(dir, 'prev') }, now: () => clockT });
    const r1 = await lt.send(VALID_MSG());
    const r2 = await lt.send(Object.assign(VALID_MSG(), { subject: '第二封' }));
    eq('log 傳輸：providerId 遞增', [r1, r2], [{ ok: true, providerId: 'log:1' }, { ok: true, providerId: 'log:2' }]);
    const files = fs.readdirSync(path.join(dir, 'prev')).sort();
    t('log 傳輸：每封信寫出 .json／.html／.txt（共 6 個檔）', files.length === 6, files.join(','));
    t('log 傳輸：檔名格式 <ts>-<tag>-<n>.<ext>', files.every((f) => /^20261008T030000000-E1_SUBMIT-\d+\.(json|html|txt)$/.test(f)), files.join(','));
    const j = JSON.parse(fs.readFileSync(path.join(dir, 'prev', files.filter((f) => /\.json$/.test(f))[0]), 'utf8'));
    eq('log 傳輸：.json 含 to 與 subject', [j.to, j.subject, j.providerId], [['user1@example.test'], '主旨 Subject', 'log:1']);
    t('log 傳輸：.html 內容與原信相同', fs.readFileSync(path.join(dir, 'prev', files.filter((f) => /\.html$/.test(f))[0]), 'utf8') === VALID_MSG().html);

    // 路徑穿越
    const parent = tmpDir();
    const sub = path.join(parent, 'prev');
    const lt2 = TR.createLogTransport({ config: { previewDir: sub }, now: () => T0 });
    const evilTags = ['../../evil', '..\\..\\evil', '/etc/passwd', 'C:\\Windows\\x', 'a/b/c', 'x\x00y', '..', '.', '', 'a'.repeat(500), '中文標籤', cp(0x202e) + 'gnp.exe'];
    for (const tag of evilTags) await lt2.send(Object.assign(VALID_MSG(), { tag }));
    const inSub = fs.readdirSync(sub);
    t('log 傳輸：惡意 tag 的檔案全部落在預覽目錄內，且上層目錄沒有多出任何東西', fs.readdirSync(parent).join() === 'prev' && inSub.length === evilTags.length * 3 && inSub.every((f) => /^[A-Za-z0-9_-]+\.(json|html|txt)$/.test(f)), fs.readdirSync(parent).join() + ' / ' + inSub.length);
    t('log 傳輸：檔名長度有上限', inSub.every((f) => f.length < 120));

    // 沒有 previewDir（Vercel）：只回 ok，完全不輸出
    const writes = [];
    const orig = { log: console.log, info: console.info, warn: console.warn, error: console.error, debug: console.debug, out: process.stdout.write, err: process.stderr.write };
    const capture = (...a) => { writes.push(a.join(' ')); };
    console.log = capture; console.info = capture; console.warn = capture; console.error = capture; console.debug = capture;
    process.stdout.write = (c) => { writes.push(String(c)); return true; };
    process.stderr.write = (c) => { writes.push(String(c)); return true; };
    let nr;
    try {
      nr = await TR.createLogTransport({ config: { previewDir: null } }).send(Object.assign(VALID_MSG(), { text: '機密內文 SECRET_CONTENT_55' }));
      const nr2 = await TR.createLogTransport({ config: getMailConfig({ VERCEL: '1' }) }).send(VALID_MSG());
      writes.push(nr2.ok ? '' : 'nr2 failed');
    } finally {
      console.log = orig.log; console.info = orig.info; console.warn = orig.warn; console.error = orig.error; console.debug = orig.debug;
      process.stdout.write = orig.out; process.stderr.write = orig.err;
    }
    t('log 傳輸（previewDir=null／Vercel）：回 ok、providerId 正確', nr.ok === true && nr.providerId === 'log:1', short(nr));
    t('log 傳輸（previewDir=null／Vercel）：完全沒有輸出任何東西到 console／stdout／stderr', writes.filter(Boolean).length === 0, writes.join('|'));
    t('getMailConfig 在 Vercel 下 previewDir 為 null（log 傳輸據此靜默）', getMailConfig({ VERCEL: '1' }).previewDir === null);

    // 相對路徑 + rootDir
    const root = tmpDir();
    await TR.createLogTransport({ config: { previewDir: '.mail-preview' }, rootDir: root, now: () => T0 }).send(VALID_MSG());
    t('log 傳輸：相對 previewDir 以 rootDir 為基準', fs.existsSync(path.join(root, '.mail-preview')) && fs.readdirSync(path.join(root, '.mail-preview')).length === 3);

    // 寫檔失敗不算寄信失敗
    const badFs = Object.assign({}, fs, { mkdirSync() { throw new Error('read-only fs'); } });
    const fr = await TR.createLogTransport({ config: { previewDir: path.join(dir, 'ro') }, fs: badFs }).send(VALID_MSG());
    t('log 傳輸：預覽檔寫入失敗仍回 ok（帶 warning）', fr.ok === true && fr.warning === 'PREVIEW_WRITE_FAILED', short(fr));

    // 亂輸入
    const lt3 = TR.createLogTransport({ config: { previewDir: null } });
    for (const g of [undefined, null, 5, 'x', [], {}, { to: 5 }, (() => { const c = {}; c.self = c; return c; })()]) {
      const rr = await within(lt3.send(g), 500);
      t('log 傳輸：亂輸入 ' + short(g).slice(0, 30) + ' → BAD_MESSAGE 且不 throw', !rr.hung && !rr.error && rr.value.ok === false && rr.value.code === 'BAD_MESSAGE', short(rr));
    }
    const boom = { get to() { throw new Error('getter boom'); } };
    const br = await within(lt3.send(boom), 500);
    t('log 傳輸：會爆炸的 getter → 包成失敗結果（不 throw）', !br.error && !br.hung && br.value.ok === false, short(br));
  }

  // ── redirect 傳輸 ──
  {
    const seen = [];
    const inner = { kind: 'inner', async send(m, o) { seen.push({ m, o }); return { ok: true, providerId: 'inner:1' }; } };
    const cfg = { redirectTo: 'Redir@Example.TEST' };
    const rt = TR.createRedirectTransport({ config: cfg, inner });
    const orig = { to: ['real.person@itts.com.tw', 'second.person@itts.com.tw'], subject: '【簽核通知】QU-1 專案', html: '<!doctype html><html><head><title>x</title></head><body style="margin:0"><p>內文</p></body></html>', text: '內文第一行\n第二行', tag: 'E1_SUBMIT', cc: ['cc.person@itts.com.tw'], bcc: ['bcc.person@itts.com.tw'], headers: { 'X-Evil': '1' } };
    const r = await rt.send(orig, { signal: 'SIG', timeouts: { totalMs: 1 } });
    t('redirect：回傳內層結果', r.ok === true && r.providerId === 'inner:1', short(r));
    const m = seen[0].m;
    eq('redirect：收件人被改寫成 redirectTo（已正規化小寫）', m.to, ['redir@example.test']);
    t('redirect：主旨加 [測試轉送] 前綴', m.subject === TR.REDIRECT_SUBJECT_PREFIX + orig.subject && /^\[測試轉送\] /.test(m.subject), m.subject);
    t('redirect：html 在 <body> 之後插入橫幅，含遮罩後的原收件人', /<body style="margin:0"><div [^>]*>【測試轉送】原收件人：r\*\*\*@itts\.com\.tw、s\*\*\*@itts\.com\.tw（redirect 模式，僅供測試）<\/div><p>內文<\/p>/.test(m.html), m.html);
    t('redirect：text 最上方加註原收件人（遮罩）', m.text.indexOf('【測試轉送】原收件人：r***@itts.com.tw、s***@itts.com.tw（redirect 模式，僅供測試）\n\n內文第一行') === 0, m.text);
    t('redirect：內層收到的內容完全不含原收件人的完整位址', !/real\.person|second\.person/.test(JSON.stringify(m)), JSON.stringify(m).slice(0, 200));
    eq('redirect：傳給內層的物件只有 to／subject／html／text／tag（cc／bcc／headers 不外流）', Object.keys(m).sort(), ['html', 'subject', 'tag', 'text', 'to']);
    t('redirect：原樣轉交 sendOpts（signal、timeouts）', seen[0].o && seen[0].o.signal === 'SIG' && seen[0].o.timeouts.totalMs === 1);
    t('redirect：不修改呼叫端傳入的物件', orig.to.length === 2 && orig.subject === '【簽核通知】QU-1 專案' && !/測試轉送/.test(orig.html));

    // html 沒有 body／只有 text／只有 html
    seen.length = 0;
    await rt.send(Object.assign({}, orig, { html: '<p>no body tag</p>' }));
    t('redirect：html 沒有 <body> 時橫幅放在最前面', /^<div [^>]*>【測試轉送】/.test(seen[0].m.html) && /<p>no body tag<\/p>$/.test(seen[0].m.html));
    seen.length = 0;
    await rt.send(Object.assign({}, orig, { html: '' }));
    t('redirect：沒有 html 時 html 維持空字串（只加 text 橫幅）', seen[0].m.html === '' && /^【測試轉送】/.test(seen[0].m.text));
    seen.length = 0;
    await rt.send(Object.assign({}, orig, { html: '<HTML><BODY CLASS="x">大寫標籤</BODY></HTML>' }));
    t('redirect：<BODY> 大寫標籤也能找到', /<BODY CLASS="x"><div /.test(seen[0].m.html), seen[0].m.html);

    // redirectTo 缺失／不合法 → NOT_CONFIGURED，且不呼叫 inner
    for (const bad of [undefined, null, '', '   ', 'not-an-email', 'a@b@c.test', 'a@example.test,b@example.test', 'x@example.test\r\nBcc: y@example.test', 5, {}]) {
      seen.length = 0;
      const rr = await TR.createRedirectTransport({ config: { redirectTo: bad }, inner }).send(orig);
      t('redirect：redirectTo=' + short(bad).slice(0, 30) + ' → NOT_CONFIGURED（permanent）且絕不呼叫內層', rr.ok === false && rr.code === 'NOT_CONFIGURED' && rr.permanent === true && seen.length === 0, short(rr));
    }
    const noInner = await TR.createRedirectTransport({ config: cfg }).send(orig);
    t('redirect：沒有內層傳輸 → NOT_CONFIGURED', noInner.ok === false && noInner.code === 'NOT_CONFIGURED');
    seen.length = 0;
    const badMsg = await rt.send(Object.assign({}, orig, { subject: 'x\r\nBcc: evil@example.test' }));
    t('redirect：主旨含換行 → BAD_MESSAGE 且不呼叫內層', badMsg.code === 'BAD_MESSAGE' && seen.length === 0);
    const badTo = await rt.send(Object.assign({}, orig, { to: ['nope'] }));
    t('redirect：收件人不合法 → BAD_MESSAGE 且不呼叫內層', badTo.code === 'BAD_MESSAGE' && seen.length === 0);

    // 內層出問題
    const throwing = TR.createRedirectTransport({ config: cfg, inner: { send() { throw new Error('sync boom'); } } });
    const rejecting = TR.createRedirectTransport({ config: cfg, inner: { send() { return Promise.reject(new Error('async boom')); } } });
    const garbage = TR.createRedirectTransport({ config: cfg, inner: { send() { return Promise.resolve('what'); } } });
    const failing = TR.createRedirectTransport({ config: cfg, inner: { send() { return Promise.resolve({ ok: false, code: 'THROTTLED', retryAfterSec: 9, message: 'slow' }); } } });
    eq('redirect：內層同步 throw → NETWORK', (await throwing.send(orig)).code, 'NETWORK');
    eq('redirect：內層 reject → NETWORK', (await rejecting.send(orig)).code, 'NETWORK');
    eq('redirect：內層回亂格式 → SERVER', (await garbage.send(orig)).code, 'SERVER');
    const fr = await failing.send(orig);
    t('redirect：內層失敗結果原樣通過（含 retryAfterSec）', fr.code === 'THROTTLED' && fr.retryAfterSec === 9 && fr.permanent === false, short(fr));

    // 多次呼叫不累積狀態
    seen.length = 0;
    await rt.send(orig); await rt.send(orig);
    t('redirect：連續兩次呼叫的橫幅各只有一個', seen.length === 2 && seen.every((s) => (s.m.html.match(/原收件人/g) || []).length === 1));
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('2 模式護欄 selectTransport', async () => {
  const REAL_SENT = [];
  const REAL = { kind: 'REAL', async send(m) { REAL_SENT.push(m); return { ok: true, providerId: 'real:1' }; } };
  const root = tmpDir();
  const deps = () => ({ realTransport: REAL, rootDir: root, now: () => T0 });
  const isLive = (raw) => typeof raw === 'string' && /^[ \t\r\n]*[Ll][Ii][Vv][Ee][ \t\r\n]*$/.test(raw);
  const modeOf = (raw) => {
    if (typeof raw !== 'string') return 'off';
    const m = /^[ \t\r\n]*([A-Za-z]{2,10})[ \t\r\n]*$/.exec(raw);
    if (!m) return 'off';
    const l = m[1].toLowerCase();
    return ['off', 'log', 'redirect', 'live'].indexOf(l) >= 0 ? l : 'off';
  };
  const RAW = [
    'live', 'LIVE', 'Live', 'lIvE', ' live', 'live ', ' live ', '\tlive\n', 'live\r\n', '\r\nLIVE\t',
    'liv', 'livee', 'xlive', 'live1', 'live;', 'live,log', 'live log', 'l ive', 'l1ve', '1ive', 'live!', '"live"', "'live'", 'live\x00', '\x00live',
    cp(0xff4c, 0xff49, 0xff56, 0xff45), cp(0x217c) + 'ive', 'l' + cp(0x456) + 've', cp(0x39a) + 'ive', 'L' + cp(0x130) + 'VE', 'live' + cp(0x200b), cp(0xa0) + 'live', 'live' + cp(0xa0), cp(0xfeff) + 'live', cp(0x202e) + 'live',
    'li' + cp(0x212a) + 'e', 'LIVE' + cp(0x2028), 'live' + cp(0x3000),
    'true', 'TRUE', '1', 'yes', 'on', 'y', 'production', 'prod', 'enable', 'enabled', 'live-mode', 'LIVE_MODE', 'go-live', 'delivery', 'olive', 'alive',
    'off', 'OFF', ' Off ', 'log', 'LOG', 'redirect', 'REDIRECT', ' Redirect ', 'redirectt', 'redir', 'logg', 'offf',
    '', ' ', '   ', '\n', 'null', 'undefined', 'false', '0', 'none', 'dev', 'test', 'demo', 'staging',
    null, undefined, 0, 1, true, false, {}, [], ['live'], { toString() { return 'live'; } }, () => 'live', Symbol('live'), 12345,
  ];

  // 逐值：手工組出的 config
  for (const raw of RAW) {
    const name = typeof raw === 'string' ? JSON.stringify(raw) : util.inspect(raw);
    const cfg = { mode: raw, redirectTo: 'redir@example.test' };
    const tr = TR.selectTransport(cfg, deps());
    const wantLive = isLive(raw);
    t('手工 config.mode=' + name + '：' + (wantLive ? '選到真傳輸' : '不會選到真傳輸'), (tr === REAL) === wantLive, 'kind=' + (tr && tr.kind));
    const wantKind = wantLive ? 'REAL' : { off: 'null', log: 'log', redirect: 'redirect' }[modeOf(raw)];
    t('手工 config.mode=' + name + '：傳輸種類＝' + wantKind, tr.kind === wantKind, 'kind=' + tr.kind);
  }
  // 逐值：經由 getMailConfig(env)
  for (const raw of RAW) {
    if (typeof raw !== 'string') continue;
    const name = JSON.stringify(raw);
    const cfg = getMailConfig({ MAIL_MODE: raw, MAIL_REDIRECT_TO: 'redir@example.test', MAIL_PREVIEW_DIR: path.join(root, 'p') });
    const tr = TR.selectTransport(cfg, deps());
    const wantLive = isLive(raw);
    t('env MAIL_MODE=' + name + '：' + (wantLive ? '選到真傳輸' : '不會選到真傳輸'), (tr === REAL) === wantLive, 'kind=' + tr.kind + ' mode=' + cfg.mode);
  }

  // 環境窮舉：MAIL_MODE × MAIL_REDIRECT_TO × Graph 設定 × VERCEL
  {
    const redirects = [undefined, 'redir@example.test', 'bad redirect'];
    const graphs = [{}, { MAIL_GRAPH_TENANT_ID: 'tenant-x', MAIL_GRAPH_CLIENT_ID: 'client-x', MAIL_GRAPH_CLIENT_SECRET: 'secret-x', MAIL_GRAPH_SENDER: 'send@itts.com.tw' }];
    const vercels = [undefined, '1'];
    let combos = 0;
    let wrongLive = 0;
    let wrongNonLive = 0;
    let redirectWithoutTarget = 0;
    for (const raw of RAW.filter((x) => typeof x === 'string')) {
      for (const rd of redirects) for (const g of graphs) for (const vc of vercels) {
        const env = Object.assign({ MAIL_MODE: raw }, g);
        if (rd !== undefined) env.MAIL_REDIRECT_TO = rd;
        if (vc !== undefined) env.VERCEL = vc;
        const cfg = getMailConfig(env);
        const tr = TR.selectTransport(cfg, deps());
        combos += 1;
        if (isLive(raw) && tr !== REAL) wrongLive += 1;
        if (!isLive(raw) && tr === REAL) wrongNonLive += 1;
        if (modeOf(raw) === 'redirect' && tr.kind === 'redirect' && !(rd === 'redir@example.test')) redirectWithoutTarget += 1;
      }
    }
    t('環境窮舉（' + combos + ' 組）：live 值一定選到真傳輸', wrongLive === 0, wrongLive);
    t('環境窮舉（' + combos + ' 組）：非 live 值絕不選到真傳輸', wrongNonLive === 0, wrongNonLive);
    t('環境窮舉：redirect 沒有有效目標時不會維持 redirect（降為 log）', redirectWithoutTarget === 0, redirectWithoutTarget);
  }

  // 隨機字串
  {
    let seed = 20261008;
    const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
    const alpha = ['l', 'i', 'v', 'e', 'L', 'I', 'V', 'E', 'o', 'f', 'g', 'r', 'd', 'c', 't', ' ', '\t', '\n', '\r', cp(0xa0), cp(0x200b), cp(0xff4c), cp(0x130), cp(0x212a), '1', '-', '_', '\x00'];
    let bad = 0;
    let liveHits = 0;
    for (let i = 0; i < 6000; i++) {
      let s = '';
      const n = 1 + Math.floor(rnd() * 8);
      for (let k = 0; k < n; k++) s += alpha[Math.floor(rnd() * alpha.length)];
      if (i % 5 === 0) s = ['live', 'LIVE', ' live', 'live '][i % 4] + (rnd() < 0.5 ? '' : alpha[Math.floor(rnd() * alpha.length)]);
      const tr = TR.selectTransport({ mode: s, redirectTo: 'redir@example.test' }, deps());
      if ((tr === REAL) !== isLive(s)) bad += 1;
      if (tr === REAL) liveHits += 1;
    }
    t('隨機 6000 個 mode 字串：選到真傳輸 ⇔ 正規化後恰為 live（不一致 ' + bad + ' 個，live 命中 ' + liveHits + ' 次）', bad === 0 && liveHits > 100);
  }

  // 各模式的實際行為（真的送一封看看）
  {
    const msg = () => ({ to: ['real.person@itts.com.tw'], subject: 's', html: '<body>h</body>', text: 't', tag: 'T' });
    for (const raw of ['off', 'OFF', 'log', ' LOG ', '', 'garbage', 'liv']) {
      REAL_SENT.length = 0;
      const tr = TR.selectTransport(getMailConfig({ MAIL_MODE: raw, MAIL_REDIRECT_TO: 'redir@example.test', MAIL_PREVIEW_DIR: path.join(root, 'p2') }), deps());
      const r = await tr.send(msg());
      t('mode=' + JSON.stringify(raw) + '：真傳輸一次都沒被呼叫', REAL_SENT.length === 0, short(r));
    }
    for (const raw of ['redirect', 'REDIRECT', ' Redirect\n']) {
      REAL_SENT.length = 0;
      const tr = TR.selectTransport(getMailConfig({ MAIL_MODE: raw, MAIL_REDIRECT_TO: 'redir@example.test' }), deps());
      const r = await tr.send(msg());
      t('mode=' + JSON.stringify(raw) + '：真傳輸收到的收件人只有 redirectTo', r.ok === true && REAL_SENT.length === 1 && REAL_SENT[0].to.length === 1 && REAL_SENT[0].to[0] === 'redir@example.test' && !/real\.person/.test(JSON.stringify(REAL_SENT[0])), short(REAL_SENT));
    }
    for (const raw of ['live', 'LIVE', ' live ']) {
      REAL_SENT.length = 0;
      const tr = TR.selectTransport(getMailConfig({ MAIL_MODE: raw }), deps());
      await tr.send(msg());
      t('mode=' + JSON.stringify(raw) + '：真傳輸收到原收件人', REAL_SENT.length === 1 && REAL_SENT[0].to[0] === 'real.person@itts.com.tw');
    }
    // redirect 設定被破壞（手工 config）：寧可不寄也不能寄給原收件人
    for (const bad of [undefined, '', 'nope', null]) {
      REAL_SENT.length = 0;
      const tr = TR.selectTransport({ mode: 'redirect', redirectTo: bad }, deps());
      const r = await tr.send(msg());
      t('手工 config：mode=redirect 但 redirectTo=' + short(bad) + ' → NOT_CONFIGURED，真傳輸未被呼叫', r.code === 'NOT_CONFIGURED' && REAL_SENT.length === 0);
    }
    // log 不管有沒有真傳輸都不碰它
    REAL_SENT.length = 0;
    await TR.selectTransport({ mode: 'log', previewDir: null }, deps()).send(msg());
    t('mode=log：即使傳入真傳輸也不使用', REAL_SENT.length === 0);
  }

  // 其他護欄
  {
    for (const c of [undefined, null, {}, 'live', 5, [], { mode: undefined }]) t('config=' + short(c) + ' → null 傳輸', TR.selectTransport(c, deps()).kind === 'null');
    const live1 = TR.selectTransport({ mode: 'live' }, {});
    t('live 但沒有真傳輸 → 使用 graph（P4 stub，回 NOT_CONFIGURED）', live1.kind === 'graph' && (await live1.send(VALID_MSG())).code === 'NOT_CONFIGURED');
    t('live 但 realTransport 沒有 send 函式 → 退回 graph stub', TR.selectTransport({ mode: 'live' }, { realTransport: {} }).kind === 'graph' && TR.selectTransport({ mode: 'live' }, { realTransport: 'x' }).kind === 'graph');
    t('redirect 且 realTransport 不合法 → 內層改用 log（不會爆）', TR.selectTransport({ mode: 'redirect', redirectTo: 'r@example.test', previewDir: null }, { realTransport: {} }).kind === 'redirect');
    const lt = TR.createLogTransport({ config: { previewDir: null } });
    t('可以重用傳入的 logTransport', TR.selectTransport({ mode: 'log' }, { logTransport: lt }) === lt);
    const seen = [];
    const spy = { send: async (m) => { seen.push(m); return { ok: true }; } };
    await TR.selectTransport({ mode: 'redirect', redirectTo: 'r@example.test' }, { realTransport: spy }).send(VALID_MSG());
    t('傳入的 realTransport 在 redirect 模式只會收到改寫後的信', seen.length === 1 && seen[0].to[0] === 'r@example.test');
  }
});

// ═════════════════════════════════════════════════════════════════════════
// 派送器測試的共用設備
// ═════════════════════════════════════════════════════════════════════════
const USERS = () => [
  { username: 'sales1', role: 'user', email: 'sales1@itts.com.tw', displayName: 'Sales One' },
  { username: 'mgr1', role: 'manager1', email: 'mgr1@itts.com.tw', nickname: 'M1' },
  { username: 'mgr2', role: 'manager1', email: 'mgr2@itts.com.tw' },
  { username: 'gm1', role: 'executive', email: 'gm1@itts.com.tw' },
  { username: 'gm2', role: 'executive', email: 'gm2@itts.com.tw' },
  { username: 'cons1', role: 'consult_manager_south', email: 'cons1@itts.com.tw' },
  { username: 'sec1', role: 'secretary', email: 'sec1@itts.com.tw' },
  { username: 'noemail1', role: 'user' },
  { username: 'badmail1', role: 'user', email: 'not an email' },
  { username: 'ext1', role: 'user', email: 'ext1@example.test' },
  { username: 'off1', role: 'user', email: 'off1@itts.com.tw', active: false },
];
const AMOUNT_NEEDLES = ['1234567800', '12,345,678', '12345678', '41.04', '506000000', '5,060,000'];
// 客戶名、專案名、業務名不可進稽核／outbox。（actorLabel＝操作者稱呼，是規格內的 outbox 欄位，所以測試用不同的字串 'Actor Person'）
const NAME_NEEDLES = ['Project Alpha', 'Acme Test Co', 'Owner Person'];

function mkEv(over) {
  return Object.assign({
    type: 'E1_SUBMIT', quoteId: 'q-1', quoteNo: 'QU-000001-001', projectName: 'Project Alpha', company: 'Acme Test Co', ownerLabel: 'Owner Person',
    step: { level: 1, label: '一級主管' },
    numbers: { revenueCents: 1234567800, gpCents: 506000000, marginText: '41.04%', marginPct: 41.04, tierLevel: 1, tierLabel: '一級主管' },
    actor: { label: 'Actor Person' }, at: '2026-10-08T03:00:00.000Z', stepKey: '2026-10-08T03:00:00.000Z#1',
  }, over || {});
}
const rcp = (names, kind) => names.map((n) => ({ username: n, kind: kind || 'mgr1' }));

/** 建立一個完整的測試世界（假時鐘、記憶體 outbox、可替換行為的假傳輸、假 render、稽核收集器）。 */
function mkWorld(over) {
  const o = over || {};
  const clock = { t: T0, now() { return this.t; }, advance(ms) { this.t += ms; } };
  const config = getMailConfig(Object.assign({ MAIL_MODE: 'live', MAIL_PREVIEW_DIR: path.join(tmpDir(), 'prev') }, o.env || {}));
  config.timeouts = o.timeouts || { connectMs: 20, totalMs: 80 };
  if (o.config) Object.assign(config, o.config);
  const adapter = o.adapter || memoryAdapter();
  const outbox = createOutbox(adapter, { now: () => clock.t, config });
  const sent = [];
  const behavior = { fn: null, inflight: 0, maxInflight: 0, signals: [] };
  const transport = {
    kind: 'FAKE',
    async send(msg, opts) {
      sent.push(msg);
      behavior.signals.push(opts && opts.signal);
      behavior.inflight += 1;
      behavior.maxInflight = Math.max(behavior.maxInflight, behavior.inflight);
      try {
        return behavior.fn ? await behavior.fn(msg, opts, sent.length) : { ok: true, providerId: 'fake:' + sent.length };
      } finally { behavior.inflight -= 1; }
    },
  };
  const users = o.users || USERS();
  const logs = [];
  const renders = [];
  const calls = { getUsers: 0 };
  const render = o.render || ((ev, viewer) => {
    renders.push({ ev, viewer });
    return { subject: '【簽核通知】' + ev.quoteNo + ' 專案', html: '<html><body><p>hi ' + viewer.label + ' ' + viewer.kind + '</p></body></html>', text: 'hi ' + viewer.label, meta: {} };
  });
  const world = {
    clock, config, adapter, outbox, sent, behavior, users, logs, renders, calls, transport,
    ev: mkEv(),
    rebuildEv: null,
  };
  const dispatcher = createDispatcher({
    config, outbox, transport: o.noTransport ? undefined : transport,
    getUsers: o.getUsers || (() => { calls.getUsers += 1; return users; }),
    isStillValid: o.isStillValid,
    writeLog: o.writeLog || ((...a) => { logs.push(a); }),
    render, now: () => clock.t,
    // 預設的 rebuild 模擬「依單據目前狀態重建事件」：事件身分（型別、單據、stepKey）取自工作的去重鍵，其餘取自 world.ev
    rebuild: o.rebuild === undefined ? ((job) => ({ ev: Object.assign({}, world.rebuildEv || world.ev, { type: job.type, quoteId: job.quoteId, quoteNo: job.quoteNo, stepKey: job.dedupeKey.split(':').slice(3).join(':') }) })) : o.rebuild,
    auxTimeoutMs: o.auxTimeoutMs || 100,
    concurrency: o.concurrency, budgetMs: o.budgetMs, rootDir: o.rootDir,
  });
  world.dispatch = dispatcher.dispatch;
  world.drainDue = dispatcher.drainDue;
  world.records = async () => (await outbox.list({ limit: 500 })).rows;
  world.dump = () => adapter.dump();
  return world;
}
const sumKeys = ['cancelled', 'errors', 'failed', 'queued', 'sent', 'skipped'];

// ═════════════════════════════════════════════════════════════════════════
section('3 dispatcher：成功、重複、模式', async () => {
  // 成功路徑
  {
    const w = mkWorld();
    const s = await w.dispatch(w.ev, rcp(['mgr1']), { actorUsername: 'sales1', operatorLabel: 'Sales One' });
    eq('成功：Summary 形狀', Object.keys(s).sort(), sumKeys);
    eq('成功：sent=1，其餘為空', [s.sent, s.queued, s.skipped, s.failed, s.cancelled, s.errors], [1, 0, [], [], 0, 0]);
    eq('成功：真傳輸收到一封，收件人是 mgr1 的真實位址', w.sent.map((m) => m.to), [['mgr1@itts.com.tw']]);
    eq('成功：標籤＝事件型別', w.sent[0].tag, 'E1_SUBMIT');
    eq('成功：render 收到 viewer（暱稱優先）與 kind', w.renders.map((r) => r.viewer), [{ username: 'mgr1', label: 'M1', kind: 'mgr1' }]);
    const recs = await w.records();
    eq('成功：outbox 一筆 sent，遮罩位址', recs.map((r) => [r.status, r.toUser, r.toMasked, r.attempts, r.type, r.meta]), [['sent', 'mgr1', 'm***@itts.com.tw', 1, 'E1_SUBMIT', { level: 1, kind: 'mgr1' }]]);
    eq('成功：不寫稽核', w.logs, []);
    eq('成功：actorLabel 記在紀錄上', recs[0].actorLabel, 'Actor Person');
  }
  // 重複觸發
  {
    const w = mkWorld();
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    const before = w.renders.length;
    const s2 = await w.dispatch(w.ev, rcp(['mgr1']), {});
    eq('重複觸發：不重寄、記 DUPLICATE', [w.sent.length, s2.sent, s2.skipped], [1, 0, [{ username: 'mgr1', reason: 'DUPLICATE' }]]);
    eq('重複觸發：不重新渲染', w.renders.length, before);
    eq('重複觸發：outbox 仍只有 1 筆', (await w.records()).length, 1);
    eq('重複觸發：不寫稽核', w.logs, []);
    const s3 = await w.dispatch(mkEv({ stepKey: '2026-10-08T04:00:00.000Z#1' }), rcp(['mgr1']), {});
    eq('不同 stepKey（再次送簽）：視為新事件，會寄', [s3.sent, w.sent.length], [1, 2]);
    const s4 = await w.dispatch(mkEv({ type: 'E3_NEXT_STEP' }), rcp(['mgr1']), {});
    eq('不同事件型別：視為新事件', [s4.sent, w.sent.length], [1, 3]);
    const s5 = await w.dispatch(w.ev, rcp(['mgr2']), {});
    eq('同事件、不同收件人：各自獨立', [s5.sent, w.sent.length], [1, 4]);
  }
  // 多收件人／kind 傳遞
  {
    const w = mkWorld();
    const s = await w.dispatch(w.ev, [{ username: 'gm1', kind: 'gm' }, { username: 'gm2', kind: 'gm' }, { username: 'sec1', kind: 'secretary' }, { username: 'cons1', kind: 'consultant' }], { actorUsername: 'sales1' });
    eq('多收件人：全部寄出', [s.sent, w.sent.length], [4, 4]);
    eq('多收件人：render 的 kind 來自呼叫端', w.renders.map((r) => r.viewer.kind).sort(), ['consultant', 'gm', 'gm', 'secretary']);
    eq('多收件人：outbox 記錄 kind', (await w.records()).map((r) => r.meta.kind).sort(), ['consultant', 'gm', 'gm', 'secretary']);
  }
  // 同一 username 出現兩次、第一個 kind 為準
  {
    const w = mkWorld();
    const s = await w.dispatch(w.ev, [{ username: 'mgr1', kind: 'mgr1' }, { username: 'mgr1', kind: 'consultant' }], {});
    eq('同帳號出現兩次：只寄一封、第二次記 DUP', [s.sent, s.skipped, w.renders[0].viewer.kind], [1, [{ username: 'mgr1', reason: 'DUP' }], 'mgr1']);
  }
  // off
  {
    const w = mkWorld({ env: { MAIL_MODE: 'off' } });
    const s = await w.dispatch(w.ev, rcp(['mgr1', 'gm1', 'noemail1', 'ghost']), { actorUsername: 'sales1' });
    eq('off：全部 skipped MODE_OFF', s.skipped.map((x) => x.reason), ['MODE_OFF', 'MODE_OFF', 'MODE_OFF', 'MODE_OFF']);
    eq('off：沒寄、沒渲染、沒查帳號、沒稽核', [w.sent.length, w.renders.length, w.calls.getUsers, w.logs.length, s.sent, s.errors], [0, 0, 0, 0, 0, 0]);
    const recs = await w.records();
    eq('off：outbox 記 4 筆 skipped/MODE_OFF（無遮罩位址）', recs.map((r) => [r.status, r.skipReason, r.toMasked]).sort(), [['skipped', 'MODE_OFF', ''], ['skipped', 'MODE_OFF', ''], ['skipped', 'MODE_OFF', ''], ['skipped', 'MODE_OFF', '']]);
    const s2 = await w.dispatch(w.ev, rcp(['mgr1', 'mgr1']), {});
    eq('off：重複的 username 記 DUP', s2.skipped.map((x) => x.reason), ['MODE_OFF', 'DUP']);
    eq('off：drainDue 什麼都不做', [(await w.drainDue({})).errors, w.sent.length], [0, 0]);
    // 先有一筆 pending，再把模式切成 off：drainDue 不可以繼續寄
    const wp = mkWorld();
    wp.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await wp.dispatch(wp.ev, rcp(['mgr1']), {});
    wp.behavior.fn = null;
    wp.config.mode = 'off';
    wp.clock.advance(3600 * SEC);
    const dp = await wp.drainDue({});
    eq('先 pending 後切成 off：drainDue 不領取、不寄', [dp.sent, dp.queued, wp.sent.length, (await wp.records())[0].status], [0, 0, 1, 'pending']);
    wp.config.mode = 'live';
    eq('切回 live 後 drainDue 照常把它寄出', [(await wp.drainDue({})).sent, wp.sent.length], [1, 2]);
    const w2 = mkWorld({ env: { MAIL_MODE: undefined } });
    const s3 = await w2.dispatch(w2.ev, rcp(['mgr1']), {});
    eq('未設定 MAIL_MODE（預設 off）：不寄', [s3.skipped[0].reason, w2.sent.length], ['MODE_OFF', 0]);
  }
  // log
  {
    const prev = path.join(tmpDir(), 'prev');
    const w = mkWorld({ env: { MAIL_MODE: 'log', MAIL_PREVIEW_DIR: prev }, rootDir: tmpDir() });
    const s = await w.dispatch(w.ev, rcp(['mgr1', 'gm1']), { actorUsername: 'sales1' });
    eq('log：真傳輸一次都沒被呼叫', w.sent.length, 0);
    eq('log：Summary 為 skipped/MODE_LOG，sent=0', [s.sent, s.skipped.map((x) => x.reason), s.failed], [0, ['MODE_LOG', 'MODE_LOG'], []]);
    const files = fs.existsSync(prev) ? fs.readdirSync(prev) : [];
    t('log：預覽檔有寫出（每封 3 個）', files.length === 6, files.length);
    eq('log：outbox 記為 skipped/MODE_LOG（可由 admin requeue）', (await w.records()).map((r) => [r.status, r.skipReason]), [['skipped', 'MODE_LOG'], ['skipped', 'MODE_LOG']]);
    eq('log：不寫稽核', w.logs, []);
    // 之後切到 redirect 並 requeue
    const rec = (await w.records())[0];
    w.config.mode = 'redirect';
    w.config.redirectTo = 'redir@example.test';
    const rq = await w.outbox.requeue(rec.id);
    t('log 紀錄可 requeue（不會被去重擋住）', rq.ok === true);
    const d = await w.drainDue({});
    eq('requeue 後切 redirect：drainDue 會重寄到測試信箱', [d.sent, w.sent.map((m) => m.to)], [1, [['redir@example.test']]]);
  }
  // redirect
  {
    const w = mkWorld({ env: { MAIL_MODE: 'redirect', MAIL_REDIRECT_TO: 'Redir@Example.test' } });
    const s = await w.dispatch(w.ev, rcp(['mgr1', 'gm1', 'sec1']), { actorUsername: 'sales1' });
    eq('redirect：全部寄出', [s.sent, w.sent.length], [3, 3]);
    t('redirect：真傳輸收到的每一封收件人都只有 redirectTo', w.sent.every((m) => m.to.length === 1 && m.to[0] === 'redir@example.test'), short(w.sent.map((m) => m.to)));
    t('redirect：真傳輸收到的內容不含任何真實收件位址', !/mgr1@|gm1@|sec1@/.test(JSON.stringify(w.sent)));
    t('redirect：主旨有 [測試轉送]，內文有橫幅（遮罩位址）', w.sent.every((m) => /^\[測試轉送\] /.test(m.subject) && /原收件人：[mgs]\*\*\*@itts\.com\.tw/.test(m.html) && /原收件人：/.test(m.text)));
    t('redirect：outbox 記的是「預定收件人」的遮罩位址（不是測試信箱）', (await w.records()).every((r) => /^[mgs]\*\*\*@itts\.com\.tw$/.test(r.toMasked)));
    t('redirect：render 仍以真實收件人的 kind／稱呼渲染', w.renders.map((r) => r.viewer.username).sort().join() === 'gm1,mgr1,sec1');
    // redirect 無目標 → config 降為 log
    const w2 = mkWorld({ env: { MAIL_MODE: 'redirect' } });
    t('redirect 沒設 MAIL_REDIRECT_TO：config 降為 log', w2.config.mode === 'log');
    const s2 = await w2.dispatch(w2.ev, rcp(['mgr1']), {});
    eq('redirect 沒設目標：不會寄給任何人', [w2.sent.length, s2.sent], [0, 0]);
    // 手工破壞 config：mode=redirect 但 redirectTo 清空（繞過 getMailConfig）
    const w3 = mkWorld({ env: { MAIL_MODE: 'redirect', MAIL_REDIRECT_TO: 'redir@example.test' } });
    w3.config.redirectTo = '';
    const s3 = await w3.dispatch(w3.ev, rcp(['mgr1']), {});
    eq('config 被破壞（mode=redirect 且 redirectTo 空）：NOT_CONFIGURED 失敗，真傳輸未被呼叫', [w3.sent.length, s3.failed, s3.sent], [0, [{ username: 'mgr1', code: 'NOT_CONFIGURED' }], 0]);
    t('config 被破壞：寫了失敗稽核', w3.logs.length === 1 && w3.logs[0][0] === 'QUOTE_MAIL_FAILED');
  }
  // live 與運行中模式切換
  {
    const w = mkWorld({ env: { MAIL_MODE: 'live' } });
    const s = await w.dispatch(w.ev, rcp(['mgr1']), {});
    eq('live：寄到真實位址', [s.sent, w.sent[0].to], [1, ['mgr1@itts.com.tw']]);
    w.config.mode = 'garbage';
    const s2 = await w.dispatch(mkEv({ stepKey: 'k2' }), rcp(['mgr1']), {});
    eq('執行中 config.mode 被改成亂值：下一次 dispatch 視為 off', [s2.skipped[0].reason, w.sent.length], ['MODE_OFF', 1]);
    w.config.mode = 'LIVE ';
    const s3 = await w.dispatch(mkEv({ stepKey: 'k3' }), rcp(['mgr1']), {});
    eq("config.mode='LIVE '：正規化後視為 live", [s3.sent, w.sent.length], [1, 2]);
    // 沒給真傳輸的 live：graph stub → NOT_CONFIGURED
    const w2 = mkWorld({ env: { MAIL_MODE: 'live' }, noTransport: true });
    const s4 = await w2.dispatch(w2.ev, rcp(['mgr1']), {});
    eq('live 但沒有真傳輸：NOT_CONFIGURED（permanent）失敗，不是靜默成功', [s4.sent, s4.failed], [0, [{ username: 'mgr1', code: 'NOT_CONFIGURED' }]]);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('3 dispatcher：收件人略過與稽核', async () => {
  const w = mkWorld();
  const names = ['mgr1', 'noemail1', 'badmail1', 'ext1', 'off1', 'ghost', 'sales1', 'mgr1'];
  const s = await w.dispatch(w.ev, rcp(names), { actorUsername: 'sales1', operatorLabel: 'Sales One' });
  eq('略過：只有 mgr1 寄出', [s.sent, w.sent.map((m) => m.to[0])], [1, ['mgr1@itts.com.tw']]);
  const reasons = {};
  s.skipped.forEach((x) => { reasons[x.username] = x.reason; });
  eq('略過：各帳號的原因', reasons, { noemail1: 'NO_EMAIL', badmail1: 'BAD_EMAIL', ext1: 'DOMAIN_NOT_ALLOWED', off1: 'INACTIVE', ghost: 'UNKNOWN_USER', sales1: 'ACTOR', mgr1: 'DUP' });
  const recs = await w.records();
  eq('略過：有 outbox 紀錄的是 mgr1(sent) 與五個值得注意的略過；ACTOR／DUP 不入紀錄', recs.map((r) => r.toUser).sort(), ['badmail1', 'ext1', 'ghost', 'mgr1', 'noemail1', 'off1']);
  eq('略過：紀錄狀態與原因', recs.filter((r) => r.status === 'skipped').map((r) => r.toUser + ':' + r.skipReason).sort(), ['badmail1:BAD_EMAIL', 'ext1:DOMAIN_NOT_ALLOWED', 'ghost:UNKNOWN_USER', 'noemail1:NO_EMAIL', 'off1:INACTIVE']);
  eq('稽核：5 筆 QUOTE_MAIL_SKIPPED（ACTOR／DUP 不寫）', w.logs.map((l) => l[0] + ' ' + l[3]).sort(), [
    'QUOTE_MAIL_SKIPPED type=E1_SUBMIT to=badmail1 reason=BAD_EMAIL', 'QUOTE_MAIL_SKIPPED type=E1_SUBMIT to=ext1 reason=DOMAIN_NOT_ALLOWED',
    'QUOTE_MAIL_SKIPPED type=E1_SUBMIT to=ghost reason=UNKNOWN_USER', 'QUOTE_MAIL_SKIPPED type=E1_SUBMIT to=noemail1 reason=NO_EMAIL', 'QUOTE_MAIL_SKIPPED type=E1_SUBMIT to=off1 reason=INACTIVE',
  ]);
  t('稽核：operator 與 target（單號）參數正確', w.logs.every((l) => l[1] === 'Sales One' && l[2] === 'QU-000001-001'));
  // 略過的收件人不影響其他人；再次 dispatch 同事件
  const s2 = await w.dispatch(w.ev, rcp(names), { actorUsername: 'sales1' });
  t('重複觸發：略過的收件人也不會重複寫稽核', w.logs.length === 5, w.logs.length);
  eq('重複觸發：略過的收件人不會變成寄出', [s2.sent, w.sent.length], [0, 1]);
  // 帳號之後補上 email：同一事件不會自動補寄（去重），但 admin 可 requeue → 之後由 drainDue 補寄
  w.users.find((u) => u.username === 'noemail1').email = 'noemail1@itts.com.tw';
  const skippedRec = recs.find((r) => r.toUser === 'noemail1');
  const rq = await w.outbox.requeue(skippedRec.id);
  t('補上 email 後 admin 可 requeue 被略過的那筆', rq.ok === true);
  const d = await w.drainDue({});
  eq('requeue 後 drainDue 補寄', [d.sent, w.sent.map((m) => m.to[0]).sort()], [1, ['mgr1@itts.com.tw', 'noemail1@itts.com.tw']]);
  // 帳號含 @ 時稽核會遮罩
  const w2 = mkWorld({ users: [{ username: 'who@example.test', role: 'user' }] });
  await w2.dispatch(w2.ev, rcp(['who@example.test']), {});
  t('稽核：帳號名稱像 email 時被遮罩', w2.logs.length === 1 && !/who@example/.test(w2.logs[0][3]) && /w\*\*\*@example\.test/.test(w2.logs[0][3]), short(w2.logs));
});

// ═════════════════════════════════════════════════════════════════════════
section('3 dispatcher：傳輸失敗分類與重試', async () => {
  const retryable = [['TIMEOUT', 60], ['NETWORK', 60], ['SERVER', 60], ['THROTTLED', 60]];
  for (const [code, delay] of retryable) {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code, message: 'boom', permanent: false });
    const s = await w.dispatch(w.ev, rcp(['mgr1']), {});
    eq(code + '：可重試 → queued，不算 failed，不寫稽核', [s.queued, s.failed, s.sent, w.logs.length], [1, [], 0, 0]);
    const r = (await w.records())[0];
    eq(code + '：紀錄 pending、排 ' + delay + ' 秒後、錯誤碼已記', [r.status, r.nextAttemptAt, r.lastErrorCode, r.attempts], ['pending', iso(T0 + delay * SEC), code, 1]);
  }
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'THROTTLED', message: 'slow down', retryAfterSec: 120 });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    eq('THROTTLED + Retry-After 120：下次排在 120 秒後（優先於 60 秒排程）', (await w.records())[0].nextAttemptAt, iso(T0 + 120 * SEC));
  }
  for (const code of ['AUTH', 'REJECTED', 'BAD_MESSAGE', 'NOT_CONFIGURED']) {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code, message: 'no' });
    const s = await w.dispatch(w.ev, rcp(['mgr1']), {});
    eq(code + '：permanent → 立刻 failed，寫失敗稽核', [s.failed, s.queued, w.logs.map((l) => l[0] + ' ' + l[3])], [[{ username: 'mgr1', code }], 0, ['QUOTE_MAIL_FAILED type=E1_SUBMIT to=mgr1 code=' + code]]);
    eq(code + '：紀錄 failed', [(await w.records())[0].status, (await w.records())[0].nextAttemptAt], ['failed', null]);
  }
  // 亂格式與未知碼
  for (const g of [undefined, null, 'ok', 5, {}, { ok: 'yes' }, { ok: false, code: 'WAT' }]) {
    const w = mkWorld();
    w.behavior.fn = async () => g;
    const s = await w.dispatch(w.ev, rcp(['mgr1']), {});
    t('傳輸回亂格式 ' + short(g) + '：當成 SERVER 可重試', s.queued === 1 && s.failed.length === 0 && (await w.records())[0].lastErrorCode === 'SERVER', short(s));
  }
  // 傳輸拋例外／拒絕
  {
    const w = mkWorld();
    w.behavior.fn = () => { throw new Error('sync boom with secret_token=abcdefghijklmnopqrstuvwxyz0123456789'); };
    const s = await w.dispatch(w.ev, rcp(['mgr1']), {});
    eq('傳輸拋例外：包成 NETWORK 失敗（可重試），dispatch 不 throw', [s.queued, s.errors], [1, 0]);
    t('傳輸拋例外：例外訊息裡的密鑰不會進 outbox', w.dump().indexOf('abcdefghijklmnopqrstuvwxyz0123456789') < 0);
    const w2 = mkWorld();
    w2.behavior.fn = () => Promise.reject(new Error('async boom'));
    eq('傳輸 reject：NETWORK', [(await w2.dispatch(w2.ev, rcp(['mgr1']), {})).queued, (await w2.records())[0].lastErrorCode], [1, 'NETWORK']);
    const w3 = mkWorld();
    w3.behavior.fn = () => 'not a promise';
    eq('傳輸回傳非 Promise 的亂值：SERVER', (await (async () => { await w3.dispatch(w3.ev, rcp(['mgr1']), {}); return (await w3.records())[0].lastErrorCode; })()), 'SERVER');
  }
  // 永不 resolve → TIMEOUT，且 abort signal 被觸發
  {
    const w = mkWorld({ timeouts: { connectMs: 10, totalMs: 60 } });
    w.behavior.fn = () => new Promise(() => { /* 永不 resolve */ });
    const t0 = Date.now();
    const r = await within(w.dispatch(w.ev, rcp(['mgr1']), {}), 3000);
    const ms = Date.now() - t0;
    t('傳輸永不 resolve：dispatch 在逾時內返回（不卡住）', !r.hung && !r.error, 'ms=' + ms);
    t('傳輸永不 resolve：約 totalMs 後返回（60ms，容忍 40–1500）', ms >= 40 && ms < 1500, 'ms=' + ms);
    eq('傳輸永不 resolve：視為 TIMEOUT，可重試', [r.value.queued, r.value.failed, (await w.records())[0].lastErrorCode], [1, [], 'TIMEOUT']);
    t('傳輸永不 resolve：已通知傳輸 abort（signal.aborted）', w.behavior.signals[0] && w.behavior.signals[0].aborted === true);
  }
  // 逾時後傳輸才完成：結果被忽略，不影響已記錄的 TIMEOUT
  {
    const w = mkWorld({ timeouts: { connectMs: 10, totalMs: 40 } });
    w.behavior.fn = () => new Promise((res) => setTimeout(() => res({ ok: true, providerId: 'late' }), 120));
    const s = await w.dispatch(w.ev, rcp(['mgr1']), {});
    await sleep(200);
    eq('傳輸比逾時慢：以 TIMEOUT 為準（之後才完成的成功結果被忽略）', [s.sent, s.queued, (await w.records())[0].status], [0, 1, 'pending']);
  }
  // 完整重試流程：首次 + 3 次重試，共 4 次，之後 failed 並寫一次稽核
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'SERVER', message: 'HTTP 503' });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    const steps = [60, 300, 900];
    for (let i = 0; i < 3; i++) {
      w.clock.advance(steps[i] * SEC - 1);
      eq('重試 ' + (i + 1) + '：差 1 毫秒不會領取', (await w.drainDue({})).sent + w.sent.length, w.sent.length);
      w.clock.advance(1);
      const d = await w.drainDue({});
      eq('重試 ' + (i + 1) + '：到期後重寄一次', [w.sent.length, d.queued + d.failed.length], [i + 2, 1]);
    }
    const rec = (await w.records())[0];
    eq('重試用盡：共 4 次嘗試，最終 failed', [w.sent.length, rec.status, rec.attempts, rec.lastErrorCode], [4, 'failed', 4, 'SERVER']);
    eq('重試用盡：只在最終失敗時寫一筆稽核', w.logs.map((l) => l[0] + ' ' + l[3]), ['QUOTE_MAIL_FAILED type=E1_SUBMIT to=mgr1 code=SERVER']);
    w.clock.advance(1000 * 86400);
    eq('failed 不會再被重寄', [(await w.drainDue({})).sent, w.sent.length], [0, 4]);
    // admin 重送
    w.behavior.fn = null;
    await w.outbox.requeue(rec.id);
    const dd = await w.drainDue({});
    eq('failed 經 requeue 重送成功', [dd.sent, (await w.records())[0].status, w.sent.length], [1, 'sent', 5]);
  }
  // 重試中途成功
  {
    const w = mkWorld();
    let n = 0;
    w.behavior.fn = async () => (++n < 3 ? { ok: false, code: 'NETWORK', message: 'reset' } : { ok: true, providerId: 'ok' });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    w.clock.advance(60 * SEC);
    await w.drainDue({});
    w.clock.advance(300 * SEC);
    const d = await w.drainDue({});
    eq('重試第 3 次成功：sent=1，紀錄 sent，總共嘗試 3 次，全程沒有稽核', [d.sent, (await w.records())[0].status, w.sent.length, w.logs.length], [1, 'sent', 3, 0]);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('3 dispatcher：isStillValid、render、帳號資料、相依爆炸', async () => {
  // isStillValid
  {
    const seenArgs = [];
    const w = mkWorld({ isStillValid: (job, ev) => { seenArgs.push([job.status, job.toUser, ev.quoteNo]); return false; } });
    const s = await w.dispatch(w.ev, rcp(['mgr1', 'gm1']), {});
    eq('isStillValid=false：cancelled，不寄', [s.cancelled, s.sent, w.sent.length, s.failed], [2, 0, 0, []]);
    eq('isStillValid 收到的 job 是 sending、帶 toUser，ev 是原事件', seenArgs.sort(), [['sending', 'gm1', 'QU-000001-001'], ['sending', 'mgr1', 'QU-000001-001']]);
    eq('isStillValid=false：紀錄 cancelled／STALE', (await w.records()).map((r) => [r.status, r.skipReason]), [['cancelled', 'STALE'], ['cancelled', 'STALE']]);
    eq('isStillValid=false：不渲染、不寫稽核', [w.renders.length, w.logs.length], [0, 0]);
  }
  {
    const w = mkWorld({ isStillValid: async () => true });
    eq('isStillValid=true（async）：照寄', (await w.dispatch(w.ev, rcp(['mgr1']), {})).sent, 1);
  }
  for (const [name, fn] of [['undefined', () => undefined], ['1', () => 1], ["'yes'", () => 'yes'], ['{}', () => ({})], ['null', () => null]]) {
    const w = mkWorld({ isStillValid: fn });
    const s = await w.dispatch(w.ev, rcp(['mgr1']), {});
    t('isStillValid 回 ' + name + '（非 true）：保守取消，不寄', s.cancelled === 1 && s.sent === 0 && w.sent.length === 0, short(s));
  }
  {
    const w = mkWorld({ isStillValid: () => { throw new Error('db down'); } });
    const s = await w.dispatch(w.ev, rcp(['mgr1']), {});
    eq('isStillValid 拋例外：不寄、稍後重試（STALE_CHECK），不是取消也不是失敗', [s.queued, s.cancelled, s.failed, w.sent.length, (await w.records())[0].lastErrorCode], [1, 0, [], 0, 'STALE_CHECK']);
    const w2 = mkWorld({ isStillValid: () => new Promise(() => { /* hang */ }) });
    const r = await within(w2.dispatch(w2.ev, rcp(['mgr1']), {}), 2000);
    t('isStillValid 永不 resolve：以輔助逾時返回，不寄、稍後重試', !r.hung && r.value.queued === 1 && w2.sent.length === 0, short(r));
  }
  // 沒提供 isStillValid → 視為有效
  eq('未提供 isStillValid：視為有效', (await mkWorld().dispatch(mkEv(), rcp(['mgr1']), {})).sent, 1);

  // render
  {
    const w = mkWorld({ render: () => { const e = new Error('bad template'); e.code = 'BAD_KIND'; throw e; } });
    const s = await w.dispatch(w.ev, rcp(['mgr1', 'gm1']), {});
    eq('render 拋例外：每位收件人 failed/RENDER（permanent），不重試', [s.failed.map((x) => x.code), s.sent, s.queued, w.sent.length], [['RENDER', 'RENDER'], 0, 0, 0]);
    eq('render 失敗：寫稽核（每人一筆）', w.logs.map((l) => l[0] + ' ' + l[3]).sort(), ['QUOTE_MAIL_FAILED type=E1_SUBMIT to=gm1 code=RENDER', 'QUOTE_MAIL_FAILED type=E1_SUBMIT to=mgr1 code=RENDER']);
    eq('render 失敗：紀錄 failed 且只嘗試 1 次', (await w.records()).map((r) => [r.status, r.attempts, r.lastErrorCode]), [['failed', 1, 'RENDER'], ['failed', 1, 'RENDER']]);
  }
  for (const [name, val] of [['undefined', undefined], ['null', null], ['字串', 'x'], ['缺 subject', { html: 'h' }], ['缺內文', { subject: 's' }], ['subject 非字串', { subject: 5, html: 'h' }]]) {
    const w = mkWorld({ render: () => val });
    const s = await w.dispatch(w.ev, rcp(['mgr1']), {});
    t('render 回傳 ' + name + '：failed/RENDER', s.failed.length === 1 && s.failed[0].code === 'RENDER' && w.sent.length === 0, short(s));
  }
  {
    const w = mkWorld({ render: () => ({ subject: '', html: '<p>h</p>', text: 't' }) });
    const s = await w.dispatch(w.ev, rcp(['mgr1']), {});
    eq('render 回空主旨：派送器自己驗證 → BAD_MESSAGE（permanent），不送到傳輸層', [s.failed, w.sent.length], [[{ username: 'mgr1', code: 'BAD_MESSAGE' }], 0]);
    const w2 = mkWorld({ render: () => ({ subject: 'x\r\nBcc: evil@example.test', html: '<p>h</p>' }) });
    const s2 = await w2.dispatch(w2.ev, rcp(['mgr1']), {});
    t('render 回含換行的主旨（標頭注入）：派送器擋下 BAD_MESSAGE，傳輸層一封都沒收到', s2.failed.length === 1 && s2.failed[0].code === 'BAD_MESSAGE' && w2.sent.length === 0, short(s2));
    const w3 = mkWorld({ render: () => ({ subject: 's', html: '<p>h</p>' + 'x'.repeat(1024 * 1024) }) });
    const s3 = await w3.dispatch(w3.ev, rcp(['mgr1']), {});
    t('render 回超過 1MB 的信：BAD_MESSAGE', s3.failed.length === 1 && s3.failed[0].code === 'BAD_MESSAGE' && w3.sent.length === 0, short(s3));
    const w4 = mkWorld({ render: () => ({ subject: 's', html: '<p>h</p>', text: 't', to: ['evil@example.test'], cc: ['evil@example.test'] }) });
    await w4.dispatch(w4.ev, rcp(['mgr1']), {});
    t('render 回傳多餘欄位（to／cc）不會影響收件人：傳輸層只看到驗證後的五個欄位', w4.sent.length === 1 && Object.keys(w4.sent[0]).sort().join() === 'html,subject,tag,text,to' && w4.sent[0].to.join() === 'mgr1@itts.com.tw', short(w4.sent[0]));
  }
  {
    const w = mkWorld({ render: () => new Promise(() => { /* hang */ }) });
    const r = await within(w.dispatch(w.ev, rcp(['mgr1']), {}), 2000);
    t('render 永不 resolve：以輔助逾時返回，記 RENDER 失敗', !r.hung && r.value.failed.length === 1 && r.value.failed[0].code === 'RENDER', short(r));
  }
  {
    const w = mkWorld({ render: async (ev, viewer) => ({ subject: 's ' + viewer.username, html: 'h', text: 't' }) });
    eq('render 可以是 async', (await w.dispatch(w.ev, rcp(['mgr1']), {})).sent, 1);
    const w2 = mkWorld({ render: (ev, viewer) => { if (viewer.username === 'gm1') throw new Error('only gm1 fails'); return { subject: 's', html: 'h', text: 't' }; } });
    const s2 = await w2.dispatch(w2.ev, rcp(['mgr1', 'gm1', 'gm2']), {});
    eq('一位收件人 render 失敗不影響其他人', [s2.sent, s2.failed.map((f) => f.username)], [2, ['gm1']]);
  }

  // getUsers 出狀況：事件先入列（email 留空），之後由 drainDue 補寄
  for (const [name, fn] of [['拋例外', () => { throw new Error('db down'); }], ['reject', () => Promise.reject(new Error('db down'))], ['回 null', () => null], ['永不 resolve', () => new Promise(() => { /* hang */ })]]) {
    const w = mkWorld({ getUsers: fn });
    const r = await within(w.dispatch(w.ev, rcp(['mgr1', 'gm1', 'sales1']), { actorUsername: 'sales1' }), 2000);
    t('getUsers ' + name + '：dispatch 不卡住、不 throw', !r.hung && !r.error, short(r));
    if (r.value) {
      eq('getUsers ' + name + '：事件入列待補（queued=2，errors=1），不寄', [r.value.queued, r.value.errors, w.sent.length, r.value.skipped.map((x) => x.reason)], [2, 1, 0, ['ACTOR']]);
      eq('getUsers ' + name + '：紀錄 pending、遮罩位址空', (await w.records()).map((x) => [x.status, x.toMasked]), [['pending', ''], ['pending', '']]);
    }
  }
  {
    let broken = true;
    const w = mkWorld({ getUsers: () => (broken ? null : USERS()) });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    broken = false;
    const d = await w.drainDue({});
    eq('帳號資料恢復後 drainDue 補寄', [d.sent, w.sent[0] && w.sent[0].to], [1, ['mgr1@itts.com.tw']]);
  }

  // writeLog 爆炸不影響派送
  for (const [name, wl] of [['拋例外', () => { throw new Error('log down'); }], ['reject', () => Promise.reject(new Error('log down'))], ['永不 resolve', () => new Promise(() => { /* hang */ })]]) {
    const w = mkWorld({ writeLog: wl });
    const r = await within(w.dispatch(w.ev, rcp(['mgr1', 'noemail1', 'ghost']), {}), 3000);
    t('writeLog ' + name + '：dispatch 照常完成', !r.hung && !r.error && r.value.sent === 1, short(r));
  }
  { const w = mkWorld({ writeLog: undefined }); delete w.logs; t('未提供 writeLog 也能運作', (await w.dispatch(w.ev, rcp(['mgr1', 'noemail1']), {})).sent === 1); }

  // outbox 壞掉
  {
    const bad = Object.assign({}, memoryAdapter(), { insert: async () => { throw new Error('db down'); } });
    const w = mkWorld({ adapter: bad });
    const r = await within(w.dispatch(w.ev, rcp(['mgr1', 'gm1']), {}), 2000);
    t('outbox insert 失敗：dispatch 不 throw、errors 計數、沒寄', !r.error && !r.hung && r.value.errors === 2 && r.value.sent === 0 && w.sent.length === 0, short(r));
    t('outbox insert 失敗：嘗試寫 QUOTE_MAIL_FAILED 稽核', w.logs.length === 2 && w.logs.every((l) => l[0] === 'QUOTE_MAIL_FAILED' && /code=INTERNAL/.test(l[3])), short(w.logs));
    const bad2 = Object.assign({}, memoryAdapter(), { claimById: async () => { throw new Error('db down'); } });
    const w2 = mkWorld({ adapter: bad2 });
    const r2 = await within(w2.dispatch(w2.ev, rcp(['mgr1']), {}), 2000);
    t('outbox claim 失敗：不 throw、errors 計數', !r2.error && r2.value.errors === 1 && w2.sent.length === 0, short(r2));
    const bad3 = Object.assign({}, memoryAdapter(), { breakerLoad: async () => { throw new Error('db down'); }, breakerSave: async () => { throw new Error('db down'); } });
    const w3 = mkWorld({ adapter: bad3 });
    const r3 = await within(w3.dispatch(w3.ev, rcp(['mgr1']), {}), 2000);
    t('熔斷狀態讀寫失敗：視為關閉，照常寄送', !r3.error && r3.value.sent === 1 && r3.value.errors === 0, short(r3));
    const bad4 = Object.assign({}, memoryAdapter(), { markSent: async () => { throw new Error('db down'); } });
    const w4 = mkWorld({ adapter: bad4 });
    const r4 = await within(w4.dispatch(w4.ev, rcp(['mgr1']), {}), 2000);
    t('信寄出但 markSent 失敗：不 throw，errors 計數，仍算 sent', !r4.error && r4.value.sent === 1 && r4.value.errors === 1, short(r4));
  }
  // 事件不合法
  {
    const w = mkWorld();
    const bads = [undefined, null, 'x', 5, {}, mkEv({ type: 'E9' }), mkEv({ quoteId: '../x' }), mkEv({ stepKey: '' }), mkEv({ quoteNo: '' }), mkEv({ at: 'yesterday' }), mkEv({ numbers: { revenueCents: 1.5 } })];
    for (const ev of bads) {
      const r = await within(w.dispatch(ev, rcp(['mgr1']), {}), 1000);
      t('事件不合法 ' + short(ev).slice(0, 40) + '：failed/BAD_EVENT，不寄', !r.error && !r.hung && r.value.errors === 1 && r.value.failed.length === 1 && r.value.failed[0].code === 'BAD_EVENT' && w.sent.length === 0, short(r));
    }
    t('事件不合法：寫稽核但 detail 不含事件內容', w.logs.length === bads.length && w.logs.every((l) => l[0] === 'QUOTE_MAIL_FAILED' && /code=BAD_EVENT$/.test(l[3])), short(w.logs[0]));
  }
  // 收件人清單怪異
  {
    const w = mkWorld();
    const weird = [null, undefined, 5, 'mgr1', {}, { username: 5 }, { username: '' }, { username: {} }, [], { kind: 'mgr1' }];
    const r = await within(w.dispatch(w.ev, weird, {}), 1000);
    t('收件人清單全是怪東西：不 throw、不寄', !r.error && !r.hung && r.value.sent === 0 && w.sent.length === 0, short(r));
    const r2 = await within(w.dispatch(w.ev, 'mgr1', {}), 1000);
    const r3 = await within(w.dispatch(w.ev, undefined, undefined), 1000);
    t('收件人不是陣列／沒有 opts：不 throw', !r2.error && !r3.error && r2.value.sent === 0 && r3.value.sent === 0);
    const w2 = mkWorld();
    const r4 = await w2.dispatch(w2.ev, Array.from({ length: 500 }, (_, i) => ({ username: 'u' + i, kind: 'mgr1' })), {});
    t('收件人超過 200 位：只處理前 200 位', r4.skipped.length === 200 && w2.sent.length === 0, r4.skipped.length);
  }
  // 控制字元帳號
  {
    const w = mkWorld({ users: [{ username: 'bad\nuser', role: 'user', email: 'x@itts.com.tw' }, { username: 'ok1', role: 'user', email: 'ok1@itts.com.tw' }] });
    const r = await within(w.dispatch(w.ev, rcp(['bad\nuser', 'ok1']), {}), 1000);
    t('帳號含換行：該收件人記 BAD_KEY 錯誤，其他人照常寄出', !r.error && r.value.sent === 1 && w.sent.length === 1 && w.sent[0].to[0] === 'ok1@itts.com.tw', short(r));
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('3 dispatcher：熔斷、並行、時間預算', async () => {
  // 連續 TIMEOUT → 熔斷
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT', message: 'slow' });
    for (let i = 0; i < 5; i++) await w.dispatch(mkEv({ stepKey: 'k' + i }), rcp(['mgr1']), {});
    eq('5 次連續逾時：真傳輸被呼叫 5 次', w.sent.length, 5);
    eq('5 次連續逾時：熔斷開啟（600 秒）', [(await w.outbox.breaker.state()).open, (await w.outbox.breaker.state()).until], [true, iso(T0 + 600 * SEC)]);
    const s = await w.dispatch(mkEv({ stepKey: 'k-after' }), rcp(['mgr1', 'gm1']), {});
    eq('熔斷開啟時：不呼叫傳輸，入列後留 pending（queued）', [w.sent.length, s.queued, s.sent], [5, 2, 0]);
    const held = (await w.records()).filter((r) => r.dedupeKey.indexOf('k-after') >= 0);
    eq('熔斷開啟時：nextAttemptAt＝熔斷結束時間，attempts=0', held.map((r) => [r.status, r.nextAttemptAt, r.attempts]), [['pending', iso(T0 + 600 * SEC), 0], ['pending', iso(T0 + 600 * SEC), 0]]);
    const d = await w.drainDue({});
    eq('熔斷開啟時 drainDue：直接返回（breakerOpen），什麼都不領', [d.breakerOpen, d.sent, d.queued, w.sent.length], [true, 0, 0, 5]);
    // 熔斷結束後恢復
    w.clock.advance(600 * SEC);
    w.behavior.fn = null;
    const d2 = await w.drainDue({ limit: 20 });
    t('熔斷結束後 drainDue 恢復運作，並把留下的信寄出', d2.sent >= 2 && !d2.breakerOpen, short(d2));
    eq('寄成功後熔斷重置', await w.outbox.breaker.state(), { open: false, until: null, failures: 0, halfOpen: false });
  }
  // AUTH 一次就開
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'AUTH', message: '401' });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    eq('AUTH 失敗一次：熔斷立刻開啟', (await w.outbox.breaker.state()).open, true);
    const s = await w.dispatch(mkEv({ stepKey: 'k2' }), rcp(['gm1']), {});
    eq('熔斷開啟後不再打傳輸', [w.sent.length, s.queued], [1, 1]);
  }
  // 與傳輸健康無關的失敗不觸發熔斷
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'REJECTED', message: 'no such user' });
    for (let i = 0; i < 8; i++) await w.dispatch(mkEv({ stepKey: 'k' + i }), rcp(['mgr1']), {});
    eq('連續 REJECTED：不會熔斷', [(await w.outbox.breaker.state()).open, w.sent.length], [false, 8]);
  }
  // 成功重置計數
  {
    const w = mkWorld();
    let n = 0;
    w.behavior.fn = async () => (n++ % 5 === 4 ? { ok: true } : { ok: false, code: 'TIMEOUT' });
    for (let i = 0; i < 20; i++) await w.dispatch(mkEv({ stepKey: 'k' + i }), rcp(['mgr1']), {});
    eq('每 5 次有 1 次成功：永遠湊不到連續 5 次失敗，不熔斷', (await w.outbox.breaker.state()).open, false);
  }
  // 並行上限
  {
    const w = mkWorld({ concurrency: 3 });
    const many = [];
    const users = [];
    for (let i = 0; i < 12; i++) { users.push({ username: 'p' + i, role: 'user', email: 'p' + i + '@itts.com.tw' }); many.push('p' + i); }
    w.users.push(...users);
    w.behavior.fn = async () => { await sleep(15); return { ok: true }; };
    const s = await w.dispatch(w.ev, rcp(many), {});
    eq('12 位收件人、並行上限 3：全部寄出，同時進行的傳輸呼叫最多 3 個', [s.sent, w.behavior.maxInflight], [12, 3]);
    eq('並行寄送：每位收件人各一封、沒有重複', new Set(w.sent.map((m) => m.to[0])).size, 12);
  }
  // 時間預算：超過預算的收件人留 pending，不遺失
  {
    const w = mkWorld({ budgetMs: 50, concurrency: 1 });
    const users = [];
    const many = [];
    for (let i = 0; i < 6; i++) { users.push({ username: 'p' + i, role: 'user', email: 'p' + i + '@itts.com.tw' }); many.push('p' + i); }
    w.users.push(...users);
    // 假時鐘不會自己走，所以讓傳輸每次把假時鐘往前撥 30ms
    w.behavior.fn = async () => { w.clock.advance(30); return { ok: true }; };
    const s = await w.dispatch(w.ev, rcp(many), {});
    t('時間預算：前幾位寄出、超過預算後的留 pending（不遺失）', s.sent >= 1 && s.sent < 6 && s.queued === 6 - s.sent && s.sent + s.queued === 6, short(s));
    const d = await w.drainDue({ limit: 20 });
    eq('時間預算：剩下的由 drainDue 補寄', [d.sent, w.sent.length], [6 - s.sent, 6]);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('3 dispatcher：稽核與 outbox 內容不洩漏', async () => {
  const w = mkWorld();
  // 各種結果混在一起：成功、可重試、permanent、略過、取消
  w.behavior.fn = async (msg) => {
    const to = msg.to[0];
    if (to.startsWith('gm1')) return { ok: false, code: 'REJECTED', message: 'recipient gm1@itts.com.tw rejected for Project Alpha / Acme Test Co / Sales One amount 1234567800 tenant 00000000-0000-4000-8000-0000000000aa Bearer abcdefghijklmnopqrstuv' };
    if (to.startsWith('gm2')) return { ok: false, code: 'TIMEOUT', message: 'timeout while sending 【簽核通知】QU-000001-001 專案 to gm2@itts.com.tw' };
    return { ok: true, providerId: 'p' };
  };
  await w.dispatch(w.ev, rcp(['mgr1', 'gm1', 'gm2', 'noemail1', 'ghost', 'sec1']), { actorUsername: 'sales1', operatorLabel: 'Sales One' });
  const dump = w.dump();
  // 稽核的 operator 參數（第 2 個）本來就是操作者名稱，是稽核欄位的一部分；這裡檢查的是 action／target／detail
  const auditText = w.logs.map((l) => [l[0], l[2], l[3]].join('|')).join('\n');
  t('稽核有內容（失敗與略過）', w.logs.length >= 3, w.logs.length);
  t('稽核：action 只會是 QUOTE_MAIL_FAILED／QUOTE_MAIL_SKIPPED', w.logs.every((l) => l[0] === 'QUOTE_MAIL_FAILED' || l[0] === 'QUOTE_MAIL_SKIPPED'));
  t('稽核 detail 格式固定（type= to= code=|reason=）', w.logs.every((l) => /^type=E\d_[A-Z_]+( to=\S+)? (code|reason)=[A-Z_]+$/.test(l[3])), short(w.logs.map((l) => l[3])));
  t('稽核不含任何 email 全址', !FULL_EMAIL_RE.test(auditText), (auditText.match(FULL_EMAIL_RE) || [''])[0]);
  AMOUNT_NEEDLES.forEach((n) => t('稽核不含金額字串「' + n + '」', auditText.indexOf(n) < 0));
  NAME_NEEDLES.forEach((n) => t('稽核不含「' + n + '」', auditText.indexOf(n) < 0));
  t('稽核不含信件內容（主旨／html）', auditText.indexOf('簽核通知') < 0 && auditText.indexOf('<p>') < 0);
  t('成功的信不寫稽核（沒有任何一行提到 mgr1／sec1）', !/mgr1|sec1/.test(auditText), auditText);
  // outbox 內容
  t('outbox 儲存體不含任何 email 全址', !FULL_EMAIL_RE.test(dump), (dump.match(FULL_EMAIL_RE) || [''])[0]);
  AMOUNT_NEEDLES.forEach((n) => t('outbox 儲存體不含金額字串「' + n + '」', dump.indexOf(n) < 0));
  NAME_NEEDLES.forEach((n) => t('outbox 儲存體不含「' + n + '」', dump.indexOf(n) < 0));
  t('outbox 儲存體不含主旨與內文', dump.indexOf('簽核通知') < 0 && dump.indexOf('<p>') < 0 && dump.indexOf('<html') < 0);
  t('outbox 儲存體不含密鑰樣式字串與 GUID', dump.indexOf('Bearer abcdef') < 0 && dump.indexOf('00000000-0000-4000') < 0);
  const rec = (await w.records()).find((r) => r.toUser === 'gm1');
  t('傳輸的錯誤訊息經過清理才存進 outbox（≤200 字、無全址）', rec.lastErrorMsg.length <= 200 && !/gm1@/.test(rec.lastErrorMsg) && /\[已遮蔽\]|\[email\]|\[id\]/.test(rec.lastErrorMsg), rec.lastErrorMsg);
  const rec2 = (await w.records()).find((r) => r.toUser === 'gm2');
  t('錯誤訊息裡回顯的主旨被遮蔽', rec2.lastErrorMsg.indexOf('簽核通知') < 0, rec2.lastErrorMsg);
  eq('outbox 紀錄欄位不含任何內容欄位', Object.keys(rec).filter((k) => /html|text|subject|body|amount|revenue|margin|email$/i.test(k)), []);
});

// ═════════════════════════════════════════════════════════════════════════
section('4 drainDue', async () => {
  // 基本：到期才領
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    w.behavior.fn = null;
    eq('drainDue：未到期不領', [(await w.drainDue({})).sent, w.sent.length], [0, 1]);
    w.clock.advance(60 * SEC);
    const rebuildArgs = [];
    const d = await w.drainDue({ rebuild: (job) => { rebuildArgs.push([job.status, job.toUser, job.type, job.attempts]); return { ev: w.ev }; } });
    eq('drainDue：到期後重寄，rebuild 收到 sending 的 job', [d.sent, rebuildArgs, w.sent.length], [1, [['sending', 'mgr1', 'E1_SUBMIT', 2]], 2]);
    eq('drainDue：重寄成功後紀錄 sent，attempts=2', [(await w.records())[0].status, (await w.records())[0].attempts], ['sent', 2]);
  }
  // rebuild 回 null → GONE；stepKey 不符 → STALE；rebuild 爆炸 → 稍後重試；沒給 rebuild → 不領取
  {
    const mk = async () => {
      const w = mkWorld();
      w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
      await w.dispatch(w.ev, rcp(['mgr1']), {});
      w.behavior.fn = null;
      w.clock.advance(60 * SEC);
      return w;
    };
    let w = await mk();
    let d = await w.drainDue({ rebuild: () => null });
    eq('rebuild 回 null：cancelled/GONE，不寄', [d.cancelled, w.sent.length, (await w.records())[0].status, (await w.records())[0].skipReason], [1, 1, 'cancelled', 'GONE']);
    w = await mk();
    d = await w.drainDue({ rebuild: () => ({ ev: mkEv({ stepKey: 'a-different-step' }) }) });
    eq('rebuild 的事件與工作的去重鍵不符（單據已進到別關）：cancelled/STALE', [d.cancelled, w.sent.length, (await w.records())[0].skipReason], [1, 1, 'STALE']);
    w = await mk();
    d = await w.drainDue({ rebuild: () => { throw new Error('quote lookup failed'); } });
    eq('rebuild 拋例外：稍後重試（REBUILD），不取消不失敗', [d.queued, d.cancelled, d.failed, (await w.records())[0].status, (await w.records())[0].lastErrorCode], [1, 0, [], 'pending', 'REBUILD']);
    w = await mk();
    d = await w.drainDue({ rebuild: () => ({ ev: { type: 'E1_SUBMIT' } }) });
    eq('rebuild 回不合法事件：failed/BAD_EVENT（permanent）', [d.failed, (await w.records())[0].status], [[{ username: 'mgr1', code: 'BAD_EVENT' }], 'failed']);
    w = await mk();
    d = await w.drainDue({ rebuild: () => 'garbage' });
    t('rebuild 回亂值：不 throw，job 不會卡在 sending', !d.hung && (await w.records())[0].status !== 'sending', short(d));
    w = await mk();
    const noRb = createDispatcher({ config: w.config, outbox: w.outbox, transport: w.transport, getUsers: () => w.users, writeLog: () => {}, render: () => ({ subject: 's', html: 'h' }), now: () => w.clock.t });
    d = await noRb.drainDue({});
    eq('沒有 rebuild：不領取任何工作（errors=1），工作仍是 pending', [d.errors, d.sent, (await w.records())[0].status], [1, 0, 'pending']);
    w = await mk();
    d = await w.drainDue({ rebuild: () => new Promise(() => { /* hang */ }) });
    eq('rebuild 永不 resolve：以輔助逾時返回並稍後重試', [d.queued, (await w.records())[0].lastErrorCode], [1, 'REBUILD']);
  }
  // 帳號之後被停用／移除 email
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w.dispatch(w.ev, rcp(['mgr1', 'gm1', 'gm2']), {});
    w.behavior.fn = null;
    w.users.find((u) => u.username === 'mgr1').active = false;
    delete w.users.find((u) => u.username === 'gm1').email;
    w.users.find((u) => u.username === 'gm2').email = 'gm2@example.test';
    w.clock.advance(60 * SEC);
    const d = await w.drainDue({ limit: 10 });
    eq('寄送當下重新檢查：停用／無 email／網域不符 → 改標 skipped，不寄', [d.sent, w.sent.length, d.skipped.map((x) => x.username + ':' + x.reason).sort()], [0, 3, ['gm1:NO_EMAIL', 'gm2:DOMAIN_NOT_ALLOWED', 'mgr1:INACTIVE']]);
    eq('寄送當下略過：紀錄 skipped', (await w.records()).map((r) => r.status), ['skipped', 'skipped', 'skipped']);
    eq('寄送當下略過：寫 QUOTE_MAIL_SKIPPED 稽核', w.logs.map((l) => l[0]).sort(), ['QUOTE_MAIL_SKIPPED', 'QUOTE_MAIL_SKIPPED', 'QUOTE_MAIL_SKIPPED']);
  }
  // 租約
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    w.behavior.fn = null;
    w.clock.advance(60 * SEC);
    const claimed = await w.outbox.claimDue({ limit: 5, leaseSec: 60 });          // 另一個工作者領走了（沒有處理完）
    eq('另一個工作者領走（sending）', claimed.length, 1);
    const d1 = await w.drainDue({});
    eq('租約未到期：drainDue 不會重複領取', [d1.sent, w.sent.length], [0, 1]);
    w.clock.advance(60 * SEC);
    const d2 = await w.drainDue({});
    eq('租約到期：drainDue 接手並寄出', [d2.sent, w.sent.length, (await w.records())[0].attempts], [1, 2, 3]);
  }
  // 並行 drainDue 不重複（建立大量 pending 需要關掉熔斷：連續 5 次逾時本來就會讓熔斷開啟）
  {
    const w = mkWorld({ config: { breaker: { failures: 1000, windowSec: 600, openSec: 600 } } });
    const users = [];
    const many = [];
    for (let i = 0; i < 10; i++) { users.push({ username: 'p' + i, role: 'user', email: 'p' + i + '@itts.com.tw' }); many.push('p' + i); }
    w.users.push(...users);
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w.dispatch(w.ev, rcp(many), {});
    w.behavior.fn = async () => { await sleep(5); return { ok: true }; };
    w.clock.advance(60 * SEC);
    const ds = await Promise.all([w.drainDue({ limit: 4 }), w.drainDue({ limit: 4 }), w.drainDue({ limit: 4 }), w.drainDue({ limit: 4 })]);
    const sentTo = w.sent.slice(10).map((m) => m.to[0]);
    eq('4 個並行 drainDue 搶 10 筆：重寄恰好 10 封、沒有重複', [sentTo.length, new Set(sentTo).size, ds.reduce((a, s) => a + s.sent, 0)], [10, 10, 10]);
  }
  // limit
  {
    const w = mkWorld({ config: { breaker: { failures: 1000, windowSec: 600, openSec: 600 } } });
    const users = [];
    const many = [];
    for (let i = 0; i < 7; i++) { users.push({ username: 'p' + i, role: 'user', email: 'p' + i + '@itts.com.tw' }); many.push('p' + i); }
    w.users.push(...users);
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w.dispatch(w.ev, rcp(many), {});
    w.behavior.fn = null;
    w.clock.advance(60 * SEC);
    eq('drainDue 預設 limit=5', (await w.drainDue({})).sent, 5);
    eq('drainDue limit=1', (await w.drainDue({ limit: 1 })).sent, 1);
    eq('drainDue 剩餘 1 筆', (await w.drainDue({ limit: 99 })).sent, 1);
  }
  // 租約過期、本輪用盡（工作者連續當掉）
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    for (let i = 0; i < 3; i++) { w.clock.advance(1000 * SEC); await w.outbox.claimDue({ limit: 5 }); }   // 當掉的工作者：領了就消失
    w.clock.advance(1000 * SEC);
    const d = await w.drainDue({});
    const rec = (await w.records())[0];
    eq('工作者連續當掉、本輪用盡：標 failed/LEASE_EXPIRED，不再寄', [d.sent, rec.status, rec.lastErrorCode, w.sent.length], [0, 'failed', 'LEASE_EXPIRED', 1]);
  }
  // drain 時 isStillValid／render／寄送 一樣會檢查
  {
    let valid = true;
    const w = mkWorld({ isStillValid: () => valid });
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    w.behavior.fn = null;
    valid = false;
    w.clock.advance(60 * SEC);
    const d = await w.drainDue({});
    eq('drainDue 寄送前也會呼叫 isStillValid：false → cancelled/STALE', [d.cancelled, w.sent.length, (await w.records())[0].skipReason], [1, 1, 'STALE']);
  }
  // kind 沿用工作紀錄上的 meta.kind
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w.dispatch(w.ev, [{ username: 'cons1', kind: 'consultant' }], {});
    w.behavior.fn = null;
    w.clock.advance(60 * SEC);
    w.renders.length = 0;
    await w.drainDue({});
    eq('drainDue：rebuild 沒給 kind 時沿用 job.meta.kind', w.renders.map((r) => r.viewer.kind), ['consultant']);
    const w2 = mkWorld();
    w2.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w2.dispatch(w2.ev, [{ username: 'cons1', kind: 'consultant' }], {});
    w2.behavior.fn = null;
    w2.clock.advance(60 * SEC);
    w2.renders.length = 0;
    await w2.drainDue({ rebuild: () => ({ ev: w2.ev, kind: 'owner' }) });
    eq('drainDue：rebuild 給的 kind 優先', w2.renders.map((r) => r.viewer.kind), ['owner']);
  }
  // drainDue 亂輸入
  {
    const w = mkWorld();
    for (const a of [undefined, null, 5, 'x', [], { limit: 'a' }, { limit: -5 }, { limit: Infinity }, { rebuild: 5 }]) {
      const r = await within(w.drainDue(a), 1000);
      t('drainDue 亂參數 ' + short(a) + '：不 throw、回傳 Summary', !r.error && !r.hung && sumKeys.every((k) => k in r.value), short(r));
    }
  }
});

// ═════════════════════════════════════════════════════════════════════════
// drainDue 的時間預算、中途熔斷、「歸還」語意，以及 LEASE_EXPIRED 可見性
// （審查發現：drainDue 一次領走 limit 筆、沒有時間預算、批次中途不重查熔斷；傳輸卡住時 Vercel maxDuration 30 秒先到，
//  沒輪到的工作停在 sending，沒寄過卻已被扣一次嘗試。現在改成「處理前才領一筆」，到預算或熔斷就停止領取，沒領的原封不動。）
// ═════════════════════════════════════════════════════════════════════════
section('4b drainDue：時間預算、熔斷、不白扣嘗試次數', async () => {
  // 直接入列 n 筆「已到期」的 pending（不經 dispatch，免得失敗的寄送先把熔斷弄開）
  async function mkDue(w, n) {
    const names = [];
    for (let i = 0; i < n; i++) {
      const u = 'p' + i;
      w.users.push({ username: u, role: 'user', email: u + '@itts.com.tw' });
      names.push(u);
      await w.outbox.enqueue({ type: w.ev.type, quoteId: w.ev.quoteId, quoteNo: w.ev.quoteNo, toUser: u, toMasked: '', dedupeKey: EVN.dedupeKey(w.ev, u), actorLabel: 'Actor Person', meta: { level: 1, kind: 'mgr1' } });
      w.clock.advance(1);                                        // createdAt 各不相同 → 領取順序確定（p0、p1、…）
    }
    return names;
  }
  const byUser = async (w) => { const m = {}; (await w.records()).forEach((r) => { m[r.toUser] = r; }); return m; };

  // 1) 預算：傳輸每封花 8 秒（假時鐘）；預算 20 秒、concurrency 1 → 在 0／8／16 秒各開始一封，第 4 封（24 秒）不再領取
  {
    const w = mkWorld({ concurrency: 1, config: { breaker: { failures: 1000, windowSec: 600, openSec: 600 } } });
    await mkDue(w, 10);
    w.behavior.fn = async () => { w.clock.advance(8000); return { ok: true, providerId: 'fake:slow' }; };
    const d = await w.drainDue({ limit: 10, budgetMs: 20000 });
    eq('預算 20 秒、每封 8 秒、10 筆到期：寄出 3 封就停止領取', [d.sent, w.sent.length], [3, 3]);
    eq('到預算停止：Summary 帶 budgetExhausted:true（不帶 breakerOpen）', [d.budgetExhausted, d.breakerOpen], [true, undefined]);
    const m = await byUser(w);
    const rest = Object.keys(m).filter((u) => m[u].status !== 'sent');
    eq('沒輪到的 7 筆：原封不動留在 pending（attempts=0、attemptsInRound=0、沒有租約）', rest.map((u) => [m[u].status, m[u].attempts, m[u].attemptsInRound, m[u].leaseUntil]), Array(7).fill(['pending', 0, 0, null]));
    eq('沒有任何工作卡在 sending', (await w.records()).filter((r) => r.status === 'sending').length, 0);
    // 下一次 drainDue 從上次停下的地方接手；全部寄完後每位收件人恰好一封
    const d2 = await w.drainDue({ limit: 10, budgetMs: 20000 });
    const d3 = await w.drainDue({ limit: 10, budgetMs: 20000 });
    const d4 = await w.drainDue({ limit: 10, budgetMs: 20000 });
    eq('分批接手：每批最多 3 封，最後把 10 筆都寄完', [d2.sent, d3.sent, d4.sent, w.sent.length], [3, 3, 1, 10]);
    eq('每位收件人恰好一封（沒有重寄、沒有漏寄）', w.sent.map((x) => x.to[0]).sort(), Array.from({ length: 10 }, (_, i) => 'p' + i + '@itts.com.tw').sort());
    eq('全部紀錄 sent，attempts 都是 1（沒有白扣）', (await w.records()).map((r) => [r.status, r.attempts]), Array(10).fill(['sent', 1]));
  }

  // 2) 預設預算 = min(20 秒, 2 × 單封總逾時)：totalMs 8000 → 16 秒；5000 → 10 秒；30000 → 20 秒（上限）
  for (const [totalMs, perSend, expectSent, expectBudget] of [[8000, 5000, 4, 16000], [5000, 5000, 2, 10000], [30000, 5000, 4, 20000]]) {
    const w = mkWorld({ concurrency: 1, timeouts: { connectMs: 20, totalMs }, config: { breaker: { failures: 1000, windowSec: 600, openSec: 600 } } });
    await mkDue(w, 10);
    w.behavior.fn = async () => { w.clock.advance(perSend); return { ok: true, providerId: 'fake:slow' }; };
    const d = await w.drainDue({ limit: 10 });
    eq('預設預算（totalMs=' + totalMs + ' → ' + expectBudget + 'ms）：每封 ' + perSend + 'ms，寄出 ' + expectSent + ' 封後停止', [d.sent, d.budgetExhausted === true], [expectSent, true]);
  }

  // 3) 熔斷在批次「中途」打開：本批自己的連續失敗（3 次 TIMEOUT 就開）→ 之後不再領取、不再寄
  {
    const w = mkWorld({ concurrency: 1, config: { breaker: { failures: 3, windowSec: 600, openSec: 600 } } });
    await mkDue(w, 10);
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT', message: 'timeout' });
    const d = await w.drainDue({ limit: 10 });
    eq('熔斷中途打開（3 次逾時）：只嘗試寄 3 封，不是 10 封', [w.sent.length, d.breakerOpen, d.budgetExhausted], [3, true, undefined]);
    const m = await byUser(w);
    eq('前 3 筆：嘗試過 1 次、排程重試（pending、attempts=1）', ['p0', 'p1', 'p2'].map((u) => [m[u].status, m[u].attempts, m[u].lastErrorCode]), Array(3).fill(['pending', 1, 'TIMEOUT']));
    eq('後 7 筆：沒有被領取、沒被扣次數（pending、attempts=0）', Array.from({ length: 7 }, (_, i) => 'p' + (i + 3)).map((u) => [m[u].status, m[u].attempts, m[u].lastErrorCode]), Array(7).fill(['pending', 0, null]));
    const d2 = await w.drainDue({ limit: 10 });
    eq('熔斷開著時再呼叫 drainDue：立刻返回 breakerOpen、不領取不寄送', [d2.breakerOpen, d2.sent, w.sent.length], [true, 0, 3]);
  }
  // 3b) 熔斷被「別的請求」在批次中途打開（本批自己沒有失敗）
  {
    const w = mkWorld({ concurrency: 1, config: { breaker: { failures: 1000, windowSec: 600, openSec: 600 } } });
    await mkDue(w, 6);
    // 成功的寄送會把熔斷重置，所以「別的請求把熔斷打開」要發生在我們這封成功回報之後：包住 markSent，第一封記完後開熔斷
    const origMarkSent = w.outbox.markSent;
    let opened = false;
    w.outbox.markSent = async (id) => { const r = await origMarkSent(id); if (!opened) { opened = true; await w.outbox.breaker.record(false, 'AUTH'); } return r; };
    const d = await w.drainDue({ limit: 6 });
    eq('別人在批次中途開了熔斷（AUTH）：寄完手上這封就停，其餘不領取', [d.sent, w.sent.length, d.breakerOpen], [1, 1, true]);
    eq('其餘 5 筆仍是 pending、attempts=0', (await w.records()).filter((r) => r.status === 'pending').map((r) => r.attempts), Array(5).fill(0));
  }
  // 3c) concurrency=3 時中途熔斷也會停（允許已經在飛的最多 3 封）
  {
    const w = mkWorld({ concurrency: 3, config: { breaker: { failures: 3, windowSec: 600, openSec: 600 } } });
    await mkDue(w, 12);
    w.behavior.fn = async () => { await sleep(3); return { ok: false, code: 'TIMEOUT', message: 'timeout' }; };
    const d = await w.drainDue({ limit: 12 });
    t('concurrency=3 熔斷中途打開：嘗試寄送數遠少於 12（<= 6）', w.sent.length >= 3 && w.sent.length <= 6 && d.breakerOpen === true, w.sent.length);
    eq('其餘工作沒被扣次數', (await w.records()).filter((r) => r.attempts === 0).length, 12 - w.sent.length);
  }

  // 4) 傳輸一直卡住（原審查情境）：20 筆到期、傳輸永不回應。預算 + 熔斷 讓一次 drainDue 在有限時間內結束，沒輪到的不被扣次數
  {
    const w = mkWorld({ concurrency: 3, timeouts: { connectMs: 20, totalMs: 40 } });          // 單封逾時 40ms（真實時間）
    await mkDue(w, 20);
    w.behavior.fn = () => new Promise(() => { /* 永不回應 */ });
    const r = await within(w.drainDue({ limit: 20 }), 5000);
    t('傳輸永不回應：drainDue 在有限時間內結束', !r.hung && !r.error, short(r));
    const recs = await w.records();
    const untouched = recs.filter((x) => x.attempts === 0);
    t('熔斷在 5 次逾時後打開：沒輪到的工作（attempts=0）至少 11 筆', untouched.length >= 11, untouched.length);
    eq('沒有任何工作卡在 sending', recs.filter((x) => x.status === 'sending').length, 0);
    eq('所有被嘗試過的都是 pending/TIMEOUT 排程重試，沒有 failed', recs.filter((x) => x.attempts > 0).map((x) => [x.status, x.attempts, x.lastErrorCode]), recs.filter((x) => x.attempts > 0).map(() => ['pending', 1, 'TIMEOUT']));
    eq('Summary 帶 breakerOpen', r.value.breakerOpen, true);
  }

  // 5) limit 語意不變：只處理 limit 筆（預算足夠時）；預設 5
  {
    const w = mkWorld({ concurrency: 3, config: { breaker: { failures: 1000, windowSec: 600, openSec: 600 } } });
    await mkDue(w, 9);
    eq('預算充足：limit=4 剛好處理 4 筆', [(await w.drainDue({ limit: 4 })).sent, w.sent.length], [4, 4]);
    eq('預設 limit=5', (await w.drainDue({})).sent, 5);
    const d = await w.drainDue({ limit: 4 });
    eq('剩 0 筆：sent=0、沒有 budgetExhausted／breakerOpen', [d.sent, d.budgetExhausted, d.breakerOpen], [0, undefined, undefined]);
  }

  // 6) budgetMs 亂值不 throw、退回預設
  {
    const w = mkWorld({ concurrency: 1, config: { breaker: { failures: 1000, windowSec: 600, openSec: 600 } } });
    await mkDue(w, 3);
    for (const b of ['x', null, -5, NaN, Infinity, {}, [], 1e12]) {
      const r = await within(w.drainDue({ limit: 1, budgetMs: b }), 2000);
      t('budgetMs=' + short(b) + '：不 throw、回傳 Summary', !r.hung && !r.error && sumKeys.every((k) => k in r.value), short(r));
    }
  }

  // 7) 領取時儲存體出錯（claimDue reject）：drainDue 不 throw、計入 errors，且回傳之後沒有任何工作者還在背景繼續寄
  {
    const w = mkWorld({ concurrency: 3, config: { breaker: { failures: 1000, windowSec: 600, openSec: 600 } } });
    await mkDue(w, 10);
    w.behavior.fn = async () => { await sleep(10); return { ok: true, providerId: 'fake:1' }; };
    const origClaimDue = w.outbox.claimDue;
    let calls = 0;
    w.outbox.claimDue = async (o) => { calls += 1; if (calls === 2) throw new Error('store down'); return origClaimDue(o); };
    const r = await within(w.drainDue({ limit: 10 }), 3000);
    t('claimDue 中途 reject：drainDue 不 throw、不卡住', !r.hung && !r.error, short(r));
    const nAtReturn = w.sent.length;
    const sendingAtReturn = (await w.records()).filter((x) => x.status === 'sending').length;      // 回傳當下還有工作者在飛的話，它的紀錄會是 sending
    await sleep(80);
    eq('claimDue reject 後：計入 errors，且 drainDue 回傳時所有已領取的工作都處理完了（沒有工作者還在背景繼續寄信）', [r.value.errors, sendingAtReturn, w.sent.length === nAtReturn, w.sent.length < 10], [1, 0, true, true]);
    eq('等一下之後也沒有任何工作卡在 sending', (await w.records()).filter((x) => x.status === 'sending').length, 0);
  }
});

section('4c drainDue：LEASE_EXPIRED 不再靜默', async () => {
  // 租約過期且本輪次數用盡 → failed/LEASE_EXPIRED：進 Summary.failed，寫 QUOTE_MAIL_FAILED（以前兩者都沒有）
  {
    const w = mkWorld();
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    for (let i = 0; i < 3; i++) { w.clock.advance(1000 * SEC); await w.outbox.claimDue({ limit: 5, sweep: false }); }   // 工作者連續當掉：領了就消失
    w.clock.advance(1000 * SEC);
    const logsBefore = w.logs.length;
    const d = await w.drainDue({});
    const rec = (await w.records())[0];
    eq('record：failed/LEASE_EXPIRED，沒有再寄', [rec.status, rec.lastErrorCode, w.sent.length, d.sent], ['failed', 'LEASE_EXPIRED', 1, 0]);
    eq('Summary.failed 帶這筆（LEASE_EXPIRED）', d.failed, [{ username: 'mgr1', code: 'LEASE_EXPIRED' }]);
    eq('稽核：寫了一筆 QUOTE_MAIL_FAILED（code=LEASE_EXPIRED），操作者 system', w.logs.slice(logsBefore).map((l) => [l[0], l[1], l[2], l[3]]), [['QUOTE_MAIL_FAILED', 'system', 'QU-000001-001', 'type=E1_SUBMIT to=mgr1 code=LEASE_EXPIRED']]);
    const logText = JSON.stringify(w.logs.slice(logsBefore));
    t('稽核不含專案名／客戶名／業務名／金額／完整 email', NAME_NEEDLES.concat(AMOUNT_NEEDLES).every((n) => logText.indexOf(n) < 0) && !FULL_EMAIL_RE.test(logText), logText);
    const d2 = await w.drainDue({});
    eq('已清掃過的不會重複回報、不重複稽核', [d2.failed, w.logs.length - logsBefore], [[], 1]);
  }
  // 與其他工作同批：failed 只列 LEASE_EXPIRED 那筆，其他照常寄出
  {
    const w = mkWorld({ config: { breaker: { failures: 1000, windowSec: 600, openSec: 600 } } });
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    for (let i = 0; i < 3; i++) { w.clock.advance(1000 * SEC); await w.outbox.claimDue({ limit: 5, sweep: false }); }
    w.clock.advance(1000 * SEC);
    w.behavior.fn = null;
    await w.dispatch(mkEv({ stepKey: '2026-10-08T05:00:00.000Z#1' }), rcp(['gm1']), {});          // 新事件，當次就寄出
    const d = await w.drainDue({});
    eq('同一批：LEASE_EXPIRED 列入 failed，不影響其他收件人', [d.failed.map((x) => x.username + ':' + x.code), (await w.records()).map((r) => r.toUser + ':' + r.status).sort()], [['mgr1:LEASE_EXPIRED'], ['gm1:sent', 'mgr1:failed']]);
  }
  // 稽核關掉（writeLog 丟例外）也不影響 drainDue
  {
    const w = mkWorld({ writeLog: () => { throw new Error('log down'); } });
    w.behavior.fn = async () => ({ ok: false, code: 'TIMEOUT' });
    await w.dispatch(w.ev, rcp(['mgr1']), {});
    for (let i = 0; i < 3; i++) { w.clock.advance(1000 * SEC); await w.outbox.claimDue({ limit: 5, sweep: false }); }
    w.clock.advance(1000 * SEC);
    const r = await within(w.drainDue({}), 2000);
    t('writeLog 丟例外：drainDue 仍回傳 Summary，且列入 failed', !r.hung && !r.error && r.value.failed.length === 1, short(r));
  }
});


// ═════════════════════════════════════════════════════════════════════════
section('5 永不 throw（隨機亂輸入）', async () => {
  let seed = 987654;
  const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
  const pick = (a) => a[Math.floor(rnd() * a.length)];
  const circ = (() => { const c = { type: 'E1_SUBMIT' }; c.self = c; return c; })();
  const boom = { get type() { throw new Error('getter boom'); }, get quoteId() { throw new Error('getter boom'); } };
  const proxyBoom = new Proxy({}, { get() { throw new Error('proxy boom'); }, has() { throw new Error('proxy boom'); }, ownKeys() { throw new Error('proxy boom'); } });
  const evs = [undefined, null, 5, 'x', [], {}, circ, boom, proxyBoom, mkEv(), mkEv({ stepKey: '' }), mkEv({ type: 'E2_COST_REQUEST', step: null }), mkEv({ numbers: null }), mkEv({ quoteId: 'a'.repeat(100) }), Object.defineProperty(mkEv(), 'projectName', { get() { throw new Error('boom'); }, enumerable: true })];
  const recs = [undefined, null, 5, 'x', {}, [], [null], [{}], [{ username: boom }], [{ username: 'mgr1' }], rcp(['mgr1']), rcp(['mgr1', 'ghost', 'noemail1']), [{ username: 'mgr1', kind: {} }], [{ username: 'mgr1', kind: 'nope' }], [proxyBoom], Array.from({ length: 300 }, (_, i) => ({ username: 'u' + i, kind: 'mgr1' }))];
  const optss = [undefined, null, 5, 'x', [], {}, { actorUsername: 5 }, { operatorLabel: {} }, { actorUsername: 'mgr1', operatorLabel: 'x'.repeat(5000) }, proxyBoom, { get actorUsername() { throw new Error('boom'); } }];
  const behaviors = [null, () => { throw new Error('boom'); }, () => Promise.reject(new Error('boom')), () => ({ ok: false, code: 'TIMEOUT' }), () => 5, () => ({ ok: true }), () => ({ ok: false, code: 'AUTH' })];
  let bad = 0;
  let total = 0;
  const firstBad = [];
  for (let i = 0; i < 300; i++) {
    const w = mkWorld({
      timeouts: { connectMs: 5, totalMs: 30 },
      isStillValid: pick([undefined, () => true, () => false, () => { throw new Error('boom'); }]),
      render: pick([undefined, () => { throw new Error('boom'); }, () => null, (ev, v) => ({ subject: 's', html: 'h', text: 't' })]),
      getUsers: pick([undefined, () => USERS(), () => { throw new Error('boom'); }, () => null, () => 'users', () => [null, 5, proxyBoom]]),
      writeLog: pick([undefined, () => { throw new Error('boom'); }, () => Promise.reject(new Error('boom')), () => {}]),
      env: { MAIL_MODE: pick(['live', 'live', 'redirect', 'log', 'off', 'garbage']), MAIL_REDIRECT_TO: 'redir@example.test' },
    });
    w.behavior.fn = pick(behaviors);
    const ev = pick(evs);
    const rc = pick(recs);
    const op = pick(optss);
    total += 1;
    const r = await within(w.dispatch(ev, rc, op), 1500);
    const okShape = r.value && sumKeys.every((k) => k in r.value) && Array.isArray(r.value.skipped) && Array.isArray(r.value.failed);
    if (r.hung || r.error || !okShape) { bad += 1; if (firstBad.length < 3) firstBad.push(short({ i, ev: ev && typeof ev === 'object' ? Object.keys(ev) : ev, r })); }
    const dr = await within(w.drainDue(pick([undefined, null, {}, { limit: 3 }, { rebuild: () => { throw new Error('boom'); } }, { rebuild: () => ({ ev: pick(evs) }) }, 5])), 1500);
    const okShape2 = dr.value && sumKeys.every((k) => k in dr.value);
    if (dr.hung || dr.error || !okShape2) { bad += 1; if (firstBad.length < 3) firstBad.push(short({ i, drain: dr })); }
  }
  t('隨機 ' + total + ' 組亂輸入／爆炸相依：dispatch 與 drainDue 全部正常返回 Summary（異常 ' + bad + '）', bad === 0, firstBad.join(' | '));
  // 工廠本身
  for (const d of [undefined, null, 5, 'x', [], {}, { config: null, outbox: null }, { outbox: {} }]) {
    let disp = null;
    let threw = false;
    try { disp = createDispatcher(d); } catch (e) { threw = true; }
    const r = disp ? await within(disp.dispatch(mkEv(), rcp(['mgr1']), {}), 1000) : null;
    t('createDispatcher(' + short(d).slice(0, 30) + ')：不 throw，dispatch 回 Summary（errors>0）', !threw && r && !r.error && r.value.errors > 0 && r.value.sent === 0, short(r));
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('6 整合', async () => {
  // JSON 檔案 outbox 全流程
  {
    const dir = tmpDir();
    const file = path.join(dir, 'mail-outbox.json');
    const w = mkWorld({ adapter: jsonFileAdapter({ file }) });
    w.behavior.fn = async (m, o, n) => (m.to[0].startsWith('gm1') ? { ok: false, code: 'TIMEOUT' } : { ok: true });
    const s = await w.dispatch(w.ev, rcp(['mgr1', 'gm1', 'noemail1']), { actorUsername: 'sales1' });
    eq('JSON 檔案 outbox：sent／queued／skipped', [s.sent, s.queued, s.skipped.map((x) => x.reason)], [1, 1, ['NO_EMAIL']]);
    const text = fs.readFileSync(file, 'utf8');
    t('JSON 檔案內容不含 email 全址、金額、專案名', !FULL_EMAIL_RE.test(text) && AMOUNT_NEEDLES.every((n) => text.indexOf(n) < 0) && NAME_NEEDLES.every((n) => text.indexOf(n) < 0));
    t('JSON 檔案目錄沒有殘留暫存檔', fs.readdirSync(dir).join() === 'mail-outbox.json');
    w.behavior.fn = null;
    w.clock.advance(60 * SEC);
    const d = await w.drainDue({});
    eq('JSON 檔案 outbox：重試成功', [d.sent, (await w.records()).map((r) => r.status).sort()], [1, ['sent', 'sent', 'skipped']]);
  }

  // 真的 render.js（若不存在則略過並標示）
  let real = null;
  try { real = load('lib/mail/render.js'); } catch (e) { real = null; }
  if (!real || typeof real.renderMail !== 'function') {
    t('【略過】lib/mail/render.js 尚未就緒，未跑真 render 整合測試（整合階段補跑）', true);
    return;
  }
  const mkReal = (over) => {
    const w = mkWorld(Object.assign({ render: real.renderMail }, over || {}));
    return w;
  };
  // E1 → 簽核人：有金額，主旨沒有
  {
    const w = mkReal();
    const s = await w.dispatch(w.ev, [{ username: 'mgr1', kind: 'mgr1' }, { username: 'gm1', kind: 'gm' }], { actorUsername: 'sales1' });
    eq('真 render：E1 兩位簽核人都寄出', [s.sent, s.failed, s.errors], [2, [], 0]);
    t('真 render：信件有 html／text，主旨以【簽核通知】開頭', w.sent.every((m) => m.html.length > 500 && m.text.length > 50 && /^【簽核通知】/.test(m.subject)), short(w.sent.map((m) => m.subject)));
    t('真 render：簽核人信含金額（12,345,678）', w.sent.every((m) => m.html.indexOf('12,345,678') >= 0 && m.text.indexOf('12,345,678') >= 0));
    t('真 render：主旨不含金額、毛利率、客戶名', w.sent.every((m) => !/12,345,678|41\.04|Acme/.test(m.subject)), short(w.sent.map((m) => m.subject)));
    t('真 render：outbox 與稽核仍不含金額／全址／名稱（即使信件本身有）', AMOUNT_NEEDLES.every((n) => w.dump().indexOf(n) < 0) && !FULL_EMAIL_RE.test(w.dump()) && NAME_NEEDLES.every((n) => w.dump().indexOf(n) < 0));
  }
  // E2 → 顧問：即使事件帶了 numbers，信件也不含金額
  {
    const w = mkReal();
    const ev = mkEv({ type: 'E2_COST_REQUEST', step: null, items: [{ desc: '現場安裝', qty: 2, unit: '式' }], stepKey: 'cost#1' });
    const s = await w.dispatch(ev, [{ username: 'cons1', kind: 'consultant' }], { actorUsername: 'sales1' });
    eq('真 render：E2 顧問信寄出', [s.sent, s.failed], [1, []]);
    const m = w.sent[0];
    const blob = m.subject + '\n' + m.html + '\n' + m.text;
    t('真 render：顧問信不含任何金額／毛利率數字', AMOUNT_NEEDLES.every((n) => blob.indexOf(n) < 0) && !/毛利|折扣/.test(blob), AMOUNT_NEEDLES.filter((n) => blob.indexOf(n) >= 0).join());
    t('真 render：顧問信連結帶 ?cost=1', /\/q\/q-1\?cost=1/.test(m.text) || /\/q\/q-1\?cost=1/.test(m.html));
  }
  // 秘書信：不帶客戶名與業務名
  {
    const w = mkReal();
    await w.dispatch(mkEv({ step: { level: 'board', label: '董事會' }, stepKey: 'board#1', type: 'E3_NEXT_STEP' }), [{ username: 'sec1', kind: 'secretary' }], { actorUsername: 'gm1' });
    const m = w.sent[0];
    const blob = m.subject + m.html + m.text;
    t('真 render：秘書（董事會關）信不含客戶名與業務名，但有金額', m && !/Acme Test Co|Sales One/.test(blob) && blob.indexOf('12,345,678') >= 0, m ? '' : '沒有寄出');
  }
  // 不允許的組合 → RENDER 失敗，不影響其他人
  {
    const w = mkReal();
    const s = await w.dispatch(w.ev, [{ username: 'cons1', kind: 'consultant' }, { username: 'mgr1', kind: 'mgr1' }], { actorUsername: 'sales1' });
    t('真 render：E1 寄給顧問被 render 拒絕（RENDER），簽核人仍正常寄出', s.sent === 1 && s.failed.length === 1 && s.failed[0].code === 'RENDER' && w.sent.length === 1, short(s));
  }
  // redirect 模式 + 真 render
  {
    const w = mkReal({ env: { MAIL_MODE: 'redirect', MAIL_REDIRECT_TO: 'redir@example.test' } });
    await w.dispatch(w.ev, [{ username: 'mgr1', kind: 'mgr1' }], { actorUsername: 'sales1' });
    t('真 render + redirect：只寄到測試信箱，橫幅在內文最上方', w.sent.length === 1 && w.sent[0].to.join() === 'redir@example.test' && /<body[^>]*><div [^>]*>【測試轉送】/.test(w.sent[0].html) && /^【測試轉送】/.test(w.sent[0].text), short(w.sent[0] && w.sent[0].html.slice(0, 300)));
    t('真 render + redirect：真實收件位址不在任何內容中', !/mgr1@itts/.test(JSON.stringify(w.sent)));
  }
  // 大小：每封信 < 100KB 且通過傳輸層驗證（BAD_MESSAGE 不會發生）
  {
    const w = mkReal();
    const s = await w.dispatch(w.ev, [{ username: 'mgr1', kind: 'mgr1' }], { actorUsername: 'sales1' });
    t('真 render：信件通過傳輸層的訊息驗證（TR.validateMessage）', TR.validateMessage(w.sent[0]).ok === true && s.failed.length === 0);
  }
});

main();
