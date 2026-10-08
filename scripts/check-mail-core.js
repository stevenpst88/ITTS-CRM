#!/usr/bin/env node
'use strict';
/**
 * 信件核心模組檢查。用法：node scripts/check-mail-core.js（不開伺服器、不連網、不寫檔、不碰 data.json／auth.json）
 *
 * 目前涵蓋（第一階段）：
 *   0) 原始碼衛生（lib/mail 內不可有不可見字元；只能 require 相對路徑或 Node 內建模組）
 *   1) safety     normalizeEmail／isAllowedDomain／maskEmail／headerSafe／escHtml／clip／safeText（含惡意輸入與 200 萬字元效能）
 *   2) config     getMailConfig／publicConfig（mode 護欄、缺設定不會誤寄、密鑰不外洩）
 *   3) recipients resolveRecipients（區分大小寫、各種略過原因、原型污染名稱）
 *   4) visibility 7 種收件人 × 6 欄位的逐格政策表
 *   5) events     validateEvent／dedupeKey
 * 後續階段追加（見檔尾「追加位置」註解）：userEmail、link／跳板頁、deep-link。
 *
 * 環境變數：
 *   MAIL_CORE_ROOT  要檢查的專案根目錄（預設＝本檔上一層）。變異測試（scripts/check-mail-mutation.js）會指向被破壞的副本。
 *
 * 撰寫備註：本檔不寫四位數的 \uXXXX（部分寫檔工具會把它轉成真正的字元），一律用 cp()／\u{...}／\x..。
 */
const fs = require('fs');
const path = require('path');
const util = require('util');
const assert = require('assert');
const Module = require('module');

const ROOT = process.env.MAIL_CORE_ROOT || path.join(__dirname, '..');
const load = (rel) => require(path.join(ROOT, rel));

// ── 迷你測試框架（章節可追加）────────────────────────────────────────────────
const results = [];
const perf = [];
let currentSection = '';

function record(name, ok, extra) {
  results.push({ section: currentSection, name, ok: !!ok, extra: extra === undefined ? '' : String(extra) });
}
const t = (name, ok, extra) => record(name, ok, extra);
function short(v) {
  let s;
  try { s = typeof v === 'string' ? JSON.stringify(v) : util.inspect(v, { depth: 4, breakLength: Infinity }); } catch (e) { s = String(v); }
  return s.length > 160 ? s.slice(0, 160) + '…' : s;
}
function eq(name, actual, expected) {
  let ok = true;
  try { assert.deepStrictEqual(actual, expected); } catch (e) { ok = false; }
  record(name, ok, ok ? '' : 'actual=' + short(actual) + ' expected=' + short(expected));
}
function noThrow(name, fn) {
  try { fn(); record(name, true); } catch (e) { record(name, false, 'threw: ' + (e && e.message)); }
}
function throwsCode(name, fn, code) {
  try { fn(); record(name, false, 'did not throw'); } catch (e) { record(name, e && e.code === code, 'code=' + (e && e.code) + ' msg=' + (e && e.message)); }
}
function minMs(fn, runs) {
  let best = Infinity;
  for (let i = 0; i < (runs || 3); i++) {
    const s = process.hrtime.bigint();
    fn();
    best = Math.min(best, Number(process.hrtime.bigint() - s) / 1e6);
  }
  return best;
}
/** 效能斷言：取最佳一次 < limitMs，並把數字記進報告 */
function fast(name, fn, limitMs) {
  const ms = minMs(fn);
  perf.push(name + ' ' + ms.toFixed(1) + 'ms');
  record('效能 ' + name + ' < ' + (limitMs || 50) + 'ms', ms < (limitMs || 50), ms.toFixed(1) + 'ms');
}
function section(name, fn) {
  currentSection = name;
  try { fn(); } catch (e) { record('章節執行中例外', false, (e && e.stack) || e); }
}
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

const cp = (...n) => String.fromCodePoint(...n);
const KELVIN = cp(0x212a);        // K（小寫化後會變成 ASCII 的 k）
const ELLIPSIS = cp(0x2026);

// ═════════════════════════════════════════════════════════════════════════
section('0 原始碼衛生', () => {
  const BAD = /[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\xa0\xad\u{61c}\u{180e}\u{2000}-\u{200f}\u{2028}-\u{202f}\u{205f}-\u{206f}\u{3000}\u{feff}\u{fffd}\u{e0000}-\u{e007f}]|\p{Cs}/u;
  const files = [];
  (function walk(d) {
    if (!fs.existsSync(d)) return;
    fs.readdirSync(d).forEach((n) => {
      const p = path.join(d, n);
      if (fs.statSync(p).isDirectory()) walk(p); else if (/\.js$/.test(n)) files.push(p);
    });
  })(path.join(ROOT, 'lib', 'mail'));
  const sd = path.join(ROOT, 'scripts');
  if (fs.existsSync(sd)) fs.readdirSync(sd).forEach((n) => { if (/^check-mail-.*\.js$|^mail-preview\.js$/.test(n)) files.push(path.join(sd, n)); });
  const dl = path.join(ROOT, '_client', 'deep-link.js');
  if (fs.existsSync(dl)) files.push(dl);
  t('找得到 lib/mail 的檔案', files.length >= 5, files.length);
  const builtins = new Set(Module.builtinModules);
  files.forEach((f) => {
    const src = fs.readFileSync(f, 'utf8');
    const rel = path.relative(ROOT, f).replace(/\\/g, '/');
    const badLine = src.split('\n').findIndex((l) => BAD.test(l));
    t(rel + ' 無不可見字元', badLine < 0, badLine < 0 ? '' : '第 ' + (badLine + 1) + ' 行');
    if (rel.startsWith('lib/mail/')) {
      const reqs = [];
      src.replace(/require\(\s*['"]([^'"]+)['"]\s*\)/g, (m, name) => { reqs.push(name); return m; });
      const outsiders = reqs.filter((n) => !(n.startsWith('./') || n.startsWith('../') || builtins.has(n.replace(/^node:/, ''))));
      t(rel + ' 只 require 相對路徑或 Node 內建模組（不新增 npm 相依）', outsiders.length === 0, outsiders.join(','));
    }
  });
});


// ═════════════════════════════════════════════════════════════════════════
// FIX-7：測試與文件的 email 位址衛生。公開 repo 裡的位址若是「人名樣式 local part ＋ 真實網域」（例如把 itts.com.tw 配上常見英文名），
// 可能碰巧對上真人信箱。規則：非保留網域（RFC 2606／6761：*.test、example.*、localhost、invalid）的位址，local part 只能是合成樣式
// （user1、test-user-1、a、x、sales1、mgr1a、gm2、first.last…），不可出現人名。需要驗證「預設網域白名單 itts.com.tw」行為的測試仍可使用該網域，
// 但 local part 一律用合成名稱。
section('0b 測試與文件的 email 位址衛生（無人名樣式位址）', () => {
  const RESERVED = /(^|\.)(example\.(test|com|org|net)|[a-z0-9-]+\.test|test|invalid|localhost|example)$/i;
  const GENERIC = new Set(['user', 'users', 'test', 'tester', 'u', 'a', 'b', 'c', 'x', 'y', 'z', 'k', 'g', 'm', 'n', 'p', 'q', 'r', 's', 't', 'tu', 'sales', 'mgr', 'gm', 'sec', 'cons', 'proxy', 'chair', 'prx', 'redir', 'redirect', 'box', 'qa',
    'shared', 'same', 'dup', 'solo', 'sender', 'send', 'ok', 'off', 'new', 'noemail', 'free', 'ghost', 'ws', 'upper', 'whatever', 'someone', 'dis', 'iv', 'lg', 'na', 'nb', 'nn', 'xss', 'evil', 'who', 'ext', 'pw', 'name', 'e', 'to', 'cc', 'bcc',
    'full', 'real', 'second', 'first', 'last', 'tag', 'person', 'address', 'sub', 'long', 'mail', 'sample', 'demo', 'foo', 'bar', 'role', 'admin', 'owner', 'consultant', 'secretary', 'board', 'exec', 'dummy', 'fake', 'nobody', 'anyone', 'mixed', 'case', 'plus']);
  // 合成角色名（套在「還帶著數字的原字串」上，所以數字洗不白縮寫）：
  //   單字母前綴 m／g／u／t／c／e／p／s 只能接數字（至少一位）：m1、g2、s10 ✓；ed、tj、mo、cy、m1x ✗（像真人縮寫）
  //   兩個字母以上的角色名只認列舉的詞 mgr／gm／sec／cons／proxy／chair／sales／prx／lead，後面可接「數字」或「數字＋一個字母」：mgr、mgr1、mgr1a ✓；mgrx、salesx、leader ✗
  const ROLE_RE = /^(?:(?:mgr|gm|sec|cons|proxy|chair|sales|prx|lead)(?:[0-9]+[a-z]?)?|[mgutceps][0-9]+)$/;
  const EMAIL_RE = /([A-Za-z0-9._%+-]+)@([A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+)/g;
  /** 回傳「人名樣式＋真實網域」的位址清單（原始碼裡的跳脫序列 \t \n \x7f 之類先拿掉，它們不是位址的一部分） */
  function scanText(text) {
    const bad = [];
    const clean = text.replace(/\\(?:[nrt]|x[0-9a-fA-F]{2}|u[0-9a-fA-F]{4}|u\{[0-9a-fA-F]+\})/g, ' ');
    let m;
    EMAIL_RE.lastIndex = 0;
    while ((m = EMAIL_RE.exec(clean))) {
      const local = m[1];
      const dom = m[2].toLowerCase();
      if (local.indexOf('*') >= 0 || RESERVED.test(dom) || /^[0-9]/.test(m[2])) continue;       // 遮罩位址／保留網域／套件版本（xlsx@0.18.5）
      // 純數字的段落略過；其餘每段：去掉數字後是 GENERIC 的詞，或整段（含數字）符合 ROLE_RE，才算合成名稱（數字不拿來洗白：ed1、e1d 的底字 ed 不在 GENERIC、也不符合 ROLE_RE）
      const parts = local.toLowerCase().split(/[._+%-]+/).filter((p) => p !== '' && !/^[0-9]+$/.test(p));
      if (!(parts.length === 0 || parts.every((p) => GENERIC.has(p.replace(/[0-9]+/g, '')) || ROLE_RE.test(p)))) bad.push(local + '@' + m[2]);
    }
    return bad;
  }
  // 掃描器本身（用植入的名字驗證它抓得到、也不會誤抓合成位址）
  // 植入的人名位址在執行時才組出來（原始碼裡不能出現這些字面值，否則下面的實際掃描會抓到本檔）
  [['ali' + 'ce', 'itts.com.tw'], ['Bo' + 'b', 'ITTS.com.tw'], ['ma' + 'ry', 'itts.com.tw'], ['da' + 've', 'gmail.com'], ['da' + 'vid', 'example2.org']].map((p) => p[0] + '@' + p[1]).forEach((x) => eq('掃描器抓得到人名樣式位址 ' + x, scanText('x ' + x + ' y'), [x]));
  ['user1@itts.com.tw', 'test-user-1@itts.com.tw', 'a@itts.com.tw', 'sales1@itts.com.tw', 'mgr1a@itts.com.tw', 'first.last@itts.com.tw', 'real.person@itts.com.tw', 'test-user-1@example.test', 'user2@itts.test', 'user3@example.com', 'u***@itts.com.tw', 'xlsx@0.18.5'].forEach((x) => eq('掃描器放行合成／保留網域／遮罩位址 ' + x, scanText('x ' + x + ' y'), []));
  eq('掃描器不把 \\t 之類的跳脫序列當成位址的一部分', scanText("'Bob\\tuser1@itts.com.tw'"), []);
  // 收緊（T3）：單字母前綴（m／g／u／t／c／e／p／s）後面只能接數字（至少一位），不能再接字母——ed、tj、mo、cy 這類像真人縮寫的 local part 搭配真實網域要被抓到；
  // 兩個字母以上的角色名只認明確列舉的詞（mgr／gm／sec／cons／proxy／chair／sales／prx／lead，後面可接數字或「數字＋一個字母」），其餘一律視為可疑。
  // 縮寫在執行時才組出來（原始碼裡不能出現「縮寫＋@＋真實網域」的字面值，否則下面的實際掃描會抓到本檔）
  const AT = String.fromCharCode(64);
  ['ed', 'tj', 'mo', 'cy', 'al', 'jo', 'ty', 'mc', 'sy', 'pj', 'ue', 'cj', 'gt', 'tm'].map((n) => n + AT + 'itts.com.tw').forEach((x) => eq('掃描器抓得到像真人縮寫的位址 ' + x, scanText('x ' + x + ' y'), [x]));
  ['ed1', 'tj2', 'mo3', 'cy10', 'e1d', 'm1x', 't2j', 'ed.lee', 'first.ed', 'ed_1', 'mgrx', 'salesx', 'managerx', 'leader'].map((n) => n + AT + 'itts.com.tw').forEach((x) => eq('掃描器：數字或分隔符號洗不白縮寫與未列舉的詞 ' + x, scanText('x ' + x + ' y'), [x]));
  ['user1', 'user', 'm1', 'g2', 'u3', 't4', 'c5', 'e6', 'p7', 's8', 'm12', 'gm1', 'gm', 'sec2', 'cons3', 'proxy1', 'chair1', 'prx2', 'sales1', 'mgr', 'mgr1', 'mgr1a', 'mgr2b', 'lead', 'lead1', 'test-user-1', 'user.1', 'a1', 'x2']
    .map((n) => n + AT + 'itts.com.tw').forEach((x) => eq('掃描器放行合成角色名 ' + x, scanText('x ' + x + ' y'), []));
  eq('掃描器：同一段文字多個位址只回報可疑的那幾個', scanText('ok ' + 'user1' + AT + 'itts.com.tw' + ' bad ' + 'ed' + AT + 'itts.com.tw' + ' ok ' + 'gm1' + AT + 'itts.com.tw' + ' bad ' + 'tj' + AT + 'itts.com.tw'), ['ed' + AT + 'itts.com.tw', 'tj' + AT + 'itts.com.tw']);
  eq('掃描器：像縮寫但在保留網域上 → 放行（保留網域不可能對應真人信箱）', ['ed', 'tj', 'mo'].map((n) => scanText('x ' + n + AT + 'itts.test y').length + scanText('x ' + n + AT + 'example.test y').length), [0, 0, 0]);

  // 實際掃描：lib/mail 全部檔案（含 README）、scripts/check-mail-*.js、scripts/mail-preview.js、_client/deep-link.js
  const files = [];
  (function walk(d) {
    if (!fs.existsSync(d)) return;
    fs.readdirSync(d).forEach((n) => { const p = path.join(d, n); if (fs.statSync(p).isDirectory()) walk(p); else if (/\.(js|md)$/.test(n)) files.push(p); });
  })(path.join(ROOT, 'lib', 'mail'));
  const sd = path.join(ROOT, 'scripts');
  if (fs.existsSync(sd)) fs.readdirSync(sd).forEach((n) => { if (/^check-mail-.*\.js$|^mail-preview\.js$/.test(n)) files.push(path.join(sd, n)); });
  const dl = path.join(ROOT, '_client', 'deep-link.js');
  if (fs.existsSync(dl)) files.push(dl);
  t('掃得到檔案（lib/mail、scripts/check-mail-*.js…）', files.length >= 10, files.length);
  let total = 0;
  files.forEach((f2) => {
    const bad = scanText(fs.readFileSync(f2, 'utf8'));
    total += bad.length;
    t(path.relative(ROOT, f2).replace(/\\/g, '/') + ' 沒有人名樣式＋真實網域的 email', bad.length === 0, bad.slice(0, 5).join(', '));
  });
  t('全部檔案合計 0 筆人名樣式位址', total === 0, total);
  // README 的聲明與實況一致
  const rd = fs.readFileSync(path.join(ROOT, 'lib', 'mail', 'README.md'), 'utf8');
  t('README 聲明「不含真實 email 位址或人名」且範例位址是合成值', /不含密鑰、租戶 ID、真實 email 位址或人名/.test(rd) && /user1@example\.test/.test(rd));
});

// ═════════════════════════════════════════════════════════════════════════
section('1 safety', () => {
  const S = load('lib/mail/safety.js');

  // ── normalizeEmail：合法 ──
  [
    ['a@itts.com.tw', 'a@itts.com.tw'],
    ['  A.B@ITTS.com.tw \n', 'a.b@itts.com.tw'],
    ['first.last+tag@itts.com.tw', 'first.last+tag@itts.com.tw'],
    ['user_1-x@sub.itts.com.tw', 'user_1-x@sub.itts.com.tw'],
    ['user1@example.test', 'user1@example.test'],
    ['a@xn--fiq228c.tw', 'a@xn--fiq228c.tw'],
    ['a@itts.com.tw' + cp(0xa0), 'a@itts.com.tw'],          // 頭尾的 NBSP 會被 trim 掉，結果仍是乾淨 ASCII
  ].forEach(([raw, want]) => {
    const r = S.normalizeEmail(raw);
    t('normalizeEmail 合法 ' + short(raw), r.ok === true && r.value === want, short(r));
  });

  // ── normalizeEmail：非法（含惡意輸入）──
  const bad = {
    '非字串 null': null, '非字串 undefined': undefined, '非字串 數字': 123, '非字串 物件': {}, '非字串 陣列': ['a@itts.com.tw'],
    '空字串': '', '全空白': '   ',
    '換行注入 CRLF+Bcc': 'a@itts.com.tw\r\nBcc: x@evil.test',
    '換行注入 LF': 'a@itts.com.tw\nb@itts.com.tw',
    '換行注入 前置': 'Bcc: x@evil.test\r\na@itts.com.tw',
    'Tab': 'a@itts.com.tw\tb',
    'NUL': 'a\x00@itts.com.tw',
    'DEL': 'a\x7f@itts.com.tw',
    '內含空白': 'a b@itts.com.tw',
    '內含空白(網域)': 'a@itts .com.tw',
    '逗號分隔多位址': 'a@itts.com.tw,b@itts.com.tw',
    '分號分隔多位址': 'a@itts.com.tw;b@itts.com.tw',
    '空白分隔多位址': 'a@itts.com.tw b@itts.com.tw',
    '角括號 名稱+位址': 'Name <a@itts.com.tw>',
    '角括號 單純': '<a@itts.com.tw>',
    '雙引號 local': '"a"@itts.com.tw',
    '單引號': "a'b@itts.com.tw",
    '括號 註解': 'a(comment)@itts.com.tw',
    '方括號 IP 網域': 'a@[127.0.0.1]',
    '反斜線': 'a\\b@itts.com.tw',
    '冒號': 'a:b@itts.com.tw',
    '百分號 percent-hack': 'a%evil.test@itts.com.tw',
    '驚嘆號': 'a!b@itts.com.tw',
    '兩個 @': 'a@b@itts.com.tw',
    '沒有 @': 'itts.com.tw',
    'local 為空': '@itts.com.tw',
    '網域為空': 'a@',
    '全形 a': cp(0xff41) + '@itts.com.tw',
    '全形 @': 'a' + cp(0xff20) + 'itts.com.tw',
    '全形句點': 'a@itts' + cp(0xff0e) + 'com' + cp(0xff0e) + 'tw',
    '中日文句點': 'a@itts' + cp(0x3002) + 'com.tw',
    '西里爾 а': cp(0x430) + '@itts.com.tw',
    'Kelvin 符號(小寫化後變 ASCII k)': KELVIN + 'user1@itts.com.tw',
    '土耳其 İ': 'user' + cp(0x130) + '@itts.com.tw',
    '中文': '業務@itts.com.tw',
    'U+2028': 'a@itts.com.tw' + cp(0x2028) + 'b',
    '內含 NBSP': 'a' + cp(0xa0) + 'b@itts.com.tw',
    '零寬字元': 'a' + cp(0x200b) + '@itts.com.tw',
    'RLO': cp(0x202e) + 'a@itts.com.tw',
    '網域沒有點': 'a@itts',
    '網域以點開頭': 'a@.itts.com.tw',
    '網域連續點': 'a@itts..com.tw',
    '網域尾端點': 'a@itts.com.tw.',
    '網域段以連字號開頭': 'a@-itts.com.tw',
    '網域段以連字號結尾': 'a@itts-.com.tw',
    '網域底線': 'a@itts_com.tw',
    '網域是 IP': 'a@1.2.3.4',
    'TLD 只有 1 字': 'a@itts.c',
    'TLD 全數字': 'a@itts.123',
    'local 以點開頭': '.a@itts.com.tw',
    'local 以點結尾': 'a.@itts.com.tw',
    'local 連續點': 'a..b@itts.com.tw',
    'local 65 字': 'l'.repeat(65) + '@itts.com.tw',
    '網域段 64 字': 'a@' + 'd'.repeat(64) + '.com',
    '總長 255': 'l'.repeat(64) + '@' + ['d'.repeat(63), 'd'.repeat(63), 'd'.repeat(58), 'com'].join('.'),
  };
  Object.keys(bad).forEach((k) => {
    const r = S.normalizeEmail(bad[k]);
    t('normalizeEmail 拒絕：' + k, r.ok === false && typeof r.error === 'string' && r.error.length > 0 && typeof r.code === 'string', short(r));
  });
  // 邊界：local 恰 64、網域段恰 63、總長恰 254 要通過
  t('normalizeEmail local 恰 64 字通過', S.normalizeEmail('l'.repeat(64) + '@itts.com.tw').ok === true);
  t('normalizeEmail 網域段恰 63 字通過', S.normalizeEmail('a@' + 'd'.repeat(63) + '.com').ok === true);
  const longOk = 'l'.repeat(64) + '@' + ['d'.repeat(63), 'd'.repeat(63), 'd'.repeat(57), 'com'].join('.');
  t('normalizeEmail 總長恰 254 通過', longOk.length === 254 && S.normalizeEmail(longOk).ok === true, longOk.length);
  t('normalizeEmail 錯誤訊息不回顯原值', !S.normalizeEmail('secret.person@evil .test').error.includes('secret.person'));
  const kel = S.normalizeEmail(KELVIN + 'user1@itts.com.tw');
  t('Kelvin 符號不會被小寫化洗成 ASCII 放行', kel.ok === false && kel.code === 'NON_ASCII', short(kel));

  // ── 效能 / ReDoS：200 萬字元 ──
  const MB2 = 2000000;
  [
    ['a×200萬', 'a'.repeat(MB2)],
    ['a.×100萬', 'a.'.repeat(MB2 / 2)],
    ['@×200萬', '@'.repeat(MB2)],
    ['空白×200萬+x', ' '.repeat(MB2) + 'x'],
    ['x+空白×200萬', 'x' + ' '.repeat(MB2)],
    ['a×200萬+@itts.com.tw', 'a'.repeat(MB2) + '@itts.com.tw'],
    ['a@+點×100萬', 'a@' + '.a'.repeat(MB2 / 2)],
    ['換行×200萬', '\n'.repeat(MB2)],
  ].forEach(([n, s]) => {
    fast('normalizeEmail ' + n, () => S.normalizeEmail(s));
    t('normalizeEmail 巨大輸入被拒絕：' + n, S.normalizeEmail(s).ok === false);
  });
  fast('isAllowedDomain 200萬字元', () => S.isAllowedDomain('a'.repeat(MB2) + '@itts.com.tw', ['itts.com.tw']));
  fast('maskEmail 200萬字元', () => S.maskEmail('a'.repeat(MB2)));

  // ── isAllowedDomain：完全相等 ──
  const AL = ['itts.com.tw'];
  [
    ['a@itts.com.tw', true], ['A@ITTS.COM.TW', true], ['  a@itts.com.tw\r\n', true],
    ['a@itts.com.tw.evil.test', false], ['a@evil-itts.com.tw', false], ['a@sub.itts.com.tw', false],
    ['a@xitts.com.tw', false], ['a@itts.com.twx', false], ['a@com.tw', false], ['a@itts.com', false],
    ['a@itts.com.tw\nb@evil.test', false], ['a@itts.com.tw,b@evil.test', false],
    [cp(0x430) + '@itts.com.tw', false], ['a' + cp(0xff20) + 'itts.com.tw', false],
    ['not an email', false], ['', false], [null, false], [undefined, false], [123, false],
  ].forEach(([e, want]) => eq('isAllowedDomain ' + short(e), S.isAllowedDomain(e, AL), want));
  eq('isAllowedDomain 白名單為空陣列＝全拒絕', S.isAllowedDomain('a@itts.com.tw', []), false);
  eq('isAllowedDomain 白名單 undefined', S.isAllowedDomain('a@itts.com.tw', undefined), false);
  eq('isAllowedDomain 白名單 null', S.isAllowedDomain('a@itts.com.tw', null), false);
  eq('isAllowedDomain 白名單是字串（非陣列）', S.isAllowedDomain('a@itts.com.tw', 'itts.com.tw'), false);
  eq('isAllowedDomain 萬用字元項目不生效', S.isAllowedDomain('a@sub.itts.com.tw', ['*.itts.com.tw']), false);
  eq('isAllowedDomain 點開頭項目不生效', S.isAllowedDomain('a@sub.itts.com.tw', ['.itts.com.tw']), false);
  eq('isAllowedDomain @ 開頭項目不生效', S.isAllowedDomain('a@itts.com.tw', ['@itts.com.tw']), false);
  eq('isAllowedDomain 白名單項目大小寫與空白被正規化', S.isAllowedDomain('a@itts.com.tw', [' ITTS.com.tw ']), true);
  eq('isAllowedDomain 白名單含非字串項目', S.isAllowedDomain('a@itts.com.tw', [null, 5, {}, ['itts.com.tw']]), false);
  eq('isAllowedDomain 多個白名單項目', S.isAllowedDomain('user1@example.test', ['itts.com.tw', 'example.test']), true);

  // ── maskEmail / maskedFromList ──
  eq('maskEmail 基本', S.maskEmail('user1@itts.com.tw'), 'u***@itts.com.tw');
  eq('maskEmail 大寫與空白先正規化', S.maskEmail(' User1@ITTS.com.tw '), 'u***@itts.com.tw');
  eq('maskEmail 單字 local', S.maskEmail('a@itts.com.tw'), 'a***@itts.com.tw');
  ['nope', '', null, undefined, 5, 'a@itts.com.tw\nBcc: x', 'a@@itts.com.tw'].forEach((x) => eq('maskEmail 非法 ' + short(x), S.maskEmail(x), '(invalid)'));
  t('maskEmail 不洩漏 local 第二字以後', !S.maskEmail('u123456@itts.com.tw').includes('23456'));
  eq('maskedFromList', S.maskedFromList(['a@itts.com.tw', 'bad', 'User2@itts.com.tw']), ['a***@itts.com.tw', '(invalid)', 'u***@itts.com.tw']);
  eq('maskedFromList 非陣列', S.maskedFromList('a@itts.com.tw'), []);
  eq('maskedFromList null', S.maskedFromList(null), []);

  // ── headerSafe ──
  const CTRL = /[\x00-\x1f\x7f-\x9f\u{2028}\u{2029}]/u;
  const INVIS = /[\u{61c}\u{200b}\u{200e}\u{200f}\u{202a}-\u{202e}\u{2060}\u{2066}-\u{2069}\u{feff}\u{e0000}-\u{e007f}]/u;
  eq('headerSafe 換行注入變單行', S.headerSafe('Subject\r\nBcc: x@evil.test'), 'Subject Bcc: x@evil.test');
  eq('headerSafe LF', S.headerSafe('a\nb'), 'a b');
  eq('headerSafe CR', S.headerSafe('a\rb'), 'a b');
  eq('headerSafe 連續空白合併並 trim', S.headerSafe('  a   b\t\tc  '), 'a b c');
  let allCtrlClean = true;
  for (let c = 0; c < 0xa0; c++) {
    if (c >= 0x20 && c < 0x7f) continue;
    const out = S.headerSafe('x' + String.fromCharCode(c) + 'y');
    if (CTRL.test(out) || out.indexOf('\r') >= 0 || out.indexOf('\n') >= 0) allCtrlClean = false;
  }
  t('headerSafe 剝除全部 C0／DEL／C1 控制字元', allCtrlClean);
  [0x85, 0x2028, 0x2029].forEach((c) => {
    const out = S.headerSafe('x' + String.fromCodePoint(c) + 'y');
    t('headerSafe 處理 U+' + c.toString(16).toUpperCase(), out === 'x y', short(out));
  });
  t('headerSafe NUL 不留在輸出', !S.headerSafe('a\x00b').includes('\x00'));
  [0x202e, 0x202d, 0x202a, 0x2066, 0x2069, 0x200b, 0x200e, 0x200f, 0x2060, 0xfeff, 0x61c, 0xe0041].forEach((c) => {
    const out = S.headerSafe('a' + String.fromCodePoint(c) + 'b');
    t('headerSafe 移除不可見／方向控制字元 U+' + c.toString(16).toUpperCase(), out === 'ab', short(out));
  });
  eq('headerSafe 預設截斷 255', Array.from(S.headerSafe('x'.repeat(400))).length, 255);
  eq('headerSafe 指定截斷', S.headerSafe('abcdef', 3), 'abc');
  eq('headerSafe 截斷後不留尾端空白', S.headerSafe('ab   cd', 3), 'ab');
  eq('headerSafe maxLen=0', S.headerSafe('abc', 0), '');
  eq('headerSafe maxLen 非法用預設', S.headerSafe('abc', NaN), 'abc');
  eq('headerSafe maxLen 小數取整', S.headerSafe('abcdef', 2.9), 'ab');
  const EMO = cp(0x1f600);
  eq('headerSafe 代理對：不切半（保留完整 emoji）', S.headerSafe('aaa' + EMO + EMO, 4), 'aaa' + EMO);
  eq('headerSafe 代理對：恰好能放下', S.headerSafe('ab' + EMO, 3), 'ab' + EMO);
  eq('headerSafe 代理對：放不下就整個不要', S.headerSafe('ab' + EMO, 2), 'ab');
  eq('headerSafe 全 emoji 以 code point 計', Array.from(S.headerSafe(EMO.repeat(50), 5)).length, 5);
  const loneHi = 'a' + String.fromCharCode(0xd83d) + 'b';
  const loneLo = 'a' + String.fromCharCode(0xde00) + 'b';
  eq('headerSafe 孤立高代理換成 U+FFFD', S.headerSafe(loneHi), 'a' + cp(0xfffd) + 'b');
  eq('headerSafe 孤立低代理換成 U+FFFD', S.headerSafe(loneLo), 'a' + cp(0xfffd) + 'b');
  eq('headerSafe null', S.headerSafe(null), '');
  eq('headerSafe undefined', S.headerSafe(undefined), '');
  eq('headerSafe 數字', S.headerSafe(123), '123');
  noThrow('headerSafe 物件的 toString 丟例外不外洩', () => { if (S.headerSafe({ toString() { throw new Error('x'); } }) !== '') throw new Error('should be empty'); });
  noThrow('headerSafe Symbol', () => S.headerSafe(Symbol('s')));
  [['a×200萬', 'a'.repeat(MB2)], ['換行×200萬', '\n'.repeat(MB2)], ['RLO×200萬', cp(0x202e).repeat(MB2)], ['emoji×100萬', EMO.repeat(MB2 / 2)], ['空白×200萬+x', ' '.repeat(MB2) + 'x']]
    .forEach(([n, s]) => { fast('headerSafe ' + n, () => S.headerSafe(s, 120)); fast('safeText ' + n, () => S.safeText(s, 200)); });
  t('headerSafe 巨大輸入結果長度受控', Array.from(S.headerSafe('a'.repeat(MB2), 120)).length === 120);

  // ── escHtml ──
  eq('escHtml 六個特殊字元', S.escHtml('&<>"\'`'), '&amp;&lt;&gt;&quot;&#39;&#96;');
  eq('escHtml 不重複跳脫順序', S.escHtml('&lt;'), '&amp;lt;');
  eq('escHtml script', S.escHtml('<script>alert(1)</script>'), '&lt;script&gt;alert(1)&lt;/script&gt;');
  eq('escHtml 屬性注入', S.escHtml('"><img src=x onerror=alert(1)>'), '&quot;&gt;&lt;img src=x onerror=alert(1)&gt;');
  t('escHtml 輸出不含 < > "', !/[<>"]/.test(S.escHtml('"><svg/onload=1>\'`')));
  eq('escHtml null', S.escHtml(null), '');
  eq('escHtml undefined', S.escHtml(undefined), '');
  eq('escHtml 數字', S.escHtml(42), '42');
  eq('escHtml 中文原樣', S.escHtml('報價單 QU-1'), '報價單 QU-1');
  eq('escHtml javascript: 不被改寫（href 本來就不該放使用者輸入）', S.escHtml('javascript:alert(1)'), 'javascript:alert(1)');
  noThrow('escHtml 物件 toString 丟例外', () => { if (S.escHtml({ toString() { throw new Error('x'); } }) !== '') throw new Error('should be empty'); });
  // escHtml 沒有正規式回溯問題，輸出本身就會膨脹到 8MB；這兩條實測約 35ms，門檻放寬到 150ms 只為避免機器忙碌時誤報，
  // 要守住的是「線性時間、不會卡死」。一般字串（無特殊字元）仍然要求 <50ms。
  fast('escHtml 200萬個 <（最壞，輸出 800 萬字元）', () => S.escHtml('<'.repeat(MB2)), 150);
  fast('escHtml 200萬字元混合特殊字元', () => S.escHtml('a<b&c"d'.repeat(MB2 / 7)), 150);
  fast('escHtml 200萬字元無特殊字元', () => S.escHtml('a'.repeat(MB2)));
  t('escHtml 巨大輸入輸出正確（200萬個 & → 每個 5 字元）', S.escHtml('&'.repeat(MB2)).length === MB2 * 5);

  // ── clip ──
  eq('clip 截斷加省略號', S.clip('abcdef', 3), 'abc' + ELLIPSIS);
  eq('clip 剛好等於上限不加', S.clip('abc', 3), 'abc');
  eq('clip 未超過', S.clip('ab', 3), 'ab');
  eq('clip n=0', S.clip('abc', 0), '');
  eq('clip 預設 200', Array.from(S.clip('x'.repeat(300))).length, 201);
  eq('clip n 非數字用預設', S.clip('abc', 'x'), 'abc');
  eq('clip 負數用預設', S.clip('abc', -5), 'abc');
  eq('clip 以 code point 計不切半', S.clip(EMO.repeat(3), 2), EMO + EMO + ELLIPSIS);
  eq('clip 中文', S.clip('一二三四五', 3), '一二三' + ELLIPSIS);
  eq('clip 省略號前不留空白', S.clip('ab  cd', 4), 'ab' + ELLIPSIS);
  eq('clip 孤立高代理不 throw 且不被當成一對', S.clip('a' + String.fromCharCode(0xd83d) + 'bcd', 3).length, 4);
  eq('clip null', S.clip(null, 5), '');
  fast('clip 200萬字元', () => S.clip('a'.repeat(MB2), 50));
  fast('clip emoji×100萬', () => S.clip(EMO.repeat(MB2 / 2), 50));

  // ── safeText ──
  eq('safeText 單行化', S.safeText('a\r\nb\tc'), 'a b c');
  eq('safeText 截斷加省略號', S.safeText('x'.repeat(30), 10), 'x'.repeat(10) + ELLIPSIS);
  eq('safeText 預設 200', Array.from(S.safeText('x'.repeat(500))).length, 201);
  eq('safeText 移除 RLO', S.safeText('a' + cp(0x202e) + 'b'), 'ab');
  eq('safeText null', S.safeText(null), '');
  const xssInputs = ['<script>alert(1)</script>', '"><img src=x onerror=alert(1)>', "'-alert(1)-'", 'javascript:alert(1)', 'a\r\nBcc: x@evil.test', cp(0x202e) + 'evil'];
  xssInputs.forEach((x) => {
    const out = S.escHtml(S.safeText(x, 200));
    t('escHtml(safeText(...)) 無 < > " 且單行：' + short(x), !/[<>"\r\n]/.test(out) && !INVIS.test(out), short(out));
  });
  t('safeText 過長輸入結果有界', Array.from(S.safeText('a'.repeat(MB2), 40)).length === 41);
});

// ═════════════════════════════════════════════════════════════════════════
section('2 config', () => {
  const C = load('lib/mail/config.js');
  const S = load('lib/mail/safety.js');
  const get = (env) => C.getMailConfig(env);
  const SECRET = 'SYNTH-SECRET-VALUE-0001';
  const GRAPH = { MAIL_GRAPH_TENANT_ID: 'tenant-synth-0001', MAIL_GRAPH_CLIENT_ID: 'client-synth-0001', MAIL_GRAPH_CLIENT_SECRET: SECRET, MAIL_GRAPH_SENDER: 'Sender@Itts.com.tw' };

  // 1) 空環境
  const d = get({});
  eq('空環境：mode 預設 off', d.mode, 'off');
  eq('空環境：redirectTo 空', d.redirectTo, '');
  eq('空環境：fromName 預設', d.fromName, 'ITTS-CRM ' + cp(0x7c3d, 0x6838, 0x901a, 0x77e5));
  eq('空環境：appBaseUrl 預設正式站', d.appBaseUrl, 'https://itts-crm.vercel.app');
  eq('空環境：allowedDomains 預設', d.allowedDomains, ['itts.com.tw']);
  eq('空環境：graph 未設定', d.graph.configured, false);
  eq('空環境：timeouts', d.timeouts, { connectMs: 3000, totalMs: 8000 });
  eq('空環境：retry', d.retry, { delaysSec: [60, 300, 900], maxAttempts: 3 });
  eq('空環境：breaker', d.breaker, { failures: 5, windowSec: 600, openSec: 600 });
  eq('空環境：previewDir 預設', d.previewDir, '.mail-preview');
  eq('空環境：outboxFile 預設', d.outboxFile, 'mail-outbox.json');
  eq('空環境：retentionDays', d.retentionDays, 90);
  eq('空環境：沒有 warnings', d.warnings, []);
  eq('null 環境等同空環境', get(null).mode, 'off');
  eq('非物件環境等同空環境', get('MAIL_MODE=live').mode, 'off');
  {
    const old = process.env.MAIL_MODE;
    process.env.MAIL_MODE = 'log';
    try { eq('省略參數時讀 process.env', C.getMailConfig().mode, 'log'); }
    finally { if (old === undefined) delete process.env.MAIL_MODE; else process.env.MAIL_MODE = old; }
  }

  // 2) mode 護欄
  [['off', 'off'], ['log', 'log'], ['live', 'live'], [' LIVE ', 'live'], ['Live', 'live'], ['\tlive\n', 'live'], ['LOG', 'log'], ['Off', 'off']]
    .forEach(([raw, want]) => eq('MAIL_MODE ' + short(raw) + ' → ' + want, get({ MAIL_MODE: raw }).mode, want));
  eq('MAIL_MODE=redirect 無目標 → 降為 log', get({ MAIL_MODE: 'redirect' }).mode, 'log');
  t('MAIL_MODE=redirect 無目標 → 有 warning', get({ MAIL_MODE: 'redirect' }).warnings.some((w) => w.includes('MAIL_REDIRECT_TO')));
  eq('MAIL_MODE=redirect 有目標 → redirect', get({ MAIL_MODE: 'redirect', MAIL_REDIRECT_TO: 'Tester@Example.test' }).mode, 'redirect');
  eq('MAIL_REDIRECT_TO 正規化為小寫', get({ MAIL_MODE: 'redirect', MAIL_REDIRECT_TO: ' Tester@Example.test ' }).redirectTo, 'tester@example.test');
  const notLive = [undefined, '', ' ', 'l1ve', 'liv', 'livee', 'live!', 'true', '1', 'yes', 'on', 'prod', 'production', 'enabled', 'LIVE=1', 'live;off', 'live off',
    'live\x00', 'live\u{a0}', cp(0xff4c, 0xff49, 0xff56, 0xff45), 'L' + cp(0x130) + 'VE', 'lIve' + KELVIN, '0', 'false', 'none', 'banana', '"live"', "'live'", 'live\\', 'live,log'];
  notLive.forEach((raw) => {
    const m = get({ MAIL_MODE: raw }).mode;
    t('MAIL_MODE ' + short(raw) + ' 絕不是 live（實際 ' + m + '）', m === 'off' || m === 'log', m);
  });
  [1, true, {}, [], ['live']].forEach((raw) => eq('MAIL_MODE 非字串 ' + short(raw) + ' → off', get({ MAIL_MODE: raw }).mode, 'off'));
  t('MAIL_MODE 非法值有 warning', get({ MAIL_MODE: 'banana' }).warnings.length === 1 && !get({ MAIL_MODE: 'banana' }).warnings[0].includes('banana'));
  eq('MAIL_MODE 未設或全空白不產生 warning', get({ MAIL_MODE: '  ' }).warnings, []);
  // 以獨立 oracle 做隨機比對：mode===live 若且唯若「去 ASCII 頭尾空白、轉小寫後恰為 live」
  {
    let seed = 20261008;
    const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
    const alpha = ['l', 'i', 'v', 'e', 'L', 'I', 'V', 'E', ' ', '\t', '\n', '\r', 'x', 'o', 'f', 'g', 'r', 'd', 'c', 't', cp(0x130), KELVIN, cp(0xff4c), cp(0xa0), '1', '-'];
    const oracle = (s) => {
      const m = s.replace(/^[ \t\r\n]+|[ \t\r\n]+$/g, '');
      if (!/^[A-Za-z]+$/.test(m)) return 'off';
      const l = m.toLowerCase();
      if (l === 'live' || l === 'log' || l === 'off') return l;
      return l === 'redirect' ? 'log' : 'off';
    };
    let mismatches = 0;
    let liveCount = 0;
    let first = '';
    const seen = [];
    for (let n = 0; n < 6000; n++) {
      let s = '';
      const len = 1 + Math.floor(rnd() * 7);
      for (let i = 0; i < len; i++) s += alpha[Math.floor(rnd() * alpha.length)];
      seen.push(s);
    }
    ['live', 'LIVE', ' Live\n', '\tLIVE\r\n', 'redirect', 'REDIRECT', 'log', 'off'].forEach((s) => seen.push(s));
    seen.forEach((s) => {
      const want = oracle(s);
      const got = get({ MAIL_MODE: s }).mode;
      if (want === 'live') liveCount++;
      if (got !== want) { mismatches++; if (!first) first = short(s) + ' got=' + got + ' want=' + want; }
    });
    t('MAIL_MODE 隨機比對獨立 oracle（' + seen.length + ' 組，其中 live ' + liveCount + ' 組）', mismatches === 0 && liveCount > 0, first);
  }
  // 窮舉環境組合：live 只在「明確設定 live」時出現；redirect 只在有合法目標時出現
  {
    const modes = [undefined, '', 'off', 'log', 'redirect', 'live', 'LIVE ', 'x', 'true'];
    const redirects = [undefined, '', 'nope', 'a@b@c', 'ok@example.test', 'a@itts.com.tw\nb'];
    const vercels = [undefined, '', '1'];
    let violations = 0;
    let combos = 0;
    modes.forEach((m) => redirects.forEach((r) => vercels.forEach((v) => {
      const env = {};
      if (m !== undefined) env.MAIL_MODE = m;
      if (r !== undefined) env.MAIL_REDIRECT_TO = r;
      if (v !== undefined) env.VERCEL = v;
      const c = get(env);
      combos++;
      const explicitLive = typeof m === 'string' && m.trim().toLowerCase() === 'live';
      if (c.mode === 'live' && !explicitLive) violations++;
      if (c.mode === 'redirect' && !c.redirectTo) violations++;
      if (explicitLive && c.mode !== 'live') violations++;
      if (!['off', 'log', 'redirect', 'live'].includes(c.mode)) violations++;
    })));
    t('環境組合窮舉（' + combos + ' 組）：不會因缺設定或亂設定而成為 live／redirect 無目標', violations === 0, violations);
  }
  eq('parseMode 無法辨識回 null', C.parseMode('banana'), null);
  eq('parseMode live', C.parseMode(' Live '), 'live');
  eq('normalizeMode 無法辨識回 off', C.normalizeMode('banana'), 'off');
  eq('normalizeMode undefined', C.normalizeMode(undefined), 'off');
  eq('normalizeMode redirect 保持 redirect（降級是 getMailConfig 的事）', C.normalizeMode('redirect'), 'redirect');

  // 3) redirect 目標
  ['a@x.test,b@y.test', 'nope', 'a@itts.com.tw\r\nBcc: x@evil.test', cp(0xff41) + '@example.test', 'a@b@c.test', '<a@example.test>', 'a@example.test;b@example.test'].forEach((raw) => {
    const c = get({ MAIL_MODE: 'redirect', MAIL_REDIRECT_TO: raw });
    t('MAIL_REDIRECT_TO 非法 ' + short(raw) + ' → log、不採用', c.mode === 'log' && c.redirectTo === '' && c.warnings.some((w) => w.includes('MAIL_REDIRECT_TO')), short(c.warnings));
  });
  eq('MAIL_REDIRECT_TO 不受網域白名單限制', get({ MAIL_MODE: 'redirect', MAIL_REDIRECT_TO: 'tester@example.test' }).mode, 'redirect');
  t('warnings 不回顯 MAIL_REDIRECT_TO 的值', !JSON.stringify(get({ MAIL_MODE: 'redirect', MAIL_REDIRECT_TO: 'secret.person@evil .test' }).warnings).includes('secret.person'));

  // 4) fromName
  eq('MAIL_FROM_NAME 自訂', get({ MAIL_FROM_NAME: ' 測試通知 ' }).fromName, '測試通知');
  t('MAIL_FROM_NAME 去換行', !/[\r\n]/.test(get({ MAIL_FROM_NAME: 'Evil\r\nBcc: x@evil.test' }).fromName));
  eq('MAIL_FROM_NAME 去 <>"\\', get({ MAIL_FROM_NAME: 'A<b>"c\\d' }).fromName, 'Abcd');
  eq('MAIL_FROM_NAME ≤60 字', Array.from(get({ MAIL_FROM_NAME: 'x'.repeat(200) }).fromName).length, 60);
  eq('MAIL_FROM_NAME 只有空白用預設', get({ MAIL_FROM_NAME: '   ' }).fromName, d.fromName);
  eq('MAIL_FROM_NAME 全是被剝除字元用預設', get({ MAIL_FROM_NAME: '<>"' }).fromName, d.fromName);

  // 5) appBaseUrl
  const base = (env) => get(env).appBaseUrl;
  eq('APP_BASE_URL 去尾端斜線', base({ APP_BASE_URL: 'https://crm.example.test/' }), 'https://crm.example.test');
  eq('APP_BASE_URL 去多個尾端斜線', base({ APP_BASE_URL: 'https://crm.example.test///' }), 'https://crm.example.test');
  eq('APP_BASE_URL 主機名稱轉小寫', base({ APP_BASE_URL: 'https://CRM.Example.test' }), 'https://crm.example.test');
  eq('VERCEL_PROJECT_PRODUCTION_URL 補 https://', base({ VERCEL_PROJECT_PRODUCTION_URL: 'crm-demo.example.test' }), 'https://crm-demo.example.test');
  eq('APP_BASE_URL 優先於 VERCEL_PROJECT_PRODUCTION_URL', base({ APP_BASE_URL: 'https://a.example.test', VERCEL_PROJECT_PRODUCTION_URL: 'b.example.test' }), 'https://a.example.test');
  eq('localhost 可用 http', base({ APP_BASE_URL: 'http://localhost:3000' }), 'http://localhost:3000');
  eq('127.0.0.1 可用 http', base({ APP_BASE_URL: 'http://127.0.0.1:3001/' }), 'http://127.0.0.1:3001');
  const DEFAULT_URL = 'https://itts-crm.vercel.app';
  ['http://crm.example.test', 'http://localhost.evil.test', 'ftp://crm.example.test', 'javascript:alert(1)', 'https://user:pw@crm.example.test', 'https://crm.example.test/path',
    'https://crm.example.test?x=1', 'https://crm.example.test#a', 'https://crm .example.test', 'not a url', 'https://crm.example.test"onmouseover=', '//crm.example.test', 'https://' + 'a'.repeat(300) + '.test']
    .forEach((raw) => {
      const c = get({ APP_BASE_URL: raw });
      t('APP_BASE_URL 非法 ' + short(raw).slice(0, 60) + ' → 預設＋warning', c.appBaseUrl === DEFAULT_URL && c.warnings.some((w) => w.includes('APP_BASE_URL')), c.appBaseUrl);
    });
  {
    const c = get({ VERCEL_PROJECT_PRODUCTION_URL: 'https://crm.example.test' });
    t('VERCEL_PROJECT_PRODUCTION_URL 已含 scheme（變成 https://https://）→ 預設＋warning', c.appBaseUrl === DEFAULT_URL && c.warnings.length === 1, short(c));
    const c2 = get({ VERCEL_PROJECT_PRODUCTION_URL: 'crm.example.test/evil' });
    t('VERCEL_PROJECT_PRODUCTION_URL 含路徑 → 預設＋warning', c2.appBaseUrl === DEFAULT_URL && c2.warnings.length === 1);
  }
  t('appBaseUrl 輸出不含引號、空白、換行', ['https://a.example.test', 'bad"url', 'https://a b'].every((u) => !/["'\s<>]/.test(base({ APP_BASE_URL: u }))));

  // 6) allowedDomains
  eq('MAIL_ALLOWED_DOMAINS 正規化＋去重', get({ MAIL_ALLOWED_DOMAINS: 'ITTS.com.tw, Example.TEST ,,itts.com.tw' }).allowedDomains, ['itts.com.tw', 'example.test']);
  {
    const c = get({ MAIL_ALLOWED_DOMAINS: 'itts.com.tw,*.evil.test,a@itts.com.tw,localhost,' + cp(0xff49) + 'tts.com.tw' });
    eq('MAIL_ALLOWED_DOMAINS 丟掉非法項目', c.allowedDomains, ['itts.com.tw']);
    t('MAIL_ALLOWED_DOMAINS 非法項目有 warning', c.warnings.some((w) => w.includes('MAIL_ALLOWED_DOMAINS')));
  }
  {
    const c = get({ MAIL_ALLOWED_DOMAINS: '*,localhost,a@b' });
    eq('MAIL_ALLOWED_DOMAINS 全非法 → 預設', c.allowedDomains, ['itts.com.tw']);
    t('MAIL_ALLOWED_DOMAINS 全非法有 warning', c.warnings.length >= 1);
  }
  eq('MAIL_ALLOWED_DOMAINS 空白 → 預設', get({ MAIL_ALLOWED_DOMAINS: '  ' }).allowedDomains, ['itts.com.tw']);

  // 7) Graph 與密鑰保護
  const full = get(Object.assign({ MAIL_MODE: 'live' }, GRAPH));
  eq('Graph 四項齊全 → configured', full.graph.configured, true);
  eq('Graph sender 正規化', full.graph.sender, 'sender@itts.com.tw');
  eq('Graph 可讀到 clientSecret（P4 傳輸層使用）', full.graph.clientSecret, SECRET);
  Object.keys(GRAPH).forEach((k) => {
    const env = Object.assign({}, GRAPH);
    delete env[k];
    const c = get(env);
    t('Graph 缺 ' + k + ' → configured=false 且有 warning', c.graph.configured === false && c.warnings.some((w) => w.includes('Graph')), short(c.warnings));
  });
  eq('Graph 完全沒設 → 沒有 warning', get({}).warnings, []);
  [['MAIL_GRAPH_TENANT_ID', 'a/b'], ['MAIL_GRAPH_TENANT_ID', 'a?b=1'], ['MAIL_GRAPH_CLIENT_ID', 'x y'], ['MAIL_GRAPH_CLIENT_SECRET', 'abc\ndef'], ['MAIL_GRAPH_SENDER', 'not-an-email'], ['MAIL_GRAPH_SENDER', 'a@b@c']]
    .forEach(([k, v]) => {
      const c = get(Object.assign({}, GRAPH, { [k]: v }));
      t('Graph ' + k + '=' + short(v) + ' 非法 → configured=false', c.graph.configured === false);
    });
  t('live 但 Graph 未設定 → warning 指出會 NOT_CONFIGURED', get({ MAIL_MODE: 'live' }).warnings.some((w) => w.includes('NOT_CONFIGURED')));
  {
    const c = full;
    const dump = [
      JSON.stringify(c),
      util.inspect(c, { depth: 10 }),
      util.inspect(c, { depth: 10, showHidden: true }),
      util.inspect(c.graph, { showHidden: true }),
      JSON.stringify(c.graph),
      JSON.stringify(Object.assign({}, c.graph)),
      JSON.stringify(Object.assign({}, c, { graph: Object.assign({}, c.graph) })),
      JSON.stringify(Object.keys(c.graph)),
      JSON.stringify(Object.entries(c.graph)),
      JSON.stringify(C.publicConfig(c)),
      JSON.stringify(c.warnings),
      String(c.graph) + `${c.graph}`,
      JSON.stringify({ x: c }),
    ].join('\n');
    t('任何序列化／inspect 都不含 clientSecret 值', !dump.includes(SECRET));
    t('序列化也不含租戶／用戶端 ID／寄件者', !dump.includes('tenant-synth-0001') && !dump.includes('client-synth-0001') && !dump.includes('sender@itts.com.tw'));
    eq('graph 的 JSON 只剩 configured', JSON.parse(JSON.stringify(c.graph)), { configured: true });
    t('clientSecret 不是 enumerable 屬性', !Object.keys(c.graph).includes('clientSecret') && Object.getOwnPropertyDescriptor(c.graph, 'clientSecret').enumerable === false);
    t('clientSecret 不可被覆寫', (() => { try { c.graph.clientSecret = 'x'; } catch (e) { /* strict mode */ } return c.graph.clientSecret === SECRET; })());
  }

  // 8) 每次呼叫回傳新物件；常數凍結
  {
    const a = get({});
    a.retry.delaysSec.push(1); a.allowedDomains.push('x.test'); a.timeouts.totalMs = 1; a.warnings.push('x'); a.breaker.failures = 99;
    const b = get({});
    eq('改動前一次的結果不污染下一次', [b.retry.delaysSec, b.allowedDomains, b.timeouts.totalMs, b.warnings, b.breaker.failures], [[60, 300, 900], ['itts.com.tw'], 8000, [], 5]);
    t('DEFAULTS 已凍結', Object.isFrozen(C.DEFAULTS) && Object.isFrozen(C.DEFAULTS.allowedDomains) && Object.isFrozen(C.MODES));
    eq('MODES', Array.from(C.MODES), ['off', 'log', 'redirect', 'live']);
  }

  // 9) 預覽目錄與 outbox
  eq('MAIL_PREVIEW_DIR 自訂', get({ MAIL_PREVIEW_DIR: 'tmp/prev' }).previewDir, 'tmp/prev');
  eq('VERCEL 有值 → previewDir 為 null', get({ VERCEL: '1' }).previewDir, null);
  eq('VERCEL 有值時即使設了 MAIL_PREVIEW_DIR 也是 null', get({ VERCEL: '1', MAIL_PREVIEW_DIR: 'x' }).previewDir, null);
  eq('VERCEL 空字串視為沒有', get({ VERCEL: '' }).previewDir, '.mail-preview');
  ['../evil', 'a/../b', 'a\\..\\b', '..', 'x\x00y', 'x\ny', 'a'.repeat(201)].forEach((raw) => {
    const c = get({ MAIL_PREVIEW_DIR: raw });
    t('MAIL_PREVIEW_DIR 非法 ' + short(raw).slice(0, 40) + ' → 預設＋warning', c.previewDir === '.mail-preview' && c.warnings.some((w) => w.includes('MAIL_PREVIEW_DIR')), short(c.previewDir));
  });
  eq('MAIL_OUTBOX_FILE 自訂', get({ MAIL_OUTBOX_FILE: 'data/out.json' }).outboxFile, 'data/out.json');
  ['../x.json', 'a/../../x', 'x\ny.json'].forEach((raw) => {
    const c = get({ MAIL_OUTBOX_FILE: raw });
    t('MAIL_OUTBOX_FILE 非法 ' + short(raw) + ' → 預設＋warning', c.outboxFile === 'mail-outbox.json' && c.warnings.some((w) => w.includes('MAIL_OUTBOX_FILE')));
  });

  // 10) publicConfig
  {
    const env = Object.assign({ MAIL_MODE: 'redirect', MAIL_REDIRECT_TO: 'tester@example.test' }, GRAPH);
    const c = get(env);
    const pc = C.publicConfig(c);
    eq('publicConfig 欄位', Object.keys(pc).sort(), ['allowedDomains', 'appBaseUrl', 'fromName', 'graphConfigured', 'mode', 'redirectConfigured', 'retry', 'timeouts', 'warnings']);
    eq('publicConfig 值', [pc.mode, pc.redirectConfigured, pc.graphConfigured], ['redirect', true, true]);
    const text = JSON.stringify(pc);
    t('publicConfig 不含 redirectTo 原值', !text.includes('tester@example.test') && !text.includes('example.test'));
    t('publicConfig 不含密鑰與租戶資訊', !text.includes(SECRET) && !text.includes('tenant-synth') && !text.includes('client-synth') && !text.includes('sender@'));
    eq('publicConfig 沒設 redirect/graph', [C.publicConfig(get({})).redirectConfigured, C.publicConfig(get({})).graphConfigured], [false, false]);
    pc.allowedDomains.push('x.test'); pc.timeouts.totalMs = 1; pc.retry.delaysSec.push(1);
    eq('改 publicConfig 結果不影響原 config', [c.allowedDomains, c.timeouts.totalMs, c.retry.delaysSec], [['itts.com.tw'], 8000, [60, 300, 900]]);
  }
  [null, undefined, {}, 'x', 5, [], { mode: 'banana', graph: null, retry: null, timeouts: null }].forEach((x) => noThrow('publicConfig 容忍殘缺輸入 ' + short(x), () => C.publicConfig(x)));
  eq('publicConfig 的 mode 重新正規化', [C.publicConfig({ mode: 'banana' }).mode, C.publicConfig({ mode: 'LIVE ' }).mode, C.publicConfig(null).mode], ['off', 'live', 'off']);
});

// ═════════════════════════════════════════════════════════════════════════
section('3 recipients', () => {
  const R = load('lib/mail/recipients.js');
  const C = load('lib/mail/config.js');
  const cfg = C.getMailConfig({});
  const users = [
    { username: 'u1', email: 'u1@itts.com.tw', displayName: 'User One', nickname: 'Uno', role: 'user' },
    { username: 'U1', email: 'Upper@ITTS.com.tw', displayName: 'Upper One', role: 'user' },
    { username: 'mgr', email: 'Mgr@ITTS.com.tw', displayName: 'Manager', role: 'manager1' },
    { username: 'noemail', displayName: 'No Email', role: 'user' },
    { username: 'nullmail', email: null, role: 'user' },
    { username: 'blank', email: '   ', role: 'user' },
    { username: 'bad', email: 'not-an-email', role: 'user' },
    { username: 'multi', email: 'a@itts.com.tw,b@itts.com.tw' },
    { username: 'inj', email: 'a@itts.com.tw\r\nBcc: x@evil.test' },
    { username: 'num', email: 12345 },
    { username: 'ext', email: 'x@example.test' },
    { username: 'evil1', email: 'x@itts.com.tw.evil.test' },
    { username: 'evil2', email: 'x@evil-itts.com.tw' },
    { username: 'sub', email: 'x@sub.itts.com.tw' },
    { username: 'off', email: 'off@itts.com.tw', active: false },
    { username: 'dis', email: 'dis@itts.com.tw', disabled: true },
    { username: 'poolacct', email: 'p@itts.com.tw', role: 'pool' },
    { username: 'xss', email: 'xss@itts.com.tw', displayName: '<script>alert(1)</script>\r\nX' + cp(0x202e) },
    { username: 'wsnick', email: 'ws@itts.com.tw', nickname: '   ', displayName: 'Disp', role: 'user' },
    { username: 'noname', email: 'nn@itts.com.tw' },
    { username: 'invis', email: 'iv@itts.com.tw', nickname: cp(0x202e, 0x200b), displayName: cp(0x200b) },
    { username: 'constructor', email: 'c@itts.com.tw' },
    { username: '__proto__', email: 'p2@itts.com.tw' },
    { username: 'long', email: 'lg@itts.com.tw', displayName: 'x'.repeat(300) },
  ];
  const res = (names, over) => R.resolveRecipients(names, Object.assign({ users, actorUsername: 'someone-else', config: cfg }, over));
  const reasonsOf = (r) => r.skipped.map((s) => s.username + ':' + s.reason);

  // 基本與順序
  {
    const r = res(['mgr', 'u1']);
    eq('deliver 保留輸入順序、email 正規化為小寫', r.deliver.map((x) => [x.username, x.email]), [['mgr', 'mgr@itts.com.tw'], ['u1', 'u1@itts.com.tw']]);
    eq('label：暱稱優先，其次顯示名稱', r.deliver.map((x) => x.label), ['Manager', 'Uno']);
    eq('沒有 skipped', r.skipped, []);
    eq('deliver 項目的欄位', Object.keys(r.deliver[0]).sort(), ['email', 'label', 'username']);
  }
  // 區分大小寫
  {
    eq('只差大小寫的兩個帳號是不同人，各自送到各自信箱', res(['u1', 'U1']).deliver.map((x) => [x.username, x.email]), [['u1', 'u1@itts.com.tw'], ['U1', 'upper@itts.com.tw']]);
    eq('同名重複記 DUP', reasonsOf(res(['u1', 'u1'])), ['u1:DUP']);
    eq('重複只送一封', res(['u1', 'u1', 'u1']).deliver.length, 1);
    eq('操作者 u1 不會排除 U1', res(['u1', 'U1'], { actorUsername: 'u1' }).deliver.map((x) => x.username), ['U1']);
    eq('操作者 U1 不會排除 u1', res(['u1', 'U1'], { actorUsername: 'U1' }).deliver.map((x) => x.username), ['u1']);
    eq('大小寫不同的帳號名 MGR → UNKNOWN_USER', reasonsOf(res(['MGR'])), ['MGR:UNKNOWN_USER']);
    eq('大小寫不同的帳號名 Mgr → UNKNOWN_USER', reasonsOf(res(['Mgr'])), ['Mgr:UNKNOWN_USER']);
  }
  // 各種略過原因
  eq('ACTOR', reasonsOf(res(['u1', 'mgr'], { actorUsername: 'u1' })), ['u1:ACTOR']);
  eq('ACTOR 不在 users 裡也算 ACTOR', reasonsOf(res(['ghost'], { actorUsername: 'ghost' })), ['ghost:ACTOR']);
  eq('UNKNOWN_USER', reasonsOf(res(['ghost'])), ['ghost:UNKNOWN_USER']);
  eq('INACTIVE active:false', reasonsOf(res(['off'])), ['off:INACTIVE']);
  eq('INACTIVE disabled:true', reasonsOf(res(['dis'])), ['dis:INACTIVE']);
  eq('INACTIVE role pool', reasonsOf(res(['poolacct'])), ['poolacct:INACTIVE']);
  eq('NO_EMAIL 未設', reasonsOf(res(['noemail'])), ['noemail:NO_EMAIL']);
  eq('NO_EMAIL null', reasonsOf(res(['nullmail'])), ['nullmail:NO_EMAIL']);
  eq('NO_EMAIL 全空白', reasonsOf(res(['blank'])), ['blank:NO_EMAIL']);
  eq('BAD_EMAIL 格式錯', reasonsOf(res(['bad'])), ['bad:BAD_EMAIL']);
  eq('BAD_EMAIL 多位址', reasonsOf(res(['multi'])), ['multi:BAD_EMAIL']);
  eq('BAD_EMAIL 換行注入', reasonsOf(res(['inj'])), ['inj:BAD_EMAIL']);
  eq('BAD_EMAIL 非字串', reasonsOf(res(['num'])), ['num:BAD_EMAIL']);
  eq('DOMAIN_NOT_ALLOWED 外部網域', reasonsOf(res(['ext'])), ['ext:DOMAIN_NOT_ALLOWED']);
  eq('DOMAIN_NOT_ALLOWED itts.com.tw.evil.test', reasonsOf(res(['evil1'])), ['evil1:DOMAIN_NOT_ALLOWED']);
  eq('DOMAIN_NOT_ALLOWED evil-itts.com.tw', reasonsOf(res(['evil2'])), ['evil2:DOMAIN_NOT_ALLOWED']);
  eq('DOMAIN_NOT_ALLOWED 子網域', reasonsOf(res(['sub'])), ['sub:DOMAIN_NOT_ALLOWED']);
  eq('白名單加入 example.test 後 ext 可送', res(['ext'], { config: Object.assign({}, cfg, { allowedDomains: ['itts.com.tw', 'example.test'] }) }).deliver.map((x) => x.email), ['x@example.test']);
  eq('白名單為空 → 全部 DOMAIN_NOT_ALLOWED', reasonsOf(res(['u1', 'mgr'], { config: Object.assign({}, cfg, { allowedDomains: [] }) })), ['u1:DOMAIN_NOT_ALLOWED', 'mgr:DOMAIN_NOT_ALLOWED']);
  eq('沒給 config 時用預設白名單', [res(['u1'], { config: undefined }).deliver.length, res(['ext'], { config: undefined }).skipped.length], [1, 1]);
  eq('skipped 順序依輸入', reasonsOf(res(['ghost', 'u1', 'bad', 'off', 'u1'])), ['ghost:UNKNOWN_USER', 'bad:BAD_EMAIL', 'off:INACTIVE', 'u1:DUP']);
  eq('第一次被略過的帳號，第二次仍記 DUP', reasonsOf(res(['bad', 'bad'])), ['bad:BAD_EMAIL', 'bad:DUP']);
  eq('收件人層永遠回真實位址（即使 config 是 redirect）', res(['u1'], { config: Object.assign({}, cfg, { mode: 'redirect', redirectTo: 'tester@example.test' }) }).deliver[0].email, 'u1@itts.com.tw');
  // label
  {
    const r = res(['xss', 'wsnick', 'noname', 'invis', 'long']);
    const lab = r.deliver.map((x) => x.label);
    t('label：顯示名稱含 script／換行／RLO → 單行且無不可見字元', !/[\r\n\u{202e}]/u.test(lab[0]) && lab[0].includes('script'), short(lab[0]));
    t('label 再經 escHtml 後無 < >', !/[<>]/.test(load('lib/mail/safety.js').escHtml(lab[0])));
    eq('label：暱稱全空白 → 顯示名稱', lab[1], 'Disp');
    eq('label：都沒有 → 帳號', lab[2], 'noname');
    eq('label：暱稱與顯示名稱都只剩不可見字元 → 帳號', lab[3], 'invis');
    t('label 過長被截斷（≤61 code point）', Array.from(lab[4]).length <= 61, Array.from(lab[4]).length);
  }
  // 原型污染名稱 / users 形態
  {
    eq('users 陣列含 constructor／__proto__ 帳號時可正常解析', res(['constructor', '__proto__']).deliver.map((x) => x.username), ['constructor', '__proto__']);
    const r = R.resolveRecipients(['__proto__', 'constructor', 'toString', 'hasOwnProperty', 'valueOf'], { users: [{ username: 'u1', email: 'u1@itts.com.tw' }], config: cfg });
    eq('不存在的原型名稱 → UNKNOWN_USER', [r.deliver.length, r.skipped.map((s) => s.reason)], [0, ['UNKNOWN_USER', 'UNKNOWN_USER', 'UNKNOWN_USER', 'UNKNOWN_USER', 'UNKNOWN_USER']]);
    const mapUsers = { u1: { username: 'u1', email: 'u1@itts.com.tw', displayName: 'U One' }, keyonly: { email: 'k@itts.com.tw' } };
    const r2 = R.resolveRecipients(['u1', 'keyonly', 'constructor'], { users: mapUsers, config: cfg });
    eq('users 也可以是以 username 為鍵的物件（quoteRoutes 的 userMap）', [r2.deliver.map((x) => x.username), r2.skipped.map((s) => s.reason)], [['u1', 'keyonly'], ['UNKNOWN_USER']]);
  }
  // 垃圾輸入
  [null, undefined, 'u1', {}, 5, true].forEach((x) => eq('usernames 非陣列 ' + short(x) + ' → 空結果', R.resolveRecipients(x, { users, config: cfg }), { deliver: [], skipped: [] }));
  noThrow('opts 為 undefined 不 throw', () => R.resolveRecipients(['u1']));
  noThrow('opts 為 null 不 throw', () => R.resolveRecipients(['u1'], null));
  eq('users 為 null → 全部 UNKNOWN_USER', reasonsOf(R.resolveRecipients(['u1'], { users: null, config: cfg })), ['u1:UNKNOWN_USER']);
  eq('清單內的非字串／空字串項目 → UNKNOWN_USER', R.resolveRecipients([null, 5, '', {}, undefined], { users, config: cfg }).skipped.map((s) => s.reason), ['UNKNOWN_USER', 'UNKNOWN_USER', 'UNKNOWN_USER', 'UNKNOWN_USER', 'UNKNOWN_USER']);
  eq('users 陣列內含壞項目不 throw', R.resolveRecipients(['u1'], { users: [null, 5, 'x', { username: 5 }, { username: 'u1', email: 'u1@itts.com.tw' }], config: cfg }).deliver.length, 1);
  eq('skipped 的 username 已清理（換行／控制字元）', R.resolveRecipients(['evil\r\nBcc: x'], { users, config: cfg }).skipped[0].username, 'evil Bcc: x');
  // 不就地修改輸入
  {
    const frozen = JSON.parse(JSON.stringify(users));
    frozen.forEach((u) => Object.freeze(u));
    Object.freeze(frozen);
    const before = JSON.stringify(frozen);
    noThrow('輸入物件被 freeze 也能運作', () => R.resolveRecipients(['u1', 'mgr', 'bad', 'off'], { users: frozen, config: Object.freeze(Object.assign({}, cfg)) }));
    eq('不就地修改 users', JSON.stringify(frozen), before);
  }
  // 大量輸入
  {
    const many = new Array(5000).fill('u1');
    let r;
    fast('resolveRecipients 5000 筆重複輸入', () => { r = res(many); }, 100);
    eq('大量重複輸入仍只送一封', r.deliver.length, 1);
    t('超長清單被截在上限內（不會無限制處理）', r.skipped.length <= 1000, r.skipped.length);
  }
  eq('SKIP_REASONS 完整', Object.keys(R.SKIP_REASONS).sort(), ['ACTOR', 'BAD_EMAIL', 'DOMAIN_NOT_ALLOWED', 'DUP', 'INACTIVE', 'NO_EMAIL', 'UNKNOWN_USER']);
});

// ═════════════════════════════════════════════════════════════════════════
section('4 visibility', () => {
  const V = load('lib/mail/visibility.js');
  // 這張矩陣是獨立抄錄的期望值（來源：規格 §1.4 與業主 2026-10-08 決定），不引用被測程式
  const FIELDS = ['amount', 'margin', 'tier', 'customer', 'owner', 'items', 'project'];
  const EXPECT = {
    mgr1:       [1, 1, 1, 1, 1, 0, 1],
    gm:         [1, 1, 1, 1, 1, 0, 1],
    chairman:   [1, 1, 1, 1, 1, 0, 1],
    secretary:  [1, 1, 1, 0, 0, 0, 0],   // 秘書／董事會代核人：不帶客戶名、業務名、專案名稱（只放單號；業主 2026-10-08 決定）
    boardProxy: [1, 1, 1, 0, 0, 0, 0],
    consultant: [0, 0, 0, 0, 1, 1, 1],
    owner:      [0, 0, 0, 1, 0, 0, 1],
  };
  eq('KINDS 清單與順序', Array.from(V.KINDS), ['mgr1', 'gm', 'chairman', 'secretary', 'boardProxy', 'consultant', 'owner']);
  t('KINDS 已凍結', Object.isFrozen(V.KINDS));
  Object.keys(EXPECT).forEach((kind) => {
    const v = V.visibilityFor(kind);
    FIELDS.forEach((f, i) => eq('visibilityFor(' + kind + ').' + f + ' = ' + !!EXPECT[kind][i], v[f], !!EXPECT[kind][i]));
    eq('visibilityFor(' + kind + ').itemPrices 恆為 false', v.itemPrices, false);
    t('visibilityFor(' + kind + ').reason 為非空字串', typeof v.reason === 'string' && v.reason.length > 0);
    eq('visibilityFor(' + kind + ') 欄位', Object.keys(v), ['amount', 'margin', 'tier', 'customer', 'owner', 'items', 'project', 'itemPrices', 'reason']);
    t('visibilityFor(' + kind + ') 回傳凍結物件', Object.isFrozen(v));
    t('visibilityFor(' + kind + ') 每次回傳新物件', V.visibilityFor(kind) !== v);
  });
  const unknown = [undefined, null, '', ' ', 'MGR1', 'Mgr1', 'GM', 'Secretary', 'admin', 'user', 'manager1', 'executive', 'signer', 'mgr1 ', ' mgr1', 'mgr1\n', '__proto__', 'constructor', 'toString', 'hasOwnProperty', 'valueOf', 123, {}, [], ['mgr1'], { toString() { return 'mgr1'; } }, true, Symbol('x')];
  unknown.forEach((k) => {
    const v = V.visibilityFor(k);
    t('未知 kind ' + short(k) + ' → 全 false（最小資料）', FIELDS.every((f) => v[f] === false) && v.itemPrices === false && typeof v.reason === 'string', short(v));
    eq('isKnownKind(' + short(k) + ')', V.isKnownKind(k), false);
  });
  V.KINDS.forEach((k) => eq('isKnownKind(' + k + ')', V.isKnownKind(k), true));
  // 不變式
  t('不變式：只有顧問看得到 items', V.KINDS.every((k) => V.visibilityFor(k).items === (k === 'consultant')));
  t('不變式：顧問看不到任何財務欄位', ['amount', 'margin', 'tier'].every((f) => V.visibilityFor('consultant')[f] === false));
  t('不變式：業務看不到金額與毛利', ['amount', 'margin', 'tier'].every((f) => V.visibilityFor('owner')[f] === false));
  t('不變式：秘書與董事會代核人看不到客戶與業務', ['secretary', 'boardProxy'].every((k) => V.visibilityFor(k).customer === false && V.visibilityFor(k).owner === false));
  t('不變式：一級主管／總經理／董事長看得到客戶與業務與金額毛利', ['mgr1', 'gm', 'chairman'].every((k) => ['amount', 'margin', 'tier', 'customer', 'owner'].every((f) => V.visibilityFor(k)[f] === true)));
  t('不變式：amount 與 margin 在所有 kind 一致（不會只給金額不給毛利）', V.KINDS.every((k) => V.visibilityFor(k).amount === V.visibilityFor(k).margin));
  t('不變式：秘書與董事會代核人看不到專案名稱（project 旗標為 false）', ['secretary', 'boardProxy'].every((k) => V.visibilityFor(k).project === false));
  t('不變式：project 只有秘書與董事會代核人為 false，其餘 5 種 kind 為 true', V.KINDS.every((k) => V.visibilityFor(k).project === !(k === 'secretary' || k === 'boardProxy')));
  t('不變式：project 旗標是布林值（不是 truthy 的其他型別）', V.KINDS.every((k) => typeof V.visibilityFor(k).project === 'boolean'));
  t('FIELDS 匯出含 project', Array.isArray(V.FIELDS) && V.FIELDS.indexOf('project') >= 0, short(V.FIELDS));
  // 改動回傳值不影響政策
  {
    const v = V.visibilityFor('consultant');
    try { v.amount = true; } catch (e) { /* strict mode 會 throw */ }
    eq('改動回傳物件不影響下一次查詢', V.visibilityFor('consultant').amount, false);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('5 events', () => {
  const E = load('lib/mail/events.js');
  const mk = (type, over) => {
    const base = {
      type,
      quoteId: 'q-0001',
      quoteNo: 'QU-000000-001',
      projectName: '測試專案',
      company: '測試公司',
      ownerLabel: '業務甲',
      step: { level: 1, label: '一級主管' },
      numbers: { revenueCents: 123456700, gpCents: 50700000, marginText: '41.04%', marginPct: 41.04, tierLevel: 1, tierLabel: '一級主管' },
      actor: { label: '操作者' },
      at: '2026-10-08T10:00:00.000Z',
      stepKey: '2026-10-08T10:00:00.000Z#1',
    };
    if (type === 'E2_COST_REQUEST') { base.step = null; base.numbers = null; base.items = [{ desc: '品項一', qty: 2, unit: '式' }, { desc: '品項二', qty: '10', unit: '人天' }]; }
    if (type === 'E4_RESULT') base.result = { kind: 'final_approved' };
    if (type === 'E5_COST_DONE') { base.step = null; base.numbers = null; }
    if (type === 'E6_WITHDRAWN') base.result = { kind: 'withdrawn', reason: '客戶要求修改' };
    return Object.assign(base, over || {});
  };
  const ok = (name, ev) => { const r = E.validateEvent(ev); t(name, r.ok === true, short(r)); };
  const no = (name, ev, field) => { const r = E.validateEvent(ev); t(name, r.ok === false && typeof r.error === 'string' && (field === undefined || r.field === field), short(r)); };

  eq('EVENT_TYPES', Array.from(E.EVENT_TYPES), ['E1_SUBMIT', 'E2_COST_REQUEST', 'E3_NEXT_STEP', 'E4_RESULT', 'E5_COST_DONE', 'E6_WITHDRAWN']);
  t('EVENT_TYPES 已凍結', Object.isFrozen(E.EVENT_TYPES));
  E.EVENT_TYPES.forEach((type) => ok('合法事件 ' + type, mk(type)));
  ok('E4 approved', mk('E4_RESULT', { result: { kind: 'approved' } }));
  ok('E4 rejected 含原因', mk('E4_RESULT', { result: { kind: 'rejected', reason: '毛利過低' } }));
  ok('E4 returned', mk('E4_RESULT', { result: { kind: 'returned', reason: null } }));
  ok('E6 voided', mk('E6_WITHDRAWN', { result: { kind: 'voided' } }));
  ok('董事會關 step.level="board"', mk('E3_NEXT_STEP', { step: { level: 'board', label: '董事會' }, numbers: { revenueCents: 6000000000, gpCents: 1, marginText: '0.00%', marginPct: 0, tierLevel: null, tierLabel: '董事會' } }));
  ok('E1 numbers 為 null 合法（renderer 會整段不輸出決策條）', mk('E1_SUBMIT', { numbers: null }));
  ok('E1 沒有 numbers 欄位合法', (() => { const e = mk('E1_SUBMIT'); delete e.numbers; return e; })());
  ok('沒有 actor 合法', (() => { const e = mk('E1_SUBMIT'); delete e.actor; return e; })());
  ok('projectName 為空字串合法', mk('E1_SUBMIT', { projectName: '' }));
  ok('projectName 缺少合法', (() => { const e = mk('E1_SUBMIT'); delete e.projectName; return e; })());
  ok('quoteId 為 uuid', mk('E1_SUBMIT', { quoteId: '3f2b8c1e-9a4d-4e6b-8c1f-0a1b2c3d4e5f' }));
  ok('quoteId 為 legacy-12', mk('E1_SUBMIT', { quoteId: 'legacy-12' }));
  ok('quoteId 恰 64 字', mk('E1_SUBMIT', { quoteId: 'x'.repeat(64) }));
  ok('at 帶時區 +08:00', mk('E1_SUBMIT', { at: '2026-10-08T18:00:00+08:00' }));
  ok('at 帶毫秒', mk('E1_SUBMIT', { at: '2026-10-08T10:00:00.123Z' }));
  ok('at 省略秒', mk('E1_SUBMIT', { at: '2026-10-08T10:00Z' }));
  ok('金額 0 分', mk('E1_SUBMIT', { numbers: { revenueCents: 0, gpCents: 0, marginText: '0.00%', marginPct: 0, tierLevel: 1, tierLabel: '一級主管' } }));
  ok('金額 1 分', mk('E1_SUBMIT', { numbers: { revenueCents: 1, gpCents: 0, marginText: '0.00%', marginPct: 0, tierLevel: 1, tierLabel: '一級主管' } }));
  ok('金額 1e12 分', mk('E1_SUBMIT', { numbers: { revenueCents: 1e12, gpCents: 4e11, marginText: '40.00%', marginPct: 40, tierLevel: 3, tierLabel: '董事長' } }));
  ok('最大安全整數', mk('E1_SUBMIT', { numbers: { revenueCents: Number.MAX_SAFE_INTEGER, gpCents: 1, marginText: '1.00%', marginPct: 1, tierLevel: 3, tierLabel: '董事長' } }));
  ok('虧損單：gpCents 為負、毛利率為負', mk('E1_SUBMIT', { numbers: { revenueCents: 100000, gpCents: -5000, marginText: '-5.00%', marginPct: -5, tierLevel: 3, tierLabel: '董事長' } }));
  ok('marginText 不帶 %', mk('E1_SUBMIT', { numbers: { revenueCents: 1, gpCents: 1, marginText: '41.04', marginPct: 41.04, tierLevel: 1, tierLabel: '一級主管' } }));
  ok('未知事件多餘欄位被忽略', mk('E1_SUBMIT', { extra: 'x', price: 5 }));

  // 非法
  [null, undefined, 'x', 5, [], true].forEach((x) => no('事件不是物件 ' + short(x), x));
  [undefined, '', 'E7', 'e1_submit', 'E1_SUBMIT ', 'E1', 5, null, ['E1_SUBMIT']].forEach((x) => no('type 非法 ' + short(x), mk('E1_SUBMIT', { type: x }), 'type'));
  ['', 'a b', 'a/b', '../x', 'a.b', 'x'.repeat(65), '業務', 'a\n', 'a?b', 'a#b', 123, null, undefined, {}].forEach((x) => no('quoteId 非法 ' + short(x), mk('E1_SUBMIT', { quoteId: x }), 'quoteId'));
  ['', '   ', undefined, null, 'x'.repeat(41), 'QU\n1', 5].forEach((x) => no('quoteNo 非法 ' + short(x), mk('E1_SUBMIT', { quoteNo: x }), 'quoteNo'));
  [undefined, null, '', '   ', 'x'.repeat(201), 'a\nb', 5].forEach((x) => no('stepKey 非法 ' + short(x), mk('E1_SUBMIT', { stepKey: x }), 'stepKey'));
  ['yesterday', '2026-10-08', '2026-10-08 10:00:00', '2026-02-30T10:00:00Z', '2026-13-01T10:00:00Z', '2026-10-08T25:00:00Z', 1696723200000, new Date(), null, undefined, '', '2026-10-08T10:00:00', 'x'.repeat(50)]
    .forEach((x) => no('at 非法 ' + short(x), mk('E1_SUBMIT', { at: x }), 'at'));
  [5, {}, 'x'.repeat(301), []].forEach((x) => no('projectName 非法 ' + short(x).slice(0, 30), mk('E1_SUBMIT', { projectName: x }), 'projectName'));
  no('company 太長', mk('E1_SUBMIT', { company: 'x'.repeat(301) }), 'company');
  no('ownerLabel 太長', mk('E1_SUBMIT', { ownerLabel: 'x'.repeat(101) }), 'ownerLabel');
  // numbers
  const nums = (over) => mk('E1_SUBMIT', { numbers: Object.assign({ revenueCents: 100, gpCents: 40, marginText: '40.00%', marginPct: 40, tierLevel: 1, tierLabel: '一級主管' }, over) });
  [-1, -100, 1.5, 0.1, '100', NaN, Infinity, -Infinity, Number.MAX_SAFE_INTEGER + 2, null, undefined, {}, [], true].forEach((x) => no('revenueCents 非法 ' + short(x), nums({ revenueCents: x }), 'numbers.revenueCents'));
  [1.5, '40', NaN, Infinity, null, undefined, Number.MAX_SAFE_INTEGER + 2].forEach((x) => no('gpCents 非法 ' + short(x), nums({ gpCents: x }), 'numbers.gpCents'));
  [NaN, Infinity, -Infinity, '41', null, undefined, {}].forEach((x) => no('marginPct 非法 ' + short(x), nums({ marginPct: x }), 'numbers.marginPct'));
  ['abc', '<script>', '', '1e5%', '41.04%%', '4 1%', 'x'.repeat(30), 5, null, undefined, '--1%', '1,000%'].forEach((x) => no('marginText 非法 ' + short(x), nums({ marginText: x }), 'numbers.marginText'));
  // 極端虧損單：lib/quoteApproval.js 的 marginText 對負毛利沒有上限（成本 ≥ 營收 1 萬倍時整數部分超過 6 位）。
  // 驗證必須放行（簽核信不能因極端數字而寄不出去），信上改由 displayMargin 顯示成 <-999999%。格式檢查（防注入）仍然有效。
  ['-1000000.00', '-100000000.00', '-100000000.00%', '-10000000000.00', '-900719925474099100.00', '12345678901234567890', '-12345678901234567890.0000', '1000000.00'].forEach((x) => {
    ok('極端毛利率 marginText ' + x + ' 合法（長度 ' + x.length + '）', nums({ marginText: x, marginPct: Number(x.replace(/%$/, '')) }));
  });
  ['123456789012345678901', '-123456789012345678901.00', '-100000000.00000', '-1e9', '- 1000000.00', '-100000000.00 ', '1000000,00', '-100000000.00<b>', '<-999999%', '>999999%', '-.5'].forEach((x) => no('極端毛利率仍檢查格式（防注入）：' + short(x), nums({ marginText: x }), 'numbers.marginText'));
  ok('極端虧損單整個事件（營收 1 分、毛利 -1 億分）可通過驗證', mk('E1_SUBMIT', { numbers: { revenueCents: 100, gpCents: -100000000, marginText: '-100000000.00', marginPct: -100000000, tierLevel: 3, tierLabel: '董事長' } }));
  // displayMargin：信上顯示用字串
  {
    const D = E.displayMargin;
    t('events 匯出 displayMargin', typeof D === 'function');
    [
      ['41.04%', '41.04%'], ['41.04', '41.04%'], ['-5.00', '-5.00%'], ['-5.00%', '-5.00%'], ['0.00', '0.00%'], ['100.00', '100.00%'],
      ['999999.99', '999999.99%'], ['-999999.99', '-999999.99%'], ['-999999.99%', '-999999.99%'],
      ['1000000.00', '>999999%'], ['-1000000.00', '<-999999%'], ['-1000000.00%', '<-999999%'], ['-100000000.00', '<-999999%'],
      ['-900719925474099100.00', '<-999999%'], ['12345678901234567890', '>999999%'],
    ].forEach(([input, want]) => eq('displayMargin(' + input + ') → ' + want, D(input), want));
    ['abc', '', '%', null, undefined, 5, {}, [], '-'].forEach((x) => {
      let r; let threw = false;
      try { r = D(x); } catch (e) { threw = true; }
      t('displayMargin(' + short(x) + ') 不 throw、回傳字串', !threw && typeof r === 'string', short(r));
    });
    t('displayMargin 的結果不含 HTML 特殊字元以外的怪字元（只有數字、小數點、負號、<、>、%）', ['-1000000.00', '1000000.00', '12.50', '-0.01'].every((x) => /^[<>\-0-9.%]+$/.test(D(x))));
  }
  [0, 4, '1', 1.5, NaN, 'board'].forEach((x) => no('tierLevel 非法 ' + short(x), nums({ tierLevel: x }), 'numbers.tierLevel'));
  ok('tierLevel 為 null 合法', nums({ tierLevel: null }));
  ok('tierLevel 未提供視為 null', nums({ tierLevel: undefined }));
  ['', undefined, null, 5, 'x'.repeat(21)].forEach((x) => no('tierLabel 非法 ' + short(x), nums({ tierLabel: x }), 'numbers.tierLabel'));
  no('numbers 不是物件', mk('E1_SUBMIT', { numbers: 'x' }), 'numbers');
  // step
  [2.5, '2', 0, 4, 'Board', undefined === 0].forEach((x) => no('step.level 非法 ' + short(x), mk('E3_NEXT_STEP', { step: { level: x, label: 'x' } }), 'step.level'));
  [1, 2, 3, 'board', null].forEach((x) => ok('step.level 合法 ' + short(x), mk('E3_NEXT_STEP', { step: { level: x, label: 'x' } })));
  no('E1 缺 step', mk('E1_SUBMIT', { step: undefined }), 'step');
  no('E1 step 為 null', mk('E1_SUBMIT', { step: null }), 'step');
  no('E3 缺 step', mk('E3_NEXT_STEP', { step: undefined }), 'step');
  no('E1 step.label 空', mk('E1_SUBMIT', { step: { level: 1, label: '' } }), 'step.label');
  no('E3 step.label 空白', mk('E3_NEXT_STEP', { step: { level: 1, label: '  ' } }), 'step.label');
  no('step.label 太長', mk('E1_SUBMIT', { step: { level: 1, label: 'x'.repeat(41) } }), 'step.label');
  no('step 不是物件', mk('E1_SUBMIT', { step: 'x' }), 'step');
  ok('E2 不需要 step', mk('E2_COST_REQUEST', { step: undefined }));
  ok('E5 不需要 step', mk('E5_COST_DONE', { step: undefined }));
  // result
  no('E4 缺 result', mk('E4_RESULT', { result: undefined }), 'result');
  no('E6 缺 result', mk('E6_WITHDRAWN', { result: undefined }), 'result');
  no('E4 kind=withdrawn（與類型不符）', mk('E4_RESULT', { result: { kind: 'withdrawn' } }), 'result.kind');
  no('E4 kind=voided（與類型不符）', mk('E4_RESULT', { result: { kind: 'voided' } }), 'result.kind');
  no('E6 kind=approved（與類型不符）', mk('E6_WITHDRAWN', { result: { kind: 'approved' } }), 'result.kind');
  no('E6 kind=rejected（與類型不符）', mk('E6_WITHDRAWN', { result: { kind: 'rejected' } }), 'result.kind');
  no('result.kind 不在列舉', mk('E4_RESULT', { result: { kind: 'maybe' } }), 'result.kind');
  no('result.kind 缺', mk('E4_RESULT', { result: {} }), 'result.kind');
  no('result.reason 非字串', mk('E4_RESULT', { result: { kind: 'rejected', reason: 5 } }), 'result.reason');
  no('result.reason 太長', mk('E4_RESULT', { result: { kind: 'rejected', reason: 'x'.repeat(2001) } }), 'result.reason');
  ok('result.reason 恰 2000 字', mk('E4_RESULT', { result: { kind: 'rejected', reason: 'x'.repeat(2000) } }));
  no('result 不是物件', mk('E4_RESULT', { result: 'approved' }), 'result');
  // items
  no('items 不是陣列', mk('E2_COST_REQUEST', { items: 'x' }), 'items');
  no('items 超過 200 筆', mk('E2_COST_REQUEST', { items: new Array(201).fill({ desc: 'x', qty: 1, unit: '式' }) }), 'items');
  ok('items 恰 200 筆', mk('E2_COST_REQUEST', { items: new Array(200).fill({ desc: 'x', qty: 1, unit: '式' }) }));
  no('item 缺 desc', mk('E2_COST_REQUEST', { items: [{ qty: 1, unit: '式' }] }), 'items[0].desc');
  no('item desc 空白', mk('E2_COST_REQUEST', { items: [{ desc: '  ', qty: 1 }] }), 'items[0].desc');
  no('item desc 太長', mk('E2_COST_REQUEST', { items: [{ desc: 'x'.repeat(501) }] }), 'items[0].desc');
  no('item unit 太長', mk('E2_COST_REQUEST', { items: [{ desc: 'x', unit: 'u'.repeat(21) }] }), 'items[0].unit');
  [NaN, Infinity, -1, {}, 'x'.repeat(21)].forEach((x) => no('item qty 非法 ' + short(x), mk('E2_COST_REQUEST', { items: [{ desc: 'x', qty: x }] }), 'items[0].qty'));
  no('item 不是物件', mk('E2_COST_REQUEST', { items: ['x'] }), 'items[0]');
  ok('items 為 null', mk('E2_COST_REQUEST', { items: null }));
  ok('item 只有 desc', mk('E2_COST_REQUEST', { items: [{ desc: 'x' }] }));
  // actor
  no('actor.label 空', mk('E1_SUBMIT', { actor: { label: '' } }), 'actor.label');
  no('actor 不是物件', mk('E1_SUBMIT', { actor: 'x' }), 'actor');
  // 永不 throw
  noThrow('validateEvent 遇到會丟例外的 getter 不 throw', () => {
    const e = mk('E1_SUBMIT');
    Object.defineProperty(e, 'quoteNo', { get() { throw new Error('boom'); }, enumerable: true });
    if (E.validateEvent(e).ok !== false) throw new Error('should be invalid');
  });
  noThrow('validateEvent 遇到會丟例外的 Proxy 不 throw', () => {
    const p = new Proxy({}, { get() { throw new Error('boom'); }, getPrototypeOf() { throw new Error('boom'); } });
    if (E.validateEvent(p).ok !== false) throw new Error('should be invalid');
  });
  t('validateEvent 錯誤訊息是繁中字串', /[\u{4e00}-\u{9fff}]/u.test(E.validateEvent(null).error));

  // isValidQuoteId
  ['abc', 'a-b_c', 'legacy-1', '3f2b8c1e-9a4d-4e6b-8c1f-0a1b2c3d4e5f', 'x'.repeat(64)].forEach((x) => eq('isValidQuoteId ' + short(x).slice(0, 30), E.isValidQuoteId(x), true));
  ['', 'x'.repeat(65), 'a b', 'a/b', 'a.b', 5, null, undefined, '../x', 'a\n'].forEach((x) => eq('isValidQuoteId 非法 ' + short(x), E.isValidQuoteId(x), false));

  // dedupeKey
  const ev1 = mk('E1_SUBMIT');
  eq('dedupeKey 格式', E.dedupeKey(ev1, 'user1'), 'E1_SUBMIT:q-0001:user1:2026-10-08T10:00:00.000Z#1');
  eq('dedupeKey 決定性', E.dedupeKey(ev1, 'user1'), E.dedupeKey(mk('E1_SUBMIT'), 'user1'));
  t('dedupeKey 不同收件人不同鍵', E.dedupeKey(ev1, 'user1') !== E.dedupeKey(ev1, 'user2'));
  t('dedupeKey 區分大小寫（U1 與 u1 是不同人）', E.dedupeKey(ev1, 'U1') !== E.dedupeKey(ev1, 'u1'));
  t('dedupeKey 不同 stepKey 不同鍵', E.dedupeKey(ev1, 'u') !== E.dedupeKey(mk('E1_SUBMIT', { stepKey: 'other#1' }), 'u'));
  t('dedupeKey 不同事件類型不同鍵', E.dedupeKey(ev1, 'u') !== E.dedupeKey(mk('E3_NEXT_STEP'), 'u'));
  t('dedupeKey 不同單據不同鍵', E.dedupeKey(ev1, 'u') !== E.dedupeKey(mk('E1_SUBMIT', { quoteId: 'q-0002' }), 'u'));
  t('dedupeKey 不受 projectName／金額變動影響（只看 type、單據、收件人、stepKey）', E.dedupeKey(ev1, 'u') === E.dedupeKey(mk('E1_SUBMIT', { projectName: 'x', numbers: null }), 'u'));
  [undefined, null, '', '   ', 5].forEach((x) => throwsCode('dedupeKey stepKey=' + short(x) + ' → throw NO_STEP_KEY（不給預設）', () => E.dedupeKey(mk('E1_SUBMIT', { stepKey: x }), 'u'), 'NO_STEP_KEY'));
  throwsCode('dedupeKey stepKey 缺欄位', () => { const e = mk('E1_SUBMIT'); delete e.stepKey; return E.dedupeKey(e, 'u'); }, 'NO_STEP_KEY');
  throwsCode('dedupeKey 事件類型非法', () => E.dedupeKey(mk('E1_SUBMIT', { type: 'E9' }), 'u'), 'BAD_EVENT');
  throwsCode('dedupeKey quoteId 非法', () => E.dedupeKey(mk('E1_SUBMIT', { quoteId: 'a/b' }), 'u'), 'BAD_EVENT');
  throwsCode('dedupeKey 事件為 null', () => E.dedupeKey(null, 'u'), 'BAD_EVENT');
  [undefined, null, '', 5, {}, 'a\nb', 'x'.repeat(201)].forEach((x) => throwsCode('dedupeKey username=' + short(x).slice(0, 20) + ' → throw BAD_USERNAME', () => E.dedupeKey(ev1, x), 'BAD_USERNAME'));
  t('dedupeKey 丟的是 MailEventError', (() => { try { E.dedupeKey(ev1, ''); } catch (e) { return e instanceof E.MailEventError && e instanceof Error && e.name === 'MailEventError'; } return false; })());
  t('帳號含冒號：(a:b, c) 與 (a, b:c) 不會撞鍵', E.dedupeKey(mk('E1_SUBMIT', { stepKey: 'c' }), 'a:b') !== E.dedupeKey(mk('E1_SUBMIT', { stepKey: 'b:c' }), 'a'));
  t('帳號含百分號：a%3Ab 與 a:b 不會撞鍵', E.dedupeKey(ev1, 'a%3Ab') !== E.dedupeKey(ev1, 'a:b'));
  // 單射性（injectivity）隨機檢查：不同的 (username, stepKey) 一定產生不同的鍵
  {
    let seed = 8675309;
    const rnd = () => { seed = (seed * 1103515245 + 12345) & 0x7fffffff; return seed / 0x7fffffff; };
    const alpha = ['a', 'b', 'A', ':', '%', '3', '2', '5', '#', '-', 'u', '1'];
    const gen = (min, max) => { let s = ''; const n = min + Math.floor(rnd() * (max - min + 1)); for (let i = 0; i < n; i++) s += alpha[Math.floor(rnd() * alpha.length)]; return s; };
    const seenKeys = new Map();
    let collisions = 0;
    let pairs = 0;
    for (let i = 0; i < 20000; i++) {
      const un = gen(1, 6);
      const sk = gen(1, 6);
      if (sk.trim() === '') continue;
      const key = E.dedupeKey(mk('E1_SUBMIT', { stepKey: sk }), un);
      const id = un + '\x00' + sk;
      if (seenKeys.has(key) && seenKeys.get(key) !== id) collisions++;
      seenKeys.set(key, id);
      pairs++;
    }
    t('dedupeKey 單射性：' + pairs + ' 組隨機 (username, stepKey) 無撞鍵', collisions === 0 && pairs > 1000, collisions);
  }
  t('LIMITS 已凍結', Object.isFrozen(E.LIMITS));
});

// ═════════════════════════════════════════════════════════════════════════
// ▼▼▼ 追加位置（後一階段 agent）▼▼▼
// userEmail ／ link（跳板頁）／ deep-link 的章節，請以 section('名稱', () => { ... }) 的形式追加在這一段註解「之下、finish() 之上」。
// 不要修改上方章節。可用的 helper：t / eq / noThrow / throwsCode / fast / minMs / cp / load(相對路徑) / short。
// deep-link 測試（_client/deep-link.js 是瀏覽器端 IIFE＋module.exports）請用 vm＋假 sessionStorage／location。
// ═════════════════════════════════════════════════════════════════════════
section('6 userEmail', () => {
  const U = load('lib/mail/userEmail.js');
  const R = load('lib/mail/recipients.js');
  const CF = load('lib/mail/config.js');
  const cfg = CF.getMailConfig({});
  const deepFreeze = (o) => {
    if (o && typeof o === 'object' && !Object.isFrozen(o)) { Object.freeze(o); Object.keys(o).forEach((k) => deepFreeze(o[k])); }
    return o;
  };
  // 合成資料（不含任何真實帳號／位址）。Bob／bob 是「只差大小寫的不同帳號」，與實際系統的情況一致
  const mkUsers = () => [
    { username: 'alice', displayName: 'Alice', role: 'user', email: 'test-user-1@itts.com.tw' },
    { username: 'Bob', displayName: 'Bob', role: 'manager1', email: 'test-user-3@itts.com.tw' },
    { username: 'bob', displayName: 'Bob Two', role: 'user' },
    { username: 'carol', displayName: 'Carol', role: 'secretary' },
    { username: 'dave', displayName: 'Dave', role: 'user', email: 'Test-user-7@ITTS.com.tw' },
    { username: 'erin', displayName: 'Erin', role: 'user', active: false, email: 'test-user-8@itts.com.tw' },
    { username: 'frank', displayName: 'Frank', role: 'manager1', email: 'frank@example.test' },
    { username: 'grace', displayName: 'Grace', role: 'user', email: 'not-an-email' },
    { username: 'pool1', role: 'pool' },
    { username: 'heidi', nickname: 'Hei', displayName: 'Heidi D', role: 'user', email: '   ' },
    { username: 'Mary Lee', displayName: 'Mary', role: 'user' },
    { username: 'view1', displayName: 'Viewer', role: 'manager1', accessMode: 'view' },
  ];
  const V = (raw, extra) => U.validateUserEmail(raw, Object.assign({ config: cfg, users: mkUsers() }, extra));

  // ── A. validateUserEmail ──
  eq('合法位址：trim＋轉小寫', V('  New.User+tag@ITTS.com.tw ', { selfUsername: 'bob' }), { ok: true, value: 'new.user+tag@itts.com.tw', code: 'OK' });
  [null, '', '   ', '\t\n', cp(0x3000), cp(0xa0)].forEach((raw) => {
    eq('清除（允許）：' + short(raw), V(raw), { ok: true, value: '', code: 'CLEARED' });
  });
  { const r = V(undefined); t('undefined → MISSING（不當成清除，避免路由漏帶欄位時清掉所有人）', r.ok === false && r.code === 'MISSING', short(r)); }
  [0, false, 123, {}, ['a@itts.com.tw'], Symbol('x')].forEach((raw) => {
    const r = V(raw);
    t('非字串不是清除：' + short(raw), r.ok === false && r.code === 'NOT_STRING', short(r));
  });

  // 唯一性
  const dupCase = [
    ['TEST-USER-1@itts.com.tw', 'bob', 'alice', '不分大小寫（TEST-USER-1@ 與 test-user-1@）'],
    ['test-user-3@itts.com.tw', 'bob', 'Bob', '只差大小寫的另一個帳號持有'],
    ['test-user-8@itts.com.tw', 'alice', 'erin', '停用帳號持有的位址也算占用'],
    ['test-user-7@itts.com.tw', 'alice', 'dave', '已存值未正規化（Test-user-7@ITTS）仍以正規化比對'],
    ['  TEST-USER-7@itts.com.tw  ', 'alice', 'dave', '輸入前後空白'],
  ];
  dupCase.forEach(([raw, self, holder, name]) => {
    const r = V(raw, { selfUsername: self });
    t('唯一性：' + name, r.ok === false && r.code === 'DUPLICATE' && r.conflictUsername === holder, short(r));
    t('唯一性：錯誤訊息是固定文字（不含帳號與位址）', typeof r.error === 'string' && r.error.indexOf(holder) < 0 && r.error.indexOf('@') < 0, short(r.error));
  });
  eq('唯一性：排除自己（alice 改成 TEST-USER-1@）', V('Test-user-1@ITTS.com.tw', { selfUsername: 'alice' }), { ok: true, value: 'test-user-1@itts.com.tw', code: 'OK' });
  eq('唯一性：Bob 本人沿用自己的位址', V('test-user-3@itts.com.tw', { selfUsername: 'Bob' }), { ok: true, value: 'test-user-3@itts.com.tw', code: 'OK' });
  eq('唯一性：未帶 selfUsername 時不排除任何人', V('test-user-1@itts.com.tw').code, 'DUPLICATE');
  eq('唯一性：髒資料（grace 的 not-an-email、heidi 的空白）不影響其他位址', V('free1@itts.com.tw', { selfUsername: 'x' }).ok, true);
  eq('users 為 username 鍵物件（ctx.users 形態）', V('test-user-1@itts.com.tw', { users: { alice: { email: 'test-user-1@itts.com.tw' } }, selfUsername: 'zed' }).conflictUsername, 'alice');
  eq('users 帳號叫 __proto__ 也安全', V('p@itts.com.tw', { users: [{ username: '__proto__', email: 'p@itts.com.tw' }], selfUsername: 'x' }).conflictUsername, '__proto__');
  eq('沒有 users 就不檢查唯一性', U.validateUserEmail('test-user-1@itts.com.tw', { config: cfg }).ok, true);
  eq('opts 缺省不 throw', U.validateUserEmail('a@itts.com.tw').ok, true);

  // 網域與格式（沿用 safety 的規則，這裡確認有串起來）
  [
    ['x@example.test', 'DOMAIN_NOT_ALLOWED'], ['x@evil-itts.com.tw', 'DOMAIN_NOT_ALLOWED'], ['x@itts.com.tw.evil.test', 'DOMAIN_NOT_ALLOWED'],
    ['x@sub.itts.com.tw', 'DOMAIN_NOT_ALLOWED'], ['x@ITTS.COM.TW.', 'BAD_DOMAIN'],
    ['a@itts.com.tw\nBcc: b@itts.com.tw', 'CONTROL_CHAR'], ['a@itts.com.tw,b@itts.com.tw', 'BAD_CHAR'], ['a@itts.com.tw;b@itts.com.tw', 'BAD_CHAR'],
    [cp(0xff41) + '@itts.com.tw', 'NON_ASCII'], ['a b@itts.com.tw', 'SPACE'], ['no-at-sign', 'BAD_FORMAT'], ['a@@itts.com.tw', 'BAD_FORMAT'],
    ['<script>alert(1)</script>@itts.com.tw', 'BAD_CHAR'], ['"><svg/onload=alert(1)>@itts.com.tw', 'BAD_CHAR'],
  ].forEach(([raw, code]) => {
    const r = V(raw);
    t('拒絕 ' + short(raw) + ' → ' + code, r.ok === false && r.code === code, short(r));
  });
  ['<img src=x onerror=alert(1)>@itts.com.tw', '"><svg/onload=alert(1)>@evil.test', 'javascript:alert(1)'].forEach((raw) => {
    const r = V(raw);
    t('錯誤訊息不回顯輸入：' + short(raw), r.ok === false && r.error.indexOf(raw.slice(0, 6)) < 0, short(r.error));
  });
  eq('自訂網域白名單', U.validateUserEmail('x@example.test', { config: { allowedDomains: ['example.test'] } }).ok, true);
  eq('自訂白名單時 itts.com.tw 不再通過', U.validateUserEmail('x@itts.com.tw', { config: { allowedDomains: ['example.test'] } }).code, 'DOMAIN_NOT_ALLOWED');
  eq('空白名單＝全部拒絕（fail closed）', U.validateUserEmail('x@itts.com.tw', { config: { allowedDomains: [] } }).code, 'DOMAIN_NOT_ALLOWED');
  eq('config 缺 allowedDomains → 用預設白名單', U.validateUserEmail('x@itts.com.tw', { config: {} }).ok, true);
  t('網域錯誤訊息列出允許的網域', V('x@example.test').error.indexOf('itts.com.tw') >= 0, V('x@example.test').error);
  {
    const big = Array.from({ length: 5000 }, (_, i) => ({ username: 'u' + i, email: 'u' + i + '@itts.com.tw' }));
    fast('validateUserEmail（5000 個帳號）', () => U.validateUserEmail('z1@itts.com.tw', { config: cfg, users: big }), 100);
    fast('validateUserEmail（200 萬字元輸入）', () => U.validateUserEmail('a'.repeat(2e6) + '@itts.com.tw', { config: cfg, users: big }), 50);
    fast('validateUserEmail（200 萬個空白＝清除）', () => U.validateUserEmail(' '.repeat(2e6), { config: cfg }), 50);
  }

  // ── B. parseBulkEmailText ──
  const P = (text, extra) => U.parseBulkEmailText(text, Object.assign({ config: cfg, users: mkUsers() }, extra));
  const sts = (res) => res.rows.map((r) => r.status).join(',');
  {
    const p = P('alice,test-user-2@itts.com.tw\nBob\ttest-user-4@itts.com.tw\ncarol test-user-6@itts.com.tw');
    eq('三種分隔（逗號／Tab／空白）', sts(p), 'ok,ok,ok');
    eq('rows[0] 形狀', p.rows[0], { line: 1, username: 'alice', email: 'test-user-2@itts.com.tw', status: 'ok' });
    eq('summary', p.summary, { ok: 3, unchanged: 0, error: 0 });
    eq('fatal=false', p.fatal, false);
  }
  {
    const p = P(cp(0xfeff) + 'alice,test-user-2@itts.com.tw\r\n\r\n# 註解\r\n   \r\nBob,test-user-4@itts.com.tw\r\n');
    eq('BOM＋CRLF＋空行＋註解：狀態', sts(p), 'ok,ok');
    eq('行號是原檔實體行號', p.rows.map((r) => r.line), [1, 5]);
    eq('BOM 沒有黏進帳號', p.rows[0].username, 'alice');
  }
  eq('只有 CR 的換行', sts(P('alice,test-user-2@itts.com.tw\rBob,test-user-4@itts.com.tw')), 'ok,ok');
  [
    ['全形逗號', 'alice' + cp(0xff0c) + 'test-user-2@itts.com.tw'], ['全形空白', 'alice' + cp(0x3000) + 'test-user-2@itts.com.tw'],
    ['分號', 'alice;test-user-2@itts.com.tw'], ['全形分號', 'alice' + cp(0xff1b) + 'test-user-2@itts.com.tw'],
    ['混合多個分隔', 'alice , \t ' + cp(0xff0c) + ' test-user-2@itts.com.tw'], ['行尾多餘逗號', 'alice,test-user-2@itts.com.tw,,'],
    ['雙引號 CSV', '"alice","test-user-2@itts.com.tw"'], ['彎引號', cp(0x201c) + 'alice' + cp(0x201d) + ',' + cp(0x201c) + 'test-user-2@itts.com.tw' + cp(0x201d)],
    ['前後空白', '   alice  ,  test-user-2@itts.com.tw   '], ['Email 大寫', 'alice,TEST-USER-2@ITTS.COM.TW'],
  ].forEach(([name, line]) => {
    const p = P(line);
    t('髒輸入可解析：' + name, sts(p) === 'ok' && p.rows[0].username === 'alice' && p.rows[0].email === 'test-user-2@itts.com.tw', short(p.rows));
  });
  ['Mary Lee,test-user-9@itts.com.tw', 'Mary Lee\ttest-user-9@itts.com.tw', 'Mary Lee test-user-9@itts.com.tw', '"Mary Lee","test-user-9@itts.com.tw"'].forEach((line) => {
    const p = P(line);
    t('帳號含空白：' + short(line), sts(p) === 'ok' && p.rows[0].username === 'Mary Lee', short(p.rows));
  });
  {
    const p = P('dave,TEST-USER-7@itts.com.tw\nalice,TEST-USER-1@itts.com.tw');
    eq('同一個位址（大小寫不同）→ unchanged', sts(p), 'unchanged,unchanged');
    eq('summary unchanged', p.summary, { ok: 0, unchanged: 2, error: 0 });
  }
  {
    const p = P('bob,test-user-5@itts.com.tw\nBob,test-user-4@itts.com.tw');
    eq('只差大小寫的兩個帳號各自獨立', p.rows.map((r) => r.username + ':' + r.status), ['bob:ok', 'Bob:ok']);
  }
  {
    const p = P('nobody,a@itts.com.tw\nALICE,x@itts.com.tw\nBOB,y@itts.com.tw');
    eq('帳號不存在的錯誤碼', p.rows.map((r) => r.code), ['UNKNOWN_USER', 'UNKNOWN_USER', 'UNKNOWN_USER']);
    t('只差大小寫時提示「區分大小寫」，否則不提示', p.rows[1].error.indexOf('大小寫') >= 0 && p.rows[0].error.indexOf('大小寫') < 0, short(p.rows.map((r) => r.error)));
    t('提示不洩漏是哪個帳號', p.rows[1].error.indexOf('alice') < 0 && p.rows[2].error.indexOf('Bob') < 0, short(p.rows[1].error));
  }
  [
    ['alice,a@evil-itts.com.tw', 'DOMAIN_NOT_ALLOWED'], ['alice,not-an-email', 'BAD_FORMAT'], ['alice', 'FORMAT'], ['alice,', 'FORMAT'], [',', 'FORMAT'],
    ['alice,""', 'FORMAT'], ['a@itts.com.tw alice', 'SWAPPED'], ['alice,a@itts.com.tw,b@itts.com.tw', 'UNKNOWN_USER'],
    ['alice,a\x00@itts.com.tw', 'CONTROL_CHAR'], ['bob,test-user-1@itts.com.tw', 'DUPLICATE'], ['bob,TEST-USER-1@ITTS.com.tw', 'DUPLICATE'],
    ['alice,' + 'a'.repeat(1001) + '@itts.com.tw', 'LINE_TOO_LONG'], ['alice,<script>@itts.com.tw', 'BAD_CHAR'],
  ].forEach(([line, code]) => {
    const p = P(line);
    t('錯誤行 ' + short(line.length > 40 ? line.slice(0, 40) + '…' : line) + ' → ' + code, sts(p) === 'error' && p.rows[0].code === code, short(p.rows[0]));
  });
  {
    const p = P('bob,test-user-1@itts.com.tw');
    eq('與現有帳號衝突時附 conflictUsername', p.rows[0].conflictUsername, 'alice');
  }
  {
    const p = P('carol,shared@itts.com.tw\nMary Lee,SHARED@itts.com.tw\nalice,solo@itts.com.tw');
    eq('同批內 Email 重複：涉及的每一行都錯誤，其他行不受影響', p.rows.map((r) => r.status + ':' + (r.code || '')), ['error:DUP_EMAIL_IN_BATCH', 'error:DUP_EMAIL_IN_BATCH', 'ok:']);
    eq('同批內重複的 summary', p.summary, { ok: 1, unchanged: 0, error: 2 });
  }
  {
    const p = P('carol,c1@itts.com.tw\ncarol,c2@itts.com.tw\nalice,solo@itts.com.tw');
    eq('同一帳號出現多次：每一行都錯誤', p.rows.map((r) => r.status + ':' + (r.code || '')), ['error:DUP_USER_IN_BATCH', 'error:DUP_USER_IN_BATCH', 'ok:']);
  }
  {
    // 已經是錯誤的行不參與同批重複判斷（它本來就不會被套用）
    const p = P('nobody,same@itts.com.tw\ncarol,same@itts.com.tw');
    eq('錯誤行不連累合法行', p.rows.map((r) => r.status), ['error', 'ok']);
  }
  {
    const p = P('alice,<img src=x onerror=alert(1)>@itts.com.tw\n<img src=x onerror=alert(1)>,a@itts.com.tw');
    t('惡意內容只會落在 username／email 欄位且無控制字元', p.rows.length === 2 && p.rows.every((r) => r.status === 'error' && !/[\x00-\x08\x0b-\x1f\x7f]/.test(r.username + r.email)), short(p.rows));
    t('錯誤訊息是固定文字，不回顯輸入', p.rows.every((r) => r.error.indexOf('onerror') < 0), short(p.rows.map((r) => r.error)));
  }
  {
    const pu = [{ username: '__proto__' }, { username: 'constructor', email: 'c@itts.com.tw' }];
    const p = P('__proto__,p@itts.com.tw\nconstructor,c@itts.com.tw\ntoString,t@itts.com.tw', { users: pu });
    eq('原型污染名稱', p.rows.map((r) => r.username + ':' + r.status), ['__proto__:ok', 'constructor:unchanged', 'toString:error']);
  }
  // 上限
  {
    const many = Array.from({ length: 600 }, (_, i) => ({ username: 'u' + i }));
    const line = (i) => 'u' + i + ',u' + i + '@itts.com.tw';
    const text500 = Array.from({ length: 500 }, (_, i) => line(i)).join('\n');
    const p500 = P(text500, { users: many });
    t('剛好 500 行資料可通過', p500.fatal === false && p500.summary.ok === 500 && p500.rows.length === 500, short(p500.summary));
    const p501 = P(text500 + '\n' + line(500), { users: many });
    t('501 行整批拒絕（fatal TOO_MANY_ROWS，沒有半套）', p501.fatal === true && p501.rows.length === 1 && p501.rows[0].line === 0 && p501.rows[0].code === 'TOO_MANY_ROWS' && p501.summary.error === 1 && p501.summary.ok === 0, short(p501));
    const withNoise = P(Array.from({ length: 1000 }, (_, i) => (i % 2 ? '# c' : '')).join('\n') + '\n' + text500, { users: many });
    t('空行與註解不計入 500 行上限', withNoise.fatal === false && withNoise.summary.ok === 500, short(withNoise.summary));
    fast('parseBulkEmailText（500 行 × 600 帳號）', () => P(text500, { users: many }), 100);
    const bigUsers = Array.from({ length: 3000 }, (_, i) => ({ username: 'v' + i, email: 'v' + i + '@itts.com.tw' }));
    const text500b = Array.from({ length: 500 }, (_, i) => 'v' + i + ',w' + i + '@itts.com.tw').join('\n');
    fast('parseBulkEmailText（500 行 × 3000 帳號）', () => P(text500b, { users: bigUsers }), 300);
  }
  {
    const pl = P('x'.repeat(200001));
    t('全文過大（>200000 字元）整批拒絕', pl.fatal === true && pl.rows[0].code === 'TOO_LARGE', short(pl.rows[0]));
    fast('parseBulkEmailText（200 萬字元單行）', () => P('a'.repeat(2e6)), 50);
    [null, undefined, 123, {}, ['a,b']].forEach((x) => {
      const r = P(x);
      t('非字串輸入 → fatal NOT_STRING：' + short(x), r.fatal === true && r.rows[0].code === 'NOT_STRING', short(r));
    });
    eq('空字串：沒有資料行', P(''), { rows: [], summary: { ok: 0, unchanged: 0, error: 0 }, fatal: false });
    eq('只有註解與空行', P('# a\n\n   \n# b').rows, []);
    t('整批拒絕時沒有任何 ok 行（不會被誤套用）', pl.rows.every((r) => r.status === 'error'), short(pl.rows));
  }
  {
    const frozen = deepFreeze(mkUsers());
    const before = JSON.stringify(frozen);
    noThrow('不修改傳入的 users（凍結物件也不 throw）', () => P('alice,test-user-2@itts.com.tw', { users: frozen }));
    eq('傳入的 users 內容不變', JSON.stringify(frozen), before);
    noThrow('users 缺省／null／字串不 throw', () => { P('alice,a@itts.com.tw', { users: null }); P('alice,a@itts.com.tw', { users: 'x' }); U.parseBulkEmailText('alice,a@itts.com.tw'); });
    eq('users 缺省時每一行都是 UNKNOWN_USER', U.parseBulkEmailText('alice,a@itts.com.tw').rows[0].code, 'UNKNOWN_USER');
  }

  // ── C. applyBulkPlan ──
  {
    const frozen = deepFreeze(mkUsers());
    const plan = P('alice,test-user-2@itts.com.tw\nbob,test-user-4@itts.com.tw\ndave,TEST-USER-7@itts.com.tw\nnobody,n@itts.com.tw\nMary Lee,test-user-9@itts.com.tw', { users: frozen });
    let ap;
    noThrow('套用（輸入已凍結也不 throw）', () => { ap = U.applyBulkPlan(plan.rows, frozen, { config: cfg }); });
    eq('updated：只有 ok 行（unchanged 與 error 不動）', ap.updated.map((x) => x.username), ['alice', 'bob', 'Mary Lee']);
    eq('skipped 為空', ap.skipped, []);
    t('回傳新陣列且每個 user 都是新物件', ap.users !== frozen && ap.users.length === frozen.length && ap.users.every((u, i) => u !== frozen[i]));
    const byName = (name) => ap.users.find((u) => u.username === name);
    eq('alice 已更新', byName('alice').email, 'test-user-2@itts.com.tw');
    eq('小寫 bob 已更新，Bob 不受影響', [byName('bob').email, byName('Bob').email], ['test-user-4@itts.com.tw', 'test-user-3@itts.com.tw']);
    eq('unchanged 的 dave 保留原值（不重寫）', byName('dave').email, 'Test-user-7@ITTS.com.tw');
    eq('只改 email：其他欄位原樣', Object.assign({}, byName('alice'), { email: 'test-user-1@itts.com.tw' }), frozen[0]);
    eq('updated 帶 previous 與稽核文字', ap.updated[0], { username: 'alice', previous: 'test-user-1@itts.com.tw', email: 'test-user-2@itts.com.tw', detail: 'email 已變更' });
    eq('原本沒有 email → 未設定→已設定', ap.updated[1].detail, 'email 未設定→已設定');
    t('稽核文字不含 @', ap.updated.every((x) => x.detail.indexOf('@') < 0), short(ap.updated.map((x) => x.detail)));
    const again = U.applyBulkPlan(plan.rows, ap.users, { config: cfg });
    eq('重複套用同一份計畫＝沒有變更（冪等）', [again.updated.length, again.skipped.length], [0, 0]);
  }
  {
    const tampered = [
      { line: 1, username: 'alice', email: 'x@evil.test', status: 'ok' },
      { line: 2, username: 'ghost', email: 'g@itts.com.tw', status: 'ok' },
      { line: 3, username: 'carol', email: 'dup@itts.com.tw', status: 'ok' },
      { line: 4, username: 'Mary Lee', email: 'DUP@itts.com.tw', status: 'ok' },
      { line: 5, username: 'bob', email: '', status: 'ok' },
      { line: 6, username: 'dave', email: 'whatever@itts.com.tw', status: 'error' },
      { line: 7, username: 'heidi', email: 'a@itts.com.tw\nBcc: z@itts.com.tw', status: 'ok' },
      { line: 8, username: 'alice', email: 'test-user-3@itts.com.tw', status: 'ok' },
      { line: 9, username: 'erin', email: 'e2@itts.com.tw', status: 'unchanged' },
      { line: 10, username: 'view1', status: 'ok' },
      null, 'str', 5,
    ];
    const src = mkUsers();
    const ap = U.applyBulkPlan(tampered, src, { config: cfg });
    eq('竄改的 rows：只有合法的 carol 被更新', ap.updated.map((x) => x.username), ['carol']);
    eq('竄改的 rows：skipped 的原因', ap.skipped.map((x) => x.username + ':' + x.code), [
      'alice:DOMAIN_NOT_ALLOWED', 'ghost:UNKNOWN_USER', 'Mary Lee:DUPLICATE', 'bob:EMPTY', 'heidi:CONTROL_CHAR', 'alice:DUPLICATE', 'view1:MISSING',
    ]);
    const g = (n) => ap.users.find((u) => u.username === n);
    t('被拒絕的人原值不變、空字串不會清除別人', g('alice').email === 'test-user-1@itts.com.tw' && g('bob').email === undefined && g('dave').email === 'Test-user-7@ITTS.com.tw');
    eq('輸入的 users 沒被就地修改', src, mkUsers());
    ap.users[0].displayName = 'changed';
    eq('改輸出不影響輸入（深度獨立到第一層）', src[0].displayName, 'Alice');
  }
  {
    const rows = [{ username: 'alice', email: 'new1@itts.com.tw', status: 'ok' }, { username: 'bob', email: 'TEST-USER-1@itts.com.tw', status: 'ok' }];
    const ap = U.applyBulkPlan(rows, mkUsers(), { config: cfg });
    eq('演進中的名冊：前一行釋出的位址，後一行可以接手', ap.updated.map((x) => x.username + '>' + x.email), ['alice>new1@itts.com.tw', 'bob>test-user-1@itts.com.tw']);
    const rows2 = [{ username: 'carol', email: 'same@itts.com.tw', status: 'ok' }, { username: 'bob', email: 'same@itts.com.tw', status: 'ok' }];
    const ap2 = U.applyBulkPlan(rows2, mkUsers(), { config: cfg });
    eq('演進中的名冊：同批後一行搶同一位址 → skipped DUPLICATE', [ap2.updated.length, ap2.skipped.map((x) => x.code)], [1, ['DUPLICATE']]);
  }
  {
    const obj = { alice: { username: 'alice' } };
    const r1 = U.applyBulkPlan([{ username: 'alice', email: 'a@itts.com.tw', status: 'ok' }], obj);
    t('users 不是陣列：原樣回傳＋error，不做任何變更', r1.users === obj && typeof r1.error === 'string' && r1.updated.length === 0 && obj.alice.email === undefined, short(r1));
    const r2 = U.applyBulkPlan('nope', mkUsers());
    t('rows 不是陣列：回傳 users 副本＋error', Array.isArray(r2.users) && r2.users.length === 12 && typeof r2.error === 'string', short(r2.error));
    const r3 = U.applyBulkPlan(Array.from({ length: 501 }, () => ({ username: 'alice', email: 'q@itts.com.tw', status: 'ok' })), mkUsers(), { config: cfg });
    t('rows 超過 500：不做任何變更', r3.updated.length === 0 && typeof r3.error === 'string' && r3.users[0].email === 'test-user-1@itts.com.tw', short(r3.error));
    noThrow('applyBulkPlan 缺參數不 throw', () => { U.applyBulkPlan(); U.applyBulkPlan([], []); U.applyBulkPlan(null, null, null); });
    const r4 = U.applyBulkPlan([{ username: 'Mary  Lee', email: 'm@itts.com.tw', status: 'ok' }], [{ username: 'Mary  Lee' }]);
    eq('帳號含連續空白時以原樣比對', r4.updated.length, 1);
  }

  // ── D. maskedAuditDetail ──
  {
    const TXT = ['email 未設定→已設定', 'email 已變更', 'email 已清除', 'email 無變更'];
    const M = U.maskedAuditDetail;
    eq('未設定→已設定', [M('', 'a@itts.com.tw'), M(undefined, 'a@itts.com.tw'), M(null, 'a@itts.com.tw'), M('   ', 'a@itts.com.tw')], Array(4).fill(TXT[0]));
    eq('已變更', M('a@itts.com.tw', 'b@itts.com.tw'), TXT[1]);
    eq('已清除', [M('a@itts.com.tw', ''), M('a@itts.com.tw', null), M('a@itts.com.tw', undefined)], Array(3).fill(TXT[2]));
    eq('無變更（大小寫不同＝同一位址）', [M('a@itts.com.tw', 'A@ITTS.com.tw'), M('', ''), M(null, undefined)], Array(3).fill(TXT[3]));
    eq('髒資料→合法位址＝已變更；同一個髒值＝無變更', [M('garbage', 'a@itts.com.tw'), M('garbage', ' garbage ')], [TXT[1], TXT[3]]);
    eq('非字串視為未設定', [M(123, 456), M({}, [])], [TXT[3], TXT[3]]);
    let leaks = 0;
    let n = 0;
    const pool = ['', null, 'a@itts.com.tw', 'Mixed.Case+tag@itts.com.tw', 'x@example.test', 'garbage', '<b>@itts.com.tw', 123];
    pool.forEach((a) => pool.forEach((b) => { n++; const s = M(a, b); if (TXT.indexOf(s) < 0 || s.indexOf('@') >= 0) leaks++; }));
    t('輸出只會是四句固定文字之一（共 ' + n + ' 組輸入），不可能含位址', leaks === 0, leaks);
  }

  // ── E. missingEmailReport ──
  const roster = {
    gm: ['dave', 'heidi', 'ghost', 'erin', 'carol'], chairman: ['alice', 'carol'], boardProxy: ['grace', 'heidi'],
    costProviders: ['carol', 'pool1', 'Mary Lee', 'carol'], sealManagers: ['bob'],
  };
  {
    const rep = U.missingEmailReport({ users: mkUsers(), roster, config: cfg });
    eq('缺 Email 名單（依帳號 code unit 排序、每人一筆）', rep.map((r) => r.username), ['Mary Lee', 'carol', 'frank', 'grace', 'heidi', 'view1']);
    const g = (n) => rep.find((r) => r.username === n);
    eq('carol：secretary 角色＋三個名冊（角色依固定順序）', [g('carol').roles, g('carol').reason], [['gm', 'chairman', 'secretary', 'costProvider'], 'NO_EMAIL']);
    eq('carol 的中文角色名', g('carol').roleLabels, ['總經理', '董事長', '秘書', '成本填寫人']);
    eq('frank：manager1＋網域不在白名單', [g('frank').roles, g('frank').reason], [['manager1'], 'DOMAIN_NOT_ALLOWED']);
    eq('grace：boardProxy＋格式非法', [g('grace').roles, g('grace').reason], [['boardProxy'], 'BAD_EMAIL']);
    eq('heidi：gm＋boardProxy，空白視為沒有，label 用暱稱', [g('heidi').roles, g('heidi').reason, g('heidi').label], [['gm', 'boardProxy'], 'NO_EMAIL', 'Hei']);
    eq('Mary Lee：只有 costProvider', [g('Mary Lee').roles, g('Mary Lee').reason], [['costProvider'], 'NO_EMAIL']);
    eq('唯讀帳號 view1 仍列出並標 readOnly', [g('view1').roles, g('view1').readOnly, g('frank').readOnly], [['manager1'], true, false]);
    t('不列：有可用 Email 的人（Bob、dave、alice）', !g('Bob') && !g('dave') && !g('alice'));
    t('不列：停用帳號（erin）與客戶池帳號（pool1）', !g('erin') && !g('pool1'));
    t('不列：名冊裡不存在的帳號（ghost）', !g('ghost'));
    t('不列：只在 sealManagers 的人（bob）與沒有任何簽核角色的人', !g('bob'));
    eq('報表物件欄位固定', Object.keys(rep[0]).sort(), ['label', 'readOnly', 'reason', 'roleLabels', 'roles', 'username']);
    eq('ROLE_KEYS 與中文標籤一一對應', U.ROLE_KEYS.map((k) => !!U.ROLE_LABELS[k]), Array(6).fill(true));
  }
  eq('沒有 roster：只有角色型（manager1／secretary）', U.missingEmailReport({ users: mkUsers(), config: cfg }).map((r) => r.username), ['carol', 'frank', 'view1']);
  eq('roster 欄位不是陣列或內含雜物：忽略',
    U.missingEmailReport({ users: mkUsers(), config: cfg, roster: { gm: 'heidi', chairman: [null, 5, {}, '', 'heidi'], boardProxy: { 0: 'grace' }, costProviders: 7 } }).map((r) => r.username + ':' + r.roles.join('/')),
    ['carol:secretary', 'frank:manager1', 'heidi:chairman', 'view1:manager1']);
  eq('users 為 username 鍵物件', U.missingEmailReport({ users: { x: { username: 'x', role: 'manager1' }, y: { role: 'secretary', email: 'y@itts.com.tw' } }, config: cfg }).map((r) => r.username), ['x']);
  eq('空輸入不 throw', [U.missingEmailReport(), U.missingEmailReport({}), U.missingEmailReport({ users: null, roster: 5 }), U.missingEmailReport('x')], [[], [], [], []]);
  noThrow('輸入已凍結也不 throw', () => U.missingEmailReport({ users: deepFreeze(mkUsers()), roster: deepFreeze(JSON.parse(JSON.stringify(roster))), config: cfg }));
  eq('admin 放進名冊仍會被列出（quoteRoutes 的 stepRecipients 不過濾 admin）', U.missingEmailReport({ users: [{ username: 'adm', role: 'admin' }], roster: { gm: ['adm'] }, config: cfg }).map((r) => r.username), ['adm']);
  {
    // 與 resolveRecipients 對拍：同一個帳號，「缺 Email 報表」與「寄信時被略過的原因」必須一致
    const emails = [undefined, null, '', '   ', 'a@itts.com.tw', 'A@ITTS.COM.TW', 'a@evil.test', 'a@itts.com.tw.evil.test', 'a b@itts.com.tw',
      'a@itts.com.tw,b@itts.com.tw', 'x'.repeat(300) + '@itts.com.tw', 123, {}, ['a@itts.com.tw'], cp(0xff41) + '@itts.com.tw', 'a@sub.itts.com.tw', 'a@itts.com.tw\n'];
    const variants = [{}, { active: false }, { disabled: true }, { role: 'pool' }];
    let diff = 0;
    let cases = 0;
    emails.forEach((email) => variants.forEach((v) => {
      cases++;
      const u = Object.assign({ username: 'u', role: 'manager1', email }, v);
      const rep = U.missingEmailReport({ users: [u], config: cfg });
      const rr = R.resolveRecipients(['u'], { users: [u], config: cfg });
      const reason = rr.skipped.length ? rr.skipped[0].reason : '';
      const expectReport = ['NO_EMAIL', 'BAD_EMAIL', 'DOMAIN_NOT_ALLOWED'].indexOf(reason) >= 0;
      const ok = expectReport ? (rep.length === 1 && rep[0].reason === reason) : rep.length === 0;
      if (!ok) { diff++; record('對拍不一致 email=' + short(email) + ' v=' + short(v), false, 'report=' + short(rep) + ' resolve=' + short(rr)); }
    }));
    t('缺 Email 報表與 resolveRecipients 的略過原因完全一致（' + cases + ' 組）', diff === 0, diff);
    const labelCases = [{ nickname: 'N', displayName: 'D' }, { nickname: '  ', displayName: 'D' }, { displayName: 'D' }, {}, { nickname: 'x'.repeat(100) },
      { displayName: 'a\nb' + cp(0x202e) + 'c' }, { nickname: 5, displayName: 'D' }, { displayName: '  \t ' }];
    let ld = 0;
    labelCases.forEach((c) => {
      const withMail = Object.assign({ username: 'lbl', role: 'manager1', email: 'a@itts.com.tw' }, c);
      const noMail = Object.assign({ username: 'lbl', role: 'manager1' }, c);
      const l1 = R.resolveRecipients(['lbl'], { users: [withMail], config: cfg }).deliver[0].label;
      const l2 = U.missingEmailReport({ users: [noMail], config: cfg })[0].label;
      if (l1 !== l2) { ld++; record('label 對拍不一致 ' + short(c), false, short([l1, l2])); }
    });
    t('label 規則與 resolveRecipients 一致（暱稱 > 顯示名稱 > 帳號，safeText）', ld === 0, ld);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('7 link／跳板頁', () => {
  const L = load('lib/mail/link.js');
  const E = load('lib/mail/events.js');
  const CF = load('lib/mail/config.js');
  const DL = load('_client/deep-link.js');
  const cfg = CF.getMailConfig({});
  const UUID = '0f8fad5b-d9cb-469f-a165-70867728950e';       // 合成的 uuid 樣式字串
  const SAFE = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-';
  const isValidOracle = (s) => {                              // 獨立於正規式的手寫判斷
    if (typeof s !== 'string' || s.length < 1 || s.length > 64) return false;
    for (let i = 0; i < s.length; i++) if (SAFE.indexOf(s.charAt(i)) < 0) return false;
    return true;
  };
  const H = (r, name) => {
    const k = Object.keys(r.headers).find((x) => x.toLowerCase() === name.toLowerCase());
    return k === undefined ? undefined : r.headers[k];
  };
  const mulberry = (a) => () => { a |= 0; a = (a + 0x6d2b79f5) | 0; let x = Math.imul(a ^ (a >>> 15), 1 | a); x = (x + Math.imul(x ^ (x >>> 7), 61 | x)) ^ x; return ((x ^ (x >>> 14)) >>> 0) / 4294967296; };

  // ── isValidQuoteId ──
  t('isValidQuoteId 與 events.js 是同一個函式（不會有兩份正規式漂移）', L.isValidQuoteId === E.isValidQuoteId);
  [UUID, 'legacy-1', 'a', 'A_b-9', 'x'.repeat(64), '-', '_'].forEach((id) => t('合法 id：' + short(id), L.isValidQuoteId(id) === true));
  [
    '', 'x'.repeat(65), 'a/b', 'a.b', '..', '%2e%2e', 'a b', ' a', 'a ', 'a\n', 'a\r\n', '\na', cp(0xff11) + '23', 'é', 'a%00', 'a\x00', 'a:cost',
    'a#b', 'a?b', 'a\\b', '<script>', cp(0x202e) + 'abc', 'a' + cp(0x2028), null, undefined, 123, [], ['a'], {}, true, Symbol('x'),
  ].forEach((id) => t('非法 id：' + short(typeof id === 'symbol' ? 'Symbol' : id), L.isValidQuoteId(id) === false));
  fast('isValidQuoteId（200 萬字元）', () => L.isValidQuoteId('a'.repeat(2e6)), 50);

  // ── buildQuoteLink ──
  eq('預設正式站連結', L.buildQuoteLink(cfg, UUID), 'https://itts-crm.vercel.app/q/' + UUID);
  eq('填成本加 ?cost=1', L.buildQuoteLink(cfg, UUID, { cost: true }), 'https://itts-crm.vercel.app/q/' + UUID + '?cost=1');
  [true, 1, '1'].forEach((c) => t('cost=' + short(c) + ' 算填成本', L.buildQuoteLink(cfg, UUID, { cost: c }).endsWith('?cost=1') && L.isCostFlag(c)));
  [false, 0, '0', 'true', 'yes', '', null, undefined, ['1'], {}, 2].forEach((c) => {
    t('cost=' + short(c) + ' 不算填成本', L.buildQuoteLink(cfg, UUID, { cost: c }) === 'https://itts-crm.vercel.app/q/' + UUID && !L.isCostFlag(c));
  });
  eq('opts 缺省／null 不 throw', [L.buildQuoteLink(cfg, 'a'), L.buildQuoteLink(cfg, 'a', null)], ['https://itts-crm.vercel.app/q/a', 'https://itts-crm.vercel.app/q/a']);
  eq('自訂網站根（Demo）', L.buildQuoteLink(CF.getMailConfig({ APP_BASE_URL: 'https://demo.example.test' }), UUID), 'https://demo.example.test/q/' + UUID);
  eq('本機 http://localhost 可用', L.buildQuoteLink(CF.getMailConfig({ APP_BASE_URL: 'http://localhost:3000' }), UUID), 'http://localhost:3000/q/' + UUID);
  eq('網站根尾端斜線不會變成雙斜線', L.buildQuoteLink({ appBaseUrl: 'https://x.example.test/' }, 'a'), 'https://x.example.test/q/a');
  eq('網站根主機名稱被正規化成小寫', L.buildQuoteLink({ appBaseUrl: 'HTTPS://Itts-CRM.Vercel.App' }, 'a'), 'https://itts-crm.vercel.app/q/a');
  [
    undefined, null, '', '   ', 123, [], {}, 'javascript:alert(1)', 'http://evil.example', 'ftp://x.test', 'https://a.test"onmouseover=',
    'https://itts-crm.vercel.app/app', 'https://u:p@itts-crm.vercel.app', 'https://itts-crm.vercel.app?x=1', 'https://itts-crm.vercel.app#x',
    '//itts-crm.vercel.app', 'itts-crm.vercel.app', 'https://itts crm.vercel.app', 'https://itts-crm.vercel.app\nX: y', 'x'.repeat(300),
  ].forEach((base) => {
    throwsCode('網站根不合法 → BAD_BASE_URL：' + short(base), () => L.buildQuoteLink({ appBaseUrl: base }, UUID), 'BAD_BASE_URL');
  });
  [undefined, null, 'x', 5, []].forEach((c) => throwsCode('config 不合法 → BAD_BASE_URL：' + short(c), () => L.buildQuoteLink(c, UUID), 'BAD_BASE_URL'));
  throwsCode('config 的 getter 丟例外 → BAD_BASE_URL（不洩漏原始例外）', () => L.buildQuoteLink(Object.defineProperty({}, 'appBaseUrl', { get() { throw new Error('boom'); } }), UUID), 'BAD_BASE_URL');
  ['', 'a/b', '../x', 'x'.repeat(65), null, undefined, 5, {}, 'a\n', '%2e%2e', 'a?cost=1'].forEach((id) => {
    throwsCode('單據 id 不合法 → BAD_QUOTE_ID：' + short(id), () => L.buildQuoteLink(cfg, id), 'BAD_QUOTE_ID');
  });
  t('錯誤是 MailLinkError 的實例', (() => { try { L.buildQuoteLink(cfg, ''); } catch (e) { return e instanceof L.MailLinkError && e.name === 'MailLinkError' && e instanceof Error; } return false; })());
  {
    const u = new URL(L.buildQuoteLink(cfg, UUID, { cost: true }));
    eq('連結結構：origin／path／query 精確', [u.origin, u.pathname, u.search, u.hash, u.username, u.password], ['https://itts-crm.vercel.app', '/q/' + UUID, '?cost=1', '', '', '']);
    const plain = new URL(L.buildQuoteLink(cfg, UUID));
    eq('一般連結沒有 query', plain.search, '');
    t('連結不帶 token／帳號等資訊', !/token|auth|session|password|secret|@/i.test(L.buildQuoteLink(cfg, UUID, { cost: true })));
  }
  eq('deepHash', [L.deepHash(UUID), L.deepHash(UUID, { cost: true }), L.deepHash('a', { cost: 'true' })], ['#quote:' + UUID, '#quote:' + UUID + ':cost', '#quote:a']);
  throwsCode('deepHash 拒絕非法 id', () => L.deepHash('a/b'), 'BAD_QUOTE_ID');
  eq('JUMP_PATH_PREFIX', L.JUMP_PATH_PREFIX, '/q/');

  // ── renderJumpPage：有效 id ──
  const ok1 = L.renderJumpPage({ quoteId: UUID });
  const okCost = L.renderJumpPage({ quoteId: UUID, cost: true });
  eq('有效 id → 200', [ok1.status, okCost.status], [200, 200]);
  t('meta refresh 導向 /index.html#quote:<id>', ok1.body.indexOf('<meta http-equiv="refresh" content="0;url=/index.html#quote:' + UUID + '">') >= 0, ok1.body);
  t('備援連結（沒有 JS 也能點）', ok1.body.indexOf('<a href="/index.html#quote:' + UUID + '">') >= 0);
  t('填成本 → #quote:<id>:cost（meta 與連結都是）', okCost.body.indexOf('url=/index.html#quote:' + UUID + ':cost"') >= 0 && okCost.body.indexOf('<a href="/index.html#quote:' + UUID + ':cost">') >= 0);
  t('沒有 cost 就不會出現 :cost', ok1.body.indexOf(':cost') < 0);
  [true, 1, '1'].forEach((c) => t('跳板頁 cost=' + short(c) + ' 算填成本', L.renderJumpPage({ quoteId: 'a', cost: c }).body.indexOf('#quote:a:cost') >= 0));
  [false, 0, '0', 'true', ['1'], {}, undefined, null].forEach((c) => t('跳板頁 cost=' + short(c) + ' 不算填成本', L.renderJumpPage({ quoteId: 'a', cost: c }).body.indexOf(':cost') < 0));
  {
    const target = /url=(\/index\.html)(#[^"]*)"/.exec(ok1.body);
    t('導向的片段被 deep-link.js 接受（兩邊格式一致）', !!target && DL.isValidHash(target[2]) && target[2] === L.deepHash(UUID), short(target));
  }
  eq('標頭：Cache-Control', H(ok1, 'cache-control'), 'no-store');
  eq('標頭：X-Robots-Tag', H(ok1, 'x-robots-tag'), 'noindex');
  eq('標頭：Referrer-Policy', H(ok1, 'referrer-policy'), 'no-referrer');
  eq('標頭：Content-Type', H(ok1, 'content-type'), 'text/html; charset=utf-8');
  eq('標頭：X-Content-Type-Options', H(ok1, 'x-content-type-options'), 'nosniff');
  {
    const csp = H(ok1, 'content-security-policy') || '';
    t('回應自帶 CSP：default-src 與 frame-ancestors 為 none', csp.indexOf("default-src 'none'") >= 0 && csp.indexOf("frame-ancestors 'none'") >= 0, csp);
    t('CSP 沒有 script-src／萬用字元／unsafe-eval（頁面不需要任何 script）', !/script-src|\*|unsafe-eval/.test(csp), csp);
    t('CSP 允許頁面自己的行內樣式（style-src unsafe-inline）', csp.indexOf("style-src 'unsafe-inline'") >= 0, csp);
    t('CSP 不比專案現行政策寬鬆：沒有任何 https: 或外部來源', !/https?:|data:|blob:/.test(csp), csp);
  }
  eq('404 與 200 的標頭鍵完全相同', Object.keys(L.renderJumpPage({ quoteId: '..' }).headers).sort(), Object.keys(ok1.headers).sort());
  {
    const a = L.renderJumpPage({ quoteId: UUID });
    a.headers['Cache-Control'] = 'public, max-age=999';
    a.headers.Extra = 'x';
    eq('headers 每次都是新物件（改了不影響下一次）', [H(L.renderJumpPage({ quoteId: UUID }), 'cache-control'), H(L.renderJumpPage({ quoteId: UUID }), 'extra')], ['no-store', undefined]);
  }
  {
    // 頁面內容衛生：沒有任何會觸發 CSP 或載入資源的東西
    const forbidden = [/<script/i, /<link/i, /<img/i, /<iframe/i, /<object/i, /<embed/i, /<form/i, /<base/i, /<input/i, /<video/i, /<audio/i, /<svg/i, /<source/i,
      /javascript:/i, /\son[a-z]+\s*=/i, /@import/i, /url\(/i, /https?:/i, /\/\//, /data:/i, /expression\(/i];
    [['有效', ok1.body], ['有效(cost)', okCost.body], ['404', L.renderJumpPage({}).body]].forEach(([name, body]) => {
      const hit = forbidden.filter((re) => re.test(body)).map(String);
      t('頁面衛生（' + name + '）：無 script／外部資源／事件屬性', hit.length === 0, hit.join(' '));
      const count = (re) => (body.match(re) || []).length;
      const tags = ['html', 'head', 'body', 'main', 'p', 'a', 'style', 'title'];
      const unbalanced = tags.filter((tg) => count(new RegExp('<' + tg + '(\\s|>)', 'gi')) !== count(new RegExp('</' + tg + '>', 'gi')));
      t('頁面衛生（' + name + '）：標籤成對', unbalanced.length === 0, unbalanced.join(','));
      t('頁面衛生（' + name + '）：小於 2KB、以 doctype 開頭、宣告 utf-8', body.length < 2048 && /^<!doctype html>/i.test(body) && body.indexOf('<meta charset="utf-8">') >= 0, body.length);
    });
    t('有效頁面只有一個 meta refresh，且目標是同站相對路徑', (ok1.body.match(/http-equiv="refresh"/g) || []).length === 1 && /content="0;url=\/index\.html#quote:/.test(ok1.body));
    t('404 頁面沒有 refresh、沒有 quote 片段', !/refresh|quote:/i.test(L.renderJumpPage({}).body));
  }
  {
    // 與專案現行 CSP（server.js:90-102，2026-10-08 抄錄）相容：頁面唯一的行內資源是 <style>，現行 style-src 允許 'unsafe-inline'；其餘全部不載入資源
    const PROJECT_STYLE_SRC = ["'self'", "'unsafe-inline'"];
    const body = ok1.body;
    const inlineStyles = (body.match(/<style>/g) || []).length;
    t('相容專案 CSP：只有 1 段行內 <style>，且現行 style-src 允許 unsafe-inline', inlineStyles === 1 && PROJECT_STYLE_SRC.indexOf("'unsafe-inline'") >= 0);
    t('相容專案 CSP：沒有 style 屬性以外的行內資源（無 script、無事件屬性、無外部載入）', !/<script|\son[a-z]+=|<link|<img|<iframe/i.test(body));
  }

  {
    // 標頭名稱與值必須能被 Node 的 http（Express 底層）接受，否則 res.set() 會在執行階段丟例外
    const http = require('http');
    const bad = [];
    [ok1, okCost, L.renderJumpPage({})].forEach((r) => {
      Object.keys(r.headers).forEach((k) => {
        try { http.validateHeaderName(k); http.validateHeaderValue(k, r.headers[k]); } catch (e) { bad.push(k + ': ' + e.message); }
      });
    });
    t('所有回應標頭的名稱與值都能被 Node http 接受', bad.length === 0, bad.join('; '));
    t('回應標頭的值都是字串、不含換行', [ok1, okCost].every((r) => Object.keys(r.headers).every((k) => typeof r.headers[k] === 'string' && !/[\r\n]/.test(r.headers[k]))));
  }

  // ── renderJumpPage：無效 id ──
  const BAD_IDS = ['', ' ', 'x'.repeat(65), 'a/b', '../../etc/passwd', '%2e%2e', 'a b', 'a\n', 'a;b', 'a:cost', '<script>alert(1)</script>', '"><img src=x onerror=alert(1)>',
    'javascript:alert(1)', cp(0xff11) + '23', 'é', cp(0x202e) + 'abc', null, undefined, 123, [], ['a'], {}, true, 'a'.repeat(5000)];
  const notFounds = BAD_IDS.map((id) => L.renderJumpPage({ quoteId: id }));
  t('無效 id 一律 404', notFounds.every((r) => r.status === 404), short(notFounds.map((r) => r.status)));
  t('所有無效 id 回傳完全相同的內容（不因輸入而不同）', new Set(notFounds.map((r) => r.body)).size === 1);
  t('所有無效 id 的標頭完全相同', new Set(notFounds.map((r) => JSON.stringify(r.headers))).size === 1);
  t('404 頁面不回顯輸入', BAD_IDS.every((id, i) => typeof id !== 'string' || id.length < 3 || notFounds[i].body.indexOf(id) < 0));
  t('404 也帶 no-store／noindex／no-referrer', H(notFounds[0], 'cache-control') === 'no-store' && H(notFounds[0], 'x-robots-tag') === 'noindex' && H(notFounds[0], 'referrer-policy') === 'no-referrer');
  {
    const a = 'aaaaaaaa-1111-4222-8333-444444444444';
    const b = 'bbbbbbbb-5555-4666-8777-888888888888';
    const ra = L.renderJumpPage({ quoteId: a }).body;
    const rb = L.renderJumpPage({ quoteId: b }).body;
    eq('不同單據 id 的頁面除了 id 之外完全相同（沒有任何逐單資訊）', ra.split(a).join('ID'), rb.split(b).join('ID'));
  }
  {
    const hostile = [undefined, null, 'str', 5, [], Symbol('s'), { get quoteId() { throw new Error('boom'); } },
      { quoteId: UUID, get cost() { throw new Error('boom'); } },
      new Proxy({}, { get() { throw new Error('boom'); } }), { quoteId: { toString() { throw new Error('boom'); } } }];
    hostile.forEach((x, i) => {
      let r;
      let threw = false;
      try { r = L.renderJumpPage(x); } catch (e) { threw = true; }
      t('惡意輸入 #' + i + ' 不 throw 且回合法回應', !threw && r && (r.status === 404 || r.status === 200) && typeof r.body === 'string', threw ? 'threw' : (r && r.status));
    });
    t('不帶參數也不 throw', (() => { try { return L.renderJumpPage().status === 404; } catch (e) { return false; } })());
    fast('renderJumpPage（200 萬字元 id）', () => L.renderJumpPage({ quoteId: 'a'.repeat(2e6) }), 50);
  }
  {
    // 模糊測試：200 與 404 的分界必須等於獨立 oracle；200 頁面的導向目標只由 id 組成；與 deep-link 格式一致
    const rnd = mulberry(20261008);
    const bad = 'abcXYZ019_-';
    const specials = ['/', '.', '%', ':', ' ', '\n', '<', '>', '"', "'", '\\', '?', '#', '&', '=', cp(0xff21), cp(0x2028), 'é', '\x00', '@'];
    const gen = () => {
      const n = Math.floor(rnd() * 70);
      let s = '';
      for (let i = 0; i < n; i++) s += SAFE.charAt(Math.floor(rnd() * SAFE.length));
      if (rnd() < 0.4 && s.length) { const p = Math.floor(rnd() * s.length); s = s.slice(0, p) + specials[Math.floor(rnd() * specials.length)] + s.slice(p + 1); }
      return s || (rnd() < 0.5 ? '' : bad);
    };
    let mismatch = 0;
    let valid = 0;
    let dlMismatch = 0;
    for (let i = 0; i < 20000; i++) {
      const id = gen();
      const r = L.renderJumpPage({ quoteId: id });
      const want = isValidOracle(id);
      if (want) valid++;
      if ((r.status === 200) !== want) mismatch++;
      if (want) {
        const m = /url=([^"]*)"/.exec(r.body);
        if (!m || m[1] !== '/index.html#quote:' + id) mismatch++;
      }
      if (DL.isValidHash('#quote:' + id) !== want) dlMismatch++;
      if (L.isValidQuoteId(id) !== want) mismatch++;
    }
    t('模糊測試 20000 組（其中合法 ' + valid + '）：跳板頁／isValidQuoteId 與獨立 oracle 一致', mismatch === 0 && valid > 3000, mismatch);
    t('模糊測試：deep-link.js 與 link.js 對 id 的判定完全一致', dlMismatch === 0, dlMismatch);
  }
});

// ═════════════════════════════════════════════════════════════════════════
section('8 deep-link（_client/deep-link.js）', () => {
  const vm = require('vm');
  const SRC = fs.readFileSync(path.join(ROOT, '_client', 'deep-link.js'), 'utf8');
  const KEY = 'itts.deepLink';
  const UUID = '0f8fad5b-d9cb-469f-a165-70867728950e';
  // 在 vm 沙箱裡執行原始碼（等同瀏覽器載入 <script>）：window＝沙箱本身，附假的 sessionStorage／location
  function mkEnv(opt) {
    const o = opt || {};
    const store = new Map();
    const calls = { get: 0, set: 0, remove: 0 };
    const storage = {
      getItem(k) { calls.get++; if (o.throwGet) throw new Error('get failed'); return store.has(k) ? store.get(k) : null; },
      setItem(k, v) { calls.set++; if (o.throwSet) throw new Error('QuotaExceededError'); store.set(k, String(v)); },
      removeItem(k) { calls.remove++; if (o.throwRemove) throw new Error('remove failed'); store.delete(k); },
    };
    // window 用一個「普通物件」而不是沙箱全域本身：Node 的 vm 會把全域上 getter 丟出的例外吞掉並回傳 undefined，
    // 那樣就測不到「存取 sessionStorage 本身會丟 SecurityError」這種真實情況（變異 M134 曾因此倖存）
    const win = {};
    if (o.locationThrows) Object.defineProperty(win, 'location', { get() { throw new Error('no location'); } });
    else win.location = { hash: o.hash || '' };
    if (o.storage === 'getterThrows') Object.defineProperty(win, 'sessionStorage', { get() { throw new Error('SecurityError'); } });
    else if (o.storage !== 'absent') win.sessionStorage = storage;
    const sandbox = { window: win };
    vm.createContext(sandbox);
    vm.runInContext(SRC, sandbox, { filename: 'deep-link.js' });
    return { api: win.ITTSDeepLink, store, calls, sandbox, win };
  }

  const env0 = mkEnv();
  t('載入後掛在 window.ITTSDeepLink', !!env0.api && typeof env0.api === 'object');
  eq('公開的函式', ['remember', 'consume', 'hashForRedirect', 'isValidHash'].map((k) => typeof env0.api[k]), Array(4).fill('function'));
  eq('STORAGE_KEY', env0.api.STORAGE_KEY, KEY);
  t('只在 window 上多一個 ITTSDeepLink，沒有洩漏其他全域', Object.keys(env0.win).sort().join() === 'ITTSDeepLink,location,sessionStorage' && Object.keys(env0.sandbox).join() === 'window', Object.keys(env0.win).join() + ' | ' + Object.keys(env0.sandbox).join());

  // ── 合法格式 ──
  const VALID = ['#quote:a', '#quote:' + UUID, '#quote:' + UUID + ':cost', '#quote:legacy-1', '#quote:A_b-9', '#quote:' + 'x'.repeat(64), '#quote:' + 'x'.repeat(64) + ':cost', '#quote:-', '#quote:_:cost'];
  VALID.forEach((h) => {
    const e = mkEnv();
    t('remember 接受 ' + short(h.length > 50 ? h.slice(0, 50) + '…' : h), e.api.remember(h) === true && e.store.get(KEY) === h && e.api.isValidHash(h) === true);
  });
  {
    const e = mkEnv();
    eq('remember 接受帶 .hash 的物件（location）', [e.api.remember({ hash: '#quote:abc', pathname: '/login.html' }), e.store.get(KEY)], [true, '#quote:abc']);
  }
  {
    const e = mkEnv({ hash: '#quote:fromloc:cost' });
    eq('省略參數 → 讀 window.location.hash', [e.api.remember(), e.store.get(KEY)], [true, '#quote:fromloc:cost']);
  }

  // ── 不合法格式：一律不存、不 throw ──
  const INVALID = [
    '', '#', '#quote', '#quote:', '#quote::cost', '#quote:abc:', '#quote:abc:cost2', '#quote:abc:COST', '#quote:abc:cost:cost', '#quote:abc:cost ',
    '#QUOTE:abc', 'quote:abc', '##quote:abc', ' #quote:abc', '#quote:abc ', '#quote:abc\n', '#quote:abc\r\n', '#quote:\nabc', '#quote:abc\t',
    '#quote:../x', '#quote:a/b', '#quote:a.b', '#quote:a%2Fb', '#quote:a b', '#quote:a?x=1', '#quote:a#b', '#quote:a\\b', '#quote:a;b', '#quote:a=b',
    '#quote:' + 'x'.repeat(65), '#quote:' + 'x'.repeat(65) + ':cost', '#quote:' + 'x'.repeat(5000), '#quote:' + cp(0xff11) + '23', '#quote:é',
    '#quote:abc' + cp(0x2028), '#quote:' + cp(0x202e) + 'abc', '#quote:abc\x00', '#quote:<script>', '#quote:abc"onload=',
    'javascript:alert(1)', 'https://evil.example/', '//evil.example', '/index.html#quote:abc', '#foo', '#/quote:abc', '#quote:abc#quote:def',
    null, undefined, 123, true, {}, [], ['#quote:abc'], Symbol('x'),
  ];
  INVALID.forEach((h) => {
    const e = mkEnv();
    let r;
    let threw = false;
    try { r = e.api.remember(h); } catch (x) { threw = true; }
    const name = typeof h === 'symbol' ? 'Symbol' : (typeof h === 'string' && h.length > 40 ? h.slice(0, 40) + '…(' + h.length + ')' : short(h));
    t('remember 拒絕 ' + name, !threw && r === false && e.calls.set === 0 && e.store.size === 0 && e.api.isValidHash(h) === false, 'threw=' + threw + ' r=' + r + ' set=' + e.calls.set);
  });
  {
    const e = mkEnv();
    e.api.remember('#quote:first');
    INVALID.forEach((h) => { try { e.api.remember(h); } catch (x) { /* 另有測試 */ } });
    eq('無效輸入不會清掉先前記住的值', e.store.get(KEY), '#quote:first');
    eq('後記住的覆蓋前一個（最新優先）', [e.api.remember('#quote:second'), e.store.get(KEY)], [true, '#quote:second']);
  }

  // ── consume：取一次即清 ──
  {
    const e = mkEnv();
    e.api.remember('#quote:' + UUID + ':cost');
    eq('consume 取出', e.api.consume(), '#quote:' + UUID + ':cost');
    eq('consume 第二次為空（取一次即清）', e.api.consume(), '');
    eq('儲存已被清除', e.store.has(KEY), false);
    eq('沒有東西時 consume 為空且不 throw', mkEnv().api.consume(), '');
  }
  {
    const e = mkEnv();
    e.api.remember('#quote:abc');
    eq('hashForRedirect 取出可直接接在網址後面', '/' + e.api.hashForRedirect(), '/#quote:abc');
    eq('hashForRedirect 同樣取一次即清', [e.api.hashForRedirect(), e.api.consume(), e.store.has(KEY)], ['', '', false]);
    eq('沒有記住時 hashForRedirect 為空字串（網址後面不會多東西）', '/' + mkEnv().api.hashForRedirect(), '/');
  }
  ['javascript:alert(1)', 'https://evil.example/', '#quote:../../x', '#quote:abc\n', '#quote:' + 'x'.repeat(65), '<script>', '', '#quote:a b'].forEach((tampered) => {
    const e = mkEnv();
    e.store.set(KEY, tampered);             // 模擬 sessionStorage 被竄改（XSS、其他程式碼）
    const r = e.api.consume();
    t('consume 對被竄改的儲存內容回空並清掉：' + short(tampered.length > 40 ? tampered.slice(0, 40) + '…' : tampered), r === '' && !e.store.has(KEY), short(r));
    e.store.set(KEY, tampered);
    t('hashForRedirect 同樣不放行：' + short(tampered.length > 40 ? tampered.slice(0, 40) + '…' : tampered), e.api.hashForRedirect() === '' && !e.store.has(KEY));
  });
  {
    const e = mkEnv();
    e.store.set('other.key', '#quote:zzz');
    e.api.remember('#quote:abc');
    e.api.consume();
    eq('只動自己的 key', e.store.get('other.key'), '#quote:zzz');
  }

  // ── sessionStorage／location 出狀況 ──
  {
    const e = mkEnv({ storage: 'getterThrows' });
    let res;
    noThrow('存取 window.sessionStorage 本身就丟例外（隱私模式／封鎖 cookie）：不 throw', () => { res = [e.api.remember('#quote:abc'), e.api.consume(), e.api.hashForRedirect()]; });
    eq('→ remember=false，consume／hashForRedirect 為空', res, [false, '', '']);
  }
  {
    const e = mkEnv({ storage: 'absent' });
    eq('沒有 sessionStorage（undefined）', [e.api.remember('#quote:abc'), e.api.consume()], [false, '']);
  }
  {
    const e = mkEnv({ throwSet: true });
    let r;
    noThrow('setItem 丟 QuotaExceededError：不 throw', () => { r = e.api.remember('#quote:abc'); });
    eq('→ remember=false，沒有殘留', [r, e.store.size, e.api.consume()], [false, 0, '']);
  }
  {
    const e = mkEnv({ throwGet: true });
    e.store.set(KEY, '#quote:abc');
    let r;
    noThrow('getItem 丟例外：不 throw', () => { r = e.api.consume(); });
    eq('→ consume 為空', r, '');
  }
  {
    const e = mkEnv({ throwRemove: true });
    e.api.remember('#quote:abc');
    let r;
    noThrow('removeItem 丟例外：不 throw', () => { r = e.api.consume(); });
    eq('→ 仍回傳已記住的值（盡力而為；清不掉是 sessionStorage 壞掉的情況）', r, '#quote:abc');
  }
  {
    const e = mkEnv({ locationThrows: true });
    let r;
    noThrow('window.location 存取丟例外：remember() 不 throw', () => { r = e.api.remember(); });
    eq('→ false', r, false);
    eq('字串參數不受 location 影響', e.api.remember('#quote:abc'), true);
  }

  // ── 登出殘留（FIX-4）：rememberForLogin／settleOnLoginPage／clear ──
  {
    const ARMED = 'itts.deepLink.armed';
    const LINK = '#quote:' + UUID;
    const e0 = mkEnv();
    eq('公開的函式（含 rememberForLogin／settleOnLoginPage／clear）', ['rememberForLogin', 'settleOnLoginPage', 'clear'].map((k) => typeof e0.api[k]), Array(3).fill('function'));
    eq('ARMED_KEY', e0.api.ARMED_KEY, ARMED);
    // remember：只存、不上膛
    { const e = mkEnv(); e.api.remember(LINK); t('remember 不上膛', e.store.get(KEY) === LINK && !e.store.has(ARMED)); }
    // rememberForLogin：存＋上膛；格式不符一律不存也不上膛
    { const e = mkEnv(); eq('rememberForLogin 合法 → 存＋上膛', [e.api.rememberForLogin(LINK), e.store.get(KEY), e.store.get(ARMED)], [true, LINK, '1']); }
    { const e = mkEnv(); eq('rememberForLogin 不合格 → false、不存、不上膛', [e.api.rememberForLogin('#evil:1'), e.store.has(KEY), e.store.has(ARMED)], [false, false, false]); }
    { const e = mkEnv({ hash: LINK + ':cost' }); eq('rememberForLogin 省略參數 → 讀 location.hash', [e.api.rememberForLogin(), e.store.get(KEY)], [true, LINK + ':cost']); }
    // clear：兩個 key 都清
    { const e = mkEnv(); e.api.rememberForLogin(LINK); eq('clear 清掉暫存與上膛旗標', [e.api.clear(), e.store.has(KEY), e.store.has(ARMED)], [true, false, false]); }
    { const e = mkEnv(); eq('clear 在沒有東西時回 true、不 throw', e.api.clear(), true); }
    { const e = mkEnv({ storage: 'absent' }); eq('clear 沒有 sessionStorage → false', e.api.clear(), false); }
    { const e = mkEnv({ storage: 'getterThrows' }); eq('clear 存取 sessionStorage 會丟例外 → false、不 throw', e.api.clear(), false); }
    { const e = mkEnv({ throwRemove: true }); e.api.remember(LINK); eq('clear 清不掉 → false、不 throw', e.api.clear(), false); }
    // consume：連上膛旗標一起清
    { const e = mkEnv(); e.api.rememberForLogin(LINK); eq('consume 取出後暫存與上膛旗標都清掉', [e.api.consume(), e.store.has(KEY), e.store.has(ARMED)], [LINK, false, false]); }
    { const e = mkEnv(); e.api.rememberForLogin(LINK); eq('hashForRedirect（登入成功導向）同樣取一次即清、含上膛旗標', [e.api.hashForRedirect(), e.api.hashForRedirect(), e.store.has(ARMED)], [LINK, '', false]); }
    // settleOnLoginPage：登入頁載入時決定要不要沿用暫存
    { const e = mkEnv({ hash: LINK }); eq('網址有合法片段 → 記住（hash），並把上膛旗標取走', [e.api.settleOnLoginPage(e.win.location.hash), e.store.get(KEY), e.store.has(ARMED)], ['hash', LINK, false]); }
    { const e = mkEnv(); e.api.remember('#quote:old'); eq('網址有合法片段 → 覆蓋舊暫存', [e.api.settleOnLoginPage(LINK), e.store.get(KEY)], ['hash', LINK]); }
    { const e = mkEnv(); e.api.rememberForLogin(LINK); eq('沒有片段、但上膛（app.js 為了登入而導過來）→ 保留暫存（kept），上膛旗標被取走（一次性）', [e.api.settleOnLoginPage(''), e.store.get(KEY), e.store.has(ARMED)], ['kept', LINK, false]); }
    { const e = mkEnv(); e.api.rememberForLogin(LINK); e.api.settleOnLoginPage(''); eq('第二次載入登入頁（沒有片段、旗標已被取走）→ 清掉暫存', [e.api.settleOnLoginPage(''), e.store.has(KEY)], ['cleared', false]); }
    { const e = mkEnv(); e.api.remember(LINK); eq('沒有片段、沒上膛（使用者主動打開登入頁、或強制改密碼畫面登出後）→ 清掉舊暫存', [e.api.settleOnLoginPage(''), e.store.has(KEY)], ['cleared', false]); }
    { const e = mkEnv(); eq('什麼都沒有 → cleared（不 throw）', e.api.settleOnLoginPage(''), 'cleared'); }
    { const e = mkEnv(); e.api.rememberForLogin(LINK); e.store.set(KEY, '#evil:1'); eq('上膛但暫存內容被竄改成不合格 → 清掉', [e.api.settleOnLoginPage(''), e.store.has(KEY)], ['cleared', false]); }
    { const e = mkEnv(); eq('網址片段不合格 → 當成沒有片段（不存）', [e.api.settleOnLoginPage('#evil:1'), e.store.has(KEY)], ['cleared', false]); }
    { const e = mkEnv({ storage: 'absent' }); eq('沒有 sessionStorage → none、不 throw', e.api.settleOnLoginPage(LINK), 'none'); }
    { const e = mkEnv({ storage: 'getterThrows' }); eq('存取 sessionStorage 丟例外 → none、不 throw', e.api.settleOnLoginPage(''), 'none'); }
    { const e = mkEnv({ throwGet: true, throwRemove: true, throwSet: true }); let ok = true; try { e.api.settleOnLoginPage(''); e.api.settleOnLoginPage(LINK); e.api.rememberForLogin(LINK); e.api.clear(); e.api.consume(); } catch (x) { ok = false; } t('storage 全部丟例外：settle／rememberForLogin／clear／consume 都不 throw', ok); }
    // 完整情境：A 在強制改密碼畫面登出 → B 登入
    {
      const e = mkEnv();
      e.api.remember(LINK);                       // app.js 強制改密碼分支：存、不上膛
      e.api.clear();                              // 登出：clear
      eq('情境：強制改密碼畫面登出後，登入頁載入 → 沒有東西可沿用；B 登入的導向是空片段', [e.api.settleOnLoginPage(''), e.api.hashForRedirect()], ['cleared', '']);
      const e2 = mkEnv();
      e2.api.remember(LINK);                      // 就算登出時沒清（例如舊版 app.js 快取）…
      eq('情境：登出沒清，登入頁載入也會清掉（第二層防護）→ B 登入導向空片段', [e2.api.settleOnLoginPage(''), e2.api.hashForRedirect()], ['cleared', '']);
      const e3 = mkEnv();
      e3.api.rememberForLogin(LINK);              // 401 分支：存＋上膛 → 登入頁載入保留 → 登入成功帶回
      eq('情境：401 分支 → 登入頁保留 → 登入成功帶回該單，之後即清', [e3.api.settleOnLoginPage(''), e3.api.hashForRedirect(), e3.api.hashForRedirect()], ['kept', LINK, '']);
    }
  }

  // ── Node 載入（module.exports）──
  {
    const API = load('_client/deep-link.js');
    t('Node require 取得同樣的 API', ['remember', 'rememberForLogin', 'settleOnLoginPage', 'clear', 'consume', 'hashForRedirect', 'isValidHash'].every((k) => typeof API[k] === 'function') && API.STORAGE_KEY === KEY);
    t('Node 載入不會污染全域', typeof globalThis.ITTSDeepLink === 'undefined' && typeof globalThis.window === 'undefined');
    eq('Node 沒有 sessionStorage：remember=false、consume 為空', [API.remember('#quote:abc'), API.consume()], [false, '']);
    eq('Node 下 isValidHash 可用', [API.isValidHash('#quote:abc'), API.isValidHash('#quote:a/b')], [true, false]);
  }

  // ── 模糊測試：與獨立 oracle 逐字比對 ──
  {
    const oracle = (s) => {
      if (typeof s !== 'string' || s.indexOf('#quote:') !== 0) return false;
      const parts = s.slice(7).split(':');
      if (parts.length > 2) return false;
      if (parts.length === 2 && parts[1] !== 'cost') return false;
      const id = parts[0];
      if (id.length < 1 || id.length > 64) return false;
      for (const ch of id) if (!/[A-Za-z0-9_-]/.test(ch)) return false;
      return true;
    };
    let a = 20261009;
    const rnd = () => { a |= 0; a = (a + 0x6d2b79f5) | 0; let x = Math.imul(a ^ (a >>> 15), 1 | a); x = (x + Math.imul(x ^ (x >>> 7), 61 | x)) ^ x; return ((x ^ (x >>> 14)) >>> 0) / 4294967296; };
    const SAFE = 'abcXYZ019_-';
    const bits = ['#', 'quote', ':', 'cost', 'a', 'Z', '0', '_', '-', '/', '.', '?', '%', '\n', ' ', '<', '>', '"', "'", '\\', cp(0xff21), cp(0x2028), 'é', '\x00', '@', ';'];
    const gen = () => {
      if (rnd() < 0.5) {
        const n = 1 + Math.floor(rnd() * 70);
        let s = '#quote:';
        for (let i = 0; i < n; i++) s += SAFE.charAt(Math.floor(rnd() * SAFE.length));
        if (rnd() < 0.4) s += ':cost';
        if (rnd() < 0.4) { const p = Math.floor(rnd() * s.length); s = s.slice(0, p) + bits[Math.floor(rnd() * bits.length)] + s.slice(p + 1); }
        return s;
      }
      const n = Math.floor(rnd() * 12);
      let s = '';
      for (let i = 0; i < n; i++) s += bits[Math.floor(rnd() * bits.length)];
      return s;
    };
    const e = mkEnv();
    let mism = 0;
    let valid = 0;
    for (let i = 0; i < 20000; i++) {
      const s = gen();
      const want = oracle(s);
      if (want) valid++;
      const got = e.api.isValidHash(s);
      const stored = e.api.remember(s) === true;
      if (got !== want || stored !== want) mism++;
      else if (want) { if (e.api.consume() !== s) mism++; }
      else if (e.api.consume() !== '' || e.store.has(KEY)) mism++;
    }
    t('模糊測試 20000 組（合法 ' + valid + '）：isValidHash／remember／consume 與獨立 oracle 一致', mism === 0 && valid > 2000, mism);
  }
  fast('isValidHash（200 萬字元）', () => env0.api.isValidHash('#quote:' + 'a'.repeat(2e6)), 50);
  fast('remember（200 萬字元）', () => env0.api.remember('#quote:' + 'a'.repeat(2e6)), 50);
});

// ▲▲▲ 追加位置結束 ▲▲▲
// ═════════════════════════════════════════════════════════════════════════

finish();
