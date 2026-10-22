// 檢查 F5 還原區段（app.js 解析期的 restoreLastSection → showSection）不會呼叫「後載 script」才定義的函式。
// 靜態：解析 index.html script 順序與 showSection 內被呼叫的 loader，
//       凡定義在 app.js 之後載入的檔案者，必須包在 _whenScriptsReady(...) 內。
// 行為：用 vm 模擬 readyState=loading，確認 loader 延後到 DOMContentLoaded 才呼叫。
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const dir = path.join(__dirname, '..', '_client');
const html = fs.readFileSync(path.join(dir, 'index.html'), 'utf8');
const app = fs.readFileSync(path.join(dir, 'app.js'), 'utf8');
let fail = 0;
const ok = (c, m) => { if (c) console.log('PASS', m); else { fail++; console.log('FAIL', m); } };

const scripts = [...html.matchAll(/<script\s+src="([^"]+\.js)"/g)].map(m => m[1]).filter(s => !/^https?:/.test(s));
const appIdx = scripts.indexOf('app.js');
ok(appIdx >= 0, 'index.html 含 app.js');
const late = scripts.slice(appIdx + 1);

const m = app.match(/function showSection\(section\) \{[\s\S]*?\n\}\n/);
ok(!!m, '找到 showSection');
const body = m ? m[0] : '';
const lines = body.split(/\r?\n/).filter(l => /^\s*if \(section === /.test(l));
// 後載檔定義的頂層函式／window 物件
const lateDefs = new Set();
for (const f of late) {
  const src = fs.readFileSync(path.join(dir, f), 'utf8');
  for (const d of src.matchAll(/^(?:async\s+)?function\s+([A-Za-z0-9_$]+)/gm)) lateDefs.add(d[1]);
  if (/window\.RoleHomes\s*=/.test(src)) lateDefs.add('RoleHomes');
}
let checked = 0;
for (const l of lines) {
  const calls = [...l.matchAll(/\b([A-Za-z_$][\w$]*)\(/g)].map(x => x[1]).filter(n => lateDefs.has(n));
  const usesRoleHomes = /RoleHomes/.test(l) && lateDefs.has('RoleHomes');
  if (calls.length || usesRoleHomes) {
    checked++;
    ok(/_whenScriptsReady\(/.test(l), '後載 loader 已守衛：' + l.trim().slice(0, 90));
  }
}
ok(checked >= 3, '至少涵蓋 quotations／marketingHome／consultHome（實際 ' + checked + '）');

// 行為模擬
function run(readyState) {
  const calls = []; const listeners = [];
  const ctx = {
    currentSection: null,
    document: { readyState, addEventListener: (n, f) => listeners.push(f) },
  };
  const helper = app.match(/function _whenScriptsReady\(section, fn\) \{[\s\S]*?\n\}\n/)[0];
  vm.createContext(ctx);
  vm.runInContext(helper + '\nthis._w=_whenScriptsReady;', ctx);
  ctx.currentSection = 'quotations';
  ctx._w('quotations', () => calls.push('x'));
  const before = calls.length;
  listeners.forEach(f => f());
  return { before, after: calls.length, n: listeners.length, ctx, listeners };
}
let r = run('loading');
ok(r.before === 0 && r.after === 1 && r.n === 1, 'loading：先不呼叫，DOMContentLoaded 後呼叫一次');
r = run('complete');
ok(r.before === 1 && r.n === 0, 'complete：立即呼叫');
{ // 期間切到別區段則放棄
  const calls = []; const ls = [];
  const ctx = { currentSection: 'quotations', document: { readyState: 'loading', addEventListener: (n, f) => ls.push(f) } };
  vm.createContext(ctx);
  vm.runInContext(app.match(/function _whenScriptsReady\(section, fn\) \{[\s\S]*?\n\}\n/)[0] + '\nthis._w=_whenScriptsReady;', ctx);
  ctx._w('quotations', () => calls.push(1));
  ctx.currentSection = 'contacts';
  ls.forEach(f => f());
  ok(calls.length === 0, '已切到別區段則不呼叫');
}
console.log(fail ? `\n${fail} 項失敗` : '\n全部通過');
process.exit(fail ? 1 : 0);
