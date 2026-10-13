#!/usr/bin/env node
'use strict';
/**
 * scripts/verify-audit-pg.js 的防護邏輯檢查（不連任何資料庫：注入假的 pg 模組）。用法：node scripts/check-verify-audit-pg-guards.js
 * 涵蓋：沒設環境變數中止、與 DATABASE_URL 相同中止、含正式站特徵字串中止、非空資料庫中止（含邊界）、確認步驟、
 *       錯誤訊息／輸出不洩漏連線字串與密碼與主機、唯讀檢查階段完全沒有寫入語句、腳本本身沒有寫檔與根目錄資料檔參照。
 * 這個檢查使用的連線字串全是假的（.invalid 網域）。
 */
const fs = require('fs');
const path = require('path');
const V = require('./verify-audit-pg.js');

let pass = 0, fail = 0;
const ok = (c, n, d) => { if (c) pass++; else { fail++; console.log('FAIL', n, d === undefined ? '' : d); } };

const PW = 'S3cretPw-for-test';
const HOST = 'db.fakeproject0123.supabase.invalid';
const URLX = `postgresql://postgres:${PW}@${HOST}:5432/postgres`;
const WRITE = /\b(INSERT|CREATE|UPDATE|DELETE|DROP|TRUNCATE|ALTER|GRANT|REVOKE)\b/i;

function fakePg(cfg) {
  cfg = cfg || {};
  const st = { pools: 0, queries: [], ended: 0 };
  class Pool {
    constructor() { st.pools++; }
    on() {}
    async query(sql) { st.queries.push(sql); return respond(sql); }
    async connect() {
      if (cfg.connectError) throw new Error(cfg.connectError);
      return { query: async (sql) => { st.queries.push(sql); return respond(sql); }, release() {} };
    }
    async end() { st.ended++; }
  }
  function respond(sql) {
    if (/to_regclass/.test(sql)) return { rows: [{ ok: cfg.noAppData ? false : true }] };
    if (/FROM app_data WHERE id = 'main'/.test(sql)) return { rows: cfg.main === null ? [] : [Object.assign({ contacts: 0, opps: 0, users: 0 }, cfg.main || {})] };
    return { rows: [] };
  }
  return { mod: { Pool }, st };
}
async function go(env, argv, cfg, extra) {
  const f = fakePg(cfg);
  const lines = [];
  const before = [console.log, console.error, console.warn];
  const r = await V.main(Object.assign({ env, argv: argv || [], out: (l) => lines.push(l), pg: f.mod, stopAfterGuards: true }, extra || {}));
  ok(before[0] === console.log && before[1] === console.error && before[2] === console.warn, 'console 在 main 結束後已還原');
  return { r, lines, st: f.st, text: lines.join('\n') };
}
const allText = [];
const note = (x) => { allText.push(x.text); return x; };

(async () => {
  let x;
  // 1 沒設環境變數
  x = note(await go({}, []));
  ok(x.r.code === 2 && /AUDIT_VERIFY_DATABASE_URL/.test(x.text) && x.st.pools === 0, '缺環境變數：中止、完全沒建連線', x.text);
  // 2 不是 postgres URL
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: 'http://example.invalid/x' }, []));
  ok(x.r.code === 2 && x.st.pools === 0, '非 postgres:// → 中止');
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: 'not a url' }, []));
  ok(x.r.code === 2 && x.st.pools === 0, '亂字串 → 中止');
  // 3 與 DATABASE_URL 相同
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX, DATABASE_URL: URLX }, ['--yes-this-is-test-db']));
  ok(x.r.code === 2 && /DATABASE_URL/.test(x.text) && x.st.pools === 0, '與 DATABASE_URL 完全相同 → 中止、沒連線', x.text);
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX, DATABASE_URL: `postgresql://postgres:otherpw@${HOST}:5432/postgres` }, ['--yes-this-is-test-db']));
  ok(x.r.code === 2 && x.st.pools === 0, '只有密碼不同、其餘（主機／帳號／資料庫）相同 → 仍視為同一個資料庫而中止');
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX, DATABASE_URL: 'postgresql://u:p@other.invalid:5432/other' }, ['--yes-this-is-test-db'], { main: null }));
  ok(x.r.guardsPassed === true, '不同資料庫不會被 DATABASE_URL 規則擋下');
  // 4 正式站特徵字串
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX, AUDIT_VERIFY_FORBIDDEN_SUBSTR: 'FAKEPROJECT0123' }, ['--yes-this-is-test-db']));
  ok(x.r.code === 2 && /FORBIDDEN/.test(x.text) && x.st.pools === 0, '含禁止字串（不分大小寫）→ 中止、沒連線', x.text);
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX, AUDIT_VERIFY_FORBIDDEN_SUBSTR: 'zzz-not-here' }, ['--yes-this-is-test-db'], { main: null }));
  ok(x.r.guardsPassed === true, '不含禁止字串 → 通過');
  // 5 非空資料庫
  for (const [label, main] of [['聯絡人 6', { contacts: 6 }], ['商機 6', { opps: 6 }], ['帳號 4', { users: 4 }], ['全都超過', { contacts: 900, opps: 300, users: 40 }]]) {
    x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX }, ['--yes-this-is-test-db'], { main }));
    ok(x.r.code === 2 && x.text.includes('這看起來不是空的測試資料庫') && !x.r.guardsPassed, `非空資料庫（${label}）→ 中止並顯示「這看起來不是空的測試資料庫」`, x.text);
    ok(x.st.queries.every((q) => !WRITE.test(q)), `非空資料庫（${label}）：只送出唯讀語句`, x.st.queries.filter((q) => WRITE.test(q)));
    ok(x.st.queries.some((q) => /BEGIN READ ONLY/.test(q)), '檢查在 READ ONLY 交易內進行');
    ok(x.st.ended >= 1, '中止時連線池有關閉');
  }
  // 邊界：剛好 5／5／3 視為空
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX }, ['--yes-this-is-test-db'], { main: { contacts: 5, opps: 5, users: 3 } }));
  ok(x.r.guardsPassed === true, '邊界 5／5／3 仍算測試資料庫');
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX }, ['--yes-this-is-test-db'], { main: null }));
  ok(x.r.guardsPassed === true, 'app_data 沒有 main 列 → 通過');
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX }, ['--yes-this-is-test-db'], { noAppData: true }));
  ok(x.r.guardsPassed === true, '完全沒有 app_data 表（全新資料庫）→ 通過');
  ok(x.st.queries.every((q) => !WRITE.test(q)), '通過防護的當下也還沒有任何寫入語句', x.st.queries.filter((q) => WRITE.test(q)));
  // 6 確認步驟
  let asked = 0;
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX }, [], { main: null }, { ask: async (p) => { asked++; return 'wrongpre'; } }));
  ok(asked === 1 && x.r.code === 2 && /前綴不符/.test(x.text) && !x.r.guardsPassed, '輸入錯誤的主機前綴 → 取消');
  ok(!x.text.includes(HOST.slice(0, 8)) , '提示文字沒有把主機前綴印出來', x.text);
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX }, [], { main: null }, { ask: async () => { asked++; return HOST.slice(0, 8); } }));
  ok(asked === 2 && x.r.guardsPassed === true, '輸入正確的主機前 8 字元 → 通過');
  asked = 0;
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX }, ['--yes-this-is-test-db'], { main: null }, { ask: async () => { asked++; return ''; } }));
  ok(asked === 0 && x.r.guardsPassed === true, '--yes-this-is-test-db 不再詢問');
  // 7 連線失敗：錯誤訊息遮蔽
  x = note(await go({ AUDIT_VERIFY_DATABASE_URL: URLX }, ['--yes-this-is-test-db'], { connectError: `connect failed for ${URLX} (password ${PW}) host ${HOST}` }));
  ok(x.r.code === 3, '連線失敗 → 結束代碼 3');
  ok(![URLX, PW, HOST, 'fakeproject0123'].some((s) => x.text.includes(s)), '連線錯誤訊息不含連線字串／密碼／主機', x.text);
  ok(/已遮蔽/.test(x.text), '錯誤訊息有遮蔽標記');
  // redactor 本身
  const red = V.makeRedactor(URLX);
  const t = red(`x ${URLX} y ${PW} z ${HOST} postgres://a:b@c/d`);
  ok(![URLX, PW, HOST, 'a:b@c'].some((s) => t.includes(s)), 'makeRedactor：整串／密碼／主機／任何 postgres:// 都被遮蔽', t);
  ok(red('plain text 沒有祕密') === 'plain text 沒有祕密', 'makeRedactor：一般文字不動');
  const red2 = V.makeRedactor(`postgresql://u:${encodeURIComponent('p@ss:w/rd')}@h.invalid/db`);
  ok(!red2('bad p@ss:w/rd here').includes('p@ss:w/rd'), 'makeRedactor：URL 編碼過的密碼解碼後也遮蔽');
  // 所有輸出總掃描
  const blob = allText.join('\n');
  ok(![URLX, PW, HOST].some((s) => blob.includes(s)), '所有測試情境的輸出都不含連線字串／密碼／主機');
  // 8 腳本本身的靜態檢查
  const src = fs.readFileSync(path.join(__dirname, 'verify-audit-pg.js'), 'utf8');
  const code = src.replace(/^#!.*\n/, '').replace(/\/\*\*[\s\S]*?\*\//, '');   // 去掉檔頭說明
  ok(!/writeFileSync|appendFileSync|createWriteStream|writeFile\(/.test(code), '腳本不寫任何檔案（連線字串不可能被寫進檔案）');
  ok(!/data\.json|auth\.json|audit\.log\.json/.test(code), '腳本程式碼不參照專案根目錄的 data／auth／audit 檔');
  ok(!/console\.(log|error|warn)\([^)]*\burl\b/.test(code), '沒有直接把 url 印出來的 console 呼叫');
  ok(/if \(!url\) return abort/.test(code) && code.indexOf('abort(2') < code.indexOf('new pgMod.Pool'), '所有字串防護都在建立連線池之前');
  ok(code.indexOf('BEGIN READ ONLY') < code.indexOf('CREATE TABLE') || !/CREATE TABLE/.test(code), '空庫檢查在任何寫入語句之前');
  console.log(`verify-audit-pg 防護檢查：PASS ${pass} / FAIL ${fail}`);
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error('例外', e); process.exit(2); });
