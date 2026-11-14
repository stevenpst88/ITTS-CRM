#!/usr/bin/env node
/**
 * verify-write-race-pg.js 的安全護欄與整體流程檢查（用假 PG 注入，不連任何遠端資料庫）。用法：node scripts/check-verify-write-race-pg-guards.js
 *   - 各種不合格的輸入必須在「任何寫入之前」中止（假 PG 內沒有 schema、沒有任何 CREATE 類敘述）
 *   - 連線字串／密碼／主機不得出現在任何輸出
 *   - 假 PG 上整套驗證 0 失敗、結束後臨時 schema 已刪除；中途失敗也要清理
 * 注意：假 PG 只驗證腳本邏輯，不等於真 PG 行為。
 */
const { FakeDb } = require('./_fake-pg');
const V = require('./verify-write-race-pg');

const results = [];
const ok = (n, c, e) => results.push([n, !!c, e === undefined ? '' : String(e)]);
const URL_OK = 'postgres://postgres.abcdefgh1234:S3cretPw%21xyz@aws-0-test.pooler.example.com:6543/postgres';
const mkPg = (fake) => ({ Pool: function (o) { return fake.pool({ max: o.max, connectionTimeoutMillis: o.connectionTimeoutMillis }); } });

async function attempt(env, argv, extra) {
  const fake = (extra && extra.fake) || new FakeDb({ latency: [0, 1] });
  const lines = [];
  const r = await V.run(Object.assign({ env, argv, out: (l) => lines.push(l), pg: mkPg(fake), scale: 0.25 }, extra && extra.opts));
  return { r, lines, fake, text: lines.join('\n') };
}
const noWrites = (a) => a.fake.schemas.size === 0 && !a.fake.log.some((q) => /^(CREATE|DROP|INSERT|UPDATE|DELETE)/.test(q));

(async () => {
  const good = ['--expect-ref=abcdefgh1234', '--yes-this-is-test-db'];
  let a = await attempt({}, good);
  ok('沒有連線字串環境變數 → 中止(2)、無寫入', a.r.code === 2 && noWrites(a));
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: 'mysql://x' }, good);
  ok('非 postgres:// → 中止(2)', a.r.code === 2 && noWrites(a));
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: URL_OK, DATABASE_URL: URL_OK }, good);
  ok('與 DATABASE_URL 相同 → 中止(2)', a.r.code === 2 && noWrites(a));
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: URL_OK, WRITE_RACE_VERIFY_FORBIDDEN_SUBSTR: 'abcdefgh' }, good);
  ok('含禁止字串 → 中止(2)', a.r.code === 2 && noWrites(a));
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: URL_OK }, ['--yes-this-is-test-db']);
  ok('沒給 --expect-ref → 中止(2)', a.r.code === 2 && noWrites(a));
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: URL_OK }, ['--expect-ref=short', '--yes-this-is-test-db']);
  ok('--expect-ref 太短 → 中止(2)', a.r.code === 2 && noWrites(a));
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: URL_OK }, ['--expect-ref=zzzzzzzzzz', '--yes-this-is-test-db']);
  ok('--expect-ref 與連線字串不符 → 中止(2)', a.r.code === 2 && noWrites(a));

  // 看起來像正式庫（app_data 很大）
  const big = new FakeDb({ latency: [0, 1] });
  big.seedMain({ contacts: new Array(600).fill({}), opportunities: [] });
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: URL_OK }, good, { fake: big });
  ok('app_data 看起來像正式庫 → 中止(2)、無寫入', a.r.code === 2 && /正式資料庫/.test(a.text) && noWrites(a));

  // 互動確認
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: URL_OK }, ['--expect-ref=abcdefgh1234'], { opts: { ask: async () => 'wrong' } });
  ok('互動確認錯誤 → 中止(2)、無寫入', a.r.code === 2 && noWrites(a));
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: URL_OK }, ['--expect-ref=abcdefgh1234'], { opts: { ask: async () => 'aws-0-te' } });
  ok('互動確認正確（主機前 8 字元）→ 跑完且 0 失敗', a.r.code === 0 && a.r.failed === 0, a.r.code + ' ' + a.r.failed);

  // 正常流程
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: URL_OK }, good);
  ok('假 PG 上整套驗證 0 失敗', a.r.code === 0 && a.r.failed === 0 && a.r.passed >= 15, `code=${a.r.code} passed=${a.r.passed} failed=${a.r.failed}`);
  ok('結束後臨時 schema 已刪除', a.fake.schemas.size === 0);
  ok('輸出不含連線字串／密碼／主機', !/S3cretPw|aws-0-test\.pooler|postgres:\/\//.test(a.text));
  ok('沒有碰 app_data（無任何對 app_data 的寫入）', !a.fake.log.some((q) => /app_data/.test(q) && /^(INSERT|UPDATE|DELETE|CREATE|ALTER|DROP)/.test(q)));
  ok('鎖鍵不是正式預設鍵', !a.fake.log.some((q) => /pg_advisory_xact_lock\(7301420260101\)/.test(q)));
  ok('也接受沿用 AUDIT_VERIFY_DATABASE_URL', (await attempt({ AUDIT_VERIFY_DATABASE_URL: URL_OK }, good)).r.code === 0);

  // 中途失敗仍清理，且錯誤訊息不外洩密碼
  const f2 = new FakeDb({ latency: [0, 1] });
  f2.failNext(/^WITH c AS/, 1, new Error('boom S3cretPw!xyz aws-0-test.pooler.example.com'));
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: URL_OK }, good, { fake: f2 });
  ok('中途失敗 → 結束碼 1、仍刪除臨時 schema、訊息已遮蔽', a.r.code === 1 && f2.schemas.size === 0 && !/S3cretPw|aws-0-test\.pooler/.test(a.text), a.text.slice(-300));
  a = await attempt({ WRITE_RACE_VERIFY_DATABASE_URL: URL_OK }, [...good, '--keep']);
  ok('--keep 保留 schema 並提示', a.fake.schemas.size === 1 && /--keep/.test(a.text));

  let bad = 0;
  for (const [n, p, e] of results) { if (!p) bad++; console.log((p ? 'PASS ' : 'FAIL ') + n + (!p && e ? '  ← ' + e : '')); }
  console.log(`\n${results.length - bad}/${results.length} 通過`);
  process.exit(bad ? 1 : 0);
})().catch((e) => { console.error('例外', e); process.exit(2); });
