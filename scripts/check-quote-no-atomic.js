#!/usr/bin/env node
/**
 * 報價單號原子取號（QUOTE_NO_ATOMIC=1）呼叫端檢查。用法：node scripts/check-quote-no-atomic.js（不需要伺服器、不連資料庫；假 app＋假 db 直接呼叫路由 handler）
 * 保護 lib/quoteRoutes.js 的 POST /api/quotations 與 GET /api/admin/quote-ledger 裡「原子取號」那段接線，三個核心不變量：
 *   1) 取號失敗＝fail-closed：回 503 QUOTE_NO_UNAVAILABLE，不新增報價單、不存檔、不退回 blob 計數器給假號
 *   2) 取號是 await：之後必須重新 db.load()，拿舊物件整份存回會蓋掉別人的資料
 *   3) 傳給 DB 的下界＝blob 現況（max(quoteSeq, 現有報價單最大號)），取得的序號要寫回 blob 的 quoteSeq（日後關掉原子模式不重用號碼）
 * 另：開關關閉／JSON 後端（沒有 db.nextQuoteSeq）時完全沿用 server.js 的 genQuoteNo；quote-ledger 取不到已發號紀錄時要標 issuedUnavailable 並重取快取。
 * 測試資料只用通用字串（公開 repo）。
 */
'use strict';
const path = require('path');
const ROOT = path.join(__dirname, '..');
const PNL = require(path.join(ROOT, 'lib/quotePnlExcel.js'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);
let summaryPrinted = false;
process.on('exit', (code) => {
  if (summaryPrinted) return;
  res.forEach(([n, p, e]) => console.log((p ? 'PASS ' : 'FAIL ') + n + (e && !p ? '  ← ' + e : '')));
  console.log('\n【未完成】測試在印出摘要前就結束了（有 Promise 永不結束），視為失敗。');
  process.exitCode = code || 3;
});

function loadRoutes(atomic) {
  if (atomic) process.env.QUOTE_NO_ATOMIC = '1'; else delete process.env.QUOTE_NO_ATOMIC;
  const p = require.resolve(path.join(ROOT, 'lib/quoteRoutes.js'));
  delete require.cache[p];
  return require(p);
}

const USERS = [
  { username: 'admin1', role: 'admin', active: true, displayName: 'Admin' },
  { username: 'own1', role: 'user', active: true, displayName: 'Owner1', supervisor: 'mgr1', bu: ['ITS'] },
  { username: 'mgr1', role: 'manager1', active: true, displayName: 'Mgr1', bu: ['ITS'] },
  { username: 'gm1', role: 'executive', active: true, displayName: 'Gm' },
  { username: 'ch1', role: 'executive', active: true, displayName: 'Chair' },
  { username: 'cons1', role: 'user', active: true, displayName: 'Cons1', bu: ['ITS'] },
];
const CLASSES = { PS: { cls: 'software', costBySales: true } };
const opp = (id) => ({ id, owner: 'own1', stage: 'C', company: 'TestCo', product: 'Prod', amount: '800', bu: 'ITS', contactId: '', contactName: '' });
const oldQuote = () => ({ id: 'old1', quoteNo: 'QU-202610-004', owner: 'own1', company: 'TestCo', projectName: 'Old', quoteDate: '2026-10-01', status: 'draft', createdAt: '2026-10-01T00:00:00Z', updatedAt: '2026-10-01T00:00:00Z', validUntil: '2026-10-30', products: ['PS'], items: [{ lid: 'i1', desc: 'x', unit: '式', qty: 1, unitPrice: 1000, cost: 100 }], costFlow: { state: 'na' } });
const formBody = (oppId) => ({ oppId, company: 'TestCo', projectName: 'P', products: ['PS'], validUntil: '2026-10-30', items: [{ desc: '品項A', unit: '式', qty: 1, unitPrice: 1000, cost: 100 }] });

/** dbExtra：覆蓋／追加 db 的函式（nextQuoteSeq、listIssuedQuoteNos…） */
function mkEnv(register, dbExtra, dataOver) {
  const routes = {}; const app = {};
  ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
  const env = {
    data: Object.assign({ opportunities: [opp('OC')], lostOpportunities: [], quotations: [oldQuote()], quoteSeq: { '202610': 2 }, quoteApproval: { roster: { gm: ['gm1'], chairman: ['ch1'], boardProxy: [], costProviders: ['cons1'], sealManagers: [] }, productClasses: CLASSES } }, dataOver || {}),
    saved: [], genCalls: 0, seq: 0, logs: [],
  };
  env.auth = { users: JSON.parse(JSON.stringify(USERS)) };
  const db = Object.assign({ load: () => env.data, save: (d) => { env.saved.push(d); return Promise.resolve(); }, flush: async () => {} }, dbExtra || {});
  const deps = {
    db, loadAuth: () => env.auth, saveAuth: () => {},
    requireAuth: (req, rs, next) => next(), requireAdmin: (req, rs, next) => next(),
    writeLog: (...a) => env.logs.push(a), pushNotification: () => {},
    getViewableOwners: (req) => { const rl = req.session.user.role; if (rl === 'admin' || rl === 'executive' || rl === 'manager1') return ['own1']; return [req.session.user.username]; },
    sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200),
    genQuoteNo: () => { env.genCalls++; return 'GEN-' + (++env.seq); },
    taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}),
    buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: PNL.buildQuotePnlExcel, QUOTE_TEMPLATE: '',
    uuidv4: () => 'u' + (++env.seq), normalizeBu: (b) => (Array.isArray(b) ? b : b ? [b] : []), getUserFeatures: () => ['quotations'],
  };
  register(app, deps);
  env.call = (user, method, p, body) => new Promise((resolve, reject) => {
    const h = [].concat(...routes[method.toUpperCase() + ' ' + p]);
    const role = (env.auth.users.find((u) => u.username === user) || {}).role;
    const req = { session: { user: { username: user, role } }, params: {}, query: {}, body: body === undefined ? {} : JSON.parse(JSON.stringify(body)) };
    const rs = { status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j }); return this; }, send() { resolve({ s: this._s || 200, j: null }); return this; }, setHeader() {}, set() {} };
    let i = 0;
    const next = (err) => { if (err) return reject(err); const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
    next();
  });
  return env;
}

(async () => {
  const atomic = loadRoutes(true);

  // ═══ 1) 成功路徑：下界、序號、寫回 blob ═══
  {
    const calls = [];
    const env = mkEnv(atomic, { nextQuoteSeq: async (ym, floor, who) => { calls.push({ ym, floor, who }); return { serial: 5, quoteNo: 'QU-202610-005' }; } });
    const r = await env.call('own1', 'POST', '/api/quotations', formBody('OC'));
    const created = env.data.quotations[env.data.quotations.length - 1];
    t('1.1 取號成功 → 201，單號用 DB 發的（QU-202610-005），沒有呼叫 genQuoteNo', r.s === 201 && created.quoteNo === 'QU-202610-005' && env.genCalls === 0, r.s + ' ' + created.quoteNo);
    t('1.2 傳給 DB 的下界＝blob 現況 max(quoteSeq=2, 現有最大號=4)＝4；年月＝台灣當月；發給建單者', calls.length === 1 && calls[0].floor === 4 && calls[0].ym === '202610' && calls[0].who === 'own1', J(calls));
    t('1.3 取得的序號寫回 blob 的 quoteSeq（日後關掉原子模式不會重用號碼）', env.data.quoteSeq['202610'] === 5, J(env.data.quoteSeq));
    t('1.4 報價單與流水帳都在同一次存檔的物件裡', env.saved.length >= 1 && env.saved[env.saved.length - 1] === env.data && env.data.quotations.some((q) => q.quoteNo === 'QU-202610-005') && (env.data.quoteNoLedger || []).some((e) => e.no === 'QU-202610-005'), J((env.data.quoteNoLedger || []).map((e) => e.no)));
  }

  // ═══ 2) fail-closed ═══
  {
    const env = mkEnv(atomic, { nextQuoteSeq: async () => { throw new Error('db down'); } });
    const before = env.data.quotations.length;
    const err = console.error; console.error = () => {};
    let r; try { r = await env.call('own1', 'POST', '/api/quotations', formBody('OC')); } finally { console.error = err; }
    t('2.1 取號失敗 → 503 QUOTE_NO_UNAVAILABLE（可重試的友善訊息）', r.s === 503 && r.j.code === 'QUOTE_NO_UNAVAILABLE' && /尚未建立/.test(r.j.error), r.s + J(r.j));
    t('2.2 沒有新增報價單、沒有存檔、沒有退回 blob 計數器給假號（genQuoteNo 沒被呼叫）', env.data.quotations.length === before && env.saved.length === 0 && env.genCalls === 0, 'n=' + env.data.quotations.length + ' saves=' + env.saved.length + ' gen=' + env.genCalls);
    t('2.3 blob 計數器沒有被動到', env.data.quoteSeq['202610'] === 2);
  }

  // ═══ 3) await 之後重取快取 ═══
  {
    let env;
    env = mkEnv(atomic, {
      nextQuoteSeq: async () => {
        // 取號的 await 期間，別的請求的 ready() 把快取換成較新的物件（多了一張別人剛建的單）
        env.data = Object.assign({}, env.data, { quotations: env.data.quotations.concat([{ id: 'other', quoteNo: 'QU-202610-006', owner: 'own2' }]) });
        return { serial: 7, quoteNo: 'QU-202610-007' };
      },
    });
    const stale = env.data;
    const r = await env.call('own1', 'POST', '/api/quotations', formBody('OC'));
    const last = env.saved[env.saved.length - 1];
    t('3.1 取號 await 後重取快取：存檔的是「新物件」（含別人剛建的單），不是 await 前拿到的舊物件', r.s === 201 && last === env.data && last !== stale && last.quotations.some((q) => q.id === 'other') && last.quotations.some((q) => q.quoteNo === 'QU-202610-007'), r.s + ' ' + J(last && last.quotations.map((q) => q.quoteNo)));
    t('3.2 舊物件沒有被存回（沒有被修改成最終狀態）', !stale.quotations.some((q) => q.quoteNo === 'QU-202610-007'));
  }

  // ═══ 4) 關閉／JSON 後端：沿用 genQuoteNo ═══
  {
    const off = loadRoutes(false);
    let called = 0;
    const env = mkEnv(off, { nextQuoteSeq: async () => { called++; return { serial: 9, quoteNo: 'QU-202610-009' }; } });
    const r = await env.call('own1', 'POST', '/api/quotations', formBody('OC'));
    t('4.1 開關關閉：單號來自 genQuoteNo，nextQuoteSeq 完全沒被呼叫', r.s === 201 && env.genCalls === 1 && called === 0 && env.data.quotations.some((q) => /^GEN-/.test(q.quoteNo)), r.s + ' called=' + called);
    const on = loadRoutes(true);
    const env2 = mkEnv(on, {});   // 開關開但 db 沒有 nextQuoteSeq（JSON 後端）
    const r2 = await env2.call('own1', 'POST', '/api/quotations', formBody('OC'));
    t('4.2 開關開、但後端沒有 nextQuoteSeq（JSON）：沿用 genQuoteNo', r2.s === 201 && env2.genCalls === 1);
  }

  // ═══ 5) quote-ledger：取不到已發號紀錄 ═══
  {
    const on = loadRoutes(true);
    let env;
    env = mkEnv(on, {
      listIssuedQuoteNos: async () => {
        env.data = Object.assign({}, env.data, { marker: 'reloaded' });   // await 期間快取被換
        throw new Error('pool timeout');
      },
    }, { quotations: [Object.assign(oldQuote(), { id: 'q9', quoteNo: 'QU-202610-009' })] });   // 單在 blob 內但流水帳沒有 → review 會補登 → 需要存檔
    const err = console.error; console.error = () => {};
    let r; try { r = await env.call('admin1', 'GET', '/api/admin/quote-ledger'); } finally { console.error = err; }
    t('5.1 取不到已發號紀錄：仍回 200，並標示 issuedUnavailable（缺號清單未併入已發號紀錄）', r.s === 200 && r.j.issuedUnavailable === true, r.s + J(Object.keys(r.j || {})));
    const last = env.saved[env.saved.length - 1];
    t('5.2 補登後存檔的是 await 後重取的新物件（不是舊物件）', !!last && last === env.data && last.marker === 'reloaded', last && last.marker);
    const ok2 = mkEnv(on, { listIssuedQuoteNos: async () => [] });
    const r2 = await ok2.call('admin1', 'GET', '/api/admin/quote-ledger');
    t('5.3 取得成功時回應沒有 issuedUnavailable，且有 issuedUnused 欄位', r2.s === 200 && r2.j.issuedUnavailable === undefined && r2.j.summary.issuedUnused === 0, J(r2.j.summary));
    // 5.4 計數器自癒：blob 的 quoteSeq(2) 比已發號紀錄的最大序號(9) 小（曾被舊快照覆蓋）→ 補到 9 並存檔；之後關掉原子模式也不會重發
    const heal = mkEnv(on, { listIssuedQuoteNos: async () => [{ quoteNo: 'QU-202610-009', ym: '202610', serial: 9, issuedTo: 'own1', issuedAt: '2020-01-01T00:00:00Z' }] });
    const r3 = await heal.call('admin1', 'GET', '/api/admin/quote-ledger');
    t('5.4 blob 計數器落後已發號最大序號 → 自動補到最大序號、存檔、回應列出 counterHealed', r3.s === 200 && heal.data.quoteSeq['202610'] === 9 && heal.saved.length >= 1 && J(r3.j.counterHealed) === J([{ ym: '202610', from: 2, to: 9 }]), J(r3.j.counterHealed) + ' ' + J(heal.data.quoteSeq));
    const noHeal = mkEnv(on, { listIssuedQuoteNos: async () => [{ quoteNo: 'QU-202610-001', ym: '202610', serial: 1, issuedTo: 'own1', issuedAt: '2020-01-01T00:00:00Z' }] });
    const r4 = await noHeal.call('admin1', 'GET', '/api/admin/quote-ledger');
    t('5.5 計數器本來就不落後：不改動（只增不減）', r4.j.counterHealed === undefined && noHeal.data.quoteSeq['202610'] === 2);
  }

  let bad = 0;
  for (const [n, p, e] of res) { if (!p) bad++; console.log((p ? 'PASS ' : 'FAIL ') + n + (e && !p ? '  ← ' + e : '')); }
  console.log(`\n${res.length - bad}/${res.length} 通過` + (bad ? `，${bad} 失敗` : ''));
  summaryPrinted = true;
  process.exit(bad ? 1 : 0);
})().catch((e) => { console.error('測試腳本例外：', e); process.exit(2); });
