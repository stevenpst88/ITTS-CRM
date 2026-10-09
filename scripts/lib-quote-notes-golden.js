'use strict';
/**
 * check-quote-item-notes.js 的「位元級相容」黃金值產生器（不是獨立的檢查）。
 *
 * 用「沒有說明／備註」的一批確定性（固定亂數種子）報價單，在指定的程式碼樹（root）上算出一組摘要：
 *   hash      contentHash／itemsSig／costLinesSig／lineStructureSig／structureSig
 *   store     POST／PUT 經 normalizeItems 後實際存進資料庫的品項 JSON
 *   serialize GET /api/quotations/:id 對擁有者與管理員的完整回應 JSON（含 preview、approval、perm）
 *   excel     給客戶的 Excel（buildQuoteWorkbook）的 sheet1.xml
 *   preview   給客戶的網頁預覽 HTML（buildQuotePreviewHtml）
 * 黃金值（check-quote-item-notes.js 內寫死的 GOLDEN）是在「品項說明／備註功能加入之前」的程式碼（git HEAD bc921ee）上算出來的；
 * 新程式碼對沒有說明／備註的品項必須算出完全相同的摘要。要重算黃金值（只有「刻意改變舊輸出」時才該做）：
 *   用 git archive 取出舊版樹，node scripts/check-quote-item-notes.js --digests <舊版樹路徑>
 */
const path = require('path');
const fs = require('fs');
const vm = require('vm');
const crypto = require('crypto');

const sha = (s) => crypto.createHash('sha256').update(String(s), 'utf8').digest('hex').slice(0, 16);

/** 確定性亂數（mulberry32） */
function rng(seed) {
  let a = seed >>> 0;
  return () => { a = (a + 0x6D2B79F5) >>> 0; let t = a; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
}

const DESCS = ['系統導入顧問', '教育訓練', '軟體授權 SAP', 'ABAP 開發', '年度維護 MA', '硬體設備', '專案管理 PM', '<b>特殊</b>字元 & "引號"', 'Very long item description '.repeat(5), '全形ＡＢＣ　空白', 'line1\nline2'];
const UNITS = ['式', '人天', '套', '台', '月'];

/** 一批舊式報價單（沒有 spec／note）。回傳 [{ q, state }]，state 決定要放進資料庫的簽核狀態 */
function makeFixtures(R, n) {
  const QA = require(path.join(R, 'lib/quoteApproval.js'));
  const r = rng(20261009);
  const pick = (a) => a[Math.floor(r() * a.length)];
  const out = [];
  for (let i = 0; i < n; i++) {
    const nItems = 1 + Math.floor(r() * 7);
    const items = [];
    for (let k = 0; k < nItems; k++) {
      const roll = r();
      const lid = 'L' + i + '_' + k;
      if (roll < 0.12) items.push({ lid, kind: 'title', desc: 'Part ' + String.fromCharCode(65 + k) });
      else if (roll < 0.22) items.push({ lid, kind: 'subtotal', desc: r() < 0.5 ? '' : '小計 ' + k });
      else {
        const it = { lid, desc: pick(DESCS), unit: pick(UNITS), qty: pick([1, 2, 3, 0.5, 10, 7.25]), unitPrice: pick([0, 1000, 6500, 12345.67, 250000]), cost: pick([0, 500, 4000]) };
        if (r() < 0.3) it.cat = pick(['consult', 'software', 'hardware', 'other']);
        items.push(it);
      }
    }
    if (!items.some((x) => x.kind !== 'title' && x.kind !== 'subtotal')) items.push({ lid: 'L' + i + '_x', desc: '保底品項', unit: '式', qty: 1, unitPrice: 1000, cost: 0 });
    const q = {
      id: 'Q' + i, quoteNo: 'QU-G-' + String(i).padStart(3, '0'), owner: 'own1', company: pick(['甲公司', 'Acme <Co>', '乙 & 丙']), projectName: pick(['專案一', 'Project "X"', '']),
      contactName: '聯絡人', phone: '02-1234', address: '台北市', quoteDate: '2026-10-08', status: 'draft', createdAt: '2026-10-08T00:00:00.000Z', updatedAt: '2026-10-08T00:00:00.000Z',
      validUntil: '2026-10-30', products: ['PS'], costBy: null, costFlow: { state: 'na' },
      discountType: pick(['none', 'percent', 'amount']), discountValue: pick([0, 90, 5, 12345]), note: pick(['', '付款條件另議']),
      items, approval: null,
    };
    if (i % 5 === 1) {   // 一部分做成「已核准」（hash 由執行中的程式碼算，所以 valid 一定為真）
      q.approval = { state: 'approved', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), submittedAt: '2026-10-08T01:00:00.000Z', submittedBy: 'own1', derived: null, steps: [{ tier: 'mgr1', label: '一級主管', status: 'approved', by: 'mgr1', at: '2026-10-08T02:00:00.000Z', comment: '' }], cur: 1, board: null, history: [{ at: '2026-10-08T01:00:00.000Z', by: 'own1', action: 'SUBMIT', comment: '', tier: '' }] };
    } else if (i % 5 === 2) {
      q.approval = { state: 'pending', rulesVersion: QA.RULES_VERSION, hash: QA.contentHash(q), submittedAt: '2026-10-08T01:00:00.000Z', submittedBy: 'own1', derived: null, steps: [{ tier: 'mgr1', label: '一級主管', status: 'pending', assignee: 'mgr1' }], cur: 0, board: null, history: [] };
    }
    out.push(q);
  }
  return out;
}

const USERS = [
  { username: 'admin1', role: 'admin', active: true, displayName: 'Admin' },
  { username: 'own1', role: 'user', active: true, displayName: 'Owner1', supervisor: 'mgr1', bu: ['ITS'] },
  { username: 'own2', role: 'user', active: true, displayName: 'Owner2', supervisor: 'mgr1', bu: ['ITS'] },
  { username: 'mgr1', role: 'manager1', active: true, displayName: 'Mgr1', bu: ['ITS'] },
  { username: 'gm1', role: 'executive', active: true, displayName: 'Gm' },
  { username: 'ch1', role: 'executive', active: true, displayName: 'Chair' },
  { username: 'cons1', role: 'user', active: true, displayName: 'Cons1', bu: ['ITS'] },
  { username: 'sec1', role: 'secretary', active: true, displayName: 'Sec', bu: ['ITS'] },
  { username: 'x1', role: 'user', active: true, displayName: 'Other', bu: ['ERP'] },
];
const CLASSES = { PC: { cls: 'consult', costBySales: false }, PS: { cls: 'software', costBySales: true } };

/** 記憶體資料庫＋直接呼叫 handler 的路由測試環境（registerQuoteRoutes 來自 root 這棵樹） */
function mkEnv(R, quotations) {
  const registerQuoteRoutes = require(path.join(R, 'lib/quoteRoutes.js'));
  const PNL = require(path.join(R, 'lib/quotePnlExcel.js'));
  const routes = {};
  const app = {};
  ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
  const env = { data: { quotations, quoteApproval: { roster: { gm: ['gm1'], chairman: ['ch1'], boardProxy: [], costProviders: ['cons1'], sealManagers: [] }, productClasses: CLASSES }, contacts: [] }, logs: [], notes: [], saves: 0, seq: 0, nseq: 0 };
  env.auth = { users: JSON.parse(JSON.stringify(USERS)) };
  const deps = {
    // 與 db/json.js 相同的語意：load() 每次回傳「重新讀檔」的新物件、save(d) 才寫入。（曾經因為測試用同一個物件，沒測到通知合併會被過期的 ctx.data 蓋掉的問題）
    db: { load: () => JSON.parse(JSON.stringify(env.data)), save: (d) => { env.saves++; env.data = JSON.parse(JSON.stringify(d)); }, flush: async () => {} }, loadAuth: () => env.auth, saveAuth: () => {},
    requireAuth: (req, rs, next) => next(), requireAdmin: (req, rs, next) => next(),
    writeLog: (...a) => env.logs.push(a),
    // 與 server.js 的 pushNotification 同形：unshift 進 data.notifications（{id,to,type,title,body,refId,read,createdAt}），測試另外用 env.notes 看呼叫參數
    pushNotification: (to, type, title, body, refId) => { env.notes.push([to, type, title, body, refId]); const d = deps.db.load(); (d.notifications = d.notifications || []).unshift({ id: 'n' + (++env.nseq), to, type, title, body, refId, read: false, createdAt: new Date().toISOString() }); deps.db.save(d); },
    getViewableOwners: (req) => { const rl = req.session.user.role; if (rl === 'admin' || rl === 'executive' || rl === 'manager1' || rl === 'secretary') return ['own1', 'own2']; return [req.session.user.username]; },
    sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-NEW', taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}),
    buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: PNL.buildQuotePnlExcel, QUOTE_TEMPLATE: '',
    uuidv4: () => 'u' + (++env.seq), normalizeBu: (b) => (Array.isArray(b) ? b : b ? [b] : []), getUserFeatures: () => ['quotations'],
  };
  registerQuoteRoutes(app, deps);
  env.routesOf = (m, p) => routes[m.toUpperCase() + ' ' + p];
  env.call = (user, method, p, params, body) => new Promise((resolve, reject) => {
    const h = [].concat(...env.routesOf(method, p));
    const role = (env.auth.users.find((u) => u.username === user) || {}).role;
    const req = { session: { user: { username: user, role } }, params: params || {}, query: {}, headers: {}, body: body === undefined ? {} : JSON.parse(JSON.stringify(body)) };
    const rs = { status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j }); return this; }, send() { resolve({ s: this._s || 200, j: null }); return this; }, setHeader() {}, set() {} };
    let i = 0;
    const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
    next();
  });
  return env;
}

/** 給客戶的預覽 HTML：把 quote-preview.js 放進 vm（只需要 escapeHtml 與 taipeiTodayClient 兩個前端全域函式） */
function loadPreview(R) {
  const src = fs.readFileSync(path.join(R, '_client/quote-preview.js'), 'utf8');
  const ctx = {
    console, window: {}, document: { getElementById: () => null },
    escapeHtml: (str) => (str === null || str === undefined ? '' : String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;')),
    taipeiTodayClient: () => '2026-10-08', API: '/api',
  };
  vm.createContext(ctx);
  vm.runInContext(src + '\n;globalThis.__build = buildQuotePreviewHtml; globalThis.__clean = (typeof _qpvCleanText === "function") ? _qpvCleanText : null;', ctx);
  return { build: ctx.__build, clean: ctx.__clean, ctx };
}

async function sheetXml(R, q, opts) {
  const JSZip = require('jszip');
  const QE = require(path.join(R, 'lib/quoteExcel.js'));
  const buf = await QE.buildQuoteWorkbook(q, path.join(R, 'templates/quotation_template.xlsx'), Object.assign({ issueDate: '2026-10-09', issuer: { name: '業務', phone: '02-1', ext: '1', mobile: '0900' }, approved: false, seal: null }, opts || {}));
  const z = await JSZip.loadAsync(buf);
  return { xml: await z.file('xl/worksheets/sheet1.xml').async('string'), buf, zip: z };
}

/** 摘要們（沒有說明／備註的舊式資料） */
async function goldenDigests(R) {
  const QA = require(path.join(R, 'lib/quoteApproval.js'));
  const N = 40;
  const fx = makeFixtures(R, N);
  const out = {};
  out.hash = sha(fx.map((q) => [QA.contentHash(q), QA.itemsSig(q), QA.costLinesSig(q), QA.lineStructureSig(q.items), QA.structureSig(q)].join('|')).join('\n'));
  // store：用 POST 建立（走 normalizeItems）再用 PUT 原樣送回（舊畫面形狀：沒有 spec／note 欄位）
  const bodies = fx.map((q) => ({ company: q.company, projectName: q.projectName, products: ['PS'], validUntil: '2026-10-30', discountType: q.discountType, discountValue: q.discountValue, items: q.items.map((it) => Object.assign({}, it, { lid: undefined })) }));
  const envS = mkEnv(R, []);
  const stored = [];
  for (const b of bodies) {
    const c = await envS.call('own1', 'POST', '/api/quotations', {}, b);
    if (c.s !== 201) throw new Error('golden POST failed ' + c.s + ' ' + JSON.stringify(c.j));
    const id = c.j.id;
    const u = await envS.call('own1', 'PUT', '/api/quotations/:id', { id }, { items: c.j.items.map((it) => Object.assign({}, it, { desc: it.desc })), rowKinds: 1 });
    if (u.s !== 200) throw new Error('golden PUT failed ' + u.s + ' ' + JSON.stringify(u.j));
    stored.push(envS.data.quotations.find((x) => x.id === id).items);
  }
  out.store = sha(JSON.stringify(stored));
  // serialize：GET 擁有者與管理員
  const envG = mkEnv(R, JSON.parse(JSON.stringify(fx)));
  const ser = [];
  for (const q of fx) {
    for (const u of ['own1', 'admin1']) {
      const g = await envG.call(u, 'GET', '/api/quotations/:id', { id: q.id });
      ser.push(g.s + JSON.stringify(g.j));
    }
  }
  out.serialize = sha(ser.join('\n'));
  // excel（sheet1.xml）與預覽 HTML
  const xs = [];
  for (const q of fx) xs.push((await sheetXml(R, q)).xml);
  out.excel = sha(xs.join('\n'));
  const P = loadPreview(R);
  out.preview = sha(fx.map((q) => P.build(q, { issueDate: '2026-10-09', issuer: { name: '業務', phone: '02-1', ext: '1', mobile: '0900' }, remarks: ['1.a', '2.b'] })).join('\n'));
  return out;
}

module.exports = { goldenDigests, makeFixtures, mkEnv, loadPreview, sheetXml, sha, USERS, CLASSES };
