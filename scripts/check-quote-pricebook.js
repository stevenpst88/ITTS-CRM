#!/usr/bin/env node
/**
 * 報價牌價簿（lib/quotePricebook.js ＋ lib/quoteRoutes.js 的 /api/quote-pricebook、/api/admin/quote-pricebook）單元／路由層測試。
 * 用法：node scripts/check-quote-pricebook.js（不需要伺服器；不碰 data.json／auth.json）
 * 動 lib/quotePricebook.js 或 lib/quoteRoutes.js 的牌價簿區段之後必跑。
 *   1) 金額解析 parseMoney：數字／純數字字串通過；NaN、Infinity、負數、超大、空字串、科學記號、布林、null、物件一律拒絕；四捨五入到小數 2 位
 *   2) normalizePricebook：名稱（空白、全空白、過長、控制字元、重複不分大小寫且先 trim）、牌價／成本、active 型別、60 筆上限、
 *      XSS 字樣的名稱原樣當純文字存放、id 穩定（保留合法 id、重複 id 拒絕、缺漏／不合法才產生、新 id 不撞既有）、unit 固定人天、多餘欄位丟棄、不改動輸入
 *   3) publicItems／adminItems／diffPricebook／summarizeDiff（新增、移除、改名、牌價成本 old→new、啟停用、順序）
 *   4) 路由層（假 app＋記憶體 db 直接呼叫 handler）：業務端只看到啟用項目（含牌價＋成本）、沒有報價單功能的角色 403、被指派的成本填寫人可讀；
 *      管理員 GET 含停用項目；非管理員 PUT／GET admin 403 且沒有任何寫入；驗證失敗 400 且不動資料；樂觀並行 409 STALE_PRICEBOOK（附現況）；
 *      成功儲存：data.pricebook 落地、id 產生、updatedBy、稽核 SAVE_QUOTE_PRICEBOOK 摘要含實際異動；已存報價單與其他命名空間完全不受影響
 *   5) 管理員路由確實掛了 requireAdmin（不是只靠 handler 內的角色檢查）；讀取路由掛 requireAuth
 * 測試資料只用通用字串（公開 repo：不放客戶名、人名、真實費率）。
 */
'use strict';
const path = require('path');
const ROOT = path.join(__dirname, '..');
const PB = require(path.join(ROOT, 'lib/quotePricebook.js'));
const registerQuoteRoutes = require(path.join(ROOT, 'lib/quoteRoutes.js'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const eq = (a, b) => JSON.stringify(a) === JSON.stringify(b);
let _n = 0;
const genId = () => 'gen' + (++_n);
const norm = (raw) => PB.normalizePricebook(raw, { genId });
const row = (name, price, cost, extra) => Object.assign({ name, price, cost }, extra || {});
const code = (r) => (r.ok ? 'OK' : r.error.code + ':' + (r.error.field || '') + '@' + r.error.index);

async function run() {
  // ═════════════════ 1) parseMoney ═════════════════
  const okMoney = [[0, 0], [7000, 7000], ['7000', 7000], [' 12.5 ', 12.5], [1e9, 1e9], [0.005, 0.01], [1.234, 1.23], ['0', 0], ['00012', 12]];
  t('1.1 parseMoney 通過：數字、純數字字串（含前後空白）、0、上限 1e9、四捨五入到小數 2 位', okMoney.every(([i, o]) => PB.parseMoney(i) === o), okMoney.map(([i]) => PB.parseMoney(i)).join(','));
  const badMoney = [NaN, Infinity, -Infinity, -1, -0.01, 1e9 + 1, 1e12, '', ' ', 'abc', '12px', '1e3', '-5', '+5', '5,000', '.5', '5.', null, undefined, true, false, {}, [], [5], '0x10'];
  t('1.2 parseMoney 拒絕：NaN、±Infinity、負數、超過上限、空字串、非純數字字串、科學記號、布林、null、物件／陣列', badMoney.every((v) => PB.parseMoney(v) === null), badMoney.filter((v) => PB.parseMoney(v) !== null).map(String).join('|'));

  // ═════════════════ 2) normalizePricebook ═════════════════
  let r = norm([row('PM 顧問經理', 9000, 6500), row('SD 顧問', '7000', '5000', { active: false })]);
  t('2.1 正常：欄位齊全；unit 固定「人天」；active 預設 true；字串金額轉數字；順序＝陣列順序', r.ok && eq(r.items.map((x) => [x.name, x.unit, x.price, x.cost, x.active]), [['PM 顧問經理', '人天', 9000, 6500, true], ['SD 顧問', '人天', 7000, 5000, false]]), JSON.stringify(r));
  t('2.2 每項都有 id（缺漏時由 genId 產生，彼此不同）', r.ok && r.items.every((x) => /^[A-Za-z0-9_-]{1,40}$/.test(x.id)) && r.items[0].id !== r.items[1].id);
  t('2.3 client 送的 unit／cat／其他欄位一律丟棄（unit 強制人天）', (() => { const x = norm([row('A', 1, 1, { unit: '式', cat: 'hardware', evil: '<x>', extra: 1 })]); return x.ok && eq(Object.keys(x.items[0]).sort(), ['active', 'cost', 'id', 'name', 'price', 'unit']) && x.items[0].unit === '人天'; })());
  t('2.4 空陣列合法（清空牌價簿）', eq(norm([]), { ok: true, items: [] }));
  const bad = [
    ['2.5a 不是陣列', 'x', 'BAD_PRICEBOOK:@undefined'], ['2.5b null', null, 'BAD_PRICEBOOK:@undefined'], ['2.5c 物件', {}, 'BAD_PRICEBOOK:@undefined'],
    ['2.6a 項目不是物件', [null], 'BAD_PRICEBOOK:@0'], ['2.6b 項目是字串', ['PM'], 'BAD_PRICEBOOK:@0'], ['2.6c 項目是陣列', [[]], 'BAD_PRICEBOOK:@0'],
    ['2.7a 名稱空字串', [row('', 1, 1)], 'BAD_PRICEBOOK:name@0'], ['2.7b 名稱全空白（含全形空白）', [row(' 　 ', 1, 1)], 'BAD_PRICEBOOK:name@0'], ['2.7c 名稱不是字串', [row(5, 1, 1)], 'BAD_PRICEBOOK:name@0'],
    ['2.7d 名稱缺漏', [{ price: 1, cost: 1 }], 'BAD_PRICEBOOK:name@0'], ['2.7e 名稱 41 字', [row('x'.repeat(41), 1, 1)], 'BAD_PRICEBOOK:name@0'], ['2.7f 名稱含換行', [row('a\nb', 1, 1)], 'BAD_PRICEBOOK:name@0'], ['2.7g 名稱含 NUL', [row('a\u0000b', 1, 1)], 'BAD_PRICEBOOK:name@0'],
    ['2.8a 名稱重複', [row('PM', 1, 1), row('PM', 2, 2)], 'BAD_PRICEBOOK:name@1'], ['2.8b 名稱重複（大小寫不同）', [row('pm', 1, 1), row('PM', 2, 2)], 'BAD_PRICEBOOK:name@1'], ['2.8c 名稱重複（前後空白不同）', [row('PM', 1, 1), row(' PM ', 2, 2)], 'BAD_PRICEBOOK:name@1'], ['2.8d 名稱重複（全形／半形）', [row('AB', 1, 1), row('ＡＢ', 2, 2)], 'BAD_PRICEBOOK:name@1'], ['2.8e 名稱重複（內部連續空白）', [row('p m', 1, 1), row('p  m', 2, 2)], 'BAD_PRICEBOOK:name@1'],
    ['2.9a 牌價 NaN', [row('A', NaN, 1)], 'BAD_PRICEBOOK:price@0'], ['2.9b 牌價負數', [row('A', -1, 1)], 'BAD_PRICEBOOK:price@0'], ['2.9c 牌價超大', [row('A', 1e9 + 1, 1)], 'BAD_PRICEBOOK:price@0'], ['2.9d 牌價非數字字串', [row('A', 'abc', 1)], 'BAD_PRICEBOOK:price@0'],
    ['2.9e 牌價缺漏', [{ name: 'A', cost: 1 }], 'BAD_PRICEBOOK:price@0'], ['2.9f 牌價 Infinity', [row('A', Infinity, 1)], 'BAD_PRICEBOOK:price@0'],
    ['2.10a 成本 NaN', [row('A', 1, NaN)], 'BAD_PRICEBOOK:cost@0'], ['2.10b 成本負數', [row('A', 1, '-3')], 'BAD_PRICEBOOK:cost@0'], ['2.10c 成本超大', [row('A', 1, 1e10)], 'BAD_PRICEBOOK:cost@0'], ['2.10d 成本空字串', [row('A', 1, '')], 'BAD_PRICEBOOK:cost@0'], ['2.10e 成本 null', [row('A', 1, null)], 'BAD_PRICEBOOK:cost@0'],
    ['2.11a active 不是布林（字串）', [row('A', 1, 1, { active: 'false' })], 'BAD_PRICEBOOK:active@0'], ['2.11b active 是 0', [row('A', 1, 1, { active: 0 })], 'BAD_PRICEBOOK:active@0'],
    ['2.12 第二列出錯 → 錯誤指向第二列（index=1）', [row('A', 1, 1), row('B', -5, 1)], 'BAD_PRICEBOOK:price@1'],
    ['2.13 重複 id', [row('A', 1, 1, { id: 'same' }), row('B', 1, 1, { id: 'same' })], 'BAD_PRICEBOOK:id@1'],
  ];
  for (const [name, input, want] of bad) t(name + ' → 拒絕', code(norm(input)) === want, code(norm(input)));
  t('2.14 恰好 60 筆通過、61 筆拒絕', norm(Array.from({ length: 60 }, (_, i) => row('R' + i, 1, 1))).ok && code(norm(Array.from({ length: 61 }, (_, i) => row('R' + i, 1, 1)))) === 'BAD_PRICEBOOK:@undefined');
  t('2.15 名稱恰好 40 字通過；前後空白被 trim；牌價成本 0 通過；成本高於牌價也通過（只有 UI 提醒）', (() => { const x = norm([row(' ' + 'x'.repeat(40) + ' ', 0, 0), row('B', 100, 900)]); return x.ok && x.items[0].name.length === 40 && x.items[1].cost === 900; })());
  const xss = ['<script>alert(1)</script>', '"><img src=x onerror=alert(1)>', "'; DROP TABLE x;--", '&lt;b&gt;', '<b>PM</b>'];
  const xr = norm(xss.map((n, i) => row(n.slice(0, 40), i, i)));
  t('2.16 XSS 字樣的名稱：原樣當純文字存放（不轉義、不過濾；顯示端負責轉義），長度內都通過', xr.ok && xr.items.every((x, i) => x.name === xss[i].slice(0, 40).trim()), JSON.stringify(xr.items && xr.items.map((x) => x.name)));
  t('2.17 id 穩定：合法 id 原樣保留（含改名、調序後）', (() => { const x = norm([row('B', 1, 1, { id: 'idB' }), row('A2', 1, 1, { id: 'idA' })]); return x.ok && x.items[0].id === 'idB' && x.items[1].id === 'idA'; })());
  t('2.18 不合法的 id（含空白／太長／非字串）重新產生；新產生的 id 不會撞上後面列指定的 id', (() => {
    const x = PB.normalizePricebook([row('A', 1, 1, { id: 'bad id!' }), row('B', 1, 1, { id: 'gen1' }), row('C', 1, 1, { id: 5 }), row('D', 1, 1, { id: 'x'.repeat(41) })], { genId: (() => { let k = 0; return () => 'gen' + (++k); })() });
    const ids = x.ok ? x.items.map((i) => i.id) : [];
    return x.ok && new Set(ids).size === 4 && ids[1] === 'gen1' && ids.every((i) => /^[A-Za-z0-9_-]{1,40}$/.test(i));
  })());
  t('2.19 不改動輸入物件', (() => { const input = [row(' A ', '5', '3', { id: 'k' })]; const snap = JSON.stringify(input); norm(input); return JSON.stringify(input) === snap; })());
  let threw = false; try { PB.normalizePricebook([], {}); } catch (e) { threw = true; }
  t('2.20 沒給 genId 直接丟錯（不會默默產生 undefined id）', threw);

  // ═════════════════ 3) public / admin / diff ═════════════════
  const stored = [{ id: 'a', name: 'PM', unit: '人天', price: 9000, cost: 6000, active: true }, { id: 'b', name: 'SD', unit: '人天', price: 7000, cost: 5000, active: false }, { id: 'c', name: 'MM', unit: '人天', price: 7500, cost: 5200, active: true }];
  t('3.1 publicItems：只留啟用項目、只有 id／name／unit／price／cost，順序不變', eq(PB.publicItems(stored), [{ id: 'a', name: 'PM', unit: '人天', price: 9000, cost: 6000 }, { id: 'c', name: 'MM', unit: '人天', price: 7500, cost: 5200 }]));
  t('3.2 adminItems：全部（含停用），多一個 active', PB.adminItems(stored).length === 3 && PB.adminItems(stored)[1].active === false);
  t('3.3 publicItems／adminItems 對壞資料不丟錯（undefined／非陣列／含 null）', eq(PB.publicItems(undefined), []) && eq(PB.adminItems('x'), []) && PB.publicItems([null, stored[0]]).length === 1);
  const next = [{ id: 'c', name: 'MM 顧問', unit: '人天', price: 8000, cost: 5200, active: true }, { id: 'a', name: 'PM', unit: '人天', price: 9000, cost: 6500, active: false }, { id: 'n', name: 'ABAP', unit: '人天', price: 6000, cost: 4000, active: true }];
  const d = PB.diffPricebook(stored, next);
  t('3.4 diff：新增 ABAP、移除 SD、改名 MM→MM 顧問、MM 牌價 7500→8000、PM 成本 6000→6500、PM 停用、順序調整', d.added.length === 1 && d.added[0].name === 'ABAP' && d.removed.length === 1 && d.removed[0].name === 'SD' && eq(d.renamed, [{ from: 'MM', to: 'MM 顧問' }])
    && d.changed.some((x) => x.name === 'MM 顧問' && x.field === 'price' && x.from === 7500 && x.to === 8000) && d.changed.some((x) => x.name === 'PM' && x.field === 'cost' && x.from === 6000 && x.to === 6500)
    && eq(d.toggled, [{ name: 'PM', active: false }]) && d.reordered === true, JSON.stringify(d));
  const sm = PB.summarizeDiff(d);
  t('3.5 summarizeDiff：文字含「新增」「移除」「改名」「牌價 7500→8000」「成本 6000→6500」「停用」「順序調整」', ['新增', '「ABAP」', '移除', '「SD」', '改名', '「MM」→「MM 顧問」', '牌價 7500→8000', '成本 6000→6500', '「PM」停用', '順序調整'].every((s) => sm.includes(s)), sm);
  t('3.6 無差異 → 「無異動」且 reordered=false；只換順序 → 只有順序調整', PB.summarizeDiff(PB.diffPricebook(stored, stored)) === '無異動' && PB.summarizeDiff(PB.diffPricebook(stored, [stored[2], stored[1], stored[0]])) === '順序調整');
  t('3.7 摘要過長會截斷（預設 1500 字）', PB.summarizeDiff(PB.diffPricebook([], Array.from({ length: 60 }, (_, i) => ({ id: 'i' + i, name: 'R' + 'x'.repeat(30) + i, price: 1000000, cost: 1000000, active: true })))).length <= 1500);

  // ═════════════════ 4) 路由層 ═════════════════
  const USERS = [
    { username: 'admin1', role: 'admin', active: true, displayName: 'Admin' },
    { username: 'own1', role: 'user', active: true, displayName: 'Owner1', supervisor: 'mgr1', bu: ['ITS'] },
    { username: 'mgr1', role: 'manager1', active: true, displayName: 'Mgr1', bu: ['ITS'] },
    { username: 'cons1', role: 'consultx', active: true, displayName: 'Cons1', bu: ['ITS'] },
    { username: 'nofeat', role: 'nofeat', active: true, displayName: 'NoFeat', bu: ['ITS'] },
    { username: 'grp1', role: 'tecopm', active: true, displayName: 'Grp', bu: ['ITS'] },
    { username: 'gone', role: 'user', active: false, displayName: 'Disabled' },
  ];
  function mkEnv() {
    const routes = {};
    const app = {};
    ['get', 'post', 'put', 'delete'].forEach((m) => { app[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
    const quote = { id: 'Q1', quoteNo: 'QU-1', owner: 'own1', company: 'TestCo', projectName: 'Proj', status: 'draft', products: [], items: [{ lid: 'i1', desc: 'PM', unit: '人天', qty: 3, unitPrice: 8000, cost: 0 }], approval: null, discountType: 'none', discountValue: 0 };
    const env = { data: { quotations: [quote], quoteApproval: { roster: { gm: [], chairman: [], boardProxy: [], costProviders: ['cons1'], sealManagers: [] }, productClasses: {} }, contacts: [{ id: 'c1' }] }, logs: [], saves: 0, seq: 0 };
    env.auth = { users: JSON.parse(JSON.stringify(USERS)) };
    const requireAuth = function requireAuthStub(req, rs, next) { next(); };
    const requireAdmin = function requireAdminStub(req, rs, next) { next(); };
    env.requireAuth = requireAuth; env.requireAdmin = requireAdmin;
    const deps = {
      db: { load: () => env.data, save: () => { env.saves++; }, flush: async () => {} }, loadAuth: () => env.auth, saveAuth: () => {},
      requireAuth, requireAdmin,
      writeLog: (...a) => env.logs.push(a), pushNotification: () => {},
      getViewableOwners: (req) => [req.session.user.username],
      sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-NEW', taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}),
      buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: async () => Buffer.from(''), QUOTE_TEMPLATE: '',
      uuidv4: () => 'u' + (++env.seq).toString().padStart(8, '0') + 'xxxx', normalizeBu: (b) => (Array.isArray(b) ? b : b ? [b] : []),
      getUserFeatures: (role) => (role === 'nofeat' || role === 'tecopm' ? [] : ['quotations']),
    };
    registerQuoteRoutes(app, deps);
    env.routesOf = (m, p) => routes[m.toUpperCase() + ' ' + p];
    env.call = (user, method, p, params, body) => new Promise((resolve, reject) => {
      const h = [].concat(...env.routesOf(method, p));
      const role = (env.auth.users.find((u) => u.username === user) || {}).role;
      const req = { session: { user: { username: user, role } }, params: params || {}, query: {}, body: body === undefined ? {} : JSON.parse(JSON.stringify(body)) };
      const rs = { status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j }); return this; }, send() { resolve({ s: this._s || 200, j: null }); return this; }, setHeader() {}, set() {} };
      let i = 0;
      const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } };
      next();
    });
    return env;
  }
  const A = '/api/admin/quote-pricebook', U = '/api/quote-pricebook';
  let env = mkEnv();
  let rg = await env.call('admin1', 'GET', A);
  t('4.1 尚未建立：管理員 GET → 200、items 空、updatedAt null、帶 limits', rg.s === 200 && eq(rg.j.items, []) && rg.j.updatedAt === null && rg.j.limits.MAX_ITEMS === 60, JSON.stringify(rg));
  let ru = await env.call('own1', 'GET', U);
  t('4.2 尚未建立：業務端 GET → 200、items 空', ru.s === 200 && eq(ru.j.items, []));
  const putBody = (items, updatedAt) => ({ items, updatedAt: updatedAt === undefined ? null : updatedAt });
  const baseSave = env.saves, snapOther = JSON.stringify({ q: env.data.quotations, qa: env.data.quoteApproval, c: env.data.contacts });
  let rp = await env.call('admin1', 'PUT', A, {}, putBody([row('PM 顧問經理', 9000, 6500), row('SD 顧問', 7000, 5000, { active: false }), row('ABAP 顧問', '6500', '4500')]));
  t('4.3 管理員第一次儲存 → 200；回傳 items（全部，含 id）、updatedAt、updatedBy', rp.s === 200 && rp.j.success === true && rp.j.items.length === 3 && rp.j.items.every((x) => x.id) && !!rp.j.updatedAt && rp.j.updatedBy === 'admin1', JSON.stringify(rp.j));
  t('4.4 落地：data.pricebook＝{items,updatedAt,updatedBy}（獨立命名空間）；db.save 呼叫 1 次', env.data.pricebook && env.data.pricebook.items.length === 3 && env.data.pricebook.updatedBy === 'admin1' && env.data.pricebook.updatedAt === rp.j.updatedAt && env.saves === baseSave + 1);
  t('4.5 已存報價單、簽核設定、其他命名空間一個位元都沒動', JSON.stringify({ q: env.data.quotations, qa: env.data.quoteApproval, c: env.data.contacts }) === snapOther);
  const lg = env.logs[env.logs.length - 1] || [];
  t('4.6 稽核：writeLog SAVE_QUOTE_PRICEBOOK、操作者、摘要含新增的三個項目與牌價／成本', lg[0] === 'SAVE_QUOTE_PRICEBOOK' && lg[1] === 'admin1' && ['新增', '「PM 顧問經理」(牌價9000/成本6500)', '「SD 顧問」(牌價7000/成本5000/停用)', '「ABAP 顧問」(牌價6500/成本4500)'].every((s) => String(lg[3]).includes(s)) && !!lg[4], JSON.stringify(lg.slice(0, 4)));
  ru = await env.call('own1', 'GET', U);
  t('4.7 業務端 GET：只有啟用項目（停用的 SD 不在）、有牌價與成本、順序不變、沒有 active／updatedBy', ru.s === 200 && eq(ru.j.items.map((x) => x.name), ['PM 顧問經理', 'ABAP 顧問']) && ru.j.items.every((x) => typeof x.price === 'number' && typeof x.cost === 'number' && x.unit === '人天' && !('active' in x)) && !('updatedBy' in ru.j), JSON.stringify(ru.j));
  const mg = await env.call('mgr1', 'GET', U);
  t('4.8 其他有報價單功能的角色（主管）也能讀', mg.s === 200 && mg.j.items.length === 2);
  const nf = await env.call('nofeat', 'GET', U);
  t('4.9 沒有報價單功能的角色 → 403 NO_PERMISSION，body 沒有項目', nf.s === 403 && nf.j.code === 'NO_PERMISSION' && !nf.j.items, JSON.stringify(nf));
  const gp = await env.call('grp1', 'GET', U);
  t('4.10 集團角色（沒有報價單功能）→ 403', gp.s === 403);
  env.auth.users.find((u) => u.username === 'cons1').role = 'nofeat';
  const cp = await env.call('cons1', 'GET', U);
  t('4.11 被指派的成本填寫人即使角色沒有報價單功能也能讀（要在顧問對話框帶入成本）', cp.s === 200 && cp.j.items.length === 2, JSON.stringify(cp));
  env.auth.users.find((u) => u.username === 'cons1').role = 'consultx';
  const dis = await env.call('gone', 'GET', U);
  t('4.12 已停用帳號 → 401 ACCOUNT_DISABLED', dis.s === 401 && dis.j.code === 'ACCOUNT_DISABLED');
  const ad = await env.call('admin1', 'GET', A);
  t('4.13 管理員 GET：全部三項（含停用的 SD）、含 active、updatedAt／updatedBy', ad.s === 200 && ad.j.items.length === 3 && ad.j.items[1].active === false && ad.j.updatedAt === rp.j.updatedAt && ad.j.updatedBy === 'admin1');
  const au = await env.call('admin1', 'GET', U);
  t('4.14 管理員走業務端路由也只看到啟用項目（兩條路由語意固定）', au.s === 200 && au.j.items.length === 2);

  // 非管理員寫入
  const snapDb = JSON.stringify(env.data.pricebook), sv = env.saves, nl = env.logs.length;
  const np = await env.call('own1', 'PUT', A, {}, putBody([row('Hack', 1, 1)], env.data.pricebook.updatedAt));
  const ng = await env.call('own1', 'GET', A);
  const mp = await env.call('mgr1', 'PUT', A, {}, putBody([], env.data.pricebook.updatedAt));
  t('4.15 非管理員 PUT／GET admin → 403 NO_PERMISSION，且沒有任何寫入、沒有稽核', np.s === 403 && np.j.code === 'NO_PERMISSION' && ng.s === 403 && mp.s === 403 && JSON.stringify(env.data.pricebook) === snapDb && env.saves === sv && env.logs.length === nl, JSON.stringify([np, ng, mp]));

  // 驗證失敗
  const live = () => env.data.pricebook;
  const badPut = [
    ['4.16a 名稱重複', [row('X', 1, 1), row('x', 1, 1)], 1, 'name'], ['4.16b 牌價 NaN', [row('X', 'NaN', 1)], 0, 'price'], ['4.16c 成本負數', [row('X', 1, -1)], 0, 'cost'],
    ['4.16d 名稱空白', [row(' ', 1, 1)], 0, 'name'], ['4.16e 超過 60 筆', Array.from({ length: 61 }, (_, i) => row('R' + i, 1, 1)), undefined, undefined],
  ];
  for (const [name, items, idx, field] of badPut) {
    const before = JSON.stringify(env.data.pricebook), s0 = env.saves, l0 = env.logs.length;
    const x = await env.call('admin1', 'PUT', A, {}, putBody(items, live().updatedAt));
    t(name + ' → 400 BAD_PRICEBOOK（附 index／field），資料與稽核都沒變', x.s === 400 && x.j.code === 'BAD_PRICEBOOK' && x.j.index === idx && x.j.field === field && JSON.stringify(env.data.pricebook) === before && env.saves === s0 && env.logs.length === l0, JSON.stringify(x));
  }
  const nb = await env.call('admin1', 'PUT', A, {}, { updatedAt: live().updatedAt });
  t('4.16f 沒帶 items → 400（不會把牌價簿清空）', nb.s === 400 && nb.j.code === 'BAD_PRICEBOOK' && live().items.length === 3);

  // 樂觀並行
  const loaded = live().updatedAt, firstIds = live().items.map((x) => x.id);
  const items2 = live().items.map((x) => Object.assign({}, x));
  items2[0].price = 9500; items2[0].cost = 6000;          // PM：牌價 9000→9500、成本 6500→6000
  items2[1].name = 'SD 資深顧問'; items2[1].active = true; // SD：改名＋啟用
  items2.splice(2, 1);                                      // 移除 ABAP
  items2.push(row('MM 顧問', 7500, 5200));                  // 新增 MM
  const r2 = await env.call('admin1', 'PUT', A, {}, putBody(items2, loaded));
  t('4.17 帶著載入時的 updatedAt 儲存 → 200；updatedAt 嚴格變新', r2.s === 200 && r2.j.updatedAt > loaded, JSON.stringify([loaded, r2.j.updatedAt]));
  t('4.18 id 穩定：未動的項目 id 不變（PM、SD 改名後仍同 id），新增的 MM 拿到新 id', r2.s === 200 && r2.j.items[0].id === firstIds[0] && r2.j.items[1].id === firstIds[1] && !firstIds.includes(r2.j.items[2].id) && r2.j.items.length === 3);
  const lg2 = env.logs[env.logs.length - 1];
  t('4.19 第二次稽核摘要是實際異動：牌價 9000→9500、成本 6500→6000、改名 SD→SD 資深顧問、SD 啟用、移除 ABAP、新增 MM', ['「PM 顧問經理」牌價 9000→9500', '「PM 顧問經理」成本 6500→6000', '「SD 顧問」→「SD 資深顧問」', '「SD 資深顧問」啟用', '移除 「ABAP 顧問」', '新增 「MM 顧問」(牌價7500/成本5200)'].every((s) => String(lg2[3]).includes(s)), lg2[3]);
  const stale = await env.call('admin1', 'PUT', A, {}, putBody([row('Other', 1, 1)], loaded));
  t('4.20 用舊的 updatedAt 儲存 → 409 STALE_PRICEBOOK，附現況（live.items／updatedAt），資料沒被覆蓋', stale.s === 409 && stale.j.code === 'STALE_PRICEBOOK' && stale.j.live.updatedAt === r2.j.updatedAt && stale.j.live.items.length === 3 && live().items.length === 3 && live().updatedAt === r2.j.updatedAt, JSON.stringify(stale).slice(0, 200));
  const stale2 = await env.call('admin1', 'PUT', A, {}, putBody([row('Other', 1, 1)], null));
  const stale3 = await env.call('admin1', 'PUT', A, {}, { items: [row('Other', 1, 1)] });
  t('4.21 沒帶 updatedAt／帶 null，但伺服器已有資料 → 409（不能無條件覆蓋）', stale2.s === 409 && stale3.s === 409 && live().items.length === 3);
  const stale4 = await env.call('admin1', 'PUT', A, {}, putBody([row('Other', 1, 1)], 123));
  t('4.22 updatedAt 型別錯誤 → 409（當作沒看過現況）', stale4.s === 409);
  // 同毫秒連續儲存：updatedAt 仍然遞增、第二個人用舊值會被擋
  const t0 = live().updatedAt;
  const q1 = await env.call('admin1', 'PUT', A, {}, putBody(live().items, t0));
  const q2 = await env.call('admin1', 'PUT', A, {}, putBody(live().items, q1.j.updatedAt));
  const q3 = await env.call('admin1', 'PUT', A, {}, putBody(live().items, q1.j.updatedAt));
  t('4.23 連續儲存 updatedAt 一律嚴格遞增；拿前一版 updatedAt 的第三次儲存被擋（409）', q1.s === 200 && q2.s === 200 && q1.j.updatedAt > t0 && q2.j.updatedAt > q1.j.updatedAt && q3.s === 409, JSON.stringify([t0, q1.j.updatedAt, q2.j.updatedAt, q3.s]));
  t('4.24 無異動儲存也合法，稽核寫「無異動」', env.logs[env.logs.length - 1][3].includes('無異動'));
  // 停用後業務端不見；清空
  const items3 = live().items.map((x) => Object.assign({}, x, { active: false }));
  const r3 = await env.call('admin1', 'PUT', A, {}, putBody(items3, live().updatedAt));
  ru = await env.call('own1', 'GET', U);
  t('4.25 全部停用 → 業務端清單為空（管理員仍看得到全部）', r3.s === 200 && eq(ru.j.items, []) && (await env.call('admin1', 'GET', A)).j.items.length === 3);
  const r4 = await env.call('admin1', 'PUT', A, {}, putBody([], live().updatedAt));
  t('4.26 可以整份清空（空陣列合法）', r4.s === 200 && eq(r4.j.items, []) && eq(live().items, []));
  // XSS 名稱：路由原樣回傳純文字（JSON），不轉義
  const xs = await env.call('admin1', 'PUT', A, {}, putBody([row('<img src=x onerror=alert(1)>', 1, 1)], live().updatedAt));
  ru = await env.call('own1', 'GET', U);
  t('4.27 XSS 字樣的名稱：儲存、回傳都是原樣純文字（轉義是顯示端的責任，admin.html／quote-pricebook.js 另有靜態檢查）', xs.s === 200 && ru.j.items[0].name === '<img src=x onerror=alert(1)>');
  t('4.28 已存報價單與簽核設定在整串牌價簿操作後仍然完全不變', JSON.stringify({ q: env.data.quotations, qa: env.data.quoteApproval, c: env.data.contacts }) === snapOther);
  t('4.29 業務端 GET 不會寫入（不呼叫 db.save）', await (async () => { const s0 = env.saves; await env.call('own1', 'GET', U); await env.call('admin1', 'GET', A); return env.saves === s0; })());

  // ═════════════════ 5) 路由掛載 ═════════════════
  const flat = (m, p) => [].concat(...env.routesOf(m, p));
  t('5.1 三條路由都有註冊', !!env.routesOf('GET', U) && !!env.routesOf('GET', A) && !!env.routesOf('PUT', A));
  t('5.2 管理員路由（GET／PUT）第一個 middleware 是 requireAdmin；業務端 GET 是 requireAuth', flat('PUT', A)[0] === env.requireAdmin && flat('GET', A)[0] === env.requireAdmin && flat('GET', U)[0] === env.requireAuth);
  t('5.3 沒有註冊 POST／DELETE（只有 GET＋PUT 整份取代）', !env.routesOf('POST', A) && !env.routesOf('DELETE', A) && !env.routesOf('POST', U) && !env.routesOf('PUT', U));

  // ═════════════════ 6) 選取後的品項＝一般品項（伺服器不認得牌價簿、沒有出處欄位） ═════════════════
  const vm = require('vm'), fs = require('fs');
  const qctx = {}; vm.createContext(qctx);
  vm.runInContext(fs.readFileSync(path.join(ROOT, '_client/quote-pricebook.js'), 'utf8'), qctx);
  const QPB = qctx.QPB;
  const QA = require(path.join(ROOT, 'lib/quoteApproval.js'));
  env = mkEnv();
  await env.call('admin1', 'PUT', A, {}, putBody([row('PM 顧問經理', 9000, 6500), row('SD 顧問', 7000, 5000)], null));
  const list = QPB.sanitizeList((await env.call('own1', 'GET', U)).j.items);
  const picked = QPB.buildItems(list, [{ id: list[0].id, qty: 3 }, { id: list[1].id, qty: 1.5 }]);
  const mkBody = (items, co) => ({ company: co || 'PbCo', projectName: 'P', products: [], items, discountType: 'none' });
  const c1 = await env.call('own1', 'POST', '/api/quotations', {}, mkBody(picked));
  const saved = env.data.quotations.find((x) => x.id === (c1.j && (c1.j.id || (c1.j.quotation && c1.j.quotation.id))));
  t('6.1 選取產生的品項可以直接存成報價單（POST /api/quotations 200）', c1.s === 200 || c1.s === 201, JSON.stringify(c1).slice(0, 200));
  t('6.2 存下來的品項：desc／unit 人天／qty／unitPrice／cat consult，加上伺服器配發的 lid 與成本 0；沒有任何牌價簿出處欄位', !!saved && saved.items.length === 2 && saved.items.every((i) => eq(Object.keys(i).sort(), ['cat', 'cost', 'desc', 'lid', 'qty', 'unit', 'unitPrice'])) && saved.items[0].desc === 'PM 顧問經理' && saved.items[0].unit === '人天' && saved.items[0].qty === 3 && saved.items[0].unitPrice === 9000 && saved.items[0].cat === 'consult' && saved.items[1].qty === 1.5, JSON.stringify(saved && saved.items));
  // 手動輸入同樣內容的品項：存下來逐欄相同（除 lid）→ 簽核雜湊也相同（牌價簿不改雜湊）
  const c2 = await env.call('own1', 'POST', '/api/quotations', {}, mkBody(picked.map((i) => Object.assign({}, i))));
  const saved2 = env.data.quotations.filter((x) => x.company === 'PbCo')[1];
  const strip = (items) => items.map((i) => { const o = Object.assign({}, i); delete o.lid; return o; });
  t('6.3 手動輸入同樣內容 → 存下來的品項逐欄相同（除 lid），且 contentHash（排除 id 差異後）相同', !!saved2 && eq(strip(saved.items), strip(saved2.items)) && (() => { const a = JSON.parse(JSON.stringify(saved)), b = JSON.parse(JSON.stringify(saved2)); [a, b].forEach((q) => { q.id = 'X'; q.items.forEach((i, k) => { i.lid = 'L' + k; }); }); return QA.contentHash(a) === QA.contentHash(b); })());
  // 伺服器不強制牌價：改過單價照存
  const edited = picked.map((i, k) => Object.assign({}, i, k === 0 ? { unitPrice: 1234 } : {}));
  await env.call('own1', 'POST', '/api/quotations', {}, mkBody(edited));
  const saved3 = env.data.quotations.filter((x) => x.company === 'PbCo')[2];
  t('6.4 伺服器不強制牌價：業務把單價改成 1234 照存（牌價簿只是建議值）', !!saved3 && saved3.items[0].unitPrice === 1234);
  // 之後牌價簿改價，已存的單不變
  const snapSaved = JSON.stringify(env.data.quotations);
  const cur = (await env.call('admin1', 'GET', A)).j;
  await env.call('admin1', 'PUT', A, {}, putBody(cur.items.map((x) => Object.assign({}, x, { price: x.price + 1000, cost: x.cost + 1000, name: x.name + '改' })), cur.updatedAt));
  t('6.5 牌價簿之後改價／改名，先前已存的報價單（含品項單價）完全不變', JSON.stringify(env.data.quotations) === snapSaved);
}

run().then(() => {
  let pass = 0, fail = 0;
  res.forEach(([n, ok, x]) => {
    console.log((ok ? 'PASS ' : 'FAIL ') + n + (x && !ok ? '  <- ' + x : ''));
    ok ? pass++ : fail++;
  });
  console.log('\n報價牌價簿（伺服器）檢查：PASS ' + pass + ' / FAIL ' + fail);
  process.exit(fail ? 1 : 0);
}).catch((e) => { res.filter((r) => !r[1]).forEach((r) => console.log('FAIL ' + r[0] + (r[2] ? '  <- ' + r[2] : ''))); console.error('例外（後面的檢查沒跑到）', e.stack); process.exit(2); });
