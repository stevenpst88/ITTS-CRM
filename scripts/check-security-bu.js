#!/usr/bin/env node
/**
 * 第五個 BU「資安」（代碼就是中文 "資安"）檢查：
 *   1) 各處寫死的 BU 清單同值同序（server.js VALID_BUS／BU 對比 API、lib/quoteRoutes.js ALL_BUS、lib/quotePricebook.js BUS、
 *      前端 QPB／QCL／admin.html／admin-users-view.js／app.js、productCatalog 種子）；html 的 <option>、核取方塊、樣式
 *   2) validateBuInput（從 server.js 原始碼抽出來在 vm 內執行）：資安 通過；多 BU 去重；無效值（"資安 "、"SEC"）被濾掉；錯誤訊息列出五個 BU
 *   3) 秘書可視範圍（lib/quoteRoutes.js inViewScope，經 GET /api/quotations）：
 *        掛滿「舊四個 BU」的秘書維持看全部（不需加掛 資安）、四個＋資安 也看全部；只掛資安 → 只看資安業務的單；
 *        只掛舊四個中的部分（如 ERP+ITS+MDM）→ 不會因為新增 資安 而多看到 資安 的單；沒有 BU → 什麼都看不到
 *   4) 簽核流程設定：GET 帶 資安（預設 hasMgr2Step:false）；PUT 可開關 資安；舊設定檔（沒有 資安 鍵）讀取時補 false
 *   5) 簽核關卡的 BU 判斷（buildActualTiers）：bu 為陣列 ['資安']／字串／多 BU（取第一個）都對得上設定
 * 用法：node scripts/check-security-bu.js（不需要伺服器；不碰 data.json／auth.json／audit.log.json）
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ROOT = path.join(__dirname, '..');
const read = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8');
const PB = require(path.join(ROOT, 'lib/quotePricebook.js'));
const PC = require(path.join(ROOT, 'lib/productCatalog.js'));
const registerQuoteRoutes = require(path.join(ROOT, 'lib/quoteRoutes.js'));
const V = require(path.join(ROOT, '_client/admin-users-view.js'));

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = JSON.stringify;
const BUS5 = ['ERP', 'ITS', 'MDM', 'CRM', '資安'];
// 從原始碼抓 "xxx = ['ERP', ...]" 的陣列字面值
const arrOf = (src, re) => { const m = re.exec(src); return m ? J(JSON.parse(m[1].replace(/'/g, '"'))) : null; };

(async () => {
  // ═════════ 1) 清單同步 ═════════
  const srv = read('server.js'), qr = read('lib/quoteRoutes.js'), adm = read('_client/admin.html'), app = read('_client/app.js'), idx = read('_client/index.html'), css = read('_client/style.css');
  t('1.1 server.js VALID_BUS＝五個 BU', arrOf(srv, /const VALID_BUS = (\[[^\]]*\])/) === J(BUS5));
  t('1.2 server.js BU 對比 API 的 BUS＝五個 BU（/api/exec/bu-comparison）', arrOf(srv, /app\.get\('\/api\/exec\/bu-comparison'[\s\S]*?const BUS = (\[[^\]]*\])/) === J(BUS5));
  t('1.3 lib/quoteRoutes.js ALL_BUS＝五個 BU；LEGACY_ALL_BUS＝舊四個（秘書「掛滿＝看全部」的基準）', arrOf(qr, /const ALL_BUS = (\[[^\]]*\])/) === J(BUS5) && arrOf(qr, /const LEGACY_ALL_BUS = (\[[^\]]*\])/) === J(BUS5.slice(0, 4)) && /LEGACY_ALL_BUS\.every\(/.test(qr));
  t('1.4 牌價簿 BUS 與 quoteRoutes／QPB／QCL／後台 PB_BUS 同值同序；MAX_TOTAL＝300', J(PB.BUS) === J(BUS5) && PB.MAX_TOTAL === 300 && arrOf(read('_client/quote-pricebook.js'), /var BUS = (\[[^\]]*\])/) === J(BUS5) && arrOf(read('_client/quote-costlines.js'), /const PB_BUS = (\[[^\]]*\])/) === J(BUS5) && arrOf(adm, /const PB_BUS = (\[[^\]]*\])/) === J(BUS5));
  t('1.5 後台其他清單：簽核流程 AF_BUS、商品目錄 PC_BUS、帳號管理 BU_LIST 都是五個 BU', arrOf(adm, /const AF_BUS = (\[[^\]]*\])/) === J(BUS5) && arrOf(adm, /const PC_BUS = (\[[^\]]*\])/) === J(BUS5) && J(V.BU_LIST) === J(BUS5));
  t('1.6 前端 app.js：setupBuSelector 的 ALL_BUS、BU 業績總覽 BUS 都含 資安', arrOf(app, /const ALL_BUS = (\[[^\]]*\])/) === J(BUS5) && arrOf(app, /function renderBuBreakdown[\s\S]*?const BUS = (\[[^\]]*\])/) === J(BUS5));
  t('1.7 商品目錄種子有 資安 鍵（空陣列）；deriveDropdowns 對沒有 資安 鍵的舊目錄不丟錯', Array.isArray(PC.SEED_PRODUCT_CATALOG.bus['資安']) && PC.SEED_PRODUCT_CATALOG.bus['資安'].length === 0 && !!PC.deriveDropdowns({ bus: { ERP: [] } }) && PC.deriveDropdowns(PC.SEED_PRODUCT_CATALOG).OPP_PRODUCTS['資安'] !== undefined);
  const optCount = (html, v) => (html.match(new RegExp('<option value="' + v + '">' + v + '</option>', 'g')) || []).length;
  t('1.8 index.html：4 個 BU 下拉（預測篩選、cOppCategory、oppCategory、oppEditCategory）都有 資安 選項', optCount(idx, '資安') === 4 && optCount(idx, 'MDM') === 4);
  t('1.9 admin.html 使用者表單有 資安 核取方塊（class bu-cb）', /class="bu-cb" value="資安"/.test(adm) && (adm.match(/class="bu-cb" value="/g) || []).length === 5);
  t('1.10 客戶類別按鈕：兩個對話框都有 pl-sec／data-pl="資安"；style.css 有 .pl-sec；BU 卡片格線 5 欄', (app.match(/pl-btn pl-sec" data-pl="資安"/g) || []).length === 2 && /\.pl-sec \{/.test(css) && /\.bu-breakdown-grid \{ display: grid; grid-template-columns: repeat\(5, 1fr\)/.test(css));
  t('1.11 資安 顏色 #e11d48 與既有四色不同（四處淺色表＋一處深色表）', ['#1a73e8', '#0a8a4a', '#e37400', '#7c3aed', '#d97706'].every((c) => c !== '#e11d48') && (app.match(/#e11d48/g) || []).length >= 3 && (adm.match(/#e11d48/g) || []).length >= 3 && /'資安':'#fb7185'/.test(app));

  // ═════════ 2) validateBuInput（原始碼抽出，在 vm 內執行）═════════
  const fnSrc = /function validateBuInput\(role, buInput\) \{[\s\S]*?\n\}\n/.exec(srv);
  const ctx = vm.createContext({ VALID_BUS: BUS5.slice(), CROSS_BU_ROLES: ['admin', 'executive', 'accounting_manager', 'finance_manager'], GROUP_SCOPED_ROLES: ['tecopm', 'groupsales'] });
  vm.runInContext(fnSrc[0] + '\nthis.v = validateBuInput;', ctx);
  const vb = (role, bu) => ctx.v(role, bu);
  t('2.1 validateBuInput：user + ["資安"] 通過 → bu=["資安"]；字串 "資安" 也通過', J(vb('user', ['資安'])) === J({ bu: ['資安'] }) && J(vb('user', '資安')) === J({ bu: ['資安'] }));
  t('2.2 多 BU 含 資安、重複去重、保留輸入順序', J(vb('user', ['ERP', '資安', '資安', 'ITS']).bu) === J(['ERP', '資安', 'ITS']));
  t('2.3 無效值被濾掉（"資安 "、"SEC"、"資安部"、小寫 erp）；全無效 → 錯誤，訊息列出五個 BU（含 資安）', J(vb('user', ['資安 ', 'SEC', '資安部', 'erp']).error ? 'err' : 'ok') === '"err"' && /資安/.test(vb('user', []).error) && /ERP \/ ITS \/ MDM \/ CRM \/ 資安/.test(vb('manager1', ['x']).error));
  t('2.4 跨 BU 角色／集團角色仍一律 bu=null，不受 資安 影響', vb('admin', ['資安']).bu === null && vb('executive', []).bu === null && vb('tecopm', ['資安']).bu === null && vb('groupsales', ['資安']).bu === null);

  // ═════════ 3) 秘書可視範圍 + 4) 流程設定 + 5) buActualTiers ═════════
  const USERS = [
    { username: 'adm', role: 'admin', active: true, displayName: 'Adm' },
    { username: 'o_erp', role: 'user', active: true, displayName: 'OErp', bu: ['ERP'] },
    { username: 'o_its', role: 'user', active: true, displayName: 'OIts', bu: ['ITS'] },
    { username: 'o_sec', role: 'user', active: true, displayName: 'OSec', bu: ['資安'] },
    { username: 'o_multi', role: 'user', active: true, displayName: 'OMulti', bu: ['資安', 'ERP'], supervisor: 'm2x' },
    { username: 'o_str', role: 'user', active: true, displayName: 'OStr', bu: '資安', supervisor: 'm2x' },
    { username: 'm2x', role: 'manager2', active: true, displayName: 'M2', bu: ['資安'] },
    { username: 's_all4', role: 'secretary', active: true, displayName: 'S4', bu: ['ERP', 'ITS', 'MDM', 'CRM'] },
    { username: 's_all5', role: 'secretary', active: true, displayName: 'S5', bu: ['ERP', 'ITS', 'MDM', 'CRM', '資安'] },
    { username: 's_sec', role: 'secretary', active: true, displayName: 'SSec', bu: ['資安'] },
    { username: 's_part', role: 'secretary', active: true, displayName: 'SPart', bu: ['ERP', 'ITS', 'MDM'] },
    { username: 's_none', role: 'secretary', active: true, displayName: 'SNone', bu: [] },
    { username: 's_its_sec', role: 'secretary', active: true, displayName: 'SIS', bu: ['ITS', '資安'] },
  ];
  const mkQ = (id, owner) => ({ id, quoteNo: 'QU-' + id, owner, company: 'TestCo', projectName: 'Proj', status: 'draft', products: [], items: [{ lid: 'i1', desc: 'PM', unit: '人天', qty: 3, unitPrice: 8000, cost: 0 }], approval: null, discountType: 'none', discountValue: 0, createdAt: '2026-10-0' + id.slice(1) + 'T00:00:00Z' });
  const env = { data: { quotations: [mkQ('q1', 'o_erp'), mkQ('q2', 'o_its'), mkQ('q3', 'o_sec'), mkQ('q4', 'o_multi')], quoteApproval: { roster: { gm: [], chairman: [], boardProxy: [], costProviders: [], sealManagers: [] }, productClasses: {} } }, auth: { users: JSON.parse(J(USERS)) }, logs: [], seq: 0 };
  const routes = {};
  const fakeApp = {}; ['get', 'post', 'put', 'delete'].forEach((m) => { fakeApp[m] = (p, ...h) => { routes[m.toUpperCase() + ' ' + p] = h; }; });
  const pass = (rq, rs, next) => next();
  const deps = {
    db: { load: () => env.data, save: () => {}, flush: async () => {} }, loadAuth: () => env.auth, saveAuth: () => {},
    requireAuth: pass, requireAdmin: pass, writeLog: (...a) => env.logs.push(a), pushNotification: () => {},
    getViewableOwners: (req) => [req.session.user.username],
    sanitizeStr: (s, n) => String(s == null ? '' : s).trim().slice(0, n || 200), genQuoteNo: () => 'QT-NEW', taipeiToday: () => '2026-10-08', resolveIssuer: () => ({}),
    buildQuoteWorkbook: async () => Buffer.from(''), buildQuotePnlExcel: async () => Buffer.from(''), QUOTE_TEMPLATE: '',
    uuidv4: () => 'u' + (++env.seq).toString().padStart(8, '0') + 'xxxx', normalizeBu: (b) => (Array.isArray(b) ? b.filter(Boolean) : (typeof b === 'string' && b.trim() ? [b.trim()] : [])),
    getUserFeatures: () => ['quotations'],
  };
  registerQuoteRoutes(fakeApp, deps);
  const call = (user, method, p, body) => new Promise((resolve, reject) => {
    const h = [].concat(...routes[method + ' ' + p]);
    const role = (env.auth.users.find((u) => u.username === user) || {}).role;
    const req = { session: { user: { username: user, role } }, params: {}, query: {}, headers: {}, body: body === undefined ? {} : JSON.parse(J(body)) };
    const rs = { headersSent: false, status(c) { this._s = c; return this; }, json(j) { resolve({ s: this._s || 200, j }); return this; }, send(b) { resolve({ s: this._s || 200, b }); return this; }, setHeader() {}, set() {} };
    let i = 0; const next = () => { const f = h[i++]; if (f) { try { const x = f(req, rs, next); if (x && x.catch) x.catch(reject); } catch (e) { reject(e); } } }; next();
  });
  const seen = async (u) => { const r = await call(u, 'GET', '/api/quotations'); return r.s === 200 ? r.j.map((x) => x.id).sort() : 'HTTP' + r.s; };
  const ALLQ = ['q1', 'q2', 'q3', 'q4'];
  t('3.1 舊行為不變：掛滿 ERP/ITS/MDM/CRM 的秘書（沒有 資安）仍看到全部報價單，包含 資安 業務的單', J(await seen('s_all4')) === J(ALLQ), J(await seen('s_all4')));
  t('3.2 掛滿五個 BU 的秘書看全部', J(await seen('s_all5')) === J(ALLQ));
  t('3.3 只掛 資安 的秘書：只看到 資安 業務（含 資安+ERP 多 BU）發起的單，看不到純 ERP／ITS 的單', J(await seen('s_sec')) === J(['q3', 'q4']), J(await seen('s_sec')));
  t('3.4 只掛舊四個的一部分（ERP+ITS+MDM）：不會因為新增 資安 而變成看全部；看 ERP／ITS 與 資安+ERP 多 BU 的單，看不到純 資安 的單', J(await seen('s_part')) === J(['q1', 'q2', 'q4']), J(await seen('s_part')));
  t('3.5 沒有 BU 的秘書看不到任何單', J(await seen('s_none')) === J([]));
  t('3.6 ITS+資安 的秘書：看 ITS 與 資安 的單（含多 BU）、不看純 ERP', J(await seen('s_its_sec')) === J(['q2', 'q3', 'q4']), J(await seen('s_its_sec')));
  t('3.7 管理員看全部；各 BU 業務只看自己的（不受 資安 影響）', J(await seen('adm')) === J(ALLQ) && J(await seen('o_sec')) === J(['q3']) && J(await seen('o_erp')) === J(['q1']));

  const g0 = await call('adm', 'GET', '/api/admin/quote-approval/flow-config');
  t('4.1 流程設定 GET：五個 BU 都有 buSettings，資安 預設 hasMgr2Step:false（舊設定檔沒有 資安 鍵也補上）', g0.s === 200 && J(Object.keys(g0.j.buSettings)) === J(BUS5) && g0.j.buSettings['資安'].hasMgr2Step === false);
  const p1 = await call('adm', 'PUT', '/api/admin/quote-approval/flow-config', { buSettings: { '資安': { hasMgr2Step: true }, ERP: { hasMgr2Step: false } }, exceptions: [] });
  const g1 = await call('adm', 'GET', '/api/admin/quote-approval/flow-config');
  t('4.2 PUT 開啟 資安 的二級主管關卡 → GET 讀回 true，其他 BU 仍 false；稽核文字含 資安', p1.s === 200 && g1.j.buSettings['資安'].hasMgr2Step === true && ['ERP', 'ITS', 'MDM', 'CRM'].every((b) => g1.j.buSettings[b].hasMgr2Step === false) && env.logs.some((l) => l[0] === 'SAVE_APPROVAL_FLOW_CONFIG' && /資安/.test(l[3])), J(g1.j));
  // 5) 簽核關卡的 BU 判斷：送簽路徑需要完整核決表與成本資料，這裡以原始碼層級確認
  t('5.1 buildActualTiers 以 normalizeBu(...)[0] 取主要 BU（陣列／字串／多 BU 都對得上 buSettings 鍵；不再把陣列直接當鍵）', /const ownerBu = normalizeBu\(\(users\[ownerUsername\] \|\| \{\}\)\.bu\)\[0\] \|\| 'ERP'/.test(qr) && /\(Array\.isArray\(u\.bu\) \? u\.bu\[0\] : u\.bu\) === bu/.test(adm));

  // ─── 輸出 ───
  let pass_ = 0, fail_ = 0;
  for (const [name, ok, extra] of res) { if (ok) pass_++; else fail_++; console.log((ok ? 'PASS ' : 'FAIL ') + name + (ok ? '' : (extra ? '  <- ' + extra.slice(0, 300) : ''))); }
  console.log('資安 BU 檢查：PASS ' + pass_ + ' / FAIL ' + fail_);
  process.exit(fail_ ? 1 : 0);
})().catch((e) => { console.error('檢查腳本例外', e); process.exit(2); });
