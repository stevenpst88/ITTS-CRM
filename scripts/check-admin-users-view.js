#!/usr/bin/env node
/**
 * 後台「帳號管理」檢視邏輯（_client/admin-users-view.js → AdminUsersView）純函式＋admin.html 接線／跳脫的靜態檢查。
 * 用法：node scripts/check-admin-users-view.js（沒有 DOM、不需要伺服器；真實瀏覽器行為另以無頭瀏覽器實跑）。
 *   1) 分割規則：多 BU 出現在每個分頁；空 bu／集團角色／pool／壞資料 → 全公司／集團；怪值（小寫、空白、null、重複、非陣列、非字串元素）不丟錯
 *   2) 篩選與計數：搜尋（NFKC、大小寫、前後空白、不比對 email、正規式字元當一般字元、超長輸入）、角色／狀態、計數隨篩選變動
 *   3) 角色順序與分組、兩種排序（穩定、中文 zh-TW、不改輸入）、sessionStorage 值驗證、收合狀態輔助函式、pickTab、buildView
 *   4) 靜態紀律：admin.html 新渲染路徑一律 escH、分頁 ARIA、mailAfterUserTable 仍被呼叫、pool 標籤在且角色下拉沒有 pool、
 *      script 標籤在使用之前、伺服器 ?v= 注入、沒有 eval／new Function／document.write、深色樣式限定在 #sec-users
 * 變異測試：AUV_SRC／AUV_ADMIN／AUV_SERVER 環境變數可指向被破壞的副本（見報告）。
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const ROOT = path.join(__dirname, '..');
const LIB = process.env.AUV_SRC || path.join(ROOT, '_client/admin-users-view.js');
const ADMIN = process.env.AUV_ADMIN || path.join(ROOT, '_client/admin.html');
const SERVER = process.env.AUV_SERVER || path.join(ROOT, 'server.js');

const res = [];
const t = (name, ok, extra) => res.push([name, !!ok, extra === undefined ? '' : String(extra)]);
const J = (x) => JSON.stringify(x);
const read = (p) => fs.readFileSync(p, 'utf8').replace(/\r\n/g, '\n');

const libSrc = read(LIB);
const adminSrc = read(ADMIN);
const serverSrc = read(SERVER);
const V = require(LIB);

// 瀏覽器路徑：用 vm 載入，確認掛到 window.AdminUsersView（沒有 module）
const sandbox = { self: {} };
sandbox.window = sandbox.self;
vm.createContext(sandbox);
vm.runInContext(libSrc, sandbox);

const mk = (username, role, bu, extra) => Object.assign({ username, displayName: username, nickname: '', email: '', role, bu, active: true }, extra || {});
const names = (list) => list.map((u) => u.username);
const noThrow = (fn) => { try { fn(); return true; } catch (e) { return false; } };

// ═════════ 0) 載入 ═════════
t('0.1 瀏覽器載入掛 window.AdminUsersView；Node 載入同一份', sandbox.self.AdminUsersView && typeof sandbox.self.AdminUsersView.buildView === 'function' && typeof V.buildView === 'function');
t('0.2 五個分頁固定且順序是 ERP、ITS、MDM、CRM、全公司／集團（沒有「全部」分頁）', J(V.TABS.map((x) => x.label)) === J(['ERP', 'ITS', 'MDM', 'CRM', '全公司／集團']) && V.TAB_IDS.length === 5);

// ═════════ 1) 分割規則 ═════════
{
  const multi = mk('multi', 'user', ['ERP', 'CRM']);
  const all4 = mk('all4', 'manager1', ['ERP', 'ITS', 'MDM', 'CRM']);
  const one = mk('one', 'user', ['ITS']);
  const adm = mk('adm', 'admin', []);
  const exe = mk('exe', 'executive', []);
  const acc = mk('acc', 'accounting_manager', []);
  const fin = mk('fin', 'finance_manager', []);
  const tpm = mk('tpm', 'tecopm', []);
  const gs = mk('gs', 'groupsales', []);
  const pool = mk('_pool', 'pool', []);
  const p = V.partition([multi, all4, one, adm, exe, acc, fin, tpm, gs, pool]);
  t('1.1 多 BU 帳號出現在每個所屬的 BU 分頁（ERP+CRM → 兩個分頁都有；四個 BU → 四個分頁都有）', p.ERP.includes(multi) && p.CRM.includes(multi) && !p.ITS.includes(multi) && !p.MDM.includes(multi) && ['ERP', 'ITS', 'MDM', 'CRM'].every((b) => p[b].includes(all4)));
  t('1.2 單一 BU 只在該分頁；各分頁數字因多 BU 而重疊（總和大於帳號數）', p.ITS.includes(one) && !p.ERP.includes(one) && Object.values(p).reduce((n, l) => n + l.length, 0) > 10);
  t('1.3 空 bu 的 admin／executive／accounting_manager／finance_manager → 只在「全公司／集團」', [adm, exe, acc, fin].every((u) => p.CORP.includes(u) && ['ERP', 'ITS', 'MDM', 'CRM'].every((b) => !p[b].includes(u))));
  t('1.4 tecopm／groupsales／pool → 全公司／集團', [tpm, gs, pool].every((u) => p.CORP.includes(u)));
  t('1.5 集團角色即使資料裡殘留 bu，仍只放全公司／集團（畫面上它們顯示的是「集團限制」而不是 BU）', J(V.tabsOf(mk('x', 'tecopm', ['ERP']))) === J(['CORP']) && J(V.tabsOf(mk('y', 'groupsales', ['ITS', 'CRM']))) === J(['CORP']));
  t('1.6 每個帳號在 CORP 以外的分頁數＝有效 BU 數；CORP 只收沒有有效 BU 的', V.tabsOf(multi).length === 2 && V.tabsOf(one).length === 1 && V.tabsOf(adm).length === 1);
}
{
  const odd = [
    mk('lower', 'user', ['erp']), mk('spaced', 'user', [' Its ']), mk('dup', 'user', ['ERP', 'ERP', 'erp']), mk('nul', 'user', null), mk('undef', 'user', undefined),
    mk('str', 'user', 'MDM'), mk('numarr', 'user', [1, null, {}, ['ERP']]), mk('junk', 'user', ['XYZ', '']), mk('obj', 'user', { 0: 'ERP' }), mk('num', 'user', 5),
    mk('mixed', 'user', ['ERP', null, 'bad', 'crm'])
  ];
  const p = V.partition(odd);
  t('1.7 小寫／前後空白的 BU 正規化後歸入對應分頁（erp→ERP、" Its "→ITS）', p.ERP.some((u) => u.username === 'lower') && p.ITS.some((u) => u.username === 'spaced'));
  t('1.8 重複的 BU 只算一次（["ERP","ERP","erp"] 在 ERP 分頁只出現一次）', p.ERP.filter((u) => u.username === 'dup').length === 1);
  t('1.9 null／undefined／非陣列物件／數字／無效值 → 全公司／集團，不丟錯', ['nul', 'undef', 'numarr', 'junk', 'obj', 'num'].every((n) => p.CORP.some((u) => u.username === n)) && noThrow(() => V.partition(odd)));
  t('1.10 單一字串 bu（舊格式）視為一個 BU；混合陣列只留有效值', p.MDM.some((u) => u.username === 'str') && p.ERP.some((u) => u.username === 'mixed') && p.CRM.some((u) => u.username === 'mixed') && !p.CORP.some((u) => u.username === 'mixed'));
  t('1.11 壞輸入不丟錯：null／非物件元素／缺欄位／非陣列的使用者清單', noThrow(() => { V.partition([null, undefined, 5, 'x', [], {}, { role: 'user' }]); V.partition(null); V.partition('abc'); V.buildView(null, null); V.buildView([null, 3], { q: 'a' }); }));
  t('1.12 沒有 role 的帳號也能分割與分組（role 為空字串群組）', noThrow(() => V.groupByRole([{ username: 'a', bu: ['ERP'] }, { username: 'b', role: null }])));
}

// ═════════ 2) 搜尋、篩選、計數 ═════════
{
  const users = [
    mk('alice', 'user', ['ERP'], { displayName: 'Alice Wang', nickname: 'Ally' }),
    mk('bob', 'user', ['ERP', 'ITS'], { displayName: '王小明', nickname: '' }),
    mk('carol', 'manager1', ['ITS'], { displayName: 'Carol', nickname: '小卡', active: false }),
    mk('dave', 'admin', [], { displayName: 'Dave', email: 'secret.person@example.com' }),
    mk('fw', 'user', ['CRM'], { displayName: 'ＡＢＣ全形', nickname: '' }),
    mk('meta', 'user', ['MDM'], { displayName: 'a.b*c(d)[e]+?', nickname: '' })
  ];
  const f = (q, extra) => names(V.applyFilters(users, Object.assign({ q }, extra || {})));
  t('2.1 搜尋比對帳號／顯示名稱／暱稱（子字串）', J(f('alice')) === J(['alice']) && J(f('王小')) === J(['bob']) && J(f('小卡')) === J(['carol']) && J(f('ally')) === J(['alice']));
  t('2.2 不分大小寫；前後空白忽略；內部連續空白壓成一個', J(f('ALICE')) === J(['alice']) && J(f('  alice  ')) === J(['alice']) && J(f('Alice   Wang')) === J(['alice']) && J(f('\t\n')) === J(names(users)));
  t('2.3 NFKC：全形英文可用半形搜尋（abc→ＡＢＣ全形），半形帳號可用全形搜尋（ＡＬＩＣＥ→alice）', J(f('abc')) === J(['fw']) && J(f('ＡＬＩＣＥ')) === J(['alice']));
  t('2.4 不搜尋 email：用 email 的片段、完整位址、網域都搜不到', J(f('secret.person')) === J([]) && J(f('secret.person@example.com')) === J([]) && J(f('example.com')) === J([]) && J(f('@')) === J([]));
  t('2.5 正規式特殊字元當一般字元：「.」不是萬用字元、「*」「(」「[」「+」「?」逐字比對，不丟錯', J(f('a.b*c(d)[e]+?')) === J(['meta']) && J(f('.*')) === J([]) && J(f('(')) === J(['meta']) && J(f('[')) === J(['meta']) && noThrow(() => f('\\')) && noThrow(() => f('((((')) && J(f('a.c')) === J([]));
  t('2.6 超長輸入（100 萬字元）快速且不丟錯、不符合任何人', (() => { const s = Date.now(); const r = f('a'.repeat(1000000)); return r.length === 0 && Date.now() - s < 2000; })());
  t('2.7 搜尋值是 null／undefined／數字／物件時不丟錯（當空字串或字串處理）', noThrow(() => { f(null); f(undefined); f(5); f({}); f([]); }) && J(f(null)) === J(names(users)));
  t('2.8 角色篩選：只留該角色；不存在的角色 → 空', J(f('', { role: 'admin' })) === J(['dave']) && J(f('', { role: 'nope' })) === J([]));
  t('2.9 狀態篩選：active／inactive；其他值＝不篩', J(f('', { status: 'inactive' })) === J(['carol']) && f('', { status: 'active' }).length === 5 && f('', { status: 'weird' }).length === 6);
  t('2.10 條件 AND：搜尋＋角色＋狀態同時成立', J(f('小', { role: 'manager1', status: 'inactive' })) === J(['carol']) && J(f('小', { role: 'manager1', status: 'active' })) === J([]));
  t('2.11 active 欄位缺漏視為停用（與表格上的「停用」徽章一致）', V.isActive({ active: undefined }) === false && V.isActive({ active: true }) === true && f('', { status: 'inactive' }).includes('carol'));
  const v0 = V.buildView(users, {});
  t('2.12 計數：無篩選時各分頁＝該分頁人數（多 BU 重複計）', v0.tabs.find((x) => x.id === 'ERP').count === 2 && v0.tabs.find((x) => x.id === 'ITS').count === 2 && v0.tabs.find((x) => x.id === 'CORP').count === 1 && v0.total === 6 && v0.matched === 6);
  const v1 = V.buildView(users, { q: '王' });
  t('2.13 計數隨篩選變動：搜尋「王」→ ERP 1、ITS 1、其餘 0；matched 是不重複人數(2)', v1.tabs.find((x) => x.id === 'ERP').count === 1 && v1.tabs.find((x) => x.id === 'ITS').count === 1 && v1.tabs.find((x) => x.id === 'CORP').count === 0 && v1.matched === 1 && v1.total === 6, J(v1.tabs));
  const v2 = V.buildView(users, { status: 'inactive' });
  t('2.14 狀態篩選也改計數；filtersOn 為真；無篩選時為假；只有空白的搜尋不算篩選', v2.tabs.find((x) => x.id === 'ITS').count === 1 && v2.filtersOn === true && v0.filtersOn === false && V.buildView(users, { q: '   ' }).filtersOn === false);
  t('2.15 空結果時 others 列出其他有帳號的分頁（供空狀態提示）', (() => { const v = V.buildView(users, { q: '王', tab: 'CRM' }); return v.users.length === 0 && v.others.some((o) => o.id === 'ERP') && v.others.every((o) => o.count > 0 && o.id !== 'CRM'); })());
}

// ═════════ 3) 角色順序、分組、排序 ═════════
{
  const roles = ['pool', 'zzz_unknown', 'tecopm', 'secretary', 'consult_manager_north', 'user', 'manager2', 'admin', 'aaa_unknown', 'groupsales', 'marketing', 'manager1', 'executive', 'finance_manager', 'accounting_manager', 'consult_manager_south'];
  const list = roles.map((r, i) => mk('u' + i, r, ['ERP']));
  const g = V.groupByRole(list);
  t('3.1 角色組序：admin, executive, accounting_manager, finance_manager, manager1, manager2, user, marketing, consult_*, secretary, groupsales, tecopm', J(g.map((x) => x.role).slice(0, 13)) === J(['admin', 'executive', 'accounting_manager', 'finance_manager', 'manager1', 'manager2', 'user', 'marketing', 'consult_manager_south', 'consult_manager_north', 'secretary', 'groupsales', 'tecopm']));
  t('3.2 未登記角色排在已登記之後（依原始值字母序）、pool 永遠最後', J(g.slice(13).map((x) => x.role)) === J(['aaa_unknown', 'zzz_unknown', 'pool']));
  t('3.3 ROLE_ORDER 涵蓋伺服器 KNOWN_ROLES 的每個角色（沒有漏掉的角色掉進「未登記」）', (() => { const m = /const KNOWN_ROLES = (\[[^\]]*\])/.exec(serverSrc); const known = JSON.parse(m[1].replace(/'/g, '"')); return known.every((r) => V.ROLE_ORDER.includes(r)) && known.every((r) => V.ROLE_LABELS[r]); })());
  t('3.4 組名：已登記用中文標籤、未登記用原始值、pool 有中文標籤', g[0].label === '系統管理員' && g.find((x) => x.role === 'zzz_unknown').label === 'zzz_unknown' && g.find((x) => x.role === 'pool').label === '客戶池（系統帳號）' && V.roleLabel('pool') !== 'pool');
  t('3.5 roleLabel 對 __proto__／toString／constructor 這類字串不誤判（回原始值）', V.roleLabel('__proto__') === '__proto__' && V.roleLabel('toString') === 'toString' && V.roleLabel('constructor') === 'constructor' && V.roleRank('constructor') === V.ROLE_ORDER.length + 1);
  t('3.6 角色篩選選項：登記的角色都在、pool 只有資料有才列、未登記但出現的角色會列', (() => { const a = V.roleOptions([mk('a', 'user', ['ERP'])]).map((o) => o.value); const b = V.roleOptions([mk('p', 'pool', []), mk('x', 'weird', [])]).map((o) => o.value); return V.ROLE_ORDER.every((r) => a.includes(r)) && !a.includes('pool') && b.includes('pool') && b.includes('weird') && b[b.length - 1] === 'pool'; })());
}
{
  const users = [
    mk('c1', 'user', ['ERP'], { displayName: '陳大文' }), mk('c2', 'user', ['ERP'], { displayName: '王小明' }), mk('c3', 'user', ['ERP'], { displayName: '李四' }),
    mk('b', 'user', ['ERP'], { displayName: 'Bob' }), mk('a', 'user', ['ERP'], { displayName: 'alice', nickname: 'Zed' }),
    mk('n10', 'user', ['ERP'], { displayName: 'user10' }), mk('n2', 'user', ['ERP'], { displayName: 'user2' }),
    mk('dupA', 'user', ['ERP'], { displayName: 'Same' }), mk('dupB', 'user', ['ERP'], { displayName: 'Same' })
  ];
  const snap = J(users);
  const sorted = V.sortByName(users);
  t('3.7 依名稱排序不改輸入；顯示名稱規則＝暱稱>顯示名稱>帳號（alice 因暱稱 Zed 排到 Zed 的位置）', J(users) === snap && sorted !== users && V.userLabel(users[4]) === 'Zed' && names(sorted).indexOf('a') > names(sorted).indexOf('b'));
  t('3.8 數字依數值排（user2 在 user10 之前）；英文不分大小寫', names(sorted).indexOf('n2') < names(sorted).indexOf('n10'));
  t('3.9 同一批資料不論輸入順序，排序結果完全相同（穩定、可重現）', J(names(V.sortByName(users))) === J(names(V.sortByName(users.slice().reverse()))));
  t('3.10 同名時依帳號再排，結果與輸入順序無關（dupA 永遠在 dupB 前）', (() => { const a = names(V.sortByName(users.slice().reverse())); return a.indexOf('dupA') < a.indexOf('dupB'); })());
  t('3.11 中文依 zh-TW 校對規則排序（ICU 的 zh-TW 預設是筆畫：王4、李7、陳16），不是字碼順序（李、王、陳）', (() => { const c = V.sortByName([mk('x', 'user', ['ERP'], { displayName: '陳' }), mk('y', 'user', ['ERP'], { displayName: '王' }), mk('z', 'user', ['ERP'], { displayName: '李' })]).map((u) => u.displayName).join(''); const want = ['陳', '王', '李'].sort(new Intl.Collator('zh-TW').compare).join(''); return c === want && c !== '李王陳'; })());
  const bv = V.buildView(users, { sort: 'name' });
  t('3.12 sort=name：groups 為 null（平面、不分組）；sort=role：groups 有且 key＝分頁|角色', bv.groups === null && V.buildView(users, { sort: 'role' }).groups[0].key === 'ERP|user' && V.buildView(users, {}).groups !== null);
  t('3.13 無效的 sort 值回預設（依角色）', V.validSort('hax') === 'role' && V.validSort(null) === 'role' && V.buildView(users, { sort: 'hax' }).groups !== null);
  const mixed = [mk('m2', 'manager2', ['ERP'], { displayName: 'Zed' }), mk('m1', 'manager1', ['ERP'], { displayName: 'Amy' }), mk('m3', 'manager1', ['ERP'], { displayName: 'Ben' }), mk('adm', 'admin', ['ERP'])];
  const rv = V.buildView(mixed, { sort: 'role' });
  t('3.14 依角色排序：組序 admin→manager1→manager2；組內依名稱；users 平面清單＝各組依序串接', J(rv.groups.map((g) => g.role)) === J(['admin', 'manager1', 'manager2']) && J(names(rv.groups[1].users)) === J(['m1', 'm3']) && J(names(rv.users)) === J(['adm', 'm1', 'm3', 'm2']));
}

// ═════════ 4) 狀態驗證、收合輔助 ═════════
{
  t('4.1 validTab：只認五個分頁 id，其他（大小寫不同、空字串、null、物件、數字）都回 null', V.validTab('ERP') === 'ERP' && V.validTab('CORP') === 'CORP' && ['erp', '', null, undefined, {}, 5, 'ALL', 'ERP '].every((x) => V.validTab(x) === null));
  t('4.2 validView／validSort／validStatus 壞值回預設（table／role／空）', V.validView('card') === 'card' && V.validView('grid') === 'table' && V.validView(null) === 'table' && V.validSort('name') === 'name' && V.validStatus('active') === 'active' && V.validStatus('x') === '' && V.validStatus(null) === '');
  t('4.3 pickTab：有效值直接用；無效 → 第一個有帳號的分頁；全空 → 第一個', V.pickTab('CRM', { ERP: 3 }) === 'CRM' && V.pickTab(null, { ERP: 0, ITS: 0, MDM: 4, CRM: 1, CORP: 2 }) === 'MDM' && V.pickTab('bad', { CORP: 1 }) === 'CORP' && V.pickTab(null, {}) === 'ERP' && V.pickTab(undefined, null) === 'ERP');
  t('4.4 buildView 預設分頁＝第一個有帳號的分頁（ERP 沒人時是 ITS）；指定有效分頁則尊重（即使是空的）', V.buildView([mk('a', 'user', ['ITS'])], {}).tab === 'ITS' && V.buildView([mk('a', 'user', ['ITS'])], { tab: 'ERP' }).tab === 'ERP' && V.buildView([mk('a', 'user', ['ITS'])], { tab: 'junk' }).tab === 'ITS');
  const k = V.collapseKey('ERP', 'user');
  t('4.5 collapseKey 為「分頁|角色」；validCollapseKey 擋非字串、無分頁前綴、過長', k === 'ERP|user' && V.validCollapseKey(k) && !V.validCollapseKey('XXX|user') && !V.validCollapseKey('|user') && !V.validCollapseKey('ERP') && !V.validCollapseKey(5) && !V.validCollapseKey(null) && !V.validCollapseKey('ERP|' + 'x'.repeat(200)));
  t('4.6 parseCollapsed：壞 JSON／非陣列／null／空字串都回 []；壞元素與重複被濾掉', ['', 'not json', '{"a":1}', 'null', '5', undefined, null, '"str"'].every((x) => J(V.parseCollapsed(x)) === '[]') && J(V.parseCollapsed('["ERP|user",5,"bad","ERP|user","CORP|admin",null]')) === J(['ERP|user', 'CORP|admin']));
  t('4.7 parseCollapsed 有上限（不會被塞一個巨大陣列撐爆）', V.parseCollapsed(J(Array.from({ length: 5000 }, (_, i) => 'ERP|r' + i))).length === 200);
  t('4.8 serializeCollapsed／parseCollapsed 來回一致；接受 Set 與陣列；壞輸入回 "[]"', (() => { const s = new Set(['ERP|user', 'CORP|admin']); return J(V.parseCollapsed(V.serializeCollapsed(s))) === J(['ERP|user', 'CORP|admin']) && V.serializeCollapsed(['ERP|a']) === '["ERP|a"]' && V.serializeCollapsed(null) === '[]' && V.serializeCollapsed(5) === '[]' && V.serializeCollapsed('abc') === '[]'; })());
  t('4.9 toggleCollapsed：加入／移除、不改輸入的 Set', (() => { const s = new Set(['ERP|user']); const a = V.toggleCollapsed(s, 'ITS|user'); const b = V.toggleCollapsed(a, 'ERP|user'); return a.has('ITS|user') && a.has('ERP|user') && !b.has('ERP|user') && s.size === 1 && !s.has('ITS|user') && V.toggleCollapsed(null, 'ERP|x').has('ERP|x'); })());
}

// ═════════ 5) 靜態紀律 ═════════
const uvA = adminSrc.indexOf('const UV = window.AdminUsersView;');
const uvB = adminSrc.indexOf('// ── 事件委派（取代所有 onclick/onchange');
const uvSrc = uvA > 0 && uvB > uvA ? adminSrc.slice(uvA, uvB) : '';
const rowSrc = (() => { const a = uvSrc.indexOf('function uvRowHtml'), b = uvSrc.indexOf('function uvCardHtml'); return a > 0 && b > a ? uvSrc.slice(a, b) : ''; })();
const cardSrc = (() => { const a = uvSrc.indexOf('function uvCardHtml'), b = uvSrc.indexOf('function uvGroupBtn'); return a > 0 && b > a ? uvSrc.slice(a, b) : ''; })();
t('5.0 找得到新渲染區塊（uvRowHtml／uvCardHtml 都有內容）', uvSrc.length > 3000 && rowSrc.length > 500 && cardSrc.length > 500);
t('5.1 新渲染路徑沒有未跳脫的 ${u.xxx} 內插（username／displayName／nickname／role／viewOwnerScope／email／supervisor）', !/\$\{\s*u\.(username|displayName|nickname|role|viewOwnerScope|viewGroupId|email|supervisor)\s*(\|\|[^}]*)?\}/.test(uvSrc), (uvSrc.match(/\$\{\s*u\.[^}]*\}/g) || []).join(' ; '));
t('5.2 username 進 HTML 前一律經 escH（const un = escH(u.username)），data-username 用 un', /const un = escH\(u\.username\)/.test(rowSrc) && /const un = escH\(u\.username\)/.test(cardSrc) && !/data-username="\$\{(?!un\})/.test(rowSrc + cardSrc) && (rowSrc.match(/data-username="\$\{un\}"/g) || []).length === 5 && (cardSrc.match(/data-username="\$\{un\}"/g) || []).length === 5);
t('5.3 displayName／nickname／標籤都經 escH（表格列與卡片）', /escH\(u\.displayName\)/.test(rowSrc) && /escH\(u\.nickname\)/.test(rowSrc) && /escH\(label\)/.test(cardSrc) && /escH\(u\.displayName\)/.test(cardSrc));
t('5.4 群組標題、分頁標籤、角色下拉選項、角色徽章的文字與 data-*／value 都經 escH；class 經 uvCls 過濾', /data-uvgroup="\$\{escH\(g\.key\)\}"/.test(uvSrc) && /\$\{escH\(g\.label\)\}/.test(uvSrc) && /\$\{escH\(t\.label\)\}/.test(uvSrc) && /value="\$\{escH\(o\.value\)\}">\$\{escH\(o\.label\)\}/.test(uvSrc) && /badge-\$\{uvCls\(u\.role\)\}/.test(uvSrc) && /escH\(UV\.roleLabel\(u\.role\)\)/.test(uvSrc));
t('5.5 集團角色那顆「看 …／集團限制」徽章的 viewOwnerScope 也跳脫（原本沒有）', /const scope = escH\(u\.viewOwnerScope/.test(uvSrc));
t('5.6 分頁 ARIA：tablist（aria-label）、tab（aria-selected、tabindex、aria-controls）、tabpanel（aria-labelledby 隨分頁更新）', /id="uvTabs"[^>]*role="tablist"/.test(adminSrc) && /role="tab"/.test(uvSrc) && /setAttribute\('aria-selected'/.test(uvSrc) &&/setAttribute\('tabindex'/.test(uvSrc) && /id="uvPanel"[^>]*role="tabpanel"/.test(adminSrc) && /aria-labelledby', 'uvTab-'/.test(uvSrc));
t('5.7 分頁鍵盤：ArrowLeft／ArrowRight／Home／End，環狀切換', /ArrowRight/.test(uvSrc) && /ArrowLeft/.test(uvSrc) && /'Home'/.test(uvSrc) && /'End'/.test(uvSrc) && /% ids\.length/.test(uvSrc));
t('5.8 群組標題是 <button aria-expanded>；檢視切換是 aria-pressed；摘要列 role=status', /class="uv-gbtn"[^>]*aria-expanded/.test(uvSrc) && /aria-pressed/.test(adminSrc) && /id="uvSummary" role="status"/.test(adminSrc));
t('5.9 loadUsers 仍是 renderUserStat → renderUserTable → mailAfterUserTable（Email 提示列在每次抓資料後更新）', /allUsers = await r\.json\(\);\s*renderUserStat\(\);\s*renderUserTable\(\);\s*if \(typeof mailAfterUserTable === 'function'\) mailAfterUserTable\(\);/.test(adminSrc));
t('5.10 表格列與卡片都呼叫 mailEmailCell(u)（Email 欄／卡片 Email 狀態）', /mailEmailCell\(u\)/.test(rowSrc) && /mailEmailCell\(u\)/.test(cardSrc));
t('5.11 表格仍是 11 欄（表頭 11 個 th；群組列 colspan=11；每列 11 個 td）', (() => { const th = (adminSrc.slice(adminSrc.indexOf('<table class="admin-table">', adminSrc.indexOf('id="uvTableWrap"')), adminSrc.indexOf('<tbody id="userTableBody">')).match(/<th[ >]/g) || []).length; const td = (rowSrc.match(/<td/g) || []).length; return th === 11 && td === 11 && /colspan="11"/.test(uvSrc); })());
t('5.12 卡片動作沿用既有 data-action／data-username／data-field（edit、resetpw、delete、toggle×2），沒有另寫處理函式', ['edit', 'resetpw', 'delete'].every((a) => new RegExp('data-action="' + a + '"').test(cardSrc)) && (cardSrc.match(/data-action="toggle"/g) || []).length === 2 && /data-field="canDownloadContacts"/.test(cardSrc) && /data-field="canSetTargets"/.test(cardSrc) && !/addEventListener\('click'[\s\S]{0,40}openEditUser/.test(uvSrc) && (adminSrc.match(/document\.addEventListener\('click', function\(e\) \{\s*const btn = e\.target\.closest\('\[data-action\]'\)/g) || []).length === 1);
t('5.13 pool 有中文標籤（庫與 SUPERVISOR_ROLE_LABEL），角色下拉 #fieldRole 沒有 pool 選項', /pool: ?'客戶池（系統帳號）'/.test(libSrc) && /SUPERVISOR_ROLE_LABEL = \{[^}]*pool:'客戶池（系統帳號）'/.test(adminSrc) && (() => { const a = adminSrc.indexOf('<select id="fieldRole"'); const b = adminSrc.indexOf('</select>', a); const sel = adminSrc.slice(a, b); return a > 0 && (sel.match(/<option/g) || []).length === 13 && !/pool/.test(sel); })());
t('5.14 新 script 標籤存在，且在使用 AdminUsersView 的內嵌 script 之前', (() => { const a = adminSrc.indexOf('<script src="admin-users-view.js"></script>'); const b = adminSrc.indexOf('const UV = window.AdminUsersView'); const c = adminSrc.indexOf("const API = '/api';"); return a > 0 && c > a && b > c && (adminSrc.match(/src="admin-users-view\.js"/g) || []).length === 1; })());
t('5.15 伺服器對 admin.html 注入 admin-users-view.js?v=（與其他腳本同法）；vercel includeFiles 含 _client/**', /\.replace\(\/src="admin-users-view\\\.js"\/g, `src="admin-users-view\.js\?v=\$\{BUILD_VERSION\}"`\)/.test(serverSrc) && /\{_client,/.test(read(path.join(ROOT, 'vercel.json'))) && /_client,templates[^"]*\/\*\*/.test(read(path.join(ROOT, 'vercel.json'))));
t('5.16 新檔與新渲染區塊沒有 eval／new Function／document.write／setTimeout(字串)', !/\beval\s*\(|new Function|document\.write/.test(libSrc) && !/\beval\s*\(|new Function|document\.write/.test(uvSrc));
t('5.17 搜尋不碰 email：matches 函式沒有 .email；fields 只有 username／displayName／nickname', (() => { const a = libSrc.indexOf('function matches'), b = libSrc.indexOf('function applyFilters'); const m = libSrc.slice(a, b); return a > 0 && b > a && !/email/i.test(m) && /u\.username, u\.displayName, u\.nickname/.test(m); })());
t('5.18 sessionStorage 讀寫都包 try/catch，讀回的值先經 valid*／parse* 驗證', /function uvStoreGet\(k\) \{ try \{ return sessionStorage\.getItem\(k\); \} catch/.test(uvSrc) && /function uvStoreSet\(k, v\) \{ try \{ sessionStorage\.setItem\(k, v\); \} catch/.test(uvSrc) && /UV\.validTab\(uvStoreGet/.test(uvSrc) && /UV\.validView\(uvStoreGet/.test(uvSrc) && /UV\.parseCollapsed\(uvStoreGet/.test(uvSrc) && !/sessionStorage\.(get|set)Item/.test(uvSrc.replace(/function uvStoreGet[^\n]*\n/, '').replace(/function uvStoreSet[^\n]*\n/, '')));
t('5.19 深色樣式全部限定在 #sec-users（新增的 body.dark 規則一行一行檢查）', (() => { const a = adminSrc.indexOf('帳號管理：搜尋／篩選工具列'); const b = adminSrc.indexOf('窄螢幕（手機）：後台原本沒有任何 RWD'); const css = adminSrc.slice(a, b); const dark = css.split('\n').filter((l) => /body\.dark/.test(l)); return a > 0 && b > a && dark.length > 15 && dark.every((l) => /body\.dark #sec-users /.test(l)) && css.split('\n').filter((l) => /^\s*[#.][^{]*\{/.test(l) && !/^\s*(@media|body\.dark)/.test(l)).every((l) => /#sec-users/.test(l)); })());
t('5.20 quickToggle 成功只更新 allUsers 與統計卡（不重畫）、失敗呼叫 loadUsers()（狀態在 uvState 不會掉）；uvState 在模組層級', /if \(u\) u\[field\] = value;\s*renderUserStat\(\);/.test(adminSrc) && /showToast\(data\.error \|\| '更新失敗'\); loadUsers\(\); return;/.test(adminSrc) && /^\s{4}const uvState = /m.test(adminSrc));
t('5.21 帳號用在 URL 路徑時 encodeURIComponent（quickToggle／儲存／重設密碼／刪除）', (adminSrc.match(/admin\/users\/\$\{encodeURIComponent\(/g) || []).length === 4 && !/admin\/users\/\$\{(username|editUn|un)\}/.test(adminSrc));
t('5.22 統計卡 renderUserStat 沒被改（四張卡、同樣的算法）', /const total   = allUsers\.length;\s*const admins  = allUsers\.filter\(u => u\.role === 'admin'\)\.length;\s*const active  = allUsers\.filter\(u => u\.active\)\.length;\s*const canDL   = allUsers\.filter\(u => u\.canDownloadContacts\)\.length;/.test(adminSrc));
t('5.23 初始化只有一次（uvWire 為即時函式、事件不重複掛）；分頁按鈕只建立一次', /\(function uvWire\(\)/.test(uvSrc) && /!== UV\.TABS\.length/.test(uvSrc));

// ═════════ 6) 未設定角色（role 空／缺漏）═════════
{
  const users = [mk('r_empty', '', ['ERP']), mk('r_null', null, ['ERP']), { username: 'r_missing', displayName: 'x', bu: ['ERP'], active: true }, mk('r_user', 'user', ['ERP']), mk('r_admin', 'admin', [])];
  const g = V.groupByRole(users);
  t('6.1 role 空字串／null／缺漏的帳號同組，組名「（未設定角色）」（不是空字串）；排在已登記角色之後', g.some((x) => x.role === '' && x.label === '（未設定角色）' && J(names(x.users).sort()) === J(['r_empty', 'r_missing', 'r_null'])) && g.every((x) => x.label !== '') && g[g.length - 1].role === '' && V.roleLabel(null) === '（未設定角色）');
  const opts = V.roleOptions(users);
  const none = opts.find((o) => o.label === '（未設定角色）');
  t('6.2 角色下拉：有未設定角色的帳號時多一個「（未設定角色）」選項，值是哨兵 __none__（不是空字串，否則會變成「全部角色」）；沒有這類帳號時不列', !!none && none.value === '__none__' && opts.every((o) => o.value !== '') && V.roleOptions([mk('a', 'user', ['ERP'])]).every((o) => o.value !== '__none__') && V.NONE_ROLE === '__none__');
  const f = (role) => names(V.applyFilters(users, { role }));
  t('6.3 選哨兵＝只篩出 role 空／null／缺漏的帳號；其他角色篩選與「全部」不受影響', J(f('__none__').sort()) === J(['r_empty', 'r_missing', 'r_null']) && J(f('user')) === J(['r_user']) && f('').length === 5);
  const v = V.buildView(users, { role: '__none__' });
  t('6.4 buildView：哨兵篩選時 ERP 分頁只有三個、群組只有「（未設定角色）」、收合鍵 ERP|（空角色）仍合法', v.tabs[0].count === 3 && v.groups.length === 1 && v.groups[0].label === '（未設定角色）' && V.validCollapseKey(v.groups[0].key));
}

// ═════════ 7) 狀態保存與接線（靜態：執行期的狀態遺失由 scratchpad e2e 實跑，這裡守住程式路徑沒被拆掉）═════════
{
  const fnBody = (name) => { const a = adminSrc.indexOf('function ' + name + '('); if (a < 0) return ''; let d = 0, i = adminSrc.indexOf('{', a); const s0 = i; for (; i < adminSrc.length; i++) { if (adminSrc[i] === '{') d++; else if (adminSrc[i] === '}') { d--; if (d === 0) break; } } return adminSrc.slice(s0, i + 1); };
  const loadUsersSrc = fnBody('loadUsers'), renderTableSrc = fnBody('renderUserTable'), uvRenderSrc = fnBody('uvRender'), uvSetTabSrc = fnBody('uvSetTab'), uvToggleSrc = fnBody('uvToggleGroup'), quickSrc = fnBody('quickToggle'), syncSrc = fnBody('uvSyncToolbar');
  t('7.0 找得到要檢查的函式本體', [loadUsersSrc, renderTableSrc, uvRenderSrc, uvSetTabSrc, uvToggleSrc, quickSrc, syncSrc].every((x) => x.length > 10));
  t('7.1 三個 sessionStorage 鍵（admin.users.tab／view／collapsed）存在，並在初始化時讀回＋驗證', /tab: 'admin\.users\.tab', view: 'admin\.users\.view', collapsed: 'admin\.users\.collapsed'/.test(uvSrc) && /tab: UV\.validTab\(uvStoreGet\(UV_KEY\.tab\)\)/.test(uvSrc) && /view: UV\.validView\(uvStoreGet\(UV_KEY\.view\)\)/.test(uvSrc) && /collapsed: new Set\(UV\.parseCollapsed\(uvStoreGet\(UV_KEY\.collapsed\)\)\)/.test(uvSrc));
  t('7.2 改變時寫回：換分頁寫 tab、收合寫 collapsed（序列化）、切換檢視寫 view', /uvStoreSet\(UV_KEY\.tab, id\)/.test(uvSetTabSrc) && /uvStoreSet\(UV_KEY\.collapsed, UV\.serializeCollapsed\(uvState\.collapsed\)\)/.test(uvToggleSrc) && /uvStoreSet\(UV_KEY\.view, uvState\.view\)/.test(uvSrc));
  t('7.3 每個改變狀態的操作之後都重畫（換分頁、收合、搜尋、角色／狀態／排序、清除、檢視）', /uvRender\(\)/.test(uvSetTabSrc) && /uvRender\(\)/.test(uvToggleSrc) && /uvState\.q = e\.target\.value; uvRender\(\)/.test(uvSrc) && /uvState\.role = e\.target\.value; uvRender\(\)/.test(uvSrc) && /uvState\.status = UV\.validStatus\(e\.target\.value\); uvRender\(\)/.test(uvSrc) && /uvState\.sort = UV\.validSort\(e\.target\.value\); uvRender\(\)/.test(uvSrc) && /function uvClearFilters\(\) \{[^}]*uvRender\(\)/.test(uvSrc));
  t('7.4 狀態物件 uvState 只宣告一次（模組層級 const），全檔沒有整個重新賦值；loadUsers／renderUserTable／uvRender 不重設 tab／搜尋／篩選／收合', (adminSrc.match(/^\s{4}const uvState = /gm) || []).length === 1 && !/\buvState\s*=[^=]/.test(adminSrc.replace(/const uvState = /, '')) && !/uvState/.test(loadUsersSrc) && renderTableSrc.replace(/\s+/g, '') === '{uvRender();}' && !/uvState\.(tab|q|status|view|sort|collapsed)\s*=[^=]/.test(uvRenderSrc) && !/uvState\.collapsed\s*=[^=]/.test(adminSrc.replace(uvToggleSrc, '')));
  t('7.5 loadUsers 抓完資料 → renderUserStat → renderUserTable（→ uvRender）；新增／編輯／刪除成功後走 loadUsers()（狀態保留）', /renderUserTable\(\)/.test(loadUsersSrc) && /uvRender\(\)/.test(renderTableSrc) && (adminSrc.match(/closeUserModal\(\);\s*loadUsers\(\);/g) || []).length === 1 && /showToast\('✅ 帳號已刪除'\);\s*closeDelModal\(\);\s*loadUsers\(\);/.test(adminSrc));
  t('7.6 開關失敗路徑（HTTP 錯誤、例外）都呼叫 loadUsers() 重抓重畫；成功路徑只就地改 allUsers＋renderUserStat', (quickSrc.match(/loadUsers\(\)/g) || []).length === 2 && /!r\.ok\) \{ showToast\(data\.error \|\| '更新失敗'\); loadUsers\(\); return; \}/.test(quickSrc) && /catch \(e\) \{ if \(e\.message !== 'Session expired'\) \{ showToast\('操作失敗，請重試'\); loadUsers\(\); \} \}/.test(quickSrc) && /if \(u\) u\[field\] = value;\s*renderUserStat\(\);/.test(quickSrc));
  t('7.7 工具列由 uvState 回填（搜尋框、角色／狀態／排序、檢視按鈕 aria-pressed、清除連結），只在值不同時寫入搜尋框', /q\.value !== uvState\.q\) q\.value = uvState\.q/.test(syncSrc) && /sel\.value !== uvState\.role/.test(syncSrc) && /uvState\.status/.test(syncSrc) && /uvState\.sort/.test(syncSrc) && /aria-pressed', String\(uvState\.view === 'table'\)/.test(syncSrc) && /\$\('uvClear'\)\.hidden = !v\.filtersOn/.test(syncSrc));
  t('7.8 表格欄寬：不再對全部 td／th 套 nowrap，只有第 5、7、8、11 欄（角色、狀態、存取模式、操作）；操作按鈕列 flex-wrap:nowrap（最小寬度約等於改版前 1417→1427；wrap 會讓 375～1500px 下列高變 118px）', (() => { const m = adminSrc.split('\n').filter((l) => /#sec-users \.uv-tablewrap \.admin-table[^{]*\{[^}]*white-space: nowrap/.test(l)); return m.length === 1 && /td:nth-child\(5\).*td:nth-child\(7\).*td:nth-child\(8\).*td:nth-child\(11\)/.test(m[0]) && !/ th[,\s{]/.test(m[0].split('{')[0]) && /<div style="display:flex;gap:6px;flex-wrap:nowrap">\s*<button class="btn btn-sm btn-primary"  data-action="edit"   data-username="\$\{un\}">編輯/.test(rowSrc); })());
  t('7.9 空狀態文字色對比：淺色 #6b7280（>=4.5:1）、深色 #8b949e', /#sec-users \.uv-empty \{ text-align: center; color: #6b7280;/.test(adminSrc) && /body\.dark #sec-users \.uv-empty \{ color: #8b949e; \}/.test(adminSrc));
}

console.log('');
let pass = 0, fail = 0;
res.forEach(([n, ok, x]) => { console.log((ok ? 'PASS ' : 'FAIL ') + n + (!ok && x ? '  <- ' + x : '')); ok ? pass++ : fail++; });
console.log(`\n帳號管理檢視檢查：PASS ${pass} / FAIL ${fail}`);
process.exit(fail ? 1 : 0);
