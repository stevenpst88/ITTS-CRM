/*!
 * 後台「帳號管理」檢視邏輯 —— 純函式（無相依、不碰 DOM），可在瀏覽器與 Node 載入
 *   瀏覽器：<script src="admin-users-view.js"> → window.AdminUsersView
 *   Node  ：require('../_client/admin-users-view.js')（scripts/check-admin-users-view.js 用它；Vercel 的 includeFiles 已含 _client/**）
 *
 * 規則（業主 2026-10 核可）：
 *   1) BU 分頁固定五個：ERP、ITS、MDM、CRM、全公司／集團。帳號的 bu 陣列有幾個 BU，就出現在幾個 BU 分頁（各分頁數字因此會重疊）。
 *      沒有任何有效 BU 的帳號（admin／executive／accounting_manager／finance_manager、系統帳號池 pool、資料壞掉的帳號）一律放「全公司／集團」。
 *      集團角色（tecopm／groupsales）以集團界定範圍、不掛 BU，也固定放「全公司／集團」。
 *   2) 分頁內依角色（職能＝帳號 role）分組，組序固定（見 ROLE_ORDER）；沒登記的角色排在後面、組名用原始值；系統帳號池 pool 永遠最後。
 *      組內與「依名稱排序」都用同一個顯示名稱規則：暱稱 > 顯示名稱 > 帳號（與 admin.html 的 adminUserLabel、後端 userLabel 一致）。
 *   3) 搜尋只比對 帳號／顯示名稱／暱稱（NFKC、不分大小寫、前後空白忽略、逐字比對不用正規式），**不比對 email**（email 在畫面上是遮罩的，搜尋不能反推）。
 *   4) 畫面狀態（分頁、檢視、收合群組）存 sessionStorage；讀回來的值一律經本檔的 valid*／parse* 驗證，壞值回預設，不丟錯。
 */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.AdminUsersView = api;
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const BU_LIST = ['ERP', 'ITS', 'MDM', 'CRM'];
  const TAB_CORP = 'CORP';
  const TABS = [
    { id: 'ERP', label: 'ERP' },
    { id: 'ITS', label: 'ITS' },
    { id: 'MDM', label: 'MDM' },
    { id: 'CRM', label: 'CRM' },
    { id: TAB_CORP, label: '全公司／集團' }
  ];
  const TAB_IDS = TABS.map(t => t.id);

  // 角色順序：管理層 → 一般主管 → 業務（含行銷）→ 顧問主管 → 秘書 → 集團角色（可編輯的集團業務在前、唯讀的集團PM在後）→ 系統帳號池
  const ROLE_ORDER = [
    'admin', 'executive', 'accounting_manager', 'finance_manager',
    'manager1', 'manager2',
    'user', 'marketing',
    'consult_manager_south', 'consult_manager_north',
    'secretary',
    'groupsales', 'tecopm'
  ];
  const POOL_ROLE = 'pool';
  const NONE_ROLE = '__none__';          // 角色篩選下拉裡「（未設定角色）」的值：篩出 role 為空／缺漏的帳號
  const NONE_LABEL = '（未設定角色）';
  const GROUP_ROLES = ['tecopm', 'groupsales'];
  const ROLE_LABELS = {
    admin: '系統管理員', executive: '董事長/總經理', accounting_manager: '會計主管', finance_manager: '財務主管',
    manager1: '一級主管', manager2: '二級主管', user: '一般業務', marketing: '行銷人員',
    consult_manager_south: '南區顧問主管', consult_manager_north: '北區顧問主管', secretary: '秘書',
    groupsales: '集團業務', tecopm: '集團PM（唯讀）',
    pool: '客戶池（系統帳號）'
  };
  const SORTS = ['role', 'name'];
  const VIEWS = ['table', 'card'];
  const STATUSES = ['', 'active', 'inactive'];
  const MAX_COLLAPSED = 200;
  const MAX_KEY_LEN = 120;

  const hasOwn = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
  const isUser = (u) => !!u && typeof u === 'object' && !Array.isArray(u);
  const str = (v) => (v === null || v === undefined ? '' : String(v));

  /** 帳號顯示名稱：暱稱 > 顯示名稱 > 帳號（與 admin.html adminUserLabel 相同） */
  function userLabel(u) { return isUser(u) ? (str(u.nickname) || str(u.displayName) || str(u.username)) : ''; }

  /** 角色的中文標籤；沒登記的回原始值（空值回空字串） */
  function roleLabel(role) {
    const r = str(role);
    if (r === '') return NONE_LABEL;
    return hasOwn(ROLE_LABELS, r) ? ROLE_LABELS[r] : r;
  }

  /** 角色排序鍵：已登記的依 ROLE_ORDER；未登記的排在已登記之後（再依原始值）；pool 永遠最後 */
  function roleRank(role) {
    const r = str(role);
    if (r === POOL_ROLE) return ROLE_ORDER.length + 2;
    const i = ROLE_ORDER.indexOf(r);
    return i >= 0 ? i : ROLE_ORDER.length + 1;
  }
  function compareRole(a, b) {
    const ra = roleRank(a), rb = roleRank(b);
    if (ra !== rb) return ra - rb;
    const x = str(a), y = str(b);
    return x < y ? -1 : x > y ? 1 : 0;
  }

  /** 搜尋／比對用正規化：NFKC、小寫、前後空白去掉（內部連續空白壓成一個） */
  function normText(s) {
    let v = str(s);
    try { v = v.normalize('NFKC'); } catch (_) { /* 舊環境沒有 normalize 也不丟錯 */ }
    return v.toLowerCase().replace(/\s+/g, ' ').trim();
  }

  /** 帳號的有效 BU：trim + 轉大寫、只留 ERP/ITS/MDM/CRM、去重；非陣列的字串視為單一 BU；其他型別當空 */
  function effectiveBus(u) {
    if (!isUser(u)) return [];
    let raw = u.bu;
    if (typeof raw === 'string') raw = [raw];
    if (!Array.isArray(raw)) return [];
    const out = [];
    raw.forEach(b => {
      if (typeof b !== 'string') return;
      const v = b.trim().toUpperCase();
      if (BU_LIST.includes(v) && !out.includes(v)) out.push(v);
    });
    return out;
  }

  /** 帳號所屬的分頁（可多個）。集團角色、沒有有效 BU 的 → 只在「全公司／集團」 */
  function tabsOf(u) {
    if (!isUser(u)) return [];
    if (GROUP_ROLES.includes(str(u.role))) return [TAB_CORP];
    const bus = effectiveBus(u);
    return bus.length ? bus : [TAB_CORP];
  }

  function isActive(u) { return !!(isUser(u) && u.active); }

  /** 篩選條件是否生效（搜尋字串只有空白算沒生效） */
  function filtersActive(f) {
    f = f || {};
    return !!(normText(f.q) || str(f.role) || str(f.status));
  }

  /** 單一帳號是否符合篩選（q／role／status；全部條件 AND） */
  function matches(u, f) {
    if (!isUser(u)) return false;
    f = f || {};
    const role = str(f.role);
    if (role === NONE_ROLE) { if (str(u.role) !== '') return false; }
    else if (role && str(u.role) !== role) return false;
    const st = str(f.status);
    if (st === 'active' && !isActive(u)) return false;
    if (st === 'inactive' && isActive(u)) return false;
    const q = normText(f.q);
    if (q) {
      const hay = [u.username, u.displayName, u.nickname].map(normText);
      if (!hay.some(h => h.includes(q))) return false;
    }
    return true;
  }

  function applyFilters(users, f) {
    return (Array.isArray(users) ? users : []).filter(u => matches(u, f));
  }

  /** 依分頁分割；多 BU 帳號會出現在每個所屬分頁（同一個物件參考） */
  function partition(users) {
    const out = {};
    TAB_IDS.forEach(id => { out[id] = []; });
    (Array.isArray(users) ? users : []).forEach(u => {
      tabsOf(u).forEach(t => out[t].push(u));
    });
    return out;
  }

  let _collator = null;
  function collator() {
    if (_collator) return _collator;
    try { _collator = new Intl.Collator('zh-TW', { numeric: true, sensitivity: 'base' }); }
    catch (_) { _collator = { compare: (a, b) => (a < b ? -1 : a > b ? 1 : 0) }; }
    return _collator;
  }

  /** 穩定排序（不改輸入）：先依顯示名稱（中文用 zh-TW 排序、數字依數值），再依帳號，最後依原順序 */
  function sortByName(list) {
    const c = collator();
    return (Array.isArray(list) ? list : []).map((u, i) => ({ u, i })).sort((a, b) => {
      const d = c.compare(userLabel(a.u), userLabel(b.u));
      if (d) return d;
      const e = c.compare(str(a.u.username), str(b.u.username));
      if (e) return e;
      const f = str(a.u.username) < str(b.u.username) ? -1 : str(a.u.username) > str(b.u.username) ? 1 : 0;
      return f || a.i - b.i;
    }).map(x => x.u);
  }

  /** 依角色分組（組序見 roleRank）；組內依名稱排序。回 [{ role, label, users }] */
  function groupByRole(list) {
    const map = new Map();
    (Array.isArray(list) ? list : []).forEach(u => {
      if (!isUser(u)) return;
      const r = str(u.role);
      if (!map.has(r)) map.set(r, []);
      map.get(r).push(u);
    });
    return Array.from(map.keys()).sort(compareRole).map(r => ({ role: r, label: roleLabel(r), users: sortByName(map.get(r)) }));
  }

  /** 角色篩選下拉的選項：登記過的角色（pool 只有資料裡真的有才列）＋資料裡出現的未登記角色；依 roleRank 排序 */
  function roleOptions(users) {
    const present = new Set();
    (Array.isArray(users) ? users : []).forEach(u => { if (isUser(u)) present.add(str(u.role)); });
    const set = new Set(ROLE_ORDER);
    present.forEach(r => set.add(r));
    return Array.from(set).filter(r => r !== '' || present.has('')).sort(compareRole).map(r => ({ value: r === '' ? NONE_ROLE : r, label: roleLabel(r) }));
  }

  // ── 驗證從 sessionStorage 讀回的值 ──
  function validTab(v) { return typeof v === 'string' && TAB_IDS.includes(v) ? v : null; }
  function validView(v) { return typeof v === 'string' && VIEWS.includes(v) ? v : 'table'; }
  function validSort(v) { return typeof v === 'string' && SORTS.includes(v) ? v : 'role'; }
  function validStatus(v) { return typeof v === 'string' && STATUSES.includes(v) ? v : ''; }

  /** 收合群組的鍵：分頁|角色 */
  function collapseKey(tab, role) { return str(tab) + '|' + str(role); }
  function validCollapseKey(k) {
    if (typeof k !== 'string' || k.length > MAX_KEY_LEN) return false;
    const i = k.indexOf('|');
    return i > 0 && TAB_IDS.includes(k.slice(0, i));
  }
  /** 讀回 JSON 字串 → 鍵陣列（壞 JSON、非陣列、壞元素、過多都安全處理） */
  function parseCollapsed(raw) {
    let v;
    try { v = typeof raw === 'string' ? JSON.parse(raw) : raw; } catch (_) { return []; }
    if (!Array.isArray(v)) return [];
    const out = [];
    for (const k of v) {
      if (out.length >= MAX_COLLAPSED) break;
      if (validCollapseKey(k) && !out.includes(k)) out.push(k);
    }
    return out;
  }
  function serializeCollapsed(keys) {
    const arr = [];
    (keys && typeof keys[Symbol.iterator] === 'function' ? Array.from(keys) : []).forEach(k => {
      if (validCollapseKey(k) && !arr.includes(k) && arr.length < MAX_COLLAPSED) arr.push(k);
    });
    return JSON.stringify(arr);
  }
  /** 切換收合：回新的 Set（不改輸入） */
  function toggleCollapsed(set, key) {
    const out = new Set(set || []);
    if (out.has(key)) out.delete(key); else out.add(key);
    return out;
  }

  /** 想要的分頁無效 → 第一個有帳號的分頁；都沒有 → 第一個分頁 */
  function pickTab(wanted, counts) {
    const w = validTab(wanted);
    if (w) return w;
    const c = counts || {};
    for (const id of TAB_IDS) if ((c[id] || 0) > 0) return id;
    return TAB_IDS[0];
  }

  /**
   * 一次算出畫面需要的全部資料。
   *   state：{ q, role, status, sort, tab }（tab 可為 null → 自動挑第一個有帳號的分頁）
   *   回傳：{ total, matched, active, tabs:[{id,label,count}], tab, users, groups|null, others:[{id,label,count}], filtersOn }
   *   total＝全部有效帳號數；matched＝符合篩選的不重複帳號數（多 BU 帳號只算一次）；users＝目前分頁排序後的平面清單；
   *   groups＝sort 為 role 時才有（[{ role, label, key, users }]），sort 為 name 時是 null（平面、不分組）。
   */
  function buildView(users, state) {
    state = state || {};
    const all = (Array.isArray(users) ? users : []).filter(isUser);
    const filtered = applyFilters(all, state);
    const parts = partition(filtered);
    const counts = {};
    TAB_IDS.forEach(id => { counts[id] = parts[id].length; });
    const tab = pickTab(state.tab, counts);
    const sort = validSort(state.sort);
    const list = parts[tab];
    let flat, groups = null;
    if (sort === 'role') {
      groups = groupByRole(list).map(g => ({ role: g.role, label: g.label, key: collapseKey(tab, g.role), users: g.users }));
      flat = [].concat.apply([], groups.map(g => g.users));
    } else {
      flat = sortByName(list);
    }
    return {
      total: all.length,
      matched: filtered.length,
      tabs: TABS.map(t => ({ id: t.id, label: t.label, count: counts[t.id] })),
      tab,
      users: flat,
      groups,
      others: TABS.filter(t => t.id !== tab && counts[t.id] > 0).map(t => ({ id: t.id, label: t.label, count: counts[t.id] })),
      filtersOn: filtersActive(state)
    };
  }

  return {
    BU_LIST, TABS, TAB_IDS, TAB_CORP, ROLE_ORDER, ROLE_LABELS, GROUP_ROLES, POOL_ROLE, NONE_ROLE, NONE_LABEL, SORTS, VIEWS,
    userLabel, roleLabel, roleRank, compareRole, normText, effectiveBus, tabsOf, isActive,
    filtersActive, matches, applyFilters, partition, sortByName, groupByRole, roleOptions,
    validTab, validView, validSort, validStatus,
    collapseKey, validCollapseKey, parseCollapsed, serializeCollapsed, toggleCollapsed,
    pickTab, buildView
  };
});
