/* 簽核流程設計器（M2）— flow-designer.js
 * 後台「簽核流程設計器」：泳道編輯（拖拉關卡、條件卡、屬性抽屜、條件造句編輯器）、模擬器、驗證面板、總覽（唯讀自由畫布，見 flow-overview.js）、草稿清單。
 * 後端：lib/quoteFlowDesignerRoutes.js（開關 QUOTE_FLOW_DESIGNER=0|1，預設關閉；關閉時這個區塊不會出現）。
 *
 * 重要規則
 *  - 草稿不能生效：本頁沒有任何「提交覆核／生效」功能（M3 才開放），只顯示灰色說明。
 *  - 資料模型就是 M1 的流程定義（FlowDef）本身；每一次編輯都先問後端 /designer/validate（11 條驗證＋保護底線），
 *    違反底線或結構無效的編輯即時拒絕，不會進入畫面狀態。後端 PUT 儲存時會再驗一次（不只靠前端）。
 *  - 所有動態內容經 h`` 模板（插值自動跳脫）或 textContent 輸出；不使用 inline script／on* 屬性／style= 屬性。
 *  - 純函式（itemText／astToEdit／editToAst／wanToYuan…）掛在 FlowDesigner._test 上，供 scripts/check-quote-flow-designer-ui.js 測試。 */
(function (root) {
  'use strict';
  const doc = root.document;

  // ───────────────────────── 跳脫與模板（插值預設跳脫）─────────────────────────
  const ESC_MAP = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ESC_MAP[c]);
  function Raw(s) { this.s = s; }
  const raw = (s) => new Raw(s);
  const part = (v) => (v instanceof Raw ? v.s : Array.isArray(v) ? v.map(part).join('') : esc(v));
  /** 標籤模板：h`<b>${userText}</b>` —— 插值一律跳脫，只有 Raw（另一個 h`` 的結果）原樣輸出 */
  function h(strs) { let out = strs[0]; for (let i = 1; i < arguments.length; i++) out += part(arguments[i]) + strs[i]; return new Raw(out); }
  const put = (el, r) => { if (el) el.innerHTML = r instanceof Raw ? r.s : esc(r); };   // 唯一的 innerHTML 出口
  const $ = (s, r) => (r || doc).querySelector(s);
  const $$ = (s, r) => Array.from((r || doc).querySelectorAll(s));
  const clone = (o) => JSON.parse(JSON.stringify(o));
  const BASE = '/api/admin/quote-flow';
  const canonStr = (v) => { if (Array.isArray(v)) return '[' + v.map(canonStr).join(',') + ']'; if (v && typeof v === 'object') return '{' + Object.keys(v).sort().map((k) => JSON.stringify(k) + ':' + canonStr(v[k])).join(',') + '}'; return JSON.stringify(v === undefined ? null : v); };

  // ───────────────────────── 常數 ─────────────────────────
  const TIERS = ['mgr2', 'mgr1', 'gm', 'chairman', 'board'];
  const TIER = {
    mgr2: { short: '二級主管', rank: 1, color: 'teal', locked: false, rule: '業務直屬的二級主管', ruleLong: '業務帳號「直屬主管」那一層，且必須是在職的二級主管。送簽當下就鎖定成具體的人。', timing: '送簽時鎖定指定人' },
    mgr1: { short: '一級主管', rank: 2, color: 'blue', locked: false, rule: '沿主管鏈往上第一位一級主管', ruleLong: '從業務往上找，第一位在職的一級主管。找不到人時無法送簽。送簽當下鎖定成具體的人。', timing: '送簽時鎖定指定人' },
    gm: { short: '總經理', rank: 3, color: 'pink', locked: false, rule: '名冊「總經理」的任一位', ruleLong: '名冊「總經理」裡任何一位在職、非唯讀的人都可以簽。核准當下才判斷資格（名冊是活的）。', timing: '核准當下判斷（活名單）' },
    chairman: { short: '董事長', rank: 4, color: 'amber', locked: false, rule: '名冊「董事長」的任一位', ruleLong: '名冊「董事長」裡任何一位在職、非唯讀的人。核准當下才判斷資格。', timing: '核准當下判斷（活名單）' },
    board: { short: '董事會代核', rank: 5, color: 'violet', locked: true, rule: '董事會決議，由秘書代為登錄', ruleLong: '董事會決議的結果，由名冊「董事會代核人」與秘書角色代為登錄；必填決議日期與文號。', timing: '核准當下判斷（活名單）', lockWhy: '代核／秘書／名冊的語意要等 Steven 與董事長溝通完才決定，所以設計器只能決定「哪些路線含這一關」，不能改它的簽核方式。' },
  };
  const COLORS = ['teal', 'blue', 'pink', 'amber', 'violet', 'green', 'slate'];
  const COLOR_ZH = { teal: '青', blue: '藍', pink: '粉', amber: '琥珀', violet: '紫', green: '綠', slate: '灰' };
  const CATS = [['consult', '顧問服務'], ['software', '軟體規劃服務'], ['hardware', '硬體規劃服務'], ['consult_software', '顧問服務+軟體服務'], ['consult_hardware', '顧問服務+硬體服務'], ['crm', '客服服務(CRM)'], ['mdm', '帳單列印(MDM)'], ['ot', 'OT 策略性專案'], ['other', '其他（表外）']];
  const CAT_ZH = Object.fromEntries(CATS);
  const CLASS_KEYS = [['consult', '顧問服務'], ['software', '軟體規劃服務'], ['hardware', '硬體規劃服務'], ['crm', '客服服務(CRM)'], ['mdm', '帳單列印(MDM)'], ['ot', 'OT 策略性專案'], ['other', '其他']];
  const CLASS_ZH = Object.fromEntries(CLASS_KEYS);
  const BUS = ['ERP', 'ITS', 'MDM', 'CRM', '資安'];
  const OP_ZH = { '==': '等於', '!=': '不等於', '>': '大於', '>=': '大於等於', '<': '小於', '<=': '小於等於', in: '屬於', not_in: '不屬於', contains: '包含', not_contains: '不包含' };
  const NUM_OPS = ['>', '>=', '==', '<', '<=', '!='];
  const STR_OPS = ['==', '!=', 'in', 'not_in'];
  // 變數與運算子沿用 M1 白名單（lib/quoteFlow.js VARS）
  const VARS = {
    amount: { label: '折後未稅金額', type: 'money', ops: NUM_OPS },
    marginLevel: { label: '毛利層級', type: 'level', ops: NUM_OPS },
    marginPct: { label: '毛利率', type: 'pct', ops: NUM_OPS },
    category: { label: '商品歸類', type: 'enum', ops: STR_OPS },
    classKeys: { label: '商品類別', type: 'set', ops: ['contains', 'not_contains'] },
    ownerBuPrimary: { label: '業務所屬 BU', type: 'enum', ops: STR_OPS },
    owner: { label: '指定業務', type: 'string', ops: STR_OPS },
  };
  const VAR_ORDER = ['amount', 'marginLevel', 'marginPct', 'category', 'classKeys', 'ownerBuPrimary', 'owner'];
  const VAR_DESC = { amount: '大於／小於／等於…', marginLevel: '層級 1／2／3（核決表算出）', marginPct: '毛利率（%，後台才看得到）', category: '等於／屬於', classKeys: '包含某類商品', ownerBuPrimary: '等於／屬於', owner: '指定某位業務' };
  const LVL_ZH = { 1: '層級 1（一般）', 2: '層級 2（偏低）', 3: '層級 3（很低）' };

  // ───────────────────────── 金額（十進位字串，不用浮點）─────────────────────────
  /** 萬元字串 → 元字串（至多 6 位小數的萬元＝2 位小數的元）；無效回 null */
  function wanToYuan(s) {
    s = String(s == null ? '' : s).trim();
    const m = /^(\d{1,12})(?:\.(\d{1,6}))?$/.exec(s);
    if (!m) return null;
    const cents = BigInt(m[1]) * 1000000n + BigInt((m[2] || '').padEnd(6, '0'));
    if (cents <= 0n) return null;
    const yuan = cents / 100n, c = cents % 100n;
    return yuan.toString() + (c ? '.' + String(c).padStart(2, '0') : '');
  }
  /** 元字串 → 萬元字串（供輸入框顯示） */
  function yuanToWan(s) {
    const m = /^(\d+)(?:\.(\d{1,2}))?$/.exec(String(s == null ? '' : s));
    if (!m) return '';
    const cents = BigInt(m[1]) * 100n + BigInt((m[2] || '').padEnd(2, '0') || '0');
    const wan = cents / 1000000n, frac = cents % 1000000n;
    return wan.toString() + (frac ? '.' + String(frac).padStart(6, '0').replace(/0+$/, '') : '');
  }
  const comma = (s) => String(s).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  /** 元字串 → 白話金額（整萬顯示「N 萬」，其他顯示「N 元」） */
  function fmtYuan(s) {
    const m = /^(\d+)(?:\.(\d{1,2}))?$/.exec(String(s));
    if (!m) return String(s);
    const yuan = BigInt(m[1]), c = (m[2] || '').padEnd(2, '0');
    if ((!c || c === '00') && yuan % 10000n === 0n && yuan >= 10000n) return comma((yuan / 10000n).toString()) + ' 萬';
    return comma(m[1]) + (c && c !== '00' ? '.' + c : '') + ' 元';
  }

  // ───────────────────────── 條件：AST ↔ 造句（平面 mode＋items）─────────────────────────
  function valText(v, val) {
    const one = (x) => (v === 'amount' ? fmtYuan(x) : v === 'marginLevel' ? '層級 ' + x : v === 'marginPct' ? x + '%' : v === 'category' ? (CAT_ZH[x] || x) : v === 'classKeys' ? (CLASS_ZH[x] || x) : String(x));
    return Array.isArray(val) ? (val.length ? '[' + val.map(one).join('、') + ']' : '[（無）]') : one(val);
  }
  /** 單一比較的白話（與後端 describeCmp 同一套寫法） */
  function itemText(it) { const vd = VARS[it.v]; return (vd ? vd.label : String(it.v)) + ' ' + (OP_ZH[it.o] || it.o) + ' ' + valText(it.v, it.val); }
  function isCmp(x) { return !!x && typeof x === 'object' && !Array.isArray(x) && Object.keys(x).length === 1 && Array.isArray(x.cmp) && x.cmp.length === 3 && typeof x.cmp[0] === 'string'; }
  function cmpToItem(x) {
    const it = { v: x.cmp[0], o: x.cmp[1], val: Array.isArray(x.cmp[2]) ? x.cmp[2].slice() : x.cmp[2] };
    if (it.v === 'amount') it.wan = yuanToWan(it.val);
    return it;
  }
  /** when（AST）→ { mode, items } ；預設路線(true) → null；超出編輯器範圍（巢狀／否定…）→ { unsupported:true } */
  function astToEdit(w) {
    if (w === true) return null;
    if (isCmp(w)) return { mode: 'all', items: [cmpToItem(w)] };
    if (w && typeof w === 'object' && Array.isArray(w.all) && Object.keys(w).length === 1 && w.all.length && w.all.every(isCmp)) return { mode: 'all', items: w.all.map(cmpToItem) };
    if (w && typeof w === 'object' && Array.isArray(w.any) && Object.keys(w).length === 1 && w.any.length && w.any.every(isCmp)) return { mode: 'any', items: w.any.map(cmpToItem) };
    return { unsupported: true };
  }
  function itemToCmp(it) { return { cmp: [it.v, it.o, Array.isArray(it.val) ? it.val.slice() : it.val] }; }
  function editToAst(e) {
    if (!e) return true;
    const cs = e.items.map(itemToCmp);
    return cs.length === 1 ? cs[0] : { [e.mode === 'any' ? 'any' : 'all']: cs };
  }
  function defaultItem(v) {
    switch (v) {
      case 'amount': return { v, o: '>', val: '10000000', wan: '1000' };
      case 'marginLevel': return { v, o: '==', val: 2 };
      case 'marginPct': return { v, o: '<', val: '10' };
      case 'category': return { v, o: '==', val: 'consult' };
      case 'classKeys': return { v, o: 'contains', val: 'consult' };
      case 'ownerBuPrimary': return { v, o: '==', val: 'ERP' };
      default: return { v: 'owner', o: '==', val: '' };
    }
  }
  /** 編輯中單一條件項目的欄位檢查；回傳錯誤訊息或 '' */
  function itemError(it) {
    const vd = VARS[it.v];
    if (!vd) return '不在白名單的變數';
    if (!vd.ops.includes(it.o)) return vd.label + '不能用「' + (OP_ZH[it.o] || it.o) + '」';
    if (vd.type === 'money') { if (!/^\d{1,16}(\.\d{1,2})?$/.test(String(it.val || '')) || /^0(\.0+)?$/.test(String(it.val))) return '金額要填大於 0 的數字（萬元，至多 6 位小數）'; }
    else if (vd.type === 'pct') { if (!/^-?\d{1,3}(\.\d{1,4})?$/.test(String(it.val == null ? '' : it.val))) return '毛利率要填數字（%，至多 4 位小數）'; }
    else if (vd.type === 'level') { if (![1, 2, 3].includes(it.val)) return '毛利層級只能是 1、2、3'; }
    else if (vd.type === 'enum' || vd.type === 'string') {
      if (it.o === 'in' || it.o === 'not_in') { if (!Array.isArray(it.val) || !it.val.length) return vd.label + '「' + (OP_ZH[it.o]) + '」要至少選一個值'; if (vd.type === 'string' && it.val.some((x) => !x || x.length > 100)) return '帳號清單內容無效'; }
      else if (typeof it.val !== 'string' || !it.val) return vd.label + '的值不能空白';
    } else if (vd.type === 'set') { if (typeof it.val !== 'string' || !it.val) return '請選擇商品類別'; }
    return '';
  }

  // ───────────────────────── 流程定義的小工具 ─────────────────────────
  const nodeOfTier = (def, tier) => def.nodes.find((n) => n.tier === tier);
  const nodeById = (def, id) => def.nodes.find((n) => n.id === id);
  const tierOfId = (def, id) => { const n = nodeById(def, id); return n ? n.tier : null; };
  const labelOfTier = (def, tier) => { const n = nodeOfTier(def, tier); return n ? n.label : TIER[tier].short; };
  const colorOfTier = (def, tier) => { const n = nodeOfTier(def, tier); return (n && n.color) || TIER[tier].color; };
  const laneOf = (def, id) => def.routes.find((r) => r.id === id);
  const pathNames = (def, path) => (path.length ? path.map((id) => (nodeById(def, id) || { label: id }).label).join(' → ') : '（無任何關卡）');
  const fmtTime = (iso) => { const t = Date.parse(iso); if (!isFinite(t)) return ''; const d = new Date(t), p = (n) => String(n).padStart(2, '0'); return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes()); };
  let idSeq = 0;
  const newRouteId = () => 'r_' + Date.now().toString(36) + (idSeq++).toString(36);

  // ───────────────────────── 狀態 ─────────────────────────
  const S = {
    enabled: false, inited: false, meta: null, tab: 'design',
    def: null, view: null, checks: null, rc: null,
    draft: { id: null, name: '', note: '', updatedAt: null, baseVer: '' }, saved: '',
    drafts: [], draftsLoaded: false,
    simIn: null, simRes: null, scan: null, simErr: '',
    sel: null, drawer: null, banner: '', rejects: 0, focusKey: null,
    ovCtl: null, ovSeq: 0, ovSimOn: true, ovSummary: '',
    queue: Promise.resolve(), skipUsers: 0, busy: false, probed: false,
  };
  const isMobile = () => !!(root.matchMedia && root.matchMedia('(max-width:640px)').matches);
  const canEdit = () => S.enabled && !!S.def && !isMobile();
  const isDirty = () => !!S.def && (canonStr(S.def) + '|' + S.draft.name + '|' + S.draft.note) !== S.saved;
  const snapshotSaved = () => { S.saved = canonStr(S.def) + '|' + S.draft.name + '|' + S.draft.note; };

  async function api(method, path, body) {
    const f = typeof root.adminFetch === 'function' ? root.adminFetch : root.fetch.bind(root);
    const opt = { method, headers: { 'Content-Type': 'application/json' } };
    if (body !== undefined) opt.body = JSON.stringify(body);
    const r = await f(BASE + path, opt);
    let data = null;
    try { data = await r.json(); } catch (_) { data = null; }
    return { ok: r.ok, status: r.status, data };
  }

  // ───────────────────────── 通知／對話框 ─────────────────────────
  let liveT, toastT;
  function announce(t) { const l = $('#fdLive'); if (!l) return; l.textContent = ''; clearTimeout(liveT); liveT = setTimeout(() => { l.textContent = t; }, 30); }
  function toast(t) { const e = $('#fdToast'); if (!e) return; e.textContent = t; e.classList.add('show'); clearTimeout(toastT); toastT = setTimeout(() => e.classList.remove('show'), 4200); announce(t); }
  function dialog(inner, mount) {
    const dlg = $('#fdDlg');
    if (dlg.open) return Promise.resolve('cancel');   // 已有對話框開著（連點）：不重複 showModal
    put(dlg, inner);
    dlg.returnValue = '';
    return new Promise((res) => {
      dlg.onclose = () => res(dlg.returnValue);
      dlg.showModal();
      $$('[data-close]', dlg).forEach((b) => b.addEventListener('click', () => { if (b.dataset.close === 'ok' && mount && mount.validate && !mount.validate(dlg)) return; dlg.close(b.dataset.close); }));
      if (mount && mount.init) mount.init(dlg);
      const f = $('[autofocus]', dlg) || $('textarea,input', dlg) || $('button[data-close="ok"]', dlg);
      if (f) f.focus();
    });
  }
  async function confirmBox(title, body, okText, danger) {
    const r = await dialog(h`<h2 id="fdDlgT">${title}</h2><div>${body}</div><div class="fd-acts"><button type="button" class="fd-btn" data-close="cancel" autofocus>取消</button><button type="button" class="fd-btn ${danger ? 'danger' : 'primary'}" data-close="ok">${okText || '確認'}</button></div>`);
    return r === 'ok';
  }

  // ───────────────────────── 版面骨架（只建一次）─────────────────────────
  const TABS = [['design', '設計器'], ['sim', '模擬器'], ['val', '驗證'], ['ov', '總覽'], ['drafts', '草稿清單']];
  function buildScaffold(host) {
    put(host, h`
<div class="fd-root" id="fdApp">
  <div class="fd-head"><span class="fd-pill warn" title="M2：只能存草稿、模擬與檢視">草稿不會生效</span><span id="fdStatusPills" class="fd-row"></span></div>
  <div class="fd-note" role="note"><b>說明</b><span>在這裡設計「報價單簽核流程」的草稿，並用模擬器試算。<b>草稿不會影響任何簽核</b>——線上簽核仍完全走現行流程。生效、版本治理與雙人覆核將於 M3 開放。</span><button type="button" class="fd-btn sm" disabled aria-disabled="true" id="fdPropose" title="M3 開放">提交覆核（M3 開放）</button></div>
  <div class="fd-note fd-mobile-note" role="note"><b>手機唯讀</b><span>手機上設計器為唯讀檢視（拖拉編輯在小螢幕易誤觸），模擬器與總覽可正常使用。要編輯流程請改用桌機。</span></div>
  <div class="fd-bar" id="fdBar"></div>
  <nav class="fd-tabs" role="tablist" aria-label="功能分頁" id="fdTabs">${TABS.map(([k, n]) => h`<button type="button" class="fd-tab" role="tab" id="fdTabBtn-${k}" aria-controls="fdTab-${k}" aria-selected="false" data-tab="${k}" tabindex="-1">${n}<span class="fd-cnt" id="fdCnt-${k}" hidden></span></button>`)}</nav>
  <section id="fdTab-design" role="tabpanel" aria-labelledby="fdTabBtn-design" hidden>
    <div class="fd-grid" id="fdGrid">
      <aside class="fd-panel fd-palette" id="fdPalette" aria-label="節點與條件面板"></aside>
      <section class="fd-panel fd-canvas" id="fdCanvas" aria-labelledby="fdCanvasH"></section>
      <aside class="fd-panel fd-drawer" id="fdDrawer" role="dialog" aria-modal="false" aria-labelledby="fdDrawerH" hidden></aside>
    </div>
  </section>
  <section id="fdTab-sim" role="tabpanel" aria-labelledby="fdTabBtn-sim" hidden>
    <div class="fd-panel">
      <h2 class="fd-title">模擬器：輸入一張假想單，看它會怎麼簽</h2>
      <p class="fd-help">輸入的是假想值（不是真實報價單）。結果由後端的流程引擎（與正式送簽同一個函式）計算；命中的路線會同步在「設計器」與「總覽」高亮，並用白話說明為什麼走這條。</p>
      <form id="fdSimForm" class="fd-simform" autocomplete="off">
        <div><label class="fd-lbl" for="fdSimAmt">折後未稅金額（萬元）</label><input type="text" inputmode="decimal" id="fdSimAmt" class="fd-w100" value="1200"></div>
        <div><span class="fd-lbl" id="fdLvlLbl">毛利層級</span><div class="fd-lvl" role="radiogroup" aria-labelledby="fdLvlLbl">
          <label><input type="radio" name="fdSimLvl" value="1" checked><span>1　一般</span></label><label><input type="radio" name="fdSimLvl" value="2"><span>2　偏低</span></label><label><input type="radio" name="fdSimLvl" value="3"><span>3　很低</span></label></div></div>
        <div><label class="fd-lbl" for="fdSimPct">或直接輸入毛利率（%，選填）</label><input type="text" inputmode="decimal" id="fdSimPct" class="fd-w100" placeholder="例：18.5（有填就以此為準）"></div>
        <div><label class="fd-lbl" for="fdSimCat">商品歸類</label><select id="fdSimCat" class="fd-w100"></select></div>
        <div><label class="fd-lbl" for="fdSimBu">業務所屬 BU</label><select id="fdSimBu" class="fd-w100"></select></div>
        <div class="fd-simchecks"><span class="fd-chkline"><input type="checkbox" id="fdSimNoMgr2"><label for="fdSimNoMgr2">這位業務找不到直屬二級主管</label></span>
          <span class="fd-chkline" id="fdSimExcWrap" hidden><input type="checkbox" id="fdSimExc"><label for="fdSimExc">其二級主管在「跳過一級主管」例外名單內</label></span>
          <span class="fd-chkline"><input type="checkbox" id="fdSimCmp"><label for="fdSimCmp">與目前生效的預設流程比較</label></span></div>
      </form>
      <div class="fd-row"><button type="button" class="fd-btn primary" id="fdSimRun">模擬</button><button type="button" class="fd-btn" id="fdSimScan">邊界掃描（1,000 萬／5,000 萬及所有金額字面值 ±0.01）</button><button type="button" class="fd-btn sm" id="fdSimClear">清除高亮</button>
        <span class="fd-mini fd-muted">快速帶入：</span>
        <button type="button" class="fd-btn sm" data-preset="800,1">800 萬・一般</button><button type="button" class="fd-btn sm" data-preset="1200,1">1,200 萬</button><button type="button" class="fd-btn sm" data-preset="300,2">300 萬・層級 2</button><button type="button" class="fd-btn sm" data-preset="6000,1">6,000 萬</button></div>
      <div class="fd-simout" id="fdSimOut" aria-live="polite"></div>
    </div>
  </section>
  <section id="fdTab-val" role="tabpanel" aria-labelledby="fdTabBtn-val" hidden><div class="fd-panel" id="fdVal"></div></section>
  <section id="fdTab-ov" role="tabpanel" aria-labelledby="fdTabBtn-ov" hidden>
    <div class="fd-panel">
      <h2 class="fd-title">總覽：自由畫布（唯讀）</h2>
      <p class="fd-help">把目前草稿畫成流程圖：最左「送簽」→ 由上而下依序判斷各路線的條件 → 符合的路線接該路線的關卡 → 「核准完成」。滾輪縮放、拖曳平移；Tab 可逐一檢視節點。這裡不能編輯——要改請回「設計器」。</p>
      <div class="fd-ovbar"><button type="button" class="fd-btn sm" id="fdOvFit">縮放至適合</button><button type="button" class="fd-btn sm" id="fdOvOut" aria-label="縮小">－</button><button type="button" class="fd-btn sm" id="fdOvIn" aria-label="放大">＋</button><span class="fd-mini fd-muted" id="fdOvZ" aria-live="off"></span>
        <span class="fd-chkline"><input type="checkbox" id="fdOvSim" checked><label for="fdOvSim">套用模擬輸入高亮命中的路線</label></span><span class="fd-mini fd-muted" id="fdOvInfo" role="status"></span></div>
      <div class="fd-ovwrap" id="fdOvHost"></div>
      <h3>文字大綱</h3><ol class="fd-outline" id="fdOutline"></ol>
    </div>
  </section>
  <section id="fdTab-drafts" role="tabpanel" aria-labelledby="fdTabBtn-drafts" hidden><div class="fd-panel" id="fdDrafts"></div></section>
  <div id="fdLive" class="fd-sr" aria-live="polite"></div>
  <div id="fdToast" class="fd-toast" role="status"></div>
  <div id="fdGhost" class="fd-ghost" aria-hidden="true"></div>
  <dialog id="fdDlg" class="fd-dlg" aria-labelledby="fdDlgT"></dialog>
</div>`);
    put($('#fdSimCat'), h`${CATS.map(([k, n]) => h`<option value="${k}">${n}</option>`)}`);
    put($('#fdSimBu'), h`${BUS.map((b) => h`<option value="${b}">${b}</option>`)}`);
  }

  // ───────────────────────── 渲染：頂列／狀態 ─────────────────────────
  function renderBar() {
    const ed = canEdit(), d = S.draft;
    put($('#fdBar'), h`
      <label class="fd-lbl" for="fdName">草稿名稱</label><input type="text" id="fdName" class="fd-name" maxlength="40" value="${d.name}" ${ed ? '' : raw('readonly')} aria-describedby="fdNameHelp">
      <label class="fd-lbl" for="fdNote">備註</label><input type="text" id="fdNote" class="fd-name" maxlength="200" value="${d.note}" placeholder="選填" ${ed ? '' : raw('readonly')}>
      <button type="button" class="fd-btn primary fd-edit-only" id="fdSave" ${ed && S.def ? '' : raw('disabled')}>儲存草稿</button>
      <button type="button" class="fd-btn fd-edit-only" id="fdSaveAs" ${ed && S.def ? '' : raw('disabled')}>另存新草稿</button>
      <button type="button" class="fd-btn fd-edit-only" id="fdNew" ${ed ? '' : raw('disabled')} title="丟掉目前編輯中的內容，從目前設定轉出的預設流程重新開始">從預設流程開始新草稿</button>
      <button type="button" class="fd-btn fd-edit-only" id="fdRevert" ${ed ? '' : raw('disabled')} title="把編輯中的內容還原成預設流程">還原成預設流程</button>
      <span class="fd-mini fd-muted" id="fdNameHelp">儲存時後端會再完整驗證一次（不只靠前端）。</span>`);
    const rc = S.rc;
    const e = rc ? rc.checks.filter((c) => c.status === 'error').length : 0, w = rc ? rc.checks.filter((c) => c.status === 'warn').length : 0;
    const dirty = isDirty();
    put($('#fdStatusPills'), h`
      <span class="fd-pill ${d.id ? (dirty ? 'cur' : 'ok') : 'gray'}" id="fdSavedPill">${d.id ? (dirty ? '● 有未儲存的變更' : '✓ 草稿已儲存') : (S.def ? '○ 尚未儲存的新草稿' : '○ 尚未載入')}</span>
      ${rc ? h`<span class="fd-pill ${e ? 'err' : w ? 'warn' : 'ok'}" id="fdValPill">${e ? '✕ 驗證 ' + e + ' 個錯誤' : w ? '▲ 驗證 ' + w + ' 個警告' : '✓ 驗證通過'}</span>` : ''}
      ${S.rejects ? h`<span class="fd-pill err" title="被保護底線拒絕的編輯次數" id="fdRejPill">底線拒絕 ${S.rejects} 次</span>` : ''}`);
    const cv = $('#fdCnt-val');
    if (cv) { const n = e + w; cv.hidden = !n; cv.textContent = String(n); cv.className = 'fd-cnt' + (e ? '' : ' warn'); }
  }

  // ───────────────────────── 渲染：左側面板 ─────────────────────────
  function renderPalette() {
    const el = $('#fdPalette');
    const ed = canEdit();
    el.hidden = !S.def || isMobile();
    $('#fdGrid').classList.toggle('no-pal', el.hidden);
    if (el.hidden) return;
    const def = S.def;
    const nodes = TIERS.map((t) => {
      const T = TIER[t], nd = nodeOfTier(def, t);
      if (!nd) return '';
      return h`<div class="fd-pal fd-c-${colorOfTier(def, t)}" ${ed ? raw('tabindex="0" role="button"') : raw('aria-disabled="true"')} data-pal="node" data-tier="${t}" ${ed ? h`data-drag="pnode:${t}"` : ''} aria-label="關卡節點：${nd.label}${T.locked ? '（鎖定）' : ''}。${ed ? '用滑鼠拖進路線，或按 Enter 選擇要加入哪條路線' : '目前唯讀'}"><span class="fd-gl">${T.locked ? '🔒' : raw('<i class="fd-sq"></i>')}</span><div><div class="fd-nm">${nd.label}</div><div class="fd-ds">${T.rule}</div></div></div>`;
    });
    const conds = VAR_ORDER.map((v) => h`<div class="fd-pal cond" ${ed ? raw('tabindex="0" role="button"') : raw('aria-disabled="true"')} data-pal="cond" data-var="${v}" ${ed ? h`data-drag="pcond:${v}"` : ''} aria-label="條件卡：${VARS[v].label}。${ed ? '拖到路線起點的條件膠囊，或按 Enter 選擇路線' : '目前唯讀'}"><span class="fd-gl">◆</span><div><div class="fd-nm">${VARS[v].label}</div><div class="fd-ds">${VAR_DESC[v]}</div></div></div>`);
    put(el, h`<div class="fd-grp"><h3>關卡節點</h3><p class="fd-help">拖進右邊的路線。只有這 5 種，<b>不能新增新種類</b>（🔒 代核關可放進路線，但不能改它的簽核方式）。</p>${nodes}</div>
      <div class="fd-grp"><h3>條件卡</h3><p class="fd-help">拖到路線最左邊的條件膠囊，會加進那條路線的條件（預設路線不能加）。</p>${conds}</div>`);
  }

  // ───────────────────────── 渲染：畫布 ─────────────────────────
  function simInfo() {
    const r = S.simRes;
    return r && r.route ? r : null;
  }
  function nodeCard(def, r, i, id, sim) {
    const tier = tierOfId(def, id), T = TIER[tier], nd = nodeById(def, id), ed = canEdit();
    let ord = '', cls = '';
    if (sim && sim.route.id === r.id) { const k = sim.tiers.indexOf(tier); if (k >= 0) { cls = ' hit-n'; ord = h`<span class="fd-ord" aria-hidden="true">${k + 1}</span>`; } }
    else if (sim) cls = ' dimn';
    return h`<li class="fd-node fd-c-${colorOfTier(def, tier)}${T.locked ? ' locked' : ''}${cls}" tabindex="0" ${ed ? h`data-drag="node:${r.id}:${i}"` : ''} data-lane="${r.id}" data-idx="${i}" data-tier="${tier}" aria-label="路線「${r.name}」第 ${i + 1} 關：${nd.label}${T.locked ? '（鎖定的代核關）' : ''}。${ed ? 'Enter 開屬性；左右方向鍵調整順序，上下方向鍵換路線；Delete 移除' : '唯讀'}">
      ${ord}<div class="fd-top2"><span class="fd-stepno">STEP ${i + 1}</span>${T.locked ? h`<span title="${T.lockWhy}" aria-label="鎖定">🔒</span>` : ''}<button type="button" class="fd-x" data-act="del-node" aria-label="從路線移除「${nd.label}」" tabindex="-1">×</button></div>
      <div class="fd-nm">${nd.label}</div><div class="fd-rl">${T.rule}</div>${T.locked ? h`<div class="fd-rl">🔒 鎖定：語意待董事長溝通</div>` : ''}</li>`;
  }
  function renderCanvas() {
    const el = $('#fdCanvas');
    if (!S.def) { put(el, h`<h2 id="fdCanvasH" class="fd-title">簽核路線畫布</h2><p class="fd-help">${S.busy ? '載入中…' : '尚未載入草稿。請到「草稿清單」載入既有草稿，或按上方「從預設流程開始新草稿」。'}</p>`); return; }
    const d = S.def, ed = canEdit(), sim = simInfo();
    const v = S.view || { routes: [], modifiers: [] };
    const vr = (id) => v.routes.find((x) => x.id === id) || { whenText: '', isDefault: false };
    const lanes = d.routes.map((r, ri) => {
      const isDef = r.when === true;
      const hit = !!(sim && sim.route.id === r.id);
      const cards = r.path.map((id, i) => nodeCard(d, r, i, id, sim));
      const slot = r.path.length ? h`<li class="fd-slot" aria-hidden="true">＋ 拖入關卡</li>` : h`<li class="fd-slot empty" role="alert">⚠ 這條路線沒有任何關卡（至少要一關）。把左邊的關卡拖進來</li>`;
      const inserted = hit ? sim.steps.filter((s) => s.inserted) : [];
      const insNote = inserted.length ? h`<span class="fd-chip add" title="由調整規則插入">＋ ${inserted.map((s) => s.label).join('、')}（調整規則插入）</span>` : '';
      const wt = vr(r.id).whenText;
      return h`<li class="fd-lane${isDef ? ' default' : ''}${hit ? ' hit' : ''}" data-lane="${r.id}" data-i="${ri}" aria-label="優先 ${ri + 1}：${r.name}">
        <div class="fd-lhead">
          <button type="button" class="fd-grip" ${ed ? h`data-drag="lane:${r.id}"` : raw('disabled')} aria-label="拖曳調整路線「${r.name}」的優先順序（目前第 ${ri + 1}；也可按上下方向鍵）">⠿</button>
          <div class="fd-lmeta"><div><span class="fd-pri">優先 ${ri + 1}</span><span class="fd-lname">${r.name}</span>${hit ? h` <span class="fd-pill cur">● 本單走這條</span>` : ''}</div>
            <div class="fd-cond${isDef ? ' dflt' : ''}" ${isDef ? '' : h`tabindex="0" role="button" aria-label="路線條件：${wt}。雙擊或按 Enter 開啟條件編輯"`} data-cond="${r.id}">${wt}</div>
            <div class="fd-ltools fd-edit-only">${ed ? h`<button type="button" class="fd-btn sm" data-act="lane-up" data-lane="${r.id}" aria-label="路線「${r.name}」優先順序上移" ${ri === 0 ? raw('disabled') : ''}>▲</button><button type="button" class="fd-btn sm" data-act="lane-down" data-lane="${r.id}" aria-label="路線「${r.name}」優先順序下移" ${ri >= d.routes.length - 2 ? raw('disabled') : ''}>▼</button>${isDef ? '' : h`<button type="button" class="fd-btn sm danger" data-act="lane-del" data-lane="${r.id}" aria-label="刪除路線「${r.name}」">刪除</button>`}<button type="button" class="fd-btn sm" data-act="lane-edit" data-lane="${r.id}" aria-label="編輯路線「${r.name}」">屬性</button>` : ''}</div>
          </div>
        </div>
        <ol class="fd-path" aria-label="路線「${r.name}」的關卡（依序）">${cards}${slot}</ol>
        <div class="fd-lend-wrap"><span class="fd-lend">✔ 完成核准</span> ${insNote}</div>
      </li>`;
    });
    const hitMods = new Set(sim ? sim.applied.map((a) => a.id) : []);
    const mods = d.modifiers.map((m) => {
      const mv = v.modifiers.find((x) => x.id === m.id) || { kind: 'other', whenText: '' };
      const hit = hitMods.has(m.id);
      const tn = nodeById(d, m.target), nn = m.node ? nodeById(d, m.node) : null;
      const act = { insert_before: h`在「${tn ? tn.label : m.target}」之前插入「${nn ? nn.label : m.node}」`, insert_after: h`在「${tn ? tn.label : m.target}」之後插入「${nn ? nn.label : m.node}」`, remove: h`移除「${tn ? tn.label : m.target}」`, replace: h`把「${tn ? tn.label : m.target}」換成「${nn ? nn.label : m.node}」` }[m.op];
      return h`<div class="fd-mod${hit ? ' hit' : ''}" tabindex="0" role="button" data-mod="${m.id}" aria-label="調整規則：${m.name}。雙擊或按 Enter 開啟屬性">
        <div class="fd-mod-main"><div><span class="fd-pill proxy">⚙ 調整規則</span> <b>${m.name}</b>${hit ? h` <span class="fd-pill cur">● 本單套用</span>` : ''}</div>
          <div class="fd-mini">如果 ${mv.kind === 'bu' ? h`「業務所屬 BU」屬於 ${mv.bus.length ? mv.bus.map((b) => h`<span class="fd-chip k">${b}</span>`) : h`<span class="fd-chip del">（未選）</span>`} 且找得到直屬二級主管` : mv.whenText} → ${act}</div></div>
        <div class="fd-mini fd-muted">套用在路線選定之後。${mv.kind === 'bu' ? '找不到二級主管時不插入、也不會擋單（現行規則）。' : ''}</div></div>`;
    });
    $('#fdCanvas').classList.toggle('ro', !ed);
    put(el, h`
      <div class="fd-ctop"><div><h2 id="fdCanvasH">簽核路線畫布　<span class="fd-pill gray">${S.draft.name || '新草稿'}</span></h2>
        <p class="fd-help">單子送簽時，由上而下找<b>第一條符合條件</b>的路線，就依那條路線的關卡順序簽。最後一條是「其他情況（預設）」，不可刪除。</p></div>
        <div class="fd-row fd-edit-only">${ed ? h`<button type="button" class="fd-btn sm" id="fdAddLane">＋ 新增路線</button>` : ''}</div></div>
      <div class="fd-banner" id="fdBanner" role="alert" ${S.banner ? '' : raw('hidden')}><span aria-hidden="true">⛔</span><div>${S.banner ? raw(S.banner) : ''}</div><button type="button" class="fd-bx" id="fdBannerX" aria-label="關閉訊息">×</button></div>
      <div class="fd-trunk"><span class="fd-start">送簽</span><span aria-hidden="true">→</span><span class="fd-decision">條件判斷　由上而下，第一條符合者勝出</span></div>
      <ol class="fd-lanes${sim ? ' simdim' : ''}" id="fdLanes">${lanes}</ol>
      <div class="fd-mods"><h3>調整規則（路線選定後再套用）</h3>${mods.length ? mods : h`<p class="fd-help">沒有調整規則。</p>`}</div>
      <p class="fd-hint">操作：拖曳卡片排序／換路線；<b>雙擊</b>卡片、條件膠囊或調整規則開屬性；鍵盤：Tab 選卡片、←→ 調順序、↑↓ 換路線、Enter 開屬性、Delete 移除、Esc 關閉。</p>`);
    if (S.focusKey) {
      const fk = S.focusKey; S.focusKey = null;
      const p = fk.split(':');
      let f = null;
      if (p[0] === 'node') f = findNode(p[1], +p[2]); else if (p[0] === 'grip') { const ln = findLane(p[1]); f = ln && $('.fd-grip', ln); }
      if (f) f.focus();
    }
  }
  function findLane(id) { return $$('#fdLanes .fd-lane').find((e) => e.dataset.lane === id) || null; }
  function findNode(l, i) { return $$('#fdLanes .fd-node').find((e) => e.dataset.lane === l && +e.dataset.idx === i) || null; }
  function applySel() {
    $$('#fdLanes .fd-node.sel').forEach((n) => n.classList.remove('sel'));
    if (S.sel) { const e = findNode(S.sel.lane, S.sel.idx); if (e) e.classList.add('sel'); }
  }

  // ───────────────────────── 變更入口：所有編輯都走這裡 ─────────────────────────
  // 先在複本上做變更 → 問後端驗證（11 條＋保護底線）→ 通過才換掉 S.def。違反底線、或結構無效，一律拒絕（畫面狀態不會進入無效定義）。
  function setBanner(r) { S.banner = r ? (r instanceof Raw ? r.s : esc(r)) : ''; }
  function floorMsg(vs) {
    const v = vs && vs[0];
    if (!v) return h`<b>不能套用（保護底線）</b>`;
    const amt = fmtYuan(String(BigInt(v.revenueCents) / 100n) + (BigInt(v.revenueCents) % 100n ? '.' + String(BigInt(v.revenueCents) % 100n).padStart(2, '0') : ''));
    return h`<b>不能套用（保護底線）：</b>這個變更會讓「金額 ${amt}、${CAT_ZH[v.category] || v.category}、${v.ownerBu}」這類單<b>沒有經過${v.need}</b>。<br>底線規則：金額 >1,000 萬必含董事長（或董事會代核）、>5,000 萬必含董事會代核；其他部分都可以自由調整。要改底線得走程式變更並由稽核簽認。`;
  }
  function firstError(rc) {
    for (const c of (rc && rc.checks) || []) for (const i of c.issues || []) if (i.level !== 'warn') return i.text;
    return '流程定義無效';
  }
  function tryApply(mut, desc, opts) {
    const run = () => doApply(mut, desc, opts || {}).catch((e) => { toast('編輯失敗：' + (e && e.message ? e.message : e)); return { ok: false, msg: String(e && e.message || e), plain: true }; });
    S.queue = S.queue.then(run, run);
    return S.queue;
  }
  async function doApply(mut, desc, opts) {
    if (!canEdit()) { toast('目前不能編輯（' + (isMobile() ? '手機為唯讀' : !S.def ? '尚未載入草稿' : '此環境不開放編輯') + '）'); return { ok: false, msg: 'locked' }; }
    const d = clone(S.def);
    const r = mut(d);
    if (r && r.error) { if (!opts.silent) { setBanner(h`<b>無法套用：</b>${r.error}`); renderCanvas(); } return { ok: false, msg: r.error, plain: true }; }
    let res;
    try { res = await api('POST', '/designer/validate', { def: d }); } catch (e) { toast('連線失敗，無法驗證這個變更'); return { ok: false, msg: '連線失敗', plain: true }; }
    if (!res.ok || !res.data) { const m = (res.data && res.data.error) || '驗證失敗'; setBanner(h`<b>無法套用：</b>${m}`); renderCanvas(); return { ok: false, msg: m, plain: true }; }
    const rc = res.data;
    if (!rc.structural) {
      const m = firstError(rc);
      if (opts.from !== 'drawer') { setBanner(h`<b>無法套用：</b>${m}`); renderCanvas(); announce('已拒絕：' + m); }
      return { ok: false, msg: m, plain: true };
    }
    if (rc.floor && rc.floor.ok === false) {
      S.rejects++;
      const msg = floorMsg(rc.floor.violations);
      if (opts.from === 'drawer') { renderBar(); return { ok: false, msg: msg.s, html: true }; }
      setBanner(msg); renderAll(); announce('已拒絕：違反保護底線。' + (desc || ''));
      return { ok: false, msg: msg.s, html: true };
    }
    S.def = d; S.view = rc.view; S.checks = rc.checks; S.rc = rc; setBanner(null);
    S.scan = null;   // 邊界掃描是針對編輯前的內容，編輯後作廢
    if (S.simIn) await runSimQuiet();
    if (!opts.noRender) renderAll();
    announce(desc || '已更新草稿');
    return { ok: true };
  }

  // ───────────────────────── 具體編輯動作 ─────────────────────────
  function actMoveNode(fl, fi, tl, ti, why) {
    return tryApply((d) => {
      const f = laneOf(d, fl), t = laneOf(d, tl);
      if (!f || !t) return { error: '找不到路線' };
      const [n] = f.path.splice(fi, 1);
      let idx = ti; if (f === t && fi < ti) idx = ti - 1;
      idx = Math.max(0, Math.min(idx, t.path.length));
      t.path.splice(idx, 0, n);
      S.focusKey = 'node:' + tl + ':' + idx;
    }, why || '已移動關卡');
  }
  async function actInsertNode(tl, ti, tier) {
    const r = await tryApply((d) => {
      const t = laneOf(d, tl); const nd = nodeOfTier(d, tier);
      if (!t || !nd) return { error: '找不到路線或關卡' };
      if (t.path.includes(nd.id)) return { error: `路線「${t.name}」已經有「${nd.label}」，同一個關卡不能出現兩次（同一人不能簽兩關）` };
      const at = Math.max(0, Math.min(ti, t.path.length));
      t.path.splice(at, 0, nd.id);
      S.focusKey = 'node:' + tl + ':' + at;
    }, '已把「' + labelOfTier(S.def, tier) + '」加入路線');
    if (r.ok) toast(`已把「${labelOfTier(S.def, tier)}」加入「${laneOf(S.def, tl).name}」`);
    return r;
  }
  async function actDelNode(l, i) {
    const d = S.def, r = laneOf(d, l);
    if (!r) return;
    const nd = nodeById(d, r.path[i]);
    const ok = await confirmBox('移除關卡', h`<p>要把「${nd.label}」從路線「${r.name}」移除嗎？</p><p class="fd-help">這只是從這條路線拿掉，關卡種類本身還在左側面板。若會違反保護底線或讓路線變空，系統會拒絕。</p>`, '移除', true);
    if (!ok) return;
    const res = await tryApply((dd) => { laneOf(dd, l).path.splice(i, 1); S.focusKey = null; }, '已移除「' + nd.label + '」');
    if (res.ok) toast('已移除關卡');
  }
  function actMoveLane(id, dir) {
    return tryApply((d) => {
      const i = d.routes.findIndex((r) => r.id === id), j = i + dir;
      if (d.routes[i].when === true) return { error: '預設路線固定在最後' };
      if (j < 0 || j >= d.routes.length - 1) return { error: '已經到邊界（預設路線固定在最後）' };
      const [x] = d.routes.splice(i, 1); d.routes.splice(j, 0, x);
      S.focusKey = 'grip:' + id;
    }, '路線優先順序已調整');
  }
  async function actDelLane(id) {
    const r = laneOf(S.def, id);
    if (!r || r.when === true) return;
    const ok = await confirmBox('刪除路線', h`<p>要刪除路線「${r.name}」嗎？它的條件與關卡都會一起移除。</p><p class="fd-help">若刪除後會違反保護底線，系統會拒絕。</p>`, '刪除', true);
    if (!ok) return;
    const res = await tryApply((d) => { d.routes = d.routes.filter((x) => x.id !== id); }, '已刪除路線');
    if (res.ok) { toast('已刪除路線「' + r.name + '」'); if (S.drawer && S.drawer.ref === id) closeDrawer(); }
  }
  async function actAddLane() {
    const res = await tryApply((d) => {
      const id = newRouteId();
      const m1 = nodeOfTier(d, 'mgr1'), gm = nodeOfTier(d, 'gm');
      if (d.routes.length >= 30) return { error: '路線數量已達上限' };
      d.routes.splice(d.routes.length - 1, 0, { id, name: '新路線 ' + d.routes.length, when: { cmp: ['ownerBuPrimary', '==', 'MDM'] }, path: [m1.id, gm.id] });
      S.focusKey = 'grip:' + id;
    }, '已新增路線');
    if (res.ok) toast('已在「預設路線」之前新增一條路線（MDM 業務加簽總經理），請雙擊條件膠囊調整');
  }
  async function actAddCond(laneId, v) {
    const r0 = laneOf(S.def, laneId);
    if (!r0 || r0.when === true) { toast('預設路線不能加條件：它是「其他情況」'); return; }
    const cur = astToEdit(r0.when);
    if (cur && cur.unsupported) { toast('這條路線的條件超出編輯器範圍（巢狀或否定），無法加入條件卡'); return; }
    const res = await tryApply((d) => {
      const r = laneOf(d, laneId); const e = astToEdit(r.when);
      if (e.items.length >= 6) return { error: '條件最多 6 個' };
      e.items.push(defaultItem(v));
      r.when = editToAst(e);
    }, '已加入條件：' + VARS[v].label);
    if (res.ok) { toast('已把「' + VARS[v].label + '」加進路線條件，可在右側調整'); openRouteDrawer(laneId); }
  }

  // ───────────────────────── 右側抽屜 ─────────────────────────
  let drawerOpener = null;
  function closeDrawer() {
    S.drawer = null; $('#fdDrawer').hidden = true; $('#fdGrid').classList.remove('drawer-open');
    if (drawerOpener && doc.contains(drawerOpener)) drawerOpener.focus();
    drawerOpener = null;
  }
  function showDrawer() { S.sel = null; applySel(); $('#fdDrawer').hidden = false; $('#fdGrid').classList.add('drawer-open'); }
  function openNodeDrawer(l, i) {
    if (isMobile()) { toast('手機為唯讀檢視，請改用桌機編輯屬性'); return; }
    const r = laneOf(S.def, l); if (!r) return;
    const nd = nodeById(S.def, r.path[i]);
    drawerOpener = doc.activeElement;
    S.drawer = { kind: 'node', ref: nd.tier, laneId: l, idx: i, edit: { label: nd.label, color: colorOfTier(S.def, nd.tier) }, err: '' };
    renderDrawer(); showDrawer();
    const f = $('#fdDrLabel') || $('#fdDrawer .fd-close'); if (f) f.focus();
  }
  function openRouteDrawer(l) {
    if (isMobile()) { toast('手機為唯讀檢視，請改用桌機編輯條件'); return; }
    const r = laneOf(S.def, l); if (!r) return;
    drawerOpener = doc.activeElement;
    S.drawer = { kind: 'route', ref: l, edit: { name: r.name, when: astToEdit(r.when) }, err: '' };
    renderDrawer(); showDrawer();
    const f = $('#fdDrName'); if (f) f.focus();
  }
  function openModDrawer(id) {
    if (isMobile()) { toast('手機為唯讀檢視，請改用桌機編輯'); return; }
    const m = S.def.modifiers.find((x) => x.id === id); if (!m) return;
    const mv = (S.view && S.view.modifiers.find((x) => x.id === id)) || { kind: 'other' };
    drawerOpener = doc.activeElement;
    S.drawer = { kind: 'mod', ref: id, edit: { bus: mv.kind === 'bu' ? mv.bus.slice() : [] }, err: '' };
    renderDrawer(); showDrawer();
    const f = $('#fdDrawer .fd-close'); if (f) f.focus();
  }
  function fieldCaption(t) { return h`<span class="fd-lbl">${t}</span>`; }
  function condBlock(it, i, n, mode) {
    const vd = VARS[it.v];
    const vsel = VAR_ORDER.map((k) => h`<option value="${k}" ${k === it.v ? raw('selected') : ''}>${VARS[k].label}</option>`);
    const osel = vd.ops.map((o) => h`<option value="${o}" ${o === it.o ? raw('selected') : ''}>${OP_ZH[o]}</option>`);
    let val;
    if (vd.type === 'money') val = h`<input type="text" inputmode="decimal" data-f="wan" value="${it.wan == null ? '' : it.wan}" aria-label="金額（萬元）"> 萬元`;
    else if (vd.type === 'pct') val = h`<input type="text" inputmode="decimal" data-f="val" class="fd-short" value="${it.val == null ? '' : it.val}" aria-label="毛利率（%）"> %`;
    else if (vd.type === 'level') val = h`<select data-f="val" aria-label="毛利層級">${[1, 2, 3].map((x) => h`<option value="${x}" ${it.val === x ? raw('selected') : ''}>${LVL_ZH[x]}</option>`)}</select>`;
    else if (vd.type === 'set') val = h`<select data-f="val" aria-label="${vd.label}">${CLASS_KEYS.map(([k, nm]) => h`<option value="${k}" ${it.val === k ? raw('selected') : ''}>${nm}</option>`)}</select>`;
    else if (vd.type === 'string') val = (it.o === 'in' || it.o === 'not_in')
      ? h`<input type="text" data-f="csv" value="${Array.isArray(it.val) ? it.val.join(',') : ''}" placeholder="帳號，以逗號分隔" aria-label="帳號清單">`
      : h`<input type="text" data-f="val" value="${it.val == null ? '' : it.val}" placeholder="帳號" aria-label="指定業務帳號">`;
    else {
      const opts = it.v === 'category' ? CATS : BUS.map((b) => [b, b]);
      if (it.o === 'in' || it.o === 'not_in') val = h`<span class="fd-vals" role="group" aria-label="${vd.label}（可複選）">${opts.map(([k, nm]) => h`<label><input type="checkbox" data-f="chk" value="${k}" ${Array.isArray(it.val) && it.val.includes(k) ? raw('checked') : ''}>${nm}</label>`)}</span>`;
      else val = h`<select data-f="val" aria-label="${vd.label}">${opts.map(([k, nm]) => h`<option value="${k}" ${it.val === k ? raw('selected') : ''}>${nm}</option>`)}</select>`;
    }
    return h`<div class="fd-blk" data-i="${i}"><span class="fd-w">${i === 0 ? '如果' : (mode === 'any' ? '或' : '且')}</span><select data-f="v" aria-label="變數">${vsel}</select><select data-f="o" aria-label="運算子">${osel}</select>${val}<button type="button" class="fd-btn sm danger" data-act="rm-item" data-i="${i}" aria-label="移除第 ${i + 1} 個條件" ${n <= 1 ? raw('title="至少要有一個條件"') : ''}>×</button></div>`;
  }
  function renderDrawer() {
    const D = S.drawer, el = $('#fdDrawer');
    if (!D) return;
    const d = S.def, ed = canEdit();
    let body, title;
    const errBox = h`<div class="fd-drerr fd-banner ${D.err ? 'show' : ''}" id="fdDrErr" role="alert">${D.err ? raw(D.err) : ''}</div>`;
    if (D.kind === 'node') {
      const T = TIER[D.ref], lk = T.locked, e = D.edit;
      title = '關卡屬性：' + labelOfTier(d, D.ref);
      const inRoutes = d.routes.filter((r) => r.path.includes(nodeOfTier(d, D.ref).id));
      body = h`${lk ? h`<div class="fd-lockbox"><span aria-hidden="true">🔒</span><div><b>鎖定節點</b><br>${T.lockWhy}<br><span class="fd-mini">（這一關只能決定「哪些路線包含它」，不能改名稱、節點色與簽核方式；也不能新增新種類的關卡。）</span></div></div>` : ''}
        <div class="fd-field"><label class="fd-lbl" for="fdDrLabel">顯示名稱</label><input type="text" id="fdDrLabel" class="fd-w100" data-f="label" value="${e.label}" maxlength="14" ${lk ? raw('disabled') : ''}><p class="fd-help">這個名字會出現在報價單的簽核進度卡片上。1–14 字，不可含 &lt; &gt; " ' &amp;。</p></div>
        <div class="fd-field">${fieldCaption('節點色')}<div class="fd-sw" role="group" aria-label="節點色">${COLORS.map((c) => h`<button type="button" class="fd-c-${c}" data-color="${c}" aria-pressed="${e.color === c}" aria-label="${COLOR_ZH[c]}" ${lk ? raw('disabled') : ''}></button>`)}</div><p class="fd-help">只影響畫布與卡片的色條，方便辨識；不影響誰能簽。</p></div>
        <div class="fd-field">${fieldCaption('簽核人規則（唯讀）')}<div class="fd-ro">${T.ruleLong}</div><p class="fd-help"><b>名單由「名冊管理」與帳號的直屬主管欄位維護，這裡不能改。</b>設計器只決定「哪一關、什麼順序、什麼條件走」，不決定「誰」。資格判斷時機：${T.timing}。</p></div>
        <div class="fd-field">${fieldCaption('無對應簽核人時（唯讀）')}<div class="fd-ro">不能送出（擋下，請管理員先補名單）</div><p class="fd-help">其他系統（如 BPM）可選「跳過」，但<b>跳過會少一次獨立檢視</b>，內控上漏放比多擋嚴重，所以只開放「不能送出」。</p></div>
        <div class="fd-field">${fieldCaption('出現在哪些路線')}${inRoutes.length ? inRoutes.map((r) => h`<div class="fd-mini">・${r.name}（STEP ${r.path.indexOf(nodeOfTier(d, D.ref).id) + 1}）</div>`) : h`<div class="fd-mini fd-muted">目前沒有路線使用這一關</div>`}</div>
        ${errBox}
        <div class="fd-row">${ed ? h`<button type="button" class="fd-btn primary" data-act="dr-apply" ${lk ? raw('disabled') : ''}>套用</button>` : ''}<button type="button" class="fd-btn" data-act="dr-close">${ed ? '取消' : '關閉'}</button>${D.laneId && ed ? h`<button type="button" class="fd-btn danger" data-act="dr-delnode">從此路線移除</button>` : ''}</div>`;
    } else if (D.kind === 'route') {
      const r = laneOf(d, D.ref), e = D.edit;
      if (!r) { title = '路線屬性'; body = h`<p>這條路線已經不存在。</p><button type="button" class="fd-btn" data-act="dr-close">關閉</button>`; }
      else {
        const isDef = r.when === true, uns = !isDef && e.when && e.when.unsupported;
        title = isDef ? '路線屬性：其他情況（預設）' : '路線條件編輯器';
        const blocks = !isDef && !uns ? e.when.items.map((it, i) => condBlock(it, i, e.when.items.length, e.when.mode)) : '';
        const n = nodeById;
        body = h`<div class="fd-field"><label class="fd-lbl" for="fdDrName">路線名稱</label><input type="text" id="fdDrName" class="fd-w100" data-f="rname" value="${e.name}" maxlength="16"><p class="fd-help">只用來辨識這條路線（出現在報價單卡片與稽核紀錄），不影響判斷。1–16 字。</p></div>
          ${isDef ? h`<div class="fd-lockbox"><span aria-hidden="true">ℹ</span><div><b>其他情況（預設）</b><br>上面所有路線都不符合時，走這一條。它<b>固定在最後、不能刪除、不能加條件</b>，這樣每張單一定走得到某一條路線。</div></div>`
            : uns ? h`<div class="fd-lockbox"><span aria-hidden="true">ℹ</span><div><b>這條路線的條件超出編輯器範圍</b><br>它使用了巢狀或否定條件（例如由程式轉出的預設流程以外的結構）。條件造句編輯器只支援「全部符合／任一符合」一層；這條路線的條件在這裡只能檢視，可改名稱，要改條件請刪除路線後重建。<div class="fd-sentence">${S.view ? ((S.view.routes.find((x) => x.id === r.id) || {}).whenText || '') : ''}</div></div></div>`
            : h`<div class="fd-field">${fieldCaption('什麼情況走這條路線')}<p class="fd-help">像造句：「如果【金額】【大於】【1,000 萬】且【毛利層級】【等於】【3】，則走這條路線」。條件之間只能選「全部符合」或「任一符合」（一層，不做複雜巢狀，稽核好讀）。順序＝優先順序：<b>第一條符合者勝出</b>。</p>
              ${e.when.items.length >= 2 ? h`<div class="fd-joiner"><label for="fdDrMode">條件之間：</label><select id="fdDrMode" data-f="mode"><option value="all" ${e.when.mode === 'all' ? raw('selected') : ''}>全部都要符合（且）</option><option value="any" ${e.when.mode === 'any' ? raw('selected') : ''}>任一個符合就算（或）</option></select></div>` : ''}
              <div class="fd-blocks" id="fdBlocks">${blocks}</div>
              <button type="button" class="fd-btn sm" data-act="add-item" ${e.when.items.length >= 6 ? raw('disabled') : ''}>＋ 新增條件</button>
              <p class="fd-help">「毛利層級」是由核決表依商品歸類算出的 1／2／3：1＝一級主管可核、2＝需總經理、3＝需董事長。「毛利率」是實際百分比（只有管理員看得到這個頁面）。</p></div>
              <div class="fd-field">${fieldCaption('白話預覽')}<div class="fd-sentence" id="fdDrSentence" aria-live="polite"></div></div>`}
          <div class="fd-field">${fieldCaption('這條路線要簽的關卡（唯讀，請在畫布上拖拉調整）')}<div>${r.path.length ? r.path.map((id, k) => h`${k ? ' → ' : ''}<span class="fd-chip k">${n(d, id).label}</span>`) : h`<span class="fd-chip del">（空）</span>`}</div></div>
          ${errBox}
          <div class="fd-row">${ed ? h`<button type="button" class="fd-btn primary" data-act="dr-apply">套用</button>` : ''}<button type="button" class="fd-btn" data-act="dr-close">${ed ? '取消' : '關閉'}</button></div>`;
      }
    } else {
      const m = d.modifiers.find((x) => x.id === D.ref);
      const mv = (S.view && S.view.modifiers.find((x) => x.id === D.ref)) || { kind: 'other', whenText: '' };
      title = '調整規則：' + (m ? m.name : '');
      body = mv.kind === 'bu'
        ? h`<div class="fd-field">${fieldCaption('哪些 BU 的業務要多一道「二級主管複核」')}<div class="fd-vals fd-row" role="group" aria-label="涵蓋的 BU">${BUS.map((b) => h`<label class="fd-chkline"><input type="checkbox" data-f="bu" value="${b}" ${D.edit.bus.includes(b) ? raw('checked') : ''} ${ed ? '' : raw('disabled')}>${b}</label>`)}</div>
          <p class="fd-help">勾選的 BU，業務送簽時會在「一級主管」前面多一關二級主管複核（二級主管＝業務帳號的直屬主管，須為在職二級主管）。這份清單原本來自「簽核流程設定」頁的各 BU 開關，草稿可獨立調整。</p></div>
          <div class="fd-field">${fieldCaption('找不到二級主管時（唯讀）')}<div class="fd-ro">不插入這一關，也不會擋單（沿用現行規則）</div><p class="fd-help">現行系統找不到直屬二級主管時是靜默略過；這裡維持相同行為，避免新舊流程不一致。</p></div>`
        : h`<div class="fd-field">${fieldCaption('這個調整規則（唯讀）')}<div class="fd-ro">${mv.whenText}</div><p class="fd-help">${mv.kind === 'skip' ? `例外名單內有 ${mv.userCount} 人（來自「簽核流程設定」頁的例外規則）；名單不在設計器維護，也不會顯示帳號。` : '這個調整規則的結構超出編輯器範圍，這裡只能檢視。'}</p></div>`;
      body = h`${body}${errBox}<div class="fd-row">${ed && mv.kind === 'bu' ? h`<button type="button" class="fd-btn primary" data-act="dr-apply">套用</button>` : ''}<button type="button" class="fd-btn" data-act="dr-close">${ed ? '取消' : '關閉'}</button></div>`;
    }
    put(el, h`<h2 id="fdDrawerH"><span>${title}</span><button type="button" class="fd-close" data-act="dr-close" aria-label="關閉屬性抽屜（Esc）">×</button></h2>${body}`);
    refreshSentence();
  }
  function refreshSentence() {
    const D = S.drawer; if (!D || D.kind !== 'route') return;
    const sent = $('#fdDrSentence'); if (!sent) return;
    const e = D.edit;
    if (!e.when || e.when.unsupported) return;
    const nm = ($('#fdDrName') ? $('#fdDrName').value : e.name);
    const r = laneOf(S.def, D.ref);
    put(sent, h`如果 ${e.when.items.map((it, i) => h`${i ? (e.when.mode === 'any' ? ' 或 ' : ' 且 ') : ''}${itemText(it)}`)}<br>則走「${nm}」：${r ? pathNames(S.def, r.path) : ''}`);
  }
  function drawerErr(html) {
    const D = S.drawer; if (!D) return;
    D.err = html ? (html instanceof Raw ? html.s : esc(html)) : '';
    const e = $('#fdDrErr'); if (e) { e.innerHTML = D.err; e.classList.toggle('show', !!D.err); }
  }
  async function applyDrawer() {
    const D = S.drawer; if (!D) return;
    if (D.kind === 'node') {
      const lb = D.edit.label.trim();
      if (!lb || lb.length > 14 || /[<>"'&]/.test(lb)) { drawerErr(h`<span aria-hidden="true">⛔</span><div><b>名稱不合法：</b>1–14 字，不可含 &lt; &gt; " ' &amp;（FLOW_BAD_LABEL）</div>`); return; }
      if (TIER[D.ref].locked) { drawerErr(h`<span aria-hidden="true">🔒</span><div><b>鎖定節點不可修改（FLOW_LOCKED_EDIT）</b></div>`); return; }
      const r = await tryApply((d) => {
        const n = nodeOfTier(d, D.ref);
        n.label = lb;
        if (D.edit.color === TIER[D.ref].color && n.color === undefined) { /* 預設色：不寫欄位，保持內容雜湊不變 */ }
        else if (D.edit.color === TIER[D.ref].color) delete n.color;
        else n.color = D.edit.color;
      }, '已更新關卡屬性', { from: 'drawer' });
      if (r.ok) { toast('已套用關卡屬性'); closeDrawer(); } else drawerErr(h`<span aria-hidden="true">⛔</span><div>${r.html ? raw(r.msg) : r.msg}</div>`);
    } else if (D.kind === 'route') {
      const nm = D.edit.name.trim();
      if (!nm || nm.length > 16) { drawerErr(h`<span aria-hidden="true">⛔</span><div><b>路線名稱要 1–16 字</b></div>`); return; }
      const w = D.edit.when;
      if (w && !w.unsupported) {
        if (!w.items.length) { drawerErr(h`<span aria-hidden="true">⛔</span><div><b>非預設路線至少要有一個條件</b></div>`); return; }
        for (const it of w.items) { const er = itemError(it); if (er) { drawerErr(h`<span aria-hidden="true">⛔</span><div><b>${er}</b>（FLOW_BAD_COND）</div>`); return; } }
      }
      const r = await tryApply((d) => {
        const x = laneOf(d, D.ref); if (!x) return { error: '這條路線已不存在' };
        x.name = nm;
        if (w && !w.unsupported) x.when = editToAst(w);
      }, '已更新路線條件', { from: 'drawer' });
      if (r.ok) { toast('已套用路線設定'); closeDrawer(); } else drawerErr(h`<span aria-hidden="true">⛔</span><div>${r.html ? raw(r.msg) : r.msg}</div>`);
    } else {
      const r = await tryApply((d) => {
        const m = d.modifiers.find((x) => x.id === D.ref); if (!m) return { error: '找不到調整規則' };
        m.when.all[0].cmp[2] = D.edit.bus.slice().sort();
      }, '已更新調整規則', { from: 'drawer' });
      if (r.ok) { toast('已套用調整規則（BU：' + (D.edit.bus.join('、') || '無') + '）'); closeDrawer(); } else drawerErr(h`<span aria-hidden="true">⛔</span><div>${r.html ? raw(r.msg) : r.msg}</div>`);
    }
  }
  function renderDrawerKeep(sel) { renderDrawer(); const e = $(sel, $('#fdDrawer')); if (e) e.focus(); }
  function drawerInput(e) {
    const D = S.drawer; if (!D) return;
    const t = e.target, f = t.dataset.f; if (!f) return;
    if (D.kind === 'node' && f === 'label') { D.edit.label = t.value; return; }
    if (D.kind === 'mod' && f === 'bu') { const v = t.value; D.edit.bus = t.checked ? [...new Set([...D.edit.bus, v])] : D.edit.bus.filter((x) => x !== v); return; }
    if (D.kind !== 'route') return;
    if (f === 'rname') { D.edit.name = t.value; refreshSentence(); return; }
    if (f === 'mode') { D.edit.when.mode = t.value; renderDrawerKeep('#fdDrMode'); return; }
    const blk = t.closest('.fd-blk'); if (!blk) return;
    const i = +blk.dataset.i, it = D.edit.when.items[i];
    if (f === 'v') { D.edit.when.items[i] = defaultItem(t.value); renderDrawerKeep(`.fd-blk[data-i="${i}"] [data-f="v"]`); return; }
    if (f === 'o') {
      const no = t.value, vd = VARS[it.v], multi = no === 'in' || no === 'not_in', wasMulti = Array.isArray(it.val);
      if (vd.type === 'enum') { if (multi && !wasMulti) it.val = [it.val]; else if (!multi && wasMulti) it.val = it.val[0] || (it.v === 'category' ? 'consult' : 'ERP'); }
      else if (vd.type === 'string') { if (multi && !wasMulti) it.val = it.val ? [it.val] : []; else if (!multi && wasMulti) it.val = it.val[0] || ''; }
      it.o = no; renderDrawerKeep(`.fd-blk[data-i="${i}"] [data-f="o"]`); return;
    }
    if (f === 'wan') { it.wan = t.value; it.val = wanToYuan(t.value) || ''; }
    else if (f === 'val') { it.val = it.v === 'marginLevel' ? +t.value : t.value; }
    else if (f === 'csv') { it.val = t.value.split(',').map((x) => x.trim()).filter(Boolean); }
    else if (f === 'chk') { const set = new Set(Array.isArray(it.val) ? it.val : []); if (t.checked) set.add(t.value); else set.delete(t.value); it.val = [...set]; }
    refreshSentence();
  }

  // ───────────────────────── 模擬器 ─────────────────────────
  function readSim() {
    const yuan = wanToYuan($('#fdSimAmt').value);
    if (!yuan) return { error: '金額要填大於 0 的數字（單位：萬元，至多 6 位小數）' };
    const input = { amount: yuan, category: $('#fdSimCat').value, ownerBu: $('#fdSimBu').value, noMgr2: $('#fdSimNoMgr2').checked };
    const pct = $('#fdSimPct').value.trim();
    if (pct) input.marginPct = pct; else input.marginLevel = +($('input[name=fdSimLvl]:checked') || { value: 1 }).value;
    if (!$('#fdSimExcWrap').hidden && $('#fdSimExc').checked) input.mgr2Exception = true;
    return { input, compare: $('#fdSimCmp').checked };
  }
  async function runSimQuiet() {
    if (!S.def || !S.simIn) { S.simRes = null; return; }
    try {
      const r = await api('POST', '/designer/simulate', { def: S.def, input: S.simIn.input, compare: !!S.simIn.compare });
      if (r.ok) { S.simRes = r.data; S.simErr = ''; } else { S.simRes = null; S.simErr = (r.data && r.data.error) || '模擬失敗'; }
    } catch (e) { S.simRes = null; S.simErr = '連線失敗'; }
  }
  async function runSim() {
    const rd = readSim();
    if (rd.error) { S.simIn = null; S.simRes = null; S.simErr = rd.error; renderSimOut(); return; }
    S.simIn = rd; S.scan = null;
    await runSimQuiet();
    renderAll();
    announce(S.simRes ? '模擬完成：走「' + S.simRes.route.name + '」' : '模擬失敗：' + S.simErr);
  }
  function stepsHtml(steps, mark) {
    return h`<ol class="fd-steps" aria-label="簽核關卡">${steps.length ? steps.map((s, i) => h`<li class="fd-c-${colorOfTier(S.def, s.tier)} ${mark && mark[i] ? mark[i] : ''}"><span class="fd-n">STEP ${i + 1}${s.inserted ? '　＋調整插入' : ''}${s.locked ? '　🔒代核' : ''}${mark && mark[i] === 'st-add' ? '　＋新增' : ''}${mark && mark[i] === 'st-del' ? '　－草稿拿掉' : ''}</span><b>${s.label}</b>${s.approver ? h`<span class="fd-mini fd-muted">${s.approver}</span>` : ''}</li>`) : h`<li><b>（沒有任何關卡）</b></li>`}</ol>`;
  }
  const opText = (t) => ({ insert_before: `在「${t.target}」之前插入「${t.node}」`, insert_after: `在「${t.target}」之後插入「${t.node}」`, remove: `移除「${t.target}」`, replace: `把「${t.target}」換成「${t.node}」` }[t.op] || '');
  function whyHtml(res) {
    const rows = res.trace.map((t) => {
      if (t.kind === 'route') {
        if (!t.evaluated) return h`<li class="skip">優先 ${t.order}「${t.name}」：前面已經命中，不再檢查。</li>`;
        if (t.isDefault) return h`<li class="${t.matched ? 'y' : 'n'}">${t.matched ? '✓' : '✕'} 優先 ${t.order}「${t.name}」：上面都不符合 → <b>走預設路線</b>。</li>`;
        return h`<li class="${t.matched ? 'y' : 'n'}">${t.matched ? raw('✓ <b>命中</b>') : '✕ 不符合'}　優先 ${t.order}「${t.name}」：${t.atoms.map((a) => a.text + (a.ok ? ' ✓' : ' ✕')).join('；')}${t.matched ? raw('。<b>第一條符合者勝出，後面不再檢查。</b>') : ''}</li>`;
      }
      return h`<li class="${t.applied ? 'y' : 'n'}">${t.applied ? '✓' : '—'} 調整規則「${t.name}」：${t.atoms.map((a) => a.text + (a.ok ? ' ✓' : ' ✕')).join('；')} → ${t.applied ? opText(t) : '不套用'}。</li>`;
    });
    return h`<ul class="fd-why">${rows}</ul>`;
  }
  function renderSimOut() {
    const out = $('#fdSimOut'); if (!out) return;
    if (S.simErr && !S.simRes) { put(out, h`<div class="fd-note err" role="alert"><b>無法模擬</b><span>${S.simErr}</span></div>`); return; }
    if (!S.simRes) { put(out, ''); return; }
    const R = S.simRes, I = R.input;
    const bl = R.blockers && R.blockers.length ? h`<div class="fd-note warn" role="alert"><b>送簽會被擋下</b><span>${R.blockers.map((b) => b.code + '：' + b.message).join('；')}</span></div>` : '';
    let cmp = '';
    if (R.against) {
      const A = R.against;
      const dm = R.steps.map((s) => (s.addedVsActive ? 'st-add' : ''));
      const am = A.steps.map((s) => (s.removedInDraft ? 'st-del' : ''));
      cmp = h`<h3>與目前生效的預設流程比較　<span class="fd-pill ${A.same ? 'ok' : A.loose ? 'err' : 'warn'}">${A.same ? '✓ 結果相同' : A.loose ? '▲ 差異：放寬（草稿較鬆）' : '△ 差異：草稿不同'}</span></h3>
        <div class="fd-cmpgrid"><div><b>目前生效的預設流程</b>　<span class="fd-mini fd-muted">走「${A.route.name}」</span>${stepsHtml(A.steps.map((s) => ({ tier: s.tier, label: s.label })), am)}</div>
        <div><b>草稿</b>　<span class="fd-mini fd-muted">走「${R.route.name}」</span>${stepsHtml(R.steps.map((s) => ({ tier: s.tier, label: s.label, inserted: s.inserted, locked: s.locked })), dm)}</div></div>
        ${A.steps.some((s) => s.removedInDraft) ? h`<p class="fd-mini">－ 虛線＋紅底的關卡是「草稿拿掉的關卡」（生效版有、草稿沒有）；＋ 綠底是草稿新增的。</p>` : ''}`;
    }
    put(out, h`
      <h3>走【優先 ${R.route.order}：${R.route.name}】，共 ${R.steps.length} 關</h3>
      <p class="fd-help">假想單：金額 ${fmtYuan(I.amount)}、毛利層級 ${I.marginLevel}（毛利率 ${I.marginPct}）、${I.categoryLabel}、BU ${I.ownerBu}${I.noMgr2 ? '、找不到二級主管' : ''}${I.mgr2Exception ? '、二級主管在例外名單內' : ''}。</p>
      ${bl}${stepsHtml(R.steps)}
      ${R.floor && !R.floor.ok ? h`<div class="fd-note err" role="alert"><b>違反保護底線</b><span>${R.floor.limit}的單沒有經過${R.floor.need}（這份草稿不能儲存）。</span></div>` : ''}
      <h3>為什麼走這條？</h3>${whyHtml(R)}${cmp}${S.scan ? scanHtml() : ''}`);
  }
  async function runScan() {
    const rd = readSim();
    const base = rd.error ? { amount: '1', marginLevel: 1, category: $('#fdSimCat').value, ownerBu: $('#fdSimBu').value } : Object.assign({}, rd.input);
    if (!S.def) return;
    const r = await api('POST', '/designer/sweep', { def: S.def, input: base });
    if (!r.ok) { S.simErr = (r.data && r.data.error) || '掃描失敗'; S.scan = null; renderSimOut(); return; }
    S.scan = r.data; S.simErr = '';
    if (!S.simIn && !rd.error) S.simIn = rd;
    if (S.simIn) await runSimQuiet();
    renderAll();
    announce('邊界掃描完成，共 ' + r.data.rows.length + ' 筆');
  }
  function scanHtml() {
    const sc = S.scan, sm = sc.summary;
    const wins = Object.keys(sm.routeWins || {}).map((id) => { const r = S.def && laneOf(S.def, id); return (r ? r.name : id) + '：' + sm.routeWins[id]; }).join('、');
    return h`<h3>邊界掃描</h3>
      <div class="fd-scroll"><table class="fd-tbl" id="fdScanTbl"><caption>固定其他欄位、只變動金額（所有金額條件字面值與 1,000 萬／5,000 萬的 −0.01／剛好／+0.01）</caption>
      <thead><tr><th scope="col">折後未稅金額</th><th scope="col">走哪條</th><th scope="col">草稿關卡</th><th scope="col">目前生效版關卡</th><th scope="col">比較</th></tr></thead><tbody>
      ${sc.rows.map((r) => h`<tr class="${r.floorBad ? 'bad' : r.loose ? 'loose' : r.same === false ? 'diff' : ''}"><td>${r.amountText}</td><td>${r.route}</td><td>${r.stepsText}</td><td>${r.activeStepsText == null ? '—' : r.activeStepsText}</td><td>${r.floorBad ? '⛔ 違反底線' : r.same == null ? '—' : r.same ? '✓ 相同' : r.loose ? '▲ 放寬' : '△ 不同'}</td></tr>`)}
      </tbody></table></div>
      <p class="fd-mini">全取樣統計（${comma(sm.samples)} 組${sm.truncated ? '，已達上限被截斷' : ''}${sm.sampledEvery > 1 ? '；「不同／無法解析」每 ' + sm.sampledEvery + ' 組抽 1 組統計' : ''}）：各路線被選中次數——${wins}。${sm.unreachableRoutes.length ? h`⚠ 永遠不會被選中的路線：${sm.unreachableRoutes.map((x) => x.name).join('、')}。` : '✓ 沒有永遠不會生效的路線。'} ${sm.unreachableMods.length ? h`⚠ 從未套用的調整規則：${sm.unreachableMods.map((x) => x.name).join('、')}。` : ''}
        ${sm.differing != null ? h`與目前生效版不同 ${comma(sm.differing)} 組、其中放寬 ${comma(sm.loosened)} 組。` : ''} 違反底線 ${sm.floorViolations} 組；無法解析簽核人 ${sm.blockedSamples} 組${Object.keys(sm.blockCodes || {}).length ? '（' + Object.keys(sm.blockCodes).join('、') + '）' : ''}。</p>
      <div class="fd-row"><button type="button" class="fd-btn sm" id="fdScanCsv">下載 CSV</button></div>`;
  }
  function csvCell(v) { let s = String(v == null ? '' : v); if (/^[=+\-@\t\r]/.test(s)) s = "'" + s; return '"' + s.replace(/"/g, '""') + '"'; }
  function downloadScanCsv() {
    if (!S.scan) return;
    const rows = [['折後未稅金額(元)', '走哪條', '草稿關卡', '目前生效版關卡', '比較']].concat(S.scan.rows.map((r) => [r.amount, r.route, r.stepsText, r.activeStepsText == null ? '' : r.activeStepsText, r.floorBad ? '違反底線' : r.same == null ? '' : r.same ? '相同' : r.loose ? '放寬' : '不同']));
    const csv = '﻿' + rows.map((r) => r.map(csvCell).join(',')).join('\r\n');
    const url = root.URL.createObjectURL(new root.Blob([csv], { type: 'text/csv;charset=utf-8' }));
    const a = doc.createElement('a'); a.href = url; a.download = '簽核流程邊界掃描.csv'; doc.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => root.URL.revokeObjectURL(url), 2000);
  }

  // ───────────────────────── 驗證面板 ─────────────────────────
  const STI = { pass: ['✓ 通過', 'ok'], warn: ['▲ 警告', 'warn'], error: ['✕ 錯誤', 'err'] };
  function renderVal() {
    const el = $('#fdVal'); if (!el) return;
    const c = S.checks;
    if (!c) { put(el, h`<h2 class="fd-title">即時驗證</h2><p class="fd-help">尚未載入草稿。</p>`); return; }
    const e = c.filter((x) => x.status === 'error').length, w = c.filter((x) => x.status === 'warn').length, p = c.filter((x) => x.status === 'pass').length;
    const rc = S.rc || {};
    put(el, h`<h2 class="fd-title">即時驗證（規格 11 條）</h2>
      <div class="fd-sumbar" role="status"><span class="fd-pill err">✕ 錯誤 ${e}</span><span class="fd-pill warn">▲ 警告 ${w}</span><span class="fd-pill ok">✓ 通過 ${p}</span><span class="fd-mini fd-muted">內容類錯誤（第 1、2、3、6、10、11 條）會讓草稿存不進去；名冊類（第 4、5 條）與警告只提醒。每次編輯都由後端重驗，違反保護底線的編輯會被即時拒絕。${rc.samples ? ' 本次邊界取樣 ' + comma(rc.samples) + ' 組。' : ''}</span></div>
      <ul class="fd-chk">${c.map((k) => h`<li class="${k.status}" data-chk="${k.n}"><div class="fd-h"><span class="fd-t">${k.n}. ${k.title}</span><span class="fd-pill ${STI[k.status][1]}">${STI[k.status][0]}</span></div><div class="fd-code">${k.code}</div><div class="fd-s">${k.summary}${k.n === 10 && S.rejects ? '；本次已拒絕 ' + S.rejects + ' 次違反底線的編輯' : ''}</div>
        ${k.issues.length ? h`<ul>${k.issues.map((i) => h`<li><span>${i.level === 'warn' ? '▲ ' : '✕ '}${i.text}</span>${i.loc ? h`<button type="button" class="fd-btn sm" data-loc="${JSON.stringify(i.loc)}">點擊定位</button>` : ''}</li>`)}</ul>` : ''}</li>`)}</ul>`);
  }
  function switchTab(t) {
    S.tab = t;
    $$('#fdTabs [role=tab]').forEach((b) => { const on = b.dataset.tab === t; b.setAttribute('aria-selected', String(on)); b.tabIndex = on ? 0 : -1; });
    TABS.forEach(([k]) => { const p = $('#fdTab-' + k); if (p) p.hidden = k !== t; });
    if (t === 'ov') refreshOverview();
    if (t === 'drafts') loadDrafts();
  }
  function locate(loc) {
    if (S.tab !== 'design') switchTab('design');
    let el = null;
    if (loc.mod) el = $$('.fd-mod').find((x) => x.dataset.mod === loc.mod) || null;
    else if (loc.idx != null && loc.idx >= 0) el = findNode(loc.lane, loc.idx);
    if (!el && loc.lane) el = findLane(loc.lane);
    if (!el && loc.node) el = $$('.fd-pal[data-pal="node"]').find((x) => x.dataset.tier === loc.node) || null;
    if (!el) return;
    el.scrollIntoView({ block: 'center', behavior: root.matchMedia && root.matchMedia('(prefers-reduced-motion:reduce)').matches ? 'auto' : 'smooth' });
    el.classList.add('flash'); setTimeout(() => el.classList.remove('flash'), 1300);
    if (el.matches('.fd-node,.fd-mod')) el.focus({ preventScroll: true });
    announce('已定位');
  }

  // ───────────────────────── 總覽（唯讀自由畫布）─────────────────────────
  async function refreshOverview() {
    if (S.tab !== 'ov') return;
    const host = $('#fdOvHost'); if (!host) return;
    if (!S.def) { put($('#fdOvInfo'), '尚未載入草稿'); return; }
    const seq = ++S.ovSeq;
    const body = { def: S.def };
    const useSim = S.ovSimOn && S.simIn;
    if (useSim) body.input = S.simIn.input;
    let r;
    try { r = await api('POST', '/designer/overview', body); } catch (e) { return; }
    if (seq !== S.ovSeq) return;
    if (!r.ok) { $('#fdOvInfo').textContent = (r.data && r.data.error) || '產生總覽失敗'; return; }
    if (!S.ovCtl) S.ovCtl = root.FlowOverview.mount(host, { onView: (v) => { const z = $('#fdOvZ'); if (z) z.textContent = Math.round(v.k * 100) + '%'; } });
    const g = r.data.graph, prev = S.ovCtl.state();
    S.ovCtl.setGraph(g, prev.nodes === g.nodes.length && prev.nodes > 0);
    $('#fdOvInfo').textContent = useSim && r.data.sim ? '高亮：模擬輸入命中「' + r.data.sim.route.name + '」，共 ' + r.data.sim.steps.length + ' 關' : (S.simIn && !S.ovSimOn ? '未套用模擬高亮' : '');
    put($('#fdOutline'), h`${g.outline.map((o) => h`<li class="${o.hit ? 'hit' : ''}"><b>優先 ${o.order}：${o.name}</b>${o.hit ? ' ● 本單走這條' : ''}<br><span class="fd-muted">${o.whenText}</span><br>${o.stepsText}</li>`)}`);
  }

  // ───────────────────────── 草稿：載入／儲存／刪除 ─────────────────────────
  async function loadDrafts() {
    try {
      const r = await api('GET', '/drafts');
      if (r.ok) { S.drafts = r.data.items || []; S.draftsLoaded = true; }
    } catch (e) { /* 略 */ }
    renderDrafts();
  }
  function renderDrafts() {
    const el = $('#fdDrafts'); if (!el) return;
    const ed = canEdit();
    put(el, h`<h2 class="fd-title">草稿清單</h2>
      <p class="fd-help">草稿存在後端（獨立資料區，與任何簽核無關），最多 ${S.meta ? S.meta.limits.drafts : 50} 份。草稿只有管理員看得到；<b>草稿不會生效</b>。</p>
      <div class="fd-row"><button type="button" class="fd-btn primary fd-edit-only" id="fdDraftNew" ${ed ? '' : raw('disabled')}>＋ 從預設流程開始新草稿</button></div>
      ${S.drafts.length ? h`<div class="fd-scroll"><table class="fd-tbl" id="fdDraftTbl"><thead><tr><th scope="col">名稱</th><th scope="col">備註</th><th scope="col">路線數</th><th scope="col">最後更新</th><th scope="col">更新者</th><th scope="col">操作</th></tr></thead><tbody>
        ${S.drafts.map((d) => h`<tr class="fd-draftrow ${S.draft.id === d.id ? 'cur' : ''}" data-draft="${d.id}"><td><b>${d.name}</b>${S.draft.id === d.id ? h` <span class="fd-pill ok">編輯中</span>` : ''}</td><td>${d.note}</td><td>${d.routes}</td><td>${fmtTime(d.updatedAt)}</td><td>${d.updatedBy}</td>
          <td><div class="fd-row"><button type="button" class="fd-btn sm" data-draft-load="${d.id}">載入</button><button type="button" class="fd-btn sm danger fd-edit-only" data-draft-del="${d.id}" ${ed ? '' : raw('disabled')}>刪除</button></div></td></tr>`)}
        </tbody></table></div>` : h`<p class="fd-help">${S.draftsLoaded ? '還沒有任何草稿。編輯後按上方「儲存草稿」就會出現在這裡。' : '載入中…'}</p>`}`);
    const c = $('#fdCnt-drafts'); if (c) { c.hidden = !S.drafts.length; c.textContent = String(S.drafts.length); c.className = 'fd-cnt warn'; }
  }
  async function refreshValidation() {
    if (!S.def) return;
    try {
      const r = await api('POST', '/designer/validate', { def: S.def });
      if (r.ok) { S.rc = r.data; S.checks = r.data.checks; S.view = r.data.view || S.view; }
    } catch (e) { /* 略 */ }
  }
  async function confirmDiscard() {
    if (!isDirty()) return true;
    return confirmBox('捨棄未儲存的變更？', h`<p>目前編輯中的草稿有尚未儲存的變更，繼續會把它們丟掉。</p>`, '捨棄並繼續', true);
  }
  async function startFromDefault(opts) {
    opts = opts || {};
    S.busy = true;
    let r;
    try { r = await api('GET', '/designer/default-flow'); } catch (e) { S.busy = false; toast('連線失敗，無法載入預設流程'); return false; }
    S.busy = false;
    if (!r.ok) { toast((r.data && r.data.error) || '載入預設流程失敗'); return false; }
    const d = clone(r.data.def);
    d.origin = 'designer'; d.id = 'draft-new'; d.name = opts.keepName ? S.draft.name : '新草稿';
    S.def = d; S.view = r.data.view; S.checks = r.data.checks; S.rc = null;
    S.draft = { id: null, name: opts.keepName ? S.draft.name : '新草稿', note: '', updatedAt: null, baseVer: r.data.baseVer };
    S.skipUsers = r.data.skipUserCount || 0;
    $('#fdSimExcWrap').hidden = !S.skipUsers;
    S.sel = null; setBanner(null); closeDrawerQuiet();
    await refreshValidation();
    snapshotSaved();
    S.simRes = null; if (S.simIn) await runSimQuiet();
    renderAll();
    return true;
  }
  function closeDrawerQuiet() { S.drawer = null; const dr = $('#fdDrawer'); if (dr) dr.hidden = true; const g = $('#fdGrid'); if (g) g.classList.remove('drawer-open'); }
  async function loadDraft(id) {
    if (!(await confirmDiscard())) return;
    let r;
    try { r = await api('GET', '/drafts/' + encodeURIComponent(id)); } catch (e) { toast('連線失敗'); return; }
    if (!r.ok) { toast((r.data && r.data.error) || '載入草稿失敗'); loadDrafts(); return; }
    const x = r.data;
    S.def = x.def; S.view = x.view; S.rc = null;
    S.draft = { id: x.id, name: x.name, note: x.note || '', updatedAt: x.updatedAt, baseVer: x.baseVer || '' };
    S.sel = null; setBanner(null); closeDrawerQuiet();
    await refreshValidation();
    snapshotSaved();
    S.simRes = null; if (S.simIn) await runSimQuiet();
    renderAll(); switchTab('design');
    toast('已載入草稿「' + x.name + '」');
  }
  async function deleteDraft(id) {
    const d = S.drafts.find((x) => x.id === id); if (!d) return;
    if (!(await confirmBox('刪除草稿', h`<p>要永久刪除草稿「${d.name}」嗎？此動作無法復原（稽核紀錄會保留刪除事件）。</p>`, '刪除', true))) return;
    let r;
    try { r = await api('DELETE', '/drafts/' + encodeURIComponent(id)); } catch (e) { toast('連線失敗'); return; }
    if (!r.ok) { toast((r.data && r.data.error) || '刪除失敗'); return; }
    if (S.draft.id === id) { S.draft.id = null; S.draft.updatedAt = null; S.saved = ''; }
    toast('已刪除草稿「' + d.name + '」');
    await loadDrafts(); renderAll();
  }
  function reasonList(data) {
    if (!data) return h`<p>儲存失敗。</p>`;
    if (data.blocking && data.blocking.length) return h`<p>後端驗證沒有通過，草稿未儲存：</p><ul>${data.blocking.map((b) => h`<li><b>第 ${b.n} 條（${b.code}）</b>：${(b.issues || []).join('；')}</li>`)}</ul>`;
    return h`<p>${data.error || '儲存失敗'}</p>`;
  }
  async function saveDraft(asNew) {
    if (!canEdit() || !S.def) return;
    const name = (S.draft.name || '').trim();
    if (!name) { toast('請先填草稿名稱'); const n = $('#fdName'); if (n) n.focus(); return; }
    if (S.rc && S.rc.ok === false) {
      await dialog(h`<h2 id="fdDlgT">還不能儲存</h2><div><p>草稿有「內容類」驗證錯誤，後端不會接受：</p><ul>${S.rc.blocking.map((b) => h`<li><b>第 ${b.n} 條（${b.code}）</b>：${(b.issues || []).join('；')}</li>`)}</ul><p class="fd-help">請先到「驗證」分頁修正（點「點擊定位」可跳到出問題的地方）。</p></div><div class="fd-acts"><button type="button" class="fd-btn primary" data-close="ok" autofocus>知道了</button></div>`);
      switchTab('val'); return;
    }
    const body = { name, note: S.draft.note || '', def: S.def, baseVer: S.draft.baseVer || '' };
    if (!asNew && S.draft.id) { body.id = S.draft.id; body.ifUpdatedAt = S.draft.updatedAt; }
    let r;
    try { r = await api('PUT', '/drafts', body); } catch (e) { toast('連線失敗，草稿未儲存'); return; }
    if (r.status === 409 && r.data && r.data.code === 'DRAFT_STALE') {
      await dialog(h`<h2 id="fdDlgT">草稿已被更新過</h2><div><p>這份草稿在你編輯期間，被其他人（或其他視窗）更新過了（${fmtTime(r.data.updatedAt)}，${r.data.updatedBy}）。為避免蓋掉別人的修改，這次沒有儲存。</p><p class="fd-help">你可以按「另存新草稿」把目前內容存成新的一份，或到「草稿清單」重新載入最新版。</p></div><div class="fd-acts"><button type="button" class="fd-btn primary" data-close="ok" autofocus>知道了</button></div>`);
      return;
    }
    if (!r.ok) {
      await dialog(h`<h2 id="fdDlgT">儲存失敗</h2><div>${reasonList(r.data)}</div><div class="fd-acts"><button type="button" class="fd-btn primary" data-close="ok" autofocus>知道了</button></div>`);
      return;
    }
    const x = r.data;
    S.def = x.def; S.view = x.view;
    S.draft = { id: x.id, name: x.name, note: x.note || '', updatedAt: x.updatedAt, baseVer: x.baseVer || S.draft.baseVer };
    await refreshValidation();
    snapshotSaved();
    renderAll();
    toast(x.created ? '已新增草稿「' + x.name + '」（草稿不會生效）' : '已儲存草稿「' + x.name + '」');
    loadDrafts();
  }

  // ───────────────────────── 總 render ─────────────────────────
  function renderAll() {
    if (!S.inited) return;
    renderBar(); renderPalette(); renderCanvas(); applySel(); renderSimOut(); renderVal(); renderDrafts();
    $$('.fd-edit-only').forEach((e) => { if (isMobile()) e.setAttribute('aria-hidden', 'true'); else e.removeAttribute('aria-hidden'); });
    if (S.drawer) { if (!S.def) closeDrawer(); else if (S.drawer.kind === 'route') refreshSentence(); }
    if (S.tab === 'ov') refreshOverview();
  }

  // ───────────────────────── 拖拉（Pointer Events；滑鼠／觸控／手寫筆）─────────────────────────
  let DR = null, suppressClick = 0;
  function ghostLabel(p) {
    const [k, a, b] = p.split(':');
    if (k === 'pnode') return labelOfTier(S.def, a);
    if (k === 'node') { const r = laneOf(S.def, a); return r ? (nodeById(S.def, r.path[+b]) || { label: '關卡' }).label : '關卡'; }
    if (k === 'pcond') return '◆ 條件：' + VARS[a].label;
    if (k === 'lane') return '路線：' + (laneOf(S.def, a) || { name: '' }).name;
    return '連線';
  }
  function clearDropMarks() { $$('#fdLanes .drop-ok,#fdLanes .drop-bad,#fdLanes .drop-before,#fdLanes .drop-after').forEach((e) => e.classList.remove('drop-ok', 'drop-bad', 'drop-before', 'drop-after')); }
  function computeDrop(x, y) {
    const kind = DR.payload.split(':')[0];
    const els = doc.elementsFromPoint(x, y);
    const laneEl = els.map((e) => e.closest && e.closest('#fdLanes .fd-lane')).find(Boolean);
    if (!laneEl) return null;
    const lid = laneEl.dataset.lane;
    if (kind === 'lane') { const r = laneEl.getBoundingClientRect(); return { type: 'lane', lane: lid, before: y < r.top + r.height / 2, el: laneEl }; }
    if (kind === 'pcond') {
      const inHead = els.some((e) => e.closest && e.closest('.fd-lhead'));
      const dflt = laneEl.classList.contains('default');
      return { type: 'cond', lane: lid, ok: inHead && !dflt, reason: dflt ? '預設路線不能加條件' : !inHead ? '請放到最左邊的條件膠囊區' : '', el: laneEl };
    }
    const cards = $$('.fd-path > .fd-node', laneEl);
    let idx = 0, ref = null, before = true;
    if (cards.length && !els.some((e) => e.closest && e.closest('.fd-lhead'))) {
      let best = 1e9;
      cards.forEach((c, i) => { const r = c.getBoundingClientRect(), cx = r.left + r.width / 2, cy = r.top + r.height / 2, dd = Math.abs(x - cx) + Math.abs(y - cy) * 2; if (dd < best) { best = dd; ref = c; idx = i; before = x < cx; } });
      if (!before) idx++;
    }
    return { type: 'node', lane: lid, idx, ref, before, el: laneEl };
  }
  function onPointerDown(e) {
    if (e.button !== 0 || isMobile() || !canEdit()) return;
    if (e.target.closest('input,select,textarea')) return;
    if (e.target.closest('.fd-x')) return;
    const t = e.target.closest('[data-drag]'); if (!t) return;
    DR = { payload: t.dataset.drag, el: t, sx: e.clientX, sy: e.clientY, started: false, drop: null };
  }
  function onPointerMove(e) {
    if (!DR) return;
    if (!DR.started) {
      if (Math.hypot(e.clientX - DR.sx, e.clientY - DR.sy) < 6) return;
      DR.started = true; doc.body.classList.add('fd-dragging');
      const g = $('#fdGhost'); g.textContent = ghostLabel(DR.payload); g.style.display = 'block';
      if (DR.payload.startsWith('node:')) DR.el.classList.add('dragging-src');
    }
    const g = $('#fdGhost'); g.style.left = e.clientX + 'px'; g.style.top = e.clientY + 'px';
    if (e.clientY < 64) root.scrollBy(0, -20); else if (e.clientY > root.innerHeight - 64) root.scrollBy(0, 20);
    clearDropMarks();
    const dp = computeDrop(e.clientX, e.clientY); DR.drop = dp;
    if (!dp) return;
    if (dp.type === 'lane') dp.el.classList.add(dp.before ? 'drop-before' : 'drop-after');
    else if (dp.type === 'cond') dp.el.classList.add(dp.ok ? 'drop-ok' : 'drop-bad');
    else { dp.el.classList.add('drop-ok'); if (dp.ref) dp.ref.classList.add(dp.before ? 'drop-before' : 'drop-after'); }
  }
  function endDrag(e, cancel) {
    if (!DR) return;
    const D = DR; DR = null;
    doc.body.classList.remove('fd-dragging'); const g = $('#fdGhost'); if (g) g.style.display = 'none'; clearDropMarks();
    if (D.el && D.el.classList) D.el.classList.remove('dragging-src');
    if (!D.started) {
      if (!cancel && D.payload.startsWith('pnode:')) openAddDialog('node', D.payload.split(':')[1]);
      else if (!cancel && D.payload.startsWith('pcond:')) openAddDialog('cond', D.payload.split(':')[1]);
      return;
    }
    suppressClick = Date.now() + 150;
    if (cancel) return;
    const dp = D.drop;
    if (!dp) { toast('沒有放到任何路線上，已取消'); return; }
    const [k, a, b] = D.payload.split(':');
    if (k === 'lane') {
      const from = S.def.routes.findIndex((r) => r.id === a), to = S.def.routes.findIndex((r) => r.id === dp.lane);
      if (S.def.routes[from].when === true) { toast('預設路線固定在最後'); return; }
      const ti = dp.before ? to : to + 1;
      tryApply((d) => { const [x] = d.routes.splice(from, 1); let idx = from < ti ? ti - 1 : ti; idx = Math.min(idx, d.routes.length - 1); d.routes.splice(idx, 0, x); S.focusKey = 'grip:' + a; }, '路線優先順序已調整');
    } else if (k === 'pcond') {
      if (!dp.ok) { toast('條件卡放不上去：' + dp.reason); return; }
      actAddCond(dp.lane, a);
    } else if (k === 'pnode') actInsertNode(dp.lane, dp.idx, a);
    else if (k === 'node') actMoveNode(a, +b, dp.lane, dp.idx, '已移動關卡');
  }
  async function openAddDialog(kind, id) {
    const d = S.def; if (!d || !canEdit()) return;
    const lanes = d.routes.filter((r) => kind === 'node' || r.when !== true);
    const v = S.view || { routes: [] };
    const title = kind === 'node' ? `把「${labelOfTier(d, id)}」加入哪條路線（加在最後一關）？` : `把「${VARS[id].label}」條件加到哪條路線？`;
    const r = await dialog(h`<h2 id="fdDlgT">${title}</h2><div class="fd-menu">${lanes.map((l) => h`<button type="button" class="fd-btn" data-close="${l.id}">優先 ${d.routes.indexOf(l) + 1}：${l.name}${l.when === true ? '（預設）' : ''}　<span class="fd-mini fd-muted">${(v.routes.find((x) => x.id === l.id) || {}).whenText || ''}</span></button>`)}</div><div class="fd-acts"><button type="button" class="fd-btn" data-close="cancel">取消</button></div>`);
    if (!r || r === 'cancel') return;
    const lane = laneOf(S.def, r);
    if (!lane) return;
    if (kind === 'node') actInsertNode(r, lane.path.length, id); else actAddCond(r, id);
  }

  // ───────────────────────── 事件 ─────────────────────────
  async function onClick(e) {
    const T = e.target;
    if (Date.now() < suppressClick) return;
    let x;
    if ((x = T.closest('[data-tab]'))) { switchTab(x.dataset.tab); return; }
    if (T.closest('#fdBannerX')) { setBanner(null); renderCanvas(); return; }
    if (T.closest('#fdAddLane')) { actAddLane(); return; }
    if (T.closest('#fdSave')) { saveDraft(false); return; }
    if (T.closest('#fdSaveAs')) { saveDraft(true); return; }
    if (T.closest('#fdNew') || T.closest('#fdDraftNew')) { if (await confirmDiscard()) { if (await startFromDefault()) { switchTab('design'); toast('已從目前設定轉出的預設流程開始新草稿（尚未儲存）'); } } return; }
    if (T.closest('#fdRevert')) { if (await confirmBox('還原成預設流程', h`<p>要捨棄目前編輯中的內容，還原成「目前設定轉出的預設流程」嗎？（草稿名稱會保留；已儲存的版本不受影響，直到你再次按「儲存草稿」。）</p>`, '還原', true)) { await startFromDefault({ keepName: true }); S.draft.id = S.draft.id; toast('已還原成預設流程'); } return; }
    if ((x = T.closest('[data-loc]'))) { try { locate(JSON.parse(x.dataset.loc)); } catch (_) { /* 略 */ } return; }
    if (T.closest('#fdSimRun')) { runSim(); return; }
    if (T.closest('#fdSimScan')) { runScan(); return; }
    if (T.closest('#fdSimClear')) { S.simIn = null; S.simRes = null; S.scan = null; S.simErr = ''; renderAll(); return; }
    if (T.closest('#fdScanCsv')) { downloadScanCsv(); return; }
    if ((x = T.closest('[data-preset]'))) { const [a, l] = x.dataset.preset.split(','); $('#fdSimAmt').value = a; $('#fdSimPct').value = ''; const rb = $$('input[name=fdSimLvl]').find((i) => i.value === l); if (rb) rb.checked = true; runSim(); return; }
    if (T.closest('#fdOvFit')) { if (S.ovCtl) S.ovCtl.fit(); return; }
    if (T.closest('#fdOvIn')) { if (S.ovCtl) S.ovCtl.zoomBy(1.25); return; }
    if (T.closest('#fdOvOut')) { if (S.ovCtl) S.ovCtl.zoomBy(0.8); return; }
    if ((x = T.closest('[data-draft-load]'))) { loadDraft(x.dataset.draftLoad); return; }
    if ((x = T.closest('[data-draft-del]'))) { deleteDraft(x.dataset.draftDel); return; }
    // 抽屜
    if ((x = T.closest('[data-color]')) && S.drawer && S.drawer.kind === 'node') { S.drawer.edit.color = x.dataset.color; $$('#fdDrawer [data-color]').forEach((b) => b.setAttribute('aria-pressed', String(b === x))); return; }
    if ((x = T.closest('[data-act]'))) {
      const act = x.dataset.act;
      if (act === 'dr-close') { closeDrawer(); return; }
      if (act === 'dr-apply') { applyDrawer(); return; }
      if (act === 'dr-delnode') { const D = S.drawer; closeDrawer(); actDelNode(D.laneId, D.idx); return; }
      if (act === 'add-item') { const w = S.drawer.edit.when; if (w.items.length < 6) { w.items.push(defaultItem('amount')); renderDrawerKeep(`.fd-blk[data-i="${w.items.length - 1}"] [data-f="v"]`); } return; }
      if (act === 'rm-item') { const w = S.drawer.edit.when; if (w.items.length <= 1) { drawerErr(h`<span aria-hidden="true">⛔</span><div>至少要有一個條件</div>`); return; } w.items.splice(+x.dataset.i, 1); renderDrawer(); return; }
      if (act === 'del-node') { const li = x.closest('.fd-node'); actDelNode(li.dataset.lane, +li.dataset.idx); return; }
      if (act === 'lane-up') { actMoveLane(x.dataset.lane, -1); return; }
      if (act === 'lane-down') { actMoveLane(x.dataset.lane, 1); return; }
      if (act === 'lane-del') { actDelLane(x.dataset.lane); return; }
      if (act === 'lane-edit') { openRouteDrawer(x.dataset.lane); return; }
    }
    const nd = T.closest('#fdLanes .fd-node');
    if (nd) { S.sel = { lane: nd.dataset.lane, idx: +nd.dataset.idx }; applySel(); return; }
    const cc = T.closest('.fd-cond[data-cond]');
    if (cc) { $$('.fd-cond.sel').forEach((c) => c.classList.remove('sel')); cc.classList.add('sel'); return; }
    if (S.sel && !T.closest('#fdDrawer')) { S.sel = null; applySel(); }
  }
  function onDblClick(e) {
    const T = e.target; let x;
    if (T.closest('[data-act]')) return;
    if ((x = T.closest('#fdLanes .fd-node'))) { openNodeDrawer(x.dataset.lane, +x.dataset.idx); return; }
    if ((x = T.closest('.fd-cond[data-cond]'))) { if (!x.classList.contains('dflt')) openRouteDrawer(x.dataset.cond); return; }
    if ((x = T.closest('.fd-lname'))) { openRouteDrawer(x.closest('.fd-lane').dataset.lane); return; }
    if ((x = T.closest('.fd-mod[data-mod]'))) { openModDrawer(x.dataset.mod); return; }
  }
  function onInput(e) {
    const t = e.target;
    if (t.closest('#fdDrawer')) { drawerInput(e); return; }
    if (t.id === 'fdName') { S.draft.name = t.value; renderStatusOnly(); return; }
    if (t.id === 'fdNote') { S.draft.note = t.value; renderStatusOnly(); return; }
  }
  function renderStatusOnly() {
    const p = $('#fdSavedPill'); if (!p) return;
    const dirty = isDirty(), d = S.draft;
    p.className = 'fd-pill ' + (d.id ? (dirty ? 'cur' : 'ok') : 'gray');
    p.textContent = d.id ? (dirty ? '● 有未儲存的變更' : '✓ 草稿已儲存') : (S.def ? '○ 尚未儲存的新草稿' : '○ 尚未載入');
  }
  function onChange(e) {
    const t = e.target;
    if (t.closest('#fdDrawer')) { drawerInput(e); return; }
    if (t.id === 'fdOvSim') { S.ovSimOn = t.checked; refreshOverview(); return; }
    if (t.closest('#fdSimForm') || t.id === 'fdSimCmp') { if (S.simIn || t.id === 'fdSimCmp') { if (S.simIn) runSim(); } return; }
  }
  function onKeyDown(e) {
    const T = e.target;
    if (e.key === 'Escape') {
      if (doc.querySelector('dialog[open]')) return;
      if (S.drawer) { closeDrawer(); e.preventDefault(); return; }
      if (S.sel) { S.sel = null; applySel(); }
      return;
    }
    const tab = T.closest && T.closest('#fdTabs [role=tab]');
    if (tab && (e.key === 'ArrowRight' || e.key === 'ArrowLeft')) { const bs = $$('#fdTabs [role=tab]'), i = bs.indexOf(tab), n = bs[(i + (e.key === 'ArrowRight' ? 1 : bs.length - 1)) % bs.length]; switchTab(n.dataset.tab); n.focus(); return; }
    if (!S.def) return;
    const pal = T.closest && T.closest('.fd-pal[data-pal]');
    if (pal && (e.key === 'Enter' || e.key === ' ') && canEdit()) { e.preventDefault(); openAddDialog(pal.dataset.pal, pal.dataset.tier || pal.dataset.var); return; }
    const grip = T.closest && T.closest('.fd-grip');
    if (grip && (e.key === 'ArrowUp' || e.key === 'ArrowDown') && canEdit()) { e.preventDefault(); actMoveLane(grip.closest('.fd-lane').dataset.lane, e.key === 'ArrowUp' ? -1 : 1); return; }
    const cc = T.closest && T.closest('.fd-cond[data-cond]');
    if (cc && e.key === 'Enter' && !cc.classList.contains('dflt')) { e.preventDefault(); openRouteDrawer(cc.dataset.cond); return; }
    const md = T.closest && T.closest('.fd-mod[data-mod]');
    if (md && e.key === 'Enter') { e.preventDefault(); openModDrawer(md.dataset.mod); return; }
    const nd = T.closest && T.closest('#fdLanes .fd-node');
    if (nd && T === nd) {
      const l = nd.dataset.lane, i = +nd.dataset.idx;
      if (e.key === 'Enter') { e.preventDefault(); openNodeDrawer(l, i); return; }
      if (e.key === ' ') { e.preventDefault(); S.sel = { lane: l, idx: i }; applySel(); return; }
      if (!canEdit()) return;
      const ri = S.def.routes.findIndex((r) => r.id === l), r = S.def.routes[ri];
      if (e.key === 'ArrowLeft') { e.preventDefault(); if (i > 0) actMoveNode(l, i, l, i - 1, '關卡已往前移'); else toast('已經是第一關'); return; }
      if (e.key === 'ArrowRight') { e.preventDefault(); if (i < r.path.length - 1) actMoveNode(l, i, l, i + 2, '關卡已往後移'); else toast('已經是最後一關'); return; }
      if (e.key === 'ArrowUp' || e.key === 'ArrowDown') { e.preventDefault(); const nr = S.def.routes[ri + (e.key === 'ArrowUp' ? -1 : 1)]; if (!nr) { toast('沒有更上／下一條路線'); return; } actMoveNode(l, i, nr.id, Math.min(i, nr.path.length), '關卡已移到路線「' + nr.name + '」'); return; }
      if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); actDelNode(l, i); return; }
    }
  }

  // ───────────────────────── 啟動 ─────────────────────────
  function wire() {
    const app = $('#fdApp');
    app.addEventListener('click', onClick);
    app.addEventListener('dblclick', onDblClick);
    app.addEventListener('input', onInput);
    app.addEventListener('change', onChange);
    app.addEventListener('keydown', onKeyDown);
    app.addEventListener('pointerdown', onPointerDown);
    doc.addEventListener('pointermove', onPointerMove);
    doc.addEventListener('pointerup', (e) => endDrag(e, false));
    doc.addEventListener('pointercancel', (e) => endDrag(e, true));
    if (root.matchMedia) {
      const mq = root.matchMedia('(max-width:640px)');
      const onMq = () => { if (S.drawer) closeDrawer(); renderAll(); };
      if (mq.addEventListener) mq.addEventListener('change', onMq); else if (mq.addListener) mq.addListener(onMq);
    }
    root.addEventListener('beforeunload', (e) => { if (S.inited && isDirty()) { e.preventDefault(); e.returnValue = ''; } });
  }
  async function init() {
    if (!S.enabled) return;
    const host = $('#fdHost'); if (!host) return;
    if (S.inited) { if (S.tab === 'drafts') loadDrafts(); return; }
    S.inited = true;
    buildScaffold(host);
    wire();
    switchTab('design');
    renderAll();
    S.busy = true; renderCanvas();
    await Promise.all([loadDrafts(), startFromDefault()]);
    S.busy = false;
    renderAll();
  }
  /** 後台載入後呼叫：開關開啟（端點存在）才顯示側欄項；關閉時端點是 404，什麼都不顯示 */
  async function probe() {
    if (S.enabled) return true;
    if (S.probed) return false;   // 已探測過且回報未開啟／無權限：這個頁面生命週期內不再重複打（loadUsers 會多次呼叫）
    try {
      const r = await root.fetch(BASE + '/designer/status', { credentials: 'same-origin', headers: { Accept: 'application/json' } });
      if (r.status === 404) S.probed = true;   // 404＝開關關閉（確定結果）；401／403 可能只是還沒登入，登入後 loadUsers 會再探測一次
      if (!r.ok) return false;
      const j = await r.json();
      if (!j || j.enabled !== true) return false;
      S.enabled = true; S.meta = j;
      const it = $('#fdNavItem');
      if (it && it.classList.contains('fd-nav-hidden')) {
        it.classList.remove('fd-nav-hidden');
        const cnt = it.closest('.nav-domain') && it.closest('.nav-domain').querySelector('.nav-domain-count');
        if (cnt && /^\d+$/.test(cnt.textContent.trim())) cnt.textContent = String(+cnt.textContent.trim() + 1);
      }
      return true;
    } catch (e) { return false; }
  }

  root.FlowDesigner = {
    probe, init, state: () => S, version: 'm2',
    _test: { esc, h, raw, wanToYuan, yuanToWan, fmtYuan, itemText, astToEdit, editToAst, defaultItem, itemError, canonStr, csvCell, tryApply, renderAll, switchTab, floorMsg },
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = root.FlowDesigner;
})(typeof window !== 'undefined' ? window : globalThis);
