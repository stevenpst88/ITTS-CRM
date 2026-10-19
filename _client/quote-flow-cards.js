/*!
 * 報價單「簽核進度卡片」(quote-flow-cards.js) —— 流程設計器 M0：唯讀視覺化，不碰簽核引擎。
 *
 *   瀏覽器：<script src="quote-flow-cards.js"> → window.QFlowCards
 *   Node  ：require('../_client/quote-flow-cards.js')（scripts/check-quote-approval-cards.js 單元測試用）
 *
 * 輸入是伺服器 serializeApproval() 已經回傳的 q.approval（steps／cur／state／valid／board／history）。
 * 這個檔案「只讀」那些欄位：
 *   · 不使用 derived（含毛利率文字 marginText／reasons／thresholds／金額），所以對任何檢視者都不會多洩漏毛利或成本。
 *   · 不新增任何要存回報價單的欄位，也不改簽核／代核／名冊／秘書的語意；代核卡片只是把現有的 tier==='board' 與 approval.board 標示出來。
 *
 * 對外：
 *   cardsOf(approval, opts) → view-model（純函式；opts.fmtTime 可注入時間格式化）
 *   render(vm, opts)        → HTML 字串（所有動態字串都 escape）
 *   bind(container)         → 在容器上綁一次事件委派（候選人名單展開／收合）
 *   CSS                     → 樣式字串（前綴 qfc-；顏色全部是不透明色，與宿主背景／六種主題無關，對比 ≥ 4.5:1）
 *   enabled(search)         → 是否顯示卡片（預設是；網址帶 ?cards=0 退回舊的條列式）
 */
(function (root) {
  'use strict';

  var TITLE = { mgr2: '二級主管', mgr1: '一級主管', gm: '總經理', chairman: '董事長', board: '董事會代核' };
  var NODE = { mgr2: '二級主管核決', mgr1: '一級主管核決', gm: '總經理核決', chairman: '董事長核決', board: '董事會核決' };
  var LABEL_FALLBACK = { mgr2: '二級主管複核', mgr1: '一級主管', gm: '總經理', chairman: '董事長', board: '董事會決議（秘書代核）' };
  var STATUS_TEXT = { done: '已完成', cur: '目前關卡', wait: '後續關卡', ret: '已退回', skip: '未進行', void: '已失效' };
  var STATUS_ICON = { done: '✓', cur: '●', wait: '○', ret: '✕', skip: '–', void: '–' };
  var COMMENT_MAX = 40;
  var HEADLINE_NAMES_MAX = 3;

  function str(v) { return v === undefined || v === null ? '' : String(v); }
  function esc(v) {
    return str(v).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
  }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function defaultFmt(v) {
    if (!v) return '';
    var d = new Date(v);
    if (isNaN(d.getTime())) return str(v);
    return pad2(d.getMonth() + 1) + '/' + pad2(d.getDate()) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
  }
  function clip(s, n) {
    s = str(s).replace(/\s+/g, ' ').trim();
    var a = Array.from(s);
    return a.length > n ? a.slice(0, n).join('') + '…' : s;
  }
  function namesOf(s) {
    return str(s).split(/[、,，]/).map(function (x) { return x.trim(); }).filter(Boolean);
  }
  function lastAction(ap) {
    var h = ap && Array.isArray(ap.history) ? ap.history : [];
    for (var i = h.length - 1; i >= 0; i--) {
      var a = h[i] && h[i].action;
      if (a === 'INVALIDATE' || a === 'WITHDRAW' || a === 'SUBMIT' || a === 'APPROVE' || a === 'RETURN') return a;
    }
    return '';
  }

  /** approval → view-model。null／沒有 approval（顧問視角）→ cards 為空、state 'none'。 */
  function cardsOf(ap, opts) {
    var fmt = (opts && opts.fmtTime) || defaultFmt;
    var vm = { state: 'none', headline: '', tag: null, banner: null, cards: [] };
    if (!ap || typeof ap !== 'object') return vm;
    var steps = Array.isArray(ap.steps) ? ap.steps : [];
    var apState = ap.state;

    if (!steps.length) {
      var la = lastAction(ap);
      if (apState === 'none' || !apState) {
        if (la === 'INVALIDATE') { vm.state = 'invalidated'; vm.banner = { kind: 'void', text: '核准已作廢，需重新送簽。' }; }
        else if (la === 'WITHDRAW') { vm.state = 'withdrawn'; vm.banner = { kind: 'wait', text: '已撤回，尚未重新送簽。' }; }
        else { vm.state = 'draft'; }
      } else { vm.state = apState; }
      return vm;
    }

    var voided = apState === 'approved' && ap.valid === false;
    var cur = Math.max(0, Math.min(steps.length - 1, parseInt(ap.cur, 10) || 0));
    var boardInfo = ap.board && typeof ap.board === 'object' ? ap.board : null;

    var cards = steps.map(function (s, i) {
      var tier = str(s.tier);
      var raw = str(s.status);
      var status;
      if (voided) status = 'void';
      else if (apState === 'returned') status = raw === 'approved' ? 'done' : (raw === 'returned' || raw === 'pending') ? 'ret' : 'skip';
      else if (apState === 'approved') status = raw === 'approved' ? 'done' : 'wait';
      else if (apState === 'pending') status = raw === 'approved' ? 'done' : (i === cur ? 'cur' : (raw === 'returned' ? 'ret' : 'wait'));
      else status = raw === 'approved' ? 'done' : raw === 'returned' ? 'ret' : 'wait';

      var people = namesOf(s.assigneeName);
      var multi = (tier === 'gm' || tier === 'chairman') && people.length >= 2;
      var proxy = tier === 'board';
      var finished = raw === 'approved' || raw === 'returned';
      var name = str(s.label) || LABEL_FALLBACK[tier] || tier;
      var card = {
        n: i + 1,
        total: steps.length,
        tier: tier,
        name: name,
        nodeLabel: NODE[tier] || (name + '核決'),
        title: TITLE[tier] || '',
        status: status,
        statusText: STATUS_TEXT[status],
        icon: STATUS_ICON[status],
        current: status === 'cur',
        people: people,
        who: multi ? (people.length + ' 位候選人') : (people.join('、') || ''),
        candidates: multi ? { count: people.length, names: people.slice() } : null,
        proxy: proxy,
        proxyBy: '',
        decidedBy: finished ? str(s.byName) : '',
        at: finished && s.at ? fmt(s.at) : '',
        comment: finished && s.comment ? clip(s.comment, COMMENT_MAX) : '',
        board: null,
        link: 'wait',
      };
      if (proxy) {
        card.proxyBy = str((boardInfo && boardInfo.byName) || (finished ? s.byName : ''));
        if (status === 'done' && boardInfo && (boardInfo.resolutionDate || boardInfo.resolutionNo)) {
          card.board = { resolutionDate: str(boardInfo.resolutionDate), resolutionNo: str(boardInfo.resolutionNo) };
        }
      }
      return card;
    });
    cards.forEach(function (c, i) { c.link = c.status === 'done' ? 'done' : c.status === 'ret' ? 'ret' : 'wait'; });

    vm.cards = cards;
    vm.state = voided ? 'void' : apState;

    if (voided) {
      vm.headline = '核准已失效，需重新簽核';
      vm.tag = { kind: 'void', text: '核准已失效' };
      vm.banner = { kind: 'void', text: '核准後內容被修改，下列為失效前的簽核軌跡。' };
    } else if (apState === 'pending') {
      var c = cards[cur];
      var who = c.people.length ? c.people : [];
      var head;
      if (c.candidates) head = c.name + '（' + c.candidates.count + ' 位候選人）';
      else if (who.length && who.length <= HEADLINE_NAMES_MAX) head = who.join('、');
      else head = c.name;
      vm.headline = '目前待 ' + head + ' 審核';
      vm.tag = { kind: 'cur', text: '目前待 ' + (c.candidates ? (c.candidates.count + ' 位候選人') : (c.who || c.name)) + '｜' + (c.title || c.name) };
    } else if (apState === 'approved') {
      vm.headline = '簽核已全部完成';
      vm.tag = { kind: 'done', text: '已核准' };
    } else if (apState === 'returned') {
      var rc = null;
      for (var k = 0; k < cards.length; k++) if (cards[k].status === 'ret') { rc = cards[k]; break; }
      vm.headline = rc ? ((rc.decidedBy ? rc.decidedBy + ' ' : '') + '已於「' + rc.name + '」退回') : '簽核已被退回';
      vm.tag = { kind: 'ret', text: '已退回' + (rc ? '｜' + rc.name : '') };
    } else {
      vm.headline = '簽核進度';
    }
    return vm;
  }

  var uid = 0;

  function cardHtml(c, idp) {
    var h = '<li class="qfc-step st-' + c.status + ' link-' + c.link + '"' + (c.current ? ' aria-current="step"' : '') + '>';
    h += '<div class="qfc-card">';
    h += '<span class="qfc-sr">第 ' + c.n + ' 關，共 ' + c.total + ' 關，' + esc(c.statusText) + '。</span>';
    h += '<div class="qfc-top"><span class="qfc-no" aria-hidden="true">STEP ' + c.n + '</span>';
    h += '<span class="qfc-badge"><span aria-hidden="true">' + esc(c.icon) + '</span> ' + esc(c.statusText) + '</span></div>';
    h += '<div class="qfc-name">' + esc(c.name) + '</div>';
    if (c.proxy) h += '<div class="qfc-proxyrow"><span class="qfc-badge qfc-proxy">代核</span>' + (c.proxyBy ? '<span class="qfc-small">代核人：' + esc(c.proxyBy) + '</span>' : '') + '</div>';
    h += '<div class="qfc-who">' + (c.who ? esc(c.who) : '<span class="qfc-small">（尚無簽核人）</span>') + (c.title && !c.proxy ? ' <span class="qfc-small">' + esc(c.title) + '</span>' : '') + '</div>';
    h += '<div class="qfc-node">流程節點：' + esc(c.nodeLabel) + '</div>';
    if (c.candidates) {
      var id = idp + '-c' + c.n;
      h += '<button type="button" class="qfc-cand-btn" aria-expanded="false" aria-controls="' + id + '" data-qfc-toggle="' + id + '">' + esc(c.candidates.count + ' 位候選人') + ' <span class="qfc-caret" aria-hidden="true">▾</span></button>';
      h += '<ul class="qfc-cand-list" id="' + id + '" hidden>' + c.candidates.names.map(function (n) { return '<li>' + esc(n) + '</li>'; }).join('') + '</ul>';
    }
    if (c.status === 'done' || c.status === 'ret') {
      var note = '';
      if (c.decidedBy || c.at) note += '<div>' + (c.status === 'ret' ? '退回：' : '簽核：') + esc(c.decidedBy) + (c.at ? '　' + esc(c.at) : '') + '</div>';
      if (c.board) note += '<div>決議：' + esc(c.board.resolutionDate) + (c.board.resolutionNo ? '　文號 ' + esc(c.board.resolutionNo) : '') + '</div>';
      if (c.comment) note += '<div class="qfc-cm">意見：' + esc(c.comment) + '</div>';
      if (note) h += '<div class="qfc-note">' + note + '</div>';
    }
    return h + '</div></li>';
  }

  function render(vm, opts) {
    if (!vm) return '';
    var idp = 'qfc' + (++uid);
    var h = '';
    if (!vm.cards.length && !vm.banner) return '';
    h += '<div class="qfc" data-state="' + esc(vm.state) + '">';
    if (vm.cards.length) {
      var k = vm.tag ? vm.tag.kind : 'wait';
      h += '<div class="qfc-now k-' + esc(k) + '"><div class="qfc-line" role="status">' + esc(vm.headline) + '</div>';
      if (vm.tag) h += '<span class="qfc-tag">' + esc(vm.tag.text) + '</span>';
      h += '</div>';
    }
    if (vm.banner) h += '<div class="qfc-banner k-' + esc(vm.banner.kind) + '" role="note">' + esc(vm.banner.text) + '</div>';
    if (vm.cards.length) {
      var n = Math.min(vm.cards.length, 5);
      h += '<ol class="qfc-flow n' + n + '" aria-label="簽核關卡，共 ' + vm.cards.length + ' 關">';
      h += vm.cards.map(function (c) { return cardHtml(c, idp); }).join('');
      h += '</ol>';
    }
    return h + '</div>';
  }

  var BOUND = typeof WeakSet === 'function' ? new WeakSet() : null;
  function bind(container) {
    if (!container || !container.addEventListener) return;
    if (BOUND) { if (BOUND.has(container)) return; BOUND.add(container); }
    container.addEventListener('click', function (ev) {
      var t = ev.target;
      var btn = t && t.closest ? t.closest('[data-qfc-toggle]') : null;
      if (!btn || !container.contains(btn)) return;
      var list = container.querySelector('#' + btn.getAttribute('data-qfc-toggle'));
      if (!list) return;
      var open = btn.getAttribute('aria-expanded') !== 'true';
      btn.setAttribute('aria-expanded', open ? 'true' : 'false');
      if (open) list.removeAttribute('hidden'); else list.setAttribute('hidden', '');
    });
  }

  function enabled(search) {
    try {
      var s = search === undefined ? (root.location && root.location.search) : search;
      if (!s) return true;
      return !/(^|[?&])cards=0(&|$)/.test(String(s));
    } catch (e) { return true; }
  }

  // 顏色 token：每個狀態一組 --bg／--fg／--fg2／--bd，全部不透明；testing 腳本會逐組算對比（≥ 4.5:1）
  var CSS = [
    '.qfc { --gap: 26px; container-type: inline-size; color: var(--fg, #14202e); }',
    '.qfc { --st-done-bg:#e3f4e8; --st-done-fg:#0f4a22; --st-done-fg2:#2d5a3b; --st-done-bd:#1e7a3a;',
    '  --st-cur-bg:#fff1dc; --st-cur-fg:#5c2c00; --st-cur-fg2:#7a4310; --st-cur-bd:#d96b00;',
    '  --st-wait-bg:#eef1f5; --st-wait-fg:#2b3644; --st-wait-fg2:#4a5565; --st-wait-bd:#6f7c8f;',
    '  --st-ret-bg:#fde7e6; --st-ret-fg:#7d1410; --st-ret-fg2:#8f2b27; --st-ret-bd:#c0342c;',
    '  --st-void-bg:#e9ebee; --st-void-fg:#3a4350; --st-void-fg2:#555f6d; --st-void-bd:#6a7686;',
    '  --st-proxy-bg:#ece5fb; --st-proxy-fg:#3f2486; --st-proxy-bd:#6f4fc7; }',
    'body.dark .qfc { --st-done-bg:#10321f; --st-done-fg:#d4f7e0; --st-done-fg2:#a9d9bb; --st-done-bd:#34d399;',
    '  --st-cur-bg:#3d2610; --st-cur-fg:#ffe9d1; --st-cur-fg2:#f2c99c; --st-cur-bd:#fb923c;',
    '  --st-wait-bg:#1e2636; --st-wait-fg:#dce4f2; --st-wait-fg2:#b4c0d6; --st-wait-bd:#6b7a93;',
    '  --st-ret-bg:#451a1a; --st-ret-fg:#ffe0dd; --st-ret-fg2:#f5b4ae; --st-ret-bd:#f87171;',
    '  --st-void-bg:#242a35; --st-void-fg:#d6dce8; --st-void-fg2:#aab4c6; --st-void-bd:#6b7686;',
    '  --st-proxy-bg:#2d2350; --st-proxy-fg:#e6dcff; --st-proxy-bd:#a78bfa; }',
    '.qfc-now { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 8px 12px; padding: 10px 14px; border-radius: 12px; border: 1px solid var(--bd); background: var(--bg); color: var(--fg); margin-bottom: 14px; }',
    '.qfc-now.k-cur { --bg: var(--st-cur-bg); --fg: var(--st-cur-fg); --bd: var(--st-cur-bd); }',
    '.qfc-now.k-done { --bg: var(--st-done-bg); --fg: var(--st-done-fg); --bd: var(--st-done-bd); }',
    '.qfc-now.k-ret { --bg: var(--st-ret-bg); --fg: var(--st-ret-fg); --bd: var(--st-ret-bd); }',
    '.qfc-now.k-void, .qfc-now.k-wait { --bg: var(--st-void-bg); --fg: var(--st-void-fg); --bd: var(--st-void-bd); }',
    '.qfc-line { font-size: 15px; font-weight: 700; line-height: 1.4; min-width: 0; overflow-wrap: anywhere; }',
    '.qfc-tag { display: inline-block; padding: 3px 12px; border-radius: 999px; font-size: 12.5px; font-weight: 700; border: 1px solid var(--bd); color: var(--fg); background: var(--bg); box-shadow: inset 0 0 0 1px var(--bd); max-width: 100%; overflow-wrap: anywhere; }',
    '.qfc-banner { padding: 8px 12px; border-radius: 10px; font-size: 13px; margin-bottom: 14px; border: 1px solid var(--bd); background: var(--bg); color: var(--fg); }',
    '.qfc-banner.k-void { --bg: var(--st-void-bg); --fg: var(--st-void-fg); --bd: var(--st-void-bd); }',
    '.qfc-banner.k-wait { --bg: var(--st-wait-bg); --fg: var(--st-wait-fg); --bd: var(--st-wait-bd); }',
    '.qfc-flow { list-style: none; margin: 0; padding: 0; display: grid; grid-template-columns: minmax(0, 1fr); gap: var(--gap); }',
    '.qfc-step { position: relative; min-width: 0; }',
    '.qfc-card { height: 100%; display: flex; flex-direction: column; gap: 5px; padding: 12px 13px; border-radius: 14px; border: 1px solid var(--bd); background: var(--bg); color: var(--fg); box-sizing: border-box; }',
    '.st-done { --bg: var(--st-done-bg); --fg: var(--st-done-fg); --fg2: var(--st-done-fg2); --bd: var(--st-done-bd); }',
    '.st-cur { --bg: var(--st-cur-bg); --fg: var(--st-cur-fg); --fg2: var(--st-cur-fg2); --bd: var(--st-cur-bd); }',
    '.st-wait { --bg: var(--st-wait-bg); --fg: var(--st-wait-fg); --fg2: var(--st-wait-fg2); --bd: var(--st-wait-bd); }',
    '.st-ret { --bg: var(--st-ret-bg); --fg: var(--st-ret-fg); --fg2: var(--st-ret-fg2); --bd: var(--st-ret-bd); }',
    '.st-skip, .st-void { --bg: var(--st-void-bg); --fg: var(--st-void-fg); --fg2: var(--st-void-fg2); --bd: var(--st-void-bd); }',
    '.st-cur .qfc-card { border-width: 2px; padding: 11px 12px; box-shadow: 0 0 0 3px color-mix(in srgb, var(--st-cur-bd) 30%, transparent); }',
    '.st-skip .qfc-card, .st-void .qfc-card { border-style: dashed; }',
    '.st-skip .qfc-name, .st-void .qfc-name { text-decoration: line-through; text-decoration-thickness: 1px; }',
    '.qfc-top { display: flex; align-items: center; justify-content: space-between; gap: 8px; flex-wrap: wrap; }',
    '.qfc-no { font-size: 11.5px; letter-spacing: .08em; font-weight: 700; color: var(--fg2); }',
    '.qfc-badge { display: inline-flex; align-items: center; gap: 4px; padding: 1px 9px; border-radius: 999px; font-size: 12px; font-weight: 700; white-space: nowrap; border: 1px solid var(--bd); color: var(--fg); }',
    '.qfc-badge.qfc-proxy { background: var(--st-proxy-bg); color: var(--st-proxy-fg); border-color: var(--st-proxy-bd); }',
    '.qfc-name { font-size: 16px; font-weight: 700; line-height: 1.3; overflow-wrap: anywhere; }',
    '.qfc-who { font-size: 14px; overflow-wrap: anywhere; }',
    '.qfc-small, .qfc-node { font-size: 12.5px; color: var(--fg2); }',
    '.qfc-proxyrow { display: flex; flex-wrap: wrap; gap: 4px 8px; align-items: center; }',
    '.qfc-note { margin-top: auto; padding-top: 7px; border-top: 1px dashed var(--bd); font-size: 12.5px; color: var(--fg2); overflow-wrap: anywhere; }',
    '.qfc-cm { margin-top: 2px; }',
    '.qfc-cand-btn { appearance: none; font: inherit; font-size: 12.5px; text-align: left; cursor: pointer; min-height: 36px; padding: 4px 10px; border-radius: 8px; border: 1px solid var(--bd); background: transparent; color: var(--fg); }',
    '.qfc-cand-btn:hover { background: color-mix(in srgb, var(--bd) 14%, transparent); }',
    '.qfc-cand-btn[aria-expanded="true"] .qfc-caret { display: inline-block; transform: rotate(180deg); }',
    '.qfc-cand-list { list-style: none; margin: 0; padding: 7px 10px; border-radius: 9px; border: 1px solid var(--bd); font-size: 13px; display: grid; gap: 3px; color: var(--fg); }',
    '.qfc-cand-list[hidden] { display: none; }',
    '.qfc-cand-btn:focus-visible { outline: 3px solid #f29900; outline-offset: 2px; }',
    '.qfc-sr { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); white-space: nowrap; }',
    /* 連線：預設直排（垂直線＋箭頭）；已完成實線綠、退回紅、其餘虛線灰 */
    '.qfc-step:not(:last-child)::after { content: ""; position: absolute; left: 26px; top: calc(100% + 3px); height: calc(var(--gap) - 12px); border-left: 3px dashed var(--st-wait-bd); }',
    '.qfc-step:not(:last-child)::before { content: ""; position: absolute; left: 20px; top: calc(100% + var(--gap) - 11px); border: 6px solid transparent; border-top: 8px solid var(--st-wait-bd); border-bottom: 0; }',
    '.qfc-step.link-done:not(:last-child)::after { border-left-style: solid; border-left-color: var(--st-done-bd); }',
    '.qfc-step.link-done:not(:last-child)::before { border-top-color: var(--st-done-bd); }',
    '.qfc-step.link-ret:not(:last-child)::after { border-left-color: var(--st-ret-bd); }',
    '.qfc-step.link-ret:not(:last-child)::before { border-top-color: var(--st-ret-bd); }',
    /* 容器夠寬才橫排（依關數）；容器查詢不支援的瀏覽器維持直排 */
    '@container (min-width: 440px) { .qfc-flow.n2 { grid-template-columns: repeat(2, minmax(0, 1fr)); } }',
    '@container (min-width: 600px) { .qfc-flow.n3 { grid-template-columns: repeat(3, minmax(0, 1fr)); } }',
    '@container (min-width: 740px) { .qfc-flow.n4 { grid-template-columns: repeat(4, minmax(0, 1fr)); } }',
    '@container (min-width: 800px) { .qfc-flow.n5 { grid-template-columns: repeat(5, minmax(0, 1fr)); } }',
    '@container (min-width: 440px) { .qfc-flow.n2 .qfc-step:not(:last-child)::after { left: calc(100% + 3px); top: 28px; width: calc(var(--gap) - 6px); height: 0; border-left: 0; border-top: 3px dashed var(--st-wait-bd); }',
    '  .qfc-flow.n2 .qfc-step:not(:last-child)::before { left: calc(100% + var(--gap) - 10px); top: 22px; border: 6px solid transparent; border-left: 8px solid var(--st-wait-bd); border-right: 0; } }',
    '@container (min-width: 600px) { .qfc-flow.n3 .qfc-step:not(:last-child)::after { left: calc(100% + 3px); top: 28px; width: calc(var(--gap) - 6px); height: 0; border-left: 0; border-top: 3px dashed var(--st-wait-bd); }',
    '  .qfc-flow.n3 .qfc-step:not(:last-child)::before { left: calc(100% + var(--gap) - 10px); top: 22px; border: 6px solid transparent; border-left: 8px solid var(--st-wait-bd); border-right: 0; } }',
    '@container (min-width: 740px) { .qfc-flow.n4 .qfc-step:not(:last-child)::after { left: calc(100% + 3px); top: 28px; width: calc(var(--gap) - 6px); height: 0; border-left: 0; border-top: 3px dashed var(--st-wait-bd); }',
    '  .qfc-flow.n4 .qfc-step:not(:last-child)::before { left: calc(100% + var(--gap) - 10px); top: 22px; border: 6px solid transparent; border-left: 8px solid var(--st-wait-bd); border-right: 0; } }',
    '@container (min-width: 800px) { .qfc-flow.n5 .qfc-step:not(:last-child)::after { left: calc(100% + 3px); top: 28px; width: calc(var(--gap) - 6px); height: 0; border-left: 0; border-top: 3px dashed var(--st-wait-bd); }',
    '  .qfc-flow.n5 .qfc-step:not(:last-child)::before { left: calc(100% + var(--gap) - 10px); top: 22px; border: 6px solid transparent; border-left: 8px solid var(--st-wait-bd); border-right: 0; } }',
    /* 橫排時已完成／退回連線的顏色（橫排用 border-top） */
    '@container (min-width: 440px) { .qfc-flow.n2 .link-done:not(:last-child)::after { border-top-style: solid; border-top-color: var(--st-done-bd); } .qfc-flow.n2 .link-done:not(:last-child)::before { border-left-color: var(--st-done-bd); border-top-color: transparent; } .qfc-flow.n2 .link-ret:not(:last-child)::after { border-top-color: var(--st-ret-bd); } .qfc-flow.n2 .link-ret:not(:last-child)::before { border-left-color: var(--st-ret-bd); border-top-color: transparent; } }',
    '@container (min-width: 600px) { .qfc-flow.n3 .link-done:not(:last-child)::after { border-top-style: solid; border-top-color: var(--st-done-bd); } .qfc-flow.n3 .link-done:not(:last-child)::before { border-left-color: var(--st-done-bd); border-top-color: transparent; } .qfc-flow.n3 .link-ret:not(:last-child)::after { border-top-color: var(--st-ret-bd); } .qfc-flow.n3 .link-ret:not(:last-child)::before { border-left-color: var(--st-ret-bd); border-top-color: transparent; } }',
    '@container (min-width: 740px) { .qfc-flow.n4 .link-done:not(:last-child)::after { border-top-style: solid; border-top-color: var(--st-done-bd); } .qfc-flow.n4 .link-done:not(:last-child)::before { border-left-color: var(--st-done-bd); border-top-color: transparent; } .qfc-flow.n4 .link-ret:not(:last-child)::after { border-top-color: var(--st-ret-bd); } .qfc-flow.n4 .link-ret:not(:last-child)::before { border-left-color: var(--st-ret-bd); border-top-color: transparent; } }',
    '@container (min-width: 800px) { .qfc-flow.n5 .link-done:not(:last-child)::after { border-top-style: solid; border-top-color: var(--st-done-bd); } .qfc-flow.n5 .link-done:not(:last-child)::before { border-left-color: var(--st-done-bd); border-top-color: transparent; } .qfc-flow.n5 .link-ret:not(:last-child)::after { border-top-color: var(--st-ret-bd); } .qfc-flow.n5 .link-ret:not(:last-child)::before { border-left-color: var(--st-ret-bd); border-top-color: transparent; } }',
    '@media (prefers-reduced-motion: no-preference) { .st-cur .qfc-card { animation: qfcpulse 2.4s ease-in-out infinite; } @keyframes qfcpulse { 0%, 100% { box-shadow: 0 0 0 3px color-mix(in srgb, var(--st-cur-bd) 18%, transparent); } 50% { box-shadow: 0 0 0 5px color-mix(in srgb, var(--st-cur-bd) 38%, transparent); } } }',
    '@media (prefers-reduced-motion: reduce) { .qfc *, .qfc *::before, .qfc *::after { animation: none !important; transition: none !important; } }',
    '@media (forced-colors: active) { .qfc-card, .qfc-now, .qfc-banner { border: 2px solid CanvasText; } .st-cur .qfc-card { border-color: Highlight; } }',
  ].join('\n');

  var api = { cardsOf: cardsOf, render: render, bind: bind, enabled: enabled, CSS: CSS, TITLE: TITLE, NODE: NODE, STATUS_TEXT: STATUS_TEXT };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.QFlowCards = api;
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
