// ── 簽核設定「兩人覆核」前端元件 (config-review.js) ───────────────
// 後端：lib/configReview.js（GET/POST /api/admin/approval-config-requests…）。開關是伺服器環境變數 APPROVAL_CONFIG_DUAL_REVIEW，畫面不能切換。
// 同一份給兩個頁面用：admin.html（簽核流程設定）與 quote-approval.js 的「報價單簽核設定」對話框。
//   const cr = CfgReview.create({ api, onApplied });      api(method, path, body) → Promise<{ ok, status, data }>，path 例如 '/admin/approval-config-requests'
//   cr.paint(hostEl)                                      畫「兩人覆核：啟用／未啟用」晶片＋（啟用時）「待覆核變更 N 筆」橫幅；點「覆核」開視窗
//   cr.refresh()                                          重新抓列表並重畫所有已登記的 host
//   cr.enabled() / cr.saveLabel('儲存設定')               啟用時儲存鈕文字變成「提出變更（待覆核）」
//   CfgReview.proposalMessage(resp.data)                  202 之後要顯示的訊息（含單號）
// 所有進 innerHTML 的字串一律 esc()。純函式（chipHtml／bannerHtml／requestsHtml／historyHtml…）掛在 CfgReview 上，scripts/check-config-dual-review.js 以 vm 測。
(function (root) {
  'use strict';

  var STATUS_LABEL = { pending: '待覆核', approved: '已核可', rejected: '已退回', cancelled: '已撤回', expired: '已逾期', stale: '已作廢（基準變動）' };
  var STATUS_COLOR = { pending: '#b45309', approved: '#15803d', rejected: '#b91c1c', cancelled: '#6b7280', expired: '#6b7280', stale: '#6b7280' };
  var ACTION_LABEL = { propose: '提出', approve: '核可', reject: '退回', cancel: '撤回', expire: '逾期', stale: '作廢' };

  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function fmtTime(iso) {
    var t = Date.parse(iso);
    if (!isFinite(t)) return '';
    var d = new Date(t), p = function (n) { return (n < 10 ? '0' : '') + n; };
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes());
  }

  /** st = { enabled, activeAdmins } */
  function chipHtml(st) {
    var on = !!(st && st.enabled);
    if (!on) return '<span class="cr-chip gc-cr-off" data-cr-chip="off" style="display:inline-block;padding:2px 10px;border-radius:999px;font-size:12px">兩人覆核：未啟用</span>';
    if ((st.activeAdmins | 0) < 2) {
      return '<span class="cr-chip gc-cr-blocked" data-cr-chip="blocked" style="display:inline-block;padding:2px 10px;border-radius:999px;font-size:12px">兩人覆核：啟用，但在職管理員不足 2 人，目前無法儲存變更</span>';
    }
    return '<span class="cr-chip gc-cr-on" data-cr-chip="on" style="display:inline-block;padding:2px 10px;border-radius:999px;font-size:12px">兩人覆核：啟用</span>';
  }
  function saveLabel(enabled, normal) { return enabled ? '提出變更（待覆核）' : normal; }
  function proposalMessage(d) {
    d = d || {};
    if (d.noChange) return '設定沒有變更，不需覆核';
    return (d.message || '已提出，等待另一位管理員覆核') + (d.requestId ? '（單號 ' + String(d.requestId).slice(0, 8) + '）' : '');
  }
  function bannerHtml(st) {
    if (!st || !st.enabled) return '';
    var n = st.pendingCount | 0;
    if (!n) return '<div style="margin-top:6px"><button type="button" data-cr="open" style="background:none;border:none;color:#1a73e8;font-size:12.5px;cursor:pointer;padding:0">覆核紀錄</button></div>';
    return '<div class="cr-banner" style="margin-top:8px;padding:8px 12px;border-radius:8px;background:#fffbeb;border:1px solid #fde68a;color:#92400e;font-size:13px;display:flex;align-items:center;gap:10px">' +
      '<span>待覆核變更 ' + n + ' 筆</span><button type="button" data-cr="open" style="background:#1a73e8;color:#fff;border:none;padding:4px 14px;border-radius:6px;font-size:13px;cursor:pointer">覆核</button></div>';
  }
  function diffListHtml(lines) {
    var arr = Array.isArray(lines) ? lines : [];
    if (!arr.length) return '<div style="font-size:12.5px;color:#6b7280">（沒有差異摘要）</div>';
    return '<ul style="margin:6px 0 0 18px;padding:0;font-size:13px;line-height:1.6">' + arr.map(function (l) { return '<li>' + esc(l) + '</li>'; }).join('') + '</ul>';
  }
  function historyLinesHtml(h) {
    return (Array.isArray(h) ? h : []).map(function (x) {
      return '<div style="font-size:12px;color:#6b7280">' + esc(fmtTime(x.at)) + ' ' + esc(x.by) + ' ' + esc(ACTION_LABEL[x.action] || x.action) + (x.comment ? '：' + esc(x.comment) : '') + '</div>';
    }).join('');
  }
  /** 待覆核卡片。isMine＝提案人本人：核可／退回停用並說明，只能撤回 */
  function requestCardHtml(r) {
    var id = esc(r.id);
    var mine = !!r.isMine;
    var h = '<div class="cr-card" data-cr-id="' + id + '" style="border:1px solid #e5e7eb;border-radius:10px;padding:12px 14px;margin-bottom:12px;background:#fff">' +
      '<div style="display:flex;justify-content:space-between;gap:8px;flex-wrap:wrap"><b>' + esc(r.kindLabel || r.kind) + '</b><span style="font-size:12px;color:#6b7280">' + esc(r.proposedBy) + ' 於 ' + esc(fmtTime(r.proposedAt)) + ' 提出；' + esc(fmtTime(r.expiresAt)) + ' 前有效</span></div>' +
      diffListHtml(r.diffSummary);
    if (r.baseChanged) h += '<div style="margin-top:6px;font-size:12.5px;color:#b91c1c">提醒：提案之後設定已被其他變更改動，核可時會被拒絕（需提案人重新提出）。</div>';
    var hist = (r.history || []).filter(function (x) { return x.action === 'propose' && x.comment; });
    if (hist.length) h += '<div style="margin-top:6px;font-size:12.5px;color:#374151">提案說明：' + esc(hist[0].comment) + '</div>';
    if (mine) {
      h += '<div style="margin-top:8px;font-size:12.5px;color:#6b7280">這是你自己提出的變更，須由另一位管理員覆核（你不能核可或退回自己的申請）。</div>' +
        '<div style="margin-top:8px"><button type="button" data-cr="cancel" data-id="' + id + '" style="background:#fff;border:1px solid #d1d5db;padding:5px 14px;border-radius:6px;cursor:pointer">撤回</button>' +
        ' <button type="button" disabled style="padding:5px 14px;border-radius:6px;border:1px solid #e5e7eb;background:#f3f4f6;color:#9ca3af">核可</button>' +
        ' <button type="button" disabled style="padding:5px 14px;border-radius:6px;border:1px solid #e5e7eb;background:#f3f4f6;color:#9ca3af">退回</button></div>';
    } else {
      h += '<div style="margin-top:8px"><input type="text" data-cr-comment="' + id + '" maxlength="200" placeholder="意見（退回必填，至少 2 字）" style="width:100%;box-sizing:border-box;padding:6px 8px;border:1px solid #d1d5db;border-radius:6px;font-size:13px"></div>' +
        '<div style="margin-top:8px"><button type="button" data-cr="approve" data-id="' + id + '" style="background:#15803d;color:#fff;border:none;padding:5px 14px;border-radius:6px;cursor:pointer">核可</button>' +
        ' <button type="button" data-cr="reject" data-id="' + id + '" style="background:#fff;color:#b91c1c;border:1px solid #fecaca;padding:5px 14px;border-radius:6px;cursor:pointer">退回</button></div>';
    }
    return h + '<div data-cr-msg="' + id + '" role="status" style="margin-top:6px;font-size:12.5px;color:#b91c1c"></div></div>';
  }
  function requestsHtml(list) {
    var p = (list || []).filter(function (r) { return r.status === 'pending'; });
    if (!p.length) return '<div style="color:#6b7280;font-size:13px;padding:8px 0">目前沒有待覆核的變更。</div>';
    return p.map(requestCardHtml).join('');
  }
  function historyHtml(list) {
    var done = (list || []).filter(function (r) { return r.status !== 'pending'; });
    if (!done.length) return '<div style="color:#6b7280;font-size:13px">尚無歷史紀錄。</div>';
    return '<div style="display:flex;flex-direction:column;gap:6px">' + done.map(function (r) {
      return '<div class="cr-hist" style="border:1px solid #eee;border-radius:8px;padding:8px 10px;font-size:13px"><div><span style="font-weight:600;color:' + (STATUS_COLOR[r.status] || '#374151') + '">' + esc(STATUS_LABEL[r.status] || r.status) + '</span> · ' +
        esc(r.kindLabel || r.kind) + ' · 提案 ' + esc(r.proposedBy) + ' ' + esc(fmtTime(r.proposedAt)) + (r.decidedBy ? ' · 處理 ' + esc(r.decidedBy) + ' ' + esc(fmtTime(r.decidedAt)) : '') + '</div>' +
        diffListHtml(r.diffSummary) + historyLinesHtml(r.history) + '</div>';
    }).join('') + '</div>';
  }
  function modalInnerHtml(list, titleId) {
    return '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px"><b class="cr-title"' + (titleId ? ' id="' + esc(titleId) + '"' : '') + ' style="font-size:16px">簽核設定變更覆核</b><button type="button" class="cr-close" data-cr="close" aria-label="關閉" style="border:none;background:none;font-size:20px;cursor:pointer">&times;</button></div>' +
      '<div style="font-size:12.5px;color:#6b7280;margin-bottom:10px">核可後才會真正生效；提案人不能覆核自己的變更。申請 3 天內無人處理會自動失效。</div>' +
      '<div data-cr="pending">' + requestsHtml(list) + '</div>' +
      '<details style="margin-top:8px"><summary style="cursor:pointer;font-size:13px;color:#374151">歷史申請</summary><div style="margin-top:8px">' + historyHtml(list) + '</div></details>';
  }

  var instanceSeq = 0;
  function create(opts) {
    var api = opts.api;
    var st = { enabled: false, activeAdmins: 0, pendingCount: 0, requests: [], me: '', loaded: false };
    var host = null;
    var overlay = null;
    var prevFocus = null;   // 開啟視窗前的焦點（關閉時還原）
    var uid = 'crTitle' + (++instanceSeq);

    function paint(h) {
      if (h && h !== host) {
        host = h;
        host.addEventListener('click', function (ev) {
          var t = ev.target && ev.target.closest ? ev.target.closest('[data-cr="open"]') : null;
          if (t) openModal();
        });
      }
      if (host) host.innerHTML = chipHtml(st) + bannerHtml(st);
    }
    async function refresh() {
      var r;
      try { r = await api('GET', '/admin/approval-config-requests'); } catch (e) { r = null; }
      if (r && r.ok && r.data) {
        st.enabled = !!r.data.enabled; st.activeAdmins = r.data.activeAdmins | 0; st.pendingCount = r.data.pendingCount | 0;
        st.requests = Array.isArray(r.data.requests) ? r.data.requests : []; st.me = r.data.me || ''; st.loaded = true;
      }
      paint();
      if (overlay) renderModal();
      return st;
    }
    function focusables() {
      if (!overlay) return [];
      return Array.prototype.slice.call(overlay.querySelectorAll('button, input, summary, [href], [tabindex]')).filter(function (e) {
        return !e.disabled && e.getAttribute('tabindex') !== '-1' && e.offsetParent !== null;
      });
    }
    function renderModal() {
      if (!overlay) return;
      // 重畫前記住：各卡片已輸入的「意見」、目前焦點在哪個控制項——重畫後還原（失敗後 refresh() 不能清掉使用者剛打的字）
      var saved = {};
      Array.prototype.forEach.call(overlay.querySelectorAll('[data-cr-comment]'), function (i) { if (i.value) saved[i.getAttribute('data-cr-comment')] = i.value; });
      var a = document.activeElement, fd = null;
      if (a && overlay.contains(a)) {
        if (a.hasAttribute('data-cr-comment')) fd = '[data-cr-comment="' + a.getAttribute('data-cr-comment') + '"]';
        else if (a.getAttribute('data-cr') && a.getAttribute('data-id')) fd = '[data-cr="' + a.getAttribute('data-cr') + '"][data-id="' + a.getAttribute('data-id') + '"]';
        else if (a.getAttribute('data-cr') === 'close') fd = '[data-cr="close"]';
      }
      overlay.firstChild.innerHTML = modalInnerHtml(st.requests, uid);
      Object.keys(saved).forEach(function (id) {
        var inp = overlay.querySelector('[data-cr-comment="' + id + '"]');
        if (inp) inp.value = saved[id];
      });
      var re = fd && overlay.querySelector(fd);
      if (re && !re.disabled) { try { re.focus(); } catch (e) { /* 忽略 */ } }
      else if (fd) { var c0 = overlay.querySelector('[data-cr="close"]'); if (c0) c0.focus(); }   // 原本的控制項已不存在（例如該筆已處理）：回到關閉鈕，避免焦點掉到頁面後面
    }
    function onKey(ev) {
      if (!overlay) return;
      if (ev.key === 'Escape') { ev.preventDefault(); ev.stopPropagation(); closeModal(); return; }
      if (ev.key !== 'Tab') return;
      var f = focusables();
      if (!f.length) { ev.preventDefault(); return; }
      var first = f[0], last = f[f.length - 1], a = document.activeElement;
      if (!overlay.contains(a)) { ev.preventDefault(); first.focus(); }
      else if (ev.shiftKey && a === first) { ev.preventDefault(); last.focus(); }
      else if (!ev.shiftKey && a === last) { ev.preventDefault(); first.focus(); }
    }
    function openModal() {
      if (overlay) return;
      prevFocus = document.activeElement;
      overlay = document.createElement('div');
      overlay.className = 'cr-overlay';
      overlay.setAttribute('role', 'dialog');
      overlay.setAttribute('aria-modal', 'true');
      overlay.setAttribute('aria-labelledby', uid);
      overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:100000;display:flex;align-items:flex-start;justify-content:center;overflow:auto;padding:30px 12px';
      var box = document.createElement('div');
      box.className = 'cr-modal-box';
      box.style.cssText = 'background:#fff;border-radius:12px;max-width:760px;width:100%;padding:18px 20px;box-sizing:border-box';
      overlay.appendChild(box);
      overlay.addEventListener('click', onClick);
      document.addEventListener('keydown', onKey, true);
      document.body.appendChild(overlay);
      renderModal();
      var f0 = focusables()[0]; if (f0) f0.focus();   // 焦點進對話框（第一個可操作的控制項）
      refresh();
    }
    function closeModal() {
      if (!overlay) return;
      document.removeEventListener('keydown', onKey, true);
      overlay.remove(); overlay = null;
      // 還原焦點：原本的「覆核」鈕若已被重畫換掉，就找 host 裡新的那顆
      var back = prevFocus && prevFocus.isConnected ? prevFocus : (host && host.querySelector('[data-cr="open"]'));
      prevFocus = null;
      if (back && typeof back.focus === 'function') { try { back.focus(); } catch (e) { /* 忽略 */ } }
    }
    function setMsg(id, text) {
      var el = overlay && overlay.querySelector('[data-cr-msg="' + id + '"]');
      if (el) el.textContent = text || '';
    }
    async function onClick(ev) {
      var t = ev.target && ev.target.closest ? ev.target.closest('[data-cr]') : null;
      if (!t) { if (ev.target === overlay) closeModal(); return; }
      var act = t.getAttribute('data-cr');
      if (act === 'close') { closeModal(); return; }
      var id = t.getAttribute('data-id');
      if (!id || t.disabled) return;
      var inp = overlay.querySelector('[data-cr-comment="' + id + '"]');
      var comment = inp ? String(inp.value || '').trim() : '';
      if (act === 'reject' && Array.from(comment).length < 2) { setMsg(id, '退回必須填寫原因（至少 2 個字）'); return; }
      if (act === 'cancel' && typeof confirm === 'function' && !confirm('確定撤回這筆變更申請？')) return;
      t.disabled = true;
      var r = null;
      try { r = await api('POST', '/admin/approval-config-requests/' + encodeURIComponent(id) + '/' + act, { comment: comment }); }
      catch (e) { r = null; }   // 網路例外：不能讓按鈕卡在 disabled、使用者看不到任何訊息
      finally { t.disabled = false; }
      if (r && r.ok) {
        await refresh();
        if (act === 'approve' && typeof opts.onApplied === 'function') opts.onApplied();
      } else {
        var failMsg = (r && r.data && r.data.error) || (r ? '操作失敗' : '操作失敗（網路異常，請稍後再試）');
        setMsg(id, failMsg);
        await refresh();
        setMsg(id, failMsg);
      }
    }
    return { paint: paint, refresh: refresh, openModal: openModal, closeModal: closeModal, enabled: function () { return st.enabled; }, state: function () { return st; }, saveLabel: function (n) { return saveLabel(st.enabled, n); } };
  }

  root.CfgReview = {
    esc: esc, fmtTime: fmtTime, chipHtml: chipHtml, saveLabel: saveLabel, proposalMessage: proposalMessage, bannerHtml: bannerHtml,
    diffListHtml: diffListHtml, requestCardHtml: requestCardHtml, requestsHtml: requestsHtml, historyHtml: historyHtml, modalInnerHtml: modalInnerHtml, create: create,
  };
})(typeof window !== 'undefined' ? window : this);
