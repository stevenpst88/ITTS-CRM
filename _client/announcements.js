/* 公告彈窗（使用者端 + 後台預覽共用）
 *
 * 使用者端：app.js 的 initUser() 完成後呼叫 ITTSAnnounce.start(user)。
 *   - 失敗一律靜默（絕不可阻擋登入／載入）。
 *   - 一次性（once）：伺服器已排除「已按我知道了」的；這裡逐則彈出，按「我知道了」→ POST ack。
 *   - 每次登入（everyLogin）：用「登入戳記 loginAt」判斷，而不是 sessionStorage 單獨判斷。
 *       loginAt 由伺服器在 /api/login 當下寫入 session（JWT），F5／還原 session 不變，重新登入必定更新。
 *       若只用 sessionStorage：同一分頁「session 逾時→重新登入」不會清掉旗標而漏彈；開新分頁又會重彈。
 *       本檔以 localStorage 記 {loginAt, 已顯示的 id:version}，loginAt 一變就整份作廢 → 新登入必彈、F5 不彈、新分頁不重彈。
 *   - 所有文字一律走 textContent（不經 innerHTML），公告內容 XSS-inert。
 * 後台：ITTSAnnounce.preview({title, body, severity, mode}) 只顯示、不送 ack。
 */
(function () {
  'use strict';
  if (window.ITTSAnnounce) return;

  // 強制閱讀秒數：每一則彈窗（once／everyLogin、含後台預覽）顯示後，「我知道了」須等這麼多秒才可按。
  // 伺服器端另有防繞過檢查（lib/announcements.js READ_MIN_MS，略小於此值以容忍網路與計時誤差）。
  var READ_SECONDS = 5;

  var CSS = [
    '.ann-overlay{position:fixed;inset:0;z-index:100000;display:flex;align-items:center;justify-content:center;padding:16px;',
    '  background:rgba(0,0,0,.5);--ann-bg:#fff;--ann-fg:#1f2937;--ann-sub:#6b7280;--ann-line:#e5e7eb;--ann-accent:#1a73e8;--ann-accent-fg:#fff;--ann-warn:#d93025;--ann-warn-bg:#fdecea;}',
    'body.dark .ann-overlay{--ann-bg:#161b22;--ann-fg:#e6edf3;--ann-sub:#8b949e;--ann-line:#30363d;--ann-accent:#1f6feb;--ann-warn:#ff7b72;--ann-warn-bg:#3b1a1a;}',
    '.ann-card{background:var(--ann-bg);color:var(--ann-fg);width:520px;max-width:100%;max-height:86vh;display:flex;flex-direction:column;',
    '  border-radius:14px;box-shadow:0 20px 60px rgba(0,0,0,.3);border-top:5px solid var(--ann-accent);overflow:hidden;animation:annIn .18s ease;font-family:inherit}',
    '.ann-card.ann-important{border-top-color:var(--ann-warn)}',
    '@keyframes annIn{from{transform:translateY(14px);opacity:0}to{transform:none;opacity:1}}',
    '.ann-head{padding:18px 22px 10px;display:flex;align-items:center;gap:10px}',
    '.ann-badge{flex:none;font-size:12px;font-weight:700;padding:2px 9px;border-radius:99px;background:var(--ann-line);color:var(--ann-sub)}',
    '.ann-important .ann-badge{background:var(--ann-warn-bg);color:var(--ann-warn)}',
    '.ann-title{margin:0;font-size:18px;font-weight:700;line-height:1.4;word-break:break-word;flex:1}',
    '.ann-body{padding:4px 22px 16px;overflow-y:auto;font-size:14px;line-height:1.75;white-space:pre-wrap;word-break:break-word}',
    '.ann-foot{padding:12px 22px 16px;border-top:1px solid var(--ann-line);display:flex;align-items:center;justify-content:space-between;gap:10px}',
    '.ann-count{font-size:12px;color:var(--ann-sub)}',
    '.ann-btn{border:none;border-radius:8px;padding:9px 22px;font-size:14px;font-weight:600;cursor:pointer;background:var(--ann-accent);color:var(--ann-accent-fg);font-family:inherit}',
    '.ann-important .ann-btn{background:var(--ann-warn);color:#fff}',
    '.ann-btn:disabled{opacity:.6;cursor:default}',
    '.ann-btn:focus-visible{outline:3px solid rgba(26,115,232,.45);outline-offset:2px}',
    '@media (max-width:600px){.ann-overlay{align-items:flex-end;padding:0}.ann-card{width:100%;border-radius:16px 16px 0 0;max-height:92vh}}'
  ].join('\n');

  function ensureStyle() {
    if (document.getElementById('annStyle')) return;
    var st = document.createElement('style');
    st.id = 'annStyle';
    st.textContent = CSS;
    (document.head || document.documentElement).appendChild(st);
  }

  // 建立單則彈窗；全部文字用 textContent。回傳 overlay 元素。
  function buildCard(a, idx, total, okLabel, onOk) {
    ensureStyle();
    var ov = document.createElement('div');
    ov.className = 'ann-overlay';
    ov.setAttribute('data-ann-id', String(a.id || 'preview'));
    var card = document.createElement('div');
    card.className = 'ann-card' + (a.severity === 'important' ? ' ann-important' : '');
    card.setAttribute('role', 'dialog');
    card.setAttribute('aria-modal', 'true');
    var head = document.createElement('div'); head.className = 'ann-head';
    var badge = document.createElement('span'); badge.className = 'ann-badge';
    badge.textContent = a.severity === 'important' ? '重要公告' : '公告';
    var h = document.createElement('h3'); h.className = 'ann-title';
    h.id = 'annTitle' + Math.random().toString(36).slice(2, 8);
    h.textContent = String(a.title == null ? '' : a.title);
    card.setAttribute('aria-labelledby', h.id);
    head.appendChild(badge); head.appendChild(h);
    var body = document.createElement('div'); body.className = 'ann-body';
    body.textContent = String(a.body == null ? '' : a.body);
    var foot = document.createElement('div'); foot.className = 'ann-foot';
    var cnt = document.createElement('span'); cnt.className = 'ann-count';
    cnt.textContent = total > 1 ? ('第 ' + (idx + 1) + ' / ' + total + ' 則') : '';
    var btn = document.createElement('button'); btn.type = 'button'; btn.className = 'ann-btn';
    // 倒數：沒有關閉鈕、不處理 Esc／點背景，倒數結束前按鈕 disabled；點擊處理器再擋一次（防 DOM 竄改 disabled）
    var left = READ_SECONDS;
    btn.disabled = true;
    btn.textContent = '請先閱讀（' + left + '）';
    var timer = setInterval(function () {
      left--;
      if (left <= 0) {
        clearInterval(timer);
        btn.disabled = false;
        btn.textContent = okLabel;
        try { btn.focus(); } catch (e) { /* ignore */ }
      } else {
        btn.textContent = '請先閱讀（' + left + '）';
      }
    }, 1000);
    btn.addEventListener('click', function () {
      if (left > 0) return;
      clearInterval(timer);
      btn.disabled = true; onOk(ov);
    });
    foot.appendChild(cnt); foot.appendChild(btn);
    card.appendChild(head); card.appendChild(body); card.appendChild(foot);
    ov.appendChild(card);
    document.body.appendChild(ov);
    try { btn.focus(); } catch (e) { /* ignore */ }
    return ov;
  }

  // ── 本機「這次登入已顯示」紀錄 ──
  function shownKey(username) { return 'annShown:' + username; }
  function loadShown(username, loginAt) {
    try {
      var o = JSON.parse(localStorage.getItem(shownKey(username)) || 'null');
      if (o && o.loginAt === loginAt && o.ids) return o;
    } catch (e) { /* ignore */ }
    return { loginAt: loginAt, ids: {} };
  }
  var _mem = {}; // localStorage 不可用時的記憶體備援（至少同一頁面生命週期內不重彈）
  function saveShown(username, o) {
    _mem[username] = o;
    try { localStorage.setItem(shownKey(username), JSON.stringify(o)); } catch (e) { /* ignore */ }
  }

  var running = false;

  async function start(user) {
    if (running) return;
    running = true;
    try {
      var username = (user && user.username) || window._myUsername || '';
      if (!username) return;
      var res = await fetch('/api/announcements/active', { credentials: 'same-origin', cache: 'no-store' });
      if (!res.ok) return;
      var d = await res.json();
      var loginAt = d && d.loginAt ? d.loginAt : 'legacy';
      var shown = (_mem[username] && _mem[username].loginAt === loginAt) ? _mem[username] : loadShown(username, loginAt);
      var queue = ((d && d.items) || []).filter(function (a) {
        return a && a.id && !shown.ids[a.id + ':' + a.version]; // once 也要擋：回條失敗時，同一次登入內不重彈
      });
      for (var i = 0; i < queue.length; i++) {
        await showOne(queue[i], i, queue.length, username, shown);
      }
    } catch (e) {
      /* 公告失敗不影響登入與使用 */
    } finally {
      running = false;
    }
  }

  function showOne(a, idx, total, username, shown) {
    return new Promise(function (resolve) {
      buildCard(a, idx, total, '我知道了', function (ov) {
        shown.ids[a.id + ':' + a.version] = 1; // 本次登入內不再彈（once/everyLogin 皆然）
        saveShown(username, shown);
        // 回條失敗也放行：once 公告最壞情況是下次登入再彈一次
        fetch('/api/announcements/' + encodeURIComponent(a.id) + '/ack', {
          method: 'POST', credentials: 'same-origin',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ version: a.version })
        }).catch(function () { /* ignore */ }).then(function () {
          if (ov.parentNode) ov.parentNode.removeChild(ov);
          resolve();
        });
      });
    });
  }

  // 後台預覽：不送 ack
  function preview(a) {
    return new Promise(function (resolve) {
      buildCard(a || {}, 0, 1, '關閉預覽', function (ov) {
        if (ov.parentNode) ov.parentNode.removeChild(ov);
        resolve();
      });
    });
  }

  window.ITTSAnnounce = { start: start, preview: preview, READ_SECONDS: READ_SECONDS };
})();
