'use strict';
/**
 * 測試環境標示（APP_ENV_LABEL，例如 DEV）。
 *
 * 目的：測試站與正式站跑同一份程式碼；環境變數 APP_ENV_LABEL 有值時，登入頁／主畫面／管理員後台顯示醒目的標示，
 *      並在瀏覽器分頁標題前加 "[DEV] "，避免測試人員誤以為在正式站。
 * 未設定（或清理後為空）時：伺服器的任何回應都與沒有此功能時逐位元相同——不註冊任何路由、不改寫任何 HTML。
 *
 * 設計：沒有「公開 API 端點」。標示字串由伺服器在送出 HTML 時直接寫進一個 <script data-env="…"> 標籤（值已清理成只剩
 *      英數、-、_、空白，屬性內不可能跳出），所以登入前就能顯示、不需要額外的未登入端點，也就沒有可被探測或洩漏資訊的面。
 *      外部腳本 /env-label.js 只在標示有值時才註冊成公開路由（未設定時此路徑維持原狀＝需登入的靜態檔區之外不存在）。
 * 用戶端腳本以字串常數內建於此（不放進 _client/）：_client/ 內的檔案在登入後一律可被取得，放在那裡會讓未設定的環境多出一個可存取的檔案。
 */

const MAX_LEN = 12;

// 只允許英數字、-、_、空白；其餘（含 < > " ' & 控制字元、全形、中文、emoji）一律剝除。
// 長度上限 12；至少要有一個英數字，否則視為未設定（避免只剩 "---" 之類無意義標示）。
function getEnvLabel(env) {
  const raw = env && env.APP_ENV_LABEL;
  if (typeof raw !== 'string') return '';
  let s = raw.replace(/[^A-Za-z0-9_\- ]/g, '').replace(/ +/g, ' ').trim();
  s = s.slice(0, MAX_LEN).trim();
  return /[A-Za-z0-9]/.test(s) ? s : '';
}

// 把腳本標籤塞進 HTML（</body> 之前；找不到就附在最後）。label 為空 → 原樣回傳（同一個字串）。
// label 必須是 getEnvLabel 的輸出；這裡再防禦性檢查一次，不合格式就不注入。
function injectEnvLabel(html, label) {
  if (typeof html !== 'string' || !label || !/^[A-Za-z0-9_\- ]{1,12}$/.test(label)) return html;
  const tag = '<script src="/env-label.js" data-env="' + label + '"></script>\n';
  const i = html.lastIndexOf('</body>');
  return i < 0 ? html + tag : html.slice(0, i) + tag + html.slice(i);
}

// 用戶端腳本（ES5、純 DOM API、textContent；樣式走 CSSOM，不產生 style 屬性字串）。
const CLIENT_JS = [
  '(function () {',
  '  try {',
  '    var me = document.currentScript;',
  '    var label = me && me.getAttribute("data-env");',
  '    if (!label || !/^[A-Za-z0-9_\\- ]{1,12}$/.test(label)) return;',
  '    var prefix = "[" + label + "] ";',
  '    function fixTitle() { if (document.title.indexOf(prefix) !== 0) document.title = prefix + document.title; }',
  '    function mount() {',
  '      if (document.getElementById("itts-env-label")) return;',
  '      var box = document.createElement("div");',
  '      box.id = "itts-env-label";',
  '      box.setAttribute("role", "note");',
  '      box.setAttribute("aria-label", "目前為 " + label + " 測試環境");',
  '      function css(el, o) { for (var k in o) el.style.setProperty(k, o[k]); }',
  '      css(box, { position: "fixed", top: "0", left: "0", right: "0", height: "0", "z-index": "2147483647", "pointer-events": "none" });',
  '      var bar = document.createElement("div");',
  '      css(bar, { position: "absolute", top: "0", left: "0", right: "0", height: "3px", background: "#f59e0b" });',
  '      var pill = document.createElement("div");',
  '      pill.textContent = label;',
  '      css(pill, { position: "absolute", top: "0", left: "50%", transform: "translateX(-50%)", padding: "2px 14px 3px", background: "#f59e0b", color: "#1f1300", font: "700 12px/1.3 system-ui,-apple-system,Segoe UI,sans-serif", "letter-spacing": ".08em", "white-space": "nowrap", "border-radius": "0 0 8px 8px", "box-shadow": "0 1px 4px rgba(0,0,0,.35)", "max-width": "60vw", overflow: "hidden", "text-overflow": "ellipsis" });',
  '      box.appendChild(bar); box.appendChild(pill);',
  '      document.body.appendChild(box);',
  '    }',
  '    fixTitle();',
  '    var t = document.querySelector("title");',
  '    if (t && window.MutationObserver) new MutationObserver(fixTitle).observe(t, { childList: true, characterData: true, subtree: true });',
  '    if (document.body) mount(); else document.addEventListener("DOMContentLoaded", mount);',
  '  } catch (e) { /* 標示失敗不影響頁面 */ }',
  '})();',
  '',
].join('\n');

module.exports = { getEnvLabel, injectEnvLabel, CLIENT_JS, MAX_LEN };
