'use strict';
/**
 * 測試啟動伺服器前，先確認埠沒有人在聽（B-03）。
 *
 * 為什麼：各測試腳本用「GET /login.html 回 200」判斷自己的 server 起來了。埠已被別人（別的 session／agent）的伺服器佔用時，
 * 自己的 server 會因 EADDRINUSE 立刻結束，可是就緒檢查卻打到「別人的 server」並通過，之後的登入打到別人的伺服器 →
 * 401／429 假失敗，甚至把測試資料寫進對方的環境。啟動前先檢查埠空閒，被佔用就明確報錯並拒絕執行。
 *
 * 做法：嘗試連線 127.0.0.1:port（連得上＝有人在聽）。比 listen 試探更可靠（Windows 對同埠 listen 的行為不一致）。
 * 丟出 Error（code='EPORTBUSY'）而不是直接 process.exit：讓呼叫端既有的 catch／finally 清掉自己剛寫的臨時 auth.json／data.json。
 */
const net = require('net');

function isPortInUse(port, host) {
  return new Promise((resolve) => {
    const s = net.connect({ port: Number(port), host: host || '127.0.0.1' });
    let done = false;
    const fin = (v) => { if (!done) { done = true; try { s.destroy(); } catch { /* 已關閉 */ } resolve(v); } };
    s.once('connect', () => fin(true));
    s.once('error', () => fin(false));
    s.setTimeout(1500, () => fin(false));
  });
}

async function assertPortFree(port, label) {
  if (await isPortInUse(port, '127.0.0.1') || await isPortInUse(port, '::1')) {
    const e = new Error('埠 ' + port + ' 已被其他程序佔用' + (label ? '（' + label + '）' : '') + '——可能是別的 session／agent 的伺服器。拒絕啟動，避免測試誤連到別人的伺服器（假失敗 401／429，或把測試資料寫進對方環境）。請換埠（看腳本檔頭的 *_PORT 環境變數）或先結束該程序。');
    e.code = 'EPORTBUSY';
    throw e;
  }
}

module.exports = { assertPortFree, isPortInUse };
