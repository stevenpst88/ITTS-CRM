'use strict';
/**
 * 「目前這個呼叫鏈屬於哪個請求」的極小封裝（AsyncLocalStorage）。
 * 只有 DB_WRITE_SERIALIZE=1 時 lib/writeGuard.js 才會 run() 進來；db/postgres.js 的 save() 透過 writeGuard 註冊的 hook 查詢，
 * 用來判斷「這次 db.save 是不是在本請求持有的寫入鎖內」以及「本請求自己的存檔有沒有失敗」，而不是看全行程的計數器。
 * 注意：multer 之類由 socket 事件驅動的回呼會丟失 context（store 為 undefined），呼叫端要能處理 undefined。
 */
const { AsyncLocalStorage } = require('async_hooks');
const als = new AsyncLocalStorage();
module.exports = {
  run(store, fn) { return als.run(store, fn); },
  current() { return als.getStore(); },
};
