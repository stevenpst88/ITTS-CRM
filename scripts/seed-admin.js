// ════════════════════════════════════════════════════════════
//  建立第一個管理員帳號
//
//  用法：
//    node scripts/seed-admin.js <username> <password> [displayName]
//
//  會連到 DB_BACKEND 指定的後端（預設 json）。
//  若 DB_BACKEND=postgres，確保 .env 有 DATABASE_URL。
// ════════════════════════════════════════════════════════════
require('dotenv').config();
const bcrypt = require('bcryptjs');
const db     = require('../db');

const [,, username, password, displayName] = process.argv;
if (!username || !password) {
  console.error('用法: node scripts/seed-admin.js <username> <password> [displayName]');
  process.exit(1);
}

(async () => {
  try {
    if (typeof db.ready === 'function') await db.ready();

    const data = db.load();
    // auth.json 結構：儲存在 DB 的 app_data.content.users（如果有）
    // 如果本來業務邏輯是分開 loadAuth()，我們也塞一份到 data.users 作為種子
    if (!Array.isArray(data.users)) data.users = [];

    if (data.users.some(u => u.username === username)) {
      console.error(`❌ 帳號 ${username} 已存在`);
      process.exit(2);
    }

    const passwordHash = await bcrypt.hash(password, 12);
    data.users.push({
      username,
      passwordHash,
      displayName: displayName || username,
      role: 'manager1',
      createdAt: new Date().toISOString(),
    });

    await db.save(data);
    console.log(`✅ 已建立帳號：${username}（角色 manager1）`);
    console.log('⚠️  請確認 server.js 的 loadAuth() 能正確讀到此帳號，若仍用 auth.json 則需額外遷移。');
    process.exit(0);
  } catch (err) {
    console.error('❌ 失敗:', err);
    process.exit(3);
  }
})();
