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
    // server.js 的 loadAuth() 在 postgres 模式讀 data._auth.users
    if (!data._auth) data._auth = { users: [] };
    if (!Array.isArray(data._auth.users)) data._auth.users = [];

    if (data._auth.users.some(u => u.username === username)) {
      console.error(`❌ 帳號 ${username} 已存在`);
      process.exit(2);
    }

    const passwordHash = await bcrypt.hash(password, 12);
    data._auth.users.push({
      username,
      password: passwordHash,          // server.js 讀 user.password
      displayName: displayName || username,
      role: 'admin',
      active: true,
      canDownloadContacts: true,
      canSetTargets: true,
      createdAt: new Date().toISOString(),
    });

    await db.save(data);
    console.log(`✅ 已建立管理員帳號：${username}（role: admin）`);
    process.exit(0);
  } catch (err) {
    console.error('❌ 失敗:', err);
    process.exit(3);
  }
})();
