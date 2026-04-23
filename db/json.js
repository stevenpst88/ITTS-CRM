// ════════════════════════════════════════════════════════════
//  JSON 檔案實作（原本的 db.js 搬來，僅供本地開發使用）
// ════════════════════════════════════════════════════════════
const fs   = require('fs');
const path = require('path');

const DB_FILE = path.join(__dirname, '..', 'data.json');

function load() {
  if (!fs.existsSync(DB_FILE)) {
    fs.writeFileSync(DB_FILE, JSON.stringify({ contacts: [] }, null, 2), 'utf8');
  }
  return JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
}

function save(data) {
  fs.writeFileSync(DB_FILE, JSON.stringify(data, null, 2), 'utf8');
}

module.exports = { load, save };
