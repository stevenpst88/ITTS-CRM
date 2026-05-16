// ════════════════════════════════════════════════════════════
//  JSON 檔案實作（本地開發用）
//
//  本地仍把 _auditLog 存在同一個 data.json，不像 postgres 拆 row。
//  原因：本地不在乎 egress，且檔案 IO 已經很快；拆檔反而增加複雜度。
//
//  對外暴露的 loadAuditLog / appendAuditLog 介面與 postgres 一致，
//  讓 server.js 不需要分支 if (_USE_DB_FOR_META)。
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

// ── Audit log（本地：直接存在 data.json 的 _auditLog 欄位）──
function loadAuditLog() {
  const d = load();
  return Array.isArray(d._auditLog) ? d._auditLog : [];
}

function appendAuditLog(entry) {
  const d = load();
  if (!Array.isArray(d._auditLog)) d._auditLog = [];
  d._auditLog.unshift(entry);
  if (d._auditLog.length > 5000) d._auditLog.length = 5000;
  save(d);
}

module.exports = { load, save, loadAuditLog, appendAuditLog };
