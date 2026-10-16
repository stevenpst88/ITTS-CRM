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

const { createAuditStore, jsonlAdapter, legacyReadAdapter, createAuditFacade } = require('../lib/auditStore');
const AuditLegacy = require('../lib/auditLegacy');
const AuditChain = require('../lib/auditChain');   // 防竄改鏈（AUDIT_CHAIN=1 才生效；預設關閉）

const ROOT = path.join(__dirname, '..');
const DB_FILE = path.join(ROOT, 'data.json');

// 稽核事件：append-only JSON Lines（無筆數上限）；舊的 audit.log.json（500 筆）與 contact_audit.json 保留為唯讀歷史
// 安全開關 AUDIT_EVENTS_TABLE：沒設成 '1'（預設）＝舊行為（audit.log.json 500 筆、contact_audit.json）；'1'＝append-only audit.log.jsonl。
const readJsonArray = (f) => { try { const a = JSON.parse(fs.readFileSync(f, 'utf8')); return Array.isArray(a) ? a : []; } catch { return []; } };
const AUDIT_TABLE_ON = process.env.AUDIT_EVENTS_TABLE === '1';
const AUDIT_CHAIN_CFG = AUDIT_TABLE_ON ? AuditChain.getConfig(process.env) : null;
const auditStore = createAuditFacade({
  enabled: AUDIT_TABLE_ON,
  chain: AUDIT_CHAIN_CFG,
  tableStore: createAuditStore(
    jsonlAdapter({
      chain: AUDIT_CHAIN_CFG,
      file: path.join(ROOT, 'audit.log.jsonl'),
      legacyLogFile: path.join(ROOT, 'audit.log.json'),
      legacyContactFile: path.join(ROOT, 'contact_audit.json'),
    }),
    { fallbackFile: path.join(ROOT, 'audit.failed.jsonl') }
  ),
  legacyStore: createAuditStore(legacyReadAdapter({ readLogs: async () => readJsonArray(path.join(ROOT, 'audit.log.json')), legacyContactFile: path.join(ROOT, 'contact_audit.json') })),
  legacyWrite: (ev) => AuditLegacy.writeJsonMode(ev, ROOT),
});
function insertAuditEvent(ev) { return auditStore.insert(ev); }

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
// 改 async 與 postgres 介面一致（呼叫端 await db.loadAuditLog()）
async function loadAuditLog() {
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

module.exports = { load, save, flush: () => auditStore.flush(), loadAuditLog, appendAuditLog, auditStore, insertAuditEvent };
