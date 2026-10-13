'use strict';
/**
 * 稽核「舊路徑」寫入（第二波之前的行為，位元級保留）。AUDIT_EVENTS_TABLE 沒設成 '1'（預設）時，所有稽核事件都走這裡：
 *  - 本機 JSON：audit.log.json（新的在最前面、最多 500 筆、JSON.stringify(…, null, 2)）
 *  - Postgres：整包 JSONB（app_data.id='audit_log'，最多 5000 筆）——由 db/postgres.js 的 appendAuditLog 處理，這裡只組出和以前相同欄位順序的項目
 *  - 名片欄位稽核：contact_audit.json（最多 3000 筆，寫入失敗靜默）
 *  - 報價單原始編輯事件（kind 'quote-edit'，第二波新增）沒有舊表可去 → 以一般日誌動作 QUOTE_EDIT 寫入舊路徑
 */
const fs = require('fs');
const path = require('path');
const { v4: uuidv4 } = require('uuid');

const ROOT = path.join(__dirname, '..');

/** 事件 → { type:'log'|'contact', ... }（quote-edit 轉成一般日誌） */
function classify(ev) {
  if (ev.kind === 'contact') return { type: 'contact', ev };
  if (ev.kind === 'quote-edit') {
    return { type: 'log', ev: { action: 'QUOTE_EDIT', operator: ev.operator, target: ev.target, detail: `${ev.action}｜${ev.detail == null ? '' : ev.detail}`, ip: ev.ip, ua: ev.ua } };
  }
  return { type: 'log', ev };
}

/** 本機 JSON 模式舊項目欄位順序：id, timestamp, action, operator, target, detail, ip, userAgent */
function jsonLogEntry(ev) {
  return { id: uuidv4(), timestamp: new Date().toISOString(), action: ev.action, operator: ev.operator, target: ev.target, detail: ev.detail, ip: ev.ip || '', userAgent: ev.ua || '' };
}
/** Postgres 模式舊項目欄位順序：id, action, operator, target, detail, ip, userAgent, timestamp */
function pgLogEntry(ev) {
  return { id: uuidv4(), action: ev.action, operator: ev.operator, target: ev.target, detail: ev.detail, ip: ev.ip || '', userAgent: ev.ua || '', timestamp: new Date().toISOString() };
}
function appendJsonLog(root, ev) {
  const logFile = path.join(root || ROOT, 'audit.log.json');
  let logs = [];
  try { logs = JSON.parse(fs.readFileSync(logFile, 'utf8')); } catch {}
  logs.unshift(jsonLogEntry(ev));
  if (logs.length > 500) logs = logs.slice(0, 500);   // keep latest 500
  fs.writeFileSync(logFile, JSON.stringify(logs, null, 2), 'utf8');
}
function appendContactFile(root, ev) {
  const file = path.join(root || ROOT, 'contact_audit.json');
  const x = ev.extra || {};
  let logs = [];
  try { logs = JSON.parse(fs.readFileSync(file, 'utf8')); } catch {}
  logs.unshift({
    id: uuidv4(),
    timestamp: new Date().toISOString(),
    userId: ev.operator,
    userName: x.userName,
    action: ev.action,
    targetId: x.targetId,
    targetName: x.targetName || '',
    targetCompany: x.targetCompany || '',
    changes: x.changes || [],
  });
  if (logs.length > 3000) logs = logs.slice(0, 3000);
  try { fs.writeFileSync(file, JSON.stringify(logs, null, 2), 'utf8'); } catch {}
}

/** 本機 JSON 模式的舊路徑寫入（同步） */
function writeJsonMode(ev, root) {
  const c = classify(ev);
  if (c.type === 'contact') appendContactFile(root, c.ev); else appendJsonLog(root, c.ev);
  return true;
}
/** Postgres 模式的舊路徑寫入：一般日誌交給 appendAuditLog（整包 JSONB），名片稽核寫檔 */
function writePgMode(ev, appendAuditLog, root) {
  const c = classify(ev);
  if (c.type === 'contact') { appendContactFile(root, c.ev); return true; }
  return appendAuditLog(pgLogEntry(c.ev));
}

module.exports = { classify, jsonLogEntry, pgLogEntry, appendJsonLog, appendContactFile, writeJsonMode, writePgMode };
