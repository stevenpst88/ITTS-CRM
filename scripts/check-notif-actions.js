#!/usr/bin/env node
'use strict';
/**
 * 鈴鐺通知「清除」與「導向」檢查（feat/bell-actions）：
 *   單元 unit  lib/notifActions.js（與基準版 calcStatus 隨機對拍、已處理鍵含迄日、清理上限、本人檢查、不改動原陣列）
 *   API   api   真實 server.js：DELETE /api/notifications/:id 與 /read（本人／他人 404／批次／未登入／唯讀／集團角色）、
 *               POST /api/contract-reminders/:id/dismiss（可視範圍、ETag 304 不誤用舊資料、改迄日後重新提醒、GET 不寫入、清理上限）
 *   開關  off   未使用新功能時，既有端點輸出與基準提交（git archive）逐位元相同（新增欄位 contractType 除外，測試內明列）
 *   UI    ui    headless Edge＋CDP 真實點擊（scripts/_notif-actions-ui.js）：✕／清除已讀／已處理／合約三種 type 導向／已刪除友善訊息／
 *               callin、lead、receivable 導向／#contract: 深連結／XSS／四種外觀對比／手機 375px／截圖（scratchpad/bell_shots/）
 *   變異 NTF_MUT=1  每個變異（後端與前端）都必須讓對應測試變紅
 *
 *   node scripts/check-notif-actions.js                    全部群組（不含變異）
 *   NTF_ONLY=unit,api,off,ui node ...                       只跑指定群組
 *   NTF_MUT=1 node ...                                      另外跑變異；NTF_MUT_ONLY=M1,M4 只跑指定變異（名稱開頭比對）
 *   NTF_PORT（預設 3452；限 3450–3459，基準版伺服器用 +1）／NTF_CDP（預設 9752；限 9750–9759）／NTF_BASE_REF（預設 8ca4fac）
 * 一次只跑一個。不含任何帳密：臨時 _s_* 帳號與隨機密碼只在記憶體；資料檔只存在於本次執行；全部假資料；結束清掉伺服器、Edge、臨時檔。
 * 守門：工作樹已有 auth.json／data.json／audit.log.json 時拒絕執行（不碰真實資料，不碰 3020）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const { spawn, spawnSync } = require('child_process');

const PORT = parseInt(process.env.NTF_PORT || '3452', 10);
if (!(PORT >= 3450 && PORT <= 3458)) { console.log('拒絕：此腳本只能用埠 3450–3458（基準版伺服器用 +1）'); process.exit(2); }
const CDP = parseInt(process.env.NTF_CDP || '9752', 10);
if (!(CDP >= 9750 && CDP <= 9759)) { console.log('拒絕：此腳本只能用 CDP 埠 9750–9759'); process.exit(2); }
process.env.B2_PORT = String(PORT);
const H = require('./_b2-harness');
const { assertPortFree } = require('./_port-free');
const ROOT = H.ROOT;
const BASE_REF = process.env.NTF_BASE_REF || '8ca4fac';
const ONLY = (process.env.NTF_ONLY || '').split(',').map((s) => s.trim()).filter(Boolean);
const want = (k) => !ONLY.length || ONLY.includes(k);
const sleep = H.sleep;
const J = JSON.stringify;

// ───────────────── 結果收集 ─────────────────
let res = [];
let quiet = false;
function t(name, cond, extra) {
  res.push([name, !!cond, extra == null ? '' : String(extra)]);
  if (!cond && !quiet) console.log('  ✗ ' + name + (extra != null && extra !== '' ? '｜' + String(extra).slice(0, 400) : ''));
}
const teq = (name, a, b) => t(name, J(a) === J(b), `得到 ${J(a)}，預期 ${J(b)}`);

// ───────────────── 假資料 ─────────────────
const pad = (n) => String(n).padStart(2, '0');
const dstr = (off) => { const d = new Date(); d.setHours(12, 0, 0, 0); d.setDate(d.getDate() + off); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const XSS1 = '<img src=x onerror=window.__xss=1>', XSS2 = '<script>window.__xss=2</script>';
const U = { adm: '_s_adm', m1: '_s_m1', u1: '_s_u1', u2: '_s_u2', u3: '_s_u3', u4: '_s_u4', sec: '_s_sec', mkt: '_s_mkt', gs: '_s_gs', tp: '_s_tp', vw: '_s_vw' };
function seedUsers() {
  return [
    H.mkUser(U.adm, 'admin', { displayName: '測試管理員' }),
    H.mkUser(U.m1, 'manager1', { bu: ['ERP'], displayName: '測試主管' }),
    H.mkUser(U.u1, 'user', { bu: ['ERP'], supervisor: U.m1, displayName: '測試業務甲' }),
    H.mkUser(U.u2, 'user', { bu: ['ERP'], supervisor: U.m1, displayName: '測試業務乙' }),
    H.mkUser(U.u3, 'user', { bu: ['ITS'], displayName: '測試業務丙' }),
    H.mkUser(U.u4, 'user', { bu: ['ERP'], supervisor: U.m1, displayName: '測試業務丁' }),
    H.mkUser(U.sec, 'secretary', { bu: ['ERP'], displayName: '測試秘書' }),
    H.mkUser(U.mkt, 'marketing', { bu: ['ERP'], displayName: '測試行銷' }),
    H.mkUser(U.gs, 'groupsales', { bu: ['ERP'], viewOwnerScope: U.u1, viewGroupId: 'g1', displayName: '測試集團業務' }),
    H.mkUser(U.tp, 'tecopm', { viewOwnerScope: U.u1, viewGroupId: 'g1', displayName: '測試集團PM' }),
    H.mkUser(U.vw, 'user', { bu: ['ERP'], supervisor: U.m1, accessMode: 'view', displayName: '測試唯讀業務' }),
  ];
}
const NOW = Date.now();
const iso = (minAgo) => new Date(NOW - minAgo * 60000).toISOString();
const mkContract = (id, owner, type, company, endOff, extra) => Object.assign({
  id, owner, contractNo: 'NT-' + id, company, contactName: '', product: 'SAP License MA', startDate: dstr(-300), endDate: dstr(endOff), renewDate: '', amount: '100', salesPerson: '', note: '', type, createdAt: iso(100000),
}, extra || {});
/** 通知：minAgo 愈小愈新（陣列順序＝新到舊，與 pushNotification 的 unshift 一致） */
const mkN = (id, to, type, title, read, refId, minAgo) => ({ id, to, type, title, body: title + '｜內文', refId: refId === undefined ? '' : refId, read: !!read, createdAt: iso(minAgo == null ? 10 : minAgo) });
function seedData(opts) {
  opts = opts || {};
  const d = {
    contacts: [], companies: [], opportunities: [], visits: [], receivables: [], callins: [], keyAccounts: [], quotations: [], groups: [], industries: [],
    contracts: [
      mkContract('ct-erp1', U.u1, 'ERP_MA', 'ERP 測試客戶', 20),
      mkContract('ct-sap1', U.u1, 'SAP_MA', 'SAP 測試客戶', -10),
      mkContract('ct-svc1', U.u1, 'SERVICE_MA', 'Service 測試客戶', 60),
      mkContract('ct-xss', U.u1, 'ERP_MA', XSS1 + '公司', 40),
      mkContract('ct-far', U.u1, 'ERP_MA', '很久以後才到期', 300),
      mkContract('ct-u2', U.u2, 'ERP_MA', '業務乙的客戶', 15),
      mkContract('ct-u3', U.u3, 'ERP_MA', '另一個 BU 的客戶', 15),
      mkContract('ct-del', U.u1, 'ERP_MA', '已刪除合約的客戶', 10, { deleted: true, deletedAt: iso(50), deletedBy: U.u1 }),
    ],
    notifications: [
      mkN('n1', U.u1, 'quote_approved', 'N1 報價單已核准', false, 'q-00000001', 1),
      mkN('n2', U.u1, 'deal_task', 'N2 成交簽約確認', true, 'task-1', 2),
      mkN('n3', U.u1, 'hotline_closed', 'N3 Hotline 結案', true, 'hc-1', 3),
      mkN('n4', U.u1, 'contract_expiring_30', 'N4 合約 30 天內到期預警', false, 'ct-erp1', 4),
      mkN('n5', U.u1, 'contract_expiring_7', 'N5 合約即將到期', false, 'ct-sap1', 5),
      mkN('n6', U.u1, 'contract_expiring_7', 'N6 Service MA 合約即將到期', true, 'ct-svc1', 6),
      mkN('n7', U.u1, 'contract_urgent', 'N7 合約已不存在', false, 'ct-gone', 7),
      mkN('n8', U.u1, 'callin_new', 'N8 新 Call-in Pass', false, 'ci-1', 8),
      mkN('n9', U.u1, 'lead_assigned', 'N9 新 Lead 指派', false, 'ld-1', 9),
      mkN('n10', U.u1, 'receivable_overdue', 'N10 帳款逾期未收', false, 'rc-1', 10),
      mkN('n11', U.u1, 'birthday', 'N11 客戶生日', false, 'c-1', 11),
      mkN('n12', U.u1, 'admin_account_change', XSS1 + 'N12 惡意標題', false, '', 12),
      mkN('u2a', U.u2, 'quote_approved', 'U2A 業務乙的未讀', false, 'q-00000002', 1),
      mkN('u2b', U.u2, 'deal_task', 'U2B 業務乙的已讀', true, 'task-2', 2),
      mkN('m1a', U.m1, 'contract_expiring_30', 'M1A 主管的已讀', true, 'ct-u2', 3),
      mkN('mk1', U.mkt, 'lead_assigned', 'MK1 行銷的新 Lead', false, 'ld-2', 1),
      mkN('vw1', U.vw, 'admin_account_change', 'VW1 唯讀者的通知', false, '', 1),
    ],
  };
  d.notifications[11].body = XSS2;
  // u4：25 則一般通知（鈴鐺只顯示 20 則；刪一則後第 21 則要補上來）
  for (let i = 1; i <= 25; i++) d.notifications.push(mkN('b' + pad(i), U.u4, 'admin_account_change', 'B' + pad(i) + ' 批量通知', i > 10, '', 100 + i));
  if (opts.data) opts.data(d);
  return d;
}
const LOGIN_ALL = [U.adm, U.m1, U.u1, U.u2, U.u3, U.u4, U.sec, U.mkt, U.gs, U.tp, U.vw];

// ───────────────── 伺服器 ─────────────────
async function fresh(opts) {
  opts = opts || {};
  await H.fresh(seedUsers(), seedData(opts), { env: opts.env || {}, file: opts.file });
}
const readData = () => JSON.parse(fs.readFileSync(path.join(ROOT, 'data.json'), 'utf8'));
const dataHash = () => crypto.createHash('sha1').update(fs.readFileSync(path.join(ROOT, 'data.json'))).digest('hex');
async function rawGet(who, url, hdr) {
  const r = await fetch(H.BASE + url, { headers: Object.assign({ Cookie: H.jar[who] || '' }, hdr || {}) });
  const text = await r.text();
  let json = null; try { json = JSON.parse(text); } catch { /* 非 JSON */ }
  return { status: r.status, etag: r.headers.get('etag'), text, json };
}
const poll = async (who) => (await H.http(who, 'GET', '/api/poll-bundle')).json || {};
const idsOf = (arr) => (Array.isArray(arr) ? arr : []).map((x) => x.id);
const notifIds = async (who) => idsOf((await H.http(who, 'GET', '/api/notifications')).json);
const crIds = (b) => ((b && b.contractReminders) || []).map((x) => x.contractId).sort();

// ═════════════ 單元 ═════════════
function groupUnit(NA) {
  NA = NA || require('../lib/notifActions');
  // U1 與基準版 server.js 內嵌的 calcStatus 隨機對拍
  const baseSrc = spawnSync('git', ['show', BASE_REF + ':server.js'], { cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 }).stdout || '';
  const i0 = baseSrc.indexOf('  const calcStatus = (c) => {');
  const i1 = baseSrc.indexOf('\n  };', i0);
  t('U0 能從基準版取出原本內嵌的 calcStatus', i0 > 0 && i1 > i0);
  let oldFn = null;
  if (i0 > 0 && i1 > i0) oldFn = new Function('today', baseSrc.slice(i0, i1 + 5) + '\n return calcStatus;');
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const oldCalc = oldFn ? oldFn(today) : null;
  let diff = 0, n = 0, rnd = 12345;
  const rand = () => { rnd = (rnd * 1103515245 + 12345) & 0x7fffffff; return rnd / 0x7fffffff; };
  for (let k = 0; k < 4000 && oldCalc; k++) {
    const c = {};
    if (rand() > 0.05) c.endDate = dstr(Math.floor(rand() * 500) - 200);
    if (rand() > 0.5) c.renewDate = dstr(Math.floor(rand() * 500) - 100);
    const a = oldCalc(c), b = NA.contractReminderStatus(c, today);
    n++;
    if (J(a) !== J(b)) diff++;
  }
  t('U1 contractReminderStatus 與基準版 calcStatus 對拍 4000 組完全相同', n === 4000 && diff === 0, `不同 ${diff}／${n}`);
  teq('U1b 沒有迄日 → null', NA.contractReminderStatus({}, today), null);
  // U2 鍵
  teq('U2 dismissKey 格式', NA.dismissKey('u1', 'c1', '2026-12-31'), 'u1|c1|2026-12-31');
  t('U2b 迄日不同鍵不同', NA.dismissKey('u1', 'c1', '2026-12-31') !== NA.dismissKey('u1', 'c1', '2027-12-31'));
  t('U2c 使用者不同鍵不同', NA.dismissKey('u1', 'c1', 'x') !== NA.dismissKey('u2', 'c1', 'x'));
  // U3 過濾
  const rem = [{ contractId: 'c1', endDate: 'E1' }, { contractId: 'c2', endDate: 'E2' }];
  t('U3 沒有紀錄：原陣列原樣回傳（同一個物件）', NA.filterDismissed(undefined, 'u1', rem) === rem && NA.filterDismissed({}, 'u1', rem) === rem);
  const map = { 'u1|c1|E1': 'x' };
  teq('U3b 本人已處理的被濾掉', NA.filterDismissed(map, 'u1', rem).map((r) => r.contractId), ['c2']);
  teq('U3c 別人已處理不影響我', NA.filterDismissed(map, 'u2', rem).map((r) => r.contractId), ['c1', 'c2']);
  teq('U3d 迄日改了就重新提醒', NA.filterDismissed(map, 'u1', [{ contractId: 'c1', endDate: 'E1b' }]).map((r) => r.contractId), ['c1']);
  t('U3e 原型污染鍵不會誤判（__proto__／constructor）', NA.isDismissed({}, 'u1', '__proto__', 'x') === false && NA.isDismissed({}, 'constructor', 'a', 'b') === false);
  // U4 清理
  const nowIso = '2026-10-26T00:00:00.000Z';
  const d4 = { contractReminderDismissals: { 'u1|old|2025-01-01': 'a', 'u1|edge|2025-09-22': 'b', 'u1|new|2026-11-30': 'c', 'u1|odd|不是日期': 'd' } };
  NA.recordDismissal(d4, 'u1', 'c9', '2026-12-01', nowIso);
  const k4 = Object.keys(d4.contractReminderDismissals).sort();
  t('U4 迄日早於今天 400 天以上者被清掉、其餘保留（含無法解析者）', !k4.includes('u1|old|2025-01-01') && k4.includes('u1|edge|2025-09-22') && k4.includes('u1|new|2026-11-30') && k4.includes('u1|odd|不是日期') && k4.includes('u1|c9|2026-12-01'), k4.join(','));
  const d5 = { contractReminderDismissals: {} };
  for (let i = 0; i < 5005; i++) d5.contractReminderDismissals['u|c' + i + '|2027-01-01'] = new Date(Date.UTC(2026, 0, 1) + i * 1000).toISOString();
  NA.recordDismissal(d5, 'u', 'cnew', '2027-01-01', nowIso);
  const k5 = Object.keys(d5.contractReminderDismissals);
  t('U4b 超過 5000 筆時由最舊的開始丟（剛寫的保留）', k5.length === NA.DISMISS_CAP && !('u|c0|2027-01-01' in d5.contractReminderDismissals) && ('u|cnew|2027-01-01' in d5.contractReminderDismissals) && ('u|c5004|2027-01-01' in d5.contractReminderDismissals), k5.length);
  const d6 = { contractReminderDismissals: [] };
  NA.recordDismissal(d6, 'u', 'c', 'e', nowIso);
  t('U4c 儲存值被改成陣列時重建為物件', !Array.isArray(d6.contractReminderDismissals) && Object.keys(d6.contractReminderDismissals).length === 1);
  // U5 removeNotification
  const nl = [{ id: 'a', to: 'u1' }, { id: 'b', to: 'u2' }, { id: 'c', to: 'u1' }];
  const d7 = { notifications: nl };
  t('U5 刪自己的一則', NA.removeNotification(d7, 'a', 'u1') === true && d7.notifications.map((x) => x.id).join() === 'b,c');
  t('U5b 不改動原陣列（回傳新陣列）', nl.length === 3 && d7.notifications !== nl);
  const before = d7.notifications;
  t('U5c 別人的：回 false 且資料不動', NA.removeNotification(d7, 'b', 'u1') === false && d7.notifications === before && d7.notifications.length === 2);
  t('U5d 不存在／沒有陣列：false', NA.removeNotification(d7, 'zzz', 'u1') === false && NA.removeNotification({}, 'a', 'u1') === false);
  // U6 removeRead
  const d8 = { notifications: [{ id: '1', to: 'u1', read: true }, { id: '2', to: 'u1', read: false }, { id: '3', to: 'u2', read: true }, { id: '4', to: 'u1', read: true }] };
  t('U6 只清自己的已讀並回筆數', NA.removeRead(d8, 'u1') === 2 && d8.notifications.map((x) => x.id).join() === '2,3');
  const ref8 = d8.notifications;
  t('U6b 沒有可清的：回 0 且不換陣列', NA.removeRead(d8, 'u1') === 0 && d8.notifications === ref8);
  t('U6c 沒有陣列：回 0', NA.removeRead({}, 'u1') === 0);
  t('U7 isValidId', NA.isValidId('abc-1_X') && !NA.isValidId('') && !NA.isValidId('a/b') && !NA.isValidId('x'.repeat(65)) && !NA.isValidId(null) && !NA.isValidId('<img>'));
}

// ═════════════ API ═════════════
async function groupApi(opts) {
  opts = opts || {};
  await fresh({ file: opts.file });
  await H.loginAll(LOGIN_ALL);
  const U1 = U.u1;

  // ── 儲存型通知：刪除自己的 ──
  const before = await notifIds(U1);
  t('A0 種子：u1 看得到自己的 12 則、看不到別人的', before.length === 12 && !before.includes('u2a'), before.join());
  const del1 = await H.http(U1, 'DELETE', '/api/notifications/n1');
  t('A1 DELETE 自己的 → 200 {success:true}', del1.status === 200 && del1.json && del1.json.success === true, J(del1.json));
  const after1 = await notifIds(U1);
  t('A1b 清單少了那一則、其餘不變', !after1.includes('n1') && after1.length === 11);
  t('A1c data.json 真的刪了（不是只在記憶體）', !readData().notifications.some((n) => n.id === 'n1') && readData().notifications.some((n) => n.id === 'u2a'));
  // ── 他人的／不存在 ──
  const o1 = await H.http(U1, 'DELETE', '/api/notifications/u2a');
  const o2 = await H.http(U1, 'DELETE', '/api/notifications/no-such-id');
  t('A2 刪別人的 → 404', o1.status === 404, o1.status);
  t('A2b 別人的與不存在回應完全相同（不洩漏他人通知是否存在）', o2.status === 404 && J(o1.json) === J(o2.json) && o1.text === o2.text, o1.text + ' vs ' + o2.text);
  t('A2c 別人的通知沒被動到', (await notifIds(U.u2)).includes('u2a'));
  // ── 已刪過再刪 ──
  t('A2d 重複刪除同一則 → 404（不是 500）', (await H.http(U1, 'DELETE', '/api/notifications/n1')).status === 404);
  // ── 批次清除已讀 ──
  const rb = await H.http(U1, 'DELETE', '/api/notifications/read');
  t('A3 DELETE /read → 200 {success:true, removed:n}（n2、n3、n6 三則已讀）', rb.status === 200 && rb.json && rb.json.success === true && rb.json.removed === 3, J(rb.json));
  const after3 = await notifIds(U1);
  t('A3b 已讀的不見、未讀的都在', !['n2', 'n3', 'n6'].some((i) => after3.includes(i)) && ['n4', 'n5', 'n7', 'n8', 'n9', 'n10', 'n11', 'n12'].every((i) => after3.includes(i)), after3.join());
  t('A3c 別人已讀的沒被清（u2b、m1a）', (await notifIds(U.u2)).includes('u2b') && (await notifIds(U.m1)).includes('m1a'));
  const rb2 = await H.http(U1, 'DELETE', '/api/notifications/read');
  t('A3d 再清一次：removed 0、仍是 200', rb2.status === 200 && rb2.json.removed === 0, J(rb2.json));
  t('A3e /read 沒有被當成通知 id（路由順序）：回的是 removed 而不是 404', rb2.json && 'removed' in rb2.json);
  // ── 權限 ──
  const un = await fetch(H.BASE + '/api/notifications/n4', { method: 'DELETE' });
  t('A4 未登入刪除 → 401', un.status === 401, un.status);
  const un2 = await fetch(H.BASE + '/api/notifications/read', { method: 'DELETE' });
  t('A4b 未登入批次清除 → 401', un2.status === 401, un2.status);
  const sec = await H.http(U.sec, 'DELETE', '/api/notifications/read');
  t('A5 秘書可清除自己的（0 則也 200）', sec.status === 200, sec.status);
  const gs = await H.http(U.gs, 'DELETE', '/api/notifications/n4');
  const tp = await H.http(U.tp, 'DELETE', '/api/notifications/n4');
  t('A6 集團業務／集團 PM 被既有路徑白名單擋（403，沒有放行）', gs.status === 403 && tp.status === 403, gs.status + '/' + tp.status);
  const vw = await H.http(U.vw, 'DELETE', '/api/notifications/read');
  t('A7 唯讀模式帳號與既有「標已讀」一致：403', vw.status === 403 && (await H.http(U.vw, 'PUT', '/api/notifications/read-all')).status === 403, vw.status);
  const ra = await H.http(U.u2, 'PUT', '/api/notifications/read-all');
  t('A8 既有 PUT read-all 不受影響', ra.status === 200 && ra.json.success === true);
  t('A8b 既有 PUT :id/read 不受影響', (await H.http(U1, 'PUT', '/api/notifications/n4/read')).status === 200 && (await H.http(U1, 'GET', '/api/notifications')).json.find((n) => n.id === 'n4').read === true);

  // ── 動態合約提醒：dismiss ──
  await fresh({ file: opts.file });
  await H.loginAll(LOGIN_ALL);
  const b0 = await poll(U1);
  t('D1 poll-bundle 的合約提醒帶 contractId 與 contractType', (b0.contractReminders || []).length === 5 && b0.contractReminders.every((r) => r.contractId && r.contractType), crIds(b0).join());
  const tmap = Object.fromEntries((b0.contractReminders || []).map((r) => [r.contractId, r.contractType]));
  t('D1b contractType 對應合約實際類型', tmap['ct-erp1'] === 'ERP_MA' && tmap['ct-sap1'] === 'SAP_MA' && tmap['ct-svc1'] === 'SERVICE_MA', J(tmap));
  const cr0 = (await H.http(U1, 'GET', '/api/contract-reminders')).json;
  t('D1c /api/contract-reminders 同樣帶 contractType', Array.isArray(cr0) && cr0.length === 5 && cr0.every((r) => r.contractType));
  // GET 不寫入
  const h0 = dataHash();
  await poll(U1); await H.http(U1, 'GET', '/api/contract-reminders'); await poll(U.m1);
  t('D2 GET poll-bundle／contract-reminders 不寫入 data.json', dataHash() === h0);
  // ETag 流程
  const g1 = await rawGet(U1, '/api/poll-bundle');
  const g304 = await rawGet(U1, '/api/poll-bundle', { 'If-None-Match': g1.etag });
  t('D3 沒變時帶 ETag → 304', g1.status === 200 && g1.etag && g304.status === 304, g304.status);
  const dm = await H.http(U1, 'POST', '/api/contract-reminders/ct-erp1/dismiss');
  t('D4 dismiss → 200 {success:true}', dm.status === 200 && dm.json && dm.json.success === true, J(dm.json));
  const g2 = await rawGet(U1, '/api/poll-bundle', { 'If-None-Match': g1.etag });
  t('D5 處理後帶舊 ETag → 200 新內容（不是 304 沿用舊的）', g2.status === 200 && g2.etag !== g1.etag, g2.status);
  t('D5b 新內容少了那份合約、其餘都在', !crIds(g2.json).includes('ct-erp1') && crIds(g2.json).length === 4, crIds(g2.json).join());
  const g3 = await rawGet(U1, '/api/poll-bundle', { 'If-None-Match': g2.etag });
  t('D5c 用新 ETag → 304', g3.status === 304);
  t('D6 只影響本人：主管仍看得到該合約的提醒', crIds(await poll(U.m1)).includes('ct-erp1'));
  const g4 = await rawGet(U1, '/api/contract-reminders');
  t('D7 /api/contract-reminders 同樣濾掉', g4.json.length === 4 && !g4.json.some((r) => r.contractId === 'ct-erp1'));
  const store = readData().contractReminderDismissals || {};
  const k = Object.keys(store);
  const c1end = readData().contracts.find((c) => c.id === 'ct-erp1').endDate;
  t('D8 紀錄鍵＝本人|合約|迄日，值是 ISO 時間', k.length === 1 && k[0] === `${U1}|ct-erp1|${c1end}` && !isNaN(Date.parse(store[k[0]])), k.join());
  const dm2 = await H.http(U1, 'POST', '/api/contract-reminders/ct-erp1/dismiss');
  t('D9 重複處理冪等（200、紀錄不增加）', dm2.status === 200 && Object.keys(readData().contractReminderDismissals).length === 1);
  // 改迄日 → 重新提醒
  const ct = readData().contracts.find((c) => c.id === 'ct-erp1');
  const put = await H.http(U1, 'PUT', '/api/contracts/ct-erp1', Object.assign({}, ct, { endDate: dstr(30) }));
  t('D10a 前置：改迄日 PUT 成功', put.status === 200, put.status + ' ' + put.text.slice(0, 120));
  t('D10 改了迄日 → 提醒自動重新出現', crIds(await poll(U1)).includes('ct-erp1'));
  const put2 = await H.http(U1, 'PUT', '/api/contracts/ct-erp1', Object.assign({}, ct, { endDate: c1end, note: '改回原迄日' }));
  t('D10b 迄日改回原值 → 先前的已處理紀錄又生效（鍵相同）', put2.status === 200 && !crIds(await poll(U1)).includes('ct-erp1'));
  // 可視範圍
  const nf = await H.http(U1, 'POST', '/api/contract-reminders/ct-gone/dismiss');
  const other = await H.http(U1, 'POST', '/api/contract-reminders/ct-u2/dismiss');
  const other3 = await H.http(U1, 'POST', '/api/contract-reminders/ct-u3/dismiss');
  t('D11 不存在 → 404', nf.status === 404, nf.status);
  t('D11b 別人的合約（看不到）→ 404，與不存在相同回應', other.status === 404 && other.text === nf.text && other3.status === 404, other.status);
  t('D11c 看不到的合約沒有留下紀錄', !Object.keys(readData().contractReminderDismissals || {}).some((x) => x.includes('ct-u2') || x.includes('ct-u3')));
  const farr = await H.http(U1, 'POST', '/api/contract-reminders/ct-far/dismiss');
  t('D12 目前沒有提醒的合約（90 天以外）→ 404', farr.status === 404, farr.status);
  t('D12b 秘書 → 404', (await H.http(U.sec, 'POST', '/api/contract-reminders/ct-erp1/dismiss')).status === 404);
  t('D12c 非法 id（過長／特殊字元）→ 404', (await H.http(U1, 'POST', '/api/contract-reminders/' + 'x'.repeat(65) + '/dismiss')).status === 404 && (await H.http(U1, 'POST', '/api/contract-reminders/%3Cimg%3E/dismiss')).status === 404);
  t('D12d 未登入 → 401', (await fetch(H.BASE + '/api/contract-reminders/ct-erp1/dismiss', { method: 'POST' })).status === 401);
  t('D12e 集團業務／集團 PM → 403（既有白名單）', (await H.http(U.gs, 'POST', '/api/contract-reminders/ct-erp1/dismiss')).status === 403 && (await H.http(U.tp, 'POST', '/api/contract-reminders/ct-erp1/dismiss')).status === 403);
  t('D12f 唯讀模式 → 403', (await H.http(U.vw, 'POST', '/api/contract-reminders/ct-erp1/dismiss')).status === 403);
  // 主管看得到下屬合約，可處理，只影響自己
  const mdm = await H.http(U.m1, 'POST', '/api/contract-reminders/ct-u2/dismiss');
  t('D13 主管可處理下屬合約的提醒（可視範圍內）', mdm.status === 200 && !crIds(await poll(U.m1)).includes('ct-u2'));
  t('D13b 業務乙自己的提醒不受主管處理影響', crIds(await poll(U.u2)).includes('ct-u2'));
  // 已刪除合約的提醒（既有行為：仍會提醒）也能處理
  t('D14 已刪除合約的提醒（既有行為仍會出現）可被處理掉', crIds(await poll(U1)).includes('ct-del') && (await H.http(U1, 'POST', '/api/contract-reminders/ct-del/dismiss')).status === 200 && !crIds(await poll(U1)).includes('ct-del'));
  // 清理上限（POST 時才做）
  await fresh({ file: opts.file, data: (d) => {
    d.contractReminderDismissals = {};
    for (let i = 0; i < 5100; i++) d.contractReminderDismissals[`${U.u3}|x${i}|${dstr(100)}`] = new Date(Date.UTC(2026, 0, 1) + i * 1000).toISOString();
    d.contractReminderDismissals[`${U.u3}|stale|${dstr(-500)}`] = iso(5);
  } });
  await H.loginAll([U.u1, U.u3]);
  const h1 = dataHash(); await poll(U1);
  t('D15 5100 筆紀錄時 GET 不清理（不寫入）', dataHash() === h1 && Object.keys(readData().contractReminderDismissals).length === 5101);
  await H.http(U1, 'POST', '/api/contract-reminders/ct-sap1/dismiss');
  const st = readData().contractReminderDismissals;
  t('D15b POST 時清理：迄日早於 400 天者被清掉，總數 ≤ 5000', Object.keys(st).length <= 5000 && !Object.keys(st).some((x) => x.includes('|stale|')) && Object.keys(st).some((x) => x.startsWith(U1 + '|ct-sap1|')), Object.keys(st).length);
}

// ═════════════ 開關：與基準版逐位元相同 ═════════════
async function startAt(dir, port, extraEnv) {
  const env = Object.assign({}, process.env, { PORT: String(port), JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', ALLOWED_ORIGINS: 'http://localhost:' + port }, extraEnv || {});
  for (const k of ['DATABASE_URL', 'MAIL_MODE', 'FEATURE_GUARD_MODE', 'MA_TRACKING', 'HOTLINE']) delete env[k];
  await assertPortFree(port);
  const child = spawn(process.execPath, [path.join(dir, 'server.js')], { cwd: dir, env, stdio: ['ignore', 'pipe', 'pipe'] });
  let out = ''; child.stdout.on('data', (d) => { out += d; }); child.stderr.on('data', (d) => { out += d; });
  for (let i = 0; i < 80; i++) {
    await sleep(250);
    try { const r = await fetch(`http://localhost:${port}/login.html`); if (r.ok) return child; } catch { /* 尚未起來 */ }
    if (child.exitCode !== null) throw new Error('基準伺服器提前結束：' + out.slice(-400));
  }
  throw new Error('基準伺服器啟動逾時：' + out.slice(-400));
}
async function groupOff() {
  await H.stopServer();
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ntf_base_'));
  let child = null;
  try {
    const ar = spawnSync('git', ['archive', BASE_REF], { cwd: ROOT, maxBuffer: 256 * 1024 * 1024 });
    t('O0 git archive 基準提交 ' + BASE_REF, ar.status === 0 && ar.stdout.length > 1000);
    const tarf = path.join(dir, '_base.tar'); fs.writeFileSync(tarf, ar.stdout);
    const tx = spawnSync('tar', ['-xf', '_base.tar'], { cwd: dir }); fs.unlinkSync(tarf);
    t('O0b 解開基準版', tx.status === 0 && fs.existsSync(path.join(dir, 'server.js')));
    const mk = spawnSync('cmd', ['/c', 'mklink', '/J', path.join(dir, 'node_modules').replace(/\//g, '\\'), path.join(ROOT, 'node_modules').replace(/\//g, '\\')], { encoding: 'utf8' });
    t('O0c node_modules junction', mk.status === 0 && fs.existsSync(path.join(dir, 'node_modules', 'bcryptjs')), mk.stdout + mk.stderr);
    // 同一份種子寫到兩邊（密碼雜湊、通知、合約完全相同）
    const users = seedUsers(), data = seedData();
    data.notifications.push(mkN('rc-x', U.u1, 'receivable_overdue', '另一則', false, 'rc-9', 20));
    for (const d of [ROOT, dir]) {
      fs.writeFileSync(path.join(d, 'auth.json'), J({ users }, null, 2));
      fs.writeFileSync(path.join(d, 'data.json'), J(data, null, 2));
    }
    child = await startAt(dir, PORT + 1);
    await H.startServer({});
    const jb = {}, jc = {};
    const lg = async (base, jr, who) => {
      const r = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: J({ username: who, password: H.PW }) });
      jr[who] = (r.headers.getSetCookie ? r.headers.getSetCookie() : []).map((c) => c.split(';')[0]).join('; ');
      return r.status;
    };
    const BB = 'http://localhost:' + (PORT + 1), CB = H.BASE;
    const who = [U.u1, U.m1, U.u2, U.u3, U.sec, U.mkt, U.adm];
    for (const w of who) { await lg(BB, jb, w); await lg(CB, jc, w); }
    const get = async (base, jr, w, url) => { const r = await fetch(base + url, { headers: { Cookie: jr[w] } }); const text = await r.text(); let json = null; try { json = JSON.parse(text); } catch { /* */ } return { status: r.status, text, json }; };
    const stripCT = (o) => { // 新增欄位 contractType：只在合約提醒物件上，比對前拿掉
      if (o && Array.isArray(o.contractReminders)) o.contractReminders.forEach((r) => { delete r.contractType; });
      if (Array.isArray(o)) o.forEach((r) => { if (r && typeof r === 'object') delete r.contractType; });
      return o;
    };
    let diffs = [];
    for (const w of who) {
      for (const url of ['/api/notifications', '/api/poll-bundle?days=3', '/api/contract-reminders', '/api/contracts', '/api/me']) {
        const a = await get(BB, jb, w, url), b = await get(CB, jc, w, url);
        if (url === '/api/me') { t(`O1 ${w} ${url} 狀態碼相同`, a.status === b.status); continue; }
        const aj = J(stripCT(a.json)), bj = J(stripCT(b.json));
        if (a.status !== b.status || aj !== bj) diffs.push(`${w} ${url} ${a.status}/${b.status}`);
      }
    }
    t('O1 7 個帳號 × 4 個既有端點：輸出（拿掉 contractType）與基準版逐位元相同', diffs.length === 0, diffs.join('；'));
    // contractType 是唯一差異：拿掉前確實有、基準版沒有
    const nb = (await get(CB, jc, U.u1, '/api/poll-bundle?days=3')).json, ob = (await get(BB, jb, U.u1, '/api/poll-bundle?days=3')).json;
    t('O2 唯一差異是新增的 contractType 欄位（本版有、基準沒有）', nb.contractReminders.length > 0 && nb.contractReminders.every((r) => r.contractType) && ob.contractReminders.every((r) => !('contractType' in r)));
    // 既有寫入端點：標已讀後兩邊仍相同
    for (const [base, jr] of [[BB, jb], [CB, jc]]) {
      await fetch(base + '/api/notifications/n4/read', { method: 'PUT', headers: { Cookie: jr[U.u1] } });
      await fetch(base + '/api/notifications/read-all', { method: 'PUT', headers: { Cookie: jr[U.u2] } });
    }
    diffs = [];
    for (const w of [U.u1, U.u2]) {
      const a = await get(BB, jb, w, '/api/notifications'), b = await get(CB, jc, w, '/api/notifications');
      if (a.text !== b.text) diffs.push(w);
    }
    t('O3 標已讀（單則／全部）後通知清單與基準版相同', diffs.length === 0, diffs.join());
    // 基準版沒有新端點；本版有
    const nd = await fetch(BB + '/api/notifications/n5', { method: 'DELETE', headers: { Cookie: jb[U.u1] } });
    t('O4 基準版確實沒有 DELETE（對照用，證明比對的是舊版）', nd.status === 404);
    // 沒使用新功能時，poll-bundle ETag 流程仍正常（不因過濾改變 304 行為）
    const e1 = await fetch(CB + '/api/poll-bundle', { headers: { Cookie: jc[U.u1] } });
    const e2 = await fetch(CB + '/api/poll-bundle', { headers: { Cookie: jc[U.u1], 'If-None-Match': e1.headers.get('etag') } });
    t('O5 沒使用新功能時 ETag／304 照舊', e1.headers.get('etag') && e2.status === 304);
    // 合約提醒的 id 順序與內容（排序）完全相同已含於 O1
    t('O6 本版沒有在 GET 時多寫 contractReminderDismissals', !('contractReminderDismissals' in readData()));
  } finally {
    await H.stopServer();
    if (child) { try { process.kill(child.pid); } catch { /* 已結束 */ } await sleep(400); }
    H.cleanupFiles();
    try { spawnSync('cmd', ['/c', 'rmdir', path.join(dir, 'node_modules').replace(/\//g, '\\')]); } catch { /* 忽略 */ }
    if (!fs.existsSync(path.join(dir, 'node_modules'))) { try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* 忽略 */ } }
    else console.log('警告：基準版 node_modules junction 未移除，保留目錄 ' + dir);
  }
}

// ═════════════ 變異 ═════════════
const SERVER_MUTS = [
  { name: 'M1 拿掉本人檢查（刪別人的通知）', file: 'lib/notifActions.js', edits: [{ find: 'n && n.id === id && n.to === username', rep: 'n && n.id === id' }], expect: /^A2/ },
  { name: 'M2 清除已讀不分本人', file: 'lib/notifActions.js', edits: [{ find: 'n && n.to === username && n.read === true', rep: 'n && n.read === true' }], expect: /^A3c/ },
  { name: 'M3 ETag 不含合約提醒（處理後仍回 304）', file: 'server.js', edits: [{ find: "const json = JSON.stringify(body);\n  const etag = 'W/\"' + crypto.createHash('sha1').update(json).digest('base64').slice(0, 22) + '\"';", rep: "const json = JSON.stringify(body);\n  const etag = 'W/\"' + crypto.createHash('sha1').update(JSON.stringify(Object.assign({}, body, { contractReminders: [] }))).digest('base64').slice(0, 22) + '\"';" }], expect: /^D5/ },
  { name: 'M4 已處理鍵不含迄日（改日期不會重新提醒）', file: 'lib/notifActions.js', edits: [{ find: "String(endDate == null ? '' : endDate);", rep: "'';" }], expect: /^D10/ },
  { name: 'M5 GET poll-bundle 會寫入', file: 'server.js', edits: [{ find: "  const body = { notifications, contractReminders, birthdayReminders };", rep: "  data.__touch = Date.now(); db.save(data);\n  const body = { notifications, contractReminders, birthdayReminders };" }], expect: /^D2/ },
  { name: 'M6 dismiss 不檢查可視範圍', file: 'server.js', edits: [{ find: "x.id === id && owners.includes(x.owner) && !(MA.enabled() && x.type === 'SERVICE_MA'));\n  if (!c) return NOT_FOUND();", rep: "x.id === id);\n  if (!c) return NOT_FOUND();" }], expect: /^D11/ },
  { name: 'M7 dismiss 不確認目前是否有提醒', file: 'server.js', edits: [{ find: "if (!NTA.contractReminderStatus(c, today)) return NOT_FOUND();", rep: "" }], expect: /^D12 / },
  { name: 'M8 /api/contract-reminders 沒有過濾已處理', file: 'server.js', edits: [{ find: "res.json(NTA.filterDismissed(data.contractReminderDismissals, username, reminders));", rep: "res.json(reminders);" }], expect: /^D7/ },
  { name: 'M9 DELETE /read 被排到 /:id 之後（路由順序）', file: 'server.js', edits: [{ find: "app.delete('/api/notifications/read', requireAuth,", rep: "app.delete('/api/notifications/readx', requireAuth," }], expect: /^A3/ },
  { name: 'M10 poll-bundle 沒有過濾已處理', file: 'server.js', edits: [{ find: "contractReminders = NTA.filterDismissed(data.contractReminderDismissals, username, contractReminders);", rep: "" }], expect: /^D5 / },
  { name: 'M11 刪除後沒有存檔', file: 'server.js', edits: [{ find: "if (!NTA.removeNotification(data, req.params.id, username)) return res.status(404).json({ error: '找不到這則通知' });\n  db.save(data);", rep: "if (!NTA.removeNotification(data, req.params.id, username)) return res.status(404).json({ error: '找不到這則通知' });" }], expect: /^A1c/ },
  { name: 'M12 清理不丟過期紀錄', file: 'lib/notifActions.js', edits: [{ find: "if (Number.isFinite(end) && end < cutoff) delete m[k];", rep: "" }], expect: /^D15b|^U4/ },
];
function writeMut(m) {
  // 變異後的 server 副本放在工作樹（require 相對路徑才對）；lib 變異另存 lib/notifActions__mut.js 並讓 server 副本改用它
  const srcServer = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  let server = srcServer;
  const files = [];
  const apply = (text, edits) => {
    edits.forEach((e) => { if (!text.includes(e.find)) throw new Error('MUTATION ANCHOR MISSING: ' + e.find.slice(0, 60)); text = text.replace(e.find, () => e.rep); });
    return text;
  };
  if (m.file === 'server.js') server = apply(server, m.edits);
  else {
    const lib = apply(fs.readFileSync(path.join(ROOT, m.file), 'utf8'), m.edits);
    const mp = path.join(ROOT, 'lib', 'notifActions__mut.js'); fs.writeFileSync(mp, lib); files.push(mp);
    server = server.replace("require('./lib/notifActions')", "require('./lib/notifActions__mut')");
  }
  const sp = path.join(ROOT, 'server.__mut.js'); fs.writeFileSync(sp, server); files.push(sp);
  return { serverFile: sp, libFile: m.file === 'server.js' ? null : path.join(ROOT, 'lib', 'notifActions__mut.js'), files };
}
async function groupMut() {
  console.log('[變異] 後端：改壞程式碼，對應測試必須變紅');
  let undetected = 0, total = 0;
  const only = process.env.NTF_MUT_ONLY ? process.env.NTF_MUT_ONLY.split(',') : null;
  for (const m of SERVER_MUTS) {
    if (only && !only.some((p) => m.name.startsWith(p))) continue;
    total++;
    let info;
    try { info = writeMut(m); } catch (e) { console.log('  ✗ 變異錨點不存在：' + m.name); undetected++; continue; }
    res = []; quiet = true;
    try {
      if (info.libFile) groupUnit(require(info.libFile));   // lib 變異：單元測試也跑變異後的模組
      await groupApi({ file: info.serverFile });
    } catch (e) { t('(例外) ' + String(e && e.message).slice(0, 100), false); }
    quiet = false;
    info.files.forEach((f) => { try { fs.unlinkSync(f); } catch { /* 忽略 */ } });
    try { delete require.cache[require.resolve(path.join(ROOT, 'lib', 'notifActions__mut.js'))]; } catch { /* 沒載入過 */ }
    const failed = res.filter((x) => !x[1]);
    const hit = failed.some((x) => m.expect.test(x[0]));
    console.log((hit ? '  ✓ ' : '  ✗ ') + m.name + '：' + failed.length + ' 項變紅' + (hit ? '' : '（預期的測試沒變紅）') + (failed.length ? '｜例：' + failed[0][0].slice(0, 60) : ''));
    if (!hit) undetected++;
  }
  console.log(`後端變異測試：${total - undetected}/${total} 被偵測到`);
  return undetected;
}

// ═════════════ 執行 ═════════════
const X = { t, teq, H, U, ROOT, PORT, CDP, J, seedUsers, seedData, fresh, readData, dataHash, rawGet, poll, dstr, mkN, mkContract, XSS1, XSS2, LOGIN_ALL, iso, get res() { return res; }, set res(v) { res = v; }, setQuiet(q) { quiet = q; } };
module.exports = X;

function report() {
  const bad = res.filter((x) => !x[1]);
  console.log(`\n鈴鐺通知檢查結果：${res.length - bad.length} 通過，${bad.length} 失敗`);
  if (bad.length) console.log(bad.map((x) => ' - ' + x[0] + (x[2] ? '｜' + x[2].slice(0, 200) : '')).join('\n'));
  return bad.length;
}
if (require.main === module) {
  (async () => {
    H.guardClean();
    let undetected = 0, allBad = 0;
    try {
      if (want('unit')) { console.log('[單元]'); groupUnit(); }
      if (want('api')) { console.log('[API]'); await groupApi(); }
      if (want('off')) { console.log('[開關：與基準版比對]'); await groupOff(); }
      allBad += report();
      if (want('ui')) {
        console.log('[UI：headless Edge＋CDP]');
        res = [];
        const UI = require('./_notif-actions-ui');
        await UI.run({ X, shots: process.env.NTF_SHOTS || path.join(ROOT, '..', 'bell_shots'), cdp: CDP }, process.env.NTF_UI_ONLY || '');
        allBad += report();
      }
      if (process.env.NTF_MUT === '1') {
        if (want('unit') || want('api')) undetected += await groupMut();
        if (want('ui')) { const UI = require('./_notif-actions-ui'); undetected += await UI.groupMut({ X, cdp: CDP }); }
      }
    } catch (e) {
      console.log('例外：', e && e.stack || e); res.push(['(例外) ' + (e && e.message), false, '']);
    } finally {
      await H.finish();
      ['server.__mut.js', 'lib/notifActions__mut.js'].forEach((f) => { try { fs.unlinkSync(path.join(ROOT, f)); } catch { /* 沒有 */ } });
    }
    process.exitCode = (allBad || undetected) ? 1 : 0;
  })();
}
