#!/usr/bin/env node
'use strict';
/**
 * BUG-03：逾期提醒（checkOverdueReminders）不再蓋掉通知 — 本地 JSON 後端檢查。不需要瀏覽器，不連網路。
 *
 * 背景：checkOverdueReminders 先 db.load() 拿一份快照，迴圈中每次 pushNotification 又各自 load＋save，最後再用「舊快照」save——
 *       本地 JSON 後端的 load() 每次都讀檔回傳新物件，所以最後那次 save 會把剛推的通知全部蓋掉（只剩提醒標記）。
 *       Postgres 後端 load() 回傳同一份快取物件，所以不受影響，但寫法脆弱。修正後：單次 load → 全部寫進同一個 data → 只 save 一次；
 *       pushNotification 多一個選用的 callerData 參數（有傳＝只改該物件，不 load／不 save／不發 Web Push）。
 *
 * 做法：在 BUG03_ROOT（預設＝本工作樹）建立臨時 auth.json／data.json（假資料＋_s_* 帳號），DB_BACKEND=json，
 *       **在本行程內** require server.js（不 listen），透過 server.js 的 _test 掛鉤直接呼叫 checkOverdueReminders／pushNotification，
 *       再「從檔案重新讀」確認最終資料。跑完刪除臨時檔。
 * 守門：BUG03_ROOT 已有 auth.json／data.json／audit.log.json（可能是真實資料）時拒絕執行。不含任何帳密（密碼只在記憶體）。
 *
 *   node scripts/check-overdue-notify.js                 # 一般檢查（不 listen、不連網，不佔任何埠）
 *   BUG03_MUT=1 node scripts/check-overdue-notify.js     # 變異測試：產生「還原修正」的 server.mut.js（臨時檔）再跑本檢查，必須變紅
 *   BUG03_ROOT=<其他程式樹> node scripts/check-overdue-notify.js     # 對其他程式樹跑（需有 _test 掛鉤）
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const SERVER_FILE = process.env.BUG03_SERVER || 'server.js';
const ROOT = path.resolve(process.env.BUG03_ROOT || path.join(__dirname, '..'));
const FILES = ['auth.json', 'data.json', 'audit.log.json', 'mail-outbox.json', 'audit.log.jsonl', 'audit.failed.jsonl', 'contact_audit.json'].map(f => path.join(ROOT, f));
const guard = FILES.filter(f => /(auth|data|audit.log).jsonl?$/.test(f)).filter(f => fs.existsSync(f));
if (guard.length) { console.log('拒絕：', ROOT, '已有', guard.map(f => path.basename(f)).join('/'), '（可能是真實資料），請在乾淨的 worktree 執行'); process.exit(2); }

// 變異測試：把修正還原（pushNotification 不傳 data＝舊行為）寫成臨時 server.mut.js，用子行程跑本檢查，必須「失敗」才算檢查有效
if (process.env.BUG03_MUT === '1') {
  const src = fs.readFileSync(path.join(ROOT, 'server.js'), 'utf8');
  const needle = 'pushNotification(to, type, title, msg, refId, data)';
  if (!src.includes(needle)) { console.log('變異失敗：找不到修正點，檢查需更新'); process.exit(2); }
  const mut = path.join(ROOT, 'server.mut.js');
  fs.writeFileSync(mut, src.replace(needle, 'pushNotification(to, type, title, msg, refId)'));
  let code = 0;
  try {
    const r = require('child_process').spawnSync(process.execPath, [__filename], { env: Object.assign({}, process.env, { BUG03_MUT: '', BUG03_SERVER: 'server.mut.js' }), encoding: 'utf8' });
    const out = (r.stdout || '') + (r.stderr || '');
    const m = out.match(/通過 (\d+)\s+失敗 (\d+)/);
    console.log('變異版（還原修正）：', m ? m[0] : '（無結果）' + out.slice(-600), '結束碼', r.status);
    if (r.status === 0) { console.log('✗ 變異版竟然通過 → 檢查無效'); code = 1; } else console.log('✓ 變異版如預期變紅');
  } finally { try { fs.unlinkSync(mut); } catch { /* 已不存在 */ } }
  process.exit(code);
}

let pass = 0, fail = 0;
const failures = [];
function ok(cond, name) { if (cond) pass++; else { fail++; failures.push(name); console.log('  ✗ FAIL:', name); } }
function eq(a, b, name) { ok(JSON.stringify(a) === JSON.stringify(b), `${name}（得到 ${JSON.stringify(a)}，預期 ${JSON.stringify(b)}）`); }

const DATA_FILE = path.join(ROOT, 'data.json');
const AUTH_FILE = path.join(ROOT, 'auth.json');
const readData = () => JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));   // 一律從檔案重新讀＝「最終資料檔」
const cleanup = () => FILES.forEach(f => { try { fs.unlinkSync(f); } catch { /* 不存在 */ } });

// 日期：全部用「相對今天」，避免測試過期（比對用本機日曆日，與 server.js 的 setHours(0,0,0,0) 一致）
const ymd = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const addDays = (n) => { const d = new Date(); d.setHours(12, 0, 0, 0); d.setDate(d.getDate() + n); return d; };
const today = new Date();

function seed() {
  const now = new Date().toISOString();
  const bcrypt = require(path.join(__dirname, '..', 'node_modules', 'bcryptjs'));
  const pw = bcrypt.hashSync('Ov' + crypto.randomBytes(9).toString('hex') + '!Z', 6);
  const mk = (username, role, extra) => Object.assign({ username, displayName: username, role, active: true, password: pw, createdAt: now, passwordChangedAt: now, mustChangePassword: false }, extra || {});
  fs.writeFileSync(AUTH_FILE, JSON.stringify({
    users: [
      mk('_s_adm', 'admin'),
      mk('_s_m1', 'manager1', { bu: ['ERP'] }),
      mk('_s_a', 'user', { bu: ['ERP'], supervisor: '_s_m1' }),
      mk('_s_b', 'user', { bu: ['ERP'], supervisor: '_s_m1' }),
    ],
  }, null, 2));
  fs.writeFileSync(DATA_FILE, JSON.stringify({
    contacts: [
      { id: 'c-bd', name: '生日客戶', company: '沙盒公司', owner: '_s_a', bu: 'ERP', personalBirthday: `${today.getMonth() + 1}/${today.getDate()}` },
    ],
    companies: [], opportunities: [], visits: [], quotations: [], callins: [], keyAccounts: [],
    contracts: [
      // D-7 內（每天提醒一次）兩份、D-30 預警一份
      { id: 'k1', owner: '_s_a', company: '甲公司', contractNo: 'K-001', product: 'ERP MA', endDate: ymd(addDays(3)), type: 'ERP_MA' },
      { id: 'k2', owner: '_s_b', company: '乙公司', contractNo: 'K-002', product: 'ERP MA', endDate: ymd(addDays(4)), type: 'ERP_MA' },
      { id: 'k3', owner: '_s_a', company: '丙公司', contractNo: 'K-003', product: 'ERP MA', endDate: ymd(addDays(20)), type: 'ERP_MA' },
    ],
    receivables: [
      // 兩筆逾期未付
      { id: 'r1', owner: '_s_a', company: '丁公司', invoiceNo: 'INV-1', dueDate: ymd(addDays(-10)), status: 'unpaid' },
      { id: 'r2', owner: '_s_b', company: '戊公司', invoiceNo: 'INV-2', dueDate: ymd(addDays(-5)), status: 'unpaid' },
    ],
    notifications: [{ id: 'old-1', to: '_s_a', type: 'x', title: '既有通知', body: '', refId: null, read: false, createdAt: now }],   // 既有通知也不能被洗掉
  }, null, 2));
}

(async () => {
  let code = 0;
  try {
    cleanup(); seed();
    process.chdir(ROOT);
    Object.assign(process.env, { JWT_SECRET: crypto.randomBytes(24).toString('hex'), SESSION_BACKEND: 'jwt', DB_BACKEND: 'json', NODE_ENV: 'development', PORT: '0' });
    for (const k of ['DATABASE_URL', 'MAIL_MODE', 'MA_TRACKING', 'HOTLINE', 'VAPID_PUBLIC_KEY', 'VAPID_PRIVATE_KEY']) delete process.env[k];
    const app = require(path.join(ROOT, SERVER_FILE));
    const T = app._test || {};
    ok(typeof T.checkOverdueReminders === 'function' && typeof T.pushNotification === 'function', '[0] server.js 的 _test 掛鉤有 checkOverdueReminders／pushNotification');

    console.log('[1] pushNotification：有傳 callerData＝只改該物件（不存檔）；沒傳＝原行為（load＋save）');
    {
      const data = JSON.parse(fs.readFileSync(DATA_FILE, 'utf8'));
      const before = readData().notifications.length;
      const n = T.pushNotification('_s_a', 'unit_a', '標題A', '內文A', 'ref-a', data) || {};
      ok(n && n.id && n.to === '_s_a' && n.type === 'unit_a' && n.read === false, '1.1 回傳新增的通知物件');
      eq(data.notifications[0].id, n.id, '1.2 通知被加進呼叫端傳入的 data（最前面）');
      eq(readData().notifications.length, before, '1.3 有傳 data 時完全沒有寫檔');
      const n2 = T.pushNotification('_s_a', 'unit_b', '標題B', '內文B', 'ref-b');
      const disk = readData().notifications;
      ok(disk[0].id === n2.id && disk.length === before + 1, '1.4 沒傳 data＝原行為：load＋save，通知已寫進檔案');
      ok(!disk.some(x => x.id === n.id), '1.5 先前「只改物件」的那則沒有被存進檔案（呼叫端沒 save）');
      // 500 筆上限在兩種模式都生效
      const big = { notifications: Array.from({ length: 500 }, (_, i) => ({ id: 'x' + i })) };
      T.pushNotification('_s_a', 'cap', 't', 'b', null, big);
      eq([big.notifications.length, big.notifications[0].type], [500, 'cap'], '1.6 傳入 data 時 500 筆上限照樣生效（最舊的被丟掉）');
    }

    console.log('[2] checkOverdueReminders：兩筆以上逾期項目 → 每一筆通知都存在於最終資料檔，提醒標記也已更新');
    seed();   // 重置資料檔
    await T.checkOverdueReminders();
    {
      const d = readData();
      const notes = d.notifications || [];
      const by = (type) => notes.filter(n => n.type === type);
      const pairs = (type) => by(type).map(n => `${n.to}:${n.refId}`).sort();
      eq(pairs('contract_expiring_7'), ['_s_a:k1', '_s_b:k2', '_s_m1:k1', '_s_m1:k2'], '2.1 兩份 D-7 合約：擁有者＋一級主管各一則（共 4 則）都在最終資料檔');
      eq(pairs('contract_expiring_30'), ['_s_a:k3', '_s_m1:k3'], '2.2 D-30 預警：擁有者＋一級主管（共 2 則）');
      eq(pairs('birthday'), ['_s_a:c-bd'], '2.3 生日提醒 1 則');
      eq(pairs('receivable_overdue'), ['_s_a:r1', '_s_b:r2', '_s_m1:r1', '_s_m1:r2'], '2.4 兩筆逾期帳款：擁有者＋一級主管各一則（共 4 則）');
      ok(notes.some(n => n.id === 'old-1'), '2.5 既有通知沒有被洗掉');
      eq(notes.length, 1 + 4 + 2 + 1 + 4, '2.6 通知總數＝既有 1＋新增 11');
      ok(notes.filter(n => n.id !== 'old-1').every(n => n.read === false && n.createdAt && n.title && n.body), '2.7 新通知欄位完整（未讀、時間、標題、內文）');
      const k = (id) => d.contracts.find(c => c.id === id);
      const t0 = new Date(); t0.setHours(0, 0, 0, 0); const ts = t0.toISOString().slice(0, 10);   // 與 server.js 同式（本機午夜轉 UTC 日期；台灣時區會是「前一天」，既有行為），不寫死時區
      eq([k('k1')._contractReminderDate, k('k2')._contractReminderDate], [ts, ts], '2.8 D-7 合約的提醒標記已更新（_contractReminderDate）');
      eq(typeof k('k3')._contract30Reminded, 'string', '2.9 D-30 合約的預警標記已更新（_contract30Reminded）');
      eq(d.contacts[0]._birthdayReminderYear, today.getFullYear(), '2.10 生日提醒標記已更新');
      eq(d.receivables.map(r => r._receivableReminderDate), [ts, ts], '2.11 兩筆帳款的提醒標記已更新');
      ok(by('contract_expiring_7').concat(by('receivable_overdue')).every(n => /天/.test(n.body)), '2.12 通知內文有帶天數');
    }

    console.log('[3] 冪等：標記有存進去，所以再跑一次不會重複推送');
    {
      const before = readData().notifications.length;
      await T.checkOverdueReminders();
      eq(readData().notifications.length, before, '3.1 第二次執行沒有新增任何通知');
    }

    console.log('[4] 沒有任何到期項目：不寫檔（沒有不必要的 save）');
    {
      const raw = fs.readFileSync(DATA_FILE, 'utf8');
      const d = JSON.parse(raw);
      d.contracts = []; d.receivables = []; d.contacts = [];
      fs.writeFileSync(DATA_FILE, JSON.stringify(d, null, 2));
      const raw2 = fs.readFileSync(DATA_FILE, 'utf8');
      await T.checkOverdueReminders();
      eq(fs.readFileSync(DATA_FILE, 'utf8'), raw2, '4.1 沒有要提醒的項目時資料檔逐位元不變');
    }
  } catch (e) { console.log('錯誤：', e.stack || e.message); fail++; failures.push('例外：' + e.message); code = 1; }
  finally { cleanup(); }
  console.log(`\n逾期提醒通知檢查：通過 ${pass}　失敗 ${fail}`);
  if (fail) { console.log(failures.map(f => ' - ' + f).join('\n')); code = 1; }
  process.exit(code);
})();
