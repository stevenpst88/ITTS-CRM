#!/usr/bin/env node
'use strict';
/**
 * 信件預覽工具：用「合成假資料」把 E1–E6 × 每種收件人 kind 渲染成 HTML／純文字檔，供人工檢視。
 * 用法：node scripts/mail-preview.js [--out <資料夾>]
 *   預設輸出到 <repo>/.mail-preview/gallery/（已在 .gitignore）；--out 只允許指向 .mail-preview 之內。
 *
 * 輸出：
 *   index.html                 目錄頁（每封信的主旨、收件人類型、連結；不允許的組合標示 KIND_NOT_ALLOWED）
 *   <id>.html / <id>.txt       信件本體（HTML 與純文字版）
 *   dark/<id>.html             深色模式模擬版：把 @media (prefers-color-scheme:dark) 改成 @media all，
 *                              讓任何瀏覽器都能直接看到深色樣式（僅供預覽，不是寄出的版本）
 *   另外有 S-* 情境信：不同核決層級色塊、虧損、超長文字、35 個品項、惡意字串（確認有被跳脫）等。
 *
 * 安全：不連網、不寄信、不讀環境變數與 data.json／auth.json；資料全是通用字串（沒有真實姓名、客戶、email）。
 * 連結一律指向 https://crm.example.test，不會點到正式站。
 * 只會刪除輸出資料夾裡自己產生的 .html／.txt 檔（不遞迴刪除目錄）。
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const { getMailConfig } = require(path.join(ROOT, 'lib', 'mail', 'config.js'));
const { renderMail, MailRenderError, ALLOWED_KINDS } = require(path.join(ROOT, 'lib', 'mail', 'render.js'));
const { EVENT_TYPES } = require(path.join(ROOT, 'lib', 'mail', 'events.js'));
const { KINDS, visibilityFor } = require(path.join(ROOT, 'lib', 'mail', 'visibility.js'));
const { escHtml } = require(path.join(ROOT, 'lib', 'mail', 'safety.js'));

const PREVIEW_ROOT = path.join(ROOT, '.mail-preview');
const DEFAULT_OUT = path.join(PREVIEW_ROOT, 'gallery');

const CONFIG = getMailConfig({ APP_BASE_URL: 'https://crm.example.test' });
const AT = '2026-10-08T06:30:00.000Z';
const QID = '00000000-0000-4000-8000-0000000000a1';

const KIND_LABEL = {
  mgr1: '簽核人甲（一級主管）', gm: '簽核人乙（總經理）', chairman: '簽核人丙（董事長）',
  secretary: '秘書丁', boardProxy: '代核人戊', consultant: '顧問乙', owner: '業務甲',
};
const STEP_FOR_KIND = {
  mgr1: { level: 1, label: '一級主管' }, gm: { level: 2, label: '總經理' }, chairman: { level: 3, label: '董事長' },
  secretary: { level: 'board', label: '董事會' }, boardProxy: { level: 'board', label: '董事會' },
};

function numbers(over) {
  return Object.assign({ revenueCents: 123456789, gpCents: 50671000, marginText: '41.04%', marginPct: 41.04, tierLevel: 1, tierLabel: '一級主管' }, over || {});
}
function baseEvent(type, over) {
  return Object.assign({
    type,
    quoteId: QID,
    quoteNo: 'QU-000000-001',
    projectName: '示範專案：機房升級與網路整合案',
    company: '範例客戶股份有限公司',
    ownerLabel: '業務甲',
    at: AT,
    stepKey: AT + '#' + type,
  }, over || {});
}
const ITEMS = [
  { desc: '伺服器主機安裝與設定', qty: 2, unit: '台' },
  { desc: '網路交換器設備上架', qty: 4, unit: '台' },
  { desc: '系統導入顧問服務', qty: 10, unit: '人天' },
  { desc: '教育訓練', qty: 1, unit: '式' },
  { desc: '一年期原廠維護', qty: 1, unit: '年' },
];

/** 事件 × kind 的標準範例 */
function standardEvent(type, kind) {
  const step = STEP_FOR_KIND[kind];
  const tier = kind === 'gm' ? { tierLevel: 2, tierLabel: '總經理' } : kind === 'chairman' ? { tierLevel: 3, tierLabel: '董事長' } : kind === 'secretary' || kind === 'boardProxy' ? { tierLevel: 3, tierLabel: '董事會' } : {};
  switch (type) {
    case 'E1_SUBMIT':
    case 'E3_NEXT_STEP':
      return baseEvent(type, { step: step || { level: 1, label: '一級主管' }, numbers: numbers(tier), actor: { label: type === 'E1_SUBMIT' ? '業務甲' : '簽核人甲' } });
    case 'E2_COST_REQUEST':
      return baseEvent(type, { items: ITEMS, actor: { label: '業務甲' }, numbers: numbers() });  // 故意帶 numbers：顧問信不得顯示
    case 'E4_RESULT':
      return baseEvent(type, { step: { level: 1, label: '一級主管' }, result: { kind: 'rejected', reason: '毛利偏低，請重新評估報價與成本後再送簽。' }, actor: { label: '簽核人甲' }, numbers: numbers() });
    case 'E5_COST_DONE':
      return baseEvent(type, { actor: { label: '顧問乙' }, numbers: numbers() });
    case 'E6_WITHDRAWN':
      return baseEvent(type, { step: step || { level: 1, label: '一級主管' }, result: { kind: 'withdrawn', reason: '客戶需求變更，業務撤回後重新報價。' }, actor: { label: '業務甲' }, numbers: numbers(tier) });
    default:
      return null;
  }
}

function scenarios() {
  const long = '超長專案名稱' + '示範'.repeat(60);
  const xss = '<script>alert(1)</script>"><img src=x onerror=alert(2)> javascript:alert(3)';
  return [
    { id: 'S-tier2-amber', title: '核決層級 2（琥珀）', ev: standardEvent('E3_NEXT_STEP', 'gm'), kind: 'gm' },
    { id: 'S-tier3-red', title: '核決層級 3（紅）', ev: standardEvent('E3_NEXT_STEP', 'chairman'), kind: 'chairman' },
    { id: 'S-tier-null-grey', title: '核決層級未知（灰）', ev: baseEvent('E1_SUBMIT', { step: { level: 'board', label: '董事會' }, numbers: numbers({ tierLevel: null, tierLabel: '董事會' }) }), kind: 'secretary' },
    { id: 'S-loss', title: '虧損單（毛利為負）', ev: baseEvent('E1_SUBMIT', { step: { level: 1, label: '一級主管' }, numbers: numbers({ gpCents: -1234500, marginText: '-3.21%', marginPct: -3.21, tierLevel: 3, tierLabel: '董事長' }) }), kind: 'mgr1' },
    { id: 'S-huge-amount', title: '超大金額＋超長文字', ev: baseEvent('E1_SUBMIT', { projectName: long, company: long, step: { level: 1, label: '一級主管' }, numbers: numbers({ revenueCents: 1000000000000, gpCents: 410400000000 }) }), kind: 'mgr1' },
    { id: 'S-cents', title: '含分的金額（99 分、1 分）', ev: baseEvent('E1_SUBMIT', { step: { level: 1, label: '一級主管' }, numbers: numbers({ revenueCents: 99, gpCents: 1, marginText: '1.01%', marginPct: 1.01 }) }), kind: 'mgr1' },
    { id: 'S-items-35', title: '顧問信 35 個品項（顯示 30 列＋另 5 項）', ev: baseEvent('E2_COST_REQUEST', { items: Array.from({ length: 35 }, (_, i) => ({ desc: '品項說明 ' + (i + 1), qty: i + 1, unit: '式' })) }), kind: 'consultant' },
    { id: 'S-final-approved', title: '最終核准（綠）', ev: baseEvent('E4_RESULT', { result: { kind: 'final_approved' }, actor: { label: '簽核人丙' } }), kind: 'owner' },
    { id: 'S-returned', title: '退回修改（琥珀）', ev: baseEvent('E4_RESULT', { step: { level: 2, label: '總經理' }, result: { kind: 'returned', reason: '請補充競爭對手報價與付款條件。' } }), kind: 'owner' },
    { id: 'S-voided', title: '核准後修改作廢（紅）', ev: baseEvent('E6_WITHDRAWN', { step: { level: 2, label: '總經理' }, result: { kind: 'voided', reason: '核准後品項被修改。' }, actor: { label: '業務甲' } }), kind: 'gm' },
    { id: 'S-xss', title: '惡意字串（應全部被跳脫成純文字）', ev: baseEvent('E1_SUBMIT', { projectName: xss, company: xss, ownerLabel: xss, step: { level: 1, label: '一級主管' }, numbers: numbers(), actor: { label: xss } }), kind: 'mgr1', label: xss },
  ];
}

function toDark(html) {
  return html.replace('@media (prefers-color-scheme:dark){', '@media all{');
}

function parseArgs(argv) {
  const a = { out: DEFAULT_OUT };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === '--out' && argv[i + 1]) a.out = path.resolve(argv[++i]);
  }
  return a;
}

function clearGenerated(dir) {
  if (!fs.existsSync(dir)) return;
  fs.readdirSync(dir).forEach((n) => {
    const p = path.join(dir, n);
    if (fs.statSync(p).isFile() && /\.(html|txt)$/.test(n)) fs.unlinkSync(p);
  });
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  const rel = path.relative(PREVIEW_ROOT, args.out);
  if (rel.startsWith('..') || path.isAbsolute(rel)) {
    console.error('--out 必須在 ' + PREVIEW_ROOT + ' 之內（這個資料夾在 .gitignore，不會被提交）');
    return 1;
  }
  fs.mkdirSync(path.join(args.out, 'dark'), { recursive: true });
  clearGenerated(args.out);
  clearGenerated(path.join(args.out, 'dark'));

  const entries = [];
  const write = (id, mail) => {
    fs.writeFileSync(path.join(args.out, id + '.html'), mail.html, 'utf8');
    fs.writeFileSync(path.join(args.out, id + '.txt'), mail.text, 'utf8');
    fs.writeFileSync(path.join(args.out, 'dark', id + '.html'), toDark(mail.html), 'utf8');
  };

  EVENT_TYPES.forEach((type) => {
    KINDS.forEach((kind) => {
      const id = type.split('_')[0] + '_' + kind;
      const ev = standardEvent(type, kind);
      try {
        const mail = renderMail(ev, { username: 'user-' + kind, label: KIND_LABEL[kind], kind }, { config: CONFIG });
        write(id, mail);
        entries.push({ id, type, kind, ok: true, subject: mail.subject, hasAmount: mail.meta.hasAmount });
      } catch (e) {
        if (!(e instanceof MailRenderError)) throw e;
        entries.push({ id, type, kind, ok: false, code: e.code });
      }
    });
  });
  const extra = [];
  scenarios().forEach((s) => {
    const mail = renderMail(s.ev, { username: 'user-' + s.kind, label: s.label || KIND_LABEL[s.kind], kind: s.kind }, { config: CONFIG });
    write(s.id, mail);
    extra.push({ id: s.id, title: s.title, type: s.ev.type, kind: s.kind, subject: mail.subject, hasAmount: mail.meta.hasAmount });
  });

  fs.writeFileSync(path.join(args.out, 'index.html'), indexHtml(entries, extra), 'utf8');
  const ok = entries.filter((e) => e.ok).length;
  console.log('已輸出 ' + (ok + extra.length) + ' 封信（標準 ' + ok + ' ＋ 情境 ' + extra.length + '）到 ' + args.out);
  console.log('不寄送的事件×收件人組合：' + entries.filter((e) => !e.ok).length + ' 種（KIND_NOT_ALLOWED）');
  return 0;
}

function indexHtml(entries, extra) {
  const e = escHtml;
  const vis = (k) => { const v = visibilityFor(k); return ['amount', 'margin', 'customer', 'owner', 'project', 'items'].filter((f) => v[f]).join('、') || '(僅單號)'; };
  const rows = entries.map((x) => x.ok
    ? '<tr><td>' + e(x.type) + '</td><td>' + e(x.kind) + '</td><td>' + e(x.subject) + '</td><td>' + e(vis(x.kind)) + '</td><td><a href="' + e(x.id) + '.html">HTML</a> · <a href="dark/' + e(x.id) + '.html">深色</a> · <a href="' + e(x.id) + '.txt">純文字</a></td></tr>'
    : '<tr class="na"><td>' + e(x.type) + '</td><td>' + e(x.kind) + '</td><td colspan="3">不寄送（' + e(x.code) + '）</td></tr>').join('\n');
  const rows2 = extra.map((x) => '<tr><td>' + e(x.id) + '</td><td>' + e(x.kind) + '</td><td>' + e(x.title) + '<br><small>' + e(x.subject) + '</small></td><td>' + (x.hasAmount ? '含金額' : '無金額') + '</td><td><a href="' + e(x.id) + '.html">HTML</a> · <a href="dark/' + e(x.id) + '.html">深色</a> · <a href="' + e(x.id) + '.txt">純文字</a></td></tr>').join('\n');
  return '<!DOCTYPE html>\n<html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>信件預覽</title>\n' +
    '<style>body{font-family:"Microsoft JhengHei",Arial,sans-serif;margin:24px;color:#111827}table{border-collapse:collapse;width:100%}td,th{border:1px solid #d1d5db;padding:6px 10px;text-align:left;font-size:14px;vertical-align:top}th{background:#f3f4f6}tr.na td{color:#6b7280;background:#f9fafb}small{color:#4b5563}</style></head><body>\n' +
    '<h1>信件預覽（合成資料，不連網、不寄信）</h1>\n' +
    '<p>連結一律指向 https://crm.example.test。「深色」是把深色模式媒體查詢強制啟用的模擬版，僅供預覽。</p>\n' +
    '<h2>事件 × 收件人類型</h2>\n<table><tr><th>事件</th><th>收件人類型</th><th>主旨</th><th>信中可見的欄位</th><th>檢視</th></tr>\n' + rows + '\n</table>\n' +
    '<h2>情境信</h2>\n<table><tr><th>編號</th><th>收件人類型</th><th>說明／主旨</th><th>金額</th><th>檢視</th></tr>\n' + rows2 + '\n</table>\n' +
    '<h2>允許矩陣</h2><pre>' + e(JSON.stringify(ALLOWED_KINDS, null, 2)) + '</pre>\n</body></html>\n';
}

process.exit(main());
