// SEC-05 匯入檔安全閘門檢查（lib/xlsxSafe.js ＋ server.js 七個匯入口的 xlsxGuard）。
//   node scripts/check-xlsx-import-guard.js          單元＋HTTP（沙盒伺服器，埠 XG_PORT 預設 3405，臨時 _s_* 帳號與假資料，跑完刪除）
//   XG_MUT=1 node scripts/check-xlsx-import-guard.js 變異：拿掉各項防護，對應測試必須變紅（共 8 個變異）
// 單元層不碰網路與檔案；HTTP 層沿用 _b2-harness（工作樹若已有 data.json／auth.json 一律拒絕執行）。
process.env.B2_PORT = process.env.XG_PORT || '3405';
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..');
const JSZip = require(path.join(ROOT, 'node_modules', 'jszip'));
const XLSX = require(path.join(ROOT, 'node_modules', 'xlsx'));
const H = require('./_b2-harness');

const MUT = process.env.XG_MUT === '1';
let R = { pass: 0, fail: 0, failures: [] };
const ok = (c, n) => { if (c) R.pass++; else { R.fail++; R.failures.push(n); if (!MUT) console.log('  ✗ FAIL:', n); } };

// ───────── 樣本 ─────────
const rows = [['姓名', '公司', '日期', '備註'], ['王小明', '甲公司', 46000, ''], ['李小華', '乙公司', 46001, '中文備註']];
const mkWb = () => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), '資料'); return wb; };
const XLSX_BUF = XLSX.write(mkWb(), { type: 'buffer', bookType: 'xlsx' });
const XLS_BUF = XLSX.write(mkWb(), { type: 'buffer', bookType: 'biff8' });
const CSV_BUF = Buffer.from('﻿姓名,公司,日期,備註\n王小明,甲公司,2026/1/1,\n', 'utf8');
const HTML_XLS = Buffer.from('<html><body><table><tr><td>姓名</td><td>公司</td></tr><tr><td>王</td><td>甲</td></tr></table></body></html>', 'utf8');
const UTF16_CSV = Buffer.concat([Buffer.from([0xFF, 0xFE]), Buffer.from('姓名\t公司\n王\t甲\n', 'utf16le')]);

async function zipWith(mutate) {
  const z = await JSZip.loadAsync(XLSX_BUF);
  await mutate(z);
  return z.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
const protoSheet = (fn) => zipWith(async (z) => { const s = await z.file('xl/workbook.xml').async('string'); z.file('xl/workbook.xml', fn(s)); });
/** 在每個 zip 中央目錄項目上改位元組（不改長度，所以位移仍有效） */
const patchCD = (buf, fn) => { const b = Buffer.from(buf); for (let i = 0; i + 4 <= b.length; i++) if (b.readUInt32LE(i) === 0x02014b50) fn(b, i); return b; };
const patchEocd = (buf, fn) => { const b = Buffer.from(buf); for (let i = b.length - 22; i >= 0; i--) if (b.readUInt32LE(i) === 0x06054b50) { fn(b, i); break; } return b; };
/** 同長度改名（本地檔頭與中央目錄兩處都改）→ 做出「同名條目」 */
const renameBytes = (buf, from, to) => { const b = Buffer.from(buf), f = Buffer.from(from), t = Buffer.from(to); let at = b.indexOf(f); while (at >= 0) { t.copy(b, at); at = b.indexOf(f, at + 1); } return b; };
const readSheet = (b) => { const wb = XLSX.read(b, { type: 'buffer' }); return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' }); };

async function samples() {
  return {
    bomb: await zipWith((z) => { z.file('xl/worksheets/sheet2.xml', Buffer.alloc(70 * 1024 * 1024)); }),             // 工作表部件解壓後 70MB（單項上限 32MB）
    cpuAttack: await zipWith((z) => { z.file('xl/worksheets/sheet2.xml', Buffer.alloc(63 * 1024 * 1024)); z.file('xl/worksheets/sheet3.xml', Buffer.alloc(63 * 1024 * 1024)); }),   // 驗收報告的 CPU 放大樣本（約 130KB）
    oneBig: await zipWith((z) => { z.file('xl/worksheets/sheet2.xml', Buffer.alloc(40 * 1024 * 1024)); }),                // 單一 40MB 工作表：只有單項上限（32MB）會擋（總量 64MB 以內）
    mediaBig: await zipWith((z) => { z.file('xl/media/big.bin', Buffer.alloc(70 * 1024 * 1024)); }),                  // 白名單以外的大部件：不解壓、不保留
    nearLimit: await zipWith((z) => { z.file('xl/worksheets/sheet2.xml', Buffer.alloc(30 * 1024 * 1024)); z.file('xl/worksheets/sheet3.xml', Buffer.alloc(30 * 1024 * 1024)); }),   // 合計 60MB：在上限內（最壞可接受情況，量測耗時用）
    protoSingle: await protoSheet((s) => s.replace(/name="[^"]*"/, "name='__proto__'")),
    protoEntity: await protoSheet((s) => s.replace(/name="[^"]*"/, 'name="&#95;&#95;proto&#95;&#95;"')),
    protoHex: await protoSheet((s) => s.replace(/name="[^"]*"/, 'name="&#x5f;_proto__"')),
    protoDouble: await protoSheet((s) => s.replace(/name="[^"]*"/, 'name="&amp;#95;_proto__"')),
    protoCtorEntity: await protoSheet((s) => s.replace(/name="[^"]*"/, 'name="constr&#117;ctor"')),
    protoNewline: await protoSheet((s) => s.replace(/<sheet /, '<sheet\n').replace(/name="[^"]*"/, 'name\n=\n"__proto__"')),
    protoUtf16: await zipWith(async (z) => { const s = await z.file('xl/workbook.xml').async('string'); z.file('xl/workbook.xml', Buffer.concat([Buffer.from([0xFF, 0xFE]), Buffer.from(s.replace(/name="[^"]*"/, 'name="__proto__"'), 'utf16le')])); }),
    doctypeUtf16: await zipWith(async (z) => { const s = await z.file('xl/worksheets/sheet1.xml').async('string'); z.file('xl/worksheets/sheet1.xml', Buffer.concat([Buffer.from([0xFF, 0xFE]), Buffer.from('<!DOCTYPE x [<!ENTITY a "b">]>' + s, 'utf16le')])); }),
    dupNames: await (async () => { const b = await zipWith((z) => { z.file('xl/worksheets/sheet2.xml', '<worksheet/>'); }); return renameBytes(b, 'sheet2.xml', 'sheet1.xml'); })(),
    encrypted: patchCD(XLSX_BUF, (b, i) => b.writeUInt16LE(b.readUInt16LE(i + 8) | 1, i + 8)),
    sizeMismatch: patchCD(XLSX_BUF, (b, i) => b.writeUInt32LE(b.readUInt32LE(i + 24) + 1, i + 24)),
    zip64: patchEocd(XLSX_BUF, (b, e) => b.writeUInt16LE(0xFFFF, e + 10)),
    manyEntries: await zipWith((z) => { for (let i = 0; i < 2001; i++) z.file('xl/media/m' + i + '.bin', 'x'); }),
    fewEntries: await zipWith((z) => { for (let i = 0; i < 100; i++) z.file('xl/media/m' + i + '.bin', 'x'); }),
    backslash: await zipWith((z) => { z.file('xl\\evil.xml', '<a/>'); }),
    bigLegit: (() => { const aoa = [['姓名', '公司', '職稱', '電話', '分機', '手機', 'Email', '地址', '網站', '統一編號', '產業屬性', '備註']]; for (let i = 0; i < 10000; i++) aoa.push(['客戶' + i, '公司' + (i % 900), '經理', '02-1234-' + i, '', '09' + String(10000000 + i), 'u' + i + '@example.test', '台北市某路' + i + '號', 'www.example.test', '', '製造業', '備註文字 ' + i]); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), '資料'); return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', bookSST: true }); })(),   // 合法大檔：1 萬列×12 欄
    extraPart: await zipWith((z) => { z.file('xl/media/x.bin', 'MZ-extra'); z.file('docProps/evil.bin', 'extra'); }),
    doctype: await zipWith(async (z) => { const s = await z.file('xl/worksheets/sheet1.xml').async('string'); z.file('xl/worksheets/sheet1.xml', '<!DOCTYPE x [<!ENTITY a "b">]>' + s); }),
    proto: await zipWith(async (z) => { const s = await z.file('xl/workbook.xml').async('string'); z.file('xl/workbook.xml', s.replace(/name="[^"]*"/, 'name="__proto__"')); }),
    traversal: await zipWith((z) => { z.file('../evil.xml', '<a/>'); }),
    docx: await (async () => { const z = new JSZip(); z.file('[Content_Types].xml', '<Types/>'); z.file('word/document.xml', '<w/>'); return z.generateAsync({ type: 'nodebuffer' }); })(),
    emptyZip: Buffer.from([0x50, 0x4B, 0x05, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
    truncated: XLSX_BUF.slice(0, Math.floor(XLSX_BUF.length / 2)),
    garbage: Buffer.concat([Buffer.from('MZ'), Buffer.alloc(200, 0), Buffer.from('This program cannot be run in DOS mode')]),
    ratio: await zipWith((z) => { for (let i = 0; i < 3; i++) z.file('xl/worksheets/s' + i + '.xml', Buffer.alloc(30 * 1024 * 1024)); }),   // 三個各 30MB（各自低於單項上限）合計 90MB > 總量上限 64MB
  };
}

// ───────── 單元層 ─────────
async function unit(X, S) {
  let r = await X.prepareUpload(XLSX_BUF);
  ok(r.ok && r.kind === 'xlsx' && JSON.stringify(readSheet(r.buf)) === JSON.stringify(readSheet(XLSX_BUF)), 'U1 正常 xlsx：通過，重新封裝後解析結果與原檔逐格相同');
  r = await X.prepareUpload(XLS_BUF);
  ok(r.ok && r.kind === 'xls' && r.buf === XLS_BUF, 'U2 正常 .xls（CFB）：原樣放行（同一個 Buffer）');
  r = await X.prepareUpload(CSV_BUF);
  ok(r.ok && r.kind === 'text' && r.buf === CSV_BUF, 'U3 CSV（含 BOM、中文）：原樣放行');
  r = await X.prepareUpload(HTML_XLS);
  ok(r.ok && r.kind === 'text' && r.buf === HTML_XLS, 'U4 舊系統 HTML 偽裝 .xls：原樣放行（行為不變，不誤擋）');
  r = await X.prepareUpload(UTF16_CSV);
  ok(r.ok && r.buf === UTF16_CSV, 'U5 UTF-16 文字檔（有 BOM）：原樣放行');
  r = await X.prepareUpload(Buffer.alloc(0));
  ok(!r.ok && r.code === 'EMPTY_FILE', 'U6 空檔 → EMPTY_FILE');
  r = await X.prepareUpload(S.bomb);
  ok(!r.ok && r.code === 'BAD_XLSX', 'U7 zip bomb（單項解壓 70MB）→ 拒絕');
  r = await X.prepareUpload(S.ratio);
  ok(!r.ok && r.code === 'BAD_XLSX', 'U8 多項合計超過總量上限 → 拒絕');
  r = await X.prepareUpload(S.doctype);
  ok(!r.ok && r.code === 'BAD_XLSX', 'U9 sheet XML 含 DOCTYPE／ENTITY → 拒絕');
  r = await X.prepareUpload(S.proto);
  ok(!r.ok && r.code === 'BAD_XLSX', 'U10 工作表名稱 __proto__ → 拒絕');
  r = await X.prepareUpload(S.traversal);
  ok(!r.ok && r.code === 'BAD_XLSX', 'U11 路徑跳脫（../）→ 拒絕');
  r = await X.prepareUpload(S.docx);
  ok(!r.ok && r.code === 'BAD_XLSX', 'U12 非 xlsx 的 zip（docx 類）→ 拒絕');
  r = await X.prepareUpload(S.emptyZip);
  ok(!r.ok, 'U13 空 zip → 拒絕');
  r = await X.prepareUpload(S.truncated);
  ok(!r.ok, 'U14 截斷的 xlsx → 拒絕');
  r = await X.prepareUpload(S.garbage);
  ok(!r.ok && r.code === 'BAD_UPLOAD', 'U15 二進位垃圾（改副檔名的執行檔）→ 拒絕');
  // 中介層
  const mk = () => { const o = { code: 0, body: null }; return { o, res: { status(c) { o.code = c; return this; }, json(b) { o.body = b; return this; } } }; };
  let m = mk(), called = 0;
  await X.uploadGuard()({ file: { buffer: XLSX_BUF, size: XLSX_BUF.length } }, m.res, () => called++);
  ok(called === 1 && m.o.code === 0, 'U16 中介層：合格檔放行');
  m = mk(); called = 0;
  const reqBad = { file: { buffer: S.bomb, size: S.bomb.length } };
  await X.uploadGuard()(reqBad, m.res, () => called++);
  ok(called === 0 && m.o.code === 400 && m.o.body && m.o.body.code === 'BAD_XLSX', 'U17 中介層：不合格 → 400＋code，不進路由');
  m = mk(); called = 0;
  await X.uploadGuard()({}, m.res, () => called++);
  ok(called === 1, 'U18 中介層：沒有檔案 → 放行（路由自己回「請上傳」）');
  // 牌價簿仍沿用原限制（單項 6MB）
  const PBX = require(path.join(ROOT, 'lib', 'quotePricebookXlsx'));
  const big7 = await zipWith((z) => { z.file('xl/media/b.bin', Buffer.alloc(7 * 1024 * 1024)); });
  r = await X.prepareUpload(big7);
  const pr = await PBX.safeRepack(big7);
  ok(r.ok && !pr.ok, 'U19 同一個 7MB 部件：一般匯入通過、牌價簿仍拒絕（牌價簿限制未放寬）');
  await unit2(X, S);
}

// 第二批：閘門內部各項檢查（每一項都有對應變異，拿掉就會紅）
async function unit2(X, S) {
  const rej = async (b) => { const r = await X.prepareUpload(b); return !r.ok && r.code === 'BAD_XLSX'; };
  ok(await rej(S.protoSingle), 'U20 工作表名 name=\'__proto__\'（單引號）→ 拒絕');
  ok(await rej(S.protoEntity), 'U21 工作表名以 &#95; 實體拼出 __proto__ → 拒絕');
  ok(await rej(S.protoHex), 'U22 工作表名以 &#x5f; 十六進位實體 → 拒絕');
  ok(await rej(S.protoDouble), 'U23 工作表名疊加編碼 &amp;#95; → 拒絕');
  ok(await rej(S.protoCtorEntity), 'U24 工作表名 constr&#117;ctor → 拒絕');
  ok(await rej(S.protoNewline), 'U25 name 屬性前後換行 → 拒絕');
  ok(await rej(S.protoUtf16), 'U26 workbook.xml 為 UTF-16 且工作表名 __proto__ → 拒絕');
  ok(await rej(S.doctypeUtf16), 'U27 sheet XML 為 UTF-16 且含 DOCTYPE／ENTITY → 拒絕');
  ok(!X.hasBadSheetName('<sheet name="A&amp;B"/><sheet name="資料" sheetId="2"/><sheets/>'), 'U28 正常工作表名（含 &amp;、中文）不誤判');
  ok(await rej(S.dupNames), 'U29 同名條目（兩個 sheet1.xml）→ 拒絕');
  ok(await rej(S.encrypted), 'U30 加密旗標 → 拒絕');
  ok(await rej(S.sizeMismatch), 'U31 解壓後大小與宣告不符 → 拒絕');
  ok(await rej(S.zip64), 'U32 ZIP64 標記 → 拒絕');
  ok(await rej(S.manyEntries), 'U33 項目數 2001（上限 2000）→ 拒絕');
  const fe = await X.prepareUpload(S.fewEntries);
  ok(fe.ok, 'U34 項目數 100 的正常檔通過（上一項不是一律拒絕）');
  ok(await rej(S.backslash), 'U35 項目名稱含反斜線 → 拒絕');
  // 白名單與重新封裝
  const mb = await X.prepareUpload(S.mediaBig);
  ok(mb.ok && !(await JSZip.loadAsync(mb.buf)).file('xl/media/big.bin'), 'U36 白名單以外的 70MB 部件：通過、但不被解壓也不保留');
  let t0 = Date.now();
  ok(await rej(S.oneBig), 'U47 單一 40MB 工作表部件（超過單項上限、未超過總量上限）→ 拒絕');
  ok(await rej(S.cpuAttack), 'U37 CPU 放大樣本（2×63MB 全零工作表）→ 拒絕');
  const tAttack = Date.now() - t0;
  ok(tAttack < 1500, 'U38 上項拒絕耗時 < 1.5 秒（實際 ' + tAttack + 'ms；修正前單路由 3.5–5.4 秒）');
  t0 = Date.now();
  const nl = await X.prepareUpload(S.nearLimit);
  const tNear = Date.now() - t0;
  ok(nl.ok && tNear < 2500, 'U39 上限內最壞樣本（合計 60MB）通過且 < 2.5 秒（實際 ' + tNear + 'ms）');
  // 中介層：重新封裝版取代原檔、解析失敗不放行
  const orig = S.extraPart;
  const rq = { file: { buffer: orig, size: orig.length } };
  let called = 0;
  await X.uploadGuard()(rq, { status() { return this; }, json() { return this; } }, () => called++);
  const z2 = rq.file.buffer === orig ? null : await JSZip.loadAsync(rq.file.buffer);
  ok(called === 1 && z2 && !z2.file('xl/media/x.bin') && !z2.file('docProps/evil.bin') && rq.file.size === rq.file.buffer.length, 'U40 中介層放行時 req.file.buffer 已換成重新封裝版（多餘部件被丟掉、size 同步）');
  let n = 0, code = 0, nx = 0;
  const rq2 = { file: Object.defineProperty({}, 'buffer', { get() { if (++n > 1) throw new Error('boom'); return Buffer.from('x'); } }) };
  await X.uploadGuard()(rq2, { status(c) { code = c; return this; }, json() { return this; } }, () => nx++);
  ok(code === 400 && nx === 0, 'U41 閘門內部例外 → 400，不 fail-open（不呼叫 next）');
  // 白名單重新封裝不可改變解析結果：多工作表、註解、超連結、合併、公式、日期、定義名稱的活頁簿，原檔與重新封裝版逐格相同
  const rich = (() => {
    const wb = XLSX.utils.book_new();
    const a = XLSX.utils.aoa_to_sheet([['姓名', '日期', '金額', '連結'], ['王', new Date(Date.UTC(2026, 9, 5)), 1234.5, 'x'], ['李', 46001, '=1+2', '']], { cellDates: true });
    a.A2.c = [{ a: 'x', t: '註解' }]; a.D2.l = { Target: 'https://example.test/' }; a['!merges'] = [{ s: { r: 2, c: 2 }, e: { r: 2, c: 3 } }];
    XLSX.utils.book_append_sheet(wb, a, '資料');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['b', 1], ['c', 2]]), '2026');
    wb.Workbook = { Names: [{ Name: 'MyName', Ref: '資料!$A$1' }] };
    return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', cellDates: true });
  })();
  const dump = (b) => { const w = XLSX.read(b, { type: 'buffer' }); return JSON.stringify(w.SheetNames.map((nm) => [nm, XLSX.utils.sheet_to_json(w.Sheets[nm], { header: 1, defval: '' }), w.Sheets[nm]['!merges'] || null])); };
  const rr = await X.prepareUpload(rich);
  ok(rr.ok && dump(rr.buf) === dump(rich), 'U45 多工作表＋註解＋超連結＋合併＋公式＋日期＋定義名稱：重新封裝前後解析結果逐格相同');
  const big = await X.prepareUpload(S.bigLegit);
  ok(big.ok && dump(big.buf) === dump(S.bigLegit), 'U46 合法大檔（1 萬列×12 欄）：通過且解析結果與原檔逐格相同');
  // 已知行為（SEC-05 實測，鎖定；若升級 SheetJS 後這兩條變了，要重新評估匯入結果是否改變）
  const mk = (bookSST) => { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['&lt;b&gt;', 'A&amp;B', 'A&B', '<x>']]), 'S'); return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', bookSST }); };
  const rd = (b) => { const w = XLSX.read(b, { type: 'buffer' }); return XLSX.utils.sheet_to_json(w.Sheets.S, { header: 1 })[0]; };
  eqj(rd(mk(false)), ['<b>', 'A&B', 'A&B', '<x>'], 'U42 已知行為：SheetJS 0.20.3 讀 t="str" 儲存格，字面實體（&lt;、&amp;）會多解一次（0.18.5 保持原樣；Excel 的共用字串不受影響）');
  eqj(rd(mk(true)), ['&lt;b&gt;', 'A&amp;B', 'A&B', '<x>'], 'U43 共用字串（Excel 慣用，t="s"）的字面實體原樣保留');
  const cw = XLSX.read(Buffer.from('"=SUM(1,2)",b\n"@cmd",c\n'), { type: 'buffer' });
  eqj(XLSX.utils.sheet_to_json(cw.Sheets[cw.SheetNames[0]], { header: 1, defval: '' }), [['=SUM(1,2)', 'b'], ['@cmd', 'c']], 'U44 已知行為：CSV 以 = 開頭的格，0.20.3 讀成原文（0.18.5 讀成空字串；公式不會被求值，匯入端當純文字）');
}
const eqj = (a, b, n) => ok(JSON.stringify(a) === JSON.stringify(b), n + '（得到 ' + JSON.stringify(a) + '）');

// ───────── HTTP 層 ─────────
const ROUTES = [
  ['/api/admin/contacts/import', '_s_adm'], ['/api/admin/opportunities/import', '_s_adm'], ['/api/admin/visits/import', '_s_adm'],
  ['/api/admin/contracts/import', '_s_adm'], ['/api/companies/import', '_s_mkt'], ['/api/yoy/import', '_s_adm'], ['/api/admin/import-contacts', '_s_adm'],
];
const cookies = {};   // 登入限流（15 分鐘 20 次）：每個帳號只登入一次
async function up(who, url, buf, name, extra) {
  if (!cookies[who]) { const lr = await H.http(null, 'POST', '/api/login', { username: who, password: H.PW }); cookies[who] = lr.setCookie.map((c) => c.split(';')[0]).join('; '); }
  const cookie = cookies[who];
  const fd = new FormData(); fd.append('file', new Blob([buf]), name); for (const [k, v] of Object.entries(extra || {})) fd.append(k, v);
  const res = await fetch(H.BASE + url + (url.includes('companies') || url.includes('yoy') ? '?dryRun=1' : ''), { method: 'POST', headers: { Cookie: cookie }, body: fd });
  return { status: res.status, json: await res.json().catch(() => ({})) };
}
const guardCodes = ['BAD_XLSX', 'BAD_UPLOAD', 'EMPTY_FILE'];
const rejected = (r) => r.status === 400 && guardCodes.includes(r.json && r.json.code);

async function http(S, serverFile) {
  for (const k of Object.keys(cookies)) delete cookies[k];
  await H.fresh([H.mkUser('_s_adm', 'admin'), H.mkUser('_s_mkt', 'marketing')], { companies: [], contacts: [], opportunities: [], industries: ['製造業'] }, { env: {}, file: serverFile });
  for (const [url, who] of ROUTES) {
    const ex = { defaultOwner: '_s_adm', targetOwner: '_s_adm' };
    let r = await up(who, url, S.bomb, 'a.xlsx', ex);
    ok(rejected(r), 'H ' + url + ' zip bomb → 400 BAD_XLSX（' + r.status + '）');
    r = await up(who, url, S.garbage, 'a.xlsx', ex);
    ok(rejected(r), 'H ' + url + ' 改副檔名的二進位 → 400 BAD_UPLOAD（' + r.status + '）');
    const tc = Date.now();
    r = await up(who, url, S.cpuAttack, 'a.xlsx', ex);
    ok(rejected(r) && Date.now() - tc < 3000, 'H ' + url + ' CPU 放大樣本 → 400 且 < 3 秒（實際 ' + (Date.now() - tc) + 'ms）');
    r = await up(who, url, S.doctype, 'a.xlsx', ex);
    ok(rejected(r), 'H ' + url + ' DOCTYPE → 400（' + r.status + '）');
    r = await up(who, url, XLSX_BUF, 'a.xlsx', ex);
    ok(!rejected(r) && r.status < 500, 'H ' + url + ' 正常 xlsx 不被閘門擋（' + r.status + '）');
    r = await up(who, url, XLS_BUF, 'a.xls', ex);
    ok(!rejected(r) && r.status < 500, 'H ' + url + ' 正常 xls 不被閘門擋（' + r.status + '）');
    r = await up(who, url, CSV_BUF, 'a.csv', ex);
    ok(!rejected(r) && r.status < 500, 'H ' + url + ' 正常 csv 不被閘門擋（' + r.status + '）');
  }
  // 合法大檔（1 萬列×12 欄）：閘門放行、企業主檔匯入預覽（dryRun）照常解析
  const tb = Date.now();
  const bl = await up('_s_mkt', '/api/companies/import', S.bigLegit, 'big.xlsx', {});
  ok(S.bigLegit.length < 10 * 1024 * 1024 && !rejected(bl) && bl.status === 200 && bl.json.totalRows === 10000, 'H 合法大檔 1 萬列（' + Math.round(S.bigLegit.length / 1024) + 'KB）企業主檔匯入預覽 200、totalRows=10000（status ' + bl.status + '、' + (Date.now() - tb) + 'ms）');
  const alive = await fetch(H.BASE + '/login.html').then((x) => x.ok).catch(() => false);
  ok(alive, 'H 惡意檔全部測完後伺服器仍存活');
}

// ───────── 變異 ─────────
function mutate(src, from, to, label) { if (!src.includes(from)) throw new Error('變異點找不到：' + label); return src.replace(from, to); }

async function main() {
  H.guardClean();
  const S = await samples();
  const srcPath = path.join(ROOT, 'lib', 'xlsxSafe.js');
  if (!MUT) {
    await unit(require(srcPath), S);
    try { await http(S); } finally { await H.finish(); }
    console.log(`\nSEC-05 匯入閘門檢查：${R.pass} 通過，${R.fail} 失敗`);
    if (R.fail) console.log(R.failures.map((f) => ' - ' + f).join('\n'));
    process.exit(R.fail ? 1 : 0);
  }
  // 變異模式
  const src = fs.readFileSync(srcPath, 'utf8');
  const muts = [
    ['M1 拿掉 DOCTYPE／ENTITY 檢查', (s) => mutate(s, "if (/<!DOCTYPE|<!ENTITY/i.test(txt)) return E;", '', 'm1'), /U9/],
    ['M2 拿掉 __proto__ 工作表名稱檢查', (s) => mutate(s, "&& hasBadSheetName(txt)", '&& false', 'm2'), /U10/],
    ['M2b 工作表名不解 XML 實體（舊版缺陷）', (s) => mutate(s, 'BAD_SHEET_NAME_RE.test(decodeEntities(a[1] !== undefined ? a[1] : a[2]).trim())', 'BAD_SHEET_NAME_RE.test((a[1] !== undefined ? a[1] : a[2]).trim())', 'm2b'), /U21/],
    ['M2c 工作表名只認雙引號、不處理換行（舊版缺陷）', (s) => mutate(s, "const a = /\\bname\\s*=\\s*(?:\"([^\"]*)\"|'([^']*)')/i.exec(m[0]);", 'const a = /name="([^"]*)"/i.exec(m[0]);', 'm2c'), /U20/],
    ['M2d UTF-16 部件當 UTF-8 掃（舊版缺陷）', (s) => mutate(s, 'const txt = xmlText(data);', "const txt = data.toString('utf8');", 'm2d'), /U26/],
    ['M10 拿掉同名條目檢查', (s) => mutate(s, 'seen.has(e.name) || BAD_NAME_RE.test(e.name)', 'BAD_NAME_RE.test(e.name)', 'm10'), /U29/],
    ['M11 拿掉加密旗標檢查', (s) => mutate(s, 'if (e.flags & 1) return null;', '', 'm11'), /U30/],
    ['M12 拿掉「解壓後大小須等於宣告」', (s) => mutate(s, 'if (out.length !== e.uncomp || out.length > maxEntryBytes) return null;', 'if (out.length > maxEntryBytes) return null;', 'm12'), /U31/],
    ['M13 項目數上限拿掉', (s) => mutate(s, 'if (total > maxEntries || cdOff + cdSize > eocd) return null;', 'if (cdOff + cdSize > eocd) return null;', 'm13'), /U33/],
    ['M14 uploadGuard 不換成重新封裝版', (s) => mutate(s, 'if (r.buf !== req.file.buffer) {', 'if (false) {', 'm14'), /U40/],
    ['M15 uploadGuard 例外時 fail-open', (s) => mutate(s, "    } catch (_) {\n      return res.status(400).json({ error: UPLOAD_LIMITS.badMessage, code: 'BAD_UPLOAD' });", '    } catch (_) {\n      return next();', 'm15'), /U41/],
    ['M16 拿掉白名單（所有部件都解壓並保留）', (s) => mutate(s, 'if (o.keepPart && !o.keepPart(e.name)) continue;', '', 'm16'), /U36/],
    ['M17 白名單丟部件但不修 [Content_Types].xml', (s) => mutate(s, 'dropOverrides(xmlText(data), o.keepPart)', 'xmlText(data)', 'm17'), /U1 |U45/],
    ['M18 拿掉 ZIP64 標記檢查（等價變異：項目數／偏移檢查已涵蓋，預期仍綠）', (s) => mutate(s, 'if (total === 0xFFFF || cdSize === 0xFFFFFFFF || cdOff === 0xFFFFFFFF) return null;   // ZIP64', '', 'm18'), null],
    ['M3 拿掉路徑跳脫檢查', (s) => mutate(s, 'seen.has(e.name) || BAD_NAME_RE.test(e.name)', 'seen.has(e.name)', 'm3'), /U11/],
    ['M4 單項解壓上限拉到天文數字', (s) => mutate(s, 'maxEntryBytes: 32 * 1024 * 1024,', 'maxEntryBytes: 4096 * 1024 * 1024,', 'm4'), /U47/],
    ['M5 拿掉總量上限檢查', (s) => mutate(s, 'if (total > o.maxTotalBytes)', 'if (false)', 'm5'), /U8/],
    ['M6 拿掉必要部件檢查（docx 類）', (s) => mutate(s, '!REQUIRED.every((n) => seen.has(n)) || !entries.some((e) => SHEET_PART_RE.test(e.name))', 'false', 'm6'), /U12/],
    ['M7 二進位垃圾一律放行', (s) => mutate(s, 'if (!looksLikeText(buf)) return bad', 'if (false) return bad', 'm7'), /U15/],
    ['M8 CFB 不放行（過度防護：誤擋正常 xls）', (s) => mutate(s, 'return { ok: true, buf, kind: \'xls\' };', "return bad('BAD_UPLOAD', 'x');", 'm8'), /U2/],
  ];
  let allRed = true;
  for (const [label, fn, expect] of muts) {
    const mp = path.join(ROOT, 'lib', 'xlsxSafe__mut_1.js');
    fs.writeFileSync(mp, fn(src));
    delete require.cache[mp];
    R = { pass: 0, fail: 0, failures: [] };
    try { await unit(require(mp), S); } catch (e) { R.fail++; R.failures.push('U1 例外（變異讓解析直接丟錯）：' + e.message); } finally { try { fs.unlinkSync(mp); } catch (_) { /* 已刪 */ } }
    const red = expect ? R.failures.some((f) => expect.test(f)) : R.failures.length === 0;   // expect=null：等價變異，預期不紅
    if (!red) allRed = false;
    console.log((red ? (expect ? '  變異被抓到（紅）' : '  等價變異（如預期仍綠）') : '  ✗ 變異未被抓到') + '：' + label + '  → ' + (R.failures.map((f) => f.slice(0, 3)).join(',') || '無紅'));
  }
  // 伺服器端：拿掉 xlsxGuard 接線，HTTP 層必須變紅
  const sp = path.join(ROOT, 'server.js');
  const ssrc = fs.readFileSync(sp, 'utf8');
  const smp = path.join(ROOT, 'server__mut_xg.js');
  fs.writeFileSync(smp, mutate(ssrc, "const xlsxGuard = require('./lib/xlsxSafe').uploadGuard();", "const xlsxGuard = (q, s, n) => n();", 'm9'));
  R = { pass: 0, fail: 0, failures: [] };
  try { await http(S, smp); } finally { try { fs.unlinkSync(smp); } catch (_) { /* 已刪 */ } await H.finish(); }
  const redH = R.failures.length >= ROUTES.length;   // 每條路由至少一個紅
  if (!redH) allRed = false;
  console.log((redH ? '  變異被抓到（紅）' : '  ✗ 變異未被抓到') + '：M9 伺服器拿掉 xlsxGuard 接線 → 紅 ' + R.failures.length + ' 項');
  console.log(allRed ? '\n變異全數被抓到' : '\n有變異沒被抓到');
  process.exit(allRed ? 0 : 1);
}
main().catch(async (e) => { console.error(e); try { await H.finish(); } catch (_) { /* ignore */ } process.exit(1); });
