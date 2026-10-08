#!/usr/bin/env node
'use strict';
/**
 * 信件模組「變異測試」：證明安全護欄真的被測試守住。
 * 用法：
 *   node scripts/check-mail-mutation.js            跑全部變異
 *   node scripts/check-mail-mutation.js --only M01,M20   只跑指定編號
 *   node scripts/check-mail-mutation.js --list      只列出清單
 *
 * 做法：把 lib/mail、scripts/check-mail-*.js（與 scripts/mail-preview.js、_client/deep-link.js 若存在）複製到系統暫存資料夾，
 * 先跑一次「未破壞」的基準（必須全過），再對每個變異：在「副本」上做一處破壞性修改，執行該變異指定的測試腳本，
 * 預期結果是測試失敗（exit code != 0）＝被殺死（KILLED）。測試仍通過＝倖存（SURVIVED），代表那道護欄沒有測試保護。
 * 原始專案檔案完全不會被修改；暫存資料夾結束時刪除。
 *
 * 變異定義的欄位：
 *   id、area、desc、file（相對專案根）、find（必須在檔案中「恰好出現一次」，否則記為 NOT-APPLIED）、replace、
 *   tests（要跑的測試腳本，預設 ['scripts/check-mail-core.js']）
 *
 * 後續階段（render／outbox／dispatch／userEmail／link）請在 MUTATIONS 陣列尾端追加自己的變異，編號沿用 M7x、M8x… 。
 *
 * 註：單一變異若只拿掉「多層防禦中的一層」，可能因為其他層仍擋得住而成為等價變異（例如 normalizeEmail 的控制字元檢查，
 * 後面的 local／domain 字元白名單也會擋掉換行）。這類請改成破壞「最後一道」或同時破壞多層（edits 陣列），不要硬算成測試漏洞。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const DEFAULT_TESTS = ['scripts/check-mail-core.js'];
const GLUE = ['scripts/check-mail-glue.js'];
const plant = (name, domain) => name + String.fromCharCode(64) + domain;     // 整合膠水（quoteMail.js／routes.js）的單元測試

// prettier-ignore
const MUTATIONS = [
  // ── safety.js ──────────────────────────────────────────────────────────────
  { id: 'M01', area: 'safety', desc: '網域白名單改成「後綴比對」（evil-itts.com.tw、sub.itts.com.tw 會通過）',
    file: 'lib/mail/safety.js', find: "if (typeof d === 'string' && d.trim().toLowerCase() === domain) return true;", replace: "if (typeof d === 'string' && domain.endsWith(d.trim().toLowerCase())) return true;" },
  { id: 'M02', area: 'safety', desc: '網域白名單改成「包含比對」（itts.com.tw.evil.test 會通過）',
    file: 'lib/mail/safety.js', find: "if (typeof d === 'string' && d.trim().toLowerCase() === domain) return true;", replace: "if (typeof d === 'string' && domain.indexOf(d.trim().toLowerCase()) >= 0) return true;" },
  { id: 'M03', area: 'safety', desc: '拿掉非 ASCII（全形／同形字）檢查',
    file: 'lib/mail/safety.js', find: "if (RE_NON_ASCII.test(s)) return fail('NON_ASCII',", replace: "if (false) return fail('NON_ASCII'," },
  { id: 'M04', area: 'safety', desc: '先轉小寫再檢查非 ASCII（Kelvin 符號被洗成 ASCII k 而放行）',
    file: 'lib/mail/safety.js', find: "if (RE_NON_ASCII.test(s)) return fail('NON_ASCII',", replace: "if (RE_NON_ASCII.test(s.toLowerCase())) return fail('NON_ASCII'," },
  { id: 'M05', area: 'safety', desc: 'local part 字元白名單放寬成任意字元（% ! 連續點都通過）',
    file: 'lib/mail/safety.js', find: 'const RE_LOCAL = /^[a-z0-9_+-]+(?:\\.[a-z0-9_+-]+)*$/;', replace: 'const RE_LOCAL = /^.+$/;' },
  { id: 'M06', area: 'safety', desc: '網域語法檢查永遠通過（無點、連字號開頭、IP 形式都通過）',
    file: 'lib/mail/safety.js', find: "function isValidDomainLower(d) {\n  if (typeof d !== 'string'", replace: "function isValidDomainLower(d) {\n  return typeof d === 'string' && d.length > 0;\n  if (typeof d !== 'string'" },
  { id: 'M07', area: 'safety', desc: 'headerSafe 不再把 C0／C1 控制字元換成空白（NUL、U+0085 殘留）',
    file: 'lib/mail/safety.js', find: 's = s.replace(RE_CONTROL_TO_SPACE, \' \');', replace: '/* mutated: control chars kept */' },
  { id: 'M08', area: 'safety', desc: 'headerSafe 不再移除方向控制／零寬／Tag 字元',
    file: 'lib/mail/safety.js', find: "s = s.replace(RE_INVISIBLE, '');", replace: '/* mutated: invisible chars kept */' },
  { id: 'M09', area: 'safety', desc: 'headerSafe 改用 UTF-16 單元截斷（會把 emoji 切成兩半）',
    file: 'lib/mail/safety.js', find: 'return takeCodePoints(s, max).text.trimEnd();', replace: 'return s.slice(0, max).trimEnd();' },
  { id: 'M10', area: 'safety', desc: 'escHtml 不再跳脫 <',
    file: 'lib/mail/safety.js', find: ".replace(/</g, '&lt;')", replace: '' },
  { id: 'M11', area: 'safety', desc: 'escHtml 不再跳脫雙引號（屬性注入）',
    file: 'lib/mail/safety.js', find: ".replace(/\"/g, '&quot;')", replace: '' },
  { id: 'M12', area: 'safety', desc: 'maskEmail 洩漏完整 local part',
    file: 'lib/mail/safety.js', find: "return n.value.charAt(0) + '***' + n.value.slice(at);", replace: "return n.value.slice(0, at) + '***' + n.value.slice(at);" },
  { id: 'M13', area: 'safety', desc: 'clip／headerSafe 的 code point 計數不處理代理對',
    file: 'lib/mail/safety.js', find: 'i += (d >= 0xdc00 && d <= 0xdfff) ? 2 : 1;', replace: 'i += 1;' },
  { id: 'M14', area: 'safety', desc: 'normalizeEmail 的 trim 被拿掉（頭尾空白換行會殘留或被拒絕）',
    file: 'lib/mail/safety.js', find: 'const s = raw.trim();\n  if (!s) return fail(\'EMPTY\'', replace: 'const s = raw;\n  if (!s) return fail(\'EMPTY\'' },
  // ── config.js ──────────────────────────────────────────────────────────────
  { id: 'M20', area: 'config', desc: 'MAIL_MODE 非法值時預設成 live（而不是 off）',
    file: 'lib/mail/config.js', find: "    mode = 'off';\n  }", replace: "    mode = 'live';\n  }" },
  { id: 'M21', area: 'config', desc: 'normalizeMode 對無法辨識的值回 live',
    file: 'lib/mail/config.js', find: "return parseMode(raw) || 'off';", replace: "return parseMode(raw) || 'live';" },
  { id: 'M22', area: 'config', desc: 'parseMode 不再轉小寫（\' LIVE \' 不被視為 live）',
    file: 'lib/mail/config.js', find: 'const m = s.toLowerCase();', replace: 'const m = s;' },
  { id: 'M23', area: 'config', desc: 'parseMode 改成「包含 live 就當 live」（livee、xlivex 變 live）',
    file: 'lib/mail/config.js', find: 'return MODES.indexOf(m) >= 0 ? m : null;', replace: "return m.indexOf('live') >= 0 ? 'live' : (MODES.indexOf(m) >= 0 ? m : null);" },
  { id: 'M24', area: 'config', desc: 'asciiTrim 改成 Unicode trim（尾端 NBSP 的 live 被當成 live）',
    file: 'lib/mail/config.js', find: 'function asciiTrim(s) {\n  return s.replace(', replace: 'function asciiTrim(s) {\n  return s.trim();\n  return s.replace(' },
  { id: 'M25', area: 'config', desc: 'redirect 缺目標時不再降級為 log（redirect 沒目標卻維持 redirect）',
    file: 'lib/mail/config.js', find: "if (mode === 'redirect' && !redirectTo) {", replace: "if (false) {" },
  { id: 'M26', area: 'config', desc: 'publicConfig 洩漏 redirectTo 原值',
    file: 'lib/mail/config.js', find: 'redirectConfigured: !!c.redirectTo,', replace: 'redirectConfigured: !!c.redirectTo, redirectTo: c.redirectTo,' },
  { id: 'M27', area: 'config', desc: 'clientSecret 變成可列舉屬性（Object.assign／展開會帶走）',
    file: 'lib/mail/config.js', find: "Object.defineProperty(g, 'clientSecret', { value: clientSecret, enumerable: false,", replace: "Object.defineProperty(g, 'clientSecret', { value: clientSecret, enumerable: true," },
  { id: 'M28', area: 'config', desc: 'graph.toJSON 洩漏 clientSecret',
    file: 'lib/mail/config.js', find: 'return { configured: configured };', replace: 'return { configured: configured, clientSecret: clientSecret };' },
  { id: 'M29', area: 'config', desc: 'APP_BASE_URL 允許非 localhost 的 http',
    file: 'lib/mail/config.js', find: "if (u.protocol === 'http:' && isLocal) return u.origin;", replace: "if (u.protocol === 'http:') return u.origin;" },
  { id: 'M30', area: 'config', desc: 'Vercel 上仍然提供 previewDir（會嘗試寫唯讀檔案系統並把信件內容留在磁碟）',
    file: 'lib/mail/config.js', find: "const onVercel = str(e, 'VERCEL') !== '';", replace: 'const onVercel = false;' },
  { id: 'M31', area: 'config', desc: 'APP_BASE_URL 主機名稱不再嚴格檢查（引號注入 href）',
    file: 'lib/mail/config.js', find: "if (!/^[a-z0-9]([a-z0-9.-]*[a-z0-9])?$/.test(u.hostname)) return '';", replace: '' },
  { id: 'M32', area: 'config', desc: '省略 env 參數時不讀 process.env',
    file: 'lib/mail/config.js', find: 'function getMailConfig(env = process.env) {', replace: 'function getMailConfig(env) {' },
  // ── recipients.js ──────────────────────────────────────────────────────────
  { id: 'M40', area: 'recipients', desc: 'username 比對改成不分大小寫（MGR 命中 mgr）',
    file: 'lib/mail/recipients.js', find: 'const u = index.get(un);', replace: 'const u = index.get(un) || index.get(un.toLowerCase());' },
  { id: 'M41', area: 'recipients', desc: '不再排除操作者本人',
    file: 'lib/mail/recipients.js', find: "if (actor !== '' && un === actor) {", replace: 'if (false) {' },
  { id: 'M42', area: 'recipients', desc: '不再檢查收件網域白名單',
    file: 'lib/mail/recipients.js', find: 'if (!isAllowedDomain(norm.value, allowed)) {', replace: 'if (false) {' },
  { id: 'M43', area: 'recipients', desc: '不再去除重複收件人',
    file: 'lib/mail/recipients.js', find: 'if (seen.has(un)) {', replace: 'if (false) {' },
  { id: 'M44', area: 'recipients', desc: '不再略過停用帳號',
    file: 'lib/mail/recipients.js', find: 'if (isInactive(u)) {', replace: 'if (false) {' },
  { id: 'M45', area: 'recipients', desc: 'deliver 使用未正規化的原始 email',
    file: 'lib/mail/recipients.js', find: 'deliver.push({ username: un, email: norm.value,', replace: 'deliver.push({ username: un, email: String(rawEmail),' },
  { id: 'M46', area: 'recipients', desc: 'email 不經 normalizeEmail 驗證就接受（多位址／換行注入）',
    file: 'lib/mail/recipients.js', find: 'if (!norm.ok) {', replace: 'if (false) {' },
  // ── visibility.js ──────────────────────────────────────────────────────────
  { id: 'M50', area: 'visibility', desc: '顧問看得到金額',
    file: 'lib/mail/visibility.js', find: "['consultant', { amount: false,", replace: "['consultant', { amount: true," },
  { id: 'M51', area: 'visibility', desc: '秘書看得到客戶名',
    file: 'lib/mail/visibility.js', find: "['secretary',  { amount: true,  margin: true,  tier: true,  customer: false,", replace: "['secretary',  { amount: true,  margin: true,  tier: true,  customer: true," },
  { id: 'M52', area: 'visibility', desc: 'itemPrices 變成 true',
    file: 'lib/mail/visibility.js', find: '    itemPrices: false,\n    reason: row.reason,', replace: '    itemPrices: true,\n    reason: row.reason,' },
  { id: 'M53', area: 'visibility', desc: '業務本人看得到毛利率',
    file: 'lib/mail/visibility.js', find: "['owner',      { amount: false, margin: false,", replace: "['owner',      { amount: false, margin: true," },
  { id: 'M54', area: 'visibility', desc: '未知 kind 退回最寬鬆的 mgr1 政策',
    file: 'lib/mail/visibility.js', find: "const row = typeof kind === 'string' ? POLICY.get(kind) : undefined;", replace: "const row = (typeof kind === 'string' ? POLICY.get(kind) : undefined) || POLICY.get('mgr1');" },
  { id: 'M55', area: 'visibility', desc: 'kind 比對改成不分大小寫',
    file: 'lib/mail/visibility.js', find: "const row = typeof kind === 'string' ? POLICY.get(kind) : undefined;", replace: "const row = typeof kind === 'string' ? POLICY.get(kind.toLowerCase()) : undefined;" },
  { id: 'M56', area: 'visibility', desc: '顧問看得到業務姓名變成看不到、且 items 關閉（顧問信失去用途）',
    file: 'lib/mail/visibility.js', find: "customer: false, owner: true,  items: true,  project: true,  reason: '顧問", replace: "customer: false, owner: false, items: false, project: true,  reason: '顧問" },
  { id: 'M57', area: 'visibility', desc: '董事長看不到金額',
    file: 'lib/mail/visibility.js', find: "['chairman',   { amount: true,", replace: "['chairman',   { amount: false," },
  // ── events.js ──────────────────────────────────────────────────────────────
  { id: 'M60', area: 'events', desc: 'dedupeKey 允許空白 stepKey（誤去重／不去重）',
    file: 'lib/mail/events.js', find: 'if (strErr(ev.stepKey, LIMITS.stepKey, { required: true, noControl: true })) {', replace: 'if (false) {' },
  { id: 'M61', area: 'events', desc: 'dedupeKey 漏掉收件人（同一單所有收件人共用一把鍵，只有第一人收到信）',
    file: 'lib/mail/events.js', find: "return ev.type + ':' + ev.quoteId + ':' + u + ':' + ev.stepKey;", replace: "return ev.type + ':' + ev.quoteId + ':' + ev.stepKey;" },
  { id: 'M62', area: 'events', desc: 'dedupeKey 不跳脫帳號中的冒號與百分號（不同組合撞同一把鍵）',
    file: 'lib/mail/events.js', find: "const u = username.replace(/%/g, '%25').replace(/:/g, '%3A');", replace: 'const u = username;' },
  { id: 'M63', area: 'events', desc: 'dedupeKey 把帳號轉小寫（只差大小寫的兩人共用一把鍵）',
    file: 'lib/mail/events.js', find: "const u = username.replace(/%/g, '%25')", replace: "const u = username.toLowerCase().replace(/%/g, '%25')" },
  { id: 'M64', area: 'events', desc: 'validateEvent 接受負的營收金額',
    file: 'lib/mail/events.js', find: 'if (!Number.isSafeInteger(n.revenueCents) || n.revenueCents < 0) return', replace: 'if (!Number.isSafeInteger(n.revenueCents)) return' },
  { id: 'M65', area: 'events', desc: 'validateEvent 接受 NaN／Infinity 的毛利率',
    file: 'lib/mail/events.js', find: "if (typeof n.marginPct !== 'number' || !isFinite(n.marginPct)) return", replace: "if (typeof n.marginPct !== 'number') return" },
  { id: 'M66', area: 'events', desc: 'validateEvent 接受小數／字串金額（不是整數 cents）',
    file: 'lib/mail/events.js', find: 'if (!Number.isSafeInteger(n.revenueCents) || n.revenueCents < 0) return', replace: "if (typeof n.revenueCents !== 'number' || n.revenueCents < 0) return" },
  { id: 'M67', area: 'events', desc: '單據 id 格式放寬（允許斜線與點，可造出路徑／連結注入）',
    file: 'lib/mail/events.js', find: 'const QUOTE_ID_RE = /^[A-Za-z0-9_-]{1,64}$/;', replace: 'const QUOTE_ID_RE = /^[A-Za-z0-9_.\\/-]{1,64}$/;' },
  { id: 'M68', area: 'events', desc: 'E4 可以沒有 result（結果通知沒有結果）',
    file: 'lib/mail/events.js', find: 'const required = !!allowedKinds;', replace: 'const required = false;' },
  { id: 'M69', area: 'events', desc: 'at 不再檢查真實日曆日期（2026-02-30 通過）',
    file: 'lib/mail/events.js', find: ' || !isRealCalendarDate(ev.at)', replace: '' },
  // ── userEmail.js（編號 M90–M107）──────────────────────────────────────────────
  { id: 'M90', area: 'userEmail', desc: '唯一性比對改用原始字串（大小寫不同的同一位址不算重複）',
    file: 'lib/mail/userEmail.js', find: "return n.ok ? n.value : '';", replace: "return n.ok ? u.email : '';" },
  { id: 'M91', area: 'userEmail', desc: '唯一性檢查不排除本人（自己沿用原位址也被當成重複）',
    file: 'lib/mail/userEmail.js', find: 'if (holders[i] !== selfUsername) {', replace: 'if (true) {' },
  { id: 'M92', area: 'userEmail', desc: '唯一性檢查整段拿掉（兩個帳號可共用同一位址）',
    file: 'lib/mail/userEmail.js', find: '  if (index) {\n    const holders', replace: '  if (false) {\n    const holders' },
  { id: 'M93', area: 'userEmail', desc: 'validateUserEmail 不檢查網域白名單',
    file: 'lib/mail/userEmail.js', find: 'if (!isAllowedDomain(n.value, allowed)) {', replace: 'if (false) {' },
  { id: 'M94', area: 'userEmail', desc: 'undefined 被當成「清除」（路由漏帶欄位會無聲清掉 email）',
    file: 'lib/mail/userEmail.js', find: "if (raw === undefined) return fail('MISSING', '未提供 Email 欄位（要清除請明確傳入空字串）');", replace: "if (raw === undefined) return { ok: true, value: '', code: 'CLEARED' };" },
  { id: 'M95', area: 'userEmail', desc: '批次匯入：同批內 Email 重複不標錯',
    file: 'lib/mail/userEmail.js', find: "if (list.length > 1) list.forEach((i) => mark(i, 'DUP_EMAIL_IN_BATCH', ", replace: "if (false) list.forEach((i) => mark(i, 'DUP_EMAIL_IN_BATCH', " },
  { id: 'M96', area: 'userEmail', desc: '批次匯入：同一帳號出現多次不標錯（後者無聲覆蓋前者）',
    file: 'lib/mail/userEmail.js', find: "if (list.length > 1) list.forEach((i) => mark(i, 'DUP_USER_IN_BATCH', ", replace: "if (false) list.forEach((i) => mark(i, 'DUP_USER_IN_BATCH', " },
  { id: 'M97', area: 'userEmail', desc: '批次匯入上限放寬成 5000 行',
    file: 'lib/mail/userEmail.js', find: 'maxRows: 500,', replace: 'maxRows: 5000,' },
  { id: 'M98', area: 'userEmail', desc: '批次匯入的帳號比對不分大小寫（ALICE 會對到 alice）',
    file: 'lib/mail/userEmail.js', find: 'const user = byName.get(f.userTok);', replace: 'const user = byName.get(f.userTok) || [...byName.entries()].find(([k]) => k.toLowerCase() === f.userTok.toLowerCase());' },
  { id: 'M99', area: 'userEmail', desc: '批次匯入：帳號只取第一欄、其餘全算 Email（帳號含空白的行解析錯誤）',
    file: 'lib/mail/userEmail.js', find: 'while (j > 0 && !isSepChar(s.charAt(j - 1))) j--;', replace: 'j = 0; while (j < s.length && !isSepChar(s.charAt(j))) j++; while (j < s.length && isSepChar(s.charAt(j))) j++;' },
  { id: 'M100', area: 'userEmail', desc: '稽核文字洩漏新位址',
    file: 'lib/mail/userEmail.js', find: 'if (!hadOld) return AUDIT_TEXT.SET;', replace: "if (!hadOld) return 'email 設定為 ' + newEmail;" },
  { id: 'M101', area: 'userEmail', desc: 'applyBulkPlan 就地修改傳入的 users',
    file: 'lib/mail/userEmail.js', find: 'const out = users.map((u) => (isObj(u) ? { ...u } : u));', replace: 'const out = users;' },
  { id: 'M102', area: 'userEmail', desc: 'applyBulkPlan 不重新驗證 rows（竄改的網域／重複位址直接寫入）',
    file: 'lib/mail/userEmail.js', find: 'if (!v.ok) { skipped.push({ username: safeText(un, 64), code: v.code }); continue; }', replace: 'if (!v.ok) { v.value = row.email; }' },
  { id: 'M103', area: 'userEmail', desc: '缺 Email 報表列出停用帳號',
    file: 'lib/mail/userEmail.js', find: 'if (isInactive(u)) return;', replace: '/* mutated: inactive listed */' },
  { id: 'M104', area: 'userEmail', desc: '缺 Email 報表把 sealManagers（報價章管理人）也當成簽核收件人',
    file: 'lib/mail/userEmail.js', find: "costProviders: 'costProvider' });", replace: "costProviders: 'costProvider', sealManagers: 'costProvider' });" },
  { id: 'M105', area: 'userEmail', desc: '缺 Email 報表把格式非法誤報成 NO_EMAIL（與 resolveRecipients 不一致）',
    file: 'lib/mail/userEmail.js', find: "if (!n.ok) return 'BAD_EMAIL';", replace: "if (!n.ok) return 'NO_EMAIL';" },
  { id: 'M106', area: 'userEmail', desc: '缺 Email 報表的稱呼順序改成 顯示名稱 > 暱稱（與信件稱呼不一致）',
    file: 'lib/mail/userEmail.js', find: 'const cands = [u.nickname, u.displayName, username];', replace: 'const cands = [u.displayName, u.nickname, username];' },
  { id: 'M107', area: 'userEmail', desc: '缺 Email 報表不排序',
    file: 'lib/mail/userEmail.js', find: 'out.sort((a, b) => (a.username < b.username ? -1 : (a.username > b.username ? 1 : 0)));', replace: '/* mutated: unsorted */' },
  // ── link.js（編號 M110–M119）──────────────────────────────────────────────────
  { id: 'M110', area: 'link', desc: '網站根網址不重新驗證（javascript: 或 http://evil 也能組進連結）',
    file: 'lib/mail/link.js', find: "if (probe.warnings.length) throw new MailLinkError('BAD_BASE_URL', '網站根網址不合法');", replace: '' },
  { id: 'M111', area: 'link', desc: 'cost 旗標改成「任何 truthy 都算」（?cost=false、陣列也變填成本）',
    file: 'lib/mail/link.js', find: "return v === true || v === 1 || v === '1';", replace: 'return !!v;' },
  { id: 'M112', area: 'link', desc: 'buildQuoteLink／deepHash 不驗證單據 id（路徑／連結注入）',
    file: 'lib/mail/link.js', find: 'function assertQuoteId(quoteId) {', replace: 'function assertQuoteId(quoteId) {\n  return;' },
  { id: 'M113', area: 'link', desc: '跳板頁不再回 Cache-Control: no-store',
    file: 'lib/mail/link.js', find: "'Cache-Control': 'no-store',", replace: "'Cache-Control': 'public, max-age=3600'," },
  { id: 'M114', area: 'link', desc: '跳板頁改用行內 script 導向（CSP 一收緊就壞）',
    file: 'lib/mail/link.js', find: "'<meta http-equiv=\"refresh\" content=\"0;url=' + t + '\">\\n',", replace: "'<script>location.replace(\"' + t + '\")</script>\\n'," },
  { id: 'M115', area: 'link', desc: '無效 id 的頁面回 200（等於告訴掃描器每個 id 都存在）',
    file: 'lib/mail/link.js', find: 'return { status: 404, headers: pageHeaders(), body: NOT_FOUND_BODY };', replace: 'return { status: 200, headers: pageHeaders(), body: NOT_FOUND_BODY };' },
  { id: 'M116', area: 'link', desc: '回應自帶的 CSP 允許任何網站嵌入（frame-ancestors *）',
    file: 'lib/mail/link.js', find: "frame-ancestors 'none'\";", replace: 'frame-ancestors *";' },
  { id: 'M117', area: 'link', desc: 'link.js 不再引用 events 的 isValidQuoteId，改用寬鬆的自訂版本',
    file: 'lib/mail/link.js', find: "const { isValidQuoteId } = require('./events');", replace: "const isValidQuoteId = (id) => typeof id === 'string' && id.length > 0;" },
  { id: 'M118', area: 'link', desc: '跳板頁漏掉 X-Robots-Tag: noindex',
    file: 'lib/mail/link.js', find: "'X-Robots-Tag': 'noindex',", replace: '' },
  // ── _client/deep-link.js（編號 M130–M139）─────────────────────────────────────
  { id: 'M130', area: 'deep-link', desc: '格式放寬到允許斜線（可塞進路徑／open redirect）',
    file: '_client/deep-link.js', find: 'var HASH_RE = /^#quote:[A-Za-z0-9_-]{1,64}(?::cost)?$/;', replace: 'var HASH_RE = /^#quote:[A-Za-z0-9_\\/-]{1,64}(?::cost)?$/;' },
  { id: 'M131', area: 'deep-link', desc: 'consume 不清除（同一分頁下次登入又被帶去舊單據）',
    file: '_client/deep-link.js', find: 'try { s.removeItem(STORAGE_KEY); } catch (e) { /* 清不掉就算了 */ }', replace: '/* mutated: not removed */' },
  { id: 'M132', area: 'deep-link', desc: 'consume 不再驗證儲存內容（被竄改的值原樣放行）',
    file: '_client/deep-link.js', find: "return isValidHash(v) ? v : '';", replace: "return typeof v === 'string' ? v : '';" },
  { id: 'M133', area: 'deep-link', desc: 'remember 不驗證格式',
    file: '_client/deep-link.js', find: 'if (!isValidHash(h)) return false;', replace: "if (typeof h !== 'string') return false;" },
  { id: 'M134', area: 'deep-link', desc: '存取 sessionStorage 的例外不處理（隱私模式下整頁壞掉）',
    file: '_client/deep-link.js', find: '    } catch (e) {\n      return null;\n    }', replace: '    } catch (e) {\n      throw e;\n    }' },
  { id: 'M135', area: 'deep-link', desc: 'setItem 的例外不處理（儲存空間滿了就丟例外）',
    file: '_client/deep-link.js', find: '      s.setItem(STORAGE_KEY, h);\n      return true;\n    } catch (e) {\n      return false;\n    }', replace: '      s.setItem(STORAGE_KEY, h);\n      return true;\n    } catch (e) {\n      throw e;\n    }' },
  { id: 'M136', area: 'deep-link', desc: 'hashForRedirect 只讀不清（登入導向後值還留著）',
    file: '_client/deep-link.js', find: 'function hashForRedirect() {\n    return consume();\n  }', replace: "function hashForRedirect() {\n    var s = getStorage();\n    try { var v = s.getItem(STORAGE_KEY); return isValidHash(v) ? v : ''; } catch (e) { return ''; }\n  }" },
  { id: 'M138', area: 'deep-link', desc: '讀取 window.location 的例外不處理（remember() 在奇怪環境下整頁壞掉）',
    file: '_client/deep-link.js', find: '} catch (e) { /* 取不到就當沒有 */ }', replace: '} catch (e) { throw e; }' },
  { id: 'M137', area: 'deep-link', desc: '無效輸入會清掉先前記住的值',
    file: '_client/deep-link.js', find: 'if (!isValidHash(h)) return false;\n    var s = getStorage();', replace: "if (!isValidHash(h)) { var s0 = getStorage(); try { s0.removeItem(STORAGE_KEY); } catch (e) {} return false; }\n    var s = getStorage();" },
  // ── render／templates（編號 M150–M219）：全部只跑 check-mail-render.js，證明「渲染層的測試」自己就守得住這些護欄 ──
  // 可見性與 headerSafe／escHtml：破壞底層模組，渲染測試（不靠 check-mail-core.js）也必須失敗
  { id: 'M150', area: 'render/visibility', desc: '顧問看得到金額（visibility 表被改）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/visibility.js', find: "['consultant', { amount: false,", replace: "['consultant', { amount: true," },
  { id: 'M151', area: 'render/visibility', desc: '秘書看得到客戶名', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/visibility.js', find: "['secretary',  { amount: true,  margin: true,  tier: true,  customer: false,", replace: "['secretary',  { amount: true,  margin: true,  tier: true,  customer: true," },
  { id: 'M152', area: 'render/visibility', desc: '業務本人看得到毛利率', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/visibility.js', find: "['owner',      { amount: false, margin: false,", replace: "['owner',      { amount: false, margin: true," },
  { id: 'M153', area: 'render/visibility', desc: '未知 kind 退回最寬鬆的 mgr1 政策', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/visibility.js', find: "const row = typeof kind === 'string' ? POLICY.get(kind) : undefined;", replace: "const row = (typeof kind === 'string' ? POLICY.get(kind) : undefined) || POLICY.get('mgr1');" },
  { id: 'M154', area: 'render/visibility', desc: '董事長看不到金額', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/visibility.js', find: "['chairman',   { amount: true,", replace: "['chairman',   { amount: false," },
  { id: 'M155', area: 'render/headerSafe', desc: 'headerSafe／safeText 不再把控制字元（CR／LF／NUL）換成空白——主旨可被注入標頭', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/safety.js', find: 's = s.replace(RE_CONTROL_TO_SPACE, \' \');', replace: '/* mutated: control chars kept */' },
  { id: 'M156', area: 'render/headerSafe', desc: 'headerSafe／safeText 不再移除方向控制（RLO）與零寬字元', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/safety.js', find: "s = s.replace(RE_INVISIBLE, '');", replace: '/* mutated: invisible chars kept */' },
  { id: 'M157', area: 'render/escHtml', desc: 'escHtml 不再跳脫 <（HTML 注入）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/safety.js', find: ".replace(/</g, '&lt;')", replace: '' },
  { id: 'M158', area: 'render/escHtml', desc: 'escHtml 不再跳脫雙引號', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/safety.js', find: ".replace(/\"/g, '&quot;')", replace: '' },
  { id: 'M159', area: 'render/escHtml', desc: 'escHtml 不再跳脫 &（實體可被偽造）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/safety.js', find: ".replace(/&/g, '&amp;')", replace: '' },
  // render.js
  { id: 'M160', area: 'render/subject', desc: '主旨帶出客戶名', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "[no, proj].filter(Boolean).join(' ')", replace: "[no, proj, safeText(ev.company, 20)].filter(Boolean).join(' ')" },
  // 主旨有兩層清理（各欄位先 headerSafe／safeText，整行再 headerSafe 一次），單拿掉一層是等價變異，所以這裡同時拿掉兩層
  { id: 'M161', area: 'render/subject', desc: '主旨完全不經 headerSafe／safeText（單號與專案名稱的方向字元、換行、U+2028 全殘留）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "const no = headerSafe(ev.quoteNo, 40);\n  const proj = showProject ? safeText(ev.projectName, SUBJECT_PROJECT_MAX) : '';\n  return headerSafe(prefix + [no, proj].filter(Boolean).join(' ') + (SUBJECT_SUFFIX[ev.type] || ''), SUBJECT_MAX);",
    replace: "const no = String(ev.quoteNo);\n  const proj = showProject ? String(ev.projectName || '') : '';\n  return prefix + [no, proj].filter(Boolean).join(' ') + (SUBJECT_SUFFIX[ev.type] || '');" },
  { id: 'M162', area: 'render/subject', desc: '主旨的專案名稱不經清理（換行可注入標頭）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "const proj = showProject ? safeText(ev.projectName, SUBJECT_PROJECT_MAX) : '';", replace: "const proj = showProject ? String(ev.projectName || '') : '';" },
  { id: 'M163', area: 'render/visibility', desc: '決策條的金額格不看 visibility（任何收件人都有金額）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'if (vis.amount) {', replace: 'if (true) {' },
  { id: 'M164', area: 'render/visibility', desc: '客戶列不看 visibility（秘書／顧問看得到客戶名）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "const company = vis.customer ? safeText(ev.company, 100) : '';", replace: 'const company = safeText(ev.company, 100);' },
  { id: 'M165', area: 'render/visibility', desc: '業務列不看 visibility（秘書／業務本人看得到業務名）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "const owner = vis.owner ? safeText(ev.ownerLabel, 60) : '';", replace: 'const owner = safeText(ev.ownerLabel, 60);' },
  { id: 'M166', area: 'render/visibility', desc: '品項表不看 visibility（簽核人信也列品項）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'if (vis.items && Array.isArray(ev.items) && ev.items.length > 0) {', replace: 'if (Array.isArray(ev.items) && ev.items.length > 0) {' },
  { id: 'M167', area: 'render/visibility', desc: '操作人列不受 visibility.owner 控管（E1／E2／E6 的操作人＝業務，秘書看得到）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'const hide = (ACTOR_IS_OWNER[type] && !vis.owner) || (ownerRaw !== \'\' && actor === ownerRaw);', replace: "const hide = (ownerRaw !== '' && actor === ownerRaw);" },
  { id: 'M168', area: 'render/visibility', desc: '操作人與業務同名時不隱藏（E3 的操作人若是業務，秘書看得到業務名）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'const hide = (ACTOR_IS_OWNER[type] && !vis.owner) || (ownerRaw !== \'\' && actor === ownerRaw);', replace: 'const hide = (ACTOR_IS_OWNER[type] && !vis.owner);' },
  { id: 'M169', area: 'render/decision', desc: 'E6（請勿簽核）也放決策條', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'const SHOWS_DECISION = Object.freeze({ E1_SUBMIT: true, E3_NEXT_STEP: true });', replace: 'const SHOWS_DECISION = Object.freeze({ E1_SUBMIT: true, E3_NEXT_STEP: true, E6_WITHDRAWN: true });' },
  { id: 'M170', area: 'render/amount', desc: '金額元的部分改用浮點除法 Math.floor(cents / 100)', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'const yuan = padded.slice(0, padded.length - 2);', replace: 'const yuan = String(Math.floor(cents / 100));' },
  { id: 'M171', area: 'render/amount', desc: '金額不加千分位', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "const grouped = yuan.replace(/\\B(?=(\\d{3})+(?!\\d))/g, ',');", replace: 'const grouped = yuan;' },
  { id: 'M172', area: 'render/tone', desc: '色塊對照反了（層級 1 紅、層級 3 綠）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "const TIER_TONE = Object.freeze({ 1: 'green', 2: 'amber', 3: 'red' });", replace: "const TIER_TONE = Object.freeze({ 1: 'red', 2: 'amber', 3: 'green' });" },
  { id: 'M173', area: 'render/tone', desc: '核決層級未知時色塊變綠（應為灰）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "? TIER_TONE[lv] : 'grey';", replace: "? TIER_TONE[lv] : 'green';" },
  { id: 'M174', area: 'render/tone', desc: '層級 1 的標籤不再是「可核」', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "if (level === 1) return tierLabel + '可核';", replace: '' },
  { id: 'M175', area: 'render/tone', desc: '負毛利不再標示「毛利為負」', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'const loss = n.gpCents < 0 || n.marginPct < 0;', replace: 'const loss = false;' },
  { id: 'M176', area: 'render/decision', desc: 'numbers 為 null 時輸出「空的決策條」而不是整段省略', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'if (!n) return null;', replace: 'if (!n) return { cells: [] };' },
  { id: 'M177', area: 'render/matrix', desc: '允許把「請填成本」寄給秘書', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "E2_COST_REQUEST: Object.freeze(['consultant']),", replace: "E2_COST_REQUEST: Object.freeze(['consultant', 'secretary'])," },
  { id: 'M178', area: 'render/robust', desc: '不做事件快照（getter 在驗證後改值可繞過驗證）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'const ev = snapshotEvent(rawEv);', replace: 'const ev = rawEv;' },
  { id: 'M179', area: 'render/robust', desc: '不做收件人快照（viewer.kind 在檢查後被改成別的值）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'const viewer = snapshotViewer(rawViewer);', replace: 'const viewer = rawViewer;' },
  { id: 'M180', area: 'render/link', desc: '站台網址不再做格式檢查（引號注入 href）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "if (typeof b !== 'string' || !RE_BASE_URL.test(b)) throw", replace: "if (typeof b !== 'string') throw" },
  { id: 'M181', area: 'render/link', desc: '站台網址允許非 localhost 的 http', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "if (b.indexOf('http://') === 0 && !m) throw", replace: 'if (false) throw' },
  { id: 'M182', area: 'render/link', desc: '請填成本的連結不帶 ?cost=1', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "(opts && opts.cost ? '?cost=1' : '')", replace: "''" },
  { id: 'M183', area: 'render/content', desc: '原因不再截成 200 字', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'safeText(ev.result.reason, 200)', replace: 'safeText(ev.result.reason, 2000)' },
  { id: 'M184', area: 'render/content', desc: '品項列數上限 30 放寬成 300', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'const MAX_ITEM_ROWS = 30;', replace: 'const MAX_ITEM_ROWS = 300;' },
  { id: 'M185', area: 'render/content', desc: '品項超過上限時不顯示「…另 N 項」', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "moreText: extra > 0 ? '…另 ' + extra + ' 項' : ''", replace: "moreText: ''" },
  { id: 'M186', area: 'render/content', desc: '時間不轉台北時區', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'new Date(ms + 8 * 3600 * 1000)', replace: 'new Date(ms)' },
  { id: 'M187', area: 'render/content', desc: 'preheader 帶出金額', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'preheader: safeText(pre + [quoteNo, project].filter(Boolean).join(\' \'), 90),', replace: "preheader: safeText(pre + [quoteNo, project, decision ? decision.cells[0].value : ''].filter(Boolean).join(' '), 90)," },
  { id: 'M188', area: 'render/robust', desc: '收件人 kind 不再檢查是否為已知類型', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "if (typeof viewer.kind !== 'string' || !isKnownKind(viewer.kind)) throw", replace: "if (typeof viewer.kind !== 'string') throw" },
  { id: 'M189', area: 'render/robust', desc: '略過 validateEvent', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'const v = validateEvent(ev);', replace: 'const v = { ok: true };' },
  { id: 'M190', area: 'render/robust', desc: '未預期例外直接往外丟（不包成 MailRenderError）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "throw new MailRenderError('INTERNAL', '信件渲染發生未預期的錯誤（' + (e && e.name ? e.name : 'Error') + '）');", replace: 'throw e;' },
  { id: 'M191', area: 'render/robust', desc: '拿掉信件大小上限', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "if (Buffer.byteLength(html, 'utf8') > MAX_MAIL_BYTES || Buffer.byteLength(text, 'utf8') > MAX_MAIL_BYTES) {", replace: 'if (false) {' },
  { id: 'M192', area: 'render/content', desc: '頁尾沒有「機密，請勿轉寄」', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "'機密，請勿轉寄；本信由 ITTS-CRM 自動發送，請勿直接回覆。',", replace: "'本信由 ITTS-CRM 自動發送。'," },
  { id: 'M193', area: 'render/meta', desc: 'meta.hasAmount 恆為 false', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: "hasAmount: !!(model.decision && model.decision.cells.some((c) => c.kind === 'amount')),", replace: 'hasAmount: false,' },
  { id: 'M194', area: 'render/meta', desc: 'meta 帶出客戶名', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'kind: viewer.kind,', replace: 'kind: viewer.kind, company: ev.company,' },
  // templates.js
  { id: 'M200', area: 'templates/escape', desc: '<title> 不跳脫（主旨內的 </title><script> 可注入）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: "o.push('<title>' + esc(m.title) + '</title>');", replace: "o.push('<title>' + m.title + '</title>');" },
  { id: 'M201', area: 'templates/escape', desc: '資訊列的值不跳脫', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: ", esc(r.value)),", replace: ', r.value),' },
  { id: 'M202', area: 'templates/escape', desc: '駁回原因不跳脫', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: 'esc(r.reason)', replace: 'r.reason' },
  { id: 'M203', area: 'templates/escape', desc: '品項說明不跳脫', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: 'esc(r[0])', replace: 'r[0]' },
  { id: 'M204', area: 'templates/escape', desc: '稱呼不跳脫（收件人名稱可注入）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: 'esc(m.greeting)', replace: 'm.greeting' },
  { id: 'M205', area: 'templates/escape', desc: '色塊標籤不跳脫（tierLabel 可注入）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: 'esc(c.tag)', replace: 'c.tag' },
  { id: 'M206', area: 'templates/layout', desc: '拿掉 color-scheme meta', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: 'o.push(\'<meta name="color-scheme" content="light dark">\');', replace: '' },
  { id: 'M207', area: 'templates/layout', desc: '拿掉深色模式媒體查詢', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: "'@media (prefers-color-scheme:dark){',", replace: "'@media not all{'," },
  { id: 'M208', area: 'templates/layout', desc: '色塊儲存格沒有 bgcolor 屬性（Outlook Word 引擎會變白底白字）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: '      bgcolor: bg,', replace: '      bgcolor: null,' },
  { id: 'M209', area: 'templates/layout', desc: '色塊沒有文字標籤（只靠顏色傳達）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: 'if (c.tag) {', replace: 'if (false) {' },
  { id: 'M210', area: 'templates/layout', desc: 'body 沒有 bgcolor', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: "o.push('<body class=\"bg-page\" bgcolor=\"' + P.pageBg + '\" style=", replace: "o.push('<body class=\"bg-page\" style=" },
  { id: 'M211', area: 'templates/layout', desc: 'preheader 沒有隱藏', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: '<div class="preheader" style="display:none;', replace: '<div class="preheader" style="display:block;' },
  { id: 'M212', area: 'templates/text', desc: '純文字版漏掉決策條', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: 'if (m.decision) {', replace: 'if (false) {' },
  { id: 'M213', area: 'templates/text', desc: '純文字版把網址折行', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: 'L.push(m.button.url);', replace: "wrapText(m.button.url, 40, '', '').forEach((x) => L.push(x));" },
  { id: 'M214', area: 'templates/text', desc: '純文字行寬上限放寬成 200', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: 'const MAX_TEXT_COLS = 76;', replace: 'const MAX_TEXT_COLS = 200;' },
  { id: 'M215', area: 'templates/text', desc: '欄寬計算把全形字當成 1 欄', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: '(cp >= 0x20000 && cp <= 0x3fffd)) return 2;', replace: '(cp >= 0x20000 && cp <= 0x3fffd)) return 1;' },
  // ── 派送層（outbox／adapter／熔斷／scrub／傳輸／派送器）：tests 指向 check-mail-outbox.js 或 check-mail-dispatch.js ──
  { id: 'M300', area: 'outbox/dedupe', desc: 'JSON／記憶體 adapter 不再依 dedupeKey 去重（同一事件可入列多次）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'if (st.jobs[i].dedupeKey === rec.dedupeKey) return { inserted: false, record: st.jobs[i] };', replace: 'if (false) return { inserted: false, record: st.jobs[i] };' },
  { id: 'M301', area: 'outbox/pg', desc: 'Postgres insert 拿掉 ON CONFLICT (dedupe_key) DO NOTHING', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'ON CONFLICT (dedupe_key) DO NOTHING\nRETURNING *', replace: 'RETURNING *' },
  { id: 'M302', area: 'outbox/lease', desc: '租約沒到期的 sending 也會被重新領取', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "if (j.status !== 'sending') return false;\n    const t = msOf(j.leaseUntil);\n    return !Number.isFinite(t) || t <= nowMs;", replace: "if (j.status !== 'sending') return false;\n    return true;" },
  { id: 'M303', area: 'outbox/claim', desc: 'claimDue 領取後不把狀態改成 sending（並行領取會重複）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "picked.forEach((j) => {\n        j.status = 'sending';", replace: "picked.forEach((j) => {\n        j.status = 'pending';" },
  { id: 'M304', area: 'outbox/claim', desc: 'claimById 不檢查是否到期', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'if (!j || !isDuePending(j, nowMs)) return null;', replace: "if (!j || j.status !== 'pending') return null;" },
  { id: 'M305', area: 'outbox/retry', desc: 'Retry-After 不再優先於排程延遲', tests: ['scripts/check-mail-outbox.js', 'scripts/check-mail-dispatch.js'],
    file: 'lib/mail/outbox.js', find: 'if (typeof ra === \'number\' && isFinite(ra) && ra > 0) delay = Math.min(Math.max(ra, 1), MAX_RETRY_AFTER_SEC);', replace: '/* mutated: Retry-After ignored */' },
  { id: 'M306', area: 'outbox/retry', desc: '退避延遲索引差一（第一次重試等 300 秒而不是 60 秒）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: 'let delay = retry.delays[Math.min(Math.max(rec.attemptsInRound, 1) - 1, retry.delays.length - 1)];', replace: 'let delay = retry.delays[Math.min(Math.max(rec.attemptsInRound, 1), retry.delays.length - 1)];' },
  { id: 'M307', area: 'outbox/retry', desc: '一輪嘗試上限少一次（首次 + 2 次重試就 failed）', tests: ['scripts/check-mail-outbox.js', 'scripts/check-mail-dispatch.js'],
    file: 'lib/mail/outbox.js', find: 'const roundLimit = 1 + retry.retries;', replace: 'const roundLimit = retry.retries;' },
  { id: 'M308', area: 'outbox/retry', desc: 'Retry-After 沒有 24 小時上限', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: 'delay = Math.min(Math.max(ra, 1), MAX_RETRY_AFTER_SEC);', replace: 'delay = Math.max(ra, 1);' },
  { id: 'M309', area: 'outbox/privacy', desc: 'toMasked 不經清理（呼叫端傳完整位址就原樣存進去）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: 'toMasked: sanitizeMasked(job.toMasked),', replace: "toMasked: typeof job.toMasked === 'string' ? job.toMasked : ''," },
  { id: 'M310', area: 'outbox/whitelist', desc: '呼叫端可以指定紀錄的初始狀態（偽造 sent）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: "status: skipped ? 'skipped' : 'pending',", replace: "status: typeof job.status === 'string' ? job.status : 'pending'," },
  { id: 'M311', area: 'outbox/whitelist', desc: '紀錄與儲存層不再用白名單（信件內容欄位 html 會被存進去）', tests: ['scripts/check-mail-outbox.js'],
    edits: [
      { file: 'lib/mail/outbox.js', find: '      id: genId(),\n      type: job.type,', replace: '      id: genId(),\n      html: job.html,\n      type: job.type,' },
      { file: 'lib/mail/outboxAdapters.js', find: "    id: r.id,\n    type: strOr(r.type, ''),", replace: "    id: r.id,\n    html: r.html,\n    type: strOr(r.type, '')," },
    ] },
  { id: 'M312', area: 'outbox/purge', desc: 'purge 連 pending／sending 也清掉', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'const keep = st.jobs.filter((j) => !(TERMINAL_STATUSES.indexOf(j.status) >= 0 && msOf(j.updatedAt) < cutMs));', replace: 'const keep = st.jobs.filter((j) => !(msOf(j.updatedAt) < cutMs));' },
  { id: 'M313', area: 'outbox/requeue', desc: 'requeue 不重置本輪次數（重送後只剩零次機會）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'j.attemptsInRound = 0;\n      j.requeues += 1;', replace: 'j.requeues += 1;' },
  { id: 'M314', area: 'outbox/breaker', desc: 'AUTH 錯誤不再直接開熔斷', tests: ['scripts/check-mail-outbox.js', 'scripts/check-mail-dispatch.js'],
    file: 'lib/mail/outbox.js', find: "const w = c === 'AUTH' ? bcfg.failures : (BREAKER_WEIGHT[c] || 1);", replace: 'const w = (BREAKER_WEIGHT[c] || 1);' },
  { id: 'M315', area: 'outbox/breaker', desc: '成功不再重置熔斷計數', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: 'st = { failures: [], openUntil: 0 };', replace: '/* mutated: no reset */' },
  { id: 'M316', area: 'outbox/breaker', desc: '與傳輸健康無關的錯誤碼（REJECTED／BAD_MESSAGE…）也計入熔斷', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: 'if (HEALTH_CODES.indexOf(c) < 0) return viewOf(st, t);', replace: '' },
  { id: 'M317', area: 'outbox/breaker', desc: '半開狀態的第一次失敗不會立刻重新開啟', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: 'if (halfOpen || total >= bcfg.failures) {', replace: 'if (total >= bcfg.failures) {' },
  { id: 'M318', area: 'outbox/breaker', desc: '失敗計數不看時間視窗（很久以前的失敗也累計）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: 'st.failures = st.failures.filter((f) => f.t > t - bcfg.windowSec * 1000);', replace: '' },
  { id: 'M319', area: 'outbox/breaker', desc: '熔斷開啟期間的失敗會延長開啟時間', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: 'if (st.openUntil > t) return viewOf(st, t);', replace: '' },
  { id: 'M320', area: 'outbox/scrub', desc: 'markFailed 存錯誤訊息前不再清理（密鑰／GUID／email 原樣進 outbox）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: 'const msg = scrubMessage(o.msg);', replace: "const msg = typeof o.msg === 'string' ? o.msg.slice(0, 200) : '';" },
  { id: 'M321', area: 'outbox/json', desc: 'JSON adapter 改成直接覆寫主檔（不是暫存檔＋rename 原子替換）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "fsx.writeFileSync(tmp, text, { encoding: 'utf8', flag: 'wx' });\n      renameWithRetry(tmp, file);", replace: "fsx.writeFileSync(file, text, 'utf8');" },
  { id: 'M322', area: 'outbox/json', desc: 'JSON 檔案損壞時丟例外（而不是備份後從空開始）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "backupCorrupt(raw, 'move', 'PARSE_OR_SHAPE');\n      return freshState();", replace: "throw new MailOutboxError('STORE_IO', 'corrupt');" },
  { id: 'M323', area: 'outbox/json', desc: '寫入失敗時不清掉暫存檔', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'try { fsx.unlinkSync(tmp); } catch (e2) { /* 暫存檔可能根本沒建立 */ }', replace: '' },
  { id: 'M324', area: 'outbox/pg', desc: 'Postgres 的 expireExhausted 不再執行（租約過期且次數用盡的工作永遠卡在 sending）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'const rows = await q(SQL.expireExhausted, [now, LEASE_EXPIRED_MSG, maxRoundAttempts]);', replace: 'const rows = [];' },
  { id: 'M325', area: 'outbox/pg', desc: 'Postgres 領取拿掉 FOR UPDATE SKIP LOCKED', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'LIMIT $3\n          FOR UPDATE SKIP LOCKED)', replace: 'LIMIT $3\n          FOR UPDATE)' },
  { id: 'M326', area: 'outbox/pg', desc: 'Postgres finishAttempt 沒有 attempts 樂觀鎖', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "WHERE id = $1 AND status = 'sending' AND attempts = $2", replace: "WHERE id = $1 AND status = 'sending' AND $2 IS NOT NULL" },
  { id: 'M327', area: 'outbox/pg', desc: 'Postgres purge 連非終態也刪', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "WHERE status IN ('sent', 'failed', 'skipped', 'cancelled') AND updated_at < $1::timestamptz", replace: 'WHERE updated_at < $1::timestamptz' },
  { id: 'M328', area: 'outbox/lease', desc: 'markFailed 不再比對領取當時的 attempts（兩層檢查同時拿掉：租約被搶走後舊工作者仍可回報）', tests: ['scripts/check-mail-outbox.js'],
    edits: [
      { file: 'lib/mail/outbox.js', find: "if (Number.isSafeInteger(o.attempts) && o.attempts !== rec.attempts) return { ok: false, reason: 'LOST_LEASE' };", replace: '' },
      { file: 'lib/mail/outbox.js', find: 'expectAttempts: Number.isSafeInteger(o.attempts) ? o.attempts : rec.attempts,', replace: 'expectAttempts: rec.attempts,' },
    ] },
  // ── 傳輸層 ──
  { id: 'M330', area: 'transport/mode', desc: '不認得的模式（含缺設定）改走真傳輸，而不是 null', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/transports.js', find: "  if (mode === 'live') return real || createGraphTransport({ config: c, fetchImpl: d.fetchImpl });\n  return createNullTransport();", replace: '  return real || createGraphTransport({ config: c, fetchImpl: d.fetchImpl });' },
  { id: 'M331', area: 'transport/mode', desc: 'log 模式改走真傳輸', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/transports.js', find: "if (mode === 'log') return makeLog();", replace: "if (mode === 'log') return real || makeLog();" },
  { id: 'M332', area: 'transport/mode', desc: 'redirect 模式不包 redirect 傳輸（直接用真傳輸，收件人不改寫）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/transports.js', find: "if (mode === 'redirect') return createRedirectTransport({ config: c, inner: real || makeLog() });", replace: "if (mode === 'redirect') return real || makeLog();" },
  { id: 'M333', area: 'transport/redirect', desc: 'redirect 傳輸不改寫收件人', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/transports.js', find: 'to: [target.value],', replace: 'to: v.msg.to,' },
  { id: 'M334', area: 'transport/redirect', desc: 'redirectTo 缺失時退回寄給原收件人（拿掉 NOT_CONFIGURED 護欄）', tests: ['scripts/check-mail-dispatch.js'],
    edits: [
      { file: 'lib/mail/transports.js', find: "if (!target.ok) return failure('NOT_CONFIGURED', 'redirect 模式需要有效的 MAIL_REDIRECT_TO');", replace: '' },
      { file: 'lib/mail/transports.js', find: 'to: [target.value],', replace: 'to: target.ok ? [target.value] : v.msg.to,' },
    ] },
  { id: 'M335', area: 'transport/mode', desc: 'selectTransport 不再正規化模式字串（\'LIVE \'、\' Redirect \' 不被認得）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/transports.js', find: 'const mode = normalizeMode(c.mode);', replace: 'const mode = c.mode;' },
  { id: 'M336', area: 'transport/mode', desc: 'live 判斷改成「包含 live」（olive、live1 都會選到真傳輸）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/transports.js', find: "if (mode === 'live') return real || createGraphTransport({ config: c, fetchImpl: d.fetchImpl });", replace: "if (mode === 'live' || String(c.mode).toLowerCase().indexOf('live') >= 0) return real || createGraphTransport({ config: c, fetchImpl: d.fetchImpl });" },
  { id: 'M337', area: 'transport/redirect', desc: 'redirect 把原訊息的多餘欄位（cc／bcc）一起轉給內層', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/transports.js', find: 'try { r = await inner.send(out, sendOpts); }', replace: 'try { r = await inner.send(Object.assign({}, msg, out), sendOpts); }' },
  { id: 'M338', area: 'transport/validate', desc: '訊息驗證放行主旨裡的換行（標頭注入）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/transports.js', find: "if (RE_HEADER_BAD.test(msg.subject)) return { ok: false, error: '主旨含換行或控制字元' };", replace: '' },
  { id: 'M339', area: 'transport/log', desc: 'log 傳輸的檔名不剝除路徑字元（路徑穿越）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/transports.js', find: "const t = String(s || '').replace(/[^A-Za-z0-9_-]/g, '_').slice(0, max);", replace: "const t = String(s || '').slice(0, max);" },
  { id: 'M340', area: 'transport/log', desc: '沒有 previewDir（Vercel）時 log 傳輸把信件內容印到 console', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/transports.js', find: 'if (!dir) return { ok: true, providerId };', replace: 'if (!dir) { console.log(JSON.stringify(v.msg)); return { ok: true, providerId }; }' },
  { id: 'M341', area: 'transport/result', desc: 'AUTH／REJECTED 等不再自動視為 permanent（完全信任傳輸自己的宣告）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/transports.js', find: 'permanent: isPermanentCode(c) || !!(extra && extra.permanent === true)', replace: 'permanent: !!(extra && extra.permanent === true)' },
  // ── 派送器 ──
  { id: 'M350', area: 'dispatch/stale', desc: '寄送前一刻不再用 isStillValid 取消過期的信', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: 'if (r.value !== true) {', replace: 'if (false) {' },
  { id: 'M351', area: 'dispatch/stale', desc: 'isStillValid 出錯（無法確認）時照寄', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "if (!r.ok) { await failJob(job, c, summary, 'STALE_CHECK', '無法確認單據狀態', { permanent: false }); return; }", replace: '' },
  { id: 'M352', area: 'dispatch/audit', desc: '稽核 detail 不再遮罩長得像 email 的帳號名稱', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "function userInDetail(u) { return (typeof u === 'string' && u.indexOf('@') >= 0) ? maskEmail(u) : safeText(u, 64); }", replace: 'function userInDetail(u) { return safeText(u, 64); }' },
  { id: 'M353', area: 'dispatch/audit', desc: '成功寄出也寫稽核', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: '      summary.sent += 1;\n      return;', replace: "      summary.sent += 1;\n      await audit('QUOTE_MAIL_SKIPPED', c.operator, c.quoteNo, c.type, job.toUser, 'reason=SENT');\n      return;" },
  { id: 'M354', area: 'dispatch/mode', desc: 'dispatch 在 off 模式仍照常渲染與寄送', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "    if (mode === 'off') {\n      const seen = new Set();", replace: "    if (false) {\n      const seen = new Set();" },
  { id: 'M355', area: 'dispatch/mode', desc: 'drainDue 在 off 模式仍照常領取並寄送', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "if (currentMode() === 'off') return;", replace: '' },
  { id: 'M356', area: 'dispatch/timeout', desc: '寄送不再套用總逾時（傳輸永不 resolve 就永遠卡住）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "result = await Promise.race([sp.then(normalizeResult, () => failure('NETWORK', '傳輸拒絕')), timeoutP]);", replace: "result = await sp.then(normalizeResult, () => failure('NETWORK', '傳輸拒絕'));" },
  { id: 'M357', area: 'dispatch/retry', desc: '所有失敗都當成 permanent（不重試）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: 'result.message, { permanent: result.permanent === true, retryAfterSec: result.retryAfterSec });', replace: 'result.message, { permanent: true, retryAfterSec: result.retryAfterSec });' },
  { id: 'M358', area: 'dispatch/breaker', desc: '寄送前不看熔斷狀態（「不當次寄」與「排到熔斷結束」兩層同時拿掉）', tests: ['scripts/check-mail-dispatch.js'],
    edits: [
      { file: 'lib/mail/dispatcher.js', find: "const hold = br.open || (nowMs() - startMs > budgetMs) || r.email === '';", replace: "const hold = (nowMs() - startMs > budgetMs) || r.email === '';" },
      { file: 'lib/mail/dispatcher.js', find: 'nextAttemptAt: br.open && br.until ? br.until : undefined,', replace: 'nextAttemptAt: undefined,' },
    ] },
  { id: 'M359', area: 'dispatch/breaker', desc: '傳輸失敗不回報給熔斷器', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: 'if (HEALTH_CODES.indexOf(result.code) >= 0) await breakerRecord(false, result.code);', replace: '' },
  { id: 'M360', area: 'dispatch/drain', desc: 'drainDue 不再比對重建事件的去重鍵（舊關卡的信會被補寄）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: 'if (key !== job.dedupeKey) {', replace: 'if (false) {' },
  { id: 'M361', area: 'dispatch/mode', desc: 'log 模式「寄成功」記為 sent（而不是 skipped/MODE_LOG，之後無法 requeue）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: 'if (logOnly) {', replace: 'if (false) {' },
  { id: 'M362', area: 'dispatch/skip', desc: '操作者本人也被記成 outbox 紀錄', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "if (s.reason === 'DUP' || s.reason === 'ACTOR') { summary.skipped.push({ username: s.username, reason: s.reason }); continue; }", replace: "if (s.reason === 'DUP') { summary.skipped.push({ username: s.username, reason: s.reason }); continue; }" },
  { id: 'M363', area: 'dispatch/concurrency', desc: '同時寄送數不設上限', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: 'const concurrency = clampInt(d.concurrency, 1, 8, 3);', replace: 'const concurrency = 100;' },
  { id: 'M364', area: 'dispatch/validate', desc: '派送器不再自己驗證訊息（主旨換行、超大、多餘欄位直接交給傳輸）', tests: ['scripts/check-mail-dispatch.js'],
    edits: [
      { file: 'lib/mail/dispatcher.js', find: "if (!vm.ok) { await failJob(job, c, summary, 'BAD_MESSAGE', vm.error, { permanent: true }); return; }", replace: '' },
      { file: 'lib/mail/dispatcher.js', find: 'const msg = vm.msg;', replace: "const msg = { to: [c.email], subject: mail.subject, html: mail.html, text: mail.text, tag: c.type };" },
    ] },
  { id: 'M365', area: 'dispatch/skip', desc: '解析收件人時不排除操作者', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: 'resolved = resolveRecipients(list, { users: ur.value, actorUsername: actor, config });', replace: "resolved = resolveRecipients(list, { users: ur.value, actorUsername: '', config });" },
  { id: 'M366', area: 'dispatch/audit', desc: '重複觸發時被略過的收件人重複寫稽核', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "return r.created ? 'created' : 'exists';", replace: "return 'created';" },
  { id: 'M367', area: 'dispatch/drain', desc: 'drainDue 沒有 rebuild 也照樣領取工作', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "if (typeof rebuild !== 'function') { summary.errors += 1; return; }", replace: '' },
  { id: 'M368', area: 'dispatch/render', desc: '渲染失敗當成可重試（應為 permanent）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "await failJob(job, c, summary, 'RENDER', why, { permanent: true });", replace: "await failJob(job, c, summary, 'RENDER', why, { permanent: false });" },
  { id: 'M369', area: 'dispatch/privacy', desc: '錯誤訊息不再遮蔽已知字串（主旨、專案名、客戶名）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: 'msg: scrubMessage(message, c.known)', replace: 'msg: scrubMessage(message)' },
  { id: 'M370', area: 'dispatch/mode', desc: 'config.mode 的第二層正規化被拿掉（\'LIVE \' 不再被視為 live）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: 'const currentMode = () => normalizeMode(config.mode);', replace: 'const currentMode = () => config.mode;' },
  { id: 'M390', area: 'dispatch/dedupe', desc: '派送器不看 enqueue 的 created（重複觸發時照樣往下走）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "if (!enq.created) { summary.skipped.push({ username: r.username, reason: 'DUPLICATE' }); return; }", replace: '' },
  { id: 'M392', area: 'dispatch/budget', desc: '不再套用單次 dispatch 的時間預算', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: '(nowMs() - startMs > budgetMs)', replace: 'false' },
  { id: 'M394', area: 'dispatch/mode', desc: 'off 模式不在 outbox 留下 skipped/MODE_OFF 紀錄', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "const out = await recordSkipped(ev, u, kindBy.get(u), 'MODE_OFF');", replace: "const out = 'created';" },
  { id: 'M396', area: 'outbox/json', desc: 'JSON adapter 不再去除檔頭 BOM（有 BOM 的檔案被當成損壞）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'if (raw.length > 0 && raw.charCodeAt(0) === 0xfeff) raw = raw.slice(1);', replace: '' },
  { id: 'M397', area: 'outbox/vercel', desc: 'JSON adapter 在 Vercel 上不再拒絕建立', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "refuseOnVercel('json', opts);", replace: '' },
  { id: 'M398', area: 'outbox/vercel', desc: '記憶體 adapter 在 Vercel 上不再拒絕建立', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "refuseOnVercel('memory', opts);", replace: '' },
  // ── 整合模擬（scripts/check-mail-integration.js 單獨負責殺死；跨模組接線與端到端行為）──
  { id: 'M400', area: 'integration/dispatch', desc: '寄送前不再呼叫 isStillValid（過期的「請簽核」信照寄）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/dispatcher.js', find: "if (typeof d.isStillValid === 'function') {", replace: 'if (false) {' },
  { id: 'M401', area: 'integration/dispatch', desc: 'off 模式不再只記錄（繼續往下解析收件人並嘗試寄送）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/dispatcher.js', find: "if (mode === 'off') {", replace: 'if (false) {' },
  { id: 'M402', area: 'integration/dispatch', desc: '重複觸發不再回報 DUPLICATE 並結束（繼續往下領取）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/dispatcher.js', find: "if (!enq.created) { summary.skipped.push({ username: r.username, reason: 'DUPLICATE' }); return; }", replace: 'if (!enq.created) { summary.queued += 1; }' },
  { id: 'M403', area: 'integration/transport', desc: 'redirect 傳輸不再改寫收件人（測試信會寄給真實收件人）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/transports.js', find: 'to: [target.value],', replace: 'to: v.msg.to,' },
  { id: 'M404', area: 'integration/transport', desc: 'log 模式改用真傳輸（log 模式會真的寄出信件）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/transports.js', find: "if (mode === 'log') return makeLog();", replace: "if (mode === 'log') return real || makeLog();" },
  { id: 'M405', area: 'integration/outbox', desc: '退避永遠用第一個延遲（60 秒），不再 60／300／900 拉長', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/outbox.js', find: 'let delay = retry.delays[Math.min(Math.max(rec.attemptsInRound, 1) - 1, retry.delays.length - 1)];', replace: 'let delay = retry.delays[0];' },
  { id: 'M406', area: 'integration/outbox', desc: 'AUTH 錯誤不再一次就開熔斷（權重 1）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/outbox.js', find: "const w = c === 'AUTH' ? bcfg.failures : (BREAKER_WEIGHT[c] || 1);", replace: 'const w = BREAKER_WEIGHT[c] || 1;' },
  { id: 'M407', area: 'integration/outbox', desc: '429 的 Retry-After 不再優先於預設退避', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/outbox.js', find: 'if (typeof ra === \'number\' && isFinite(ra) && ra > 0) delay = Math.min(Math.max(ra, 1), MAX_RETRY_AFTER_SEC);   // Retry-After 優先', replace: '' },
  { id: 'M408', area: 'integration/recipients', desc: '停用帳號不再被信件層略過（膠水沒先過濾時會寄給停用帳號）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/recipients.js', find: 'if (isInactive(u)) {', replace: 'if (false) {' },
  { id: 'M409', area: 'integration/visibility', desc: '秘書（董事會關）看得到客戶名', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/visibility.js', find: "['secretary',  { amount: true,  margin: true,  tier: true,  customer: false,", replace: "['secretary',  { amount: true,  margin: true,  tier: true,  customer: true," },
  { id: 'M410', area: 'integration/visibility', desc: '顧問信失去兩道防線：E1（簽核通知）可以寄給顧問，且顧問的可見性表開放金額與毛利率', tests: ['scripts/check-mail-integration.js'],
    edits: [
      { file: 'lib/mail/render.js', find: 'E1_SUBMIT: Object.freeze(APPROVERS.slice()),', replace: "E1_SUBMIT: Object.freeze(APPROVERS.concat(['consultant']))," },
      { file: 'lib/mail/visibility.js', find: "['consultant', { amount: false, margin: false, tier: false,", replace: "['consultant', { amount: true,  margin: true,  tier: true," },
    ] },
  { id: 'M411', area: 'integration/render', desc: '撤回／作廢（E6）的信也放決策條（請勿簽核的信出現金額與毛利率）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/render.js', find: 'const SHOWS_DECISION = Object.freeze({ E1_SUBMIT: true, E3_NEXT_STEP: true });', replace: 'const SHOWS_DECISION = Object.freeze({ E1_SUBMIT: true, E3_NEXT_STEP: true, E6_WITHDRAWN: true });' },
  { id: 'M412', area: 'integration/dispatch', desc: '傳輸健康錯誤不再記入熔斷（連續失敗永遠不會開熔斷）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/dispatcher.js', find: 'if (HEALTH_CODES.indexOf(result.code) >= 0) await breakerRecord(false, result.code);', replace: '' },
  { id: 'M413', area: 'integration/dispatch', desc: '熔斷開著時仍然當次嘗試寄送（同時拆掉「hold」與「下次嘗試時間＝熔斷結束」兩層；每次送簽都吃逾時）', tests: ['scripts/check-mail-integration.js'],
    edits: [
      { file: 'lib/mail/dispatcher.js', find: "const hold = br.open || (nowMs() - startMs > budgetMs) || r.email === '';", replace: "const hold = (nowMs() - startMs > budgetMs) || r.email === '';" },
      { file: 'lib/mail/dispatcher.js', find: 'nextAttemptAt: br.open && br.until ? br.until : undefined,', replace: '' },
    ] },
  { id: 'M414', area: 'integration/dispatch', desc: '重試時不再重新解析收件人（排隊後被停用／移除 email 的帳號仍會被寄）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/dispatcher.js', find: 'if (!rs.deliver.length) {', replace: 'if (false) {' },
  { id: 'M415', area: 'integration/link', desc: 'link.js 的連結前綴與 render.js 的 /q/ 漂移', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/link.js', find: "const JUMP_PATH_PREFIX = '/q/';", replace: "const JUMP_PATH_PREFIX = '/quote/';" },
  { id: 'M416', area: 'integration/config', desc: 'MAIL_MODE 非法值時預設成 live（安全預設被破壞）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/config.js', find: "    mode = 'off';\n  }", replace: "    mode = 'live';\n  }" },
  { id: 'M417', area: 'integration/events', desc: '去重鍵不含收件人（同事件的第二位收件人被當成重複）', tests: ['scripts/check-mail-integration.js'],
    file: 'lib/mail/events.js', find: "return ev.type + ':' + ev.quoteId + ':' + u + ':' + ev.stepKey;", replace: "return ev.type + ':' + ev.quoteId + ':' + ev.stepKey;" },
  // ── scrub ──
  { id: 'M380', area: 'scrub', desc: 'scrub 不再遮蔽 GUID（租戶／應用程式 ID）', tests: ['scripts/check-mail-outbox.js', 'scripts/check-mail-dispatch.js'],
    file: 'lib/mail/scrub.js', find: "\n    .replace(RE_GUID, '[id]')", replace: '' },
  { id: 'M381', area: 'scrub', desc: 'scrub 不再遮蔽 email', tests: ['scripts/check-mail-outbox.js', 'scripts/check-mail-dispatch.js'],
    file: 'lib/mail/scrub.js', find: "\n    .replace(RE_EMAIL_LIKE, '[email]')", replace: '' },
  { id: 'M382', area: 'scrub', desc: 'scrub 不再遮蔽金額樣式的數字', tests: ['scripts/check-mail-outbox.js', 'scripts/check-mail-dispatch.js'],
    file: 'lib/mail/scrub.js', find: "\n    .replace(RE_BIGNUM, '[數字]')", replace: '' },
  { id: 'M383', area: 'scrub', desc: 'scrub 不再遮蔽 Bearer 權杖', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/scrub.js', find: "s = s.replace(RE_BEARER, 'Bearer [已遮蔽]')\n    .replace(RE_JWT", replace: 's = s\n    .replace(RE_JWT' },
  { id: 'M384', area: 'scrub', desc: 'scrub 不再遮蔽 client_secret=… 這類名稱＝值', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/scrub.js', find: "\n    .replace(RE_KV_SECRET, '$1=[已遮蔽]')", replace: '' },
  { id: 'M385', area: 'scrub', desc: 'scrub 不再遮蔽長的不透明字串（金鑰）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/scrub.js', find: "\n    .replace(RE_OPAQUE, '[已遮蔽]')", replace: '' },
  // ── 獨立審查後的修正（M420–M459）：JSON 備份上限、Postgres RLS／動態 SQL 守衛／actorLabel、drainDue 預算與熔斷、LEASE_EXPIRED 可見、極端毛利率、深色模式邊線 ──
  { id: 'M420', area: 'review/json-backup', desc: 'JSON 損壞備份不再去重（同一個損壞狀態，每次唯讀操作都複製一份備份，檔案無上限增長）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'if (diag.seen.has(sig)) return null;', replace: '' },
  { id: 'M421', area: 'review/json-backup', desc: 'JSON 損壞備份不再修剪（備份檔數量沒有上限）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: '      pruneBackups();\n    }\n    note({', replace: '    }\n    note({' },
  { id: 'M422', area: 'review/json-backup', desc: '備份失敗也記成「已備份」（之後永遠不再重試備份）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'if (backup) {\n      diag.seen.add(sig);', replace: 'if (true) {\n      diag.seen.add(sig);' },
  { id: 'M423', area: 'review/json-backup', desc: 'diagnostics.recoveries 沒有上限', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'if (diag.recoveries.length > MAX_RECOVERY_NOTES) diag.recoveries.splice(0, diag.recoveries.length - MAX_RECOVERY_NOTES);', replace: '' },
  { id: 'M424', area: 'review/json-backup', desc: '修剪備份時刪掉最新的而不是最舊的', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'for (let i = 0; i < names.length - MAX_CORRUPT_BACKUPS; i++) {', replace: 'for (let i = MAX_CORRUPT_BACKUPS; i < names.length; i++) {' },
  { id: 'M425', area: 'review/pg-rls', desc: 'Postgres 建表流程不再執行兩條 RLS 語句（mail_outbox／mail_breaker 會被 Supabase PostgREST 匿名金鑰讀到）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "'createdIdx', 'rlsOutbox', 'createBreaker', 'rlsBreaker']", replace: "'createdIdx', 'createBreaker']" },
  { id: 'M426', area: 'review/pg-rls', desc: 'Postgres 建表流程只漏掉 mail_breaker 的 RLS', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "'createBreaker', 'rlsBreaker']", replace: "'createBreaker']" },
  { id: 'M427', area: 'review/pg-rls', desc: 'mail_outbox 的 RLS 語句變成 DISABLE', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "rlsOutbox: 'ALTER TABLE mail_outbox ENABLE ROW LEVEL SECURITY',", replace: "rlsOutbox: 'ALTER TABLE mail_outbox DISABLE ROW LEVEL SECURITY'," },
  { id: 'M428', area: 'review/pg-guard', desc: 'Postgres q() 拿掉「SQL 必須是常數」守衛（動態 SQL 可以送到資料庫）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "if (!SQL_SET.has(sql)) throw new MailOutboxError('DYNAMIC_SQL', '只允許執行 outboxAdapters.js 內定義的 SQL 常數');", replace: '' },
  { id: 'M429', area: 'review/pg-guard', desc: 'Postgres q() 拿掉參數型別守衛（物件參數可以送到資料庫）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: "if (!Array.isArray(p) || !p.every(isParamOk)) throw new MailOutboxError('BAD_PARAM', 'SQL 參數只能是字串、數字、布林、null 或字串陣列');", replace: '' },
  { id: 'M430', area: 'review/actorLabel', desc: 'actorLabel 不經 safeText 就入庫（控制字元、方向控制字元、超長字串進 outbox）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: "actorLabel: typeof job.actorLabel === 'string' ? safeText(job.actorLabel, 100) : ''," , replace: "actorLabel: typeof job.actorLabel === 'string' ? job.actorLabel : ''," },
  { id: 'M431', area: 'review/templates-attr', desc: 'escAttr 不再跳脫雙引號（屬性值可以跳出引號注入事件屬性）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: ".replace(/>/g, '&gt;').replace(/\"/g, '&quot;');\n}\nfunction attrs", replace: ".replace(/>/g, '&gt;');\n}\nfunction attrs" },
  { id: 'M432', area: 'review/drain', desc: 'drainDue 不看時間預算（傳輸卡住時一直領取到 limit 為止）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "if (nowMs() - startMs >= budget) { halted = true; summary.budgetExhausted = true; return; }", replace: '' },
  { id: 'M433', area: 'review/drain', desc: 'drainDue 每筆之間不重查熔斷（批次中途熔斷打開，仍把整批寄完）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "if (b.open) { halted = true; summary.breakerOpen = true; return; }", replace: '' },
  { id: 'M434', area: 'review/drain', desc: 'drainDue 的 limit 失效（忽略處理筆數上限）', tests: ['scripts/check-mail-dispatch.js'],
    edits: [
      { file: 'lib/mail/dispatcher.js', find: "        if (halted || taken >= limit) return;\n        if (nowMs()", replace: "        if (halted) return;\n        if (nowMs()" },
      { file: 'lib/mail/dispatcher.js', find: "        if (halted || taken >= limit) return;\n        if (b.open)", replace: "        if (halted) return;\n        if (b.open)" },
    ] },
  { id: 'M435', area: 'review/drain', desc: 'drainDue 預設預算上限放寬到 60 秒（超過 Vercel maxDuration）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: 'const defaultDrainBudgetMs = () => Math.min(20000, totalMs() * 2);', replace: 'const defaultDrainBudgetMs = () => Math.min(60000, totalMs() * 2);' },
  { id: 'M436', area: 'review/drain', desc: 'drainDue 不再先清掃 LEASE_EXPIRED（租約過期且用盡的 sending 永遠卡著）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: 'const swept = await outbox.expireStale();', replace: 'const swept = [];' },
  { id: 'M437', area: 'review/lease-expired', desc: '清掃出的 LEASE_EXPIRED 不列入 Summary.failed', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "summary.failed.push({ username: rec.toUser, code: 'LEASE_EXPIRED' });", replace: '' },
  { id: 'M438', area: 'review/lease-expired', desc: '清掃出的 LEASE_EXPIRED 不寫 QUOTE_MAIL_FAILED 稽核', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "await audit('QUOTE_MAIL_FAILED', 'system', safeText(rec.quoteNo, 40), EVENT_TYPES.indexOf(rec.type) >= 0 ? rec.type : 'UNKNOWN', rec.toUser, 'code=LEASE_EXPIRED');", replace: '' },
  { id: 'M439', area: 'review/lease-expired', desc: 'JSON／記憶體 adapter 的 expireExhausted 不回傳被改掉的紀錄（又變回靜默轉換）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: '      if (flipped.length) io.save(st);\n      return flipped;', replace: '      if (flipped.length) io.save(st);\n      return [];' },
  { id: 'M440', area: 'review/lease-expired', desc: 'Postgres expireExhausted 沒有 RETURNING *（被清掃的列回不來）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outboxAdapters.js', find: 'AND attempts_in_round >= $3\nRETURNING *', replace: 'AND attempts_in_round >= $3' },
  { id: 'M441', area: 'review/lease-expired', desc: 'outbox.claimDue 領取前不再預設清掃（租約過期且用盡的 sending 永遠不會變 failed）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: 'if (o.sweep !== false) await adapter.expireExhausted({ now: iso(t), maxRoundAttempts: roundLimit });', replace: '' },
  { id: 'M442', area: 'review/lease-expired', desc: 'outbox.claimDue 忽略 sweep:false（drainDue 想先看到清掃結果時被搶先靜默清掉）', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: 'if (o.sweep !== false) await adapter.expireExhausted(', replace: 'if (true) await adapter.expireExhausted(' },
  { id: 'M443', area: 'review/lease-expired', desc: 'createOutbox 不再要求 adapter 有 expireExhausted', tests: ['scripts/check-mail-outbox.js'],
    file: 'lib/mail/outbox.js', find: " || typeof adapter.expireExhausted !== 'function') {", replace: ') {' },
  { id: 'M444', area: 'review/margin', desc: 'marginText 整數位數上限退回 6 位（極端虧損單整封簽核信 BAD_EVENT）', tests: ['scripts/check-mail-core.js', 'scripts/check-mail-render.js', 'scripts/check-mail-integration.js'],
    file: 'lib/mail/events.js', find: 'const RE_MARGIN_TEXT = /^-?[0-9]{1,20}(?:\\.[0-9]{1,4})?%?$/;', replace: 'const RE_MARGIN_TEXT = /^-?[0-9]{1,6}(?:\\.[0-9]{1,4})?%?$/;' },
  { id: 'M445', area: 'review/margin', desc: 'marginText 長度上限退回 16 字', tests: ['scripts/check-mail-core.js', 'scripts/check-mail-render.js'],
    file: 'lib/mail/events.js', find: '  marginText: 32,', replace: '  marginText: 16,' },
  { id: 'M446', area: 'review/margin', desc: 'displayMargin 不再截斷極端值（信上印出一長串數字）', tests: ['scripts/check-mail-core.js', 'scripts/check-mail-render.js'],
    file: 'lib/mail/events.js', find: 'if (/^[0-9]+$/.test(intPart) && intPart.length > MARGIN_DISPLAY_MAX_INT_DIGITS) {', replace: 'if (false) {' },
  { id: 'M447', area: 'review/margin', desc: 'render 不再用 displayMargin 顯示毛利率', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/render.js', find: 'return displayMargin(marginText);', replace: "return String(marginText).replace(/%$/, '') + '%';" },
  { id: 'M448', area: 'review/margin', desc: 'marginText 格式檢查放寬到任意字串（可從毛利率文字注入 HTML）', tests: ['scripts/check-mail-core.js'],
    file: 'lib/mail/events.js', find: 'const RE_MARGIN_TEXT = /^-?[0-9]{1,20}(?:\\.[0-9]{1,4})?%?$/;', replace: 'const RE_MARGIN_TEXT = /^[^]*$/;' },
  { id: 'M449', area: 'review/dark-mode', desc: '深色模式 CSS 不再覆蓋決策條格子的邊線色（深色卡片上出現白線）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: "    '  .stack{border-color:' + D.cardBg + ' !important;}',\n", replace: '' },
  { id: 'M450', area: 'review/dark-mode', desc: '品項表外框沒有 bd class（深色模式下仍是亮灰框）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: "return table({ class: 'bd', width: '100%', style: 'border:1px solid ' + P.border + ';' }, headRow + bodyRows + more);", replace: "return table({ width: '100%', style: 'border:1px solid ' + P.border + ';' }, headRow + bodyRows + more);" },
  { id: 'M451', area: 'review/dark-mode', desc: '決策條格子的分隔線又寫死成另一種淺灰（不在深色覆蓋的顏色表內）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: "const border = i < n - 1 ? 'border-right:2px solid ' + P.cardBg + ';' : '';", replace: "const border = i < n - 1 ? 'border-right:2px solid ' + P.border + ';' : '';" },
  { id: 'M453', area: 'review/margin', desc: '決策條窄格子（30% 寬）的大字上限放寬回 12 字（「<-999999%」折成孤字）', tests: ['scripts/check-mail-render.js'],
    file: 'lib/mail/templates.js', find: 'const fits = widthPct >= 40 ? 12 : 8;', replace: 'const fits = 12;' },
  { id: 'M452', area: 'review/drain', desc: 'drainDue 的工作者領取失敗時例外往外丟（Promise.all 提早 reject，其他工作者在 drainDue 回傳後還在背景寄信）', tests: ['scripts/check-mail-dispatch.js'],
    file: 'lib/mail/dispatcher.js', find: "try { got = await outbox.claimDue({ limit: 1, leaseSec: leaseSec(), sweep: false }); } catch (e) { summary.errors += 1; halted = true; return; }", replace: "got = await outbox.claimDue({ limit: 1, leaseSec: leaseSec(), sweep: false });" },
  // ── 整合膠水（lib/mail/quoteMail.js、lib/mail/routes.js；測試：scripts/check-mail-glue.js）──────────────
  { id: 'M460', area: 'glue/kind', desc: "董事會關的秘書被當成 gm（信裡出現客戶名與業務名）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "if (tier === 'board') return user && user.role === 'secretary' ? 'secretary' : 'boardProxy';", replace: "if (tier === 'board') return 'gm';" },
  { id: 'M461', area: 'glue/numbers', desc: "tierLabel 用完整標籤「董事會決議（秘書代核）」而不是短標籤（色塊變成「需董事會決議（秘書代核）核准」）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "const tierLabel = d.board ? '董事會' : (lv ? TIER_SHORT[lv] : '');", replace: "const tierLabel = d.board ? '董事會決議（秘書代核）' : (lv ? TIER_SHORT[lv] : '');" },
  { id: 'M462', area: 'glue/numbers', desc: "numbersOf 不再檢查金額是整數（小數／負數／字串進信件）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "if (!Number.isSafeInteger(d.revenueCents) || d.revenueCents < 0 || !Number.isSafeInteger(d.gpCents)) return null;", replace: "if (false) return null;" },
  { id: 'M463', area: 'glue/step', desc: "stepOf 用原型鏈取 level（__proto__ 之類的 tier 命中原型屬性）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "const lv = Object.prototype.hasOwnProperty.call(STEP_LEVEL, step.tier) ? STEP_LEVEL[step.tier] : null;", replace: "const lv = STEP_LEVEL[step.tier];" },
  { id: 'M464', area: 'glue/items', desc: "E2 的品項夾帶單價（顧問信會看到價格）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "const row = { desc: clip(it && it.desc, 500) || '(未命名品項)', unit: clip(it && it.unit, 20) };", replace: "const row = { desc: clip(it && it.desc, 500) || '(未命名品項)', unit: clip(it && it.unit, 20), unitPrice: it && it.unitPrice };" },
  { id: 'M465', area: 'glue/mode', desc: "MAIL_MODE=off 時整合層仍然呼叫 dispatch（會碰儲存體）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "const enabled = () => !!outbox && normalizeMode(config.mode) !== 'off';", replace: "const enabled = () => !!outbox;" },
  { id: 'M466', area: 'glue/recipients', desc: "收件人過濾不排除操作者，同時 dispatch 也不傳 actorUsername（兩層一起壞，作廢信會寄給操作者本人）", tests: GLUE,
    edits: [{"file":"lib/mail/quoteMail.js","find":"if (!un || typeof un !== 'string' || un === actor || seen.has(un) || !H.isActive(ctx.users[un])) return;","replace":"if (!un || typeof un !== 'string' || seen.has(un) || !H.isActive(ctx.users[un])) return;"},{"file":"lib/mail/quoteMail.js","find":"const opts = { actorUsername: ctx.me, operatorLabel: operator };","replace":"const opts = { actorUsername: '', operatorLabel: operator };"}] },
  { id: 'M467', area: 'glue/recipients', desc: "收件人過濾不排除停用帳號（停用帳號也入列、記稽核噪音）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "if (!un || typeof un !== 'string' || un === actor || seen.has(un) || !H.isActive(ctx.users[un])) return;", replace: "if (!un || typeof un !== 'string' || un === actor || seen.has(un)) return;" },
  { id: 'M468', area: 'glue/stepKey', desc: "作廢信的去重鍵改用作廢「後」已被清空的 submittedAt（null#voided）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "stepKey: snap.submittedAt + '#' + spec.resultKind });", replace: "stepKey: String(q.approval && q.approval.submittedAt) + '#' + spec.resultKind });" },
  { id: 'M469', area: 'glue/stepKey', desc: "E2 的去重鍵不含 requestedAt（換顧問再換回、結構變動後重發都被當重複）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "{ stepKey: cf.requestedAt + '#' + q.costBy });", replace: "{ stepKey: 'x#' + q.costBy });" },
  { id: 'M470', area: 'glue/stepKey', desc: "E4 的去重鍵不含關卡序號（同一次送簽的第二關核准被當重複）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "stepKey: ap.submittedAt + '#r:' + spec.resultKind + ':' + idx });", replace: "stepKey: ap.submittedAt + '#r:' + spec.resultKind });" },
  { id: 'M471', area: 'glue/event', desc: "E4 approved 也帶原因（核准意見進信）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "ev.result = o.resultKind === 'rejected' && o.reason ?", replace: "ev.result = o.reason ?" },
  { id: 'M472', area: 'glue/valid', desc: "isStillValid 對送簽事件不檢查 approval.cur（單據已走到下一關仍照寄）", tests: GLUE,
    file: 'lib/mail/validity.js', find: "if (ap.cur !== idx) return no('STEP_MOVED');", replace: "/* mutated */" },
  { id: 'M473', area: 'glue/valid', desc: "isStillValid 對送簽事件不檢查收件人（被改派走的一級主管仍會收到）", tests: GLUE,
    file: 'lib/mail/validity.js', find: "return Array.isArray(list) && list.indexOf(to) >= 0 ? yes() : no('RECIPIENT_CHANGED');", replace: "return yes();" },
  { id: 'M474', area: 'glue/valid', desc: "isStillValid 對 E2 不比 requestedAt（舊的請求在重新請求後仍照寄）", tests: GLUE,
    file: 'lib/mail/validity.js', find: "return sk === cf.requestedAt + '#' + q.costBy ? yes() : no('COST_STATE');", replace: "return yes();" },
  { id: 'M475', area: 'glue/rebuild', desc: "rebuild 的 at 改用現在時間（重試信與首次寄出的內容不同）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "const base = { stepKey: sk, at: job.createdAt };", replace: "const base = { stepKey: sk, at: new Date(nowMs()).toISOString() };" },
  { id: 'M476', area: 'glue/rebuild', desc: "rebuild 的 E4 approved 也帶核准意見（makeEvent 的 rejected 才帶原因與 rebuild 的過濾兩層一起壞；單壞一層是等價變異）", tests: GLUE,
    edits: [
      { file: 'lib/mail/quoteMail.js', find: "reason: step && m[1] === 'rejected' ? step.comment : undefined", replace: "reason: step ? step.comment : undefined" },
      { file: 'lib/mail/quoteMail.js', find: "ev.result = o.resultKind === 'rejected' && o.reason ?", replace: "ev.result = o.reason ?" }] },
  { id: 'M477', area: 'glue/poll', desc: "pollDrain 沒有節流（每個 poll-bundle 都清理）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "if (poll.running || (poll.lastStartMs && t - poll.lastStartMs < POLL_MIN_INTERVAL_MS)) return false;", replace: "if (poll.running) return false;" },
  { id: 'M478', area: 'glue/poll', desc: "pollDrain 沒有逾時（傳輸卡住就卡住整個 poll-bundle）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "const r = await Promise.race([dispatcher.drainDue(POLL_DRAIN), timeout]);", replace: "const r = await dispatcher.drainDue(POLL_DRAIN);" },
  { id: 'M479', area: 'glue/status', desc: "emailStatusMap 一律回 OK（缺 email 的人沒有徽章）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "rep.forEach((r) => { out[r.username] = r.reason; });", replace: "/* mutated */" },
  { id: 'M480', area: 'glue/drain', desc: "同一個請求內每次 notifyMail 都各自清理一次（核准會清理兩次）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "if (!req._mailDrain) req._mailDrain = dispatcher.drainDue(ROUTE_DRAIN);", replace: "req._mailDrain = dispatcher.drainDue(ROUTE_DRAIN);" },
  { id: 'M481', area: 'glue/snap', desc: "snapApproval 忘了目前關卡的收件人（撤回時目前輪到的人不會收到「請勿簽核」）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "const current = curStep ? H.stepRecipients(curStep, ctx).filter(Boolean).map((un) => ({ username: un, tier: curStep.tier })) : [];", replace: "const current = [];" },
  { id: 'M482', area: 'glue/never-throw', desc: "notifyMail 的例外往外丟（寄信模組出錯會讓簽核動作失敗）", tests: GLUE,
    file: 'lib/mail/quoteMail.js', find: "warn('notifyMail', e);", replace: "throw e;" },
  { id: 'M483', area: 'glue/cron', desc: "Cron 的 secret 比對永遠通過", tests: GLUE,
    file: 'lib/mail/routes.js', find: "return crypto.timingSafeEqual(a, b);", replace: "return true;" },
  { id: 'M484', area: 'glue/cron', desc: "CRON_SECRET 未設時不再一律拒絕（雜湊 undefined 會 throw 而不是 401）", tests: GLUE,
    file: 'lib/mail/routes.js', find: "if (typeof secret !== 'string' || secret === '') return false;", replace: "/* mutated */" },
  { id: 'M485', area: 'glue/cron', desc: "Cron 摘要附上略過的收件人帳號", tests: GLUE,
    file: 'lib/mail/routes.js', find: "skipped: Array.isArray(r.skipped) ? r.skipped.length : 0,", replace: "skipped: Array.isArray(r.skipped) ? r.skipped.length : 0, skippedUsers: r.skipped," },
  { id: 'M486', area: 'glue/cron', desc: "MAIL_MODE=off 時 Cron 仍然 drain 與 purge（會碰儲存體）", tests: GLUE,
    file: 'lib/mail/routes.js', find: "if (!quoteMail.enabled()) {", replace: "if (false) {" },
  { id: 'M487', area: 'glue/cron', desc: "Cron 的 purge 失敗讓整個清理回 500", tests: GLUE,
    file: 'lib/mail/routes.js', find: "else if (pr.error) console.warn('[mail cron] purge failed:', pr.error && pr.error.message);", replace: "else if (pr.error) throw pr.error;" },
  { id: 'M488', area: 'glue/import', desc: "Email 批次匯入套用時換掉整個 users 陣列（改密碼路由持有的舊參照落在被丟棄的物件上）", tests: GLUE,
    file: 'lib/mail/routes.js', find: "plan.updated.forEach((u) => { const target = byName.get(u.username); if (target) target.email = u.email; });", replace: "auth.users = plan.users;" },
  { id: 'M489', area: 'glue/import', desc: "Email 批次匯入的稽核寫入完整位址", tests: GLUE,
    file: 'lib/mail/routes.js', find: "const names = plan.updated.slice(0, 20).map((u) => safeText(u.username, 40)).join('、');", replace: "const names = plan.updated.slice(0, 20).map((u) => safeText(u.email, 80)).join('、');" },
  { id: 'M490', area: 'glue/requeue', desc: "requeue 不驗證 id 格式", tests: GLUE,
    file: 'lib/mail/routes.js', find: "if (!ID_RE.test(id)) return res.status(400).json({ error: '寄信紀錄代碼格式不正確', code: 'BAD_ID' });", replace: "/* mutated */" },
  { id: 'M491', area: 'glue/requeue', desc: "requeue 之後不立即清理（「立即重送」不立即）", tests: GLUE,
    file: 'lib/mail/routes.js', find: "if (quoteMail.enabled()) drain = summaryCounts(await quoteMail.drainDue(REQUEUE_DRAIN));", replace: "drain = null;" },
  { id: 'M492', area: 'glue/config', desc: "後台設定頁洩漏 redirectTo 原值", tests: GLUE,
    file: 'lib/mail/routes.js', find: "config: pub,", replace: "config: Object.assign({}, pub, { redirectTo: config.redirectTo })," },
  { id: 'M493', area: 'glue/admin', desc: "寄件匣列表拿掉 requireAdmin", tests: GLUE,
    file: 'lib/mail/routes.js', find: "app.get('/api/admin/mail/outbox', requireAdmin, async (req, res) => {", replace: "app.get('/api/admin/mail/outbox', async (req, res) => {" },
  // ── FIX-1（2026-10-08 業主決定：秘書與董事會代核人的信不顯示專案名稱）──
  { id: 'M500', area: 'visibility/project', desc: "秘書的 project 旗標改回 true（信帶專案名稱）— 由 core 的可見性矩陣擋", tests: ["scripts/check-mail-core.js"],
    edits: [{"file":"lib/mail/visibility.js","find":"['secretary',  { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: false,","replace":"['secretary',  { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: true, "}] },
  { id: 'M501', area: 'visibility/project', desc: "秘書的 project 旗標改回 true — 由 render 的矩陣與哨兵字串擋", tests: ["scripts/check-mail-render.js"],
    edits: [{"file":"lib/mail/visibility.js","find":"['secretary',  { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: false,","replace":"['secretary',  { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: true, "}] },
  { id: 'M502', area: 'visibility/project', desc: "秘書的 project 旗標改回 true — 由整合測試的信件內容擋", tests: ["scripts/check-mail-integration.js"],
    edits: [{"file":"lib/mail/visibility.js","find":"['secretary',  { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: false,","replace":"['secretary',  { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: true, "}] },
  { id: 'M503', area: 'visibility/project', desc: "秘書的 project 旗標改回 true — 由膠水測試（真 quoteMail＋真 render）擋", tests: ["scripts/check-mail-glue.js"],
    edits: [{"file":"lib/mail/visibility.js","find":"['secretary',  { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: false,","replace":"['secretary',  { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: true, "}] },
  { id: 'M504', area: 'visibility/project', desc: "董事會代核人的 project 旗標改回 true — core", tests: ["scripts/check-mail-core.js"],
    edits: [{"file":"lib/mail/visibility.js","find":"['boardProxy', { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: false,","replace":"['boardProxy', { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: true, "}] },
  { id: 'M505', area: 'visibility/project', desc: "董事會代核人的 project 旗標改回 true — render", tests: ["scripts/check-mail-render.js"],
    edits: [{"file":"lib/mail/visibility.js","find":"['boardProxy', { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: false,","replace":"['boardProxy', { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: true, "}] },
  { id: 'M506', area: 'visibility/project', desc: "董事會代核人的 project 旗標改回 true — 膠水測試", tests: ["scripts/check-mail-glue.js"],
    edits: [{"file":"lib/mail/visibility.js","find":"['boardProxy', { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: false,","replace":"['boardProxy', { amount: true,  margin: true,  tier: true,  customer: false, owner: false, items: false, project: true, "}] },
  { id: 'M507', area: 'visibility/project', desc: "未知 kind 的 project 變成 true（最小資料原則失效）", tests: ["scripts/check-mail-core.js"],
    edits: [{"file":"lib/mail/visibility.js","find":"amount: false, margin: false, tier: false, customer: false, owner: false, items: false, project: false,\n      itemPrices: false, reason: UNKNOWN_REASON,","replace":"amount: false, margin: false, tier: false, customer: false, owner: false, items: false, project: true,\n      itemPrices: false, reason: UNKNOWN_REASON,"}] },
  { id: 'M508', area: 'visibility/project', desc: "visibilityFor 不回傳 project 欄位（render 讀到 undefined）", tests: ["scripts/check-mail-core.js"],
    edits: [{"file":"lib/mail/visibility.js","find":"    project: row.project,\n","replace":""}] },
  { id: 'M509', area: 'render/project', desc: "model 的專案名稱不看 visibility（秘書信的「專案名稱」列、preheader 都帶出專案名稱）", tests: ["scripts/check-mail-render.js"],
    edits: [{"file":"lib/mail/render.js","find":"const project = vis.project ? safeText(ev.projectName, 120) : '';","replace":"const project = safeText(ev.projectName, 120);"}] },
  { id: 'M510', area: 'render/project', desc: "model 的專案名稱不看 visibility — 由整合測試擋", tests: ["scripts/check-mail-integration.js"],
    edits: [{"file":"lib/mail/render.js","find":"const project = vis.project ? safeText(ev.projectName, 120) : '';","replace":"const project = safeText(ev.projectName, 120);"}] },
  { id: 'M511', area: 'render/project', desc: "主旨不看 visibility（秘書信的主旨帶出專案名稱）", tests: ["scripts/check-mail-render.js"],
    edits: [{"file":"lib/mail/render.js","find":"title: subjectLine(ev, vis.project === true),","replace":"title: subjectLine(ev, true),"}] },
  { id: 'M512', area: 'render/project', desc: "主旨不看 visibility — 由膠水測試擋", tests: ["scripts/check-mail-glue.js"],
    edits: [{"file":"lib/mail/render.js","find":"title: subjectLine(ev, vis.project === true),","replace":"title: subjectLine(ev, true),"}] },
  { id: 'M513', area: 'render/project', desc: "subjectFor(ev, kind) 忽略 kind（沒給 kind 也放專案名稱）", tests: ["scripts/check-mail-render.js"],
    edits: [{"file":"lib/mail/render.js","find":"return subjectLine(ev, visibilityFor(kind).project === true);","replace":"return subjectLine(ev, true);"}] },
  { id: 'M514', area: 'render/project', desc: "model 與主旨兩處旗標一起拿掉 — 由膠水測試擋", tests: ["scripts/check-mail-glue.js"],
    edits: [{"file":"lib/mail/render.js","find":"const project = vis.project ? safeText(ev.projectName, 120) : '';","replace":"const project = safeText(ev.projectName, 120);"},{"file":"lib/mail/render.js","find":"title: subjectLine(ev, vis.project === true),","replace":"title: subjectLine(ev, true),"}] },
  // ── FIX-4（登出殘留：深層連結暫存）──
  { id: 'M570', area: 'client/deep-link', desc: "登入頁載入時沒有上膛也沒有片段 → 不清舊暫存（下一位登入的人被帶去上一位的單據）", tests: ['scripts/check-mail-core.js'],
    edits: [{"file":"_client/deep-link.js","find":"return s ? (clear(), 'cleared') : 'none';","replace":"return s ? 'kept' : 'none';"}] },
  { id: 'M571', area: 'client/deep-link', desc: "clear() 不清上膛旗標", tests: ['scripts/check-mail-core.js'],
    edits: [{"file":"_client/deep-link.js","find":"try { s.removeItem(ARMED_KEY); } catch (e) { ok = false; }","replace":"/* mutated */"}] },
  { id: 'M572', area: 'client/deep-link', desc: "consume() 不清上膛旗標（登入成功後旗標殘留）", tests: ['scripts/check-mail-core.js'],
    edits: [{"file":"_client/deep-link.js","find":"try { s.removeItem(ARMED_KEY); } catch (e) { /* 同上 */ }","replace":"/* mutated */"}] },
  { id: 'M573', area: 'client/deep-link', desc: "rememberForLogin 不上膛（401 分支的暫存到了登入頁被當成「主動打開」而清掉）", tests: ['scripts/check-mail-core.js'],
    edits: [{"file":"_client/deep-link.js","find":"try { if (s) s.setItem(ARMED_KEY, '1'); }","replace":"try { if (s) s.setItem(ARMED_KEY, '0'); }"}] },
  { id: 'M574', area: 'client/deep-link', desc: "settleOnLoginPage 不取走上膛旗標（不是一次性：之後每次打開登入頁都沿用）", tests: ['scripts/check-mail-core.js'],
    edits: [{"file":"_client/deep-link.js","find":"try { s.removeItem(ARMED_KEY); } catch (e) { /* 一次性旗標清不掉就算了 */ }","replace":"/* mutated */"}] },
  { id: 'M575', area: 'client/deep-link', desc: "settleOnLoginPage 不看上膛旗標（401 分支的暫存一律被清掉）", tests: ['scripts/check-mail-core.js'],
    edits: [{"file":"_client/deep-link.js","find":"if (armed) {","replace":"if (false) {"}] },
  // ── FIX-7（測試與文件的 email 位址衛生）──
  { id: 'M560', area: 'hygiene/email', desc: '測試檔裡混進人名樣式位址（人名＋真實網域）', tests: ['scripts/check-mail-core.js'],
    edits: [{ file: 'scripts/check-mail-core.js', find: 'const ELLIPSIS = cp(0x2026);', replace: 'const ELLIPSIS = cp(0x2026); // planted: ' + plant('car' + 'ol', 'itts.com.tw') }] },
  { id: 'M561', area: 'hygiene/email', desc: 'README 裡混進人名樣式位址', tests: ['scripts/check-mail-core.js'],
    edits: [{ file: 'lib/mail/README.md', find: '範例位址一律是 `user1@example.test` 這類合成值', replace: '範例位址一律是 `user1@example.test` 這類合成值（例如 ' + plant('ma' + 'ry', 'itts.com.tw') + '）' }] },
  { id: 'M562', area: 'hygiene/email', desc: 'lib/mail 原始碼註解裡混進人名樣式位址', tests: ['scripts/check-mail-core.js'],
    edits: [{ file: 'lib/mail/safety.js', find: "'use strict';", replace: "'use strict';\n// planted: " + plant('da' + 've', 'gmail.com') }] },
  // ── T3（位址衛生掃描器收緊：單字母前綴只接數字、角色名只認列舉的詞）──
  { id: 'M563', area: 'hygiene/scanner', desc: '把 ROLE_RE 放寬回原樣（單字母前綴後可再接字母：ed、tj、mo、cy 配真實網域被放行）', tests: ['scripts/check-mail-core.js'],
    edits: [{ file: 'scripts/check-mail-core.js', find: 'const ROLE_RE = /^(?:(?:mgr|gm|sec|cons|proxy|chair|sales|prx|lead)(?:[0-9]+[a-z]?)?|[mgutceps][0-9]+)$/;', replace: 'const ROLE_RE = /^(mgr|gm|sec|cons|proxy|chair|sales|prx|lead|m|g|u|t|c|e|p|s)[0-9]*[a-z]?$/;' }] },
  { id: 'M564', area: 'hygiene/scanner', desc: '單字母前綴後面允許接字母（m1x、ed 這類像縮寫的 local part 被放行）', tests: ['scripts/check-mail-core.js'],
    edits: [{ file: 'scripts/check-mail-core.js', find: '|[mgutceps][0-9]+)$/;', replace: '|[mgutceps][0-9a-z]*)$/;' }] },
  { id: 'M565', area: 'hygiene/scanner', desc: '角色名後面允許接任意字母（mgrx、salesx 被放行）', tests: ['scripts/check-mail-core.js'],
    edits: [{ file: 'scripts/check-mail-core.js', find: '(?:[0-9]+[a-z]?)?|[mgutceps]', replace: '[0-9a-z]*|[mgutceps]' }] },
  { id: 'M566', area: 'hygiene/scanner', desc: 'ROLE_RE 沒有結尾錨定（以列舉的角色名開頭的任何字串都被放行）', tests: ['scripts/check-mail-core.js'],
    edits: [{ file: 'scripts/check-mail-core.js', find: '|[mgutceps][0-9]+)$/;', replace: '|[mgutceps][0-9]+)/;' }] },
  // ── FIX-2／3／6（過期判斷、時限、改派去重）──
  { id: 'M520', area: 'validity/E4', desc: "E4 駁回信不檢查 state=returned（業務重新送簽後仍寄「已被駁回」）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"if (ap.state !== 'returned') return no('NOT_RETURNED');","replace":"/* mutated */"}] },
  { id: 'M521', area: 'validity/E4', desc: "E4 本關通過信不檢查單據仍在簽核中（全部簽完／被駁回／撤回之後仍寄「將送往下一關」）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"if (ap.state !== 'pending') return no('NOT_PENDING');\n        if (!isObj(step) || step.status !== 'approved'","replace":"if (!isObj(step) || step.status !== 'approved'"}] },
  { id: 'M522', area: 'validity/E4', desc: "E4 最終核准信不檢查 state=approved（作廢後仍寄「已完成核准」）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"if (ap.state !== 'approved') return no('NOT_APPROVED');\n        return isObj(step)","replace":"return isObj(step)"}] },
  { id: 'M523', area: 'validity/E5', desc: "E5 不比對 filledAt（顧問重填一次，舊的「成本已完成」信仍寄）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"return cf.state === 'filled' && typeof cf.filledAt === 'string' && cf.filledAt === m[1] ? yes() : no('COST_STATE');","replace":"return cf.state === 'filled' ? yes() : no('COST_STATE');"}] },
  { id: 'M524', area: 'validity/E6', desc: "E6 撤回信不比對被撤回的那一次送簽（重新送簽後仍寄「請勿簽核」）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"if (m[2] === 'withdrawn') return typeof ap.submittedAt === 'string' && ap.submittedAt === m[1] ? yes() : no('ROUND_CHANGED');","replace":"if (m[2] === 'withdrawn') return yes();"}] },
  { id: 'M525', area: 'validity/E6', desc: "E6 作廢信不檢查 submittedAt 為空（重新送簽後仍寄「核准已作廢」）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"return ap.submittedAt === null || ap.submittedAt === undefined || ap.submittedAt === '' ? yes() : no('ROUND_CHANGED');","replace":"return yes();"}] },
  { id: 'M526', area: 'validity/E4', desc: "E4 不比對送簽輪次（S 不符的舊信照寄）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"if (typeof ap.submittedAt !== 'string' || m[1] !== ap.submittedAt) return no('ROUND_CHANGED');\n      if (kind === 'approved') {","replace":"if (kind === 'approved') {"}] },
  { id: 'M527', area: 'validity/E4', desc: "E4 不檢查收件人仍是單據的業務", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"if (!q.owner || q.owner !== to) return no('NOT_OWNER');\n      const kind = m[2];","replace":"const kind = m[2];"}] },
  { id: 'M528', area: 'validity/E1', desc: "E1 不比對改派標記（改派前的舊 E1 在改派後仍照寄）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"if (job.type === 'E1_SUBMIT' && (m[3] || '') !== reassignEpoch(ap)) return no('REASSIGNED');","replace":"/* mutated */"}] },
  { id: 'M529', area: 'glue/validity', desc: "quoteMail.isStillValid 對 E4／E5／E6 一律放行（FIX-2 之前的行為）", tests: GLUE,
    edits: [{"file":"lib/mail/quoteMail.js","find":"return checkJob(job).valid === true;","replace":"return /^E[456]_/.test(job.type) ? true : checkJob(job).valid === true;"}] },
  { id: 'M530', area: 'validity/E1', desc: "e1StepKey 不帶改派標記（A→B→A 時 A 的第二封 E1 被去重擋掉）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"return String(ap.submittedAt) + '#' + idx + (epoch ? '@' + epoch : '');","replace":"return String(ap.submittedAt) + '#' + idx;"}] },
  { id: 'M531', area: 'validity/E1', desc: "reassignEpoch 不在 SUBMIT 停下（上一輪的改派算進新一輪）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"if (h.action === 'SUBMIT') return '';","replace":"/* mutated */"}] },
  { id: 'M532', area: 'glue/E1', desc: "送簽的 E1 不經 e1StepKey（改派不帶標記）", tests: GLUE,
    edits: [{"file":"lib/mail/quoteMail.js","find":"stepKey: spec.type === 'E1_SUBMIT' ? e1StepKey(ap, idx) : ap.submittedAt + '#' + idx","replace":"stepKey: ap.submittedAt + '#' + idx"}] },
  // ── T1（改派給目前的承辦人本人不重複寄）──
  { id: 'M533', area: 'validity/E1', desc: "同人改派也算新的改派（isNoopReassign 永遠 false：A→A 產生新的去重鍵，多寄一封 E1，還在重試的舊 E1 被取消）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"return !!m && typeof m.from === 'string' && m.from !== '' && m.from === m.to;","replace":"return false;"}] },
  { id: 'M534', area: 'validity/E1', desc: "isNoopReassign 不檢查 from 是非空字串（meta 缺欄位／空字串的舊紀錄被誤認為同人改派，真的換人時漏寄 E1）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"return !!m && typeof m.from === 'string' && m.from !== '' && m.from === m.to;","replace":"return !!m && m.from === m.to;"}] },
  { id: 'M535', area: 'validity/E1', desc: "reassignEpoch 遇到同人改派就停下回傳空字串（而不是略過它往前找：A→B→B 時 B 的標記掉回無標記）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"if (isNoopReassign(h)) continue;","replace":"if (isNoopReassign(h)) return '';"}] },
  { id: 'M536', area: 'validity/E1', desc: "reassignEpoch 完全不略過同人改派（等同改派功能最初的行為）", tests: GLUE,
    edits: [{"file":"lib/mail/validity.js","find":"if (isNoopReassign(h)) continue;","replace":"/* mutated */"}] },
  { id: 'M540', area: 'glue/deadline', desc: "notifyMail 沒有整體時限（outbox 卡住時簽核回應無限期等下去）", tests: GLUE,
    edits: [{"file":"lib/mail/quoteMail.js","find":"const p = deadline(work, lim.notifyMs).then((r) => {","replace":"const p = deadline(work, 1e9).then((r) => {"}] },
  { id: 'M541', area: 'glue/deadline', desc: "waitPending 沒有外層時限", tests: GLUE,
    edits: [{"file":"lib/mail/quoteMail.js","find":"return deadline(Promise.allSettled(list), lim.waitMs).then((r) => {","replace":"return deadline(Promise.allSettled(list), 1e9).then((r) => {"}] },
  { id: 'M542', area: 'glue/deadline', desc: "drainDue（Cron／後台重送）沒有整體時限", tests: GLUE,
    edits: [{"file":"lib/mail/quoteMail.js","find":"const ms = Math.min(lim.drainCapMs, budget + totalMs + 1000);","replace":"const ms = 1e9;"}] },
  { id: 'M543', area: 'glue/deadline', desc: "pollDrain 之後的 db.flush 沒有時限", tests: GLUE,
    edits: [{"file":"lib/mail/quoteMail.js","find":"const fr = await deadline(Promise.resolve().then(() => d.db.flush()), lim.flushMs);","replace":"const fr = { timedOut: false }; await d.db.flush();"}] },
  { id: 'M544', area: 'glue/deadline', desc: "deadline() 完成後不清除計時器", tests: GLUE,
    edits: [{"file":"lib/mail/quoteMail.js","find":"if (timer) clearTimeout(timer); resolve(r); };","replace":"resolve(r); };"}] },
  { id: 'M545', area: 'glue/deadline', desc: "deadline() 不接住被丟下的 promise 之後的 rejection（unhandled rejection）", tests: GLUE,
    edits: [{"file":"lib/mail/quoteMail.js","find":"Promise.resolve(p).then((value) => finish({ timedOut: false, value }), (error) => finish({ timedOut: false, error }));","replace":"Promise.resolve(p).then((value) => finish({ timedOut: false, value }));"}] },
  { id: 'M546', area: 'glue/deadline', desc: "Cron 的 purge 沒有時限", tests: GLUE,
    edits: [{"file":"lib/mail/routes.js","find":"const pr = await deadline(Promise.resolve().then(() => outbox.purge()), CRON_STEP_MS);","replace":"const pr = { timedOut: false, value: await outbox.purge() };"}] },
  { id: 'M547', area: 'glue/deadline', desc: "Cron 的 db.flush 沒有時限", tests: GLUE,
    edits: [{"file":"lib/mail/routes.js","find":"await deadline(Promise.resolve().then(() => db.flush()), CRON_STEP_MS);","replace":"await db.flush();"}] },
  { id: 'M548', area: 'glue/deadline', desc: "drainDue 逾時時不標 timedOut（Cron 看不出來）", tests: GLUE,
    edits: [{"file":"lib/mail/quoteMail.js","find":"return { queued: 0, sent: 0, skipped: [], failed: [], cancelled: 0, errors: 1, timedOut: !!r.timedOut };","replace":"return { queued: 0, sent: 0, skipped: [], failed: [], cancelled: 0, errors: 1 };"}] },
];

// ── 執行────────────────────────────────────────────────────────────────────
// fs.cpSync 在含中文路徑的 Windows 環境會讓 Node 無聲當掉（exit 127），所以自己遞迴複製
function copyRecursive(from, to) {
  const st = fs.statSync(from);
  if (st.isDirectory()) {
    fs.mkdirSync(to, { recursive: true });
    fs.readdirSync(from).forEach((n) => copyRecursive(path.join(from, n), path.join(to, n)));
  } else {
    fs.mkdirSync(path.dirname(to), { recursive: true });
    fs.copyFileSync(from, to);
  }
}

function copyTree(dst) {
  const rels = ['lib/mail'];
  const sd = path.join(ROOT, 'scripts');
  fs.readdirSync(sd).forEach((n) => { if (/^check-mail-.*\.js$|^mail-preview\.js$/.test(n)) rels.push('scripts/' + n); });
  if (fs.existsSync(path.join(ROOT, '_client', 'deep-link.js'))) rels.push('_client/deep-link.js');
  if (fs.existsSync(path.join(ROOT, 'lib', 'quoteItems.js'))) rels.push('lib/quoteItems.js');       // quoteMail.js 的相依（純函式、零相依）
  rels.forEach((rel) => {
    copyRecursive(path.join(ROOT, rel), path.join(dst, rel));
  });
}

function countOccurrences(hay, needle) {
  if (!needle) return 0;
  let n = 0;
  let i = 0;
  while ((i = hay.indexOf(needle, i)) >= 0) { n++; i += needle.length; }
  return n;
}

function runTest(dir, rel) {
  const r = spawnSync(process.execPath, [path.join(dir, rel)], { cwd: dir, encoding: 'utf8', timeout: 180000, env: Object.assign({}, process.env, { MAIL_CORE_ROOT: '' }) });
  const out = (r.stdout || '') + (r.stderr || '');
  const failedLines = out.split('\n').filter((l) => /^\s+✗ /.test(l));
  const m = out.match(/(FAILED|PASSED)：(\d+) \/ (\d+)/);
  return { code: r.status === null ? 1 : r.status, out, failedLines, summary: m ? m[0] : '(無摘要：腳本可能當掉)' };
}

function main() {
  const args = process.argv.slice(2);
  const onlyArg = args.indexOf('--only');
  const only = onlyArg >= 0 ? new Set((args[onlyArg + 1] || '').split(',').filter(Boolean)) : null;
  if (args.includes('--list')) {
    MUTATIONS.forEach((m) => console.log(m.id + '  [' + m.area + '] ' + m.desc));
    console.log('共 ' + MUTATIONS.length + ' 個變異');
    return 0;
  }
  const ids = new Set();
  MUTATIONS.forEach((m) => { if (ids.has(m.id)) { console.log('編號重複：' + m.id); process.exitCode = 1; } ids.add(m.id); });

  const tmpRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'mail-mut-'));
  let bad = 0;
  try {
    // 基準：未破壞的副本必須全過
    const baseDir = path.join(tmpRoot, 'base');
    copyTree(baseDir);
    const tests = new Set(DEFAULT_TESTS);
    MUTATIONS.forEach((m) => (m.tests || DEFAULT_TESTS).forEach((x) => tests.add(x)));
    tests.forEach((rel) => {
      if (!fs.existsSync(path.join(baseDir, rel))) { console.log('基準  ✗ 找不到測試腳本 ' + rel); bad++; return; }
      const r = runTest(baseDir, rel);
      console.log('基準  ' + (r.code === 0 ? '✓' : '✗') + ' ' + rel + '  ' + r.summary);
      if (r.code !== 0) bad++;
    });
    if (bad) { console.log('基準測試沒有全過，變異結果無意義，停止。'); return 1; }

    let killed = 0;
    let survived = 0;
    let notApplied = 0;
    let ran = 0;
    MUTATIONS.forEach((m) => {
      if (only && !only.has(m.id)) return;
      ran++;
      const dir = path.join(tmpRoot, m.id);
      copyTree(dir);
      // 一個變異可以是單一處（file／find／replace），也可以是同時改多處的 edits:[{file?, find, replace}]（用來破壞多層防禦）
      const edits = (m.edits || [{ file: m.file, find: m.find, replace: m.replace }]).map((e) => ({ file: e.file || m.file, find: e.find, replace: e.replace }));
      let bad = null;
      for (const e of edits) {
        const n = countOccurrences(fs.readFileSync(path.join(dir, e.file), 'utf8'), e.find);
        if (n !== 1) { bad = '（' + e.file + ' 的 find 出現 ' + n + ' 次，必須恰好 1 次）'; break; }
      }
      if (bad) {
        notApplied++;
        console.log('NOT-APPLIED ' + m.id + ' [' + m.area + '] ' + m.desc + '  ' + bad);
        return;
      }
      for (const e of edits) {
        const f = path.join(dir, e.file);
        fs.writeFileSync(f, fs.readFileSync(f, 'utf8').replace(e.find, () => e.replace));
      }
      const results = (m.tests || DEFAULT_TESTS).map((rel) => ({ rel, r: runTest(dir, rel) }));
      const failing = results.filter((x) => x.r.code !== 0);
      if (failing.length) {
        killed++;
        const first = failing[0];
        const firstLine = (first.r.failedLines[0] || first.r.summary).trim().slice(0, 150);
        console.log('KILLED   ' + m.id + ' [' + m.area + '] ' + m.desc + '\n           → ' + first.r.summary + '；首項：' + firstLine);
      } else {
        survived++;
        console.log('SURVIVED ' + m.id + ' [' + m.area + '] ' + m.desc + '  ← 測試仍全過，這道護欄沒有被測試保護');
      }
    });
    console.log('\n變異測試：共 ' + ran + ' 個；被殺死 ' + killed + '；倖存 ' + survived + '；未套用 ' + notApplied);
    return survived || notApplied ? 1 : 0;
  } finally {
    try { fs.rmSync(tmpRoot, { recursive: true, force: true }); } catch (e) { /* 暫存資料夾清不掉就算了 */ }
  }
}

// 被 require 時（例如預先驗證每個變異的 find 是否恰好出現一次）只匯出清單，不執行
if (require.main === module) process.exit(main());
else module.exports = { MUTATIONS, countOccurrences };
