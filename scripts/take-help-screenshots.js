/**
 * 自動截圖腳本 v3 — 依角色分別截取各頁面
 * 輸出目錄：_client/help-img/{user|manager|secretary|admin}/
 */
// 注意：puppeteer 沒有宣告在 package.json（不是專案依賴，也不會隨部署安裝）；要跑本腳本請先在本機 `npm i --no-save puppeteer`。
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

const BASE_URL = process.env.HELP_BASE_URL || 'http://localhost:3000';
const OUT_BASE = path.join(__dirname, '../_client/help-img');

// 帳密一律從環境變數讀取（此檔案在公開 repo；先前寫死的帳密已移除，那些密碼必須視為外洩並輪替）。
// 每個角色一組：HELP_<ROLE>_USER／HELP_<ROLE>_PASS，ROLE = USER｜MANAGER｜SECRETARY｜ADMIN。
// 沒設定的角色會被略過（並在開頭列出）。範例（PowerShell）：
//   $env:HELP_USER_USER='…'; $env:HELP_USER_PASS='…'; node scripts/take-help-screenshots.js
const roleCred = (r) => ({ username: process.env['HELP_' + r.toUpperCase() + '_USER'], password: process.env['HELP_' + r.toUpperCase() + '_PASS'] });
const CREDS = {
  user:      roleCred('user'),
  manager:   roleCred('manager'),
  secretary: roleCred('secretary'),
  admin:     roleCred('admin'),
};

// 各角色要截的頁面
const ROLE_SHOTS = {
  user: [
    { id: 'interface',   navId: 'navContacts' },
    { id: 'contacts',    navId: 'navContacts' },
    { id: 'visits',      navId: 'navVisits' },
    { id: 'pipeline',    navId: 'navPipeline' },
    { id: 'erp-ma',      navId: 'navErpMa' },
    { id: 'sap-ma',      navId: 'navSapMa' },
    { id: 'receivables', navId: 'navReceivables' },
    { id: 'quotations',  navId: 'navQuotations' },
    { id: 'targets',     navId: 'navTargets' },
  ],
  manager: [
    { id: 'interface',      navId: 'navManagerHome' },
    { id: 'contacts',       navId: 'navContacts' },
    { id: 'visits',         navId: 'navVisits' },
    { id: 'pipeline',       navId: 'navPipeline' },
    { id: 'forecast',       navId: 'navForecast' },
    { id: 'targets',        navId: 'navTargets' },
    { id: 'lostOpp',        navId: 'navLostOpp' },
    { id: 'pipelineReport', navId: 'navPipelineReport' },
    { id: 'managerHome',    navId: 'navManagerHome' },
    { id: 'erp-ma',         navId: 'navErpMa' },
    { id: 'sap-ma',         navId: 'navSapMa' },
    { id: 'receivables',    navId: 'navReceivables' },
    { id: 'quotations',     navId: 'navQuotations' },
  ],
  secretary: [
    { id: 'interface',      navId: 'navForecast' },
    { id: 'forecast',       navId: 'navForecast' },
    { id: 'erp-ma',         navId: 'navErpMa' },
    { id: 'sap-ma',         navId: 'navSapMa' },
    { id: 'receivables',    navId: 'navReceivables' },
    { id: 'quotations',     navId: 'navQuotations' },
    { id: 'pipelineReport', navId: 'navPipelineReport' },
  ],
  admin: [
    { id: 'interface',      navId: 'navContacts' },
    { id: 'contacts',       navId: 'navContacts' },
    { id: 'visits',         navId: 'navVisits' },
    { id: 'pipeline',       navId: 'navPipeline' },
    { id: 'forecast',       navId: 'navForecast' },
    { id: 'targets',        navId: 'navTargets' },
    { id: 'lostOpp',        navId: 'navLostOpp' },
    { id: 'quotations',     navId: 'navQuotations' },
    { id: 'receivables',    navId: 'navReceivables' },
    { id: 'erp-ma',         navId: 'navErpMa' },
    { id: 'sap-ma',         navId: 'navSapMa' },
    { id: 'campaigns',      navId: 'navCampaigns' },
    { id: 'leads',          navId: 'navLeads' },
    { id: 'pipelineReport', navId: 'navPipelineReport' },
    { id: 'managerHome',    navId: 'navManagerHome' },
    { id: 'admin-users',    url: '/admin.html', adminSec: 'users' },
    { id: 'admin-logs',     url: '/admin.html', adminSec: 'logs' },
    { id: 'admin-api',      url: '/admin.html', adminSec: 'api-stats' },
  ],
};

async function wait(ms) { return new Promise(r => setTimeout(r, ms)); }

async function doLogin(page, creds) {
  console.log(`  → 前往登入頁...`);
  await page.goto(`${BASE_URL}/login.html`, { waitUntil: 'domcontentloaded', timeout: 15000 });
  await wait(1000);
  await page.evaluate(() => {
    document.getElementById('username').value = '';
    document.getElementById('password').value = '';
  });
  await page.type('#username', creds.username, { delay: 60 });
  await page.type('#password', creds.password, { delay: 60 });
  await wait(300);
  await page.click('.login-btn');
  console.log(`  → 等待跳轉...`);
  try {
    await page.waitForFunction(
      () => !window.location.pathname.includes('login'),
      { timeout: 10000 }
    );
  } catch(e) {
    const url = page.url();
    if (url.includes('login')) {
      console.error(`  ❌ 登入失敗（仍在 ${url}）`);
      return false;
    }
  }
  await wait(2000);
  console.log(`  ✅ 登入成功：${page.url()}`);
  return true;
}

async function ensureMainApp(page) {
  const url = page.url();
  if (url.includes('login') || url.includes('admin.html')) {
    await page.goto(`${BASE_URL}/`, { waitUntil: 'networkidle2' });
    await wait(2500);
  }
}

async function clickNav(page, navId) {
  const ok = await page.evaluate((id) => {
    const el = document.getElementById(id);
    if (!el) return false;
    if (el.style.display === 'none') return 'hidden';
    el.click();
    return true;
  }, navId);
  if (ok === 'hidden') { console.warn(`  ⚠ #${navId} 不可見（此角色無此功能）`); return false; }
  if (!ok) { console.warn(`  ⚠ 找不到 #${navId}`); return false; }
  await wait(2500);
  return true;
}

async function takeShot(page, folder, id) {
  const dir = path.join(OUT_BASE, folder);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, `${id}.png`);
  await page.screenshot({ path: file, fullPage: false });
  console.log(`  📸 ${folder}/${id}.png`);
}

(async () => {
  const missing = Object.keys(CREDS).filter(r => !CREDS[r].username || !CREDS[r].password);
  if (missing.length) console.log('以下角色沒有設定帳密環境變數（HELP_<ROLE>_USER／HELP_<ROLE>_PASS），將略過：' + missing.join(', '));
  if (missing.length === Object.keys(CREDS).length) { console.error('沒有任何角色設定帳密，結束。'); process.exit(2); }

  // 登入頁（共用）
  {
    const browser = await puppeteer.launch({
      headless: true,
      executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu', '--window-size=1440,900'],
      defaultViewport: { width: 1440, height: 900 },
    });
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(20000);
    await page.goto(`${BASE_URL}/login.html`, { waitUntil: 'networkidle2' });
    await wait(1500);
    await takeShot(page, '.', 'login');  // 存到根目錄
    await browser.close();
  }

  // 各角色截圖
  for (const [roleName, shots] of Object.entries(ROLE_SHOTS)) {
    console.log(`\n╔══════════════════════╗`);
    console.log(`║  角色：${roleName.padEnd(14)}║`);
    console.log(`╚══════════════════════╝`);

    const browser = await puppeteer.launch({
      headless: true,
      executablePath: process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu', '--window-size=1440,900'],
      defaultViewport: { width: 1440, height: 900 },
    });
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(20000);

    // 登入
    if (!CREDS[roleName].username || !CREDS[roleName].password) { console.log('  略過（未設定帳密環境變數）'); await browser.close(); continue; }
    const ok = await doLogin(page, CREDS[roleName]);
    if (!ok) { await browser.close(); continue; }

    for (const shot of shots) {
      console.log(`\n  [${shot.id}]`);
      try {
        if (shot.adminSec) {
          // 後台頁面
          await page.goto(`${BASE_URL}/admin.html`, { waitUntil: 'networkidle2' });
          await wait(1500);
          await page.evaluate((sec) => {
            const el = document.querySelector(`[data-sec="${sec}"]`);
            if (el) el.click();
          }, shot.adminSec);
          await wait(2500);
          await takeShot(page, roleName, shot.id);
        } else if (shot.url) {
          // 指定 URL
          await page.goto(`${BASE_URL}${shot.url}`, { waitUntil: 'networkidle2' });
          await wait(2000);
          await takeShot(page, roleName, shot.id);
        } else {
          // 一般 nav 點擊
          await ensureMainApp(page);
          const clicked = await clickNav(page, shot.navId);
          if (clicked) await takeShot(page, roleName, shot.id);
        }
      } catch(e) {
        console.error(`  ❌ ${shot.id} 失敗：${e.message}`);
      }
    }

    await browser.close();
    console.log(`\n✅ 角色 ${roleName} 完成`);
  }

  console.log('\n🎉 全部截圖完成');
})();
