/**
 * 自動截圖腳本 v3 — 依角色分別截取各頁面
 * 輸出目錄：_client/help-img/{user|manager|secretary|admin}/
 */
const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

const BASE_URL = 'http://localhost:3000';
const OUT_BASE = path.join(__dirname, '../_client/help-img');

const CREDS = {
  user:      { username: 'Stevenlee', password: '890807Aa' },
  manager:   { username: 'Sabrina',   password: 'Sabrina1234' },
  secretary: { username: 'Debbie',    password: 'Debbie1234' },
  admin:     { username: 'admin',     password: 'Admin1234' },
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
  // 登入頁（共用）
  {
    const browser = await puppeteer.launch({
      headless: true,
      executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
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
      executablePath: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-gpu', '--window-size=1440,900'],
      defaultViewport: { width: 1440, height: 900 },
    });
    const page = await browser.newPage();
    page.setDefaultNavigationTimeout(20000);

    // 登入
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
