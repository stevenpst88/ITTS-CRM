const API = '/api';

// ── XSS 防護：HTML 字符轉義 ───────────────────────────────
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ── 安全 href：防止 javascript: scheme XSS ───────────────
function safeHref(url) {
  if (!url) return '#';
  const trimmed = String(url).trim();
  return /^https?:\/\//i.test(trimmed) ? trimmed : '#';
}

let userPermissions = { role: 'user', canDownloadContacts: false, canSetTargets: false };

async function loadPermissions() {
  try {
    const r = await fetch(`${API}/me/permissions`);
    if (r.ok) {
      userPermissions = await r.json();
      applyPermissions();
    }
  } catch {}
}

function applyPermissions() {
  // 業績目標：只有有權限才能儲存
  const saveTargetBtn = $('saveTargetBtn');
  if (saveTargetBtn) {
    if (!userPermissions.canSetTargets && userPermissions.role !== 'admin') {
      saveTargetBtn.style.display = 'none';
      const notice = document.createElement('div');
      notice.style.cssText = 'font-size:12px;color:#e65100;margin-top:8px;';
      notice.textContent = '⚠️ 您沒有設定業績目標的權限，請聯繫管理者';
      notice.id = 'noTargetPermNotice';
      if (!$('noTargetPermNotice')) saveTargetBtn.parentNode.appendChild(notice);
    } else {
      saveTargetBtn.style.display = '';
    }
  }
  // 下載客戶名單：只有有權限才顯示匯出按鈕
  const exportBtn = $('exportBtn');
  if (exportBtn) {
    exportBtn.style.display = (userPermissions.canDownloadContacts || userPermissions.role === 'admin') ? '' : 'none';
  }
  // 管理者後台入口：只有 admin 才顯示
  const adminLink = $('adminPanelLink');
  if (adminLink) adminLink.style.display = (userPermissions.role === 'admin') ? 'block' : 'none';
  // 名單移轉：manager1、manager2、admin 才顯示
  const navTransfer = $('navTransfer');
  if (navTransfer) navTransfer.style.display = ['admin','manager1','manager2'].includes(userPermissions.role) ? '' : 'none';
  // 角色導覽限制
  const role = userPermissions.role;
  if (role === 'secretary') {
    // 秘書可看首頁看板、銷售預測、帳務管理、Call-in Pass，隱藏其他功能
    ['navProspects','navContacts','navVisits','navTargets','navPipeline','navContractGroup'].forEach(id => {
      const el = $(id); if (el) el.style.display = 'none';
    });
    // navHome、navAccountingGroup、navForecast、navCallin 保留
  } else if (role === 'manager1' || role === 'manager2') {
    // 主管可以看全部，但隱藏目標設定（除非有 canSetTargets 權限）
    // 顯示所有導覽
    ['navHome','navProspects','navContacts','navVisits','navTargets','navPipeline','navContractGroup','navForecast','navLostOpp'].forEach(id => {
      const el = $(id); if (el) el.style.display = '';
    });
  }
  // 管理儀表板 + 主管首頁：主管 / admin 才顯示
  if (role === 'manager1' || role === 'manager2' || role === 'admin') {
    const e1 = $('navExecDash');     if (e1) e1.style.display = '';
    const e2 = $('navManagerHome');  if (e2) e2.style.display = '';
  }
  // 行銷功能
  if (role === 'marketing') {
    ['navProspects','navContacts','navVisits','navTargets','navPipeline',
     'navContractGroup','navForecast','navLostOpp','navExecDash','navQuotations',
     'navPipelineReport','navAccountingGroup','navCallin'].forEach(id => {
      const el = $(id); if (el) el.style.display = 'none';
    });
  }
  // 行銷活動 & Lead 管理：僅行銷人員可見
  if (role === 'marketing') {
    ['navCampaigns','navLeads'].forEach(id => { const el=$(id); if(el) el.style.display=''; });
  }
}

let allContacts = [];
let allOpportunities = [];
let currentViewId = null;

// ── 職能分類 ─────────────────────────────────────────────
const JOB_FUNCTION_CATEGORIES = [
  {
    key: 'management',
    label: '決策領導層',
    sublabel: 'Top Management',
    color: '#c0392b',
    bg: '#fce8e6',
    keywords: ['董事長','副董事長','總經理','副總','協理','總監','CEO','COO','CFO','CTO','CMO','VP','Vice President','President','Director','執行長','執行副總']
  },
  {
    key: 'operations',
    label: '營運與生產核心',
    sublabel: 'Production & Operations',
    color: '#e67e22',
    bg: '#fff3e0',
    keywords: ['廠長','生管','物管','資材','MC','課長','組長','班長','Supervisor','作業員','技術員','OP','Technician','生產','製造','倉管','物料','採購主管','廠務','現場']
  },
  {
    key: 'engineering',
    label: '研發與技術工程',
    sublabel: 'R&D & Engineering',
    color: '#1565c0',
    bg: '#e8f0fe',
    keywords: ['研發工程師','R&D','研發','製程工程師','製程','PE','設備工程師','設備','EE','工業工程師','工業工程','IE','產品經理','PM','Product Manager','軟體工程師','系統工程師','MIS','IT工程','架構師','開發工程師','韌體','硬體','機械工程師']
  },
  {
    key: 'quality',
    label: '品質控管體系',
    sublabel: 'Quality Assurance',
    color: '#1b7e34',
    bg: '#e6f4ea',
    keywords: ['品保','品管','IQC','IPQC','OQC','測試工程師','QA','QC','TE','品質','品控','稽核','驗證','認證','可靠度']
  },
  {
    key: 'admin',
    label: '行政與市場幕僚',
    sublabel: 'Administration & Sales',
    color: '#6a1b9a',
    bg: '#f3e5f5',
    keywords: ['業務','Sales','採購','Buyer','財務','會計','人力資源','HR','環安衛','ESH','行政','秘書','助理','公關','行銷','Marketing','法務','企劃','客服','業務專員','業務經理','業務副理']
  }
];

// 自動 Mapping 職稱到職能分類
function autoMapJobFunction(title) {
  if (!title) return '';
  const t = title.toLowerCase();
  for (const cat of JOB_FUNCTION_CATEGORIES) {
    if (cat.keywords.some(kw => t.includes(kw.toLowerCase()))) {
      return cat.key;
    }
  }
  return '';
}

// ── 系統選單資料 ─────────────────────────────────────────
const SYSTEMS = {
  '鼎新系統': ['Tiptop ERP', 'Workflow ERP', 'Smart ERP'],
  'SAP': ['SAP ECC', 'SAP Public Cloud', 'SAP Private Cloud', 'SAP OP S/4 HANA', 'SAP B1'],
  'Oracle': ['Fusion', 'EBS', 'NetSuite'],
  '文中': [],
  '正航': [],
  '自行開發': null,
  '其他': null
};

// ── 工具函式 ─────────────────────────────────────────────
function $(id) { return document.getElementById(id); }

function showToast(msg, duration = 2500) {
  const t = $('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), duration);
}

function getInitial(name) {
  if (!name) return '?';
  return name.charAt(0);
}

function avatarColor(name) {
  const colors = [
    ['#1a73e8','#34a853'], ['#ea4335','#fbbc04'],
    ['#9c27b0','#e91e63'], ['#00bcd4','#009688'],
    ['#ff5722','#ff9800'], ['#3f51b5','#2196f3']
  ];
  const idx = (name || '?').charCodeAt(0) % colors.length;
  return colors[idx];
}

// ── 頁面區塊切換 ─────────────────────────────────────────
let currentSection = null; // null = dashboard, 'contacts', 'visits'

// ── 色盤（與漏斗 gray/blue/orange/red 色系呼應）──────────
const PALETTE_3D = [
  '#2563eb','#e11d48','#d97706','#16a34a',
  '#7c3aed','#0891b2','#dc2626','#ca8a04',
  '#059669','#9333ea','#0284c7','#db2777',
  '#65a30d','#0d9488'
];

function hexDarken(hex, amt) {
  const r=parseInt(hex.slice(1,3),16), g=parseInt(hex.slice(3,5),16), b=parseInt(hex.slice(5,7),16);
  const f=1-amt;
  return `rgb(${~~(r*f)},${~~(g*f)},${~~(b*f)})`;
}
function hexLighten(hex, amt) {
  const r=parseInt(hex.slice(1,3),16), g=parseInt(hex.slice(3,5),16), b=parseInt(hex.slice(5,7),16);
  return `rgb(${~~Math.min(255,r+(255-r)*amt)},${~~Math.min(255,g+(255-g)*amt)},${~~Math.min(255,b+(255-b)*amt)})`;
}

// ── 3D 圓餅 SVG 生成 ──────────────────────────────────────
function build3DPieSVG(labels, data, palette) {
  const total = data.reduce((a,b)=>a+b,0);
  if (!total) return '';
  const W=270, cx=135, cy=80, Rx=100, Ry=38, depth=22;
  const H = cy + Ry + depth + 12;
  const font="-apple-system,BlinkMacSystemFont,'Microsoft JhengHei',sans-serif";
  const epx = a => cx + Rx*Math.cos(a);
  const epy = a => cy + Ry*Math.sin(a);

  let cum = -Math.PI/2;
  const slices = data.map((d,i) => {
    const s=cum, sw=(d/total)*Math.PI*2; cum+=sw;
    return { start:s, end:cum, mid:s+sw/2, d, i, c:palette[i%palette.length] };
  });

  let defs = `<defs><radialGradient id="topShine" cx="42%" cy="38%" r="55%">
    <stop offset="0%" stop-color="rgba(255,255,255,0.38)"/>
    <stop offset="100%" stop-color="rgba(255,255,255,0)"/>
  </radialGradient>`;

  let arcSvg='', radSvg='', topSvg='';

  // ① 弧牆（前半圓）
  slices.forEach(sl => {
    const cd=hexDarken(sl.c, 0.38);
    const steps=Math.max(4, Math.ceil(Math.abs(sl.end-sl.start)/(Math.PI/10)));
    for (let j=0;j<steps;j++) {
      const a1=sl.start+(sl.end-sl.start)*j/steps;
      const a2=sl.start+(sl.end-sl.start)*(j+1)/steps;
      const ma=(a1+a2)/2;
      if (Math.sin(ma) > -0.05) {
        const x1=epx(a1),y1=epy(a1),x2=epx(a2),y2=epy(a2);
        const op=(0.52+Math.max(0,Math.sin(ma))*0.36).toFixed(2);
        arcSvg+=`<path d="M${x1},${y1} L${x1},${y1+depth} L${x2},${y2+depth} L${x2},${y2} Z" fill="${cd}" opacity="${op}"/>`;
      }
    }
    // ② 徑向側面
    [sl.start, sl.end].forEach(a => {
      if (Math.sin(a) > -0.12) {
        const ax=epx(a),ay=epy(a);
        radSvg+=`<path d="M${cx},${cy} L${cx},${cy+depth} L${ax},${ay+depth} L${ax},${ay} Z" fill="${hexDarken(sl.c,0.44)}" opacity="0.72"/>`;
      }
    });
  });

  // ③ 頂面切片
  slices.forEach(sl => {
    const la=(sl.end-sl.start)>Math.PI?1:0;
    const x1=epx(sl.start),y1=epy(sl.start),x2=epx(sl.end),y2=epy(sl.end);
    const gId=`tg${sl.i}`;
    defs+=`<linearGradient id="${gId}" x1="0%" y1="0%" x2="85%" y2="100%">
      <stop offset="0%" stop-color="${hexLighten(sl.c,0.34)}"/>
      <stop offset="100%" stop-color="${sl.c}"/>
    </linearGradient>`;
    topSvg+=`<path d="M${cx},${cy} L${x1},${y1} A${Rx},${Ry} 0 ${la} 1 ${x2},${y2} Z"
      fill="url(#${gId})" stroke="rgba(255,255,255,0.55)" stroke-width="0.9"/>`;
  });

  // ④ 橢圓高光 + 邊框
  topSvg+=`<ellipse cx="${cx}" cy="${cy}" rx="${Rx}" ry="${Ry}" fill="url(#topShine)"/>
    <ellipse cx="${cx}" cy="${cy}" rx="${Rx}" ry="${Ry}" fill="none" stroke="rgba(255,255,255,0.28)" stroke-width="1.2"/>`;

  defs+='</defs>';
  return `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg"
    style="width:100%;display:block;filter:drop-shadow(0 5px 16px rgba(0,0,0,0.24))">
    ${defs}${arcSvg}${radSvg}${topSvg}</svg>`;
}

// 共用：建立圖例清單 HTML
function buildLegendHTML(labels, data, colors, unit = '家') {
  const total = data.reduce((a, b) => a + b, 0);
  return labels.map((label, i) => {
    const pct = total > 0 ? Math.round(data[i] / total * 100) : 0;
    return `<div class="chart-legend-row">
      <span class="chart-legend-dot" style="background:${colors[i]}"></span>
      <span class="chart-legend-label">${label}</span>
      <span class="chart-legend-count">${data[i]} ${unit}</span>
      <span class="chart-legend-pct">${pct}%</span>
    </div>`;
  }).join('');
}

function countOppStage(stage) {
  // 只統計商機記錄中的等級（不含名片的 opportunityStage 欄位）
  return allOpportunities.filter(o => o.stage === stage).length;
}

// ── 生日提醒 ──────────────────────────────────────────────
async function loadBirthdayReminders() {
  const card = $('birthdayReminderCard');
  if (!card) return;
  try {
    const res = await fetch(`${API}/birthday-reminders?days=3`);
    if (!res.ok) { card.style.display = 'none'; return; }
    const list = await res.json();
    if (!list.length) { card.style.display = 'none'; return; }

    // 是否被本次 session 關閉過（當天）
    const dismissedKey = 'bdayDismissed_' + new Date().toISOString().slice(0,10);
    if (sessionStorage.getItem(dismissedKey)) { card.style.display = 'none'; return; }

    const dayLabel = d => d === 0 ? '🎉 今天' : d === 1 ? '明天' : `${d} 天後`;
    const urgency  = d => d === 0 ? 'bday-today' : d === 1 ? 'bday-tomorrow' : '';

    $('bdayList').innerHTML = list.map(p => `
      <div class="bday-item ${urgency(p.daysLeft)}">
        <div class="bday-item-left">
          <span class="bday-days ${urgency(p.daysLeft)}">${dayLabel(p.daysLeft)}</span>
          <div class="bday-name">${escapeHtml(p.name)}
            ${p.nameEn ? `<span class="bday-en">${escapeHtml(p.nameEn)}</span>` : ''}
          </div>
          <div class="bday-company">${escapeHtml(p.company)}${p.title ? ' · ' + escapeHtml(p.title) : ''}</div>
        </div>
        <div class="bday-item-right">
          <span class="bday-date">🎂 ${p.personalBirthday}</span>
          <span class="bday-owner">負責：${escapeHtml(p.ownerName)}</span>
        </div>
      </div>`).join('');

    card.style.display = '';

    const btn = $('bdayDismiss');
    if (btn) btn.onclick = () => {
      sessionStorage.setItem(dismissedKey, '1');
      card.style.display = 'none';
    };
  } catch(e) { console.warn('[birthday card]', e); }
}

// ── 殭屍商機：首頁警示卡 ──────────────────────────────────
async function loadZombieAlertCard() {
  const card = $('zombieAlertCard');
  if (!card) return;
  try {
    const res = await fetch(`${API}/zombie-opportunities`);
    if (!res.ok) { card.style.display = 'none'; return; }
    allZombieOpps = await res.json();
    if (!allZombieOpps.length) { card.style.display = 'none'; return; }

    const dismissedKey = 'zombieDismissed_' + new Date().toDateString();
    if (sessionStorage.getItem(dismissedKey)) { card.style.display = 'none'; return; }

    const danger = allZombieOpps.filter(z => z.severity === 'danger').length;
    const warn   = allZombieOpps.length - danger;
    const stageCount = { A:0, B:0, C:0 };
    allZombieOpps.forEach(z => { if (stageCount[z.stage] !== undefined) stageCount[z.stage]++; });

    $('zombieAlertSub').textContent = `共 ${allZombieOpps.length} 筆商機需要關注` +
      (danger ? `，其中 ${danger} 筆情況緊急` : '');

    $('zombieAlertChips').innerHTML = [
      stageCount.A ? `<span class="z-chip z-chip-a z-chip-danger">A 級 ${stageCount.A} 筆</span>` : '',
      stageCount.B ? `<span class="z-chip z-chip-b ${stageCount.B>0?'':''}">B 級 ${stageCount.B} 筆</span>` : '',
      stageCount.C ? `<span class="z-chip z-chip-c">C 級 ${stageCount.C} 筆</span>` : '',
    ].join('');

    card.style.display = '';

    $('zombieAlertGoto').onclick = () => showSection('targets');
    $('zombieAlertDismiss').onclick = () => {
      sessionStorage.setItem(dismissedKey, '1');
      card.style.display = 'none';
    };
  } catch { if (card) card.style.display = 'none'; }
}

// ── 殭屍商機：年度目標頁詳細區塊 ───────────────────────────
function renderZombieSection() {
  const sec  = $('zombieSection');
  const tbody = $('zombieTableBody');
  const cnt   = $('zombieSectionCount');
  if (!sec || !tbody) return;

  if (!allZombieOpps.length) { sec.style.display = 'none'; return; }
  sec.style.display = '';
  cnt.textContent = `${allZombieOpps.length} 筆`;

  const STAGE_COLOR = { A:'#c5221f', B:'#e37400', C:'#1a73e8', D:'#888' };
  const VISIT_TYPE_ICONS = { '親訪':'🤝','電話':'📞','視訊':'💻','Email':'📧','展覽':'🎪' };

  tbody.innerHTML = allZombieOpps.map(z => {
    const sevCls = z.severity === 'danger' ? 'z-sev-danger' : 'z-sev-warn';
    const sevLabel = z.severity === 'danger'
      ? '<span class="z-sev-badge z-sev-danger">🔴 緊急</span>'
      : '<span class="z-sev-badge z-sev-warn">🟡 警示</span>';
    const stageStyle = `color:${STAGE_COLOR[z.stage]||'#333'};font-weight:700`;
    const lastVisitTxt = z.lastVisit
      ? `${z.lastVisit}<br><small style="color:#aaa">${z.daysSinceAny} 天前</small>`
      : '<span style="color:#e74c3c">從未</span>';
    const reasonHtml = z.reasons.map(r => `<div class="z-reason-item">• ${escapeHtml(r)}</div>`).join('');
    return `<tr class="z-row ${sevCls}" data-id="${z.id}">
      <td>${sevLabel}</td>
      <td><strong>${escapeHtml(z.company||'-')}</strong></td>
      <td>${escapeHtml(z.contactName||'-')}</td>
      <td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${z.product||''}">${escapeHtml(z.product||'-')}</td>
      <td style="text-align:right">${z.amount?Number(z.amount).toLocaleString():'-'}</td>
      <td><span style="${stageStyle}">${z.stage}</span></td>
      <td style="white-space:nowrap">${lastVisitTxt}</td>
      <td class="z-reasons">${reasonHtml}</td>
      <td>${escapeHtml(z.ownerName||z.owner)}</td>
    </tr>`;
  }).join('');

  // 收合按鈕
  const colBtn = $('zombieCollapseBtn');
  if (colBtn) {
    colBtn.onclick = () => {
      const body = $('zombieSectionBody');
      const collapsed = body.style.display === 'none';
      body.style.display = collapsed ? '' : 'none';
      colBtn.textContent = collapsed ? '▲' : '▼';
    };
  }
}

function updateStatCards() {
  const uniqueCompanies = new Set(allContacts.map(c => c.company).filter(Boolean)).size;
  $('statCompanies').textContent = uniqueCompanies;
  $('statContacts').textContent  = allContacts.length;
  $('statCommit').textContent    = countOppStage('A');
  $('statUpside').textContent    = countOppStage('B');
  $('statPipeline').textContent  = countOppStage('C');
}

function renderDashboardCharts() {
  // ── 產業圓餅圖（以公司為單位，同公司只算一次）──
  // 先建立每家公司的代表聯絡人（主要聯繫人優先，否則取第一筆）
  const companyRepMap = new Map();
  allContacts.forEach(c => {
    const key = (c.company || '').trim();
    if (!key) {
      // 無公司名稱：個人各自計入
      companyRepMap.set('__individual__' + c.id, c);
    } else if (!companyRepMap.has(key) || c.isPrimary) {
      companyRepMap.set(key, c);
    }
  });
  const industryMap = {};
  companyRepMap.forEach(c => {
    const ind = c.industry || '未分類';
    industryMap[ind] = (industryMap[ind] || 0) + 1;
  });
  // 依數量排序（多的在前）
  const iSorted = Object.entries(industryMap).sort((a, b) => b[1] - a[1]);
  const iLabels  = iSorted.map(e => e[0]);
  const iData    = iSorted.map(e => e[1]);
  const iColors = iLabels.map((_, i) => PALETTE_3D[i % PALETTE_3D.length]);

  const iEmpty  = $('industryChartEmpty');
  const iChart  = $('industryChart3D');
  const iLegend = $('industryLegend');
  if (iLabels.length === 0) {
    iChart.innerHTML = '';
    iEmpty.classList.add('visible');
    iLegend.innerHTML = '';
  } else {
    iEmpty.classList.remove('visible');
    iChart.innerHTML = build3DPieSVG(iLabels, iData, iColors);
    iLegend.innerHTML = buildLegendHTML(iLabels, iData, iColors, '家公司');
  }

  // ── 商機漏斗 SVG 3D（D頂→A底）──
  const OPP_FUNNEL = [
    { key:'D', name:'靜止中',   base:'#8d8d8d', light:'#c8c8c8', dark:'#4a4a4a', rim:'#b0b0b0' },
    { key:'C', name:'Pipeline', base:'#1a73e8', light:'#6ab0ff', dark:'#0a3d8f', rim:'#4d9fff' },
    { key:'B', name:'Upside',   base:'#f57c00', light:'#ffb74d', dark:'#8f3e00', rim:'#ffa040' },
    { key:'A', name:'Commit',   base:'#d32f2f', light:'#ff6b6b', dark:'#7f0000', rim:'#f05050' },
  ];
  // 商機漏斗來源：只使用實際商機記錄（與商機推進進度一致）
  const oppMap = {};
  allOpportunities.forEach(o => {
    if (o.stage) oppMap[o.stage] = (oppMap[o.stage] || 0) + 1;
  });
  const oppTotal = OPP_FUNNEL.reduce((s, st) => s + (oppMap[st.key] || 0), 0);

  const oEmpty  = $('opportunityChartEmpty');
  const oFunnel = $('opportunityFunnel');

  if (oppTotal === 0) {
    oFunnel.innerHTML = '';
    oEmpty.classList.add('visible');
  } else {
    oEmpty.classList.remove('visible');

    const W = 380, H = 360, cx = W / 2;
    const f  = "-apple-system,BlinkMacSystemFont,'Microsoft JhengHei',sans-serif";
    const ryF = 0.19; // 橢圓扁率（模擬透視）

    // 各邊界的 y 座標與半寬（加寬後等比放大）
    const bounds = [
      { y: 6,   hw: 168 },
      { y: 96,  hw: 124 },
      { y: 186, hw: 84  },
      { y: 276, hw: 48  },
      { y: 350, hw: 22  },
    ];

    let defs = '<defs>';
    // 共用垂直高光漸層
    defs += `<linearGradient id="shineV" x1="0%" y1="0%" x2="0%" y2="100%">
      <stop offset="0%" stop-color="white" stop-opacity="0.22"/>
      <stop offset="55%" stop-color="white" stop-opacity="0.04"/>
      <stop offset="100%" stop-color="white" stop-opacity="0"/>
    </linearGradient>`;
    // 文字陰影 filter（讓白字在淺色區段也清晰）
    defs += `<filter id="txtShadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="1" stdDeviation="2" flood-color="rgba(0,0,0,0.65)"/>
    </filter>`;

    let fills = '', rims = '', shines = '', labels = '';

    OPP_FUNNEL.forEach((s, i) => {
      const b0 = bounds[i], b1 = bounds[i + 1];
      const ry0 = b0.hw * ryF, ry1 = b1.hw * ryF;
      const midY = (b0.y + b1.y) / 2 + 2;
      const count = oppMap[s.key] || 0;
      const pct   = Math.round(count / oppTotal * 100);
      const op    = count === 0 ? 0.38 : 1;
      const gId   = `hg${s.key}`;

      // 水平漸層：左暗→中亮→右暗（製造圓柱感）
      defs += `<linearGradient id="${gId}" x1="0%" y1="0%" x2="100%" y2="0%">
        <stop offset="0%"   stop-color="${s.dark}"/>
        <stop offset="28%"  stop-color="${s.base}"/>
        <stop offset="52%"  stop-color="${s.light}"/>
        <stop offset="75%"  stop-color="${s.base}"/>
        <stop offset="100%" stop-color="${s.dark}"/>
      </linearGradient>`;

      // ① 梯形主體
      fills += `<polygon opacity="${op}"
        points="${cx-b0.hw},${b0.y} ${cx+b0.hw},${b0.y} ${cx+b1.hw},${b1.y} ${cx-b1.hw},${b1.y}"
        fill="url(#${gId})"/>`;

      // ② 高光覆蓋（頂部反光條）
      shines += `<polygon opacity="${op * 0.9}"
        points="${cx-b0.hw},${b0.y} ${cx+b0.hw},${b0.y} ${cx+b1.hw},${b1.y} ${cx-b1.hw},${b1.y}"
        fill="url(#shineV)"/>`;

      // ③ 頂部橢圓環（模擬開口深度）— 在所有梯形之後畫
      rims += `
        <ellipse cx="${cx}" cy="${b0.y}" rx="${b0.hw}" ry="${ry0}"
          fill="${s.dark}" opacity="${op * 0.55}"/>
        <ellipse cx="${cx}" cy="${b0.y}" rx="${b0.hw * 0.72}" ry="${ry0 * 0.72}"
          fill="${s.light}" opacity="${op * 0.35}"/>
        <ellipse cx="${cx}" cy="${b0.y}" rx="${b0.hw}" ry="${ry0}"
          fill="none" stroke="${s.rim}" stroke-width="1.2" opacity="${op * 0.7}"/>`;

      // ④ 文字標籤（套用 SVG filter 陰影，白字在各顏色段皆清晰）
      labels += `
        <text x="${cx}" y="${midY - 8}" text-anchor="middle" opacity="${op}"
          fill="white" font-size="19" font-weight="bold" font-family="${f}"
          filter="url(#txtShadow)">${count} 家</text>
        <text x="${cx}" y="${midY + 14}" text-anchor="middle" opacity="${op}"
          fill="white" font-size="13" font-family="${f}"
          filter="url(#txtShadow)">${s.key}｜${s.name}　${pct}%</text>`;
    });

    // 底部橢圓封口
    const bLast = bounds[bounds.length - 1];
    rims += `<ellipse cx="${cx}" cy="${bLast.y}" rx="${bLast.hw}" ry="${bLast.hw * ryF}"
      fill="rgba(0,0,0,0.4)"/>`;

    defs += '</defs>';
    oFunnel.innerHTML = `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg"
      style="width:100%;display:block;max-height:400px;filter:drop-shadow(0 6px 18px rgba(0,0,0,0.28))">
      ${defs}${fills}${shines}${rims}${labels}
    </svg>`;
  }

  // ── 首頁商機推進進度摘要 ──
  const DASH_STAGES = [
    { key:'D', name:'D｜靜止中',   cls:'dpc-d' },
    { key:'C', name:'C｜Pipeline', cls:'dpc-c' },
    { key:'B', name:'B｜Upside',   cls:'dpc-b' },
    { key:'A', name:'A｜Commit',   cls:'dpc-a' },
  ];
  const pipelineCols = $('dashPipelineCols');
  if (pipelineCols) {
    pipelineCols.innerHTML = '';
    DASH_STAGES.forEach(({ key, name, cls }) => {
      const stageOpps = allOpportunities.filter(o => o.stage === key);
      const total = stageOpps.reduce((s, o) => s + (parseFloat(o.amount) || 0), 0);
      const col = document.createElement('div');
      col.className = `dash-pipeline-col ${cls}`;
      col.innerHTML = `
        <div class="dash-pipeline-col-name">${name}</div>
        <div class="dash-pipeline-col-count">${stageOpps.length}</div>
        <div class="dash-pipeline-col-label">筆商機</div>
        <div class="dash-pipeline-col-amt">$ ${total.toLocaleString()} 萬</div>`;
      col.addEventListener('click', () => showSection('pipeline'));
      pipelineCols.appendChild(col);
    });
  }
}

function getWonContactIds() {
  const ids = new Set(allOpportunities.filter(o => o.stage === 'Won').map(o => o.contactId));
  // 以公司為單位：同公司任一人設為 customer，整公司都算
  const customerCompanies = new Set(
    allContacts.filter(c => c.customerType === 'customer' && c.company).map(c => c.company)
  );
  allContacts.forEach(c => {
    if (c.customerType === 'customer' || (c.company && customerCompanies.has(c.company))) {
      ids.add(c.id);
    }
  });
  return ids;
}

function setActiveNav(section) {
  ['navHome','navManagerHome','navProspects','navContacts','navVisits','navTargets','navPipeline','navForecast','navLostOpp','navExecDash','navCampaigns','navLeads','navQuotations','navPipelineReport','navErpMa','navSapMa','navReceivables','navCallin'].forEach(id => { const el=$(id); if(el) el.classList.remove('active'); });
  const map = { null:'navHome', managerHome:'navManagerHome', prospects:'navProspects', contacts:'navContacts', visits:'navVisits', targets:'navTargets', pipeline:'navPipeline', forecast:'navForecast', lostOpp:'navLostOpp', execDash:'navExecDash', campaigns:'navCampaigns', leads:'navLeads', quotations:'navQuotations', pipelineReport:'navPipelineReport', 'erp-ma':'navErpMa', 'sap-ma':'navSapMa', receivables:'navReceivables', callin:'navCallin' };
  const el = $(map[section]);
  if (el) el.classList.add('active');
  const titles = { null:'首頁', managerHome:'主管首頁', prospects:'潛在客戶', contacts:'我的客戶', visits:'業務日報', targets:'業務年度目標', pipeline:'商機推進進度', forecast:'銷售預測報表', lostOpp:'流失商機分析', execDash:'管理儀表板', campaigns:'行銷活動', leads:'Lead 管理', quotations:'報價單管理', pipelineReport:'商機動態報表', 'erp-ma':'ERP MA 合約管理', 'sap-ma':'SAP License MA 管理', receivables:'應收帳款逾期', callin:'Call-in Pass 管理' };
  $('topbarTitle').textContent = titles[section] ?? '首頁';
  if (section === 'erp-ma' || section === 'sap-ma') {
    $('subMenuContract').classList.add('open');
    $('navContractGroup').querySelector('.nav-arrow').classList.add('open');
  }
  if (section === 'receivables') {
    $('subMenuAccounting').classList.add('open');
    $('navAccountingGroup').querySelector('.nav-arrow').classList.add('open');
  }
}

function showDashboard() {
  currentSection = null;
  localStorage.setItem('lastSection', 'dashboard');
  $('dashboardView').style.display = '';
  // 把所有 section view 全部隱藏
  ['prospectsView','contactsView','visitsView','targetsView','pipelineView',
   'forecastView','lostOppView','execDashView','managerHomeView','campaignsView','leadsView',
   'quotationsView','pipelineReportView','erpMaView','sapMaView',
   'receivablesView','callinView','transferView'].forEach(id => {
    const el = $(id); if (el) el.style.display = 'none';
  });
  $('prospectsToolbar').style.display = 'none';
  $('contactsToolbar').style.display  = 'none';
  $('visitsToolbar').style.display    = 'none';
  setActiveNav(null);
  updateStatCards();
  Promise.all([loadOpportunities(), loadTargets()]).then(() => {
    renderDashboardCharts();
    updateTargetCard();
  });
  loadBirthdayReminders();
  loadZombieAlertCard();
}

function showSection(section) {
  currentSection = section;
  localStorage.setItem('lastSection', section);
  $('dashboardView').style.display    = 'none';
  $('prospectsView').style.display    = section === 'prospects'   ? '' : 'none';
  $('contactsView').style.display     = section === 'contacts'    ? '' : 'none';
  $('visitsView').style.display       = section === 'visits'      ? '' : 'none';
  $('targetsView').style.display      = section === 'targets'     ? '' : 'none';
  $('pipelineView').style.display     = section === 'pipeline'    ? '' : 'none';
  $('forecastView').style.display     = section === 'forecast'    ? '' : 'none';
  $('execDashView').style.display     = section === 'execDash'    ? '' : 'none';
  $('managerHomeView').style.display  = section === 'managerHome' ? '' : 'none';
  $('campaignsView').style.display    = section === 'campaigns'   ? '' : 'none';
  $('leadsView').style.display        = section === 'leads'       ? '' : 'none';
  $('lostOppView').style.display          = section === 'lostOpp'        ? '' : 'none';
  $('quotationsView').style.display       = section === 'quotations'     ? '' : 'none';
  $('pipelineReportView').style.display   = section === 'pipelineReport' ? '' : 'none';
  $('erpMaView').style.display        = section === 'erp-ma'      ? '' : 'none';
  $('sapMaView').style.display        = section === 'sap-ma'      ? '' : 'none';
  $('receivablesView').style.display  = section === 'receivables' ? '' : 'none';
  $('callinView').style.display       = section === 'callin'      ? '' : 'none';
  $('transferView').style.display     = section === 'transfer'    ? '' : 'none';
  $('prospectsToolbar').style.display = section === 'prospects' ? 'flex' : 'none';
  $('contactsToolbar').style.display  = section === 'contacts'  ? 'flex' : 'none';
  $('visitsToolbar').style.display    = section === 'visits'    ? 'flex' : 'none';
  setActiveNav(section);
  if (section === 'prospects')   loadContacts();
  if (section === 'contacts')    loadContacts();
  if (section === 'visits')      { loadVisits(); if (allContacts.length === 0) loadContacts(); }
  if (section === 'targets')     loadTargetsView();
  if (section === 'pipeline')    loadPipelineView();
  if (section === 'forecast')    loadForecastView();
  if (section === 'execDash')    loadExecDash();
  if (section === 'managerHome') loadManagerHome();
  if (section === 'campaigns')   loadCampaignsView();
  if (section === 'leads')       loadLeadsView();
  if (section === 'lostOpp')         loadLostOppView();
  if (section === 'quotations')      loadQuotationsView();
  if (section === 'pipelineReport')  loadPipelineReport();
  if (section === 'erp-ma')      loadErpMaView();
  if (section === 'sap-ma')      loadSapMaView();
  if (section === 'receivables') loadReceivablesView();
  if (section === 'callin')      loadCallinView();
  if (section === 'transfer')    loadTransferView();
}

$('navHome').addEventListener('click', showDashboard);
$('navProspects').addEventListener('click', () => showSection('prospects'));
$('navContacts').addEventListener('click',  () => showSection('contacts'));
$('navVisits').addEventListener('click',    () => showSection('visits'));
$('navTargets').addEventListener('click',   () => showSection('targets'));
$('navPipeline').addEventListener('click',  () => showSection('pipeline'));
$('navForecast').addEventListener('click',  () => showSection('forecast'));
$('navLostOpp').addEventListener('click',         () => showSection('lostOpp'));
$('navExecDash').addEventListener('click',        () => showSection('execDash'));
$('navManagerHome').addEventListener('click',     () => showSection('managerHome'));
$('navCampaigns').addEventListener('click',       () => showSection('campaigns'));
$('navLeads').addEventListener('click',           () => showSection('leads'));
$('navQuotations').addEventListener('click',      () => showSection('quotations'));
$('navPipelineReport').addEventListener('click',  () => showSection('pipelineReport'));
$('navErpMa').addEventListener('click',     () => showSection('erp-ma'));
$('navSapMa').addEventListener('click',     () => showSection('sap-ma'));
$('navReceivables').addEventListener('click',() => showSection('receivables'));
$('navCallin').addEventListener('click',    () => showSection('callin'));
$('goSetTargetBtn').addEventListener('click', () => showSection('targets'));

// ── 合約管理側邊欄群組展開/收合 ─────────────────────────
$('navContractGroup').addEventListener('click', () => {
  const sub   = $('subMenuContract');
  const arrow = $('navContractGroup').querySelector('.nav-arrow');
  const isOpen = sub.classList.toggle('open');
  arrow.classList.toggle('open', isOpen);
});

// ── 帳務管理側邊欄群組展開/收合 ──────────────────────────
$('navAccountingGroup').addEventListener('click', () => {
  const sub   = $('subMenuAccounting');
  const arrow = $('navAccountingGroup').querySelector('.nav-arrow');
  const isOpen = sub.classList.toggle('open');
  arrow.classList.toggle('open', isOpen);
});

// ── 載入聯絡人 ───────────────────────────────────────────
let isSearchMode = false;

async function loadContacts(search = '') {
  try {
    isSearchMode = search.length > 0;
    const url = search ? `${API}/contacts?search=${encodeURIComponent(search)}` : `${API}/contacts`;
    const [res] = await Promise.all([
      fetch(url),
      allVisits.length === 0        ? fetch(`${API}/visits`).then(r => r.json()).then(d => { allVisits = d; }) : Promise.resolve(),
      allOpportunities.length === 0 ? fetch(`${API}/opportunities`).then(r => r.json()).then(d => { allOpportunities = d; }) : Promise.resolve(),
    ]);
    const data = await res.json();
    if (!isSearchMode) allContacts = data;
    updateCompanyDatalist(allContacts);

    // 依 section 過濾
    const wonIds = getWonContactIds();
    let displayed = data;
    if (currentSection === 'prospects') {
      displayed = data.filter(c => !wonIds.has(c.id));
    } else if (currentSection === 'contacts') {
      displayed = data.filter(c => wonIds.has(c.id));
    }

    renderContacts(displayed);

    // 更新計數
    if (currentSection === 'prospects') {
      const cos = new Set(displayed.filter(c => c.company).map(c => c.company)).size;
      $('prospectCompanyCount').textContent = `${cos} 家公司`;
      $('prospectContactCount').textContent = `${displayed.length} 位潛在客戶`;
    } else {
      const uniqueCompanies = new Set(allContacts.map(c => c.company).filter(Boolean)).size;
      $('companyCount2').textContent = `${uniqueCompanies} 家公司`;
      $('contactCount2').textContent = `${allContacts.length} 位聯絡人`;
    }
    updateStatCards();
    if (currentSection === null) renderDashboardCharts();
  } catch {
    showToast('無法連線至伺服器');
  }
}

// ── 建立單張聯絡人卡片 ───────────────────────────────────
function buildContactCard(c, extraClass = '') {
  const [c1, c2] = avatarColor(c.name);
  const card = document.createElement('div');
  card.className = 'contact-card' + (extraClass ? ' ' + extraClass : '');
  card.dataset.id = c.id;
  const titleOnly = c.title || '';
  const imageThumb = c.cardImage ? `<img class="card-image-thumb" src="${c.cardImage}" alt="名片">` : '';
  const oppBadge = c.opportunityStage
    ? `<span class="opp-badge opp-card ${OPPORTUNITY_LABELS[c.opportunityStage]?.cls || ''}">${c.opportunityStage}｜${OPPORTUNITY_LABELS[c.opportunityStage]?.label.split('｜')[1] || ''}</span>`
    : '';
  const primaryMark = c.isPrimary ? `<span class="primary-badge">&#11088; 主要</span>` : '';
  const resignedBadge = c.isResigned ? `<span class="resigned-badge">&#128683; 離職</span>` : '';
  const isCustomer = c.customerType === 'customer';
  const relLabel = isCustomer ? `我的客戶${c.productLine ? '｜' + c.productLine : ''}` : '潛在客戶';
  const relClass = isCustomer ? 'rel-customer' : 'rel-prospect';
  card.innerHTML = `
    <div class="card-top" style="position:relative">
      <div class="avatar" style="background:linear-gradient(135deg,${c1},${c2})">${getInitial(c.name)}</div>
      <div class="card-name-block">
        <div class="card-name">${escapeHtml(c.name) || '-'} ${primaryMark}</div>
        <div class="card-title-company">${escapeHtml(titleOnly) || '&nbsp;'}</div>
      </div>
      ${resignedBadge}
      ${imageThumb}
    </div>
    <div class="card-info">
      ${c.phone  ? `<div class="card-info-row"><span class="info-icon">&#128222;</span><span class="info-text">${escapeHtml(c.phone)}</span></div>` : ''}
      ${c.mobile ? `<div class="card-info-row"><span class="info-icon">&#128241;</span><span class="info-text">${escapeHtml(c.mobile)}</span></div>` : ''}
      ${c.email  ? `<div class="card-info-row"><span class="info-icon">&#128140;</span><span class="info-text">${escapeHtml(c.email)}</span></div>` : ''}
    </div>
    ${oppBadge ? `<div class="card-opp-row">${oppBadge}</div>` : ''}
    <div class="card-rel-row">
      <span class="contact-rel-badge ${relClass}" data-cid="${escapeHtml(c.id)}">${escapeHtml(relLabel)}</span>
    </div>`;
  card.addEventListener('click', () => openView(c.id));
  const relBadge = card.querySelector('.contact-rel-badge');
  relBadge.addEventListener('click', e => { e.stopPropagation(); openRelPicker(c, relBadge); });
  return card;
}

// ── 客戶關係徽章互動 ─────────────────────────────────────
function openRelPicker(c, anchor) {
  document.querySelectorAll('.rel-picker').forEach(p => p.remove());
  const picker = document.createElement('div');
  picker.className = 'rel-picker';
  picker.innerHTML = `
    <div class="rel-picker-item" data-val="prospect">潛在客戶</div>
    <div class="rel-picker-item" data-val="customer">我的客戶</div>`;
  const rect = anchor.getBoundingClientRect();
  picker.style.top  = (rect.bottom + 4) + 'px';
  picker.style.left = rect.left + 'px';
  document.body.appendChild(picker);
  picker.querySelectorAll('.rel-picker-item').forEach(item => {
    item.addEventListener('click', async e => {
      e.stopPropagation();
      picker.remove();
      if (item.dataset.val === 'customer') {
        openProductLineModal(c);
      } else {
        await setCustomerType(c.id, 'prospect', '');
      }
    });
  });
  setTimeout(() => document.addEventListener('click', () => picker.remove(), { once: true }), 0);
}

function openProductLineModal(c) {
  document.querySelectorAll('.product-line-overlay').forEach(p => p.remove());
  const overlay = document.createElement('div');
  overlay.className = 'product-line-overlay';
  overlay.innerHTML = `
    <div class="product-line-modal">
      <div class="product-line-title">選擇客戶類別</div>
      <div class="product-line-sub">請選擇此客戶所屬的產品線</div>
      <div class="product-line-options">
        <button class="pl-btn pl-erp" data-pl="ERP">ERP</button>
        <button class="pl-btn pl-its" data-pl="ITS">ITS</button>
      </div>
      <button class="pl-cancel">取消</button>
    </div>`;
  document.body.appendChild(overlay);
  overlay.querySelectorAll('.pl-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const pl = btn.dataset.pl;
      overlay.remove();
      await setCustomerType(c.id, 'customer', pl);
    });
  });
  overlay.querySelector('.pl-cancel').addEventListener('click', () => overlay.remove());
  overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
}

async function setCustomerType(contactId, customerType, productLine) {
  const contact = allContacts.find(c => c.id === contactId);
  if (!contact) return;
  // 同公司所有聯絡人一起更新
  const targets = contact.company
    ? allContacts.filter(c => c.company === contact.company)
    : [contact];
  try {
    const results = await Promise.all(targets.map(c =>
      fetch(`/api/contacts/${c.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ...c, customerType, productLine })
      })
    ));
    // 檢查每個回應是否成功
    for (const r of results) {
      if (!r.ok) {
        const err = await r.json().catch(() => ({}));
        throw new Error(err.error || `HTTP ${r.status}`);
      }
    }
    const label = customerType === 'customer'
      ? `已將「${contact.company || contact.name}」設為我的客戶（${productLine}）`
      : `已將「${contact.company || contact.name}」移回潛在客戶`;
    showToast(label);
    await loadContacts();
  } catch(e) {
    console.error('setCustomerType error:', e);
    showToast('更新失敗：' + e.message);
  }
}

// ── 渲染名片列表 ─────────────────────────────────────────
function renderContacts(contacts) {
  const isProspects = currentSection === 'prospects';
  const list = $(isProspects ? 'prospectList' : 'contactList');
  list.innerHTML = '';

  if (contacts.length === 0) {
    list.className = 'contact-list';
    list.innerHTML = isProspects
      ? `<div class="empty-state"><div class="empty-icon">&#128270;</div><p>尚無潛在客戶</p><p class="empty-sub">點擊右上角「新增潛在客戶」開始建立</p></div>`
      : `<div class="empty-state"><div class="empty-icon">&#127942;</div><p>尚無成交客戶</p><p class="empty-sub">商機成交後客戶將自動移入此處</p></div>`;
    return;
  }

  if (isSearchMode) {
    // 搜尋模式：平鋪所有結果
    list.className = 'contact-list';
    contacts.forEach(c => list.appendChild(buildContactCard(c)));
  } else {
    // 預設模式：依公司分組堆疊
    list.className = 'contact-list grouped';
    const withCompany = new Map();
    const noCompany = [];
    contacts.forEach(c => {
      if (c.company) {
        if (!withCompany.has(c.company)) withCompany.set(c.company, []);
        withCompany.get(c.company).push(c);
      } else {
        noCompany.push(c);
      }
    });

    // 依公司名排序
    [...withCompany.entries()]
      .sort(([a], [b]) => a.localeCompare(b, 'zh-TW'))
      .forEach(([company, members]) => {
        const primary = members.find(c => c.isPrimary) || members[0];
        const opp = members.find(c => c.opportunityStage)?.opportunityStage;
        const industry = members.find(c => c.industry)?.industry;
        const [c1, c2] = avatarColor(company);
        const oppBadge = opp
          ? `<span class="opp-badge opp-card ${OPPORTUNITY_LABELS[opp]?.cls || ''}">${opp}｜${OPPORTUNITY_LABELS[opp]?.label.split('｜')[1] || ''}</span>` : '';
        const industryBadge = industry ? `<span class="industry-badge">${industry}</span>` : '';

        // 最近拜訪記錄
        const memberIds = new Set(members.map(c => c.id));
        const recentVisit = [...allVisits]
          .filter(v => memberIds.has(v.contactId))
          .sort((a, b) => (b.visitDate || '').localeCompare(a.visitDate || ''))
          [0];

        // 最近商機
        const recentOpp = [...allOpportunities]
          .filter(o => (o.company || '') === company && o.stage !== 'Won')
          .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))
          [0];

        const recentVisitHtml = recentVisit
          ? `<div class="cg-recent-row">
               <span class="cg-recent-icon">&#128203;</span>
               <span class="cg-recent-date">${recentVisit.visitDate || ''}</span>
               <span class="cg-recent-text">${recentVisit.topic || ''}</span>
             </div>`
          : `<div class="cg-recent-row cg-recent-empty">尚無拜訪記錄</div>`;

        const recentOppHtml = recentOpp
          ? `<div class="cg-recent-row">
               <span class="cg-recent-icon">&#128161;</span>
               <span class="cg-opp-stage opp-${(recentOpp.stage||'').toLowerCase()}">${recentOpp.stage}</span>
               <span class="cg-recent-text">${recentOpp.product || recentOpp.category || ''}</span>
               ${recentOpp.amount ? `<span class="cg-opp-amt">${Number(recentOpp.amount).toLocaleString()} 萬</span>` : ''}
             </div>`
          : '';

        const group = document.createElement('div');
        group.className = 'company-group';
        group.innerHTML = `
          <div class="company-group-header">
            <div class="company-avatar" style="background:linear-gradient(135deg,${c1},${c2})">${company.charAt(0)}</div>
            <div class="company-group-info">
              <div class="company-group-name">${escapeHtml(company)}</div>
              <div class="company-group-primary">&#11088; ${escapeHtml(primary.name)}${primary.title ? '・' + escapeHtml(primary.title) : ''}</div>
              <div class="company-group-recent">
                ${recentVisitHtml}
                ${recentOppHtml}
              </div>
            </div>
            <div class="company-group-badges">${industryBadge}${oppBadge}</div>
            <div class="company-group-count">${members.length} 人</div>
            <button class="cg-add-contact-btn" title="新增此公司聯絡人">&#43; 新增聯絡人</button>
            <div class="company-group-chevron">&#9660;</div>
          </div>
          <div class="company-group-members collapsed"></div>`;

        const membersDiv = group.querySelector('.company-group-members');
        members.forEach(c => membersDiv.appendChild(buildContactCard(c, 'contact-card-sub')));

        // 「新增聯絡人」按鈕：帶入同公司資訊
        group.querySelector('.cg-add-contact-btn').addEventListener('click', e => {
          e.stopPropagation(); // 不觸發展開/收合
          // 從同公司找出共用資訊（電話、分機、地址、網站、統編、產業、系統廠商）
          const ref = primary;
          openModalAddContact({
            company:      company,
            phone:        ref.phone        || '',
            ext:          '',
            address:      ref.address      || '',
            website:      ref.website      || '',
            taxId:        ref.taxId        || '',
            industry:     ref.industry     || (members.find(c => c.industry)?.industry || ''),
            systemVendor: ref.systemVendor || (members.find(c => c.systemVendor)?.systemVendor || ''),
            systemProduct:ref.systemProduct|| '',
          });
        });

        group.querySelector('.company-group-header').addEventListener('click', () => {
          const isOpen = !membersDiv.classList.contains('collapsed');
          membersDiv.classList.toggle('collapsed', isOpen);
          group.classList.toggle('expanded', !isOpen);
        });

        list.appendChild(group);
      });

    // 無公司的聯絡人放最後
    if (noCompany.length > 0) {
      const section = document.createElement('div');
      section.className = 'no-company-section';
      section.innerHTML = `<div class="no-company-label">無公司聯絡人</div><div class="no-company-grid"></div>`;
      const grid = section.querySelector('.no-company-grid');
      noCompany.forEach(c => grid.appendChild(buildContactCard(c)));
      list.appendChild(section);
    }
  }
}

// ── 頁簽切換 ─────────────────────────────────────────────
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', function () {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
    this.classList.add('active');
    $(this.dataset.tab).classList.remove('hidden');

    if (this.dataset.tab === 'tab-company') {
      const tid = $('taxId').value.trim();
      $('lookupTaxIdDisplay').textContent = tid || '尚未填入統編';
    }

    // 切到商機頁簽
    if (this.dataset.tab === 'tab-opp') {
      $('saveBtn').style.display = 'none';
      $('cOppSaveBtn').style.display = '';
      const hasId = !!$('contactId').value;
      $('cOppNoIdNote').style.display = hasId ? 'none' : '';
      $('cOppForm').style.display     = hasId ? '' : 'none';
      $('cOppSaveBtn').disabled       = !hasId;
    } else {
      $('saveBtn').style.display    = '';
      $('cOppSaveBtn').style.display = 'none';
    }
  });
});

// ── 商品「其他」自訂欄位 ─────────────────────────────────
function handleProductOther(selectId, customId) {
  const sel = $(selectId), inp = $(customId);
  if (!sel || !inp) return;
  sel.addEventListener('change', () => {
    const isOther = sel.value === '其他';
    inp.style.display = isOther ? '' : 'none';
    if (isOther) { inp.value = ''; inp.focus(); }
    else inp.value = '';
  });
}

// 取得最終商品值（若選「其他」則回傳自訂輸入）
function getProductValue(selectId, customId) {
  const sel = $(selectId), inp = $(customId);
  if (sel && sel.value === '其他' && inp) return inp.value.trim() || '其他';
  return sel ? sel.value : '';
}

handleProductOther('oppProduct',     'oppProductCustom');
handleProductOther('cOppProduct',    'cOppProductCustom');
handleProductOther('oppEditProduct', 'oppEditProductCustom');

// ── 聯絡人 Modal 商機 TAB ─────────────────────────────────
$('cOppCategory').addEventListener('change', function () {
  const ps = $('cOppProduct');
  const pg = $('cOppProductGroup');
  ps.innerHTML = '<option value="">-- 請選擇 --</option>';
  const catData = OPP_PRODUCTS[this.value];
  const groups  = catData && typeof catData === 'object' && !Array.isArray(catData) ? catData : null;
  const hasItems = groups && Object.values(groups).some(arr => arr.length > 0);
  if (hasItems) {
    Object.entries(groups).forEach(([groupLabel, items]) => {
      if (!items.length) return;
      const grp = document.createElement('optgroup');
      grp.label = groupLabel;
      items.forEach(p => {
        const o = document.createElement('option');
        o.value = p; o.textContent = p;
        grp.appendChild(o);
      });
      ps.appendChild(grp);
    });
    pg.style.display = '';
  } else {
    pg.style.display = 'none';
  }
});

$('cOppSaveBtn').addEventListener('click', async () => {
  const contactId = $('contactId').value;
  if (!contactId) { showToast('請先儲存聯絡人資料'); return; }
  const category = $('cOppCategory').value;
  if (!category)  { showToast('請選擇商機類別'); return; }

  const contact = allContacts.find(c => c.id === contactId);
  const payload = {
    contactId,
    contactName:    contact ? contact.name    : $('name').value.trim(),
    company:        contact ? contact.company : $('company').value.trim(),
    category,
    product:        getProductValue('cOppProduct', 'cOppProductCustom'),
    stage:          $('cOppStage').value || 'C',
    amount:         $('cOppAmount').value,
    grossMarginRate:$('cOppGrossMargin').value,
    expectedDate:   $('cOppExpectedDate').value,
    description:    $('cOppDescription').value.trim(),
  };

  try {
    await fetch(`${API}/opportunities`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    showToast('✅ 商機已新增');
    // 重置表單
    $('cOppCategory').value    = '';
    $('cOppProduct').innerHTML = '<option value="">-- 請選擇 --</option>';
    $('cOppProductGroup').style.display = 'none';
    $('cOppStage').value       = 'C';
    $('cOppAmount').value      = '';
    $('cOppGrossMargin').value = '';
    $('cOppExpectedDate').value= '';
    $('cOppDescription').value = '';
    await loadOpportunities();
    updateStatCards();
  } catch { showToast('新增失敗，請重試'); }
});

function resetContactOppTab() {
  $('cOppCategory').value    = '';
  $('cOppProduct').innerHTML = '<option value="">-- 請選擇 --</option>';
  $('cOppProductGroup').style.display = 'none';
  $('cOppStage').value       = 'C';
  $('cOppAmount').value      = '';
  $('cOppGrossMargin').value = '';
  $('cOppExpectedDate').value= '';
  $('cOppDescription').value = '';
  $('saveBtn').style.display    = '';
  $('cOppSaveBtn').style.display = 'none';
  // 切回第一個 tab
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
  document.querySelector('[data-tab="tab-card"]').classList.add('active');
  $('tab-card').classList.remove('hidden');
}

// ── Email Domain 自動完成 ────────────────────────────────
function getHistoryDomains() {
  const domains = allContacts
    .map(c => (c.email || '').split('@')[1])
    .filter(Boolean);
  return [...new Set(domains)]; // 去重複
}

$('email').addEventListener('input', function () {
  const val = this.value;
  const atIdx = val.indexOf('@');
  const dropdown = $('domainDropdown');

  if (atIdx === -1) { dropdown.innerHTML = ''; dropdown.classList.remove('open'); return; }

  const typed = val.slice(atIdx + 1).toLowerCase();
  const domains = getHistoryDomains().filter(d => d.toLowerCase().startsWith(typed));

  if (domains.length === 0) { dropdown.innerHTML = ''; dropdown.classList.remove('open'); return; }

  dropdown.innerHTML = domains.map(d =>
    `<div class="domain-item" data-domain="${escapeHtml(d)}">${escapeHtml(val.slice(0, atIdx + 1))}<strong>${escapeHtml(d)}</strong></div>`
  ).join('');
  dropdown.classList.add('open');
});

$('email').addEventListener('blur', () => {
  setTimeout(() => { $('domainDropdown').classList.remove('open'); }, 150);
});

$('domainDropdown').addEventListener('click', function (e) {
  const item = e.target.closest('.domain-item');
  if (!item) return;
  const atIdx = $('email').value.indexOf('@');
  $('email').value = $('email').value.slice(0, atIdx + 1) + item.dataset.domain;
  this.classList.remove('open');
  this.innerHTML = '';
});

// ── 語音輸入統編 ─────────────────────────────────────────
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;

function chineseToDigits(str) {
  const map = {
    '零':'0','一':'1','二':'2','三':'3','四':'4','五':'5','六':'6','七':'7','八':'8','九':'9',
    '壹':'1','貳':'2','參':'3','肆':'4','伍':'5','陸':'6','柒':'7','捌':'8','玖':'9',
    'O':'0','o':'0','Ｏ':'0','ｏ':'0'
  };
  return str.replace(/[零一二三四五六七八九壹貳參肆伍陸柒捌玖OoＯｏ]/g, c => map[c] || c)
            .replace(/[^0-9]/g, '');
}

if (SpeechRecognition) {
  const recognition = new SpeechRecognition();
  recognition.lang = 'zh-TW';
  recognition.interimResults = false;
  recognition.maxAlternatives = 1;

  const micBtn = $('taxIdMicBtn');

  micBtn.addEventListener('click', () => {
    if (micBtn.classList.contains('listening')) { recognition.stop(); return; }
    recognition.start();
  });

  recognition.addEventListener('start', () => {
    $('taxIdMicBtn').classList.add('listening');
    $('taxIdMicBtn').textContent = '⏹';
    $('taxIdMicBtn').title = '點擊停止';
    showToast('🎙 請說出統一編號數字...', 3000);
  });

  recognition.addEventListener('result', (e) => {
    const transcript = e.results[0][0].transcript;
    const digits = chineseToDigits(transcript).slice(0, 8);
    if (digits) {
      $('taxId').value = digits;
      $('taxId').dispatchEvent(new Event('input'));
      showToast(`✓ 語音辨識：${digits}`);
    } else {
      showToast('未辨識到數字，請再試一次');
    }
  });

  recognition.addEventListener('end', () => {
    $('taxIdMicBtn').classList.remove('listening');
    $('taxIdMicBtn').textContent = '🎙';
    $('taxIdMicBtn').title = '語音輸入統編';
  });

  recognition.addEventListener('error', (e) => {
    $('taxIdMicBtn').classList.remove('listening');
    $('taxIdMicBtn').textContent = '🎙';
    const msg = e.error === 'not-allowed' ? '請允許麥克風權限' : '語音辨識失敗，請再試';
    showToast(msg);
  });
} else {
  if ($('taxIdMicBtn')) $('taxIdMicBtn').style.display = 'none';
}

// ── 統編自動帶入公司資料 ─────────────────────────────────
let taxIdTimer = null;
$('taxId').addEventListener('input', function () {
  const tid = this.value.trim();
  $('lookupTaxIdDisplay').textContent = tid || '尚未填入統編';
  $('taxIdOkBadge').style.display = 'none';
  clearTimeout(taxIdTimer);
  if (tid.length !== 8 || !/^\d{8}$/.test(tid)) return;

  $('taxIdLoadingBadge').style.display = 'inline';
  taxIdTimer = setTimeout(async () => {
    try {
      const res = await fetch(`/api/company-lookup?taxId=${tid}`);
      if (!res.ok) return;
      const d = await res.json();

      // 只填入公司相關欄位，不覆蓋聯絡人欄位
      if (d.companyName) $('company').value = d.companyName;
      if (d.address)     $('address').value = d.address;

      // 產業自動判斷
      if (d.companyName) {
        const detected = detectIndustry(d.companyName);
        if (detected) {
          $('industry').value = detected;
          $('industry').dataset.manual = 'false';
          $('autoDetectBadge').style.display = 'inline';
        }
      }

      // 同步公司資訊頁簽的查詢結果
      renderCompanyInfo(d);

      $('taxIdOkBadge').style.display = 'inline';
      showToast(`✓ 已帶入：${d.companyName || '查無資料'}`);
    } catch { /* ignore */ }
    finally { $('taxIdLoadingBadge').style.display = 'none'; }
  }, 600);
});

// ── 公司查詢 ─────────────────────────────────────────────
$('lookupBtn').addEventListener('click', async () => {
  const taxId = $('taxId').value.trim();
  if (!taxId || taxId.length !== 8) {
    showToast('請先在名片資料頁簽填入 8 碼統一編號');
    return;
  }
  const result = $('companyInfoResult');
  result.innerHTML = `<div class="ci-loading"><span class="ci-spinner"></span>查詢中，請稍候...</div>`;
  try {
    const res = await fetch(`/api/company-lookup?taxId=${taxId}`);
    if (!res.ok) { showToast('查詢失敗，請稍後再試'); return; }
    const d = await res.json();
    renderCompanyInfo(d);
    // 自動回填 Tab1 欄位（如果是空的）
    if (d.companyName && !$('company').value) $('company').value = d.companyName;
    if (d.address   && !$('address').value)   $('address').value = d.address;
    if (d.website   && !$('website').value)   $('website').value = d.website;
    // 同步更新 AI 公司分析的網址欄位
    if (d.website) $('companyInsightUrl').value = d.website;
  } catch { showToast('查詢失敗，請確認網路連線'); result.innerHTML = '<div class="lookup-hint">查詢失敗，請稍後再試</div>'; }
});

function renderCompanyInfo(d) {
  const listedBadge = d.stockCode
    ? `<span class="${d.listedType === '上市' ? 'badge-listed' : 'badge-otc'}">${d.listedType}</span><span class="stock-code">${d.stockCode}</span><small style="color:#888;margin-left:6px">${d.exchange}</small>`
    : `<span class="badge-unlisted">未上市櫃</span>`;

  const yr1 = d.dataYear1 || 2025;
  const yr2 = d.dataYear2 || 2024;
  function finVal(v, strong) {
    if (!v || v === 'N/A' || v === '無法取得') return `<span class="fin-na">${v || 'N/A'}</span>`;
    return strong ? `<strong>${v}</strong>` : v;
  }
  const epsDisplay = (d.eps && d.eps !== 'N/A')
    ? `<span class="eps-badge">EPS &nbsp;${d.eps} 元<small style="opacity:.65;margin-left:6px">${d.epsYear} 全年</small></span>`
    : '';
  const finSection = d.stockCode ? `
    <div class="ci-section">
      <div class="ci-section-title">財務資訊</div>
      ${epsDisplay ? `<div class="ci-row" style="margin-bottom:10px"><span class="ci-label">最新 EPS</span><span class="ci-value">${epsDisplay}</span></div>` : ''}
      <table class="fin-table">
        <tr><th>項目</th><th>${yr1} 年度</th><th>${yr2} 年度</th></tr>
        <tr>
          <td>年度營業額</td>
          <td>${finVal(d.revenue2025, false)}</td>
          <td>${finVal(d.revenue2024, false)}</td>
        </tr>
        <tr>
          <td>毛利率</td>
          <td>${finVal(d.grossMargin2025, true)}</td>
          <td>${finVal(d.grossMargin2024, true)}</td>
        </tr>
      </table>
    </div>` : '';

  $('companyInfoResult').innerHTML = `
    <div class="company-info-card">
      <div class="ci-section">
        <div class="ci-section-title">基本資料</div>
        <div class="ci-row"><span class="ci-label">公司名稱</span><span class="ci-value">${d.companyName || '-'}</span></div>
        <div class="ci-row"><span class="ci-label">負責人</span><span class="ci-value">${d.representative || '-'}</span></div>
        <div class="ci-row"><span class="ci-label">公司狀態</span><span class="ci-value">${d.companyStatus || '-'}</span></div>
        <div class="ci-row"><span class="ci-label">資本額</span><span class="ci-value">${d.capital || '-'}</span></div>
        <div class="ci-row"><span class="ci-label">地址</span><span class="ci-value">${d.address || '-'}</span></div>
      </div>
      <div class="ci-section">
        <div class="ci-section-title">上市／上櫃資訊</div>
        <div class="ci-row"><span class="ci-label">市場別</span><span class="ci-value"><span class="listed-badge">${listedBadge}</span></span></div>
      </div>
      ${finSection}
    </div>`;
}

// ── Feature 6：KA 客戶五構面分析 ─────────────────────────
function renderKaDim(title, dim) {
  const signalMap = {
    green:  { icon: '🟢', label: '正向', cls: 'ka-signal-green' },
    yellow: { icon: '🟡', label: '觀察', cls: 'ka-signal-yellow' },
    red:    { icon: '🔴', label: '風險', cls: 'ka-signal-red' }
  };
  const s = signalMap[dim?.signal] || signalMap.yellow;
  const rows = Object.entries(dim || {})
    .filter(([k]) => k !== 'signal' && k !== 'salesHook')
    .map(([, v]) => `<div class="ka-dim-row">${escapeHtml(String(v))}</div>`)
    .join('');
  return `
    <div class="ka-dim-card ${s.cls}">
      <div class="ka-dim-header">
        <span>${s.icon}</span>
        <span class="ka-dim-title">${title}</span>
        <span class="ka-signal-label">${s.label}</span>
      </div>
      <div class="ka-dim-body">${rows}</div>
      <div class="ka-dim-hook">💡 ${escapeHtml(dim?.salesHook || '')}</div>
    </div>`;
}

$('companyInsightBtn').addEventListener('click', async () => {
  const btn = $('companyInsightBtn');
  const url = $('companyInsightUrl').value.trim();
  if (!url) return showToast('請輸入公司官網網址', 'error');
  if (!/^https?:\/\//i.test(url)) return showToast('網址格式錯誤，請包含 https://', 'error');

  btn.disabled = true; btn.textContent = '分析中…';
  const resultDiv = $('companyInsightResult');
  resultDiv.style.display = '';
  resultDiv.innerHTML = '<div style="color:#888;font-size:13px;padding:8px 0">🤖 AI 正在進行五構面分析，請稍候（約 15-20 秒）…</div>';

  try {
    const r = await fetch(`${API}/ai/company-insight`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ url })
    });
    const d = await r.json();
    if (!r.ok) {
      resultDiv.innerHTML = `<div class="ai-insight-error">${escapeHtml(d.error || 'AI 發生錯誤，請重試')}</div>`;
      return;
    }
    const dims = [
      ['戰略與市場', d.strategic],
      ['財務構面',   d.financial],
      ['營運與風險', d.operational],
      ['人力資本',   d.humanCapital],
      ['客戶與品牌', d.customerBrand],
    ];
    const opps = (d.topOpportunities || [])
      .map((o, i) => `<span class="ka-opp-tag">${['①','②','③'][i]||'·'} ${escapeHtml(String(o))}</span>`)
      .join('');
    resultDiv.innerHTML = `
      <div class="ka-insight-wrap">
        <div class="ka-company-name">${escapeHtml(d.companyName||'')}</div>
        <div class="ka-analysis-base">📄 ${escapeHtml(d.analysisBase||'')}</div>
        ${dims.map(([t, dim]) => renderKaDim(t, dim)).join('')}
        <div class="ka-summary">
          <div class="ka-summary-title">📋 整體建議</div>
          <div class="ka-summary-text">${escapeHtml(d.executiveSummary||'')}</div>
        </div>
        <div class="ka-opps">
          <div class="ka-opps-title">⭐ 優先機會點</div>
          <div class="ka-opps-list">${opps}</div>
        </div>
      </div>`;
  } catch (e) {
    resultDiv.innerHTML = `<div class="ai-insight-error">網路錯誤：${escapeHtml(e.message)}</div>`;
  } finally {
    btn.disabled = false; btn.textContent = '🔍 分析';
  }
});

// ── 商機分類徽章 ─────────────────────────────────────────
const OPPORTUNITY_LABELS = {
  A: { label: 'A｜Commit', sub: '本月有案子可簽約',        cls: 'opp-a' },
  B: { label: 'B｜Upside', sub: '已報價，三個月內可落袋',  cls: 'opp-b' },
  C: { label: 'C｜Pipeline', sub: '有商機，需求已明確',    cls: 'opp-c' },
  D: { label: 'D｜靜止中', sub: '暫無商機',               cls: 'opp-d' },
};
function renderOpportunityBadge(stage) {
  const o = OPPORTUNITY_LABELS[stage];
  if (!o) return stage;
  return `<span class="opp-badge ${o.cls}">${o.label}</span><span class="opp-sub">${o.sub}</span>`;
}

// ── 產業自動判斷 ─────────────────────────────────────────
const INDUSTRY_RULES = [
  { industry: '半導體', keywords: ['半導體','晶圓','晶片','積體電路','封測','TSMC','台積','聯電','聯發科','日月光','矽品','力積電','世界先進','瑞昱','novatek','聯詠','群聯','慧榮','矽統','南亞科','華邦','旺宏','winbond','realtek','mediatek','semiconductor'] },
  { industry: '科技業', keywords: ['科技','資訊','軟體','系統','網路','數位','雲端','AI','tech','software','IT','cloud','data','solution','solutions','資安','智慧','IoT','SaaS','platform','微軟','google','apple','Meta','IBM','intel','nvidia','amd','cisco','oracle','SAP'] },
  { industry: '電子製造', keywords: ['電子','鴻海','廣達','仁寶','緯創','英業達','和碩','台達','光寶','鴻準','富士康','foxconn','quanta','compal','wistron','pegatron','delta','liteon','PCB','電路板','主機板','ASUS','華碩','宏碁','acer','HTC','研華','advantech'] },
  { industry: '製造業', keywords: ['工業','製造','機械','自動化','零件','精密','模具','鑄造','沖壓','焊接','組裝','生產','工廠','automotive','汽車','車輛','輪胎','鋼鐵','鋁','金屬','塑膠','化工','石化','中鋼','台塑','台化'] },
  { industry: '金融業', keywords: ['銀行','保險','證券','投信','投顧','金融','期貨','基金','資產','信託','租賃','bank','finance','insurance','富邦','國泰','中信','玉山','永豐','兆豐','第一','台新','遠東商銀','凱基'] },
  { industry: '醫療／生技', keywords: ['醫院','醫療','生技','製藥','藥廠','藥局','健康','biotech','pharma','medical','health','診所','長庚','台大醫','榮總','慈濟','聯合醫','中醫','牙科','器材','基因','疫苗'] },
  { industry: '零售／電商', keywords: ['百貨','零售','電商','購物','超市','超商','便利','量販','momo','pchome','蝦皮','shopee','amazon','7-11','全家','全聯','costco','ikea','outlet','商場','連鎖'] },
  { industry: '建設／不動產', keywords: ['建設','地產','房屋','建築','營造','開發','estate','realty','住宅','豪宅','工程','承包','土木','信義房屋','永慶','遠雄','國泰建設','長虹','興富發'] },
  { industry: '顧問／服務業', keywords: ['顧問','會計','審計','法律','律師','諮詢','consultant','consulting','kpmg','pwc','deloitte','ey','管理','人資','獵頭','公關','廣告','行銷','媒體','傳播','創意'] },
  { industry: '政府／公家機關', keywords: ['政府','公所','市政','縣政','部','局','院','署','處','委員會','行政院','立法院','經濟部','財政部','教育部','衛生局','警察','消防','台電','中華電信','中油','台水'] },
  { industry: '教育', keywords: ['大學','學校','學院','高中','國中','國小','教育','university','college','school','institute','研究院','學術','補習班','培訓'] },
  { industry: '傳產／食品', keywords: ['食品','飲料','農業','畜牧','水產','紡織','成衣','皮革','木材','造紙','印刷','菸酒','統一','味全','泰山','黑松','台糖','大成','卜蜂','桂格','義美'] },
];

function detectIndustry(companyName) {
  if (!companyName) return '';
  const name = companyName.toLowerCase();
  for (const rule of INDUSTRY_RULES) {
    if (rule.keywords.some(kw => name.includes(kw.toLowerCase()))) {
      return rule.industry;
    }
  }
  return '';
}

// ── 公司名稱自動完成 ─────────────────────────────────────
function updateCompanyDatalist(contacts) {
  const seen = new Set();
  const recent = [...contacts]
    .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))
    .filter(c => c.company && !seen.has(c.company) && seen.add(c.company))
    .slice(0, 10)
    .map(c => c.company);

  const dl = $('companyList');
  dl.innerHTML = recent.map(name => `<option value="${name}">`).join('');
}

// ── 搜尋 ─────────────────────────────────────────────────
$('searchInput').addEventListener('input', function () {
  const val = this.value.trim();
  $('clearSearch').classList.toggle('visible', val.length > 0);
  loadContacts(val);
});
$('clearSearch').addEventListener('click', () => {
  $('searchInput').value = '';
  $('clearSearch').classList.remove('visible');
  loadContacts();
});

// ── 職能分類選擇器渲染 ────────────────────────────────────
function renderJfSelector(selectedKey) {
  const container = $('jfSelector');
  if (!container) return;
  container.innerHTML = '';
  JOB_FUNCTION_CATEGORIES.forEach(cat => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'jf-btn' + (cat.key === selectedKey ? ' active' : '');
    btn.dataset.key = cat.key;
    btn.style.setProperty('--jf-color', cat.color);
    btn.style.setProperty('--jf-bg', cat.bg);
    btn.innerHTML = `<span class="jf-btn-label">${cat.label}</span><span class="jf-btn-sub">${cat.sublabel}</span>`;
    btn.addEventListener('click', () => {
      const isActive = btn.classList.contains('active');
      container.querySelectorAll('.jf-btn').forEach(b => b.classList.remove('active'));
      if (!isActive) {
        btn.classList.add('active');
        $('jobFunction').value = cat.key;
      } else {
        $('jobFunction').value = '';
      }
    });
    container.appendChild(btn);
  });
}

// ── 新增/編輯 Modal ──────────────────────────────────────
function openModal(contact = null) {
  $('contactForm').reset();
  $('contactId').value = '';
  $('cardImageUrl').value = '';
  $('previewImg').style.display = 'none';
  $('previewImg').src = '';
  $('cardImagePreview').querySelector('.upload-hint').style.display = 'flex';
  $('imageActions').style.display = 'none';

  if (contact) {
    $('modalTitle').textContent = '編輯名片';
    $('contactId').value = contact.id;
    $('name').value = contact.name || '';
    $('nameEn').value = contact.nameEn || '';
    $('company').value = contact.company || '';
    $('title').value = contact.title || '';
    $('phone').value = contact.phone || '';
    $('mobile').value = contact.mobile || '';
    $('ext').value = contact.ext || '';
    $('email').value = contact.email || '';
    $('address').value = contact.address || '';
    $('website').value = contact.website || '';
    $('taxId').value = contact.taxId || '';
    $('industry').value = contact.industry || '';
    $('industry').dataset.manual = contact.industry ? 'true' : 'false';
    $('autoDetectBadge').style.display = 'none';
    $('systemVendor').value = contact.systemVendor || '';
    updateSystemProduct(contact.systemVendor || '', contact.systemProduct || '');
    if ($('opportunityStage')) $('opportunityStage').value = contact.opportunityStage || '';
    $('note').value = contact.note || '';
    // 帶入主要聯繫窗口
    $('isPrimaryYes').checked = !!contact.isPrimary;
    $('isPrimaryNo').checked  = !contact.isPrimary;
    // 帶入離職狀態
    $('isResigned').checked = !!contact.isResigned;
    // 帶入個人資訊
    loadPersonalInfo(contact);
    if (contact.cardImage) {
      $('cardImageUrl').value = contact.cardImage;
      $('previewImg').src = contact.cardImage;
      $('previewImg').style.display = 'block';
      $('cardImagePreview').querySelector('.upload-hint').style.display = 'none';
      $('imageActions').style.display = 'flex';
    }
    // 職能分類：有存就用，沒有就嘗試自動 mapping
    const jfKey = contact.jobFunction || autoMapJobFunction(contact.title || '');
    $('jobFunction').value = jfKey;
    renderJfSelector(jfKey);
  } else {
    $('modalTitle').textContent = currentSection === 'prospects' ? '新增潛在客戶' : '新增名片';
    $('industry').value = '';
    $('industry').dataset.manual = 'false';
    $('autoDetectBadge').style.display = 'none';
    if ($('opportunityStage')) $('opportunityStage').value = 'C'; // 預設 Pipeline
    $('isPrimaryNo').checked = true;
    $('isResigned').checked = false;
    loadPersonalInfo({});
    updateSystemProduct('');
    $('jobFunction').value = '';
    renderJfSelector('');
  }
  // AI 公司背景分析：帶入官網網址、清空上次結果
  $('companyInsightUrl').value            = contact?.website || '';
  $('companyInsightResult').style.display = 'none';
  $('companyInsightResult').innerHTML     = '';

  $('modalOverlay').classList.add('open');
  $('name').focus();
}

function closeModal() {
  $('modalOverlay').classList.remove('open');
  if (typeof resetContactOppTab === 'function') resetContactOppTab();
}

// ── 個人資訊 Tab 讀寫 ─────────────────────────────────────
function loadPersonalInfo(c) {
  // 清除所有勾選
  document.querySelectorAll('#tab-personal input[type="checkbox"]').forEach(cb => cb.checked = false);
  // 飲料偏好
  const drinks = (c.personalDrink || '').split(',').map(s => s.trim()).filter(Boolean);
  document.querySelectorAll('input[name="drinkPref"]').forEach(cb => {
    if (drinks.includes(cb.value)) cb.checked = true;
  });
  // 興趣愛好
  const hobbies = (c.personalHobbies || '').split(',').map(s => s.trim()).filter(Boolean);
  document.querySelectorAll('input[name="hobbies"]').forEach(cb => {
    if (hobbies.includes(cb.value)) cb.checked = true;
  });
  // 飲食禁忌
  const diet = (c.personalDiet || '').split(',').map(s => s.trim()).filter(Boolean);
  document.querySelectorAll('input[name="diet"]').forEach(cb => {
    if (diet.includes(cb.value)) cb.checked = true;
  });
  // 生日 & 備忘
  $('personalBirthday').value = c.personalBirthday || '';
  $('personalMemo').value     = c.personalMemo     || '';
}

function getPersonalInfo() {
  const drinks  = [...document.querySelectorAll('input[name="drinkPref"]:checked')].map(cb => cb.value).join(',');
  const hobbies = [...document.querySelectorAll('input[name="hobbies"]:checked')].map(cb => cb.value).join(',');
  const diet    = [...document.querySelectorAll('input[name="diet"]:checked')].map(cb => cb.value).join(',');
  return {
    personalDrink:    drinks,
    personalHobbies:  hobbies,
    personalDiet:     diet,
    personalBirthday: $('personalBirthday').value.trim(),
    personalMemo:     $('personalMemo').value.trim()
  };
}

// ── 從既有公司快速新增聯絡人（帶入公司基本資訊）──────────
function openModalAddContact(prefill = {}) {
  $('contactForm').reset();
  $('contactId').value = '';
  $('cardImageUrl').value = '';
  $('previewImg').style.display = 'none';
  $('previewImg').src = '';
  $('cardImagePreview').querySelector('.upload-hint').style.display = 'flex';
  $('imageActions').style.display = 'none';

  $('modalTitle').textContent = '新增聯絡人';

  // 帶入公司資訊，其餘留空
  $('company').value  = prefill.company  || '';
  $('phone').value    = prefill.phone    || '';
  $('ext').value      = ''; // 分機不帶入，由使用者填寫
  $('address').value  = prefill.address  || '';
  $('website').value  = prefill.website  || '';
  $('taxId').value    = prefill.taxId    || '';
  // 帶入產業（從同公司繼承）
  if (prefill.industry) {
    $('industry').value = prefill.industry;
    $('industry').dataset.manual = 'true';
  } else {
    $('industry').value = '';
    $('industry').dataset.manual = 'false';
  }
  // 帶入系統廠商
  if (prefill.systemVendor) {
    $('systemVendor').value = prefill.systemVendor;
    updateSystemProduct(prefill.systemVendor, prefill.systemProduct || '');
  } else {
    updateSystemProduct('');
  }
  $('autoDetectBadge').style.display = 'none';
  if ($('opportunityStage')) $('opportunityStage').value = '';
  $('isPrimaryNo').checked = true;
  $('jobFunction').value = '';
  renderJfSelector('');

  // 職稱、姓名、英文名、手機、Email 留空讓使用者填寫
  $('title').value   = '';
  $('name').value    = '';
  $('nameEn').value  = '';
  $('mobile').value  = '';
  $('email').value   = '';

  $('modalOverlay').classList.add('open');
  $('name').focus();
}

// ── 公司名稱輸入：自動帶入歷史資料 ──────────────────────
let companyFillTimer = null;
function triggerCompanyFill() {
  clearTimeout(companyFillTimer);
  const val = $('company').value.trim();
  if (!val) return;
  companyFillTimer = setTimeout(() => {
    if ($('contactId').value) return; // 編輯模式不觸發
    const match = allContacts.find(c => (c.company || '').trim().toLowerCase() === val.toLowerCase());
    if (!match) return;

    const fill = (id, val) => { if (val && !$(id).value) $(id).value = val; };
    fill('phone',   match.phone);
    fill('ext',     match.ext);
    fill('address', match.address);
    fill('website', match.website);
    fill('taxId',   match.taxId);
    // 產業
    if (match.industry && !$('industry').value) {
      $('industry').value = match.industry;
      $('industry').dataset.manual = 'true';
      $('autoDetectBadge').style.display = 'none';
    }
    // 使用中系統
    if (match.systemVendor && !$('systemVendor').value) {
      $('systemVendor').value = match.systemVendor;
      updateSystemProduct(match.systemVendor, match.systemProduct || '');
    }
    // 同步統編顯示
    if (match.taxId) $('lookupTaxIdDisplay').textContent = match.taxId;

    showToast(`✓ 已帶入「${match.company}」的歷史資料`);
  }, 300);
}
$('company').addEventListener('input',  triggerCompanyFill);
$('company').addEventListener('change', triggerCompanyFill);

// 公司欄位變更時：自動判斷是否為全新公司，若是則預設為主要聯繫窗口
let primaryAutoTimer = null;
$('company').addEventListener('input', function () {
  if ($('contactId').value) return; // 編輯模式不自動覆蓋
  clearTimeout(primaryAutoTimer);
  primaryAutoTimer = setTimeout(() => {
    const val = this.value.trim().toLowerCase();
    if (!val) { $('isPrimaryNo').checked = true; return; }
    const exists = allContacts.some(c => (c.company || '').trim().toLowerCase() === val);
    if (!exists) {
      $('isPrimaryYes').checked = true; // 新公司：自動設為主要
    } else {
      $('isPrimaryNo').checked = true;  // 已有公司：預設否
    }
  }, 350);
});

// ── 公司輸入自動判斷產業 ─────────────────────────────────
let industryAutoTimer = null;
$('company').addEventListener('input', function () {
  clearTimeout(industryAutoTimer);
  industryAutoTimer = setTimeout(() => {
    const manual = $('industry').dataset.manual === 'true';
    if (manual) return;
    const detected = detectIndustry(this.value.trim());
    if (detected) {
      $('industry').value = detected;
      $('autoDetectBadge').style.display = 'inline';
    } else {
      $('autoDetectBadge').style.display = 'none';
    }
  }, 400);
});

$('industry').addEventListener('change', function () {
  this.dataset.manual = 'true';
  $('autoDetectBadge').style.display = 'none';
});

// ── 系統子選單聯動 ───────────────────────────────────────
function updateSystemProduct(vendor, selectedProduct = '') {
  const group = $('systemProductGroup');
  const select = $('systemProduct');
  const text = $('systemProductText');
  const products = SYSTEMS[vendor];

  // 無子選單（文中、正航、空值）
  if (!vendor || (Array.isArray(products) && products.length === 0)) {
    group.style.display = 'none';
    select.style.display = 'select';
    text.style.display = 'none';
    select.innerHTML = '<option value="">-- 請選擇 --</option>';
    text.value = '';
    return;
  }

  group.style.display = 'flex';
  group.style.flexDirection = 'column';

  // 自由輸入（自行開發、其他）
  if (products === null) {
    select.style.display = 'none';
    text.style.display = 'block';
    text.value = selectedProduct;
    return;
  }

  // 下拉選單（鼎新、SAP、Oracle）
  select.style.display = 'block';
  text.style.display = 'none';
  text.value = '';
  select.innerHTML = '<option value="">-- 請選擇 --</option>' +
    products.map(p => `<option value="${p}" ${p === selectedProduct ? 'selected' : ''}>${p}</option>`).join('');
}

$('systemVendor').addEventListener('change', function () {
  updateSystemProduct(this.value);
});

$('addBtn').addEventListener('click', () => openModal());
$('addProspectBtn').addEventListener('click', () => openModal());

// 潛在客戶搜尋
$('prospectSearchInput').addEventListener('input', function () {
  $('prospectClearSearch').classList.toggle('visible', this.value.length > 0);
  loadContacts(this.value);
});
$('prospectClearSearch').addEventListener('click', () => {
  $('prospectSearchInput').value = '';
  $('prospectClearSearch').classList.remove('visible');
  loadContacts();
});
$('modalClose').addEventListener('click', closeModal);

// ── 職稱輸入時自動判斷職能分類 ────────────────────────────
let jfAutoTimer = null;
$('title').addEventListener('input', function () {
  clearTimeout(jfAutoTimer);
  jfAutoTimer = setTimeout(() => {
    // 若已手動選取職能分類則不覆蓋
    if ($('jobFunction').value) return;
    const mapped = autoMapJobFunction(this.value.trim());
    if (mapped) {
      $('jobFunction').value = mapped;
      renderJfSelector(mapped);
    }
  }, 400);
});
$('cancelBtn').addEventListener('click', closeModal);
$('modalOverlay').addEventListener('click', e => { if (e.target === $('modalOverlay')) closeModal(); });

// 名片圖片上傳
$('cardImagePreview').addEventListener('click', () => $('cardImageInput').click());
$('cardImageInput').addEventListener('change', async function () {
  const file = this.files[0];
  if (!file) return;
  const formData = new FormData();
  formData.append('card', file);
  try {
    const res = await fetch(`${API}/upload`, { method: 'POST', body: formData });
    const data = await res.json();
    if (data.url) {
      $('cardImageUrl').value = data.url;
      $('previewImg').src = data.url;
      $('previewImg').style.display = 'block';
      $('cardImagePreview').querySelector('.upload-hint').style.display = 'none';
      $('imageActions').style.display = 'flex';
    }
  } catch { showToast('圖片上傳失敗'); }
  this.value = '';
});

$('removeImage').addEventListener('click', () => {
  $('cardImageUrl').value = '';
  $('previewImg').style.display = 'none';
  $('previewImg').src = '';
  $('cardImagePreview').querySelector('.upload-hint').style.display = 'flex';
  $('imageActions').style.display = 'none';
});

// 儲存
$('saveBtn').addEventListener('click', async () => {
  const name = $('name').value.trim();
  if (!name) { showToast('請輸入姓名'); $('name').focus(); return; }
  const email = $('email').value.trim();
  if (!email) { showToast('請輸入 Email'); $('email').focus(); return; }

  const payload = {
    name,
    nameEn: $('nameEn').value.trim(),
    company: $('company').value.trim(),
    title: $('title').value.trim(),
    phone: $('phone').value.trim(),
    mobile: $('mobile').value.trim(),
    ext: $('ext').value.trim(),
    email: $('email').value.trim(),
    address: $('address').value.trim(),
    website: $('website').value.trim(),
    taxId: $('taxId').value.trim(),
    industry: $('industry').value,
    systemVendor: $('systemVendor').value,
    systemProduct: SYSTEMS[$('systemVendor').value] === null
      ? $('systemProductText').value.trim()
      : $('systemProduct').value,
    opportunityStage: $('opportunityStage') ? $('opportunityStage').value : '',
    isPrimary: $('isPrimaryYes').checked,
    isResigned: $('isResigned').checked,
    note: $('note').value.trim(),
    cardImage: $('cardImageUrl').value,
    jobFunction: $('jobFunction').value,
    ...getPersonalInfo()
  };

  const id = $('contactId').value;
  try {
    let r;
    if (id) {
      r = await fetch(`/api/contacts/${id}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    } else {
      r = await fetch(`/api/contacts`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    }
    if (!r.ok) {
      const err = await r.json().catch(() => ({}));
      throw new Error(err.error || `HTTP ${r.status}`);
    }
    showToast(id ? '已更新聯絡人' : '已新增聯絡人');
    closeModal();
    loadContacts($('searchInput').value.trim());
  } catch(e) { showToast('儲存失敗：' + e.message); }
});

// ── 檢視 Modal ───────────────────────────────────────────
function openView(id) {
  const c = allContacts.find(x => x.id === id);
  if (!c) return;
  currentViewId = id;
  $('viewName').textContent = c.name || '聯絡人';

  const [c1, c2] = avatarColor(c.name);
  let html = '';

  if (c.cardImage) {
    html += `<img class="view-card-image" src="${c.cardImage}" alt="名片圖片">`;
  }

  html += `
    <div class="view-section">
      <div class="view-row"><span class="view-label">姓名</span><span class="view-value">${escapeHtml(c.name) || '-'}${c.nameEn ? `<span style="color:#888;font-size:13px;margin-left:8px">${escapeHtml(c.nameEn)}</span>` : ''}${c.isPrimary ? ' <span class="primary-badge">&#11088; 主要聯繫窗口</span>' : ''}</span></div>
      <div class="view-row"><span class="view-label">公司</span><span class="view-value">${escapeHtml(c.company) || '-'}</span></div>
      <div class="view-row"><span class="view-label">統編</span><span class="view-value">${escapeHtml(c.taxId) || '-'}</span></div>
      <div class="view-row"><span class="view-label">職稱</span><span class="view-value">${escapeHtml(c.title) || '-'}</span></div>
      <div class="view-row"><span class="view-label">職能分類</span><span class="view-value">${(() => { const jfCat = JOB_FUNCTION_CATEGORIES.find(cat => cat.key === (c.jobFunction || autoMapJobFunction(c.title || ''))); return jfCat ? `<span class="jf-badge" style="background:${jfCat.bg};color:${jfCat.color};border-color:${jfCat.color}20">${jfCat.label}<small style="opacity:.7;margin-left:6px">${jfCat.sublabel}</small></span>` : '-'; })()}</span></div>
      <div class="view-row"><span class="view-label">產業</span><span class="view-value">${c.industry ? `<span class="industry-badge">${escapeHtml(c.industry)}</span>` : '-'}</span></div>
      <div class="view-row"><span class="view-label">商機分類</span><span class="view-value">${c.opportunityStage ? renderOpportunityBadge(c.opportunityStage) : '-'}</span></div>
    </div>
    <div class="view-section">
      <div class="view-section-title">聯絡方式</div>
      <div class="view-row"><span class="view-label">電話</span><span class="view-value">${c.phone ? `<a href="tel:${escapeHtml(c.phone)}">${escapeHtml(c.phone)}</a>` : '-'}${c.ext ? `<span style="color:#888;font-size:13px;margin-left:8px">分機 ${escapeHtml(c.ext)}</span>` : ''}</span></div>
      <div class="view-row"><span class="view-label">手機</span><span class="view-value">${c.mobile ? `<a href="tel:${escapeHtml(c.mobile)}">${escapeHtml(c.mobile)}</a>` : '-'}</span></div>
      <div class="view-row"><span class="view-label">Email</span><span class="view-value">${c.email ? `<a href="mailto:${escapeHtml(c.email)}">${escapeHtml(c.email)}</a>` : '-'}</span></div>
      <div class="view-row"><span class="view-label">地址</span><span class="view-value">${escapeHtml(c.address) || '-'}</span></div>
      <div class="view-row"><span class="view-label">網站</span><span class="view-value">${c.website ? `<a href="${safeHref(c.website)}" target="_blank" rel="noopener noreferrer">${escapeHtml(c.website)}</a>` : '-'}</span></div>
    </div>
    ${(c.systemVendor || c.systemProduct) ? `
    <div class="view-section">
      <div class="view-section-title">使用中系統</div>
      ${c.systemVendor ? `<div class="view-row"><span class="view-label">系統</span><span class="view-value">${escapeHtml(c.systemVendor)}</span></div>` : ''}
      ${c.systemProduct ? `<div class="view-row"><span class="view-label">產品</span><span class="view-value system-badge">${escapeHtml(c.systemProduct)}</span></div>` : ''}
    </div>` : ''}
    ${c.note ? `<div class="view-section"><div class="view-section-title">備註</div><div class="view-row"><span class="view-value" style="white-space:pre-wrap">${escapeHtml(c.note)}</span></div></div>` : ''}`;

  $('viewBody').innerHTML = html;
  $('viewOverlay').classList.add('open');
}

$('viewClose').addEventListener('click', () => $('viewOverlay').classList.remove('open'));
$('viewOverlay').addEventListener('click', e => { if (e.target === $('viewOverlay')) $('viewOverlay').classList.remove('open'); });

$('viewEditBtn').addEventListener('click', () => {
  $('viewOverlay').classList.remove('open');
  const c = allContacts.find(x => x.id === currentViewId);
  if (c) openModal(c);
});

// ── 刪除 ─────────────────────────────────────────────────
$('viewDeleteBtn').addEventListener('click', () => {
  const c = allContacts.find(x => x.id === currentViewId);
  if (!c) return;
  $('deleteContactName').textContent = c.name;
  $('viewOverlay').classList.remove('open');
  $('confirmOverlay').classList.add('open');
});

$('confirmCancel').addEventListener('click', () => $('confirmOverlay').classList.remove('open'));
$('confirmOverlay').addEventListener('click', e => { if (e.target === $('confirmOverlay')) $('confirmOverlay').classList.remove('open'); });

$('confirmDelete').addEventListener('click', async () => {
  try {
    await fetch(`${API}/contacts/${currentViewId}`, { method: 'DELETE' });
    showToast('已刪除聯絡人');
    $('confirmOverlay').classList.remove('open');
    loadContacts($('searchInput').value.trim());
  } catch { showToast('刪除失敗，請重試'); }
});

// ── 匯出 Excel ───────────────────────────────────────────
$('exportBtn').addEventListener('click', async () => {
  // 先重新取得最新權限
  await loadPermissions();
  if (!userPermissions.canDownloadContacts && userPermissions.role !== 'admin') {
    showToast('⛔ 您沒有下載客戶名單的權限，請聯繫管理者');
    return;
  }
  try {
    const r = await fetch(`${API}/export`);
    if (!r.ok) {
      let errMsg = '⛔ 無法下載客戶名單';
      try { const j = await r.json(); errMsg = '⛔ ' + (j.error || errMsg); } catch {}
      showToast(errMsg, 3000);
      // 重新載入最新權限並套用
      await loadPermissions();
      return;
    }
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const today = new Date();
    const ds = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}${String(today.getDate()).padStart(2,'0')}`;
    a.href = url;
    a.download = `客戶名單_${ds}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  } catch {
    showToast('⛔ 下載失敗，請重試');
  }
});

// ── 登入者資訊 & 登出 ────────────────────────────────────
window._myRole = 'user'; // 全域存角色，供其他函式判斷
async function initUser() {
  try {
    const res = await fetch('/api/me');
    if (res.status === 401) { window.location.href = '/login.html'; return; }
    const user = await res.json();
    window._myRole = user.role || 'user';
    $('sidebarUser').textContent = user.displayName;
    const ROLE_LABEL = { admin:'管理者', manager1:'一級主管', manager2:'二級主管', secretary:'秘書', user:'' };
    const roleLabel = ROLE_LABEL[user.role] || '';
    if (roleLabel) {
      const roleEl = document.createElement('div');
      roleEl.style.cssText = 'font-size:10px;color:rgba(255,255,255,0.55);margin-top:2px;';
      roleEl.textContent = roleLabel;
      $('sidebarUser').parentNode.appendChild(roleEl);
    }
    const h = new Date().getHours();
    const greet = h < 12 ? '早安' : h < 18 ? '午安' : '晚安';
    $('dashboardGreeting').textContent = greet + '，' + user.displayName + ' 👋';
  } catch { window.location.href = '/login.html'; }
}

$('logoutBtn').addEventListener('click', async () => {
  await fetch('/api/logout', { method: 'POST' });
  window.location.href = '/login.html';
});

// ── 更改密碼 ─────────────────────────────────────────────
function openChangePw() {
  $('changePwOld').value     = '';
  $('changePwNew').value     = '';
  $('changePwConfirm').value = '';
  $('changePwError').style.display = 'none';
  $('changePwOverlay').classList.add('open');
  setTimeout(() => $('changePwOld').focus(), 60);
}
function closeChangePw() {
  $('changePwOverlay').classList.remove('open');
}

$('changePwBtn').addEventListener('click', openChangePw);
$('changePwClose').addEventListener('click', closeChangePw);
$('changePwCancel').addEventListener('click', closeChangePw);
$('changePwOverlay').addEventListener('click', e => { if (e.target === $('changePwOverlay')) closeChangePw(); });

// Enter 鍵送出
['changePwOld','changePwNew','changePwConfirm'].forEach(id => {
  $(id).addEventListener('keydown', e => { if (e.key === 'Enter') $('changePwSubmit').click(); });
});

$('changePwSubmit').addEventListener('click', async () => {
  const errEl  = $('changePwError');
  const oldPw  = $('changePwOld').value.trim();
  const newPw  = $('changePwNew').value;
  const confPw = $('changePwConfirm').value;
  const btn    = $('changePwSubmit');

  const showErr = msg => { errEl.textContent = msg; errEl.style.display = ''; };
  errEl.style.display = 'none';

  if (!oldPw)           return showErr('請輸入舊密碼');
  if (!newPw)           return showErr('請輸入新密碼');
  if (newPw.length < 6) return showErr('新密碼至少需要 6 個字元');
  if (newPw !== confPw) return showErr('兩次輸入的新密碼不一致');

  btn.disabled = true;
  btn.textContent = '更新中...';
  try {
    const r = await fetch(`${API}/user/password`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ currentPassword: oldPw, newPassword: newPw })
    });
    const data = await r.json();
    if (!r.ok) return showErr(data.error || '更改失敗，請稍後再試');
    closeChangePw();
    showToast('✅ 密碼已成功更新');
  } catch {
    showErr('網路錯誤，請稍後再試');
  } finally {
    btn.disabled = false;
    btn.textContent = '確認更改';
  }
});

// ── 拜訪記錄 ─────────────────────────────────────────────
let allVisits = [];
let currentVisitId = null;

function $v(id) { return document.getElementById(id); }

// 載入商機
async function loadOpportunities() {
  try {
    const res = await fetch(`${API}/opportunities`);
    allOpportunities = await res.json();
  } catch { /* ignore */ }
}

// ── 業務日報日曆 ─────────────────────────────────────────
let calYear  = new Date().getFullYear();
let calMonth = new Date().getMonth(); // 0-based

async function loadVisits() {
  try {
    const res = await fetch(`${API}/visits`);
    allVisits = await res.json();
    renderCalendar();
  } catch { showToast('無法載入拜訪記錄'); }
}

// 保留供相容性（舊地方呼叫 renderVisits 時改為呼叫 renderCalendar）
function renderVisits() { renderCalendar(); }

function renderCalendar() {
  const today = new Date();
  const firstDay = new Date(calYear, calMonth, 1);
  const lastDay  = new Date(calYear, calMonth + 1, 0);
  const startDow = firstDay.getDay(); // 0=Sun

  // 月份標題
  $('calMonthTitle').textContent =
    `${calYear} 年 ${calMonth + 1} 月`;

  // 建立日期→拜訪 map
  const visitMap = {};
  allVisits.forEach(v => {
    if (!v.visitDate) return;
    (visitMap[v.visitDate] = visitMap[v.visitDate] || []).push(v);
  });

  const grid = $('calGrid');
  grid.innerHTML = '';

  // 前置空格（上月尾日）
  for (let i = 0; i < startDow; i++) {
    const prev = new Date(calYear, calMonth, -startDow + i + 1);
    grid.appendChild(makeCell(prev, visitMap, today, true));
  }
  // 本月
  for (let d = 1; d <= lastDay.getDate(); d++) {
    const date = new Date(calYear, calMonth, d);
    grid.appendChild(makeCell(date, visitMap, today, false));
  }
  // 後置空格（下月初）
  const totalCells = startDow + lastDay.getDate();
  const remainder  = totalCells % 7 === 0 ? 0 : 7 - (totalCells % 7);
  for (let i = 1; i <= remainder; i++) {
    const next = new Date(calYear, calMonth + 1, i);
    grid.appendChild(makeCell(next, visitMap, today, true));
  }
}

function makeCell(date, visitMap, today, otherMonth) {
  const cell = document.createElement('div');
  const dow  = date.getDay();
  const dateStr = `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`;
  const isToday = dateStr === `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;

  let cls = 'cal-cell';
  if (otherMonth) cls += ' other-month';
  if (isToday)    cls += ' today';
  if (dow === 0)  cls += ' sunday';
  if (dow === 6)  cls += ' saturday';
  cell.className = cls;

  // 日期數字
  const dateDiv = document.createElement('div');
  dateDiv.className = 'cal-date';
  dateDiv.textContent = date.getDate();
  cell.appendChild(dateDiv);

  // 拜訪記錄
  const visitsDiv = document.createElement('div');
  visitsDiv.className = 'cal-visits';
  const dayVisits = visitMap[dateStr] || [];
  const MAX_SHOW = 3;
  dayVisits.slice(0, MAX_SHOW).forEach(v => {
    const contact = allContacts.find(c => c.id === v.contactId);
    const name = contact ? contact.name : v.contactName || '';
    const company = contact?.company || '';
    const vt = v.visitType || '親訪';
    const badge = document.createElement('div');
    badge.className = `cal-visit-badge vt-${vt}`;
    badge.title = `${vt}｜${name}${company ? '・' + company : ''}\n${v.topic || ''}`;
    badge.innerHTML = `<span class="badge-type">${vt}</span>${name}${company ? '・' + company : ''}`;
    badge.addEventListener('click', e => { e.stopPropagation(); openVisitView(v.id); });
    visitsDiv.appendChild(badge);
  });
  if (dayVisits.length > MAX_SHOW) {
    const more = document.createElement('div');
    more.className = 'cal-more';
    more.textContent = `+${dayVisits.length - MAX_SHOW} 筆`;
    more.addEventListener('click', e => { e.stopPropagation(); openAllVisitsOfDay(dateStr, dayVisits); });
    visitsDiv.appendChild(more);
  }
  cell.appendChild(visitsDiv);

  // + 新增按鈕
  const addBtn = document.createElement('button');
  addBtn.className = 'cal-add-btn';
  addBtn.textContent = '+';
  addBtn.title = '新增此日拜訪記錄';
  addBtn.addEventListener('click', e => {
    e.stopPropagation();
    openVisitModalOnDate(dateStr);
  });
  cell.appendChild(addBtn);

  return cell;
}

// 點格子內 +，帶入指定日期開啟 Modal
function openVisitModalOnDate(dateStr) {
  if (allContacts.length === 0) {
    loadContacts().then(() => { openVisitModal(); $v('visitDate').value = dateStr; });
  } else {
    openVisitModal();
    $v('visitDate').value = dateStr;
  }
}

// 超過 3 筆：展示當日所有記錄（重用 openVisitView 流程，先用 toast 提示）
function openAllVisitsOfDay(dateStr, visits) {
  // 簡單：把當天的用一個輕量 overlay 列表展示
  // 用現有 visitViewOverlay 方式，點擊各筆再開詳情
  const list = visits.map(v => {
    const contact = allContacts.find(c => c.id === v.contactId);
    const name = contact ? contact.name : v.contactName || '';
    return `<div class="day-visit-row" data-id="${v.id}">
      <span class="visit-type-badge">${v.visitType||'親訪'}</span>
      <span style="font-weight:600;margin:0 6px">${name}</span>
      <span style="color:#555">${v.topic||''}</span>
    </div>`;
  }).join('');
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay open';
  overlay.innerHTML = `<div class="modal modal-sm" style="width:420px">
    <div class="modal-header"><h2>${dateStr} 拜訪記錄</h2>
      <button class="modal-close" id="dayListClose">✕</button></div>
    <div class="modal-body" style="display:flex;flex-direction:column;gap:10px">${list}</div>
  </div>`;
  document.body.appendChild(overlay);
  overlay.querySelector('#dayListClose').addEventListener('click', () => overlay.remove());
  overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });
  overlay.querySelectorAll('.day-visit-row').forEach(row => {
    row.style.cursor = 'pointer';
    row.style.padding = '8px 10px';
    row.style.borderRadius = '8px';
    row.style.transition = 'background .15s';
    row.addEventListener('mouseenter', () => row.style.background = '#f0f6ff');
    row.addEventListener('mouseleave', () => row.style.background = '');
    row.addEventListener('click', () => { overlay.remove(); openVisitView(row.dataset.id); });
  });
}

// 日曆導覽
$('calPrevBtn').addEventListener('click', () => {
  calMonth--;
  if (calMonth < 0) { calMonth = 11; calYear--; }
  renderCalendar();
});
$('calNextBtn').addEventListener('click', () => {
  calMonth++;
  if (calMonth > 11) { calMonth = 0; calYear++; }
  renderCalendar();
});
$('calTodayBtn').addEventListener('click', () => {
  calYear  = new Date().getFullYear();
  calMonth = new Date().getMonth();
  renderCalendar();
});

// 填充客戶公司下拉選單
function populateCompanySelect(selectedCompany = '') {
  const sel = $v('visitCompany');
  sel.innerHTML = '<option value="">-- 請選擇客戶 --</option>';
  const companies = [...new Set(allContacts.map(c => c.company).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'zh-TW'));
  companies.forEach(name => {
    const opt = document.createElement('option');
    opt.value = name;
    opt.textContent = name;
    if (name === selectedCompany) opt.selected = true;
    sel.appendChild(opt);
  });
}

// 依公司填充聯絡人下拉選單
function populateContactSelect(company = '', selectedContactId = '') {
  const sel = $v('visitContactId');
  if (!company) {
    sel.innerHTML = '<option value="">-- 請先選擇客戶 --</option>';
    sel.disabled = true;
    return;
  }
  const contacts = allContacts.filter(c => c.company === company);
  sel.innerHTML = '<option value="">-- 請選擇聯絡人 --</option>';
  contacts.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'zh-TW')).forEach(c => {
    const opt = document.createElement('option');
    opt.value = c.id;
    opt.textContent = c.name + (c.title ? `（${c.title}）` : '');
    if (c.id === selectedContactId) opt.selected = true;
    sel.appendChild(opt);
  });
  sel.disabled = false;
  // 若該公司只有一位聯絡人，自動選取
  if (contacts.length === 1) sel.value = contacts[0].id;
}

// 公司選單變動時聯動聯絡人選單
$v('visitCompany').addEventListener('change', function () {
  populateContactSelect(this.value);
});

// ── 商機類別 → 商品選單 ──────────────────────────────────
const OPP_PRODUCTS = {
  ERP: {
    '一般商品': [
      'SAP Public Cloud License',
      'SAP Private Cloud License',
      'ERP 顧問導入專案 PE',
      'ERP 顧問導入專案 PCE',
      '其他',
    ],
    '─── License MA ───': [
      'SAP PCOE License MA',
      'SAP PCE License MA',
      'SAP PE License MA',
    ],
    '─── Service MA ───': [
      'Service MA（Basis）',
      'Service MA（AP）',
      'Service MA（Basis & AP）',
    ],
    '─── SAP 延伸解決方案 ───': [
      'SAP CRM',
      'SAP Analytics Cloud（SAC）',
      'SAP Customer Data Platform（CDP）',
      'SAP Engagement Cloud（Emarsys）',
    ],
  },
  ITS: {
    'ITS 產品': [
      'MES',
      'WMS',
      'ESG',
      'DMS',
    ],
  }
};

$v('oppCategory').addEventListener('change', function () {
  const cat = this.value;
  const pg = $v('oppProductGroup');
  const ps = $v('oppProduct');
  ps.innerHTML = '<option value="">-- 請選擇 --</option>';
  const catData = OPP_PRODUCTS[cat];
  const groups = catData && typeof catData === 'object' && !Array.isArray(catData) ? catData : null;
  const hasItems = groups && Object.values(groups).some(arr => arr.length > 0);
  if (hasItems) {
    Object.entries(groups).forEach(([groupLabel, items]) => {
      if (!items.length) return;
      const grp = document.createElement('optgroup');
      grp.label = groupLabel;
      items.forEach(p => {
        const o = document.createElement('option');
        o.value = p; o.textContent = p;
        grp.appendChild(o);
      });
      ps.appendChild(grp);
    });
    pg.style.display = '';
  } else {
    pg.style.display = 'none';
  }
});

// ── 拜訪 Modal Tab 切換 ───────────────────────────────────
document.querySelectorAll('[data-vtab]').forEach(btn => {
  btn.addEventListener('click', function () {
    document.querySelectorAll('[data-vtab]').forEach(b => b.classList.remove('active'));
    this.classList.add('active');
    const target = this.dataset.vtab;
    ['vtab-visit', 'vtab-opportunity'].forEach(id => {
      $v(id).style.display = id === target ? '' : 'none';
    });
  });
});

function resetVisitModalTabs() {
  document.querySelectorAll('[data-vtab]').forEach(b => b.classList.remove('active'));
  const first = document.querySelector('[data-vtab="vtab-visit"]');
  if (first) first.classList.add('active');
  $v('vtab-visit').style.display = '';
  $v('vtab-opportunity').style.display = 'none';
  // reset opportunity fields
  $v('oppCategory').value = '';
  $v('oppProduct').innerHTML = '<option value="">-- 請選擇 --</option>';
  $v('oppProductGroup').style.display = 'none';
  $v('oppAmount').value = '';
  $v('oppExpectedDate').value = '';
  $v('oppDescription').value = '';
}

// 開啟新增/編輯 Modal
function openVisitModal(visit = null) {
  $v('visitForm').reset();
  $v('visitId').value = '';
  $v('visitDate').value = new Date().toISOString().slice(0, 10);
  resetVisitModalTabs();

  if (visit) {
    const contact = allContacts.find(c => c.id === visit.contactId);
    const company = contact?.company || '';
    populateCompanySelect(company);
    populateContactSelect(company, visit.contactId || '');
    $v('visitModalTitle').textContent = '編輯拜訪記錄';
    $v('visitId').value = visit.id;
    $v('visitDate').value = visit.visitDate || '';
    $v('visitType').value = visit.visitType || '親訪';
    $v('visitTopic').value = visit.topic || '';
    $v('visitContent').value = visit.content || '';
    $v('visitNextAction').value = visit.nextAction || '';
  } else {
    populateCompanySelect();
    populateContactSelect();
    $v('visitModalTitle').textContent = '新增拜訪記錄';
  }
  $v('visitModalOverlay').classList.add('open');
}

function closeVisitModal() { $v('visitModalOverlay').classList.remove('open'); }

// 儲存拜訪記錄（+ 商機）
$v('visitSaveBtn').addEventListener('click', async () => {
  const contactId = $v('visitContactId').value;
  const topic = $v('visitTopic').value.trim();
  const visitDate = $v('visitDate').value;
  if (!contactId) { showToast('請選擇聯絡人'); return; }
  if (!visitDate) { showToast('請填入拜訪日期'); return; }
  if (!topic)     { showToast('請填入拜訪主題'); return; }

  const contact = allContacts.find(c => c.id === contactId);
  const visitPayload = {
    contactId,
    contactName: contact ? contact.name : '',
    visitDate,
    visitType: $v('visitType').value,
    topic,
    content: $v('visitContent').value.trim(),
    nextAction: $v('visitNextAction').value.trim(),
  };

  const id = $v('visitId').value;
  try {
    let visitId = id;
    if (id) {
      await fetch(`${API}/visits/${id}`, { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(visitPayload) });
    } else {
      const r = await fetch(`${API}/visits`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(visitPayload) });
      const saved = await r.json();
      visitId = saved.id;
    }

    // 若商機類別有填 → 同步建立商機
    const oppCat = $v('oppCategory').value;
    if (oppCat) {
      const oppPayload = {
        contactId,
        contactName: contact ? contact.name : '',
        company: contact ? (contact.company || '') : '',
        category: oppCat,
        product: getProductValue('oppProduct', 'oppProductCustom'),
        amount: $v('oppAmount').value,
        expectedDate: $v('oppExpectedDate').value,
        description: $v('oppDescription').value.trim(),
        visitId,
      };
      await fetch(`${API}/opportunities`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(oppPayload) });
      await loadOpportunities();
      updateStatCards();
      showToast(id ? '已更新拜訪記錄' : '已新增拜訪記錄 及 商機（Pipeline）');
    } else {
      showToast(id ? '已更新拜訪記錄' : '已新增拜訪記錄');
    }

    closeVisitModal();
    loadVisits();
    if (currentSection === null) renderDashboardCharts();
  } catch { showToast('儲存失敗，請重試'); }
});

$v('addVisitBtn').addEventListener('click', () => {
  if (allContacts.length === 0) loadContacts().then(openVisitModal);
  else openVisitModal();
});
$v('visitModalClose').addEventListener('click', closeVisitModal);
$v('visitCancelBtn').addEventListener('click', closeVisitModal);
$v('visitModalOverlay').addEventListener('click', e => { if (e.target === $v('visitModalOverlay')) closeVisitModal(); });

// 檢視拜訪記錄
function openVisitView(id) {
  const v = allVisits.find(x => x.id === id);
  if (!v) return;
  currentVisitId = id;
  const contact = allContacts.find(c => c.id === v.contactId);
  const contactLabel = contact
    ? `${escapeHtml(contact.name)}${contact.company ? '<span style="color:#888;font-size:13px;margin-left:6px">' + escapeHtml(contact.company) + '</span>' : ''}`
    : escapeHtml(v.contactName) || '-';
  $v('visitViewTitle').textContent = v.topic || '拜訪記錄';
  $v('visitViewBody').innerHTML = `
    <div class="view-section">
      <div class="view-row"><span class="view-label">日期</span><span class="view-value">${escapeHtml(v.visitDate) || '-'}</span></div>
      <div class="view-row"><span class="view-label">方式</span><span class="view-value">${escapeHtml(v.visitType) || '-'}</span></div>
      <div class="view-row"><span class="view-label">聯絡人</span><span class="view-value">${contactLabel}</span></div>
      <div class="view-row"><span class="view-label">主題</span><span class="view-value" style="font-weight:700">${escapeHtml(v.topic) || '-'}</span></div>
    </div>
    ${v.content ? `<div class="view-section"><div class="view-section-title">會談內容</div><div style="font-size:14px;color:#333;line-height:1.7;white-space:pre-wrap">${escapeHtml(v.content)}</div></div>` : ''}
    ${v.nextAction ? `<div class="view-section"><div class="view-section-title">下一步行動</div><div style="font-size:14px;color:#34a853;font-weight:600;line-height:1.7;white-space:pre-wrap">${escapeHtml(v.nextAction)}</div></div>` : ''}`;
  $v('visitViewOverlay').classList.add('open');
}

$v('visitViewClose').addEventListener('click', () => $v('visitViewOverlay').classList.remove('open'));
$v('visitViewOverlay').addEventListener('click', e => { if (e.target === $v('visitViewOverlay')) $v('visitViewOverlay').classList.remove('open'); });
$v('visitViewEditBtn').addEventListener('click', () => {
  $v('visitViewOverlay').classList.remove('open');
  const v = allVisits.find(x => x.id === currentVisitId);
  if (v) openVisitModal(v);
});
$v('visitViewDeleteBtn').addEventListener('click', () => {
  $v('visitViewOverlay').classList.remove('open');
  $v('visitConfirmOverlay').classList.add('open');
});
$v('visitConfirmCancel').addEventListener('click', () => $v('visitConfirmOverlay').classList.remove('open'));
$v('visitConfirmOverlay').addEventListener('click', e => { if (e.target === $v('visitConfirmOverlay')) $v('visitConfirmOverlay').classList.remove('open'); });
$v('visitConfirmDelete').addEventListener('click', async () => {
  try {
    await fetch(`${API}/visits/${currentVisitId}`, { method: 'DELETE' });
    showToast('已刪除拜訪記錄');
    $v('visitConfirmOverlay').classList.remove('open');
    loadVisits();
  } catch { showToast('刪除失敗，請重試'); }
});

// ── 商機推進看板（Kanban）────────────────────────────────
const KANBAN_STAGES = [
  { key:'D',   label:'D｜靜止中',  color:'#9e9e9e' },
  { key:'C',   label:'C｜Pipeline',color:'#1a73e8' },
  { key:'B',   label:'B｜Upside',  color:'#f57c00' },
  { key:'A',   label:'A｜Commit',  color:'#d32f2f' },
  { key:'Won', label:'🏆 Won',     color:'#1e8e3e' },
];

let dragOppId = null;

async function loadPipelineView() {
  await loadOpportunities();
  renderKanban();
}

function renderKanban() {
  const activeOpps = allOpportunities; // 所有階段（含 Won）都在看板顯示

  KANBAN_STAGES.forEach(({ key }) => {
    const stageOpps = activeOpps.filter(o => o.stage === key);
    const total = stageOpps.reduce((s, o) => s + (parseFloat(o.amount) || 0), 0);

    $('kanbanBadge' + key).textContent = stageOpps.length;
    $('kanbanTotal' + key).innerHTML =
      '$' + total.toLocaleString() + '<span class="kanban-col-total-label">萬</span>';

    const container = $('kanbanCards' + key);
    container.innerHTML = '';

    if (stageOpps.length === 0) {
      container.innerHTML = '<div class="kanban-empty-col">尚無商機</div>';
    } else {
      stageOpps
        .sort((a, b) => (parseFloat(b.amount)||0) - (parseFloat(a.amount)||0))
        .forEach(o => container.appendChild(buildKanbanCard(o)));
    }

    // Drag-and-drop 事件（用 AbortController 避免 renderKanban 多次呼叫疊加監聽器）
    if (container._dndController) container._dndController.abort();
    container._dndController = new AbortController();
    const sig = { signal: container._dndController.signal };

    container.addEventListener('dragover', e => {
      e.preventDefault();
      container.classList.add('drag-over');
    }, sig);
    container.addEventListener('dragleave', () => container.classList.remove('drag-over'), sig);
    container.addEventListener('drop', async e => {
      e.preventDefault();
      container.classList.remove('drag-over');
      if (!dragOppId) return;
      const newStage = container.dataset.stage;
      const opp = allOpportunities.find(x => x.id === dragOppId);
      if (!opp || opp.stage === newStage) return;
      try {
        const body = { stage: newStage };
        if (newStage === 'Won' && !opp.achievedDate) {
          body.achievedDate = new Date().toISOString().slice(0, 10);
        }
        const r = await fetch(`${API}/opportunities/${dragOppId}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(body)
        });
        if (!r.ok) {
          const err = await r.json().catch(() => ({}));
          showToast('❌ 更新失敗：' + (err.error || r.status));
          return;
        }
        // 先在記憶體直接更新，保證 renderKanban 立即看到新狀態
        Object.assign(opp, body);
        renderKanban();
        updateStatCards();
        renderDashboardCharts();
        if (newStage === 'Won') {
          celebrateWon(opp);
        } else {
          const stageLabel = KANBAN_STAGES.find(s => s.key === newStage)?.label.split('｜')[1] || newStage;
          showToast(`✅ 已移至 ${newStage}｜${stageLabel}`);
        }
        // 背景同步，確保資料一致
        loadOpportunities();
      } catch (err) { showToast('❌ 網路錯誤，請重試'); }
    }, sig);
  });
}

// ── 簽約慶祝動畫 ─────────────────────────────────────────
function celebrateWon(opp) {
  const overlay = document.getElementById('wonCelebration');
  const card    = document.getElementById('wonCelebCard');
  const canvas  = document.getElementById('wonConfettiCanvas');
  if (!overlay || !card || !canvas) return;

  // 填入商機資訊
  document.getElementById('wonCelebCompany').textContent = opp.company || '';
  const amt = opp.amount ? `合約金額：$${Number(opp.amount).toLocaleString()} 萬` : '';
  document.getElementById('wonCelebAmount').textContent = amt;

  // 顯示 overlay
  overlay.style.display = '';
  card.className = '';
  void card.offsetWidth; // reflow
  card.className = 'won-celeb-in';

  // Canvas confetti
  const ctx = canvas.getContext('2d');
  canvas.width  = window.innerWidth;
  canvas.height = window.innerHeight;
  const COLORS = ['#1e8e3e','#34a853','#fbbc04','#ea4335','#4285f4','#ff6d00','#ab47bc'];
  const pieces = Array.from({ length: 120 }, () => ({
    x: Math.random() * canvas.width,
    y: Math.random() * canvas.height - canvas.height,
    w: 8 + Math.random() * 10,
    h: 14 + Math.random() * 10,
    color: COLORS[Math.floor(Math.random() * COLORS.length)],
    rot: Math.random() * 360,
    vx: (Math.random() - .5) * 3,
    vy: 3 + Math.random() * 5,
    vr: (Math.random() - .5) * 8,
    alpha: 1,
  }));

  let frame, start = null;
  const DURATION = 3200; // ms

  function draw(ts) {
    if (!start) start = ts;
    const elapsed = ts - start;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    pieces.forEach(p => {
      p.x  += p.vx;
      p.y  += p.vy;
      p.vy += 0.12; // gravity
      p.rot += p.vr;
      if (elapsed > DURATION * .6) p.alpha = Math.max(0, p.alpha - .02);
      ctx.save();
      ctx.globalAlpha = p.alpha;
      ctx.translate(p.x + p.w / 2, p.y + p.h / 2);
      ctx.rotate(p.rot * Math.PI / 180);
      ctx.fillStyle = p.color;
      ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h);
      ctx.restore();
    });
    if (elapsed < DURATION) {
      frame = requestAnimationFrame(draw);
    } else {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      // 關閉彈窗
      card.className = 'won-celeb-out';
      setTimeout(() => { overlay.style.display = 'none'; card.className = ''; }, 400);
    }
  }
  cancelAnimationFrame(frame);
  frame = requestAnimationFrame(draw);

  // 點擊提早關閉
  overlay.style.pointerEvents = 'auto';
  overlay.onclick = () => {
    cancelAnimationFrame(frame);
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    card.className = 'won-celeb-out';
    setTimeout(() => { overlay.style.display = 'none'; card.className = ''; overlay.style.pointerEvents = 'none'; overlay.onclick = null; }, 400);
  };
}

function buildKanbanCard(o) {
  const card = document.createElement('div');
  const isWon = o.stage === 'Won';
  card.className = 'kanban-card' + (isWon ? ' kanban-card-won' : '');
  card.draggable = !isWon; // Won 卡片不可再拖移
  card.dataset.id = o.id;

  const amt  = o.amount ? '$' + Number(o.amount).toLocaleString() : '金額未填';
  const cat  = o.category ? `<span class="kanban-card-cat">${escapeHtml(o.category)}</span>` : '';
  const prod = o.product  ? `<div class="kanban-card-product" title="${escapeHtml(o.product)}">${escapeHtml(o.product)}</div>` : '';
  const amtStyle = o.amount ? '' : 'color:#bbb;font-weight:400';

  // Won 顯示成交日；其他顯示預計結案日
  const dateLabel = isWon
    ? (o.achievedDate ? `🏆 ${o.achievedDate}` : '🏆 已成交')
    : (o.expectedDate || '');

  card.innerHTML = `
    <div class="kanban-card-company">${escapeHtml(o.company) || '（未填公司）'}</div>
    <div class="kanban-card-contact">${escapeHtml(o.contactName) || ''}</div>
    ${cat}${prod}
    <div class="kanban-card-footer">
      <span class="kanban-card-amount" style="${amtStyle}">${escapeHtml(amt)} 萬</span>
      <span class="kanban-card-date">${escapeHtml(dateLabel)}</span>
    </div>
    <div class="kanban-card-edit-hint">點擊編輯</div>`;

  if (!isWon) {
    card.addEventListener('dragstart', () => {
      dragOppId = o.id;
      setTimeout(() => card.classList.add('dragging'), 0);
    });
    card.addEventListener('dragend', () => {
      dragOppId = null;
      card.classList.remove('dragging');
    });
  }
  // 點擊開啟編輯（拖曳時不觸發）
  card.addEventListener('click', () => {
    if (dragOppId) return;
    openOppEdit(o.id);
  });

  return card;
}

// ── 商機編輯 Modal ────────────────────────────────────────
function openOppEdit(id) {
  const o = allOpportunities.find(x => x.id === id);
  if (!o) return;

  $('oppEditId').value             = o.id;
  $('oppEditCompany').value        = o.company         || '';
  $('oppEditContact').value        = o.contactName     || '';
  $('oppEditCategory').value       = o.category        || '';
  $('oppEditStage').value          = o.stage           || 'C';
  $('oppEditAmount').value         = o.amount          || '';
  $('oppEditGrossMargin').value    = o.grossMarginRate !== undefined ? o.grossMarginRate : '';
  $('oppEditExpectedDate').value   = o.expectedDate    || '';
  $('oppEditDescription').value    = o.description     || '';

  // 商品選單
  populateOppEditProduct(o.category, o.product);

  // 重置 AI 贏率面板
  const _panel = $('oppWinRatePanel');
  if (_panel) _panel.style.display = 'none';

  $('oppEditOverlay').classList.add('open');
}

function populateOppEditProduct(category, selected = '') {
  const ps = $('oppEditProduct');
  const pg = $('oppEditProductGroup');
  ps.innerHTML = '<option value="">-- 請選擇 --</option>';
  const catData = OPP_PRODUCTS[category];
  const groups = catData && typeof catData === 'object' && !Array.isArray(catData) ? catData : null;
  const hasItems = groups && Object.values(groups).some(arr => arr.length > 0);
  if (hasItems) {
    Object.entries(groups).forEach(([groupLabel, items]) => {
      if (!items.length) return;
      const grp = document.createElement('optgroup');
      grp.label = groupLabel;
      items.forEach(p => {
        const o = document.createElement('option');
        o.value = p; o.textContent = p;
        if (p === selected) o.selected = true;
        grp.appendChild(o);
      });
      ps.appendChild(grp);
    });
    pg.style.display = '';
    // 若 selected 不在清單中（自訂商機名稱）→ 選「其他」並填入自訂欄
    const allValues = groups ? Object.values(groups).flat() : [];
    const customInp = $('oppEditProductCustom');
    if (selected && selected !== '其他' && !allValues.includes(selected)) {
      ps.value = '其他';
      if (customInp) { customInp.style.display = ''; customInp.value = selected; }
    } else {
      if (customInp) { customInp.style.display = ps.value === '其他' ? '' : 'none'; customInp.value = ''; }
    }
  } else {
    pg.style.display = 'none';
  }
}

$('oppEditCategory').addEventListener('change', function () {
  $('oppEditProductCustom').style.display = 'none';
  $('oppEditProductCustom').value = '';
  populateOppEditProduct(this.value);
});

function closeOppEdit() { $('oppEditOverlay').classList.remove('open'); }
$('oppEditClose').addEventListener('click', closeOppEdit);
$('oppEditCancel').addEventListener('click', closeOppEdit);
$('oppEditOverlay').addEventListener('click', e => { if (e.target === $('oppEditOverlay')) closeOppEdit(); });

// 刪除案件 → 開啟填寫原因 Modal
$('oppEditDelete').addEventListener('click', () => {
  const company = $('oppEditCompany').value || '';
  const product = getProductValue('oppEditProduct', 'oppEditProductCustom') ||
                  $('oppEditProduct').value || '';
  $('oppDeleteCompany').value = company;
  $('oppDeleteProduct').value = product;
  $('oppDeleteReasonSel').value = '';
  $('oppDeleteReasonNote').value = '';
  $('oppDeleteModal').classList.add('open');
});

function closeOppDeleteModal() {
  $('oppDeleteModal').classList.remove('open');
}
$('oppDeleteClose').addEventListener('click', closeOppDeleteModal);
$('oppDeleteCancel').addEventListener('click', closeOppDeleteModal);
$('oppDeleteModal').addEventListener('click', e => {
  if (e.target === $('oppDeleteModal')) closeOppDeleteModal();
});

$('oppDeleteConfirm').addEventListener('click', async () => {
  const reasonSel  = $('oppDeleteReasonSel').value.trim();
  const reasonNote = $('oppDeleteReasonNote').value.trim();
  if (!reasonSel) { alert('請選擇刪除原因'); return; }
  const deleteReason = reasonNote ? `${reasonSel}：${reasonNote}` : reasonSel;
  const id = $('oppEditId').value;
  try {
    const r = await fetch(`/api/opportunities/${id}`, {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ deleteReason })
    });
    if (!r.ok) throw new Error((await r.json()).error || '刪除失敗');
    closeOppDeleteModal();
    closeOppEdit();
    await loadOpportunities();
    renderForecastTable();
  } catch (e) { alert('刪除失敗：' + e.message); }
});

$('oppEditSave').addEventListener('click', async () => {
  const id = $('oppEditId').value;
  const newStage = $('oppEditStage').value;
  const gmRaw = $('oppEditGrossMargin').value;
  const payload = {
    category:        $('oppEditCategory').value,
    product:         getProductValue('oppEditProduct', 'oppEditProductCustom'),
    stage:           newStage,
    amount:          $('oppEditAmount').value,
    grossMarginRate: gmRaw !== '' ? parseFloat(gmRaw) : '',
    expectedDate:    $('oppEditExpectedDate').value,
    description:     $('oppEditDescription').value.trim(),
  };
  if (newStage === 'Won' && !payload.achievedDate) {
    payload.achievedDate = new Date().toISOString().slice(0, 10);
  }
  try {
    await fetch(`${API}/opportunities/${id}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    await loadOpportunities();
    renderKanban();
    updateStatCards();
    renderDashboardCharts();
    updateTargetCard();
    if (currentSection === 'targets') renderOppTable();
    if (currentSection === 'prospects' || currentSection === 'contacts') loadContacts();
    closeOppEdit();
    showToast('商機已更新');
  } catch { showToast('儲存失敗，請重試'); }
});

// ── 主管業績達成率總覽 ────────────────────────────────────
let managerAchYearBuilt = false;

async function loadManagerAchievement(year) {
  const section = $('managerAchSection');
  if (!section) return;
  try {
    const res = await fetch(`/api/manager/achievement?year=${year}`);
    if (!res.ok) return;
    const { rows } = await res.json();
    const container = $('managerAchTable');

    if (!rows.length) { container.innerHTML = '<div style="color:#aaa;text-align:center;padding:20px">暫無資料</div>'; return; }

    const fmt = n => (n || 0).toLocaleString();
    const rateColor = r => r === null ? '#bbb' : r >= 100 ? '#0a8a4a' : r >= 70 ? '#1a73e8' : r >= 40 ? '#f59e0b' : '#e53e3e';
    const rateText  = r => r === null ? '–' : r + '%';
    const ROLE_BADGE = { manager1:'一級主管', manager2:'二級主管' };

    // ── SVG 半圓油表產生器 ──────────────────────────────────
    // center(100,108) radius=80, 弧長=π*80≈251.3
    function makeGaugeSvg(rate, color) {
      const R = 80, cx = 100, cy = 108;
      const arcLen = Math.PI * R;                      // ≈251.33
      const progress = rate === null ? 0 : Math.min(rate, 100) / 100 * arcLen;
      const lx = cx - R, rx = cx + R;                  // (20,108),(180,108)
      // M lx cy  A R R 0 0 0 rx cy  → 從左經頂部到右（sweep=0=逆時針在screen=往上）
      const d = `M ${lx} ${cy} A ${R} ${R} 0 0 0 ${rx} ${cy}`;

      // 顏色分段刻度（背景分5格）
      const ticks = [0, 0.25, 0.5, 0.75, 1.0].map(p => {
        const a = Math.PI * (1 - p);
        return { x: cx + R * Math.cos(a), y: cy - R * Math.sin(a) };
      });
      const tickMarks = ticks.map(t =>
        `<circle cx="${t.x.toFixed(1)}" cy="${t.y.toFixed(1)}" r="3" fill="#e0e7ef"/>`
      ).join('');

      // 中心文字
      const pctTxt = rate === null ? '—' : (rate > 999 ? '>999' : rate) + '%';
      const txtColor = rate === null ? '#bbb' : color;

      return `<svg viewBox="0 0 200 120" width="170" height="102" style="display:block;margin:0 auto">
        <!-- 背景弧 -->
        <path d="${d}" stroke="#e8ecf4" stroke-width="18" fill="none" stroke-linecap="round"/>
        <!-- 進度弧 -->
        <path d="${d}" stroke="${color}" stroke-width="18" fill="none" stroke-linecap="round"
          stroke-dasharray="${progress.toFixed(1)} ${arcLen.toFixed(1)}"/>
        ${tickMarks}
        <!-- 百分比文字 -->
        <text x="100" y="94" text-anchor="middle" font-size="26" font-weight="800"
          font-family="-apple-system,sans-serif" fill="${txtColor}">${pctTxt}</text>
        <!-- 0% / 100% 標籤 -->
        <text x="${lx - 4}" y="${cy + 18}" text-anchor="middle" font-size="9" fill="#bbb">0%</text>
        <text x="${rx + 4}" y="${cy + 18}" text-anchor="middle" font-size="9" fill="#bbb">100%</text>
      </svg>`;
    }

    // ── 卡片格 ──────────────────────────────────────────────
    container.innerHTML = `
      <div style="display:flex;flex-wrap:wrap;gap:16px;padding:4px 0">
        ${rows.map(r => {
          const color = rateColor(r.rate);
          const badge = ROLE_BADGE[r.role] ? `<span style="font-size:10px;background:#e8f0fe;color:#1a73e8;border-radius:10px;padding:1px 7px;margin-left:6px;font-weight:600">${ROLE_BADGE[r.role]}</span>` : '';
          return `
          <div class="ach-gauge-card" style="background:#fff;border:1.5px solid #e8ecf4;border-radius:14px;padding:16px 18px;min-width:210px;flex:1 1 210px;max-width:280px;box-shadow:0 2px 8px rgba(0,0,0,0.06)">
            <!-- 姓名 -->
            <div style="font-size:15px;font-weight:700;color:#1a2d52;margin-bottom:6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">
              ${r.displayName}${badge}
            </div>
            <!-- 油表 -->
            ${makeGaugeSvg(r.rate, color)}
            <!-- 目標（可編輯）-->
            <div class="mgr-target-cell" data-user="${r.username}" data-year="${year}" data-amount="${r.target || ''}"
                style="display:flex;align-items:center;justify-content:center;gap:6px;margin:8px 0 4px">
              <span style="font-size:11px;color:#888">目標</span>
              <span class="mgr-target-display" title="點擊編輯目標"
                style="font-size:14px;font-weight:700;color:${r.target ? '#1a2d52' : '#ccc'};border-bottom:1px dashed #bbb;cursor:pointer;padding-bottom:1px">
                ${r.target ? r.target.toLocaleString() + ' 萬' : '未設定'}
              </span>
              <span style="font-size:11px;color:#aaa">✎</span>
            </div>
            <!-- 數字列 -->
            <div style="display:flex;justify-content:space-around;margin-top:10px;padding-top:10px;border-top:1px solid #f0f0f0;font-size:12px;color:#555;text-align:center">
              <div>
                <div style="font-weight:700;font-size:14px;color:#0a8a4a">${fmt(r.achieved)}</div>
                <div style="color:#aaa;margin-top:2px">已成交（萬）</div>
              </div>
              <div style="width:1px;background:#f0f0f0"></div>
              <div>
                <div style="font-weight:700;font-size:14px;color:#1a73e8">${fmt(r.pipeline)}</div>
                <div style="color:#aaa;margin-top:2px">在手商機（萬）</div>
              </div>
              <div style="width:1px;background:#f0f0f0"></div>
              <div>
                <div style="font-weight:700;font-size:14px;color:#555">${r.wonCount}</div>
                <div style="color:#aaa;margin-top:2px">成交件數</div>
              </div>
            </div>
          </div>`;
        }).join('')}
      </div>`;

    // ── Inline 編輯：點擊目標欄位 ──
    container.querySelectorAll('.mgr-target-cell').forEach(cell => {
      const display = cell.querySelector('.mgr-target-display');
      display.addEventListener('click', () => {
        if (cell.querySelector('input')) return; // 已在編輯中
        const username = cell.dataset.user;
        const yr       = parseInt(cell.dataset.year);
        const curAmt   = cell.dataset.amount || '';

        // 換成 input + 確認/取消
        cell.innerHTML = `
          <div style="display:flex;align-items:center;gap:4px;justify-content:flex-end">
            <input type="number" value="${curAmt}" min="0" placeholder="金額"
              style="width:90px;font-size:13px;padding:3px 6px;border:1.5px solid #1a73e8;border-radius:5px;text-align:right">
            <button class="mgr-save-btn" style="padding:3px 8px;font-size:12px;background:#1a73e8;color:#fff;border:none;border-radius:4px;cursor:pointer">✓</button>
            <button class="mgr-cancel-btn" style="padding:3px 8px;font-size:12px;background:#eee;color:#555;border:none;border-radius:4px;cursor:pointer">✕</button>
          </div>`;

        const input  = cell.querySelector('input');
        const saveBtn   = cell.querySelector('.mgr-save-btn');
        const cancelBtn = cell.querySelector('.mgr-cancel-btn');
        input.focus(); input.select();

        const cancel = () => loadManagerAchievement(yr); // 取消重新渲染

        const save = async () => {
          const newAmt = parseFloat(input.value);
          if (isNaN(newAmt) || newAmt < 0) { showToast('請輸入有效金額'); return; }
          saveBtn.disabled = true; saveBtn.textContent = '…';
          try {
            const r = await fetch(`/api/manager/target/${encodeURIComponent(username)}`, {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ year: yr, amount: newAmt }),
            });
            if (!r.ok) { const e = await r.json(); showToast(e.error || '儲存失敗'); return; }
            showToast('✅ 目標已更新');
            loadManagerAchievement(yr); // 重新載入整個表格
          } catch { showToast('儲存失敗，請重試'); }
        };

        saveBtn.addEventListener('click', save);
        cancelBtn.addEventListener('click', cancel);
        input.addEventListener('keydown', e => {
          if (e.key === 'Enter') save();
          if (e.key === 'Escape') cancel();
        });
      });
    });

  } catch (e) {
    console.error('loadManagerAchievement error', e);
  }
}

// ── 年度目標管理 ─────────────────────────────────────────
let allTargets = [];
let allLostOppsForTargets = [];
let allZombieOpps = [];

async function loadTargets() {
  try {
    const res = await fetch(`${API}/targets`);
    allTargets = await res.json();
  } catch { /* ignore */ }
}

function getAchievedAmount(year) {
  return allOpportunities
    .filter(o => o.stage === 'Won')
    .filter(o => {
      const y = o.achievedDate ? new Date(o.achievedDate).getFullYear()
                               : new Date(o.createdAt).getFullYear();
      return y === year;
    })
    .reduce((sum, o) => sum + (parseFloat(o.amount) || 0), 0);
}

function updateTargetCard() {
  const year = new Date().getFullYear();
  const target = allTargets.find(t => t.year === year);
  $('targetAchYear').textContent = `${year} 年度業績目標`;
  const achieved = getAchievedAmount(year);
  $('targetAchievedDisplay').textContent = achieved.toLocaleString() + ' 萬';
  if (!target || !target.amount) {
    $('targetAmountDisplay').textContent = '尚未設定';
    $('targetRateDisplay').textContent = '--';
    $('targetProgressFill').style.width = '0%';
    return;
  }
  const rate = Math.min(100, Math.round(achieved / target.amount * 100));
  $('targetAmountDisplay').textContent = target.amount.toLocaleString() + ' 萬';
  $('targetRateDisplay').textContent = rate + '%';
  $('targetProgressFill').style.width = rate + '%';
}

async function loadTargetsView() {
  await loadTargets();
  await loadOpportunities();
  try {
    const res = await fetch(`${API}/lost-opportunities`);
    allLostOppsForTargets = res.ok ? await res.json() : [];
  } catch { allLostOppsForTargets = []; }
  try {
    const res = await fetch(`${API}/zombie-opportunities`);
    allZombieOpps = res.ok ? await res.json() : [];
  } catch { allZombieOpps = []; }
  renderZombieSection();

  // 年度選單
  const yearSel = $('targetYearSel');
  const curYear = new Date().getFullYear();
  yearSel.innerHTML = '';
  for (let y = curYear + 1; y >= curYear - 3; y--) {
    const o = document.createElement('option');
    o.value = y; o.textContent = y + ' 年';
    if (y === curYear) o.selected = true;
    yearSel.appendChild(o);
  }
  // 帶入目前年度目標
  const cur = allTargets.find(t => t.year === curYear);
  if (cur) $('targetAmountInput').value = cur.amount;

  yearSel.addEventListener('change', function () {
    const t = allTargets.find(x => x.year === parseInt(this.value));
    $('targetAmountInput').value = t ? t.amount : '';
  });

  // 歷史目標
  renderTargetHistory();

  // ── 主管業績達成率總覽（manager1 / manager2 / admin 才顯示）──
  if (['manager1', 'manager2', 'admin'].includes(window._myRole)) {
    const section = $('managerAchSection');
    if (section) {
      section.style.display = 'block';
      // 建立年度選單（只建一次）
      if (!managerAchYearBuilt) {
        managerAchYearBuilt = true;
        const achYearSel = $('managerAchYear');
        achYearSel.innerHTML = '';
        for (let y = curYear + 1; y >= curYear - 3; y--) {
          const opt = document.createElement('option');
          opt.value = y; opt.textContent = y + ' 年';
          if (y === curYear) opt.selected = true;
          achYearSel.appendChild(opt);
        }
        achYearSel.addEventListener('change', () => loadManagerAchievement(parseInt(achYearSel.value)));
      }
      loadManagerAchievement(curYear);
    }
  }

  // 商機列表篩選年度選單
  const oppYearSel = $('oppFilterYear');
  const years = [...new Set(allOpportunities.map(o =>
    new Date(o.createdAt).getFullYear()
  ))].sort((a,b) => b - a);
  oppYearSel.innerHTML = '<option value="">全部年度</option>';
  years.forEach(y => {
    const o = document.createElement('option');
    o.value = y; o.textContent = y + ' 年';
    if (y === curYear) o.selected = true;
    oppYearSel.appendChild(o);
  });

  renderOppTable();
  updateTargetCard();
}

function renderTargetHistory() {
  const hist = $('targetHistory');
  if (allTargets.length === 0) { hist.innerHTML = ''; return; }
  hist.innerHTML = allTargets
    .sort((a,b) => b.year - a.year)
    .map(t => {
      const ach = getAchievedAmount(t.year);
      const rate = t.amount ? Math.round(ach / t.amount * 100) : 0;
      return `<div class="target-history-chip">
        <span class="target-history-year">${t.year} 年</span>
        <span class="target-history-amt">目標 ${t.amount.toLocaleString()} 萬</span>
        <span class="target-history-ach">已達成 ${ach.toLocaleString()} 萬（${rate}%）</span>
      </div>`;
    }).join('');
}

$('saveTargetBtn').addEventListener('click', async () => {
  const year   = parseInt($('targetYearSel').value);
  const amount = parseFloat($('targetAmountInput').value);
  if (!year || isNaN(amount) || amount < 0) { showToast('請填入正確的目標金額'); return; }
  try {
    await fetch(`${API}/targets`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ year, amount })
    });
    await loadTargets();
    renderTargetHistory();
    updateTargetCard();
    showToast(`${year} 年度目標已儲存：${amount.toLocaleString()} 萬元`);
  } catch { showToast('儲存失敗，請重試'); }
});

// ── 商機列表渲染 ─────────────────────────────────────────
const OPP_STAGE_LABELS = {
  A: 'A｜Commit', B: 'B｜Upside', C: 'C｜Pipeline',
  D: 'D｜靜止中', Won: '🏆 Won'
};
const OPP_STAGE_COLORS = {
  A: '#c5221f', B: '#e37400', C: '#1a73e8', D: '#888', Won: '#1e8e3e'
};

function renderOppTable() {
  const filterYear  = $('oppFilterYear').value;
  const filterStage = $('oppFilterStage').value;
  const tbody = $('oppTableBody');
  const empty = $('oppEmpty');

  const showingDeleted = filterStage === '已刪除';

  // ── 正常商機 ──
  let list = [];
  if (!showingDeleted) {
    list = [...allOpportunities];
    if (filterYear)  list = list.filter(o => new Date(o.createdAt).getFullYear() === parseInt(filterYear));
    if (filterStage) list = list.filter(o => o.stage === filterStage);
    list.sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt));
  }

  // ── 已刪除商機（底部附加 or 獨立顯示）──
  let lostList = [];
  if (showingDeleted || !filterStage) {
    lostList = [...allLostOppsForTargets];
    if (filterYear) lostList = lostList.filter(o => new Date(o.createdAt||o.deletedAt).getFullYear() === parseInt(filterYear));
    lostList.sort((a,b) => new Date(b.deletedAt||0) - new Date(a.deletedAt||0));
  }

  const totalRows = list.length + lostList.length;
  if (totalRows === 0) {
    tbody.innerHTML = '';
    empty.style.display = 'block';
    return;
  }
  empty.style.display = 'none';

  // 殭屍商機 ID Set（供列標記用）
  const zombieMap = new Map(allZombieOpps.map(z => [z.id, z]));

  // ── 正常列渲染 ──
  const activeRows = list.map(o => {
    const won     = o.stage === 'Won';
    const zombie  = zombieMap.get(o.id);
    const zombieCls = zombie ? (zombie.severity === 'danger' ? 'opp-zombie-danger' : 'opp-zombie-warn') : '';
    const zombieBadge = zombie
      ? `<span class="opp-zombie-badge ${zombie.severity === 'danger' ? 'z-badge-danger' : 'z-badge-warn'}"
             title="${escapeHtml(zombie.reasons.join(' | '))}">🧟</span> `
      : '';
    const stageOpts = ['A','B','C','D','Won'].map(s =>
      `<option value="${s}" ${o.stage===s?'selected':''}>${OPP_STAGE_LABELS[s]||s}</option>`
    ).join('');
    return `<tr class="${won ? 'opp-won-row' : ''} ${zombieCls}" data-id="${o.id}">
      <td data-label="建立日期">${o.createdAt ? o.createdAt.slice(0,10) : ''}</td>
      <td data-label="客戶公司">${zombieBadge}${escapeHtml(o.company||'-')}</td>
      <td data-label="聯絡人">${escapeHtml(o.contactName||'-')}</td>
      <td data-label="商機類別">${escapeHtml(o.category||'-')}</td>
      <td data-label="商品項目" style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${o.product||''}">${escapeHtml(o.product||'-')}</td>
      <td data-label="預估金額（萬）" style="text-align:right;font-weight:700">${o.amount ? Number(o.amount).toLocaleString() : '-'}</td>
      <td data-label="預計成交日">${o.expectedDate || '-'}</td>
      <td data-label="狀態">
        <select class="opp-stage-sel" data-id="${o.id}" style="color:${OPP_STAGE_COLORS[o.stage]||'#333'}">${stageOpts}</select>
      </td>
      <td data-label="操作">
        ${won
          ? `<span style="font-size:12px;color:#34a853">已計入業績</span>`
          : `<button class="btn btn-sm btn-primary opp-won-btn" data-id="${o.id}">標記成交</button>`
        }
      </td>
    </tr>`;
  }).join('');

  // ── 已刪除列渲染 ──
  // 分隔列：只要有 lost 資料就顯示（不管是否混合顯示）
  const deletedSeparator = lostList.length > 0
    ? `<tr class="opp-lost-separator"><td colspan="9">💔 已刪除商機（可還原）</td></tr>`
    : '';

  const lostRows = lostList.map(o => {
    const deletedDate = o.deletedAt ? o.deletedAt.slice(0,10) : '';
    const reason = o.deleteReason ? `<span class="opp-lost-reason" title="${o.deleteReason}">原因：${o.deleteReason}</span>` : '';
    return `<tr class="opp-lost-row" data-id="${o.id}">
      <td data-label="建立日期"><span class="opp-lost-date">${o.createdAt ? o.createdAt.slice(0,10) : '-'}</span></td>
      <td data-label="客戶公司">${o.company || '-'}</td>
      <td data-label="聯絡人">${o.contactName || '-'}</td>
      <td data-label="商機類別">${o.category || '-'}</td>
      <td data-label="商品項目" style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${o.product||''}">${o.product || '-'}</td>
      <td data-label="預估金額（萬）" style="text-align:right">${o.amount ? Number(o.amount).toLocaleString() : '-'}</td>
      <td data-label="預計成交日">${o.expectedDate || '-'}</td>
      <td data-label="狀態">
        <span class="opp-lost-badge">💔 流失</span>
        ${reason}
        ${deletedDate ? `<div style="font-size:11px;color:#aaa;margin-top:2px">刪除於 ${deletedDate}</div>` : ''}
      </td>
      <td data-label="操作">
        <button class="btn btn-sm opp-restore-btn" data-id="${o.id}" title="還原此商機至 Pipeline 階段">↩️ 還原</button>
      </td>
    </tr>`;
  }).join('');

  tbody.innerHTML = activeRows
    + (lostList.length > 0 ? deletedSeparator + lostRows : '');

  // 狀態下拉更新
  tbody.querySelectorAll('.opp-stage-sel').forEach(sel => {
    sel.addEventListener('change', async function () {
      await updateOppStage(this.dataset.id, this.value);
    });
  });

  // 標記成交按鈕
  tbody.querySelectorAll('.opp-won-btn').forEach(btn => {
    btn.addEventListener('click', async function () {
      await updateOppStage(this.dataset.id, 'Won', true);
    });
  });

  // 還原按鈕
  tbody.querySelectorAll('.opp-restore-btn').forEach(btn => {
    btn.addEventListener('click', async function () {
      const id = this.dataset.id;
      const opp = allLostOppsForTargets.find(o => o.id === id);
      const name = opp ? (opp.company || opp.product || id) : id;
      if (!confirm(`確定要還原商機「${name}」嗎？\n還原後將回到 C｜Pipeline 階段。`)) return;
      try {
        const r = await fetch(`${API}/opportunities/restore/${id}`, { method: 'POST' });
        if (!r.ok) { const e = await r.json(); showToast('還原失敗：' + (e.error||'未知錯誤')); return; }
        showToast('✅ 商機已還原至 C｜Pipeline 階段');
        // 重新載入
        await loadOpportunities();
        const res2 = await fetch(`${API}/lost-opportunities`);
        allLostOppsForTargets = res2.ok ? await res2.json() : [];
        renderOppTable();
        updateTargetCard();
        updateStatCards();
        renderDashboardCharts();
      } catch { showToast('還原失敗，請重試'); }
    });
  });
}

async function updateOppStage(id, stage, confirmWon = false) {
  const isWon = stage === 'Won';
  if (confirmWon && !confirm(`確認將此商機標記為【${stage}】並計入年度業績？`)) return;
  try {
    const body = { stage };
    if (isWon) body.achievedDate = new Date().toISOString().slice(0,10);
    await fetch(`${API}/opportunities/${id}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    await loadOpportunities();
    renderOppTable();
    renderTargetHistory();
    updateTargetCard();
    updateStatCards();
    renderDashboardCharts();
    if (currentSection === 'prospects' || currentSection === 'contacts') loadContacts();
    if (stage === 'Won') {
      const opp = allOpportunities.find(x => x.id === id);
      if (opp) celebrateWon(opp);
    } else {
      showToast(isWon ? `✅ 已標記成交，計入年度業績` : '商機狀態已更新');
    }
  } catch { showToast('更新失敗，請重試'); }
}

$('oppFilterYear').addEventListener('change', renderOppTable);
$('oppFilterStage').addEventListener('change', renderOppTable);

// ── 銷售預測報表 ─────────────────────────────────────────
let forecastYear = new Date().getFullYear();
let forecastSelectedStages = [];   // 空陣列 = 全部階段

// 階段 → 把握度%
const STAGE_CONFIDENCE = { D: 10, C: 25, B: 50, A: 90, Won: 100 };
// 階段 → 把握度顯示標籤
const STAGE_CONF_LABEL = { A: 'Commit', B: 'Upside', C: 'Pipeline', Won: 'Won' };

// 業務人員（從登入資訊取得）
let forecastSalesPerson = '';
// 用戶名稱對應表（username → displayName），供多用戶報表顯示
let forecastUserMap = {};

async function loadForecastView() {
  await Promise.all([
    loadOpportunities(),
    allVisits.length === 0 ? fetch(`${API}/visits`).then(r=>r.json()).then(d=>{ allVisits=d; }) : Promise.resolve()
  ]);
  if (!forecastSalesPerson) {
    try {
      const r = await fetch(`${API}/me`);
      const u = await r.json();
      forecastSalesPerson = u.displayName || u.username || '';
    } catch { /* ignore */ }
  }
  // 取得可見用戶名稱對應表（秘書/主管才有多筆）
  try {
    const r = await fetch(`${API}/usermap`);
    if (r.ok) forecastUserMap = await r.json();
  } catch { /* ignore */ }
  initForecastYearSel();
  initForecastSalesFilter();   // 載入 userMap 後再填入業務選單
  renderForecastTable();
}

function initForecastYearSel() {
  const sel = $('forecastYear');
  if (sel.options.length) return;
  const cur = new Date().getFullYear();
  for (let y = cur - 1; y <= cur + 2; y++) {
    const o = document.createElement('option');
    o.value = y; o.textContent = `${y} 年`;
    if (y === forecastYear) o.selected = true;
    sel.appendChild(o);
  }
  sel.addEventListener('change', () => { forecastYear = parseInt(sel.value); renderForecastTable(); });

}

function initForecastSalesFilter() {
  const canFilter = ['admin','manager1','manager2','secretary'].includes(userPermissions.role);
  const salesFilter = $('forecastSalesFilter');

  // 業務人員篩選：只限主管/秘書
  if (canFilter && salesFilter) {
    salesFilter.style.display = '';
    const existingVals = new Set(Array.from(salesFilter.options).map(o => o.value));
    Object.entries(forecastUserMap).forEach(([username, displayName]) => {
      if (!existingVals.has(username)) {
        const o = document.createElement('option');
        o.value = username;
        o.textContent = displayName || username;
        salesFilter.appendChild(o);
      }
    });
    if (!salesFilter.dataset.bound) {
      salesFilter.addEventListener('change', renderForecastTable);
      salesFilter.dataset.bound = '1';
    }
  }

  // 階段多選：所有角色都顯示
  initForecastStageMultiSel();
}

function initForecastStageMultiSel() {
  const wrap  = $('forecastStageWrap');
  const btn   = $('forecastStageBtn');
  const panel = $('forecastStagePanel');
  if (!wrap || !btn || !panel) return;
  if (wrap.dataset.bound) return;
  wrap.dataset.bound = '1';

  wrap.style.display = '';

  // 按鈕開/關 panel
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    const open = panel.style.display !== 'none';
    panel.style.display = open ? 'none' : '';
  });

  // 點 checkbox → 更新 state + 重算
  panel.addEventListener('change', (e) => {
    if (!e.target.classList.contains('fc-stage-chk')) return;
    forecastSelectedStages = Array.from(
      panel.querySelectorAll('.fc-stage-chk:checked')
    ).map(c => c.value);
    _updateForecastStageBtn();
    renderForecastTable();
  });

  // 點外部關閉
  document.addEventListener('click', (e) => {
    if (!wrap.contains(e.target)) panel.style.display = 'none';
  }, true);

  function _updateForecastStageBtn() {
    const n = forecastSelectedStages.length;
    const STAGE_NAMES = { A:'Commit', B:'Upside', C:'Pipeline', Won:'Won' };
    if (n === 0) {
      btn.textContent = '全部階段 ▾';
      btn.classList.remove('active');
    } else if (n === 1) {
      btn.textContent = (STAGE_NAMES[forecastSelectedStages[0]] || forecastSelectedStages[0]) + ' ▾';
      btn.classList.add('active');
    } else {
      btn.textContent = `已選 ${n} 個階段 ▾`;
      btn.classList.add('active');
    }
  }
}

function getForecastOpps(year) {
  const salesVal = $('forecastSalesFilter') ? $('forecastSalesFilter').value : '';

  return allOpportunities.filter(o => {
    if (!o.expectedDate) return false;
    if (o.stage === 'D') return false;          // D 階段不納入銷售預測
    if (new Date(o.expectedDate).getFullYear() !== year) return false;
    if (salesVal && o.owner !== salesVal) return false;   // 業務人員篩選
    if (forecastSelectedStages.length && !forecastSelectedStages.includes(o.stage)) return false;  // 多選階段篩選
    return true;
  });
}

function renderForecastTable() {
  const year = forecastYear;
  const opps = getForecastOpps(year);

  // ── 摘要卡 ──
  const totContract = opps.reduce((s, o) => s + (parseFloat(o.amount) || 0) * 10, 0); // 萬→NT$K
  const totGrossProfit = opps.reduce((s, o) => {
    const amt = (parseFloat(o.amount) || 0) * 10;
    const gm  = parseFloat(o.grossMarginRate) || 0;
    return s + amt * gm / 100;
  }, 0);
  const wonOpps = opps.filter(o => o.stage === 'Won');
  const totWon = wonOpps.reduce((s, o) => s + (parseFloat(o.amount) || 0) * 10, 0);

  // ── 篩選標籤（顯示目前套用的條件）──
  const salesVal = $('forecastSalesFilter') ? $('forecastSalesFilter').value : '';
  const salesLabel = salesVal ? (forecastUserMap[salesVal] || salesVal) : '';
  const STAGE_TAG_NAMES = { A:'A｜Commit', B:'B｜Upside', C:'C｜Pipeline', Won:'Won' };
  const stageLabels = forecastSelectedStages.map(s => STAGE_TAG_NAMES[s] || s);
  const filterTags = [salesLabel, ...stageLabels].filter(Boolean);
  const filterHint = filterTags.length
    ? `<div style="font-size:12px;color:#1a73e8;margin-bottom:6px">
        篩選中：${filterTags.map(t=>`<span style="background:#e8f0fe;border-radius:4px;padding:2px 8px;margin-right:4px">${t}</span>`).join('')}
       </div>`
    : '';
  const hintEl = $('forecastFilterHint');
  if (hintEl) hintEl.innerHTML = filterHint;

  $('forecastSummaryCards').innerHTML = `
    <div class="forecast-summary-card fsc-blue">
      <div class="fsc-label">商機筆數</div>
      <div class="fsc-value">${opps.length} 筆</div>
    </div>
    <div class="forecast-summary-card fsc-orange">
      <div class="fsc-label">合約金額合計（NT$K）</div>
      <div class="fsc-value">${totContract.toLocaleString()}</div>
    </div>
    <div class="forecast-summary-card fsc-green">
      <div class="fsc-label">預估毛利合計（NT$K）</div>
      <div class="fsc-value">${Math.round(totGrossProfit).toLocaleString()}</div>
    </div>
    <div class="forecast-summary-card fsc-red">
      <div class="fsc-label">已成交金額（NT$K）</div>
      <div class="fsc-value">${totWon.toLocaleString()}</div>
    </div>`;

  // ── 表頭 ──
  $('forecastThead').innerHTML = `
    <tr>
      <th class="ft-company-col">客戶名稱</th>
      <th class="ft-case-col">銷售案名</th>
      <th class="ft-bu-col">BU</th>
      <th>預定簽約日</th>
      <th>業務人員</th>
      <th class="ft-head-pink">把握度%</th>
      <th class="ft-head-pink">預估<br>毛利率</th>
      <th class="ft-head-yellow">合約金額<br>（NT$K）</th>
      <th class="ft-head-yellow">毛利金額<br>（NT$K）</th>
    </tr>`;

  // ── 資料列 ──
  const tbody = $('forecastTbody');
  tbody.innerHTML = '';

  if (opps.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center;padding:40px;color:#bbb">
      ${year} 年尚無商機資料，請在商機編輯中填入「預定簽約日期」</td></tr>`;
    return;
  }

  // 依預定簽約日排序
  const sorted = [...opps].sort((a, b) => (a.expectedDate || '').localeCompare(b.expectedDate || ''));

  sorted.forEach(o => {
    const amt    = (parseFloat(o.amount) || 0) * 10;           // 萬 → NT$K
    const gm     = parseFloat(o.grossMarginRate) || 0;
    const profit = Math.round(amt * gm / 100);
    const confLabel   = STAGE_CONF_LABEL[o.stage] || null;
    const confDisplay = confLabel ? `<span class="ft-conf-label">${confLabel}</span>` : '—';
    const salesName = o.product || o.description || '—';
    const stageClass = o.stage === 'Won' ? 'ft-won' : (o.stage === 'A' ? 'ft-stage-a' : '');
    // 業務人員：優先用 owner 查對應表，找不到才用登入者名稱
    const salesPerson = (o.owner && forecastUserMap[o.owner])
      ? forecastUserMap[o.owner]
      : forecastSalesPerson || '—';

    const tr = document.createElement('tr');
    tr.className = stageClass;
    tr.style.cursor = 'pointer';
    tr.innerHTML = `
      <td class="ft-company-col">${escapeHtml(o.company) || '—'}</td>
      <td class="ft-case-col" title="${escapeHtml(salesName)}">${escapeHtml(salesName)}</td>
      <td class="ft-bu-col">${escapeHtml(o.category) || '—'}</td>
      <td>${escapeHtml(o.expectedDate) || '—'}</td>
      <td>${escapeHtml(salesPerson)}</td>
      <td class="ft-cell-pink ft-center">${confDisplay}</td>
      <td class="ft-cell-pink ft-center">${gm ? escapeHtml(String(gm)) + '%' : '—'}</td>
      <td class="ft-cell-yellow ft-right">${amt ? amt.toLocaleString() : '—'}</td>
      <td class="ft-cell-yellow ft-right">${profit ? profit.toLocaleString() : '—'}</td>`;
    tr.addEventListener('click', () => openOppEdit(o.id));

    // ── 浮動拜訪記錄 tooltip ──
    tr.addEventListener('mouseenter', (e) => showOppVisitTooltip(e, o));
    tr.addEventListener('mousemove',  (e) => moveOppVisitTooltip(e));
    tr.addEventListener('mouseleave', ()  => hideOppVisitTooltip());

    tbody.appendChild(tr);
  });

  // 合計列
  const tr = document.createElement('tr');
  tr.className = 'ft-total-row';
  const avgConf = opps.length
    ? Math.round(opps.reduce((s, o) => s + (STAGE_CONFIDENCE[o.stage] ?? 0), 0) / opps.length)
    : 0;
  tr.innerHTML = `
    <td class="ft-company-col" colspan="5">合　計（${opps.length} 筆）</td>
    <td class="ft-cell-pink ft-center">—</td>
    <td class="ft-cell-pink ft-center">—</td>
    <td class="ft-cell-yellow ft-right">${Math.round(totContract).toLocaleString()}</td>
    <td class="ft-cell-yellow ft-right">${Math.round(totGrossProfit).toLocaleString()}</td>`;
  tbody.appendChild(tr);
}

// ── 商機拜訪記錄浮動 Tooltip ─────────────────────────────
const _ovTip = () => $('oppVisitTooltip');

function showOppVisitTooltip(e, opp) {
  // 找最新一筆：先比對 contactId，沒有再比對公司名稱
  const latest = [...allVisits]
    .filter(v => (opp.contactId && v.contactId === opp.contactId) ||
                 (opp.company   && (allContacts.find(c => c.id === v.contactId)?.company === opp.company)))
    .sort((a, b) => (b.visitDate || '').localeCompare(a.visitDate || ''))[0];

  const tip = _ovTip();
  if (!latest) {
    tip.innerHTML = `
      <div class="ovt-header">📋 最新業務日報</div>
      <div class="ovt-empty">尚無拜訪記錄</div>`;
  } else {
    const typeIcon = { '親訪':'🤝','電話':'📞','視訊':'💻','Email':'📧','展覽':'🎪' }[latest.visitType] || '📋';
    const contentSnip = latest.content
      ? latest.content.length > 80 ? latest.content.slice(0, 80) + '…' : latest.content
      : '—';
    const nextAct = latest.nextAction || '—';
    tip.innerHTML = `
      <div class="ovt-header">📋 最新業務日報</div>
      <div class="ovt-meta">
        <span class="ovt-date">${latest.visitDate || ''}</span>
        <span class="ovt-type">${typeIcon} ${latest.visitType || '親訪'}</span>
      </div>
      <div class="ovt-topic">${escapeHtml(latest.topic || '—')}</div>
      ${contentSnip !== '—' ? `<div class="ovt-content">${escapeHtml(contentSnip)}</div>` : ''}
      <div class="ovt-next-label">⚡ 下一步行動</div>
      <div class="ovt-next">${escapeHtml(nextAct)}</div>`;
  }
  tip.style.display = 'block';
  moveOppVisitTooltip(e);
}

function moveOppVisitTooltip(e) {
  const tip = _ovTip();
  const vw = window.innerWidth, vh = window.innerHeight;
  const tw = tip.offsetWidth || 280, th = tip.offsetHeight || 160;
  let x = e.clientX + 18, y = e.clientY + 12;
  if (x + tw > vw - 12) x = e.clientX - tw - 12;
  if (y + th > vh - 12) y = e.clientY - th - 12;
  tip.style.left = x + 'px';
  tip.style.top  = y + 'px';
}

function hideOppVisitTooltip() {
  _ovTip().style.display = 'none';
}

// 匯出 Excel
$('forecastExportBtn').addEventListener('click', () => {
  window.location.href = `${API}/forecast/export?year=${forecastYear}`;
});

// ── ERP MA 合約管理 ───────────────────────────────────────
let allContracts = [];
let contractDeleteId = null;

function contractStatus(c) {
  const today = new Date(); today.setHours(0,0,0,0);
  const endDate   = (c && typeof c === 'object') ? c.endDate   : c;
  const renewDate = (c && typeof c === 'object') ? c.renewDate : null;
  const startDate = (c && typeof c === 'object') ? c.startDate : null;

  if (!endDate) return { key: 'active', label: '有效', days: null, pct: 100, color: 'green', isRenewed: false };

  const end     = new Date(endDate);
  const endDiff = Math.ceil((end - today) / 86400000);

  let effectiveEnd, effectiveStart, isRenewed = false;

  if (endDiff < 0 && renewDate) {
    // 已到期但有續約 → 改用續約日期
    effectiveEnd   = new Date(renewDate);
    const afterEnd = new Date(endDate);
    afterEnd.setDate(afterEnd.getDate() + 1);
    effectiveStart = afterEnd;
    isRenewed = true;
  } else if (endDiff < 0) {
    // 真正逾期，無續約
    return { key: 'expired', label: '逾期', days: Math.abs(endDiff), pct: 0, color: 'red', isRenewed: false };
  } else {
    effectiveEnd   = end;
    effectiveStart = startDate ? new Date(startDate) : null;
  }

  const diff = Math.ceil((effectiveEnd - today) / 86400000);

  // 進度條：剩餘時間百分比
  let pct = 100;
  if (effectiveStart) {
    const total   = Math.ceil((effectiveEnd - effectiveStart) / 86400000);
    const elapsed = Math.ceil((today - effectiveStart) / 86400000);
    pct = total > 0 ? Math.max(0, Math.min(100, Math.round(((total - elapsed) / total) * 100))) : 0;
  }

  const color = diff <= 25 ? 'red' : diff <= 90 ? 'yellow' : 'green';

  if (diff <= 25)  return { key: 'urgent',   label: isRenewed ? '續約即將到期' : '即將到期', days: diff, pct, color, isRenewed };
  if (diff <= 90)  return { key: 'expiring', label: isRenewed ? '續約有效'     : '即將到期', days: diff, pct, color, isRenewed };
  return                  { key: 'active',   label: isRenewed ? '續約有效'     : '有效',     days: diff, pct, color, isRenewed };
}

// ── 流失商機分析 ──────────────────────────────────────────
let allLostOpps = [];
let lostOppListenersSet = false;

async function loadLostOppView() {
  try {
    const res = await fetch('/api/lost-opportunities');
    allLostOpps = await res.json();
    renderLostOppTable();

    if (!lostOppListenersSet) {
      lostOppListenersSet = true;
      $('lostOppSearch').addEventListener('input', renderLostOppTable);
      $('lostOppReasonFilter').addEventListener('change', renderLostOppTable);
      $('lostOppMonth').addEventListener('change', renderLostOppTable);
      $('lostOppClear').addEventListener('click', () => {
        $('lostOppSearch').value = '';
        $('lostOppReasonFilter').value = '';
        $('lostOppMonth').value = '';
        renderLostOppTable();
      });
    }
  } catch (e) { console.error('loadLostOppView', e); }
}

function renderLostOppTable() {
  const keyword = ($('lostOppSearch').value || '').toLowerCase();
  const reasonF = $('lostOppReasonFilter').value;
  const monthF  = $('lostOppMonth').value;  // "YYYY-MM"

  let rows = allLostOpps.filter(o => {
    if (keyword && !((o.company||'').toLowerCase().includes(keyword) ||
                     (o.product||'').toLowerCase().includes(keyword))) return false;
    if (reasonF && !(o.deleteReason||'').startsWith(reasonF)) return false;
    if (monthF) {
      const d = (o.deletedAt||'').slice(0,7); // "YYYY-MM"
      if (d !== monthF) return false;
    }
    return true;
  });

  // 統計卡
  const totalAmt = rows.reduce((s,o) => s + (parseFloat(o.amount)||0), 0);
  const byReason = {};
  rows.forEach(o => {
    const r = (o.deleteReason||'未填').split('：')[0];
    byReason[r] = (byReason[r]||0) + 1;
  });
  const topReason = Object.entries(byReason).sort((a,b)=>b[1]-a[1])[0];

  $('lostOppStats').innerHTML = `
    <div style="background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.08);padding:14px 20px;min-width:140px">
      <div style="font-size:11px;color:#888;margin-bottom:4px">流失案件數</div>
      <div style="font-size:22px;font-weight:700;color:#ea4335">${rows.length}</div>
    </div>
    <div style="background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.08);padding:14px 20px;min-width:140px">
      <div style="font-size:11px;color:#888;margin-bottom:4px">流失金額(萬)</div>
      <div style="font-size:22px;font-weight:700;color:#f57c00">${totalAmt.toLocaleString()}</div>
    </div>
    ${topReason ? `<div style="background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.08);padding:14px 20px;min-width:140px">
      <div style="font-size:11px;color:#888;margin-bottom:4px">最常見原因</div>
      <div style="font-size:16px;font-weight:600;color:#333">${escapeHtml(topReason[0])}</div>
      <div style="font-size:12px;color:#888">${topReason[1]} 件</div>
    </div>` : ''}
  `;

  $('lostOppCount').textContent = `共 ${rows.length} 筆`;

  const STAGE_COLORS = {
    '潛在機會':'#607d8b','初步接觸':'#2196f3','需求確認':'#9c27b0',
    '方案展示':'#ff9800','提案報價':'#ff5722','議價中':'#f44336','結案失敗':'#795548'
  };

  const tbody = $('lostOppTbody');
  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;padding:40px;color:#aaa">沒有符合條件的流失商機</td></tr>`;
    return;
  }

  tbody.innerHTML = rows.map(o => {
    const date = o.deletedAt ? new Date(o.deletedAt).toLocaleDateString('zh-TW') : '';
    const stageColor = STAGE_COLORS[o.stage] || '#aaa';
    const amt = parseFloat(o.amount) || 0;
    return `<tr style="border-bottom:1px solid #f0f0f0">
      <td style="padding:10px 14px;white-space:nowrap;color:#888;font-size:12px">${date}</td>
      <td style="padding:10px 14px;font-weight:500">${escapeHtml(o.company||'')}</td>
      <td style="padding:10px 14px">${escapeHtml(o.product||'')}</td>
      <td style="padding:10px 14px">
        <span style="background:${stageColor}20;color:${stageColor};padding:2px 8px;border-radius:12px;font-size:12px;font-weight:500">${escapeHtml(o.stage||'')}</span>
      </td>
      <td style="padding:10px 14px;text-align:right;font-weight:500">${amt ? amt.toLocaleString() : '—'}</td>
      <td style="padding:10px 14px;color:#555">${escapeHtml(o.deleteReason||'—')}</td>
      <td style="padding:10px 14px;color:#888;font-size:12px">${escapeHtml(o.deletedByName||o.deletedBy||'')}</td>
    </tr>`;
  }).join('');
}

// ── 商機動態報表 ──────────────────────────────────────────
let prChart = null;          // Chart.js instance
let prCurrentPeriod = 'month';
let prCurrentOwner  = '';   // '' = 全部業務
let prListenersSet  = false;
let prOwnerBuilt    = false; // 業務下拉只建立一次

const STAGE_COLOR = {
  'A': '#d32f2f',   // Commit   紅
  'B': '#f57c00',   // Upside   橘
  'C': '#1a73e8',   // Pipeline 藍
  'D': '#9e9e9e',   // 靜止中   灰
  'Won':  '#1e8e3e',
  'Won': '#1e8e3e', // Won      深綠
};
const STAGE_LBL = {
  'A':'A｜Commit','B':'B｜Upside','C':'C｜Pipeline','D':'D｜靜止中','Won':'🏆 Won'
};

function getPrDateRange(period) {
  const now = new Date();
  let from, to;
  if (period === 'month') {
    from = new Date(now.getFullYear(), now.getMonth(), 1);
    to   = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
  } else if (period === 'lastmonth') {
    from = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    to   = new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59);
  } else if (period === 'quarter') {
    const q = Math.floor(now.getMonth() / 3);
    from = new Date(now.getFullYear(), q * 3, 1);
    to   = new Date(now.getFullYear(), q * 3 + 3, 0, 23, 59, 59);
  } else if (period === 'custom') {
    const f = $('prFromDate').value, t = $('prToDate').value;
    from = f ? new Date(f) : new Date(now.getFullYear(), now.getMonth(), 1);
    to   = t ? new Date(t + 'T23:59:59') : new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
  }
  return { from, to };
}

function fmtPeriodLabel(from, to) {
  const f = d => `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,'0')}/${String(d.getDate()).padStart(2,'0')}`;
  return `${f(from)} ～ ${f(to)}`;
}

async function loadPipelineReport() {
  const { from, to } = getPrDateRange(prCurrentPeriod);
  $('prPeriodLabel').textContent = fmtPeriodLabel(from, to);

  try {
    let qs = `from=${from.toISOString()}&to=${to.toISOString()}`;
    if (prCurrentOwner) qs += `&owner=${encodeURIComponent(prCurrentOwner)}`;
    const res = await fetch(`/api/pipeline-report?${qs}`);
    if (!res.ok) throw new Error('載入失敗');
    const data = await res.json();

    // 第一次載入時建立業務人員篩選下拉
    if (!prOwnerBuilt && data.ownerOptions && data.ownerOptions.length > 1) {
      prOwnerBuilt = true;
      const wrap   = $('prOwnerWrap');
      const select = $('prOwnerSelect');
      // 清除舊 options 再填入
      select.innerHTML = '<option value="">全部業務</option>';
      data.ownerOptions.forEach(o => {
        const opt = document.createElement('option');
        opt.value = o.username;
        opt.textContent = o.displayName;
        select.appendChild(opt);
      });
      select.value = prCurrentOwner;
      wrap.style.display = 'flex';
      select.addEventListener('change', () => {
        prCurrentOwner = select.value;
        loadPipelineReport();
      });
    }

    renderPrSummary(data.summary);
    renderPrFunnel(data.funnel);
    renderPrMoves(data);
  } catch (e) {
    showToast('無法載入動態報表：' + e.message);
  }

  if (!prListenersSet) {
    prListenersSet = true;

    // 期間按鈕切換
    document.querySelectorAll('.pr-period-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.pr-period-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        prCurrentPeriod = btn.dataset.period;
        $('prCustomRange').style.display = prCurrentPeriod === 'custom' ? 'flex' : 'none';
        if (prCurrentPeriod !== 'custom') loadPipelineReport();
      });
    });

    $('prCustomApply').addEventListener('click', () => {
      prCurrentPeriod = 'custom';
      loadPipelineReport();
    });
  }
}

function renderPrSummary(s) {
  const cards = [
    { icon:'💼', label:'在手商機總額', value:`${(s.totalPipeline||0).toLocaleString()} 萬`, sub:`共 ${s.totalCount} 件`, color:'#1a73e8', bg:'#e8f0fe' },
    { icon:'✨', label:'本期新增商機', value:`${(s.newAmount||0).toLocaleString()} 萬`, sub:`${s.newCount} 件`, color:'#34a853', bg:'#e6f4ea' },
    { icon:'⬆️', label:'階段晉升', value:`${(s.promotedAmount||0).toLocaleString()} 萬`, sub:`${s.promotedCount} 次`, color:'#1e8e3e', bg:'#e6f4ea' },
    { icon:'⬇️', label:'階段退後', value:`${(s.demotedAmount||0).toLocaleString()} 萬`, sub:`${s.demotedCount} 次`, color:'#f57c00', bg:'#fff3e0' },
    { icon:'💔', label:'本期流失', value:`${(s.lostAmount||0).toLocaleString()} 萬`, sub:`${s.lostCount} 件`, color:'#c62828', bg:'#fce8e6' },
  ];
  $('prSummaryCards').innerHTML = cards.map(c => `
    <div class="pr-summary-card" style="border-top:3px solid ${c.color}">
      <div class="pr-summary-icon" style="background:${c.bg};color:${c.color}">${c.icon}</div>
      <div class="pr-summary-body">
        <div class="pr-summary-label">${c.label}</div>
        <div class="pr-summary-value" style="color:${c.color}">${c.value}</div>
        <div class="pr-summary-sub">${c.sub}</div>
      </div>
    </div>`).join('');
}

function renderPrFunnel(funnel) {
  // 橫條漏斗
  const maxAmt = Math.max(...funnel.map(f => f.amount), 1);
  $('prFunnelBars').innerHTML = funnel.slice().reverse().map(f => {
    const pct = Math.max((f.amount / maxAmt) * 100, f.count ? 4 : 0);
    const color = STAGE_COLOR[f.stage] || '#90a4ae';
    return `
      <div class="pr-funnel-row">
        <div class="pr-funnel-lbl">${STAGE_LBL[f.stage]||f.stage}</div>
        <div class="pr-funnel-bar-wrap">
          <div class="pr-funnel-bar" style="width:${pct}%;background:${color}"></div>
        </div>
        <div class="pr-funnel-meta">
          <span class="pr-funnel-amt">${f.amount ? f.amount.toLocaleString()+'萬' : '—'}</span>
          <span class="pr-funnel-cnt">${f.count}件</span>
        </div>
      </div>`;
  }).join('');

  // Chart.js 圓形圖（只顯示有金額的）
  const active = funnel.filter(f => f.amount > 0);
  if (prChart) { prChart.destroy(); prChart = null; }
  const canvas = $('prFunnelChart');
  if (active.length === 0) {
    canvas.style.display = 'none';
    return;
  }
  canvas.style.display = '';
  prChart = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels: active.map(f => STAGE_LBL[f.stage]||f.stage),
      datasets: [{
        data: active.map(f => f.amount),
        backgroundColor: active.map(f => STAGE_COLOR[f.stage]||'#90a4ae'),
        borderWidth: 2,
        borderColor: '#fff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position:'bottom', labels:{ font:{size:10}, padding:8, boxWidth:10 } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.label}：${ctx.parsed.toLocaleString()} 萬`
          }
        }
      },
      cutout: '52%'
    }
  });
}

function renderPrMoves(data) {
  const fmt = o => {
    const stageTag = o.stage ? `<span class="pr-stage-tag" style="background:${STAGE_COLOR[o.stage]||'#90a4ae'}20;color:${STAGE_COLOR[o.stage]||'#90a4ae'};border:1px solid ${STAGE_COLOR[o.stage]||'#90a4ae'}40">${STAGE_LBL[o.stage]||o.stage}</span>` : '';
    const amt = o.amount ? `<span class="pr-deal-amt">${o.amount.toLocaleString()}萬</span>` : '';
    return `<div class="pr-deal-row">
      <div class="pr-deal-info">
        <div class="pr-deal-company">${escapeHtml(o.company||'')}</div>
        <div class="pr-deal-product">${escapeHtml(o.product||'')}</div>
      </div>
      <div class="pr-deal-right">${stageTag}${amt}</div>
    </div>`;
  };

  const fmtMove = o => {
    const fi = STAGE_COLOR[o.from]||'#90a4ae', ti = STAGE_COLOR[o.to]||'#90a4ae';
    return `<div class="pr-deal-row">
      <div class="pr-deal-info">
        <div class="pr-deal-company">${escapeHtml(o.company||'')}</div>
        <div class="pr-deal-product">${escapeHtml(o.product||'')}</div>
      </div>
      <div class="pr-deal-right">
        <span class="pr-stage-tag" style="background:${fi}20;color:${fi};border:1px solid ${fi}40">${STAGE_LBL[o.from]||o.from}</span>
        <span style="color:#888;font-size:11px">→</span>
        <span class="pr-stage-tag" style="background:${ti}20;color:${ti};border:1px solid ${ti}40">${STAGE_LBL[o.to]||o.to}</span>
        <span class="pr-deal-amt">${o.amount.toLocaleString()}萬</span>
      </div>
    </div>`;
  };

  // 晉升
  $('prPromotedCount').textContent = data.promoted.length;
  $('prPromotedList').innerHTML = data.promoted.length
    ? data.promoted.map(fmtMove).join('')
    : '<div class="pr-empty-row">本期無階段晉升記錄</div>';

  // 退後
  $('prDemotedCount').textContent = data.demoted.length;
  $('prDemotedList').innerHTML = data.demoted.length
    ? data.demoted.map(fmtMove).join('')
    : '<div class="pr-empty-row">本期無階段退後記錄</div>';

  // 新增
  $('prNewCount').textContent = data.newDeals.length;
  $('prNewList').innerHTML = data.newDeals.length
    ? data.newDeals.map(fmt).join('')
    : '<div class="pr-empty-row">本期無新增商機</div>';

  // 流失
  $('prLostCount').textContent = data.lostDeals.length;
  $('prLostList').innerHTML = data.lostDeals.length
    ? data.lostDeals.map(o => {
        const amt = o.amount ? `<span class="pr-deal-amt">${o.amount.toLocaleString()}萬</span>` : '';
        return `<div class="pr-deal-row">
          <div class="pr-deal-info">
            <div class="pr-deal-company">${escapeHtml(o.company||'')}</div>
            <div class="pr-deal-product">${escapeHtml(o.product||'')} ${o.deleteReason ? `<span style="color:#ea4335;font-size:11px">（${escapeHtml(o.deleteReason)}）</span>` : ''}</div>
          </div>
          <div class="pr-deal-right">${amt}</div>
        </div>`;
      }).join('')
    : '<div class="pr-empty-row">本期無流失商機</div>';
}

// ── 合約資料載入（共用）──────────────────────────────────
async function fetchAllContracts() {
  const res = await fetch(`${API}/contracts`);
  allContracts = await res.json();
}

async function loadErpMaView() {
  try {
    await fetchAllContracts();
    const list = allContracts.filter(c => !c.type || c.type === 'ERP_MA');
    renderContractStat(list, 'contractStatRow');
    renderContractTable(list, 'contractTbody',
      $('contractSearch').value, $('contractStatusFilter').value);
  } catch { showToast('無法載入合約資料'); }
}

async function loadSapMaView() {
  try {
    await fetchAllContracts();
    const list = allContracts.filter(c => c.type === 'SAP_MA');
    renderSapNoticeBar(list);
    renderContractStat(list, 'sapStatRow');
    renderContractTable(list, 'sapContractTbody',
      $('sapContractSearch').value, $('sapContractStatusFilter').value);
  } catch { showToast('無法載入合約資料'); }
}

// ── 年度通知判斷 ─────────────────────────────────────────
// PCE / PE License MA：起迄日超過1年 → 每年到期前90天顯示發票開立通知
// PCOE License MA：合約到期日前90天顯示客戶端續約通知
function getAnnualNotification(c) {
  if (!c.product || !c.endDate) return null;
  const isPCE  = c.product.includes('PCE License MA') && !c.product.includes('PCOE');
  const isPE   = c.product.includes('PE License MA')  && !c.product.includes('PCE') && !c.product.includes('PCOE');
  const isPCOE = c.product.includes('PCOE License MA');
  const isInvoiceType = isPCE || isPE;
  if (!isInvoiceType && !isPCOE) return null;

  const today = new Date(); today.setHours(0,0,0,0);
  const end   = new Date(c.endDate);

  // ── PCOE：直接判斷合約到期日前90天 ──────────────────────
  if (isPCOE) {
    const daysUntil = Math.ceil((end - today) / 86400000);
    if (daysUntil < 0 || daysUntil > 90) return null;
    return {
      type:     'renewal',
      label:    '🔄 客戶端續約通知',
      color:    '#e65100',
      bg:       '#fff3e0',
      nextDate: c.endDate,
      daysLeft: daysUntil
    };
  }

  // ── PCE / PE：合約超過1年，以年度到期日前90天通知 ───────
  if (c.startDate) {
    const durationDays = (end - new Date(c.startDate)) / 86400000;
    if (durationDays <= 365) return null;
  }

  const m = end.getMonth(), d = end.getDate();
  let nextAnnual = new Date(today.getFullYear(), m, d);
  if (nextAnnual <= today) nextAnnual = new Date(today.getFullYear() + 1, m, d);

  if (nextAnnual > end) return null;

  const daysUntil = Math.ceil((nextAnnual - today) / 86400000);
  if (daysUntil > 90) return null;

  return {
    type:     'invoice',
    label:    '📄 發票開立通知',
    color:    '#1565c0',
    bg:       '#e3f2fd',
    nextDate: nextAnnual.toISOString().slice(0, 10),
    daysLeft: daysUntil
  };
}

// ── 通知橫幅（SAP MA 專用）──────────────────────────────
function renderSapNoticeBar(list) {
  const bar = $('sapNoticeBar');
  if (!bar) return;
  const notices = list
    .map(c => ({ c, n: getAnnualNotification(c) }))
    .filter(x => x.n !== null);

  if (notices.length === 0) { bar.innerHTML = ''; bar.style.display = 'none'; return; }

  bar.style.display = 'block';
  bar.innerHTML = `
    <div class="sap-notice-title">🔔 年度合約通知（${notices.length} 筆）</div>
    ${notices.map(({ c, n }) => `
      <div class="sap-notice-item" style="border-left:4px solid ${n.color};background:${n.bg}">
        <span class="sap-notice-badge" style="color:${n.color}">${n.label}</span>
        <span class="sap-notice-co">${c.company || '—'}</span>
        <span class="sap-notice-prod">【${c.product}】</span>
        年度到期日：<strong>${n.nextDate}</strong>，
        還有 <strong style="color:${n.daysLeft <= 30 ? '#c62828' : n.color}">${n.daysLeft} 天</strong>
        ${c.contractNo ? `<span class="sap-notice-no">（${c.contractNo}）</span>` : ''}
      </div>`).join('')}`;
}

async function reloadCurrentContractView() {
  if (currentSection === 'erp-ma') await loadErpMaView();
  else if (currentSection === 'sap-ma') await loadSapMaView();
}

function renderContractStat(all, statRowId) {
  const active   = all.filter(c => contractStatus(c).key === 'active').length;
  const expiring = all.filter(c => ['expiring','urgent'].includes(contractStatus(c).key)).length;
  const expired  = all.filter(c => contractStatus(c).key === 'expired').length;
  const totAmt   = all.reduce((s, c) => s + (parseFloat(c.amount) || 0), 0);

  $(statRowId).innerHTML = `
    <div class="contract-stat-card csc-total">
      <div class="csc-label">合約總數</div>
      <div class="csc-value">${all.length}</div>
      <div class="csc-sub">年費合計 NT$${totAmt.toLocaleString()}K</div>
    </div>
    <div class="contract-stat-card csc-active">
      <div class="csc-label">&#9989; 有效合約</div>
      <div class="csc-value">${active}</div>
    </div>
    <div class="contract-stat-card csc-expiring">
      <div class="csc-label">&#9888; 即將到期（90天內）</div>
      <div class="csc-value">${expiring}</div>
    </div>
    <div class="contract-stat-card csc-expired">
      <div class="csc-label">&#10060; 逾期</div>
      <div class="csc-value">${expired}</div>
    </div>`;
}

function renderContractTable(baseList, tbodyId, search = '', filterStatus = '') {
  let list = [...baseList];

  // 排序：逾期最後，urgent/expiring 優先，依有效到期日
  list.sort((a, b) => {
    const sa = contractStatus(a), sb = contractStatus(b);
    const order = { urgent: 0, expiring: 1, active: 2, expired: 3 };
    if (order[sa.key] !== order[sb.key]) return order[sa.key] - order[sb.key];
    return (a.endDate || '').localeCompare(b.endDate || '');
  });

  if (search) {
    const kw = search.toLowerCase();
    list = list.filter(c =>
      (c.company || '').toLowerCase().includes(kw) ||
      (c.contractNo || '').toLowerCase().includes(kw) ||
      (c.product || '').toLowerCase().includes(kw) ||
      (c.salesPerson || '').toLowerCase().includes(kw)
    );
  }
  if (filterStatus) {
    if (filterStatus === 'expiring') {
      list = list.filter(c => ['expiring','urgent'].includes(contractStatus(c).key));
    } else {
      list = list.filter(c => contractStatus(c).key === filterStatus);
    }
  }

  const tbody = $(tbodyId);
  if (list.length === 0) {
    tbody.innerHTML = `<tr><td colspan="10" style="text-align:center;padding:40px;color:#bbb">
      尚無合約資料，點擊「+ 新增合約」建立第一筆</td></tr>`;
    return;
  }

  tbody.innerHTML = '';
  list.forEach(c => {
    const st = contractStatus(c);
    const badgeCls = {
      active:   'ct-badge-active',
      expiring: 'ct-badge-expiring',
      urgent:   'ct-badge-urgent',
      expired:  'ct-badge-expired'
    }[st.key] || 'ct-badge-active';
    const rowCls = { urgent:'ct-urgent', expiring:'ct-expiring', expired:'ct-expired' }[st.key] || '';

    // 到期日欄：原合約迄日 + 續約資訊 + 進度條 + 剩餘天數
    let endCellHtml = c.endDate || '—';
    if (st.isRenewed) {
      endCellHtml += `<div class="ct-renew-label">→ 續約至: ${c.renewDate}</div>`;
    }
    if (st.days !== null) {
      const barHtml = `<div class="ct-progress-wrap">
        <div class="ct-progress-bar ct-progress-${st.color}" style="width:${st.pct}%"></div>
      </div>`;
      const dayCls  = st.key === 'expired' ? 'overdue' : (st.days <= 25 ? 'urgent' : '');
      const dayText = st.key === 'expired'
        ? `已逾期 ${st.days} 天`
        : `剩餘 ${st.days} 天`;
      endCellHtml += barHtml + `<div class="ct-days ${dayCls}">${dayText}</div>`;
    }

    const tr = document.createElement('tr');
    tr.className = rowCls;
    tr.innerHTML = `
      <td><strong>${c.contractNo || '—'}</strong></td>
      <td><strong>${c.company || '—'}</strong>${c.contactName ? `<div style="font-size:11px;color:#888">${c.contactName}</div>` : ''}</td>
      <td>${c.product || '—'}</td>
      <td>${c.startDate || '—'}</td>
      <td>${endCellHtml}</td>
      <td style="text-align:right">
        ${c.amount ? Number(c.amount).toLocaleString() : '—'}
        ${c.tcv ? `<div style="font-size:11px;color:#1565c0;font-weight:600;margin-top:2px">TCV ${Number(c.tcv).toLocaleString()}K</div>` : ''}
      </td>
      <td>${c.salesPerson || '—'}</td>
      <td>
        <span class="ct-badge ${badgeCls}">${st.label}</span>
        ${(() => { const n = getAnnualNotification(c); return n ? `<div class="ct-notice-badge" style="color:${n.color};background:${n.bg}">${n.label}<br><span style="font-weight:400">${n.daysLeft}天後到期</span></div>` : ''; })()}
      </td>
      <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis" title="${c.note || ''}">${c.note || '—'}</td>
      <td>
        <div class="ct-actions">
          <button class="ct-btn ct-btn-edit" data-id="${c.id}">編輯</button>
          <button class="ct-btn ct-btn-renew" data-id="${c.id}">續約</button>
          <button class="ct-btn ct-btn-delete" data-id="${c.id}">刪除</button>
        </div>
      </td>`;
    tbody.appendChild(tr);
  });

  // 事件綁定
  tbody.querySelectorAll('.ct-btn-edit').forEach(btn =>
    btn.addEventListener('click', () => openContractModal(btn.dataset.id)));
  tbody.querySelectorAll('.ct-btn-renew').forEach(btn =>
    btn.addEventListener('click', () => renewContract(btn.dataset.id)));
  tbody.querySelectorAll('.ct-btn-delete').forEach(btn =>
    btn.addEventListener('click', () => confirmDeleteContract(btn.dataset.id)));
}

// ── 多年合約費用明細 ──────────────────────────────────────
function updateMultiYearSection(prefillAmounts = []) {
  const start = $('contractStartDate').value;
  const end   = $('contractEndDate').value;
  const sec   = $('contractMultiYearSection');

  if (!start || !end) { sec.style.display = 'none'; return; }

  const diffDays = (new Date(end) - new Date(start)) / 86400000;
  if (diffDays <= 365) { sec.style.display = 'none'; return; }

  const years = Math.ceil(diffDays / 365);
  sec.style.display = '';

  const fieldsDiv = $('contractYearFields');
  // 保留已填數值（切換日期時不清空）
  const existing = [];
  fieldsDiv.querySelectorAll('.year-amount-input').forEach(i => existing.push(i.value));

  fieldsDiv.innerHTML = '';
  for (let i = 1; i <= years; i++) {
    const val = prefillAmounts[i - 1] ?? existing[i - 1] ?? '';
    const row = document.createElement('div');
    row.className = 'year-amount-row';
    row.innerHTML = `
      <label>第 ${i} 年（NT$K）</label>
      <input type="number" class="year-amount-input" data-year="${i}"
        value="${val}" placeholder="0" min="0">`;
    fieldsDiv.appendChild(row);
  }

  fieldsDiv.querySelectorAll('.year-amount-input').forEach(inp =>
    inp.addEventListener('input', calcTCV));
  calcTCV();
}

function calcTCV() {
  let total = 0;
  $('contractYearFields').querySelectorAll('.year-amount-input').forEach(inp => {
    total += parseFloat(inp.value) || 0;
  });
  $('contractTCV').value = total > 0 ? total : '';
}

$('contractStartDate').addEventListener('change', () => updateMultiYearSection());
$('contractEndDate').addEventListener('change',   () => updateMultiYearSection());

// ── 合約 Modal ───────────────────────────────────────────
function openContractModal(id = null, type = null) {
  const isEdit = !!id;
  const contractType = type || (id ? (allContracts.find(x => x.id === id)?.type || 'ERP_MA') : 'ERP_MA');
  const typeLabel = contractType === 'SAP_MA' ? 'SAP License MA' : 'ERP MA';
  $('contractModalTitle').textContent = isEdit ? `編輯 ${typeLabel} 合約` : `新增 ${typeLabel} 合約`;
  $('contractTypeHidden').value = contractType;
  $('contractId').value = id || '';

  if (isEdit) {
    const c = allContracts.find(x => x.id === id);
    if (!c) return;
    $('contractNo').value          = c.contractNo   || '';
    $('contractCompany').value     = c.company      || '';
    $('contractContact').value     = c.contactName  || '';
    $('contractProduct').value     = c.product      || '';
    $('contractStartDate').value   = c.startDate    || '';
    $('contractEndDate').value     = c.endDate      || '';
    $('contractRenewDate').value   = c.renewDate    || '';
    $('contractAmount').value      = c.amount       || '';
    $('contractSalesPerson').value = c.salesPerson  || '';
    $('contractNote').value        = c.note         || '';
    updateMultiYearSection(c.yearAmounts || []);
    if (c.tcv) $('contractTCV').value = c.tcv;
  } else {
    ['contractNo','contractCompany','contractContact','contractAmount','contractSalesPerson','contractNote'].forEach(id => $(id).value = '');
    $('contractProduct').value    = '';
    $('contractStartDate').value  = '';
    $('contractEndDate').value    = '';
    $('contractRenewDate').value  = '';
    $('contractMultiYearSection').style.display = 'none';
    $('contractYearFields').innerHTML = '';
    $('contractTCV').value = '';
  }
  $('contractModal').classList.add('open');
}

function closeContractModal() { $('contractModal').classList.remove('open'); }

$('contractModalClose').addEventListener('click', closeContractModal);
$('contractCancelBtn').addEventListener('click',  closeContractModal);
$('contractModal').addEventListener('click', e => { if (e.target === $('contractModal')) closeContractModal(); });

$('contractSaveBtn').addEventListener('click', async () => {
  try {
    const company = ($('contractCompany').value || '').trim();
    if (!company) { showToast('請填入客戶名稱'); return; }
    const endDate = ($('contractEndDate').value || '').trim();
    if (!endDate) { showToast('請填入合約迄日'); return; }

    const id = $('contractId').value || '';
    const yearAmountInputs = $('contractYearFields')
      ? Array.from($('contractYearFields').querySelectorAll('.year-amount-input'))
      : [];

    const payload = {
      contractNo:  ($('contractNo').value || '').trim(),
      company,
      contactName: ($('contractContact').value || '').trim(),
      product:     $('contractProduct').value || '',
      startDate:   $('contractStartDate').value || '',
      endDate,
      renewDate:   $('contractRenewDate').value || '',
      amount:      $('contractAmount').value || '',
      yearAmounts: yearAmountInputs.map(i => parseFloat(i.value) || 0),
      tcv:         parseFloat($('contractTCV').value) || null,
      salesPerson: ($('contractSalesPerson').value || '').trim(),
      note:        ($('contractNote').value || '').trim(),
      type:        ($('contractTypeHidden').value || 'ERP_MA')
    };

    if (id) {
      const r = await fetch(`${API}/contracts/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (!r.ok) {
        const e = await r.json().catch(() => ({}));
        showToast('❌ 儲存失敗：' + (e.error || r.status));
        return;
      }
      showToast('合約已更新');
    } else {
      const r = await fetch(`${API}/contracts`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (!r.ok) {
        const e = await r.json().catch(() => ({}));
        showToast('❌ 新增失敗：' + (e.error || r.status));
        return;
      }
      showToast('合約已新增');
    }
    closeContractModal();
    await reloadCurrentContractView();
  } catch (err) {
    console.error('[contractSave]', err);
    showToast('❌ 儲存錯誤：' + (err.message || '請重試'));
  }
});

// ── 續約（複製並延長一年）────────────────────────────────
async function renewContract(id) {
  const c = allContracts.find(x => x.id === id);
  if (!c) return;
  const newStart = c.endDate ? (() => {
    const d = new Date(c.endDate); d.setDate(d.getDate() + 1);
    return d.toISOString().slice(0,10);
  })() : '';
  const newEnd = c.endDate ? (() => {
    const d = new Date(c.endDate); d.setFullYear(d.getFullYear() + 1);
    return d.toISOString().slice(0,10);
  })() : '';
  // 開啟新增 modal 並預填資料
  const typeLabel = c.type === 'SAP_MA' ? 'SAP License MA' : 'ERP MA';
  openContractModal(null, c.type || 'ERP_MA');
  $('contractNo').value          = c.contractNo ? c.contractNo + '-R' : '';
  $('contractCompany').value     = c.company      || '';
  $('contractContact').value     = c.contactName  || '';
  $('contractProduct').value     = c.product      || '';
  $('contractStartDate').value   = newStart;
  $('contractEndDate').value     = newEnd;
  $('contractAmount').value      = c.amount       || '';
  $('contractSalesPerson').value = c.salesPerson  || '';
  $('contractNote').value        = `（續約自 ${c.contractNo || id}）`;
  $('contractModalTitle').textContent = `續約 — ${typeLabel} 合約`;
}

// ── 刪除確認 ─────────────────────────────────────────────
function confirmDeleteContract(id) {
  contractDeleteId = id;
  $('contractDeleteModal').classList.add('open');
}
$('contractDeleteClose').addEventListener('click',  () => $('contractDeleteModal').classList.remove('open'));
$('contractDeleteCancel').addEventListener('click', () => $('contractDeleteModal').classList.remove('open'));
$('contractDeleteModal').addEventListener('click', e => { if (e.target === $('contractDeleteModal')) $('contractDeleteModal').classList.remove('open'); });
$('contractDeleteConfirm').addEventListener('click', async () => {
  if (!contractDeleteId) return;
  try {
    await fetch(`${API}/contracts/${contractDeleteId}`, { method: 'DELETE' });
    $('contractDeleteModal').classList.remove('open');
    contractDeleteId = null;
    showToast('合約已刪除');
    await reloadCurrentContractView();
  } catch { showToast('刪除失敗'); }
});

// ── 搜尋 & 篩選 ─── ERP MA ───────────────────────────────
$('contractSearch').addEventListener('input', () => {
  const list = allContracts.filter(c => !c.type || c.type === 'ERP_MA');
  renderContractTable(list, 'contractTbody', $('contractSearch').value, $('contractStatusFilter').value);
});
$('contractStatusFilter').addEventListener('change', () => {
  const list = allContracts.filter(c => !c.type || c.type === 'ERP_MA');
  renderContractTable(list, 'contractTbody', $('contractSearch').value, $('contractStatusFilter').value);
});
$('addContractBtn').addEventListener('click', () => openContractModal(null, 'ERP_MA'));

// ── 搜尋 & 篩選 ─── SAP MA ───────────────────────────────
$('sapContractSearch').addEventListener('input', () => {
  const list = allContracts.filter(c => c.type === 'SAP_MA');
  renderContractTable(list, 'sapContractTbody', $('sapContractSearch').value, $('sapContractStatusFilter').value);
});
$('sapContractStatusFilter').addEventListener('change', () => {
  const list = allContracts.filter(c => c.type === 'SAP_MA');
  renderContractTable(list, 'sapContractTbody', $('sapContractSearch').value, $('sapContractStatusFilter').value);
});
$('addSapContractBtn').addEventListener('click', () => openContractModal(null, 'SAP_MA'));

// ── 初始化 ───────────────────────────────────────────────
// ── 手機漢堡選單 ──────────────────────────────────────────
(function initMobileMenu() {
  const btn      = $('mobileMenuBtn');
  const sidebar  = document.querySelector('.sidebar');
  const overlay  = $('sidebarOverlay');
  if (!btn || !sidebar || !overlay) return;

  function openSidebar()  { sidebar.classList.add('mobile-open');  overlay.classList.add('show'); }
  function closeSidebar() { sidebar.classList.remove('mobile-open'); overlay.classList.remove('show'); }

  btn.addEventListener('click', openSidebar);
  overlay.addEventListener('click', closeSidebar);

  // 點選導覽項目後自動關閉 sidebar（手機）
  document.querySelectorAll('.sidebar-nav-item:not(.sidebar-nav-group-hd)').forEach(el => {
    el.addEventListener('click', () => { if (window.innerWidth <= 900) closeSidebar(); });
  });
})();

initUser();
loadPermissions();
loadContacts(); // 背景載入聯絡人，供圖表與拜訪記錄使用

// ── 側邊欄拖動排序 ─────────────────────────────────────────
(function initSidebarDrag() {
  var nav = document.querySelector('.sidebar-nav');
  if (!nav) return;

  var STORAGE_KEY = 'sidebarNavOrder';
  var dragSrc = null;

  // 取得可拖動的直接子元素（排除 display:none 的）
  function getNavChildren() {
    return Array.from(nav.children).filter(function(el) {
      return el.style.display !== 'none';
    });
  }

  function saveOrder() {
    var order = Array.from(nav.children).map(function(el) {
      return el.id || '';
    });
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(order)); } catch(e) {}
  }

  function restoreOrder() {
    try {
      var stored = localStorage.getItem(STORAGE_KEY);
      if (!stored) return;
      var order = JSON.parse(stored);
      if (!Array.isArray(order) || !order.length) return;
      order.forEach(function(id) {
        if (!id) return;
        var el = document.getElementById(id);
        if (el && el.parentNode === nav) nav.appendChild(el);
      });
    } catch(e) {}
  }

  function clearIndicators() {
    Array.from(nav.children).forEach(function(el) {
      el.classList.remove('sidebar-drag-over-top', 'sidebar-drag-over-bottom', 'sidebar-drag-src');
    });
  }

  function onDragStart(e) {
    dragSrc = this;
    this.classList.add('sidebar-drag-src');
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', this.id || '');
  }

  function onDragOver(e) {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    if (!dragSrc || this === dragSrc) return;
    clearIndicators();
    dragSrc.classList.add('sidebar-drag-src');
    // 判斷滑鼠在元素上半或下半
    var rect = this.getBoundingClientRect();
    var midY = rect.top + rect.height / 2;
    if (e.clientY < midY) {
      this.classList.add('sidebar-drag-over-top');
    } else {
      this.classList.add('sidebar-drag-over-bottom');
    }
  }

  function onDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    if (!dragSrc || dragSrc === this) { clearIndicators(); return; }
    var rect = this.getBoundingClientRect();
    var midY = rect.top + rect.height / 2;
    if (e.clientY < midY) {
      nav.insertBefore(dragSrc, this);
    } else {
      nav.insertBefore(dragSrc, this.nextSibling);
    }
    clearIndicators();
    saveOrder();
  }

  function onDragEnd() {
    clearIndicators();
    dragSrc = null;
  }

  function applyDrag() {
    Array.from(nav.children).forEach(function(el) {
      if (!el._dragBound) {
        el._dragBound = true;
        el.setAttribute('draggable', 'true');
        el.addEventListener('dragstart', onDragStart);
        el.addEventListener('dragover',  onDragOver);
        el.addEventListener('drop',      onDrop);
        el.addEventListener('dragend',   onDragEnd);
      }
    });
  }

  restoreOrder();
  applyDrag();
})();

// 還原上次所在頁面
(function restoreLastSection() {
  const last = localStorage.getItem('lastSection');
  if (!last || last === 'dashboard') {
    showDashboard();
  } else {
    showSection(last);
  }
})();

// ════════════════════════════════════════════════════════
// ── 通知系統 ─────────────────────────────────────────────
// ════════════════════════════════════════════════════════
let _notifPollTimer = null;

// 合約提醒已讀 key 格式：contract_{id}_{YYYY-MM-DD}
function _contractReadKey(id) {
  return `cr_read_${id}_${new Date().toISOString().slice(0,10)}`;
}
function isContractReminderRead(id) {
  return !!localStorage.getItem(_contractReadKey(id));
}
function markContractReminderRead(id) {
  localStorage.setItem(_contractReadKey(id), '1');
}

// ── 生日提醒已讀狀態（每日重置）────────────────────────────
function _birthdayReadKey(id) {
  return `bday_read_${id}_${new Date().toISOString().slice(0,10)}`;
}
function isBirthdayReminderRead(id) {
  return !!localStorage.getItem(_birthdayReadKey(id));
}
function markBirthdayReminderRead(id) {
  localStorage.setItem(_birthdayReadKey(id), '1');
}

async function pollNotifications() {
  try {
    // 1. 一般通知
    const r1 = await fetch(`${API}/notifications`);
    const regularList = r1.ok ? await r1.json() : [];

    // 2. 合約到期提醒（靜默失敗）
    let contractList = [];
    try {
      const r2 = await fetch(`${API}/contract-reminders`);
      if (r2.ok) contractList = await r2.json();
    } catch {}

    // 3. 生日提醒（靜默失敗）
    let birthdayList = [];
    try {
      const r3 = await fetch(`${API}/birthday-reminders?days=3`);
      if (r3.ok) birthdayList = await r3.json();
    } catch {}

    // 合約提醒轉換成統一格式（加 isContract flag 供 render 識別）
    const contractNotifs = contractList.map(c => ({
      id:        c.id,
      type:      c.type,
      title:     c.title,
      body:      c.body,
      read:      isContractReminderRead(c.id),
      createdAt: null,
      isContract: true,
      days:      c.days,
    }));

    // 生日提醒轉換成統一格式
    const birthdayNotifs = birthdayList.map(p => {
      const bdayId = `bday_${p.id}`;
      const whenTxt = p.daysLeft === 0 ? '今天！🎉' : p.daysLeft === 1 ? '明天' : `${p.daysLeft} 天後`;
      return {
        id:         bdayId,
        type:       'birthday',
        title:      `🎂 ${p.name} 生日${p.daysLeft === 0 ? '快樂！' : `即將到來（${whenTxt}）`}`,
        body:       `${p.company ? p.company + '・' : ''}${p.title ? p.title + '・' : ''}${p.personalBirthday}` +
                    (p.ownerName ? `　負責：${p.ownerName}` : ''),
        read:       isBirthdayReminderRead(bdayId),
        createdAt:  null,
        isBirthday: true,
        daysLeft:   p.daysLeft,
        contactId:  p.id
      };
    });

    // 合併：生日 → 合約 → 一般通知
    const merged = [...birthdayNotifs, ...contractNotifs, ...regularList];

    // 未讀計算
    const unreadCount = merged.filter(n => !n.read).length;
    const badge = $('notifBadge');
    if (unreadCount > 0) {
      badge.textContent = unreadCount > 9 ? '9+' : unreadCount;
      badge.style.display = '';
    } else {
      badge.style.display = 'none';
    }

    renderNotifList(merged);
  } catch {}
}

function renderNotifList(list) {
  const el = $('notifList');
  if (!list.length) { el.innerHTML = '<div class="notif-empty">暫無通知</div>'; return; }

  const typeIcon = {
    callin_new:'📞', callin_assigned:'📋', callin_overdue:'⏰', callin_responded:'✅',
    contract_urgent:'🟠', contract_expiring:'🟡', contract_expired:'🔴'
  };

  // 分組：生日 / 合約提醒 / 一般通知
  const birthdayItems  = list.filter(n => n.isBirthday);
  const contractItems  = list.filter(n => n.isContract);
  const regularItems   = list.filter(n => !n.isContract && !n.isBirthday);

  let html = '';

  // ── 生日提醒群組 ──
  if (birthdayItems.length) {
    html += `<div class="notif-group-hd">🎂 客戶生日提醒</div>`;
    html += birthdayItems.map(n => {
      const urgCls = n.daysLeft === 0 ? 'notif-bday-today'
                   : n.daysLeft === 1 ? 'notif-bday-soon'
                   : 'notif-bday-upcoming';
      return `
        <div class="notif-item notif-bday ${urgCls} ${n.read ? '' : 'unread'}" data-id="${escapeHtml(n.id)}" data-birthday="1">
          <div class="notif-icon">${n.daysLeft === 0 ? '🎉' : '🎂'}</div>
          <div class="notif-content">
            <div class="notif-title">${escapeHtml(n.title)}</div>
            <div class="notif-body">${escapeHtml(n.body)}</div>
          </div>
        </div>`;
    }).join('');
  }

  if (contractItems.length) {
    html += `<div class="notif-group-hd">📄 合約到期提醒</div>`;
    html += contractItems.map(n => {
      const urgClass = n.type === 'contract_expired'  ? 'notif-contract-expired'
                     : n.type === 'contract_urgent'   ? 'notif-contract-urgent'
                     : 'notif-contract-expiring';
      return `
        <div class="notif-item notif-contract ${urgClass} ${n.read ? '' : 'unread'}" data-id="${escapeHtml(n.id)}" data-contract="1">
          <div class="notif-icon">${typeIcon[n.type] || '📄'}</div>
          <div class="notif-content">
            <div class="notif-title">${escapeHtml(n.title)}</div>
            <div class="notif-body">${escapeHtml(n.body)}</div>
            ${n.days != null ? `<div class="notif-days-bar">
              <div class="notif-days-fill ${urgClass}" style="width:${Math.min(100, Math.max(4, 100 - n.days))}%"></div>
            </div>` : ''}
          </div>
        </div>`;
    }).join('');
  }

  if (regularItems.length) {
    if (contractItems.length) html += `<div class="notif-group-hd">🔔 系統通知</div>`;
    html += regularItems.slice(0, 20).map(n => `
      <div class="notif-item ${n.read ? '' : 'unread'}" data-id="${escapeHtml(n.id)}">
        <div class="notif-icon">${typeIcon[n.type] || '🔔'}</div>
        <div class="notif-content">
          <div class="notif-title">${escapeHtml(n.title)}</div>
          <div class="notif-body">${escapeHtml(n.body)}</div>
          <div class="notif-time">${n.createdAt ? new Date(n.createdAt).toLocaleString('zh-TW') : ''}</div>
        </div>
      </div>`).join('');
  }

  el.innerHTML = html;

  // 點擊已讀
  el.querySelectorAll('.notif-item').forEach(item => {
    item.addEventListener('click', async () => {
      const id = item.dataset.id;
      if (item.dataset.birthday) {
        markBirthdayReminderRead(id);
        item.classList.remove('unread');
        pollNotifications();
      } else if (item.dataset.contract) {
        markContractReminderRead(id);
        item.classList.remove('unread');
        pollNotifications();
      } else {
        await fetch(`${API}/notifications/${id}/read`, { method: 'PUT' });
        item.classList.remove('unread');
        pollNotifications();
      }
    });
  });
}

$('notifReadAll').addEventListener('click', async () => {
  // 生日提醒全標已讀（localStorage）
  document.querySelectorAll('.notif-item[data-birthday="1"]').forEach(item => {
    markBirthdayReminderRead(item.dataset.id);
  });
  // 合約提醒全標已讀（localStorage）
  document.querySelectorAll('.notif-item[data-contract="1"]').forEach(item => {
    markContractReminderRead(item.dataset.id);
  });
  await fetch(`${API}/notifications/read-all`, { method: 'PUT' });
  pollNotifications();
});

$('notifBellBtn').addEventListener('click', () => {
  const dd = $('notifDropdown');
  dd.style.display = dd.style.display === 'none' ? '' : 'none';
  if (dd.style.display !== 'none') pollNotifications();
});


document.addEventListener('click', e => {
  if (!$('notifBellWrap').contains(e.target)) $('notifDropdown').style.display = 'none';
});

// 每 30 秒輪詢一次
_notifPollTimer = setInterval(pollNotifications, 30000);
pollNotifications(); // 立即執行一次

// ════════════════════════════════════════════════════════
// ── 應收帳款管理 ──────────────────────────────────────────
// ════════════════════════════════════════════════════════
let allReceivables = [];
let currentRtab = 'all';

async function loadReceivablesView() {
  try {
    const r = await fetch(`${API}/receivables`);
    if (!r.ok) return;
    allReceivables = await r.json();
  } catch { showToast('無法載入帳款資料'); return; }
  renderReceivables();
}

function overdueDays(dueDateStr) {
  if (!dueDateStr) return 0;
  const diff = Date.now() - new Date(dueDateStr).getTime();
  return Math.max(0, Math.floor(diff / 86400000));
}

function receivableStatus(item) {
  const balance = item.amount - (item.paidAmount || 0);
  if (balance <= 0) return 'paid';
  if (overdueDays(item.dueDate) > 0) return 'overdue';
  return 'pending';
}

function renderReceivables() {
  const today = new Date();
  const list = allReceivables.filter(r => {
    const st = receivableStatus(r);
    if (currentRtab === 'all') return true;
    return st === currentRtab;
  });

  // 摘要統計
  const totalAmt = allReceivables.reduce((s, r) => s + (r.amount - (r.paidAmount || 0)), 0);
  const overdueAmt = allReceivables.filter(r => receivableStatus(r) === 'overdue')
    .reduce((s, r) => s + (r.amount - (r.paidAmount || 0)), 0);
  $('receivablesSummary').innerHTML = `
    <div class="recv-stat"><span class="recv-stat-label">總應收餘額</span><span class="recv-stat-val">${totalAmt.toLocaleString()}</span></div>
    <div class="recv-stat overdue-stat"><span class="recv-stat-label">逾期金額</span><span class="recv-stat-val" style="color:#d32f2f">${overdueAmt.toLocaleString()}</span></div>
    <div class="recv-stat"><span class="recv-stat-label">逾期筆數</span><span class="recv-stat-val">${allReceivables.filter(r => receivableStatus(r) === 'overdue').length}</span></div>`;

  const STATUS_LABEL = { paid:'已收款', pending:'待收款', overdue:'逾期' };
  const STATUS_CLS   = { paid:'badge-paid', pending:'badge-pending', overdue:'badge-overdue' };
  $('receivablesTbody').innerHTML = list.map(r => {
    const st = receivableStatus(r);
    const days = overdueDays(r.dueDate);
    const balance = r.amount - (r.paidAmount || 0);
    return `<tr>
      <td>${r.company || '-'}</td>
      <td>${r.contactName || '-'}</td>
      <td>${r.invoiceNo || '-'}</td>
      <td>${r.invoiceDate || '-'}</td>
      <td>${r.dueDate || '-'}</td>
      <td>${st === 'overdue' ? `<span style="color:#d32f2f;font-weight:700">${days} 天</span>` : '-'}</td>
      <td>${r.amount.toLocaleString()} ${r.currency || 'NTD'}</td>
      <td>${balance.toLocaleString()} ${r.currency || 'NTD'}</td>
      <td><span class="recv-badge ${STATUS_CLS[st]}">${STATUS_LABEL[st]}</span></td>
      <td>
        <button class="btn btn-sm" onclick="openEditReceivable('${r.id}')">編輯</button>
        <button class="btn btn-sm btn-danger" onclick="deleteReceivable('${r.id}')">刪除</button>
      </td>
    </tr>`;
  }).join('') || `<tr><td colspan="10" style="text-align:center;color:#bbb;padding:32px">暫無資料</td></tr>`;
}

// Tab 切換
document.querySelectorAll('[data-rtab]').forEach(btn => {
  btn.addEventListener('click', function() {
    document.querySelectorAll('[data-rtab]').forEach(b => b.classList.remove('active'));
    this.classList.add('active');
    currentRtab = this.dataset.rtab;
    renderReceivables();
  });
});

// 新增帳款
$('addReceivableBtn').addEventListener('click', () => openReceivableModal());

function openReceivableModal(id) {
  const item = id ? allReceivables.find(r => r.id === id) : null;
  $('receivableId').value = id || '';
  $('receivableModalTitle').textContent = id ? '編輯帳款' : '新增帳款';
  $('rCompany').value   = item ? item.company : '';
  $('rContact').value   = item ? item.contactName : '';
  $('rInvoiceNo').value = item ? item.invoiceNo : '';
  $('rCurrency').value  = item ? (item.currency || 'NTD') : 'NTD';
  $('rInvoiceDate').value = item ? item.invoiceDate : '';
  $('rDueDate').value   = item ? item.dueDate : '';
  $('rAmount').value    = item ? item.amount : '';
  $('rPaidAmount').value= item ? (item.paidAmount || 0) : 0;
  $('rNote').value      = item ? (item.note || '') : '';
  $('receivableModalOverlay').style.display = 'flex';
}
function openEditReceivable(id) { openReceivableModal(id); }
$('receivableModalClose').addEventListener('click', () => $('receivableModalOverlay').style.display = 'none');
$('receivableModalCancel').addEventListener('click', () => $('receivableModalOverlay').style.display = 'none');

$('receivableModalSave').addEventListener('click', async () => {
  const id = $('receivableId').value;
  const payload = {
    company:     $('rCompany').value.trim(),
    contactName: $('rContact').value.trim(),
    invoiceNo:   $('rInvoiceNo').value.trim(),
    currency:    $('rCurrency').value,
    invoiceDate: $('rInvoiceDate').value,
    dueDate:     $('rDueDate').value,
    amount:      parseFloat($('rAmount').value) || 0,
    paidAmount:  parseFloat($('rPaidAmount').value) || 0,
    note:        $('rNote').value.trim()
  };
  if (!payload.company) { showToast('請填入公司名稱'); return; }
  if (!payload.amount)  { showToast('請填入應收金額'); return; }
  try {
    const r = await fetch(`${API}/receivables${id ? '/'+id : ''}`, {
      method: id ? 'PUT' : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!r.ok) { showToast('儲存失敗'); return; }
    showToast(id ? '✅ 帳款已更新' : '✅ 帳款已新增');
    $('receivableModalOverlay').style.display = 'none';
    loadReceivablesView();
  } catch { showToast('操作失敗'); }
});

async function deleteReceivable(id) {
  if (!confirm('確定刪除此帳款？')) return;
  await fetch(`${API}/receivables/${id}`, { method: 'DELETE' });
  showToast('✅ 已刪除');
  loadReceivablesView();
}

// ════════════════════════════════════════════════════════
// ── Call-in Pass 管理 ─────────────────────────────────────
// ════════════════════════════════════════════════════════
let allCallins = [];
let currentCtab = 'all';
let callinUsers = []; // 可指派的業務清單

async function loadCallinView() {
  try {
    const r = await fetch(`${API}/callins`);
    if (!r.ok) return;
    allCallins = await r.json();
  } catch { showToast('無法載入 Call-in 資料'); return; }

  // 判斷角色決定是否顯示新增按鈕
  const role = userPermissions.role;
  $('addCallinBtn').style.display = (role === 'secretary' || role === 'admin' || role === 'manager1' || role === 'manager2') ? '' : 'none';

  renderCallins();
  updateCallinDashCard();
  updateCallinBadge();
}

function renderCallins() {
  const STATUS_LABEL = { pending:'待指派', assigned:'待聯繫', contacted:'已聯繫', qualified:'合格商機', unqualified:'非合格', overdue:'逾時' };
  const STATUS_CLS   = { pending:'ci-pending', assigned:'ci-assigned', contacted:'ci-contacted', qualified:'ci-qualified', unqualified:'ci-unqualified', overdue:'ci-overdue' };
  const role = userPermissions.role;

  const list = allCallins.filter(c => {
    if (currentCtab === 'all') return true;
    if (currentCtab === 'done') return ['contacted','qualified','unqualified'].includes(c.status);
    return c.status === currentCtab;
  });

  $('callinTbody').innerHTML = list.map(c => {
    const canAssign  = (role === 'manager1' || role === 'manager2' || role === 'admin') && c.status === 'pending';
    const canRespond = role === 'user' && c.status === 'assigned';
    const dl = c.deadline ? new Date(c.deadline).toLocaleString('zh-TW', {month:'numeric',day:'numeric',hour:'2-digit',minute:'2-digit'}) : '-';
    return `<tr>
      <td>${escapeHtml(new Date(c.createdAt).toLocaleString('zh-TW',{month:'numeric',day:'numeric',hour:'2-digit',minute:'2-digit'}))}</td>
      <td><strong>${escapeHtml(c.company) || ''}</strong>${c.contactName ? '<br><span style="color:#888;font-size:12px">'+escapeHtml(c.contactName)+'</span>' : ''}</td>
      <td>${escapeHtml(c.phone) || '-'}</td>
      <td style="max-width:180px;white-space:normal">${escapeHtml(c.topic) || '-'}</td>
      <td>${escapeHtml(c.assignedTo) || '<span style="color:#bbb">未指派</span>'}</td>
      <td>${escapeHtml(dl)}</td>
      <td><span class="ci-badge ${STATUS_CLS[c.status] || ''}">${STATUS_LABEL[c.status] || escapeHtml(c.status)}</span></td>
      <td>
        ${canAssign  ? `<button class="btn btn-sm btn-primary" onclick="openAssignModal('${escapeHtml(c.id)}')">指派</button>` : ''}
        ${canRespond ? `<button class="btn btn-sm btn-success" onclick="openRespondModal('${escapeHtml(c.id)}')">回應</button>` : ''}
        ${(!canAssign && !canRespond) ? '<span style="color:#bbb;font-size:12px">-</span>' : ''}
      </td>
    </tr>`;
  }).join('') || `<tr><td colspan="8" style="text-align:center;color:#bbb;padding:32px">暫無資料</td></tr>`;
}

// Tab 切換
document.querySelectorAll('[data-ctab]').forEach(btn => {
  btn.addEventListener('click', function() {
    document.querySelectorAll('[data-ctab]').forEach(b => b.classList.remove('active'));
    this.classList.add('active');
    currentCtab = this.dataset.ctab;
    renderCallins();
  });
});

// 新增 Call-in
$('addCallinBtn').addEventListener('click', () => {
  $('ciCompany').value = $('ciContact').value = $('ciPhone').value = $('ciTopic').value = $('ciNote').value = '';
  $('ciSource').value = '電話';
  $('callinModalOverlay').style.display = 'flex';
});
$('callinModalClose').addEventListener('click', () => $('callinModalOverlay').style.display = 'none');
$('callinModalCancel').addEventListener('click', () => $('callinModalOverlay').style.display = 'none');

$('callinModalSave').addEventListener('click', async () => {
  const topic = $('ciTopic').value.trim();
  if (!topic) { showToast('請填入來電事由'); return; }
  const payload = { company:$('ciCompany').value.trim(), contactName:$('ciContact').value.trim(), phone:$('ciPhone').value.trim(), topic, source:$('ciSource').value, note:$('ciNote').value.trim() };
  try {
    const r = await fetch(`${API}/callins`, { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload) });
    if (!r.ok) { const e = await r.json(); showToast(e.error || '建立失敗'); return; }
    showToast('✅ Call-in 已建立，已通知主管');
    $('callinModalOverlay').style.display = 'none';
    loadCallinView();
  } catch { showToast('操作失敗'); }
});

// 指派 Modal
async function openAssignModal(id) {
  const item = allCallins.find(c => c.id === id);
  if (!item) return;
  $('assignCallinId').value = id;
  $('assignCallinInfo').innerHTML = `<b>${escapeHtml(item.company || item.contactName)}</b> — ${escapeHtml(item.topic)}`;

  // 取得可指派業務清單
  if (!callinUsers.length) {
    try {
      const r = await fetch(`${API}/admin/users`);
      if (r.ok) callinUsers = (await r.json()).filter(u => u.role === 'user' || u.role === 'manager2');
    } catch {}
  }
  const sel = $('assignToUser');
  sel.innerHTML = '<option value="">請選擇業務...</option>' +
    callinUsers.map(u => `<option value="${u.username}">${u.displayName || u.username}</option>`).join('');
  $('callinAssignOverlay').style.display = 'flex';
}
$('callinAssignClose').addEventListener('click', () => $('callinAssignOverlay').style.display = 'none');
$('callinAssignCancel').addEventListener('click', () => $('callinAssignOverlay').style.display = 'none');

$('callinAssignSave').addEventListener('click', async () => {
  const id = $('assignCallinId').value;
  const assignedTo = $('assignToUser').value;
  if (!assignedTo) { showToast('請選擇業務'); return; }
  try {
    const r = await fetch(`${API}/callins/${id}/assign`, { method:'PUT', headers:{'Content-Type':'application/json'}, body:JSON.stringify({ assignedTo }) });
    if (!r.ok) { const e = await r.json(); showToast(e.error || '指派失敗'); return; }
    showToast('✅ 已指派，業務已收到通知');
    $('callinAssignOverlay').style.display = 'none';
    loadCallinView();
    pollNotifications();
  } catch { showToast('操作失敗'); }
});

// 業務回應 Modal
function openRespondModal(id) {
  const item = allCallins.find(c => c.id === id);
  if (!item) return;
  $('respondCallinId').value = id;
  $('respondCallinInfo').innerHTML = `<b>${escapeHtml(item.company || item.contactName)}</b> — ${escapeHtml(item.topic)}<br><span style="color:#e65100">截止：${item.deadline ? escapeHtml(new Date(item.deadline).toLocaleString('zh-TW')) : '-'}</span>`;
  $('oppName').value = $('respondNote').value = '';
  $('oppStage').value = 'C';
  document.querySelectorAll('input[name="respondAction"]').forEach(r => r.checked = false);
  $('qualifiedSection').style.display = 'none';
  $('callinRespondOverlay').style.display = 'flex';
}
$('callinRespondClose').addEventListener('click', () => $('callinRespondOverlay').style.display = 'none');
$('callinRespondCancel').addEventListener('click', () => $('callinRespondOverlay').style.display = 'none');

document.querySelectorAll('input[name="respondAction"]').forEach(r => {
  r.addEventListener('change', () => {
    $('qualifiedSection').style.display = $('raQualified').checked ? '' : 'none';
  });
});

$('callinRespondSave').addEventListener('click', async () => {
  const id = $('respondCallinId').value;
  const action = document.querySelector('input[name="respondAction"]:checked')?.value;
  if (!action) { showToast('請選擇聯繫結果'); return; }
  if (action === 'qualified' && !$('oppName').value.trim()) { showToast('請填入商機名稱'); return; }
  const payload = { action, opportunityName:$('oppName').value.trim(), opportunityStage:$('oppStage').value, note:$('respondNote').value.trim() };
  try {
    const r = await fetch(`${API}/callins/${id}/respond`, { method:'PUT', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload) });
    if (!r.ok) { const e = await r.json(); showToast(e.error || '回應失敗'); return; }
    showToast('✅ 已回應，主管已收到通知');
    $('callinRespondOverlay').style.display = 'none';
    loadCallinView();
    pollNotifications();
  } catch { showToast('操作失敗'); }
});

// Dashboard Call-in 待辦卡
function updateCallinDashCard() {
  const role = userPermissions.role;
  const pending = allCallins.filter(c => c.status === 'assigned' || c.status === 'overdue');
  const card = $('dashCallinCard');
  if (!pending.length || (role !== 'user')) { card.style.display = 'none'; return; }
  card.style.display = '';
  $('dashCallinList').innerHTML = pending.map(c => {
    const isOverdue = c.status === 'overdue';
    return `<div class="dash-callin-item ${isOverdue ? 'overdue' : ''}">
      <div class="dash-callin-ico">${isOverdue ? '⏰' : '📞'}</div>
      <div class="dash-callin-info">
        <strong>${escapeHtml(c.company || c.contactName)}</strong> — ${escapeHtml(c.topic)}
        <div style="font-size:11px;color:${isOverdue?'#d32f2f':'#e65100'}">截止：${c.deadline ? escapeHtml(new Date(c.deadline).toLocaleString('zh-TW')) : '-'}</div>
      </div>
      <button class="btn btn-sm btn-primary" onclick="showSection('callin')">處理</button>
    </div>`;
  }).join('');
}

function updateCallinBadge() {
  const pending = allCallins.filter(c => c.status === 'pending' || c.status === 'assigned' || c.status === 'overdue');
  const badge = $('callinBadge');
  if (pending.length > 0) {
    badge.textContent = pending.length > 9 ? '9+' : pending.length;
    badge.style.display = '';
  } else {
    badge.style.display = 'none';
  }
}

// 也在 dashboard 載入時更新 Call-in
const _origShowDashboard = showDashboard;
// 擴展 showDashboard 以載入 callin
const _dashCallinInit = async () => {
  try {
    const r = await fetch(`${API}/callins`);
    if (r.ok) {
      allCallins = await r.json();
      updateCallinDashCard();
      updateCallinBadge();
    }
  } catch {}
};

// ════════════════════════════════════════════════════════
// ── 責任業務名單移轉 ──────────────────────────────────────
// ════════════════════════════════════════════════════════
let transferViewLoaded = false;

async function loadTransferView() {
  if (!transferViewLoaded) {
    await populateTransferOwnerSelects();
    $('appTransferFromOwner').addEventListener('change', loadAppTransferFrom);
    $('appTransferToOwner').addEventListener('change', () => { loadAppTransferTo(); updateAppTransferUI(); });
    $('appTransferSelectAll').addEventListener('click', () => {
      const cbs = document.querySelectorAll('#appTransferFromList .transfer-cb');
      const allChecked = [...cbs].every(c => c.checked);
      cbs.forEach(cb => {
        cb.checked = !allChecked;
        cb.closest('.transfer-item').classList.toggle('selected', !allChecked);
      });
      $('appTransferSelectAll').textContent = allChecked ? '全選' : '全消';
      updateAppTransferUI();
    });
    $('appTransferConfirmBtn').addEventListener('click', doAppTransfer);
    transferViewLoaded = true;
  }
}

async function populateTransferOwnerSelects() {
  try {
    const r = await fetch(`${API}/admin/users`);
    if (!r.ok) return;
    const users = await r.json();
    const active = users.filter(u => u.active);
    ['appTransferFromOwner','appTransferToOwner'].forEach(id => {
      const sel = $(id);
      sel.innerHTML = '<option value="">-- 選取業務人員 --</option>';
      active.forEach(u => {
        const o = document.createElement('option');
        o.value = u.username;
        o.textContent = u.displayName || u.username;
        sel.appendChild(o);
      });
    });
  } catch {}
}

// 將聯絡人列表依公司分組，回傳 [{company, contacts:[]}] 陣列
function groupContactsByCompany(contacts) {
  const map = new Map();
  contacts.forEach(c => {
    const key = (c.company || '').trim() || '__no_company__';
    if (!map.has(key)) map.set(key, { company: key === '__no_company__' ? '' : key, contacts: [] });
    map.get(key).contacts.push(c);
  });
  // 有公司名稱的先排，再依名稱排序；無公司的放最後
  return [...map.values()].sort((a, b) => {
    if (!a.company && b.company) return 1;
    if (a.company && !b.company) return -1;
    return a.company.localeCompare(b.company, 'zh-TW');
  });
}

function renderAppTransferList(contacts, listId, checkable) {
  const el = $(listId);
  if (!contacts.length) {
    el.innerHTML = '<div class="transfer-item-empty">此業務目前無客戶資料</div>';
    return;
  }

  const groups = groupContactsByCompany(contacts);

  el.innerHTML = groups.map(g => {
    const companyLabel = g.company || '（無公司）';
    const names = g.contacts.map(c => escapeHtml(c.name || '？')).join('、');
    const ids = g.contacts.map(c => c.id).join(',');
    return `
    <div class="transfer-item transfer-company-item" data-ids="${escapeHtml(ids)}">
      ${checkable ? `<input type="checkbox" class="transfer-cb" data-ids="${escapeHtml(ids)}">` : ''}
      <div class="transfer-item-content">
        <div class="transfer-item-name">${escapeHtml(companyLabel)}</div>
        <div class="transfer-item-company">${g.contacts.length} 位聯絡人：${names}</div>
      </div>
    </div>`;
  }).join('');

  if (checkable) {
    el.querySelectorAll('.transfer-item').forEach(item => {
      item.addEventListener('click', e => {
        if (e.target.type === 'checkbox') return;
        const cb = item.querySelector('.transfer-cb');
        cb.checked = !cb.checked;
        item.classList.toggle('selected', cb.checked);
        updateAppTransferUI();
      });
      item.querySelector('.transfer-cb').addEventListener('change', function() {
        item.classList.toggle('selected', this.checked);
        updateAppTransferUI();
      });
    });
  }
}

function updateAppTransferUI() {
  const checked = document.querySelectorAll('#appTransferFromList .transfer-cb:checked');
  const totalContacts = [...checked].reduce((sum, cb) => sum + cb.dataset.ids.split(',').length, 0);
  $('appTransferSelectedCount').textContent = checked.length ? `已選 ${checked.length} 家公司（${totalContacts} 位聯絡人）` : '';
  $('appTransferConfirmBtn').disabled = !(checked.length > 0 && $('appTransferToOwner').value);
}

async function loadAppTransferFrom() {
  const owner = $('appTransferFromOwner').value;
  if (!owner) { $('appTransferFromList').innerHTML = '<div class="transfer-item-empty">請先選擇來源業務人員</div>'; return; }
  try {
    const r = await fetch(`${API}/contacts-by-owner?owner=${encodeURIComponent(owner)}`);
    const contacts = await r.json();
    const companyCount = new Set(contacts.map(c => c.company || '__')).size;
    $('appTransferFromCount').textContent = `${companyCount} 家公司・${contacts.length} 位聯絡人`;
    renderAppTransferList(contacts, 'appTransferFromList', true);
    updateAppTransferUI();
  } catch { showToast('載入失敗'); }
}

async function loadAppTransferTo() {
  const owner = $('appTransferToOwner').value;
  if (!owner) { $('appTransferToList').innerHTML = '<div class="transfer-item-empty">請先選擇目標業務人員</div>'; return; }
  try {
    const r = await fetch(`${API}/contacts-by-owner?owner=${encodeURIComponent(owner)}`);
    const contacts = await r.json();
    const companyCount = new Set(contacts.map(c => c.company || '__')).size;
    $('appTransferToCount').textContent = `${companyCount} 家公司・${contacts.length} 位聯絡人`;
    renderAppTransferList(contacts, 'appTransferToList', false);
  } catch {}
}

async function doAppTransfer() {
  const fromOwner = $('appTransferFromOwner').value;
  const toOwner   = $('appTransferToOwner').value;
  const checked   = [...document.querySelectorAll('#appTransferFromList .transfer-cb:checked')];
  if (!fromOwner || !toOwner || !checked.length) return;

  // 展開所有選中公司的聯絡人 ID
  const contactIds = checked.flatMap(cb => cb.dataset.ids.split(',').filter(Boolean));
  $('appTransferConfirmBtn').disabled = true;
  $('appTransferConfirmBtn').textContent = '移轉中...';

  try {
    const r = await fetch(`${API}/transfer-contacts`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fromOwner, toOwner, contactIds })
    });
    const result = await r.json();
    const box = $('appTransferResultBox');
    box.style.display = 'block';
    if (result.success) {
      box.className = 'transfer-result success';
      box.innerHTML = `✅ 移轉完成！共移轉 <strong>${result.contactCount}</strong> 位客戶，
        拜訪記錄 <strong>${result.visitCount}</strong> 筆、
        商機 <strong>${result.oppCount}</strong> 筆、
        應收帳款 <strong>${result.recvCount}</strong> 筆`;
      showToast(`✅ 成功移轉 ${result.contactCount} 位客戶`);
      await loadAppTransferFrom();
      await loadAppTransferTo();
    } else {
      box.className = 'transfer-result error';
      box.innerHTML = `❌ 移轉失敗：${escapeHtml(result.error || '未知錯誤')}`;
    }
  } catch { showToast('移轉失敗，請重試'); }
  finally {
    $('appTransferConfirmBtn').disabled = false;
    $('appTransferConfirmBtn').textContent = '確認移轉';
    updateAppTransferUI();
  }
}

// ── AI 名片拍照辨識（業務用）─────────────────────────────
(function setupOcrCard() {
  const btn    = $('ocrCardBtn');
  const input  = $('ocrCardInput');
  const status = $('ocrCardStatus');
  if (!btn || !input) return;

  // 點按鈕 → 開啟相機或檔案選取
  btn.addEventListener('click', () => input.click());

  input.addEventListener('change', async () => {
    const file = input.files[0];
    if (!file) return;
    input.value = '';   // 允許重複選同一張

    // 前端壓縮到 1200px
    const compressed = await compressCardImage(file, 1200);

    btn.disabled = true;
    status.textContent = '🤖 AI 辨識中，請稍候…';

    try {
      const fd = new FormData();
      fd.append('card', compressed, 'card.jpg');
      const r = await fetch(`${API}/ai/ocr-card`, { method: 'POST', body: fd });
      const data = await r.json();

      if (!r.ok) {
        status.textContent = '❌ ' + (data.error || 'AI 辨識失敗，請重試');
        return;
      }

      const c = data.contact || {};
      // 填入表單欄位
      const fill = (id, val) => { const el = $(id); if (el && val) el.value = val; };
      fill('name',    c.name);
      fill('nameEn',  c.nameEn);
      fill('company', c.company);
      fill('title',   c.title);
      fill('phone',   c.phone);
      fill('mobile',  c.mobile);
      fill('ext',     c.ext);
      fill('email',   c.email);
      fill('address', c.address);
      fill('website', c.website);
      fill('taxId',   c.taxId);

      // 若有統編，觸發自動查詢
      if (c.taxId && c.taxId.length === 8) {
        $('taxId').dispatchEvent(new Event('input', { bubbles: true }));
      }
      // 若有產業，設定
      if (c.industry) {
        const indEl = $('industry');
        if (indEl) { indEl.value = c.industry; indEl.dataset.manual = 'true'; }
      }

      status.textContent = '✅ 辨識完成，請確認並補充資料後儲存';
      status.style.color = '#1e8e3e';
    } catch {
      status.textContent = '❌ 網路錯誤，請稍後再試';
    } finally {
      btn.disabled = false;
    }
  });

  // 壓縮圖片（Canvas resize）
  function compressCardImage(file, maxPx) {
    return new Promise(resolve => {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => {
        URL.revokeObjectURL(url);
        let { width: w, height: h } = img;
        if (w > maxPx || h > maxPx) {
          if (w > h) { h = Math.round(h * maxPx / w); w = maxPx; }
          else       { w = Math.round(w * maxPx / h); h = maxPx; }
        }
        const canvas = document.createElement('canvas');
        canvas.width = w; canvas.height = h;
        canvas.getContext('2d').drawImage(img, 0, 0, w, h);
        canvas.toBlob(blob => resolve(new File([blob], 'card.jpg', { type: 'image/jpeg' })), 'image/jpeg', 0.88);
      };
      img.onerror = () => resolve(file);
      img.src = url;
    });
  }
})();

// ── AI 功能 ───────────────────────────────────────────────

// 拜訪記錄 AI 建議
const _visitAiBtn = $v('visitAiBtn');
if (_visitAiBtn) {
  _visitAiBtn.addEventListener('click', async () => {
    const topic       = $v('visitTopic').value.trim();
    const content     = $v('visitContent').value.trim();
    const visitType   = $v('visitType').value;
    const contactId   = $v('visitContactId').value;
    const contact     = allContacts.find(c => c.id === contactId);
    const contactName = contact?.name || '';
    const company     = contact?.company || '';

    if (!content && !topic) { showToast('請先填寫拜訪主題或內容再使用 AI 建議'); return; }

    _visitAiBtn.disabled = true;
    _visitAiBtn.textContent = '✨ 分析中…';
    $v('visitAiTakeaways').style.display = 'none';

    try {
      const r = await fetch(`${API}/ai/visit-suggest`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ topic, content, visitType, contactName, company })
      });
      const data = await r.json();
      if (!r.ok) {
        if (data.error === 'AI_NOT_CONFIGURED') showToast('⚠️ 未設定 GEMINI_API_KEY，請聯絡管理員');
        else showToast('AI 分析失敗：' + (data.error || '請重試'));
        return;
      }
      // 填入下一步建議（若空白則自動填入）
      if (data.nextAction && !$v('visitNextAction').value.trim()) {
        $v('visitNextAction').value = data.nextAction;
      }
      // 顯示關鍵重點
      if (data.keyTakeaways && data.keyTakeaways.length) {
        const taDiv = $v('visitAiTakeaways');
        taDiv.innerHTML = '<strong>🤖 AI 關鍵重點：</strong><ul style="margin:4px 0 0 16px;padding:0">' +
          data.keyTakeaways.map(t => `<li>${escapeHtml(t)}</li>`).join('') + '</ul>';
        if (data.nextAction) {
          taDiv.innerHTML += `<div style="margin-top:6px"><strong>💡 建議下一步：</strong>${escapeHtml(data.nextAction)}</div>`;
        }
        taDiv.style.display = '';
      }
    } catch { showToast('網路錯誤，請稍後再試'); }
    finally {
      _visitAiBtn.disabled = false;
      _visitAiBtn.textContent = '✨ AI 建議';
    }
  });
}

// ── Feature 1b：生成跟進信件草稿 ────────────────────────
const _visitEmailBtn = $v('visitEmailBtn');
if (_visitEmailBtn) {
  _visitEmailBtn.addEventListener('click', async () => {
    const topic   = $v('visitTopic').value.trim();
    const content = $v('visitContent').value.trim();
    if (!content && !topic) { showToast('請先填寫拜訪主題或會談內容再生成信件'); return; }

    const contactId = $v('visitContactId').value;
    const contact   = allContacts.find(c => c.id === contactId);

    _visitEmailBtn.disabled = true;
    _visitEmailBtn.textContent = '✉️ 產生中…';
    $v('visitEmailDraft').style.display = 'none';

    try {
      const r = await fetch(`${API}/ai/follow-up-email`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contactName: contact?.name  || '',
          company:     contact?.company || '',
          title:       contact?.title  || '',
          visitType:   $v('visitType').value,
          topic,
          content,
          nextAction:  $v('visitNextAction').value.trim()
        })
      });
      const data = await r.json();
      if (!r.ok) {
        if (data.error === 'AI_NOT_CONFIGURED') showToast('⚠️ 未設定 GEMINI_API_KEY，請聯絡管理員');
        else showToast('AI 生成失敗：' + (data.error || '請重試'));
        return;
      }
      $v('visitEmailSubject').textContent = data.subject || '';
      $v('visitEmailBody').innerText      = data.body || '';
      $v('visitEmailDraft').style.display = '';
      // 捲動到草稿框
      $v('visitEmailDraft').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    } catch { showToast('網路錯誤，請稍後再試'); }
    finally {
      _visitEmailBtn.disabled = false;
      _visitEmailBtn.textContent = '✉️ 生成跟進信件';
    }
  });

  // 複製全文
  $v('visitEmailCopyBtn').addEventListener('click', () => {
    const subject = $v('visitEmailSubject').textContent.trim();
    const body    = $v('visitEmailBody').innerText.trim();
    const text    = `主旨：${subject}\n\n${body}`;
    navigator.clipboard.writeText(text)
      .then(() => showToast('已複製到剪貼簿', 'success'))
      .catch(() => showToast('複製失敗，請手動選取'));
  });

  // 開啟郵件程式（mailto）
  $v('visitEmailMailtoBtn').addEventListener('click', () => {
    const contactId = $v('visitContactId').value;
    const contact   = allContacts.find(c => c.id === contactId);
    const toEmail   = contact?.email || '';
    const subject   = encodeURIComponent($v('visitEmailSubject').textContent.trim());
    const body      = encodeURIComponent($v('visitEmailBody').innerText.trim());
    window.location.href = `mailto:${toEmail}?subject=${subject}&body=${body}`;
  });
}

// 商機贏率預測
const _oppWinRateBtn = $('oppWinRateBtn');
if (_oppWinRateBtn) {
  _oppWinRateBtn.addEventListener('click', async () => {
    const oppId = $('oppEditId').value;
    if (!oppId) return;

    _oppWinRateBtn.disabled = true;
    _oppWinRateBtn.textContent = '🤖 預測中…';

    try {
      const r = await fetch(`${API}/ai/opp-win-rate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ oppId })
      });
      const data = await r.json();
      if (!r.ok) {
        if (data.error === 'AI_NOT_CONFIGURED') showToast('⚠️ 未設定 GEMINI_API_KEY，請聯絡管理員');
        else showToast('AI 分析失敗：' + (data.error || '請重試'));
        return;
      }
      // 顯示贏率面板
      const panel = $('oppWinRatePanel');
      if (panel) {
        $('oppWinRateValue').textContent = (data.winRate ?? '--') + '%';
        $('oppWinRateReasoning').textContent = data.reasoning || '';

        const f = data.factors || {};
        const factorsEl = $('oppWinRateFactors');
        if (factorsEl) {
          factorsEl.textContent = [
            f.stage    ? `階段：${f.stage}`       : '',
            f.activity ? `活躍度：${f.activity}` : '',
            f.timeline ? `時程：${f.timeline}`   : '',
            f.amount   ? `金額：${f.amount}`     : ''
          ].filter(Boolean).join('　');
        }
        panel.style.display = '';
        // 同步更新記憶體中的商機（快取）
        const oInMem = allOpportunities.find(x => x.id === oppId);
        if (oInMem) {
          oInMem.aiWinRate   = data.winRate;
          oInMem.aiWinRateAt = new Date().toISOString();
        }
      }
    } catch { showToast('網路錯誤，請稍後再試'); }
    finally {
      _oppWinRateBtn.disabled = false;
      _oppWinRateBtn.textContent = '🤖 預測贏率';
    }
  });
}

// 聯絡人 AI 摘要 —— tab 切換顯示快取、按鈕重新生成
(function setupContactAiTab() {
  // tab 切換時載入快取摘要
  const tabBtn = document.querySelector('[data-tab="tab-ai"]');
  if (!tabBtn) return;

  tabBtn.addEventListener('click', () => {
    const contactId = $('contactId').value;
    if (!contactId) return;
    const c = allContacts.find(x => x.id === contactId);
    if (!c) return;
    _renderAiSummary(c);
  });

  // 生成按鈕
  const genBtn = $('aiSummaryBtn');
  if (genBtn) {
    genBtn.addEventListener('click', async () => {
      const contactId = $('contactId').value;
      if (!contactId) return;

      genBtn.disabled = true;
      genBtn.textContent = '🤖 分析中…';
      $('aiSummaryStatus').textContent = '正在呼叫 AI，請稍候（約 5–15 秒）…';

      try {
        const r = await fetch(`${API}/ai/contact-summary`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ contactId })
        });
        const data = await r.json();
        if (!r.ok) {
          if (data.error === 'AI_NOT_CONFIGURED') {
            $('aiSummaryStatus').textContent = '⚠️ 未設定 GEMINI_API_KEY，請聯絡管理員';
          } else {
            $('aiSummaryStatus').textContent = '❌ ' + (data.error || 'AI 分析失敗，請重試');
          }
          return;
        }
        // 更新記憶體快取
        const c = allContacts.find(x => x.id === contactId);
        if (c) {
          c.aiSummary       = data.summary;
          c.aiSummaryHealth = data.health;
          c.aiSummaryAt     = new Date().toISOString();
          _renderAiSummary(c);
        }
        $('aiSummaryStatus').textContent = '';
      } catch {
        $('aiSummaryStatus').textContent = '❌ 網路錯誤，請稍後再試';
      } finally {
        genBtn.disabled = false;
        genBtn.textContent = '🤖 生成客戶輪廓摘要';
      }
    });
  }
})();

function _renderAiSummary(c) {
  const healthEl = $('aiSummaryHealth');
  const textEl   = $('aiSummaryText');
  const metaEl   = $('aiSummaryMeta');

  if (c.aiSummary) {
    textEl.textContent = c.aiSummary;
    if (metaEl && c.aiSummaryAt) {
      const d = new Date(c.aiSummaryAt);
      metaEl.textContent = `上次分析：${d.toLocaleDateString('zh-TW')} ${d.toLocaleTimeString('zh-TW', { hour: '2-digit', minute: '2-digit' })}`;
    }
    if (healthEl) {
      const hMap = { '良好': { bg: '#e8f5e9', color: '#1e8e3e', text: '✅ 關係良好' },
                     '普通': { bg: '#fff8e1', color: '#f9a825', text: '⚠️ 關係普通' },
                     '需關注': { bg: '#fce4ec', color: '#c62828', text: '🔴 需要關注' } };
      const h = hMap[c.aiSummaryHealth] || { bg: '#f5f5f5', color: '#666', text: c.aiSummaryHealth || '' };
      healthEl.style.background = h.bg;
      healthEl.style.color      = h.color;
      healthEl.textContent      = h.text;
      healthEl.style.display    = h.text ? 'inline-block' : 'none';
    }
  } else {
    textEl.textContent = '尚未生成摘要，點擊下方按鈕開始分析';
    if (healthEl) { healthEl.textContent = ''; healthEl.style.display = 'none'; }
    if (metaEl)   metaEl.textContent = '';
  }
}

// openModal 時控制 AI tab 顯示（僅編輯現有聯絡人時顯示）
// 利用 MutationObserver 監聽 modalOverlay class 變化
(function patchOpenModalForAiTab() {
  const tabAiBtn = document.getElementById('tabAiBtn');
  if (!tabAiBtn) return;
  const mo = new MutationObserver(() => {
    const isOpen = document.getElementById('modalOverlay').classList.contains('open');
    if (isOpen) {
      const hasId = !!document.getElementById('contactId').value;
      tabAiBtn.style.display = hasId ? '' : 'none';
      if (!hasId) {
        // 若當前在 tab-ai，切回 tab-card
        if (!document.getElementById('tab-ai').classList.contains('hidden')) {
          document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
          document.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
          document.querySelector('[data-tab="tab-card"]').classList.add('active');
          document.getElementById('tab-card').classList.remove('hidden');
        }
      }
    }
  });
  mo.observe(document.getElementById('modalOverlay'), { attributes: true, attributeFilter: ['class'] });
})();

// ════════════════════════════════════════════════════════
//  管理儀表板（Executive Dashboard）
// ════════════════════════════════════════════════════════
let execCurrentTab = 'conversion';
let execYear = new Date().getFullYear();
let execOwnerFilter = '';
let execData = { conversion: null, trend: null, product: null };

async function loadExecDash() {
  // 初始化年份選單
  const yrSel = $('execYearSel');
  if (!yrSel.options.length) {
    const cy = new Date().getFullYear();
    for (let y = cy + 1; y >= cy - 3; y--) {
      const o = document.createElement('option');
      o.value = y; o.textContent = y + ' 年';
      if (y === cy) o.selected = true;
      yrSel.appendChild(o);
    }
    yrSel.addEventListener('change', () => { execYear = parseInt(yrSel.value); loadExecDash(); });

    // Tab 切換
    document.querySelectorAll('.exec-tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        document.querySelectorAll('.exec-tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        execCurrentTab = btn.dataset.tab;
        ['conversion','trend','product'].forEach(t => {
          const el = $('execTab' + t.charAt(0).toUpperCase() + t.slice(1));
          if (el) el.style.display = t === execCurrentTab ? '' : 'none';
        });
        renderExecCurrentTab();
      });
    });
  }

  // 拉取資料（三個 API 並行）
  const qs = `year=${execYear}${execOwnerFilter ? '&owner=' + execOwnerFilter : ''}`;
  const trendQs = execOwnerFilter ? `?owner=${execOwnerFilter}` : '';

  const [conv, trend, prod] = await Promise.all([
    fetch(`${API}/exec/conversion?${qs}`).then(r => r.json()).catch(() => null),
    fetch(`${API}/exec/trend${trendQs}`).then(r => r.json()).catch(() => []),
    fetch(`${API}/exec/product-analysis?${qs}`).then(r => r.json()).catch(() => []),
  ]);
  execData = { conversion: conv, trend: trend || [], product: prod || [] };

  // 業務篩選器（主管才顯示）
  const ownerWrap = $('execOwnerWrap');
  if (conv && conv.ownerOptions && conv.ownerOptions.length > 0) {
    ownerWrap.style.display = 'flex';
    const sel = $('execOwnerSel');
    const curVal = sel.value;
    sel.innerHTML = '<option value="">全部業務</option>';
    conv.ownerOptions.forEach(u => {
      const o = document.createElement('option');
      o.value = u.username; o.textContent = u.displayName;
      if (u.username === curVal) o.selected = true;
      sel.appendChild(o);
    });
    if (!sel._bound) {
      sel._bound = true;
      sel.addEventListener('change', () => { execOwnerFilter = sel.value; loadExecDash(); });
    }
  } else {
    ownerWrap.style.display = 'none';
  }

  renderExecCurrentTab();
}

function renderExecCurrentTab() {
  if (execCurrentTab === 'conversion') renderExecConversion(execData.conversion);
  else if (execCurrentTab === 'trend')  renderExecTrend(execData.trend);
  else if (execCurrentTab === 'product') renderExecProduct(execData.product);
}

// ── 轉換率漏斗 ──────────────────────────────────────────
function renderExecConversion(d) {
  const el = $('execTabConversion');
  if (!d) { el.innerHTML = '<div class="exec-empty">載入中…</div>'; return; }

  const STAGE_INFO = {
    C: { label: 'C｜Pipeline', color: '#1a73e8', bg: '#e8f0fe' },
    B: { label: 'B｜Upside',   color: '#e37400', bg: '#fef3e2' },
    A: { label: 'A｜Commit',   color: '#c5221f', bg: '#fce8e6' },
    Won: { label: '🏆 Won',    color: '#1e8e3e', bg: '#e6f4ea' },
  };
  const stages = ['C','B','A','Won'];

  // 摘要卡
  const wrCard = (label, val, sub) => `
    <div class="exec-kpi-card">
      <div class="exec-kpi-label">${label}</div>
      <div class="exec-kpi-val">${val ?? '—'}</div>
      ${sub ? `<div class="exec-kpi-sub">${sub}</div>` : ''}
    </div>`;

  const winRateColor = d.winRate === null ? '#888' : d.winRate >= 60 ? '#1e8e3e' : d.winRate >= 40 ? '#e37400' : '#c5221f';
  const summaryHtml = `
    <div class="exec-kpi-row">
      ${wrCard('整體 Win Rate', d.winRate !== null ? `<span style="color:${winRateColor}">${d.winRate}%</span>` : '—',
        `成交 ${d.totalWon} / 已關閉 ${d.totalClosed}`)}
      ${wrCard('平均成交週期', d.avgCycleDays !== null ? `${d.avgCycleDays} 天` : '—', '從建立到成交')}
      ${wrCard('本年成交件數', d.totalWon, `流失 ${d.totalLost} 件`)}
    </div>`;

  // 漏斗視覺（每個 stage box + 轉換箭頭）
  let funnelHtml = '<div class="exec-funnel">';
  stages.forEach((s, i) => {
    const info = STAGE_INFO[s];
    const count = d.stageCounts[s] || 0;
    funnelHtml += `
      <div class="exec-funnel-stage">
        <div class="exec-funnel-box" style="border-color:${info.color};background:${info.bg}">
          <div class="exec-funnel-box-label" style="color:${info.color}">${info.label}</div>
          <div class="exec-funnel-box-count" style="color:${info.color}">${count}</div>
          <div class="exec-funnel-box-unit">件在手</div>
        </div>
      </div>`;
    if (i < stages.length - 1) {
      const trans = d.stages[i];
      const rateStr = trans && trans.rate !== null ? `${trans.rate}%` : '—';
      const daysStr = trans && trans.avgDays !== null ? `avg ${trans.avgDays}d` : '';
      funnelHtml += `
        <div class="exec-funnel-arrow">
          <div class="exec-funnel-rate">${rateStr}</div>
          <div class="exec-funnel-arrow-line">→</div>
          <div class="exec-funnel-days">${daysStr}</div>
        </div>`;
    }
  });
  funnelHtml += '</div>';

  // 轉換率詳細表格
  const tableRows = (d.stages || []).map(s => `
    <tr>
      <td>${s.from} → ${s.to}</td>
      <td style="text-align:right">${s.total || 0}</td>
      <td style="text-align:right">${s.count || 0}</td>
      <td style="text-align:right">
        <span class="exec-rate-badge" style="background:${s.rate >= 60 ? '#e6f4ea' : s.rate >= 40 ? '#fef3e2' : '#fce8e6'};color:${s.rate >= 60 ? '#1e8e3e' : s.rate >= 40 ? '#e37400' : '#c5221f'}">
          ${s.rate !== null ? s.rate + '%' : '—'}
        </span>
      </td>
      <td style="text-align:right">${s.avgDays !== null ? s.avgDays + ' 天' : '—'}</td>
    </tr>`).join('');

  el.innerHTML = summaryHtml + funnelHtml + `
    <div class="exec-section-title">📋 階段轉換明細</div>
    <div style="overflow-x:auto;background:#fff;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-top:8px">
      <table style="width:100%;border-collapse:collapse;font-size:13px">
        <thead>
          <tr style="background:#fafafa;border-bottom:2px solid #eee">
            <th style="padding:10px 16px;text-align:left;font-weight:600">轉換路徑</th>
            <th style="padding:10px 16px;text-align:right;font-weight:600">進入數</th>
            <th style="padding:10px 16px;text-align:right;font-weight:600">晉升數</th>
            <th style="padding:10px 16px;text-align:right;font-weight:600">轉換率</th>
            <th style="padding:10px 16px;text-align:right;font-weight:600">平均停留</th>
          </tr>
        </thead>
        <tbody>${tableRows}</tbody>
      </table>
    </div>`;
}

// ── 月度業績趨勢 ─────────────────────────────────────────
function renderExecTrend(data) {
  const el = $('execTabTrend');
  if (!data || !data.length) { el.innerHTML = '<div class="exec-empty">暫無資料</div>'; return; }

  const cy = new Date().getFullYear();
  const maxAmt = Math.max(...data.map(d => d.amount), 1);
  const W = 700, H = 220, PAD_L = 56, PAD_B = 38, PAD_T = 20, PAD_R = 16;
  const chartW = W - PAD_L - PAD_R;
  const chartH = H - PAD_B - PAD_T;
  const barW = Math.max(4, Math.floor(chartW / data.length) - 3);

  // Y axis gridlines
  const yTicks = 4;
  let gridLines = '', yLabels = '';
  for (let i = 0; i <= yTicks; i++) {
    const y = PAD_T + chartH - (i / yTicks) * chartH;
    const val = Math.round(maxAmt * i / yTicks);
    gridLines += `<line x1="${PAD_L}" y1="${y}" x2="${W - PAD_R}" y2="${y}" stroke="#eee" stroke-width="1"/>`;
    yLabels += `<text x="${PAD_L - 6}" y="${y + 4}" font-size="10" fill="#aaa" text-anchor="end">${val}</text>`;
  }

  // Bars
  let bars = '', xLabels = '', tooltipData = [];
  data.forEach((d, i) => {
    const x = PAD_L + i * (chartW / data.length) + (chartW / data.length - barW) / 2;
    const barH = d.amount > 0 ? Math.max(2, (d.amount / maxAmt) * chartH) : 0;
    const y = PAD_T + chartH - barH;
    const isThisYear = d.month && d.month.startsWith(String(cy));
    const fill = isThisYear ? '#1a73e8' : '#c5d9f5';
    bars += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${barW}" height="${barH.toFixed(1)}"
      rx="2" fill="${fill}" data-i="${i}" class="exec-bar">
      <title>${d.month}：${d.amount.toLocaleString()} 萬（${d.count} 件）</title>
    </rect>`;
    // X axis: show label every 3 months
    if (i % 3 === 0) {
      const lx = x + barW / 2;
      xLabels += `<text x="${lx.toFixed(1)}" y="${H - 6}" font-size="10" fill="#888" text-anchor="middle">${d.month.slice(0, 7)}</text>`;
    }
    tooltipData.push(d);
  });

  // Legend
  const legend = `
    <div class="exec-trend-legend">
      <span class="exec-legend-dot" style="background:#1a73e8"></span><span>${cy} 年</span>
      <span class="exec-legend-dot" style="background:#c5d9f5"></span><span>${cy - 1} 年</span>
    </div>`;

  // Total / summary
  const thisYearData = data.filter(d => d.month.startsWith(String(cy)));
  const totalAmt = thisYearData.reduce((s, d) => s + d.amount, 0);
  const totalCnt = thisYearData.reduce((s, d) => s + d.count, 0);
  const prevYearData = data.filter(d => d.month.startsWith(String(cy - 1)));
  const prevAmt = prevYearData.reduce((s, d) => s + d.amount, 0);
  const growth = prevAmt > 0 ? ((totalAmt - prevAmt) / prevAmt * 100).toFixed(1) : null;
  const growthStr = growth !== null
    ? `<span style="color:${parseFloat(growth) >= 0 ? '#1e8e3e' : '#c5221f'}">${parseFloat(growth) >= 0 ? '▲' : '▼'} ${Math.abs(growth)}%</span> vs 去年`
    : '';

  const kpis = `
    <div class="exec-kpi-row" style="margin-bottom:16px">
      <div class="exec-kpi-card">
        <div class="exec-kpi-label">${cy} 年 YTD 成交</div>
        <div class="exec-kpi-val">${totalAmt.toLocaleString()} 萬</div>
        <div class="exec-kpi-sub">${totalCnt} 件 ${growthStr}</div>
      </div>
      <div class="exec-kpi-card">
        <div class="exec-kpi-label">${cy - 1} 年全年成交</div>
        <div class="exec-kpi-val">${prevAmt.toLocaleString()} 萬</div>
        <div class="exec-kpi-sub">${prevYearData.reduce((s,d)=>s+d.count,0)} 件</div>
      </div>
    </div>`;

  el.innerHTML = kpis + legend + `
    <div class="exec-chart-wrap">
      <svg viewBox="0 0 ${W} ${H}" style="width:100%;max-width:${W}px;display:block">
        ${gridLines}${yLabels}
        <line x1="${PAD_L}" y1="${PAD_T}" x2="${PAD_L}" y2="${PAD_T + chartH}" stroke="#ddd" stroke-width="1"/>
        <line x1="${PAD_L}" y1="${PAD_T + chartH}" x2="${W - PAD_R}" y2="${PAD_T + chartH}" stroke="#ddd" stroke-width="1"/>
        ${bars}${xLabels}
      </svg>
    </div>`;
}

// ════════════════════════════════════════════════════════
//  行銷管理：行銷活動 & Lead
// ════════════════════════════════════════════════════════

const CAMPAIGN_TYPE_LABEL = { seminar:'研討會', webinar:'Webinar', digital:'數位廣告', exhibition:'展覽', other:'其他' };
const CAMPAIGN_STATUS_LABEL = { planned:'計畫中', active:'進行中', completed:'已完成', cancelled:'已取消' };
const CAMPAIGN_STATUS_COLOR = { planned:'#1a73e8', active:'#1e8e3e', completed:'#888', cancelled:'#c5221f' };
const LEAD_STATUS_LABEL = { new:'🆕 新進', assigned:'📋 已指派', contacted:'📞 已聯繫', converted:'✅ 已轉商機', disqualified:'❌ 不合格' };
const LEAD_STATUS_COLOR = { new:'#1a73e8', assigned:'#e37400', contacted:'#9c27b0', converted:'#1e8e3e', disqualified:'#888' };

let allCampaigns = [];
let allLeads     = [];
let salesUserList = []; // { username, displayName }

async function fetchSalesUsers() {
  if (salesUserList.length) return;
  try {
    const r = await fetch('/api/admin/users');
    if (!r.ok) return;
    const users = await r.json();
    salesUserList = users.filter(u => u.role === 'user' && u.active !== false)
      .map(u => ({ username: u.username, displayName: u.displayName || u.username }));
  } catch {}
}

// ── Lead Badge（業務看被分配的未處理 Lead）───────────────
async function refreshLeadsBadge() {
  try {
    const r = await fetch(`${API}/leads`);
    if (!r.ok) return;
    const leads = await r.json();
    const pending = leads.filter(l => l.status === 'assigned').length;
    const badge = $('leadsBadge');
    if (badge) {
      badge.textContent = pending;
      badge.style.display = pending > 0 ? '' : 'none';
    }
  } catch {}
}

// ════════════════════════════════════════════════════════
//  主管首頁（Manager Home）
// ════════════════════════════════════════════════════════
let mgrYear = new Date().getFullYear();
let mgrOwnerFilter = '';
const _mgrCharts = { gauge: null, aging: null, topCust: null };

async function loadManagerHome() {
  // 初始化年份選單與事件（只綁一次）
  const yrSel = $('mgrYearSel');
  if (!yrSel.options.length) {
    const cy = new Date().getFullYear();
    for (let y = cy + 1; y >= cy - 3; y--) {
      const o = document.createElement('option');
      o.value = y; o.textContent = y + ' 年';
      if (y === cy) o.selected = true;
      yrSel.appendChild(o);
    }
    yrSel.addEventListener('change', () => { mgrYear = parseInt(yrSel.value); loadManagerHome(); });
    $('mgrRefreshBtn').addEventListener('click', () => loadManagerHome());
  }

  const qs = `year=${mgrYear}${mgrOwnerFilter ? '&owner=' + mgrOwnerFilter : ''}`;
  let d;
  try {
    const r = await fetch(`${API}/manager-home?${qs}`);
    if (!r.ok) {
      const err = await r.json().catch(() => ({}));
      showToast(err.error || '載入主管首頁失敗', 'error');
      return;
    }
    d = await r.json();
  } catch (e) {
    showToast('網路錯誤：' + e.message, 'error');
    return;
  }

  // 業務篩選器
  const ownerWrap = $('mgrOwnerWrap');
  if (d.ownerOptions && d.ownerOptions.length > 1) {
    ownerWrap.style.display = 'flex';
    const sel = $('mgrOwnerSel');
    const cur = sel.value;
    sel.innerHTML = '<option value="">全公司</option>';
    d.ownerOptions.forEach(u => {
      const o = document.createElement('option');
      o.value = u.username; o.textContent = u.displayName;
      if (u.username === cur) o.selected = true;
      sel.appendChild(o);
    });
    if (!sel._bound) {
      sel._bound = true;
      sel.addEventListener('change', () => { mgrOwnerFilter = sel.value; loadManagerHome(); });
    }
  } else {
    ownerWrap.style.display = 'none';
  }

  renderMgrGauge(d.achievement);
  renderMgrCommit(d.thisMonthCommit);
  renderMgrAging(d.aging);
  renderMgrTopCust(d.topCustomers);
}

// ── 1. 達成儀表盤（doughnut） ─────────────────────────
function renderMgrGauge(a) {
  const pct = a.pct;
  const fmt = (n) => Number(n || 0).toLocaleString('zh-TW');
  $('mgrGaugePct').textContent   = pct === null ? '—' : pct + '%';
  $('mgrGaugeLabel').textContent = a.target > 0 ? '本年達成度' : '尚未設定目標';
  $('mgrGaugeDetail').innerHTML = a.target > 0
    ? `已達成 <b>${fmt(a.achieved)}</b> 萬 / 目標 <b>${fmt(a.target)}</b> 萬`
    : `本年成交 <b>${fmt(a.achieved)}</b> 萬（無目標可比對）`;

  // 顏色：>=90 綠 / 60-89 橘 / <60 紅
  const color = pct === null ? '#bbb' : pct >= 90 ? '#1e8e3e' : pct >= 60 ? '#e37400' : '#c5221f';
  const drawPct = pct === null ? 0 : Math.min(pct, 100);
  const ctx = $('mgrGaugeChart').getContext('2d');
  if (_mgrCharts.gauge) _mgrCharts.gauge.destroy();
  _mgrCharts.gauge = new Chart(ctx, {
    type: 'doughnut',
    data: {
      datasets: [{
        data: [drawPct, 100 - drawPct],
        backgroundColor: [color, '#eef0f3'],
        borderWidth: 0
      }]
    },
    options: {
      cutout: '78%',
      circumference: 270,
      rotation: -135,
      plugins: { legend: { display: false }, tooltip: { enabled: false } },
      responsive: true,
      maintainAspectRatio: false
    }
  });
}

// ── 2. 本月可望成交 ─────────────────────────────────────
function renderMgrCommit(c) {
  const fmt = (n) => Number(n || 0).toLocaleString('zh-TW');
  const block = (label, color, group, hint) => {
    const items = group.items.slice(0, 4).map(o => `
      <div class="mgr-commit-item">
        <span class="mgr-commit-stage" style="background:${color}">${o.stage}</span>
        <span class="mgr-commit-co">${escapeHtml(o.company || '—')}</span>
        <span class="mgr-commit-prod">${escapeHtml(o.product || '')}</span>
        <span class="mgr-commit-amt">${fmt(o.amount)}萬</span>
      </div>`).join('');
    const more = group.items.length > 4 ? `<div class="mgr-commit-more">…還有 ${group.items.length - 4} 案</div>` : '';
    return `
      <div class="mgr-commit-block" style="border-left-color:${color}">
        <div class="mgr-commit-head">
          <span class="mgr-commit-label">${label}</span>
          <span class="mgr-commit-stat">${group.count} 案 · <b>${fmt(group.amount)}萬</b></span>
        </div>
        <div class="mgr-commit-hint">${hint}</div>
        ${items || '<div class="mgr-commit-empty">本月無此類案件</div>'}
        ${more}
      </div>`;
  };
  $('mgrCommitContent').innerHTML =
    block('✅ 確定可成交', '#1e8e3e', c.confirmed, 'A 階段且推進中（建立 ≤60 天）') +
    block('⚠️ 風險案件',   '#c5221f', c.atRisk,    'A/B 階段但建立 >60 天，需介入') +
    block('❓ 變數較大',   '#e37400', c.uncertain, 'B 階段，需推進至 A 才有把握');
}

// ── 3. 商機 Aging（堆疊長條圖） ─────────────────────────
function renderMgrAging(a) {
  const bucketColors = {
    '0-7':   '#34a853',
    '8-30':  '#7cb342',
    '31-60': '#fbbc04',
    '61-90': '#fb8c00',
    '90+':   '#ea4335',
  };
  const datasets = a.buckets.map(b => ({
    label: b + ' 天',
    data: a.stages.map(s => a.data[s][b]),
    backgroundColor: bucketColors[b],
    borderWidth: 0,
    stack: 'stack-aging',
  }));
  const ctx = $('mgrAgingChart').getContext('2d');
  if (_mgrCharts.aging) _mgrCharts.aging.destroy();
  _mgrCharts.aging = new Chart(ctx, {
    type: 'bar',
    data: { labels: a.stages.map(s => s + ' 階段'), datasets },
    options: {
      responsive: true, maintainAspectRatio: false,
      indexAxis: 'y',
      scales: {
        x: { stacked: true, title: { display: true, text: '案件數' } },
        y: { stacked: true }
      },
      plugins: {
        legend: { position: 'bottom' },
        tooltip: { callbacks: { label: (c) => `${c.dataset.label}：${c.parsed.x} 件` } }
      }
    }
  });
  $('mgrAgingHint').innerHTML = a.stalledCount > 0
    ? `<span style="color:#c5221f;font-weight:600">⚠️ 共 ${a.stalledCount} 案在 C/B/A 階段停滯超過 60 天，建議優先檢視</span>`
    : `<span style="color:#1e8e3e">✅ 各階段案件健康，無顯著停滯</span>`;
}

// ── 4. 客戶 TOP 10（橫條圖） ───────────────────────────
function renderMgrTopCust(list) {
  const ctx = $('mgrTopCustChart').getContext('2d');
  if (_mgrCharts.topCust) _mgrCharts.topCust.destroy();
  if (!list || list.length === 0) {
    ctx.clearRect(0, 0, ctx.canvas.width, ctx.canvas.height);
    return;
  }
  const labels = list.map(c => c.company.length > 20 ? c.company.slice(0, 20) + '…' : c.company);
  _mgrCharts.topCust = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: '已成交（Won）', data: list.map(c => c.won),    backgroundColor: '#1e8e3e', stack: 's' },
        { label: '在手商機',       data: list.map(c => c.active), backgroundColor: '#1a73e8', stack: 's' },
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      indexAxis: 'y',
      scales: {
        x: { stacked: true, title: { display: true, text: '金額（萬）' } },
        y: { stacked: true }
      },
      plugins: {
        legend: { position: 'bottom' },
        tooltip: {
          callbacks: {
            label: (c) => `${c.dataset.label}：${Number(c.parsed.x).toLocaleString('zh-TW')} 萬`,
            footer: (items) => {
              const idx = items[0].dataIndex;
              return `共 ${list[idx].count} 件商機`;
            }
          }
        }
      }
    }
  });
}

// ════════════════════════════════════════════════════════
//  行銷活動
// ════════════════════════════════════════════════════════

async function loadCampaignsView() {
  try {
    const r = await fetch(`${API}/campaigns`);
    allCampaigns = r.ok ? await r.json() : [];
  } catch { allCampaigns = []; }
  renderCampaignCards();

  // 新增按鈕（行銷 / manager / admin 才顯示）
  const addBtn = $('addCampaignBtn');
  const role = userPermissions.role;
  if (addBtn) {
    addBtn.style.display = role === 'marketing' ? '' : 'none';
    addBtn.onclick = () => openCampaignModal(null);
  }
  // 篩選
  ['campaignTypeFilter','campaignStatusFilter'].forEach(id => {
    const el = $(id); if (el) { el.onchange = renderCampaignCards; }
  });
}

function renderCampaignCards() {
  const container = $('campaignCards');
  if (!container) return;
  const typeF   = $('campaignTypeFilter')?.value   || '';
  const statusF = $('campaignStatusFilter')?.value || '';
  let list = allCampaigns.filter(c =>
    (!typeF   || c.type   === typeF) &&
    (!statusF || c.status === statusF)
  );
  if (!list.length) {
    container.innerHTML = '<div style="grid-column:1/-1;padding:48px;text-align:center;color:#aaa">尚無行銷活動，點「新增活動」開始建立</div>';
    return;
  }
  const role = userPermissions.role;
  const canEdit = role === 'marketing';
  container.innerHTML = list.map(c => {
    const pct = c.targetCount > 0 ? Math.min(100, Math.round(c.leadCount / c.targetCount * 100)) : 0;
    const typeLabel   = CAMPAIGN_TYPE_LABEL[c.type]     || c.type    || '—';
    const statusLabel = CAMPAIGN_STATUS_LABEL[c.status] || c.status  || '—';
    const statusColor = CAMPAIGN_STATUS_COLOR[c.status] || '#888';
    const convRate    = c.leadCount > 0 ? Math.round(c.convertedCount / c.leadCount * 100) : 0;
    return `
      <div class="campaign-card" data-id="${c.id}">
        <div class="campaign-card-header">
          <span class="campaign-type-badge">${typeLabel}</span>
          <span class="campaign-status-badge" style="color:${statusColor};background:${statusColor}18">${statusLabel}</span>
        </div>
        <div class="campaign-card-name">${escapeHtml(c.name)}</div>
        <div class="campaign-card-date">${c.startDate || ''}${c.endDate && c.endDate !== c.startDate ? ' ～ ' + c.endDate : ''}</div>
        ${c.description ? `<div class="campaign-card-desc">${escapeHtml(c.description)}</div>` : ''}
        <div class="campaign-card-stats">
          <div class="campaign-stat">
            <div class="campaign-stat-val">${c.leadCount}</div>
            <div class="campaign-stat-lbl">Lead 數</div>
          </div>
          <div class="campaign-stat">
            <div class="campaign-stat-val" style="color:#1e8e3e">${c.convertedCount}</div>
            <div class="campaign-stat-lbl">已轉商機</div>
          </div>
          <div class="campaign-stat">
            <div class="campaign-stat-val">${c.leadCount > 0 ? convRate + '%' : '—'}</div>
            <div class="campaign-stat-lbl">轉換率</div>
          </div>
          ${c.budget ? `<div class="campaign-stat"><div class="campaign-stat-val">$${c.budget}萬</div><div class="campaign-stat-lbl">預算</div></div>` : ''}
        </div>
        ${c.targetCount > 0 ? `
          <div style="margin-top:10px">
            <div style="display:flex;justify-content:space-between;font-size:11px;color:#888;margin-bottom:4px">
              <span>Lead 進度</span><span>${c.leadCount} / ${c.targetCount}</span>
            </div>
            <div style="height:6px;background:#eee;border-radius:3px;overflow:hidden">
              <div style="width:${pct}%;height:100%;background:#1a73e8;border-radius:3px;transition:width .3s"></div>
            </div>
          </div>` : ''}
        <div class="campaign-card-actions">
          <button class="btn btn-sm btn-secondary" onclick="showSection('leads');filterLeadsByCampaign('${c.id}')">查看 Lead</button>
          ${canEdit ? `<button class="btn btn-sm btn-secondary" onclick="openCampaignModal('${c.id}')">編輯</button>` : ''}
        </div>
      </div>`;
  }).join('');
}

function filterLeadsByCampaign(campaignId) {
  const sel = $('leadCampaignFilter');
  if (sel) { sel.value = campaignId; renderLeadsTable(); }
}

function openCampaignModal(id) {
  const c = id ? allCampaigns.find(x => x.id === id) : null;
  $('campaignModalId').value   = c?.id   || '';
  $('campaignModalTitle').textContent = c ? '編輯行銷活動' : '新增行銷活動';
  $('campaignModalName').value   = c?.name        || '';
  $('campaignModalType').value   = c?.type        || 'seminar';
  $('campaignModalStatus').value = c?.status      || 'planned';
  $('campaignModalStart').value  = c?.startDate   || '';
  $('campaignModalEnd').value    = c?.endDate     || '';
  $('campaignModalTarget').value = c?.targetCount || '';
  $('campaignModalBudget').value = c?.budget      || '';
  $('campaignModalDesc').value   = c?.description || '';
  $('campaignModalOverlay').classList.add('open');
  setTimeout(() => $('campaignModalName').focus(), 60);
}
function closeCampaignModal() { $('campaignModalOverlay').classList.remove('open'); }

$('campaignModalSaveBtn').addEventListener('click', async () => {
  const name = $('campaignModalName').value.trim();
  if (!name) { showToast('請填入活動名稱'); return;  }
  const id = $('campaignModalId').value;
  const payload = {
    name,
    type:        $('campaignModalType').value,
    status:      $('campaignModalStatus').value,
    startDate:   $('campaignModalStart').value,
    endDate:     $('campaignModalEnd').value,
    targetCount: parseInt($('campaignModalTarget').value) || 0,
    budget:      parseFloat($('campaignModalBudget').value) || 0,
    description: $('campaignModalDesc').value.trim(),
  };
  try {
    const r = await fetch(`${API}/campaigns${id ? '/' + id : ''}`, {
      method: id ? 'PUT' : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!r.ok) { const e = await r.json().catch(()=>({})); showToast('❌ ' + (e.error||r.status)); return; }
    showToast(id ? '活動已更新' : '活動已建立');
    closeCampaignModal();
    await loadCampaignsView();
  } catch (e) { showToast('❌ 網路錯誤'); }
});

// ════════════════════════════════════════════════════════
//  Lead 管理
// ════════════════════════════════════════════════════════

async function loadLeadsView() {
  await fetchSalesUsers();
  try {
    const [lr, cr] = await Promise.all([
      fetch(`${API}/leads`),
      fetch(`${API}/campaigns`)
    ]);
    allLeads     = lr.ok ? await lr.json() : [];
    allCampaigns = cr.ok ? await cr.json() : allCampaigns;
  } catch { allLeads = []; }

  // 活動篩選下拉
  const campSel = $('leadCampaignFilter');
  if (campSel) {
    const curVal = campSel.value;
    campSel.innerHTML = '<option value="">全部活動</option>';
    allCampaigns.forEach(c => {
      const o = document.createElement('option');
      o.value = c.id; o.textContent = c.name;
      if (c.id === curVal) o.selected = true;
      campSel.appendChild(o);
    });
    campSel.onchange = renderLeadsTable;
  }

  // Lead Modal 活動下拉也更新
  const leadCampSel = $('leadModalCampaign');
  if (leadCampSel) {
    leadCampSel.innerHTML = '<option value="">-- 無活動 --</option>';
    allCampaigns.forEach(c => {
      const o = document.createElement('option');
      o.value = c.id; o.textContent = c.name;
      leadCampSel.appendChild(o);
    });
  }

  // 新增/搜尋事件
  const addBtn = $('addLeadBtn');
  const role = userPermissions.role;
  if (addBtn) {
    addBtn.style.display = role === 'marketing' ? '' : 'none';
    addBtn.onclick = () => openLeadModal(null);
  }
  const srch = $('leadSearch');
  if (srch) srch.oninput = renderLeadsTable;
  const statusF = $('leadStatusFilter');
  if (statusF) statusF.onchange = renderLeadsTable;

  renderLeadsTable();
}

function renderLeadsTable() {
  const role     = userPermissions.role;
  const canManage = ['manager1','manager2','admin'].includes(role);
  const srch     = ($('leadSearch')?.value || '').toLowerCase();
  const statusF  = $('leadStatusFilter')?.value || '';
  const campF    = $('leadCampaignFilter')?.value || '';

  let list = allLeads.filter(l =>
    (!statusF || l.status === statusF) &&
    (!campF   || l.campaignId === campF) &&
    (!srch    || (l.company||'').toLowerCase().includes(srch) || (l.contactName||'').toLowerCase().includes(srch))
  );

  const tbody = $('leadsTbody');
  if (!tbody) return;
  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:#aaa">暫無 Lead 資料</td></tr>';
    return;
  }

  tbody.innerHTML = list.map(l => {
    const statusLabel = LEAD_STATUS_LABEL[l.status] || l.status;
    const statusColor = LEAD_STATUS_COLOR[l.status] || '#888';
    const assignedName = l.assignedTo
      ? (salesUserList.find(u => u.username === l.assignedTo)?.displayName || l.assignedTo)
      : '—';
    const actions = canManage ? buildLeadActions(l) : '';
    return `<tr>
      <td style="font-weight:500">${escapeHtml(l.company || '—')}</td>
      <td>${escapeHtml(l.contactName || '—')}</td>
      <td style="color:#888;font-size:12px">${escapeHtml(l.title || '—')}</td>
      <td style="font-size:12px;color:#555">${escapeHtml(l.campaignName || '—')}</td>
      <td><span class="lead-status-badge" style="color:${statusColor};background:${statusColor}18">${statusLabel}</span></td>
      <td style="font-size:12px">${escapeHtml(assignedName)}</td>
      <td style="font-size:12px;color:#aaa">${(l.createdAt||'').slice(0,10)}</td>
      <td style="white-space:nowrap">${actions}</td>
    </tr>`;
  }).join('');
}

function buildLeadActions(l) {
  const btns = [];
  btns.push(`<button class="btn btn-xs btn-secondary" onclick="openLeadModal('${l.id}')">編輯</button>`);
  if (l.status !== 'converted' && l.status !== 'disqualified') {
    btns.push(`<button class="btn btn-xs btn-primary" onclick="openAssignLeadModal('${l.id}')">指派</button>`);
    btns.push(`<button class="btn btn-xs" style="background:#1e8e3e;color:#fff" onclick="openConvertLeadModal('${l.id}')">轉商機</button>`);
    btns.push(`<button class="btn btn-xs btn-secondary" onclick="disqualifyLead('${l.id}')">不合格</button>`);
  }
  return btns.join(' ');
}

function openLeadModal(id) {
  const l = id ? allLeads.find(x => x.id === id) : null;
  $('leadModalId').value         = l?.id          || '';
  $('leadModalHeading').textContent = l ? '編輯 Lead' : '新增 Lead';
  $('leadModalCompany').value    = l?.company     || '';
  $('leadModalContact').value    = l?.contactName || '';
  $('leadModalJobTitle').value   = l?.title       || '';
  $('leadModalPhone').value      = l?.phone       || '';
  $('leadModalEmail').value      = l?.email       || '';
  $('leadModalInterest').value   = l?.interest    || '';
  $('leadModalNote').value       = l?.note        || '';
  const campSel = $('leadModalCampaign');
  if (campSel) campSel.value = l?.campaignId || '';
  $('leadModalOverlay').classList.add('open');
  setTimeout(() => $('leadModalCompany').focus(), 60);
}
function closeLeadModal() { $('leadModalOverlay').classList.remove('open'); }

$('leadModalSaveBtn').addEventListener('click', async () => {
  const company = $('leadModalCompany').value.trim();
  if (!company) { showToast('請填入公司名稱'); return; }
  const id = $('leadModalId').value;
  const campSel = $('leadModalCampaign');
  const campId  = campSel?.value || '';
  const campName = campId ? (allCampaigns.find(c => c.id === campId)?.name || '') : '';
  const payload = {
    company,
    contactName:  $('leadModalContact').value.trim(),
    title:        ($('leadModalJobTitle')?.value || '').trim(),
    phone:        $('leadModalPhone').value.trim(),
    email:        $('leadModalEmail').value.trim(),
    interest:     $('leadModalInterest').value.trim(),
    note:         $('leadModalNote').value.trim(),
    campaignId:   campId,
    campaignName: campName,
  };
  try {
    const r = await fetch(`${API}/leads${id ? '/' + id : ''}`, {
      method: id ? 'PUT' : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!r.ok) { const e = await r.json().catch(()=>({})); showToast('❌ ' + (e.error||r.status)); return; }
    showToast(id ? 'Lead 已更新' : 'Lead 已建立');
    closeLeadModal();
    await loadLeadsView();
  } catch { showToast('❌ 網路錯誤'); }
});

// ── 指派 Lead ─────────────────────────────────────────
async function openAssignLeadModal(leadId) {
  await fetchSalesUsers();
  const l = allLeads.find(x => x.id === leadId);
  if (!l) return;
  $('assignLeadId').value = leadId;
  $('assignLeadName').textContent = `${l.company || ''}${l.contactName ? ' / ' + l.contactName : ''}`;
  const sel = $('assignSalesSel');
  sel.innerHTML = '<option value="">請選擇業務…</option>';
  salesUserList.forEach(u => {
    const o = document.createElement('option');
    o.value = u.username; o.textContent = u.displayName;
    if (u.username === l.assignedTo) o.selected = true;
    sel.appendChild(o);
  });
  $('assignLeadOverlay').classList.add('open');
}
function closeAssignLeadModal() { $('assignLeadOverlay').classList.remove('open'); }

$('assignLeadConfirmBtn').addEventListener('click', async () => {
  const leadId = $('assignLeadId').value;
  const assignedTo = $('assignSalesSel').value;
  if (!assignedTo) { showToast('請選擇業務'); return; }
  try {
    const r = await fetch(`${API}/leads/${leadId}/assign`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ assignedTo })
    });
    if (!r.ok) { const e = await r.json().catch(()=>({})); showToast('❌ ' + (e.error||r.status)); return; }
    showToast('✅ Lead 已指派');
    closeAssignLeadModal();
    await loadLeadsView();
  } catch { showToast('❌ 網路錯誤'); }
});

// ── 轉換 Lead → 商機 ──────────────────────────────────
async function openConvertLeadModal(leadId) {
  await fetchSalesUsers();
  const l = allLeads.find(x => x.id === leadId);
  if (!l) return;
  $('convertLeadId').value = leadId;
  $('convertLeadCompany').value  = l.company     || '';
  $('convertLeadContact').value  = l.contactName || '';
  $('convertOppName').value      = l.interest    || '';
  $('convertOppCategory').value  = '';
  $('convertOppStage').value     = 'C';
  const sel = $('convertSalesSel');
  sel.innerHTML = '<option value="">請選擇業務…</option>';
  salesUserList.forEach(u => {
    const o = document.createElement('option');
    o.value = u.username; o.textContent = u.displayName;
    if (u.username === l.assignedTo) o.selected = true;
    sel.appendChild(o);
  });
  $('convertLeadOverlay').classList.add('open');
  setTimeout(() => $('convertOppName').focus(), 60);
}
function closeConvertLeadModal() { $('convertLeadOverlay').classList.remove('open'); }

$('convertLeadConfirmBtn').addEventListener('click', async () => {
  const leadId     = $('convertLeadId').value;
  const oppName    = $('convertOppName').value.trim();
  const salesPerson = $('convertSalesSel').value;
  if (!oppName)      { showToast('請填入商機名稱'); return; }
  if (!salesPerson)  { showToast('請選擇負責業務'); return; }
  try {
    const r = await fetch(`${API}/leads/${leadId}/convert`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        oppName,
        category:    $('convertOppCategory').value.trim(),
        stage:       $('convertOppStage').value,
        salesPerson,
      })
    });
    if (!r.ok) { const e = await r.json().catch(()=>({})); showToast('❌ ' + (e.error||r.status)); return; }
    showToast('🎉 Lead 已成功轉換為商機！');
    closeConvertLeadModal();
    await loadLeadsView();
  } catch { showToast('❌ 網路錯誤'); }
});

// ── 不合格 ────────────────────────────────────────────
async function disqualifyLead(leadId) {
  const l = allLeads.find(x => x.id === leadId);
  if (!l) return;
  if (!confirm(`確認將「${l.company || l.contactName}」標記為不合格 Lead？`)) return;
  try {
    const r = await fetch(`${API}/leads/${leadId}/disqualify`, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ reason: '' })
    });
    if (!r.ok) { const e = await r.json().catch(()=>({})); showToast('❌ ' + (e.error||r.status)); return; }
    showToast('Lead 已標記為不合格');
    await loadLeadsView();
  } catch { showToast('❌ 網路錯誤'); }
}

// ── 產品 / BU 分析 ────────────────────────────────────────
let execProductSortKey = 'wonAmount';
let execProductSortAsc = false;

function renderExecProduct(data) {
  const el = $('execTabProduct');
  if (!data || !data.length) { el.innerHTML = '<div class="exec-empty">暫無資料</div>'; return; }

  const sorted = [...data].sort((a, b) => {
    const va = a[execProductSortKey] ?? -1;
    const vb = b[execProductSortKey] ?? -1;
    return execProductSortAsc ? va - vb : vb - va;
  });
  const maxWin = Math.max(...sorted.map(r => r.winRate || 0), 1);

  const thStyle = (key) => {
    const active = key === execProductSortKey;
    return `style="padding:10px 14px;text-align:${key==='category'?'left':'right'};font-weight:600;cursor:pointer;${active ? 'color:#1a73e8;' : ''};white-space:nowrap;user-select:none"
      data-sortkey="${key}"`;
  };

  const rows = sorted.map(r => {
    const winColor = r.winRate === null ? '#888' : r.winRate >= 60 ? '#1e8e3e' : r.winRate >= 40 ? '#e37400' : '#c5221f';
    const barPct = r.winRate !== null ? Math.round(r.winRate / 100 * 100) : 0;
    return `<tr class="exec-prod-row">
      <td style="padding:10px 14px;font-weight:500">${escapeHtml(r.category)}</td>
      <td style="padding:10px 14px;text-align:right">${r.count}</td>
      <td style="padding:10px 14px;text-align:right">${r.pipelineAmount.toLocaleString()}</td>
      <td style="padding:10px 14px;text-align:right;color:#1e8e3e;font-weight:600">${r.wonAmount.toLocaleString()}</td>
      <td style="padding:10px 14px;text-align:right">
        <div style="display:flex;align-items:center;gap:8px;justify-content:flex-end">
          <div style="width:64px;height:8px;background:#eee;border-radius:4px;overflow:hidden">
            <div style="width:${barPct}%;height:100%;background:${winColor};border-radius:4px"></div>
          </div>
          <span style="color:${winColor};font-weight:600;min-width:40px;text-align:right">${r.winRate !== null ? r.winRate + '%' : '—'}</span>
        </div>
      </td>
      <td style="padding:10px 14px;text-align:right">${r.avgGrossMargin !== null ? r.avgGrossMargin + '%' : '—'}</td>
    </tr>`;
  }).join('');

  const sortArrow = (key) => key === execProductSortKey ? (execProductSortAsc ? ' ▲' : ' ▼') : '';

  el.innerHTML = `
    <div style="overflow-x:auto;background:#fff;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.08)">
      <table style="width:100%;border-collapse:collapse;font-size:13px" id="execProdTable">
        <thead>
          <tr style="background:#fafafa;border-bottom:2px solid #eee">
            <th ${thStyle('category')}>產品線 / BU${sortArrow('category')}</th>
            <th ${thStyle('count')}>商機數${sortArrow('count')}</th>
            <th ${thStyle('pipelineAmount')}>在手金額(萬)${sortArrow('pipelineAmount')}</th>
            <th ${thStyle('wonAmount')}>成交金額(萬)${sortArrow('wonAmount')}</th>
            <th ${thStyle('winRate')}>Win Rate${sortArrow('winRate')}</th>
            <th ${thStyle('avgGrossMargin')}>平均毛利率${sortArrow('avgGrossMargin')}</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;

  // 排序事件
  el.querySelectorAll('th[data-sortkey]').forEach(th => {
    th.addEventListener('click', () => {
      const key = th.dataset.sortkey;
      if (execProductSortKey === key) execProductSortAsc = !execProductSortAsc;
      else { execProductSortKey = key; execProductSortAsc = false; }
      renderExecProduct(execData.product);
    });
  });
}
