const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();

pptx.layout = 'LAYOUT_WIDE'; // 16:9

// ── 色彩系統 ──────────────────────────────────────────────
const C = {
  bg:      '0d1117',
  card:    '161b22',
  card2:   '1c2333',
  cyan:    '00d4ff',
  cyanD:   '0099bb',
  purple:  '7c3aed',
  purpleL: 'a855f7',
  green:   '3fb950',
  orange:  'f97316',
  red:     'f85149',
  white:   'ffffff',
  gray:    '8b949e',
  grayL:   'c9d1d9',
  yellow:  'ffd60a',
};

// ── 共用 helper ───────────────────────────────────────────
function setBg(slide) {
  slide.background = { color: C.bg };
}

function addTitle(slide, text, y = 0.35, size = 28) {
  slide.addText(text, {
    x: 0.5, y, w: 12.33, h: 0.6,
    fontSize: size, bold: true, color: C.cyan,
    fontFace: 'Arial',
  });
}

function addSubtitle(slide, text, y = 0.9) {
  slide.addText(text, {
    x: 0.5, y, w: 12.33, h: 0.35,
    fontSize: 13, color: C.gray,
    fontFace: 'Arial',
  });
}

function addBody(slide, text, x, y, w, h, size = 13, color = C.grayL) {
  slide.addText(text, {
    x, y, w, h,
    fontSize: size, color, fontFace: 'Arial',
    valign: 'top', wrap: true,
  });
}

function addCard(slide, x, y, w, h, color = C.card2) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h,
    fill: { color },
    line: { color: C.cyan, width: 0.5, transparency: 70 },
    rectRadius: 0.08,
  });
}

function addBadge(slide, text, x, y, color = C.cyan) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w: 1.6, h: 0.3,
    fill: { color, transparency: 80 },
    line: { color, width: 1 },
    rectRadius: 0.15,
  });
  slide.addText(text, {
    x, y, w: 1.6, h: 0.3,
    fontSize: 10, bold: true, color,
    align: 'center', valign: 'middle', fontFace: 'Arial',
  });
}

function addScreenshotBox(slide, x, y, w, h, label) {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: '0a0f1a' },
    line: { color: C.cyan, width: 1.5, dashType: 'dash' },
  });
  slide.addText(`📷 截圖：${label}`, {
    x, y: y + h / 2 - 0.2, w, h: 0.4,
    fontSize: 11, color: C.gray,
    align: 'center', valign: 'middle', fontFace: 'Arial',
    italic: true,
  });
}

function addDivider(slide, y) {
  slide.addShape(pptx.ShapeType.line, {
    x: 0.5, y, w: 12.33, h: 0,
    line: { color: C.cyan, width: 0.5, transparency: 60 },
  });
}

function addIcon(slide, icon, x, y) {
  slide.addText(icon, {
    x, y, w: 0.5, h: 0.5,
    fontSize: 22, align: 'center', valign: 'middle',
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 1 — 封面
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);

  // 左側發光條
  s.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 0.08, h: 5.63,
    fill: { color: C.cyan },
    line: { color: C.cyan, width: 0 },
  });

  // 背景裝飾圓
  s.addShape(pptx.ShapeType.ellipse, {
    x: 8.5, y: -1, w: 5, h: 5,
    fill: { color: C.purple, transparency: 90 },
    line: { color: C.purple, width: 0 },
  });
  s.addShape(pptx.ShapeType.ellipse, {
    x: 9.5, y: 2, w: 3.5, h: 3.5,
    fill: { color: C.cyan, transparency: 92 },
    line: { color: C.cyan, width: 0 },
  });

  // Tag
  addBadge(s, '🚀 內部產品簡報', 0.6, 0.5, C.cyan);

  // 主標題
  s.addText('ITTS CRM', {
    x: 0.6, y: 1.1, w: 8, h: 1.1,
    fontSize: 64, bold: true, color: C.white,
    fontFace: 'Arial',
  });
  s.addText('智能業務管理平台', {
    x: 0.6, y: 2.1, w: 8, h: 0.7,
    fontSize: 32, bold: true, color: C.cyan,
    fontFace: 'Arial',
  });

  // 副標題
  s.addText('聯絡人管理 ｜ 商機看板 ｜ 拜訪記錄 ｜ AI 智能分析', {
    x: 0.6, y: 2.85, w: 9, h: 0.4,
    fontSize: 14, color: C.gray,
    fontFace: 'Arial',
  });

  addDivider(s, 3.4);

  // 底部 stats
  const stats = [
    { icon: '🤖', val: '4 項 AI', sub: '智能功能' },
    { icon: '☁️', val: '100%', sub: '雲端部署' },
    { icon: '📱', val: '手機友善', sub: '隨時存取' },
    { icon: '🔒', val: '多層', sub: '權限管理' },
  ];
  stats.forEach((st, i) => {
    const x = 0.6 + i * 3.1;
    s.addText(st.icon + ' ' + st.val, { x, y: 3.6, w: 2.8, h: 0.45, fontSize: 18, bold: true, color: C.white, fontFace: 'Arial' });
    s.addText(st.sub, { x, y: 4.0, w: 2.8, h: 0.3, fontSize: 12, color: C.gray, fontFace: 'Arial' });
  });

  // 日期
  s.addText('2026.04', { x: 10.5, y: 5.1, w: 2.3, h: 0.3, fontSize: 11, color: C.gray, align: 'right', fontFace: 'Arial' });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 2 — 業務痛點 → 解方
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '我們在解決什麼問題？');
  addDivider(s, 1.05);

  const pains = [
    { icon: '😩', title: '名片一堆，找不到人', color: C.red },
    { icon: '🤷', title: '商機進度不透明', color: C.orange },
    { icon: '📝', title: '拜訪記錄靠記憶', color: C.yellow },
    { icon: '📊', title: '業績分析曠日廢時', color: C.purple },
  ];
  const solutions = [
    { icon: '🗂️', title: '數位化名片管理', sub: '一秒找到任何聯絡人', color: C.cyan },
    { icon: '🏆', title: '視覺化商機看板', sub: 'Kanban 拖曳即更新進度', color: C.green },
    { icon: '📋', title: '結構化拜訪記錄', sub: '日期、主題、下一步全紀錄', color: C.cyan },
    { icon: '🤖', title: 'AI 自動分析', sub: '贏率預測、客戶摘要一鍵生成', color: C.purpleL },
  ];

  // 左側痛點
  s.addText('現況痛點', { x: 0.5, y: 1.15, w: 5.8, h: 0.35, fontSize: 13, bold: true, color: C.red, fontFace: 'Arial' });
  pains.forEach((p, i) => {
    addCard(s, 0.5, 1.55 + i * 0.88, 5.6, 0.75, '1a0a0a');
    s.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.55 + i * 0.88, w: 0.06, h: 0.75, fill: { color: p.color }, line: { color: p.color, width: 0 } });
    s.addText(p.icon + '  ' + p.title, { x: 0.7, y: 1.62 + i * 0.88, w: 5.2, h: 0.35, fontSize: 15, bold: true, color: C.white, fontFace: 'Arial' });
  });

  // 箭頭
  s.addText('➜', { x: 6.2, y: 2.7, w: 0.8, h: 0.5, fontSize: 28, bold: true, color: C.cyan, align: 'center', fontFace: 'Arial' });

  // 右側解方
  s.addText('ITTS CRM 解方', { x: 7.1, y: 1.15, w: 5.8, h: 0.35, fontSize: 13, bold: true, color: C.green, fontFace: 'Arial' });
  solutions.forEach((sol, i) => {
    addCard(s, 7.1, 1.55 + i * 0.88, 5.7, 0.75, '0a1a0a');
    s.addShape(pptx.ShapeType.rect, { x: 7.1, y: 1.55 + i * 0.88, w: 0.06, h: 0.75, fill: { color: sol.color }, line: { color: sol.color, width: 0 } });
    s.addText(sol.icon + '  ' + sol.title, { x: 7.3, y: 1.62 + i * 0.88, w: 5.2, h: 0.35, fontSize: 15, bold: true, color: C.white, fontFace: 'Arial' });
    s.addText(sol.sub, { x: 7.3, y: 1.95 + i * 0.88, w: 5.2, h: 0.25, fontSize: 11, color: C.gray, fontFace: 'Arial' });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 3 — 系統功能總覽
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '系統功能全覽');
  addDivider(s, 1.05);

  const features = [
    { icon: '🗂️', title: '名片管理', items: ['數位化聯絡人', 'AI 拍照辨識', '多欄位搜尋'], color: C.cyan },
    { icon: '📊', title: '商機看板', items: ['Kanban 拖曳', 'Won 成交慶祝', '贏率預測 AI'], color: C.green },
    { icon: '📋', title: '拜訪記錄', items: ['結構化記錄', 'AI 建議下一步', '業務日報'], color: C.purpleL },
    { icon: '🎯', title: '目標管理', items: ['月度業績設定', '達成率追蹤', '多維度報表'], color: C.orange },
    { icon: '👥', title: '管理員功能', items: ['用戶管理', '資料匯出入', '帳號轉移'], color: C.yellow },
    { icon: '🤖', title: 'AI 智能分析', items: ['OCR 辨識', '客戶摘要', '商機預測'], color: C.purple },
  ];

  features.forEach((f, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.4 + col * 4.27;
    const y = 1.2 + row * 2.05;

    addCard(s, x, y, 4.05, 1.85);
    // 頂部色條
    s.addShape(pptx.ShapeType.rect, { x, y, w: 4.05, h: 0.06, fill: { color: f.color }, line: { color: f.color, width: 0 } });

    s.addText(f.icon + ' ' + f.title, { x: x + 0.15, y: y + 0.12, w: 3.7, h: 0.42, fontSize: 17, bold: true, color: C.white, fontFace: 'Arial' });
    f.items.forEach((item, j) => {
      s.addText('• ' + item, { x: x + 0.2, y: y + 0.6 + j * 0.36, w: 3.6, h: 0.33, fontSize: 12, color: C.grayL, fontFace: 'Arial' });
    });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 4 — 名片管理
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '🗂️  名片 & 聯絡人管理');
  addSubtitle(s, '所有業務往來對象，集中數位管理');
  addDivider(s, 1.22);

  // 左側說明
  const points = [
    { icon: '🔍', title: '多欄位快速搜尋', desc: '公司、姓名、職稱、產業即時篩選' },
    { icon: '🏷️', title: '智能分類管理', desc: '主要聯繫窗口標記、離職狀態追蹤' },
    { icon: '📱', title: '個人資訊記錄', desc: '飲食偏好、興趣、生日、備忘一站管理' },
    { icon: '🔗', title: '關聯商機與拜訪', desc: '一鍵查看與該聯絡人的所有往來記錄' },
  ];
  points.forEach((p, i) => {
    addCard(s, 0.4, 1.35 + i * 0.98, 5.5, 0.83);
    s.addText(p.icon, { x: 0.55, y: 1.42 + i * 0.98, w: 0.5, h: 0.5, fontSize: 20, align: 'center', fontFace: 'Arial' });
    s.addText(p.title, { x: 1.1, y: 1.44 + i * 0.98, w: 4.6, h: 0.32, fontSize: 14, bold: true, color: C.white, fontFace: 'Arial' });
    s.addText(p.desc, { x: 1.1, y: 1.74 + i * 0.98, w: 4.6, h: 0.28, fontSize: 11, color: C.gray, fontFace: 'Arial' });
  });

  // 右側截圖區
  addScreenshotBox(s, 6.1, 1.35, 6.6, 3.9, '聯絡人列表頁面');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 5 — 商機看板
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '📊  商機看板（Kanban）');
  addSubtitle(s, '視覺化掌握每筆商機推進狀態');
  addDivider(s, 1.22);

  // Kanban 示意
  const stages = [
    { key: 'A', label: '探索接觸', color: C.gray },
    { key: 'B', label: '需求確認', color: C.cyan },
    { key: 'C', label: '方案報價', color: C.purple },
    { key: 'D', label: '議約收尾', color: C.orange },
    { key: 'Won', label: '🏆 成交', color: C.green },
  ];

  stages.forEach((st, i) => {
    const x = 0.35 + i * 2.56;
    // 欄位標題
    s.addShape(pptx.ShapeType.roundRect, {
      x, y: 1.3, w: 2.38, h: 0.38,
      fill: { color: st.color, transparency: 75 },
      line: { color: st.color, width: 1 },
      rectRadius: 0.05,
    });
    s.addText(st.label, { x, y: 1.3, w: 2.38, h: 0.38, fontSize: 12, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: 'Arial' });

    // 卡片示意
    const cardCount = [2, 3, 1, 2, 1][i];
    for (let c = 0; c < cardCount; c++) {
      addCard(s, x + 0.06, 1.78 + c * 0.82, 2.26, 0.7);
      s.addShape(pptx.ShapeType.rect, { x: x + 0.06, y: 1.78 + c * 0.82, w: 0.05, h: 0.7, fill: { color: st.color }, line: { color: st.color, width: 0 } });
      s.addText(['鼎新 ERP 導入', '正航系統升級', 'SAP 評估', '雲端移轉', '年約維護', '新客開發', 'ERP 整合', 'CRM 建置', '資安服務'][i * 2 + c] || '商機案件', {
        x: x + 0.15, y: 1.82 + c * 0.82, w: 2.1, h: 0.28, fontSize: 10, bold: true, color: C.white, fontFace: 'Arial',
      });
      s.addText(['$280萬', '$150萬', '$320萬', '$90萬', '$180萬', '$240萬', '$110萬', '$450萬', '$200萬'][i * 2 + c] || '', {
        x: x + 0.15, y: 2.08 + c * 0.82, w: 2.1, h: 0.22, fontSize: 9, color: st.color, fontFace: 'Arial',
      });
    }
  });

  // 底部說明
  addCard(s, 0.35, 5.0, 12.63, 0.4, C.card);
  s.addText('💡  拖曳卡片即可更新商機階段｜成交後自動觸發慶祝動畫｜支援同時管理多業務人員商機', {
    x: 0.35, y: 5.0, w: 12.63, h: 0.4, fontSize: 11, color: C.cyan, align: 'center', valign: 'middle', fontFace: 'Arial',
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 6 — Won 成交慶祝
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '🏆  Won — 成交里程碑慶祝');
  addSubtitle(s, '當商機進入「成交」階段，系統自動觸發視覺慶祝效果');
  addDivider(s, 1.22);

  // 左側說明
  const pts = [
    { icon: '🎊', title: '彩帶動畫', desc: '120 顆彩色粒子從畫面飄落，強化成就感' },
    { icon: '💰', title: '自動計入業績', desc: '已成交金額即時統計在 Dashboard' },
    { icon: '📅', title: '記錄成交日期', desc: '自動帶入成交時間，完整商機履歷' },
    { icon: '📈', title: '業績預測更新', desc: '預測達成率即時反映新成交案件' },
  ];
  pts.forEach((p, i) => {
    addCard(s, 0.4, 1.35 + i * 0.98, 5.5, 0.83);
    s.addText(p.icon, { x: 0.55, y: 1.42 + i * 0.98, w: 0.5, h: 0.5, fontSize: 20, align: 'center', fontFace: 'Arial' });
    s.addText(p.title, { x: 1.1, y: 1.44 + i * 0.98, w: 4.6, h: 0.32, fontSize: 14, bold: true, color: C.green, fontFace: 'Arial' });
    s.addText(p.desc, { x: 1.1, y: 1.74 + i * 0.98, w: 4.6, h: 0.28, fontSize: 11, color: C.gray, fontFace: 'Arial' });
  });

  addScreenshotBox(s, 6.1, 1.35, 6.6, 3.9, 'Won 成交慶祝動畫畫面');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 7 — 拜訪記錄
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '📋  拜訪記錄管理');
  addSubtitle(s, '每次客戶互動都有跡可循');
  addDivider(s, 1.22);

  const pts = [
    { icon: '📅', title: '日曆視覺化', desc: '業務日報以月曆呈現，快速回顧拜訪密度' },
    { icon: '🏷️', title: '多種拜訪類型', desc: '親訪、電訪、視訊、展覽、Email 分類記錄' },
    { icon: '✅', title: '下一步行動', desc: '每筆記錄附帶後續行動提醒，不漏接商機' },
    { icon: '🔗', title: '連結商機', desc: '拜訪同時可關聯商機，進度一致更新' },
  ];
  pts.forEach((p, i) => {
    addCard(s, 0.4, 1.35 + i * 0.98, 5.5, 0.83);
    s.addText(p.icon, { x: 0.55, y: 1.42 + i * 0.98, w: 0.5, h: 0.5, fontSize: 20, align: 'center', fontFace: 'Arial' });
    s.addText(p.title, { x: 1.1, y: 1.44 + i * 0.98, w: 4.6, h: 0.32, fontSize: 14, bold: true, color: C.white, fontFace: 'Arial' });
    s.addText(p.desc, { x: 1.1, y: 1.74 + i * 0.98, w: 4.6, h: 0.28, fontSize: 11, color: C.gray, fontFace: 'Arial' });
  });

  addScreenshotBox(s, 6.1, 1.35, 6.6, 3.9, '拜訪記錄日曆頁面');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 8 — AI 功能總覽
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);

  // 背景裝飾
  s.addShape(pptx.ShapeType.ellipse, { x: 3, y: 0.5, w: 7, h: 7, fill: { color: C.purple, transparency: 96 }, line: { color: C.purple, width: 0 } });

  addTitle(s, '🤖  AI 智能功能');
  s.addText('搭載 Google Gemini 1.5 Flash，免費額度充裕，無需信用卡', {
    x: 0.5, y: 0.9, w: 12.33, h: 0.35, fontSize: 13, color: C.gray, fontFace: 'Arial',
  });
  addDivider(s, 1.3);

  const ais = [
    {
      icon: '📷', num: '01', title: '手機拍照辨識名片',
      desc: '業務用手機對準名片拍照，AI 自動辨識填入\n姓名、公司、電話、Email 等所有欄位',
      tags: ['Gemini Vision', '手機相機', '秒填欄位'],
      color: C.cyan,
    },
    {
      icon: '✨', num: '02', title: '拜訪記錄 AI 助手',
      desc: '填完拜訪內容後一鍵分析，自動生成\n關鍵重點摘要與下一步行動建議',
      tags: ['關鍵重點', '行動建議', '省時省力'],
      color: C.green,
    },
    {
      icon: '📊', num: '03', title: '商機贏率預測',
      desc: '綜合商機階段、拜訪頻率、時程分析，\nAI 預測成交機率並說明關鍵因素',
      tags: ['勝率百分比', '推理說明', '24h 快取'],
      color: C.orange,
    },
    {
      icon: '👤', num: '04', title: '客戶輪廓 AI 摘要',
      desc: '聚合拜訪記錄與商機狀態，自動生成\n客戶關係健康度評估與150字摘要',
      tags: ['健康度評級', '關係摘要', '快取更新'],
      color: C.purpleL,
    },
  ];

  ais.forEach((ai, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 6.45;
    const y = 1.45 + row * 1.95;

    addCard(s, x, y, 6.2, 1.75);
    s.addShape(pptx.ShapeType.rect, { x, y, w: 6.2, h: 0.05, fill: { color: ai.color }, line: { color: ai.color, width: 0 } });

    s.addText(ai.icon, { x: x + 0.15, y: y + 0.12, w: 0.55, h: 0.55, fontSize: 24, align: 'center', fontFace: 'Arial' });
    s.addText(ai.num, { x: x + 5.4, y: y + 0.12, w: 0.65, h: 0.3, fontSize: 28, bold: true, color: ai.color, align: 'right', fontFace: 'Arial', transparency: 60 });
    s.addText(ai.title, { x: x + 0.75, y: y + 0.15, w: 5.0, h: 0.38, fontSize: 15, bold: true, color: C.white, fontFace: 'Arial' });
    s.addText(ai.desc, { x: x + 0.75, y: y + 0.55, w: 5.2, h: 0.55, fontSize: 11, color: C.grayL, fontFace: 'Arial', wrap: true });

    ai.tags.forEach((tag, j) => {
      s.addShape(pptx.ShapeType.roundRect, { x: x + 0.15 + j * 1.9, y: y + 1.35, w: 1.75, h: 0.25, fill: { color: ai.color, transparency: 85 }, line: { color: ai.color, width: 0.5 }, rectRadius: 0.12 });
      s.addText(tag, { x: x + 0.15 + j * 1.9, y: y + 1.35, w: 1.75, h: 0.25, fontSize: 9, bold: true, color: ai.color, align: 'center', valign: 'middle', fontFace: 'Arial' });
    });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 9 — AI 名片辨識（手機）
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '📷  手機拍照辨識名片');
  addSubtitle(s, '業務在外拜訪，拿到名片當場建立聯絡人');
  addDivider(s, 1.22);

  // 流程圖
  const steps = [
    { icon: '📱', label: '點「新增名片」', sub: '在手機 CRM 開啟' },
    { icon: '📷', label: '點拍照辨識', sub: '直接開後鏡頭相機' },
    { icon: '🤖', label: 'AI 辨識（5-10秒）', sub: 'Gemini Vision 分析' },
    { icon: '✅', label: '欄位自動填入', sub: '確認後一鍵儲存' },
  ];

  steps.forEach((st, i) => {
    const x = 0.5 + i * 3.15;
    // 圓形圖示
    s.addShape(pptx.ShapeType.ellipse, { x: x + 0.65, y: 1.4, w: 1.1, h: 1.1, fill: { color: C.cyan, transparency: 80 }, line: { color: C.cyan, width: 1.5 } });
    s.addText(st.icon, { x: x + 0.65, y: 1.4, w: 1.1, h: 1.1, fontSize: 28, align: 'center', valign: 'middle', fontFace: 'Arial' });

    s.addText(st.label, { x, y: 2.65, w: 2.75, h: 0.38, fontSize: 13, bold: true, color: C.white, align: 'center', fontFace: 'Arial' });
    s.addText(st.sub, { x, y: 3.0, w: 2.75, h: 0.28, fontSize: 11, color: C.gray, align: 'center', fontFace: 'Arial' });

    if (i < 3) {
      s.addText('→', { x: x + 2.8, y: 1.82, w: 0.5, h: 0.4, fontSize: 20, bold: true, color: C.cyan, align: 'center', fontFace: 'Arial' });
    }
  });

  // 可辨識欄位
  addCard(s, 0.5, 3.45, 12.3, 1.5);
  s.addText('🎯 可自動辨識的欄位', { x: 0.7, y: 3.55, w: 4, h: 0.38, fontSize: 14, bold: true, color: C.cyan, fontFace: 'Arial' });
  const fields = ['姓名（中/英）', '公司名稱', '職稱', '市話 + 分機', '手機', 'Email', '地址', '網址', '統一編號', '產業別'];
  fields.forEach((f, i) => {
    const col = i % 5;
    const row = Math.floor(i / 5);
    s.addShape(pptx.ShapeType.roundRect, { x: 0.7 + col * 2.38, y: 4.0 + row * 0.38, w: 2.2, h: 0.3, fill: { color: C.cyan, transparency: 88 }, line: { color: C.cyan, width: 0.5 }, rectRadius: 0.05 });
    s.addText('✓ ' + f, { x: 0.7 + col * 2.38, y: 4.0 + row * 0.38, w: 2.2, h: 0.3, fontSize: 10, color: C.cyan, align: 'center', valign: 'middle', fontFace: 'Arial' });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 10 — AI 拜訪建議
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '✨  拜訪記錄 AI 助手');
  addSubtitle(s, '填完拜訪內容，AI 立刻給出專業建議');
  addDivider(s, 1.22);

  // 左側說明
  addCard(s, 0.4, 1.35, 5.5, 1.0);
  s.addText('💬 AI 分析輸入', { x: 0.6, y: 1.45, w: 5.0, h: 0.35, fontSize: 13, bold: true, color: C.cyan, fontFace: 'Arial' });
  ['拜訪主題', '會談內容', '拜訪類型', '聯絡人姓名', '公司名稱'].forEach((item, i) => {
    s.addText('• ' + item, { x: 0.7, y: 1.85 + i * 0.0, w: 5.0, h: 0.26, fontSize: 11, color: C.grayL, fontFace: 'Arial' });
    // reduce overlap
    s.addText('• ' + item, { x: 0.7, y: 1.82 + i * 0.22, w: 5.0, h: 0.26, fontSize: 11, color: C.grayL, fontFace: 'Arial' });
  });

  s.addText('→', { x: 6.05, y: 2.55, w: 0.5, h: 0.4, fontSize: 22, bold: true, color: C.green, align: 'center', fontFace: 'Arial' });

  addCard(s, 6.65, 1.35, 6.1, 3.9);
  s.addText('🤖 AI 回傳結果', { x: 6.85, y: 1.45, w: 5.7, h: 0.35, fontSize: 13, bold: true, color: C.green, fontFace: 'Arial' });

  // 模擬 AI 回應
  addCard(s, 6.85, 1.88, 5.7, 0.9, '0a1e0a');
  s.addText('💡 建議下一步行動', { x: 7.0, y: 1.95, w: 5.4, h: 0.28, fontSize: 11, bold: true, color: C.green, fontFace: 'Arial' });
  s.addText('安排產品 Demo，邀請客戶資訊長參與，準備三個客戶成功案例', { x: 7.0, y: 2.22, w: 5.4, h: 0.45, fontSize: 10, color: C.grayL, fontFace: 'Arial', wrap: true });

  addCard(s, 6.85, 2.92, 5.7, 1.65, '0a0a1e');
  s.addText('🎯 關鍵重點摘要', { x: 7.0, y: 2.99, w: 5.4, h: 0.28, fontSize: 11, bold: true, color: C.cyan, fontFace: 'Arial' });
  ['客戶對現行系統效能不滿，採購預算已核准', '決策者為資訊長，本月底前需提案', '競品已報價，需強調技術支援差異化'].forEach((item, i) => {
    s.addText('• ' + item, { x: 7.0, y: 3.3 + i * 0.38, w: 5.4, h: 0.34, fontSize: 10, color: C.grayL, fontFace: 'Arial', wrap: true });
  });

  addCard(s, 0.4, 2.45, 5.5, 0.7);
  s.addText('⚡ 自動填入「下一步行動」欄位\n（若欄位空白時）', { x: 0.6, y: 2.52, w: 5.0, h: 0.56, fontSize: 11, color: C.yellow, fontFace: 'Arial' });

  addCard(s, 0.4, 3.25, 5.5, 0.8);
  s.addText('🆓 免費額度\n每日 100萬 tokens，15 次/分鐘', { x: 0.6, y: 3.32, w: 5.0, h: 0.66, fontSize: 11, color: C.green, fontFace: 'Arial' });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 11 — AI 贏率預測
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '📊  商機贏率 AI 預測');
  addSubtitle(s, '讓 AI 告訴你這筆單子把握幾成');
  addDivider(s, 1.22);

  // 左側分析因素
  s.addText('AI 分析維度', { x: 0.5, y: 1.3, w: 5.8, h: 0.38, fontSize: 14, bold: true, color: C.orange, fontFace: 'Arial' });
  const factors = [
    { icon: '🎯', label: '商機階段', desc: 'A探索 → B需求 → C報價 → D議約' },
    { icon: '📅', label: '拜訪活躍度', desc: '近30天/60天拜訪頻率分析' },
    { icon: '⏱️', label: '距截止天數', desc: '預計成交日與今日距離' },
    { icon: '💰', label: '案件金額', desc: '大案件需更多關注點' },
  ];
  factors.forEach((f, i) => {
    addCard(s, 0.5, 1.75 + i * 0.82, 5.7, 0.68);
    s.addText(f.icon + '  ' + f.label, { x: 0.7, y: 1.82 + i * 0.82, w: 5.3, h: 0.3, fontSize: 13, bold: true, color: C.white, fontFace: 'Arial' });
    s.addText(f.desc, { x: 0.7, y: 2.1 + i * 0.82, w: 5.3, h: 0.25, fontSize: 11, color: C.gray, fontFace: 'Arial' });
  });

  // 右側預測結果示意
  addCard(s, 6.5, 1.3, 6.3, 3.9);
  s.addText('🤖 AI 預測結果', { x: 6.7, y: 1.4, w: 5.9, h: 0.38, fontSize: 14, bold: true, color: C.orange, fontFace: 'Arial' });

  // 大數字
  s.addText('72%', { x: 6.7, y: 1.85, w: 3, h: 1.0, fontSize: 72, bold: true, color: C.orange, fontFace: 'Arial' });
  s.addText('預測贏率', { x: 9.7, y: 2.2, w: 2.8, h: 0.4, fontSize: 14, color: C.gray, fontFace: 'Arial' });

  // 進度條
  s.addShape(pptx.ShapeType.rect, { x: 6.7, y: 2.95, w: 5.8, h: 0.22, fill: { color: '1a1a1a' }, line: { color: C.orange, width: 0.5 } });
  s.addShape(pptx.ShapeType.rect, { x: 6.7, y: 2.95, w: 5.8 * 0.72, h: 0.22, fill: { color: C.orange, transparency: 20 }, line: { color: C.orange, width: 0 } });

  addCard(s, 6.7, 3.28, 5.8, 1.15, '1a1200');
  s.addText('📝 AI 分析說明', { x: 6.9, y: 3.35, w: 5.5, h: 0.28, fontSize: 11, bold: true, color: C.orange, fontFace: 'Arial' });
  s.addText('本商機已進入報價階段，近30天拜訪3次顯示關係活躍。距預計成交日尚有45天，建議本週確認預算核准流程，加速進入議約。', {
    x: 6.9, y: 3.65, w: 5.5, h: 0.72, fontSize: 10, color: C.grayL, fontFace: 'Arial', wrap: true,
  });

  addCard(s, 0.5, 5.05, 12.3, 0.35, C.card);
  s.addText('💡 結果自動快取 24 小時，不重複呼叫 AI，節省額度', { x: 0.5, y: 5.05, w: 12.3, h: 0.35, fontSize: 11, color: C.gray, align: 'center', valign: 'middle', fontFace: 'Arial' });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 12 — AI 客戶摘要
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '👤  客戶輪廓 AI 摘要');
  addSubtitle(s, '一秒了解客戶關係現狀，交接不怕資訊斷層');
  addDivider(s, 1.22);

  // 三個健康度卡
  const health = [
    { label: '✅ 關係良好', desc: '近期拜訪頻繁\n商機積極推進中', color: C.green },
    { label: '⚠️ 關係普通', desc: '近60天未拜訪\n商機停滯需跟進', color: C.yellow },
    { label: '🔴 需要關注', desc: '超過90天未聯繫\n商機流失風險高', color: C.red },
  ];
  health.forEach((h, i) => {
    addCard(s, 0.4 + i * 4.2, 1.3, 3.95, 1.5);
    s.addShape(pptx.ShapeType.rect, { x: 0.4 + i * 4.2, y: 1.3, w: 3.95, h: 0.05, fill: { color: h.color }, line: { color: h.color, width: 0 } });
    s.addShape(pptx.ShapeType.roundRect, { x: 0.6 + i * 4.2, y: 1.45, w: 3.5, h: 0.35, fill: { color: h.color, transparency: 80 }, line: { color: h.color, width: 1 }, rectRadius: 0.17 });
    s.addText(h.label, { x: 0.6 + i * 4.2, y: 1.45, w: 3.5, h: 0.35, fontSize: 13, bold: true, color: h.color, align: 'center', valign: 'middle', fontFace: 'Arial' });
    s.addText(h.desc, { x: 0.6 + i * 4.2, y: 1.9, w: 3.5, h: 0.7, fontSize: 11, color: C.grayL, align: 'center', fontFace: 'Arial', wrap: true });
  });

  // AI 摘要示例
  addCard(s, 0.4, 2.95, 12.4, 1.55);
  s.addText('🤖 AI 摘要範例', { x: 0.6, y: 3.05, w: 3, h: 0.35, fontSize: 13, bold: true, color: C.purpleL, fontFace: 'Arial' });
  s.addText(
    '王大明為永聯科技資訊部資深經理，是本公司核心決策聯絡人。過去三個月共拜訪5次，關係積極。目前進行中商機2筆（ERP導入$280萬、雲端移轉$150萬），預計Q3完成簽約。最近一次拜訪於7天前，討論報價細節，客戶反應正面。建議本週確認採購流程時程，安排高階主管拜會以加速決策。',
    { x: 0.6, y: 3.45, w: 12.0, h: 0.95, fontSize: 11, color: C.grayL, fontFace: 'Arial', wrap: true }
  );

  addCard(s, 0.4, 4.6, 12.4, 0.75, C.card);
  const features2 = ['聚合最近5筆拜訪', '進行中商機狀態', '上次拜訪天數', '關係健康度評級', '24小時快取'];
  features2.forEach((f, i) => {
    s.addShape(pptx.ShapeType.roundRect, { x: 0.6 + i * 2.42, y: 4.72, w: 2.25, h: 0.28, fill: { color: C.purple, transparency: 85 }, line: { color: C.purpleL, width: 0.5 }, rectRadius: 0.05 });
    s.addText('✦ ' + f, { x: 0.6 + i * 2.42, y: 4.72, w: 2.25, h: 0.28, fontSize: 9, color: C.purpleL, align: 'center', valign: 'middle', fontFace: 'Arial' });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 13 — 目標管理
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '🎯  目標 & 業績管理');
  addSubtitle(s, '設定目標、追蹤達成、預測趨勢');
  addDivider(s, 1.22);

  const pts = [
    { icon: '📅', title: '月度 / 季度目標設定', desc: '管理員為每位業務設定業績目標金額' },
    { icon: '📈', title: '即時達成率', desc: '已成交金額 vs 目標，百分比即時計算' },
    { icon: '🏆', title: 'Won 自動計入', desc: '商機移入 Won 階段，業績立即反映' },
    { icon: '📊', title: '多維度分析', desc: '個人 / 團隊 / 產品線多角度業績報表' },
  ];
  pts.forEach((p, i) => {
    addCard(s, 0.4, 1.35 + i * 0.98, 5.5, 0.83);
    s.addText(p.icon, { x: 0.55, y: 1.42 + i * 0.98, w: 0.5, h: 0.5, fontSize: 20, align: 'center', fontFace: 'Arial' });
    s.addText(p.title, { x: 1.1, y: 1.44 + i * 0.98, w: 4.6, h: 0.32, fontSize: 14, bold: true, color: C.white, fontFace: 'Arial' });
    s.addText(p.desc, { x: 1.1, y: 1.74 + i * 0.98, w: 4.6, h: 0.28, fontSize: 11, color: C.gray, fontFace: 'Arial' });
  });

  addScreenshotBox(s, 6.1, 1.35, 6.6, 3.9, 'Dashboard 業績統計');
}

// ═══════════════════════════════════════════════════════════
// SLIDE 14 — 管理員功能
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '👥  管理員功能');
  addSubtitle(s, '完整的後台控制中心');
  addDivider(s, 1.22);

  const features3 = [
    { icon: '👤', title: '用戶管理', items: ['新增/停用帳號', '角色設定（業務/主管/管理員）', '密碼重置'] },
    { icon: '📤', title: '資料匯出入', items: ['聯絡人 Excel 匯出', '批次匯入建立', '商機資料匯出'] },
    { icon: '🔄', title: '帳號移轉', items: ['業務離職客戶轉移', '商機/拜訪/應收同步移轉', '一鍵完成交接'] },
    { icon: '📷', title: 'AI 名片批次匯入', items: ['上傳名片圖片', 'AI 批次辨識', '一次匯入多筆聯絡人'] },
  ];

  features3.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 6.5;
    const y = 1.35 + row * 2.1;

    addCard(s, x, y, 6.2, 1.9);
    s.addShape(pptx.ShapeType.rect, { x, y, w: 6.2, h: 0.05, fill: { color: C.yellow }, line: { color: C.yellow, width: 0 } });
    s.addText(f.icon + ' ' + f.title, { x: x + 0.2, y: y + 0.12, w: 5.8, h: 0.42, fontSize: 16, bold: true, color: C.white, fontFace: 'Arial' });
    f.items.forEach((item, j) => {
      s.addText('• ' + item, { x: x + 0.3, y: y + 0.62 + j * 0.38, w: 5.7, h: 0.34, fontSize: 12, color: C.grayL, fontFace: 'Arial' });
    });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 15 — 系統架構
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);
  addTitle(s, '☁️  雲端系統架構');
  addSubtitle(s, '企業級可靠性，零維護成本');
  addDivider(s, 1.22);

  // 架構圖示意
  const layers = [
    { label: '前端', items: ['HTML / CSS / JavaScript', '響應式設計，手機電腦皆適用'], color: C.cyan, icon: '🖥️' },
    { label: 'API 伺服器', items: ['Node.js / Express', 'JWT 驗證，多層權限控管'], color: C.purple, icon: '⚙️' },
    { label: '資料庫', items: ['Supabase PostgreSQL', '資料即時同步，自動備份'], color: C.green, icon: '🗄️' },
    { label: 'AI 引擎', items: ['Google Gemini 1.5 Flash', '免費額度：每日100萬 tokens'], color: C.orange, icon: '🤖' },
  ];

  layers.forEach((l, i) => {
    const x = 0.4 + i * 3.2;
    addCard(s, x, 1.35, 2.95, 2.5);
    s.addShape(pptx.ShapeType.rect, { x, y: 1.35, w: 2.95, h: 0.05, fill: { color: l.color }, line: { color: l.color, width: 0 } });
    s.addText(l.icon, { x, y: 1.5, w: 2.95, h: 0.6, fontSize: 28, align: 'center', fontFace: 'Arial' });
    s.addText(l.label, { x, y: 2.12, w: 2.95, h: 0.38, fontSize: 15, bold: true, color: l.color, align: 'center', fontFace: 'Arial' });
    l.items.forEach((item, j) => {
      s.addText(item, { x: x + 0.1, y: 2.55 + j * 0.38, w: 2.75, h: 0.34, fontSize: 11, color: C.grayL, align: 'center', fontFace: 'Arial', wrap: true });
    });

    if (i < 3) {
      s.addText('↔', { x: x + 2.95, y: 2.35, w: 0.25, h: 0.4, fontSize: 16, bold: true, color: C.gray, align: 'center', fontFace: 'Arial' });
    }
  });

  // 部署平台
  addCard(s, 0.4, 4.0, 12.4, 1.35);
  s.addText('🚀 部署平台：Vercel', { x: 0.6, y: 4.1, w: 4, h: 0.38, fontSize: 14, bold: true, color: C.white, fontFace: 'Arial' });
  const vercelFeatures = ['全球 CDN 加速', 'Git Push 自動部署', '30秒無伺服器函數', '免費 SSL 憑證', '自訂網域（itts-crm.vercel.app）'];
  vercelFeatures.forEach((f, i) => {
    s.addText('✓ ' + f, { x: 0.6 + i * 2.45, y: 4.55, w: 2.3, h: 0.6, fontSize: 11, color: C.green, fontFace: 'Arial' });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 16 — 結語 / 價值主張
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s);

  // 背景裝飾
  s.addShape(pptx.ShapeType.ellipse, { x: -1, y: -1, w: 6, h: 6, fill: { color: C.cyan, transparency: 96 }, line: { color: C.cyan, width: 0 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 9, y: 1, w: 5, h: 5, fill: { color: C.purple, transparency: 94 }, line: { color: C.purple, width: 0 } });

  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.08, h: 5.63, fill: { color: C.green }, line: { color: C.green, width: 0 } });

  s.addText('ITTS CRM', { x: 0.5, y: 0.6, w: 12.33, h: 0.7, fontSize: 42, bold: true, color: C.white, fontFace: 'Arial' });
  s.addText('讓每位業務都有 AI 輔助的超級助手', { x: 0.5, y: 1.3, w: 10, h: 0.55, fontSize: 24, color: C.cyan, fontFace: 'Arial' });

  addDivider(s, 2.0);

  const values = [
    { icon: '⏱️', val: '50%↓', label: '名片建檔時間' },
    { icon: '📊', val: '100%', label: '商機視覺化' },
    { icon: '🤖', val: '4項', label: 'AI 智能功能' },
    { icon: '☁️', val: '24/7', label: '雲端隨時存取' },
  ];
  values.forEach((v, i) => {
    addCard(s, 0.5 + i * 3.2, 2.15, 2.95, 1.6);
    s.addText(v.icon, { x: 0.5 + i * 3.2, y: 2.25, w: 2.95, h: 0.5, fontSize: 26, align: 'center', fontFace: 'Arial' });
    s.addText(v.val, { x: 0.5 + i * 3.2, y: 2.75, w: 2.95, h: 0.55, fontSize: 30, bold: true, color: C.cyan, align: 'center', fontFace: 'Arial' });
    s.addText(v.label, { x: 0.5 + i * 3.2, y: 3.3, w: 2.95, h: 0.35, fontSize: 12, color: C.gray, align: 'center', fontFace: 'Arial' });
  });

  s.addText('系統持續迭代，功能由實際業務需求驅動', { x: 0.5, y: 3.95, w: 12.33, h: 0.38, fontSize: 14, color: C.gray, align: 'center', fontFace: 'Arial' });

  addCard(s, 0.5, 4.45, 12.33, 0.8, C.card2);
  s.addText('🌐  https://itts-crm.vercel.app', { x: 0.5, y: 4.45, w: 12.33, h: 0.8, fontSize: 22, bold: true, color: C.cyan, align: 'center', valign: 'middle', fontFace: 'Arial' });
}

// ── 輸出 ──────────────────────────────────────────────────
const outPath = 'C:/Users/steven.lee/Documents/ITTS_CRM_產品亮點.pptx';
pptx.writeFile({ fileName: outPath })
  .then(() => console.log('✅ PPT 已儲存至：' + outPath))
  .catch(e => console.error('❌ 錯誤：', e));
