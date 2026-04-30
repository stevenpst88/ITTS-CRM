const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';

// ── 色彩系統（商務簡潔：深藍 + 白 + 淺灰）─────────────────
const C = {
  navy:    '1e3a5f',
  blue:    '2563eb',
  blueL:   '3b82f6',
  blueBg:  'eff6ff',
  white:   'ffffff',
  gray1:   'f8fafc',   // 背景
  gray2:   'f1f5f9',   // 淺卡片
  gray3:   'e2e8f0',   // 分隔線
  gray4:   '94a3b8',   // 次要文字
  text:    '1e293b',   // 主文字
  textS:   '475569',   // 次文字
  green:   '16a34a',
  greenBg: 'f0fdf4',
  orange:  'ea580c',
  purple:  '7c3aed',
  cyan:    '0891b2',
};

// ── helper ────────────────────────────────────────────────
function setBg(slide, color = C.white) {
  slide.background = { color };
}

function hline(slide, y, x = 0.5, w = 12.33, color = C.gray3) {
  slide.addShape(pptx.ShapeType.line, { x, y, w, h: 0, line: { color, width: 0.8 } });
}

function tag(slide, text, x, y, color = C.blue) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w: 1.5, h: 0.28, fill: { color, transparency: 88 }, line: { color, width: 0 }, rectRadius: 0.14 });
  slide.addText(text, { x, y, w: 1.5, h: 0.28, fontSize: 9, bold: true, color, align: 'center', valign: 'middle', fontFace: 'Arial' });
}

function card(slide, x, y, w, h, fillColor = C.white, borderColor = C.gray3) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color: fillColor }, line: { color: borderColor, width: 0.8 }, rectRadius: 0.1 });
}

function titleBar(slide) {
  // 左側藍色豎線裝飾
  slide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.38, w: 0.06, h: 0.52, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
}

function slideTitle(slide, title, sub = '') {
  titleBar(slide);
  slide.addText(title, { x: 0.68, y: 0.33, w: 11.5, h: 0.55, fontSize: 22, bold: true, color: C.navy, fontFace: 'Arial' });
  if (sub) slide.addText(sub, { x: 0.68, y: 0.87, w: 11.5, h: 0.3, fontSize: 12, color: C.gray4, fontFace: 'Arial' });
  hline(slide, sub ? 1.2 : 1.05);
}

function icon(slide, emoji, x, y, size = 24) {
  slide.addText(emoji, { x, y, w: 0.6, h: 0.6, fontSize: size, align: 'center', valign: 'middle', fontFace: 'Arial' });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 1 — 封面
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s, C.navy);

  // 右側幾何裝飾
  s.addShape(pptx.ShapeType.rect, { x: 8.8, y: 0, w: 4, h: 5.63, fill: { color: '162d4a' }, line: { color: '162d4a', width: 0 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 9.5, y: 0.8, w: 2.8, h: 2.8, fill: { color: C.blue, transparency: 85 }, line: { color: C.blue, width: 0 } });
  s.addShape(pptx.ShapeType.ellipse, { x: 10.2, y: 2.8, w: 1.8, h: 1.8, fill: { color: C.blueL, transparency: 80 }, line: { color: C.blueL, width: 0 } });

  // 底部白色橫條
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 5.1, w: 13.33, h: 0.53, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
  s.addText('ITTS — 東捷資訊服務股份有限公司　｜　2026 Q2', {
    x: 0.5, y: 5.1, w: 12.5, h: 0.53, fontSize: 11, color: 'cce4ff', fontFace: 'Arial', valign: 'middle',
  });

  // 主標題
  s.addText('ITTS CRM', { x: 0.7, y: 0.9, w: 8, h: 0.95, fontSize: 56, bold: true, color: C.white, fontFace: 'Arial' });
  s.addShape(pptx.ShapeType.rect, { x: 0.7, y: 1.82, w: 3.5, h: 0.05, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
  s.addText('智能業務管理平台', { x: 0.7, y: 1.95, w: 8, h: 0.6, fontSize: 26, color: 'cce4ff', fontFace: 'Arial' });
  s.addText('業務數位化 × AI 智能分析 × 雲端協作', { x: 0.7, y: 2.65, w: 8, h: 0.38, fontSize: 14, color: '8db8e8', fontFace: 'Arial' });

  // 四個 KPI
  const kpis = [
    { icon: '🗂️', val: '名片管理' },
    { icon: '📊', val: '商機看板' },
    { icon: '🤖', val: 'AI 智能分析' },
    { icon: '☁️', val: '雲端部署' },
  ];
  kpis.forEach((k, i) => {
    const x = 0.7 + i * 2.12;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 3.35, w: 1.95, h: 1.2, fill: { color: '162d4a' }, line: { color: C.blue, width: 0.8, transparency: 50 }, rectRadius: 0.1 });
    s.addText(k.icon, { x, y: 3.42, w: 1.95, h: 0.52, fontSize: 26, align: 'center', fontFace: 'Arial' });
    s.addText(k.val, { x, y: 3.92, w: 1.95, h: 0.35, fontSize: 11, bold: true, color: 'cce4ff', align: 'center', fontFace: 'Arial' });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 2 — 系統功能一覽
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s, C.white);
  slideTitle(s, '系統功能一覽', '六大核心模組，涵蓋業務全流程');

  const features = [
    { icon: '🗂️', title: '聯絡人管理', desc: '名片數位化、多欄位搜尋、個人資訊記錄', color: C.blue },
    { icon: '📊', title: '商機看板', desc: 'Kanban 拖曳、階段追蹤、Won 成交統計', color: C.green },
    { icon: '📋', title: '拜訪記錄', desc: '結構化記錄、日曆視覺化、下一步行動', color: C.purple },
    { icon: '🎯', title: '目標業績', desc: '月度目標設定、達成率追蹤、報表分析', color: C.orange },
    { icon: '🤖', title: 'AI 智能分析', desc: '名片辨識、拜訪建議、贏率預測、客戶摘要', color: C.cyan },
    { icon: '👥', title: '管理員後台', desc: '用戶管理、資料匯出入、帳號移轉', color: '64748b' },
  ];

  features.forEach((f, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.5 + col * 4.28;
    const y = 1.35 + row * 1.88;

    card(s, x, y, 4.05, 1.68, C.white, C.gray3);
    s.addShape(pptx.ShapeType.rect, { x, y, w: 4.05, h: 0.05, fill: { color: f.color }, line: { color: f.color, width: 0 } });
    s.addText(f.icon, { x: x + 0.15, y: y + 0.12, w: 0.6, h: 0.6, fontSize: 24, align: 'center', fontFace: 'Arial' });
    s.addText(f.title, { x: x + 0.8, y: y + 0.15, w: 3.1, h: 0.42, fontSize: 15, bold: true, color: C.text, fontFace: 'Arial' });
    s.addText(f.desc, { x: x + 0.15, y: y + 0.72, w: 3.7, h: 0.7, fontSize: 11, color: C.textS, fontFace: 'Arial', wrap: true });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 3 — 業務工作流程
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s, C.white);
  slideTitle(s, '業務數位化工作流程', '從拿到名片到成交，全程在 CRM 管理');

  const steps = [
    { icon: '📷', step: '01', title: '名片辨識', desc: '手機拍照\nAI 自動建立聯絡人' },
    { icon: '💼', step: '02', title: '建立商機', desc: '掛載商機\n設定金額與預計日期' },
    { icon: '📋', step: '03', title: '拜訪記錄', desc: '每次拜訪\nAI 建議下一步行動' },
    { icon: '📊', step: '04', title: '看板推進', desc: '拖曳卡片\n即時更新商機階段' },
    { icon: '🏆', step: '05', title: '成交 Won！', desc: '業績自動計入\n慶祝動畫觸發' },
  ];

  steps.forEach((st, i) => {
    const x = 0.5 + i * 2.57;

    // 圓形圖示
    s.addShape(pptx.ShapeType.ellipse, { x: x + 0.55, y: 1.35, w: 1.3, h: 1.3, fill: { color: C.blueBg }, line: { color: C.blue, width: 1.5 } });
    s.addText(st.icon, { x: x + 0.55, y: 1.35, w: 1.3, h: 1.3, fontSize: 30, align: 'center', valign: 'middle', fontFace: 'Arial' });

    // 步驟號碼
    s.addShape(pptx.ShapeType.ellipse, { x: x + 1.5, y: 1.25, w: 0.38, h: 0.38, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
    s.addText(st.step, { x: x + 1.5, y: 1.25, w: 0.38, h: 0.38, fontSize: 9, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: 'Arial' });

    // 箭頭
    if (i < 4) {
      s.addText('→', { x: x + 2.1, y: 1.72, w: 0.55, h: 0.4, fontSize: 18, color: C.gray3, align: 'center', bold: true, fontFace: 'Arial' });
    }

    // 文字
    s.addText(st.title, { x, y: 2.82, w: 2.4, h: 0.38, fontSize: 13, bold: true, color: C.navy, align: 'center', fontFace: 'Arial' });
    s.addText(st.desc, { x, y: 3.22, w: 2.4, h: 0.65, fontSize: 11, color: C.textS, align: 'center', fontFace: 'Arial', wrap: true });
  });

  // 底部提示
  s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 4.1, w: 12.33, h: 0.55, fill: { color: C.blueBg }, line: { color: C.blue, width: 0 }, rectRadius: 0.08 });
  s.addText('💡  整個流程都在瀏覽器完成，手機電腦皆可使用，無需安裝任何軟體', {
    x: 0.5, y: 4.1, w: 12.33, h: 0.55, fontSize: 12, color: C.blue, align: 'center', valign: 'middle', fontFace: 'Arial',
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 4 — 商機看板示意圖
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s, C.white);
  slideTitle(s, '商機看板（Kanban）', '拖曳即更新，進度透明可視');

  const stages = [
    { label: 'A　探索接觸', color: '94a3b8', cards: ['三晃實業 ERP', '台灣水泥系統'], amounts: ['$120萬', '$85萬'] },
    { label: 'B　需求確認', color: C.blue, cards: ['永聯科技導入', '長庚資訊整合'], amounts: ['$280萬', '$160萬'] },
    { label: 'C　方案報價', color: C.purple, cards: ['台塑集團升級'], amounts: ['$450萬'] },
    { label: 'D　議約收尾', color: C.orange, cards: ['中鋼雲端移轉', '遠東新訂閱'], amounts: ['$320萬', '$95萬'] },
    { label: '🏆　Won 成交', color: C.green, cards: ['中華電信 MA'], amounts: ['$240萬'] },
  ];

  stages.forEach((st, i) => {
    const x = 0.42 + i * 2.55;

    // 欄位標題
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.28, w: 2.35, h: 0.38, fill: { color: st.color, transparency: 15 }, line: { color: st.color, width: 0 }, rectRadius: 0.06 });
    s.addText(st.label, { x, y: 1.28, w: 2.35, h: 0.38, fontSize: 10, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: 'Arial' });

    // 卡片
    st.cards.forEach((c2, j) => {
      card(s, x + 0.06, 1.76 + j * 1.0, 2.23, 0.82, C.white, C.gray3);
      s.addShape(pptx.ShapeType.rect, { x: x + 0.06, y: 1.76 + j * 1.0, w: 0.05, h: 0.82, fill: { color: st.color }, line: { color: st.color, width: 0 } });
      s.addText(c2, { x: x + 0.16, y: 1.82 + j * 1.0, w: 2.0, h: 0.3, fontSize: 10, bold: true, color: C.text, fontFace: 'Arial', wrap: true });
      s.addText(st.amounts[j], { x: x + 0.16, y: 2.1 + j * 1.0, w: 2.0, h: 0.25, fontSize: 10, color: st.color, bold: true, fontFace: 'Arial' });
      s.addText('• 近期有拜訪', { x: x + 0.16, y: 2.33 + j * 1.0, w: 2.0, h: 0.2, fontSize: 9, color: C.gray4, fontFace: 'Arial' });
    });
  });

  // 底部統計
  hline(s, 4.05);
  const tots = [['進行中', '6 筆'], ['預估金額', '$1,510萬'], ['本月已成交', '$240萬'], ['成交率', '14%']];
  tots.forEach((t, i) => {
    s.addText(t[0], { x: 1.2 + i * 3.0, y: 4.15, w: 2.4, h: 0.28, fontSize: 10, color: C.gray4, align: 'center', fontFace: 'Arial' });
    s.addText(t[1], { x: 1.2 + i * 3.0, y: 4.42, w: 2.4, h: 0.38, fontSize: 18, bold: true, color: C.navy, align: 'center', fontFace: 'Arial' });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 5 — AI 四大功能
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s, C.gray1);
  slideTitle(s, '🤖  AI 四大智能功能', '搭載 Google Gemini，免費方案即可使用');

  const ais = [
    {
      icon: '📷', title: '手機拍照辨識名片',
      before: '拿到名片手動輸入10分鐘',
      after: '拍照 → AI 辨識 → 10秒建檔',
      color: C.blue,
    },
    {
      icon: '✨', title: '拜訪記錄 AI 建議',
      before: '拜訪後腦中回想下一步',
      after: 'AI 自動摘要重點 + 建議行動',
      color: C.purple,
    },
    {
      icon: '📊', title: '商機贏率預測',
      before: '靠經驗主觀判斷把握度',
      after: 'AI 綜合分析給出百分比',
      color: C.orange,
    },
    {
      icon: '👤', title: '客戶輪廓摘要',
      before: '交接時翻記錄整理費時',
      after: 'AI 一鍵生成150字關係摘要',
      color: C.green,
    },
  ];

  ais.forEach((ai, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 6.45;
    const y = 1.35 + row * 1.98;

    card(s, x, y, 6.2, 1.78, C.white, C.gray3);
    s.addShape(pptx.ShapeType.rect, { x, y, w: 6.2, h: 0.04, fill: { color: ai.color }, line: { color: ai.color, width: 0 } });

    // 圖示
    s.addShape(pptx.ShapeType.ellipse, { x: x + 0.18, y: y + 0.14, w: 0.72, h: 0.72, fill: { color: ai.color, transparency: 88 }, line: { color: ai.color, width: 0 } });
    s.addText(ai.icon, { x: x + 0.18, y: y + 0.14, w: 0.72, h: 0.72, fontSize: 22, align: 'center', valign: 'middle', fontFace: 'Arial' });

    s.addText(ai.title, { x: x + 1.05, y: y + 0.16, w: 5.0, h: 0.36, fontSize: 14, bold: true, color: C.navy, fontFace: 'Arial' });

    // Before / After
    s.addShape(pptx.ShapeType.roundRect, { x: x + 0.18, y: y + 0.96, w: 2.68, h: 0.56, fill: { color: 'fef2f2' }, line: { color: 'fca5a5', width: 0.5 }, rectRadius: 0.06 });
    s.addText('❌  ' + ai.before, { x: x + 0.18, y: y + 0.96, w: 2.68, h: 0.56, fontSize: 10, color: 'dc2626', valign: 'middle', fontFace: 'Arial', wrap: true });

    s.addText('→', { x: x + 2.94, y: y + 1.05, w: 0.35, h: 0.38, fontSize: 16, color: C.gray4, align: 'center', bold: true, fontFace: 'Arial' });

    s.addShape(pptx.ShapeType.roundRect, { x: x + 3.35, y: y + 0.96, w: 2.68, h: 0.56, fill: { color: C.greenBg }, line: { color: '86efac', width: 0.5 }, rectRadius: 0.06 });
    s.addText('✅  ' + ai.after, { x: x + 3.35, y: y + 0.96, w: 2.68, h: 0.56, fontSize: 10, color: C.green, valign: 'middle', fontFace: 'Arial', wrap: true });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 6 — 權限架構
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s, C.white);
  slideTitle(s, '多層權限管理', '依角色控管可視範圍，資料安全有保障');

  const roles = [
    {
      icon: '🛡️', title: '系統管理員', color: C.navy,
      perms: ['查看所有用戶資料', '新增/停用帳號', '客戶資料批次匯出入', '帳號移轉作業', 'AI 名片批次辨識', '所有報表存取'],
    },
    {
      icon: '👔', title: '主管', color: C.blue,
      perms: ['查看所有業務資料', '商機看板總覽', '業績目標設定', '拜訪記錄查閱', '下載客戶資料', '業務績效報表'],
    },
    {
      icon: '💼', title: '業務', color: C.cyan,
      perms: ['管理自己的聯絡人', '維護自身商機', '記錄拜訪內容', 'AI 輔助功能', '手機拍照建檔', '自助更改密碼'],
    },
  ];

  roles.forEach((r, i) => {
    const x = 0.5 + i * 4.28;
    card(s, x, 1.35, 4.05, 3.85, C.white, C.gray3);
    s.addShape(pptx.ShapeType.rect, { x, y: 1.35, w: 4.05, h: 0.05, fill: { color: r.color }, line: { color: r.color, width: 0 } });

    s.addShape(pptx.ShapeType.ellipse, { x: x + 1.5, y: 1.52, w: 1.05, h: 1.05, fill: { color: r.color, transparency: 88 }, line: { color: r.color, width: 1.5 } });
    s.addText(r.icon, { x: x + 1.5, y: 1.52, w: 1.05, h: 1.05, fontSize: 28, align: 'center', valign: 'middle', fontFace: 'Arial' });
    s.addText(r.title, { x, y: 2.68, w: 4.05, h: 0.4, fontSize: 15, bold: true, color: r.color, align: 'center', fontFace: 'Arial' });

    hline(s, 3.15, x + 0.3, 3.45, C.gray3);

    r.perms.forEach((p, j) => {
      s.addText('✓  ' + p, { x: x + 0.25, y: 3.26 + j * 0.31, w: 3.55, h: 0.28, fontSize: 11, color: C.textS, fontFace: 'Arial' });
    });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 7 — 系統架構
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s, C.white);
  slideTitle(s, '雲端系統架構', '零維護成本，全球 CDN 加速，99.9% 可用性');

  // 三層架構示意
  const layers = [
    { icon: '🌐', label: '使用者端', sub: '任何瀏覽器 / 手機', items: ['Chrome / Safari / Edge', '手機直接拍照上傳'], color: C.blue },
    { icon: '⚙️', label: 'API 伺服器', sub: 'Node.js on Vercel', items: ['JWT 身份驗證', 'Serverless 函數 30s'], color: C.purple },
    { icon: '🗄️', label: '資料庫', sub: 'Supabase PostgreSQL', items: ['自動備份', '連線池加速'], color: C.green },
    { icon: '🤖', label: 'AI 引擎', sub: 'Google Gemini', items: ['免費 1,500次/天', '視覺辨識支援'], color: C.orange },
  ];

  layers.forEach((l, i) => {
    const x = 0.5 + i * 3.22;
    card(s, x, 1.35, 3.0, 2.75, l.color === C.blue ? C.blueBg : C.gray1, C.gray3);

    s.addShape(pptx.ShapeType.ellipse, { x: x + 1.05, y: 1.5, w: 0.9, h: 0.9, fill: { color: l.color, transparency: 82 }, line: { color: l.color, width: 1.5 } });
    s.addText(l.icon, { x: x + 1.05, y: 1.5, w: 0.9, h: 0.9, fontSize: 24, align: 'center', valign: 'middle', fontFace: 'Arial' });

    s.addText(l.label, { x, y: 2.52, w: 3.0, h: 0.35, fontSize: 14, bold: true, color: l.color, align: 'center', fontFace: 'Arial' });
    s.addText(l.sub, { x, y: 2.85, w: 3.0, h: 0.28, fontSize: 10, color: C.gray4, align: 'center', fontFace: 'Arial' });
    hline(s, 3.18, x + 0.2, 2.6, C.gray3);
    l.items.forEach((item, j) => {
      s.addText('• ' + item, { x: x + 0.2, y: 3.28 + j * 0.32, w: 2.6, h: 0.28, fontSize: 10, color: C.textS, fontFace: 'Arial' });
    });

    if (i < 3) {
      s.addText('→', { x: x + 3.0, y: 2.55, w: 0.22, h: 0.35, fontSize: 16, color: C.gray4, align: 'center', bold: true, fontFace: 'Arial' });
    }
  });

  // Vercel 說明
  s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 4.25, w: 12.33, h: 0.65, fill: { color: C.blueBg }, line: { color: C.blue, width: 0 }, rectRadius: 0.08 });
  s.addText('🚀  部署於 Vercel　｜　自動 HTTPS　｜　Git Push 即部署　｜　全球 CDN 加速　｜　網址：itts-crm.vercel.app', {
    x: 0.5, y: 4.25, w: 12.33, h: 0.65, fontSize: 12, color: C.blue, align: 'center', valign: 'middle', fontFace: 'Arial',
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 8 — 導入效益 & 結語
// ═══════════════════════════════════════════════════════════
{
  const s = pptx.addSlide();
  setBg(s, C.navy);

  // 底部藍條
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 5.1, w: 13.33, h: 0.53, fill: { color: C.blue }, line: { color: C.blue, width: 0 } });
  s.addText('itts-crm.vercel.app', { x: 0, y: 5.1, w: 13.33, h: 0.53, fontSize: 13, color: 'cce4ff', align: 'center', valign: 'middle', fontFace: 'Arial' });

  s.addText('導入效益', { x: 0.5, y: 0.35, w: 12.33, h: 0.55, fontSize: 22, bold: true, color: 'cce4ff', fontFace: 'Arial' });
  hline(s, 0.95, 0.5, 12.33, '2d5a8e');

  const benefits = [
    { icon: '⏱️', num: '80%', label: '建檔時間節省', desc: 'AI 拍照辨識取代手動輸入' },
    { icon: '📊', num: '100%', label: '商機透明化', desc: 'Kanban 即時反映所有進度' },
    { icon: '🤖', num: '4 項', label: 'AI 輔助功能', desc: '免費方案即可全數使用' },
    { icon: '📱', num: '0', label: '額外軟體安裝', desc: '瀏覽器直接使用，手機友善' },
  ];

  benefits.forEach((b, i) => {
    const x = 0.5 + i * 3.22;
    s.addShape(pptx.ShapeType.roundRect, { x, y: 1.1, w: 3.0, h: 2.8, fill: { color: '162d4a' }, line: { color: '2d5a8e', width: 0.8 }, rectRadius: 0.1 });

    s.addText(b.icon, { x, y: 1.22, w: 3.0, h: 0.62, fontSize: 30, align: 'center', fontFace: 'Arial' });
    s.addText(b.num, { x, y: 1.88, w: 3.0, h: 0.72, fontSize: 38, bold: true, color: C.white, align: 'center', fontFace: 'Arial' });
    s.addText(b.label, { x, y: 2.6, w: 3.0, h: 0.35, fontSize: 12, bold: true, color: 'cce4ff', align: 'center', fontFace: 'Arial' });
    hline(s, 3.0, x + 0.3, 2.4, '2d5a8e');
    s.addText(b.desc, { x, y: 3.1, w: 3.0, h: 0.55, fontSize: 10, color: '8db8e8', align: 'center', fontFace: 'Arial', wrap: true });
  });

  // 結語
  s.addShape(pptx.ShapeType.roundRect, { x: 0.5, y: 4.05, w: 12.33, h: 0.85, fill: { color: C.blue }, line: { color: C.blue, width: 0 }, rectRadius: 0.08 });
  s.addText('ITTS CRM 以業務實際需求為核心，持續迭代升級，讓每位業務都有 AI 加持的競爭優勢', {
    x: 0.5, y: 4.05, w: 12.33, h: 0.85, fontSize: 14, bold: true, color: C.white, align: 'center', valign: 'middle', fontFace: 'Arial',
  });
}

// ── 輸出 ──────────────────────────────────────────────────
const outPath = 'C:/Users/steven.lee/Documents/ITTS_CRM_商務版.pptx';
pptx.writeFile({ fileName: outPath })
  .then(() => console.log('✅ 已儲存：' + outPath))
  .catch(e => console.error('❌', e));
