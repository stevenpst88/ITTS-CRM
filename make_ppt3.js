const PptxGenJS = require('pptxgenjs');
const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';

const C = {
  navy:   '1e3a5f', blue:   '2563eb', blueL:  '3b82f6',
  blueBg: 'eff6ff', white:  'ffffff', gray1:  'f8fafc',
  gray2:  'f1f5f9', gray3:  'e2e8f0', gray4:  '94a3b8',
  text:   '1e293b', textS:  '475569',
  green:  '16a34a', greenB: 'f0fdf4',
  orange: 'ea580c', red:    'dc2626', redB: 'fef2f2',
  purple: '7c3aed', cyan:   '0891b2',
};

const F = 'Arial';

function bg(s, c = C.white) { s.background = { color: c }; }
function hln(s, y, x=0.5, w=12.33, c=C.gray3) {
  s.addShape(pptx.ShapeType.line,{x,y,w,h:0,line:{color:c,width:0.8}});
}
function rnd(s,x,y,w,h,fill,line,r=0.1) {
  s.addShape(pptx.ShapeType.roundRect,{x,y,w,h,fill:{color:fill},line:{color:line,width:0.8},rectRadius:r});
}
function txt(s,t,x,y,w,h,opts={}) {
  s.addText(t,{x,y,w,h,fontFace:F,...opts});
}
function emo(s,e,x,y,sz=36) {
  s.addText(e,{x,y,w:sz/36*1.1,h:sz/36*1.1,fontSize:sz,align:'center',valign:'middle',fontFace:F});
}
function hdr(s,title,sub='') {
  s.addShape(pptx.ShapeType.rect,{x:0.5,y:0.36,w:0.06,h:0.52,fill:{color:C.blue},line:{color:C.blue,width:0}});
  txt(s,title,0.68,0.32,11.5,0.58,{fontSize:22,bold:true,color:C.navy});
  if(sub) txt(s,sub,0.68,0.88,11.5,0.3,{fontSize:12,color:C.gray4});
  hln(s,sub?1.22:1.08);
}
function bigNum(s,num,label,x,y,color=C.blue) {
  txt(s,num,x,y,3,0.85,{fontSize:52,bold:true,color,align:'center'});
  txt(s,label,x,y+0.82,3,0.35,{fontSize:13,color:C.textS,align:'center'});
}

// ══════════════════════════════════════════════════════════
// S1 封面
// ══════════════════════════════════════════════════════════
{
  const s=pptx.addSlide(); bg(s,C.navy);
  s.addShape(pptx.ShapeType.rect,{x:8.6,y:0,w:4.73,h:5.63,fill:{color:'162d4a'},line:{color:'162d4a',width:0}});
  s.addShape(pptx.ShapeType.ellipse,{x:9.3,y:0.5,w:3.2,h:3.2,fill:{color:C.blue,transparency:87},line:{color:C.blue,width:0}});
  s.addShape(pptx.ShapeType.ellipse,{x:10.5,y:3.0,w:2.0,h:2.0,fill:{color:C.blueL,transparency:83},line:{color:C.blueL,width:0}});

  // 大圖示組合
  txt(s,'🗂️',1.0,0.7,1.5,1.5,{fontSize:60,align:'center'});
  txt(s,'📊',2.3,1.4,1.5,1.5,{fontSize:60,align:'center'});
  txt(s,'🤖',0.4,1.9,1.5,1.5,{fontSize:60,align:'center'});

  txt(s,'ITTS CRM',3.5,0.75,8,1.05,{fontSize:58,bold:true,color:C.white});
  s.addShape(pptx.ShapeType.rect,{x:3.5,y:1.76,w:3.8,h:0.05,fill:{color:C.blue},line:{color:C.blue,width:0}});
  txt(s,'主管看得到　業務省下來',3.5,1.88,8,0.62,{fontSize:26,color:'cce4ff'});
  txt(s,'即時數據 × AI 輔助 × 告別 Excel 地獄',3.5,2.6,8,0.38,{fontSize:14,color:'8db8e8'});

  // 底列
  s.addShape(pptx.ShapeType.rect,{x:0,y:5.1,w:13.33,h:0.53,fill:{color:C.blue},line:{color:C.blue,width:0}});
  txt(s,'ITTS 東捷資訊服務股份有限公司　|　2026 Q2',0,5.1,13.33,0.53,{fontSize:11,color:'cce4ff',align:'center',valign:'middle'});

  // KPI 列
  const ks=[{e:'🕐',v:'80%',l:'建檔時間節省'},{e:'📡',v:'即時',l:'數據同步'},{e:'🤖',v:'4項',l:'AI功能免費'},{e:'❌',v:'0',l:'紙本Excel傳遞'}];
  ks.forEach((k,i)=>{
    const x=3.5+i*2.55;
    rnd(s,x,3.15,2.35,1.45,'162d4a','2d5a8e',0.1);
    txt(s,k.e,x,3.2,2.35,0.55,{fontSize:26,align:'center'});
    txt(s,k.v,x,3.72,2.35,0.45,{fontSize:22,bold:true,color:C.white,align:'center'});
    txt(s,k.l,x,4.12,2.35,0.3,{fontSize:10,color:'8db8e8',align:'center'});
  });
}

// ══════════════════════════════════════════════════════════
// S2 主管的煩惱（大圖示）
// ══════════════════════════════════════════════════════════
{
  const s=pptx.addSlide(); bg(s,C.gray1);
  hdr(s,'每週一，這個場景您熟悉嗎？','週會前的 Excel 地獄');

  // 時間軸場景
  const scenes=[
    {e:'📩',time:'週一 08:00',title:'秘書發信給業務',desc:'請在今天中午前\n回傳業績Excel'},
    {e:'🤦',time:'週一 11:58',title:'業務趕工填表',desc:'手忙腳亂\n數字東拼西湊'},
    {e:'🔀',time:'週一 12:30',title:'版本混亂',desc:'收到5個不同版本\n「最終版_v3_真的最終」'},
    {e:'😤',time:'週一 14:00',title:'主管看到舊數據',desc:'有人在開會前\n又更新了一筆'},
  ];
  scenes.forEach((sc,i)=>{
    const x=0.42+i*3.25;
    rnd(s,x,1.32,3.05,3.55,C.white,C.gray3,0.12);
    s.addShape(pptx.ShapeType.rect,{x,y:1.32,w:3.05,h:0.05,fill:{color:C.red},line:{color:C.red,width:0}});
    txt(s,sc.e,x,1.45,3.05,1.15,{fontSize:50,align:'center'});
    rnd(s,x+0.25,2.65,2.55,0.3,C.redB,'fca5a5',0.15);
    txt(s,sc.time,x+0.25,2.65,2.55,0.3,{fontSize:9,bold:true,color:C.red,align:'center',valign:'middle'});
    txt(s,sc.title,x,3.0,3.05,0.42,{fontSize:13,bold:true,color:C.text,align:'center'});
    txt(s,sc.desc,x,3.45,3.05,0.62,{fontSize:11,color:C.textS,align:'center',wrap:true});

    if(i<3){
      txt(s,'→',x+3.05,2.85,0.2,0.4,{fontSize:18,color:C.gray4,bold:true,align:'center'});
    }
  });

  // 底部
  rnd(s,0.5,5.02,12.33,0.45,C.redB,'fca5a5',0.08);
  txt(s,'❌  這樣的會議，決策依據是落後的資料，主管無法即時掌握真實業務狀況',0.5,5.02,12.33,0.45,{fontSize:12,bold:true,color:C.red,align:'center',valign:'middle'});
}

// ══════════════════════════════════════════════════════════
// S3 解方：主管即時掌握
// ══════════════════════════════════════════════════════════
{
  const s=pptx.addSlide(); bg(s,C.white);
  hdr(s,'有了 ITTS CRM，主管這樣開會');

  // 左：舊方式
  rnd(s,0.42,1.28,5.6,4.1,'fef2f2','fca5a5',0.12);
  txt(s,'❌  以前',0.6,1.38,5.2,0.45,{fontSize:16,bold:true,color:C.red});
  const olds=[
    {e:'📧',t:'等業務寄Excel'},
    {e:'🔀',t:'手動合併多個版本'},
    {e:'🕐',t:'整理花2小時'},
    {e:'📉',t:'開會看的是昨天的數字'},
  ];
  olds.forEach((o,i)=>{
    txt(s,o.e,0.6,1.98+i*0.72,0.75,0.65,{fontSize:32,align:'center'});
    txt(s,o.t,1.38,2.08+i*0.72,4.4,0.45,{fontSize:14,color:C.textS,valign:'middle'});
  });

  // 中間箭頭
  s.addShape(pptx.ShapeType.ellipse,{x:6.05,y:2.8,w:1.22,h:1.22,fill:{color:C.blue},line:{color:C.blue,width:0}});
  txt(s,'→',6.05,2.8,1.22,1.22,{fontSize:32,bold:true,color:C.white,align:'center',valign:'middle'});

  // 右：新方式
  rnd(s,7.3,1.28,5.6,4.1,C.greenB,'86efac',0.12);
  txt(s,'✅  現在',7.5,1.38,5.2,0.45,{fontSize:16,bold:true,color:C.green});
  const news=[
    {e:'📱',t:'打開手機即時查看'},
    {e:'📊',t:'所有業務數據自動匯總'},
    {e:'⚡',t:'0秒整理，永遠是最新'},
    {e:'🎯',t:'週會直接投螢幕討論'},
  ];
  news.forEach((n,i)=>{
    txt(s,n.e,7.5,1.98+i*0.72,0.75,0.65,{fontSize:32,align:'center'});
    txt(s,n.t,8.28,2.08+i*0.72,4.4,0.45,{fontSize:14,color:C.textS,valign:'middle'});
  });
}

// ══════════════════════════════════════════════════════════
// S4 主管儀表板（即時數據示意）
// ══════════════════════════════════════════════════════════
{
  const s=pptx.addSlide(); bg(s,C.white);
  hdr(s,'主管儀表板　—　週會直接看這頁','所有業務數據即時匯總，無需任何整理');

  // 上排 KPI 卡
  const kpis=[
    {e:'💰',label:'本月已成交',val:'$1,240萬',sub:'↑ 18% vs 上月',color:C.green,bg:C.greenB},
    {e:'📋',label:'進行中商機',val:'23 筆',sub:'總金額 $4,580萬',color:C.blue,bg:C.blueBg},
    {e:'👥',label:'本月新增聯絡人',val:'47 位',sub:'AI辨識建檔 31 位',color:C.purple,bg:'faf5ff'},
    {e:'🏃',label:'本月拜訪次數',val:'89 次',sub:'平均每業務 8.1 次',color:C.orange,bg:'fff7ed'},
  ];
  kpis.forEach((k,i)=>{
    const x=0.42+i*3.25;
    rnd(s,x,1.28,3.05,1.45,k.bg,C.gray3,0.1);
    txt(s,k.e,x,1.32,0.75,0.75,{fontSize:32,align:'center',valign:'middle'});
    txt(s,k.label,x+0.75,1.35,2.15,0.3,{fontSize:10,color:C.gray4});
    txt(s,k.val,x+0.75,1.62,2.15,0.52,{fontSize:22,bold:true,color:k.color});
    txt(s,k.sub,x+0.15,2.42,2.75,0.25,{fontSize:10,color:C.gray4});
  });

  // 中段：業務排行 + 商機分布
  // 業務排行
  rnd(s,0.42,2.88,6.0,2.35,C.white,C.gray3,0.1);
  txt(s,'📊  各業務本月業績',0.62,2.98,5.6,0.38,{fontSize:13,bold:true,color:C.navy});

  const members=[
    {name:'陳業務',val:340,pct:92,color:C.green},
    {name:'林業務',val:280,pct:76,color:C.blue},
    {name:'王業務',val:210,pct:57,color:C.blue},
    {name:'張業務',val:180,pct:49,color:C.orange},
    {name:'李業務',val:100,pct:27,color:C.red},
  ];
  members.forEach((m,i)=>{
    const y=3.48+i*0.34;
    txt(s,m.name,0.55,y,1.1,0.3,{fontSize:10,color:C.textS,valign:'middle'});
    s.addShape(pptx.ShapeType.rect,{x:1.62,y:y+0.04,w:4.3,h:0.22,fill:{color:C.gray2},line:{color:C.gray3,width:0}});
    s.addShape(pptx.ShapeType.rect,{x:1.62,y:y+0.04,w:4.3*m.pct/100,h:0.22,fill:{color:m.color,transparency:20},line:{color:m.color,width:0}});
    txt(s,'$'+m.val+'萬',5.98,y,0.85,0.3,{fontSize:10,bold:true,color:m.color,align:'right',valign:'middle'});
  });

  // 商機分布
  rnd(s,6.6,2.88,6.12,2.35,C.white,C.gray3,0.1);
  txt(s,'🎯  商機階段分布',6.8,2.98,5.7,0.38,{fontSize:13,bold:true,color:C.navy});

  const stages=[
    {l:'A 探索',n:8,color:'94a3b8'},
    {l:'B 需求',n:6,color:C.blue},
    {l:'C 報價',n:5,color:C.purple},
    {l:'D 議約',n:3,color:C.orange},
    {l:'Won ✓',n:1,color:C.green},
  ];
  stages.forEach((st,i)=>{
    const barW=st.n*0.42;
    const y=3.42+i*0.35;
    txt(s,st.l,6.72,y,1.1,0.3,{fontSize:10,color:C.textS,valign:'middle'});
    s.addShape(pptx.ShapeType.roundRect,{x:7.82,y:y+0.04,w:barW,h:0.22,fill:{color:st.color,transparency:15},line:{color:st.color,width:0},rectRadius:0.05});
    txt(s,st.n+'筆',7.82+barW+0.08,y,0.5,0.3,{fontSize:10,bold:true,color:st.color,valign:'middle'});
  });

  rnd(s,0.42,5.3,12.3,0.18,C.blueBg,'bee3f8',0.05);
  txt(s,'✦ 所有數據即時更新，主管打開頁面即是最新狀態，無需任何人工整理',0.42,5.3,12.3,0.18,{fontSize:10,color:C.blue,align:'center',valign:'middle'});
}

// ══════════════════════════════════════════════════════════
// S5 業務省心（大圖示）
// ══════════════════════════════════════════════════════════
{
  const s=pptx.addSlide(); bg(s,C.gray1);
  hdr(s,'業務的一天　—　省心又省時','AI 幫你做繁瑣的事，你專注在客戶身上');

  const scenes=[
    {e:'📷',time:'09:00',title:'早會拿到名片',action:'手機拍照\n10秒 AI 建檔完成',color:C.blue},
    {e:'☎️',time:'11:30',title:'電訪客戶',action:'記錄拜訪\nAI 自動給下一步建議',color:C.purple},
    {e:'📊',time:'14:00',title:'商機推進',action:'拖曳看板\n階段立即更新通知主管',color:C.orange},
    {e:'🏆',time:'17:00',title:'成交簽約',action:'拖到 Won\n業績即時計入 + 慶祝',color:C.green},
  ];

  scenes.forEach((sc,i)=>{
    const x=0.42+i*3.25;
    rnd(s,x,1.28,3.05,4.1,C.white,C.gray3,0.12);
    s.addShape(pptx.ShapeType.rect,{x,y:1.28,w:3.05,h:0.05,fill:{color:sc.color},line:{color:sc.color,width:0}});

    // 時間徽章
    rnd(s,x+0.62,1.4,1.8,0.3,sc.color,sc.color,0.15);
    txt(s,sc.time,x+0.62,1.4,1.8,0.3,{fontSize:10,bold:true,color:sc.color,align:'center',valign:'middle'});

    // 大圖示
    s.addShape(pptx.ShapeType.ellipse,{x:x+0.78,y:1.88,w:1.5,h:1.5,fill:{color:sc.color,transparency:88},line:{color:sc.color,width:1.5}});
    txt(s,sc.e,x+0.78,1.88,1.5,1.5,{fontSize:46,align:'center',valign:'middle'});

    txt(s,sc.title,x,3.55,3.05,0.4,{fontSize:13,bold:true,color:C.text,align:'center'});
    hln(s,4.0,x+0.3,2.45,C.gray3);
    txt(s,sc.action,x,4.08,3.05,0.72,{fontSize:12,color:C.textS,align:'center',wrap:true});
  });
}

// ══════════════════════════════════════════════════════════
// S6 AI 四大功能（大圖示版）
// ══════════════════════════════════════════════════════════
{
  const s=pptx.addSlide(); bg(s,C.white);
  hdr(s,'🤖  AI 四大功能','Google Gemini 驅動　免費方案即可使用　無需信用卡');

  const ais=[
    {e:'📷',title:'拍照辨識名片',tag:'名片 → 10秒建檔',desc:'手機對準名片拍照，AI 自動填入姓名、公司、電話、Email 等所有欄位',color:C.blue},
    {e:'✍️',title:'拜訪記錄助手',tag:'省下整理時間',desc:'填完拜訪內容，AI 自動摘要關鍵重點並建議下一步跟進行動',color:C.purple},
    {e:'🎯',title:'商機贏率預測',tag:'AI 客觀評估',desc:'綜合商機階段、拜訪頻率、時程分析，AI 給出百分比勝率',color:C.orange},
    {e:'📄',title:'客戶輪廓摘要',tag:'交接不怕斷層',desc:'一鍵生成 150 字客戶關係摘要，健康度評級（良好/普通/需關注）',color:C.green},
  ];

  ais.forEach((ai,i)=>{
    const col=i%2, row=Math.floor(i/2);
    const x=0.42+col*6.5, y=1.32+row*2.05;

    rnd(s,x,y,6.2,1.85,C.white,C.gray3,0.12);
    s.addShape(pptx.ShapeType.rect,{x,y,w:6.2,h:0.05,fill:{color:ai.color},line:{color:ai.color,width:0}});

    // 大圖示圓
    s.addShape(pptx.ShapeType.ellipse,{x:x+0.18,y:y+0.14,w:1.35,h:1.35,fill:{color:ai.color,transparency:88},line:{color:ai.color,width:1.5}});
    txt(s,ai.e,x+0.18,y+0.14,1.35,1.35,{fontSize:44,align:'center',valign:'middle'});

    // Tag
    rnd(s,x+1.72,y+0.18,1.9,0.28,ai.color,ai.color,0.14);
    txt(s,ai.tag,x+1.72,y+0.18,1.9,0.28,{fontSize:9,bold:true,color:ai.color,align:'center',valign:'middle'});

    txt(s,ai.title,x+1.72,y+0.55,4.3,0.38,{fontSize:15,bold:true,color:C.navy});
    txt(s,ai.desc,x+1.72,y+0.98,4.3,0.65,{fontSize:11,color:C.textS,wrap:true});
  });

  rnd(s,0.42,5.48,12.3,0.18,C.blueBg,'bee3f8',0.05);
  txt(s,'免費額度：每天 1,500 次請求　|　15 次/分鐘　|　業務每日正常使用約佔 4% 額度',0.42,5.48,12.3,0.18,{fontSize:10,color:C.blue,align:'center',valign:'middle'});
}

// ══════════════════════════════════════════════════════════
// S7 商機看板（視覺示意）
// ══════════════════════════════════════════════════════════
{
  const s=pptx.addSlide(); bg(s,C.white);
  hdr(s,'商機看板　—　拖曳即更新','主管一眼掌握全公司商機進度，業務不用另外回報');

  const stages=[
    {l:'A　探索接觸',c:'94a3b8',cards:[{t:'三晃實業',a:'$85萬'},{t:'台灣水泥',a:'$120萬'}]},
    {l:'B　需求確認',c:C.blue,cards:[{t:'永聯科技',a:'$280萬'},{t:'長庚資訊',a:'$160萬'},{t:'遠東新',a:'$95萬'}]},
    {l:'C　方案報價',c:C.purple,cards:[{t:'台塑集團',a:'$450萬'}]},
    {l:'D　議約收尾',c:C.orange,cards:[{t:'中鋼雲端',a:'$320萬'},{t:'友達光電',a:'$180萬'}]},
    {l:'🏆 成交',c:C.green,cards:[{t:'中華電信',a:'$240萬'}]},
  ];

  stages.forEach((st,i)=>{
    const x=0.38+i*2.57;
    rnd(s,x,1.28,2.38,0.38,st.c,'ffffff',0.06);
    txt(s,st.l,x,1.28,2.38,0.38,{fontSize:9,bold:true,color:C.white,align:'center',valign:'middle'});

    st.cards.forEach((c2,j)=>{
      rnd(s,x+0.06,1.76+j*1.05,2.26,0.88,C.white,C.gray3,0.08);
      s.addShape(pptx.ShapeType.rect,{x:x+0.06,y:1.76+j*1.05,w:0.05,h:0.88,fill:{color:st.c},line:{color:st.c,width:0}});
      txt(s,c2.t,x+0.16,1.82+j*1.05,2.05,0.32,{fontSize:11,bold:true,color:C.text,wrap:true});
      txt(s,c2.a,x+0.16,2.12+j*1.05,2.05,0.28,{fontSize:12,bold:true,color:st.c});
      txt(s,'● 拜訪中',x+0.16,2.38+j*1.05,2.05,0.2,{fontSize:9,color:C.gray4});
    });
  });

  // 底部提示
  const hints=[
    {e:'🖱️',t:'拖曳卡片即更新階段'},
    {e:'🔔',t:'成員異動主管即時可見'},
    {e:'💰',t:'Won 自動計入已成交金額'},
  ];
  hints.forEach((h,i)=>{
    rnd(s,0.42+i*4.3,5.02,4.1,0.48,C.blueBg,C.gray3,0.08);
    txt(s,h.e+'  '+h.t,0.42+i*4.3,5.02,4.1,0.48,{fontSize:12,color:C.blue,align:'center',valign:'middle'});
  });
}

// ══════════════════════════════════════════════════════════
// S8 週會場景 before/after
// ══════════════════════════════════════════════════════════
{
  const s=pptx.addSlide(); bg(s,C.white);
  hdr(s,'週一業務會議　—　這樣就夠了','不再等 Excel，打開系統就開會');

  // 左：舊
  rnd(s,0.42,1.28,5.8,4.1,'fef2f2','fca5a5',0.12);
  txt(s,'😩  過去的週一',0.62,1.38,5.4,0.45,{fontSize:15,bold:true,color:C.red});
  const befores=[
    {e:'📧',t:'秘書群發索取 Excel'},
    {e:'⏳',t:'等待業務回傳 2~3 小時'},
    {e:'🔀',t:'手動合併5個不同版本'},
    {e:'🤔',t:'主管看著舊數據做決策'},
    {e:'🔁',t:'會後再更新一版流串'},
  ];
  befores.forEach((b,i)=>{
    txt(s,b.e,0.62,2.0+i*0.58,0.65,0.52,{fontSize:26,align:'center'});
    txt(s,b.t,1.3,2.1+i*0.58,4.6,0.35,{fontSize:13,color:C.textS,valign:'middle'});
  });

  // 箭頭
  s.addShape(pptx.ShapeType.ellipse,{x:6.12,y:2.88,w:1.1,h:1.1,fill:{color:C.blue},line:{color:C.blue,width:0}});
  txt(s,'→',6.12,2.88,1.1,1.1,{fontSize:30,bold:true,color:C.white,align:'center',valign:'middle'});

  // 右：新
  rnd(s,7.42,1.28,5.5,4.1,C.greenB,'86efac',0.12);
  txt(s,'😊  現在的週一',7.62,1.38,5.1,0.45,{fontSize:15,bold:true,color:C.green});
  const afters=[
    {e:'📱',t:'打開 CRM 儀表板'},
    {e:'⚡',t:'所有數據自動匯總完畢'},
    {e:'🖥️',t:'投影直接討論，無需準備'},
    {e:'📊',t:'即時更新，決策有依據'},
    {e:'✅',t:'會後數據繼續跑，不用整理'},
  ];
  afters.forEach((a,i)=>{
    txt(s,a.e,7.62,2.0+i*0.58,0.65,0.52,{fontSize:26,align:'center'});
    txt(s,a.t,8.3,2.1+i*0.58,4.5,0.35,{fontSize:13,color:C.textS,valign:'middle'});
  });
}

// ══════════════════════════════════════════════════════════
// S9 導入效益（大數字）
// ══════════════════════════════════════════════════════════
{
  const s=pptx.addSlide(); bg(s,C.navy);

  s.addShape(pptx.ShapeType.rect,{x:0,y:5.1,w:13.33,h:0.53,fill:{color:C.blue},line:{color:C.blue,width:0}});
  txt(s,'itts-crm.vercel.app　|　ITTS 東捷資訊服務',0,5.1,13.33,0.53,{fontSize:11,color:'cce4ff',align:'center',valign:'middle'});

  txt(s,'導入後，主管與業務都得到了什麼？',0.5,0.22,12.33,0.58,{fontSize:20,bold:true,color:'cce4ff'});
  hln(s,0.88,0.5,12.33,'2d5a8e');

  // 上排 3 大數字
  const top3=[
    {e:'⏱️',num:'80%',label:'業務建檔時間節省',sub:'AI拍照10秒完成'},
    {e:'📡',num:'即時',label:'主管掌握業務狀況',sub:'打開即是最新數據'},
    {e:'❌',num:'0',label:'Excel 版本困擾',sub:'唯一平台唯一版本'},
  ];
  top3.forEach((t,i)=>{
    const x=0.5+i*4.3;
    rnd(s,x,0.98,3.95,2.2,'162d4a','2d5a8e',0.12);
    txt(s,t.e,x,1.08,3.95,0.72,{fontSize:36,align:'center'});
    txt(s,t.num,x,1.78,3.95,0.72,{fontSize:40,bold:true,color:C.white,align:'center'});
    txt(s,t.label,x,2.46,3.95,0.35,{fontSize:12,bold:true,color:'cce4ff',align:'center'});
    txt(s,t.sub,x,2.79,3.95,0.28,{fontSize:10,color:'8db8e8',align:'center'});
  });

  // 下排 4 個次要
  const bot4=[
    {e:'🤖',v:'4項',l:'AI免費功能'},
    {e:'📱',v:'0',l:'額外APP安裝'},
    {e:'☁️',v:'24/7',l:'雲端隨時存取'},
    {e:'🔒',v:'3層',l:'角色權限管控'},
  ];
  bot4.forEach((b,i)=>{
    const x=0.5+i*3.25;
    rnd(s,x,3.4,3.05,1.48,'162d4a','2d5a8e',0.1);
    txt(s,b.e,x,3.48,3.05,0.5,{fontSize:26,align:'center'});
    txt(s,b.v,x,3.96,3.05,0.42,{fontSize:26,bold:true,color:C.white,align:'center'});
    txt(s,b.l,x,4.36,3.05,0.3,{fontSize:11,color:'8db8e8',align:'center'});
  });
}

// ── 輸出 ─────────────────────────────────────────────────
pptx.writeFile({fileName:'C:/Users/steven.lee/Documents/ITTS_CRM_主管版v3.pptx'})
  .then(()=>console.log('✅ 完成：ITTS_CRM_主管版v3.pptx'))
  .catch(e=>console.error('❌',e));
