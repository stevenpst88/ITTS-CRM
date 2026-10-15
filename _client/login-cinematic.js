/* 登入頁「電影感開場」（ITTS-CRM）。外部檔（無 inline script），原生 WebGL＋Canvas2D，無外部依賴。
   分級：full（WebGL 粒子）→ lite（Canvas2D 星塵）→ static（無動畫；減少動態效果或太慢時）。 */
(function(){
/* ── 信件深層連結：登入頁載入時決定要不要沿用暫存的 #quote:…（登入成功後帶回去）──
   網址上有合法片段 → 記住；沒有，但 app.js 是為了登入而導過來的（上膛）→ 保留；其他（使用者主動打開登入頁、登出後回來）→ 清掉舊暫存，
   下一位登入的人不會被帶去上一位的單據。deep-link.js 載入失敗只是少了這個功能，不影響登入。 */
try { if (window.ITTSDeepLink) window.ITTSDeepLink.settleOnLoginPage(window.location.hash); } catch (_) { /* 取不到就當沒有 */ }

/* ── Service Worker 註冊與更新提示（與舊登入頁相同；提示列改用 class，不寫 inline style） ── */
if ('serviceWorker' in navigator) {
  addEventListener('load', function () {
    navigator.serviceWorker.register('/sw.js').then(function (reg) {
      reg.addEventListener('updatefound', function () {
        var nw = reg.installing; if (!nw) return;
        nw.addEventListener('statechange', function () {
          if (nw.state === 'installed' && navigator.serviceWorker.controller) {
            window.__swUpdateReg = reg;
            if (typeof window.showSwUpdateToast === 'function') { window.showSwUpdateToast(reg); return; }
            var bar = document.createElement('div'); bar.className = 'sw-bar';
            var sp = document.createElement('span'); sp.textContent = '\u{1F504} 有新版本可用';
            var b = document.createElement('button'); b.type = 'button'; b.textContent = '立即更新';
            b.addEventListener('click', function () { if (reg.waiting) reg.waiting.postMessage({ type: 'SKIP_WAITING' }); });
            bar.appendChild(sp); bar.appendChild(b); document.body.appendChild(bar);
          }
        });
      });
    }).catch(function (err) { console.warn('[SW] register failed:', err); });
    /* 第一次安裝 SW（載入時還沒有 controller）時 clients.claim() 也會觸發 controllerchange；那時重新整理只會把使用者剛打的字和開場打斷，所以只在「換新版」時才 reload；登入中／成功轉場中也不 reload */
    var hadController = !!navigator.serviceWorker.controller, refreshing = false;
    navigator.serviceWorker.addEventListener('controllerchange', function () {
      if (!hadController || refreshing) return;
      try { if (ITTS.state === 'loading' || ITTS.state === 'success') return; } catch (_) {}
      refreshing = true; window.location.reload();
    });
  });
}
/* 從 bfcache 還原（登入成功後按上一頁）時，頁面會停在「成功轉場」畫面 → 重新載入 */
addEventListener('pageshow', function (e) { if (e.persisted) location.reload(); });

/* ═══ 共用：參數／主題／表單行為／分級降級工具 ═══ */
const $=s=>document.querySelector(s);
const Q=new URLSearchParams(location.search);
const ROOT=document.documentElement;
const BOOT=window.ITTSLoginBoot,THEMES=BOOT.THEMES;
const hex=h=>[1,3,5].map(i=>parseInt(h.slice(i,i+2),16)/255);
const mixc=(a,b,t)=>a.map((v,i)=>v+(b[i]-v)*t);
const rgbCss=(c,a)=>'rgba('+c.map(v=>Math.round(v*255)).join(',')+','+(a===undefined?1:a)+')';
const hexCss=c=>'#'+c.map(v=>Math.round(v*255).toString(16).padStart(2,'0')).join('');
const clamp=(v,a,b)=>Math.max(a,Math.min(b,v));
let THEME=BOOT.theme;
let MODE=BOOT.mode;                                         /* 預設深色、不看 OS；偏好記在 localStorage（glassTheme／darkMode，與玻璃 UI 共用） */
let MOTION_REDUCE=false;                                    /* 尊重系統「減少動態效果」；?motion=reduce 亦可手動關閉 */
try{MOTION_REDUCE=Q.get('motion')==='reduce'||!!(window.matchMedia&&matchMedia('(prefers-reduced-motion: reduce)').matches)}catch(e){}
if(MOTION_REDUCE)ROOT.classList.add('rm');

/* 調色盤（供 shader／canvas 使用）— 每個主題×模式各一組，切換時 JS 內插 */
function paletteOf(theme,mode){
  const T=THEMES[theme],bg=hex(T.bg),b1=hex(T.b1),b2=hex(T.b2),b3=hex(T.b3),ac=hex(T.ac),acl=hex(T.acl);
  if(mode==='dark')return{bg,c0:b1,c1:ac,c2:b2,c3:mixc(b3,ac,.4),ac,btn:ac,app:mixc(bg,b1,.15),tint:mixc(bg,b1,.10),s:[.56,.74,0]};
  const lb=mixc(hex('#dfe9ff'),b1,T.lmix*.18);
  return{bg:lb,c0:mixc(b1,[0,0,0],.15),c1:acl,c2:mixc(b2,[0,0,0],.1),c3:mixc(b3,b1,.5),ac:acl,btn:acl,app:mixc(lb,[1,1,1],.35),tint:[1,1,1],s:[.72,1,1]};
}
let PAL=paletteOf(THEME,MODE),PAL_T=PAL;
function palStep(dt){            /* 指數內插到目標 */
  const k=1-Math.exp(-dt*5);let moved=false;
  for(const key in PAL_T){const a=PAL[key],b=PAL_T[key];PAL[key]=a.map((v,i)=>{const d=b[i]-v;if(Math.abs(d)>.002)moved=true;return v+d*k});}
  return moved;
}
function applyTheme(initial){
  BOOT.paint(ROOT,THEME,MODE);
  const p=paletteOf(THEME,MODE);
  ROOT.style.setProperty('--app',hexCss(p.app));
  PAL_T=p;if(initial)PAL=paletteOf(THEME,MODE);
  document.querySelectorAll('#dock [data-th]').forEach(b=>b.setAttribute('aria-pressed',b.dataset.th===THEME?'true':'false'));
  if(ITTS.onTheme)ITTS.onTheme();
}
function buildDock(){
  const d=$('#dock');if(!d)return;
  Object.keys(THEMES).forEach(k=>{const b=document.createElement('button');b.type='button';b.dataset.th=k;b.title=THEMES[k].n;b.setAttribute('aria-label','主題：'+THEMES[k].n);b.style.setProperty('--c','linear-gradient(135deg,'+THEMES[k].b1+','+THEMES[k].ac+')');b.style.background='linear-gradient(135deg,'+THEMES[k].b1+','+THEMES[k].ac+')';
    b.onclick=()=>{THEME=k;savePref();applyTheme()};d.appendChild(b)});
  const s=document.createElement('span');s.className='sep';d.appendChild(s);
  const m=document.createElement('button');m.type='button';m.className='mode';m.setAttribute('aria-label','切換深色／淺色');
  m.innerHTML='<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M2 12h2M20 12h2M5 5l1.4 1.4M17.6 17.6L19 19M5 19l1.4-1.4M17.6 6.4L19 5"/></svg>';
  m.onclick=()=>{MODE=MODE==='dark'?'light':'dark';savePref();applyTheme()};d.appendChild(m);
  function savePref(){try{localStorage.setItem('glassTheme',THEME);localStorage.setItem('darkMode',MODE==='dark'?'1':'0')}catch(e){}}
}

/* ── 表單行為（對應 login.html：#username／#password／#errorMsg／送出鈕；顯示密碼、大寫提示為新增）── */
const ITTS={onState:null,onTheme:null,state:'idle',tier:null};
function setState(s){
  ITTS.state=s;
  ROOT.classList.remove('st-loading','st-error','st-success');
  if(s!=='idle')ROOT.classList.add('st-'+s);
  const em=$('#errorMsg'),btn=$('#loginBtn'),lbl=$('#btnLbl');
  em.classList.toggle('show',s==='error');
  btn.disabled=(s==='loading'||s==='success');
  lbl.textContent=s==='loading'?'驗證中…':(s==='success'?'登入成功':'登　入');
  btn.classList.toggle('hot',s==='loading'||s==='success'||btn.dataset.hover==='1');
  if(ITTS.onState)ITTS.onState(s);
}
/* 被伺服器登出時（閒置逾時／連線失效）由 app.js 留下原因，這裡顯示友善提示（只顯示一次；來自 feat/audit3-batch2） */
function showLoginNotice(){
  let code='';
  try{code=sessionStorage.getItem('loginNotice')||'';sessionStorage.removeItem('loginNotice')}catch(e){return}
  const MSG={SESSION_IDLE_TIMEOUT:'閒置過久已登出，請重新登入',SESSION_EXPIRED:'登入已逾時，請重新登入',SESSION_INVALID:'您的連線已失效（帳號停用或已不存在），請重新登入'};
  const el=$('#loginNotice');
  if(el&&MSG[code]){el.textContent=MSG[code];el.classList.add('show')}
}
function initForm(){
  const u=$('#username'),p=$('#password'),eye=$('#eyeBtn'),caps=$('#capsHint'),btn=$('#loginBtn'),form=$('#loginForm');
  eye.onclick=()=>{const show=p.type==='password';p.type=show?'text':'password';eye.setAttribute('aria-pressed',show);eye.setAttribute('aria-label',show?'隱藏密碼':'顯示密碼')};
  const capf=e=>{if(e.getModifierState)caps.hidden=!e.getModifierState('CapsLock')};
  p.addEventListener('keydown',capf);p.addEventListener('keyup',capf);p.addEventListener('blur',()=>caps.hidden=true);
  const hot=on=>{btn.dataset.hover=on?'1':'0';btn.classList.toggle('hot',on||ITTS.state==='loading'||ITTS.state==='success')};
  btn.addEventListener('pointerenter',()=>hot(true));btn.addEventListener('pointerleave',()=>hot(document.activeElement===btn));
  btn.addEventListener('focus',()=>hot(true));btn.addEventListener('blur',()=>hot(false));
  [u,p].forEach(i=>i.addEventListener('input',()=>{if(ITTS.state==='error')setState('idle')}));
  let busy=false;
  const showErr=(msg,clearPw)=>{
    busy=false;$('#errorMsg').textContent=msg;setState('error');
    if(clearPw){p.value=''}
    try{p.focus()}catch(_){}
  };
  function onLoginOk(data){
    try{sessionStorage.setItem('justLoggedIn','1')}catch(_){}
    /* 深層連結：非 admin 登入後帶回暫存的 #quote:…（hashForRedirect 取一次即清）；admin 不附片段，但仍要清掉暫存 */
    let dl='';try{dl=window.ITTSDeepLink?window.ITTSDeepLink.hashForRedirect():''}catch(_){dl=''}
    const target=data.role==='admin'?'/admin.html':'/'+dl;
    setState('success');
    let gone=false;const go=()=>{if(gone)return;gone=true;location.href=target};
    /* 動畫約 2.2 秒＋收尾；靜態版（減少動態／降級）縮短；4 秒必定導向（安全逾時） */
    setTimeout(go,ITTS.tier==='static'?900:2500);
    setTimeout(go,4000);
  }
  form.addEventListener('submit',async e=>{
    e.preventDefault();
    if(busy||ITTS.state==='loading'||ITTS.state==='success')return;
    if(!form.checkValidity()){form.reportValidity();return}
    busy=true;setState('loading');
    const username=u.value.trim(),password=p.value;
    const ctl=window.AbortController?new AbortController():null;
    const tm=setTimeout(()=>{if(ctl)ctl.abort()},20000);
    try{
      const res=await fetch('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({username,password}),signal:ctl?ctl.signal:undefined});
      let data=null;try{data=await res.json()}catch(_){}
      clearTimeout(tm);
      if(res.ok&&data&&data.success){onLoginOk(data);return}
      if(res.status===401)showErr('帳號或密碼錯誤，請重新輸入',true);
      else showErr((data&&data.message)||'登入失敗，請稍後再試',false);
    }catch(err){clearTimeout(tm);showErr('連線失敗，請稍後再試',false)}
  });
}

/* ── 分級／效能工具 ── */
function webglSupported(){
  try{const c=document.createElement('canvas');const g=c.getContext('webgl',{failIfMajorPerformanceCaveat:false})||c.getContext('experimental-webgl');
    if(!g)return false;const ok=!!g.getParameter(g.VERSION);const e=g.getExtension('WEBGL_lose_context');if(e)e.loseContext();return ok}catch(e){return false}
}
/* 主迴圈：頁籤隱藏即停（回來後 dt 不會爆衝）。cb(time, dt) 回傳 false 可停止 */
function loop(cb){
  let last=0,id=0,clock=0,run=true;
  const f=ts=>{if(!run)return;const dt=last?Math.min((ts-last)/1000,.1):.016;last=ts;clock+=dt;if(cb(clock,dt,ts)===false){run=false;return}id=requestAnimationFrame(f)};
  id=requestAnimationFrame(f);
  document.addEventListener('visibilitychange',()=>{if(!run)return;if(document.hidden){cancelAnimationFrame(id);id=0}else if(!id||true){cancelAnimationFrame(id);last=0;id=requestAnimationFrame(f)}});
  return{stop(){run=false;cancelAnimationFrame(id)},get clock(){return clock}};
}
/* 量測前 N 幀的牆鐘時間（略過前 4 幀暖機），回傳平均 ms；用來決定要不要降級 */
function makePerf(n,onDone){
  let c=0,t0=0,done=false,sum=0,k=0;
  return ts=>{if(done)return;c++;if(c<=4){t0=ts;return}const d=ts-t0;t0=ts;if(d>250)return;sum+=d;k++;if(k>=n){done=true;onDone(sum/k)}}
}
const pointer={x:innerWidth/2,y:innerHeight/2,nx:0,ny:0,moved:false,down:false,t:0};
addEventListener('pointermove',e=>{pointer.x=e.clientX;pointer.y=e.clientY;pointer.nx=e.clientX/innerWidth*2-1;pointer.ny=e.clientY/innerHeight*2-1;pointer.moved=true;pointer.t=performance.now();
  const c=$('#card');if(c){const r=c.getBoundingClientRect();c.style.setProperty('--mx',((e.clientX-r.left)/r.width*100).toFixed(1)+'%');c.style.setProperty('--my',((e.clientY-r.top)/r.height*100).toFixed(1)+'%')}},{passive:true});
addEventListener('pointerdown',()=>{pointer.down=true});addEventListener('pointerup',()=>{pointer.down=false});

/* ── GLSL：粒子（點＋速度線共用頂點著色器）、模糊、合成 ── */
const PREC='#ifdef GL_FRAGMENT_PRECISION_HIGH\nprecision highp float;\n#else\nprecision mediump float;\n#endif\n';
const VS_P=`
attribute vec3 aP;attribute vec4 aR;attribute float aS;
uniform vec2 uCss,uCtr,uPtr,uPar;
uniform float uT,uAsm,uCam,uScale,uWarp,uBurst,uShim,uPtrOn,uDpr,uTrail,uDustA,uType,uLogoA;
uniform vec3 uC0,uC1;
varying vec4 vC;varying float vE;
const float F=1000.;
void main(){
  float wv=max(uWarp-aS*uTrail,0.);
  vec3 pos;float size=1.;vec4 col;
  if(uType<.5){
    float k=clamp(uAsm*1.55-aR.x*.55,0.,1.);float e=1.-pow(1.-k,3.5);
    vec3 tg=vec3(aP.xy*uScale+uCtr,aP.z*26.);
    float ang=aR.z*6.28318;float rad=380.+1000.*aR.w;
    vec3 st=vec3(cos(ang)*rad*1.25,sin(ang)*rad*.78,260.+1100.*aR.y);
    pos=mix(st,tg,e);
    pos.xy+=vec2(-sin(ang),cos(ang))*sin(e*3.14159)*240.*(.25+aR.y);
    float sh=uShim*smoothstep(.8,1.,e);
    pos.xy+=vec2(sin(uT*1.2+aR.z*50.),cos(uT*.95+aR.w*43.))*(1.1+aR.y*1.7)*sh;
    pos.z+=sin(uT*.8+aR.y*31.)*9.*sh;
    vec2 dp=pos.xy-uPtr;float dd=length(dp);
    float f=exp(-dd*dd/(2.*105.*105.))*uPtrOn*smoothstep(.8,1.,e);
    pos.xy+=dp/(dd+1.)*f*48.*(.6+aR.y*.8);pos.z-=f*70.;
    vec2 rel=pos.xy-uCtr;vec2 dir=normalize(rel+(aR.zw-.5)*40.+vec2(.001));
    float bb=1.-pow(1.-clamp(uBurst-aS*uTrail*.4,0.,1.),3.);
    pos.xy+=dir*bb*(240.+900.*aR.y);pos.z+=bb*(aR.w-.5)*500.;
    size=(1.0+aR.y*1.5)*(1.+(1.-e)*.5+f*.7);
    float tw=.78+.22*sin(uT*3.+aR.z*30.);
    float a=mix(.22,.62,e)*tw*uLogoA*(1.+(1.-e)*.9+f*.8);
    col=vec4(mix(uC0,uC1,aR.w*aR.w+f*.6),a);
  }else{
    float z=mod(aP.z-uT*(16.+34.*aR.y)+300.,1700.)-300.;
    pos=vec3(aP.xy,z);
    float tw=.65+.35*sin(uT*(1.+aR.x*3.)+aR.z*40.);
    float a=(.18+.82*aR.w*aR.w)*tw*uDustA*smoothstep(-300.,-120.,z)*smoothstep(1400.,1050.,z);
    size=.9+aR.y*1.9;
    col=vec4(mix(vec3(.74,.84,1.),uC1,aR.z*.5)*(uC0.x<-.5?1.:1.),a);
    col.rgb=mix(uC0,uC1,aR.z*.55)*.85+vec3(.08);
  }
  pos.z-=wv*wv*(2400.+1700.*aR.y);
  float dist=F+pos.z-uCam;
  if(dist<40.){gl_Position=vec4(3.,3.,3.,1.);gl_PointSize=0.;vC=vec4(0.);vE=0.;return;}
  float sc=F/dist;
  vec2 scr=(pos.xy-uPar)*sc+uPar*.22;
  gl_Position=vec4(scr.x/(uCss.x*.5),-scr.y/(uCss.y*.5),0.,1.);
  gl_PointSize=max(1.,size*sc*uDpr*(uType<.5?3.0:4.2));
  vC=col;vE=1.-aS;
}`;
const FS_PT=PREC+`varying vec4 vC;varying float vE;
void main(){vec2 c=gl_PointCoord-.5;float r2=dot(c,c)*4.;if(r2>1.)discard;
  float core=exp(-r2*7.)+exp(-r2*2.2)*.22;gl_FragColor=vec4(vC.rgb*core*vC.a,1.);}`;
const FS_LN=PREC+`varying vec4 vC;varying float vE;void main(){gl_FragColor=vec4(vC.rgb*vC.a*vE*.9,1.);}`;
const VS_Q='attribute vec2 a;void main(){gl_Position=vec4(a,0.,1.);}';
const FS_BLUR=PREC+`uniform sampler2D uTex;uniform vec2 uStep,uOut;uniform float uThr;
vec3 tx(vec2 uv){vec3 c=texture2D(uTex,uv).rgb;if(uThr>0.){float l=max(c.r,max(c.g,c.b));c*=smoothstep(uThr,uThr+.25,l);}return c;}
void main(){vec2 uv=gl_FragCoord.xy/uOut;
  vec3 s=tx(uv)*.2270;
  s+=(tx(uv+uStep)+tx(uv-uStep))*.1946;
  s+=(tx(uv+uStep*2.)+tx(uv-uStep*2.))*.1216;
  s+=(tx(uv+uStep*3.)+tx(uv-uStep*3.))*.0541;
  s+=(tx(uv+uStep*4.)+tx(uv-uStep*4.))*.0162;
  gl_FragColor=vec4(s,1.);}`;
const FS_COMP=PREC+`uniform sampler2D uA,uB1,uB2,uAn;uniform vec2 uRes;
uniform float uT,uBloom,uFade,uLight,uWarp,uAn1;uniform vec3 uBg,uN1,uN2,uApp;
float hash(vec2 p){return fract(sin(dot(p,vec2(12.9898,78.233)))*43758.5453);}
void main(){
  vec2 uv=gl_FragCoord.xy/uRes;vec2 q=uv-.5;float asp=uRes.x/uRes.y;vec2 qa=vec2(q.x*asp,q.y);
  float vig=smoothstep(1.05,.28,length(qa));
  vec2 ca=q*(.0012+uWarp*.016);
  vec3 a=vec3(texture2D(uA,uv+ca).r,texture2D(uA,uv).g,texture2D(uA,uv-ca).b);
  vec3 bl=texture2D(uB1,uv).rgb*.6+texture2D(uB2,uv).rgb*.75+texture2D(uAn,uv).rgb*uAn1;
  vec3 nb=uBg+uN1*.20*exp(-dot(qa+vec2(.18,.0),qa+vec2(.18,.0))*5.5)+uN2*.14*exp(-dot(qa-vec2(-.35,.28),qa-vec2(-.35,.28))*4.5);
  vec3 col;
  if(uLight<.5){
    col=nb+a+bl*uBloom;
    col*=mix(.55,1.,vig);
    col=col/(1.+col*.12);
  }else{
    col=nb-(a*.9+bl*uBloom*.5);
    col=mix(col*.9,col,vig);
    col=clamp(col,0.,1.);
  }
  col=mix(col,uApp,uFade);
  col+=(hash(gl_FragCoord.xy+fract(uT))-.5)/255.*1.5;
  gl_FragColor=vec4(col,1.);
}`;

/* ═══ 電影感開場：主程式 ═══ */
(function(){
showLoginNotice();initForm();buildDock();applyTheme(true);
const cvGL=$('#gl'),cvLite=$('#lite'),card=$('#card'),stage=$('.stage'),tagEl=$('#tagline'),subEl=$('#tagSub'),logoSvg=$('#logoSvg'),ring=$('#ring'),flash=$('#flash'),skipEl=$('#skip'),fadeEl=$('#fade'),done=$('#done');
const hero=document.createElement('div');hero.className='hero';stage.insertBefore(hero,card);
const INTRO=3.2,TAG='讓每一次拜訪都有價值';
const sstep=(a,b,x)=>{x=clamp((x-a)/(b-a),0,1);return x*x*(3-2*x)};
const eo=x=>1-Math.pow(1-clamp(x,0,1),3);
const lerp=(a,b,t)=>a+(b-a)*t;

/* ── 開場狀態 ── */
let seen=false;/* 每次開啟登入頁都播開場（使用者決定）；reduced-motion 仍走靜態 */
let releasing=false,tier=null,why='',gl=null,G=null,frames=0,perfMs=0,mainLoop=null,lctx=null;
let T=0,Tabs=0,skipping=false,introDone=false,introSkippedBecauseSeen=false,fadeIn=1;
const staticTier=()=>tier==='static';
let playIntro=!seen;
if(MOTION_REDUCE){playIntro=false}
if(!playIntro){T=INTRO;introSkippedBecauseSeen=seen&&!MOTION_REDUCE;fadeIn=introSkippedBecauseSeen?0:1}
const succ={t0:-1,p:0,dur:2200};
const mobile=()=>innerWidth<=860;
let L=null,LOGO=null;

/* tagline 逐字 */
const letters=[...TAG].map(ch=>{const s=document.createElement('span');s.textContent=ch;s.setAttribute('aria-hidden','true');tagEl.appendChild(s);return s});
const sr=document.createElement('span');sr.className='sr';sr.textContent=TAG;tagEl.appendChild(sr);

/* ── Logo 取樣（用 Canvas2D 直接畫與 itts-logo.svg 相同的幾何，再取亮點成粒子；不載入外部圖檔，無 CSP／tainted canvas 問題）── */
function rng(seed){return()=>{seed|=0;seed=seed+0x6D2B79F5|0;let t=Math.imul(seed^seed>>>15,1|seed);t=t+Math.imul(t^t>>>7,61|t)^t;return((t^t>>>14)>>>0)/4294967296}}
function sampleLogo(N){
  const cv=document.createElement('canvas');cv.width=960;cv.height=360;const c=cv.getContext('2d',{willReadFrequently:true});
  c.setTransform(6,0,0,6,0,0);c.strokeStyle='#fff';c.fillStyle='#fff';c.lineCap='round';
  c.font='900 28px "Arial Black",Arial,sans-serif';if('letterSpacing'in c)c.letterSpacing='1px';c.fillText('ITTS',8,40);
  c.lineWidth=.8;c.strokeText('ITTS',8,40);
  const d=c.getImageData(0,0,960,360).data;const pts=[];let x0=1e9,x1=-1,y0=1e9,y1=-1;
  for(let y=0;y<360;y+=1)for(let x=0;x<960;x+=1){if(d[(y*960+x)*4+3]>110){pts.push(x,y);if(x<x0)x0=x;if(x>x1)x1=x;if(y<y0)y0=y;if(y>y1)y1=y}}
  const bw=x1-x0,bh=y1-y0,cx=(x0+x1)/2,cy=(y0+y1)/2,R=rng(20260715),out=[];
  const n=pts.length/2;
  for(let i=0;i<N;i++){const k=Math.floor(R()*n)*2;out.push((pts[k]-cx)/bw+(R()-.5)*.0016,(pts[k+1]-cy)/bw+(R()-.5)*.0016)}
  return{pts:out,hh:bh/bw,n:N};
}
function layout(){
  const W=innerWidth,H=innerHeight,mob=mobile(),hr=hero.getBoundingClientRect();
  L={W,H,mob,sx:W/2,sy:H/2-(mob?34:14),sw:mob?Math.min(W*.74,330):Math.min(W*.40,540),
     ex:hr.left+hr.width/2,ey:hr.top+hr.height/2-(mob?18:34),ew:mob?(H<=520?Math.min(hr.width*.66,170):Math.min(W*.58,236)):Math.min(hr.width*.84,430)};   /* H<=520＝手機橫放（左右並排，hero 只有一小塊） */
}

/* ── 時間軸 → DOM（所有分級共用）── */
let sc0=1,logoC={x:0,y:0,w:300};
function domTL(T){
  if(!L)layout();
  const shift=sstep(1.9,2.85,T),cam=camOf(T);
  sc0=1000/(1000-cam);
  const cx=lerp(L.sx,L.ex,shift),cy=lerp(L.sy,L.ey,shift),w=lerp(L.sw,L.ew,shift);
  logoC={x:cx,y:cy,w};
  const hh=(LOGO?LOGO.hh:.34);
  /* 以畫面中心為原點套用相機推近（與粒子投影一致） */
  const X=L.W/2+(cx-L.W/2)*sc0,Y=L.H/2+(cy-L.H/2)*sc0;
  const ty=Y+(w*hh/2)*sc0+(L.mob?30:42)*sc0;
  const sp=ITTS.state==='success'?1-sstep(0,.25,succ.p):1;
  tagEl.style.opacity=sp;subEl.style.opacity=sp*sstep(2.0,2.6,T)*.9;
  tagEl.style.transform='translate('+X.toFixed(1)+'px,'+ty.toFixed(1)+'px) translate(-50%,-50%) scale('+sc0.toFixed(4)+')';
  subEl.style.transform='translate('+X.toFixed(1)+'px,'+(ty+(L.mob?30:38)*sc0).toFixed(1)+'px) translate(-50%,-50%) scale('+sc0.toFixed(4)+')';
  const n=letters.length,sweep=(T-2.15)/.95*1.4-.2,dark=MODE==='dark';
  letters.forEach((s,i)=>{
    const o=clamp((T-1.45-i*.06)/.42,0,1),e=eo(o);
    const xi=(i+.5)/n,b=Math.exp(-Math.pow((xi-sweep)/.085,2))*(T>2.1&&T<3.4?1:0);
    s.style.opacity=e.toFixed(3);s.style.transform='translateY('+((1-e)*14).toFixed(1)+'px)';
    s.style.filter=o<1?'blur('+((1-e)*9).toFixed(1)+'px)':'';
    if(dark){const v=Math.round(220+35*b);s.style.color='rgb('+v+','+Math.round(232+23*b)+',255)';s.style.textShadow=b>.03?'0 0 '+(b*18).toFixed(1)+'px '+THEMES[THEME].ac+', 0 0 '+(b*4).toFixed(1)+'px #fff':'';}
    else{s.style.color='#14264f';s.style.textShadow=b>.03?'0 0 '+(b*14).toFixed(1)+'px '+THEMES[THEME].acl:''}
  });
  /* 卡片凝結 */
  const cR=(T-2.45)/.85,c0=eo(cR),c=c0>.99?1:c0;   /* 開場結束時（T=3.2）曲線停在 .998，收成 1，避免殘留 0.04px 模糊 */
  if(ITTS.state==='success'){
    const k=1-sstep(0,.3,succ.p);card.style.opacity=k;card.style.transform='scale('+(1+succ.p*.08).toFixed(3)+')';card.style.filter=succ.p>0?'blur('+(succ.p*8).toFixed(1)+'px)':'';
  }else{
    card.style.opacity=(c*fadeIn).toFixed(3);
    card.style.transform=c<1?'scale('+(.9+.1*c).toFixed(4)+')':'';
    card.style.filter=c<1?'blur('+((1-c)*22).toFixed(1)+'px)':'';
  }
  card.style.pointerEvents='auto';   /* 開場期間表單就能點、能打字，不被擋 */
  if(cR>0&&cR<1.05){
    const r=card.getBoundingClientRect(),g=clamp(cR,0,1);
    const pad=-8+g*70;ring.style.cssText='left:'+(r.left-pad)+'px;top:'+(r.top-pad)+'px;width:'+(r.width+pad*2)+'px;height:'+(r.height+pad*2)+'px;opacity:'+((1-g)*(1-g)*.95).toFixed(3)+';border-radius:'+(32+g*60)+'px';
    const fs=900*eo(g);flash.style.cssText='left:'+(r.left+r.width/2-fs/2)+'px;top:'+(r.top+r.height/2-fs/2)+'px;width:'+fs+'px;height:'+fs+'px;opacity:'+((1-g)*.55).toFixed(3);
  }else{ring.style.opacity=0;flash.style.opacity=0}
  skipEl.classList.toggle('show',playIntro&&T<2.6&&!skipping&&!introDone);
  /* lite/static 的 SVG logo */
  if(tier!=='full'){
    const p=clamp((T-.15)/1.5,0,1),e=eo(p);
    logoSvg.style.opacity=(sstep(.05,.5,T)*fadeIn*(ITTS.state==='success'?Math.max(0,1-succ.p*1.2):1)).toFixed(3);
    const wpx=w,hpx=w*hh;
    logoSvg.style.width=wpx+'px';logoSvg.style.height=hpx+'px';
    logoSvg.style.transform='translate('+X.toFixed(1)+'px,'+Y.toFixed(1)+'px) translate(-50%,-50%) scale('+(sc0*(ITTS.state==='success'?1+succ.p*3:1)).toFixed(4)+')';
    logoSvg.style.filter='blur('+((1-e)*14).toFixed(1)+'px) drop-shadow(0 0 14px '+THEMES[THEME].ac+'88)';
    logoSvg.querySelectorAll('.g *').forEach(el=>{el.style.strokeDashoffset=String(1-e)});
    logoSvg.querySelector('.t').style.opacity=sstep(.8,1.7,T);
  }
}
const camOf=T=>120*eo(Math.min(T,4.2)/4.2)+(ITTS.state==='success'?Math.pow(succ.p,2)*500:0);

/* ── 分級 ── */
function setTier(t,w){
  tier=t;why=w||'';ROOT.dataset.tier=t;window.__tier=t;
  logoSvg.classList.toggle('on',t!=='full');ITTS.tier=t;
  if(t==='full'){cvGL.hidden=false;cvLite.hidden=true}
  else{
    cvGL.hidden=true;cvLite.hidden=false;
    if(gl){releasing=true;try{const e=gl.getExtension('WEBGL_lose_context');if(e)e.loseContext()}catch(_){}gl=null;G=null}
    sizeLite();
    if(t==='static'){T=INTRO;fadeIn=1;playIntro=false;drawLite(T,0);domTL(T)}
  }
}

/* ── WebGL ── */
function compile(type,src){const s=gl.createShader(type);gl.shaderSource(s,src);gl.compileShader(s);if(!gl.getShaderParameter(s,gl.COMPILE_STATUS))throw new Error(gl.getShaderInfoLog(s));return s}
function program(vs,fs,attrs){const p=gl.createProgram();gl.attachShader(p,compile(gl.VERTEX_SHADER,vs));gl.attachShader(p,compile(gl.FRAGMENT_SHADER,fs));attrs.forEach((a,i)=>gl.bindAttribLocation(p,i,a));gl.linkProgram(p);if(!gl.getProgramParameter(p,gl.LINK_STATUS))throw new Error(gl.getProgramInfoLog(p));
  const u={};const n=gl.getProgramParameter(p,gl.ACTIVE_UNIFORMS);for(let i=0;i<n;i++){const a=gl.getActiveUniform(p,i);u[a.name]=gl.getUniformLocation(p,a.name)}return{p,u}}
function fbo(w,h){const tex=gl.createTexture();gl.bindTexture(gl.TEXTURE_2D,tex);gl.texImage2D(gl.TEXTURE_2D,0,gl.RGBA,w,h,0,gl.RGBA,gl.UNSIGNED_BYTE,null);
  gl.texParameteri(gl.TEXTURE_2D,gl.TEXTURE_MIN_FILTER,gl.LINEAR);gl.texParameteri(gl.TEXTURE_2D,gl.TEXTURE_MAG_FILTER,gl.LINEAR);gl.texParameteri(gl.TEXTURE_2D,gl.TEXTURE_WRAP_S,gl.CLAMP_TO_EDGE);gl.texParameteri(gl.TEXTURE_2D,gl.TEXTURE_WRAP_T,gl.CLAMP_TO_EDGE);
  const f=gl.createFramebuffer();gl.bindFramebuffer(gl.FRAMEBUFFER,f);gl.framebufferTexture2D(gl.FRAMEBUFFER,gl.COLOR_ATTACHMENT0,gl.TEXTURE_2D,tex,0);return{tex,f,w,h}}
function buildBuffers(){
  const mob=mobile();const NL=mob?4200:6200,ND=mob?1500:2600;
  LOGO=LOGO&&LOGO.n===NL?LOGO:sampleLogo(NL);
  const R=rng(777);
  const lg=new Float32Array(NL*2*8);
  for(let i=0;i<NL;i++){const r0=R(),r1=R(),r2=R(),r3=R();for(let s=0;s<2;s++){const o=(i*2+s)*8;lg[o]=LOGO.pts[i*2];lg[o+1]=LOGO.pts[i*2+1];lg[o+2]=(R()-.5)*2;lg[o+3]=r0;lg[o+4]=r1;lg[o+5]=r2;lg[o+6]=r3;lg[o+7]=s}}
  const ds=new Float32Array(ND*2*8);
  for(let i=0;i<ND;i++){const r0=R(),r1=R(),r2=R(),r3=R(),x=(R()*2-1)*innerWidth*1.05,y=(R()*2-1)*innerHeight*1.05,z=-300+R()*1700;for(let s=0;s<2;s++){const o=(i*2+s)*8;ds[o]=x;ds[o+1]=y;ds[o+2]=z;ds[o+3]=r0;ds[o+4]=r1;ds[o+5]=r2;ds[o+6]=r3;ds[o+7]=s}}
  const mk=(data,n)=>{const b=gl.createBuffer();gl.bindBuffer(gl.ARRAY_BUFFER,b);gl.bufferData(gl.ARRAY_BUFFER,data,gl.STATIC_DRAW);
    const idx=new Uint16Array(n);for(let i=0;i<n;i++)idx[i]=i*2;const ib=gl.createBuffer();gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER,ib);gl.bufferData(gl.ELEMENT_ARRAY_BUFFER,idx,gl.STATIC_DRAW);return{b,ib,n}};
  G.logo=mk(lg,NL);G.dust=mk(ds,ND);
}
function initGL(){
  gl=cvGL.getContext('webgl',{antialias:false,alpha:false,depth:false,stencil:false,powerPreference:'high-performance'});
  if(!gl)throw new Error('no webgl');
  G={pt:program(VS_P,FS_PT,['aP','aR','aS']),ln:program(VS_P,FS_LN,['aP','aR','aS']),bl:program(VS_Q,FS_BLUR,['a']),co:program(VS_Q,FS_COMP,['a'])};
  const q=gl.createBuffer();gl.bindBuffer(gl.ARRAY_BUFFER,q);gl.bufferData(gl.ARRAY_BUFFER,new Float32Array([-1,-1,3,-1,-1,3]),gl.STATIC_DRAW);G.quad=q;
  buildBuffers();
  cvGL.addEventListener('webglcontextlost',e=>{e.preventDefault();if(!releasing)toStatic('context lost')},false);
  resizeGL();
}
function resizeGL(){
  const dpr=Math.min(devicePixelRatio||1,mobile()?2:1.5);
  const w=Math.max(2,Math.round(innerWidth*dpr)),h=Math.max(2,Math.round(innerHeight*dpr));
  if(G.w===w&&G.h===h)return;
  cvGL.width=w;cvGL.height=h;G.w=w;G.h=h;G.dpr=w/innerWidth;
  G.A=fbo(w,h);const q4=[Math.max(2,w>>2),Math.max(2,h>>2)],q8=[Math.max(2,w>>3),Math.max(2,h>>3)];
  G.T1=fbo(q4[0],q4[1]);G.T1b=fbo(q4[0],q4[1]);G.T2=fbo(q8[0],q8[1]);G.T2b=fbo(q8[0],q8[1]);G.An=fbo(q4[0],q4[1]);
  gl.bindFramebuffer(gl.FRAMEBUFFER,null);
}
function uni(p,n,v){const l=p.u[n];if(l===undefined||l===null)return;if(typeof v==='number')gl.uniform1f(l,v);else if(v.length===2)gl.uniform2f(l,v[0],v[1]);else if(v.length===3)gl.uniform3f(l,v[0],v[1],v[2])}
function bindPts(buf){gl.bindBuffer(gl.ARRAY_BUFFER,buf.b);gl.enableVertexAttribArray(0);gl.enableVertexAttribArray(1);gl.enableVertexAttribArray(2);
  gl.vertexAttribPointer(0,3,gl.FLOAT,false,32,0);gl.vertexAttribPointer(1,4,gl.FLOAT,false,32,12);gl.vertexAttribPointer(2,1,gl.FLOAT,false,32,28)}
function pass(prog,target,w,h,setup){gl.bindFramebuffer(gl.FRAMEBUFFER,target?target.f:null);gl.viewport(0,0,w,h);gl.useProgram(prog.p);
  gl.bindBuffer(gl.ARRAY_BUFFER,G.quad);gl.disableVertexAttribArray(1);gl.disableVertexAttribArray(2);gl.enableVertexAttribArray(0);gl.vertexAttribPointer(0,2,gl.FLOAT,false,0,0);setup();gl.drawArrays(gl.TRIANGLES,0,3)}
function tex(unit,t,prog,name){gl.activeTexture(gl.TEXTURE0+unit);gl.bindTexture(gl.TEXTURE_2D,t.tex);gl.uniform1i(prog.u[name],unit)}
const ptr={x:0,y:0,on:0,px:0,py:0};

function sceneParams(){
  const light=PAL.s[2]>.5;
  const inv=c=>c.map(v=>1-v);
  /* 暗：粒子色＝白＋強調色；亮：粒子色＝(1−墨色)，合成時用減法 */
  const c0=light?inv(PAL.ac):mixc(PAL.ac,[1,1,1],.72),c1=light?inv(mixc(PAL.ac,PAL.c0,.4)):mixc(PAL.ac,PAL.c0,.15);
  return{light,c0,c1};
}
function frameGL(t,dt){
  palStep(dt);
  const sp=sceneParams(),W=G.w,H=G.h,k=G.dpr;
  const sP=succ.t0>=0?succ.p:0;
  const asm=clamp((T-.2)/1.5,0,1),cam=camOf(T);
  const shift=sstep(1.9,2.85,T);
  const cx=lerp(L.sx,L.ex,shift)-L.W/2,cy=lerp(L.sy,L.ey,shift)-L.H/2,scale=lerp(L.sw,L.ew,shift);
  /* 指標（相對畫面中心，css px） */
  const on=pointer.moved&&performance.now()-pointer.t<7000?1:0;ptr.on+=(on-ptr.on)*Math.min(1,dt*2.5);
  const tx=pointer.x-L.W/2,ty=pointer.y-L.H/2;ptr.x+=(tx-ptr.x)*Math.min(1,dt*9);ptr.y+=(ty-ptr.y)*Math.min(1,dt*9);
  ptr.px+=(pointer.nx*20-ptr.px)*Math.min(1,dt*2);ptr.py+=(pointer.ny*14-ptr.py)*Math.min(1,dt*2);
  const burst=sP>0?eo(sP/.4):0,warp=sP>0?Math.pow(clamp((sP-.18)/.78,0,1),1.5):0;
  const bloom=1+.75*Math.exp(-Math.pow((T-2.5)/.35,2))*(playIntro?1:0)+sP*1.4+(1-sstep(0,1.4,T))*.0;
  const common=p=>{
    uni(p,'uCss',[L.W,L.H]);uni(p,'uCtr',[cx,cy]);uni(p,'uPtr',[ptr.x,ptr.y]);uni(p,'uPar',[ptr.px,ptr.py]);
    uni(p,'uT',Tabs);uni(p,'uAsm',asm);uni(p,'uCam',cam);uni(p,'uScale',scale);uni(p,'uWarp',warp);uni(p,'uBurst',burst);
    uni(p,'uShim',sstep(1.6,2.6,T));uni(p,'uPtrOn',ptr.on*(sP>0?0:1));uni(p,'uDpr',k);uni(p,'uTrail',.17);uni(p,'uDustA',sstep(0,.8,T)*fadeIn);uni(p,'uLogoA',sstep(0,.4,T)*fadeIn);
    uni(p,'uC0',sp.c0);uni(p,'uC1',sp.c1)};
  /* 1) 粒子 → A */
  gl.bindFramebuffer(gl.FRAMEBUFFER,G.A.f);gl.viewport(0,0,W,H);gl.clearColor(0,0,0,1);gl.clear(gl.COLOR_BUFFER_BIT);
  gl.enable(gl.BLEND);gl.blendFunc(gl.ONE,gl.ONE);
  gl.useProgram(G.pt.p);common(G.pt);
  uni(G.pt,'uType',1);bindPts(G.dust);gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER,G.dust.ib);gl.drawElements(gl.POINTS,G.dust.n,gl.UNSIGNED_SHORT,0);
  uni(G.pt,'uType',0);bindPts(G.logo);gl.bindBuffer(gl.ELEMENT_ARRAY_BUFFER,G.logo.ib);gl.drawElements(gl.POINTS,G.logo.n,gl.UNSIGNED_SHORT,0);
  if(warp>.015){
    gl.useProgram(G.ln.p);common(G.ln);
    uni(G.ln,'uType',1);bindPts(G.dust);gl.drawArrays(gl.LINES,0,G.dust.n*2);
    uni(G.ln,'uType',0);bindPts(G.logo);gl.drawArrays(gl.LINES,0,G.logo.n*2);
  }
  gl.disable(gl.BLEND);
  /* 2) 泛光：亮部 → 1/4 H → V；1/8 H → V；變形（水平拉絲） */
  const A=G.A,T1=G.T1,T1b=G.T1b,T2=G.T2,T2b=G.T2b,An=G.An;
  const blur=(src,dst,dx,dy,thr,scale)=>pass(G.bl,dst,dst.w,dst.h,()=>{tex(0,src,G.bl,'uTex');uni(G.bl,'uStep',[dx/src.w*scale,dy/src.h*scale]);uni(G.bl,'uOut',[dst.w,dst.h]);uni(G.bl,'uThr',thr)});
  blur(A,T1,1,0,.12,1.2);blur(T1,T1b,0,1,0,1.2);
  blur(T1b,T2,1,0,0,1.6);blur(T2,T2b,0,1,0,1.6);
  blur(T1,An,1,0,0,5.5);
  /* 3) 合成 */
  const light=sp.light;
  pass(G.co,null,W,H,()=>{
    tex(0,A,G.co,'uA');tex(1,T1b,G.co,'uB1');tex(2,T2b,G.co,'uB2');tex(3,An,G.co,'uAn');
    uni(G.co,'uRes',[W,H]);uni(G.co,'uT',Tabs);uni(G.co,'uBloom',bloom);uni(G.co,'uFade',sP>0?sstep(.78,1,sP):0);uni(G.co,'uLight',light?1:0);uni(G.co,'uWarp',warp);
    uni(G.co,'uAn1',(.2+sP*1.2)*(light?.4:1));
    uni(G.co,'uBg',PAL.bg);uni(G.co,'uN1',light?PAL.c0.map(v=>v*.5):PAL.c0);uni(G.co,'uN2',light?PAL.c2.map(v=>v*.5):PAL.c2);uni(G.co,'uApp',PAL.app);
  });
  if(!ROOT.classList.contains('gl'))ROOT.classList.add('gl');
}

/* ── Canvas2D（lite／static）：星塵＋CSS logo ── */
let dust2=null;
function sizeLite(){const d=Math.min(devicePixelRatio||1,1.5);cvLite.width=Math.round(innerWidth*d);cvLite.height=Math.round(innerHeight*d);lctx=cvLite.getContext('2d');
  const R=rng(99);dust2=Array.from({length:mobile()?140:260},()=>({x:(R()*2-1)*innerWidth,y:(R()*2-1)*innerHeight,z:-300+R()*1700,s:.6+R()*1.6,a:.3+R()*.7,sp:16+R()*34,h:R()}))}
function drawLite(t,dt){
  if(!lctx)return;const c=lctx,w=cvLite.width,h=cvLite.height,d=w/innerWidth,light=PAL.s[2]>.5;
  c.globalCompositeOperation='source-over';c.clearRect(0,0,w,h);
  const sP=succ.t0>=0?succ.p:0,warp=sP>0?Math.pow(clamp((sP-.18)/.78,0,1),1.5):0,cam=camOf(T);
  const ac=THEMES[THEME].ac;const da=sstep(0,.8,T)*fadeIn;
  c.globalCompositeOperation=light?'source-over':'lighter';
  for(const s of dust2){
    if(dt)s.z-=dt*s.sp;if(s.z<-300)s.z+=1700;
    const z=s.z-warp*warp*(2400+1700*s.h),dist=1000+z-cam;if(dist<40)continue;const sc=1000/dist;
    const px=(innerWidth/2+(s.x-ptrL.x)*sc+ptrL.x*.22)*d,py=(innerHeight/2+(s.y-ptrL.y)*sc+ptrL.y*.22)*d;
    const al=s.a*da*Math.min(1,(s.z+300)/180)*Math.min(1,(1400-s.z)/350)*(.65+.35*Math.sin(t*(1+s.h*3)+s.h*40));
    c.fillStyle=light?'rgba(21,88,214,'+(al*.6).toFixed(3)+')':(s.h>.55?'rgba(200,225,255,':'rgba(150,200,255,')+al.toFixed(3)+')';
    if(warp>.02){const z2=s.z-Math.pow(Math.max(warp-.17,0),2)*(2400+1700*s.h),d2=1000+z2-cam,sc2=d2>40?1000/d2:sc;c.strokeStyle=c.fillStyle;c.lineWidth=Math.max(1,s.s*sc*d);c.beginPath();c.moveTo(px,py);c.lineTo((innerWidth/2+(s.x-ptrL.x)*sc2+ptrL.x*.22)*d,(innerHeight/2+(s.y-ptrL.y)*sc2+ptrL.y*.22)*d);c.stroke()}
    else{const r=Math.max(.8,s.s*sc*d);c.fillRect(px-r/2,py-r/2,r,r)}
  }
  c.globalCompositeOperation='source-over';
}
const ptrL={x:0,y:0};
function frameLite(t,dt){
  palStep(dt);ptrL.x+=(pointer.nx*20-ptrL.x)*Math.min(1,dt*2);ptrL.y+=(pointer.ny*14-ptrL.y)*Math.min(1,dt*2);
  drawLite(Tabs,dt);
}
function toStatic(w){if(tier==='static')return;setTier('static',w);if(mainLoop)mainLoop.stop();mainLoop=null;
  if(succ.t0>=0){fadeEl.style.transition='opacity .5s';fadeEl.style.opacity='1';done.classList.add('show')}}

/* ── 事件 ── */
function markSeen(){}
function skip(){if(introDone||skipping||!playIntro||T>=INTRO)return;skipping=true;markSeen()}
addEventListener('pointerdown',skip,{capture:true});
addEventListener('keydown',e=>{if(!/^(Shift|Control|Alt|Meta|CapsLock)$/.test(e.key))skip()},{capture:true});
ITTS.onState=s=>{
  if(s==='success'&&succ.t0<0){
    succ.t0=performance.now();
    if(tier==='static'){fadeEl.style.transition='opacity .5s';fadeEl.style.opacity='1';setTimeout(()=>done.classList.add('show'),500)}
    setTimeout(()=>done.classList.add('show'),succ.dur*.85);
  }
};
ITTS.onTheme=()=>{if(tier==='static'){PAL=PAL_T;drawLite(T,0);domTL(T)}else if(tier!=='full')domTL(T)};
addEventListener('resize',()=>{layout();if(tier==='full'&&G){resizeGL();}else{sizeLite()}if(tier==='static'){drawLite(T,0);domTL(T)}});

/* ── 啟動 ── */
function start(){
  layout();
  let t0=null;if(MOTION_REDUCE)t0='static';
  if(!t0){t0=webglSupported()?'full':'lite';if(t0==='lite')why='no webgl'}
  if(t0==='full'){
    try{LOGO=LOGO||sampleLogo(mobile()?4200:6200);initGL();setTier('full','auto')}catch(e){console.warn('[cinematic] WebGL init failed → lite:',e.message);gl=null;G=null;t0='lite';why='init failed'}
  }
  if(t0!=='full'){LOGO=LOGO||{hh:.33};setTier(t0,why||'auto')}
  if(tier==='static')return;
  /* 看門狗：1.2 秒內動畫迴圈沒起來 → 立刻讓表單可見可用 */
  setTimeout(()=>{if(frames===0){fadeIn=1;T=INTRO;card.style.opacity=1;card.style.transform='';card.style.filter='';card.style.pointerEvents='auto';console.warn('[cinematic] watchdog: loop not running, form revealed')}},1200);
  let perf=null;
  /* 時間式自動降級：full 平均 >33ms/幀 → lite；lite 平均 >50ms/幀 → static（舊電腦） */
  const armLite=()=>{perf=makePerf(40,avg=>{perfMs=avg;if(avg>50&&tier==='lite')toStatic('slow lite '+avg.toFixed(0)+'ms/frame')})};
  const armFull=()=>{perf=makePerf(40,avg=>{perfMs=avg;if(avg>33.4&&tier==='full'){setTier('lite','slow '+avg.toFixed(0)+'ms/frame');armLite()}})};
  if(tier==='full')armFull();else armLite();
  mainLoop=loop((t,dt,ts)=>{
    try{
      frames++;Tabs=t;
      if(perf)perf(ts);
      if(playIntro&&!introDone){T+=dt*(skipping?5:1);if(T>=INTRO){T=INTRO;introDone=true;markSeen()}}
      else if(introSkippedBecauseSeen&&fadeIn<1){fadeIn=Math.min(1,fadeIn+dt/.9)}
      if(succ.t0>=0)succ.p=clamp((performance.now()-succ.t0)/succ.dur,0,1);
      domTL(T);
      if(tier==='full')frameGL(t,dt);else frameLite(t,dt);
      if(succ.t0>=0&&tier!=='full'){fadeEl.style.opacity=String(sstep(.78,1,succ.p))}
    }catch(err){console.error('[cinematic] frame error → fail-safe',err);card.style.opacity=1;card.style.transform='';card.style.filter='';card.style.pointerEvents='auto';return false}
  });
}
window.__cin={get tier(){return tier},get why(){return why},get frames(){return frames},get T(){return T},get perfMs(){return perfMs},get seen(){return seen},get playIntro(){return playIntro},get skipped(){return skipping}};
start();
try{if(matchMedia('(pointer:fine)').matches)$('#username').focus({preventScroll:true})}catch(e){}
})();
})();
