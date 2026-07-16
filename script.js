// ── cached DOM refs (avoid repeated getElementById lookups) ──
const $intro   = document.getElementById('intro');
const $iname   = document.getElementById('iname');
const $isub    = document.getElementById('isub');
const $wrap    = document.getElementById('wrap');
const $sidebar = document.getElementById('sidebar');
const $dim     = document.getElementById('dim');
const $hbg     = document.getElementById('hbg');
const $metro   = document.getElementById('metro');
const $metroWrap = document.getElementById('metro-wrap');
const $trans   = document.getElementById('trans');
const $transTitle = document.getElementById('trans-title');
const $tbBack  = document.getElementById('tb-back');
const $tbSep   = document.getElementById('tb-sep');
const $tbCat   = document.getElementById('tb-cat');
const $lb      = document.getElementById('lb');
const $lbImg   = document.getElementById('lb-img');
const $lbThumb = document.getElementById('lb-thumb');
const $lbVid   = document.getElementById('lb-vid');
const $lbMt    = document.getElementById('lb-mt');
const $lbMd    = document.getElementById('lb-md');
const $lbCtr   = document.getElementById('lb-ctr');

// ── utils ──────────────────────────────────
function shuffle(a){
  const b=[...a];
  for(let i=b.length-1;i>0;i--){
    const j=Math.floor(Math.random()*(i+1));
    [b[i],b[j]]=[b[j],b[i]];
  }
  return b;
}
function setAppHeight() {
  document.documentElement.style.setProperty(
    '--app-height',
    `${window.innerHeight}px`
  );
}
// debounce resize — recalculating on every pixel of a drag-resize/rotate is wasted work
function debounce(fn, wait){
  let t;
  return (...args)=>{ clearTimeout(t); t=setTimeout(()=>fn(...args), wait); };
}
window.addEventListener('resize', debounce(setAppHeight, 100));

// prevent easy copying — block right-click / drag-and-drop on images & videos
document.addEventListener('contextmenu', e=>{
  if(e.target.tagName==='IMG'||e.target.tagName==='VIDEO') e.preventDefault();
});
document.addEventListener('dragstart', e=>{
  if(e.target.tagName==='IMG'||e.target.tagName==='VIDEO') e.preventDefault();
});

// fallback: some mobile browsers block autoplay entirely until the user
// interacts with the page — retry any still-paused metro videos then
function kickAutoplayVideos(){
  document.querySelectorAll('#metro video').forEach(v=>{
    if(v.paused && v.src){ const p=v.play(); if(p&&p.catch)p.catch(()=>{}); }
  });
}
['touchstart','click'].forEach(evt=>document.addEventListener(evt,kickAutoplayVideos,{once:true,passive:true}));

// ── IntersectionObserver: pause offscreen gallery videos ──
// The justified grid can hold many autoplay videos at once; only the ones
// actually visible need to be decoding/playing. Saves CPU/battery/bandwidth
// on long scrolling categories.
const _videoVisObserver = new IntersectionObserver((entries)=>{
  entries.forEach(entry=>{
    const v = entry.target;
    if(entry.isIntersecting){
      if(v.paused){ const p=v.play(); if(p&&p.catch)p.catch(()=>{}); }
    } else {
      v.pause();
    }
  });
},{root:document.getElementById('pg-gallery'), rootMargin:'200px 0px'});

// ══════════════════════════════════════════════════════
// DATA — از فایل data.xlsx خونده می‌شه (کنار همین index.html روی هاست بذارش)
// شیت "Categories": ستون‌های Id, Title
// هر شیت دیگه (اسمش = همون Id بالا): ستون‌های Number, Type, Ext, Caption, Year, Camera, Lens
// Number = شماره فایل توی thumbs/<id>/ و hq/<id>/ (مثلاً 3 یعنی 03.jpg)
// Ext خالی باشه یعنی پیش‌فرض: jpg برای photo، mp4 برای video
// ══════════════════════════════════════════════════════
async function loadSiteData(){
  try{
    const resp = await fetch('data.xlsx');
    if(!resp.ok) throw new Error('data.xlsx not found ('+resp.status+')');
    const buf = await resp.arrayBuffer();
    const wb = XLSX.read(buf,{type:'array'});

    const catsSheet = wb.Sheets['Categories'];
    if(!catsSheet) throw new Error('Categories sheet missing in data.xlsx');
    const catsRows = XLSX.utils.sheet_to_json(catsSheet);

    return catsRows.map(row=>{
      const id = String(row.Id||'').trim();
      const title = String(row.Title||id).trim();
      const sheet = wb.Sheets[id];
      const rows = sheet ? XLSX.utils.sheet_to_json(sheet) : [];

      const items = rows.map(r=>{
        const num = String(r.Number).padStart(2,'0');
        const type = String(r.Type||'photo').toLowerCase()==='video' ? 'video' : 'photo';
        const ext = (r.Ext && String(r.Ext).trim()) || (type==='video' ? 'mp4' : 'jpg');
        return {
          thumb: `thumbs/${id}/${num}.${ext}`,
          hq: `hq/${id}/${num}.${ext}`,
          type,
          title: r.Caption ? String(r.Caption) : '',
          year: r.Year ? String(r.Year) : '',
          camera: r.Camera ? String(r.Camera) : '',
          lens: r.Lens ? String(r.Lens) : ''
        };
      });

      return {id, title, items};
    });
  }catch(err){
    console.error('Failed to load data.xlsx:', err);
    return [];
  }
}
// ══════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════

// ── state ──────────────────────────────────
let DATA=[];
let curCat=0;
let filtItems=[];
let lbIdx=0;
const _ratioCache=new Map(); // thumb path -> {ratio, isVthumb} — avoids re-probing video/image metadata on repeat visits

// ── INIT ───────────────────────────────────
window.addEventListener('load',()=>{
setAppHeight();
  setTimeout(()=>{$iname.classList.add('in');$isub.classList.add('in')},150);

  const dataPromise = loadSiteData(); // kick off fetch now, doesn't block the intro

  setTimeout(async ()=>{
    $iname.style.transition='transform .9s cubic-bezier(.77,0,.175,1),opacity .9s';
    $iname.style.transform='translateY(-24px)';
    $iname.style.opacity='0';
    $isub.style.opacity='0';
    DATA = await dataPromise;
    buildMetro();
  },1500);
  setTimeout(()=>$intro.classList.add('out'),1850);
  setTimeout(()=>$intro.style.display='none',2700);
});

// ── LAYOUT TEMPLATES (8 tiles, 12-col grid) ─
// هر template دقیقاً ۱۲ستون × N ردیف رو پر می‌کنه — بدون gap
const TEMPLATES = [
  { rows:3, areas:[
    {c:'1/5',  r:'1/3'},
    {c:'5/13', r:'1/2'},
    {c:'5/9',  r:'2/3'},
    {c:'9/13', r:'2/3'},
    {c:'1/4',  r:'3/4'},
    {c:'4/7',  r:'3/4'},
    {c:'7/10', r:'3/4'},
    {c:'10/13',r:'3/4'},
  ]},
  { rows:3, areas:[
    {c:'9/13', r:'1/3'},
    {c:'1/9',  r:'1/2'},
    {c:'1/5',  r:'2/3'},
    {c:'5/9',  r:'2/3'},
    {c:'1/4',  r:'3/4'},
    {c:'4/7',  r:'3/4'},
    {c:'7/10', r:'3/4'},
    {c:'10/13',r:'3/4'},
  ]},
  { rows:3, areas:[
    {c:'1/5',  r:'1/2'},
    {c:'5/9',  r:'1/2'},
    {c:'9/13', r:'1/2'},
    {c:'1/7',  r:'2/4'},
    {c:'7/10', r:'2/3'},
    {c:'10/13',r:'2/3'},
    {c:'7/10', r:'3/4'},
    {c:'10/13',r:'3/4'},
  ]},
  { rows:2, areas:[
    {c:'1/4',  r:'1/2'},
    {c:'4/7',  r:'1/2'},
    {c:'7/10', r:'1/2'},
    {c:'10/13',r:'1/2'},
    {c:'1/4',  r:'2/3'},
    {c:'4/7',  r:'2/3'},
    {c:'7/10', r:'2/3'},
    {c:'10/13',r:'2/3'},
  ]},
  { rows:3, areas:[
    {c:'1/13', r:'1/2'},
    {c:'1/5',  r:'2/3'},
    {c:'5/9',  r:'2/3'},
    {c:'9/13', r:'2/3'},
    {c:'1/7',  r:'3/4'},
    {c:'7/9',  r:'3/4'},
    {c:'9/11', r:'3/4'},
    {c:'11/13',r:'3/4'},
  ]},
  { rows:3, areas:[
    {c:'1/7',  r:'1/2'},
    {c:'7/13', r:'1/2'},
    {c:'1/4',  r:'2/4'},
    {c:'10/13',r:'2/4'},
    {c:'4/7',  r:'2/3'},
    {c:'7/10', r:'2/3'},
    {c:'4/7',  r:'3/4'},
    {c:'7/10', r:'3/4'},
  ]},
];

const MOBILE_TEMPLATE = { rows:6, rowTemplate:'0.75fr 0.75fr 1fr 1fr 1fr 1fr', areas:[
  {c:'1/13', r:'1/3'},
  {c:'1/7',  r:'3/4'},
  {c:'7/13', r:'3/4'},
  {c:'1/7',  r:'4/6'},
  {c:'7/13', r:'4/5'},
  {c:'7/13', r:'5/6'},
  {c:'1/7',  r:'6/7'},
  {c:'7/13', r:'6/7'},
]};

// ── TILE MEDIA (staged load + random slide/roll transitions) ─
function buildTileMedia(tile,cat){
  const isVideoCat = cat.items.some(it=>it.type==='video');
  const cap = isVideoCat ? 3 : 6;
  const allItems=cat.items.slice(0,cap);
  if(allItems.length===0)return;

  const slides=[];

  function makeSlide(item){
    const isVideo = item.type==='video' || /\.(mp4|mov|webm)$/i.test(item.thumb||'');
    const s=document.createElement('div');
    s.className='sl';
    tile.appendChild(s);
    if(isVideo){
      const posterPath = item.poster ? `public/${item.poster}` : `public/${item.thumb.replace(/\.(mp4|mov|webm)$/i,'.jpg')}`;
      s.innerHTML=`<video src="public/${item.thumb}" poster="${posterPath}" loop muted playsinline webkit-playsinline="true" preload="auto"></video>`;
      const v=s.querySelector('video');
      const tryPlay=()=>{ const p=v.play(); if(p&&p.catch)p.catch(()=>{}); };
      tryPlay();
      v.addEventListener('loadeddata',tryPlay);
      v.addEventListener('canplay',tryPlay);
    } else {
      s.innerHTML=`<img src="public/${item.thumb}" alt="${cat.title}" decoding="async"/>`;
    }
    slides.push(s);
    return s;
  }

  const eagerCount=Math.min(3,allItems.length);
  for(let i=0;i<eagerCount;i++) makeSlide(allItems[i]);

  let curIdx=Math.floor(Math.random()*slides.length);
  slides[curIdx].classList.add('on');
  slides[curIdx].style.transform='translate(0,0)';

  if(allItems.length>eagerCount){
    setTimeout(()=>{
      for(let i=eagerCount;i<allItems.length;i++) makeSlide(allItems[i]);
    },3000+Math.random()*1500);
  }

  if(allItems.length<2)return;

  const DIRS=[{x:1,y:0},{x:-1,y:0},{x:0,y:1},{x:0,y:-1}];

  function roll(){
    if(slides.length<2)return;
    let nextIdx;
    do{ nextIdx=Math.floor(Math.random()*slides.length); }while(nextIdx===curIdx);
    const d=DIRS[Math.floor(Math.random()*DIRS.length)];
    const outEl=slides[curIdx], inEl=slides[nextIdx];

    const inVid=inEl.querySelector('video');
    if(inVid){ const p=inVid.play(); if(p&&p.catch)p.catch(()=>{}); }

    inEl.style.transition='none';
    inEl.style.transform=`translate(${d.x*100}%, ${d.y*100}%)`;
    inEl.classList.add('on');
    void inEl.offsetWidth;

    inEl.style.transition='transform .85s cubic-bezier(.65,0,.35,1)';
    outEl.style.transition='transform .85s cubic-bezier(.65,0,.35,1)';
    requestAnimationFrame(()=>{
      inEl.style.transform='translate(0,0)';
      outEl.style.transform=`translate(${-d.x*100}%, ${-d.y*100}%)`;
    });

    setTimeout(()=>{
      outEl.classList.remove('on');
      outEl.style.transition='none';
      outEl.style.transform='translate(0,0)';
    },870);

    curIdx=nextIdx;
  }

  function scheduleNext(delay){
    setTimeout(()=>{
      roll();
      scheduleNext(5000+Math.random()*3000);
    },delay);
  }

  scheduleNext(4000+Math.random()*4000);
}

// ── BUILD METRO ────────────────────────────
let _savedTpl = null;
let _savedCats = null;

function buildMetro(reuse=false){
  $metro.innerHTML='';

  $metroWrap.style.transition='none';
  $metroWrap.style.transform='scale(1)';
  $metroWrap.style.opacity='1';
  $metroWrap.style.transformOrigin='center center';

  const isMobile=window.innerWidth<=700;
  if(!reuse || !_savedTpl){
    _savedTpl = isMobile ? MOBILE_TEMPLATE : TEMPLATES[Math.floor(Math.random()*TEMPLATES.length)];
    if(isMobile){
      const obs = DATA.find(c=>c.id==='observations');
      const rest = shuffle(DATA.filter(c=>c.id!=='observations'));
      _savedCats = obs ? [obs, ...rest] : shuffle([...DATA]);
    } else {
      _savedCats = shuffle([...DATA]);
    }
  }
  const tpl = _savedTpl;
  const cats = _savedCats;

  $metro.style.gridTemplateRows = tpl.rowTemplate ? tpl.rowTemplate : `repeat(${tpl.rows},1fr)`;

  const frag = document.createDocumentFragment();
  cats.forEach((cat,i)=>{
    if(i>7)return;
    const area=tpl.areas[i];
    const tile=document.createElement('div');
    tile.className='tile';
    tile.style.gridColumn=area.c;
    tile.style.gridRow=area.r;

    buildTileMedia(tile,cat);

    const lbl=document.createElement('div');
    lbl.className='tile-lbl';
    lbl.innerHTML=`<div class="tile-lbl-t">${cat.title}</div>
      <div class="tile-lbl-c">${cat.items.length} Projects</div>`;
    tile.appendChild(lbl);

    tile.addEventListener('click',()=>zoomTransition(tile,DATA.indexOf(cat)));
    frag.appendChild(tile);
  });
  $metro.appendChild(frag);
}

// ── iOS ZOOM + CINEMATIC TRANSITION ────────
function zoomTransition(tile,catIdx){
  const cat=DATA[catIdx];

  const tRect=tile.getBoundingClientRect();
  const wRect=$metroWrap.getBoundingClientRect();
  const ox=((tRect.left+tRect.width/2)-wRect.left)/wRect.width*100;
  const oy=((tRect.top+tRect.height/2)-wRect.top)/wRect.height*100;

  $metroWrap.style.transformOrigin=`${ox}% ${oy}%`;
  $metroWrap.style.transition='transform .55s cubic-bezier(.4,0,.2,1),opacity .45s ease';

  $metroWrap.style.transform='scale(1.18)';
  $metroWrap.style.opacity='0';

  setTimeout(()=>{
    $trans.classList.add('fade-in');
    $transTitle.textContent=cat.title;
    setTimeout(()=>$transTitle.classList.add('show'),100);
  },300);

  setTimeout(()=>{
    openCat(catIdx);
  },500);

  setTimeout(()=>{
    $transTitle.classList.remove('show');
    $trans.style.transition='opacity .5s ease';
    $trans.classList.remove('fade-in');
    setTimeout(()=>{
      document.querySelector('.photo-grid').classList.add('in');
    },200);
  },1100);

  setTimeout(()=>{
    $trans.style.transition='';
    $transTitle.textContent='';
  },1700);
}

// ── OPEN CATEGORY ──────────────────────────
let _catGen=0;
async function openCat(ci, fromPop=false){
  const myGen=++_catGen;
  curCat=ci;
  const cat=DATA[ci];
  filtItems=[...cat.items];

  $tbBack.style.display='block';
  $tbSep.style.display='block';
  $tbCat.style.display='block';
  $tbCat.textContent=cat.title;

  await renderGrid(filtItems,myGen);
  if(myGen!==_catGen)return;

  setTimeout(()=>{
    document.getElementById('pg-gallery').scrollTop = 0;
  },0);

  showPg('gallery',false);

  if(!fromPop){
    history.pushState(
        {page:'gallery',cat:ci},
        '',
        '#'+cat.id
    );
  }
}

// ── JUSTIFIED GRID (masonry columns) ──────
async function renderGrid(items,myGen){
  const grid = document.getElementById('photo-grid');
  grid.className = 'photo-grid';
  grid.innerHTML = '';
  grid.style.cssText = 'display:flex;gap:10px;padding:3px;align-items:flex-start;opacity:0;transition:opacity .5s ease';

  const itemsWithRatio = await Promise.all(items.map((item,index) => new Promise(resolve => {
    const cached=_ratioCache.get(item.thumb);
    if(cached){
      resolve({...item, ...cached, originalIndex:index});
      return;
    }
    const isVthumb = item.thumb && /\.(mp4|mov|webm)$/i.test(item.thumb);
    if(isVthumb){
        const v = document.createElement('video');
        v.muted = true;
        v.preload = 'metadata';
        v.onloadedmetadata = () => {
          const r = v.videoWidth && v.videoHeight ? v.videoWidth/v.videoHeight : 1.25;
          const result={ratio:r,isVthumb:true};
          _ratioCache.set(item.thumb,result);
          resolve({...item, ...result, originalIndex:index});
        };
        v.onerror = () => resolve({...item, ratio:1.25, isVthumb:true, originalIndex:index});
        v.src = 'public/' + item.thumb;
        return;
      }
    const img = new Image();
    img.onload = () => {
      const result={ratio: img.naturalWidth / img.naturalHeight, isVthumb:false};
      _ratioCache.set(item.thumb,result);
      resolve({...item, ...result, originalIndex: index});
    };
    img.onerror = () => resolve({...item, ratio: 1.25, isVthumb: false, originalIndex: index});
    img.src = 'public/' + item.thumb;
  })));

  if(myGen!==undefined && myGen!==_catGen)return;

  const W    = window.innerWidth;
  const cols = W < 600 ? 2 : W < 1000 ? 3 : 4;
  const GAP  = 3;
  const colW = Math.floor((W - GAP*(cols+1)) / cols);

  const colEls = [], colH = [];
  for(let i = 0; i < cols; i++){
    const c = document.createElement('div');
    c.style.cssText = 'flex:1;display:flex;flex-direction:column;gap:10px;min-width:0';
    grid.appendChild(c);
    colEls.push(c); colH.push(0);
  }

  const sorted = [...itemsWithRatio].reverse();

  sorted.forEach((item) => {
    const h   = Math.round(colW / item.ratio);
    const card = document.createElement('div');
    card.style.cssText = 'position:relative;overflow:hidden;cursor:pointer;background:#111;width:100%;border-radius:10px';

    const origIdx = item.originalIndex;
    if(item.isVthumb){
      const posterPath = item.poster ? `public/${item.poster}` : `public/${item.thumb.replace(/\.(mp4|mov|webm)$/i,'.jpg')}`;
      card.innerHTML = `<video src="public/${item.thumb}" poster="${posterPath}" muted loop playsinline style="width:100%;height:${h}px;object-fit:cover;display:block;filter:brightness(.85)"></video><div class="pi-ov"><span>▶</span></div>`;
    } else {
      card.innerHTML = `<img src="public/${item.thumb}" loading="lazy" decoding="async" style="width:100%;height:${h}px;object-fit:cover;display:block;filter:brightness(.85)"/><div class="pi-ov"><span>+</span></div>`;
    }
    const media = card.querySelector('img,video');
    card.addEventListener('mouseenter', () => { media.style.transform='scale(1.04)'; media.style.filter='brightness(.5)'; media.style.transition='transform .5s,filter .3s'; });
    card.addEventListener('mouseleave', () => { media.style.transform=''; media.style.filter='brightness(.85)'; });
    card.onclick = () => openLb(origIdx);

    // only autoplay/decode when actually scrolled into view
    if(item.isVthumb){ _videoVisObserver.observe(media); }

    let t = 0;
    for(let i = 1; i < cols; i++) if(colH[i] < colH[t]) t = i;
    colH[t] += h + GAP;
    colEls[t].appendChild(card);
  });

  requestAnimationFrame(() => requestAnimationFrame(() => { grid.style.opacity='1'; }));
}

// ── GO HOME with cinematic back transition ──
function goHome(fromPop=false){
  const grid=document.querySelector('.photo-grid');

  if(grid) grid.classList.remove('in');

  if(fromPop){
    $wrap.style.transition='opacity .22s ease';
    $wrap.style.opacity='0';
    setTimeout(()=>{
      buildMetro(true);
      showPg('home',true);
      requestAnimationFrame(()=>{ $wrap.style.opacity='1'; });
    },160);
    setTimeout(()=>{ $wrap.style.transition=''; },500);
  } else {
    $trans.style.transition='opacity .35s ease';
    $trans.classList.add('fade-in');

    setTimeout(()=>{
      buildMetro(true);
      showPg('home',true);
    },380);

    setTimeout(()=>{
      $trans.style.transition='opacity .6s ease';
      $trans.classList.remove('fade-in');
    },600);

    setTimeout(()=>{
      $trans.style.transition='';
      $transTitle.textContent='';
    },1300);
  }

  document.querySelectorAll('#sidebar nav a').forEach(a=>a.classList.remove('act'));
  document.querySelector('#sidebar nav a[onclick*="\'home\'"]')?.classList.add('act');

  if(!fromPop){
    history.pushState({page:'home'}, '', '#');
  }
}

// ── PAGES ──────────────────────────────────
function showPg(name, resetTopbar=true){
  document.querySelectorAll('.pg').forEach(p=>p.classList.remove('act'));
  document.getElementById('pg-'+name).classList.add('act');
  if(name!=='gallery') document.getElementById('pg-'+name).scrollTop=0;
  if(resetTopbar){
    $tbBack.style.display='none';
    $tbSep.style.display='none';
    $tbCat.style.display='none';
  }
  closeSB();
}

function nav(name){
  if(name==='home'){goHome();return}
  showPg(name,true);
  document.querySelectorAll('#sidebar nav a').forEach(a=>a.classList.remove('act'));
  const t=document.querySelector(`#sidebar nav a[onclick*="'${name}'"]`);
  if(t)t.classList.add('act');
}

// ── SIDEBAR ────────────────────────────────
function toggleSB(){
  $sidebar.classList.contains('open')?closeSB():openSB();
}
function openSB(){
  $sidebar.classList.add('open');
  $hbg.classList.add('open');
  $dim.classList.add('on');
  $wrap.classList.add('pushed');
}
function closeSB(){
  $sidebar.classList.remove('open');
  $hbg.classList.remove('open');
  $dim.classList.remove('on');
  $wrap.classList.remove('pushed');
}

// ── LIGHTBOX ───────────────────────────────
function openLb(i){
  lbIdx=i;renderLb();
  if(window.__lbResetZoom)window.__lbResetZoom(false);

  $lb.animate(
    [{transform:'scale(.97)',opacity:0},{transform:'scale(1)',opacity:1}],
    {duration:350, easing:'cubic-bezier(.25,.8,.25,1)'}
  );
  $lb.classList.add('on');
  document.body.style.overflow='hidden';
}
function renderLb(){
  const item=filtItems[lbIdx];
  if(!item)return;

  const isV = item.type==='video';

  $lbImg.src=''; $lbImg.style.opacity='0';
  $lbThumb.src=''; $lbThumb.style.opacity='0';

  if(isV){
    $lbImg.style.display='none';
    $lbThumb.style.display='none';
    $lbVid.style.display='block';
    $lbVid.src='public/'+item.hq;
    $lbVid.play();
  } else {
    $lbVid.style.display='none'; $lbVid.pause(); $lbVid.src='';
    $lbImg.style.display='block';
    $lbThumb.style.display='block';

    $lbThumb.onload = () => { $lbThumb.style.opacity='1'; };
    $lbThumb.src = 'public/'+item.thumb;

    const full = new Image();
    const targetHq = 'public/'+item.hq;
    full.onload = () => {
      if(filtItems[lbIdx] !== item) return;
      $lbImg.src = targetHq;
      $lbImg.style.opacity = '1';
      $lbThumb.style.opacity = '0';
    };
    full.onerror = () => {
      if(filtItems[lbIdx] !== item) return;
      $lbThumb.style.opacity = '1';
    };
    full.src = targetHq;
  }

  $lbMt.textContent = item.title||'';
  $lbMd.textContent = [item.year,item.camera,item.lens].filter(Boolean).join(' · ');
  $lbCtr.textContent = `${lbIdx+1} / ${filtItems.length}`;
}
function closeLb(){
  $lb.animate(
    [{transform:'scale(1)',opacity:1},{transform:'scale(.97)',opacity:0}],
    {duration:250, easing:'ease'}
  );
  setTimeout(() => {
    $lb.classList.remove('on');
    $lbVid.pause();
    $lbVid.src = '';
    document.body.style.overflow = '';
  },220);
}
function lbNav(d){
  const stageEls=[$lbImg,$lbThumb,$lbVid];
  if(window.__lbResetZoom)window.__lbResetZoom(false);

  stageEls.forEach(el=>{
    el.style.transition='transform .22s cubic-bezier(.4,0,.2,1)';
    el.style.transform=`translateX(${d>0?-40:40}px)`;
  });

  setTimeout(()=>{
    lbIdx=(lbIdx+d+filtItems.length)%filtItems.length;
    renderLb();

    stageEls.forEach(el=>{
      el.style.transition='none';
      el.style.transform=`translateX(${d>0?40:-40}px)`;
    });
    void stageEls[0].offsetWidth;
    stageEls.forEach(el=>{
      el.style.transition='transform .3s cubic-bezier(.4,0,.2,1)';
      el.style.transform='translateX(0)';
    });
  },180);
}
document.addEventListener('keydown',e=>{
  if(!$lb.classList.contains('on'))return;
  if(e.key==='Escape')closeLb();
  if(e.key==='ArrowLeft')lbNav(-1);
  if(e.key==='ArrowRight')lbNav(1);
});
$lb.addEventListener('wheel',e=>{
  lbNav(e.deltaY>0?1:-1);
},{passive:true});
$lb.addEventListener('click',e=>{
  if(e.target===e.currentTarget)closeLb();
});

// pinch-to-zoom + pan + swipe-navigation for the lightbox image
(function(){
  const img=$lbImg;
  let scale=1, panX=0, panY=0;
  let pinchStartDist=0, pinchStartScale=1;
  let isPinching=false, isPanning=false;
  let panStartX=0, panStartY=0, startPanX=0, startPanY=0;
  let swipeStartX=0, swipeStartY=0;
  let lastTap=0;

  function applyTransform(smooth){
    img.style.transition = smooth ? 'opacity .4s ease, transform .25s ease' : 'opacity .4s ease';
    img.style.transform = `translate(${panX}px, ${panY}px) scale(${scale})`;
    img.style.cursor = scale>1 ? 'grab' : 'zoom-in';
    if(smooth) setTimeout(()=>{ img.style.transition='opacity .4s ease'; },250);
  }
  function resetZoom(smooth){
    scale=1; panX=0; panY=0;
    applyTransform(smooth!==false);
  }
  window.__lbResetZoom = resetZoom;

  function dist(a,b){ return Math.hypot(a.clientX-b.clientX, a.clientY-b.clientY); }

  $lb.addEventListener('touchstart',e=>{
    if(e.touches.length===2){
      isPinching=true; isPanning=false;
      pinchStartDist=dist(e.touches[0],e.touches[1]);
      pinchStartScale=scale;
      return;
    }
    if(e.touches.length===1){
      const t=e.touches[0];
      if(scale>1){
        isPanning=true;
        panStartX=t.clientX; panStartY=t.clientY;
        startPanX=panX; startPanY=panY;
      } else {
        swipeStartX=t.clientX; swipeStartY=t.clientY;
      }
      const now=Date.now();
      if(now-lastTap<300){
        if(scale>1){ resetZoom(true); }
        else { scale=2.4; applyTransform(true); }
        isPanning=false;
      }
      lastTap=now;
    }
  },{passive:true});

  $lb.addEventListener('touchmove',e=>{
    if(isPinching && e.touches.length===2){
      e.preventDefault();
      const d=dist(e.touches[0],e.touches[1]);
      scale=Math.min(4,Math.max(1,pinchStartScale*(d/pinchStartDist)));
      applyTransform(false);
    } else if(isPanning && e.touches.length===1){
      e.preventDefault();
      const t=e.touches[0];
      panX=startPanX+(t.clientX-panStartX);
      panY=startPanY+(t.clientY-panStartY);
      applyTransform(false);
    }
  },{passive:false});

  $lb.addEventListener('touchend',e=>{
    if(isPinching){
      isPinching=false;
      if(scale<1.02)resetZoom(true);
      return;
    }
    if(isPanning){
      isPanning=false;
      return;
    }
    if(scale>1)return;
    if(e.touches.length>0)return;
    const dx=e.changedTouches[0].clientX-swipeStartX;
    const dy=e.changedTouches[0].clientY-swipeStartY;
    if(Math.abs(dx)>50 && Math.abs(dx)>Math.abs(dy)*1.5){
      lbNav(dx<0?1:-1);
    }
  },{passive:true});

  let mouseDown=false, dragMoved=false;
  let mStartX=0, mStartY=0, mStartPanX=0, mStartPanY=0;

  img.addEventListener('mousedown',e=>{
    if(scale<=1)return;
    mouseDown=true; dragMoved=false;
    mStartX=e.clientX; mStartY=e.clientY;
    mStartPanX=panX; mStartPanY=panY;
    e.preventDefault();
  });
  window.addEventListener('mousemove',e=>{
    if(!mouseDown)return;
    const dx=e.clientX-mStartX, dy=e.clientY-mStartY;
    if(Math.abs(dx)>3||Math.abs(dy)>3)dragMoved=true;
    panX=mStartPanX+dx;
    panY=mStartPanY+dy;
    applyTransform(false);
  });
  window.addEventListener('mouseup',()=>{
    if(mouseDown){ mouseDown=false; img.style.cursor=scale>1?'grab':'zoom-in'; }
  });

  img.addEventListener('click',()=>{
    if(dragMoved){ dragMoved=false; return; }
    if(scale>1){ resetZoom(true); }
    else { scale=2.4; applyTransform(true); }
  });
})();

window.addEventListener('popstate', e => {
  if(!e.state || e.state.page === 'home'){
    goHome(true);
    return;
  }
  if(e.state.page === 'gallery'){
    openCat(e.state.cat, true);
  }
});
