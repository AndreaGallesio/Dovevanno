const CATS=["Bar e caffè","Ristoranti e cibo","Tabacchi","Spesa","Trasporti","Shopping","Persone","Risparmi e investimenti","Ricariche e banca","Altro"];
const COL=["#7c6cff","#ff8a3d","#12b5a5","#f0477a","#2d9cff","#ffc233","#9bdc4d","#6f6c92","#5a5778","#b65cf5"];
const NONSPEND=["Risparmi e investimenti","Ricariche e banca"];
const IC={"Bar e caffè":"M17 8h1a4 4 0 0 1 0 8h-1M3 8h14v9a4 4 0 0 1-4 4H7a4 4 0 0 1-4-4ZM6 2v3M10 2v3M14 2v3","Ristoranti e cibo":"M3 2v7a2 2 0 0 0 2 2h4a2 2 0 0 0 2-2V2M7 2v20M21 15V2a5 5 0 0 0-5 5v6a2 2 0 0 0 2 2h3Zm0 0v7","Tabacchi":"M2 14h14v4H2zM20 14v4M18 10c0-2 2-2 2-5","Spesa":"M3 4h2l2.4 10.2a2 2 0 0 0 2 1.5h7.7a2 2 0 0 0 2-1.5L21 8H6M9 20h.01M18 20h.01","Trasporti":"M5 17h14M6 17V8a3 3 0 0 1 3-3h6a3 3 0 0 1 3 3v9M8 21v-4M16 21v-4M9 12h6","Shopping":"M6 7h12l1 13H5ZM9 7a3 3 0 0 1 6 0","Persone":"M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2M9 11a4 4 0 1 0 0-8 4 4 0 0 0 0 8M22 21v-2a4 4 0 0 0-3-3.9M16 3.1a4 4 0 0 1 0 7.8","Risparmi e investimenti":"M12 21a9 9 0 1 0 0-18 9 9 0 0 0 0 18M8 12h6M9 9.5a3.5 3.5 0 0 1 5-1M9 14.5a3.5 3.5 0 0 0 5 1","Ricariche e banca":"M3 10l9-6 9 6M5 10v8M9 10v8M15 10v8M19 10v8M3 21h18","Altro":"M5 12h.01M12 12h.01M19 12h.01"};
function ico(c){
  const k=IC[c]?c:"Altro",color=c==="Da confermare"?"#b9a36a":col(c);
  const e=h("span",{class:"ic",style:"color:"+color+";background:"+color+"26"});
  e.innerHTML='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="'+IC[k]+'"/></svg>';
  return e;
}
const col=c=>COL[Math.max(0,CATS.indexOf(c))];
const SEED={"Caffè Portici":"Bar e caffè","SumUp  *Cafe Noir":"Bar e caffè","AvGvstA":"Bar e caffè","KOMOREBI IMMOBILIARE S":"Bar e caffè","SUSHI STAR SRL":"Ristoranti e cibo","I Particular Osteria":"Ristoranti e cibo","MC DONALD'S":"Ristoranti e cibo","Qualeaty":"Ristoranti e cibo","Tabaccheria Calandri Domenica":"Tabacchi"};
const $=id=>document.getElementById(id);
const eur=n=>n.toLocaleString("it-IT",{style:"currency",currency:"EUR"});
const eur0=n=>n.toLocaleString("it-IT",{style:"currency",currency:"EUR",maximumFractionDigits:0});
const norm=s=>String(s==null?"":s).trim();
const h=(t,a={},...c)=>{const e=document.createElement(t);for(const[k,v]of Object.entries(a)){if(k==="class")e.className=v;else if(k.startsWith("on"))e[k]=v;else e.setAttribute(k,v)}for(const x of c.flat())if(x!=null)e.append(x.nodeType?x:document.createTextNode(x));return e};
const mLabel=m=>{if(m[0]==="w")return weekLabel(m);if(m[0]==="y")return "Anno "+m.slice(1);const[y,mo]=m.split("-");return new Date(+y,+mo-1,1).toLocaleDateString("it-IT",{month:"long",year:"numeric"})};
const mShort=m=>{const[y,mo]=m.split("-");return new Date(+y,+mo-1,1).toLocaleDateString("it-IT",{month:"short"}).replace(".","")};

let tx=new Map(),learned={},store=null,period="all",open=null,pending={},aiSet=new Set();
let src="satispay";try{src=localStorage.getItem("sp-src")||"satispay"}catch(e){}
if(src!=="tot"&&!("satispay,contanti".split(",").includes(src)))src="satispay";
const SRCN={satispay:"Satispay",contanti:"Contanti"};
const SRCS=[["tot","Totale"],["satispay","Satispay"],["contanti","Contanti"]];
const SRCUI={
  tot:{intro:"Qui vedi insieme Satispay e contanti. Scegli un conto per caricare i movimenti."},
  satispay:{intro:"Carica l’export dei movimenti di Satispay (file .xlsx), anche ogni settimana: i movimenti già presenti vengono saltati e la tabella per categoria si crea da sola.",imp:"Aggiungi un export di Satispay (.xlsx)",btn:"+ Carica spese"},
  contanti:{intro:"Qui inserisci a mano le spese in contanti: tocca il pulsante qui sotto."}
};
const srcOf=t=>t.s||"satispay";
const T=()=>src==="tot"?[...tx.values()]:[...tx.values()].filter(t=>srcOf(t)===src);

/* ---- archivio: account (db) oppure solo questo browser ---- */
async function initStore(){store=null;try{if(navigator.storage&&navigator.storage.persist)await navigator.storage.persist()}catch(e){}}
async function loadAll(){
  if(store){
    try{
      const s=await store.get();
      for(const d of s.docs){const x=d.data()||{};if(x.map)learned=x.map;if(x.tx)for(const t of x.tx)tx.set(t.id,t)}
      return;
    }catch(e){store=null}
  }
  try{
    for(const t of JSON.parse(localStorage.getItem("sp-tx")||"[]"))tx.set(t.id,t);
    learned=JSON.parse(localStorage.getItem("sp-rules")||"{}");
  }catch(e){}
}
async function persist(months){
  try{
    if(store){
      for(const m of months)await store.doc("m-"+m).set({tx:[...tx.values()].filter(t=>t.d.slice(0,7)===m)});
    }else localStorage.setItem("sp-tx",JSON.stringify([...tx.values()]));
  }catch(e){showErr("Non sono riuscito a salvare i dati. Li vedi ora ma potrebbero non restare.")}
}
async function persistRules(){
  try{
    if(store)await store.doc("rules").set({map:learned});
    else localStorage.setItem("sp-rules",JSON.stringify(learned));
  }catch(e){showErr("Non sono riuscito a salvare le categorie.")}
}
function showErr(t){const e=$("err");e.textContent=t;e.hidden=!t}

/* ---- categorie ---- */
function byTipo(t){
  if(/Persona/i.test(t))return "Persone";
  if(/Investimento/i.test(t))return "Risparmi e investimenti";
  if(/Deposito Risparmi/i.test(t))return "Altro";
  if(/Banca|Bonifico/i.test(t))return "Ricariche e banca";
  return null;
}
function categoria(t){
  return t.c||learned[t.n]||SEED[t.n]||(/Salvadanaio/i.test(t.n)?"Risparmi e investimenti":null)||byTipo(t.t)||null;
}

/* ---- lettura file ---- */
const pad=n=>String(n).padStart(2,"0");
function fmtDate(v){
  if(v instanceof Date)return v.getFullYear()+"-"+pad(v.getMonth()+1)+"-"+pad(v.getDate())+" "+pad(v.getHours())+":"+pad(v.getMinutes());
  return String(v).replace("T"," ").slice(0,16);
}
function parse(wb){
  const ws=wb.Sheets["Transactions"]||wb.Sheets[wb.SheetNames[0]];
  const out=[];
  for(const d of XLSX.utils.sheet_to_json(ws,{defval:null,raw:true})){
    const i=Number(d["Importo"]);
    if(!isFinite(i)||d["Nome"]==null||d["Data"]==null)continue;
    if(d["Stato"]&&/Annullato|corso/i.test(String(d["Stato"])))continue;
    const idk=Object.keys(d).find(k=>k.startsWith("ID"));
    const dt=fmtDate(d["Data"]);
    const n=norm(d["Nome"]);
    out.push({id:idk&&d[idk]?String(d[idk]):dt+"|"+n+"|"+i,d:dt,n,i,t:norm(d["Tipo"])});
  }
  return out;
}
async function load(file){
  showErr("");$("msg").hidden=true;
  try{
    const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
    const rows=parse(wb);
    if(!rows.length)throw new Error("Non trovo movimenti nel file. Serve l’export di Satispay con le colonne Nome, Importo e Tipo.");
    let add=0,dup=0;const months=new Set();
    for(const r of rows){if(tx.has(r.id)){dup++;continue}tx.set(r.id,r);months.add(r.d.slice(0,7));add++}
    await persist(months);
    const m=$("msg");m.textContent=add+" movimenti aggiunti"+(dup?", "+dup+" già presenti":"")+".";m.hidden=false;
    if(add){const ds=rows.map(r=>r.d.slice(0,10)).sort();period="w"+mondayOf(ds[ds.length-1])}
    open=null;render();
    if(window.gsap&&!RM)gsap.fromTo("#hero",{scale:.96},{scale:1,duration:.9,ease:"elastic.out(1,.45)"});
  }catch(e){showErr(e.message||"Non riesco a leggere il file.")}
  $("file").value="";
}
$("file").addEventListener("change",e=>{const f=e.target.files[0];if(f)load(f)});

/* ---- calcoli ---- */
function mondayOf(d){const x=new Date(+d.slice(0,4),+d.slice(5,7)-1,+d.slice(8,10));x.setDate(x.getDate()-(x.getDay()+6)%7);return x.getFullYear()+"-"+pad(x.getMonth()+1)+"-"+pad(x.getDate())}
const weekKey=t=>"w"+mondayOf(t.d.slice(0,10));
function wkAdd(iso,n){const x=new Date(+iso.slice(0,4),+iso.slice(5,7)-1,+iso.slice(8,10));x.setDate(x.getDate()+n);return x.getFullYear()+"-"+pad(x.getMonth()+1)+"-"+pad(x.getDate())}
function weekLabel(k){
  const a=new Date(+k.slice(1,5),+k.slice(6,8)-1,+k.slice(9,11)),b=new Date(a);b.setDate(b.getDate()+6);
  const mo=d=>d.toLocaleDateString("it-IT",{month:"short"}).replace(".","");
  return a.getMonth()===b.getMonth()?a.getDate()+"–"+b.getDate()+" "+mo(b)+" "+b.getFullYear():a.getDate()+" "+mo(a)+" – "+b.getDate()+" "+mo(b)+" "+b.getFullYear();
}
const allWeeks=()=>[...new Set(T().map(weekKey))].sort();
const inKey=(t,k)=>k[0]==="w"?weekKey(t)===k:t.d.slice(0,7)===k;
function prevPeriod(){
  if(/^\d{4}-\d{2}$/.test(period)){const m=allMonths(),i=m.indexOf(period);return i>0?m[i-1]:null}
  if(period[0]==="w"){const w=allWeeks(),i=w.indexOf(period);return i>0?w[i-1]:null}
  return null;
}
const inPeriod=t=>period==="all"||(period[0]==="w"?weekKey(t)===period:period[0]==="y"?t.d.slice(0,4)===period.slice(1):t.d.slice(0,7)===period);
const allMonths=()=>[...new Set(T().map(t=>t.d.slice(0,7)))].sort();
function totals(list){
  const o={};
  for(const t of list){const c=categoria(t);if(t.i<0)o[c||"Da confermare"]=(o[c||"Da confermare"]||0)-t.i}
  return Object.entries(o).sort((a,b)=>b[1]-a[1]);
}

/* ---- grafici ---- */
function donut(tot,all){
  const R=50,P=a=>(70+R*Math.cos(a)).toFixed(2)+","+(70+R*Math.sin(a)).toFixed(2);
  let a0=-Math.PI/2;
  let s='<svg viewBox="0 0 140 140" width="150" height="150" role="img" aria-label="Uscite per categoria">';
  for(const[c,v]of tot){
    const sw=v/all*2*Math.PI,gap=tot.length>1?.04:0,st=a0+gap/2,en=a0+sw-gap/2,color=c==="Da confermare"?"#b9a36a":col(c);
    const at='class="seg" data-c="'+c+'" pathLength="100" fill="none" stroke="'+color+'" stroke-width="22" style="cursor:pointer;stroke-dasharray:100"';
    if(sw>6.28)s+='<circle '+at+' cx="70" cy="70" r="50"/>';
    else if(en>st)s+='<path '+at+' d="M'+P(st)+"A"+R+","+R+" 0 "+(en-st>Math.PI?1:0)+" 1 "+P(en)+'"/>';
    a0+=sw;
  }
  return s+'<text x="70" y="67" text-anchor="middle" font-size="15" font-weight="800" style="fill:var(--ink)">'+eur0(all)+'</text><text x="70" y="82" text-anchor="middle" font-size="8" style="fill:var(--mute)">Totale</text></svg>';
}
function bars(catFilter){
  const ms=allMonths().slice(-12);
  if(ms.length<2)return null;
  const data=ms.map(m=>{const o={};for(const t of T())if(t.i<0&&t.d.slice(0,7)===m){const c=categoria(t)||"Da confermare";if(!catFilter||c===catFilter)o[c]=(o[c]||0)-t.i}return o});
  const sums=data.map(o=>Object.values(o).reduce((a,b)=>a+b,0));
  const max=Math.max(...sums,1),W=320,H=150,top=14,bot=20,bw=Math.min(34,(W-16)/ms.length*.66);
  let s='<svg viewBox="0 0 '+W+" "+H+'" width="100%" role="img" aria-label="Uscite mese per mese">';
  ms.forEach((m,k)=>{
    const x=8+(W-16)/ms.length*(k+.5)-bw/2;let y=H-bot;
    const dim=period!=="all"&&period[0]!=="y"&&period[0]!=="w"&&period!==m?' opacity=".35"':"";
    s+='<g data-m="'+m+'" style="cursor:pointer"'+dim+'>';
    const keys=catFilter?[catFilter]:[...CATS,"Da confermare"];
    for(const c of keys){const v=data[k][c];if(!v)continue;const hh=v/max*(H-top-bot);y-=hh;s+='<rect class="gb" style="animation-delay:'+(k*45)+'ms" x="'+x.toFixed(1)+'" y="'+y.toFixed(1)+'" width="'+bw.toFixed(1)+'" height="'+hh.toFixed(1)+'" rx="2" fill="'+(c==="Da confermare"?"#b9a36a":catFilter?"var(--acc)":col(c))+'"/>'}
    s+='<text x="'+(x+bw/2).toFixed(1)+'" y="'+(y-3).toFixed(1)+'" text-anchor="middle" font-size="8" style="fill:var(--mute)">'+Math.round(sums[k])+'</text>';
    s+='<text x="'+(x+bw/2).toFixed(1)+'" y="'+(H-6)+'" text-anchor="middle" font-size="9" style="fill:var(--ink)">'+mShort(m)+'</text>';
    s+='<rect x="'+(x-4).toFixed(1)+'" y="0" width="'+(bw+8).toFixed(1)+'" height="'+H+'" fill="transparent"/></g>';
  });
  return s+"</svg>";
}

const RM=matchMedia("(prefers-reduced-motion: reduce)").matches;
function countUp(el,to){
  if(RM||!el)return;
  const t0=performance.now(),D=900;
  const f=now=>{const p=Math.max(0,Math.min(1,(now-t0)/D));el.textContent=eur(to*(1-Math.pow(1-p,3)));if(p<1)requestAnimationFrame(f)};
  requestAnimationFrame(f);
}

/* ---- schermata ---- */
function renderSrc(){
  const b=$("srcs");b.textContent="";
  for(const[k,l]of SRCS)b.append(h("button",{type:"button",class:"sg"+(k===src?" on":""),onclick:()=>{if(k===src)return;src=k;try{localStorage.setItem("sp-src",k)}catch(e){}open=null;pending={};aiSet=new Set();showErr("");$("msg").hidden=true;render()}},l));
}
function render(){
  renderSrc();
  const list=T(),ui=SRCUI[src];
  $("intro").hidden=list.length>0;$("intro").textContent=ui.intro;
  $("result").hidden=!list.length;
  $("addm0").hidden=list.length>0;$("addm0").textContent=src==="contanti"?"+ Aggiungi una spesa in contanti":"+ Aggiungi una spesa a mano";
  const dz=$("drop");dz.classList.toggle("mini",list.length>0||!ui.imp);
  if(ui.imp)dz.querySelector("span").textContent=ui.imp;
  if(!list.length)return;
  const ms=allMonths();
  if(period[0]==="w"){const wv=allWeeks();if(!wv.includes(period))period=wv[wv.length-1]||"all"}
  else if(period!=="all"&&period[0]!=="y"&&!ms.includes(period))period=ms[ms.length-1];
  const wks=allWeeks(),years=[...new Set(ms.map(m=>m.slice(0,4)))];
  const mode=period==="all"?"all":period[0]==="y"?"year":period[0]==="w"?"week":"month";
  const cap=x=>x[0].toUpperCase()+x.slice(1);
  const yrOf=mode==="week"?period.slice(1,5):mode==="month"?period.slice(0,4):mode==="year"?period.slice(1):years[years.length-1];
  const setMode=m=>{
    if(m===mode)return;
    if(m==="all")period="all";
    else if(m==="year")period="y"+yrOf;
    else if(m==="month")period=mode==="week"&&ms.includes(period.slice(1,8))?period.slice(1,8):(ms.filter(x=>x.startsWith(yrOf)).pop()||ms[ms.length-1]);
    else period=(mode==="month"?wks.filter(w=>w.slice(1,8)===period).pop():wks.filter(w=>w.slice(1,5)===yrOf).pop())||wks[wks.length-1];
    open=null;navDir=0;render();
  };
  const seg=$("seg3");seg.textContent="";
  for(const[m,l]of [["week","Settimana"],["month","Mese"],["year","Anno"],["all","Tutto"]])seg.append(h("button",{type:"button",class:"sg"+(m===mode?" on":""),onclick:()=>setMode(m)},l));
  const nav=$("nav");nav.textContent="";nav.hidden=mode==="all";navGo=null;
  if(mode!=="all"){
    const opts=mode==="week"?wks:mode==="month"?ms:years.map(y=>"y"+y),i=opts.indexOf(period);
    const lab=p=>mode==="week"?weekLabel(p):mode==="month"?cap(mLabel(p)):p.slice(1);
    const go=d=>{const j=i+d;if(j<0||j>=opts.length)return;period=opts[j];open=null;navDir=d;render()};
    navGo=go;
    const sel=h("select",{class:"ghostsel","aria-label":"Scegli il periodo"});
    opts.slice().reverse().forEach(p=>sel.add(new Option(lab(p),p)));
    sel.value=period;sel.onchange=()=>{period=sel.value;open=null;navDir=0;render()};
    const prev=h("button",{type:"button",class:"na","aria-label":"Periodo precedente",onclick:()=>go(-1)},"‹"),next=h("button",{type:"button",class:"na","aria-label":"Periodo successivo",onclick:()=>go(1)},"›");
    prev.disabled=i<=0;next.disabled=i>=opts.length-1;
    nav.append(prev,h("div",{class:"nl"},h("span",{id:"navL"},lab(period)),h("span",{class:"chev"},"▾"),sel),next);
    if(navDir&&window.gsap&&!RM)gsap.from("#navL",{x:navDir*30,opacity:0,duration:.4,ease:"power3.out"});
    navDir=0;
  }

  const cur=list.filter(inPeriod),tot=totals(cur),all=tot.reduce((a,[,v])=>a+v,0);
  let cmp=null;
  const pp=prevPeriod();
  if(pp){const pt=agg(list.filter(t=>inKey(t,pp))).out;if(pt){const d=(all-pt)/pt*100;cmp=(d>=0?"▲ +":"▼ ")+d.toFixed(0)+"% rispetto "+(pp[0]==="w"?(pp.slice(1)===wkAdd(period.slice(1),-7)?"alla settimana scorsa":"a "+weekLabel(pp)):"a "+mShort(pp))}}
  const hero=$("heroIn");hero.textContent="";const real=tot.filter(([c])=>!NONSPEND.includes(c)&&c!=="Da confermare").reduce((a,[,v])=>a+v,0);const sav=(tot.find(x=>x[0]==="Risparmi e investimenti")||[0,0])[1];
  hero.append(h("div",{class:"mute"},period==="all"?"Uscite totali":(period[0]==="y"?"Uscite "+mLabel(period).toLowerCase():period[0]==="w"?"Uscite della settimana "+mLabel(period):"Uscite di "+mLabel(period))),h("div",{class:"big"},eur(all)),
    h("div",{class:"mrow"},cmp?h("span",{class:"pill"},cmp):null,h("span",{class:"mute"},(sav?"di cui risparmi "+eur(sav)+" · ":"")+cur.filter(t=>t.i<0).length+" pagamenti")),
    h("div",{class:"hbtns"},SRCUI[src].btn?h("button",{type:"button",class:"gbtn",onclick:()=>$("file").click()},SRCUI[src].btn):null,h("button",{type:"button",class:"gbtn",onclick:openForm},"+ Spesa a mano")));
  countUp(hero.querySelector(".big"),all);


  const t=$("table");t.textContent="";let ri=0;
  for(const[c,v]of [...tot].sort((a,b)=>NONSPEND.includes(a[0])-NONSPEND.includes(b[0]))){
    const known=c!=="Da confermare";
    t.append(h("button",{type:"button",class:"row"+(NONSPEND.includes(c)?" soft":""),"aria-label":c+" "+eur(v),onclick:()=>{if(known){open=c;render()}}},
      ico(c),h("b",{},c),h("span",{class:"num amt","data-v":String(v)},eur(v)),h("span",{class:"mute num"},(v/all*100).toFixed(0)+"%"),h("span",{class:"rbar"},h("i",{style:"width:"+(v/all*100).toFixed(1)+"%;background:"+(known?col(c):"#b9a36a")}))));
  }
  t.append(h("div",{class:"tot"},h("span",{},"Totale uscite"),h("span",{},eur(all))));
  renderDetail(cur);
  renderUnknown();
  renderTx();renderStats(cur,list);renderBrk(tot,all);renderCharts(cur);
  animateIn();
}

function closeSheet(){
  if(!open)return;
  if(window.gsap&&!RM){gsap.to("#detail",{yPercent:100,duration:.4,ease:"power3.in"});gsap.to("#scrim",{opacity:0,duration:.4,onComplete:()=>{open=null;render()}})}
  else{open=null;render()}
}
let sheetOpen=false;
function agg(l){let inc=0,out=0,sav=0;for(const t of l){if(t.i>0)inc+=t.i;else{out-=t.i;if(categoria(t)==="Risparmi e investimenti")sav-=t.i}}return{inc,out,sav}}
function pctTxt(a,b){if(!b)return null;const d=(a-b)/b*100;return (d>=0?"▲ +":"▼ ")+d.toFixed(0)+"%"}
function renderStats(cur,list){
  const a=agg(cur);let p=null;
  const pq=prevPeriod();if(pq)p=agg(list.filter(t=>inKey(t,pq)));
  const f=v=>v>=10000?eur0(v):eur(v);
  const defs=[["Entrate",a.inc,p&&p.inc,"#4ade80","M3 17l6-6 4 4 8-8M15 7h6v6"],["Uscite",a.out,p&&p.out,"#ff5d8f","M3 7h18v10H3zM3 11h18"],["Risparmi",a.sav,p&&p.sav,"#5aa2ff",IC["Risparmi e investimenti"]]];
  const box=$("stats");box.textContent="";
  for(const[n,v,pv,c,path]of defs){
    const e=h("div",{class:"st",style:"--c:"+c});
    e.innerHTML='<span class="sic"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="'+path+'"/></svg></span>';
    e.append(h("div",{class:"mute"},n),h("div",{class:"sv amt","data-v":String(v)},f(v)),h("div",{class:"mute"},(p&&pctTxt(v,pv))||"nel periodo"));
    box.append(e);
  }
}
function renderBrk(tot,all){
  const b=$("brk");b.textContent="";
  b.hidden=!all;if(!all)return;
  const lg=h("div",{});
  for(const[c,v]of tot.slice(0,6)){
    const known=c!=="Da confermare";
    lg.append(h("button",{type:"button",class:"lg",onclick:()=>{if(known){open=c;render()}}},h("span",{class:"dot",style:"background:"+(known?col(c):"#b9a36a")}),h("span",{class:"ln"},c),h("span",{class:"mute num"},(v/all*100).toFixed(0)+"%")));
  }
  const dn=h("div",{class:"dn"});dn.innerHTML=donut(tot,all);
  b.append(h("div",{class:"head"},h("b",{},"Ripartizione spese"),h("span",{class:"mute"},period==="all"?"Tutto":mLabel(period))),h("div",{class:"brkbody"},lg,dn));
}
function lineChart(cur){
  cur=cur.filter(t=>categoria(t)!=="Risparmi e investimenti");
  const isM=/^\d{4}-\d{2}$/.test(period),isW=period[0]==="w";
  if(!(isM||isW)||!cur.some(t=>t.i<0))return null;
  let nd,idx,labs;
  if(isW){
    nd=7;const m0=new Date(+period.slice(1,5),+period.slice(6,8)-1,+period.slice(9,11));
    idx=t=>Math.round((new Date(+t.d.slice(0,4),+t.d.slice(5,7)-1,+t.d.slice(8,10))-m0)/864e5);
    labs=["Lun","Mar","Mer","Gio","Ven","Sab","Dom"].map((l,i)=>[l,i]);
  }else{
    const[y,mo]=period.split("-").map(Number);nd=new Date(y,mo,0).getDate();
    idx=t=>+t.d.slice(8,10)-1;labs=[1,8,15,22,nd].map(d=>[d,d-1]);
  }
  const day=Array(nd).fill(0);
  for(const t of cur)if(t.i<0)day[idx(t)]-=t.i;
  const ld=Math.max(...cur.filter(t=>t.i<0).map(idx))+1;
  let a=0;const cum=day.slice(0,ld).map(v=>a+=v);
  const W=320,H=150,L=8,T=14,B=22,mx=Math.max(cum[ld-1],1)*1.08;
  const X=i=>L+(W-2*L)*(i/(nd-1)),Y=v=>H-B-v/mx*(H-T-B);
  const line="M"+cum.map((v,i)=>X(i).toFixed(1)+","+Y(v).toFixed(1)).join("L"),area=line+"L"+X(ld-1).toFixed(1)+","+(H-B)+"L"+X(0).toFixed(1)+","+(H-B)+"Z";
  let s='<svg viewBox="0 0 '+W+" "+H+'" width="100%" role="img" aria-label="Spesa accumulata"><defs><linearGradient id="ag" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#8b6dff" stop-opacity=".5"/><stop offset="1" stop-color="#8b6dff" stop-opacity="0"/></linearGradient></defs>';
  s+='<path class="ar" d="'+area+'" fill="url(#ag)"/><path class="ln" d="'+line+'" pathLength="1" fill="none" stroke="#a58bff" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" style="stroke-dasharray:1"/><circle cx="'+X(ld-1).toFixed(1)+'" cy="'+Y(cum[ld-1]).toFixed(1)+'" r="4" fill="#fff"/>';
  for(const[l,i]of labs)s+='<text x="'+X(i).toFixed(1)+'" y="'+(H-6)+'" text-anchor="middle" font-size="9" style="fill:var(--mute)">'+l+'</text>';
  return s+"</svg>";
}
function renderCharts(cur){
  const ln=lineChart(cur);$("lnT").textContent=period[0]==="w"?"Andamento nella settimana":"Andamento nel mese";$("lnT").hidden=$("lnc").hidden=!ln;if(ln)$("lnc").innerHTML=ln;
}
const TRASH='<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18M8 6V4h8v2M6 6l1 14h10l1-14M10 11v6M14 11v6"/></svg>';
function delBtn(t){
  const b=h("button",{type:"button",class:"xbtn","aria-label":"Elimina questa spesa"});
  b.innerHTML=TRASH;
  let armed=false,timer;
  b.onclick=async e=>{
    e.stopPropagation();
    if(!armed){armed=true;b.classList.add("armed");b.textContent="Elimina?";timer=setTimeout(()=>{armed=false;b.classList.remove("armed");b.innerHTML=TRASH},3000);return}
    clearTimeout(timer);
    tx.delete(t.id);
    await persist(new Set([t.d.slice(0,7)]));
    if(open&&![...tx.values()].some(x=>categoria(x)===open))open=null;
    render();
  };
  return b;
}
let formOpen=false,mType=-1;
const isoDay=d=>d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate());
function setType(k){mType=k;$("mOut").classList.toggle("on",k<0);$("mIn").classList.toggle("on",k>0)}
function openForm(){
  formOpen=true;setType(-1);
  $("mAmt").value="";$("mName").value="";$("mDate").value=isoDay(new Date());$("mCat").value="Altro";$("mSrc").value=src==="tot"?"contanti":src;$("mErr").hidden=true;
  $("mform").hidden=false;$("scrim").hidden=false;document.body.style.overflow="hidden";
  if(window.gsap&&!RM){gsap.fromTo("#mform",{yPercent:100},{yPercent:0,duration:.55,ease:"power4.out"});gsap.fromTo("#scrim",{opacity:0},{opacity:1,duration:.3})}
}
function closeForm(){
  const done=()=>{formOpen=false;$("mform").hidden=true;$("scrim").hidden=!open;document.body.style.overflow=open?"hidden":""};
  if(window.gsap&&!RM){gsap.to("#mform",{yPercent:100,duration:.35,ease:"power3.in",onComplete:done});if(!open)gsap.to("#scrim",{opacity:0,duration:.35})}
  else done();
}
async function saveForm(){
  const v=parseFloat(String($("mAmt").value).replace(",",".")),name=norm($("mName").value),day=$("mDate").value,e=$("mErr");
  const bad=m=>{e.textContent=m;e.hidden=false};
  if(!(v>0))return bad("Inserisci un importo maggiore di zero.");
  if(!name)return bad("Scrivi una descrizione.");
  if(!day)return bad("Scegli la data.");
  const now=new Date(),d=day+" "+pad(now.getHours())+":"+pad(now.getMinutes());
  const t={id:"man-"+Date.now(),d,n:name,i:mType*Math.round(v*100)/100,t:"Manuale",c:$("mCat").value,s:$("mSrc").value};
  tx.set(t.id,t);
  if(src!=="tot"&&src!==t.s){src=t.s;try{localStorage.setItem("sp-src",src)}catch(e){}}
  await persist(new Set([d.slice(0,7)]));
  closeForm();
  period="w"+mondayOf(d.slice(0,10));open=null;
  render();
  const m=$("msg");m.textContent="Spesa aggiunta.";m.hidden=false;
}
function renderTx(){
  const q=norm($("q").value).toLowerCase(),fc=$("fc").value,min=parseFloat($("fmin").value)||0,d1=$("fd1").value,d2=$("fd2").value;
  const rows=T().filter(inPeriod).filter(t=>{
    const c=categoria(t)||"Da confermare";
    if(q&&!(t.n.toLowerCase().includes(q)||c.toLowerCase().includes(q)))return false;
    if(fc&&c!==fc)return false;
    if(min&&Math.abs(t.i)<min)return false;
    const d=t.d.slice(0,10);
    return !(d1&&d<d1)&&!(d2&&d>d2);
  }).sort((a,b)=>b.d.localeCompare(a.d));
  $("fsum").textContent=rows.length+" pagamenti · saldo "+eur(rows.reduce((a,t)=>a+t.i,0));
  const box=$("flist");box.textContent="";
  for(const t of rows.slice(0,50)){
    const c=categoria(t)||"Da confermare",[,m,d]=t.d.slice(0,10).split("-");
    const man=String(t.id).startsWith("man-");
    const row=h("div",{class:"tx2"},ico(c),h("div",{},h("div",{},t.n),h("div",{class:"mute"},d+"/"+m+" · "+c+(src==="tot"?" · "+SRCN[srcOf(t)]:"")+(man?" · manuale":""))),h("span",{class:"num"+(t.i>0?" pos":"")},eur(t.i)));
    if(man){row.style.gridTemplateColumns="auto 1fr auto auto";row.append(delBtn(t))}
    box.append(row);
  }
  if(rows.length>50)box.append(h("p",{class:"mute"},"Mostro i primi 50: affina la ricerca."));
}
function renderDetail(cur){
  const box=$("detail");box.textContent="";
  const was=sheetOpen;sheetOpen=!!open;
  box.hidden=!open;$("scrim").hidden=!open;document.body.style.overflow=open?"hidden":"";
  if(!open)return;
  const rows=cur.filter(t=>categoria(t)===open).sort((a,b)=>b.d.localeCompare(a.d));
  const sum=rows.reduce((a,t)=>a+(t.i<0?-t.i:0),0);
  const g={};for(const t of rows)(g[t.n]=g[t.n]||{s:0,n:0}),g[t.n].s+=t.i,g[t.n].n++;
  box.append(h("h2",{},open),h("div",{class:"card"},
    h("div",{class:"head"},h("div",{},h("div",{class:"big",style:"font-size:24px;font-weight:800"},eur(sum)),h("div",{class:"mute"},rows.length+" movimenti"+(period==="all"?"":" in "+mLabel(period)))),h("button",{class:"ghost",type:"button",onclick:closeSheet},"Chiudi"))));
  const b=bars(open);
  if(b){const c=h("div",{class:"card",style:"margin-top:10px"});c.innerHTML=b;c.addEventListener("click",e=>{const x=e.target.closest("[data-m]");if(x){period=period===x.dataset.m?"all":x.dataset.m;render()}});box.append(c)}
  const mc=h("div",{class:"card",style:"margin-top:10px"});
  mc.append(h("b",{},"Per esercente"));
  for(const[n,v]of Object.entries(g).sort((a,b)=>Math.abs(b[1].s)-Math.abs(a[1].s))){
    const sel=h("select",{"aria-label":"Sposta "+n+" in un’altra categoria"});
    for(const c of CATS)sel.add(new Option(c,c));sel.value=open;
    sel.onchange=async()=>{learned[n]=sel.value;await persistRules();render()};
    mc.append(h("div",{class:"q"},h("div",{class:"top"},h("span",{},n+" ("+v.n+")"),h("span",{class:"num"},eur(-v.s))),sel));
  }
  box.append(mc);
  const lc=h("div",{class:"card",style:"margin-top:10px"});
  lc.append(h("b",{},"Pagamenti"),h("div",{class:"mute"},"Tocca un pagamento per modificarlo."));
  for(const t of rows.slice(0,60)){
    const [y,m,d]=t.d.slice(0,10).split("-");
    const r=h("div",{class:"tx"},h("span",{class:"mute"},d+"/"+m),h("span",{},t.n),h("span",{class:"num"},eur(t.i)));
    if(String(t.id).startsWith("man-")){r.style.gridTemplateColumns="52px 1fr auto auto";r.append(delBtn(t))}
    lc.append(r);
  }
  if(rows.length>60)lc.append(h("p",{class:"mute"},"Mostro i 60 più recenti."));
  box.append(lc);
  box.prepend(h("div",{class:"grab"}));
  if(window.gsap&&!RM&&!was){
    gsap.fromTo("#detail",{yPercent:100},{yPercent:0,duration:.6,ease:"power4.out"});
    gsap.fromTo("#scrim",{opacity:0},{opacity:1,duration:.4});
    gsap.from("#detail > *",{y:24,opacity:0,duration:.6,stagger:.08,ease:"power3.out",delay:.15});
    if(document.querySelector("#detail .gb"))gsap.from("#detail .gb",{scaleY:0,transformOrigin:"50% 100%",duration:.8,stagger:.05,delay:.3,ease:"power3.out"});
  }
}

/* ---- esercenti da confermare ---- */
function unknownMap(){
  const u={};for(const t of T())if(!categoria(t))u[t.n]=(u[t.n]||0)+t.i;return u;
}
function renderUnknown(){
  const u=unknownMap(),names=Object.keys(u);
  $("todo").hidden=!names.length;
  const box=$("unknown");box.textContent="";
  for(const n of names){
    const sel=h("select",{"aria-label":"Categoria per "+n},new Option("Scegli categoria…",""));
    for(const c of CATS)sel.add(new Option(c,c));
    sel.value=pending[n]||"";
    sel.onchange=()=>{pending[n]=sel.value;aiSet.delete(n)};
    box.append(h("div",{class:"q"},h("div",{class:"top"},h("span",{},h("b",{},n),aiSet.has(n)?h("span",{class:"tag"},"suggerito da Claude"):null),h("span",{class:"num"},eur(u[n]))),sel));
  }
}
$("save").onclick=async()=>{
  for(const[k,v]of Object.entries(pending))if(v)learned[k]=v;
  pending={};aiSet=new Set();await persistRules();render();
};

let navDir=0,navGo=null;
function animateIn(){
  if(!window.gsap||RM)return;
  const has=q=>document.querySelector(q);
  const tl=gsap.timeline({defaults:{ease:"power3.out"}});
  tl.from("#hero",{y:34,opacity:0,scale:.95,duration:.85},0);
  if(has("#brk .seg"))tl.fromTo("#brk .seg",{strokeDashoffset:100},{strokeDashoffset:0,duration:1.2,stagger:.09},.15);
  if(has("#stats .st"))tl.from("#stats .st",{y:26,opacity:0,scale:.94,duration:.6,stagger:.1},.2);
  if(has("#lnc .ln")){tl.fromTo("#lnc .ln",{strokeDashoffset:1},{strokeDashoffset:0,duration:1.4},.35);tl.from("#lnc .ar",{opacity:0,duration:1},.7)}
  if(has("#table .row"))tl.from("#table .row",{y:22,opacity:0,duration:.55,stagger:.07},.35);
  document.querySelectorAll("#table .amt,#stats .amt").forEach(e=>countUp(e,+e.dataset.v));
  if(has("#table .rbar i"))tl.from("#table .rbar i",{width:0,duration:1,stagger:.07},.5);
}
function intro(){
  if(!window.gsap||RM)return;
  gsap.from([".brand","#intro"],{y:28,opacity:0,duration:.8,stagger:.12,ease:"power3.out"});gsap.from("#drop",{opacity:0,duration:.8,delay:.25,ease:"power3.out"});
  gsap.to(".drop svg",{y:-4,duration:1.1,repeat:-1,yoyo:true,ease:"sine.inOut"});
  gsap.to(".o1",{x:-34,y:34,duration:7,repeat:-1,yoyo:true,ease:"sine.inOut"});
  gsap.to(".o2",{x:44,y:-24,duration:9,repeat:-1,yoyo:true,ease:"sine.inOut"});
  const io=new IntersectionObserver(es=>{for(const e of es)if(e.isIntersecting){io.unobserve(e.target);gsap.from(e.target,{y:40,opacity:0,duration:.85,ease:"power3.out"})}},{threshold:.12});
  document.querySelectorAll("#result h2,#table,#todo,#srch,#brk,#lnc,#stats").forEach(el=>io.observe(el));
}
let sx=0,sy=0;
$("hero").addEventListener("touchstart",e=>{sx=e.touches[0].clientX;sy=e.touches[0].clientY},{passive:true});
$("hero").addEventListener("touchend",e=>{const dx=e.changedTouches[0].clientX-sx,dy=e.changedTouches[0].clientY-sy;if(navGo&&Math.abs(dx)>60&&Math.abs(dy)<40)navGo(dx<0?1:-1)},{passive:true});
$("scrim").onclick=()=>{if(formOpen)closeForm();else closeSheet()};
for(const c of CATS)$("mCat").add(new Option(c,c));
$("mOut").onclick=()=>setType(-1);$("mIn").onclick=()=>setType(1);
$("mSave").onclick=saveForm;$("mCancel").onclick=closeForm;$("addm0").onclick=openForm;
for(const c of [...CATS,"Da confermare"])$("fc").add(new Option(c,c));
for(const id of ["q","fmin"])$(id).addEventListener("input",renderTx);
for(const id of ["fc","fd1","fd2"])$(id).addEventListener("change",renderTx);
$("brk").addEventListener("click",e=>{const g=e.target.closest("[data-c]");if(g&&g.dataset.c!=="Da confermare"){open=g.dataset.c;render()}});
(async()=>{await initStore();await loadAll();{const w0=allWeeks();if(w0.length)period=w0[w0.length-1]}for(const[k,t]of [...tx])if(t.s==="banca")tx.delete(k);render();intro()})();
