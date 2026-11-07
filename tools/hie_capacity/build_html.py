"""Render metrics.json into one self-contained HTML report (no external libraries)."""
import json, os

D = os.path.dirname(os.path.abspath(__file__))
OUT = '/home/user/Chase-s-Repository/reports/HIE_Capacity_Actuals.html'
m = json.load(open(os.path.join(D, 'metrics.json')))
cp = os.path.join(D, '..', 'cost', 'cost.json')
if os.path.exists(cp):
    c = json.load(open(cp)); m['cost'] = {k: c[k] for k in ('meta', 'yoy', 'index', 'lines', 'skus', 'spread')}
k = m['kpi']

HTML = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>HIE Capacity Actuals</title>
<style>
:root{color-scheme:light;--bg:#f6f6f4;--surface:#fcfcfb;--ink:#0b0b0b;--ink2:#52514e;--muted:#8a8984;--rule:#e4e3df;--grid:#ecebe7;
--s1:#2a78d6;--s2:#eb6834;--neutral:#c9c8c2;--o1:#86b6ef;--o2:#3987e5;--o3:#1c5cab;--o4:#0d366b;--crit:#d03b3b;--good:#0ca30c;--hl:#fff4d6}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){color-scheme:dark;--bg:#121211;--surface:#1a1a19;--ink:#f0efec;--ink2:#c3c2b7;--muted:#8f8e86;--rule:#2e2e2b;--grid:#262624;
--s1:#3987e5;--s2:#d95926;--neutral:#6a6964;--o1:#86b6ef;--o2:#5598e7;--o3:#2a78d6;--o4:#184f95;--hl:#3a3218}}
:root[data-theme="dark"]{color-scheme:dark;--bg:#121211;--surface:#1a1a19;--ink:#f0efec;--ink2:#c3c2b7;--muted:#8f8e86;--rule:#2e2e2b;--grid:#262624;
--s1:#3987e5;--s2:#d95926;--neutral:#6a6964;--o1:#86b6ef;--o2:#5598e7;--o3:#2a78d6;--o4:#184f95;--hl:#3a3218}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif}
.wrap{max-width:1280px;margin:0 auto;padding:24px 16px 64px}
h1{font-size:24px;margin:0 0 4px}h2{font-size:17px;margin:0 0 4px}.sub{color:var(--ink2);margin:0 0 20px}
.card{background:var(--surface);border:1px solid var(--rule);border-radius:10px;padding:18px 18px 14px;margin:0 0 16px}
.cap{color:var(--ink2);font-size:13px;margin:0 0 12px}
.story{font-size:15px}.story li{margin:4px 0}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin:0 0 16px}
.kpi{background:var(--surface);border:1px solid var(--rule);border-radius:10px;padding:14px}
.kpi .v{font-size:26px;font-weight:650;letter-spacing:-.01em}.kpi .l{color:var(--ink2);font-size:12.5px}.kpi .n{color:var(--muted);font-size:12px;margin-top:2px}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px}@media(max-width:900px){.grid2{grid-template-columns:1fr}}
svg{display:block;width:100%;height:auto;overflow:visible}svg text{fill:var(--ink2);font-size:11px}
.axis line,.axis path{stroke:var(--rule)}.gridl{stroke:var(--grid)}
.legend{display:flex;flex-wrap:wrap;gap:14px;font-size:12.5px;color:var(--ink2);margin:0 0 8px}.legend i{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:6px;vertical-align:-1px}
#tip{position:fixed;pointer-events:none;background:var(--surface);color:var(--ink);border:1px solid var(--rule);border-radius:8px;padding:8px 10px;font-size:12.5px;box-shadow:0 4px 18px rgba(0,0,0,.12);display:none;z-index:9;max-width:320px}
#tip b{display:block;margin-bottom:2px}#tip .r{display:flex;justify-content:space-between;gap:16px}
.ctrl{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:0 0 10px}
input,select,button{font:inherit;color:var(--ink);background:var(--surface);border:1px solid var(--rule);border-radius:6px;padding:6px 8px}
button{cursor:pointer}button.on{background:var(--ink);color:var(--surface)}
.tw{overflow:auto;max-height:560px;border:1px solid var(--rule);border-radius:8px}
table{border-collapse:collapse;width:100%;font-size:12.5px}th,td{padding:6px 8px;border-bottom:1px solid var(--rule);text-align:right;white-space:nowrap}
th{position:sticky;top:0;background:var(--surface);cursor:pointer;user-select:none;color:var(--ink2);font-weight:600}
th:first-child,td:first-child,th.l,td.l{text-align:left}tr:hover td{background:var(--hl)}
.bad{color:var(--crit);font-weight:600}.ok{color:var(--good)}.mut{color:var(--muted)}
.note{font-size:12.5px;color:var(--ink2)}.note li{margin:3px 0}
.tag{display:inline-block;font-size:11px;padding:1px 6px;border-radius:9px;border:1px solid var(--rule);color:var(--ink2)}
.tabs{display:flex;gap:6px;margin:0 0 10px}
</style></head><body><div class="wrap">
<h1>HIE Capacity Actuals</h1>
<p class="sub" id="sub"></p>

<div class="card story"><h2>The story</h2><ul id="story"></ul></div>

<div class="kpis" id="kpis"></div>

<div class="card"><h2>Revenue HIE delivered late, by the month it was promised</h2>
<p class="cap">Value at selling price of units HIE committed (PI date) for each month, split by how late they actually finished. "Still open" = past its PI date and not yet produced. This is revenue <b>pushed out</b>, not necessarily lost — lost sales depend on whether Jacksonville had cover.</p>
<div class="legend" id="lg-rev"></div><div id="ch-rev"></div></div>

<div class="grid2">
<div class="card"><h2>Overdue backlog at HIE</h2><p class="cap">Value of units past their original PI date and not yet produced, at each weekly report.</p><div id="ch-back"></div></div>
<div class="card"><h2>How late HIE finishes, per PO line</h2><p class="cap">Finish date vs original PI date, units. Finished lines only.</p><div id="ch-hist"></div></div>
</div>

<div class="grid2">
<div class="card"><h2>HIE output per month</h2><p class="cap">Units through production (good + rejects), from week-over-week movement on the open order list. Months with missing weekly reports are spread evenly.</p><div id="ch-cap"></div></div>
<div class="card"><h2>Reject rate per month</h2><p class="cap">B-grade + defect units as % of units through production.</p><div id="ch-rej"></div></div>
</div>


<div class="card" id="costcard"><h2>HIE price changes year over year</h2>
<p class="cap">HIE's own unit price per SKU per PO, from the weekly warehouse inventory list. Year = when the PO was placed. Change is measured SKU by SKU on items bought in both years, weighted by units, so mix shifts don't distort it. $ impact = price change × units ordered that year.</p>
<div class="kpis" id="costk" style="margin-bottom:8px"></div>
<div class="grid2"><div><div id="ch-yoy"></div></div><div class="tw" style="max-height:300px"><table id="t-yoy"></table></div></div></div>
<div class="card"><h2>Price change by product line</h2><p class="cap">Unit-weighted % change vs prior year; $ = extra (or saved) cost on that year's units.</p><div class="tw"><table id="t-cl"></table></div></div>
<div class="card"><h2>Price history by SKU</h2><p class="cap">Unit-weighted HIE price by PO year. Sorted by cumulative $ impact of increases.</p>
<div class="ctrl"><input id="qc" placeholder="Search SKU…" size="28"><select id="flc"></select><span class="note" id="cntc"></span></div><div class="tw"><table id="t-cs"></table></div></div>
<div class="card"><h2>Same SKU, different price in the same year</h2><p class="cap">SKUs where HIE charged more than 3% apart on different POs within one year — worth checking against the PI.</p><div class="tw" style="max-height:360px"><table id="t-spread"></table></div></div>

<div class="card"><h2>By product line</h2><p class="cap">Sorted by late revenue. Click a header to sort.</p><div class="tw"><table id="t-pl"></table></div></div>

<div class="card"><h2>By SKU</h2><p class="cap">Every HIE item on the open order list since <span id="sfirst"></span>. Slip = days finished after original PI date (negative = early). Late $ = value of units finished more than <span class="gr"></span> days late, plus value overdue today.</p>
<div class="ctrl"><input id="q" placeholder="Search SKU…" size="28"><select id="fl"></select><label class="note"><input type="checkbox" id="onlyod"> Only SKUs overdue now</label><span class="note" id="cnt"></span></div>
<div class="tw"><table id="t-sku"></table></div></div>

<div class="card"><h2>Open PO lines past their PI date today</h2><p class="cap">As of the latest HIE report. Latest PI = the PI date on HIE's most recent list (HIE sometimes moves it earlier as well as later); Re-sched = times it changed.</p>
<div class="ctrl"><input id="q2" placeholder="Search SKU or PO…" size="28"><span class="note" id="cnt2"></span></div>
<div class="tw"><table id="t-od"></table></div></div>

<div class="card note"><h2>How this was built</h2><ul id="method"></ul></div>
</div><div id="tip"></div>
<script>
const M = __DATA__;
const K = M.kpi;
const fmt = n => n==null||n===''?'–':Math.round(n).toLocaleString('en-US');
const usd = n => n==null?'–':(Math.abs(n)>=1e6?'$'+(n/1e6).toFixed(2)+'M':Math.abs(n)>=1e3?'$'+(n/1e3).toFixed(0)+'K':'$'+Math.round(n));
const usdF = n => n==null?'–':'$'+Math.round(n).toLocaleString('en-US');
const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const mlab = s => {const [y,m]=s.split('-');return ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+m-1]+' '+y.slice(2)};
const dlab = s => {const d=new Date(s+'T00:00');return d.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'2-digit'})};

// ---------- header, story, KPIs
document.getElementById('sub').textContent = `${K.n_reports} weekly HIE open order lists, ${dlab(K.first)} – ${dlab(K.last)} · ${fmt(K.n_pos)} POs · ${fmt(K.n_lines)} PO lines · ${fmt(K.n_skus)} SKUs · as of ${dlab(K.today)}`;
document.getElementById('sfirst').textContent = dlab(K.first);
document.querySelectorAll('.gr').forEach(e=>e.textContent=K.grace);
const worstM = [...M.rev_month].sort((a,b)=>b.late_rev-a.late_rev)[0];
const worstL = M.pl[0];
const story = [
 `<b>HIE hits its own PI date on just over half of what it builds.</b> ${K.ontime_units_pct}% of units finished within ${K.grace} days of the PI date HIE agreed to; the median PO line lands ${K.med_slip} days late and one in four lands ${K.p75_slip}+ days late.`,
 `<b>${usd(K.late_rev_total)} of product (${K.late_rev_pct}% of the ${usd(K.due_rev_total)} HIE committed) arrived late</b> — about ${usd(K.late_rev_month_avg)} a month pushed out of the month we were promised it. Worst month: ${mlab(worstM.m)}, ${usd(worstM.late_rev)} late.`,
 `<b>Right now ${fmt(K.od_units_now)} units (${usd(K.od_rev_now)}) are past their PI date</b> on ${K.od_lines_now} open PO lines.`,
 `<b>${worstL.line} carries the most late revenue</b> (${usd(worstL.late_rev)}). Lines with the longest average slip: ${M.pl.filter(p=>p.avg_slip!=null&&p.made>500).sort((a,b)=>b.avg_slip-a.avg_slip).slice(0,3).map(p=>`${p.line} ${Math.round(p.avg_slip)}d`).join(', ')}.`,
 `<b>Capacity:</b> HIE has averaged ${fmt(K.avg_week_all)} units/week across the period and ${fmt(K.avg_week_recent)}/week over the last three full months. Rejects run ${K.rej_rate}% of output (${fmt(K.rej_units)} units B-grade or defect).`,
];
document.getElementById('story').innerHTML = story.map(s=>`<li>${s}</li>`).join('');
const kp = [
 [K.ontime_units_pct+'%','Units on time','within '+K.grace+'d of PI date'],
 [K.med_slip+'d','Median slip per PO line','75th pct '+K.p75_slip+'d'],
 [usd(K.late_rev_total),'Revenue delivered late',K.late_rev_pct+'% of committed'],
 [usd(K.late_rev_month_avg),'Late revenue per month','average'],
 [usd(K.od_rev_now),'Overdue now',fmt(K.od_units_now)+' units · '+K.od_lines_now+' lines'],
 [fmt(K.avg_week_recent),'Units / week, last 3 mo','period avg '+fmt(K.avg_week_all)],
 [K.rej_rate+'%','Reject rate',fmt(K.rej_units)+' B-grade + defect'],
];
document.getElementById('kpis').innerHTML = kp.map(([v,l,n])=>`<div class="kpi"><div class="v">${v}</div><div class="l">${l}</div><div class="n">${n}</div></div>`).join('');

// ---------- tooltip
const tip = document.getElementById('tip');
function showTip(e, html){tip.innerHTML=html;tip.style.display='block';const x=e.clientX+14,y=e.clientY+14;const r=tip.getBoundingClientRect();
 tip.style.left=Math.min(x,innerWidth-r.width-8)+'px';tip.style.top=Math.min(y,innerHeight-r.height-8)+'px'}
function hideTip(){tip.style.display='none'}
const row=(l,v)=>`<div class="r"><span>${l}</span><span>${v}</span></div>`;

// ---------- svg helpers
const NS='http://www.w3.org/2000/svg';
function el(t,a,p){const e=document.createElementNS(NS,t);for(const k in a)e.setAttribute(k,a[k]);if(p)p.appendChild(e);return e}
function nice(max){if(max<=0)return 1;const p=Math.pow(10,Math.floor(Math.log10(max)));const f=max/p;return (f<=1?1:f<=2?2:f<=2.5?2.5:f<=5?5:10)*p}
function frame(id,W,H,pad,ymax,yfmt){const host=document.getElementById(id);host.innerHTML='';const s=el('svg',{viewBox:`0 0 ${W} ${H}`,role:'img'},host);
 const top=nice(ymax),y=v=>pad.t+(H-pad.t-pad.b)*(1-v/top);
 for(let i=0;i<=4;i++){const v=top*i/4,yy=y(v);el('line',{x1:pad.l,x2:W-pad.r,y1:yy,y2:yy,class:'gridl'},s);const t=el('text',{x:pad.l-6,y:yy+3.5,'text-anchor':'end'},s);t.textContent=yfmt(v)}
 return {s,y,top}}
function xlabels(s,xs,labs,H,pad,every){labs.forEach((l,i)=>{if(i%every)return;const t=el('text',{x:xs[i],y:H-pad.b+14,'text-anchor':'middle'},s);t.textContent=l})}
function bar(s,x,y0,y1,w,fill,round){const h=Math.max(0,y0-y1);if(h<=0)return null;
 const r=round?Math.min(4,w/2,h):0;const d=`M${x},${y0}V${y1+r}${r?`Q${x},${y1} ${x+r},${y1}`:''}H${x+w-r}${r?`Q${x+w},${y1} ${x+w},${y1+r}`:''}V${y0}Z`;return el('path',{d,fill},s)}

// ---------- 1. revenue by commit month (stacked)
(function(){
 const D=M.rev_month.filter(r=>r.due_rev>0);const segs=[['ontime','On time / early','--neutral'],['late30','8–30d late','--o1'],['late90','31–90d late','--o2'],['late_more','90d+ late','--o3'],['still_open','Still open, past PI','--o4']];
 D.forEach(r=>r.ontime=Math.max(0,r.due_rev-r.late_rev));
 document.getElementById('lg-rev').innerHTML=segs.map(([k,l,c])=>`<span><i style="background:var(${c})"></i>${l}</span>`).join('');
 const W=1200,H=320,pad={l:56,r:8,t:10,b:26};const {s,y}=frame('ch-rev',W,H,pad,Math.max(...D.map(r=>r.due_rev)),v=>usd(v));
 const n=D.length,bw=(W-pad.l-pad.r)/n,w=Math.max(3,bw-4);const xs=D.map((_,i)=>pad.l+i*bw+bw/2);
 D.forEach((r,i)=>{let acc=0;const x=pad.l+i*bw+2;segs.forEach(([k,,c],j)=>{const v=r[k];if(!v)return;const p=bar(s,x,y(acc)-(acc?1:0),y(acc+v),w,`var(${c})`,j===segs.length-1||segs.slice(j+1).every(([kk])=>!r[kk]));acc+=v});
  const hit=el('rect',{x:pad.l+i*bw,y:pad.t,width:bw,height:H-pad.t-pad.b,fill:'transparent'},s);
  hit.addEventListener('mousemove',e=>showTip(e,`<b>${mlab(r.m)} commitments</b>${row('Committed',usdF(r.due_rev))}${row('Late (>'+K.grace+'d)',usdF(r.late_rev)+' · '+r.pct_late+'%')}${row('Avg days late',r.avg_days_late+'d')}${row('Units late',fmt(r.late_units)+' / '+fmt(r.due_units))}${row('Still open',usdF(r.still_open))}`));hit.addEventListener('mouseleave',hideTip)});
 xlabels(s,xs,D.map(r=>mlab(r.m)),H,pad,n>30?3:n>16?2:1)})();

// ---------- 2. backlog line
(function(){
 const D=M.backlog;const W=600,H=260,pad={l:56,r:10,t:10,b:26};const {s,y}=frame('ch-back',W,H,pad,Math.max(...D.map(r=>r.rev)),v=>usd(v));
 const t0=new Date(D[0].d).getTime(),t1=new Date(D[D.length-1].d).getTime();const x=d=>pad.l+(W-pad.l-pad.r)*((new Date(d).getTime()-t0)/(t1-t0||1));
 let path='';let prev=null;D.forEach(r=>{const gap=prev&&(new Date(r.d)-new Date(prev.d))/864e5>21;path+=(path&&!gap?'L':'M')+x(r.d)+','+y(r.rev);prev=r});
 el('path',{d:path,fill:'none',stroke:'var(--s1)','stroke-width':2,'stroke-linejoin':'round'},s);
 const yrs=[...new Set(D.map(r=>r.d.slice(0,4)))];yrs.forEach(yr=>{const d=yr+'-01-01';if(new Date(d)<new Date(D[0].d))return;const t=el('text',{x:x(d),y:H-pad.b+14,'text-anchor':'middle'},s);t.textContent=yr;el('line',{x1:x(d),x2:x(d),y1:pad.t,y2:H-pad.b,class:'gridl'},s)});
 const cross=el('line',{y1:pad.t,y2:H-pad.b,stroke:'var(--muted)','stroke-dasharray':'3 3',visibility:'hidden'},s);const dot=el('circle',{r:4.5,fill:'var(--s1)',stroke:'var(--surface)','stroke-width':2,visibility:'hidden'},s);
 const hit=el('rect',{x:pad.l,y:pad.t,width:W-pad.l-pad.r,height:H-pad.t-pad.b,fill:'transparent'},s);
 hit.addEventListener('mousemove',e=>{const r=s.getBoundingClientRect();const px=(e.clientX-r.left)*W/r.width;let best=D[0];D.forEach(d=>{if(Math.abs(x(d.d)-px)<Math.abs(x(best.d)-px))best=d});
  cross.setAttribute('x1',x(best.d));cross.setAttribute('x2',x(best.d));cross.setAttribute('visibility','visible');dot.setAttribute('cx',x(best.d));dot.setAttribute('cy',y(best.rev));dot.setAttribute('visibility','visible');
  showTip(e,`<b>${dlab(best.d)}</b>${row('Overdue value',usdF(best.rev))}${row('Overdue units',fmt(best.units))}${row('PO lines',fmt(best.lines))}`)});
 hit.addEventListener('mouseleave',()=>{hideTip();cross.setAttribute('visibility','hidden');dot.setAttribute('visibility','hidden')});
 const g=K.gaps.filter(g=>g[2]>21);if(g.length){const t=el('text',{x:W-pad.r,y:H-2,'text-anchor':'end'},s);t.textContent='Line breaks = weeks with no HIE report'}})();

// ---------- 3. histogram
(function(){
 const D=M.hist;const W=600,H=260,pad={l:48,r:8,t:10,b:40};const {s,y}=frame('ch-hist',W,H,pad,Math.max(...D.map(r=>r.units)),v=>fmt(v));
 const bw=(W-pad.l-pad.r)/D.length;D.forEach((r,i)=>{const late=i>=3;bar(s,pad.l+i*bw+3,y(0),y(r.units),bw-6,late?'var(--s2)':'var(--neutral)',true);
  const t=el('text',{x:pad.l+i*bw+bw/2,y:H-pad.b+14,'text-anchor':'middle'},s);t.textContent=r.b.replace(' late','');
  const hit=el('rect',{x:pad.l+i*bw,y:pad.t,width:bw,height:H-pad.t-pad.b,fill:'transparent'},s);hit.addEventListener('mousemove',e=>showTip(e,`<b>${r.b}</b>${row('Units',fmt(r.units))}${row('PO lines',fmt(r.lines))}`));hit.addEventListener('mouseleave',hideTip)});
 const t=el('text',{x:pad.l,y:H-6},s);t.textContent='Orange = finished more than '+K.grace+' days after PI date';})();

// ---------- 4/5. capacity + reject rate
function monthBars(id,D,key,fmtv,tipf,color){const W=600,H=260,pad={l:52,r:8,t:10,b:26};const {s,y}=frame(id,W,H,pad,Math.max(...D.map(r=>r[key]))||1,fmtv);
 const n=D.length,bw=(W-pad.l-pad.r)/n;const xs=D.map((_,i)=>pad.l+i*bw+bw/2);
 D.forEach((r,i)=>{bar(s,pad.l+i*bw+1,y(0),y(r[key]),Math.max(2,bw-2),color,true);const hit=el('rect',{x:pad.l+i*bw,y:pad.t,width:bw,height:H-pad.t-pad.b,fill:'transparent'},s);
  hit.addEventListener('mousemove',e=>showTip(e,tipf(r)));hit.addEventListener('mouseleave',hideTip)});xlabels(s,xs,D.map(r=>mlab(r.m)),H,pad,n>30?6:n>16?3:1)}
monthBars('ch-cap',M.cap_month,'done',v=>fmt(v),r=>`<b>${mlab(r.m)}</b>${row('Units through production',fmt(r.done))}${row('Good',fmt(r.good))}${row('Rejects',fmt(r.rej))}${row('Per week',fmt(r.perweek))}${r.coverage<20?row('Note','partial month in data'):''}`,'var(--s1)');
monthBars('ch-rej',M.rej_month,'rate',v=>v.toFixed(1)+'%',r=>`<b>${mlab(r.m)}</b>${row('Reject rate',r.rate+'%')}${row('Rejects',fmt(r.rej))}${row('Through production',fmt(r.done))}`,'var(--s2)');

// ---------- tables
function table(id,cols,rows,state){const t=document.getElementById(id);
 const head='<tr>'+cols.map((c,i)=>`<th data-i="${i}" class="${c.l?'l':''}">${c.h}${state.s===i?(state.d>0?' ▲':' ▼'):''}</th>`).join('')+'</tr>';
 const sorted=[...rows];if(state.s!=null){const c=cols[state.s];sorted.sort((a,b)=>{const va=c.v(a),vb=c.v(b);if(va==null&&vb==null)return 0;if(va==null)return 1;if(vb==null)return -1;return (va>vb?1:va<vb?-1:0)*state.d})}
 t.innerHTML=head+sorted.slice(0,state.limit||2000).map(r=>'<tr>'+cols.map(c=>`<td class="${(c.l?'l ':'')+(c.cls?c.cls(r):'')}">${c.f?c.f(r):(c.v(r)??'–')}</td>`).join('')+'</tr>').join('');
 t.querySelectorAll('th').forEach(th=>th.onclick=()=>{const i=+th.dataset.i;state.d=state.s===i?-state.d:-1;state.s=i;state.render()})}
const slipCls=v=>v==null?'mut':v>30?'bad':v<=K.grace?'ok':'';
const plCols=[{h:'Product line',v:r=>r.line,l:1},{h:'SKUs',v:r=>r.skus,f:r=>fmt(r.skus)},{h:'Units ordered',v:r=>r.ordered,f:r=>fmt(r.ordered)},{h:'Units made',v:r=>r.made,f:r=>fmt(r.made)},
 {h:'Avg slip (d)',v:r=>r.avg_slip,f:r=>r.avg_slip==null?'–':Math.round(r.avg_slip),cls:r=>slipCls(r.avg_slip)},{h:'On time %',v:r=>r.ontime,f:r=>r.ontime==null?'–':r.ontime+'%'},
 {h:'Reject %',v:r=>r.rej_rate,f:r=>r.rej_rate==null?'–':r.rej_rate.toFixed(2)+'%'},{h:'Overdue units',v:r=>r.od_units,f:r=>fmt(r.od_units)},{h:'Overdue $',v:r=>r.od_rev,f:r=>usdF(r.od_rev)},{h:'Late $',v:r=>r.late_rev,f:r=>usdF(r.late_rev)}];
const plS={s:9,d:-1};plS.render=()=>table('t-pl',plCols,M.pl,plS);plS.render();

const skuCols=[{h:'SKU',v:r=>r.sku,l:1},{h:'Line',v:r=>r.line,l:1},{h:'POs',v:r=>r.pos},{h:'Ordered',v:r=>r.ordered,f:r=>fmt(r.ordered)},{h:'Made',v:r=>r.made,f:r=>fmt(r.made)},
 {h:'Avg slip',v:r=>r.avg_slip,f:r=>r.avg_slip==null?'–':Math.round(r.avg_slip),cls:r=>slipCls(r.avg_slip)},{h:'Median',v:r=>r.med_slip,f:r=>r.med_slip==null?'–':Math.round(r.med_slip)},{h:'Worst',v:r=>r.max_slip,f:r=>r.max_slip==null?'–':r.max_slip},
 {h:'On time %',v:r=>r.ontime,f:r=>r.ontime==null?'–':r.ontime+'%'},{h:'Re-sched',v:r=>r.resched},{h:'B-grade',v:r=>r.bgrade,f:r=>fmt(r.bgrade)},{h:'Defect',v:r=>r.defect,f:r=>fmt(r.defect)},
 {h:'Reject %',v:r=>r.rej_rate,f:r=>r.rej_rate==null?'–':r.rej_rate.toFixed(1)+'%',cls:r=>r.rej_rate>3?'bad':''},{h:'Overdue now',v:r=>r.od_units,f:r=>fmt(r.od_units)},{h:'Oldest overdue (d)',v:r=>r.max_od,f:r=>r.max_od??'–',cls:r=>r.max_od>60?'bad':''},
 {h:'Sell price',v:r=>r.price,f:r=>`<span title="${r.pbasis}">${usdF(r.price)}${r.pbasis==='Open SO rate'?'':r.pbasis.startsWith('Cost')?'*':'†'}</span>`},{h:'Late $',v:r=>r.late_rev,f:r=>usdF(r.late_rev)}];
const fl=document.getElementById('fl');fl.innerHTML='<option value="">All product lines</option>'+[...new Set(M.sku.map(r=>r.line))].sort().map(l=>`<option>${l}</option>`).join('');
const skuS={s:16,d:-1};skuS.render=()=>{const q=document.getElementById('q').value.trim().toUpperCase(),l=fl.value,od=document.getElementById('onlyod').checked;
 const rows=M.sku.filter(r=>(!q||r.sku.includes(q))&&(!l||r.line===l)&&(!od||r.od_units>0));document.getElementById('cnt').textContent=rows.length+' SKUs';table('t-sku',skuCols,rows,skuS)};
['q','fl','onlyod'].forEach(i=>document.getElementById(i).addEventListener('input',skuS.render));skuS.render();

const odCols=[{h:'PO',v:r=>r.po,l:1},{h:'SKU',v:r=>r.sku,l:1},{h:'Line',v:r=>r.line,l:1},{h:'Units open',v:r=>r.qty_open,f:r=>fmt(r.qty_open)},{h:'Original PI',v:r=>r.orig,f:r=>dlab(r.orig)},{h:'Latest PI',v:r=>r.current,f:r=>r.current?dlab(r.current):'–'},
 {h:'Days past PI',v:r=>r.days,cls:r=>r.days>60?'bad':''},{h:'Re-sched',v:r=>r.resched},{h:'Value',v:r=>r.rev,f:r=>usdF(r.rev)}];
const odS={s:8,d:-1};odS.render=()=>{const q=document.getElementById('q2').value.trim().toUpperCase();const rows=M.od_lines.filter(r=>!q||r.sku.includes(q)||r.po.includes(q));document.getElementById('cnt2').textContent=rows.length+' lines';table('t-od',odCols,rows,odS)};
document.getElementById('q2').addEventListener('input',odS.render);odS.render();


// ---------- cost section
if(M.cost){const C=M.cost;const Y=C.yoy;const pct=v=>v==null?'–':(v>0?'+':'')+v.toFixed(1)+'%';
 const last=Y[Y.length-1]||{};const tot=Y.reduce((a,r)=>a+r.impact,0);const idx=C.index[C.index.length-1];
 document.getElementById('costk').innerHTML=[
  [pct(last.wpct),`Price change ${last.from}→${last.to}`,`${last.skus} SKUs bought both years`],
  [idx?(idx.index-100>0?'+':'')+(idx.index-100).toFixed(1)+'%':'–',`Cumulative since ${C.index[0].year}`,'chain-linked, unit-weighted'],
  [usd(tot),'Extra cost from price changes','sum of yearly $ impact'],
  [fmt(C.meta.n_skus),'SKUs priced',`${fmt(C.meta.n_sku_po)} SKU-PO prices · ${C.meta.n_files} weekly lists`]
 ].map(([v,l,n])=>`<div class="kpi"><div class="v">${v}</div><div class="l">${l}</div><div class="n">${n}</div></div>`).join('');
 // index chart
 (function(){const D=C.index;const W=600,H=230,pad={l:44,r:10,t:10,b:26};const lo=Math.min(...D.map(r=>r.index))-2,hi=Math.max(...D.map(r=>r.index))+2;
  const host=document.getElementById('ch-yoy');host.innerHTML='';const s=el('svg',{viewBox:`0 0 ${W} ${H}`},host);const y=v=>pad.t+(H-pad.t-pad.b)*(1-(v-lo)/(hi-lo));
  for(let i=0;i<=4;i++){const v=lo+(hi-lo)*i/4;el('line',{x1:pad.l,x2:W-pad.r,y1:y(v),y2:y(v),class:'gridl'},s);const t=el('text',{x:pad.l-6,y:y(v)+3.5,'text-anchor':'end'},s);t.textContent=v.toFixed(0)}
  const x=i=>pad.l+(W-pad.l-pad.r)*(D.length>1?i/(D.length-1):.5);el('path',{d:D.map((r,i)=>(i?'L':'M')+x(i)+','+y(r.index)).join(''),fill:'none',stroke:'var(--s1)','stroke-width':2},s);
  D.forEach((r,i)=>{const c=el('circle',{cx:x(i),cy:y(r.index),r:5,fill:'var(--s1)',stroke:'var(--surface)','stroke-width':2},s);const t=el('text',{x:x(i),y:H-pad.b+14,'text-anchor':'middle'},s);t.textContent=r.year;
   const lb=el('text',{x:x(i),y:y(r.index)-10,'text-anchor':'middle'},s);lb.textContent=r.index.toFixed(1);
   const hit=el('circle',{cx:x(i),cy:y(r.index),r:16,fill:'transparent'},s);const yy=Y.find(q=>q.to===r.year);hit.addEventListener('mousemove',e=>showTip(e,`<b>${r.year}</b>${row('HIE price index',r.index.toFixed(1))}${yy?row('vs prior year',pct(yy.wpct))+row('$ impact',usdF(yy.impact)):''}`));hit.addEventListener('mouseleave',hideTip)});
  const t=el('text',{x:pad.l,y:H-2},s);t.textContent='HIE price index, '+D[0].year+' = 100'})();
 const ycols=[{h:'Years',v:r=>r.from+'→'+r.to,l:1},{h:'SKUs',v:r=>r.skus},{h:'Units',v:r=>r.units,f:r=>fmt(r.units)},{h:'Weighted %',v:r=>r.wpct,f:r=>pct(r.wpct),cls:r=>r.wpct>0.5?'bad':''},{h:'Median %',v:r=>r.median_pct,f:r=>pct(r.median_pct)},{h:'Up / flat / down',v:r=>r.up,f:r=>`${r.up} / ${r.flat} / ${r.down}`},{h:'$ impact',v:r=>r.impact,f:r=>usdF(r.impact)}];
 const yS={s:null,d:-1};yS.render=()=>table('t-yoy',ycols,Y,yS);yS.render();
 const yrs=Y.map(r=>String(r.to));
 const lcols=[{h:'Product line',v:r=>r.line,l:1},...yrs.flatMap(y=>[{h:y+' %',v:r=>r[y]?.pct,f:r=>r[y]?pct(r[y].pct):'–',cls:r=>r[y]?.pct>0.5?'bad':''},{h:y+' $',v:r=>r[y]?.impact,f:r=>r[y]?usdF(r[y].impact):'–'}]),{h:'Total $',v:r=>r.impact,f:r=>usdF(r.impact)}];
 const lS={s:lcols.length-1,d:-1};lS.render=()=>table('t-cl',lcols,C.lines,lS);lS.render();
 const allY=[...new Set(C.skus.flatMap(r=>Object.keys(r.years)))].sort();
 const scols=[{h:'SKU',v:r=>r.sku,l:1},{h:'Line',v:r=>r.line,l:1},{h:'POs',v:r=>r.n_po},...allY.map(y=>({h:y,v:r=>r.years[y],f:r=>r.years[y]!=null?'$'+r.years[y].toFixed(2):'–'})),
  {h:'Total %',v:r=>r.total_pct,f:r=>r.total_pct==null?'–':pct(r.total_pct),cls:r=>r.total_pct>3?'bad':''},{h:'CAGR',v:r=>r.cagr,f:r=>r.cagr==null?'–':pct(r.cagr)},{h:'$ impact',v:r=>r.impact,f:r=>usdF(r.impact)}];
 const flc=document.getElementById('flc');flc.innerHTML='<option value="">All product lines</option>'+[...new Set(C.skus.map(r=>r.line))].sort().map(l=>`<option>${l}</option>`).join('');
 const sS={s:scols.length-1,d:-1};sS.render=()=>{const q=document.getElementById('qc').value.trim().toUpperCase(),l=flc.value;const rows=C.skus.filter(r=>(!q||r.sku.includes(q))&&(!l||r.line===l));document.getElementById('cntc').textContent=rows.length+' SKUs';table('t-cs',scols,rows,sS)};
 ['qc','flc'].forEach(i=>document.getElementById(i).addEventListener('input',sS.render));sS.render();
 const pcols=[{h:'SKU',v:r=>r.sku,l:1},{h:'Year',v:r=>r.year},{h:'Low',v:r=>r.low,f:r=>'$'+r.low.toFixed(2)},{h:'High',v:r=>r.high,f:r=>'$'+r.high.toFixed(2)},{h:'Spread',v:r=>r.pct,f:r=>'+'+r.pct+'%',cls:r=>r.pct>10?'bad':''}];
 const pS={s:4,d:-1};pS.render=()=>table('t-spread',pcols,C.spread,pS);pS.render();
} else {document.getElementById('costcard').innerHTML='<h2>HIE price changes</h2><p class="cap">No price data loaded.</p>'}

// ---------- method
const gaps=K.gaps.map(g=>`${dlab(g[0])} → ${dlab(g[1])} (${g[2]}d)`).join('; ');
document.getElementById('method').innerHTML=[
 `<b>Source:</b> every "Predator open order list" attached to HIE's Weekly reports email (Lydia, sal1_07@hamsonint.com), ${dlab(K.first)} – ${dlab(K.last)}, ${K.n_reports} reports. Each row is one PO delivery line: order qty, PI date, scheduled finish, shipped / packing / warehouse, B-grade, defect, online (in production).`,
 `<b>Commitment = original PI date</b> — the date on HIE's proforma invoice the first time the line appeared. HIE's "Scheduled Finished Date" is its internal plan and often sits months after the PI date, and "Actually Finished Date" holds a forecast on open lines, so neither is used as the commitment. Finish = the report where the full line quantity was through production (or HIE's actual-finish date once the line is complete).`,
 `<b>Slip</b> = finish − original PI date. On time = within ${K.grace} days. Lines already finished when first seen, and lines that vanished with nothing produced (cancelled or re-split by HIE), are excluded from slip. ${fmt(K.n_fin)} finished lines are measured.`,
 `<b>Output</b> = week-over-week increase in good + rejects per PO line. Intervals with missing reports are spread evenly across their days.`,
 `<b>Revenue</b> = units × selling price. Price = average open sales-order rate for the SKU (${K.priced_skus} SKUs); otherwise NetSuite average cost × that product line's median price-to-cost ratio (${K.cost_skus} SKUs, marked * in the SKU table); HIE codes not in NetSuite take the median price of their closest SKU family (${K.family_skus} SKUs, e.g. BCP PRE BK RUSH …); anything still unmatched uses the overall median ${usdF(K.fallback_price)} (${K.unpriced_skus} SKUs). "Late revenue" is value delivered late, not proven lost sales — that needs Jacksonville stock-out history.`,
 `<b>Price changes:</b> HIE unit price per SKU and PO from the weekly 'Predator warehouse inventory list' (regular inventory sheet). PO year = first appearance on the open order list; POs already open in Jun 2023 are dated from their PO number. PO 7666 (HIE overstock bucket) excluded. YoY compares the same SKU across consecutive years, unit-weighted; index is chain-linked.`,
 `<b>Report gaps</b> (no weekly file found): ${gaps||'none'}.`,
].map(s=>`<li>${s}</li>`).join('');
</script></body></html>"""

html = HTML.replace('__DATA__', json.dumps(m, separators=(',', ':'), default=str))
os.makedirs(os.path.dirname(OUT), exist_ok=True)
open(OUT, 'w', encoding='utf-8').write(html)
print('wrote', OUT, round(len(html) / 1024), 'KB')
