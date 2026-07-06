/* ============================================================
   Revenue MIS Dashboard — app.js v3
   Tabs: Dashboard · Revenue Summary V2 · INS · FEES · PMS · Talk
   ============================================================ */
(function () {
"use strict";

var FY_SERIALS=(function(){var d=new Date(),y=d.getMonth()>=3?d.getFullYear():d.getFullYear()-1,s=[];for(var i=0;i<12;i++){var m=(3+i)%12;s.push(Math.round((new Date(Date.UTC(y+(m<3?1:0),m,1))-new Date(Date.UTC(1899,11,30)))/86400000));}return s;})();
var ACCENTS = ['#4f46e5','#0891b2','#059669','#ec4899','#8b5cf6','#d97706','#ef4444','#14b8a6','#f97316','#a855f7','#0ea5e9','#84cc16'];
var MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

/* ─── state ─── */
var DATA = null;
var RM2TEAM = {};
var charts = {};
var state = {
  tab:'dashboard', month:'All', chartTeam:'All',
  talkMonth:'YTD', revsumView:'mtd', revsumTeam:'All',
  insView:'rm', feesView:'rm', pmsView:'rm', talkView:'team',
  talkMetric:'netSale',
  insSearch:'', feesSearch:'', pmsSearch:'',
  insRawFilters:{rm:'All',insType:'All',partner:'All',status:'All'},
  insRawSort:{col:'',dir:1}
};

/* ── GitHub config ─────────────────────────────────────────────
   The dashboard auto-fetches the .xlsb from this repo on every load.
   For a private repo set token to a Personal Access Token with `repo` scope.
   ──────────────────────────────────────────────────────────── */
var GH = {
  owner: 'FincartOptima',
  repo:  'RevenueMISDashboard',
  branch:'main',
  token: ''   /* optional PAT for private repo */
};

/* ─── DOM/number helpers ─── */
function $(id){ return document.getElementById(id); }
function el(tag,cls,html){ var e=document.createElement(tag); if(cls)e.className=cls; if(html!=null)e.innerHTML=html; return e; }
function num(v){ if(v==null)return 0; if(typeof v==='number')return v; var s=(''+v).replace(/,/g,'').trim(); var f=parseFloat(s); return isNaN(f)?0:f; }

function inGroup(n){
  var neg=n<0; n=Math.abs(Math.round(n));
  var s=''+n; if(s.length<=3)return (neg?'-':'')+s;
  var last3=s.slice(-3),rest=s.slice(0,-3).replace(/\B(?=(\d{2})+(?!\d))/g,',');
  return (neg?'-':'')+rest+','+last3;
}
function inr(n){ return '₹'+inGroup(n); }
function inrCr(n){
  var neg=n<0?'-':''; var a=Math.abs(n);
  if(a>=1e7)return neg+'₹'+(a/1e7).toFixed(2)+' Cr';
  if(a>=1e5)return neg+'₹'+(a/1e5).toFixed(2)+' L';
  if(a>=1e3)return neg+'₹'+(a/1e3).toFixed(1)+'K';
  return neg+'₹'+Math.round(a);
}
function numFmt(n){ return inGroup(n); }
function pct(n){ return (n*100).toFixed(1)+'%'; }
function hours(h){ var hh=Math.floor(h); var mm=Math.round((h-hh)*60); return hh+'h '+(mm<10?'0':'')+mm+'m'; }
function avg(tot,n){ return n?Math.round(tot/n):0; }

/* MTD label — appends " — MTD" if month matches current calendar month */
function mtdLabel(m){
  var now=new Date();
  var cur=MONTH_ABBR[now.getMonth()]+'-'+now.getFullYear();
  return m===cur?m+' — MTD':m;
}

/* Achievement colour helpers */
function achBarColor(ratio){
  return ratio>=0.8?'rgba(5,150,105,.85)':ratio>=0.5?'rgba(217,119,6,.85)':'rgba(225,29,72,.85)';
}
function achRowCls(ratio){
  return ratio>=0.8?'heat-high':ratio>=0.5?'heat-low':'heat-vlow';
}

function uniqInFyOrder(months){
  var present={}; months.forEach(function(m){present[m]=1;});
  return DATA.meta.fyMonths.filter(function(m){return present[m];});
}
function fyOrder(arr){
  var ord={}; DATA.meta.fyMonths.forEach(function(m,i){ord[m]=i;});
  return arr.slice().sort(function(a,b){return (ord[a.key]==null?99:ord[a.key])-(ord[b.key]==null?99:ord[b.key]);});
}
function pivot(rows,keyFn,valFns){
  var map={};
  rows.forEach(function(r){
    var k=keyFn(r); if(!k||k==='')k='(blank)';
    if(!map[k]){map[k]={key:k,_count:0}; Object.keys(valFns).forEach(function(f){map[k][f]=0;});}
    map[k]._count++;
    Object.keys(valFns).forEach(function(f){map[k][f]+=valFns[f](r);});
  });
  return Object.keys(map).map(function(k){return map[k];});
}
function byRev(a,b){return b.rev-a.rev;}
function sum(arr,fn){return arr.reduce(function(a,r){return a+(fn?fn(r):r);},0);}

/* ─── RM→Team ─── */
function buildRmTeam(){
  RM2TEAM={};
  DATA.talk.rows.forEach(function(r){ if(r.name)RM2TEAM[r.name.toLowerCase().trim()]=r.team; });
}
function teamOf(rm){ if(!rm)return 'Unassigned'; return RM2TEAM[rm.toLowerCase().trim()]||'Unassigned'; }

/* team helpers */
function filterByTeam(rows,team){
  if(!team||team==='All')return rows;
  return rows.filter(function(r){ return teamOf(r.rm)===team; });
}
function allTeams(){
  var t={}; DATA.talk.rows.forEach(function(r){if(r.team&&r.team.toUpperCase()!=='B2B')t[r.team]=1;});
  return Object.keys(t).sort();
}

/* ─── filters ─── */
function getINS(){
  return filterByTeam(DATA.ins.filter(function(r){ return state.month==='All'||r.month===state.month; }),state.chartTeam);
}
function getFEES(){
  return filterByTeam(DATA.fees.filter(function(r){ return state.month==='All'||r.month===state.month; }),state.chartTeam);
}
function getPMS(){
  return filterByTeam((DATA.pms||[]).filter(function(r){ return state.month==='All'||r.month===state.month; }),state.chartTeam);
}
function getTalkRows(){
  var rows=DATA.talk.rows;
  if(state.chartTeam!=='All') rows=rows.filter(function(r){return r.team===state.chartTeam;});
  return rows;
}

/* ─── TABLE BUILDER ─── */
function buildTable(cols,rows,opts){
  opts=opts||{};
  var wrap=el('div'); var t=el('table','data compact');
  var thead=el('thead'); var htr=el('tr');
  cols.forEach(function(c){
    var th=el('th',null,c.label);
    if(opts.onSort&&c.sortKey){
      th.className='sortable';
      if(opts.sortState&&opts.sortState.col===c.sortKey)
        th.innerHTML+=('<span class="sort-icon">'+(opts.sortState.dir>0?'▲':'▼')+'</span>');
      th.onclick=(function(k){return function(){opts.onSort(k);};})(c.sortKey);
    }
    htr.appendChild(th);
  });
  thead.appendChild(htr); t.appendChild(thead);
  var tb=el('tbody');
  var heatVals=null,hMax=0;
  if(opts.heat!=null&&!opts.rowCls){
    heatVals=rows.map(function(r,ri){return num(cols[opts.heat].get(r,ri));});
    hMax=Math.max.apply(null,heatVals.concat([0]));
  }
  rows.forEach(function(r,ri){
    var tr=el('tr');
    if(opts.rowCls){tr.className=opts.rowCls(r,ri);}
    else if(heatVals&&hMax>0){
      var ratio=heatVals[ri]/hMax;
      tr.className=ratio>=0.66?'heat-high':ratio>=0.33?'heat-mid':ratio>0?'heat-low':'heat-vlow';
    }
    cols.forEach(function(c){
      var td=el('td'); var v=c.get(r,ri);
      td.innerHTML=(c.fmt?c.fmt(v,r):v);
      if(c.align)td.style.textAlign=c.align;
      if(c.isRank)td.className='rank';
      tr.appendChild(td);
    });
    tb.appendChild(tr);
  });
  if(opts.grand){
    var gtr=el('tr','grand');
    cols.forEach(function(c,ci){
      var td=el('td');
      if(ci===0){td.innerHTML='TOTAL';}
      else if(c.sum){
        var tot=rows.reduce(function(a,r){return a+num(c.get(r,0));},0);
        td.innerHTML=c.fmt?c.fmt(tot,null):inGroup(tot);
      }
      if(c.align)td.style.textAlign=c.align;
      gtr.appendChild(td);
    });
    tb.appendChild(gtr);
  }
  t.appendChild(tb); wrap.appendChild(t); return wrap;
}

/* standard Name/Count/Revenue/Share table */
function revShareTable(rows,totalRev){
  totalRev=totalRev||rows.reduce(function(a,r){return a+r.rev;},0)||1;
  return buildTable([
    {label:'Name',get:function(r){return r.key;}},
    {label:'Count',get:function(r){return r._count;},fmt:numFmt,align:'right',sum:true},
    {label:'Revenue',get:function(r){return r.rev;},fmt:inr,align:'right',sum:true},
    {label:'Share %',get:function(r){return r.rev/totalRev;},fmt:pct,align:'right'}
  ],rows,{grand:true,heat:2});
}

/* ─── KPI CARD ─── */
function kpi(value,label,color){
  var d=el('div','kpi kpi-'+(color||'primary'));
  d.appendChild(el('div','kpi-value',value));
  d.appendChild(el('div','kpi-label',label));
  return d;
}
function kpiRow(items){
  var d=el('div','kpis');
  items.forEach(function(i){d.appendChild(kpi(i[0],i[1],i[2]));});
  return d;
}

/* ─── VIEW BAR ─── */
function makeViewBar(views,current,onSelect){
  var bar=el('div','view-bar');
  views.forEach(function(v){
    var btn=el('button','vbtn'+(v.id===current?' active':''),v.label);
    btn.onclick=function(){onSelect(v.id);};
    bar.appendChild(btn);
  });
  return bar;
}

/* ─── SECTION HELPER ─── */
function sect(title,barCls,infoKey){
  var s=el('div','section '+(barCls||''));
  var h=el('h2',null,title);
  if(infoKey){
    var btn=document.createElement('button');
    btn.className='info-btn'; btn.dataset.info=infoKey; btn.title='Chart info'; btn.textContent='i';
    h.appendChild(document.createTextNode(' '));
    h.appendChild(btn);
  }
  s.appendChild(h);
  return s;
}

/* ─── CANVAS WRAPPER ─── */
function cw(id,heightPx){
  var c=document.createElement('canvas'); c.id=id;
  var w=el('div','chart-wrap'); if(heightPx)w.style.height=heightPx+'px';
  w.appendChild(c); return w;
}

/* ─── CHARTS ─── */
function _chart(canvasId,config){
  var ctx=$(canvasId); if(!ctx)return;
  if(typeof Chart==='undefined'){
    if(ctx.parentNode)ctx.parentNode.innerHTML='<div style="color:var(--muted);font-size:12px;padding:20px;text-align:center">Chart requires internet (Chart.js CDN).</div>';
    return;
  }
  if(charts[canvasId])charts[canvasId].destroy();
  charts[canvasId]=new Chart(ctx,config);
}

/* vertical / stacked bar — datalabels enabled */
function drawBar(id,labels,datasets,opts){
  opts=opts||{};
  var tooMany=labels.length>18;
  _chart(id,{
    type:opts.type||'bar',
    data:{labels:labels,datasets:datasets},
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:opts.legend!==false,position:'top',labels:{font:{size:11}}},
        datalabels:{
          display:!tooMany&&opts.datalabels!==false,
          anchor:'end',align:'end',offset:2,
          font:{size:9,weight:'bold'},
          color:'#374151',
          formatter:function(v){return opts.money?inrCr(v):(v>0?numFmt(v):'');}
        },
        tooltip:{callbacks:{label:function(c){
          var v=c.parsed.y!=null?c.parsed.y:c.parsed;
          var lines=[c.dataset.label+': '+(opts.money?inrCr(v):numFmt(v))];
          if(opts.countData)lines.push('Count: '+numFmt(opts.countData[c.dataIndex]));
          return lines;
        }}}
      },
      scales:{
        x:{stacked:!!opts.stacked,ticks:{font:{size:10},maxRotation:40}},
        y:{stacked:!!opts.stacked,ticks:{font:{size:10},callback:function(v){return opts.money?inrCr(v):numFmt(v);}}}
      }
    }
  });
}

/* horizontal bar — for RM rankings, datalabels enabled */
function drawHBar(id,labels,datasets,opts){
  opts=opts||{};
  _chart(id,{
    type:'bar',
    data:{labels:labels,datasets:datasets},
    options:{
      indexAxis:'y',
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{display:opts.legend!==false,position:'top',labels:{font:{size:11}}},
        datalabels:{
          display:opts.datalabels!==false,
          anchor:'end',align:'end',offset:2,
          font:{size:9,weight:'bold'},
          color:'#374151',
          formatter:function(v){return opts.money?inrCr(v):(v>0?numFmt(v):'');}
        },
        tooltip:{callbacks:{label:function(c){
          var v=c.parsed.x!=null?c.parsed.x:c.parsed;
          return c.dataset.label+': '+(opts.money?inrCr(v):numFmt(v));
        }}}
      },
      scales:{
        y:{ticks:{font:{size:10}}},
        x:{ticks:{font:{size:10},callback:function(v){return opts.money?inrCr(v):numFmt(v);}}}
      }
    }
  });
}

/* doughnut — datalabels show % */
function drawDoughnut(id,labels,data,opts){
  opts=opts||{};
  var total=data.reduce(function(a,b){return a+b;},0)||1;
  _chart(id,{
    type:'doughnut',
    data:{labels:labels,datasets:[{
      data:data,
      backgroundColor:labels.map(function(_,i){return ACCENTS[i%ACCENTS.length];}),
      borderWidth:2,borderColor:'#fff'
    }]},
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{
        legend:{position:'right',labels:{font:{size:10},boxWidth:14}},
        datalabels:{
          display:true,
          formatter:function(v){return v>0?pct(v/total):'';},
          font:{size:9,weight:'bold'},
          color:'#fff'
        },
        tooltip:{callbacks:{label:function(c){
          return c.label+': '+(opts.money?inrCr(c.parsed):numFmt(c.parsed))+' ('+pct(c.parsed/total)+')';
        }}}
      }
    }
  });
}

/* achievement legend HTML element */
function achLegendEl(){
  var d=el('div','ach-legend');
  d.innerHTML='<span class="al-dot" style="background:#059669"></span>≥80% &nbsp;'+
    '<span class="al-dot" style="background:#d97706"></span>50–79% &nbsp;'+
    '<span class="al-dot" style="background:#e11d48"></span>&lt;50%';
  return d;
}

/* ============================================================
   TABS
   ============================================================ */
var TABS=[
  {id:'dashboard',label:'📊 Dashboard',cls:'t-dash'},
  {id:'revsum',label:'📋 Revenue Summary V2'},
  {id:'ins',label:'🛡️ INS — Insurance'},
  {id:'fees',label:'💰 FEES'},
  {id:'pms',label:'📈 PMS'},
  {id:'talk',label:'📞 Talk — RM Performance'}
];
function renderTabs(){
  var c=$('tabs'); c.innerHTML='';
  TABS.forEach(function(t){
    var b=el('button','tab '+(t.cls||'')+(state.tab===t.id?' active':''),t.label);
    b.onclick=function(){setTab(t.id);};
    c.appendChild(b);
  });
}
function setTab(id){
  state.tab=id;
  document.querySelectorAll('.tab-panel').forEach(function(p){p.classList.remove('active');});
  $('tab-'+id).classList.add('active');
  renderTabs();
  /* show month+team filter for all tabs except revsum */
  $('header-dash-filters').style.display=id==='revsum'?'none':'';
  $('header-talk-filters').style.display=id==='talk'?'':'none';
  $('header-revsum-filters').style.display=id==='revsum'?'':'none';
  renderActive();
}
function renderActive(){
  var t=state.tab;
  if(t==='dashboard')renderDashboard();
  else if(t==='revsum')renderRevsum();
  else if(t==='ins')renderINS();
  else if(t==='fees')renderFEES();
  else if(t==='pms')renderPMS();
  else if(t==='talk')renderTalk();
}

/* ============================================================
   DASHBOARD
   ============================================================ */
function renderDashboard(){
  var ins=getINS(), fees=getFEES(), pms=getPMS();
  var insRev=sum(ins,function(r){return r.revenue;});
  var feeRev=sum(fees,function(r){return r.revenue;});
  var pmsRev=sum(pms,function(r){return r.revenue;});
  var insPrem=sum(ins,function(r){return r.premium;});
  var talkRows=getTalkRows();
  var talkMonths=(DATA.talk||{}).months||[];

  /* ── Reconciled revenue from REVENUE REPORT (product-wise incl. Trail) ── */
  var rr=DATA.revreport;
  var k=$('dash-kpis'); k.innerHTML='';
  var tgl=$('dash-rev-toggle'); if(tgl)tgl.innerHTML='';
  if(rr&&(rr.monthly||rr.ytd)){
    var view=state.dashRevView||'mtd';
    var rrows=(view==='ytd'?rr.ytd:rr.monthly)||[];
    /* B2C, non-pool only — matches Revenue Summary V2.0 Column I grand total */
    rrows=rrows.filter(function(r){return (r.team||'').toUpperCase()!=='B2B' && !r.pool;});
    if(state.chartTeam!=='All') rrows=rrows.filter(function(r){return r.team===state.chartTeam;});
    function rsum(key){return rrows.reduce(function(a,r){return a+(r[key]||0);},0);}
    var periodLbl=view==='ytd'?'YTD':((rr.mtdMonth||'MTD')+' · MTD');

    if(tgl){
      var bar=el('div','view-bar');
      [['mtd','📅 MTD ('+(rr.mtdMonth||'Current Month')+')'],['ytd','📆 YTD (Full Year)']].forEach(function(v){
        var b=el('button','vbtn'+(view===v[0]?' active':''),v[1]);
        b.onclick=function(){state.dashRevView=v[0];renderDashboard();};
        bar.appendChild(b);
      });
      tgl.appendChild(bar);
    }

    /* Product KPIs from raw sheets — computed first so Total Revenue can use them */
    var mtdMo=(rr.mtdMonth||'').toUpperCase();
    function rawByView(arr){
      var out=view==='ytd'?arr:arr.filter(function(r){return (r.month||'').toUpperCase()===mtdMo;});
      if(state.chartTeam!=='All') out=filterByTeam(out,state.chartTeam);
      return out;
    }
    var dIns=rawByView(DATA.ins), dFees=rawByView(DATA.fees), dPms=rawByView(DATA.pms||[]);
    var rawInsRev=sum(dIns,function(r){return r.revenue;});
    var rawFeesRev=sum(dFees,function(r){return r.revenue;});
    var rawPmsRev=sum(dPms,function(r){return r.revenue;});
    var trailRev=rsum('trail');
    var rawTotal=rawInsRev+rawFeesRev+rawPmsRev+trailRev;

    k.appendChild(kpi(inrCr(rawTotal),'Total Revenue · '+periodLbl,'primary'));

    /* Revenue Target from REVSUM */
    var vk=view==='ytd'?'ytd':'mtd';
    var ot=DATA.revsum.overallTarget||{};
    var revTarget;
    if(state.chartTeam==='All'){
      revTarget=vk==='ytd'?(ot.mtdTarget||0)*12:(ot.mtdTarget||0);
    } else {
      var rsRows=(DATA.revsum.rows||[]).filter(function(r){return r.team===state.chartTeam;});
      revTarget=rsRows.reduce(function(a,r){return a+(r[vk].revTarget||0);},0);
    }
    var achPct=revTarget?rawTotal/revTarget:0;
    k.appendChild(kpi(inrCr(revTarget),'Revenue Target · '+periodLbl,'red'));
    k.appendChild(kpi(revTarget?pct(achPct):'—','% Achievement',achPct>=0.8?'green':achPct>=0.5?'yellow':'red'));

    k.appendChild(kpi(inrCr(rawInsRev),'Insurance Revenue','blue'));
    k.appendChild(kpi(inrCr(rawFeesRev),'FEES Revenue','green'));
    k.appendChild(kpi(inrCr(rawPmsRev),'PMS Revenue','violet'));
    k.appendChild(kpi(inrCr(trailRev),'Trail Revenue (MF)','orange'));
    k.appendChild(kpi(numFmt(rrows.length),'RMs Reporting','cyan'));
  }else{
    /* fallback: old raw-sheet KPIs when revreport isn't in data.js */
    k.appendChild(kpi(inrCr(insRev+feeRev+pmsRev),'Total Revenue (INS+FEES+PMS)','primary'));
    k.appendChild(kpi(inrCr(insRev),'Insurance Revenue','blue'));
    k.appendChild(kpi(inrCr(feeRev),'FEES Revenue','green'));
    k.appendChild(kpi(inrCr(pmsRev),'PMS Revenue','violet'));
    k.appendChild(kpi(numFmt(ins.length),'Insurance Policies','orange'));
    k.appendChild(kpi(inrCr(insPrem),'Total Premium Booked','cyan'));
  }

  var k2=$('dash-kpis2'); k2.innerHTML='';
  var dView=state.dashRevView||'mtd';
  var rrMtd=((DATA.revreport||{}).mtdMonth||'');
  function talkMonthSum(field){
    return talkRows.reduce(function(acc,r){
      var arr=r[field]||[];
      if(state.month!=='All'){
        var idx=talkMonths.indexOf(state.month);
        return acc+(idx>=0?(arr[idx]||0):0);
      } else if(dView==='mtd'){
        var idx=-1;
        for(var i=0;i<talkMonths.length;i++){if(talkMonths[i].toUpperCase()===rrMtd.toUpperCase()){idx=i;break;}}
        return acc+(idx>=0?(arr[idx]||0):0);
      } else {
        return acc+arr.reduce(function(a,v){return a+(v||0);},0);
      }
    },0);
  }
  var netSaleVal=talkMonthSum('netSale');
  var sipVal=talkMonthSum('sip');
  var cliVal=talkMonthSum('clients');
  var talkLbl=state.month!=='All'?mtdLabel(state.month):dView==='mtd'?'MTD ('+rrMtd+')':'YTD';
  k2.appendChild(kpi(inrCr(netSaleVal),'Net Sale — '+talkLbl,'green'));
  k2.appendChild(kpi(inrCr(sipVal),'SIP Book — '+talkLbl,'blue'));
  k2.appendChild(kpi(numFmt(cliVal),'Clients Added — '+talkLbl,'orange'));
  k2.appendChild(kpi(numFmt(talkRows.length),'Active RMs','violet'));

  /* monthly trend — MTD labels */
  var allMonths=uniqInFyOrder(
    [].concat(ins.map(function(r){return r.month;}),fees.map(function(r){return r.month;}),pms.map(function(r){return r.month;}))
  );
  function msum(arr,m){return arr.filter(function(r){return r.month===m;}).reduce(function(a,r){return a+r.revenue;},0);}
  drawBar('dash-trend-total',allMonths.map(mtdLabel),[
    {label:'Insurance',data:allMonths.map(function(m){return msum(ins,m);}),backgroundColor:ACCENTS[0]},
    {label:'FEES',data:allMonths.map(function(m){return msum(fees,m);}),backgroundColor:ACCENTS[1]},
    {label:'PMS',data:allMonths.map(function(m){return msum(pms,m);}),backgroundColor:ACCENTS[2]}
  ],{money:true,datalabels:false,stacked:true});
  drawBar('dash-trend-ins',allMonths.map(mtdLabel),[
    {label:'Insurance',data:allMonths.map(function(m){return msum(ins,m);}),backgroundColor:ACCENTS[0]}
  ],{money:true,datalabels:false});
  drawBar('dash-trend-fees',allMonths.map(mtdLabel),[
    {label:'FEES',data:allMonths.map(function(m){return msum(fees,m);}),backgroundColor:ACCENTS[1]}
  ],{money:true,datalabels:false});
  drawBar('dash-trend-pms',allMonths.map(mtdLabel),[
    {label:'PMS',data:allMonths.map(function(m){return msum(pms,m);}),backgroundColor:ACCENTS[2]}
  ],{money:true,datalabels:false});

  /* team table */
  var tmap={};
  function addTeam(arr,key){
    arr.forEach(function(r){
      var t=teamOf(r.rm); if(!tmap[t])tmap[t]={team:t,ins:0,fee:0,pms:0,cnt:0};
      tmap[t][key]+=r.revenue; tmap[t].cnt++;
    });
  }
  addTeam(ins,'ins'); addTeam(fees,'fee'); addTeam(pms,'pms');
  var teamRows=Object.keys(tmap).map(function(k){return tmap[k];}).sort(function(a,b){return (b.ins+b.fee+b.pms)-(a.ins+a.fee+a.pms);});
  var tt=$('dash-team'); tt.innerHTML='';
  var teamNote=el('p','table-note','Filtered by Month dropdown — '+(state.month==='All'?'All Months':mtdLabel(state.month)));
  teamNote.style.cssText='font-size:0.75rem;color:#667;margin:0 0 0.4rem;font-style:italic;';
  tt.appendChild(teamNote);
  tt.appendChild(buildTable([
    {label:'Team',get:function(r){return r.team;}},
    {label:'INS Rev',get:function(r){return r.ins;},fmt:inrCr,align:'right',sum:true},
    {label:'FEES Rev',get:function(r){return r.fee;},fmt:inrCr,align:'right',sum:true},
    {label:'PMS Rev',get:function(r){return r.pms;},fmt:inrCr,align:'right',sum:true},
    {label:'Total',get:function(r){return r.ins+r.fee+r.pms;},fmt:inrCr,align:'right',sum:true},
    {label:'Txns',get:function(r){return r.cnt;},fmt:numFmt,align:'right',sum:true}
  ],teamRows,{grand:true,heat:4}));

  /* monthly table with MTD label */
  var months=uniqInFyOrder([].concat(ins.map(function(r){return r.month;}),fees.map(function(r){return r.month;}),pms.map(function(r){return r.month;})));
  var mt=$('dash-month'); mt.innerHTML='';
  mt.appendChild(buildTable([
    {label:'Month',get:function(r){return mtdLabel(r);}},
    {label:'INS',get:function(r){return msum(ins,r);},fmt:inr,align:'right',sum:true},
    {label:'FEES',get:function(r){return msum(fees,r);},fmt:inr,align:'right',sum:true},
    {label:'PMS',get:function(r){return msum(pms,r);},fmt:inr,align:'right',sum:true},
    {label:'Total',get:function(r){return msum(ins,r)+msum(fees,r)+msum(pms,r);},fmt:inr,align:'right',sum:true}
  ],months,{grand:true}));

  $('app-subtitle').textContent=DATA.meta.source+'  ·  Generated: '+DATA.meta.generated;
}

/* ============================================================
   REVENUE SUMMARY V2.0
   ============================================================ */
function renderRevsum(){
  var view=state.revsumView; var label=view==='mtd'?(DATA.revsum.mtdMonth||'MTD'):'YTD';
  var all=DATA.revsum.rows;
  var rows=state.revsumTeam==='All'?all:all.filter(function(r){return r.team===state.revsumTeam;});
  var V=function(r){return r[view];};
  $('revsum-period').textContent=label;

  var revT=sum(rows,function(r){return V(r).revTarget;}), revA=sum(rows,function(r){return V(r).revAch;});
  var aum=sum(rows,function(r){return V(r).totalAum;}), mf=sum(rows,function(r){return V(r).mfNetSale;});
  var cli=sum(rows,function(r){return V(r).totalClients;});
  var k=$('revsum-kpis'); k.innerHTML='';
  k.appendChild(kpi(inrCr(revA),'Revenue Achieved ('+label+')','orange'));
  k.appendChild(kpi(inrCr(revT),'Revenue Target','blue'));
  k.appendChild(kpi(revT?pct(revA/revT):'—','Overall % Achievement',achColor(revT?revA/revT:0)));
  k.appendChild(kpi(inrCr(aum),'Total AUM','violet'));
  k.appendChild(kpi(inrCr(mf),'MF Net Sale','green'));
  k.appendChild(kpi(numFmt(cli),'Total Clients','pink'));

  /* team summary — chart coloured by %Ach */
  var tmap={}; rows.forEach(function(r){
    var d=V(r); var t=r.team;
    if(!tmap[t])tmap[t]={team:t,n:0,revT:0,revA:0,aum:0,mf:0,sip:0,cli:0};
    var o=tmap[t]; o.n++;o.revT+=d.revTarget;o.revA+=d.revAch;o.aum+=d.totalAum;o.mf+=d.mfNetSale;o.sip+=d.sipAdd;o.cli+=d.totalClients;
  });
  var teamRows=Object.keys(tmap).map(function(k){return tmap[k];}).sort(function(a,b){return b.revA-a.revA;});
  var tm=$('revsum-team'); tm.innerHTML='';
  tm.appendChild(buildTable([
    {label:'Team',get:function(r){return r.team;}},
    {label:'RMs',get:function(r){return r.n;},fmt:numFmt,align:'right',sum:true},
    {label:'Rev Target',get:function(r){return r.revT;},fmt:inrCr,align:'right',sum:true},
    {label:'Rev Achieved',get:function(r){return r.revA;},fmt:inrCr,align:'right',sum:true},
    {label:'% Ach',get:function(r){return r.revT?r.revA/r.revT:0;},fmt:function(v){return achBadge(v);},align:'right'},
    {label:'Total AUM',get:function(r){return r.aum;},fmt:inrCr,align:'right',sum:true},
    {label:'MF Net Sale',get:function(r){return r.mf;},fmt:inrCr,align:'right',sum:true},
    {label:'Net SIP',get:function(r){return r.sip;},fmt:inr,align:'right',sum:true},
    {label:'Clients',get:function(r){return r.cli;},fmt:numFmt,align:'right',sum:true}
  ],teamRows,{grand:true,rowCls:function(r){return achRowCls(r.revT?r.revA/r.revT:0);}}));

  /* Top RMs chart — labelled bars */
  var topRm=rows.slice().sort(function(a,b){return V(b).revAch-V(a).revAch;}).slice(0,15);
  drawBar('revsum-chart',topRm.map(function(r){return r.name;}),[
    {label:'Achieved',data:topRm.map(function(r){return V(r).revAch;}),backgroundColor:topRm.map(function(r){return achBarColor(V(r).revTarget?V(r).revAch/V(r).revTarget:0);})},
    {label:'Target',data:topRm.map(function(r){return V(r).revTarget;}),backgroundColor:'#cbd5e1'}
  ],{money:true,datalabels:true});

  /* append ach legend below revsum-chart */
  var chartSect=$('revsum-chart').closest?$('revsum-chart').closest('.section'):null;
  if(chartSect&&!chartSect.querySelector('.ach-legend'))chartSect.appendChild(achLegendEl());

  /* RM detail — rows coloured by %Ach */
  var detail=rows.slice().sort(function(a,b){return V(b).revAch-V(a).revAch;});
  var cols=[
    {label:'RM',get:function(r){return r.name;}},
    {label:'Team',get:function(r){return r.team;}},
    {label:'Lvl',get:function(r){return r.level;}},
    {label:'KRA',get:function(r){return r.kra;}},
    {label:'Rev Target',get:function(r){return V(r).revTarget;},fmt:inr,align:'right',sum:true},
    {label:'Rev Achieved',get:function(r){return V(r).revAch;},fmt:inr,align:'right',sum:true},
    {label:'% Ach',get:function(r){return V(r).revTarget?V(r).revAch/V(r).revTarget:0;},fmt:function(v){return achBadge(v);},align:'right'},
    {label:'Clients',get:function(r){return V(r).totalClients;},fmt:numFmt,align:'right',sum:true},
    {label:'AUM Target',get:function(r){return V(r).aumTarget;},fmt:inrCr,align:'right',sum:true},
    {label:'Total AUM',get:function(r){return V(r).totalAum;},fmt:inrCr,align:'right',sum:true},
    {label:'MF Net Sale',get:function(r){return V(r).mfNetSale;},fmt:inrCr,align:'right',sum:true},
    {label:'PMS/AIF',get:function(r){return V(r).pmsAif;},fmt:inrCr,align:'right',sum:true},
    {label:'SIP Target',get:function(r){return V(r).sipTarget;},fmt:inr,align:'right',sum:true},
    {label:'Net SIP',get:function(r){return V(r).sipAdd;},fmt:inr,align:'right',sum:true},
    {label:'GT Hours',get:function(r){return V(r).gtTime;},fmt:hours,align:'right'}
  ];
  if(view==='ytd') cols.splice(6,0,{label:'Deficit',get:function(r){return V(r).revDeff;},fmt:inr,align:'right',sum:true});
  var rd=$('revsum-detail'); rd.innerHTML='';
  rd.appendChild(buildTable(cols,detail,{grand:true,rowCls:function(r){
    var ratio=V(r).revTarget?V(r).revAch/V(r).revTarget:0;
    return achRowCls(ratio);
  }}));

  /* ach legend below RM detail */
  if(!rd.nextElementSibling||!rd.nextElementSibling.classList.contains('ach-legend')){
    rd.parentNode.appendChild(achLegendEl());
  }
}
function achColor(r){return r>=0.8?'green':r>=0.5?'yellow':'red';}
function achBadge(v){
  var c=v>=0.8?'var(--green)':v>=0.5?'var(--yellow)':'var(--red)';
  return '<span style="font-weight:700;color:'+c+'">'+pct(v)+'</span>';
}

/* ============================================================
   INS TAB
   ============================================================ */
var INS_VIEWS=[
  {id:'rm',    label:'🏅 RM Analysis'},
  {id:'product',label:'📦 Product Breakdown'},
  {id:'partner',label:'🏢 Partner & Source'},
  {id:'monthly',label:'📅 Monthly Trend'},
  {id:'raw',   label:'🔍 Raw Data'}
];

function renderINS(){
  var pane=$('tab-ins'); pane.innerHTML='';
  var ins=getINS();
  var totalRev=sum(ins,function(r){return r.revenue;});
  var totalPrem=sum(ins,function(r){return r.premium;});
  var confRev=sum(ins.filter(function(r){return r.confirmed==='CONFIRMED';}),function(r){return r.revenue;});
  var cancelled=ins.filter(function(r){return r.status==='Cancelled'||r.confirmed==='Cancelled';}).length;

  pane.appendChild(kpiRow([
    [inrCr(totalRev),'Insurance Revenue','primary'],
    [numFmt(ins.length),'Total Policies','blue'],
    [inrCr(totalPrem),'Total Premium','cyan'],
    [inrCr(confRev),'Confirmed Revenue','green'],
    [inrCr(avg(totalRev,ins.length)),'Avg Rev / Policy','orange'],
    [numFmt(cancelled),'Cancelled','red']
  ]));
  pane.appendChild(makeViewBar(INS_VIEWS,state.insView,function(v){state.insView=v;renderINS();}));

  if(state.insView==='rm')     insRmView(pane,ins,totalRev);
  else if(state.insView==='product') insProductView(pane,ins,totalRev);
  else if(state.insView==='partner') insPartnerView(pane,ins,totalRev);
  else if(state.insView==='monthly') insMonthlyView(pane,ins);
  else insRawView(pane,ins);
}

/* assign a fixed revenue rank (1 = highest) to each pivot row, so the #
   column stays put no matter how the table is later re-sorted */
function assignRevRank(byRM){
  byRM.slice().sort(byRev).forEach(function(r,i){r._rank=i+1;});
  return byRM;
}
/* Revenue-by-Team bar chart, shared by INS / FEES / PMS RM Analysis */
function teamRevChart(pane,rows,canvasId,barCls){
  var byTeam=pivot(rows,function(r){return teamOf(r.rm);},{rev:function(r){return r.revenue;}}).sort(byRev);
  var s=sect('Revenue by Team',barCls,canvasId);
  s.appendChild(cw(canvasId,300)); pane.appendChild(s);
  drawBar(canvasId,byTeam.map(function(r){return r.key;}),[
    {label:'Revenue',data:byTeam.map(function(r){return r.rev;}),
     backgroundColor:byTeam.map(function(_,i){return ACCENTS[i%ACCENTS.length];})}
  ],{money:true,legend:false,countData:byTeam.map(function(r){return r._count;})});
}

function insRmView(pane,ins,totalRev){
  var byRM=assignRevRank(pivot(ins,function(r){return r.rm;},{rev:function(r){return r.revenue;},prem:function(r){return r.premium;}}));
  var tot=totalRev||1;
  var teamSel=state.chartTeam;

  /* chart always ranked by Revenue (desc); title reflects team filter */
  var byRev15=byRM.slice().sort(byRev).slice(0,15);
  var chartTitle=(teamSel&&teamSel!=='All')
    ? 'RM Rankings — '+teamSel
    : 'Top 15 RMs by Revenue';
  var s1=sect(chartTitle,'bar-orange','ins-rm-chart');
  s1.appendChild(cw('ins-rm-chart',420));
  pane.appendChild(s1);
  drawHBar('ins-rm-chart',byRev15.map(function(r){return r.key;}),[
    {label:'Revenue',data:byRev15.map(function(r){return r.rev;}),backgroundColor:ACCENTS[0]},
    {label:'Premium',data:byRev15.map(function(r){return r.prem;}),backgroundColor:ACCENTS[1]}
  ],{money:true});

  /* revenue by team */
  teamRevChart(pane,ins,'ins-team-chart','bar-green');

  /* full list — sortable columns (click header to toggle asc/desc) */
  if(!state.insRmSort) state.insRmSort={col:'rev',dir:-1};
  var ss=state.insRmSort;
  var sorted=byRM.slice().sort(function(a,b){
    var va,vb;
    if(ss.col==='rm'){ va=(a.key||'').toLowerCase(); vb=(b.key||'').toLowerCase(); return (va<vb?-1:va>vb?1:0)*ss.dir; }
    if(ss.col==='team'){ va=teamOf(a.key).toLowerCase(); vb=teamOf(b.key).toLowerCase(); return (va<vb?-1:va>vb?1:0)*ss.dir; }
    if(ss.col==='count'){ va=a._count; vb=b._count; }
    else if(ss.col==='prem'){ va=a.prem; vb=b.prem; }
    else if(ss.col==='avg'){ va=avg(a.rev,a._count); vb=avg(b.rev,b._count); }
    else { va=a.rev; vb=b.rev; }   /* rev, pct */
    return (va-vb)*ss.dir;
  });

  var s2=sect('RM Rankings — Full List','bar-blue','ins-rm-table');
  var tw=el('div','table-wrap tall');
  tw.appendChild(buildTable([
    {label:'#',   get:function(r){return r._rank;},align:'right',isRank:true},
    {label:'RM',  sortKey:'rm',   get:function(r){return r.key;}},
    {label:'Team',sortKey:'team', get:function(r){return teamOf(r.key);}},
    {label:'Policies',sortKey:'count',get:function(r){return r._count;},fmt:numFmt,align:'right',sum:true},
    {label:'Premium',sortKey:'prem',get:function(r){return r.prem;},fmt:inr,align:'right',sum:true},
    {label:'Revenue',sortKey:'rev',get:function(r){return r.rev;},fmt:inr,align:'right',sum:true},
    {label:'% Total',sortKey:'pct',get:function(r){return r.rev/tot;},fmt:pct,align:'right'},
    {label:'Avg/Policy',sortKey:'avg',get:function(r){return avg(r.rev,r._count);},fmt:inr,align:'right'}
  ],sorted,{grand:true,heat:5,
    sortState:{col:ss.col,dir:ss.dir},
    onSort:function(k){
      if(state.insRmSort.col===k) state.insRmSort.dir*=-1;
      else state.insRmSort={col:k,dir:-1};
      renderINS();
    }
  }));
  s2.appendChild(tw); pane.appendChild(s2);

  /* ── Insurance Product Mix — team-wise or RM-wise ── */
  insProductMix(pane,ins,teamSel);
}

function insProductMix(pane,ins,teamSel){
  var isTeamView=(!teamSel||teamSel==='All');
  var title=isTeamView?'Insurance Product Mix — by Team':'Insurance Product Mix — '+teamSel+' (RM-wise)';
  var s=sect(title,'bar-cyan','ins-product-mix');
  var tw=el('div','table-wrap tall');

  if(isTeamView){
    var map={};
    ins.forEach(function(r){
      var team=teamOf(r.rm); var it=r.insType||'Unknown';
      var k=team+'|||'+it;
      if(!map[k]) map[k]={team:team,insType:it,count:0,prem:0};
      map[k].count++; map[k].prem+=r.premium;
    });
    var rows=Object.keys(map).map(function(k){return map[k];}).sort(function(a,b){
      var tc=a.team.localeCompare(b.team); if(tc!==0)return tc;
      return b.prem-a.prem;
    });
    tw.appendChild(buildTable([
      {label:'Team',get:function(r){return r.team;}},
      {label:'Insurance Type',get:function(r){return r.insType;}},
      {label:'Policies',get:function(r){return r.count;},fmt:numFmt,align:'right',sum:true},
      {label:'Total Premium',get:function(r){return r.prem;},fmt:inr,align:'right',sum:true},
      {label:'Avg Premium/Client',get:function(r){return avg(r.prem,r.count);},fmt:inr,align:'right'}
    ],rows,{grand:true,heat:3}));
  }else{
    var map2={};
    ins.forEach(function(r){
      var rm=r.rm||'Unknown'; var it=r.insType||'Unknown';
      var k=rm+'|||'+it;
      if(!map2[k]) map2[k]={rm:rm,insType:it,count:0,prem:0};
      map2[k].count++; map2[k].prem+=r.premium;
    });
    var rows2=Object.keys(map2).map(function(k){return map2[k];}).sort(function(a,b){
      var rc=a.rm.localeCompare(b.rm); if(rc!==0)return rc;
      return b.prem-a.prem;
    });
    tw.appendChild(buildTable([
      {label:'RM',get:function(r){return r.rm;}},
      {label:'Insurance Type',get:function(r){return r.insType;}},
      {label:'Policies',get:function(r){return r.count;},fmt:numFmt,align:'right',sum:true},
      {label:'Total Premium',get:function(r){return r.prem;},fmt:inr,align:'right',sum:true},
      {label:'Avg Premium/Client',get:function(r){return avg(r.prem,r.count);},fmt:inr,align:'right'}
    ],rows2,{grand:true,heat:3}));
  }
  s.appendChild(tw); pane.appendChild(s);
}

/* Product Breakdown — only Insurance Category chart with count in tooltip */
function insProductView(pane,ins,totalRev){
  var byInsType=pivot(ins,function(r){return r.insType;},{rev:function(r){return r.revenue;}}).sort(byRev);
  var byType=pivot(ins,function(r){return r.type;},{rev:function(r){return r.revenue;}}).sort(byRev);
  var tot=totalRev||1;

  /* single chart — Revenue by Insurance Category with count in tooltip */
  var s1=sect('Revenue by Insurance Category','bar-orange','ins-instype-chart');
  s1.appendChild(cw('ins-instype-chart',320)); pane.appendChild(s1);
  drawBar('ins-instype-chart',byInsType.map(function(r){return r.key;}),[
    {label:'Revenue',data:byInsType.map(function(r){return r.rev;}),backgroundColor:byInsType.map(function(_,i){return ACCENTS[i%ACCENTS.length];})}
  ],{money:true,legend:false,countData:byInsType.map(function(r){return r._count;})});

  /* two tables */
  var g2=el('div','grid cols-2');
  var t1=sect('By Insurance Category','bar-green','ins-instype-table');
  var tw1=el('div','table-wrap'); tw1.appendChild(revShareTable(byInsType,tot)); t1.appendChild(tw1); g2.appendChild(t1);
  var t2=sect('By Product Type','bar-violet','ins-type-table');
  var tw2=el('div','table-wrap'); tw2.appendChild(revShareTable(byType,tot)); t2.appendChild(tw2); g2.appendChild(t2);
  pane.appendChild(g2);
}

function insPartnerView(pane,ins,totalRev){
  var byPartner=pivot(ins,function(r){return r.partner;},{rev:function(r){return r.revenue;}}).sort(byRev);
  var byCategory=pivot(ins,function(r){return r.category;},{rev:function(r){return r.revenue;}}).sort(byRev);
  var tot=totalRev||1; var top10=byPartner.slice(0,10);

  var s1=sect('Top 10 Insurance Partners by Revenue','bar-orange','ins-partner-chart');
  s1.appendChild(cw('ins-partner-chart',320));
  pane.appendChild(s1);
  drawBar('ins-partner-chart',top10.map(function(r){return r.key;}),[
    {label:'Revenue',data:top10.map(function(r){return r.rev;}),backgroundColor:top10.map(function(_,i){return ACCENTS[i%ACCENTS.length];})}
  ],{money:true,legend:false,countData:top10.map(function(r){return r._count;})});

  var g=el('div','grid cols-2');
  var t1=sect('By Insurance Partner','bar-blue','ins-partner-table');
  var tw1=el('div','table-wrap'); tw1.appendChild(revShareTable(byPartner,tot)); t1.appendChild(tw1); g.appendChild(t1);
  var t2=sect('By Source / Category','bar-green','ins-category-table');
  var tw2=el('div','table-wrap'); tw2.appendChild(revShareTable(byCategory,tot)); t2.appendChild(tw2); g.appendChild(t2);
  pane.appendChild(g);
}

/* ── Shared MTD day-window helpers (INS / FEES / PMS Monthly Trend) ── */
function dayWindowBar(prefix){
  var sK=prefix+'MtdStart', eK=prefix+'MtdEnd';
  if(state[sK]==null) state[sK]=1;
  if(state[eK]==null) state[eK]=31;
  var render={ins:renderINS,fees:renderFEES,pms:renderPMS}[prefix];
  var fbar=el('div','section bar-teal');
  fbar.appendChild(el('h2',null,'Day Window — MTD by Create Date'));
  var ctr=el('div','filters'); ctr.style.marginTop='8px';
  ctr.appendChild(el('label',null,'Day'));
  var inS=el('input'); inS.type='number'; inS.min=1; inS.max=31; inS.value=state[sK]; inS.style.width='70px';
  ctr.appendChild(inS);
  ctr.appendChild(el('span',null,'<span style="color:var(--muted);padding:0 4px">to</span>'));
  var inE=el('input'); inE.type='number'; inE.min=1; inE.max=31; inE.value=state[eK]; inE.style.width='70px';
  ctr.appendChild(inE);
  ctr.appendChild(el('span',null,'<span style="color:var(--muted);font-size:12px;margin-left:10px">Compares the same day range across every month (by Create Date)</span>'));
  inS.onchange=function(){state[sK]=Math.max(1,Math.min(31,parseInt(this.value,10)||1)); render();};
  inE.onchange=function(){state[eK]=Math.max(1,Math.min(31,parseInt(this.value,10)||31)); render();};
  fbar.appendChild(ctr);
  return fbar;
}
function applyDayWindow(rows,prefix){
  var sd=state[prefix+'MtdStart']||1, ed=state[prefix+'MtdEnd']||31;
  var win=rows.filter(function(r){
    var cd=r.createDate||'';
    if(cd.length<10) return false;
    var day=parseInt(cd.substring(8,10),10);
    return day>=sd && day<=ed;
  });
  return {rows:win, days:(ed-sd+1)||1};
}

function insMonthlyView(pane,ins){
  pane.appendChild(dayWindowBar('ins'));
  var w=applyDayWindow(ins,'ins');
  var win=w.rows;
  var sd=state.insMtdStart, ed=state.insMtdEnd;

  var months=uniqInFyOrder(win.map(function(r){return r.month;}));
  function mrev(m){return sum(win.filter(function(r){return r.month===m;}),function(r){return r.revenue;});}
  function mprem(m){return sum(win.filter(function(r){return r.month===m;}),function(r){return r.premium;});}

  var s1=sect('Monthly Insurance Revenue (Day '+sd+'–'+ed+')','bar-blue','ins-monthly-chart');
  s1.appendChild(cw('ins-monthly-chart',320)); pane.appendChild(s1);
  drawBar('ins-monthly-chart',months.map(mtdLabel),[
    {label:'Revenue',data:months.map(mrev),backgroundColor:ACCENTS[0]},
    {label:'Premium',data:months.map(mprem),backgroundColor:ACCENTS[1]}
  ],{money:true});

  var g=el('div','grid cols-2');
  var byMonth=fyOrder(pivot(win,function(r){return r.month;},{rev:function(r){return r.revenue;},prem:function(r){return r.premium;}}));
  var sm=sect('Monthly Breakdown','bar-green','ins-monthly-table');
  var tmw=el('div','table-wrap');
  tmw.appendChild(buildTable([
    {label:'Month',get:function(r){return mtdLabel(r.key);}},
    {label:'Policies',get:function(r){return r._count;},fmt:numFmt,align:'right',sum:true},
    {label:'Premium',get:function(r){return r.prem;},fmt:inr,align:'right',sum:true},
    {label:'Revenue',get:function(r){return r.rev;},fmt:inr,align:'right',sum:true}
  ],byMonth,{grand:true}));
  sm.appendChild(tmw); g.appendChild(sm);

  var byStatus=pivot(win,function(r){return r.status;},{rev:function(r){return r.revenue;}}).sort(byRev);
  var tot=sum(win,function(r){return r.revenue;})||1;
  var ss=sect('Status Distribution','bar-orange','ins-status-table');
  var tw2=el('div','table-wrap'); tw2.appendChild(revShareTable(byStatus,tot)); ss.appendChild(tw2); g.appendChild(ss);
  pane.appendChild(g);
}

/* INS Raw Data — Excel-like per-column filter menus (Marketing MIS style) */
function insRawView(pane,ins){
  var s=sect('INS — Raw Policy Data','bar-red','ins-raw');

  /* state.insColFilters: {colKey: [vals]} (exact match) OR {colKey: {type:'text',value:'q'}} */
  if(!state.insColFilters) state.insColFilters={};
  var sortState=state.insRawSort;

  /* column definitions */
  var COLS=[
    {key:'rm',         label:'RM',             type:'cat', fmt:function(r){return r.rm;}},
    {key:'team',       label:'Team',           type:'cat', fmt:function(r){return teamOf(r.rm);}},
    {key:'client',     label:'Client',         type:'cat', fmt:function(r){return r.client;}},
    {key:'insType',    label:'Insurance Type', type:'cat', fmt:function(r){return r.insType;}},
    {key:'partner',    label:'Partner',        type:'cat', fmt:function(r){return r.partner;}},
    {key:'type',       label:'Product Type',   type:'cat', fmt:function(r){return r.type;}},
    {key:'category',   label:'Category',       type:'cat', fmt:function(r){return r.category;}},
    {key:'month',      label:'Month',          type:'cat', fmt:function(r){return r.month;}},
    {key:'premium',    label:'Premium',        type:'num', fmt:function(r){return r.premium;}},
    {key:'revenue',    label:'Revenue',        type:'num', fmt:function(r){return r.revenue;}},
    {key:'status',     label:'Status',         type:'cat', fmt:function(r){return r.status;}}
  ];

  /* search input + clear filters */
  var frow=el('div','raw-filters');
  var inp=el('input'); inp.type='text'; inp.placeholder='Global search…'; inp.style.minWidth='240px';
  inp.value=state.insSearch;
  inp.oninput=function(){state.insSearch=this.value;renderTable();};
  frow.appendChild(inp);
  var clearBtn=document.createElement('button'); clearBtn.className='secondary'; clearBtn.textContent='Clear Column Filters';
  clearBtn.onclick=function(){state.insColFilters={};renderTable();};
  frow.appendChild(clearBtn);
  var meta=el('span','meta'); meta.style.marginLeft='8px';
  frow.appendChild(meta);
  s.appendChild(frow);

  var tbody=el('div'); s.appendChild(tbody); pane.appendChild(s);

  /* compute rows that pass ALL OTHER column filters (dependent filtering) */
  function rowsAfterFilters(skipCol){
    var q=(state.insSearch||'').toLowerCase();
    return ins.filter(function(r){
      if(q){
        var hay=(r.rm+' '+r.client+' '+r.insType+' '+r.partner+' '+r.type+' '+r.category+' '+r.status+' '+r.month).toLowerCase();
        if(hay.indexOf(q)<0)return false;
      }
      for(var i=0;i<COLS.length;i++){
        var c=COLS[i]; if(c.key===skipCol)continue;
        var f=state.insColFilters[c.key];
        if(!f)continue;
        var v=c.fmt(r);
        if(f.type==='text'){
          if(((''+v).toLowerCase()).indexOf(f.value.toLowerCase())<0)return false;
        } else {
          /* array of selected values; empty = none allowed */
          if(!f.length)return false;
          if(f.indexOf(''+v)<0)return false;
        }
      }
      return true;
    });
  }

  function uniqVals(colKey){
    var rows=rowsAfterFilters(colKey);
    var col=COLS.filter(function(c){return c.key===colKey;})[0];
    var seen={}; var out=[];
    rows.forEach(function(r){var v=''+col.fmt(r); if(!(v in seen)){seen[v]=1;out.push(v);}});
    out.sort(function(a,b){
      if(col.type==='num'){return parseFloat(a)-parseFloat(b);}
      return a.localeCompare(b);
    });
    return out;
  }

  function closeMenu(){
    var m=document.getElementById('raw-filter-menu-active');
    if(m)m.remove();
  }

  function openMenu(colKey,anchor){
    closeMenu();
    var col=COLS.filter(function(c){return c.key===colKey;})[0];
    var existing=state.insColFilters[colKey];
    var allVals=uniqVals(colKey);
    var selected; var isText=existing&&existing.type==='text';
    if(isText){selected=allVals.slice();}
    else if(existing){selected=existing.slice();}
    else selected=allVals.slice();
    /* dedupe selected to those that exist */
    var validSet={}; allVals.forEach(function(v){validSet[v]=1;});
    selected=selected.filter(function(v){return v in validSet;});

    var menu=document.createElement('div');
    menu.className='raw-filter-menu';
    menu.id='raw-filter-menu-active';

    var title=document.createElement('div');
    title.className='raw-filter-title';
    title.textContent='Filter: '+col.label;
    menu.appendChild(title);

    var search=document.createElement('input');
    search.className='raw-filter-search';
    search.type='text';
    search.placeholder='Search values…  (Enter for "contains")';
    if(isText)search.value=existing.value;
    menu.appendChild(search);

    var actions=document.createElement('div');
    actions.className='raw-filter-actions';
    var bAll=document.createElement('button'); bAll.textContent='All';
    var bNone=document.createElement('button'); bNone.textContent='None';
    actions.appendChild(bAll); actions.appendChild(bNone);
    menu.appendChild(actions);

    var optsBox=document.createElement('div');
    optsBox.className='raw-filter-options';
    menu.appendChild(optsBox);

    var footer=document.createElement('div');
    footer.className='raw-filter-footer';
    var bReset=document.createElement('button'); bReset.textContent='Reset';
    var bApply=document.createElement('button'); bApply.textContent='Apply';
    footer.appendChild(bReset); footer.appendChild(bApply);
    menu.appendChild(footer);

    var filtered=allVals.slice();
    function renderOpts(){
      optsBox.innerHTML='';
      if(!filtered.length){
        var em=document.createElement('div'); em.className='raw-filter-empty'; em.textContent='No values';
        optsBox.appendChild(em); return;
      }
      filtered.forEach(function(v){
        var opt=document.createElement('label'); opt.className='raw-filter-option';
        var cb=document.createElement('input'); cb.type='checkbox'; cb.value=v;
        cb.checked=selected.indexOf(v)>=0;
        cb.onchange=function(){
          if(this.checked){if(selected.indexOf(v)<0)selected.push(v);}
          else{selected=selected.filter(function(x){return x!==v;});}
        };
        var sp=document.createElement('span');
        sp.textContent= col.type==='num'? (v==='0'?'0':inr(parseFloat(v))) : v;
        sp.title=v;
        opt.appendChild(cb); opt.appendChild(sp);
        optsBox.appendChild(opt);
      });
    }
    renderOpts();

    search.oninput=function(){
      var q=this.value.toLowerCase();
      filtered=allVals.filter(function(v){return v.toLowerCase().indexOf(q)>=0;});
      renderOpts();
    };
    search.onkeydown=function(ev){
      if(ev.key==='Enter'&&this.value){
        state.insColFilters[colKey]={type:'text',value:this.value};
        closeMenu(); renderTable();
      }
    };
    bAll.onclick=function(){selected=filtered.slice();renderOpts();};
    bNone.onclick=function(){selected=[];renderOpts();};
    bReset.onclick=function(){delete state.insColFilters[colKey];closeMenu();renderTable();};
    bApply.onclick=function(){
      if(selected.length===allVals.length){delete state.insColFilters[colKey];}
      else{state.insColFilters[colKey]=selected.slice();}
      closeMenu();renderTable();
    };

    document.body.appendChild(menu);
    /* position */
    var r=anchor.getBoundingClientRect();
    var top=r.bottom+4, left=r.left;
    var mw=menu.offsetWidth, mh=menu.offsetHeight;
    if(left+mw>window.innerWidth-8) left=window.innerWidth-mw-8;
    if(top+mh>window.innerHeight-8) top=Math.max(8,r.top-mh-4);
    menu.style.top=top+'px'; menu.style.left=left+'px';

    /* outside click + resize closes */
    setTimeout(function(){
      function onDoc(e){
        if(menu.contains(e.target))return;
        if(e.target.classList&&e.target.classList.contains('raw-filter-btn'))return;
        closeMenu(); document.removeEventListener('click',onDoc); window.removeEventListener('resize',onResize);
      }
      function onResize(){closeMenu();window.removeEventListener('resize',onResize);document.removeEventListener('click',onDoc);}
      document.addEventListener('click',onDoc);
      window.addEventListener('resize',onResize);
    },10);
  }

  function buildHeadEl(col){
    var wrap=document.createElement('div'); wrap.className='raw-filter-head';
    var lbl=document.createElement('span');
    lbl.textContent=col.label;
    lbl.style.cursor='pointer';
    if(sortState.col===col.key) lbl.innerHTML+=' <span class="sort-icon">'+(sortState.dir>0?'▲':'▼')+'</span>';
    lbl.onclick=function(ev){
      ev.stopPropagation();
      if(sortState.col===col.key)sortState.dir*=-1;
      else{sortState.col=col.key;sortState.dir=1;}
      state.insRawSort=sortState; renderTable();
    };
    var btn=document.createElement('button');
    btn.className='raw-filter-btn'+(state.insColFilters[col.key]?' active':'');
    btn.textContent='▼';
    btn.title='Filter '+col.label;
    btn.onclick=function(ev){ev.stopPropagation();openMenu(col.key,btn);};
    wrap.appendChild(lbl); wrap.appendChild(btn);
    var f=state.insColFilters[col.key];
    if(f){
      var sel=document.createElement('div'); sel.className='raw-filter-selected';
      if(f.type==='text') sel.textContent='~'+f.value;
      else sel.textContent=f.length+' selected';
      wrap.appendChild(sel);
    }
    return wrap;
  }

  function renderTable(){
    var data=rowsAfterFilters(null);
    /* sort */
    if(sortState.col){
      var col=COLS.filter(function(c){return c.key===sortState.col;})[0];
      data=data.slice().sort(function(a,b){
        var va=col.fmt(a), vb=col.fmt(b);
        if(col.type==='num') return sortState.dir*(num(vb)-num(va));
        return sortState.dir*(''+va).localeCompare(''+vb);
      });
    }
    meta.textContent=data.length+' / '+ins.length+' rows';
    tbody.innerHTML='';
    var tw=el('div','table-wrap tall');
    var table=document.createElement('table'); table.className='data compact raw-data-table';
    var thead=document.createElement('thead'); var htr=document.createElement('tr');
    /* # column */
    var thN=document.createElement('th'); thN.textContent='#'; htr.appendChild(thN);
    COLS.forEach(function(c){var th=document.createElement('th');th.appendChild(buildHeadEl(c));htr.appendChild(th);});
    thead.appendChild(htr); table.appendChild(thead);
    var tb=document.createElement('tbody');
    data.forEach(function(r,i){
      var tr=document.createElement('tr');
      var tdN=document.createElement('td'); tdN.className='rank'; tdN.textContent=(i+1); tr.appendChild(tdN);
      COLS.forEach(function(c){
        var td=document.createElement('td'); var v=c.fmt(r);
        if(c.type==='num'){td.style.textAlign='right'; td.textContent=inr(v);}
        else{td.textContent=v;}
        tr.appendChild(td);
      });
      tb.appendChild(tr);
    });
    table.appendChild(tb); tw.appendChild(table); tbody.appendChild(tw);
  }
  renderTable();
}

/* ============================================================
   FEES TAB
   ============================================================ */
var FEES_VIEWS=[
  {id:'rm',    label:'🏅 RM Analysis'},
  {id:'product',label:'📦 Product Breakdown'},
  {id:'monthly',label:'📅 Monthly Trend'},
  {id:'raw',   label:'🔍 Raw Data'}
];

function renderFEES(){
  var pane=$('tab-fees'); pane.innerHTML='';
  var fees=getFEES();
  var totalRev=sum(fees,function(r){return r.revenue;});
  var totalAmt=sum(fees,function(r){return r.amount;});
  var confRev=sum(fees.filter(function(r){return r.confirmed==='CONFIRMED';}),function(r){return r.revenue;});

  pane.appendChild(kpiRow([
    [inrCr(totalRev),'Fee Revenue (Net)','primary'],
    [inrCr(totalAmt),'Gross Amount','blue'],
    [numFmt(fees.length),'Fee Transactions','green'],
    [inrCr(confRev),'Confirmed Revenue','green'],
    [inrCr(avg(totalRev,fees.length)),'Avg Fee / Txn','violet'],
    [inrCr(totalAmt-totalRev),'Net Deduction','red']
  ]));
  pane.appendChild(makeViewBar(FEES_VIEWS,state.feesView,function(v){state.feesView=v;renderFEES();}));

  if(state.feesView==='rm')      feesRmView(pane,fees,totalRev);
  else if(state.feesView==='product') feesProductView(pane,fees,totalRev);
  else if(state.feesView==='monthly') feesMonthlyView(pane,fees);
  else feesRawView(pane,fees);
}

function feesRmView(pane,fees,totalRev){
  var byRM=assignRevRank(pivot(fees,function(r){return r.rm;},{rev:function(r){return r.revenue;},amt:function(r){return r.amount;}}));
  var tot=totalRev||1;
  var teamSel=state.chartTeam;

  /* chart always ranked by Revenue (desc); title reflects team filter */
  var byRev15=byRM.slice().sort(byRev).slice(0,15);
  var chartTitle=(teamSel&&teamSel!=='All') ? 'RM Rankings — '+teamSel : 'Top 15 RMs by Fee Revenue';
  var s1=sect(chartTitle,'bar-orange','fees-rm-chart');
  s1.appendChild(cw('fees-rm-chart',420)); pane.appendChild(s1);
  drawHBar('fees-rm-chart',byRev15.map(function(r){return r.key;}),[
    {label:'Net Revenue',data:byRev15.map(function(r){return r.rev;}),backgroundColor:ACCENTS[2]},
    {label:'Gross Amount',data:byRev15.map(function(r){return r.amt;}),backgroundColor:ACCENTS[1]}
  ],{money:true});

  /* revenue by team */
  teamRevChart(pane,fees,'fees-team-chart','bar-green');

  /* full list — sortable columns */
  if(!state.feesRmSort) state.feesRmSort={col:'rev',dir:-1};
  var ss=state.feesRmSort;
  var sorted=byRM.slice().sort(function(a,b){
    var va,vb;
    if(ss.col==='rm'){ va=(a.key||'').toLowerCase(); vb=(b.key||'').toLowerCase(); return (va<vb?-1:va>vb?1:0)*ss.dir; }
    if(ss.col==='team'){ va=teamOf(a.key).toLowerCase(); vb=teamOf(b.key).toLowerCase(); return (va<vb?-1:va>vb?1:0)*ss.dir; }
    if(ss.col==='count'){ va=a._count; vb=b._count; }
    else if(ss.col==='amt'){ va=a.amt; vb=b.amt; }
    else if(ss.col==='avg'){ va=avg(a.rev,a._count); vb=avg(b.rev,b._count); }
    else { va=a.rev; vb=b.rev; }   /* rev, pct */
    return (va-vb)*ss.dir;
  });

  var s2=sect('RM Rankings — Full List','bar-blue','fees-rm-table');
  var tw=el('div','table-wrap tall');
  tw.appendChild(buildTable([
    {label:'#',   get:function(r){return r._rank;},align:'right',isRank:true},
    {label:'RM',  sortKey:'rm',   get:function(r){return r.key;}},
    {label:'Team',sortKey:'team', get:function(r){return teamOf(r.key);}},
    {label:'Txns',sortKey:'count',get:function(r){return r._count;},fmt:numFmt,align:'right',sum:true},
    {label:'Gross',sortKey:'amt', get:function(r){return r.amt;},fmt:inr,align:'right',sum:true},
    {label:'Net Revenue',sortKey:'rev',get:function(r){return r.rev;},fmt:inr,align:'right',sum:true},
    {label:'% Total',sortKey:'pct',get:function(r){return r.rev/tot;},fmt:pct,align:'right'},
    {label:'Avg/Txn',sortKey:'avg',get:function(r){return avg(r.rev,r._count);},fmt:inr,align:'right'}
  ],sorted,{grand:true,heat:5,
    sortState:{col:ss.col,dir:ss.dir},
    onSort:function(k){
      if(state.feesRmSort.col===k) state.feesRmSort.dir*=-1;
      else state.feesRmSort={col:k,dir:-1};
      renderFEES();
    }
  }));
  s2.appendChild(tw); pane.appendChild(s2);
}

function feesProductView(pane,fees,totalRev){
  var byTrans=pivot(fees,function(r){return r.transType;},{rev:function(r){return r.revenue;}}).sort(byRev);
  var byAsset=pivot(fees,function(r){return r.asset+' / '+r.prodType;},{rev:function(r){return r.revenue;}}).sort(byRev);
  var byCategory=pivot(fees,function(r){return r.category;},{rev:function(r){return r.revenue;}}).sort(byRev);
  var tot=totalRev||1;

  var s1=sect('Revenue by Transaction Type','bar-orange','fees-trans-chart');
  s1.appendChild(cw('fees-trans-chart',280)); pane.appendChild(s1);
  drawBar('fees-trans-chart',byTrans.map(function(r){return r.key;}),[
    {label:'Revenue',data:byTrans.map(function(r){return r.rev;}),backgroundColor:byTrans.map(function(_,i){return ACCENTS[i%ACCENTS.length];})}
  ],{money:true,legend:false,countData:byTrans.map(function(r){return r._count;})});

  var g=el('div','grid cols-2');
  var t1=sect('By Transaction Type','bar-blue','fees-trans-table');
  var tw1=el('div','table-wrap'); tw1.appendChild(revShareTable(byTrans,tot)); t1.appendChild(tw1); g.appendChild(t1);
  var t2=sect('By Asset / Product','bar-green','fees-asset-table');
  var tw2=el('div','table-wrap'); tw2.appendChild(revShareTable(byAsset,tot)); t2.appendChild(tw2); g.appendChild(t2);
  pane.appendChild(g);

  var s2=sect('By Source / Category','bar-violet','fees-category-table');
  var tw3=el('div','table-wrap'); tw3.appendChild(revShareTable(byCategory,tot)); s2.appendChild(tw3); pane.appendChild(s2);
}

function feesMonthlyView(pane,fees){
  pane.appendChild(dayWindowBar('fees'));
  var w=applyDayWindow(fees,'fees');
  var win=w.rows;
  var sd=state.feesMtdStart, ed=state.feesMtdEnd;

  var months=uniqInFyOrder(win.map(function(r){return r.month;}));
  function msum(m){return sum(win.filter(function(r){return r.month===m;}),function(r){return r.revenue;});}
  function asum(m){return sum(win.filter(function(r){return r.month===m;}),function(r){return r.amount;});}

  var s1=sect('Monthly Fee Revenue Trend (Day '+sd+'–'+ed+')','bar-blue','fees-monthly-chart');
  s1.appendChild(cw('fees-monthly-chart',300)); pane.appendChild(s1);
  drawBar('fees-monthly-chart',months.map(mtdLabel),[
    {label:'Net Revenue',data:months.map(msum),backgroundColor:ACCENTS[2]},
    {label:'Gross Amount',data:months.map(asum),backgroundColor:ACCENTS[1]}
  ],{money:true});

  var byMonth=fyOrder(pivot(win,function(r){return r.month;},{rev:function(r){return r.revenue;},amt:function(r){return r.amount;}}));
  var s2=sect('Monthly Breakdown','bar-green','fees-monthly-table');
  var tw=el('div','table-wrap');
  tw.appendChild(buildTable([
    {label:'Month',get:function(r){return mtdLabel(r.key);}},
    {label:'Txns',get:function(r){return r._count;},fmt:numFmt,align:'right',sum:true},
    {label:'Gross Amount',get:function(r){return r.amt;},fmt:inr,align:'right',sum:true},
    {label:'Net Revenue',get:function(r){return r.rev;},fmt:inr,align:'right',sum:true}
  ],byMonth,{grand:true}));
  s2.appendChild(tw); pane.appendChild(s2);
}

function feesRawView(pane,fees){
  var s=sect('FEES — Raw Transaction Data','bar-red','fees-raw');
  var frow=el('div','filters'); frow.style.marginBottom='8px';
  var inp=el('input'); inp.type='text'; inp.placeholder='Search RM, client, asset, category…'; inp.style.minWidth='300px';
  inp.value=state.feesSearch;
  var meta=el('span','meta'); meta.style.marginLeft='8px';
  var tbody=el('div');
  function refresh(){
    var q=state.feesSearch.toLowerCase();
    var data=fees.filter(function(r){
      return !q||(r.rm+' '+r.client+' '+r.asset+' '+r.category+' '+r.transType).toLowerCase().indexOf(q)>=0;
    });
    meta.textContent=data.length+' rows';
    tbody.innerHTML='';
    var tw=el('div','table-wrap tall');
    tw.appendChild(buildTable([
      {label:'#',get:function(r,i){return i+1;},align:'right',isRank:true},
      {label:'RM',get:function(r){return r.rm;}},
      {label:'Client',get:function(r){return r.client;}},
      {label:'Asset',get:function(r){return r.asset;}},
      {label:'Prod Type',get:function(r){return r.prodType;}},
      {label:'Category',get:function(r){return r.category;}},
      {label:'Month',get:function(r){return r.month;}},
      {label:'Gross',get:function(r){return r.amount;},fmt:inr,align:'right'},
      {label:'Net Revenue',get:function(r){return r.revenue;},fmt:inr,align:'right'},
      {label:'Trans Type',get:function(r){return r.transType;}}
    ],data,{}));
    tbody.appendChild(tw);
  }
  inp.oninput=function(){state.feesSearch=this.value;refresh();};
  frow.appendChild(inp); frow.appendChild(meta);
  s.appendChild(frow); s.appendChild(tbody); pane.appendChild(s);
  refresh();
}

/* ============================================================
   PMS TAB
   ============================================================ */
var PMS_VIEWS=[
  {id:'rm',    label:'🏅 RM Analysis'},
  {id:'asset', label:'📦 Asset Breakdown'},
  {id:'monthly',label:'📅 Monthly Trend'},
  {id:'raw',   label:'🔍 Raw Data'}
];

function renderPMS(){
  var pane=$('tab-pms'); pane.innerHTML='';
  var pms=getPMS();
  var totalRev=sum(pms,function(r){return r.revenue;});
  var totalInv=sum(pms,function(r){return r.punchAmount;});
  var confRev=sum(pms.filter(function(r){return r.confirmed==='CONFIRMED';}),function(r){return r.revenue;});

  pane.appendChild(kpiRow([
    [inrCr(totalRev),'PMS Revenue','primary'],
    [inrCr(totalInv),'Total Investment (Punch)','blue'],
    [numFmt(pms.length),'Total Transactions','green'],
    [inrCr(confRev),'Confirmed Revenue','teal'],
    [inrCr(avg(totalRev,pms.length)),'Avg Revenue / Txn','cyan'],
    [inrCr(avg(totalInv,pms.length)),'Avg Investment / Txn','orange']
  ]));
  pane.appendChild(makeViewBar(PMS_VIEWS,state.pmsView,function(v){state.pmsView=v;renderPMS();}));

  if(state.pmsView==='rm')      pmsRmView(pane,pms,totalRev);
  else if(state.pmsView==='asset') pmsAssetView(pane,pms,totalRev);
  else if(state.pmsView==='monthly') pmsMonthlyView(pane,pms);
  else pmsRawView(pane,pms);
}

function pmsRmView(pane,pms,totalRev){
  var byRM=assignRevRank(pivot(pms,function(r){return r.rm;},{rev:function(r){return r.revenue;},inv:function(r){return r.punchAmount;}}));
  var tot=totalRev||1;
  var teamSel=state.chartTeam;

  /* chart always ranked by Revenue (desc); title reflects team filter */
  var byRev15=byRM.slice().sort(byRev).slice(0,15);
  var chartTitle=(teamSel&&teamSel!=='All') ? 'RM Rankings — '+teamSel : 'Top 15 RMs by PMS Revenue';
  var s1=sect(chartTitle,'bar-violet','pms-rm-chart');
  s1.appendChild(cw('pms-rm-chart',420)); pane.appendChild(s1);
  drawHBar('pms-rm-chart',byRev15.map(function(r){return r.key;}),[
    {label:'Revenue',data:byRev15.map(function(r){return r.rev;}),backgroundColor:ACCENTS[4]},
    {label:'Investment',data:byRev15.map(function(r){return r.inv;}),backgroundColor:ACCENTS[1]}
  ],{money:true});

  /* revenue by team */
  teamRevChart(pane,pms,'pms-team-chart','bar-green');

  /* full list — sortable columns */
  if(!state.pmsRmSort) state.pmsRmSort={col:'rev',dir:-1};
  var ss=state.pmsRmSort;
  var sorted=byRM.slice().sort(function(a,b){
    var va,vb;
    if(ss.col==='rm'){ va=(a.key||'').toLowerCase(); vb=(b.key||'').toLowerCase(); return (va<vb?-1:va>vb?1:0)*ss.dir; }
    if(ss.col==='team'){ va=teamOf(a.key).toLowerCase(); vb=teamOf(b.key).toLowerCase(); return (va<vb?-1:va>vb?1:0)*ss.dir; }
    if(ss.col==='count'){ va=a._count; vb=b._count; }
    else if(ss.col==='inv'){ va=a.inv; vb=b.inv; }
    else if(ss.col==='avg'){ va=avg(a.rev,a._count); vb=avg(b.rev,b._count); }
    else { va=a.rev; vb=b.rev; }   /* rev, pct */
    return (va-vb)*ss.dir;
  });

  var s2=sect('RM Rankings — Full List','bar-blue','pms-rm-table');
  var tw=el('div','table-wrap tall');
  tw.appendChild(buildTable([
    {label:'#',   get:function(r){return r._rank;},align:'right',isRank:true},
    {label:'RM',  sortKey:'rm',   get:function(r){return r.key;}},
    {label:'Team',sortKey:'team', get:function(r){return teamOf(r.key);}},
    {label:'Txns',sortKey:'count',get:function(r){return r._count;},fmt:numFmt,align:'right',sum:true},
    {label:'Investment',sortKey:'inv',get:function(r){return r.inv;},fmt:inr,align:'right',sum:true},
    {label:'Revenue',sortKey:'rev',get:function(r){return r.rev;},fmt:inr,align:'right',sum:true},
    {label:'% Total',sortKey:'pct',get:function(r){return r.rev/tot;},fmt:pct,align:'right'},
    {label:'Avg/Txn',sortKey:'avg',get:function(r){return avg(r.rev,r._count);},fmt:inr,align:'right'}
  ],sorted,{grand:true,heat:5,
    sortState:{col:ss.col,dir:ss.dir},
    onSort:function(k){
      if(state.pmsRmSort.col===k) state.pmsRmSort.dir*=-1;
      else state.pmsRmSort={col:k,dir:-1};
      renderPMS();
    }
  }));
  s2.appendChild(tw); pane.appendChild(s2);
}

function pmsAssetView(pane,pms,totalRev){
  var byAsset=pivot(pms,function(r){return r.asset;},{rev:function(r){return r.revenue;},inv:function(r){return r.punchAmount;}}).sort(byRev);
  var byScheme=pivot(pms,function(r){return r.scheme;},{rev:function(r){return r.revenue;}}).sort(byRev);
  var tot=totalRev||1;

  var g=el('div','grid cols-2');
  var s1=sect('Revenue by Asset Type','bar-violet','pms-asset-chart');
  s1.appendChild(cw('pms-asset-chart',280)); g.appendChild(s1);
  var s2=sect('Asset Distribution (Revenue)','bar-cyan','pms-asset-donut');
  s2.appendChild(cw('pms-asset-donut',280)); g.appendChild(s2);
  pane.appendChild(g);

  drawBar('pms-asset-chart',byAsset.map(function(r){return r.key;}),[
    {label:'Revenue',data:byAsset.map(function(r){return r.rev;}),backgroundColor:byAsset.map(function(_,i){return ACCENTS[i%ACCENTS.length];})}
  ],{money:true,legend:false,countData:byAsset.map(function(r){return r._count;})});
  drawDoughnut('pms-asset-donut',byAsset.map(function(r){return r.key;}),byAsset.map(function(r){return r.rev;}),{money:true});

  var g2=el('div','grid cols-2');
  var t1=sect('By Asset Type','bar-green','pms-asset-table');
  var tw1=el('div','table-wrap');
  tw1.appendChild(buildTable([
    {label:'Asset',get:function(r){return r.key;}},
    {label:'Txns',get:function(r){return r._count;},fmt:numFmt,align:'right',sum:true},
    {label:'Investment',get:function(r){return r.inv;},fmt:inr,align:'right',sum:true},
    {label:'Revenue',get:function(r){return r.rev;},fmt:inr,align:'right',sum:true},
    {label:'Share %',get:function(r){return r.rev/tot;},fmt:pct,align:'right'}
  ],byAsset,{grand:true,heat:3}));
  t1.appendChild(tw1); g2.appendChild(t1);

  var t2=sect('By Scheme','bar-orange','pms-scheme-table');
  var tw2=el('div','table-wrap'); tw2.appendChild(revShareTable(byScheme,tot)); t2.appendChild(tw2); g2.appendChild(t2);
  pane.appendChild(g2);
}

function pmsMonthlyView(pane,pms){
  pane.appendChild(dayWindowBar('pms'));
  var w=applyDayWindow(pms,'pms');
  var win=w.rows;
  var sd=state.pmsMtdStart, ed=state.pmsMtdEnd;

  var months=uniqInFyOrder(win.map(function(r){return r.month;}));
  function mrev(m){return sum(win.filter(function(r){return r.month===m;}),function(r){return r.revenue;});}
  function minv(m){return sum(win.filter(function(r){return r.month===m;}),function(r){return r.punchAmount;});}

  var s1=sect('Monthly PMS Revenue & Investment (Day '+sd+'–'+ed+')','bar-violet','pms-monthly-chart');
  s1.appendChild(cw('pms-monthly-chart',320)); pane.appendChild(s1);
  drawBar('pms-monthly-chart',months.map(mtdLabel),[
    {label:'Revenue',data:months.map(mrev),backgroundColor:ACCENTS[4]},
    {label:'Investment',data:months.map(minv),backgroundColor:ACCENTS[1]}
  ],{money:true});

  var g=el('div','grid cols-2');
  var byMonth=fyOrder(pivot(win,function(r){return r.month;},{rev:function(r){return r.revenue;},inv:function(r){return r.punchAmount;}}));
  var s2=sect('Monthly Breakdown','bar-green','pms-monthly-table');
  var tw=el('div','table-wrap');
  tw.appendChild(buildTable([
    {label:'Month',get:function(r){return mtdLabel(r.key);}},
    {label:'Txns',get:function(r){return r._count;},fmt:numFmt,align:'right',sum:true},
    {label:'Investment',get:function(r){return r.inv;},fmt:inr,align:'right',sum:true},
    {label:'Revenue',get:function(r){return r.rev;},fmt:inr,align:'right',sum:true}
  ],byMonth,{grand:true}));
  s2.appendChild(tw); g.appendChild(s2);

  var byStatus=pivot(win,function(r){return r.confirmed;},{rev:function(r){return r.revenue;}}).sort(byRev);
  var tot=sum(win,function(r){return r.revenue;})||1;
  var s3=sect('Confirmation Status','bar-orange','pms-confstatus-table');
  var tw2=el('div','table-wrap'); tw2.appendChild(revShareTable(byStatus,tot)); s3.appendChild(tw2); g.appendChild(s3);
  pane.appendChild(g);
}

function pmsRawView(pane,pms){
  var s=sect('PMS — Raw Transaction Data','bar-red','pms-raw');
  var frow=el('div','filters'); frow.style.marginBottom='8px';
  var inp=el('input'); inp.type='text'; inp.placeholder='Search RM, client, asset, scheme…'; inp.style.minWidth='300px';
  inp.value=state.pmsSearch;
  var meta=el('span','meta'); meta.style.marginLeft='8px';
  var tbody=el('div');
  function refresh(){
    var q=state.pmsSearch.toLowerCase();
    var data=pms.filter(function(r){
      return !q||(r.rm+' '+r.client+' '+r.asset+' '+r.scheme+' '+r.category).toLowerCase().indexOf(q)>=0;
    });
    meta.textContent=data.length+' rows';
    tbody.innerHTML='';
    var tw=el('div','table-wrap tall');
    tw.appendChild(buildTable([
      {label:'#',get:function(r,i){return i+1;},align:'right',isRank:true},
      {label:'RM',get:function(r){return r.rm;}},
      {label:'Client',get:function(r){return r.client;}},
      {label:'Asset',get:function(r){return r.asset;}},
      {label:'Scheme',get:function(r){return r.scheme;}},
      {label:'Category',get:function(r){return r.category;}},
      {label:'Month',get:function(r){return r.month;}},
      {label:'Investment',get:function(r){return r.punchAmount;},fmt:inr,align:'right'},
      {label:'Revenue',get:function(r){return r.revenue;},fmt:inr,align:'right'},
      {label:'Status',get:function(r){return r.confirmed;}}
    ],data,{}));
    tbody.appendChild(tw);
  }
  inp.oninput=function(){state.pmsSearch=this.value;refresh();};
  frow.appendChild(inp); frow.appendChild(meta);
  s.appendChild(frow); s.appendChild(tbody); pane.appendChild(s);
  refresh();
}

/* ============================================================
   TALK TAB — all 5 metric groups in chart + table form
   ============================================================ */
var TALK_VIEWS=[
  {id:'team',    label:'👥 Team Summary'},
  {id:'rm',      label:'🏅 RM Rankings'},
  {id:'monthly', label:'📅 Monthly Trend'},
  {id:'talktime',label:'⏱️ Talktime'}
];
var TALK_METRICS=[
  {id:'netSale',label:'Net Sale',money:true},
  {id:'sip',label:'SIP Book',money:true},
  {id:'pms',label:'PMS',money:true},
  {id:'clients',label:'Clients Added',money:false},
  {id:'talktime',label:'Talktime',money:false}
];

function talkVal(r,metric){
  if(state.talkMonth==='YTD')return sum(r[metric]);
  var idx=DATA.talk.months.indexOf(state.talkMonth);
  return idx>=0?r[metric][idx]:0;
}

function renderTalk(){
  var pane=$('tab-talk'); pane.innerHTML='';
  var rows=getTalkRows();
  function sv(metric){return sum(rows,function(r){return talkVal(r,metric);});}
  var netSale=sv('netSale'),sipV=sv('sip'),cliV=sv('clients'),pmsV=sv('pms'),tt=sv('talktime');

  pane.appendChild(kpiRow([
    [inrCr(netSale),'Net Sale','primary'],
    [inrCr(sipV),'SIP Book','blue'],
    [inrCr(pmsV),'PMS','violet'],
    [numFmt(cliV),'Clients Added','green'],
    [hours(tt),'Total Talktime','teal'],
    [numFmt(rows.length),'Active RMs','orange']
  ]));
  pane.appendChild(makeViewBar(TALK_VIEWS,state.talkView,function(v){state.talkView=v;renderTalk();}));

  if(state.talkView==='team')    talkTeamView(pane,rows);
  else if(state.talkView==='rm') talkRmView(pane,rows);
  else if(state.talkView==='monthly') talkMonthlyView(pane,rows);
  else talkTalktimeView(pane,rows);
}

/* ── Team Summary — 5 metric charts + full table ── */
function talkTeamView(pane,rows){
  var tmap={};
  rows.forEach(function(r){
    var t=r.team||'Unassigned';
    if(!tmap[t])tmap[t]={team:t,n:0,netSale:0,sip:0,pms:0,clients:0,talktime:0};
    var o=tmap[t]; o.n++;
    ['netSale','sip','pms','clients','talktime'].forEach(function(k){o[k]+=talkVal(r,k);});
  });
  var teamRows=Object.keys(tmap).map(function(k){return tmap[k];}).sort(function(a,b){return b.netSale-a.netSale;});
  var teams=teamRows.map(function(r){return r.team;});

  /* 2-col grid of metric charts */
  var g1=el('div','grid cols-2');

  function teamMetricChart(id,metric,label,color,money,heightPx){
    var s=sect(label+' by Team','bar-'+color,'talk-team-'+metric);
    s.appendChild(cw(id,heightPx||240)); return s;
  }

  var c1=teamMetricChart('talk-team-netsale','netSale','Net Sale','blue',true);
  var c2=teamMetricChart('talk-team-sip','sip','SIP Book','green',true);
  var c3=teamMetricChart('talk-team-pms','pms','PMS','violet',true);
  var c4=teamMetricChart('talk-team-clients','clients','Clients Added','orange',false);
  g1.appendChild(c1); g1.appendChild(c2); g1.appendChild(c3); g1.appendChild(c4);
  pane.appendChild(g1);

  drawBar('talk-team-netsale',teams,[{label:'Net Sale',data:teamRows.map(function(r){return r.netSale;}),backgroundColor:ACCENTS[0]}],{money:true});
  drawBar('talk-team-sip',teams,[{label:'SIP Book',data:teamRows.map(function(r){return r.sip;}),backgroundColor:ACCENTS[2]}],{money:true});
  drawBar('talk-team-pms',teams,[{label:'PMS',data:teamRows.map(function(r){return r.pms;}),backgroundColor:ACCENTS[4]}],{money:true});
  drawBar('talk-team-clients',teams,[{label:'Clients',data:teamRows.map(function(r){return r.clients;}),backgroundColor:ACCENTS[5]}],{money:false});

  /* Talktime separate (full width) */
  var st=sect('Talktime by Team','bar-teal','talk-team-talktime');
  st.appendChild(cw('talk-team-tt',220)); pane.appendChild(st);
  drawBar('talk-team-tt',teams,[{label:'Talktime (hrs)',data:teamRows.map(function(r){return r.talktime;}),backgroundColor:ACCENTS[7]}],{money:false,
    datalabels:true});
  /* override formatter for hours */
  if(charts['talk-team-tt']){
    charts['talk-team-tt'].options.plugins.datalabels.formatter=function(v){return hours(v);};
    charts['talk-team-tt'].update();
  }

  /* comprehensive team table */
  var s2=sect('Team Summary — All Metrics','bar-green','talk-team-table');
  var tw=el('div','table-wrap');
  tw.appendChild(buildTable([
    {label:'Team',get:function(r){return r.team;}},
    {label:'RMs',get:function(r){return r.n;},fmt:numFmt,align:'right',sum:true},
    {label:'Net Sale',get:function(r){return r.netSale;},fmt:inr,align:'right',sum:true},
    {label:'SIP Book',get:function(r){return r.sip;},fmt:inr,align:'right',sum:true},
    {label:'PMS',get:function(r){return r.pms;},fmt:inr,align:'right',sum:true},
    {label:'Clients',get:function(r){return r.clients;},fmt:numFmt,align:'right',sum:true},
    {label:'Talktime',get:function(r){return r.talktime;},fmt:hours,align:'right'}
  ],teamRows,{grand:true,heat:2}));
  s2.appendChild(tw); pane.appendChild(s2);
}

/* ── RM Rankings — metric picker + top chart + full all-metrics table ── */
function talkRmView(pane,rows){
  var mbar=el('div','metric-bar');
  TALK_METRICS.forEach(function(m){
    var btn=el('button','mpill'+(m.id===state.talkMetric?' active':''),m.label);
    btn.onclick=function(){state.talkMetric=m.id;renderTalk();};
    mbar.appendChild(btn);
  });

  var mDef=TALK_METRICS.filter(function(m){return m.id===state.talkMetric;})[0]||TALK_METRICS[0];
  var sorted=rows.slice().sort(function(a,b){return talkVal(b,state.talkMetric)-talkVal(a,state.talkMetric);});
  var top=sorted.slice(0,15);

  var s1=sect('Top 15 RMs — '+mDef.label,'bar-orange','talk-rm-chart');
  s1.insertBefore(mbar,s1.querySelector('h2').nextSibling);
  s1.appendChild(cw('talk-rm-chart',420)); pane.appendChild(s1);

  var fmtFn=mDef.money?inrCr:function(v){return mDef.id==='talktime'?hours(v):numFmt(v);};
  drawHBar('talk-rm-chart',top.map(function(r){return r.name;}),[
    {label:mDef.label,data:top.map(function(r){return talkVal(r,state.talkMetric);}),backgroundColor:ACCENTS[0]}
  ],{money:mDef.money});

  /* full RM table — all metrics side by side */
  var s2=sect('RM Rankings — All Metrics','bar-blue','talk-rm-table');
  var tw=el('div','table-wrap tall');
  tw.appendChild(buildTable([
    {label:'#',   get:function(r,i){return i+1;},align:'right',isRank:true},
    {label:'RM',  get:function(r){return r.name;}},
    {label:'Team',get:function(r){return r.team;}},
    {label:'Level',get:function(r){return r.level;}},
    {label:'Role',get:function(r){return r.role;}},
    {label:'Net Sale',get:function(r){return talkVal(r,'netSale');},fmt:inr,align:'right',sum:true},
    {label:'SIP Book',get:function(r){return talkVal(r,'sip');},fmt:inr,align:'right',sum:true},
    {label:'PMS',get:function(r){return talkVal(r,'pms');},fmt:inr,align:'right',sum:true},
    {label:'Clients',get:function(r){return talkVal(r,'clients');},fmt:numFmt,align:'right',sum:true},
    {label:'Talktime',get:function(r){return talkVal(r,'talktime');},fmt:hours,align:'right'}
  ],sorted,{grand:true,heat:5}));
  s2.appendChild(tw); pane.appendChild(s2);
}

/* ── Monthly Trend — all 5 metrics by month ── */
function talkMonthlyView(pane,rows){
  var months=DATA.talk.months;
  var period=state.talkMonth==='YTD'?'Full Year':state.talkMonth;
  if(state.talkMonth!=='YTD'){
    /* single month — show per-RM breakdown */
    talkRmView(pane,rows); return;
  }

  function teamColor(team){
    var teams=allTeams(); var idx=teams.indexOf(team);
    return ACCENTS[idx%ACCENTS.length];
  }

  function monthDatasets(metric,money){
    /* aggregate by month across all rows */
    return [{
      label:metric,
      data:months.map(function(mo,mi){
        return sum(rows,function(r){return r[metric][mi]||0;});
      }),
      backgroundColor:ACCENTS[TALK_METRICS.findIndex(function(m){return m.id===metric;})%ACCENTS.length]
    }];
  }

  var metrics=[
    {id:'netSale',label:'Net Sale',money:true,color:'blue'},
    {id:'sip',label:'SIP Book',money:true,color:'green'},
    {id:'pms',label:'PMS',money:true,color:'violet'},
    {id:'clients',label:'Clients Added',money:false,color:'orange'},
    {id:'talktime',label:'Talktime',money:false,color:'teal'}
  ];

  metrics.forEach(function(m,mi){
    var s=sect('Monthly '+m.label,'bar-'+m.color,'talk-monthly-'+m.id);
    var cid='talk-monthly-'+m.id;
    s.appendChild(cw(cid,220)); pane.appendChild(s);
    drawBar(cid,months.map(mtdLabel),monthDatasets(m.id,m.money),{money:m.money,legend:false,datalabels:false});
    if(m.id==='talktime'&&charts[cid]){
      charts[cid].options.scales.y.ticks.callback=function(v){return hours(v);};
      charts[cid].options.plugins.tooltip.callbacks.label=function(c){return hours(c.parsed.y);};
      charts[cid].update();
    }
  });

  /* comprehensive monthly table */
  var s=sect('Monthly Summary — All Metrics','bar-green','talk-monthly-table');
  var tw=el('div','table-wrap tall');
  tw.appendChild(buildTable([
    {label:'Month',get:function(m){return mtdLabel(m);}},
    {label:'Net Sale',get:function(m,mi){return sum(rows,function(r){return r.netSale[mi]||0;});},fmt:inr,align:'right',sum:true},
    {label:'SIP Book',get:function(m,mi){return sum(rows,function(r){return r.sip[mi]||0;});},fmt:inr,align:'right',sum:true},
    {label:'PMS',get:function(m,mi){return sum(rows,function(r){return r.pms[mi]||0;});},fmt:inr,align:'right',sum:true},
    {label:'Clients',get:function(m,mi){return sum(rows,function(r){return r.clients[mi]||0;});},fmt:numFmt,align:'right',sum:true},
    {label:'Talktime',get:function(m,mi){return sum(rows,function(r){return r.talktime[mi]||0;});},fmt:hours,align:'right'}
  ],months,{grand:true}));
  s.appendChild(tw); pane.appendChild(s);
}

/* ── Talktime view ── */
function talkTalktimeView(pane,rows){
  var sorted=rows.slice().sort(function(a,b){return talkVal(b,'talktime')-talkVal(a,'talktime');});
  var top=sorted.slice(0,15);

  var s1=sect('Top 15 RMs by Talktime','bar-teal','talk-tt-chart');
  s1.appendChild(cw('talk-tt-chart',420)); pane.appendChild(s1);
  drawHBar('talk-tt-chart',top.map(function(r){return r.name;}),[
    {label:'Talktime (hrs)',data:top.map(function(r){return talkVal(r,'talktime');}),backgroundColor:ACCENTS[7]}
  ],{money:false});
  if(charts['talk-tt-chart']){
    charts['talk-tt-chart'].options.plugins.datalabels.formatter=function(v){return hours(v);};
    charts['talk-tt-chart'].update();
  }

  var s2=sect('Talktime Breakdown — All RMs','bar-cyan','talk-tt-table');
  var tw=el('div','table-wrap tall');
  tw.appendChild(buildTable([
    {label:'#',   get:function(r,i){return i+1;},align:'right',isRank:true},
    {label:'RM',  get:function(r){return r.name;}},
    {label:'Team',get:function(r){return r.team;}},
    {label:'Level',get:function(r){return r.level;}},
    {label:'Role',get:function(r){return r.role;}},
    {label:'Talktime',get:function(r){return talkVal(r,'talktime');},fmt:hours,align:'right'},
    {label:'Net Sale',get:function(r){return talkVal(r,'netSale');},fmt:inr,align:'right',sum:true},
    {label:'SIP Book',get:function(r){return talkVal(r,'sip');},fmt:inr,align:'right',sum:true},
    {label:'Clients',get:function(r){return talkVal(r,'clients');},fmt:numFmt,align:'right',sum:true}
  ],sorted,{grand:false,heat:5}));
  s2.appendChild(tw); pane.appendChild(s2);
}

/* ============================================================
   IN-BROWSER .xlsb / .xlsx PARSER (mirrors extract.py)
   ============================================================ */
function ser2month(n){
  var d=new Date(Date.UTC(1899,11,30)+n*86400000);
  var mn=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return mn[d.getUTCMonth()]+'-'+d.getUTCFullYear();
}
var FY_MONTHS=FY_SERIALS.map(ser2month);

/* Derive the 12 FY months from the MTD month string (e.g. "JUN-2026" or "Jun-2026").
   Returns Apr-YYYY … Mar-YYYY+1. Falls back to hardcoded FY_MONTHS on parse failure. */
function fyFromMtdMonth(mtd){
  var mn=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  var MN=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  var s=(mtd||'').trim().toLowerCase();
  var mo=-1;
  for(var j=0;j<12;j++){if(s.indexOf(mn[j])===0){mo=j;break;}}
  var m4=s.match(/(\d{4})/); var yr=m4?parseInt(m4[1]):NaN;
  if(mo<0||isNaN(yr))return FY_MONTHS;
  var fyStart=mo>=3?yr:yr-1;          // Apr+ → same year; Jan-Mar → previous year
  var out=[];
  for(var i=0;i<12;i++){var mm=(3+i)%12; out.push(MN[mm]+'-'+(fyStart+(mm<3?1:0)));}
  return out;
}

function cell(row,i){return (row&&i<row.length)?row[i]:null;}
function sstr(v){ if(v==null)return ''; if(typeof v==='number'&&v%1===0)return ''+v; return (''+v).trim(); }
function serToIsoDate(v){
  if(v==null||v==='')return '';
  function p2(n){return (n<10?'0':'')+n;}
  if(typeof v==='number'){
    var d=new Date(Date.UTC(1899,11,30)+Math.round(v)*86400000);
    return d.getUTCFullYear()+'-'+p2(d.getUTCMonth()+1)+'-'+p2(d.getUTCDate());
  }
  var dt=new Date((''+v).trim());
  if(!isNaN(dt.getTime()))return dt.getFullYear()+'-'+p2(dt.getMonth()+1)+'-'+p2(dt.getDate());
  return (''+v).trim();
}
function titleCase(s){return s?s.charAt(0).toUpperCase()+s.slice(1).toLowerCase():'';}
function talktimeHours(v){
  if(v==null)return 0;
  if(typeof v==='number')return Math.round(v*24*10000)/10000;
  var t=(''+v).trim(); if(!t)return 0;
  var p=t.split(':').map(parseFloat); while(p.length<3)p.push(0);
  return Math.round((p[0]+p[1]/60+p[2]/3600)*10000)/10000;
}
function parseWorkbook(wb){
  function sheet(name){var s=wb.Sheets[name];return s?XLSX.utils.sheet_to_json(s,{header:1,raw:true,defval:null}):[]; }
  var ins=[],fees=[],pmsArr=[],talk=[];

  /* ── Read REVSUM first so we can detect the fiscal year ── */
  var R=sheet('REVENUE SUMMARY V 2.0');
  var rsMonth=R.length?sstr(cell(R[0],3)):'';
  var fyM=rsMonth?fyFromMtdMonth(rsMonth):FY_MONTHS;   // dynamic FY months array

  /* INS — skip rows outside current FY */
  var I=sheet('INS');
  for(var i=1;i<I.length;i++){var r=I[i];if(!r||cell(r,2)==null||cell(r,2)==='')continue;
    if((sstr(cell(r,1))||'').toUpperCase()==='B2B')continue;
    var mo=sstr(cell(r,48)); if(mo&&fyM.indexOf(mo)<0)continue;
    ins.push({dtype:sstr(cell(r,1))||'B2C',rm:sstr(cell(r,2)),client:sstr(cell(r,3)),
      platform:sstr(cell(r,8))||'Unknown',category:sstr(cell(r,10))||'Unknown',insType:sstr(cell(r,11))||'Unknown',
      partner:sstr(cell(r,14))||'Unknown',premium:num(cell(r,27)),createDate:serToIsoDate(cell(r,47)),month:mo,confirmed:sstr(cell(r,50)),
      type:sstr(cell(r,51))||'Unknown',revenue:num(cell(r,52)),pct:num(cell(r,54)),
      teamLeader:sstr(cell(r,56)),status:sstr(cell(r,63))||'Unknown'});}

  /* FEES — skip rows outside current FY */
  var F=sheet('FEES');
  for(i=1;i<F.length;i++){var fr=F[i];if(!fr||cell(fr,2)==null||cell(fr,2)==='')continue;
    if((sstr(cell(fr,1))||'').toUpperCase()==='B2B')continue;
    var fmo=sstr(cell(fr,29)); fmo=fmo?fmo.charAt(0).toUpperCase()+fmo.slice(1).toLowerCase():'';
    if(fmo&&fyM.indexOf(fmo)<0)continue;
    fees.push({dtype:sstr(cell(fr,1))||'B2C',rm:sstr(cell(fr,2)),client:sstr(cell(fr,3)),
      platform:sstr(cell(fr,8))||'Unknown',category:sstr(cell(fr,10))||'Unknown',
      asset:sstr(cell(fr,11))||'Unknown',prodType:sstr(cell(fr,12))||'Unknown',
      amount:num(cell(fr,14)),createDate:serToIsoDate(cell(fr,28)),month:fmo,transType:sstr(cell(fr,31))||'Unknown',
      confirmed:sstr(cell(fr,32)),revenue:num(cell(fr,33))});}

  /* PMS — skip rows outside current FY */
  var P=sheet('PMS');
  for(i=1;i<P.length;i++){var pr=P[i];if(!pr||cell(pr,2)==null||cell(pr,2)==='')continue;
    if((sstr(cell(pr,1))||'').toUpperCase()==='B2B')continue;
    var pmo=sstr(cell(pr,28)); pmo=pmo?titleCase(pmo):'';
    if(pmo&&fyM.indexOf(pmo)<0)continue;
    pmsArr.push({dtype:sstr(cell(pr,1))||'B2C',rm:sstr(cell(pr,2)),client:sstr(cell(pr,3)),
      platform:sstr(cell(pr,8))||'Unknown',category:sstr(cell(pr,10))||'Unknown',
      asset:sstr(cell(pr,11))||'Unknown',scheme:sstr(cell(pr,12))||'Unknown',
      punchAmount:num(cell(pr,16)),createDate:serToIsoDate(cell(pr,27)),month:pmo,
      confirmed:sstr(cell(pr,32))||'Unknown',revenue:num(cell(pr,33))});}

  /* Talk — month-indexed arrays, not filtered by date */
  var T=sheet('Talk'); var groups={talktime:6,netSale:21,sip:35,pms:50,clients:64};
  for(i=2;i<T.length;i++){var tr=T[i];if(!tr)continue;
    if((cell(tr,0)==null||cell(tr,0)==='')&&(cell(tr,2)==null||cell(tr,2)===''))continue;
    var rec={empCode:sstr(cell(tr,0)),team:sstr(cell(tr,1))||'Unassigned',name:sstr(cell(tr,2)),level:sstr(cell(tr,4)),role:sstr(cell(tr,5))};
    Object.keys(groups).forEach(function(key){var st=groups[key];var a=[];for(var m=0;m<12;m++){var cv=cell(tr,st+m);a.push(key==='talktime'?talktimeHours(cv):num(cv));}rec[key]=a;});
    talk.push(rec);}

  /* REVSUM rows (header already read) */
  var revsum=[]; var revsumOverall={mtdTarget:0,mtdAch:0,ytdTarget:0,ytdAch:0};
  for(i=3;i<R.length;i++){var rr=R[i];if(!rr)continue;
    var nm=sstr(cell(rr,3)); if(!nm||nm.toLowerCase().indexOf('pool')===0)continue;
    if(nm.toLowerCase().indexOf('overall')===0){
      revsumOverall={mtdTarget:num(cell(rr,7)),mtdAch:num(cell(rr,8)),ytdTarget:num(cell(rr,22)),ytdAch:num(cell(rr,23))};
      continue;}
    if(typeof cell(rr,0)!=='number')continue;
    revsum.push({sn:num(cell(rr,0)),empCode:sstr(cell(rr,1)),team:sstr(cell(rr,2))||'Unassigned',name:nm,level:sstr(cell(rr,5)),kra:sstr(cell(rr,6)),
      mtd:{revTarget:num(cell(rr,7)),revAch:num(cell(rr,8)),pct:num(cell(rr,9)),cliTarget:num(cell(rr,10)),totalClients:num(cell(rr,11)),cliAcq:num(cell(rr,12)),
        aumTarget:num(cell(rr,13)),totalAum:num(cell(rr,14)),mfNetSale:num(cell(rr,15)),pmsAif:num(cell(rr,16)),sipTarget:num(cell(rr,17)),sipAdd:num(cell(rr,18)),gtTarget:num(cell(rr,19)),gtTime:talktimeHours(cell(rr,20))},
      ytd:{revTarget:num(cell(rr,22)),revAch:num(cell(rr,23)),revDeff:num(cell(rr,24)),cliTarget:num(cell(rr,25)),totalClients:num(cell(rr,26)),cliAcq:num(cell(rr,27)),
        aumTarget:num(cell(rr,28)),totalAum:num(cell(rr,29)),mfNetSale:num(cell(rr,30)),pmsAif:num(cell(rr,31)),sipTarget:num(cell(rr,32)),sipAdd:num(cell(rr,33)),gtTarget:num(cell(rr,34)),gtTime:talktimeHours(cell(rr,35))}});}

  /* REVENUE REPORT MONTHLY / YTD — reconciled product-wise (different layouts) */
  function readReport(name,map){
    var S=sheet(name); var out=[]; var hdrMonth=S.length&&S[0]?sstr(cell(S[0],3)):'';
    for(var r=2;r<S.length;r++){var row=S[r];if(!row)continue;
      if(cell(row,0)==null||cell(row,0)==='')continue;
      var rec={empCode:sstr(cell(row,0)),team:sstr(cell(row,1))||'Unassigned',name:sstr(cell(row,3))||sstr(cell(row,2)),
        pool:sstr(cell(row,2)).toLowerCase().indexOf('pool')===0};
      Object.keys(map).forEach(function(k){var s=0;map[k].forEach(function(c){s+=num(cell(row,c));});rec[k]=Math.round(s*100)/100;});
      out.push(rec);}
    return {mtdMonth:hdrMonth,rows:out};
  }
  var rrMonthly=readReport('REVENUE REPORT MONTHLY',{fees:[4,5],insurance:[13],pms:[15,16,17],trail:[18],total:[19]});
  var rrYtd=readReport('REVENUE REPORT YTD',{fees:[4,5],insurance:[6,7,8,9,10,11,12],pms:[13,14,15],trail:[16],total:[17]});

  return {
    meta:{source:'(live)',fyMonths:fyM,generated:new Date().toLocaleString(),
      insRows:ins.length,feesRows:fees.length,pmsRows:pmsArr.length,talkRows:talk.length,revsumRows:revsum.length,revsumMtdMonth:rsMonth},
    ins:ins,fees:fees,pms:pmsArr,
    talk:{months:fyM,rows:talk},
    revsum:{mtdMonth:rsMonth,rows:revsum,overallTarget:revsumOverall},
    revreport:{mtdMonth:rrMonthly.mtdMonth,monthly:rrMonthly.rows,ytd:rrYtd.rows}
  };
}

/* ============================================================
   FILTERS & BOOT
   ============================================================ */
function populateMonthFilter(){
  var allMonths=[];
  (DATA.ins||[]).forEach(function(r){allMonths.push(r.month);});
  (DATA.fees||[]).forEach(function(r){allMonths.push(r.month);});
  (DATA.pms||[]).forEach(function(r){allMonths.push(r.month);});
  var months=uniqInFyOrder(allMonths);

  var sel=$('filter-month'); sel.innerHTML='';
  var ao=el('option',null,'All Months'); ao.value='All'; sel.appendChild(ao);
  months.forEach(function(m){var o=el('option',null,mtdLabel(m));o.value=m;sel.appendChild(o);});
  sel.value=state.month;

  /* team filter */
  var tsel=$('filter-chart-team'); tsel.innerHTML='';
  var to=el('option',null,'All Teams'); to.value='All'; tsel.appendChild(to);
  allTeams().forEach(function(t){var o=el('option',null,t);o.value=t;tsel.appendChild(o);});
  tsel.value=state.chartTeam;

  var ttsel=$('filter-talkmonth'); ttsel.innerHTML='';
  var yo=el('option',null,'YTD (Full Year)'); yo.value='YTD'; ttsel.appendChild(yo);
  DATA.talk.months.forEach(function(m){var o=el('option',null,mtdLabel(m));o.value=m;ttsel.appendChild(o);});
  ttsel.value=state.talkMonth;

  var rsel=$('filter-revsum-team'); if(rsel){rsel.innerHTML='';
    var ro=el('option',null,'All Teams'); ro.value='All'; rsel.appendChild(ro);
    var teams={}; (DATA.revsum.rows||[]).forEach(function(r){teams[r.team]=1;});
    Object.keys(teams).sort().forEach(function(t){var o=el('option',null,t);o.value=t;rsel.appendChild(o);});
    rsel.value=state.revsumTeam;}
  var vsel=$('filter-revsum-view'); if(vsel)vsel.value=state.revsumView;
}

function bindFilters(){
  $('filter-month').onchange=function(){state.month=this.value;renderActive();};
  $('filter-chart-team').onchange=function(){state.chartTeam=this.value;renderActive();};
  $('filter-talkmonth').onchange=function(){state.talkMonth=this.value;renderActive();};
  $('filter-revsum-view').onchange=function(){state.revsumView=this.value;renderActive();};
  $('filter-revsum-team').onchange=function(){state.revsumTeam=this.value;renderActive();};
  $('reupload-file').onchange=handleUpload;
}

function handleUpload(e){
  var f=e.target.files[0]; if(!f)return;
  var rd=new FileReader();
  rd.onload=function(ev){
    try{
      var wb=XLSX.read(new Uint8Array(ev.target.result),{
        type:'array',
        sheets:['INS','FEES','PMS','Talk','REVENUE SUMMARY V 2.0','REVENUE REPORT MONTHLY','REVENUE REPORT YTD'],
        cellFormula:false,cellHTML:false,cellDates:false
      });
      DATA=parseWorkbook(wb);
      buildRmTeam(); populateMonthFilter(); state.month='All'; state.chartTeam='All';
      renderActive();
    }catch(err){alert('Could not parse file: '+err.message);}
  };
  rd.readAsArrayBuffer(f);
}

/* ── GitHub fetch: data.js only (fast, ~2 MB) ── */
async function loadFromGitHub(setStatus){
  var H=GH.token?{'Authorization':'token '+GH.token}:{};
  var rawBase='https://raw.githubusercontent.com/'+GH.owner+'/'+GH.repo;

  setStatus('Loading data from GitHub…');
  /* Try main branch first, then master.
     cache-bust query + no-store so GitHub's CDN / browser never serve a stale
     copy after a fresh data.js upload. */
  var bust='?t='+Date.now();
  var branches=['main','master'];
  for(var i=0;i<branches.length;i++){
    var url=rawBase+'/'+branches[i]+'/data.js'+bust;
    var res=await fetch(url,{headers:H,cache:'no-store'});
    if(res.ok){
      var txt=await res.text();
      var json=txt.replace(/^\s*window\.REVENUE_DATA\s*=\s*/,'').replace(/;\s*$/,'');
      var d=JSON.parse(json);
      d.meta=d.meta||{}; d.meta.source=d.meta.source||'GitHub data.js';
      return d;
    }
  }
  throw new Error('data.js not found in GitHub repo. Run extract.py → upload data.js to the repo.');
}

async function boot(){
  if(typeof Chart!=='undefined'&&typeof ChartDataLabels!=='undefined'){
    Chart.register(ChartDataLabels);
  }
  var ld=$('loading'), lmsg=$('loading-msg');
  function setStatus(t){if(lmsg)lmsg.innerHTML=t;}

  /* When testing locally (localhost / file://), prefer the freshly-generated
     local data.js so you can preview new data before pushing to GitHub. */
  var isLocal=/^(localhost|127\.0\.0\.1|\[::1\])$/.test(location.hostname)||location.protocol==='file:';
  if(isLocal&&window.REVENUE_DATA){
    setStatus('Loading local data…');
    DATA=window.REVENUE_DATA;
    buildRmTeam();
    if(ld)ld.style.display='none';
    $('app').style.display='block';
    renderTabs(); populateMonthFilter(); bindFilters(); setTab('dashboard');
    return;
  }

  /* Try GitHub first */
  if(GH.owner&&GH.repo){
    try{
      DATA=await loadFromGitHub(setStatus);
    }catch(err){
      console.warn('GitHub load failed:',err.message);
      if(window.REVENUE_DATA){
        setStatus('GitHub unavailable (<i>'+err.message+'</i>) — using local data…');
        DATA=window.REVENUE_DATA;
        await new Promise(function(r){setTimeout(r,900);});
      }else{
        setStatus('<span style="color:#ef4444;font-weight:600">'+err.message+'</span><br>'
          +'<small style="color:var(--muted)">Push the .xlsb to GitHub or run <code>python extract.py</code> locally.</small>');
        return;
      }
    }
  }else if(window.REVENUE_DATA){
    DATA=window.REVENUE_DATA;
  }else{
    setStatus('<span style="color:#ef4444">No data source. Set GH.owner/GH.repo in app.js or run extract.py.</span>');
    return;
  }

  buildRmTeam();
  if(ld)ld.style.display='none';
  $('app').style.display='block';
  renderTabs();
  populateMonthFilter();
  bindFilters();
  setTab('dashboard');
}
if(document.readyState!=='loading')boot(); else document.addEventListener('DOMContentLoaded',boot);

/* ============================================================
   INFO POPUPS — describe each chart/table, its data source sheet
   and the source columns used
   ============================================================ */
var TABLE_INFO={
  /* DASHBOARD */
  'dash-trend-total':{title:'Total Revenue Trend by Month',
    desc:'Stacked monthly bar chart combining Insurance + FEES + PMS revenue. Each segment shows one product line; the total bar height is the combined revenue for that month. Respects Month and Team filters.',
    cols:'INS revenue (col 52), FEES revenue (col 33), PMS revenue (col 33) — aggregated by month.',
    source:'INS, FEES, PMS sheets'},
  'dash-trend-ins':{title:'Insurance Revenue Trend by Month',
    desc:'Monthly bar chart of Insurance revenue. Respects the Month and Team filters.',
    cols:'INS sheet — col 52 REVENUE · col 48 MONTH · col 2 RmName. Rows where col 1 (B Type) ≠ "B2C" are excluded.',
    source:'INS sheet'},
  'dash-trend-fees':{title:'FEES Revenue Trend by Month',
    desc:'Monthly bar chart of FEES (Financial Planning / Fin-Plan) revenue. Respects the Month and Team filters.',
    cols:'FEES sheet — col 33 REVENUE · col 29 FEES CAME IN (month) · col 2 RmName. Rows where col 1 (B Type) ≠ "B2C" are excluded.',
    source:'FEES sheet'},
  'dash-trend-pms':{title:'PMS Revenue Trend by Month',
    desc:'Monthly bar chart of PMS (Portfolio Management Services) revenue. Respects the Month and Team filters.',
    cols:'PMS sheet — col 33 REVENUE · col 28 MONTH · col 2 RM. Rows where col 1 (B Type) ≠ "B2C" are excluded.',
    source:'PMS sheet'},
  'dash-team':{title:'Revenue by Team (INS + FEES + PMS)',
    desc:'Per-team aggregation of INS, FEES, PMS revenue and total. Team is resolved from RM name via the Talk sheet (RM → Team mapping).',
    cols:'INS col 52 REVENUE · FEES col 33 REVENUE · PMS col 33 REVENUE · RM names from col 2 of each sheet · Team from Talk sheet col 1 (Team) joined on Talk col 2 (Name).',
    source:'INS, FEES, PMS, Talk sheets'},
  'dash-month':{title:'Revenue by Month',
    desc:'Per-month totals from each product line. The current calendar month appears with an " — MTD" suffix.',
    cols:'INS col 52 REVENUE · col 48 MONTH | FEES col 33 REVENUE · col 29 FEES CAME IN | PMS col 33 REVENUE · col 28 MONTH. Month strings are title-cased to match FY format Apr-2026 … Mar-2027.',
    source:'INS, FEES, PMS sheets'},

  /* REVENUE SUMMARY V 2.0 */
  'revsum-chart':{title:'Top RMs — Revenue Achieved vs Target',
    desc:'Top 15 RMs ranked by REVENUE ACHIEVEMENT in the selected view (MTD or YTD). Bar colour encodes % Achievement: green ≥80%, yellow 50–79%, red <50%. Grey bars are the corresponding REVENUE TARGET.',
    cols:'MTD view → col 7 (MTD REVENUE TARGET), col 8 (MTD REVENUE ACHIEVEMENT) | YTD view → col 22 (YTD REVENUE TARGET), col 23 (YTD REVENUE ACHIEVEMENT). RM name from col 3.',
    source:'REVENUE SUMMARY V 2.0 sheet'},
  'revsum-team':{title:'Team Summary — Target vs Achievement',
    desc:'Per-team aggregation of revenue target, achievement, % achievement, AUM, MF Net Sale, Net SIP, and total clients. Row colour encodes %Ach: green ≥80%, yellow 50–79%, red <50%.',
    cols:'Team from col 2 · RMs counted from col 3 · MTD: cols 7 (Rev Target), 8 (Rev Ach), 14 (Total AUM), 15 (MF Net Sale), 18 (New SIP Add), 11 (Total Clients) | YTD: cols 22 (Rev Target), 23 (Rev Ach), 29 (Total AUM), 30 (MF Net Sale), 33 (Net SIP Add), 26 (Total Client). "Pool" rows are excluded.',
    source:'REVENUE SUMMARY V 2.0 sheet'},
  'revsum-detail':{title:'RM Detail — KRA Scorecard',
    desc:'Full per-RM KRA scorecard with revenue, clients, AUM, MF Net Sale, PMS/AIF, SIP, and Greentape Time. Row colour by %Ach.',
    cols:'Cols 0 (SN), 1 (Emp Code), 2 (Team), 3 (Name), 5 (RM Level), 6 (KRA) · MTD block cols 7–20 · YTD block cols 22–35. GT Time fields are h:m:s in MTD and Excel-day-fraction in YTD — both converted to hours.',
    source:'REVENUE SUMMARY V 2.0 sheet'},

  /* INS */
  'ins-rm-chart':{title:'Top 15 RMs by Insurance Revenue',
    desc:'Horizontal bar chart of top 15 RMs ranked by INS revenue. Shows Revenue and Premium side-by-side.',
    cols:'col 2 RmName · col 27 Premium · col 52 REVENUE (post-confirmation revenue). B2C only (col 1).',
    source:'INS sheet'},
  'ins-team-chart':{title:'Insurance Revenue by Team',
    desc:'Total INS revenue aggregated per team. Team is resolved from each RM via the Talk sheet.',
    cols:'col 2 RmName → Team (Talk sheet) · col 52 REVENUE.',
    source:'INS sheet + Talk sheet'},
  'ins-rm-table':{title:'INS RM Rankings — Full List',
    desc:'Every RM with policy count, premium, revenue, %-of-total and average revenue per policy.',
    cols:'col 2 RmName · col 27 Premium · col 52 REVENUE · row count = policy count. Team from Talk sheet.',
    source:'INS, Talk sheets'},
  'ins-product-mix':{title:'Insurance Product Mix',
    desc:'Shows how many types of insurance each team (or RM, when a team is selected) has sold, along with total premium and average premium per client. Team view groups by Team × Insurance Type; RM view groups by RM × Insurance Type.',
    cols:'col 2 RmName · col 11 InsuranceType · col 27 Premium. Team from Talk sheet.',
    source:'INS sheet + Talk sheet'},
  'ins-instype-chart':{title:'Revenue by Insurance Category (chart)',
    desc:'Vertical bar chart of revenue grouped by Insurance Type. Hover for policy count.',
    cols:'col 11 InsuranceType · col 52 REVENUE.',
    source:'INS sheet'},
  'ins-instype-table':{title:'By Insurance Category',
    desc:'Revenue + count grouped by Insurance Type with share-of-total.',
    cols:'col 11 InsuranceType · col 52 REVENUE.',
    source:'INS sheet'},
  'ins-type-table':{title:'By Product Type (Revenue Type)',
    desc:'Revenue grouped by the "TYPE" column from the post-processing block (e.g. New / Renewal / 1st Year).',
    cols:'col 51 TYPE · col 52 REVENUE.',
    source:'INS sheet'},
  'ins-partner-chart':{title:'Top 10 Insurance Partners by Revenue',
    desc:'Bar chart of top 10 PolicyPartner values ranked by revenue.',
    cols:'col 14 PolicyPartner · col 52 REVENUE.',
    source:'INS sheet'},
  'ins-partner-table':{title:'By Insurance Partner',
    desc:'All partners with revenue and share-of-total.',
    cols:'col 14 PolicyPartner · col 52 REVENUE.',
    source:'INS sheet'},
  'ins-category-table':{title:'By Source / Category',
    desc:'Revenue grouped by Category column (lead source / campaign category).',
    cols:'col 10 Category · col 52 REVENUE.',
    source:'INS sheet'},
  'ins-monthly-chart':{title:'Monthly Insurance Revenue',
    desc:'Per-month revenue + premium bars. Use the Day Window to compare the same day range (e.g. day 1–11) across every month — rows are filtered by their Create Date.',
    cols:'col 47 CREATEDATE (day filter) · col 48 MONTH · col 27 Premium · col 52 REVENUE.',
    source:'INS sheet'},
  'ins-monthly-table':{title:'Monthly Breakdown',
    desc:'Per-month policy count, premium and revenue totals for the selected Day Window (filtered by Create Date).',
    cols:'col 47 CREATEDATE (day filter) · col 48 MONTH · col 27 Premium · col 52 REVENUE.',
    source:'INS sheet'},
  'ins-status-table':{title:'Status Distribution',
    desc:'Policy count and revenue split by Status (Issued / Pending / Cancelled).',
    cols:'col 63 Status · col 52 REVENUE.',
    source:'INS sheet'},
  'ins-raw':{title:'INS — Raw Policy Data',
    desc:'Every B2C INS row with Excel-style per-column filters. Click the ▼ on any header to choose values or type a "contains" filter. Filters interact dependently — values in one column reflect what other filters have already excluded.',
    cols:'Cols 2 (RmName), 3 (Client Name), 10 (Category), 11 (InsuranceType), 14 (PolicyPartner), 27 (Premium), 48 (MONTH), 50 (CONFIRMED/NOT), 51 (TYPE), 52 (REVENUE), 63 (Status). B2C only (col 1).',
    source:'INS sheet'},

  /* FEES */
  'fees-rm-chart':{title:'Top 15 RMs by Fee Revenue',
    desc:'Horizontal bar chart of top 15 RMs ranked by Net Revenue, showing Gross Amount alongside.',
    cols:'col 2 RmName · col 14 Amount (gross) · col 33 REVENUE (net).',
    source:'FEES sheet'},
  'fees-team-chart':{title:'Fee Revenue by Team',
    desc:'Total FEES net revenue aggregated per team. Team is resolved from each RM via the Talk sheet.',
    cols:'col 2 RmName → Team (Talk sheet) · col 33 REVENUE.',
    source:'FEES sheet + Talk sheet'},
  'fees-rm-table':{title:'FEES RM Rankings — Full List',
    desc:'Every RM with txn count, gross, net revenue, %-of-total and avg/txn.',
    cols:'col 2 RmName · col 14 Amount · col 33 REVENUE. Team from Talk sheet.',
    source:'FEES, Talk sheets'},
  'fees-trans-chart':{title:'Revenue by Transaction Type',
    desc:'Revenue grouped by TRANS. TYPE (e.g. Lumpsum / SIP / Switch). Hover for txn count.',
    cols:'col 31 TRANS. TYPE · col 33 REVENUE.',
    source:'FEES sheet'},
  'fees-trans-table':{title:'By Transaction Type',
    desc:'Tabular view of revenue by transaction type with share-of-total.',
    cols:'col 31 TRANS. TYPE · col 33 REVENUE.',
    source:'FEES sheet'},
  'fees-asset-table':{title:'By Asset / Product',
    desc:'Revenue grouped by Asset + productType combined.',
    cols:'col 11 Asset · col 12 productType · col 33 REVENUE.',
    source:'FEES sheet'},
  'fees-category-table':{title:'By Source / Category',
    desc:'Revenue grouped by Category.',
    cols:'col 10 Category · col 33 REVENUE.',
    source:'FEES sheet'},
  'fees-monthly-chart':{title:'Monthly Fee Revenue Trend',
    desc:'Per-month Net Revenue + Gross bars. Current month labelled " — MTD".',
    cols:'col 29 FEES CAME IN · col 14 Amount · col 33 REVENUE.',
    source:'FEES sheet'},
  'fees-monthly-table':{title:'FEES Monthly Breakdown',
    desc:'Txn count, gross, net revenue by month.',
    cols:'col 29 FEES CAME IN · col 14 Amount · col 33 REVENUE.',
    source:'FEES sheet'},
  'fees-raw':{title:'FEES — Raw Transaction Data',
    desc:'Every B2C FEES row with global search.',
    cols:'Cols 2 (RmName), 3 (ClientName), 10 (Category), 11 (Asset), 12 (productType), 14 (Amount), 29 (FEES CAME IN), 31 (TRANS. TYPE), 32 (CONFIRMED/NOT), 33 (REVENUE).',
    source:'FEES sheet'},

  /* PMS */
  'pms-rm-chart':{title:'Top 15 RMs by PMS Revenue',
    desc:'Horizontal bars of top 15 RMs ranked by PMS revenue with Investment (Punch Amount) alongside.',
    cols:'col 2 RM · col 16 Punch Amount · col 33 REVENUE.',
    source:'PMS sheet'},
  'pms-team-chart':{title:'PMS Revenue by Team',
    desc:'Total PMS revenue aggregated per team. Team is resolved from each RM via the Talk sheet.',
    cols:'col 2 RM → Team (Talk sheet) · col 33 REVENUE.',
    source:'PMS sheet + Talk sheet'},
  'pms-rm-table':{title:'PMS RM Rankings — Full List',
    desc:'Every RM with txn count, investment, revenue, %-of-total, avg/txn.',
    cols:'col 2 RM · col 16 Punch Amount · col 33 REVENUE.',
    source:'PMS, Talk sheets'},
  'pms-asset-chart':{title:'Revenue by Asset Type',
    desc:'Bar chart of revenue grouped by Asset.',
    cols:'col 11 Asset · col 33 REVENUE.',
    source:'PMS sheet'},
  'pms-asset-donut':{title:'Asset Distribution (Revenue)',
    desc:'Doughnut chart showing each asset\'s share of total PMS revenue.',
    cols:'col 11 Asset · col 33 REVENUE.',
    source:'PMS sheet'},
  'pms-asset-table':{title:'By Asset Type',
    desc:'Tabular view of investment, revenue and share-of-total by asset.',
    cols:'col 11 Asset · col 16 Punch Amount · col 33 REVENUE.',
    source:'PMS sheet'},
  'pms-scheme-table':{title:'By Scheme',
    desc:'Revenue grouped by Scheme name.',
    cols:'col 12 Scheme · col 33 REVENUE.',
    source:'PMS sheet'},
  'pms-monthly-chart':{title:'Monthly PMS Revenue & Investment',
    desc:'Per-month revenue and investment bars. Current month labelled " — MTD".',
    cols:'col 28 MONTH · col 16 Punch Amount · col 33 REVENUE.',
    source:'PMS sheet'},
  'pms-monthly-table':{title:'PMS Monthly Breakdown',
    desc:'Txn count, investment, revenue per month.',
    cols:'col 28 MONTH · col 16 Punch Amount · col 33 REVENUE.',
    source:'PMS sheet'},
  'pms-confstatus-table':{title:'Confirmation Status',
    desc:'Split of revenue by CONFIRMED/NOT status. Unconfirmed transactions are still included in the dashboard\'s gross PMS Revenue KPI — only the "Confirmed Revenue" KPI excludes them.',
    cols:'col 32 CONFIRMED/NOT · col 33 REVENUE.',
    source:'PMS sheet',
    note:'2 unconfirmed rows in current data contribute ~₹15L to gross. Use the Confirmed Revenue KPI when validating against accounts.'},
  'pms-raw':{title:'PMS — Raw Transaction Data',
    desc:'Every B2C PMS row with global search.',
    cols:'Cols 2 (RM), 3 (ClientName), 10 (Category), 11 (Asset), 12 (Scheme), 16 (Punch Amount), 28 (MONTH), 32 (CONFIRMED/NOT), 33 (REVENUE).',
    source:'PMS sheet'},

  /* TALK */
  'talk-team-netSale':{title:'Net Sale by Team',
    desc:'Per-team total of NET SALE for the selected period (YTD or single month). Net Sale = gross investment book added.',
    cols:'Talk sheet — col 1 Team · cols 21–32 (12 monthly Net Sale columns, Apr–Mar).',
    source:'Talk sheet'},
  'talk-team-sip':{title:'SIP Book by Team',
    desc:'Per-team total of SIP for the selected period.',
    cols:'Talk sheet — col 1 Team · cols 35–46 (12 monthly SIP columns).',
    source:'Talk sheet'},
  'talk-team-pms':{title:'PMS Book by Team',
    desc:'Per-team total of PMS sale (Talk sheet) for the selected period.',
    cols:'Talk sheet — col 1 Team · cols 50–61 (12 monthly PMS columns).',
    source:'Talk sheet',
    note:'This is the Talk sheet\'s PMS sale (book), not the PMS sheet\'s commission revenue. The two measures differ — Talk PMS for YTD ≈ ₹2.96 Cr while PMS sheet revenue ≈ ₹52L.'},
  'talk-team-clients':{title:'Clients Added by Team',
    desc:'Per-team total of new clients added for the selected period.',
    cols:'Talk sheet — col 1 Team · cols 64–75 (12 monthly Total Client columns).',
    source:'Talk sheet'},
  'talk-team-talktime':{title:'Talktime by Team',
    desc:'Per-team total Talktime in hours. h:m:s strings and Excel time fractions both convert to hours.',
    cols:'Talk sheet — col 1 Team · cols 6–17 (12 monthly Talktime columns).',
    source:'Talk sheet'},
  'talk-team-table':{title:'Team Summary — All Metrics',
    desc:'Per-team totals for Net Sale, SIP, PMS, Clients, Talktime for the selected period.',
    cols:'Talk sheet — col 1 Team · col 2 Name · all 5 metric blocks (cols 6, 21, 35, 50, 64).',
    source:'Talk sheet'},
  'talk-rm-chart':{title:'Top 15 RMs — current metric',
    desc:'Top 15 RMs ranked by the metric chosen on the pill bar above (Net Sale / SIP / PMS / Clients / Talktime).',
    cols:'Talk sheet — col 2 Name · whichever metric block is active.',
    source:'Talk sheet'},
  'talk-rm-table':{title:'RM Rankings — All Metrics',
    desc:'Every RM with all 5 metrics side-by-side, plus Team / Level / Role.',
    cols:'Talk sheet — col 0 Emp Code · col 1 Team · col 2 Name · col 4 RM Level · col 5 Role · cols 6, 21, 35, 50, 64 (metric blocks).',
    source:'Talk sheet'},
  'talk-monthly-netSale':{title:'Monthly Net Sale',
    desc:'Sum of Net Sale across all RMs per month.',
    cols:'Talk sheet — cols 21–32.',source:'Talk sheet'},
  'talk-monthly-sip':{title:'Monthly SIP Book',
    desc:'Sum of SIP across all RMs per month.',
    cols:'Talk sheet — cols 35–46.',source:'Talk sheet'},
  'talk-monthly-pms':{title:'Monthly PMS',
    desc:'Sum of PMS sale across all RMs per month.',
    cols:'Talk sheet — cols 50–61.',source:'Talk sheet'},
  'talk-monthly-clients':{title:'Monthly Clients Added',
    desc:'Sum of clients added across all RMs per month.',
    cols:'Talk sheet — cols 64–75.',source:'Talk sheet'},
  'talk-monthly-talktime':{title:'Monthly Talktime',
    desc:'Sum of Talktime hours across all RMs per month.',
    cols:'Talk sheet — cols 6–17.',source:'Talk sheet'},
  'talk-monthly-table':{title:'Monthly Summary — All Metrics',
    desc:'All 5 metrics rolled up by month across every RM.',
    cols:'Talk sheet — all 5 metric blocks summed over the 12 monthly columns.',
    source:'Talk sheet'},
  'talk-tt-chart':{title:'Top 15 RMs by Talktime',
    desc:'Top 15 RMs by total Talktime hours for the selected period.',
    cols:'Talk sheet — col 2 Name · cols 6–17 (Talktime).',
    source:'Talk sheet'},
  'talk-tt-table':{title:'Talktime Breakdown — All RMs',
    desc:'Every RM with Talktime alongside Net Sale, SIP and Clients for context.',
    cols:'Talk sheet — col 0 Emp Code · col 1 Team · col 2 Name · col 4 Level · col 5 Role · cols 6 (TT), 21 (Net Sale), 35 (SIP), 64 (Clients).',
    source:'Talk sheet'}
};

function showInfoPopup(key){
  var info=TABLE_INFO[key];
  if(!info)return;
  var overlay=document.createElement('div'); overlay.className='info-overlay';
  var popup=document.createElement('div'); popup.className='info-popup';
  var noteHtml=info.note?'<div class="info-label" style="color:#b45309">⚠ Methodology Note</div><p style="color:#92400e;background:#fffbeb;border:1px solid #fde68a;border-radius:6px;padding:8px 10px;margin-top:4px;font-size:12px">'+info.note+'</p>':'';
  popup.innerHTML='<button class="close-info">&times;</button>'+
    '<h3>'+info.title+'</h3>'+
    '<p>'+info.desc+'</p>'+
    '<div class="info-label">Columns Used</div><p>'+info.cols+'</p>'+
    '<div class="info-label">Data Source</div><p>'+info.source+'</p>'+
    noteHtml;
  document.body.appendChild(overlay); document.body.appendChild(popup);
  function close(){overlay.remove();popup.remove();}
  overlay.onclick=close;
  popup.querySelector('.close-info').onclick=close;
}
document.addEventListener('click',function(e){
  var btn=e.target.closest&&e.target.closest('.info-btn');
  if(btn){e.stopPropagation();showInfoPopup(btn.dataset.info);}
});

})();
