"""Page filter for Analysis 4 - stock-out frequency and duration.

The worked example for the pf_a<N> modules: build(pg, R) registers the rows
the page's figures and tables are computed from, and returns the script that
recomputes each of them from the filtered rows. R is the report module, so
every rule (months affected, the 92-day window, MIN_N) is read from it, not
restated.

Outputs: FIG 4.1-4.5, tables 4.1 and 4.7.
"""


def build(pg, R):
    M, s = R.M, R.s
    so_ids = {id(r) for r in R.SO}
    pg.dataset("mg", R.MG, R.pf_fac, R.pf_med, {
        "lvl": lambda r: r["_lvl"],
        "cls": R._class_label,
        "prog": lambda r: s(r[M.PROGRAM]),
        "med": lambda r: s(r[M.MED]),
        "fac": R.mfac,
        "so": lambda r: R.months_affected(r) > 0,
        "wd": R._wdays,
        # average days out: total_stockout_days (R._tdays), 2026-10-26
        "td": R._tdays,
        "out": lambda r: id(r) in so_ids,
        # FIG 4.5-4.6, table 4.8: bit k = month k (Jun, Jul, Aug) was out /
        # had a request above 0
        "sm": lambda r: sum(1 << k for k, m in enumerate(R.MONTHS)
                            if (M.num(r[M.SO_MONTH[m]]) or 0) > 0),
        "rm": lambda r: sum(1 << k for k, m in enumerate(R.MONTHS)
                            if (M.num(r[M.REQ[m]]) or 0) > 0)})
    return (JS.replace("__LEVELS__", R.json.dumps(R.LEVELS))
              .replace("__SOREQ__", R.json.dumps([R.SOREQ_SERIES,
                                                  R.SOREQ_COLS]))
              .replace("__ROSE__", R.json.dumps([R.CH.ROSE_COLS]))
              .replace("__SOMAP__", R.json.dumps(
                  {"steps": [list(x) for x in R.SO_STEPS],
                   "dark": list(R.SO_DARK_TEXT), "good": R.SO_GOOD,
                   "bad": R.SO_BAD})))


JS = r"""
var LEVELS=__LEVELS__,SOREQ=__SOREQ__,ROSE=__ROSE__,SOMAP=__SOMAP__;
// gauge cards, matching charts.gauge()/gauges() (FIG 4.1)
function gauge(value,caption){var T=42,w=250,h=148,cx=w/2,cy=h-18,ro=w/2-21,
    ri=w/2-47,lit=P.pyRound(Math.max(0,Math.min(1,value))*T),
    o=['<svg viewBox="0 0 '+w+' '+h+'" role="img" aria-label="'
      +P.fx(value*100,1)+' percent">'];
  for(var i=0;i<T;i++){var a=Math.PI*(1+(i+0.5)/T),c=Math.cos(a),s=Math.sin(a);
    o.push('<line x1="'+P.fx(cx+ri*c,1)+'" y1="'+P.fx(cy+ri*s,1)+'" x2="'
      +P.fx(cx+ro*c,1)+'" y2="'+P.fx(cy+ro*s,1)+'" stroke="'
      +(i<lit?SOMAP.bad:SOMAP.good)+'" stroke-width="3.4" '
      +'stroke-linecap="round"/>');}
  o.push('<text x="'+P.fx(cx,0)+'" y="'+P.fx(cy-34,0)+'" class="gval" '
    +'text-anchor="middle">'+P.fx(value*100,1)+'%</text>');
  o.push('<text x="'+P.fx(cx,0)+'" y="'+P.fx(cy-15,0)+'" class="gcap" '
    +'text-anchor="middle">'+P.esc(caption)+'</text>');
  return o.join('')+'</svg>';}
function gauges(rows,caption){if(!rows.length)return P.EMPTY;
  return '<div class="ggrid">'+rows.map(function(r){var v=r[1]?r[2]/r[1]:0;
    return '<div class="gcard">'+gauge(v,caption)+'<p class="gname">'
      +P.esc(r[0])+'</p><p class="gcount"><b>'+P.int(r[2])+'</b> <span>of '
      +P.int(r[1])+'</span></p></div>';}).join('')+'</div>';}
// FIG 4.1: half rose, matching charts.half_rose()
function rose(rows,centre,sub){rows=rows.filter(function(r){return r[1];});
  if(!rows.length)return '<p class="fempty">Nothing to show.</p>';
  var w=940,h=500,r0=118,rmax=205,cx=330,cy=h/2,k=rows.length,ri=r0-14,
      cols=ROSE[0],vmax=0,o=[],f=function(v){return P.fx(v,1);};
  rows.forEach(function(r){vmax=Math.max(vmax,r[2]/r[1]);});vmax=vmax||1;
  o.push('<svg class="chart" viewBox="0 0 '+w+' '+h+'" role="img">');
  rows.forEach(function(r,i){var v=r[2]/r[1],
      a0=(-90+i*180/k)*Math.PI/180,a1=(-90+(i+1)*180/k)*Math.PI/180,
      am=(a0+a1)/2,R=r0+22+(rmax-r0-22)*v/vmax,
      X=function(rr,a){return f(cx+rr*Math.cos(a));},
      Y=function(rr,a){return f(cy+rr*Math.sin(a));};
    o.push('<path d="M'+X(ri,a0)+','+Y(ri,a0)+' L'+X(R,a0)+','+Y(R,a0)
      +' A'+f(R)+','+f(R)+' 0 0 1 '+X(R,a1)+','+Y(R,a1)+' L'+X(ri,a1)+','
      +Y(ri,a1)+' A'+f(ri)+','+f(ri)+' 0 0 0 '+X(ri,a0)+','+Y(ri,a0)
      +' Z" fill="'+stepCol(v)+'" stroke="#000" stroke-width="2">'
      +'<title>'+P.esc(r[0])+': '+P.pc1(v)+' with a stock-out - '
      +P.int(r[2])+' of '+P.int(r[1])+'</title></path>');
    var rm=(r0+R)/2+6;
    o.push('<text x="'+X(rm,am)+'" y="'+f(cy+rm*Math.sin(am)+5)
      +'" text-anchor="middle" style="font-size:15px;font-weight:700;'
      +'fill:'+(SOMAP.dark.indexOf(stepCol(v))>=0?'#12303a':'#fff')+'">'
      +P.fx(v*100,1)+'%</text>');
    var lx=cx+(R+14)*Math.cos(am),ly=cy+(R+14)*Math.sin(am);
    o.push('<text x="'+f(lx)+'" y="'+f(ly-2)+'" class="ct" style="'
      +'font-weight:700">'+P.esc(r[0])+'</text><text x="'+f(lx)+'" y="'
      +f(ly+12)+'" class="cn">'+P.int(r[2])+' of '+P.int(r[1])+'</text>');});
  o.push('<circle cx="'+cx+'" cy="'+f(cy)+'" r="'+(r0+6)+'" fill="#e9e9e6" '
    +'opacity=".35"/><circle cx="'+cx+'" cy="'+f(cy)+'" r="'+r0
    +'" fill="#f4f4f2"/><circle cx="'+cx+'" cy="'+f(cy)+'" r="'+(r0-26)
    +'" fill="#fff" stroke="#dcdcd8"/>');
  o.push('<text x="'+cx+'" y="'+f(cy-2)+'" text-anchor="middle" style="'
    +'font-size:34px;font-weight:700;fill:#222">'+P.esc(centre)+'</text>'
    +'<text x="'+cx+'" y="'+f(cy+20)+'" text-anchor="middle" style="'
    +'font-size:11.5px;fill:#555">'+P.esc(sub)+'</text><rect x="'+(cx-34)
    +'" y="'+f(cy+30)+'" width="68" height="3" rx="1.5" fill="url(#rosebar)"/>');
  o.push('<defs><linearGradient id="rosebar"><stop offset="0" stop-color="'
    +cols[0]+'"/><stop offset=".5" stop-color="'+cols[2]+'"/><stop '
    +'offset="1" stop-color="'+cols[4]+'"/></linearGradient></defs></svg>');
  return o.join('');}
// FIG 4.4: recolour the two-view map in place (paths are in the page)
function stepCol(v){if(v===null)return '#2c2c2a';
  for(var i=0;i<SOMAP.steps.length;i++){var s=SOMAP.steps[i];
    if(v>=s[0]&&v<s[1])return s[2];}
  return SOMAP.steps[SOMAP.steps.length-1][2];}
function paintMap(dv,pv,overall,n){
  var root=window.__PF_CAPTURE?null:document.getElementById('somap');
  if(!root)return;
  [['d',dv],['p',pv]].forEach(function(L){var box=root.querySelector('.mp-'+L[0]);
    if(!box)return;var vals=L[1];
    box.querySelectorAll('path[data-n]').forEach(function(el){
      var g=vals[el.getAttribute('data-n')],v=g&&g[0]>=P.MINN?g[1]/g[0]:null;
      el.setAttribute('fill',stepCol(v));var t=el.querySelector('title');
      if(t){if(!el.dataset.lab)el.dataset.lab=t.textContent.split(' - ')[0];
        t.textContent=el.dataset.lab+' - '+(v===null?'fewer than '+P.MINN
          :P.pc1(v))+'\n'+P.int(g?g[1]:0)+' of '+P.int(g?g[0]:0)
          +' managed pairs with a stock-out';}});
    box.querySelectorAll('text[data-nv]').forEach(function(el){
      var g=vals[el.getAttribute('data-nv')];
      el.textContent=g&&g[0]>=P.MINN?P.pc1(g[1]/g[0]):'n/a';});
    var sb=box.querySelector('.msub');
    if(sb)sb.textContent=P.int(n)+' managed pairs - overall '+P.pc1(overall);});
  var sel=root.querySelector('.lyrsel');
  if(sel)Array.prototype.forEach.call(sel.options,function(op){if(!op.value)return;
    var nm=op.textContent.split(' · ')[0],g=pv[nm],
        v=g&&g[0]>=P.MINN?g[1]/g[0]:null;
    op.textContent=nm+' · '+(v===null?'n/a':P.pc1(v));});}
P.render(function(){
  var ix=P.rows('mg'),C=P.D.mg.c,
      lvl=function(i){return C.lvl[i];},cls=function(i){return C.cls[i];},
      prg=function(i){return C.prog[i]||'(none)';},
      fac=function(i){return C.fac[i];},so=function(i){return C.so[i];};
  var H='with a stock-out',B='managed pairs';
  var any0=ix.filter(so);
  P.put('4|FIG_4.1',rose(P.worstFirst(P.rateBy(ix,prg,so,fac)),
    P.pc1(P.pct(any0.length,ix.length)),P.int(any0.length)+' of '
    +P.int(ix.length)+' managed pairs'));
  function geo(keyfn){var o={};
    ix.forEach(function(i){var k=keyfn(i),g=o[k]||(o[k]=[0,0]);g[0]++;
      if(so(i))g[1]++;});return o;}
  paintMap(geo(function(i){return P.fac('mg',i)[1];}),
    geo(function(i){return P.fac('mg',i)[2];}),
    P.pct(any0.length,ix.length),ix.length);
  P.put('4|FIG_4.2',gauges(P.barRows(P.rateBy(ix,lvl,so,fac,LEVELS,null,
    false)),'with a stock-out'));
  // _days_fig: mean window days per affected pair, longest first
  function days(keyfn){var aff=ix.filter(function(i){return C.td[i]>0;}),
      b=P.by(aff,keyfn),rows=[];
    b.keys.forEach(function(k){var v=b.g[k];if(k&&v.length>=P.MINN)
      rows.push([k,v.length,P.mean(v.map(function(i){return C.td[i];}))]);});
    return P.sort(rows,function(a,b){return b[2]-a[2];});}
  // scale to the next 30 days above the longest mean, never under 92
  var d44=days(prg),top=d44.reduce(function(a,r){return Math.max(a,r[2]);},0);
  P.put('4|FIG_4.4',P.dayBars(d44,top<=92?92:Math.ceil(top/30)*30,null,
    [SOMAP.bad,SOMAP.good]));
  // 4.1
  var any=ix.filter(so),out=P.cnt(ix,function(i){return C.out[i];});
  // 4.1 is off the dashboard (HTML_DROP, 2026-10-19)
  // ---- FIG 4.5-4.6 and table 4.8: stock-out months requested or not
  function mix(rs){var a=0,b=0;rs.forEach(function(i){
    for(var k=0;k<3;k++)if(C.sm[i]&(1<<k)){if(C.rm[i]&(1<<k))a++;else b++;}});
    return [a,b];}
  var lb=P.by(ix,lvl),lrows=[];
  LEVELS.forEach(function(l){var v=lb.g[l]?mix(lb.g[l]):[0,0];
    if(v[0]+v[1]>=1)lrows.push([l,v]);});
  P.put('4|FIG_4.5',P.stacked(lrows,SOREQ[0],SOREQ[1]));
  var pb=P.by(ix,function(i){return C.prog[i];}),prows=[];
  pb.keys.slice().sort(P.pyStrCmp).forEach(function(k){if(!k)return;
    var v=mix(pb.g[k]);if(v[0]+v[1]>=1)prows.push([k,v]);});
  prows=P.sort(prows,function(a,b){return b[1][1]/(b[1][0]+b[1][1])
    -a[1][1]/(a[1][0]+a[1][1]);});
  P.put('4|FIG_4.6',P.stacked(prows,SOREQ[0],SOREQ[1]));
  var fb=P.by(ix,fac),frows=[];
  fb.keys.slice().sort(P.pyStrCmp).forEach(function(f){var rs=fb.g[f],
      v=mix(rs),n=v[0]+v[1];if(!n)return;
    frows.push([f,P.fac('mg',rs[0])[1],C.lvl[rs[0]],n,v[0],v[1],
      P.R(v[1]/n)]);});
  frows=P.sort(frows,function(a,b){return b[5]-a[5]||b[3]-a[3];});
  P.put('4|4.8',P.table(['Facility','District','Level of care',
    'Stock-out months','Requested that month','No request that month',
    'NO REQUEST %'],frows));
  // 4.7 average: total_stockout_days over the pairs where it is > 0,
  // round(., 1) as Python
  function tavg(rs){var o=rs.filter(function(i){return C.td[i]>0;});
    return o.length?P.pyRound(P.sum(o,C.td)/o.length,1):null;}
  // 4.7: every medicine, worst incidence first
  var bm=P.by(ix,function(i){return C.med[i];}),rows=[];
  bm.keys.slice().sort(P.pyStrCmp).forEach(function(m){var rs=bm.g[m],
      o=rs.filter(function(i){return C.wd[i]>0;}),
      d=Math.round(P.sum(rs,C.wd));
    rows.push([m,C.prog[rs[0]],rs.length,o.length,
      P.R(P.pct(o.length,rs.length)),d,rs.length*92,
      P.R(P.pct(d,rs.length*92)),tavg(rs)]);});
  rows=P.sort(rows,function(a,b){return (b[4].r||0)-(a[4].r||0);});
  P.put('4|4.7',P.table(['Medicine','Programme','Managed pairs',
    'Pairs ever out','INCIDENCE','Total days out',
    'Days possible (denominator)','DAYS OUT RATE','Average days out'],rows));
  // scorecards pc4-0..2, as page_cards() builds them, on the filtered rows
  function card(id,v,l,s){[['-v',v],['-l',l],['-s',s]].forEach(function(p){
    var e=document.getElementById(id+p[0]);if(e)e.textContent=p[1];});}
  var pc=function(x){return P.fx(x*100,1)+'%';},
      raw=function(i){return C.prog[i];};
  if(ix.length)card('pc4-0',pc(any.length/ix.length),
    'of managed medicines stocked out','at least one day, June to August');
  else card('pc4-0','-','of managed medicines stocked out',
    'no managed medicines match these filters');
  // rate_by over programmes, non-empty key, n >= MIN_N, highest rate first
  var pg=P.sort(P.rateBy(ix,raw,so,fac).filter(function(x){
      return x[0]&&x[4]!==null;}),function(a,b){return b[4]-a[4];});
  if(pg.length)card('pc4-1',pc(pg[0][4]),
    pg[0][0]+' - most frequent stock-outs',
    pg[pg.length-1][0]+' least at '+pc(pg[pg.length-1][4]));
  else card('pc4-1','-','most frequent stock-outs',
    'no programme with '+P.MINN+' or more managed pairs');
  // max of (mean window days, programme) over programmes with MIN_N+
  // affected pairs; a tie on the mean goes to the later programme name
  var bd=P.by(ix.filter(function(i){return C.td[i]>0;}),raw),lg=null;
  bd.keys.forEach(function(k){var v=bd.g[k];if(!k||v.length<P.MINN)return;
    var m=P.mean(v.map(function(i){return C.td[i];}));
    if(!lg||m>lg[0]||(m===lg[0]&&k>lg[1]))lg=[m,k];});
  if(lg)card('pc4-2',P.fx(lg[0],0)+' days',lg[1]+' - longest stock-outs',
    'mean days out per medicine that ran out');
  else card('pc4-2','-','longest stock-outs',
    'no programme with '+P.MINN+' or more medicines that ran out');
});
"""
