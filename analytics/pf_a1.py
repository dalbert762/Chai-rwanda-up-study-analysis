"""Page filter for Analysis 1 - availability.

The page's own explorer script (explorer.py) already redraws the maps,
FIG 1-6 and the all-medicines table from the same filter bar (fx-*). This
module redraws everything else on the page:

    FIG 7       upstream by level, one denominator (rows answered for both
                tiers), dots only
    FIG 8-9     RMS Regional against Central (UP rows)
    FIG 10-11   the form's distribution-issue / nationwide attribution
    1.2, 1.3    and their managed-only (method 2) variants, each from its
                own base (MOLS / MOLS_B) - both are on the page, the method
                switch only shows one
    1.9         nationwide stock-outs
    1.10        why medicines are not managed

Every rule is read from the report module R; the JS mirrors it line by
line. Charts the engine lacks (grouped, columns, area_band) are ported from
charts.py so their printed text matches.
"""


def build(pg, R):
    M, s = R.M, R.s

    def yn(r, col):
        return 1 if M.yes(r[col]) else (0 if M.no(r[col]) else None)

    # 1.2 / 1.3, method 1 (MOLS - every row answered) and method 2 (MOLS_B)
    pg.dataset("ma", R.MOLS, R.pf_fac, R.pf_med,
               {"av": lambda r: M.yes(r[M.AVAIL])})
    pg.dataset("mb", R.MOLS_B, R.pf_fac, R.pf_med,
               {"av": lambda r: M.yes(r[M.AVAIL])})

    # 1.10: the not-managed molecules, Essential Medicines then programmes
    why = [(r, t, 0) for r, t in R._WHY_EM] + [(r, t, 1) for r, t in R._WHY_PG]
    th = {id(r): (t, k) for r, t, k in why}
    pg.dataset("why", [w[0] for w in why], R.pf_fac, R.pf_med, {
        "th": lambda r: th[id(r)][0], "pg": lambda r: th[id(r)][1]})

    # FIG 7-11 and 1.9: the unavailable rows
    pg.dataset("up", R.UP, R.pf_fac, R.pf_med, {
        "reg": lambda r: yn(r, M.RMS_REG),
        "cen": lambda r: yn(r, M.RMS_CEN),
        "dn": lambda r: {R.DN_ISSUE_T: 1, R.DN_NS_T: 2}.get(
            s(r[M.DIST_FLAG]), 0),
        "dcl": lambda r: s(r[M.DCLASS]),
        "prog": lambda r: s(r[M.PROGRAM])})
    # 1.9's "facilities managing it": item-level managed rows, only for the
    # medicines that can appear in the table
    ns_meds = {s(r[M.MED]) for r in R.ns_rows}
    pg.dataset("mg", [r for r in R.MG if s(r[M.MED]) in ns_meds],
               R.pf_fac, R.pf_med)

    rep = {"__LEVELS__": R.json.dumps(R.LEVELS),
           "__HIDDEN__": R.json.dumps(list(R.WHY_HIDDEN)),
           "__LAST__": R.json.dumps(list(R.WHY_LAST)),
           "__GROUP_COLS__": R.json.dumps(R.CH.GROUP_COLS),
           "__BAND__": R.json.dumps([R.CH.BAND_A, R.CH.BAND_B,
                                     R.CH.BAND_FILL])}
    js = JS
    for k, v in rep.items():
        js = js.replace(k, v)
    return js


JS = r"""
var LEVELS=__LEVELS__,WHY_HIDDEN=__HIDDEN__,WHY_LAST=__LAST__,
    GCOLS=__GROUP_COLS__,BAND=__BAND__,W1=940;
function fname(ds){return function(i){return P.fac(ds,i)[0];};}
function mname(ds){return function(i){return P.med(ds,i)[0];};}
function sortedKeys(b){return b.keys.slice().sort(P.pyStrCmp);}
// ---- charts.grouped()
function grouped(rows,series){
  if(!rows.length)return '<p class="fempty">Nothing to show.</p>';
  var LW=210,BH=15,IN=3,GP=12,PT=42,k=series.length,grp=k*BH+(k-1)*IN,
      h=PT+rows.length*(grp+GP)+18,pw=W1-LW-78,
      o=['<svg class="chart" viewBox="0 0 '+W1+' '+h+'" role="img">'];
  series.forEach(function(nm,i){var x=LW+i*150;
    o.push('<rect x="'+x+'" y="'+(PT-32)+'" width="11" height="11" rx="2" '
      +'fill="'+GCOLS[i%GCOLS.length]+'"/><text x="'+(x+16)+'" y="'+(PT-22)
      +'" class="cl">'+P.esc(nm)+'</text>');});
  [0,25,50,75,100].forEach(function(t){var x=LW+pw*t/100;
    o.push('<line class="gl'+(t===0?' z':'')+'" x1="'+P.fx(x,1)+'" y1="'
      +(PT-8)+'" x2="'+P.fx(x,1)+'" y2="'+(h-16)+'"/><text x="'+P.fx(x,1)
      +'" y="'+(h-4)+'" class="cn" text-anchor="middle">'+t+'%</text>');});
  rows.forEach(function(r,i){var lab=r[0],vals=r[1],y0=PT+i*(grp+GP);
    o.push('<text x="'+(LW-10)+'" y="'+P.fx(y0+grp/2+4,0)+'" class="ct" '
      +'text-anchor="end">'+P.esc(lab)+'</text>');
    vals.forEach(function(p,j){var n=p[0],hit=p[1],y=y0+j*(BH+IN),
        v=n?hit/n:0,w=pw*Math.min(v,1);
      o.push('<rect x="'+LW+'" y="'+y+'" width="'+P.fx(pw,1)+'" height="'
        +BH+'" fill="'+CFG.rest+'"/>');
      o.push('<rect class="seg" x="'+LW+'" y="'+y+'" width="'+P.fx(w,1)
        +'" height="'+BH+'" fill="'+GCOLS[j%GCOLS.length]+'"><title>'
        +P.esc(lab)+' - '+P.esc(series[j])+': '+P.pc1(v)+' ('+P.int(hit)
        +' of '+P.int(n)+')</title></rect>');
      o.push('<text x="'+P.fx(LW+w+6,1)+'" y="'+(y+BH-3)+'" class="cv" '
        +'style="font-size:10px">'+P.pc1(v)+'</text>');});
    o.push('<text x="'+(W1-6)+'" y="'+P.fx(y0+grp/2+4,0)+'" class="cn" '
      +'text-anchor="end">n='+P.int(vals[0][0])+' / '
      +P.int(vals[vals.length-1][0])+'</text>');});
  return o.join('')+'</svg>';}
function wrapLab(lab,lim){var words=String(lab).split(/\s+/).filter(Boolean),
    lines=[],cur='';
  words.forEach(function(w){if(cur.length+w.length+1<=lim)
      cur=(cur+' '+w).trim();else{lines.push(cur);cur=w;}});
  lines.push(cur);return lines.slice(0,3);}
// ---- charts.columns()
function columns(rows,series){
  if(!rows.length)return '<p class="fempty">Nothing to show.</p>';
  var H=330,PL=56,PB=74,PT=46,BW=26,IN=4,k=series.length,ph=H-PT-PB,
      pw=W1-PL-16,step=pw/rows.length,
      o=['<svg class="chart" viewBox="0 0 '+W1+' '+H+'" role="img">'];
  series.forEach(function(nm,i){var x=PL+i*150;
    o.push('<rect x="'+x+'" y="'+(PT-30)+'" width="11" height="11" rx="2" '
      +'fill="'+GCOLS[i%GCOLS.length]+'"/><text x="'+(x+16)+'" y="'+(PT-20)
      +'" class="cl">'+P.esc(nm)+'</text>');});
  [0,25,50,75,100].forEach(function(t){var y=PT+ph*(1-t/100);
    o.push('<line class="gl'+(t===0?' z':'')+'" x1="'+PL+'" y1="'+P.fx(y,1)
      +'" x2="'+(W1-16)+'" y2="'+P.fx(y,1)+'"/><text x="'+(PL-8)+'" y="'
      +P.fx(y+4,1)+'" class="cn" text-anchor="end">'+t+'%</text>');});
  rows.forEach(function(r,i){var lab=r[0],vals=r[1],cx=PL+step*(i+0.5),
      gw=k*BW+(k-1)*IN,x0=cx-gw/2;
    vals.forEach(function(p,j){var n=p[0],hit=p[1],
        v=Math.min(n?hit/n:0,1),bh=ph*v,x=x0+j*(BW+IN),y=PT+ph-bh;
      o.push('<rect x="'+P.fx(x,1)+'" y="'+P.fx(PT,1)+'" width="'+BW
        +'" height="'+P.fx(ph,1)+'" fill="'+CFG.rest+'" opacity=".5"/>');
      o.push('<rect class="seg" x="'+P.fx(x,1)+'" y="'+P.fx(y,1)+'" width="'
        +BW+'" height="'+P.fx(bh,1)+'" fill="'+GCOLS[j%GCOLS.length]
        +'"><title>'+P.esc(lab)+' - '+P.esc(series[j])+': '+P.pc1(v)+' ('
        +P.int(hit)+' of '+P.int(n)+')</title></rect>');
      o.push('<text x="'+P.fx(x+BW/2,1)+'" y="'+P.fx(y-4,1)+'" class="cv" '
        +'style="font-size:10px" text-anchor="middle">'+P.fx(v*100,1)
        +'%</text>');});
    var lines=wrapLab(lab,16);
    lines.forEach(function(line,li){o.push('<text x="'+P.fx(cx,1)+'" y="'
      +P.fx(PT+ph+16+li*11,0)+'" class="cn" text-anchor="middle">'
      +P.esc(line)+'</text>');});
    o.push('<text x="'+P.fx(cx,1)+'" y="'+P.fx(PT+ph+16+lines.length*11,0)
      +'" class="cs" text-anchor="middle">n='+P.int(vals[0][0])+' / '
      +P.int(vals[vals.length-1][0])+'</text>');});
  return o.join('')+'</svg>';}
// ---- charts.area_band()
function areaBand(rows,series,connect){
  if(!rows.length||series.length!==2)
    return '<p class="fempty">Nothing to show.</p>';
  var H=360,PL=56,PB=78,PT=50,vals=[];
  rows.forEach(function(r){r[1].forEach(function(p){
    vals.push(p[0]?p[1]/p[0]:0);});});
  var lo=Math.max(0,Math.min.apply(null,vals)-.08),
      hi=Math.min(1,Math.max.apply(null,vals)+.06),
      ph=H-PT-PB,pw=W1-PL-20,IN=70,step=(pw-2*IN)/Math.max(1,rows.length-1),
      nmax=Math.max.apply(null,rows.map(function(r){
        return Math.max.apply(null,r[1].map(function(p){return p[0];}));}))||1;
  function X(i){return PL+IN+step*i;}
  function Y(v){return PT+ph*(1-(v-lo)/(hi-lo));}
  var DR=11.0,LIFT=connect?13:DR+7,DROP=connect?18:DR+12,SIDE=connect?7:DR+5;
  function Rr(n){return connect?3.2+6.0*Math.sqrt(n/nmax):DR;}
  var o=['<svg class="chart" viewBox="0 0 '+W1+' '+H+'" role="img">'],
      t=Math.ceil(lo*20)/20;
  while(t<=hi+1e-9){var y=Y(t);
    o.push('<line class="gl" x1="'+PL+'" y1="'+P.fx(y,1)+'" x2="'+(W1-20)
      +'" y2="'+P.fx(y,1)+'"/><text x="'+(PL-8)+'" y="'+P.fx(y+4,1)
      +'" class="cn" text-anchor="end">'+P.fx(t*100,0)+'%</text>');
    t+=.05;}
  var a=rows.map(function(r){var p=r[1][0];return p[0]?p[1]/p[0]:0;}),
      b=rows.map(function(r){var p=r[1][1];return p[0]?p[1]/p[0]:0;});
  function pts(v,rev){var ix=v.map(function(x,i){return i;});
    if(rev)ix.reverse();
    return ix.map(function(i){return P.fx(X(i),1)+','+P.fx(Y(v[i]),1);})
      .join(' L');}
  if(connect){o.push('<path d="M'+pts(b)+' L'+pts(a,true)+' Z" fill="'
    +BAND[2]+'" opacity=".22"/>');
  [[b,BAND[1]],[a,BAND[0]]].forEach(function(s){
    o.push('<path d="M'+pts(s[0])+'" fill="none" stroke="'+s[1]
      +'" stroke-width="2.6" stroke-linejoin="round"/>');});}
  rows.forEach(function(r,i){var lab=r[0],vs=r[1];
    [[b,BAND[1],1],[a,BAND[0],0]].forEach(function(s){var n=vs[s[2]][0],
        hit=vs[s[2]][1];
      o.push('<circle cx="'+P.fx(X(i),1)+'" cy="'+P.fx(Y(s[0][i]),1)+'" r="'
        +P.fx(Rr(n),1)+'" fill="'+s[1]+'" stroke="#0e0e0d" stroke-width="1.4">'
        +'<title>'+P.esc(lab)+' - '+P.esc(series[s[2]])+': '+P.pc1(s[0][i])
        +' ('+P.int(hit)+' of '+P.int(n)+')</title></circle>');});
    var gap=a[i]-b[i],top=Math.max(a[i],b[i]),bot=Math.min(a[i],b[i]);
    // each value carries its count (hit/n), the gap its n (2026-11-05)
    var tj=a[i]>=b[i]?0:1,gn=vs[0][0]===vs[1][0]?'n='+P.int(vs[0][0])
      :'n='+P.int(vs[0][0])+' / '+P.int(vs[1][0]);
    o.push('<text x="'+P.fx(X(i),1)+'" y="'+P.fx(Y(top)-LIFT,1)+'" class="cv" '
      +'style="font-size:10.5px" text-anchor="middle">'+P.fx(top*100,1)
      +'% ('+P.int(vs[tj][1])+'/'+P.int(vs[tj][0])+')</text>');
    o.push('<text x="'+P.fx(X(i),1)+'" y="'+P.fx(Y(bot)+DROP,1)+'" class="cv" '
      +'style="font-size:10.5px" text-anchor="middle">'+P.fx(bot*100,1)
      +'% ('+P.int(vs[1-tj][1])+'/'+P.int(vs[1-tj][0])+')</text>');
    if(Math.abs(gap)>=.02){var g=P.fx(gap*100,1),tx=X(i)+SIDE;
      if(!connect){var bx=X(i)+DR+4,bw=9,y1=Y(top),y2=Y(bot),ym=(y1+y2)/2,
          q=Math.min(6,Math.abs(y2-y1)/4),m=bx+bw/2,f=function(v){
            return P.fx(v,1);};
        o.push('<path d="M'+f(bx)+','+f(y1)+' Q'+f(m)+','+f(y1)+' '+f(m)+','
          +f(y1+q)+' L'+f(m)+','+f(ym-q)+' Q'+f(m)+','+f(ym)+' '+f(bx+bw)+','
          +f(ym)+' Q'+f(m)+','+f(ym)+' '+f(m)+','+f(ym+q)+' L'+f(m)+','
          +f(y2-q)+' Q'+f(m)+','+f(y2)+' '+f(bx)+','+f(y2)+'" fill="none" '
          +'stroke="#c3c2b7" stroke-width="1.5" stroke-linecap="round"/>');
        tx=bx+bw+5;}
      // the gap written out (2026-11-05): first minus second series
      var na=vs[0][0],ha=vs[0][1],nb=vs[1][0],hb=vs[1][1],ym=(Y(top)+Y(bot))/2,
          _u=0;
      var nd=na===nb?'n='+P.int(Math.abs(ha-hb)):gn;
      o.push('<text x="'+P.fx(tx,1)+'" y="'+P.fx(ym+4,1)
        +'" class="cs" text-anchor="start">'+(g[0]==='-'?g:'+'+g)
        +' pts ('+nd+')</text>');}
    var lines=wrapLab(lab,15);
    lines.forEach(function(line,li){o.push('<text x="'+P.fx(X(i),1)+'" y="'
      +P.fx(PT+ph+18+li*11,0)+'" class="cn" text-anchor="middle">'
      +P.esc(line)+'</text>');});
    o.push('<text x="'+P.fx(X(i),1)+'" y="'+P.fx(PT+ph+18+lines.length*11,0)
      +'" class="cs" text-anchor="middle">n='+P.int(vs[0][0])
      +(vs[0][0]===vs[1][0]?'':' / '+P.int(vs[1][0]))+'</text>');
});
  series.forEach(function(nm,i){var x=PL+i*170,col=i===0?BAND[0]:BAND[1];
    if(connect)o.push('<line x1="'+x+'" y1="'+(PT-26)+'" x2="'+(x+22)
      +'" y2="'+(PT-26)+'" stroke="'+col+'" stroke-width="2.6"/>');
    o.push('<circle cx="'+(x+11)+'" cy="'
      +(PT-26)+'" r="'+(connect?4:6)+'" fill="'+col+'"/><text x="'+(x+29)+'" y="'+(PT-22)
      +'" class="cl">'+P.esc(nm)+'</text>');});
  if(connect)o.push('<text x="'+(PL+360)+'" y="'+(PT-22)+'" class="cs">'
    +'shaded band = the gap &#183; marker size = observations</text>');
  return o.join('')+'</svg>';}

// ---- the scorecards: _rate_of(), _low_high(), _avail_cards()
function rateOf(ds,ix){var C=P.D[ds].c;return ix.length
  ?P.cnt(ix,function(i){return C.av[i];})/ix.length:null;}
// [rate, key, n] lowest and highest of the qualifying groups, sorted like
// Python's tuples (rate, key, n); null when no group qualifies
function lowHigh(ds,ix,keyfn,minFac){var b=P.by(ix,keyfn),rs=[];
  b.keys.forEach(function(k){var g=b.g[k];
    if(k&&g.length>=P.MINN&&P.uniq(g,fname(ds))>=minFac)
      rs.push([rateOf(ds,g),k,g.length]);});
  if(!rs.length)return null;
  rs.sort(function(a,c){return a[0]-c[0]||P.pyStrCmp(a[1],c[1])
    ||a[2]-c[2];});
  return [rs[0],rs[rs.length-1]];}
function setCard(cid,v,l,sd){[['-v',v],['-l',l],['-s',sd]].forEach(
  function(p){var e=document.getElementById(cid+p[0]);
    if(e)e.textContent=p[1];});}
function availCards(ds,other,basis,otherName,cid){
  var pc=function(x){return P.fx(x*100,1)+'%';},ix=P.rows(ds),
      ox=P.rows(other),r0=rateOf(ds,ix),r1=rateOf(other,ox);
  setCard(cid+'-0',r0===null?'-':pc(r0),'availability - '+basis,
    r1===null?'no rows on '+otherName:pc(r1)+' on '+otherName);
  var lv=lowHigh(ds,ix,function(i){return P.fac(ds,i)[3];},5),
      pg=lowHigh(ds,ix,function(i){var m=P.med(ds,i);return m?m[1]:'';},1),
      di=lowHigh(ds,ix,function(i){return P.fac(ds,i)[1];},1);
  function none(what){return [ '-','no '+what+' with enough data',
    'too few rows under these filters'];}
  var c1=lv?[pc(lv[0][0]),lv[0][1].toLowerCase()+'s - lowest level',
      lv[1][1].toLowerCase()+'s highest at '+pc(lv[1][0])]:none('level'),
    c2=pg?[pc(pg[0][0]),pg[0][1]+' - lowest programme',
      pg[1][1]+' highest at '+pc(pg[1][0])]:none('programme'),
    c3=di?[pc(di[0][0]),di[0][1]+' - lowest district',
      di[1][1]+' highest at '+pc(di[1][0])]:none('district');
  setCard(cid+'-1',c1[0],c1[1],c1[2]);
  setCard(cid+'-2',c2[0],c2[1],c2[2]);
  setCard(cid+'-3',c3[0],c3[1],c3[2]);}

P.render(function(){
  // FIG 6: every hospital on the method's base, lowest first (as _hosp_rows)
  (function(){var ds=P.meth()===2?'mb':'ma',ix=P.rows(ds),AV=P.D[ds].c.av,
      H=['District hospital','Level 2 teaching hospital',
         'National referral hospital'],g={};
    ix.forEach(function(i){var f=P.fac(ds,i);if(H.indexOf(f[3])<0)return;
      var v=g[f[0]]||(g[f[0]]=[0,0]);v[0]++;v[1]+=AV[i]?1:0;});
    var rows=Object.keys(g).map(function(k){return [k,g[k][0],g[k][1]];});
    rows.sort(function(a,b){return a[2]/a[1]-b[2]/b[1]||P.pyStrCmp(a[0],b[0]);});
    P.put('1|FIG_6',P.bars(rows,'available','expected',null,[12,4]));})();
  // ---------- scorecards, both methods (the switch shows one set)
  availCards('ma','mb','expected medicines','managed products only','pc1a');
  availCards('mb','ma','managed products only','expected medicines','pc1b');
  // ---------- 1.2 / 1.3 from one base
  function leastMeds(ds,sorted){var ix=P.rows(ds),C=P.D[ds].c,
      b=P.by(ix,mname(ds)),rows=[];
    (sorted?sortedKeys(b):b.keys).forEach(function(m){var rs=b.g[m];
      if(rs.length<P.MINN)return;
      var av=P.cnt(rs,function(i){return C.av[i];});
      rows.push([m,P.med(ds,rs[0])[1],P.uniq(rs,fname(ds)),rs.length,av,
                 av/rs.length]);});
    return P.sort(rows,function(a,b){return a[5]-b[5];}).slice(0,10)
      .map(function(r){r[5]=P.R(r[5]);return r;});}
  function weakFacs(ds){var ix=P.rows(ds),C=P.D[ds].c,
      b=P.by(ix,fname(ds)),rows=[];
    sortedKeys(b).forEach(function(f){var rs=b.g[f],F=P.fac(ds,rs[0]),
        av=P.cnt(rs,function(i){return C.av[i];});
      rows.push([f,F[1],F[3],rs.length,av,rs.length-av,av/rs.length]);});
    return P.sort(rows,function(a,b){return a[6]-b[6];}).slice(0,20)
      .map(function(r){r[6]=P.R(r[6]);return r;});}
  var H12=['Medicine','Programme','Facilities','Answered','Available',
           'AVAILABLE %'];
  P.put('1|1.2',P.table(H12,leastMeds('ma',true)));
  P.put('1|1.2#2',P.table(H12,leastMeds('mb',false)));
  P.put('1|1.3',P.table(['Facility','District','Level of care',
    'Expected assessed','Available','Missing','FACILITY AVAILABILITY %'],
    weakFacs('ma')));
  P.put('1|1.3#2',P.table(['Facility','District','Level of care',
    'Managed assessed','Available','Missing','FACILITY AVAILABILITY %'],
    weakFacs('mb')));

  // ---------- FIG 7-11 and 1.9: the unavailable rows
  var ux=P.rows('up'),U=P.D.up.c,
      lvl=function(i){return P.fac('up',i)[3];},
      cls=function(i){return U.dcl[i]||'(unclassified)';},
      prg=function(i){return U.prog[i]||'(none)';};
  // up_by(): {k: [n, hits]} and the order Python's list would have
  function upBy(col,keyfn,keys,floor){var c=U[col],
      b=P.by(ux.filter(function(i){return c[i]!==null;}),keyfn),out=[];
    (keys||sortedKeys(b)).forEach(function(k){var rs=b.g[k];
      if(!rs||rs.length<floor)return;
      out.push([k,rs.length,P.cnt(rs,function(i){return c[i]===1;})]);});
    var m={};out.forEach(function(r){m[r[0]]=[r[1],r[2]];});return m;}
  // _grouped(): Regional and Central side by side
  function grp(keyfn,keys,floor){var reg=upBy('reg',keyfn,keys,floor),
      cen=upBy('cen',keyfn,keys,floor),order=keys;
    if(!order){order=Object.keys(reg).filter(function(k){return k in cen;})
      .sort(P.pyStrCmp);
      order=P.sort(order,function(a,b){return reg[b][1]/reg[b][0]
        -reg[a][1]/reg[a][0];});}
    return order.filter(function(k){return k in reg||k in cen;})
      .map(function(k){return [k,[reg[k]||[0,0],cen[k]||[0,0]]];});}
  // FIG 7-10 share one base: rows answered for BOTH tiers (UP_BOTH)
  var bx=ux.filter(function(i){return U.reg[i]!==null&&U.cen[i]!==null;});
  (function(){var b=P.by(bx,lvl),rows=[];
    LEVELS.forEach(function(k){var rs=b.g[k];if(!rs)return;
      rows.push([k,[[rs.length,P.cnt(rs,function(i){return U.cen[i]===1;})],
        [rs.length,P.cnt(rs,function(i){return U.reg[i]===1;})]]]);});
    P.put('1|FIG_7',areaBand(rows,['RMS Central','RMS Regional'],false));
  })();
  // _both_by(): sorted keys, then Regional share descending (stable)
  (function(){var b=P.by(bx,prg),rows=[];
    sortedKeys(b).forEach(function(k){var rs=b.g[k];if(rs.length<P.MINN)return;
      rows.push([k,[[rs.length,P.cnt(rs,function(i){return U.reg[i]===1;})],
        [rs.length,P.cnt(rs,function(i){return U.cen[i]===1;})]]]);});
    rows=P.sort(rows,function(a,c){return c[1][0][1]/c[1][0][0]
      -a[1][0][1]/a[1][0][0];});
    P.put('1|FIG_8',grouped(rows,['RMS Regional','RMS Central']));})();
  // dn_by()
  var dx=bx.filter(function(i){return U.dn[i]>0;});
  function dnBy(keyfn,keys,floor){var b=P.by(dx,keyfn),order=keys;
    function ns(k){return P.cnt(b.g[k],function(i){return U.dn[i]===2;})
      /b.g[k].length;}
    if(!order)order=P.sort(b.keys,function(a,c){return ns(c)-ns(a);});
    return order.filter(function(k){return (b.g[k]||[]).length>=floor;})
      .map(function(k){var rs=b.g[k];return [k,[
        [rs.length,P.cnt(rs,function(i){return U.dn[i]===1;})],
        [rs.length,P.cnt(rs,function(i){return U.dn[i]===2;})]]];});}
  var DS=['Distribution Gap','Nationwide stock-out'];
  P.put('1|FIG_9',columns(dnBy(lvl,LEVELS,1),DS));
  P.put('1|FIG_10',grouped(dnBy(prg,null,P.MINN),DS));
  // 1.9: NS_TABLE
  (function(){var ns=ux.filter(function(i){return U.dn[i]===2;}),
      b=P.by(ns,mname('up')),
      mg=P.by(P.rows('mg'),mname('mg')),rows=[];
    sortedKeys(b).forEach(function(m){var rs=b.g[m],
        nf=P.uniq(rs,fname('up')),nm=(mg.g[m]||[]).length;
      rows.push([m,U.prog[rs[0]],U.dcl[rs[0]],nf,
        P.uniq(rs,function(i){return P.fac('up',i)[1];}),nm,
        P.R(P.pct(nf,nm))]);});
    rows=P.sort(rows,function(a,c){return c[3]-a[3];});
    P.put('1|1.9',P.table(['Medicine','Programme','Drug class',
      'Facilities reporting it','Districts','Facilities managing it',
      'SHARE OF MANAGING FACILITIES'],rows));
  })();

  // ---------- 1.10: why medicines are not managed
  (function(){var ix=P.rows('why'),C=P.D.why.c,em={},pg={},seen=[],
      meds={},ne=0,np=0;
    ix.forEach(function(i){var t=C.th[i],m=P.med('why',i)[0];
      if(!(t in em)){em[t]=0;pg[t]=0;seen.push(t);}
      if(C.pg[i]){pg[t]++;np++;}else{em[t]++;ne++;}
      if(!meds[t])meds[t]={k:[],n:{}};
      if(!(m in meds[t].n)){meds[t].n[m]=0;meds[t].k.push(m);}
      meds[t].n[m]++;});
    var order=P.sort(seen.filter(function(t){return WHY_LAST.indexOf(t)<0
        &&WHY_HIDDEN.indexOf(t)<0;}).sort(P.pyStrCmp),
      function(a,b){return (em[b]+pg[b])-(em[a]+pg[a]);})
      .concat(WHY_LAST.filter(function(t){return t in em;})),rows=[];
    order.forEach(function(t){var mc=meds[t],
        top=P.sort(mc.k,function(a,b){return mc.n[b]-mc.n[a];}).slice(0,3)
          .map(function(m){return m.slice(0,30)+' ('+mc.n[m]+')';}).join(', ');
      rows.push([t,em[t],P.R(P.pct(em[t],ne)),pg[t],P.R(P.pct(pg[t],np)),
        em[t]+pg[t],top]);});
    rows.push(['TOTAL not managed',ne,P.R(1),np,
      P.R(1),ne+np,'']);
    P.put('1|1.10',P.table(['Reason (coded from the free-text answer)',
      'Essential Medicines - authorised, not managed','% of these',
      'Programme products - not managed','% of these','Total',
      'Medicines citing it most'],rows));
  })();
});
"""
