"""Page filter for Analysis 5 - prescription fulfilment.

Two datasets from the PRESCRIPTION tool (not the stock tool):
  rx  every prescription (R.PRES); its medicines are the items on it, so a
      prescription matches a medicine / programme filter when any item does.
  it  every prescribed item (R.ITEMS).
The facility record comes from the prescription submission - name, district,
province, level of care and the management model FIG 5.3 borrows from the
stock assessment by name. The medicine record is the item's readable name,
derived programme and EML drug class, as the page groups them.

Outputs: FIG 5.1-5.7, TABLE 5A, tables 5.1-4.5, 5.7, 5.8, 5.9, 5.10, 5.11,
5.12. FIG 5.2's district / province outlines are geometry that never changes
with the filters; the script reads them from the page once instead of
embedding them a second time.
"""


def build(pg, R):
    P, s, FL, MAP = R.P, R.s, R.FL, R.MAP

    def fac_rec(r):
        return [R.pfac(r), R.pdist(r), R.pprov(r), R.rx_level(r),
                R.rx_mgmt(r) or "(not recorded)"]

    def med_rec(r):
        return [R.rx_med_name(r), R.rx_prog(r), R.rx_class(r)]

    # LEAN PAYLOAD. A prescription's medicines, item count, unserved count
    # and outcome are all read off its items in the browser, so the rx
    # dataset carries only its facility and item count. Items arrive grouped
    # by prescription in PRES order, so the item -> prescription link is the
    # run of item counts. If a data refresh ever breaks either assumption the
    # explicit columns are embedded instead, and the script uses them.
    items_of = {}
    for r in R.ITEMS:
        items_of.setdefault(id(r["_pres"]), []).append(r)
    pidx = {id(p): i for i, p in enumerate(R.PRES)}
    seq = [pidx[id(r["_pres"])] for r in R.ITEMS]
    grouped = all(a <= b for a, b in zip(seq, seq[1:])) and all(
        p["_n_items"] == len(items_of.get(id(p), [])) for p in R.PRES)

    def derivable(p):
        it = items_of.get(id(p), [])
        nu = sum(1 for r in it if r["_avail"] is False)
        return (p["_n_unavail"] == nu
                and bool(p["_any_unserved"]) == (nu > 0)
                and bool(p["_fully_served"]) == (
                    bool(it) and all(r["_avail"] is True for r in it)))
    flags = all(derivable(p) for p in R.PRES)

    rcols = {"ni": lambda r: r["_n_items"]}
    if not flags:
        rcols.update({"fs": lambda r: bool(r["_fully_served"]),
                      "any": lambda r: bool(r["_any_unserved"]),
                      "nu": lambda r: r["_n_unavail"]})
    # a prescription's name types are its items' - set from the items in
    # the browser, like its medicine list
    pg.dataset("rx", R.PRES, fac_rec, None, rcols, tag=lambda r: [])
    icols = {}
    if not grouped:
        icols["pi"] = lambda r: pidx[id(r["_pres"])]
    pg.dataset("it", R.ITEMS, fac_rec, med_rec, icols,
               tag=lambda r: s(r.get(P.NAME_TYPE)) or None)
    # rebuilt in the browser: every rx medicine list, and - when items are
    # grouped - each item's facility, which is its prescription's
    pg.D["rx"]["m"] = []
    if grouped:
        pg.D["it"]["f"] = []
    # availability as one character per item: 1 available, 0 not, - blank
    av = "".join("-" if r["_avail"] is None else "1" if r["_avail"] is True
                 else "0" for r in R.ITEMS)

    # Per-item attributes that are nearly constant per medicine: the
    # medicine record carries programme and class as first seen; the raw
    # form code and the form's class text are given once per medicine, and
    # only the items that differ are listed.
    K0 = sorted({r["_class"] for r in R.ITEMS})
    k0 = {k: i for i, k in enumerate(K0)}
    first, exc = {}, {}
    for i, r in enumerate(R.ITEMS):
        mi = pg._mi[R.rx_med_name(r)]
        v = [s(r["_med"]), k0[r["_class"]], R.rx_class(r), R.rx_prog(r)]
        if mi not in first:
            first[mi] = v
        elif v != first[mi]:
            exc[i] = v
    dflt = [first[i][:2] for i in range(len(pg.M))]
    # quantity and value only matter on unserved items
    qc = {i: [P.num(r.get(P.QTY)), P.num(r.get(P.COST_UNSERVED))]
          for i, r in enumerate(R.ITEMS) if r["_avail"] is False}

    dun = [[d["district"], d["prov"], MAP.slug(d["prov"] or d["district"]),
            d["cx_cy"][0], d["cx_cy"][1]]
           for d in MAP.DROWS]
    pun = []
    for p in MAP.PROWS:
        cx, cy = MAP._prov_centroid(p["prov"])
        pun.append([p["prov"], p["prov"], MAP.slug(p["prov"]),
                    cx, cy])
    W, H, PAD = MAP.W, MAP.H, MAP.PAD
    tail = (MAP.north_arrow(W - PAD - 16, 66) + MAP.scale_bar(PAD + 4, H + 10)
            + f'<text x="{PAD + 4}" y="{H + 36}" class="msrc">Boundaries: '
            f'geoBoundaries ADM2 &#183; source Open Data Rwanda / NISR '
            f'&#183; CC BY 4.0.</text></svg>')
    keys = "".join(f'<p class="key"><span class="sw" style="background:{c}">'
                   f'</span>{R.esc(l)}</p>'
                   for (_, _, c), l in zip(R.RX_STEPS, R.RX_LABELS))
    # 5.13: one row per feedback answer, its themes as a bit mask
    pg.dataset("fb", R.FB_ROWS, fac_rec, None,
               {"th": lambda r: sum(1 << i for i in r["_th"])})
    cfg = {"fb": [[nm, sm] for nm, sm, _ in R.FB_THEMES],
           "fbhide": sorted(R.FB_HIDE),
           "fbn": list(R.FB_NONE), "levels": R.LEVELS, "av": av, "k0": K0, "dflt": dflt, "exc": exc, "qc": qc,
           "ilevels": FL.ordered(r["_lvl"] for r in R.ITEMS),
           "map": {"W": W, "H": H, "PAD": PAD, "d": dun, "p": pun,
                   "tail": tail, "steps": R.RX_STEPS, "keys": keys}}
    return JS.replace("__CFG5__", R.json.dumps(cfg, separators=(",", ":")))


JS = r"""
var C5=__CFG5__;
// ---- rebuild what the payload leaves out (see build())
(function(){var IT=P.D.it,RX=P.D.rx,IC=IT.c,RC=RX.c,n=IT.n,j,i;
  if(!IC.pi){IC.pi=new Array(n);i=0;
    for(j=0;j<RX.n;j++)for(var k=0;k<RC.ni[j];k++)IC.pi[i++]=j;}
  IC.av=C5.av.split('').map(function(c){return c==='-'?null:+c;});
  if(!IT.f.length)for(i=0;i<n;i++)IT.f.push(RX.f[IC.pi[i]]);
  RX.m=[];for(j=0;j<RX.n;j++)RX.m.push([]);
  for(i=0;i<n;i++)RX.m[IC.pi[i]].push(IT.m[i]);
  // the prescription carries the name types of its items (Type of drug)
  if(RX.t&&IT.t)for(i=0;i<n;i++){var tv=IT.t[i];
    if(tv>=0&&RX.t[IC.pi[i]].indexOf(tv)<0)RX.t[IC.pi[i]].push(tv);}
  for(j=0;j<RX.n;j++)if(!RX.m[j].length)RX.m[j]=-1;
  if(!RC.nu){RC.nu=[];RC.fs=[];RC.any=[];
    for(j=0;j<RX.n;j++){RC.nu.push(0);RC.fs.push(RC.ni[j]>0?1:0);}
    for(i=0;i<n;i++){j=IC.pi[i];if(IC.av[i]===0)RC.nu[j]++;
      if(IC.av[i]!==1)RC.fs[j]=0;}
    for(j=0;j<RX.n;j++)RC.any.push(RC.nu[j]>0?1:0);}
  IC.mr=[];IC.cl0=[];IC.cls=[];IC.prog=[];IC.q=[];IC.cost=[];
  for(i=0;i<n;i++){var m=IT.m[i],e=C5.exc[i],d=C5.dflt[m],M=P.M[m],
      q=C5.qc[i];
    IC.mr.push(e?e[0]:d[0]);IC.cl0.push(C5.k0[e?e[1]:d[1]]);
    IC.cls.push(e?e[2]:M[2]);IC.prog.push(e?e[3]:M[1]);
    IC.q.push(q?q[0]:null);IC.cost.push(q?q[1]:null);}})();
// ---- charts.ring3d()
function _shade(h,f){h=h.replace('#','');var o='#';
  [0,2,4].forEach(function(i){var c=Math.floor(parseInt(h.substr(i,2),16)*f);
    c=Math.max(0,Math.min(255,c));o+=(c<16?'0':'')+c.toString(16);});return o;}
function ring3d(slices,colours,title,sub){
  var width=940,height=430,cx=320,cy=175,R=205,r_in=78,k=0.48,depth=40,
      min_deg=4.0,PI=Math.PI,steps=90;
  var tot=slices.reduce(function(a,s){return a+s[1];},0);
  if(!tot)return '<p class="fempty">Nothing to show.</p>';
  function pt(r,a,dy,ox,oy){return [cx+ox+r*Math.cos(a),cy+oy+dy+r*k*Math.sin(a)];}
  function arc(r,a0,a1,dy,ox,oy){var n=Math.max(2,Math.floor(steps*Math.abs(a1-a0)/PI)+2),
      o=[];for(var i=0;i<n;i++)o.push(pt(r,a0+(a1-a0)*i/(n-1),dy,ox,oy));return o;}
  function poly(pts,fill,stroke,tt){
    return '<polygon points="'+pts.map(function(p){return P.fx(p[0],1)+','
      +P.fx(p[1],1);}).join(' ')+'" fill="'+fill+'"'+(stroke?' stroke="'+stroke
      +'" stroke-width="1"':'')+'>'+(tt?'<title>'+P.esc(tt)+'</title>':'')
      +'</polygon>';}
  var spans=slices.map(function(s){return s[1]?Math.max(2*PI*s[1]/tot,
      min_deg*PI/180):0;}),big=0;
  slices.forEach(function(s,i){if(s[1]>slices[big][1])big=i;});
  spans[big]-=spans.reduce(function(a,b){return a+b;},0)-2*PI;
  var rest=0;spans.forEach(function(sp,i){if(i!==big)rest+=sp;});
  var a=PI/2-rest/2-spans[big],segs=[];
  slices.forEach(function(s,i){var span=spans[i],mid=a+span/2,
      pull=s[1]/tot<0.05?22:0;
    segs.push([s[0],s[1],a,a+span,colours[i%colours.length],
      pull*Math.cos(mid),pull*k*Math.sin(mid)]);a+=span;});
  var walls=[],inner=[],tops=[];
  segs.forEach(function(sg){var lab=sg[0],v=sg[1],a0=sg[2],a1=sg[3],col=sg[4],
      ox=sg[5],oy=sg[6],
      tip=lab+': '+P.int(v)+' of '+P.int(tot)+' ('+P.fx(v/tot*100,1)+'%)';
    [-1,0,1].forEach(function(m){var lo=2*PI*m,hi=2*PI*m+PI,
        b0=Math.max(a0,lo),b1=Math.min(a1,hi);
      if(b1>b0)walls.push(poly(arc(R,b0,b1,0,ox,oy).concat(arc(R,b1,b0,depth,ox,oy)),
        _shade(col,.72),_shade(col,.6),tip));});
    [-1,0,1].forEach(function(m){var lo=2*PI*m-PI,hi=2*PI*m,
        b0=Math.max(a0,lo),b1=Math.min(a1,hi);
      if(b1>b0)inner.push(poly(arc(r_in,b0,b1,0,ox,oy).concat(arc(r_in,b1,b0,depth,ox,oy)),
        _shade(col,.55),_shade(col,.45),tip));});
    if(ox||oy)[a0,a1].forEach(function(e){walls.push(poly([pt(r_in,e,0,ox,oy),
        pt(R,e,0,ox,oy),pt(R,e,depth,ox,oy),pt(r_in,e,depth,ox,oy)],
        _shade(col,.62),_shade(col,.5),tip));});
    tops.push(poly(arc(R,a0,a1,0,ox,oy).concat(arc(r_in,a1,a0,0,ox,oy)),col,
      _shade(col,.8),tip));});
  var o=['<svg class="chart" viewBox="0 0 '+width+' '+height+'" role="img" '
    +'aria-label="'+P.esc(title||'composition')+'">'];
  o.push('<ellipse cx="'+cx+'" cy="'+(cy+depth+10)+'" rx="'+(R+14)+'" ry="'
    +String((R+14)*k)+'" fill="#000" opacity=".35"/>');
  o=o.concat(inner,walls,tops);
  if(title)o.push('<text x="'+cx+'" y="'+P.fx(cy+depth+R*k+46,0)
    +'" text-anchor="middle" style="font-size:15px;font-weight:700;fill:#fff">'
    +P.esc(title)+'</text>');
  if(sub)o.push('<text x="'+cx+'" y="'+P.fx(cy+depth+R*k+64,0)
    +'" text-anchor="middle" class="cn">'+P.esc(sub)+'</text>');
  var kx=cx+R+70,ky=cy-70;
  segs.forEach(function(sg,i){var col=colours[i%colours.length],y=ky+i*62;
    o.push('<rect x="'+kx+'" y="'+y+'" width="16" height="16" rx="3" fill="'
      +col+'"/>');
    o.push('<text x="'+(kx+26)+'" y="'+(y+13)+'" style="font-size:14px;'
      +'font-weight:700;fill:#fff">'+P.esc(sg[0])+'</text>');
    o.push('<text x="'+(kx+26)+'" y="'+(y+36)+'" style="font-size:24px;'
      +'font-weight:700;fill:'+col+'">'+P.fx(sg[1]/tot*100,1)+'%</text>');
    o.push('<text x="'+(kx+120)+'" y="'+(y+36)+'" class="cn">'+P.int(sg[1])
      +' of '+P.int(tot)+'</text>');});
  return o.join('')+'</svg>';}
// ---- charts.donut1()
// pie, matching charts.pie(sort=False, w=560, h=330, r=104)
function pie(slices,cols,unit){var w=560,h=330,r=90,cx=w/2,cy=h/2+8,tot=0,
    f=function(v){return P.fx(v,1);},o=[],a=0,labels=[];
  slices.forEach(function(s){tot+=s[1];});
  if(!tot)return '<p class="fempty">Nothing to show.</p>';
  function pt(rr,an){var rad=(an-90)*Math.PI/180;
    return [cx+rr*Math.cos(rad),cy+rr*Math.sin(rad)];}
  o.push('<svg class="chart" viewBox="0 0 '+w+' '+h+'" role="img" '
    +'aria-label="composition">');
  slices.forEach(function(s,i){var span=360*s[1]/tot,p0=pt(r,a),
      p1=pt(r,a+span),col=cols[i%cols.length];
    o.push('<path d="M'+f(cx)+','+f(cy)+' L'+f(p0[0])+','+f(p0[1])+' A'+r+','
      +r+' 0 '+(span>180?1:0)+' 1 '+f(p1[0])+','+f(p1[1])+' Z" fill="'+col
      +'" stroke="#0e0e0d" stroke-width="2"><title>'+P.esc(s[0])+': '
      +P.int(s[1])+' of '+P.int(tot)+' ('+P.pc1(s[1]/tot)+')</title></path>');
    labels.push([s[0],s[1],a+span/2,col]);a+=span;});
  labels.forEach(function(L){var e=pt(r+4,L[2]),b=pt(r+30,L[2]),
      right=Math.cos((L[2]-90)*Math.PI/180)>=0,tx=b[0]+(right?46:-46),
      an=right?'start':'end',lx=tx+(right?6:-6);
    o.push('<path d="M'+f(e[0])+','+f(e[1])+' L'+f(b[0])+','+f(b[1])+' L'
      +f(tx)+','+f(b[1])+'" fill="none" stroke="'+L[3]+'" stroke-width="1.6"/>');
    o.push('<text x="'+P.fx(lx,0)+'" y="'+P.fx(b[1]-3,0)+'" class="ct" '
      +'text-anchor="'+an+'">'+P.esc(L[0])+'</text>');
    o.push('<text x="'+P.fx(lx,0)+'" y="'+P.fx(b[1]+13,0)+'" class="cs" '
      +'text-anchor="'+an+'">'+P.pc1(L[1]/tot)+' &#183; '+P.int(L[1])+' '
      +P.esc(unit)+'</text>');});
  o.push('<text x="'+P.fx(cx,0)+'" y="'+P.fx(cy+r+46,0)+'" class="cn" '
    +'text-anchor="middle">'+P.int(tot)+' '+P.esc(unit)+' in total</text>');
  return o.join('')+'</svg>';}
function donut1(hits,n,caption){var w=420,h=330,r=118,stroke=30,dec=2,
    colour='#c9544a',v=n?hits/n:0,cx='210.0',cy='161.0',cyn=161,
    circ=2*Math.PI*r,arc=v>0?Math.max(circ*v,3.0):0,
    o=['<svg class="chart" viewBox="0 0 '+w+' '+h+'" role="img" aria-label="'
      +P.fx(v*100,dec)+' percent">','<circle cx="'+cx+'" cy="'+cy+'" r="'+r
      +'" fill="none" stroke="'+CFG.rest+'" stroke-width="'+stroke+'"/>'];
  if(arc)o.push('<circle cx="'+cx+'" cy="'+cy+'" r="'+r+'" fill="none" stroke="'
    +colour+'" stroke-width="'+stroke+'" stroke-dasharray="'+P.fx(arc,2)+' '
    +P.fx(circ,2)+'" transform="rotate(-90 '+cx+' '+cy+')"><title>'
    +P.fx(v*100,dec)+'% - '+P.int(hits)+' of '+P.int(n)+'</title></circle>');
  o.push('<text x="'+cx+'" y="'+(cyn+6)+'.0" text-anchor="middle" '
    +'style="font-size:46px;font-weight:700;fill:#fff">'+P.fx(v*100,dec)
    +'%</text>');
  if(caption)o.push('<text x="'+cx+'" y="'+(cyn+32)+'.0" text-anchor="middle" '
    +'class="cn" style="font-size:13px">'+P.esc(caption)+'</text>');
  o.push('<text x="'+cx+'" y="'+(h-8)+'" text-anchor="middle" class="cn" '
    +'style="font-size:13px">'+P.int(hits)+' of '+P.int(n)
    +' prescribed medicines</text>');
  return o.join('')+'</svg>';}
// ---- FIG 5.2: two_view_map(), the outlines read once from the page
var MP=C5.map,PATHS=null;
function mapPaths(){if(PATHS)return PATHS;PATHS={d:[],p:[]};
  var b=document.getElementById('rxmap');
  if(b&&b.querySelectorAll)['d','p'].forEach(function(k){
    var ps=b.querySelectorAll('.mp-'+k+' path.dist');
    for(var i=0;i<ps.length;i++)PATHS[k].push(ps[i].getAttribute('d')||'');});
  return PATHS;}
function stepCol(v){if(v===null||v===undefined)return '#2c2c2a';
  for(var i=0;i<MP.steps.length;i++){var s=MP.steps[i];
    if(s[0]<=v&&v<s[1])return s[2];}return MP.steps[MP.steps.length-1][2];}
function pc(v){return v===null||v===undefined?'-':P.fx(v*100,1)+'%';}
function choro(units,dpaths,vals,title,sub,noun){
  var o=['<svg viewBox="0 0 '+MP.W+' '+(MP.H+54)+'" role="img" aria-label="'
    +P.esc(title)+'">','<text x="'+MP.PAD+'" y="18" class="mtitle">'
    +P.esc(title)+'</text>','<text x="'+MP.PAD+'" y="34" class="msub">'
    +P.esc(sub)+'</text>'];
  units.forEach(function(u,i){var x=vals[u[0]]||[0,0,null],
      where=u[1]&&u[1]!==u[0]?u[0]+', '+u[1]:u[0];
    o.push('<path class="dist pv-'+u[2]+'" d="'+(dpaths[i]||'')+'" fill="'
      +stepCol(x[2])+'"><title>'+P.esc(where)+' - '
      +(x[2]!==null?pc(x[2]):'fewer than '+P.MINN)+'\\n'+P.int(x[1])+' of '
      +P.int(x[0])+' '+P.esc(noun)+'</title></path>');});
  var big=units.length<=5;
  units.forEach(function(u){var x=vals[u[0]]||[0,0,null],pv='pv-'+u[2];
    o.push('<text class="dlabel '+pv+'" x="'+P.fx(u[3],0)+'" y="'+P.fx(u[4],0)
      +'"'+(big?' style="font-size:13px"':'')+'>'+P.esc(u[0])+'</text>');
    o.push('<text class="dlabel '+pv+'" x="'+P.fx(u[3],0)+'" y="'
      +P.fx(u[4]+(big?14:11),0)+'" style="font-size:'+(big?12:10.5)+'px">'
      +(x[2]!==null?pc(x[2]):'n/a')+'</text>');});
  return o.join('')+MP.tail;}
function rxMap(rs,keyd,keyp,ok,fac){
  function vals(kf){var o={};P.rateBy(rs,kf,ok,fac).forEach(function(r){
    o[r[0]]=[r[2],r[3],r[4]];});return o;}
  var dv=vals(keyd),pv=vals(keyp),pp=mapPaths(),
      tot=P.pct(P.cnt(rs,ok),rs.length),title='Complete prescriptions',
      sub=P.int(rs.length)+' prescriptions - overall '+pc(tot),
      noun='prescriptions complete';
  var opts=MP.p.map(function(u){return '<option value="'+u[2]+'">'
    +P.esc(u[0])+' &#183; '+(u[0] in pv?pc(pv[u[0]][2]):'n/a')+'</option>';})
    .join('');
  return '<div class="vmap maprow" id="rxmap"><div class="mapbox"><div class="mp-d">'
    +choro(MP.d,pp.d,dv,title+' - districts',sub,noun)+'</div><div class="mp-p">'
    +choro(MP.p,pp.p,pv,title+' - provinces',sub,noun)+'</div></div>'
    +'<div><div class="legend"><h4>Prescriptions complete</h4>'+MP.keys
    +'</div><div class="layers"><p class="lyrh">View</p>'
    +'<label class="lyr"><input type="radio" class="v-d" name="rxmap-v" checked>'
    +' Districts <span class="lyrn">'+MP.d.length+'</span></label>'
    +'<label class="lyr"><input type="radio" class="v-p" name="rxmap-v">'
    +' Provinces <span class="lyrn">'+MP.p.length+'</span></label>'
    +'<p class="lyrh" style="margin-top:14px">Filter &#183; province</p>'
    +'<select class="lyrsel" aria-label="Filter by province"><option value="">'
    +'All provinces</option>'+opts+'</select></div></div></div>';}
function putMap(h){var b=window.__PF_CAPTURE?null:document.getElementById('rxmap'),
    vp=null,sel=null;
  if(b){var r=b.querySelector('.v-p'),s=b.querySelector('.lyrsel');
    vp=r?r.checked:null;sel=s?s.value:null;}
  P.put('5|FIG_5.2',h);
  b=window.__PF_CAPTURE?null:document.getElementById('rxmap');
  if(b){if(vp){var r2=b.querySelector('.v-p');if(r2)r2.checked=true;}
    if(sel){var s2=b.querySelector('.lyrsel');if(s2)s2.value=sel;}}}
P.render(function(){mapPaths();
  var rx=P.rows('rx'),RC=P.D.rx.c,it=P.rows('it'),IC=P.D.it.c,
      FR=P.D.rx.f,FI=P.D.it.f,
      rfac=function(i){return FR[i];},ifac=function(i){return FI[i];},
      rF=function(i){return P.F[FR[i]];},iF=function(i){return P.F[FI[i]];},
      rlvl=function(i){return rF(i)[3];},rdist=function(i){return rF(i)[1];},
      rprov=function(i){return rF(i)[2];},
      rmg=function(i){var m=rF(i)[4];return m==='(not recorded)'?'':m;},
      ilvl=function(i){return iF(i)[3];},iprov=function(i){return iF(i)[2];},
      idist=function(i){return iF(i)[1];},
      inm=function(i){return P.med('it',i)[0];},
      any=function(i){return RC.any[i];},
      un=function(i){return IC.av[i]===0;},ok=function(i){return IC.av[i]===1;};
  var RX=rx.filter(function(i){return RC.ni[i]>0;}),
      ANS=it.filter(function(i){return IC.av[i]!==null;}),
      MISS=ANS.filter(un),UNS=it.filter(un);
  // FIG 5.1
  var c1=P.cnt(RX,function(i){return RC.fs[i];}),
      c2=P.cnt(RX,function(i){return RC.any[i]&&RC.nu[i]<RC.ni[i];}),
      c3=P.cnt(RX,function(i){return RC.ni[i]&&RC.nu[i]===RC.ni[i];});
  P.put('5|FIG_5.1',ring3d([['Complete',c1],['Partially served',c2],
    ['Nothing served',c3]],['#199e70','#eda100','#c9544a'],
    pc(P.pct(c1,RX.length))+' of prescriptions complete',
    P.int(RX.length)+' prescriptions'));
  // FIG 5.2
  putMap(rxMap(RX,rdist,rprov,function(i){return RC.fs[i];},rfac));
  // FIG 5.3, 5.4
  var H1='with a missing item',B1='prescriptions';
  // FIG 5.3 by programme: prescriptions holding one of its medicines, and
  // those where one of ITS medicines was unavailable (as _rx_prog_rows())
  (function(){var g={};ANS.forEach(function(i){var p=IC.prog[i],
      e=g[p]||(g[p]={n:{},m:{},f:{}});e.n[IC.pi[i]]=1;e.f[ifac(i)]=1;
      if(un(i))e.m[IC.pi[i]]=1;});
    var rows=Object.keys(g).sort(P.pyStrCmp).map(function(p){
      var n=Object.keys(g[p].n).length,h=Object.keys(g[p].m).length;
      return [p,Object.keys(g[p].f).length,n,h,n>=P.MINN?h/n:null];});
    rows=P.sort(rows,function(a,b){return (a[4]===null)-(b[4]===null)
      ||(a[4]===null?0:a[4]-b[4]);});
    P.put('5|FIG_5.3',P.bars(P.worstFirst(rows),H1,B1));})();
  P.put('5|FIG_5.4',P.bars(P.barRows(P.rateBy(RX,rlvl,any,rfac,C5.levels,
    null,false)),H1,B1));
  // FIG 5.5-5.7
  P.put('5|FIG_5.5',donut1(MISS.length,ANS.length,'of medicines unserved'));
  function miss(kf,mn){return P.worstFirst(P.rateBy(ANS,kf,un,ifac,null,mn)
    .filter(function(x){return x[3];}));}
  var H2='unserved',B2='medicines prescribed';
  // the drug-class chart was removed on request (2026-10-19); the
  // programme chart is FIG 5.6
  P.put('5|FIG_5.6',P.bars(miss(function(i){return IC.prog[i];},1),H2,B2));
  // TABLE 5A
  var bm=P.by(MISS,inm),presc={},rows=[];
  ANS.forEach(function(i){var m=inm(i);presc[m]=(presc[m]||0)+1;});
  bm.keys.forEach(function(m){var rs=bm.g[m],d={};
    rs.forEach(function(i){d[idist(i)]=1;});
    rows.push([m,IC.cls[rs[0]],IC.prog[rs[0]],presc[m],rs.length,
      P.R(P.pct(rs.length,presc[m])),P.uniq(rs,function(i){return IC.pi[i];}),
      P.uniq(rs,ifac),Object.keys(d).sort(P.pyStrCmp).join(', '),
      P.pyRound(P.sum(rs,IC.q)),P.pyRound(P.sum(rs,IC.cost))]);});
  rows=P.sort(rows,function(a,b){return b[4]-a[4]||b[10]-a[10];});
  P.put('5|TABLE_5A',P.table(['Medicine','Drug class','Programme',
    'Times prescribed','Times unserved','UNSERVED %','Prescriptions affected',
    'Facilities','Districts','Qty unserved','Value unserved (RWF)'],rows));
  // 5.1-4.5
  var nP=rx.length,nU=P.cnt(rx,any),
      nA=P.cnt(rx,function(i){return RC.ni[i]&&RC.nu[i]===RC.ni[i];}),
      nM=P.cnt(rx,function(i){return RC.any[i]&&!(RC.ni[i]&&RC.nu[i]===RC.ni[i]);});
  // 5.1-4.5 is off the dashboard (HTML_DROP, 2026-10-20)
  // 5.7 is off the dashboard (HTML_DROP, 2026-10-20)
  // 5.8
  var bl=P.by(it,ilvl),bp=P.by(rx,rlvl);
  P.put('5|5.8',P.table(['Level of care','Facilities','Prescriptions',
    'Prescriptions not fully served','%','Items','Unserved items',
    'Unserved %'],C5.ilevels.filter(function(l){return l in bl.g;})
    .map(function(l){var ri=bl.g[l],rp=bp.g[l]||[],pu=P.cnt(rp,any),
        iu=P.cnt(ri,un);
      return [l,P.uniq(ri,ifac),rp.length,pu,P.R(P.pct(pu,rp.length)),
        ri.length,iu,P.R(P.pct(iu,ri.length))];})));
  // 5.9: grouped by the readable medicine name, not the form's code
  // (user request 2026-10-20), as rx_med_name() in the report
  var mr=function(i){return P.med('it',i)[0];},bu=P.by(UNS,mr),ba=P.by(it,mr);
  rows=bu.keys.slice().sort(P.pyStrCmp).map(function(m){var u=bu.g[m],
      a=(ba.g[m]||[]).length;
    return [m,u.length,a,P.R(P.pct(u.length,a)),P.uniq(u,ifac),
      P.pyRound(P.sum(u,IC.q))];});
  rows=P.sort(rows,function(a,b){return b[1]-a[1]||(b[3].r||0)-(a[3].r||0);})
    .slice(0,15);
  P.put('5|5.9',P.table(['Medicine','Unserved','Times prescribed',
    'Unserved %','Facilities','Quantity'],rows));
  // 5.10-5.12
  var RH=['Group','Facilities','Observations','Numerator','RATE'];
  function rt(t){return t.map(function(r){return [r[0],r[1],r[2],r[3],
    P.R(r[4])];});}
  P.put('5|5.10',P.table(RH,rt(P.rateBy(ANS,iprov,ok,ifac))));
  P.put('5|5.11',P.table(RH,rt(P.rateBy(ANS,idist,ok,ifac).slice(0,10))));
  // 5.12 is off the dashboard (HTML_DROP, 2026-10-20)
  // 5.13: feedback themes, facilities once per theme. Facility filters
  // only - an answer is not about one medicine, so the medicine and name
  // type filters leave it alone.
  var FD=P.D.fb,fbi=[];
  for(var q=0;q<FD.n;q++){var ff=FD.f[q];if(ff<0||P._okF[ff])fbi.push(q);}
  var fth=C5.fb.map(function(){return {};}),fall={},fthm={};
  fbi.forEach(function(q){var nm=P.fac('fb',q)[0],m=FD.c.th[q];fall[nm]=1;
    for(var j=0;j<C5.fb.length;j++)if(m&(1<<j)){fth[j][nm]=1;fthm[nm]=1;}});
  var fr=[];C5.fb.forEach(function(t,j){var n=Object.keys(fth[j]).length;
    if(n&&C5.fbhide.indexOf(t[0])<0)fr.push([t[0],n,t[1]]);});
  fr=P.sort(fr,function(a,b){return b[1]-a[1];});
  var nall=Object.keys(fall).length,
      nno=Object.keys(fall).filter(function(k){return !fthm[k];}).length;
  P.put('5|5.13',nall?P.table(['Theme','Facilities','What was reported'],
    [['Facilities that gave feedback',nall,'']].concat(fr)):P.EMPTY);
  // scorecards pc5-0..3, as page_cards() writes them
  function card(k,v,l,d){var a=['v','l','s'],x=[v,l,d];
    for(var j=0;j<3;j++){var e=document.getElementById('pc5-'+k+'-'+a[j]);
      if(e)e.textContent=x[j];}}
  var NONE='nothing matches these filters';
  card(0,RX.length?pc(c1/RX.length):'-','of prescriptions complete',
    RX.length?'high by design - prescribers write what the pharmacy holds; '
      +'not a stock-availability rate':NONE);
  card(1,ANS.length?P.fx(MISS.length/ANS.length*100,1)+'%':'-',
    'of prescribed medicines unserved',ANS.length?P.int(MISS.length)+' of '
      +P.int(ANS.length)+' items - understated for the same reason; use the '
      +'stock pages for availability':NONE);
  var nrh=RX.filter(function(i){return rlvl(i)==='National referral hospital';});
  card(2,nrh.length?pc(P.cnt(nrh,any)/nrh.length):'-',
    'national referral prescriptions with a gap',nrh.length?nrh.length
      +' prescriptions'
      :'no national referral prescriptions match these filters');
  var rd=P.rateBy(RX,rdist,function(i){return RC.fs[i];},rfac)
    .filter(function(x){return x[4]!==null;});
  if(rd.length)card(3,pc(rd[0][4]),rd[0][0]+' - lowest district',
    'share of prescriptions complete');
  else card(3,'-','lowest district','no district with '+P.MINN
    +' or more prescriptions');
});
"""
