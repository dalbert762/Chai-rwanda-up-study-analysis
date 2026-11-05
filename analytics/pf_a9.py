"""Page filter for Analysis 9 - community-level (CHW) medicine availability.

build(pg, R) registers the CHW rows (and, for 9.4, the facility stock rows
of the eight shared products) and returns the script that recomputes
FIG 9.1 and tables 9.1-9.6 from the filtered rows.

Filter records
  facility  the CHW's parent health centre, as the stock tool records it
            (name, district, province, level group, management); the
            management type comes from the stock tool when that facility
            was also assessed there, "(not recorded)" otherwise.
  medicine  the CHW product's short name. A product with a facility
            comparator (R.SHARED) also carries the facility product's name,
            and each facility row carries the CHW short name, so typing
            either name selects both sides of the 9.4 pairing. Programme:
            the CHW form records none; the ten malaria products (ACTs, DHAP,
            pyronaridine/artesunate, RDT) take the facility tool's
            "Malaria", ORS / zinc / amoxicillin have no programme.
  sub       one row per CHW submission (9.1's CHW and health-centre counts),
            medicines = every product in that CHW's basket.
"""


def build(pg, R):
    C, M, s = R.C, R.M, R.s
    shared = list(R.SHARED.items())
    fac_names = set(R.SHARED.values())
    frows = [r for r in R.ANS if s(r[M.MED]) in fac_names]

    # programme / class of each facility comparator
    fmeta = {}
    for r in frows:
        fmeta.setdefault(s(r[M.MED]), (s(r[M.PROGRAM]), R._class_label(r)))
    fam_cls = {}
    for sh, fn in shared:
        if fn in fmeta:
            fam_cls.setdefault(C.family(sh), fmeta[fn][1])
    malaria = {k for _, k in C.FAMILY} | {"Malaria RDT"}
    prog_malaria = next((p for p, _ in fmeta.values() if p), "Malaria")

    def chw_med_rec(sh):
        fn = R.SHARED.get(sh)
        if fn in fmeta:
            prg, cls = fmeta[fn]
        else:
            fam = C.family(sh)
            prg = prog_malaria if (fam in malaria or sh in malaria) \
                else "(none)"
            cls = fam_cls.get(fam, fam)
        rec = [sh, prg, cls]
        return [rec, [fn, prg, cls]] if fn else rec

    mgmt = {}
    for r in R.MED_ALL:
        mgmt.setdefault(s(r["_p"][M.FAC]), R.mmgt(r))

    def chw_fac(p):
        f = C.s(p[C.FAC])
        return [f, C.s(p[C.DIST]), C.s(p[C.PROV]), R.FL.group(p[C.LEVEL]),
                mgmt.get(f, "(not recorded)")]

    def chw_key(p):
        return C.s(p[C.FAC]) + "\u0001" + C.s(p[C.CHW])

    def av(r):
        v = r[C.AVAIL]
        return None if C.blank(v) else (1 if C.yes(v) else
                                        (0 if C.no(v) else -1))

    # CHW rows first, so the CHW form's facility records are the ones kept
    pg.dataset("c", C.MEDROWS, lambda r: chw_fac(r["_p"]),
               lambda r: chw_med_rec(r["_short"]), {
                   "mg": lambda r: C.yes(r[C.MANAGED]),
                   "av": av,
                   "sh": lambda r: r["_short"],
                   # the least-available card groups on the form's own
                   # name cut to 30 characters, not the short name
                   "mn": lambda r: C.s(r[C.MED])[:30],
                   "k": lambda r: chw_key(r["_p"])})

    basket = {}
    for r in C.MEDROWS:
        basket.setdefault(id(r["_p"]), [])
        if r["_short"] not in basket[id(r["_p"])]:
            basket[id(r["_p"])].append(r["_short"])

    def sub_meds(p):
        b = basket.get(id(p))
        if not b:
            return None
        out = []
        for sh in b:
            rec = chw_med_rec(sh)
            out.extend(rec if isinstance(rec[0], list) else [rec])
        return out

    pg.dataset("s", C.SUB, chw_fac, sub_meds, {"k": chw_key})

    short_of = {fn: sh for sh, fn in shared}
    pg.dataset("f", frows, R.pf_fac,
               lambda r: [R.pf_med(r), chw_med_rec(short_of[s(r[M.MED])])[0]],
               {"av": lambda r: M.yes(r[M.AVAIL]),
                "nm": lambda r: s(r[M.MED])})
    return (JS.replace("__SHARED__", R.json.dumps(shared))
              .replace("__STEPS__", R.json.dumps(
                  [list(x) for x in R.CHW_STEPS])))


JS = r"""
var SHARED=__SHARED__,STEPS=__STEPS__;
// FIG 9.1: recolour the two-view map in place (paths are in the page)
function stepCol(v){if(v===null)return '#2c2c2a';
  for(var i=0;i<STEPS.length;i++){var s=STEPS[i];
    if(v>=s[0]&&v<s[1])return s[2];}
  return STEPS[STEPS.length-1][2];}
function paintMap(dv,pv,overall,n){
  var root=window.__PF_CAPTURE?null:document.getElementById('chwmap');
  if(!root)return;
  [['d',dv],['p',pv]].forEach(function(L){var box=root.querySelector('.mp-'+L[0]);
    if(!box)return;var vals=L[1];
    box.querySelectorAll('path[data-n]').forEach(function(el){
      var g=vals[el.getAttribute('data-n')],v=g&&g[0]>0?g[1]/g[0]:null;
      el.setAttribute('fill',stepCol(v));var t=el.querySelector('title');
      if(t){if(!el.dataset.lab)el.dataset.lab=t.textContent.split(' - ')[0];
        t.textContent=el.dataset.lab+' - '+(v===null?'no CHW records'
          :P.pc1(v))+'\n'+P.int(g?g[1]:0)+' of '+P.int(g?g[0]:0)
          +' managed products available';}});
    box.querySelectorAll('text[data-nv]').forEach(function(el){
      var g=vals[el.getAttribute('data-nv')];
      el.textContent=g&&g[0]>0?P.pc1(g[1]/g[0]):'n/a';});
    var sb=box.querySelector('.msub');
    if(sb)sb.textContent=P.int(n)+' managed products - overall '+P.pc1(overall);});
  var sel=root.querySelector('.lyrsel');
  if(sel)Array.prototype.forEach.call(sel.options,function(op){if(!op.value)return;
    var nm=op.textContent.split(' · ')[0],g=pv[nm],
        v=g&&g[0]>0?g[1]/g[0]:null;
    op.textContent=nm+' · '+(v===null?'n/a':P.pc1(v));});}
// FIG 9.2: charts.lollipop, as pf_a7 draws it
var fx1=function(v){return P.fx(v,1);};
function lollipop(rows,hit,base){
  if(!rows.length)return '<p class="fempty">Nothing to show.</p>';
  var lw=210,rh=34,pt=30,width=940,colour='#4f9be8',
      h=pt+rows.length*rh+26,pw=width-lw-90,x0=lw,
      o=['<svg class="chart" viewBox="0 0 '+width+' '+h+'" role="img">'];
  [25,50,75,100].forEach(function(t){var x=x0+pw*t/100;
    o.push('<line class="gl" x1="'+fx1(x)+'" y1="'+(pt-10)+'" x2="'+fx1(x)
      +'" y2="'+(h-18)+'"/>');
    o.push('<text x="'+fx1(x)+'" y="'+(h-4)+'" class="cn" '
      +'text-anchor="middle">'+t+'%</text>');});
  rows.forEach(function(r,i){var lab=r[0],n=r[1],hv=r[2],v=n?hv/n:0,
      y=pt+i*rh+rh/2,xv=x0+pw*Math.min(v,1);
    o.push('<text x="'+(x0-14)+'" y="'+fx1(y+4.5)+'" class="ct" '
      +'text-anchor="end">'+P.esc(lab)+'</text>');
    o.push('<line x1="'+x0+'" y1="'+fx1(y)+'" x2="'+fx1(xv-7)+'" y2="'
      +fx1(y)+'" stroke="'+colour+'" stroke-width="2"/>');
    o.push('<circle cx="'+fx1(xv)+'" cy="'+fx1(y)+'" r="7" '
      +'fill="var(--plane,#000)" stroke="'+colour+'" stroke-width="3">'
      +'<title>'+P.esc(lab)+': '+P.pc1(v)+' '+P.esc(hit)+' - '+P.int(hv)
      +' of '+P.int(n)+' '+P.esc(base)+'</title></circle>');
    o.push('<text x="'+fx1(xv+14)+'" y="'+fx1(y+4.5)+'" class="cv">'
      +P.pc1(v)+'</text>');
    o.push('<text x="'+fx1(xv+66)+'" y="'+fx1(y+4.5)+'" class="cn">n='
      +P.int(n)+'</text>');});
  o.push('<line x1="'+x0+'" y1="'+(pt-12)+'" x2="'+x0+'" y2="'+(h-22)
    +'" stroke="#e8e8e2" stroke-width="5"/>');
  return o.join('')+'</svg>';}
P.render(function(){
  var C=P.D.c.c,ix=P.rows('c'),
      fac=function(i){return P.fac('c',i)[0];},
      key=function(i){return C.k[i];},
      cm=ix.filter(function(i){return C.mg[i];}),
      cans=cm.filter(function(i){return C.av[i]!==null;}),
      yes=function(i){return C.av[i]===1;},
      cav=cans.filter(yes);
  // FIG 9.1 map, FIG 9.2 province bars
  var prov=function(i){return P.fac('c',i)[2];},
      lvl=function(i){return P.fac('c',i)[3];};
  function geo(keyfn){var o={};
    cans.forEach(function(i){var k=keyfn(i),g=o[k]||(o[k]=[0,0]);g[0]++;
      if(yes(i))g[1]++;});return o;}
  paintMap(geo(function(i){return P.fac('c',i)[1];}),geo(prov),
    P.pct(cav.length,cans.length),cans.length);
  // FIG 9.2: lollipop per product, highest first, ties by name
  var bp=P.by(cans,function(i){return C.sh[i];}),lp=[];
  bp.keys.slice().sort(P.pyStrCmp).forEach(function(m){var rs=bp.g[m];
    if(rs.length>=P.MINN)lp.push([m,rs.length,P.cnt(rs,yes)]);});
  lp=P.sort(lp,function(a,b){return -(a[2]/a[1])+(b[2]/b[1]);});
  P.put('9|FIG_9.2',lollipop(lp,'available','managed products'));
  // 9.1
  var sx=P.rows('s'),SK=P.D.s.c.k;
  // 9.1 is off the dashboard (HTML_DROP, 2026-10-20)
  // 9.2: per product, lowest availability first
  var bs=P.by(cm,function(i){return C.sh[i];}),r2=[];
  bs.keys.slice().sort(P.pyStrCmp).forEach(function(m){var rs=bs.g[m],
      a=P.cnt(rs,yes);
    r2.push([m,rs.length,a,P.R(P.pct(a,rs.length)),P.uniq(rs,key)]);});
  r2=P.sort(r2,function(a,b){return (a[3].r||0)-(b[3].r||0);});
  // 9.2 is off the dashboard (HTML_DROP, 2026-10-20)
  // 9.3: complete basket per CHW
  var bc=P.by(cm,key),nc=0,n1=0,n3=0;
  bc.keys.forEach(function(k){var rs=bc.g[k],
      no=P.cnt(rs,function(i){return C.av[i]===0;});
    if(rs.every(yes))nc++;if(no===1)n1++;if(no>=3)n3++;});
  // 9.3 is off the dashboard (HTML_DROP, 2026-10-20)
  // 9.4: community against facility on the shared products
  var F=P.D.f.c,fx=P.rows('f'),r4=[];
  function cmpRow(lab,c,f){var ca=P.cnt(c,yes),
      fa=P.cnt(f,function(i){return F.av[i];}),
      cp=P.pct(ca,c.length),fp=P.pct(fa,f.length),
      d=(cp===null&&fp===null)?0:P.R((cp||0)-(fp||0));
    return [lab,c.length,P.R(cp),f.length,P.R(fp),d];}
  var shs={},fns={};
  SHARED.forEach(function(p){shs[p[0]]=1;fns[p[1]]=1;
    var c=cm.filter(function(i){return C.sh[i]===p[0];}),
        f=fx.filter(function(i){return F.nm[i]===p[1];});
    if(c.length||f.length||!P.active())r4.push(cmpRow(p[0],c,f));});
  r4.push(cmpRow('ALL EIGHT SHARED PRODUCTS',
    cm.filter(function(i){return shs[C.sh[i]];}),
    fx.filter(function(i){return fns[F.nm[i]];})));
  P.put('9|9.4',r4.length>1?P.table(['Product','CHW observations',
    'CHW available %','Facility observations','Facility available %',
    'Difference (percentage points)'],r4):P.EMPTY);
  // 9.5: by district, lowest first
  var dist=function(i){return P.fac('c',i)[1];},bd=P.by(cm,dist),r5=[];
  bd.keys.slice().sort(P.pyStrCmp).forEach(function(d){var rs=bd.g[d],
      a=P.cnt(rs,yes);
    r5.push([d,P.uniq(rs,key),rs.length,a,P.R(P.pct(a,rs.length))]);});
  r5=P.sort(r5,function(a,b){return (a[4].r||0)-(b[4].r||0);});
  // 9.5 is off the dashboard (HTML_DROP, 2026-10-20)
  // 9.6: province, then level of care
  function rr(t){return t.map(function(r){return [r[0],r[1],r[2],r[3],
    P.R(r[4])];});}
  // 9.6 is off the dashboard (HTML_DROP, 2026-10-20)
  // scorecards pc9-0..2, as page_cards() writes them
  function card(k,v,l,d){var a=['v','l','s'],x=[v,l,d];
    for(var j=0;j<3;j++){var e=document.getElementById('pc9-'+k+'-'+a[j]);
      if(e)e.textContent=x[j];}}
  var nchw=P.uniq(sx,function(i){return SK[i];});
  if(cans.length)card(0,P.pc1(cav.length/cans.length),
    'CHW medicine availability',nchw+' CHWs, products they manage');
  else card(0,'-','CHW medicine availability',
    'no managed products match these filters');
  var bm=P.by(cans,function(i){return C.mn[i];}),low=null;
  bm.keys.forEach(function(m){var rs=bm.g[m];if(rs.length<P.MINN)return;
    var v=P.cnt(rs,yes)/rs.length;
    if(low===null||v<low[0]||(v===low[0]&&P.pyStrCmp(m,low[1])<0))
      low=[v,m];});
  if(low)card(1,P.pc1(low[0]),low[1]+' - least available','at CHW level');
  else card(1,'-','least available','no medicine with '+P.MINN
    +' or more managed observations');
  var nhc=P.uniq(sx,function(i){return P.fac('s',i)[0];});
  card(2,sx.length?String(nhc):'-','health centres supervising the CHWs',
    sx.length?'every CHW reports to a health centre'
    :'no CHWs match these filters');
});
"""
