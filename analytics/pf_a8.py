"""Page filter for Analysis 8 - financial value of unserved prescriptions
(displayed as tab 7).

Two halves, three datasets:
  un   - unserved prescription lines (R.UNSERVED): tables 8.1-7.3, 8.4,
         8.5, 8.6. Facility from the prescription submission (_sub), the
         management type borrowed from the stock tool by name (rx_mgmt);
         medicine [rx_med_name (readable, so the name filter finds it - the
         tables keep the report's own _med label), rx_prog, rx_class].
  pres - prescriptions not fully served (R._pres_un), for the "Prescriptions
         not fully served" count and the average per affected prescription.
         Its medicines are those of its unserved lines, so a medicine filter
         keeps the prescriptions in which that medicine went unserved.
  lo   - managed stock rows with at least one stock-out day (_wdays > 0),
         for the estimated lost sales, tables 8.7-8.12. The estimate itself
         (_loss_row) is recomputed in the browser from its raw inputs - price,
         summed consumption, days out - so the floating-point arithmetic is
         the report's, operation for operation.
Grouping keys are carried as their own columns rather than read from the
shared facility/medicine records, which are keyed by name and keep the first
spelling seen.
"""


def build(pg, R):
    M, P, s = R.M, R.P, R.s

    def rx_fac(r):
        sub = r["_sub"]
        return [s(sub[P.FACIL]), s(sub[P.DISTR]), s(sub[P.PROV]),
                R.rx_level(r), R.rx_mgmt(r) or "(not recorded)"]

    def rx_med(r):
        return [R.rx_med_name(r), R.rx_prog(r), R.rx_class(r)]

    def cost(r):
        return P.num(r.get(P.COST_UNSERVED))

    pg.dataset("un", R.UNSERVED, rx_fac, rx_med, {
        "c": cost,
        "pr": lambda r: not P.blank(r.get(P.UNIT_PRICE)),
        "lvl": lambda r: r["_lvl"],
        "prov": lambda r: s(r["_sub"][P.PROV]),
        "cls": lambda r: r["_class"],
        "med": lambda r: s(r["_med"])})

    lines = {}
    for r in R.UNSERVED:
        if r.get("_pres") is not None:
            lines.setdefault(id(r["_pres"]), []).append(r)
    pg.dataset("pres", R._pres_un, rx_fac,
               lambda p: [rx_med(i) for i in lines.get(id(p), [])] or None)

    def cons(r):
        c = [M.num(r[M.CONS[m]]) for m in R.MONTHS]
        return None if all(x is None for x in c) else sum(x or 0 for x in c)

    lo = [r for r in R.MG if R._wdays(r) > 0]
    pg.dataset("lo", lo, R.pf_fac, R.pf_med, {
        "wd": R._wdays,
        "pr": lambda r: M.num(r[M.PRICE]),
        "cs": cons,
        "lvl": lambda r: r["_lvl"],
        "cls": R._class_label,
        "med": lambda r: s(r[M.MED]),
        "prov": lambda r: s(r["_p"][M.PROV]),
        "dist": lambda r: s(r["_p"][M.DIST])})

    order = R.FL.ordered(r["_lvl"] for r in R.ITEMS)
    return (JS.replace("__LEVELS__", R.json.dumps(R.LEVELS))
              .replace("__RXLV__", R.json.dumps(order))
              .replace("__MARKUP__", repr(R.MARKUP)))


JS = r"""
var LEVELS=__LEVELS__,RXLV=__RXLV__,MARKUP=__MARKUP__;
function cmpStr(a,b){return P.pyStrCmp(String(a),String(b));}
P.render(function(){
  var C=P.D.un.c,ix=P.rows('un'),
      cost=function(i){return C.c[i]||0;},
      fac=function(i){return P.fac('un',i)[0];};
  var val=0;ix.forEach(function(i){val+=cost(i);});
  var priced=ix.filter(function(i){return C.pr[i];}),
      npres=P.rows('pres').length;
  // 8.1-7.3 is off the dashboard (HTML_DROP, 2026-10-20); val, priced and
  // npres are still used by the scorecards below
  // 8.4 Pareto: medicines by value, cumulative share of the total
  var bm=P.by(ix,function(i){return C.med[i];}),bv=[];
  bm.keys.slice().sort(cmpStr).forEach(function(m){var rs=bm.g[m],t=0;
    rs.forEach(function(i){t+=cost(i);});bv.push([m,rs.length,P.pyRound(t)]);});
  bv=P.sort(bv,function(a,b){return b[2]-a[2];});
  var cum=0,par=bv.map(function(x){cum+=x[2];
    return [x[0],x[1],x[2],P.R(P.pct(x[2],val)),P.R(P.pct(cum,val))];});
  P.put('8|8.4',P.table(['Medicine','Unserved lines','Value (RWF)',
    '% of total','Cumulative %'],par.slice(0,15)));
  // 8.5 level of care, then province
  var rows=[];
  function grp(keyfn,keys,dim){var b=P.by(ix,keyfn);
    keys.forEach(function(k){var rs=b.g[k]||[],t=0;
      rs.forEach(function(i){t+=cost(i);});
      rows.push([k,dim,rs.length,P.pyRound(t),P.R(P.pct(t,val))]);});}
  var bl=P.by(ix,function(i){return C.lvl[i];});
  grp(function(i){return C.lvl[i];},
      RXLV.filter(function(l){return l in bl.g;}),'Level of care');
  grp(function(i){return C.prov[i];},
      P.by(ix,function(i){return C.prov[i];}).keys.slice().sort(cmpStr),
      'Province');
  P.put('8|8.5',P.table(['Group','Dimension','Unserved lines','Value (RWF)',
    '% of total'],rows));
  // 8.6 vol_by() over priced unserved lines, by recorded drug class
  var pos=ix.filter(function(i){return cost(i)>0;}),grand=0;
  pos.forEach(function(i){grand+=cost(i);});
  var bc=P.by(pos,function(i){return C.cls[i];}),vb=[];
  bc.keys.slice().sort(cmpStr).forEach(function(k){var rs=bc.g[k],t=0;
    rs.forEach(function(i){t+=cost(i);});
    vb.push([k,P.uniq(rs,fac),rs.length,P.pyRound(t),
      P.median(rs.map(function(i){return C.c[i];})),P.R(P.pct(t,grand))]);});
  vb=P.sort(vb,function(a,b){return b[3]-a[3];});
  P.put('8|8.6',P.table(['Group','Facilities','Observations','Total','Median',
    'SHARE OF TOTAL'],vb.slice(0,10)));

  // ---- estimated lost sales: _loss_row() per stock-out pair
  var L=P.D.lo.c,lx=P.rows('lo'),loss=[],excl={},tot=0,top=null;
  lx.forEach(function(i){var out=Math.min(L.wd[i],92),pr=L.pr[i]||0,why=null;
    if(pr<=0)why='no price';
    else if(L.cs[i]===null||L.cs[i]===undefined)why='no consumption recorded';
    else if(out>=92)why='out all 92 days - no demand signal';
    if(why){excl[why]=(excl[why]||0)+1;return;}
    var u=L.cs[i]/(92-out)*out,x=[i,u,u*pr];loss.push(x);
    if(!top||x[2]>top[2])top=x;});
  var base0=0;loss.forEach(function(x){base0+=x[2];});tot=base0*(1+MARKUP);
  var lfac=function(i){return P.fac('lo',i)[0];};
  function lossBy(keyfn,keys){var g={},ks=[];
    loss.forEach(function(x){var k=keyfn(x[0]);
      if(!(k in g)){g[k]=[];ks.push(k);}g[k].push(x);});
    var out=[];(keys||ks.slice().sort(cmpStr)).forEach(function(k){
      var rs=g[k];if(!rs||!rs.length)return;var b=0,u=0,f={};
      rs.forEach(function(x){b+=x[2];u+=x[1];f[lfac(x[0])]=1;});
      out.push([k,Object.keys(f).length,rs.length,P.pyRound(u),P.pyRound(b),
        P.pyRound(b*MARKUP),P.pyRound(b*(1+MARKUP)),
        P.R(P.pct(b*(1+MARKUP),tot))]);});
    return out;}
  function byValue(t){return P.sort(t,function(a,b){return b[6]-a[6];});}
  var uall=0,fall={};loss.forEach(function(x){uall+=x[1];fall[lfac(x[0])]=1;});
  P.put('8|8.7',P.table(['Measure','Value'],[
    ['Managed pairs with at least one stock-out day, Jun-Aug',lx.length],
    ['  estimated (priced, with in-stock consumption)',loss.length],
    ['  excluded - no price (programme products supplied free; insulin, '
     +'carbetocin)',excl['no price']||0],
    ['  excluded - out all 92 days, no demand signal',
     excl['out all 92 days - no demand signal']||0],
    ['  excluded - no consumption recorded',
     excl['no consumption recorded']||0],
    ['Estimated units lost',P.pyRound(uall)],
    ['Value at tariff price (RWF)',P.pyRound(tot/(1+MARKUP))],
    ['20% markup - the facility\'s own margin lost (RWF)',
     P.pyRound(tot/(1+MARKUP)*MARKUP)],
    ['LOST SALES incl. 20% markup (RWF)',P.pyRound(tot)],
    ['Facilities with an estimated loss',Object.keys(fall).length],
    ['Largest single facility x medicine - '+(top?L.med[top[0]]+' at '
     +lfac(top[0])+', out '+P.fx(Math.min(L.wd[top[0]],92),0)+' days':''),
     top?P.pyRound(top[2]*(1+MARKUP)):null],
    ['  its share of all lost sales',
     P.R(top?P.pct(top[2]*(1+MARKUP),tot):null)]]));
  // ---- the page's scorecards follow the filters too
  function card(id,v,d){var e=document.getElementById(id+'-v');
    if(e)e.textContent=v;e=document.getElementById(id+'-s');
    if(e&&d!==undefined)e.textContent=d;}
  card('pc8-val',P.int(P.pyRound(val))+' RWF',P.int(ix.length)
    +' unserved items, the form\'s own cost');
  card('pc8-loss',P.fx(tot/1e6,1)+'M RWF');
  card('pc8-top',top&&tot?P.pc1(top[2]*(1+MARKUP)/tot):'-',
    top?String(L.med[top[0]]).slice(0,28)+' at '+lfac(top[0])
       :'no estimated loss under these filters');
  var FH=['Group','Facilities','Pairs out','Estimated units lost',
    'Value at tariff price (RWF)','20% markup (RWF)',
    'LOST SALES incl. markup (RWF)','SHARE OF TOTAL'];
  P.put('8|8.8',P.table(FH,lossBy(function(i){return L.lvl[i];},LEVELS)));
  P.put('8|8.9',P.table(FH,byValue(lossBy(function(i){return L.cls[i];}))));
  P.put('8|8.10',P.table(FH,byValue(lossBy(function(i){return L.med[i];}))));
  P.put('8|8.11',P.table(FH,byValue(lossBy(function(i){return L.prov[i];}))));
  P.put('8|8.12',P.table(FH,byValue(lossBy(function(i){return L.dist[i];}))));
});
"""
