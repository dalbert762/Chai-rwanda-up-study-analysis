"""Page filter for Analysis 3 - medicine consumption.

Outputs: tables 3.1, 3.1b, 3.1c (AMC_STATS), 3.2 (totals by level and
programme) and 3.3 (variability by drug class, _cvok), all from CROWS.

The row set carries only the three monthly quantities; total, AMC and CV are
derived in the browser with the same float operations as the report.
statistics.mean / stdev / pstdev are EXACT in Python (they sum fractions),
so the script reproduces them with BigInt fractions and a correctly rounded
conversion back to a double - otherwise round(mean, 1) and the CV cut-offs
could tip on the last bit.
"""


def build(pg, R):
    M, s = R.M, R.s
    pg.dataset("cr", R.CROWS, R.pf_fac, R.pf_med, {
        "c0": lambda r: r["_c"][0], "c1": lambda r: r["_c"][1],
        "c2": lambda r: r["_c"][2]})
    # exact whatever comes; a whole float is written as an int (the same
    # number in JavaScript, fewer bytes)
    for j in range(3):
        pg.D["cr"]["c"][f"c{j}"] = [
            int(r["_c"][j]) if r["_c"][j].is_integer() else r["_c"][j]
            for r in R.CROWS]
    # 3.3 groups on the RECORDED drug class (not the display label), which
    # is one value per medicine: a lookup by medicine index
    dc = {s(r[M.MED]): s(r[M.DCLASS]) or "(unclassified)" for r in R.CROWS}
    return (JS.replace("__LEVELS__", R.json.dumps(R.LEVELS))
            .replace("__DC__", R.json.dumps([dc.get(m[0], "(unclassified)")
                                             for m in pg.M])))


JS = r"""
var LEVELS=__LEVELS__,DC=__DC__;
// ---- exact arithmetic for Python's statistics module
var _dv=new DataView(new ArrayBuffer(8));
function _ex(x){_dv.setFloat64(0,x);var hi=_dv.getUint32(0),lo=_dv.getUint32(4),
  e=(hi>>>20)&0x7ff,m=(BigInt(hi&0xfffff)<<32n)|BigInt(lo);
  if(e===0)e=1;else m|=1n<<52n;if(hi>>>31)m=-m;return [m,e-1075];}
// exact sum of doubles -> [num, den] (den a power of two)
function _xsum(a){var p=a.map(_ex).filter(function(q){return q[0]!==0n;});
  if(!p.length)return [0n,1n];
  var k=Math.min.apply(null,p.map(function(q){return q[1];})),N=0n;
  p.forEach(function(q){N+=q[0]<<BigInt(q[1]-k);});
  return k>=0?[N<<BigInt(k),1n]:[N,1n<<BigInt(-k)];}
function _bl(x){return x.toString(2).length;}
function _p2(e){var v=1;while(e>1000){v*=Math.pow(2,1000);e-=1000;}
  while(e<-1000){v*=Math.pow(2,-1000);e+=1000;}return v*Math.pow(2,e);}
// num/den (den > 0) -> the nearest double, ties to even (float(Fraction))
function _q2d(N,D){if(N===0n)return 0;var neg=N<0n;if(neg)N=-N;
  var s=55-(_bl(N)-_bl(D)),num=s>=0?N<<BigInt(s):N,den=s>=0?D:D<<BigInt(-s),
      q=num/den,r=num%den,ex=_bl(q)-53,low=q&((1n<<BigInt(ex))-1n),
      half=1n<<BigInt(ex-1);
  q>>=BigInt(ex);
  if(low>half||(low===half&&(r!==0n||(q&1n)===1n)))q+=1n;
  var v=Number(q)*_p2(ex-s);return neg?-v:v;}
function pmean(v){var t=_xsum(v);return _q2d(t[0],t[1]*BigInt(v.length));}
// _ss() then variance (ddof 1) or pvariance (ddof 0), then sqrt
function psd(v,ddof){var n=v.length,c=pmean(v),
    a=_xsum(v.map(function(x){return (x-c)*(x-c);})),
    b=_xsum(v.map(function(x){return x-c;})),N=BigInt(n);
  return Math.sqrt(_q2d(a[0]*b[1]*b[1]*N-b[0]*b[0]*a[1],
                        a[1]*b[1]*b[1]*N*BigInt(n-ddof)));}
function med2(v){v=v.filter(function(x){return x!==null&&x!==undefined;});
  if(!v.length)return null;v=v.slice().sort(function(a,b){return a-b;});
  var m=v.length>>1,x=v.length%2?v[m]:(v[m-1]+v[m])/2;
  return P.pyRound(x,2);}
function F(v){return v===null||v===undefined?null:P.R(v);}
// ---- per row, once: total, AMC, CV
var D=P.D.cr,C=D.c,TOT=[],AMC=[],CV=[];
for(var i=0;i<D.n;i++){var v=[C.c0[i],C.c1[i],C.c2[i]],t=v[0]+v[1]+v[2];
  TOT.push(t);AMC.push(t/3);var m=pmean(v);
  CV.push(m===0?null:psd(v,0)/m*100);}
function medOf(i){return P.M[D.m[i]][0];}
function prgOf(i){return P.M[D.m[i]][1];}
function facOf(i){return P.F[D.f[i]][0];}
function lvlOf(i){return P.F[D.f[i]][3];}
function tsum(rs){var s=0;rs.forEach(function(i){s+=TOT[i];});return s;}
var AMC_HDR=['Medicine','Programme','Facilities','Total AMC','Mean AMC',
  'Median AMC','SD of AMC','CV','Total Jun-Aug'];
P.render(function(){
  var ix=P.rows('cr');
  // 3.1 / 3.1b / 3.1c: _amc_stats per medicine, names in sorted order
  var bm=P.by(ix,medOf),st=[];
  bm.keys.slice().sort(P.pyStrCmp).forEach(function(m){if(!m)return;
    var rs=bm.g[m],v=rs.map(function(i){return AMC[i];}),mean=pmean(v),
        sd=v.length>1?psd(v,1):null;
    // total AMC: summed in row order, as Python's sum() does
    var tot=0;v.forEach(function(x){tot+=x;});
    st.push([m,prgOf(rs[0]),P.uniq(rs,facOf),P.pyRound(tot,1),
      P.pyRound(mean,1),med2(v),
      sd!==null?P.pyRound(sd,1):null,
      sd!==null&&mean?P.pyRound(sd/mean*100,1):null,
      P.pyRound(tsum(rs),0)]);});
  function rowsOut(t){return t.map(function(x){return [x[0],x[1],x[2],F(x[3]),
    F(x[4]),F(x[5]),F(x[6]),F(x[7]),x[8]];});}
  var hi=P.sort(st,function(a,b){return b[3]-a[3];}),
      lo=P.sort(st,function(a,b){return a[3]-b[3];});
  P.put('3|3.1',P.table(AMC_HDR,rowsOut(hi.slice(0,20))));
  P.put('3|3.1b',P.table(AMC_HDR,rowsOut(lo.slice(0,20))));
  P.put('3|3.1c',P.table(AMC_HDR,rowsOut(hi)));
  // 3.2p: AMC by programme, the 3.1c columns, one row per programme
  var bp2=P.by(ix,prgOf),pr=[];
  bp2.keys.slice().sort(P.pyStrCmp).forEach(function(p){if(!p)return;
    var rs=bp2.g[p],v=rs.map(function(i){return AMC[i];}),mean=pmean(v),
        sd=v.length>1?psd(v,1):null,tot=0;v.forEach(function(x){tot+=x;});
    pr.push([p,P.uniq(rs,medOf),P.uniq(rs,facOf),P.pyRound(tot,1),
      P.pyRound(mean,1),med2(v),sd!==null?P.pyRound(sd,1):null,
      sd!==null&&mean?P.pyRound(sd/mean*100,1):null,P.pyRound(tsum(rs),0)]);});
  pr=P.sort(pr,function(a,b){return b[3]-a[3];});
  P.put('3|3.2p',P.table(['Programme','Medicines','Facilities'].concat(
    AMC_HDR.slice(3)),rowsOut(pr)));
  // 3.2: by level of care, then by programme
  var all=tsum(ix),rows=[];
  function grp(rs,k){rows.push([k,rs.length,P.pyRound(tsum(rs),0),
    F(med2(rs.map(function(i){return TOT[i];}))),F(P.pct(tsum(rs),all))]);}
  var bl=P.by(ix,lvlOf);
  LEVELS.forEach(function(l){if(bl.g[l])grp(bl.g[l],l);});
  rows.push(['','','','','']);
  var bp=P.by(ix,prgOf);
  bp.keys.slice().sort(P.pyStrCmp).forEach(function(p){grp(bp.g[p],p);});
  // 3|3.2 is off the dashboard (HTML_DROP, 2026-10-19)
  // 3.3: _cvok by recorded drug class, groups under MIN_N dropped
  var ok=ix.filter(function(i){return CV[i]!==null;}),
      bc=P.by(ok,function(i){return DC[D.m[i]];}),t3=[];
  bc.keys.slice().sort(P.pyStrCmp).forEach(function(c){var g=bc.g[c];
    if(g.length<P.MINN)return;
    var x=g.map(function(i){return CV[i];}),
        er=x.filter(function(y){return y>=50;}).length;
    t3.push([c,g.length,F(med2(x)),x.filter(function(y){return y<25;}).length,
      er,F(P.pct(er,g.length))]);});
  t3=P.sort(t3,function(a,b){return ((b[5]&&b[5].r)||0)-((a[5]&&a[5].r)||0);});
  // 3|3.3 is off the dashboard (HTML_DROP, 2026-10-19)
  // ---- scorecards pc3-0..2, as page_cards() builds them, on these rows
  function card(id,v,l,s){[['-v',v],['-l',l],['-s',s]].forEach(function(p){
    var e=document.getElementById(id+p[0]);if(e)e.textContent=p[1];});}
  var pc=function(x){return P.fx(x*100,1)+'%';};
  // Python str.title(): a cased letter after an uncased character goes up,
  // after a cased one goes down; [:30] counts code points
  function title(t){var o='',prev=false;Array.from(t).forEach(function(c){
    var cased=c.toLowerCase()!==c.toUpperCase();
    o+=cased?(prev?c.toLowerCase():c.toUpperCase()):c;prev=cased;});
    return o;}
  // max(wide, key=Total AMC): the first maximum in medicine-name order
  var top=null;st.forEach(function(x){if(x[2]>=10&&(!top||x[3]>top[3]))top=x;});
  if(top)card('pc3-0',P.fix(top[3],0),
    title(Array.from(top[0]).slice(0,30).join(''))+' - highest total AMC',
    'units a month, summed across '+top[2]+' facilities');
  else card('pc3-0','-','highest total AMC',
    'no medicine reported by 10 or more facilities');
  var er=ok.filter(function(i){return CV[i]>=50;}).length;
  if(ok.length)card('pc3-1',pc(er/ok.length),
    'of facility x medicine pairs erratic',
    'month-to-month variation of 50% or more');
  else card('pc3-1','-','of facility x medicine pairs erratic',
    'no pairs with consumption match these filters');
  var hc=bl.g['Health centre']?tsum(bl.g['Health centre']):0;
  if(all)card('pc3-2',pc(hc/all),'of all consumption at health centres',
    'June to August, all medicines');
  else card('pc3-2','-','of all consumption at health centres',
    'no consumption matches these filters');
});
"""
