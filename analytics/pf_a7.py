"""Page filter for Analysis 7 - requisition and order fulfilment (tab 6).

build(pg, R) registers the managed facility x medicine rows (R.MG) with the
kept medicine-month quantities - requested and delivered per month where
requested >= delivered (R._mreq / R._mdel), consumption over the same kept
months (R._amc_triplet) - and returns the script that recomputes every
figure and table on the page from the filtered rows.

Outputs: FIG 7.1-7.9 (dial, two-view map, level x month columns, class
dumbbell, programme lollipop, month and management bars, request/delivery
vs AMC paired dots and bands), TABLE 7A and table 7.8.

The map's boundary paths are not embedded a second time: the script reads
them from the map already on the page, once, and only the colours, counts
and labels are recomputed.
"""


def _map_cfg(R):
    MAP = R.MAP
    if MAP is None:
        return None

    def units(rows, big):
        out = []
        for name, (cx, cy), prov in rows:
            out.append([name, prov, MAP.slug(prov or name), f"{cx:.0f}",
                        f"{cy:.0f}", f"{cy + (14 if big else 11):.0f}"])
        return out
    d = units([(x["district"], x["cx_cy"], x["prov"]) for x in MAP.DROWS],
              len(MAP.DROWS) <= 5)
    p = units([(x["prov"], MAP._prov_centroid(x["prov"]), x["prov"])
               for x in MAP.PROWS], len(MAP.PROWS) <= 5)
    W, H, PAD = MAP.W, MAP.H, MAP.PAD
    tail = (MAP.north_arrow(W - PAD - 16, 66) + MAP.scale_bar(PAD + 4, H + 10)
            + f'<text x="{PAD + 4}" y="{H + 36}" class="msrc">Boundaries: '
            f'geoBoundaries ADM2 &#183; source Open Data Rwanda / NISR '
            f'&#183; CC BY 4.0.</text>')
    keys = "".join(f'<p class="key"><span class="sw" style="background:{c}">'
                   f'</span>{R.esc(lab)}</p>'
                   for (_, _, c), lab in zip(R.FILL_STEPS, R.FILL_LABELS))
    return {"W": W, "H": H, "PAD": PAD, "d": d, "p": p, "tail": tail,
            "keys": keys, "steps": [list(x) for x in R.FILL_STEPS],
            "uid": "fillmap", "title": "Order fill rate by geography",
            "legend": "Order fill rate", "noun": "requested units delivered"}


def build(pg, R):
    M = R.M
    mos = R.MONTHS
    so_ids = {id(r) for r in R.SO}
    cols = {"out": lambda r: id(r) in so_ids,
            "c": lambda r: R._amc_triplet(r)[2]}
    for j, m in enumerate(mos):
        cols[f"q{j}"] = (lambda mm: lambda r: R._mreq(r, mm))(m)
        cols[f"d{j}"] = (lambda mm: lambda r: R._mdel(r, mm))(m)
    # FIG 7.2: the three months' outcome codes (0-4) as one number
    cols["fu"] = lambda r: sum(c * 5 ** j for j, c in
                               enumerate(R._fulfil_cats(r)))
    pg.dataset("mg", R.MG, R.pf_fac, R.pf_med, cols)
    cfg = {"LEVELS": R.LEVELS, "MONTHS": list(mos), "MAP": _map_cfg(R),
           "BANDS": [list(b) for b in R.AMC_BANDS],
           "BAND_COLS": R.AMC_BAND_COLS, "REST": R.CH.REST,
           "DB_LOW": R.CH.DB_LOW, "DB_HIGH": R.CH.DB_HIGH,
           "FS": R.FULFIL_SERIES, "FC": R.FULFIL_COLS,
           "AVG": [R.CH.AVG_COL, R.CH.AVG_LABEL]}
    return JS.replace("__CFG__", R.json.dumps(cfg, separators=(",", ":")))


JS = r"""
var A7=__CFG__,LEVELS=A7.LEVELS,MONTHS=A7.MONTHS;
// ---- per-row quantities over the kept months (requested >= delivered)
var C7=P.D.mg.c,NR=P.D.mg.n,RQ=new Array(NR),DV=new Array(NR);
for(var _i=0;_i<NR;_i++){
  RQ[_i]=(C7.q0[_i]||0)+(C7.q1[_i]||0)+(C7.q2[_i]||0);
  DV[_i]=(C7.d0[_i]||0)+(C7.d1[_i]||0)+(C7.d2[_i]||0);}
function f7(i){return P.fac('mg',i);}
function m7(i){return P.med('mg',i);}
function sumA(idx,a){var s=0;idx.forEach(function(i){s+=a[i]||0;});return s;}
var fx1=function(v){return P.fx(v,1);};
// _fill_vals(): {key: [round(q), round(d), d/q]} for groups with q, in
// first-appearance order
function fillVals(idx,keyfn){var b=P.by(idx,keyfn),v={},keys=[];
  b.keys.forEach(function(k){var q=0,d=0;b.g[k].forEach(function(i){
      q+=RQ[i];d+=DV[i];});
    if(q){v[k]=[P.pyRound(q),P.pyRound(d),d/q];keys.push(k);}});
  return {v:v,keys:keys};}
// _fill_rows(): weakest first unless keys give the order
function fillRows(idx,keyfn,keys,drop){var fv=fillVals(idx,keyfn),v=fv.v,
    ks=keys?keys.filter(function(k){return k in v;})
      :P.sort(fv.keys,function(a,b){return v[a][2]-v[b][2];});
  drop=drop||[];
  return ks.filter(function(k){return k&&drop.indexOf(k)<0;})
    .map(function(k){return [k,v[k][0],v[k][1]];});}
// ---- charts.py ports
var NOTHING='<p class="fempty">Nothing to show.</p>';
function columns(rows,series,colours,avg){
  if(!rows.length)return NOTHING;
  // avg: a last pooled column per cluster, as charts.columns(avg=True)
  var nlab=rows.map(function(r){return [r[1][0][0],r[1][r[1].length-1][0]];});
  if(avg){rows=rows.map(function(r){var tn=0,th=0;
      r[1].forEach(function(nh){tn+=nh[0];th+=nh[1];});
      return [r[0],r[1].concat([[tn,th]])];});
    series=series.concat([A7.AVG[1]]);
    colours=colours.slice(0,series.length-1).concat([A7.AVG[0]]);}
  var width=940,height=330,pl=56,pb=74,pt=46,bw=26,inner=4,cols=colours,
      k=series.length,ph=height-pt-pb,pw=width-pl-16,step=pw/rows.length,
      o=['<svg class="chart" viewBox="0 0 '+width+' '+height+'" role="img">'];
  series.forEach(function(nm,i){var x=pl+i*150;
    o.push('<rect x="'+x+'" y="'+(pt-30)+'" width="11" height="11" rx="2" '
      +'fill="'+cols[i%cols.length]+'"/><text x="'+(x+16)+'" y="'+(pt-20)
      +'" class="cl">'+P.esc(nm)+'</text>');});
  [0,25,50,75,100].forEach(function(t){var y=pt+ph*(1-t/100);
    o.push('<line class="gl'+(t===0?' z':'')+'" x1="'+pl+'" y1="'+fx1(y)
      +'" x2="'+(width-16)+'" y2="'+fx1(y)+'"/><text x="'+(pl-8)+'" y="'
      +fx1(y+4)+'" class="cn" text-anchor="end">'+t+'%</text>');});
  rows.forEach(function(r,i){var lab=r[0],vals=r[1],cx=pl+step*(i+0.5),
      gw=k*bw+(k-1)*inner,x0=cx-gw/2;
    vals.forEach(function(nh,j){var n=nh[0],hit=nh[1],
        v=Math.min(n?hit/n:0,1),bh=ph*v,x=x0+j*(bw+inner),y=pt+ph-bh;
      o.push('<rect x="'+fx1(x)+'" y="'+fx1(pt)+'" width="'+bw+'" height="'
        +fx1(ph)+'" fill="'+A7.REST+'" opacity=".5"/>');
      o.push('<rect class="seg" x="'+fx1(x)+'" y="'+fx1(y)+'" width="'+bw
        +'" height="'+fx1(bh)+'" fill="'+cols[j%cols.length]+'"><title>'
        +P.esc(lab)+' - '+P.esc(series[j])+': '+P.pc1(v)+' ('+P.int(hit)
        +' of '+P.int(n)+')</title></rect>');
      o.push('<text x="'+fx1(x+bw/2)+'" y="'+fx1(y-4)+'" class="cv" '
        +'style="font-size:10px" text-anchor="middle">'+P.fx(v*100,1)
        +'%</text>');});
    var words=String(lab).split(/\s+/).filter(function(w){return w;}),
        lines=[],cur='';
    words.forEach(function(w){if(cur.length+w.length+1<=16)
        cur=(cur+' '+w).trim();else{lines.push(cur);cur=w;}});
    lines.push(cur);lines=lines.slice(0,3);
    lines.forEach(function(line,li){o.push('<text x="'+fx1(cx)+'" y="'
      +P.fx(pt+ph+16+li*11,0)+'" class="cn" text-anchor="middle">'
      +P.esc(line)+'</text>');});
    o.push('<text x="'+fx1(cx)+'" y="'+P.fx(pt+ph+16+lines.length*11,0)
      +'" class="cs" text-anchor="middle">n='+P.int(nlab[i][0])+' / '
      +P.int(nlab[i][1])+'</text>');});
  return o.join('')+'</svg>';}
function dumbbell(rows,ref,refLabel,unit){
  if(!rows.length)return NOTHING;
  var lw=250,rh=25,pt=48,width=940,hi=1.0,
      vals=rows.map(function(r){return r[2];}).concat([ref]),
      lo=Math.max(0,Math.floor(Math.min.apply(null,vals)*10)/10),
      h=pt+rows.length*rh+22,pw=width-lw-70,
      X=function(v){return lw+pw*(v-lo)/(hi-lo);},
      o=['<svg class="chart" viewBox="0 0 '+width+' '+h+'" role="img">'],
      step=hi-lo>0.3?0.1:0.05,t=lo;
  while(t<=hi+1e-9){
    o.push('<line class="gl" x1="'+fx1(X(t))+'" y1="'+(pt-10)+'" x2="'
      +fx1(X(t))+'" y2="'+(h-18)+'"/>');
    o.push('<text x="'+fx1(X(t))+'" y="'+(pt-16)+'" class="cn" '
      +'text-anchor="middle">'+P.fx(t*100,0)+'%</text>');
    t+=step;}
  o.push('<line x1="'+fx1(X(ref))+'" y1="'+(pt-8)+'" x2="'+fx1(X(ref))
    +'" y2="'+(h-18)+'" stroke="#fff" stroke-dasharray="3 4" opacity=".35"/>');
  o.push('<circle cx="'+(lw+8)+'" cy="14" r="6" fill="#fff" stroke="#222" '
    +'stroke-width="1.5"/><text x="'+(lw+20)+'" y="18" class="cn">'
    +P.esc(refLabel)+' '+P.fx(ref*100,1)+'%</text><circle cx="'+(lw+230)
    +'" cy="14" r="6" fill="'+A7.DB_LOW+'"/><text x="'+(lw+242)+'" y="18" '
    +'class="cn">below</text><circle cx="'+(lw+300)+'" cy="14" r="6" fill="'
    +A7.DB_HIGH+'"/><text x="'+(lw+312)+'" y="18" class="cn">at or above'
    +'</text>');
  rows.forEach(function(r,i){var lab=r[0],n=r[1],v=r[2],
      y=pt+i*rh+rh/2,col=v>=ref?A7.DB_HIGH:A7.DB_LOW;
    o.push('<line x1="'+lw+'" y1="'+fx1(y)+'" x2="'+(width-70)+'" y2="'
      +fx1(y)+'" stroke="rgba(255,255,255,.06)"/>');
    o.push('<text x="'+(lw-12)+'" y="'+fx1(y+4)+'" text-anchor="end" '
      +'style="font-size:12.5px;font-weight:'+(v>=ref?700:400)+';fill:'+col
      +'">'+P.esc(lab)+'</text>');
    o.push('<line x1="'+fx1(X(ref))+'" y1="'+fx1(y)+'" x2="'+fx1(X(v))
      +'" y2="'+fx1(y)+'" stroke="#8a8a86" stroke-width="2.4"/>');
    o.push('<circle cx="'+fx1(X(ref))+'" cy="'+fx1(y)+'" r="6.5" '
      +'fill="#fff" stroke="#222" stroke-width="1.5"/>');
    o.push('<circle cx="'+fx1(X(v))+'" cy="'+fx1(y)+'" r="7" fill="'+col
      +'" stroke="#111" stroke-width="1.2"><title>'+P.esc(lab)+': '
      +P.pc1(v)+' against '+refLabel+' '+P.pc1(ref)+' - n='+P.int(n)
      +(unit?' '+unit:'')+'</title></circle>');
    o.push('<text x="'+(width-6)+'" y="'+fx1(y+4)+'" class="cn" '
      +'text-anchor="end">'+P.fx(v*100,1)+'%</text>');});
  return o.join('')+'</svg>';}
function lollipop(rows,hit,base){
  if(!rows.length)return NOTHING;
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
      +P.pc1(v)+'</text>');});
  o.push('<line x1="'+x0+'" y1="'+(pt-12)+'" x2="'+x0+'" y2="'+(h-22)
    +'" stroke="#e8e8e2" stroke-width="5"/>');
  return o.join('')+'</svg>';}
function pairdots(rows,names){
  if(!rows.length)return NOTHING;
  var lw=210,rh=27,pt=50,width=940,colour='#4f9be8',vals=[];
  rows.forEach(function(r){if(r[0].indexOf('--')!==0)vals.push(r[2],r[3]);});
  var hi=Math.max(1.5,Math.ceil(Math.max.apply(null,vals)*2)/2),
      h=pt+rows.length*rh+24,pw=width-lw-120,
      X=function(v){return lw+pw*Math.min(v,hi)/hi;},
      o=['<svg class="chart" viewBox="0 0 '+width+' '+h+'" role="img">'],t=0;
  while(t<=hi+1e-9){
    o.push('<line class="gl'+(t===0?' z':'')+'" x1="'+fx1(X(t))+'" y1="'
      +(pt-10)+'" x2="'+fx1(X(t))+'" y2="'+(h-20)+'"/><text x="'+fx1(X(t))
      +'" y="'+(h-5)+'" class="cn" text-anchor="middle">'+P.fx(t*100,0)
      +'%</text>');
    t+=0.5;}
  o.push('<line x1="'+fx1(X(1))+'" y1="'+(pt-12)+'" x2="'+fx1(X(1))+'" y2="'
    +(h-20)+'" stroke="#fff" stroke-dasharray="4 4" opacity=".55"/>');
  o.push('<circle cx="'+(lw+8)+'" cy="16" r="6" fill="none" stroke="'+colour
    +'" stroke-width="2.5"/><text x="'+(lw+20)+'" y="20" class="cn">'
    +P.esc(names[0])+'</text><circle cx="'+(lw+230)+'" cy="16" r="6" fill="'
    +colour+'"/><text x="'+(lw+242)+'" y="20" class="cn">'+P.esc(names[1])
    +'</text><text x="'+fx1(X(1))+'" y="'+(pt-16)+'" class="cn" '
    +'text-anchor="middle">in line with use</text>');
  function p0(v){return P.fx(v*100,1)+'%';}
  rows.forEach(function(r,i){var y=pt+i*rh+rh/2;
    if(r[0].indexOf('--')===0){
      o.push('<text x="'+(lw-12)+'" y="'+fx1(y+5)+'" text-anchor="end" '
        +'style="font-size:11px;letter-spacing:.08em;fill:#d95926;'
        +'font-weight:700">'+P.esc(r[0].slice(2).toUpperCase())+'</text>');
      return;}
    var lab=r[0],n=r[1],a=r[2],b=r[3];
    o.push('<text x="'+(lw-12)+'" y="'+fx1(y+4.5)+'" class="ct" '
      +'text-anchor="end">'+P.esc(lab)+'</text>');
    o.push('<line x1="'+fx1(X(a))+'" y1="'+fx1(y)+'" x2="'+fx1(X(b))
      +'" y2="'+fx1(y)+'" stroke="#8a8a86" stroke-width="2.4"/>');
    o.push('<circle cx="'+fx1(X(a))+'" cy="'+fx1(y)+'" r="6.5" fill="#000" '
      +'stroke="'+colour+'" stroke-width="2.5"><title>'+P.esc(lab)+': '
      +P.esc(names[0])+' '+p0(a)+'</title></circle>');
    o.push('<circle cx="'+fx1(X(b))+'" cy="'+fx1(y)+'" r="6.5" fill="'
      +colour+'"><title>'+P.esc(lab)+': '+P.esc(names[1])+' '+p0(b)
      +' - n='+P.int(n)+' pairs</title></circle>');
    o.push('<text x="'+(width-6)+'" y="'+fx1(y+4.5)+'" class="cn" '
      +'text-anchor="end">'+p0(a)+' / '+p0(b)+'</text>');});
  return o.join('')+'</svg>';}
// ---- FIG 7.2: two_view_map() / _choro_svg(), paths read from the page
var MP=A7.MAP,PATHS=null;
function mapPaths(){
  if(PATHS&&(PATHS.d.length||PATHS.p.length))return PATHS;
  var out={d:[],p:[]};
  if(!window.__PF_CAPTURE){var root=document.getElementById(MP.uid);
    if(root&&root.querySelectorAll){
      ['d','p'].forEach(function(l){
        var ps=root.querySelectorAll('.mp-'+l+' path.dist');
        for(var i=0;i<ps.length;i++)out[l].push(ps[i].getAttribute('d'));});}}
  return (PATHS=out);}
function stepColour(v){if(v===null||v===undefined)return '#2c2c2a';
  for(var i=0;i<MP.steps.length;i++){var s=MP.steps[i];
    if(s[0]<=v&&v<s[1])return s[2];}
  return MP.steps[MP.steps.length-1][2];}
function choro(units,paths,vals,title,sub){
  var o=['<svg viewBox="0 0 '+MP.W+' '+(MP.H+54)+'" role="img" aria-label="'
    +P.esc(title)+'">','<text x="'+MP.PAD+'" y="18" class="mtitle">'
    +P.esc(title)+'</text>','<text x="'+MP.PAD+'" y="34" class="msub">'
    +P.esc(sub)+'</text>'],big=units.length<=5;
  function get(nm){return vals[nm]||[0,0,null];}
  units.forEach(function(u,i){var x=get(u[0]),v=x[2],
      where=u[1]&&u[1]!==u[0]?u[0]+', '+u[1]:u[0];
    o.push('<path class="dist pv-'+u[2]+'" d="'+P.esc(paths[i]||'')
      +'" fill="'+stepColour(v)+'"><title>'+P.esc(where)+' - '
      +(v!==null?P.pc1(v):'fewer than '+P.MINN)+'\\n'+P.int(x[1])+' of '
      +P.int(x[0])+' '+P.esc(MP.noun)+'</title></path>');});
  units.forEach(function(u){var x=get(u[0]),v=x[2],pv='pv-'+u[2];
    o.push('<text class="dlabel '+pv+'" x="'+u[3]+'" y="'+u[4]+'"'
      +(big?' style="font-size:13px"':'')+'>'+P.esc(u[0])+'</text>');
    o.push('<text class="dlabel '+pv+'" x="'+u[3]+'" y="'+u[5]+'" '
      +'style="font-size:'+(big?12:10.5)+'px">'+(v!==null?P.pc1(v):'n/a')
      +'</text>');});
  return o.join('')+MP.tail+'</svg>';}
function fillMap(dv,pv,tot,nLabel){var ps=mapPaths(),uid=MP.uid,
    sub=nLabel+' - overall '+P.pc1(tot),
    opts=MP.p.map(function(u){return '<option value="'+u[2]+'">'
      +P.esc(u[0])+' &#183; '+(u[0] in pv?P.pc1(pv[u[0]][2]):'n/a')
      +'</option>';}).join(''),
    panel='<div><div class="legend"><h4>'+P.esc(MP.legend)+'</h4>'+MP.keys
      +'</div><div class="layers"><p class="lyrh">View</p>'
      +'<label class="lyr"><input type="radio" class="v-d" name="'+uid
      +'-v" checked> Districts <span class="lyrn">'+MP.d.length
      +'</span></label><label class="lyr"><input type="radio" class="v-p" '
      +'name="'+uid+'-v"> Provinces <span class="lyrn">'+MP.p.length
      +'</span></label><p class="lyrh" style="margin-top:14px">Filter '
      +'&#183; province</p><select class="lyrsel" aria-label="Filter by '
      +'province"><option value="">All provinces</option>'+opts
      +'</select></div></div>';
  return '<div class="vmap maprow" id="'+uid+'"><div class="mapbox">'
    +'<div class="mp-d">'+choro(MP.d,ps.d,dv,MP.title+' - districts',sub)
    +'</div><div class="mp-p">'+choro(MP.p,ps.p,pv,MP.title+' - provinces',
    sub)+'</div></div>'+panel+'</div>';}
// keep the map's own View / province choice across a redraw
function putMap(key,h){var root=window.__PF_CAPTURE?null:
    document.getElementById(MP.uid),st=null;
  if(root&&root.querySelector){var r=root.querySelector('input.v-p'),
      s=root.querySelector('select.lyrsel');
    st={p:!!(r&&r.checked),s:s?s.value:''};}
  P.put(key,h);
  if(st){root=document.getElementById(MP.uid);if(!root)return;
    var r2=root.querySelector(st.p?'input.v-p':'input.v-d'),
        s2=root.querySelector('select.lyrsel');
    if(r2)r2.checked=true;if(s2)s2.value=st.s;}}
// ---- AMC
function amcRatio(rs){var c=sumA(rs,C7.c);
  return c?[sumA(rs,RQ)/c,sumA(rs,DV)/c]:[null,null];}
function amcGroup(base,keyfn,keys,drop){var b=P.by(base,keyfn),out=[];
  drop=drop||[];
  (keys?keys.filter(function(k){return k in b.g;})
       :b.keys.slice().sort(P.pyStrCmp)).forEach(function(k){
    if(!k||drop.indexOf(k)>=0||b.g[k].length<P.MINN)return;
    var r=amcRatio(b.g[k]);out.push([k,b.g[k].length,r[0],r[1]]);});
  return out;}
function amcBands(base,a){var c=A7.BANDS.map(function(){return 0;});
  base.forEach(function(i){var v=a[i]/C7.c[i];
    for(var j=0;j<A7.BANDS.length;j++){var bd=A7.BANDS[j];
      if(bd[1]<=v&&v<bd[2]){c[j]++;break;}}});
  return c;}
// ---- scorecards pc7-0..3, as page_cards() builds them
function card7(id,v,l,s){[['-v',v],['-l',l],['-s',s]].forEach(function(p){
  var e=document.getElementById(id+p[0]);if(e)e.textContent=p[1];});}
function cards7(ix){var pc=function(x){return P.fx(x*100,1)+'%';},
    rq=sumA(ix,RQ),dv=sumA(ix,DV);
  if(rq)card7('pc7-0',pc(dv/rq),'order fill rate',
    'quantity delivered / requested, Jun-Aug');
  else card7('pc7-0','-','order fill rate','nothing requested under these '
    +'filters');
  // min of (fill rate, programme) over programmes with a request
  var fv=fillVals(ix,function(i){return m7(i)[1]||'(none)';}),lo=null;
  fv.keys.forEach(function(k){if(k==='(none)')return;var r=fv.v[k][2];
    if(!lo||r<lo[0]||(r===lo[0]&&k<lo[1]))lo=[r,k];});
  if(lo)card7('pc7-1',pc(lo[0]),lo[1]+' - lowest fill rate',
    'of all programmes');
  else card7('pc7-1','-','lowest fill rate','no programme placed an order');
  var base=ix.filter(function(i){return C7.c[i]>0;});
  if(base.length){var c=sumA(base,C7.c);
    card7('pc7-2',pc(sumA(base,RQ)/c),'requested against consumption',
      'delivered '+pc(sumA(base,DV)/c)+' of consumption');}
  else card7('pc7-2','-','requested against consumption',
    'no consumption recorded under these filters');
  var bands=amcBands(base,RQ),nb=bands.reduce(function(a,b){return a+b;},0);
  if(nb)card7('pc7-3',pc(bands[2]/nb),'of medicine orders in line with use',
    'facility x medicine requests within 10% of consumption');
  else card7('pc7-3','-','of medicine orders in line with use',
    'no consumption recorded under these filters');}
P.render(function(){
  var ix=P.rows('mg'),
      lvl=function(i){return f7(i)[3];},
      cls=function(i){return m7(i)[2];},
      prg=function(i){return m7(i)[1]||'(none)';},
      mgt=function(i){return f7(i)[4];},
      rq=sumA(ix,RQ),dv=sumA(ix,DV),
      HIT='delivered',BASE='units requested';
  if(!ix.length){['FIG_7.1','FIG_7.2','FIG_7.3','FIG_7.4','FIG_7.5','FIG_7.6',
      'FIG_7.7','FIG_7.8','FIG_7.9','FIG_7.10','TABLE_7A','7.8']
      .forEach(function(k){P.put('7|'+k,P.EMPTY);});}
  cards7(ix);
  if(!ix.length)return;
  // FIG 7.8: requisition lines fully / partially / not (code 4, delivered
  // without a request, is off the figure)
  function fmix(rs){var o=[0,0,0];rs.forEach(function(i){var v=C7.fu[i];
    for(var j=0;j<3;j++){var k=v%5;v=(v-k)/5;if(k&&k<4)o[k-1]++;}});return o;}
  var frows=[['All facilities',fmix(ix)]],lb7=P.by(ix,lvl);
  LEVELS.forEach(function(l){if(lb7.g[l])frows.push([l,fmix(lb7.g[l])]);});
  P.put('7|FIG_7.8',P.stacked(frows,A7.FS,A7.FC));
  // FIG 7.1
  P.put('7|FIG_7.1',P.dials([['All facilities',P.pyRound(rq),
    P.pyRound(dv)]],'of requested quantity delivered','units requested',
    true));
  // FIG 7.2
  if(MP)putMap('7|FIG_7.2',fillMap(
    fillVals(ix,function(i){return f7(i)[1];}).v,
    fillVals(ix,function(i){return f7(i)[2];}).v,
    P.pct(dv,rq),P.int(P.pyRound(rq))+' units requested'));
  // FIG 7.3: level x month columns
  var bl=P.by(ix,lvl),crow=[];
  LEVELS.forEach(function(l){var rs=bl.g[l];if(!rs||!(sumA(rs,RQ)>0))return;
    crow.push([l,MONTHS.map(function(m,j){
      return [P.pyRound(sumA(rs,C7['q'+j])),P.pyRound(sumA(rs,C7['d'+j]))];
    })]);});
  P.put('7|FIG_7.3',columns(crow,MONTHS,['#1c4f91','#4f9be8','#a9cfdf'],true));
  // FIG 7.4: class dumbbell against the (filtered) overall fill rate
  var fv=fillVals(ix,cls),np={},ref=P.pct(dv,rq);
  ix.forEach(function(i){if(RQ[i]>0){var k=cls(i);np[k]=(np[k]||0)+1;}});
  var drow=P.sort(fv.keys.filter(function(k){return k&&(np[k]||0)>=P.MINN;})
    .map(function(k){return [k,np[k],fv.v[k][2]];}),
    function(a,b){return b[2]-a[2];});
  P.put('7|FIG_7.4',ref===null?NOTHING:dumbbell(drow,ref,'national',
    'pairs ordering'));
  // FIG 7.5: programme lollipop, highest first
  P.put('7|FIG_7.5',lollipop(P.sort(fillRows(ix,prg,null,['(none)']),
    function(a,b){return -(a[2]/a[1])+(b[2]/b[1]);}),HIT,BASE));
  // FIG 7.6: by month
  P.put('7|FIG_7.6',P.bars(MONTHS.map(function(m,j){
    return [m,P.pyRound(sumA(ix,C7['q'+j])),P.pyRound(sumA(ix,C7['d'+j]))];
  }),HIT,BASE));
  // FIG 7.7: management, weakest first
  P.put('7|FIG_7.7',P.bars(fillRows(ix,mgt,null,['(not recorded)']),HIT,
    BASE));
  // FIG 7.8 / 7.9 / TABLE 7A: against consumption
  var base=ix.filter(function(i){return C7.c[i]>0;});
  if(!base.length){['FIG_7.9','FIG_7.10','TABLE_7A'].forEach(function(k){
      P.put('7|'+k,P.EMPTY);});}
  else{var ab=amcRatio(base),
      prows=[['All facilities',base.length,ab[0],ab[1]],['--Level of care']]
      .concat(amcGroup(base,lvl,LEVELS),[['--Programme']],
        P.sort(amcGroup(base,prg,null,['(none)']),function(a,b){
          return b[2]-a[2];}));
    P.put('7|FIG_7.9',pairdots(prows,['Request vs AMC','Delivery vs AMC']));
    P.put('7|FIG_7.10',P.stacked([['Request vs AMC',amcBands(base,RQ)],
      ['Delivery vs AMC',amcBands(base,DV)]],
      A7.BANDS.map(function(b){return b[0];}),A7.BAND_COLS));
    var pc=function(v){return v!==null&&v!==undefined?P.fx(v*100,1)+'%':'';},
        bc=P.by(base,cls);
    var trs=P.sort(amcGroup(base,cls),function(a,b){return b[2]-a[2];})
      .map(function(r){var g=bc.g[r[0]];
        return [r[0],r[1],P.pyRound(sumA(g,C7.c)),P.pyRound(sumA(g,RQ)),
          P.pyRound(sumA(g,DV)),pc(r[2]),pc(r[3])];});
    P.put('7|TABLE_7A',trs.length?P.table(['Drug class',
      'Facility-medicine pairs','Consumed','Requested','Delivered',
      'Request vs AMC','Delivery vs AMC'],trs):P.EMPTY);}
  // 7.8: most under-supplied medicines
  var bm=P.by(ix,function(i){return m7(i)[0];}),trow=[];
  bm.keys.slice().sort(P.pyStrCmp).forEach(function(m){var rs=bm.g[m],
      q=sumA(rs,RQ),d=sumA(rs,DV);
    if(q<1000)return;
    trow.push([m,P.pyRound(q),P.pyRound(d),P.pyRound(q-d),P.R(P.pct(d,q)),
      P.cnt(rs,function(i){return C7.out[i];})]);});
  P.put('7|7.8',trow.length?P.table(['Medicine','Requested','Received',
    'Shortfall','FILL RATE','Stocked out on visit day'],P.sort(trow,
      function(a,b){return a[4].r-b[4].r;}).slice(0,15)):P.EMPTY);
});
"""
