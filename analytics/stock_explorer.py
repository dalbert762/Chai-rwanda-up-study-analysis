"""Client-side filter bar for the stock adequacy and inventory accuracy page.

The twin of explorer.py for Analysis 2: the rows are embedded and the browser
filters and redraws, for the same reason - the dashboard is one file opened
offline. Five filters: province, facility level, programme, facility name and
medicine name (type-ahead, matching the start of any word).

THE ENCODING
Facilities and medicines are lookup lists; each row is a fixed-width record.
    adequacy  fac (2) + medicine (2) + band (1) + months of stock x 100 (4)
    accuracy  fac (2) + medicine (2) + 1 if physical = system else 0 (1)
The band (0 under, 1 adequate, 2 over, 3 under AND at or below the
emergency order point, 4 stocked out) is computed at build time by the
report's _band(), so the teaching/referral-hospital cut-off lives in one place and the browser
never restates it. Months of stock travel only for the medians in 2A/2B.

WHAT THE FILTERS DRIVE
FIG 2.1-2.11, tables 2A/2B and the four cards at the top of the page
(pc2-0..3, rewritten at the end of run() exactly as page_cards() words them).
"""
import json

import rms_hubs as HUB

from explorer import CSS, b36, esc    # noqa: F401 - CSS re-used by the page


def build(adequacy, accuracy, fac_rec, med_rec, band, mos, exact, qacc=None,
          system=None):
    """-> (facilities, medicines, adequacy string, accuracy string,
    system names). system(r) -> the cleaned digital-system name of an
    accuracy record, for the "Digital system" filter.

    fac_rec(r) -> [name, district, province, level, management]
    med_rec(r) -> [name, programme, drug class label]"""
    facs, meds = {}, {}

    def ix(table, rec):
        if rec[0] not in table:
            table[rec[0]] = (len(table), rec)
        return table[rec[0]][0]

    ad = []
    for r in adequacy:
        m = max(0, min(36 ** 4 - 1, round((mos(r) or 0) * 100)))
        ad.append(b36(ix(facs, fac_rec(r)), 2) + b36(ix(meds, med_rec(r)), 2)
                  + str(band(r)) + b36(m, 4))
    ac, sysl, sysi = [], [], {}
    for r in accuracy:
        q = qacc(r) if qacc else None     # x 1e6; zzzz = undefined
        nm = system(r) if system else ""
        if nm not in sysi:
            sysi[nm] = len(sysl)
            sysl.append(nm)
        ac.append(b36(ix(facs, fac_rec(r)), 2) + b36(ix(meds, med_rec(r)), 2)
                  + ("1" if exact(r) else "0")
                  + ("zzzz" if q is None else b36(round(q * 1e6), 4))
                  + b36(sysi[nm], 2))
    fl = [v[1] for v in sorted(facs.values(), key=lambda x: x[0])]
    ml = [v[1] for v in sorted(meds.values(), key=lambda x: x[0])]
    return fl, ml, "".join(ad), "".join(ac), sysl


def system_parts(names):
    """The individual systems inside the combined names ("eLMIS + SAP" ->
    eLMIS, SAP), most used first is not needed - alphabetical."""
    return sorted({p.strip() for n in names for p in n.split("+")
                   if p.strip()}, key=str.lower)


def filters(facs, meds, systems=()):
    provs = sorted({f[2] for f in facs})
    dists = sorted({f[1] for f in facs if f[1]})
    lvls = sorted({f[3] for f in facs})
    prgs = sorted({m[1] for m in meds if m[1] and m[1] != "(none)"},
                  key=lambda x: x.lower())

    def opts(vals):
        return "".join(f'<option value="{esc(v)}">{esc(v)}</option>'
                       for v in vals)

    return f'''<section class="ind"><div class="fbar">
<h4>Filter</h4>
<div class="fgrid">
  <div><label for="sx-prov">Province</label>
    <select id="sx-prov" class="ms" multiple><option value="">All provinces</option>
    {opts(provs)}</select></div>
  <div><label for="sx-dist">District</label>
    <select id="sx-dist" class="ms" multiple><option value="">All districts</option>
    {opts(dists)}</select></div>
  <div><label for="sx-hub">RMS hub</label>
    <select id="sx-hub" class="ms" multiple><option value="">All RMS hubs</option>
    {opts(HUB.ORDER)}</select></div>
  <div><label for="sx-lvl">Facility level</label>
    <select id="sx-lvl" class="ms" multiple><option value="">All levels</option>
    {opts(lvls)}</select></div>
  <div><label for="sx-prg">Programme</label>
    <select id="sx-prg" class="ms" multiple><option value="">All programmes</option>
    {opts(prgs)}</select></div>
  <div><label for="sx-sys">Digital system (FIG 2.5 on)</label>
    <select id="sx-sys"><option value="">All systems</option>
    {opts(system_parts(systems))}</select></div>
  <div class="fsg"><label for="sx-fac">Facility name</label>
    <input id="sx-fac" type="search" placeholder="type the first letters"
     autocomplete="off"></div>
  <div class="fsg"><label for="sx-med">Medicine name</label>
    <input id="sx-med" type="search" placeholder="type the first letters"
     autocomplete="off"></div>
</div>
<div class="fsum">
  <button type="button" class="freset" onclick="sxReset()">Reset</button>
  <p class="fcmp" id="sx-sum"></p>
</div>
</div></section>'''


def script(facs, meds, ad, ac, systems, cfg):
    """cfg: min_n, levels, provs (map order), series, colours, blurbs,
    acc_steps [[lo, hi, colour]], acc_none, rest, avail, g_on, g_off."""
    data = json.dumps({"f": facs, "m": meds, "a": ad, "c": ac,
                       "hub": HUB.HUB_OF, "sys": list(systems)},
                      separators=(",", ":"))
    return """<script>
(function(){
var D=__DATA__,K=__CFG__;
var F=D.f,MD=D.m,MINN=K.min_n;
var NA=D.a.length/9,A=new Float64Array(NA*4);
for(var i=0;i<NA;i++){var o=i*9;
  A[i*4]=parseInt(D.a.substr(o,2),36);A[i*4+1]=parseInt(D.a.substr(o+2,2),36);
  // band 3 = understock at/below the emergency order point
  A[i*4+2]=D.a.charCodeAt(o+4)-48;A[i*4+3]=parseInt(D.a.substr(o+5,4),36)/100;}
// accuracy record: fac 2, med 2, exact 1, quantity accuracy 4, system 2
var NC=D.c.length/11,AC=new Int32Array(NC*3),AQ=new Float64Array(NC),
    AS=new Int32Array(NC);
for(var i=0;i<NC;i++){var o=i*11;AS[i]=parseInt(D.c.substr(o+9,2),36);
  AC[i*3]=parseInt(D.c.substr(o,2),36);AC[i*3+1]=parseInt(D.c.substr(o+2,2),36);
  AC[i*3+2]=D.c.charCodeAt(o+4)-48;
  // quantity accuracy x 1e6, NaN where the system quantity is 0
  var qv=D.c.substr(o+5,4);AQ[i]=qv==='zzzz'?NaN:parseInt(qv,36)/1e6;}
function starts(name,q){var n=name.toLowerCase();
  if(n.indexOf(q)===0)return true;
  var w=n.split(/[^a-z0-9]+/);
  for(var i=0;i<w.length;i++){if(w[i]&&w[i].indexOf(q)===0)return true;}
  return n.indexOf(' '+q)>=0;}
// multi-select filter lists (multiselect.py): ticked values, lower-cased,
// [] = all. A plain single select still works (its .value).
function msVals(id){var e=document.getElementById(id);if(!e)return [];
  if(e.multiple&&e.options)return Array.prototype.filter.call(e.options,
    function(o){return o.selected&&o.value;}).map(function(o){
    return o.value.trim().toLowerCase();});
  var v=String(e.value||'').trim().toLowerCase();return v?[v]:[];}
function msHas(a,x){return !a.length||
  a.indexOf(String(x).trim().toLowerCase())>=0;}
function val(id){var e=document.getElementById(id);
  return e?e.value.trim().toLowerCase():'';}
function esc(t){return String(t).replace(/[&<>"]/g,function(c){
  return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function pc1(v){return (v*100).toFixed(1)+'%';}
// Python's f'{x:.{d}f}' (pagefilter.py P.fx): an EXACT half goes to the
// even digit, where toFixed sends it up. Used by the scorecards.
function _gcd(a,b){while(b){var t=b;b=a%b;a=t;}return a;}
function fx(x,d){var s=x.toFixed(d),k=Math.pow(10,d),n2=Math.round(Math.abs(x)*2*k);
  if(n2%2===1&&Math.abs(Math.abs(x)*2*k-n2)<1e-6){
    var den=2*k/_gcd(n2,2*k);
    if((den&(den-1))===0){var lo=Math.floor(Math.abs(x)*k+1e-9),
        n=lo%2===0?lo:lo+1;s=((x<0?-1:1)*n/k).toFixed(d);}}
  return s;}
function put(id,h){var e=document.getElementById(id);if(e)e.innerHTML=h;}
var EMPTY='<p class="fempty">Nothing matches these filters.</p>';
var W=940;

// ---- stacked bars, matching charts.stacked()
function stacked(rows){
  if(!rows.length)return EMPTY;
  var LW=210,BH=22,GP=8,PT=46,cols=K.colours,ser=K.series,
      h=PT+rows.length*(BH+GP)+18,pw=W-LW-78,
      o=['<svg class="chart" viewBox="0 0 '+W+' '+h+'" role="img">'];
  ser.forEach(function(nm,i){var x=LW+i*160;
    o.push('<rect x="'+x+'" y="'+(PT-30)+'" width="11" height="11" rx="2" '
      +'fill="'+cols[i]+'"/><text x="'+(x+16)+'" y="'+(PT-20)+'" class="cl">'
      +esc(nm)+'</text>');});
  [0,25,50,75,100].forEach(function(t){var x=LW+pw*t/100;
    o.push('<line class="gl'+(t===0?' z':'')+'" x1="'+x.toFixed(1)+'" y1="'
      +(PT-8)+'" x2="'+x.toFixed(1)+'" y2="'+(h-16)+'"/><text x="'+x.toFixed(1)
      +'" y="'+(h-4)+'" class="cn" text-anchor="middle">'+t+'%</text>');});
  rows.forEach(function(r,i){var y=PT+i*(BH+GP),v=r[1],
      tot=v.reduce(function(a,b){return a+b;},0)||1,x=LW;
    o.push('<text x="'+(LW-10)+'" y="'+(y+BH-6)+'" class="ct" '
      +'text-anchor="end">'+esc(r[0])+'</text>');
    v.forEach(function(c,j){var w=pw*c/tot;if(w<=0)return;
      o.push('<rect class="seg" x="'+x.toFixed(1)+'" y="'+y+'" width="'
        +w.toFixed(1)+'" height="'+BH+'" fill="'+cols[j]+'"><title>'
        +esc(r[0])+' - '+esc(ser[j])+': '+c.toLocaleString()+' of '
        +tot.toLocaleString()+' ('+pc1(c/tot)+')</title></rect>');
      if(c/tot>=0.07)o.push('<text x="'+(x+w/2).toFixed(1)+'" y="'
        +(y+BH-7).toFixed(1)+'" text-anchor="middle" style="font-size:10.5px;'
        +'font-weight:700;fill:#fff">'+(c/tot*100).toFixed(1)+'%</text>');
      x+=w;});
    o.push('<text x="'+(W-6)+'" y="'+(y+BH-6)+'" class="cn" '
      +'text-anchor="end">n='+tot.toLocaleString()+'</text>');});
  return o.join('')+'</svg>';
}
// ---- radial rings, matching charts.radial()
function radial(vals,sub,hls){
  var tot=vals.reduce(function(a,b){return a+b;},0);if(!tot)return EMPTY;
  var w=940,h=470,RW=K.ring_w||34,RG=K.ring_gap||9,RC=84,n=vals.length,
      cx=250,cy=h/2,cols=K.colours,
      ro=RC+18+n*(RW+RG)-RG/2,
      o=['<svg class="chart" viewBox="0 0 '+w+' '+h+'" role="img">'];
  // FIG 2.1 pie infographic, as charts.pie_info() (user template 2026-10-26)
  var ang=-90,HUB=K.pie_hub,f1=function(x){return x.toFixed(1);};
  K.series.forEach(function(lab,i){var col=cols[i],sh=vals[i]/tot,
      sw=Math.min(sh,0.9999)*360,r=K.pie_reach[i%K.pie_reach.length],
      a0=ang*Math.PI/180,a1=(ang+sw)*Math.PI/180,am=(ang+sw/2)*Math.PI/180,
      rl=(HUB+r)/2+8;
    o.push('<path d="M'+f1(cx)+','+f1(cy)+' L'+f1(cx+r*Math.cos(a0))+','
      +f1(cy+r*Math.sin(a0))+' A'+r+','+r+' 0 '+(sw>180?1:0)+' 1 '
      +f1(cx+r*Math.cos(a1))+','+f1(cy+r*Math.sin(a1))+' Z" fill="'+col
      +'" stroke="#000" stroke-width="3" stroke-linejoin="round"><title>'
      +esc(lab)+': '+vals[i].toLocaleString()+' of '+tot.toLocaleString()
      +' ('+pc1(sh)+')</title></path>');
    o.push('<text x="'+f1(cx+rl*Math.cos(am))+'" y="'+f1(cy+rl*Math.sin(am)+7)
      +'" text-anchor="middle" style="font-size:21px;font-weight:700;'
      +'fill:#fff;paint-order:stroke;stroke:rgba(0,0,0,.25);stroke-width:2px">'
      +(sh*100).toFixed(1)+'%</text>');
    ang+=sw;});
  o.push('<circle cx="'+cx+'" cy="'+f1(cy)+'" r="'+(HUB+9)
    +'" fill="rgba(0,0,0,.45)"/><circle cx="'+cx+'" cy="'+f1(cy)+'" r="'+HUB
    +'" fill="#e9ecf1" stroke="#c9ced8" stroke-width="2"/>');
  o.push('<text x="'+cx+'" y="'+f1(cy-6)+'" text-anchor="middle" '
    +'style="font-size:14px;font-weight:700;fill:#1e2b45;letter-spacing:.03em">'
    +'STOCK ADEQUACY</text>');
  sub.forEach(function(ln,i){o.push('<text x="'+cx+'" y="'+f1(cy+13+i*13)
    +'" text-anchor="middle" style="font-size:10.5px;fill:#4a5568">'+esc(ln)
    +'</text>');});
  var kx=560,rh=Math.min(128,(h-40)/n),y0=cy-rh*n/2;
  K.series.forEach(function(lab,i){var col=cols[i],by=y0+rh*i+rh/2,v=vals[i];
    o.push('<circle cx="'+kx+'" cy="'+by.toFixed(1)+'" r="35" fill="none" '
      +'stroke="'+col+'" stroke-width="4"/>');
    o.push('<text x="'+kx+'" y="'+(by+5).toFixed(1)+'" text-anchor="middle" '
      +'style="font-size:'+(v<10000?15:13)+'px;font-weight:700;fill:#fff">'
      +v.toLocaleString()+'</text>');
    o.push('<text x="'+(kx+54)+'" y="'+(by-14).toFixed(1)+'" style="'
      +'font-size:19px;font-weight:700;fill:'+col+';letter-spacing:.04em">'
      +esc(lab.toUpperCase())+'</text>');
    var ws=K.blurbs[i].split(' '),ls=[],cur='';
    ws.forEach(function(x){if(cur.length+x.length+1<=46)cur=(cur+' '+x).trim();
      else{ls.push(cur);cur=x;}});if(cur)ls.push(cur);
    ls.slice(0,3).forEach(function(ln,j){o.push('<text x="'+(kx+54)+'" y="'
      +(by+6+j*16).toFixed(1)+'" style="font-size:12.5px;fill:#aeb6c6">'
      +esc(ln)+'</text>');});
    var hl=hls&&hls[i];
    if(hl){var hs=hl.split(' - ').filter(Boolean),
        hy=by+6+Math.min(3,ls.length)*16+5,mx=0;
      hs.forEach(function(x){mx=Math.max(mx,x.length);});
      o.push('<rect x="'+(kx+50)+'" y="'+(hy-13).toFixed(1)+'" width="'
        +Math.min(w-kx-56,10+mx*7.4).toFixed(0)+'" height="'+(6+16*hs.length)
        +'" rx="4" fill="'+K.hl+'" fill-opacity=".16" stroke="'+K.hl
        +'" stroke-opacity=".7"/>');
      hs.forEach(function(x,j){o.push('<text x="'+(kx+55)+'" y="'
        +(hy+j*16).toFixed(1)+'" style="font-size:12.5px;font-weight:700;'
        +'fill:'+K.hl+'">'+esc(x)+'</text>');});}});
  return o.join('')+'</svg>';
}
// ---- bars, matching charts.bars()
// mean=true: rows are [label, n, mean 0-1] (charts.mean_bars)
function bars(rows,hit,base,Wd,LWd,mean){
  if(!rows.length)return EMPTY;
  var W=Wd||940,LW=LWd||210,BH=19,GP=7,PT=40,h=PT+rows.length*(BH+GP)+16,
      pw=W-LW-78,
      o=['<svg class="chart" viewBox="0 0 '+W+' '+h+'" role="img">'];
  [0,25,50,75,100].forEach(function(t){var x=LW+pw*t/100;
    o.push('<line class="gl'+(t===0?' z':'')+'" x1="'+x.toFixed(1)+'" y1="'
      +(PT-8)+'" x2="'+x.toFixed(1)+'" y2="'+(h-14)+'"/><text x="'
      +x.toFixed(1)+'" y="'+(h-3)+'" class="cn" text-anchor="middle">'+t
      +'%</text>');});
  rows.forEach(function(r,i){var y=PT+i*(BH+GP),v=mean?r[2]:r[2]/r[1],
      w=Math.max(0,Math.min(v,1));
    o.push('<text x="'+(LW-10)+'" y="'+(y+BH-5)+'" class="ct" '
      +'text-anchor="end">'+esc(r[0])+'</text>');
    o.push('<rect x="'+LW+'" y="'+y+'" width="'+pw.toFixed(1)+'" height="'+BH
      +'" fill="'+K.rest+'"/>');
    o.push('<rect class="seg" x="'+LW+'" y="'+y+'" width="'+(pw*w).toFixed(1)
      +'" height="'+BH+'" fill="'+K.avail+'"><title>'+esc(r[0])+': '+pc1(v)
      +' '+hit+' - '+(mean?'over '+r[1].toLocaleString()
        :r[2].toLocaleString()+' of '+r[1].toLocaleString())+' '
      +base+'</title></rect>');
    o.push(w<0.86
      ? '<text x="'+(LW+pw*w+6).toFixed(1)+'" y="'+(y+BH-5)+'" class="cv">'
        +pc1(v)+'</text>'
      : '<text x="'+(LW+pw*w-6).toFixed(1)+'" y="'+(y+BH-5)
        +'" class="cv" text-anchor="end">'+pc1(v)+'</text>');
    o.push('<text x="'+(W-6)+'" y="'+(y+BH-5)+'" class="cn" '
      +'text-anchor="end">n='+r[1].toLocaleString()+'</text>');});
  return o.join('')+'</svg>';
}
// ---- one dial card, matching charts.dial()/dials()
var A0=150,SW=240;
function dpt(cx,cy,r,t){var a=(A0+SW*t)*Math.PI/180;
  return [cx+r*Math.cos(a),cy+r*Math.sin(a)];}
function darc(cx,cy,r,t0,t1){var a=dpt(cx,cy,r,t0),b=dpt(cx,cy,r,t1);
  return 'M'+a[0].toFixed(1)+','+a[1].toFixed(1)+' A'+r.toFixed(1)+','
    +r.toFixed(1)+' 0 '+((t1-t0)*SW>180?1:0)+' 1 '+b[0].toFixed(1)+','
    +b[1].toFixed(1);}
// FIG 2.5: the exact-match dial and the quantity-accuracy dial
function dial(n,hit,qn,qs){
  if(!n)return EMPTY;
  return '<div class="dgrid">'
    +dialCard(hit/n,'of records match the shelf','Exact match',
              '<b>'+hit.toLocaleString()+'</b> of '+n.toLocaleString()
              +' records')
    +dialCard(qn?qs/qn:0,'average quantity accuracy','Quantity accuracy',
              '<b>'+qn.toLocaleString()+'</b> records; a system quantity '
              +'of 0 counts as 0%')+'</div>';
}
function dialCard(value,caption,name,count){
  var w=272,h=226,T=41,cx=w/2,cy=h/2-6,rb=w/2-58,rt0=rb+16,
      rt1=rb+28,v=Math.max(0,Math.min(1,value)),
      o=['<div class="dcard"><svg viewBox="0 0 '+w+' '+h+'" role="img">'];
  for(var i=0;i<T;i++){var t=i/(T-1),a=dpt(cx,cy,rt0,t),b=dpt(cx,cy,rt1,t);
    o.push('<line x1="'+a[0].toFixed(1)+'" y1="'+a[1].toFixed(1)+'" x2="'
      +b[0].toFixed(1)+'" y2="'+b[1].toFixed(1)+'" stroke="'
      +(t<=v+1e-9?K.g_on:K.g_off)+'" stroke-width="'+(i%5?3.2:4.4)
      +'" stroke-linecap="round"/>');}
  o.push('<path d="'+darc(cx,cy,rb,0,1)+'" fill="none" stroke="'+K.g_off
    +'" stroke-width="13" stroke-linecap="round" opacity=".55"/>');
  if(v>0)o.push('<path d="'+darc(cx,cy,rb,0,v)+'" fill="none" stroke="'
    +K.g_on+'" stroke-width="13" stroke-linecap="round"/>');
  [[0,'0'],[.25,'25'],[.5,'50'],[.75,'75'],[1,'100']].forEach(function(s){
    var p=dpt(cx,cy,rt1+15,s[0]);o.push('<text x="'+p[0].toFixed(1)+'" y="'
      +(p[1]+4).toFixed(1)+'" class="dscale" text-anchor="middle">'+s[1]
      +'</text>');});
  var nd=dpt(cx,cy,rb-9,v);
  o.push('<line x1="'+cx.toFixed(1)+'" y1="'+cy.toFixed(1)+'" x2="'
    +nd[0].toFixed(1)+'" y2="'+nd[1].toFixed(1)+'" stroke="'+K.g_on
    +'" stroke-width="3.2" stroke-linecap="round"/>');
  o.push('<circle cx="'+cx.toFixed(1)+'" cy="'+cy.toFixed(1)+'" r="7.5" '
    +'fill="#141c2e" stroke="'+K.g_on+'" stroke-width="3.4"/>');
  o.push('<text x="'+cx.toFixed(0)+'" y="'+(cy+52).toFixed(0)+'" class="dbig" '
    +'text-anchor="middle">'+(value*100).toFixed(1)+'%</text>');
  o.push('<text x="'+cx.toFixed(0)+'" y="'+(cy+70).toFixed(0)+'" class="dcap" '
    +'text-anchor="middle">'+caption+'</text>');
  o.push('</svg><p class="dname">'+name+'</p><p class="dcount">'+count
    +'</p></div>');
  return o.join('');
}
// ---- 2A / 2B tables, matching html_table()
function median(v){if(!v.length)return null;v=v.slice().sort(function(a,b){
  return a-b;});var m=v.length>>1;
  return Math.round((v.length%2?v[m]:(v[m-1]+v[m])/2)*100)/100;}
function fnum(v){return v===null?'':v.toFixed(1)
  .replace(/\\B(?=(\\d{3})+(?!\\d))/g,',');}
function pcell(v){var w=Math.max(2,Math.round(v*54)),
  c=v>=.85?'':(v>=.75?' mid':' low');
  return pc1(v)+'<span class="bar'+c+'" style="width:'+w+'px"></span>';}
function table(rows){
  if(!rows.length)return '<p class="fempty">No medicine has '+MINN
    +' stocked rows under these filters.</p>';
  var hd=['Medicine','Programme','Stocked rows','Facilities','Understocked',
          'UNDERSTOCK %','Overstocked','OVERSTOCK %','Median months'];
  var h='<div class="scroll"><table><thead><tr>'+hd.map(function(x){
    return '<th>'+x+'</th>';}).join('')+'</tr></thead><tbody>';
  rows.forEach(function(r){h+='<tr><td>'+esc(r[0])+'</td><td>'+esc(r[1])
    +'</td><td class="num">'+r[2].toLocaleString()+'</td><td class="num">'
    +r[3]+'</td><td class="num">'+r[4].toLocaleString()+'</td><td class="num">'
    +pcell(r[5])+'</td><td class="num">'+r[6].toLocaleString()
    +'</td><td class="num">'+pcell(r[7])+'</td><td class="num">'
    +fnum(r[8])+'</td></tr>';});
  return h+'</tbody></table></div>';
}
function accColour(v){if(v===null)return K.acc_none;
  for(var i=0;i<K.acc_steps.length;i++){var s=K.acc_steps[i];
    if(v>=s[0]&&v<s[1])return s[2];}
  return K.acc_steps[K.acc_steps.length-1][2];}

function run(){
  // province / hub / level / programme are multi-select lists (multiselect.py)
  var pv=msVals('sx-prov'),lv=msVals('sx-lvl'),pg=msVals('sx-prg'),
      fq=val('sx-fac'),mq=val('sx-med'),hv=msVals('sx-hub'),sv=val('sx-sys'),
      dv=msVals('sx-dist');
  // Digital system: a record matches if its (possibly combined) system
  // name includes the chosen one; accuracy figures only
  var okS=D.sys.map(function(n){return !sv||n.toLowerCase().split('+')
    .map(function(x){return x.trim();}).indexOf(sv)>=0;});
  // RMS hub: the facility's district looked up in D.hub (rms_hubs.py)
  var okF=F.map(function(f){return msHas(pv,f[2])&&msHas(dv,f[1])
    &&msHas(hv,D.hub[f[1]]||'')
    &&msHas(lv,f[3])&&(!fq||starts(f[0],fq));});
  var okM=MD.map(function(m){return msHas(pg,m[1])
    &&(!mq||starts(m[0],mq));});
  // tallies are [stocked out, under, adequate, over]
  function add3(o,k,b){var g=o[k]||(o[k]=[0,0,0,0]);g[b]++;}
  var tot=[0,0,0,0],nE=0,byL={},byP={},byV={},byMed={},fs={};
  for(var i=0;i<NA;i++){var fi=A[i*4],mi=A[i*4+1];
    if(!okF[fi]||!okM[mi])continue;
    var b=A[i*4+2],f=F[fi],m=MD[mi];fs[fi]=1;
    if(b===3){nE++;b=0;}
    var c4=b===4?0:b+1;
    tot[c4]++;add3(byL,f[3],c4);add3(byP,m[1],c4);add3(byV,f[2],c4);
    if(b===4)continue;           // 2A/2B: stocked rows only
    var g=byMed[mi]||(byMed[mi]={b:[0,0,0],f:{},mos:[]});
    g.b[b]++;g.f[fi]=1;g.mos.push(A[i*4+3]);}
  var n=tot[0]+tot[1]+tot[2]+tot[3];
  put('sx-f21',radial(tot,[n.toLocaleString()+' medicines',
    'out or on the shelf'],
    [null,tot[1]?nE.toLocaleString()+' ('+(nE/tot[1]*100).toFixed(1)+'%) at or '
      +'below '+K.eop+' months - emergency order point':'',null,null]));
  put('sx-f22',stacked(K.levels.filter(function(l){return byL[l];})
    .map(function(l){return [l,byL[l]];})));
  // keys sorted first, then a stable sort on the overstock share, as the
  // Python does
  function overFirst(o,min){var out=[];
    Object.keys(o).sort(function(a,b){return a<b?-1:(a>b?1:0);})
    .forEach(function(k){var v=o[k],t=v[0]+v[1]+v[2]+v[3];
      if(k&&k!=="(none)"&&t>=min)out.push([k,v,0]);});
    function sh(x){var v=x[1];return v[3]/Math.max(1,v[0]+v[1]+v[2]+v[3]);}
    out.forEach(function(x,i){x[2]=i;});
    out.sort(function(a,b){return sh(b)-sh(a)||a[2]-b[2];});
    return out.map(function(x){return [x[0],x[1]];});}
  put('sx-f23',stacked(overFirst(byP,MINN)));
  put('sx-f24',stacked(overFirst(byV,1)));
  var mr=[];for(var k in byMed){var g=byMed[k],t=g.b[0]+g.b[1]+g.b[2];
    if(t<MINN)continue;
    mr.push([MD[k][0],MD[k][1],t,Object.keys(g.f).length,g.b[0],g.b[0]/t,
             g.b[2],g.b[2]/t,median(g.mos)]);}
  put('sx-t2a',table(mr.slice().sort(function(a,b){return b[5]-a[5];})
    .slice(0,15)));
  put('sx-t2b',table(mr.slice().sort(function(a,b){return b[7]-a[7];})
    .slice(0,15)));

  // ---- accuracy
  // [records, exact, facilities, quantity-accuracy sum]
  function bump(o,k,e,fi,q){var g=o[k]||(o[k]=[0,0,{},0]);g[0]++;g[1]+=e;
    g[2][fi]=1;g[3]+=q;}
  var cn=0,ch=0,qn=0,qs=0,aL={},aC={},aP={},aV={},aD={},aM={};
  for(var i=0;i<NC;i++){var fi=AC[i*3],mi=AC[i*3+1];
    if(!okF[fi]||!okM[mi]||!okS[AS[i]])continue;
    var e=AC[i*3+2],f=F[fi],m=MD[mi];cn++;ch+=e;fs[fi]=1;
    var q=isNaN(AQ[i])?0:AQ[i];qn++;qs+=q;
    bump(aL,f[3],e,fi,q);bump(aC,m[2],e,fi,q);bump(aP,m[1]||'(none)',e,fi,q);
    bump(aV,f[2],e,fi,q);bump(aD,f[1],e,fi,q);bump(aM,f[4],e,fi,q);}
  function rows(o,keys,sort){var out=[];(keys||Object.keys(o).sort())
    .forEach(function(k){var g=o[k];if(!g||!k||g[0]<MINN)return;
      out.push([k,g[0],g[1]]);});
    if(sort)out.sort(function(a,b){return a[2]/a[1]-b[2]/b[1];});
    return out;}
  put('sx-f25',dial(cn,ch,qn,qs));
  var H='accurate',B='records';
  // FIG 2.6-2.8: exact match left, average quantity accuracy right, the
  // right in the left's row order
  var PW=K.pair_w,PL=K.pair_lw,QH='average quantity accuracy';
  function pair(o,keys,sort,lid,rid){var l=rows(o,keys,sort);
    put(lid,bars(l,H,B,PW,PL));
    put(rid,bars(l.map(function(r){return [r[0],r[1],o[r[0]][3]/o[r[0]][0]];}),
      QH,B,PW,PL,true));}
  pair(aL,K.levels,false,'sx-f26','sx-q26');
  pair(aP,null,true,'sx-f28','sx-q28');
  pair(aV,null,true,'sx-f29','sx-q29');
  // FIG 2.9: the two district maps
  function paint(id,val,tip,det,subCls,overall){
    var map=document.getElementById(id);if(!map)return;
    map.querySelectorAll('path[data-d]').forEach(function(el){
      var g=aD[el.getAttribute('data-d')],v=g&&g[0]>=MINN?val(g):null;
      el.setAttribute('fill',accColour(v));
      var t=el.querySelector('title');
      if(t&&!el.dataset.lab)el.dataset.lab=t.textContent.split(' - ')[0];
      if(t)t.textContent=el.dataset.lab+' - '
        +(v===null?'fewer than '+MINN+' records':pc1(v)+' '+tip)+'\\n'
        +det(g||[0,0,{},0]);});
    map.querySelectorAll('text[data-dv]').forEach(function(el){
      var g=aD[el.getAttribute('data-dv')];
      el.textContent=g&&g[0]>=MINN?pc1(val(g)):'n/a';});
    var st=map.querySelector('.'+subCls);
    if(st)st.textContent=st.textContent.replace(/overall .*$/,
      'overall '+overall);}
  paint('sx-f210',function(g){return g[1]/g[0];},'accurate',
    function(g){return g[1].toLocaleString()+' of '+g[0].toLocaleString()
      +' records match the shelf';},'msub-acc',cn?pc1(ch/cn):'-');
  paint('sx-q210',function(g){return g[3]/g[0];},QH,
    function(g){return 'average over '+g[0].toLocaleString()+' records';},
    'msub-qacc',qn?pc1(qs/qn):'-');
  // ---- scorecards pc2-0..3, as page_cards() builds them, on these rows
  function card(id,v,l,s){[['-v',v],['-l',l],['-s',s]].forEach(function(p){
    var e=document.getElementById(id+p[0]);if(e)e.textContent=p[1];});}
  function pc(x){return fx(x*100,1)+'%';}
  function ci(x){return String(x).replace(/\\B(?=(\\d{3})+(?!\\d))/g,',');}
  var eop=Number.isInteger(K.eop)?K.eop.toFixed(1):String(K.eop);
  if(n){
    card('pc2-0',pc(tot[0]/n),'stocked out',ci(tot[0])+' of '+ci(n)
      +' medicines not on the shelf on the visit day');
    card('pc2-1',pc(tot[1]/n),'understocked','under one month; '+ci(nE)+' ('
      +fx(nE/Math.max(1,tot[1])*100,1)+'%) at or below '+eop
      +' months - emergency order point');
    card('pc2-2',pc(tot[3]/n),'overstocked','more than two months (three at '
      +'teaching and referral hospitals); only '+pc(tot[2]/n)+' adequate');
  }else{
    card('pc2-0','-','stocked out','no medicines match these filters');
    card('pc2-1','-','understocked','no medicines match these filters');
    card('pc2-2','-','overstocked','no medicines match these filters');}
  // rate_by(_acc, district): every district incl. a blank one, n >= MIN_N,
  // names sorted then a stable sort on the rate - the lowest first
  var dl=[];Object.keys(aD).sort().forEach(function(k){var g=aD[k];
    if(g[0]>=MINN)dl.push([k,g[1]/g[0]]);});
  dl.sort(function(a,b){return a[1]-b[1];});
  if(cn)card('pc2-3',pc(ch/cn),'inventory record accuracy',
    'physical count equals the system count'+(dl.length?'; '+dl[0][0]
      +' least accurate at '+pc(dl[0][1]):''));
  else card('pc2-3','-','inventory record accuracy',
    'no accuracy records match these filters');
  var on=pv.length||dv.length||lv.length||pg.length||hv.length||fq||mq;
  put('sx-sum',on?esc(Object.keys(fs).length+' facilities \\u00b7 '
    +n.toLocaleString()+' stocked rows \\u00b7 '+cn.toLocaleString()
    +' accuracy records match these filters'):'');
}
window.sxReset=function(){
  ['sx-prov','sx-dist','sx-hub','sx-lvl','sx-prg','sx-sys','sx-fac','sx-med'].forEach(function(id){
    document.getElementById(id).value='';});run();
  // the page-wide engine (pagefilter.py) listens for this
  document.getElementById('sx-prov').dispatchEvent(new Event('change'));};
['sx-prov','sx-dist','sx-hub','sx-lvl','sx-prg','sx-sys'].forEach(function(id){
  document.getElementById(id).addEventListener('change',run);});
['sx-fac','sx-med'].forEach(function(id){
  document.getElementById(id).addEventListener('input',run);});
function suggest(id,names){
  var inp=document.getElementById(id),box=document.createElement('div');
  box.className='sugg';inp.parentNode.appendChild(box);
  var uniq=names.filter(function(n,i){return names.indexOf(n)===i;})
    .sort(function(a,b){return a.toLowerCase()<b.toLowerCase()?-1:1;});
  function hide(){box.className='sugg';}
  function show(){var q=inp.value.trim().toLowerCase();
    if(!q){hide();return;}
    var hit=uniq.filter(function(n){return starts(n,q);});
    hit.sort(function(a,b){var x=a.toLowerCase().indexOf(q)===0?0:1,
      y=b.toLowerCase().indexOf(q)===0?0:1;return x-y;});
    if(!hit.length){box.innerHTML='<p class="snone">No name starts with '
      +'that</p>';box.className='sugg on';return;}
    box.innerHTML=hit.slice(0,14).map(function(n){return '<button '
      +'type="button" data-n="'+esc(n)+'">'+esc(n)+'</button>';}).join('')
      +(hit.length>14?'<p class="snone">'+(hit.length-14)
      +' more - keep typing</p>':'');
    box.className='sugg on';}
  inp.addEventListener('input',show);inp.addEventListener('focus',show);
  inp.addEventListener('blur',function(){setTimeout(hide,150);});
  inp.addEventListener('keydown',function(e){if(e.key==='Escape')hide();});
  box.addEventListener('mousedown',function(e){
    var b=e.target.closest('button');if(!b)return;
    e.preventDefault();inp.value=b.getAttribute('data-n');hide();
    inp.dispatchEvent(new Event('input'));});
}
suggest('sx-fac',F.map(function(f){return f[0];}));
suggest('sx-med',MD.map(function(m){return m[0];}));
run();
})();
</script>""".replace("__DATA__", data).replace(
        "__CFG__", json.dumps(cfg, separators=(",", ":")))
