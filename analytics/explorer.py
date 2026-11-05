"""Client-side filter bar driving every finding on the availability page.

WHY THE DATA IS EMBEDDED
The dashboard is a single file opened from Box or from an email attachment,
often with no network, so there is no server to query. Every combination of
five filters cannot be pre-rendered either - that is thousands of maps and
charts. So the facility x molecule rows are embedded and the browser does the
filtering and the redrawing.

THE ENCODING
One fixed-width record per row, concatenated into a single string with no
separators: 2 base36 chars for the facility, 2 for the molecule, 1 for
availability. That is 5 bytes a row against roughly 25 for the same rows as
JSON, which keeps a 10,776-row dataset near 53KB instead of 270KB. Province,
district and level are looked up from the facility; drug class and programme
from the molecule, so none of them are repeated 10,000 times.

WHAT THE FILTERS DRIVE
Everything on the page: the headline figures, both maps, all four figures and
the medicine table. A district the filter empties is drawn as no-data rather
than as zero availability, and a facility with no matching rows is hidden -
an empty selection is not a bad result and must not be coloured like one.

Scales are not restated here. The maps' colour ramps and bubble-size anchors
come from Tracer_availability_mapping.scales_js(), and the chart colours from
charts.py, so a browser redraw cannot drift from the legend beside it or from
the build-time render it replaces.
"""
import html
import json

import charts as CH
import rms_hubs as HUB

B36 = "0123456789abcdefghijklmnopqrstuvwxyz"

# The column naming who runs the facility - government, faith based
# or public-private partnership.
MGMT = "Facility management"


def b36(n, width):
    out = ""
    while n:
        n, r = divmod(n, 36)
        out = B36[r] + out
    return out.rjust(width, "0")


def esc(x):
    return html.escape("" if x is None else str(x), quote=True)


def build(rows, M, FL, group_level, rows_b=None, items=None):
    """rows (expected medicines) [+ rows_b (managed only)] -> (facilities,
    molecules, packed record string).

    Each record is 6 characters: facility (2, base 36), molecule (2), 1 if
    available else 0, and a METHOD FLAG - 1 in the study base only, 2 in
    the managed-only base only, 3 in both. Most rows are in both; the
    browser keeps the rows whose flag matches the method switch, so one
    payload serves both methods.

    `items` are per-STRENGTH records for the medicine table, as tuples
    (facility record, strength name, programme, flags) with flags = 1 if
    available + 2 if managed and answered + 4 if expected (method 1).
    They pack as facility (2) + strength (2) + flags (1). Each molecule
    also carries the names of the strengths it combines, so a strength
    typed in the medicine search narrows the molecule-level charts too."""
    def s(v):
        return "" if v is None else str(v).strip()

    facs, meds, recs = {}, {}, {}
    for flag, src in ((1, rows), (2, rows_b or [])):
        for r in src:
            f = s(r["_p"][M.FAC])
            if f not in facs:
                facs[f] = [f, s(r["_p"][M.DIST]), s(r["_p"][M.PROV]),
                           group_level(r["_p"][M.LEVEL]),
                           s(r["_p"].get(MGMT)) or "(not recorded)"]
            m = s(r[M.MED])
            if m not in meds:
                meds[m] = [m, s(r[M.DCLASS]) or "(unclassified)",
                           s(r[M.PROGRAM]) or "(none)", []]
            for nm in r.get("_contributing") or [m]:
                if nm not in meds[m][3]:
                    meds[m][3].append(nm)
            k = (f, m)
            a = 1 if M.yes(r[M.AVAIL]) else 0
            if k in recs:
                recs[k][1] |= flag
            else:
                recs[k] = [a, flag]
    fi = {k: i for i, k in enumerate(facs)}
    mi = {k: i for i, k in enumerate(meds)}
    packed = "".join(b36(fi[f], 2) + b36(mi[m], 2) + str(a) + str(fl)
                     for (f, m), (a, fl) in recs.items())
    its, ii, ipacked = [], {}, []
    for frec, name, prog, flags in items or []:
        if frec[0] not in fi:
            fi[frec[0]] = len(facs)
            facs[frec[0]] = frec
        if name not in ii:
            ii[name] = len(its)
            its.append([name, prog])
        ipacked.append(b36(fi[frec[0]], 2) + b36(ii[name], 2) + str(flags))
    return (list(facs.values()), list(meds.values()), packed, its,
            "".join(ipacked))

CSS = """
.fbar{margin:0 0 22px;padding:16px 18px;border:1px solid var(--rule);
 border-radius:10px;background:var(--surface)}
.fbar h4{font-size:12px;letter-spacing:.09em;text-transform:uppercase;
 color:var(--accent);margin:0 0 12px}
.fgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(168px,1fr));
 gap:10px}
.fgrid label{display:block;font-size:11px;color:var(--muted);margin:0 0 4px;
 letter-spacing:.04em;text-transform:uppercase}
.fgrid select,.fgrid input{width:100%;font:inherit;font-size:13px;padding:7px 9px;
 color:var(--ink);background:#141413;border:1px solid var(--rule);
 border-radius:6px}
.fgrid input::placeholder{color:#6d6b66}
.fsg{position:relative}
.sugg{display:none;position:absolute;z-index:30;left:0;right:0;top:100%;
 margin-top:2px;background:#141413;border:1px solid var(--rule);
 border-radius:6px;max-height:280px;overflow:auto;
 box-shadow:0 6px 18px rgba(0,0,0,.5)}
.sugg.on{display:block}
.sugg button{display:block;width:100%;text-align:left;font:inherit;
 font-size:12.5px;padding:6px 10px;background:none;border:0;color:var(--ink);
 cursor:pointer}
.sugg button:hover{background:#2a2a28;color:var(--accent)}
.sugg .snone{margin:0;padding:6px 10px;font-size:12px;color:var(--muted)}
.fgrid select:focus,.fgrid input:focus{outline:none;border-color:var(--accent)}
.fsum{display:flex;flex-wrap:wrap;gap:26px;align-items:baseline;
 margin:16px 0 0;padding:14px 0 0;border-top:1px solid var(--rule)}
.fsum b{font-size:26px;font-weight:700;letter-spacing:-.01em}
.fsum span{font-size:12.5px;color:var(--ink2)}
.fsum .fnote{color:var(--muted);font-size:12px;margin-left:auto;max-width:44ch;
 text-align:right}
.fcmp{flex-basis:100%;margin:0;font-size:13px;color:#fff;font-weight:600}
.only-b{display:none}
.meth-b .only-a{display:none}
.meth-b .only-b{display:block}
.freset{font:inherit;font-size:12px;padding:6px 12px;cursor:pointer;
 color:var(--ink);background:#1f1f1e;border:1px solid var(--rule);
 border-radius:6px}
.freset:hover{border-color:var(--accent);color:var(--accent)}
.mtab td.av{font-variant-numeric:tabular-nums;font-weight:700}
.mtab .trk{display:inline-block;width:92px;height:9px;border-radius:99px;
 background:#3a3a38;vertical-align:middle;margin-right:9px;overflow:hidden}
.mtab .trk i{display:block;height:100%;background:#1c4f91}
.mtab tbody tr:hover{background:#17171a}
.fempty{color:var(--muted);font-size:13px;padding:18px 0}
"""


def filters(facs, meds):
    """The filter bar and the live summary. Sits at the top of the page."""
    provs = sorted({f[2] for f in facs})
    dists = sorted({f[1] for f in facs if f[1]})
    lvls = sorted({f[3] for f in facs})
    prgs = sorted({m[2] for m in meds if m[2] and m[2] != "(none)"},
                  key=lambda x: x.lower())

    def opts(vals):
        return "".join(f'<option value="{esc(v)}">{esc(v)}</option>'
                       for v in vals)

    return f'''<section class="ind"><div class="fbar">
<h4>Filter</h4>
<div class="fgrid">
  <div><label for="fx-meth">Availability method</label>
    <select id="fx-meth">
    <option value="a">1 - Expected medicines</option>
    <option value="b" selected>2 - Managed products only</option></select></div>
  <div><label for="fx-prov">Province</label>
    <select id="fx-prov" class="ms" multiple><option value="">All provinces</option>
    {opts(provs)}</select></div>
  <div><label for="fx-dist">District</label>
    <select id="fx-dist" class="ms" multiple><option value="">All districts</option>
    {opts(dists)}</select></div>
  <div><label for="fx-hub">RMS hub</label>
    <select id="fx-hub" class="ms" multiple><option value="">All RMS hubs</option>
    {opts(HUB.ORDER)}</select></div>
  <div><label for="fx-lvl">Facility level</label>
    <select id="fx-lvl" class="ms" multiple><option value="">All levels</option>
    {opts(lvls)}</select></div>
  <div><label for="fx-prg">Programme</label>
    <select id="fx-prg" class="ms" multiple><option value="">All programmes</option>
    {opts(prgs)}</select></div>
  <div class="fsg"><label for="fx-fac">Facility name</label>
    <input id="fx-fac" type="search" placeholder="type the first letters"
     autocomplete="off"></div>
  <div class="fsg"><label for="fx-med">Medicine name</label>
    <input id="fx-med" type="search" placeholder="type the first letters"
     autocomplete="off"></div>
</div>
<div class="fsum">
  <button type="button" class="freset" onclick="fxReset()">Reset</button>
  <p class="fcmp" id="fx-cmp"></p>
</div>
</div></section>'''


def table(min_n):
    """The medicine table. Sits at the end of the page."""
    return (f'<section class="ind">'
            f'<h4><span class="mid">TABLE</span>All medicines by level of '
            f'availability</h4>'
            f'<div id="fx-table" class="scroll tall"></div></section>')


def script(facs, meds, packed, min_n, min_fac, scales, chart_colours,
           items=(), ipacked=""):
    """One pass over the filtered rows, redrawing every finding on the page."""
    data = json.dumps({"f": facs, "m": meds, "p": packed, "n": min_n,
                       "nf": min_fac, "it": list(items), "q": ipacked,
                       "hub": HUB.HUB_OF},
                      separators=(",", ":"))
    # The script is full of percent signs, so it is injected by replacement
    # rather than %-formatting or .format().
    return """<script>
(function(){
var D=__DATA__, S=__SCALES__, C=__COLOURS__, RADIAL=__RADIAL__;
var NOMGT="(not recorded)";
var F=D.f,MD=D.m,P=D.p,MINN=D.n,MINF=D.nf,N=P.length/6,R=new Int32Array(N*4);
for(var i=0;i<N;i++){
  R[i*4]=parseInt(P.substr(i*6,2),36);
  R[i*4+1]=parseInt(P.substr(i*6+2,2),36);
  R[i*4+2]=P.charCodeAt(i*6+4)-48;
  R[i*4+3]=P.charCodeAt(i*6+5)-48;
}
var IT=D.it||[],Q=D.q||'',NQ=Q.length/5,RQ=new Int32Array(NQ*3);
for(var i=0;i<NQ;i++){
  RQ[i*3]=parseInt(Q.substr(i*5,2),36);
  RQ[i*3+1]=parseInt(Q.substr(i*5+2,2),36);
  RQ[i*3+2]=Q.charCodeAt(i*5+4)-48;
}
// A name matches when it, or any word in it, starts with what was typed:
// "kic" finds "Kicukiro CS", "amlo" both amlodipine strengths.
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
function pc(a,b){return b?(a/b*100).toFixed(1)+'%':'-';}
function esc(t){return String(t).replace(/[&<>"]/g,function(c){
  return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function ramp(steps,v,none){if(v===null)return none;
  for(var i=0;i<steps.length;i++){if(v>=steps[i][0]&&v<steps[i][1])
    return steps[i][2];}return steps[steps.length-1][2];}
function radius(v){if(v===null)return S.rmin;
  var f=Math.max(0,Math.min(1,(v-S.vlo)/(S.vhi-S.vlo)));
  return S.rmin+(S.rmax-S.rmin)*Math.sqrt(f);}
// A group thinner than the reporting minimum gets no rate at all. Drawing one
// anyway would put a confident colour on two observations.
function rate(g){return g&&g[0]>=MINN?g[1]/g[0]:null;}

// ---- chart primitives, matching charts.py so a redraw does not change style
var W=940,LW=210,BH=19,GP=7;
function svgOpen(w,h){return '<svg class="chart" viewBox="0 0 '+w+' '+h
  +'" role="img">';}
function bars(rows,el){
  if(!rows.length){el.innerHTML='<p class="fempty">Nothing matches these '
    +'filters.</p>';return;}
  var h=40+rows.length*(BH+GP)+16,pw=W-LW-78,o=[svgOpen(W,h)];
  [0,25,50,75,100].forEach(function(t){var x=LW+pw*t/100;
    o.push('<line class="gl'+(t===0?' z':'')+'" x1="'+x.toFixed(1)+'" y1="32" x2="'
      +x.toFixed(1)+'" y2="'+(h-14)+'"/><text x="'+x.toFixed(1)+'" y="'+(h-3)
      +'" class="cn" text-anchor="middle">'+t+'%</text>');});
  rows.forEach(function(r,i){var y=40+i*(BH+GP),v=r[2]/r[1],w=Math.min(v,1);
    o.push('<text x="'+(LW-10)+'" y="'+(y+BH-5)+'" class="ct" text-anchor="end">'
      +esc(r[0])+'</text>');
    o.push('<rect x="'+LW+'" y="'+y+'" width="'+pw.toFixed(1)+'" height="'+BH
      +'" fill="'+C.rest+'"/>');
    o.push('<rect class="seg" x="'+LW+'" y="'+y+'" width="'+(pw*w).toFixed(1)
      +'" height="'+BH+'" fill="'+C.avail+'"><title>'+esc(r[0])+': '
      +(v*100).toFixed(1)+'% available - '+r[2]+' of '+r[1]
      +' expected</title></rect>');
    o.push(w<0.86
      ? '<text x="'+(LW+pw*w+6).toFixed(1)+'" y="'+(y+BH-5)+'" class="cv">'
        +(v*100).toFixed(1)+'%</text>'
      : '<text x="'+(LW+pw*w-6).toFixed(1)+'" y="'+(y+BH-5)
        +'" class="cv" text-anchor="end">'+(v*100).toFixed(1)+'%</text>');
    o.push('<text x="'+(W-6)+'" y="'+(y+BH-5)+'" class="cn" text-anchor="end">n='
      +r[1].toLocaleString()+'</text>');});
  el.innerHTML=o.join('')+'</svg>';
}
// FIG 1 track columns, matching charts.track_columns()
function trackColumns(rows,el){
  if(!rows.length){el.innerHTML='<p class="fempty">Nothing matches these '
    +'filters.</p>';return;}
  var RR=RADIAL[0],RT=RADIAL[1],DK=RADIAL[2],k=rows.length,cw=64,pad=40,
      top=30,yb=290,ph=yb-top,step=(W-2*pad)/k,rad=cw*0.45,H=400,
      f1=function(x){return x.toFixed(1);},
      jx=function(i){return k>1&&k<=RR.length?Math.floor(i*(RR.length-1)/(k-1)
        +0.5):Math.min(i,RR.length-1);},
      o=[svgOpen(W,H),'<defs>'];
  for(var i=0;i<k;i++){var j=jx(i);o.push('<linearGradient id="tc'+i
    +'" x1="0" x2="1" y1="0" y2="0"><stop offset="0" stop-color="'+RR[j]
    +'"/><stop offset="1" stop-color="'+RT[j]+'"/></linearGradient>');}
  o.push('</defs><line x1="'+(pad-20)+'" y1="'+(yb+22)+'" x2="'+(W-pad+20)
    +'" y2="'+(yb+22)+'" stroke="#55554f" stroke-width="1.5"/>');
  function shape(x,y,w,h){var r=Math.min(rad,h);return 'M'+f1(x)+','+yb
    +' V'+f1(y)+' H'+f1(x+w-r)+' A'+f1(r)+','+f1(r)+' 0 0 1 '+f1(x+w)+','
    +f1(y+r)+' V'+yb+' Z';}
  function wrap(lab){var ls=[],cur='';String(lab).split(/\s+/).forEach(
    function(w){if(!w)return;if(cur&&cur.length+w.length+1>14){ls.push(cur);
      cur=w;}else cur=(cur+' '+w).trim();});ls.push(cur);return ls.slice(0,2);}
  rows.slice().reverse().forEach(function(r,pos){var i=k-1-pos,j=jx(i),
      v=r[1]?r[2]/r[1]:0,cx=pad+step*(pos+0.5),x=cx-cw/2,
      h=ph*Math.min(v,1),y=yb-h,p=(v*100).toFixed(1)+'%',
      tc=j<DK?'#14243d':'#fff';
    o.push('<path d="'+shape(x,top,cw,ph)+'" fill="#5a5a55"/>');
    o.push('<path class="seg" d="'+shape(x,y,cw,h)+'" fill="url(#tc'+i
      +')"><title>'+esc(r[0])+': '+p+' available - '+r[2].toLocaleString()
      +' of '+r[1].toLocaleString()+' expected</title></path>');
    o.push(h>=40?'<text x="'+f1(cx)+'" y="'+f1(y+26)+'" text-anchor="middle" '
      +'style="font-size:15px;font-weight:700;fill:'+tc+'">'+p+'</text>'
      :'<text x="'+f1(cx)+'" y="'+f1(y-8)+'" class="cv" text-anchor="middle">'
      +p+'</text>');
    o.push('<rect x="'+f1(cx-5)+'" y="'+(yb+17)+'" width="10" height="10" '
      +'transform="rotate(45 '+f1(cx)+' '+(yb+22)+')" fill="'+RR[j]+'"/>');
    var ls=wrap(r[0]);
    ls.forEach(function(line,li){o.push('<text x="'+f1(cx)+'" y="'
      +(yb+50+li*15)+'" class="ct" text-anchor="middle" style="font-weight:'
      +'700">'+esc(line)+'</text>');});
    o.push('<text x="'+f1(cx)+'" y="'+(yb+50+ls.length*15)+'" class="cn" '
      +'text-anchor="middle">n='+r[1].toLocaleString()+'</text>');});
  el.innerHTML=o.join('')+'</svg>';
}
// FIG 1 radial bars, matching charts.radial_bars()
function radialBars(rows,el){
  if(!rows.length){el.innerHTML='<p class="fempty">Nothing matches these '
    +'filters.</p>';return;}
  var k=rows.length,SW=14,ST=24,R0=34,rmax=R0+(k-1)*ST,cx=34+rmax,
      cy=30+rmax,h=cy+rmax+30,span=W-cx-90,o=[svgOpen(W,h)],f1=function(x){
        return x.toFixed(1);};
  rows.forEach(function(r,i){var v=r[1]?r[2]/r[1]:0,rr=R0+i*ST,
      col=RADIAL[0][k>1&&k<=RADIAL[0].length?Math.round(i*(RADIAL[0].length-1)
        /(k-1)):Math.min(i,RADIAL[0].length-1)],
      x1=cx+span*Math.min(v,1),p=(v*100).toFixed(1)+'%';
    o.push('<path d="M'+f1(cx+12)+','+f1(cy+rr)+' H'+f1(cx)+' A'+f1(rr)+','
      +f1(rr)+' 0 0 1 '+f1(cx)+','+f1(cy-rr)+' H'+f1(x1)+'" fill="none" '
      +'stroke="'+col+'" stroke-width="'+SW+'" stroke-linecap="round"><title>'
      +esc(r[0])+': '+p+' available - '+r[2].toLocaleString()+' of '
      +r[1].toLocaleString()+' expected</title></path>');
    o.push('<text x="'+f1(x1+14)+'" y="'+f1(cy-rr+4.5)+'" class="cv">'+p
      +'</text>');
    o.push('<text x="'+f1(cx+30)+'" y="'+f1(cy+rr+4.5)+'" class="ct" '
      +'style="font-weight:700">'+esc(r[0])+' <tspan class="cn" '
      +'style="font-weight:400">n='+r[1].toLocaleString()+'</tspan></text>');});
  el.innerHTML=o.join('')+'</svg>';
}
// Gauge cards, matching charts.gauges() so a redraw keeps the same geometry.
function gauge(v,w,h){
  var T=42,cx=w/2,cy=h-18,ro=w/2-21,ri=w/2-47,
      lit=Math.round(Math.max(0,Math.min(1,v))*T),o=[];
  o.push('<svg viewBox="0 0 '+w+' '+h+'" role="img">');
  for(var i=0;i<T;i++){
    var a=Math.PI*(1+(i+0.5)/T),c=Math.cos(a),s=Math.sin(a);
    o.push('<line x1="'+(cx+ri*c).toFixed(1)+'" y1="'+(cy+ri*s).toFixed(1)
      +'" x2="'+(cx+ro*c).toFixed(1)+'" y2="'+(cy+ro*s).toFixed(1)
      +'" stroke="'+(i<lit?C.gon:C.goff)
      +'" stroke-width="3.4" stroke-linecap="round"/>');}
  o.push('<text x="'+cx+'" y="'+(cy-34)+'" class="gval" text-anchor="middle">'
    +(v*100).toFixed(1)+'%</text>');
  o.push('<text x="'+cx+'" y="'+(cy-15)+'" class="gcap" text-anchor="middle">'
    +'available of those expected</text>');
  return o.join('')+'</svg>';
}
function gauges(rows,el){
  if(!rows.length){el.innerHTML='<p class="fempty">Nothing matches these '
    +'filters.</p>';return;}
  var o=['<div class="ggrid">'];
  rows.forEach(function(r){var v=r[1]?r[2]/r[1]:0;
    o.push('<div class="gcard">'+gauge(v,250,148)
      +'<p class="gname">'+esc(r[0])+'</p>'
      +'<p class="gcount"><b>'+r[2].toLocaleString()+'</b> <span>of '
      +r[1].toLocaleString()+'</span></p></div>');});
  el.innerHTML=o.join('')+'</div>';
}
// Dial gauges, matching charts.dial()/dials() exactly.
var A0=150,SW=240;
function dpt(cx,cy,r,t){var a=(A0+SW*t)*Math.PI/180;
  return [cx+r*Math.cos(a),cy+r*Math.sin(a)];}
function darc(cx,cy,r,t0,t1){var a=dpt(cx,cy,r,t0),b=dpt(cx,cy,r,t1),
  big=(t1-t0)*SW>180?1:0;
  return 'M'+a[0].toFixed(1)+','+a[1].toFixed(1)+' A'+r.toFixed(1)+','
    +r.toFixed(1)+' 0 '+big+' 1 '+b[0].toFixed(1)+','+b[1].toFixed(1);}
function dial(value,w,h){
  var T=41,cx=w/2,cy=h/2-6,rb=w/2-58,rt0=rb+16,rt1=rb+28,
      v=Math.max(0,Math.min(1,value)),o=[];
  o.push('<svg viewBox="0 0 '+w+' '+h+'" role="img">');
  for(var i=0;i<T;i++){var t=i/(T-1),a=dpt(cx,cy,rt0,t),b=dpt(cx,cy,rt1,t);
    o.push('<line x1="'+a[0].toFixed(1)+'" y1="'+a[1].toFixed(1)+'" x2="'
      +b[0].toFixed(1)+'" y2="'+b[1].toFixed(1)+'" stroke="'
      +(t<=v+1e-9?C.gon:C.goff)+'" stroke-width="'+(i%5?3.2:4.4)
      +'" stroke-linecap="round"/>');}
  o.push('<path d="'+darc(cx,cy,rb,0,1)+'" fill="none" stroke="'+C.goff
    +'" stroke-width="13" stroke-linecap="round" opacity=".55"/>');
  if(v>0)o.push('<path d="'+darc(cx,cy,rb,0,v)+'" fill="none" stroke="'+C.gon
    +'" stroke-width="13" stroke-linecap="round"/>');
  [[0,'0'],[0.25,'25'],[0.5,'50'],[0.75,'75'],[1,'100']].forEach(function(s){
    var p=dpt(cx,cy,rt1+15,s[0]);
    o.push('<text x="'+p[0].toFixed(1)+'" y="'+(p[1]+4).toFixed(1)
      +'" class="dscale" text-anchor="middle">'+s[1]+'</text>');});
  var n=dpt(cx,cy,rb-9,v);
  o.push('<line x1="'+cx.toFixed(1)+'" y1="'+cy.toFixed(1)+'" x2="'
    +n[0].toFixed(1)+'" y2="'+n[1].toFixed(1)+'" stroke="'+C.gon
    +'" stroke-width="3.2" stroke-linecap="round"/>');
  o.push('<circle cx="'+cx.toFixed(1)+'" cy="'+cy.toFixed(1)
    +'" r="7.5" fill="#141c2e" stroke="'+C.gon+'" stroke-width="3.4"/>');
  o.push('<text x="'+cx+'" y="'+(cy+52)+'" class="dbig" text-anchor="middle">'
    +(value*100).toFixed(1)+'%</text>');
  o.push('<text x="'+cx+'" y="'+(cy+70)+'" class="dcap" text-anchor="middle">'
    +'of medicines expected</text>');
  return o.join('')+'</svg>';
}
function dials(rows,el){
  if(!rows.length){el.innerHTML='<p class="fempty">Nothing matches these '
    +'filters.</p>';return;}
  var o=['<div class="dgrid">'];
  rows.forEach(function(r){var v=r[1]?r[2]/r[1]:0;
    o.push('<div class="dcard">'+dial(v,272,226)
      +'<p class="dname">'+esc(r[0])+'</p>'
      +'<p class="dcount"><b>'+r[2].toLocaleString()+'</b> of '
      +r[1].toLocaleString()+' expected'
      +(r[3]?' &#183; '+r[3]+' districts':'')+'</p></div>');});
  el.innerHTML=o.join('')+'</div>';
}
function sing(w){return /ies$/.test(w)?w.slice(0,-3)+'y':w.replace(/s$/,'');}
// Heat map - the twin of charts.heatmap(): rows x columns of rates on the
// page's red-amber-green scale, cells under MINN grey with their n.
var HEAT=[[0,.70,'#c9544a'],[.70,.80,'#e07b3c'],[.80,.90,'#d9a21b'],
          [.90,.95,'#6fae5c'],[.95,1.01,'#199e70']];
var HEATL=['below 70%','70-79%','80-89%','90-94%','95% and above'];
function heat(rows,cols,el){
  if(!el)return;
  if(!rows.length){el.innerHTML='<p class="fempty">Nothing matches these '
    +'filters.</p>';return;}
  var W=940,LW=190,CH=40,PT=80,G=4,k=cols.length,
      CW=(W-LW-8-G*(k-1))/k,H=PT+rows.length*(CH+G)+34,o=[];
  o.push('<svg class="chart" viewBox="0 0 '+W+' '+H+'" role="img">');
  HEAT.forEach(function(s,i){var x=LW+i*120;
    o.push('<rect x="'+x+'" y="6" width="12" height="12" rx="2" fill="'+s[2]
      +'"/><text x="'+(x+17)+'" y="16" class="cn">'+HEATL[i]+'</text>');});
  o.push('<rect x="'+(LW+600)+'" y="6" width="12" height="12" rx="2" fill="'
    +C.rest+'"/><text x="'+(LW+617)+'" y="16" class="cn">fewer than '+MINN
    +'</text>');
  cols.forEach(function(c,j){var cx=LW+j*(CW+G)+CW/2,w=String(c).split(' '),
      ls=[],cur='';
    w.forEach(function(x){if((cur+' '+x).trim().length<=14)cur=(cur+' '+x).trim();
      else{ls.push(cur);cur=x;}});ls.push(cur);
    ls.slice(0,3).forEach(function(ln,li){o.push('<text x="'+cx.toFixed(1)
      +'" y="'+(PT-36+li*12)+'" class="ct" text-anchor="middle" style="'
      +'font-weight:'+(j===k-1?700:400)+'">'+esc(ln)+'</text>');});
    // the column's N, as charts.heatmap()
    var cnn=0;rows.forEach(function(r){cnn+=r[1][j][0];});
    o.push('<text x="'+cx.toFixed(1)+'" y="'+(PT-36+Math.min(3,ls.length)*12)
      +'" class="cn" text-anchor="middle" style="font-size:10px">n='
      +cnn.toLocaleString('en-US')+'</text>');});
  rows.forEach(function(r,i){var y=PT+i*(CH+G);
    o.push('<text x="'+(LW-10)+'" y="'+(y+CH/2-1)+'" class="ct" '
      +'text-anchor="end">'+esc(r[0])+'</text><text x="'+(LW-10)+'" y="'
      +(y+CH/2+12)+'" class="cn" text-anchor="end" style="font-size:10px">n='
      +r[1][r[1].length-1][0].toLocaleString('en-US')+'</text>');
    r[1].forEach(function(c,j){var x=LW+j*(CW+G),n=c[0],a=c[1];
      if(!n){o.push('<rect x="'+x.toFixed(1)+'" y="'+y+'" width="'
        +CW.toFixed(1)+'" height="'+CH+'" rx="4" fill="#1f1f1e"/>');return;}
      var v=a/n,thin=n<MINN,f=thin?C.rest:ramp(HEAT,v,C.rest);
      o.push('<rect class="seg" x="'+x.toFixed(1)+'" y="'+y+'" width="'
        +CW.toFixed(1)+'" height="'+CH+'" rx="4" fill="'+f+'"'
        +(j===k-1?' stroke="#fff" stroke-width="1.5"':'')+'><title>'
        +esc(r[0])+' - '+esc(cols[j])+': '+(v*100).toFixed(1)+'% ('
        +a.toLocaleString()+' of '+n.toLocaleString()+')</title></rect>');
      o.push('<text x="'+(x+CW/2).toFixed(1)+'" y="'+(y+CH/2+(thin?0:5))
        +'" text-anchor="middle" style="font-size:'+(thin?12:14)
        +'px;font-weight:700;fill:#fff;paint-order:stroke;'
        +'stroke:rgba(0,0,0,.35);stroke-width:2px">'+(v*100).toFixed(1)
        +'%</text>');
      if(thin)o.push('<text x="'+(x+CW/2).toFixed(1)+'" y="'+(y+CH/2+14)
        +'" text-anchor="middle" class="cn" style="font-size:10px">n='+n
        +'</text>');});});
  el.innerHTML=o.join('')+'</svg>';
}
function dotplot(rows,el,unit){
  unit=unit||'districts';
  rows=rows.filter(function(r){return r[1]>0;});
  if(!rows.length){el.innerHTML='<p class="fempty">Nothing matches these '
    +'filters.</p>';return;}
  var vals=[];rows.forEach(function(r){vals.push(r[2]);
    for(var k in r[3])vals.push(r[3][k]);});
  var lo=Math.min.apply(null,vals)-0.03,hi=Math.max.apply(null,vals)+0.03;
  if(hi-lo<0.08){lo-=0.04;hi+=0.04;}
  var dense=0;rows.forEach(function(r){
    dense=Math.max(dense,Object.keys(r[3]).length);});
  var dotR=dense<=24?4.5:(dense<=80?3.4:2.8),
      dotO=dense<=24?1:(dense<=80?0.72:0.5);
  var RH=46,h=52+rows.length*RH+24,LW2=180,pw=W-LW2-70;
  function X(v){return LW2+pw*(v-lo)/(hi-lo);}
  var o=[svgOpen(W,h)];
  o.push('<rect x="'+LW2+'" y="9" width="11" height="11" rx="2" fill="'+C.avail
    +'"/><text x="'+(LW2+16)+'" y="18" class="cl">group availability</text>'
    +'<rect x="'+(LW2+170)+'" y="9" width="11" height="11" rx="2" fill="'+C.rest
    +'"/><text x="'+(LW2+186)+'" y="18" class="cl">one of the '+unit
    +'</text>');
  var step=(hi-lo)<=0.30?0.05:0.10;
  var t=Math.ceil(lo/step)*step;
  for(;t<hi;t+=step){o.push('<line class="gl" x1="'+X(t).toFixed(1)
    +'" y1="44" x2="'+X(t).toFixed(1)+'" y2="'+(h-20)+'"/><text x="'
    +X(t).toFixed(1)+'" y="'+(h-6)+'" class="cn" text-anchor="middle">'
    +(t*100).toFixed(0)+'%</text>');}
  rows.forEach(function(r,i){var y=52+i*RH+RH/2-8,ds=r[3],ks=Object.keys(ds);
    o.push('<text x="'+(LW2-10)+'" y="'+(y+4)+'" class="ct" text-anchor="end">'
      +esc(r[0])+'</text>');
    if(ks.length){var vs=ks.map(function(k){return ds[k];});
      o.push('<line x1="'+X(Math.min.apply(null,vs)).toFixed(1)+'" y1="'+y
        +'" x2="'+X(Math.max.apply(null,vs)).toFixed(1)+'" y2="'+y
        +'" stroke="rgba(255,255,255,.14)" stroke-width="3" '
        +'stroke-linecap="round"/>');
      ks.forEach(function(k){o.push('<circle cx="'+X(ds[k]).toFixed(1)+'" cy="'
        +y+'" r="'+dotR+'" fill="none" stroke="#9a9892" stroke-width="1.4" '
        +'opacity="'+dotO+'"><title>'+esc(k)+' - '+(ds[k]*100).toFixed(1)
        +'%</title></circle>');});
      var ord=ks.slice().sort(function(a,b){return ds[a]-ds[b];});
      var lod=ord[0],hid=ord[ord.length-1];
      o.push('<text x="'+X(ds[lod]).toFixed(1)+'" y="'+(y-11)
        +'" class="cs" text-anchor="middle">'+esc(lod)+' '
        +(ds[lod]*100).toFixed(1)+'%</text>');
      if(hid!==lod)o.push('<text x="'+X(ds[hid]).toFixed(1)+'" y="'+(y-11)
        +'" class="cs" text-anchor="middle">'+esc(hid)+' '
        +(ds[hid]*100).toFixed(1)+'%</text>');}
    var x=X(r[2]);
    o.push('<path d="M'+x.toFixed(1)+','+(y-7)+' L'+(x+6.5).toFixed(1)+','+y
      +' L'+x.toFixed(1)+','+(y+7)+' L'+(x-6.5).toFixed(1)+','+y+' Z" fill="'
      +C.avail+'"><title>'+esc(r[0])+' - '+(r[2]*100).toFixed(1)
      +'%</title></path>');
    o.push('<text x="'+(W-6)+'" y="'+(y+4)+'" class="cv" text-anchor="end">'
      +(r[2]*100).toFixed(1)+'%</text>');
    if(ks.length)o.push('<text x="'+(LW2-10)+'" y="'+(y+17)+'" class="cn" '
      +'text-anchor="end">'+ks.length+' '
      +(ks.length===1?sing(unit):unit)+' &#183; '
      +r[1].toLocaleString()+' obs</text>');});
  el.innerHTML=o.join('')+'</svg>';
}
// Sankey, matching charts.sankey(). Size is the count of available
// medicines; colour is the level of care a ribbon comes from.

function sankey(av,tt,leftOrder,el){
  var W=940,H=580,NW=15,GAP=7,PAD=120,TOPR=12,OTHER='Other classes',tot=0,k;
  for(k in av)tot+=av[k];
  if(!tot){el.innerHTML='<p class="fempty">Nothing matches these filters.'
    +'</p>';return;}
  var rs={};
  for(k in av){var p=k.split('\u0000');rs[p[1]]=(rs[p[1]]||0)+av[k];}
  var keep=Object.keys(rs).sort(function(a,b){return rs[b]-rs[a];})
    .slice(0,TOPR),keepS={};
  keep.forEach(function(x){keepS[x]=1;});
  function fold(src){var o={};for(var q in src){var p=q.split('\u0000'),
    r=keepS[p[1]]?p[1]:OTHER,kk=p[0]+'\u0000'+r;o[kk]=(o[kk]||0)+src[q];}
    return o;}
  var fa=fold(av),ft=fold(tt),ls={},rs2={},lt={},rt={};
  for(k in fa){var p=k.split('\u0000');
    ls[p[0]]=(ls[p[0]]||0)+fa[k];rs2[p[1]]=(rs2[p[1]]||0)+fa[k];}
  for(k in ft){var p2=k.split('\u0000');
    lt[p2[0]]=(lt[p2[0]]||0)+ft[k];rt[p2[1]]=(rt[p2[1]]||0)+ft[k];}
  var rights=keep.filter(function(x){return rs2[x];});
  if(rs2[OTHER])rights.push(OTHER);
  var lefts=leftOrder.filter(function(x){return ls[x];});
  function place(keys,tv){
    var ah=H-PAD-GAP*Math.max(0,keys.length-1),pos={},y=PAD/2;
    keys.forEach(function(x){var hh=ah*tv[x]/tot;pos[x]=[y,hh];y+=hh+GAP;});
    return pos;}
  var lp=place(lefts,ls),rp=place(rights,rs2),x0=2,x1=W-NW-2,o=[];
  o.push('<svg class="sk" viewBox="0 0 '+W+' '+H+'" role="img">');
  var lc={},rc={};
  lefts.forEach(function(x){lc[x]=lp[x][0];});
  rights.forEach(function(x){rc[x]=rp[x][0];});
  Object.keys(fa).sort(function(a,b){
    var pa=a.split('\u0000'),pb=b.split('\u0000');
    return lefts.indexOf(pa[0])-lefts.indexOf(pb[0])
        || rights.indexOf(pa[1])-rights.indexOf(pb[1]);})
  .forEach(function(kk){
    var p=kk.split('\u0000'),l=p[0],r=p[1],v=fa[kk],n=ft[kk]||v,
        hl=lp[l][1]*v/ls[l],hr=rp[r][1]*v/rs2[r],y0=lc[l],y1=rc[r];
    lc[l]+=hl;rc[r]+=hr;
    var xa=x0+NW,xb=x1,xm=((xa+xb)/2).toFixed(0),rate=n?v/n:null;
    o.push('<path class="rb" d="M'+xa+','+y0.toFixed(1)+' C'+xm+','
      +y0.toFixed(1)+' '+xm+','+y1.toFixed(1)+' '+xb+','+y1.toFixed(1)
      +' L'+xb+','+(y1+hr).toFixed(1)+' C'+xm+','+(y1+hr).toFixed(1)+' '+xm
      +','+(y0+hl).toFixed(1)+' '+xa+','+(y0+hl).toFixed(1)+' Z" fill="'
      +C.sk[lefts.indexOf(l)%C.sk.length]+'" opacity=".46"><title>'+esc(l)+' \u2192 '+esc(r)
      +'\\n'+(rate===null?'-':(rate*100).toFixed(1)+'%')+' available &#183; '
      +v.toLocaleString()+' of '+n.toLocaleString()
      +' expected</title></path>');});
  lefts.forEach(function(x,li){var y=lp[x][0],hh=lp[x][1],
    rate=lt[x]?ls[x]/lt[x]:null,ty=hh<24?y+hh/2-4:y+hh/2;
    o.push('<rect class="nd" x="'+x0+'" y="'+y.toFixed(1)+'" width="'+NW
      +'" height="'+hh.toFixed(1)+'" fill="'+C.sk[li%C.sk.length]+'"/>');
    o.push('<text class="sknode" x="'+(x0+NW+7)+'" y="'+ty.toFixed(1)+'">'
      +esc(x)+' <tspan class="skpc">'
      +(rate===null?'-':(rate*100).toFixed(1)+'%')+'</tspan></text>');
    o.push('<text class="sknum" x="'+(x0+NW+7)+'" y="'+(ty+10).toFixed(1)
      +'">'+ls[x].toLocaleString()+' of '+lt[x].toLocaleString()
      +' expected</text>');});
  rights.forEach(function(x){var y=rp[x][0],hh=rp[x][1],
    rate=rt[x]?rs2[x]/rt[x]:null;
    o.push('<rect class="nd" x="'+x1+'" y="'+y.toFixed(1)+'" width="'+NW
      +'" height="'+hh.toFixed(1)+'" fill="'+C.skd+'"/>');
    if(rs2[x]/tot>=0.012){
      o.push('<text class="sknode" x="'+(x1-7)+'" y="'+(y+hh/2).toFixed(1)
        +'" text-anchor="end">'+esc(x)+' <tspan class="skpc">'
        +(rate===null?'-':(rate*100).toFixed(1)+'%')+'</tspan></text>');
      o.push('<text class="sknum" x="'+(x1-7)+'" y="'+(y+hh/2+10).toFixed(1)
        +'" text-anchor="end">'+rs2[x].toLocaleString()+' of '
        +rt[x].toLocaleString()+' expected</text>');}});
  var kx=x0+NW+7;
  lefts.forEach(function(x,i){
    o.push('<rect x="'+kx.toFixed(0)+'" y="'+(H-26)+'" width="12" height="12" '
      +'rx="2" fill="'+C.sk[i%C.sk.length]+'"/><text class="sknum" x="'
      +(kx+17).toFixed(0)+'" y="'+(H-16)+'">'+esc(x)+'</text>');
    kx+=24+5.5*x.length;});
  o.push('<text class="sknum" x="'+(x0+NW+7)+'" y="'+(H-34)
    +'">Colour = level of care &#183; size = number of medicines available '
    +'&#183; the rate is on each label</text>');
  el.innerHTML=o.join('')+'</svg>';
}
function run(){
  // province / hub / level / programme are multi-select lists (multiselect.py)
  var pv=msVals('fx-prov'),lv=msVals('fx-lvl'),cl=msVals('fx-prg'),
      fq=val('fx-fac'),mq=val('fx-med'),hv=msVals('fx-hub'),
      dv=msVals('fx-dist');
  // RMS hub: the facility's district looked up in D.hub (rms_hubs.py)
  var okF=F.map(function(f){
    return msHas(pv,f[2])&&msHas(dv,f[1])&&msHas(lv,f[3])
      &&msHas(hv,D.hub[f[1]]||'')
      &&(!fq||starts(f[0],fq));});
  var okM=MD.map(function(m){
    return msHas(cl,m[2])
      &&(!mq||starts(m[0],mq)||(m[3]||[]).some(function(n){
        return starts(n,mq);}));});
  var meth=document.getElementById('fx-meth').value==='b'?2:1;
  document.body.classList.toggle('meth-b',meth===2);
  var tn=0,ta=0,fs={},ms={},byMed={},byCls={},byLvl={},byProv={},byDist={},
      byFac={},byMgt={},byFacMgt={},skf={},skt={},cn=[0,0,0],ca=[0,0,0],
      byPrg={},byPL={},byPP={};
  function bump(o,k){var g=o[k]||(o[k]=[0,0]);g[0]++;return g;}
  for(var i=0;i<N;i++){
    var fi=R[i*4],mi=R[i*4+1];
    if(!okF[fi]||!okM[mi])continue;
    var fl=R[i*4+3];
    if(fl&1){cn[1]++;ca[1]+=R[i*4+2];}
    if(fl&2){cn[2]++;ca[2]+=R[i*4+2];}
    if(!(fl&meth))continue;
    var a=R[i*4+2],f=F[fi],m=MD[mi];
    tn++;ta+=a;fs[fi]=1;ms[mi]=1;
    bump(byMed,mi)[1]+=a;
    if(!byMed[mi][2])byMed[mi][2]={};byMed[mi][2][fi]=1;
    bump(byCls,m[1])[1]+=a;
    bump(byPrg,m[2])[1]+=a;
    bump(byPL,m[2]+'\u0000'+f[3])[1]+=a;
    bump(byPP,m[2]+'\u0000'+f[2])[1]+=a;
    bump(byLvl,f[3])[1]+=a;
    bump(byProv,f[2])[1]+=a;
    bump(byDist,f[1])[1]+=a;
    bump(byFac,fi)[1]+=a;
    bump(byMgt,f[4])[1]+=a;
    var fm=byFacMgt[f[4]]||(byFacMgt[f[4]]={});
    var q=fm[fi]||(fm[fi]=[0,0]);q[0]++;q[1]+=a;
  }
  // ---- the headline counters were removed from the filter bar; set them
  // only if a build still has them, so this does not throw on either layout
  function put(id,txt){var e=document.getElementById(id);
    if(e)e.textContent=txt;}
  var e1=cn[1]?ca[1]/cn[1]:null,e2=cn[2]?ca[2]/cn[2]:null;
  put('fx-cmp',(e1===null||e2===null)?'':
    'Method 1 (expected medicines): '+pc(ca[1],cn[1])+' of '+cn[1].toLocaleString()
    +'  \u00b7  Method 2 (managed only): '+pc(ca[2],cn[2])+' of '
    +cn[2].toLocaleString()+'  \u00b7  difference '
    +((e2-e1)*100>=0?'+':'')+((e2-e1)*100).toFixed(1)+' pts');
  put('fx-natl',pc(ca[meth],cn[meth]));   // MAP 1 subtitle
  put('fx-pct',pc(ta,tn));
  put('fx-num',ta.toLocaleString());
  put('fx-den','available of '+tn.toLocaleString()+' expected');
  put('fx-fac-n',Object.keys(fs).length);
  put('fx-med-n',Object.keys(ms).length);
  var on=[];
  ['fx-prov','fx-dist','fx-hub','fx-lvl','fx-prg'].forEach(function(id){
    var e=document.getElementById(id),t=Array.prototype.filter.call(e.options,
      function(o){return o.selected&&o.value;}).map(function(o){return o.text;});
    if(t.length)on.push(t.join(' or '));});
  if(fq)on.push('facility "'+document.getElementById('fx-fac').value+'"');
  if(mq)on.push('medicine "'+document.getElementById('fx-med').value+'"');
  put('fx-state',on.length
    ? 'Filtered to '+on.join(' + ')+'. Every figure, map and table below '
      +'reflects this selection.'
    : 'Showing every facility and every medicine.');

  // ---- maps
  var dm=document.getElementById('distmap');
  if(dm){
    dm.querySelectorAll('path[data-d]').forEach(function(el){
      var v=rate(byDist[el.getAttribute('data-d')]);
      el.setAttribute('fill',ramp(meth===2?S.steps_b:S.steps,v,S.none));});
    dm.querySelectorAll('text[data-dv]').forEach(function(el){
      var v=rate(byDist[el.getAttribute('data-dv')]);
      el.textContent=v===null?'-':(v*100).toFixed(1)+'%';});
    dm.querySelectorAll('path[data-p]').forEach(function(el){
      var v=rate(byProv[el.getAttribute('data-p')]);
      el.setAttribute('fill',ramp(meth===2?S.steps_b:S.steps,v,S.none));});
    dm.querySelectorAll('text[data-pv]').forEach(function(el){
      var v=rate(byProv[el.getAttribute('data-pv')]);
      el.textContent=v===null?'-':(v*100).toFixed(1)+'%';});
  }
  var fm=document.getElementById('facmap');
  if(fm){
    var byName={};
    for(var k in byFac)byName[F[k][0]]=byFac[k];
    fm.querySelectorAll('circle[data-f]').forEach(function(el){
      var g=byName[el.getAttribute('data-f')];
      // No matching rows is not zero availability - hide it.
      if(!g){el.style.display='none';return;}
      el.style.display='';
      var v=g[1]/g[0];
      el.setAttribute('r',radius(v).toFixed(1));
      el.setAttribute('fill',ramp(S.pt,v,S.ptnone));});
  }
  // The Mapbox version of MAP 2 (when loaded) takes the same values.
  var gl={};
  for(var k3 in byFac){var g3=byFac[k3],v3=g3[1]/g3[0];
    gl[F[k3][0]]=[g3[0],g3[1],ramp(S.pt,v3,S.ptnone),radius(v3)];}
  window.fxFacLast=gl;
  if(window.fxFacHook)window.fxFacHook(gl);

  // ---- figures
  function rank(o){var out=[];for(var k in o){if(o[k][0]>=MINN)
    out.push([k,o[k][0],o[k][1]]);}
    out.sort(function(a,b){return a[2]/a[1]-b[2]/b[1];});return out;}
  var prw=[];for(var pk in byPrg){if(pk&&pk!=='(none)')
    prw.push([pk,byPrg[pk][0],byPrg[pk][1]]);}
  prw.sort(function(a,b){return a[2]/a[1]-b[2]/b[1];});
  trackColumns(prw,document.getElementById('fx-figp'));
  // FIG 2: programme x level heat map, all levels last
  var hlv=S.levels.filter(function(l){return byLvl[l];});
  var hrows=prw.map(function(r){
    var cells=hlv.map(function(l){var g=byPL[r[0]+'\u0000'+l];
      return g?[g[0],g[1]]:[0,0];});
    cells.push([r[1],r[2]]);
    return [r[0],cells];});
  heat(hrows,hlv.concat(['All levels']),document.getElementById('fx-figpl'));
  var lrows=[];S.levels.forEach(function(l){if(byLvl[l])
    lrows.push([l,byLvl[l][0],byLvl[l][1]]);});
  // FIG 3: level of care as dials (the province dials were removed)
  dials(lrows.map(function(r){return [r[0],r[1],r[2],0];}),
    document.getElementById('fx-fig2'));
  var prows=[];S.prov.forEach(function(p){var g=byProv[p];if(!g)return;
    var ds={};for(var d in byDist){if(S.dist[d]!==p)continue;
      if(byDist[d][0]>=MINN)ds[d]=byDist[d][1]/byDist[d][0];}
    prows.push([p,g[0],g[1]/g[0],ds]);});
  prows.sort(function(a,b){return a[2]-b[2];});
  // the province dials (old FIG 4) were removed on request (2026-10-19)
  // FIG 5: management type, facilities as the spread. The unrecorded
  // category is dropped here only - it cannot be compared with the three.
  var m4=[];for(var mk in byMgt){if(mk===NOMGT)continue;
    var g=byMgt[mk],fsp={};
    for(var k in byFacMgt[mk]||{}){var q=byFacMgt[mk][k];
      if(q[0]>=MINF)fsp[F[k][0]]=q[1]/q[0];}
    m4.push([mk,g[0],g[1]/g[0],fsp]);}
  m4.sort(function(a,b){return a[2]-b[2];});
  dotplot(m4,document.getElementById('fx-fig4'),'facilities');
  // FIG 6: programme x province heat map, all provinces last
  var hpv=S.prov.filter(function(p){return byProv[p];});
  heat(prw.map(function(r){
    var cells=hpv.map(function(p){var g=byPP[r[0]+'\u0000'+p];
      return g?[g[0],g[1]]:[0,0];});
    cells.push([r[1],r[2]]);
    return [r[0],cells];}),hpv.concat(['All provinces']),
    document.getElementById('fx-fig5'));
  // FIG 7-10 and the tables are redrawn by pf_a1.py (pagefilter.py).

  // ---- table: one row per STRENGTH, from the per-strength records
  var okI=IT.map(function(x){
    return msHas(cl,x[1])&&(!mq||starts(x[0],mq));});
  var bi={};
  for(var j=0;j<NQ;j++){
    var fj=RQ[j*3],ij=RQ[j*3+1],fg=RQ[j*3+2];
    if(!okF[fj]||!okI[ij])continue;
    var g=bi[ij]||(bi[ij]={n:0,m:0,a:0,e:0,ae:0,f:{}});
    g.n++;g.f[fj]=1;
    if(fg&2)g.m++;
    if(fg&1)g.a++;
    if(fg&4){g.e++;if(fg&1)g.ae++;}
  }
  var rows=[];
  for(var k2 in bi){var g=bi[k2],den=meth===2?g.m:g.e,num=meth===2?g.a:g.ae;
    rows.push([IT[k2][0],IT[k2][1],Object.keys(g.f).length,g.m,g.e,num,
               den?num/den:null,den]);}
  // highest availability first (user request 2026-10-18); medicines with
  // no denominator last; ties go to the larger denominator
  rows.sort(function(a,b){return (a[6]===null)-(b[6]===null)
    ||b[6]-a[6]||b[7]-a[7];});
  var el=document.getElementById('fx-table');
  if(!rows.length){el.innerHTML='<p class="fempty">No medicine matches these '
    +'filters.</p>';return;}
  var h='<table class="mtab"><thead><tr><th>Medicine (strength)</th>'
    +'<th>Programme</th><th>Facilities assessed</th><th>Managed</th>'
    +'<th>Expected</th><th>Available</th><th>AVAILABILITY ('
    +(meth===2?'of managed':'of expected')+')</th></tr></thead><tbody>';
  rows.forEach(function(x){var v=x[6],thin=x[7]<MINN;
    h+='<tr><td>'+esc(x[0])+'</td><td>'+esc(x[1])+'</td><td class="num">'
      +x[2]+'</td><td class="num">'+x[3].toLocaleString()+'</td>'
      +'<td class="num">'+x[4].toLocaleString()+'</td><td class="num">'
      +x[5].toLocaleString()+'</td>'
      +(v===null?'<td class="av">-</td>':'<td class="av"'
      +(thin?' style="color:#898781"':'')+'><span class="trk"><i style="width:'
      +(v*100).toFixed(1)+'%"></i></span>'+(v*100).toFixed(1)+'%</td>')
      +'</tr>';});
  el.innerHTML=h+'</tbody></table>';
}
window.fxReset=function(){
  ['fx-prov','fx-dist','fx-hub','fx-lvl','fx-prg','fx-fac','fx-med'].forEach(function(id){
    document.getElementById(id).value='';});
  run();
  // the page-wide engine (pagefilter.py) listens for this
  document.getElementById('fx-prov').dispatchEvent(new Event('change'));};
['fx-meth','fx-prov','fx-dist','fx-hub','fx-lvl','fx-prg'].forEach(function(id){
  document.getElementById(id).addEventListener('change',run);});
['fx-fac','fx-med'].forEach(function(id){
  document.getElementById(id).addEventListener('input',run);});
// Type-ahead: the names whose start - or the start of any word in them -
// matches what has been typed, listed under the box; a click picks one.
function suggest(id,names){
  var inp=document.getElementById(id),box=document.createElement('div');
  box.className='sugg';inp.parentNode.appendChild(box);
  var uniq=names.filter(function(n,i){return names.indexOf(n)===i;})
    .sort(function(a,b){return a.toLowerCase()<b.toLowerCase()?-1:1;});
  function hide(){box.className='sugg';}
  function show(){
    var q=inp.value.trim().toLowerCase();
    if(!q){hide();return;}
    var hit=uniq.filter(function(n){return starts(n,q);});
    hit.sort(function(a,b){
      var x=a.toLowerCase().indexOf(q)===0?0:1,y=b.toLowerCase().indexOf(q)===0?0:1;
      return x-y;});
    if(!hit.length){box.innerHTML='<p class="snone">No name starts with '
      +'that</p>';box.className='sugg on';return;}
    box.innerHTML=hit.slice(0,14).map(function(n){
      return '<button type="button" data-n="'+esc(n)+'">'+esc(n)+'</button>';
    }).join('')+(hit.length>14?'<p class="snone">'+(hit.length-14)
      +' more - keep typing</p>':'');
    box.className='sugg on';
  }
  inp.addEventListener('input',show);
  inp.addEventListener('focus',show);
  inp.addEventListener('blur',function(){setTimeout(hide,150);});
  inp.addEventListener('keydown',function(e){if(e.key==='Escape')hide();});
  box.addEventListener('mousedown',function(e){
    var b=e.target.closest('button');if(!b)return;
    e.preventDefault();inp.value=b.getAttribute('data-n');hide();
    inp.dispatchEvent(new Event('input'));});
}
suggest('fx-fac',F.map(function(f){return f[0];}));
suggest('fx-med',IT.length?IT.map(function(x){return x[0];})
  :MD.map(function(m){return m[0];}));
run();
})();
</script>""".replace("__DATA__", data).replace("__SCALES__", scales) \
            .replace("__COLOURS__",
                     json.dumps(chart_colours, separators=(",", ":"))) \
            .replace("__RADIAL__", json.dumps(
                [CH.RADIAL_RAMP, CH.RADIAL_TINT, CH.RADIAL_DARK_TEXT]))
