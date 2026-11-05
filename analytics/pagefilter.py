"""Filters for every page: one engine, one filter bar per page, and the rows
each page's figures are computed from, embedded so the browser can recompute
them offline.

    pg = Page("a4")                                   # own filter bar
    pg.dataset("mg", R.MG, fac_rec, med_rec, {"so": lambda r: ...})
    html = pg.bar() + ... + pg.script(PAGE_JS)

THE CONTRACT WITH THE PAGE
write_html() tags every <section class="ind"> with data-k="<page>|<label>",
the label being the indicator id AS WRITTEN IN THE SOURCE (before the
display renumbering), spaces as underscores: "4|FIG_4.1", "4|4.7". A
repeated label gets "#2", "#3". The page script replaces a section's chart
or table with P.put(key, html); the slot is the section's last element that
is not its heading or a note, so the builders need no ids.

THE FILTERS
Province, facility level, programme, facility name, medicine name. A row
passes when its facility passes the first, second and fourth, and - when a
programme or medicine filter is set - at least one of its medicines passes
them. A row with no medicine (a facility-level record) is excluded while a
programme or medicine filter is set: it cannot be said to match.

THE DATA
Columnar. Facilities and medicines are lookup lists shared by the page's
datasets: facility [name, district, province, level, management], medicine
[name, programme, drug class]. A column of strings is dictionary-coded.

TYPES IN TABLES
Python's fmt() formats a float in [0, 1] as a percentage and an int with
commas; JavaScript has one number type, so a RATE must be passed to
P.table() wrapped as P.R(v). Unwrapped numbers format like Python ints when
whole and like non-whole floats otherwise.
"""
import html
import json

import rms_hubs as HUB

FILTERS = ("prov", "dist", "hub", "lvl", "prg", "fac", "med")


def esc(x):
    return html.escape("" if x is None else str(x), quote=True)


class Page:
    def __init__(self, key, ids=None, hide=(), tag=None):
        """key: page id ("a3"). ids: map FILTERS -> element ids of an
        existing filter bar to follow instead of drawing a new one. hide:
        filters left off this page's bar (they then never restrict).
        tag: an optional extra filter on a per-row value, as
        {"label": ..., "all": ..., "options": [[value, text], ...]}; a
        dataset opts in with dataset(..., tag=fn), fn(r) -> a value, a list
        of values (the row matches if any does) or None (never matches)."""
        self.key = key
        self.hide = set(hide)
        self.tag = tag
        self.own = ids is None
        self.ids = ids or {f: f"pf{key}-{f}" for f in FILTERS}
        if tag:
            self.ids = dict(self.ids, tag=f"pf{key}-tag")
        self.F, self._fi = [], {}
        self.M, self._mi = [], {}
        self.D = {}

    def _fac(self, rec):
        if rec is None:
            return -1
        rec = ["" if x is None else str(x) for x in rec]
        k = rec[0]
        if k not in self._fi:
            self._fi[k] = len(self.F)
            self.F.append(rec)
        return self._fi[k]

    def _med(self, rec):
        if rec is None:
            return -1
        if rec and isinstance(rec[0], (list, tuple)):
            out = sorted({self._med(x) for x in rec} - {-1})
            return out if len(out) != 1 else out[0]
        rec = ["" if x is None else str(x) for x in rec]
        k = rec[0]
        if k not in self._mi:
            self._mi[k] = len(self.M)
            self.M.append(rec)
        return self._mi[k]

    def dataset(self, name, rows, fac, med=None, cols=None, tag=None):
        """fac(r) -> facility record; med(r) -> medicine record, a list of
        them, or None; cols: {name: fn(r) -> number / bool / str / None}."""
        cols = cols or {}
        rows = list(rows)
        d = {"n": len(rows), "f": [self._fac(fac(r)) for r in rows],
             "m": [self._med(med(r)) if med else -1 for r in rows], "c": {}}
        for c, fn in cols.items():
            vals = [fn(r) for r in rows]
            if any(isinstance(v, str) for v in vals):
                keys = sorted({"" if v is None else str(v) for v in vals})
                ix = {k: i for i, k in enumerate(keys)}
                d["c"][c] = {"k": keys,
                             "v": [ix["" if v is None else str(v)]
                                   for v in vals]}
            else:
                d["c"][c] = [None if v is None else
                             (int(v) if isinstance(v, bool) else
                              (round(v, 6) if isinstance(v, float) else v))
                             for v in vals]
        if tag and self.tag:
            ix = {v: i for i, (v, _) in enumerate(self.tag["options"])}

            def code(v):
                if isinstance(v, (list, tuple, set)):
                    return sorted({ix[x] for x in v if x in ix})
                return ix.get(v, -1)
            d["t"] = [code(tag(r)) for r in rows]
        self.D[name] = d
        return self

    def bar(self):
        if not self.own:
            return ""
        provs = sorted({f[2] for f in self.F if f[2]})
        dists = sorted({f[1] for f in self.F if f[1]})
        lvls = sorted({f[3] for f in self.F if f[3]})
        prgs = sorted({m[1] for m in self.M if m[1] and m[1] != "(none)"},
                      key=str.lower)

        def opts(vals):
            return "".join(f'<option value="{esc(v)}">{esc(v)}</option>'
                           for v in vals)
        i = self.ids
        return f'''<section class="ind"><div class="fbar">
<h4>Filter</h4>
<div class="fgrid">
  <div><label for="{i['prov']}">Province</label>
    <select id="{i['prov']}" class="ms" multiple><option value="">All provinces</option>
    {opts(provs)}</select></div>
  {"" if "dist" in self.hide else f"""<div><label for="{i['dist']}">District</label>
    <select id="{i['dist']}" class="ms" multiple><option value="">All districts</option>
    {opts(dists)}</select></div>"""}
  {"" if "hub" in self.hide else f"""<div><label for="{i['hub']}">RMS hub</label>
    <select id="{i['hub']}" class="ms" multiple><option value="">All RMS hubs</option>
    {opts(HUB.ORDER)}</select></div>"""}
  {"" if "lvl" in self.hide else f"""<div><label for="{i['lvl']}">Facility level</label>
    <select id="{i['lvl']}" class="ms" multiple><option value="">All levels</option>
    {opts(lvls)}</select></div>"""}
  <div><label for="{i['prg']}">Programme</label>
    <select id="{i['prg']}" class="ms" multiple><option value="">All programmes</option>
    {opts(prgs)}</select></div>
  {self._tag_select()}
  <div class="fsg"><label for="{i['fac']}">Facility name</label>
    <input id="{i['fac']}" type="search" placeholder="type the first letters"
     autocomplete="off"></div>
  <div class="fsg"><label for="{i['med']}">Medicine name</label>
    <input id="{i['med']}" type="search" placeholder="type the first letters"
     autocomplete="off"></div>
</div>
<div class="fsum">
  <button type="button" class="freset" id="pf{self.key}-reset">Reset</button>
  <p class="fcmp" id="pf{self.key}-sum"></p>
</div>
</div></section>'''

    def _tag_select(self):
        if not self.tag:
            return ""
        t, i = self.tag, self.ids["tag"]
        opts = "".join(f'<option value="{k}">{esc(txt)}</option>'
                       for k, (_, txt) in enumerate(t["options"]))
        return (f'<div><label for="{i}">{esc(t["label"])}</label>'
                f'<select id="{i}"><option value="">{esc(t["all"])}</option>'
                f'{opts}</select></div>')

    def script(self, js, cfg=None):
        data = json.dumps({"F": self.F, "M": self.M, "D": self.D},
                          separators=(",", ":"))
        conf = json.dumps(dict(cfg or {}, key=self.key, ids=self.ids,
                               hubs=HUB.HUB_OF,
                               own=self.own), separators=(",", ":"))
        return ("<script>\n(function(){\nvar DATA=" + data + ",CFG=" + conf
                + ";\n" + ENGINE + "\n" + js + "\nP.start();\n})();\n"
                "</script>")


# ---------------------------------------------------------------
# section keys
# ---------------------------------------------------------------
def tag_sections(body, page_no):
    """Add data-k to every indicator section in one page's HTML."""
    import re
    seen = {}

    def one(m):
        sec = m.group(0)
        lab = re.search(r'<span class="mid">([^<]+)</span>', sec)
        lab = (lab.group(1) if lab else "untitled").replace(" ", "_")
        seen[lab] = seen.get(lab, 0) + 1
        k = f"{page_no}|{lab}" + (f"#{seen[lab]}" if seen[lab] > 1 else "")
        return sec.replace('<section class="ind">',
                           f'<section class="ind" data-k="{esc(k)}">', 1)
    return re.sub(r'<section class="ind">.*?</section>', one, body,
                  flags=re.S)


# ---------------------------------------------------------------
# the browser engine
# ---------------------------------------------------------------
ENGINE = r"""
var P={F:DATA.F,M:DATA.M,K:CFG,MINN:CFG.min_n||30};
// ---- decode columns
P.D={};
for(var dn in DATA.D){var d=DATA.D[dn],c={};
  for(var cn in d.c){var v=d.c[cn];
    if(v&&v.k){var a=new Array(v.v.length);
      for(var i=0;i<v.v.length;i++)a[i]=v.k[v.v[i]];c[cn]=a;}
    else c[cn]=v;}
  P.D[dn]={n:d.n,f:d.f,m:d.m,c:c,t:d.t};}
P.col=function(ds,c){return P.D[ds].c[c];};
P.fac=function(ds,i){var f=P.D[ds].f[i];return f<0?null:P.F[f];};
P.med=function(ds,i){var m=P.D[ds].m[i];
  if(m===-1||m===null)return null;return P.M[typeof m==='number'?m:m[0]];};
P.esc=function(t){return String(t===null||t===undefined?'':t)
  .replace(/[&<>"']/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;',
  '"':'&quot;',"'":'&#x27;'}[c];});};
P.starts=function(name,q){var n=String(name).toLowerCase();
  if(n.indexOf(q)===0)return true;
  var w=n.split(/[^a-z0-9]+/);
  for(var i=0;i<w.length;i++){if(w[i]&&w[i].indexOf(q)===0)return true;}
  return n.indexOf(' '+q)>=0;};
// multi-select filter lists (multiselect.py): ticked values, lower-cased,
// [] = all. A plain single select still works (its .value).
function msVals(id){var e=document.getElementById(id);if(!e)return [];
  if(e.multiple&&e.options)return Array.prototype.filter.call(e.options,
    function(o){return o.selected&&o.value;}).map(function(o){
    return o.value.trim().toLowerCase();});
  var v=String(e.value||'').trim().toLowerCase();return v?[v]:[];}
function msHas(a,x){return !a.length||
  a.indexOf(String(x).trim().toLowerCase())>=0;}
function _val(id){var e=document.getElementById(id);
  return e?String(e.value||'').trim().toLowerCase():'';}
// province / hub / level / programme are multi-select lists (multiselect.py)
P.state=function(){var i=CFG.ids;return {prov:msVals(i.prov),
  dist:i.dist?msVals(i.dist):[],
  hub:i.hub?msVals(i.hub):[],lvl:msVals(i.lvl),
  prg:msVals(i.prg),fac:_val(i.fac),med:_val(i.med),tag:i.tag?_val(i.tag):''};};
P.meth=function(){var e=document.getElementById('fx-meth');
  return e&&e.value==='b'?2:1;};
P.active=function(){var s=P.state();
  return !!(s.prov.length||s.dist.length||s.hub.length||s.lvl.length||s.prg.length||s.fac
    ||s.med||s.tag);};
// ---- the filter itself
P._okF=[];P._okM=[];P._sel={};
P._compute=function(){var s=P.state();
  // RMS hub: the facility's district looked up in CFG.hubs (rms_hubs.py)
  P._okF=P.F.map(function(f){return msHas(s.prov,f[2])&&msHas(s.dist,f[1])
    &&msHas(s.hub,CFG.hubs[f[1]]||'')
    &&msHas(s.lvl,f[3])&&(!s.fac||P.starts(f[0],s.fac));});
  P._medOn=!!(s.prg.length||s.med);
  P._tag=s.tag===''?null:+s.tag;
  P._okM=P.M.map(function(m){return msHas(s.prg,m[1])
    &&(!s.med||P.starts(m[0],s.med));});
  P._sel={};};
P.okRow=function(ds,i){var d=P.D[ds],f=d.f[i],m=d.m[i];
  if(f>=0&&!P._okF[f])return false;
  // the optional extra filter: only datasets that carry a tag column
  if(P._tag!==null&&d.t){var t=d.t[i];
    if(typeof t==='number'){if(t!==P._tag)return false;}
    else if(!t||t.indexOf(P._tag)<0)return false;}
  if(!P._medOn)return true;
  if(m===-1||m===null)return false;
  if(typeof m==='number')return P._okM[m];
  for(var j=0;j<m.length;j++)if(P._okM[m[j]])return true;return false;};
P.rows=function(ds){if(P._sel[ds])return P._sel[ds];
  var d=P.D[ds],out=[];for(var i=0;i<d.n;i++)if(P.okRow(ds,i))out.push(i);
  return (P._sel[ds]=out);};
// ---- aggregation helpers
P.by=function(idx,fn){var g={},keys=[];idx.forEach(function(i){var k=fn(i);
  if(!(k in g)){g[k]=[];keys.push(k);}g[k].push(i);});return {g:g,keys:keys};};
P.sum=function(idx,col){var s=0;idx.forEach(function(i){var v=col[i];
  if(v)s+=v;});return s;};
P.cnt=function(idx,fn){var n=0;idx.forEach(function(i){if(fn(i))n++;});return n;};
P.uniq=function(idx,fn){var o={};idx.forEach(function(i){o[fn(i)]=1;});
  return Object.keys(o).length;};
// Python's statistics.median, then med()'s round(., 2)
P.median=function(v){v=v.filter(function(x){return x!==null&&x!==undefined;});
  if(!v.length)return null;v=v.slice().sort(function(a,b){return a-b;});
  var m=v.length>>1,x=v.length%2?v[m]:(v[m-1]+v[m])/2;
  return P.pyRound(x,2);};
P.mean=function(v){return v.length?v.reduce(function(a,b){return a+b;},0)/v.length:null;};
P.pct=function(a,b){return b?a/b:null;};
// a sort that never reorders ties, whatever the browser
P.sort=function(a,cmp){return a.map(function(x,i){return [x,i];})
  .sort(function(p,q){return cmp(p[0],q[0])||p[1]-q[1];})
  .map(function(p){return p[0];});};
P.pyStrCmp=function(a,b){return a<b?-1:(a>b?1:0);};
// ---- formatting, matching Python fmt() / html_table()
P.R=function(v){return {r:v};};
// Python's f'{x:.{d}f}': correctly rounded, and an EXACT half goes to the
// even digit (JavaScript's toFixed sends it up). Every number printed as
// text goes through this, so a browser redraw reads like the build.
function _gcd(a,b){while(b){var t=b;b=a%b;a=t;}return a;}
P.fx=function(x,d){if(x===null||x===undefined||isNaN(x))return '';
  var s=x.toFixed(d),k=Math.pow(10,d),n2=Math.round(Math.abs(x)*2*k);
  if(n2%2===1&&Math.abs(Math.abs(x)*2*k-n2)<1e-6){
    var den=2*k/_gcd(n2,2*k);
    if((den&(den-1))===0){var lo=Math.floor(Math.abs(x)*k+1e-9),
        n=lo%2===0?lo:lo+1;s=((x<0?-1:1)*n/k).toFixed(d);}}
  return s;};
P.pyRound=function(x,d){return Number(P.fx(x,d||0));};
function _comma(s){return s.replace(/\B(?=(\d{3})+(?!\d))/g,',');}
P.fix=function(v,d){var s=P.fx(v,d),neg=s[0]==='-';if(neg)s=s.slice(1);
  var p=s.split('.');return (neg?'-':'')+_comma(p[0])+(d?'.'+p[1]:'');};
P.int=function(v){return P.fix(v,0);};
P.pc1=function(v){return v===null||v===undefined?'-':P.fx(v*100,1)+'%';};
P.fmt=function(h,v){if(v===null||v===undefined||v==='')return '';
  var rate=v&&typeof v==='object';if(rate)v=v.r;
  if(v===null||v===undefined)return '';
  if(typeof v!=='number')return String(v);
  if((CFG.not_pct||[]).indexOf(h)>=0)return P.fix(v,1);
  if((rate||!Number.isInteger(v))&&v>=0&&v<=1)return P.fx(v*100,1)+'%';
  return Number.isInteger(v)?P.int(v):P.fix(v,1);};
P.table=function(hd,rows){var b=[];
  rows.forEach(function(row){var t=[];
    hd.forEach(function(h,j){var v=row[j],rate=v&&typeof v==='object',
        raw=rate?v.r:v,num=typeof raw==='number',
        cell=P.esc(P.fmt(h,v)),H=String(h).toUpperCase();
      if(num&&raw!==null&&(rate||!Number.isInteger(raw))&&raw>=0&&raw<=1
         &&(H.indexOf('%')>=0||H.indexOf('RATE')>=0)){
        var w=Math.max(2,Math.round(raw*54)),
            cl=raw>=.85?'':(raw>=.75?' mid':' low');
        cell+='<span class="bar'+cl+'" style="width:'+w+'px"></span>';}
      t.push('<td class="'+(num&&raw!==null?'num':'')+'">'+cell+'</td>');});
    b.push('<tr>'+t.join('')+'</tr>');});
  return '<div class="scroll'+(rows.length>18?' tall':'')+'"><table><thead>'
    +'<tr>'+hd.map(function(h){return '<th>'+P.esc(h)+'</th>';}).join('')
    +'</tr></thead><tbody>'+b.join('')+'</tbody></table></div>';};
P.EMPTY='<p class="fempty">Nothing matches these filters.</p>';
// ---- charts, matching charts.py
var W=940;
// size: optional [bar height, gap], as charts.bars(bar_h=, gap=)
P.bars=function(rows,hit,base,cols,size){hit=hit||'available';base=base||'expected';
  var FILL=cols?cols[0]:CFG.avail,TRACK=cols?cols[1]:CFG.rest;
  if(!rows.length)return P.EMPTY;
  var LW=210,BH=size?size[0]:19,GP=size?size[1]:7,PT=40,
      h=PT+rows.length*(BH+GP)+16,pw=W-LW-78,
      o=['<svg class="chart" viewBox="0 0 '+W+' '+h+'" role="img">'];
  [0,25,50,75,100].forEach(function(t){var x=LW+pw*t/100;
    o.push('<line class="gl'+(t===0?' z':'')+'" x1="'+x.toFixed(1)+'" y1="'
      +(PT-8)+'" x2="'+x.toFixed(1)+'" y2="'+(h-14)+'"/><text x="'
      +x.toFixed(1)+'" y="'+(h-3)+'" class="cn" text-anchor="middle">'+t
      +'%</text>');});
  rows.forEach(function(r,i){var y=PT+i*(BH+GP),v=r[1]?r[2]/r[1]:0,
      w=Math.min(v,1);
    o.push('<text x="'+(LW-10)+'" y="'+(y+BH-5)+'" class="ct" '
      +'text-anchor="end">'+P.esc(r[0])+'</text>');
    o.push('<rect x="'+LW+'" y="'+y+'" width="'+pw.toFixed(1)+'" height="'
      +BH+'" fill="'+TRACK+'"/>');
    o.push('<rect class="seg" x="'+LW+'" y="'+y+'" width="'+(pw*w).toFixed(1)
      +'" height="'+BH+'" fill="'+FILL+'"><title>'+P.esc(r[0])+': '
      +P.pc1(v)+' '+P.esc(hit)+' - '+P.int(r[2])+' of '+P.int(r[1])+' '
      +P.esc(base)+'</title></rect>');
    o.push(w<0.86
      ? '<text x="'+(LW+pw*w+6).toFixed(1)+'" y="'+(y+BH-5)+'" class="cv">'
        +P.pc1(v)+'</text>'
      : '<text x="'+(LW+pw*w-6).toFixed(1)+'" y="'+(y+BH-5)
        +'" class="cv" text-anchor="end">'+P.pc1(v)+'</text>');
    o.push('<text x="'+(W-6)+'" y="'+(y+BH-5)+'" class="cn" '
      +'text-anchor="end">n='+P.int(r[1])+'</text>');});
  return o.join('')+'</svg>';};
P.stacked=function(rows,series,cols){cols=cols||CFG.stack_cols;
  if(!rows.length)return '<p class="fempty">Nothing to show.</p>';
  var LW=210,BH=22,GP=8,PT=46,h=PT+rows.length*(BH+GP)+18,pw=W-LW-78,
      o=['<svg class="chart" viewBox="0 0 '+W+' '+h+'" role="img">'];
  series.forEach(function(nm,i){var x=LW+i*160;
    o.push('<rect x="'+x+'" y="'+(PT-30)+'" width="11" height="11" rx="2" '
      +'fill="'+cols[i%cols.length]+'"/><text x="'+(x+16)+'" y="'+(PT-20)
      +'" class="cl">'+P.esc(nm)+'</text>');});
  [0,25,50,75,100].forEach(function(t){var x=LW+pw*t/100;
    o.push('<line class="gl'+(t===0?' z':'')+'" x1="'+x.toFixed(1)+'" y1="'
      +(PT-8)+'" x2="'+x.toFixed(1)+'" y2="'+(h-16)+'"/><text x="'
      +x.toFixed(1)+'" y="'+(h-4)+'" class="cn" text-anchor="middle">'+t
      +'%</text>');});
  rows.forEach(function(r,i){var y=PT+i*(BH+GP),v=r[1],
      tot=v.reduce(function(a,b){return a+b;},0)||1,x=LW;
    o.push('<text x="'+(LW-10)+'" y="'+(y+BH-6)+'" class="ct" '
      +'text-anchor="end">'+P.esc(r[0])+'</text>');
    v.forEach(function(c,j){var w=pw*c/tot;if(w<=0)return;
      o.push('<rect class="seg" x="'+x.toFixed(1)+'" y="'+y+'" width="'
        +w.toFixed(1)+'" height="'+BH+'" fill="'+cols[j%cols.length]
        +'"><title>'+P.esc(r[0])+' - '+P.esc(series[j])+': '+P.int(c)
        +' of '+P.int(tot)+' ('+P.pc1(c/tot)+')</title></rect>');
      if(c/tot>=0.07)o.push('<text x="'+(x+w/2).toFixed(1)+'" y="'
        +(y+BH-7).toFixed(1)+'" text-anchor="middle" style="font-size:10.5px;'
        +'font-weight:700;fill:#fff">'+P.fx(c/tot*100,1)+'%</text>');
      x+=w;});
    o.push('<text x="'+(W-6)+'" y="'+(y+BH-6)+'" class="cn" '
      +'text-anchor="end">n='+P.int(tot)+'</text>');});
  return o.join('')+'</svg>';};
// rows [label, n, mean], matching charts.day_bars()
P.dayBars=function(rows,maxV,unit,cols){maxV=maxV||92;unit=unit||'days';
  var FILL=cols?cols[0]:CFG.avail,TRACK=cols?cols[1]:CFG.rest;
  if(!rows.length)return P.EMPTY;
  var LW=210,BH=19,GP=7,PT=40,h=PT+rows.length*(BH+GP)+16,pw=W-LW-78,
      o=['<svg class="chart" viewBox="0 0 '+W+' '+h+'" role="img">'],
      ticks=maxV===92?[0,15,30,45,60,75,90]:[0,1,2,3,4].map(function(k){
        return Math.round(maxV*k/4);});
  ticks.forEach(function(t){var x=LW+pw*t/maxV;
    o.push('<line class="gl'+(t===0?' z':'')+'" x1="'+x.toFixed(1)+'" y1="'
      +(PT-8)+'" x2="'+x.toFixed(1)+'" y2="'+(h-14)+'"/><text x="'
      +x.toFixed(1)+'" y="'+(h-3)+'" class="cn" text-anchor="middle">'+t
      +'</text>');});
  rows.forEach(function(r,i){var y=PT+i*(BH+GP),v=r[2],
      w=maxV?Math.min(v/maxV,1):0,txt=P.fx(v,1)+' '+unit;
    o.push('<text x="'+(LW-10)+'" y="'+(y+BH-5)+'" class="ct" '
      +'text-anchor="end">'+P.esc(r[0])+'</text>');
    o.push('<rect x="'+LW+'" y="'+y+'" width="'+pw.toFixed(1)+'" height="'
      +BH+'" fill="'+TRACK+'"/>');
    o.push('<rect class="seg" x="'+LW+'" y="'+y+'" width="'+(pw*w).toFixed(1)
      +'" height="'+BH+'" fill="'+FILL+'"><title>'+P.esc(r[0])+': '
      +P.fx(v,1)+' '+P.esc(unit)+' on average, over '+P.int(r[1])
      +' pairs</title></rect>');
    o.push(w<0.80
      ? '<text x="'+(LW+pw*w+6).toFixed(1)+'" y="'+(y+BH-5)+'" class="cv">'
        +P.esc(txt)+'</text>'
      : '<text x="'+(LW+pw*w-6).toFixed(1)+'" y="'+(y+BH-5)
        +'" class="cv" text-anchor="end">'+P.esc(txt)+'</text>');
    o.push('<text x="'+(W-6)+'" y="'+(y+BH-5)+'" class="cn" '
      +'text-anchor="end">n='+P.int(r[1])+'</text>');});
  return o.join('')+'</svg>';};
// rate_by(): [key, facilities, n, hits, rate|null] per group, keys sorted
// (or in `keys` order), groups under MINN get a null rate; sort=true
// orders by rate ascending, nulls last, ties keeping key order.
P.rateBy=function(idx,keyfn,hitfn,facfn,keys,minN,sort){
  var mn=minN===undefined||minN===null?P.MINN:minN,b=P.by(idx,keyfn),out=[];
  (keys||b.keys.slice().sort(P.pyStrCmp)).forEach(function(k){
    var rs=b.g[k];if(!rs||!rs.length)return;
    var hit=P.cnt(rs,hitfn);
    out.push([k,P.uniq(rs,facfn),rs.length,hit,
              rs.length>=mn?hit/rs.length:null]);});
  if(sort!==false)out=P.sort(out,function(a,b){
    return (a[4]===null)-(b[4]===null)||(a[4]===null?0:a[4]-b[4]);});
  return out;};
// _bar_rows(): rate_by rows -> bars rows, suppressed groups dropped
P.barRows=function(t,top){var o=t.filter(function(r){return r[0]&&r[4]!==null;})
  .map(function(r){return [r[0],r[2],r[3]];});return top?o.slice(0,top):o;};
// _worst_first(): highest rate first
P.worstFirst=function(t){return P.barRows(P.sort(t,function(a,b){
  return (a[4]===null)-(b[4]===null)||(a[4]===null?0:b[4]-a[4]);}));};
// dial cards, matching charts.dial()/dials()
var A0=150,SW=240;
function _dpt(cx,cy,r,t){var a=(A0+SW*t)*Math.PI/180;
  return [cx+r*Math.cos(a),cy+r*Math.sin(a)];}
function _darc(cx,cy,r,t0,t1){var a=_dpt(cx,cy,r,t0),b=_dpt(cx,cy,r,t1);
  return 'M'+a[0].toFixed(1)+','+a[1].toFixed(1)+' A'+r.toFixed(1)+','
    +r.toFixed(1)+' 0 '+((t1-t0)*SW>180?1:0)+' 1 '+b[0].toFixed(1)+','
    +b[1].toFixed(1);}
P.dial=function(value,caption){var w=272,h=226,T=41,cx=w/2,cy=h/2-6,
    rb=w/2-58,rt0=rb+16,rt1=rb+28,v=Math.max(0,Math.min(1,value)),
    o=['<svg viewBox="0 0 '+w+' '+h+'" role="img" aria-label="'
      +(value*100).toFixed(1)+' percent of medicines available">'];
  for(var i=0;i<T;i++){var t=i/(T-1),a=_dpt(cx,cy,rt0,t),b=_dpt(cx,cy,rt1,t);
    o.push('<line x1="'+a[0].toFixed(1)+'" y1="'+a[1].toFixed(1)+'" x2="'
      +b[0].toFixed(1)+'" y2="'+b[1].toFixed(1)+'" stroke="'
      +(t<=v+1e-9?CFG.g_on:CFG.g_off)+'" stroke-width="'+(i%5?3.2:4.4)
      +'" stroke-linecap="round"/>');}
  o.push('<path d="'+_darc(cx,cy,rb,0,1)+'" fill="none" stroke="'+CFG.g_off
    +'" stroke-width="13" stroke-linecap="round" opacity=".55"/>');
  if(v>0)o.push('<path d="'+_darc(cx,cy,rb,0,v)+'" fill="none" stroke="'
    +CFG.g_on+'" stroke-width="13" stroke-linecap="round"/>');
  [[0,'0'],[.25,'25'],[.5,'50'],[.75,'75'],[1,'100']].forEach(function(s){
    var p=_dpt(cx,cy,rt1+15,s[0]);o.push('<text x="'+p[0].toFixed(1)+'" y="'
      +(p[1]+4).toFixed(1)+'" class="dscale" text-anchor="middle">'+s[1]
      +'</text>');});
  var nd=_dpt(cx,cy,rb-9,v);
  o.push('<line x1="'+cx.toFixed(1)+'" y1="'+cy.toFixed(1)+'" x2="'
    +nd[0].toFixed(1)+'" y2="'+nd[1].toFixed(1)+'" stroke="'+CFG.g_on
    +'" stroke-width="3.2" stroke-linecap="round"/>');
  o.push('<circle cx="'+cx.toFixed(1)+'" cy="'+cy.toFixed(1)+'" r="7.5" '
    +'fill="#141c2e" stroke="'+CFG.g_on+'" stroke-width="3.4"/>');
  o.push('<text x="'+cx.toFixed(0)+'" y="'+(cy+52).toFixed(0)+'" class="dbig" '
    +'text-anchor="middle">'+P.fx(value*100,1)+'%</text>');
  o.push('<text x="'+cx.toFixed(0)+'" y="'+(cy+70).toFixed(0)+'" class="dcap" '
    +'text-anchor="middle">'+P.esc(caption)+'</text>');
  return o.join('')+'</svg>';};
// rows [label, n, hits, subtitle?]; solo for a single centred card
P.dials=function(rows,caption,noun,solo){
  if(!rows.length)return P.EMPTY;
  return '<div class="dgrid'+(solo?' solo':'')+'">'+rows.map(function(r){
    var v=r[1]?r[2]/r[1]:0;return '<div class="dcard">'+P.dial(v,caption)
      +'<p class="dname">'+P.esc(r[0])+'</p><p class="dcount"><b>'
      +P.int(r[2])+'</b> of '+P.int(r[1])+' '+P.esc(noun)+(r[3]||'')
      +'</p></div>';}).join('')+'</div>';};
// ---- putting it on the page
P.slot=function(key){
  var s=document.querySelector('section[data-k="'+key+'"]');if(!s)return null;
  for(var i=s.children.length-1;i>=0;i--){var c=s.children[i],t=c.tagName;
    if(t!=='H4'&&t!=='P')return c;}
  return null;};
P.put=function(key,h){
  if(window.__PF_CAPTURE){window.__PF_CAPTURE[key]=h;return;}  // tests
  var c=P.slot(key);if(!c)return;
  var d=document.createElement('div');d.innerHTML=h;
  var n=d.firstElementChild&&d.children.length===1?d.firstElementChild:null;
  if(n&&c.tagName===n.tagName&&c.className===n.className){
    c.innerHTML=n.innerHTML;return;}
  if(c.tagName==='DIV'&&!c.className){c.innerHTML=h;return;}
  c.outerHTML='<div>'+h+'</div>';};
P._fns=[];
P.render=function(fn){P._fns.push(fn);};
P.run=function(){P._compute();P._fns.forEach(function(fn){fn();});
  var e=document.getElementById('pf'+CFG.key+'-sum');
  if(e)e.textContent='';};
P._suggest=function(id,names){var inp=document.getElementById(id);
  if(!inp||inp.dataset.sg)return;inp.dataset.sg='1';
  var box=document.createElement('div');
  box.className='sugg';inp.parentNode.appendChild(box);
  var seen={},uniq=[];names.forEach(function(n){if(!seen[n]){seen[n]=1;
    uniq.push(n);}});
  uniq.sort(function(a,b){return a.toLowerCase()<b.toLowerCase()?-1:1;});
  function hide(){box.className='sugg';}
  function show(){var q=inp.value.trim().toLowerCase();if(!q){hide();return;}
    var hit=uniq.filter(function(n){return P.starts(n,q);});
    hit.sort(function(a,b){var x=a.toLowerCase().indexOf(q)===0?0:1,
      y=b.toLowerCase().indexOf(q)===0?0:1;return x-y;});
    if(!hit.length){box.innerHTML='<p class="snone">No name starts with '
      +'that</p>';box.className='sugg on';return;}
    box.innerHTML=hit.slice(0,14).map(function(n){return '<button '
      +'type="button" data-n="'+P.esc(n)+'">'+P.esc(n)+'</button>';}).join('')
      +(hit.length>14?'<p class="snone">'+(hit.length-14)
      +' more - keep typing</p>':'');
    box.className='sugg on';}
  inp.addEventListener('input',show);inp.addEventListener('focus',show);
  inp.addEventListener('blur',function(){setTimeout(hide,150);});
  inp.addEventListener('keydown',function(e){if(e.key==='Escape')hide();});
  box.addEventListener('mousedown',function(e){
    var b=e.target.closest('button');if(!b)return;e.preventDefault();
    inp.value=b.getAttribute('data-n');hide();
    inp.dispatchEvent(new Event('input'));});};
P.start=function(){var i=CFG.ids;
  ['prov','dist','hub','lvl','prg','tag'].forEach(function(k){
    var e=i[k]&&document.getElementById(i[k]);
    if(e)e.addEventListener('change',P.run);});
  ['fac','med'].forEach(function(k){var e=document.getElementById(i[k]);
    if(e)e.addEventListener('input',P.run);});
  var m=document.getElementById('fx-meth');
  if(m&&CFG.key==='a1')m.addEventListener('change',P.run);
  var r=document.getElementById('pf'+CFG.key+'-reset');
  if(r)r.addEventListener('click',function(){
    ['prov','dist','hub','lvl','prg','fac','med','tag'].forEach(function(k){
      var e=i[k]&&document.getElementById(i[k]);if(e)e.value='';});P.run();});
  if(CFG.own){P._suggest(i.fac,P.F.map(function(f){return f[0];}));
    P._suggest(i.med,P.M.map(function(m){return m[0];}));}
  P.run();};
"""
