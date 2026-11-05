"""Multi-select for the filter bars (user request 2026-10-26).

Every Province / RMS hub / Facility level / Programme <select class="ms">
on every page becomes a button with a checkbox list. The <select multiple>
stays in the page, hidden, as the one source of truth: ticking a box sets
that option's .selected and fires the select's own 'change', so the three
filter engines (explorer.py, stock_explorer.py, pagefilter.py) keep their
listeners and read the ticked values with msVals(). Nothing ticked = all.

A Reset sets select.value = '' - for a multiple select that selects only
the "All ..." option, which msVals() ignores - so the widget only has to
redraw its labels after any .freset click.
"""

CSS = """
.ms-wrap{position:relative}
.ms-wrap select.ms{display:none}
.msb{width:100%;font:inherit;font-size:13px;padding:7px 26px 7px 9px;
 color:var(--ink);background:#141413;border:1px solid var(--rule);
 border-radius:6px;text-align:left;cursor:pointer;white-space:nowrap;
 overflow:hidden;text-overflow:ellipsis;position:relative}
.msb::after{content:"";position:absolute;right:10px;top:50%;margin-top:-3px;
 border:4px solid transparent;border-top-color:var(--ink2)}
.msb.on{border-color:var(--accent)}
.msp{display:none;position:absolute;z-index:40;left:0;min-width:100%;
 top:100%;margin-top:2px;background:#141413;border:1px solid var(--rule);
 border-radius:6px;max-height:300px;overflow:auto;padding:4px 0;
 box-shadow:0 6px 18px rgba(0,0,0,.5)}
.msp.open{display:block}
.msp label{display:flex;gap:8px;align-items:center;margin:0;padding:5px 10px;
 font-size:12.5px;color:var(--ink);text-transform:none;letter-spacing:0;
 cursor:pointer;white-space:nowrap}
.msp label:hover{background:#2a2a28}
.msp input{width:auto;margin:0;accent-color:var(--accent)}
.msp .msc{display:block;width:100%;text-align:left;font:inherit;
 font-size:12px;padding:6px 10px;margin-top:2px;background:none;border:0;
 border-top:1px solid var(--rule);color:var(--accent);cursor:pointer}
"""

SCRIPT = r"""<script>
(function(){
// the ticked values of a filter select, lower-cased; [] = all
window.msVals=function(id){var e=document.getElementById(id);if(!e)return [];
  return Array.prototype.filter.call(e.options,function(o){
    return o.selected&&o.value;}).map(function(o){
    return o.value.trim().toLowerCase();});};
// true if x passes a msVals() list
window.msHas=function(a,x){return !a.length||
  a.indexOf(String(x).trim().toLowerCase())>=0;};
var all=[];
function label(s){var on=Array.prototype.filter.call(s.options,function(o){
    return o.selected&&o.value;});
  return on.length?(on.length<=2?on.map(function(o){return o.text;})
    .join(', '):on.length+' selected'):s.options[0].text;}
function build(s){
  s.multiple=true;
  Array.prototype.forEach.call(s.options,function(o){o.selected=false;});
  var w=document.createElement('div');w.className='ms-wrap';
  s.parentNode.insertBefore(w,s);w.appendChild(s);
  var b=document.createElement('button');b.type='button';b.className='msb';
  b.setAttribute('aria-haspopup','listbox');
  var p=document.createElement('div');p.className='msp';
  w.appendChild(b);w.appendChild(p);
  var boxes=[];
  Array.prototype.forEach.call(s.options,function(o){if(!o.value)return;
    var l=document.createElement('label'),c=document.createElement('input');
    c.type='checkbox';c.value=o.value;
    c.addEventListener('change',function(){o.selected=c.checked;
      if(s.options[0]&&!s.options[0].value)s.options[0].selected=false;
      sync();s.dispatchEvent(new Event('change'));});
    l.appendChild(c);l.appendChild(document.createTextNode(o.text));
    p.appendChild(l);boxes.push([c,o]);});
  var clr=document.createElement('button');clr.type='button';
  clr.className='msc';clr.textContent='Clear';p.appendChild(clr);
  clr.addEventListener('click',function(){
    Array.prototype.forEach.call(s.options,function(o){o.selected=false;});
    sync();s.dispatchEvent(new Event('change'));});
  function sync(){boxes.forEach(function(x){x[0].checked=x[1].selected&&
      !!x[1].value;});
    b.textContent=label(s);
    b.classList.toggle('on',boxes.some(function(x){return x[1].selected;}));}
  b.addEventListener('click',function(ev){ev.stopPropagation();
    var o=p.classList.contains('open');
    document.querySelectorAll('.msp.open').forEach(function(x){
      x.classList.remove('open');});
    if(!o)p.classList.add('open');});
  p.addEventListener('click',function(ev){ev.stopPropagation();});
  var lab=document.querySelector('label[for="'+s.id+'"]');
  if(lab)lab.addEventListener('click',function(ev){ev.preventDefault();
    b.click();});
  sync();all.push(sync);}
function init(){document.querySelectorAll('select.ms').forEach(build);}
document.addEventListener('click',function(ev){
  document.querySelectorAll('.msp.open').forEach(function(x){
    x.classList.remove('open');});
  // the Reset buttons' own handlers have already cleared the selects
  if(ev.target.closest&&ev.target.closest('.freset'))
    all.forEach(function(f){f();});});
if(document.readyState==='loading')
  document.addEventListener('DOMContentLoaded',init);else init();
})();
</script>"""
