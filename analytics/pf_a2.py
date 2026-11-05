"""Page filter for Analysis 2 - stock adequacy and inventory accuracy.

FIG 2.1-2.11 and TABLE 2A/2B are redrawn by stock_explorer.py's own script;
this module follows that page's filter bar (sx-prov ... sx-med) and redraws
the two indicator tables below them:

    2|2.1      months of stock - distribution and stock status
               (SA_OUT / SA_IN by R._band / AMC_ZERO)
    2|2.2-2.4  inventory record accuracy (_acc, carried on the same rows)

Every Python float in a table is passed as P.R(v) so it prints exactly as
fmt() prints a float (a whole 0.0 or 1.0 is a percentage there).
"""


def _int(v):
    """A whole float as an int: the same number in JavaScript, fewer bytes."""
    return int(v) if isinstance(v, float) and v.is_integer() else v


def build(pg, R):
    M = R.M
    # both tables are hidden from the dashboard (HTML_DROP, 2026-10-16):
    # nothing to redraw, so embed nothing
    if {(2, "2.1"), (2, "2.2-2.4")} <= R.HTML_DROP:
        return None
    out_ids = {id(r) for r in R.SA_OUT}
    zero_ids = {id(r) for r in R.AMC_ZERO}
    rows = R.SA_OUT + R.SA_IN + R.AMC_ZERO
    # k: 0 stocked out, 1-3 on the shelf under/adequate/over, 4 excluded
    pg.dataset("st", rows, R.pf_fac, R.pf_med, {
        "k": lambda r: 0 if id(r) in out_ids else
        (4 if id(r) in zero_ids else R._band(r) + 1)})
    # months of stock at full precision: the median is rounded to 2 places
    # after interpolation, so the 6-place column rounding could move it
    pg.D["st"]["c"]["mos"] = [
        None if id(r) in out_ids or id(r) in zero_ids else R._mos(r)
        for r in rows]
    # _acc is a subset of the on-shelf rows, so its two counts ride on the
    # same rows (null = not in _acc) instead of a second facility/medicine
    # index. A missing count is stored as 0, which is what (num or 0) reads.
    acc = {id(r) for r in R._acc}
    assert acc <= {id(r) for r in rows}

    def qty(col):
        return [_int((M.num(r[col]) or 0)) if id(r) in acc else None
                for r in rows]
    pg.D["st"]["c"]["ph"] = qty(M.QTY_PHYS)
    pg.D["st"]["c"]["sy"] = qty(M.QTY_SYS)
    return JS


JS = r"""
function med2(v){v=v.filter(function(x){return x!==null&&x!==undefined;});
  if(!v.length)return null;v=v.slice().sort(function(a,b){return a-b;});
  var m=v.length>>1,x=v.length%2?v[m]:(v[m-1]+v[m])/2;
  return P.pyRound(x,2);}
P.render(function(){
  // 2.1
  var ix=P.rows('st'),K=P.D.st.c.k,MO=P.D.st.c.mos,n=[0,0,0,0,0],mos=[];
  ix.forEach(function(i){n[K[i]]++;if(K[i]>=1&&K[i]<=3)mos.push(MO[i]);});
  var base=n[0]+n[1]+n[2]+n[3],pc=function(a){return P.R(P.pct(a,base));};
  P.put('2|2.1',P.table(['Stock category','Definition','Observations',
    '% of base'],[
    ['Stocked out','not available on the visit day',n[0],pc(n[0])],
    ['Understock','on the shelf, under 1 month of stock',n[1],pc(n[1])],
    ['Adequate stock','1 to 2 months (1 to 3 at teaching and referral hospitals)',n[2],pc(n[2])],
    ['Overstock','more than 2 months (3 at teaching and referral hospitals)',n[3],pc(n[3])],
    ['','','',null],
    ['BASE','every stock-out, plus stocked rows with consumption',base,P.R(1)],
    ['Excluded: on the shelf with zero consumption',
     'months of stock is undefined at AMC = 0 and the medicine was present',
     n[4],null],
    ['Median months of stock','stocked rows only',P.R(med2(mos)),null]]));
  // 2.2-2.4
  var PH=P.D.st.c.ph,SY=P.D.st.c.sy,
      ax=ix.filter(function(i){return PH[i]!==null;}),ex=0,gt=0,lt=0,
      diff=[],ape=[];
  ax.forEach(function(i){var p=PH[i]||0,d=p-(SY[i]||0);
    if(d===0)ex++;if(d>0)gt++;if(d<0)lt++;diff.push(d);
    if(p>0)ape.push(Math.abs(d)/(PH[i]||1));});
  P.put('2|2.2-2.4',P.table(['Metric','Value','Basis'],[
    ['Observations with both quantities',ax.length,
     'available rows only - counts are not recorded where the medicine is '
     +'absent'],
    ['Physical equals recorded (exact accuracy)',ex,''],
    ['EXACT ACCURACY RATE',P.R(P.pct(ex,ax.length)),''],
    ['Physical exceeds recorded',gt,'the shelf holds more than the system'],
    ['Recorded exceeds physical',lt,
     'the system overstates - hides a shortage, reorder never triggers'],
    ['Median discrepancy (physical - recorded)',P.R(med2(diff)),'units'],
    ['Median absolute percentage error',P.R(med2(ape)),
     '|physical - recorded| / physical, physical > 0'],
    ['Within +/- 5%',ape.filter(function(e){return e<=0.05;}).length,
     'of '+ape.length+' rows with stock > 0']]));
});
"""
