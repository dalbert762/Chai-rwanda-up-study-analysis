"""Two-page summary brief of the indicator dashboard.

    python3 Summary_brief.py      -> Summary_brief.html

Modelled on a policy-brief layout: a headline strip, numbered thematic
sections of small charts, a district-by-district table, priorities for
action, data gaps and numbered sources. Every number is read from
Full_indicator_report's row sets, so a rebuild of the dashboard and of this
brief cannot disagree. Print to PDF from the browser (A4, two pages).

openpyxl + the standard library, like the rest of this project.
"""
import _bootstrap as B   # path setup - must come first
import contextlib
import html
import io
import os
import statistics
from collections import Counter, defaultdict

with contextlib.redirect_stdout(io.StringIO()):
    import Full_indicator_report as R

M, C, s, pct = R.M, R.C, R.s, R.pct
OUT = B.out("Summary_brief.html")

TEAL, TEAL_L, ORANGE, INK, MUTED = ("#1f7a72", "#8cc5bd", "#c1622c",
                                    "#1d2b2c", "#5d6b6b")
CARD = "#eef2f1"


def esc(x):
    return html.escape("" if x is None else str(x), quote=True)


def pc(v, d=1):
    return "-" if v is None else f"{v * 100:.{d}f}%"


def num(v):
    return f"{v:,.0f}"


def mean(v):
    v = list(v)
    return statistics.mean(v) if v else None


# ---------------------------------------------------------------
# numbers
# ---------------------------------------------------------------
LV = [l for l in R.LEVELS]
SHORT = {"Health post": "Health post", "Health centre": "Health centre",
         "District hospital": "District hospital",
         "Provincial hospital": "Provincial hosp.",
         "Level 2 teaching hospital": "L2 teaching hosp.",
         "National referral hospital": "National referral"}


def avail(rows):
    a = [r for r in rows if not M.blank(r[M.AVAIL])]
    return pct(sum(1 for r in a if M.yes(r[M.AVAIL])), len(a)), len(a)


def by(rows, keyfn):
    g = defaultdict(list)
    for r in rows:
        g[keyfn(r)].append(r)
    return g


lvl = lambda r: r["_lvl"]
prov = lambda r: s(r["_p"][M.PROV])
dist = lambda r: s(r["_p"][M.DIST])
prog = lambda r: s(r[M.PROGRAM]) or "(none)"

AV_B, N_B = avail(R.MOLS_B)          # managed products only (default)
AV_A, N_A = avail(R.MOLS)            # expected medicines
N_FAC = len(R.M.FACILITIES)
LVL_FAC = Counter({s(r["_p"][M.FAC]): r["_lvl"] for r in R.MED_ALL}.values())

av_lvl = []
for l in LV:
    a, na = avail(by(R.MOLS, lvl).get(l, []))
    b, nb = avail(by(R.MOLS_B, lvl).get(l, []))
    if nb:
        av_lvl.append((l, a, b, LVL_FAC[l]))
av_prog = sorted(((k, *avail(v)) for k, v in by(R.MOLS_B, prog).items()
                  if k != "(none)" and len(v) >= R.MIN_N),
                 key=lambda x: x[1])
av_prov = sorted(((k, *avail(v)) for k, v in by(R.MOLS_B, prov).items()),
                 key=lambda x: x[1])

# stock adequacy (FIG 2.1 base)
MIX = R.adequacy_mix4()
MIX_N = sum(MIX)
N_EOP = sum(1 for r in R.SA_SHELF if R._eop(r))
mix_lvl = [(l, R.adequacy_mix4(lambda r, l=l: r["_lvl"] == l)) for l in LV]
mix_lvl = [(l, v) for l, v in mix_lvl if sum(v)]

# record accuracy
ACC = R._acc
EXACT = pct(sum(1 for r in ACC if R._exact(r)), len(ACC))
QACC = mean(R._qacc(r) for r in ACC)
N_SYS0 = sum(1 for r in ACC if (M.num(r[M.QTY_SYS]) or 0) <= 0)
acc_lvl = []
for l, rs in ((l, by(ACC, lvl).get(l, [])) for l in LV):
    if len(rs) >= R.MIN_N:
        acc_lvl.append((l, pct(sum(1 for r in rs if R._exact(r)), len(rs)),
                        mean(R._qacc(r) for r in rs), len(rs)))

# supply chain
UPB = R.UP_BOTH
N_GAP = sum(1 for r in UPB if s(r[M.DIST_FLAG]) == R.DN_ISSUE_T)
N_NAT = sum(1 for r in UPB if s(r[M.DIST_FLAG]) == R.DN_NS_T)
up_lvl = [(k, v[0][0], v[0][1], v[1][1]) for k, v in R.up_lvl_both]
SO_INC = pct(len(R._any_so), len(R.MG))
so_prog = sorted(((k, pct(sum(1 for r in v if R._somonths(r)), len(v)),
                   len(v)) for k, v in by(R.MG, prog).items()
                  if k != "(none)" and len(v) >= R.MIN_N),
                 key=lambda x: -x[1])
RQ = sum(r["_rq"] for r in R.MG)
DV = sum(r["_dv"] for r in R.MG)
FILL = pct(DV, RQ)
fill_lvl = []
for l in LV:
    rs = by(R.MG, lvl).get(l, [])
    q = sum(r["_rq"] for r in rs)
    if q:
        fill_lvl.append((l, sum(r["_dv"] for r in rs) / q))
LOSS = R._loss_tot
UNSERVED_VAL = R._val

# prescriptions and community
N_PRES = len(R.PRES)
N_PRES_UN = len(R._pres_un)
RX_OK = pct(sum(1 for r in R.RX_ANS if R._rx_ok(r)), len(R.RX_ANS))
N_RX_FAC = len({R.pfac(r) for r in R.PRES})
CHW = pct(len(R.CAV), len(R.CANS))
N_CHW = len(C.CHWS)
chw_prov = []
for k, v in by(R.CANS, lambda r: C.s(r["_p"][C.PROV])).items():
    if len(v) >= R.MIN_N:
        chw_prov.append((k, pct(sum(1 for r in v if C.yes(r[C.AVAIL])),
                                len(v)), len(v)))
chw_prov.sort(key=lambda x: x[1])

# districts
DIST = []
avb, mg, ac = by(R.MOLS_B, dist), by(R.MG, dist), by(ACC, dist)
for d in sorted(avb):
    a, _ = avail(avb[d])
    m = mg.get(d, [])
    q = sum(r["_rq"] for r in m)
    rs = ac.get(d, [])
    DIST.append({
        "d": d, "prov": prov(avb[d][0]),
        "fac": len({s(r["_p"][M.FAC]) for r in avb[d]}),
        "av": a,
        "so": pct(sum(1 for r in m if R._somonths(r)), len(m)),
        "acc": pct(sum(1 for r in rs if R._exact(r)), len(rs))
        if len(rs) >= R.MIN_N else None,
        "fill": sum(r["_dv"] for r in m) / q if q else None})
DIST.sort(key=lambda x: -(x["av"] or 0))
NATIONAL = {"av": AV_B, "so": SO_INC, "acc": EXACT, "fill": FILL}
LOWER_IS_BETTER = {"so"}


# ---------------------------------------------------------------
# small SVG charts
# ---------------------------------------------------------------
def hbars(rows, w=470, lw=120, bh=11, gap=5, colour=TEAL, ref=None,
          ref_lab="", vmax=1.0, fmt=pc, note=None):
    """rows: (label, value 0-1, tooltip, colour?)"""
    top = 18 if ref is not None else 4
    h = top + len(rows) * (bh + gap) + 4
    pw = w - lw - 48
    o = [f'<svg viewBox="0 0 {w} {h}" class="ch" role="img">']
    for i, row in enumerate(rows):
        lab, v, tip = row[:3]
        col = row[3] if len(row) > 3 else colour
        y = top + i * (bh + gap)
        bw = pw * max(0, min(v or 0, vmax)) / vmax
        o.append(f'<text x="{lw - 8}" y="{y + bh - 3}" class="lab" '
                 f'text-anchor="end">{esc(lab)}</text>'
                 f'<rect x="{lw}" y="{y}" width="{bw:.1f}" height="{bh}" '
                 f'rx="2" fill="{col}"><title>{esc(tip)}</title></rect>'
                 f'<text x="{lw + bw + 5:.1f}" y="{y + bh - 3}" class="val">'
                 f'{esc(fmt(v))}</text>')
    if ref is not None:
        x = lw + pw * ref / vmax
        o.append(f'<line x1="{x:.1f}" y1="10" x2="{x:.1f}" y2="{h - 2}" '
                 f'class="ref"/><text x="{x:.1f}" y="9" class="refl" '
                 f'text-anchor="middle">{esc(ref_lab)}</text>')
    return "".join(o) + "</svg>"


def pairbars(rows, names, w=470, lw=118, bh=7, gap=7, cols=(TEAL_L, TEAL),
             vmax=1.0):
    """rows: (label, v1, v2) - two thin bars per group."""
    h = 20 + len(rows) * (2 * bh + 2 + gap)
    pw = w - lw - 46
    o = [f'<svg viewBox="0 0 {w} {h}" class="ch" role="img">']
    lx = 4
    for j, nm in enumerate(names):
        o.append(f'<rect x="{lx}" y="2" width="9" height="9" '
                 f'rx="2" fill="{cols[j]}"/><text x="{lx + 13}" '
                 f'y="10" class="lg">{esc(nm)}</text>')
        lx += 13 + len(nm) * 5.2 + 16
    for i, (lab, *vs) in enumerate(rows):
        y = 20 + i * (2 * bh + 2 + gap)
        o.append(f'<text x="{lw - 8}" y="{y + bh + 4}" class="lab" '
                 f'text-anchor="end">{esc(lab)}</text>')
        for j, v in enumerate(vs):
            yy = y + j * (bh + 2)
            bw = pw * max(0, min(v or 0, vmax)) / vmax
            o.append(f'<rect x="{lw}" y="{yy}" width="{bw:.1f}" '
                     f'height="{bh}" rx="2" fill="{cols[j]}"><title>'
                     f'{esc(lab)} - {esc(names[j])}: {pc(v)}</title></rect>'
                     f'<text x="{lw + bw + 4:.1f}" y="{yy + bh - 1}" '
                     f'class="val sm">{pc(v)}</text>')
    return "".join(o) + "</svg>"


MIX_COLS = ["#7a1f33", ORANGE, TEAL, "#d9a21b"]
MIX_NAMES = ["Stocked out", "Understock", "Adequate", "Overstock"]


def stackbars(rows, w=640, lw=118, bh=13, gap=5, legend=True):
    top = 20 if legend else 2
    h = top + len(rows) * (bh + gap)
    pw = w - lw - 6
    o = [f'<svg viewBox="0 0 {w} {h}" class="ch" role="img">']
    if legend:
        lx = 4
        for j, nm in enumerate(MIX_NAMES):
            o.append(f'<rect x="{lx}" y="2" width="9" height="9" '
                     f'rx="2" fill="{MIX_COLS[j]}"/><text x="{lx + 13}"'
                     f' y="10" class="lg">{nm}</text>')
            lx += 13 + len(nm) * 5.2 + 14
    for i, (lab, vals) in enumerate(rows):
        y = top + i * (bh + gap)
        tot = sum(vals) or 1
        x = lw
        o.append(f'<text x="{lw - 8}" y="{y + bh - 3}" class="lab" '
                 f'text-anchor="end">{esc(lab)}</text>')
        for j, v in enumerate(vals):
            bw = pw * v / tot
            o.append(f'<rect x="{x:.1f}" y="{y}" width="{bw:.1f}" '
                     f'height="{bh}" fill="{MIX_COLS[j]}"><title>{esc(lab)} - '
                     f'{MIX_NAMES[j]}: {v:,} of {tot:,} ({pc(v / tot)})'
                     f'</title></rect>')
            if bw >= 22:
                o.append(f'<text x="{x + bw / 2:.1f}" y="{y + bh - 4}" '
                         f'class="inb" text-anchor="middle">'
                         f'{v / tot * 100:.0f}%</text>')
            x += bw
    return "".join(o) + "</svg>"


def splitbar(parts, w=640, h=40):
    """parts: (label, value, colour) - one 100% bar with labels inside."""
    tot = sum(p[1] for p in parts) or 1
    x, o = 0, [f'<svg viewBox="0 0 {w} {h}" class="ch" role="img">']
    for lab, v, col in parts:
        bw = w * v / tot
        o.append(f'<rect x="{x:.1f}" y="0" width="{bw:.1f}" height="{h}" '
                 f'fill="{col}"><title>{esc(lab)}: {v:,} of {tot:,}</title>'
                 f'</rect>')
        if bw > 60:
            o.append(f'<text x="{x + 9:.1f}" y="{h / 2 + 5:.0f}" '
                     f'class="sbig">{v / tot * 100:.0f}%</text>'
                     f'<text x="{x + 46:.1f}" y="{h / 2 + 4:.0f}" '
                     f'class="ssm">{esc(lab)}</text>')
        x += bw
    return "".join(o) + "</svg>"


# ---------------------------------------------------------------
# page parts
# ---------------------------------------------------------------
def sup(*n):
    # the sources list was removed on request (2026-10-16), so the numbered
    # reference markers are no longer printed
    return ""


def card(title, body, foot="", cls=""):
    return (f'<div class="card {cls}"><h3>{title}</h3>{body}'
            + (f'<p class="foot">{foot}</p>' if foot else "") + '</div>')


def section(no, title, tag, body):
    return (f'<section><div class="sh"><h2>{no} &#183; {esc(title)}</h2>'
            f'<span class="tag">{esc(tag)}</span></div>{body}</section>')


def kpi(v, lab, src, n):
    return (f'<div class="kpi"><p class="kv">{v}</p><p class="kl">{lab}</p>'
            f'<p class="ks">{src}{sup(n)}</p></div>')


KPIS = "".join([
    kpi(pc(AV_B), "of the medicines facilities manage were on the shelf",
        f"{N_FAC} facilities · visit day", 1),
    kpi(f"{N_GAP / (N_GAP + N_NAT):.0%}", "of facility stock-outs were "
        "<b>distribution gaps</b>: the stock was at RMS",
        f"{N_GAP} of {N_GAP + N_NAT} answered", 2),
    kpi(pc(MIX[3] / MIX_N), "of medicines overstocked (over 2 months); only "
        f"{pc(MIX[2] / MIX_N)} adequate", f"{MIX_N:,} medicines", 3),
    kpi(pc(EXACT, 0), "of stock records match the shelf exactly",
        f"{pc(QACC, 0)} quantity accuracy", 4),
    kpi(pc(FILL), "of units requested were delivered",
        "Jun–Aug 2026 · order fill", 5),
])

# compare levels with at least 5 facilities, so one hospital is not "a level"
_lv5 = [x for x in av_lvl if x[3] >= 5]
worst_lvl = min(_lv5, key=lambda x: x[2])      # managed products only
best_lvl = max(_lv5, key=lambda x: x[2])
worst_a = min(_lv5, key=lambda x: x[1])        # expected medicines

def points(items):
    """The 'key points' card that opens each section."""
    return ('<div class="card kp"><h3>Key points</h3><ul>'
            + "".join(f"<li>{x}</li>" for x in items) + '</ul></div>')


def B(x):
    return f"<b>{x}</b>"


# --- extra numbers for the key points
_lvl_a = {l: a for l, a, b, _ in av_lvl}
_lvl_b = {l: b for l, a, b, _ in av_lvl}
HP_A, HP_B = _lvl_a.get("Health post"), _lvl_b.get("Health post")
AFF_DAYS = [R._wdays(r) for r in R.MG if R._wdays(r) > 0]
OUT30 = sum(1 for r in R.SO if (M.num(r[M.SO_NOW]) or 0) > 30)
def _short_med(m):
    """'Amlodipine / Losartan/ HCT - Tab. 5mg...' -> 'Amlodipine / Losartan'."""
    m = m.split(" - ")[0].strip()
    return m if len(m) <= 30 else m[:28].rstrip(" /,") + "…"


N_UN_ITEMS = len(R.UNSERVED)
UN_FAC = Counter(R.pfac(r) for r in R.UNSERVED).most_common()
N_UN_FAC = len(UN_FAC)
UN_MED = Counter(R.rx_med_name(r) for r in R.UNSERVED).most_common()
rx_lvl = sorted(((k, pct(sum(1 for r in v if R._rx_ok(r)), len(v)), len(v))
                 for k, v in by(R.RX_ANS, lvl).items() if len(v) >= R.MIN_N),
                key=lambda x: x[1])
fill_prog = []
for k, v in by(R.MG, prog).items():
    q = sum(r["_rq"] for r in v)
    if k != "(none)" and q:
        fill_prog.append((k, sum(r["_dv"] for r in v) / q))
fill_prog.sort(key=lambda x: x[1])
fill_mon = [(m, sum(R._mdel(r, m) for r in R.MG)
             / sum(R._mreq(r, m) for r in R.MG)) for m in R.MONTHS]
N_CHW_ASSESSED = len(R._chw_by)
CHW_COMPLETE = R._chw_complete
CMP_ALL = next((x for x in R._cmp if str(x[0]).upper().startswith("ALL")),
               None)
acc_rows = [d for d in DIST if d["acc"] is not None]
so_rows = [d for d in DIST if d["so"] is not None]
fill_rows = [d for d in DIST if d["fill"] is not None]


def _worse(d):
    """How many of the four district indicators are worse than national."""
    n = 0
    for k in ("av", "acc", "fill"):
        n += d[k] is not None and d[k] < NATIONAL[k]
    n += d["so"] is not None and d["so"] > NATIONAL["so"]
    return n


WEAK = [d["d"] for d in DIST if _worse(d) >= 3]
lo_av, hi_av = DIST[-1], DIST[0]
lo_acc = min(acc_rows, key=lambda d: d["acc"])
hi_acc = max(acc_rows, key=lambda d: d["acc"])
hi_so = max(so_rows, key=lambda d: d["so"])
lo_fill = min(fill_rows, key=lambda d: d["fill"])

S1 = section(1, "Medicine availability", "AVAILABILITY",
    '<div class="g g3">'
    + points([
        f"{B(pc(AV_B))} of the medicines facilities manage were on the "
        f"shelf on the visit day; {B(pc(AV_A))} of the medicines "
        f"<i>expected</i> at their level.",
        f"Health posts hold {B(pc(HP_A))} of expected medicines but "
        f"{B(pc(HP_B))} of what they manage: the gap is authorised "
        f"Essential Medicines they do not stock, not stock-outs.",
        f"Lowest programmes: {av_prog[0][0]} {pc(av_prog[0][1])} and "
        f"{av_prog[1][0]} {pc(av_prog[1][1])}; highest "
        f"{av_prog[-1][0]} {pc(av_prog[-1][1])}.",
        f"Provinces sit in a narrow band, {pc(av_prov[0][1])} "
        f"({av_prov[0][0]}) to {pc(av_prov[-1][1])} ({av_prov[-1][0]})."])
    + card("By level of care, both methods",
           pairbars([(SHORT.get(l, l), a, b) for l, a, b, _ in av_lvl],
                    ["Expected (method 1)", "Managed only (method 2)"],
                    w=330, lw=104),
           "Method 1 counts an Essential Medicine authorised at the level "
           "but not managed as unavailable; method 2 asks only of what the "
           "facility manages." + sup(1))
    + card("By programme",
           hbars([(k, v, f"{k}: {pc(v)} of {n:,} managed pairs",
                   TEAL if v >= AV_B else ORANGE) for k, v, n in av_prog],
                 ref=AV_B, ref_lab=f"national {pc(AV_B)}", w=330, lw=118),
           "Managed products only. Orange = below the national figure."
           + sup(1))
    + '</div>')

_acc5 = [x for x in acc_lvl if LVL_FAC[x[0]] >= 5]   # 5+ facilities
_acc_lo = min(_acc5, key=lambda x: x[1])
_acc_hi = max(_acc5, key=lambda x: x[1])

S2 = section(2, "Stock adequacy and inventory accuracy",
             "STOCK ADEQUACY · RECORDS",
    '<div class="g g3">'
    + points([
        f"Only {B(pc(MIX[2] / MIX_N))} of medicines were adequately "
        f"stocked; {B(pc(MIX[3] / MIX_N))} were overstocked (over 2 "
        f"months) and {pc(MIX[1] / MIX_N)} understocked.",
        f"{B(f'{N_EOP:,}')} understocked medicines ({N_EOP / MIX[1]:.0%}) "
        f"were at or below the 0.5-month emergency order point.",
        f"Only {B(pc(EXACT))} of stock records matched the shelf exactly; "
        f"average quantity accuracy was {B(pc(QACC))}.",
        f"{_acc_lo[0]}s keep the least accurate records "
        f"({pc(_acc_lo[1])} exact, {pc(_acc_lo[2])} quantity); "
        f"{_acc_hi[0].lower()}s the most ({pc(_acc_hi[1])} exact)."])
    + card("Stock adequacy by level of care",
           stackbars([("All facilities", MIX)]
                     + [(SHORT.get(l, l), v) for l, v in mix_lvl],
                     w=330, lw=104),
           f"Months of stock = quantity on hand ÷ average monthly "
           f"consumption (Jun–Aug); adequate 1–2 months (1–3 at teaching and referral "
           f"hospitals). Base {MIX_N:,} medicines, stock-outs included." + sup(3))
    + card("Inventory accuracy by level of care",
           pairbars([(SHORT.get(l, l), e, q) for l, e, q, _ in acc_lvl],
                    ["Exact match", "Quantity accuracy"], w=330, lw=104),
           "Quantity accuracy = 1 − |system − physical| ÷ system per record, "
           f"0% at worst and where the system shows 0. {len(ACC):,} records "
           "naming a digital system." + sup(4))
    + '</div>')

S3 = section(3, "Stock-outs", "STOCK-OUT · SUPPLY CHAIN",
    '<div class="g g3">'
    + points([
        f"{B(pc(SO_INC))} of managed medicines ran out at least once in "
        f"June–August; {B(f'{MIX[0]:,}')} were out on the visit day.",
        f"Stock-outs are long: {B(f'{statistics.mean(AFF_DAYS):.0f} days')}"
        f" on average per affected medicine; {OUT30} of those out on the "
        f"visit day had been out over 30 days.",
        f"{B(f'{N_GAP / (N_GAP + N_NAT):.0%}')} were distribution gaps - "
        f"the medicine was at RMS Central or Regional ({N_GAP} of "
        f"{N_GAP + N_NAT}); only {N_NAT} were nationwide."])
    + card("Distribution gap against nationwide stock-out",
           splitbar([("Distribution gap", N_GAP, TEAL),
                     ("Nationwide", N_NAT, ORANGE)], w=330, h=34)
           + pairbars([(SHORT.get(k, k), reg / n, cen / n)
                       for k, n, reg, cen in up_lvl],
                      ["At RMS Regional", "At RMS Central"],
                      cols=(ORANGE, TEAL), w=330, lw=104),
           f"Unavailable medicines answered Yes/No for both stores "
           f"({len(UPB)})." + sup(2))
    + card("Stock-out incidence by programme",
           hbars([(k, v, f"{k}: {pc(v)} of {n:,} managed pairs",
                   ORANGE if v > SO_INC else TEAL) for k, v, n in so_prog],
                 ref=SO_INC, ref_lab=f"all {pc(SO_INC)}", w=330, lw=118,
                 vmax=max(v for _, v, _ in so_prog) * 1.15),
           "Managed pairs with at least one stock-out day, June–August. "
           "Orange = above the national rate." + sup(1))
    + '</div>')

S4 = section(4, "Prescription fulfilment", "PRESCRIPTIONS",
    '<div class="g g3">'
    + points([
        f"{B(pc(RX_OK))} of {len(R.RX_ANS):,} prescribed items were "
        f"available; {B(N_PRES_UN)} of {N_PRES:,} prescriptions "
        f"({pc(N_PRES_UN / N_PRES)}) were not fully served.",
        f"Lowest at {rx_lvl[0][0].lower()}s ({pc(rx_lvl[0][1])}); every "
        f"other level is above {pc(min(v for _, v, _ in rx_lvl[1:]), 0)}.",
        f"The {N_UN_ITEMS} unserved items came from only {N_UN_FAC} "
        f"facilities; {B(UN_FAC[0][0])} alone accounts for "
        f"{B(UN_FAC[0][1])} ({UN_FAC[0][1] / N_UN_ITEMS:.0%}).",
        "Prescribers tend to write what the pharmacy holds, so this rate "
        "overstates supply and is not a stock measure."])
    + card("Prescribed items available, by level of care",
           hbars([(SHORT.get(k, k), v, f"{k}: {pc(v)} of {n:,} items",
                   TEAL if v >= RX_OK else ORANGE) for k, v, n in rx_lvl],
                 ref=RX_OK, ref_lab=f"all {pc(RX_OK)}", w=330, lw=104),
           f"{N_PRES:,} prescriptions at {N_RX_FAC} facilities." + sup(7))
    + card("Unserved items are concentrated in a few facilities",
           hbars([(f, n / N_UN_ITEMS, f"{f}: {n} of {N_UN_ITEMS} unserved "
                   f"items", ORANGE) for f, n in UN_FAC[:5]], w=330, lw=104,
                 vmax=max(n for _, n in UN_FAC[:5]) / N_UN_ITEMS * 1.25,
                 fmt=lambda v: f"{round(v * N_UN_ITEMS)}")
           + '<p class="txt">Most often unserved: '
           + "; ".join(f"{esc(_short_med(m))} ({n})" for m, n in UN_MED[:4])
           + '.</p>',
           f"Count of prescribed items recorded not available, by facility "
           f"(top 5 of {N_UN_FAC}).")
    + '</div>')

S5 = section(5, "Order fill rate", "REQUISITIONS",
    '<div class="g g3">'
    + points([
        f"{B(pc(FILL))} of units requested were delivered in June–August "
        f"({DV / 1e6:,.1f}M of {RQ / 1e6:,.1f}M units).",
        f"Vertical programmes are filled least: {fill_prog[0][0]} "
        f"{B(pc(fill_prog[0][1]))}, {fill_prog[1][0]} "
        f"{pc(fill_prog[1][1])}.",
        f"By level, {min(fill_lvl, key=lambda x: x[1])[0].lower()}s are "
        f"lowest at {pc(min(v for _, v in fill_lvl))}.",
        f"Monthly fill moved between {pc(min(v for _, v in fill_mon))} and "
        f"{pc(max(v for _, v in fill_mon))}."])
    + card("Fill rate by programme",
           hbars([(k, v, f"{k}: {pc(v)}", TEAL if v >= FILL else ORANGE)
                  for k, v in fill_prog], ref=FILL,
                 ref_lab=f"all {pc(FILL)}", w=330, lw=118, vmax=1.2,
                 fmt=lambda v: pc(v, 0)),
           "Units delivered ÷ units requested; can exceed 100%." + sup(5))
    + card("Fill rate by level of care",
           hbars([(SHORT.get(l, l), v, f"{l}: {pc(v)}",
                   TEAL if v >= FILL else ORANGE) for l, v in fill_lvl],
                 ref=FILL, ref_lab=f"all {pc(FILL)}", w=330, lw=104,
                 vmax=1.2, fmt=lambda v: pc(v, 0)),
           "Orange = below the national fill rate." + sup(5))
    + '</div>')

S6 = section(6, "Community-level medicine availability", "COMMUNITY",
    '<div class="g g3">'
    + points([
        f"Community health workers had {B(pc(CHW))} of the medicines they "
        f"manage ({N_CHW} CHWs).",
        f"Only {B(f'{CHW_COMPLETE}')} of {N_CHW_ASSESSED} CHWs "
        f"({pc(pct(CHW_COMPLETE, N_CHW_ASSESSED))}) held every medicine "
        f"they manage.",
        f"Lowest in {chw_prov[0][0]} ({pc(chw_prov[0][1])}), highest in "
        f"{chw_prov[-1][0]} ({pc(chw_prov[-1][1])}).",
        (f"On the products both carry, CHWs had {pc(CMP_ALL[2])} against "
         f"{pc(CMP_ALL[4])} at facilities." if CMP_ALL else
         "MFT zoning allocates each zone one antimalarial, so availability "
         "is over managed products only.")])
    + card("CHW availability by province",
           hbars([(k, v, f"{k}: {pc(v)} of {n:,} managed CHW × medicine "
                   f"rows", TEAL if v >= CHW else ORANGE)
                  for k, v, n in chw_prov], ref=CHW,
                 ref_lab=f"all {pc(CHW)}", w=330, lw=90),
           "Managed products only (MFT zoning)." + sup(8))
    + card("Complete CHW basket",
           f'<p class="big">{pc(pct(CHW_COMPLETE, N_CHW_ASSESSED))}</p>'
           f'<p class="bigs">{CHW_COMPLETE} of {N_CHW_ASSESSED} CHWs held '
           f'every medicine they manage</p>'
           + (f'<p class="big o">{(CMP_ALL[4] - CMP_ALL[2]) * 100:+.1f} '
              f'pts</p><p class="bigs">facility minus CHW availability on '
              f'the shared products</p>' if CMP_ALL else ""), "" )
    + '</div>')


def dcell(k, v):
    if v is None:
        return '<td class="na">–</td>'
    nat = NATIONAL[k]
    good = v <= nat if k in LOWER_IS_BETTER else v >= nat
    return f'<td class="{"ok" if good else "bad"}">{v * 100:.1f}</td>'


cols = [DIST[:10], DIST[10:20], DIST[20:]]
tables = "".join(
    '<table class="dt"><thead><tr><th>District</th><th>Avail.</th>'
    '<th>Ever had stock-out</th><th>Quantity matching</th><th>Fill</th>'
    '</tr></thead><tbody>'
    + "".join(f'<tr><td>{esc(d["d"])} <span class="pv">{esc(d["prov"])}'
              f'</span></td>' + dcell("av", d["av"]) + dcell("so", d["so"])
              + dcell("acc", d["acc"]) + dcell("fill", d["fill"]) + '</tr>'
              for d in col) + '</tbody></table>' for col in cols)
S7 = section(7, "District by district", "HETEROGENEITY",
    '<div class="card kp wide"><h3>Key points</h3><ul class="two">'
    f'<li>Availability ranges from {B(pc(lo_av["av"]))} in {lo_av["d"]} '
    f'to {B(pc(hi_av["av"]))} in {hi_av["d"]}.</li>'
    f'<li>Quantity matching varies most: {B(pc(lo_acc["acc"]))} in '
    f'{lo_acc["d"]} against {pc(hi_acc["acc"])} in {hi_acc["d"]}.</li>'
    f'<li>Most often ever had a stock-out: {hi_so["d"]} '
    f'({B(pc(hi_so["so"]))}); lowest fill rate: {lo_fill["d"]} '
    f'({pc(lo_fill["fill"])}).</li>'
    f'<li>{B(len(WEAK))} districts are worse than national on at least '
    f'three of the four indicators: {", ".join(WEAK) or "none"}.</li>'
    '</ul></div>'
    f'<div class="card wide"><div class="dts">{tables}</div><p class="foot">'
    f'Ranked by availability (managed products only, %). Ever had stock-out '
    f'= % of managed pairs out at least one day Jun–Aug; Quantity matching = '
    f'% of stock '
    f'records matching the shelf exactly; Fill = % of units requested '
    f'delivered. <b class="gt">Teal</b> = better than the national figure '
    f'(availability {pc(AV_B)}, ever had stock-out {pc(SO_INC)}, quantity '
    f'matching '
    f'{pc(EXACT)}, fill {pc(FILL)}); <b class="ot">orange</b> = worse; – = '
    f'fewer than 30 records.{sup(1, 4, 5)}</p></div>')

# Priorities for action removed on request (2026-10-16); the data gaps
# stay, under section 7.
GAPS = (
    '<div class="gaps"><b>Data gaps:</b> the upstream question has a usable '
    f'answer for both stores on {len(UPB)} of {len(R.UP)} unavailable '
    f'medicines; {N_SYS0:,} stock records show a system quantity of 0, scored 0% '
    'for quantity accuracy; stock-out days are recorded per month, so '
    'episodes cannot be counted; consumption covers June–August only; '
    'prescription availability reflects what prescribers chose to write.'
    + sup(2, 4, 7) + '</div>')

SOURCES = [
    f"Medicine stock assessment (KoboToolbox), {N_FAC} facilities, "
    "June–August 2026 reference period, availability on the visit day. "
    "Strengths combined per molecule; method 2 (managed products only) "
    "unless stated.",
    "Asked of medicines unavailable at the facility: was it at RMS "
    f"Regional / RMS Central? Base {len(UPB)} answered Yes or No for both; "
    "distribution gap matches the form's own classification on every one.",
    "Months of stock = quantity on hand ÷ average monthly consumption "
    "Jun–Aug; teaching and referral hospitals adequate at 1–3 months; emergency order point "
    "≤ 0.5 months.",
    "Records naming a digital recording system. Quantity accuracy = mean of "
    "max(0, 1 − |system − physical| ÷ system), 0% where the system shows 0.",
    "Requisitions June–August: units delivered ÷ units requested, over "
    "medicine-months with a request.",
    "Lost sales: daily demand (Jun–Aug consumption ÷ days in stock) × days "
    "out × unit tariff price × 1.2; priced pairs with in-stock consumption "
    "only.",
    f"Prescription assessment: {N_PRES:,} prescriptions, {len(R.RX_ANS):,} "
    f"items, {N_RX_FAC} facilities; unserved value at unit cost.",
    f"Community health worker assessment: {N_CHW} CHWs; availability over "
    "medicines the CHW manages.",
]
SRC = ('<ol class="src">' + "".join(f"<li>{esc(x)}</li>" for x in SOURCES)
       + '</ol><p class="src2">Figures are read from the indicator dashboard '
       '(Indicator_dashboard.html), where every chart can be filtered by '
       'province, level, programme, facility and medicine. Prepared '
       'October 2026.</p>')

CSS = f"""
:root{{--teal:{TEAL};--orange:{ORANGE};--ink:{INK};--muted:{MUTED};
 --card:{CARD};--head:#0d3a3f}}
*{{box-sizing:border-box}}
body{{margin:0;background:#dfe5e4;color:var(--ink);
 font:12.5px/1.4 -apple-system,"Segoe UI",Helvetica,Arial,sans-serif}}
.page{{max-width:1040px;margin:18px auto;background:#fff;padding:26px 30px;
 box-shadow:0 2px 14px rgba(0,0,0,.12)}}
h1,h2,h3,.big,.kv{{font-family:Georgia,"Times New Roman",serif}}
.mast{{background:var(--head);color:#fff;padding:22px 26px;margin:-26px -30px
 18px}}
.kick{{font-size:10.5px;letter-spacing:.16em;color:#b9d3cf;margin:0 0 8px}}
.mast h1{{font-size:28px;line-height:1.15;margin:0 0 10px}}
.mast p.lede{{margin:0;font-size:14px;color:#e7f0ee}}
.mast p.lede b{{color:#7fd3c4}}
.kpis{{display:grid;grid-template-columns:repeat(5,1fr);border:1px solid
 #cfd8d6;margin-bottom:18px}}
.kpi{{padding:10px 12px;border-right:1px solid #cfd8d6}}
.kpi:last-child{{border-right:0}}
.kv{{font-size:23px;color:var(--teal);margin:0;font-weight:700}}
.kl{{margin:2px 0 4px;font-size:12px}}
.ks{{margin:0;font-size:10.5px;color:var(--muted)}}
section{{margin:0 0 12px}}
.sh{{display:flex;justify-content:space-between;align-items:baseline;
 border-bottom:2px solid var(--teal);margin:0 0 10px;padding-bottom:3px}}
.sh h2{{font-size:17px;margin:0}}
.tag{{font-size:10.5px;letter-spacing:.12em;color:var(--teal)}}
.g{{display:grid;gap:10px;margin-bottom:6px}}
.g2{{grid-template-columns:1fr 1fr}}
.g3{{grid-template-columns:1fr 1fr 1fr}}
.span2{{grid-column:span 2}}
.card{{background:var(--card);padding:9px 11px}}
.card h3{{font-size:13.5px;margin:0 0 8px;font-family:inherit;
 font-weight:700}}
.foot{{font-size:10.5px;color:var(--muted);margin:8px 0 0}}
.txt{{font-size:12px;margin:6px 0 0}}
.big{{font-size:22px;font-weight:700;color:var(--teal);margin:2px 0 0}}
.big.o{{color:var(--orange)}}
.bigs{{font-size:11.5px;color:var(--muted);margin:0 0 6px}}
svg.ch{{width:100%;height:auto;display:block}}
.ch .lab{{font-size:10.5px;fill:var(--ink)}}
.ch .val{{font-size:10.5px;fill:var(--ink);font-weight:700}}
.ch .val.sm{{font-size:9.5px}}
.ch .lg{{font-size:9.5px;fill:var(--muted)}}
.ch .inb{{font-size:9.5px;fill:#fff;font-weight:700}}
.ch .ref{{stroke:#555;stroke-dasharray:3 3}}
.ch .refl{{font-size:9px;fill:var(--muted)}}
.ch .sbig{{font-size:15px;font-weight:700;fill:#fff}}
.ch .ssm{{font-size:10px;fill:#fff}}
sup{{font-size:8px;color:var(--teal)}}
.dts{{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}}
table{{border-collapse:collapse;width:100%}}
.dt th,.dt td{{padding:2px 5px;font-size:11px;border-bottom:1px solid
 #d6dedc;text-align:right}}
.dt th:first-child,.dt td:first-child{{text-align:left}}
.dt th{{font-weight:600;color:var(--muted);font-size:10px}}
.dt td.ok{{color:var(--teal);font-weight:700}}
.dt td.bad{{color:var(--orange)}}
.dt td.na{{color:#999}}
.pv{{font-size:9px;color:var(--muted)}}
.gt{{color:var(--teal)}}.ot{{color:var(--orange)}}
.act th{{text-align:left;font-size:10.5px;letter-spacing:.08em;
 color:var(--muted);border-bottom:2px solid var(--ink);padding:5px 6px}}
.act td{{vertical-align:top;padding:8px 6px;border-bottom:1px solid #d6dedc;
 font-size:12px}}
.act td.no{{font-family:Georgia,serif;font-weight:700;color:var(--teal);
 width:22px}}
.act td.m{{color:var(--muted);width:24%}}
.kp{{background:#fff;border:1px solid #cfd8d6;border-top:3px solid var(--teal)}}
.kp ul{{margin:0;padding-left:16px}}
.kp li{{margin:0 0 4px;font-size:12px}}
.kp ul.two{{columns:2;column-gap:28px}}
.kp ul.two li{{break-inside:avoid}}
.wide{{margin-bottom:8px}}
.cta{{background:var(--head);color:#fff;padding:12px 16px;margin:12px 0 8px;
 font-size:12.5px}}
.cta b{{color:#7fd3c4}}
.gaps{{border-left:4px solid var(--orange);background:#f6efe9;
 padding:8px 12px;font-size:11.5px}}
.src{{columns:3;column-gap:22px;font-size:9.5px;color:var(--muted);
 padding-left:16px;margin:8px 0 2px}}
.src li{{margin:0 0 3px;break-inside:avoid}}
.src2{{font-size:10px;color:var(--muted);margin:0}}
@media (max-width:820px){{.kpis{{grid-template-columns:repeat(2,1fr)}}
 .g2,.g3,.dts{{grid-template-columns:1fr}}.span2{{grid-column:auto}}
 .page{{padding:16px}}.mast{{margin:-16px -16px 14px}}.src{{columns:1}}}}
@media print{{body{{background:#fff}}.page{{box-shadow:none;margin:0;
 max-width:none;padding:10mm 12mm}}.mast{{margin:-10mm -12mm 14px}}
 .p2{{break-before:page}}}}
@page{{size:A4;margin:0}}
"""

HEAD = (
    '<header class="mast"><p class="kick">SUMMARY BRIEF · MEDICINE STOCK AND '
    'PRESCRIPTION ASSESSMENT · RWANDA · OCTOBER 2026</p>'
    '<h1>Medicine Availability in Rwanda\'s Health Facilities: Stock, '
    'Supply and Records</h1>'
    '</header>')

doc = ("<!doctype html><html lang='en'><head><meta charset='utf-8'>"
       "<meta name='viewport' content='width=device-width,initial-scale=1'>"
       "<title>Medicine Availability Brief</title><style>" + CSS
       + "</style></head><body>"
       + '<div class="page">' + HEAD + f'<div class="kpis">{KPIS}</div>'
       + S1 + S2 + S3 + '</div>'
       + '<div class="page p2">' + S4 + S5 + S6 + S7 + '</div>'
       + "</body></html>")

if os.path.exists(OUT):
    os.chmod(OUT, 0o644)
with open(OUT, "w", encoding="utf-8") as fh:
    fh.write(doc)
print(f"wrote {os.path.basename(OUT)} ({os.path.getsize(OUT):,} bytes)")
print(f"  availability {pc(AV_B)} (method 1 {pc(AV_A)}); gap "
      f"{N_GAP}/{N_GAP + N_NAT}; mix {MIX}; exact {pc(EXACT)}; qacc "
      f"{pc(QACC)}; fill {pc(FILL)}; stock-out {pc(SO_INC)}; CHW {pc(CHW)}; "
      f"loss {LOSS:,.0f}; {len(DIST)} districts")
