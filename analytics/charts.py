"""Inline SVG charts for the availability disaggregations.

No chart library and no CDN: the dashboards are opened from a Box folder and
from email attachments, often with no network, so every mark is an SVG element
written here.

WHAT THE BARS SHOW
One number per category, and only one: of the medicines EXPECTED at that
facility (EML-authorised for Essential Medicines, managed for programme
products), the share that were physically present on the day of the visit. The filled part of the bar is that share; the grey remainder
is everything else, whether the facility had run out or does not stock the
medicine at all. The reason split is still in the tables for anyone who needs
it, but it is deliberately absent from these charts.

    bars()      one bar per category, availability only
    dotplot()   one row per province, its districts as dots on a range line
    heat()      a category x level grid, when both dimensions matter at once
"""
import html
import math
from collections import defaultdict

# Dark blue for the value, grey for the rest. Two colours and nothing else:
# the bars carry a single measure, so a third hue would imply a distinction
# the chart is no longer making.
AVAIL = "#1c4f91"
AVAIL_HI = "#2a68bd"
REST = "#3a3a38"
INK = "#e8e7e1"
MUT = "#9a9892"
RULE = "rgba(255,255,255,.14)"


def esc(x):
    return html.escape("" if x is None else str(x), quote=True)


def fmt(v):
    return "-" if v is None else f"{v * 100:.1f}%"


def singular(word):
    """facilities -> facility, districts -> district. Only ever used on the
    two words this module passes, so it does not need to be clever."""
    return word[:-3] + "y" if word.endswith("ies") else word.rstrip("s")


def _legend(x, y, items):
    out = []
    for i, (c, lab) in enumerate(items):
        cx = x + i * 150
        out.append(f'<rect x="{cx}" y="{y - 9}" width="11" height="11" '
                   f'rx="2" fill="{c}"/>'
                   f'<text x="{cx + 16}" y="{y}" class="cl">{esc(lab)}</text>')
    return "".join(out)


# Every piece of text on a chart is white. The grey that used to carry the
# axis ticks, the counts and the heatmap's column headers was legible on a
# monitor and not on a projector or in a printed deck, which is where these
# charts are actually read. Sizes go up a step with it: a 10px grey tick was
# doing two things wrong at once.
CSS = """
.chart{width:100%;height:auto;display:block}
.chart text{fill:#fff}
.chart .cl{fill:#fff;font-size:11.5px}
.chart .cv{fill:#fff;font-size:11.5px;font-weight:700;
 font-variant-numeric:tabular-nums}
.chart .ct{fill:#fff;font-size:12px}
.chart .cn{fill:#fff;font-size:11px;font-variant-numeric:tabular-nums}
.chart .cs{fill:#fff;font-size:10px;font-variant-numeric:tabular-nums;
 opacity:.78}
.chart .gl{stroke:rgba(255,255,255,.14);stroke-width:1}
.chart .gl.z{stroke:rgba(255,255,255,.34)}
.chart rect.seg:hover{opacity:.82}
"""


def bars(rows, label_w=210, bar_h=19, gap=7, pad_t=40, width=940,
         note_right=True, hit="available", base="expected", fill=AVAIL,
         track=REST):
    """Horizontal bars for a rate.

    Rows are (label, n, hits). The filled dark-blue bar is hits / n; the grey
    track behind it is the full 100%, so a short bar is read against the same
    reference everywhere. The caller sorts. `hit` and `base` are the words in
    the tooltip ("12% available - 3 of 25 expected"), so a stock-out rate
    can say what it counts.
    """
    h = pad_t + len(rows) * (bar_h + gap) + 16
    plot_w = width - label_w - (78 if note_right else 8)
    out = [f'<svg class="chart" viewBox="0 0 {width} {h}" role="img">',
           f'<style>{CSS}</style>']
    for i in (0, 25, 50, 75, 100):
        x = label_w + plot_w * i / 100
        out.append(f'<line class="gl{" z" if i == 0 else ""}" x1="{x:.1f}" '
                   f'y1="{pad_t - 8}" x2="{x:.1f}" y2="{h - 14}"/>')
        out.append(f'<text x="{x:.1f}" y="{h - 3}" class="cn" '
                   f'text-anchor="middle">{i}%</text>')
    for i, (lab, n, av) in enumerate(rows):
        y = pad_t + i * (bar_h + gap)
        v = av / n if n else 0
        # A fill rate can exceed 100% - more was delivered than requested -
        # and the bar must stop at the end of its track rather than running
        # past it. The printed figure keeps the true value.
        w = min(v, 1.0)
        out.append(f'<text x="{label_w - 10}" y="{y + bar_h - 5}" class="ct" '
                   f'text-anchor="end">{esc(lab)}</text>')
        out.append(f'<rect x="{label_w}" y="{y}" width="{plot_w:.1f}" '
                   f'height="{bar_h}" fill="{track}"/>')
        out.append(f'<rect class="seg" x="{label_w}" y="{y}" '
                   f'width="{plot_w * w:.1f}" height="{bar_h}" '
                   f'fill="{fill}"><title>{esc(lab)}: {fmt(v)} {esc(hit)} '
                   f'- {av:,} of {n:,} {esc(base)}</title></rect>')
        if w < .86:
            out.append(f'<text x="{label_w + plot_w * w + 6:.1f}" '
                       f'y="{y + bar_h - 5}" class="cv">{fmt(v)}</text>')
        else:
            out.append(f'<text x="{label_w + plot_w * w - 6:.1f}" '
                       f'y="{y + bar_h - 5}" class="cv" text-anchor="end">'
                       f'{fmt(v)}</text>')
        if note_right:
            out.append(f'<text x="{width - 6}" y="{y + bar_h - 5}" '
                       f'class="cn" text-anchor="end">n={n:,}</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# radial bars (user template, 2026-10-26)
# ---------------------------------------------------------------
# One ring per row: a short stub at the bottom carries the label, the ring
# turns round the left and runs out along the top as a bar whose length is
# the rate. Every top bar starts at the same x, so their lengths compare
# directly. Rows in the order given, the first innermost - the caller sorts
# weakest first, so the best is the outer ring. A row's colour is fixed by
# its rank: the user's light-to-dark blue ramp (2026-10-26), lightest on the
# inner (weakest) ring, darkest on the outer (best).
# The six swatches as sent, plus one darker step (#33496b) for a seventh ring.
RADIAL_RAMP = ["#c4cddf", "#a5b3d1", "#7a93c1", "#567aaf", "#4c6c9d",
               "#405c86", "#33496b"]


def radial_colour(i, k):
    """Ring i of k (0 = innermost) -> a ramp step, spread over the ramp so the
    outer ring is always the darkest."""
    n = len(RADIAL_RAMP)
    # half-up, as Math.round in the browser redraw
    j = int(i * (n - 1) / (k - 1) + 0.5) if 1 < k <= n else min(i, n - 1)
    return RADIAL_RAMP[j]


def radial_bars(rows, width=940, stroke=14, step=24, r0=34, hit="available",
                base="expected"):
    """rows are (label, n, hits); the top bar's length is hits / n."""
    if not rows:
        return '<p class="fempty">Nothing to show.</p>'
    k = len(rows)
    rmax = r0 + (k - 1) * step
    cx = 34 + rmax
    cy = 30 + rmax
    h = cy + rmax + 30
    span = width - cx - 90
    out = [f'<svg class="chart" viewBox="0 0 {width} {h}" role="img">']
    for i, (lab, n, av) in enumerate(rows):
        v = av / n if n else 0
        r = r0 + i * step
        col = radial_colour(i, k)
        x1 = cx + span * min(v, 1.0)
        out.append(f'<path d="M{cx + 12:.1f},{cy + r:.1f} H{cx:.1f} '
                   f'A{r:.1f},{r:.1f} 0 0 1 {cx:.1f},{cy - r:.1f} '
                   f'H{x1:.1f}" fill="none" stroke="{col}" '
                   f'stroke-width="{stroke}" stroke-linecap="round">'
                   f'<title>{esc(lab)}: {fmt(v)} {esc(hit)} - {av:,} of '
                   f'{n:,} {esc(base)}</title></path>')
        out.append(f'<text x="{x1 + 14:.1f}" y="{cy - r + 4.5:.1f}" '
                   f'class="cv">{fmt(v)}</text>')
        out.append(f'<text x="{cx + 30:.1f}" y="{cy + r + 4.5:.1f}" '
                   f'class="ct" style="font-weight:700">'
                   f'{esc(lab)} <tspan class="cn" style="font-weight:400">'
                   f'n={n:,}</tspan></text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# track columns (user template, 2026-10-26)
# ---------------------------------------------------------------
# One column per row on a grey track to 100%, the top-right corner rounded,
# filled with a left-to-right gradient of its ramp colour; the value inside
# the top of the fill, the label under a baseline with a diamond marker.
# Rows arrive weakest first (as radial_bars takes them) and are drawn best
# first, left to right; colour by rank on RADIAL_RAMP, darkest = best.
def _tint(hexcol, f=0.38):
    c = [int(hexcol[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{int(v + (255 - v) * f + 0.5):02x}" for v in c)


RADIAL_TINT = [_tint(c) for c in RADIAL_RAMP]
# the two lightest steps take dark text
RADIAL_DARK_TEXT = 2


def _wrap(lab, n=14):
    lines, cur = [], ""
    for w in str(lab).split():
        if cur and len(cur) + len(w) + 1 > n:
            lines.append(cur)
            cur = w
        else:
            cur = (cur + " " + w).strip()
    lines.append(cur)
    return lines[:2]


def track_columns(rows, width=940, height=400, col_w=64, pad=40,
                  hit="available", base="expected", gid="tc"):
    """rows are (label, n, hits), weakest first; drawn best first."""
    if not rows:
        return '<p class="fempty">Nothing to show.</p>'
    k = len(rows)
    top, yb = 30, 290
    ph = yb - top
    step = (width - 2 * pad) / k
    rad = col_w * 0.45
    out = [f'<svg class="chart" viewBox="0 0 {width} {height}" role="img">',
           '<defs>']
    for i in range(k):
        j = int(i * (len(RADIAL_RAMP) - 1) / (k - 1) + 0.5) if 1 < k <= len(
            RADIAL_RAMP) else min(i, len(RADIAL_RAMP) - 1)
        out.append(f'<linearGradient id="{gid}{i}" x1="0" x2="1" y1="0" '
                   f'y2="0"><stop offset="0" stop-color="{RADIAL_RAMP[j]}"/>'
                   f'<stop offset="1" stop-color="{RADIAL_TINT[j]}"/>'
                   f'</linearGradient>')
    out.append('</defs>')
    out.append(f'<line x1="{pad - 20}" y1="{yb + 22}" x2="{width - pad + 20}" '
               f'y2="{yb + 22}" stroke="#55554f" stroke-width="1.5"/>')

    def shape(x, y, w, hgt):
        r = min(rad, hgt)
        return (f'M{x:.1f},{yb} V{y:.1f} H{x + w - r:.1f} A{r:.1f},{r:.1f} '
                f'0 0 1 {x + w:.1f},{y + r:.1f} V{yb} Z')
    for pos, (lab, n, av) in enumerate(reversed(rows)):
        i = k - 1 - pos                      # rank, 0 = weakest
        j = int(i * (len(RADIAL_RAMP) - 1) / (k - 1) + 0.5) if 1 < k <= len(
            RADIAL_RAMP) else min(i, len(RADIAL_RAMP) - 1)
        v = av / n if n else 0
        cx = pad + step * (pos + 0.5)
        x = cx - col_w / 2
        hgt = ph * min(v, 1.0)
        y = yb - hgt
        out.append(f'<path d="{shape(x, top, col_w, ph)}" fill="#5a5a55"/>')
        out.append(f'<path class="seg" d="{shape(x, y, col_w, hgt)}" '
                   f'fill="url(#{gid}{i})"><title>{esc(lab)}: {fmt(v)} '
                   f'{esc(hit)} - {av:,} of {n:,} {esc(base)}</title></path>')
        tcol = "#14243d" if j < RADIAL_DARK_TEXT else "#fff"
        if hgt >= 40:
            out.append(f'<text x="{cx:.1f}" y="{y + 26:.1f}" '
                       f'text-anchor="middle" style="font-size:15px;'
                       f'font-weight:700;fill:{tcol}">{fmt(v)}</text>')
        else:
            out.append(f'<text x="{cx:.1f}" y="{y - 8:.1f}" class="cv" '
                       f'text-anchor="middle">{fmt(v)}</text>')
        out.append(f'<rect x="{cx - 5:.1f}" y="{yb + 17}" width="10" '
                   f'height="10" transform="rotate(45 {cx:.1f} {yb + 22})" '
                   f'fill="{RADIAL_RAMP[j]}"/>')
        ls = _wrap(lab)
        for li, line in enumerate(ls):
            out.append(f'<text x="{cx:.1f}" y="{yb + 50 + li * 15}" '
                       f'class="ct" text-anchor="middle" '
                       f'style="font-weight:700">{esc(line)}</text>')
        out.append(f'<text x="{cx:.1f}" y="{yb + 50 + len(ls) * 15}" '
                   f'class="cn" text-anchor="middle">n={n:,}</text>')
    out.append("</svg>")
    return "".join(out)


def day_bars(rows, max_v=92, unit="days", label_w=210, bar_h=19, gap=7,
             pad_t=40, width=940, fill=AVAIL, track=REST):
    """Horizontal bars for an AVERAGE, not a rate - the twin of bars().

    Rows are (label, n, value). The track runs 0 to `max_v` (92 = the days in
    June-August, the most a pair can be out), so a bar's length is read
    against the whole window. n is the number of pairs the mean is over.
    """
    h = pad_t + len(rows) * (bar_h + gap) + 16
    plot_w = width - label_w - 78
    out = [f'<svg class="chart" viewBox="0 0 {width} {h}" role="img">',
           f'<style>{CSS}</style>']
    ticks = [0, 15, 30, 45, 60, 75, 90] if max_v == 92 else \
        [int(max_v * k / 4 + 0.5) for k in range(5)]   # as JS Math.round
    for i in ticks:
        x = label_w + plot_w * i / max_v
        out.append(f'<line class="gl{" z" if i == 0 else ""}" x1="{x:.1f}" '
                   f'y1="{pad_t - 8}" x2="{x:.1f}" y2="{h - 14}"/>')
        out.append(f'<text x="{x:.1f}" y="{h - 3}" class="cn" '
                   f'text-anchor="middle">{i}</text>')
    for i, (lab, n, v) in enumerate(rows):
        y = pad_t + i * (bar_h + gap)
        w = min(v / max_v, 1.0) if max_v else 0
        out.append(f'<text x="{label_w - 10}" y="{y + bar_h - 5}" class="ct" '
                   f'text-anchor="end">{esc(lab)}</text>')
        out.append(f'<rect x="{label_w}" y="{y}" width="{plot_w:.1f}" '
                   f'height="{bar_h}" fill="{track}"/>')
        out.append(f'<rect class="seg" x="{label_w}" y="{y}" '
                   f'width="{plot_w * w:.1f}" height="{bar_h}" '
                   f'fill="{fill}"><title>{esc(lab)}: {v:.1f} {esc(unit)} on '
                   f'average, over {n:,} pairs</title></rect>')
        txt = f"{v:.1f} {unit}"
        if w < .80:
            out.append(f'<text x="{label_w + plot_w * w + 6:.1f}" '
                       f'y="{y + bar_h - 5}" class="cv">{esc(txt)}</text>')
        else:
            out.append(f'<text x="{label_w + plot_w * w - 6:.1f}" '
                       f'y="{y + bar_h - 5}" class="cv" text-anchor="end">'
                       f'{esc(txt)}</text>')
        out.append(f'<text x="{width - 6}" y="{y + bar_h - 5}" '
                   f'class="cn" text-anchor="end">n={n:,}</text>')
    out.append("</svg>")
    return "".join(out)


def dotplot(rows, label_w=180, row_h=46, pad_t=52, width=940, lo=None,
            hi=None, unit="districts"):
    """One row per group: the units inside it as dots on a range line.

    A single rate for a group says where it sits and nothing about whether it
    is one problem or several. A level of care whose facilities run 29% to
    100% does not have a level-wide problem, it has a tail of weak facilities
    wearing an average. The line spans min to max, the hollow dots are the
    units, the filled diamond is the group's own rate.

    Dot size and opacity scale down as a row gets crowded. A level of care
    pools two hundred facilities where a province pools eight districts, and
    4.5px circles at full strength turn that row into one solid bar.
    """
    vals = [v for _, _, pooled, ds in rows
            for v in list(ds.values()) + [pooled]]
    dense = max((len(ds) for _, _, _, ds in rows), default=0)
    dot_r = 4.5 if dense <= 24 else (3.4 if dense <= 80 else 2.8)
    dot_op = 1.0 if dense <= 24 else (0.72 if dense <= 80 else 0.5)
    lo = lo if lo is not None else min(vals) - .03
    hi = hi if hi is not None else max(vals) + .03
    h = pad_t + len(rows) * row_h + 24
    plot_w = width - label_w - 70

    def X(v):
        return label_w + plot_w * (v - lo) / (hi - lo)

    out = [f'<svg class="chart" viewBox="0 0 {width} {h}" role="img">',
           f'<style>{CSS}</style>']
    out.append(_legend(label_w, 18, [(AVAIL, "group availability"),
                                     (REST, f"one of the {unit}")]))
    # A 5-point step is right across a narrow range and unreadable across a
    # wide one, so the step follows the range rather than being fixed.
    step = .05 if (hi - lo) <= .30 else .10
    t = (int(lo / step) + 1) * step
    while t < hi:
        out.append(f'<line class="gl" x1="{X(t):.1f}" y1="{pad_t - 8}" '
                   f'x2="{X(t):.1f}" y2="{h - 20}"/>')
        out.append(f'<text x="{X(t):.1f}" y="{h - 6}" class="cn" '
                   f'text-anchor="middle">{t * 100:.0f}%</text>')
        t += step
    for i, (prov, n, pooled, ds) in enumerate(rows):
        y = pad_t + i * row_h + row_h / 2 - 8
        out.append(f'<text x="{label_w - 10}" y="{y + 4}" class="ct" '
                   f'text-anchor="end">{esc(prov)}</text>')
        if ds:
            out.append(f'<line x1="{X(min(ds.values())):.1f}" y1="{y}" '
                       f'x2="{X(max(ds.values())):.1f}" y2="{y}" '
                       f'stroke="{RULE}" '
                       f'stroke-width="3" stroke-linecap="round"/>')
        order = sorted(ds.items(), key=lambda x: x[1])
        for d, v in order:
            out.append(f'<circle cx="{X(v):.1f}" cy="{y}" r="{dot_r}" '
                       f'fill="none" stroke="{MUT}" stroke-width="1.4" '
                       f'opacity="{dot_op}">'
                       f'<title>{esc(d)} - {fmt(v)}</title></circle>')
        # Name the ends of the range. Labelling every dot would overlap into
        # noise, and the lowest and highest district are the two a reader
        # actually wants off the chart.
        if order:
            lo_d, lo_v = order[0]
            hi_d, hi_v = order[-1]
            out.append(f'<text x="{X(lo_v):.1f}" y="{y - 11}" class="cs" '
                       f'text-anchor="middle">{esc(lo_d)} '
                       f'{lo_v * 100:.1f}%</text>')
            if hi_d != lo_d:
                out.append(f'<text x="{X(hi_v):.1f}" y="{y - 11}" class="cs" '
                           f'text-anchor="middle">{esc(hi_d)} '
                           f'{hi_v * 100:.1f}%</text>')
        x = X(pooled)
        out.append(f'<path d="M{x:.1f},{y - 7} L{x + 6.5:.1f},{y} '
                   f'L{x:.1f},{y + 7} L{x - 6.5:.1f},{y} Z" fill="{AVAIL}">'
                   f'<title>{esc(prov)} pooled - {fmt(pooled)}</title></path>')
        out.append(f'<text x="{width - 6}" y="{y + 4}" class="cv" '
                   f'text-anchor="end">{fmt(pooled)}</text>')
        if ds:
            out.append(f'<text x="{label_w - 10}" y="{y + 17}" class="cn" '
                       f'text-anchor="end">{len(ds)} '
                       f'{singular(unit) if len(ds) == 1 else unit} &#183; '
                       f'{n:,} obs</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# gauge cards
# ---------------------------------------------------------------
# One card per level of care, each a semicircular tick gauge. The arc is drawn
# as discrete ticks rather than a solid band because a tick arc reads as a
# proportion of a countable whole, which is what an availability rate is, and
# because the boundary between filled and unfilled is exact instead of being
# softened by a stroke cap.
#
# Colours follow the page: the filled ticks are the same dark blue as the bars,
# the unfilled ones grey. The reference layout this was modelled on uses teal
# on navy; the shape, the card grid and the typography are taken from it, the
# palette is not, so this figure matches the rest of the dashboard.
GAUGE_CSS = """
.ggrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(248px,1fr));
 gap:14px;margin:4px 0 0}
.gcard{background:#141c2e;border:1px solid rgba(255,255,255,.09);
 border-radius:14px;padding:14px 16px 18px}
.gcard svg{display:block;width:100%;height:auto;margin:0 auto}
.gcard .gname{font-size:14.5px;font-weight:700;color:#fff;margin:6px 0 0}
.gcard .gcount{font-size:13px;color:#fff;margin:5px 0 0}
.gcard .gcount b{font-size:17px;font-weight:700}
.gcard .gcount span{color:#aeb6c6}
.gval{fill:#fff;font-size:30px;font-weight:700;letter-spacing:-.01em}
.gcap{fill:#aeb6c6;font-size:10.5px}
"""

G_ON = "#4f9be8"
G_OFF = "#4a4f5c"


def gauge(value, ticks=42, w=250, h=148,
          caption="available of those expected", on=G_ON, off=G_OFF):
    """Semicircular tick gauge. `value` is 0-1 and may exceed 1."""
    cx, cy, ro, ri = w / 2, h - 18, w / 2 - 21, w / 2 - 47
    frac = max(0.0, min(1.0, value))
    lit = int(round(frac * ticks))
    out = [f'<svg viewBox="0 0 {w} {h}" role="img" '
           f'aria-label="{value * 100:.0f} percent">']
    for i in range(ticks):
        a = math.pi * (1 + (i + 0.5) / ticks)
        c, sn = math.cos(a), math.sin(a)
        out.append(f'<line x1="{cx + ri * c:.1f}" y1="{cy + ri * sn:.1f}" '
                   f'x2="{cx + ro * c:.1f}" y2="{cy + ro * sn:.1f}" '
                   f'stroke="{on if i < lit else off}" stroke-width="3.4" '
                   f'stroke-linecap="round"/>')
    out.append(f'<text x="{cx:.0f}" y="{cy - 34:.0f}" class="gval" '
               f'text-anchor="middle">{value * 100:.1f}%</text>')
    out.append(f'<text x="{cx:.0f}" y="{cy - 15:.0f}" class="gcap" '
               f'text-anchor="middle">{esc(caption)}</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# half rose (overall figure on the stock-out page)
# ---------------------------------------------------------------
# A white disc with the headline in its centre and one wedge per group fanned
# over its right half, from twelve to six o'clock; a wedge's LENGTH is its
# rate against the largest one. Modelled on the user's reference infographic
# (2026-10-19).
ROSE_COLS = ["#f2b134", "#e2603a", "#8e3f8f", "#2b3a8c", "#1f8a8a",
             "#6b7a1f"]


def half_rose(rows, centre, sub, w=940, h=500, r0=118, rmax=205,
              caption="with a stock-out", colours=None, colour_fn=None,
              dark_text=()):
    """rows are (label, n, hits); centre/sub are the disc's two lines.
    colour_fn(rate) -> fill colours a wedge by its value (else colours by
    position); fills listed in dark_text get a dark percentage label."""
    rows = [r for r in rows if r[1]]
    if not rows:
        return '<p class="fempty">Nothing to show.</p>'
    cols = colours or ROSE_COLS
    cx, cy, k = 330, h / 2, len(rows)
    vmax = max(hit / n for _, n, hit in rows) or 1
    ri = r0 - 14
    out = [f'<svg class="chart" viewBox="0 0 {w} {h}" role="img">']
    for i, (lab, n, hit) in enumerate(rows):
        v = hit / n
        a0 = math.radians(-90 + i * 180 / k)
        a1 = math.radians(-90 + (i + 1) * 180 / k)
        am = (a0 + a1) / 2
        R = r0 + 22 + (rmax - r0 - 22) * v / vmax
        p = (f"M{cx + ri * math.cos(a0):.1f},{cy + ri * math.sin(a0):.1f} "
             f"L{cx + R * math.cos(a0):.1f},{cy + R * math.sin(a0):.1f} "
             f"A{R:.1f},{R:.1f} 0 0 1 {cx + R * math.cos(a1):.1f},"
             f"{cy + R * math.sin(a1):.1f} "
             f"L{cx + ri * math.cos(a1):.1f},{cy + ri * math.sin(a1):.1f} "
             f"A{ri:.1f},{ri:.1f} 0 0 0 {cx + ri * math.cos(a0):.1f},"
             f"{cy + ri * math.sin(a0):.1f} Z")
        fill = colour_fn(v) if colour_fn else cols[i % len(cols)]
        out.append(f'<path d="{p}" fill="{fill}" '
                   f'stroke="#000" stroke-width="2"><title>{esc(lab)}: '
                   f'{fmt(v)} {esc(caption)} - {hit:,} of {n:,}</title>'
                   f'</path>')
        rm = (r0 + R) / 2 + 6
        out.append(f'<text x="{cx + rm * math.cos(am):.1f}" '
                   f'y="{cy + rm * math.sin(am) + 5:.1f}" text-anchor="middle"'
                   f' style="font-size:15px;font-weight:700;fill:'
                   f'{"#12303a" if fill in dark_text else "#fff"}">'
                   f'{v * 100:.1f}%</text>')
        lx, ly = cx + (R + 14) * math.cos(am), cy + (R + 14) * math.sin(am)
        out.append(f'<text x="{lx:.1f}" y="{ly - 2:.1f}" class="ct" '
                   f'style="font-weight:700">{esc(lab)}</text>'
                   f'<text x="{lx:.1f}" y="{ly + 12:.1f}" class="cn">'
                   f'{hit:,} of {n:,}</text>')
    out.append(f'<circle cx="{cx}" cy="{cy:.1f}" r="{r0 + 6}" fill="#e9e9e6" '
               f'opacity=".35"/><circle cx="{cx}" cy="{cy:.1f}" r="{r0}" '
               f'fill="#f4f4f2"/><circle cx="{cx}" cy="{cy:.1f}" '
               f'r="{r0 - 26}" fill="#fff" stroke="#dcdcd8"/>')
    out.append(f'<text x="{cx}" y="{cy - 2:.1f}" text-anchor="middle" '
               f'style="font-size:34px;font-weight:700;fill:#222">'
               f'{esc(centre)}</text>'
               f'<text x="{cx}" y="{cy + 20:.1f}" text-anchor="middle" '
               f'style="font-size:11.5px;fill:#555">{esc(sub)}</text>'
               f'<rect x="{cx - 34}" y="{cy + 30:.1f}" width="68" height="3" '
               f'rx="1.5" fill="url(#rosebar)"/>')
    out.append('<defs><linearGradient id="rosebar"><stop offset="0" '
               f'stop-color="{cols[0]}"/><stop offset=".5" stop-color='
               f'"{cols[2]}"/><stop offset="1" stop-color="{cols[4]}"/>'
               '</linearGradient></defs></svg>')
    return "".join(out)


def gauges(rows, caption="available of those expected", on=G_ON, off=G_OFF):
    """rows are (label, n, hits) - one gauge card per row, grouped in a
    grid; `caption` is the line under the percentage."""
    cards = []
    for lab, n, av in rows:
        v = av / n if n else 0
        cards.append(
            f'<div class="gcard">{gauge(v, caption=caption, on=on, off=off)}'
            f'<p class="gname">{esc(lab)}</p>'
            f'<p class="gcount"><b>{av:,}</b> <span>of {n:,}</span></p>'
            f'</div>')
    return f'<div class="ggrid">{"".join(cards)}</div>'


# ---------------------------------------------------------------
# dial gauges
# ---------------------------------------------------------------
# A 240-degree dial per province: 0 at the lower left, 50 at the top, 100 at
# the lower right, with a tick ring outside a solid band and a needle on the
# value. The extra furniture over the half-gauge used for level of care is
# the point - a labelled 0/25/50/75/100 scale lets a reader place a province
# on an absolute range rather than only against the other cards, which is what
# the needle is for.
#
# The band and the lit ticks are the page's blue, the rest grey. The reference
# design is orange on white; the geometry is taken from it, the palette is the
# dashboard's.
DIAL_CSS = """
.dgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(272px,1fr));
 gap:14px;margin:4px 0 0}
.dcard{background:#141c2e;border:1px solid rgba(255,255,255,.09);
 border-radius:14px;padding:16px 16px 18px}
.dcard svg{display:block;width:100%;height:auto;margin:0 auto}
.dcard .dname{font-size:15px;font-weight:700;color:#fff;margin:4px 0 0;
 text-align:center}
.dcard .dcount{font-size:12.5px;color:#aeb6c6;margin:4px 0 0;text-align:center}
.dcard .dcount b{color:#fff;font-size:14px;font-weight:700}
.dscale{fill:#fff;font-size:11px;font-variant-numeric:tabular-nums}
.dbig{fill:#fff;font-size:34px;font-weight:700;letter-spacing:-.015em}
.dcap{fill:#aeb6c6;font-size:11px}
.dgrid.solo{grid-template-columns:minmax(272px,520px);justify-content:center}
"""

A0, SWEEP = 150.0, 240.0      # degrees: 0 lower-left, 50 top, 100 lower-right


def _pt(cx, cy, r, t):
    a = math.radians(A0 + SWEEP * t)
    return cx + r * math.cos(a), cy + r * math.sin(a)


def _arc(cx, cy, r, t0, t1):
    x0, y0 = _pt(cx, cy, r, t0)
    x1, y1 = _pt(cx, cy, r, t1)
    big = 1 if (t1 - t0) * SWEEP > 180 else 0
    return f"M{x0:.1f},{y0:.1f} A{r:.1f},{r:.1f} 0 {big} 1 {x1:.1f},{y1:.1f}"


def dial(value, w=272, h=226, ticks=41,
         caption="of medicines expected"):
    """240-degree dial. `value` is 0-1; the needle clamps at the ends."""
    cx, cy = w / 2, h / 2 - 6
    rb = w / 2 - 58          # band radius
    rt0, rt1 = rb + 16, rb + 28   # tick ring
    v = max(0.0, min(1.0, value))
    out = [f'<svg viewBox="0 0 {w} {h}" role="img" '
           f'aria-label="{value * 100:.0f} percent of medicines available">']
    # tick ring
    for i in range(ticks):
        t = i / (ticks - 1)
        x0, y0 = _pt(cx, cy, rt0, t)
        x1, y1 = _pt(cx, cy, rt1, t)
        on = t <= v + 1e-9
        out.append(f'<line x1="{x0:.1f}" y1="{y0:.1f}" x2="{x1:.1f}" '
                   f'y2="{y1:.1f}" stroke="{G_ON if on else G_OFF}" '
                   f'stroke-width="{3.2 if i % 5 else 4.4}" '
                   f'stroke-linecap="round"/>')
    # band: full track, then the filled part over it
    out.append(f'<path d="{_arc(cx, cy, rb, 0, 1)}" fill="none" '
               f'stroke="{G_OFF}" stroke-width="13" stroke-linecap="round" '
               f'opacity=".55"/>')
    if v > 0:
        out.append(f'<path d="{_arc(cx, cy, rb, 0, v)}" fill="none" '
                   f'stroke="{G_ON}" stroke-width="13" '
                   f'stroke-linecap="round"/>')
    # scale labels
    for t, lab in ((0, "0"), (.25, "25"), (.5, "50"), (.75, "75"), (1, "100")):
        lx, ly = _pt(cx, cy, rt1 + 15, t)
        out.append(f'<text x="{lx:.1f}" y="{ly + 4:.1f}" class="dscale" '
                   f'text-anchor="middle">{lab}</text>')
    # needle
    nx, ny = _pt(cx, cy, rb - 9, v)
    out.append(f'<line x1="{cx:.1f}" y1="{cy:.1f}" x2="{nx:.1f}" '
               f'y2="{ny:.1f}" stroke="{G_ON}" stroke-width="3.2" '
               f'stroke-linecap="round"/>')
    out.append(f'<circle cx="{cx:.1f}" cy="{cy:.1f}" r="7.5" fill="#141c2e" '
               f'stroke="{G_ON}" stroke-width="3.4"/>')
    # readout, below the hub where the dial leaves room
    out.append(f'<text x="{cx:.0f}" y="{cy + 52:.0f}" class="dbig" '
               f'text-anchor="middle">{value * 100:.1f}%</text>')
    out.append(f'<text x="{cx:.0f}" y="{cy + 70:.0f}" class="dcap" '
               f'text-anchor="middle">{esc(caption)}</text>')
    out.append("</svg>")
    return "".join(out)


def dials(rows, caption="of medicines expected", noun="expected"):
    """rows are (label, n, available) - one dial card per row.

    `caption` is the line under the big number inside the dial and `noun` the
    word in the count beneath it, so the same card serves "76% of medicines
    expected" and "81% were available at RMS Regional" without a second
    implementation drifting away from this one.
    """
    cards = []
    for lab, n, av in rows:
        v = av / n if n else 0
        cards.append(
            f'<div class="dcard">{dial(v, caption=caption)}'
            f'<p class="dname">{esc(lab)}</p>'
            f'<p class="dcount"><b>{av:,}</b> of {n:,} {esc(noun)}'
            f'{_sub(lab)}</p></div>')
    return f'<div class="dgrid">{"".join(cards)}</div>'


_SUBTITLES = {}


def set_districts(d):
    """{label: district count} - the province cards' own subtitle."""
    _SUBTITLES.clear()
    _SUBTITLES.update({k: f"{v} districts" for k, v in d.items()})


def set_subtitles(d):
    """{label: free text} for any other dial grid."""
    _SUBTITLES.clear()
    _SUBTITLES.update(d)


def _sub(label):
    t = _SUBTITLES.get(label)
    return f" &#183; {t}" if t else ""


# ---------------------------------------------------------------
# sankey
# ---------------------------------------------------------------
# Level of care on the left, drug class on the right. TWO variables are
# carried at once, and keeping them apart is the whole design:
#
#   SIZE   - node height and ribbon width are a COUNT of medicines that were
#            expected and on the shelf. Size answers "how much".
#   COLOUR - one hue per LEVEL OF CARE, carried by the left node and by every
#            ribbon leaving it, so a band can be traced back to the level it
#            came from where the ribbons cross.
#
# The availability RATE is still on every node and in every tooltip, as a
# number rather than a shade. Colour is identity here, not magnitude: it says
# where a flow starts, and nothing about how well that level is stocked.
SANKEY_CSS = """
.sk{width:100%;height:auto;display:block}
.sk text{fill:#fff}
.sk .sknode{font-size:9.5px}
.sk .sknum{font-size:8.5px;opacity:.75;font-variant-numeric:tabular-nums}
.sk .skpc{font-size:9.5px;font-weight:700;font-variant-numeric:tabular-nums}
.sk path.rb:hover{opacity:.9!important}
.sk rect.nd{stroke:rgba(255,255,255,.35);stroke-width:.8}
"""

# One hue per level of care, in reporting order. Six distinguishable colours
# that stay apart at 60% opacity where ribbons overlap.
SK_SRC = ["#4f9be8", "#2f6fd0", "#3fb0a0", "#8f7fe0", "#d98f4a", "#c25f8a"]
SK_DST = "#5a6275"


def sankey(avail, total, left_order, width=940, node_w=15, gap=7, pad=120,
           top_right=12, height=580, min_label=0.012):
    """`avail` and `total` are {(left, right): count}; ribbons use `avail`.

    `top_right` caps the right-hand column - the full class list would give
    one-pixel nodes with unreadable labels, so the smallest are pooled into a
    single 'Other classes' node rather than being dropped, which keeps the
    column total equal to the real total.
    """
    tot = sum(avail.values())
    if not tot:
        return '<p class="fempty">Nothing matches these filters.</p>'
    rsum = defaultdict(int)
    for (l, r), v in avail.items():
        rsum[r] += v
    keep = [k for k, _ in sorted(rsum.items(), key=lambda x: -x[1])[:top_right]]
    other = "Other classes"

    def fold(src):
        o = defaultdict(int)
        for (l, r), v in src.items():
            o[(l, r if r in keep else other)] += v
        return o

    fa, ft = fold(avail), fold(total)
    rsum, lsum, rtot, ltot = (defaultdict(int), defaultdict(int),
                              defaultdict(int), defaultdict(int))
    for (l, r), v in fa.items():
        rsum[r] += v
        lsum[l] += v
    for (l, r), v in ft.items():
        rtot[r] += v
        ltot[l] += v
    rights = [k for k in keep if rsum.get(k)]
    if rsum.get(other):
        rights.append(other)
    lefts = [k for k in left_order if lsum.get(k)]

    def place(keys, totals):
        avail_h = height - pad - gap * max(0, len(keys) - 1)
        pos, y = {}, pad / 2
        for k in keys:
            hgt = avail_h * totals[k] / tot
            pos[k] = (y, hgt)
            y += hgt + gap
        return pos

    lp, rp = place(lefts, lsum), place(rights, rsum)
    x0, x1 = 2, width - node_w - 2
    out = [f'<svg class="sk" viewBox="0 0 {width} {height}" role="img" '
           f'aria-label="Available medicines by level of care and drug '
           f'class">']
    lcur = {k: lp[k][0] for k in lefts}
    rcur = {k: rp[k][0] for k in rights}
    ordered = sorted(fa.items(), key=lambda kv: (lefts.index(kv[0][0]),
                                                 rights.index(kv[0][1])))
    for (l, r), v in ordered:
        hl = lp[l][1] * v / lsum[l]
        hr = rp[r][1] * v / rsum[r]
        y0, y1 = lcur[l], rcur[r]
        lcur[l] += hl
        rcur[r] += hr
        xa, xb = x0 + node_w, x1
        xm = (xa + xb) / 2
        n = ft.get((l, r), v)
        rate = v / n if n else None
        out.append(
            f'<path class="rb" d="M{xa},{y0:.1f} C{xm:.0f},{y0:.1f} '
            f'{xm:.0f},{y1:.1f} {xb},{y1:.1f} L{xb},{y1 + hr:.1f} '
            f'C{xm:.0f},{y1 + hr:.1f} {xm:.0f},{y0 + hl:.1f} '
            f'{xa},{y0 + hl:.1f} Z" fill="{SK_SRC[lefts.index(l) % len(SK_SRC)]}" '
            f'opacity=".46">'
            f'<title>{esc(l)} &#8594; {esc(r)}\n{fmt(rate)} available '
            f'&#183; {v:,} of {n:,} expected</title></path>')
    for i, k in enumerate(lefts):
        y, hgt = lp[k]
        rate = lsum[k] / ltot[k] if ltot[k] else None
        out.append(f'<rect class="nd" x="{x0}" y="{y:.1f}" width="{node_w}" '
                   f'height="{hgt:.1f}" '
                   f'fill="{SK_SRC[i % len(SK_SRC)]}"/>')
        ty = y + hgt / 2 - (4 if hgt < 24 else 0)
        out.append(f'<text class="sknode" x="{x0 + node_w + 7}" '
                   f'y="{ty:.1f}">{esc(k)} '
                   f'<tspan class="skpc">{fmt(rate)}</tspan></text>')
        out.append(f'<text class="sknum" x="{x0 + node_w + 7}" '
                   f'y="{ty + 10:.1f}">{lsum[k]:,} of {ltot[k]:,} '
                   f'expected</text>')
    for k in rights:
        y, hgt = rp[k]
        rate = rsum[k] / rtot[k] if rtot[k] else None
        out.append(f'<rect class="nd" x="{x1}" y="{y:.1f}" width="{node_w}" '
                   f'height="{hgt:.1f}" fill="{SK_DST}"/>')
        if rsum[k] / tot >= min_label:
            out.append(f'<text class="sknode" x="{x1 - 7}" '
                       f'y="{y + hgt / 2:.1f}" text-anchor="end">'
                       f'{esc(k)} <tspan class="skpc">{fmt(rate)}</tspan>'
                       f'</text>')
            out.append(f'<text class="sknum" x="{x1 - 7}" '
                       f'y="{y + hgt / 2 + 10:.1f}" text-anchor="end">'
                       f'{rsum[k]:,} of {rtot[k]:,} expected</text>')
    # key: the colour names the level a ribbon comes from
    kx = x0 + node_w + 7
    for i, k in enumerate(lefts):
        out.append(f'<rect x="{kx}" y="{height - 26}" width="12" height="12" '
                   f'rx="2" fill="{SK_SRC[i % len(SK_SRC)]}"/>'
                   f'<text class="sknum" x="{kx + 17}" y="{height - 16}">'
                   f'{esc(k)}</text>')
        kx += 24 + 5.5 * len(k)
    out.append(f'<text class="sknum" x="{x0 + node_w + 7}" '
               f'y="{height - 34}">Colour = level of care &#183; '
               f'size = number of medicines available &#183; '
               f'the rate is on each label</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# grouped bars
# ---------------------------------------------------------------
# One group per category, one bar per series inside it. Used where two
# measures of the same category must be read against each other - a single
# bar list with "- Regional" and "- Central" suffixes made the reader do that
# pairing by eye, and across twenty rows that is where errors come from.
GROUP_COLS = ["#1c4f91", "#5693e0"]


def grouped(rows, series, label_w=210, bar_h=15, inner=3, gap=12, pad_t=42,
            width=940):
    """rows is [(label, [(n, hits), ...])] in the order of `series`."""
    if not rows:
        return '<p class="fempty">Nothing to show.</p>'
    k = len(series)
    grp = k * bar_h + (k - 1) * inner
    h = pad_t + len(rows) * (grp + gap) + 18
    plot_w = width - label_w - 78
    out = [f'<svg class="chart" viewBox="0 0 {width} {h}" role="img">']
    for i, nm in enumerate(series):
        x = label_w + i * 150
        out.append(f'<rect x="{x}" y="{pad_t - 32}" width="11" height="11" '
                   f'rx="2" fill="{GROUP_COLS[i % len(GROUP_COLS)]}"/>'
                   f'<text x="{x + 16}" y="{pad_t - 22}" class="cl">'
                   f'{esc(nm)}</text>')
    for t in (0, 25, 50, 75, 100):
        x = label_w + plot_w * t / 100
        out.append(f'<line class="gl{" z" if t == 0 else ""}" x1="{x:.1f}" '
                   f'y1="{pad_t - 8}" x2="{x:.1f}" y2="{h - 16}"/>'
                   f'<text x="{x:.1f}" y="{h - 4}" class="cn" '
                   f'text-anchor="middle">{t}%</text>')
    for i, (lab, vals) in enumerate(rows):
        y0 = pad_t + i * (grp + gap)
        out.append(f'<text x="{label_w - 10}" y="{y0 + grp / 2 + 4:.0f}" '
                   f'class="ct" text-anchor="end">{esc(lab)}</text>')
        for j, (n, hit) in enumerate(vals):
            y = y0 + j * (bar_h + inner)
            v = hit / n if n else 0
            w_ = plot_w * min(v, 1.0)
            out.append(f'<rect x="{label_w}" y="{y}" width="{plot_w:.1f}" '
                       f'height="{bar_h}" fill="{REST}"/>')
            out.append(f'<rect class="seg" x="{label_w}" y="{y}" '
                       f'width="{w_:.1f}" height="{bar_h}" '
                       f'fill="{GROUP_COLS[j % len(GROUP_COLS)]}">'
                       f'<title>{esc(lab)} - {esc(series[j])}: {fmt(v)} '
                       f'({hit:,} of {n:,})</title></rect>')
            out.append(f'<text x="{label_w + w_ + 6:.1f}" '
                       f'y="{y + bar_h - 3}" class="cv" '
                       f'style="font-size:10px">{fmt(v)}</text>')
        out.append(f'<text x="{width - 6}" y="{y0 + grp / 2 + 4:.0f}" '
                   f'class="cn" text-anchor="end">'
                   f'n={vals[0][0]:,} / {vals[-1][0]:,}</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# grouped columns
# ---------------------------------------------------------------
# The same data as grouped(), drawn vertically. Columns suit a short ordered
# sequence - the six levels of care run lowest to highest, and a vertical
# arrangement reads as a progression in a way horizontal bars do not.
AVG_COL = "#eda100"          # columns(avg=True): the pooled column
AVG_LABEL = "Average Jun-Aug"


def columns(rows, series, width=940, height=330, pad_l=56, pad_b=74,
            pad_t=46, bar_w=26, inner=4, colours=None, avg=False):
    """rows is [(label, [(n, hits), ...])] in the order of `series`. avg
    adds a last column per cluster, the pooled sum(hits) / sum(n), drawn in
    AVG_COL; n under the cluster stays the first / last real series."""
    if not rows:
        return '<p class="fempty">Nothing to show.</p>'
    cols = colours or GROUP_COLS
    if avg:
        nlab = {lab: (vals[0][0], vals[-1][0]) for lab, vals in rows}
        rows = [(lab, list(vals) + [(sum(n for n, _ in vals),
                                     sum(h for _, h in vals))])
                for lab, vals in rows]
        series = list(series) + [AVG_LABEL]
        cols = list(cols[:len(series) - 1]) + [AVG_COL]
    k = len(series)
    plot_h = height - pad_t - pad_b
    plot_w = width - pad_l - 16
    step = plot_w / len(rows)
    out = [f'<svg class="chart" viewBox="0 0 {width} {height}" role="img">']
    for i, nm in enumerate(series):
        x = pad_l + i * 150
        out.append(f'<rect x="{x}" y="{pad_t - 30}" width="11" height="11" '
                   f'rx="2" fill="{cols[i % len(cols)]}"/>'
                   f'<text x="{x + 16}" y="{pad_t - 20}" class="cl">'
                   f'{esc(nm)}</text>')
    for t in (0, 25, 50, 75, 100):
        y = pad_t + plot_h * (1 - t / 100)
        out.append(f'<line class="gl{" z" if t == 0 else ""}" x1="{pad_l}" '
                   f'y1="{y:.1f}" x2="{width - 16}" y2="{y:.1f}"/>'
                   f'<text x="{pad_l - 8}" y="{y + 4:.1f}" class="cn" '
                   f'text-anchor="end">{t}%</text>')
    for i, (lab, vals) in enumerate(rows):
        cx = pad_l + step * (i + 0.5)
        grp_w = k * bar_w + (k - 1) * inner
        x0 = cx - grp_w / 2
        for j, (n, hit) in enumerate(vals):
            v = min(hit / n if n else 0, 1.0)
            bh = plot_h * v
            x = x0 + j * (bar_w + inner)
            y = pad_t + plot_h - bh
            out.append(f'<rect x="{x:.1f}" y="{pad_t:.1f}" width="{bar_w}" '
                       f'height="{plot_h:.1f}" fill="{REST}" opacity=".5"/>')
            out.append(f'<rect class="seg" x="{x:.1f}" y="{y:.1f}" '
                       f'width="{bar_w}" height="{bh:.1f}" '
                       f'fill="{cols[j % len(cols)]}">'
                       f'<title>{esc(lab)} - {esc(series[j])}: {fmt(v)} '
                       f'({hit:,} of {n:,})</title></rect>')
            out.append(f'<text x="{x + bar_w / 2:.1f}" y="{y - 4:.1f}" '
                       f'class="cv" style="font-size:10px" '
                       f'text-anchor="middle">{v * 100:.1f}%</text>')
        # label, wrapped onto two lines so the long tier names fit
        words = str(lab).split()
        lines, cur = [], ""
        for w in words:
            if len(cur) + len(w) + 1 <= 16:
                cur = (cur + " " + w).strip()
            else:
                lines.append(cur)
                cur = w
        lines.append(cur)
        for li, line in enumerate(lines[:3]):
            out.append(f'<text x="{cx:.1f}" y="{pad_t + plot_h + 16 + li * 11:.0f}" '
                       f'class="cn" text-anchor="middle">{esc(line)}</text>')
        out.append(f'<text x="{cx:.1f}" '
                   f'y="{pad_t + plot_h + 16 + len(lines[:3]) * 11:.0f}" '
                   f'class="cs" text-anchor="middle">'
                   f'n={(nlab[lab] if avg else (vals[0][0], vals[-1][0]))[0]:,}'
                   f' / {(nlab[lab] if avg else (vals[0][0], vals[-1][0]))[1]:,}'
                   f'</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# line with a shaded band between two series
# ---------------------------------------------------------------
# Two lines across an ORDERED axis, with the area between them filled. The
# band is the point: where Central runs above Regional, the stock reached the
# national store and stopped before the region, and the thickness of the
# shading is the size of that loss at each tier.
#
# A line is only honest here because level of care is a real sequence, lowest
# to highest. It still implies something untrue - that there are values
# BETWEEN the tiers - so the markers are drawn and sized by how many
# observations sit behind each point, and the counts are printed under the
# axis. Without that, a tier resting on four observations draws the same
# confident line as one resting on five hundred.
BAND_A = "#5693e0"          # upper series
BAND_B = "#1c4f91"          # lower series
BAND_FILL = "#2a68bd"
DOT_R = 11.0                # dot radius when area_band(connect=False)


def brace(x, y1, y2, w):
    """SVG path of a right-facing curly bracket from (x, y1) down to
    (x, y2), its tip pointing right by w at the midpoint."""
    ym, q = (y1 + y2) / 2, min(6.0, abs(y2 - y1) / 4)
    m = x + w / 2
    return (f"M{x:.1f},{y1:.1f} Q{m:.1f},{y1:.1f} {m:.1f},{y1 + q:.1f} "
            f"L{m:.1f},{ym - q:.1f} Q{m:.1f},{ym:.1f} {x + w:.1f},{ym:.1f} "
            f"Q{m:.1f},{ym:.1f} {m:.1f},{ym + q:.1f} L{m:.1f},{y2 - q:.1f} "
            f"Q{m:.1f},{y2:.1f} {x:.1f},{y2:.1f}")


def area_band(rows, series, width=940, height=360, pad_l=56, pad_b=78,
              pad_t=50, lo=None, hi=None, connect=True):
    """rows is [(label, [(n, hits), ...])] - exactly two series.

    connect=False draws the dots alone - no lines, no shaded band - for
    categories that are not a sequence; when both series share one
    denominator the count under each category is printed once."""
    if not rows or len(series) != 2:
        return '<p class="fempty">Nothing to show.</p>'
    vals = [h / n if n else 0 for _, vs in rows for n, h in vs]
    lo = lo if lo is not None else max(0.0, min(vals) - .08)
    hi = hi if hi is not None else min(1.0, max(vals) + .06)
    plot_h = height - pad_t - pad_b
    plot_w = width - pad_l - 20
    # end categories sit 70 px inside the frame so their count labels
    # clear the axis on the left and the edge on the right (2026-11-05)
    inset = 70
    step = (plot_w - 2 * inset) / max(1, len(rows) - 1)
    ns = [max(n for n, _ in vs) for _, vs in rows]
    nmax = max(ns) or 1

    def X(i):
        return pad_l + inset + step * i

    def Y(v):
        return pad_t + plot_h * (1 - (v - lo) / (hi - lo))

    def R(n):
        # dots alone are all one size (2026-10-16): the count is printed
        # under each category instead
        return DOT_R if not connect else 3.2 + 6.0 * math.sqrt(n / nmax)

    # labels clear the dots: offsets grow with the fixed dot size
    lift = DOT_R + 7 if not connect else 13
    drop = DOT_R + 12 if not connect else 18
    side = DOT_R + 5 if not connect else 7

    out = [f'<svg class="chart" viewBox="0 0 {width} {height}" role="img">']
    t = math.ceil(lo * 20) / 20
    while t <= hi + 1e-9:
        y = Y(t)
        out.append(f'<line class="gl" x1="{pad_l}" y1="{y:.1f}" '
                   f'x2="{width - 20}" y2="{y:.1f}"/>'
                   f'<text x="{pad_l - 8}" y="{y + 4:.1f}" class="cn" '
                   f'text-anchor="end">{t * 100:.0f}%</text>')
        t += .05

    pts = [[(h / n if n else 0) for n, h in (vs[k],)][0]
           for _, vs in rows for k in (0,)]
    a_vals = [vs[0][1] / vs[0][0] if vs[0][0] else 0 for _, vs in rows]
    b_vals = [vs[1][1] / vs[1][0] if vs[1][0] else 0 for _, vs in rows]

    band = ("M" + " L".join(f"{X(i):.1f},{Y(v):.1f}"
                            for i, v in enumerate(b_vals))
            + " L" + " L".join(f"{X(i):.1f},{Y(v):.1f}"
                               for i, v in reversed(list(enumerate(a_vals))))
            + " Z")
    if connect:
        out.append(f'<path d="{band}" fill="{BAND_FILL}" opacity=".22"/>')
        for vals_, col, nm in ((b_vals, BAND_B, series[1]),
                               (a_vals, BAND_A, series[0])):
            d = "M" + " L".join(f"{X(i):.1f},{Y(v):.1f}"
                                for i, v in enumerate(vals_))
            out.append(f'<path d="{d}" fill="none" stroke="{col}" '
                       f'stroke-width="2.6" stroke-linejoin="round"/>')

    for i, (lab, vs) in enumerate(rows):
        for vals_, col, j in ((b_vals, BAND_B, 1), (a_vals, BAND_A, 0)):
            n, hit = vs[j]
            out.append(f'<circle cx="{X(i):.1f}" cy="{Y(vals_[i]):.1f}" '
                       f'r="{R(n):.1f}" fill="{col}" stroke="#0e0e0d" '
                       f'stroke-width="1.4">'
                       f'<title>{esc(lab)} - {esc(series[j])}: '
                       f'{fmt(vals_[i])} ({hit:,} of {n:,})</title></circle>')
        gap = a_vals[i] - b_vals[i]
        top, bot = max(a_vals[i], b_vals[i]), min(a_vals[i], b_vals[i])
        # each value carries its count (hit/n), the gap its n (2026-11-05)
        tj = 0 if a_vals[i] >= b_vals[i] else 1
        th, tn_ = vs[tj][1], vs[tj][0]
        bh, bn_ = vs[1 - tj][1], vs[1 - tj][0]
        out.append(f'<text x="{X(i):.1f}" y="{Y(top) - lift:.1f}" class="cv" '
                   f'style="font-size:10.5px" text-anchor="middle">'
                   f'{top * 100:.1f}% ({th:,}/{tn_:,})</text>')
        out.append(f'<text x="{X(i):.1f}" y="{Y(bot) + drop:.1f}" class="cv" '
                   f'style="font-size:10.5px" text-anchor="middle">'
                   f'{bot * 100:.1f}% ({bh:,}/{bn_:,})</text>')
        gn = (f"n={vs[0][0]:,}" if vs[0][0] == vs[1][0]
              else f"n={vs[0][0]:,} / {vs[1][0]:,}")
        if abs(gap) >= .02:
            tx = X(i) + side
            if not connect:
                # a "}" from one dot centre to the other, the gap at its tip
                bx, bw = X(i) + DOT_R + 4, 9
                out.append(f'<path d="{brace(bx, Y(top), Y(bot), bw)}" '
                           f'fill="none" stroke="#c3c2b7" stroke-width="1.5" '
                           f'stroke-linecap="round"/>')
                tx = bx + bw + 5
            # the gap written out (2026-11-05): first series minus second,
            # as percentages and as items
            (na, ha), (nb, hb) = vs[0], vs[1]
            ym = (Y(top) + Y(bot)) / 2
            nd = f"n={abs(ha - hb):,}" if na == nb else gn
            out.append(f'<text x="{tx:.1f}" y="{ym + 4:.1f}" class="cs" '
                       f'text-anchor="start">{gap * 100:+.1f} pts ({nd})</text>')
        words = str(lab).split()
        lines, cur = [], ""
        for w in words:
            if len(cur) + len(w) + 1 <= 15:
                cur = (cur + " " + w).strip()
            else:
                lines.append(cur)
                cur = w
        lines.append(cur)
        for li, line in enumerate(lines[:3]):
            out.append(f'<text x="{X(i):.1f}" '
                       f'y="{pad_t + plot_h + 18 + li * 11:.0f}" class="cn" '
                       f'text-anchor="middle">{esc(line)}</text>')
        nlab = (f"n={vs[0][0]:,}" if vs[0][0] == vs[1][0]
                else f"n={vs[0][0]:,} / {vs[1][0]:,}")
        out.append(f'<text x="{X(i):.1f}" '
                   f'y="{pad_t + plot_h + 18 + len(lines[:3]) * 11:.0f}" '
                   f'class="cs" text-anchor="middle">{nlab}</text>')

    for i, nm in enumerate(series):
        x = pad_l + i * 170
        col = BAND_A if i == 0 else BAND_B
        if connect:
            out.append(f'<line x1="{x}" y1="{pad_t - 26}" x2="{x + 22}" '
                       f'y2="{pad_t - 26}" stroke="{col}" '
                       f'stroke-width="2.6"/>')
        out.append(f'<circle cx="{x + 11}" cy="{pad_t - 26}" r="'
                   f'{4 if connect else 6}" fill="{col}"/>'
                   f'<text x="{x + 29}" y="{pad_t - 22}" class="cl">'
                   f'{esc(nm)}</text>')
    if connect:
        out.append(f'<text x="{pad_l + 360}" y="{pad_t - 22}" class="cs">'
                   'shaded band = the gap &#183; marker size = observations'
                   '</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# 100% stacked bars
# ---------------------------------------------------------------
# For a composition that always sums to the whole: each bar is the full
# width and the segments are the shares within it. Used for stock adequacy,
# where under / adequate / over are mutually exclusive and exhaust the base,
# so the question is how the mix differs between groups rather than how big
# each group is. The observation count sits at the right, because a 100% bar
# deliberately hides it.
def mean_bars(rows, label_w=210, bar_h=19, gap=7, pad_t=40, width=940,
              what="average quantity accuracy", unit="records"):
    """bars() for an AVERAGE of 0-1 scores rather than a count rate.

    Rows are (label, n, mean). The bar is the mean on the same 0-100%
    track; the tooltip says what it averages over."""
    h = pad_t + len(rows) * (bar_h + gap) + 16
    plot_w = width - label_w - 78
    out = [f'<svg class="chart" viewBox="0 0 {width} {h}" role="img">',
           f'<style>{CSS}</style>']
    for i in (0, 25, 50, 75, 100):
        x = label_w + plot_w * i / 100
        out.append(f'<line class="gl{" z" if i == 0 else ""}" x1="{x:.1f}" '
                   f'y1="{pad_t - 8}" x2="{x:.1f}" y2="{h - 14}"/>')
        out.append(f'<text x="{x:.1f}" y="{h - 3}" class="cn" '
                   f'text-anchor="middle">{i}%</text>')
    for i, (lab, n, v) in enumerate(rows):
        y = pad_t + i * (bar_h + gap)
        w = max(0.0, min(v, 1.0))
        out.append(f'<text x="{label_w - 10}" y="{y + bar_h - 5}" class="ct" '
                   f'text-anchor="end">{esc(lab)}</text>')
        out.append(f'<rect x="{label_w}" y="{y}" width="{plot_w:.1f}" '
                   f'height="{bar_h}" fill="{REST}"/>')
        out.append(f'<rect class="seg" x="{label_w}" y="{y}" '
                   f'width="{plot_w * w:.1f}" height="{bar_h}" '
                   f'fill="{AVAIL}"><title>{esc(lab)}: {fmt(v)} {esc(what)} '
                   f'- over {n:,} {esc(unit)}</title></rect>')
        if w < .86:
            out.append(f'<text x="{label_w + plot_w * w + 6:.1f}" '
                       f'y="{y + bar_h - 5}" class="cv">{fmt(v)}</text>')
        else:
            out.append(f'<text x="{label_w + plot_w * w - 6:.1f}" '
                       f'y="{y + bar_h - 5}" class="cv" text-anchor="end">'
                       f'{fmt(v)}</text>')
        out.append(f'<text x="{width - 6}" y="{y + bar_h - 5}" '
                   f'class="cn" text-anchor="end">n={n:,}</text>')
    out.append("</svg>")
    return "".join(out)


STACK_COLS = ["#c9544a", "#199e70", "#eda100"]


def stacked(rows, series, colours=None, label_w=210, bar_h=22, gap=8,
            pad_t=46, width=940, min_label=0.07):
    """rows is [(label, [v1, v2, ...])] in the order of `series`."""
    if not rows:
        return '<p class="fempty">Nothing to show.</p>'
    cols = colours or STACK_COLS
    h = pad_t + len(rows) * (bar_h + gap) + 18
    plot_w = width - label_w - 78
    out = [f'<svg class="chart" viewBox="0 0 {width} {h}" role="img">']
    for i, nm in enumerate(series):
        x = label_w + i * 160
        out.append(f'<rect x="{x}" y="{pad_t - 30}" width="11" height="11" '
                   f'rx="2" fill="{cols[i % len(cols)]}"/>'
                   f'<text x="{x + 16}" y="{pad_t - 20}" class="cl">'
                   f'{esc(nm)}</text>')
    for t in (0, 25, 50, 75, 100):
        x = label_w + plot_w * t / 100
        out.append(f'<line class="gl{" z" if t == 0 else ""}" x1="{x:.1f}" '
                   f'y1="{pad_t - 8}" x2="{x:.1f}" y2="{h - 16}"/>'
                   f'<text x="{x:.1f}" y="{h - 4}" class="cn" '
                   f'text-anchor="middle">{t}%</text>')
    for i, (lab, vals) in enumerate(rows):
        y = pad_t + i * (bar_h + gap)
        tot = sum(vals) or 1
        out.append(f'<text x="{label_w - 10}" y="{y + bar_h - 6}" class="ct" '
                   f'text-anchor="end">{esc(lab)}</text>')
        x = label_w
        for j, v in enumerate(vals):
            w_ = plot_w * v / tot
            if w_ <= 0:
                continue
            out.append(f'<rect class="seg" x="{x:.1f}" y="{y}" '
                       f'width="{w_:.1f}" height="{bar_h}" '
                       f'fill="{cols[j % len(cols)]}">'
                       f'<title>{esc(lab)} - {esc(series[j])}: {v:,} of '
                       f'{tot:,} ({v / tot:.1%})</title></rect>')
            if v / tot >= min_label:
                out.append(f'<text x="{x + w_ / 2:.1f}" '
                           f'y="{y + bar_h - 7:.1f}" text-anchor="middle" '
                           f'style="font-size:10.5px;font-weight:700;'
                           f'fill:#fff">{v / tot * 100:.1f}%</text>')
            x += w_
        out.append(f'<text x="{width - 6}" y="{y + bar_h - 6}" class="cn" '
                   f'text-anchor="end">n={tot:,}</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# pie
# ---------------------------------------------------------------
# A single composition that sums to the whole, with few enough slices to
# read. Three categories is the case a pie handles well; past about five it
# stops being readable and a bar is better, which is why this is used for the
# overall adequacy mix and not for the group breakdowns beneath it.
#
# ARRANGEMENT: slices are ordered LARGEST FIRST, starting at twelve o'clock
# and running clockwise. That is the convention a reader expects, and it puts
# the biggest share against the vertical where the eye judges angle best.
# The caller may pass slices in any order; they are sorted here.
#
# LABELS SIT OUTSIDE on leader lines - name, share and count together - so
# nothing depends on matching a colour back to a key, and a small slice is
# labelled as legibly as a large one.
#
# COLOURS are ordinal rather than traffic-light: amber for too little,
# through teal for the right amount, to indigo for too much. Under and over
# are different KINDS of wrong, not degrees of the same one, so they are
# given hues that do not read as "bad" and "worse".
PIE_COLS = ["#d98030", "#1f9e7a", "#4a5fb5", "#8c5fb5", "#5f8cb5"]


def pie(slices, colours=None, w=760, h=380, r=104, title="",
        sort=True, unit="observations"):
    """slices is [(label, value)]; drawn largest first unless sort=False."""
    tot = sum(v for _, v in slices)
    if not tot:
        return '<p class="fempty">Nothing to show.</p>'
    cols = colours or PIE_COLS
    # keep each label's colour tied to its ORIGINAL position, so re-ordering
    # the slices does not silently recolour the categories
    idx = {lab: i for i, (lab, _) in enumerate(slices)}
    data = sorted(slices, key=lambda kv: -kv[1]) if sort else list(slices)
    cx, cy = w / 2, h / 2 + 8
    out = [f'<svg class="chart" viewBox="0 0 {w} {h}" role="img" '
           f'aria-label="{esc(title or "composition")}">']

    def pt(rr, a):
        rad = math.radians(a - 90)
        return cx + rr * math.cos(rad), cy + rr * math.sin(rad)

    a = 0.0
    labels = []
    for lab, v in data:
        span = 360 * v / tot
        x0, y0 = pt(r, a)
        x1, y1 = pt(r, a + span)
        big = 1 if span > 180 else 0
        col = cols[idx[lab] % len(cols)]
        out.append(
            f'<path d="M{cx:.1f},{cy:.1f} L{x0:.1f},{y0:.1f} '
            f'A{r},{r} 0 {big} 1 {x1:.1f},{y1:.1f} Z" fill="{col}" '
            f'stroke="#0e0e0d" stroke-width="2"><title>{esc(lab)}: {v:,} of '
            f'{tot:,} ({v / tot:.1%})</title></path>')
        labels.append((lab, v, a + span / 2, col))
        a += span

    for lab, v, mid, col in labels:
        ex, ey = pt(r + 4, mid)
        bx, by = pt(r + 30, mid)
        right = math.cos(math.radians(mid - 90)) >= 0
        tx = bx + (46 if right else -46)
        out.append(f'<path d="M{ex:.1f},{ey:.1f} L{bx:.1f},{by:.1f} '
                   f'L{tx:.1f},{by:.1f}" fill="none" stroke="{col}" '
                   f'stroke-width="1.6"/>')
        anchor = "start" if right else "end"
        lx = tx + (6 if right else -6)
        out.append(f'<text x="{lx:.0f}" y="{by - 3:.0f}" class="ct" '
                   f'text-anchor="{anchor}">{esc(lab)}</text>')
        out.append(f'<text x="{lx:.0f}" y="{by + 13:.0f}" class="cs" '
                   f'text-anchor="{anchor}">{v / tot:.1%} &#183; '
                   f'{v:,} {esc(unit)}</text>')

    out.append(f'<text x="{cx:.0f}" y="{cy + r + 46:.0f}" class="cn" '
               f'text-anchor="middle">{tot:,} {esc(unit)} in total</text>')
    out.append("</svg>")
    return "".join(out)


RING_COLS = ["#e8a33d", "#17a67f", "#5465c4", "#b5559e", "#2f9bc4"]


# ---------------------------------------------------------------
# radial bars
# ---------------------------------------------------------------
# One concentric ring per category, each starting at twelve o'clock and
# running anticlockwise for its share of a full turn, with the share printed
# where the ring starts and a keyed explanation of each category to the
# right. Unlike ring(), every category has its own track, so a share is read
# against a full circle (the faint track behind it) rather than against its
# neighbours - which is what makes 24% and 26% distinguishable at a glance.
# Rings run outside-in in the order given, so the order of the series is the
# order of the key.
RADIAL_HL = "#ffd166"       # highlighted line under a key entry


def radial(slices, colours=None, centre_title="", centre_sub="", w=940,
           h=470, ring_w=34, ring_gap=9, r_core=84, highlights=None,
           pie=False):
    """slices is [(label, value, blurb)]; colours follow the slice order.
    `highlights` optionally gives, per slice, one extra line printed under
    its explanation on an amber band - a sub-count the reader must not miss.
    pie=True draws the user's pie infographic (2026-10-26) in place of the
    rings: wedges of slightly different reach round a light centre disc,
    each wedge's share printed inside it; the key on the right is the same."""
    tot = sum(sl[1] for sl in slices)
    if not tot:
        return '<p class="fempty">Nothing to show.</p>'
    cols = colours or RING_COLS
    n = len(slices)
    cx, cy = 250, h / 2
    r_outer = r_core + 18 + n * (ring_w + ring_gap) - ring_gap / 2
    out = [f'<svg class="chart" viewBox="0 0 {w} {h}" role="img" '
           f'aria-label="{esc(centre_title or "composition")}">']
    if pie:
        out.append(pie_info(slices, cols, tot, cx, cy, centre_title,
                            centre_sub))
    for i, (lab, v, blurb) in enumerate([] if pie else slices):
        col = cols[i % len(cols)]
        r = r_outer - ring_w / 2 - i * (ring_w + ring_gap)
        share = v / tot
        out.append(f'<circle cx="{cx}" cy="{cy:.1f}" r="{r:.1f}" fill="none" '
                   f'stroke="rgba(255,255,255,.07)" stroke-width="{ring_w}"/>')
        sweep = min(share, 0.9999) * 360
        x0, y0 = cx, cy - r
        rad = math.radians(-90 - sweep)
        x1, y1 = cx + r * math.cos(rad), cy + r * math.sin(rad)
        big = 1 if sweep > 180 else 0
        out.append(f'<path d="M{x0:.1f},{y0:.1f} A{r:.1f},{r:.1f} 0 {big} 0 '
                   f'{x1:.1f},{y1:.1f}" fill="none" stroke="{col}" '
                   f'stroke-width="{ring_w}"><title>{esc(lab)}: {v:,} of '
                   f'{tot:,} ({share:.1%})</title></path>')
        out.append(f'<text x="{cx + 10}" y="{cy - r + 9:.1f}" '
                   f'style="font-size:25px;font-weight:700;fill:{col}">'
                   f'{share * 100:.1f}%</text>')

    if not pie:
        out.append(f'<circle cx="{cx}" cy="{cy:.1f}" r="{r_core}" '
                   f'fill="#141c2e" stroke="rgba(255,255,255,.18)" '
                   f'stroke-width="2"/>')
    if centre_title and not pie:
        out.append(f'<text x="{cx}" y="{cy - 4:.1f}" text-anchor="middle" '
                   f'style="font-size:17px;font-weight:700;fill:#fff;'
                   f'letter-spacing:.03em">{esc(centre_title)}</text>')
    for i, line in enumerate(str(centre_sub).split("|")
                             if centre_sub and not pie else []):
        out.append(f'<text x="{cx}" y="{cy + 16 + i * 14:.1f}" '
                   f'text-anchor="middle" style="font-size:11px;'
                   f'fill:#aeb6c6">{esc(line.strip())}</text>')

    # the key: badge with the count, coloured heading, explanation
    kx, row_h = 560, min(128, (h - 40) / n)
    y0 = cy - row_h * n / 2
    for i, (lab, v, blurb) in enumerate(slices):
        col = cols[i % len(cols)]
        by = y0 + row_h * i + row_h / 2
        out.append(f'<circle cx="{kx}" cy="{by:.1f}" r="35" fill="none" '
                   f'stroke="{col}" stroke-width="4"/>')
        out.append(f'<text x="{kx}" y="{by + 5:.1f}" text-anchor="middle" '
                   f'style="font-size:{15 if v < 10000 else 13}px;'
                   f'font-weight:700;fill:#fff">{v:,}</text>')
        out.append(f'<text x="{kx + 54}" y="{by - 14:.1f}" '
                   f'style="font-size:19px;font-weight:700;fill:{col};'
                   f'letter-spacing:.04em">{esc(lab.upper())}</text>')
        words, lines, cur = str(blurb).split(), [], ""
        for word in words:
            if len(cur) + len(word) + 1 <= 46:
                cur = (cur + " " + word).strip()
            else:
                lines.append(cur)
                cur = word
        if cur:
            lines.append(cur)
        for j, line in enumerate(lines[:3]):
            out.append(f'<text x="{kx + 54}" y="{by + 6 + j * 16:.1f}" '
                       f'style="font-size:12.5px;fill:#aeb6c6">'
                       f'{esc(line)}</text>')
        hl = (highlights or [None] * n)[i]
        if hl:
            # " - " splits the highlight onto two lines inside one box
            hls = [x for x in hl.split(" - ") if x]
            hy = by + 6 + min(3, len(lines)) * 16 + 5
            out.append(f'<rect x="{kx + 50}" y="{hy - 13:.1f}" '
                       f'width="{min(w - kx - 56, 10 + max(map(len, hls)) * 7.4):.0f}" '
                       f'height="{6 + 16 * len(hls)}" rx="4" '
                       f'fill="{RADIAL_HL}" fill-opacity=".16" '
                       f'stroke="{RADIAL_HL}" stroke-opacity=".7"/>')
            for j, line in enumerate(hls):
                out.append(f'<text x="{kx + 55}" y="{hy + j * 16:.1f}" '
                           f'style="font-size:12.5px;font-weight:700;'
                           f'fill:{RADIAL_HL}">{esc(line)}</text>')
    out.append("</svg>")
    return "".join(out)


# Pie infographic for radial(pie=True). The reach of each wedge cycles
# through PIE_REACH for the template's look; it carries no data - the share
# is the ANGLE, and it is printed inside the wedge.
PIE_REACH = [205, 182, 196, 174, 190, 178]
PIE_HUB = 74


def pie_info(slices, cols, tot, cx, cy, title, sub):
    out, a = [], -90.0
    for i, (lab, v, _) in enumerate(slices):
        share = v / tot
        sweep = min(share, 0.9999) * 360
        r = PIE_REACH[i % len(PIE_REACH)]
        a0, a1 = math.radians(a), math.radians(a + sweep)
        x0, y0 = cx + r * math.cos(a0), cy + r * math.sin(a0)
        x1, y1 = cx + r * math.cos(a1), cy + r * math.sin(a1)
        big = 1 if sweep > 180 else 0
        col = cols[i % len(cols)]
        out.append(f'<path d="M{cx:.1f},{cy:.1f} L{x0:.1f},{y0:.1f} '
                   f'A{r},{r} 0 {big} 1 {x1:.1f},{y1:.1f} Z" fill="{col}" '
                   f'stroke="#000" stroke-width="3" stroke-linejoin="round">'
                   f'<title>{esc(lab)}: {v:,} of {tot:,} ({share:.1%})'
                   f'</title></path>')
        am = math.radians(a + sweep / 2)
        rl = (PIE_HUB + r) / 2 + 8
        out.append(f'<text x="{cx + rl * math.cos(am):.1f}" '
                   f'y="{cy + rl * math.sin(am) + 7:.1f}" text-anchor="middle" '
                   f'style="font-size:21px;font-weight:700;fill:#fff;'
                   f'paint-order:stroke;stroke:rgba(0,0,0,.25);'
                   f'stroke-width:2px">{share * 100:.1f}%</text>')
        a += sweep
    out.append(f'<circle cx="{cx}" cy="{cy:.1f}" r="{PIE_HUB + 9}" '
               f'fill="rgba(0,0,0,.45)"/>')
    out.append(f'<circle cx="{cx}" cy="{cy:.1f}" r="{PIE_HUB}" fill="#e9ecf1" '
               f'stroke="#c9ced8" stroke-width="2"/>')
    if title:
        out.append(f'<text x="{cx}" y="{cy - 6:.1f}" text-anchor="middle" '
                   f'style="font-size:14px;font-weight:700;fill:#1e2b45;'
                   f'letter-spacing:.03em">{esc(title)}</text>')
    for i, line in enumerate(str(sub).split("|") if sub else []):
        out.append(f'<text x="{cx}" y="{cy + 13 + i * 13:.1f}" '
                   f'text-anchor="middle" style="font-size:10.5px;'
                   f'fill:#4a5568">{esc(line.strip())}</text>')
    return "".join(out)


# ---------------------------------------------------------------
# 3D ring (extruded donut)
# ---------------------------------------------------------------
# Drawn on request for the overall prescription outcome. A tilted pie
# distorts - slices at the front show a wall and look bigger than slices at
# the back - so the shares are printed in the key and on hover, and the key,
# not the area, is the figure to quote. Slices under 5% are pulled out from
# the ring so a 0.2% sliver is still seen at all.
#
# Geometry: an ellipse seen from above at tilt `k`. Painter's order - the
# inside of the hole (only its far half shows), then the outer walls (only
# the near half shows), then the top faces - so nothing needs depth sorting.
def _shade(hexcol, f):
    h = hexcol.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return "#%02x%02x%02x" % tuple(max(0, min(255, int(c * f)))
                                   for c in (r, g, b))


def ring3d(slices, colours, title="", sub="", width=940, height=430,
           cx=320, cy=175, R=205, r_in=78, k=0.48, depth=40,
           min_deg=4.0):
    """slices is [(label, value)]; colours follow the slice order.

    The largest slice goes to the back and the rest are centred at the
    front, where their walls show. A slice narrower than `min_deg` is DRAWN
    at `min_deg` so it can be seen at all - the caller's note must say so;
    the key and the hover text always carry the true share."""
    tot = sum(v for _, v in slices)
    if not tot:
        return '<p class="fempty">Nothing to show.</p>'
    steps = 90

    def pt(r, a, dy=0.0, ox=0.0, oy=0.0):
        return (cx + ox + r * math.cos(a), cy + oy + dy + r * k * math.sin(a))

    def arc(r, a0, a1, dy=0.0, ox=0.0, oy=0.0):
        n = max(2, int(steps * abs(a1 - a0) / math.pi) + 2)
        return [pt(r, a0 + (a1 - a0) * i / (n - 1), dy, ox, oy)
                for i in range(n)]

    def poly(points, fill, stroke=None, title_txt=""):
        d = " ".join(f"{x:.1f},{y:.1f}" for x, y in points)
        st = f' stroke="{stroke}" stroke-width="1"' if stroke else ""
        tt = f"<title>{esc(title_txt)}</title>" if title_txt else ""
        return f'<polygon points="{d}" fill="{fill}"{st}>{tt}</polygon>'

    spans = [max(2 * math.pi * v / tot, math.radians(min_deg)) if v else 0
             for _, v in slices]
    big_i = max(range(len(slices)), key=lambda i: slices[i][1])
    spans[big_i] -= sum(spans) - 2 * math.pi      # the big slice absorbs it
    rest = sum(sp for i, sp in enumerate(spans) if i != big_i)
    # order: largest first, then the rest - their centre sits at the front
    start = math.pi / 2 - rest / 2 - spans[big_i]
    segs, a = [], start
    for i, (lab, v) in enumerate(slices):
        span = spans[i]
        mid = a + span / 2
        pull = 22 if v / tot < 0.05 else 0
        segs.append((lab, v, a, a + span, colours[i % len(colours)],
                     pull * math.cos(mid), pull * k * math.sin(mid)))
        a += span

    walls, inner, tops = [], [], []
    for lab, v, a0, a1, col, ox, oy in segs:
        tip = f"{lab}: {v:,} of {tot:,} ({v / tot:.1%})"
        # near half of the outer wall: angles where sin > 0 (0..pi)
        for lo, hi in ((2 * math.pi * m, 2 * math.pi * m + math.pi)
                       for m in (-1, 0, 1)):
            b0, b1 = max(a0, lo), min(a1, hi)
            if b1 > b0:
                walls.append(poly(arc(R, b0, b1, 0, ox, oy)
                                  + arc(R, b1, b0, depth, ox, oy),
                                  _shade(col, .72), _shade(col, .6), tip))
        # far half of the hole wall: angles where sin < 0 (pi..2pi)
        for lo, hi in ((2 * math.pi * m - math.pi, 2 * math.pi * m)
                       for m in (-1, 0, 1)):
            b0, b1 = max(a0, lo), min(a1, hi)
            if b1 > b0:
                inner.append(poly(arc(r_in, b0, b1, 0, ox, oy)
                                  + arc(r_in, b1, b0, depth, ox, oy),
                                  _shade(col, .55), _shade(col, .45), tip))
        # pulled-out slices also show their cut faces
        if ox or oy:
            for e in (a0, a1):
                walls.append(poly([pt(r_in, e, 0, ox, oy), pt(R, e, 0, ox, oy),
                                   pt(R, e, depth, ox, oy),
                                   pt(r_in, e, depth, ox, oy)],
                                  _shade(col, .62), _shade(col, .5), tip))
        tops.append(poly(arc(R, a0, a1, 0, ox, oy)
                         + arc(r_in, a1, a0, 0, ox, oy),
                         col, _shade(col, .8), tip))

    out = [f'<svg class="chart" viewBox="0 0 {width} {height}" role="img" '
           f'aria-label="{esc(title or "composition")}">',
           f'<style>{CSS}</style>']
    # floor shadow
    out.append(f'<ellipse cx="{cx}" cy="{cy + depth + 10}" rx="{R + 14}" '
               f'ry="{(R + 14) * k}" fill="#000" opacity=".35"/>')
    out += inner + walls + tops
    # headline on the largest slice's side of the figure
    big = max(segs, key=lambda x: x[1])
    if title:
        out.append(f'<text x="{cx}" y="{cy + depth + R * k + 46:.0f}" '
                   f'text-anchor="middle" style="font-size:15px;'
                   f'font-weight:700;fill:#fff">{esc(title)}</text>')
    if sub:
        out.append(f'<text x="{cx}" y="{cy + depth + R * k + 64:.0f}" '
                   f'text-anchor="middle" class="cn">{esc(sub)}</text>')
    # key: swatch, label, share, count
    kx, ky = cx + R + 70, cy - 70
    for i, (lab, v, *_rest) in enumerate(segs):
        col = colours[i % len(colours)]
        y = ky + i * 62
        out.append(f'<rect x="{kx}" y="{y}" width="16" height="16" rx="3" '
                   f'fill="{col}"/>')
        out.append(f'<text x="{kx + 26}" y="{y + 13}" style="font-size:14px;'
                   f'font-weight:700;fill:#fff">{esc(lab)}</text>')
        out.append(f'<text x="{kx + 26}" y="{y + 36}" style="font-size:24px;'
                   f'font-weight:700;fill:{col}">{v / tot:.1%}</text>')
        out.append(f'<text x="{kx + 120}" y="{y + 36}" class="cn">'
                   f'{v:,} of {tot:,}</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# single-value donut
# ---------------------------------------------------------------
# One share of one whole: a grey track for the whole, a coloured arc for the
# share, the figure in the centre. For a small share this is the honest pie
# - the number carries the value and the arc only shows how little of the
# ring it is; a sliced pie would show a hairline nobody can judge.
def donut1(hits, n, caption="", colour="#c9544a", w=420, h=330, r=118,
           stroke=30, decimals=2):
    v = hits / n if n else 0
    cx, cy = w / 2, h / 2 - 4
    circ = 2 * math.pi * r
    arc = max(circ * v, 3.0) if v > 0 else 0      # a 0.1% arc still shows
    out = [f'<svg class="chart" viewBox="0 0 {w} {h}" role="img" '
           f'aria-label="{v * 100:.{decimals}f} percent">',
           f'<style>{CSS}</style>',
           f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" stroke="{REST}" '
           f'stroke-width="{stroke}"/>']
    if arc:
        out.append(f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="none" '
                   f'stroke="{colour}" stroke-width="{stroke}" '
                   f'stroke-dasharray="{arc:.2f} {circ:.2f}" '
                   f'transform="rotate(-90 {cx} {cy})">'
                   f'<title>{v:.{decimals}%} - {hits:,} of {n:,}</title>'
                   f'</circle>')
    out.append(f'<text x="{cx}" y="{cy + 6}" text-anchor="middle" '
               f'style="font-size:46px;font-weight:700;fill:#fff">'
               f'{v * 100:.{decimals}f}%</text>')
    if caption:
        out.append(f'<text x="{cx}" y="{cy + 32}" text-anchor="middle" '
                   f'class="cn" style="font-size:13px">{esc(caption)}</text>')
    out.append(f'<text x="{cx}" y="{h - 8}" text-anchor="middle" class="cn" '
               f'style="font-size:13px">{hits:,} of {n:,} prescribed '
               f'medicines</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# dumbbell against a benchmark
# ---------------------------------------------------------------
# Each row: a white dot at the benchmark, a coloured dot at the row's own
# value, a grey bar between. Orange where the row falls short of the
# benchmark, blue where it meets or beats it - and the label takes the same
# colour, so the shortfalls can be read down the left edge alone.
DB_LOW = "#e07b3c"
DB_HIGH = "#3f7fd4"


def dumbbell(rows, ref, ref_label="benchmark", unit="", label_w=250,
             row_h=25, pad_t=48, width=940, lo=None, hi=1.0):
    """rows are (label, n, value 0-1), drawn in the order given."""
    if not rows:
        return '<p class="fempty">Nothing to show.</p>'
    vals = [v for _, _, v in rows] + [ref]
    lo = lo if lo is not None else max(0.0, math.floor(min(vals) * 10) / 10)
    h = pad_t + len(rows) * row_h + 22
    plot_w = width - label_w - 70
    x = lambda v: label_w + plot_w * (v - lo) / (hi - lo)
    out = [f'<svg class="chart" viewBox="0 0 {width} {h}" role="img">',
           f'<style>{CSS}</style>']
    step = 0.1 if hi - lo > 0.3 else 0.05
    t = lo
    while t <= hi + 1e-9:
        out.append(f'<line class="gl" x1="{x(t):.1f}" y1="{pad_t - 10}" '
                   f'x2="{x(t):.1f}" y2="{h - 18}"/>')
        out.append(f'<text x="{x(t):.1f}" y="{pad_t - 16}" class="cn" '
                   f'text-anchor="middle">{t * 100:.0f}%</text>')
        t += step
    # benchmark line and key
    out.append(f'<line x1="{x(ref):.1f}" y1="{pad_t - 8}" x2="{x(ref):.1f}" '
               f'y2="{h - 18}" stroke="#fff" stroke-dasharray="3 4" '
               f'opacity=".35"/>')
    out.append(f'<circle cx="{label_w + 8}" cy="14" r="6" fill="#fff" '
               f'stroke="#222" stroke-width="1.5"/>'
               f'<text x="{label_w + 20}" y="18" class="cn">{esc(ref_label)} '
               f'{ref * 100:.1f}%</text>'
               f'<circle cx="{label_w + 230}" cy="14" r="6" fill="{DB_LOW}"/>'
               f'<text x="{label_w + 242}" y="18" class="cn">below</text>'
               f'<circle cx="{label_w + 300}" cy="14" r="6" fill="{DB_HIGH}"/>'
               f'<text x="{label_w + 312}" y="18" class="cn">at or above</text>')
    for i, (lab, n, v) in enumerate(rows):
        y = pad_t + i * row_h + row_h / 2
        col = DB_HIGH if v >= ref else DB_LOW
        out.append(f'<line x1="{label_w}" y1="{y:.1f}" x2="{width - 70}" '
                   f'y2="{y:.1f}" stroke="rgba(255,255,255,.06)"/>')
        out.append(f'<text x="{label_w - 12}" y="{y + 4:.1f}" '
                   f'text-anchor="end" style="font-size:12.5px;'
                   f'font-weight:{700 if v >= ref else 400};fill:{col}">'
                   f'{esc(lab)}</text>')
        out.append(f'<line x1="{x(ref):.1f}" y1="{y:.1f}" x2="{x(v):.1f}" '
                   f'y2="{y:.1f}" stroke="#8a8a86" stroke-width="2.4"/>')
        out.append(f'<circle cx="{x(ref):.1f}" cy="{y:.1f}" r="6.5" '
                   f'fill="#fff" stroke="#222" stroke-width="1.5"/>')
        out.append(f'<circle cx="{x(v):.1f}" cy="{y:.1f}" r="7" fill="{col}" '
                   f'stroke="#111" stroke-width="1.2"><title>{esc(lab)}: '
                   f'{v:.1%} against {ref_label} {ref:.1%} - n={n:,}'
                   f'{(" " + unit) if unit else ""}</title></circle>')
        out.append(f'<text x="{width - 6}" y="{y + 4:.1f}" class="cn" '
                   f'text-anchor="end">{v * 100:.1f}%</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# lollipop
# ---------------------------------------------------------------
# The same values as bars with a fraction of the ink: a heavy baseline, a
# hairline stem, a hollow head at the value. Rows are drawn in the order
# given - the caller sorts, largest first.
def lollipop(rows, label_w=210, row_h=34, pad_t=30, width=940,
             hit="", base="", colour="#4f9be8", show_n=False):
    """rows are (label, n, hits); the plotted value is hits / n. show_n
    writes n=<n> after each value."""
    if not rows:
        return '<p class="fempty">Nothing to show.</p>'
    h = pad_t + len(rows) * row_h + 26
    plot_w = width - label_w - 90
    x0 = label_w
    out = [f'<svg class="chart" viewBox="0 0 {width} {h}" role="img">',
           f'<style>{CSS}</style>']
    for i in (25, 50, 75, 100):
        x = x0 + plot_w * i / 100
        out.append(f'<line class="gl" x1="{x:.1f}" y1="{pad_t - 10}" '
                   f'x2="{x:.1f}" y2="{h - 18}"/>')
        out.append(f'<text x="{x:.1f}" y="{h - 4}" class="cn" '
                   f'text-anchor="middle">{i}%</text>')
    for i, (lab, n, hv) in enumerate(rows):
        v = hv / n if n else 0
        y = pad_t + i * row_h + row_h / 2
        xv = x0 + plot_w * min(v, 1.0)
        out.append(f'<text x="{x0 - 14}" y="{y + 4.5:.1f}" class="ct" '
                   f'text-anchor="end">{esc(lab)}</text>')
        out.append(f'<line x1="{x0}" y1="{y:.1f}" x2="{xv - 7:.1f}" '
                   f'y2="{y:.1f}" stroke="{colour}" stroke-width="2"/>')
        out.append(f'<circle cx="{xv:.1f}" cy="{y:.1f}" r="7" '
                   f'fill="var(--plane,#000)" stroke="{colour}" '
                   f'stroke-width="3"><title>{esc(lab)}: {v:.1%} {esc(hit)} '
                   f'- {hv:,} of {n:,} {esc(base)}</title></circle>')
        out.append(f'<text x="{xv + 14:.1f}" y="{y + 4.5:.1f}" class="cv">'
                   f'{v:.1%}</text>')
        if show_n:
            out.append(f'<text x="{xv + 66:.1f}" y="{y + 4.5:.1f}" '
                       f'class="cn">n={n:,}</text>')
    # the heavy baseline, drawn last so the stems start under it
    out.append(f'<line x1="{x0}" y1="{pad_t - 12}" x2="{x0}" '
               f'y2="{h - 22}" stroke="#e8e8e2" stroke-width="5"/>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# paired dots against a 100% line
# ---------------------------------------------------------------
# Two ratios per row on one axis: a hollow dot for the first (requested /
# consumed), a filled dot for the second (delivered / consumed), a grey bar
# between them - the gap is what was asked for and not received. The dashed
# line at 100% is "in line with use": left of it the stock is drawn down,
# right of it it builds up. The axis runs past 100% because ratios do.
def pairdots(rows, names=("first", "second"), label_w=210, row_h=27,
             pad_t=50, width=940, colour="#4f9be8"):
    """rows are (label, n, ratio_a, ratio_b); a group label starting with
    '--' draws as a section heading instead of a row."""
    if not rows:
        return '<p class="fempty">Nothing to show.</p>'
    vals = [v for r in rows if not r[0].startswith("--") for v in r[2:4]]
    hi = max(1.5, math.ceil(max(vals) * 2) / 2)
    h = pad_t + len(rows) * row_h + 24
    plot_w = width - label_w - 120
    x = lambda v: label_w + plot_w * min(v, hi) / hi
    out = [f'<svg class="chart" viewBox="0 0 {width} {h}" role="img">',
           f'<style>{CSS}</style>']
    t = 0.0
    while t <= hi + 1e-9:
        out.append(f'<line class="gl{" z" if t == 0 else ""}" '
                   f'x1="{x(t):.1f}" y1="{pad_t - 10}" x2="{x(t):.1f}" '
                   f'y2="{h - 20}"/><text x="{x(t):.1f}" y="{h - 5}" '
                   f'class="cn" text-anchor="middle">{t * 100:.0f}%</text>')
        t += 0.5
    out.append(f'<line x1="{x(1):.1f}" y1="{pad_t - 12}" x2="{x(1):.1f}" '
               f'y2="{h - 20}" stroke="#fff" stroke-dasharray="4 4" '
               f'opacity=".55"/>')
    out.append(f'<circle cx="{label_w + 8}" cy="16" r="6" fill="none" '
               f'stroke="{colour}" stroke-width="2.5"/><text '
               f'x="{label_w + 20}" y="20" class="cn">{esc(names[0])}</text>'
               f'<circle cx="{label_w + 230}" cy="16" r="6" '
               f'fill="{colour}"/><text x="{label_w + 242}" y="20" '
               f'class="cn">{esc(names[1])}</text><text x="{x(1):.1f}" '
               f'y="{pad_t - 16}" class="cn" text-anchor="middle">in line '
               f'with use</text>')
    for i, row in enumerate(rows):
        y = pad_t + i * row_h + row_h / 2
        if row[0].startswith("--"):
            out.append(f'<text x="{label_w - 12}" y="{y + 5:.1f}" '
                       f'text-anchor="end" style="font-size:11px;'
                       f'letter-spacing:.08em;fill:#d95926;font-weight:700">'
                       f'{esc(row[0][2:].upper())}</text>')
            continue
        lab, n, a, b = row
        out.append(f'<text x="{label_w - 12}" y="{y + 4.5:.1f}" class="ct" '
                   f'text-anchor="end">{esc(lab)}</text>')
        out.append(f'<line x1="{x(a):.1f}" y1="{y:.1f}" x2="{x(b):.1f}" '
                   f'y2="{y:.1f}" stroke="#8a8a86" stroke-width="2.4"/>')
        out.append(f'<circle cx="{x(a):.1f}" cy="{y:.1f}" r="6.5" '
                   f'fill="#000" stroke="{colour}" stroke-width="2.5">'
                   f'<title>{esc(lab)}: {esc(names[0])} {a:.1%}</title>'
                   f'</circle>')
        out.append(f'<circle cx="{x(b):.1f}" cy="{y:.1f}" r="6.5" '
                   f'fill="{colour}"><title>{esc(lab)}: {esc(names[1])} '
                   f'{b:.1%} - n={n:,} pairs</title></circle>')
        out.append(f'<text x="{width - 6}" y="{y + 4.5:.1f}" class="cn" '
                   f'text-anchor="end">{a:.1%} / {b:.1%}</text>')
    out.append("</svg>")
    return "".join(out)


# ---------------------------------------------------------------
# heat map: rows x columns of rates
# ---------------------------------------------------------------
# One cell per row x column, coloured by the rate on the page's own red ->
# amber -> green scale and labelled with the figure, so a row reads as one
# group across the columns and a column as one column across the groups.
# A cell on fewer than `min_n` observations is drawn grey with its n - it
# exists, but should not be compared. The last column is the row total.
HEAT_STEPS = [(0.00, 0.70, "#c9544a"), (0.70, 0.80, "#e07b3c"),
              (0.80, 0.90, "#d9a21b"), (0.90, 0.95, "#6fae5c"),
              (0.95, 1.01, "#199e70")]
HEAT_LABELS = ["below 70%", "70-79%", "80-89%", "90-94%", "95% and above"]


def heat_colour(v):
    for lo, hi, c in HEAT_STEPS:
        if lo <= v < hi:
            return c
    return HEAT_STEPS[-1][2]


def heatmap(rows, cols, min_n=30, label_w=190, cell_h=40, width=940,
            pad_t=80):
    """rows are (label, [(n, hits) per column]); cols are the headers."""
    if not rows:
        return '<p class="fempty">Nothing to show.</p>'
    k = len(cols)
    gap = 4
    cell_w = (width - label_w - 8 - gap * (k - 1)) / k
    h = pad_t + len(rows) * (cell_h + gap) + 34
    out = [f'<svg class="chart" viewBox="0 0 {width} {h}" role="img">',
           f'<style>{CSS}</style>']
    for i, ((_, _, c), lab) in enumerate(zip(HEAT_STEPS, HEAT_LABELS)):
        x = label_w + i * 120
        out.append(f'<rect x="{x}" y="6" width="12" height="12" rx="2" '
                   f'fill="{c}"/><text x="{x + 17}" y="16" class="cn">'
                   f'{esc(lab)}</text>')
    out.append(f'<rect x="{label_w + 600}" y="6" width="12" height="12" '
               f'rx="2" fill="{REST}"/><text x="{label_w + 617}" y="16" '
               f'class="cn">fewer than {min_n}</text>')
    for j, c in enumerate(cols):
        cx = label_w + j * (cell_w + gap) + cell_w / 2
        words, lines, cur = str(c).split(), [], ""
        for w in words:
            if len(cur) + len(w) + 1 <= 14:
                cur = (cur + " " + w).strip()
            else:
                lines.append(cur)
                cur = w
        lines.append(cur)
        for li, ln in enumerate(lines[:3]):
            out.append(f'<text x="{cx:.1f}" y="{pad_t - 36 + li * 12}" '
                       f'class="ct" text-anchor="middle" '
                       f'style="font-weight:{700 if j == k - 1 else 400}">'
                       f'{esc(ln)}</text>')
        # the column's N: every observation in it, across the rows
        cn = sum(cells[j][0] for _, cells in rows)
        out.append(f'<text x="{cx:.1f}" y="{pad_t - 36 + len(lines[:3]) * 12}"'
                   f' class="cn" text-anchor="middle" style="font-size:10px">'
                   f'n={cn:,}</text>')
    for i, (lab, cells) in enumerate(rows):
        y = pad_t + i * (cell_h + gap)
        # the row's N: its last cell, the all-columns total
        out.append(f'<text x="{label_w - 10}" y="{y + cell_h / 2 - 1:.1f}" '
                   f'class="ct" text-anchor="end">{esc(lab)}</text>'
                   f'<text x="{label_w - 10}" y="{y + cell_h / 2 + 12:.1f}" '
                   f'class="cn" text-anchor="end" style="font-size:10px">'
                   f'n={cells[-1][0]:,}</text>')
        for j, (n, hit) in enumerate(cells):
            x = label_w + j * (cell_w + gap)
            if not n:
                out.append(f'<rect x="{x:.1f}" y="{y}" width="{cell_w:.1f}" '
                           f'height="{cell_h}" rx="4" fill="#1f1f1e"/>')
                continue
            v = hit / n
            thin = n < min_n
            fill = REST if thin else heat_colour(v)
            out.append(f'<rect class="seg" x="{x:.1f}" y="{y}" '
                       f'width="{cell_w:.1f}" height="{cell_h}" rx="4" '
                       f'fill="{fill}"'
                       + (' stroke="#fff" stroke-width="1.5"' if j == k - 1
                          else '')
                       + f'><title>{esc(lab)} - {esc(cols[j])}: {v:.1%} '
                       f'({hit:,} of {n:,})</title></rect>')
            out.append(f'<text x="{x + cell_w / 2:.1f}" '
                       f'y="{y + cell_h / 2 + (0 if thin else 5):.1f}" '
                       f'text-anchor="middle" style="font-size:'
                       f'{12 if thin else 14}px;font-weight:700;fill:#fff;'
                       f'paint-order:stroke;stroke:rgba(0,0,0,.35);'
                       f'stroke-width:2px">{v * 100:.1f}%</text>')
            if thin:
                out.append(f'<text x="{x + cell_w / 2:.1f}" '
                           f'y="{y + cell_h / 2 + 14:.1f}" '
                           f'text-anchor="middle" class="cn" '
                           f'style="font-size:10px">n={n}</text>')
    out.append("</svg>")
    return "".join(out)
