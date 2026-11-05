"""The picture figures of the slide deck, redrawn from the current data.

    python3 deck_figures.py <out_dir>
      -> gauges.png        slide 7   availability by level (managed products)
         map_tracer.png    slide 9   availability by district (managed)
         map_prov.png      slide 9   availability by province (managed)
         map_so.png        slide 20  stock-out incidence by district
         map_chw.png       slide 28  CHW availability by district
      and <name>.html beside each, the page the PNG is shot from.

The deck's pictures were made outside this codebase; this reproduces their
look (palette, class breaks, white district lines, dark province lines,
Trebuchet labels, legend row) so a data refresh can redraw them. The PNGs
are rendered by ~/UP_study_script_backups/tools/shot (WebKit, 2x), at the
pixel size of the picture they replace.
"""

import _bootstrap as B   # path setup - must come first
import io
import contextlib
import math
import os
import re
import subprocess
import sys
from collections import defaultdict

with contextlib.redirect_stdout(io.StringIO()):
    import Full_indicator_report as R
    import Tracer_availability_mapping as MAP

s, M, C = R.s, R.M, R.C
OUT = sys.argv[1] if len(sys.argv) > 1 else "."
SHOT = os.path.expanduser("~/UP_study_script_backups/tools/shot")

NAVY, TEAL, TURQ, LIGHT = "#003e78", "#117996", "#6edbcd", "#d5e7ef"
GOLD, MAROON, GREY, INK = "#f3b71b", "#7c1220", "#e6e6e6", "#1a1a1a"
DARK_FILLS = {NAVY, TEAL, MAROON}
FONT = '"Trebuchet MS","Lucida Grande",Tahoma,sans-serif'
# Kigali's three districts are too small for a plain label: boxed, nudged
BOXED = {"Gasabo", "Nyarugenge", "Kicukiro"}
# label anchors (picture pixels, name baseline) where the centroid does not
# work - taken from the original deck pictures
AT = {"Rubavu": (322, 412), "Rusizi": (110, 1008), "Kamonyi": (670, 666),
      "Nyarugenge": (755, 622), "Kicukiro": (853, 652), "Gasabo": (842, 545),
      "Rutsiro": (330, 578)}


def esc(t):
    return (str(t).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


def pc(v):
    return "-" if v is None else f"{v * 100:.1f}%"


def klass(v, steps):
    """steps: [(upper bound, colour, bold)], last bound None = open."""
    for hi, col, bold in steps:
        if hi is None or v < hi:
            return col, bold
    return steps[-1][1], steps[-1][2]


# ---------------------------------------------------------------
# geometry: dashboard map coordinates -> 1343 x 1222 picture
# ---------------------------------------------------------------
PW, PH, LEG = 1343, 1222, 62
_nums = [float(x) for d in MAP.DROWS
         for x in re.findall(r"-?\d+\.?\d*", d["path"])]
_xs, _ys = _nums[0::2], _nums[1::2]
X0, X1, Y0, Y1 = min(_xs), max(_xs), min(_ys), max(_ys)
K = min((PW - 40) / (X1 - X0), (PH - LEG - 30) / (Y1 - Y0))
OX = (PW - (X1 - X0) * K) / 2 - X0 * K
OY = 14 - Y0 * K


def tp(x, y):
    return x * K + OX, y * K + OY


def render(name, w, h, svg):
    page = (f'<!doctype html><html><head><meta charset="utf-8"><style>'
            f'html,body{{margin:0;background:#fff}}svg{{display:block}}'
            f'text{{font-family:{FONT}}}</style></head><body>'
            f'<svg xmlns="http://www.w3.org/2000/svg" width="{w / 2}" '
            f'height="{h / 2}" viewBox="0 0 {w} {h}">{svg}</svg>'
            f'</body></html>')
    html_p = os.path.join(OUT, name + ".html")
    png_p = os.path.join(OUT, name + ".png")
    open(html_p, "w").write(page)
    r = subprocess.run([SHOT, os.path.abspath(html_p), png_p,
                        str(w / 2), "0", str(h / 2)],
                       capture_output=True, text=True)
    if "ok" not in r.stdout:
        raise SystemExit(f"shot failed for {name}: {r.stdout}{r.stderr}")
    from PIL import Image
    im = Image.open(png_p)
    if im.size != (w, h):
        im.resize((w, h), Image.LANCZOS).save(png_p)
    print("wrote", png_p)


def legend(items, y):
    """items: [(colour, label)] centred in one row."""
    sw, gap, fs = 20, 34, 18
    widths = [sw + 10 + len(lab) * fs * 0.52 for _, lab in items]
    x = (PW - sum(widths) - gap * (len(items) - 1)) / 2
    out = []
    for (col, lab), wd in zip(items, widths):
        out.append(f'<rect x="{x:.1f}" y="{y - 15}" width="{sw}" '
                   f'height="{sw - 4}" fill="{col}"/>')
        out.append(f'<text x="{x + sw + 10:.1f}" y="{y}" font-size="{fs}" '
                   f'fill="#333">{esc(lab)}</text>')
        x += wd + gap
    return "".join(out)


def label(x, y, lines, fill, bold, size=16, boxed=False):
    col = "#fff" if fill in DARK_FILLS and not boxed else INK
    wt = "700" if bold else "400"
    out = []
    if boxed:
        wd = max(len(t) for t in lines) * size * 0.56 + 10
        out.append(f'<rect x="{x - wd / 2:.1f}" y="{y - size - 2:.1f}" '
                   f'width="{wd:.1f}" height="{size * 2 + 10}" rx="3" '
                   f'fill="#fff" fill-opacity=".85"/>')
    for i, t in enumerate(lines):
        out.append(f'<text x="{x:.1f}" y="{y + i * (size + 1):.1f}" '
                   f'font-size="{size}" font-weight="{wt}" fill="{col}" '
                   f'text-anchor="middle">{esc(t)}</text>')
    return "".join(out)


def district_map(name, values, steps, items, shown=None):
    """values: district -> rate or None (no data, grey)."""
    out = [f'<g transform="translate({OX:.3f},{OY:.3f}) scale({K:.5f})">']
    fills = {}
    for d in MAP.DROWS:
        v = values.get(d["district"])
        fill, bold = (GREY, False) if v is None else klass(v, steps)
        fills[d["district"]] = (fill, bold)
        out.append(f'<path d="{d["path"]}" fill="{fill}" stroke="#fff" '
                   f'stroke-width="{2.2 / K:.3f}" stroke-linejoin="round"/>')
    for r in MAP.PROWS:
        out.append(f'<path d="{r["path"]}" fill="none" stroke="{INK}" '
                   f'stroke-width="{3 / K:.3f}" stroke-linejoin="round"/>')
    out.append("</g>")
    for d in MAP.DROWS:
        n = d["district"]
        x, y = tp(*d["cx_cy"])
        v = values.get(n)
        fill, bold = fills[n]
        if v is None:
            out.append(f'<text x="{x:.1f}" y="{y + 6:.1f}" font-size="16" '
                       f'fill="#666" text-anchor="middle">{esc(n)}</text>')
            continue
        txt = (shown or {}).get(n, pc(v))
        x, y = AT.get(n, (x, y))
        out.append(label(x, y, [n, txt], fill, bold,
                         size=14 if n in BOXED else 16, boxed=n in BOXED))
    out.append(legend(items, PH - 22))
    render(name, PW, PH, "".join(out))


def province_map(name, values, steps, items):
    out = [f'<g transform="translate({OX:.3f},{OY:.3f}) scale({K:.5f})">']
    for r in MAP.PROWS:
        fill, _ = klass(values[r["prov"]], steps)
        out.append(f'<path d="{r["path"]}" fill="{fill}" stroke="{INK}" '
                   f'stroke-width="{3 / K:.3f}" stroke-linejoin="round"/>')
    out.append("</g>")
    for r in MAP.PROWS:
        p = r["prov"]
        fill, _ = klass(values[p], steps)
        x, y = tp(*MAP._prov_centroid(p))
        if p == "Kigali City":
            out.append(label(x, y, [p, pc(values[p])], fill, True, size=16,
                             boxed=True))
        else:
            out.append(label(x, y, [p, pc(values[p])], fill, True, size=22))
    out.append(legend(items, PH - 22))
    render(name, PW, PH + 3, "".join(out))


# ---------------------------------------------------------------
# gauges (slide 7)
# ---------------------------------------------------------------
def gauge(cx, cy, v, col, txt_col):
    r, a0, sweep = 162, 225, 270

    def pt(rad, frac):
        a = math.radians(a0 - sweep * frac)
        return cx + rad * math.cos(a), cy - rad * math.sin(a)

    def arc(f0, f1, colour, wd):
        x0, y0 = pt(r, f0)
        x1, y1 = pt(r, f1)
        big = 1 if (f1 - f0) * sweep > 180 else 0
        return (f'<path d="M{x0:.1f},{y0:.1f} A{r},{r} 0 {big} 1 '
                f'{x1:.1f},{y1:.1f}" fill="none" stroke="{colour}" '
                f'stroke-width="{wd}" stroke-linecap="round"/>')
    out = [arc(v, 1, "#d9d9d9", 30), arc(0, v, col, 30)]
    for k in range(41):
        f = k / 40
        major = k % 10 == 0
        r1, r2 = (182, 202) if major else (187, 197)
        x1, y1 = pt(r1, f)
        x2, y2 = pt(r2, f)
        out.append(f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" '
                   f'y2="{y2:.1f}" stroke="{col if f <= v else "#bfbfbf"}" '
                   f'stroke-width="{4.5 if major else 2.5}"/>')
        if major:
            lx, ly = pt(220, f)
            out.append(f'<text x="{lx:.1f}" y="{ly + 6:.1f}" font-size="16" '
                       f'fill="#595959" text-anchor="middle">'
                       f'{int(f * 100)}</text>')
    nx, ny = pt(128, v)
    out.append(f'<line x1="{cx}" y1="{cy}" x2="{nx:.1f}" y2="{ny:.1f}" '
               f'stroke="#333" stroke-width="4" stroke-linecap="round"/>')
    out.append(f'<circle cx="{cx}" cy="{cy}" r="13" fill="#f2f2f2" '
               f'stroke="#333" stroke-width="3.5"/>')
    out.append(f'<text x="{cx}" y="{cy + 92}" font-size="58" '
               f'font-weight="700" fill="{txt_col}" text-anchor="middle">'
               f'{v * 100:.1f}%</text>')
    out.append(f'<text x="{cx}" y="{cy + 150}" font-size="19" fill="#595959"'
               f' text-anchor="middle">of medicines expected</text>')
    return "".join(out)


def gauges(name, rows):
    """rows: [(title, n, available)] in reading order; the lowest of the
    levels is gold, the last (all levels) teal, the others navy."""
    w, h = 1645, 1170
    cw, ch, gx, gy = 508, 552, 560, 606
    low = min(rows[:-1], key=lambda t: t[2] / t[1])
    out = []
    for i, (title, n, a) in enumerate(rows):
        x, y = 6 + (i % 3) * gx + 0.5, 6 + (i // 3) * gy
        v = a / n
        col = TEAL if i == len(rows) - 1 else (GOLD if (title, n, a) == low
                                               else NAVY)
        out.append(f'<rect x="{x}" y="{y}" width="{cw}" height="{ch}" '
                   f'rx="18" fill="#f2f2f2"/>')
        out.append(gauge(x + cw / 2, y + 236, v, col,
                         INK if col == GOLD else NAVY))
        out.append(f'<text x="{x + cw / 2}" y="{y + 450}" font-size="27" '
                   f'font-weight="700" fill="{INK}" text-anchor="middle">'
                   f'{esc(title)}</text>')
        out.append(f'<text x="{x + cw / 2}" y="{y + 505}" font-size="21" '
                   f'fill="#404040" text-anchor="middle">{a:,} of {n:,} '
                   f'expected</text>')
    render(name, w, h, "".join(out))


# ---------------------------------------------------------------
# data
# ---------------------------------------------------------------
def rate(rows, ok):
    return sum(1 for r in rows if ok(r)) / len(rows) if rows else None


B_ = R.MOLS_B
yes = lambda r: M.yes(r[M.AVAIL])
by = defaultdict(list)
for r in B_:
    by[s(r["_p"][M.DIST])].append(r)
AV_D = {k: rate(v, yes) for k, v in by.items()}
byp = defaultdict(list)
for r in B_:
    byp[s(r["_p"][M.PROV])].append(r)
AV_P = {k: rate(v, yes) for k, v in byp.items()}

so = defaultdict(list)
for r in R.MG:
    so[s(r["_p"][M.DIST])].append(r)
SO_D = {k: rate(v, lambda r: bool(R._somonths(r))) for k, v in so.items()}

ch = defaultdict(list)
for r in R.CANS:
    ch[C.s(r["_p"][C.DIST])].append(r)
CHW_D = {k: rate(v, lambda r: C.yes(r[C.AVAIL])) for k, v in ch.items()}

LEVEL_TITLES = {"Health centre": "Health center",
                "District hospital": "District hospital*"}
GAUGE_ROWS = [(LEVEL_TITLES.get(l, l), len(rs), sum(map(yes, rs)))
              for l in R.LEVELS for rs in [[r for r in B_ if r["_lvl"] == l]]
              if rs]
GAUGE_ROWS.append(("All levels", len(B_), sum(map(yes, B_))))

if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    gauges("gauges", GAUGE_ROWS)
    district_map("map_tracer", AV_D,
                 [(0.8995, GOLD, True), (0.9295, LIGHT, False),
                  (0.9495, TURQ, False), (0.9695, TEAL, False),
                  (None, NAVY, False)],
                 [(GOLD, "below 90%"), (LIGHT, "90 to 92.9%"),
                  (TURQ, "93 to 94.9%"), (TEAL, "95 to 96.9%"),
                  (NAVY, "97% and above")])
    province_map("map_prov", AV_P,
                 [(0.9295, LIGHT, True), (0.9395, TURQ, True),
                  (None, TEAL, True)],
                 [(LIGHT, "92 to 92.9%"), (TURQ, "93 to 93.9%"),
                  (TEAL, "94 to 94.9%")])
    district_map("map_so", SO_D,
                 [(0.0495, NAVY, False), (0.0795, TEAL, False),
                  (0.0995, LIGHT, False), (0.1495, GOLD, False),
                  (None, MAROON, True)],
                 [(NAVY, "below 5%"), (TEAL, "5 to 7.9%"),
                  (LIGHT, "8 to 9.9%"), (GOLD, "10 to 14.9%"),
                  (MAROON, "15% and above")],
                 shown={"Nyagatare": "10-14.9%*"}
                 if 0.0995 <= (SO_D.get("Nyagatare") or 0) < 0.1495 else None)
    district_map("map_chw", {d["district"]: CHW_D.get(d["district"])
                             for d in MAP.DROWS},
                 [(0.5995, MAROON, True), (0.6995, GOLD, False),
                  (0.7995, LIGHT, False), (0.8995, TEAL, False),
                  (None, NAVY, False)],
                 [(MAROON, "below 60%"), (GOLD, "60 to 69.9%"),
                  (LIGHT, "70 to 79.9%"), (TEAL, "80 to 89.9%"),
                  (NAVY, "90% and above"), (GREY, "no CHW data")])
    for t, n, a in GAUGE_ROWS:
        print(f"  gauge {t:<28} {a:,} of {n:,} = {a / n:.1%}")
    print("  provinces", {k: pc(v) for k, v in sorted(AV_P.items())})
    print("  SO Nyagatare", pc(SO_D.get("Nyagatare")),
          "| SO top", sorted(((pc(v), k) for k, v in SO_D.items()),
                             reverse=True)[:4])
