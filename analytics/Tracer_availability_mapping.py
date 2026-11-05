"""Tracer medicine availability mapped: districts and facilities.

    python3 Tracer_availability_mapping.py
      -> Tracer_availability_map.html    district choropleth + facility points
      -> Tracer_availability_map.xlsx    the underlying tables and coordinates

BOUNDARIES
There is no shapefile in the project folder, so Rwanda's 30 district
polygons were downloaded from geoBoundaries ADM2, whose source is Open Data
Rwanda / NISR under CC BY 4.0. The file and its provenance are kept in
analytics/gis/. All 30 district names match the assessment data exactly, so
the join needs no fuzzy matching. If you have the official NISR shapefile,
drop a GeoJSON at the same path and this script will use it instead.

FACILITY COORDINATES
Taken from "Facility GIS_.xlsx" (sheets HC, HP, HOSP, DH Referrals, NRH,
HC Referral), joined to the assessed facilities on a normalised name within
district. Where the GIS file has no match, the submission's own GPS point is
used as a fallback and the source is labelled per facility, so nothing is
silently imputed and nothing is silently dropped.

AVAILABILITY
Availability = available on visit day / (managed AND availability answered),
the same definition as every other output here. "All medicines availability
(%)" for a facility is its own basket score: the share of the medicines
assessed at that facility that were physically present.

Plain SVG, no mapping library, so the page works offline.
"""
import _bootstrap as B

import json
import math
import os
import re
import statistics
from collections import Counter, defaultdict

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_medicine_assessment as M
import facility_levels as FL
import molecules as MOL

OUT_HTML = B.out("Tracer_availability_map.html")
OUT_XLSX = B.out("Tracer_availability_map.xlsx")
GEOJSON = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "gis", "rwanda_districts_adm2.geojson")
GIS_FILE = "Facility GIS_.xlsx"
GIS_SHEETS = ["HC", "HC Referral", "HP", "HOSP", "DH Referrals", "NRH"]
MIN_N = 10


def s(v):
    return "" if v is None else str(v).strip()


def pct(a, b):
    return a / b if b else None


def norm(n):
    """facility name key: drop level suffix, punctuation and case"""
    t = s(n).lower().replace("__", " ")
    t = re.sub(r"\(.*?\)", " ", t)
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    toks = [x for x in t.split()
            if x not in ("cs", "hc", "hp", "sghp", "dh", "ph", "mhc", "l2th",
                         "th", "rh", "hospital", "health", "center", "centre",
                         "post")]
    return "".join(toks)


def ntoks(n):
    """Same cleaning as norm(), but kept as a TOKEN LIST.

    norm() joins the tokens with no separator, which is right for an exact
    key and useless for asking whether one name sits inside another, so the
    health-post fallback needs this form.
    """
    t = s(n).lower().replace("__", " ").replace("-", " ")
    t = re.sub(r"\(.*?\)", " ", t)
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    return [x for x in t.split()
            if x not in ("cs", "hc", "hp", "sghp", "dh", "ph", "mhc", "l2th",
                         "th", "rh", "hospital", "health", "center", "centre",
                         "post", "de", "la", "notre", "st", "saint")]


# ---------------------------------------------------------------
# availability
# ---------------------------------------------------------------
# Method 3: multi-strength medicines collapsed to one molecule per facility
# DEFINITION (study, 2026-10-13) - two tracks, set by medicine_program:
# Essential Medicines on every molecule the EML authorises at that level
# (not managed counts as UNAVAILABLE); every other programme on the
# molecules the facility manages. MOL.two_track_base() applies both.
ELIG = [r for r in MOL.two_track_base(M.MEDROWS)
        if M.yes(r[M.MANAGED]) and not M.blank(r[M.AVAIL])]
for r in ELIG:
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])

FAC = defaultdict(list)
for r in ELIG:
    FAC[s(r["_p"][M.FAC])].append(r)

DIST = defaultdict(list)
for r in ELIG:
    DIST[s(r["_p"][M.DIST])].append(r)

NAT = pct(sum(1 for r in ELIG if M.yes(r[M.AVAIL])), len(ELIG))

# ---------------------------------------------------------------
# coordinates
# ---------------------------------------------------------------
def load_gis():
    path = os.path.join(B.ROOT, GIS_FILE)
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out = {}
    for sh in GIS_SHEETS:
        if sh not in wb.sheetnames:
            continue
        rows = list(wb[sh].iter_rows(values_only=True))
        hdr = [s(x) for x in rows[0]]
        def idx(*names):
            for n in names:
                for i, h in enumerate(hdr):
                    if h.lower() == n:
                        return i
            return None
        i_n = idx("hf", "facility name", "health center", "hospital name")
        i_d = idx("district", "district name")
        i_la, i_lo = idx("latitude"), idx("longitude")
        if None in (i_n, i_la, i_lo):
            continue
        for r in rows[1:]:
            name = s(r[i_n])
            try:
                la, lo = float(r[i_la]), float(r[i_lo])
            except (TypeError, ValueError):
                continue
            if not (-3.1 < la < -0.9 and 28.8 < lo < 30.95):
                continue
            key = norm(name)
            dist = s(r[i_d]).lower() if i_d is not None else ""
            out.setdefault((key, dist), (la, lo, name, sh))
            out.setdefault((key, ""), (la, lo, name, sh))
    return out


GIS = load_gis()
SUB_GPS = {}
for r in M.SUB:
    la = M.num(r.get("_start-geopoint_latitude"))
    lo = M.num(r.get("_start-geopoint_longitude"))
    if la and lo and -3.1 < la < -0.9 and 28.8 < lo < 30.95:
        SUB_GPS.setdefault(s(r[M.FAC]), (la, lo))

# Health posts are named inconsistently between the assessment and the GIS
# file: the assessment writes "KIBERINKA - GAHONDO HP" or "Cyeza Notre Damma
# de Lourde HP" where the GIS sheet carries the bare village, "GAHONDO" and
# "Cyeza". An exact normalised match misses those, which is why five posts
# were unmapped. This tier looks inside the SAME DISTRICT for an entry whose
# whole name appears as a run of tokens in the facility name, and accepts it
# only when exactly one entry qualifies - an ambiguous match is left unmapped
# rather than guessed, because a post placed in the wrong village is worse on
# a map than a post that is absent from it.
GIS_BY_DIST = defaultdict(list)
for (key, dist), v in GIS.items():
    if dist:
        GIS_BY_DIST[dist].append((ntoks(v[2]), v))


def token_match(fac, dist):
    toks = ntoks(fac)
    cands = []
    for kt, v in GIS_BY_DIST.get(dist.lower(), []):
        if not kt or sum(len(x) for x in kt) < 5:
            continue
        n = len(kt)
        if any(toks[i:i + n] == kt for i in range(len(toks) - n + 1)):
            cands.append(v)
    uniq = {(round(v[0], 5), round(v[1], 5)): v for v in cands}
    return next(iter(uniq.values())) if len(uniq) == 1 else None


# ---------------------------------------------------------------
# coordinates supplied by hand
# ---------------------------------------------------------------
# Four facilities had no usable position: three are absent from the GIS file
# under any matchable name, and BWIZA HP matched a different BWIZA 90km away
# in Gisagara, which the containment guard below rejected. None of the four
# carried a GPS reading on its submission.
#
# Supplied by the user 2026-10-13 and checked before use: each point falls
# inside the polygon of the district the assessment recorded for it. They
# take precedence over the GIS file, so the bad BWIZA match cannot come back.
MANUAL_COORDS = {
    ("BURUNGA HP", "Karongi"): (-2.059019, 29.371271),
    ("Gasenga HP", "Bugesera"): (-2.148523, 30.106181),
    ("BWIZA HP", "Kayonza"): (-1.881564, 30.495838),
    ("Gikongoro Prison CS", "Nyamagabe"): (-2.4723, 29.5815),
}

POINTS, UNLOCATED = [], []
AMBIGUOUS = []
for f, rs in sorted(FAC.items()):
    dist = s(rs[0]["_p"][M.DIST])
    av = sum(1 for r in rs if M.yes(r[M.AVAIL]))
    manual = MANUAL_COORDS.get((f, dist))
    hit = (((manual[0], manual[1], "supplied by hand", "SUPPLIED") if manual
            else None)
           or GIS.get((norm(f), dist.lower())) or GIS.get((norm(f), ""))
           or token_match(f, dist))
    if hit:
        la, lo, matched, sh = hit
        src = f"Facility GIS ({sh})"
    elif f in SUB_GPS:
        la, lo = SUB_GPS[f]
        matched, src = "", "submission GPS"
    else:
        UNLOCATED.append([f, dist, rs[0]["_lvl"], len(rs), av,
                          pct(av, len(rs))])
        continue
    POINTS.append({"fac": f, "dist": dist, "lvl": rs[0]["_lvl"],
                   "n": len(rs), "av": av, "pct": pct(av, len(rs)),
                   "la": la, "lo": lo, "src": src, "matched": matched})

# ---------------------------------------------------------------
# geometry
# ---------------------------------------------------------------
GJ = json.load(open(GEOJSON))
W, H, PAD = 980, 760, 24


# ---------------------------------------------------------------
# containment guard
# ---------------------------------------------------------------
# A name match is only as good as the name. The GIS file has a BWIZA in
# Gisagara and the assessment has a BWIZA HP in Kayonza, and the
# district-blind fallback happily matched the two, dropping the point 90km
# from the facility. Every NAME-matched point is therefore tested against
# the polygon of the district the assessment recorded, and a point that
# lands outside it is discarded rather than drawn: a marker in the wrong
# district is a false statement about where a medicine is missing, while a
# facility absent from the map is only a gap the caption already declares.
#
# Points that come from the submission's own GPS are left alone. Those are
# measurements, not guesses, and the one that fails this test (MBABE HP)
# sits a kilometre over the Kicukiro/Nyarugenge line - an honest reading
# near a boundary, not a mismatched name.
def _rings(g):
    return [g["coordinates"]] if g["type"] == "Polygon" else g["coordinates"]


def _inside(lo, la, geom):
    for poly in _rings(geom):
        ring = poly[0]
        c, n = False, len(ring)
        for i in range(n):
            x1, y1 = ring[i][:2]
            x2, y2 = ring[(i + 1) % n][:2]
            if (y1 > la) != (y2 > la) and lo < (x2 - x1) * (la - y1) / (y2 - y1) + x1:
                c = not c
        if c:
            return True
    return False


_GEOM = {s(f["properties"]["shapeName"]).lower(): f["geometry"]
         for f in GJ["features"]}
MISPLACED = []
for _p in list(POINTS):
    _g = _GEOM.get(_p["dist"].lower())
    if _p["src"] in ("submission GPS", "Facility GIS (SUPPLIED)") or not _g:
        continue
    if not _inside(_p["lo"], _p["la"], _g):
        MISPLACED.append([_p["fac"], _p["dist"], _p["lvl"], _p["matched"]])
        POINTS.remove(_p)
        UNLOCATED.append([_p["fac"], _p["dist"], _p["lvl"], _p["n"], _p["av"],
                          _p["pct"]])


def bounds():
    xs, ys = [], []
    for f in GJ["features"]:
        for poly in (f["geometry"]["coordinates"]
                     if f["geometry"]["type"] == "MultiPolygon"
                     else [f["geometry"]["coordinates"]]):
            for ring in poly:
                for x, y in ring:
                    xs.append(x)
                    ys.append(y)
    return min(xs), max(xs), min(ys), max(ys)


X0, X1, Y0, Y1 = bounds()
MIDLAT = math.radians((Y0 + Y1) / 2)
SX = (W - 2 * PAD) / ((X1 - X0) * math.cos(MIDLAT))
SY = (H - 2 * PAD) / (Y1 - Y0)
SC = min(SX, SY)


def proj(lon, lat):
    """equirectangular with a cos(lat) correction - fine at this extent"""
    x = PAD + (lon - X0) * math.cos(MIDLAT) * SC
    y = PAD + (Y1 - lat) * SC
    return x, y


def path_for(feature):
    d = []
    polys = (feature["geometry"]["coordinates"]
             if feature["geometry"]["type"] == "MultiPolygon"
             else [feature["geometry"]["coordinates"]])
    for poly in polys:
        for ring in poly:
            pts = [proj(x, y) for x, y in ring]
            d.append("M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in pts) + " Z")
    return " ".join(d)


def centroid(feature):
    polys = (feature["geometry"]["coordinates"]
             if feature["geometry"]["type"] == "MultiPolygon"
             else [feature["geometry"]["coordinates"]])
    best, area = None, -1
    for poly in polys:
        ring = poly[0]
        a = abs(sum(ring[i][0] * ring[i + 1][1] - ring[i + 1][0] * ring[i][1]
                    for i in range(len(ring) - 1)) / 2)
        if a > area:
            area, best = a, ring
    xs = [p[0] for p in best]
    ys = [p[1] for p in best]
    return proj(sum(xs) / len(xs), sum(ys) / len(ys))


# Colour scale: 5 steps, sequential blue, DARKEST = BEST.
#
# Two things were wrong before and both are fixed here.
#
# 1. DIRECTION. The ramp ran dark-for-worst, so on a map titled "availability"
#    the heaviest ink marked the least of the thing being mapped. A sequential
#    ramp reads as "more ink, more of it", so the ramp is reversed: the darkest
#    districts are now the best stocked.
#
# 2. BREAKS, recalibrated 2026-10-12 after the S6/S7 cleaning.
#    They have been reset twice, each time because the distribution moved
#    under them, which is the point: a fixed ramp on round numbers stops
#    describing the data the moment the data changes.
#      - 85/90/93/96 were set when availability was measured over managed
#        rows and sat in the nineties.
#      - 72/76/80/85 were set for the EML denominator, when the districts ran
#        67.5% to 94.7%.
#      - The EML applicability correction and the zero-quantity correction
#        moved the districts down to 65.0%-83.9%, median 74.5%. That emptied
#        the top class completely and crowded 12 of 30 districts into one
#        band, so the map was again losing most of its resolution.
#      - 70/73/76/79 were quintiles of that spread (3/8/8/8/3 districts).
#      - 2026-10-13 the definition became two-track (Essential Medicines on
#        the EML denominator, programme products on managed rows), which
#        lifted the districts to 80.1%-92.3%, national 85.7%.
#    The cuts below are near-quintiles of that spread: 7/4/9/6/4 districts,
#    with the national figure inside the middle class.
STEPS = [(0.00, 0.83, "#d7e4f7"), (0.83, 0.85, "#9cc0ee"),
         (0.85, 0.87, "#5693e0"), (0.87, 0.89, "#2a68bd"),
         (0.89, 1.01, "#15407c")]
LABELS = ["below 83%", "83 to 84%", "85 to 86%", "87 to 88%", "89% and above"]

# Method 2 (managed products only) runs higher - districts 87.0%-97.5% in
# Nov 2026 - so it has its own breaks (user request 2026-11-04).
# Classes hold 6/4/6/10/4 districts.
# Cuts sit 0.05 points low so a district labelled "97.0%" (96.95-96.99)
# is coloured with 97% and above, as its label reads.
# Same blue ramp as Method 1 (user request 2026-11-05) - only the breaks,
# and so the legend labels, change with the method.
STEPS_B = [(lo, hi, c) for (lo, hi), (_, _, c) in zip(
    [(0.00, 0.8995), (0.8995, 0.9295), (0.9295, 0.9495), (0.9495, 0.9695),
     (0.9695, 1.01)], STEPS)]
LABELS_B = ["below 90%", "90 to 92.9%", "93 to 94.9%", "95 to 96.9%",
            "97% and above"]


def colour(v):
    if v is None:
        return "#2c2c2a"
    for lo, hi, c in STEPS:
        if lo <= v < hi:
            return c
    return STEPS[-1][2]


# The facility bubbles sit on a near-black canvas, so they cannot share the
# choropleth ramp: now that dark means best, the best-stocked facilities would
# be drawn in #15407c on #0e0e0d and vanish. The bubbles get their own ramp
# with the same direction of meaning - MORE LIGHT, MORE AVAILABILITY - running
# from a dim slate to a bright aqua, which stays legible at 3px.
PT_STEPS = [(0.00, 0.83, "#4a5a6e"), (0.83, 0.85, "#2f7fb5"),
            (0.85, 0.87, "#2a9fc4"), (0.87, 0.89, "#2fc39e"),
            (0.89, 1.01, "#6ef0c4")]


def pt_colour(v):
    if v is None:
        return "#3a3a38"
    for lo, hi, c in PT_STEPS:
        if lo <= v < hi:
            return c
    return PT_STEPS[-1][2]


# ---------------------------------------------------------------
# district table
# ---------------------------------------------------------------
DROWS = []
for f in GJ["features"]:
    d = s(f["properties"]["shapeName"])
    rs = DIST.get(d, [])
    av = sum(1 for r in rs if M.yes(r[M.AVAIL]))
    facs = {s(r["_p"][M.FAC]) for r in rs}
    scores = [pct(sum(1 for x in FAC[x0] if M.yes(x[M.AVAIL])), len(FAC[x0]))
              for x0 in facs for x in [None]][:0] or [
        pct(sum(1 for x in FAC[ff] if M.yes(x[M.AVAIL])), len(FAC[ff]))
        for ff in facs]
    worst = sorted(({s(r[M.MED]) for r in rs}),
                   key=lambda m: pct(
                       sum(1 for x in rs if s(x[M.MED]) == m
                           and M.yes(x[M.AVAIL])),
                       sum(1 for x in rs if s(x[M.MED]) == m)) or 0)
    DROWS.append({
        "district": d, "prov": s(rs[0]["_p"][M.PROV]) if rs else "",
        "facs": len(facs), "n": len(rs), "av": av, "miss": len(rs) - av,
        "pct": pct(av, len(rs)) if len(rs) >= MIN_N else None,
        "mean_fac": statistics.mean(scores) if scores else None,
        "lowest_fac": min(scores) if scores else None,
        "complete": sum(1 for ff in facs
                        if all(M.yes(x[M.AVAIL]) for x in FAC[ff])),
        "worst_med": worst[0][:44] if worst else "",
        "path": path_for(f), "cx_cy": centroid(f)})
DROWS.sort(key=lambda r: (r["pct"] is None, r["pct"]))


# ---------------------------------------------------------------
# province dissolve
# ---------------------------------------------------------------
# The boundary file is ADM2 only - it carries no province attribute - so the
# province geometry is DISSOLVED from the districts rather than loaded. The
# district-to-province mapping comes from the assessment's own submissions,
# where all 30 districts carry exactly one province label and every district
# in the file appears in the data, so the mapping is complete and unambiguous.
#
# The dissolve works on edges: a segment that two districts of the SAME
# province share is interior and is dropped; what remains is the province
# outline. geoBoundaries stores shared edges with identical coordinates
# (2,815 of 3,830 segments appear exactly twice, none more), so this is an
# exact operation and not a tolerance-based merge.
D2P = {}
for _r in ELIG:
    D2P[s(_r["_p"][M.DIST])] = s(_r["_p"][M.PROV])

PROV_ORDER = ["Kigali City", "North", "South", "East", "West"]


def _segments(feature):
    g = feature["geometry"]
    polys = [g["coordinates"]] if g["type"] == "Polygon" else g["coordinates"]
    for poly in polys:
        for ring in poly:
            for i in range(len(ring) - 1):
                a = tuple(round(c, 7) for c in ring[i][:2])
                b = tuple(round(c, 7) for c in ring[i + 1][:2])
                yield a, b


def _chain(edges):
    """Walk undirected edges into closed rings."""
    adj = defaultdict(list)
    for a, b in edges:
        adj[a].append(b)
        adj[b].append(a)
    used, rings = set(), []
    for start in list(adj):
        if not adj[start]:
            continue
        ring, cur, prev = [start], start, None
        while True:
            nxt = None
            for cand in adj[cur]:
                e = tuple(sorted((cur, cand)))
                if cand != prev and e not in used:
                    nxt = cand
                    break
            if nxt is None:
                break
            used.add(tuple(sorted((cur, nxt))))
            ring.append(nxt)
            prev, cur = cur, nxt
            if nxt == start:
                break
        if len(ring) > 3:
            rings.append(ring)
    return rings


def province_paths():
    by_prov = defaultdict(list)
    for f in GJ["features"]:
        by_prov[D2P.get(s(f["properties"]["shapeName"]), "")].append(f)
    out = {}
    for prov, feats in by_prov.items():
        cnt = Counter()
        for f in feats:
            for a, b in _segments(f):
                cnt[tuple(sorted((a, b)))] += 1
        border = [e for e, k in cnt.items() if k == 1]
        d = []
        for ring in _chain(border):
            pts = [proj(x, y) for x, y in ring]
            d.append("M" + " L".join(f"{x:.1f},{y:.1f}" for x, y in pts) + " Z")
        out[prov] = " ".join(d)
    return out


PPATH = province_paths()

PROWS = []
for prov in PROV_ORDER:
    rs = [r for r in ELIG if s(r["_p"][M.PROV]) == prov]
    if not rs:
        continue
    av = sum(1 for r in rs if M.yes(r[M.AVAIL]))
    facs = {s(r["_p"][M.FAC]) for r in rs}
    dists = {s(r["_p"][M.DIST]) for r in rs}
    gap = sum(1 for r in rs if r.get("_gap"))
    PROWS.append({"prov": prov, "dists": len(dists), "facs": len(facs),
                  "n": len(rs), "av": av, "gap": gap,
                  "miss": len(rs) - av, "pct": pct(av, len(rs)),
                  "path": PPATH.get(prov, "")})

# ---------------------------------------------------------------
# Excel
# ---------------------------------------------------------------
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HF = PatternFill("solid", fgColor="1F3864")
HFONT = Font(bold=True, size=9, color="FFFFFF")
BAD = PatternFill("solid", fgColor="FCE4E4")
MID = PatternFill("solid", fgColor="FFF4D6")
OK = PatternFill("solid", fgColor="E7F3E7")


def xsheet(wb, title, headers, rows, notes=None, pctcols=(), widths=None):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    for n in (notes or []):
        c = ws.cell(row=r, column=1, value=n)
        c.font = Font(size=9, italic=True, color="555555")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 13 * (1 + len(n) // 140)
        r += 1
    if notes:
        r += 1
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font, c.fill, c.border = HFONT, HF, BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = 34
    hdr = r
    r += 1
    for row in rows:
        for ci, v in enumerate(row, start=1):
            c = ws.cell(row=r, column=ci, value=v)
            c.border = BORDER
            c.alignment = Alignment(vertical="top",
                                    wrap_text=isinstance(v, str)
                                    and len(str(v)) > 32)
            if (ci - 1) in pctcols and isinstance(v, float):
                c.number_format = "0.0%"
                c.fill = BAD if v < .85 else (MID if v < .95 else OK)
            elif isinstance(v, float):
                c.number_format = "0.000000"
            elif isinstance(v, int):
                c.number_format = "#,##0"
        r += 1
    ws.freeze_panes = f"B{hdr + 1}"
    ws.auto_filter.ref = f"A{hdr}:{get_column_letter(len(headers))}{r - 1}"
    for ci, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(ci)].width = w


def write_xlsx():
    wb = Workbook()
    wb.remove(wb.active)
    xsheet(wb, "District availability",
           ["District", "Province", "Facilities assessed",
            "Eligible observations", "Available", "Missing",
            "TRACER AVAILABILITY", "Mean facility score",
            "Lowest facility score", "Facilities with a complete basket",
            "Weakest medicine in the district"],
           [[d["district"], d["prov"], d["facs"], d["n"], d["av"], d["miss"],
             d["pct"], d["mean_fac"], d["lowest_fac"], d["complete"],
             d["worst_med"]] for d in DROWS],
           notes=[f"All 30 districts, worst first. National tracer "
                  f"availability {NAT:.1%} over {len(ELIG):,} eligible "
                  f"observations.",
                  "Availability = available on visit day / (managed AND "
                  "answered). Boundaries: geoBoundaries ADM2, source Open "
                  "Data Rwanda / NISR, CC BY 4.0."],
           pctcols=(6, 7, 8), widths={1: 16, 2: 13, 3: 12, 4: 13, 5: 10,
                                      6: 9, 7: 13, 8: 13, 9: 13, 10: 14,
                                      11: 46})
    xsheet(wb, "Facility availability and GPS",
           ["Facility", "District", "Province", "Level of care",
            "Medicines assessed", "Available", "Missing",
            "ALL MEDICINES AVAILABILITY", "Latitude", "Longitude",
            "Coordinate source", "Matched GIS name"],
           [[p["fac"], p["dist"],
             s(FAC[p["fac"]][0]["_p"][M.PROV]), p["lvl"], p["n"], p["av"],
             p["n"] - p["av"], p["pct"], p["la"], p["lo"], p["src"],
             p["matched"]]
            for p in sorted(POINTS, key=lambda x: x["pct"])],
           notes=[f"{len(POINTS)} of {len(FAC)} assessed facilities located. "
                  f"Coordinates come from Facility GIS_.xlsx where the name "
                  f"matched within district, otherwise from the submission's "
                  f"own GPS point - the source is named per row.",
                  "'All medicines availability' is the facility's own basket "
                  "score: medicines present / medicines assessed there."],
           pctcols=(7,), widths={1: 32, 2: 15, 3: 13, 4: 24, 5: 12, 6: 10,
                                 7: 9, 8: 14, 9: 12, 10: 12, 11: 24, 12: 28})
    if UNLOCATED:
        xsheet(wb, "Facilities without coordinates",
               ["Facility", "District", "Level of care", "Medicines assessed",
                "Available", "ALL MEDICINES AVAILABILITY"], UNLOCATED,
               notes=["Assessed but not locatable: no name match in the "
                      "Facility GIS file and no GPS on the submission. They "
                      "are excluded from the map but NOT from the district "
                      "figures, which are computed from all observations.",
                      "Adding these to the GIS file, or enabling the GPS "
                      "question, would complete the map."],
               pctcols=(5,), widths={1: 34, 2: 15, 3: 24, 4: 12, 5: 10, 6: 14})
    prov = defaultdict(lambda: [0, 0])
    for r in ELIG:
        p = s(r["_p"][M.PROV])
        prov[p][1] += 1
        if M.yes(r[M.AVAIL]):
            prov[p][0] += 1
    xsheet(wb, "Province summary",
           ["Province", "Districts", "Facilities", "Eligible observations",
            "Available", "TRACER AVAILABILITY"],
           sorted([[p, len({d["district"] for d in DROWS if d["prov"] == p}),
                    len({s(r["_p"][M.FAC]) for r in ELIG
                         if s(r["_p"][M.PROV]) == p}), v[1], v[0],
                    pct(v[0], v[1])] for p, v in prov.items()],
                  key=lambda x: x[5]), pctcols=(5,),
           widths={1: 14, 2: 10, 3: 11, 4: 13, 5: 10, 6: 13})
    wb.active = 0
    if os.path.exists(OUT_XLSX):
        os.chmod(OUT_XLSX, 0o644)
    wb.save(OUT_XLSX)


# ---------------------------------------------------------------
# HTML map
# ---------------------------------------------------------------
import html as _html


def esc(v):
    return _html.escape("" if v is None else str(v), quote=True)


def fmt(v):
    if v is None:
        return "n/a"
    if isinstance(v, float) and 0 <= v <= 1:
        return f"{v * 100:.1f}%"
    return f"{v:,}" if isinstance(v, int) else str(v)


CSS = """
<style>
:root{color-scheme:dark;--plane:#000;--surface:#1a1a19;--raise:#222221;
 --ink:#fff;--ink2:#c3c2b7;--muted:#898781;--rule:rgba(255,255,255,.14);
 --grid:#2c2c2a;--accent:#2a78d6;--bad:#e66767;--mid:#eda100;--good:#199e70;

 --font:"Trebuchet MS","Lucida Grande","Lucida Sans Unicode",Tahoma,sans-serif;}
*{box-sizing:border-box}
body{margin:0;background:var(--plane);color:var(--ink);font-family:var(--font);
 font-size:16px;line-height:1.55}
.wrap{max-width:1300px;margin:0 auto;padding:0 20px 72px}
h1,h2,h4{margin:0;font-family:var(--font);text-wrap:balance}
.mast{border-bottom:1px solid var(--rule);padding:40px 0 24px}
.eyebrow{font-size:12.5px;letter-spacing:.14em;text-transform:uppercase;
 color:var(--muted);margin:0 0 12px;font-weight:600}
.mast h1{font-size:clamp(26px,4.4vw,40px);letter-spacing:-.015em}
.mast p{color:var(--ink2);max-width:86ch;margin:14px 0 0;font-size:15px}
.tabs{display:flex;gap:2px;margin:28px 0 0;border-bottom:1px solid var(--rule)}
.tab{font:inherit;font-size:14.5px;color:var(--ink2);background:none;border:0;
 border-bottom:2px solid transparent;padding:11px 16px;margin-bottom:-1px;
 cursor:pointer}
.tab:hover{color:var(--ink)}
.tab.on{color:var(--ink);font-weight:600;border-bottom-color:var(--accent)}
.js .page{display:none}.js .page.on{display:block}
.page{margin-top:28px}
.maprow{display:grid;grid-template-columns:minmax(0,1fr) 300px;gap:20px;
 align-items:start}
.mapbox{border:1px solid var(--rule);background:var(--surface);padding:6px}
svg{display:block;width:100%;height:auto}
.dist{stroke:#000;stroke-width:1;cursor:default}
.dist:hover{stroke:#fff;stroke-width:2}
.mtitle{fill:#fff;font-size:15px;font-weight:700}
.msub{fill:#c3c2b7;font-size:11px}
.msrc{fill:#898781;font-size:9.5px}
.layers{margin:10px 0 0;padding:12px 14px;border:1px solid var(--rule);
 border-radius:8px;background:#141413}
.lyrh{margin:0 0 8px;font-size:11px;letter-spacing:.08em;text-transform:uppercase;
 color:var(--accent);font-weight:700}
.lyr{display:flex;align-items:center;gap:8px;font-size:12.5px;padding:3px 0;
 cursor:pointer}
.lyr input{accent-color:var(--accent);cursor:pointer}
.lyrn{margin-left:auto;color:var(--muted);font-variant-numeric:tabular-nums}
.lyrbtn{margin:10px 0 0;display:flex;gap:6px}
.lyrbtn button{flex:1;font:inherit;font-size:11.5px;padding:5px 8px;cursor:pointer;
 color:var(--ink);background:#1f1f1e;border:1px solid var(--rule);border-radius:5px}
.lyrbtn button:hover{border-color:var(--accent);color:var(--accent)}
.lyrsel{width:100%;font:inherit;font-size:12.5px;padding:6px 8px;color:var(--ink);
 background:#1f1f1e;border:1px solid var(--rule);border-radius:5px;cursor:pointer}
.lyrsel:disabled{opacity:.4;cursor:not-allowed}
{LAYER_CSS}
.mlab{fill:#c3c2b7;font-size:10px}
.slegend rect,.sbar rect{shape-rendering:crispEdges}
.dlabel{fill:#fff;font-size:9.5px;text-anchor:middle;pointer-events:none;
 paint-order:stroke;stroke:#000;stroke-width:2.4px;stroke-linejoin:round}
.border{stroke-linejoin:round}
.border:hover{stroke:#fff;stroke-width:1.8}
.pt{stroke:#000;stroke-width:.8;cursor:default}
.pt:hover{stroke:#fff;stroke-width:1.8}
.legend{border:1px solid var(--rule);background:var(--surface);padding:16px}
.legend h4{font-size:14px;margin:0 0 10px}
.key{display:flex;align-items:center;gap:9px;font-size:13px;color:var(--ink2);
 margin:0 0 6px}
.sw{width:22px;height:12px;border:1px solid #000;flex:0 0 auto}
.dot{border-radius:50%;border:1px solid #000}
.legend p{font-size:12.5px;color:var(--muted);margin:12px 0 0;line-height:1.55}
.scroll{overflow-x:auto;border:1px solid var(--rule);background:var(--surface);
 margin-top:22px}
.scroll.tall{max-height:460px;overflow-y:auto}
table{border-collapse:collapse;width:100%;font-size:13.5px}
thead th{position:sticky;top:0;background:var(--raise);text-align:left;
 font-weight:600;font-size:12.5px;color:var(--ink2);padding:8px 11px;
 border-bottom:1px solid #383835;white-space:nowrap;z-index:1}
td{padding:6px 11px;border-bottom:1px solid var(--grid);vertical-align:top;
 color:var(--ink2)}
td.num{text-align:right;font-variant-numeric:tabular-nums;color:var(--ink);
 white-space:nowrap}
tbody tr:hover td{background:#141413}
.bar{display:inline-block;height:9px;border-radius:99px;margin-left:8px}
.note{font-size:13px;color:var(--muted);margin:10px 0 0;max-width:96ch}
@media (max-width:900px){.maprow{grid-template-columns:1fr}}
@media print{.js .page{display:block!important}.tabs{display:none}
 body{background:#fff;color:#000}}
</style>
"""


def legend(kind):
    """Side panel: the colour ramp, plus a size note on the facility map
    where the bubbles also carry the value in their area."""
    ramp = PT_STEPS if kind == "fac" else STEPS

    def keys(ramp, labels):
        return "".join(
            f'<p class="key"><span class="sw" style="background:{c}"></span>'
            f'{esc(l)}</p>' for (lo, hi, c), l in zip(ramp, labels))
    rows = keys(ramp, LABELS)
    if kind != "fac":
        # one key per availability method; the page's method switch shows
        # the matching one (explorer.py's only-a / only-b classes)
        rows = (f'<div class="only-a"><p class="key" style="font-weight:600">'
                f'Expected medicines</p>{rows}</div>'
                f'<div class="only-b"><p class="key" style="font-weight:600">'
                f'Managed products only</p>{keys(STEPS_B, LABELS_B)}</div>')
    if kind == "fac":
        rows += ('<p class="key" style="margin-top:12px">'
                 '<span class="sw dot" style="width:18px;height:18px;'
                 'background:#6ef0c4"></span>large and bright = high '
                 'availability</p>'
                 '<p class="key"><span class="sw dot" style="width:7px;'
                 'height:7px;background:#4a5a6e"></span>small and dim = low '
                 'availability</p>')
    # The three explanatory paragraphs that used to sit under the ramp were
    # removed on request. The boundary credit still appears in each map's own
    # footer (district_svg / facility_svg), which is what CC BY 4.0 requires.
    return f'<div class="legend"><h4>Tracer availability</h4>{rows}</div>'


def _prov_centroid(prov):
    """Label anchor: the mean of the member districts' own centroids."""
    pts = [d["cx_cy"] for d in DROWS if d["prov"] == prov]
    return (sum(p[0] for p in pts) / len(pts),
            sum(p[1] for p in pts) / len(pts)) if pts else (W / 2, H / 2)


def district_svg():
    out = [f'<svg id="distmap" viewBox="0 0 {W} {H + 54}" role="img" '
           f'aria-label="Rwanda tracer availability by district or province">']
    out.append(f'<text x="{PAD}" y="18" class="mtitle">Tracer medicine '
               f'availability</text>')
    out.append(f'<text x="{PAD}" y="34" class="msub">Pooled availability '
               f'&#183; {len(DROWS)} districts in {len(PROWS)} provinces '
               # explorer.py rewrites this for the method and filters shown
               f'&#183; national <tspan id="fx-natl">{fmt(NAT)}</tspan>'
               f'</text>')

    # --- district layer
    out.append('<g id="lyr-dist">')
    for d in DROWS:
        out.append(f'<path class="dist pv-{slug(d["prov"])}" '
                   f'data-d="{esc(d["district"])}" d="{d["path"]}" '
                   f'fill="{colour(d["pct"])}"><title>{esc(d["district"])}'
                   f' district, {esc(d["prov"])}'
                   f' - {fmt(d["pct"])} available\n'
                   f'{d["facs"]} facilities, {d["n"]:,} observations, '
                   f'{d["miss"]} missing\nweakest: {esc(d["worst_med"])}'
                   f'</title></path>')
    for d in DROWS:
        cx, cy = d["cx_cy"]
        out.append(f'<text class="dlabel pv-{slug(d["prov"])}" '
                   f'x="{cx:.0f}" y="{cy:.0f}">{esc(d["district"])}</text>')
        out.append(f'<text class="dlabel pv-{slug(d["prov"])}" x="{cx:.0f}" '
                   f'y="{cy + 11:.0f}" style="font-size:10.5px" '
                   f'data-dv="{esc(d["district"])}">{fmt(d["pct"])}</text>')
    out.append('</g>')

    # --- province layer, drawn from the dissolved outlines
    out.append('<g id="lyr-prov">')
    for r in PROWS:
        out.append(f'<path class="dist" data-p="{esc(r["prov"])}" '
                   f'd="{r["path"]}" '
                   f'fill="{colour(r["pct"])}"><title>{esc(r["prov"])}'
                   f' province - {fmt(r["pct"])} available\n'
                   f'{r["dists"]} districts, {r["facs"]} facilities, '
                   f'{r["n"]:,} observations\n{r["gap"]:,} not managed, '
                   f'{r["miss"] - r["gap"]:,} managed but empty'
                   f'</title></path>')
    for r in PROWS:
        cx, cy = _prov_centroid(r["prov"])
        out.append(f'<text class="dlabel" x="{cx:.0f}" y="{cy:.0f}" '
                   f'style="font-size:13px">{esc(r["prov"])}</text>')
        out.append(f'<text class="dlabel" x="{cx:.0f}" y="{cy + 14:.0f}" '
                   f'style="font-size:12px" data-pv="{esc(r["prov"])}">'
                   f'{fmt(r["pct"])}</text>')
    out.append('</g>')

    out.append(north_arrow(W - PAD - 16, 66))
    out.append(scale_bar(PAD + 4, H + 10))
    out.append(f'<text x="{PAD + 4}" y="{H + 36}" class="msrc">Boundaries: '
               f'geoBoundaries ADM2 &#183; source Open Data Rwanda / NISR '
               f'&#183; CC BY 4.0. Equirectangular projection, cos(lat) '
               f'corrected.</text>')
    out.append(f'<text x="{W - PAD - 4}" y="{H + 36}" class="msrc" '
               f'text-anchor="end">CHAI Rwanda &#183; UP study</text>')
    out.append("</svg>")
    return "".join(out)


# Facility scores cluster in a narrow band, so an area scale anchored at
# zero makes every bubble the same size - 9.2px at 60% against 11.0px at
# 100%, which encodes nothing the eye can read. The scale is therefore
# stretched across the OBSERVED range instead, and the size key is labelled
# with real values so it stays decodable. This exaggerates differences
# relative to a from-zero scale, which is the trade for making the variable
# visible at all; it is stated on the map.
R_MIN, R_MAX = 3.0, 15.0
_vals = [p["pct"] for p in POINTS if p["pct"] is not None]
# Anchor the bubble scale on the 5th percentile, not the minimum. One facility
# sits near 29% while the middle half runs 71-88%, and anchoring on that single
# outlier compressed every other facility into the top of the scale. Values
# below the anchor draw at R_MIN; the key says so.
_sv = sorted(_vals)
V_LO = _sv[int(0.05 * (len(_sv) - 1))] if _sv else 0.0
V_HI = _sv[-1] if _sv else 1.0


def radius(v):
    """Marker radius encodes AVAILABILITY - large bubble, well stocked.

    Area is proportional to the position within the observed range, because
    the eye reads a circle by area rather than radius.
    """
    if v is None:
        return R_MIN
    if V_HI <= V_LO:
        return (R_MIN + R_MAX) / 2
    frac = max(0.0, min(1.0, (v - V_LO) / (V_HI - V_LO)))
    return R_MIN + (R_MAX - R_MIN) * math.sqrt(frac)


def scale_bar(x, y, km=50):
    """Scale bar. SC is pixels per degree, and one degree of latitude is
    110.574 km at this latitude, which is accurate enough for a country this
    size on an equirectangular projection."""
    px = km * SC / 110.574
    return (f'<g class="sbar"><rect x="{x}" y="{y}" width="{px:.1f}" '
            f'height="5" fill="#c3c2b7" stroke="#000" stroke-width=".6"/>'
            f'<rect x="{x}" y="{y}" width="{px / 2:.1f}" height="5" '
            f'fill="#1a1a19" stroke="#000" stroke-width=".6"/>'
            f'<text x="{x}" y="{y - 5}" class="mlab">0</text>'
            f'<text x="{x + px:.1f}" y="{y - 5}" class="mlab" '
            f'text-anchor="middle">{km} km</text></g>')


def north_arrow(x, y):
    return (f'<g class="narr"><path d="M{x},{y - 20} L{x - 6},{y} '
            f'L{x},{y - 5} L{x + 6},{y} Z" fill="#c3c2b7" stroke="#000" '
            f'stroke-width=".6"/><text x="{x}" y="{y + 13}" class="mlab" '
            f'text-anchor="middle">N</text></g>')


def size_legend(x, y):
    """Bubble-size key, labelled with observed values so the stretched scale
    can still be decoded."""
    span = V_HI - V_LO
    ticks = [V_HI, V_LO + span * 0.66, V_LO + span * 0.33, V_LO]
    out = [f'<g class="slegend"><text x="{x}" y="{y}" class="mlab" '
           f'style="font-size:10.5px;fill:#fff">Availability</text>']
    cx = x + 17
    for i, v in enumerate(ticks):
        r = radius(v)
        cy = y + 22 + i * 32
        out.append(f'<circle cx="{cx}" cy="{cy}" r="{r:.1f}" '
                   f'fill="{pt_colour(v)}" stroke="#000" stroke-width=".8"/>')
        out.append(f'<text x="{cx + 20}" y="{cy + 3.5}" class="mlab">'
                   f'{v * 100:.0f}%</text>')
    out.append(f'<text x="{x}" y="{y + 22 + len(ticks) * 32 + 4}" '
               f'class="msrc">scaled across the observed</text>')
    out.append(f'<text x="{x}" y="{y + 22 + len(ticks) * 32 + 15}" '
               f'class="msrc">range {V_LO * 100:.0f}-{V_HI * 100:.0f}%, '
               f'not from zero</text>')
    out.append("</g>")
    return "".join(out)


def slug(x):
    return re.sub(r"[^a-z0-9]+", "-", s(x).lower()).strip("-")


LVL_COUNTS = Counter(p["lvl"] for p in POINTS)


def layer_control():
    """Checkbox per level of care, so one layer can be read on its own.

    Health posts are 24 of 264 points and the smallest bubbles on the map,
    so they disappear under the health centres; turning the other layers off
    is the only way to read them. The toggle is a class on the <svg> and a
    CSS rule per level - no script runs to draw anything, so the map still
    renders correctly if the file is opened somewhere script is blocked, it
    simply starts with every layer on.
    """
    boxes = []
    for lvl in FL.ordered(LVL_COUNTS):
        sg = slug(lvl)
        boxes.append(
            f'<label class="lyr"><input type="checkbox" checked '
            f'data-lv="{sg}" onchange="facLayer(this)"> {esc(lvl)} '
            f'<span class="lyrn">{LVL_COUNTS[lvl]}</span></label>')
    return (f'<div class="layers"><p class="lyrh">Layers &#183; level of '
            f'care</p>{"".join(boxes)}'
            f'<p class="lyrbtn"><button type="button" '
            f'onclick="facOnly(null)">all</button>'
            f'<button type="button" onclick="facOnly(\'health-post\')">'
            f'health posts only</button></p></div>'
            f'<script>\n'
            f'function facLayer(cb){{\n'
            f'  document.getElementById("facmap").classList'
            f'.toggle("off-"+cb.dataset.lv, !cb.checked);\n'
            f'}}\n'
            f'function facOnly(only){{\n'
            f'  document.querySelectorAll(".layers input[data-lv]")'
            f'.forEach(function(cb){{\n'
            f'    cb.checked = (only === null) || (cb.dataset.lv === only);\n'
            f'    facLayer(cb);\n'
            f'  }});\n'
            f'}}\n'
            f'</script>')


def view_control():
    """District / province view switch, plus a province filter.

    Two different questions. The VIEW radio changes the unit the map is
    drawn in - 30 districts or 5 dissolved provinces, same colour ramp so
    the two are directly comparable. The FILTER narrows the district view
    to one province, which is the only way to read the small Kigali City
    districts, and it is disabled in province view where it would mean
    nothing. As with the facility layers this is CSS plus a class on the
    <svg>; with script blocked the map still renders, showing districts.
    """
    opts = "".join(f'<option value="{slug(r["prov"])}">{esc(r["prov"])}'
                   f' &#183; {fmt(r["pct"])}</option>' for r in PROWS)
    return (f'<div class="layers"><p class="lyrh">View</p>'
            f'<label class="lyr"><input type="radio" name="dview" checked '
            f'onchange="distView(0)"> Districts '
            f'<span class="lyrn">{len(DROWS)}</span></label>'
            f'<label class="lyr"><input type="radio" name="dview" '
            f'onchange="distView(1)"> Provinces '
            f'<span class="lyrn">{len(PROWS)}</span></label>'
            f'<p class="lyrh" style="margin-top:14px">Filter &#183; '
            f'province</p>'
            f'<select id="pvfilter" class="lyrsel" onchange="distProv(this)">'
            f'<option value="">All provinces</option>{opts}</select>'
            f'<script>\n'
            f'function distView(p){{\n'
            f'  var m=document.getElementById("distmap");\n'
            f'  m.classList.toggle("view-prov", !!p);\n'
            f'  document.getElementById("pvfilter").disabled = !!p;\n'
            f'}}\n'
            f'function distProv(sel){{\n'
            f'  var m=document.getElementById("distmap");\n'
            f'  Array.prototype.slice.call(m.classList)\n'
            f'    .filter(function(c){{return c.indexOf("only-")===0}})\n'
            f'    .forEach(function(c){{m.classList.remove(c)}});\n'
            f'  if(sel.value) m.classList.add("only-"+sel.value);\n'
            f'}}\n'
            f'</script></div>')


def layer_css():
    css = ["".join(f'#facmap.off-{slug(l)} .pt.lv-{slug(l)}{{display:none}}'
                   for l in LVL_COUNTS)]
    css.append('#distmap #lyr-prov{display:none}'
               '#distmap.view-prov #lyr-prov{display:inline}'
               '#distmap.view-prov #lyr-dist{display:none}')
    for r in PROWS:
        sg = slug(r["prov"])
        # Districts outside the chosen province are dimmed, not removed: the
        # country outline has to stay or the selected province floats with no
        # geographic anchor.
        css.append(f'#distmap.only-{sg} .dist:not(.pv-{sg})'
                   f'{{fill:#1a1a19;opacity:.45}}'
                   f'#distmap.only-{sg} .dlabel:not(.pv-{sg}){{display:none}}')
    return "".join(css)


def facility_svg():
    """Facility bubbles over district borders, with the map furniture the
    framework expects: title, size key, colour key, scale bar, north arrow
    and a source credit."""
    out = [f'<svg id="facmap" viewBox="0 0 {W} {H + 54}" role="img" '
           f'aria-label="All-medicines availability by facility, Rwanda">']
    # title block
    out.append(f'<text x="{PAD}" y="18" class="mtitle">All-medicines '
               f'availability by facility</text>')
    out.append(f'<text x="{PAD}" y="34" class="msub">Share of assessed tracer '
               f'medicines physically present on the visit day &#183; '
               f'{len(POINTS)} facilities &#183; bubble size and colour both '
               f'show availability, scaled across the observed '
               f'{V_LO * 100:.0f}-{V_HI * 100:.0f}% range</text>')
    # district borders, drawn visibly - a point means nothing until you can
    # see which district it falls in
    for d in DROWS:
        out.append(f'<path class="border" d="{d["path"]}" fill="#0e0e0d" '
                   f'stroke="#6a6864" stroke-width="1.1">'
                   f'<title>{esc(d["district"])} district - '
                   f'{fmt(d["pct"])} tracer availability, '
                   f'{d["facs"]} facilities assessed</title></path>')
    for d in DROWS:
        cx, cy = d["cx_cy"]
        out.append(f'<text class="dlabel" x="{cx:.0f}" y="{cy:.0f}" '
                   f'style="font-size:9.5px;fill:#c3c2b7">'
                   f'{esc(d["district"])}</text>')
    # markers, largest first so a small struggling facility is never buried
    for p in sorted(POINTS, key=lambda x: -(x["pct"] or 0)):
        x, y = proj(p["lo"], p["la"])
        out.append(f'<circle class="pt lv-{slug(p["lvl"])}" cx="{x:.1f}" '
                   f'cy="{y:.1f}" data-f="{esc(p["fac"])}" '
                   f'r="{radius(p["pct"]):.1f}" fill="{pt_colour(p["pct"])}">'
                   f'<title>{esc(p["fac"])} ({esc(p["lvl"])})\n'
                   f'{esc(p["dist"])} district\n'
                   f'{fmt(p["pct"])} available - {p["av"]} of {p["n"]} '
                   f'medicines\nsource: {esc(p["src"])}</title></circle>')
    # furniture
    out.append(size_legend(PAD + 4, 78))
    out.append(north_arrow(W - PAD - 16, 66))
    out.append(scale_bar(PAD + 4, H + 10))
    out.append(f'<text x="{PAD + 4}" y="{H + 36}" class="msrc">Boundaries: '
               f'geoBoundaries ADM2 &#183; source Open Data Rwanda / NISR '
               f'&#183; CC BY 4.0. Facility positions: Facility GIS_.xlsx, '
               f'with submission GPS as fallback and four points supplied by '
               f'hand. Equirectangular projection, '
               f'cos(lat) corrected.</text>')
    out.append(f'<text x="{W - PAD - 4}" y="{H + 36}" class="msrc" '
               f'text-anchor="end">CHAI Rwanda &#183; UP study</text>')
    out.append("</svg>")
    return "".join(out)


def table(headers, rows, pctcol=None):
    body = []
    for row in rows:
        tds = []
        for i, v in enumerate(row):
            num = isinstance(v, (int, float)) and not isinstance(v, bool)
            cell = esc(fmt(v) if not isinstance(v, float) or 0 <= v <= 1
                       else f"{v:.6f}")
            if i == pctcol and isinstance(v, float):
                w = max(2, round(v * 46))
                cell += (f'<span class="bar" style="width:{w}px;'
                         f'background:{colour(v)}"></span>')
            tds.append(f'<td class="{"num" if num else ""}">{cell}</td>')
        body.append("<tr>" + "".join(tds) + "</tr>")
    head = "".join(f"<th>{esc(h)}</th>" for h in headers)
    tall = " tall" if len(rows) > 18 else ""
    return (f'<div class="scroll{tall}"><table><thead><tr>{head}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody></table></div>')


def write_html():
    dtab = table(
        ["District", "Province", "Facilities", "Observations", "Missing",
         "Tracer availability", "Mean facility score", "Weakest medicine"],
        [[d["district"], d["prov"], d["facs"], d["n"], d["miss"], d["pct"],
          d["mean_fac"], d["worst_med"]] for d in DROWS], pctcol=5)
    ftab = table(
        ["Facility", "District", "Level of care", "Medicines assessed",
         "Missing", "All medicines availability", "Coordinate source"],
        [[p["fac"], p["dist"], p["lvl"], p["n"], p["n"] - p["av"], p["pct"],
          p["src"]] for p in sorted(POINTS, key=lambda x: x["pct"])],
        pctcol=5)
    unl = (f'<p class="note">{len(UNLOCATED)} assessed facilities could not '
           f'be located - no name match in the Facility GIS file and no GPS '
           f'on the submission. They are absent from the map but included in '
           f'every district figure. The list is in the workbook.</p>'
           if UNLOCATED else "")
    body = f"""
<div class="wrap">
<header class="mast">
  <p class="eyebrow">CHAI Rwanda &middot; UP study &middot; tracer availability mapping</p>
  <h1>Tracer medicine availability across Rwanda</h1>
  <p>District choropleth and facility points, both showing the same measure:
  the share of EXPECTED medicines physically present on the visit day -
  for Essential Medicines every one the EML authorises at that level of care,
  for programme products the ones the facility manages. National figure
  {fmt(NAT)} across {len(DROWS)} districts, {len(FAC)} facilities and
  {len(ELIG):,} facility-medicine observations.</p>
</header>
<nav class="tabs">
  <button class="tab on" id="tab-d" data-page="d">Districts</button>
  <button class="tab" id="tab-f" data-page="f">Facilities</button>
</nav>
<div class="page on" id="d">
  <div class="maprow"><div class="mapbox">{district_svg()}</div>
  <div>{legend("d")}{view_control()}</div></div>
  {dtab}
  <p class="note">Each district is shaded by its pooled availability and
  labelled with the figure. Hover for the facility count, the number of
  missing observations and the weakest medicine in that district.</p>
</div>
<div class="page" id="f">
  <div class="maprow"><div class="mapbox">{facility_svg()}</div>
  <div>{legend("fac")}{layer_control()}</div></div>
  {ftab}{unl}
  <p class="note">One marker per located facility; size and brightness both
  show its own basket score. Use the layer box to read a single level of care
  on its own - health posts are the smallest bubbles and sit under the health
  centres until the other layers are switched off. Hover for the facility,
  its level of care and the coordinate source.</p>
</div>
</div>
<script>
(function(){{
 var ids=["d","f"];
 function show(id){{ids.forEach(function(o){{
   document.getElementById(o).classList.toggle('on',o===id);
   document.getElementById('tab-'+o).classList.toggle('on',o===id);}});}}
 document.addEventListener('click',function(e){{
   var b=e.target.closest('[data-page]'); if(!b)return;
   show(b.dataset.page); history.replaceState(null,'','#'+b.dataset.page);}});
 var h=location.hash.slice(1); if(ids.indexOf(h)>=0) show(h);
}})();
</script>
"""
    doc = ('<!doctype html>\n<html lang="en" class="js">\n<head>\n'
           '<meta charset="utf-8">\n<meta name="viewport" '
           'content="width=device-width,initial-scale=1">\n'
           '<title>Tracer availability - Rwanda districts and facilities</title>\n'
           + CSS.replace("{LAYER_CSS}", layer_css())
           + "</head>\n<body>\n" + body + "\n</body>\n</html>\n")
    if os.path.exists(OUT_HTML):
        os.chmod(OUT_HTML, 0o644)
    with open(OUT_HTML, "w", encoding="utf-8") as fh:
        fh.write(doc)


if __name__ == "__main__":
    write_xlsx()
    write_html()
    print(f"wrote {os.path.basename(OUT_HTML)} and "
          f"{os.path.basename(OUT_XLSX)}")
    print(f"  national tracer availability {NAT:.1%} over {len(ELIG):,} "
          f"eligible observations")
    print(f"  districts mapped: {len(DROWS)} | facilities located "
          f"{len(POINTS)} of {len(FAC)} "
          f"({sum(1 for p in POINTS if p['src'].startswith('Facility GIS') and 'SUPPLIED' not in p['src'])} "
          f"from the GIS file, "
          f"{sum(1 for p in POINTS if 'SUPPLIED' in p['src'])} supplied by "
          f"hand, "
          f"{sum(1 for p in POINTS if p['src'] == 'submission GPS')} from "
          f"submission GPS) | unlocated {len(UNLOCATED)}")
    print("  worst five districts:")
    for d in DROWS[:5]:
        print(f"    {d['district']:14} {fmt(d['pct']):>7}  "
              f"{d['facs']:>3} facilities, weakest: {d['worst_med'][:40]}")


def scales_js():
    """The colour and size scales, as data the dashboard's script can read.

    The maps are redrawn in the browser when a filter changes, and a redraw
    that used a different ramp from the build-time render would silently
    contradict the legend beside it. So the scales are exported from here
    rather than restated in JavaScript.
    """
    return json.dumps({
        "steps": [[lo, hi, c] for lo, hi, c in STEPS],
        "steps_b": [[lo, hi, c] for lo, hi, c in STEPS_B],
        "pt": [[lo, hi, c] for lo, hi, c in PT_STEPS],
        "none": "#2c2c2a", "ptnone": "#3a3a38",
        "rmin": R_MIN, "rmax": R_MAX, "vlo": V_LO, "vhi": V_HI,
        "minn": MIN_N,
        "prov": [r["prov"] for r in PROWS],
        "levels": FL.ordered({r["_lvl"] for r in ELIG}),
        "dist": {d["district"]: d["prov"] for d in DROWS},
    }, separators=(",", ":"))
