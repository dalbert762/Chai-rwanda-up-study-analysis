"""Product-level availability by level of care, district and province.

    python3 Product_availability_by_geography.py
      -> Product_availability_by_geography.xlsx

The existing availability analysis gives availability BY level, BY province and
BY district (summed over all medicines), and separately BY medicine (summed
over the country). This file crosses the two: for every product, availability
at each level of care, in each province and in each district.

Availability is the share of observations where the facility MANAGES the
medicine and answered the availability question - never over all observations,
or a medicine the facility has chosen not to stock counts as a stock-out.

Every cell carries its own denominator in an adjacent column, and any cell
below MIN_N answered rows is left blank rather than shown as a percentage: a
product x district cell averages four observations, and a rate on four
observations is noise. The counts are still there to be summed.

Level of care is reported two ways. The recorded values split health posts
across three interchangeable labels ("Second Generation Health Post (SGHP)",
"Health Post (HP)" and a bare "HP"), which makes three thin tiers out of one -
so a tidied LEVEL GROUP is used for the crosstabs, with the recorded values
kept on their own sheet.

openpyxl + the standard library, like the rest of this project.
"""

import _bootstrap as B   # path setup - must come first
import os
import statistics
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_medicine_assessment as M
import analyze_medicine_assessment_chw as C

OUT = B.out("Product_availability_by_geography.xlsx")
MIN_N = 10        # below this a cell shows counts only, no percentage

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, size=9, color="FFFFFF")
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")
BAD_FILL = PatternFill("solid", fgColor="FCE4E4")
MID_FILL = PatternFill("solid", fgColor="FFF4D6")
OK_FILL = PatternFill("solid", fgColor="E7F3E7")
TITLE_FONT = Font(bold=True, size=12, color="1F3864")

# Level tiers come from facility_levels, so every output in this project
# groups them the same way: health posts merged into one tier, the maternity
# health centre folded in with the health centres, hospitals kept apart.
import facility_levels as FL
import molecules as MOL

def s(v):
    return "" if v is None else str(v).strip()


def pct(a, b):
    return a / b if b else None


COMMUNITY = "Community (CHW)"

# The CHW assessment is a separate tool with its own 13-item basket, so its
# rows are brought in as a tier BELOW the health post rather than appended
# blindly. Eight products exist in both baskets; those eight are what makes a
# like-for-like comparison across all tiers possible, and they are the only
# basis on which the community tier should be read against the others.
#
# Not shared: ORS, Zinc 10mg and Amox 125mg dispersible are CHW-only (the
# facility basket carries none of them), and ALu 2x6 and 3x6 are CHW-only
# presentations. So the CHW basket is NOT a subset of the facility basket.
SHARED = {
    "ACT (ALu) First Presentation:ARTE 20MG+LUME 120MG DISP TAB (1X6) B/30":
        "ARTE 20MG+LUME 120MG (1X6)",
    "ACT (ALu) Fourth Presentation: ARTE 20MG + LUME 120MG TAB (4X6) B/30":
        "ARTE 20MG + LUME 120MG (4X6)",
    "Rapid diagnostic tests: MALARIA RAPID DIAGNOSTICS Ag P.f/Pan 30T":
        "MALARIA RAPID DIAGNOSTICS Ag P.f/Pan 30T",
    "Dihydroartemisinin/Piperaquine (DHAP) 20/160 mg Disp Tablet":
        "Dihydroartemisinin/Piperaquine (DHAP) 20/160 mg Disp Tablet",
    "Dihydroartemisinin/Piperaquine DHAP) 30/240 mg Dispersible Tablet":
        "Dihydroartemisinin/Piperaquine DHAP) 30/240 mg Dispersible Tablet",
    "Dihydroartemisinin/Piperaquine DHAP) 80/640 mg Dispersible Tablet":
        "Dihydroartemisinin/Piperaquine DHAP) 80/640 mg Dispersible Tablet",
    "Pyronaridine/Artesunate 180/60 mg Film-Coated Tablet":
        "Pyronaridine/Artesunate 180/60 mg Film-Coated Tablet",
    "Pyronaridine/Artesunate 60/20 mg Granules for Suspension":
        "Pyronaridine/Artesunate 60/20 mg Granules for Suspension",
}


def chw_rows():
    """CHW rows shaped like facility rows, so one table can hold both."""
    out = []
    for r in C.MEDROWS:
        if not (C.yes(r[C.MANAGED]) and not C.blank(r[C.AVAIL])):
            continue
        name = str(r[C.MED]).strip()
        out.append({
            M.MED: SHARED.get(name, name),
            M.AVAIL: r[C.AVAIL],
            M.PROGRAM: "CHW package",
            "_lvl": COMMUNITY,
            "_shared": name in SHARED,
            "_p": {M.FAC: C.s(r["_p"][C.FAC]), M.PROV: C.s(r["_p"][C.PROV]),
                   M.DIST: C.s(r["_p"][C.DIST]), M.LEVEL: COMMUNITY},
        })
    return out


# managed-and-answered rows are the universe for every figure here
# Method 3: multi-strength medicines collapse to one molecule per facility,
# available if ANY strength was on the shelf - two strengths of a molecule are
# therapeutically interchangeable.
# DEFINITION (study, 2026-10-05): the denominator is every molecule the
# EML authorises at that level of care. A medicine the facility is
# supposed to manage but does not manage at all counts as UNAVAILABLE,
# not as out of scope. MOL.eml_base() applies that and keeps the real
# management flag on `_managed` / the reason on `_gap`.
FAC_BASE = [r for r in MOL.eml_base(M.MEDROWS)
            if M.yes(r[M.MANAGED]) and not M.blank(r[M.AVAIL])]
for r in FAC_BASE:
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])
    r["_shared"] = str(r[M.MED]).strip() in set(SHARED.values())
CHW_BASE = chw_rows()
BASE = FAC_BASE + CHW_BASE

PRODUCTS = sorted({s(r[M.MED]) for r in BASE},
                  key=lambda m: -sum(1 for r in BASE if s(r[M.MED]) == m))
PROVINCES = sorted({s(r["_p"][M.PROV]) for r in BASE})
DISTRICTS = sorted({s(r["_p"][M.DIST]) for r in BASE})
LEVELS = ([COMMUNITY] if CHW_BASE else []) + \
         FL.ordered(r["_lvl"] for r in FAC_BASE)


def cell(rows):
    """-> available, answered, availability %"""
    n = len(rows)
    a = sum(1 for r in rows if M.yes(r[M.AVAIL]))
    return a, n, (pct(a, n) if n >= MIN_N else None)


def crosstab(colvals, colkey):
    """Product rows x one dimension, each column a (% , n) pair."""
    idx = defaultdict(list)
    for r in BASE:
        idx[(s(r[M.MED]), colkey(r))].append(r)
    rows = []
    for m in PRODUCTS:
        tot = [r for r in BASE if s(r[M.MED]) == m]
        a, n, p = cell(tot)
        row = [m, s(tot[0][M.PROGRAM]), n, a, p]
        for cv in colvals:
            a, n, p = cell(idx.get((m, cv), []))
            row += [p, n]
        rows.append(row)
    # the marginal row
    a, n, p = cell(BASE)
    total = ["ALL PRODUCTS", "", n, a, p]
    for cv in colvals:
        a, n, p = cell([r for r in BASE if colkey(r) == cv])
        total += [p, n]
    return rows, total


def sheet(wb, title, blocks, notes=None, widths=None, freeze="C"):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    if notes:
        for n in notes:
            c = ws.cell(row=r, column=1, value=n)
            c.font = Font(size=9, italic=True, color="555555")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
            ws.row_dimensions[r].height = 13 * (1 + len(n) // 150)
            r += 1
        r += 1
    for btitle, headers, rows, total, pctcols in blocks:
        if btitle:
            ws.cell(row=r, column=1, value=btitle).font = TITLE_FONT
            r += 1
        for ci, h in enumerate(headers, start=1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
            c.alignment = Alignment(wrap_text=True, vertical="center",
                                    horizontal="center")
        ws.row_dimensions[r].height = 42
        hdr = r
        r += 1
        for row in list(rows) + ([total] if total else []):
            is_tot = total is not None and row is total
            for ci, v in enumerate(row, start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.border = BORDER
                c.alignment = Alignment(vertical="top",
                                        wrap_text=isinstance(v, str)
                                        and len(str(v)) > 30)
                if (ci - 1) in pctcols and isinstance(v, float):
                    c.number_format = "0.0%"
                    if v < 0.85:
                        c.fill = BAD_FILL
                    elif v < 0.95:
                        c.fill = MID_FILL
                    else:
                        c.fill = OK_FILL
                elif isinstance(v, (int, float)):
                    c.number_format = "#,##0"
                if is_tot:
                    c.font = Font(bold=True)
                    if not ((ci - 1) in pctcols and isinstance(v, float)):
                        c.fill = TOT_FILL
            r += 1
        ws.freeze_panes = f"{freeze}{hdr + 1}"
        ws.auto_filter.ref = f"A{hdr}:{get_column_letter(len(headers))}{r - 1}"
        r += 2
    for ci, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    return ws


def pair_headers(label, vals):
    h = ["Product", "Programme", "Answered (all)", "Available (all)",
         "AVAILABLE % (all)"]
    for v in vals:
        h += [f"{v}\n%", f"{v}\nn"]
    return h


def build():
    wb = Workbook()
    wb.remove(wb.active)
    npct = tuple([4] + [5 + 2 * i for i in range(40)])

    note_denom = ("Availability = available / (managed AND answered). "
                  f"{len(BASE):,} observations qualify, from "
                  f"{len({s(r['_p'][M.FAC]) for r in BASE})} facilities.")
    note_thin = (f"A percentage is shown only where the cell has at least "
                 f"{MIN_N} answered observations - below that the counts are "
                 f"given but the rate would be noise. Red below 85%, amber "
                 f"85-95%, green 95%+.")

    # ---------- 1. by level of care
    rows, total = crosstab(LEVELS, lambda r: r["_lvl"])
    sheet(wb, "Molecule x level of care",
          [(None, pair_headers("level", LEVELS), rows, total, npct)],
          notes=[note_denom, note_thin,
                 "Columns run from the lowest tier to the highest, so a "
                 "gradient reads left to right.",
                 "The community tier comes from the CHW assessment and is "
                 "measured over the 13-item CHW package, while the facility "
                 "tiers are measured over the 79-item facility basket - so "
                 "these columns are NOT like for like. The 'Community vs "
                 "facility' sheet restricts every tier to the eight shared "
                 "products, which is the comparable version.",
                 "Multi-strength medicines are combined: a molecule counts "
                 "as available where ANY strength was on the shelf, because "
                 "the strengths are interchangeable. 81 items become 71 "
                 "molecules; the nine combined groups are in "
                 "Combined_strength_availability.xlsx.",
                 "Health post = " + FL.MERGED[FL.HEALTH_POST] + ". "
                 "Health centre = " + FL.MERGED[FL.HEALTH_CENTRE] + ". "
                 "Hospitals are reported separately. The ungrouped values are "
                 "on the 'Recorded level values' sheet."],
          widths={1: 46, 2: 18, 3: 12, 4: 12, 5: 14})

    # ---------- 2. by province
    rows, total = crosstab(PROVINCES, lambda r: s(r["_p"][M.PROV]))
    sheet(wb, "Molecule x province",
          [(None, pair_headers("province", PROVINCES), rows, total, npct)],
          notes=[note_denom, note_thin],
          widths={1: 46, 2: 18, 3: 12, 4: 12, 5: 14})

    # ---------- 3. by district, long format (30 districts is too wide)
    rows = []
    idx = defaultdict(list)
    for r in BASE:
        idx[(s(r[M.MED]), s(r["_p"][M.DIST]))].append(r)
    for (m, d), rs in sorted(idx.items()):
        a, n, p = cell(rs)
        rows.append([m, s(rs[0][M.PROGRAM]), s(rs[0]["_p"][M.PROV]), d,
                     len({s(r["_p"][M.FAC]) for r in rs}), n, a, n - a, p])
    sheet(wb, "Molecule x district",
          [(None, ["Product", "Programme", "Province", "District",
                   "Facilities", "Answered", "Available", "Stocked out",
                   "AVAILABLE %"], rows, None, (8,))],
          notes=[note_denom,
                 "Long format, not a grid: 81 products across 30 districts "
                 "would be 2,430 columns-worth of mostly empty cells. Filter "
                 "on Product or District, or pivot it.",
                 f"A percentage appears only where the cell has {MIN_N}+ "
                 f"answered observations. Most product x district cells are "
                 f"smaller than that, which is why the counts matter more "
                 f"than the rate here."],
          widths={1: 46, 2: 18, 3: 13, 4: 15, 5: 10, 6: 10, 7: 10, 8: 11,
                  9: 13}, freeze="E")

    # ---------- 4. the marginals, so the three dimensions can be read alone
    def marg(label, vals, keyfn):
        out = []
        for v in vals:
            rs = [r for r in BASE if keyfn(r) == v]
            a, n, p = cell(rs)
            out.append([v, len({s(r["_p"][M.FAC]) for r in rs}),
                        len({s(r[M.MED]) for r in rs}), n, a, n - a,
                        pct(a, n) if n else None])
        a, n, p = cell(BASE)
        return out, [f"ALL {label}", len({s(r['_p'][M.FAC]) for r in BASE}),
                     len(PRODUCTS), n, a, n - a, pct(a, n)]

    hdr = ["Group", "Facilities", "Products", "Answered", "Available",
           "Stocked out", "AVAILABLE %"]
    l_rows, l_tot = marg("levels", LEVELS, lambda r: r["_lvl"])
    p_rows, p_tot = marg("provinces", PROVINCES, lambda r: s(r["_p"][M.PROV]))
    d_rows, d_tot = marg("districts", DISTRICTS, lambda r: s(r["_p"][M.DIST]))
    d_rows.sort(key=lambda x: (x[6] is None, x[6]))
    sheet(wb, "Totals by dimension",
          [("By level of care (lowest tier first)", hdr, l_rows, l_tot, (6,)),
           ("By province", hdr, p_rows, p_tot, (6,)),
           ("By district, worst availability first", hdr, d_rows, d_tot, (6,))],
          notes=[note_denom,
                 "These are the margins of the three crosstabs - the same "
                 "numbers read one dimension at a time."],
          widths={1: 34, 2: 11, 3: 10, 4: 11, 5: 11, 6: 11, 7: 13}, freeze="B")

    # ---------- 5. level x province, summed over products
    rows = []
    for l in LEVELS:
        row = [l]
        for p in PROVINCES:
            a, n, pc = cell([r for r in BASE if r["_lvl"] == l
                             and s(r["_p"][M.PROV]) == p])
            row += [pc, n]
        a, n, pc = cell([r for r in BASE if r["_lvl"] == l])
        row += [pc, n]
        rows.append(row)
    total = ["ALL LEVELS"]
    for p in PROVINCES:
        a, n, pc = cell([r for r in BASE if s(r["_p"][M.PROV]) == p])
        total += [pc, n]
    a, n, pc = cell(BASE)
    total += [pc, n]
    sheet(wb, "Level x province",
          [(None, ["Level of care"] + [x for p in PROVINCES
                                       for x in (f"{p}\n%", f"{p}\nn")]
            + ["All provinces\n%", "All provinces\nn"], rows, total,
            tuple(range(1, 2 * len(PROVINCES) + 3, 2)))],
          notes=[note_denom,
                 "Summed over all products - this is where the level gradient "
                 "and the geographic pattern meet. An empty cell means no "
                 "facility of that tier was assessed in that province."],
          widths={1: 32}, freeze="B")

    # ---------- 5b. like-for-like: only the products both baskets carry
    shared = [r for r in BASE if r.get("_shared")]
    prods = sorted({s(r[M.MED]) for r in shared})
    rows = []
    for m in prods:
        row = [m]
        for l in LEVELS:
            a, n, p = cell([r for r in shared if s(r[M.MED]) == m
                            and r["_lvl"] == l])
            row += [p, n]
        a, n, p = cell([r for r in shared if s(r[M.MED]) == m])
        row += [p, n]
        rows.append(row)
    total = ["ALL 8 SHARED PRODUCTS"]
    for l in LEVELS:
        a, n, p = cell([r for r in shared if r["_lvl"] == l])
        total += [p, n]
    a, n, p = cell(shared)
    total += [p, n]
    sheet(wb, "Community vs facility",
          [(None, ["Product"] + [x for l in LEVELS
                                 for x in (f"{l}\n%", f"{l}\nn")]
            + ["All tiers\n%", "All tiers\nn"], rows, total,
            tuple(range(1, 2 * len(LEVELS) + 3, 2)))],
          notes=["THE ONLY LIKE-FOR-LIKE COMPARISON IN THIS WORKBOOK. The "
                 "community tier is restricted here to the eight products "
                 "that exist in BOTH baskets, so the same medicines are "
                 "compared at every tier.",
                 "Elsewhere the community tier is measured over the 13-item "
                 "CHW package and the facility tiers over the 79-item "
                 "facility basket - those columns answer 'how is each tier "
                 "doing on its own list', not 'who is better stocked'.",
                 "The CHW basket is not a subset of the facility basket: ORS, "
                 "Zinc 10mg and Amox 125mg dispersible are carried by CHWs "
                 "and by no facility in this study, and ALu 2x6 and 3x6 are "
                 "CHW-only presentations. Those five are excluded here and "
                 "appear on the main sheets."],
          widths={1: 52}, freeze="B")

    # ---------- 6. the recorded level values, for transparency
    rows = []
    for lvl, n in Counter(s(r["_p"][M.LEVEL]) for r in BASE).most_common():
        rs = [r for r in BASE if s(r["_p"][M.LEVEL]) == lvl]
        a, nn, p = cell(rs)
        rows.append([lvl, FL.group(lvl),
                     len({s(r["_p"][M.FAC]) for r in rs}), nn, a,
                     pct(a, nn) if nn else None])
    sheet(wb, "Recorded level values",
          [(None, ["Level exactly as recorded", "Grouped as in this workbook",
                   "Facilities", "Answered", "Available", "AVAILABLE %"],
            rows, None, (5,))],
          notes=["The form records some tiers under more than one label. The "
                 "crosstabs use the grouped column so a tier is not split "
                 "into two thin rows; this sheet shows what was merged.",
                 "SGHP, 'Health Post (HP)' and the bare 'HP' are reported "
                 "together as one Health post tier; MHC is reported with the "
                 "health centres, which is what the form's own label "
                 "('MHC (health-centre level)') says it is.",
                 "Facility_level_flags.xlsx still lists the facilities whose "
                 "recorded tier is ambiguous - grouping hides that question "
                 "rather than answering it."],
          widths={1: 40, 2: 32, 3: 11, 4: 11, 5: 11, 6: 13}, freeze="B")

    wb.active = 0
    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)      # Box resets synced files to read-only
    wb.save(OUT)

    a, n, p = cell(BASE)
    print(f"wrote {OUT}")
    print(f"  {len(PRODUCTS)} products x {len(LEVELS)} levels, "
          f"{len(PROVINCES)} provinces, {len(DISTRICTS)} districts")
    print(f"  base: {n:,} managed-and-answered observations, "
          f"{a:,} available ({p:.1%})")
    sh = [r for r in BASE if r.get("_shared")]
    print(f"  like-for-like on the {len({s(r[M.MED]) for r in sh})} shared "
          f"products:")
    for l in LEVELS:
        a, n, p = cell([r for r in sh if r["_lvl"] == l])
        print(f"    {l:34} {a:>5,}/{n:<6,} "
              + (f"{p:.1%}" if p is not None else f"n<{MIN_N}"))
    print(f"  availability by level on each tier's OWN basket:")
    for l in LEVELS:
        a, n, p = cell([r for r in BASE if r["_lvl"] == l])
        print(f"    {l:34} {a:>5,}/{n:<6,} "
              + (f"{p:.1%}" if p is not None else f"n<{MIN_N}"))


if __name__ == "__main__":
    build()
