"""Unserved prescriptions by geography and by level of care.

    python3 Unserved_by_geography.py  ->  Unserved_by_geography.xlsx

Covers plan indicators 4.1, 4.2, 5.1, 6.1, 6.2 and 7.1, cut by level of care,
province, district and facility.

READ THIS BEFORE QUOTING ANY RATE HERE
The unserved rate is a DESIGN ARTEFACT, not a measure of facility
availability. A prescriber writes what the pharmacy holds, so a medicine the
facility does not stock is rarely prescribed and never enters the denominator;
the form also cannot record "available, but not enough of it". The rate
therefore sits near 0.5%, against 6.5% stocked out on the medicine assessment
for the same facilities in the same weeks. For availability, use the medicine
assessment.

What this file IS good for: the 69 unserved items are real events, each with a
medicine, a quantity, a price, whether it was ordered, whether it was supplied
and whether it was in stock upstream. That is objectives 5, 6 and 7, and those
are worth reporting facility by facility.

COUNTS BEFORE RATES
69 items across 265 facilities means almost every cell is a single digit. Every
table leads with the count, gives the denominator, and shows a percentage only
where the cell has at least MIN_N answered items. Districts with zero unserved
items are listed rather than omitted - "no unserved items recorded" is a
finding, not an empty row.

openpyxl + the standard library, like the rest of this project.
"""

import _bootstrap as B   # path setup - must come first
import os
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_kobo_prescription as P
import facility_levels as FL

OUT = B.out("Unserved_by_geography.xlsx")
MIN_N = 200        # answered items below which a rate is not shown

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, size=9, color="FFFFFF")
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")
BAD_FILL = PatternFill("solid", fgColor="FCE4E4")
MID_FILL = PatternFill("solid", fgColor="FFF4D6")
OK_FILL = PatternFill("solid", fgColor="E7F3E7")
TITLE_FONT = Font(bold=True, size=12, color="1F3864")


def s(v):
    return "" if v is None else str(v).strip()


def pct(a, b):
    return a / b if b else None


def lvl(r):
    return FL.group(r["_sub"][P.FTYPE])


ITEMS = [r for r in P.item if r["_sub"] and r["_avail"] is not None]
UNSERVED = [r for r in ITEMS if r["_avail"] is False]
PRES = [r for r in P.pres if r["_sub"]]


def measure(label, items, pres):
    un = [r for r in items if r["_avail"] is False]
    qty = sum(P.num(r.get(P.QTY)) or 0 for r in un)
    val = sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in un)
    aff = [r for r in pres if r["_any_unserved"]]
    return [label,
            len({s(r["_sub"][P.FACIL]) for r in items}),
            len({s(r["_sub"][P.FACIL]) for r in un}),
            len(pres), len(aff), pct(len(aff), len(pres)) if pres else None,
            len(items), len(un),
            pct(len(un), len(items)) if len(items) >= MIN_N else None,
            round(qty), round(val),
            len({s(r["_med"]) for r in un})]


HDR = ["Group", "Facilities assessed", "Facilities with an unserved item",
       "Prescriptions", "Prescriptions with an unserved item",
       "% of prescriptions", "Medicine items answered", "UNSERVED ITEMS",
       "Unserved %", "Quantity unserved", "Value (RWF)",
       "Distinct medicines involved"]
PCTCOLS = (5, 8)
NUM0 = (1, 2, 3, 4, 6, 7, 9, 10, 11)


def by(keyfn, keys=None, sort_count=True):
    gi, gp = defaultdict(list), defaultdict(list)
    for r in ITEMS:
        gi[keyfn(r)].append(r)
    for r in PRES:
        gp[keyfn(r)].append(r)
    ks = keys if keys is not None else sorted(set(gi) | set(gp))
    rows = [measure(k or "(not recorded)", gi.get(k, []), gp.get(k, []))
            for k in ks]
    if sort_count:
        rows.sort(key=lambda x: (-x[7], -x[6]))
    return rows


ALL = measure("ALL", ITEMS, PRES)


def sheet(wb, title, blocks, notes=None, widths=None, freeze="B"):
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
    for btitle, headers, rows, total, pctcols, num0 in blocks:
        if btitle:
            ws.cell(row=r, column=1, value=btitle).font = TITLE_FONT
            r += 1
        for ci, h in enumerate(headers, start=1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
            c.alignment = Alignment(wrap_text=True, vertical="center",
                                    horizontal="center")
        ws.row_dimensions[r].height = 44
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
                    c.number_format = "0.00%"
                    if v >= 0.02:
                        c.fill = BAD_FILL
                    elif v >= 0.01:
                        c.fill = MID_FILL
                elif (ci - 1) in num0 and isinstance(v, (int, float)):
                    c.number_format = "#,##0"
                if is_tot:
                    c.font = Font(bold=True)
                    c.fill = TOT_FILL
            r += 1
        ws.freeze_panes = f"{freeze}{hdr + 1}"
        ws.auto_filter.ref = f"A{hdr}:{get_column_letter(len(headers))}{r - 1}"
        r += 2
    for ci, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    return ws


NOTE_ARTEFACT = ("The unserved rate is a design artefact, not availability: a "
                 "prescriber writes what the pharmacy holds, so an unstocked "
                 "medicine never enters the denominator. It sits near 0.5% "
                 "here against 6.5% stocked out on the medicine assessment "
                 "for the same facilities. Quote the medicine assessment for "
                 "availability; use this file for WHICH items went unserved, "
                 "where, and what happened next.")
NOTE_THIN = (f"A percentage is shown only where the group has {MIN_N}+ "
             f"answered items. With {len(UNSERVED)} unserved items in total, "
             f"most cells are single digits - read the counts.")


def build():
    wb = Workbook()
    wb.remove(wb.active)
    levels = FL.ordered(lvl(r) for r in ITEMS)
    provs = sorted({s(r["_sub"][P.PROV]) for r in ITEMS})
    dists = sorted({s(r["_sub"][P.DISTR]) for r in ITEMS})

    # ---------- 1. overview
    conc = [
        ["Medicine items answered", ALL[6], ""],
        ["UNSERVED items", ALL[7],
         f"{pct(ALL[7], ALL[6]):.2%} of answered items"],
        ["Prescriptions", ALL[3], ""],
        ["Prescriptions with at least one unserved item", ALL[4],
         f"{pct(ALL[4], ALL[3]):.2%} of prescriptions"],
        ["Facilities assessed", ALL[1], ""],
        ["Facilities with at least one unserved item", ALL[2],
         f"{pct(ALL[2], ALL[1]):.1%} of facilities - the rest recorded none"],
        ["Districts assessed", len(dists), ""],
        ["Districts with at least one unserved item",
         len({s(r["_sub"][P.DISTR]) for r in UNSERVED}),
         f"{len(dists) - len({s(r['_sub'][P.DISTR]) for r in UNSERVED})} "
         f"districts recorded none"],
        ["Quantity of medicine unserved", ALL[9], "units"],
        ["Recorded value of the unserved items", ALL[10], "RWF, at tariff price"],
        ["Distinct medicines involved", ALL[11], ""],
    ]
    ordered_out = sorted(Counter(s(r["_sub"][P.FACIL]) for r in UNSERVED
                                 ).items(), key=lambda x: -x[1])
    sheet(wb, "Overview",
          [("Headline", ["Measure", "Value", "Note"], conc, None, (), (1,)),
           ("Where the 69 unserved items sit",
            ["Facility", "District", "Level of care", "Unserved items",
             "Items answered at that facility", "Unserved %"],
            [[f, s(next(r["_sub"][P.DISTR] for r in UNSERVED
                        if s(r["_sub"][P.FACIL]) == f)),
              lvl(next(r for r in UNSERVED if s(r["_sub"][P.FACIL]) == f)),
              n, sum(1 for r in ITEMS if s(r["_sub"][P.FACIL]) == f),
              pct(n, sum(1 for r in ITEMS if s(r["_sub"][P.FACIL]) == f))]
             for f, n in ordered_out], None, (5,), (3, 4))],
          notes=[NOTE_ARTEFACT,
                 f"The {len(UNSERVED)} unserved items are concentrated in "
                 f"{ALL[2]} of {ALL[1]} facilities. That concentration is the "
                 f"most useful thing in this file - it is a short, specific "
                 f"list to follow up rather than a national rate."],
          widths={1: 44, 2: 16, 3: 28, 4: 14, 5: 16, 6: 12})

    # ---------- 2. by level of care
    sheet(wb, "By level of care",
          [(None, HDR, by(lvl, keys=levels, sort_count=False), ALL, PCTCOLS,
            NUM0)],
          notes=[NOTE_ARTEFACT, NOTE_THIN,
                 "Rows run from the lowest tier to the highest. The "
                 "prescription tool has no community tier - CHWs do not write "
                 "prescriptions - so this is facility levels only.",
                 "Health post = " + FL.MERGED[FL.HEALTH_POST] + "; Health "
                 "centre = " + FL.MERGED[FL.HEALTH_CENTRE] + "."],
          widths={1: 30, 2: 12, 3: 14, 4: 12, 5: 14, 6: 12, 7: 13, 8: 12,
                  9: 11, 10: 13, 11: 13, 12: 13})

    # ---------- 3. by province and district
    sheet(wb, "By province and district",
          [("By province", HDR, by(lambda r: s(r["_sub"][P.PROV]), keys=provs),
            ALL, PCTCOLS, NUM0),
           ("By district - every district, including those with none", HDR,
            by(lambda r: s(r["_sub"][P.DISTR]), keys=dists), ALL, PCTCOLS,
            NUM0)],
          notes=[NOTE_ARTEFACT, NOTE_THIN,
                 "Districts with no unserved items are kept in the table. "
                 "Zero is a result - and with only 69 items nationally, most "
                 "districts have it."],
          widths={1: 20, 2: 12, 3: 14, 4: 12, 5: 14, 6: 12, 7: 13, 8: 12,
                  9: 11, 10: 13, 11: 13, 12: 13})

    # ---------- 4. every unserved item, with what happened next
    rows = []
    for r in UNSERVED:
        sub = r["_sub"]
        rows.append([
            s(sub[P.PROV]), s(sub[P.DISTR]), lvl(r), s(sub[P.FACIL]),
            s(r["_med"]), s(r[P.CLASS])[:40],
            round(P.num(r.get(P.QTY)) or 0),
            round(P.num(r.get(P.UNIT_PRICE)) or 0),
            round(P.num(r.get(P.COST_UNSERVED)) or 0),
            s(r.get(P.ORDERED)) or "(not answered)",
            s(r.get(P.SUPPLIED)) or "(not asked - never ordered)",
            s(r.get(P.RMS_REG)) or "-", s(r.get(P.RMS_CEN)) or "-",
            s(r.get(P.DIST_NAT)) or "-",
            s(r["_pres"][P.DATE])[:10] if r["_pres"] else "",
        ])
    rows.sort(key=lambda x: (-x[8], x[0], x[1]))
    sheet(wb, "Every unserved item",
          [(None, ["Province", "District", "Level of care", "Facility",
                   "Medicine", "Drug class", "Qty prescribed",
                   "Unit price (RWF)", "Value (RWF)", "Ordered that month?",
                   "Supplied fully / partially / not at all",
                   "At RMS Regional?", "At RMS Central?", "Attribution",
                   "Prescription date"], rows, None, (), (6, 7, 8))],
          notes=["All " + str(len(UNSERVED)) + " of them, sorted by value. "
                 "This is objectives 5, 6 and 7 at the row level.",
                 "'Supplied' is only asked where the medicine was ordered, so "
                 "'(not asked - never ordered)' is a finding, not missing "
                 "data: the facility never requested it.",
                 "Attribution: available at either warehouse means a "
                 "distribution problem; available at neither means a national "
                 "stock-out."],
          widths={1: 13, 2: 14, 3: 24, 4: 28, 5: 40, 6: 34, 7: 11, 8: 12,
                  9: 12, 10: 15, 11: 24, 12: 13, 13: 13, 14: 22, 15: 14},
          freeze="E")

    # ---------- 5. by medicine, and what happened next
    meds = sorted({s(r["_med"]) for r in UNSERVED})
    rows = []
    for m in meds:
        un = [r for r in UNSERVED if s(r["_med"]) == m]
        rows.append([m, s(un[0][P.CLASS])[:40], len(un),
                     sum(1 for r in ITEMS if s(r["_med"]) == m),
                     len({s(r["_sub"][P.FACIL]) for r in un}),
                     round(sum(P.num(r.get(P.QTY)) or 0 for r in un)),
                     round(sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in un)),
                     sum(1 for r in un
                         if s(r.get(P.ORDERED)).lower() == "yes"),
                     sum(1 for r in un
                         if s(r.get(P.RMS_REG)).lower() == "yes"
                         or s(r.get(P.RMS_CEN)).lower() == "yes")])
    rows.sort(key=lambda x: (-x[2], -x[6]))
    sheet(wb, "By medicine",
          [(None, ["Medicine", "Drug class", "Unserved items",
                   "Times prescribed", "Facilities", "Quantity unserved",
                   "Value (RWF)", "Of which were ordered",
                   "Of which were in stock upstream"], rows, None, (),
            (2, 3, 4, 5, 6, 7, 8))],
          notes=["'In stock upstream' means RMS Regional or Central answered "
                 "Yes - those are distribution failures rather than national "
                 "stock-outs.",
                 "A medicine prescribed often but unserved rarely is working; "
                 "the ones to look at are those unserved on a large share of "
                 "the times they were prescribed."],
          widths={1: 44, 2: 34, 3: 12, 4: 12, 5: 11, 6: 13, 7: 13, 8: 14,
                  9: 16})

    wb.active = 0
    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)      # Box resets synced files to read-only
    wb.save(OUT)

    print(f"wrote {OUT}")
    print(f"  {ALL[7]} unserved items of {ALL[6]:,} answered "
          f"({pct(ALL[7], ALL[6]):.2%}) | {ALL[4]} of {ALL[3]:,} prescriptions "
          f"affected")
    print(f"  concentrated in {ALL[2]} of {ALL[1]} facilities and "
          f"{len({s(r['_sub'][P.DISTR]) for r in UNSERVED})} of {len(dists)} "
          f"districts")
    print(f"  {ALL[9]:,} units, {ALL[10]:,} RWF, {ALL[11]} distinct medicines")
    print("  by level of care:")
    for row in by(lvl, keys=levels, sort_count=False):
        rate = f"{row[8]:.2%}" if row[8] is not None else f"n<{MIN_N}"
        print(f"    {row[0]:30} {row[7]:>3} unserved of {row[6]:>6,} "
              f"answered  {rate}")


if __name__ == "__main__":
    build()
