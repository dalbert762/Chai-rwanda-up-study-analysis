"""Three indicator sets in one workbook:

  1. Highly prescribed medicines, overall and by level of care
  2. Order fill rate over June-August: requested from RMS against received
  3. Commonly unserved products

    python3 Prescribing_fill_rate_unserved.py
      -> Prescribing_fill_rate_unserved.xlsx

The three come from two different tools and are NOT joinable at medicine level:
the prescription tool names medicines from its own prescribing list
("paracetamol_tab_500_mg") and the stock tool from the procurement catalogue
("PARACETAMOL 500MG TAB"). So "commonly unserved" is reported twice - once as
prescriptions that could not be dispensed, once as stock-outs on the shelf -
side by side rather than merged.

A WARNING THAT AFFECTS INDICATOR 1
19.7% of prescription items (2,751 of 13,946, across 168 distinct codes) carry
an unresolved medicine CODE instead of a name - m1332, m0824, m1179 and so on.
They are not in any choice list in the project folder, so they cannot be named
from available files. The ranked list of named medicines is therefore built on
the 80.3% that resolve, and the codes are ranked separately: m1332 alone would
place about 7th. One lookup file would fix the whole list.

openpyxl + the standard library, like the rest of this project.
"""

import _bootstrap as B   # path setup - must come first
import os
import re
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_kobo_prescription as P
import analyze_medicine_assessment as M
import facility_levels as FL

OUT = B.out("Prescribing_fill_rate_unserved.xlsx")
TOP = 30             # how many medicines to rank
TOP_PER_LEVEL = 10
MIN_REQ = 1000       # units requested below which a fill rate is not shown
CODE = re.compile(r"m\d{3,4}$")

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, size=9, color="FFFFFF")
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")
BAD_FILL = PatternFill("solid", fgColor="FCE4E4")
MID_FILL = PatternFill("solid", fgColor="FFF4D6")
OK_FILL = PatternFill("solid", fgColor="E7F3E7")
TITLE_FONT = Font(bold=True, size=12, color="1F3864")
MONTHS = ("June", "July", "August")


def s(v):
    return "" if v is None else str(v).strip()


def pct(a, b):
    return a / b if b else None


# ---------------------------------------------------------------
# sources
# ---------------------------------------------------------------
ITEMS = [r for r in P.item if r["_sub"]]
for r in ITEMS:
    r["_lvl"] = FL.group(r["_sub"][P.FTYPE])
    r["_code"] = bool(CODE.match(s(r["_med"])))
NAMED = [r for r in ITEMS if not r["_code"]]
CODED = [r for r in ITEMS if r["_code"]]
UNSERVED = [r for r in ITEMS if r["_avail"] is False]

MG = [r for r in M.MEDROWS if M.yes(r[M.MANAGED])]
for r in MG:
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])
LEVELS_P = FL.ordered(r["_lvl"] for r in ITEMS)
LEVELS_M = FL.ordered(r["_lvl"] for r in MG)


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
                                        and len(str(v)) > 32)
                if (ci - 1) in pctcols and isinstance(v, float):
                    c.number_format = "0.0%"
                    if v < 0.85:
                        c.fill = BAD_FILL
                    elif v < 0.95:
                        c.fill = MID_FILL
                    else:
                        c.fill = OK_FILL
                elif (ci - 1) in num0 and isinstance(v, (int, float)):
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


def build():
    wb = Workbook()
    wb.remove(wb.active)

    # ---------- 0. method
    method = [
        ["Indicator 1 - highly prescribed medicines",
         "Prescription assessment, medicine_item rows. Count of times each "
         "medicine was prescribed, with the share of items and of "
         "prescriptions, and a breakdown by level of care.",
         f"{len(ITEMS):,} items, {P.N_PRES:,} prescriptions"],
        ["Indicator 2 - order fill rate",
         "Medicine assessment. Quantity delivered as a share of quantity "
         "requested from RMS, June-August 2026, from SUMMED quantities - "
         "never a mean of per-row percentages.",
         f"{len(MG):,} managed facility x medicine rows"],
        ["Indicator 3 - commonly unserved products",
         "Two different measures, reported separately: prescriptions that "
         "could not be dispensed (prescription tool) and medicines absent "
         "from the shelf on the visit day (medicine assessment).",
         f"{len(UNSERVED)} unserved items; "
         f"{sum(1 for r in MG if M.no(r[M.AVAIL]))} stock-outs"],
        ["Why the three are not joined",
         "The two tools name medicines from different lists - "
         "'paracetamol_tab_500_mg' in the prescription tool against "
         "'PARACETAMOL 500MG TAB' in the stock tool. There is no shared "
         "identifier, so a merged medicine ranking would be wrong.", ""],
        ["Unresolved medicine codes",
         f"{len(CODED):,} of {len(ITEMS):,} prescription items "
         f"({pct(len(CODED), len(ITEMS)):.1%}) carry a code such as m1332 "
         f"instead of a name, across "
         f"{len({s(r['_med']) for r in CODED})} distinct codes. They are not "
         f"in any choice list in the project folder. The named ranking below "
         f"excludes them; they are ranked separately.",
         "ACTION: supply the code-to-name lookup and the ranking is complete"],
        ["Fill rate suppression",
         f"A fill rate is shown only where at least {MIN_REQ:,} units were "
         f"requested. Below that the ratio swings on a handful of units.", ""],
    ]
    sheet(wb, "Method",
          [(None, ["Indicator", "How it is computed", "Scale"], method, None,
            (), ())],
          widths={1: 40, 2: 96, 3: 42})

    # ---------- 1. top prescribed, named
    counts = Counter(s(r["_med"]) for r in NAMED)
    rows = []
    for rank, (m, n) in enumerate(counts.most_common(TOP), 1):
        sub = [r for r in NAMED if s(r["_med"]) == m]
        un = sum(1 for r in sub if r["_avail"] is False)
        pres = len({id(r["_pres"]) for r in sub if r["_pres"]})
        row = [rank, m, s(sub[0][P.CLASS])[:40], n, pct(n, len(ITEMS)),
               pct(pres, P.N_PRES),
               len({s(r["_sub"][P.FACIL]) for r in sub}),
               sum(1 for r in sub if r["_abx"]), un,
               pct(un, n) if n else None]
        for l in LEVELS_P:
            row.append(sum(1 for r in sub if r["_lvl"] == l))
        rows.append(row)
    hdr = ["Rank", "Medicine", "Drug class", "Times prescribed",
           "% of all items", "% of prescriptions", "Facilities",
           "Items flagged antibiotic", "Unserved", "Unserved %"] + \
          [f"{l}\n(times)" for l in LEVELS_P]
    sheet(wb, "Top prescribed medicines",
          [(None, hdr, rows, None, (4, 5, 9), tuple([3, 6, 7, 8] +
            list(range(10, 10 + len(LEVELS_P)))))],
          notes=[f"Top {TOP} by number of times prescribed, among the "
                 f"{len(NAMED):,} items whose medicine resolves to a name.",
                 f"The {len(CODED):,} coded items are excluded here - see the "
                 f"'Unresolved codes' sheet. m1332 would place around 7th, so "
                 f"the ordering below is provisional until the codes are "
                 f"named.",
                 "The level columns count times prescribed at each tier, so a "
                 "medicine concentrated at one level shows up immediately."],
          widths={1: 6, 2: 42, 3: 34, 4: 13, 5: 11, 6: 12, 7: 11, 8: 13,
                  9: 10, 10: 11}, freeze="C")

    # ---------- 2. top prescribed within each level
    rows = []
    for l in LEVELS_P:
        sub = [r for r in NAMED if r["_lvl"] == l]
        tot = len([r for r in ITEMS if r["_lvl"] == l])
        for rank, (m, n) in enumerate(
                Counter(s(r["_med"]) for r in sub).most_common(TOP_PER_LEVEL), 1):
            ms = [r for r in sub if s(r["_med"]) == m]
            rows.append([l, rank, m, n, pct(n, tot),
                         len({s(r["_sub"][P.FACIL]) for r in ms}),
                         sum(1 for r in ms if r["_avail"] is False)])
    sheet(wb, "Top prescribed by level",
          [(None, ["Level of care", "Rank", "Medicine", "Times prescribed",
                   "% of that level's items", "Facilities", "Unserved"],
            rows, None, (4,), (3, 5, 6))],
          notes=[f"The top {TOP_PER_LEVEL} medicines within each tier, so the "
                 f"ranking is not dominated by the health centres' volume.",
                 "Percentages are of that level's own items, so they compare "
                 "across tiers of very different size.",
                 "Coded medicines are excluded, as above."],
          widths={1: 28, 2: 6, 3: 42, 4: 13, 5: 14, 6: 11, 7: 10}, freeze="D")

    # ---------- 3. unresolved codes
    rows = []
    for rank, (m, n) in enumerate(
            Counter(s(r["_med"]) for r in CODED).most_common(), 1):
        sub = [r for r in CODED if s(r["_med"]) == m]
        rows.append([rank, m, n, pct(n, len(ITEMS)),
                     len({s(r["_sub"][P.FACIL]) for r in sub}),
                     s(sub[0][P.CLASS])[:50],
                     sum(1 for r in sub if r["_avail"] is False)])
    sheet(wb, "Unresolved codes",
          [(None, ["Rank among codes", "Code as recorded", "Times prescribed",
                   "% of all items", "Facilities",
                   "Drug class as recorded (may identify it)", "Unserved"],
            rows, None, (3,), (2, 4, 6))],
          notes=[f"{len(rows)} codes covering {len(CODED):,} items "
                 f"({pct(len(CODED), len(ITEMS)):.1%} of everything "
                 f"prescribed). None appear in any choice list in the project "
                 f"folder.",
                 "The drug class column is recorded alongside the code and "
                 "often identifies the medicine, so this list is a usable "
                 "starting point for reconstructing the mapping.",
                 "ACTION: supply the medicine list used by the deployed form "
                 "version and the 'Top prescribed' sheet becomes complete."],
          widths={1: 14, 2: 16, 3: 13, 4: 11, 5: 11, 6: 52, 7: 10})

    # ---------- 4. order fill rate, by month
    def fill_block(label, keyfn, keys):
        out = []
        for k in keys:
            sub = [r for r in MG if keyfn(r) == k]
            row = [k, len({s(r["_p"][M.FAC]) for r in sub}), len(sub)]
            for mth in MONTHS:
                req = sum(M.num(r[M.REQ[mth]]) or 0 for r in sub)
                dlv = sum(M.num(r[M.DEL[mth]]) or 0 for r in sub)
                row += [round(req), round(dlv),
                        pct(dlv, req) if req >= MIN_REQ else None]
            req = sum(M.num(r[M.REQ_TOT]) or 0 for r in sub)
            dlv = sum(M.num(r[M.DEL_TOT]) or 0 for r in sub)
            row += [round(req), round(dlv), round(req - dlv),
                    pct(dlv, req) if req >= MIN_REQ else None]
            out.append(row)
        return out

    fhdr = ["Group", "Facilities", "Rows"]
    for mth in MONTHS:
        fhdr += [f"{mth}\nrequested", f"{mth}\nreceived", f"{mth}\nfill rate"]
    fhdr += ["Total requested", "Total received", "Shortfall",
             "FILL RATE Jun-Aug"]
    fpct = (5, 8, 11, 15)
    fnum = (1, 2, 3, 4, 6, 7, 9, 10, 12, 13, 14)
    allrow = fill_block("ALL", lambda r: "ALL", ["ALL"])[0]
    sheet(wb, "Order fill rate",
          [("Overall and by month", fhdr, [allrow], None, fpct, fnum),
           ("By level of care", fhdr,
            fill_block("lvl", lambda r: r["_lvl"], LEVELS_M), allrow, fpct,
            fnum),
           ("By province", fhdr,
            fill_block("prov", lambda r: s(r["_p"][M.PROV]),
                       sorted({s(r["_p"][M.PROV]) for r in MG})), allrow,
            fpct, fnum),
           ("By programme", fhdr,
            fill_block("prog", lambda r: s(r[M.PROGRAM]),
                       sorted({s(r[M.PROGRAM]) for r in MG})), allrow, fpct,
            fnum)],
          notes=["Quantity requested from RMS against quantity received, "
                 "June-August 2026. Computed from summed quantities, not as a "
                 "mean of per-row percentages - that would weight a 30-unit "
                 "line like a 50,000-unit one.",
                 f"A fill rate is blank where fewer than {MIN_REQ:,} units "
                 f"were requested. Rows where nothing was requested "
                 f"contribute to the counts but not to the rate.",
                 "Red below 85%, amber 85-95%, green above."],
          widths={1: 30, 2: 10, 3: 8}, freeze="D")

    # ---------- 5. fill rate by medicine
    rows = []
    for med in sorted({s(r[M.MED]) for r in MG}):
        sub = [r for r in MG if s(r[M.MED]) == med]
        req = sum(M.num(r[M.REQ_TOT]) or 0 for r in sub)
        dlv = sum(M.num(r[M.DEL_TOT]) or 0 for r in sub)
        if req <= 0:
            continue
        rows.append([med, s(sub[0][M.PROGRAM]), len(sub),
                     len({s(r["_p"][M.FAC]) for r in sub}), round(req),
                     round(dlv), round(req - dlv),
                     pct(dlv, req) if req >= MIN_REQ else None,
                     sum(1 for r in sub if M.no(r[M.AVAIL]))])
    rows.sort(key=lambda x: (x[7] is None, x[7] if x[7] is not None else 9))
    sheet(wb, "Fill rate by medicine",
          [(None, ["Medicine", "Programme", "Rows", "Facilities",
                   "Requested", "Received", "Shortfall", "FILL RATE",
                   "Stocked out on visit day"], rows, None, (7,),
            (2, 3, 4, 5, 6, 8))],
          notes=["Worst fill rate first. The last column is from the same "
                 "rows, so a medicine with a low fill rate AND stock-outs is "
                 "failing on both the order and the shelf.",
                 f"Medicines with fewer than {MIN_REQ:,} units requested show "
                 f"their quantities but no rate."],
          widths={1: 48, 2: 18, 3: 8, 4: 10, 5: 13, 6: 13, 7: 12, 8: 11,
                  9: 14}, freeze="C")

    # ---------- 6. commonly unserved, both measures
    rows = []
    for m, n in Counter(s(r["_med"]) for r in UNSERVED).most_common():
        sub = [r for r in UNSERVED if s(r["_med"]) == m]
        times = sum(1 for r in ITEMS if s(r["_med"]) == m)
        rows.append([m, s(sub[0][P.CLASS])[:40], n, times,
                     pct(n, times) if times else None,
                     len({s(r["_sub"][P.FACIL]) for r in sub}),
                     round(sum(P.num(r.get(P.QTY)) or 0 for r in sub)),
                     round(sum(P.num(r.get(P.COST_UNSERVED)) or 0
                               for r in sub)),
                     sum(1 for r in sub
                         if s(r.get(P.ORDERED)).lower() == "yes"),
                     sum(1 for r in sub
                         if s(r.get(P.RMS_REG)).lower() == "yes"
                         or s(r.get(P.RMS_CEN)).lower() == "yes")])
    rows.sort(key=lambda x: (-x[2], -(x[4] or 0)))

    so = [r for r in MG if M.no(r[M.AVAIL])]
    rows2 = []
    for m, n in Counter(s(r[M.MED]) for r in so).most_common():
        sub = [r for r in so if s(r[M.MED]) == m]
        managed = sum(1 for r in MG if s(r[M.MED]) == m)
        rows2.append([m, s(sub[0][M.PROGRAM]), n, managed,
                      pct(n, managed) if managed else None,
                      len({s(r["_p"][M.FAC]) for r in sub}),
                      round(sum(M.num(r[M.SO_TOTAL]) or 0 for r in sub)),
                      round(sum(M.num(r[M.LOSS]) or 0 for r in sub)),
                      sum(1 for r in sub
                          if s(r[M.RMS_CEN]).lower() == "yes"
                          or s(r[M.RMS_REG]).lower() == "yes")])
    rows2.sort(key=lambda x: (-x[2], -(x[4] or 0)))

    sheet(wb, "Commonly unserved",
          [("A. Prescriptions that could not be dispensed (prescription tool)",
            ["Medicine", "Drug class", "Unserved items", "Times prescribed",
             "Unserved %", "Facilities", "Quantity", "Value (RWF)",
             "Of which ordered", "Of which in stock upstream"], rows, None,
            (4,), (2, 3, 5, 6, 7, 8, 9)),
           ("B. Medicines absent from the shelf on the visit day (medicine "
            "assessment)",
            ["Medicine", "Programme", "Stock-outs", "Managed observations",
             "Stock-out %", "Facilities", "Stock-out days Jun-Aug",
             "Recorded loss (RWF)", "Of which in stock upstream"], rows2,
            None, (4,), (2, 3, 5, 6, 7, 8))],
          notes=["TWO DIFFERENT MEASURES, deliberately not merged. A is a "
                 "patient-level event: someone presented a prescription and "
                 "left without the medicine. B is a facility-level state: the "
                 "medicine was not on the shelf when the team visited.",
                 "The two tools name medicines from different lists, so the "
                 "same product can appear in both tables under different "
                 "names. There is no shared identifier to join them on.",
                 "'In stock upstream' means RMS Regional or Central answered "
                 "Yes - a distribution failure rather than a national "
                 "stock-out."],
          widths={1: 48, 2: 34, 3: 12, 4: 13, 5: 11, 6: 11, 7: 12, 8: 13,
                  9: 13, 10: 16}, freeze="C")

    wb.active = 0
    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)      # Box resets synced files to read-only
    wb.save(OUT)

    req = sum(M.num(r[M.REQ_TOT]) or 0 for r in MG)
    dlv = sum(M.num(r[M.DEL_TOT]) or 0 for r in MG)
    print(f"wrote {OUT}")
    print(f"  1. prescribed: {len(ITEMS):,} items, "
          f"{len({s(r['_med']) for r in NAMED})} named medicines "
          f"(+{len({s(r['_med']) for r in CODED})} unresolved codes covering "
          f"{pct(len(CODED), len(ITEMS)):.1%} of items)")
    print(f"     top 5: " + ", ".join(
        f"{m} ({n:,})" for m, n in
        Counter(s(r['_med']) for r in NAMED).most_common(5)))
    print(f"  2. order fill rate Jun-Aug: {dlv:,.0f} received of "
          f"{req:,.0f} requested = {pct(dlv, req):.1%}")
    for mth in MONTHS:
        rq = sum(M.num(r[M.REQ[mth]]) or 0 for r in MG)
        dv = sum(M.num(r[M.DEL[mth]]) or 0 for r in MG)
        print(f"     {mth:8} {dv:>12,.0f} / {rq:>12,.0f} = {pct(dv, rq):.1%}")
    print(f"  3. unserved: {len(UNSERVED)} prescription items across "
          f"{len({s(r['_med']) for r in UNSERVED})} medicines; "
          f"{len(so)} shelf stock-outs across "
          f"{len({s(r[M.MED]) for r in so})} medicines")


if __name__ == "__main__":
    build()
