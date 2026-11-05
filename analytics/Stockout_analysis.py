"""Stock-out analysis across every medicine assessed.

    python3 Stockout_analysis.py  ->  Stockout_analysis.xlsx

The indicator dashboard reports stock-outs nationally and for the fifteen
worst medicines. This covers ALL 79 medicines in the basket, each against its
own denominator, and crosses them with level of care, province, district,
management type, programme and drug class.

THE WINDOW TOTAL IS COMPUTED HERE, NOT READ FROM THE FORM
The instrument carries `total_stockout_days`, and it is

    June + July + August + CURRENT stock-out days (this visit)

The current count is an ongoing run that reaches back before June, so adding
it double-counts the window and pushes 355 rows past the 92 days the window
actually holds - the maximum in the file is 684. Every figure here therefore
sums the three monthly fields, which are individually sound (no value exceeds
the days in its own month). The form's field is reported once, on the data
quality sheet, so the difference is visible rather than silently dropped.

DENOMINATORS
Stock-out questions are asked only where the facility MANAGES the medicine,
so every rate here is over managed facility x medicine pairs. A medicine a
facility does not stock cannot run out of it, and counting that as a
stock-out would turn a scope decision into a supply failure.

Three different measures are kept apart, because they rank medicines
differently and answer different questions:

    INCIDENCE  did this pair run out at all in the window?        (how common)
    DAYS       how many days was it out, summed over the window?  (how long)
    CURRENT    was it out on the day of the visit?                (how now)

Item level, not molecule level: a stock-out is a property of the product on
the shelf, and two strengths of one molecule run out separately.
"""

import _bootstrap as B   # path setup - must come first
import statistics
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_medicine_assessment as M
import explorer as EX
import facility_levels as FL

OUT = B.out("Stockout_analysis.xlsx")
MIN_N = 10          # below this a cell shows counts only, no percentage
MONTHS = ("June", "July", "August")
DAYS = {"June": 30, "July": 31, "August": 31}
WINDOW = sum(DAYS.values())          # 92

HDR_FONT = Font(bold=True, size=9, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="2A68BD")
TITLE_FONT = Font(bold=True, size=11)
TOT_FILL = PatternFill("solid", fgColor="EAEAEA")
BAD_FILL = PatternFill("solid", fgColor="F6C9C4")
MID_FILL = PatternFill("solid", fgColor="FBE3BA")
OK_FILL = PatternFill("solid", fgColor="CDE8DA")
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def s(v):
    return "" if v is None else str(v).strip()


def pct(a, b):
    return a / b if b else None


def med(v):
    v = [x for x in v if x is not None]
    return round(statistics.median(v), 1) if v else None


# ---------------------------------------------------------------
# rows
# ---------------------------------------------------------------
MG = [r for r in M.MEDROWS if M.yes(r[M.MANAGED])]
for r in MG:
    d = [M.num(r[M.SO_MONTH[m]]) for m in MONTHS]
    r["_sod"] = [x for x in d]
    r["_days"] = sum(x or 0 for x in d)
    r["_complete"] = all(x is not None for x in d)
    r["_months"] = sum(1 for x in d if (x or 0) > 0)
    r["_any"] = r["_days"] > 0
    r["_now"] = M.yes(r[M.AVAIL]) is False
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])
    r["_mgt"] = s(r["_p"].get(EX.MGMT)) or "(not recorded)"

COMPLETE = [r for r in MG if r["_complete"]]
MEDS = sorted({s(r[M.MED]) for r in MG})


def block(rows, label):
    """One summary line for any group of managed rows."""
    if not rows:
        return None
    comp = [r for r in rows if r["_complete"]]
    days = sum(r["_days"] for r in rows)
    return [label,
            len({s(r["_p"][M.FAC]) for r in rows}),
            len(rows),
            sum(1 for r in rows if r["_any"]),
            pct(sum(1 for r in rows if r["_any"]), len(rows))
            if len(rows) >= MIN_N else None,
            days,
            med([r["_days"] for r in rows]),
            max((r["_days"] for r in rows), default=0),
            pct(days, len(rows) * WINDOW) if rows else None,
            sum(1 for r in rows if r["_now"]),
            pct(sum(1 for r in rows if r["_now"]), len(rows))
            if len(rows) >= MIN_N else None]


HDR = ["Group", "Facilities", "Managed pairs", "Pairs ever out",
       "INCIDENCE", "Total days out", "Median days", "Worst pair (days)",
       "DAYS OUT RATE", "Out on visit day", "CURRENTLY OUT"]
PCTCOLS = {4, 8, 10}


def by(keyfn, keys=None, rows=None, sort=True):
    g = defaultdict(list)
    for r in (rows if rows is not None else MG):
        g[keyfn(r)].append(r)
    out = [block(g[k], k) for k in (keys or sorted(g)) if g.get(k)]
    if sort:
        out.sort(key=lambda x: -(x[4] or 0))
    return out


# ---------------------------------------------------------------
# workbook
# ---------------------------------------------------------------
def sheet(wb, title, blocks, notes=None, widths=None, freeze="B"):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    if notes:
        for n in notes:
            c = ws.cell(row=r, column=1, value=n)
            c.font = Font(size=9, italic=True, color="555555")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r,
                           end_column=11)
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
        ws.row_dimensions[r].height = 40
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
                    # high is BAD here - these are stock-out measures
                    if v >= 0.20:
                        c.fill = BAD_FILL
                    elif v >= 0.05:
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


def build():
    wb = Workbook()
    wb.remove(wb.active)
    nat = block(MG, "ALL MEDICINES")
    wide = {1: 46, 2: 11, 3: 13, 4: 12, 5: 11, 6: 13, 7: 11, 8: 14, 9: 12,
            10: 12, 11: 12}

    # --- 1. method
    form_tot = sum(M.num(r[M.SO_TOTAL]) or 0 for r in MG)
    over = sum(1 for r in MG if (M.num(r[M.SO_TOTAL]) or 0) > WINDOW)
    sheet(wb, "Method", [
        ("What is counted",
         ["Item", "Value"],
         [["Managed facility x medicine pairs", len(MG)],
          ["Of those, all three months recorded", len(COMPLETE)],
          ["Medicines in the basket", len(MEDS)],
          ["Facilities", len({s(r["_p"][M.FAC]) for r in MG})],
          ["Window", "June-August 2026"],
          ["Days in the window", WINDOW],
          ["Observation-days (pairs x window)", len(MG) * WINDOW],
          ["", ""],
          ["Total stock-out days, SUMMED FROM THE MONTHS (used here)",
           sum(r["_days"] for r in MG)],
          ["Total from the form's total_stockout_days field", round(form_tot)],
          ["Rows where the form field exceeds the 92-day window", over],
          ["Reason", "the form field adds CURRENT stock-out days, which run "
                     "back before June, to the three months - so it "
                     "double-counts. The monthly fields are individually "
                     "sound; none exceeds the days in its own month."]],
         None, set()),
    ], notes=[
        "Stock-out analysis across every medicine assessed. All rates are "
        "over MANAGED facility x medicine pairs - a facility cannot run out "
        "of a medicine it does not stock.",
        "INCIDENCE = share of pairs out at least one day in the window. "
        "DAYS OUT RATE = total days out / (pairs x 92 days). CURRENTLY OUT = "
        "share not on the shelf on the visit day.",
        f"Cells on fewer than {MIN_N} pairs show counts only, no percentage.",
    ], widths={1: 58, 2: 20})

    # --- 2. every medicine
    rows = by(lambda r: s(r[M.MED]), rows=MG)
    sheet(wb, "All medicines", [
        ("Every medicine assessed, worst incidence first", HDR, rows, nat,
         PCTCOLS)],
        notes=[f"All {len(MEDS)} medicines, each against its own denominator.",
               "Sorted by INCIDENCE. Sorting by DAYS OUT RATE or by CURRENTLY "
               "OUT gives a different order - a medicine that is out briefly "
               "but often is not the same problem as one that is out for "
               "months, and the columns are kept apart so both are visible."],
        widths=wide)

    # --- 3. the three measures ranked side by side
    top = sorted(rows, key=lambda x: -(x[4] or 0))[:25]
    byd = sorted(rows, key=lambda x: -(x[8] or 0))[:25]
    byn = sorted(rows, key=lambda x: -(x[10] or 0))[:25]
    sheet(wb, "Worst 25 by each measure", [
        ("By INCIDENCE - ran out at all", HDR, top, None, PCTCOLS),
        ("By DAYS OUT RATE - how long", HDR, byd, None, PCTCOLS),
        ("By CURRENTLY OUT - on the visit day", HDR, byn, None, PCTCOLS)],
        notes=["The same medicines ranked three ways. A medicine high on all "
               "three is a chronic failure; one high only on CURRENTLY OUT "
               "may be a stock-out that started recently."],
        widths=wide)

    # --- 4. geography and facility type
    sheet(wb, "By level and management", [
        ("By level of care", HDR,
         by(lambda r: r["_lvl"], keys=FL.ordered({r["_lvl"] for r in MG}),
            sort=False), nat, PCTCOLS),
        ("By management type", HDR, by(lambda r: r["_mgt"]), nat, PCTCOLS),
        ("By programme", HDR, by(lambda r: s(r[M.PROGRAM]) or "(none)"), nat,
         PCTCOLS),
        ("By drug class", HDR,
         by(lambda r: s(r[M.DCLASS]) or "(unclassified)"), nat, PCTCOLS)],
        widths=wide)

    sheet(wb, "By geography", [
        ("By province", HDR, by(lambda r: s(r["_p"][M.PROV])), nat, PCTCOLS),
        ("By district", HDR, by(lambda r: s(r["_p"][M.DIST])), nat, PCTCOLS)],
        widths=wide)

    # --- 5. medicine x level, medicine x province
    def cross(colkey, colvals, title):
        hdr = ["Medicine", "Managed pairs", "INCIDENCE (all)"]
        for v in colvals:
            hdr += [f"{v}\npairs", f"{v}\nincidence"]
        out = []
        for m in MEDS:
            rs = [r for r in MG if s(r[M.MED]) == m]
            line = [m, len(rs), pct(sum(1 for r in rs if r["_any"]), len(rs))]
            for v in colvals:
                g = [r for r in rs if colkey(r) == v]
                line += [len(g),
                         pct(sum(1 for r in g if r["_any"]), len(g))
                         if len(g) >= MIN_N else None]
            out.append(line)
        out.sort(key=lambda x: -(x[2] or 0))
        pc = {2} | {3 + 2 * i + 1 for i in range(len(colvals))}
        return (title, hdr, out, None, pc)

    LVLS = FL.ordered({r["_lvl"] for r in MG})
    PROVS = sorted({s(r["_p"][M.PROV]) for r in MG})
    sheet(wb, "Medicine x level",
          [cross(lambda r: r["_lvl"], LVLS,
                 "Stock-out incidence, every medicine by level of care")],
          notes=["Blank where that medicine x level cell has fewer than "
                 f"{MIN_N} managed pairs."],
          widths={1: 46}, freeze="D")
    sheet(wb, "Medicine x province",
          [cross(lambda r: s(r["_p"][M.PROV]), PROVS,
                 "Stock-out incidence, every medicine by province")],
          notes=["Blank where that medicine x province cell has fewer than "
                 f"{MIN_N} managed pairs."],
          widths={1: 46}, freeze="D")

    # --- 6. duration profile
    dist = defaultdict(int)
    for r in COMPLETE:
        dist[r["_months"]] += 1
    bands = [("Never out", lambda d: d == 0),
             ("1-7 days", lambda d: 0 < d <= 7),
             ("8-30 days", lambda d: 7 < d <= 30),
             ("31-60 days", lambda d: 30 < d <= 60),
             ("61-91 days", lambda d: 60 < d < WINDOW),
             ("Out the whole window (92)", lambda d: d >= WINDOW)]
    sheet(wb, "Duration profile", [
        ("Months affected, of the three",
         ["Months with at least one stock-out day", "Pairs", "Share"],
         [[k, dist.get(k, 0), pct(dist.get(k, 0), len(COMPLETE))]
          for k in (0, 1, 2, 3)],
         ["TOTAL", len(COMPLETE), 1.0], {2}),
        ("Total days out, banded",
         ["Days out across June-August", "Pairs", "Share"],
         [[lab, sum(1 for r in MG if f(r["_days"])),
           pct(sum(1 for r in MG if f(r["_days"])), len(MG))]
          for lab, f in bands],
         ["TOTAL", len(MG), 1.0], {2}),
        ("Days out by month",
         ["Month", "Days in month", "Pairs with any day out",
          "Share of pairs", "Total days out", "Mean days per pair"],
         [[m, DAYS[m],
           sum(1 for r in MG if (M.num(r[M.SO_MONTH[m]]) or 0) > 0),
           pct(sum(1 for r in MG if (M.num(r[M.SO_MONTH[m]]) or 0) > 0),
               len(MG)),
           round(sum(M.num(r[M.SO_MONTH[m]]) or 0 for r in MG)),
           round(sum(M.num(r[M.SO_MONTH[m]]) or 0 for r in MG) / len(MG), 2)]
          for m in MONTHS], None, {3}),
    ], notes=["The months-affected table uses only pairs with all three "
              f"months recorded ({len(COMPLETE):,} of {len(MG):,})."],
        widths={1: 34, 2: 14, 3: 22, 4: 14, 5: 16, 6: 18})

    # --- 7. facilities
    frows = by(lambda r: s(r["_p"][M.FAC]))
    sheet(wb, "Facilities", [
        ("Every facility, worst incidence first", HDR, frows, nat, PCTCOLS)],
        notes=["One line per facility over its own basket."],
        widths=wide)

    # --- 8. currently out
    now = [r for r in MG if r["_now"]]
    sheet(wb, "Out on the visit day", [
        ("Facility x medicine pairs not on the shelf when the team visited",
         ["Facility", "District", "Province", "Level of care", "Medicine",
          "Programme", "Days out in the window", "Months affected",
          "Current run (days, form field)"],
         sorted([[s(r["_p"][M.FAC]), s(r["_p"][M.DIST]), s(r["_p"][M.PROV]),
                  r["_lvl"], s(r[M.MED]), s(r[M.PROGRAM]), r["_days"],
                  r["_months"], M.num(r[M.SO_NOW])] for r in now],
                key=lambda x: -(x[6] or 0)),
         None, set())],
        notes=[f"{len(now):,} pairs were stocked out on the day of the visit.",
               "The last column is the form's own current-run field, which "
               "counts back beyond June - that is why it can exceed the "
               "window total beside it."],
        widths={1: 32, 2: 16, 3: 14, 4: 24, 5: 46, 6: 18, 7: 14, 8: 12, 9: 18})

    wb.save(OUT)
    return nat, rows


if __name__ == "__main__":
    nat, rows = build()
    print(f"wrote {OUT}")
    print(f"  {len(MEDS)} medicines, {len(MG):,} managed pairs, "
          f"{sum(r['_days'] for r in MG):,} stock-out days")
    print(f"  incidence {nat[4]:.1%} | days out rate {nat[8]:.2%} | "
          f"currently out {nat[10]:.1%}")
    print("  worst five by incidence:")
    for r in rows[:5]:
        print(f"     {r[0][:46]:<48}{r[4]:>7.1%}  {r[5]:>6,} days")
