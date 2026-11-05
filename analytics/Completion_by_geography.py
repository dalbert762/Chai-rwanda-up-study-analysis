"""Medicine assessment vs Prescription assessment: completion by geography.

    python3 Completion_by_geography.py   ->  Completion_by_geography.xlsx

Answers two different questions at four levels (province, district, sector,
facility):

  1. COMPLETION against the plan - how many of the facilities
     Sites_to_be_visited_combined_1.xlsx lists for that area has each tool
     reached? The plan has no sector column, so this is only possible at
     province, district and facility level.
  2. AGREEMENT between the two tools - where has one been and not the other?
     That is the actionable column: a facility with a stock assessment but no
     prescription assessment is a return trip, not a finding.

Facility and plan matching reuses Site_coverage_check.py, so the (core name,
level) rules and the user-confirmed equivalences apply here too - "Kagugu HC"
in the plan is "Kagugu CS" in the tools.

Sector needs its own normalising: the prescription tool writes
"Masaka__Kicukiro" and the stock tool writes "Masaka", and either may carry a
"(Burera District)" qualifier.

openpyxl + the standard library, like the rest of this project.
"""

import _bootstrap as B   # path setup - must come first
import os
import re
import sys
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, ".")
import Site_coverage_check as S   # key(), load_plan(), CONFIRMED

OUT = B.out("Completion_by_geography.xlsx")
TOOLS = ("Medicine assessment", "Prescription")

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="DDE5F0")
GAP_FILL = PatternFill("solid", fgColor="FCE4E4")     # one tool missing
BOTH_FILL = PatternFill("solid", fgColor="E7F3E7")    # both present
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")


# The plan and the tools name provinces differently; districts agree.
PROVINCE = {"city of kigali": "Kigali City", "kigali city": "Kigali City",
            "eastern": "East", "east": "East", "northern": "North",
            "north": "North", "southern": "South", "south": "South",
            "western": "West", "west": "West"}


def prov(v):
    s = re.sub(r"\s+", " ", str(v or "")).strip()
    return PROVINCE.get(s.lower(), s)


def sector(v):
    """'Masaka__Kicukiro' and 'Masaka (Kicukiro District)' -> 'masaka'."""
    s = str(v or "").split("__")[0]
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"\b(district|sector)\b", " ", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip().lower()


def fkey(name, district):
    """Facility key: (core name, level, district).

    The prescription tool's "__District" disambiguator is removed first - it
    writes "Gabiro HP__Nyamasheke" where the stock tool writes "Gabiro HP",
    and those are one facility. The district comes from the record instead,
    which also keeps the three separate RUBONA CS apart.
    """
    return S.dkey(re.sub(r"__\w+", " ", str(name or "")), district)


def clean(v):
    return re.sub(r"\s+", " ", str(v or "")).strip()


def collect():
    """One record per facility per tool: where it is and what its key is."""
    import analyze_kobo_prescription as P
    import analyze_medicine_assessment as M
    out = {}
    out["Prescription"] = {
        fkey(r[P.FACIL], r[P.DISTR]): (clean(r[P.FACIL]), prov(r[P.PROV]),
                                 clean(r[P.DISTR]), sector(r.get("Sector")))
        for r in P.sub}
    out["Medicine assessment"] = {
        fkey(r[M.FAC], r[M.DIST]): (clean(r[M.FAC]), prov(r[M.PROV]),
                               clean(r[M.DIST]), sector(r[M.SECT]))
        for r in M.SUB}
    return out


def plan_matches(plan, seen):
    """Which plan keys this tool reached.

    Delegates to Site_coverage_check.reconcile so this file and
    Site_coverage_check.xlsx can never report different completion - the exact
    match, the confirmed aliases, the level fallback and the district-bounded
    fuzzy pass all apply here too.
    """
    matched, _, _ = S.reconcile(plan, seen)
    return set(matched)


def sheet(wb, title, headers, rows, notes=None, fills=None, total=None):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    if notes:
        for n in notes:
            c = ws.cell(row=r, column=1, value=n)
            c.font = Font(size=9, italic=True, color="555555")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r,
                           end_column=max(len(headers), 3))
            ws.row_dimensions[r].height = 13 * (1 + len(n) // 110)
            r += 1
        r += 1
    widths = {}
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font, c.fill, c.border = Font(bold=True, size=10), HDR_FILL, BORDER
        c.alignment = Alignment(wrap_text=True, vertical="bottom")
        widths[ci] = min(len(str(h)) + 2, 30)
    hdr_row = r
    body = list(rows) + ([total] if total else [])
    for ri, row in enumerate(body):
        r += 1
        is_total = total is not None and ri == len(body) - 1
        for ci, v in enumerate(row, start=1):
            c = ws.cell(row=r, column=ci, value=v)
            c.border = BORDER
            c.alignment = Alignment(vertical="top",
                                    wrap_text=isinstance(v, str) and len(str(v)) > 38)
            if isinstance(v, float) and 0 <= v <= 1:
                c.number_format = "0%"
            if is_total:
                c.font = Font(bold=True)
                c.fill = TOT_FILL
            elif fills:
                f = fills(row, ci - 1, v)
                if f:
                    c.fill = f
            widths[ci] = max(widths.get(ci, 9), min(len(str(v)) + 2, 40))
        widths.setdefault(1, 14)
    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)
    ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(len(headers))}{r}"
    for ci, w in widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w


def build():
    plan = S.load_plan()
    tools = collect()
    got = {t: plan_matches(plan, tools[t]) for t in TOOLS}
    med, pres = tools["Medicine assessment"], tools["Prescription"]
    allk = set(med) | set(pres)

    wb = Workbook()
    wb.remove(wb.active)

    # ---- helper: aggregate the facility union at one geographic level
    def group(level):
        g = defaultdict(lambda: {"med": set(), "pres": set()})
        for k in allk:
            for tag, src in (("med", med), ("pres", pres)):
                if k in src:
                    name, pv, dist, sec = src[k]
                    where = {"province": (pv,), "district": (pv, dist),
                             "sector": (pv, dist, sec)}[level]
                    g[where][tag].add(k)
        return g

    # ---- planned facilities per area, for the completion columns
    plan_at = defaultdict(set)
    for k, p in plan.items():
        plan_at[("province", (prov(p["province"]),))].add(k)
        plan_at[("district", (prov(p["province"]), p["district"]))].add(k)

    def rows_for(level, label_cols):
        out = []
        for where, d in sorted(group(level).items()):
            m, p = d["med"], d["pres"]
            both, only_m, only_p = m & p, m - p, p - m
            planned = plan_at.get((level, where), set()) if level != "sector" else set()
            row = list(where) + [len(m | p), len(m), len(p), len(both),
                                 len(only_m), len(only_p)]
            if level != "sector":
                row += [len(planned),
                        len(planned & got["Medicine assessment"]),
                        len(planned & got["Prescription"]),
                        (len(planned & got["Medicine assessment"]) / len(planned)
                         if planned else None),
                        (len(planned & got["Prescription"]) / len(planned)
                         if planned else None)]
            out.append(row)
        return out

    base = ["Facilities touched", "Medicine assessment", "Prescription",
            "Both tools", "Medicine only", "Prescription only"]
    plan_cols = ["Planned facilities", "Planned done - medicine",
                 "Planned done - prescription", "% complete - medicine",
                 "% complete - prescription"]

    def gapfill(row, i, v):
        # highlight a level where the two tools disagree
        if i in (4, 5) and isinstance(v, int) and v > 0:
            return GAP_FILL
        return None

    def total_row(level, ncols):
        m, p = set(med), set(pres)
        both, om, op = m & p, m - p, p - m
        row = ["ALL"] + [""] * (ncols - 1) + [len(allk), len(m), len(p),
                                              len(both), len(om), len(op)]
        if level != "sector":
            row += [len(plan), len(got["Medicine assessment"]),
                    len(got["Prescription"]),
                    len(got["Medicine assessment"]) / len(plan),
                    len(got["Prescription"]) / len(plan)]
        return row

    NOTE_AGREE = ("'Medicine only' / 'Prescription only' are facilities one tool "
                  "reached and the other did not - the return-trip list. They are "
                  "shaded red wherever they are above zero.")
    NOTE_PLAN = ("'% complete' is against Sites_to_be_visited_combined_1.xlsx, "
                 "which lists 170 facilities in 19 districts. Areas with no "
                 "planned facilities show blank, not 0% - the plan simply does "
                 "not cover them.")

    sheet(wb, "By province", ["Province"] + base + plan_cols,
          rows_for("province", 1), notes=[NOTE_AGREE, NOTE_PLAN],
          fills=lambda r, i, v: gapfill(r, i - 0, v) if i >= 1 else None,
          total=total_row("province", 1))
    sheet(wb, "By district", ["Province", "District"] + base + plan_cols,
          rows_for("district", 2), notes=[NOTE_AGREE, NOTE_PLAN],
          fills=lambda r, i, v: gapfill(r, i - 1, v) if i >= 2 else None,
          total=total_row("district", 2))
    sheet(wb, "By sector", ["Province", "District", "Sector"] + base,
          rows_for("sector", 3),
          notes=[NOTE_AGREE,
                 "The plan has no sector column, so there are no completion "
                 "columns here. Sector names are normalised: the prescription "
                 "tool writes 'Masaka__Kicukiro', the stock tool 'Masaka'."],
          fills=lambda r, i, v: gapfill(r, i - 2, v) if i >= 3 else None,
          total=total_row("sector", 3))

    # ---- facility level
    rows = []
    for k in allk:
        name, pv, dist, sec = med.get(k) or pres.get(k)
        mn = med[k][0] if k in med else ""
        pn = pres[k][0] if k in pres else ""
        planned = k in plan
        rows.append([pv, dist, sec, mn or pn,
                     "yes" if k in med else "NO",
                     "yes" if k in pres else "NO",
                     "both" if (k in med and k in pres)
                     else ("medicine only" if k in med else "prescription only"),
                     "yes" if planned else "no",
                     mn if mn and pn and mn != pn else ""])
    rows.sort(key=lambda r: (r[6] == "both", r[0], r[1], r[3]))

    def ffill(row, i, v):
        if i in (4, 5):
            return GAP_FILL if v == "NO" else BOTH_FILL
        return None

    sheet(wb, "By facility",
          ["Province", "District", "Sector", "Facility", "Medicine assessment",
           "Prescription", "Status", "On the plan?", "Name differs (medicine)"],
          rows,
          notes=["One row per facility either tool has reached. Sorted so the "
                 "mismatches come first.",
                 "The last column is filled only where the two tools spell the "
                 "same facility differently."],
          fills=ffill)

    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)      # Box resets synced files to read-only
    wb.save(OUT)

    om, op = set(med) - set(pres), set(pres) - set(med)
    print(f"wrote {OUT}")
    print(f"  facilities touched by either tool : {len(allk)}")
    print(f"    both tools                      : {len(set(med) & set(pres))}")
    print(f"    medicine assessment only        : {len(om)}")
    print(f"    prescription only               : {len(op)}")
    print(f"  of the {len(plan)} planned facilities: "
          f"medicine {len(got['Medicine assessment'])} "
          f"({100*len(got['Medicine assessment'])/len(plan):.0f}%), "
          f"prescription {len(got['Prescription'])} "
          f"({100*len(got['Prescription'])/len(plan):.0f}%)")


if __name__ == "__main__":
    build()
