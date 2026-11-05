"""Availability of carbamazepine and valproate, on the MANAGED denominator.

    python3 Anticonvulsant_availability.py
      -> Anticonvulsant_availability.xlsx

Requested 2026-10-11. The denominator is the facilities that MANAGE the
medicine and answered the availability question - not every facility the EML
authorises it for, which is the denominator the main dashboard uses. The two
answer different questions and this file reports both side by side so the
difference is visible:

    MANAGED denominator     of the facilities that carry it, who had it?
    EML denominator         of the facilities authorised to carry it, who
                            could have dispensed it?

VALPROATE IS NOT IN THE STOCK ASSESSMENT
The stock tool assesses 79 medicines and valproate is not one of them, so it
has no managed denominator and no availability figure from that instrument.
Carbamazepine is the only anticonvulsant in the basket. Valproate is reported
here from the PRESCRIPTION tool only, on its own sheet, with the warning that
the two instruments' rates are not comparable.

SPELLING
The stock tool writes "Carbamezepine" (e, not a) throughout. Both spellings
are matched so a later correction to the form does not silently drop the
medicine from this file.
"""

import _bootstrap as B   # path setup - must come first
import re
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_kobo_prescription as P
import analyze_medicine_assessment as M
import explorer as EX
import facility_levels as FL
import molecules as MOL

OUT = B.out("Anticonvulsant_availability.xlsx")
MIN_N = 10

CARBA = re.compile(r"carbam[ae]z", re.I)
VALPRO = re.compile(r"valpro", re.I)
NOT_VALPRO = re.compile(r"methocarbamol", re.I)

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


# ---------------------------------------------------------------
# which target molecules exist in the stock assessment
# ---------------------------------------------------------------
# Written to take a LIST of target molecules so that if valproate is added to
# the form later this file picks it up with no edit. Each target reports
# whether it was found, so an absent medicine is stated rather than silently
# producing an empty table.
TARGETS = [("Carbamazepine", CARBA, None),
           ("Valproate sodium / valproic acid", VALPRO, NOT_VALPRO)]


def match(name, rx, anti):
    return bool(rx.search(name)) and not (anti and anti.search(name))


FOUND, MISSING = [], []
for tname, rx, anti in TARGETS:
    strengths = sorted({s(r[M.MED]) for r in M.MEDROWS
                        if match(s(r[M.MED]), rx, anti)})
    (FOUND if strengths else MISSING).append((tname, rx, anti, strengths))

ROWS = [r for r in M.MEDROWS
        if any(match(s(r[M.MED]), rx, anti) for _, rx, anti, _ in FOUND)]
STRENGTHS = sorted({s(r[M.MED]) for r in ROWS})
for r in ROWS:
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])
    r["_mgt"] = s(r["_p"].get(EX.MGMT)) or "(not recorded)"
    for tname, rx, anti, _ in FOUND:
        if match(s(r[M.MED]), rx, anti):
            r["_target"] = tname

MGD = [r for r in ROWS if M.yes(r[M.MANAGED]) and not M.blank(r[M.AVAIL])]


# ---------------------------------------------------------------
# combining strengths: Method 3, by where the medicine is AVAILABLE
# ---------------------------------------------------------------
# Two strengths of one molecule are therapeutically interchangeable, so the
# clinical question is whether the facility could treat the patient at all -
# not whether it held each strength separately. For every facility:
#
#     manages the molecule   if it manages ANY strength
#     has the molecule       if ANY strength it manages was on the shelf
#
# A facility holding only the 200mg therefore counts as served. Averaging the
# two strengths instead would drag the molecule down wherever one strength is
# rare, which is exactly the case here: the 400mg is managed by 17 facilities
# against 233 for the 200mg.
#
# The COMBINED group applies the same rule once more across molecules: a
# facility is counted as able to treat if any target anticonvulsant it
# manages was present. With valproate absent from the form this equals the
# carbamazepine line; it is built properly so the figure stays correct when
# valproate is added.
def combine(rows, keyfn):
    """-> {facility: (managed_any, available_any)} per key."""
    out = defaultdict(lambda: defaultdict(lambda: [False, False]))
    for r in rows:
        if not M.yes(r[M.MANAGED]) or M.blank(r[M.AVAIL]):
            continue
        g = out[keyfn(r)][s(r["_p"][M.FAC])]
        g[0] = True
        g[1] = g[1] or M.yes(r[M.AVAIL])
    return out


BY_MOL = combine(ROWS, lambda r: r["_target"])
BY_ALL = combine(ROWS, lambda r: "ANY TARGET ANTICONVULSANT")
FACSTATE = BY_ALL["ANY TARGET ANTICONVULSANT"]
for r in MGD:
    r["_fac_ok"] = FACSTATE[s(r["_p"][M.FAC])][1]


def fac_line(state, label):
    n = len(state)
    av = sum(1 for v in state.values() if v[1])
    return [label, n, av, n - av, pct(av, n)]


# Facility-level header: the denominator is facilities, so there is one count
# column, not the two the observation-level tables carry.
FHDR = ["Group", "Facilities managing it", "Had it on the visit day",
        "Did not have it", "AVAILABILITY"]
FPCT = {4}


HDR = ["Group", "Facilities", "Managed and answered", "Available",
       "Not available", "AVAILABILITY"]
PCT = {5}


def line(rows, label):
    if not rows:
        return None
    av = sum(1 for r in rows if M.yes(r[M.AVAIL]))
    return [label, len({s(r["_p"][M.FAC]) for r in rows}), len(rows), av,
            len(rows) - av, pct(av, len(rows))]


def by(keyfn, keys=None, rows=None, sort=True):
    """Observation-level cut - each strength counted on its own."""
    g = defaultdict(list)
    for r in (rows if rows is not None else MGD):
        g[keyfn(r)].append(r)
    out = [line(g[k], k) for k in (keys or sorted(g)) if g.get(k)]
    if sort:
        out.sort(key=lambda x: (x[5] is None, x[5]))
    return out


def by_fac(keyfn, keys=None, sort=True):
    """Facility-level cut with the strengths COMBINED.

    One line per facility, served if any strength it manages was present.
    This is the cut that matches the combined headline; `by()` above is kept
    for the per-strength tables, where the unit really is the observation.
    """
    g = defaultdict(dict)
    for r in MGD:
        f = s(r["_p"][M.FAC])
        cur = g[keyfn(r)].get(f, False)
        g[keyfn(r)][f] = cur or M.yes(r[M.AVAIL])
    out = []
    for k in (keys or sorted(g)):
        st = g.get(k)
        if not st:
            continue
        av = sum(1 for v in st.values() if v)
        out.append([k, len(st), av, len(st) - av, pct(av, len(st))])
    if sort:
        out.sort(key=lambda x: (x[4] is None, x[4]))
    return out


def sheet(wb, title, blocks, notes=None, widths=None, freeze="B"):
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
                    c.fill = (OK_FILL if v >= 0.95
                              else MID_FILL if v >= 0.85 else BAD_FILL)
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
    nat = line(MGD, "ALL CARBAMAZEPINE")
    wide = {1: 40, 2: 12, 3: 20, 4: 12, 5: 14, 6: 14}

    # --- 1. headline, both denominators
    rows = []
    for m in STRENGTHS:
        rs = [r for r in ROWS if s(r[M.MED]) == m]
        app = [r for r in rs if M.yes(r[M.APPLICABLE])]
        mg = [r for r in rs if M.yes(r[M.MANAGED]) and not M.blank(r[M.AVAIL])]
        av = sum(1 for r in mg if M.yes(r[M.AVAIL]))
        rows.append([m, len(rs), len(app), len(mg), av,
                     pct(av, len(mg)), pct(av, len(app))])
    allmg = len(MGD)
    allav = sum(1 for r in MGD if M.yes(r[M.AVAIL]))
    allapp = sum(1 for r in ROWS if M.yes(r[M.APPLICABLE]))
    tot = ["BOTH STRENGTHS", len(ROWS), allapp, allmg, allav,
           pct(allav, allmg), pct(allav, allapp)]

    mol = [r for r in MOL.collapse(M.MEDROWS)
           if CARBA.search(s(r[M.MED]))
           and M.yes(r[M.MANAGED]) and not M.blank(r[M.AVAIL])]
    molav = sum(1 for r in mol if M.yes(r[M.AVAIL]))

    # facility-level, strengths combined by where available
    combined = fac_line(FACSTATE, "COMBINED - any target anticonvulsant")
    permol = [fac_line(BY_MOL[t], t) for t in sorted(BY_MOL)]

    sheet(wb, "Summary", [
        ("COMBINED - strengths merged by where the medicine is available",
         FHDR, permol, combined, FPCT),
        ("Per strength, counted separately",
         ["Medicine", "Observations", "Applicable per EML",
          "MANAGED and answered", "Available",
          "AVAILABILITY (managed denominator)",
          "Availability (EML denominator)"],
         rows, tot, {5, 6}),
        ("What the combining changes",
         ["Measure", "Value", "Share"],
         [["Facilities managing at least one strength", len(FACSTATE), ""],
          ["Had at least one strength on the shelf",
           sum(1 for v in FACSTATE.values() if v[1]),
           pct(sum(1 for v in FACSTATE.values() if v[1]), len(FACSTATE))],
          ["Had none of the strengths they manage",
           sum(1 for v in FACSTATE.values() if not v[1]),
           pct(sum(1 for v in FACSTATE.values() if not v[1]),
               len(FACSTATE))],
          ["", "", ""],
          ["Observation-level rate, strengths kept separate", allmg,
           pct(allav, allmg)],
          ["Facility-level rate, strengths combined", len(FACSTATE),
           pct(sum(1 for v in FACSTATE.values() if v[1]), len(FACSTATE))]],
         None, {2}),
    ], notes=[
        "MEDICINES REQUESTED: carbamazepine, and valproate sodium / valproic "
        "acid. "
        + ("FOUND IN THE STOCK ASSESSMENT: "
           + "; ".join(f"{t} ({len(st)} strength"
                       + ("s" if len(st) > 1 else "") + ")"
                       for t, _, _, st in FOUND) + ". ")
        + ("NOT IN THE STOCK ASSESSMENT: "
           + "; ".join(t for t, _, _, _ in MISSING)
           + " - the form assesses a fixed basket and this medicine is not "
             "in it, so it has no managed denominator and no stock "
             "availability figure. It appears only on the Prescription tool "
             "sheet." if MISSING else ""),
        "COMBINING RULE: for each facility, a molecule counts as MANAGED if "
        "the facility manages any strength of it, and AVAILABLE if any "
        "strength it manages was on the shelf on the visit day. So a facility "
        "holding only the 200mg counts as served. The denominator is "
        "facilities, not observations.",
        "DENOMINATOR AS REQUESTED: facilities that MANAGE the medicine. The "
        "EML column in the second block counts every facility the EML "
        "authorises it for, which is what the main dashboard uses.",
        "The stock tool spells the medicine 'Carbamezepine'. Both spellings "
        "are matched here.",
    ], widths={1: 42, 2: 20, 3: 20, 4: 20, 5: 16, 6: 20, 7: 20})

    # --- 2. disaggregations
    LV = FL.ordered({r["_lvl"] for r in MGD})
    sheet(wb, "By level and management", [
        ("By level of care - strengths combined", FHDR,
         by_fac(lambda r: r["_lvl"], keys=LV, sort=False), combined, FPCT),
        ("By management type - strengths combined", FHDR,
         by_fac(lambda r: r["_mgt"]), combined, FPCT),
        ("By level of care - each strength on its own", HDR,
         by(lambda r: r["_lvl"], keys=LV, sort=False), nat, PCT),
        ("By strength x level of care", HDR,
         [l for m in STRENGTHS
          for l in by(lambda r: f"{m} - {r['_lvl']}",
                      rows=[r for r in MGD if s(r[M.MED]) == m], sort=False)],
         None, PCT),
    ], notes=["The first two blocks are the COMBINED view - one line per "
              "facility, served if any strength it manages was present. The "
              "last two keep the strengths apart, which is where the 400mg's "
              "small base is visible."],
        widths=wide)

    sheet(wb, "By geography", [
        ("By province - strengths combined", FHDR,
         by_fac(lambda r: s(r["_p"][M.PROV])), combined, FPCT),
        ("By district - strengths combined", FHDR,
         by_fac(lambda r: s(r["_p"][M.DIST])), combined, FPCT),
        ("By province - each strength on its own", HDR,
         by(lambda r: s(r["_p"][M.PROV])), nat, PCT),
    ], notes=["District cells rest on very few facilities - the whole "
              "analysis covers 235 facilities across 30 districts - so read "
              "the counts beside every rate rather than ranking the "
              "percentages."],
        widths=wide)

    # --- 3. facility list
    frows = sorted([[s(r["_p"][M.FAC]), s(r["_p"][M.DIST]),
                     s(r["_p"][M.PROV]), r["_lvl"], r["_mgt"], s(r[M.MED]),
                     "Available" if M.yes(r[M.AVAIL]) else "NOT AVAILABLE",
                     M.num(r[M.QTY_PHYS]), M.num(r[M.AMC]),
                     M.num(r[M.MOS]), s(r[M.STOCKCLASS])]
                    for r in MGD],
                   key=lambda x: (x[6] == "Available", x[0]))
    sheet(wb, "Facilities managing it", [
        ("Every facility x strength that manages carbamazepine",
         ["Facility", "District", "Province", "Level of care",
          "Management", "Strength", "On the visit day", "Quantity on hand",
          "Avg monthly consumption", "Months of stock", "Stock class"],
         frows, None, set())],
        notes=["Not-available rows are listed first.",
               "Quantity, AMC and months of stock are item level and are read "
               "straight from the form."],
        widths={1: 30, 2: 15, 3: 13, 4: 24, 5: 24, 6: 22, 7: 17, 8: 15,
                9: 20, 10: 15, 11: 16})

    # --- 4. not available
    na = [r for r in MGD if M.no(r[M.AVAIL])]
    sheet(wb, "Not available", [
        ("Facilities that manage carbamazepine and did not have it",
         ["Facility", "District", "Province", "Level of care", "Strength",
          "Stock-out days Jun-Aug", "Months affected",
          "Requested Jun-Aug", "Delivered Jun-Aug"],
         sorted([[s(r["_p"][M.FAC]), s(r["_p"][M.DIST]), s(r["_p"][M.PROV]),
                  r["_lvl"], s(r[M.MED]),
                  sum(M.num(r[M.SO_MONTH[m]]) or 0
                      for m in ("June", "July", "August")),
                  sum(1 for m in ("June", "July", "August")
                      if (M.num(r[M.SO_MONTH[m]]) or 0) > 0),
                  M.num(r[M.REQ_TOT]), M.num(r[M.DEL_TOT])]
                 for r in na], key=lambda x: -(x[5] or 0)),
         None, set())],
        notes=[f"{len(na)} of {len(MGD)} managed observations.",
               "Stock-out days are summed from the three monthly fields, not "
               "from the form's total_stockout_days, which adds the ongoing "
               "current run and can exceed the 92-day window."],
        widths={1: 30, 2: 15, 3: 13, 4: 24, 5: 22, 6: 20, 7: 15, 8: 17,
                9: 17})

    # --- 5. prescription tool, both medicines
    def pgroup(r):
        t = (s(r[P.MED_EML]) + " " + s(r[P.MED_FREE]) + " "
             + s(r.get("Medicine (name)")))
        if NOT_VALPRO.search(t):
            return None
        if CARBA.search(t):
            return "Carbamazepine"
        if VALPRO.search(t):
            return "Valproic acid / valproate sodium"
        return None

    pitems = defaultdict(list)
    for r in P.item:
        if not r["_sub"]:
            continue
        k = pgroup(r)
        if k:
            pitems[k].append(r)
    prows = []
    for k in ("Carbamazepine", "Valproic acid / valproate sodium"):
        rs = pitems[k]
        ans = [r for r in rs if r["_avail"] is not None]
        av = sum(1 for r in ans if r["_avail"] is True)
        prows.append([k, len(rs), len({s(r["_sub"][P.FACIL]) for r in rs}),
                      len(ans), av, pct(av, len(ans))])
    det = sorted([[k, s(r["_sub"][P.FACIL]), s(r["_sub"][P.DISTR]),
                   s(r.get("Medicine (name)")) or s(r[P.MED_EML])
                   or s(r[P.MED_FREE]),
                   "Available" if r["_avail"] else "NOT AVAILABLE"]
                  for k, rs in pitems.items() for r in rs],
                 key=lambda x: (x[4] == "Available", x[0], x[1]))
    sheet(wb, "Prescription tool", [
        ("As prescribed, from the unserved prescription assessment",
         ["Medicine", "Items prescribed", "Facilities", "Answered",
          "Available", "AVAILABILITY"], prows, None, {5}),
        ("Every line", ["Medicine", "Facility", "District",
                        "As recorded", "Outcome"], det, None, set()),
    ], notes=[
        "VALPROATE IS NOT IN THE STOCK ASSESSMENT BASKET. The stock tool "
        "assesses 79 medicines and valproate is not among them, so it has no "
        "managed denominator and no stock availability figure. This sheet is "
        "the only place it appears.",
        "DO NOT COMPARE these rates with the stock sheets. The denominator "
        "here is prescriptions written, not facilities stocking, and an item "
        "is mostly recorded when a prescription was being filled - so almost "
        "every medicine scores about 98% in this instrument. It is a design "
        "artefact, not a supply result.",
    ], widths={1: 34, 2: 17, 3: 12, 4: 12, 5: 12, 6: 14})

    wb.save(OUT)
    return rows, tot, len(mol), molav, prows


if __name__ == "__main__":
    rows, tot, nmol, molav, prows = build()
    print(f"wrote {OUT}")
    for r in rows:
        print(f"  {r[0][:30]:<32}managed {r[3]:>4}  available {r[4]:>4}  "
              f"{r[5]:>7.1%}")
    print(f"  {'BOTH STRENGTHS':<32}managed {tot[3]:>4}  available {tot[4]:>4}"
          f"  {tot[5]:>7.1%}")
    print(f"  {'molecule (either strength)':<32}managed {nmol:>4}  "
          f"available {molav:>4}  {molav / nmol:>7.1%}")
    for p in prows:
        print(f"  prescription tool: {p[0][:34]:<36}{p[4]}/{p[3]} = {p[5]:.1%}")
