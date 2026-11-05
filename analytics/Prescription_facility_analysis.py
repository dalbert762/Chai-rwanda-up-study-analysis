"""Facility-level analysis of the prescription assessment.

    python3 Prescription_facility_analysis.py
      -> Prescription_facility_analysis.xlsx

Imports analyze_kobo_prescription and reuses its joined rows and derived flags
(_avail, _abx, _is_inj, _med, _n_items, _any_unserved ...) rather than
recomputing anything, so this cannot disagree with the main workbook. Facility
names are post-reattribution - the FACILITY_OVERRIDES corrections apply here.

The unit is the FACILITY. Each sheet answers a different question about it:

  Scorecard        one row per facility, every measure side by side
  WHO indicators   each facility against the WHO/INRUD optimum ranges
  By facility type what differs between HC, DH, L2TH, HP ...
  Outliers         the ranked tails worth a phone call
  Data quality     facilities whose data should be checked before it is used

Rates are computed over that facility's own denominators, and any facility
with fewer than MIN_PRES prescriptions is marked "thin" rather than dropped:
a 2-prescription facility can show 0% or 100% on anything, and that is a
sample-size artefact, not a finding.

openpyxl + the standard library, like the rest of this project.
"""

import _bootstrap as B   # path setup - must come first
import os
import statistics
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_kobo_prescription as P

OUT = B.out("Prescription_facility_analysis.xlsx")
MIN_PRES = 5          # below this a facility's rates are not interpretable

# WHO / INRUD optimum ranges, as the main workbook states them.
WHO = {
    "items_per_enc": (1.6, 1.8, "1.6 - 1.8"),
    "inn": (1.0, 1.0, "100%"),
    "abx": (0.20, 0.268, "20 - 26.8%"),
    "inj": (0.134, 0.241, "13.4 - 24.1%"),
    "eml": (1.0, 1.0, "100%"),
}

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="DDE5F0")
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")
BAD_FILL = PatternFill("solid", fgColor="FCE4E4")
OK_FILL = PatternFill("solid", fgColor="E7F3E7")
THIN_FILL = PatternFill("solid", fgColor="FFF4D6")


def s(v):
    return "" if v is None else str(v).strip()


def pct(a, b):
    return a / b if b else None


def facility_rows():
    """Group the three joined levels by facility."""
    subs = defaultdict(list)
    for r in P.sub:
        subs[s(r[P.FACIL])].append(r)
    pres = defaultdict(list)
    for r in P.pres:
        if r["_sub"]:
            pres[s(r["_sub"][P.FACIL])].append(r)
    items = defaultdict(list)
    for r in P.item:
        if r["_sub"]:
            items[s(r["_sub"][P.FACIL])].append(r)
    return subs, pres, items


SUBS, PRES, ITEMS = facility_rows()
FACILITIES = sorted(SUBS)


def measures(fac):
    """Every measure for one facility, as a dict."""
    sb, pr, it = SUBS[fac], PRES.get(fac, []), ITEMS.get(fac, [])
    n_pres, n_item = len(pr), len(it)
    avail = [r for r in it if r["_avail"] is True]
    unavail = [r for r in it if r["_avail"] is False]
    answered = len(avail) + len(unavail)
    cost = sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in unavail)
    inn = sum(1 for r in it if s(r.get(P.NAME_TYPE)).lower().startswith("inn"))
    eml = sum(1 for r in it if P.yes(r.get(P.EML)))
    abx_enc = sum(1 for r in pr if r["_abx"])
    inj_enc = sum(1 for r in pr if r["_inj"])
    dates = sorted(s(r[P.SUB_TIME])[:10] for r in sb)
    return {
        "facility": fac,
        "province": s(sb[0][P.PROV]),
        "district": s(sb[0][P.DISTR]),
        "sector": s(sb[0].get("Sector")).split("__")[0],
        "type": s(sb[0][P.FTYPE]),
        "visits": len(sb),
        "prescriptions": n_pres,
        "items": n_item,
        "items_per_enc": pct(n_item, n_pres),
        "inn": pct(inn, n_item),
        "abx": pct(abx_enc, n_pres),
        "inj": pct(inj_enc, n_pres),
        "eml": pct(eml, n_item),
        "available": pct(len(avail), answered),
        "unserved_items": len(unavail),
        "unserved_cost": round(cost),
        "pres_with_unserved": sum(1 for r in pr if r["_any_unserved"]),
        "count_mismatch": sum(1 for r in pr if r["_count_mismatch"]),
        "collectors": len({s(r[P.COLLECTOR]) for r in sb}),
        "collector_names": ", ".join(sorted({s(r[P.COLLECTOR]) for r in sb
                                             if s(r[P.COLLECTOR])})[:3]),
        "first": dates[0] if dates else "",
        "last": dates[-1] if dates else "",
        "server": "/".join(sorted({s(r.get("_server")) for r in sb})),
        "thin": n_pres < MIN_PRES,
        "medicines": len({r["_med"] for r in it}),
    }


ROWS = [measures(f) for f in FACILITIES]


def sheet(wb, title, headers, rows, notes=None, fills=None, total=None,
          pctcols=(), num0=()):
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
        widths[ci] = min(max(len(str(h)) + 2, 9), 26)
    hdr_row = r
    body = list(rows) + ([total] if total else [])
    for ri, row in enumerate(body):
        r += 1
        is_total = total is not None and ri == len(body) - 1
        for ci, v in enumerate(row, start=1):
            c = ws.cell(row=r, column=ci, value=v)
            c.border = BORDER
            c.alignment = Alignment(vertical="top",
                                    wrap_text=isinstance(v, str) and len(str(v)) > 34)
            if (ci - 1) in pctcols and isinstance(v, (int, float)):
                c.number_format = "0.0%"
            elif (ci - 1) in num0 and isinstance(v, (int, float)):
                c.number_format = "#,##0"
            elif isinstance(v, float):
                c.number_format = "0.00"
            if is_total:
                c.font = Font(bold=True)
                c.fill = TOT_FILL
            elif fills:
                f = fills(row, ci - 1, v)
                if f:
                    c.fill = f
            widths[ci] = max(widths.get(ci, 9), min(len(str(v)) + 2, 34))
    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=4)
    ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(len(headers))}{r}"
    for ci, w in widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w


def build():
    wb = Workbook()
    wb.remove(wb.active)
    thin_n = sum(1 for m in ROWS if m["thin"])

    # ---------- 1. scorecard
    hdr = ["Facility", "Province", "District", "Sector", "Type", "Visits",
           "Prescriptions", "Items", "Distinct medicines", "Items / encounter",
           "% generic (INN)", "% enc. with antibiotic", "% enc. with injection",
           "% from EML", "% items available", "Unserved items",
           "Unserved cost (RWF)", "Prescriptions with an unserved item",
           "Collectors", "First visit", "Last visit", "Server", "Thin sample?"]
    rows = [[m["facility"], m["province"], m["district"], m["sector"], m["type"],
             m["visits"], m["prescriptions"], m["items"], m["medicines"],
             m["items_per_enc"], m["inn"], m["abx"], m["inj"], m["eml"],
             m["available"], m["unserved_items"], m["unserved_cost"],
             m["pres_with_unserved"], m["collectors"], m["first"], m["last"],
             m["server"], "yes" if m["thin"] else ""]
            for m in sorted(ROWS, key=lambda m: -m["items"])]
    tot = ["ALL FACILITIES", "", "", "", "", len(P.sub), P.N_PRES, P.N_ITEM,
           len({r["_med"] for r in P.item}),
           pct(P.N_ITEM, P.N_PRES),
           pct(sum(1 for r in P.item if s(r.get(P.NAME_TYPE)).lower().startswith("inn")),
               P.N_ITEM),
           pct(sum(1 for r in P.pres if r["_abx"]), P.N_PRES),
           pct(sum(1 for r in P.pres if r["_inj"]), P.N_PRES),
           pct(sum(1 for r in P.item if P.yes(r.get(P.EML))), P.N_ITEM),
           pct(P.ITEM_AV, P.ITEM_AV + P.ITEM_NA), P.ITEM_NA,
           round(sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in P.item
                     if r["_avail"] is False)),
           sum(1 for r in P.pres if r["_any_unserved"]),
           len({s(r[P.COLLECTOR]) for r in P.sub}), "", "", "", ""]

    def scfill(row, i, v):
        if row[22] == "yes" and i >= 9:
            return THIN_FILL
        return None

    sheet(wb, "Facility scorecard", hdr, rows,
          notes=[f"One row per facility, {len(ROWS)} in all, sorted by number of "
                 f"medicine items. Facility names are post-reattribution.",
                 f"Amber = fewer than {MIN_PRES} prescriptions ({thin_n} "
                 f"facilities). Their rates are shown but should not be compared "
                 f"- a 2-prescription facility reads 0% or 100% on anything."],
          fills=scfill, total=tot,
          pctcols=(10, 11, 12, 13, 14), num0=(5, 6, 7, 8, 15, 16, 17, 18))

    # ---------- 2. WHO indicators per facility
    hdr = ["Facility", "District", "Type", "Prescriptions",
           "Items / encounter", f"vs {WHO['items_per_enc'][2]}",
           "% generic (INN)", f"vs {WHO['inn'][2]}",
           "% enc. with antibiotic", f"vs {WHO['abx'][2]}",
           "% enc. with injection", f"vs {WHO['inj'][2]}",
           "% from EML", f"vs {WHO['eml'][2]}", "Indicators outside range"]

    def verdict(v, k):
        if v is None:
            return ""
        lo, hi, _ = WHO[k]
        if v < lo:
            return "below"
        if v > hi:
            return "above"
        return "in range"

    rows = []
    for m in sorted(ROWS, key=lambda m: -m["prescriptions"]):
        if m["thin"]:
            continue
        vs = [verdict(m[k], k) for k in ("items_per_enc", "inn", "abx", "inj", "eml")]
        rows.append([m["facility"], m["district"], m["type"], m["prescriptions"],
                     m["items_per_enc"], vs[0], m["inn"], vs[1], m["abx"], vs[2],
                     m["inj"], vs[3], m["eml"], vs[4],
                     sum(1 for x in vs if x not in ("in range", ""))])
    rows.sort(key=lambda r: -r[14])

    def wfill(row, i, v):
        if i in (5, 7, 9, 11, 13):
            return OK_FILL if v == "in range" else (BAD_FILL if v else None)
        return None

    sheet(wb, "WHO indicators", hdr, rows,
          notes=[f"Facilities with at least {MIN_PRES} prescriptions only "
                 f"({len(rows)} of {len(ROWS)}), sorted by how many of the five "
                 f"indicators fall outside the WHO/INRUD optimum.",
                 "Green = inside the optimum range, red = outside. 'Above' on "
                 "items per encounter or antibiotics is the direction of concern; "
                 "on generic and EML prescribing the optimum is 100%, so anything "
                 "short of it reads 'below' - which is why almost every facility "
                 "fails at least one. Read the count of indicators out of range, "
                 "and the direction, rather than pass/fail."],
          fills=wfill, pctcols=(6, 8, 10, 12), num0=(3, 14))

    # ---------- 3. by facility type
    g = defaultdict(list)
    for m in ROWS:
        g[m["type"] or "(not recorded)"].append(m)
    rows = []
    for t, ms in sorted(g.items(), key=lambda x: -sum(m["items"] for m in x[1])):
        np_, ni = sum(m["prescriptions"] for m in ms), sum(m["items"] for m in ms)
        it = [r for f in ms for r in ITEMS.get(f["facility"], [])]
        pr = [r for f in ms for r in PRES.get(f["facility"], [])]
        av = sum(1 for r in it if r["_avail"] is True)
        na = sum(1 for r in it if r["_avail"] is False)
        rows.append([t, len(ms), sum(m["visits"] for m in ms), np_, ni,
                     pct(ni, np_),
                     pct(sum(1 for r in it
                             if s(r.get(P.NAME_TYPE)).lower().startswith("inn")), ni),
                     pct(sum(1 for r in pr if r["_abx"]), np_),
                     pct(sum(1 for r in pr if r["_inj"]), np_),
                     pct(sum(1 for r in it if P.yes(r.get(P.EML))), ni),
                     pct(av, av + na), na,
                     round(sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in it
                               if r["_avail"] is False))])
    sheet(wb, "By facility type",
          ["Facility type", "Facilities", "Visits", "Prescriptions", "Items",
           "Items / encounter", "% generic (INN)", "% enc. with antibiotic",
           "% enc. with injection", "% from EML", "% items available",
           "Unserved items", "Unserved cost (RWF)"], rows,
          notes=["Rates are computed from summed numerators and denominators, "
                 "not as a mean of facility rates - otherwise a 2-prescription "
                 "health post would weigh the same as a 200-prescription hospital."],
          pctcols=(5, 6, 7, 8, 9, 10), num0=(1, 2, 3, 4, 11, 12))

    # ---------- 4. outliers
    ok = [m for m in ROWS if not m["thin"]]
    blocks = [
        ("Most medicine items per encounter", "items_per_enc", True, "0.00"),
        ("Lowest generic (INN) prescribing", "inn", False, "0.0%"),
        ("Highest antibiotic prescribing", "abx", True, "0.0%"),
        ("Highest injection prescribing", "inj", True, "0.0%"),
        ("Lowest EML adherence", "eml", False, "0.0%"),
        ("Lowest item availability", "available", False, "0.0%"),
        ("Highest unserved cost (RWF)", "unserved_cost", True, "#,##0"),
    ]
    rows = []
    for label, k, desc, _ in blocks:
        top = sorted([m for m in ok if m[k] is not None],
                     key=lambda m: m[k], reverse=desc)[:10]
        for i, m in enumerate(top):
            rows.append([label if i == 0 else "", m["facility"], m["district"],
                         m["type"], m["prescriptions"], m[k]])
    sheet(wb, "Outliers",
          ["Measure", "Facility", "District", "Type", "Prescriptions", "Value"],
          rows,
          notes=[f"Top 10 facilities on each measure, among the {len(ok)} with at "
                 f"least {MIN_PRES} prescriptions. The Value column carries "
                 f"whatever that measure is - a count, a ratio or a percentage."],
          num0=(4,))

    # ---------- 5. data quality
    rows = []
    for m in sorted(ROWS, key=lambda m: (-m["count_mismatch"], m["prescriptions"])):
        flags = []
        if m["thin"]:
            flags.append(f"fewer than {MIN_PRES} prescriptions")
        if m["count_mismatch"]:
            flags.append(f"{m['count_mismatch']} prescription(s) where the declared "
                         f"medicine count differs from the items entered")
        if m["collectors"] == 1 and m["prescriptions"] >= 20:
            flags.append("all data from a single collector")
        if not m["sector"]:
            flags.append("sector blank")
        if not m["type"]:
            flags.append("facility type blank")
        if m["items"] and m["available"] is None:
            flags.append("availability never answered")
        if flags:
            rows.append([m["facility"], m["district"], m["type"],
                         m["visits"], m["prescriptions"], m["items"],
                         m["collectors"], m["collector_names"], "; ".join(flags)])
    sheet(wb, "Data quality",
          ["Facility", "District", "Type", "Visits", "Prescriptions", "Items",
           "Collectors", "Collector(s)", "Flags"], rows,
          notes=["Only facilities carrying at least one flag. A flag is a reason "
                 "to check the record, not a finding in itself - a single "
                 "collector at a facility is normal where one person covered that "
                 "district."],
          num0=(3, 4, 5, 6))

    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)      # Box resets synced files to read-only
    wb.save(OUT)
    print(f"wrote {OUT}")
    print(f"  {len(ROWS)} facilities | {P.N_PRES:,} prescriptions | {P.N_ITEM:,} items")
    print(f"  {thin_n} facilities with fewer than {MIN_PRES} prescriptions "
          f"(rates flagged, not dropped)")
    # "outside at least one optimum" is not worth printing - two of the five
    # optimums are 100%, so nearly every facility fails one. The direction per
    # indicator is the finding.
    ok = [m for m in ROWS if not m["thin"]]
    print(f"  of the {len(ok)} interpretable facilities, against WHO/INRUD:")
    for k, label in (("items_per_enc", "items per encounter"),
                     ("inn", "generic (INN)"), ("abx", "antibiotic"),
                     ("inj", "injection"), ("eml", "from the EML")):
        c = Counter(verdict(m[k], k) for m in ok)
        print(f"    {label:20} above {c['above']:>3} | in range "
              f"{c['in range']:>3} | below {c['below']:>3}")


if __name__ == "__main__":
    build()
