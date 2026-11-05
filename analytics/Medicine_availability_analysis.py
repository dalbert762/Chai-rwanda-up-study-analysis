"""Medicine availability in the stock assessment, disaggregated every way the
data supports.

    python3 Medicine_availability_analysis.py
      -> Medicine_availability_analysis.xlsx

Imports analyze_medicine_assessment and reuses its joined rows and derived
fields, so nothing is recomputed and this cannot drift from the main workbook.

THE DENOMINATOR
'Available on visit day' is asked only where the facility MANAGES the medicine,
so availability is computed over answered-and-managed rows. Two earlier filters
matter and are shown as a funnel rather than hidden:

    observations -> applicable to this facility level per EML -> managed by the
    facility -> availability answered -> available / stocked out

A medicine that is not applicable to a health-centre level, or that a facility
has chosen not to stock, is not a stock-out. Reporting availability over all
observations would mix policy with supply.

WHAT CAN BE DISAGGREGATED
Every variable below is populated for effectively every row. The count in
brackets is its distinct values:

  facility side  province (5), district (30), sector (193), facility (267),
                 health facility level (9), facility_class (8),
                 management (4), hf_level_rank (3), server (2)
  medicine side  medicine (81), programme (8), drug class (33),
                 EML applicability (2), managed (2), stock classification (5)
  system         the electronic stock system named - free text, 162 raw
                 values, normalised here to a handful

RATES ON SMALL DENOMINATORS
A rate over a handful of rows is noise. Every table carries its denominator,
and any group with fewer than MIN_N answered rows is marked "thin" rather than
dropped - so a reader can see it exists without comparing it.

openpyxl + the standard library, like the rest of this project.
"""

import _bootstrap as B   # path setup - must come first
import os
import re
import statistics
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_medicine_assessment as M
import facility_levels as FL
import molecules as MOL

OUT = B.out("Medicine_availability_analysis.xlsx")
MIN_N = 30            # answered rows below which a rate is flagged as thin
TOP_FREETEXT = 12     # distinct free-text values kept before lumping to "other"

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


def emr(v):
    """The stock system named, as free text: 'ELMIS', 'eLMIS', 'E-LMIS' are one
    system. Normalised to a comparable label; anything rare becomes 'other'."""
    t = re.sub(r"[^a-z0-9]", "", s(v).lower())
    if not t:
        return ""
    for pat, label in (("elmis", "eLMIS"), ("ebuzima", "eBuzima"),
                       ("openmrs", "OpenMRS"), ("dhis", "DHIS2"),
                       ("excel", "Excel / spreadsheet"),
                       ("stockcard", "Stock cards (paper)"),
                       ("paper", "Stock cards (paper)"),
                       ("register", "Register (paper)"),
                       ("hmis", "HMIS"), ("smartcare", "SmartCare")):
        if pat in t:
            return label
    return s(v)


# ---------------------------------------------------------------
# the measure set, computed once per group
# ---------------------------------------------------------------
HDR = ["Group", "Observations", "Facilities", "Medicines",
       "Applicable per EML", "Managed", "Managed % of applicable",
       "Availability answered", "Available", "AVAILABLE %", "Stocked out",
       "Stock-out %", "Stock out", "Understock", "Plan stock", "Overstock",
       "Median months of stock", "Stock-out days (Jun-Aug)",
       "Qty requested", "Qty delivered", "Fulfilment %",
       "Recorded loss (RWF)", "Thin?"]
PCTCOLS = (6, 9, 11, 20)
NUM0 = (1, 2, 3, 4, 5, 7, 8, 10, 12, 13, 14, 15, 17, 18, 19, 21)


def measure(label, rows):
    applicable = [r for r in rows if M.yes(r[M.APPLICABLE])]
    managed = [r for r in rows if M.yes(r[M.MANAGED])]
    answered = [r for r in managed if not M.blank(r[M.AVAIL])]
    avail = [r for r in answered if M.yes(r[M.AVAIL])]
    out = [r for r in answered if M.no(r[M.AVAIL])]
    cls = Counter(s(r[M.STOCKCLASS]) for r in rows if not M.blank(r[M.STOCKCLASS]))
    mos = [M.num(r[M.MOS]) for r in managed]
    mos = [x for x in mos if x is not None]
    req = sum(M.num(r[M.REQ_TOT]) or 0 for r in rows)
    dlv = sum(M.num(r[M.DEL_TOT]) or 0 for r in rows)
    return [label, len(rows),
            len({s(r["_p"][M.FAC]) for r in rows}),
            len({s(r[M.MED]) for r in rows}),
            len(applicable), len(managed), pct(len(managed), len(applicable)),
            len(answered), len(avail), pct(len(avail), len(answered)),
            len(out), pct(len(out), len(answered)),
            cls.get("Stock out", 0), cls.get("Understock", 0),
            cls.get("Plan stock", 0), cls.get("Overstock", 0),
            round(statistics.median(mos), 2) if mos else None,
            round(sum(M.num(r[M.SO_TOTAL]) or 0 for r in rows)),
            round(req), round(dlv), pct(dlv, req),
            round(sum(M.num(r[M.LOSS]) or 0 for r in rows)),
            "thin" if len(answered) < MIN_N else ""]


def by(keyfn, rows=None, sort="avail"):
    g = defaultdict(list)
    for r in (rows if rows is not None else MEDROWS_M3):
        g[keyfn(r)].append(r)
    out = [measure(k if k else "(not recorded)", rs) for k, rs in g.items()]
    if sort == "avail":
        out.sort(key=lambda x: (x[22] == "thin", x[9] if x[9] is not None else 2))
    elif sort == "n":
        out.sort(key=lambda x: -x[1])
    else:
        out.sort(key=lambda x: str(x[0]))
    return out


# ---------------------------------------------------------------
# writer
# ---------------------------------------------------------------
def sheet(wb, title, blocks, notes=None, freeze_col=2):
    """blocks: list of (block title, headers, rows, total or None, pctcols, num0)"""
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    if notes:
        for n in notes:
            c = ws.cell(row=r, column=1, value=n)
            c.font = Font(size=9, italic=True, color="555555")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
            ws.row_dimensions[r].height = 13 * (1 + len(n) // 130)
            r += 1
        r += 1
    widths = {}
    for btitle, headers, rows, total, pctcols, num0 in blocks:
        if btitle:
            c = ws.cell(row=r, column=1, value=btitle)
            c.font = Font(bold=True, size=12, color="1F3864")
            r += 1
        for ci, h in enumerate(headers, start=1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font, c.fill, c.border = Font(bold=True, size=10), HDR_FILL, BORDER
            c.alignment = Alignment(wrap_text=True, vertical="bottom")
            widths[ci] = min(max(len(str(h)) + 2, 9), 24)
        r += 1
        body = list(rows) + ([total] if total else [])
        for ri, row in enumerate(body):
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
                elif row and row[-1] == "thin":
                    c.fill = THIN_FILL
                elif (ci - 1) == 9 and isinstance(v, float):
                    # assign only when there is a fill - openpyxl rejects None
                    if v < 0.90:
                        c.fill = BAD_FILL
                    elif v >= 0.95:
                        c.fill = OK_FILL
                widths[ci] = max(widths.get(ci, 9), min(len(str(v)) + 2, 34))
            r += 1
        r += 2
    for ci, w in widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    return ws


# DEFINITION (study, 2026-10-05): the denominator is every molecule the
# EML authorises at that level of care. A medicine the facility is
# supposed to manage but does not manage at all counts as UNAVAILABLE,
# not as out of scope. MOL.eml_base() applies that and keeps the real
# management flag on `_managed` / the reason on `_gap`.
MEDROWS_M3 = MOL.eml_base(M.MEDROWS)   # Method 3 molecules, EML denominator
ALL = measure("ALL OBSERVATIONS", MEDROWS_M3)
NOTE_DENOM = ("AVAILABLE % is over rows where the facility manages the medicine "
              "AND answered the availability question - never over all "
              "observations. 'Managed % of applicable' is the earlier filter: a "
              "medicine applicable to that facility level per the EML that the "
              "facility does not stock.")
NOTE_THIN = (f"Amber rows have fewer than {MIN_N} answered rows; the figures are "
             f"shown but a rate over a handful of observations is noise. Green "
             f"= 95% or better available, red = below 90%.")


def build():
    wb = Workbook()
    wb.remove(wb.active)

    # ---- 1. method + funnel
    funnel = [
        ["Medicine observations (facility x medicine)", M.N_MED, ""],
        ["Applicable to this facility level per the EML", ALL[4],
         f"{pct(ALL[4], M.N_MED):.1%} of observations"],
        ["Managed by the facility", ALL[5],
         f"{pct(ALL[5], ALL[4]):.1%} of applicable"],
        ["Availability answered", ALL[7],
         f"{pct(ALL[7], ALL[5]):.1%} of managed"],
        ["AVAILABLE on visit day", ALL[8],
         f"{pct(ALL[8], ALL[7]):.1%} of answered  <-- the headline"],
        ["Stocked out on visit day", ALL[10],
         f"{pct(ALL[10], ALL[7]):.1%} of answered"],
        ["", "", ""],
        ["Applicable per EML but NOT managed", ALL[4] - ALL[5],
         "not a stock-out - a facility selection or policy gap"],
        ["Not applicable to this facility level", M.N_MED - ALL[4],
         "correctly absent"],
    ]
    sheet(wb, "Method and funnel",
          [("How availability is derived", ["Step", "Rows", "Share"], funnel,
            None, (), (1,)),
           ("Stock classification, all classified observations",
            ["Class", "Observations", "% of classified"],
            [[c, n, pct(n, sum(M.CLASS_N.values()))]
             for c, n in M.CLASS_N.most_common()], None, (2,), (1,)),
           ("Supply-chain attribution for a stocked-out medicine",
            ["Attribution", "Observations", "% answered"],
            [[c, n, pct(n, sum(M.ATTRIB_N.values()))]
             for c, n in M.ATTRIB_N.most_common()], None, (2,), (1,))],
          notes=[NOTE_DENOM,
                 "METHOD 3: multi-strength medicines are collapsed to one "
                 "molecule per facility and count as available where ANY "
                 "strength was on the shelf - 81 items become 71 molecules. "
                 "Quantities, stock-out days and requisitions are NOT "
                 "collapsed; they stay at item level, because summing 500mg "
                 "and 850mg tablets gives a meaningless number.",
                 "Every figure is computed from the merged export at build time - "
                 "both Kobo servers, post facility-reattribution."])

    # ---- 2. geography
    sheet(wb, "By geography",
          [("Province", HDR, by(lambda r: s(r["_p"][M.PROV])), ALL, PCTCOLS, NUM0),
           ("District", HDR, by(lambda r: s(r["_p"][M.DIST])), ALL, PCTCOLS, NUM0),
           ("Sector", HDR, by(lambda r: s(r["_p"][M.SECT])), None, PCTCOLS, NUM0)],
          notes=[NOTE_DENOM, NOTE_THIN,
                 "Sorted worst availability first, with thin groups pushed to the "
                 "bottom. Sector is the finest geography the form carries."])

    # ---- 3. facility characteristics
    sheet(wb, "By facility type",
          [("Level of care (grouped - the reporting tiers)", HDR,
            by(lambda r: FL.group(r["_p"][M.LEVEL]),
               sort="name"), ALL, PCTCOLS, NUM0),
           ("Health facility level exactly as recorded", HDR,
            by(lambda r: s(r["_p"][M.LEVEL])), None, PCTCOLS, NUM0),
           ("Facility class (short code)", HDR, by(lambda r: s(r["_p"][M.FCLASS])),
            None, PCTCOLS, NUM0),
           ("Level rank (1 = health centre ... 3 = referral)", HDR,
            by(lambda r: s(r["_p"][M.RANK]), sort="name"), None, PCTCOLS, NUM0),
           ("Facility management", HDR, by(lambda r: s(r["_p"][M.MGMT])),
            None, PCTCOLS, NUM0),
           ("Kobo server (collection team)", HDR,
            by(lambda r: s(r["_p"].get("_server"))), None, PCTCOLS, NUM0)],
          notes=[NOTE_DENOM, NOTE_THIN,
                 "The server split is a data-collection artefact, not a health "
                 "system variable - it is here only so a difference between the "
                 "two teams' data can be seen if there is one."])

    # ---- 4. medicine characteristics
    sheet(wb, "By medicine",
          [("Programme", HDR, by(lambda r: s(r[M.PROGRAM])), ALL, PCTCOLS, NUM0),
           ("Drug class", HDR, by(lambda r: s(r[M.DCLASS])), None, PCTCOLS, NUM0),
           ("Applicable to the facility level per the EML", HDR,
            by(lambda r: s(r[M.APPLICABLE]) or "(blank)"), None, PCTCOLS, NUM0),
           ("Individual medicine", HDR, by(lambda r: s(r[M.MED])),
            None, PCTCOLS, NUM0)],
          notes=[NOTE_DENOM, NOTE_THIN,
                 "The individual-medicine block is the procurement view: 81 "
                 "medicines, worst availability first."])

    # ---- 5. recording system
    sheet(wb, "By stock system",
          [("Electronic system named for stock recording", HDR,
            by(lambda r: emr(r.get(M.EMR))), None, PCTCOLS, NUM0)],
          notes=[NOTE_DENOM, NOTE_THIN,
                 "Free text, so spellings are normalised: ELMIS / eLMIS / E-LMIS "
                 "are one system. Anything unrecognised is shown as entered.",
                 "This is what the facility said it records stock in - read it as "
                 "a correlate, not a cause: system choice tracks facility level "
                 "and programme."])

    # ---- 6. crosstabs
    provs = sorted({s(r["_p"][M.PROV]) for r in MEDROWS_M3})
    levels = sorted({s(r["_p"][M.LEVEL]) for r in MEDROWS_M3})

    def cross(rowkey, colvals, colkey, rowvals):
        out = []
        for rv in rowvals:
            row = [rv]
            for cv in colvals:
                sel = [r for r in MEDROWS_M3 if rowkey(r) == rv and colkey(r) == cv]
                ans = [r for r in sel if M.yes(r[M.MANAGED]) and not M.blank(r[M.AVAIL])]
                av = sum(1 for r in ans if M.yes(r[M.AVAIL]))
                row += [pct(av, len(ans)), len(ans)]
            out.append(row)
        return out

    prog = [p for p, _ in Counter(s(r[M.PROGRAM]) for r in MEDROWS_M3).most_common()]
    hdr_p = ["Programme"] + [x for p in provs for x in (f"{p} avail %", f"{p} n")]
    hdr_l = ["Programme"] + [x for l in levels for x in (f"{l[:18]} avail %", f"{l[:18]} n")]
    sheet(wb, "Crosstabs",
          [("Availability % by programme and province", hdr_p,
            cross(lambda r: s(r[M.PROGRAM]), provs,
                  lambda r: s(r["_p"][M.PROV]), prog), None,
            tuple(range(1, len(hdr_p), 2)), tuple(range(2, len(hdr_p), 2))),
           ("Availability % by programme and facility level", hdr_l,
            cross(lambda r: s(r[M.PROGRAM]), levels,
                  lambda r: s(r["_p"][M.LEVEL]), prog), None,
            tuple(range(1, len(hdr_l), 2)), tuple(range(2, len(hdr_l), 2)))],
          notes=[NOTE_DENOM,
                 "Each cell carries its own denominator in the 'n' column beside "
                 "it. A percentage on a single-digit n is not a finding."])

    # ---- 7. every stock-out
    out_rows = [r for r in MEDROWS_M3 if M.no(r[M.AVAIL])]
    rows = [[s(r["_p"][M.FAC]), s(r["_p"][M.DIST]), s(r["_p"][M.LEVEL]),
             s(r[M.MED]), s(r[M.PROGRAM]), s(r[M.DCLASS]),
             round(M.num(r[M.AMC]) or 0), round(M.num(r[M.SO_TOTAL]) or 0),
             s(r[M.RMS_CEN]), s(r[M.RMS_REG]), s(r[M.DIST_FLAG]),
             round(M.num(r[M.LOSS]) or 0), s(r["_p"].get("_server"))]
            for r in out_rows]
    rows.sort(key=lambda x: -x[11])
    sheet(wb, "Every stock-out",
          [("Every observation recorded as not available on the visit day",
            ["Facility", "District", "Level", "Medicine", "Programme",
             "Drug class", "Avg monthly consumption", "Stock-out days Jun-Aug",
             "At RMS Central?", "At RMS Regional?", "Attribution",
             "Recorded loss (RWF)", "Server"], rows, None, (), (6, 7, 11))],
          notes=["Sorted by recorded loss. 'At RMS Central / Regional' is what "
                 "separates a distribution failure from a national stock-out - "
                 "available upstream but absent here is the former.",
                 "Only these rows carry the RMS and attribution questions; they "
                 "are blank everywhere else by skip logic."])

    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)      # Box resets synced files to read-only
    wb.save(OUT)

    print(f"wrote {OUT}")
    print(f"  {M.N_MED:,} observations | {ALL[2]} facilities | {ALL[3]} medicines")
    print(f"  funnel: {ALL[4]:,} applicable -> {ALL[5]:,} managed -> "
          f"{ALL[7]:,} answered -> {ALL[8]:,} available ({ALL[9]:.1%})")
    print(f"  {ALL[10]:,} stocked out | {ALL[21]:,} RWF recorded loss")
    worst = [b for b in by(lambda r: s(r[M.PROGRAM])) if b[22] != "thin"][:1]
    if worst:
        print(f"  weakest programme: {worst[0][0]} at {worst[0][9]:.1%} "
              f"(n={worst[0][7]:,})")


if __name__ == "__main__":
    build()
