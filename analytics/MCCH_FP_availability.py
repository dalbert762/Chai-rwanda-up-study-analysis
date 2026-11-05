"""Availability of MCCH and family-planning medicines, from the CLEANED data.

    python3 MCCH_FP_availability.py
      -> MCCH_FP_availability.xlsx

Reads "Data cleaning/01_Stock_assessment_clean.xlsx" (and the CHW file for the
community child-health sheet) directly, so every correction in the cleaning
log - EML applicability (S6), available-with-zero-quantity (S7), duplicates
(S2), reattributions, level grouping - is already in force. The earlier
MCCH_*.xlsx workbooks were built on the raw export and predate that.

THE BASKETS (the stock tool's own `medicine_program` field)
    MCCH  oxytocin, MgSO4, tranexamic acid, heat-stable carbetocin, caffeine
          citrate - five obstetric / newborn emergency injectables
    FP    DMPA (Depo-Provera), levonorgestrel implant (Jadelle), combined
          oral contraceptive (Microgynon), copper IUD

TWO DENOMINATORS, ALWAYS SIDE BY SIDE
    Available % of managed     available / managed-and-answered. The project
                               standard: is the shelf stocked where the
                               facility runs the product?
    Available % of applicable  available / applicable per EML. Analysis rule
                               A1: a product the facility is authorised to
                               hold but does not hold counts as unavailable.
The gap between the two is the "management gap", and for both baskets it is
the bigger story than stock-outs.

WHY "NOT MANAGED" IS CODED, NOT COUNTED
For FP, most not-managed rows are faith-based facilities that do not offer
modern contraception by policy - the FP analogue of MFT zoning in the CHW
tool. For MCCH they are mostly a roll-out gap for the three newer products.
The free-text reason is coded into themes (not_managed.py) so the two are not
reported as one number.

openpyxl + the standard library, like the rest of this project.
"""

import _bootstrap as B   # path setup - must come first
import not_managed as NM
import os
import re
import statistics
from collections import Counter, defaultdict

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

CLEAN = os.path.join(B.ROOT, "Data cleaning")
STOCK = os.path.join(CLEAN, "01_Stock_assessment_clean.xlsx")
CHW = os.path.join(CLEAN, "03_CHW_assessment_clean.xlsx")
OUT = B.out("MCCH_FP_availability.xlsx")
MIN_N = 30            # answered rows below which a rate is flagged as thin
FAITH = "Faith based facility based"

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="DDE5F0")
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")
BAD_FILL = PatternFill("solid", fgColor="FCE4E4")
OK_FILL = PatternFill("solid", fgColor="E7F3E7")
THIN_FILL = PatternFill("solid", fgColor="FFF4D6")
GREY_FILL = PatternFill("solid", fgColor="EDEDED")

# short name, clinical role - in the order the tables print
MEDS = {
    "MCCH": [
        ("Oxytocin (10 IU injectable)", "Oxytocin",
         "PPH prevention/treatment - first-line uterotonic"),
        ("MgSO4 (50% injectable)", "MgSO4",
         "severe pre-eclampsia / eclampsia"),
        ("Tranexamic Acid Injection", "Tranexamic acid",
         "PPH treatment (antifibrinolytic)"),
        ("Heat Stable Carbetocin Injection", "Carbetocin (HS)",
         "PPH prevention - heat-stable uterotonic"),
        ("Caffeine Citrate Injection", "Caffeine citrate",
         "apnoea of prematurity (newborn)"),
    ],
    "FP": [
        ("Medroxyprogesterone Acetate 150 mg Vial/Depo provera)", "DMPA (Depo)",
         "short-acting - injectable"),
        ("Levonorgestrel/Ethinyl Estradiol 150/30 mcg + Fe 75 mg, 28 "
         "Tablets/Cycle/Microgynon", "COC (Microgynon)",
         "short-acting - oral pill"),
        ("Levonorgestrel 75mg/rod, 2rod Implant+Inserter/Jadelle",
         "Implant (Jadelle)", "long-acting reversible (LARC)"),
        ("Copper-bearing intrauterine devices (IUD)", "Copper IUD",
         "long-acting reversible (LARC)"),
    ],
}
SHORT = {full: sh for prog in MEDS.values() for full, sh, _ in prog}
ROLE = {full: role for prog in MEDS.values() for full, _, role in prog}
ORDER = {full: i for prog in MEDS.values() for i, (full, _, _) in enumerate(prog)}

LEVEL_ORDER = ["Health post", "Health centre", "District hospital",
               "Provincial hospital", "Level 2 teaching hospital",
               "National referral hospital"]
MGMT_LABEL = {"Government": "Government", FAITH: "Faith-based",
              "PPP": "PPP", "": "(not recorded)"}


# ---------------------------------------------------------------
# free-text "why not managed" -> theme: the shared coder, so this workbook
# and the dashboard code every answer the same way.
# ---------------------------------------------------------------
NO_REASON = NM.NO_REASON
reason_theme = NM.theme


# ---------------------------------------------------------------
# helpers
# ---------------------------------------------------------------
def s(v):
    return "" if v is None else str(v).strip()


def num(v):
    """Numeric value or None; rejects the literal 'NaN' Kobo calculates emit."""
    if v is None or v == "":
        return None
    try:
        x = float(v)
    except (TypeError, ValueError):
        return None
    return x if x == x and x not in (float("inf"), float("-inf")) else None


def yes(v):
    return s(v).lower() == "yes"


def no(v):
    return s(v).lower() == "no"


def pct(a, b):
    return a / b if b else None


def median(xs):
    xs = [x for x in xs if x is not None]
    return statistics.median(xs) if xs else None


def col(header, *needles, exclude=()):
    """The one column whose name contains every needle - fails loudly, because
    a stale constant otherwise reads as an empty column (see project notes)."""
    hits = [c for c in header if c and all(n in c for n in needles)
            and not any(x in c for x in exclude)]
    if len(hits) != 1:
        raise KeyError(f"{needles}: expected 1 column, found {hits}")
    return hits[0]


def read_sheet(path, sheet):
    ws = load_workbook(path, read_only=True)[sheet]
    it = ws.iter_rows(values_only=True)
    header = list(next(it))
    return header, [dict(zip(header, r)) for r in it]


# ---------------------------------------------------------------
# load
# ---------------------------------------------------------------
HDR, ALL = read_sheet(STOCK, "medicine_rows")
APPLIC = col(HDR, "applicable to this facility level")
MANAGED = col(HDR, "managed by the facility?", exclude=("Why",))
WHY = col(HDR, "Why is", "not managed")
AVAIL = col(HDR, "available on visit day")
QTY = col(HDR, "Quantity of", "available", exclude=("system",))
QTY_SYS = col(HDR, "available in the system")
SO_NOW = col(HDR, "Current stock-out days")
SO_MONTHS = [col(HDR, f"{m} 2026 stock-out days") for m in ("June", "July", "August")]
RMS_C = col(HDR, "available at RMS Central")
RMS_R = col(HDR, "available at RMS Regional")
CONS = [col(HDR, f"{m} consumption total") for m in ("June", "July", "August")]

ROWS = [r for r in ALL if s(r["medicine_program"]) in MEDS]
for r in ROWS:
    r["_prog"] = s(r["medicine_program"])
    r["_med"] = s(r["medicine_name"])
    r["_short"] = SHORT.get(r["_med"], r["_med"])
    r["_mgmt"] = MGMT_LABEL.get(s(r["management"]), s(r["management"]))
    r["_lvl"] = s(r["level_group"]) or "(not recorded)"
    r["_app"] = yes(r[APPLIC])
    r["_mg"] = yes(r[MANAGED])
    r["_ans"] = r["_mg"] and (yes(r[AVAIL]) or no(r[AVAIL]))
    r["_av"] = r["_ans"] and yes(r[AVAIL])
    r["_so"] = r["_ans"] and no(r[AVAIL])
    r["_reason"] = reason_theme(r[WHY]) if no(r[MANAGED]) else ""
    so = [num(r[c]) for c in SO_MONTHS]
    r["_so_days"] = sum(x for x in so if x is not None) if any(
        x is not None for x in so) else None
    r["_mos"] = num(r["stock_level_months"])
    r["_amc"] = num(r["avg_monthly_consumption"])
missing = [m for m in ORDER if m not in {x["_med"] for x in ROWS}]
if missing:
    raise KeyError(f"medicines not found in the cleaned data: {missing}")

FACS = sorted({s(r["facility"]) for r in ROWS})


def status(r):
    """One facility x medicine outcome, mutually exclusive."""
    if r["_av"]:
        return "Available"
    if r["_so"]:
        return "Stocked out"
    if r["_mg"]:
        return "Managed - not answered"
    return "Not managed" + ("" if r["_app"] else " (not applicable)")


# ---------------------------------------------------------------
# the measure set
# ---------------------------------------------------------------
M_HDR = ["Group", "Observations", "Facilities", "Applicable per EML",
         "Managed", "Managed % of applicable", "Answered", "Available",
         "Stocked out", "AVAILABLE % OF MANAGED", "Available % of applicable",
         "Not managed though applicable", "of which faith-based policy",
         "Thin?"]
M_PCT = (5, 9, 10)


def measure(label, rows):
    app = [r for r in rows if r["_app"]]
    mg = [r for r in rows if r["_mg"]]
    ans = [r for r in mg if r["_ans"]]
    av = [r for r in ans if r["_av"]]
    so = [r for r in ans if r["_so"]]
    gap = [r for r in app if not r["_mg"]]
    faith = [r for r in gap if r["_reason"].startswith("Faith")]
    return [label, len(rows), len({s(r["facility"]) for r in rows}),
            len(app), len(mg), pct(len([r for r in app if r["_mg"]]), len(app)),
            len(ans), len(av), len(so), pct(len(av), len(ans)),
            pct(len([r for r in app if r["_av"]]), len(app)),
            len(gap), len(faith), "thin" if len(ans) < MIN_N else ""]


S_HDR = ["Group", "Managed rows", "Stock out", "Understock (<1 mo)",
         "Plan stock (1-2 mo)", "Overstock (>2 mo)",
         "No consumption (MOS undefined)", "Median months of stock",
         "Rows with Jun-Aug stock-out days answered",
         "Rows with ANY stock-out day Jun-Aug", "% with a stock-out Jun-Aug",
         "Mean stock-out days (of those with one)", "Total stock-out days",
         "Qty requested Jun-Aug", "Qty delivered Jun-Aug", "Fulfilment %",
         "Lines short-delivered",
         "Kobo financial_loss (RWF) - mostly undelivered qty x tariff"]
S_PCT = (10, 15)


def stock_measure(label, rows):
    mg = [r for r in rows if r["_mg"]]
    cls = Counter(s(r["stock_classification"]) for r in mg)
    so_ans = [r for r in mg if r["_so_days"] is not None]
    so_any = [r for r in so_ans if r["_so_days"] > 0]
    req = [num(r["total_requested_qty"]) for r in mg]
    dlv = [num(r["total_delivered_qty"]) for r in mg]
    short = sum(1 for a, b in zip(req, dlv)
                if a is not None and b is not None and b < a)
    rq = sum(x for x in req if x)
    dv = sum(x for x in dlv if x)
    loss = sum(num(r["financial_loss"]) or 0 for r in mg)
    return [label, len(mg), cls["Stock out"], cls["Understock"],
            cls["Plan stock"], cls["Overstock"], cls[""],
            median([r["_mos"] for r in mg]), len(so_ans), len(so_any),
            pct(len(so_any), len(so_ans)),
            (statistics.mean(r["_so_days"] for r in so_any) if so_any else None),
            sum(r["_so_days"] for r in so_ans), rq, dv, pct(dv, rq), short,
            loss]


def grouped(rows, key, order=None, total=None, fn=measure):
    g = defaultdict(list)
    for r in rows:
        g[key(r)].append(r)
    keys = order and [k for k in order if k in g] + sorted(
        k for k in g if k not in order) or sorted(g)
    out = [fn(k, g[k]) for k in keys]
    if total:
        out.append(fn(total, rows))
    return out


def prog_rows(p):
    return [r for r in ROWS if r["_prog"] == p]


def by_med(p, fn=measure):
    rows = prog_rows(p)
    out = [fn(f"{sh}", [r for r in rows if r["_med"] == full])
           for full, sh, _ in MEDS[p]]
    out.append(fn(f"ALL {p}", rows))
    return out


# ---------------------------------------------------------------
# facility-level indicators
# ---------------------------------------------------------------
FAC = defaultdict(dict)        # facility -> short name -> row
FAC_META = {}
for r in ROWS:
    f = s(r["facility"])
    FAC[f][r["_short"]] = r
    FAC_META[f] = (s(r["province"]), s(r["district"]), r["_lvl"], r["_mgmt"])


def fac_indicator(label, eligible, success, note=""):
    """eligible(facility rows) -> bool; success(facility rows) -> bool."""
    el = [f for f in FAC if eligible(FAC[f])]
    ok = [f for f in el if success(FAC[f])]
    return [label, len(el), len(ok), pct(len(ok), len(el)), note]


def av(d, sh):
    return sh in d and d[sh]["_av"]


def app(d, sh):
    return sh in d and d[sh]["_app"]


def mg(d, sh):
    return sh in d and d[sh]["_mg"]


FP_SH = [sh for _, sh, _ in MEDS["FP"]]
MC_SH = [sh for _, sh, _ in MEDS["MCCH"]]


def offers_fp(d):
    return any(mg(d, sh) for sh in FP_SH)


def n_fp_av(d):
    return sum(av(d, sh) for sh in FP_SH)


def mcch_indicators(rows_filter=lambda f: True):
    keep = {f for f in FAC if rows_filter(f)}

    def ind(label, el, ok, note=""):
        e = [f for f in keep if el(FAC[f])]
        o = [f for f in e if ok(FAC[f])]
        return [label, len(e), len(o), pct(len(o), len(e)), note]
    return [
        ind("Oxytocin available", lambda d: app(d, "Oxytocin"),
            lambda d: av(d, "Oxytocin"),
            "of facilities where oxytocin is EML-applicable"),
        ind("MgSO4 available", lambda d: app(d, "MgSO4"),
            lambda d: av(d, "MgSO4"),
            "of facilities where MgSO4 is EML-applicable"),
        ind("Oxytocin AND MgSO4 both available",
            lambda d: app(d, "Oxytocin") and app(d, "MgSO4"),
            lambda d: av(d, "Oxytocin") and av(d, "MgSO4"),
            "the two UN life-saving maternal commodities in this basket"),
        ind("Any uterotonic available (oxytocin or carbetocin)",
            lambda d: app(d, "Oxytocin") or app(d, "Carbetocin (HS)"),
            lambda d: av(d, "Oxytocin") or av(d, "Carbetocin (HS)"),
            "PPH prevention possible on the day"),
        ind("Full PPH set: uterotonic + tranexamic acid",
            lambda d: app(d, "Tranexamic acid"),
            lambda d: (av(d, "Oxytocin") or av(d, "Carbetocin (HS)"))
            and av(d, "Tranexamic acid"),
            "WHO 2023 PPH bundle needs both"),
        ind("All MCCH items applicable to the facility available",
            lambda d: any(app(d, sh) for sh in MC_SH),
            lambda d: all(av(d, sh) for sh in MC_SH if app(d, sh)),
            "applicable per (cleaned) EML"),
        ind("All MCCH items the facility MANAGES available",
            lambda d: any(mg(d, sh) for sh in MC_SH),
            lambda d: all(av(d, sh) for sh in MC_SH if mg(d, sh)),
            "no stock-out among what it holds"),
    ]


def fp_indicators(rows_filter=lambda f: True):
    keep = {f for f in FAC if rows_filter(f)}

    def ind(label, el, ok, note=""):
        e = [f for f in keep if el(FAC[f])]
        o = [f for f in e if ok(FAC[f])]
        return [label, len(e), len(o), pct(len(o), len(e)), note]
    anyfp = lambda d: any(sh in d for sh in FP_SH)
    return [
        ind("Offers FP (manages at least one method)", anyfp, offers_fp,
            "of all facilities assessed for FP"),
        ind("At least one method available", offers_fp,
            lambda d: n_fp_av(d) >= 1, "of facilities offering FP"),
        ind("A short-acting method available (DMPA or COC)", offers_fp,
            lambda d: av(d, "DMPA (Depo)") or av(d, "COC (Microgynon)"),
            "of facilities offering FP"),
        ind("A LARC available (implant or IUD)", offers_fp,
            lambda d: av(d, "Implant (Jadelle)") or av(d, "Copper IUD"),
            "of facilities offering FP"),
        ind(">= 3 of the 4 tracked methods available", offers_fp,
            lambda d: n_fp_av(d) >= 3,
            "UNFPA SDP 'method choice' indicator, primary level"),
        ind("All 4 tracked methods available", offers_fp,
            lambda d: n_fp_av(d) == 4,
            "proxy for the >= 5 methods expected at hospital level"),
        ind("Every method it manages available (no FP stock-out)", offers_fp,
            lambda d: all(av(d, sh) for sh in FP_SH if mg(d, sh)),
            "of facilities offering FP"),
    ]


# ---------------------------------------------------------------
# workbook writer
# ---------------------------------------------------------------
wb = Workbook()
wb.remove(wb.active)


def fmt(cell, v, is_pct):
    if isinstance(v, float) and is_pct:
        cell.number_format = "0.0%"
    elif isinstance(v, float):
        cell.number_format = "#,##0.0" if abs(v) < 1000 else "#,##0"
    elif isinstance(v, int):
        cell.number_format = "#,##0"


def sheet(name, blocks, notes=(), widths=None):
    """blocks: (title, header, rows, pct_cols, avail_col) - avail_col (1-based)
    is shaded red/green, thin rows amber."""
    ws = wb.create_sheet(name)
    row = 1
    for n in notes:
        c = ws.cell(row, 1, n)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=10)
        ws.row_dimensions[row].height = 15 * max(1, len(n) // 140 + 1)
        row += 1
    if notes:
        row += 1
    for title, header, rows, pcols, acol in blocks:
        if title:
            ws.cell(row, 1, title).font = Font(bold=True, size=12)
            row += 1
        for j, h in enumerate(header, 1):
            c = ws.cell(row, j, h)
            c.font = Font(bold=True)
            c.fill = HDR_FILL
            c.border = BORDER
            c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[row].height = 45
        row += 1
        for vals in rows:
            total = s(vals[0]).startswith(("ALL", "TOTAL", "All facilities"))
            thin = "thin" in [s(v) for v in vals]
            for j, v in enumerate(vals, 1):
                c = ws.cell(row, j, v)
                c.border = BORDER
                fmt(c, v, (j - 1) in pcols)
                if total:
                    c.fill = TOT_FILL
                    c.font = Font(bold=True)
                elif thin:
                    c.fill = THIN_FILL
            if acol and isinstance(vals[acol - 1], float) and not total and not thin:
                v = vals[acol - 1]
                ws.cell(row, acol).fill = (BAD_FILL if v < 0.8 else
                                           OK_FILL if v >= 0.95 else
                                           PatternFill())
            row += 1
        row += 2
    ws.column_dimensions["A"].width = 42
    for j in range(2, 30):
        ws.column_dimensions[get_column_letter(j)].width = 13
    for j, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = None
    return ws


DEF_NOTES = (
    "AVAILABLE % OF MANAGED = available on visit day / rows the facility "
    "manages and answered (project standard). 'Available % of applicable' "
    "counts an EML-applicable product the facility does not manage as "
    "unavailable (analysis rule A1). Their difference is the management gap.",
    f"Amber rows rest on fewer than {MIN_N} answered observations - read, do "
    "not compare. Red < 80%, green >= 95%. Source: Data cleaning/"
    "01_Stock_assessment_clean.xlsx (all cleaning rules S1-S8 applied).",
)


# ---------------------------------------------------------------
# 1. Summary
# ---------------------------------------------------------------
def headline(p):
    rows = prog_rows(p)
    m = measure(p, rows)
    st = stock_measure(p, rows)
    return m, st


summ = []
for p in MEDS:
    m, st = headline(p)
    summ.append([f"{p} - facilities assessed", m[2]])
    summ.append([f"{p} - facility x medicine observations", m[1]])
    summ.append([f"{p} - managed % of EML-applicable", m[5]])
    summ.append([f"{p} - AVAILABLE % of managed (shelf)", m[9]])
    summ.append([f"{p} - available % of applicable (access)", m[10]])
    summ.append([f"{p} - managed rows with a stock-out day Jun-Aug", st[10]])
    summ.append([f"{p} - median months of stock (managed)", st[7]])
    summ.append(["", None])

# FP: government vs faith-based is the headline cut
fp_gov = [r for r in prog_rows("FP") if r["_mgmt"] != "Faith-based"]
fp_fb = [r for r in prog_rows("FP") if r["_mgmt"] == "Faith-based"]
mi = {x[0]: x for x in mcch_indicators()}
fi = {x[0]: x for x in fp_indicators()}
key_ind = [
    ["MCCH: oxytocin AND MgSO4 both available",
     *mi["Oxytocin AND MgSO4 both available"][1:4]],
    ["MCCH: full PPH set (uterotonic + TXA) available",
     *mi["Full PPH set: uterotonic + tranexamic acid"][1:4]],
    ["MCCH: every applicable item available",
     *mi["All MCCH items applicable to the facility available"][1:4]],
    ["FP: facilities offering FP",
     *fi["Offers FP (manages at least one method)"][1:4]],
    ["FP: >= 3 methods available (of facilities offering FP)",
     *fi[">= 3 of the 4 tracked methods available"][1:4]],
    ["FP: all 4 methods available (of facilities offering FP)",
     *fi["All 4 tracked methods available"][1:4]],
]
fp_split = [measure("Non-faith-based (Government, PPP)", fp_gov),
            measure("Faith-based", fp_fb),
            measure("ALL FP", prog_rows("FP"))]

sheet("Summary", [
    ("Headline figures", ["Measure", "Value"], summ, (), None),
    ("Facility-level indicators", ["Indicator", "Eligible facilities",
                                   "Meeting it", "%"], key_ind, (3,), 4),
    ("FP depends on who runs the facility", M_HDR, fp_split, M_PCT, 10),
], notes=(
    "MCCH and family-planning medicine availability - UP study stock "
    "assessment, cleaned dataset. MCCH = five obstetric/newborn emergency "
    "injectables; FP = four contraceptive methods. See 'Read me' for "
    "definitions.",) + DEF_NOTES, widths={1: 58, 2: 14})

# ---------------------------------------------------------------
# 2. Per medicine (both baskets)
# ---------------------------------------------------------------
med_blocks = []
for p in MEDS:
    rows = by_med(p)
    for x, (full, sh, role) in zip(rows, MEDS[p]):
        x.insert(1, role)
    rows[-1].insert(1, "")
    med_blocks.append((f"{p} - availability funnel by medicine",
                       [M_HDR[0], "Clinical role"] + M_HDR[1:], rows,
                       tuple(c + 1 for c in M_PCT), 11))
sheet("By medicine", med_blocks, notes=DEF_NOTES, widths={2: 34})

# ---------------------------------------------------------------
# 3. Facility-level indicators, with cuts
# ---------------------------------------------------------------
IND_HDR = ["Indicator", "Eligible facilities", "Meeting it", "%", "Basis"]
blocks = [("MCCH - all facilities", IND_HDR, mcch_indicators(), (3,), 4),
          ("FP - all facilities", IND_HDR, fp_indicators(), (3,), 4),
          ("FP - excluding faith-based facilities", IND_HDR,
           fp_indicators(lambda f: FAC_META[f][3] != "Faith-based"), (3,), 4),
          ("FP - faith-based facilities only", IND_HDR,
           fp_indicators(lambda f: FAC_META[f][3] == "Faith-based"), (3,), 4)]
for lvl in LEVEL_ORDER:
    if sum(1 for f in FAC if FAC_META[f][2] == lvl) >= 5:
        blocks.append((f"MCCH - {lvl}", IND_HDR,
                       mcch_indicators(lambda f, l=lvl: FAC_META[f][2] == l),
                       (3,), 4))
        blocks.append((f"FP - {lvl}", IND_HDR,
                       fp_indicators(lambda f, l=lvl: FAC_META[f][2] == l),
                       (3,), 4))
sheet("Facility indicators", blocks, notes=(
    "Unit = facility, not observation. A facility is eligible for an MCCH "
    "indicator if the product(s) are EML-applicable at its level (cleaned "
    "applicability); for the FP method-mix indicators, if it manages at "
    "least one FP method - a facility that offers no FP cannot have a method "
    "mix. Only 4 FP methods are tracked, so 'all 4' is a floor for the "
    "hospital-level >= 5 methods standard, not the standard itself.",
    "Levels with fewer than 5 facilities are not split out."),
    widths={1: 52, 5: 52})

# FP methods-available distribution
dist_rows = []
for lab, flt in (("All offering FP", lambda f: True),
                 ("Non-faith-based", lambda f: FAC_META[f][3] != "Faith-based"),
                 ("Faith-based", lambda f: FAC_META[f][3] == "Faith-based")):
    fs = [f for f in FAC if flt(f) and offers_fp(FAC[f])]
    c = Counter(n_fp_av(FAC[f]) for f in fs)
    dist_rows.append([lab, len(fs)] + [c[k] for k in range(5)] +
                     [pct(c[k], len(fs)) for k in range(5)])
ws = wb["Facility indicators"]
start = ws.max_row + 3
ws.cell(start, 1, "FP - number of the 4 methods available on visit day "
        "(facilities offering FP)").font = Font(bold=True, size=12)
h = ["Group", "Facilities"] + [f"{k} methods" for k in range(5)] + \
    [f"{k} methods %" for k in range(5)]
for j, v in enumerate(h, 1):
    c = ws.cell(start + 1, j, v)
    c.font = Font(bold=True)
    c.fill = HDR_FILL
    c.border = BORDER
    c.alignment = Alignment(wrap_text=True)
for i, vals in enumerate(dist_rows):
    for j, v in enumerate(vals, 1):
        c = ws.cell(start + 2 + i, j, v)
        c.border = BORDER
        fmt(c, v, j >= 8)

# ---------------------------------------------------------------
# 4. Why not managed
# ---------------------------------------------------------------
reason_blocks = []
for p in MEDS:
    rows = [r for r in prog_rows(p) if no(r[MANAGED])]
    themes = [t for t, _ in Counter(r["_reason"] for r in rows).most_common()]
    shorts = [sh for _, sh, _ in MEDS[p]]
    tab = []
    for t in themes:
        line = [t] + [sum(1 for r in rows if r["_reason"] == t
                          and r["_short"] == sh) for sh in shorts]
        n = sum(line[1:])
        tab.append(line + [n, pct(n, len(rows))])
    tab.append(["ALL not-managed rows"] +
               [sum(1 for r in rows if r["_short"] == sh) for sh in shorts] +
               [len(rows), 1.0])
    reason_blocks.append((f"{p} - why the facility does not manage the product "
                          "(coded from free text)", ["Reason theme"] + shorts +
                          ["Rows", "Share"], tab, (len(shorts) + 2,), None))
    # applicable only
    rows_a = [r for r in rows if r["_app"]]
    tab = []
    for t in themes:
        line = [t] + [sum(1 for r in rows_a if r["_reason"] == t
                          and r["_short"] == sh) for sh in shorts]
        n = sum(line[1:])
        if n:
            tab.append(line + [n, pct(n, len(rows_a))])
    tab.append(["ALL applicable-but-not-managed"] +
               [sum(1 for r in rows_a if r["_short"] == sh) for sh in shorts] +
               [len(rows_a), 1.0])
    reason_blocks.append((f"{p} - same, EML-applicable rows only (the "
                          "management gap)", ["Reason theme"] + shorts +
                          ["Rows", "Share"], tab, (len(shorts) + 2,), None))

# management type x not managed
mg_tab = []
for p in MEDS:
    for mgmt in ("Government", "Faith-based", "PPP"):
        rs = [r for r in prog_rows(p) if r["_mgmt"] == mgmt]
        nm = [r for r in rs if no(r[MANAGED])]
        mg_tab.append([f"{p} - {mgmt}", len({s(r['facility']) for r in rs}),
                       len(rs), len(nm), pct(len(nm), len(rs)),
                       sum(1 for r in nm if r["_reason"].startswith("Faith"))])
reason_blocks.append(("Not-managed share by facility management",
                      ["Programme - management", "Facilities", "Rows",
                       "Not managed", "Not managed %",
                       "of which faith-based reason"], mg_tab, (4,), None))
sheet("Why not managed", reason_blocks, notes=(
    "'Why is <medicine> not managed?' is free text. It is coded here into "
    "themes by keyword (not_managed.py, shared with the dashboard); first matching theme "
    "wins. Every verbatim answer and its theme is on the 'Not-managed "
    "verbatims' sheet so the coding can be checked.",
    "Most blank reasons are on rows that are NOT EML-applicable - the "
    "collector moved on because the product does not belong at that level. "
    "The second table in each pair keeps applicable rows only."),
    widths={1: 56})

# verbatims
verb = Counter((r["_prog"], r["_short"], s(r[WHY]), r["_reason"],
                "Yes" if r["_app"] else "No", r["_mgmt"])
               for r in ROWS if no(r[MANAGED]) and s(r[WHY]))
sheet("Not-managed verbatims", [(None, ["Programme", "Medicine",
                                        "Reason as recorded", "Coded theme",
                                        "EML-applicable?", "Management",
                                        "Rows"],
                                 [list(k) + [v] for k, v in
                                  sorted(verb.items(),
                                         key=lambda kv: (kv[0][0], kv[0][3],
                                                         kv[0][1], -kv[1]))],
                                 (), None)],
      widths={1: 10, 2: 18, 3: 70, 4: 46, 5: 11, 6: 13, 7: 8})

# ---------------------------------------------------------------
# 5. Disaggregations
# ---------------------------------------------------------------
for p in MEDS:
    rows = prog_rows(p)
    blocks = [
        (f"{p} - by level of care", M_HDR,
         grouped(rows, lambda r: r["_lvl"], LEVEL_ORDER, f"ALL {p}"), M_PCT, 10),
        (f"{p} - by facility management", M_HDR,
         grouped(rows, lambda r: r["_mgmt"], ["Government", "Faith-based",
                                              "PPP"], f"ALL {p}"), M_PCT, 10),
        (f"{p} - by province", M_HDR,
         grouped(rows, lambda r: s(r["province"]), None, f"ALL {p}"),
         M_PCT, 10),
    ]
    # medicine x level: available % of managed and managed % of applicable
    lv = [l for l in LEVEL_ORDER if any(r["_lvl"] == l for r in rows)]
    for metric, idx in (("AVAILABLE % OF MANAGED", 9),
                        ("Managed % of applicable", 5),
                        ("Available % of applicable", 10)):
        tab = []
        for full, sh, _ in MEDS[p]:
            line = [sh]
            for l in lv:
                m = measure("", [r for r in rows if r["_med"] == full
                                 and r["_lvl"] == l])
                line.append(m[idx] if (m[6] if idx == 9 else m[3]) >= 5 else None)
            tab.append(line)
        blocks.append((f"{p} - {metric.lower()} by medicine and level "
                       "(blank where < 5 rows)", ["Medicine"] + lv, tab,
                       tuple(range(1, len(lv) + 1)), None))
    if p == "FP":
        tab = []
        for full, sh, _ in MEDS[p]:
            for lab, flt in (("non-faith-based", lambda r: r["_mgmt"] != "Faith-based"),
                             ("faith-based", lambda r: r["_mgmt"] == "Faith-based")):
                tab.append(measure(f"{sh} - {lab}",
                                   [r for r in rows if r["_med"] == full and flt(r)]))
        blocks.insert(2, ("FP - by medicine and management", M_HDR, tab,
                          M_PCT, 10))
    sheet(f"{p} by level & mgmt", blocks, notes=DEF_NOTES)

    by_d = defaultdict(list)
    for r in rows:
        by_d[s(r["district"])].append(r)
    dist = [[s(g[0]["province"])] + measure(k, g) for k, g in
            sorted(by_d.items(), key=lambda kv: (s(kv[1][0]["province"]), kv[0]))]
    dist.append([""] + measure(f"ALL {p}", rows))
    sheet(f"{p} by district", [(f"{p} - by district", ["Province"] + M_HDR,
                                dist, tuple(c + 1 for c in M_PCT), 11)],
          notes=DEF_NOTES + ("District rates rest on a handful of facilities "
                             "each; most are thin. Use for targeting, not "
                             "ranking.",), widths={1: 14, 2: 18})

# ---------------------------------------------------------------
# 6. Stock status, stock-outs, supply
# ---------------------------------------------------------------
st_blocks = []
for p in MEDS:
    st_blocks.append((f"{p} - stock status and stock-outs by medicine "
                      "(managed rows)", S_HDR, by_med(p, stock_measure),
                      S_PCT, None))
    st_blocks.append((f"{p} - by level of care", S_HDR,
                      grouped(prog_rows(p), lambda r: r["_lvl"], LEVEL_ORDER,
                              f"ALL {p}", stock_measure), S_PCT, None))
sheet("Stock status & supply", st_blocks, notes=(
    "Managed rows only. Classification is the tool's own: months of stock = "
    "quantity on hand / average monthly consumption Jun-Aug; < 1 month "
    "understock, 1-2 plan stock, > 2 overstock. Where consumption is zero "
    "the ratio is undefined and the row is 'No consumption', never "
    "overstock.",
    "Stock-out days are the June, July and August answers summed (92 days "
    "possible). Requested / delivered are the Jun-Aug requisition totals; "
    "'short-delivered' = delivered < requested on that line.",
    "financial_loss is the tool's own calculate: (estimated qty unfilled "
    "during stock-out + undelivered requisition qty) x tariff. On these "
    "products it is driven almost entirely by undelivered requisitions, so "
    "read it as the value of supply not received, not as lost dispensing. "
    "FP products carry no tariff, hence zero."),
    widths={1: 30})

# RMS upstream (asked only on some rows)
rms = []
for p in MEDS:
    for full, sh, _ in MEDS[p]:
        rs = [r for r in prog_rows(p) if r["_med"] == full
              and (s(r[RMS_C]) or s(r[RMS_R]))]
        if not rs:
            continue
        so_rs = [r for r in rs if r["_so"]]
        rms.append([f"{p} - {sh}", len(rs),
                    sum(yes(r[RMS_C]) for r in rs), sum(no(r[RMS_C]) for r in rs),
                    sum(yes(r[RMS_R]) for r in rs), sum(no(r[RMS_R]) for r in rs),
                    len(so_rs), sum(yes(r[RMS_C]) or yes(r[RMS_R]) for r in so_rs)])
ws = wb["Stock status & supply"]
start = ws.max_row + 3
ws.cell(start, 1, "Upstream: was it available at RMS? (asked on a subset of "
        "rows)").font = Font(bold=True, size=12)
h = ["Medicine", "Rows answered", "RMS Central yes", "RMS Central no",
     "RMS Regional yes", "RMS Regional no", "of which facility stocked out",
     "Stocked out while RMS had it"]
for j, v in enumerate(h, 1):
    c = ws.cell(start + 1, j, v)
    c.font = Font(bold=True)
    c.fill = HDR_FILL
    c.border = BORDER
    c.alignment = Alignment(wrap_text=True)
for i, vals in enumerate(rms):
    for j, v in enumerate(vals, 1):
        ws.cell(start + 2 + i, j, v).border = BORDER

# ---------------------------------------------------------------
# 7. Inventory accuracy (facilities with a digital system - rule A3)
# ---------------------------------------------------------------
acc = []
for p in MEDS:
    for full, sh, _ in MEDS[p] + [(None, f"ALL {p}", None)]:
        rs = [r for r in prog_rows(p) if (full is None or r["_med"] == full)
              and r["_mg"] and s(r["recording_system_clean"])
              and s(r["recording_system_clean"]) != "(no system)"]
        pairs = [(num(r[QTY]), num(r[QTY_SYS])) for r in rs]
        pairs = [(a, b) for a, b in pairs if a is not None and b is not None]
        exact = sum(1 for a, b in pairs if a == b)
        within = sum(1 for a, b in pairs if max(a, b) == 0 or
                     abs(a - b) / max(a, b) <= 0.10)
        over = sum(1 for a, b in pairs if b > a)
        under = sum(1 for a, b in pairs if b < a)
        acc.append([f"{p} - {sh}" if full else sh, len(rs), len(pairs),
                    exact, pct(exact, len(pairs)), within,
                    pct(within, len(pairs)), over, under])
sheet("Inventory accuracy", [(None, ["Medicine", "Managed rows with a digital "
                                     "system", "Both counts recorded",
                                     "Physical = system", "Exact match %",
                                     "Within 10%", "Within 10% %",
                                     "System > physical", "System < physical"],
                              acc, (4, 6), 5)],
      notes=("Physical count on the shelf vs the quantity in the electronic "
             "system. Facilities that named no digital system are excluded "
             "(analysis rule A3). 'System > physical' is stock the system "
             "believes exists but the shelf does not hold - the direction "
             "that hides a stock-out.",), widths={1: 30})

# ---------------------------------------------------------------
# 8. What cleaning changed for these two baskets
# ---------------------------------------------------------------
cc = []
for p in MEDS:
    for full, sh, _ in MEDS[p]:
        rs = [r for r in prog_rows(p) if r["_med"] == full]
        app_fix = [r for r in rs if s(r["applicable_before_correction"])]
        av_fix = [r for r in rs if s(r["available_before_correction"])]
        before_app = sum(1 for r in rs if yes(r["applicable_before_correction"])
                         or (not s(r["applicable_before_correction"])
                             and r["_app"]))
        # availability over managed before S7: those rows were 'Yes'
        ans = [r for r in rs if r["_ans"]]
        before_av = sum(1 for r in ans if r["_av"]) + len(av_fix)
        cc.append([f"{p} - {sh}", len(rs), before_app, sum(r["_app"] for r in rs),
                   len(app_fix), len(av_fix),
                   pct(before_av, len(ans)), pct(sum(r["_av"] for r in ans), len(ans))])
sheet("Cleaning impact", [(None, ["Medicine", "Observations",
                                  "EML-applicable BEFORE (as answered)",
                                  "EML-applicable AFTER (S6)",
                                  "Applicability flags corrected",
                                  "Available=Yes with 0 units -> No (S7)",
                                  "Available % of managed BEFORE",
                                  "Available % of managed AFTER"],
                           cc, (6, 7), None)],
      notes=("S6 checked the collector's 'applicable per EML' answer against "
             "the Final EML list. For caffeine citrate it moved 136+ health-"
             "centre rows to applicable on an ingredient-only match ('Caffeine "
             "Citrate - Inject. All strength', listed at HC). If the EML's HC "
             "listing is not meant to cover neonatal injectables at HC, the "
             "BEFORE column is the better denominator for that product - "
             "worth confirming with the MCCH programme.",
             "S7 recoded 'available = Yes' with 0 units on the shelf to No."),
      widths={1: 30})

# ---------------------------------------------------------------
# 9. Facility matrix and every gap
# ---------------------------------------------------------------
CODE = {"Available": "A", "Stocked out": "SO", "Managed - not answered": "?",
        "Not managed": "NM", "Not managed (not applicable)": "n/a"}
mat = []
for f in sorted(FAC, key=lambda f: (FAC_META[f][0], FAC_META[f][1], f)):
    d = FAC[f]
    line = list(FAC_META[f][:2]) + [f, FAC_META[f][2], FAC_META[f][3]]
    line += [CODE[status(d[sh])] if sh in d else "" for sh in MC_SH + FP_SH]
    line += [sum(av(d, sh) for sh in MC_SH),
             sum(av(d, sh) for sh in FP_SH)]
    mat.append(line)
ws = sheet("Facility matrix", [(None, ["Province", "District", "Facility",
                                       "Level", "Management"] + MC_SH + FP_SH +
                                ["MCCH available (of 5)", "FP available (of 4)"],
                                mat, (), None)],
           notes=("A = available on visit day; SO = managed but stocked out; "
                  "NM = EML-applicable but not managed; n/a = not managed and "
                  "not applicable at this level; ? = managed, availability "
                  "not answered; blank = not assessed.",),
           widths={1: 12, 2: 13, 3: 30, 4: 18, 5: 13})
for row in ws.iter_rows(min_row=3, min_col=6, max_col=5 + len(MC_SH + FP_SH)):
    for c in row:
        if c.value == "SO":
            c.fill = BAD_FILL
        elif c.value == "NM":
            c.fill = THIN_FILL
        elif c.value == "A":
            c.fill = OK_FILL
        elif c.value == "n/a":
            c.fill = GREY_FILL
        c.alignment = Alignment(horizontal="center")

gaps = []
for r in sorted(ROWS, key=lambda r: (r["_prog"], ORDER[r["_med"]],
                                     s(r["province"]), s(r["district"]),
                                     s(r["facility"]))):
    if r["_av"] or not r["_app"] and not r["_mg"]:
        continue
    st = status(r)
    if st == "Available":
        continue
    gaps.append([r["_prog"], r["_short"], s(r["province"]), s(r["district"]),
                 s(r["facility"]), r["_lvl"], r["_mgmt"], st,
                 r["_reason"] if st.startswith("Not managed") else "",
                 s(r[WHY]) if st.startswith("Not managed") else "",
                 num(r[QTY]), r["_so_days"], num(r[SO_NOW]), r["_amc"]])
sheet("Every gap", [(None, ["Programme", "Medicine", "Province", "District",
                            "Facility", "Level", "Management", "Status",
                            "Reason theme", "Reason as recorded",
                            "Qty on shelf", "Stock-out days Jun-Aug",
                            "Current stock-out days", "AMC Jun-Aug"],
                     gaps, (), None)],
      notes=("Every facility x medicine where an EML-applicable product was "
             "not on the shelf: stocked out, or not managed. Filter by "
             "programme, district or status to build a follow-up list.",),
      widths={1: 9, 2: 16, 3: 11, 4: 12, 5: 28, 6: 16, 7: 12, 8: 14,
              9: 40, 10: 50})

# ---------------------------------------------------------------
# 10. Community child health - CHW tool
# ---------------------------------------------------------------
CH_ITEMS = [("ORS (alone)", "ORS"),
            ("Zinc 10mg dispersible tablet (alone)", "Zinc 10 mg"),
            ("Amox 125 mg tablet disp", "Amoxicillin 125 mg DT")]
ch_hdr, ch_all = read_sheet(CHW, "medicine_rows")
C_MG = col(ch_hdr, "managed by this CHW")
C_AV = col(ch_hdr, "available on visit day")
ch_tab = []
ch_lvl = []
for name, sh in CH_ITEMS:
    rs = [r for r in ch_all if s(r["medicine_name"]) == name]
    mgd = [r for r in rs if yes(r[C_MG])]
    ans = [r for r in mgd if yes(r[C_AV]) or no(r[C_AV])]
    avl = [r for r in ans if yes(r[C_AV])]
    ch_tab.append([sh, len(rs), len({s(r["facility"]) for r in rs}), len(mgd),
                   pct(len(mgd), len(rs)), len(ans), len(avl),
                   pct(len(avl), len(ans)),
                   median(num(r["stock_level_months"]) for r in mgd
                          if (num(r["avg_monthly_consumption"]) or 0) > 0)])
ch_sub = {s(r["_submission__uuid"]) for r in ch_all
          if s(r["medicine_name"]) in dict(CH_ITEMS)}
per_chw = defaultdict(dict)
for r in ch_all:
    if s(r["medicine_name"]) in dict(CH_ITEMS):
        per_chw[s(r["_submission__uuid"])][s(r["medicine_name"])] = r
all3 = [u for u, d in per_chw.items() if len(d) == 3
        and all(yes(x[C_MG]) for x in d.values())]
all3_av = [u for u in all3 if all(yes(x[C_AV]) for x in per_chw[u].values())]
sheet("CHW child health", [
    (None, ["Item", "CHW x item rows", "Parent facilities", "Managed",
            "Managed %", "Answered", "Available", "AVAILABLE % OF MANAGED",
            "Median months of stock (consuming CHWs)"], ch_tab, (4, 7), 8),
    (None, ["Indicator", "CHWs eligible", "Meeting it", "%"],
     [["ORS + zinc + amoxicillin all available (CHWs managing all three)",
       len(all3), len(all3_av), pct(len(all3_av), len(all3))]], (3,), 4)],
    notes=("Community child-health (iCCM) commodities from the CHW tool - "
           "a different instrument and unit (one CHW, not a facility). The "
           "stock tool's MCCH basket holds no child-health oral products, so "
           "this is the only child-health availability evidence in the "
           "study. Source: Data cleaning/03_CHW_assessment_clean.xlsx.",),
    widths={1: 60})

# ---------------------------------------------------------------
# 11. Read me
# ---------------------------------------------------------------
readme = [
    ["Source", "Data cleaning/01_Stock_assessment_clean.xlsx (medicine_rows), "
     "all cleaning rules applied; CHW sheet from 03_CHW_assessment_clean.xlsx."],
    ["Unit", "One facility x medicine observation, except on 'Facility "
     "indicators' (one facility) and 'CHW child health' (one CHW)."],
    ["Baskets", "From the tool's medicine_program field. MCCH: " +
     ", ".join(MC_SH) + ". FP: " + ", ".join(FP_SH) + "."],
    ["Drug class labels", "Not used. The tool records tranexamic acid as "
     "'Anti-apneic agent' and caffeine citrate as 'Hemostatic agent' - the "
     "labels are swapped - so medicines are named directly."],
    ["Available % of managed", "available on visit day / managed and "
     "answered. The shelf measure."],
    ["Available % of applicable", "available / EML-applicable (cleaned, S6). "
     "The access measure - counts the management gap as unavailable."],
    ["Faith-based FP", "Faith-based (mostly Catholic) facilities do not offer "
     "modern contraception by policy. Their not-managed FP rows are a "
     "service-design fact, not a supply failure, and are split out wherever "
     "FP is reported. Clients there are normally served by a nearby "
     "secondary health post, which this tool does not link."],
    ["Thin", f"Fewer than {MIN_N} answered rows: shaded amber, not "
     "comparable."],
    ["Not weighted", "Facilities follow a visit plan, not a probability "
     "sample; nothing here is a national estimate."],
    ["Rebuild", "python3 analytics/MCCH_FP_availability.py"],
]
sheet("Read me", [(None, ["Item", "Detail"], readme, (), None)],
      widths={1: 26, 2: 120})
for row in wb["Read me"].iter_rows(min_row=2):
    for c in row:
        c.alignment = Alignment(wrap_text=True, vertical="top")

wb.move_sheet("Read me", offset=-(len(wb.sheetnames) - 1))
wb.move_sheet("Summary", offset=-(len(wb.sheetnames) - 1))
wb.active = 0
if os.path.exists(OUT):
    os.chmod(OUT, 0o644)
wb.save(OUT)

# ---------------------------------------------------------------
# console summary
# ---------------------------------------------------------------
print(f"wrote {os.path.basename(OUT)}")
for p in MEDS:
    print(f"\n{p}")
    for x in by_med(p):
        print(f"  {x[0]:<18} obs {x[1]:>4}  managed%app {x[5] or 0:6.1%}  "
              f"avail%mgd {x[9] or 0:6.1%}  avail%app {x[10] or 0:6.1%}  "
              f"so {x[8]:>3}  gap {x[11]:>3} (faith {x[12]})")
    for x in by_med(p, stock_measure):
        print(f"  {x[0]:<18} SO-cls {x[2]:>3} under {x[3]:>3} plan {x[4]:>3} "
              f"over {x[5]:>3} nocons {x[6]:>3} MOS {x[7] or 0:5.1f}  "
              f"anySO {x[10] or 0:6.1%} meanSOd {x[11] or 0:5.1f} "
              f"fulfil {x[15] or 0:6.1%} loss {x[17]:,.0f}")
print("\nfacility indicators")
for x in mcch_indicators() + fp_indicators():
    print(f"  {x[0]:<55} {x[2]:>3}/{x[1]:<3} {x[3] or 0:6.1%}")
print("FP non-faith:")
for x in fp_indicators(lambda f: FAC_META[f][3] != "Faith-based"):
    print(f"  {x[0]:<55} {x[2]:>3}/{x[1]:<3} {x[3] or 0:6.1%}")
print("FP faith-based:")
for x in fp_indicators(lambda f: FAC_META[f][3] == "Faith-based"):
    print(f"  {x[0]:<55} {x[2]:>3}/{x[1]:<3} {x[3] or 0:6.1%}")
oth = Counter(s(r[WHY]) for r in ROWS
              if r["_reason"] == NM.OTHER)
print(f"\nreasons left uncoded: {sum(oth.values())}")
for k, v in oth.most_common(30):
    print(f"  {v:>3}  {k}")
