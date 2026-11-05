"""
Full analysis of the 'Medicine stock assessment for CHWs' KoboToolbox project.

    python3 fetch_kobo_export.py chw          # refresh the export first
    python3 analyze_medicine_assessment_chw.py

Reads   Medicine_stock_CHW_export.xlsx        (the synchronous XLSX export)
Writes  Medicine_stock_CHW_analysis.xlsx      (one worksheet per analysis section)

This is the community-health-worker companion to analyze_medicine_assessment.py.
The two forms share their availability / consumption questions, but NOT their
shape, so this is a separate script rather than a flag on that one:

  * the unit of observation is a CHW attached to a health centre, not the
    facility itself;
  * the basket is the 13-item CHW package (ACT/ALu x4, malaria RDT, Zinc, ORS,
    Amox 125, DHAP x3, Pyronaridine/Artesunate x2);
  * there is no EML applicability question, no system-vs-physical count, no
    stock-out day history, no requisition, no RMS availability and no unit
    price - so there is no inventory accuracy, fulfilment, attribution or
    financial-loss analysis to do. Anything this workbook does not contain is
    absent from the source, not dropped.

UNIT OF ANALYSIS
Rates are computed over CHW x medicine observations in the medicine_group
sheet, never over submissions - the basket is not the same size for every CHW
(the export shows 13 medicines for most and 5-12 for the rest), so a
submission-level rate would weight a 5-medicine CHW like a 13-medicine one.

THE DENOMINATOR THAT MATTERS
'Available on visit day' is asked only where the CHW manages the medicine, so
availability is computed over MANAGED rows alone. Most 'not managed' rows are
not a supply failure: RBC's MFT (multiple first-line therapy) zoning assigns
each area one of ALu / DHAP / Pyronaridine-Artesunate, so the other two are
correctly absent. The feedback coder below separates those from genuine gaps.

openpyxl + the Python standard library only (no pandas).
"""

import math
import os
import sys
import re
import statistics
from collections import Counter, defaultdict

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SRC = "Medicine_stock_CHW_export.xlsx"
OUT = "Medicine_stock_CHW_analysis.xlsx"
S_MED = "medicine_group"

# --- submission-level columns
PROV, DIST, SECT = "Province", "District", "Sector"
FAC = "Facility (select)"
LEVEL, FCLASS = "Health facility level", "facility_class"
CHW = "Name of CHW"
ENUM = "Please select your name"
BASKET = "Which medicines are being assessed at CHW? (Select all that apply)"
UNDER_T, OVER_T = "understock_threshold", "overstock_threshold"

# --- medicine_group columns (the form embeds ${medicine_name} in the labels)
MED = "medicine_name"
MANAGED = "Is **${medicine_name}** managed by this CHW?"
AVAIL = "Is **${medicine_name}** available on visit day?"
QTY = "Quantity of **${medicine_name}** available"
MONTHS = ("June", "July", "August")
CONS = {m: f"{m} consumption total quantity of **{{n}}**".replace("{n}", "${medicine_name}")
        for m in MONTHS}
AMC = "avg_monthly_consumption"
MOS = "stock_level_months"
FEEDBACK = ("Please share any feedback, concerns, or recommendations related to "
            "the supply chain and other relevant areas.")

# Short names for the 13-item basket. The full labels carry pack sizes and run
# to 70 characters, which no chart axis or crosstab column can show.
SHORT = [
    ("first presentation", "1x6", "ALu 1x6"),
    ("first presentation", "2x6", "ALu 2x6"),
    ("third presentation", "", "ALu 3x6"),
    ("fourth presentation", "", "ALu 4x6"),
    ("rapid diagnostic", "", "Malaria RDT"),
    ("zinc", "", "Zinc 10mg"),
    ("ors", "", "ORS"),
    ("amox", "", "Amox 125mg"),
    ("dhap", "20/160", "DHAP 20/160"),
    ("dhap", "30/240", "DHAP 30/240"),
    ("dhap", "80/640", "DHAP 80/640"),
    ("pyronaridine", "60/20", "Pyr/Art 60/20"),
    ("pyronaridine", "180/60", "Pyr/Art 180/60"),
]
# Which first-line antimalarial family a row belongs to, for the MFT-zoning read.
FAMILY = [("alu", "ACT / ALu"), ("dhap", "DHAP"), ("pyr", "Pyronaridine/Artesunate")]


def short(name):
    low = str(name).lower()
    for key, extra, out in SHORT:
        if key in low and (not extra or extra in low):
            return out
    return str(name)[:34]


def family(shortname):
    low = shortname.lower()
    for key, out in FAMILY:
        if low.startswith(key):
            return out
    return "Non-malaria (RDT / Zinc / ORS / Amox)"


# ---------------------------------------------------------------
# load
# ---------------------------------------------------------------
def load(sheet):
    wb = openpyxl.load_workbook(SRC, data_only=True, read_only=True)
    rows = list(wb[sheet].iter_rows(values_only=True))
    hdr = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = [dict(zip(hdr, r)) for r in rows[1:]
            if any(c is not None and str(c).strip() != "" for c in r)]
    return hdr, data


def blank(v):
    return v is None or str(v).strip() in ("", "None")


def num(v):
    if blank(v):
        return None
    try:
        f = float(str(v).strip())
    except ValueError:
        return None
    # Kobo calculate fields can emit the literal string "NaN" when a denominator
    # is zero or blank; float() accepts it and one such value turns every
    # downstream sum into NaN. Treat non-finite as missing, and count it in
    # Data quality so a new occurrence is visible rather than silent.
    return f if math.isfinite(f) else None


def yes(v):
    return str(v).strip().lower() == "yes"


def no(v):
    return str(v).strip().lower() == "no"


def s(v):
    return "" if v is None else str(v).strip()


def day(v):
    return s(v)[:10]


if not os.path.exists(SRC):
    raise SystemExit(f"{SRC} not found - run: python3 fetch_kobo_export.py chw")

wb_names = openpyxl.load_workbook(SRC, read_only=True).sheetnames
S_SUB = next(n for n in wb_names if n != S_MED)
HDR_SUB, SUB = load(S_SUB)
HDR_MED, MEDROWS = load(S_MED)

PARENT = {r["_index"]: r for r in SUB}
ORPHANS = sum(1 for r in MEDROWS if r["_parent_index"] not in PARENT)
for r in MEDROWS:
    r["_p"] = PARENT.get(r["_parent_index"])
    r["_short"] = short(r[MED])
    r["_family"] = family(r["_short"])
MEDROWS = [r for r in MEDROWS if r["_p"]]

# S10 - as in analyze_medicine_assessment.py: a first-line antimalarial the
# CHW manages, NOT available, and outside the district's MFT zone
# (analytics/mft_zoning.py) is recoded to NOT MANAGED (2026-11-02).
ZONE_RECODED = []
try:
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "analytics"))
    import mft_zoning as _ZN
except Exception:
    _ZN = None
if _ZN is not None:
    for r in MEDROWS:
        if (yes(r[MANAGED]) and no(r[AVAIL])
                and _ZN.out_of_zone(r["_short"], r["_p"][DIST])):
            r["_managed_was"], r["_available_was"] = s(r[MANAGED]), s(r[AVAIL])
            r[MANAGED], r[AVAIL] = "No", None
            # other answers not considered once NOT MANAGED (2026-11-03);
            # the general feedback question on the row is kept
            _k = list(r)
            _rest = [k for k in _k[_k.index(AVAIL) + 1:_k.index("_index")]
                     if not k.startswith("Please share any feedback")]
            r["_blanked"] = {k: r[k] for k in _rest}
            for k in _rest:
                r[k] = None
            ZONE_RECODED.append(r)

N_SUB, N_MED = len(SUB), len(MEDROWS)
FACILITIES = {s(r[FAC]) for r in SUB}
CHWS = {(s(r[FAC]), s(r[CHW])) for r in SUB}

# Thresholds are form settings, identical on every row so far but read from the
# data rather than written down - a form edit must move the analysis with it.
T_UNDER = num(Counter(s(r[UNDER_T]) for r in SUB).most_common(1)[0][0]) or 1.0
T_OVER = num(Counter(s(r[OVER_T]) for r in SUB).most_common(1)[0][0]) or 2.0
T_VARY = len({s(r[UNDER_T]) for r in SUB}) > 1 or len({s(r[OVER_T]) for r in SUB}) > 1


# ---------------------------------------------------------------
# derived: stock classification
# ---------------------------------------------------------------
# The CHW form has no stock_classification field (the facility form does), so
# it is derived here from months of stock against the form's own thresholds.
# "No consumption" is kept as its own class rather than folded into overstock:
# with AMC = 0 the months-of-stock ratio is undefined, not infinite, and it is
# the single largest group.
CLASSES = ["Stock out", "Understock", "Plan stock", "Overstock", "No consumption"]


def classify(r):
    if not yes(r[MANAGED]):
        return ""
    if no(r[AVAIL]):
        return "Stock out"
    if not yes(r[AVAIL]):
        return ""
    a = num(r[AMC])
    m = num(r[MOS])
    if m is None:
        return "No consumption" if (a is not None and a == 0) else ""
    if m < T_UNDER:
        return "Understock"
    if m <= T_OVER:
        return "Plan stock"
    return "Overstock"


for r in MEDROWS:
    r["_class"] = classify(r)


# ---------------------------------------------------------------
# derived: feedback themes
# ---------------------------------------------------------------
# Free text explaining why a medicine is absent. Rules are ordered - the first
# match wins - and every unmatched comment is listed verbatim in the Feedback
# sheet so new phrasings surface instead of disappearing into "other".
FB_THEMES = [
    ("Transport cost / access to the health centre",
     r"pocket money|transport|distance|far from"),
    ("Storage or equipment need",
     r"\bbox(?:es|e)?\b|storing|storage|shelf|shelv"),
    ("MFT zoning / another first-line in use",
     r"\bmft\b|zoning|\bzone\b|\bblock\b|catchment|other (?:region|zone)|moved to|"
     r"no longer in use|use[sd]? dhap|use[sd]? instead|switch|transition|"
     r"coartem|not yet enrolled|"
     r"out of zone|not allowed|never (?:been )?(?:in )?use|not\s+used|they use|"
     r"we use|ano(?:t)?her line|arthemeter|lumefantrine|zinc of \d|"
     r"^(?:pyronaridine|artem|as[iy]?py|dhap|the distributed acts)|"
     r"\b(?:as[iy]?py|artem)\b|rotational therapy|other strengths"),
    ("Not managed by this CHW",
     r"not\s*managed|never managed|doesn'?t manage|do(?:es)? not manage|"
     r"^n/?a\.?$"),
    ("CHW absent or unable to serve",
     r"was sick|\bsick\b|did ?n[o']t manage to (?:provide|serve)|maternity"),
    ("CHW new or not yet trained",
     r"\bnew\b|not (?:yet )?trained|training|not yet start"),
    ("Medicines returned / CHW stopped",
     r"return|handed (?:back|over)|stopped|no longer serv"),
    ("Shared with or held by another CHW",
     r"another chw|other chw|neighbo\w*ring chw|cell coordi|same village|"
     r"gave (?:the )?quantit|sent to"),
    ("Recently or imminently resupplied",
     r"received (?:the product )?this month|received in september|"
     r"product received this month|will be served along|^received\s+\d"),
    ("Stock-out or supply delay",
     r"stock(?:ed)?\s*-?\s*out|stok\s*out|rupture|shortage|"
     r"not (?:yet )?(?:supplied|received|delivered)|delay|out of stock|"
     r"insufficient|no supply|no stock|not available|small quantity|^waiting|"
     r"don'?t have|never had|was finished|finished in|received in 20\d\d|"
     r"no medication|does not have any medication|awaiting|still waiting|"
     r"forgot to request|lack of|could not (?:treat|serve)"),
    ("Asks for more or more timely supply",
     r"turasaba|turifuza|ku gihe|increase|avail enough|\bis needed\b|"
     r"requisition|have to avail|need(?:s|ed)? (?:to|more)|key even though|"
     r"improvement on|suggested to get|enough products|on time"),
    ("Expiry / quality concern", r"expir|damag|quality"),
    ("Low malaria burden / few cases",
     r"few (?:patient|case)|no (?:positive )?(?:malaria )?(?:case|patient|child|"
     r"consumption)|prevalence|reduc\w+|low (?:demand|consumption)|\brare\b|"
     r"not getting the case|no transaction|"
     r"received (?:few|a few|no|two|a moderate|\d)|"
     r"not many|\bnegative\b|incident|\bcase[sd]?\b|treated|don'?t use frequently|"
     r"only \d|only one|did ?n[o']t expect|dry season|no use so far|no patien\w*|"
     r"do not move without|don'?t move without"),
    ("Reports no issue / stock adequate",
     r"^(ok|okay|none|no|nothing|good|fine|no comment|n/a|-|0)\.?$|"
     r"no (?:issue|problem|concern|challenge)|well managed|good (?:stock|and well)|"
     r"stock (?:is )?(?:available|enough|kept)|available|keep good|sufficient|"
     r"so far so good|nothing is|correct quantit|is good|^goo[dg]$|well kept|"
     r"good ?ly? managed|well done"),
]


def theme(text):
    t = s(text).lower()
    if not t:
        return ""
    for label, pattern in FB_THEMES:
        if re.search(pattern, t):
            return label
    return "(other - read verbatim)"


for r in MEDROWS:
    r["_theme"] = theme(r[FEEDBACK])


# ---------------------------------------------------------------
# table container + writers
# ---------------------------------------------------------------
class Table:
    def __init__(self, title, headers, rows, note=None, total=None, chart=None):
        self.title, self.headers, self.rows = title, headers, rows
        self.note, self.total, self.chart = note, total, chart
        self.pct = {i for i, h in enumerate(headers)
                    if "%" in str(h) or str(h).strip().lower().endswith("rate")}

    def chart_data(self):
        if not self.chart:
            return []
        _, lc, vc, mx = self.chart
        out = [(s(r[lc]), float(r[vc])) for r in self.rows
               if len(r) > max(lc, vc)
               and isinstance(r[vc], (int, float)) and not isinstance(r[vc], bool)]
        return out[:mx] if mx else out


def freq(rows, keyfn, denom=None, sort_by_count=True, top=None):
    c = Counter(keyfn(r) for r in rows)
    n = denom or sum(c.values())
    items = c.most_common(top) if sort_by_count else sorted(c.items())
    return [[k, v, v / n if n else 0] for k, v in items]


def stat_row(label, values):
    v = sorted(x for x in values if x is not None)
    if not v:
        return [label, 0, "", "", "", "", "", "", ""]
    k = len(v)

    def q(p):
        i = p * (k - 1)
        lo, hi = int(i), min(int(i) + 1, k - 1)
        return v[lo] + (v[hi] - v[lo]) * (i - lo)

    return [label, k, round(sum(v), 1), round(sum(v) / k, 2), round(v[0], 2),
            round(q(.25), 2), round(statistics.median(v), 2), round(q(.75), 2),
            round(v[-1], 2)]


STAT_HDR = ["Measure", "n", "Sum", "Mean", "Min", "P25", "Median", "P75", "Max"]


# ---------------------------------------------------------------
# the group profile - one consistent measure set, reused for every dimension
# ---------------------------------------------------------------
PROFILE_HDR = ["Group", "Observations", "CHWs", "Facilities", "Medicines",
               "Managed", "Managed %", "Available", "Available %",
               "Stock out", "Understock", "Plan stock", "Overstock",
               "No consumption", "Median months of stock", "Qty on hand",
               "AMC (total)"]


def profile(rows, keyfn, sort_key=None):
    g = defaultdict(list)
    for r in rows:
        g[keyfn(r)].append(r)
    out = []
    for k, rs in g.items():
        managed = [r for r in rs if yes(r[MANAGED])]
        av = sum(1 for r in managed if yes(r[AVAIL]))
        cls = Counter(r["_class"] for r in rs if r["_class"])
        mos = [num(r[MOS]) for r in managed]
        mos = [x for x in mos if x is not None]
        out.append([
            k, len(rs),
            len({(s(r["_p"][FAC]), s(r["_p"][CHW])) for r in rs}),
            len({s(r["_p"][FAC]) for r in rs}),
            len({r["_short"] for r in rs}),
            len(managed), len(managed) / len(rs) if rs else None,
            av, av / len(managed) if managed else None,
            cls.get("Stock out", 0), cls.get("Understock", 0),
            cls.get("Plan stock", 0), cls.get("Overstock", 0),
            cls.get("No consumption", 0),
            round(statistics.median(mos), 2) if mos else None,
            round(sum(num(r[QTY]) or 0 for r in managed)),
            round(sum(num(r[AMC]) or 0 for r in managed)),
        ])
    out.sort(key=sort_key or (lambda r: -r[1]))
    return out


def total_row(rows, label="ALL"):
    return profile(rows, lambda _: label)[0]


# ---------------------------------------------------------------
# shared figures - computed, never written down: this export is a live interim
# cut and every one of these moves between refreshes.
# ---------------------------------------------------------------
MANAGED_ROWS = [r for r in MEDROWS if yes(r[MANAGED])]
NOTMANAGED_ROWS = [r for r in MEDROWS if no(r[MANAGED])]
AVAIL_ROWS = [r for r in MANAGED_ROWS if yes(r[AVAIL])]
OUT_ROWS = [r for r in MANAGED_ROWS if no(r[AVAIL])]
N_MANAGED, N_AVAIL, N_OUT = len(MANAGED_ROWS), len(AVAIL_ROWS), len(OUT_ROWS)
CLASS_N = Counter(r["_class"] for r in MEDROWS if r["_class"])
N_CLASSIFIED = sum(CLASS_N.values())
MOS_ALL = [num(r[MOS]) for r in MANAGED_ROWS]
MOS_ALL = [x for x in MOS_ALL if x is not None]
AMC_ZERO = sum(1 for r in MANAGED_ROWS if (num(r[AMC]) or 0) == 0)
STIMES = sorted(s(r["_submission_time"]) for r in SUB)
N_VERSIONS = len({s(r.get("__version__")) for r in SUB})
N_GEO = sum(1 for r in SUB if not blank(r.get("_start_location_latitude")))
CONS_TOT = {m: sum(num(r[CONS[m]]) or 0 for r in MANAGED_ROWS) for m in MONTHS}


def pct(a, b):
    return a / b if b else None


# ---------------------------------------------------------------
# 1. README
# ---------------------------------------------------------------
readme = [
    Table("What this workbook is", ["Item", "Value"], [
        ["Source project", "Medicine stock assessment for CHWs (KoboToolbox)"],
        ["Kobo asset", "adtEfgzyyULozJVAxfjRLD, export setting estFypyXnbDhuXNnmGmfbiB "
                       "(\"Stock_CHWs\")"],
        ["Source file", SRC],
        ["Submissions", N_SUB],
        ["CHW x medicine observations", N_MED],
        ["Submission window", f"{STIMES[0]} to {STIMES[-1]}" if STIMES else "n/a"],
        ["Generated by", "analyze_medicine_assessment_chw.py"],
        ["Companion analysis", "analyze_medicine_assessment.py - the facility-level "
                               "stock assessment, a separate Kobo project"],
    ]),
    Table("Unit of analysis - read before quoting any rate", ["Point", "Detail"], [
        ["The unit is the CHW x medicine row",
         f"{N_MED} observations from {N_SUB} submissions. Baskets differ in size "
         f"(most CHWs are assessed on 13 medicines, some on fewer), so a "
         f"submission-level rate would weight a short basket like a full one."],
        ["Availability is computed over MANAGED rows only",
         f"The form asks 'available on visit day' only where the CHW manages the "
         f"medicine: {N_MANAGED} of {N_MED} rows. The other "
         f"{N_MED - N_MANAGED} are blank by skip logic, not missing data - "
         f"counting them as unavailable would understate availability badly."],
        ["'Not managed' is mostly MFT zoning, not a supply gap",
         "RBC's multiple first-line therapy roll-out assigns each area one of "
         "ALu / DHAP / Pyronaridine-Artesunate. A CHW outside a DHAP zone should "
         "have no DHAP. See the 'First-line antimalarial families' table."],
        ["Months of stock is undefined where consumption is zero",
         f"{AMC_ZERO} of {N_MANAGED} managed rows report zero June-August "
         f"consumption, so Kobo leaves stock_level_months blank. Those rows are "
         f"classified 'No consumption', never 'Overstock' - the ratio does not "
         f"exist rather than being large."],
        ["Stock classification is derived here, not in the form",
         f"Understock < {T_UNDER:g} months <= Plan stock <= {T_OVER:g} months < "
         f"Overstock, using the form's own thresholds."
         + (" WARNING: thresholds vary between submissions - check the form."
            if T_VARY else "")],
        ["Facility means the CHW's parent health centre",
         f"All {N_SUB} submissions are attached to a health centre "
         f"({len(FACILITIES)} distinct). The health-post reattributions applied in "
         f"the prescription analysis do not apply to this form."],
    ]),
    Table("What this form does not collect", ["Absent", "Consequence"], [
        ["EML applicability", "No applicable-but-not-managed gap analysis."],
        ["System stock count", "No physical-vs-system inventory accuracy."],
        ["Stock-out day history", "Stock-out is a visit-day state only - no duration, "
                                  "no stock-out rate over the quarter."],
        ["Requisitions and deliveries", "No fulfilment analysis."],
        ["RMS central/regional availability", "No supply-chain attribution for an "
                                              "unavailable medicine."],
        ["Unit price", "No financial loss estimate."],
    ], note="All six exist on the facility form. If they are wanted for CHWs they "
            "have to be added to the CHW tool - they cannot be derived."),
    Table("Refreshing", ["Step", "Command"], [
        ["1. Re-fetch the export", "python3 fetch_kobo_export.py chw"],
        ["2. Re-run this analysis", "python3 analyze_medicine_assessment_chw.py"],
        ["If the save fails with PermissionError",
         f"chmod u+w '{OUT}'  (Box resets synced files to read-only)"],
    ]),
]

# ---------------------------------------------------------------
# 2. Overview
# ---------------------------------------------------------------
overview = [
    Table("Headline figures", ["Measure", "Value"], [
        ["Submissions", N_SUB],
        ["CHWs assessed", len(CHWS)],
        ["Parent health centres", len(FACILITIES)],
        ["Districts", len({s(r[DIST]) for r in SUB})],
        ["Medicines in the basket", len({r["_short"] for r in MEDROWS})],
        ["CHW x medicine observations", N_MED],
        ["Managed by the CHW", f"{N_MANAGED} ({pct(N_MANAGED, N_MED):.1%})"],
        ["Available on visit day (of managed)",
         f"{N_AVAIL} ({pct(N_AVAIL, N_MANAGED):.1%})"],
        ["Stocked out on visit day (of managed)",
         f"{N_OUT} ({pct(N_OUT, N_MANAGED):.1%})"],
        ["Median months of stock (where defined)",
         round(statistics.median(MOS_ALL), 2) if MOS_ALL else "n/a"],
        ["Managed rows with zero Jun-Aug consumption",
         f"{AMC_ZERO} ({pct(AMC_ZERO, N_MANAGED):.1%})"],
        ["Total quantity on hand (units, all medicines)",
         round(sum(num(r[QTY]) or 0 for r in MANAGED_ROWS))],
        ["Total Jun-Aug consumption (units)", round(sum(CONS_TOT.values()))],
        ["Submission window", f"{STIMES[0]} to {STIMES[-1]}" if STIMES else "n/a"],
    ]),
    Table("Stock classification across all classified observations",
          ["Class", "Observations", "% of classified"],
          [[c, CLASS_N.get(c, 0), pct(CLASS_N.get(c, 0), N_CLASSIFIED)]
           for c in CLASSES if CLASS_N.get(c, 0)],
          note=f"{N_CLASSIFIED} of {N_MED} observations classify. The rest are rows "
               f"the CHW does not manage, where the question is not asked.",
          chart=("v", 0, 1, 8)),
    Table("First-line antimalarial families - the MFT-zoning read",
          ["Family", "Observations", "Managed", "Managed %", "Available",
           "Available % of managed"],
          [[f,
            sum(1 for r in MEDROWS if r["_family"] == f),
            sum(1 for r in MANAGED_ROWS if r["_family"] == f),
            pct(sum(1 for r in MANAGED_ROWS if r["_family"] == f),
                sum(1 for r in MEDROWS if r["_family"] == f)),
            sum(1 for r in AVAIL_ROWS if r["_family"] == f),
            pct(sum(1 for r in AVAIL_ROWS if r["_family"] == f),
                sum(1 for r in MANAGED_ROWS if r["_family"] == f))]
           for f in sorted({r["_family"] for r in MEDROWS})],
          note="A low 'managed %' on an antimalarial family is expected outside its "
               "MFT zone. A low 'available %' within the managed rows is the real "
               "supply signal.",
          chart=("v", 0, 3, 6)),
]

# ---------------------------------------------------------------
# 3. Coverage
# ---------------------------------------------------------------
_per_fac = Counter(s(r[FAC]) for r in SUB)
coverage = [
    Table("Submissions by province", ["Province", "Submissions", "%"],
          freq(SUB, lambda r: s(r[PROV])), chart=("v", 0, 1, 8)),
    Table("Submissions by district", ["District", "Submissions", "%"],
          freq(SUB, lambda r: s(r[DIST])), chart=("h", 0, 1, 20)),
    Table("CHWs assessed per health centre",
          ["Health centre", "District", "CHWs assessed", "Medicine observations"],
          sorted([[f,
                   next(s(r[DIST]) for r in SUB if s(r[FAC]) == f),
                   _per_fac[f],
                   sum(1 for r in MEDROWS if s(r["_p"][FAC]) == f)]
                  for f in sorted(FACILITIES)], key=lambda x: (-x[2], x[0])),
          note="One CHW per submission; a health centre appearing twice means two "
               "CHWs were assessed there."),
    Table("Submissions by day", ["Date", "Submissions"],
          sorted(Counter(day(r["_submission_time"]) for r in SUB).items()),
          chart=("v", 0, 1, 40)),
    Table("Data collectors", ["Collector", "Submissions", "Health centres",
                              "Medicine observations"],
          sorted([[e, n,
                   len({s(r[FAC]) for r in SUB if s(r[ENUM]) == e}),
                   sum(1 for r in MEDROWS if s(r["_p"][ENUM]) == e)]
                  for e, n in Counter(s(r[ENUM]) for r in SUB).items()],
                 key=lambda x: -x[1]),
          note="A blank collector name means the field was left empty on that "
               "submission."),
    Table("Basket size per CHW", ["Medicines assessed", "CHWs", "%"],
          [[k, v, v / N_SUB] for k, v in
           sorted(Counter(Counter(r["_parent_index"] for r in MEDROWS).values()).items())],
          note="The repeat group is driven by the multi-select on the cover page, so "
               "a short basket means fewer medicines were selected for assessment, "
               "not that rows are missing."),
    Table("Form versions in use", ["__version__", "Submissions"],
          sorted(Counter(s(r["__version__"]) for r in SUB).items(), key=lambda x: -x[1]),
          note=f"{N_VERSIONS} versions. Columns present on only some versions are "
               f"blank on the rest - the export was taken with "
               f"fields_from_all_versions on."),
]

# ---------------------------------------------------------------
# 4. Data quality
# ---------------------------------------------------------------
empty_cols = ([["submission", h] for h in HDR_SUB
               if h and all(blank(r.get(h)) for r in SUB)]
              + [["medicine_group", h] for h in HDR_MED
                 if h and all(blank(r.get(h)) for r in MEDROWS)])
nonnum = [[c, sum(1 for r in MEDROWS
                  if not blank(r.get(c)) and num(r.get(c)) is None)]
          for c in (AMC, MOS, QTY) + tuple(CONS.values())]
nonnum = [r for r in nonnum if r[1]]

dupes = [[f, c, n] for (f, c), n in
         Counter((s(r["_p"][FAC]) + " / " + s(r["_p"][CHW]), r["_short"])
                 for r in MEDROWS).items() if n > 1]

skip_checks = [
    ["Availability answered where managed",
     sum(1 for r in MANAGED_ROWS if not blank(r[AVAIL])), N_MANAGED],
    ["Availability blank where NOT managed",
     sum(1 for r in NOTMANAGED_ROWS if blank(r[AVAIL])), len(NOTMANAGED_ROWS)],
    ["Quantity recorded where available",
     sum(1 for r in AVAIL_ROWS if not blank(r[QTY])), N_AVAIL],
    ["Quantity blank where stocked out",
     sum(1 for r in OUT_ROWS if blank(r[QTY])), N_OUT],
    ["Consumption recorded where managed",
     sum(1 for r in MANAGED_ROWS if not blank(r[CONS["June"]])), N_MANAGED],
]

qty_zero_avail = [r for r in AVAIL_ROWS if (num(r[QTY]) or 0) == 0]
mos_extreme = [r for r in MANAGED_ROWS if (num(r[MOS]) or 0) > 12]
short_visit = [r for r in SUB if r.get("start_time") and r.get("end_time")
               and (r["end_time"] - r["start_time"]).total_seconds() < 900]
suspect_chw = [r for r in SUB if len(s(r[CHW])) < 5 or
               re.search(r"\b(hc|cs|health)\b", s(r[CHW]).lower())]

quality = [
    Table("Structural checks", ["Check", "Result", "Expected"], [
        ["Submissions in export", N_SUB, "matches the Kobo submission count"],
        ["Medicine rows linked to a submission", N_MED, f"{N_MED} (0 orphans)"],
        ["Orphan medicine rows dropped", ORPHANS, 0],
        ["Distinct _id values", len({r["_id"] for r in SUB}), N_SUB],
        ["Distinct CHW x health centre pairs", len(CHWS), N_SUB],
        ["Submissions with a GPS fix", N_GEO, N_SUB],
        ["Facility level values", ", ".join(sorted({s(r[LEVEL]) for r in SUB})),
         "Health Centre only - CHWs attach to an HC"],
    ]),
    Table("Skip-logic consistency", ["Check", "Rows", "Expected"], skip_checks,
          note="Rows and Expected should match. A gap means the form was answered "
               "against its own logic, or the logic changed between versions."),
    Table("Repeated CHW x medicine observations",
          ["CHW (health centre / name)", "Medicine", "Rows"], dupes or [["none", "", 0]],
          note="A medicine assessed twice for the same CHW. None expected; any row "
               "here is a duplicate submission or a re-entry."),
    Table("Value-range flags", ["Flag", "Rows", "What it means"], [
        ["Available but quantity is zero", len(qty_zero_avail),
         "Answered 'available' then recorded 0 units - one of the two is wrong, and "
         "the row is classified from the quantity, so it reads as understock."],
        ["Months of stock above 12", len(mos_extreme),
         "Consumption near zero against a normal quantity. Real, but it distorts a "
         "mean - the profiles use the median throughout."],
        ["Managed with zero Jun-Aug consumption", AMC_ZERO,
         "Classified 'No consumption'. Either the CHW genuinely dispensed none, or "
         "consumption was not transcribed from the register."],
        ["Visit under 15 minutes", len(short_visit),
         "start_time to end_time. A 13-medicine basket in that time is worth a check."],
        ["CHW name looks like a placeholder or a facility", len(suspect_chw),
         "Names under 5 characters, or containing HC / CS / Health."],
    ]),
    Table("CHW names needing a look",
          ["Health centre", "District", "Name as entered", "Submitted"],
          [[s(r[FAC]), s(r[DIST]), s(r[CHW]), day(r["_submission_time"])]
           for r in suspect_chw] or [["none", "", "", ""]]),
    Table("Derived fields - values that will not parse as numbers",
          ["Column", "Rows"], nonnum or [["none", 0]],
          note="Kobo can export the literal string 'NaN' from a calculate field; "
               "num() rejects it so it cannot poison a sum."),
    Table("Columns that are completely empty", ["Sheet", "Column"],
          empty_cols or [["none", ""]]),
    Table("Missingness - submission sheet", ["Column", "Blank", "% blank"],
          sorted([[h, sum(1 for r in SUB if blank(r.get(h))),
                   sum(1 for r in SUB if blank(r.get(h))) / N_SUB]
                  for h in HDR_SUB if h and any(not blank(r.get(h)) for r in SUB)],
                 key=lambda x: -x[1])[:25],
          note="Top 25 by blank count; fully empty columns are listed above instead."),
    Table("Missingness - medicine_group sheet", ["Column", "Blank", "% blank"],
          sorted([[h, sum(1 for r in MEDROWS if blank(r.get(h))),
                   sum(1 for r in MEDROWS if blank(r.get(h))) / N_MED]
                  for h in HDR_MED if h and any(not blank(r.get(h)) for r in MEDROWS)],
                 key=lambda x: -x[1])[:25],
          note="High blank rates on availability and quantity are skip logic, not "
               "missing data - see the skip-logic table above."),
]

# ---------------------------------------------------------------
# 5. Medicine coverage
# ---------------------------------------------------------------
med_names = sorted({r["_short"] for r in MEDROWS},
                   key=lambda m: -sum(1 for r in MEDROWS if r["_short"] == m))
medcov = [
    Table("Medicine roster",
          ["Medicine", "Family", "Full label as exported", "Observations",
           "Managed", "Managed %"],
          [[m,
            family(m),
            next(s(r[MED]) for r in MEDROWS if r["_short"] == m)[:60],
            sum(1 for r in MEDROWS if r["_short"] == m),
            sum(1 for r in MANAGED_ROWS if r["_short"] == m),
            pct(sum(1 for r in MANAGED_ROWS if r["_short"] == m),
                sum(1 for r in MEDROWS if r["_short"] == m))]
           for m in med_names],
          note="Observations below the submission count mean the medicine was not "
               "selected for assessment for every CHW.",
          chart=("h", 0, 5, 13)),
    Table("Management by medicine", ["Medicine", "Managed", "Not managed",
                                     "Managed %"],
          sorted([[m,
                   sum(1 for r in MANAGED_ROWS if r["_short"] == m),
                   sum(1 for r in NOTMANAGED_ROWS if r["_short"] == m),
                   pct(sum(1 for r in MANAGED_ROWS if r["_short"] == m),
                       sum(1 for r in MEDROWS if r["_short"] == m))]
                  for m in med_names], key=lambda x: x[3] or 0)),
    Table("Why a medicine is not managed - coded from the comment",
          ["Theme", "Rows", "% of not-managed rows"],
          freq(NOTMANAGED_ROWS, lambda r: r["_theme"] or "(no comment)",
               denom=len(NOTMANAGED_ROWS)),
          note="Coded by keyword; the rules are in FB_THEMES. Anything landing in "
               "'(other - read verbatim)' needs a new rule - the verbatim text is "
               "on the Feedback sheet.",
          chart=("h", 0, 1, 10)),
    Table("Not-managed rows by medicine and province",
          ["Medicine"] + sorted({s(r[PROV]) for r in SUB}) + ["Total"],
          [[m] + [sum(1 for r in NOTMANAGED_ROWS
                      if r["_short"] == m and s(r["_p"][PROV]) == p)
                  for p in sorted({s(r[PROV]) for r in SUB})]
           + [sum(1 for r in NOTMANAGED_ROWS if r["_short"] == m)]
           for m in med_names],
          note="Read down a column: an antimalarial absent across a whole province "
               "is MFT zoning. Absent from scattered CHWs within a province is not."),
]

# ---------------------------------------------------------------
# 6. Availability profiles
# ---------------------------------------------------------------
def prof_table(title, rows, keyfn, note=None, sort_key=None, chart=None):
    body = profile(rows, keyfn, sort_key)
    return Table(title, PROFILE_HDR, body, note=note,
                 total=total_row(rows), chart=chart)


availability = [
    prof_table("By province", MEDROWS, lambda r: s(r["_p"][PROV]),
               chart=("v", 0, 8, 8)),
    prof_table("By district", MEDROWS, lambda r: s(r["_p"][DIST]),
               chart=("h", 0, 8, 25)),
    prof_table("By medicine", MEDROWS, lambda r: r["_short"],
               chart=("h", 0, 8, 13)),
    prof_table("By antimalarial family", MEDROWS, lambda r: r["_family"]),
    prof_table("By data collector", MEDROWS, lambda r: s(r["_p"][ENUM]),
               note="Wide spread between collectors on the same measure can be a "
                    "real geographic difference or an interviewing difference - "
                    "worth checking before it is read as either."),
    Table("Lowest availability - medicines with at least 10 managed observations",
          ["Medicine", "Managed", "Available", "Stocked out", "Available %"],
          sorted([[m,
                   sum(1 for r in MANAGED_ROWS if r["_short"] == m),
                   sum(1 for r in AVAIL_ROWS if r["_short"] == m),
                   sum(1 for r in OUT_ROWS if r["_short"] == m),
                   pct(sum(1 for r in AVAIL_ROWS if r["_short"] == m),
                       sum(1 for r in MANAGED_ROWS if r["_short"] == m))]
                  for m in med_names
                  if sum(1 for r in MANAGED_ROWS if r["_short"] == m) >= 10],
                 key=lambda x: x[4] or 0),
          chart=("h", 0, 4, 13)),
    Table("Every stock-out on visit day",
          ["Health centre", "District", "CHW", "Medicine", "Jun-Aug consumption",
           "Submitted"],
          sorted([[s(r["_p"][FAC]), s(r["_p"][DIST]), s(r["_p"][CHW]), r["_short"],
                   round(sum(num(r[CONS[m]]) or 0 for m in MONTHS)),
                   day(r["_p"]["_submission_time"])]
                  for r in OUT_ROWS], key=lambda x: (-x[4], x[0])),
          note="Sorted by quarter consumption: a stock-out of a medicine the CHW was "
               "actually dispensing is the one that cost a patient a treatment."),
]

# ---------------------------------------------------------------
# 7. Stock levels
# ---------------------------------------------------------------
def mos_band(v):
    if v is None:
        return "not computed (zero consumption)"
    if v == 0:
        return "0 (nothing on hand)"
    if v < 1:
        return "under 1 month"
    if v <= 2:
        return "1 to 2 months"
    if v <= 6:
        return "2 to 6 months"
    if v <= 12:
        return "6 to 12 months"
    return "over 12 months"


BANDS = ["0 (nothing on hand)", "under 1 month", "1 to 2 months", "2 to 6 months",
         "6 to 12 months", "over 12 months", "not computed (zero consumption)"]
band_n = Counter(mos_band(num(r[MOS])) for r in MANAGED_ROWS)

stock = [
    Table("Months of stock on hand", STAT_HDR,
          [stat_row("All managed rows where computed", [num(r[MOS]) for r in MANAGED_ROWS])]
          + [stat_row(m, [num(r[MOS]) for r in MANAGED_ROWS if r["_short"] == m])
             for m in med_names],
          note=f"Computed on {len(MOS_ALL)} of {N_MANAGED} managed rows; the rest "
               f"have zero consumption so the ratio is undefined. Read the median, "
               f"not the mean - a near-zero denominator produces values in the "
               f"hundreds."),
    Table("Months of stock, banded", ["Band", "Observations", "% of managed"],
          [[b, band_n.get(b, 0), pct(band_n.get(b, 0), N_MANAGED)]
           for b in BANDS if band_n.get(b, 0)],
          chart=("v", 0, 1, 8)),
    Table("Stock classification by province",
          ["Province", "Classified"] + CLASSES + ["Stock out %", "Understock %"],
          [[p, sum(1 for r in MEDROWS if s(r["_p"][PROV]) == p and r["_class"])]
           + [sum(1 for r in MEDROWS if s(r["_p"][PROV]) == p and r["_class"] == c)
              for c in CLASSES]
           + [pct(sum(1 for r in MEDROWS if s(r["_p"][PROV]) == p
                      and r["_class"] == "Stock out"),
                  sum(1 for r in MEDROWS if s(r["_p"][PROV]) == p and r["_class"])),
              pct(sum(1 for r in MEDROWS if s(r["_p"][PROV]) == p
                      and r["_class"] == "Understock"),
                  sum(1 for r in MEDROWS if s(r["_p"][PROV]) == p and r["_class"]))]
           for p in sorted({s(r[PROV]) for r in SUB})]),
    Table("Stock classification by medicine",
          ["Medicine", "Classified"] + CLASSES + ["Stock out %"],
          sorted([[m, sum(1 for r in MEDROWS if r["_short"] == m and r["_class"])]
                  + [sum(1 for r in MEDROWS if r["_short"] == m and r["_class"] == c)
                     for c in CLASSES]
                  + [pct(sum(1 for r in MEDROWS if r["_short"] == m
                             and r["_class"] == "Stock out"),
                         sum(1 for r in MEDROWS if r["_short"] == m and r["_class"]))]
                  for m in med_names], key=lambda x: -(x[-1] or 0)),
          chart=("h", 0, 7, 13)),
    Table("Quantity on hand", STAT_HDR,
          [stat_row("All available rows", [num(r[QTY]) for r in AVAIL_ROWS])]
          + [stat_row(m, [num(r[QTY]) for r in AVAIL_ROWS if r["_short"] == m])
             for m in med_names],
          note="Units as entered - packs for the ALu presentations, tests for RDTs, "
               "tablets or sachets elsewhere. Do not add across medicines."),
    Table("CHWs with the most stock-outs",
          ["Health centre", "CHW", "District", "Managed", "Stocked out",
           "Stock out %"],
          sorted([[f, c, next(s(r[DIST]) for r in SUB
                              if s(r[FAC]) == f and s(r[CHW]) == c),
                   sum(1 for r in MANAGED_ROWS
                       if (s(r["_p"][FAC]), s(r["_p"][CHW])) == (f, c)),
                   sum(1 for r in OUT_ROWS
                       if (s(r["_p"][FAC]), s(r["_p"][CHW])) == (f, c)),
                   pct(sum(1 for r in OUT_ROWS
                           if (s(r["_p"][FAC]), s(r["_p"][CHW])) == (f, c)),
                       sum(1 for r in MANAGED_ROWS
                           if (s(r["_p"][FAC]), s(r["_p"][CHW])) == (f, c)))]
                  for f, c in CHWS], key=lambda x: (-x[4], -(x[5] or 0)))[:30],
          note="Top 30 by count of stocked-out medicines.",
          chart=("h", 1, 4, 20)),
]

# ---------------------------------------------------------------
# 8. Consumption
# ---------------------------------------------------------------
consumption = [
    Table("Consumption by month (all managed rows)",
          ["Month", "Total quantity", "Rows reporting", "Rows reporting zero",
           "Mean where non-zero"],
          [[m, round(CONS_TOT[m]),
            sum(1 for r in MANAGED_ROWS if num(r[CONS[m]]) is not None),
            sum(1 for r in MANAGED_ROWS if (num(r[CONS[m]]) or 0) == 0),
            round(statistics.mean([num(r[CONS[m]]) for r in MANAGED_ROWS
                                   if (num(r[CONS[m]]) or 0) > 0]), 1)
            if any((num(r[CONS[m]]) or 0) > 0 for r in MANAGED_ROWS) else None]
           for m in MONTHS],
          chart=("v", 0, 1, 3)),
    Table("Monthly consumption distribution", STAT_HDR,
          [stat_row(m, [num(r[CONS[m]]) for r in MANAGED_ROWS]) for m in MONTHS]
          + [stat_row("Average monthly (AMC)", [num(r[AMC]) for r in MANAGED_ROWS])]),
    Table("Consumption by medicine",
          ["Medicine", "Managed rows", "June", "July", "August",
           "Quarter total", "Rows with zero consumption"],
          sorted([[m, sum(1 for r in MANAGED_ROWS if r["_short"] == m)]
                  + [round(sum(num(r[CONS[mo]]) or 0 for r in MANAGED_ROWS
                               if r["_short"] == m)) for mo in MONTHS]
                  + [round(sum(sum(num(r[CONS[mo]]) or 0 for mo in MONTHS)
                               for r in MANAGED_ROWS if r["_short"] == m)),
                     sum(1 for r in MANAGED_ROWS if r["_short"] == m
                         and (num(r[AMC]) or 0) == 0)]
                  for m in med_names], key=lambda x: -x[5]),
          chart=("h", 0, 5, 13)),
    Table("Consumption by province",
          ["Province", "Managed rows", "June", "July", "August", "Quarter total"],
          sorted([[p, sum(1 for r in MANAGED_ROWS if s(r["_p"][PROV]) == p)]
                  + [round(sum(num(r[CONS[mo]]) or 0 for r in MANAGED_ROWS
                               if s(r["_p"][PROV]) == p)) for mo in MONTHS]
                  + [round(sum(sum(num(r[CONS[mo]]) or 0 for mo in MONTHS)
                               for r in MANAGED_ROWS if s(r["_p"][PROV]) == p))]
                  for p in sorted({s(r[PROV]) for r in SUB})], key=lambda x: -x[5])),
    Table("Zero-consumption rows by medicine",
          ["Medicine", "Managed", "Zero Jun-Aug", "% of managed"],
          sorted([[m, sum(1 for r in MANAGED_ROWS if r["_short"] == m),
                   sum(1 for r in MANAGED_ROWS if r["_short"] == m
                       and (num(r[AMC]) or 0) == 0),
                   pct(sum(1 for r in MANAGED_ROWS if r["_short"] == m
                           and (num(r[AMC]) or 0) == 0),
                       sum(1 for r in MANAGED_ROWS if r["_short"] == m))]
                  for m in med_names], key=lambda x: -(x[3] or 0)),
          note="A medicine held but never dispensed in a quarter is either genuinely "
               "unneeded in that area, or a register that was not consulted. The two "
               "cannot be told apart from this form.",
          chart=("h", 0, 3, 13)),
]

# ---------------------------------------------------------------
# 9. CHW scorecard
# ---------------------------------------------------------------
chw_scorecard = [
    Table("One row per CHW assessed",
          ["Health centre", "CHW", "District", "Sector", "Collector", "Submitted",
           "Basket", "Managed", "Available", "Available %", "Stocked out",
           "Median months of stock", "Quarter consumption"],
          sorted([[s(p[FAC]), s(p[CHW]), s(p[DIST]), s(p[SECT]), s(p[ENUM]),
                   day(p["_submission_time"]),
                   len([r for r in MEDROWS if r["_p"] is p]),
                   len([r for r in MEDROWS if r["_p"] is p and yes(r[MANAGED])]),
                   len([r for r in MEDROWS if r["_p"] is p and yes(r[MANAGED])
                        and yes(r[AVAIL])]),
                   pct(len([r for r in MEDROWS if r["_p"] is p and yes(r[MANAGED])
                            and yes(r[AVAIL])]),
                       len([r for r in MEDROWS if r["_p"] is p and yes(r[MANAGED])])),
                   len([r for r in MEDROWS if r["_p"] is p and no(r[AVAIL])]),
                   (lambda v: round(statistics.median(v), 2) if v else None)(
                       [num(r[MOS]) for r in MEDROWS
                        if r["_p"] is p and num(r[MOS]) is not None]),
                   round(sum(sum(num(r[CONS[m]]) or 0 for m in MONTHS)
                             for r in MEDROWS if r["_p"] is p and yes(r[MANAGED])))]
                  for p in SUB], key=lambda x: (x[0], x[1])),
          note="Sorted by health centre. 'Basket' is how many medicines were selected "
               "for assessment; 'Managed' how many the CHW actually holds."),
]

# ---------------------------------------------------------------
# 10. Medicine scorecard
# ---------------------------------------------------------------
medicine_scorecard = [
    Table("One row per medicine",
          ["Medicine", "Family", "Observations", "Managed", "Managed %",
           "Available", "Available %", "Stocked out", "Median months of stock",
           "Qty on hand", "Quarter consumption", "Zero-consumption rows"],
          sorted([[m, family(m),
                   sum(1 for r in MEDROWS if r["_short"] == m),
                   sum(1 for r in MANAGED_ROWS if r["_short"] == m),
                   pct(sum(1 for r in MANAGED_ROWS if r["_short"] == m),
                       sum(1 for r in MEDROWS if r["_short"] == m)),
                   sum(1 for r in AVAIL_ROWS if r["_short"] == m),
                   pct(sum(1 for r in AVAIL_ROWS if r["_short"] == m),
                       sum(1 for r in MANAGED_ROWS if r["_short"] == m)),
                   sum(1 for r in OUT_ROWS if r["_short"] == m),
                   (lambda v: round(statistics.median(v), 2) if v else None)(
                       [num(r[MOS]) for r in MANAGED_ROWS
                        if r["_short"] == m and num(r[MOS]) is not None]),
                   round(sum(num(r[QTY]) or 0 for r in AVAIL_ROWS if r["_short"] == m)),
                   round(sum(sum(num(r[CONS[mo]]) or 0 for mo in MONTHS)
                             for r in MANAGED_ROWS if r["_short"] == m)),
                   sum(1 for r in MANAGED_ROWS if r["_short"] == m
                       and (num(r[AMC]) or 0) == 0)]
                  for m in med_names], key=lambda x: -x[3])),
]

# ---------------------------------------------------------------
# 11. Feedback
# ---------------------------------------------------------------
fb_rows = [r for r in MEDROWS if not blank(r[FEEDBACK])]
other = [r for r in fb_rows if r["_theme"] == "(other - read verbatim)"]
feedback = [
    Table("Coded themes", ["Theme", "Comments", "% of comments"],
          freq(fb_rows, lambda r: r["_theme"], denom=len(fb_rows)),
          note=f"{len(fb_rows)} of {N_MED} rows carry a comment. Coded by keyword "
               f"(FB_THEMES); the first matching rule wins.",
          chart=("h", 0, 1, 10)),
    Table("Themes by management status",
          ["Theme", "On a managed row", "On a not-managed row", "Total"],
          sorted([[t,
                   sum(1 for r in fb_rows if r["_theme"] == t and yes(r[MANAGED])),
                   sum(1 for r in fb_rows if r["_theme"] == t and no(r[MANAGED])),
                   sum(1 for r in fb_rows if r["_theme"] == t)]
                  for t in {r["_theme"] for r in fb_rows}], key=lambda x: -x[3])),
    Table("Most frequent comments verbatim", ["Comment", "Times"],
          Counter(s(r[FEEDBACK]) for r in fb_rows).most_common(30)),
    Table("Uncoded comments - every one, verbatim",
          ["Health centre", "CHW", "Medicine", "Managed", "Comment"],
          [[s(r["_p"][FAC]), s(r["_p"][CHW]), r["_short"], s(r[MANAGED]),
            s(r[FEEDBACK])] for r in other] or [["none", "", "", "", ""]],
          note="These matched no rule. Each one is either a new phrasing that needs a "
               "rule in FB_THEMES, or a genuinely singular remark."),
]

# ---------------------------------------------------------------
# 12. Findings
# ---------------------------------------------------------------
_worst_med = min(((m, pct(sum(1 for r in AVAIL_ROWS if r["_short"] == m),
                          sum(1 for r in MANAGED_ROWS if r["_short"] == m)))
                  for m in med_names
                  if sum(1 for r in MANAGED_ROWS if r["_short"] == m) >= 10),
                 key=lambda x: x[1], default=("n/a", None))
_worst_prov = min(((p, pct(sum(1 for r in AVAIL_ROWS if s(r["_p"][PROV]) == p),
                           sum(1 for r in MANAGED_ROWS if s(r["_p"][PROV]) == p)))
                   for p in {s(r[PROV]) for r in SUB}),
                  key=lambda x: x[1] or 1, default=("n/a", None))
_zero_top = max(((m, sum(1 for r in MANAGED_ROWS if r["_short"] == m
                         and (num(r[AMC]) or 0) == 0))
                 for m in med_names), key=lambda x: x[1], default=("n/a", 0))

findings = [
    Table("Findings and recommended actions",
          ["Area", "Finding", "Evidence", "Action"], [
        ["Availability",
         f"{pct(N_AVAIL, N_MANAGED):.1%} of medicines a CHW manages were physically "
         f"present on the visit day.",
         f"{N_AVAIL} available of {N_MANAGED} managed rows; {N_OUT} stocked out.",
         "Treat this as the headline CHW figure. It is not comparable with the "
         "facility form's availability - different basket, different denominator."],
        ["Availability",
         f"The weakest item is {_worst_med[0]}"
         + (f" at {_worst_med[1]:.1%} available." if _worst_med[1] is not None else "."),
         "Lowest-availability table, medicines with 10+ managed observations.",
         "Check whether the parent health centre was also out on the same date - the "
         "facility form covers the same districts."],
        ["Geography",
         f"Availability is lowest in {_worst_prov[0]}"
         + (f" ({_worst_prov[1]:.1%})." if _worst_prov[1] is not None else "."),
         "Province profile on the Availability sheet.",
         "Compare against the district profile before acting - some provinces carry "
         "only a handful of CHWs so far."],
        ["MFT zoning",
         "Most 'not managed' rows are the two first-line antimalarials the CHW's zone "
         "does not use, not a gap.",
         f"{len(NOTMANAGED_ROWS)} not-managed rows of {N_MED}; the coded reasons are "
         f"led by MFT zoning and 'not managed'.",
         "Report availability over managed rows only. Quoting it over all rows would "
         "report the zoning policy as a stock-out."],
        ["Consumption data",
         f"{AMC_ZERO} of {N_MANAGED} managed rows report zero consumption across all "
         f"three months, which leaves months of stock undefined for them.",
         f"Largest single item: {_zero_top[0]} with {_zero_top[1]} zero rows.",
         "Establish with the field team whether these are true zeroes or registers "
         "that were not consulted. Until then, months of stock describes only the "
         f"{len(MOS_ALL)} rows where it computes."],
        ["Stock levels",
         f"Median months of stock is "
         f"{round(statistics.median(MOS_ALL), 2) if MOS_ALL else 'n/a'} where it can "
         f"be computed, against a plan band of {T_UNDER:g}-{T_OVER:g} months.",
         f"{CLASS_N.get('Understock', 0)} understock, "
         f"{CLASS_N.get('Plan stock', 0)} in band, "
         f"{CLASS_N.get('Overstock', 0)} overstock of {N_CLASSIFIED} classified.",
         "Use the median, not the mean - a handful of near-zero consumption "
         "denominators put the mean in the tens of months."],
        ["Data quality",
         f"{len(qty_zero_avail)} rows say a medicine is available and then record a "
         f"quantity of zero.",
         "Value-range flags on the Data quality sheet.",
         "Add a constraint to the form: quantity must be greater than zero when "
         "availability is Yes."],
        ["Data quality",
         f"{len(suspect_chw)} CHW names are placeholders or facility names.",
         "'CHW names needing a look' on the Data quality sheet.",
         "Correct at source in Kobo; CHW is the unit of analysis, so a placeholder "
         "name breaks any follow-up visit matching."],
        ["Coverage",
         f"{len(CHWS)} CHWs across {len(FACILITIES)} health centres in "
         f"{len({s(r[DIST]) for r in SUB})} districts.",
         f"Submissions ran {STIMES[0][:10]} to {STIMES[-1][:10]}."
         if STIMES else "n/a",
         "Collection is still open - this is an interim cut, and every rate here "
         "will move on the next refresh."],
    ]),
]

SECTIONS = [
    ("README", readme),
    ("Overview", overview),
    ("Coverage", coverage),
    ("Data quality", quality),
    ("Medicine coverage", medcov),
    ("Availability profiles", availability),
    ("Stock levels", stock),
    ("Consumption", consumption),
    ("CHW scorecard", chw_scorecard),
    ("Medicine scorecard", medicine_scorecard),
    ("Feedback", feedback),
    ("Findings", findings),
]

# ---------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="DDE5F0")
TOT_FILL = PatternFill("solid", fgColor="F2F2F2")
TITLE_FONT = Font(bold=True, size=12, color="1F3864")
HDR_FONT = Font(bold=True, size=10)
NOTE_FONT = Font(size=9, italic=True, color="555555")


def write():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for name, tables in SECTIONS:
        ws = wb.create_sheet(name[:31])
        ws.sheet_view.showGridLines = False
        ws.cell(row=1, column=1, value=name).font = Font(bold=True, size=14,
                                                         color="1F3864")
        row = 3
        widths = {}
        for t in tables:
            ws.cell(row=row, column=1, value=t.title).font = TITLE_FONT
            row += 1
            for ci, h in enumerate(t.headers, start=1):
                c = ws.cell(row=row, column=ci, value=str(h))
                c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
                c.alignment = Alignment(wrap_text=True, vertical="bottom")
                widths[ci] = max(widths.get(ci, 9), min(len(str(h)) + 2, 34))
            hdr_row = row
            row += 1
            body = list(t.rows) + ([t.total] if t.total else [])
            for ri, r in enumerate(body):
                is_total = t.total is not None and ri == len(body) - 1
                for ci, v in enumerate(r, start=1):
                    c = ws.cell(row=row, column=ci, value=v)
                    c.border = BORDER
                    c.alignment = Alignment(vertical="top",
                                            wrap_text=isinstance(v, str)
                                            and len(str(v)) > 55)
                    if is_total:
                        c.font = Font(bold=True)
                        c.fill = TOT_FILL
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        c.number_format = "0.0%" if (ci - 1) in t.pct else "#,##0.##"
                    widths[ci] = max(widths.get(ci, 9), min(len(str(v)) + 2, 46))
                row += 1
            if len(body) > 8:
                ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)
            if t.note:
                c = ws.cell(row=row, column=1, value="Note: " + t.note)
                c.font = NOTE_FONT
                c.alignment = Alignment(wrap_text=True, vertical="top")
                ws.merge_cells(start_row=row, start_column=1, end_row=row,
                               end_column=max(len(t.headers), 3))
                ws.row_dimensions[row].height = 12 * (1 + len(t.note) // 120)
                row += 1
            row += 2
        for ci, w in widths.items():
            ws.column_dimensions[get_column_letter(ci)].width = w

    # charts
    ws = wb.create_sheet("Charts")
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1, value="Charts").font = Font(bold=True, size=14,
                                                         color="1F3864")
    ws.cell(row=2, column=1,
            value="Native Excel charts over the data block to the left of each - "
                  "click any chart to edit or copy it.").font = NOTE_FONT
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 14
    row, n = 4, 0
    for section, tables in SECTIONS:
        for t in tables:
            pairs_ = t.chart_data()
            if not pairs_:
                continue
            title = f"{t.title} - {t.headers[t.chart[2]]}"
            ws.cell(row=row, column=1, value=title).font = TITLE_FONT
            ws.cell(row=row, column=2, value=section).font = NOTE_FONT
            h0 = row + 1
            ws.cell(row=h0, column=1, value=str(t.headers[t.chart[1]])).font = HDR_FONT
            ws.cell(row=h0, column=2, value=str(t.headers[t.chart[2]])).font = HDR_FONT
            for i, (lab, val) in enumerate(pairs_, start=1):
                ws.cell(row=h0 + i, column=1, value=lab)
                cc = ws.cell(row=h0 + i, column=2, value=val)
                cc.number_format = "0.0%" if t.chart[2] in t.pct else "#,##0.##"
            ch = BarChart()
            ch.type = "col" if t.chart[0] == "v" else "bar"
            ch.style = 2
            ch.title = title
            ch.legend = None
            ch.gapWidth = 40
            ch.add_data(Reference(ws, min_col=2, min_row=h0, max_row=h0 + len(pairs_)),
                        titles_from_data=True)
            ch.set_categories(Reference(ws, min_col=1, min_row=h0 + 1,
                                        max_row=h0 + len(pairs_)))
            ch.height = max(7.5, min(0.6 * len(pairs_) + 4, 19))
            ch.width = 21
            ws.add_chart(ch, f"D{row}")
            row = max(h0 + len(pairs_) + 2, h0 + int(ch.height / 0.53) + 2)
            n += 1

    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)  # Box resets synced files to read-only
    wb.save(OUT)
    print(f"wrote {OUT}")
    print(f"  {len(SECTIONS)} analysis sheets + Charts")
    print(f"  {sum(len(ts) for _, ts in SECTIONS)} tables, {n} charts")
    print(f"  unit of analysis: {N_MED} CHW x medicine observations, "
          f"{len(CHWS)} CHWs, {len(FACILITIES)} health centres")
    print(f"  availability {N_AVAIL}/{N_MANAGED} managed rows "
          f"({pct(N_AVAIL, N_MANAGED):.1%}); {N_OUT} stocked out")
    if other:
        print(f"  WARNING: {len(other)} comments matched no FB_THEMES rule - "
              f"see the Feedback sheet")


if __name__ == "__main__":
    write()
