"""
Full analysis of the 'Medicine stock assessment _DC' KoboToolbox project.

    python3 analyze_medicine_assessment.py

Reads   Medicine_stock_assessment_export.xlsx   (the synchronous XLSX export)
Writes  Medicine_assessment_analysis.xlsx       (one worksheet per analysis section)

To refresh the source from Kobo, set the export URL and (once access is locked down)
a token, then re-run - the script downloads only if the local file is missing:

    export KOBO_EXPORT_URL='https://kf.kobotoolbox.org/api/v2/assets/<uid>/export-settings/<uid>/data.xlsx'
    export KOBO_TOKEN='...'
    rm Medicine_stock_assessment_export.xlsx && python3 analyze_medicine_assessment.py

The URL is deliberately NOT hard-coded here: it is a capability that grants access to
the whole dataset, so it does not belong in a file that may be committed or shared.

UNIT OF ANALYSIS
The submission is not the unit of assessment in this project. 128 of 163 submissions
carry a single medicine while 18 carry 53-55, because three facilities (Rwinkwavu DH,
Kirinda CS, Ngeruka MHC) entered one medicine per submission. Rates are therefore
computed over the 1,359 medicine rows (facility x medicine observations), never over
submissions. 1,353 of those are distinct facility x medicine pairs; the 6 duplicated
pairs are listed in the Data quality sheet and left in, since they are real repeat
observations rather than a join artefact.

Percentages that aggregate quantities (fulfilment, request-vs-AMC) are computed from
summed quantities, NOT as a mean of per-row percentages - averaging ratios would weight
a 30-unit line equally with a 50,000-unit one.

openpyxl + the Python standard library only (no pandas).
"""

import datetime as dt
import math
import os
import sys
import statistics
import urllib.request
from collections import Counter, defaultdict

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SRC = "Medicine_stock_assessment_export.xlsx"
OUT = "Medicine_assessment_analysis.xlsx"
S_MED = "medicine_group"
# Sheet 1 is named after the FORM, so renaming the project in Kobo renames it
# (which is exactly what broke the prescription analysis on 2026-09-23). The
# repeat group keeps its XLSForm name, so sheet 1 is simply the other one.
S_SUB = next(n for n in openpyxl.load_workbook(SRC, read_only=True).sheetnames
             if n != S_MED)

# --- submission-level columns
PROV, DIST, SECT = "Province", "District", "Sector"
FAC = "Facility (select)"
LEVEL, FCLASS, RANK = "Health facility level", "facility_class", "hf_level_rank"
MGMT = "Facility management"
ENUM = "Please select your name"
FEEDBACK = "Please share any feedback, concerns, or recommendations related to the supply chain and other relevant areas."

# --- medicine_group columns (the form embeds ${medicine_name} in the labels)
MED = "medicine_name"
PROGRAM, DCLASS = "medicine_program", "drug_class_name"
PRICE = "unit_tariff_price"
APPLICABLE = "Is **${medicine_name}** applicable to this facility level as per EML?"
MANAGED = "Is **${medicine_name}** managed by the facility?"
WHY_NOT = "Why is **${medicine_name}** not managed by the facility?"
AVAIL = "Is **${medicine_name}** available on visit day?"
QTY_PHYS = "Quantity of **${medicine_name}** available"
QTY_SYS = "Quantity of **${medicine_name}** available in the system"
EMR = "Name of the electronical medical system used for recording"
ACCURACY = "inventory_accuracy_pct"
CONS = {m: f"{m} consumption total quantity of **{{medicine_name}}**".replace("{medicine_name}", "${medicine_name}")
        for m in ("June", "July", "August")}
AMC = "avg_monthly_consumption"
MOS = "stock_level_months"
STOCKCLASS = "stock_classification"
SO_NOW = "Current stock-out days of **${medicine_name}** (this visit)"
SO_MONTH = {m: f"{m} 2026 stock-out days of **{{medicine_name}}**".replace("{medicine_name}", "${medicine_name}")
            for m in ("June", "July", "August")}
SO_TOTAL, SO_RATE = "total_stockout_days", "stockout_rate_pct"
BILLED = "Total Quantity Billed **${medicine_name}**"  # the export header has a trailing space; load() strips it
RX_YEAR = "Prescriptions of **${medicine_name}** / year (RSSB)"
QTY_PER_RX = "qty_per_prescription"
RX_MISSED = "prescriptions_missed_stockout"
QTY_UNFILLED = "est_qty_unfilled_stockout"
RMS_CEN = "Is **${medicine_name}** available at RMS Central?"
RMS_REG = "Is **${medicine_name}** available at RMS Regional?"
DIST_FLAG = "distribution_nationwide"
REQ = {m: f"{m} 2026 quantity of **{{medicine_name}}** requested".replace("{medicine_name}", "${medicine_name}")
       for m in ("June", "July", "August")}
DEL = {m: f"{m} 2026 quantity of **{{medicine_name}}** delivered".replace("{medicine_name}", "${medicine_name}")
       for m in ("June", "July", "August")}
REQ_TOT, DEL_TOT, UNDEL_TOT = "total_requested_qty", "total_delivered_qty", "total_undelivered_qty"
FULFIL = "fulfilment_pct"
REQ_AMC, DEL_AMC = "request_vs_amc_pct", "delivery_vs_amc_pct"
GAP = "est_unserved_requisition_gap"
LOSS = "financial_loss"
UNMET = "overall_unmet_demand_rx"


# ---------------------------------------------------------------
# load
# ---------------------------------------------------------------
def fetch_if_missing():
    if os.path.exists(SRC):
        return
    url = os.environ.get("KOBO_EXPORT_URL")
    if not url:
        raise SystemExit(f"{SRC} not found and KOBO_EXPORT_URL is not set - see the module docstring.")
    req = urllib.request.Request(url)
    tok = os.environ.get("KOBO_TOKEN")
    if tok:
        req.add_header("Authorization", f"Token {tok}")
    print(f"downloading {SRC} from KOBO_EXPORT_URL ...")
    with urllib.request.urlopen(req, timeout=180) as r, open(SRC, "wb") as f:
        f.write(r.read())


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
    # Kobo's calculate fields can emit the literal string "NaN" (a division by a
    # zero or blank denominator). float() accepts "NaN"/"inf" happily, and one
    # such value silently turns every downstream sum into NaN, so treat a
    # non-finite result as missing - the same as a blank. Counted in Data quality.
    return f if math.isfinite(f) else None


def yes(v):
    return str(v).strip().lower() == "yes"


def no(v):
    return str(v).strip().lower() == "no"


def s(v):
    return "" if v is None else str(v).strip()


fetch_if_missing()
HDR_SUB, SUB = load(S_SUB)
HDR_MED, MEDROWS = load(S_MED)
PARENT = {r["_index"]: r for r in SUB}
for r in MEDROWS:
    r["_p"] = PARENT.get(r["_parent_index"])
MEDROWS = [r for r in MEDROWS if r["_p"]]

# ---------------------------------------------------------------
# facilities removed from every analysis
# ---------------------------------------------------------------
# User decision 2026-10-18: "MPENGE HP" (Musanze, health post; 1 visit, 20
# medicine rows, no prescription or CHW record) is dropped here, before any
# cleaning rule, count or breakdown reads the rows, so every output built on
# this module leaves it out. Matched on the facility name, case-insensitive.
EXCLUDED_FACILITIES = {"MPENGE HP"}
_excl = {r["_index"] for r in SUB
         if s(r.get(FAC)).upper() in EXCLUDED_FACILITIES}
EXCLUDED_COUNTS = (len(_excl),
                   sum(1 for r in MEDROWS if r["_parent_index"] in _excl))
SUB = [r for r in SUB if r["_index"] not in _excl]
MEDROWS = [r for r in MEDROWS if r["_parent_index"] not in _excl]


# ---------------------------------------------------------------
# de-duplicate facility x medicine
# ---------------------------------------------------------------
# A facility submitted twice on the same day leaves the same medicine assessed
# twice, and the pair is the unit of analysis - so a duplicate double-counts
# the observation and, where the two visits disagree, puts contradictory
# quantities into the same total.
#
# Rule: keep ONE row per (facility, medicine), preferring
#   1. the row that actually carries an assessment - a "not managed" row with
#      nothing else filled in loses to a managed row with quantities, whichever
#      came first; then
#   2. the LATER submission, on the basis that a repeat visit supersedes the
#      first.
#
# This is applied in the analysis, NOT by editing the export: the merged
# export is rebuilt from both servers on every fetch, so an edit there would be
# silently undone. Every dropped row is printed, and the pairs where the two
# records materially disagree are listed on the Data quality sheet - the rule
# picks one, but only the field team can say which visit was right.
DEDUPE_FIELDS = (MANAGED, AVAIL, QTY_PHYS, QTY_SYS, AMC, MOS, STOCKCLASS,
                 SO_NOW, SO_TOTAL, REQ_TOT, DEL_TOT, LOSS)


def _filled(r):
    return sum(1 for c in DEDUPE_FIELDS if not blank(r.get(c)))


def dedupe_med_rows(rows):
    """-> kept rows, and a record of what was dropped."""
    groups = defaultdict(list)
    for r in rows:
        groups[(s(r["_p"][FAC]), s(r[MED]))].append(r)

    kept, dropped = [], []
    for key, rs in groups.items():
        if len(rs) == 1:
            kept.append(rs[0])
            continue
        # an informative row beats an empty one; otherwise the later submission
        ranked = sorted(rs, key=lambda r: (_filled(r) > 1,
                                           s(r["_p"]["_submission_time"])), reverse=True)
        keep, rest = ranked[0], ranked[1:]
        kept.append(keep)
        differs = [c for c in DEDUPE_FIELDS
                   if any(str(o.get(c)) != str(keep.get(c)) for o in rest)]
        for o in rest:
            dropped.append({
                "facility": key[0], "medicine": key[1],
                "kept_sub": s(keep["_p"]["_submission_time"])[:16],
                "kept_id": keep["_p"].get("_id"), "kept_filled": _filled(keep),
                "dropped_sub": s(o["_p"]["_submission_time"])[:16],
                "dropped_id": o["_p"].get("_id"), "dropped_filled": _filled(o),
                "differs": ", ".join(differs) if differs else "",
                "identical": not differs,
                "kept_row": keep, "dropped_row": o})
    return kept, dropped


MEDROWS, DUPES_DROPPED = dedupe_med_rows(MEDROWS)
# material disagreement = the two visits recorded different facts, so the
# choice changes a figure rather than just removing a copy
DUPES_CONFLICTING = [d for d in DUPES_DROPPED if not d["identical"]]


# ---------------------------------------------------------------
# cleaning rules S6-S8, applied before anything reads these rows
# ---------------------------------------------------------------
# These live here rather than in the cleaning pipeline so that ONE set of
# corrected rows feeds the pipeline's exports, every analytics script and
# every dashboard. Putting them downstream would let the published figures
# and the cleaned dataset disagree.
#
# The helper modules sit in analytics/, which is not on the path when this
# file is imported directly, so the import is guarded: if they cannot be
# found the rules are skipped and the counters stay at zero rather than the
# whole analysis failing.
EML_FIXED, ZERO_FIXED, EMR_CHANGED, EMR_NUMERIC = [], [], [], []
STRENGTH_RECODED = []
ZONE_RECODED = []           # S10
EML_OPPOSITE, EML_UNMATCHED = [], {}

try:
    sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "analytics"))
    import eml_list as _EML
    import emr_systems as _EMRSYS
    import molkey as _MK
    import mft_zoning as _ZN
except Exception:            # helpers unavailable - leave the data as read
    _EML = _EMRSYS = _MK = _ZN = None

if _EML is not None:
    # S6 - the form asks whether the medicine is applicable at this level
    # "as per EML" and the collector answers from memory. The EML list is the
    # authority. ONE DIRECTION ONLY: No -> Yes where the list holds the
    # medicine at that level. The reverse is recorded, not changed, because a
    # facility managing something outside its level is a real finding.
    _by_ing, _names = _EML.load_eml()
    _seen = {}
    for r in MEDROWS:
        _k = (s(r[MED]), s(r["_p"][LEVEL]))
        if _k not in _seen:
            _seen[_k] = _EML.applicable(_k[0], _k[1], _by_ing, _names)
        _verdict, _how, _levels, _prods = _seen[_k]
        r["_eml_levels"] = ", ".join(sorted(_levels))
        r["_eml_match"] = _how
        r["_applicable_was"] = ""
        if _verdict is None:
            if _how == "no name match":
                EML_UNMATCHED[s(r[MED])] = EML_UNMATCHED.get(s(r[MED]), 0) + 1
            continue
        if _verdict and not yes(r[APPLICABLE]):
            r["_applicable_was"] = s(r[APPLICABLE]) or "(blank)"
            r[APPLICABLE] = "Yes"
            EML_FIXED.append(r)
        elif yes(r[APPLICABLE]) and not _verdict:
            EML_OPPOSITE.append(r)

# S7 - available = Yes with nothing on the shelf contradicts itself. The
# quantity is counted or read off a stock card, so it is taken as the true
# answer. The opposite case (No with stock) does not occur in the data.
for r in MEDROWS:
    r["_available_was"] = ""
    if yes(r[AVAIL]) and num(r[QTY_PHYS]) == 0:
        r["_available_was"] = s(r[AVAIL])
        r[AVAIL] = "No"
        ZERO_FIXED.append(r)

# S9 - one medicine held at two or more strengths, every strength MANAGED
# and answered, at least one on the shelf and another not. The facility
# works with the strength it has; the empty one is recorded as NOT MANAGED
# rather than as a stock-out (decision of 2026-10-15). The managed and
# available answers are kept in _managed_was / _available_was, and a reason
# naming the strength actually used is written into WHY_NOT where it was
# blank. Strengths are grouped by molkey.strength_key(), which also treats
# the two thiamine products as one medicine.
# A recoded row's other answers (availability through the requisition
# fields) are blanked - once a row is NOT MANAGED its other responses are
# not considered (user decision 2026-11-03; S10 does the same) - and kept
# in r["_blanked"] so the as-recorded extracts can read them.
# S9_MIN_DAYS / S9_EXCLUDE can narrow the rule (tried 2026-11-03 with 60
# days and DHAP/AL excluded, then withdrawn): None / empty = every case.
S9_MIN_DAYS = None
S9_EXCLUDE = set()
_S9_DAYS = [f"{m} 2026 stock-out days of **${{medicine_name}}**"
            for m in ("June", "July", "August")]


def blank_rest(r, last="_index"):
    """Blank every answer after AVAIL up to `last`; keep them in _blanked."""
    keys = list(r)
    rest = keys[keys.index(AVAIL) + 1:keys.index(last)]
    r["_blanked"] = {k: r[k] for k in rest}
    for k in rest:
        r[k] = None


if _MK is not None:
    _grp = {}
    for r in MEDROWS:
        _grp.setdefault((s(r["_p"][FAC]), _MK.strength_key(r[MED])),
                        []).append(r)
    for _k, _rs in _grp.items():
        if _k[1] in S9_EXCLUDE:
            continue
        _m = [r for r in _rs if yes(r[MANAGED]) and not blank(r[AVAIL])]
        if len({s(r[MED]) for r in _m}) < 2:
            continue
        _on = [r for r in _m if yes(r[AVAIL])]
        _off = [r for r in _m if no(r[AVAIL])]
        if not (_on and _off):
            continue
        _used = " / ".join(sorted(s(r[MED]) for r in _on))
        for r in _off:
            if (S9_MIN_DAYS is not None and
                    sum(num(r[k]) or 0 for k in _S9_DAYS) <= S9_MIN_DAYS):
                continue
            r["_managed_was"] = s(r[MANAGED])
            r["_available_was"] = r.get("_available_was") or s(r[AVAIL])
            r[MANAGED] = "No"
            r[AVAIL] = None
            if blank(r[WHY_NOT]):
                r[WHY_NOT] = f"Recoded S9: manages {_used}"
            blank_rest(r)
            STRENGTH_RECODED.append(r)

# S10 - a first-line antimalarial (AL / ASPY / DHAP) recorded MANAGED and
# NOT AVAILABLE, in a district whose MFT zone allocates another one
# (analytics/mft_zoning.py), is recoded to NOT MANAGED: the empty shelf is
# the zoning, not a stock-out (decision of 2026-11-02). Answers kept in
# _managed_was / _available_was, as S9.
if _ZN is not None:
    for r in MEDROWS:
        if (yes(r[MANAGED]) and no(r[AVAIL])
                and _ZN.out_of_zone(r[MED], r["_p"][DIST])):
            r["_managed_was"] = s(r[MANAGED])
            r["_available_was"] = r.get("_available_was") or s(r[AVAIL])
            r[MANAGED] = "No"
            r[AVAIL] = None
            if blank(r[WHY_NOT]):
                r[WHY_NOT] = ("Recoded S10: out of MFT zone ("
                              + _ZN.ZONE.get(s(r["_p"][DIST]), "") + ")")
            blank_rest(r)
            ZONE_RECODED.append(r)

# S11 - faith-based facilities that do not provide family planning: every FP
# product at the facilities below is recoded to NOT MANAGED and its other
# answers blanked, whatever was recorded (user decision 2026-11-05). Names
# are matched exactly - Kamubuga CS (government) is a different facility.
FB_NO_FP = {"Mubuga CS", "Birambo CS"}
FP_RECODED = []
for r in MEDROWS:
    if s(r["_p"][FAC]) in FB_NO_FP and s(r[PROGRAM]) == "FP":
        r["_managed_was"] = r.get("_managed_was") or s(r[MANAGED])
        r["_available_was"] = r.get("_available_was") or s(r[AVAIL])
        r[MANAGED] = "No"
        r[AVAIL] = None
        r[WHY_NOT] = "Recoded S11: faith-based facility, FP not provided"
        if "_blanked" not in r:
            blank_rest(r)
        FP_RECODED.append(r)

if _EMRSYS is not None:
    # S8 - the recording system name is free text; 162 spellings for seven
    # systems, and 32 rows holding a bare number. Normalised into its own
    # column; the recorded value is left untouched in EMR.
    for r in MEDROWS:
        _raw = s(r[EMR])
        _name, _note = _EMRSYS.canonical(_raw)
        r["_emr_raw"] = _raw
        r["_emr_clean"] = _name
        r["_emr_note"] = _note
        if _raw and _note == "numeric value removed":
            EMR_NUMERIC.append(r)
        if _raw and _name != _raw:
            EMR_CHANGED.append(r)

N_SUB, N_MED = len(SUB), len(MEDROWS)
FACILITIES = {s(r[FAC]) for r in SUB}

# Profile of how the repeat group was filled, for the "unit of analysis" note.
# Some facilities submit one medicine per submission, others the whole basket,
# so these counts must be measured per export rather than written down once.
_per_sub = Counter(r["_parent_index"] for r in MEDROWS)
SUB_ONE_MED = sum(1 for n in _per_sub.values() if n == 1)
_bulk = [n for n in _per_sub.values() if n > 1]
SUB_MANY_MED = len(_bulk)
SUB_BULK_LO, SUB_BULK_HI = (min(_bulk), max(_bulk)) if _bulk else (0, 0)
# Which facilities drive the single-medicine submissions, named from the data.
_one_fac = Counter(s(r["_p"][FAC]) for r in MEDROWS if _per_sub[r["_parent_index"]] == 1)
SUB_ONE_MED_TOP = ", ".join(f"{f} ({n})" for f, n in _one_fac.most_common(3)) or "none"


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
               if isinstance(r[vc], (int, float)) and r[vc] is not None]
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

    return [label, k, round(sum(v), 1), round(sum(v) / k, 2), v[0],
            round(q(.25), 2), round(statistics.median(v), 2), round(q(.75), 2), v[-1]]


STAT_HDR = ["Measure", "n", "Sum", "Mean", "Min", "P25", "Median", "P75", "Max"]


# ---------------------------------------------------------------
# the group profile - one consistent measure set, reused for every dimension
# ---------------------------------------------------------------
PROFILE_HDR = ["Group", "Observations", "Facilities", "Medicines", "Managed",
               "Availability answered", "Available", "Available %",
               "Stock out", "Understock", "Plan stock", "Overstock",
               "Median months of stock", "Stock-out days (total)", "Stock-out rate %",
               "Qty requested", "Qty delivered", "Fulfilment %", "Financial loss (RWF)"]


def profile(rows, keyfn, sort_key=None):
    g = defaultdict(list)
    for r in rows:
        g[keyfn(r)].append(r)
    out = []
    for k, rs in g.items():
        answered = [r for r in rs if not blank(r[AVAIL])]
        av = sum(1 for r in answered if yes(r[AVAIL]))
        cls = Counter(s(r[STOCKCLASS]) for r in rs if not blank(r[STOCKCLASS]))
        mos = [num(r[MOS]) for r in rs]
        mos = [x for x in mos if x is not None]
        req = sum(num(r[REQ_TOT]) or 0 for r in rs)
        dlv = sum(num(r[DEL_TOT]) or 0 for r in rs)
        sod = sum(num(r[SO_TOTAL]) or 0 for r in rs)
        sor = [num(r[SO_RATE]) for r in rs]
        sor = [x for x in sor if x is not None]
        out.append([
            k, len(rs), len({s(r["_p"][FAC]) for r in rs}), len({s(r[MED]) for r in rs}),
            sum(1 for r in rs if yes(r[MANAGED])), len(answered), av,
            av / len(answered) if answered else None,
            cls.get("Stock out", 0), cls.get("Understock", 0),
            cls.get("Plan stock", 0), cls.get("Overstock", 0),
            round(statistics.median(mos), 2) if mos else None,
            round(sod) if sod else 0,
            round(sum(sor) / len(sor), 1) if sor else None,
            round(req), round(dlv),
            dlv / req if req else None,
            round(sum(num(r[LOSS]) or 0 for r in rs)),
        ])
    out.sort(key=sort_key or (lambda r: -r[1]))
    return out


def total_row(rows, label="ALL"):
    r = profile(rows, lambda _: label)
    return r[0]


# ---------------------------------------------------------------
# 1. README + overview
# ---------------------------------------------------------------
mg_per_sub = Counter(r["_parent_index"] for r in MEDROWS)
answered_all = [r for r in MEDROWS if not blank(r[AVAIL])]
avail_all = sum(1 for r in answered_all if yes(r[AVAIL]))
unavail = [r for r in MEDROWS if no(r[AVAIL])]
req_all = sum(num(r[REQ_TOT]) or 0 for r in MEDROWS)
del_all = sum(num(r[DEL_TOT]) or 0 for r in MEDROWS)
loss_all = sum(num(r[LOSS]) or 0 for r in MEDROWS)
stimes = sorted(s(r["_submission_time"]) for r in SUB)

# --- figures the Findings sheet quotes. Computed, never written down: this
# export is a live interim cut and every one of them moves between refreshes.
CLASS_N = Counter(s(r[STOCKCLASS]) for r in MEDROWS if not blank(r[STOCKCLASS]))
N_CLASSIFIED = sum(CLASS_N.values())
ATTRIB_N = Counter(s(r[DIST_FLAG]) for r in unavail if not blank(r[DIST_FLAG]))
N_APPLICABLE = sum(1 for r in MEDROWS if yes(r[APPLICABLE]))
N_NOT_APPLICABLE = sum(1 for r in MEDROWS if no(r[APPLICABLE]))
N_MANAGED = sum(1 for r in MEDROWS if yes(r[MANAGED]))
N_NOT_MANAGED = sum(1 for r in MEDROWS if no(r[MANAGED]))
# Applicable per EML yet not managed - the reconciliation gap.
N_EML_GAP = sum(1 for r in MEDROWS if yes(r[APPLICABLE]) and no(r[MANAGED]))
N_LOSS_ROWS = sum(1 for r in MEDROWS if num(r[LOSS]) is not None)
N_VERSIONS = len({s(r.get("__version__")) for r in SUB})
N_GEO = sum(1 for r in SUB if not blank(r.get("_start-geopoint_latitude")))
N_NO_ENUM = sum(1 for r in SUB if blank(r.get(ENUM)))
N_USERNAME_BROKEN = sum(1 for r in SUB
                        if s(r.get("username")) == "username not found")
# Columns that exist in the instrument but hold no data in this export.
EMPTY_COLS = [c for c in HDR_MED
              if c and all(blank(r.get(c)) for r in MEDROWS)
              and not c.startswith(("_", "**"))]
# Medicine observations per facility, to show coverage evenness.
_per_fac = Counter(s(r["_p"][FAC]) for r in MEDROWS)
FAC_OBS_LO, FAC_OBS_HI = (min(_per_fac.values()), max(_per_fac.values())) \
    if _per_fac else (0, 0)


def cls(name):
    return CLASS_N.get(name, 0)


def _colmax(col):
    v = [num(r.get(col)) for r in MEDROWS]
    v = [x for x in v if x is not None]
    return max(v) if v else 0


SO_CUR_MAX = _colmax(SO_NOW)
SO_TOT_MAX = _colmax(SO_TOTAL)
N_MED_NAMES = len({s(r[MED]) for r in MEDROWS if not blank(r[MED])})

readme = [
    Table("What this workbook is", ["Item", "Value"], [
        ["Source", f"{SRC} - KoboToolbox synchronous XLSX export"],
        ["Kobo project", "Medicine stock assessment _DC"],
        ["Submissions", N_SUB],
        ["Medicine observations (facility x medicine)", N_MED],
        ["Distinct facility x medicine pairs", len({(s(r['_p'][FAC]), s(r[MED])) for r in MEDROWS})],
        ["Facilities", len(FACILITIES)],
        ["Medicines in the instrument", len({s(r[MED]) for r in MEDROWS})],
        ["Submission window", f"{stimes[0]} .. {stimes[-1]}"],
        ["Analysis generated", dt.datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Generated by", "analyze_medicine_assessment.py (openpyxl + stdlib)"],
    ]),
    Table("Unit of analysis - read before quoting any rate", ["Point", "Detail"], [
        ["The submission is NOT the assessment unit",
         f"{SUB_ONE_MED} of {len(SUB)} submissions carry a single medicine, while "
         f"{SUB_MANY_MED} carry {SUB_BULK_LO}-{SUB_BULK_HI}: some facilities entered one "
         f"medicine per submission and others the whole basket at once, so submission "
         f"counts say more about data-entry style than about coverage. "
         f"Most single-medicine submissions: {SUB_ONE_MED_TOP}."],
        ["Rates use medicine rows",
         f"All percentages below are over the {N_MED} medicine observations, or over the "
         "subset that answered the relevant question."],
        ["Quantity percentages are ratios of sums",
         "Fulfilment %, request-vs-AMC and delivery-vs-AMC are computed from summed "
         "quantities, not as the mean of per-row percentages - a 30-unit line must not "
         "weigh the same as a 50,000-unit line."],
        ["Denominators shift with skip logic",
         f"Availability is asked only where the medicine is managed "
         f"({len(answered_all):,} of {N_MED:,}); quantity and consumption only where "
         f"available; RMS and distribution questions only where unavailable "
         f"({len(unavail):,}). Each table states its own n."],
        ["Derived fields arrive as text",
         "Kobo calculates (avg_monthly_consumption, fulfilment_pct, financial_loss, ...) "
         "export as strings. They are converted to numbers here; anything non-numeric is "
         "treated as missing and counted in Data quality."],
    ]),
    Table("Refreshing the source data", ["Step", "Command"], [
        ["1. Pull the latest export", "python3 fetch_kobo_export.py stock"],
        ["2. Rebuild this workbook", "python3 analyze_medicine_assessment.py"],
        ["Both studies at once", "python3 fetch_kobo_export.py"],
        ["If access is later locked down", "export KOBO_API_TOKEN='<your Kobo API key>'"],
    ], note="fetch_kobo_export.py holds the asset and export-setting uids and writes "
            f"straight to {SRC}. It sends a token only when one is found, because a "
            "publicly shared export rejects a bad Authorization header that it would "
            "have served with no header at all."),
]

overview = [
    Table("Headline figures", ["Measure", "Value"], [
        ["Facilities assessed", len(FACILITIES)],
        ["Provinces / districts covered", f"{len({s(r[PROV]) for r in SUB})} / {len({s(r[DIST]) for r in SUB})}"],
        ["Medicine observations", N_MED],
        ["Applicable to facility level per EML", sum(1 for r in MEDROWS if yes(r[APPLICABLE]))],
        ["Managed by the facility", sum(1 for r in MEDROWS if yes(r[MANAGED]))],
        ["Not managed by the facility", sum(1 for r in MEDROWS if no(r[MANAGED]))],
        ["Availability answered", len(answered_all)],
        ["Available on visit day", avail_all],
        ["Not available on visit day", len(unavail)],
        ["Availability rate", avail_all / len(answered_all) if answered_all else None],
        ["Stock-out days recorded (Jun-Aug, total)",
         round(sum(num(r[SO_TOTAL]) or 0 for r in MEDROWS))],
        ["Quantity requested (Jun-Aug)", round(req_all)],
        ["Quantity delivered (Jun-Aug)", round(del_all)],
        ["Quantity undelivered", round(req_all - del_all)],
        ["Fulfilment rate (delivered / requested)", del_all / req_all if req_all else None],
        ["Estimated financial loss (RWF)", round(loss_all)],
    ], note=f"Availability rate is over the {len(answered_all):,} observations where the "
            f"medicine is managed by the facility and the question was therefore asked."),
    Table("Stock classification across all classified observations",
          ["Classification", "Observations", "%"],
          freq([r for r in MEDROWS if not blank(r[STOCKCLASS])], lambda r: s(r[STOCKCLASS])),
          note=f"Classified: {sum(1 for r in MEDROWS if not blank(r[STOCKCLASS]))} of {N_MED} "
               "observations. Thresholds recorded in the form: understock < 1 month, "
               "overstock > 2 months, over a 90-day window.",
          chart=("h", 0, 1, None)),
    Table("Where the gap arises - the supply chain attribution",
          ["Point in the chain", "Observations", "%"],
          freq(unavail, lambda r: s(r[DIST_FLAG]) or "(not answered)"),
          note=f"Asked only of the {len(unavail):,} observations that were unavailable on "
               f"visit day. 'Distribution issue' means the medicine was in stock at RMS "
               f"but not at the facility; 'Nationwide stock-out' means RMS did not have "
               f"it either."),
]

# ---------------------------------------------------------------
# 2. coverage
# ---------------------------------------------------------------
fac_prov = defaultdict(set)
fac_dist = defaultdict(set)
for r in SUB:
    fac_prov[s(r[PROV])].add(s(r[FAC]))
    fac_dist[s(r[PROV])].add(s(r[DIST]))
prov_rows = []
for p in sorted(fac_prov, key=lambda p: -len(fac_prov[p])):
    subs = [r for r in SUB if s(r[PROV]) == p]
    meds = [r for r in MEDROWS if s(r["_p"][PROV]) == p]
    prov_rows.append([p, len(fac_prov[p]), len(fac_dist[p]), len(subs), len(meds)])

fac_rows = []
for f in sorted(FACILITIES):
    subs = [r for r in SUB if s(r[FAC]) == f]
    meds = [r for r in MEDROWS if s(r["_p"][FAC]) == f]
    p0 = subs[0]
    fac_rows.append([s(p0[PROV]), s(p0[DIST]), s(p0[SECT]), f, s(p0[LEVEL]), s(p0[MGMT]),
                     len(subs), len(meds)])
fac_rows.sort(key=lambda r: (r[0], r[1], r[3]))

coverage = [
    Table("Facilities that have submitted, by province",
          ["Province", "Facilities", "Districts", "Submissions", "Medicine observations"],
          prov_rows,
          total=["TOTAL", len(FACILITIES), len({s(r[DIST]) for r in SUB}), N_SUB, N_MED],
          note="Facility counts sum cleanly to the total - no facility appears in two "
               "provinces. Submissions are NOT comparable across provinces: East and West "
               "are inflated by the three one-medicine-per-submission facilities. Medicine "
               "observations are the comparable column, and they are far more even "
               "(387/331/326/207/108) than the submission counts suggest.",
          chart=("v", 0, 1, None)),
    Table("Facility roster",
          ["Province", "District", "Sector", "Facility", "Level", "Management",
           "Submissions", "Medicine observations"], fac_rows),
    Table("Facilities by level", ["Level", "Facilities", "%"],
          freq(SUB, lambda r: s(r[LEVEL]), denom=len(FACILITIES),
               sort_by_count=True) and
          [[k, len({s(r[FAC]) for r in SUB if s(r[LEVEL]) == k}),
            len({s(r[FAC]) for r in SUB if s(r[LEVEL]) == k}) / len(FACILITIES)]
           for k in sorted({s(r[LEVEL]) for r in SUB},
                           key=lambda k: -len({s(r[FAC]) for r in SUB if s(r[LEVEL]) == k}))],
          total=["TOTAL", len(FACILITIES), 1.0]),
    Table("Facilities by management", ["Management", "Facilities", "%"],
          [[k, len({s(r[FAC]) for r in SUB if s(r[MGMT]) == k}),
            len({s(r[FAC]) for r in SUB if s(r[MGMT]) == k}) / len(FACILITIES)]
           for k in sorted({s(r[MGMT]) for r in SUB},
                           key=lambda k: -len({s(r[FAC]) for r in SUB if s(r[MGMT]) == k}))],
          total=["TOTAL", len(FACILITIES), 1.0]),
    Table("Medicine observations per facility",
          ["Observations in the facility", "Facilities"],
          [[k, v] for k, v in sorted(Counter(
              len([r for r in MEDROWS if s(r["_p"][FAC]) == f]) for f in FACILITIES).items())],
          note="Coverage is even: every facility carries 42-56 medicine observations, and "
               "each province covers all 55 medicines. The wide spread in submission counts "
               "is purely data-entry style - the three one-medicine-per-submission "
               "facilities still ended up assessing the full instrument. Facility-level "
               "rates are therefore comparable."),
    Table("Submissions by day", ["Date", "Submissions"],
          [[k, v] for k, v in sorted(Counter(s(r["_submission_time"])[:10] for r in SUB).items())],
          note="A two-day window - this is an active rollout, not a completed round."),
    Table("Data collectors", ["Enumerator", "Submissions", "Facilities", "Medicine observations"],
          sorted([[k,
                   sum(1 for r in SUB if s(r[ENUM]) == k),
                   len({s(r[FAC]) for r in SUB if s(r[ENUM]) == k}),
                   len([r for r in MEDROWS if s(r["_p"][ENUM]) == k])]
                  for k in {s(r[ENUM]) for r in SUB}], key=lambda r: -r[3]),
          note="Blank enumerator = the question was left empty on that submission."),
    Table("Form versions in use", ["__version__", "Submissions"],
          [[k, v] for k, v in Counter(s(r["__version__"]) for r in SUB).most_common()],
          note="Four versions are in circulation. The newest adds May stock-out days and "
               "March-May requisition columns, which is why those columns are empty for "
               "every submission collected on an older version."),
]

# ---------------------------------------------------------------
# 3. data quality
# ---------------------------------------------------------------
def missingness(hdr, rows, skip_selectall=None):
    out = []
    for h in hdr:
        if h.startswith("_") or h.startswith("meta/") or h.startswith("**") or h == "":
            continue
        if skip_selectall and h.startswith(skip_selectall + "/"):
            continue
        m = sum(1 for r in rows if blank(r.get(h)))
        out.append([h, len(rows) - m, m, m / len(rows)])
    return out


SELALL = "Which medicines are being assessed at this facility? (Select all that apply)"
empty_sub = [r[0] for r in missingness(HDR_SUB, SUB, SELALL) if r[1] == 0]
empty_med = [r[0] for r in missingness(HDR_MED, MEDROWS) if r[1] == 0]

pairs = defaultdict(list)
for r in MEDROWS:
    pairs[(s(r["_p"][FAC]), s(r[MED]))].append(r)
dup_pairs = {k: v for k, v in pairs.items() if len(v) > 1}   # after de-duplication: expected 0
# What de-duplication removed, so the cleaning is auditable rather than invisible.
dup_rows = []
for d in sorted(DUPES_DROPPED, key=lambda d: (d["facility"], d["medicine"])):
    k, o = d["kept_row"], d["dropped_row"]
    dup_rows.append([
        d["facility"], d["medicine"],
        f'{d["kept_sub"]} (id {d["kept_id"]}, {d["kept_filled"]}/12 filled)',
        f'{s(k[AVAIL]) or "(not asked)"} / qty {s(k[QTY_PHYS]) or "-"}',
        f'{d["dropped_sub"]} (id {d["dropped_id"]}, {d["dropped_filled"]}/12 filled)',
        f'{s(o[AVAIL]) or "(not asked)"} / qty {s(o[QTY_PHYS]) or "-"}',
        "identical copy" if d["identical"] else "DISAGREE: " + d["differs"]])

nonnum = []
for c in (AMC, MOS, SO_TOTAL, SO_RATE, ACCURACY, REQ_TOT, DEL_TOT, UNDEL_TOT, FULFIL,
          REQ_AMC, DEL_AMC, GAP, LOSS, UNMET, PRICE, QTY_PER_RX, RX_MISSED, QTY_UNFILLED):
    present = [r for r in MEDROWS if not blank(r[c])]
    bad = [r for r in present if num(r[c]) is None]
    nonnum.append([c, len(present), len(bad)])

skip_checks = [
    ["Availability asked but medicine not managed",
     sum(1 for r in MEDROWS if not yes(r[MANAGED]) and not blank(r[AVAIL])), "expected 0"],
    ["Medicine managed but availability blank",
     sum(1 for r in MEDROWS if yes(r[MANAGED]) and blank(r[AVAIL])), "expected 0"],
    ["Quantity recorded but not available on visit day",
     sum(1 for r in MEDROWS if not yes(r[AVAIL]) and not blank(r[QTY_PHYS])), "expected 0"],
    ["Available but quantity blank",
     sum(1 for r in MEDROWS if yes(r[AVAIL]) and blank(r[QTY_PHYS])), "expected 0"],
    ["RMS question answered but medicine was available",
     sum(1 for r in MEDROWS if not no(r[AVAIL]) and not blank(r[RMS_CEN])), "expected 0"],
    ["Unavailable but RMS Central blank",
     sum(1 for r in MEDROWS if no(r[AVAIL]) and blank(r[RMS_CEN])), "expected 0"],
    ["Unavailable but current stock-out days blank",
     sum(1 for r in MEDROWS if no(r[AVAIL]) and blank(r[SO_NOW])), "investigate"],
    ["Not managed but no reason given",
     sum(1 for r in MEDROWS if no(r[MANAGED]) and blank(r[WHY_NOT])), "investigate"],
]

neg_acc = [r for r in MEDROWS if (num(r[ACCURACY]) or 0) < 0]
over_acc = [r for r in MEDROWS if (num(r[ACCURACY]) or 0) > 100]
neg_gap = [r for r in MEDROWS if (num(r[GAP]) or 0) < 0]
del_gt_req = [r for r in MEDROWS
              if (num(r[DEL_TOT]) or 0) > (num(r[REQ_TOT]) or 0) and not blank(r[REQ_TOT])]

quality = [
    Table("Structural checks", ["Check", "Result", "Expected"], [
        ["Submissions", N_SUB, ""],
        ["Medicine rows with a resolvable parent submission", f"{N_MED} / {N_MED}", "all"],
        ["Distinct facility x medicine pairs", len(pairs), ""],
        ["Pairs assessed more than once (after de-duplication)", len(dup_pairs), "0"],
        ["Duplicate rows removed by this script", len(DUPES_DROPPED), ""],
        ["...of which the two visits disagreed", len(DUPES_CONFLICTING), "confirm with the collector"],
        ["Submissions carrying exactly one medicine", sum(1 for v in mg_per_sub.values() if v == 1), ""],
        ["Submissions carrying 50+ medicines", sum(1 for v in mg_per_sub.values() if v >= 50), ""],
        ["Submissions with no medicine rows", N_SUB - len(mg_per_sub), "0"],
    ]),
    Table("Duplicate facility x medicine assessments removed by this script",
          ["Facility", "Medicine", "KEPT submission", "Kept: available / qty",
           "DROPPED submission", "Dropped: available / qty",
           "Did the two visits agree?"], dup_rows,
          note=f"{len(DUPES_DROPPED)} duplicate row(s) removed of {len(pairs) + len(DUPES_DROPPED):,} "
               f"facility x medicine observations. The pair is the unit of analysis, so a "
               f"duplicate double-counts the observation. The rule keeps the row carrying an "
               f"actual assessment, then the later submission; it is applied in the analysis, "
               f"not in the export, because the export is rebuilt from both servers on every "
               f"fetch. {len(DUPES_CONFLICTING)} of them DISAGREE with the row kept - the rule "
               f"picked one, but only the collector can say which visit was right, so those "
               f"medicines should be confirmed before they are reported on."),
    Table("Skip-logic consistency", ["Check", "Rows", "Expected"], skip_checks,
          note="Non-zero counts on the 'expected 0' rows would mean the form's relevance "
               "rules were bypassed. The two 'investigate' rows are legitimate gaps to chase."),
    Table("Columns that are completely empty", ["Sheet", "Column"],
          [[S_SUB, c] for c in empty_sub] + [[S_MED, c] for c in empty_med],
          note="These exist in the newest form version but hold no data yet - either the "
               "question was added after most submissions, or it is never reachable. "
               "'username' is separately broken: it reads 'username not found' on all 163 "
               "submissions, so enumerator identity survives only in the name question."),
    Table("Derived fields - values that will not parse as numbers",
          ["Field", "Rows with a value", "Non-numeric"], nonnum,
          note="Kobo exports calculate fields as text. Every one of these parses cleanly "
               "here, so the conversion is safe - but any tool reading the export directly "
               "must cast them, or it will silently sort '1000' before '9'."),
    Table("Value-range flags", ["Flag", "Rows", "What it means"], [
        ["Negative inventory accuracy", len(neg_acc),
         "system quantity exceeds the physical count"],
        ["Inventory accuracy above 100%", len(over_acc),
         "physical count exceeds the system quantity"],
        ["Negative unserved requisition gap", len(neg_gap),
         "delivered more than the modelled requirement"],
        ["Delivered more than requested", len(del_gt_req),
         "over-delivery, or a data-entry transposition"],
        ["GPS captured", sum(1 for r in SUB if not blank(r.get("_start-geopoint_latitude"))),
         f"of {N_SUB} submissions - location is not being captured reliably"],
    ]),
    Table("Missingness - submission sheet",
          ["Column", "Present", "Missing", "% missing"], missingness(HDR_SUB, SUB, SELALL),
          note=f"The {sum(1 for c in HDR_SUB if c.startswith(SELALL + '/'))} "
               "'select all that apply' dummy columns are excluded from this list. Note "
               "that several medicines appear twice in that choice list (Atazanavir, "
               "Abacavir, Tranexamic Acid, Caffeine Citrate, Carbetocin, Insulin), which "
               "is why the dummies do not reconcile to the medicine rows."),
    Table("Missingness - medicine_group sheet",
          ["Column", "Present", "Missing", "% missing"], missingness(HDR_MED, MEDROWS),
          note="High missingness on most fields is the skip logic working, not data loss - "
               "compare each against the Skip-logic consistency table above."),
]

# ---------------------------------------------------------------
# 4. medicine coverage / applicability
# ---------------------------------------------------------------
med_cov = []
for m in sorted({s(r[MED]) for r in MEDROWS}):
    rs = [r for r in MEDROWS if s(r[MED]) == m]
    ans = [r for r in rs if not blank(r[AVAIL])]
    av = sum(1 for r in ans if yes(r[AVAIL]))
    med_cov.append([m, s(rs[0][PROGRAM]), s(rs[0][DCLASS]), num(rs[0][PRICE]),
                    len(rs), len({s(r["_p"][FAC]) for r in rs}),
                    sum(1 for r in rs if yes(r[APPLICABLE])),
                    sum(1 for r in rs if yes(r[MANAGED])),
                    len(ans), av, av / len(ans) if ans else None])
med_cov.sort(key=lambda r: (r[10] is not None, r[10] if r[10] is not None else 9, -r[4]))

why_rows = freq([r for r in MEDROWS if not blank(r[WHY_NOT])],
                lambda r: s(r[WHY_NOT])[:90], top=40)

medcov = [
    Table("Medicine roster with availability",
          ["Medicine", "Programme", "Drug class", "Unit tariff price (RWF)",
           "Observations", "Facilities", "Applicable per EML", "Managed",
           "Availability answered", "Available", "Available %"], med_cov,
          note="Sorted worst availability first. 'Applicable per EML' and 'Managed' are "
               "asked of every observation; availability only where managed."),
    Table("Observations by programme", ["Programme", "Observations", "%"],
          freq(MEDROWS, lambda r: s(r[PROGRAM])),
          total=["TOTAL", N_MED, 1.0], chart=("h", 0, 1, None)),
    Table("Observations by drug class", ["Drug class", "Observations", "%"],
          freq(MEDROWS, lambda r: s(r[DCLASS])),
          total=["TOTAL", N_MED, 1.0]),
    Table("Applicability and management",
          ["Question", "Yes", "No", "Yes %"], [
              ["Applicable to this facility level per EML",
               sum(1 for r in MEDROWS if yes(r[APPLICABLE])),
               sum(1 for r in MEDROWS if no(r[APPLICABLE])),
               sum(1 for r in MEDROWS if yes(r[APPLICABLE])) / N_MED],
              ["Managed by the facility",
               sum(1 for r in MEDROWS if yes(r[MANAGED])),
               sum(1 for r in MEDROWS if no(r[MANAGED])),
               sum(1 for r in MEDROWS if yes(r[MANAGED])) / N_MED],
          ],
          note="139 observations are flagged not applicable to the facility level per EML, "
               "yet 394 are not managed - so 255 observations concern medicines the EML "
               "expects the facility to carry but which it does not manage. That gap is "
               "the substantive finding in this table."),
    Table("Applicable per EML but not managed by the facility",
          ["Medicine", "Programme", "Observations", "Facilities"],
          sorted([[m, s([r for r in MEDROWS if s(r[MED]) == m][0][PROGRAM]),
                   len([r for r in MEDROWS if s(r[MED]) == m
                        and yes(r[APPLICABLE]) and no(r[MANAGED])]),
                   len({s(r["_p"][FAC]) for r in MEDROWS if s(r[MED]) == m
                        and yes(r[APPLICABLE]) and no(r[MANAGED])})]
                  for m in {s(r[MED]) for r in MEDROWS}],
                 key=lambda r: -r[2]),
          note="Where the EML says the facility level should carry the medicine but the "
               "facility does not manage it. The stated reasons are in the next table."),
    Table("Stated reasons for not managing a medicine (first 90 characters)",
          ["Reason as entered", "Observations", "%"], why_rows,
          note=f"Free text, {sum(1 for r in MEDROWS if not blank(r[WHY_NOT]))} of "
               f"{sum(1 for r in MEDROWS if no(r[MANAGED]))} not-managed observations give "
               "a reason. Worth coding into a short categorical list for the next round."),
]

# ---------------------------------------------------------------
# 5-8. the profile sheets
# ---------------------------------------------------------------
def profile_table(title, keyfn, note=None, chart=None, sort_key=None):
    rows = profile(MEDROWS, keyfn, sort_key)
    return Table(title, PROFILE_HDR, rows, note=note,
                 total=total_row(MEDROWS), chart=chart)


availability = [
    profile_table("By province", lambda r: s(r["_p"][PROV]),
                  note="Availability % is over observations where the medicine is managed. "
                       "Compare 'Availability answered' before reading any rate - North and "
                       "Kigali City rest on very few observations.",
                  chart=("v", 0, 7, None)),
    profile_table("By district", lambda r: s(r["_p"][DIST])),
    profile_table("By facility level", lambda r: s(r["_p"][LEVEL])),
    profile_table("By facility management", lambda r: s(r["_p"][MGMT])),
    profile_table("By programme", lambda r: s(r[PROGRAM]), chart=("h", 0, 7, None)),
    profile_table("By drug class", lambda r: s(r[DCLASS])),
]

facility_scorecard = [
    profile_table("By facility", lambda r: s(r["_p"][FAC]),
                  note="One row per facility, comparable across facilities: every facility "
                       "carries 42-56 observations. 'Availability answered' is the "
                       "denominator for 'Available %' and varies with how many medicines "
                       "the facility manages.",
                  chart=("h", 0, 7, None)),
]

medicine_scorecard = [
    profile_table("By medicine", lambda r: s(r[MED]),
                  note="One row per medicine across all facilities. Sorted by observation "
                       "count; sort on 'Available %' in Excel to rank by availability.",
                  sort_key=lambda r: -r[1]),
    Table("Lowest availability - medicines with at least 5 answered observations",
          ["Medicine", "Availability answered", "Available", "Available %"],
          sorted([[r[0], r[5], r[6], r[7]]
                  for r in profile(MEDROWS, lambda x: s(x[MED])) if r[5] >= 5],
                 key=lambda r: (r[3] if r[3] is not None else 9))[:20],
          note="Restricted to medicines with 5 or more answered observations, so a single "
               "stock-out at one facility cannot produce a 0% headline.",
          chart=("h", 0, 3, 15)),
]

# ---------------------------------------------------------------
# 6. stock levels and stock-outs
# ---------------------------------------------------------------
stock = [
    Table("Stock classification by province",
          ["Province", "Classified", "Stock out", "Understock", "Plan stock", "Overstock",
           "Stock out %", "Understock %", "Overstock %"],
          [[p] + (lambda c, n: [n, c.get("Stock out", 0), c.get("Understock", 0),
                                c.get("Plan stock", 0), c.get("Overstock", 0),
                                c.get("Stock out", 0) / n if n else None,
                                c.get("Understock", 0) / n if n else None,
                                c.get("Overstock", 0) / n if n else None])(
              Counter(s(r[STOCKCLASS]) for r in MEDROWS
                      if s(r["_p"][PROV]) == p and not blank(r[STOCKCLASS])),
              sum(1 for r in MEDROWS if s(r["_p"][PROV]) == p and not blank(r[STOCKCLASS])))
           for p in sorted({s(r[PROV]) for r in SUB})],
          note="Understock is below 1 month of stock, overstock above 2 months, on a 90-day "
               "window - the thresholds recorded in the form itself."),
    Table("Stock classification by programme",
          ["Programme", "Classified", "Stock out", "Understock", "Plan stock", "Overstock",
           "Stock out %", "Understock %", "Overstock %"],
          [[p] + (lambda c, n: [n, c.get("Stock out", 0), c.get("Understock", 0),
                                c.get("Plan stock", 0), c.get("Overstock", 0),
                                c.get("Stock out", 0) / n if n else None,
                                c.get("Understock", 0) / n if n else None,
                                c.get("Overstock", 0) / n if n else None])(
              Counter(s(r[STOCKCLASS]) for r in MEDROWS
                      if s(r[PROGRAM]) == p and not blank(r[STOCKCLASS])),
              sum(1 for r in MEDROWS if s(r[PROGRAM]) == p and not blank(r[STOCKCLASS])))
           for p in sorted({s(r[PROGRAM]) for r in MEDROWS})]),
    Table("Months of stock on hand", STAT_HDR,
          [stat_row("stock_level_months (all classified)", [num(r[MOS]) for r in MEDROWS])] +
          [stat_row(f"  {p}", [num(r[MOS]) for r in MEDROWS if s(r["_p"][PROV]) == p])
           for p in sorted({s(r[PROV]) for r in SUB})],
          note="A median far above 2 months alongside a large Overstock count points at "
               "over-ordering rather than scarcity - the opposite problem to a stock-out, "
               "and it ties up working capital and risks expiry."),
    Table("Stock-out days, June-August 2026", STAT_HDR,
          [stat_row("June stock-out days", [num(r[SO_MONTH['June']]) for r in MEDROWS]),
           stat_row("July stock-out days", [num(r[SO_MONTH['July']]) for r in MEDROWS]),
           stat_row("August stock-out days", [num(r[SO_MONTH['August']]) for r in MEDROWS]),
           stat_row("Total across the window", [num(r[SO_TOTAL]) for r in MEDROWS]),
           stat_row("Stock-out rate % of days", [num(r[SO_RATE]) for r in MEDROWS]),
           stat_row("Current stock-out days (this visit)", [num(r[SO_NOW]) for r in MEDROWS])],
          note="'Current stock-out days (this visit)' is asked only where the medicine was "
               "unavailable. Its maximum of 365 days on a 90-day window is not possible and "
               "indicates at least one data-entry error."),
    Table("Medicines with the most stock-out days",
          ["Medicine", "Programme", "Observations", "Stock-out days (total)",
           "Mean stock-out rate %"],
          sorted([[m, s([r for r in MEDROWS if s(r[MED]) == m][0][PROGRAM]),
                   len([r for r in MEDROWS if s(r[MED]) == m]),
                   round(sum(num(r[SO_TOTAL]) or 0 for r in MEDROWS if s(r[MED]) == m)),
                   (lambda v: round(sum(v) / len(v), 1) if v else None)(
                       [num(r[SO_RATE]) for r in MEDROWS
                        if s(r[MED]) == m and num(r[SO_RATE]) is not None])]
                  for m in {s(r[MED]) for r in MEDROWS}], key=lambda r: -r[3])[:25],
          chart=("h", 0, 3, 15)),
    Table("Facilities with the most stock-out days",
          ["Facility", "Province", "Observations", "Stock-out days (total)",
           "Mean stock-out rate %"],
          sorted([[f, s([r for r in SUB if s(r[FAC]) == f][0][PROV]),
                   len([r for r in MEDROWS if s(r["_p"][FAC]) == f]),
                   round(sum(num(r[SO_TOTAL]) or 0 for r in MEDROWS if s(r["_p"][FAC]) == f)),
                   (lambda v: round(sum(v) / len(v), 1) if v else None)(
                       [num(r[SO_RATE]) for r in MEDROWS
                        if s(r["_p"][FAC]) == f and num(r[SO_RATE]) is not None])]
                  for f in FACILITIES], key=lambda r: -r[3])),
]

# ---------------------------------------------------------------
# 7. consumption and inventory accuracy
# ---------------------------------------------------------------
emr_rows = freq([r for r in MEDROWS if not blank(r[EMR])], lambda r: s(r[EMR]), top=40)

consumption = [
    Table("Consumption and stock quantities", STAT_HDR,
          [stat_row("June consumption", [num(r[CONS['June']]) for r in MEDROWS]),
           stat_row("July consumption", [num(r[CONS['July']]) for r in MEDROWS]),
           stat_row("August consumption", [num(r[CONS['August']]) for r in MEDROWS]),
           stat_row("Average monthly consumption (AMC)", [num(r[AMC]) for r in MEDROWS]),
           stat_row("Physical quantity available", [num(r[QTY_PHYS]) for r in MEDROWS]),
           stat_row("Quantity in the system", [num(r[QTY_SYS]) for r in MEDROWS])]),
    Table("Highest-consumption medicines",
          ["Medicine", "Programme", "Observations", "Total Jun-Aug consumption", "Total AMC"],
          sorted([[m, s([r for r in MEDROWS if s(r[MED]) == m][0][PROGRAM]),
                   len([r for r in MEDROWS if s(r[MED]) == m]),
                   round(sum((num(r[CONS['June']]) or 0) + (num(r[CONS['July']]) or 0) +
                             (num(r[CONS['August']]) or 0)
                             for r in MEDROWS if s(r[MED]) == m)),
                   round(sum(num(r[AMC]) or 0 for r in MEDROWS if s(r[MED]) == m))]
                  for m in {s(r[MED]) for r in MEDROWS}], key=lambda r: -r[3])[:25],
          chart=("h", 0, 3, 15)),
    Table("Inventory accuracy - physical count vs the system", STAT_HDR,
          [stat_row("inventory_accuracy_pct", [num(r[ACCURACY]) for r in MEDROWS])],
          note="100% means the physical count matches the system. Below 100% means the "
               "system holds more than the shelf; above 100% means the shelf holds more "
               "than the system. Negative values occur where the system figure exceeds the "
               "count by more than the count itself."),
    Table("Inventory accuracy, banded", ["Band", "Observations", "%"],
          (lambda vals: [[b, sum(1 for v in vals if lo <= v < hi), sum(1 for v in vals if lo <= v < hi) / len(vals)]
                         for b, lo, hi in [("Negative (system far exceeds shelf)", -10 ** 9, 0),
                                           ("0-49%", 0, 50), ("50-79%", 50, 80),
                                           ("80-94%", 80, 95), ("95-99%", 95, 100),
                                           ("Exactly 100%", 100, 100.000001),
                                           ("Above 100%", 100.000001, 10 ** 9)]])(
              [num(r[ACCURACY]) for r in MEDROWS if num(r[ACCURACY]) is not None]),
          note="Anything other than 'Exactly 100%' is a discrepancy between the shelf and "
               "the electronic record, and is actionable at the facility.",
          chart=("h", 0, 1, None)),
    Table("Electronic system named for stock recording", ["System as entered", "Observations", "%"],
          emr_rows,
          note="Free text, so eLMIS / ELMIS / E-LMIS style variants split the count. Worth "
               "converting to a select_one before the next round."),
]

# ---------------------------------------------------------------
# 8. requisition fulfilment
# ---------------------------------------------------------------
def month_fulfil(m):
    rq = sum(num(r[REQ[m]]) or 0 for r in MEDROWS)
    dl = sum(num(r[DEL[m]]) or 0 for r in MEDROWS)
    return [m + " 2026", round(rq), round(dl), round(rq - dl), dl / rq if rq else None]


fulfilment = [
    Table("Requisition fulfilment by month",
          ["Month", "Requested", "Delivered", "Undelivered", "Fulfilment %"],
          [month_fulfil(m) for m in ("June", "July", "August")],
          total=["Jun-Aug", round(req_all), round(del_all), round(req_all - del_all),
                 del_all / req_all if req_all else None],
          note="Ratios of summed quantities, not averages of per-row percentages.",
          chart=("v", 0, 4, None)),
    Table("Fulfilment by province",
          ["Province", "Observations", "Requested", "Delivered", "Undelivered", "Fulfilment %"],
          sorted([[p, len([r for r in MEDROWS if s(r["_p"][PROV]) == p]),
                   round(sum(num(r[REQ_TOT]) or 0 for r in MEDROWS if s(r["_p"][PROV]) == p)),
                   round(sum(num(r[DEL_TOT]) or 0 for r in MEDROWS if s(r["_p"][PROV]) == p)),
                   round(sum((num(r[UNDEL_TOT]) or 0) for r in MEDROWS if s(r["_p"][PROV]) == p)),
                   (lambda rq, dl: dl / rq if rq else None)(
                       sum(num(r[REQ_TOT]) or 0 for r in MEDROWS if s(r["_p"][PROV]) == p),
                       sum(num(r[DEL_TOT]) or 0 for r in MEDROWS if s(r["_p"][PROV]) == p))]
                  for p in {s(r[PROV]) for r in SUB}], key=lambda r: -r[2]),
          chart=("v", 0, 5, None)),
    Table("Fulfilment by facility",
          ["Facility", "Province", "Observations", "Requested", "Delivered", "Fulfilment %"],
          sorted([[f, s([r for r in SUB if s(r[FAC]) == f][0][PROV]),
                   len([r for r in MEDROWS if s(r["_p"][FAC]) == f]),
                   round(sum(num(r[REQ_TOT]) or 0 for r in MEDROWS if s(r["_p"][FAC]) == f)),
                   round(sum(num(r[DEL_TOT]) or 0 for r in MEDROWS if s(r["_p"][FAC]) == f)),
                   (lambda rq, dl: dl / rq if rq else None)(
                       sum(num(r[REQ_TOT]) or 0 for r in MEDROWS if s(r["_p"][FAC]) == f),
                       sum(num(r[DEL_TOT]) or 0 for r in MEDROWS if s(r["_p"][FAC]) == f))]
                  for f in FACILITIES], key=lambda r: -r[3])),
    Table("Fulfilment by medicine",
          ["Medicine", "Programme", "Observations", "Requested", "Delivered",
           "Undelivered", "Fulfilment %"],
          sorted([[m, s([r for r in MEDROWS if s(r[MED]) == m][0][PROGRAM]),
                   len([r for r in MEDROWS if s(r[MED]) == m]),
                   round(sum(num(r[REQ_TOT]) or 0 for r in MEDROWS if s(r[MED]) == m)),
                   round(sum(num(r[DEL_TOT]) or 0 for r in MEDROWS if s(r[MED]) == m)),
                   round(sum(num(r[UNDEL_TOT]) or 0 for r in MEDROWS if s(r[MED]) == m)),
                   (lambda rq, dl: dl / rq if rq else None)(
                       sum(num(r[REQ_TOT]) or 0 for r in MEDROWS if s(r[MED]) == m),
                       sum(num(r[DEL_TOT]) or 0 for r in MEDROWS if s(r[MED]) == m))]
                  for m in {s(r[MED]) for r in MEDROWS}], key=lambda r: -r[5])),
    Table("Requisition ratios against consumption", STAT_HDR,
          [stat_row("fulfilment_pct (per observation)", [num(r[FULFIL]) for r in MEDROWS]),
           stat_row("request_vs_amc_pct", [num(r[REQ_AMC]) for r in MEDROWS]),
           stat_row("delivery_vs_amc_pct", [num(r[DEL_AMC]) for r in MEDROWS])],
          note="request_vs_amc far above 100% means facilities order well beyond their own "
               "consumption - which is the mechanism behind the overstock finding. "
               "fulfilment_pct is populated on 660 of 1,359 observations only."),
]

# ---------------------------------------------------------------
# 9. supply attribution and cost of the gap
# ---------------------------------------------------------------
attribution = [
    Table("Supply chain attribution for unavailable medicines",
          ["Point in the chain", "Observations", "%"],
          freq(unavail, lambda r: s(r[DIST_FLAG]) or "(not answered)"),
          total=["TOTAL", len(unavail), 1.0],
          note="A 'Distribution issue' means RMS held the medicine but the facility did "
               "not - the gap is between the warehouse and the shelf, not in national "
               "supply. That is where the pilot's leverage is.",
          chart=("h", 0, 1, None)),
    Table("RMS availability when the facility was out",
          ["Question", "Yes", "No", "Yes %"], [
              ["Available at RMS Central",
               sum(1 for r in unavail if yes(r[RMS_CEN])),
               sum(1 for r in unavail if no(r[RMS_CEN])),
               sum(1 for r in unavail if yes(r[RMS_CEN])) / len(unavail)],
              ["Available at RMS Regional",
               sum(1 for r in unavail if yes(r[RMS_REG])),
               sum(1 for r in unavail if no(r[RMS_REG])),
               sum(1 for r in unavail if yes(r[RMS_REG])) / len(unavail)],
          ]),
    Table("Every unavailable observation",
          ["Province", "District", "Facility", "Level", "Medicine", "Programme",
           "Current stock-out days", "Total stock-out days", "RMS Central", "RMS Regional",
           "Attribution", "Unit price (RWF)", "Financial loss (RWF)"],
          sorted([[s(r["_p"][PROV]), s(r["_p"][DIST]), s(r["_p"][FAC]), s(r["_p"][LEVEL]),
                   s(r[MED]), s(r[PROGRAM]), num(r[SO_NOW]), num(r[SO_TOTAL]),
                   s(r[RMS_CEN]), s(r[RMS_REG]), s(r[DIST_FLAG]), num(r[PRICE]),
                   num(r[LOSS])] for r in unavail],
                 key=lambda r: (r[0], r[2], r[4])),
          note=f"All {len(unavail)} observations recorded as unavailable on the visit day."),
    Table("Unmet demand and cost", STAT_HDR,
          [stat_row("Quantity billed", [num(r[BILLED]) for r in MEDROWS]),
           stat_row("Prescriptions per year (RSSB)", [num(r[RX_YEAR]) for r in MEDROWS]),
           stat_row("Quantity per prescription", [num(r[QTY_PER_RX]) for r in MEDROWS]),
           stat_row("Prescriptions missed to stock-out", [num(r[RX_MISSED]) for r in MEDROWS]),
           stat_row("Estimated quantity unfilled", [num(r[QTY_UNFILLED]) for r in MEDROWS]),
           stat_row("Unserved requisition gap", [num(r[GAP]) for r in MEDROWS]),
           stat_row("Overall unmet demand (Rx)", [num(r[UNMET]) for r in MEDROWS]),
           stat_row("Financial loss (RWF)", [num(r[LOSS]) for r in MEDROWS])],
          note="Quantity billed and RSSB prescriptions are populated on only 163-164 of "
               "1,359 observations - roughly one per submission - so these are not "
               "comparable across medicines yet. Treat every derived cost built on them as "
               "provisional until that is explained."),
    Table("Financial loss by medicine",
          ["Medicine", "Programme", "Observations with a loss figure", "Financial loss (RWF)"],
          sorted([[m, s([r for r in MEDROWS if s(r[MED]) == m][0][PROGRAM]),
                   len([r for r in MEDROWS if s(r[MED]) == m and num(r[LOSS]) is not None]),
                   round(sum(num(r[LOSS]) or 0 for r in MEDROWS if s(r[MED]) == m))]
                  for m in {s(r[MED]) for r in MEDROWS}], key=lambda r: -r[3])[:25],
          chart=("h", 0, 3, 15)),
    Table("Financial loss by province and programme",
          ["Group", "Observations with a loss figure", "Financial loss (RWF)"],
          [[f"Province: {p}",
            len([r for r in MEDROWS if s(r["_p"][PROV]) == p and num(r[LOSS]) is not None]),
            round(sum(num(r[LOSS]) or 0 for r in MEDROWS if s(r["_p"][PROV]) == p))]
           for p in sorted({s(r[PROV]) for r in SUB})] +
          [[f"Programme: {p}",
            len([r for r in MEDROWS if s(r[PROGRAM]) == p and num(r[LOSS]) is not None]),
            round(sum(num(r[LOSS]) or 0 for r in MEDROWS if s(r[PROGRAM]) == p))]
           for p in sorted({s(r[PROGRAM]) for r in MEDROWS})],
          total=["TOTAL", sum(1 for r in MEDROWS if num(r[LOSS]) is not None), round(loss_all)]),
]

# ---------------------------------------------------------------
# 10. feedback + findings
# ---------------------------------------------------------------
feedback = [
    Table("Collector feedback on the supply chain",
          ["Province", "Facility", "Enumerator", "Feedback as entered"],
          sorted([[s(r[PROV]), s(r[FAC]), s(r[ENUM]), s(r[FEEDBACK])]
                  for r in SUB if not blank(r.get(FEEDBACK))], key=lambda r: (r[0], r[1])),
          note=f"{sum(1 for r in SUB if not blank(r.get(FEEDBACK)))} of {N_SUB} submissions "
               "left feedback. Read alongside the attribution table - much of it describes "
               "the same last-mile distribution problem in the collectors' own words."),
]

findings = [
    Table("Findings and recommended actions", ["Area", "Finding", "Evidence", "Action"], [
        ["Coverage",
         f"{len(FACILITIES)} facilities have submitted, across {len(SUB)} submissions - "
         f"report facilities, not submissions",
         f"{SUB_ONE_MED} of {len(SUB)} submissions carry a single medicine while "
         f"{SUB_MANY_MED} carry {SUB_BULK_LO}-{SUB_BULK_HI}; most single-medicine "
         f"submissions come from {SUB_ONE_MED_TOP}. Facilities hold "
         f"{FAC_OBS_LO}-{FAC_OBS_HI} medicine observations each.",
         "Report facilities and medicine observations, never submission counts. The split "
         "submissions are a data-entry style, not a coverage gap, so no facility needs "
         "re-visiting - but standardise on one submission per facility visit next round."],
        ["Availability",
         "Availability on the visit day is high",
         f"{avail_all:,} of {len(answered_all):,} managed observations available "
         f"({avail_all / len(answered_all) * 100:.1f}%); {len(unavail):,} unavailable.",
         "Availability is not the binding constraint at these facilities. Shift the "
         "analytical focus to stock adequacy and requisition fulfilment."],
        ["Stock adequacy",
         "Overstock is the dominant stock problem, not scarcity",
         f"Of {N_CLASSIFIED:,} classified observations: {cls('Overstock'):,} Overstock, "
         f"{cls('Understock'):,} Understock, {cls('Plan stock'):,} Plan stock, "
         f"{cls('Stock out'):,} Stock out. request_vs_amc_pct runs far above 100%.",
         "Investigate over-ordering: it ties up working capital and creates expiry risk. "
         "This is a more actionable finding than the stock-out count."],
        ["Attribution",
         "The gap is last-mile distribution, not national supply",
         f"Of {len(unavail):,} unavailable observations, "
         f"{ATTRIB_N.get('Distribution issue', 0):,} are flagged 'Distribution issue' "
         f"(RMS held stock) and only "
         f"{ATTRIB_N.get('Nationwide stock-out', 0):,} 'Nationwide stock-out'.",
         "Target facility ordering behaviour and RMS-to-facility delivery. This mirrors "
         "the prescription study, where most unserved medicines were held at RMS."],
        ["EML alignment",
         f"{N_EML_GAP:,} observations cover medicines the EML expects but the facility "
         f"does not manage",
         f"{N_APPLICABLE:,} observations are applicable per EML, yet {N_NOT_MANAGED:,} "
         f"are not managed - only {N_NOT_APPLICABLE:,} are flagged not applicable.",
         "Reconcile the facility-level EML expectation against what facilities actually "
         "stock. The free-text reasons should be coded into a categorical list."],
        ["Data quality",
         "'username' is broken and GPS is mostly absent",
         f"username reads 'username not found' on all {N_USERNAME_BROKEN:,} submissions; "
         f"GPS captured on {N_GEO:,} of {len(SUB):,} "
         f"({N_GEO / len(SUB) * 100:.0f}%); {N_NO_ENUM} submissions have no enumerator "
         f"name.",
         "Fix the username calculate and make the geopoint required, or drop it. Enumerator "
         "identity currently depends on one free-choice question."],
        ["Data quality",
         "Kobo calculate fields can export the literal string 'NaN'",
         "Five calculated columns (qty_per_prescription, est_qty_unfilled_stockout, "
         "est_unserved_requisition_gap, financial_loss, overall_unmet_demand_rx) carry "
         "'NaN' on one row - a divide-by-zero in the form. float('NaN') parses "
         "silently, so one such value turns a whole column's total into NaN.",
         "Guard the form's calculates against a zero or blank denominator. This script "
         "now treats any non-finite value as missing."],
        ["Data quality",
         "Stock-out day counts exceed the window they are measured over",
         f"'Current stock-out days (this visit)' reaches {SO_CUR_MAX:,.0f} and "
         f"total_stockout_days reaches {SO_TOT_MAX:,.0f}, both impossible on a 90-day "
         f"(Jun-Aug) window. The monthly columns behave correctly (max 30/31).",
         "Constrain the current-stock-out question to 0-90 and the monthly ones to the "
         "length of their month, then correct the affected rows."],
        ["Data quality",
         f"Cost fields rest on {N_LOSS_ROWS:,} of {N_MED:,} observations",
         "Quantity billed and RSSB prescriptions per year are populated roughly once per "
         f"submission, and financial_loss on {N_LOSS_ROWS:,} rows.",
         "Establish whether these are meant to be per-medicine or per-facility before any "
         "financial loss figure is published."],
        ["Instrument",
         f"{N_VERSIONS} form versions are live and several columns are empty",
         f"{N_VERSIONS} __version__ values; {len(EMPTY_COLS)} columns hold no data in this "
         f"export"
         + (f" (including {', '.join(c[:40] for c in EMPTY_COLS[:3])})" if EMPTY_COLS
            else "")
         + f". The medicine choice list is now clean: {N_MED_NAMES} distinct medicines "
           f"with no duplicated entries (the earlier duplicates are resolved).",
         "Freeze the instrument for the round and drop the unreachable columns."],
        ["Access",
         "The dataset is readable without credentials",
         f"The export URL returns HTTP 200 with no Authorization header (verified "
         f"{dt.date.today().isoformat()}), exposing "
         f"{len({s(r[ENUM]) for r in SUB if not blank(r.get(ENUM))})} enumerator names, "
         f"device IDs and {N_GEO} GPS points. Anyone with the link can download the "
         f"whole dataset.",
         "Close anonymous access in the project's Sharing settings, then recreate the "
         "export setting so the current URL stops working. Use KOBO_API_TOKEN instead."],
    ]),
]

SECTIONS = [
    ("README", readme),
    ("Overview", overview),
    ("Coverage", coverage),
    ("Data quality", quality),
    ("Medicine coverage", medcov),
    ("Availability profiles", availability),
    ("Stock and stock-outs", stock),
    ("Consumption and accuracy", consumption),
    ("Requisition fulfilment", fulfilment),
    ("Attribution and cost", attribution),
    ("Facility scorecard", facility_scorecard),
    ("Medicine scorecard", medicine_scorecard),
    ("Collector feedback", feedback),
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
        ws.cell(row=1, column=1, value=name).font = Font(bold=True, size=14, color="1F3864")
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
                                            wrap_text=isinstance(v, str) and len(str(v)) > 55)
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
    ws.cell(row=1, column=1, value="Charts").font = Font(bold=True, size=14, color="1F3864")
    ws.cell(row=2, column=1, value="Native Excel charts over the data block to the left of "
                                   "each - click any chart to edit or copy it.").font = NOTE_FONT
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
            ch.set_categories(Reference(ws, min_col=1, min_row=h0 + 1, max_row=h0 + len(pairs_)))
            ch.height = max(7.5, min(0.6 * len(pairs_) + 4, 19))
            ch.width = 21
            ws.add_chart(ch, f"D{row}")
            row = max(h0 + len(pairs_) + 2, h0 + int(ch.height / 0.53) + 2)
            n += 1

    wb.save(OUT)
    for d in sorted(DUPES_DROPPED, key=lambda d: (d["facility"], d["medicine"])):
        tag = "identical copy" if d["identical"] else f"DISAGREED on {d['differs']}"
        print(f"  de-duplicated {d['facility']} / {d['medicine'][:38]}: kept "
              f"{d['kept_sub']}, dropped {d['dropped_sub']} - {tag}")
    if DUPES_CONFLICTING:
        print(f"  WARNING: {len(DUPES_CONFLICTING)} of {len(DUPES_DROPPED)} "
              f"duplicates disagreed with the row kept - confirm those "
              f"medicines with the collector (Data quality sheet)")
    print(f"wrote {OUT}")
    print(f"  {len(SECTIONS)} analysis sheets + Charts")
    print(f"  {sum(len(ts) for _, ts in SECTIONS)} tables, {n} charts")
    print(f"  unit of analysis: {N_MED} medicine observations across {len(FACILITIES)} facilities")


if __name__ == "__main__":
    write()
