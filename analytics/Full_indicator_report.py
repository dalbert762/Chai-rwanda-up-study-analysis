"""Every feasible indicator in the master analytical prompt, in one place.

    python3 Full_indicator_report.py
      -> Full_indicator_set.xlsx      every indicator as a table
      -> Indicator_dashboard.html     the same indicators as a tabbed page

One computation engine, two renderers, so the workbook and the dashboard can
never disagree. Indicators are grouped by the prompt's eight analyses and
carry their metric number, numerator, denominator and reference period.

Scope follows Master_prompt_feasibility.xlsx: the 28 metrics that are
computable as specified plus the 12 that can only be proxied, each labelled.
The 7 with no source variable are listed on the Method tab as gaps rather
than estimated - there is no dispensed quantity, no service type, no stock-out
episode count, no requisition dates and no warehouse records in this data.

This is the ALL-INDICATORS dashboard. Combined_analysis_report.html remains
the short curated one.

openpyxl + the standard library, like the rest of this project.
"""
import _bootstrap as B   # path setup - must come first

import datetime as dt
import html
import json
import os
import sys
import re
import math
import statistics
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_kobo_prescription as P
import analyze_medicine_assessment as M
import analyze_medicine_assessment_chw as C
import facility_levels as FL
import molecules as MOL
import not_managed as NM
import charts as CH
import explorer as EX
import stock_explorer as SX
import multiselect as MS
import importlib
import pagefilter as PF

# The two maps are built by Tracer_availability_mapping, which owns the
# projection, the colour scale and the facility coordinate join. Imported
# rather than duplicated, so the maps here and the standalone map file can
# never diverge. If the boundary GeoJSON is missing the dashboard still
# builds - it just has no maps.
try:
    import Tracer_availability_mapping as MAP
except Exception as _map_err:          # noqa: BLE001 - any failure is non-fatal
    MAP, MAP_ERR = None, _map_err
else:
    MAP_ERR = None

OUT_XLSX = B.out("Full_indicator_set.xlsx")
# DASH_OUT lets a test build write somewhere other than the real dashboard
OUT_HTML = os.environ.get("DASH_OUT") or B.out("Indicator_dashboard.html")
PERIOD = "June-August 2026 (visit day for availability)"
MIN_N = 30


def s(v):
    return "" if v is None else str(v).strip()


def pct(a, b):
    return a / b if b else None


def med(v):
    v = [x for x in v if x is not None]
    return round(statistics.median(v), 2) if v else None


# ---------------------------------------------------------------
# prepared row sets
# ---------------------------------------------------------------
MED_ALL = M.MEDROWS
MG = [r for r in MED_ALL if M.yes(r[M.MANAGED])]
ANS = [r for r in MG if not M.blank(r[M.AVAIL])]
AV = [r for r in ANS if M.yes(r[M.AVAIL])]
SO = [r for r in ANS if M.no(r[M.AVAIL])]
APP = [r for r in MED_ALL if M.yes(r[M.APPLICABLE])]
for r in MED_ALL:
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])

# Method 3: availability is asked of the MOLECULE, not of each strength.
# These sets replace MED_ALL/MG/ANS/AV wherever the question is "could the
# facility treat the patient"; quantities, stock-out days and requisitions
# keep the item-level sets above, because two strengths do not add up.
# DEFINITION (study, 2026-10-13) - two tracks, set by medicine_program:
#   Essential Medicines  denominator = every molecule the EML authorises at
#                        that level; authorised but not managed counts as
#                        UNAVAILABLE.
#   any other programme  denominator = the molecules the facility manages;
#                        programme products are allocated (MFT zoning,
#                        faith-based FP, level-restricted MCCH), so not
#                        holding one is not a shortage.
# MOL.two_track_base() applies both and keeps the real management flag on
# `_managed`, the reason on `_gap` and the track on `_track`.
MOLS = MOL.two_track_base(MED_ALL)
for r in MOLS:
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])
# MOLS is already the expected set, with the Essential Medicines gap forced
# to unavailable. MANAGED is forced "Yes" on those rows so the filters below
# still read, so anything about MANAGEMENT must use `_managed`.
APPM = MOLS
ANSM = [r for r in MOLS if not M.blank(r[M.AVAIL])]
AVM = [r for r in ANSM if M.yes(r[M.AVAIL])]
SOM = [r for r in ANSM if M.no(r[M.AVAIL])]
MGM = [r for r in MOLS if r["_managed"]]
GAPM = [r for r in MOLS if r["_gap"]]

ITEMS = [r for r in P.item if r["_sub"]]
for r in ITEMS:
    r["_lvl"] = FL.group(r["_sub"][P.FTYPE])
UNSERVED = [r for r in ITEMS if r["_avail"] is False]
PRES = [r for r in P.pres if r["_sub"]]

CM = [r for r in C.MEDROWS if C.yes(r[C.MANAGED])]
CANS = [r for r in CM if not C.blank(r[C.AVAIL])]
CAV = [r for r in CANS if C.yes(r[C.AVAIL])]

LEVELS = FL.ordered(r["_lvl"] for r in MED_ALL)
MONTHS = ("June", "July", "August")

# The stock tool's drug-class labels for the MCCH products are wrong or
# split: tranexamic acid is filed as an anti-apneic and caffeine citrate as a
# haemostatic (swapped), carbetocin's "uterotonic" sits apart from
# oxytocin's "Uterotonics". Corrected for display only - the recorded value
# is left alone, as the cleaning log's drug-class decisions are still open.
CLASS_LABEL = {"Anti-apneic agent": "Haemostatics (tranexamic acid)",
               "Hemostatic agent": "Anti-apnoeics (caffeine citrate)",
               "uterotonic": "Uterotonics",
               "Treatment of eclampsia.": "Eclampsia (MgSO4)"}


def _class_label(r):
    c = s(r[M.DCLASS]) or "(unclassified)"
    return CLASS_LABEL.get(c, c)


TABLES = []          # (analysis, metric, title, status, note, headers, rows)


def add(analysis, metric, title, headers, rows, note="", status="Feasible"):
    TABLES.append((analysis, metric, title, status, note, headers, rows))


def grouped(keyfn, keys=None, rows=None):
    """availability measure set for one dimension (Method 3 rows by default)"""
    src = rows if rows is not None else MOLS
    g = defaultdict(list)
    for r in src:
        g[keyfn(r)].append(r)
    ks = keys or sorted(g)
    out = []
    for k in ks:
        rs = g.get(k, [])
        a = [x for x in rs if not M.blank(x[M.AVAIL])]
        av = sum(1 for x in a if M.yes(x[M.AVAIL]))
        gap = sum(1 for x in rs if x.get("_gap"))
        out.append([k, len({s(x["_p"][M.FAC]) for x in rs}), len(rs), gap,
                    len(rs) - gap, len(a) - av - gap, av,
                    pct(av, len(a)) if len(a) >= MIN_N else None])
    return out


# ---------------------------------------------------------------
# shared disaggregation
# ---------------------------------------------------------------
# Page 1 cuts availability by programme, level of care, province, district,
# management type and drug class. The other pages were carrying one cut each
# or none, which made them unanswerable to the obvious follow-up question -
# "where?". `rate_by` gives any rate indicator the same set of cuts, and
# `vol_by` does the same for a quantity. Both keep the reporting minimum, so
# a cell computed on three observations is left blank rather than ranked.
RATE_HDR = ["Group", "Facilities", "Observations", "Numerator", "RATE"]


def rate_by(rows, keyfn, num_fn, facfn, keys=None, min_n=None, sort=True):
    mn = MIN_N if min_n is None else min_n
    g = defaultdict(list)
    for r in rows:
        g[keyfn(r)].append(r)
    out = []
    for k in (keys or sorted(g)):
        rs = g.get(k, [])
        if not rs:
            continue
        hit = sum(1 for r in rs if num_fn(r))
        out.append([k, len({facfn(r) for r in rs}), len(rs), hit,
                    pct(hit, len(rs)) if len(rs) >= mn else None])
    if sort:
        out.sort(key=lambda x: (x[4] is None, x[4]))
    return out


VOL_HDR = ["Group", "Facilities", "Observations", "Total", "Median",
           "SHARE OF TOTAL"]


def vol_by(rows, keyfn, valfn, facfn, keys=None):
    g = defaultdict(list)
    for r in rows:
        g[keyfn(r)].append(r)
    grand = sum(valfn(r) or 0 for r in rows)
    out = []
    for k in (keys or sorted(g)):
        rs = g.get(k, [])
        if not rs:
            continue
        tot = sum(valfn(r) or 0 for r in rs)
        out.append([k, len({facfn(r) for r in rs}), len(rs), round(tot),
                    med([valfn(r) for r in rs]), pct(tot, grand)])
    out.sort(key=lambda x: -x[3])
    return out


def mfac(r):
    return s(r["_p"][M.FAC])


def mmgt(r):
    return s(r["_p"].get(EX.MGMT)) or "(not recorded)"


def fig(fid, rows, title, note, kind="bars", box=None, **kw):
    """A figure section: heading, note, and a container the browser can refill."""
    body = CH.bars(rows, **kw) if kind == "bars" else CH.gauges(rows)
    bid = f' id="{box}"' if box else ""
    return (f'<section class="ind"><h4><span class="mid">{esc(fid)}</span>'
            f'{esc(title)}</h4>'
            + (f'<p class="note">{esc(note)}</p>' if note else '')
            + f'<div{bid}>{body}</div></section>')


# ---------------------------------------------------------------
# upstream: where the stock was when the facility did not have it
# ---------------------------------------------------------------
# At module scope because Analysis 1's nationwide stock-out table
# and the figures on the same page both read these sets.
# Asked only of UNAVAILABLE rows, so these sit outside MOLS entirely.
UP = [r for r in MED_ALL if M.no(r[M.AVAIL])]
for r in UP:
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])

def _ans(r):
    return (M.yes(r[M.RMS_REG]) or M.no(r[M.RMS_REG])
            or M.yes(r[M.RMS_CEN]) or M.no(r[M.RMS_CEN]))

def _had(r):
    return M.yes(r[M.RMS_REG]) or M.yes(r[M.RMS_CEN])

UPA = [r for r in UP if _ans(r)]

# Regional and Central are kept APART from here on. They are separate
# questions with separate denominators - a row can carry an answer for
# one and not the other - so pooling them into "upstream" hid which of
# the two tiers actually held the stock.
def _ansd(r, col):
    return M.yes(r[col]) or M.no(r[col])

def up_by(col, keyfn, keys=None, floor=MIN_N):
    g = defaultdict(list)
    for r in UP:
        if _ansd(r, col):
            g[keyfn(r)].append(r)
    out = [[k, len(g[k]), sum(1 for r in g[k] if M.yes(r[col]))]
           for k in (keys or sorted(g)) if len(g.get(k, [])) >= floor]
    if keys is None:
        out.sort(key=lambda x: -(x[2] / x[1]))
    return out

def _pair(keyfn, keys=None, floor=MIN_N):
    """One row per category per tier, so the two sit side by side."""
    reg = {k: (n, a) for k, n, a in up_by(M.RMS_REG, keyfn, keys, floor)}
    cen = {k: (n, a) for k, n, a in up_by(M.RMS_CEN, keyfn, keys, floor)}
    order = keys or sorted(sorted(set(reg) | set(cen)),
                           key=lambda k: -(reg.get(k, (1, 0))[1]
                                           / reg.get(k, (1, 0))[0]))
    out = []
    for k in order:
        if k in reg:
            out.append([f"{k} - Regional", reg[k][0], reg[k][1]])
        if k in cen:
            out.append([f"{k} - Central", cen[k][0], cen[k][1]])
    return out

REG_A = [r for r in UP if _ansd(r, M.RMS_REG)]
CEN_A = [r for r in UP if _ansd(r, M.RMS_CEN)]
up_reg_lvl = up_by(M.RMS_REG, lambda r: r["_lvl"], keys=LEVELS, floor=1)
up_cen_lvl = up_by(M.RMS_CEN, lambda r: r["_lvl"], keys=LEVELS, floor=1)
def _grouped(keyfn, keys=None, floor=MIN_N):
    """[(label, [(n, hits) Regional, (n, hits) Central])]"""
    reg = {k: (n, a) for k, n, a in up_by(M.RMS_REG, keyfn, keys, floor)}
    cen = {k: (n, a) for k, n, a in up_by(M.RMS_CEN, keyfn, keys, floor)}
    order = keys or sorted(sorted(set(reg) & set(cen)),
                           key=lambda k: -(reg[k][1] / reg[k][0]))
    return [[k, [reg.get(k, (0, 0)), cen.get(k, (0, 0))]]
            for k in order if k in reg or k in cen]

def _donut_rows(col):
    g = defaultdict(lambda: [0, 0])
    for r in UP:
        if not _ansd(r, col):
            continue
        g[r["_lvl"]][0 if M.yes(r[col]) else 1] += 1
    return [[k, g[k][0], g[k][1]] for k in LEVELS if k in g]

up_cls = _grouped(lambda r: s(r[M.DCLASS]) or "(unclassified)")
up_prog = _grouped(lambda r: s(r[M.PROGRAM]) or "(none)")
up_lvl = _grouped(lambda r: r["_lvl"], keys=LEVELS, floor=1)
# FIG 7 (upstream by level) puts the two tiers on ONE denominator per level (2026-10-16): the
# unavailable rows answered Yes or No for BOTH Regional and Central, so the
# two dots of a level are shares of the same rows.
UP_BOTH = [r for r in UP if _ansd(r, M.RMS_REG) and _ansd(r, M.RMS_CEN)]
up_lvl_both = [
    [k, [(len(rs), sum(1 for r in rs if M.yes(r[M.RMS_REG]))),
         (len(rs), sum(1 for r in rs if M.yes(r[M.RMS_CEN])))]]
    for k in LEVELS
    for rs in ([r for r in UP_BOTH if r["_lvl"] == k],) if rs]
don_reg = _donut_rows(M.RMS_REG)
don_cen = _donut_rows(M.RMS_CEN)

# --- the form's own attribution field
DN_ISSUE_T, DN_NS_T = "Distribution issue", "Nationwide stock-out"
dn = [r for r in UP if s(r[M.DIST_FLAG]) in (DN_ISSUE_T, DN_NS_T)]

def dn_by(keyfn, keys=None, floor=MIN_N, rows=None):
    g = defaultdict(list)
    for r in (dn if rows is None else rows):
        g[keyfn(r)].append(r)
    order = keys or sorted(
        g, key=lambda k: -sum(1 for r in g[k]
                              if s(r[M.DIST_FLAG]) == DN_NS_T) / len(g[k]))
    return [[k, [(len(g[k]), sum(1 for r in g[k]
                                 if s(r[M.DIST_FLAG]) == DN_ISSUE_T)),
                 (len(g[k]), sum(1 for r in g[k]
                                 if s(r[M.DIST_FLAG]) == DN_NS_T))]]
            for k in order if len(g.get(k, [])) >= floor]

dn_lvl = dn_by(lambda r: r["_lvl"], keys=LEVELS, floor=1)
dn_prog = dn_by(lambda r: s(r[M.PROGRAM]) or "(none)")

# FIG 7-10 share ONE base (2026-10-16): the unavailable rows answered Yes or
# No for both Regional and Central (UP_BOTH). Every one of them also carries
# the form's attribution flag, so within a group the Regional, Central,
# distribution-issue and nationwide shares all have the same denominator.
def _both_by(keyfn, floor=MIN_N):
    g = defaultdict(list)
    for r in UP_BOTH:
        g[keyfn(r)].append(r)
    rows = [[k, [(len(g[k]), sum(1 for r in g[k] if M.yes(r[M.RMS_REG]))),
                 (len(g[k]), sum(1 for r in g[k] if M.yes(r[M.RMS_CEN])))]]
            for k in sorted(g) if len(g[k]) >= floor]
    return sorted(rows, key=lambda x: -(x[1][0][1] / x[1][0][0]))


# Display names and definitions for FIG 9-10 (user wording, 2026-10-16).
# The form's value "Distribution issue" is shown as "Distribution Gap"; on
# the FIG 7-10 base it matches the definition exactly (406 rows, every one
# with Central or Regional = Yes; the 64 nationwide rows have both = No).
DN_SERIES = ["Distribution Gap", "Nationwide stock-out"]
DEF_GAP = ("<b>Distribution Gap:</b> the medicine was not available at the "
           "facility but was available at RMS Central or RMS Regional level. "
           "<b>Nationwide stock-out:</b> not available at the facility, nor "
           "at RMS Regional or RMS Central.")

up_prog_both = _both_by(lambda r: s(r[M.PROGRAM]) or "(none)")
dn_lvl_both = dn_by(lambda r: r["_lvl"], keys=LEVELS, floor=1, rows=UP_BOTH)
dn_prog_both = dn_by(lambda r: s(r[M.PROGRAM]) or "(none)", rows=UP_BOTH)
DN_LAB = M.DIST_FLAG
DN_ISSUE = sum(1 for r in UP if s(r[M.DIST_FLAG]) == DN_ISSUE_T)
DN_NS = sum(1 for r in UP if s(r[M.DIST_FLAG]) == DN_NS_T)
DN_AGREE = sum(1 for r in UP if s(r[M.DIST_FLAG]) == DN_ISSUE_T
               and M.yes(r[M.RMS_CEN]))
DN_NS_AGREE = sum(1 for r in UP if s(r[M.DIST_FLAG]) == DN_NS_T
                  and M.no(r[M.RMS_CEN]))
DN_BLANK = len(UP) - DN_ISSUE - DN_NS

# the medicines a facility could not get anywhere
ns_rows = [r for r in UP if s(r[M.DIST_FLAG]) == DN_NS_T]
NS_TABLE = sorted(
    [[m,
      s(next(r[M.PROGRAM] for r in ns_rows if s(r[M.MED]) == m)),
      s(next(r[M.DCLASS] for r in ns_rows if s(r[M.MED]) == m)),
      len({s(r["_p"][M.FAC]) for r in ns_rows if s(r[M.MED]) == m}),
      len({s(r["_p"][M.DIST]) for r in ns_rows if s(r[M.MED]) == m}),
      sum(1 for r in M.MEDROWS
          if s(r[M.MED]) == m and M.yes(r[M.MANAGED])),
      pct(len({s(r["_p"][M.FAC]) for r in ns_rows if s(r[M.MED]) == m}),
          sum(1 for r in M.MEDROWS
              if s(r[M.MED]) == m and M.yes(r[M.MANAGED])))]
     for m in sorted({s(r[M.MED]) for r in ns_rows})],
    key=lambda x: -x[3])
up_had = sum(1 for r in UPA if _had(r))


AVAIL_HDR = ["Group", "Facilities", "Expected (denominator)",
             "EML item not managed (counted unavailable)", "Managed",
             "Managed but empty", "Available", "AVAILABLE %"]

# =============== ANALYSIS 1 =====================================
add(1, "1.1", "Overall tracer medicine availability",
    ["Measure", "Value"],
    [["EXPECTED (DENOMINATOR)", len(APPM)],
     ["  Essential Medicines - EML-authorised at the level",
      sum(1 for r in MOLS if r["_track"] == "EML")],
     ["  Programme products - managed by the facility",
      sum(1 for r in MOLS if r["_track"] == "managed")],
     ["Essential Medicines authorised but NOT managed - counted unavailable",
      len(GAPM)],
     ["Managed and on the shelf (NUMERATOR)", len(AVM)],
     ["Managed but empty on the visit day", len(SOM) - len(GAPM)],
     ["OVERALL AVAILABILITY", pct(len(AVM), len(APPM))],
     ["  Essential Medicines",
      pct(sum(1 for r in AVM if r["_track"] == "EML"),
          sum(1 for r in MOLS if r["_track"] == "EML"))],
     ["  Programme products",
      pct(sum(1 for r in AVM if r["_track"] == "managed"),
          sum(1 for r in MOLS if r["_track"] == "managed"))],
     ["Memo: availability over managed rows only", pct(len(AVM), len(MGM))],
     ["Facilities", len(M.FACILITIES)],
     ["Molecules", len({s(r[M.MED]) for r in MOLS})],
     ["Items assessed before collapsing strengths", M.N_MED]],
    "Numerator: the facility manages the medicine AND it was on the shelf. "
    "Denominator, by medicine_program: for Essential Medicines, every "
    "medicine the EML authorises at that level of care - one the facility is "
    "supposed to manage but does not counts as unavailable; for every other "
    "programme (Malaria, MCCH, FP, ARV, TB, Consumables), only the medicines "
    "the facility manages, because programme products are allocated and not "
    "holding one is usually policy, not a shortage. "
    f"{len(GAPM):,} of {len(APPM):,} pairs ({len(GAPM) / len(APPM):.1%}) are "
    f"Essential Medicines not managed; {len(SOM) - len(GAPM):,} more were "
    f"managed but empty. "
    "METHOD 3 - a medicine assessed at more than one strength counts once "
    "per facility and is available if ANY strength was on the shelf; "
    f"{len({s(r[M.MED]) for r in MED_ALL})} items become "
    f"{len({s(r[M.MED]) for r in MOLS})} molecules. 'Unexpired and usable' "
    "cannot be applied - no expiry or usability variable is collected, so "
    "availability means physically present.")

add(1, "1.2", "Individual medicine availability - ten least available",
    ["Medicine", "Programme", "Facilities", "Answered", "Available",
     "AVAILABLE %"],
    sorted([[m, s(next(r[M.PROGRAM] for r in MOLS if s(r[M.MED]) == m)),
             len({s(r["_p"][M.FAC]) for r in ANSM if s(r[M.MED]) == m}),
             sum(1 for r in ANSM if s(r[M.MED]) == m),
             sum(1 for r in AVM if s(r[M.MED]) == m),
             pct(sum(1 for r in AVM if s(r[M.MED]) == m),
                 sum(1 for r in ANSM if s(r[M.MED]) == m))]
            for m in sorted({s(r[M.MED]) for r in ANSM})
            if sum(1 for r in ANSM if s(r[M.MED]) == m) >= MIN_N],
           key=lambda x: x[5])[:10],
    "Each molecule against its own denominator. Restricted to molecules with "
    f"{MIN_N}+ answered observations. The full ranking, every medicine and "
    "strength, is the all-medicines table on this page.")

_fac_basket = []
for f in sorted(M.FACILITIES):
    rs = [r for r in ANSM if s(r["_p"][M.FAC]) == f]
    if not rs:
        continue
    av = sum(1 for r in rs if M.yes(r[M.AVAIL]))
    _fac_basket.append([f, s(rs[0]["_p"][M.DIST]), rs[0]["_lvl"], len(rs), av,
                        len(rs) - av, pct(av, len(rs))])
_fac_basket.sort(key=lambda x: x[6])
add(1, "1.3", "Facility basket availability - the twenty weakest facilities",
    ["Facility", "District", "Level of care", "Expected assessed", "Available",
     "Missing", "FACILITY AVAILABILITY %"], _fac_basket[:20],
    "One rate per facility over its OWN basket. Mean facility availability is "
    f"{pct(sum(x[4] for x in _fac_basket), sum(x[3] for x in _fac_basket)):.1%} "
    f"pooled, against an unweighted mean of "
    f"{statistics.mean(x[6] for x in _fac_basket):.1%} across "
    f"{len(_fac_basket)} facilities - the prompt asks for both, because "
    f"baskets are not identical.")

_complete = sum(1 for x in _fac_basket if x[5] == 0)
add(1, "1.4", "Complete basket availability",
    ["Measure", "Value"],
    [["Facilities with a complete assessment", len(_fac_basket)],
     ["Facilities where EVERY assessed medicine was available", _complete],
     ["COMPLETE BASKET AVAILABILITY", pct(_complete, len(_fac_basket))],
     ["Facilities missing exactly one medicine",
      sum(1 for x in _fac_basket if x[5] == 1)],
     ["Facilities missing five or more",
      sum(1 for x in _fac_basket if x[5] >= 5)]],
    "A facility counts only if every medicine in its own assessed basket was "
    "present on the visit day.")

# ---------------------------------------------------------------
# Method 2 - managed products only - for the availability page's switch
# ---------------------------------------------------------------
# The same four tables on the managed-only base, shown in place of 1.1-1.4
# when the page's "Availability method" filter is set to 2. Dashboard only:
# the workbook keeps the expected-medicines method.
MOLS_B = MOL.managed_base(MED_ALL)
for r in MOLS_B:
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])


def _avail_tables_b():
    av = [r for r in MOLS_B if M.yes(r[M.AVAIL])]
    em = [r for r in MOLS_B if s(r[M.PROGRAM]) == MOL.ESSENTIAL]
    pg = [r for r in MOLS_B if s(r[M.PROGRAM]) != MOL.ESSENTIAL]
    out = {}
    out["1.1"] = (
        "Overall tracer medicine availability - managed products only",
        ["Measure", "Value"],
        [["MANAGED AND ANSWERED (DENOMINATOR)", len(MOLS_B)],
         ["  Essential Medicines", len(em)],
         ["  Programme products", len(pg)],
         ["On the shelf (NUMERATOR)", len(av)],
         ["Managed but empty on the visit day", len(MOLS_B) - len(av)],
         ["OVERALL AVAILABILITY", pct(len(av), len(MOLS_B))],
         ["  Essential Medicines",
          pct(sum(1 for r in em if M.yes(r[M.AVAIL])), len(em))],
         ["  Programme products",
          pct(sum(1 for r in pg if M.yes(r[M.AVAIL])), len(pg))],
         ["Memo: expected medicines (method 1)", pct(len(AVM), len(APPM))],
         ["Facilities", len({s(r["_p"][M.FAC]) for r in MOLS_B})],
         ["Molecules", len({s(r[M.MED]) for r in MOLS_B})]],
        "METHOD 2. Denominator: every molecule the facility manages and "
        "answered for, whatever its programme and whatever the EML says "
        "about its level. A medicine the facility does not stock is out of "
        "scope, not unavailable - so this measures the shelf, not access. "
        f"{len(GAPM):,} Essential Medicines authorised but not managed, "
        "counted unavailable under method 1, drop out here.")
    meds = defaultdict(list)
    for r in MOLS_B:
        meds[s(r[M.MED])].append(r)
    out["1.2"] = (
        "Individual medicine availability - ten least available "
        "(managed products only)",
        ["Medicine", "Programme", "Facilities", "Answered", "Available",
         "AVAILABLE %"],
        sorted([[m, s(rs[0][M.PROGRAM]),
                 len({s(r["_p"][M.FAC]) for r in rs}), len(rs),
                 sum(1 for r in rs if M.yes(r[M.AVAIL])),
                 pct(sum(1 for r in rs if M.yes(r[M.AVAIL])), len(rs))]
                for m, rs in meds.items() if len(rs) >= MIN_N],
               key=lambda x: x[5])[:10],
        f"Each molecule over the facilities that manage it. Restricted to "
        f"molecules with {MIN_N}+ answered observations.")
    fb = []
    for f in sorted(M.FACILITIES):
        rs = [r for r in MOLS_B if s(r["_p"][M.FAC]) == f]
        if not rs:
            continue
        a = sum(1 for r in rs if M.yes(r[M.AVAIL]))
        fb.append([f, s(rs[0]["_p"][M.DIST]), rs[0]["_lvl"], len(rs), a,
                   len(rs) - a, pct(a, len(rs))])
    fb.sort(key=lambda x: x[6])
    out["1.3"] = (
        "Facility basket availability - the twenty weakest facilities "
        "(managed products only)",
        ["Facility", "District", "Level of care", "Managed assessed",
         "Available", "Missing", "FACILITY AVAILABILITY %"], fb[:20],
        "One rate per facility over the medicines it manages. Pooled "
        f"{pct(sum(x[4] for x in fb), sum(x[3] for x in fb)):.1%}, "
        f"unweighted mean {statistics.mean(x[6] for x in fb):.1%} across "
        f"{len(fb)} facilities.")
    comp = sum(1 for x in fb if x[5] == 0)
    out["1.4"] = (
        "Complete basket availability (managed products only)",
        ["Measure", "Value"],
        [["Facilities with a complete assessment", len(fb)],
         ["Facilities where EVERY managed medicine was available", comp],
         ["COMPLETE BASKET AVAILABILITY", pct(comp, len(fb))],
         ["Facilities missing exactly one medicine",
          sum(1 for x in fb if x[5] == 1)],
         ["Facilities missing five or more",
          sum(1 for x in fb if x[5] >= 5)]],
        "A facility counts only if every medicine it manages was present on "
        "the visit day.")
    return out


AVAIL_B = _avail_tables_b()


add(1, "1.5", "Availability by programme (therapeutic category)",
    AVAIL_HDR, grouped(lambda r: s(r[M.PROGRAM])),
    "The tool carries programme and drug class; 'therapeutic category' is not "
    "a collected variable, so these are the available equivalents.")
add(1, "1.6", "Availability by level of care", AVAIL_HDR,
    grouped(lambda r: r["_lvl"], keys=LEVELS),
    "Health post = SGHP + HP; Health centre = HC + MHC; hospitals separate.")
add(1, "1.7", "Availability by province", AVAIL_HDR,
    grouped(lambda r: s(r["_p"][M.PROV])))
add(1, "1.8", "Availability by district", AVAIL_HDR,
    sorted(grouped(lambda r: s(r["_p"][M.DIST])),
           key=lambda x: (x[7] is None, x[7])),
    "Worst availability first.")

add(1, "1.9", "Medicines with a NATIONWIDE STOCK-OUT",
    ["Medicine", "Programme", "Drug class", "Facilities reporting it",
     "Districts", "Facilities managing it", "SHARE OF MANAGING FACILITIES"],
    NS_TABLE,
    f"The {len(ns_rows)} rows the instrument classifies as a nationwide "
    f"stock-out - the medicine was not on the shelf AND not available at RMS "
    f"Central - covering {len(NS_TABLE)} medicines at "
    f"{len({s(r['_p'][M.FAC]) for r in ns_rows})} facilities. These are the "
    "shortages no facility-level or district-level action can fix: the "
    "national store did not have the product either. The last column is the "
    "share of facilities managing that medicine which reported it, so a "
    "medicine stocked by few facilities does not look national on a handful "
    "of reports.")

# Why medicines are not managed - at MOLECULE level, like every availability
# figure on this page: a molecule counts as not managed only when the
# facility holds NO strength of it, so "they manage the 500mg" never shows
# up as a gap. Each not-managed molecule takes the coded reason from its
# items (not_managed.py), the first one actually explained.
_items_by_mol = defaultdict(list)
for _r in MED_ALL:
    _items_by_mol[(s(_r["_p"][M.FAC]), MOL.key(_r[M.MED]))].append(_r)


def _mol_reason(r):
    themes = [NM.theme(i[M.WHY_NOT])
              for i in _items_by_mol[(s(r["_p"][M.FAC]), r["_mol"])]
              if M.no(i[M.MANAGED])]
    told = [t for t in themes if t != NM.NO_REASON]
    return told[0] if told else NM.NO_REASON


_WHY_EM = [(r, _mol_reason(r)) for r in GAPM]
_WHY_PG = [(r, _mol_reason(r)) for r in MOL.collapse(MED_ALL)
           if s(r[M.PROGRAM]) != MOL.ESSENTIAL and not M.yes(r[M.MANAGED])]


def _why_rows():
    em, pg = Counter(t for _, t in _WHY_EM), Counter(t for _, t in _WHY_PG)
    meds = defaultdict(Counter)
    for r, t in _WHY_EM + _WHY_PG:
        meds[t][s(r[M.MED])] += 1
    last = WHY_LAST
    order = sorted(sorted(t for t in set(em) | set(pg)
                           if t not in last and t not in WHY_HIDDEN),
                   key=lambda t: -(em[t] + pg[t])) + \
        [t for t in last if t in em or t in pg]
    rows = []
    for t in order:
        top = ", ".join(f"{m[:30]} ({n})" for m, n in meds[t].most_common(3))
        rows.append([t, em[t], pct(em[t], len(_WHY_EM)), pg[t],
                     pct(pg[t], len(_WHY_PG)), em[t] + pg[t], top])
    rows.append(["TOTAL not managed",
                 len(_WHY_EM), 1.0, len(_WHY_PG), 1.0,
                 len(_WHY_EM) + len(_WHY_PG), ""])
    return rows


# Every reason is shown (user request 2026-10-16), including the medicines
# with no reason recorded, so the rows add up to the TOTAL. They go last,
# after "Other". WHY_HIDDEN stays as an (empty) switch.
WHY_HIDDEN = ()
WHY_LAST = (NM.OTHER, NM.NO_REASON)
_why_hid = Counter(t for _, t in _WHY_EM + _WHY_PG)


add(1, "1.10", "Why medicines are not managed",
    ["Reason (coded from the free-text answer)",
     "Essential Medicines - authorised, not managed",
     "% of these", "Programme products - not managed", "% of these",
     "Total", "Medicines citing it most"],
    _why_rows(),
    "One row per facility x molecule the facility does not stock at any "
    "strength. ESSENTIAL MEDICINES columns: authorised at the facility's "
    "level, so each one is counted UNAVAILABLE on this page. PROGRAMME "
    "columns: allocated products the facility does not hold - outside the "
    "availability denominator, shown for completeness. The free-text answer "
    "is coded by keyword (not_managed.py); 'Other' answers are one-off "
    "phrasings. MFT zoning and faith-based FP are policy, not shortage. "
    f"Listed last: {_why_hid[NM.NO_REASON]:,} with no reason recorded - the "
    "rows add up to the TOTAL.")

# =============== ANALYSIS 2 =====================================
_cls = Counter(s(r[M.STOCKCLASS]) for r in MED_ALL if not M.blank(r[M.STOCKCLASS]))
_valid = sum(_cls.values())
# ---------------------------------------------------------------
# stock adequacy: the base for Analysis 2
# ---------------------------------------------------------------
# TWO CHANGES, 2026-10-13.
#
# 1. ROWS WITH ZERO CONSUMPTION ARE EXCLUDED. Months of stock is quantity
#    divided by average monthly consumption, so at AMC = 0 it is undefined -
#    not infinite, and not overstock. 1,534 of the 9,870 managed rows are in
#    that state. Reporting them inside the distribution made every adequacy
#    rate a share of a denominator that contained rows the measure cannot
#    describe. They are counted once, as an exclusion, and then left out.
#
# 2. STOCK-OUT IS DEFINED BY AVAILABILITY, NOT BY THE FORM'S LABEL. The
#    instrument's stock_classification said "Stock out" on 643 rows, but it
#    was computed before the availability correction and disagrees with the
#    answer the facility gave: 129 rows it calls "Understock" and 44 it
#    leaves blank were NOT on the shelf. A medicine that is not there is
#    stocked out, whatever the derived field says, so the availability
#    answer governs.
# EVERY STOCK-OUT COUNTS, whatever the consumption figure. A medicine that
# was out of stock for the whole window often records zero consumption
# BECAUSE it was never there to dispense - excluding those rows dropped 380
# of the 816 stock-outs, nearly half, and understated the rate. Zero
# consumption is a consequence of the stock-out, not a reason to ignore it.
#
# The AMC = 0 exclusion therefore applies ONLY to rows that were on the
# shelf, where months of stock cannot be computed and the medicine was
# present anyway. Those 1,154 rows carry no adequacy information at all.
SA_OUT = [r for r in MG if M.no(r[M.AVAIL])]
SA_IN = [r for r in MG if M.yes(r[M.AVAIL]) and (M.num(r[M.AMC]) or 0) > 0]
AMC_ZERO = [r for r in MG
            if M.yes(r[M.AVAIL]) and (M.num(r[M.AMC]) or 0) == 0]
AMC_POS = SA_OUT + SA_IN


def _mos(r):
    return M.num(r[M.MOS])


# Teaching and referral hospitals - level 2 teaching hospitals and the two
# national referral hospitals (CHUK, CHUB) - are held to a wider band: 1 to 3
# months of stock is adequate there (CHUK/CHUB 2026-10-16; level 2 teaching
# hospitals added 2026-11-05, user decision). Every other facility keeps 1 to
# 2 months. _band() is the one place the cut-offs live - 0 understock,
# 1 adequate, 2 overstock.
WIDE_LEVELS = {"Level 2 teaching hospital", "National referral hospital"}


def _over_at(r):
    return 3 if r["_lvl"] in WIDE_LEVELS else 2


def _band(r):
    m = _mos(r) or 0
    return 0 if m < 1 else (1 if m <= _over_at(r) else 2)


_sa_under = [r for r in SA_IN if _band(r) == 0]
_sa_plan = [r for r in SA_IN if _band(r) == 1]
_sa_over = [r for r in SA_IN if _band(r) == 2]
_sa_n = len(AMC_POS)

# STOCK ADEQUACY COMES FIRST, and it is measured on STOCKED rows only.
# Under / adequate / over describe how much of a medicine a facility is
# holding, which is a question that only exists where the medicine is there
# at all. Folding stock-outs into the same denominator mixes "we hold the
# wrong amount" with "we hold none", and the two call for different action.
# Stock-outs are reported immediately below, on their own base.
SA_SHELF = SA_IN
_ad_under = [r for r in SA_SHELF if _band(r) == 0]
_ad_plan = [r for r in SA_SHELF if _band(r) == 1]
_ad_over = [r for r in SA_SHELF if _band(r) == 2]
_ad_n = len(SA_SHELF)
ADEQ = ["Understock", "Adequate", "Overstock"]
# Emergency order point: an understocked medicine holding half a month or
# less (user request 2026-10-16). Shown as a highlighted sub-count of
# Understock in FIG 2.1.
EOP = 0.5


def _eop(r):
    return _band(r) == 0 and (_mos(r) or 0) <= EOP


def eop_line(n_eop, n_under):
    return (f"{n_eop:,} ({n_eop / n_under:.1%}) at or below {EOP} months - "
            f"emergency order point" if n_under else "")


# FIG 2.1-2.4 carry the stock-outs as a fourth category (2026-10-16), on
# the same base as table 2.1: every stock-out plus every stocked row with
# consumption. Tables 2A/2B stay on stocked rows.
ADEQ4 = ["Stocked out"] + ADEQ
STOCKOUT_COL = "#8b1a3a"
ADEQ4_COLS = [STOCKOUT_COL] + CH.STACK_COLS
SA_OUT_IDS = {id(r) for r in SA_OUT}


def adequacy_mix4(keep=lambda r: True):
    """[stocked out, under, adequate, over] over SA_OUT + SA_SHELF."""
    return ([sum(1 for r in SA_OUT if keep(r))]
            + adequacy_mix([r for r in SA_SHELF if keep(r)]))


def _mix4_rows(keyfn, rows_keys, min_n):
    out = []
    for k in rows_keys:
        v = adequacy_mix4(lambda r: keyfn(r) == k)
        if sum(v) >= min_n:
            out.append([k, v])
    return out


ADEQ_BLURBS = ["under one month of stock on hand - cannot meet a month of "
               "demand",
               "one to two months - the level the form itself sets as the "
               "target; one to three at teaching and referral hospitals",
               "more than two months (three at teaching and referral hospitals) - capital tied "
               "up, and expiry risk"]
ADEQ4_BLURBS = ["not on the shelf on the visit day - no stock to dispense"] \
    + ADEQ_BLURBS


def _med_adequacy(floor=MIN_N):
    """Per medicine, the under/over mix among its STOCKED rows."""
    g = defaultdict(list)
    for r in SA_SHELF:
        g[s(r[M.MED])].append(r)
    out = []
    for m, v in g.items():
        if len(v) < floor:
            continue
        u = sum(1 for r in v if _band(r) == 0)
        o = sum(1 for r in v if _band(r) == 2)
        out.append([m, s(v[0][M.PROGRAM]), len(v),
                    len({s(r["_p"][M.FAC]) for r in v}),
                    u, pct(u, len(v)), o, pct(o, len(v)),
                    med([_mos(r) for r in v])])
    return out


def _adequacy_med_tables():
    """The medicines at each end of the distribution, after FIG 2.4."""
    rows = _med_adequacy()
    hdr = ["Medicine", "Programme", "Stocked rows", "Facilities",
           "Understocked", "UNDERSTOCK %", "Overstocked", "OVERSTOCK %",
           "Median months"]
    under = sorted(rows, key=lambda x: -(x[5] or 0))[:15]
    over = sorted(rows, key=lambda x: -(x[7] or 0))[:15]
    return (
        '<section class="ind">'
        '<h4><span class="mid">TABLE 2A</span>Medicines most often '
        'understocked</h4>'
        f'<p class="note">Among the rows where the medicine was ON THE SHELF, '
        f'the share holding under one month of stock. Ranked worst first, '
        f'restricted to the {len(rows)} medicines with at least {MIN_N} '
        'stocked rows so a rate is not built on a handful of facilities. '
        'Stock-outs are not counted here - a medicine that is absent is in '
        '2.1, not in this table.</p>'
        + '<div id="sx-t2a">' + html_table(hdr, under) + '</div></section>'
        '<section class="ind">'
        '<h4><span class="mid">TABLE 2B</span>Medicines most often '
        'overstocked</h4>'
        + '<div id="sx-t2b">' + html_table(hdr, over) + '</div></section>')


def adequacy_mix(rows):
    return [sum(1 for r in rows if _band(r) == b) for b in (0, 1, 2)]


add(2, "2.1", "Months of stock - distribution and stock status",
    ["Stock category", "Definition", "Observations", "% of base"],
    [["Stocked out", "not available on the visit day", len(SA_OUT),
      pct(len(SA_OUT), _sa_n)],
     ["Understock", "on the shelf, under 1 month of stock", len(_sa_under),
      pct(len(_sa_under), _sa_n)],
     ["Adequate stock", "1 to 2 months (1 to 3 at teaching and referral hospitals)", len(_sa_plan),
      pct(len(_sa_plan), _sa_n)],
     ["Overstock", "more than 2 months (3 at teaching and referral hospitals)", len(_sa_over),
      pct(len(_sa_over), _sa_n)],
     ["", "", "", None],
     ["BASE", "every stock-out, plus stocked rows with consumption",
      _sa_n, 1.0],
     ["Excluded: on the shelf with zero consumption",
      "months of stock is undefined at AMC = 0 and the medicine was present",
      len(AMC_ZERO), None],
     ["Median months of stock", "stocked rows only",
      med([_mos(r) for r in SA_IN]), None]],
    f"STOCK-OUT IS THE AVAILABILITY ANSWER, not the form's "
    f"stock_classification field. That field was computed before the "
    f"availability correction and labels only "
    f"{sum(1 for r in MG if s(r[M.STOCKCLASS]) == 'Stock out')} rows a stock "
    f"out where {len(SA_OUT)} medicines were in fact not on the shelf. "
    f"EVERY STOCK-OUT IS COUNTED, whatever the consumption figure: a medicine "
    f"out for the whole window often records zero consumption because it was "
    f"never there to dispense, so excluding those rows would drop the longest "
    f"stock-outs and understate the rate. The AMC = 0 exclusion applies only "
    f"to the {len(AMC_ZERO):,} rows that were ON THE SHELF with no recorded "
    f"consumption, where months of stock cannot be computed and the medicine "
    f"was present anyway. The four categories are mutually exclusive and sum "
    f"to the base.")

# ---------------------------------------------------------------
# inventory accuracy: who is in the denominator
# ---------------------------------------------------------------
# Accuracy compares a physical count against the quantity the SYSTEM holds,
# so it only means something where there is a system. The form asks for the
# name of the digital system immediately after the system quantity, and some
# facilities never answer it on any row - they write "none", "N/A", a bare
# number, or leave it blank throughout. Their "system quantity" came from a
# paper register or from nothing at all, and scoring them for the accuracy
# of a digital record they do not keep is not a fair measure.
#
# Excluded 2026-10-13, ROW BY ROW: every record that does not name a digital
# system is dropped, not only the facilities that never named one. The
# earlier rule kept a row whose own answer was "none" as long as the facility
# had named a system somewhere else in its basket, which measured the wrong
# thing - accuracy is a property of the record for THAT medicine, and where
# no system was named for it there is no digital record to be accurate.
#
# The raw field is never literally blank on this base: the form only asks it
# where a system quantity was entered. What it holds instead is "none",
# "N/A", "No system", a bare number, or an unrelated remark - all of which
# S8 resolves to no system named, and all of which are excluded here.
def _named_system(r):
    """Did THIS record name a digital system? S8 has already normalised the
    free text, so an empty cleaned value is the whole test."""
    return bool(s(r.get("_emr_clean")))


_acc_all = [r for r in AV
            if not M.blank(r[M.QTY_SYS]) and not M.blank(r[M.QTY_PHYS])]
_acc = [r for r in _acc_all if _named_system(r)]
_acc_dropped = len(_acc_all) - len(_acc)
NO_SYSTEM = sorted({s(r["_p"][M.FAC]) for r in _acc_all
                    if not _named_system(r)})
_FAC_ALL_BLANK = sorted(
    f for f in NO_SYSTEM
    if not any(_named_system(r) for r in _acc_all
               if s(r["_p"][M.FAC]) == f))
_exact = sum(1 for r in _acc
             if (M.num(r[M.QTY_PHYS]) or 0) == (M.num(r[M.QTY_SYS]) or 0))
_diff = [(M.num(r[M.QTY_PHYS]) or 0) - (M.num(r[M.QTY_SYS]) or 0) for r in _acc]
_ape = [abs(d) / (M.num(r[M.QTY_PHYS]) or 1) for d, r in zip(_diff, _acc)
        if (M.num(r[M.QTY_PHYS]) or 0) > 0]
add(2, "2.2-2.4", "Inventory record accuracy",
    ["Metric", "Value", "Basis"],
    [["Observations with both quantities", len(_acc),
      "available rows only - counts are not recorded where the medicine is "
      "absent"],
     ["Physical equals recorded (exact accuracy)", _exact, ""],
     ["EXACT ACCURACY RATE", pct(_exact, len(_acc)), ""],
     ["Physical exceeds recorded", sum(1 for d in _diff if d > 0),
      "the shelf holds more than the system"],
     ["Recorded exceeds physical", sum(1 for d in _diff if d < 0),
      "the system overstates - hides a shortage, reorder never triggers"],
     ["Median discrepancy (physical - recorded)", med(_diff), "units"],
     ["Median absolute percentage error", med(_ape),
      "|physical - recorded| / physical, physical > 0"],
     ["Within +/- 5%",
      sum(1 for e in _ape if e <= 0.05), f"of {len(_ape)} rows with stock > 0"]],
    "DENOMINATOR CORRECTION: the prompt implies all assessed observations, but "
    "physical and recorded counts exist only where the medicine is available. "
    f"The base is {len(_acc):,} available rows, not {len(MG):,} managed rows. "
    f"EXCLUSION: every record that does not name a digital recording system "
    f"is dropped - {_acc_dropped:,} rows - because where no system was named "
    f"there is no digital record for a physical count to be checked against. "
    f"The field is never literally blank on this base; it holds 'none', "
    f"'N/A', a bare number or an unrelated remark instead. The rows come from "
    f"{len(NO_SYSTEM)} facilities, of which {len(_FAC_ALL_BLANK)} named no "
    f"system anywhere in their basket: "
    + ", ".join(_FAC_ALL_BLANK)
    + ". All of them remain in every other indicator.")

add(2, "2.5", "Inventory accuracy by level of care",
    ["Level of care", "Rows counted", "Exact match", "EXACT ACCURACY %",
     "Median absolute % error"],
    [[l, len([r for r in _acc if r["_lvl"] == l]),
      sum(1 for r in _acc if r["_lvl"] == l
          and (M.num(r[M.QTY_PHYS]) or 0) == (M.num(r[M.QTY_SYS]) or 0)),
      pct(sum(1 for r in _acc if r["_lvl"] == l
              and (M.num(r[M.QTY_PHYS]) or 0) == (M.num(r[M.QTY_SYS]) or 0)),
          len([r for r in _acc if r["_lvl"] == l])),
      med([abs((M.num(r[M.QTY_PHYS]) or 0) - (M.num(r[M.QTY_SYS]) or 0))
           / (M.num(r[M.QTY_PHYS]) or 1) for r in _acc if r["_lvl"] == l
           and (M.num(r[M.QTY_PHYS]) or 0) > 0])]
     for l in LEVELS if any(r["_lvl"] == l for r in _acc)])

def _exact(r):
    return (M.num(r[M.QTY_PHYS]) or 0) == (M.num(r[M.QTY_SYS]) or 0)


# Quantity accuracy (user request 2026-10-16), shown beside the exact-match
# dial in FIG 2.5: per record (1 - |system - physical| / system), floored at
# 0 - a count off by more than the whole system quantity is 0%, not
# negative - then averaged. A system quantity of 0 scores 0% (user decision
# 2026-10-16), so every record is in the measure - unless the shelf is also
# 0: system and physical both zero is a perfect match and scores 100% (user
# decision 2026-10-20).
def _qacc(r):
    sq = M.num(r[M.QTY_SYS]) or 0
    if sq <= 0:
        return 1.0 if (M.num(r[M.QTY_PHYS]) or 0) == 0 else 0.0
    return max(0.0, 1 - abs(sq - (M.num(r[M.QTY_PHYS]) or 0)) / sq)


def _dial_card(value, caption, name, count_html):
    return (f'<div class="dcard">{CH.dial(value, caption=caption)}'
            f'<p class="dname">{esc(name)}</p>'
            f'<p class="dcount">{count_html}</p></div>')


def acc_dials(rows):
    """FIG 2.5: exact-match dial and quantity-accuracy dial side by side."""
    hit = sum(1 for r in rows if _exact(r))
    q = [v for v in (_qacc(r) for r in rows) if v is not None]
    return ('<div class="dgrid">'
            + _dial_card(hit / len(rows) if rows else 0,
                         "of records match the shelf", "Exact match",
                         f"<b>{hit:,}</b> of {len(rows):,} records")
            + _dial_card(statistics.mean(q) if q else 0,
                         "average quantity accuracy", "Quantity accuracy",
                         f"<b>{len(q):,}</b> records; a system quantity "
                         f"of 0 counts as 0%")
            + '</div>')


# FIG 2.6-2.9 put the two measures side by side (2026-10-16): exact match
# on the left, average quantity accuracy on the right, each its own chart.
# The right chart keeps the left chart's row order so a group reads across.
PAIR_W, PAIR_LW = 620, 165


def _pair_html(left, right, lt="Exact match",
               rt="Average quantity accuracy"):
    return (f'<div class="pair2"><div><p class="ptitle">{esc(lt)}</p>{left}'
            f'</div><div><p class="ptitle">{esc(rt)}</p>{right}</div></div>')


def acc_pair(fid, title, keyfn, lbox, rbox, keys=None, sort=True):
    left = _bar_rows(rate_by(_acc, keyfn, _exact, mfac, keys=keys,
                             sort=sort))
    g = defaultdict(list)
    for r in _acc:
        g[keyfn(r)].append(_qacc(r))
    right = [[k, n, statistics.mean(g[k])] for k, n, _ in left]
    return (f'<section class="ind"><h4><span class="mid">{esc(fid)}</span>'
            f'{esc(title)}</h4>'
            + _pair_html(
                f'<div id="{lbox}">' + CH.bars(left, width=PAIR_W,
                                               label_w=PAIR_LW,
                                               hit="accurate",
                                               base="records") + '</div>',
                f'<div id="{rbox}">' + CH.mean_bars(right, width=PAIR_W,
                                                    label_w=PAIR_LW)
                + '</div>')
            + '</section>')


add(2, "2.6", "Inventory accuracy by province", RATE_HDR,
    rate_by(_acc, lambda r: s(r["_p"][M.PROV]), _exact, mfac),
    "Exact match between the physical count and the system count, weakest "
    "province first. Same exclusion as 2.4-2.6: facilities that never named a "
    "digital recording system are not in this table.")

add(2, "2.7", "Inventory accuracy by management type", RATE_HDR,
    rate_by(_acc, mmgt, _exact, mfac),
    "Who runs the facility, as recorded on the form. A collected variable, "
    "not derived from the level of care.")

add(2, "2.8", "Inventory accuracy by programme", RATE_HDR,
    rate_by(_acc, lambda r: s(r[M.PROGRAM]) or "(none)", _exact, mfac),
    "Vertical programmes keep their own records in several facilities, so "
    "this separates a facility-wide record problem from a programme one.")

_mos_rows = [r for r in MG if M.num(r[M.MOS]) is not None]


def _understock(r):
    return (M.num(r[M.MOS]) or 0) < 1


# =============== ANALYSIS 3: MEDICINE CONSUMPTION ===============
# Six indicators from section 2B of the analysis plan - C1, C2, C3, C4, C6
# and C9. The plan's own numbering skips C5, C7 and C8; they are not in the
# sheet, so they are not invented here.
#
# Consumption stays at ITEM level, not molecule level. Adding 500mg and 850mg
# tablets gives a number with no meaning, which is the same rule the rest of
# the quantity work follows, so these rows are M.MEDROWS and not MOLS.
CMONTHS = ("June", "July", "August")


def cvals(r):
    """The three monthly quantities, or None if any month is missing."""
    v = [M.num(r[M.CONS[m]]) for m in CMONTHS]
    return None if any(x is None for x in v) else v


CROWS = []
for r in MED_ALL:
    v = cvals(r)
    if v is None:
        continue
    r["_c"] = v
    r["_ctot"] = sum(v)
    r["_amc"] = sum(v) / len(v)
    CROWS.append(r)
C_MEDS = sorted({s(r[M.MED]) for r in CROWS})

def _amc_stats(m):
    """One medicine's AMC across the facilities that report it."""
    rs = [r for r in CROWS if s(r[M.MED]) == m]
    v = [r["_amc"] for r in rs]
    mean = statistics.mean(v)
    sd = statistics.stdev(v) if len(v) > 1 else None
    # Total AMC (2026-10-18): the AMCs of every facility reporting the
    # medicine, summed - its national monthly consumption. Tables 3.1-3.1c
    # rank on it.
    return [m, s(rs[0][M.PROGRAM]), len({s(r["_p"][M.FAC]) for r in rs}),
            round(sum(v), 1), round(mean, 1), med(v),
            round(sd, 1) if sd is not None else None,
            round(sd / mean * 100, 1) if sd is not None and mean else None,
            round(sum(r["_ctot"] for r in rs))]


AMC_STATS = [_amc_stats(m) for m in C_MEDS if m]
AMC_HDR = ["Medicine", "Programme", "Facilities", "Total AMC", "Mean AMC",
           "Median AMC", "SD of AMC", "CV", "Total Jun-Aug"]

add(3, "3.1", "Average monthly consumption - the twenty highest", AMC_HDR,
    sorted(AMC_STATS, key=lambda x: -x[3])[:20],
    "AMC = the three monthly quantities divided by the number of months with "
    "a value. The form carries its own avg_monthly_consumption and it equals "
    f"this on all {len(CROWS):,} complete rows, so the two agree. Mean and "
    "median are both given because consumption is heavily skewed and they "
    "rank medicines differently.")

add(3, "3.1b", "Average monthly consumption - the twenty lowest", AMC_HDR,
    sorted(AMC_STATS, key=lambda x: x[3])[:20],
    "The other end of the same ranking, least consumed first. A low mean "
    "with few facilities is a thin estimate - read it with the Facilities "
    "column.")

add(3, "3.1c", "Average monthly consumption - every medicine", AMC_HDR,
    sorted(AMC_STATS, key=lambda x: -x[3]),
    "One row per medicine, highest total AMC first. Total AMC = the AMCs of "
    "every facility reporting the medicine, summed. SD is the standard "
    "deviation of AMC across the facilities that report the medicine; CV = "
    "SD / mean x 100, a percentage - how much facilities differ from one "
    "another, not how much one facility varies month to month (that is "
    "3.3). Units differ between medicines, so compare CV, not SD, across "
    "rows.")

add(3, "3.2", "Total period consumption by level of care and programme",
    ["Group", "Medicine-facility pairs", "Total Jun-Aug", "Median per pair",
     "Share of all consumption"],
    [[l, sum(1 for r in CROWS if r["_lvl"] == l),
      round(sum(r["_ctot"] for r in CROWS if r["_lvl"] == l)),
      med([r["_ctot"] for r in CROWS if r["_lvl"] == l]),
      pct(sum(r["_ctot"] for r in CROWS if r["_lvl"] == l),
          sum(r["_ctot"] for r in CROWS))]
     for l in LEVELS if any(r["_lvl"] == l for r in CROWS)]
    + [["", "", "", "", ""]]
    + [[p, sum(1 for r in CROWS if s(r[M.PROGRAM]) == p),
        round(sum(r["_ctot"] for r in CROWS if s(r[M.PROGRAM]) == p)),
        med([r["_ctot"] for r in CROWS if s(r[M.PROGRAM]) == p]),
        pct(sum(r["_ctot"] for r in CROWS if s(r[M.PROGRAM]) == p),
            sum(r["_ctot"] for r in CROWS))]
       for p in sorted({s(r[M.PROGRAM]) for r in CROWS})],
    "Quantities summed across June, July and August. Units are not "
    "comparable between medicines - a tablet and a vial both count as one - "
    "so a total across medicines measures volume handled, not therapeutic "
    "quantity, and the shares matter more than the absolute numbers.")


# AMC by programme (user request 2026-10-19), the 3.1c columns with one row
# per programme: each facility x medicine pair's AMC is one observation.
# Units differ between medicines, so totals are volume handled, not doses.
PROG_AMC_HDR = ["Programme", "Medicines", "Facilities"] + AMC_HDR[3:]


def _prog_amc(p):
    rs = [r for r in CROWS if s(r[M.PROGRAM]) == p]
    v = [r["_amc"] for r in rs]
    mean = statistics.mean(v)
    sd = statistics.stdev(v) if len(v) > 1 else None
    return [p, len({s(r[M.MED]) for r in rs}),
            len({s(r["_p"][M.FAC]) for r in rs}),
            round(sum(v), 1), round(mean, 1), med(v),
            round(sd, 1) if sd is not None else None,
            round(sd / mean * 100, 1) if sd is not None and mean else None,
            round(sum(r["_ctot"] for r in rs))]


add(3, "3.2p", "Average monthly consumption by programme", PROG_AMC_HDR,
    sorted([_prog_amc(p) for p in sorted({s(r[M.PROGRAM]) for r in CROWS})
            if p], key=lambda x: -x[3]),
    "One row per programme, highest total AMC first; each facility x "
    "medicine pair is one observation.")


def cv(v):
    """Coefficient of variation, as a percentage. Undefined at zero mean."""
    m = statistics.mean(v)
    return None if m == 0 else statistics.pstdev(v) / m * 100


_cvrows = [(r, cv(r["_c"])) for r in CROWS]
_cvok = [(r, c) for r, c in _cvrows if c is not None]
add(3, "3.3", "Consumption variability by therapeutic class",
    ["Drug class", "Pairs with a CV", "Median CV %", "Stable (CV<25%)",
     "Erratic (CV>=50%)", "SHARE ERRATIC"],
    sorted([[c, len(g),
             med([x for _, x in g]),
             sum(1 for _, x in g if x < 25),
             sum(1 for _, x in g if x >= 50),
             pct(sum(1 for _, x in g if x >= 50), len(g))]
            for c, g in [(c, [(r, x) for r, x in _cvok
                              if (s(r[M.DCLASS]) or "(unclassified)") == c])
                         for c in sorted({s(r[M.DCLASS]) or "(unclassified)"
                                   for r, _ in _cvok})]
            if len(g) >= MIN_N], key=lambda x: -(x[5] or 0)),
    f"CV = standard deviation of the three monthly values / their mean x 100. "
    f"It is undefined where a pair consumed nothing at all, so the base is "
    f"{len(_cvok):,} of {len(CROWS):,} complete pairs. A high share erratic "
    "means demand at that class cannot be planned from a three-month average, "
    "which is what AMC-based ordering assumes.")

_tr = [r for r in CROWS if r["_c"][0] > 0]
add(3, "3.4", "Consumption trend, August against June",
    ["Group", "Pairs with June consumption", "Total June", "Total August",
     "TREND RATIO", "Pairs rising", "Pairs falling"],
    [[l, sum(1 for r in _tr if r["_lvl"] == l),
      round(sum(r["_c"][0] for r in _tr if r["_lvl"] == l)),
      round(sum(r["_c"][2] for r in _tr if r["_lvl"] == l)),
      (sum(r["_c"][2] for r in _tr if r["_lvl"] == l)
       / sum(r["_c"][0] for r in _tr if r["_lvl"] == l)
       if sum(r["_c"][0] for r in _tr if r["_lvl"] == l) else None),
      sum(1 for r in _tr if r["_lvl"] == l and r["_c"][2] > r["_c"][0]),
      sum(1 for r in _tr if r["_lvl"] == l and r["_c"][2] < r["_c"][0])]
     for l in LEVELS if any(r["_lvl"] == l for r in _tr)],
    "Trend ratio = August quantity / June quantity, computed on the pooled "
    f"totals. Restricted to the {len(_tr):,} pairs that consumed something in "
    "June, because the ratio is undefined against a zero base. Above 1.0 is "
    "rising demand, below 1.0 falling. Three months is a short window and "
    "covers no seasonal cycle, so this is a direction, not a forecast.")

_nm = [r for r in CROWS if r["_ctot"] == 0]
add(3, "3.5", "Non-moving medicines by level of care",
    ["Level of care", "Pairs with complete data", "Zero in all three months",
     "NON-MOVING RATE", "Of those, held in stock"],
    [[l, sum(1 for r in CROWS if r["_lvl"] == l),
      sum(1 for r in _nm if r["_lvl"] == l),
      pct(sum(1 for r in _nm if r["_lvl"] == l),
          sum(1 for r in CROWS if r["_lvl"] == l)),
      sum(1 for r in _nm if r["_lvl"] == l and M.yes(r[M.AVAIL]))]
     for l in LEVELS if any(r["_lvl"] == l for r in CROWS)]
    + [["ALL LEVELS", len(CROWS), len(_nm), pct(len(_nm), len(CROWS)),
        sum(1 for r in _nm if M.yes(r[M.AVAIL]))]],
    "A medicine the facility manages and recorded no consumption for in any "
    "of the three months. The last column is the part that costs money: "
    "stock sitting on a shelf that nothing was dispensed from. The plan also "
    "asks for a slow-moving rate; no threshold is defined for it anywhere in "
    "the instrument or the plan, so only the zero-consumption case is "
    "reported rather than a cut invented here.")

_val = [(r, r["_ctot"] * (M.num(r[M.PRICE]) or 0)) for r in CROWS]
_vtot = sum(v for _, v in _val)
_bymed = sorted(((m, sum(v for r, v in _val if s(r[M.MED]) == m))
                 for m in C_MEDS), key=lambda x: -x[1])
_cum, _pareto = 0.0, []
for i, (m, v) in enumerate(_bymed[:20], start=1):
    _cum += v
    _pareto.append([i, m, round(v), pct(v, _vtot), pct(_cum, _vtot)])
add(3, "3.6", "Consumption value concentration - the twenty largest",
    ["Rank", "Medicine", "Consumption value (RWF)", "Share", "CUMULATIVE"],
    _pareto,
    f"Value = quantity consumed x unit tariff price, over "
    f"{sum(1 for r, _ in _val if M.num(r[M.PRICE]) is not None):,} of "
    f"{len(CROWS):,} pairs that carry a price. Total {round(_vtot):,} RWF. "
    "The plan asks for the top-ten share as an ABC analysis; the top ten are "
    f"{pct(sum(v for _, v in _bymed[:10]), _vtot):.1%} of all consumption "
    "value, and the cumulative column is the Pareto line.")


add(3, "3.7", "Total consumption by province", VOL_HDR,
    vol_by(CROWS, lambda r: s(r["_p"][M.PROV]), lambda r: r["_ctot"], mfac),
    "Volume handled, not therapeutic quantity - a tablet and a vial both "
    "count as one unit, so the share column is the readable part.")

add(3, "3.8", "Non-moving medicines by management type and province",
    RATE_HDR,
    rate_by(CROWS, mmgt, lambda r: r["_ctot"] == 0, mfac)
    + [["", "", "", "", None]]
    + rate_by(CROWS, lambda r: s(r["_p"][M.PROV]), lambda r: r["_ctot"] == 0,
              mfac),
    "Managed, and nothing dispensed in any of the three months. Highest rate "
    "first - the worst group is the one tying up the most shelf space in "
    "product that does not move.")

add(3, "3.9", "Consumption variability by level of care and province",
    ["Group", "Pairs with a CV", "Median CV %", "Erratic (CV>=50%)",
     "SHARE ERRATIC"],
    [[k, len(g), med([x for _, x in g]),
      sum(1 for _, x in g if x >= 50),
      pct(sum(1 for _, x in g if x >= 50), len(g)) if len(g) >= MIN_N else None]
     for k, g in ([(l, [(r, x) for r, x in _cvok if r["_lvl"] == l])
                   for l in LEVELS]
                  + [("", [])]
                  + [(p, [(r, x) for r, x in _cvok
                          if s(r["_p"][M.PROV]) == p])
                     for p in sorted({s(r["_p"][M.PROV]) for r, _ in _cvok})])
     if g or k == ""],
    "A high erratic share means a three-month average is a poor basis for "
    "ordering in that group, which is exactly what AMC-based resupply assumes.")

# =============== ANALYSIS 4 =====================================
def months_affected(r):
    return sum(1 for m in MONTHS if (M.num(r[M.SO_MONTH[m]]) or 0) > 0)


_any_so = [r for r in MG if months_affected(r) > 0]
add(4, "4.1", "Three-month stock-out incidence",
    ["Measure", "Value"],
    [["Facility x medicine pairs with complete history", len(MG)],
     ["Pairs with at least one stock-out day, June-August", len(_any_so)],
     ["STOCK-OUT INCIDENCE", pct(len(_any_so), len(MG))],
     ["Pairs stocked out on the visit day", len(SO)],
     ["Facilities with at least one stock-out in the period",
      len({s(r["_p"][M.FAC]) for r in _any_so})]],
    f"Reference period: {PERIOD}. March-May columns exist in the export but "
    f"are empty on both servers, so the window is three months.")

_ma = Counter(months_affected(r) for r in MG)
add(4, "4.2", "Stock-out frequency - PROXY, months affected",
    ["Months with any stock-out", "Observations", "% of managed"],
    [[k, _ma.get(k, 0), pct(_ma.get(k, 0), len(MG))] for k in (0, 1, 2, 3)],
    "THE EPISODE COUNT THE PROMPT ASKS FOR IS NOT COLLECTED. The form records "
    "stock-out DAYS per month, so two 5-day stock-outs cannot be told from "
    "one 10-day stock-out. This counts months affected, 0-3, and must be "
    "labelled as a proxy, not as frequency.", "Proxy only")

_sod = [M.num(r[M.SO_TOTAL]) or 0 for r in MG]
_cur = [M.num(r[M.SO_NOW]) for r in SO if not M.blank(r[M.SO_NOW])]
add(4, "4.3-3.6", "Stock-out duration",
    ["Metric", "Value", "Basis"],
    [["Total stock-out days, all managed rows", round(sum(_sod)), PERIOD],
     ["Median days per affected pair",
      med([x for x in _sod if x > 0]), f"{len(_any_so)} affected pairs"],
     ["Mean days per affected pair",
      round(statistics.mean([x for x in _sod if x > 0]), 1) if _any_so else None,
      "NOT an episode mean - see 3.2"],
     ["Stock-out days ratio",
      pct(sum(_sod), len(MG) * 92), "total days out / observation-days (92)"],
     ["Currently stocked out at the visit", len(SO), ""],
     ["Median current duration", med(_cur), "days, on currently out rows"],
     ["Out more than 30 days now",
      sum(1 for x in _cur if x and x > 30), "escalation list"]],
    "Metric 3.4 as specified (days per completed episode) has no denominator "
    "without an episode count; the mean per affected pair is given instead "
    "and labelled.")

# The window total is SUMMED FROM THE MONTHS, not read from the form's
# total_stockout_days. That field adds the CURRENT stock-out run - which
# reaches back before June - to the three months, so it double-counts and
# exceeds the 92-day window on 355 rows. The monthly fields are sound.
def _wdays(r):
    return sum(M.num(r[M.SO_MONTH[m]]) or 0 for m in MONTHS)


# AVERAGE stock-out days use the form's total_stockout_days, on the pairs
# where it is above 0 (user request 2026-10-26): FIG 4.4, card pc4-2 and
# the 4.7 "Average days out" column. Incidence, total days out, the 92-day
# rate and the lost-sales days stay on the monthly fields (_wdays).
def _tdays(r):
    return M.num(r[M.SO_TOTAL]) or 0


def _burden(m):
    rs = [r for r in MG if s(r[M.MED]) == m]
    out = [r for r in rs if _wdays(r) > 0]
    days = round(sum(_wdays(r) for r in rs))
    return [m, s(rs[0][M.PROGRAM]), len(rs), len(out),
            pct(len(out), len(rs)), days, len(rs) * 92,
            pct(days, len(rs) * 92),
            round(sum(_tdays(r) for r in tout) / len(tout), 1)
            if (tout := [r for r in rs if _tdays(r) > 0]) else None]


# ---------------------------------------------------------------
# stock-out months with and without a request (user request 2026-10-19)
# ---------------------------------------------------------------
# One observation = a managed facility x medicine x month (June-August) with
# at least one stock-out day. "Requested" = a quantity above 0 requested in
# that same month, read from the raw field (the requested >= delivered rule
# that cleans fill rates is not applied: a request was made either way).
SOREQ_SERIES = ["Requested that month", "No request that month"]
SOREQ_COLS = ["#1f9e7a", "#d9592a"]


def _so_months(r):
    """-> [(month, requested?)] for the months this pair was out."""
    return [(m, (M.num(r[M.REQ[m]]) or 0) > 0) for m in MONTHS
            if (M.num(r[M.SO_MONTH[m]]) or 0) > 0]


def soreq_mix(rows):
    ev = [q for r in rows for _, q in _so_months(r)]
    return [sum(ev), len(ev) - sum(ev)]


def soreq_rows(keyfn, keys=None, min_n=1):
    g = defaultdict(list)
    for r in MG:
        g[keyfn(r)].append(r)
    out = [[k, soreq_mix(g[k])] for k in (keys or sorted(g)) if k]
    out = [x for x in out if sum(x[1]) >= min_n]
    if keys is None:
        out.sort(key=lambda x: -x[1][1] / sum(x[1]))
    return out


def _soreq_fac():
    g = defaultdict(list)
    for r in MG:
        g[mfac(r)].append(r)
    rows = []
    for f in sorted(g):
        req, nreq = soreq_mix(g[f])
        if req + nreq:
            p = g[f][0]["_p"]
            rows.append([f, s(p[M.DIST]), g[f][0]["_lvl"], req + nreq, req,
                         nreq, pct(nreq, req + nreq)])
    rows.sort(key=lambda x: (-x[5], -x[3]))
    return rows


add(4, "4.8", "Facilities with stock-outs - requested in that month or not",
    ["Facility", "District", "Level of care", "Stock-out months",
     "Requested that month", "No request that month", "NO REQUEST %"],
    _soreq_fac(),
    "One row per facility that had at least one stock-out day in June-"
    "August. A stock-out month is one medicine out at least one day in "
    "that month; requested = a quantity above 0 requested in the same "
    "month. Most unrequested stock-out months first.")

add(4, "4.7", "Every medicine assessed, by stock-out burden",
    ["Medicine", "Programme", "Managed pairs", "Pairs ever out",
     "INCIDENCE", "Total days out", "Days possible (denominator)",
     "DAYS OUT RATE", "Average days out"],
    sorted([_burden(m) for m in sorted({s(r[M.MED]) for r in MG})],
           key=lambda x: -(x[4] or 0)),
    f"All {len({s(r[M.MED]) for r in MG})} medicines in the basket, each "
    "against its own denominator, worst incidence first. INCIDENCE is the "
    "share of managed pairs out at least one day in June-August. DAYS "
    "POSSIBLE is managed pairs x 92 days - every day each pair could have "
    "been out - and DAYS OUT RATE is total days out over it. AVERAGE DAYS "
    "OUT is the form's total_stockout_days averaged over the pairs where it "
    "is above 0, the same measure as FIG 4.4. Incidence and duration rank medicines differently - out "
    "briefly but often is not the same problem as out for months - so both "
    "are shown. Total days out and the rate are summed from the three "
    "monthly fields; total_stockout_days also adds the ongoing current run, "
    "so it can exceed 92. Full crosstabs by level, "
    "province, district and management type are in Stockout_analysis.xlsx.")


def _somonths(r):
    return months_affected(r) > 0


add(4, "4.4", "Stock-out incidence by level of care and management type",
    RATE_HDR,
    rate_by(MG, lambda r: r["_lvl"], _somonths, mfac, keys=LEVELS, sort=False)
    + [["", "", "", "", None]]
    + rate_by(MG, mmgt, _somonths, mfac),
    "Share of managed facility x medicine pairs with at least one stock-out "
    "day across June to August. Incidence, not duration - 4.3 carries the "
    "days.")

add(4, "4.5", "Stock-out incidence by province", RATE_HDR,
    rate_by(MG, lambda r: s(r["_p"][M.PROV]), _somonths, mfac),
    "Worst province first.")

add(4, "4.6", "Stock-out days by province and programme", VOL_HDR,
    vol_by(MG, lambda r: s(r["_p"][M.PROV]),
           lambda r: M.num(r[M.SO_TOTAL]) or 0, mfac)
    + [["", "", "", "", None, None]]
    + vol_by(MG, lambda r: s(r[M.PROGRAM]) or "(none)",
             lambda r: M.num(r[M.SO_TOTAL]) or 0, mfac),
    "Total stock-out days summed across the window. The median is per "
    "facility x medicine pair, so a high total with a zero median means the "
    "burden sits on a few pairs rather than being spread.")

# =============== ANALYSIS 5 =====================================
_pres_un = [r for r in PRES if r["_any_unserved"]]
_all_un = [r for r in PRES if r["_n_items"] and r["_n_unavail"] == r["_n_items"]]
_mixed = [r for r in _pres_un if r not in _all_un]
add(5, "5.1-4.5", "Prescription fulfilment - PROXY on availability",
    ["Metric", "Numerator", "Denominator", "Rate"],
    [["Prescriptions with nothing missing (proxy for complete fulfilment)",
      len(PRES) - len(_pres_un), len(PRES),
      pct(len(PRES) - len(_pres_un), len(PRES))],
     ["Prescriptions with SOME items unavailable (proxy for partially served)",
      len(_mixed), len(PRES), pct(len(_mixed), len(PRES))],
     ["Prescriptions with EVERY item unavailable (proxy for completely "
      "unserved)", len(_all_un), len(PRES), pct(len(_all_un), len(PRES))],
     ["Prescriptions not fully served", len(_pres_un), len(PRES),
      pct(len(_pres_un), len(PRES))],
     ["Medicine lines available (proxy for line fulfilment)",
      len(ITEMS) - len(UNSERVED), len(ITEMS),
      pct(len(ITEMS) - len(UNSERVED), len(ITEMS))]],
    "THE TOOL RECORDS AVAILABILITY, NOT DISPENSING. There is no dispensed "
    "quantity anywhere, so 'fulfilled' here means 'the medicine was available "
    "at the facility', not 'the patient received it'. Metrics 4.6 and 4.7 as "
    "specified cannot be produced at all.", "Proxy only")

add(5, "5.7", "Quantity prescribed on unavailable items",
    ["Measure", "Value"],
    [["Unserved medicine lines", len(UNSERVED)],
     ["Total quantity prescribed on those lines",
      round(sum(P.num(r.get(P.QTY)) or 0 for r in UNSERVED))],
     ["Median quantity per unserved line",
      med([P.num(r.get(P.QTY)) for r in UNSERVED])],
     ["Distinct medicines involved", len({s(r["_med"]) for r in UNSERVED})],
     ["Facilities involved", len({s(r["_sub"][P.FACIL]) for r in UNSERVED})]],
    "This is quantity PRESCRIBED on items recorded unavailable - an upper "
    "bound on the gap. The true unserved quantity (prescribed minus "
    "dispensed) needs a dispensed quantity the form does not collect.",
    "Proxy only")

add(5, "5.8", "Prescription fulfilment by level of care",
    ["Level of care", "Facilities", "Prescriptions",
     "Prescriptions not fully served", "%", "Items", "Unserved items",
     "Unserved %"],
    [[l, len({s(r["_sub"][P.FACIL]) for r in ITEMS if r["_lvl"] == l}),
      len([r for r in PRES if FL.group(r["_sub"][P.FTYPE]) == l]),
      len([r for r in _pres_un if FL.group(r["_sub"][P.FTYPE]) == l]),
      pct(len([r for r in _pres_un if FL.group(r["_sub"][P.FTYPE]) == l]),
          len([r for r in PRES if FL.group(r["_sub"][P.FTYPE]) == l])),
      len([r for r in ITEMS if r["_lvl"] == l]),
      len([r for r in UNSERVED if r["_lvl"] == l]),
      pct(len([r for r in UNSERVED if r["_lvl"] == l]),
          len([r for r in ITEMS if r["_lvl"] == l]))]
     for l in FL.ordered(r["_lvl"] for r in ITEMS)],
    "The unserved rate runs opposite to stock availability - hospitals "
    "prescribe longer, more specialised lists, so they have more opportunity "
    "to prescribe something they do not hold.")

add(5, "5.9", "Most frequently unserved medicines",
    ["Medicine", "Unserved", "Times prescribed", "Unserved %", "Facilities",
     "Quantity"],
    sorted([[m, sum(1 for r in UNSERVED if s(r["_med"]) == m),
             sum(1 for r in ITEMS if s(r["_med"]) == m),
             pct(sum(1 for r in UNSERVED if s(r["_med"]) == m),
                 sum(1 for r in ITEMS if s(r["_med"]) == m)),
             len({s(r["_sub"][P.FACIL]) for r in UNSERVED
                  if s(r["_med"]) == m}),
             round(sum(P.num(r.get(P.QTY)) or 0 for r in UNSERVED
                       if s(r["_med"]) == m))]
            for m in sorted({s(r["_med"]) for r in UNSERVED})],
           key=lambda x: (-x[1], -(x[3] or 0)))[:15])

# =============== ANALYSIS 5 =====================================
def upstream(r):
    reg, cen = s(r.get(P.RMS_REG)).lower(), s(r.get(P.RMS_CEN)).lower()
    if "yes" in (reg, cen):
        return "Distribution constraint (stock existed upstream)"
    if reg == "no" and cen == "no":
        return "System-wide supply constraint (stock nowhere)"
    return "Insufficient evidence (don't know / not answered)"


_attr = Counter(upstream(r) for r in UNSERVED)
def pfac(r):
    return s(r["_sub"][P.FACIL])


def pprov(r):
    return s(r["_sub"][P.PROV])


def _unserved(r):
    return r["_avail"] is False


add(5, "5.10", "Prescription item fulfilment by province", RATE_HDR,
    rate_by([r for r in ITEMS if r["_avail"] is not None], pprov,
            lambda r: r["_avail"] is True, pfac),
    "Share of prescribed items recorded as available at the facility. This "
    "instrument's availability figure is far higher than the stock "
    "assessment's and is a design artefact of who gets recorded - read the "
    "ranking between provinces, not the level.")

add(5, "5.11", "Prescription item fulfilment by district - ten weakest",
    RATE_HDR,
    rate_by([r for r in ITEMS if r["_avail"] is not None],
            lambda r: s(r["_sub"][P.DISTR]), lambda r: r["_avail"] is True,
            pfac)[:10],
    "Same measure, cut by district, weakest first.")

add(5, "5.12", "Unserved items by drug class - the ten most affected",
    RATE_HDR,
    rate_by([r for r in ITEMS if r["_avail"] is not None],
            lambda r: r["_class"], _unserved, pfac, sort=False)
    and sorted(rate_by([r for r in ITEMS if r["_avail"] is not None],
                       lambda r: r["_class"], _unserved, pfac, sort=False),
               key=lambda x: -(x[4] or 0))[:10],
    "The EML class as coded on the form, including the respiratory "
    "correction. Rate here is the share of items in that class that could "
    "NOT be served, so high is bad.")

# ---------------------------------------------------------------
# 5.13 Feedback and comments (user request 2026-10-20)
# ---------------------------------------------------------------
# The form's free-text "feedback, concerns, or recommendations" question,
# asked per prescription and once per visit. Both are read; a facility counts
# once per theme however many answers it gave. Literal substring rules over
# the lower-cased answer, typos as typed; an answer can carry several themes.
# The summaries are written from the verbatim answers at 2026-10-20.
FB_THEMES = [
    ("Prescribing practice: brand names and abbreviations",
     "Brands (Brufen, Flagyl, Priton, Vermox, Coartem, Depakine) and "
     "abbreviations (Amoxy, Cipro, PCT) written instead of INN; collectors "
     "reminded prescribers of prescription principles.",
     ("brand", "abbrevia", "abreviat", "priton", "brufen", "flagyl", "fragyl",
      "lasix", "cromosol", "cromsol", "depakine", "haldol", "vermox",
      "coartem", "depo is", "microgynon", "buscopan", "denk", "antiramine",
      "(cloxa)", "chlorph.")),
    ("RSSB reimbursement and medicine prices",
     "Some medicines (beclomethasone spray, diclofenac suppository) are not "
     "reimbursed, so facilities avoid dispensing them; RSSB and RMS asked to "
     "harmonise prices and invoice verification.",
     ("reimburs", "price", "not paid", "not being paid", "pay all", "be paid",
      "invoices", "tariff")),
    ("EML list and assessment tool gaps",
     "Strengths on the tool differ from those stocked (thiamine 50/100 mg, "
     "cimetidine 400 mg); items managed but not on the EML (ferrous sulfate "
     "200 mg, valproic acid syrup); drug classes to correct; specialists ask "
     "for additions.",
     ("eml", "essential medicines list", "on the list", "add to the list",
      "drug class", "class drug", "class of", "konotoolbox", "on the tool",
      "diarrhoea", "strength", "is 10mg not", "verariate")),
    ("System downtime: prescriptions read from paper registers",
     "eBuzima down, network or power cuts on the visit day, so prescriptions "
     "were taken from OPD registers or patient files.",
     ("network", "system down", "system is down", "shut down", "shutdown",
      "was down", "cut off", "cutt off", "not online", "power", "register",
      "medsoft because files")),
    ("Digitalise records and fully develop eBuzima",
     "Health posts still keep paper records; parallel paper and system "
     "entry; MoH asked to add an offline module and link KWIVUZA / RSSB "
     "billing to eBuzima.",
     ("automat", "autamat", "autpmat", "digitali", "electronic system",
      "paperless", "paper less", "fully develop", "kwivuza", "manual records",
      "paper based", "parallel", "double", "handwritten",
      "health care invoice", "paper prescription", "recorded on paper",
      "elmis")),
    ("System records disagree with stock or the prescription",
     "Items shown out of stock in the system but dispensed; wrong quantities "
     "or strengths in eBuzima; stock reconciliation recommended.",
     ("reconcil", "appears out of stock", "appearing not",
      "appeared out of stock", "quantity reported", "quantity indicated",
      "instead to be 1.5", "it strengths", "zipline",
      "descriptions are different", "update the medicine", "don't display",
      "in the system", "openmrs there", "in ebuzima")),
    ("Timely supply of medicines",
     "Recommendation that RMS and partners make medicines available on time "
     "and in the right quantity.",
     ("avail medicine", "avail this medicine", "avail health",
      "avail medical", "avail timely", "timely", "availing",
      "availability of commodities", "encourage availability",
      "request availability", "on time")),
    ("Prescribing to stock hides unmet need",
     "Prescribers write what the facility holds, so a medicine the patient "
     "had to buy outside cannot be seen - full service may overstate "
     "availability.",
     ("outside", "prescribe existing", "prescribe what they have",
      "prescribes according to the medicine", "are what are available",
      "no room to verify", "necessarily", "bias")),
    ("Stock-outs described",
     "Specific medicines out on the day (insulin, salbutamol spray, "
     "lidocaine, thiamine), expired stock awaiting RMS delivery, part "
     "quantities dispensed.",
     ("stock out due", "out of stock at", "not available:",
      "not available at", "but not available", "programmatic are not",
      "due to stock out", "not managed", "expired")),
    ("Dosing and rational use",
     "Paracetamol quantities that risk toxicity, age-appropriate forms for "
     "children, tapering doses, polypharmacy, under-dosing imposed for "
     "dental patients.",
     ("polypharmacy", "toxicity", "tapering", "regressive", "underdose",
      "treatment duration", "adapted to the age", "should be 6",
      "how patient")),
    ("Enforce system use and refresher training",
     "Staff have the system but do not use it fully; refresher training "
     "requested.",
     ("usability", "enforc", "refresher")),
    ("Referral and patient transfer",
     "Ambulance delays at border health posts; nurses taken off duty to "
     "escort transfers.",
     ("ambulance", "transfer")),
]
FB_NONE = ("No issue raised",
           "Answers such as 'ok', 'good', 'medicines available', 'fully "
           "served'.")
# not feedback: a pasted IT notice, a person's name, a bare number
FB_JUNK = ("dear management team", "mukandoli")


def fb_themes(text):
    """Theme indices for one answer; None = not feedback, [] = no issue."""
    t = " ".join(str(text).split()).lower()
    if any(j in t for j in FB_JUNK) or t.strip().isdigit():
        return None
    t = t.replace("no stock out", "")
    return [i for i, (_, _, ks) in enumerate(FB_THEMES)
            if any(k in t for k in ks)]


# one record per answer, the visit-level ones wrapped so pfac() reads them
FB_ROWS = [{"_sub": r["_sub"], "_th": th}
           for r in PRES if not P.blank(r.get(P.FEEDBACK))
           for th in [fb_themes(r[P.FEEDBACK])] if th is not None] + [
          {"_sub": sb, "_th": th}
          for sb in P.sub if not P.blank(sb.get(P.FEEDBACK))
          for th in [fb_themes(sb[P.FEEDBACK])] if th is not None]


def fb_table(rows):
    fac = defaultdict(set)
    for r in rows:
        for i in r["_th"]:
            fac[i].add(pfac(r))
    themed = set().union(*fac.values()) if fac else set()
    allf = {pfac(r) for r in rows}
    out = [["Facilities that gave feedback", len(allf), ""]]
    out += sorted([[nm, len(fac[i]), sm]
                   for i, (nm, sm, _) in enumerate(FB_THEMES)
                   if fac[i] and nm not in FB_HIDE],
                  key=lambda x: -x[1])
    # "No issue raised" is off the table too (user request 2026-10-26)
    return out


# rows taken off table 5.13 on request (2026-10-26); the answers still count
# towards "Facilities that gave feedback"
FB_HIDE = {"Prescribing practice: brand names and abbreviations",
           "EML list and assessment tool gaps"}


add(5, "5.13", "Feedback and comments from data collectors",
    ["Theme", "Facilities", "What was reported"], fb_table(FB_ROWS),
    "Free-text answers grouped by theme; a facility counts once per theme.")

add(6, "6.1-5.4", "Supply chain attribution of unserved medicines",
    ["Attribution", "Unserved lines", "% of unserved", "Quantity",
     "Value (RWF)"],
    [[k, n, pct(n, len(UNSERVED)),
      round(sum(P.num(r.get(P.QTY)) or 0 for r in UNSERVED
                if upstream(r) == k)),
      round(sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in UNSERVED
                if upstream(r) == k))]
     for k, n in _attr.most_common()],
    "NOT VERIFIED WAREHOUSE RECORDS. There is no RMS dataset in this study - "
    "upstream availability is a facility recall answer (Yes/No/Don't know) on "
    "each unserved item. 'Sufficient' and 'usable' stock cannot be assessed. "
    "The prompt's six categories collapse to these three plus insufficient "
    "evidence: facility inventory-management and central-to-regional "
    "replenishment are not separable with these variables.", "Proxy only")

add(6, "6.5", "Reported upstream availability, Regional against Central",
    ["At RMS Regional", "At RMS Central", "Unserved lines"],
    sorted([[a, b, n] for (a, b), n in Counter(
        (s(r.get(P.RMS_REG)) or "(blank)", s(r.get(P.RMS_CEN)) or "(blank)")
        for r in UNSERVED).items()], key=lambda x: -x[2]),
    "'Don't know' is kept as its own category - folding it into No would "
    "overstate national stock-outs.", "Proxy only")

# =============== ANALYSIS 6 =====================================
_ordered = [r for r in UNSERVED if s(r.get(P.ORDERED)).lower() == "yes"]
_supp = Counter(s(r.get(P.SUPPLIED)) for r in _ordered)
add(7, "7.1", "Requisition rate for missing medicines",
    ["Metric", "Numerator", "Denominator", "Rate"],
    [["Unserved items that had been ordered that month", len(_ordered),
      len(UNSERVED), pct(len(_ordered), len(UNSERVED))],
     ["Unserved items never ordered", len(UNSERVED) - len(_ordered),
      len(UNSERVED), pct(len(UNSERVED) - len(_ordered), len(UNSERVED))]],
    "An unserved medicine never ordered is a facility-side failure; ordered "
    "but not supplied is a supply-side one. The stock tool has no equivalent "
    "yes/no - only monthly quantities.")

add(7, "7.2", "Outcome of the order, where one was placed",
    ["Supply outcome", "Lines", "% of ordered"],
    [[k or "(not recorded)", v, pct(v, len(_ordered))]
     for k, v in _supp.most_common()],
    f"Denominator is the {len(_ordered)} ordered items, not all "
    f"{len(UNSERVED)} unserved - the question is skipped where nothing was "
    f"ordered.")

# Requisition base (2026-10-13): only medicine-months where the quantity
# REQUESTED is greater than or equal to the quantity DELIVERED. A month
# delivered above its request - 216 lines, 186 of them deliveries against a
# zero request - is dropped, so no fill rate can exceed 100% and a delivery
# nobody ordered does not inflate one. Filtered at the month, the finest
# grain the form has, so an over-delivered July does not discard the same
# facility's June and August; every total below is rebuilt from the kept
# months (the form's own totals equal the month sums on every row).
def _mreq(r, mth):
    q = M.num(r[M.REQ[mth]]) or 0
    d = M.num(r[M.DEL[mth]]) or 0
    return q if q >= d else 0


def _mdel(r, mth):
    q = M.num(r[M.REQ[mth]]) or 0
    d = M.num(r[M.DEL[mth]]) or 0
    return d if q >= d else 0


for _r in MG:
    _r["_rq"] = sum(_mreq(_r, m) for m in MONTHS)
    _r["_dv"] = sum(_mdel(_r, m) for m in MONTHS)
_over_lines = sum(1 for r in MG for m in MONTHS
                  if (M.num(r[M.DEL[m]]) or 0) > (M.num(r[M.REQ[m]]) or 0))

_lines = []
for r in MG:
    for mth in MONTHS:
        req = M.num(r[M.REQ[mth]])
        dlv = M.num(r[M.DEL[mth]]) or 0
        if req is not None and req > 0 and dlv <= req:
            _lines.append((req, dlv))
_full = sum(1 for q, d in _lines if d == q)
_part = sum(1 for q, d in _lines if 0 < d < q)
_zero = sum(1 for q, d in _lines if d == 0)
_req = sum(q for q, d in _lines)
_dlv = sum(d for q, d in _lines)
add(7, "7.2-6.5", "Order fill performance, medicine-month lines",
    ["Metric", "Numerator", "Denominator", "Rate"],
    [["Complete order fill rate (delivered = requested)", _full, len(_lines),
      pct(_full, len(_lines))],
     ["Partial fulfilment (some but not all)", _part, len(_lines),
      pct(_part, len(_lines))],
     ["Non-fulfilment (nothing delivered)", _zero, len(_lines),
      pct(_zero, len(_lines))],
     ["QUANTITY FILL RATE", round(_dlv), round(_req), pct(_dlv, _req)]],
    "An 'order line' here is a medicine-month, not a real requisition line - "
    "the tool collects monthly totals, not orders. Lines with nothing "
    "requested are excluded, and so are the "
    f"{_over_lines} months delivered above their request. Metric 7.6, order "
    "lead time, cannot be computed: there are no requisition or delivery "
    "dates.")

add(7, "7.6", "Quantity fill rate by month",
    ["Month", "Requested", "Received", "Shortfall", "FILL RATE"],
    [[mth, round(sum(_mreq(r, mth) for r in MG)),
      round(sum(_mdel(r, mth) for r in MG)),
      round(sum(_mreq(r, mth) for r in MG)
            - sum(_mdel(r, mth) for r in MG)),
      pct(sum(_mdel(r, mth) for r in MG),
          sum(_mreq(r, mth) for r in MG))] for mth in MONTHS],
    "From summed quantities, never a mean of per-row percentages.")

add(7, "7.7", "Quantity fill rate by level of care",
    ["Level of care", "Facilities", "Requested", "Received", "FILL RATE"],
    [[l, len({s(r["_p"][M.FAC]) for r in MG if r["_lvl"] == l}),
      round(sum(r["_rq"] for r in MG if r["_lvl"] == l)),
      round(sum(r["_dv"] for r in MG if r["_lvl"] == l)),
      pct(sum(r["_dv"] for r in MG if r["_lvl"] == l),
          sum(r["_rq"] for r in MG if r["_lvl"] == l))]
     for l in LEVELS if any(r["_lvl"] == l for r in MG)])

add(7, "7.8", "Most under-supplied medicines",
    ["Medicine", "Requested", "Received", "Shortfall", "FILL RATE",
     "Stocked out on visit day"],
    sorted([[m, round(sum(r["_rq"] for r in MG
                          if s(r[M.MED]) == m)),
             round(sum(r["_dv"] for r in MG
                       if s(r[M.MED]) == m)),
             round(sum(r["_rq"] for r in MG if s(r[M.MED]) == m)
                   - sum(r["_dv"] for r in MG
                         if s(r[M.MED]) == m)),
             pct(sum(r["_dv"] for r in MG if s(r[M.MED]) == m),
                 sum(r["_rq"] for r in MG if s(r[M.MED]) == m)),
             sum(1 for r in SO if s(r[M.MED]) == m)]
            for m in sorted({s(r[M.MED]) for r in MG})
            if sum(r["_rq"] for r in MG
                   if s(r[M.MED]) == m) >= 1000],
           key=lambda x: x[4])[:15],
    "Medicines with at least 1,000 units requested, worst fill rate first.")

# =============== ANALYSIS 7 =====================================
_val = sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in UNSERVED)
_priced = [r for r in UNSERVED if not P.blank(r.get(P.UNIT_PRICE))]
FILL_HDR = ["Group", "Facilities", "Observations", "Requested", "Delivered", "FILL RATE"]
add(7, "7.9", "Quantity fill rate by province", FILL_HDR,
    [[k, len({mfac(r) for r in g}), len(g),
      round(sum(r["_rq"] for r in g)),
      round(sum(r["_dv"] for r in g)),
      pct(sum(r["_dv"] for r in g),
          sum(r["_rq"] for r in g))]
     for k, g in sorted(
         ((p, [r for r in MG if s(r["_p"][M.PROV]) == p])
          for p in sorted({s(r["_p"][M.PROV]) for r in MG})),
         key=lambda kv: (pct(sum(r["_dv"] for r in kv[1]),
                             sum(r["_rq"] for r in kv[1]))
                         or 0))],
    "Delivered against requested, June to August, pooled over the province. "
    "Columns four and five are the quantities behind the rate.",
    )

add(7, "7.10", "Quantity fill rate by programme and management type", FILL_HDR,
    [[k, len({mfac(r) for r in g}), len(g),
      round(sum(r["_rq"] for r in g)),
      round(sum(r["_dv"] for r in g)),
      pct(sum(r["_dv"] for r in g),
          sum(r["_rq"] for r in g))]
     for k, g in (
         [(p, [r for r in MG if s(r[M.PROGRAM]) == p])
          for p in sorted({s(r[M.PROGRAM]) for r in MG})]
         + [("", [])]
         + [(t, [r for r in MG if mmgt(r) == t])
            for t in sorted({mmgt(r) for r in MG})])
     if g or k == ""],
    "Same measure, cut by programme and by who runs the facility.")

add(8, "8.1-7.3", "Financial value of unserved prescriptions",
    ["Metric", "Value"],
    [["Unserved medicine lines", len(UNSERVED)],
     ["Lines carrying a tariff price", len(_priced)],
     ["TOTAL UNSERVED VALUE (RWF)", round(_val)],
     ["Median value per unserved line",
      med([P.num(r.get(P.COST_UNSERVED)) for r in _priced])],
     ["Prescriptions not fully served", len(_pres_un)],
     ["AVERAGE VALUE PER AFFECTED PRESCRIPTION (RWF)",
      round(_val / len(_pres_un)) if _pres_un else None]],
    "Observed sample value only. Facilities follow a visit plan, not a "
    "probability sample, so this must NOT be extrapolated nationally. It is "
    "also not patient expenditure - no expenditure data were collected.")

_byval = sorted([[m, sum(1 for r in UNSERVED if s(r["_med"]) == m),
                  round(sum(P.num(r.get(P.COST_UNSERVED)) or 0
                            for r in UNSERVED if s(r["_med"]) == m))]
                 for m in sorted({s(r["_med"]) for r in UNSERVED})],
                key=lambda x: -x[2])
_cum, _pareto = 0, []
for m, n, v in _byval:
    _cum += v
    _pareto.append([m, n, v, pct(v, _val), pct(_cum, _val)])
add(8, "8.4", "Pareto - medicine contribution to the unserved value",
    ["Medicine", "Unserved lines", "Value (RWF)", "% of total",
     "Cumulative %"], _pareto[:15],
    f"{sum(1 for x in _pareto if x[4] and x[4] <= 0.8)} medicines account for "
    f"the first 80% of the value.")

add(8, "8.5", "Unserved value by level of care and province",
    ["Group", "Dimension", "Unserved lines", "Value (RWF)", "% of total"],
    [[l, "Level of care", len([r for r in UNSERVED if r["_lvl"] == l]),
      round(sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in UNSERVED
                if r["_lvl"] == l)),
      pct(sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in UNSERVED
              if r["_lvl"] == l), _val)]
     for l in FL.ordered(r["_lvl"] for r in UNSERVED)]
    + [[p, "Province",
        len([r for r in UNSERVED if s(r["_sub"][P.PROV]) == p]),
        round(sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in UNSERVED
                  if s(r["_sub"][P.PROV]) == p)),
        pct(sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in UNSERVED
                if s(r["_sub"][P.PROV]) == p), _val)]
       for p in sorted({s(r["_sub"][P.PROV]) for r in UNSERVED})],
    "Metric 7.5 as specified also asks for value by supply chain constraint - "
    "that is on indicator 5.1-5.4, which carries a value column.")

# =============== ANALYSIS 8 =====================================
add(8, "8.6", "Unserved value by drug class - the ten largest", VOL_HDR,
    vol_by([r for r in UNSERVED
            if (P.num(r.get(P.COST_UNSERVED)) or 0) > 0],
           lambda r: r["_class"],
           lambda r: P.num(r.get(P.COST_UNSERVED)) or 0, pfac)[:10],
    "The form's own cost-of-unserved field, on unserved items that carry "
    "one. Share is of the priced unserved value, not of all prescriptions.")

# ---------------------------------------------------------------
# Estimated sales lost to stock-outs - stock assessment, priced products
# ---------------------------------------------------------------
# For every managed item that was out at least one day in June-August:
#   daily demand = Jun-Aug consumption / days IN stock (92 - days out)
#   units lost   = daily demand x days out
#   lost sales   = units lost x unit_tariff_price x (1 + MARKUP)
# Demand is taken over in-stock days, not all 92: consumption stops while a
# medicine is out, so dividing by the full window understates demand by
# exactly the share of time it was missing. A pair out for all 92 days has no
# in-stock consumption to estimate demand from, and is excluded and counted.
# unit_tariff_price is the RMS price; facilities sell at that plus a 20%
# markup, so the markup is carried as its own column - it is the margin the
# facility itself lost. Programme products (TB, FP, Malaria, ARV) carry no
# tariff - they are supplied free - so a stock-out there loses no sales and
# they cannot appear here. Days come from the monthly fields (_wdays), as
# in 4.7. This is NOT the form's financial_loss, which is driven by
# undelivered requisition quantities.
MARKUP = 0.20
FL_HDR = ["Group", "Facilities", "Pairs out", "Estimated units lost",
          "Value at tariff price (RWF)", "20% markup (RWF)",
          "LOST SALES incl. markup (RWF)", "SHARE OF TOTAL"]


def _loss_row(r):
    """(units lost, value at tariff) for one managed item, or the reason it
    cannot be estimated; None if it was never out."""
    out = min(_wdays(r), 92)
    if out <= 0:
        return None
    price = M.num(r[M.PRICE]) or 0
    if price <= 0:
        return "no price"
    cons = [M.num(r[M.CONS[m]]) for m in MONTHS]
    if all(c is None for c in cons):
        return "no consumption recorded"
    if out >= 92:
        return "out all 92 days - no demand signal"
    units = sum(c or 0 for c in cons) / (92 - out) * out
    return units, units * price


_loss_all = [(r, _loss_row(r)) for r in MG]
LOSS_ROWS = [(r, v) for r, v in _loss_all if isinstance(v, tuple)]
_loss_excl = Counter(v for _, v in _loss_all if isinstance(v, str))
_loss_tot = sum(v[1] for _, v in LOSS_ROWS) * (1 + MARKUP)
_loss_top = max(LOSS_ROWS, key=lambda x: x[1][1])


def loss_by(keyfn, keys=None):
    g = defaultdict(list)
    for r, v in LOSS_ROWS:
        g[keyfn(r)].append((r, v))
    out = []
    for k in (keys or sorted(g, key=str)):
        rs = g.get(k)
        if not rs:
            continue
        base = sum(v[1] for _, v in rs)
        out.append([k, len({s(r["_p"][M.FAC]) for r, _ in rs}), len(rs),
                    round(sum(v[0] for _, v in rs)), round(base),
                    round(base * MARKUP), round(base * (1 + MARKUP)),
                    pct(base * (1 + MARKUP), _loss_tot)])
    return out


def _by_value(rows):
    return sorted(rows, key=lambda x: -x[6])


_n_out = sum(1 for r in MG if _wdays(r) > 0)
add(8, "8.7", "Estimated sales lost to stock-outs - overall",
    ["Measure", "Value"],
    [["Managed pairs with at least one stock-out day, Jun-Aug", _n_out],
     ["  estimated (priced, with in-stock consumption)", len(LOSS_ROWS)],
     ["  excluded - no price (programme products supplied free; insulin, "
      "carbetocin)", _loss_excl["no price"]],
     ["  excluded - out all 92 days, no demand signal",
      _loss_excl["out all 92 days - no demand signal"]],
     ["  excluded - no consumption recorded",
      _loss_excl["no consumption recorded"]],
     ["Estimated units lost", round(sum(v[0] for _, v in LOSS_ROWS))],
     ["Value at tariff price (RWF)", round(_loss_tot / (1 + MARKUP))],
     ["20% markup - the facility's own margin lost (RWF)",
      round(_loss_tot / (1 + MARKUP) * MARKUP)],
     ["LOST SALES incl. 20% markup (RWF)", round(_loss_tot)],
     ["Facilities with an estimated loss",
      len({s(r["_p"][M.FAC]) for r, _ in LOSS_ROWS})],
     ["Largest single facility x medicine - "
      + (lambda r: f"{s(r[M.MED])} at {s(r['_p'][M.FAC])}, out "
         f"{min(_wdays(r), 92):.0f} days")(_loss_top[0]),
      round(_loss_top[1][1] * (1 + MARKUP))],
     ["  its share of all lost sales",
      pct(_loss_top[1][1] * (1 + MARKUP), _loss_tot)]],
    "Daily demand = June-August consumption / days in stock; units lost = "
    "daily demand x days out; lost sales = units lost x unit tariff price x "
    "1.20. Days out are the three monthly fields summed. An estimate of the "
    "sales the facility could not make over three months - not a measured "
    "loss, and not the form's financial_loss field, which is driven by "
    "undelivered requisitions.")
add(8, "8.8", "Estimated lost sales by level of care", FL_HDR,
    loss_by(lambda r: r["_lvl"], keys=LEVELS),
    "Each level's lost sales and its share of the total.")
add(8, "8.9", "Estimated lost sales by drug class", FL_HDR,
    _by_value(loss_by(_class_label)),
    "Largest first. MCCH class labels corrected for display, as in FIG 2.7.")
add(8, "8.10", "Estimated lost sales by medicine", FL_HDR,
    _by_value(loss_by(lambda r: s(r[M.MED]))),
    "Every priced medicine with an estimated loss, largest first.")
add(8, "8.11", "Estimated lost sales by province", FL_HDR,
    _by_value(loss_by(lambda r: s(r["_p"][M.PROV]))),
    "Largest first.")
add(8, "8.12", "Estimated lost sales by district", FL_HDR,
    _by_value(loss_by(lambda r: s(r["_p"][M.DIST]))),
    "Largest first.")

add(9, "9.1", "Overall CHW medicine availability",
    ["Measure", "Value"],
    [["CHW x medicine observations", C.N_MED],
     ["Managed by the CHW (denominator)", len(CM)],
     ["Available on the visit day (numerator)", len(CAV)],
     ["CHW AVAILABILITY", pct(len(CAV), len(CANS))],
     ["Stocked out", len(CANS) - len(CAV)],
     ["CHWs assessed", len(C.CHWS)],
     ["Parent health centres", len(C.FACILITIES)],
     ["Not managed - MFT zoning and similar", C.N_MED - len(CM)]],
    "Denominator is rows the CHW manages. About 40% of rows are 'not managed' "
    "because RBC's MFT zoning assigns each area one first-line antimalarial - "
    "counting those as unavailable would report the zoning policy as a "
    "stock-out.")

add(9, "9.2", "Individual CHW medicine availability",
    ["Medicine", "Managed", "Available", "AVAILABLE %", "CHWs"],
    sorted([[m, sum(1 for r in CM if r["_short"] == m),
             sum(1 for r in CAV if r["_short"] == m),
             pct(sum(1 for r in CAV if r["_short"] == m),
                 sum(1 for r in CM if r["_short"] == m)),
             len({(C.s(r["_p"][C.FAC]), C.s(r["_p"][C.CHW])) for r in CM
                  if r["_short"] == m})]
            for m in sorted({r["_short"] for r in CM})], key=lambda x: x[3] or 0))

_chw_by = defaultdict(list)
for r in CM:
    _chw_by[(C.s(r["_p"][C.FAC]), C.s(r["_p"][C.CHW]))].append(r)
_chw_complete = sum(1 for k, rs in _chw_by.items()
                    if all(C.yes(r[C.AVAIL]) for r in rs))
add(9, "9.3", "Complete CHW basket availability",
    ["Measure", "Value"],
    [["CHWs with a complete assessment", len(_chw_by)],
     ["CHWs holding EVERY medicine they manage", _chw_complete],
     ["COMPLETE BASKET AVAILABILITY", pct(_chw_complete, len(_chw_by))],
     ["CHWs missing exactly one",
      sum(1 for k, rs in _chw_by.items()
          if sum(1 for r in rs if C.no(r[C.AVAIL])) == 1)],
     ["CHWs missing three or more",
      sum(1 for k, rs in _chw_by.items()
          if sum(1 for r in rs if C.no(r[C.AVAIL])) >= 3)]],
    "Over the medicines each CHW actually manages, not the full 13-item "
    "package.")

SHARED = {
    "ALu 1x6": "ARTE 20MG+LUME 120MG (1X6)",
    "ALu 4x6": "ARTE 20MG + LUME 120MG (4X6)",
    "Malaria RDT": "MALARIA RAPID DIAGNOSTICS Ag P.f/Pan 30T",
    "DHAP 20/160": "Dihydroartemisinin/Piperaquine (DHAP) 20/160 mg Disp Tablet",
    "DHAP 30/240": "Dihydroartemisinin/Piperaquine DHAP) 30/240 mg Dispersible Tablet",
    "DHAP 80/640": "Dihydroartemisinin/Piperaquine DHAP) 80/640 mg Dispersible Tablet",
    "Pyr/Art 180/60": "Pyronaridine/Artesunate 180/60 mg Film-Coated Tablet",
    "Pyr/Art 60/20": "Pyronaridine/Artesunate 60/20 mg Granules for Suspension",
}
_cmp = []
for short, facname in SHARED.items():
    c_m = [r for r in CM if r["_short"] == short]
    f_m = [r for r in ANS if s(r[M.MED]) == facname]
    _cmp.append([short, len(c_m), pct(sum(1 for r in c_m if C.yes(r[C.AVAIL])),
                                      len(c_m)),
                 len(f_m), pct(sum(1 for r in f_m if M.yes(r[M.AVAIL])),
                               len(f_m)),
                 (pct(sum(1 for r in c_m if C.yes(r[C.AVAIL])), len(c_m)) or 0)
                 - (pct(sum(1 for r in f_m if M.yes(r[M.AVAIL])), len(f_m)) or 0)])
_ctot = [r for r in CM if r["_short"] in SHARED]
_ftot = [r for r in ANS if s(r[M.MED]) in SHARED.values()]
_cmp.append(["ALL EIGHT SHARED PRODUCTS", len(_ctot),
             pct(sum(1 for r in _ctot if C.yes(r[C.AVAIL])), len(_ctot)),
             len(_ftot),
             pct(sum(1 for r in _ftot if M.yes(r[M.AVAIL])), len(_ftot)),
             (pct(sum(1 for r in _ctot if C.yes(r[C.AVAIL])), len(_ctot)) or 0)
             - (pct(sum(1 for r in _ftot if M.yes(r[M.AVAIL])), len(_ftot)) or 0)])
add(9, "9.4", "Community against facility, on the products both carry",
    ["Product", "CHW observations", "CHW available %",
     "Facility observations", "Facility available %",
     "Difference (percentage points)"], _cmp,
    "Only eight of the 13 CHW items have a facility comparator: ORS, Zinc "
    "10mg and Amox 125mg dispersible are carried by CHWs and by no facility "
    "in this study, and ALu 2x6 and 3x6 are CHW-only presentations.")

add(9, "9.5", "CHW availability by district", AVAIL_HDR[:1] + ["CHWs",
    "Managed", "Available", "AVAILABLE %"],
    sorted([[d, len({(C.s(r["_p"][C.FAC]), C.s(r["_p"][C.CHW])) for r in CM
                     if C.s(r["_p"][C.DIST]) == d}),
             sum(1 for r in CM if C.s(r["_p"][C.DIST]) == d),
             sum(1 for r in CAV if C.s(r["_p"][C.DIST]) == d),
             pct(sum(1 for r in CAV if C.s(r["_p"][C.DIST]) == d),
                 sum(1 for r in CM if C.s(r["_p"][C.DIST]) == d))]
            for d in sorted({C.s(r["_p"][C.DIST]) for r in CM})],
           key=lambda x: x[4] or 0))

ANALYSES = {
    1: "Tracer medicine availability",
    2: "Stock adequacy and Inventory accuracy",
    3: "Medicine consumption",
    4: "Stock-out frequency and duration",
    5: "Prescription fulfilment",
    6: "Supply chain attribution",
    7: "Requisition and order fulfilment",
    8: "Financial value of unserved prescriptions",
    9: "Community-level medicine availability",
}
GAPS = [
    ("5.6 / 5.7", "Quantity fulfilment and unserved quantity",
     "No dispensed quantity is collected anywhere."),
    ("5.x / 8.x", "Inpatient versus outpatient",
     "No service, ward or admission variable at any level."),
    ("4.2 / 4.4", "Stock-out episodes and average episode duration",
     "Days are collected, episodes are not. A proxy is reported as 4.2."),
    ("7.6", "Order lead time",
     "No requisition or delivery dates - monthly quantities only."),
    ("6.x", "Verified warehouse records",
     "No RMS dataset. Upstream availability is a facility recall answer."),
    ("1.1", "Unexpired and usable stock",
     "No expiry or usability variable; availability means physically present."),
]


# ---------------------------------------------------------------
# renderer 1: Excel
# ---------------------------------------------------------------
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, size=9, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=12, color="1F3864")
PROXY_FILL = PatternFill("solid", fgColor="FFF4D6")


def is_pct(h, v):
    return isinstance(v, float) and 0 <= v <= 1 and (
        "%" in str(h).upper() or "RATE" in str(h).upper()
        or str(h).strip().lower() in ("value", "rate"))


def write_xlsx():
    wb = Workbook()
    wb.remove(wb.active)

    # The calculation sheet goes FIRST, so the workbook opens on how the
    # figures were derived rather than on a number with no provenance.
    ws = wb.create_sheet("How figures are calculated")
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1,
            value="HOW THESE FIGURES ARE CALCULATED").font = Font(
        bold=True, size=14, color="1F3864")
    c = ws.cell(row=2, column=1,
                value="Values are read from the same row sets the indicators "
                      "use, so they cannot drift from the figures they "
                      "explain. Only the definitions and the reasons are "
                      "written by hand.")
    c.font = Font(size=9, italic=True, color="555555")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
    r = 4
    block = None
    for b, item, how, val in method_rows():
        if b != block:
            ws.cell(row=r, column=1, value=b).font = Font(bold=True, size=11)
            r += 1
            for ci, h in enumerate(["Measure", "How it is calculated",
                                    "Value"], start=1):
                cc = ws.cell(row=r, column=ci, value=h)
                cc.font, cc.fill, cc.border = HDR_FONT, HDR_FILL, BORDER
                cc.alignment = Alignment(horizontal="center",
                                         wrap_text=True)
            r += 1
            block = b
        for ci, v in enumerate([item, how, val], start=1):
            cc = ws.cell(row=r, column=ci, value=v)
            cc.border = BORDER
            cc.alignment = Alignment(wrap_text=ci == 2, vertical="top")
            if ci == 3:
                cc.font = Font(bold=True)
        r += 1
    for ci, w in {1: 38, 2: 96, 3: 14}.items():
        ws.column_dimensions[get_column_letter(ci)].width = w

    for a, name in ANALYSES.items():
        ws = wb.create_sheet(f"{a}. {name}"[:31])
        ws.sheet_view.showGridLines = False
        r = 1
        ws.cell(row=r, column=1,
                value=f"ANALYSIS {a} - {name}").font = Font(
            bold=True, size=14, color="1F3864")
        r += 2
        widths = {}
        for an, mid, title, status, note, headers, rows in TABLES:
            if an != a:
                continue
            c = ws.cell(row=r, column=1, value=f"{mid}  {title}")
            c.font = TITLE_FONT
            if status != "Feasible":
                ws.cell(row=r, column=2, value=f"[{status}]").font = Font(
                    bold=True, size=10, color="9C6500")
            r += 1
            if note:
                c = ws.cell(row=r, column=1, value=note)
                c.font = Font(size=9, italic=True, color="555555")
                c.alignment = Alignment(wrap_text=True, vertical="top")
                ws.merge_cells(start_row=r, start_column=1, end_row=r,
                               end_column=max(len(headers), 4))
                ws.row_dimensions[r].height = 13 * (1 + len(note) // 120)
                r += 1
            for ci, h in enumerate(headers, start=1):
                cc = ws.cell(row=r, column=ci, value=h)
                cc.font, cc.fill, cc.border = HDR_FONT, HDR_FILL, BORDER
                cc.alignment = Alignment(wrap_text=True, vertical="center")
                widths[ci] = max(widths.get(ci, 10), min(len(str(h)) + 2, 30))
            r += 1
            for row in rows:
                for ci, v in enumerate(row, start=1):
                    cc = ws.cell(row=r, column=ci, value=v)
                    cc.border = BORDER
                    cc.alignment = Alignment(vertical="top",
                                             wrap_text=isinstance(v, str)
                                             and len(str(v)) > 34)
                    if isinstance(v, float) and 0 <= v <= 1 \
                            and headers[ci - 1] not in NOT_PCT:
                        cc.number_format = "0.0%"
                    elif isinstance(v, (int, float)):
                        cc.number_format = "#,##0.##"
                    if status != "Feasible" and ci == 1:
                        cc.fill = PROXY_FILL
                    widths[ci] = max(widths.get(ci, 10),
                                     min(len(str(v)) + 2, 46))
                r += 1
            r += 2
        for ci, w in widths.items():
            ws.column_dimensions[get_column_letter(ci)].width = w

    ws = wb.create_sheet("Method and gaps")
    ws.sheet_view.showGridLines = False
    rows = [["Reference period", PERIOD],
            ["Availability denominator",
             "rows the facility manages AND answered - never all observations"],
            ["Accuracy denominator",
             "available rows only; counts are not recorded where absent"],
            ["Supply outcome denominator", "items that were ordered"],
            ["Cleaning applied",
             "two servers merged; 12 duplicate facility x medicine rows "
             "removed; 7 facility reattributions; levels grouped; non-finite "
             "calculate fields rejected; window restricted to Jun-Aug"],
            ["Sampling", "facilities follow a visit plan, not a probability "
             "sample - report as a census of visited facilities, do not "
             "extrapolate nationally"], ["", ""]]
    rows += [["GAP - " + g[0], f"{g[1]}: {g[2]}"] for g in GAPS]
    for i, (k, v) in enumerate(rows, start=1):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True, size=10)
        c = ws.cell(row=i, column=2, value=v)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 120
    wb.active = 0
    if os.path.exists(OUT_XLSX):
        os.chmod(OUT_XLSX, 0o644)
    wb.save(OUT_XLSX)


# ---------------------------------------------------------------
# renderer 2: the all-indicator dashboard
# ---------------------------------------------------------------
def esc(v):
    return html.escape("" if v is None else str(v), quote=True)


# Columns of quantities that can legitimately fall between 0 and 1. Every
# other float in that range is a rate and prints as a percentage.
NOT_PCT = {"Total AMC", "Mean AMC", "Median AMC", "SD of AMC", "CV",
           "Average days out",
           "Median months"}


def fmt(h, v):
    if v is None or v == "":
        return ""
    if h in NOT_PCT and isinstance(v, (int, float)):
        return f"{v:,.1f}"
    if isinstance(v, float) and 0 <= v <= 1:
        return f"{v * 100:.1f}%"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        # one decimal everywhere a value is not a whole count (2026-10-26)
        return f"{v:,.0f}" if float(v).is_integer() else f"{v:,.1f}"
    return str(v)


CSS = """
<style>
:root{color-scheme:dark;
 --plane:#000;--surface:#1a1a19;--raise:#222221;--ink:#fff;--ink2:#c3c2b7;
 --muted:#898781;--grid:#2c2c2a;--axis:#383835;--rule:rgba(255,255,255,.14);
 --accent:#d95926;--good:#199e70;--warn:#eda100;--bad:#e66767;
 --font:"Trebuchet MS","Lucida Grande","Lucida Sans Unicode",Tahoma,sans-serif;}
*{box-sizing:border-box}
body{margin:0;background:var(--plane);color:var(--ink);font-family:var(--font);
 font-size:16px;line-height:1.55}
.wrap{max-width:1240px;margin:0 auto;padding:0 20px 72px}
h1,h2,h3,h4{font-family:var(--font);margin:0;text-wrap:balance}
.mast{border-bottom:1px solid var(--rule);padding:44px 0 26px}
.mast h1{font-size:clamp(28px,4.6vw,44px);letter-spacing:-.015em}
.mast p{color:var(--ink2);max-width:80ch;margin:14px 0 0;font-size:15px}
.cover{display:grid;grid-template-columns:2fr 1fr 1fr;gap:1px;
 background:var(--rule);border:1px solid var(--rule);border-top:0}
.cv1{background:var(--surface);padding:12px 16px}
.mast .cv1 .cvv{margin:0;font-size:24px;font-weight:800;color:#fff;
 line-height:1.15}
.mast .cv1 .cvl{margin:2px 0 0;font-size:13.5px;color:#fff}
.mast .cv1 .cvs{margin:4px 0 0;font-size:12.5px;color:#c3c2b7;line-height:1.5}
.mast .cv1 .cvs b{color:#fff}
@media (max-width:760px){.cover{grid-template-columns:1fr}}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));
 gap:1px;background:var(--rule);border:1px solid var(--rule);margin:26px 0 0}
.stat{background:var(--surface);padding:16px}
.stat .v{margin:0;font-size:27px;font-weight:600}
/* header counts: white and larger (2026-10-18) */
.mast .stat .v{color:#fff;font-size:32px;font-weight:800;line-height:1.15;
 letter-spacing:.01em}
.stat .l{margin:3px 0 0;font-size:13.5px;color:#fff}
.stat .s{margin:2px 0 0;font-size:12px;color:#fff}
.tabs{display:flex;flex-wrap:wrap;gap:2px;margin:30px 0 0;
 border-bottom:1px solid var(--rule)}
.tab{font:inherit;font-size:14px;color:#fff;font-weight:700;
 text-transform:uppercase;
 background:none;border:0;
 border-bottom:2px solid transparent;padding:11px 14px;margin-bottom:-1px;
 cursor:pointer;white-space:nowrap}
.tab:hover{color:#fff}
.tab.on{color:#fff;font-weight:700;border-bottom-color:var(--accent)}
.js .page{display:none}.js .page.on{display:block}
.page{margin-top:34px}
.page>h2{font-size:24px;padding-bottom:8px;border-bottom:2px solid var(--accent)}
.ind{margin-top:30px}
.ind h4{font-size:15.5px;font-weight:600;display:flex;gap:10px;
 align-items:baseline;flex-wrap:wrap}
.mid{font-family:var(--font);font-size:12px;color:var(--accent);font-weight:700;
 letter-spacing:.04em}
.badge{font-size:11px;font-weight:600;padding:2px 8px;border-radius:99px;
 background:#3a2e12;color:var(--warn);text-transform:uppercase;
 letter-spacing:.05em}
.note{margin:7px 0 10px;font-size:13px;color:var(--muted);max-width:92ch;
 line-height:1.6}
.calc{margin:6px 0 4px;font-size:12.5px;color:var(--ink2);max-width:92ch;
 line-height:1.5;border-left:3px solid var(--accent);padding:3px 0 3px 10px}
.calc b{color:var(--ink);font-weight:700}
.scroll{overflow-x:auto;border:1px solid var(--rule);background:var(--surface)}
.scroll.tall{max-height:460px;overflow-y:auto}
table{border-collapse:collapse;width:100%;font-size:13.5px}
thead th{position:sticky;top:0;background:var(--raise);text-align:left;
 font-weight:600;font-size:12.5px;color:var(--ink2);padding:8px 11px;
 border-bottom:1px solid var(--axis);white-space:nowrap;z-index:1}
td{padding:6px 11px;border-bottom:1px solid var(--grid);vertical-align:top;
 color:var(--ink2)}
td.num{text-align:right;font-variant-numeric:tabular-nums;color:var(--ink);
 white-space:nowrap}
tbody tr:last-child td{border-bottom:0}
tbody tr:hover td{background:#141413}
.bar{display:inline-block;height:9px;border-radius:99px;background:var(--good);
 vertical-align:middle;margin-left:8px}
.bar.mid{background:var(--warn)}.bar.low{background:var(--bad)}
.layers{margin:10px 0 0;padding:12px 14px;border:1px solid var(--rule);
 border-radius:8px;background:#141413}
.lyrh{margin:0 0 8px;font-size:11px;letter-spacing:.08em;text-transform:uppercase;
 color:var(--accent);font-weight:700}
.lyr{display:flex;align-items:center;gap:8px;font-size:12.5px;padding:3px 0;
 cursor:pointer}
.lyr input{accent-color:var(--accent);cursor:pointer}
.lyrn{margin-left:auto;color:var(--muted);font-variant-numeric:tabular-nums}
.lyrbtn{margin:10px 0 0;display:flex;gap:6px}
.lyrbtn button{flex:1;font:inherit;font-size:11.5px;padding:5px 8px;cursor:pointer;
 color:var(--ink);background:#1f1f1e;border:1px solid var(--rule);border-radius:5px}
.lyrbtn button:hover{border-color:var(--accent);color:var(--accent)}
.lyrsel{width:100%;font:inherit;font-size:12.5px;padding:6px 8px;color:var(--ink);
 background:#1f1f1e;border:1px solid var(--rule);border-radius:5px;cursor:pointer}
.lyrsel:disabled{opacity:.4;cursor:not-allowed}
{LAYER_CSS}
.gap{margin:26px 0 0;border:1px solid var(--rule);border-left:3px solid var(--bad);
 background:var(--surface);padding:16px 18px}
.gap h3{font-size:16px;margin:0 0 8px}
.gap li{font-size:13.5px;color:var(--ink2);margin:0 0 6px}
.maprow{display:grid;grid-template-columns:minmax(0,1fr) 290px;gap:18px;
 align-items:start;margin-top:14px}
.mapbox{border:1px solid var(--rule);background:var(--surface);padding:6px}
.mapbox svg{display:block;width:100%;height:auto}
.pair2{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:18px;
 align-items:start;margin-top:10px}
.pair2 .ptitle{margin:0 0 6px;font-size:12px;letter-spacing:.06em;
 text-transform:uppercase;color:var(--muted)}
.legend-row{display:flex;flex-wrap:wrap;gap:6px 18px;align-items:center;
 margin-top:12px}
.legend-row h4{margin:0 8px 0 0}
.legend-row .key{margin:0}
@media (max-width:900px){.pair2{grid-template-columns:1fr}}
.dist{stroke:#000;stroke-width:1}
.dist:hover{stroke:#fff;stroke-width:2}
.border{stroke-linejoin:round}
.border:hover{stroke:#fff;stroke-width:1.8}
.mtitle{fill:#fff;font-size:15px;font-weight:700}
.msub{fill:#c3c2b7;font-size:11px}
.msrc{fill:#898781;font-size:9.5px}
.mlab{fill:#c3c2b7;font-size:10px}
.slegend rect,.sbar rect{shape-rendering:crispEdges}
.dlabel{fill:#fff;font-size:9.5px;text-anchor:middle;pointer-events:none;
 paint-order:stroke;stroke:#000;stroke-width:2.4px;stroke-linejoin:round}
.pt{stroke:#000;stroke-width:.8}
.pt:hover{stroke:#fff;stroke-width:1.8}
.legend{border:1px solid var(--rule);background:var(--surface);padding:15px}
.legend h4{font-size:13.5px;margin:0 0 9px}
.pcards{margin:4px 0 22px}
.szrow{display:flex;align-items:flex-end;gap:16px}
.szk{display:inline-flex;align-items:center;gap:6px}
.mapboxgl-popup-content{background:#fff!important;color:#111!important;
 padding:9px 12px!important;border-radius:6px!important;
 box-shadow:0 4px 14px rgba(0,0,0,.45)!important;font-family:var(--font)}
.mapboxgl-popup-tip{border-top-color:#fff!important;border-bottom-color:#fff!important}
.fpop{font-size:13px;line-height:1.45;color:#111}
.fpop b{color:#111}
.fpop span{display:block;color:#333}
.fac-style button{width:auto!important;padding:0 10px!important;
 font:12px var(--font)!important;color:#333}
.fac-style button.on{background:#d95926!important;color:#fff}
.vmap .mp-p{display:none}
.vmap:has(.v-p:checked) .mp-p{display:block}
.vmap:has(.v-p:checked) .mp-d{display:none}
.vmap:has(.v-p:checked) .lyrsel{opacity:.4;pointer-events:none}
.key{display:flex;align-items:center;gap:9px;font-size:12.5px;
 color:var(--ink2);margin:0 0 6px}
.sw{width:22px;height:12px;border:1px solid #000;flex:0 0 auto}
.sw.dot{border-radius:50%}
.legend p{font-size:12px;color:var(--muted);margin:11px 0 0;line-height:1.55}
@media (max-width:900px){.maprow{grid-template-columns:1fr}}
@media (max-width:760px){.tab{padding:10px 10px;font-size:13px}}
@media print{.js .page{display:block!important}.tabs{display:none}
 .page{break-before:page}body{background:#fff;color:#000}}
</style>
"""


def html_table(headers, rows):
    body = []
    for row in rows:
        tds = []
        for h, v in zip(headers, row):
            num = isinstance(v, (int, float)) and not isinstance(v, bool)
            cell = esc(fmt(h, v))
            if isinstance(v, float) and 0 <= v <= 1 and (
                    "%" in str(h).upper() or "RATE" in str(h).upper()):
                w = max(2, round(v * 54))
                # Rescaled with the EML denominator: the national rate is
                # 78.5%, so the old 95/85 cuts painted nearly every bar
                # red and the colour carried no information. Green now
                # means above 85%, amber 75-85% (the band the national
                # figure sits in), red below 75%.
                cls = "" if v >= .85 else (" mid" if v >= .75 else " low")
                cell += f'<span class="bar{cls}" style="width:{w}px"></span>'
            tds.append(f'<td class="{"num" if num else ""}">{cell}</td>')
        body.append("<tr>" + "".join(tds) + "</tr>")
    head = "".join(f"<th>{esc(h)}</th>" for h in headers)
    tall = " tall" if len(rows) > 18 else ""
    return (f'<div class="scroll{tall}"><table><thead><tr>{head}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody></table></div>')


HOSP_LEVELS = ("District hospital", "Level 2 teaching hospital",
               "National referral hospital")


def _hosp_rows(base):
    """FIG 6: one bar per hospital - [name, expected, available] - lowest
    availability first, ties by name."""
    g = defaultdict(lambda: [0, 0])
    for r in base:
        if r["_lvl"] in HOSP_LEVELS:
            v = g[s(r["_p"][M.FAC])]
            v[0] += 1
            v[1] += M.yes(r[M.AVAIL])
    return sorted(([k, n, a] for k, (n, a) in g.items()),
                  key=lambda x: (x[2] / x[1], x[0]))


def _split(rs):
    """(n, available) for one group.

    n is every medicine EXPECTED at that facility - EML-authorised at its
    level for Essential Medicines, managed for every other programme;
    available is the subset physically present on the visit day. The reason a
    medicine was not available - run out, or not stocked at all - stays in the
    tables and is deliberately not carried into the charts.
    """
    return len(rs), sum(1 for r in rs if M.yes(r[M.AVAIL]))


# MIN_N is calibrated for GROUP rates. Applied to a single facility's basket
# it does real damage: health posts are assessed on 16 to 45 medicines each
# (median 27), so a 30-observation cut dropped 15 of the 26 posts from FIG 5,
# and dropped them non-randomly - it kept only the best-covered ones, which
# flattered the level that is already the weakest. A facility rate over 16
# medicines is a usable rate, so the per-facility floor is its own number.
MIN_FAC = 10

# Per-STRENGTH records for the all-medicines table: every assessed item row,
# flagged available (1), managed and answered (2), expected under method 1
# (4) - Essential Medicines authorised at the level, or a programme product
# the facility manages.
def _item_rec(r):
    mg = M.yes(r[M.MANAGED]) and not M.blank(r[M.AVAIL])
    av = mg and M.yes(r[M.AVAIL])
    em = s(r[M.PROGRAM]) == MOL.ESSENTIAL
    ex = (em and M.yes(r[M.APPLICABLE])) or (not em and mg)
    p = r["_p"]
    return ([s(p[M.FAC]), s(p[M.DIST]), s(p[M.PROV]), FL.group(p[M.LEVEL]),
             s(p.get(EX.MGMT)) or "(not recorded)"],
            s(r[M.MED]), s(r[M.PROGRAM]) or "(none)",
            int(av) + 2 * int(mg) + 4 * int(ex))


EX_ITEMS = [_item_rec(r) for r in MED_ALL if s(r[M.MED])]
EX_DATA = EX.build(MOLS, M, FL, FL.group, rows_b=MOLS_B, items=EX_ITEMS)


def explorer_script():
    """The one script that redraws every finding on the availability page.

    Emitted last, after the maps, the figure containers and the table
    container exist in the document.
    """
    facs, meds, packed, its, ipacked = EX_DATA
    return EX.script(facs, meds, packed, MIN_N, MIN_FAC, MAP.scales_js(),
                     {"avail": CH.AVAIL, "rest": CH.REST,
                      "gon": CH.G_ON, "goff": CH.G_OFF,
                      "sk": CH.SK_SRC, "skd": CH.SK_DST},
                     items=its, ipacked=ipacked)


def charts_block():
    """Three disaggregations of availability, each in the shape that suits it.

    Therapeutic drug class is a RANKING of many categories, so it gets stacked
    bars - the eye compares bar ends down a column. Level of care is a short
    ordered ladder, so the same bars read as a gradient. Province is already
    mapped, so repeating it as a bar adds nothing; what the map cannot show is
    the spread of districts inside each province, which is what decides whether
    an action should be provincial or district-level - hence the dot plot.

    All three carry the same three-band split, because an availability bar on
    its own cannot distinguish a supply failure from a medicine the facility
    never stocks, and those need different responses.
    """

    # --- 2. level of care
    byl = defaultdict(list)
    for r in MOLS:
        byl[r["_lvl"]].append(r)
    lrows = [[k] + list(_split(byl[k])) for k in LEVELS if byl.get(k)]

    # --- 3. province, with its districts
    byp, bypd = defaultdict(list), defaultdict(lambda: defaultdict(list))
    for r in MOLS:
        byp[s(r["_p"][M.PROV])].append(r)
        bypd[s(r["_p"][M.PROV])][s(r["_p"][M.DIST])].append(r)
    prows = []
    for k in MAP.PROV_ORDER if MAP else sorted(byp):
        rs = byp.get(k)
        if not rs:
            continue
        ds = {d: sum(1 for x in v if M.yes(x[M.AVAIL])) / len(v)
              for d, v in bypd[k].items() if len(v) >= MIN_N}
        prows.append([k, len(rs),
                      sum(1 for x in rs if M.yes(x[M.AVAIL])) / len(rs), ds])
    prows.sort(key=lambda x: x[2])
    prows3 = [[k, n, round(v * n)] for k, n, v, _ in prows]
    CH.set_districts({k: len(ds) for k, _, _, ds in prows})

    # --- 4. who runs the facility, with its facilities as the spread. A
    # collected variable, not derived, so it stands on its own rather than
    # being read off the level of care.
    #
    # The one facility with no management type recorded is excluded here and
    # only here. It cannot be assigned to any of the three categories, and a
    # fourth row holding a single facility would read as a comparable group.
    # Every other figure on this page still counts it.
    bym = defaultdict(lambda: defaultdict(list))
    for r in MOLS:
        k = s(r["_p"].get(EX.MGMT))
        if not k:
            continue
        bym[k][s(r["_p"][M.FAC])].append(r)
    mrows = []
    for k, facs_ in bym.items():
        rs = [x for v in facs_.values() for x in v]
        fsp = {f: sum(1 for x in v if M.yes(x[M.AVAIL])) / len(v)
               for f, v in facs_.items() if len(v) >= MIN_FAC}
        mrows.append([k, len(rs),
                      sum(1 for x in rs if M.yes(x[M.AVAIL])) / len(rs), fsp])
    mrows.sort(key=lambda x: x[2])

    # --- 5. programme x province heat map (replaced the level -> drug
    # class Sankey, 2026-10-16), all provinces last; under MIN_N grey
    pv_cols = [k for k in (MAP.PROV_ORDER if MAP else sorted(byp))
               if byp.get(k)]
    pv_rows = []
    for p in sorted({s(r[M.PROGRAM]) for r in MOLS if s(r[M.PROGRAM])}):
        rs = [r for r in MOLS if s(r[M.PROGRAM]) == p]
        cells = [(sum(1 for r in rs if s(r["_p"][M.PROV]) == k),
                  sum(1 for r in rs if s(r["_p"][M.PROV]) == k
                      and M.yes(r[M.AVAIL]))) for k in pv_cols]
        cells.append((len(rs), sum(1 for r in rs if M.yes(r[M.AVAIL]))))
        pv_rows.append([p, cells])
    pv_rows.sort(key=lambda x: x[1][-1][1] / x[1][-1][0])

    # --- 6/7. managed but NOT applicable at that level per the EML.
    # MOLS is the applicable set, so these rows are not in it - they come
    # from the uncollapsed molecule view.
    OFF_ALL = MOL.collapse(MED_ALL)
    for r in OFF_ALL:
        r["_lvl"] = FL.group(r["_p"][M.LEVEL])
    off_mg = [r for r in OFF_ALL if M.yes(r[M.MANAGED])]
    off = [r for r in off_mg if not M.yes(r[M.APPLICABLE])]
    # One row carries no programme at all; below the reporting minimum it
    # would draw as an unlabelled bar, so it is suppressed like any other
    # thin cell rather than shown with an empty name.
    off_prog = sorted(
        [[p, sum(1 for r in off_mg if s(r[M.PROGRAM]) == p),
          sum(1 for r in off if s(r[M.PROGRAM]) == p)]
         for p in sorted({s(r[M.PROGRAM]) for r in off_mg}) if p],
        key=lambda x: -(x[2] / x[1] if x[1] else 0))
    off_prog = [r for r in off_prog if r[1] >= MIN_N]

    # --- 2. by programme (medicine_program); the one row with none is left
    # out like any other unlabelled cell
    prows = sorted(
        [[p, sum(1 for r in MOLS if s(r[M.PROGRAM]) == p),
          sum(1 for r in MOLS if s(r[M.PROGRAM]) == p
              and M.yes(r[M.AVAIL]))]
         for p in sorted({s(r[M.PROGRAM]) for r in MOLS}) if p],
        key=lambda x: x[2] / x[1])

    # --- 2. programme x level of care, as a heat map: one row per
    # programme (weakest overall first), one column per level, the row total
    # last; cells under MIN_N pairs grey
    hm_lv = [l for l in LEVELS if any(r["_lvl"] == l for r in MOLS)]
    hm_rows = []
    for p in sorted({s(r[M.PROGRAM]) for r in MOLS if s(r[M.PROGRAM])}):
        rs = [r for r in MOLS if s(r[M.PROGRAM]) == p]
        cells = [(sum(1 for r in rs if r["_lvl"] == l),
                  sum(1 for r in rs if r["_lvl"] == l and M.yes(r[M.AVAIL])))
                 for l in hm_lv]
        cells.append((len(rs), sum(1 for r in rs if M.yes(r[M.AVAIL]))))
        hm_rows.append([p, cells])
    hm_rows.sort(key=lambda x: x[1][-1][1] / x[1][-1][0])

    nat = pct(len(AVM), len(APPM))
    return (
        '<section class="ind">'
        '<h4><span class="mid">FIG 1</span>Availability by programme</h4>'
        # track columns, the user's template (2026-10-26)
        + f'<div id="fx-figp">' + CH.track_columns(prows) + '</div></section>'
        '<section class="ind">'
        '<h4><span class="mid">FIG 2</span>Availability by level of care</h4>'
        # dials, as the province figure that was removed on request
        # (2026-10-19) used to be drawn
        + f'<div id="fx-fig2">' + CH.dials(lrows) + '</div></section>'
        '<section class="ind">'
        '<h4><span class="mid">FIG 3</span>Availability of programmes by '
        'level of care</h4>'
        f'<p class="note">One row per programme, weakest first; one column '
        f'per level of care, all levels on the right. Cells on fewer than '
        f'{MIN_N} expected pairs are grey with their n. Hover a cell for the '
        f'counts.</p>'
        + f'<div id="fx-figpl">'
        + CH.heatmap(hm_rows, hm_lv + ["All levels"], min_n=MIN_N)
        + '</div></section>'
        '<section class="ind">'
        '<h4><span class="mid">FIG 4</span>Availability by management '
        'type</h4>'
        + f'<div id="fx-fig4">'
        + CH.dotplot(mrows, unit="facilities") + '</div></section>'
        '<section class="ind">'
        '<h4><span class="mid">FIG 5</span>Availability of programmes by '
        'province</h4>'
        + f'<div id="fx-fig5">'
        + CH.heatmap(pv_rows, pv_cols + ["All provinces"], min_n=MIN_N)
        + '</div></section>'
        # every hospital, lowest first (user request 2026-10-20)
        '<section class="ind">'
        '<h4><span class="mid">FIG 6</span>Availability by hospital</h4>'
        # compact bars: 36 rows (user request 2026-10-20)
        + CH.bars(_hosp_rows(MOLS), bar_h=12, gap=4) + '</section>'
        '<section class="ind">'
        '<h4><span class="mid">FIG 7</span>Unavailable at the facility - was '
        'the stock upstream? By level of care</h4>'
        + f'<div id="fx-fig8">'
        + CH.area_band([[k, [v[1], v[0]]] for k, v in up_lvl_both],
                       ["RMS Central", "RMS Regional"], connect=False)
        + '</div></section>'
        '<section class="ind">'
        '<h4><span class="mid">FIG 8</span>Regional against Central, by '
        'programme</h4>'
        '<p class="note">The same pairing by vertical programme. A programme '
        'where both bars are short is short at the national store itself.</p>'
        + f'<div id="fx-fig10">'
        + CH.grouped(up_prog_both, ["RMS Regional", "RMS Central"])
        + '</div></section>'
        '<section class="ind">'
        '<h4><span class="mid">FIG 9</span>Distribution Gap against '
        'nationwide stock-out, by level of care</h4>'
        f'<p class="note">{DEF_GAP}</p>'
        + f'<div id="fx-fig11">'
        + CH.columns(dn_lvl_both, DN_SERIES)
        + '</div></section>'
        '<section class="ind">'
        '<h4><span class="mid">FIG 10</span>Distribution Gap against '
        'nationwide stock-out, by programme</h4>'
        f'<p class="note">{DEF_GAP}</p>'
        + f'<div id="fx-fig12">'
        + CH.grouped(dn_prog_both, DN_SERIES)
        + '</div></section>')


def fmt_pc(v):
    return "-" if v is None else f"{v * 100:.1f}%"


# ---------------------------------------------------------------
# MAP 2 on a Mapbox base map
# ---------------------------------------------------------------
# The facility points are drawn with Mapbox GL JS over Mapbox's dark style,
# same colours and bubble sizes as the SVG map. The token is a PUBLIC (pk.)
# token read at build time from analytics/mapbox_token.txt or $MAPBOX_TOKEN
# - never written into this source, because the repository is public; the
# file is listed in .git/info/exclude. It does end up in the generated HTML,
# as every Mapbox web page's token does: restrict it to your own URLs in the
# Mapbox account. The SVG map stays in the page underneath and is only
# swapped out once Mapbox has loaded, so offline - or with no token - the
# map still shows.
MAPBOX_GL = "https://api.mapbox.com/mapbox-gl-js/v3.12.0"


def _mapbox_token():
    t = os.environ.get("MAPBOX_TOKEN", "").strip()
    if not t:
        p = os.path.join(B.HERE, "mapbox_token.txt")
        if os.path.exists(p):
            t = open(p).read().strip()
    return t if t.startswith("pk.") else ""


# MAP 2: COLOUR = level of care, SIZE = availability. One fixed colour per
# level so every facility of a level looks alike; the bubble's area carries
# the facility's own availability (MAP.radius - smallest at or below the
# 5th percentile, largest at 100%). Colours chosen to stay distinct on the
# satellite, light Streets and dark maps alike.
LEVEL_COLS = {"Health post": "#F28E2B", "Health centre": "#4E79A7",
              "District hospital": "#E15759", "Provincial hospital": "#59A14F",
              "Level 2 teaching hospital": "#B07AA1",
              "National referral hospital": "#EDC948"}


def _fac_legend():
    rows = "".join(
        f'<p class="key"><span class="sw dot" style="width:14px;height:14px;'
        f'background:{LEVEL_COLS.get(l, "#999")}"></span>{esc(l)} '
        f'<span class="lyrn">{MAP.LVL_COUNTS[l]}</span></p>'
        for l in FL.ordered(MAP.LVL_COUNTS))
    sizes = "".join(
        f'<span class="szk"><span class="sw dot" style="width:'
        f'{2 * MAP.radius(v):.0f}px;height:{2 * MAP.radius(v):.0f}px;'
        f'background:#9aa3ad"></span>{v:.0%}</span>'
        for v in (0.70, 0.85, 1.0))
    return (f'<div class="legend"><h4>Level of care (colour)</h4>{rows}'
            f'<h4 style="margin-top:14px">Availability (size)</h4>'
            f'<p class="key szrow">{sizes}</p>'
            f'<p class="key" style="color:var(--muted);font-size:11.5px">'
            f'Smallest at {MAP.V_LO:.0%} or below. Hover a bubble for the '
            f'facility name and its figures.</p></div>')


def facility_mapbox():
    """(map markup, script) - the script is empty when there is no token."""
    tok = _mapbox_token()
    # The fallback SVG is the standalone map page's, which colours and keys
    # by availability; here colour is the level of care, so its own
    # subtitle and availability key are dropped - the side panel's legend
    # covers both encodings.
    svg = MAP.facility_svg()
    svg = re.sub(r'<g class="slegend">.*?</g>', "", svg, count=1, flags=re.S)
    svg = re.sub(r'<text [^>]*class="msub">.*?</text>', "", svg, count=1,
                 flags=re.S)
    box = (f'<div id="facgl" style="display:none;height:640px"></div>'
           f'<div id="facsvg">{svg}</div>')
    if not tok:
        return box, ""
    # The token is NOT written into the HTML (GitHub push protection blocks
    # it, 2026-11-05): it goes to mapbox_token.js beside the page, a file
    # git ignores, and the page loads it. Without that file - e.g. a copy on
    # GitHub - the SVG map simply stays.
    with open(os.path.join(B.HERE, "mapbox_token.js"), "w") as fh:
        fh.write(f'window.MAPBOX_TOKEN = "{tok}";\n')
    feats = [{"type": "Feature",
              "geometry": {"type": "Point", "coordinates": [p["lo"], p["la"]]},
              "properties": {"f": p["fac"], "d": p["dist"], "l": p["lvl"],
                             "lv": MAP.slug(p["lvl"]), "n": p["n"],
                             "a": p["av"],
                             "p": round(p["pct"] * 100, 1)
                             if p["pct"] is not None else None,
                             "c": LEVEL_COLS.get(p["lvl"], "#999999"),
                             "r": round(MAP.radius(p["pct"]), 1), "h": 0}}
             for p in MAP.POINTS]
    data = json.dumps({"type": "FeatureCollection", "features": feats},
                      separators=(",", ":")).replace("</", "<\\/")
    # District borders from the same geoBoundaries ADM2 file as the SVG maps,
    # rounded to 4 decimals (~10 m) to keep the page small, plus one label
    # point per district at the mean of its outline.
    def rnd(c):
        return [round(c[0], 4), round(c[1], 4)] if isinstance(c[0], float) \
            else [rnd(x) for x in c]
    dfeat, lfeat = [], []
    for f in MAP.GJ["features"]:
        geom = {"type": f["geometry"]["type"],
                "coordinates": rnd(f["geometry"]["coordinates"])}
        name = s(f["properties"]["shapeName"])
        dfeat.append({"type": "Feature", "geometry": geom,
                      "properties": {"d": name}})
        ring = (f["geometry"]["coordinates"][0]
                if f["geometry"]["type"] == "Polygon"
                else max(f["geometry"]["coordinates"],
                         key=lambda p: len(p[0]))[0])
        lfeat.append({"type": "Feature", "properties": {"d": name},
                      "geometry": {"type": "Point", "coordinates": [
                          round(sum(x for x, _ in ring) / len(ring), 4),
                          round(sum(y for _, y in ring) / len(ring), 4)]}})
    dist = json.dumps({"type": "FeatureCollection", "features": dfeat},
                      separators=(",", ":")).replace("</", "<\\/")
    dlab = json.dumps({"type": "FeatureCollection", "features": lfeat},
                      separators=(",", ":")).replace("</", "<\\/")
    script = (
        f'<link href="{MAPBOX_GL}/mapbox-gl.css" rel="stylesheet">'
        f'<script src="{MAPBOX_GL}/mapbox-gl.js"></script>'
        '<script src="mapbox_token.js"></script>'
        '<script>\n'
        '(function(){\n'
        '  if (typeof mapboxgl === "undefined" || !window.MAPBOX_TOKEN) return;\n'
        '  mapboxgl.accessToken = window.MAPBOX_TOKEN;\n'
        f'  var pts = {data};\n'
        f'  var dist = {dist};\n'
        f'  var dlab = {dlab};\n'
        '  var map;\n'
        '  try {\n'
        '    map = new mapboxgl.Map({container: "facgl",\n'
        '      style: "mapbox://styles/mapbox/satellite-streets-v12",\n'
        '      bounds: [[28.84, -2.84], [30.90, -1.05]],\n'
        # locked to Rwanda: no panning or zooming out to the world map
        # (user request 2026-10-26)
        '      maxBounds: [[28.3, -3.2], [31.4, -0.7]],\n'
        '      renderWorldCopies: false, dragRotate: false,\n'
        '      fitBoundsOptions: {padding: 24}});\n'
        '  } catch (e) { return; }\n'
        '  map.addControl(new mapboxgl.NavigationControl({showCompass: false}));\n'
        '  map.addControl(new mapboxgl.ScaleControl({unit: "metric"}));\n'
        '  var STYLES = [["Satellite", "satellite-streets-v12"],\n'
        '                ["Streets", "streets-v12"],\n'
        '                ["Dark", "dark-v11"]];\n'
        '  var sw = document.createElement("div");\n'
        '  sw.className = "mapboxgl-ctrl mapboxgl-ctrl-group fac-style";\n'
        '  STYLES.forEach(function(s, i){\n'
        '    var b = document.createElement("button");\n'
        '    b.type = "button"; b.textContent = s[0];\n'
        '    if (i === 0) b.className = "on";\n'
        '    b.onclick = function(){\n'
        '      sw.querySelectorAll("button").forEach(function(x){ x.className = ""; });\n'
        '      b.className = "on";\n'
        '      map.setStyle("mapbox://styles/mapbox/" + s[1]);\n'
        '    };\n'
        '    sw.appendChild(b);\n'
        '  });\n'
        '  map.addControl({onAdd: function(){ return sw; },\n'
        '    onRemove: function(){}}, "top-left");\n'
        '  var pop = new mapboxgl.Popup({closeButton: false,\n'
        '    closeOnClick: false});\n'
        '  function levels(){\n'
        '    if (!map.getLayer("fac")) return;\n'
        '    var on = [];\n'
        '    document.querySelectorAll(".layers input[data-lv]")\n'
        '      .forEach(function(cb){ if (cb.checked) on.push(cb.dataset.lv); });\n'
        '    map.setFilter("fac", ["all", ["in", ["get", "lv"],\n'
        '      ["literal", on]], ["==", ["get", "h"], 0]]);\n'
        '  }\n'
        '  // The page filters (method switch included) recompute each\n'
        '  // facility and hand the values here: [n, available, colour, r].\n'
        '  window.fxFacHook = function(gl){\n'
        '    if (!map.getSource("fac")) return;\n'
        '    pts.features.forEach(function(f){\n'
        '      var q = f.properties, g = gl[q.f];\n'
        '      if (!g) { q.h = 1; return; }\n'
        '      q.h = 0; q.n = g[0]; q.a = g[1];\n'
        '      q.p = (g[1] / g[0] * 100).toFixed(1); q.r = g[3];\n'
        '    });\n'
        '    map.getSource("fac").setData(pts); levels();\n'
        '  };\n'
        '  function addFac(){\n'
        '    if (map.getSource("fac")) return;\n'
        '    map.addSource("dist", {type: "geojson", data: dist});\n'
        '    map.addSource("dlab", {type: "geojson", data: dlab});\n'
        '    map.addLayer({id: "dist-halo", type: "line", source: "dist",\n'
        '      paint: {"line-color": "#ffffff", "line-width": 4,\n'
        '        "line-opacity": 0.55}});\n'
        '    map.addLayer({id: "dist-line", type: "line", source: "dist",\n'
        '      paint: {"line-color": "#2b2b2b", "line-width": 1.6}});\n'
        '    map.addLayer({id: "dist-name", type: "symbol", source: "dlab",\n'
        '      layout: {"text-field": ["get", "d"], "text-size": 12,\n'
        '        "text-font": ["DIN Pro Bold", "Arial Unicode MS Bold"],\n'
        '        "text-allow-overlap": false},\n'
        '      paint: {"text-color": "#1a1a1a", "text-halo-color": "#ffffff",\n'
        '        "text-halo-width": 1.6}});\n'
        '    map.addSource("fac", {type: "geojson", data: pts});\n'
        '    map.addLayer({id: "fac", type: "circle", source: "fac",\n'
        '      layout: {"circle-sort-key": ["-", 0, ["get", "r"]]},\n'
        '      paint: {"circle-color": ["get", "c"],\n'
        '        "circle-radius": ["get", "r"],\n'
        '        "circle-stroke-color": "#000",\n'
        '        "circle-stroke-width": 1.2,\n'
        '        "circle-opacity": 0.95}});\n'
        '    if (window.fxFacLast) window.fxFacHook(window.fxFacLast);\n'
        '    else levels();\n'
        '  }\n'
        '  map.on("style.load", addFac);\n'
        '  map.on("mousemove", "fac", function(e){\n'
        '    map.getCanvas().style.cursor = "pointer";\n'
        '    var q = e.features[0].properties;\n'
        '    pop.setLngLat(e.features[0].geometry.coordinates).setHTML(\n'
        '      "<div class=\\"fpop\\"><b>" + q.f + "</b>" +\n'
        '      "<span>Level: " + q.l + "</span>" +\n'
        '      "<span>District: " + q.d + "</span>" +\n'
        '      "<span>Availability: <b>" + q.p + "%</b></span>" +\n'
        '      "<span>" + q.a + " of " + q.n + " expected medicines on the "\n'
        '      + "shelf</span></div>").addTo(map);\n'
        '  });\n'
        '  map.on("mouseleave", "fac", function(){\n'
        '    map.getCanvas().style.cursor = ""; pop.remove();\n'
        '  });\n'
        '  map.on("load", function(){\n'
        '    document.getElementById("facgl").style.display = "block";\n'
        '    document.getElementById("facsvg").style.display = "none";\n'
        '    map.resize();\n'
        '    addFac();\n'

        '    document.querySelectorAll(".layers input[data-lv]")\n'
        '      .forEach(function(cb){ cb.addEventListener("change", levels); });\n'
        '    document.querySelectorAll(".lyrbtn button")\n'
        '      .forEach(function(b){ b.addEventListener("click", levels); });\n'
        '    document.querySelectorAll(".tab").forEach(function(t){\n'
        '      t.addEventListener("click", function(){\n'
        '        setTimeout(function(){ map.resize(); }, 50); });\n'
        '    });\n'
        '  });\n'
        '})();\n'
        '</script>')
    return box, script


def _fac_svg_css():
    """The SVG fallback map colours its circles by level too - CSS outranks
    the fill attribute the availability ramp writes."""
    return "".join(f'#facmap .pt.lv-{MAP.slug(l)}{{fill:{c}}}'
                   for l, c in LEVEL_COLS.items())


def maps_block():
    """The district choropleth and the facility points, for Analysis 1."""
    if MAP is None:
        return ('<p class="note">Maps unavailable: '
                f'{esc(str(MAP_ERR))}. Run Tracer_availability_mapping.py '
                'once to fetch the district boundaries.</p>')
    located = len(MAP.POINTS)
    _fbox, _fscript = facility_mapbox()
    return (
        '<section class="ind">'
        '<h4><span class="mid">MAP 1</span>Tracer availability by district</h4>'

        f'<div class="maprow"><div class="mapbox">{MAP.district_svg()}</div>'
        f'<div>{MAP.legend("d")}{MAP.view_control()}</div></div></section>'
        '<section class="ind">'
        '<h4><span class="mid">MAP 2</span>All-medicines availability by '
        'facility</h4>'
        f'<p class="note">{located} of {len(MAP.FAC)} assessed facilities '
        f'located. '
        + (f'{len(MAP.UNLOCATED)} facilities are absent from the map because '
           f'neither source has a position for them; they remain in every '
           f'district figure.' if MAP.UNLOCATED else '')
        + f'</p><div class="maprow"><div class="mapbox">{_fbox}</div>'
        f'<div>{_fac_legend()}{MAP.layer_control()}</div>'
        f'</div><style>{_fac_svg_css()}</style>{_fscript}</section>')


def cfac(r):
    return C.s(r["_p"][C.FAC])


add(9, "9.6", "CHW availability by province and level of care", RATE_HDR,
    rate_by(CANS, lambda r: C.s(r["_p"][C.PROV]),
            lambda r: C.yes(r[C.AVAIL]), cfac)
    + [["", "", "", "", None]]
    + rate_by(CANS, lambda r: FL.group(r["_p"][C.LEVEL]),
              lambda r: C.yes(r[C.AVAIL]), cfac),
    "Over the products each CHW manages and answered for. 'Not managed' at "
    "CHW level is MFT zoning - which first-line therapy that zone uses - and "
    "is policy, not a stock-out, so it stays out of the denominator.")

add(6, "6.6", "Attribution by level of care", RATE_HDR,
    rate_by([r for r in UNSERVED if not P.blank(r.get(P.RMS_REG))],
            lambda r: FL.group(r["_sub"][P.FTYPE]),
            lambda r: P.yes(r.get(P.RMS_REG)), pfac),
    "Of the unserved items where an upstream answer was given, the share "
    "reported available at RMS Regional - a facility-level problem rather "
    "than a national one. Small denominators; read with the counts.")


# ---------------------------------------------------------------
# one visual per page
# ---------------------------------------------------------------
# Page 1 opens with maps and figures; every other page opened with a wall of
# tables. Each page now leads with the one chart that carries its headline
# measure across the cut a reader asks for first. Bars take (label, n, hits),
# which is the same shape rate_by() already produces, so the figure and the
# table underneath it cannot disagree.
def _bar_rows(tbl, top=None):
    """rate_by() output -> CH.bars() rows, dropping suppressed cells."""
    out = [[r[0], r[2], r[3]] for r in tbl if r[0] and r[4] is not None]
    return out[:top] if top else out





# ---------------------------------------------------------------
# FIG 2.9: record accuracy on the district map
# ---------------------------------------------------------------
# Same boundaries, projection and styling as MAP 1 on Analysis 1, but its own
# colour classes: availability spans 65-84% across districts while record
# accuracy spans roughly 17-100%, so the availability breaks would put almost
# every district in the end classes. Colours are the user's single teal
# ramp, light to dark, so lighter means less accurate. Six shades, so the
# 60-79% band is split in two - it holds the most districts.
ACC_STEPS = [(0.00, 0.30, "#c6dce6"), (0.30, 0.45, "#a9cfdf"),
             (0.45, 0.60, "#85bccf"), (0.60, 0.70, "#5ea9c0"),
             (0.70, 0.80, "#4d8fa3"), (0.80, 1.01, "#3f7a8c")]
ACC_LABELS = ["below 30%", "30 to 44%", "45 to 59%", "60 to 69%",
              "70 to 79%", "80% and above"]


def _acc_colour(v):
    if v is None:
        return "#2c2c2a"
    for lo, hi, c in ACC_STEPS:
        if lo <= v < hi:
            return c
    return ACC_STEPS[-1][2]


def _acc_map_svg(byd, title, sub_cls, sub, tip, lab):
    """One district map. byd: district -> (n, value or None, detail)."""
    W, H, PAD = MAP.W, MAP.H, MAP.PAD
    out = [f'<svg viewBox="0 0 {W} {H + 54}" role="img" '
           f'aria-label="Rwanda {esc(title)} by district">',
           f'<text x="{PAD}" y="18" class="mtitle">{esc(title)}</text>',
           f'<text x="{PAD}" y="34" class="msub {sub_cls}">{sub}</text>']
    for d in MAP.DROWS:
        n, v, det = byd.get(d["district"], (0, None, ""))
        out.append(f'<path class="dist" data-d="{esc(d["district"])}" '
                   f'd="{d["path"]}" '
                   f'fill="{_acc_colour(v)}"><title>{esc(d["district"])} '
                   f'district, {esc(d["prov"])} - '
                   + (f'{fmt_pc(v)} {tip}' if v is not None
                      else f'fewer than {MIN_N} records')
                   + f'\n{det}</title></path>')
    for d in MAP.DROWS:
        n, v, _ = byd.get(d["district"], (0, None, ""))
        cx, cy = d["cx_cy"]
        out.append(f'<text class="dlabel" x="{cx:.0f}" y="{cy:.0f}">'
                   f'{esc(d["district"])}</text>')
        out.append(f'<text class="dlabel" x="{cx:.0f}" y="{cy + 11:.0f}" '
                   f'data-dv="{esc(d["district"])}" style="font-size:10.5px">'
                   f'{fmt_pc(v) if v is not None else "n/a"}</text>')
    out.append(MAP.north_arrow(W - PAD - 16, 66))
    out.append(MAP.scale_bar(PAD + 4, H + 10))
    out.append(f'<text x="{PAD + 4}" y="{H + 36}" class="msrc">Boundaries: '
               f'geoBoundaries ADM2 &#183; source Open Data Rwanda / NISR '
               f'&#183; CC BY 4.0.</text>')
    out.append("</svg>")
    return "".join(out)


def accuracy_map():
    """FIG 2.9: exact match and average quantity accuracy by district, two
    maps side by side on the same colour classes."""
    if MAP is None:
        return ('<p class="note">Map unavailable: '
                f'{esc(str(MAP_ERR))}.</p>')
    ex = {k: (n, v, f"{hit:,} of {n:,} records match the shelf")
          for k, _, n, hit, v in
          rate_by(_acc, lambda r: s(r["_p"][M.DIST]), _exact, mfac)}
    g = defaultdict(list)
    for r in _acc:
        g[s(r["_p"][M.DIST])].append(_qacc(r))
    qa = {k: (len(v), statistics.mean(v) if len(v) >= MIN_N else None,
              f"average over {len(v):,} records") for k, v in g.items()}
    nd = len(MAP.DROWS)
    left = _acc_map_svg(
        ex, "Exact match", "msub-acc",
        f"Physical count equals the system count &#183; {nd} districts "
        f"&#183; overall "
        f"{fmt_pc(pct(sum(1 for r in _acc if _exact(r)), len(_acc)))}",
        "accurate", "exact")
    right = _acc_map_svg(
        qa, "Average quantity accuracy", "msub-qacc",
        f"(1 - |system - physical| / system) &#183; {nd} districts &#183; "
        f"overall {fmt_pc(statistics.mean([_qacc(r) for r in _acc]))}",
        "average quantity accuracy", "qty")
    keys = "".join(f'<p class="key"><span class="sw" style="background:{c}">'
                   f'</span>{esc(l)}</p>'
                   for (_, _, c), l in zip(ACC_STEPS, ACC_LABELS))
    keys += ('<p class="key"><span class="sw" style="background:#2c2c2a">'
             f'</span>fewer than {MIN_N} records</p>')
    return ('<section class="ind"><h4><span class="mid">FIG 2.9</span>'
            'Inventory record accuracy by district</h4>'
            + _pair_html(f'<div class="mapbox" id="sx-f210">{left}</div>',
                         f'<div class="mapbox" id="sx-q210">{right}</div>')
            + f'<div class="legend legend-row"><h4>Record accuracy</h4>'
            f'{keys}</div></section>')


def _worst_first(tbl):
    """rate_by() output -> bar rows, HIGHEST rate first - for a rate where
    more is worse, which rate_by's ascending sort gets the wrong way round."""
    return _bar_rows(sorted(tbl, key=lambda x: (x[4] is None, -(x[4] or 0))))


def _day_max(rows):
    top = max((r[2] for r in rows), default=0)
    return 92 if top <= 92 else int(math.ceil(top / 30) * 30)


def _days_fig(fid, keyfn, title, note):
    """Mean stock-out days per AFFECTED pair, by group, longest first. Days
    are the form's total_stockout_days (_tdays), on pairs where it is > 0."""
    g = defaultdict(list)
    for r in MG:
        d = _tdays(r)
        if d > 0:
            g[keyfn(r)].append(d)
    rows = sorted(([k, len(v), statistics.mean(v)] for k, v in g.items()
                   if k and len(v) >= MIN_N), key=lambda x: -x[2])
    return (f'<section class="ind"><h4><span class="mid">{esc(fid)}</span>'
            f'{esc(title)}</h4><p class="note">{esc(note)}</p>'
            # total_stockout_days can pass 92, so the scale runs to the next
            # 30 days above the longest mean (never under 92)
            f'<div>{CH.day_bars(rows, fill=SO_BAD, track=SO_GOOD, max_v=_day_max(rows))}</div>'
            f'</section>')


# ---------------------------------------------------------------
# Analysis 5 opening: prescription fulfilment by place, class, programme
# ---------------------------------------------------------------
# Fulfilment = prescribed items recorded available / items with an answer.
# Its LEVEL is a design artefact of the instrument (prescribers write what
# the pharmacy holds), so these figures compare places and classes against
# each other; they are not a supply rate.
RX_ANS = [r for r in ITEMS if r["_avail"] is not None]


def _rx_ok(r):
    return r["_avail"] is True


def pdist(r):
    return s(r["_sub"][P.DISTR])


# The prescription form records no management type; it is taken from the
# stock assessment, which visited the same facilities, by facility name.
# Names that differ between the two tools do not match and are left out of
# the management figure only.
_MGMT_BY_FAC = {s(r["_p"][M.FAC]).lower(): s(r["_p"].get(EX.MGMT))
                for r in M.MEDROWS if s(r["_p"].get(EX.MGMT))}


def rx_mgmt(r):
    return _MGMT_BY_FAC.get(s(r["_sub"][P.FACIL]).lower())


def rx_level(r):
    return FL.group(r["_sub"][P.FTYPE])


# The prescription form records no programme. It is DERIVED from the EML
# section, plus the five MCCH emergency medicines by name, so the figure can
# be read beside the stock assessment's medicine_program.
_RX_PROG = [("Malaria", r"^11\.7\.4\.2|antimalar"),
            ("TB", r"^11\.5\b|tubercul"),
            ("ARV", r"^11\.7\.2\b|antiretrovir"),
            ("FP", r"^23\.2\b|contracep|family planning"),
            ("MCCH", r"^17\.[12]\b|oxyto|carbetocin|caffeine_citrate|"
                     r"caffeine citrate|tranexamic|magnesium_sulfate_inj"),
            ("Consumables", r"^2[78]\.|consumable")]


def rx_prog(r):
    txt = f'{s(r["_class_no"])} {s(r["_class"])} {s(r["_med"])}'
    for prog, pat in _RX_PROG:
        if re.search(pat, txt, re.I):
            return prog
    return "Essential Medicines"


# Map colours: the user's teal ramp, light = weakest. Fulfilment bunches at
# the top (15 of 30 districts at 100%), so the breaks are tight there and one
# wide class catches the outliers below 95%.
RX_STEPS = [(0.000, 0.950, "#c6dce6"), (0.950, 0.990, "#a9cfdf"),
            (0.990, 0.995, "#5ea9c0"), (0.995, 0.99999, "#4d8fa3"),
            (0.99999, 1.01, "#3f7a8c")]
RX_LABELS = ["below 95%", "95 to 98.9%", "99 to 99.4%", "99.5 to 99.9%",
             "100%"]


def _step_colour(v, steps):
    if v is None:
        return "#2c2c2a"
    for lo, hi, c in steps:
        if lo <= v < hi:
            return c
    return steps[-1][2]


def _choro_svg(units, vals, steps, title, sub, noun):
    """One choropleth layer. `units` are (name, path, (cx, cy), province)
    and `vals` {name: (n, hits, rate)}; rate None = below MIN_N."""
    W, H, PAD = MAP.W, MAP.H, MAP.PAD
    out = [f'<svg viewBox="0 0 {W} {H + 54}" role="img" '
           f'aria-label="{esc(title)}">',
           f'<text x="{PAD}" y="18" class="mtitle">{esc(title)}</text>',
           f'<text x="{PAD}" y="34" class="msub">{esc(sub)}</text>']
    for name, path, _, prov in units:
        n, hit, v = vals.get(name, (0, 0, None))
        where = f"{name}, {prov}" if prov and prov != name else name
        out.append(f'<path class="dist pv-{MAP.slug(prov or name)}" '
                   f'data-n="{esc(name)}" d="{path}" '
                   f'fill="{_step_colour(v, steps)}"><title>{esc(where)} - '
                   + (fmt_pc(v) if v is not None else f"fewer than {MIN_N}")
                   + f'\\n{hit:,} of {n:,} {esc(noun)}</title></path>')
    big = len(units) <= 5
    for name, _, (cx, cy), prov in units:
        n, hit, v = vals.get(name, (0, 0, None))
        pv = f"pv-{MAP.slug(prov or name)}"
        out.append(f'<text class="dlabel {pv}" x="{cx:.0f}" y="{cy:.0f}"'
                   + (' style="font-size:13px"' if big else '')
                   + f'>{esc(name)}</text>')
        out.append(f'<text class="dlabel {pv}" x="{cx:.0f}" y="{cy + (14 if big else 11):.0f}" '
                   f'data-nv="{esc(name)}" '
                   f'style="font-size:{12 if big else 10.5}px">'
                   f'{fmt_pc(v) if v is not None else "n/a"}</text>')
    out.append(MAP.north_arrow(W - PAD - 16, 66))
    out.append(MAP.scale_bar(PAD + 4, H + 10))
    out.append(f'<text x="{PAD + 4}" y="{H + 36}" class="msrc">Boundaries: '
               f'geoBoundaries ADM2 &#183; source Open Data Rwanda / NISR '
               f'&#183; CC BY 4.0.</text>')
    out.append("</svg>")
    return "".join(out)


def two_view_map(uid, fid, title, note, rows, keyd, keyp, okfn, facfn,
                 steps, labels, legend_title, noun, unit="prescribed items",
                 dvals=None, pvals=None, total=None, n_label=None):
    """Province and District as two tabs of one figure - plain CSS radios,
    so the switch works with the page script blocked. A COUNT rate is
    computed here from rows / okfn; a quantity-weighted one (fill rate) is
    passed in ready-made as dvals / pvals {name: (n, hits, rate)}."""
    def vals(keyfn):
        return {k: (n, hit, v) for k, _, n, hit, v in
                rate_by(rows, keyfn, okfn, facfn)}
    dv = dvals if dvals is not None else vals(keyd)
    pv = pvals if pvals is not None else vals(keyp)
    tot = total if total is not None else \
        pct(sum(1 for r in rows if okfn(r)), len(rows))
    dunits = [(d["district"], d["path"], d["cx_cy"], d["prov"])
              for d in MAP.DROWS]
    punits = [(p["prov"], p["path"], MAP._prov_centroid(p["prov"]),
               p["prov"]) for p in MAP.PROWS]
    sub = f"{n_label or f'{len(rows):,} {unit}'} - overall {fmt_pc(tot)}"
    keys = "".join(f'<p class="key"><span class="sw" style="background:{c}">'
                   f'</span>{esc(l)}</p>'
                   for (_, _, c), l in zip(steps, labels))
    provs = [p["prov"] for p in MAP.PROWS]
    opts = "".join(
        f'<option value="{MAP.slug(p)}">{esc(p)} &#183; '
        f'{fmt_pc(pv[p][2]) if p in pv else "n/a"}</option>' for p in provs)
    # The side panel copies MAP 1's: VIEW radios with counts, then a
    # province filter that dims the other districts. MAP 1 does this in
    # script; here it is CSS :has(), so it works with script blocked.
    css = "".join(
        f'#{uid}:has(option[value="{MAP.slug(p)}"]:checked) .mp-d '
        f':is(.dist,.dlabel):not(.pv-{MAP.slug(p)}){{opacity:.18}}'
        for p in provs)
    panel = (f'<div><div class="legend"><h4>{esc(legend_title)}</h4>{keys}'
             f'</div><div class="layers"><p class="lyrh">View</p>'
             f'<label class="lyr"><input type="radio" class="v-d" '
             f'name="{uid}-v" checked> Districts <span class="lyrn">'
             f'{len(dunits)}</span></label>'
             f'<label class="lyr"><input type="radio" class="v-p" '
             f'name="{uid}-v"> Provinces <span class="lyrn">{len(punits)}'
             f'</span></label>'
             f'<p class="lyrh" style="margin-top:14px">Filter &#183; '
             f'province</p><select class="lyrsel" aria-label="Filter by '
             f'province"><option value="">All provinces</option>{opts}'
             f'</select></div></div>')
    return (f'<section class="ind"><h4><span class="mid">{esc(fid)}</span>'
            f'{esc(title)}</h4>'
            + (f'<p class="note">{esc(note)}</p>' if note else '')
            + f'<style>{css}</style>'
            f'<div class="vmap maprow" id="{uid}"><div class="mapbox">'
            f'<div class="mp-d">'
            + _choro_svg(dunits, dv, steps, f"{title} - districts", sub, noun)
            + '</div><div class="mp-p">'
            + _choro_svg(punits, pv, steps, f"{title} - provinces", sub, noun)
            + f'</div></div>{panel}</div></section>')


# The unit is the PRESCRIPTION - what the patient walks out with - not the
# item. A prescription is COMPLETE when every item on it was available; it
# has a MISSING ITEM when at least one was not (partially served if some
# were, nothing served if none were). Every item carries an answer, so no
# prescription drops out.
RX_PRES = [r for r in PRES if r["_n_items"] > 0]


def _rx_complete(r):
    return r["_fully_served"]


def _rx_outcome(rs):
    """[complete, partially served, nothing served] counts."""
    return [sum(1 for r in rs if r["_fully_served"]),
            sum(1 for r in rs if r["_any_unserved"]
                and r["_n_unavail"] < r["_n_items"]),
            sum(1 for r in rs if r["_n_items"]
                and r["_n_unavail"] == r["_n_items"])]


# ---------------------------------------------------------------
# Unserved medicines - the unit is the prescribed ITEM
# ---------------------------------------------------------------
# An unserved medicine is a prescribed medicine recorded NOT available at the
# facility on the day. Overall, by drug class, by programme, then every
# medicine that went missing.
RX_MISS = [r for r in RX_ANS if r["_avail"] is False]

# Drug class is the EML sub-section, grouped on the section CODE because
# the label is spelled more than one way, and shown under its commonest
# spelling.
_rx_lab = defaultdict(Counter)
for _r in RX_ANS:
    _rx_lab[s(_r["_class_no"])][s(_r["_class"])] += 1


def rx_class(r):
    no = s(r["_class_no"])
    if not no:
        return "(class not recorded)"
    lab = _rx_lab[no].most_common(1)[0][0]
    return lab if len(lab) <= 46 else lab[:44].rstrip() + "..."


# Medicine names arrive as an EML code (m1105), a slug
# (ciprofloxacin_tab_500mg) or free text. Newer form versions also carry a
# readable "Medicine (name)"; any row that has one names its code for all.
_rx_name = defaultdict(Counter)
for _r in ITEMS:
    _n = s(_r.get("Medicine (name)"))
    if _n:
        _rx_name[_r["_med"]][_n] += 1


def _rx_raw_name(r):
    m = r["_med"]
    if _rx_name.get(m):
        return _rx_name[m].most_common(1)[0][0]
    if m.endswith(" (free text)"):
        return m
    t = m.replace("_", " ").strip()
    return t[:1].upper() + t[1:]


# One medicine, one name (user request 2026-10-20). The same product arrives
# under several spellings - the form's readable label ("Long-acting insulins
# - Inj.100 IU/ml") and a slug-derived one ("Long acting insulins inj 100 iu
# ml"). Names that match once case, spaces and punctuation are ignored are
# one medicine; the readable label (the " - " form) is kept, else the most
# common spelling.
def _rx_key(n):
    return re.sub(r"[^a-z0-9]", "", n.lower())


_rx_spell = defaultdict(Counter)
for _r in ITEMS:
    _n = _rx_raw_name(_r)
    _rx_spell[_rx_key(_n)][_n] += 1
_RX_CANON = {}
for _k, _c in _rx_spell.items():
    _lab = [n for n in _c if " - " in n]
    _RX_CANON[_k] = (max(_lab, key=lambda n: (_c[n], n)) if _lab
                     else _c.most_common(1)[0][0])


def rx_med_name(r):
    n = _rx_raw_name(r)
    return _RX_CANON.get(_rx_key(n), n)


# Table 5.9 is built before rx_med_name() exists, keyed on the form's code
# (m1147, m1105, ...). Rebuilt here on the readable name (user request
# 2026-10-20), in place, so its position on the page is unchanged.
def _rx_top_unserved():
    names = sorted({rx_med_name(r) for r in UNSERVED})
    rows = []
    for m in names:
        un = [r for r in UNSERVED if rx_med_name(r) == m]
        n = sum(1 for r in ITEMS if rx_med_name(r) == m)
        rows.append([m, len(un), n, pct(len(un), n),
                     len({s(r["_sub"][P.FACIL]) for r in un}),
                     round(sum(P.num(r.get(P.QTY)) or 0 for r in un))])
    return sorted(rows, key=lambda x: (-x[1], -(x[3] or 0)))[:15]


TABLES[:] = [t[:6] + (_rx_top_unserved(),) if (t[0], t[1]) == (5, "5.9")
             else t for t in TABLES]


def _rx_prog_rows():
    """FIG 5.3 by programme (user request 2026-10-20), as rate_by() rows:
    per programme, the prescriptions holding at least one of its medicines,
    and of those, the ones where one of ITS medicines was unavailable."""
    g = defaultdict(lambda: [set(), set(), set()])
    for r in RX_ANS:
        p, pid = rx_prog(r), id(r["_pres"])
        g[p][0].add(pid)
        g[p][2].add(pfac(r))
        if r["_avail"] is False:
            g[p][1].add(pid)
    out = [[p, len(v[2]), len(v[0]), len(v[1]),
            pct(len(v[1]), len(v[0])) if len(v[0]) >= MIN_N else None]
           for p, v in sorted(g.items())]
    out.sort(key=lambda x: (x[4] is None, x[4]))
    return out


def _miss_rows(keyfn, min_n=1):
    """Classes / programmes that had at least one unserved medicine, by unserved
    share, highest first - a long list of 0% bars says nothing."""
    tbl = rate_by(RX_ANS, keyfn, lambda r: r["_avail"] is False, pfac,
                  min_n=min_n)
    return _worst_first([x for x in tbl if x[3]])


def _missing_table():
    by = defaultdict(list)
    for r in RX_MISS:
        by[rx_med_name(r)].append(r)
    presc = Counter(rx_med_name(r) for r in RX_ANS)
    rows = []
    for m, rs in by.items():
        rows.append([m, rx_class(rs[0]), rx_prog(rs[0]), presc[m], len(rs),
                     pct(len(rs), presc[m]),
                     len({id(r["_pres"]) for r in rs}),
                     len({pfac(r) for r in rs}),
                     ", ".join(sorted({pdist(r) for r in rs})),
                     round(sum(P.num(r.get(P.QTY)) or 0 for r in rs)),
                     round(sum(P.num(r.get(P.COST_UNSERVED)) or 0
                               for r in rs))])
    rows.sort(key=lambda x: (-x[4], -x[10]))
    return html_table(
        ["Medicine", "Drug class", "Programme", "Times prescribed",
         "Times unserved", "UNSERVED %", "Prescriptions affected",
         "Facilities", "Districts", "Qty unserved", "Value unserved (RWF)"],
        rows)


NT_LABELS = [("INN", "Generic (INN)"), ("Brand", "Brand")]
NT_COLS = ["#1f9e7a", "#d98030"]


def _name_type_mix(items):
    """[(label, count)] of name types among `items`, INN first."""
    c = Counter(s(r.get(P.NAME_TYPE)) for r in items)
    return [(lab, c.get(k, 0)) for k, lab in NT_LABELS]


def _missing_block():
    return (
        '<section class="ind"><h4><span class="mid">FIG 5.5</span>'
        'Unserved prescriptions - overall unserved medicines</h4>'
        '<div style="max-width:420px;margin:0 auto">'
        + CH.donut1(len(RX_MISS), len(RX_ANS),
                    caption="of medicines unserved")
        + '</div></section>'
        + fig("FIG 5.6", _miss_rows(rx_prog),
              "Unserved medicines by programme",
              "Programmes with at least one unserved medicine, highest first. "
              "Programme is derived from the EML section and the MCCH "
              "emergency medicines by name; the form records none.",
              hit="unserved", base="medicines prescribed")
        + '<section class="ind"><h4><span class="mid">TABLE 5A</span>'
        'Every unserved medicine</h4>'
        '<p class="note">Each medicine that was unserved at least once, most '
        'often unserved first. UNSERVED % is times unserved over times '
        'prescribed anywhere in the study. Value is the form\'s own '
        'cost-of-unserved field.</p>'
        + _missing_table() + '</section>')


def _rx_figs():
    return (
        '<section class="ind"><h4><span class="mid">FIG 5.1</span>'
        'Overall complete prescriptions</h4><p class="note">Every '
        'prescription in the study: complete (every item available), '
        'partially served (some items missing) or nothing served. The two '
        'small slices are pulled out, turned to the front and drawn at a '
        'minimum width so they can be seen at all - read the shares from '
        'the key, not from the slice sizes.</p>'
        + CH.ring3d(list(zip(["Complete", "Partially served",
                              "Nothing served"], _rx_outcome(RX_PRES))),
                    ["#199e70", "#eda100", "#c9544a"],
                    title=f"{fmt_pc(pct(sum(1 for r in RX_PRES if r['_fully_served']), len(RX_PRES)))} of prescriptions complete",
                    sub=f"{len(RX_PRES):,} prescriptions")
        + '</section>'
        + two_view_map(
            "rxmap", "FIG 5.2", "Complete prescriptions", "",
            RX_PRES, lambda r: s(r["_sub"][P.DISTR]), pprov, _rx_complete,
            pfac, RX_STEPS, RX_LABELS, "Prescriptions complete",
            "prescriptions complete", unit="prescriptions")
        + fig("FIG 5.3", _worst_first(_rx_prog_rows()),
              "Prescriptions with a missing item, by programme", "",
              hit="with a missing item", base="prescriptions")
        + fig("FIG 5.4",
              _bar_rows(rate_by(RX_PRES, rx_level,
                                lambda r: r["_any_unserved"], pfac,
                                keys=LEVELS, sort=False)),
              "Prescriptions with a missing item, by level of care",
              f"Share of prescriptions with at least one item unavailable. A "
              f"longer bar is worse. Levels with fewer than {MIN_N} "
              f"prescriptions are left out.",
              hit="with a missing item", base="prescriptions")
        + _missing_block())


# ---------------------------------------------------------------
# Analysis 7 opening: order fill rate overall, on the map, by level
# ---------------------------------------------------------------
# Quantity-weighted - delivered / requested over the kept medicine-months
# (_rq / _dv) - so the map is fed totals, not a count of rows.
# Same blue ramp as MAP 1 (tracer availability), its own breaks:
# near-quintiles of the district spread, 84.0%-100%.
FILL_STEPS = [(0.00, 0.88, "#d7e4f7"), (0.88, 0.92, "#9cc0ee"),
              (0.92, 0.95, "#5693e0"), (0.95, 0.98, "#2a68bd"),
              (0.98, 1.01, "#15407c")]
FILL_LABELS = ["below 88%", "88 to 91.9%", "92 to 94.9%", "95 to 97.9%",
               "98% and above"]
def _chw_geo(keyfn):
    """FIG 9.1 values {name: (n, hits, rate)} with NO minimum n - every
    district with a CHW record is coloured on the ramp (user, 2026-10-20)."""
    g = defaultdict(lambda: [0, 0])
    for r in CANS:
        g[keyfn(r)][0] += 1
        g[keyfn(r)][1] += C.yes(r[C.AVAIL])
    return {k: (n, h, pct(h, n)) for k, (n, h) in g.items()}


def _chw_lollipop():
    """FIG 9.2: availability per CHW product, highest first (ties by name),
    as the user's lollipop template (2026-10-20)."""
    g = defaultdict(lambda: [0, 0])
    for r in CANS:
        g[r["_short"]][0] += 1
        g[r["_short"]][1] += C.yes(r[C.AVAIL])
    rows = sorted(((k, n, h) for k, (n, h) in sorted(g.items())
                   if n >= MIN_N), key=lambda x: -(x[2] / x[1]))
    return ('<section class="ind"><h4><span class="mid">FIG 9.2</span>'
            'CHW medicine availability by product</h4>'
            + CH.lollipop(rows, hit="available", base="managed products",
                          show_n=True)
            + '</section>')


# FIG 9.1 CHW map: MAP 1's blue ramp, breaks on the CHW district spread
# (45.5%-100% at 2026-10-20)
CHW_STEPS = [(0.00, 0.65, "#d7e4f7"), (0.65, 0.72, "#9cc0ee"),
             (0.72, 0.80, "#5693e0"), (0.80, 0.90, "#2a68bd"),
             (0.90, 1.01, "#15407c")]
CHW_LABELS = ["below 65%", "65 to 71.9%", "72 to 79.9%", "80 to 89.9%",
              "90% and above"]


# FIG 4.4 classes: darker = MORE stock-outs (districts span 2.5-22.4%)
# Page 4 colours (user request 2026-10-19): the green / orange of the
# "requested / no request" bars across the whole page - green = no
# stock-out, orange = stock-out. The map and the overall rose use a green ->
# light -> orange ramp on the same classes; gauges and bars light the
# stock-out share in orange on a green track.
SO_GOOD, SO_BAD = "#1f9e7a", "#d9592a"
SO_STEPS = [(0.00, 0.05, "#1f9e7a"), (0.05, 0.08, "#86c4a8"),
            (0.08, 0.10, "#ecd3c1"), (0.10, 0.15, "#e98d5f"),
            (0.15, 1.01, "#d9592a")]
SO_DARK_TEXT = ("#86c4a8", "#ecd3c1")
SO_LABELS = ["below 5%", "5 to 7.9%", "8 to 9.9%", "10 to 14.9%",
             "15% and above"]


def _fill_vals(keyfn):
    g = defaultdict(lambda: [0.0, 0.0])
    for r in MG:
        g[keyfn(r)][0] += r["_rq"]
        g[keyfn(r)][1] += r["_dv"]
    return {k: (round(q), round(d), d / q) for k, (q, d) in g.items() if q}


def _fill_rows(keyfn, keys=None, drop=()):
    """(label, requested, delivered) bar rows for a quantity fill rate -
    calendar / level order when `keys` is given, else weakest first."""
    v = _fill_vals(keyfn)
    ks = [k for k in keys if k in v] if keys else \
        sorted(v, key=lambda k: v[k][2])
    return [[k, v[k][0], v[k][1]] for k in ks if k and k not in drop]


def _fill_class_fig():
    """Fill rate per drug class against the national fill rate, as a
    dumbbell. Classes with fewer than MIN_N managed pairs that requested
    anything are left out."""
    v = _fill_vals(_class_label)
    npairs = Counter(_class_label(r) for r in MG if r["_rq"] > 0)
    ref = pct(sum(r["_dv"] for r in MG), sum(r["_rq"] for r in MG))
    rows = sorted(([k, npairs[k], v[k][2]] for k in v
                   if k and npairs[k] >= MIN_N), key=lambda x: -x[2])
    return ('<section class="ind"><h4><span class="mid">FIG 7.4</span>'
            'Order fill rate by drug class</h4>'
            + CH.dumbbell(rows, ref, ref_label="national",
                          unit="pairs ordering")
            + '</section>')


# ---------------------------------------------------------------
# Requests and deliveries against consumption (AMC)
# ---------------------------------------------------------------
# Request vs AMC  = quantity requested / quantity consumed
# Delivery vs AMC = quantity delivered / quantity consumed
# over the SAME kept medicine-months as the fill rate (requested >=
# delivered), so the three measures on this page share one base. Over a
# three-month window this equals average monthly request / AMC. The form's
# own request_vs_amc_pct divides three months of requests by ONE month of
# consumption, which runs about three times too high, so it is not used.
# Pairs with nothing consumed in the kept months have no ratio and are left
# out (counted in the note).
def _amc_triplet(r):
    q = d = c = 0.0
    for m in MONTHS:
        rq = M.num(r[M.REQ[m]]) or 0
        dl = M.num(r[M.DEL[m]]) or 0
        if rq >= dl:
            q += rq
            d += dl
            c += M.num(r[M.CONS[m]]) or 0
    return q, d, c


for _r in MG:
    _r["_amcq"], _r["_amcd"], _r["_amcc"] = _amc_triplet(_r)
AMC_BASE = [r for r in MG if r["_amcc"] > 0]
_amc_nocons = sum(1 for r in MG if r["_amcc"] <= 0 and r["_amcq"] > 0)


def _amc_ratio(rs):
    c = sum(r["_amcc"] for r in rs)
    return (sum(r["_amcq"] for r in rs) / c if c else None,
            sum(r["_amcd"] for r in rs) / c if c else None)


def _amc_group(keyfn, keys=None, drop=()):
    g = defaultdict(list)
    for r in AMC_BASE:
        g[keyfn(r)].append(r)
    ks = [k for k in keys if k in g] if keys else sorted(g, key=str)
    out = []
    for k in ks:
        if not k or k in drop or len(g[k]) < MIN_N:
            continue
        a, b = _amc_ratio(g[k])
        out.append([k, len(g[k]), a, b])
    return out


AMC_BANDS = [("far below use (under 50%)", 0, 0.5),
             ("below use (50-89%)", 0.5, 0.9),
             ("in line with use (90-110%)", 0.9, 1.1000001),
             ("above use (111-150%)", 1.1000001, 1.5000001),
             ("far above use (over 150%)", 1.5000001, 1e9)]
AMC_BAND_COLS = ["#c9544a", "#eda100", "#199e70", "#5693e0", "#1c4f91"]


def _amc_bands(field):
    c = [0] * len(AMC_BANDS)
    for r in AMC_BASE:
        v = r[field] / r["_amcc"]
        for i, (_, lo, hi) in enumerate(AMC_BANDS):
            if lo <= v < hi:
                c[i] += 1
                break
    return c


def _amc_figs():
    a, b = _amc_ratio(AMC_BASE)
    rows = ([["All facilities", len(AMC_BASE), a, b], ["--Level of care"]]
            + _amc_group(lambda r: r["_lvl"], keys=LEVELS)
            + [["--Programme"]]
            + sorted(_amc_group(lambda r: s(r[M.PROGRAM]) or "(none)",
                                drop=("(none)",)), key=lambda x: -x[2]))
    pc = lambda v: f"{v * 100:.1f}%" if v is not None else ""
    cls = sorted(_amc_group(_class_label), key=lambda x: -x[2])
    tbl = html_table(
        ["Drug class", "Facility-medicine pairs", "Consumed",
         "Requested", "Delivered", "Request vs AMC", "Delivery vs AMC"],
        [[k, n, round(sum(r["_amcc"] for r in AMC_BASE
                          if _class_label(r) == k)),
          round(sum(r["_amcq"] for r in AMC_BASE if _class_label(r) == k)),
          round(sum(r["_amcd"] for r in AMC_BASE if _class_label(r) == k)),
          pc(x), pc(y)] for k, n, x, y in cls])
    return (
        '<section class="ind"><h4><span class="mid">FIG 7.9</span>'
        'Request vs AMC and delivery vs AMC</h4>'
        + CH.pairdots(rows, names=("Request vs AMC", "Delivery vs AMC"))
        + '</section>'
        '<section class="ind"><h4><span class="mid">FIG 7.10</span>'
        'How each facility\'s requests and deliveries compare with its own '
        'use</h4><p class="note">Every facility x medicine pair, banded by '
        'its own ratio. The pooled figures in FIG 7.8 sit near 100% because '
        'pairs ordering too little and pairs ordering too much cancel out - '
        'this is where they are.</p>'
        + CH.stacked([["Request vs AMC", _amc_bands("_amcq")],
                      ["Delivery vs AMC", _amc_bands("_amcd")]],
                     [b[0] for b in AMC_BANDS], colours=AMC_BAND_COLS)
        + '</section>'
        '<section class="ind"><h4><span class="mid">TABLE 7A</span>'
        'Request and delivery vs AMC by drug class</h4>'
        f'<p class="note">Highest request ratio first. Classes under {MIN_N} '
        'pairs are left out. Quantities are units of each class\'s own '
        'products and are not comparable across rows; the ratios are.</p>'
        + tbl + '</section>')


# One requisition line = a managed facility x medicine x month (June-
# August), read from the RAW requested / delivered fields. Lines with 0
# requested AND 0 delivered are left out (no order, no delivery).
# "Delivered without a request" (code 4) is off the figure (user, 2026-10-20)
FULFIL_SERIES = ["Fully delivered", "Partially delivered", "Not delivered"]
FULFIL_COLS = ["#1f9e7a", "#eda100", "#c9544a"]


def fulfil_cat(q, d):
    """0 left out, 1 fully (incl. more than requested), 2 partially,
    3 not at all, 4 delivered without a request."""
    if q <= 0:
        return 0 if d <= 0 else 4
    if d <= 0:
        return 3
    return 2 if d < q else 1


def _fulfil_cats(r):
    return [fulfil_cat(M.num(r[M.REQ[m]]) or 0, M.num(r[M.DEL[m]]) or 0)
            for m in MONTHS]


def fulfil_mix(rows):
    c = Counter(k for r in rows for k in _fulfil_cats(r))
    return [c[1], c[2], c[3]]


def _fill_figs():
    rq = sum(r["_rq"] for r in MG)
    dv = sum(r["_dv"] for r in MG)
    return (
        '<section class="ind"><h4><span class="mid">FIG 7.1</span>'
        'Order fill rate overall</h4>'
        '<p class="note">Quantity delivered against quantity requested, '
        'June to August, across every facility.</p>'
        + CH.dials([("All facilities", round(rq), round(dv))],
                   caption="of requested quantity delivered",
                   noun="units requested")
        .replace('class="dgrid"', 'class="dgrid solo"', 1)
        + '</section>'
        + two_view_map(
            "fillmap", "FIG 7.2", "Order fill rate by geography",
            "Quantity delivered against quantity requested, pooled by "
            "district or province - switch with View, narrow with the "
            "province filter, as on the tracer availability map. Hover a "
            "shape for the quantities.",
            MG, None, None, None, mfac, FILL_STEPS, FILL_LABELS,
            "Order fill rate", "requested units delivered",
            dvals=_fill_vals(lambda r: s(r["_p"][M.DIST])),
            pvals=_fill_vals(lambda r: s(r["_p"][M.PROV])),
            total=pct(dv, rq), n_label=f"{round(rq):,} units requested")
        + '<section class="ind"><h4><span class="mid">FIG 7.3</span>'
        'Order fill rate by level of care</h4>'
        + CH.columns(
            [[l, [(round(sum(_mreq(r, m) for r in MG if r["_lvl"] == l)),
                   round(sum(_mdel(r, m) for r in MG if r["_lvl"] == l)))
                  for m in MONTHS]]
             for l in LEVELS
             if sum(r["_rq"] for r in MG if r["_lvl"] == l) > 0],
            # pooled Jun-Aug average per level (user request 2026-10-20)
            list(MONTHS), colours=["#1c4f91", "#4f9be8", "#a9cfdf"],
            avg=True)
        + '</section>'
        + _fill_class_fig()
        + '<section class="ind"><h4><span class="mid">FIG 7.5</span>'
        'Order fill rate by programme</h4>'
        + CH.lollipop(sorted(_fill_rows(lambda r: s(r[M.PROGRAM]) or "(none)",
                                        drop=("(none)",)),
                             key=lambda x: -(x[2] / x[1])),
                      hit="delivered", base="units requested")
        + '</section>'
        + fig("FIG 7.6",
              [[mth, round(sum(_mreq(r, mth) for r in MG)),
                round(sum(_mdel(r, mth) for r in MG))] for mth in MONTHS],
              "Order fill rate by month",
              "Each month's delivered against requested, in calendar order.",
              hit="delivered", base="units requested")
        + fig("FIG 7.7",
              _fill_rows(mmgt, drop=("(not recorded)",)),
              "Order fill rate by management model",
              "Weakest first. Who runs the facility, as recorded on the stock "
              "form; the one facility with no management type is left out.",
              hit="delivered", base="units requested")
        # fully / partially / not at all (user request 2026-10-20)
        + '<section class="ind"><h4><span class="mid">FIG 7.8</span>'
        'Requisitions delivered fully, partially or not at all</h4>'
        + CH.stacked([["All facilities", fulfil_mix(MG)]]
                     + [[l, fulfil_mix([r for r in MG if r["_lvl"] == l])]
                        for l in LEVELS
                        if any(r["_lvl"] == l for r in MG)],
                     FULFIL_SERIES, colours=FULFIL_COLS)
        + '</section>'
        + _amc_figs())


PAGE_FIG = {
    2: lambda: (
        '<section class="ind"><h4><span class="mid">FIG 2.1</span>'
        'Stock adequacy overall</h4>'
        + '<div id="sx-f21">' + CH.radial(
            list(zip(ADEQ4, adequacy_mix4(), ADEQ4_BLURBS)),
            colours=ADEQ4_COLS,          # same key as the bars below
            centre_title="STOCK ADEQUACY",
            centre_sub=f"{_sa_n:,} medicines | out or on the shelf",
            ring_w=26, ring_gap=7, pie=True,   # user template, 2026-10-26
            highlights=[None, eop_line(sum(1 for r in SA_SHELF if _eop(r)),
                                       adequacy_mix(SA_SHELF)[0]),
                        None, None])
        + '</div></section>'
        '<section class="ind"><h4><span class="mid">FIG 2.2</span>'
        'Stock adequacy mix by level of care</h4>'
        + '<div id="sx-f22">'
        + CH.stacked(_mix4_rows(lambda r: r["_lvl"], LEVELS, 1), ADEQ4,
                     colours=ADEQ4_COLS)
        + '</div></section>'
        '<section class="ind"><h4><span class="mid">FIG 2.3</span>'
        'Stock adequacy mix by programme</h4>'
        '<p class="note">Vertical programmes manage their own ordering, so '
        'the mix differs sharply between them.</p>'
        + '<div id="sx-f23">' + CH.stacked(sorted(
            _mix4_rows(lambda r: s(r[M.PROGRAM]),
                       [p for p in sorted({s(r[M.PROGRAM])
                                           for r in SA_OUT + SA_SHELF}) if p],
                       MIN_N),
            key=lambda x: -x[1][3] / max(1, sum(x[1]))), ADEQ4,
            colours=ADEQ4_COLS)
        + '</div></section>'
        '<section class="ind"><h4><span class="mid">FIG 2.4</span>'
        'Stock adequacy mix by province</h4>'
        '<p class="note">Most overstocked province first.</p>'
        + '<div id="sx-f24">' + CH.stacked(sorted(
            _mix4_rows(lambda r: s(r["_p"][M.PROV]),
                       sorted({s(r["_p"][M.PROV])
                               for r in SA_OUT + SA_SHELF}), 1),
            key=lambda x: -x[1][3] / max(1, sum(x[1]))), ADEQ4,
            colours=ADEQ4_COLS)
        + '</div></section>'
        + _adequacy_med_tables()
        + '<section class="ind"><h4><span class="mid">FIG 2.5</span>'
        'Inventory record accuracy overall</h4>'
        '<p class="note">Left: physical count equals the system count. '
        'Right: quantity accuracy = (1 - |system - physical| / system) x '
        '100 per record, 0% at worst and 0% where the system quantity is 0, '
        'averaged. Both across every record that names a digital recording '
        'system.</p>'
        + '<div id="sx-f25">' + acc_dials(_acc)
        + '</div></section>'
        + acc_pair("FIG 2.6", "Inventory record accuracy by level of care",
                   lambda r: r["_lvl"], "sx-f26", "sx-q26",
                   keys=LEVELS, sort=False)
        + acc_pair("FIG 2.7", "Inventory record accuracy by programme",
                   lambda r: s(r[M.PROGRAM]) or "(none)", "sx-f28", "sx-q28")
        + acc_pair("FIG 2.8", "Inventory record accuracy by province",
                   lambda r: s(r["_p"][M.PROV]), "sx-f29", "sx-q29")
        + accuracy_map()
),
    4: lambda: (
        # overall, as the user's half-rose infographic (2026-10-19)
        '<section class="ind"><h4><span class="mid">FIG 4.1</span>'
        'Stock-out incidence overall</h4>'
        # one wedge per programme, worst first, coloured as the FIG 4.3 map
        + CH.half_rose(_worst_first(rate_by(MG, lambda r: s(r[M.PROGRAM])
                                            or "(none)", _somonths, mfac)),
                       fmt_pc(pct(len(_any_so), len(MG))),
                       f"{len(_any_so):,} of {len(MG):,} managed pairs",
                       colour_fn=lambda v: _step_colour(v, SO_STEPS),
                       dark_text=SO_DARK_TEXT)
        + '</section>'
        # gauge cards, one per level of care (user request 2026-10-19)
        '<section class="ind"><h4><span class="mid">FIG 4.2</span>'
        'Stock-out incidence by level of care</h4>'
        + CH.gauges(_bar_rows(rate_by(MG, lambda r: r["_lvl"], _somonths,
                                      mfac, keys=LEVELS, sort=False)),
                    caption="with a stock-out", on=SO_BAD, off=SO_GOOD)
        + '</section>'
        + two_view_map(
            "somap", "FIG 4.3", "Stock-out incidence by geography",
            "Share of managed facility x medicine pairs with at least one "
            "stock-out day, June-August, by district or province - switch "
            "with View, narrow with the province filter. Green is fewer "
            "stock-outs, orange more.",
            MG, lambda r: s(r["_p"][M.DIST]), lambda r: s(r["_p"][M.PROV]),
            _somonths, mfac, SO_STEPS, SO_LABELS, "Stock-out incidence",
            "managed pairs with a stock-out", unit="managed pairs")
        + _days_fig("FIG 4.4", lambda r: s(r[M.PROGRAM]) or "(none)",
                    "Average stock-out days by programme",
                    "Mean total_stockout_days per affected pair, longest "
                    "first.")
        + '<section class="ind"><h4><span class="mid">FIG 4.5</span>'
        'Stock-out months with and without a request, by level of care'
        '</h4>' + CH.stacked(soreq_rows(lambda r: r["_lvl"], keys=LEVELS),
                             SOREQ_SERIES, colours=SOREQ_COLS)
        + '</section>'
        '<section class="ind"><h4><span class="mid">FIG 4.6</span>'
        'Stock-out months with and without a request, by programme</h4>'
        + CH.stacked(soreq_rows(lambda r: s(r[M.PROGRAM])), SOREQ_SERIES,
                     colours=SOREQ_COLS)
        + '</section>'),
    5: _rx_figs,
    7: lambda: _fill_figs(),
    # opens on the district / province map, as MAP 1 (user request 2026-10-20)
    9: lambda: two_view_map(
        "chwmap", "FIG 9.1", "CHW medicine availability by geography", "",
        CANS, lambda r: C.s(r["_p"][C.DIST]), lambda r: C.s(r["_p"][C.PROV]),
        lambda r: C.yes(r[C.AVAIL]), cfac, CHW_STEPS, CHW_LABELS,
        "CHW availability", "managed products available",
        unit="managed products",
        dvals=_chw_geo(lambda r: C.s(r["_p"][C.DIST])),
        pvals=_chw_geo(lambda r: C.s(r["_p"][C.PROV]))) + _chw_lollipop(),
}


def check_scripts(doc):
    """Fail the build on a JavaScript string literal broken across a line.

    The dashboard's behaviour - the filters, the map redraws, every chart
    rebuild - lives in one inline <script>. A single unterminated string is a
    parse error that kills the whole block, and nothing about the page looks
    wrong afterwards: the build-time SVGs are still there, so the charts and
    maps render and only the interactivity is dead. That is exactly the kind
    of failure a human reviewer does not catch.
    //
    It happened: writing '\n' in the Python source of the Sankey tooltip
    emitted a REAL newline into the JavaScript instead of the two characters
    backslash-n, and the filter bar silently stopped working.
    //
    The check walks each script line and tracks whether a quote opened and
    never closed on that line, skipping escaped quotes and // comments.
    """
    bad = []
    for blk in re.findall(r"<script>(.*?)</script>", doc, re.S):
        for ln, line in enumerate(blk.split("\n"), start=1):
            code = line
            q = None
            i = 0
            prev = ""
            while i < len(code):
                c = code[i]
                if q:
                    if c == "\\":
                        i += 2
                        continue
                    if c == q:
                        q = None
                elif c in "'\"":
                    q = c
                elif c == "/" and i + 1 < len(code):
                    # Either a // comment - nothing after it matters - or a
                    # regex literal, whose quotes and slashes are not string
                    # delimiters. /[&<>"]/g reads as an opened quote without
                    # this, which is a false alarm on the one line that
                    # escapes HTML.
                    if code[i + 1] == "/":
                        break
                    if prev in "" or prev in "(,=:[!&|?{};+return":
                        j = i + 1
                        while j < len(code):
                            if code[j] == "\\":
                                j += 2
                                continue
                            if code[j] == "/":
                                break
                            j += 1
                        if j < len(code):
                            i = j + 1
                            prev = "/"
                            continue
                if not c.isspace():
                    prev = c
                i += 1
            if q:
                bad.append((ln, line.strip()[:90]))
    if bad:
        raise SystemExit(
            "BUILD FAILED - unterminated JavaScript string literal:\n"
            + "\n".join(f"  script line {ln}: {t}" for ln, t in bad)
            + "\n  A '\\n' in the Python source emits a real newline. "
              "Write '\\\\n' to send backslash-n to the browser.")


# ---------------------------------------------------------------
# how every figure is calculated
# ---------------------------------------------------------------
# The counts below are read from the live row sets, not typed in, so this
# page cannot drift from the figures it explains: if a filter changes, the
# number here changes with it. Anything hand-written is a definition or a
# reason, never a quantity.
def _n(x):
    return f"{len(x):,}"


def method_rows():
    """(block, item, formula or filter, value) for the Method page."""
    out = []

    def add_m(block, item, how, val=""):
        out.append([block, item, how, val])

    add_m("1. What one row is", "Stock assessment, sheet 2",
          "one row per facility x medicine assessed", _n(MED_ALL))
    add_m("1. What one row is", "Facilities visited",
          "distinct facility names in the submissions", f"{len(M.FACILITIES)}")
    add_m("1. What one row is", "Molecules",
          "medicines after strengths of one molecule are collapsed (Method 3)",
          f"{len({s(r[M.MED]) for r in MOLS})}")
    add_m("1. What one row is", "Items before collapsing",
          "medicines as the form lists them",
          f"{len({s(r[M.MED]) for r in MED_ALL})}")

    add_m("2. Cleaning applied before counting",
          "S2 duplicate rows removed",
          "same facility x medicine recorded twice; the row with more fields "
          "filled is kept", f"{len(M.DUPES_DROPPED)}")
    add_m("2. Cleaning applied before counting",
          "S6 EML applicability corrected",
          "form said No where the official EML list holds the medicine at "
          "that level of care -> Yes", f"{len(M.EML_FIXED)}")
    add_m("2. Cleaning applied before counting",
          "S6 not changed (opposite direction)",
          "form says Yes where the list does not hold it at that level - "
          "reported, left as recorded", f"{len(M.EML_OPPOSITE)}")
    add_m("2. Cleaning applied before counting",
          "S7 availability corrected",
          "answered 'available = Yes' with 0 units on the shelf -> No",
          f"{len(M.ZERO_FIXED)}")
    add_m("2. Cleaning applied before counting",
          "S8 recording system names normalised",
          "free text; one spelling per system, bare numbers removed",
          f"{len(M.EMR_CHANGED)}")

    add_m("3. Method 3 - combining strengths", "The rule",
          "per facility: a molecule is MANAGED if any strength is managed, "
          "and AVAILABLE if any managed strength was on the shelf", "")
    add_m("3. Method 3 - combining strengths", "Why",
          "two strengths are interchangeable - a facility holding Metformin "
          "500mg can treat a patient needing metformin", "")
    add_m("3. Method 3 - combining strengths", "What does NOT collapse",
          "quantities, months of stock, stock-out days, requisitions and "
          "financial loss stay at item level - summing 500mg and 850mg "
          "tablets gives a meaningless number", "")

    add_m("4. The availability denominator", "EXPECTED (the denominator)",
          "set by medicine_program - two tracks, below", _n(APPM))
    add_m("4. The availability denominator",
          "Essential Medicines track",
          "every molecule the EML authorises at that facility's level of "
          "care, managed or not",
          f"{sum(1 for r in MOLS if r['_track'] == 'EML'):,}")
    add_m("4. The availability denominator",
          "Programme track (Malaria, MCCH, FP, ARV, TB, Consumables)",
          "only the molecules the facility manages - programme products are "
          "allocated, so not holding one is not a shortage",
          f"{sum(1 for r in MOLS if r['_track'] == 'managed'):,}")
    add_m("4. The availability denominator", "of which MANAGED",
          "the facility actually stocks it", _n(MGM))
    add_m("4. The availability denominator",
          "of which NOT managed (Essential Medicines only)",
          "authorised but the facility does not stock it - counted "
          "UNAVAILABLE under the study definition", _n(GAPM))
    add_m("4. The availability denominator", "AVAILABLE (the numerator)",
          "managed AND on the shelf on the visit day", _n(AVM))
    add_m("4. The availability denominator", "Managed but empty",
          "stocked, but nothing on the shelf", f"{len(SOM) - len(GAPM):,}")
    add_m("4. The availability denominator", "OVERALL AVAILABILITY",
          "AVAILABLE / EXPECTED", f"{pct(len(AVM), len(APPM)):.1%}")
    add_m("4. The availability denominator",
          "Memo: over managed rows only",
          "AVAILABLE / MANAGED - the figure if a medicine the facility "
          "chooses not to stock is left out", f"{pct(len(AVM), len(MGM)):.1%}")

    add_m("5. Other denominators", "Inventory accuracy",
          "physical count = system count, over AVAILABLE rows only - a count "
          "exists only where there is stock; facilities that named no digital "
          "system are excluded", _n(_acc))
    add_m("5. Other denominators", "Stock-out incidence",
          "pairs with at least one stock-out day June-August, over MANAGED "
          "pairs", _n(MG))
    add_m("5. Other denominators", "Stock-out days",
          "summed from the three monthly fields, NOT the form's "
          "total_stockout_days, which adds the ongoing current run and "
          "exceeds the 92-day window on 355 rows", "")
    add_m("5. Other denominators", "Order fill rate",
          "quantity delivered / quantity requested, Jun-Aug, over the "
          "medicine-months where requested >= delivered (over-deliveries "
          "dropped)", f"{_over_lines} months dropped")
    add_m("5. Other denominators", "CHW availability",
          "available / (managed and answered); 'not managed' is MFT zoning, "
          "not a stock-out, so it stays out of the denominator", _n(CANS))
    add_m("5. Other denominators", "Prescription fulfilment",
          "items recorded available / items with an answer - a design "
          "artefact of who gets recorded; compare groups, not the level",
          _n([r for r in ITEMS if r["_avail"] is not None]))

    add_m("6. Suppression", f"Minimum {MIN_N} observations",
          "a rate computed on fewer is left blank rather than ranked; the "
          "counts are still shown", "")
    add_m("6. Suppression", f"Minimum {MIN_FAC} per facility",
          "for the facility dots in FIG 5 - a facility basket of 16 "
          "medicines is a usable rate, 30 would drop 15 of 26 health posts",
          "")
    return out


def method_block():
    rows = method_rows()
    blocks, seen = [], None
    out = []
    for b, item, how, val in rows:
        if b != seen:
            if seen is not None:
                out.append("</tbody></table></div></section>")
            out.append(
                f'<section class="ind"><h4><span class="mid">METHOD</span>'
                f'{esc(b)}</h4><div class="scroll"><table><thead><tr>'
                f'<th>Measure</th><th>How it is calculated</th>'
                f'<th>Value</th></tr></thead><tbody>')
            seen = b
        out.append(f'<tr><td>{esc(item)}</td><td>{esc(how)}</td>'
                   f'<td class="num">{esc(val)}</td></tr>')
    out.append("</tbody></table></div></section>")
    return "".join(out)


# One line under every figure title saying how its number is built:
# numerator / denominator and anything excluded. Kept here in one table,
# keyed by the figure id, so a figure's wording and its method line cannot
# be edited apart. A figure with no entry gets no line, and the build prints
# a warning naming it.
FIG_CALC = {
    # a (method 1, method 2) pair: the method switch shows one of the two
    "MAP 1": ("Method 1 - Expected medicines: available on the visit day / "
              "expected medicines, pooled over every facility in the "
              "district. Expected = for Essential Medicines, EML-authorised "
              "at the level (not managed counts as unavailable); for other "
              "programmes, managed only.",
              "Method 2 - Managed products only: available on the visit day "
              "/ medicines the facility manages, pooled over every facility "
              "in the district. Every programme, Essential Medicines "
              "included, is on the managed base: a medicine the facility "
              "does not manage is left out, not counted unavailable."),
    "MAP 2": "Per facility: medicines available on the visit day / its "
             "expected medicines (EML-authorised Essential Medicines + "
             "managed programme products).",
    "FIG 1": "Available on the visit day / expected pairs, by programme "
             "(medicine_program).",
    "FIG 2": "Available on the visit day / expected pairs (EML-authorised "
             "Essential Medicines + managed programme products), by level of "
             "care.",
    "FIG 3": "Each cell: available / expected pairs for that programme at "
             "that level of care; last column all levels.",
    "FIG 4": "Available / expected pairs, by management type (diamond); "
             "each dot is one facility's own rate.",
    "FIG 5": "Available on the visit day / expected pairs, per programme "
             "within each province; last column all provinces.",
    "FIG 6": "Per hospital (district, Level 2 teaching and national "
             "referral): available on the visit day / expected medicines; "
             "lowest first.",
    "FIG 7": "Unavailable rows answered 'Yes' at RMS Regional (or Central) / "
             "unavailable rows with a Yes or No answer for BOTH stores.",
    "FIG 8": "Same base as FIG 7 - answered 'Yes' at the store / unavailable "
              "rows with a Yes or No answer for BOTH stores - by programme.",
    "FIG 9": "Same base as FIG 7. Rows classed Distribution Gap (or "
              "nationwide stock-out) "
              "/ unavailable rows with a Yes or No answer for BOTH stores, by "
              "level of care.",
    "FIG 10": "Same base as FIG 7. Rows classed Distribution Gap (or "
              "'nationwide stock-out') / unavailable rows with a Yes or No "
              "answer for BOTH stores, by programme.",
    "TABLE": "Per molecule: facilities with it available / facilities "
             "expected to hold it - EML-authorised for Essential Medicines, "
             "managing it for other programmes.",
    "FIG 2.1": "Base = every stock-out (not on the shelf on the visit day). "
               "Under 1 = understock, 1-2 = adequate, over 2 = overstock "
               "(at teaching and referral hospitals 1-3 = adequate, over 3 = overstock).",
    "FIG 2.2": "As FIG 2.1 - stocked out, or months of stock classed "
               "under / adequate / over - shares within each level of care.",
    "FIG 2.3": "As FIG 2.1 - stocked out, or months of stock classed "
               "under / adequate / over - shares within each programme.",
    "FIG 2.4": "As FIG 2.1 - stocked out, or months of stock classed "
               "under / adequate / over - shares within each province.",
    "TABLE 2A": "Per medicine: rows with under 1 month of stock / rows in "
                "stock with consumption > 0. Medicines under 30 rows left "
                "out.",
    "TABLE 2B": "Per medicine: rows with over 2 months of stock (3 at "
                "teaching and referral hospitals) / rows in "
                "stock with consumption > 0. Medicines under 30 rows left "
                "out.",
    "FIG 2.5": "Left: records where physical count = system count / records "
               "with both counts that name a digital recording system. "
               "Right: mean over the same records of max(0, 1 - |system - "
               "physical| / system) x 100, with 0% where the system "
               "quantity is 0.",
    "FIG 2.6": "Left: physical = system / records with both counts and a named "
               "digital system, by level of care."
               " Right: mean of max(0, 1 - |system - physical| / system), 0% where the system quantity is 0, in the same order as the left.",
    "FIG 2.7": "Left: physical = system / records with both counts and a named "
               "digital system, by programme."
               " Right: mean of max(0, 1 - |system - physical| / system), 0% where the system quantity is 0, in the same order as the left.",
    "FIG 2.8": "Left: physical = system / records with both counts and a named "
               "digital system, by province."
               " Right: mean of max(0, 1 - |system - physical| / system), 0% where the system quantity is 0, in the same order as the left.",
    "FIG 2.9": "Left: physical = system / records with both counts and a named "
               "digital system, pooled by district."
               " Right: mean of max(0, 1 - |system - physical| / system), 0% where the system quantity is 0, in the same order as the left.",
    "FIG 4.1": "Managed rows with at least one stock-out day, Jun-Aug / all "
               "managed rows: the centre is all facilities, each wedge one "
               "programme (its length against the highest; colour as FIG "
               "4.3).",
    "FIG 4.3": "Managed rows with at least one stock-out day, Jun-Aug / all "
               "managed rows, pooled by district or province.",
    "FIG 4.2": "Managed rows with at least one stock-out day, Jun-Aug / all "
               "managed rows, by level of care.",
    "FIG 4.5": "Managed medicine-months (Jun-Aug) with at least one stock-out "
               "day, split by whether a quantity above 0 was requested in "
               "the same month; shares within each level of care.",
    "FIG 4.6": "As FIG 4.5, within each programme; most unrequested first.",
    "FIG 4.4": "Total stock-out days (total_stockout_days) per managed "
               "pair; mean over the pairs with at least one day out, by "
               "programme.",
    "FIG 5.1": "Prescriptions by outcome / all prescriptions: complete = "
               "every item available; partially served = some missing; "
               "nothing served = all missing.",
    "FIG 5.2": "Prescriptions with every item available / all prescriptions, "
               "pooled by province and by district.",
    "FIG 5.3": "Per programme: prescriptions where one of its medicines was "
               "unavailable / prescriptions holding at least one of its "
               "medicines.",
    "FIG 5.4": "Prescriptions with at least one item unavailable / all "
               "prescriptions, by level of care.",
    "FIG 5.5": "Unserved medicines / all prescribed medicines.",
    "FIG 5.6": "Unserved medicines / medicines prescribed, by programme derived from "
               "the EML section and medicine name.",
    "TABLE 5A": "Per medicine: times unserved / times prescribed; "
                "prescriptions, facilities and districts where it was "
                "unserved.",
    "FIG 7.1": "Quantity delivered / quantity requested, June-August, over "
               "medicine-months where requested >= delivered, all "
               "facilities.",
    "FIG 7.8": "Requisition lines (facility x medicine x month, Jun-Aug) "
               "by outcome: delivered >= requested, 0 < delivered < "
               "requested, delivered 0, or delivered with 0 requested / all "
               "lines except those with 0 requested and 0 delivered.",
    "FIG 7.2": "Quantity delivered / quantity requested over the kept "
               "medicine-months, pooled by province and by district.",
    "FIG 7.3": "Quantity delivered / quantity requested in each month, "
               "medicine-months where requested >= delivered, by level of "
               "care.",
    "FIG 7.4": "Quantity delivered / quantity requested over the kept "
               "medicine-months, by drug class; white dot = the same ratio "
               "for all classes.",
    "FIG 7.5": "Quantity delivered / quantity requested over the kept "
               "medicine-months, by programme.",
    "FIG 7.6": "Quantity delivered / quantity requested in each month, "
               "months where requested >= delivered only.",
    "FIG 7.7": "Quantity delivered / quantity requested over the kept "
               "medicine-months, by management model.",
    "FIG 7.9": "Requested / consumed and delivered / consumed, quantities "
               "pooled over the kept medicine-months, by group.",
    "FIG 7.10": "Each pair's requested / consumed and delivered / consumed, "
               "banded: <50%, 50-89%, 90-110%, 111-150%, >150%.",
    "TABLE 7A": "Requested / consumed and delivered / consumed, pooled by "
                "drug class over the kept medicine-months.",
    "FIG 9.1": "CHW products available on the visit day / products the CHW "
               "manages and answered for, by district or province.",
    "FIG 9.2": "CHW products available on the visit day / products the CHW "
               "manages and answered for, by product.",
}


def add_calc(doc):
    """Put each figure's method line directly under its title."""
    seen = set()

    def ins(m):
        fid = m.group(2)
        line = FIG_CALC.get(fid)
        if not line:
            return m.group(0)
        seen.add(fid)
        if isinstance(line, tuple):
            return (m.group(0) + "".join(
                f'<div class="only-{k}"><p class="calc"><b>How it is '
                f'calculated:</b> {esc(x)}</p></div>'
                for k, x in zip("ab", line)))
        return (m.group(0) + f'<p class="calc"><b>How it is calculated:</b> '
                f'{esc(line)}</p>')
    doc = re.sub(r'(<span class="mid">([^<]+)</span>[^<]*(?:<span '
                 r'class="badge">[^<]*</span>)?</h4>)', ins, doc)
    titled = set(re.findall(r'<span class="mid">((?:FIG|MAP|TABLE)[^<]*)'
                            r'</span>', doc))
    for fid in sorted(titled - seen):
        print(f"  WARNING: no method line for {fid}")
    for fid in sorted(set(FIG_CALC) - seen):
        print(f"  WARNING: method line for {fid} matched no figure")
    return doc


# Tables the dashboard leaves out - 2.5-2.8 because a figure on the same page
# already shows them, 1.1, 1.4, 1.5-1.8, 3.4-3.9, 4.2-4.6, 7.1-7.7 and 7.9-7.10 on
# request.
# The workbook keeps every table - it is the full record.
HTML_DROP = {(8, "8.1-7.3"), (5, "5.1-4.5"), (5, "5.12"), (5, "5.7"),
             (9, "9.1"), (9, "9.2"), (9, "9.3"), (9, "9.5"), (9, "9.6"),
             (1, "1.1"), (1, "1.4"), (1, "1.5"), (1, "1.6"), (1, "1.7"), (1, "1.8"),
             (2, "2.1"), (2, "2.2-2.4"),
             (2, "2.5"), (2, "2.6"), (2, "2.7"), (2, "2.8"),
             (3, "3.2"), (3, "3.3"), (3, "3.4"), (3, "3.5"), (3, "3.6"), (3, "3.7"), (3, "3.8"),
             (3, "3.9"), (4, "4.1"), (4, "4.2"), (4, "4.3-3.6"),
             (4, "4.4"), (4, "4.5"), (4, "4.6"),
             (7, "7.1"), (7, "7.2"), (7, "7.2-6.5"), (7, "7.6"), (7, "7.7"),
             (7, "7.9"), (7, "7.10")}


# Tables shown on the dashboard WITHOUT their note, on request; the
# workbook keeps the note.
HTML_NO_NOTE = {(5, "5.13"), (1, "1.10"), (1, "1.9"), (4, "4.7"), (3, "3.1c"), (3, "3.2p")}

# Whole analyses left off the dashboard on request - their tab and page.
# The workbook keeps their sheets.
HTML_PAGES_DROP = {6}
PAGES = {a: n for a, n in ANALYSES.items() if a not in HTML_PAGES_DROP}


# ---------------------------------------------------------------
# Page scorecards: two to four headline findings at the top of each page
# ---------------------------------------------------------------
# Every figure is computed from the same row sets as the page below it, so
# a refresh moves the cards with the charts. Each card: the number, what it
# is, and one line of context. The availability page carries one set per
# method; its method switch shows the matching set.
def _rate_of(rows):
    return sum(1 for r in rows if M.yes(r[M.AVAIL])) / len(rows) if rows \
        else None


def _low_high(rows, keyfn, min_n=MIN_N, min_fac=1):
    g = defaultdict(list)
    for r in rows:
        g[keyfn(r)].append(r)
    rs = sorted((_rate_of(v), k, len(v)) for k, v in g.items()
                if k and len(v) >= min_n
                and len({s(x["_p"][M.FAC]) for x in v}) >= min_fac)
    return rs[0], rs[-1]


def _card(v, label, detail, cid=None):
    """cid: ids "<cid>-v", "<cid>-l", "<cid>-s" on the value, label and
    detail, so the page's filter script can rewrite them. Every scorecard
    follows the filters (user request 2026-10-18)."""
    iv, il, isd = ((f' id="{cid}-v"', f' id="{cid}-l"', f' id="{cid}-s"')
                   if cid else ("", "", ""))
    return (f'<div class="stat"><p class="v"{iv}>{esc(v)}</p>'
            f'<p class="l"{il}>{esc(label)}</p><p class="s"{isd}>'
            f'{esc(detail)}</p></div>')


def _avail_cards(base, other, basis, other_name, cid="pc1"):
    """Availability-page cards on one method's base; `other` is the other
    method's base, quoted on the headline card for comparison. Levels need
    5+ facilities, so a single provincial hospital is not 'the' level."""
    pc = lambda x: f"{x * 100:.1f}%"
    lo_lv, hi_lv = _low_high(base, lambda r: r["_lvl"], min_fac=5)
    lo_pg, hi_pg = _low_high(base, lambda r: s(r[M.PROGRAM]))
    lo_d, hi_d = _low_high(base, lambda r: s(r["_p"][M.DIST]))
    return [
        _card(pc(_rate_of(base)), f"availability - {basis}",
              f"{pc(_rate_of(other))} on {other_name}", cid=f"{cid}-0"),
        _card(pc(lo_lv[0]), f"{lo_lv[1].lower()}s - lowest level",
              f"{hi_lv[1].lower()}s highest at {pc(hi_lv[0])}", cid=f"{cid}-1"),
        _card(pc(lo_pg[0]), f"{lo_pg[1]} - lowest programme",
              f"{hi_pg[1]} highest at {pc(hi_pg[0])}", cid=f"{cid}-2"),
        _card(pc(lo_d[0]), f"{lo_d[1]} - lowest district",
              f"{hi_d[1]} highest at {pc(hi_d[0])}", cid=f"{cid}-3")]


def page_cards():
    pc = lambda x: f"{x * 100:.1f}%"
    cards = {}
    method_cards = {
        "a": _avail_cards(MOLS, MOLS_B, "expected medicines",
                          "managed products only", cid="pc1a"),
        "b": _avail_cards(MOLS_B, MOLS, "managed products only",
                          "expected medicines", cid="pc1b")}
    # 2. stock adequacy and inventory accuracy - the adequacy cards are on
    # FIG 2.1's base: every stock-out + every stocked row with consumption
    mix = adequacy_mix4()
    tot = sum(mix) or 1
    n_eop = sum(1 for r in SA_SHELF if _eop(r))
    acc_d = rate_by(_acc, lambda r: s(r["_p"][M.DIST]), _exact, mfac)
    acc_d = sorted([x for x in acc_d if x[4] is not None], key=lambda x: x[4])
    cards[2] = [
        _card(pc(mix[0] / tot), "stocked out",
              f"{mix[0]:,} of {tot:,} medicines not on the shelf on the "
              f"visit day", cid="pc2-0"),
        _card(pc(mix[1] / tot), "understocked",
              f"under one month; {n_eop:,} ({n_eop / max(1, mix[1]):.1%}) at "
              f"or below {EOP} months - emergency order point", cid="pc2-1"),
        _card(pc(mix[3] / tot), "overstocked",
              f"more than two months (three at teaching and referral hospitals); only "
              f"{pc(mix[2] / tot)} adequate", cid="pc2-2"),
        _card(pc(pct(sum(1 for r in _acc if _exact(r)), len(_acc))),
              "inventory record accuracy",
              f"physical count equals the system count; {acc_d[0][0]} least "
              f"accurate at {pc(acc_d[0][4])}", cid="pc2-3")]
    # 3. consumption
    wide = [x for x in AMC_STATS if x[2] >= 10]
    top = max(wide, key=lambda x: x[3])
    erratic = pct(sum(1 for _, c in _cvok if c >= 50), len(_cvok))
    hc = pct(sum(r["_ctot"] for r in CROWS if r["_lvl"] == "Health centre"),
             sum(r["_ctot"] for r in CROWS))
    cards[3] = [
        _card(f"{top[3]:,.0f}", f"{top[0][:30].title()} - highest total AMC",
              f"units a month, summed across {top[2]} facilities", cid="pc3-0"),
        _card(pc(erratic), "of facility x medicine pairs erratic",
              "month-to-month variation of 50% or more", cid="pc3-1"),
        _card(pc(hc), "of all consumption at health centres",
              "June to August, all medicines", cid="pc3-2")]
    # 4. stock-outs
    inc = pct(sum(1 for r in MG if _somonths(r)), len(MG))
    pg = sorted([x for x in rate_by(MG, lambda r: s(r[M.PROGRAM]),
                                    _somonths, mfac)
                 if x[0] and x[4] is not None], key=lambda x: -x[4])
    dd = defaultdict(list)
    for r in MG:
        if _tdays(r) > 0:
            dd[s(r[M.PROGRAM])].append(_tdays(r))
    longest = max(((statistics.mean(v), k) for k, v in dd.items()
                   if k and len(v) >= MIN_N))
    cards[4] = [
        _card(pc(inc), "of managed medicines stocked out",
              "at least one day, June to August", cid="pc4-0"),
        _card(pc(pg[0][4]), f"{pg[0][0]} - most frequent stock-outs",
              f"{pg[-1][0]} least at {pc(pg[-1][4])}", cid="pc4-1"),
        _card(f"{longest[0]:.0f} days", f"{longest[1]} - longest stock-outs",
              "mean days out per medicine that ran out", cid="pc4-2")]
    # 5. prescriptions - high by design, so every card says so
    nrh = [r for r in RX_PRES if rx_level(r) == "National referral hospital"]
    rd = sorted([x for x in rate_by(RX_PRES, lambda r: s(r["_sub"][P.DISTR]),
                                    _rx_complete, pfac) if x[4] is not None],
                key=lambda x: x[4])
    cards[5] = [
        _card(pc(pct(sum(1 for r in RX_PRES if r["_fully_served"]),
                     len(RX_PRES))), "of prescriptions complete",
              "high by design - prescribers write what the pharmacy holds; "
              "not a stock-availability rate", cid="pc5-0"),
        _card(f"{len(RX_MISS) / len(RX_ANS) * 100:.1f}%",
              "of prescribed medicines unserved",
              f"{len(RX_MISS):,} of {len(RX_ANS):,} items - understated for "
              "the same reason; use the stock pages for availability", cid="pc5-1"),
        _card(pc(pct(sum(1 for r in nrh if r["_any_unserved"]), len(nrh))),
              "national referral prescriptions with a gap",
              f"{len(nrh)} prescriptions", cid="pc5-2"),
        _card(pc(rd[0][4]), f"{rd[0][0]} - lowest district",
              "share of prescriptions complete", cid="pc5-3")]
    # 6 (analysis 7). requisition and order fulfilment
    fv = _fill_vals(lambda r: s(r[M.PROGRAM]) or "(none)")
    lowp = min(((v[2], k) for k, v in fv.items() if k != "(none)"))
    q, d, c = (sum(r["_amcq"] for r in AMC_BASE),
               sum(r["_amcd"] for r in AMC_BASE),
               sum(r["_amcc"] for r in AMC_BASE))
    bands = _amc_bands("_amcq")
    cards[7] = [
        _card(pc(pct(sum(r["_dv"] for r in MG), sum(r["_rq"] for r in MG))),
              "order fill rate", "quantity delivered / requested, Jun-Aug", cid="pc7-0"),
        _card(pc(lowp[0]), f"{lowp[1]} - lowest fill rate",
              "of all programmes", cid="pc7-1"),
        _card(pc(q / c), "requested against consumption",
              f"delivered {pc(d / c)} of consumption", cid="pc7-2"),
        _card(pc(bands[2] / sum(bands)), "of medicine orders in line with use",
              "facility x medicine requests within 10% of consumption", cid="pc7-3")]
    # 7 (analysis 8). financial value
    # these three follow the page filters: pf_a8.py rewrites them
    cards[8] = [
        _card(f"{round(_val):,} RWF", "unserved prescription value",
              f"{len(UNSERVED):,} unserved items, the form's own cost",
              cid="pc8-val"),
        _card(f"{round(_loss_tot / 1e6, 1)}M RWF",
              "estimated sales lost to stock-outs",
              "tariff price + 20% markup, priced products", cid="pc8-loss"),
        _card(pc(_loss_top[1][1] * (1 + MARKUP) / _loss_tot),
              "of lost sales from one item",
              f"{s(_loss_top[0][M.MED])[:28]} at "
              f"{s(_loss_top[0]['_p'][M.FAC])}", cid="pc8-top")]
    # 8 (analysis 9). CHWs - availability over MANAGED rows (MFT zoning)
    cm = defaultdict(lambda: [0, 0])
    for r in CANS:
        g = cm[C.s(r[C.MED])[:30]]
        g[0] += 1
        g[1] += C.yes(r[C.AVAIL])
    lowc = min(((v[1] / v[0], k) for k, v in cm.items() if v[0] >= MIN_N))
    cards[9] = [
        _card(pc(pct(len(CAV), len(CANS))), "CHW medicine availability",
              f"{len(C.CHWS)} CHWs, products they manage", cid="pc9-0"),
        _card(pc(lowc[0]), f"{lowc[1]} - least available",
              "at CHW level", cid="pc9-1"),
        _card(f"{len(C.FACILITIES)}", "health centres supervising the CHWs",
              "every CHW reports to a health centre", cid="pc9-2")]
    out = {a: '<div class="stats pcards">' + "".join(v) + "</div>"
           for a, v in cards.items()}
    out[1] = (f'<div class="only-a"><div class="stats pcards">'
              f'{"".join(method_cards["a"])}</div></div>'
              f'<div class="only-b"><div class="stats pcards">'
              f'{"".join(method_cards["b"])}</div></div>')
    return out


def renumber_display(doc):
    """Number what the dashboard SHOWS in sequence, after pages and tables
    were hidden: tabs and page headings 1..n, each page's figures under its
    own tab number, each page's tables n.1, n.2, ... in order of appearance.
    The workbook keeps the indicator ids; this touches only the page.

    Only labels are rewritten - the <span class="mid"> markers, the tab and
    heading text - plus table cross-references inside notes. Numbers in
    table cells are data and are never matched."""
    newno = {a: i for i, a in enumerate(PAGES, start=1)}
    for a, n in newno.items():
        name = PAGES[a]
        doc = doc.replace(f'data-page="a{a}">{a}. {esc(name)}<',
                          f'data-page="a{a}">{n}. {esc(name)}<')
        doc = doc.replace(f'<h2>Analysis {a} &middot; {esc(name)}</h2>',
                          f'<h2>Analysis {n} &middot; {esc(name)}</h2>')
    pieces = re.split(r'(<div class="page[^"]*" id="a\d+">)', doc)
    out = [pieces[0]]
    for marker, body in zip(pieces[1::2], pieces[2::2]):
        a = int(re.search(r'id="a(\d+)"', marker).group(1))
        n = newno.get(a, a)
        if n != a:
            body = re.sub(rf'(FIG |TABLE ){a}(\.|[A-Z]\b)',
                          lambda m: f"{m.group(1)}{n}{m.group(2)}", body)
        mp, k = {}, 0
        for mid in re.findall(r'<span class="mid">([^<]+)</span>', body):
            if mid.startswith(("FIG", "MAP", "TABLE")) or mid in mp:
                continue
            k += 1
            mp[mid] = f"{n}.{k}"
        body = re.sub(r'<span class="mid">([^<]+)</span>',
                      lambda m: f'<span class="mid">{mp.get(m.group(1), m.group(1))}</span>',
                      body)

        def fix_note(m):
            txt = m.group(0)
            for old in sorted(mp, key=len, reverse=True):
                txt = re.sub(rf'(?<![\d.,]){re.escape(old)}(?![\d%])',
                             mp[old], txt)
            return txt
        body = re.sub(r'<p class="note">.*?</p>', fix_note, body, flags=re.S)
        out += [marker, body]
    return "".join(out)


def _sx_fac(r):
    p = r["_p"]
    return [s(p[M.FAC]), s(p[M.DIST]), s(p[M.PROV]), r["_lvl"], mmgt(r)]


def _sx_med(r):
    return [s(r[M.MED]), s(r[M.PROGRAM]), _class_label(r)]


def _sx_cfg():
    """Scales and labels the page-2 filter script redraws with, read from
    the same constants the build-time figures use."""
    return {"min_n": MIN_N, "levels": LEVELS, "series": ADEQ4,
            "colours": ADEQ4_COLS, "blurbs": ADEQ4_BLURBS,
            "ring_w": 26, "ring_gap": 7,
            "pie_reach": CH.PIE_REACH, "pie_hub": CH.PIE_HUB,
            "pair_w": PAIR_W, "pair_lw": PAIR_LW,
            "eop": EOP, "hl": CH.RADIAL_HL,
            "acc_steps": [list(x) for x in ACC_STEPS], "acc_none": "#2c2c2a",
            "rest": CH.REST, "avail": CH.AVAIL, "g_on": CH.G_ON,
            "g_off": CH.G_OFF}


# ---------------------------------------------------------------
# page filters: pf_a<N>.py builds each page's datasets and script
# ---------------------------------------------------------------
# filters left off a page's bar on request: the community page's "level" was
# only the CHW's parent facility level (2026-10-18)
PF_HIDE = {9: ("lvl",)}
# an extra filter on one page: the prescription form's "Name type" per item
# (user request 2026-10-19). A prescription matches if any of its items does.
PF_TAG = {5: {"label": "Type of drug", "all": "Generic and brand",
              "options": [["INN", "Generic (INN)"], ["Brand", "Brand"]]}}
PF_IDS = {1: {k: f"fx-{k}" for k in PF.FILTERS},
          2: {k: f"sx-{k}" for k in PF.FILTERS}}


def _page_filter(a):
    """-> (Page, js) from pf_a<a>.build(), or None if that page has no
    filter module yet."""
    only = os.environ.get("PF_ONLY")      # test builds: just these pages
    if only and str(a) not in only.split(","):
        return None
    try:
        mod = importlib.import_module(f"pf_a{a}")
    except ModuleNotFoundError:
        return None
    pg = PF.Page(f"a{a}", ids=PF_IDS.get(a), hide=PF_HIDE.get(a, ()),
                 tag=PF_TAG.get(a))
    js = mod.build(pg, sys.modules[__name__])
    if js is None:          # the module has nothing left to redraw
        return None
    return pg, js


def pf_fac(r):
    """Stock-tool row -> the facility record every page filter uses."""
    p = r["_p"]
    return [s(p[M.FAC]), s(p[M.DIST]), s(p[M.PROV]),
            FL.group(p[M.LEVEL]), mmgt(r)]


def pf_med(r):
    """Stock-tool row -> [medicine, programme, drug class label]."""
    return [s(r[M.MED]), s(r[M.PROGRAM]), _class_label(r)]


def _pf_cfg():
    return {"min_n": MIN_N, "not_pct": sorted(NOT_PCT), "rest": CH.REST,
            "avail": CH.AVAIL, "g_on": CH.G_ON, "g_off": CH.G_OFF,
            "stack_cols": CH.STACK_COLS}


def write_html():
    # the headline uses the same totals as indicator 7.6; the
    # line-based rate in 6.2-6.5 has a different denominator
    # (medicine-months with a request above zero) and would not match
    _rq = sum(r["_rq"] for r in MG)
    _dv = sum(r["_dv"] for r in MG)
    _fill_total = pct(_dv, _rq)
    av = pct(len(AVM), len(ANSM))
    # Reach, not results: how much each tool covered. Facilities are counted
    # once however many times they were visited. The CHW card breaks the
    # CHWs' parent facilities down by level of care.
    _chw_fac = {}
    for _r in C.MEDROWS:
        _chw_fac[C.s(_r["_p"][C.FAC])] = FL.group(_r["_p"][C.LEVEL])
    _lv = Counter(_chw_fac.values())
    _hosp = sum(v for k, v in _lv.items()
                if k not in ("Health post", "Health centre"))
    stats = [
        (f"{len(M.FACILITIES):,}", "facilities visited",
         "medicine assessment"),
        (f"{len({pfac(r) for r in PRES}):,}", "facilities visited",
         "prescription assessment"),
        (f"{len(C.CHWS):,}", "CHWs assessed", ""),
    ]
    # Facility type, from the medicine assessment (the study's facility
    # frame), with the prescription assessment's count underneath.
    _sf, _pf = {}, {}
    for _r in M.MEDROWS:
        _sf[s(_r["_p"][M.FAC])] = FL.group(_r["_p"][M.LEVEL])
    for _r in PRES:
        _pf[pfac(_r)] = FL.group(_r["_sub"][P.FTYPE])
    _sc, _pc = Counter(_sf.values()), Counter(_pf.values())
    _HOSP = {"District hospital": "DH", "Provincial hospital": "PH",
             "Level 2 teaching hospital": "L2TH",
             "National referral hospital": "NRH"}
    _sh = sum(_sc[k] for k in _HOSP)
    _ph = sum(_pc[k] for k in _HOSP)
    # the health-post card was removed on request (2026-11-05)
    stats += [
        (f"{_sc['Health centre']:,}", "health centres",
         "medicine assessment"),
        (f"{_sh:,}", "hospitals",
         " \u00b7 ".join(f"{_sc[k]} {v}" for k, v in _HOSP.items() if _sc[k])
         + f" \u00b7 {_ph} in prescription"),
    ]
    hero = "".join(f'<div class="stat"><p class="v">{esc(v)}</p>'
                   f'<p class="l">{esc(l)}</p><p class="s">{esc(x)}</p></div>'
                   for v, l, x in stats)
    # Coverage strip under the cards (user request 2026-10-20): medicines
    # assessed per programme (each product / strength once), districts
    # visited, prescriptions and items reviewed.
    # a record with no medicine name is not a product, so it is not counted
    _pm = defaultdict(set)
    for _r in M.MEDROWS:
        if s(_r[M.MED]):
            _pm[s(_r[M.PROGRAM]) or "no programme"].add(s(_r[M.MED]))
    _progs = sorted(_pm.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    _nmed = sum(len(v) for v in _pm.values())
    # Figures as given by the user (2026-10-26): 55 medicines assessed, 25 of
    # them Essential Medicines; the other programme counts stay computed.
    _nmed = 55
    _shown = {"Essential Medicines": 25}
    _dist = {s(_r["_p"][M.DIST]) for _r in M.MEDROWS} | {pdist(_r)
                                                        for _r in PRES}
    _prov = {s(_r["_p"][M.PROV]) for _r in M.MEDROWS}
    cover = (
        '<div class="cover">'
        f'<div class="cv1"><p class="cvv">{_nmed}</p><p class="cvl">medicines '
        'assessed</p><p class="cvs">'
        + " &#183; ".join(f"<b>{_shown.get(k, len(v))}</b> {esc(k)}"
                          for k, v in _progs)
        + '</p></div>'
        f'<div class="cv1"><p class="cvv">{len(_dist)}</p><p class="cvl">'
        f'districts visited</p><p class="cvs">all {len(_prov)} provinces'
        '</p></div>'
        f'<div class="cv1"><p class="cvv">{len(PRES):,}</p><p class="cvl">'
        f'prescriptions reviewed</p><p class="cvs"><b>{len(ITEMS):,}</b> '
        'prescribed items</p></div></div>')
    tabs = "".join(
        f'<button class="tab{" on" if i == 0 else ""}" id="tab-a{a}" '
        f'data-page="a{a}">{a}. {esc(n)}</button>'
        for i, (a, n) in enumerate(PAGES.items()))
    # The "How these figures are calculated" page was removed on request
    # (2026-10-18); method_block() is kept for the workbook's sake.
    pages = []
    _cards = page_cards()
    for i, (a, name) in enumerate(PAGES.items()):
        pf = _page_filter(a)
        parts = [f'<div class="page{" on" if i == 0 else ""}" id="a{a}">'
                 f'<h2>Analysis {a} &middot; {esc(name)}</h2>',
                 _cards.get(a, ""), pf[0].bar() if pf else ""]
        if a == 1:
            # Order on the page: filters, then the maps, then the figures,
            # then the full medicine table. The script goes last because it
            # addresses every container above it.
            facs, meds, _, _, _ = EX_DATA
            parts.append(EX.filters(facs, meds))
            parts.append(maps_block())
            parts.append(charts_block())
            parts.append(EX.table(MIN_N))
            parts.append(explorer_script())
        elif a == 2:
            # the page's own filter bar, then the figures, then the script
            # band 4 = stocked out; shelf rows first so medicine order
            # follows SA_SHELF, as tables 2A/2B do
            sx = SX.build(SA_SHELF + SA_OUT, _acc, _sx_fac, _sx_med,
                          lambda r: 4 if id(r) in SA_OUT_IDS else
                          (3 if _eop(r) else _band(r)), _mos, _exact,
                          qacc=_qacc,
                          system=lambda r: s(r.get("_emr_clean")))
            parts.append(SX.filters(sx[0], sx[1], sx[4]))
            parts.append(PAGE_FIG[a]())
            parts.append(SX.script(*sx, _sx_cfg()))
        elif a in PAGE_FIG:
            parts.append(PAGE_FIG[a]())
        for an, mid, title, status, note, headers, rows in TABLES:
            if an != a or (an, mid) in HTML_DROP:
                continue
            badge = ('<span class="badge">proxy</span>'
                     if status != "Feasible" else "")
            sec = (f'<section class="ind"><h4><span class="mid">{esc(mid)}'
                   f'</span>{esc(title)}{badge}</h4>'
                   + (f'<p class="note">{esc(note)}</p>'
                      if note and (an, mid) not in HTML_NO_NOTE else "")
                   + html_table(headers, rows) + "</section>")
            if an == 1 and mid in AVAIL_B:
                # the availability method switch shows one or the other
                bt, bh, br, bn = AVAIL_B[mid]
                sec = (f'<div class="only-a">{sec}</div>'
                       f'<div class="only-b"><section class="ind"><h4>'
                       f'<span class="mid">{esc(mid)}</span>{esc(bt)}</h4>'
                       f'<p class="note">{esc(bn)}</p>'
                       + html_table(bh, br) + '</section></div>')
            parts.append(sec)
        page = PF.tag_sections("".join(parts), a)
        if pf:
            page += pf[0].script(pf[1], _pf_cfg())
        pages.append(page + "</div>")
    ids_js = ",".join(f'"a{a}"' for a in PAGES)
    body = f"""
<div class="wrap">
<header class="mast">
  <h1>MEDICINE STOCK AND PRESCRIPTION ASSESSMENT FINDINGS</h1>
  <div class="stats">{hero}</div>
  {cover}
</header>
<nav class="tabs">{tabs}</nav>
{''.join(pages)}
</div>
<script>
(function(){{
 var ids=[{ids_js}];
 function show(id){{ids.forEach(function(o){{
   document.getElementById(o).classList.toggle('on',o===id);
   var t=document.getElementById('tab-'+o);
   t.classList.toggle('on',o===id);}});}}
 document.addEventListener('click',function(e){{
   var b=e.target.closest('[data-page]'); if(!b)return;
   show(b.dataset.page); history.replaceState(null,'','#'+b.dataset.page);
   window.scrollTo(0,0);}});
 var h=location.hash.slice(1); if(h&&ids.indexOf(h)>=0) show(h);
}})();
</script>
"""
    body = body.replace("{ids_js}", ids_js)
    doc = ('<!doctype html>\n<html lang="en" class="js">\n<head>\n'
           '<meta charset="utf-8">\n<meta name="viewport" '
           'content="width=device-width,initial-scale=1">\n'
           '<title>Medicine stock and Prescription assessment findings</title>\n'
           + CSS.replace("{LAYER_CSS}", MAP.layer_css() + EX.CSS + CH.CSS
                            + CH.GAUGE_CSS + CH.DIAL_CSS
                            + CH.SANKEY_CSS + MS.CSS)
           + MS.SCRIPT          # multi-select filters (2026-10-26)
           + "</head>\n<body>\n" + body + "\n</body>\n</html>\n")
    doc = add_calc(doc)
    doc = renumber_display(doc)
    check_scripts(doc)
    if os.path.exists(OUT_HTML):
        os.chmod(OUT_HTML, 0o644)
    with open(OUT_HTML, "w", encoding="utf-8") as fh:
        fh.write(doc)


if __name__ == "__main__":
    write_xlsx()
    write_html()
    prox = sum(1 for t in TABLES if t[3] != "Feasible")
    print(f"wrote {os.path.basename(OUT_XLSX)} and "
          f"{os.path.basename(OUT_HTML)}")
    print(f"  {len(TABLES)} indicator tables across {len(ANALYSES)} analyses "
          f"({prox} proxy)")
    print(f"  {os.path.getsize(OUT_HTML):,} bytes of HTML")
    print(f"  headline: availability {pct(len(AVM), len(ANSM)):.1%} | "
          f"fill rate {pct(sum(r['_dv'] for r in MG), sum(r['_rq'] for r in MG)):.1%} | CHW "
          f"{pct(len(CAV), len(CANS)):.1%} | unserved value {round(_val):,} RWF")
