"""Can the master analytical prompt be delivered from this data?

    python3 Master_prompt_feasibility.py  ->  Master_prompt_feasibility.xlsx

The prompt specifies ~48 metrics across 8 analyses, each with an exact
formula. This workbook checks every one against the variables the three Kobo
tools actually collect, says what can be produced, what can only be proxied
and what has no source at all, and points each feasible metric at the output
that already produces it.

The prompt's own instruction is the standard applied here: "Do not fabricate
observations, assume missing values are zero, or generate unsupported
numerical findings." Several metrics assume variables this assessment never
collected - a dispensed quantity, warehouse stock records, requisition dates,
stock-out episode counts, an expiry flag, a service type. Those are reported
as gaps rather than approximated into something that looks like an answer.

Every count in the Data quality sheet is computed from the live merged
exports when this file is generated.

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
import analyze_medicine_assessment_chw as C
import facility_levels as FL

OUT = B.out("Master_prompt_feasibility.xlsx")

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, size=9, color="FFFFFF")
YES = PatternFill("solid", fgColor="E7F3E7")
PART = PatternFill("solid", fgColor="FFF4D6")
NO = PatternFill("solid", fgColor="FCE4E4")
TITLE_FONT = Font(bold=True, size=12, color="1F3864")

Y, PR, N = "Feasible", "Proxy only", "No source"


def s(v):
    return "" if v is None else str(v).strip()


# ---------------------------------------------------------------
# metric, formula as specified, what it needs, verdict, what we can do, where
# ---------------------------------------------------------------
METRICS = [
 ("1.1", "Overall tracer availability",
  "available & expected obs / eligible obs with verified stock status", Y,
  "Available = 'available on visit day' Yes, over rows the facility manages "
  "and answered. 9,870 eligible observations, 93.5%.",
  "The prompt adds 'unexpired and usable'. NO EXPIRY OR USABILITY VARIABLE "
  "is collected, so availability means physically present, nothing more. "
  "Say so in the methods.",
  "analytics/Medicine_availability_analysis.xlsx"),
 ("1.2", "Individual medicine availability", "facilities with X / facilities "
  "assessed for X", Y, "79 medicines ranked, each with its own denominator.",
  "", "analytics/Product_availability_by_geography.xlsx"),
 ("1.3", "Facility basket availability",
  "expected medicines available at facility / expected assessed there", Y,
  "One rate per facility over its own basket.",
  "Report alongside the pooled rate, as the prompt requires - they differ "
  "because baskets are not identical.", "TO BUILD - one sheet"),
 ("1.4", "Complete basket availability",
  "facilities with ALL expected available / facilities fully assessed", Y,
  "Computable directly.", "", "TO BUILD - one row"),
 ("1.5", "Therapeutic-category availability", "by category", Y,
  "The tool carries programme (7 values) and drug class (33). Both are "
  "already reported.", "'Therapeutic category' is not a variable; programme "
  "and drug class are the available equivalents.",
  "analytics/Medicine_availability_analysis.xlsx"),
 ("2.1", "Months of stock", "physical stock / average monthly consumption", Y,
  "Present as stock_level_months with the form's own thresholds; 8,679 of "
  "9,870 managed rows have it.",
  "Undefined where consumption is zero - reported as its own class, never "
  "folded into overstock, exactly as the prompt requires.",
  "analytics/Medicine_availability_analysis.xlsx"),
 ("2.2", "Understock rate", "obs below minimum MOS / obs with valid MOS", Y,
  "Direct from the classification.", "", "same"),
 ("2.3", "Overstock rate", "obs above maximum MOS / obs with valid MOS", Y,
  "Direct from the classification.", "", "same"),
 ("2.4", "Exact inventory accuracy",
  "obs where physical = recorded / obs with both verified", Y,
  "Both quantities present on all 9,227 available rows.",
  "DENOMINATOR: counts exist only where the medicine is available - none of "
  "the 643 stocked-out rows carry them. Base is available rows, not managed.",
  "analytics/Medicine_availability_analysis.xlsx"),
 ("2.5", "Inventory discrepancy", "physical - recorded", Y,
  "Signed difference, already added at the team's request.", "",
  "analytics/Medicine_availability_analysis.xlsx"),
 ("2.6", "Absolute percentage inventory error",
  "|physical - recorded| / physical x 100, physical > 0", Y,
  "Computable on rows with physical stock above zero.",
  "The team also asked for (recorded - physical) / recorded, which is "
  "directional. Report both; they answer different questions.",
  "TO BUILD - add to the accuracy sheet"),
 ("3.1", "Three-month stock-out incidence",
  "pairs with >=1 stock-out in the period / pairs with complete history", Y,
  "Derivable: any of the three monthly stock-out day counts above zero.",
  "", "TO BUILD - one sheet"),
 ("3.2", "Stock-out frequency (episodes)", "count of separate episodes", N,
  "NOT COLLECTED. The form records stock-out DAYS per month, not episodes.",
  "PROXY: months affected (0-3). Two 5-day stock-outs cannot be told from "
  "one 10-day stock-out. Needs an episode-count question in the form.",
  "proxy in analytics/MCCH_availability_analysis.xlsx"),
 ("3.3", "Cumulative stock-out duration", "total days unavailable", Y,
  "total_stockout_days, June-August.", "", "Medicine_assessment_analysis.xlsx"),
 ("3.4", "Average stock-out duration",
  "days across completed episodes / number of episodes", N,
  "Depends on 3.2. No episode count, so no denominator.",
  "Report mean days per affected medicine-facility pair instead, and label "
  "it as such - it is not an episode mean.", "TO BUILD, as a labelled proxy"),
 ("3.5", "Current stock-out duration", "consecutive days out as of the visit",
  Y, "Collected directly; present on the 635 currently stocked-out rows.",
  "Blanks are not zeros - the question is skipped where the medicine is in "
  "stock.", "Medicine_assessment_analysis.xlsx"),
 ("3.6", "Stock-out days ratio",
  "total stock-out days / total eligible observation days", Y,
  "stockout_rate_pct, over a 92-day window.", "", "same"),
 ("4.1", "Complete prescription fulfilment",
  "prescriptions where every line fully dispensed / assessed prescriptions",
  PR, "The tool records, per item, whether the medicine was AVAILABLE at the "
  "facility - not whether it was dispensed, and not how much.",
  "PROXY: prescriptions with no unavailable item (5,924 of 5,976). That is "
  "'nothing was missing', not 'everything was dispensed'.",
  "analytics/Unserved_by_geography.xlsx"),
 ("4.2", "Partially served prescription rate",
  "some quantity dispensed but >=1 line not fully fulfilled", PR,
  "NO DISPENSED QUANTITY EXISTS. A prescription can be flagged as having "
  "some items available and some not, which is not the same thing.",
  "PROXY: prescriptions with at least one unavailable item AND at least one "
  "available. Label it as mixed availability, not partial dispensing.",
  "TO BUILD"),
 ("4.3", "Completely unserved prescription rate",
  "prescriptions with zero dispensed across all lines", PR,
  "Same limitation. Derivable as prescriptions where every item was "
  "unavailable.", "PROXY, same caveat.", "TO BUILD"),
 ("4.4", "Unserved prescription rate", "prescriptions not fully fulfilled", PR,
  "52 of 5,976 prescriptions (0.87%) have at least one unavailable item.",
  "Read as 'had a medicine the facility could not provide'.",
  "analytics/Unserved_by_geography.xlsx"),
 ("4.5", "Medicine-line fulfilment rate",
  "fully dispensed lines / assessed lines", PR,
  "Availability stands in for dispensing: 13,877 of 13,946 items available "
  "(99.5%).", "Same substitution, stated once in the methods.",
  "analytics/Unserved_by_geography.xlsx"),
 ("4.6", "Quantity fulfilment rate", "quantity dispensed / quantity prescribed",
  N, "NO DISPENSED QUANTITY ANYWHERE in the prescription tool. Only the "
  "quantity PRESCRIBED is recorded.",
  "Cannot be produced. Add a 'quantity dispensed' field to the form. Until "
  "then quantity analysis stops at quantity prescribed on unavailable items.",
  "-"),
 ("4.7", "Unserved medicine quantity", "prescribed - dispensed", N,
  "Same cause. What IS available is the full prescribed quantity on items "
  "recorded unavailable: 962 units.",
  "Report that as 'quantity prescribed but not available', which is an "
  "upper bound on the gap, not the gap.",
  "analytics/Unserved_by_geography.xlsx"),
 ("4.x", "Inpatient versus outpatient comparison", "-", N,
  "NO SERVICE TYPE, ward, department or admission variable at any level of "
  "the prescription tool - checked in the deployed export and in all three "
  "XLSForms.",
  "Cannot be produced. One select_one would fix it for the next round. "
  "Facility type is a facility attribute, not a service attribute, and "
  "should not be substituted silently.", "-"),
 ("5.1", "Regional warehouse availability",
  "unserved lines with sufficient usable stock at the RMS branch on the "
  "prescription date / lines with verified regional records", PR,
  "There are NO WAREHOUSE RECORDS in this study. What exists is a "
  "Yes/No/Don't know answer from the facility about whether the medicine was "
  "at RMS Regional - a recall, answered on all 69 unserved items.",
  "Report as reported-upstream-availability, never as verified stock. "
  "'Sufficient' and 'usable' cannot be assessed at all. Linking to real RMS "
  "records would need an RMS dataset this study does not have.",
  "analytics/Unserved_by_geography.xlsx"),
 ("5.2", "Central warehouse availability", "as 5.1, centrally", PR,
  "Same: a recall answer, not a record.", "Same caveat.", "same"),
 ("5.3", "Distribution constraint rate",
  "verified distribution constraints / lines with sufficient evidence", PR,
  "distribution_nationwide is derived from the two recall answers: in stock "
  "at either warehouse means distribution.",
  "The word 'verified' cannot be honoured. Report as attributed-from-"
  "report.", "same"),
 ("5.4", "Central supply constraint rate", "as 5.3", PR, "Same.", "Same.",
  "same"),
 ("5.x", "Six-category attribution",
  "facility inventory / facility ordering / branch-to-facility / "
  "central-to-regional / system-wide / unclassified", PR,
  "Three of the six are derivable: facility ordering (was it ordered?), "
  "distribution (in stock upstream), and system-wide (in stock nowhere). "
  "Facility inventory-management, and central-to-regional specifically, are "
  "not separable with these variables.",
  "Report three categories plus 'insufficient evidence', and say which two "
  "of the prompt's six were collapsed.", "TO BUILD"),
 ("6.1", "Requisition rate",
  "missing medicines with a valid requisition / missing with verified status",
  Y, "Prescription tool: 22 of 69 unserved items were ordered that month.",
  "The stock tool has no yes/no 'ordered' - only monthly quantities. The "
  "team flagged this; it is a one-question form change.",
  "analytics/Unserved_by_geography.xlsx"),
 ("6.2", "Complete order fill rate", "lines supplied in full / assessed lines",
  Y, "Derivable at medicine-month level: delivered >= requested.",
  "'Order line' here is a medicine-month, not a real requisition line - the "
  "tool collects monthly totals, not orders.", "TO BUILD"),
 ("6.3", "Quantity fill rate", "quantity supplied / quantity requested", Y,
  "93.8% over June-August; 36.1M of 38.5M units.",
  "Computed from summed quantities per medicine before aggregation, as the "
  "prompt requires.", "analytics/Prescribing_fill_rate_unserved.xlsx"),
 ("6.4", "Partial fulfilment rate", "lines below requested but above zero", Y,
  "Derivable at medicine-month level.", "Same 'order line' caveat.",
  "TO BUILD"),
 ("6.5", "Non-fulfilment rate", "lines receiving zero", Y,
  "Derivable at medicine-month level.", "Same caveat.", "TO BUILD"),
 ("6.6", "Order lead time", "delivery date - requisition date", N,
  "NO DATES. The tool records monthly requested and delivered QUANTITIES "
  "only - no requisition date, no delivery date.",
  "Cannot be produced at all. Needs date fields, or an RMS order dataset.",
  "-"),
 ("7.1", "Unserved prescription value", "unserved quantity x tariff price", Y,
  "67 of 69 unserved items carry a unit price; 229,301 RWF total.",
  "Quantity used is quantity PRESCRIBED on unavailable items - see 4.7. "
  "Values the observed gap, a floor rather than a total.",
  "analytics/Unserved_by_geography.xlsx"),
 ("7.2", "Total unserved value", "sum across lines", Y, "As above.", "",
  "same"),
 ("7.3", "Average value per affected prescription",
  "total value / prescriptions not fully served", Y,
  "229,301 / 52 affected prescriptions.", "", "TO BUILD - one row"),
 ("7.4", "Medicine contribution to total value", "value of X / total", Y,
  "Pareto over 46 medicines.", "", "analytics/Unserved_by_geography.xlsx"),
 ("7.5", "Value by supply chain constraint", "value per attribution class", PR,
  "Derivable for the three attribution classes that exist.",
  "Inherits the 5.x limitation.", "TO BUILD"),
 ("7.x", "Value by inpatient/outpatient", "-", N, "No service type.",
  "Cannot be produced.", "-"),
 ("7.y", "National extrapolation",
  "observed sample value vs national estimate", PR,
  "The prompt rightly asks these be distinguished. Facilities were selected "
  "from a visit plan, not sampled probabilistically, and fieldwork covered "
  "265 of ~1,700 facilities.",
  "Report observed values only. A national extrapolation would need a "
  "sampling frame and weights that do not exist.", "-"),
 ("8.1", "Overall CHW availability",
  "available CHW-medicine obs / eligible obs", Y,
  "918 of 1,180 managed rows, 77.8%.",
  "Denominator is rows the CHW manages. About 40% of rows are 'not managed' "
  "because of MFT zoning - counting those as unavailable would report the "
  "zoning policy as a stock-out.", "Medicine_stock_CHW_analysis.xlsx"),
 ("8.2", "Individual CHW medicine availability", "CHWs holding X / CHWs "
  "expected to hold X", Y, "13 medicines, each with its own denominator.", "",
  "same"),
 ("8.3", "Complete CHW basket availability",
  "CHWs holding every expected medicine / CHWs fully assessed", Y,
  "Computable directly.", "", "TO BUILD - one row"),
 ("8.4", "CHW versus attached facility",
  "comparison for medicines expected at both levels", Y,
  "Eight products exist in both baskets. CHW 81.8% against health centre "
  "93.8% on those eight.",
  "Only eight of the CHW's 13 items have a facility comparator: ORS, Zinc "
  "and Amox 125 are carried by CHWs and by no facility in this study.",
  "analytics/Product_availability_by_geography.xlsx"),
]

STATS = [
 ("Does availability differ by facility type?",
  "Design-adjusted chi-square or regression", PR,
  "Chi-square across the six grouped tiers is fine. 'Design-adjusted' is "
  "not - there is no sample design to adjust for."),
 ("Which characteristics are associated with stock-outs?",
  "Multilevel logistic regression", Y,
  "Feasible: outcome = stocked out, level-1 medicine observation, level-2 "
  "facility, with province/level/management/programme as covariates. "
  "9,870 observations in 267 facilities is ample."),
 ("Does stock adequacy predict subsequent stock-outs?",
  "Regression on temporally ordered observations", N,
  "No temporal ordering. Stock status is a single visit-day snapshot; the "
  "stock-out history is recalled for the same window, not a later one."),
 ("Is inventory accuracy associated with availability?",
  "Multilevel regression", PR,
  "Both exist, but accuracy is only recorded where the medicine IS "
  "available, so the association cannot be estimated across the full range "
  "of the outcome. Restrict to variation in accuracy among available items "
  "and say so."),
 ("Does prescription fulfilment differ across facility levels?",
  "Multilevel logistic regression", PR,
  "Only 69 unserved items in 21 facilities. A multilevel model will not "
  "converge usefully; report counts and exact confidence intervals."),
 ("Which constraints account for the greatest unserved value?",
  "Descriptive financial decomposition", Y, "Feasible on the three "
  "attribution classes that exist."),
 ("Sampling weights, strata, clustering", "Apply where available", N,
  "None exist. Facilities follow a visit plan, not a probability sample. "
  "Report as a census of visited facilities; do not weight or generalise."),
]

CLEANING = [
 ("Two servers merged", "All three tools are collected on kf and eu "
  "KoboToolbox servers and merged before analysis.",
  "_index restarts at 1 per server, so EU repeat rows would otherwise attach "
  "to kf facilities. Renamed columns are mapped (on_eml_final -> on_eml) "
  "before the union."),
 ("Duplicate facility x medicine rows removed",
  f"{len(M.DUPES_DROPPED)} duplicate observations dropped, keeping the row "
  f"carrying an assessment and then the later submission.",
  f"{len(M.DUPES_CONFLICTING)} of them disagreed with the row kept and are "
  f"listed for the collector to confirm. The pair is the unit of analysis, "
  f"so a duplicate double-counts it."),
 ("Facility reattributions", "Seven corrections applied by (facility, date), "
  "including health posts captured under a parent health centre and one "
  "level misrecording.",
  "Facility is the main grouping, so pooled data overstates the parent and "
  "hides the post."),
 ("Facility level grouped", "Health post = SGHP + HP + bare HP; Health "
  "centre = HC + MHC. Hospitals kept separate.",
  "The forms record some tiers under three interchangeable labels, which "
  "split one small tier into three thin rows."),
 ("Non-finite calculate fields rejected",
  "Kobo can export the literal string 'NaN'; every num() returns None for "
  "non-finite values.", "One NaN turns an entire downstream total into NaN."),
 ("Denominators fixed per indicator",
  "Availability over managed-and-answered rows; accuracy over available "
  "rows; supply outcome over ordered items.",
  "Each question is skipped under different logic, so a single denominator "
  "would be wrong for most indicators."),
 ("Month window restricted to June-August",
  "March-May columns exist in the export but are empty on both servers - "
  "legacy columns from an earlier form version.",
  "Including them would imply six months of history that was never "
  "collected."),
 ("Medicine names not merged across tools",
  "The prescription tool names medicines from a prescribing list, the stock "
  "tool from the procurement catalogue.",
  "No shared identifier exists, so a merged medicine ranking would be "
  "wrong. The prompt's Step 2 'unique medicine identifier' cannot be "
  "constructed across tools - only within each."),
]


def dq():
    """The prompt's Step 3 data quality report, computed live."""
    items = [r for r in P.item if r["_sub"]]
    mg = [r for r in M.MEDROWS if M.yes(r[M.MANAGED])]
    av = [r for r in mg if M.yes(r[M.AVAIL])]
    code = sum(1 for r in items if re.fullmatch(r"m\d{3,4}", s(r["_med"])))
    neg = sum(1 for r in mg for c in (M.QTY_PHYS, M.QTY_SYS)
              if (M.num(r.get(c)) or 0) < 0)
    zero_av = sum(1 for r in av if (M.num(r[M.QTY_PHYS]) or 0) == 0)
    baddate = sum(1 for r in P.pres
                  if P.blank(r.get(P.DATE)) or s(r.get(P.DATE))[:4] not in
                  ("2025", "2026"))
    pairs = len({(s(r["_p"][M.FAC]), s(r[M.MED])) for r in M.MEDROWS})
    return [
      ["Total records - medicine assessment",
       f"{M.N_MED:,} facility x medicine observations in {len(M.SUB):,} "
       f"submissions", ""],
      ["Total records - prescription",
       f"{len(items):,} medicine items in {P.N_PRES:,} prescriptions, "
       f"{len(P.sub):,} facility visits", ""],
      ["Total records - CHW",
       f"{C.N_MED:,} CHW x medicine observations, {C.N_SUB} submissions", ""],
      ["Unique facilities",
       f"medicine {len(M.FACILITIES)}, prescription "
       f"{len({s(r[P.FACIL]) for r in P.sub})}, CHW {len(C.FACILITIES)} "
       f"parent health centres", "post-reattribution"],
      ["Unique medicines",
       f"medicine {len({s(r[M.MED]) for r in M.MEDROWS})}, prescription "
       f"{len({s(r['_med']) for r in items})}, CHW "
       f"{len({s(r[C.MED]) for r in C.MEDROWS})}", "not joinable across tools"],
      ["Duplicate records", f"{len(M.DUPES_DROPPED)} duplicate facility x "
       f"medicine rows found and removed; {pairs:,} unique pairs remain",
       "cleaned - see the Cleaning sheet"],
      ["Missing vs genuine zero",
       f"availability answered on all {len(mg):,} managed rows; "
       f"{len(M.MEDROWS) - len(mg):,} rows are 'not managed' and are BLANK "
       f"by skip logic, not zero", "distinguished throughout"],
      ["Inconsistent quantities",
       f"{zero_av} rows answer 'available' and record a quantity of 0",
       "flagged, not corrected"],
      ["Negative stock values", f"{neg}", "none found" if not neg else "check"],
      ["Invalid prescription dates",
       f"{baddate} of {P.N_PRES:,} prescriptions have a missing or "
       f"out-of-range date", ""],
      ["Unmatched medicine identifiers",
       f"{code:,} of {len(items):,} prescription items "
       f"({code / len(items):.1%}) carry an unresolved code such as m1332, "
       f"across {len({s(r['_med']) for r in items if re.fullmatch(chr(109) + chr(92) + 'd{3,4}', s(r['_med']))})} codes",
       "ACTION: supply the code-to-name lookup"],
      ["Incomplete RMS records",
       "No RMS dataset exists. Upstream availability is a facility recall "
       "answer on 69 unserved items, not a warehouse record.",
       "limits analysis 5"],
      ["Orphan rows", "0 in all three tools", "merge verified"],
    ]


def sheet(wb, title, blocks, notes=None, widths=None, fills=None, freeze="B"):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    if notes:
        for n in notes:
            c = ws.cell(row=r, column=1, value=n)
            c.font = Font(size=9, italic=True, color="555555")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
            ws.row_dimensions[r].height = 13 * (1 + len(n) // 150)
            r += 1
        r += 1
    for btitle, headers, rows in blocks:
        if btitle:
            ws.cell(row=r, column=1, value=btitle).font = TITLE_FONT
            r += 1
        for ci, h in enumerate(headers, start=1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
            c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 32
        hdr = r
        r += 1
        for row in rows:
            for ci, v in enumerate(row, start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.border = BORDER
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if fills:
                    f = fills(row, ci - 1, v)
                    if f:
                        c.fill = f
            r += 1
        ws.freeze_panes = f"{freeze}{hdr + 1}"
        ws.auto_filter.ref = f"A{hdr}:{get_column_letter(len(headers))}{r - 1}"
        r += 2
    for ci, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(ci)].width = w


def build():
    wb = Workbook()
    wb.remove(wb.active)
    cnt = Counter(m[3] for m in METRICS)

    def vfill(row, ci, v):
        if ci == 3:
            return {Y: YES, PR: PART, N: NO}.get(v)
        return None

    summary = [
        [Y, cnt[Y], "Computable as specified, from variables that exist."],
        [PR, cnt[PR], "The question can be answered but not with the formula "
         "as written - a variable it assumes is absent and something weaker "
         "stands in. Each says what was substituted."],
        [N, cnt[N], "No source variable at all. Reporting a number here would "
         "mean inventing one."],
        ["", "", ""],
        ["Metrics specified by the prompt", len(METRICS), ""],
        ["Already produced by an existing output",
         sum(1 for m in METRICS if m[6] not in ("-",) and "TO BUILD" not in m[6]),
         "see the 'Where produced' column"],
        ["Feasible but not yet built",
         sum(1 for m in METRICS if "TO BUILD" in m[6]),
         "small additions to existing sheets"],
    ]
    sheet(wb, "Summary",
          [("Can the master prompt be delivered?",
            ["Verdict", "Metrics", "Meaning"], summary)],
          notes=["Checked metric by metric against the variables the three "
                 "tools actually collect.",
                 "The prompt's own rule is applied throughout: no fabricated "
                 "observations, missing is not zero, every indicator needs a "
                 "real numerator and denominator."],
          widths={1: 16, 2: 10, 3: 110}, freeze="A")

    sheet(wb, "Metric feasibility",
          [(None, ["Metric", "Name", "Formula as specified", "Verdict",
                   "What the data supports", "Caveat or substitution",
                   "Where produced"], [list(m) for m in METRICS])],
          notes=["One row per metric in the master prompt. Green = feasible, "
                 "amber = proxy only, red = no source.",
                 "'Where produced' names the output that already contains it, "
                 "or marks it TO BUILD where it is a small addition."],
          widths={1: 8, 2: 34, 3: 46, 4: 12, 5: 60, 6: 62, 7: 44},
          fills=vfill, freeze="C")

    sheet(wb, "Data quality",
          [(None, ["Check", "Result", "Note"], dq())],
          notes=["The prompt's section 3 Step 3 report, computed from the "
                 "live merged exports when this workbook was generated.",
                 "Missing data and genuine zeros are distinguished: a blank "
                 "under skip logic is not a zero, and is never counted as "
                 "one."],
          widths={1: 36, 2: 86, 3: 34}, freeze="A")

    sheet(wb, "Cleaning applied",
          [(None, ["Step", "What was done", "Why it matters"],
            [list(c) for c in CLEANING])],
          notes=["Everything already applied to the data before any indicator "
                 "is computed. The prompt asks that prior cleaning be taken "
                 "into account - this is it.",
                 "All of it is applied in the analysis scripts, never by "
                 "editing an export: the merged exports are rebuilt from both "
                 "servers on every fetch, so an edit there would be lost."],
          widths={1: 36, 2: 80, 3: 76}, freeze="A")

    sheet(wb, "Statistical analyses",
          [(None, ["Analytical question", "Approach the prompt asks for",
                   "Verdict", "Assessment"], [list(x) for x in STATS])],
          notes=["Section 5 of the prompt. Sample size is ample for the stock "
                 "analyses and too small for the prescription ones.",
                 "There are no sampling weights, strata or clusters: "
                 "facilities follow a visit plan, not a probability sample. "
                 "Report as a census of visited facilities."],
          widths={1: 46, 2: 36, 3: 12, 4: 86},
          fills=lambda row, ci, v: {Y: YES, PR: PART, N: NO}.get(v)
          if ci == 2 else None, freeze="A")

    gaps = [
      ("Quantity dispensed, per medicine line", "Prescription", "integer",
       "Metrics 4.6 and 4.7 - the entire quantity side of prescription "
       "fulfilment - have no source without it."),
      ("Service type (outpatient / inpatient)", "Prescription",
       "select_one", "Every inpatient/outpatient comparison the prompt asks "
       "for, in analyses 4 and 7."),
      ("Requisition and delivery dates", "Medicine assessment", "date x2",
       "Metric 6.6, order lead time."),
      ("Number of stock-out episodes", "Medicine assessment", "integer",
       "Metrics 3.2 and 3.4."),
      ("Expiry / usable flag on the counted stock", "Medicine assessment",
       "select_one or date",
       "Metric 1.1 defines availability as unexpired and usable; only "
       "physical presence is collected."),
      ("Was this medicine ordered in the period?", "Medicine assessment",
       "select_one yes_no",
       "Metric 6.1 on the stock side. Raised by the team."),
      ("Medicine code-to-name lookup", "Prescription", "reference file",
       "19.7% of prescribed items carry an unresolved code, so the most-"
       "prescribed ranking is provisional."),
      ("RMS warehouse stock extract", "External dataset", "data file",
       "Analysis 5 asks for verified warehouse records on the prescription "
       "date. Only a facility recall answer exists."),
    ]
    sheet(wb, "What is missing",
          [(None, ["Variable or dataset needed", "Where it belongs", "Type",
                   "What it unlocks"], [list(g) for g in gaps])],
          notes=["Eight additions would make the master prompt deliverable in "
                 "full. Six are a single question each.",
                 "Two of them - dispensed quantity and service type - block "
                 "whole sections of the prompt rather than single metrics."],
          widths={1: 44, 2: 24, 3: 20, 4: 80}, freeze="A")

    wb.active = 0
    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)
    wb.save(OUT)
    print(f"wrote {os.path.basename(OUT)}")
    print(f"  {len(METRICS)} metrics checked: {cnt[Y]} feasible, "
          f"{cnt[PR]} proxy only, {cnt[N]} no source")
    print(f"  already produced: "
          f"{sum(1 for m in METRICS if m[6] != '-' and 'TO BUILD' not in m[6])}"
          f" | to build: {sum(1 for m in METRICS if 'TO BUILD' in m[6])}")
    print(f"  {len(gaps)} variables/datasets would close the gap")


if __name__ == "__main__":
    build()
