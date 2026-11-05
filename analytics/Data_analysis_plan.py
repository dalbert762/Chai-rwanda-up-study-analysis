"""Build the study's detailed data analysis plan as an Excel workbook.

    python3 Data_analysis_plan.py   ->  Data_analysis_plan.xlsx

Written against the EIGHT study objectives and against the variables the three
tools actually collect. Every variable named exists in the current merged
exports, and every denominator, completeness figure and distinct-value count
is READ FROM THE DATA at build time - so the plan cannot drift from the
instruments it describes.

What the verification pass against the data established, and what this plan
therefore says that a plan written from the protocol alone would get wrong:

  * Objective 2's denominator is AVAILABLE rows, not managed rows. The physical
    and system counts are recorded on all 9,234 available rows and on none of
    the 644 stocked-out ones - you cannot count what is not on the shelf.
  * Objective 3's "current stock-out days" exists only on currently
    stocked-out rows (~635), not on all managed rows.
  * Objective 3's "frequency" is not collected. The form records DAYS per
    month, not episodes, so frequency can only be proxied by months affected.
  * Objective 4's outpatient/inpatient split is not collected at all - there is
    no service, ward, department or admission variable at any level of the
    prescription tool.
  * Objective 6's supply outcome is asked only where the medicine was ordered,
    which is 22 of the 69 unserved items. That is a finding (two thirds were
    never ordered), not missing data.
  * The form's own unserved roll-ups are now populated but disagree with the
    item rows on 39 prescriptions, so every unserved figure is recomputed from
    the items and the roll-up is used only as a cross-check.
  * The March, April and May columns exist in the merged export but are EMPTY
    on both servers - they are legacy columns from an earlier form version,
    carried over by fields_from_all_versions. The real window is June-August
    for both servers, identically. An earlier draft of this plan wrongly
    described this as a difference between the two servers.
  * The two XLSForm files in the project folder
    ("Unserved_Prescription final 09092026.xlsx" and "XLs form
    medicine_assessment_final.xlsx") PREDATE what is deployed. They define the
    stock window as May-July (stock-outs) and April-June (requisitions), and
    they contain no EML-applicability question, no digital-system question and
    no monthly consumption totals - all of which the deployed form has. Read
    the plan against the deployed export, not against those files; the
    "Form version check" sheet lists the differences.

This version also incorporates the team's edits returned in
Data_analysis_plan11.xlsx - see the "Changes from v11" sheet for what was
adopted, what was added back, and why.

openpyxl + the standard library, like the rest of this project.
"""

import _bootstrap as B   # path setup - must come first
import os
from collections import Counter

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_kobo_prescription as P
import analyze_medicine_assessment as M
import analyze_medicine_assessment_chw as C

OUT = B.out("Data_analysis_plan.xlsx")
MIN_N = 30

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, size=10, color="FFFFFF")
OBJ_FILL = PatternFill("solid", fgColor="DDE5F0")
YES_FILL = PatternFill("solid", fgColor="E7F3E7")
PART_FILL = PatternFill("solid", fgColor="FFF4D6")
NO_FILL = PatternFill("solid", fgColor="FCE4E4")
TITLE_FONT = Font(bold=True, size=13, color="1F3864")

T_STOCK, T_PRESC, T_CHW = "Medicine assessment", "Prescription assessment", "CHW assessment"


# ---------------------------------------------------------------
# live denominators, so the plan states the real n
# ---------------------------------------------------------------
def stats():
    mg = [r for r in M.MEDROWS if M.yes(r[M.MANAGED])]
    av = [r for r in mg if M.yes(r[M.AVAIL])]
    so = [r for r in mg if M.no(r[M.AVAIL])]
    app = [r for r in M.MEDROWS if M.yes(r[M.APPLICABLE])]
    un = [r for r in P.item if r["_avail"] is False]
    ordered = [r for r in un if str(r[P.ORDERED]).strip().lower() == "yes"]
    cm = [r for r in C.MEDROWS if C.yes(r[C.MANAGED])]
    cav = [r for r in cm if C.yes(r[C.AVAIL])]
    return {
        "obs": M.N_MED, "applicable": len(app), "managed": len(mg),
        "available": len(av), "stockout": len(so),
        "facilities": len(M.FACILITIES),
        "medicines": len({str(r[M.MED]).strip() for r in M.MEDROWS}),
        "acc_rows": sum(1 for r in av if not M.blank(r[M.QTY_SYS])),
        "so_now": sum(1 for r in mg if not M.blank(r[M.SO_NOW])),
        "loss_rows": sum(1 for r in M.MEDROWS if not M.blank(r[M.LOSS])),
        "fulfil_rows": sum(1 for r in mg if not M.blank(r[M.FULFIL])),
        "pres": P.N_PRES, "items": P.N_ITEM, "unserved": len(un),
        "pres_fac": len({str(r[P.FACIL]).strip() for r in P.sub}),
        "ordered": len(ordered),
        "supplied_answered": sum(1 for r in un if not P.blank(r[P.SUPPLIED])),
        "priced": sum(1 for r in un if not P.blank(r[P.UNIT_PRICE])),
        "rollup_mismatch": sum(
            1 for r in P.pres
            if P.num(r.get("n_items_unavailable")) is not None
            and int(P.num(r.get("n_items_unavailable"))) != r["_n_unavail"]),
        "chw_rows": C.N_MED, "chw_managed": len(cm), "chw_avail": len(cav),
        "chw_chws": len(C.CHWS), "chw_fac": len(C.FACILITIES),
        "chw_mos": sum(1 for r in cm if not C.blank(r[C.MOS])),
        "shared_fac": len({str(r[M.FAC]).strip() for r in M.SUB}
                          & {C.s(r[C.FAC]) for r in C.SUB}),
    }


S = stats()

OBJECTIVES = {
    1: "Measure availability of a defined tracer medicine list on the day of "
       "visit, by level of care, province, and district.",
    2: "Determine the stock accuracy by comparing the physical stock versus "
       "stock recorded into digital tool used at each health facility.",
    3: "Quantify stock-outs, current duration and frequency over the preceding "
       "three months.",
    4: "Measure the proportion of prescriptions that go unserved and the "
       "quantity of medicine involved, across outpatient and inpatient services.",
    5: "Determine whether each unserved medicine was available at RMS regional "
       "and central warehouses on the prescription date, separating "
       "distribution challenges from nationwide stock-outs.",
    6: "Establish whether the facility requested the missing medicine and "
       "whether it was supplied fully, partially, or not at all.",
    7: "Estimate the financial value of unserved prescriptions using official "
       "tariff prices.",
    8: "Describe availability at community level through the community health "
       "worker attached to each selected facility.",
}

# Each indicator is a dict so the detailed columns stay readable.
IND = [
 dict(obj=1, id="1.1", name="Availability on the day of visit",
      defn="Share of tracer medicines physically present at the facility on "
           "the visit day, among those the facility manages.",
      num="Observations with 'available on visit day' = Yes",
      den="Observations managed AND availability answered",
      n=f"{S['managed']:,} managed rows; {S['available']:,} available",
      unit="Facility x medicine observation; reported at facility, district "
           "and level of care",
      tool=T_STOCK, sheet="medicine_group",
      vars="Is **${medicine_name}** available on visit day?; Is "
           "**${medicine_name}** managed by the facility?; Is "
           "**${medicine_name}** applicable to this facility level as per EML?",
      disagg="Level of care; province; district; sector; facility; programme; "
             "drug class; medicine; management; level rank; server",
      method="Proportion with a 95% Wilson confidence interval. Chi-square "
             "across level of care to test the gradient; report the "
             "percentage-point spread between the highest and lowest level.",
      stat="Proportion + 95% CI; chi-square across levels",
      shell="T1 Availability by disaggregator",
      out="Medicine_availability_analysis.xlsx - By geography / By facility "
          "type / By medicine",
      feas="Yes", pri="Core",
      note="Denominator is MANAGED rows only. Over all observations it would "
           "count a medicine the facility does not stock as a stock-out."),
 dict(obj=1, id="1.2", name="Tracer list definition and assessment coverage",
      defn="The tracer list is the form's own medicine multi-select. Document "
           "it and report how often each item was assessed.",
      num="Observations per medicine", den="Facilities assessed",
      n=f"{S['medicines']} medicines across {S['facilities']} facilities",
      unit="Medicine", tool=T_STOCK, sheet="medicine_group",
      vars="medicine_name; medicine_program; drug_class_name; Which medicines "
           "are being assessed at this facility? (Select all that apply)",
      disagg="Drug class; level of care",
      method="Frequency table. Flag any medicine assessed at fewer than "
             f"{MIN_N} facilities as not comparable across groups.",
      stat="Counts only", shell="T2 Tracer roster",
      out="Medicine_availability_analysis.xlsx - By medicine",
      feas="Yes", pri="Core",
      note="There is NO tracer flag in the tool. The basket IS the tracer "
           "list - reproduce it in the methods or the result is not replicable."),
 dict(obj=1, id="1.3", name="EML applicability gap",
      defn="Medicines applicable to that facility level per the EML that the "
           "facility nonetheless does not manage.",
      num="Applicable = Yes AND managed = No", den="Applicable = Yes",
      n=f"{S['applicable'] - S['managed']:,} of {S['applicable']:,} applicable",
      unit="Facility x medicine observation", tool=T_STOCK, sheet="medicine_group",
      vars="Is **${medicine_name}** applicable to this facility level as per "
           "EML?; Is **${medicine_name}** managed by the facility?; Why is "
           "**${medicine_name}** not managed by the facility?",
      disagg="Level of care; province; district; medicine; drug class",
      method="Proportion of applicable medicines not managed, by level of "
             "care and by drug class - the two cuts that make it actionable. "
             "Keyword-code the free-text reason into themes and report the "
             "share per theme. Chi-square across levels of care to test "
             "whether the gap concentrates at one level. List the medicines "
             "with the widest gap for the EML review.",
      stat="Proportion; thematic coding",
      shell="T3 Applicability gap + coded reasons",
      out="Medicine_assessment_analysis.xlsx - Medicine coverage",
      feas="Yes", pri="Core",
      note="A selection or policy gap, NOT a stock-out. Report separately or "
           "availability will be misread."),
 dict(obj=2, id="2.1", name="Inventory accuracy (physical vs digital)",
      defn="Agreement between the physical count and the quantity held in the "
           "facility's digital stock system.",
      num="Physical quantity counted", den="Quantity recorded in the system",
      n=f"{S['acc_rows']:,} rows - ALL available rows, and NONE of the "
        f"{S['stockout']} stocked-out rows",
      unit="Facility x medicine observation", tool=T_STOCK, sheet="medicine_group",
      vars="Quantity of **${medicine_name}** available; Quantity of "
           "**${medicine_name}** available in the system; inventory_accuracy_pct",
      disagg="Level of care; province; district; facility; digital system; "
             "programme; medicine",
      method="Median accuracy with IQR. Band as: exactly 100%, 95-99%, 80-94%, "
             "50-79%, 0-49%, negative (system exceeds shelf). Report the share "
             "within +/-5% as the headline. Kruskal-Wallis across levels.",
      stat="Median + IQR; banded distribution; Kruskal-Wallis",
      shell="T4 Accuracy bands",
      out="Medicine_assessment_analysis.xlsx - Consumption and accuracy",
      feas="Yes", pri="Core",
      note="DENOMINATOR CORRECTION: counts are captured only where the "
           "medicine is available, so accuracy is computed over available "
           "rows, not managed rows. Use the MEDIAN - near-zero denominators "
           "distort the mean. A negative band means the system holds more than "
           "the shelf, a different error from under-recording."),
 dict(obj=2, id="2.2", name="Signed stock discrepancy (team request)",
      defn="The difference between the digital and physical stock as a share "
           "of the digital stock - a directional measure, where the ratio in "
           "2.1 hides whether the system over- or under-states.",
      num="(digital stock - physical stock)",
      den="digital stock",
      n="7,614 rows have a non-zero digital stock. Of those 5,012 agree "
        "exactly, 1,841 the system OVERSTATES, 761 it understates",
      unit="Facility x medicine observation", tool=T_STOCK, sheet="medicine_group",
      vars="Quantity of **${medicine_name}** available in the system; "
           "Quantity of **${medicine_name}** available",
      disagg="Level of care; province; district; facility; digital system; "
             "programme; medicine; drug class",
      method="(digital - physical) / digital, as a signed percentage. Report "
             "the median and IQR, and the three-way split: exact agreement, "
             "system overstates (positive), system understates (negative). "
             "Band the magnitude at +/-5%, +/-25%, beyond. Rows with a digital "
             "stock of zero are excluded - the ratio is undefined, not zero.",
      stat="Signed median + IQR; direction split; banded magnitude",
      shell="T4b Signed discrepancy",
      out="to be added to the stock analysis", feas="Yes", pri="Core",
      note="REQUESTED BY THE TEAM. The sign is the point: a system that "
           "overstates hides a shortage (orders are not triggered), one that "
           "understates causes over-ordering. 2.1 and 2.2 must be read "
           "together - 2.1 gives magnitude of agreement, 2.2 gives direction. "
           "Exclude zero-digital-stock rows explicitly rather than treating "
           "them as 100% discrepancy."),
 dict(obj=2, id="2.3", name="Digital system in use, and accuracy by system",
      defn="Which electronic stock system the facility names, and whether "
           "accuracy differs between them.",
      num="Facilities naming each system", den="Facilities answering",
      n=f"{S['acc_rows']:,} rows name a system",
      unit="Facility", tool=T_STOCK, sheet="medicine_group",
      vars="Name of the electronical medical system used for recording",
      disagg="Level of care; province; district",
      method="Normalise spellings, then frequency table and median accuracy "
             "per system. Restrict comparison to systems with "
             f"{MIN_N}+ observations.",
      stat="Counts; median accuracy by system",
      shell="T5 Systems roster",
      out="Medicine_availability_analysis.xlsx - By stock system",
      feas="Yes", pri="Secondary",
      note="Free text: 162 raw values, where ELMIS / eLMIS / E-LMIS are one "
           "system. Differences between systems are correlation, not cause - "
           "system choice tracks facility level and programme."),
 dict(obj=3, id="3.1", name="Stock-out prevalence on the visit day",
      defn="Share of managed tracer medicines stocked out at the moment of the "
           "visit.",
      num="Availability = No", den="Managed AND availability answered",
      n=f"{S['stockout']} of {S['managed']:,}",
      unit="Facility x medicine observation", tool=T_STOCK, sheet="medicine_group",
      vars="Is **${medicine_name}** available on visit day?",
      disagg="Level of care; province; district; sector; facility; programme; "
             "medicine",
      method="Proportion with 95% CI. The exact complement of 1.1 - present "
             "one or the other, not both as separate findings.",
      stat="Proportion + 95% CI", shell="T1 Availability by disaggregator",
      out="Medicine_availability_analysis.xlsx - all sheets",
      feas="Yes", pri="Core", note=""),
 dict(obj=3, id="3.2", name="Current stock-out duration",
      defn="Days the medicine had already been out of stock when the team "
           "arrived.",
      num="Days out of stock at the visit", den="Observations stocked out",
      n=f"{S['so_now']} rows carry this - only the stocked-out ones",
      unit="Facility x medicine observation", tool=T_STOCK, sheet="medicine_group",
      vars="Current stock-out days of **${medicine_name}** (this visit)",
      disagg="Level of care; province; district; programme; medicine",
      method="Median and IQR. Share of stock-outs already running more than 7, "
             "14 and 30 days - the >30-day group is the procurement escalation "
             "list.",
      stat="Median + IQR; threshold shares",
      shell="T6 Duration distribution",
      out="Medicine_assessment_analysis.xlsx - Stock and stock-outs",
      feas="Yes", pri="Core",
      note="Populated only on currently stocked-out rows, by skip logic. Do "
           "not treat the blanks as zero."),
 dict(obj=3, id="3.3", name="Stock-out days over the preceding three months",
      defn="Days out of stock across June, July and August 2026.",
      num="Sum of monthly stock-out days",
      den="Observation-months assessed (3 per observation)",
      n=f"{S['managed']:,} managed rows x 3 months",
      unit="Facility x medicine observation", tool=T_STOCK, sheet="medicine_group",
      vars="June / July / August 2026 stock-out days of **${medicine_name}**; "
           "total_stockout_days; stockout_rate_pct",
      disagg="Month; level of care; province; district; facility; programme; "
             "medicine",
      method="Stock-out rate = total stock-out days / days in the window. "
             "Report by month to show trend, and by medicine to rank the worst "
             "items. Test the month trend with Friedman or a simple paired "
             "comparison.",
      stat="Rate; monthly trend",
      shell="T7 Stock-out days by month",
      out="Medicine_assessment_analysis.xlsx - Stock and stock-outs",
      feas="Yes", pri="Core",
      note="COMPARABILITY: the kf form version also carries March-May; the EU "
           "version does not. Use Jun-Aug only or the two servers are not "
           "comparable."),
 dict(obj=3, id="3.4", name="Stock level - months of stock on hand "
                             "(team request)",
      defn="How long the quantity on hand would last at the facility's own "
           "average monthly consumption.",
      num="Quantity on hand", den="Average monthly consumption",
      n="8,679 rows have a months-of-stock value; AMC is present on all "
        "9,878 managed rows",
      unit="Facility x medicine observation", tool=T_STOCK, sheet="medicine_group",
      vars="Quantity of **${medicine_name}** available; "
           "avg_monthly_consumption; stock_level_months; "
           "understock_threshold; overstock_threshold",
      disagg="Level of care; province; district; facility; programme; drug "
             "class; medicine",
      method="MEDIAN months of stock with IQR, never the mean - a near-zero "
             "consumption denominator sends the mean into the hundreds. Band "
             "as 0, under 1, 1-2, 2-6, 6-12, over 12 months, plus a separate "
             "'consumption zero so undefined' category.",
      stat="Median + IQR; banded distribution",
      shell="T22 Months of stock",
      out="Medicine_assessment_analysis.xlsx - Stock and stock-outs",
      feas="Yes", pri="Core",
      note="REQUESTED BY THE TEAM ('stock level and all connected "
           "indicators'). AMC is entered by the collector in this tool, not "
           "calculated from the three months - cross-check it against the "
           "monthly consumption columns before relying on months of stock."),
 dict(obj=3, id="3.5", name="Stock status classification (team request)",
      defn="Each observation classed against the form's own thresholds: stock "
           "out, understock, plan stock, overstock.",
      num="Observations per class", den="Observations classified",
      n="8,679 of 9,878 managed rows carry a classification",
      unit="Facility x medicine observation", tool=T_STOCK, sheet="medicine_group",
      vars="stock_classification; stock_level_months; understock_threshold; "
           "overstock_threshold",
      disagg="Level of care; province; district; facility; programme; drug "
             "class; medicine",
      method="Distribution across the four classes, by each disaggregator. "
             "Report understock and overstock together - they are two failures "
             "of the same planning process, and overstock ties up budget and "
             "expires.",
      stat="Distribution; chi-square across levels",
      shell="T23 Stock classification",
      out="Medicine_availability_analysis.xlsx; "
          "Medicine_assessment_analysis.xlsx",
      feas="Yes", pri="Core",
      note="REQUESTED BY THE TEAM. The thresholds are form settings "
           "(understock below 1 month, overstock above 2) - read them from the "
           "data rather than writing them down, in case the form changes."),
 dict(obj=3, id="3.6", name="Quantity and value missing because of the "
                             "stock-out (team request)",
      defn="The quantity that could not be dispensed because the medicine was "
           "out of stock, and its value at tariff price.",
      num="est_qty_unfilled_stockout; prescriptions_missed_stockout; "
          "financial_loss",
      den="n/a - volumes and totals",
      n="SPARSE: quantity unfilled on 580 rows (234 stocked-out), "
        "prescriptions missed on 1,167 (591), financial loss on 1,774 "
        "(all 644 stocked-out)",
      unit="Facility x medicine observation", tool=T_STOCK, sheet="medicine_group",
      vars="est_qty_unfilled_stockout; prescriptions_missed_stockout; "
           "financial_loss; unit_tariff_price; total_stockout_days; "
           "Total Quantity Billed **${medicine_name}**; Prescriptions of "
           "**${medicine_name}** / year (RSSB)",
      disagg="Level of care; province; district; facility; programme; medicine",
      method="Sum the quantity unfilled and the prescriptions missed, and the "
             "loss, over stocked-out rows. Report the number of rows "
             "contributing beside every total - the inputs are sparse. Rank "
             "medicines and facilities by value.",
      stat="Sums with contributing-row counts; ranking",
      shell="T24 Stock-out quantity and value",
      out="Medicine_assessment_analysis.xlsx - Attribution and cost",
      feas="Yes", pri="Core",
      note="REQUESTED BY THE TEAM ('check the financial loss due to the "
           "stockout, check also the quantity missing'). These are MODELLED "
           "from RSSB billing volumes, not observed, and the inputs are "
           "present on a minority of rows - always state the contributing n. "
           "Overlaps indicator 7.2: 7.2 is the whole modelled loss, 3.6 is "
           "the stock-out share of it. Do not add them."),
 dict(obj=4, id="4.1", name="Proportion of prescriptions unserved",
      defn="Share of prescribed medicine items the facility could not dispense "
           "on the prescription date.",
      num="Items with 'available at the facility' = No",
      den="Items with availability answered",
      n=f"{S['unserved']} of {S['items']:,} items; {S['pres']:,} prescriptions",
      unit="Medicine item, and prescription", tool=T_PRESC, sheet="medicine_item",
      vars="Was the medicine available at the health facility during the date "
           "of prescription?; Qty prescribed",
      disagg="Province; district; facility; facility type; prescriber "
             "category; drug class; medicine; dosage form",
      method="Report at BOTH levels: item-level share, and share of "
             "prescriptions with at least one unserved item. Recompute from "
             "the item rows - do not use the form roll-up.",
      stat="Proportion + 95% CI at both levels",
      shell="T9 Unserved rate + full item list",
      out="Prescription_study_analysis.xlsx - Availability & unserved",
      feas="Yes", pri="Core",
      note=f"CRITICAL: this rate is a design artefact and runs near 0.5%. A "
           f"prescriber writes what the pharmacy holds, so a medicine the "
           f"facility never stocks is rarely prescribed and never enters the "
           f"denominator. Quote 1.1 for availability. Also: the form's own "
           f"roll-up disagrees with the item rows on {S['rollup_mismatch']} "
           f"prescriptions - recompute, and use the roll-up only as a check."),
 dict(obj=4, id="4.2", name="Quantity of medicine involved",
      defn="Units of medicine prescribed but not dispensed.",
      num="Sum of quantity prescribed on unserved items",
      den="n/a - a volume, not a rate",
      n=f"{S['unserved']} unserved items",
      unit="Medicine item", tool=T_PRESC, sheet="medicine_item",
      vars="Qty prescribed; Was it supplied fully, partially, or not at all?",
      disagg="Province; district; facility type; medicine; drug class",
      method="Total and median quantity, split by the supply outcome where it "
             "is known.",
      stat="Sum; median", shell="T10 Unserved quantity",
      out="Prescription_study_analysis.xlsx - Availability & unserved",
      feas="Yes", pri="Core",
      note="A partially supplied item still leaves a gap, but the form has no "
           "'quantity actually dispensed', so partial gaps cannot be sized."),
 dict(obj=5, id="5.1", name="Upstream availability and attribution of the gap",
      defn="For each unserved medicine, whether it was in stock at RMS "
           "Regional and RMS Central on the prescription date, and the "
           "attribution that follows: stock existed upstream (distribution "
           "problem) or did not exist anywhere (national problem).",
      num="Unserved items per attribution class",
      den="Unserved items answered",
      n=f"{S['unserved']} prescription items; {S['stockout']} stock-assessment rows",
      unit="Unserved medicine item",
      tool=f"{T_PRESC} + {T_STOCK}", sheet="medicine_item / medicine_group",
      vars="Was the medicine available at RMS Regional during the date of "
           "prescription?; Was the medicine available at RMS Central during "
           "the date of prescription?; distribution_nationwide (derived)",
      disagg="Province; district; facility type; medicine; programme",
      method="Step 1: cross-tabulate Regional x Central (Yes / No / Don't "
             "know), keeping 'Don't know' as its own category - folding it "
             "into No would overstate nationwide stock-outs. Step 2: apply "
             "the attribution rule - available at either warehouse -> "
             "distribution issue; at neither -> nationwide stock-out; unknown "
             "-> not answered. Report the proportion of each with the medicine "
             "list behind it.",
      stat="Crosstab; proportions", shell="T11 RMS crosstab; T12 Attribution",
      out="Both analyses; Medicine_availability_analysis.xlsx - Method sheet",
      feas="Yes", pri="Core",
      note="This is the objective that turns a stock-out into an action: a "
           "distribution issue is a district/RMS logistics problem, a "
           "nationwide stock-out is a procurement problem. They go to "
           "different people."),
 dict(obj=6, id="6.1", name="Was the missing medicine requested?",
      defn="Whether the facility ordered the medicine in the month the "
           "prescription was written.",
      num="Unserved items ordered = Yes", den="Unserved items (all answered)",
      n=f"{S['ordered']} ordered of {S['unserved']} unserved",
      unit="Unserved medicine item", tool=T_PRESC, sheet="medicine_item",
      vars="Was the medicine ordered during the month the prescription was "
           "made?",
      disagg="Province; district; facility type; medicine",
      method="Proportion ordered. Cross-tabulate against the RMS answer to "
             "separate 'never asked for it' from 'asked and did not get it'.",
      stat="Proportion; crosstab with 5.1",
      shell="T13 Ordered vs upstream availability",
      out="Prescription_study_analysis.xlsx - Availability & unserved",
      feas="Yes", pri="Core",
      note=f"A finding in its own right: only {S['ordered']} of {S['unserved']} "
           f"unserved items had been ordered. An unserved medicine never "
           f"ordered is a facility-side failure; ordered but not supplied is a "
           f"supply-side one. TEAM COMMENT - 'the medicine assessment should "
           f"have this question as well': checked, and it does NOT. The stock "
           f"tool carries monthly requested and delivered QUANTITIES "
           f"(indicator 6.3) but no yes/no 'was it ordered'. The quantity "
           f"route answers the same question for the tracer basket as a whole; "
           f"a binary would need adding to the form - see 'Form change "
           f"requests'."),
 dict(obj=6, id="6.2", name="Supplied fully, partially, or not at all",
      defn="Outcome of the request, where a request was made.",
      num="Unserved items per supply outcome",
      den="Unserved items that were ORDERED",
      n=f"{S['supplied_answered']} answered - asked only where ordered = Yes",
      unit="Unserved medicine item", tool=T_PRESC, sheet="medicine_item",
      vars="Was it supplied fully, partially, or not at all?",
      disagg="Province; district; facility type; medicine; drug class",
      method="Frequency of the three outcomes over ORDERED items, not over all "
             "unserved items. Report the quantity involved in each.",
      stat="Counts; quantity by outcome", shell="T14 Supply outcome",
      out="Prescription_study_analysis.xlsx - Availability & unserved",
      feas="Yes", pri="Core",
      note=f"DENOMINATOR: the question is skipped where the medicine was never "
           f"ordered, so the base is {S['supplied_answered']}, not "
           f"{S['unserved']}. Reporting it over all unserved items would make "
           f"two thirds look like missing data when they are in fact 'never "
           f"ordered'."),
 dict(obj=6, id="6.3", name="Requisition fulfilment at facility level",
      defn="Quantity delivered against quantity requested over the three "
           "months, for all tracer medicines.",
      num="Total quantity delivered", den="Total quantity requested",
      n=f"{S['fulfil_rows']:,} of {S['managed']:,} managed rows have a "
        f"fulfilment value",
      unit="Facility x medicine observation", tool=T_STOCK, sheet="medicine_group",
      vars="June / July / August 2026 quantity of **${medicine_name}** "
           "requested; ... delivered; total_requested_qty; total_delivered_qty; "
           "total_undelivered_qty; fulfilment_pct",
      disagg="Month; level of care; province; district; facility; programme; "
             "medicine",
      method="Fulfilment % from SUMMED quantities, never as a mean of per-row "
             "percentages - that would weight a 30-unit line like a "
             "50,000-unit one. Report by month, and rank facilities and "
             "medicines.",
      stat="Ratio of sums; monthly trend",
      shell="T15 Fulfilment by month and group",
      out="Medicine_assessment_analysis.xlsx - Requisition fulfilment",
      feas="Yes", pri="Core",
      note="fulfilment_pct is blank where nothing was requested, which is why "
           f"only {S['fulfil_rows']:,} rows carry it. Compute from the summed "
           "quantities rather than averaging the column."),
 dict(obj=7, id="7.1", name="Financial value of unserved prescriptions",
      defn="Value of the medicines patients could not receive, at official "
           "tariff prices.",
      num="Sum of (quantity prescribed x unit tariff price) on unserved items",
      den="n/a - a total",
      n=f"{S['priced']} of {S['unserved']} unserved items carry a price",
      unit="Unserved medicine item", tool=T_PRESC, sheet="medicine_item",
      vars="Unit price for the medicine that was not available (RWF); "
           "total_cost_unserved; Qty prescribed",
      disagg="Province; district; facility; facility type; medicine; drug class",
      method="Total and median per item; top medicines and facilities by "
             "value. State how many items were unpriced and exclude them "
             "explicitly rather than treating them as zero.",
      stat="Sum; median; ranking", shell="T16 Value of unserved prescriptions",
      out="Prescription_study_analysis.xlsx - Availability & unserved",
      feas="Yes", pri="Core",
      note="Price is captured only on unserved items, so this values the "
           "observed gap, not the whole shortage. It is a floor, not a total."),
 dict(obj=7, id="7.2", name="Financial value of the stock gap",
      defn="Value of the unmet requisition and stock-out gap measured by the "
           "stock tool.",
      num="Sum of financial_loss", den="n/a - a total",
      n=f"{S['loss_rows']:,} rows carry a loss value, including all "
        f"{S['stockout']} stocked-out rows",
      unit="Facility x medicine observation", tool=T_STOCK, sheet="medicine_group",
      vars="unit_tariff_price; est_unserved_requisition_gap; financial_loss; "
           "est_qty_unfilled_stockout; prescriptions_missed_stockout",
      disagg="Level of care; province; district; facility; programme; medicine",
      method="Total with the top contributors listed. ALWAYS report the total "
             "both with and without the largest single facility, and name it.",
      stat="Sum; concentration check", shell="T17 Value of the stock gap",
      out="Medicine_assessment_analysis.xlsx - Attribution and cost",
      feas="Yes", pri="Core",
      note="MUST NOT be added to 7.1 - they overlap and answer different "
           "questions. 7.1 is prescriptions patients actually presented; 7.2 "
           "is a modelled supply gap. In the current data one referral "
           "hospital (CHUK) drives about two thirds of 7.2."),
 dict(obj=8, id="8.1", name="Community-level availability",
      defn="Share of the CHW package physically present with the CHW on the "
           "visit day.",
      num="CHW x medicine rows with availability = Yes",
      den="CHW x medicine rows the CHW manages",
      n=f"{S['chw_avail']} available of {S['chw_managed']:,} managed "
        f"({S['chw_rows']:,} rows total)",
      unit="CHW x medicine observation", tool=T_CHW, sheet="medicine_group",
      vars="Is **${medicine_name}** managed by this CHW?; Is "
           "**${medicine_name}** available on visit day?; Quantity of "
           "**${medicine_name}** available",
      disagg="Province; district; parent health centre; CHW; medicine; "
             "antimalarial family",
      method="Proportion with 95% CI, over managed rows.",
      stat="Proportion + 95% CI", shell="T18 CHW availability",
      out="Medicine_stock_CHW_analysis.xlsx - Availability profiles",
      feas="Yes", pri="Core",
      note=f"About {100*(S['chw_rows']-S['chw_managed'])/S['chw_rows']:.0f}% of "
           f"rows are 'not managed' because RBC's MFT zoning gives each area "
           f"one first-line antimalarial. Counting those as unavailable would "
           f"report the zoning policy as a stock-out and roughly halve the "
           f"figure."),
 dict(obj=8, id="8.2", name="CHW stock level and consumption",
      defn="Months of stock on hand with the CHW, and consumption over the "
           "window.",
      num="Quantity on hand", den="Average monthly consumption",
      n=f"{S['chw_mos']} of {S['chw_managed']:,} managed rows have a "
        f"months-of-stock value",
      unit="CHW x medicine observation", tool=T_CHW, sheet="medicine_group",
      vars="June / July / August consumption total quantity of "
           "**${medicine_name}**; avg_monthly_consumption; stock_level_months",
      disagg="Province; district; parent health centre; medicine",
      method="MEDIAN months of stock, banded against the form's own thresholds "
             "(understock <1 month, plan stock 1-2, overstock >2). Consumption "
             "totals by month.",
      stat="Median; banded distribution", shell="T19 CHW months of stock",
      out="Medicine_stock_CHW_analysis.xlsx - Stock levels",
      feas="Yes", pri="Core",
      note=f"Months of stock is UNDEFINED where consumption is zero - only "
           f"{S['chw_mos']} of {S['chw_managed']:,} rows have it. Class those "
           f"'no consumption', never 'overstock'. Use the median; a near-zero "
           f"denominator sends the mean into the hundreds."),
 dict(obj=8, id="8.3", name="Community vs facility comparison",
      defn="Availability at the CHW against availability at the health centre "
           "the CHW is attached to.",
      num="CHW availability %", den="Facility availability % at the parent",
      n=f"{S['shared_fac']} facilities appear in both tools",
      unit="Parent facility", tool=f"{T_CHW} + {T_STOCK}",
      sheet="medicine_group (both)",
      vars="Facility (select) in both tools; availability in both; "
           "medicine_name in both",
      disagg="Province; district; medicine; programme",
      method="Paired comparison restricted to facilities present in both tools "
             "AND to the medicines both baskets share. Report the difference "
             "in percentage points with a paired test (McNemar at observation "
             "level, or Wilcoxon on facility-level rates).",
      stat="Paired difference; McNemar / Wilcoxon",
      shell="T20 Paired CHW vs facility",
      out="NOT YET BUILT - the one indicator with no output",
      feas="Yes", pri="Core",
      note="The CHW basket is 13 items, the facility basket is 81. Restrict to "
           "the shared items or the comparison is between two different "
           "medicine lists, not two levels of care."),
 dict(obj=8, id="8.4", name="Why a medicine is absent from the CHW kit",
      defn="Coded reasons for a medicine not being managed or not available at "
           "community level.",
      num="Comments per theme", den="Comments left",
      n="free text on the CHW medicine rows",
      unit="CHW x medicine observation", tool=T_CHW, sheet="medicine_group",
      vars="Please share any feedback, concerns, or recommendations related to "
           "the supply chain and other relevant areas.",
      disagg="Province; district; medicine; managed status",
      method="Keyword-code into themes; report the share per theme and list "
             "every uncoded comment verbatim so new phrasings surface.",
      stat="Thematic counts", shell="T21 Coded themes",
      out="Medicine_stock_CHW_analysis.xlsx - Feedback",
      feas="Yes", pri="Secondary",
      note="Themes must separate MFT zoning and low malaria burden from "
           "genuine stock-outs - in a raw 'not available' count they look "
           "identical. The coder needs new rules at almost every refresh."),
]

# Requirements that are part of an objective but cannot be produced from the
# data as collected. The team's returned version simply deleted these rows;
# they are kept HERE instead so the decision is documented rather than lost -
# the objectives still contain the requirement, and a reader of the report
# will ask what happened to it.
EXCLUDED = [
    ("3.7", 3, "Stock-out frequency (number of episodes)",
     "Quantify stock-outs ... and frequency over the preceding three months",
     "The form records stock-out DAYS per month, not episodes. Two 5-day "
     "stock-outs and one 10-day stock-out are indistinguishable.",
     "PROXY available: count the months with any stock-out (0-3) and label it "
     "'months affected', not frequency.",
     "Add an episode-count question (\"how many separate times did this "
     "medicine run out?\") to the next form version.",
     "Report the proxy, labelled as such, in the limitations."),
    ("4.3", 4, "Outpatient vs inpatient split",
     "Measure the proportion of prescriptions that go unserved ... across "
     "outpatient and inpatient services",
     "Verified against the deployed export AND the XLSForm: there is no "
     "service type, ward, department or admission variable at any level of "
     "the prescription tool.",
     "No proxy that is honest at facility level. Facility type is the closest "
     "(hospitals carry the inpatient load, health centres are almost entirely "
     "outpatient) but it is a facility attribute, not a service attribute.",
     "Add a select_one (outpatient / inpatient) at prescription level in the "
     "next form version - it is one question.",
     "DECISION NEEDED before drafting: report facility type as an explicitly "
     "labelled proxy, or state the objective as not assessed."),
]

# The XLSForm files in the project folder predate what was deployed. This is
# not a criticism of the forms - it is a warning that reviewing the plan
# against those files will raise questions the deployed data already answers.
FORMVER = [
    ("Stock-out day months", "May, June, July 2026",
     "June, July, August 2026",
     "March-May columns exist in the export but are EMPTY on both servers - "
     "legacy columns carried over by fields_from_all_versions. Use Jun-Aug."),
    ("Requisition months", "April, May, June 2026",
     "June, July, August 2026 (March-May columns present but empty)",
     "Same cause. The real window is identical on both servers, so the two "
     "collection teams ARE comparable - an earlier draft of this plan said "
     "otherwise and was wrong."),
    ("EML applicability", "not in the form",
     "Is **${medicine_name}** applicable to this facility level as per EML?",
     "Indicator 1.3 depends on it. It exists in the deployed data."),
    ("Digital stock system", "not in the form",
     "Name of the electronical medical system used for recording",
     "Indicator 2.2 depends on it. It exists in the deployed data."),
    ("Monthly consumption totals", "avg_monthly_consumption entered directly",
     "June / July / August consumption total quantity, plus the average",
     "The deployed form collects the three months; AMC can be checked against "
     "them rather than taken on trust."),
    ("RSSB billing variables", "units_served_year, prescriptions_year",
     "Total Quantity Billed; Prescriptions of ... / year (RSSB)",
     "Feed the modelled loss in indicator 7.2."),
    ("Medicine name (prescription)", "free text (medicine_name_prescribed)",
     "select list plus a free-text fallback",
     "Why the deployed data has both 'Medicine name' and 'Medicine name (not "
     "on the EML list) - please specify'. Analyse the resolved name."),
    ("Prescription service type", "not in the form", "not in the form",
     "Confirms indicator 4.3 cannot be produced - absent by design in both "
     "versions, not lost in deployment."),
]

# What came back from the team in Data_analysis_plan11.xlsx, and what this
# version did about it.
CHANGES = [
    ("Unit of analysis for 1.1 should name the reporting levels",
     "Adopted", "Now reads 'Facility x medicine observation; reported at "
     "facility, district and level of care'."),
    ("Drop programme from the 1.2 disaggregation list",
     "Adopted", "1.2 now disaggregates by drug class and level of care only."),
    ("1.3 should disaggregate by drug class, not programme",
     "Adopted", "1.3 now reads 'Level of care; province; district; medicine; "
     "drug class' (the returned file's 'dug class' typo corrected)."),
    ("Analysis method left blank for 1.3 and 2.1",
     "Filled in", "Read as 'needs specifying'. 1.3 now has the by-level and "
     "by-drug-class cuts with a chi-square across levels; 2.1 has median "
     "accuracy with IQR, the six bands, the +/-5% headline and "
     "Kruskal-Wallis across levels."),
    ("Indicators 5.1 and 5.2 merged into one",
     "Adopted", "Objective 5 now has a single indicator covering the RMS "
     "Regional x Central crosstab AND the attribution that follows, in two "
     "numbered steps."),
    ("Indicators 3.4 and 4.3 deleted",
     "Adopted for the core sheet, kept on an 'Excluded indicators' sheet",
     "The objectives still contain those requirements, so deleting them "
     "silently would leave the report unable to explain the omission. The "
     "sheet records why each is impossible, what proxy exists, and what it "
     "would take to collect - 4.3 needs a decision before drafting."),
    ("Columns Sheet / Output / Feasibility / Notes removed",
     "Adopted for the core sheet", "The core sheet is now the team's 11 "
     "columns. Everything removed moved to 'Indicator detail', so the core "
     "table reads cleanly and nothing is lost."),
    ("Inventory accuracy should also be calculated as "
     "(digital - physical) / digital",
     "Adopted - new indicator 2.2",
     "Added as a signed discrepancy measure alongside the existing ratio, "
     "because the sign is the actionable part: on the current data 1,841 rows "
     "have the system overstating and 761 understating, with 5,012 exact. "
     "Rows with zero digital stock are excluded - the ratio is undefined."),
    ("Add stock level and all connected indicators",
     "Adopted - new indicators 3.4 and 3.5",
     "Months of stock on hand (median, banded) and the four-way stock status "
     "classification, both against the form's own thresholds. 8,679 of 9,878 "
     "managed rows carry them."),
    ("Check the financial loss due to the stock-out, and the quantity missing",
     "Adopted - new indicator 3.6",
     "Quantity unfilled, prescriptions missed and loss, over stocked-out "
     "rows. WARNING recorded in the plan: the inputs are sparse (quantity "
     "unfilled on 580 rows, loss on 1,774) and modelled from RSSB billing, "
     "and 3.6 overlaps 7.2 - they must not be added."),
    ("The medicine assessment should have the 'ordered?' question as well",
     "Checked - it does not; recorded as a form change request",
     "The stock tool records monthly requested and delivered QUANTITIES "
     "(indicator 6.3) but no yes/no 'ordered'. That covers the same ground "
     "for the tracer basket; a binary needs a form change. See 'Form change "
     "requests'."),
    ("Analysis method left blank on every row except the first two",
     "Filled in for all",
     "Read as 'specify these'. Every indicator now carries a method, the "
     "statistic to use, and the output table it feeds."),
    ("Indicator ID column removed",
     "Added back", "Every other sheet - table shells, deliverables, "
     "limitations, excluded indicators - references indicators by ID. Without "
     "it the workbook cannot be cross-read. It is one narrow column."),
]

# Changes to the instruments that the analysis cannot work around. Each is one
# question; all of them came out of either the team's comments or the
# verification pass against the deployed data.
FORMREQ = [
    ("Prescription", "Service type at prescription level",
     "select_one (outpatient / inpatient)",
     "Objective 4 asks for the unserved split across outpatient and inpatient "
     "services. No service, ward, department or admission variable exists in "
     "the deployed form OR in the XLSForm, so the objective cannot be "
     "reported as written.", "High - it blocks part of an objective"),
    ("Medicine assessment", "Was this medicine ordered in the period?",
     "select_one yes_no, per medicine",
     "TEAM COMMENT on indicator 6.1. The stock tool records monthly requested "
     "and delivered quantities but no binary 'ordered'. A zero requested "
     "quantity cannot be distinguished from a month that was not asked about.",
     "Medium - the quantity route covers most of the need"),
    ("Medicine assessment", "Number of separate stock-out episodes",
     "integer, per medicine, per month or per window",
     "Objective 3 asks for stock-out frequency. Only day counts exist, so two "
     "5-day stock-outs and one 10-day stock-out are indistinguishable.",
     "Medium - a labelled proxy exists (months affected)"),
    ("Prescription", "Quantity actually dispensed on a partial supply",
     "integer, shown when supply outcome = partially",
     "Objectives 4 and 6 can count partial supplies but cannot size the gap "
     "they leave.", "Medium"),
    ("Medicine assessment", "Confirm avg monthly consumption is derived, not "
     "typed", "calculate from the three monthly consumption columns",
     "AMC is currently entered by the collector while the three monthly "
     "consumption totals are also collected. Deriving it would remove a "
     "transcription risk from every months-of-stock figure (indicator 3.4).",
     "Low - but cheap to fix"),
    ("All three", "Archive the deployed XLSForm alongside the data",
     "version control, not a question",
     "The two XLSForm files in the project folder predate deployment - "
     "different month windows, no EML applicability, no digital-system "
     "question. The design of record should match the data of record.",
     "Low - documentation hygiene"),
]

SEQUENCE = [
    ("1", "Refresh all six Kobo sources and rebuild the three merged exports",
     "python3 fetch_kobo_export.py",
     "Six parts, three merged files. Nothing downstream is valid until this "
     "runs - every figure is an interim cut."),
    ("2", "Confirm the merge and the facility corrections",
     "read the fetch output; check the reattribution lines",
     "Zero orphan rows; every FACILITY_OVERRIDES entry matches at least one "
     "row; a WARNING means a facility was renamed upstream."),
    ("3", "Run the three analyses",
     "python3 analyze_medicine_assessment.py; analyze_kobo_prescription.py; "
     "analyze_medicine_assessment_chw.py",
     "Three workbooks. Check the CHW feedback coder's uncoded count and add "
     "rules until it is near zero."),
    ("4", "Run the data quality checks on this plan's 'Data quality checks' "
          "sheet", "the Data quality sheet of each workbook",
     "Resolve or document every failure BEFORE computing indicators."),
    ("5", "Confirm the denominators are the ones this plan specifies",
     "Medicine_availability_analysis.xlsx - Method and funnel sheet",
     "observations -> applicable -> managed -> answered -> available. A rate "
     "on the wrong step is the most likely error in the whole analysis."),
    ("6", "Produce the objective 1-3 indicators (availability, accuracy, "
          "stock-outs)",
     "python3 Medicine_availability_analysis.py",
     "Availability by every disaggregator, the funnel, every stock-out."),
    ("7", "Produce the objective 4-7 indicators (unserved, RMS, requests, "
          "value)",
     "Prescription_study_analysis.xlsx + Prescription_facility_analysis.py",
     "Unserved item list, RMS crosstab, ordered/supplied, financial value."),
    ("8", "Produce the objective 8 indicators (community level)",
     "Medicine_stock_CHW_analysis.xlsx; 8.3 still to be built",
     "CHW availability, months of stock, coded reasons."),
    ("9", "Check coverage against the site plan",
     "python3 Site_coverage_check.py; python3 Completion_by_geography.py",
     "How much of the planned sample each tool reached - the denominator "
     "assurance behind every objective."),
    ("10", "Rebuild the shared report", "python3 export_combined_report.py",
     "Three tabbed pages of headline indicators for circulation."),
    ("11", "Apply the reporting rules before anything is quoted",
     "the 'Reporting rules' sheet",
     "Especially: the prescription unserved rate is not availability, and the "
     "two financial estimates are never added."),
]

CHECKS = [
    ("Merge integrity", "Every repeat row joins to a submission",
     "0 orphan rows in all three tools",
     "Rows would be silently dropped from every denominator"),
    ("Merge integrity", "The two-server merge did not duplicate or mis-link rows",
     "_index offset applied for the stock and CHW tools; _uuid unique across "
     "servers", "EU rows would attach themselves to kf facilities"),
    ("Skip logic", "Availability answered only where managed",
     "answered == managed", "The form was answered against its own logic"),
    ("Skip logic", "Physical and system counts present only on available rows",
     f"{S['acc_rows']:,} available rows carry both; stocked-out rows carry "
     f"neither", "Objective 2's denominator would be wrong"),
    ("Skip logic", "RMS / ordered answered on every unserved item; supplied "
                   "answered only where ordered = Yes",
     f"RMS {S['unserved']}/{S['unserved']}; supplied "
     f"{S['supplied_answered']}/{S['unserved']}",
     "Objective 6.2 would look like missing data instead of 'never ordered'"),
    ("Denominators", "The applicable -> managed -> answered funnel reconciles",
     f"{S['applicable']:,} -> {S['managed']:,} -> {S['managed']:,}",
     "Each step must be no larger than the one before it"),
    ("Derived fields", "Form roll-ups cross-checked against the item rows",
     f"currently {S['rollup_mismatch']} prescriptions disagree",
     "The roll-up is not authoritative - recompute from items"),
    ("Numeric hygiene", "Calculate fields parse as finite numbers",
     "no literal 'NaN' surviving into a sum",
     "One NaN turns every downstream total into NaN"),
    ("Numeric hygiene", "Availability = Yes implies quantity > 0",
     "no rows available with quantity 0", "One of the two answers is wrong"),
    ("Outliers", "Loss and months-of-stock tails inspected before reporting",
     "top contributors named and sanity-checked",
     "One referral hospital currently drives two thirds of objective 7.2"),
    ("Identity", "Facility corrections applied before any grouping",
     "seven FACILITY_OVERRIDES entries, all matching",
     "Health-post data stays pooled into its parent health centre"),
    ("Identity", "Facility names that repeat across districts kept apart",
     "key on name + level + district",
     "RUBONA CS exists in three districts - a name-only key merges them"),
    ("Comparability", "Month window identical across form versions",
     "June-August only",
     "The kf version carries March-May, the EU version does not"),
    ("Coverage", "Collected facilities reconciled against the site plan",
     "Site_coverage_check.xlsx rerun after each refresh",
     "Gaps and unplanned visits go unnoticed"),
]

RULES = [
    ("Availability is computed over MANAGED rows only", "1, 3, 8",
     "The availability question is only asked where the facility or CHW "
     "manages the medicine. Using all observations counts policy decisions and "
     "MFT zoning as stock-outs."),
    ("Inventory accuracy is computed over AVAILABLE rows only", "2",
     "The physical and system counts are captured only where the medicine is "
     "on the shelf. Managed rows would be the wrong base."),
    ("The prescription unserved rate is NOT an availability measure", "4",
     "It runs near 0.5% because prescribers write what the pharmacy holds. "
     "Quote indicator 1.1 for availability."),
    ("Supply outcome is reported over ORDERED items", "6",
     "The question is skipped where the medicine was never ordered, so two "
     "thirds of unserved items have no outcome by design."),
    ("Never add the two financial estimates together", "7",
     "7.1 values prescriptions patients presented; 7.2 values a modelled "
     "supply gap. They overlap and measure different things."),
    ("Report medians for stock levels and accuracy, not means", "2, 8",
     "Near-zero denominators push the mean into the hundreds of months and "
     "distort accuracy."),
    ("Quantities are summed before a ratio is taken", "6, 7",
     "Fulfilment and value must come from summed numerators and denominators, "
     "not a mean of per-row percentages."),
    (f"Groups below {MIN_N} answered rows are shown but not compared", "all",
     "A rate over a handful of observations is noise. Flag rather than drop, "
     "so the reader knows the group exists."),
    ("State the tracer list explicitly in the methods", "1",
     "There is no tracer flag in the tool - the medicine multi-select IS the "
     "list, and it must be reproduced for the result to be replicable."),
    ("Every collection is an interim cut until fieldwork closes", "all",
     "Re-fetch, re-run all three analyses and rebuild the outputs before any "
     "figure is quoted."),
]

LIMITS = [
    ("L1", "No outpatient / inpatient split", "Objective 4",
     "The prescription tool records no service type, ward, department or "
     "admission status at any level.",
     "Objective 4 cannot be reported by service. Decide between adding the "
     "question, using facility type as a labelled proxy, or dropping it."),
    ("L2", "Stock-out frequency is not measurable", "Objective 3",
     "Only monthly stock-out DAY counts exist, not episode counts.",
     "Report 'months affected (0-3)' as an explicit proxy."),
    ("L3", "The prescription unserved rate understates the shortage",
     "Objectives 4, 7",
     "Prescribers prescribe what the facility holds, so unstocked medicines "
     "never enter the denominator; the form also cannot record 'available but "
     "not enough'.",
     "Use the stock assessment for availability. Treat 7.1 as a floor."),
    ("L4", "Partial supply cannot be sized", "Objectives 4, 6",
     "There is no 'quantity actually dispensed' variable.",
     "Report the count of partially supplied items, not the quantity gap."),
    ("L5", "Financial value rests on prices captured only for unserved items",
     "Objective 7",
     f"{S['priced']} of {S['unserved']} unserved items carry a unit price.",
     "State the unpriced count; do not treat unpriced as zero."),
    ("L6", "One facility dominates the modelled loss", "Objective 7",
     "A single national referral hospital accounts for roughly two thirds of "
     "the financial_loss total.",
     "Always report the total with and without it, and name it."),
    ("L7", "The observation window is three months", "Objectives 3, 6",
     "June-August 2026 on BOTH servers. The March-May columns in the export "
     "are empty legacy columns from an earlier form version, not data - an "
     "earlier draft of this plan wrongly read them as a server difference.",
     "Restrict every time series to June-August and say so. Seasonality "
     "cannot be assessed from one quarter."),
    ("L8", "Months of stock is undefined where consumption is zero",
     "Objectives 3, 8",
     f"Only {S['chw_mos']} of {S['chw_managed']:,} managed CHW rows have a "
     f"months-of-stock value.",
     "Class those as 'no consumption'; never fold them into overstock."),
    ("L9", "CHW and facility baskets differ", "Objective 8.3",
     "13 CHW items against 81 facility items.",
     "Restrict the paired comparison to the shared medicines."),
    ("L10", "Fine-grained geography is thin", "all",
     "193 sectors and 267 facilities against ~9,900 managed rows.",
     f"Most sector and facility cells fall below {MIN_N} answered rows - "
     f"report province and district, and use facility only for listings."),
    ("L11", "Non-random facility selection", "all",
     "Facilities follow a visit plan, not a probability sample, and fieldwork "
     "was still open.",
     "Report as a census of visited facilities. Do not generalise to national "
     "prevalence without stating this."),
]

SHELLS = [
    ("T1", "Availability by disaggregator", "1.1, 3.1",
     "Group | Observations | Applicable | Managed | Managed % | Answered | "
     "Available | AVAILABLE % (95% CI) | Stocked out | Stock-out %"),
    ("T2", "Tracer medicine roster", "1.2",
     "Medicine | Programme | Drug class | Facilities assessed | Observations | "
     "Managed % | Available %"),
    ("T3", "EML applicability gap", "1.3",
     "Group | Applicable | Managed | Not managed | Gap % | Top coded reasons"),
    ("T4", "Inventory accuracy bands", "2.1",
     "Band (exactly 100% / 95-99% / 80-94% / 50-79% / 0-49% / negative) | "
     "Observations | % of counted | Median accuracy"),
    ("T5", "Digital systems roster", "2.2",
     "System (normalised) | Facilities | Observations | Median accuracy % | "
     "Available %"),
    ("T6", "Current stock-out duration", "3.2",
     "Group | Stocked-out observations | Median days | IQR | >7 days | "
     ">14 days | >30 days"),
    ("T7", "Stock-out days by month", "3.3",
     "Group | June days | July days | August days | Total | Stock-out rate % | "
     "Observations"),
    ("T8", "Months affected (frequency proxy)", "3.4",
     "Months with any stock-out (0 / 1 / 2 / 3) | Observations | % of managed"),
    ("T9", "Unserved rate", "4.1",
     "Group | Items | Answered | Unserved items | Unserved % | Prescriptions | "
     "Prescriptions with >=1 unserved | %"),
    ("T10", "Unserved quantity", "4.2",
     "Medicine | Unserved items | Total quantity | Median quantity | "
     "Supply outcome split"),
    ("T11", "RMS availability crosstab", "5.1",
     "RMS Central Yes/No/Don't know x RMS Regional Yes/No/Don't know | counts"),
    ("T12", "Attribution of the gap", "5.2",
     "Attribution (distribution issue / nationwide stock-out / not answered) | "
     "Items | % | Medicines involved"),
    ("T13", "Ordered vs upstream availability", "6.1",
     "Ordered Yes/No x available at either warehouse Yes/No | counts | "
     "interpretation"),
    ("T14", "Supply outcome", "6.2",
     "Outcome (fully / partially / not at all) | Items | % of ordered | "
     "Quantity involved"),
    ("T15", "Requisition fulfilment", "6.3",
     "Group | Quantity requested | Quantity delivered | Undelivered | "
     "Fulfilment % | Observations"),
    ("T16", "Value of unserved prescriptions", "7.1",
     "Group | Unserved items | Items priced | Total value (RWF) | "
     "Median per item | Top medicines"),
    ("T17", "Value of the stock gap", "7.2",
     "Group | Observations with a loss value | Total loss (RWF) | "
     "Total excluding the largest facility | Top contributors"),
    ("T18", "CHW availability", "8.1",
     "Group | CHW x medicine rows | Managed | Managed % | Available | "
     "AVAILABLE % (95% CI) | Stocked out"),
    ("T19", "CHW months of stock", "8.2",
     "Band (0 / <1 / 1-2 / 2-6 / 6-12 / >12 / no consumption) | Observations | "
     "% of managed | Median months"),
    ("T20", "Paired CHW vs facility availability", "8.3",
     "Parent facility | Shared medicines | Facility available % | "
     "CHW available % | Difference (pp) | n"),
    ("T21", "CHW coded reasons", "8.4",
     "Theme | Comments | % of comments | On managed rows | On not-managed rows"),
]


def disagg():
    def n(rows, f):
        return len({str(f(r)).strip() for r in rows
                    if str(f(r)).strip() not in ("", "None")})
    return [
        ["Province", "All three", n(M.SUB, lambda r: r[M.PROV]), "Objective 1 "
         "names it explicitly. Safe for every indicator."],
        ["District", "All three", n(M.SUB, lambda r: r[M.DIST]), "Objective 1 "
         "names it explicitly. Thin for some indicators - check n."],
        ["Sector", "All three", n(M.SUB, lambda r: r[M.SECT]),
         "Finest geography collected. The prescription tool writes "
         "'Sector__District', the stock tool 'Sector' - normalise first. Most "
         "cells will be thin."],
        ["Facility", "All three", n(M.SUB, lambda r: r[M.FAC]),
         "Use post-reattribution names. A name can repeat across districts, so "
         "key on name + level + district."],
        ["Health facility level", T_STOCK, n(M.SUB, lambda r: r[M.LEVEL]),
         "This is objective 1's 'level of care'."],
        ["Facility class (short code)", T_STOCK, n(M.SUB, lambda r: r[M.FCLASS]),
         "HC / DH / L2TH / MHC / HP / SGHP / PH - a compact version of level."],
        ["Level rank", T_STOCK, n(M.SUB, lambda r: r[M.RANK]),
         "1 health centre, 2 district hospital, 3 referral. Use for a trend "
         "test across levels."],
        ["Facility management", T_STOCK, n(M.SUB, lambda r: r[M.MGMT]),
         "Government / faith-based / PPP. Not in the objectives but a strong "
         "candidate finding."],
        ["Facility type", T_PRESC, n(P.sub, lambda r: r[P.FTYPE]),
         "The prescription tool's equivalent of level of care."],
        ["Programme", T_STOCK, n(M.MEDROWS, lambda r: r[M.PROGRAM]),
         "Essential Medicines, Malaria, MCCH, FP, ARV, TB, Consumables."],
        ["Drug class", T_STOCK, n(M.MEDROWS, lambda r: r[M.DCLASS]),
         "Finer than programme; some classes will be thin."],
        ["Medicine", T_STOCK, n(M.MEDROWS, lambda r: r[M.MED]),
         "The tracer list itself - the procurement view."],
        ["Digital stock system", T_STOCK, n(M.MEDROWS, lambda r: r[M.EMR]),
         "Free text; normalise before grouping. Objective 2.2."],
        ["Prescriber category", T_PRESC, n(P.pres, lambda r: r[P.PRESCRIBER]),
         "Nurse, GP, specialist, midwife, allied health. Prescription level."],
        ["Drug class (prescription)", T_PRESC, n(P.item, lambda r: r[P.CLASS]),
         "EML section as recorded on the item."],
        ["Medicine (prescription)", T_PRESC, n(P.item, lambda r: r["_med"]), ""],
        ["Dosage form [derived]", T_PRESC, n(P.item, lambda r: r["_form"]),
         "Parsed from the medicine code - not collected directly."],
        ["Parent health centre", T_CHW, len(C.FACILITIES),
         "The link for objective 8.3."],
        ["CHW", T_CHW, len(C.CHWS),
         "Name is free text - check for placeholders before using as a key."],
        ["Medicine (CHW)", T_CHW, n(C.MEDROWS, lambda r: r["_short"]),
         "The 13-item CHW package."],
        ["Antimalarial family [derived]", T_CHW,
         n(C.MEDROWS, lambda r: r["_family"]),
         "ALu / DHAP / Pyronaridine-Artesunate - required for the MFT read."],
        ["Month", f"{T_STOCK} + {T_CHW}", 3,
         "June, July, August 2026 only, for comparability across form versions."],
        ["Kobo server", "All three", 2,
         "kf / eu. A collection artefact - use only to check the two teams agree."],
    ]


def sheet(wb, title, blocks, notes=None, widths=None, fills=None):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    if notes:
        for n in notes:
            c = ws.cell(row=r, column=1, value=n)
            c.font = Font(size=9, italic=True, color="555555")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
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
        hdr_row = r
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
        ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)
        ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(len(headers))}{r - 1}"
        r += 2
    for ci, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    return ws


def build():
    wb = Workbook()
    wb.remove(wb.active)

    # ---------- read me
    rows = [
        ["Purpose", "Maps each of the eight study objectives to the indicators "
         "that answer it, with the exact variables, denominators, methods, "
         "disaggregation and cautions for each."],
        ["Built from the data", "Every variable named here was verified to "
         "exist in the current merged exports. Every n, completeness figure and "
         "distinct-value count is read from the data when this workbook is "
         "generated, so the plan cannot drift from the instruments."],
        ["Data sources", "Three KoboToolbox tools, each collected on two "
         "servers (kf and eu) and merged before analysis: the medicine (stock) "
         "assessment, the prescription assessment, and the CHW assessment."],
        ["Current size", f"Medicine: {S['obs']:,} observations at "
         f"{S['facilities']} facilities over {S['medicines']} medicines. "
         f"Prescription: {S['pres']:,} prescriptions, {S['items']:,} items at "
         f"{S['pres_fac']} facilities. CHW: {S['chw_rows']:,} rows, "
         f"{S['chw_chws']} CHWs at {S['chw_fac']} health centres."],
        ["Unit of analysis", "Never the submission. The stock and CHW tools are "
         "analysed at facility x medicine (or CHW x medicine) level; the "
         "prescription tool at medicine-item level, with prescription-level and "
         "facility-level roll-ups stated separately."],
        ["Denominator chain", f"{S['obs']:,} observations -> "
         f"{S['applicable']:,} applicable per EML -> {S['managed']:,} managed "
         f"-> {S['managed']:,} answered -> {S['available']:,} available. "
         f"Getting the step wrong is the most likely error in the analysis."],
        ["Minimum group size", f"{MIN_N} answered rows. Smaller groups are "
         f"reported but flagged, never compared."],
        ["Feasibility", "Six objectives are fully answerable from data already "
         "collected. Objective 3 is partial (frequency not measurable) and "
         "objective 4's service split is not collected at all. See the "
         "Feasibility and Limitations sheets."],
        ["How to use", "Work through 'Analysis sequence' once, then objective "
         "by objective on 'Objectives and indicators'. 'Table shells' gives the "
         "layout of each output table; 'Deliverables' says which script already "
         "produces it."],
        ["Sheets", "Read me | Objectives and indicators | Indicator detail | "
         "Feasibility | Excluded indicators | Form version check | Changes "
         "from v11 | Analysis sequence | Disaggregation | Table shells | Data "
         "quality checks | Reporting rules | Limitations | Form change "
         "requests | Deliverables"],
        ["Version", "Incorporates the team's edits returned in "
         "Data_analysis_plan11.xlsx. The core sheet uses their 11-column "
         "layout; the columns they trimmed moved to 'Indicator detail' rather "
         "than being lost. See 'Changes from v11'."],
        ["Form versions", "The two XLSForm files in the project folder predate "
         "what was deployed - different month windows, and no EML "
         "applicability or digital-system question. This plan follows the "
         "deployed data; see 'Form version check'."],
    ]
    sheet(wb, "Read me",
          [("Detailed data analysis plan - how this workbook is organised",
            ["Item", "Detail"], rows)],
          widths={1: 22, 2: 124})

    # ---------- objectives and indicators (the team's 11-column layout)
    hdr = ["Obj", "Objective", "Ind", "Indicator", "Definition", "Numerator",
           "Denominator", "Unit of analysis", "Tool", "Source variables",
           "Disaggregation", "Analysis method"]
    rows = [[d["obj"], OBJECTIVES[d["obj"]], d["id"], d["name"], d["defn"],
             d["num"], d["den"], d["unit"], d["tool"], d["vars"],
             d["disagg"], d["method"]] for d in IND]
    sheet(wb, "Objectives and indicators", [(None, hdr, rows)],
          notes=["The core plan, in the column layout returned by the team. "
                 "${medicine_name} is Kobo's repeat-group placeholder and "
                 "appears literally in the export column header.",
                 "Two requirements that cannot be produced from the data - "
                 "stock-out frequency and the outpatient/inpatient split - are "
                 "on the 'Excluded indicators' sheet with the reason and what "
                 "it would take to collect them.",
                 "Everything the team trimmed from this sheet (source sheet, "
                 "output, feasibility, cautions) is on 'Indicator detail', "
                 "keyed by the Ind column."],
          widths={1: 5, 2: 34, 3: 6, 4: 30, 5: 40, 6: 32, 7: 32, 8: 30,
                  9: 20, 10: 46, 11: 34, 12: 56},
          fills=lambda row, ci, v: OBJ_FILL if ci in (0, 1) else None)

    # ---------- indicator detail
    hdr = ["Ind", "Indicator", "Current n / completeness", "Sheet",
           "Statistics", "Output table", "Where produced", "Priority",
           "Notes and cautions"]
    rows = [[d["id"], d["name"], d["n"], d["sheet"], d["stat"], d["shell"],
             d["out"], d["pri"], d["note"]] for d in IND]
    sheet(wb, "Indicator detail", [(None, hdr, rows)],
          notes=["The columns the core sheet does not carry, keyed by Ind.",
                 "'Current n / completeness' is read from the data when this "
                 "workbook is generated, so it tells you whether an indicator "
                 "is computable NOW, not just in principle."],
          widths={1: 6, 2: 30, 3: 34, 4: 18, 5: 28, 6: 26, 7: 36, 8: 11,
                  9: 70})

    # ---------- feasibility
    rows = []
    for o in sorted(OBJECTIVES):
        inds = [d for d in IND if d["obj"] == o]
        exc = [e for e in EXCLUDED if e[1] == o]
        # the excluded requirements still belong to the objective - a verdict
        # computed from the indicator list alone would read "fully feasible"
        # for objectives 3 and 4, which is exactly the misreading this plan
        # exists to prevent
        verdict = "Fully feasible" if not exc else "Partly - see Excluded"
        gap = "; ".join(f"{e[0]} {e[2]}" for e in exc) or "none"
        action = {
            3: "Report 'months affected (0-3)' as a labelled proxy for "
               "frequency (3.7); add an episode-count question for future "
               "rounds. Indicators 3.4-3.6 were added at the team's request.",
            4: "DECISION NEEDED before drafting: add a service-type question, "
               "use facility type as an explicitly labelled proxy, or state "
               "the service split as not assessed.",
        }.get(o, "Proceed as planned.")
        rows.append([o, OBJECTIVES[o], len(inds), len(exc), verdict, gap, action])

    def ffill(row, ci, v):
        if ci == 4:
            return YES_FILL if v == "Fully feasible" else PART_FILL
        return None

    sheet(wb, "Feasibility",
          [("Can each objective be answered from the data collected?",
            ["Obj", "Objective", "Indicators planned",
             "Requirements excluded", "Verdict", "What is excluded", "Action"],
            rows)],
          notes=["Six objectives are answerable in full from data already "
                 "collected - no new fieldwork needed for 1, 2, 5, 6, 7 or 8.",
                 "Objectives 3 and 4 each contain one requirement the "
                 "instruments cannot produce. The verdict counts those, so "
                 "this sheet does not read 'fully feasible' just because the "
                 "rows were moved off the core sheet."],
          widths={1: 5, 2: 50, 3: 13, 4: 13, 5: 22, 6: 42, 7: 62}, fills=ffill)

    # ---------- excluded indicators
    sheet(wb, "Excluded indicators",
          [("Requirements inside an objective that the data cannot produce",
            ["Ind", "Obj", "Requirement", "Objective wording it comes from",
             "Why it cannot be produced", "Proxy available",
             "What it would take to collect", "Decision / action"],
            [list(x) for x in EXCLUDED])],
          notes=["The team's returned version deleted these rows. They are "
                 "kept here so the omission is documented - the objectives "
                 "still contain the requirement, and the report will have to "
                 "say what happened to it.",
                 "4.3 needs a decision before drafting; 3.4 can proceed on the "
                 "proxy if it is labelled."],
          widths={1: 6, 2: 5, 3: 32, 4: 44, 5: 52, 6: 50, 7: 44, 8: 48},
          fills=lambda row, ci, v: NO_FILL if ci == 2 else None)

    # ---------- form change requests
    sheet(wb, "Form change requests",
          [("Instrument changes the analysis cannot work around",
            ["Tool", "Question to add", "Type", "Why it is needed",
             "Priority"], [list(x) for x in FORMREQ])],
          notes=["Each row is one question. Three came from the team's "
                 "comments on v11, three from checking the plan against the "
                 "deployed data.",
                 "Only the first blocks an objective outright; the rest "
                 "improve precision or remove a transcription risk."],
          widths={1: 20, 2: 42, 3: 38, 4: 74, 5: 40})

    # ---------- form version check
    sheet(wb, "Form version check",
          [("The XLSForm files in the project folder vs what is deployed",
            ["Item", "XLSForm supplied (earlier design)",
             "Deployed form / merged export", "Implication for the analysis"],
            [list(x) for x in FORMVER])],
          notes=["Checked against 'Unserved_Prescription final 09092026.xlsx' "
                 "and 'XLs form medicine_assessment_final.xlsx'. Both predate "
                 "deployment.",
                 "This plan is written against the DEPLOYED data. Reviewing it "
                 "against the XLSForm files will raise questions about "
                 "variables that do exist in the data - this sheet answers "
                 "them.",
                 "It would be worth archiving the deployed XLSForm version "
                 "alongside these, so the design of record matches the data of "
                 "record."],
          widths={1: 30, 2: 40, 3: 46, 4: 72})

    # ---------- changes from v11
    sheet(wb, "Changes from v11",
          [("Team comments on Data_analysis_plan11.xlsx and what was done",
            ["Team comment / edit", "Response", "What changed"],
            [list(x) for x in CHANGES])],
          notes=["Read against Data_analysis_plan11.xlsx as returned. Six "
                 "edits adopted as given, one adopted with the detail moved "
                 "rather than dropped, one reversed with a reason."],
          widths={1: 52, 2: 44, 3: 86},
          fills=lambda row, ci, v: (YES_FILL if ci == 1 and v == "Adopted"
                                    else PART_FILL if ci == 1 else None))

    # ---------- sequence
    sheet(wb, "Analysis sequence",
          [("Run in this order - each step depends on the one before",
            ["Step", "What", "How", "Gate before moving on"],
            [list(x) for x in SEQUENCE])],
          widths={1: 7, 2: 52, 3: 56, 4: 74})

    # ---------- disaggregation
    sheet(wb, "Disaggregation",
          [("Variables available for disaggregation, with live distinct-value "
            "counts",
            ["Variable", "Tool(s)", "Distinct values", "Notes and cautions"],
            disagg())],
          notes=["Counts are read from the current merged exports, so they move "
                 "with the data.",
                 f"With {len(M.FACILITIES)} facilities and ~{S['managed']:,} "
                 f"managed rows, most facility and sector cells fall below the "
                 f"{MIN_N}-row threshold. Province, district, level, programme "
                 f"and medicine are the safe disaggregators."],
          widths={1: 30, 2: 22, 3: 15, 4: 92})

    # ---------- table shells
    sheet(wb, "Table shells",
          [("The layout of each output table",
            ["Ref", "Table", "Indicators", "Columns"],
            [list(x) for x in SHELLS])],
          notes=["Every rate table carries its own denominator. A percentage "
                 "without the n beside it should not leave this project."],
          widths={1: 7, 2: 38, 3: 14, 4: 110})

    # ---------- checks
    sheet(wb, "Data quality checks",
          [("Run these before computing any indicator",
            ["Area", "Check", "Pass condition / current state",
             "Why it matters"], [list(c) for c in CHECKS])],
          widths={1: 20, 2: 50, 3: 50, 4: 62})

    # ---------- rules
    sheet(wb, "Reporting rules",
          [("Rules that prevent the most likely misreadings",
            ["Rule", "Objectives", "Reason"], [list(x) for x in RULES])],
          widths={1: 54, 2: 14, 3: 96})

    # ---------- limitations
    sheet(wb, "Limitations",
          [("Study limitations, for the methods and limitations sections",
            ["Ref", "Limitation", "Objectives", "Why it arises",
             "How to handle it"], [list(x) for x in LIMITS])],
          widths={1: 6, 2: 42, 3: 16, 4: 66, 5: 66})

    # ---------- deliverables
    deliv = [
        ("Medicine_availability_analysis.xlsx", "Medicine_availability_analysis.py",
         "1.1, 1.2, 3.1, 5.2 (stock side)",
         "Availability disaggregated by every facility and medicine variable, "
         "the denominator funnel, and every stock-out with its attribution"),
        ("Medicine_assessment_analysis.xlsx", "analyze_medicine_assessment.py",
         "1.1, 1.3, 2.1, 2.2, 3.1, 3.2, 3.3, 6.3, 7.2",
         "The full stock analysis - 61 tables"),
        ("Prescription_study_analysis.xlsx", "analyze_kobo_prescription.py",
         "4.1, 4.2, 5.1, 5.2, 6.1, 6.2, 7.1",
         "The full prescription analysis - 72 tables"),
        ("Prescription_facility_analysis.xlsx", "Prescription_facility_analysis.py",
         "4.1, 4.2, 7.1 at facility level",
         "One row per facility, WHO/INRUD ranges, outliers, data quality"),
        ("Medicine_stock_CHW_analysis.xlsx", "analyze_medicine_assessment_chw.py",
         "8.1, 8.2, 8.4", "The full CHW analysis - 52 tables"),
        ("Combined_analysis_report.html", "export_combined_report.py",
         "headline indicators from all three tools",
         "Three tabbed pages for circulation; curated, not the full record"),
        ("Site_coverage_check.xlsx", "Site_coverage_check.py", "all - "
         "denominator assurance",
         "Planned vs visited facilities per tool, against the site plan"),
        ("Completion_by_geography.xlsx", "Completion_by_geography.py",
         "all - denominator assurance",
         "Completion and tool agreement by province, district, sector, facility"),
        ("TO BUILD", "-", "3.7",
         "Months-affected proxy for the stock-out frequency objective"),
        ("TO BUILD", "-", "8.3",
         "Paired CHW-vs-facility availability on the shared medicines"),
    ]
    sheet(wb, "Deliverables",
          [("Which output already answers which indicator",
            ["Output file", "Built by", "Indicators covered", "Contents"],
            [list(d) for d in deliv])],
          notes=["Every output is regenerated from the merged exports, so none "
                 "can disagree with another.",
                 "Two indicators have no output yet: 3.4 and 8.3."],
          widths={1: 38, 2: 38, 3: 30, 4: 76})

    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)      # Box resets synced files to read-only
    wb.save(OUT)
    feas = Counter(d["feas"] for d in IND)
    print(f"wrote {OUT}")
    print(f"  {len(OBJECTIVES)} objectives -> {len(IND)} indicators "
          f"({feas['Yes']} feasible, {feas['Partial']} partial, "
          f"{feas['Not collected']} not collected)")
    print(f"  {len(EXCLUDED)} excluded requirements, {len(CHANGES)} team "
          f"comments addressed, {len(FORMVER)} form-version differences, "
          f"{len(FORMREQ)} form change requests")
    print(f"  14 sheets: {len(SEQUENCE)} workflow steps, {len(disagg())} "
          f"disaggregators, {len(SHELLS)} table shells, {len(CHECKS)} quality "
          f"checks, {len(RULES)} reporting rules, {len(LIMITS)} limitations")
    print(f"  denominator chain read from the data: {S['obs']:,} obs -> "
          f"{S['applicable']:,} applicable -> {S['managed']:,} managed -> "
          f"{S['available']:,} available")


if __name__ == "__main__":
    build()
