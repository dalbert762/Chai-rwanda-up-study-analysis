"""Full national report: descriptive results (dashboard) + inferential and
predictive analysis, one content source, written to LaTeX (compiled to PDF)
and Word.

    python3 descriptive_figures.py
    python3 ../analytics/deck_figures.py figures
    python3 build_full_report.py   -> QMIA2026_full_report.tex / .pdf / .docx
"""
import contextlib
import io
import os
import re
import sys

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
INF = os.path.join(ROOT, "Inferential analysis")
sys.path.insert(0, INF)
sys.path.insert(0, HERE)
with contextlib.redirect_stdout(io.StringIO()):
    import build_report as BR          # inferential content and helpers
import reportlib as RL
import wbtables as W

import shutil
FIGDIRS = [os.path.join(HERE, "figures")]
for _f in os.listdir(os.path.join(INF, "results")):      # one local folder: LaTeX
    if _f.endswith(".png"):                                # cannot read paths with spaces
        shutil.copy2(os.path.join(INF, "results", _f), os.path.join(HERE, "figures", _f))
NAME = "QMIA2026_full_report"
TITLE = "Medicine Stock and Prescription Assessment 2026"
SUBTITLE = ("National report: medicine availability, stock management, prescription "
            "fulfilment and community-level supply, with inferential and predictive analysis")
META = ("Quality Management Improvement Approach (QMIA) 2026 · Joint assessment by "
        "MoH, RMS, RSSB and CHAI · Fieldwork 14-25 September 2026 · Report "
        "version of 5 November 2026 (draft for validation)")

RJ, E, T, orci, ci, f1, pv = BR.RJ, BR.E, BR.T, BR.orci, BR.ci, BR.f1, BR.pv
B = []
h1 = lambda t: B.append(("h1", t))
h2 = lambda t: B.append(("h2", t))
h3 = lambda t: B.append(("h3", t))
p = lambda t: B.append(("p", t))
bl = lambda *t: B.append(("bullets", list(t)))
fg = lambda fn, cap, w=16: B.append(("fig", fn, cap, w))


def tb(title, hdr, rows, w, note=""):
    B.append(("table", title, hdr, rows, w, note))


def wb(tid, title=None, max_rows=None, cols=None, w=None, note=None, rename=None):
    t, h, r, n = W.table(tid, max_rows=max_rows, cols=cols, rename=rename)
    tb(title or t, h, r, w or [1] * len(h), note or "")   # workbook notes are internal
    return r


def wbval(tid, label):
    for r in W.table(tid)[2]:
        if r[0].strip().upper().startswith(label.upper()):
            return r[1]
    raise KeyError(label)


# ---------- reuse the inferential sections
def br_section(prefix, table_titles, demote="h3", keep_heading=True, new_title=None):
    out, on, ti = [], False, 0
    for blk in BR.C:
        if blk[0] in ("h1", "h2"):
            if on:
                break
            if blk[0] == "h2" and blk[1].startswith(prefix):
                on = True
                if keep_heading:
                    out.append((demote, new_title or re.sub(r"^\d+\.\d+\s*", "", blk[1])))
            continue
        if not on:
            continue
        if blk[0] == "table":
            hdr, rows, w, cap = blk[1:]
            out.append(("table", table_titles[ti], hdr, rows, w, cap))
            ti += 1
        elif blk[0] == "fig":
            out.append(("fig", blk[1], re.sub(r"^Figure \d+\.\s*", "", blk[2]), blk[3]))
        else:
            out.append(blk)
    B.extend(out)


mg = pd.read_csv(os.path.join(INF, "data", "stock_managed_rows.csv"))
chw = pd.read_csv(os.path.join(INF, "data", "chw_rows.csv"))
m1, m2 = E["Availability, expected medicines (M1)"], E["Availability, managed products (M2)"]
LAG = BR.LAG
RD = BR.RD
WHO = BR.WHO
PA, PB = BR.PA, BR.PB

# headline descriptive numbers, read from the dashboard workbook
gap_rows = mg[(mg.available == 0) & mg.rms_central.isin(["Yes", "No"]) & mg.rms_regional.isin(["Yes", "No"])]
GAP_N, GAP_Y = len(gap_rows), int(((gap_rows.rms_central == "Yes") | (gap_rows.rms_regional == "Yes")).sum())
LOST = wbval("8.7", "LOST SALES")
_ad = mg[mg.adequacy.notna() & (mg.adequacy != "")]
ADQ = {k: 100 * (_ad.adequacy == k).mean() for k in ("out", "under", "adequate", "over")}
MAL_OVER = 100 * (_ad[_ad.programme == "Malaria"].adequacy == "over").mean()
L2_ADQ = 100 * (_ad[_ad.level == "Level 2 teaching hospital"].adequacy == "adequate").mean()
UNITS_LOST = wbval("8.7", "Estimated units lost")

# ================================================================== front matter
B.append(("unnumbered_h1", "Executive summary"))
p("The 2026 Quality Management Improvement Approach (QMIA) assessment was conducted "
  "jointly by the Ministry of Health (MoH), Rwanda Medical Supply (RMS), the Rwanda Social "
  "Security Board (RSSB) and the Clinton Health Access Initiative (CHAI). Between 14 and "
  "25 September 2026, 27 teams visited **266 health facilities in all 30 districts**, "
  "assessed **55 tracer products** (80 strengths; 14,378 facility x product records), "
  "reviewed **6,006 prescriptions** (14,059 prescribed items) and visited **175 community "
  "health workers (CHWs)**. This report presents the descriptive findings shown on the "
  "indicator dashboard together with inferential statistics (design-based confidence "
  "intervals and regression models) and predictive models.")
tb("Headline indicators (95% confidence intervals account for clustering by facility)",
   ["Objective", "Indicator", "Result"],
   [["1. Availability", "Managed products available on the visit day", ci(m2["unweighted"])],
    ["1. Availability", "Expected (EML-authorised) medicines available", ci(m1["unweighted"])],
    ["2. Records", "Physical count equals the digital record (exact match)",
     ci(E["Inventory record accuracy (exact match)"]["unweighted"])],
    ["2. Stock adequacy", "Adequately stocked (1-2 months; 1-3 at teaching and referral hospitals)", ci(E["Adequately stocked (1-2 months)"]["unweighted"])],
    ["2. Stock adequacy", "Overstocked / understocked / stocked out",
     f"{f1(E['Overstocked (>2 months)']['unweighted']['est'])}% / {f1(E['Understocked (<1 month)']['unweighted']['est'])}% / {f1(E['Stocked out on the visit day']['unweighted']['est'])}%"],
    ["3. Stock-outs", "Managed pairs with a stock-out day, June-August", ci(E["Stock-out incidence, Jun-Aug"]["unweighted"])],
    ["4. Prescriptions", "Prescriptions fully served", ci(E["Prescriptions fully served"]["unweighted"])],
    ["5. Attribution", "Missing medicines in stock at RMS Central or Regional",
     f"{100 * GAP_Y / GAP_N:.1f}% ({GAP_Y} of {GAP_N})"],
    ["6. Orders", "Quantity fill rate (delivered / requested)", ci(E["Order fill rate (quantity)"]["unweighted"])],
    ["7. Value", "Value of unserved prescription items (observed)", "834,418 RWF"],
    ["7. Value", "Estimated sales lost to stock-outs, June-August", f"{LOST} RWF"],
    ["8. Community", "CHW product availability (managed)", ci(E["CHW product availability (managed)"]["unweighted"])],
    ["8. Community", "CHWs able to treat malaria, diarrhoea and pneumonia", ci(RD["All three conditions"])]],
   [2.6, 7.6, 4.4])
B.append(("h2u", "Key findings"))
bl(f"**Availability is high but uneven.** {f1(m2['unweighted']['est'])}% of managed products "
   f"and {f1(m1['unweighted']['est'])}% of EML-authorised medicines were on the shelf. "
   f"Weighted to the national facility mix, expected availability is "
   f"{f1(m1['weighted']['est'])}%, because health posts (about 69% of facilities) hold only 71% "
   f"of the medicines authorised for them. Insulin (27.6%), gliclazide (44.7%) and "
   f"beclomethasone (54.1%) are the least available medicines.",
   f"**Most gaps are distribution gaps.** {100 * GAP_Y / GAP_N:.1f}% of medicines missing at the "
   f"facility were reported in stock at RMS Central or Regional. ARV gaps are the exception: "
   f"about half were national stock-outs (notably HIV Determine test kits).",
   f"**Records do not match the shelf in almost half of cases.** Exact-match accuracy is "
   f"{f1(E['Inventory record accuracy (exact match)']['unweighted']['est'])}%, down from the 85% "
   f"QMIA reported for 2019. Accuracy is lowest at health posts and in facilities recording in "
   f"eBuzima (adjusted OR {orci(T('gee_accuracy', 'eBuzima'))} against eLMIS).",
   f"**Stock is poorly balanced.** Only {ADQ['adequate']:.1f}% of items sat in the target band "
   f"(1-2 months; 1-3 at teaching and referral hospitals); {ADQ['over']:.1f}% were overstocked "
   f"({MAL_OVER:.0f}% of malaria items) and {ADQ['under']:.1f}% understocked, of which 928 items "
   "were at or below the 0.5-month emergency order point.",
   f"**One managed product in ten ran out in June-August** "
   f"({f1(E['Stock-out incidence, Jun-Aug']['unweighted']['est'])}%), and stock-outs were long: "
   f"TB products were out 113 days on average. A requested order that received nothing raised "
   f"the odds of a stock-out the following month {T('gee_lag', 'Requested, nothing delivered')['OR']:.1f}-fold "
   f"({f1(LAG['Requested, nothing delivered']['mean'])}% against {f1(LAG['Delivered in full']['mean'])}% "
   f"after a full delivery); not ordering at all raised them "
   f"{T('gee_lag', 'No request')['OR']:.1f}-fold.",
   "**Prescriptions are almost always fully served (98.8%)**, but this understates unmet need "
   "because prescribers write what the pharmacy holds. One in three prescriptions at national "
   "referral hospitals (32.9%) had a missing item. Antibiotics appear in "
   f"{f1(WHO['% encounters with an antibiotic']['est'])}% of encounters, above the WHO reference of 30%.",
   "**Orders are mostly filled (93.1% of quantity)**, but ARV (66.2%) and family planning "
   "(76.6%) lag, and national referral hospitals received 77.6% of what they ordered.",
   f"**The financial cost is concentrated.** Observed unserved items were worth 834,418 RWF, "
   f"73% of it at national referral hospitals; stock-outs cost an estimated {LOST} RWF in lost "
   f"sales over three months.",
   f"**At community level, availability is {f1(E['CHW product availability (managed)']['unweighted']['est'])}% "
   f"but only {f1(RD['All three conditions']['est'])}% of CHWs could treat all three iCCM "
   f"conditions on the day.** Where a CHW was out of an antimalarial or RDT, the supervising "
   f"health centre had it in every case: the gap is resupply from the health centre.",
   f"**Stock-outs can be anticipated.** A model using June-July data predicts August "
   f"stock-outs with an AUC of {BR.best(PA)['auc']:.2f}; most of its power comes from the "
   f"persistence of stock-outs, so a monthly list of products out last month plus the "
   f"model's top-risk products is a practical early-warning tool.")
B.append(("h2u", "Priority recommendations"))
bl("**RMS:** treat a requisition delivered at 0% as an alert and expedite it within the month; "
   "report nil-filled orders by hub. Escalate national stock-outs of HIV test kits, insulin, "
   "thiamine and paediatric RHZ with programmes and central procurement.",
   "**Facilities and districts:** never skip a requisition cycle; use the min-max rule and the "
   "emergency order point; redistribute excess programme stock.",
   "**QMIA mentorship:** prioritise monthly physical-count reconciliation at health posts, health "
   "centres using eBuzima, and PPP facilities; target the facility profiles in Section 7.",
   "**Health centres:** resupply CHWs at the monthly CHW meeting, with ORS and zinc as a pair.",
   "**MoH and RSSB:** align the reimbursement list with the EML; finalise the official MFT zoning "
   "list; strengthen antimicrobial stewardship in primary care.")
B.append(("unnumbered_h1", "Abbreviations"))
tb("Abbreviations used in this report", ["Abbreviation", "Meaning"],
   [["ACT", "Artemisinin-based combination therapy"], ["AL / ALu", "Artemether-lumefantrine"],
    ["AMC", "Average monthly consumption"], ["ARV", "Antiretroviral (programme)"],
    ["AUC", "Area under the ROC curve"], ["CHUB / CHUK", "University Teaching Hospitals of Butare / Kigali"],
    ["CHW", "Community health worker"], ["CI", "Confidence interval"],
    ["DH", "District hospital"], ["DHAP", "Dihydroartemisinin-piperaquine"],
    ["EML", "National Essential Medicines List"], ["eLMIS", "Electronic logistics management information system"],
    ["FP", "Family planning"], ["GEE", "Generalised estimating equations"],
    ["HC / HP", "Health centre / health post"], ["iCCM", "Integrated community case management"],
    ["L2TH", "Level 2 teaching hospital"], ["MCCH", "Maternal, child and community health"],
    ["MFT", "Multiple first-line therapies (malaria)"], ["MoH", "Ministry of Health"],
    ["NRH", "National referral hospital"], ["OR", "Odds ratio"],
    ["PPP", "Public-private partnership facility"], ["QMIA", "Quality Management Improvement Approach"],
    ["RDT", "Rapid diagnostic test"], ["RMS", "Rwanda Medical Supply"],
    ["RSSB", "Rwanda Social Security Board"], ["Pyr/Art, ASPY", "Pyronaridine-artesunate"]],
   [3.5, 12.5])

# ================================================================== 1
h1("Introduction")
h2("Background")
p("Rwanda's pharmaceutical supply chain runs from central level through RMS branches to "
  "health facilities and on to community health workers. The MoH and its partners use the "
  "QMIA to monitor supply chain performance, identify gaps and their root causes, and follow "
  "corrective action. In 2026, RMS and RSSB were also planning separate prescription "
  "assessments; to avoid duplication, these were merged with QMIA into a single coordinated "
  "assessment.")
p("Availability is usually measured as the share of tracer medicines on the shelf on the day "
  "of a visit. That indicator does not show whether medicines prescribed to patients are "
  "actually dispensed. An unfilled prescription is a direct patient-access event: the patient "
  "may buy the medicine privately, return later, go elsewhere or go without treatment. Because "
  "routine reporting does not link prescriptions to stock and supply data, the size and value "
  "of unfilled prescriptions are not measured nationally. This assessment links them.")
h2("Objectives")
p("The assessment aims to quantify medicine availability and the volume, causes and cost of "
  "unserved prescriptions across all levels of care, and to attribute each gap to the point in "
  "the supply chain where it arises. Specifically, it will:")
bl("1. Measure availability of a defined tracer medicine list on the day of visit, by level of "
   "care, province and district.",
   "2. Determine stock accuracy by comparing physical stock with stock recorded in the digital "
   "tool used at each facility.",
   "3. Quantify stock-outs, their current duration and frequency over the preceding three months.",
   "4. Measure the proportion of prescriptions that go unserved and the quantity of medicine involved.",
   "5. Determine whether each unserved medicine was available at RMS regional and central "
   "warehouses, separating distribution challenges from nationwide stock-outs.",
   "6. Establish whether the facility requested the missing medicine and whether it was supplied "
   "fully, partially or not at all.",
   "7. Estimate the financial value of unserved prescriptions at official tariff prices.",
   "8. Describe availability at community level through the CHW attached to each selected facility.")
p("Beyond these descriptive objectives, this report tests which differences are statistically "
  "robust, identifies factors associated with availability, stock-outs and record accuracy, and "
  "assesses whether stock-outs can be predicted from routine data.")

# ================================================================== 2
h1("Methodology")
h2("Design and sampling")
p("The assessment used a multistage stratified cluster design. The sample size was derived with "
  "the Yamane formula (95% confidence, 5% margin of error) and allocated proportionally by level "
  "of care, province and district, with facilities drawn by computer-generated random sampling. "
  "One health post per district was added to ensure representation at the most peripheral tier. "
  "At each facility, 20 prescriptions were reviewed at health centres and 30 at hospitals: 10 "
  "from the visit day and the rest from the preceding 10 days. The CHW attached to each selected "
  "health centre was visited.")
h2("Tools, team and fieldwork")
p("Three KoboToolbox forms were used: a medicine stock assessment (55 tracer products covering "
  "essential medicines and programme commodities, selected from the country's top diagnoses), a "
  "prescription assessment, and a CHW stock assessment. Tools were piloted at three Kigali "
  "facilities on 21 August 2026 and revised. 27 two-person teams pairing central-level staff "
  "(MoH, RMS, RBC, RSSB, CHAI) with RMS branch managers collected data over 10 working days. "
  "Data sources were facility physical counts, stock cards and digital records, dispensing logs, "
  "prescription registers and requisition records, with RMS and RSSB data where available.")
h2("Achieved sample")
fac = mg.drop_duplicates("facility")
ct = pd.crosstab(fac.level, fac.province)
LV = ["Health post", "Health centre", "District hospital", "Level 2 teaching hospital", "National referral hospital"]
ct = ct.reindex(LV)
tb("Facilities assessed by level of care and province",
   ["Level of care"] + list(ct.columns) + ["Total"],
   [[l] + [str(int(v)) for v in ct.loc[l]] + [str(int(ct.loc[l].sum()))] for l in LV]
   + [["Total"] + [str(int(ct[c].sum())) for c in ct.columns] + [str(int(ct.values.sum()))]],
   [4.6, 1.8, 2.0, 1.8, 1.8, 1.8, 1.6])
mt = fac.management.value_counts()
tb("Facilities by management type", ["Management", "Facilities", "Share"],
   [[k, str(v), f"{100 * v / mt.sum():.1f}%"] for k, v in mt.items()], [8, 3, 3])
wb("5.8", "Prescriptions and items reviewed by level of care", cols=[0, 1, 2, 5],
   w=[6, 3, 3, 3], note="")
cp = chw.drop_duplicates("chw").province.value_counts()
tb("CHWs assessed by province", ["Province", "CHWs"], [[k, str(v)] for k, v in cp.items()]
   + [["Total", str(cp.sum())]], [8, 3])
h2("Data management and cleaning")
p("Data from two Kobo servers were merged, linked across forms and cleaned with a documented, "
  "reproducible pipeline. Each rule below is applied identically in the dashboard, the cleaned "
  "datasets and this report; the cleaning log lists every affected record.")
tb("Cleaning rules applied", ["Rule", "Description", "Records"],
   [["S1", "Medicine rows linked to their parent submission", "14,378"],
    ["S2", "Duplicate facility x medicine rows removed", "12"],
    ["S3", "Non-finite calculated fields treated as missing", "-"],
    ["S4-S5", "Level of care grouped; two servers merged", "454 visits"],
    ["S6", "EML applicability corrected against the official EML", "608"],
    ["S7", "'Available' with zero quantity on the shelf recoded to not available", "173"],
    ["S8", "Digital-system names normalised (e.g. 21 spellings of OpenMRS)", "6,866"],
    ["S9", "Unavailable strength recoded to not managed where another strength was available; other answers blanked", "199"],
    ["S10", "Out-of-zone antimalarial (managed, not available) recoded to not managed; other answers blanked", "16 stock, 34 CHW"],
    ["S11", "FP products at two faith-based facilities (Mubuga, Birambo) recoded to not managed", "8"],
    ["P1", "Prescription facility names reattributed", "108"],
    ["P2", "Prescription drug classes corrected (e.g. respiratory medicines)", "69"]],
   [1.6, 11.4, 3.0])
h2("Indicator definitions")
tb("Main indicators", ["Indicator", "Definition"],
   [["Expected availability (M1)", "Available on the visit day / medicines expected: Essential Medicines authorised by the EML at the level (not managed counted unavailable); programme products managed by the facility. Strengths combined per molecule."],
    ["Managed availability (M2)", "Available on the visit day / products the facility manages (answered)."],
    ["Record accuracy", "Physical count equals the quantity in the digital system, among records naming a system."],
    ["Stock adequacy", "Months of stock = stock on hand / AMC: under 1 month understock; 1-2 months adequate (1-3 at level 2 teaching and national referral hospitals, over 3 overstock); over 2 overstock elsewhere; plus stocked out."],
    ["Stock-out incidence", "Managed pairs with at least one stock-out day in June, July or August / managed pairs."],
    ["Average stock-out days", "Mean of the form's total stock-out days over pairs with at least one day out."],
    ["Prescription fully served", "Every prescribed item recorded available at the facility."],
    ["Distribution gap", "Unavailable managed items reported in stock at RMS Central or Regional / unavailable items with both answers."],
    ["Quantity fill rate", "Quantity delivered / quantity requested, June-August (months where delivered did not exceed requested)."],
    ["CHW availability", "Available / products the CHW manages (not managed is mostly MFT zoning policy)."]],
   [4.2, 11.8])
h2("Statistical analysis")
p("Descriptive indicators are pooled proportions as shown on the dashboard. Inferential and "
  "predictive methods are summarised below and detailed with each result.")
br_section("3.2", [], demote="h3")
br_section("3.3", [], demote="h3")
br_section("3.4", [], demote="h3")
br_section("3.5", [], demote="h3")
br_section("3.6", [], demote="h3")

# ================================================================== 3 availability
h1("Medicine availability (objective 1)")
h2("Descriptive findings")
p(f"Of {m1['unweighted']['n_rows']:,} expected facility x molecule observations, "
  f"{f1(m1['unweighted']['est'])}% were available on the visit day; on managed products only, "
  f"availability was {f1(m2['unweighted']['est'])}%. The 5.9-point gap is medicines facilities "
  f"are authorised for but do not stock. Health posts have the largest gap: 92.6% of what they "
  f"manage is available, but only 71.0% of what the EML authorises.")
fg("gauges.png", "Availability of managed products by level of care.", 14)
fg("d1_availability.png", "Availability by programme and by level of care.")
wb("1.6", "Availability by level of care (expected medicines)", cols=[0, 1, 2, 3, 6, 7],
   w=[4.5, 2, 2.4, 3, 2, 2.2])
wb("1.5", "Availability by programme (expected medicines)", cols=[0, 1, 2, 3, 6, 7],
   w=[4.5, 2, 2.4, 3, 2, 2.2])
wb("1.7", "Availability by province (expected medicines)", cols=[0, 1, 2, 6, 7], w=[4.5, 2, 2.6, 2.4, 2.4])
fg("map_tracer.png", "Managed-product availability by district.", 11)
fg("map_prov.png", "Managed-product availability by province.", 9)
p("By district, Gicumbi, Ngoma, Gakenke, Rubavu and Bugesera are below 90% on managed products; "
  "Burera, Musanze, Rwamagana and Nyamasheke are at or above 97%.")
wb("1.2", "The ten least available medicines (expected basis)", cols=[0, 1, 2, 4, 5],
   w=[7, 3, 2, 2, 2.2])
wb("1.3", "The twenty facilities with the lowest expected availability",
   cols=[0, 1, 2, 3, 4, 6], w=[4.6, 2.6, 3, 2, 2, 2.4])
wb("1.4", "Complete basket availability", w=[10, 4])
p("Only 13 of 266 facilities (4.9%) had every assessed medicine available; 110 were missing "
  "five or more.")
wb("1.10", "Why medicines are not managed (coded from free text)", cols=[0, 1, 3, 5, 6],
   w=[4.2, 2.2, 2.2, 1.6, 6.4])
p("The main reasons are no demand or referral, MFT zoning (policy), rejection by RSSB or cost, "
  "and use of an alternative product. Faith-based facilities do not provide modern family "
  "planning by policy (122 pairs).")
h2("Inferential findings")
br_section("4.4", ["Model 1: adjusted odds ratios for availability of a managed medicine"],
           new_title="Factors associated with availability")

# ================================================================== 4 records & adequacy
h1("Inventory record accuracy and stock adequacy (objective 2)")
h2("Descriptive findings: record accuracy")
p(f"Among {E['Inventory record accuracy (exact match)']['unweighted']['n_rows']:,} records naming "
  f"a digital system, the physical count equalled the system count in "
  f"{f1(E['Inventory record accuracy (exact match)']['unweighted']['est'])}%. Accuracy rises "
  f"with level of care, from 38.1% at health posts to 85.1% at district hospitals.")
fg("d3_accuracy.png", "Inventory record accuracy by level of care and by digital system.")
wb("2.5", "Inventory record accuracy by level of care", cols=[0, 1, 2, 3], w=[5, 3, 3, 3])
wb("2.6", "Inventory record accuracy by province", cols=[0, 1, 2, 3], w=[5, 3, 3, 3])
wb("2.7", "Inventory record accuracy by management type", cols=[0, 1, 2, 3], w=[5, 3, 3, 3])
h2("Descriptive findings: stock adequacy")
wb("2.1", "Stock status of managed products (months of stock)", w=[3.2, 6.5, 2.6, 2.4])
fg("d2_adequacy.png", "Stock adequacy by level of care and by programme.")
p(f"Overstock dominates: {ADQ['over']:.1f}% of items held more than the maximum (two months; "
  f"three at level 2 teaching and national referral hospitals), rising to {MAL_OVER:.0f}% for "
  f"malaria products. With the wider band, {L2_ADQ:.1f}% of items at level 2 teaching hospitals "
  f"are adequately stocked. Understock affects {ADQ['under']:.1f}%, and 928 items (49% of the understocked) were at or "
  "below 0.5 months, the emergency order point.")
h2("Inferential findings")
br_section("4.7", ["Model 3: adjusted odds ratios for an exact record match"],
           new_title="Factors associated with record accuracy")

# ================================================================== 5 consumption
h1("Medicine consumption")
p("Consumption (June-August) underpins quantification and the stock adequacy and lost-sales "
  "estimates. Average monthly consumption (AMC) is the mean of the months with a value.")
fg("d7_consumption.png", "Medicines with the highest total AMC across facilities.", 13)
wb("3.1", "Average monthly consumption: the twelve highest", max_rows=12,
   cols=[0, 1, 2, 3, 4, 5], w=[6, 3, 1.8, 2.4, 2.2, 2.2])
wb("3.4", "Consumption trend, August against June", cols=[0, 1, 4, 5, 6], w=[5, 3, 2.4, 2.4, 2.4])
wb("3.5", "Non-moving medicines (zero consumption in all three months)", w=[4.5, 3, 3, 2.6, 2.8])
wb("3.6", "Consumption value concentration (ABC analysis): the ten largest", max_rows=10,
   w=[1.2, 7.6, 3.4, 2, 2.2])
p("Consumption is heavily concentrated and skewed. Ten medicines account for most of the "
  "consumption value, led by examination gloves and sodium chloride infusion; means sit well "
  "above medians, so a few large facilities drive totals. 14.5% of managed pairs recorded no "
  "consumption in three months, most of them still holding stock, which is dead stock and an "
  "expiry risk.")

# ================================================================== 6 stock-outs
h1("Stock-outs: frequency and duration (objective 3)")
h2("Descriptive findings")
wb("4.1", "Three-month stock-out incidence", w=[10, 4])
wb("4.2", "Months affected by a stock-out (frequency proxy)", w=[5, 4, 4])
fg("d4_stockouts.png", "Stock-out incidence and mean duration by programme.")
wb("4.4", "Stock-out incidence by level of care and management type", cols=[0, 1, 2, 3, 4],
   w=[5, 2.4, 2.8, 2.6, 2.4])
fg("map_so.png", "Stock-out incidence by district.", 11)
wb("4.7", "Medicines with the heaviest stock-out burden (30 or more managed pairs)",
   max_rows=14, cols=[0, 1, 2, 3, 4, 8], w=[6.4, 3, 2, 2, 2, 2.2])
p("Stock-outs are uncommon for programme products but long once they occur (65-113 days on "
  "average). Thiamine (both strengths), insulin, HIV test kits, beclomethasone and paediatric RHZ "
  "carry the heaviest burden. 219 of 266 facilities had at least one stock-out in the period.")
h2("Inferential findings")
br_section("4.5", ["Model 2: adjusted odds ratios for any stock-out day, June-August"],
           new_title="Factors associated with stock-outs")
br_section("4.6", ["Stock-out next month by this month's order outcome",
                   "Model 4: adjusted odds ratios for a stock-out next month"],
           new_title="Does order fulfilment drive stock-outs?")

# ================================================================== 7 prescriptions
h1("Prescription fulfilment (objective 4)")
h2("Descriptive findings")
p("5,934 of 6,006 prescriptions (98.8%) were fully served; 58 were partly served (1.0%) and 14 had "
  "nothing available (0.2%). 107 of 14,059 prescribed items (0.8%) were unserved. National "
  "referral hospitals stand out: 23 of 70 prescriptions (32.9%) and 15.6% of items were unserved.")
wb("5.8", "Prescription fulfilment by level of care", w=[4.2, 1.8, 2.2, 2.4, 1.6, 1.8, 2, 1.8])
wb("5.12", "Unserved items by drug class: the most affected", max_rows=8, cols=[0, 2, 3, 4],
   w=[8, 2.6, 2.6, 2.4])
wb("5.9", "Most frequently unserved medicines", max_rows=10, cols=[0, 1, 2, 3, 4],
   w=[8, 2, 2.4, 2, 2])
p("Unserved prescribing clusters in insulins and diabetes medicines and in respiratory medicines. "
  "Because prescribers tend to write what the pharmacy holds, these rates understate true unmet "
  "need; the stock-based indicators are the better guide to availability.")
h2("Inferential findings")
br_section("4.8", ["WHO/INRUD prescribing indicators", "Prescribing indicators by level of care",
                   "Models 5 and 6: unserved prescriptions and antibiotic prescribing"],
           new_title="Prescribing quality and unserved prescriptions")

# ================================================================== 8 attribution
h1("Supply chain attribution (objective 5)")
p(f"For medicines unavailable at the facility, teams recorded whether the product was in stock "
  f"at RMS Regional and RMS Central. Of {GAP_N} unavailable items with both answers, {GAP_Y} "
  f"({100 * GAP_Y / GAP_N:.1f}%) were in stock at one or both RMS levels: a distribution gap "
  f"between RMS and the facility rather than a national shortage. Missing items were more often in "
  f"stock at RMS Central than at RMS Regional, by 4 to 16 points depending on level.")
lev = []
for l in LV:
    g = gap_rows[gap_rows.level == l]
    if len(g):
        y = ((g.rms_central == "Yes") | (g.rms_regional == "Yes"))
        lev.append([l, str(len(g)), f"{100 * (g.rms_central == 'Yes').mean():.1f}%",
                    f"{100 * (g.rms_regional == 'Yes').mean():.1f}%", f"{100 * y.mean():.1f}%"])
tb("Unavailable items in stock upstream, by level of care",
   ["Level of care", "Items", "At RMS Central", "At RMS Regional", "Distribution gap"], lev,
   [5, 2, 3, 3, 3])
prg = []
for k, g in gap_rows.groupby("programme"):
    if len(g) >= 10:
        y = ((g.rms_central == "Yes") | (g.rms_regional == "Yes"))
        prg.append([k, str(len(g)), f"{100 * y.mean():.1f}%", f"{100 * (1 - y.mean()):.1f}%"])
tb("Distribution gap against nationwide stock-out, by programme (10 or more items)",
   ["Programme", "Items", "Distribution gap", "Nationwide stock-out"], prg, [5, 2.4, 3.4, 3.4])
wb("1.9", "Medicines with a nationwide stock-out (not at the facility and not at RMS Central)",
   cols=[0, 1, 3, 4, 5, 6], w=[6.4, 2.4, 1.8, 1.6, 2, 2])
p("ARV products are the exception to the distribution pattern: about half of ARV gaps were also "
  "out at RMS, a national supply problem driven by HIV Determine test kits (32 facilities in 13 "
  "districts). For prescriptions, RMS status was often unknown: 41 of 107 unserved lines had "
  "\"don't know\" for both stores, a gap in the attribution evidence.")

# ================================================================== 9 orders
h1("Requisition and order fulfilment (objective 6)")
p("Across June-August, facilities requested 38.1 million units and received 35.4 million: a "
  "quantity fill rate of 93.1%, steady across months (92.6% in June, 92.0% in July, 94.7% in "
  "August).")
fg("d5_fill.png", "Quantity fill rate by programme and by month.")
wb("7.7", "Quantity fill rate by level of care", w=[5, 2, 3, 3, 2.4])
wb("7.10", "Quantity fill rate by programme", max_rows=8, cols=[0, 1, 3, 4, 5], w=[4.6, 2, 3, 3, 2.4])
wb("7.8", "The most under-supplied medicines", max_rows=10, w=[6.6, 2.4, 2.4, 2.4, 1.8, 2])
p("ARV (66.2%), family planning (76.6%) and MCCH (85.3%) are filled least, and national referral "
  "hospitals received 77.6% of their requests, falling to 20% in July and 35% in August. Among "
  "unserved prescription items, only 31.8% had been ordered that month, and three quarters of "
  "those orders were not supplied at all.")
wb("7.1", "Requisition rate for unserved prescription items", w=[8, 2.4, 2.4, 2])
wb("7.2", "Outcome of the order where one was placed (unserved items)", w=[6, 3, 3])

# ================================================================== 10 financial
h1("Financial value (objective 7)")
p("Unserved prescription items were valued at official tariff prices: 834,418 RWF in total, "
  f"heavily concentrated at national referral hospitals and in a few high-cost combination "
  f"products. Separately, sales lost to stock-outs were estimated for each managed pair from its "
  f"in-stock consumption and the days it was out: {UNITS_LOST} units worth {LOST} RWF over "
  f"June-August including the facility's 20% margin.")
wb("8.5", "Unserved value by level of care and province", cols=[0, 2, 3, 4], w=[5, 3, 3.4, 2.6])
wb("8.4", "Medicine contribution to the unserved value (Pareto)", max_rows=8, w=[8, 2, 2.4, 2, 2.2])
wb("8.7", "Estimated sales lost to stock-outs", w=[11, 4])
wb("8.8", "Estimated lost sales by level of care", cols=[0, 1, 2, 3, 6, 7],
   w=[4.4, 1.8, 1.8, 2.6, 3.4, 2.2])
p("Lost sales are concentrated: one facility x medicine (a triple antihypertensive at CHUK, out "
  "74 days) accounts for 45.9% of the estimate, so the total is sensitive to a few large pairs. "
  "Programme products supplied free and insulin and carbetocin without a tariff price are excluded.")

# ================================================================== 11 community
h1("Community-level availability (objective 8)")
h2("Descriptive findings")
wb("9.1", "Overall CHW medicine availability", w=[10, 4])
fg("d6_chw.png", "CHW availability by product and by province.")
wb("9.2", "Availability by CHW product", w=[5, 2.4, 2.4, 2.4, 2])
wb("9.3", "Complete CHW basket availability", w=[10, 4])
fg("map_chw.png", "CHW availability by district.", 11)
wb("9.4", "Community against facility on the products both carry", w=[3.4, 2.4, 2.4, 2.6, 2.6, 2.6])
p("The denominator excludes products the CHW does not manage (842 rows, about 42%), mostly MFT "
  "zoning, which is policy rather than a stock-out. ORS, DHAP, zinc and amoxicillin are the least "
  "available products; North (65.9%) lags while the West reaches 92.4%.")
h2("Inferential findings")
br_section("4.9", ["CHW treatment readiness by iCCM condition", "CHW readiness by province",
                   "CHW against supervising health centre on the same product"],
           new_title="Treatment readiness and the health centre-to-CHW link")

# ================================================================== 12 cross-cutting
h1("Cross-cutting analyses")
br_section("4.1", ["Headline indicators: pooled and weighted estimates with 95% confidence intervals"],
           demote="h2", new_title="Precision and representativeness of the estimates")
br_section("4.2", ["Managed availability and stock-out incidence by level of care (95% CI)",
                   "Managed availability and stock-out incidence by management type (95% CI)",
                   "Managed availability and stock-out incidence by province (95% CI)",
                   "Managed availability and stock-out incidence by RMS hub (95% CI)"],
           demote="h2", new_title="Which group differences are statistically robust")
br_section("4.3", ["Variance partition: share of variation between facilities and districts"],
           demote="h2", new_title="Where the variation lies")
br_section("4.10", ["Spatial autocorrelation of district indicators (Moran's I)"], demote="h2",
           new_title="Geographic clustering")
br_section("4.11", ["Facility profiles from k-means clustering"], demote="h2",
           new_title="Facility profiles for targeting")
br_section("4.12", ["Predictive model performance (facility-grouped cross-validation)",
                    "Early-warning operating points: model against a simple rule"], demote="h2",
           new_title="Predicting stock-outs")

# ================================================================== 13 discussion
h1("Discussion")
h2("Main messages")
br_section("5.1", [], keep_heading=False)
h2("Comparison with other studies")
for pre in ("2.1", "2.2", "2.3", "2.4", "2.5", "2.6", "2.7", "2.8"):
    br_section(pre, [], demote="h3")
h2("Strengths and limitations")
br_section("5.2", [], keep_heading=False)
bl("**Visit-day measurement.** Availability is a single-day snapshot; stock-out history relies "
   "on stock cards and records for June-August only (March-May fields were empty).",
   "**Partial RSSB data.** RSSB billing fields cover 215 facilities, mostly health centres.",
   "**CHW sample.** 175 CHWs against 237 planned; Kirehe has no CHW data and no cadre was recorded.")

# ================================================================== 14 recommendations
h1("Recommendations")
br_section("6.", [], keep_heading=False)
for blk in BR.C:
    if blk[0] == "h1" and blk[1].startswith("6."):
        break
idx = [i for i, b_ in enumerate(BR.C) if b_[0] == "h1" and b_[1].startswith("6.")][0]
B.append(BR.C[idx + 1])
bl("**RSSB and MoH: align the reimbursement list with the EML** (beclomethasone, diclofenac "
   "suppository and gliclazide are the medicines most often not stocked because they are not "
   "reimbursed or cost too much).",
   "**MoH: complete eBuzima** with an offline module and a link to RSSB billing, and end parallel "
   "paper and system recording at health posts.",
   f"**RMS and programmes: redistribute excess programme stock** ({MAL_OVER:.0f}% of malaria items overstocked) "
   "and review push quantities.",
   "**Districts: follow up the 928 items at or below the emergency order point** and the twenty "
   "weakest facilities in this report.")

# ================================================================== refs
B.append(("unnumbered_h1", "References"))
B.append(("refs", BR.REFS))

if __name__ == "__main__":
    tex = os.path.join(HERE, NAME + ".tex")
    RL.to_tex(B, tex, TITLE, SUBTITLE, META, FIGDIRS)
    pdf, warns = RL.compile_pdf(tex)
    RL.to_docx(B, os.path.join(HERE, NAME + ".docx"), TITLE, SUBTITLE, META, FIGDIRS)
    print("wrote", os.path.basename(pdf), "and .docx;", len([b for b in B if b[0] == "table"]),
          "tables,", len([b for b in B if b[0] == "fig"]), "figures;", len(warns), "LaTeX warnings")
    for w_ in warns[:8]:
        print("  ", w_)
