"""Builds the further-analysis plan in two formats from one content source.

    python3 build_further_analysis.py
      -> Further_analysis_plan.docx   (python-docx)
         Further_analysis_plan.tex    (compile with pdflatex, twice)

Content lives in CONTENT below as plain Python structures, so the Word and
LaTeX versions cannot drift apart. Inline **bold** is supported in both.
"""

import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
TITLE = "Medicine Stock and Prescription Assessment 2026"
SUBTITLE = "Further analysis plan: what else the final dataset can tell us"
META = ("QMIA 2026 (MoH, RMS, RSSB, CHAI) · Prepared by the CHAI Rwanda "
        "analytics team · 5 November 2026 · Draft for discussion")

# ------------------------------------------------------------------
# content blocks: ("h1", text) ("h2", text) ("p", text) ("bullets", [..])
# ("table", header, rows, widths_in_cm) ("analysis", dict)
# ------------------------------------------------------------------
C = []
h1 = lambda t: C.append(("h1", t))
h2 = lambda t: C.append(("h2", t))
p = lambda t: C.append(("p", t))
bl = lambda *t: C.append(("bullets", list(t)))
tb = lambda hdr, rows, w: C.append(("table", hdr, rows, w))
AN = []


def an(code, title, question, why, variables, method, output, caveats,
       objective, priority, effort, done=""):
    d = dict(code=code, title=title, question=question, why=why,
             variables=variables, method=method, output=output,
             caveats=caveats, objective=objective, priority=priority,
             effort=effort, done=done)
    AN.append(d)
    C.append(("analysis", d))


# ================================================================== 1
h1("1. Purpose of this document")
p("The QMIA 2026 assessment collected three linked datasets across 266 "
  "health facilities, 175 community health workers (CHWs) and 6,006 "
  "prescriptions in all 30 districts. The indicator dashboard "
  "(Indicator_dashboard.html) and the preliminary-findings deck answer the "
  "eight objectives of the concept note at headline level. This document "
  "maps what the final, cleaned dataset can still answer, and proposes a "
  "set of further analyses, each tied to a concept-note objective, to the "
  "variables that support it, and to what the dashboard already shows.")
p("Each proposal states the question, why it matters for quality "
  "improvement, the variables used, the method, the output, the caveats, "
  "and a priority and effort rating. Section 6 ranks them into two phases: "
  "analyses that strengthen the national report before it is finalised, "
  "and deeper analyses for publication and for the next QMIA round.")
p("**Reference data.** All figures quoted come from the cleaned dataset "
  "of 5 November 2026 (Data cleaning/Cleaned dataset 2026-11-05), after "
  "cleaning rules S1 to S11 described in 00_Cleaning_log.")

# ================================================================== 2
h1("2. The data available")
h2("2.1 Datasets and units of analysis")
tb(["Dataset", "Unit of analysis", "Records", "Main content"],
   [["Stock assessment: submissions", "Facility visit", "454 visits, 266 facilities",
     "Location, level, management, medicines assessed, GPS (129 facilities), "
     "visit time stamps, data collector, general feedback (301 comments)"],
    ["Stock assessment: medicine rows", "Facility x medicine", "14,378 rows, 80 products (70 molecules)",
     "EML applicability, managed, reason not managed, available on visit day, physical and "
     "system quantity, digital system, monthly consumption and stock-out days (Jun-Aug), "
     "current stock-out days, RMS Central/Regional availability, monthly requested and "
     "delivered quantities, tariff price, RSSB billed quantity and prescriptions/year"],
    ["Prescription: prescriptions", "Prescription", "6,006 prescriptions",
     "Date of prescription, prescriber category, number of medicines, items available "
     "and unavailable, feedback (1,024 comments)"],
    ["Prescription: medicine items", "Prescribed item", "14,059 items, 514 medicine names",
     "Medicine, EML status, drug class, INN or brand, quantity prescribed, available at "
     "facility, ordered, supplied fully/partially/not, unit price, RMS availability, "
     "distribution vs nationwide classification"],
    ["CHW assessment: submissions", "CHW visit", "175 CHWs, 172 health centres",
     "Supervising health centre, district, products assessed"],
    ["CHW assessment: product rows", "CHW x product", "1,988 rows, 13 products",
     "Managed, available, quantity, monthly consumption, months of stock, feedback "
     "(1,015 comments)"]],
   [3.6, 2.6, 3.4, 7.4])
h2("2.2 External files that can be linked")
bl("**Sampling frame** (Facilities.xlsx): 577 facilities by province, district and sector. "
   "This is the basis for sampling weights.",
   "**Facility GIS** (Facility GIS_.xlsx): coordinates for health centres, health posts, "
   "district and referral hospitals; an urban/rural setting field; and district-hospital "
   "referral links (\"Refer to\").",
   "**RMS hub lookup** (analytics/rms_hubs.py): each district mapped to one of five RMS "
   "hubs (Musanze, Kayonza, Kigali, Karongi, Huye).",
   "**National EML** (Final EML list.csv): authorised level of care per product, "
   "already used for the expected-medicine denominator.",
   "**MFT zoning** (analytics/mft_zoning.py): the first-line antimalarial per district "
   "(9 official, 21 inferred until the official list is received).")
h2("2.3 Linkage keys")
p("The datasets join on facility name (harmonised and reattributed during "
  "cleaning), on medicine through the molecule key that combines strengths, "
  "on district for RMS hub and GIS, and on supervising health centre for "
  "CHWs. 168 of the 172 health centres supervising an assessed CHW were "
  "also assessed in the stock form, which allows paired CHW and facility "
  "analysis. Prescription medicine names use a different coding "
  "(e.g. ibuprofen_tab_200_mg) and need a crosswalk to the 70 tracer "
  "molecules before facility x medicine linkage.")
h2("2.4 Completeness of the key fields")
tb(["Field", "Filled", "Implication"],
   [["Available on visit day / consumption / stock-out days", "9,628 managed rows",
     "Complete for all managed rows; the core analyses are well powered."],
    ["Physical and system quantity", "9,162 rows",
     "Inventory accuracy on 8,856 records naming a digital system."],
    ["Current stock-out days", "460 rows",
     "Only for rows out on the visit day; supports a duration analysis."],
    ["RMS Central / Regional availability", "about 465 rows",
     "Only for unavailable rows, as designed."],
    ["RSSB billed quantity, prescriptions per year", "1,005 rows, 215 facilities",
     "Partial; mostly health centres (838 rows). Usable for triangulation."],
    ["March-May stock-out and order fields", "0 rows",
     "Empty on both servers: the window is three months, not six."],
    ["Facility GPS from the form", "129 of 266 facilities",
     "Use the GIS file instead for distance analyses."],
    ["Outpatient / inpatient flag on prescriptions", "not collected",
     "Objective 4 cannot be split by service as the concept note intended."]],
   [6.2, 3.6, 7.2])

# ================================================================== 3
h1("3. What the dashboard already covers")
p("The dashboard has eight pages that recompute in the browser for any "
  "province, district, RMS hub, level, programme, facility or medicine "
  "selection. The table below maps each concept-note objective to what "
  "is shown, the current headline value, and the main gap that a further "
  "analysis would close.")
tb(["Objective", "Dashboard coverage", "Headline (5 Nov)", "Main gap"],
   [["1. Tracer availability", "Page 1: maps 1-2, FIG 1-6, least available medicines, "
     "weakest facilities, reasons not managed",
     "87.8% expected; 93.7% managed",
     "No confidence intervals or weights; no explanation of why facilities differ"],
    ["2. Stock accuracy", "Page 2: FIG 2.5-2.9 (exact match, quantity accuracy)",
     "56.1% exact match",
     "Direction and value of discrepancies (plan 2.5/2.6) not shown"],
    ["2B. Stock adequacy and consumption", "Page 2: FIG 2.1-2.4, tables 2A-2B; Page 3: AMC tables",
     "24.0% adequate; 46.7% overstocked",
     "Trend, non-moving stock, ABC value analysis (plan C4, C6, C9) not shown"],
    ["3. Stock-outs", "Page 4: FIG 4.1-4.6, burden table",
     "9.7% incidence; 7.0% out on visit day",
     "Current-duration distribution (plan 3.5), medicine x district heatmap (3.6), causes"],
    ["4. Unserved prescriptions", "Page 5: FIG 5.1-5.6, tables 5.1-5.5",
     "98.8% complete; 0.8% items unserved",
     "Rate understated by design; prescriber effects; rational-use indicators"],
    ["5. Supply chain attribution", "Page 1: FIG 7-10; table of nationwide stock-outs",
     "85.7% of gaps had stock at RMS",
     "No RMS hub comparison; medicine-level escalation list"],
    ["6. Requisition and fill", "Page 6: FIG 6.1-6.10, table 6.1",
     "93.1% quantity fill rate",
     "Link from under-delivery to later stock-out not tested"],
    ["7. Financial value", "Page 7: Pareto, value by level/class, lost sales",
     "834,418 RWF unserved (observed)",
     "Not combined with overstock and discrepancy value; not annualised"],
    ["8. Community level", "Page 8: FIG 8.1-8.2, CHW vs facility",
     "80.1% CHW availability",
     "Condition-level readiness, CHW stock adequacy, paired CHW-HC analysis"]],
   [3.0, 5.4, 3.4, 5.2])

# ================================================================== 4
h1("4. Proposed further analyses")
p("Analyses are grouped in twelve themes (A to L). Codes are used in the "
  "prioritisation table (Section 6). Priority: **High** = needed for a "
  "robust national report; **Medium** = strengthens interpretation; "
  "**Low** = research or methodological interest. Effort: **L** under "
  "two days, **M** two to five days, **H** more than five days.")

# ---------------- A
h2("Theme A. Making the estimates representative")
an("A1", "Design-based national estimates with 95% confidence intervals",
   "What is the national value of each headline indicator, and how precise is it?",
   "The dashboard reports pooled point estimates. Health posts were added on top of "
   "the proportional sample (one per district) and hospitals are over-represented "
   "relative to the frame, so pooled rates are not self-weighting. The analysis plan "
   "itself asks for \"KPI tile + 95% CI\".",
   "Sampling frame (577 facilities by level and district); facility level, district, "
   "province; every headline indicator at facility x medicine level.",
   "Construct inverse-probability weights per level x district stratum; treat the "
   "facility as the primary sampling unit; estimate ratios with Taylor-linearised "
   "standard errors (survey package in R or statsmodels/samplics in Python). Report "
   "design effects and the intra-facility correlation.",
   "Table of weighted vs unweighted headline indicators with 95% CI; footnote for the "
   "report and deck.",
   "Requires agreement on the frame (577 vs 583 in the concept note). Weights change "
   "little for health-centre-dominated indicators, more for referral-level ones.",
   "All", "High", "M")
an("A2", "Reliability of district estimates",
   "Which district values on the maps are precise enough to rank districts?",
   "District maps drive targeting, but some districts rest on 5 facilities.",
   "District, facility, indicator numerators and denominators.",
   "Cluster-robust CI per district; flag districts whose CI crosses the national "
   "value; optionally small-area (empirical Bayes) shrinkage of district rates.",
   "District table with CI and a reliability flag; map legend note.",
   "With 5 to 11 facilities per district, many differences will not be significant; "
   "this is itself a useful message for targeting.",
   "1, 3, 6, 8", "High", "L")

# ---------------- B
h2("Theme B. Explaining availability")
an("B1", "Drivers of medicine availability (multilevel model)",
   "Which facility and medicine characteristics predict that a managed medicine is on "
   "the shelf?",
   "Descriptive breakdowns show differences by level, management and province one at "
   "a time; a joint model separates them and quantifies how much variation sits at "
   "facility versus district versus medicine level.",
   "Available on visit day (outcome); level, management type, province/RMS hub, "
   "programme, digital system, record accuracy, fill rate and requisition for that "
   "medicine, AMC, distance to RMS hub (GIS), urban/rural setting.",
   "Mixed-effects logistic regression with random intercepts for facility and "
   "district (and a crossed effect for medicine); report odds ratios and the variance "
   "partition coefficient.",
   "Forest plot of adjusted odds ratios; share of variance at each level.",
   "Cross-sectional: associations, not causes. Fill rate and availability are "
   "measured over overlapping periods.",
   "1", "Medium", "M")
an("B2", "Facility performance typology for QI targeting",
   "Which facilities combine several weaknesses, and do they form recognisable groups?",
   "QMIA mentorship needs a short, defensible list of facilities to visit first.",
   "Per facility: availability (managed), record accuracy, stock-out incidence, fill "
   "rate, share overstocked/understocked, prescription completeness.",
   "Standardise each indicator; build a composite index (equal weights and a "
   "sensitivity variant); cluster with k-means or latent profile analysis; test rank "
   "stability by bootstrap.",
   "Ranked facility list with profile labels (e.g. poor records, poor resupply, "
   "overstocked); map of profiles.",
   "Indicators rest on different numbers of medicines per facility; use shrinkage "
   "for small denominators.",
   "1, 2, 3, 6", "High", "M")
an("B3", "Geography and access: distance, remoteness and spatial clustering",
   "Are remote facilities less well supplied, and do weak facilities cluster in space?",
   "Distribution gaps (85.7% of missing items had RMS stock) point to last-mile "
   "delivery, which distance should capture.",
   "Facility GIS coordinates and setting; RMS hub locations; district hospital "
   "referral links; availability, stock-outs and fill rate per facility.",
   "Straight-line or road distance to the RMS branch and to the district hospital; "
   "regression of indicators on distance; Moran's I and Getis-Ord hot-spot analysis "
   "on facility rates.",
   "Scatter of availability against distance; hot-spot map.",
   "Name matching between GIS and assessment lists is needed; road distance requires "
   "an OpenStreetMap routing step.",
   "1, 5", "Medium", "M")
an("B4", "EML compliance: authorised but not managed",
   "Which EML-authorised medicines are not stocked, where, and for what stated reason?",
   "The 5.9-point gap between expected (87.8%) and managed (93.7%) availability is "
   "medicines facilities are authorised for but do not stock; part of it is policy "
   "(MFT zoning, faith-based FP), part is a service gap.",
   "EML applicability, managed, coded reason not managed (table 1.4), medicine, level, "
   "management type.",
   "Classify each not-managed pair as policy (zoning, faith-based FP, strength "
   "substitution) or service gap (no demand, rejected by RSSB, not introduced); rate "
   "by medicine and level.",
   "Table of service-gap medicines by level; list for the EML/RSSB reimbursement "
   "dialogue.",
   "Reasons are free text coded into 11 categories; about 15% remain uncodable.",
   "1", "Medium", "L",
   done="Reasons are coded on page 1 (table 1.4); this adds the policy vs service split.")

# ---------------- C
h2("Theme C. Stock-outs: dynamics and causes")
an("C1", "Does under-delivery lead to a stock-out the following month?",
   "When RMS delivers less than requested in June or July, is the medicine more "
   "likely to be out in the next month?",
   "This is the most direct test of whether order fulfilment drives stock-outs, and "
   "the data hold three monthly observations per pair.",
   "Monthly requested and delivered quantity and monthly stock-out days for June, "
   "July and August; AMC; facility and medicine.",
   "Panel of facility x medicine x month; logistic model of any stock-out in month "
   "t+1 on fill ratio in month t, with facility-medicine fixed effects or random "
   "effects; sensitivity with stock-out days as count outcome.",
   "Odds ratio of a next-month stock-out after a partial or nil delivery.",
   "Only two lagged transitions; reporting of monthly days may be approximate.",
   "3, 6", "High", "M")
an("C2", "Current stock-out duration",
   "How long have current stock-outs lasted, and which medicines stay out longest?",
   "Plan indicator 3.5 (histogram with 30/60/90-day markers) is not on the dashboard.",
   "Current stock-out days (460 rows), medicine, programme, level, RMS availability.",
   "Distribution with 30/60/90-day markers; treat ongoing spells as right-censored and "
   "compare medicines with Kaplan-Meier curves.",
   "Histogram and table of medicines out more than 60 days now.",
   "Recorded only for rows out on the visit day.",
   "3", "Medium", "L")
an("C3", "Stock-out days ratio by medicine and district",
   "Where are stock-out days concentrated?",
   "Plan indicator 3.6 asks for a medicine x district heatmap; the dashboard has the "
   "burden table only.",
   "Monthly stock-out days, managed pairs, medicine, district.",
   "Days out / (pairs x 92) per cell; suppress cells under 5 pairs.",
   "Heatmap for the report annex.",
   "Sparse cells; use molecule level.",
   "3", "Medium", "L")
an("C4", "Combined stock-out indicator: visit day or past three months",
   "What share of managed pairs were out at any point, including the visit day?",
   "Some pairs were empty on the visit day but recorded no stock-out days in "
   "June-August (a check on 3 November found about 180), so the 9.7% incidence misses "
   "them.",
   "Available on visit day, monthly stock-out days.",
   "Pair counted if out on the visit day or with any stock-out day; three-part "
   "breakdown (both, period only and recovered, visit day only).",
   "Revised FIG 4.1 option; a short note on recording completeness.",
   "Label must read \"June-August or on the visit day\".",
   "3", "Medium", "L")
an("C5", "Root-cause attribution of each stock-out",
   "For each stocked-out pair, where did the chain break?",
   "Quality improvement needs a cause per event, not separate rates.",
   "Requested quantity, delivered quantity, RMS Central/Regional availability, "
   "distribution vs nationwide classification, AMC vs request.",
   "Decision tree: not requested; requested but under-quantified relative to AMC; "
   "requested and not delivered although RMS had stock (distribution); nationwide "
   "stock-out; delivered in full but consumed faster than planned.",
   "Waterfall chart of stock-outs by cause, by level and programme.",
   "RMS fields are answered by the facility team, not verified against RMS records; "
   "\"Don't know\" answers form their own category.",
   "3, 5, 6", "High", "M")

# ---------------- D
h2("Theme D. Inventory management and quantification")
an("D1", "Direction and value of record discrepancies",
   "When physical and system counts differ, which way, by how much, and at what value?",
   "Plan indicators 2.5 and 2.6. Physical above system suggests unrecorded receipts; "
   "below suggests unrecorded dispensing, losses or leakage. The dashboard shows only "
   "exact match and quantity accuracy.",
   "Physical quantity, system quantity, digital system, tariff price, level, medicine.",
   "Signed discrepancy and percentage variance; value at tariff; top 15 medicines by "
   "absolute and net value; distribution by system and level.",
   "Diverging bar chart; value table.",
   "Large variances where the system count is near zero; cap or use log ratios.",
   "2", "High", "L")
an("D2", "Record accuracy by digital system",
   "Do some systems (eLMIS, eBuzima, OpenMRS, Medisoft, OpenClinic, IMARA) keep more "
   "accurate records once level and programme are accounted for?",
   "Informs the eBuzima roll-out and refresher-training priorities raised in the "
   "feedback.",
   "Recording system (15 cleaned names), exact match, quantity accuracy, level, "
   "programme, province.",
   "Multilevel logistic model of exact match with system as fixed effect and facility "
   "random effect.",
   "Adjusted accuracy by system.",
   "System is confounded with level (hospitals use different systems).",
   "2", "Medium", "L")
an("D3", "Do poor records lead to poor ordering?",
   "Are facilities with inaccurate records more often stocked out or overstocked?",
   "Links objective 2 to objectives 3 and 6 and justifies reconciliation as a QI action.",
   "Facility-level accuracy; stock-out incidence; adequacy mix; request vs AMC.",
   "Facility-level correlation and regression controlling for level and province.",
   "Scatter and adjusted coefficients.",
   "Ecological at facility level.",
   "2, 3, 6", "Medium", "L")
an("D4", "Ordering behaviour against a min-max rule",
   "How far do actual requests depart from what a min-max rule would order?",
   "FIG 6.9/6.10 show request vs AMC; a min-max comparison tells facilities what they "
   "should have ordered.",
   "AMC, stock on hand, monthly requests; band rule (1-2 months; 1-3 at teaching and referral hospitals).",
   "Ideal order = max level x AMC minus stock on hand; classify each request as under, "
   "in line or over; value the excess at tariff.",
   "Share of requests over/under the rule by level and programme; value of over-ordering.",
   "Stock on hand is observed on the visit day, not at each order date; use as an "
   "approximation.",
   "6", "Medium", "M")
an("D5", "Value of overstock and expiry risk",
   "How much capital is tied up above the maximum stock level?",
   "46.7% of stocked rows are overstocked, 80% for malaria; the value makes the case "
   "for redistribution.",
   "Stock on hand, AMC, tariff price, programme, medicine, months of stock.",
   "Excess = stock above max level; value at tariff; rank medicines and facilities; "
   "flag rows above 6 and 12 months of stock.",
   "Value table and a redistribution candidate list.",
   "No expiry dates were collected; risk is inferred from months of stock.",
   "2B, 7", "High", "L")
an("D6", "Emergency order list",
   "Which facility x medicine pairs need resupply now?",
   "928 understocked items sit at or below 0.5 months, the emergency order point.",
   "Months of stock, medicine, facility, district, RMS hub.",
   "Operational extract sorted by hub and district.",
   "Excel list for RMS branches.",
   "Snapshot at the visit date; needs a refresh before use.",
   "2B", "High", "L")

# ---------------- E
h2("Theme E. Consumption and value")
an("E1", "Consumption trend (August against June)",
   "Is demand rising or falling between June and August?",
   "Plan indicator C4, not on the dashboard.",
   "Monthly consumption, medicine, class, level.",
   "Trend ratio Aug/Jun with median and interquartile range per class; flag outliers.",
   "Small multiples by therapeutic class.",
   "Three months only; seasonality cannot be separated.",
   "2B", "Low", "L")
an("E2", "Non-moving and slow-moving medicines",
   "Which medicines sit on shelves with no consumption?",
   "Plan indicator C6; dead stock is wasted capital and expiry risk.",
   "Monthly consumption, stock on hand, EML applicability, tariff price.",
   "Pairs with zero consumption in all three months and stock above zero; value at tariff.",
   "Bar chart by level and a medicine list.",
   "Some zero-consumption rows may be recording gaps rather than no use.",
   "2B", "Medium", "L")
an("E3", "ABC analysis by consumption value (and ABC-VEN)",
   "Which few medicines account for most of the value consumed?",
   "Plan indicator C9; focuses supervision on high-value items.",
   "Monthly consumption, tariff price, medicine; VEN class from the EML if available.",
   "Pareto of consumption value; A/B/C classes at 80/15/5%; cross with VEN to build "
   "the ABC-VEN matrix.",
   "Pareto chart and ABC-VEN matrix.",
   "VEN classification needs to be sourced.",
   "2B, 7", "Medium", "L")
an("E4", "Triangulating facility consumption with RSSB billing",
   "Does what facilities report consuming match what is billed to RSSB?",
   "A large gap signals unbilled dispensing (revenue loss) or over-reported "
   "consumption (quantification error).",
   "Total quantity billed and prescriptions per year (RSSB), monthly consumption, "
   "quantity per prescription; 1,005 rows at 215 facilities.",
   "Ratio of billed to consumed quantity over the same period; distribution by "
   "medicine and level; outlier review.",
   "Billing completeness ratio by medicine and level.",
   "Billing period must be aligned with June-August; programme products are often "
   "not billed to RSSB.",
   "4, 7", "Medium", "M")

# ---------------- F
h2("Theme F. Prescriptions and rational use")
an("F1", "WHO/INRUD prescribing indicators",
   "How does prescribing compare with WHO reference values?",
   "The prescription data contain everything needed for the standard core "
   "indicators, which the dashboard does not report.",
   "Number of medicines per prescription (mean 2.34), name type (96.2% INN), drug "
   "class (antibacterials), injection forms, EML status, level, prescriber category, "
   "district.",
   "Average medicines per encounter; % generic; % encounters with an antibiotic; % with "
   "an injection; % from the EML; by level and prescriber category, against WHO "
   "reference values (e.g. 1.6-1.8 medicines, under 30% antibiotics, under 20% "
   "injections).",
   "Indicator table by level and prescriber; antibiotic-use map.",
   "Inpatient prescriptions inflate injections; no service flag was collected. "
   "Drug-class coding needs a review (870 items coded as peritoneal dialysis "
   "solutions look misclassified).",
   "4", "High", "M")
an("F2", "Prescriber category and unserved items",
   "Are specialists' or general practitioners' prescriptions more often unserved than "
   "nurses'?",
   "Points to where prescribing and stocking lists diverge (e.g. referral hospitals, "
   "32.9% of prescriptions with a gap).",
   "Prescriber category (nurse 4,884; GP 830; specialist 139; others), unserved items, "
   "level, number of medicines.",
   "Logistic model of a prescription with any unserved item on prescriber category, "
   "controlling for level and number of items, facility random effect.",
   "Adjusted odds ratios.",
   "Few unserved events (107 items), so wide intervals.",
   "4", "Medium", "L")
an("F3", "Concordance between prescription and stock availability",
   "When the stock form says a medicine was unavailable, were prescriptions for it "
   "still written and recorded as served?",
   "Quantifies the design artefact behind the 0.8% unserved rate (prescribers write "
   "what the pharmacy holds) and estimates hidden demand.",
   "Prescription items (medicine, availability, date) and stock rows (medicine, "
   "availability, stock-out days) for the same facility.",
   "Crosswalk prescription names to the 70 tracer molecules; compare item availability "
   "with stock-form availability and stock-out days in the prescription window.",
   "Concordance table and a statement of how understated the unserved rate is.",
   "Requires a name crosswalk; prescription dates span 74 days.",
   "4", "High", "M")
an("F4", "Latent unmet demand from stock-outs",
   "How many prescriptions would have gone unserved during recorded stock-outs?",
   "Gives a realistic counterpart to the observed 0.8%.",
   "RSSB prescriptions per year, quantity per prescription, stock-out days, AMC "
   "(fields prescriptions_missed_stockout, est_qty_unfilled_stockout).",
   "Prescriptions per day x stock-out days per pair; sum and value at tariff; compare "
   "with observed unserved items.",
   "Estimated missed prescriptions and value, by level and medicine.",
   "Relies on RSSB data for 215 facilities; extrapolate with weights (A1).",
   "4, 7", "Medium", "M")
an("F5", "Visit-day against preceding-day prescriptions",
   "Do prescriptions from the visit day differ from those sampled from registers for "
   "earlier days?",
   "Checks for selection and recording bias in the prescription sample.",
   "Date of prescription, visit date, availability, number of items.",
   "Compare unserved rates and items per prescription by day offset.",
   "Short methods note.",
   "Offsets depend on register quality.",
   "4", "Low", "L")
an("F6", "Brand-name prescribing",
   "Which medicines and prescribers use brand names, and at what cost difference?",
   "Brand names and abbreviations were raised in data-collector feedback; brands can "
   "block dispensing of an available generic.",
   "Name type (540 brand items), medicine, prescriber category, level.",
   "Rate by prescriber and level; list of brand-prescribed molecules.",
   "Short table for the rational-use section.",
   "Brand vs INN was judged by the collector.",
   "4", "Low", "L")

# ---------------- G
h2("Theme G. Supply chain attribution and RMS performance")
an("G1", "RMS hub scorecard",
   "How do the five RMS hubs compare on fill rate, distribution gaps and stock-outs?",
   "The dashboard offers an RMS hub filter but no side-by-side comparison; hubs are "
   "the accountable unit for resupply.",
   "District to hub lookup; fill rate, partial/non-fulfilment, delivered without "
   "request, distribution vs nationwide, stock-out incidence.",
   "One row per hub with CI (A1); funnel plot against the national value.",
   "Hub scorecard table and funnel plot.",
   "Hub is assigned by district, not by actual supplying branch.",
   "5, 6", "High", "L")
an("G2", "Medicine-level escalation list",
   "Which medicines are nationwide stock-outs rather than distribution problems?",
   "ARV gaps are half nationwide stock-outs; HIV Determine kits, insulin, thiamine "
   "and paediatric RHZ were out at RMS too.",
   "RMS Central and Regional availability, distribution classification, programme, "
   "medicine; prescription-level RMS fields.",
   "Combine stock-form and prescription attribution; rank medicines by facilities "
   "affected and nationwide share; include \"Don't know\" rates as a data gap.",
   "Escalation table for programmes and central procurement.",
   "Facility-reported RMS status; validate against RMS warehouse records if they can "
   "be obtained.",
   "5", "High", "L")
an("G3", "Validation against RMS warehouse data",
   "Do facility-reported RMS stock statuses match RMS records on the same dates?",
   "The concept note planned RMS warehouse data as a source; this has not been linked.",
   "RMS stock and delivery records (to request); stock-form RMS fields.",
   "Match by medicine, branch and date; agreement statistics (kappa).",
   "Validation note; corrected attribution if needed.",
   "Depends on data access from RMS.",
   "5", "Medium", "M")

# ---------------- H
h2("Theme H. Financial value")
an("H1", "One table of the cost of supply chain gaps",
   "What is the combined value of unserved prescriptions, lost sales from stock-outs, "
   "overstock and record discrepancies?",
   "Each value is computed separately today; decision-makers need one number with "
   "clear components.",
   "Unserved value (834,418 RWF observed), lost sales (page 7), overstock value (D5), "
   "discrepancy value (D1), tariff prices.",
   "Harmonise to the same period and prices; annualise and extrapolate to the 577 "
   "facilities with weights (A1); show ranges.",
   "Summary table and waterfall chart for the executive summary.",
   "Components are not additive in every case (lost sales and unserved value "
   "overlap); present them side by side, not summed.",
   "7", "High", "M")

# ---------------- I
h2("Theme I. Community level")
an("I1", "Condition-level treatment readiness of CHWs",
   "What share of CHWs can treat each iCCM condition on the visit day?",
   "Product availability (80.1%) hides whether a CHW can complete a treatment: "
   "malaria needs an RDT and an age-appropriate ACT; diarrhoea needs ORS and zinc; "
   "pneumonia needs amoxicillin.",
   "CHW product rows: RDT, AL presentations, DHAP, pyronaridine/artesunate, ORS, "
   "zinc, amoxicillin 125 mg; district zone.",
   "Per CHW, ready for a condition if every required product is managed and "
   "available (zone-appropriate ACT); rate by district.",
   "Readiness by condition and district.",
   "CHW cadre (binome/ASM) was not recorded, so plan indicator 8.3 by cadre is not "
   "possible.",
   "8", "High", "L")
an("I2", "CHW stock adequacy",
   "Are CHWs over- or under-stocked relative to their consumption?",
   "Stock adequacy is reported for facilities but not for CHWs.",
   "Months of stock (703 rows), quantity, monthly consumption.",
   "Same bands as facilities; by product and district.",
   "Adequacy mix for CHWs.",
   "Small quantities make months of stock unstable.",
   "8", "Medium", "L")
an("I3", "Test-and-treat proxy",
   "Do CHWs dispense ACTs in proportion to RDTs used?",
   "A high ACT to RDT ratio suggests treatment without testing.",
   "Monthly consumption of RDTs and of all ACT presentations per CHW.",
   "Ratio of ACT courses to RDTs per CHW and district.",
   "Distribution and district map.",
   "Positivity varies by area; ratio is a screening signal, not proof.",
   "8", "Medium", "L")
an("I4", "Paired CHW and supervising health centre",
   "When the health centre is stocked out, is its CHW too? Does the health centre hold "
   "stock its CHW lacks?",
   "Tests the resupply link from health centre to community (168 linked pairs).",
   "CHW product availability and quantity; same product at the supervising health "
   "centre (availability, months of stock).",
   "Paired cross-tabulation and McNemar test per product; list of health centres "
   "overstocked while their CHW is out.",
   "Paired table; redistribution list.",
   "Only eight products overlap between the two forms.",
   "8", "High", "L",
   done="Page 8 compares CHW and facility availability in aggregate; this pairs them.")

# ---------------- J
h2("Theme J. Qualitative feedback")
an("J1", "Thematic analysis of the remaining free text",
   "What do facility teams and CHWs say about supply problems?",
   "Prescription-form feedback is coded (table 5.5); the stock-form (301) and CHW "
   "(1,015) comments are not.",
   "Feedback fields on the three forms; facility, level, district.",
   "Code a frame from the existing prescription themes; double-code a 20% sample for "
   "agreement; link themes to quantitative outcomes (e.g. facilities citing late "
   "deliveries vs their fill rate).",
   "Theme frequency table with illustrative quotes; joint display with indicators.",
   "Many CHW comments are short (\"few patients\"); some are in Kinyarwanda.",
   "All", "Medium", "M")

# ---------------- K
h2("Theme K. Data quality and process")
an("K1", "Data-collector and device effects",
   "Do indicators vary with the data collector more than expected?",
   "45 collectors on 78 devices; collector effects would bias district comparisons.",
   "Collector, device, form version (14), interview start/end (median 154 minutes), "
   "indicators.",
   "Intra-collector correlation from a cross-classified model; interview duration "
   "against completeness; digit preference in counts.",
   "Methods annex table.",
   "Collectors and districts are partly confounded by team assignment.",
   "Methods", "Medium", "L")
an("K2", "Cleaning and consistency report",
   "How much did each cleaning rule change, and what remains uncertain?",
   "Rules S1-S11 changed availability, management and consumption fields; reviewers "
   "will ask.",
   "Cleaning log, before/after fields (available, managed, applicable).",
   "Table of rule, rows affected, effect on headline indicators.",
   "Methods annex.",
   "None.",
   "Methods", "High", "L",
   done="The cleaning log exists; this turns it into a reportable annex.")

# ---------------- L
h2("Theme L. Policy questions already flagged")
an("L1", "MFT zoning compliance", "Do facilities and CHWs hold the antimalarial of "
   "their zone?",
   "Out-of-zone products were recoded (S10); compliance is a programme indicator.",
   "Antimalarial rows, district, mft_zoning lookup.",
   "Final once the official list for all 30 districts is received.",
   "Compliance by district; flag list (Antimalarial_zoning_flags).",
   "21 districts' zones are inferred from the data.",
   "1, 8", "High", "L", done="Flag workbook produced; awaiting the official list.")
an("L2", "Strength substitution", "Where is one strength stocked out while another "
   "is available?",
   "154 cases at 101 facilities (91 out for more than 60 days, DHAP and AL excluded); "
   "informs which strengths each level should hold.",
   "Strength-level rows, molecule key, stock-out days.",
   "Already extracted; add a policy recommendation per molecule.",
   "Strength_mismatch_extract and Strength_flags_over60.",
   "None.", "1", "Medium", "L", done="Extracts produced (S9).")

# ================================================================== 5
h1("5. Suggested analytical approach")
h2("5.1 Software and reproducibility")
bl("Keep the existing Python pipeline (analyze_* loaders, Data_cleaning_pipeline.py) "
   "as the single source of cleaned data; every new analysis reads the dated cleaned "
   "files, never the raw exports.",
   "Survey estimation (A1, A2, G1) in R (survey, srvyr) or Python (samplics); mixed "
   "models (B1, C1, D2, F2) in R (lme4, glmmTMB) or Python (statsmodels).",
   "Spatial work (B3) in R (sf, spdep) or Python (geopandas, libpysal).",
   "Add new outputs as dashboard pages only where they will be monitored; one-off "
   "analyses go to the report annex.")
h2("5.2 Conventions to keep")
bl("Availability uses two tracks: Essential Medicines on the EML denominator, "
   "programme products on managed rows only.",
   "CHW availability is computed over managed rows only (not managed is MFT zoning).",
   "Average stock-out days use total_stockout_days; incidence and lost sales use the "
   "monthly fields.",
   "Prescription unserved rates are understated by design and should never be read "
   "as a supply rate on their own.",
   "Report counts beside every percentage and suppress groups under 30 observations.")

# ================================================================== 6
h1("6. Prioritisation")
p("Phase 1 analyses strengthen the national report and the dissemination; "
  "Phase 2 analyses support publication, the QI plan and the next round.")


def _phase(a):
    return "1" if a["priority"] == "High" else "2"


tb(["Code", "Analysis", "Objective", "Priority", "Effort", "Phase"],
   [[a["code"], a["title"], a["objective"], a["priority"], a["effort"], _phase(a)]
    for a in AN],
   [1.3, 8.2, 2.2, 1.9, 1.5, 1.5])

# ================================================================== 7
h1("7. Limitations of the dataset")
bl("**Three-month window.** March-May fields are empty, so frequency and trends rest "
   "on June-August only.",
   "**No service flag.** Prescriptions cannot be split into outpatient and inpatient, "
   "as objective 4 intended.",
   "**Prescription design.** Prescribers tend to write what the pharmacy holds; "
   "unserved rates understate unmet need (see F3, F4).",
   "**Facility-reported RMS status.** RMS availability is reported at the facility and "
   "not yet verified against RMS warehouse records (G3).",
   "**Partial RSSB data.** RSSB fields cover 215 facilities, mostly health centres.",
   "**Zoning.** 21 of 30 districts' antimalarial zones are inferred until the official "
   "list is received; S10 results may change.",
   "**CHW sample.** 175 CHWs against 237 planned, no cadre field, and 29 districts "
   "(Kirehe has no CHW data).",
   "**Representativeness.** Pooled estimates are unweighted until A1 is done.")

# ================================================================== 8
h1("8. Next steps")
bl("Agree the Phase 1 list and owners at the dissemination workshop.",
   "Confirm the sampling frame and compute weights (A1) first, since several analyses "
   "reuse them.",
   "Request the official MFT zoning list and RMS warehouse extracts (L1, G3).",
   "Build the prescription-to-tracer name crosswalk (F3), which unlocks several "
   "prescription analyses.",
   "Produce Phase 1 outputs as a report annex, and add the RMS hub scorecard and the "
   "CHW readiness indicators to the dashboard if they will be monitored.")

# ================================================================== annex
h1("Annex A. Dashboard index")
tb(["Page", "Figures and tables"],
   [["1. Tracer medicine availability", "MAP 1-2; FIG 1-10; all medicines by availability; "
     "least available medicines; weakest facilities; nationwide stock-outs; reasons not managed"],
    ["2. Stock adequacy and inventory accuracy", "FIG 2.1-2.9; tables 2A, 2B"],
    ["3. Medicine consumption", "Tables 3.1-3.4 (AMC highest, lowest, every medicine, by programme)"],
    ["4. Stock-out frequency and duration", "FIG 4.1-4.6; requested or not; stock-out burden by medicine"],
    ["5. Prescription fulfilment", "FIG 5.1-5.6; table 5A; tables 5.1-5.5 (incl. feedback themes)"],
    ["6. Requisition and order fulfilment", "FIG 6.1-6.10; table 6A; most under-supplied medicines"],
    ["7. Financial value", "Pareto of unserved value; value by level, province, class; lost sales (7.4-7.9)"],
    ["8. Community level", "FIG 8.1-8.2; CHW against facility on shared products"]],
   [5.5, 11.5])

h1("Annex B. Key variables by dataset")
tb(["Dataset", "Variable (cleaned name or form question)", "Use"],
   [["Stock rows", "Is ... applicable as per EML / managed / available on visit day", "Availability tracks"],
    ["Stock rows", "Quantity available; quantity in the system; recording_system_clean", "Accuracy"],
    ["Stock rows", "June/July/August consumption; avg_monthly_consumption; stock_level_months", "Adequacy, ABC, trend"],
    ["Stock rows", "June/July/August stock-out days; total_stockout_days; current stock-out days", "Stock-outs"],
    ["Stock rows", "Available at RMS Central / Regional; distribution_nationwide", "Attribution"],
    ["Stock rows", "June/July/August requested and delivered; total_requested_qty; total_delivered_qty", "Fill rate, ordering"],
    ["Stock rows", "unit_tariff_price; financial_loss; Total Quantity Billed; Prescriptions/year (RSSB)", "Value, triangulation"],
    ["Stock rows", "Why ... not managed; molecule_key; eml_levels", "EML compliance, strengths"],
    ["Stock submissions", "Facility, level, management, district, start/end, collector, deviceid, GPS, feedback", "Design, quality, geography"],
    ["Prescriptions", "Date; prescriber category; number of medicines; n_items_unavailable", "Fulfilment, rational use"],
    ["Prescription items", "Medicine; on_eml; drug_class_cleaned; Name type; Qty prescribed; availability", "Rational use, concordance"],
    ["Prescription items", "Ordered; supplied fully/partially/not; unit price; RMS availability", "Attribution, value"],
    ["CHW rows", "Managed; available; quantity; monthly consumption; stock_level_months; feedback", "Readiness, adequacy"]],
   [3.2, 9.8, 4.0])


# ------------------------------------------------------------------
# Word
# ------------------------------------------------------------------
def build_docx(path):
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor

    NAVY = RGBColor(0x1F, 0x3A, 0x68)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Cm(2.0))
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    st = doc.styles
    st["Normal"].font.name = "Arial"
    st["Normal"].font.size = Pt(10)
    st["Normal"].paragraph_format.space_after = Pt(5)
    st["Normal"].element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    for name, size in (("Heading 1", 15), ("Heading 2", 12), ("Heading 3", 10.5)):
        s = st[name]
        s.font.name, s.font.size, s.font.bold = "Arial", Pt(size), True
        s.font.color.rgb = NAVY
        s.element.rPr.rFonts.set(qn("w:asciiTheme"), "")
        s.paragraph_format.space_before = Pt(12 if name != "Heading 3" else 8)
        s.paragraph_format.space_after = Pt(4)
        s.paragraph_format.keep_with_next = True

    def runs(par, text, size=None, color=None):
        for i, part in enumerate(re.split(r"\*\*", text)):
            if part:
                r = par.add_run(part)
                r.bold = bool(i % 2)
                if size:
                    r.font.size = Pt(size)
                if color:
                    r.font.color.rgb = color
        return par

    def shade(cell, hexcol):
        tcPr = cell._tc.get_or_add_tcPr()
        sh = OxmlElement("w:shd")
        sh.set(qn("w:val"), "clear")
        sh.set(qn("w:color"), "auto")
        sh.set(qn("w:fill"), hexcol)
        tcPr.append(sh)

    def table(hdr, rows, widths):
        t = doc.add_table(rows=1, cols=len(hdr))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        for j, h in enumerate(hdr):
            c = t.rows[0].cells[j]
            c.text = ""
            runs(c.paragraphs[0], "**" + h + "**", 8.5, RGBColor(255, 255, 255))
            shade(c, "1F3A68")
        for k, row in enumerate(rows):
            cells = t.add_row().cells
            for j, v in enumerate(row):
                cells[j].text = ""
                runs(cells[j].paragraphs[0], str(v), 8.5)
                if k % 2:
                    shade(cells[j], "EEF2F7")
        for row in t.rows:
            for j, w in enumerate(widths):
                row.cells[j].width = Cm(w)
            for c in row.cells:
                for par in c.paragraphs:
                    par.paragraph_format.space_after = Pt(1)
        # repeat header row
        trPr = t.rows[0]._tr.get_or_add_trPr()
        el = OxmlElement("w:tblHeader")
        el.set(qn("w:val"), "true")
        trPr.append(el)
        doc.add_paragraph()

    # title block
    tp = doc.add_paragraph()
    tp.paragraph_format.space_before = Pt(60)
    runs(tp, "**" + TITLE + "**", 22, NAVY)
    runs(doc.add_paragraph(), SUBTITLE, 15, NAVY)
    runs(doc.add_paragraph(), META, 9.5, RGBColor(0x55, 0x55, 0x55))
    q = doc.add_paragraph()
    q.paragraph_format.space_before = Pt(24)
    runs(q, "**Contents.** 1 Purpose · 2 The data available · 3 What the dashboard "
         "already covers · 4 Proposed further analyses (themes A-L, "
         f"{len(AN)} analyses) · 5 Analytical approach · 6 Prioritisation · "
         "7 Limitations · 8 Next steps · Annexes", 9.5)
    doc.add_page_break()

    for blk in C:
        kind = blk[0]
        if kind == "h1":
            doc.add_heading(blk[1], level=1)
        elif kind == "h2":
            doc.add_heading(blk[1], level=2)
        elif kind == "p":
            runs(doc.add_paragraph(), blk[1])
        elif kind == "bullets":
            for b in blk[1]:
                runs(doc.add_paragraph(style="List Bullet"), b)
        elif kind == "table":
            table(blk[1], blk[2], blk[3])
        elif kind == "analysis":
            a = blk[1]
            doc.add_heading(f"{a['code']}. {a['title']}", level=3)
            tag = doc.add_paragraph()
            runs(tag, f"Objective {a['objective']} · Priority **{a['priority']}** · "
                      f"Effort **{a['effort']}**", 8.5, RGBColor(0x55, 0x55, 0x55))
            for lab, key in (("Question", "question"), ("Why it matters", "why"),
                             ("Variables", "variables"), ("Method", "method"),
                             ("Output", "output"), ("Caveats", "caveats"),
                             ("Already done", "done")):
                if a[key]:
                    par = doc.add_paragraph()
                    par.paragraph_format.left_indent = Cm(0.4)
                    par.paragraph_format.space_after = Pt(2)
                    runs(par, f"**{lab}.** {a[key]}")
    # footer page numbers
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = fp.add_run()
    for t_, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if t_:
            f = OxmlElement("w:fldChar")
            f.set(qn("w:fldCharType"), t_)
            r._r.append(f)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = txt
            r._r.append(it)
    r.font.size = Pt(8)
    doc.save(path)


# ------------------------------------------------------------------
# LaTeX
# ------------------------------------------------------------------
def tex_esc(s):
    s = str(s)
    rep = {"\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$",
           "#": r"\#", "_": r"\_", "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}",
           "^": r"\textasciicircum{}"}
    s = "".join(rep.get(ch, ch) for ch in s)
    s = s.replace("·", r"\textperiodcentered{}").replace("\"", "''")
    s = re.sub(r"\*\*(.+?)\*\*", r"\\textbf{\1}", s)
    s = re.sub(r"(\d) x (\d)", r"\1 $\\times$ \2", s)
    return s


def build_tex(path):
    out = [r"""\documentclass[10pt,a4paper]{article}
\usepackage[T1]{fontenc}
\usepackage[utf8]{inputenc}
\usepackage[margin=2cm]{geometry}
\usepackage{helvet}
\renewcommand{\familydefault}{\sfdefault}
\usepackage{booktabs,longtable,array,xcolor,colortbl,enumitem,titlesec}
\usepackage[hidelinks]{hyperref}
\definecolor{navy}{HTML}{1F3A68}
\definecolor{band}{HTML}{EEF2F7}
\titleformat{\section}{\large\bfseries\color{navy}}{}{0pt}{}
\titleformat{\subsection}{\normalsize\bfseries\color{navy}}{}{0pt}{}
\titleformat{\subsubsection}{\small\bfseries\color{navy}}{}{0pt}{}
\setlist{nosep,leftmargin=1.2em}
\setlength{\parindent}{0pt}\setlength{\parskip}{4pt}
\newcommand{\field}[2]{\par\hangindent=1em\textbf{#1.} #2}
\begin{document}
\begin{titlepage}\vspace*{3cm}
{\Huge\bfseries\color{navy} """ + tex_esc(TITLE) + r"""\par}\vspace{8pt}
{\Large\color{navy} """ + tex_esc(SUBTITLE) + r"""\par}\vspace{14pt}
{\small\color{gray} """ + tex_esc(META) + r"""\par}
\end{titlepage}
\tableofcontents\newpage
"""]
    for blk in C:
        kind = blk[0]
        if kind == "h1":
            out.append(r"\section{" + tex_esc(re.sub(r"^\d+\.\s*", "", blk[1])
                                             if not blk[1].startswith("Annex")
                                             else blk[1]) + "}")
        elif kind == "h2":
            out.append(r"\subsection{" + tex_esc(blk[1]) + "}")
        elif kind == "p":
            out.append(tex_esc(blk[1]) + "\n")
        elif kind == "bullets":
            out.append(r"\begin{itemize}" + "".join(
                r"\item " + tex_esc(b) + "\n" for b in blk[1]) + r"\end{itemize}")
        elif kind == "table":
            hdr, rows, w = blk[1], blk[2], blk[3]
            tot = sum(w)
            spec = "".join(r">{\raggedright\arraybackslash}p{" +
                           f"{0.98 * x / tot:.3f}" + r"\linewidth}" for x in w)
            head = " & ".join(r"\textbf{\color{white}" + tex_esc(h) + "}" for h in hdr)
            body = []
            for k, r in enumerate(rows):
                pre = r"\rowcolor{band}" if k % 2 else ""
                body.append(pre + " & ".join(tex_esc(v) for v in r) + r" \\")
            out.append(r"{\footnotesize\setlength{\tabcolsep}{3pt}" "\n"
                       r"\begin{longtable}{" + spec + "}\n"
                       r"\rowcolor{navy}" + head + r" \\" + "\n" r"\endhead" "\n"
                       + "\n".join(body) + "\n"
                       r"\end{longtable}}")
        elif kind == "analysis":
            a = blk[1]
            out.append(r"\subsubsection*{" + tex_esc(f"{a['code']}. {a['title']}") + "}")
            out.append(r"{\footnotesize\color{gray} Objective " + tex_esc(a["objective"])
                       + r" \textperiodcentered{} Priority \textbf{" + a["priority"]
                       + r"} \textperiodcentered{} Effort \textbf{" + a["effort"] + "}}")
            for lab, key in (("Question", "question"), ("Why it matters", "why"),
                             ("Variables", "variables"), ("Method", "method"),
                             ("Output", "output"), ("Caveats", "caveats"),
                             ("Already done", "done")):
                if a[key]:
                    out.append(r"\field{" + lab + "}{" + tex_esc(a[key]) + "}")
    out.append(r"\end{document}")
    open(path, "w", encoding="utf-8").write("\n".join(out))


if __name__ == "__main__":
    build_docx(os.path.join(HERE, "Further_analysis_plan.docx"))
    build_tex(os.path.join(HERE, "Further_analysis_plan.tex"))
    print(f"{len(AN)} analyses; wrote Further_analysis_plan.docx and .tex in {HERE}")
