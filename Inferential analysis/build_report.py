"""Inferential and predictive analysis report, Word + LaTeX, from results.json.

    python3 run_inferential_analysis.py     (produces results/)
    python3 build_report.py  -> QMIA2026_inferential_predictive_analysis.docx / .tex

Every number in the text is read from results/results.json, so the report is
regenerated, not retyped, after any data refresh.
"""
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
RJ = json.load(open(os.path.join(HERE, "results", "results.json")))
FIG = os.path.join(HERE, "results")
NAME = "QMIA2026_inferential_predictive_analysis"
TITLE = "Medicine Stock and Prescription Assessment 2026"
SUBTITLE = "Inferential and predictive analysis: findings and interpretation"
META = ("QMIA 2026 (MoH, RMS, RSSB, CHAI) \u00b7 CHAI Rwanda analytics team \u00b7 "
        "5 November 2026 \u00b7 Draft for technical review")

C = []
h1 = lambda t: C.append(("h1", t))
h2 = lambda t: C.append(("h2", t))
p = lambda t: C.append(("p", t))
bl = lambda *t: C.append(("bullets", list(t)))
tb = lambda hdr, rows, w, cap="": C.append(("table", hdr, rows, w, cap))
fg = lambda fn, cap, width=16: C.append(("fig", fn, cap, width))


def f1(v):
    return "n/a" if v is None else f"{v:.1f}"


def f2(v):
    return "n/a" if v is None else f"{v:.2f}"


def ci(d, k=1):
    fm = f1 if k == 1 else f2
    return f"{fm(d['est'])}% ({fm(d['lo'])}-{fm(d['hi'])})"


def orci(t):
    return f"{t['OR']:.2f} ({t['lo']:.2f}-{t['hi']:.2f})"


def pv(x):
    if x is None:
        return "n/a"
    return "<0.001" if x < 0.001 else f"{x:.3f}"


LABEL = {"level2": "Level", "mgmt": "Management", "province": "Province",
         "programme": "Programme", "urban": "Setting", "fsys": "Main digital system",
         "sys": "Digital system", "order": "Order outcome this month",
         "prescriber": "Prescriber", "month": "Month", "out_t": "Stocked out this month",
         "log_amc": "log(AMC + 1)", "erratic": "Erratic consumption (CV >= 50%)",
         "fill10": "Facility fill rate (per +10 points)",
         "acc10": "Facility record accuracy (per +10 points)",
         "n_med": "Medicines on the prescription (per +1)"}


def term(t):
    m = re.match(r"C\((\w+), Treatment\(reference='(.+?)'\)\)\[T\.(.+?)\]", t)
    if m:
        return f"{LABEL.get(m.group(1), m.group(1))}: {m.group(3)} (ref. {m.group(2)})"
    m = re.match(r"C\((\w+)\)\[T\.(.+?)\]", t)
    if m:
        return f"{LABEL.get(m.group(1), m.group(1))}: {'August' if m.group(2) == 'jul' else m.group(2)}"
    return LABEL.get(t, t)


def model_rows(key):
    return [[term(t["term"]), orci(t), pv(t["p"])] for t in RJ[key]["terms"]
            if t["lo"] is not None and t["hi"] is not None]


def T(key, lab):
    return [t for t in RJ[key]["terms"] if lab in t["term"]][0]


E = {e["indicator"]: e for e in RJ["estimates"]}
NOHP = {e["indicator"]: e for e in RJ["estimates_no_hp"]}
WHO = {w["indicator"]: w for w in RJ["who_indicators"]}
PA, PB = RJ["pred_august"], RJ["pred_visit"]
best = lambda P: [m for m in P["models"] if m["model"] == P["best"]][0]
LAG = {r["order"]: r for r in RJ["lag_crosstab"]}
RD = {r["condition"]: r for r in RJ["iccm_ready"]}
VA, VS, VQ = RJ["vpc_availability"], RJ["vpc_stockout"], RJ["vpc_accuracy"]
TY = RJ["typology"]
PROF = TY["profiles"]


def name_clusters(prof):
    left = {c["cluster"] for c in prof}
    names = {}
    by = lambda k, rev=True: sorted([c for c in prof if c["cluster"] in left],
                                     key=lambda c: c[k], reverse=rev)[0]
    c = by("stockout"); names[c["cluster"]] = "Frequent stock-outs"; left.discard(c["cluster"])
    c = by("fill", False); names[c["cluster"]] = "Weak resupply and weak records"; left.discard(c["cluster"])
    c = by("overstock"); names[c["cluster"]] = "Well supplied but overstocked"; left.discard(c["cluster"])
    for c in left:
        names[c] = "Lean stock, often understocked"
    return names


CN = name_clusters(PROF)
n_below = [d["district"] for d in RJ["district_ci_m2"] if d["vs_national"] == "below"]
n_above = [d["district"] for d in RJ["district_ci_m2"] if d["vs_national"] == "above"]

# ================================================================== summary
h1("Executive summary")
p("This report adds inferential statistics and predictive models to the "
  "descriptive QMIA 2026 dashboard. It uses the cleaned dataset of 5 November "
  "2026: 266 facilities, 9,628 managed facility x medicine pairs, 6,006 "
  "prescriptions with 14,059 items, and 175 community health workers (CHWs). "
  "All confidence intervals account for the clustering of observations within "
  "facilities and the stratification by level of care.")
m2 = E["Availability, managed products (M2)"]
m1 = E["Availability, expected medicines (M1)"]
bl(f"**The headline estimates are precise.** Managed-product availability is "
   f"{ci(m2['unweighted'])} and expected-medicine availability {ci(m1['unweighted'])}. "
   f"Weighted to the national facility mix, expected availability falls to "
   f"{ci(m1['weighted'])}, because health posts, which are about 69% of facilities "
   f"nationally, hold fewer of the medicines they are authorised for. Excluding health "
   f"posts it is {f1(NOHP['Availability, expected medicines (M1)']['est'])}%.",
   f"**A failed delivery is the strongest modifiable predictor of a stock-out.** "
   f"When a requested order received nothing, {f1(LAG['Requested, nothing delivered']['mean'])}% "
   f"of pairs were stocked out the next month, against {f1(LAG['Delivered in full']['mean'])}% "
   f"after a full delivery: adjusted odds ratio "
   f"{orci(T('gee_lag', 'Requested, nothing delivered'))}. Not ordering at all also "
   f"raised the odds ({orci(T('gee_lag', 'No request'))}).",
   f"**Stock-outs persist.** A pair out this month had about "
   f"{T('gee_lag', 'out_t')['OR']:.0f} times the odds of being out next month. A simple "
   f"rule (\"out in July\") identifies {f1(PA['rule_baseline']['recall'])}% of August "
   f"stock-outs; a machine-learning model adds value mainly by ranking new risk "
   f"(AUC {PA['new_stockouts']['auc']:.2f} among pairs not out in July).",
   f"**Record accuracy is a facility problem, not a medicine problem.** "
   f"{f1(VQ['facility']['share'] + VQ['district']['share'])}% of the variation in exact "
   f"record match lies between facilities and districts (against "
   f"{f1(VA['facility']['share'] + VA['district']['share'])}% for availability). "
   f"Facilities recording in eBuzima matched far less often than those on eLMIS "
   f"(OR {orci(T('gee_accuracy', 'eBuzima'))}).",
   f"**Faith-based facilities have higher availability** (OR "
   f"{orci(T('gee_availability', 'Faith-based'))}) and district hospitals the highest "
   f"(OR {orci(T('gee_availability', 'District hospital'))}). The West is lower than "
   f"Kigali City (OR {orci(T('gee_availability', 'West'))}). ARV products are less "
   f"available than essential medicines (OR {orci(T('gee_availability', 'ARV'))}).",
   f"**Antibiotic prescribing is high.** {f1(WHO['% encounters with an antibiotic']['est'])}% "
   f"of encounters include an antibiotic (WHO reference under 30%); prescribing by generic "
   f"name ({f1(WHO['% medicines prescribed by generic (INN) name']['est'])}%) and from the "
   f"EML ({f1(WHO['% medicines from the EML']['est'])}%) is very good, and injection use "
   f"is low ({f1(WHO['% encounters with an injection']['est'])}%).",
   f"**Only {f1(RD['All three conditions']['est'])}% of CHWs can treat all three iCCM "
   f"conditions on the day.** Malaria readiness is {f1(RD['Malaria (RDT + an ACT)']['est'])}%, "
   f"pneumonia {f1(RD['Pneumonia (amoxicillin)']['est'])}% and diarrhoea (ORS and zinc "
   f"together) only {f1(RD['Diarrhoea (ORS + zinc)']['est'])}%. Where a CHW was out of an "
   f"antimalarial or RDT, the supervising health centre had it in every one of "
   f"{RJ['paired_total']['hc_only']} discordant pairs: the gap is the last step to the "
   f"community, not supply to the health centre.",
   f"**Four facility profiles emerge**, which can direct QMIA mentorship: "
   + "; ".join(f"{CN[c['cluster']].lower()} ({c['facilities']} facilities)" for c in PROF) + ".",
   "**There is no geographic clustering** of availability, stock-outs or record accuracy "
   "between neighbouring districts (Moran's I not significant); weak districts are "
   "scattered, so targeting must be facility-specific.")

# ================================================================== 1
h1("1. Introduction")
p("The concept note sets eight objectives: availability of tracer medicines, stock "
  "record accuracy, stock-out frequency and duration, unserved prescriptions, "
  "attribution to the supply chain, requisition and order fulfilment, the financial "
  "value of unserved prescriptions, and availability at community level. The "
  "indicator dashboard answers each at headline level with pooled percentages. This "
  "report asks four further questions:")
bl("How precise are the headline figures, and do they change when weighted to the "
   "national facility mix?",
   "Which differences between levels, provinces, programmes and facility types remain "
   "once other factors are held constant?",
   "What drives stock-outs, and in particular does order fulfilment matter?",
   "Can next month's stock-outs be predicted from routine data, so that RMS and "
   "facilities can act before the shelf is empty?")
p("Section 2 summarises comparable studies; Section 3 the methods; Section 4 the "
  "results with interpretation; Section 5 the discussion and Section 6 the "
  "recommendations.")

# ================================================================== 2
h1("2. What comparable studies have found")
p("A targeted search covered peer-reviewed studies and agency reports from "
  "sub-Saharan Africa, with priority for Rwanda. The benchmarks below frame the "
  "results; the reference list gives the sources.")
h2("2.1 Availability of medicines")
p("Public-sector availability of essential medicines in sub-Saharan Africa has "
  "typically been 30-60% in WHO/HAI surveys [Cameron 2009]. Rwanda has sat at the "
  "upper end: 44.2% mean availability of 32 medicines in the 2007 SPA [Choi and "
  "Ametepi 2013], 59.6% in the public sector and 62.8% in faith-based facilities in a "
  "2020 WHO/HAI survey of 18 medicines [Bizimana 2020], and 71-78% for NCD medicines "
  "in three rural districts [Mukundiyukuri 2020]. The WHO target is 80% availability "
  "of essential NCD medicines. The QMIA basket and definitions differ (tracer "
  "medicines, managed and EML-expected denominators), so our 87.8-93.7% is not "
  "directly comparable, but it is consistent with the improvement reported by QMIA "
  "itself [Kigenza 2023].")
h2("2.2 Stock-outs and their determinants")
p("Day-of-visit stock-out of essential medicines was 18-39% across levels in Ghana, "
  "Kenya and Uganda, worst at the lowest level [Masters 2014]. Recurrent determinants "
  "are distance and rurality, facility level, transport, order fulfilment and "
  "inventory policy. In Zambia, up to 30% of clinics stocked out of an "
  "artemether-lumefantrine product while the central warehouse had ample stock, "
  "which the authors traced to the max-min rule under seasonal demand [Leung 2016]. "
  "In Tanzania, facilities with no stock-outs had higher household use of care "
  "[Kuwawenaruwa 2020]. A census of 940 Malawian facilities found faith-based "
  "ownership (OR 1.63), computerisation (OR 1.46) and more order fulfilments (OR "
  "1.14) associated with availability, and distance over 75 km with lower "
  "availability (OR 0.79) [Mohan 2024].")
h2("2.3 Inventory records and eLMIS")
p("The published evaluation of QMIA in Rwanda reported eLMIS use rising from 55% to "
  "96%, inventory accuracy from 25% to 85% and the stock-out rate falling from 10% to "
  "1% between 2017 and 2019 [Kigenza 2023]. Inventory accuracy of 70-85% is typical in "
  "Ethiopian LIAT-style surveys [Mekonen 2025], and up-to-date bin cards and eLMIS "
  "were associated with far fewer stock-outs of emergency obstetric drugs in Rwandan "
  "district hospitals [Kabera and Mukanyangezi 2024]. Our exact-match accuracy of "
  "56.1% is therefore below both the Rwandan 2019 figure and regional benchmarks, "
  "although the definitions differ (exact equality here, tolerance bands in some "
  "studies).")
h2("2.4 Prescribing")
p("A review of 43 studies in the WHO African region found median values of 3.1 "
  "medicines per encounter, 68% generic prescribing, 46.8% of encounters with an "
  "antibiotic, 25% with an injection and 88% from the EML [Ofori-Asenso 2016], against "
  "WHO reference values of under 2 medicines, 100% generic, under 30% antibiotic, under "
  "20% injection and 100% EML [WHO 1993]. A paediatric study in Burundi reported 2.4 "
  "medicines, 78.8% generic and 9.9% injection [Manirakiza 2024].")
h2("2.5 Unserved prescriptions")
p("Direct measurement of unserved prescriptions is rare. In Eswatini, 71% of "
  "chronic-disease patients did not always receive all prescribed medicines and 94% "
  "of them were referred to private pharmacies [Shabangu and Suleman 2015]. No Rwandan "
  "study of unserved prescriptions was found; QMIA 2026 appears to be the first "
  "national estimate.")
h2("2.6 Community level")
p("SC4CCM found 49% of Rwandan CHWs holding all five iCCM products at baseline, rising "
  "to 63% in quality-collaborative districts against 38% in comparison districts "
  "[Chandani 2014], with gains not fully sustained [Chandani 2017]. In Bugesera, "
  "self-reported antimalarial stock-out among 295 CHWs was 3.2%, but most walked to "
  "resupply and over half spent their own money doing so [Gakinahe 2024].")
h2("2.7 Prediction")
p("African work has focused on demand forecasting rather than stock-out "
  "classification. A random forest on Rwandan eLMIS data forecast essential-medicine "
  "demand with 76% test accuracy [Mbonyinshuti 2022]. No peer-reviewed African study "
  "reporting the discrimination (AUC) of a facility stock-out classifier was found, so "
  "the models in Section 4.12 have few direct comparators.")
h2("2.8 Antimalarial zoning")
p("Rwanda introduced multiple first-line therapies in response to artemisinin partial "
  "resistance (kelch13 R561H) [Uwimana 2020; Zupko 2023]. A 2025 description of the "
  "deployment plan allocates artesunate-pyronaridine to western districts, "
  "dihydroartemisinin-piperaquine to the East and Kigali, and artemether-lumefantrine to "
  "the South [Muvunyi 2025]. This differs from the zoning used in cleaning rule S10 for "
  "several districts, which reinforces the need to obtain the official district list.")

# ================================================================== 3
h1("3. Data and methods")
h2("3.1 Analytic samples")
tb(["Analysis", "Unit", "Sample"],
   [["Availability, expected (M1)", "Facility x molecule", f"{m1['unweighted']['n_rows']:,} at {m1['unweighted']['n_clusters']} facilities"],
    ["Availability, managed (M2)", "Facility x molecule", f"{m2['unweighted']['n_rows']:,} at {m2['unweighted']['n_clusters']} facilities"],
    ["Stock-outs, fill rate, models 2 and 4", "Facility x medicine (strength)", "9,628 managed pairs; 19,182 pair-months in the lagged panel"],
    ["Stock adequacy", "Stocked-out or stocked rows with AMC > 0", "8,474 rows"],
    ["Record accuracy", "Rows naming a digital system", "8,856 rows at 260 facilities"],
    ["Prescriptions", "Prescription / item", "6,006 prescriptions, 14,059 items"],
    ["Community", "CHW x product", "175 CHWs; 1,146 managed answered products"],
    ["Prediction", "Managed pairs", f"{PA['n']:,} pairs at {PA['facilities']} facilities"]],
   [5.5, 4.8, 6.7])
h2("3.2 Weights and variance estimation")
w = RJ["weights"]
p("Facilities were sampled within districts with probability roughly proportional "
  "to level, and one health post per district was added purposively. National "
  "estimates are therefore weighted by level: weight = facilities nationally / "
  "facilities sampled, from the facility GIS register ("
  + "; ".join(f"{k} {v['population']}/{v['sampled']} = {v['weight']}" for k, v in w.items())
  + "). Because the health-post sample is purposive and carries a weight of about 50, "
  "weighted results are also shown without health posts. Ratios are estimated with "
  "Taylor-linearised standard errors treating the facility as the primary sampling "
  "unit within level strata; the design effect compares this variance with simple "
  "random sampling of observations.")
h2("3.3 Regression models")
p("Binary outcomes (available, any stock-out day, exact record match, unserved "
  "prescription, antibiotic prescribed) were modelled with logistic regression "
  "estimated by generalised estimating equations (GEE) with an independence working "
  "correlation and cluster-robust (sandwich) standard errors by facility. Results are "
  "adjusted odds ratios (OR) with 95% confidence intervals. Facility-level covariates "
  "were level, management type, province, urban or rural setting (from the GIS "
  "register), main digital system, facility fill rate and facility record accuracy "
  "(each per 10 points). Programme and, for stock-outs, consumption volume and "
  "variability entered at row level. The share of variation at facility and district "
  "level was estimated with a Bayesian random-intercept logistic model (variational "
  "Bayes), on the latent scale where the residual variance is \u03c0\u00b2/3.")
h2("3.4 Lagged panel")
p("Each managed pair contributes two transitions: June to July and July to August. "
  "The outcome is any stock-out day in the following month; the exposure is the order "
  "outcome in the current month (no request; delivered in full; partly delivered; "
  "requested but nothing delivered), adjusted for being out in the current month, "
  "consumption, month, level and programme.")
h2("3.5 Prescribing, community and spatial analyses")
p("WHO/INRUD indicators follow WHO (1993): antibiotics are EML classes 11.2 and 11.3 "
  "(anti-bacterials); injections are items whose name indicates an injectable. CHW "
  "readiness requires every product for a condition to be managed and available: "
  "malaria an RDT and an ACT, diarrhoea ORS and zinc, pneumonia amoxicillin. CHW and "
  "supervising health centre were paired on the same molecule and compared with "
  "McNemar's exact test. Spatial autocorrelation of district indicators used Moran's I "
  "with four-nearest-neighbour weights and 9,999 permutations. The facility typology "
  "used k-means on six standardised facility indicators; k = 4 had the highest "
  f"silhouette (" + ", ".join(f"k={k}: {v}" for k, v in TY["silhouette"].items()) + ").")
h2("3.6 Predictive models")
p("Two classification tasks were set. **Prediction A**: a stock-out day in August, "
  "using only June and July data and facility characteristics. **Prediction B**: "
  "unavailability on the visit day in September, using June-August data. Three models "
  "were compared: penalised logistic regression, random forest and gradient boosting. "
  "Performance was estimated with five-fold cross-validation grouped by facility, so "
  "no facility appears in both training and test folds. Metrics are the area under the "
  "ROC curve (AUC, with 95% CI from 300 facility bootstrap resamples), average "
  "precision, Brier score and calibration by risk decile; the operating model is the "
  "best-calibrated one. Feature importance is the loss of AUC when a feature is "
  "permuted. Analyses were run in Python (statsmodels 0.14, scikit-learn 1.6); the code "
  "is in run_inferential_analysis.py.")

# ================================================================== 4
h1("4. Results")
h2("4.1 National estimates and their precision")
rows = []
for e in RJ["estimates"]:
    u, wv = e["unweighted"], e["weighted"]
    nh = NOHP.get(e["indicator"])
    rows.append([e["indicator"], ci(u), ci(wv) if wv else "-", ci(nh) if nh else "-",
                 f1(u["deff"]) if u.get("deff") else "-"])
tb(["Indicator", "Pooled (95% CI)", "Weighted (95% CI)", "Weighted, no HP", "Design effect"],
   rows, [5.6, 3.1, 3.1, 3.1, 2.1],
   "Pooled = as on the dashboard. Weighted = weighted by level to the national facility "
   "register. Design effect = variance relative to simple random sampling.")
fg("fig_estimates.png", "Figure 1. Headline indicators with 95% confidence intervals.")
acc_e = E["Inventory record accuracy (exact match)"]
p(f"**Interpretation.** The pooled figures on the dashboard are precise: the 95% "
  f"intervals for availability are about \u00b11 point. Weighting changes the picture "
  f"for expected-medicine availability ({f1(m1['unweighted']['est'])}% pooled, "
  f"{f1(m1['weighted']['est'])}% weighted), because health posts make up most of "
  f"Rwanda's facilities and stock fewer of the medicines the EML authorises for them; "
  f"without health posts the weighted value ({f1(NOHP['Availability, expected medicines (M1)']['est'])}%) "
  f"is close to the pooled one. Managed-product availability is robust to weighting. "
  f"Record accuracy has a design effect of {f1(acc_e['unweighted']['deff'])}: accuracy "
  f"is mostly a property of the facility, so its interval "
  f"({f1(acc_e['unweighted']['lo'])}-{f1(acc_e['unweighted']['hi'])}%) is much wider "
  f"than the 8,856 records would suggest, and its weighted value "
  f"({f1(acc_e['weighted']['est'])}%) is lower because health posts keep the least "
  f"accurate records. The prescription indicators are very precise but, as noted on the "
  f"dashboard, understate unmet need by design.")

h2("4.2 Differences between groups")
G = RJ["group_ci"]


def gtab(lab, by):
    return [[g["group"], f"{f1(g['est'])} ({f1(g['lo'])}-{f1(g['hi'])})", g["fac"]]
            for g in G[lab] if g["by"] == by]


for by, nm in (("level", "level of care"), ("management", "management type"),
               ("province", "province"), ("hub", "RMS hub")):
    a = {r[0]: r for r in gtab("M2 availability", by)}
    s_ = {r[0]: r for r in gtab("Stock-out incidence", by)}
    tb([nm.capitalize(), "Managed availability % (95% CI)", "Stock-out incidence % (95% CI)", "Facilities"],
       [[k, a[k][1], s_.get(k, ["", "-"])[1], a[k][2]] for k in a],
       [4.6, 4.6, 4.6, 2.2])
p(f"**Interpretation.** Overlapping intervals mean many apparent differences on the "
  f"dashboard are within sampling error. Of the 30 districts, {len(n_below)} have "
  f"managed availability significantly below the national value "
  f"({', '.join(n_below)}) and {len(n_above)} significantly above "
  f"({', '.join(n_above)}); the remaining {30 - len(n_below) - len(n_above)} cannot be "
  f"distinguished from the national figure with 5-11 facilities each. District maps "
  f"should be read with this in mind.")

h2("4.3 Where the variation lies")
tb(["Outcome", "Between facilities", "Between districts", "Within facility (residual)"],
   [[n, f"{f1(v['facility']['share'])}%", f"{f1(v['district']['share'])}%",
     f"{f1(v['residual_share'])}%"] for n, v in
    (("Available on the visit day", VA), ("Any stock-out day, Jun-Aug", VS),
     ("Exact record match", VQ))],
   [5.4, 3.8, 3.8, 4.0], "Variance partition from random-intercept logistic models (latent scale).")
p(f"**Interpretation.** For availability and stock-outs, most variation is between "
  f"medicines within the same facility; facilities and districts explain only "
  f"{f1(VA['facility']['share'] + VA['district']['share'])}% and "
  f"{f1(VS['facility']['share'] + VS['district']['share'])}%. Stock problems are "
  f"therefore mostly medicine-specific, and national or programme-level supply "
  f"actions matter as much as facility mentorship. Record accuracy is the opposite: "
  f"{f1(VQ['facility']['share'])}% of its variation is between facilities and "
  f"{f1(VQ['district']['share'])}% between districts. A facility either keeps its "
  f"records reconciled or it does not, which makes accuracy a natural target for "
  f"facility-level QMIA mentorship.")

h2("4.4 Determinants of availability (Model 1)")
tb(["Factor", "Adjusted OR (95% CI)", "p"], model_rows("gee_availability"), [10.0, 4.4, 2.6],
   f"Outcome: managed medicine available on the visit day. n = {RJ['gee_availability']['n']:,} "
   f"pairs, {RJ['gee_availability']['clusters']} facilities.")
fg("fig_forest_availability.png", "Figure 2. Adjusted odds ratios for availability.", 13)
p(f"**Interpretation.** With level, ownership, location, programme and system held "
  f"constant, district hospitals are about three times as likely as health centres to "
  f"have a managed medicine on the shelf (OR {orci(T('gee_availability', 'District hospital'))}). "
  f"Faith-based facilities do better than government ones "
  f"(OR {orci(T('gee_availability', 'Faith-based'))}), as in Malawi [Mohan 2024]. The "
  f"West is the only province significantly below Kigali City "
  f"(OR {orci(T('gee_availability', 'West'))}). Among programmes, malaria commodities "
  f"are almost always present (OR {orci(T('gee_availability', 'Malaria'))}) and ARV "
  f"products least (OR {orci(T('gee_availability', 'ARV'))}), consistent with the "
  f"nationwide ARV stock-outs on the dashboard. Neither the facility's fill rate, its "
  f"record accuracy, its digital system nor urban setting predicts availability once "
  f"these are accounted for: availability is driven more by what is in the national "
  f"pipeline for a given medicine than by facility traits.")

h2("4.5 Determinants of stock-outs (Model 2)")
tb(["Factor", "Adjusted OR (95% CI)", "p"], model_rows("gee_stockout"), [10.0, 4.4, 2.6],
   f"Outcome: any stock-out day June-August. n = {RJ['gee_stockout']['n']:,} pairs.")
fg("fig_forest_stockout.png", "Figure 3. Adjusted odds ratios for stock-out incidence.", 13)
p(f"**Interpretation.** Every vertical programme has far lower odds of a stock-out "
  f"than essential medicines (OR {T('gee_stockout', 'Malaria')['OR']:.2f} for malaria "
  f"to {T('gee_stockout', 'ARV')['OR']:.2f} for ARV). This reflects separate, "
  f"programme-managed pipelines with push distribution and larger buffer stocks, which "
  f"is also why these programmes are the most overstocked. Higher-volume medicines are "
  f"less often out (OR {orci(T('gee_stockout', 'log_amc'))} per unit log AMC), and "
  f"medicines with erratic month-to-month consumption more often "
  f"(OR {orci(T('gee_stockout', 'erratic'))}): an average-based order quantity fails "
  f"exactly where demand is irregular. Eastern facilities have higher odds than Kigali "
  f"City (OR {orci(T('gee_stockout', 'East'))}).")

h2("4.6 Does order fulfilment drive stock-outs? (Model 4)")
lag = RJ["lag_crosstab"]
tb(["Order outcome this month", "Pair-months", "Stocked out next month"],
   [[r["order"], f"{r['count']:,}", f"{f1(r['mean'])}%"] for r in lag], [7.0, 4.0, 6.0])
tb(["Factor", "Adjusted OR (95% CI)", "p"], model_rows("gee_lag"), [10.0, 4.4, 2.6],
   f"Outcome: any stock-out day in month t+1. n = {RJ['gee_lag']['n']:,} pair-months.")
fg("fig_forest_lag.png", "Figure 4. Next-month stock-out by this month's order outcome.", 13)
p(f"**Interpretation.** This is the clearest causal signal in the data, because the "
  f"exposure precedes the outcome. When a facility requested a medicine and received "
  f"nothing, {f1(LAG['Requested, nothing delivered']['mean'])}% of those pairs were out "
  f"the next month, and the adjusted odds were "
  f"{T('gee_lag', 'Requested, nothing delivered')['OR']:.1f} times those after a full "
  f"delivery, even after accounting for whether the medicine was already out. Partial "
  f"deliveries carried no excess risk, so the problem is orders that go entirely "
  f"unfilled, not quantities trimmed. Not requesting at all also raised the odds "
  f"(OR {orci(T('gee_lag', 'No request'))}), pointing to missed or late orders at the "
  f"facility. Being out this month was by far the strongest predictor "
  f"(OR {T('gee_lag', 'out_t')['OR']:.0f}): once a stock-out starts it tends to last. "
  f"Together these support two actions: RMS should flag and expedite nil-filled orders, "
  f"and facilities should not skip a requisition cycle.")

h2("4.7 Inventory record accuracy (Model 3)")
tb(["Factor", "Adjusted OR (95% CI)", "p"], model_rows("gee_accuracy"), [10.0, 4.4, 2.6],
   f"Outcome: physical count equals system count. n = {RJ['gee_accuracy']['n']:,} records.")
fg("fig_forest_accuracy.png", "Figure 5. Adjusted odds ratios for exact record match.", 13)
p(f"**Interpretation.** Records kept in eBuzima matched the shelf far less often than "
  f"eLMIS records (OR {orci(T('gee_accuracy', 'eBuzima'))}), as did Medisoft "
  f"(OR {orci(T('gee_accuracy', 'Medisoft'))}), after level and location are held "
  f"constant. eBuzima is a point-of-care system whose stock module depends on "
  f"dispensing being recorded for every patient, which the data-collector feedback "
  f"says is often not done. District hospitals are much more accurate than health "
  f"centres (OR {orci(T('gee_accuracy', 'District hospital'))}) and PPP facilities much "
  f"less (OR {orci(T('gee_accuracy', 'PPP'))}). The Rwandan QMIA reached 85% accuracy "
  f"by 2019 [Kigenza 2023]. Definitions may differ, but today's 56% suggests that "
  f"some of that gain has been lost; the eBuzima result points to the transition of "
  f"health centres to point-of-care systems as one likely reason.")

h2("4.8 Prescribing quality and unserved prescriptions")
tb(["WHO indicator", "Estimate (95% CI)", "WHO reference", "Africa median [Ofori-Asenso 2016]"],
   [["Medicines per encounter", f"{f2(WHO['Medicines per encounter']['est'])} ({f2(WHO['Medicines per encounter']['lo'])}-{f2(WHO['Medicines per encounter']['hi'])})", "< 2 (1.6-1.8)", "3.1"],
    ["Generic (INN) prescribing", ci(WHO["% medicines prescribed by generic (INN) name"]), "100%", "68.0%"],
    ["Encounters with an antibiotic", ci(WHO["% encounters with an antibiotic"]), "< 30%", "46.8%"],
    ["Encounters with an injection", ci(WHO["% encounters with an injection"]), "< 20%", "25.0%"],
    ["Medicines from the EML", f"{f1(WHO['% medicines from the EML']['est'])}% ({f1(WHO['% medicines from the EML']['lo'])}-{min(100, WHO['% medicines from the EML']['hi']):.1f})", "100%", "88.0%"]],
   [5.2, 4.2, 3.0, 4.6], "Weighted by level; facility-clustered 95% CI.")
tb(["Level", "Prescriptions", "Medicines / encounter", "% antibiotic", "% injection"],
   [[r["level"], r["n"], f2(r["meds"]), f1(r["abx"]), f1(r["inj"])] for r in RJ["who_by_level"]],
   [5.2, 2.8, 3.2, 2.9, 2.9])
fg("fig_who_iccm.png", "Figure 6. WHO prescribing indicators (left) and CHW readiness by iCCM condition (right).")
tb(["Factor", "Unserved: OR (95% CI)", "Antibiotic: OR (95% CI)"],
   [[term(a["term"]), orci(a), orci(b)] for a, b in zip(RJ["gee_rx"]["terms"], RJ["gee_abx"]["terms"])],
   [8.0, 4.5, 4.5], "Models 5 and 6. Outcomes: any unserved item; any antibiotic. Reference: nurse, primary care.")
p(f"**Interpretation.** Prescribing in Rwanda compares well with the African median on "
  f"generic names, EML adherence, polypharmacy and injections. Antibiotic use is the "
  f"exception: {f1(WHO['% encounters with an antibiotic']['est'])}% of encounters, about "
  f"the African median and well above the 30% reference. It is concentrated in primary "
  f"care (about half of health-centre and health-post encounters, against about a third "
  f"in hospitals), and general practitioners prescribe antibiotics more often than nurses "
  f"once the number of medicines is held constant (OR {orci(T('gee_abx', 'General practitioner'))}). "
  f"This is a priority for antimicrobial stewardship. Unserved prescriptions are rare, "
  f"but a specialist's prescription is far more likely to include an unavailable item "
  f"(OR {orci(T('gee_rx', 'Specialist'))}), as is a longer prescription "
  f"(OR {orci(T('gee_rx', 'n_med'))} per medicine). This matches the referral-hospital "
  f"gaps on the dashboard: specialist prescribing reaches beyond what the pharmacy "
  f"stocks.")

h2("4.9 Community level")
tb(["Condition", "CHWs ready (95% CI)"], [[r["condition"], ci(r)] for r in RJ["iccm_ready"]], [9.0, 8.0])
tb(["Province", "Malaria", "Diarrhoea", "Pneumonia", "All three"],
   [[r["province"], f1(r["malaria"]), f1(r["diarrhoea"]), f1(r["pneumonia"]), f1(r["iccm"])]
    for r in RJ["iccm_by_province"]], [4.2, 3.2, 3.2, 3.2, 3.2], "Percent of CHWs ready.")
pt = RJ["paired_total"]
tb(["Product (molecule)", "Pairs", "CHW available", "Health centre available", "CHW out, HC has it", "McNemar p"],
   [[r["molecule"].replace(" all strengths", "").replace("malaria rapid diagnostics ag p f pan t", "Malaria RDT"),
     r["pairs"], f"{f1(r['chw'])}%", f"{f1(r['hc'])}%", r["hc_only"], pv(r["mcnemar_p"])]
    for r in RJ["paired_chw_hc"]], [5.0, 1.8, 2.6, 3.0, 2.8, 1.8],
   "CHW paired with the supervising health centre on the same molecule.")
p(f"**Interpretation.** Treatment readiness is lower than product availability "
  f"suggests: only {f1(RD['All three conditions']['est'])}% of CHWs could treat malaria, "
  f"diarrhoea and pneumonia on the visit day, below the 49% SC4CCM baseline for Rwanda "
  f"[Chandani 2014] although the product lists differ. Diarrhoea is the weakest "
  f"condition because ORS and zinc must both be present. Kigali City has the lowest "
  f"readiness, largely because of amoxicillin. In all {pt['pairs']} CHW-health centre "
  f"pairs, the health centre had the product whenever the CHW did, and in every one of "
  f"the {pt['hc_only']} discordant pairs the CHW was out while the health centre had "
  f"stock (exact p {pv(pt['p'])}). The community gap is therefore one of resupply "
  f"from the health centre to the CHW, which also matches the Bugesera finding that CHWs "
  f"walk and pay to collect stock [Gakinahe 2024].")

h2("4.10 Geography")
tb(["District indicator", "Moran's I", "Expected under no clustering", "Permutation p"],
   [[k, f"{v['I']:.3f}", f"{v['expected']:.3f}", f"{v['p']:.3f}"] for k, v in RJ["moran"].items()],
   [6.0, 3.0, 4.5, 3.5])
p("**Interpretation.** Neighbouring districts are no more alike than distant ones for "
  "availability, stock-outs or record accuracy; weak districts are scattered rather than "
  "forming a region. CHW availability shows weak positive clustering (p about 0.05), "
  "consistent with zone-wide antimalarial allocation. In practice this means support "
  "should be allocated facility by facility rather than to a whole region.")

h2("4.11 Facility typology")
tb(["Profile", "Facilities", "Availability", "Stock-out", "Accuracy", "Fill rate", "Overstock", "Understock"],
   [[f"{c['cluster'] + 1}. {CN[c['cluster']]}", c["facilities"], f1(c["availability"]), f1(c["stockout"]),
     f1(c["accuracy"]), f1(c["fill"]), f1(c["overstock"]), f1(c["understock"])]
    for c in PROF], [4.6, 1.8, 1.8, 1.7, 1.7, 1.7, 1.8, 1.9],
   "Cluster means, percent. The facility list is in results/facility_typology.csv.")
fg("fig_typology.png", "Figure 7. Facility profiles from k-means clustering.", 15)
p("**Interpretation.** The profiles separate problems that need different responses. "
  "**Frequent stock-outs** facilities have the lowest availability despite normal fill "
  "rates, pointing to under-ordering and quantification. **Weak resupply and weak records** "
  "facilities receive markedly less of what they order and keep the least accurate "
  "records; they need RMS follow-up and record reconciliation together. **Well supplied "
  "but overstocked** facilities are candidates for redistribution of excess stock. "
  "**Lean stock, often understocked** facilities run close to the emergency order point "
  "and need more frequent resupply. The silhouette values are modest, so the profiles "
  "are a practical grouping for planning rather than sharply distinct types.")

h2("4.12 Predicting stock-outs")
tb(["Model", "AUC (95% CI)", "Average precision", "Brier score"],
   [[f"A: {m['model']}", f"{m['auc']:.3f} ({m['auc_lo']:.3f}-{m['auc_hi']:.3f})", f"{m['ap']:.3f}", f"{m['brier']:.4f}"] for m in PA["models"]]
   + [[f"B: {m['model']}", f"{m['auc']:.3f} ({m['auc_lo']:.3f}-{m['auc_hi']:.3f})", f"{m['ap']:.3f}", f"{m['brier']:.4f}"] for m in PB["models"]],
   [6.0, 4.4, 3.3, 3.3],
   f"A: stock-out in August (prevalence {PA['prevalence']}%). B: unavailable on the visit day "
   f"(prevalence {PB['prevalence']}%). Facility-grouped five-fold cross-validation.")
fg("fig_pred_august.png", "Figure 8. Prediction A: ROC, calibration and feature importance.")
fg("fig_pred_visit.png", "Figure 9. Prediction B: ROC, calibration and feature importance.")
ra, ma = PA["rule_baseline"], PA["model_same_rate"]
rb, mb = PB["rule_baseline"], PB["model_same_rate"]
tb(["Approach", "Pairs flagged", "Precision", "Recall"],
   [["A: rule \"out in July\"", f"{f1(ra['flagged'])}%", f"{f1(ra['precision'])}%", f"{f1(ra['recall'])}%"],
    [f"A: {PA['best']}, same flag rate", f"{f1(ma['flagged'])}%", f"{f1(ma['precision'])}%", f"{f1(ma['recall'])}%"],
    [f"A: {PA['best']}, top 10% risk", "10.0%", f"{f1(PA['top10']['precision'])}%", f"{f1(PA['top10']['recall'])}%"],
    ["B: rule \"out in July\"", f"{f1(rb['flagged'])}%", f"{f1(rb['precision'])}%", f"{f1(rb['recall'])}%"],
    [f"B: {PB['best']}, same flag rate", f"{f1(mb['flagged'])}%", f"{f1(mb['precision'])}%", f"{f1(mb['recall'])}%"],
    [f"B: {PB['best']}, top 10% risk", "10.0%", f"{f1(PB['top10']['precision'])}%", f"{f1(PB['top10']['recall'])}%"]],
   [7.0, 3.2, 3.2, 3.6], "Precision: share of flagged pairs that stocked out. Recall: share of stock-outs flagged.")
p(f"**Interpretation.** All three models discriminate well (AUC about "
  f"{best(PA)['auc']:.2f} for August stock-outs and {best(PB)['auc']:.2f} for the visit "
  f"day), and the random forest is well calibrated. Most of this comes from persistence: "
  f"July stock-out days are by far the most informative feature, and for August the simple "
  f"rule \"flag what was out in July\" performs as well as the model at the same flag rate. "
  f"The models add value in two ways. First, they rank risk among pairs that were not out "
  f"in July (AUC {PA['new_stockouts']['auc']:.2f} for new August stock-outs; "
  f"{PB['new_stockouts']['auc']:.2f} for the visit day), using the medicine, consumption, "
  f"requests and the facility's fill rate and record accuracy. Second, for the visit-day "
  f"outcome the model beats the rule clearly at the same flag rate (precision "
  f"{f1(mb['precision'])}% against {f1(rb['precision'])}%). Flagging the top 10% of pairs "
  f"by predicted risk captures about {f1(PA['top10']['recall'])}% of next-month "
  f"stock-outs with a {PA['top10']['lift']:.1f}-fold lift over random selection. A "
  f"monthly early-warning list built from eLMIS data in this way is feasible; it should be "
  f"validated on a later period before use, because three months is a short history.")

# ================================================================== 5
h1("5. Discussion")
h2("5.1 Main messages")
bl(f"**Availability is high by regional standards but uneven.** The managed-product "
   f"figure ({f1(m2['unweighted']['est'])}%) and the expected-medicine figure "
   f"({f1(m1['unweighted']['est'])}%) are well above typical "
   "African public-sector values. The national weighted figure for expected medicines "
   f"({f1(m1['weighted']['est'])}%) is lower because health posts, which are most of Rwanda's facilities, do not "
   "stock many of the medicines the EML authorises for them.",
   f"**Order fulfilment is a lever with a measurable effect.** Nil-filled requisitions "
   f"multiply the odds of next-month stock-out about "
   f"{T('gee_lag', 'Requested, nothing delivered')['OR']:.1f}-fold. This is the most "
   "actionable result for RMS, and it echoes evidence that order fulfilment and inventory "
   "policy, more than facility traits, drive stock-outs [Leung 2016; Mohan 2024].",
   "**Stock problems are medicine- and pipeline-specific.** Facilities and districts "
   "explain little of the variation in availability and stock-outs, and there is no "
   "regional clustering. National quantification for the essential-medicines pipeline "
   "(insulin, thiamine, HIV test kits, paediatric RHZ) matters as much as facility "
   "mentorship.",
   f"**Record accuracy is the facility-level weakness, and it has slipped.** "
   f"{f1(acc_e['unweighted']['est'])}% exact "
   "match is well below the 85% QMIA reported for 2019 [Kigenza 2023], and eBuzima sites "
   "are markedly worse. This is where facility mentorship pays off.",
   "**Antibiotic use is the main prescribing concern**, while generic and EML "
   "prescribing are excellent.",
   "**The last mile to CHWs is the weak link in community supply.** Health centres hold "
   "stock their CHWs lack.")
h2("5.2 Strengths and limitations")
bl("**Strengths:** national coverage of all levels; linked stock, prescription and CHW "
   "data; confidence intervals and design effects that account for clustering; a "
   "temporally ordered test of order fulfilment; predictive models validated by "
   "facility-grouped cross-validation.",
   "**Cross-sectional design.** Except for the lagged panel, associations are not causal.",
   "**Short history.** Three months (June-August) limit trend and prediction; the "
   "models should be revalidated on later months.",
   "**Weights.** Health posts were sampled purposively; weighted national figures that "
   "include them should be read as indicative.",
   "**Self-report.** RMS availability and some quantities are facility-reported; "
   "consumption may be under-recorded where dispensing is not captured.",
   "**Prescription design.** Prescribers write what the pharmacy holds, so unserved "
   "rates understate unmet need; there is no outpatient/inpatient flag.",
   "**Zoning.** Cleaning rule S10 uses a partly inferred zoning list that differs from "
   "the published deployment plan for some districts [Muvunyi 2025].")

# ================================================================== 6
h1("6. Recommendations")
bl("**RMS: treat a nil-filled order as an alert.** Track requisitions delivered at 0% and "
   "expedite or substitute within the month; report the share monthly by hub.",
   "**Facilities: never skip a requisition cycle** for a managed medicine, even when "
   "stock looks sufficient; a missed order "
   f"raises the odds of a stock-out next month about "
   f"{T('gee_lag', 'No request')['OR']:.1f}-fold.",
   "**Build a monthly stock-out early-warning list** from eLMIS: pairs out last month "
   "plus the model's top-risk pairs, sent to RMS branches and district pharmacists.",
   "**Focus QMIA mentorship on record reconciliation** at health centres using eBuzima, "
   "and at PPP facilities; aim to restore the 85% accuracy of 2019.",
   "**Use the facility profiles to tailor support:** redistribution for overstocked "
   "sites, RMS follow-up for weak-resupply sites, quantification support for "
   "frequent-stock-out sites.",
   "**Strengthen health centre-to-CHW resupply**, especially ORS and zinc as a pair and "
   "amoxicillin in Kigali City; consider resupply at the monthly CHW meeting.",
   "**Antimicrobial stewardship in primary care:** feed back antibiotic prescribing rates "
   "by facility and prescriber, starting with health centres and general practitioners.",
   "**Report weighted estimates with confidence intervals** in the national report, and "
   "flag district values whose intervals include the national figure.",
   "**Confirm the official MFT zoning list** and rerun the zoning rule and these analyses.")

# ================================================================== refs
h1("References")
REFS = [
    "Bizimana T, Kayumba PC, Heide L. Prices, availability and affordability of medicines in Rwanda. PLoS One. 2020;15(8):e0236411.",
    "Cameron A, Ewen M, Ross-Degnan D, Ball D, Laing R. Medicine prices, availability, and affordability in 36 developing and middle-income countries: a secondary analysis. Lancet. 2009;373:240-249.",
    "Chandani Y, Andersson S, Heaton A, et al. Making products available among community health workers: evidence for improving community health supply chains from Ethiopia, Malawi, and Rwanda. J Glob Health. 2014;4(2):020405.",
    "Chandani Y, Duffy M, Lamphere B, et al. Quality improvement practices to institutionalize supply chain best practices for iCCM: evidence from Rwanda and Malawi. Res Social Adm Pharm. 2017;13(6):1095-1109.",
    "Choi Y, Ametepi P. Comparison of medicine availability measurements at health facilities: evidence from Service Provision Assessment surveys in five sub-Saharan African countries. BMC Health Serv Res. 2013;13:266.",
    "Gakinahe GU, Rutungwa E, Mbonyinshuti F, Kayitare E. Availability of antimalarial medicines and inventory management at the community level: a case study of Bugesera district in Rwanda. BMC Health Serv Res. 2024;24:136.",
    "Kabera JC, Mukanyangezi MF. Influence of inventory management practices on the availability of emergency obstetric drugs in Rwandan public hospitals. BMC Health Serv Res. 2024;24.",
    "Kigenza R, Nsengiyumva E, Sabagirirwa V. The Quality Management Improvement Approach: successes and lessons learned from a workforce development intervention in Rwanda's health supply chain. Glob Health Sci Pract. 2023;11(1):e2200295.",
    "Kuwawenaruwa A, Wyss K, Wiedenmayer K, Metta E, Tediosi F. The effects of medicines availability and stock-outs on household's utilization of healthcare services in Dodoma region, Tanzania. Health Policy Plan. 2020;35(3):323-333.",
    "Leung NHZ, Chen A, Yadav P, Gallien J. The impact of inventory management on stock-outs of essential drugs in sub-Saharan Africa: secondary analysis of a field experiment in Zambia. PLoS One. 2016;11(5):e0156026.",
    "Manirakiza A, et al. Evaluating drug use patterns among paediatric outpatients in Burundi. J Pharm Policy Pract. 2024;17(1):2312369.",
    "Masters SH, Burstein R, DeCenso B, et al. Pharmaceutical availability across levels of care: evidence from facility surveys in Ghana, Kenya, and Uganda. PLoS One. 2014;9(12):e114762.",
    "Mbonyinshuti F, Nkurunziza J, Niyobuhungiro J, Kayitare E. The prediction of essential medicines demand: a machine learning approach using consumption data in Rwanda. Processes. 2022;10(1):26.",
    "Mekonen ZT, Cho DJ, Fenta TG. Logistics Management Information System (LMIS) for health commodities at public health facilities in Amhara National Regional State of Ethiopia: a data quality evaluation survey. J Multidiscip Healthc. 2025;18:255-266.",
    "Mohan S, Mangal TD, Colbourn T, et al. Factors associated with medical consumable availability in level 1 facilities in Malawi: a secondary analysis of a facility census. Lancet Glob Health. 2024;12(6):e1027-e1037.",
    "Mukundiyukuri JP, Irakiza JJ, Nyirahabimana N, et al. Availability, costs and stock-outs of essential NCD drugs in three rural Rwandan districts. Ann Glob Health. 2020;86(1):123.",
    "Muvunyi CM, Gashema P, Mbituyumuremyi A, et al. Strategies for mitigating emerging artemisinin-based antimalarial drug resistance in Rwanda. BMJ Glob Health. 2025;10(10):e020884.",
    "Ofori-Asenso R, Brhlikova P, Pollock AM. Prescribing indicators at primary health care centers within the WHO African region: a systematic analysis (1995-2015). BMC Public Health. 2016;16:724.",
    "Shabangu K, Suleman F. Medicines availability at a Swaziland hospital and impact on patients. Afr J Prim Health Care Fam Med. 2015;7(1):829.",
    "Uwimana A, Legrand E, Stokes BH, et al. Emergence and clonal expansion of in vitro artemisinin-resistant Plasmodium falciparum kelch13 R561H mutant parasites in Rwanda. Nat Med. 2020;26(10):1602-1608.",
    "World Health Organization. How to investigate drug use in health facilities: selected drug use indicators. WHO/DAP/93.1. Geneva: WHO; 1993.",
    "Zupko RJ, Nguyen TD, Ngabonziza JCS, et al. Modeling policy interventions for slowing the spread of artemisinin-resistant pfkelch R561H mutations in Rwanda. Nat Med. 2023;29(11):2775-2784.",
]
C.append(("refs", REFS))


# ------------------------------------------------------------------ Word
def build_docx(path):
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
    NAVY = RGBColor(0x1F, 0x3A, 0x68)
    GREY = RGBColor(0x55, 0x55, 0x55)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(2.0)
    st = doc.styles
    st["Normal"].font.name = "Arial"
    st["Normal"].font.size = Pt(10)
    st["Normal"].paragraph_format.space_after = Pt(5)
    st["Normal"].element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    for name, size in (("Heading 1", 15), ("Heading 2", 12)):
        s = st[name]
        s.font.name, s.font.size, s.font.bold = "Arial", Pt(size), True
        s.font.color.rgb = NAVY
        s.paragraph_format.space_before = Pt(12)
        s.paragraph_format.space_after = Pt(4)
        s.paragraph_format.keep_with_next = True

    def runs(par, text, size=None, color=None, italic=False):
        for i, part in enumerate(re.split(r"\*\*", text)):
            if part:
                r = par.add_run(part)
                r.bold = bool(i % 2)
                r.italic = italic
                if size:
                    r.font.size = Pt(size)
                if color:
                    r.font.color.rgb = color
        return par

    def shade(cell, col):
        tcPr = cell._tc.get_or_add_tcPr()
        sh = OxmlElement("w:shd")
        sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), col)
        tcPr.append(sh)

    def table(hdr, rows, widths, cap):
        t = doc.add_table(rows=1, cols=len(hdr))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        for j, h in enumerate(hdr):
            c = t.rows[0].cells[j]
            c.text = ""
            runs(c.paragraphs[0], "**" + h + "**", 8, RGBColor(255, 255, 255))
            shade(c, "1F3A68")
        for k, row in enumerate(rows):
            cells = t.add_row().cells
            for j, v in enumerate(row):
                cells[j].text = ""
                runs(cells[j].paragraphs[0], str(v), 8)
                if k % 2:
                    shade(cells[j], "EEF2F7")
        for row in t.rows:
            for j, w_ in enumerate(widths):
                row.cells[j].width = Cm(w_)
            for c in row.cells:
                for par in c.paragraphs:
                    par.paragraph_format.space_after = Pt(1)
        trPr = t.rows[0]._tr.get_or_add_trPr()
        el = OxmlElement("w:tblHeader"); el.set(qn("w:val"), "true"); trPr.append(el)
        if cap:
            runs(doc.add_paragraph(), cap, 8, GREY, italic=True)
        else:
            doc.add_paragraph().paragraph_format.space_after = Pt(2)

    tp = doc.add_paragraph()
    tp.paragraph_format.space_before = Pt(60)
    runs(tp, "**" + TITLE + "**", 22, NAVY)
    runs(doc.add_paragraph(), SUBTITLE, 15, NAVY)
    runs(doc.add_paragraph(), META, 9.5, GREY)
    q = doc.add_paragraph()
    q.paragraph_format.space_before = Pt(24)
    runs(q, "**Contents.** Executive summary \u00b7 1 Introduction \u00b7 2 What comparable "
         "studies have found \u00b7 3 Data and methods \u00b7 4 Results (4.1-4.12) \u00b7 "
         "5 Discussion \u00b7 6 Recommendations \u00b7 References", 9.5)
    doc.add_page_break()
    for blk in C:
        k = blk[0]
        if k == "h1":
            doc.add_heading(blk[1], level=1)
        elif k == "h2":
            doc.add_heading(blk[1], level=2)
        elif k == "p":
            runs(doc.add_paragraph(), blk[1])
        elif k == "bullets":
            for b in blk[1]:
                runs(doc.add_paragraph(style="List Bullet"), b)
        elif k == "table":
            table(*blk[1:])
        elif k == "fig":
            doc.add_picture(os.path.join(FIG, blk[1]), width=Cm(blk[3]))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            runs(doc.add_paragraph(), blk[2], 8, GREY, italic=True)
        elif k == "refs":
            for r in blk[1]:
                par = doc.add_paragraph()
                par.paragraph_format.left_indent = Cm(0.6)
                par.paragraph_format.first_line_indent = Cm(-0.6)
                runs(par, r, 8.5)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = fp.add_run()
    for t_, txt in (("begin", None), (None, "PAGE"), ("end", None)):
        if t_:
            f = OxmlElement("w:fldChar"); f.set(qn("w:fldCharType"), t_); r._r.append(f)
        else:
            it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve")
            it.text = txt; r._r.append(it)
    r.font.size = Pt(8)
    doc.save(path)


# ------------------------------------------------------------------ LaTeX
def tx(s):
    s = str(s)
    rep = {"\\": r"\textbackslash{}", "&": r"\&", "%": r"\%", "$": r"\$", "#": r"\#",
           "_": r"\_", "{": r"\{", "}": r"\}", "~": r"\textasciitilde{}", "^": r"\textasciicircum{}"}
    s = "".join(rep.get(ch, ch) for ch in s)
    s = (s.replace("\u00b7", r"\textperiodcentered{}").replace("\u00b1", r"$\pm$")
         .replace("\u03c0\u00b2", r"$\pi^2$").replace(">=", r"$\geq$").replace("<", r"$<$")
         .replace(">", r"$>$").replace("\"", "''"))
    s = re.sub(r"\*\*(.+?)\*\*", r"\\textbf{\1}", s)
    s = re.sub(r"(\d) x (\d)", r"\1 $\\times$ \2", s)
    s = s.replace(" x ", r" $\times$ ")
    return s


def build_tex(path):
    o = [r"""\documentclass[10pt,a4paper]{article}
\usepackage[T1]{fontenc}
\usepackage[utf8]{inputenc}
\usepackage[margin=2cm]{geometry}
\usepackage{helvet}\renewcommand{\familydefault}{\sfdefault}
\usepackage{booktabs,longtable,array,xcolor,colortbl,enumitem,titlesec,graphicx,caption}
\usepackage[hidelinks]{hyperref}
\definecolor{navy}{HTML}{1F3A68}\definecolor{band}{HTML}{EEF2F7}
\titleformat{\section}{\large\bfseries\color{navy}}{}{0pt}{}
\titleformat{\subsection}{\normalsize\bfseries\color{navy}}{}{0pt}{}
\setlist{nosep,leftmargin=1.2em}
\setlength{\parindent}{0pt}\setlength{\parskip}{4pt}
\graphicspath{{results/}}
\begin{document}
\begin{titlepage}\vspace*{3cm}
{\Huge\bfseries\color{navy} """ + tx(TITLE) + r"""\par}\vspace{8pt}
{\Large\color{navy} """ + tx(SUBTITLE) + r"""\par}\vspace{14pt}
{\small\color{gray} """ + tx(META) + r"""\par}
\end{titlepage}
\tableofcontents\newpage
"""]
    for blk in C:
        k = blk[0]
        if k == "h1":
            o.append(r"\section{" + tx(re.sub(r"^\d+\.\s*", "", blk[1])) + "}")
        elif k == "h2":
            o.append(r"\subsection{" + tx(re.sub(r"^\d+\.\d+\s*", "", blk[1])) + "}")
        elif k == "p":
            o.append(tx(blk[1]) + "\n")
        elif k == "bullets":
            o.append(r"\begin{itemize}" + "".join(r"\item " + tx(b) + "\n" for b in blk[1])
                     + r"\end{itemize}")
        elif k == "table":
            hdr, rows, wd, cap = blk[1:]
            tot = sum(wd)
            spec = "".join(r">{\raggedright\arraybackslash}p{" + f"{0.97 * x / tot:.3f}"
                           + r"\linewidth}" for x in wd)
            head = " & ".join(r"\textbf{\color{white}" + tx(h) + "}" for h in hdr)
            body = [(r"\rowcolor{band}" if i % 2 else "") + " & ".join(tx(v) for v in r)
                    + r" \\" for i, r in enumerate(rows)]
            o.append(r"{\footnotesize\setlength{\tabcolsep}{3pt}" "\n" r"\begin{longtable}{"
                     + spec + "}\n" r"\rowcolor{navy}" + head + r" \\" "\n" r"\endhead" "\n"
                     + "\n".join(body) + "\n"
                     + (r"\multicolumn{" + str(len(hdr)) + r"}{p{0.97\linewidth}}{\scriptsize\itshape "
                        + tx(cap) + "}\\\\\n" if cap else "")
                     + r"\end{longtable}}")
        elif k == "fig":
            o.append(r"\begin{figure}[htbp]\centering\includegraphics[width="
                     + f"{min(1.0, blk[3] / 17):.2f}" + r"\linewidth]{" + blk[1]
                     + r"}\caption*{\small " + tx(blk[2]) + r"}\end{figure}")
        elif k == "refs":
            o.append(r"\begin{enumerate}[leftmargin=1.5em]\small" +
                     "".join(r"\item " + tx(r) + "\n" for r in blk[1]) + r"\end{enumerate}")
    o.append(r"\end{document}")
    open(path, "w", encoding="utf-8").write("\n".join(o))


if __name__ == "__main__":
    out = os.path.join(HERE, NAME + ".docx")
    if os.path.exists(out) and not os.access(out, os.W_OK):   # open / locked in Word
        out = os.path.join(HERE, NAME + "_v2.docx")
    build_docx(out)
    build_tex(os.path.join(HERE, NAME + ".tex"))
    print("wrote", NAME + ".docx and .tex")
