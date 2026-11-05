"""Inferential and predictive analysis of the QMIA 2026 assessment.

    python3 extract_analysis_data.py      (once, after any data refresh)
    python3 run_inferential_analysis.py   -> results/results.json, results/fig_*.png

Sections
  1  Design-based estimates (stratified, clustered, level-weighted) with 95% CI
  2  Variance partition (facility / district) - random-intercept logistic
  3  GEE logistic models: availability, stock-out incidence, record accuracy
  4  Lagged panel: does a partial or nil delivery predict next month's stock-out
  5  Prescriptions: WHO/INRUD indicators with CI; model of unserved prescriptions
  6  Community level: iCCM readiness, paired CHW vs supervising health centre
  7  Spatial autocorrelation (Moran's I) of district indicators
  8  Facility typology (k-means)
  9  Predictive models: August stock-out from June-July data; visit-day
     unavailability from June-August data (grouped cross-validation)
"""
import json
import math
import os
import re
import warnings

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import statsmodels.api as sm
import statsmodels.formula.api as smf
from scipy import stats

warnings.filterwarnings("ignore")
HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(HERE, "data")
OUT = os.path.join(HERE, "results")
os.makedirs(OUT, exist_ok=True)
ROOT = os.path.dirname(HERE)
CLEAN = os.path.join(ROOT, "Data cleaning", "Cleaned dataset 2026-11-05")
RES = {}
RNG = np.random.default_rng(2026)
NAVY, TEAL, GOLD, MAROON, GREY = "#1F3A68", "#117996", "#F3B71B", "#7C1220", "#9AA5B1"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9,
                     "axes.spines.top": False, "axes.spines.right": False})

# ------------------------------------------------------------------ data
mg = pd.read_csv(os.path.join(D, "stock_managed_rows.csv"))
avm = pd.read_csv(os.path.join(D, "availability_managed.csv"))
ave = pd.read_csv(os.path.join(D, "availability_expected.csv"))
chw = pd.read_csv(os.path.join(D, "chw_rows.csv"))
cen = pd.read_csv(os.path.join(D, "district_centroids.csv"))

# population counts by level (Facility GIS_.xlsx; NRH = national teaching
# hospitals; L2TH sampled exhaustively so its weight is 1)
POP = {"Health post": 1253, "Health centre": 516, "District hospital": 37,
       "Level 2 teaching hospital": 9, "National referral hospital": 5}
nfac = mg.groupby("level").facility.nunique().to_dict()
W = {l: max(1.0, POP[l] / nfac[l]) for l in nfac}
RES["weights"] = {l: {"population": POP[l], "sampled": int(nfac[l]),
                      "weight": round(W[l], 2)} for l in nfac}

# urban / rural from the GIS file, by normalised name
def norm(n):
    n = str(n).lower()
    n = re.sub(r"\(.*?\)", " ", n)
    n = re.sub(r"\b(cs|hc|hp|dh|mhc|ph|health|centre|center|post|hospital|"
               r"district|referral|provincial)\b", " ", n)
    return re.sub(r"[^a-z]", "", n)


gis = pd.read_excel(os.path.join(ROOT, "Facility GIS_.xlsx"), sheet_name=None)
setting = {}
for sh, ncol, scol in (("HC Referral", "HF", "Settings"),
                       ("DH Referrals", "HF", "Setting"),
                       ("NRH", "HF", "Setting")):
    g = gis[sh]
    for a, b in zip(g[ncol], g[scol]):
        if isinstance(b, str):
            setting[norm(a)] = b.strip().lower()
facs = mg[["facility", "district", "province", "hub", "level", "management"]] \
    .drop_duplicates("facility").set_index("facility")
facs["setting"] = [setting.get(norm(f), "unknown") for f in facs.index]
facs.loc[facs.level == "Health post", "setting"] = "rural (HP)"
RES["setting_match"] = facs.setting.value_counts().to_dict()


# ------------------------------------------------------------------ helpers
def ratio_est(df, y, x, cluster="facility", strata="level", weighted=True):
    """Ratio sum(y)/sum(x) with Taylor-linearised SE, clusters within strata."""
    g = df.groupby(cluster).agg(Y=(y, "sum"), X=(x, "sum"),
                                h=(strata, "first")).reset_index()
    g["w"] = g.h.map(W).fillna(1.0) if weighted else 1.0
    R = (g.w * g.Y).sum() / (g.w * g.X).sum()
    g["z"] = g.w * (g.Y - R * g.X) / (g.w * g.X).sum()
    var = 0.0
    for _, gh in g.groupby("h"):
        n = len(gh)
        if n > 1:
            var += n / (n - 1) * ((gh.z - gh.z.mean()) ** 2).sum()
    se = math.sqrt(var)
    # design effect against simple random sampling of the rows
    p = R
    n_rows = df[x].sum()
    srs = math.sqrt(max(p * (1 - p), 1e-12) / max(n_rows, 1)) if df[x].max() <= 1 else None
    return dict(est=R, lo=R - 1.96 * se, hi=R + 1.96 * se, se=se,
                deff=(se / srs) ** 2 if srs else None,
                n_rows=int(n_rows), n_clusters=int(len(g)))


def pct(d, k=1):
    return {kk: (round(100 * v, k) if isinstance(v, float) and kk in
                 ("est", "lo", "hi") else v) for kk, v in d.items()}


def gee_logit(formula, data, group, name, ref_note=""):
    m = smf.gee(formula, group, data, family=sm.families.Binomial(),
                cov_struct=sm.cov_struct.Independence())
    r = m.fit()
    ci = r.conf_int()
    rows = []
    for k in r.params.index:
        if k == "Intercept":
            continue
        rows.append(dict(term=k, OR=float(np.exp(r.params[k])),
                         lo=float(np.exp(ci.loc[k, 0])),
                         hi=float(np.exp(ci.loc[k, 1])),
                         p=float(r.pvalues[k])))
    rho = None
    RES[name] = dict(n=int(r.nobs), clusters=int(data[group].nunique()),
                     exch_corr=rho, terms=rows, formula=formula, note=ref_note)
    return r, rows


LABEL = {"level2": "Level", "mgmt": "Management", "province": "Province",
         "programme": "Programme", "urban": "Setting", "fsys": "Main system",
         "sys": "System", "order": "Order this month", "prescriber": "Prescriber",
         "month": "Month", "out_t": "Stocked out this month", "log_amc":
         "log(AMC + 1)", "erratic": "Erratic consumption (CV >= 50%)",
         "fill10": "Facility fill rate (+10 pts)", "acc10": "Facility record accuracy (+10 pts)",
         "n_med": "Medicines on the prescription (+1)", "jul": "August (vs July)"}
FEAT = {"so_jul": "Stock-out days, July", "so_jun": "Stock-out days, June",
        "so_aug": "Stock-out days, August", "molecule": "Medicine",
        "fill10": "Facility fill rate", "acc10": "Facility record accuracy",
        "lcons_jul": "Consumption, July", "lcons_jun": "Consumption, June",
        "lcons_aug": "Consumption, August", "lreq_jul": "Requested, July",
        "lreq_jun": "Requested, June", "lreq_aug": "Requested, August",
        "ldel_jul": "Delivered, July", "ldel_jun": "Delivered, June",
        "ldel_aug": "Delivered, August", "fr_jul": "Fill ratio, July",
        "fr_jun": "Fill ratio, June", "fr_aug": "Fill ratio, August",
        "level2": "Level", "province": "Province", "mgmt": "Management",
        "programme": "Programme", "urban": "Setting", "fsys": "Main system",
        "lprice": "Tariff price"}


def nice(term):
    t = term
    t = re.sub(r"C\((\w+), Treatment\(reference=['\"](.+?)['\"]\)\)\[T\.(.+?)\]",
               r"\1: \3 (vs \2)", t)
    t = re.sub(r"C\((\w+)\)\[T\.(.+?)\]", r"\1: \2", t)
    for k, v in LABEL.items():
        t = re.sub(rf"^{k}(?=:|$)", v, t)
    t = t.replace("Month: jul", "Month: August (vs July)")
    return t.replace("_", " ")


def forest(rows, title, fn, xlab="Adjusted odds ratio (95% CI, log scale)"):
    rows = [r for r in rows if np.isfinite(r["lo"]) and np.isfinite(r["hi"])]
    fig, ax = plt.subplots(figsize=(7.2, 0.32 * len(rows) + 1.2))
    y = np.arange(len(rows))[::-1]
    for yi, r in zip(y, rows):
        col = MAROON if r["hi"] < 1 else (TEAL if r["lo"] > 1 else GREY)
        ax.plot([r["lo"], r["hi"]], [yi, yi], color=col, lw=2)
        ax.plot(r["OR"], yi, "o", color=col, ms=5)
    ax.axvline(1, color="#555", lw=0.8, ls="--")
    ax.set_xscale("log")
    ax.set_yticks(y)
    ax.set_yticklabels([nice(r["term"]) for r in rows], fontsize=8)
    ax.set_xlabel(xlab)
    fig.suptitle(title, fontsize=10, x=0.01, ha="left", color=NAVY, fontweight="bold")
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, fn), dpi=200)
    plt.close(fig)


# ------------------------------------------------------------------ 1 estimates
mg["one"] = 1
mg["anyso"] = ((mg.so_jun > 0) | (mg.so_jul > 0) | (mg.so_aug > 0)).astype(int)
mg["rq"] = sum(np.where(mg[f"req_{m}"] >= mg[f"del_{m}"], mg[f"req_{m}"], 0)
               for m in ("jun", "jul", "aug"))
mg["dv"] = sum(np.where(mg[f"req_{m}"] >= mg[f"del_{m}"], mg[f"del_{m}"], 0)
               for m in ("jun", "jul", "aug"))
ad = mg[mg.adequacy.notna() & (mg.adequacy != "")].copy()
for k in ("out", "under", "adequate", "over"):
    ad[k] = (ad.adequacy == k).astype(int)
acc = mg[mg.in_accuracy == 1].copy()
acc["exact"] = acc.exact.fillna(0).astype(int)
avm["one"] = ave["one"] = 1

IND = [("Availability, expected medicines (M1)", ave, "available", "one"),
       ("Availability, managed products (M2)", avm, "available", "one"),
       ("Stock-out incidence, Jun-Aug", mg, "anyso", "one"),
       ("Stocked out on the visit day", ad, "out", "one"),
       ("Understocked (<1 month)", ad, "under", "one"),
       ("Adequately stocked (1-2 months)", ad, "adequate", "one"),
       ("Overstocked (>2 months)", ad, "over", "one"),
       ("Inventory record accuracy (exact match)", acc, "exact", "one"),
       ("Order fill rate (quantity)", mg, "dv", "rq")]
est = []
for lab, df, y, x in IND:
    u = ratio_est(df, y, x, weighted=False)
    w = ratio_est(df, y, x, weighted=True)
    est.append(dict(indicator=lab, unweighted=pct(u), weighted=pct(w)))

# CHW (cluster = CHW, no weights)
cm = chw[(chw.managed == 1) & (chw.answered == 1)].copy()
cm["one"] = 1
cm["level"] = "CHW"
u = ratio_est(cm, "available", "one", cluster="chw", strata="level", weighted=False)
est.append(dict(indicator="CHW product availability (managed)", unweighted=pct(u),
                weighted=None))

# prescriptions
px = pd.read_excel(os.path.join(CLEAN, "02_Prescription_clean_2026-11-05.xlsx"),
                   sheet_name="prescriptions")
it = pd.read_excel(os.path.join(CLEAN, "02_Prescription_clean_2026-11-05.xlsx"),
                   sheet_name="medicine_items")
px["complete"] = (px.n_items_unavailable == 0).astype(int)
px["one"] = 1
px["level"] = px.level_group
it["level"] = it.level_group
it["unserved"] = it.is_unavailable_flag.astype(int)
it["one"] = 1
for lab, df, y in (("Prescriptions fully served", px, "complete"),
                   ("Prescribed items unserved", it, "unserved")):
    est.append(dict(indicator=lab,
                    unweighted=pct(ratio_est(df, y, "one", weighted=False)),
                    weighted=pct(ratio_est(df, y, "one", weighted=True))))
RES["estimates"] = est
nohp = []
for lab, df, y, x in IND:
    d = df[df.level != "Health post"]
    nohp.append(dict(indicator=lab, **{k: round(100 * v, 1) for k, v in
                ratio_est(d, y, x, weighted=True).items() if k in ("est", "lo", "hi")}))
RES["estimates_no_hp"] = nohp

# CI figure
fig, ax = plt.subplots(figsize=(7.2, 4.6))
lab = [e["indicator"] for e in est]
y = np.arange(len(est))[::-1]
for yi, e in zip(y, est):
    u = e["unweighted"]
    ax.plot([u["lo"], u["hi"]], [yi + 0.12] * 2, color=GREY, lw=2)
    ax.plot(u["est"], yi + 0.12, "o", color=GREY, ms=4)
    if e["weighted"]:
        w = e["weighted"]
        ax.plot([w["lo"], w["hi"]], [yi - 0.12] * 2, color=NAVY, lw=2)
        ax.plot(w["est"], yi - 0.12, "o", color=NAVY, ms=4)
ax.set_yticks(y)
ax.set_yticklabels(lab, fontsize=8)
ax.set_xlabel("Percent (95% CI)")
ax.plot([], [], "o-", color=GREY, label="Unweighted (pooled, as on the dashboard)")
ax.plot([], [], "o-", color=NAVY, label="Weighted to the national facility mix")
ax.legend(fontsize=7, loc="lower right", frameon=False)
ax.set_title("Headline indicators with design-based 95% confidence intervals",
             fontsize=10, loc="left", color=NAVY, fontweight="bold")
fig.tight_layout()
fig.savefig(os.path.join(OUT, "fig_estimates.png"), dpi=200)
plt.close(fig)

# by level / province with CI (managed availability, stock-out incidence)
grp = {}
for lab, df, y in (("M2 availability", avm, "available"),
                   ("Stock-out incidence", mg, "anyso")):
    rows = []
    for col in ("level", "province", "management", "hub"):
        for k, g in df.groupby(col):
            if g.facility.nunique() < 2:
                continue
            e = ratio_est(g, y, "one", weighted=False)
            rows.append(dict(by=col, group=k, est=round(100 * e["est"], 1),
                             lo=round(100 * e["lo"], 1), hi=round(100 * e["hi"], 1),
                             n=e["n_rows"], fac=e["n_clusters"]))
    grp[lab] = rows
RES["group_ci"] = grp

# district reliability
drel = []
for k, g in avm.groupby("district"):
    e = ratio_est(g, "available", "one", weighted=False)
    drel.append(dict(district=k, est=round(100 * e["est"], 1),
                     lo=round(100 * e["lo"], 1), hi=round(100 * e["hi"], 1),
                     fac=e["n_clusters"]))
nat = ratio_est(avm, "available", "one", weighted=False)["est"] * 100
for d in drel:
    d["vs_national"] = ("below" if d["hi"] < nat else
                        "above" if d["lo"] > nat else "not different")
RES["district_ci_m2"] = sorted(drel, key=lambda d: d["est"])

# ------------------------------------------------------------------ 2 variance partition
def vpc(df, y, name):
    from statsmodels.genmod.bayes_mixed_glm import BinomialBayesMixedGLM
    d = df[[y, "facility", "district"]].copy()
    m = BinomialBayesMixedGLM.from_formula(
        f"{y} ~ 1", {"facility": "0 + C(facility)", "district": "0 + C(district)"}, d)
    r = m.fit_vb()
    sd = np.exp(r.vcp_mean)
    names = list(m.vcp_names)
    var = {n: float(s ** 2) for n, s in zip(names, sd)}
    tot = sum(var.values()) + math.pi ** 2 / 3
    RES[name] = {n: dict(variance=round(v, 3), share=round(100 * v / tot, 1))
                 for n, v in var.items()}
    RES[name]["residual_share"] = round(100 * (math.pi ** 2 / 3) / tot, 1)


vpc(avm, "available", "vpc_availability")
vpc(mg, "anyso", "vpc_stockout")
vpc(acc, "exact", "vpc_accuracy")

# ------------------------------------------------------------------ 3 GEE models
# facility-level covariates
fa = mg.groupby("facility").agg(rq=("rq", "sum"), dv=("dv", "sum"))
fa["fill"] = np.where(fa.rq > 0, fa.dv / fa.rq, np.nan)
fx = acc.groupby("facility").exact.mean()
sysmode = (mg[mg.system.notna() & (mg.system != "")].groupby("facility").system
           .agg(lambda s: s.value_counts().index[0]))
facs["fill10"] = (fa.fill.reindex(facs.index) * 10)
facs["acc10"] = (fx.reindex(facs.index) * 10)
facs["fsys"] = sysmode.reindex(facs.index).fillna("None recorded")
facs["fsys"] = facs.fsys.where(facs.fsys.isin(
    ["eLMIS", "eBuzima", "OpenMRS", "Medisoft"]), "Other / none")
for c in ("fill10", "acc10"):
    facs[c] = facs[c].fillna(facs[c].median())
facs["level2"] = facs.level.replace({"Level 2 teaching hospital": "Referral / teaching",
                                     "National referral hospital": "Referral / teaching"})
# one facility has no GIS setting; it is counted as rural
facs["urban"] = facs.setting.map(lambda s: "urban" if s == "urban" else "rural")
facs["mgmt"] = facs.management.replace({"Faith based facility based": "Faith-based"})
FCOLS = ["level2", "mgmt", "fill10", "acc10", "fsys", "urban"]


def attach(df):
    d = df.merge(facs[FCOLS], left_on="facility", right_index=True, how="left")
    d = d[d.mgmt.isin(["Government", "Faith-based", "PPP"])]
    if "programme" in d:
        d = d[d.programme.notna() & ~d.programme.isin(["(none)", ""])]
    return d


A = attach(avm)
A["programme"] = A.programme.fillna("(none)")
f_av = ("available ~ C(level2, Treatment(reference='Health centre'))"
        " + C(mgmt, Treatment(reference='Government'))"
        " + C(province, Treatment(reference='Kigali City'))"
        " + C(programme, Treatment(reference='Essential Medicines'))"
        " + C(urban, Treatment(reference='rural'))"
        " + C(fsys, Treatment(reference='eLMIS')) + fill10 + acc10")
_, rows = gee_logit(f_av, A, "facility", "gee_availability")
forest(rows, "Model 1. Availability of a managed medicine on the visit day",
       "fig_forest_availability.png")

S = attach(mg)
S["log_amc"] = np.log1p(S.amc.fillna(0))
S["cv"] = S[["cons_jun", "cons_jul", "cons_aug"]].std(axis=1) / \
    S[["cons_jun", "cons_jul", "cons_aug"]].mean(axis=1).replace(0, np.nan)
S["erratic"] = (S.cv.fillna(0) >= 0.5).astype(int)
f_so = ("anyso ~ C(level2, Treatment(reference='Health centre'))"
        " + C(mgmt, Treatment(reference='Government'))"
        " + C(province, Treatment(reference='Kigali City'))"
        " + C(programme, Treatment(reference='Essential Medicines'))"
        " + C(urban, Treatment(reference='rural'))"
        " + log_amc + erratic + fill10 + acc10")
_, rows = gee_logit(f_so, S, "facility", "gee_stockout")
forest(rows, "Model 2. Any stock-out day, June-August (managed pairs)",
       "fig_forest_stockout.png")

Q = attach(acc)
Q["sys"] = Q.system.where(Q.system.isin(["eLMIS", "eBuzima", "OpenMRS", "Medisoft"]),
                          "Other")
f_ac = ("exact ~ C(sys, Treatment(reference='eLMIS'))"
        " + C(level2, Treatment(reference='Health centre'))"
        " + C(mgmt, Treatment(reference='Government'))"
        " + C(province, Treatment(reference='Kigali City'))"
        " + C(programme, Treatment(reference='Essential Medicines'))")
_, rows = gee_logit(f_ac, Q, "facility", "gee_accuracy")
forest(rows, "Model 3. Physical count equals the system count", "fig_forest_accuracy.png")

# ------------------------------------------------------------------ 4 lagged panel
pan = []
for t, t1 in (("jun", "jul"), ("jul", "aug")):
    d = S[["facility", "molecule", "programme", "level2", "mgmt", "province",
           "log_amc", f"so_{t}", f"so_{t1}", f"req_{t}", f"del_{t}"]].copy()
    d.columns = ["facility", "molecule", "programme", "level2", "mgmt", "province",
                 "log_amc", "so_t", "so_t1", "req", "dlv"]
    d["month"] = t1
    pan.append(d)
pan = pd.concat(pan)
pan["out_t"] = (pan.so_t > 0).astype(int)
pan["out_t1"] = (pan.so_t1 > 0).astype(int)
pan["order"] = np.select(
    [pan.req <= 0, pan.dlv >= pan.req, pan.dlv <= 0],
    ["No request", "Delivered in full", "Requested, nothing delivered"],
    "Partly delivered")
pan["pair"] = pan.facility + "|" + pan.molecule
f_lag = ("out_t1 ~ C(order, Treatment(reference='Delivered in full')) + out_t"
         " + log_amc + C(month) + C(level2, Treatment(reference='Health centre'))"
         " + C(programme, Treatment(reference='Essential Medicines'))")
_, rows = gee_logit(f_lag, pan, "facility", "gee_lag",
                    "Outcome: any stock-out day in month t+1 (July or August).")
RES["lag_crosstab"] = (pan.groupby("order").out_t1.agg(["mean", "count"])
                       .assign(mean=lambda x: (100 * x["mean"]).round(1))
                       .reset_index().to_dict("records"))
forest(rows, "Model 4. Stock-out next month by this month's order outcome",
       "fig_forest_lag.png")

# ------------------------------------------------------------------ 5 prescriptions
it["rxkey"] = it._server.astype(str) + "|" + it._parent_index.astype(str) + "|" + \
    it._submission__uuid.astype(str)
px["rxkey"] = px._server.astype(str) + "|" + px._index.astype(str) + "|" + \
    px._submission__uuid.astype(str)
cls = it.drug_class_cleaned.astype(str)
it["abx"] = (cls.str.match(r"^11\.2") | cls.str.match(r"^11\.3 Other anti-bact")).astype(int)
nm = (it["Medicine name"].astype(str) + " " + it["Medicine (name)"].astype(str) + " " +
      it["Medicine name (not on the EML list) - please specify"].astype(str))
it["inj"] = nm.str.contains(r"inj|injection|\biv\b", case=False).astype(int)
it["generic"] = (it["Name type"] == "INN").astype(int)
it["eml"] = it.on_eml.astype(str).str.lower().eq("yes").astype(int)
agg = it.groupby("rxkey").agg(abx=("abx", "max"), inj=("inj", "max"))
agg["n_unav"] = it.groupby("rxkey").unserved.sum()
px = px.merge(agg, left_on="rxkey", right_index=True, how="left")
RES["rx_item_link"] = dict(prescriptions=len(px), linked=int(px.n_unav.notna().sum()))
px["complete"] = (px.n_unav.fillna(px.n_items_unavailable) == 0).astype(int)
px["n_med"] = px["Number of medicines prescribed"]
who = []
for lab, df, y, x in (("Medicines per encounter", px, "n_med", "one"),
                      ("% medicines prescribed by generic (INN) name", it, "generic", "one"),
                      ("% encounters with an antibiotic", px, "abx", "one"),
                      ("% encounters with an injection", px, "inj", "one"),
                      ("% medicines from the EML", it, "eml", "one")):
    e = ratio_est(df.fillna({y: 0}), y, x, weighted=True)
    k = 1 if lab.startswith("Medicines") else 100
    who.append(dict(indicator=lab, est=round(k * e["est"], 2 if k == 1 else 1),
                    lo=round(k * e["lo"], 2 if k == 1 else 1),
                    hi=round(k * e["hi"], 2 if k == 1 else 1)))
RES["who_indicators"] = who
for e in RES["estimates"]:
    if e["indicator"] == "Prescriptions fully served":
        e["unweighted"] = pct(ratio_est(px, "complete", "one", weighted=False))
        e["weighted"] = pct(ratio_est(px, "complete", "one", weighted=True))
lvl_who = []
for lv, g in px.groupby("level"):
    lvl_who.append(dict(level=lv, n=len(g), meds=round(g.n_med.mean(), 2),
                        abx=round(100 * g.abx.mean(), 1), inj=round(100 * g.inj.mean(), 1)))
RES["who_by_level"] = lvl_who
pre = []
for pc, g in px.groupby("Prescriber category"):
    pre.append(dict(prescriber=pc, n=len(g), meds=round(g.n_med.mean(), 2),
                    abx=round(100 * g.abx.mean(), 1), inj=round(100 * g.inj.mean(), 1),
                    unserved=round(100 * (1 - g.complete.mean()), 1)))
RES["who_by_prescriber"] = pre
px["any_unserved"] = 1 - px.complete
px["prescriber"] = px["Prescriber category"].replace(
    {"Allied Health Professional": "Other", "Dental Surgeon": "Other", "Midwife": "Other",
     "General Practioner": "General practitioner"})
px["level2"] = px.level.replace({"Level 2 teaching hospital": "Referral / teaching",
                                 "National referral hospital": "Referral / teaching",
                                 "Health post": "Primary care (HC/HP)",
                                 "Health centre": "Primary care (HC/HP)"})
f_rx = ("any_unserved ~ C(prescriber, Treatment(reference='Nurse')) + n_med"
        " + C(level2, Treatment(reference='Primary care (HC/HP)'))")
_, rows = gee_logit(f_rx, px, "facility", "gee_rx")
forest(rows, "Model 5. Prescription with at least one unserved item", "fig_forest_rx.png")
f_ab = ("abx ~ C(prescriber, Treatment(reference='Nurse')) + n_med"
        " + C(level2, Treatment(reference='Primary care (HC/HP)'))")
_, rows = gee_logit(f_ab, px.fillna({"abx": 0}), "facility", "gee_abx")

# ------------------------------------------------------------------ 6 community
cm = chw[chw.managed == 1].copy()
fam = {"Malaria RDT": "RDT", "Zinc 10mg": "Zinc", "ORS": "ORS", "Amox 125mg": "Amox"}
cm["grp"] = cm["product"].map(lambda p: fam.get(p, "ACT"))
cw = []
for c, g in chw.groupby("chw"):
    m = g[g.managed == 1]
    av = lambda prod: bool(((m["product"] == prod) & (m.available == 1)).any())
    act = m[~m["product"].isin(fam)]
    act_any = bool((act.available == 1).any())
    act_all = bool(len(act) and (act.available == 1).all())
    cw.append(dict(chw=c, facility=g.facility.iloc[0], district=g.district.iloc[0],
                   province=g.province.iloc[0],
                   malaria=int(av("Malaria RDT") and act_any),
                   malaria_all=int(av("Malaria RDT") and act_all),
                   diarrhoea=int(av("ORS") and av("Zinc 10mg")),
                   pneumonia=int(av("Amox 125mg"))))
cw = pd.DataFrame(cw)
cw["iccm"] = cw.malaria * cw.diarrhoea * cw.pneumonia
cw["one"] = 1
cw["level"] = "CHW"
ready = []
for k, lab in (("malaria", "Malaria (RDT + an ACT)"),
               ("malaria_all", "Malaria (RDT + every managed ACT)"),
               ("diarrhoea", "Diarrhoea (ORS + zinc)"), ("pneumonia", "Pneumonia (amoxicillin)"),
               ("iccm", "All three conditions")):
    e = ratio_est(cw, k, "one", cluster="facility", strata="level", weighted=False)
    ready.append(dict(condition=lab, **{kk: round(100 * e[kk], 1) for kk in ("est", "lo", "hi")}))
RES["iccm_ready"] = ready
RES["iccm_by_province"] = (cw.groupby("province")[["malaria", "diarrhoea", "pneumonia", "iccm"]]
                           .mean().mul(100).round(1).reset_index().to_dict("records"))
# paired CHW vs supervising health centre, molecule level
mapc = {"Malaria RDT": "malaria rapid diagnostics ag p f pan t",
        "Pyr/Art 60/20": "pyronaridine artesunate all strengths", "Pyr/Art 180/60": "pyronaridine artesunate all strengths",
        "DHAP 20/160": "dihydroartemisinin piperaquine all strengths",
        "DHAP 30/240": "dihydroartemisinin piperaquine all strengths",
        "DHAP 80/640": "dihydroartemisinin piperaquine all strengths",
        "ALu 1x6": "arte all strengths", "ALu 2x6": "arte all strengths",
        "ALu 3x6": "arte all strengths", "ALu 4x6": "arte all strengths"}
c2 = chw[(chw.managed == 1) & (chw.answered == 1) & chw["product"].isin(mapc)].copy()
c2["molecule"] = c2["product"].map(mapc)
c2 = c2.groupby(["chw", "facility", "molecule"]).available.max().reset_index()
hc = avm.groupby(["facility", "molecule"]).available.max().reset_index() \
    .rename(columns={"available": "hc_av"})
pr = c2.merge(hc, on=["facility", "molecule"], how="inner")
paired = []
for mol, g in pr.groupby("molecule"):
    b = int(((g.available == 1) & (g.hc_av == 0)).sum())
    c_ = int(((g.available == 0) & (g.hc_av == 1)).sum())
    p_ = stats.binomtest(min(b, c_), b + c_, 0.5).pvalue if b + c_ else 1.0
    paired.append(dict(molecule=mol, pairs=len(g), chw=round(100 * g.available.mean(), 1),
                       hc=round(100 * g.hc_av.mean(), 1), chw_only=b, hc_only=c_,
                       both_out=int(((g.available == 0) & (g.hc_av == 0)).sum()),
                       mcnemar_p=float(p_)))
RES["paired_chw_hc"] = paired
tot_b = sum(x["chw_only"] for x in paired)
tot_c = sum(x["hc_only"] for x in paired)
RES["paired_total"] = dict(pairs=len(pr), chw_only=tot_b, hc_only=tot_c,
                           p=float(stats.binomtest(min(tot_b, tot_c), tot_b + tot_c, 0.5).pvalue))
# does HC unavailability raise CHW unavailability?
pr["chw_out"] = 1 - pr.available
pr["hc_out"] = 1 - pr.hc_av
RES["chw_given_hc"] = (pr.groupby("hc_out").chw_out.agg(["mean", "count"])
                       .assign(mean=lambda x: (100 * x["mean"]).round(1)).reset_index()
                       .to_dict("records"))

# ------------------------------------------------------------------ 7 Moran's I
cen2 = cen.set_index("district")
def morans(vals, k=4, perms=9999):
    dd = [d for d in vals if d in cen2.index]
    xy = cen2.loc[dd, ["cx", "cy"]].values
    x = np.array([vals[d] for d in dd], float)
    n = len(x)
    dist = np.sqrt(((xy[:, None, :] - xy[None, :, :]) ** 2).sum(-1))
    Wm = np.zeros((n, n))
    for i in range(n):
        nn = np.argsort(dist[i])[1:k + 1]
        Wm[i, nn] = 1
    Wm = (Wm + Wm.T) > 0
    Wm = Wm / Wm.sum(1, keepdims=True)
    z = x - x.mean()
    I = n / Wm.sum() * (z @ Wm @ z) / (z @ z)
    sims = np.array([(lambda zz: n / Wm.sum() * (zz @ Wm @ zz) / (zz @ zz))(RNG.permutation(z))
                     for _ in range(perms)])
    p = (np.sum(sims >= I) + 1) / (perms + 1)
    return dict(I=round(float(I), 3), expected=round(-1 / (n - 1), 3), p=round(float(p), 4),
                n=n)


dv = lambda df, y: df.groupby("district")[y].mean().to_dict()
RES["moran"] = {
    "M2 availability": morans(dv(avm, "available")),
    "M1 availability": morans(dv(ave, "available")),
    "Stock-out incidence": morans(dv(mg, "anyso")),
    "Record accuracy": morans(dv(acc, "exact")),
    "CHW availability": morans(dv(chw[(chw.managed == 1) & (chw.answered == 1)], "available")),
}

# ------------------------------------------------------------------ 8 typology
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler
fv = pd.DataFrame({
    "availability": avm.groupby("facility").available.mean(),
    "stockout": mg.groupby("facility").anyso.mean(),
    "accuracy": acc.groupby("facility").exact.mean(),
    "fill": fa.fill,
    "overstock": ad.groupby("facility").over.mean(),
    "understock": ad.groupby("facility").under.mean()})
fv = fv.fillna(fv.median())
Z = StandardScaler().fit_transform(fv)
sil = {}
for k in range(3, 7):
    km = KMeans(n_clusters=k, n_init=50, random_state=1).fit(Z)
    sil[k] = silhouette_score(Z, km.labels_)
K = 4
km = KMeans(n_clusters=K, n_init=50, random_state=1).fit(Z)
fv["cluster"] = km.labels_
prof = fv.groupby("cluster").mean().mul(100).round(1)
prof["facilities"] = fv.groupby("cluster").size()
fv2 = fv.join(facs[["level", "province", "management"]])
RES["typology"] = dict(silhouette={k: round(v, 3) for k, v in sil.items()}, k=K,
                       profiles=prof.reset_index().to_dict("records"),
                       by_level=pd.crosstab(fv2.cluster, fv2.level).to_dict(),
                       by_province=pd.crosstab(fv2.cluster, fv2.province).to_dict())
fv2.to_csv(os.path.join(OUT, "facility_typology.csv"))
fig, ax = plt.subplots(figsize=(7.2, 2.8))
pz = (fv.groupby("cluster").mean()[["availability", "stockout", "accuracy", "fill",
                                     "overstock", "understock"]])
zz = (pz - fv.drop(columns="cluster").mean()) / fv.drop(columns="cluster").std()
im = ax.imshow(zz.values, cmap="RdBu", vmin=-1.5, vmax=1.5, aspect="auto")
ax.set_xticks(range(6))
ax.set_xticklabels(["Availability", "Stock-out\nincidence", "Record\naccuracy",
                    "Fill rate", "Overstock", "Understock"], fontsize=8)
ax.set_yticks(range(K))
ax.set_yticklabels([f"Cluster {i + 1} (n={int(prof.loc[i, 'facilities'])})" for i in range(K)],
                   fontsize=8)
for i in range(K):
    for j in range(6):
        ax.text(j, i, f"{pz.values[i, j] * 100:.0f}%", ha="center", va="center", fontsize=7.5,
                color="white" if abs(zz.values[i, j]) > 0.9 else "black")
ax.set_title("Facility typology: cluster means (colour = standardised deviation)",
             fontsize=10, loc="left", color=NAVY, fontweight="bold")
fig.tight_layout()
fig.savefig(os.path.join(OUT, "fig_typology.png"), dpi=200)
plt.close(fig)

# ------------------------------------------------------------------ 9 prediction
from sklearn.ensemble import HistGradientBoostingClassifier, RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (average_precision_score, brier_score_loss, roc_auc_score,
                             roc_curve)
from sklearn.model_selection import GroupKFold
from sklearn.inspection import permutation_importance

P = S.copy()
for m in ("jun", "jul", "aug"):
    P[f"fr_{m}"] = np.where(P[f"req_{m}"] > 0, np.minimum(P[f"del_{m}"] / P[f"req_{m}"].replace(0, np.nan), 2), -1)
    P[f"lreq_{m}"] = np.log1p(P[f"req_{m}"])
    P[f"ldel_{m}"] = np.log1p(P[f"del_{m}"])
    P[f"lcons_{m}"] = np.log1p(P[f"cons_{m}"].fillna(0))
P["lprice"] = np.log1p(P.price.fillna(P.price.median()))
CAT = ["level2", "mgmt", "province", "programme", "urban", "fsys", "molecule"]
for c in CAT:
    P[c] = P[c].fillna("NA").astype(str)


def run_pred(target, num, label, fn):
    X = P[num + CAT].copy()
    y = P[target].values
    groups = P.facility.values
    Xd = pd.get_dummies(X, columns=CAT, drop_first=False).astype(float)
    Xc = X.copy()
    for c in CAT:
        Xc[c] = Xc[c].astype("category").cat.codes
    cat_mask = [c in CAT for c in Xc.columns]
    models = {
        "Logistic regression": lambda: LogisticRegression(max_iter=3000, C=0.5),
        "Random forest": lambda: RandomForestClassifier(n_estimators=400, min_samples_leaf=5,
                                                        n_jobs=-1, random_state=1),
        "Gradient boosting": lambda: HistGradientBoostingClassifier(
            max_iter=300, learning_rate=0.05, max_leaf_nodes=31,
            categorical_features=cat_mask, random_state=1)}
    gkf = GroupKFold(n_splits=5)
    oof = {k: np.zeros(len(y)) for k in models}
    for tr, te in gkf.split(Xc, y, groups):
        for k, mk in models.items():
            Xa = Xc if k == "Gradient boosting" else Xd
            m = mk().fit(Xa.iloc[tr], y[tr])
            oof[k][te] = m.predict_proba(Xa.iloc[te])[:, 1]
    res = dict(target=label, n=int(len(y)), prevalence=round(100 * y.mean(), 1),
               facilities=int(len(set(groups))), models=[])
    for k in models:
        auc = roc_auc_score(y, oof[k])
        # bootstrap CI over facilities
        fac_ix = pd.Series(range(len(y))).groupby(groups).apply(list).tolist()
        boots = []
        for _ in range(300):
            pick = RNG.integers(0, len(fac_ix), len(fac_ix))
            ix = np.concatenate([fac_ix[i] for i in pick])
            if y[ix].min() != y[ix].max():
                boots.append(roc_auc_score(y[ix], oof[k][ix]))
        res["models"].append(dict(model=k, auc=round(auc, 3),
                                  auc_lo=round(float(np.percentile(boots, 2.5)), 3),
                                  auc_hi=round(float(np.percentile(boots, 97.5)), 3),
                                  ap=round(average_precision_score(y, oof[k]), 3),
                                  brier=round(brier_score_loss(y, oof[k]), 4)))
    # operating model: best calibrated (lowest Brier score)
    best = min(res["models"], key=lambda d: d["brier"])["model"]
    res["best"] = best
    pb = oof[best]
    # operating point: flag the top 10% risk
    thr = np.quantile(pb, 0.90)
    flag = pb >= thr
    res["top10"] = dict(threshold=round(float(thr), 3),
                        precision=round(100 * y[flag].mean(), 1),
                        recall=round(100 * y[flag].sum() / max(y.sum(), 1), 1),
                        lift=round(y[flag].mean() / y.mean(), 2))
    # simple rule baseline
    if "so_jul" in num:
        rule = (P.so_jul > 0).values
        res["rule_baseline"] = dict(rule="stocked out in July",
                                    precision=round(100 * y[rule].mean(), 1),
                                    recall=round(100 * y[rule].sum() / y.sum(), 1),
                                    flagged=round(100 * rule.mean(), 1))
        # model at the same flag rate as the rule
        thr2 = np.quantile(pb, 1 - rule.mean())
        f2 = pb >= thr2
        res["model_same_rate"] = dict(flagged=round(100 * f2.mean(), 1),
                                      precision=round(100 * y[f2].mean(), 1),
                                      recall=round(100 * y[f2].sum() / y.sum(), 1))
        # how well does the model rank pairs NOT out in July (new stock-outs)
        new = ~rule
        res["new_stockouts"] = dict(n=int(new.sum()), events=int(y[new].sum()),
                                    auc=round(roc_auc_score(y[new], pb[new]), 3))
    # calibration by decile
    q = pd.qcut(pb, 10, labels=False, duplicates="drop")
    res["calibration"] = (pd.DataFrame({"q": q, "p": pb, "y": y}).groupby("q")
                          .agg(pred=("p", "mean"), obs=("y", "mean")).mul(100).round(1)
                          .reset_index().to_dict("records"))
    # permutation importance (gradient boosting, one held-out fold)
    tr, te = next(gkf.split(Xc, y, groups))
    m = HistGradientBoostingClassifier(max_iter=300, learning_rate=0.05, max_leaf_nodes=31,
                                       categorical_features=cat_mask, random_state=1
                                       ).fit(Xc.iloc[tr], y[tr])
    pi = permutation_importance(m, Xc.iloc[te], y[te], scoring="roc_auc", n_repeats=10,
                                random_state=1)
    imp = sorted(zip(Xc.columns, pi.importances_mean), key=lambda t: -t[1])
    res["importance"] = [dict(feature=f, auc_drop=round(float(v), 4)) for f, v in imp[:12]]
    RES[fn] = res
    # figure: ROC + calibration + importance
    fig, axs = plt.subplots(1, 3, figsize=(10.5, 3.3))
    for k, col in zip(models, (GREY, TEAL, NAVY)):
        fpr, tpr, _ = roc_curve(y, oof[k])
        a = [d for d in res["models"] if d["model"] == k][0]["auc"]
        axs[0].plot(fpr, tpr, color=col, lw=1.6, label=f"{k} (AUC {a:.2f})")
    axs[0].plot([0, 1], [0, 1], ls="--", color="#999", lw=0.8)
    axs[0].set_xlabel("False positive rate")
    axs[0].set_ylabel("True positive rate")
    axs[0].legend(fontsize=6.5, frameon=False, loc="lower right")
    axs[0].set_title("ROC (out-of-fold)", fontsize=9, loc="left")
    cal = pd.DataFrame(res["calibration"])
    axs[1].plot(cal.pred, cal.obs, "o-", color=NAVY)
    mx = max(cal.pred.max(), cal.obs.max()) * 1.05
    axs[1].plot([0, mx], [0, mx], ls="--", color="#999", lw=0.8)
    axs[1].set_xlabel("Predicted risk (%)")
    axs[1].set_ylabel("Observed (%)")
    axs[1].set_title(f"Calibration ({best})", fontsize=9, loc="left")
    im_ = res["importance"][:10][::-1]
    axs[2].barh([FEAT.get(d["feature"], d["feature"]) for d in im_],
                [d["auc_drop"] for d in im_], color=TEAL)
    axs[2].set_xlabel("AUC lost when permuted")
    axs[2].set_title("Most informative features", fontsize=9, loc="left")
    axs[2].tick_params(axis="y", labelsize=7)
    fig.suptitle(label, fontsize=10, x=0.01, ha="left", color=NAVY, fontweight="bold")
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, f"fig_{fn}.png"), dpi=200)
    plt.close(fig)


P["so_aug_any"] = (P.so_aug > 0).astype(int)
num_a = ["so_jun", "so_jul", "lcons_jun", "lcons_jul", "lreq_jun", "lreq_jul",
         "ldel_jun", "ldel_jul", "fr_jun", "fr_jul", "lprice", "fill10", "acc10"]
run_pred("so_aug_any", num_a, "Prediction A. Stock-out in August from June-July data",
         "pred_august")
P["unavail"] = 1 - P.available
num_b = num_a + ["so_aug", "lcons_aug", "lreq_aug", "ldel_aug", "fr_aug"]
run_pred("unavail", num_b,
         "Prediction B. Unavailable on the visit day (September) from June-August data",
         "pred_visit")


# ------------------------------------------------------------------ extra figures
fig, axs = plt.subplots(1, 2, figsize=(10, 3.0))
wi = RES["who_indicators"]
refs = {"Medicines per encounter": "< 2", "% encounters with an antibiotic": "< 30%",
        "% encounters with an injection": "< 20%",
        "% medicines prescribed by generic (INN) name": "100%",
        "% medicines from the EML": "100%"}
ax = axs[0]
labs = ["Antibiotic\nencounters", "Injection\nencounters", "Generic\n(INN)", "From EML"]
vals = [wi[2]["est"], wi[3]["est"], wi[1]["est"], wi[4]["est"]]
err = [[wi[i]["est"] - wi[i]["lo"] for i in (2, 3, 1, 4)],
       [min(wi[i]["hi"], 100) - wi[i]["est"] for i in (2, 3, 1, 4)]]
ax.bar(labs, vals, color=[MAROON, TEAL, NAVY, NAVY], yerr=err, capsize=3)
for x_, r_ in zip(range(4), (30, 20, 100, 100)):
    ax.plot([x_ - 0.4, x_ + 0.4], [r_, r_], color=GOLD, lw=2)
for x_, v in zip(range(4), vals):
    ax.text(x_, v / 2, f"{v:.1f}%", ha="center", fontsize=8.5, color="white", fontweight="bold")
ax.set_ylim(0, 112)
ax.set_ylabel("Percent (95% CI)")
ax.set_title("WHO prescribing indicators (gold = WHO reference)", fontsize=9, loc="left")
ax = axs[1]
rd = RES["iccm_ready"]
ax.barh([r_["condition"] for r_ in rd][::-1], [r_["est"] for r_ in rd][::-1],
        xerr=[[r_["est"] - r_["lo"] for r_ in rd][::-1], [r_["hi"] - r_["est"] for r_ in rd][::-1]],
        color=TEAL, capsize=3)
for y_, r_ in enumerate(rd[::-1]):
    ax.text(r_["hi"] + 1.5, y_, f"{r_['est']:.1f}%", va="center", fontsize=8)
ax.set_xlim(0, 105)
ax.set_xlabel("CHWs ready to treat (%, 95% CI)")
ax.tick_params(axis="y", labelsize=7.5)
ax.set_title("CHW treatment readiness by iCCM condition", fontsize=9, loc="left")
fig.tight_layout()
fig.savefig(os.path.join(OUT, "fig_who_iccm.png"), dpi=200)
plt.close(fig)

# ------------------------------------------------------------------ save
def clean(o):
    if isinstance(o, dict):
        return {str(k): clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [clean(v) for v in o]
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (np.floating,)):
        return None if not np.isfinite(o) else float(o)
    if isinstance(o, float) and not np.isfinite(o):
        return None
    return o


json.dump(clean(RES), open(os.path.join(OUT, "results.json"), "w"), indent=1)
print("done:", ", ".join(RES))
