"""Two-tier tracer medicine availability: basket index and diagnostic index.

    python3 Two_tier_availability.py
      -> Two_tier_availability.html    tabbed dashboard
      -> Two_tier_availability.xlsx    the same tables

Tier 1 is the performance index: how much of each facility's own expected
basket was physically present. Tier 2 is the diagnostic index: which
medicines and which therapeutic categories are short, including where the
Tier 1 score looks healthy.

Both tiers run off the same rows, so an absence reduces the facility's Tier 1
score AND shows in that medicine's Tier 2 rate, as the framework requires.

THE CRITICAL MEDICINE LIST IS PREDEFINED, not derived from the results. It is
declared in CRITICAL below by clinical function - medicines whose absence
causes immediate, irreversible harm - and is printed on its own tab so it can
be reviewed and edited before the findings are read. Nothing in the analysis
chooses it retrospectively.

Denominator rules follow the framework: a medicine not expected at a
facility's level of care is excluded (applicable = No), and a medicine never
assessed is excluded rather than counted as unavailable.

openpyxl + the standard library, like the rest of this project.
"""
import _bootstrap as B   # path setup - must come first

import html
import os
import statistics
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_kobo_prescription as P
import analyze_medicine_assessment as M
import facility_levels as FL
import molecules as MOL

OUT_HTML = B.out("Two_tier_availability.html")
OUT_XLSX = B.out("Two_tier_availability.xlsx")
MIN_N = 30
PERIOD = "visit day, within June-August 2026 fieldwork"

# ---------------------------------------------------------------
# PREDEFINED classifications - declared before any result is seen
# ---------------------------------------------------------------
# Therapeutic category, built from the tool's own drug_class and programme.
# The MCCH drug-class labels are known to be unreliable (caffeine citrate is
# recorded as a haemostatic and tranexamic acid as an anti-apnoeic, which is
# the reverse of the clinical truth), so maternal and newborn medicines are
# taken from the PROGRAMME field instead, which is correct.
CATEGORY_BY_PROGRAMME = {
    "MCCH": "Maternal, newborn and reproductive health",
    "FP": "Maternal, newborn and reproductive health",
    "Malaria": "Antimalarial medicines",
    "TB": "HIV and TB medicines",
    "ARV": "HIV and TB medicines",
}
CATEGORY_BY_CLASS = {
    "Antibiotics": "Antibiotics and anti-infectives",
    "Anti-amoebic": "Antibiotics and anti-infectives",
    "Antihypertensives": "Non-communicable disease medicines",
    "Antidiabetics": "Non-communicable disease medicines",
    "Anticonvulsants": "Non-communicable disease medicines",
    "Lipid-lowering agents": "Non-communicable disease medicines",
    "Antiarrhythmics": "Non-communicable disease medicines",
    "Diuretics": "Non-communicable disease medicines",
    "Corticosteroids": "Non-communicable disease medicines",
    "IV fluids": "Emergency and critical care",
    "Anaesthetics": "Emergency and critical care",
    "Antidotes": "Emergency and critical care",
    "Anticoagulants": "Emergency and critical care",
    "Electrolytes & minerals": "Emergency and critical care",
    "Antiemetics": "Emergency and critical care",
    "Laboratory Diagnostics": "Diagnostics and consumables",
    "Medical supplies": "Diagnostics and consumables",
}
OTHER = "Other essential medicines"

# Critical medicines: absence causes immediate, irreversible harm. Matched on
# a case-insensitive substring of the medicine name so a renamed presentation
# still matches. Reviewed on the "Critical list" tab.
CRITICAL = {
    "Oxytocin": "post-partum haemorrhage",
    "MgSO4": "eclampsia",
    "Carbetocin": "post-partum haemorrhage",
    "Tranexamic": "obstetric haemorrhage",
    "Caffeine Citrate": "neonatal apnoea",
    "ARTE 20MG": "severe/uncomplicated malaria, first line",
    "Dihydroartemisinin": "malaria, first line (MFT)",
    "Pyronaridine": "malaria, first line (MFT)",
    "MALARIA RAPID": "malaria diagnosis - no test, no treatment",
    "ARTESUNATE": "severe malaria",
    "INSULIN": "diabetic emergency",
    "NALOXONE": "opioid overdose antidote",
    "SODIUM CHLORIDE": "resuscitation fluid",
    "ADRENALINE": "anaphylaxis",
    "CALCIUM GLUCONATE": "hyperkalaemia, magnesium toxicity",
}


def s(v):
    return "" if v is None else str(v).strip()


def pct(a, b):
    return a / b if b else None


def med_(v):
    v = [x for x in v if x is not None]
    return round(statistics.median(v), 1) if v else None


def category(r):
    prog = s(r[M.PROGRAM])
    if prog in CATEGORY_BY_PROGRAMME:
        return CATEGORY_BY_PROGRAMME[prog]
    return CATEGORY_BY_CLASS.get(s(r[M.DCLASS]), OTHER)


def is_critical(name):
    low = s(name).lower()
    for k in CRITICAL:
        if k.lower() in low:
            return k
    return None


# ---------------------------------------------------------------
# rows
# ---------------------------------------------------------------
# Method 3: strengths of one molecule are collapsed per facility before
# anything is counted, so a facility holding Metformin 500mg is not recorded
# as missing metformin.
# DEFINITION (study, 2026-10-05): the denominator is every molecule the
# EML authorises at that level of care. A medicine the facility is
# supposed to manage but does not manage at all counts as UNAVAILABLE,
# not as out of scope. MOL.eml_base() applies that and keeps the real
# management flag on `_managed` / the reason on `_gap`.
ALL = MOL.eml_base(M.MEDROWS)
# eligible = expected at this level AND managed AND availability answered
ELIG = [r for r in ALL if M.yes(r[M.APPLICABLE]) and M.yes(r[M.MANAGED])
        and not M.blank(r[M.AVAIL])]
for r in ALL:
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])
    r["_cat"] = category(r)
    r["_crit"] = is_critical(r[M.MED])
AVAIL = [r for r in ELIG if M.yes(r[M.AVAIL])]
OUTS = [r for r in ELIG if M.no(r[M.AVAIL])]

FACS = defaultdict(list)
for r in ELIG:
    FACS[s(r["_p"][M.FAC])].append(r)
FAC_SCORE = {}
for f, rs in FACS.items():
    FAC_SCORE[f] = (sum(1 for r in rs if M.yes(r[M.AVAIL])), len(rs))

LEVELS = FL.ordered(r["_lvl"] for r in ELIG)
CATS = sorted({r["_cat"] for r in ELIG})
UNSERVED = [r for r in P.item if r["_sub"] and r["_avail"] is False]

TABLES = []


def add(tier, tid, title, headers, rows, note="", viz=None):
    TABLES.append((tier, tid, title, headers, rows, note, viz))


# =================== TIER 1 ====================================
pooled = pct(len(AVAIL), len(ELIG))
scores = [pct(a, n) for a, n in FAC_SCORE.values() if n]
complete = sum(1 for a, n in FAC_SCORE.values() if n and a == n)
add("T1", "T1.1-T1.4", "The four Tier 1 indicators",
    ["Indicator", "Numerator", "Denominator", "Result"],
    [["T1.2 Overall pooled tracer availability", len(AVAIL), len(ELIG), pooled],
     ["T1.3 Mean facility basket availability",
      f"sum of {len(scores)} facility scores", len(scores),
      statistics.mean(scores) if scores else None],
     ["T1.4 Complete basket availability", complete, len(FAC_SCORE),
      pct(complete, len(FAC_SCORE))],
     ["T1.1 Facility basket availability",
      "one score per facility - see the distribution below", len(FAC_SCORE),
      None]],
    "Pooled and mean are reported separately because facility baskets are not "
    "the same size, exactly as the framework requires. Eligible = expected at "
    "that level per the EML, managed by the facility, and availability "
    f"answered: {len(ELIG):,} of {len(ALL):,} observations. Reference: {PERIOD}.")

bands = [(1.0, 1.0, "100% - complete basket"), (0.95, 1.0, "95 to 99%"),
         (0.90, 0.95, "90 to 94%"), (0.80, 0.90, "80 to 89%"),
         (0.60, 0.80, "60 to 79%"), (0.0, 0.60, "below 60%")]
dist = []
for lo, hi, label in bands:
    n = sum(1 for v in scores
            if (v >= lo if lo == hi else (lo <= v < hi if hi < 1.0 else lo <= v < 1.0)))
    if label.startswith("100%"):
        n = sum(1 for v in scores if v >= 1.0)
    dist.append([label, n, pct(n, len(scores))])
add("T1", "T1.1", "Distribution of facility basket availability scores",
    ["Band", "Facilities", "% of facilities"], dist,
    "The shape behind the mean. A facility scoring 85% is not a failure - it "
    "is three medicines short of twenty.", "bar")


def tier1_by(keyfn, keys=None):
    g = defaultdict(list)
    for r in ELIG:
        g[keyfn(r)].append(r)
    out = []
    for k in (keys or sorted(g)):
        rs = g.get(k, [])
        facs = {s(r["_p"][M.FAC]) for r in rs}
        fs = [pct(*FAC_SCORE[f]) for f in facs if FAC_SCORE.get(f, (0, 0))[1]]
        comp = sum(1 for f in facs if FAC_SCORE[f][1]
                   and FAC_SCORE[f][0] == FAC_SCORE[f][1])
        av = sum(1 for r in rs if M.yes(r[M.AVAIL]))
        out.append([k, len(facs), len(rs), av, len(rs) - av,
                    pct(av, len(rs)) if len(rs) >= MIN_N else None,
                    statistics.mean(fs) if fs else None,
                    med_(fs), min(fs) if fs else None,
                    pct(comp, len(facs))])
    return out


T1_HDR = ["Group", "Facilities", "Eligible observations", "Available",
          "Missing", "POOLED AVAILABILITY", "Mean facility score",
          "Median facility score", "Lowest facility", "Complete basket %"]
add("T1", "T1.2-T1.4", "By level of care", T1_HDR,
    tier1_by(lambda r: r["_lvl"], LEVELS),
    "Pooled and mean differ wherever baskets differ in size within a tier.",
    "bar")
add("T1", "T1.2-T1.4", "By province", T1_HDR,
    tier1_by(lambda r: s(r["_p"][M.PROV])), "", "bar")
add("T1", "T1.2-T1.4", "By district, ranked",
    T1_HDR, sorted(tier1_by(lambda r: s(r["_p"][M.DIST])),
                   key=lambda x: (x[5] is None, x[5])),
    "Worst pooled availability first.", "bar")

_flist = sorted(([f, s(FACS[f][0]["_p"][M.DIST]), FACS[f][0]["_lvl"],
                  FAC_SCORE[f][1], FAC_SCORE[f][0],
                  FAC_SCORE[f][1] - FAC_SCORE[f][0], pct(*FAC_SCORE[f])]
                 for f in FACS if FAC_SCORE[f][1]), key=lambda x: x[6])
add("T1", "T1.1", "Twenty facilities with the lowest basket availability",
    ["Facility", "District", "Level of care", "Expected assessed", "Available",
     "Missing", "FACILITY SCORE"], _flist[:20],
    "The Tier 1 action list.", "bar")

# =================== TIER 2 ====================================
def med_rows():
    out = []
    for m in {s(r[M.MED]) for r in ELIG}:
        rs = [r for r in ELIG if s(r[M.MED]) == m]
        if len(rs) < MIN_N:
            continue
        av = sum(1 for r in rs if M.yes(r[M.AVAIL]))
        out.append([m, rs[0]["_cat"], "yes" if rs[0]["_crit"] else "",
                    len({s(r["_p"][M.FAC]) for r in rs}), len(rs), av,
                    len(rs) - av, pct(av, len(rs))])
    return sorted(out, key=lambda x: x[7])


MEDS = med_rows()
add("T2", "T2.1", "Individual tracer medicine availability - all medicines",
    ["Medicine", "Therapeutic category", "Critical?", "Facilities",
     "Eligible observations", "Available", "Missing", "AVAILABILITY"],
    MEDS, f"Every medicine with at least {MIN_N} eligible observations, worst "
    f"first. This is the diagnostic index: the pooled Tier 1 score of "
    f"{pooled:.1%} hides everything in the top rows of this table.", "bar")
add("T2", "T2.1", "The ten least available tracer medicines",
    ["Medicine", "Therapeutic category", "Critical?", "Facilities",
     "Eligible observations", "Available", "Missing", "AVAILABILITY"],
    MEDS[:10], "", "bar")

cat_rows = []
for c in CATS:
    rs = [r for r in ELIG if r["_cat"] == c]
    av = sum(1 for r in rs if M.yes(r[M.AVAIL]))
    facs = {s(r["_p"][M.FAC]) for r in rs}
    worst = sorted(({s(x[M.MED]) for x in rs}), key=lambda m: pct(
        sum(1 for x in rs if s(x[M.MED]) == m and M.yes(x[M.AVAIL])),
        sum(1 for x in rs if s(x[M.MED]) == m)) or 0)
    cat_rows.append([c, len({s(r[M.MED]) for r in rs}), len(facs), len(rs), av,
                     len(rs) - av, pct(av, len(rs)),
                     worst[0][:44] if worst else ""])
cat_rows.sort(key=lambda x: x[6])
add("T2", "T2.2", "Therapeutic category availability",
    ["Therapeutic category", "Medicines", "Facilities",
     "Eligible observations", "Available", "Missing", "CATEGORY AVAILABILITY",
     "Weakest medicine in the category"], cat_rows,
    "Categories are predefined: maternal/newborn and reproductive health from "
    "the MCCH and FP programmes, antimalarials from the Malaria programme, "
    "HIV and TB from their programmes, and the rest from drug class. The MCCH "
    "drug-class labels are unreliable, so the programme field is used for "
    "them.", "bar")

crit_rows = []
for m in sorted({s(r[M.MED]) for r in ELIG if r["_crit"]}):
    rs = [r for r in ELIG if s(r[M.MED]) == m]
    out = sum(1 for r in rs if M.no(r[M.AVAIL]))
    crit_rows.append([m, CRITICAL[rs[0]["_crit"]], rs[0]["_cat"],
                      len({s(r["_p"][M.FAC]) for r in rs}), len(rs),
                      len(rs) - out, out, pct(out, len(rs))])
crit_rows.sort(key=lambda x: -(x[7] or 0))
add("T2", "T2.3", "Critical medicine shortage rate",
    ["Critical medicine", "Why it is critical", "Therapeutic category",
     "Facilities", "Eligible observations", "Available", "SHORTAGES",
     "SHORTAGE RATE"], crit_rows,
    "Shortage rate, not availability - higher is worse. The list is "
    "predefined by clinical function and is on the 'Critical list' tab.",
    "bar")

crit_fac = defaultdict(lambda: [0, 0])
for r in ELIG:
    if r["_crit"]:
        f = s(r["_p"][M.FAC])
        crit_fac[f][1] += 1
        if M.no(r[M.AVAIL]):
            crit_fac[f][0] += 1
exposed = sum(1 for f, (o, n) in crit_fac.items() if o > 0)
add("T2", "T2.4", "Facilities exposed to at least one critical shortage",
    ["Measure", "Value"],
    [["Facilities assessed on at least one critical medicine", len(crit_fac)],
     ["Facilities missing at least one critical medicine", exposed],
     ["CRITICAL SHORTAGE EXPOSURE", pct(exposed, len(crit_fac))],
     ["Facilities missing two or more",
      sum(1 for f, (o, n) in crit_fac.items() if o >= 2)],
     ["Facilities missing three or more",
      sum(1 for f, (o, n) in crit_fac.items() if o >= 3)]],
    "A facility counts as exposed if any one predefined critical medicine it "
    "was expected to hold was absent on the visit day.")

burden = []
for m, cat, crit, nf, n, av, miss, rate in MEDS:
    rs = [r for r in M.MEDROWS if s(r[M.MED]) == m and M.yes(r[M.MANAGED])]
    affected = [r for r in rs
                if sum(1 for mo in ("June", "July", "August")
                       if (M.num(r[M.SO_MONTH[mo]]) or 0) > 0) > 0]
    days = [M.num(r[M.SO_TOTAL]) or 0 for r in affected]
    recur = sum(1 for r in affected
                if sum(1 for mo in ("June", "July", "August")
                       if (M.num(r[M.SO_MONTH[mo]]) or 0) > 0) >= 2)
    un = [r for r in UNSERVED if s(r["_med"]).lower()[:12] == m.lower()[:12]]
    burden.append([m, crit, rate, pct(len(affected), len(rs)),
                   round(statistics.mean(days), 1) if days else None,
                   med_(days), recur, len(un),
                   round(sum(P.num(r.get(P.COST_UNSERVED)) or 0 for r in un))])
burden.sort(key=lambda x: (x[2] if x[2] is not None else 2))
add("T2", "T2.5", "Medicine-specific stock-out burden",
    ["Medicine", "Critical?", "Availability on the day",
     "Three-month stock-out incidence", "Mean stock-out days",
     "Median stock-out days", "Facilities with recurrent stock-outs",
     "Unserved prescriptions", "Unserved value (RWF)"], burden[:25],
    "Twenty-five weakest medicines. Incidence is the share of managed pairs "
    "with any stock-out in the window; recurrent means two or three months "
    "affected. Mean and median are days per affected pair, NOT per episode - "
    "the form counts days, not episodes. The prescription link is approximate: "
    "the two tools name medicines differently, so it matches on a name prefix "
    "and will miss some.", "bar")

# =================== INTEGRATION ===============================
hidden = []
for f, (a, n) in FAC_SCORE.items():
    if not n or pct(a, n) < 0.90:
        continue
    miss = [r for r in FACS[f] if M.no(r[M.AVAIL]) and r["_crit"]]
    if miss:
        hidden.append([f, s(FACS[f][0]["_p"][M.DIST]), FACS[f][0]["_lvl"],
                       pct(a, n), len(miss),
                       ", ".join(sorted({s(r[M.MED])[:28] for r in miss}))[:90]])
hidden.sort(key=lambda x: -x[3])
add("INT", "Integration",
    "Facilities scoring 90%+ on Tier 1 that are still missing a critical medicine",
    ["Facility", "District", "Level of care", "Tier 1 score",
     "Critical medicines missing", "Which"], hidden,
    "THE REASON THE TWO TIERS MUST BE READ TOGETHER. Each of these facilities "
    "would pass any Tier 1 threshold and is still unable to manage an "
    "emergency. A high basket score is not a statement about patient access.")

add("INT", "Integration", "Integrated reporting table",
    ["Analytical level", "Indicator", "Result"],
    [["Tier 1", "Overall pooled tracer basket availability", pooled],
     ["Tier 1", "Mean facility basket availability",
      statistics.mean(scores) if scores else None],
     ["Tier 1", "Facilities with complete basket availability",
      pct(complete, len(FAC_SCORE))],
     ["Tier 2", "Critical shortage exposure (facilities)",
      pct(exposed, len(crit_fac))]]
    + [["Tier 2", f"{c[0]} availability", c[6]] for c in cat_rows],
    "The framework's integrated table, with real results. Read the Tier 1 "
    "rows as supply chain performance and the Tier 2 rows as clinical "
    "vulnerability - they do not move together.")

_district_cat = []
for d in sorted({s(r["_p"][M.DIST]) for r in ELIG}):
    rs = [r for r in ELIG if s(r["_p"][M.DIST]) == d]
    row = [d, pct(sum(1 for r in rs if M.yes(r[M.AVAIL])), len(rs))]
    for c in CATS:
        sub = [r for r in rs if r["_cat"] == c]
        row.append(pct(sum(1 for r in sub if M.yes(r[M.AVAIL])), len(sub))
                   if len(sub) >= 10 else None)
    _district_cat.append(row)
_district_cat.sort(key=lambda x: (x[1] is None, x[1]))
add("INT", "Integration", "District heatmap - Tier 1 score and category availability",
    ["District", "Tier 1 pooled"] + [c[:26] for c in CATS], _district_cat,
    "Two districts can share a Tier 1 score and be short in entirely "
    "different categories - that is what this table is for. Blank where the "
    "district has fewer than 10 eligible observations in a category.")

add("INT", "Method", "Classifications used, and the rules applied",
    ["Item", "Detail"],
    [["Eligible denominator",
      "expected at that level per the EML, managed by the facility, "
      "availability answered. Medicines not expected at a level are excluded; "
      "unassessed medicines are excluded, never counted as unavailable."],
     ["Tier 1 / Tier 2 consistency",
      "one absence reduces that facility's Tier 1 score and appears in that "
      "medicine's Tier 2 rate - both tiers read the same rows."],
     ["Critical list", "predefined by clinical function, declared before any "
      "result was computed, listed on the 'Critical list' tab."],
     ["Therapeutic categories",
      "maternal/newborn and reproductive health, antimalarials and HIV/TB "
      "from the programme field; the rest from drug class. MCCH drug-class "
      "labels are unreliable (caffeine citrate recorded as a haemostatic, "
      "tranexamic acid as an anti-apnoeic) so programme is used for them."],
     ["What a Tier 1 score is NOT",
      "it is not the probability a patient receives a prescribed medicine. "
      "Patient access is the prescription fulfilment analysis, which is a "
      "different denominator and a different tool."],
     ["Cleaning in force",
      "two servers merged, 12 duplicate facility x medicine rows removed, 7 "
      "facility reattributions, levels grouped, non-finite calculate fields "
      "rejected."],
     ["Sampling", "facilities follow a visit plan, not a probability sample, "
      "so no weights are applied and nothing is extrapolated nationally. The "
      "framework's call for sampling weights cannot be met."]])

add("INT", "Critical list", "The predefined critical medicine list",
    ["Match on name contains", "Why it is critical", "Medicines matched",
     "Eligible observations"],
    [[k, v, ", ".join(sorted({s(r[M.MED])[:36] for r in ELIG
                              if r["_crit"] == k})) or "none in this dataset",
      sum(1 for r in ELIG if r["_crit"] == k)]
     for k, v in CRITICAL.items()],
    "Declared in the script before any result is computed, and matched on a "
    "case-insensitive substring so a renamed presentation still matches. Edit "
    "CRITICAL in Two_tier_availability.py to change it; entries matching "
    "nothing in this dataset are shown so the list stays auditable.")

TIERS = {"T1": "Tier 1 - basket availability index",
         "T2": "Tier 2 - medicine and category diagnostics",
         "INT": "Integration, method and critical list"}


# ---------------------------------------------------------------
# renderers - one set of tables, two outputs
# ---------------------------------------------------------------
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, size=9, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=12, color="1F3864")


def esc(v):
    return html.escape("" if v is None else str(v), quote=True)


def fmt(v):
    if v is None or v == "":
        return ""
    if isinstance(v, float) and 0 <= v <= 1:
        return f"{v * 100:.1f}%"
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return f"{v:,.0f}" if float(v).is_integer() else f"{v:,.1f}"
    return str(v)


SHORT = {"T1": "Tier 1", "T2": "Tier 2", "INT": "Integration"}


def write_xlsx():
    wb = Workbook()
    wb.remove(wb.active)
    for tier, name in TIERS.items():
        ws = wb.create_sheet(SHORT[tier][:31])
        ws.sheet_view.showGridLines = False
        ws.cell(row=1, column=1, value=name).font = Font(
            bold=True, size=14, color="1F3864")
        r, widths = 3, {}
        for t, tid, title, headers, rows, note, viz in TABLES:
            if t != tier:
                continue
            ws.cell(row=r, column=1, value=f"{tid}  {title}").font = TITLE_FONT
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
                c = ws.cell(row=r, column=ci, value=h)
                c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
                c.alignment = Alignment(wrap_text=True, vertical="center")
                widths[ci] = max(widths.get(ci, 10), min(len(str(h)) + 2, 30))
            r += 1
            for row in rows:
                for ci, v in enumerate(row, start=1):
                    c = ws.cell(row=r, column=ci, value=v)
                    c.border = BORDER
                    c.alignment = Alignment(vertical="top",
                                            wrap_text=isinstance(v, str)
                                            and len(str(v)) > 34)
                    if isinstance(v, float) and 0 <= v <= 1:
                        c.number_format = "0.0%"
                    elif isinstance(v, (int, float)):
                        c.number_format = "#,##0.##"
                    widths[ci] = max(widths.get(ci, 10),
                                     min(len(str(v)) + 2, 48))
                r += 1
            r += 2
        for ci, w in widths.items():
            ws.column_dimensions[get_column_letter(ci)].width = w
    wb.active = 0
    if os.path.exists(OUT_XLSX):
        os.chmod(OUT_XLSX, 0o644)
    wb.save(OUT_XLSX)


CSS = """
<style>
:root{color-scheme:dark;
 --plane:#000;--surface:#1a1a19;--raise:#222221;--ink:#fff;--ink2:#c3c2b7;
 --muted:#898781;--grid:#2c2c2a;--axis:#383835;--rule:rgba(255,255,255,.14);
 --t1:#3987e5;--t2:#d95926;--int:#199e70;--accent:var(--t1);
 --good:#199e70;--warn:#eda100;--bad:#e66767;
 --font:"Trebuchet MS","Lucida Grande","Lucida Sans Unicode",Tahoma,sans-serif;}
*{box-sizing:border-box}
body{margin:0;background:var(--plane);color:var(--ink);font-family:var(--font);
 font-size:16px;line-height:1.55}
.wrap{max-width:1240px;margin:0 auto;padding:0 20px 72px}
h1,h2,h3,h4{font-family:var(--font);margin:0;text-wrap:balance}
.mast{border-bottom:1px solid var(--rule);padding:44px 0 26px}
.eyebrow{font-size:12.5px;letter-spacing:.14em;text-transform:uppercase;
 color:var(--muted);margin:0 0 12px;font-weight:600}
.mast h1{font-size:clamp(28px,4.6vw,44px);letter-spacing:-.015em}
.mast p{color:var(--ink2);max-width:82ch;margin:14px 0 0;font-size:15px}
.two{display:grid;grid-template-columns:1fr 1fr;gap:1px;background:var(--rule);
 border:1px solid var(--rule);margin:26px 0 0}
.card{background:var(--surface);padding:20px}
.card .who{font-size:12px;letter-spacing:.1em;text-transform:uppercase;
 font-weight:700;margin:0 0 10px}
.card.a .who{color:var(--t1)}.card.b .who{color:var(--t2)}
.card .big{font-size:42px;font-weight:600;margin:0;line-height:1}
.card.a .big{color:var(--t1)}.card.b .big{color:var(--t2)}
.card .cap{margin:8px 0 0;font-size:13.5px;color:var(--ink2)}
.verdict{background:#16202b;border:1px solid var(--rule);border-top:0;
 padding:16px 20px;font-size:14px;color:var(--ink2)}
.verdict b{color:var(--ink)}
.tabs{display:flex;flex-wrap:wrap;gap:2px;margin:30px 0 0;
 border-bottom:1px solid var(--rule)}
.tab{font:inherit;font-size:14.5px;color:var(--ink2);background:none;border:0;
 border-bottom:2px solid transparent;padding:11px 16px;margin-bottom:-1px;
 cursor:pointer;white-space:nowrap}
.tab:hover{color:var(--ink)}
.tab.on{color:var(--ink);font-weight:600}
.tab[data-page="T1"].on{border-bottom-color:var(--t1)}
.tab[data-page="T2"].on{border-bottom-color:var(--t2)}
.tab[data-page="INT"].on{border-bottom-color:var(--int)}
.js .page{display:none}.js .page.on{display:block}
.page{margin-top:34px}
#T1{--accent:var(--t1)}#T2{--accent:var(--t2)}#INT{--accent:var(--int)}
.page>h2{font-size:23px;padding-bottom:8px;border-bottom:2px solid var(--accent)}
.ind{margin-top:30px}
.ind h4{font-size:15.5px;font-weight:600;display:flex;gap:10px;
 align-items:baseline;flex-wrap:wrap}
.mid{font-size:12px;color:var(--accent);font-weight:700;letter-spacing:.04em}
.note{margin:7px 0 10px;font-size:13px;color:var(--muted);max-width:94ch;
 line-height:1.6}
.scroll{overflow-x:auto;border:1px solid var(--rule);background:var(--surface)}
.scroll.tall{max-height:470px;overflow-y:auto}
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
td.heat{text-align:right;font-variant-numeric:tabular-nums;color:#06130c}
.bar{display:inline-block;height:9px;border-radius:99px;background:var(--good);
 vertical-align:middle;margin-left:8px}
.bar.mid{background:var(--warn)}.bar.low{background:var(--bad)}
.crit{color:var(--bad);font-weight:700}
@media (max-width:760px){.two{grid-template-columns:1fr}
 .tab{padding:10px 10px;font-size:13px}}
@media print{.js .page{display:block!important}.tabs{display:none}
 .page{break-before:page}body{background:#fff;color:#000}}
</style>
"""


def html_table(headers, rows, heat=False):
    body = []
    for row in rows:
        tds = []
        for i, (h, v) in enumerate(zip(headers, row)):
            num = isinstance(v, (int, float)) and not isinstance(v, bool)
            cell = esc(fmt(v))
            if v == "yes" and "Critical" in str(h):
                cell = '<span class="crit">CRITICAL</span>'
            if heat and isinstance(v, float) and 0 <= v <= 1 and i > 0:
                g = int(60 + 150 * v)
                tds.append(f'<td class="heat" style="background:rgb('
                           f'{int(230 - 160 * v)},{g},90)">{cell}</td>')
                continue
            if isinstance(v, float) and 0 <= v <= 1 and not heat:
                w = max(2, round(v * 52))
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


def write_html():
    worst_cat = cat_rows[0] if cat_rows else ["", 0, 0, 0, 0, 0, 0, ""]
    body_pages = []
    for i, (tier, name) in enumerate(TIERS.items()):
        parts = [f'<div class="page{" on" if i == 0 else ""}" id="{tier}">'
                 f'<h2>{esc(name)}</h2>']
        for t, tid, title, headers, rows, note, viz in TABLES:
            if t != tier:
                continue
            parts.append(
                f'<section class="ind"><h4><span class="mid">{esc(tid)}</span>'
                f'{esc(title)}</h4>'
                + (f'<p class="note">{esc(note)}</p>' if note else "")
                + html_table(headers, rows, heat="heatmap" in title.lower())
                + "</section>")
        parts.append("</div>")
        body_pages.append("".join(parts))
    tabs = "".join(
        f'<button class="tab{" on" if i == 0 else ""}" id="tab-{t}" '
        f'data-page="{t}">{esc(n)}</button>'
        for i, (t, n) in enumerate(TIERS.items()))
    body = f"""
<div class="wrap">
<header class="mast">
  <p class="eyebrow">CHAI Rwanda &middot; UP study &middot; two-tier availability</p>
  <h1>Tracer medicine availability: performance index and clinical diagnostics</h1>
  <p>Tier 1 measures how much of each facility's own expected basket was on
  the shelf. Tier 2 asks which medicines and which therapeutic categories are
  short &mdash; including where the Tier 1 score looks healthy. The two are
  reported together because a high basket score can conceal the absence of a
  medicine that decides whether a patient survives the day.</p>
  <div class="two">
    <div class="card a"><p class="who">Tier 1 &middot; performance index</p>
      <p class="big">{pooled * 100:.1f}%</p>
      <p class="cap">of {len(ELIG):,} eligible facility-medicine observations
      were available. Mean facility basket score
      {statistics.mean(scores) * 100:.1f}%, and
      {pct(complete, len(FAC_SCORE)) * 100:.1f}% of facilities held their
      entire basket.</p></div>
    <div class="card b"><p class="who">Tier 2 &middot; diagnostic index</p>
      <p class="big">{pct(exposed, len(crit_fac)) * 100:.1f}%</p>
      <p class="cap">of facilities were missing at least one predefined
      critical medicine. Weakest category:
      {esc(worst_cat[0])} at {worst_cat[6] * 100:.1f}%.</p></div>
  </div>
  <div class="verdict"><b>Read them together.</b> {len(hidden)} facilities
  score 90% or better on Tier 1 and are still missing a critical medicine.
  A basket score is a supply chain performance measure; it is not the
  probability that a patient receives what was prescribed.</div>
</header>
<nav class="tabs">{tabs}</nav>
{''.join(body_pages)}
</div>
<script>
(function(){{
 var ids=["T1","T2","INT"];
 function show(id){{ids.forEach(function(o){{
   document.getElementById(o).classList.toggle('on',o===id);
   document.getElementById('tab-'+o).classList.toggle('on',o===id);}});}}
 document.addEventListener('click',function(e){{
   var b=e.target.closest('[data-page]'); if(!b)return;
   show(b.dataset.page); history.replaceState(null,'','#'+b.dataset.page);
   window.scrollTo(0,0);}});
 var h=location.hash.slice(1); if(ids.indexOf(h)>=0) show(h);
}})();
</script>
"""
    doc = ('<!doctype html>\n<html lang="en" class="js">\n<head>\n'
           '<meta charset="utf-8">\n<meta name="viewport" '
           'content="width=device-width,initial-scale=1">\n'
           '<title>Two-tier tracer medicine availability</title>\n' + CSS
           + "</head>\n<body>\n" + body + "\n</body>\n</html>\n")
    if os.path.exists(OUT_HTML):
        os.chmod(OUT_HTML, 0o644)
    with open(OUT_HTML, "w", encoding="utf-8") as fh:
        fh.write(doc)


if __name__ == "__main__":
    write_xlsx()
    write_html()
    print(f"wrote {os.path.basename(OUT_HTML)} and "
          f"{os.path.basename(OUT_XLSX)}")
    print(f"  {len(TABLES)} tables across Tier 1, Tier 2 and Integration")
    print(f"  TIER 1 pooled {pooled:.1%} | mean facility "
          f"{statistics.mean(scores):.1%} | complete basket "
          f"{pct(complete, len(FAC_SCORE)):.1%}")
    print(f"  TIER 2 critical shortage exposure "
          f"{pct(exposed, len(crit_fac)):.1%} of {len(crit_fac)} facilities")
    print(f"  weakest category: {cat_rows[0][0]} at {cat_rows[0][6]:.1%}")
    print(f"  {len(hidden)} facilities at 90%+ Tier 1 still missing a "
          f"critical medicine")
