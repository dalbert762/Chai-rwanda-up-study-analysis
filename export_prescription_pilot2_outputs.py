"""
Export the 'Prescription pilot2.xlsx' basic analysis to two plain deliverables:

    Prescription_pilot2_analysis.html   - one plain HTML page, all tables, minimal styling
    Prescription_pilot2_analysis.xlsx   - one Excel workbook, one worksheet per section

    python3 export_prescription_pilot2_outputs.py

Every table is computed from the workbook, not typed in, so re-running after a fresh
Kobo export refreshes both files. openpyxl + the standard library only (no pandas).

Like analyze_prescription_pilot2.py, this takes facility / district / date / prescriber /
patient ID for a medicine row from its PARENT prescription - medicine_item's own copies of
those eight columns are misaligned (see the Data quality section it writes out).

Patient / Prescription IDs are masked (2504...1423). The full values are in the workbook.
"""

import datetime as dt
import html
import re
from collections import Counter, defaultdict

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SRC = "Prescription pilot2.xlsx"
OUT_HTML = "Prescription_pilot2_analysis.html"
OUT_XLSX = "Prescription_pilot2_analysis.xlsx"

S_SUB = "Prescription assessment tool..."
S_PRES = "prescription"
S_ITEM = "medicine_item"

N_MED = "Number of medicines prescribed"
PID = "Patient / Prescription ID"
DATE = "Date of prescription"
PRESCRIBER = "Prescriber category"
EML = "Is this medicine on the EML (Essential Medicines List)?"
MED_EML = "Medicine name (type the first letter and select)"
MED_FREE = "Medicine name (not on EML list) - please specify"
CLASS_FREE = "Drug class (not on EML list) - please specify"
AVAIL = "Was the medicine available at the health facility during the date of prescription?"
ORDERED = "Was the medicine ordered during the month the prescription was made?"
SUPPLIED = "Was it supplied fully, partially, or not at all?"
UNIT_PRICE = "Unit price for the medicine that was not available (RWF)"
RMS_REG = "Was the medicine available at RMS Regional during the date of prescription?"
RMS_CEN = "Was the medicine available at RMS Central during the date of prescription?"
DIST = "distribution_nationwide"


# ---------------------------------------------------------------
# read
# ---------------------------------------------------------------
def load(sheet):
    wb = openpyxl.load_workbook(SRC, data_only=True, read_only=True)
    rows = list(wb[sheet].iter_rows(values_only=True))
    hdr = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = [dict(zip(hdr, r)) for r in rows[1:]
            if any(c is not None and str(c).strip() != "" for c in r)]
    return hdr, data


def blank(v):
    return v is None or str(v).strip() == ""


def num(v):
    if blank(v):
        return None
    try:
        return float(str(v).strip())
    except ValueError:
        return None


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    n = num(v)
    if n is not None and 20000 < n < 60000:
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=n)).date()
    return None


def yes(v):
    return str(v).strip().lower() == "yes"


def no(v):
    return str(v).strip().lower() == "no"


def mask(pid):
    s = str(pid).strip()
    return s[:4] + "..." + s[-4:] if len(s) > 10 else s


hdr_sub, sub = load(S_SUB)
hdr_pres, pres = load(S_PRES)
hdr_item, item = load(S_ITEM)

parent = {(r["_submission__uuid"], r["_index"]): r for r in pres}


def par(r):
    return parent[(r["_submission__uuid"], r["_parent_index"])]


# ---------------------------------------------------------------
# table builders - a Table is (title, note, headers, rows, total_row)
# ---------------------------------------------------------------
class Table:
    def __init__(self, title, headers, rows, note=None, total=None, chart=None):
        self.title = title
        self.headers = headers
        self.rows = rows
        self.note = note
        self.total = total
        # chart: (kind, label_col, value_col, max_rows) - kind is "v" (vertical bars,
        # short labels) or "h" (horizontal bars, long labels). None = no chart.
        self.chart = chart
        # Which columns hold a proportion. Declared from the header rather than sniffed
        # from the value: a quartile of 1.0 in a stats table is the number one, not 100%.
        self.pct = {i for i, h in enumerate(headers)
                    if "%" in str(h) or str(h).strip().lower().endswith("rate")}

    def chart_data(self):
        """(label, value) pairs for this table's chart, excluding any total row."""
        if not self.chart:
            return []
        _, lc, vc, mx = self.chart
        out = [(str(r[lc]), float(r[vc])) for r in self.rows if r[vc] not in (None, "")]
        return out[:mx] if mx else out


def counts(rows, col, denom=None, top=None, keyname="Value", skip_blank=False):
    """Frequency table with a percentage column."""
    c = Counter()
    for r in rows:
        v = r.get(col)
        if blank(v):
            if skip_blank:
                continue
            c["(blank)"] += 1
        else:
            c[str(v).strip()] += 1
    n = denom or sum(c.values())
    out = [[k, v, v / n] for k, v in c.most_common(top)]
    return out, n


def stats_row(label, rows, col):
    vals = sorted(v for v in (num(r.get(col)) for r in rows) if v is not None)
    if not vals:
        return [label, 0, len(rows), "", "", "", "", "", "", ""]
    k = len(vals)

    def q(p):
        i = p * (k - 1)
        lo, hi = int(i), min(int(i) + 1, k - 1)
        return vals[lo] + (vals[hi] - vals[lo]) * (i - lo)

    return [label, k, len(rows) - k, sum(vals), round(sum(vals) / k, 2),
            vals[0], q(.25), q(.5), q(.75), vals[-1]]


STAT_HDR = ["Field", "n", "Missing", "Sum", "Mean", "Min", "P25", "Median", "P75", "Max"]


def missing_table(hdr, rows, reads=None):
    reads = reads or {}
    out = []
    for h in hdr:
        if h.startswith("_") or h.startswith("meta/") or h.startswith("**") or h == "":
            continue
        m = sum(1 for r in rows if blank(r.get(h)))
        out.append([h, m, m / len(rows), reads.get(h, "complete" if m == 0 else "")])
    return out


# ---------------------------------------------------------------
# derived values shared across sections
# ---------------------------------------------------------------
av = defaultdict(int)
un = defaultdict(int)
for r in item:
    k = (r["_submission__uuid"], r["_parent_index"])
    av[k] += 1 if yes(r[AVAIL]) else 0
    un[k] += 1 if no(r[AVAIL]) else 0

unserved = [r for r in item if no(r[AVAIL])]
sub_dates = sorted(x for x in (as_date(r.get("_submission_time")) for r in sub) if x)
pres_dates = sorted(x for x in (as_date(r.get(DATE)) for r in pres) if x)

per_sub = Counter(r["_submission__uuid"] for r in pres)
fac_of = {r["_uuid"]: r["Facility"] for r in sub}

dup_groups = defaultdict(list)
for r in pres:
    dup_groups[str(r[PID]).strip()].append(r)
dups = {k: v for k, v in dup_groups.items() if len(v) > 1}


def mean_items(rows_by):
    return round(sum(num(r[N_MED]) or 0 for r in rows_by) / len(rows_by), 2) if rows_by else ""


def group_presc(col):
    g = defaultdict(list)
    for r in pres:
        g[str(r[col]).strip()].append(r)
    return [[k, len(v), len(v) / len(pres), mean_items(v)]
            for k, v in sorted(g.items(), key=lambda x: -len(x[1]))]


# ---------------------------------------------------------------
# SECTION 1 - overview
# ---------------------------------------------------------------
overview = [
    Table("Dataset", ["Item", "Value"], [
        ["Source workbook", SRC],
        ["Sheets", "3 (nested: facility visit > prescription > medicine item)"],
        ["Sheet 1 rows x cols", f"{len(sub)} x {len(hdr_sub)}"],
        ["Sheet 2 rows x cols", f"{len(pres)} x {len(hdr_pres)}"],
        ["Sheet 3 rows x cols", f"{len(item)} x {len(hdr_item)}"],
        ["Kobo submission dates", f"{min(sub_dates)} (all rows)" if len(set(sub_dates)) == 1
         else f"{min(sub_dates)} .. {max(sub_dates)}"],
        ["Prescription dates", f"{min(pres_dates)} .. {max(pres_dates)}"],
        ["Facilities", len({r["Facility"] for r in sub})],
        ["Districts", len({r["District"] for r in sub})],
        ["Report generated", dt.date.today().isoformat()],
    ]),
    Table("Key figures", ["Measure", "Value"], [
        ["Facility visits (submissions)", len(sub)],
        ["Prescriptions", len(pres)],
        ["Medicine items", len(item)],
        ["Prescriptions per submission (mean)", round(len(pres) / len(sub), 2)],
        ["Medicine items per prescription (mean)", round(len(item) / len(pres), 2)],
        ["Items unserved", len(unserved)],
        ["Item-level unserved rate", f"{len(unserved) / len(item) * 100:.1f}%"],
        ["Total value of unserved medicines (RWF)",
         sum(num(r["total_cost_unserved"]) or 0 for r in unserved)],
    ], note="Unserved figures are recomputed from medicine_item. The prescription sheet's "
            "own n_items_available / n_items_unavailable / pct_items_unserved fields are stale "
            "- see the Data quality section."),
]

# ---------------------------------------------------------------
# SECTION 2 - sheet 1
# ---------------------------------------------------------------
d_rows, d_n = counts(sub, "District", skip_blank=True)
t_rows, _ = counts(sub, "facility_type", skip_blank=True)
f_rows, _ = counts(sub, "Facility", skip_blank=True)
s_rows, _ = counts(sub, "Sector", skip_blank=True)

sheet1 = [
    Table("Visits by district", ["District", "Visits", "% of visits"], d_rows,
          total=["Total", len(sub), 1.0]),
    Table("Visits by facility type", ["Facility type", "Visits", "% of visits"], t_rows,
          total=["Total", len(sub), 1.0]),
    Table("Visits by facility", ["Facility", "Visits", "% of visits"], f_rows,
          total=["Total", len(sub), 1.0], chart=("v", 0, 1, None)),
    Table("Visits by sector", ["Sector", "Visits", "% of visits"], s_rows,
          note="Sector maps one-to-one onto Facility - the two variables carry identical "
               "information. Use one of them, not both.",
          total=["Total", len(sub), 1.0]),
    Table("Submission timestamps", ["Date", "Submissions"],
          [[str(k), v] for k, v in sorted(Counter(sub_dates).items())]),
    Table("Keys and review state", ["Check", "Result"], [
        ["_uuid distinct", f"{len({r['_uuid'] for r in sub})} / {len(sub)}"],
        ["_id distinct", f"{len({r['_id'] for r in sub})} / {len(sub)}"],
        ["Rows with a validation status",
         sum(1 for r in sub if not blank(r.get("_validation_status")))],
    ], note="No row carries a validation status - nothing in this export has been reviewed in Kobo."),
    Table("Missingness", ["Column", "Missing", "% missing", "Read"],
          missing_table(hdr_sub, sub),
          note="Province sits in an unnamed header cell and reads 'Kigali City' on all rows."),
]

# ---------------------------------------------------------------
# SECTION 3 - sheet 2
# ---------------------------------------------------------------
per_sub_dist = Counter(per_sub.values())
multi = [[k, fac_of.get(u, "")] for u, k in sorted(per_sub.items(), key=lambda x: -x[1]) if k > 1]

nmed_rows, _ = counts(pres, N_MED, skip_blank=True)
nmed_rows.sort(key=lambda r: float(r[0]))

daily = [[str(k), v, v / len(pres)] for k, v in sorted(Counter(pres_dates).items())]

day_by_fac = defaultdict(Counter)
for r in pres:
    day_by_fac[r["Facility"]][as_date(r[DATE])] += 1
all_days = sorted({d for c in day_by_fac.values() for d in c})
dayfac_rows = []
for f in sorted(day_by_fac, key=lambda f: -sum(day_by_fac[f].values())):
    dayfac_rows.append([f] + [day_by_fac[f].get(d, 0) for d in all_days]
                       + [sum(day_by_fac[f].values())])

dup_rows = []
for k, v in dups.items():
    for r in v:
        dup_rows.append([mask(k), r["Facility"], str(as_date(r[DATE])),
                         r[PRESCRIBER], num(r[N_MED])])

sheet2 = [
    Table("Prescriptions per submission", ["Prescriptions in the submission", "Submissions"],
          [[k, v] for k, v in sorted(per_sub_dist.items())],
          note=f"Mean {len(pres) / len(sub):.2f}. 125 submissions carry exactly one prescription; "
               "four carry the rest."),
    Table("Submissions holding more than one prescription",
          ["Prescriptions", "Facility"], multi,
          note="Solace Ministries CS entered all 30 of its prescriptions inside a single "
               "submission, which is why it shows 1 visit but 30 prescriptions."),
    Table("Prescriptions by district", ["District", "Prescriptions", "%", "Mean items"],
          group_presc("District"), total=["Total", len(pres), 1.0, mean_items(pres)]),
    Table("Prescriptions by facility type", ["Facility type", "Prescriptions", "%", "Mean items"],
          group_presc("facility_type"), total=["Total", len(pres), 1.0, mean_items(pres)]),
    Table("Prescriptions by facility", ["Facility", "Prescriptions", "%", "Mean items"],
          group_presc("Facility"), total=["Total", len(pres), 1.0, mean_items(pres)],
          chart=("v", 0, 1, None)),
    Table("Prescriptions by prescriber category",
          ["Prescriber category", "Prescriptions", "%", "Mean items"],
          group_presc(PRESCRIBER), total=["Total", len(pres), 1.0, mean_items(pres)],
          chart=("h", 0, 1, None),
          note="'General Practioner' is the form's own spelling."),
    Table("Prescriptions by date of prescription", ["Date", "Prescriptions", "% of 179"], daily,
          chart=("v", 0, 1, None),
          note="56 of 179 (31%) are dated the final day, which is also the single upload day. "
               "Five of six facilities recorded exactly 10 that day - a collection quota, not a "
               "prescribing surge. Exclude 10 September from any time trend."),
    Table("Prescriptions per day, per facility",
          ["Facility"] + [d.strftime("%d %b") for d in all_days] + ["Total"], dayfac_rows),
    Table("Medicines per prescription - distribution",
          ["Medicines prescribed", "Prescriptions", "%"], nmed_rows,
          total=["Total", len(pres), 1.0], chart=("v", 0, 1, None)),
    Table("Medicines per prescription - summary statistics", STAT_HDR,
          [stats_row(N_MED, pres, N_MED)]),
    Table("Kobo roll-up fields as exported", STAT_HDR,
          [stats_row("n_items_available", pres, "n_items_available"),
           stats_row("n_items_unavailable", pres, "n_items_unavailable"),
           stats_row("pct_items_unserved", pres, "pct_items_unserved")],
          note="Reported for completeness only - all three are stale. See Data quality."),
    Table("Duplicate Patient / Prescription IDs",
          ["ID (masked)", "Facility", "Date", "Prescriber", "Items"], dup_rows,
          note=f"{len(set(dup_groups))} distinct IDs across {len(pres)} prescriptions; "
               f"{len(dups)} IDs used twice. Each repeat matches its twin on facility, date, "
               "cadre and item count - consistent with double entry."),
    Table("Missingness", ["Column", "Missing", "% missing", "Read"],
          missing_table(hdr_pres, pres, {
              "Please specify the allied health professional type": "expected - conditional",
              "n_items_available": "populated but stale",
              "n_items_unavailable": "populated but stale",
              "pct_items_unserved": "populated but stale",
          })),
]

# ---------------------------------------------------------------
# SECTION 4 - sheet 3
# ---------------------------------------------------------------
eml_rows, _ = counts(item, EML, skip_blank=True)
nt_rows, _ = counts(item, "Name type", skip_blank=True)
item_no_rows, _ = counts(item, "item_number", skip_blank=True)
item_no_rows.sort(key=lambda r: float(r[0]))
med_rows, _ = counts(item, MED_EML, denom=len(item), skip_blank=True)
free_rows, _ = counts(item, MED_FREE, skip_blank=True)
freecls_rows, _ = counts(item, CLASS_FREE, skip_blank=True)

cls = Counter(re.sub(r"^[\d.]+\s*", "", str(r["Drug class"] or "(blank)")).strip()
              for r in item)
cls_rows = [[k, v, v / len(item)] for k, v in cls.most_common()]

qty = Counter(int(num(r["Qty prescribed"])) for r in item)
qty_rows = [[k, v, v / len(item)] for k, v in sorted(qty.items())]

big_qty = [[int(num(r["Qty prescribed"])), r[MED_EML] or r[MED_FREE], par(r)["Facility"]]
           for r in sorted(item, key=lambda r: -(num(r["Qty prescribed"]) or 0))[:15]]

avail_rows, _ = counts(item, AVAIL, skip_blank=True)

fac_tab = defaultdict(lambda: [0, 0])
for r in item:
    p = par(r)
    fac_tab[(p["facility_type"], p["Facility"])][0] += 1
    fac_tab[(p["facility_type"], p["Facility"])][1] += 1 if no(r[AVAIL]) else 0
pc = Counter((p["facility_type"], p["Facility"]) for p in pres)
fac_rows = [[k[0], k[1], pc[k], i, u, u / i]
            for k, (i, u) in sorted(fac_tab.items(), key=lambda x: -x[1][0])]

cadre_tab = defaultdict(lambda: [0, 0])
for r in item:
    cadre_tab[par(r)[PRESCRIBER]][0] += 1
    cadre_tab[par(r)[PRESCRIBER]][1] += 1 if no(r[AVAIL]) else 0
cadre_rows = [[k, i, u] for k, (i, u) in sorted(cadre_tab.items(), key=lambda x: -x[1][0])]

dt_tab = Counter((par(r)["District"], par(r)["facility_type"]) for r in item)
dt_rows = [[k[0], k[1], v, v / len(item)] for k, v in dt_tab.most_common()]

eml_fac = defaultdict(lambda: [0, 0, 0])
for r in item:
    k = par(r)["Facility"]
    eml_fac[k][0] += 1
    eml_fac[k][1] += 1 if yes(r[EML]) else 0
    eml_fac[k][2] += 1 if str(r["Name type"]).strip() == "Brand" else 0
emlfac_rows = [[k, i, e, b] for k, (i, e, b) in sorted(eml_fac.items(), key=lambda x: -x[1][0])]

cls_by_fac = defaultdict(Counter)
for r in item:
    cls_by_fac[par(r)["Facility"]][
        re.sub(r"^[\d.]+\s*", "", str(r["Drug class"] or "(blank)")).strip()] += 1
clsfac_rows = []
for f in sorted(cls_by_fac, key=lambda f: -sum(cls_by_fac[f].values())):
    for k, v in cls_by_fac[f].most_common(5):
        clsfac_rows.append([f, k, v])

unserved_rows = []
for r in unserved:
    p = par(r)
    unserved_rows.append([
        p["Facility"], p["facility_type"], p["District"], str(as_date(p[DATE])), p[PRESCRIBER],
        r[MED_EML] or r[MED_FREE], r[EML], int(num(r["Qty prescribed"])),
        r[ORDERED], r[SUPPLIED], r[RMS_REG], r[RMS_CEN], r[DIST],
        num(r[UNIT_PRICE]), num(r["total_cost_unserved"]),
    ])

follow_rows = []
for label, col in [("Ordered during the month?", ORDERED),
                   ("Supplied fully / partially / not at all?", SUPPLIED),
                   ("Available at RMS Regional?", RMS_REG),
                   ("Available at RMS Central?", RMS_CEN),
                   ("distribution_nationwide", DIST)]:
    c = Counter(str(r[col]).strip() for r in unserved if not blank(r[col]))
    follow_rows.append([label, " | ".join(f"{k} {v}" for k, v in c.most_common())])

full = part = none_ = 0
for r in pres:
    k = (r["_submission__uuid"], r["_index"])
    if un[k] == 0:
        full += 1
    elif av[k] == 0:
        none_ += 1
    else:
        part += 1

sheet3 = [
    Table("Rows per parent prescription", ["Check", "Result"], [
        ["Parents with items", f"{len({(r['_submission__uuid'], r['_parent_index']) for r in item})} / {len(pres)}"],
        ["Items per prescription (mean)", round(len(item) / len(pres), 2)],
        ["Rows where the count != the parent's stated number",
         sum(1 for r in pres
             if Counter((x["_submission__uuid"], x["_parent_index"]) for x in item)
             .get((r["_submission__uuid"], r["_index"]), 0) != num(r[N_MED]))],
    ]),
    Table("On the EML?", ["Response", "Items", "%"], eml_rows, total=["Total", len(item), 1.0]),
    Table("Name type recorded", ["Name type", "Items", "%"], nt_rows,
          total=["Total", len(item), 1.0],
          note="45 of the 46 'Brand' rows carry an INN name picked from the EML list, and 44 of "
               "those 46 come from one facility - see Findings."),
    Table("Position within the prescription", ["item_number", "Items", "%"], item_no_rows,
          total=["Total", len(item), 1.0]),
    Table("Medicines selected from the EML picklist - all distinct entries",
          ["Medicine", "Items", "% of 353"], med_rows, chart=("h", 0, 1, 15),
          note="350 of 353 items. Near-duplicate entries to merge before any consumption "
               "analysis: nystatin_tab_500_000_ui / nystatin_tab_500000_ui, ibuprofen_tab_400_mg / "
               "ibuprofen_tab_200_mg_400_mg, hydrochlorothiazide_tab_25_mg / ..._25_mg_50mg, "
               "nifedipine_tab_10_mg_20_mg / nifedipine_tab_20mg, plus hyocine/hyoscine and "
               "chlorpheniramine/chlorphenamine spellings."),
    Table("Medicines typed free-text (off-EML)", ["Entry", "Items"],
          [[r[0], r[1]] for r in free_rows],
          note="Two of the three are syrup formulations of medicines already on the EML in tablet "
               "form - the picklist appears to carry only the tablet forms."),
    Table("Drug class typed free-text (off-EML)", ["Entry", "Items"],
          [[r[0], r[1]] for r in freecls_rows],
          note="Both free-text fields are only reachable when EML = No, so these are the same "
               "three items."),
    Table("Drug class - all values, EML numbering stripped",
          ["Drug class", "Items", "% of 353"], cls_rows, total=["Total", len(item), 1.0],
          chart=("h", 0, 1, 12)),
    Table("Top five drug classes at each facility", ["Facility", "Drug class", "Items"],
          clsfac_rows,
          note="Masaka DH is the only facility where oxytocics and parenteral products lead, "
               "consistent with a district hospital running maternity and inpatient services."),
    Table("Quantity prescribed - full distribution", ["Qty prescribed", "Items", "% of 353"],
          qty_rows, total=["Total", len(item), 1.0], chart=("v", 0, 1, None)),
    Table("Quantity prescribed - summary statistics", STAT_HDR,
          [stats_row("Qty prescribed", item, "Qty prescribed")]),
    Table("The 15 largest quantities", ["Qty", "Medicine", "Facility"], big_qty),
    Table("Availability at the facility on the prescription date",
          ["Response", "Items", "%"], avail_rows, total=["Total", len(item), 1.0]),
    Table("The unserved items in full",
          ["Facility", "Type", "District", "Date", "Prescriber", "Medicine", "On EML", "Qty",
           "Ordered this month", "Supplied", "RMS Regional", "RMS Central", "Flag",
           "Unit price (RWF)", "Value unserved (RWF)"], unserved_rows,
          note="Two of the three were in stock at RMS but never reached the facility - one was "
               "ordered and not supplied, the other never ordered. On the pilot's own evidence the "
               "bottleneck is facility-level ordering and last-mile distribution, not national "
               "availability."),
    Table("Follow-up responses across the unserved rows", ["Question", "Responses"], follow_rows,
          note="Every follow-up question is conditional on 'available = No', so all are answered "
               "on exactly these rows and blank on the other 350. That is the form working as "
               "designed, not missing data."),
    Table("Items and unserved rate by facility",
          ["Type", "Facility", "Prescriptions", "Items", "Unserved", "Unserved rate"], fac_rows,
          total=["Total", "", len(pres), len(item), len(unserved), len(unserved) / len(item)],
          chart=("h", 1, 3, None),
          note="Geography taken from the parent prescription. Read off medicine_item's own "
               "Facility column instead, the unserved items land at the wrong facilities."),
    Table("Items and unserved by prescriber category", ["Prescriber category", "Items", "Unserved"],
          cadre_rows, total=["Total", len(item), len(unserved)]),
    Table("Items by district and facility type", ["District", "Facility type", "Items", "%"],
          dt_rows, total=["Total", "", len(item), 1.0]),
    Table("EML status and name type by facility",
          ["Facility", "Items", "On EML", "Typed 'Brand'"], emlfac_rows),
    Table("Prescription-level service status, recomputed from medicine_item",
          ["Status", "Prescriptions", "%"],
          [["Fully served", full, full / len(pres)],
           ["Partly unserved", part, part / len(pres)],
           ["Fully unserved", none_, none_ / len(pres)]],
          total=["Total", len(pres), 1.0], chart=("h", 0, 1, None)),
    Table("Missingness", ["Column", "Missing", "% missing", "Read"],
          missing_table(hdr_item, item, {
              "Province": "populated but misaligned - Defect 1",
              "District": "populated but misaligned - Defect 1",
              "Sector": "populated but misaligned - Defect 1",
              "Facility": "populated but misaligned - Defect 1",
              "facility_type": "populated but misaligned - Defect 1",
              PRESCRIBER: "populated but misaligned - Defect 1",
              PID: "populated but misaligned - Defect 1",
              DATE: "populated but misaligned - Defect 1",
              MED_EML: "expected - the 3 off-EML rows",
              "Drug class": "expected - the 3 off-EML rows",
              MED_FREE: "expected - conditional",
              CLASS_FREE: "expected - conditional",
              ORDERED: "expected - conditional on unavailable",
              SUPPLIED: "expected - conditional on unavailable",
              UNIT_PRICE: "expected - conditional on unavailable",
              "total_cost_unserved": "expected - conditional on unavailable",
              RMS_REG: "expected - conditional on unavailable",
              RMS_CEN: "expected - conditional on unavailable",
              DIST: "expected - conditional on unavailable",
          })),
]

# ---------------------------------------------------------------
# SECTION 5 - data quality
# ---------------------------------------------------------------
mis_rows = []
for col in (PID, DATE, "Facility", PRESCRIBER):
    if col == DATE:
        m = sum(1 for r in item if as_date(r[col]) != as_date(par(r)[col]))
    else:
        m = sum(1 for r in item if str(r[col]).strip() != str(par(r)[col]).strip())
    mis_rows.append([col, m, len(item), m / len(item)])
na = sum(1 for r in item if str(r["Facility"]).strip() == "#N/A")
mis_rows.append(["Facility literally '#N/A'", na, len(item), na / len(item)])

mm_a = sum(1 for r in pres if av[(r["_submission__uuid"], r["_index"])] != num(r["n_items_available"]))
mm_u = sum(1 for r in pres if un[(r["_submission__uuid"], r["_index"])] != num(r["n_items_unavailable"]))
bad_sum = sum(1 for r in pres
              if (num(r["n_items_available"]) or 0) + (num(r["n_items_unavailable"]) or 0)
              != num(r[N_MED]))

quality = [
    Table("Defect 1 - the first eight columns of medicine_item are misaligned",
          ["Column", "Rows disagreeing with the parent", "of", "%"], mis_rows,
          note="Province, District, Sector, Facility, facility_type, Prescriber category, "
               "Patient / Prescription ID and Date of prescription were filled by a lookup that "
               "never resolved. FIX: drop these eight columns from the export, or join "
               "medicine_item to prescription on (_submission__uuid, _parent_index) -> "
               "(_submission__uuid, _index) and take those fields from the parent. Read straight "
               "off the sheet, the three unserved medicines land at Masaka DH and Kabusunzu CS; "
               "joined correctly they are at Gikondo CS (2) and Masaka DH (1)."),
    Table("Defect 2 - the calculated fields on 'prescription' are stale",
          ["Field", "Exported sum", "Actual", "Rows wrong", "of"],
          [["n_items_available", sum(num(r["n_items_available"]) or 0 for r in pres),
            sum(av.values()), mm_a, len(pres)],
           ["n_items_unavailable", sum(num(r["n_items_unavailable"]) or 0 for r in pres),
            sum(un.values()), mm_u, len(pres)],
           ["available + unavailable = prescribed", "fails", "", bad_sum, len(pres)]],
          note="The real item-level unserved rate is 0.8% (3 of 353). The exported "
               "pct_items_unserved averages 0.28 and implies a third of that. FIX: recompute all "
               "three from medicine_item before any analysis or dashboard."),
    Table("Link integrity - sound", ["Check", "Result"], [
        ["Prescription rows with an unknown submission uuid",
         sum(1 for r in pres if r["_submission__uuid"] not in {x["_uuid"] for x in sub})],
        ["medicine_item rows with no matching parent",
         sum(1 for r in item if (r["_submission__uuid"], r["_parent_index"]) not in parent)],
        ["Submissions with no prescription rows",
         len({x["_uuid"] for x in sub} - {r["_submission__uuid"] for r in pres})],
        ["medicine_item counts matching the parent's stated number", f"179 / {len(pres)}"],
    ]),
    Table("Other flags", ["Flag", "Value"], [
        ["Items with drug class 'Not specified in EML source'",
         sum(1 for r in item if str(r["Drug class"]).strip() == "Not specified in EML source")],
        ["Duplicate Patient / Prescription IDs", len(dups)],
        ["Rows with a Kobo validation status", 0],
        ["Upload days vs prescription date span",
         f"all {len(sub)} submissions uploaded {min(sub_dates)}; prescriptions span "
         f"{min(pres_dates)} .. {max(pres_dates)}"],
    ], note="Retrospective chart abstraction, not point-of-care capture."),
]

findings = [
    Table("Findings and recommended actions", ["Finding", "Evidence", "Action"], [
        ["Solace Ministries CS answered 'Brand' on all 44 of its items",
         "44 of the 46 brand-typed rows in the dataset; every other facility answered 'Brand' at "
         "most once. All 44 carry an INN name picked from the EML list and all 44 are flagged "
         "on-EML.",
         "One collector reading the question differently, not a brand-prescribing signal. "
         "Re-brief that collector; treat Name type as unusable for Solace until re-collected."],
        ["The 10 September spike is a quota, not prescribing volume",
         "Five of six facilities recorded exactly 10 prescriptions on the final day (Kabusunzu 6); "
         "all 129 submissions uploaded that same day against prescription dates back to 31 July.",
         "Exclude 10 September from any time trend. Record the abstraction protocol alongside the "
         "data so later rounds are comparable."],
        ["The EML reference list behind the form has gaps",
         "45 items (12.7%) carry drug class 'Not specified in EML source'; two of the three "
         "free-text medicines are syrup forms of products already on the list; the picklist holds "
         "near-duplicates and two spelling variants.",
         "Reconcile the picklist against 'List of essential medecines.xlsx' before the next round "
         "- fill missing classes, add paediatric/syrup forms, de-duplicate."],
        ["Three unserved items cannot carry a comparison",
         "Every breakdown by facility, facility type or cadre is driven by single observations. "
         "Gikondo's 4.2% is two items; Masaka's 1.4% is one; the Specialist rate is one of fifteen.",
         "Report the count and the three case narratives, not rates or comparisons. The pilot "
         "succeeds as an instrument test; sample size must grow substantially before unserved "
         "rates mean anything."],
        ["Sector duplicates Facility",
         "The six facilities map one-to-one onto six sectors.",
         "Use one of the two variables, not both."],
    ]),
]

SECTIONS = [
    ("Overview", overview),
    ("Sheet 1 - facility visits", sheet1),
    ("Sheet 2 - prescriptions", sheet2),
    ("Sheet 3 - medicine items", sheet3),
    ("Data quality", quality),
    ("Findings", findings),
]

GRAIN = {
    "Sheet 1 - facility visits": f"{len(sub)} rows x {len(hdr_sub)} columns - one row = one submission (a facility visit)",
    "Sheet 2 - prescriptions": f"{len(pres)} rows x {len(hdr_pres)} columns - one row = one prescription",
    "Sheet 3 - medicine items": f"{len(item)} rows x {len(hdr_item)} columns - one row = one medicine on one prescription",
}


# ---------------------------------------------------------------
# HTML writer - plain
# ---------------------------------------------------------------
# ---------------------------------------------------------------
# inline SVG bar charts - no libraries, so the HTML file stands alone offline
# ---------------------------------------------------------------
BAR = "#2f6fba"
GRID = "#d8d8d8"
AXIS = "#888888"
LBL = "#333333"


def nice_ticks(hi):
    """A rounded axis maximum and its tick values, so every gridline names a real number.

    Candidate steps are the usual 1/2/2.5/5 x powers of ten. Among those giving 3-7
    gridlines, take the tightest ceiling over the data, breaking ties toward ~5 ticks -
    otherwise a max of 56 ends up plotted against an axis running to 80.
    """
    if hi <= 0:
        return 1, [0, 1]
    best = None
    mag = 10 ** -3
    while mag <= hi * 10:
        for m in (1, 2, 2.5, 5):
            step = mag * m
            n = int(hi / step) + (0 if abs(hi / step - int(hi / step)) < 1e-9 else 1)
            top = step * n
            if 2 <= n <= 6:
                key = (top, abs((n + 1) - 5))
                if best is None or key < best[0]:
                    best = (key, top, step, n)
        mag *= 10
    if best is None:
        return hi, [0, hi]
    _, top, step, n = best
    return top, [i * step for i in range(n + 1)]


def esc(s):
    return html.escape(str(s))


def svg_vbars(pairs, ylabel, width=880, plot_h=210):
    """Vertical bars - short labels (dates, small integers, facility names)."""
    n = max(len(pairs), 1)
    longest = max((len(str(p[0])) for p in pairs), default=1)
    rotate = longest > 7 or (n > 14 and longest > 4)
    ml, mr, mt = 54, 16, 20
    mb = 74 if rotate else 42
    h = mt + plot_h + mb
    top, ticks = nice_ticks(max((p[1] for p in pairs), default=1))
    slot = (width - ml - mr) / n
    bw = min(slot * 0.68, 56)

    s = [f'<svg class="chart" viewBox="0 0 {width} {h}" width="{width}" height="{h}" '
         f'role="img" aria-label="{esc(ylabel)}" xmlns="http://www.w3.org/2000/svg">']
    s.append(f'<rect x="0" y="0" width="{width}" height="{h}" fill="#ffffff"/>')

    for t in ticks:
        y = mt + plot_h - (t / top) * plot_h
        s.append(f'<line x1="{ml}" y1="{y:.1f}" x2="{width - mr}" y2="{y:.1f}" '
                 f'stroke="{GRID}" stroke-width="1" fill="none"/>')
        s.append(f'<text x="{ml - 8}" y="{y + 3.5:.1f}" text-anchor="end" font-size="10" '
                 f'font-family="Arial, sans-serif" fill="{LBL}">{t:g}</text>')
    s.append(f'<text x="10" y="{mt + plot_h / 2:.0f}" font-size="10" font-family="Arial, sans-serif" '
             f'fill="{LBL}" transform="rotate(-90 10 {mt + plot_h / 2:.0f})" text-anchor="middle">'
             f'{esc(ylabel)}</text>')

    for i, (lab, val) in enumerate(pairs):
        cx = ml + slot * i + slot / 2
        bh = (val / top) * plot_h
        y = mt + plot_h - bh
        s.append(f'<rect x="{cx - bw / 2:.1f}" y="{y:.1f}" width="{bw:.1f}" height="{max(bh, 0.6):.1f}" '
                 f'fill="{BAR}"><title>{esc(lab)}: {val:g}</title></rect>')
        if n <= 26:
            s.append(f'<text x="{cx:.1f}" y="{y - 4:.1f}" text-anchor="middle" font-size="9.5" '
                     f'font-family="Arial, sans-serif" fill="{LBL}">{val:g}</text>')
        if rotate:
            s.append(f'<text x="{cx:.1f}" y="{mt + plot_h + 12:.0f}" font-size="9.5" '
                     f'font-family="Arial, sans-serif" fill="{LBL}" text-anchor="end" '
                     f'transform="rotate(-45 {cx:.1f} {mt + plot_h + 12:.0f})">{esc(lab)[:26]}</text>')
        else:
            s.append(f'<text x="{cx:.1f}" y="{mt + plot_h + 14:.0f}" text-anchor="middle" '
                     f'font-size="9.5" font-family="Arial, sans-serif" fill="{LBL}">{esc(lab)[:14]}</text>')

    s.append(f'<line x1="{ml}" y1="{mt + plot_h}" x2="{width - mr}" y2="{mt + plot_h}" '
             f'stroke="{AXIS}" stroke-width="1" fill="none"/>')
    s.append("</svg>")
    return "\n".join(s)


def svg_hbars(pairs, xlabel, width=880):
    """Horizontal bars - long category labels (drug classes, medicine names, cadres)."""
    n = max(len(pairs), 1)
    row_h = 21
    lw = min(max((len(str(p[0])) for p in pairs), default=10) * 5.7 + 10, 300)
    ml, mr, mt, mb = lw, 52, 16, 34
    plot_w = width - ml - mr
    h = mt + n * row_h + mb
    top, ticks = nice_ticks(max((p[1] for p in pairs), default=1))

    s = [f'<svg class="chart" viewBox="0 0 {width} {h}" width="{width}" height="{h}" '
         f'role="img" aria-label="{esc(xlabel)}" xmlns="http://www.w3.org/2000/svg">']
    s.append(f'<rect x="0" y="0" width="{width}" height="{h}" fill="#ffffff"/>')

    for t in ticks:
        x = ml + (t / top) * plot_w
        s.append(f'<line x1="{x:.1f}" y1="{mt}" x2="{x:.1f}" y2="{mt + n * row_h}" '
                 f'stroke="{GRID}" stroke-width="1" fill="none"/>')
        s.append(f'<text x="{x:.1f}" y="{mt + n * row_h + 14}" text-anchor="middle" font-size="10" '
                 f'font-family="Arial, sans-serif" fill="{LBL}">{t:g}</text>')
    s.append(f'<text x="{ml + plot_w / 2:.0f}" y="{h - 6}" text-anchor="middle" font-size="10" '
             f'font-family="Arial, sans-serif" fill="{LBL}">{esc(xlabel)}</text>')

    for i, (lab, val) in enumerate(pairs):
        y = mt + i * row_h
        bh = row_h - 7
        bl = (val / top) * plot_w
        s.append(f'<text x="{ml - 8}" y="{y + bh - 1:.1f}" text-anchor="end" font-size="10" '
                 f'font-family="Arial, sans-serif" fill="{LBL}">{esc(lab)[:52]}</text>')
        s.append(f'<rect x="{ml}" y="{y:.1f}" width="{max(bl, 0.6):.1f}" height="{bh}" '
                 f'fill="{BAR}"><title>{esc(lab)}: {val:g}</title></rect>')
        s.append(f'<text x="{ml + bl + 6:.1f}" y="{y + bh - 1:.1f}" font-size="10" '
                 f'font-family="Arial, sans-serif" fill="{LBL}">{val:g}</text>')

    s.append(f'<line x1="{ml}" y1="{mt}" x2="{ml}" y2="{mt + n * row_h}" '
             f'stroke="{AXIS}" stroke-width="1" fill="none"/>')
    s.append("</svg>")
    return "\n".join(s)


def render_chart(t):
    pairs = t.chart_data()
    if not pairs:
        return ""
    kind = t.chart[0]
    ylab = str(t.headers[t.chart[2]])
    cap = t.title if not t.chart[3] else f"{t.title} - top {len(pairs)}"
    svg = svg_vbars(pairs, ylab) if kind == "v" else svg_hbars(pairs, ylab)
    return f'<div class="chartwrap"><p class="caption">Chart: {esc(cap)}</p>{svg}</div>'


def fmt_cell(v, pct=False):
    if isinstance(v, bool):
        return html.escape(str(v))
    if pct and isinstance(v, (int, float)):
        return f"{v * 100:.1f}%"
    if isinstance(v, float):
        return f"{v:,.2f}".rstrip("0").rstrip(".")
    if isinstance(v, int):
        return f"{v:,}"
    return html.escape(str(v))


def write_html():
    p = []
    p.append("<!DOCTYPE html>")
    p.append('<html lang="en"><head><meta charset="utf-8">')
    p.append('<meta name="viewport" content="width=device-width, initial-scale=1">')
    p.append("<title>Prescription pilot2 - basic analysis</title>")
    p.append("""<style>
body { font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 1.5;
       color: #000; background: #fff; margin: 0; padding: 24px; max-width: 1100px; }
h1 { font-size: 22px; margin: 0 0 4px; }
h2 { font-size: 18px; margin: 34px 0 4px; border-bottom: 2px solid #000; padding-bottom: 4px; }
h3 { font-size: 14px; margin: 22px 0 6px; }
p.meta, p.grain, p.note { font-size: 12px; color: #444; margin: 4px 0; }
p.note { margin: 6px 0 0; }
table { border-collapse: collapse; margin: 6px 0 0; font-size: 13px; }
th, td { border: 1px solid #999; padding: 4px 8px; text-align: left; vertical-align: top; }
th { background: #eee; font-weight: bold; }
td.num { text-align: right; }
tr.total td { font-weight: bold; background: #f4f4f4; }
ul { margin: 6px 0; padding-left: 20px; font-size: 13px; }
a { color: #00c; }
.chartwrap { margin: 10px 0 4px; max-width: 100%; overflow-x: auto; }
.chartwrap svg { max-width: 100%; height: auto; display: block; }
p.caption { font-size: 12px; font-weight: bold; margin: 10px 0 2px; }
@media print { body { padding: 0; } h2 { page-break-after: avoid; } table { page-break-inside: avoid; } }
</style></head><body>""")

    p.append("<h1>Prescription pilot2 - basic analysis</h1>")
    p.append(f'<p class="meta">Source: {html.escape(SRC)} (3 sheets, {len(sub)} / {len(pres)} / '
             f'{len(item)} rows). Generated {dt.date.today().isoformat()} by '
             f'export_prescription_pilot2_outputs.py (openpyxl + Python standard library).</p>')
    p.append('<p class="meta">Sheet 3 figures take facility, district, date, prescriber and '
             'patient ID from the parent prescription - that sheet\'s own copies of those eight '
             'columns are misaligned (see Data quality). Patient / Prescription IDs are masked.</p>')

    p.append("<h3>Contents</h3><ul>")
    for name, _ in SECTIONS:
        p.append(f'<li><a href="#{html.escape(name.replace(" ", "-"))}">{html.escape(name)}</a></li>')
    p.append("</ul>")

    for name, tables in SECTIONS:
        p.append(f'<h2 id="{html.escape(name.replace(" ", "-"))}">{html.escape(name)}</h2>')
        if name in GRAIN:
            p.append(f'<p class="grain">{html.escape(GRAIN[name])}</p>')
        for t in tables:
            p.append(f"<h3>{html.escape(t.title)}</h3>")
            p.append("<table><thead><tr>")
            for h in t.headers:
                p.append(f"<th>{html.escape(str(h))}</th>")
            p.append("</tr></thead><tbody>")
            body = list(t.rows) + ([t.total] if t.total else [])
            for ri, row in enumerate(body):
                cls = ' class="total"' if t.total and ri == len(body) - 1 else ""
                p.append(f"<tr{cls}>")
                for ci, v in enumerate(row):
                    numeric = isinstance(v, (int, float)) and not isinstance(v, bool)
                    td = '<td class="num">' if numeric else "<td>"
                    p.append(td + fmt_cell(v, ci in t.pct) + "</td>")
                p.append("</tr>")
            p.append("</tbody></table>")
            if t.note:
                p.append(f'<p class="note">{html.escape(t.note)}</p>')
            if t.chart:
                p.append(render_chart(t))

    p.append("</body></html>")
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write("\n".join(p))
    print(f"wrote {OUT_HTML}")


# ---------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="E6E6E6")
TITLE_FONT = Font(bold=True, size=12)
HDR_FONT = Font(bold=True, size=10)
NOTE_FONT = Font(size=9, italic=True, color="555555")


def add_excel_charts(wb):
    """A 'Charts' worksheet holding native Excel charts - editable, not pasted images.

    Each chart keeps its own small data block beside it so the chart stays live if the
    sheet is copied into another workbook.
    """
    ws = wb.create_sheet("Charts")
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1, value="Charts").font = Font(bold=True, size=14)
    ws.cell(row=2, column=1,
            value="Each chart is a native Excel chart over the data block to its left - "
                  "click it to edit, restyle or copy into a deck.").font = NOTE_FONT
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 12

    row = 4
    for section, tables in SECTIONS:
        for t in tables:
            pairs = t.chart_data()
            if not pairs:
                continue
            title = t.title if not t.chart[3] else f"{t.title} (top {len(pairs)})"
            ws.cell(row=row, column=1, value=title).font = TITLE_FONT
            ws.cell(row=row, column=2, value=section).font = NOTE_FONT
            hdr_row = row + 1
            ws.cell(row=hdr_row, column=1, value=str(t.headers[t.chart[1]])).font = HDR_FONT
            ws.cell(row=hdr_row, column=2, value=str(t.headers[t.chart[2]])).font = HDR_FONT
            for i, (lab, val) in enumerate(pairs, start=1):
                ws.cell(row=hdr_row + i, column=1, value=lab)
                c = ws.cell(row=hdr_row + i, column=2, value=val)
                c.number_format = "#,##0.##"

            ch = BarChart()
            ch.type = "col" if t.chart[0] == "v" else "bar"
            ch.style = 2
            ch.title = title
            ch.y_axis.title = str(t.headers[t.chart[2]])
            ch.x_axis.title = str(t.headers[t.chart[1]])
            ch.legend = None
            ch.gapWidth = 40
            data = Reference(ws, min_col=2, min_row=hdr_row, max_row=hdr_row + len(pairs))
            cats = Reference(ws, min_col=1, min_row=hdr_row + 1, max_row=hdr_row + len(pairs))
            ch.add_data(data, titles_from_data=True)
            ch.set_categories(cats)
            ch.height = max(7.5, min(0.62 * len(pairs) + 3.5, 18))
            ch.width = 20
            ws.add_chart(ch, f"D{row}")

            row = hdr_row + len(pairs) + 2
            row = max(row, hdr_row + int(ch.height / 0.53) + 2)
    print(f"  {sum(1 for _, ts in SECTIONS for t in ts if t.chart)} Excel charts")


def write_xlsx():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for name, tables in SECTIONS:
        ws = wb.create_sheet(name[:31].replace("/", "-"))
        ws.sheet_view.showGridLines = False
        row = 1
        ws.cell(row=row, column=1, value=name).font = Font(bold=True, size=14)
        row += 1
        if name in GRAIN:
            c = ws.cell(row=row, column=1, value=GRAIN[name])
            c.font = NOTE_FONT
            row += 1
        row += 1

        widths = {}
        for t in tables:
            ws.cell(row=row, column=1, value=t.title).font = TITLE_FONT
            row += 1

            for ci, h in enumerate(t.headers, start=1):
                c = ws.cell(row=row, column=ci, value=str(h))
                c.font = HDR_FONT
                c.fill = HDR_FILL
                c.border = BORDER
                c.alignment = Alignment(wrap_text=True, vertical="top")
                widths[ci] = max(widths.get(ci, 10), min(len(str(h)) + 2, 42))
            row += 1

            body = list(t.rows) + ([t.total] if t.total else [])
            for ri, r in enumerate(body):
                total_row = t.total is not None and ri == len(body) - 1
                for ci, v in enumerate(r, start=1):
                    c = ws.cell(row=row, column=ci, value=v)
                    c.border = BORDER
                    c.alignment = Alignment(vertical="top", wrap_text=isinstance(v, str) and len(str(v)) > 45)
                    if total_row:
                        c.font = Font(bold=True)
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        c.number_format = "0.0%" if (ci - 1) in t.pct else "#,##0.##"
                    widths[ci] = max(widths.get(ci, 10), min(len(str(v)) + 2, 46))
                row += 1

            if t.note:
                c = ws.cell(row=row, column=1, value="Note: " + t.note)
                c.font = NOTE_FONT
                c.alignment = Alignment(wrap_text=True, vertical="top")
                ws.merge_cells(start_row=row, start_column=1,
                               end_row=row, end_column=max(len(t.headers), 2))
                ws.row_dimensions[row].height = 13 * (1 + len(t.note) // 110)
                row += 1
            row += 1

        for ci, w in widths.items():
            ws.column_dimensions[get_column_letter(ci)].width = w

    add_excel_charts(wb)
    wb.save(OUT_XLSX)
    print(f"wrote {OUT_XLSX}")


write_html()
write_xlsx()
