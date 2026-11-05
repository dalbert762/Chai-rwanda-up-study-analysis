"""
Basic analysis of 'Prescription pilot2.xlsx' (Unserved Prescriptions - prescription
assessment pilot, KoboToolbox export).

Three sheets, one per level of the form's nesting:
- 'Prescription assessment tool...' : one row per submission (a facility visit)
- 'prescription'                    : one row per prescription (repeat group)
- 'medicine_item'                   : one row per medicine on a prescription (nested repeat)

    python3 analyze_prescription_pilot2.py

Uses openpyxl + the standard library only (no pandas), so it runs against the system
python3 on this machine.

IMPORTANT - medicine_item's first 8 columns (Province..Prescriber category, Patient /
Prescription ID, Date of prescription) do NOT line up with the rows they sit next to:
they still carry '#N/A' from a lookup and disagree with the parent prescription on
~95% of rows. Everything below therefore takes facility / district / date / prescriber /
patient ID for a medicine row from its PARENT prescription, joined on
(_submission__uuid, _parent_index) -> (_submission__uuid, _index).
"""

import datetime as dt
import re
from collections import Counter, defaultdict

import openpyxl

FILE = "Prescription pilot2.xlsx"

S_SUB = "Prescription assessment tool..."
S_PRES = "prescription"
S_ITEM = "medicine_item"

# Long Kobo column labels, named once.
N_MED = "Number of medicines prescribed"
PID = "Patient / Prescription ID"
DATE = "Date of prescription"
PRESCRIBER = "Prescriber category"
EML = "Is this medicine on the EML (Essential Medicines List)?"
MED_EML = "Medicine name (type the first letter and select)"
MED_FREE = "Medicine name (not on EML list) - please specify"
AVAIL = "Was the medicine available at the health facility during the date of prescription?"
ORDERED = "Was the medicine ordered during the month the prescription was made?"
SUPPLIED = "Was it supplied fully, partially, or not at all?"
UNIT_PRICE = "Unit price for the medicine that was not available (RWF)"
RMS_REG = "Was the medicine available at RMS Regional during the date of prescription?"
RMS_CEN = "Was the medicine available at RMS Central during the date of prescription?"


# ---------------------------------------------------------------
# helpers
# ---------------------------------------------------------------
def load(sheet):
    wb = openpyxl.load_workbook(FILE, data_only=True, read_only=True)
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
    """Kobo dates arrive as datetimes on 'prescription' and as Excel serials on 'medicine_item'."""
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


def freq(rows, col, top=None, show_blank=True):
    c = Counter()
    for r in rows:
        v = r.get(col)
        c["(blank)" if blank(v) else str(v).strip()] += 1
    if not show_blank:
        c.pop("(blank)", None)
    n = sum(c.values()) or 1
    for k, v in c.most_common(top):
        print(f"    {k[:58]:<58} {v:>5}  {v / n * 100:5.1f}%")
    if top and len(c) > top:
        print(f"    ... {len(c) - top} more distinct values")


def desc(rows, col):
    vals = sorted(v for v in (num(r.get(col)) for r in rows) if v is not None)
    miss = len(rows) - len(vals)
    if not vals:
        print(f"    {col}: no numeric values ({miss} missing)")
        return
    n = len(vals)

    def q(p):
        i = p * (n - 1)
        lo, hi = int(i), min(int(i) + 1, n - 1)
        return vals[lo] + (vals[hi] - vals[lo]) * (i - lo)

    print(f"    {col}")
    print(f"      n={n} missing={miss} sum={sum(vals):,.0f} mean={sum(vals) / n:,.2f} "
          f"min={vals[0]:,.0f} p25={q(.25):,.1f} median={q(.5):,.1f} p75={q(.75):,.1f} max={vals[-1]:,.0f}")


def missingness(hdr, rows):
    print("    column                                                    missing      %")
    for h in hdr:
        if h.startswith("_") or h.startswith("meta/") or h.startswith("**") or h == "":
            continue
        m = sum(1 for r in rows if blank(r.get(h)))
        print(f"    {h[:55]:<55} {m:>7} {m / len(rows) * 100:6.1f}%")


def h1(t):
    print("\n" + "=" * 88 + f"\n{t}\n" + "=" * 88)


def h2(t):
    print(f"\n  -- {t}")


# ---------------------------------------------------------------
# load + build the parent join
# ---------------------------------------------------------------
hdr_sub, sub = load(S_SUB)
hdr_pres, pres = load(S_PRES)
hdr_item, item = load(S_ITEM)

parent = {(r["_submission__uuid"], r["_index"]): r for r in pres}


def par(r):
    return parent[(r["_submission__uuid"], r["_parent_index"])]


# ---------------------------------------------------------------
# 1. submissions
# ---------------------------------------------------------------
h1(f"SHEET 1 - '{S_SUB}' (facility visits): {len(sub)} rows x {len(hdr_sub)} cols")

h2("By district")
freq(sub, "District")
h2("By facility type")
freq(sub, "facility_type")
h2("By facility")
freq(sub, "Facility")
h2("By sector")
freq(sub, "Sector")

d = sorted(x for x in (as_date(r.get("_submission_time")) for r in sub) if x)
h2("Submission timestamps")
print(f"    {min(d)} .. {max(d)}  ({len(set(d))} distinct day(s))")

h2("Keys")
print(f"    _uuid distinct: {len(set(r['_uuid'] for r in sub))}/{len(sub)}    "
      f"_id distinct: {len(set(r['_id'] for r in sub))}/{len(sub)}")
print(f"    rows with a validation status: {sum(1 for r in sub if not blank(r.get('_validation_status')))}")

h2("Missingness")
missingness(hdr_sub, sub)

# ---------------------------------------------------------------
# 2. prescriptions
# ---------------------------------------------------------------
h1(f"SHEET 2 - '{S_PRES}' (prescriptions): {len(pres)} rows x {len(hdr_pres)} cols")

per_sub = Counter(r["_submission__uuid"] for r in pres)
h2("Prescriptions per submission")
print(f"    mean {len(pres) / len(sub):.2f}; distribution "
      f"{dict(sorted(Counter(per_sub.values()).items()))}")
fac_of = {r["_uuid"]: r["Facility"] for r in sub}
for u, k in sorted(per_sub.items(), key=lambda x: -x[1])[:4]:
    print(f"      {fac_of.get(u)}: {k} prescriptions in one submission")

h2("By district")
freq(pres, "District")
h2("By facility type")
freq(pres, "facility_type")
h2("By facility")
freq(pres, "Facility")
h2("By sector")
freq(pres, "Sector")
h2("By prescriber category")
freq(pres, PRESCRIBER)

dp = sorted(x for x in (as_date(r.get(DATE)) for r in pres) if x)
h2("Date of prescription")
print(f"    {min(dp)} .. {max(dp)}")
for k, v in sorted(Counter(dp).items()):
    print(f"    {k}  {v:>4}")

h2("Medicines per prescription")
desc(pres, N_MED)
freq(pres, N_MED)

h2("Mean medicines per prescription, by group")
for col in (PRESCRIBER, "facility_type", "Facility"):
    print(f"    by {col}:")
    g = defaultdict(list)
    for r in pres:
        g[str(r[col]).strip()].append(num(r[N_MED]) or 0)
    for k, v in sorted(g.items(), key=lambda x: -len(x[1])):
        print(f"      {k[:34]:<34} n={len(v):>4}  mean={sum(v) / len(v):.2f}")

h2("Kobo roll-up fields as exported")
for c in ("n_items_available", "n_items_unavailable", "pct_items_unserved"):
    desc(pres, c)

h2("Duplicate Patient / Prescription IDs")
g = defaultdict(list)
for r in pres:
    g[str(r[PID]).strip()].append(r)
dups = {k: v for k, v in g.items() if len(v) > 1}
print(f"    {len(set(g))} distinct IDs over {len(pres)} prescriptions; {len(dups)} ID(s) used twice")
for k, v in dups.items():
    for r in v:
        print(f"      {k} | {r['Facility']} | {as_date(r[DATE])} | {r[PRESCRIBER]} | n_med={r[N_MED]}")

h2("Missingness")
missingness(hdr_pres, pres)

# ---------------------------------------------------------------
# 3. medicine items
# ---------------------------------------------------------------
h1(f"SHEET 3 - '{S_ITEM}' (prescribed medicines): {len(item)} rows x {len(hdr_item)} cols")

cnt = Counter((r["_submission__uuid"], r["_parent_index"]) for r in item)
h2("Rows per parent prescription")
v = sorted(cnt.values())
print(f"    parents with items: {len(cnt)}/{len(pres)}   min={v[0]} median={v[len(v) // 2]} max={v[-1]}")
print(f"    rows where the count != the parent's '{N_MED}': "
      f"{sum(1 for r in pres if cnt.get((r['_submission__uuid'], r['_index']), 0) != num(r[N_MED]))}")

h2("On the EML?")
freq(item, EML)
h2("Name type recorded")
freq(item, "Name type")
brand = [r for r in item if str(r["Name type"]).strip() == "Brand"]
print(f"    rows typed 'Brand' but with EML='Yes': "
      f"{sum(1 for r in brand if yes(r[EML]))}/{len(brand)}")

h2("All medicines selected from the EML picklist")
freq(item, MED_EML, show_blank=False)
h2("Medicines typed free-text (off-EML)")
freq(item, MED_FREE, show_blank=False)

h2("Drug class, EML numbering stripped (all)")
c = Counter(re.sub(r"^[\d.]+\s*", "", str(r["Drug class"] or "(blank)")).strip().upper() for r in item)
for k, v in c.most_common():
    print(f"    {k[:60]:<60} {v:>4}  {v / len(item) * 100:4.1f}%")

h2("Quantity prescribed")
desc(item, "Qty prescribed")
h2("Qty prescribed - full frequency distribution")
freq(item, "Qty prescribed")
print("    the 15 largest quantities:")
for r in sorted(item, key=lambda r: -(num(r["Qty prescribed"]) or 0))[:15]:
    print(f"      {num(r['Qty prescribed']):>6.0f}  {r[MED_EML] or r[MED_FREE]}  ({par(r)['Facility']})")

h2("Availability at the facility on the prescription date")
freq(item, AVAIL)

unserved = [r for r in item if no(r[AVAIL])]
h2(f"Unserved items: {len(unserved)}/{len(item)} ({len(unserved) / len(item) * 100:.1f}%)")
for c in (ORDERED, SUPPLIED, RMS_REG, RMS_CEN, "distribution_nationwide"):
    print(f"    {c[:70]}")
    freq(unserved, c, show_blank=False)
desc(unserved, UNIT_PRICE)
desc(unserved, "total_cost_unserved")
tc = sum(num(r["total_cost_unserved"]) or 0 for r in unserved)
print(f"    total value of unserved medicines: {tc:,.0f} RWF")

h2("The unserved items in full")
for r in unserved:
    p = par(r)
    print(f"    {p['Facility']} ({p['facility_type']}, {p['District']}) | {as_date(p[DATE])} | {p[PRESCRIBER]}")
    print(f"      {r[MED_EML] or r[MED_FREE]}  EML={r[EML]}  qty={r['Qty prescribed']}  "
          f"ordered={r[ORDERED]}  supplied={r[SUPPLIED]}")
    print(f"      RMS regional={r[RMS_REG]}  RMS central={r[RMS_CEN]}  flag={r['distribution_nationwide']}  "
          f"unit={r[UNIT_PRICE]} RWF  total={r['total_cost_unserved']} RWF")

h2("Items and unserved rate by facility (geography taken from the parent)")
tab = defaultdict(lambda: [0, 0])
for r in item:
    p = par(r)
    k = (p["facility_type"], p["Facility"])
    tab[k][0] += 1
    tab[k][1] += 1 if no(r[AVAIL]) else 0
pc = Counter((p["facility_type"], p["Facility"]) for p in pres)
print(f"    {'type':<18}{'facility':<24}{'presc':>6}{'items':>7}{'unserved':>10}{'rate':>8}")
for k, (i, u) in sorted(tab.items(), key=lambda x: -x[1][0]):
    print(f"    {k[0]:<18}{k[1]:<24}{pc[k]:>6}{i:>7}{u:>10}{u / i * 100:>7.1f}%")
tot_u = sum(x[1] for x in tab.values())
print(f"    {'TOTAL':<42}{len(pres):>6}{len(item):>7}{tot_u:>10}{tot_u / len(item) * 100:>7.1f}%")

h2("Items and unserved by prescriber category (from the parent)")
t2 = defaultdict(lambda: [0, 0])
for r in item:
    k = par(r)[PRESCRIBER]
    t2[k][0] += 1
    t2[k][1] += 1 if no(r[AVAIL]) else 0
for k, (i, u) in sorted(t2.items(), key=lambda x: -x[1][0]):
    print(f"    {str(k)[:30]:<30}{i:>5} items{u:>5} unserved")

h2("item_number within prescription")
freq(item, "item_number")

h2("Items by district and facility type (from the parent)")
t3 = Counter((par(r)["District"], par(r)["facility_type"]) for r in item)
for k, v in sorted(t3.items(), key=lambda x: -x[1]):
    print(f"    {k[0]:<14}{k[1]:<20}{v:>5}  {v / len(item) * 100:4.1f}%")

h2("EML status and name type by facility (from the parent)")
t4 = defaultdict(lambda: [0, 0, 0])
for r in item:
    k = par(r)["Facility"]
    t4[k][0] += 1
    t4[k][1] += 1 if yes(r[EML]) else 0
    t4[k][2] += 1 if str(r["Name type"]).strip() == "Brand" else 0
print(f"    {'facility':<24}{'items':>6}{'on EML':>8}{'brand-typed':>13}")
for k, (i, e, b) in sorted(t4.items(), key=lambda x: -x[1][0]):
    print(f"    {k:<24}{i:>6}{e:>8}{b:>13}")

h2("Top drug classes by facility (from the parent)")
t5 = defaultdict(Counter)
for r in item:
    t5[par(r)["Facility"]][re.sub(r"^[\d.]+\s*", "", str(r["Drug class"] or "(blank)")).strip().upper()] += 1
for f in sorted(t5, key=lambda f: -sum(t5[f].values())):
    print(f"    {f}:")
    for k, v in t5[f].most_common(5):
        print(f"      {k[:56]:<56}{v:>4}")

h2("Prescriptions per day per facility")
t6 = defaultdict(Counter)
for r in pres:
    t6[r["Facility"]][as_date(r[DATE])] += 1
for f in sorted(t6, key=lambda f: -sum(t6[f].values())):
    days = ", ".join(f"{d} ({n})" for d, n in sorted(t6[f].items()))
    print(f"    {f}: {days}")

h2("Prescription-level service status, recomputed from medicine_item")
av = defaultdict(int)
un = defaultdict(int)
for r in item:
    k = (r["_submission__uuid"], r["_parent_index"])
    av[k] += 1 if yes(r[AVAIL]) else 0
    un[k] += 1 if no(r[AVAIL]) else 0
full = part = none_ = 0
for r in pres:
    k = (r["_submission__uuid"], r["_index"])
    if un[k] == 0:
        full += 1
    elif av[k] == 0:
        none_ += 1
    else:
        part += 1
print(f"    fully served     {full:>4} ({full / len(pres) * 100:.1f}%)")
print(f"    partly unserved  {part:>4} ({part / len(pres) * 100:.1f}%)")
print(f"    fully unserved   {none_:>4} ({none_ / len(pres) * 100:.1f}%)")

h2("Missingness")
missingness(hdr_item, item)

# ---------------------------------------------------------------
# 4. data quality / integrity
# ---------------------------------------------------------------
h1("DATA QUALITY AND CROSS-SHEET INTEGRITY")

h2("Link integrity")
sub_u = {r["_uuid"] for r in sub}
print(f"    prescription rows with an unknown submission uuid : "
      f"{sum(1 for r in pres if r['_submission__uuid'] not in sub_u)}")
print(f"    medicine_item rows with no matching parent        : "
      f"{sum(1 for r in item if (r['_submission__uuid'], r['_parent_index']) not in parent)}")
print(f"    submissions with no prescription rows             : "
      f"{len(sub_u - {r['_submission__uuid'] for r in pres})}")

h2("medicine_item's own copies of parent fields, vs the parent")
for col, pcol in ((PID, PID), ("Facility", "Facility"), (PRESCRIBER, PRESCRIBER)):
    m = sum(1 for r in item if str(r[col]).strip() != str(par(r)[pcol]).strip())
    print(f"    {col[:45]:<45} {m:>4}/{len(item)} disagree")
m = sum(1 for r in item if as_date(r[DATE]) != as_date(par(r)[DATE]))
print(f"    {DATE:<45} {m:>4}/{len(item)} disagree")
print(f"    rows where medicine_item.Facility is literally '#N/A': "
      f"{sum(1 for r in item if str(r['Facility']).strip() == '#N/A')}")
print("    -> treat these 8 columns as unusable; join to 'prescription' instead.")

h2("Kobo roll-ups on 'prescription', vs recomputed from medicine_item")
mm_a = sum(1 for r in pres if av[(r["_submission__uuid"], r["_index"])] != num(r["n_items_available"]))
mm_u = sum(1 for r in pres if un[(r["_submission__uuid"], r["_index"])] != num(r["n_items_unavailable"]))
print(f"    n_items_available   disagrees on {mm_a}/{len(pres)} "
      f"(exported sum {sum(num(r['n_items_available']) or 0 for r in pres):.0f} vs actual {sum(av.values())})")
print(f"    n_items_unavailable disagrees on {mm_u}/{len(pres)} "
      f"(exported sum {sum(num(r['n_items_unavailable']) or 0 for r in pres):.0f} vs actual {sum(un.values())})")
print(f"    available+unavailable != '{N_MED}' on "
      f"{sum(1 for r in pres if (num(r['n_items_available']) or 0) + (num(r['n_items_unavailable']) or 0) != num(r[N_MED]))}"
      f"/{len(pres)} rows")
print("    -> recompute n_items_available / n_items_unavailable / pct_items_unserved from medicine_item.")

h2("Other flags")
print(f"    'Drug class' = 'Not specified in EML source': "
      f"{sum(1 for r in item if str(r['Drug class']).strip() == 'Not specified in EML source')} items")
print(f"    duplicate Patient / Prescription IDs: {len(dups)}")
print(f"    all {len(sub)} submissions were uploaded on {min(d)} (single sync), while prescription "
      f"dates span {min(dp)}..{max(dp)} - retrospective chart abstraction, not same-day capture.")
