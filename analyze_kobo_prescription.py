"""
Full analysis of the prescription-study Kobo export -> one Excel workbook.

    python3 fetch_kobo_export.py          # refresh Prescription_kobo_export.xlsx
    python3 analyze_kobo_prescription.py  # -> Prescription_study_analysis.xlsx

The workbook holds, in order:

  Overview                 dataset grain, coverage, headline indicators
  1 Facility visits        source sheet 1, all columns
  2 Prescriptions          source sheet 2 + general info pulled down from sheet 1
  3 Medicine items         source sheet 3 + general info AND prescription context
  Geography ... Findings   analysis sections
  Charts                   native Excel charts

"Pull the general information down" is the point of the three data sheets: in the
raw export, Province / District / Sector / Facility / facility type / collector /
submission date live ONLY on sheet 1, and sheet 3 does not even carry the
prescription's date or prescriber. Each level is joined back to its parent so
every row is self-contained and filterable on its own:

    medicine_item (_submission__uuid, _parent_index)
        -> prescription (_submission__uuid, _index)
        -> submission (_uuid)

Derived columns are added to the data sheets and marked "[derived]" in the header:

  - Prescriptions: items available / not available / % unserved, recomputed.
    The form's own n_items_available, n_items_unavailable and pct_items_unserved
    calculate fields are EMPTY on all rows of this export (the form has three
    __version__ values, so these were almost certainly added after collection
    started), so every unserved figure here is recomputed from medicine_item.
  - Medicine items: dosage form and injection flag parsed from the medicine code
    (paracetamol_tab_500_mg -> tab), antibiotic flag from the drug-class number.

openpyxl + the standard library only (no pandas), like the other scripts here.
"""

import datetime as dt
import math
import re
from collections import Counter, defaultdict

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SRC = "Prescription_kobo_export.xlsx"
OUT_XLSX = "Prescription_study_analysis.xlsx"

S_PRES = "prescription"
S_ITEM = "medicine_item"
# Sheet 1 is named after the FORM, so renaming the project in Kobo renames it:
# it arrived as "Prescription assessment_DC" and became "Prescription
# Assessment" on 2026-09-23. The two repeat groups keep their XLSForm names, so
# sheet 1 is simply the one that is neither - resolved rather than written down.
S_SUB = next(n for n in openpyxl.load_workbook(SRC, read_only=True).sheetnames
             if n not in (S_PRES, S_ITEM))

# --- sheet 1 (submission / facility visit) -------------------------------
COLLECTOR = "Data collection name"
PROV, DISTR, SECTOR, FACIL = "Province", "District", "Sector", "Facility"
FTYPE = "facility_type"
SUB_TIME = "_submission_time"
VISIT_DATE = "today"

# --- sheet 2 (prescription) ----------------------------------------------
DATE = "Date of prescription"
PID = "Patient / Prescription ID"
PRESCRIBER = "Prescriber category"
ALLIED = "Please specify the allied health professional type"
N_MED = "Number of medicines prescribed"
FEEDBACK = ("Please share any feedback, concerns, or recommendations related to the "
            "supply chain and other relevant areas.")

# --- sheet 3 (medicine item) ---------------------------------------------
ITEM_NO = "item_number"
# The form has renamed these more than once, and a constant that no longer
# matches reads as an empty column rather than an error: every medicine became
# "(not recorded)" and EML adherence fell to 7% before this was caught on
# 2026-09-25. Each name below is the one present in the CURRENT merged export.
# "Is this medicine on the EML?" is asked only on the newer (EU) form version,
# so the EML indicator uses the calculated flag both versions carry.
EML = "on_eml"
MED_EML = "Medicine name"
MED_FREE = "Medicine name (not on the EML list) - please specify"
CLASS = "Drug class"
CLASS_FREE = "Drug class - please specify"
NAME_TYPE = "Name type"
QTY = "Qty prescribed"
AVAIL = "Was the medicine available at the health facility during the date of prescription?"
AV_FLAG, UNAV_FLAG = "is_available_flag", "is_unavailable_flag"
ORDERED = "Was the medicine ordered during the month the prescription was made?"
SUPPLIED = "Was it supplied fully, partially, or not at all?"
UNIT_PRICE = "Unit price for the medicine that was not available (RWF)"
COST_UNSERVED = "total_cost_unserved"
RMS_REG = "Was the medicine available at RMS Regional during the date of prescription?"
RMS_CEN = "Was the medicine available at RMS Central during the date of prescription?"
DIST_NAT = "distribution_nationwide"

# Kobo "note" widgets - display-only template strings, empty on every row here.
DROP_COLS = {
    "**${n_medicines_prescribed}** medicine(s) prescribed - **${n_items_available}**"
    " available, **${n_items_unavailable}** not available.",
    "**Medicine ${item_number} - prescription ${patient_prescription_id}**",
}


# ---------------------------------------------------------------
# read
# ---------------------------------------------------------------
def load(sheet):
    wb = openpyxl.load_workbook(SRC, data_only=True, read_only=True)
    rows = list(wb[sheet].iter_rows(values_only=True))
    hdr = [str(h).strip() if h is not None else "" for h in rows[0]]
    data = [dict(zip(hdr, r)) for r in rows[1:]
            if any(c is not None and str(c).strip() != "" for c in r)]
    return [h for h in hdr if h not in DROP_COLS], data


def blank(v):
    return v is None or str(v).strip() == ""


def num(v):
    if blank(v):
        return None
    try:
        f = float(str(v).strip())
    except ValueError:
        return None
    # Kobo calculate fields can emit the literal string "NaN", which float()
    # accepts and which then turns every downstream sum into NaN. Treat a
    # non-finite value as missing. (Seen in the stock-assessment export; guarded
    # here too since both forms share the same calculate-field machinery.)
    return f if math.isfinite(f) else None


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


def pct(a, b):
    return (a / b) if b else None


hdr_sub, sub = load(S_SUB)
hdr_pres, pres = load(S_PRES)
hdr_item, item = load(S_ITEM)

# ---------------------------------------------------------------
# facility corrections
# ---------------------------------------------------------------
# A health post that has no entry in the Kobo facility list gets captured under
# its parent health centre. The field team identifies which submissions those
# are; each is keyed by (facility as recorded, submission date) - the same terms
# the correction is given in - so a re-fetch reapplies it and a late-arriving
# submission on that date is caught too.
#
# Every health post in this dataset is typed SGHP, so a reattributed post takes
# that type; leaving it as HC would misplace it in the facility-type tables.
# District and sector are left alone: a post sits in its parent's catchment.
FACILITY_OVERRIDES = {
    ("Nyanza CS", "2026-09-16"): {FACIL: "Kiberinka Gahondo HP", FTYPE: "SGHP"},
    # All 17 Rwimbogo CS submissions fall on this date, so Rwimbogo CS leaves
    # the dataset entirely once this is applied.
    ("Rwimbogo CS", "2026-09-16"): {FACIL: "Kiburara HP", FTYPE: "SGHP"},
    # Named "Ituze SGHP" as the user gave it. Every other post here is named
    # "<name> HP" with facility_type SGHP, so this name alone carries the
    # level twice - kept as given rather than normalised to "Ituze HP".
    # Likewise all 21 Kigese CS submissions fall on this date.
    ("Kigese CS", "2026-09-15"): {FACIL: "Ituze SGHP", FTYPE: "SGHP"},
    # All 20 Ruhango CS submissions fall on this date too, so Ruhango CS also
    # leaves the dataset entirely. "Ruhango PH" is a different facility in
    # the same district and is NOT touched - the key matches the name exactly.
    ("Ruhango CS", "2026-09-16"): {FACIL: "Bunyogombe HP", FTYPE: "SGHP"},
    # Mbabe HP still has no entry in the prescription tool's facility list, so
    # its data goes in under Masaka CS. All 20 Masaka CS submissions share one
    # date and one collector (Kampeta Edith) and no new ones have arrived since
    # 2026-09-16, so the whole group is treated as Mbabe - the same rule the
    # three above use. Typed HP, not SGHP: the STOCK tool carries this facility
    # in its own list as "MBABE HP", level HP, Kicukiro / Masaka - the same
    # district and sector these rows already have.
    ("Masaka CS", "2026-09-16"): {FACIL: "Mbabe HP", FTYPE: "HP"},
    # Not a health post this time - a level misrecorded. The 14 Sep visits went
    # to the health centre and were entered against the district hospital;
    # the 24 Sep ones are genuinely the DH and must NOT move, which is the
    # first case where the date half of the key does real work. Corroborated
    # by the other two instruments: the stock assessment has Kirinda CS on
    # 2026-09-14 (51 submissions) and Kirinda DH on 2026-09-24, and the CHW
    # assessment has Kirinda CS on 2026-09-14.
    ("Kirinda DH", "2026-09-14"): {FACIL: "Kirinda CS", FTYPE: "HC"},
    # Keyed on the VISIT date, not the submission date: the team was at the
    # facility on 18 Sep and the form did not reach Kobo until the 21st. The
    # collector (Kareranzira Placide) also entered a genuine "Gikongoro
    # Prison" visit on 14 Sep, which this must not disturb - hence the exact
    # (facility, date) pair rather than a blanket rename.
    ("Kigeme Refugee Camp CS", "2026-09-18"): {FACIL: "Gikongoro Prison CS",
                                               FTYPE: "HC", "_on": "visit"},
}


def apply_facility_overrides(rows):
    """Rewrite Facility before anything else reads it, so every breakdown,
    scorecard and data sheet sees the correction. Keeps the original in
    _facility_was for the audit table on the Data quality sheet.

    The date in a key is the SUBMISSION date by default. A correction is
    usually given in terms of the day the team was at the facility, though,
    and a form filled on paper can be submitted days later - so an entry may
    set "_on": "visit" to key on the visit date (`today`) instead. Keying on
    the wrong one either matches nothing or, worse, sweeps up a different
    visit that happens to share the other date.
    """
    applied = Counter()
    fields = {"submitted": SUB_TIME, "visit": VISIT_DATE}
    for r in rows:
        fac = str(r.get(FACIL)).strip()
        for tag, field in fields.items():
            key = (fac, str(r.get(field))[:10])
            fix = FACILITY_OVERRIDES.get(key)
            if not fix or fix.get("_on", "submitted") != tag:
                continue
            r["_facility_was"] = r.get(FACIL)
            r.update({k: v for k, v in fix.items() if not k.startswith("_")})
            applied[key] += 1
            break
    return applied


OVERRIDES_APPLIED = apply_facility_overrides(sub)

# ---------------------------------------------------------------
# facilities removed from every analysis
# ---------------------------------------------------------------
# Facilities named here are dropped after the overrides, before anything is
# joined or counted, so every output built on this module leaves them out.
# Matched on the facility name exactly. Empty since 2026-10-18: "Gatsibo CS"
# was removed and then restored at the user's request.
EXCLUDED_FACILITIES = set()
_excl_uuids = {r["_uuid"] for r in sub
               if str(r.get(FACIL)).strip() in EXCLUDED_FACILITIES}
EXCLUDED_COUNTS = (len(_excl_uuids),
                   sum(1 for r in pres
                       if r["_submission__uuid"] in _excl_uuids),
                   sum(1 for r in item
                       if r["_submission__uuid"] in _excl_uuids))
sub[:] = [r for r in sub if r["_uuid"] not in _excl_uuids]
pres[:] = [r for r in pres if r["_submission__uuid"] not in _excl_uuids]
item[:] = [r for r in item if r["_submission__uuid"] not in _excl_uuids]

# ---------------------------------------------------------------
# join the three levels
# ---------------------------------------------------------------
sub_by_uuid = {r["_uuid"]: r for r in sub}
pres_by_key = {(r["_submission__uuid"], r["_index"]): r for r in pres}

for r in pres:
    r["_sub"] = sub_by_uuid.get(r["_submission__uuid"])
for r in item:
    p = pres_by_key.get((r["_submission__uuid"], r["_parent_index"]))
    r["_pres"] = p
    r["_sub"] = p["_sub"] if p else sub_by_uuid.get(r["_submission__uuid"])

ORPHAN_PRES = sum(1 for r in pres if r["_sub"] is None)
ORPHAN_ITEM = sum(1 for r in item if r["_pres"] is None)

items_of = defaultdict(list)
for r in item:
    if r["_pres"] is not None:
        items_of[id(r["_pres"])].append(r)


def g(row, col):
    """A general-information field, wherever the row sits in the hierarchy."""
    s = row.get("_sub")
    return s.get(col) if s else None


# ---------------------------------------------------------------
# derived fields
# ---------------------------------------------------------------
# Dosage form from the medicine code (paracetamol_tab_500_mg -> tab). The form
# collects no dosage-form field, and the code is the only place it appears, so
# the WHO "% prescriptions with an injection" indicator has to come from here.
FORMS = [
    ("inj", r"(^|_)(inj|injection|injections)(_|$)"),
    ("infusion", r"(^|_)infusion(_|$)"),
    ("tab", r"(^|_)(tab|tabs|tablet)(_|$)"),
    ("caps", r"(^|_)(caps|capsule)(_|$)"),
    ("syrup", r"(^|_)(syr|syrup)(_|$)"),
    ("suspension", r"(^|_)(susp|suspension)(_|$)"),
    ("suppository", r"(^|_)suppository(_|$)"),
    ("cream/ointment", r"(^|_)(cream|ointment|gel)(_|$)"),
    ("drops", r"(^|_)(drops|drop)(_|$)"),
    ("solution", r"(^|_)solution(_|$)"),
    ("powder", r"(^|_)powder(_|$)"),
]


def dosage_form(code):
    c = str(code).strip().lower()
    if blank(code):
        return "(not recorded)"
    for name, pat in FORMS:
        if re.search(pat, c):
            return name
    return "(other/unclassified)"


def class_no(v):
    """Leading EML section number of a drug class, normalised ('25.2. PARENTERAL
    ROUTE' and '25.2 Parenteral route' are the same section typed two ways)."""
    m = re.match(r"\s*(\d+(?:\.\d+)*)", str(v or ""))
    return m.group(1).rstrip(".") if m else ""


# Systemic antibacterials: EML section 11.2 (beta-lactams) and 11.3 (other
# anti-bacterials). Topical/ophthalmic anti-infectives (12.1.2, 14.1) are
# reported separately rather than folded into the WHO antibiotic indicator.
def is_antibiotic(r):
    n = class_no(r.get(CLASS))
    return n.startswith("11.2") or n.startswith("11.3")


def is_topical_antiinfective(r):
    n = class_no(r.get(CLASS))
    return n.startswith("12.1.2") or n.startswith("14.1")


def med_label(r):
    """Resolved medicine name: the EML code, else the free-text specify field."""
    if not blank(r.get(MED_EML)):
        return str(r[MED_EML]).strip()
    if not blank(r.get(MED_FREE)):
        return str(r[MED_FREE]).strip() + " (free text)"
    return "(not recorded)"


# ---------------------------------------------------------------
# drug class corrections
# ---------------------------------------------------------------
# Eight respiratory medicines were coded in the form under
# "15.6.2 Medicines for diarrhoea". The misfiling is in the instrument's own
# choice list, not in the data collectors' answers, so it affects every
# submission that recorded one of them and cannot be fixed by re-visiting a
# facility. Reported 2026-10-06.
#
# The medicines are keyed by their EML code - the stable identifier - and NOT
# by their label, because the same product appears both as a code (m1495) and
# as a slugged name (salbutamol_inhalation_solution100_g_dose) depending on
# which server and form version the submission came from, and both spellings
# have to be caught.
#
# Zinc sulphate, which also sits under 15.6.2, is left exactly where it is:
# it IS a diarrhoea medicine and is the only thing in that class that belongs
# there.
RESPIRATORY = "26. MEDICINES ACTING ON THE RESPIRATORY TRACT"

CLASS_OVERRIDES = {
    # salbutamol, all forms
    "m1495": RESPIRATORY,
    "salbutamol_inhalation_solution100_g_dose": RESPIRATORY,
    "m1497": RESPIRATORY,
    "salbutamol_nebuliser_solution_5mg_ml": RESPIRATORY,
    "m1494": RESPIRATORY,
    "salbutamol_bromexine_guaifenesine_syrup_2mg_4mg_100mg": RESPIRATORY,
    # inhaled corticosteroid
    "m0189": RESPIRATORY,
    "beclomethasone_inhalation_solution_250_g": RESPIRATORY,
    # xanthine bronchodilator
    "m0111": RESPIRATORY,
    "aminophylline_tab_100_mg": RESPIRATORY,
    # cough and cold combinations
    "m0334": RESPIRATORY,
    "chlorpheniramine_diphenyhydramine_phenylephrine_syrup_0_1gr_0_015gr_0_1_gr":
        RESPIRATORY,
    # injectable corticosteroid, listed on the respiratory sheet
    "triamcinolone_acetonide_inj_40_80mg_1_ml": RESPIRATORY,
}

CLASS_FIXED = 0


def class_label(r):
    global CLASS_FIXED
    key = str(r.get(MED_EML) or "").strip()
    if key in CLASS_OVERRIDES:
        CLASS_FIXED += 1
        return CLASS_OVERRIDES[key]
    if not blank(r.get(CLASS)):
        return str(r[CLASS]).strip()
    if not blank(r.get(CLASS_FREE)):
        return str(r[CLASS_FREE]).strip() + " (free text)"
    return "(not recorded)"


for r in item:
    r["_form"] = dosage_form(r.get(MED_EML))
    r["_is_inj"] = r["_form"] in ("inj", "infusion")
    r["_abx"] = is_antibiotic(r)
    r["_med"] = med_label(r)
    r["_class"] = class_label(r)
    # From the CORRECTED class, so CLASS_OVERRIDES reaches the section code
    # too - read from the raw column it kept the respiratory medicines under
    # 15.6.2 for anything that groups on the code.
    r["_class_no"] = class_no(r["_class"])
    # Availability: prefer the text answer, fall back to the form's own flag.
    if not blank(r.get(AVAIL)):
        r["_avail"] = yes(r.get(AVAIL))
    else:
        f = num(r.get(AV_FLAG))
        r["_avail"] = (f == 1) if f is not None else None

# Recomputed prescription roll-ups (the exported calculate fields are empty).
for r in pres:
    its = items_of.get(id(r), [])
    r["_n_items"] = len(its)
    r["_n_avail"] = sum(1 for i in its if i["_avail"] is True)
    r["_n_unavail"] = sum(1 for i in its if i["_avail"] is False)
    r["_pct_unserved"] = pct(r["_n_unavail"], r["_n_items"])
    r["_fully_served"] = bool(its) and r["_n_unavail"] == 0
    r["_any_unserved"] = r["_n_unavail"] > 0
    r["_abx"] = any(i["_abx"] for i in its)
    r["_inj"] = any(i["_is_inj"] for i in its)
    r["_n_declared"] = num(r.get(N_MED))
    r["_count_mismatch"] = (r["_n_declared"] is not None
                            and int(r["_n_declared"]) != r["_n_items"])

for r in sub:
    mine = [p for p in pres if p["_sub"] is r]
    r["_n_pres"] = len(mine)
    r["_n_items"] = sum(p["_n_items"] for p in mine)


# ---------------------------------------------------------------
# table builders
# ---------------------------------------------------------------
class Table:
    def __init__(self, title, headers, rows, note=None, total=None, chart=None):
        self.title = title
        self.headers = headers
        self.rows = rows
        self.note = note
        self.total = total
        # chart: (kind, label_col, value_col, max_rows); kind "v" or "h".
        self.chart = chart
        self.pct = {i for i, h in enumerate(headers)
                    if "%" in str(h) or str(h).strip().lower().endswith("rate")}

    def chart_data(self):
        if not self.chart:
            return []
        _, lc, vc, cap = self.chart
        out = []
        for r in self.rows:
            v = r[vc]
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                out.append((str(r[lc]), v))
        return out[:cap] if cap else out


def counts(rows, key, denom=None, top=None, keyname="Value", skip_blank=False):
    """Frequency table over key(row) (a column name or a callable)."""
    f = key if callable(key) else (lambda r: r.get(key))
    vals = []
    for r in rows:
        v = f(r)
        if skip_blank and blank(v):
            continue
        vals.append("(blank)" if blank(v) else str(v).strip())
    c = Counter(vals)
    d = denom if denom is not None else len(vals)
    ordered = c.most_common(top) if top else sorted(c.items(), key=lambda kv: (-kv[1], kv[0]))
    return [[k, n, pct(n, d)] for k, n in ordered], len(c), sum(c.values())


def freq_table(title, rows, key, keyname, note=None, top=None, chart=None,
               skip_blank=False):
    body, ndist, tot = counts(rows, key, top=top, keyname=keyname, skip_blank=skip_blank)
    return Table(title, [keyname, "n", "% of rows"], body, note=note,
                 total=["Total" + (f" ({ndist} distinct)" if not top else ""), tot,
                        1.0 if tot else None],
                 chart=chart)


def stats_row(label, values):
    v = sorted(x for x in values if x is not None)
    if not v:
        return [label, 0, None, None, None, None, None, None]

    def q(p):
        if len(v) == 1:
            return v[0]
        i = p * (len(v) - 1)
        lo, hi = int(i), min(int(i) + 1, len(v) - 1)
        return v[lo] + (v[hi] - v[lo]) * (i - lo)

    return [label, len(v), sum(v) / len(v), v[0], q(0.25), q(0.5), q(0.75), v[-1]]


STAT_HDR = ["Measure", "n", "Mean", "Min", "P25", "Median", "P75", "Max"]


def avail_by(title, rows, key, keyname, note=None, top=None, chart=None,
             min_items=1):
    """Availability / unserved breakdown of medicine items by some grouping."""
    f = key if callable(key) else (lambda r: r.get(key))
    grp = defaultdict(list)
    for r in rows:
        v = f(r)
        grp["(blank)" if blank(v) else str(v).strip()].append(r)
    body = []
    for k, rs in grp.items():
        known = [r for r in rs if r["_avail"] is not None]
        na = sum(1 for r in known if not r["_avail"])
        if len(rs) < min_items:
            continue
        body.append([k, len(rs), sum(1 for r in known if r["_avail"]), na,
                     pct(na, len(known))])
    body.sort(key=lambda r: (-r[1], r[0]))
    if top:
        body = body[:top]
    t_items = sum(r[1] for r in body)
    t_na = sum(r[3] for r in body)
    t_a = sum(r[2] for r in body)
    return Table(title,
                 [keyname, "Items", "Available", "Not available", "% not available"],
                 body, note=note,
                 total=["Total", t_items, t_a, t_na, pct(t_na, t_a + t_na)],
                 chart=chart)


def missing_table(hdr, rows, title, drop_prefix=("_",)):
    body = []
    for h in hdr:
        if h.startswith(drop_prefix):
            continue
        n = sum(1 for r in rows if blank(r.get(h)))
        body.append([h, n, pct(n, len(rows))])
    body.sort(key=lambda r: (-r[1], r[0]))
    return Table(title, ["Column", "Blank rows", "% blank"], body,
                 note="Skip-logic fields are legitimately blank when the question was "
                      "never asked - on this form the unserved follow-ups (ordered, "
                      "supplied, unit price, RMS, distribution) are only asked for an "
                      "item recorded as not available, so ~99.5% blank is expected there.")


# ---------------------------------------------------------------
# headline numbers
# ---------------------------------------------------------------
N_SUB, N_PRES, N_ITEM = len(sub), len(pres), len(item)
ITEM_AV = sum(1 for r in item if r["_avail"] is True)
ITEM_NA = sum(1 for r in item if r["_avail"] is False)
ITEM_UNK = sum(1 for r in item if r["_avail"] is None)
PRES_ANY_NA = sum(1 for r in pres if r["_any_unserved"])
N_VERSIONS = len(set(str(r.get("__version__")) for r in sub))
N_FEEDBACK = sum(1 for r in pres if not blank(r.get(FEEDBACK)))
# Per-theme feedback counts quoted in Findings. fb_theme is defined further
# down, so these are filled in by _count_themes() once it exists.
FB_OUTSIDE = FB_TIMELY = FB_AUTOMATE = FB_OPENMRS = FB_REGISTER = 0
FB_ENFORCE = 0
N_PID_DISTINCT = len({str(r.get(PID)).strip() for r in pres if not blank(r.get(PID))})
sub_times = [r[SUB_TIME] for r in sub if not blank(r.get(SUB_TIME))]
pres_dates = [as_date(r.get(DATE)) for r in pres]
pres_dates = [d for d in pres_dates if d]
COST = sum(num(r.get(COST_UNSERVED)) or 0 for r in item)
sub_days = [as_date(t) for t in sub_times if as_date(t)]
SUB_SPAN_DAYS = (max(sub_days) - min(sub_days)).days + 1 if sub_days else 0
PRES_SPAN_DAYS = (max(pres_dates) - min(pres_dates)).days + 1 if pres_dates else 0


def overview():
    return [
        Table("Dataset", ["Item", "Value"], [
            ["Source workbook", SRC],
            ["Downloaded / analysed", dt.date.today().isoformat()],
            ["Facility visits (submissions)", N_SUB],
            ["Prescriptions", N_PRES],
            ["Medicine items", N_ITEM],
            ["Provinces / districts / sectors", f"{len(set(r[PROV] for r in sub))}"
             f" / {len(set(r[DISTR] for r in sub))} / {len(set(r[SECTOR] for r in sub))}"],
            ["Facilities", len(set(r[FACIL] for r in sub))],
            ["Facility types", len(set(str(r[FTYPE]) for r in sub))],
            ["Data collectors", len(set(str(r[COLLECTOR]) for r in sub))],
            ["Form versions in this export", len(set(str(r.get("__version__")) for r in sub))],
            ["Submission window", f"{min(sub_times)} -> {max(sub_times)}" if sub_times else "n/a"],
            ["Prescription dates", f"{min(pres_dates)} -> {max(pres_dates)}" if pres_dates else "n/a"],
        ], note="Collection was still open when this export was taken - the submission "
                "window is far shorter than the span of prescription dates, so treat "
                "totals as an interim cut, not a closed dataset."),

        Table("Headline indicators", ["Indicator", "Value"], [
            ["Medicine items per prescription (mean)", N_ITEM / N_PRES if N_PRES else None],
            ["Prescriptions per facility visit (mean)", N_PRES / N_SUB if N_SUB else None],
            ["Items recorded available", ITEM_AV],
            ["Items recorded NOT available", ITEM_NA],
            ["% of items not available", pct(ITEM_NA, ITEM_AV + ITEM_NA)],
            ["Prescriptions with >=1 item not available", PRES_ANY_NA],
            ["% of prescriptions with >=1 item not available", pct(PRES_ANY_NA, N_PRES)],
            ["Total recorded cost of unserved items (RWF)", COST],
        ], note="The unserved rate is strikingly low for a study designed around "
                "unserved prescriptions - see the Data quality sheet before reading it "
                "as a result."),
    ]


# ---------------------------------------------------------------
# sections
# ---------------------------------------------------------------
def geography():
    return [
        freq_table("Facility visits by province", sub, PROV, "Province",
                   chart=("v", 0, 1, 12)),
        freq_table("Facility visits by district", sub, DISTR, "District",
                   chart=("h", 0, 1, 25)),
        freq_table("Facility visits by facility type", sub, FTYPE, "Facility type",
                   chart=("v", 0, 1, 12)),
        freq_table("Prescriptions by province", pres, lambda r: g(r, PROV), "Province",
                   chart=("v", 0, 1, 12)),
        freq_table("Prescriptions by district", pres, lambda r: g(r, DISTR), "District",
                   chart=("h", 0, 1, 25)),
        freq_table("Medicine items by province", item, lambda r: g(r, PROV), "Province"),
        Table("Coverage by province", ["Province", "Districts", "Facilities", "Visits",
                                       "Prescriptions", "Items", "Items / prescription"],
              _coverage(lambda r: g(r, PROV), PROV)),
        Table("Coverage by facility type", ["Facility type", "Districts", "Facilities",
                                            "Visits", "Prescriptions", "Items",
                                            "Items / prescription"],
              _coverage(lambda r: g(r, FTYPE), FTYPE)),
    ]


def _coverage(key_pres, key_sub):
    grp = defaultdict(lambda: {"d": set(), "f": set(), "v": 0, "p": 0, "i": 0})
    for r in sub:
        k = str(r.get(key_sub))
        e = grp[k]
        e["d"].add(str(r.get(DISTR)))
        e["f"].add(str(r.get(FACIL)))
        e["v"] += 1
    for r in pres:
        k = str(key_pres(r))
        grp[k]["p"] += 1
        grp[k]["i"] += r["_n_items"]
    body = [[k, len(e["d"]), len(e["f"]), e["v"], e["p"], e["i"],
             pct(e["i"], e["p"])] for k, e in grp.items()]
    body.sort(key=lambda r: -r[4])
    return body


def facilities():
    grp = defaultdict(lambda: {"v": 0, "p": 0, "i": 0, "na": 0, "known": 0,
                               "prov": "", "dist": "", "type": "", "coll": set()})
    for r in sub:
        e = grp[str(r.get(FACIL))]
        e["v"] += 1
        e["prov"], e["dist"], e["type"] = r.get(PROV), r.get(DISTR), r.get(FTYPE)
        e["coll"].add(str(r.get(COLLECTOR)))
    for r in pres:
        grp[str(g(r, FACIL))]["p"] += 1
    for r in item:
        e = grp[str(g(r, FACIL))]
        e["i"] += 1
        if r["_avail"] is not None:
            e["known"] += 1
            if not r["_avail"]:
                e["na"] += 1
    body = [[k, e["prov"], e["dist"], e["type"], e["v"], e["p"], e["i"],
             pct(e["i"], e["p"]), e["na"], pct(e["na"], e["known"]), len(e["coll"])]
            for k, e in grp.items()]
    body.sort(key=lambda r: (-r[6], r[0]))
    return [
        Table("Every facility", ["Facility", "Province", "District", "Type", "Visits",
                                 "Prescriptions", "Items", "Items / prescription",
                                 "Items not available", "% not available", "Collectors"],
              body, chart=("h", 0, 6, 30),
              total=["Total", "", "", "", N_SUB, N_PRES, N_ITEM,
                     pct(N_ITEM, N_PRES), ITEM_NA, pct(ITEM_NA, ITEM_AV + ITEM_NA), ""]),
        Table("Prescriptions per facility visit", STAT_HDR,
              [stats_row("Prescriptions per visit", [r["_n_pres"] for r in sub]),
               stats_row("Medicine items per visit", [r["_n_items"] for r in sub])],
              note="One submission = one facility visit by one data collector; a visit "
                   "carries 1 to 22 prescriptions."),
    ]


def prescribers():
    grp = defaultdict(list)
    for r in pres:
        grp[str(r.get(PRESCRIBER))].append(r)
    body = []
    for k, rs in grp.items():
        its = [i for p in rs for i in items_of.get(id(p), [])]
        known = [i for i in its if i["_avail"] is not None]
        na = sum(1 for i in known if not i["_avail"])
        body.append([k, len(rs), pct(len(rs), N_PRES), len(its),
                     pct(len(its), len(rs)), na, pct(na, len(known)),
                     pct(sum(1 for p in rs if p["_abx"]), len(rs)),
                     pct(sum(1 for p in rs if p["_inj"]), len(rs))])
    body.sort(key=lambda r: -r[1])
    return [
        freq_table("Prescriber category", pres, PRESCRIBER, "Prescriber category",
                   chart=("h", 0, 1, 12)),
        freq_table("Allied health professional type (where specified)", pres, ALLIED,
                   "Allied professional type", skip_blank=True,
                   note="Only asked when prescriber category is 'Allied Health "
                        "Professional'."),
        Table("Prescribing profile by prescriber category",
              ["Prescriber category", "Prescriptions", "% of prescriptions", "Items",
               "Items / prescription", "Items not available", "% items not available",
               "% with an antibiotic", "% with an injection"], body,
              note="Read the small categories with care - Specialist and Midwife have "
                   "a single prescription each, so their percentages are 0% or 100% by "
                   "construction."),
        avail_by("Availability by prescriber category", item,
                 lambda r: (r["_pres"] or {}).get(PRESCRIBER), "Prescriber category"),
    ]


def prescription_profile():
    declared = [r["_n_declared"] for r in pres]
    return [
        freq_table("Number of medicines per prescription", pres,
                   lambda r: int(r["_n_items"]), "Medicines on the prescription",
                   chart=("v", 0, 1, 12),
                   note="Counted from the medicine_item rows actually present."),
        Table("Medicines per prescription - distribution", STAT_HDR, [
            stats_row("Items counted from medicine_item", [r["_n_items"] for r in pres]),
            stats_row("'Number of medicines prescribed' as recorded", declared),
        ], note=("The two rows agree on every prescription in this export (the declared "
                 f"field totals exactly the {N_ITEM:,} medicine_item rows)."
                 if not any(r["_count_mismatch"] for r in pres) else
                 f"{sum(1 for r in pres if r['_count_mismatch'])} prescriptions declare "
                 f"a medicine count that differs from the number of medicine_item rows "
                 f"actually present - see Data quality.")),
        Table("Polypharmacy", ["Threshold", "Prescriptions", "% of prescriptions"], [
            [f">= {k} medicines", sum(1 for r in pres if r["_n_items"] >= k),
             pct(sum(1 for r in pres if r["_n_items"] >= k), N_PRES)]
            for k in (2, 3, 4, 5, 6)
        ], note="WHO flags >=5 medicines per encounter as a polypharmacy marker."),
        freq_table("Item position used on the form", item, ITEM_NO,
                   "item_number",
                   note="How deep the repeat group was filled; mirrors the medicines-"
                        "per-prescription distribution."),
        Table("Quantity prescribed", STAT_HDR,
              [stats_row("Qty prescribed (all items)", [num(r.get(QTY)) for r in item])],
              note="Quantity is recorded in mixed units (tablets, ml, bottles) with no "
                   "unit field, so the max of 420 is not comparable across dosage "
                   "forms - use it per medicine, not in aggregate."),
        Table("Quantity prescribed - top medicines", STAT_HDR,
              _qty_by_med(15),
              note="Limited to medicines with at least 10 items, where a median is "
                   "meaningful."),
    ]


def _qty_by_med(top):
    grp = defaultdict(list)
    for r in item:
        grp[r["_med"]].append(num(r.get(QTY)))
    rows = [(k, v) for k, v in grp.items() if len([x for x in v if x is not None]) >= 10]
    rows.sort(key=lambda kv: -len(kv[1]))
    return [stats_row(k, v) for k, v in rows[:top]]


def medicines():
    return [
        freq_table("Top 40 medicines prescribed", item, "_med", "Medicine",
                   top=40, chart=("h", 0, 1, 25)),
        freq_table("On the Essential Medicines List", item, EML, "On EML",
                   chart=("v", 0, 1, 6)),
        freq_table("Medicines not found on the EML list (free text)", item, MED_FREE,
                   "Free-text medicine name", skip_blank=True,
                   note="Only asked when the medicine is not on the EML, or the "
                        "collector could not find it in the picker."),
        freq_table("Name type - generic (INN) vs brand", item, NAME_TYPE, "Name type",
                   chart=("v", 0, 1, 6)),
        freq_table("Dosage form [derived from the medicine code]", item, "_form",
                   "Dosage form", chart=("v", 0, 1, 12),
                   note="Parsed from the medicine code (paracetamol_tab_500_mg -> tab). "
                        "The form collects no dosage-form field."),
        Table("Brand-name prescribing by facility type",
              ["Facility type", "Items", "Brand", "% brand"],
              _brand_by(lambda r: g(r, FTYPE))),
        Table("Brand-name prescribing by prescriber category",
              ["Prescriber category", "Items", "Brand", "% brand"],
              _brand_by(lambda r: (r["_pres"] or {}).get(PRESCRIBER))),
    ]


def _brand_by(key):
    grp = defaultdict(list)
    for r in item:
        grp[str(key(r))].append(r)
    body = [[k, len(rs), sum(1 for r in rs if str(r.get(NAME_TYPE)).strip() == "Brand"),
             pct(sum(1 for r in rs if str(r.get(NAME_TYPE)).strip() == "Brand"), len(rs))]
            for k, rs in grp.items()]
    body.sort(key=lambda r: -r[1])
    return body


def drug_classes():
    return [
        freq_table("Drug class as recorded", item, "_class", "Drug class",
                   top=50, chart=("h", 0, 1, 25)),
        freq_table("Drug class by EML section number [derived]", item, "_class_no",
                   "EML section", top=50,
                   note="Normalised from the class text, so '25.2 Parenteral route' and "
                        "'25.2. PARENTERAL ROUTE' collapse into one section."),
        freq_table("Drug class free text (not on EML)", item, CLASS_FREE,
                   "Free-text drug class", skip_blank=True),
        Table("Therapeutic groups of interest [derived]",
              ["Group", "Items", "% of items", "Prescriptions touched",
               "% of prescriptions"], [
            ["Systemic antibacterial (EML 11.2, 11.3)",
             sum(1 for r in item if r["_abx"]),
             pct(sum(1 for r in item if r["_abx"]), N_ITEM),
             sum(1 for r in pres if r["_abx"]), pct(sum(1 for r in pres if r["_abx"]), N_PRES)],
            ["Topical / ophthalmic anti-infective (12.1.2, 14.1)",
             sum(1 for r in item if is_topical_antiinfective(r)),
             pct(sum(1 for r in item if is_topical_antiinfective(r)), N_ITEM), "", ""],
            ["Analgesic / NSAID (EML 2.x, 18.2)",
             sum(1 for r in item if r["_class_no"].startswith(("2.", "18.2"))),
             pct(sum(1 for r in item if r["_class_no"].startswith(("2.", "18.2"))), N_ITEM),
             "", ""],
            ["Injection or infusion [from dosage form]",
             sum(1 for r in item if r["_is_inj"]),
             pct(sum(1 for r in item if r["_is_inj"]), N_ITEM),
             sum(1 for r in pres if r["_inj"]), pct(sum(1 for r in pres if r["_inj"]), N_PRES)],
        ], note="Group membership is derived from the EML section number on each item "
                "(and, for injections, from the dosage form in the medicine code), not "
                "from a field the form collects."),
        avail_by("Availability by drug class", item, "_class", "Drug class",
                 top=30, min_items=5),
    ]


def availability():
    unserved = [r for r in item if r["_avail"] is False]
    detail = [[g(r, FACIL), g(r, DISTR), g(r, FTYPE),
               as_date((r["_pres"] or {}).get(DATE)),
               (r["_pres"] or {}).get(PRESCRIBER), r["_med"], r["_class"], r["_form"],
               num(r.get(QTY)), r.get(ORDERED), r.get(SUPPLIED),
               num(r.get(UNIT_PRICE)), num(r.get(COST_UNSERVED)),
               r.get(RMS_REG), r.get(RMS_CEN), r.get(DIST_NAT)] for r in unserved]
    detail.sort(key=lambda r: (str(r[0]), str(r[5])))
    return [
        Table("Item availability at the facility on the prescription date",
              ["Status", "Items", "% of items"], [
            ["Available", ITEM_AV, pct(ITEM_AV, N_ITEM)],
            ["Not available", ITEM_NA, pct(ITEM_NA, N_ITEM)],
            ["Not recorded", ITEM_UNK, pct(ITEM_UNK, N_ITEM)],
        ], total=["Total", N_ITEM, 1.0], chart=("v", 0, 1, 6)),

        Table("Prescription-level service status",
              ["Status", "Prescriptions", "% of prescriptions"], [
            ["Fully served (all items available)",
             sum(1 for r in pres if r["_fully_served"]),
             pct(sum(1 for r in pres if r["_fully_served"]), N_PRES)],
            ["Partially served (some items not available)",
             sum(1 for r in pres if r["_any_unserved"] and r["_n_avail"] > 0),
             pct(sum(1 for r in pres if r["_any_unserved"] and r["_n_avail"] > 0), N_PRES)],
            ["Fully unserved (no item available)",
             sum(1 for r in pres if r["_n_items"] and r["_n_avail"] == 0),
             pct(sum(1 for r in pres if r["_n_items"] and r["_n_avail"] == 0), N_PRES)],
        ], total=["Total", N_PRES, 1.0],
           note="Recomputed from medicine_item - the form's own n_items_available / "
                "n_items_unavailable / pct_items_unserved fields are empty in this "
                "export."),

        Table("Every unserved medicine item (all %d)" % len(unserved),
              ["Facility", "District", "Facility type", "Prescription date",
               "Prescriber", "Medicine", "Drug class", "Form", "Qty",
               "Ordered that month?", "Supplied?", "Unit price (RWF)",
               "Cost unserved (RWF)", "At RMS Regional?", "At RMS Central?",
               "Distribution issue"], detail,
              note="Small enough to list in full rather than summarise. Every unserved "
                   "item in the dataset is on this table."),

        freq_table("Was the medicine ordered that month? (unserved items only)",
                   unserved, ORDERED, "Ordered", skip_blank=True),
        freq_table("Supplied fully / partially / not at all (unserved items only)",
                   unserved, SUPPLIED, "Supplied", skip_blank=True),
        freq_table("Available at RMS Regional (unserved items only)", unserved,
                   RMS_REG, "At RMS Regional", skip_blank=True),
        freq_table("Available at RMS Central (unserved items only)", unserved,
                   RMS_CEN, "At RMS Central", skip_blank=True),
        freq_table("Cause of non-availability (unserved items only)", unserved,
                   DIST_NAT, "Cause", skip_blank=True),

        Table("Financial loss from unserved items", ["Measure", "Value"], [
            ["Unserved items with a unit price", sum(1 for r in unserved
                                                     if num(r.get(UNIT_PRICE)) is not None)],
            ["Total cost of unserved items (RWF)", COST],
            ["Mean cost per unserved item (RWF)", pct(COST, len(unserved))],
        ], note="Cost is only collected for items recorded as not available, so this is "
                "the loss on those items - not a facility-wide figure."),

        avail_by("Availability by province", item, lambda r: g(r, PROV), "Province"),
        avail_by("Availability by district", item, lambda r: g(r, DISTR), "District"),
        avail_by("Availability by facility type", item, lambda r: g(r, FTYPE),
                 "Facility type"),
        avail_by("Availability by facility (facilities with >=10 items)", item,
                 lambda r: g(r, FACIL), "Facility", min_items=10),
        avail_by("Availability by dosage form [derived]", item, "_form", "Dosage form"),
        avail_by("Availability of the 30 most-prescribed medicines", item, "_med",
                 "Medicine", top=30, min_items=5),
    ]


def who_indicators():
    n_inn = sum(1 for r in item if str(r.get(NAME_TYPE)).strip() == "INN")
    n_eml = sum(1 for r in item if yes(r.get(EML)))
    n_nt = sum(1 for r in item if not blank(r.get(NAME_TYPE)))
    return [
        Table("WHO / INRUD core prescribing indicators",
              ["Indicator", "Numerator", "Denominator", "Value", "WHO optimum"], [
            ["Average number of medicines per encounter", N_ITEM, N_PRES,
             N_ITEM / N_PRES if N_PRES else None, "1.6 - 1.8"],
            ["% medicines prescribed by generic name (INN)", n_inn, n_nt,
             pct(n_inn, n_nt), "100%"],
            ["% encounters with an antibiotic prescribed",
             sum(1 for r in pres if r["_abx"]), N_PRES,
             pct(sum(1 for r in pres if r["_abx"]), N_PRES), "20 - 26.8%"],
            ["% encounters with an injection prescribed",
             sum(1 for r in pres if r["_inj"]), N_PRES,
             pct(sum(1 for r in pres if r["_inj"]), N_PRES), "13.4 - 24.1%"],
            ["% medicines prescribed from the EML", n_eml, N_ITEM,
             pct(n_eml, N_ITEM), "100%"],
        ], chart=("h", 0, 3, 6),
           note="Antibiotic = EML section 11.2/11.3; injection = dosage form parsed "
                "from the medicine code. Neither is a field the form collects, so both "
                "are derived - see the Drug classes sheet. WHO optimum values are the "
                "standard INRUD reference ranges, included for orientation only; this "
                "is a facility-stock study, not a rational-use survey, and the "
                "denominator is prescriptions sampled for stock assessment."),

        Table("Core indicators by province",
              ["Province", "Prescriptions", "Medicines / encounter", "% generic (INN)",
               "% with antibiotic", "% with injection", "% from EML"],
              _who_by(lambda r: g(r, PROV), lambda r: g(r, PROV))),
        Table("Core indicators by facility type",
              ["Facility type", "Prescriptions", "Medicines / encounter",
               "% generic (INN)", "% with antibiotic", "% with injection", "% from EML"],
              _who_by(lambda r: g(r, FTYPE), lambda r: g(r, FTYPE))),
        Table("Core indicators by facility (>=5 prescriptions)",
              ["Facility", "Prescriptions", "Medicines / encounter", "% generic (INN)",
               "% with antibiotic", "% with injection", "% from EML"],
              _who_by(lambda r: g(r, FACIL), lambda r: g(r, FACIL), min_pres=5)),
    ]


def _who_by(key_pres, key_item, min_pres=1):
    gp, gi = defaultdict(list), defaultdict(list)
    for r in pres:
        gp[str(key_pres(r))].append(r)
    for r in item:
        gi[str(key_item(r))].append(r)
    body = []
    for k, ps in gp.items():
        its = gi.get(k, [])
        if len(ps) < min_pres:
            continue
        nt = [r for r in its if not blank(r.get(NAME_TYPE))]
        body.append([k, len(ps), pct(len(its), len(ps)),
                     pct(sum(1 for r in nt if str(r.get(NAME_TYPE)).strip() == "INN"), len(nt)),
                     pct(sum(1 for r in ps if r["_abx"]), len(ps)),
                     pct(sum(1 for r in ps if r["_inj"]), len(ps)),
                     pct(sum(1 for r in its if yes(r.get(EML))), len(its))])
    body.sort(key=lambda r: -r[1])
    return body


def time_trend():
    by_day = defaultdict(lambda: {"p": 0, "i": 0, "na": 0})
    for r in pres:
        d = as_date(r.get(DATE))
        if d:
            by_day[d]["p"] += 1
    for r in item:
        d = as_date((r["_pres"] or {}).get(DATE))
        if d:
            by_day[d]["i"] += 1
            if r["_avail"] is False:
                by_day[d]["na"] += 1
    body = [[d.isoformat(), e["p"], e["i"], e["na"]] for d, e in sorted(by_day.items())]
    sub_day = Counter()
    for r in sub:
        d = as_date(r.get(SUB_TIME))
        if d:
            sub_day[d] += 1
    return [
        Table("Prescriptions by date of prescription",
              ["Date of prescription", "Prescriptions", "Medicine items",
               "Items not available"], body, chart=("v", 0, 1, 60),
              total=["Total", sum(r[1] for r in body), sum(r[2] for r in body),
                     sum(r[3] for r in body)],
              note="Date of prescription is what was written on the prescription, so it "
                   "runs well before the submission dates."),
        Table("Submissions by date submitted", ["Date submitted", "Submissions"],
              [[d.isoformat(), n] for d, n in sorted(sub_day.items())],
              chart=("v", 0, 1, 30),
              note=f"All submissions land in a {SUB_SPAN_DAYS}-day window - the "
                   f"collectors entered back-logged paper prescriptions, so submission "
                   f"date carries no information about when care happened."),
        Table("Lag between prescription date and submission", STAT_HDR,
              [stats_row("Days from prescription to submission",
                         [(as_date(r.get(SUB_TIME)) - as_date(r.get(DATE))).days
                          for r in pres
                          if as_date(r.get(DATE)) and as_date(r.get(SUB_TIME))])],
              note="A negative value would mean a prescription dated after it was "
                   "submitted; see Data quality."),
    ]


# Keyword themes for the free-text feedback field. Ordered - first match wins.
# Deliberately literal: these are substring rules over the lower-cased answer,
# including the misspellings the collectors actually typed (OPEMRS, OPNMRS,
# OPENRMS all mean OpenMRS), not an interpretation of what was meant.
# Ordered most-specific first: the first matching rule wins, so a sentence
# carrying real content is never swallowed by a generic positive phrase.
FB_THEMES = [
    # The study's own blind spot, stated by the collectors: a prescriber writes
    # what the pharmacy has, so a medicine the patient had to buy elsewhere
    # never appears as "not available". Bears directly on the unserved rate.
    ("Cannot observe medicines sourced outside the facility",
     ("outside facility", "outside the facility", "find outside",
      "prescribe existing medicines", "outside of hc", "purchased outside",
      "no way to know the products")),
    # Stock record said out-of-stock yet the item was dispensed (or the reverse).
    ("Stock record and dispensing disagree - reconciliation needed",
     ("but it was isuued", "but it was issued", "but dispensed",
      "appears out of stock", "appearing not being in stock", "reconciliat")),
    ("e-Ubuzima quantity does not match the prescription",
     ("e-ubuzima", "e ubuzima", "quantity reported is", "quantity indicated")),
    # Paper register used, usually because the electronic system was
    # unreachable. Some answers name the network, others only the register.
    ("Recorded from paper register / network issue",
     ("network issue", "network", "system down", "system not used",
      "from register", "in register", "treatment done in register")),
    ("Prescription came from OpenMRS / an electronic record",
     ("openmrs", "opemrs", "opnmrs", "openrms", "open mrs")),
    ("Medicine names written as abbreviations", ("abbrevia",)),
    # Typos as typed ("Autpmation", "autamation") and the synonym the health
    # posts used ("digitalise the prescription data recording").
    ("Recommend automating / digitalising records",
     ("automat", "autamat", "autpmat", "digitali", "digitis", "digitiz",
      "electronic system")),
    # Distinct from automating records: the system exists, staff are not using
    # it. Typos as typed ("Enforcment", "Ebuzima" for e-Ubuzima).
    ("Recommend enforcing use of the electronic system",
     ("system usability", "usability enforcement", "usability  .", "enforcment",
      "enforcement ebuzima", "enforcement of system", "system usability "
      "enforcement")),
    ("Recommend timely supply of medicines",
     ("avail medicine", "avail medicines", "avail this medicine", "avail timely",
      "timely", "encourage availability", "avail health", "avail medical",
      "continue availing", "availing health")),
    ("Brand dispensed against generic prescribed (or vice versa)",
     ("brufen prescribed", "prescribed but issued")),
    # "Sodium nasal drops prescribed but not available" - an unserved item
    # reported in free text, worth catching rather than leaving unclassified.
    ("Stock-out / partial dispensing described",
     ("stock out", "stock-out", "stockout", "prescribed but not available",
      "but not available")),
    ("Possible measurement bias in availability", ("bias",)),
    ("Double data entry in the system", ("double data",)),
    ("Polypharmacy concern", ("polypharmacy",)),
    ("Dosing / treatment-duration concern",
     ("tapering", "regressive", "treatment duration", "first dose",
      "follow up", "follow-up")),
    ("Patient identifier missing (UPID)", ("upid",)),
    ("Price taken from insurer tariff, not RMS", ("tariff",)),
    ("Referral / pre-transfer treatment", ("pre-transfer", "transfer treatment")),
    ("Power outage / prescriptions not in RSSB", ("power", "rssb")),
    ("Programmatic medicines not available", ("programmatic",)),
]

# Answers that carry no content. Matched against the WHOLE answer (punctuation
# stripped), so "good" alone is no-issue while "good, but X" falls through to a
# content rule above. Typos included as typed ("no essue").
FB_NO_ISSUE = {
    "ok", "okay", "it's okay", "its okay", "oky", "none", "no", "n/a", "na",
    "nothing", "no issue", "no issues", "no any issue", "no essue", "no challenge",
    "no challenges", "no feedback", "no comment", "no comments", "good", "goo",
    "correct", "available", "fine", "well", "good and well", "well and good",
    "all good", "no problem", "no problems", "taken", "medicine are available",
    "medicines are available", "medicine available", "medicines available",
}
# Longer answers that merely elaborate on "all fine".
FB_NO_ISSUE_PREFIX = ("well and good", "good and well", "it's okay", "its okay")
# A short answer confirming the prescription was filled - "well served",
# "Drugs received", "Rabies vaccine available", "Thiamine available 100mg".
# Only applied when no content rule matched, and only when the answer carries
# no negation, so "prescribed but not available" is never swept in here.
FB_POSITIVE_WORDS = ("available", "served", "received", "issued", "dispensed")
FB_NEGATIONS = ("not ", "no ", "out of", "n't", "lack", "missing", "short")


def fb_theme(text):
    t = " ".join(str(text).split()).strip().lower()
    for label, keys in FB_THEMES:
        if any(k in t for k in keys):
            return label
    bare = t.strip(" .!,;:-")
    if bare in FB_NO_ISSUE or bare.startswith(FB_NO_ISSUE_PREFIX):
        return "No issue reported"
    if (len(bare.split()) <= 5
            and any(w in bare for w in FB_POSITIVE_WORDS)
            and not any(n in bare for n in FB_NEGATIONS)):
        return "No issue reported"
    return "(other - read verbatim)"


def _count_themes():
    """Fill the FB_* counters Findings quotes. Called before the sections are
    built, since fb_theme is defined after the module-level constants."""
    global FB_OUTSIDE, FB_TIMELY, FB_AUTOMATE, FB_OPENMRS, FB_REGISTER
    global FB_ENFORCE
    c = Counter(fb_theme(r.get(FEEDBACK)) for r in pres
                if not blank(r.get(FEEDBACK)))
    FB_OUTSIDE = c.get("Cannot observe medicines sourced outside the facility", 0)
    FB_TIMELY = c.get("Recommend timely supply of medicines", 0)
    FB_AUTOMATE = c.get("Recommend automating / digitalising records", 0)
    FB_OPENMRS = c.get("Prescription came from OpenMRS / an electronic record", 0)
    FB_REGISTER = c.get("Recorded from paper register / network issue", 0)
    FB_ENFORCE = c.get("Recommend enforcing use of the electronic system", 0)


def feedback():
    rows = [r for r in pres if not blank(r.get(FEEDBACK))]
    verbatim = [[g(r, FACIL), g(r, DISTR), g(r, FTYPE),
                 as_date(r.get(DATE)), fb_theme(r.get(FEEDBACK)),
                 str(r.get(FEEDBACK)).strip()] for r in rows]
    verbatim.sort(key=lambda x: (str(x[4]), str(x[0])))
    by_fac = Counter(str(g(r, FACIL)) for r in rows)
    return [
        Table("Free-text feedback - response rate", ["Measure", "Value"], [
            ["Prescriptions with feedback", len(rows)],
            ["% of prescriptions", pct(len(rows), N_PRES)],
            ["Distinct answers", len({str(r.get(FEEDBACK)).strip() for r in rows})],
            ["Facilities that gave any feedback",
             f"{len(by_fac)} of {len(set(str(r[FACIL]) for r in sub))}"],
        ], note="The field was added to the form mid-collection, so a blank is usually "
                "'the question did not exist yet', not 'no comment'."),

        freq_table("Feedback by theme [keyword-coded]", rows,
                   lambda r: fb_theme(r.get(FEEDBACK)), "Theme",
                   chart=("h", 0, 1, 12),
                   note="Coded by literal keyword match over the answer text, including "
                        "the misspellings as typed (OPEMRS / OPNMRS / OPENRMS = "
                        "OpenMRS). Read the verbatim table below before relying on any "
                        "theme count - 46 answers is small enough to read in full."),

        Table("Feedback by facility", ["Facility", "District", "Facility type",
                                       "Answers"],
              sorted([[k, next((g(r, DISTR) for r in rows if str(g(r, FACIL)) == k), ""),
                       next((g(r, FTYPE) for r in rows if str(g(r, FACIL)) == k), ""), n]
                      for k, n in by_fac.items()], key=lambda x: -x[3]),
              note="Feedback is concentrated in a handful of facilities - it reflects "
                   "which collectors used the field, not where problems are."),

        Table("Every feedback answer, verbatim (all %d)" % len(rows),
              ["Facility", "District", "Facility type", "Prescription date",
               "Theme [keyword-coded]", "Answer as typed"], verbatim,
              note="Three answers bear directly on how the availability figures should "
                   "be read: the Ruhengeri L2TH comment that 'either items are available "
                   "or bias in treatment', the Nyamata L2TH note that a patient was "
                   "given 49 of 60 pills 'due to stock out', and the Kigese CS report "
                   "that 'some of programmatic are not available'. The second describes "
                   "a partial dispense on a prescription whose items are all recorded "
                   "as available."),
    ]


def collectors():
    grp = defaultdict(lambda: {"v": 0, "p": 0, "i": 0, "na": 0, "f": set(), "d": set()})
    for r in sub:
        e = grp[str(r.get(COLLECTOR))]
        e["v"] += 1
        e["f"].add(str(r.get(FACIL)))
        e["d"].add(str(r.get(DISTR)))
    for r in pres:
        grp[str(g(r, COLLECTOR))]["p"] += 1
    for r in item:
        e = grp[str(g(r, COLLECTOR))]
        e["i"] += 1
        if r["_avail"] is False:
            e["na"] += 1
    body = [[k, e["v"], len(e["f"]), len(e["d"]), e["p"], e["i"], pct(e["i"], e["p"]),
             e["na"]] for k, e in grp.items()]
    body.sort(key=lambda r: -r[4])
    return [
        Table("Data collectors", ["Data collector", "Visits", "Facilities", "Districts",
                                  "Prescriptions", "Items", "Items / prescription",
                                  "Items not available"], body, chart=("h", 0, 4, 30),
              note="'None' is a genuine gap - the collector-name field is blank on those "
                   "submissions. The device 'username' field reads 'username not found' "
                   "on all 343 rows, so it cannot substitute."),
    ]


def quality():
    ids = [str(r.get(PID)).strip() for r in pres if not blank(r.get(PID))]
    ci = Counter(ids)
    dup_exact = {k: n for k, n in ci.items() if n > 1}
    ci_fold = Counter(s.upper() for s in ids)
    dup_fold = {k: n for k, n in ci_fold.items() if n > 1}
    neg_lag = [r for r in pres
               if as_date(r.get(DATE)) and as_date(r.get(SUB_TIME))
               and as_date(r.get(SUB_TIME)) < as_date(r.get(DATE))]
    class_variants = defaultdict(set)
    for r in item:
        if r["_class_no"]:
            class_variants[r["_class_no"]].add(str(r.get(CLASS)).strip())
    cv = {k: v for k, v in class_variants.items() if len(v) > 1}
    med_variants = defaultdict(set)
    for r in item:
        if not blank(r.get(MED_EML)):
            key = re.sub(r"[_\s]+", "", str(r[MED_EML]).strip().lower())
            med_variants[key].add(str(r[MED_EML]).strip())
    mv = {k: v for k, v in med_variants.items() if len(v) > 1}

    checks = [
        ["Rows read: submissions / prescriptions / items",
         f"{N_SUB} / {N_PRES} / {N_ITEM}", ""],
        ["Prescriptions not matching a submission", ORPHAN_PRES,
         "OK" if not ORPHAN_PRES else "CHECK"],
        ["Medicine items not matching a prescription", ORPHAN_ITEM,
         "OK" if not ORPHAN_ITEM else "CHECK"],
        ["Submissions with no prescription",
         sum(1 for r in sub if r["_n_pres"] == 0),
         "OK" if all(r["_n_pres"] for r in sub) else "CHECK"],
        ["Prescriptions with no medicine item",
         sum(1 for r in pres if r["_n_items"] == 0),
         "OK" if all(r["_n_items"] for r in pres) else "CHECK"],
        ["'Number of medicines prescribed' != medicine_item rows",
         sum(1 for r in pres if r["_count_mismatch"]),
         "OK" if not any(r["_count_mismatch"] for r in pres) else "CHECK"],
        ["Items with availability not recorded", ITEM_UNK,
         "OK" if not ITEM_UNK else "CHECK"],
        ["Items where availability text and flag disagree",
         sum(1 for r in item
             if not blank(r.get(AVAIL)) and num(r.get(AV_FLAG)) is not None
             and yes(r.get(AVAIL)) != (num(r.get(AV_FLAG)) == 1)), "OK"],
        ["Exported n_items_available / n_items_unavailable / pct_items_unserved",
         "empty on all %d rows" % N_PRES, "CHECK"],
        ["Form versions present in one export",
         len(set(str(r.get("__version__")) for r in sub)), "CHECK"],
        ["Submissions with no data-collector name",
         sum(1 for r in sub if blank(r.get(COLLECTOR))), "CHECK"],
        ["Rows with a Kobo validation status",
         sum(1 for r in sub if not blank(r.get("_validation_status"))),
         "CHECK"],
        ["Patient / Prescription ID blank",
         sum(1 for r in pres if blank(r.get(PID))), ""],
        ["Patient / Prescription ID exact duplicates",
         f"{len(dup_exact)} values, {sum(dup_exact.values())} rows",
         "CHECK" if dup_exact else "OK"],
        ["Patient / Prescription ID duplicates ignoring case",
         f"{len(dup_fold)} values, {sum(dup_fold.values())} rows",
         "CHECK" if dup_fold else "OK"],
        ["Prescription dated after its submission", len(neg_lag),
         "OK" if not neg_lag else "CHECK"],
        ["Drug-class sections typed more than one way", len(cv),
         "CHECK" if cv else "OK"],
        ["Medicine codes differing only in separators", len(mv),
         "CHECK" if mv else "OK"],
    ]

    moved = [r for r in sub if r.get("_facility_was")]
    ov_rows = []
    for r in moved:
        mine = [p for p in pres if p["_sub"] is r]
        ov_rows.append([str(r.get("_facility_was")), str(r.get(FACIL)),
                        str(r.get(SUB_TIME))[:10], r.get("_id"),
                        r.get(COLLECTOR), len(mine),
                        sum(p["_n_items"] for p in mine)])
    ov_rows.sort(key=lambda x: (x[0], x[2]))

    return [
        Table("Facility reattributions applied by this script",
              ["Recorded in Kobo as", "Counted here as", "Submitted",
               "Submission ID", "Collector", "Prescriptions", "Medicine items"],
              ov_rows,
              note="Health posts with no entry in the Kobo facility list are captured "
                   "under their parent health centre, and the field team says which "
                   "submissions those are. Each correction is keyed by (facility as "
                   "recorded, submission date) in FACILITY_OVERRIDES at the top of "
                   "analyze_kobo_prescription.py, so it survives a re-fetch. Facility "
                   "type is set to SGHP to match every other health post; district and "
                   "sector are left as recorded, since a post sits inside its parent's "
                   "catchment. Every other sheet in this workbook already reflects "
                   "these changes."
              if ov_rows else
                   "No reattributions are configured, or none matched this export - "
                   "check FACILITY_OVERRIDES if one was expected."),

        Table("Integrity checks", ["Check", "Value", "Flag"], checks,
              note="'CHECK' marks something to look at before publishing figures, not "
                   "necessarily an error."),

        Table("Patient / Prescription ID duplicates (ignoring case)",
              ["ID (upper-cased)", "Rows", "As typed"],
              sorted([[k, n, " | ".join(sorted({s for s in ids if s.upper() == k}))]
                      for k, n in dup_fold.items()], key=lambda r: (-r[1], r[0])),
              note=f"{N_PID_DISTINCT} distinct IDs across {N_PRES} prescriptions. Most "
                   f"duplicates are the same ID typed in different case, so the ID "
                   f"cannot be used as a key until it is normalised."),

        Table("Drug-class sections typed more than one way",
              ["EML section", "Variants as typed"],
              sorted([[k, " | ".join(sorted(v))] for k, v in cv.items()]),
              note="Same section, different spelling/case - these split the class "
                   "frequency table unless normalised."),

        Table("Medicine codes differing only in separators",
              ["Normalised", "Variants as typed"],
              sorted([[k, " | ".join(sorted(v))] for k, v in mv.items()]),
              note="These are the same product entered under two codes, so each one's "
                   "count is understated in the medicine frequency table."),

        missing_table(hdr_sub, sub, "Blank values - sheet 1 (facility visits)"),
        missing_table(hdr_pres, pres, "Blank values - sheet 2 (prescriptions)"),
        missing_table(hdr_item, item, "Blank values - sheet 3 (medicine items)"),

        freq_table("Form version by submission", sub, "__version__", "Form version",
                   note="Three versions in one export. The empty roll-up calculate "
                        "fields are the visible consequence: they were almost certainly "
                        "added to the form after collection had begun."),
        Table("Constant columns (no analytical value in this export)",
              ["Column", "Single value"],
              _constant_cols(),
              note="Columns where every row carries the same value, including empties."),
    ]


def _constant_cols():
    out = []
    for name, hdr, rows in (("sheet 1", hdr_sub, sub), ("sheet 2", hdr_pres, pres),
                            ("sheet 3", hdr_item, item)):
        for h in hdr:
            if h.startswith("_") and h not in (SUB_TIME,):
                pass
            vals = {str(r.get(h)) for r in rows}
            if len(vals) == 1:
                v = next(iter(vals))
                out.append([f"{h} ({name})", "(blank)" if v == "None" else v])
    return out


def findings():
    rate = pct(ITEM_NA, ITEM_AV + ITEM_NA)
    return [
        Table("What the data shows", ["#", "Finding"], [
            [1, f"{N_SUB} facility visits produced {N_PRES} prescriptions and "
                f"{N_ITEM} medicine items across {len(set(r[FACIL] for r in sub))} "
                f"facilities in {len(set(r[DISTR] for r in sub))} districts and all "
                f"{len(set(r[PROV] for r in sub))} provinces - national in scope, "
                f"unlike the earlier Kicukiro-only pilot."],
            [2, f"Only {ITEM_NA} of {N_ITEM} items ({rate:.2%}) were recorded as not "
                f"available, affecting {PRES_ANY_NA} of {N_PRES} prescriptions. Both "
                f"the text answer and the form's own flag agree, so this is what was "
                f"entered - but it should not be reported as a stock-availability "
                f"result. The collectors' own feedback identifies the mechanism: "
                f"{FB_OUTSIDE} answers say it is impossible to see what patients had to "
                f"buy elsewhere ('prescribers prescribe existing medicines within "
                f"facility and it's hard to determine any medicine patient had to find "
                f"outside'), one names the problem directly ('either items are available "
                f"or bias in treatment'), and one records a patient given 49 of 60 pills "
                f"'due to stock out' on a prescription whose items are all marked "
                f"available. The instrument measures what was dispensed against what was "
                f"prescribed at the same facility, so an unavailable medicine is largely "
                f"absent from the denominator by construction."],
            [3, f"Total recorded financial loss from unserved items is {COST:,.0f} RWF, "
                f"from the {ITEM_NA} priced items only."],
            [4, f"Prescribing is dominated by nurses ({sum(1 for r in pres if str(r.get(PRESCRIBER)).strip() == 'Nurse')} "
                f"of {N_PRES} prescriptions) and by analgesics - paracetamol and "
                f"ibuprofen alone account for "
                f"{sum(1 for r in item if r['_med'].startswith(('paracetamol', 'ibuprofen')))} "
                f"of {N_ITEM} items."],
            [5, f"Generic (INN) prescribing is high at "
                f"{pct(sum(1 for r in item if str(r.get(NAME_TYPE)).strip() == 'INN'), N_ITEM):.1%} "
                f"and EML adherence near-total at "
                f"{pct(sum(1 for r in item if yes(r.get(EML))), N_ITEM):.1%}. Mean "
                f"medicines per encounter is {N_ITEM / N_PRES:.2f}, above the WHO "
                f"reference range of 1.6-1.8."],
            [6, f"The form's own unserved roll-ups (n_items_available, "
                f"n_items_unavailable, pct_items_unserved) are empty on all {N_PRES} "
                f"rows. {N_VERSIONS} form versions appear in one export, so they were "
                f"likely added mid-collection. Everything unserved in this workbook is "
                f"recomputed from medicine_item."],
            [7, f"Patient / Prescription ID is not yet a usable key: {N_PID_DISTINCT} "
                f"distinct values across {N_PRES} rows, with duplicates that differ "
                f"only in case, plus format outliers. Normalise before any "
                f"patient-level linkage."],
            [8, "Two coding inconsistencies inflate the long tail: EML section 25.2 is "
                "typed both '25.2 Parenteral route' and '25.2. PARENTERAL ROUTE', and "
                "sodium chloride appears under two codes differing by one underscore."],
            [9, f"Collection was still open at export time - all {N_SUB} submissions "
                f"landed in a {SUB_SPAN_DAYS}-day window covering prescriptions written "
                f"over {PRES_SPAN_DAYS} days. Treat every total here as an interim cut."],
            [10, f"{N_FEEDBACK} prescriptions carry free-text feedback. Beyond the "
                 f"availability caveat above, the recurring operational themes are "
                 f"requests to supply medicines more promptly ({FB_TIMELY}) and to "
                 f"automate or digitalise records ({FB_AUTOMATE}); prescriptions arriving "
                 f"from OpenMRS ({FB_OPENMRS}); calls to enforce actual use of the "
                 f"electronic system ({FB_ENFORCE}); and data captured from the paper "
                 f"register because the network or system was down ({FB_REGISTER}) - the "
                 f"last of these means a share of the dataset was transcribed rather "
                 f"than read from the electronic record."],
            [11, "Two answers report the stock record and the dispensing record "
                 "disagreeing - cloxacillin and diclofenac injections showed as out of "
                 "stock yet were dispensed - and two report e-Ubuzima quantities not "
                 "matching the prescription (one tablet recorded where fifteen were "
                 "prescribed). Both are reconciliation problems between systems, and "
                 "both cut against taking either system's stock figure at face value."],
        ], note="Derived indicators (antibiotic, injection, dosage form) are computed "
                "from the EML section number and the medicine code, since the form "
                "collects neither dosage form nor a therapeutic-group field."),
    ]


# ---------------------------------------------------------------
# the three data sheets, with general information pulled down
# ---------------------------------------------------------------
GEN = [("Submission ID", "_id"), ("Submitted at", SUB_TIME), ("Visit date", VISIT_DATE),
       ("Data collector", COLLECTOR), ("Province", PROV), ("District", DISTR),
       ("Sector", SECTOR), ("Facility", FACIL), ("Facility type", FTYPE)]

PRES_CTX = [("Prescription date", DATE), ("Patient / Prescription ID", PID),
            ("Prescriber category", PRESCRIBER),
            ("Medicines on prescription", N_MED)]


def own_cols(hdr, skip):
    return [h for h in hdr if h not in skip]


def data_sheets():
    """(sheet name, note, headers, row-builder) for each of the three levels."""
    gen_h = [f"{lab} [general info]" for lab, _ in GEN]

    s1_cols = own_cols(hdr_sub, set())
    s1 = ("1 Facility visits",
          f"{N_SUB} rows - one row = one facility visit (submission). Source sheet "
          f"'{S_SUB}', all columns, plus two derived counts.",
          s1_cols + ["Facility as recorded in Kobo [derived]",
                     "Prescriptions on this visit [derived]",
                     "Medicine items on this visit [derived]"],
          lambda r: [r.get(h) for h in s1_cols]
                    + [r.get("_facility_was") or "", r["_n_pres"], r["_n_items"]])

    s2_cols = own_cols(hdr_pres, {c for _, c in GEN})
    s2 = ("2 Prescriptions",
          f"{N_PRES} rows - one row = one prescription. Source sheet '{S_PRES}' with "
          f"the general information from sheet 1 pulled down onto every row, plus "
          f"recomputed roll-ups.",
          gen_h + s2_cols + ["Items counted [derived]", "Items available [derived]",
                             "Items not available [derived]", "% items unserved [derived]",
                             "Fully served [derived]", "Has antibiotic [derived]",
                             "Has injection [derived]"],
          lambda r: [g(r, c) for _, c in GEN] + [r.get(h) for h in s2_cols]
                    + [r["_n_items"], r["_n_avail"], r["_n_unavail"], r["_pct_unserved"],
                       "Yes" if r["_fully_served"] else "No",
                       "Yes" if r["_abx"] else "No", "Yes" if r["_inj"] else "No"])

    s3_cols = own_cols(hdr_item, {c for _, c in GEN})
    s3 = ("3 Medicine items",
          f"{N_ITEM} rows - one row = one medicine on one prescription. Source sheet "
          f"'{S_ITEM}' with the general information from sheet 1 AND the prescription "
          f"context from sheet 2 pulled down onto every row, plus derived coding.",
          gen_h + [f"{lab} [from prescription]" for lab, _ in PRES_CTX] + s3_cols
          + ["Medicine resolved [derived]", "Drug class resolved [derived]",
             "EML section [derived]", "Dosage form [derived]", "Injection [derived]",
             "Systemic antibacterial [derived]", "Available [derived]"],
          lambda r: [g(r, c) for _, c in GEN]
                    + [(r["_pres"] or {}).get(c) for _, c in PRES_CTX]
                    + [r.get(h) for h in s3_cols]
                    + [r["_med"], r["_class"], r["_class_no"], r["_form"],
                       "Yes" if r["_is_inj"] else "No", "Yes" if r["_abx"] else "No",
                       "" if r["_avail"] is None else ("Yes" if r["_avail"] else "No")])

    return [(s1, sub), (s2, pres), (s3, item)]


SECTIONS = [
    ("Overview", overview),
    ("Geography", geography),
    ("Facilities", facilities),
    ("Prescribers", prescribers),
    ("Prescription profile", prescription_profile),
    ("Medicines", medicines),
    ("Drug classes", drug_classes),
    ("Availability & unserved", availability),
    ("WHO indicators", who_indicators),
    ("Time", time_trend),
    ("Field feedback", feedback),
    ("Data collectors", collectors),
    ("Data quality", quality),
    ("Findings", findings),
]

# ---------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="E6E6E6")
GEN_FILL = PatternFill("solid", fgColor="DCE9F5")
DER_FILL = PatternFill("solid", fgColor="E8F0DC")
TITLE_FONT = Font(bold=True, size=12)
HDR_FONT = Font(bold=True, size=10)
NOTE_FONT = Font(size=9, italic=True, color="555555")
DATE_FMT = "yyyy-mm-dd"


def sheet_name(s):
    return s[:31].replace("/", "-").replace("[", "(").replace("]", ")")


def write_data_sheet(wb, spec, rows):
    name, note, headers, build = spec
    ws = wb.create_sheet(sheet_name(name))
    ws.freeze_panes = "A4"
    ws.cell(row=1, column=1, value=name).font = Font(bold=True, size=14)
    c = ws.cell(row=2, column=1, value=note)
    c.font = NOTE_FONT
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=min(len(headers), 14))
    ws.row_dimensions[2].height = 28

    widths = {}
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=ci, value=str(h))
        c.font = HDR_FONT
        c.fill = (GEN_FILL if "[general info]" in h or "[from prescription]" in h
                  else DER_FILL if "[derived]" in h else HDR_FILL)
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="top")
        widths[ci] = min(max(12, len(str(h).split("[")[0]) + 2), 34)

    for ri, r in enumerate(rows, start=4):
        for ci, v in enumerate(build(r), start=1):
            if isinstance(v, str) and len(v) > 250:
                v = v[:250] + "..."
            c = ws.cell(row=ri, column=ci, value=v)
            c.alignment = Alignment(vertical="top")
            if isinstance(v, (dt.date, dt.datetime)):
                c.number_format = DATE_FMT
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                c.number_format = "0.0%" if "%" in str(headers[ci - 1]) else "#,##0.##"

    for ci, w in widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.auto_filter.ref = (f"A3:{get_column_letter(len(headers))}"
                          f"{3 + len(rows)}")
    return len(rows), len(headers)


def write_section(wb, name, tables):
    ws = wb.create_sheet(sheet_name(name))
    ws.sheet_view.showGridLines = False
    row = 1
    ws.cell(row=row, column=1, value=name).font = Font(bold=True, size=14)
    row += 2

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
            widths[ci] = max(widths.get(ci, 10), min(len(str(h)) + 2, 44))
        row += 1

        body = list(t.rows) + ([t.total] if t.total else [])
        for ri, r in enumerate(body):
            is_total = t.total is not None and ri == len(body) - 1
            for ci, v in enumerate(r, start=1):
                if isinstance(v, str) and len(v) > 300:
                    v = v[:300] + "..."
                c = ws.cell(row=row, column=ci, value=v)
                c.border = BORDER
                c.alignment = Alignment(vertical="top",
                                        wrap_text=isinstance(v, str) and len(str(v)) > 45)
                if is_total:
                    c.font = Font(bold=True)
                if isinstance(v, (dt.date, dt.datetime)):
                    c.number_format = DATE_FMT
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    c.number_format = "0.0%" if (ci - 1) in t.pct else "#,##0.##"
                widths[ci] = max(widths.get(ci, 10), min(len(str(v)) + 2, 48))
            row += 1

        if t.note:
            c = ws.cell(row=row, column=1, value="Note: " + t.note)
            c.font = NOTE_FONT
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=row, start_column=1, end_row=row,
                           end_column=max(len(t.headers), 3))
            ws.row_dimensions[row].height = 13 * (1 + len(t.note) // 100)
            row += 1
        row += 1

    for ci, w in widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w


def add_charts(wb, built):
    ws = wb.create_sheet("Charts")
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1, value="Charts").font = Font(bold=True, size=14)
    ws.cell(row=2, column=1,
            value="Each chart is a native Excel chart over the data block to its left - "
                  "click to edit, restyle or copy into a deck.").font = NOTE_FONT
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 14

    row, n = 4, 0
    for section, tables in built:
        for t in tables:
            pairs = t.chart_data()
            if not pairs:
                continue
            title = t.title if not t.chart[3] else f"{t.title} (top {len(pairs)})"
            ws.cell(row=row, column=1, value=title).font = TITLE_FONT
            ws.cell(row=row, column=2, value=section).font = NOTE_FONT
            h = row + 1
            ws.cell(row=h, column=1, value=str(t.headers[t.chart[1]])).font = HDR_FONT
            ws.cell(row=h, column=2, value=str(t.headers[t.chart[2]])).font = HDR_FONT
            for i, (lab, val) in enumerate(pairs, start=1):
                ws.cell(row=h + i, column=1, value=lab)
                cc = ws.cell(row=h + i, column=2, value=val)
                cc.number_format = "0.0%" if t.chart[2] in t.pct else "#,##0.##"

            ch = BarChart()
            ch.type = "col" if t.chart[0] == "v" else "bar"
            ch.style = 2
            ch.title = title
            ch.y_axis.title = str(t.headers[t.chart[2]])
            ch.x_axis.title = str(t.headers[t.chart[1]])
            ch.legend = None
            ch.gapWidth = 40
            data = Reference(ws, min_col=2, min_row=h, max_row=h + len(pairs))
            cats = Reference(ws, min_col=1, min_row=h + 1, max_row=h + len(pairs))
            ch.add_data(data, titles_from_data=True)
            ch.set_categories(cats)
            ch.height = max(7.5, min(0.62 * len(pairs) + 3.5, 18))
            ch.width = 20
            ws.add_chart(ch, f"D{row}")
            n += 1
            row = max(h + len(pairs) + 2, h + int(ch.height / 0.53) + 2)
    return n


def write_contents(wb, built, data_info):
    ws = wb.create_sheet("Contents", 0)
    ws.sheet_view.showGridLines = False
    ws.cell(row=1, column=1, value="Prescription study - full analysis").font = Font(bold=True, size=15)
    ws.cell(row=2, column=1,
            value=f"Source: {SRC} (KoboToolbox export)   |   "
                  f"{N_SUB} visits, {N_PRES} prescriptions, {N_ITEM} medicine items   |   "
                  f"built {dt.date.today().isoformat()}").font = NOTE_FONT
    ws.cell(row=4, column=1, value="Sheet").font = HDR_FONT
    ws.cell(row=4, column=2, value="Rows").font = HDR_FONT
    ws.cell(row=4, column=3, value="What it holds").font = HDR_FONT
    rows = [("Contents", "", "This page.")]
    for (name, note, _, _), nrow, ncol in data_info:
        rows.append((sheet_name(name), nrow, f"{ncol} columns. {note}"))
    for section, tables in built:
        rows.append((sheet_name(section), len(tables),
                     f"{len(tables)} tables: " + "; ".join(t.title for t in tables)[:220]))
    rows.append(("Charts", "", "Native Excel charts for every charted table."))
    for i, (a, b, c) in enumerate(rows, start=5):
        ws.cell(row=i, column=1, value=a).font = Font(bold=True, size=10)
        ws.cell(row=i, column=2, value=b)
        cc = ws.cell(row=i, column=3, value=c)
        cc.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 13 * (1 + len(c) // 95)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 110
    ws.cell(row=len(rows) + 6, column=1,
            value="Column shading on the data sheets: blue = general information pulled "
                  "down from a parent sheet, green = derived in this script, grey = as "
                  "exported by Kobo.").font = NOTE_FONT


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    data_info = []
    for spec, rows in data_sheets():
        nrow, ncol = write_data_sheet(wb, spec, rows)
        data_info.append((spec, nrow, ncol))
        print(f"  {sheet_name(spec[0])}: {nrow} rows x {ncol} cols")

    for (was, when), n in sorted(OVERRIDES_APPLIED.items()):
        print(f"  reattributed {n} submission(s): {was} ({when}) -> "
              f"{FACILITY_OVERRIDES[(was, when)][FACIL]}")
    for key in FACILITY_OVERRIDES:
        if key not in OVERRIDES_APPLIED:
            print(f"  WARNING: override {key} matched nothing in this export")

    _count_themes()
    built = [(name, fn()) for name, fn in SECTIONS]
    for name, tables in built:
        write_section(wb, name, tables)
    print(f"  {len(built)} analysis sheets, "
          f"{sum(len(t) for _, t in built)} tables")

    n = add_charts(wb, built)
    write_contents(wb, built, data_info)
    wb.save(OUT_XLSX)
    print(f"  {n} Excel charts")
    print(f"wrote {OUT_XLSX}")


if __name__ == "__main__":
    main()
