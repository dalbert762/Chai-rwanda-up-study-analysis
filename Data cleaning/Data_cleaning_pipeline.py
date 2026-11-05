"""The data cleaning pipeline: every correction, applied and exported.

    python3 Data_cleaning_pipeline.py
      -> "Data cleaning/" :  00_Cleaning_log.xlsx
                             01_Stock_assessment_clean.xlsx
                             02_Prescription_clean.xlsx
                             03_CHW_assessment_clean.xlsx

Until now the corrections lived inside the analysis scripts, which is right
for the analysis and wrong for everything else: nobody outside this codebase
could get a cleaned dataset, and no one could see what had been changed
without reading Python. This runs the same rules and writes out the result,
plus a log naming every affected row.

The analysis scripts remain the source of truth. This file imports their
rules rather than restating them, so the exported data and the dashboards
cannot drift apart - if a correction changes there, it changes here on the
next run.

WHAT IS AND IS NOT CLEANING
Applied to the data, because the recorded value is wrong:
    facility reattributions, drug class corrections, duplicate rows,
    non-finite calculate fields, level-label grouping.
NOT applied, because the value is right and the question is analytical:
    the EML-applicable denominator, the Method 3 strength collapse, the
    exclusion of facilities with no digital system from inventory accuracy.
Those three are decisions about how to COUNT, not about what was recorded,
and burning them into a dataset would stop anyone computing it another way.
They are listed in the log for completeness and marked "analysis rule".
"""

import os as _os
import sys as _sys

# This file now lives in "Data cleaning/", beside its own outputs, while the
# modules it imports live in analytics/. Add that folder to the path before
# anything else so the script runs from wherever it is kept.
_HERE = _os.path.dirname(_os.path.abspath(__file__))
_ANALYTICS = _os.path.join(_os.path.dirname(_HERE), "analytics")
if _ANALYTICS not in _sys.path:
    _sys.path.insert(0, _ANALYTICS)

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
import eml_list as EML
import emr_systems as EMRSYS
import facility_levels as FL
import molecules as MOL

OUTDIR = os.path.join(B.ROOT, "Data cleaning")

HDR_FONT = Font(bold=True, size=9, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="2A68BD")
TITLE_FONT = Font(bold=True, size=11)
WARN_FILL = PatternFill("solid", fgColor="FBE3BA")
INFO_FILL = PatternFill("solid", fgColor="E8EEF8")
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def s(v):
    return "" if v is None else str(v).strip()


def ensure_dir():
    os.makedirs(OUTDIR, exist_ok=True)


def out(name):
    path = os.path.join(OUTDIR, name)
    # Box resets synced files to read-only; make them writable before saving.
    if os.path.exists(path):
        os.chmod(path, 0o644)
    return path


def write(wb, name):
    wb.save(out(name))
    return name


def table(ws, start, title, headers, rows, widths=None, fill=None):
    r = start
    if title:
        ws.cell(row=r, column=1, value=title).font = TITLE_FONT
        r += 1
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center",
                                horizontal="center")
    ws.row_dimensions[r].height = 34
    hdr = r
    r += 1
    for row in rows:
        for ci, v in enumerate(row, start=1):
            c = ws.cell(row=r, column=ci, value=v)
            c.border = BORDER
            c.alignment = Alignment(vertical="top",
                                    wrap_text=isinstance(v, str)
                                    and len(str(v)) > 40)
            if isinstance(v, float):
                c.number_format = "0.0%"
            elif isinstance(v, int):
                c.number_format = "#,##0"
            if fill:
                c.fill = fill
        r += 1
    if rows:
        ws.freeze_panes = f"A{hdr + 1}"
        ws.auto_filter.ref = (f"A{hdr}:{get_column_letter(len(headers))}"
                              f"{r - 1}")
    for ci, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    return r + 2


def data_sheet(wb, title, headers, rows, widths=None):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 42
    for ri, row in enumerate(rows, start=2):
        for ci, v in enumerate(row, start=1):
            ws.cell(row=ri, column=ci, value=v)
    ws.freeze_panes = "B2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows)+1}"
    for ci, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    return ws


# ---------------------------------------------------------------
# the rules, as they are actually applied
# ---------------------------------------------------------------
RULES = []


def rule(tool, rid, name, kind, rows_affected, detail):
    RULES.append([tool, rid, name, kind, rows_affected, detail])


# ---------------------------------------------------------------
# S6 - EML applicability, corrected against the official list
# ---------------------------------------------------------------
# The form asks the collector whether the medicine is applicable at that
# facility level "as per EML", and the answer is given from memory. The EML
# list is the authority, so column I is checked against it.
#
# ONE DIRECTION ONLY, as instructed: No -> Yes where the list holds the
# medicine at that level. The opposite case - the form says Yes and the list
# does not hold it there - is reported but NOT changed, because a facility
# legitimately managing something outside its level is a real finding and
# overwriting it would hide it.
#
# This is not cosmetic: APPLICABLE is the denominator of the study's
# availability definition, so every row corrected here enters that
# denominator and moves the headline rate.
EML_BY_ING, EML_NAMES = EML.load_eml()
_EML_CACHE = {}


def eml_verdict(med, level):
    k = (med, level)
    if k not in _EML_CACHE:
        _EML_CACHE[k] = EML.applicable(med, level, EML_BY_ING, EML_NAMES)
    return _EML_CACHE[k]


def apply_eml_applicability(rows):
    """-> (corrected, opposite direction, unmatched medicines)."""
    fixed, opposite, unmatched = [], [], Counter()
    for r in rows:
        med, lvl = s(r[M.MED]), s(r["_p"][M.LEVEL])
        verdict, how, levels, prods = eml_verdict(med, lvl)
        r["_eml_levels"] = ", ".join(sorted(levels))
        r["_eml_match"] = how
        r["_eml_products"] = "; ".join(sorted(prods)[:3])
        # analyze_medicine_assessment applies S6 on import and records the
        # old value; keep it rather than blanking it here
        r.setdefault("_applicable_was", "")
        if verdict is None:
            if how == "no name match":
                unmatched[med] += 1
            continue
        cur = M.yes(r[M.APPLICABLE])
        if verdict and not cur:
            r["_applicable_was"] = s(r[M.APPLICABLE]) or "(blank)"
            r[M.APPLICABLE] = "Yes"
            fixed.append(r)
        elif cur and not verdict:
            opposite.append(r)
    return fixed, opposite, unmatched


# ---------------------------------------------------------------
# S7 - available = Yes with a physical quantity of zero
# ---------------------------------------------------------------
# The form asks whether the medicine is available on the visit day, then asks
# for the quantity on the shelf. A row answering Yes and then 0 contradicts
# itself: nothing can be dispensed from an empty shelf, so the medicine was
# not available. The quantity is the harder number to get wrong - it is read
# off a stock card or counted - so it is treated as the true one and the
# availability flag is corrected to No.
#
# Only this direction occurs. There is no row answering No with a quantity
# above zero, so the correction is unambiguous: 173 rows move, none the other
# way.
#
# Like S6 this moves the availability numerator, so it is logged row by row.
def apply_zero_quantity(rows):
    """-> rows whose availability flag was corrected from Yes to No."""
    fixed = []
    for r in rows:
        if M.yes(r[M.AVAIL]) and M.num(r[M.QTY_PHYS]) == 0:
            r["_available_was"] = s(r[M.AVAIL])
            r[M.AVAIL] = "No"
            fixed.append(r)
    return fixed


# ---------------------------------------------------------------
# S8 - the recording system name, normalised
# ---------------------------------------------------------------
# The column is free text and 162 distinct values were typed for what is
# really seven systems - "OpenMRS" alone appears twenty-one different ways.
# Any count grouped on the raw value is meaningless, and 32 rows hold a bare
# number where a system name should be.
#
# Numbers are removed, each name is mapped to one canonical spelling, and a
# row naming several systems is joined in a fixed order so the same
# combination always reads the same way. The raw value is kept beside it, and
# any qualifier the collector added ("only used for ordering") goes in a note
# rather than being dropped - it says something about how far the system is
# really used.
def apply_emr_names(rows):
    """-> (rows changed, rows whose numeric value was removed)."""
    changed, numeric = [], []
    for r in rows:
        raw = s(r[M.EMR])
        name, note = EMRSYS.canonical(raw)
        r["_emr_raw"] = raw
        r["_emr_clean"] = name
        r["_emr_note"] = note
        if not raw:
            continue
        if note == "numeric value removed":
            numeric.append(r)
        if name != raw:
            changed.append(r)
    return changed, numeric


def build_stock():
    """Stock assessment: already de-duplicated and parent-linked on import."""
    sub = M.SUB
    rows = M.MEDROWS
    # S6, S7 and S9 are applied by analyze_medicine_assessment on import, so
    # its lists hold the rows it changed; re-running the rules here finds
    # nothing new and would log zero. The pipeline's own passes still run -
    # they annotate every row (EML match, recording system) for the export.
    eml_new, eml_opposite, eml_unmatched = apply_eml_applicability(rows)
    eml_fixed = list(M.EML_FIXED) + eml_new
    zero_fixed = list(M.ZERO_FIXED) + apply_zero_quantity(rows)
    strength_recoded = list(M.STRENGTH_RECODED)
    emr_changed, emr_numeric = apply_emr_names(rows)
    dropped = M.DUPES_DROPPED
    conflicting = [d for d in dropped if not d["identical"]]

    rule("Stock assessment", "S1", "Rows linked to a parent submission",
         "applied", len(rows),
         "Medicine rows whose parent submission is missing are dropped - "
         "they carry no facility, district or level and cannot be placed.")
    rule("Stock assessment", "S2", "Duplicate facility x medicine rows removed",
         "applied", len(dropped),
         f"{len(dropped)} rows removed across {len({(d['facility'], d['medicine']) for d in dropped})} "
         f"facility x medicine pairs. The row kept is the one with more "
         f"fields filled; ties go to the later submission. "
         f"{len(conflicting)} of the dropped rows disagreed with the kept row "
         f"on at least one field - those are listed in full.")
    rule("Stock assessment", "S3", "Non-finite calculate fields treated as "
         "missing", "applied", None,
         "Kobo calculate fields emit the literal string 'NaN' when a "
         "denominator is zero or blank. float() accepts it, and one such "
         "value turns every downstream sum into NaN, so a non-finite number "
         "is read as missing.")
    rule("Stock assessment", "S4", "Level of care grouped", "applied",
         len(sub),
         "SGHP, 'Health Post (HP)' and a bare 'HP' are reported as one "
         "'Health post' tier; MHC joins 'Health centre'. The recorded value "
         "is kept in its own column - nothing is overwritten.")
    rule("Stock assessment", "S5", "Two servers merged", "applied",
         len(sub),
         "kf and EU exports are merged by fetch_kobo_export.py, with the EU "
         "_index offset so the two do not collide.")
    rule("Stock assessment", "S6",
         "EML applicability corrected against the official list", "applied",
         len(eml_fixed),
         f"Column I - 'Is <medicine> applicable to this facility level as per "
         f"EML?' - is answered from memory by the collector. Checked against "
         f"Final EML list.csv: {len(eml_fixed):,} rows said No where the list "
         f"holds that medicine at that level, and are corrected to Yes. ONE "
         f"DIRECTION ONLY: {len(eml_opposite):,} rows say Yes where the list "
         f"does not hold it at that level - reported, not changed. "
         f"{sum(eml_unmatched.values()):,} rows could not be checked because "
         f"the medicine name did not match the list by ingredient and "
         f"strength.")
    rule("Stock assessment", "S7",
         "Available = Yes with a quantity of zero corrected to No", "applied",
         len(zero_fixed),
         f"{len(zero_fixed):,} rows answered Yes to 'is it available on the "
         f"visit day?' and then recorded 0 units on the shelf. Nothing can be "
         f"dispensed from an empty shelf, so the quantity is taken as the "
         f"true answer and the flag is corrected to No. The opposite case - "
         f"No with a quantity above zero - does not occur, so the correction "
         f"is one-directional and unambiguous.")
    rule("Stock assessment", "S9",
         "Unavailable strength recoded to not managed where another "
         "strength was available", "applied", len(strength_recoded),
         f"A medicine held at two or more strengths, every strength answered "
         f"as managed, at least one on the shelf and another not: the "
         f"facility works with the strength it has, so the empty one is "
         f"recoded from managed = Yes / available = No to NOT MANAGED "
         f"(decision of 2026-10-15). Since 2026-11-03 every other answer "
         f"on a recoded row (availability through requisitions) is blanked: "
         f"a row that is not managed has no other responses considered. "
         f"{len(strength_recoded):,} rows at "
         f"{len({s(r['_p'][M.FAC]) for r in strength_recoded})} facilities. "
         f"Where the reason was blank it is set to 'Recoded S9: manages "
         f"<strength used>'. The two thiamine products (100mg 'THIAMINE (VIT "
         f"B1)' and 50mg 'Vitamin B1 (Thiamine)') are treated as one "
         f"medicine. The answers as recorded are kept in "
         f"managed_before_correction / available_before_correction.")
    _zr, _zc = list(M.ZONE_RECODED), list(getattr(C, "ZONE_RECODED", []))
    rule("Stock assessment", "S10",
         "Out-of-zone antimalarial, managed but not available, recoded to "
         "not managed", "applied", len(_zr),
         f"A first-line antimalarial (AL / ASPY / DHAP) answered managed = "
         f"Yes and available = No in a district whose MFT zone allocates "
         f"another one is recoded to NOT MANAGED: the empty shelf reflects "
         f"the zoning, not a stock-out (decision of 2026-11-02). "
         f"{len(_zr):,} rows at {len({s(r['_p'][M.FAC]) for r in _zr})} "
         f"facilities. Zoning in analytics/mft_zoning.py: official for 9 "
         f"districts, inferred from the data for the other 21 until the "
         f"official list is received. Where the reason was blank it is set "
         f"to 'Recoded S10: out of MFT zone (<zone>)'. Every other answer "
         f"on a recoded row is blanked (2026-11-03).")
    rule("CHW assessment", "S10",
         "Out-of-zone antimalarial, managed but not available, recoded to "
         "not managed", "applied", len(_zc),
         f"Same rule on the CHW form: {len(_zc):,} rows at "
         f"{len({(s(r['_p'][C.FAC]), s(r['_p'][C.CHW])) for r in _zc})} "
         f"CHWs. Other answers on the row are blanked (2026-11-03); the "
         f"general feedback question is kept.")
    _fp = list(getattr(M, "FP_RECODED", []))
    rule("Stock assessment", "S11",
         "FP products at faith-based facilities that do not provide FP "
         "recoded to not managed", "applied", len(_fp),
         f"Every FP product at {', '.join(sorted(M.FB_NO_FP))} - both "
         f"faith-based - is recoded to NOT MANAGED whatever was recorded, "
         f"and every other answer on the row is blanked (decision of "
         f"2026-11-05). {len(_fp):,} rows. The reason is set to 'Recoded "
         f"S11: faith-based facility, FP not provided'.")
    rule("Stock assessment", "S8",
         "Recording system name normalised", "applied", len(emr_changed),
         f"'Name of the electronical medical system used for recording' is "
         f"free text: {len({s(r[M.EMR]) for r in rows if s(r[M.EMR])})} "
         f"distinct values were typed for what is really "
         f"{len({r['_emr_clean'] for r in rows if r.get('_emr_clean')})} "
         f"systems - OpenMRS alone appears twenty-one ways. Each is mapped "
         f"to one spelling, matched on letters only so case, spacing and "
         f"hyphens collapse. {len(emr_numeric)} rows held a bare number "
         f"instead of a name and are emptied. A row naming several systems "
         f"is joined with ' + ' in a fixed order. The raw value is kept in "
         f"its own column and qualifiers go to a note.")

    wb = Workbook()
    wb.remove(wb.active)

    # --- submissions
    shdr = list(M.HDR_SUB) + ["level_group"]
    srows = [[r.get(h) for h in M.HDR_SUB] + [FL.group(r[M.LEVEL])]
             for r in sub]
    data_sheet(wb, "submissions", shdr, srows, {1: 20})

    # --- medicine rows, with the parent's identifiers carried on
    carry = [M.FAC, M.DIST, M.PROV, M.SECT, M.LEVEL, M.MGMT, M.ENUM]
    mhdr = (["facility", "district", "province", "sector", "level_recorded",
             "level_group", "management", "collector"]
            + list(M.HDR_MED) + ["molecule_key", "molecule_label",
                                 "applicable_before_correction",
                                 "eml_levels", "eml_match_type",
                                 "eml_products_matched",
                                 "available_before_correction",
                                 "managed_before_correction",
                                 "recording_system_clean",
                                 "recording_system_raw",
                                 "recording_system_note"])
    g = MOL.groups(rows)
    mrows = []
    for r in rows:
        k = MOL.key(r[M.MED])
        mrows.append([s(r["_p"][c]) for c in carry[:5]]
                     + [FL.group(r["_p"][M.LEVEL]), s(r["_p"][M.MGMT]),
                        s(r["_p"][M.ENUM])]
                     + [r.get(h) for h in M.HDR_MED]
                     + [k, MOL.label(g[k]), r.get("_applicable_was", ""),
                        r.get("_eml_levels", ""), r.get("_eml_match", ""),
                        r.get("_eml_products", ""),
                        r.get("_available_was", ""),
                        r.get("_managed_was", ""),
                        r.get("_emr_clean", ""), r.get("_emr_raw", ""),
                        r.get("_emr_note", "")])
    data_sheet(wb, "medicine_rows", mhdr, mrows,
               {1: 30, 2: 15, 3: 13, 4: 16, 5: 26, 6: 22, 7: 24, 8: 22})
    return (wb, dropped, conflicting, len(sub), len(rows),
            eml_fixed, eml_opposite, eml_unmatched, zero_fixed,
            emr_changed, emr_numeric, rows, strength_recoded)


def build_prescription():
    sub, pres, item = P.sub, P.pres, P.item
    moved = [r for r in sub if r.get("_facility_was")]
    reclass = [r for r in item
               if s(r.get(P.MED_EML)) in P.CLASS_OVERRIDES]

    rule("Prescription", "P1", "Facility reattributions", "applied",
         len(moved),
         f"{len(P.FACILITY_OVERRIDES)} rules, {len(moved)} submissions moved. "
         "Each is keyed on (facility, date) so a correction cannot sweep up a "
         "different visit that shares one of them. One rule keys on the VISIT "
         "date rather than the submission date, because the form was filled "
         "on paper and reached Kobo three days later.")
    rule("Prescription", "P2", "Drug class corrections", "applied",
         len(reclass),
         f"{len(reclass)} item rows moved from '15.6.2 Medicines for "
         f"diarrhoea' to '{P.RESPIRATORY}'. The misfiling is in the form's "
         "own choice list, so it is in every submission that recorded one of "
         "these eight respiratory medicines and will keep arriving until the "
         "XLSForm is corrected. Zinc sulphate stays under 15.6.2 - it belongs "
         "there.")
    rule("Prescription", "P3", "Level of care grouped", "applied", len(sub),
         "Same grouping as the stock assessment, so the two instruments'"
         " tiers line up.")
    rule("Prescription", "P4", "Two servers merged", "applied", len(sub),
         "Joined on (_submission__uuid, _index) - unlike the stock tool, no "
         "offset is applied, because the prescription export carries a uuid "
         "that is unique across both servers.")

    wb = Workbook()
    wb.remove(wb.active)
    shdr = list(P.hdr_sub) + ["level_group", "facility_before_correction"]
    srows = [[r.get(h) for h in P.hdr_sub]
             + [FL.group(r[P.FTYPE]), r.get("_facility_was", "")]
             for r in sub]
    data_sheet(wb, "submissions", shdr, srows, {1: 20})

    phdr = (["facility", "district", "province", "level_group"]
            + list(P.hdr_pres))
    prows = [[s(r["_sub"][P.FACIL]), s(r["_sub"][P.DISTR]),
              s(r["_sub"][P.PROV]), FL.group(r["_sub"][P.FTYPE])]
             + [r.get(h) for h in P.hdr_pres]
             for r in pres if r.get("_sub")]
    data_sheet(wb, "prescriptions", phdr, prows, {1: 30, 2: 15, 3: 13, 4: 22})

    ihdr = (["facility", "district", "province", "level_group"]
            + list(P.hdr_item)
            + ["drug_class_cleaned", "drug_class_corrected"])
    irows = [[s(r["_sub"][P.FACIL]), s(r["_sub"][P.DISTR]),
              s(r["_sub"][P.PROV]), FL.group(r["_sub"][P.FTYPE])]
             + [r.get(h) for h in P.hdr_item]
             + [r["_class"],
                "YES" if s(r.get(P.MED_EML)) in P.CLASS_OVERRIDES else ""]
             for r in item if r.get("_sub")]
    data_sheet(wb, "medicine_items", ihdr, irows,
               {1: 30, 2: 15, 3: 13, 4: 22})
    return wb, moved, reclass, len(sub), len(item)


def build_chw():
    sub, rows = C.SUB, C.MEDROWS
    rule("CHW assessment", "W1", "Level of care grouped", "applied", len(sub),
         "Same grouping as the other two instruments.")
    rule("CHW assessment", "W2", "Two servers merged", "applied", len(sub),
         "Same _index offset as the stock tool.")
    rule("CHW assessment", "W3", "'Not managed' left as recorded", "applied",
         None,
         "A CHW not managing a first-line therapy is MFT zoning - which "
         "therapy that zone uses - not a stock-out. Nothing is rewritten; "
         "the flag is kept so availability can be computed over managed rows "
         "only.")

    wb = Workbook()
    wb.remove(wb.active)
    shdr = list(C.HDR_SUB) + ["level_group"]
    srows = [[r.get(h) for h in C.HDR_SUB] + [FL.group(r[C.LEVEL])]
             for r in sub]
    data_sheet(wb, "submissions", shdr, srows, {1: 20})
    mhdr = (["facility", "district", "province", "level_group"]
            + list(C.HDR_MED))
    mrows = [[C.s(r["_p"][C.FAC]), C.s(r["_p"][C.DIST]),
              C.s(r["_p"][C.PROV]), FL.group(r["_p"][C.LEVEL])]
             + [r.get(h) for h in C.HDR_MED]
             for r in rows if r.get("_p")]
    data_sheet(wb, "medicine_rows", mhdr, mrows, {1: 30, 2: 15, 3: 13, 4: 22})
    return wb, len(sub), len(rows)


def build_log(stock, presc, chw):
    wb = Workbook()
    wb.remove(wb.active)

    # --- rules
    ws = wb.create_sheet("Rules")
    ws.sheet_view.showGridLines = False
    r = table(ws, 1, "Every cleaning rule applied, by tool",
              ["Tool", "ID", "Rule", "Type", "Rows affected", "Detail"],
              RULES, {1: 20, 2: 6, 3: 38, 4: 12, 5: 13, 6: 92})
    r = table(ws, r, "Analysis rules - NOT applied to the exported data",
              ["Tool", "ID", "Rule", "Type", "Rows affected", "Detail"],
              [["Stock assessment", "A1", "EML-applicable denominator",
                "analysis rule", None,
                "A medicine the facility is authorised to manage but does "
                "not manage counts as unavailable. A way of counting, not a "
                "correction - the exported data keeps the recorded "
                "management flag so either denominator can be built."],
               ["Stock assessment", "A2",
                "Method 3 - strengths combined per molecule",
                "analysis rule", None,
                "Two strengths of one molecule are interchangeable, so a "
                "facility holding either counts as served. The exported "
                "medicine_rows sheet carries molecule_key and molecule_label "
                "so this can be rebuilt, but the rows stay at item level."],
               ["Stock assessment", "A3",
                "Facilities with no digital system excluded from inventory "
                "accuracy", "analysis rule", None,
                "Accuracy compares a physical count with a system count; six "
                "facilities named no digital recording system on any row. "
                "Excluded from that indicator only, not from the data."]],
              {1: 20, 2: 6, 3: 38, 4: 12, 5: 13, 6: 92}, fill=INFO_FILL)

    # --- the actual changed rows
    dropped, conflicting = stock[1], stock[2]
    ws = wb.create_sheet("S2 duplicates removed")
    ws.sheet_view.showGridLines = False
    table(ws, 1, "Duplicate facility x medicine rows dropped",
          ["Facility", "Medicine", "Kept: submitted", "Kept: fields filled",
           "Dropped: submitted", "Dropped: fields filled",
           "Fields that disagreed", "Identical?"],
          [[d["facility"], d["medicine"], d["kept_sub"], d["kept_filled"],
            d["dropped_sub"], d["dropped_filled"], d["differs"],
            "yes" if d["identical"] else "NO - see fields"]
           for d in sorted(dropped, key=lambda x: (not x["identical"],
                                                   x["facility"]))],
          {1: 30, 2: 46, 3: 18, 4: 14, 5: 18, 6: 14, 7: 40, 8: 16})

    moved, reclass = presc[1], presc[2]
    ws = wb.create_sheet("P1 facility reattributions")
    ws.sheet_view.showGridLines = False
    table(ws, 1, "Submissions moved to a different facility",
          ["Recorded as", "Corrected to", "Level after", "District",
           "Visit date", "Submitted", "Collector"],
          sorted([[s(r.get("_facility_was")), s(r[P.FACIL]), s(r[P.FTYPE]),
                   s(r[P.DISTR]), s(r.get(P.VISIT_DATE))[:10],
                   s(r.get(P.SUB_TIME))[:16], s(r.get(P.COLLECTOR))]
                  for r in moved]),
          {1: 28, 2: 28, 3: 12, 4: 16, 5: 13, 6: 18, 7: 26})

    ws = wb.create_sheet("P2 drug class corrections")
    ws.sheet_view.showGridLines = False
    cnt = Counter(s(r.get(P.MED_EML)) for r in reclass)
    table(ws, 1, "Respiratory medicines moved out of the diarrhoea class",
          ["Medicine code as recorded", "Readable name", "Items", "From", "To"],
          [[k, next((s(r.get("Medicine (name)")) for r in reclass
                     if s(r.get(P.MED_EML)) == k and s(r.get("Medicine (name)"))),
                    ""), n, "15.6.2 Medicines for diarrhoea", P.RESPIRATORY]
           for k, n in cnt.most_common()],
          {1: 58, 2: 52, 3: 8, 4: 32, 5: 46})

    # --- S6: the EML applicability corrections
    fixed, opposite, unmatched = stock[5], stock[6], stock[7]
    ws = wb.create_sheet("S6 EML applicability")
    ws.sheet_view.showGridLines = False
    cnt = Counter((s(r[M.MED]), s(r["_p"][M.LEVEL]), r.get("_eml_match", ""),
                   r.get("_eml_levels", "")) for r in fixed)
    r0 = table(ws, 1,
               "Corrected No -> Yes: the EML list holds this medicine at this "
               "level",
               ["Medicine (as recorded)", "Facility level", "Rows corrected",
                "Match quality", "Levels the EML lists it at",
                "EML product matched"],
               [[m, l, n, how, lv,
                 next((x.get("_eml_products", "") for x in fixed
                       if s(x[M.MED]) == m), "")]
                for (m, l, how, lv), n in cnt.most_common()],
               {1: 44, 2: 30, 3: 12, 4: 20, 5: 34, 6: 60})
    r0 = table(ws, r0,
               "NOT changed: the form says Yes, the EML list does not hold it "
               "at this level",
               ["Medicine (as recorded)", "Facility level", "Rows",
                "Levels the EML lists it at"],
               [[m, l, n, lv] for (m, l, lv), n in
                Counter((s(r[M.MED]), s(r["_p"][M.LEVEL]),
                         r.get("_eml_levels", "")) for r in opposite
                        ).most_common()],
               {1: 44, 2: 30, 3: 10, 4: 34}, fill=WARN_FILL)
    table(ws, r0,
          "Could not be checked: the medicine name did not match the EML list",
          ["Medicine (as recorded)", "Rows", "Why"],
          [[m, n, "fixed-dose combination, commodity or brand name - not "
            "mapped by eye"] for m, n in unmatched.most_common()],
          {1: 58, 2: 10, 3: 62}, fill=INFO_FILL)

    # --- S7: availability corrected where the shelf was empty
    zero = stock[8]
    ws = wb.create_sheet("S7 available but zero qty")
    ws.sheet_view.showGridLines = False
    r0 = table(ws, 1,
               "Answered 'available = Yes' with 0 units on the shelf - "
               "corrected to No",
               ["Facility", "District", "Level of care", "Medicine",
                "Programme", "Quantity on shelf", "Quantity in system",
                "Stock-out days Jun-Aug", "Was", "Now"],
               sorted([[s(r["_p"][M.FAC]), s(r["_p"][M.DIST]),
                        FL.group(r["_p"][M.LEVEL]), s(r[M.MED]),
                        s(r[M.PROGRAM]), M.num(r[M.QTY_PHYS]),
                        M.num(r[M.QTY_SYS]),
                        sum(M.num(r[M.SO_MONTH[m]]) or 0
                            for m in ("June", "July", "August")),
                        r.get("_available_was", ""), "No"]
                       for r in zero], key=lambda x: (x[0], x[3])),
               {1: 30, 2: 15, 3: 24, 4: 46, 5: 20, 6: 16, 7: 16, 8: 20,
                9: 8, 10: 8})
    table(ws, r0, "Most affected medicines",
          ["Medicine", "Rows corrected"],
          Counter(s(r[M.MED]) for r in zero).most_common(),
          {1: 56, 2: 14}, fill=INFO_FILL)

    # --- S11: FP at faith-based facilities
    ws = wb.create_sheet("S11 FP faith-based")
    ws.sheet_view.showGridLines = False
    table(ws, 1,
          "FP products at faith-based facilities that do not provide FP - "
          "recoded to NOT MANAGED, other answers blanked",
          ["Facility", "District", "Management", "Product", "Managed was",
           "Available was", "Now"],
          sorted([[s(r["_p"][M.FAC]), s(r["_p"][M.DIST]),
                   s(r["_p"].get(M.MGMT)), s(r[M.MED]),
                   r.get("_managed_was", ""), r.get("_available_was", ""),
                   "Not managed"]
                  for r in getattr(M, "FP_RECODED", [])]),
          {1: 16, 2: 13, 3: 26, 4: 60, 5: 12, 6: 12, 7: 13})

    # --- S10: out-of-zone antimalarial, managed but unavailable
    import mft_zoning as _ZN
    ws = wb.create_sheet("S10 zone recoded")
    ws.sheet_view.showGridLines = False
    table(ws, 1,
          "First-line antimalarial managed but not available, outside the "
          "district's MFT zone - recoded to NOT MANAGED",
          ["Form", "District", "District zone", "Zone basis", "Facility",
           "CHW", "Product", "Managed was", "Available was", "Now"],
          sorted([["Stock", s(r["_p"][M.DIST]),
                   _ZN.ZONE.get(s(r["_p"][M.DIST]), ""),
                   "official" if s(r["_p"][M.DIST]) in _ZN.OFFICIAL
                   else "inferred", s(r["_p"][M.FAC]), "", s(r[M.MED]),
                   r.get("_managed_was", ""), "No", "Not managed"]
                  for r in M.ZONE_RECODED]
                 + [["CHW", s(r["_p"][C.DIST]),
                     _ZN.ZONE.get(s(r["_p"][C.DIST]), ""),
                     "official" if s(r["_p"][C.DIST]) in _ZN.OFFICIAL
                     else "inferred", s(r["_p"][C.FAC]), s(r["_p"][C.CHW]),
                     r["_short"], r.get("_managed_was", ""), "No",
                     "Not managed"]
                    for r in getattr(C, "ZONE_RECODED", [])],
                 key=lambda x: (x[0], x[1], x[4], x[5], x[6])),
          {1: 8, 2: 13, 3: 11, 4: 10, 5: 26, 6: 24, 7: 50, 8: 12, 9: 12,
           10: 13})

    # --- S9: unavailable strength recoded to not managed
    s9 = stock[12]
    ws = wb.create_sheet("S9 strength recoded")
    ws.sheet_view.showGridLines = False
    r0 = table(ws, 1,
               "Managed but unavailable, while another strength of the same "
               "medicine was available - recoded to NOT MANAGED",
               ["Facility", "District", "Level of care", "Medicine recoded",
                "Strength(s) available instead", "Stock-out days Jun-Aug",
                "Managed was", "Available was", "Now"],
               sorted([[s(r["_p"][M.FAC]), s(r["_p"][M.DIST]),
                        FL.group(r["_p"][M.LEVEL]), s(r[M.MED]),
                        s(r[M.WHY_NOT]).replace("Recoded S9: manages ", ""),
                        sum(M.num(r["_blanked"].get(M.SO_MONTH[m])) or 0
                            for m in ("June", "July", "August")),
                        r.get("_managed_was", ""), "No", "Not managed"]
                       for r in s9], key=lambda x: (x[0], x[3])),
               {1: 30, 2: 15, 3: 18, 4: 46, 5: 46, 6: 18, 7: 12, 8: 12,
                9: 14})
    table(ws, r0, "Rows recoded by medicine",
          ["Medicine", "Rows recoded"],
          Counter(s(r[M.MED]) for r in s9).most_common(),
          {1: 56, 2: 14}, fill=INFO_FILL)

    # --- S8: recording system names
    emr_changed, emr_numeric = stock[9], stock[10]
    ws = wb.create_sheet("S8 recording system")
    ws.sheet_view.showGridLines = False
    pairs = Counter((r.get("_emr_raw", ""), r.get("_emr_clean", ""),
                     r.get("_emr_note", "")) for r in emr_changed)
    r0 = table(ws, 1, "Every raw spelling and what it was mapped to",
               ["As recorded", "Normalised to", "Rows", "Note kept"],
               [[raw, clean or "(no system)", n, note]
                for (raw, clean, note), n in
                sorted(pairs.items(), key=lambda kv: (kv[0][1], -kv[1]))],
               {1: 56, 2: 26, 3: 9, 4: 46})
    tot = Counter()
    for r in stock[11]:
        tot[r.get("_emr_clean", "") or "(no system recorded)"] += 1
    r0 = table(ws, r0, "Distinct systems after normalising",
               ["System", "Rows"], tot.most_common(),
               {1: 30, 2: 10}, fill=INFO_FILL)
    table(ws, r0, "Numbers removed - a quantity typed into the name box",
          ["Facility", "Medicine", "Value removed"],
          sorted([[s(r["_p"][M.FAC]), s(r[M.MED]), r.get("_emr_raw", "")]
                  for r in emr_numeric]),
          {1: 30, 2: 46, 3: 16}, fill=WARN_FILL)

    # --- counts before and after
    ws = wb.create_sheet("Row counts")
    ws.sheet_view.showGridLines = False
    table(ws, 1, "What went in and what came out",
          ["Tool", "Sheet", "Rows exported", "Note"],
          [["Stock assessment", "submissions", stock[3], ""],
           ["Stock assessment", "medicine_rows", stock[4],
            f"{len(dropped)} duplicate rows removed; "
            f"{len(fixed)} EML applicability flags corrected; "
            f"{len(zero)} availability flags corrected; "
            f"{len(stock[12])} strengths recoded to not managed (S9); "
            f"{len(emr_changed)} recording system names normalised"],
           ["Prescription", "submissions", presc[3],
            f"{len(moved)} moved to a different facility"],
           ["Prescription", "medicine_items", presc[4],
            f"{len(reclass)} drug classes corrected"],
           ["CHW assessment", "submissions", chw[1], ""],
           ["CHW assessment", "medicine_rows", chw[2], ""]],
          {1: 20, 2: 20, 3: 15, 4: 44})
    return wb


def main():
    ensure_dir()
    stock = build_stock()
    presc = build_prescription()
    chw = build_chw()
    log = build_log(stock, presc, chw)
    names = [write(log, "00_Cleaning_log.xlsx"),
             write(stock[0], "01_Stock_assessment_clean.xlsx"),
             write(presc[0], "02_Prescription_clean.xlsx"),
             write(chw[0], "03_CHW_assessment_clean.xlsx")]
    print(f"wrote to {OUTDIR}")
    for n in names:
        print(f"   {n}")
    print(f"\n{len(RULES)} cleaning rules applied:")
    for tool, rid, name, kind, n, _ in RULES:
        print(f"   {rid}  {tool:<18}{name[:46]:<48}"
              f"{'' if n is None else format(n, ',')}")


if __name__ == "__main__":
    main()
