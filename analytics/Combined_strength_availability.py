"""Availability with multi-strength medicines combined into one molecule.

    python3 Combined_strength_availability.py
      -> Combined_strength_availability.xlsx

A patient who needs metformin can be treated with 500mg or 850mg, so counting
the two strengths as separate tracer items understates what the facility can
actually do. This combines them.

THREE WAYS TO COMBINE, because they answer different questions and give
different numbers:

  Method 1 - MEAN OF STRENGTHS (the method requested)
    availability of each strength, then the unweighted average of those.
    Treats every strength as equally important regardless of how widely it
    is stocked.

  Method 2 - POOLED
    all observations of every strength in one numerator and denominator.
    Weights each strength by how often it was assessed, so the commonly
    stocked strength dominates.

  Method 3 - ANY STRENGTH AT THE FACILITY
    for each facility, was AT LEAST ONE strength of the molecule available?
    This is the clinical question - it is the only one of the three that
    answers "could this facility treat the patient".

Method 1 is reported first because it is what was asked for. Method 3 is the
one to quote clinically, and the gap between them is large for exactly the
molecules where one strength is rare.

openpyxl + the standard library, like the rest of this project.
"""
import _bootstrap as B

import os
import re
import statistics
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_medicine_assessment as M
import facility_levels as FL

OUT = B.out("Combined_strength_availability.xlsx")

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HF = PatternFill("solid", fgColor="1F3864")
HFONT = Font(bold=True, size=9, color="FFFFFF")
TOTF = PatternFill("solid", fgColor="F2F2F2")
BAD = PatternFill("solid", fgColor="FCE4E4")
MIDF = PatternFill("solid", fgColor="FFF4D6")
OKF = PatternFill("solid", fgColor="E7F3E7")
TITLE = Font(bold=True, size=12, color="1F3864")

# words that describe a strength, pack or form rather than the molecule
FORMS = (r"tab|tabs|tablet|tablets|caps|capsule|capsules|inj|injection|"
         r"injectable|powder|for|disp|dispersible|susp|suspension|syrup|"
         r"solution|sol|oint|ointment|cream|supp|suppository|drops|film|"
         r"coated|granules|iv|im|sc|po|bottle|bottleof|vial|amp|ampoule|"
         r"pr|dec|of|with|solvent")


def molecule(name):
    """Medicine name -> molecule key, with strength, pack size and dosage
    form removed. Stray brackets are stripped BEFORE the parenthetical rule,
    because one DHAP entry is written 'DHAP)' with no opening bracket and
    would otherwise key differently from its own siblings."""
    t = str(name).lower().replace("(", " ").replace(")", " ")
    t = re.sub(r"\b\d+[\d.,/]*\s*(mg|g|gr|mcg|µg|ui|iu|ml|l|%|mgs)\b", " ", t)
    t = re.sub(r"\b\d+\s*x\s*\d+\b", " ", t)
    t = re.sub(r"\bb\s*/?\s*\d+\b", " ", t)
    t = re.sub(r"\d+[\d.,/]*", " ", t)
    t = re.sub(r"\b(" + FORMS + r")\b", " ", t)
    t = re.sub(r"[^a-z ]", " ", t)
    return " ".join(t.split())


def s(v):
    return "" if v is None else str(v).strip()


def pct(a, b):
    return a / b if b else None


ROWS = M.MEDROWS
for r in ROWS:
    r["_mol"] = molecule(r[M.MED])
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])

GROUPS = defaultdict(set)
for r in ROWS:
    GROUPS[r["_mol"]].add(s(r[M.MED]))
MULTI = {k: sorted(v) for k, v in GROUPS.items() if len(v) > 1}
SINGLE = {k: sorted(v)[0] for k, v in GROUPS.items() if len(v) == 1}


def elig(rows):
    return [r for r in rows if M.yes(r[M.MANAGED]) and not M.blank(r[M.AVAIL])]


def rate(rows):
    e = elig(rows)
    return pct(sum(1 for r in e if M.yes(r[M.AVAIL])), len(e)), len(e)


def molecule_label(mol):
    names = sorted(GROUPS[mol])
    if len(names) == 1:
        return names[0]
    stem = names[0].split()[0].title()
    return f"{stem} (all {len(names)} strengths)"


def three_methods(mol, rows=None):
    """-> mean-of-strengths, pooled, any-strength-at-facility, plus counts"""
    src = [r for r in (rows if rows is not None else ROWS) if r["_mol"] == mol]
    names = sorted({s(r[M.MED]) for r in src})
    per = []
    for n in names:
        rt, nn = rate([r for r in src if s(r[M.MED]) == n])
        if rt is not None:
            per.append((n, rt, nn))
    m1 = statistics.mean(x[1] for x in per) if per else None
    m2, n2 = rate(src)
    # method 3: by facility, any strength available
    byfac = defaultdict(list)
    for r in elig(src):
        byfac[s(r["_p"][M.FAC])].append(r)
    any_ok = sum(1 for f, rs in byfac.items()
                 if any(M.yes(x[M.AVAIL]) for x in rs))
    m3 = pct(any_ok, len(byfac))
    return m1, m2, m3, per, n2, len(byfac), any_ok


def xsheet(wb, title, blocks, notes=None, widths=None):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    for n in (notes or []):
        c = ws.cell(row=r, column=1, value=n)
        c.font = Font(size=9, italic=True, color="555555")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        ws.row_dimensions[r].height = 13 * (1 + len(n) // 135)
        r += 1
    if notes:
        r += 1
    for btitle, headers, rows, total, pctcols in blocks:
        if btitle:
            ws.cell(row=r, column=1, value=btitle).font = TITLE
            r += 1
        for ci, h in enumerate(headers, start=1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font, c.fill, c.border = HFONT, HF, BORDER
            c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 44
        hdr = r
        r += 1
        for row in list(rows) + ([total] if total else []):
            istot = total is not None and row is total
            for ci, v in enumerate(row, start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.border = BORDER
                c.alignment = Alignment(vertical="top",
                                        wrap_text=isinstance(v, str)
                                        and len(str(v)) > 34)
                if (ci - 1) in pctcols and isinstance(v, float):
                    c.number_format = "0.0%"
                    c.fill = BAD if v < .85 else (MIDF if v < .95 else OKF)
                elif isinstance(v, (int, float)):
                    c.number_format = "#,##0.##"
                if istot:
                    c.font = Font(bold=True)
                    if not ((ci - 1) in pctcols and isinstance(v, float)):
                        c.fill = TOTF
            r += 1
        ws.freeze_panes = f"B{hdr + 1}"
        ws.auto_filter.ref = f"A{hdr}:{get_column_letter(len(headers))}{r - 1}"
        r += 2
    for ci, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(ci)].width = w


def build():
    wb = Workbook()
    wb.remove(wb.active)

    # ---- 1. the multi-strength molecules, three methods
    rows, detail = [], []
    for mol in sorted(MULTI):
        m1, m2, m3, per, n2, nfac, nok = three_methods(mol)
        rows.append([molecule_label(mol), len(per), n2, nfac, m1, m2, m3,
                     (m3 - m1) if (m1 is not None and m3 is not None) else None,
                     "; ".join(f"{n.split()[0][:14]}… {r:.0%}" for n, r, _ in per)])
        for n, rt, nn in sorted(per, key=lambda x: x[1]):
            detail.append([molecule_label(mol), n, nn, rt,
                           pct(nn, n2), m1, m2, m3])
    rows.sort(key=lambda x: x[4] if x[4] is not None else 2)
    xsheet(wb, "Combined molecules",
           [(None, ["Molecule", "Strengths", "Eligible observations",
                    "Facilities", "METHOD 1 mean of strengths",
                    "METHOD 2 pooled", "METHOD 3 any strength at the facility",
                    "Method 3 minus Method 1", "Per-strength availability"],
             rows, None, (4, 5, 6, 7))],
           notes=[f"The {len(MULTI)} molecules in the basket that are "
                  f"assessed at more than one strength or presentation.",
                  "Method 1 is the requested calculation: average the "
                  "strengths. Method 2 pools all observations. Method 3 asks "
                  "whether the facility had ANY strength - the clinical "
                  "question, and the one that should be quoted when the "
                  "strengths are interchangeable.",
                  "Where Method 3 is far above Method 1, the molecule is "
                  "available but one strength is rare - that is a formulary "
                  "choice, not a stock-out."],
           widths={1: 34, 2: 10, 3: 13, 4: 11, 5: 14, 6: 12, 7: 15, 8: 13,
                   9: 54})

    xsheet(wb, "Strength detail",
           [(None, ["Molecule", "Strength as recorded",
                    "Eligible observations", "Availability of this strength",
                    "Share of the molecule's observations",
                    "Method 1 molecule", "Method 2 molecule",
                    "Method 3 molecule"], detail, None, (3, 4, 5, 6, 7))],
           notes=["Every strength behind the combined figures, weakest first "
                  "within each molecule.",
                  "The 'share of observations' column shows why Methods 1 and "
                  "2 diverge: a strength assessed at 25 facilities counts as "
                  "much as one assessed at 234 under Method 1, and far less "
                  "under Method 2."],
           widths={1: 32, 2: 46, 3: 13, 4: 14, 5: 15, 6: 13, 7: 13, 8: 13})

    # ---- 2. full basket, molecules instead of items
    full = []
    for mol in sorted(GROUPS):
        m1, m2, m3, per, n2, nfac, nok = three_methods(mol)
        if n2 == 0:
            continue
        ex = [r for r in ROWS if r["_mol"] == mol][0]
        full.append([molecule_label(mol), s(ex[M.PROGRAM]), len(per), n2,
                     m1, m2, m3])
    full.sort(key=lambda x: x[4] if x[4] is not None else 2)
    xsheet(wb, "Full basket by molecule",
           [("All molecules, weakest first on Method 1", HDR_FULL := [
               "Molecule", "Programme", "Strengths", "Eligible observations",
               "METHOD 1 mean of strengths", "METHOD 2 pooled",
               "METHOD 3 any strength"], full, None, (4, 5, 6))],
           notes=[f"The basket collapses from "
                  f"{len({s(r[M.MED]) for r in ROWS})} items to "
                  f"{len(full)} molecules. Single-strength medicines are "
                  f"unchanged - all three methods give the same number for "
                  f"them, which is why most rows agree.",
                  "This is the sheet to use for a 'ten least available "
                  "medicines' list once strengths are combined."],
           widths={1: 46, 2: 20, 3: 10, 4: 13, 5: 14, 6: 12, 7: 14})

    # ---- 3. how the ten-least-available list changes
    item_rates = []
    for n in sorted({s(r[M.MED]) for r in ROWS}):
        rt, nn = rate([r for r in ROWS if s(r[M.MED]) == n])
        if rt is not None and nn >= 30:
            item_rates.append([n, nn, rt])
    item_rates.sort(key=lambda x: x[2])
    mol_rates = sorted([f for f in full if f[3] >= 30],
                       key=lambda x: x[4] if x[4] is not None else 2)
    comp = []
    for i in range(10):
        a = item_rates[i] if i < len(item_rates) else ["", "", None]
        b = mol_rates[i] if i < len(mol_rates) else ["", "", "", "", None]
        comp.append([i + 1, a[0][:44], a[2], b[0][:44], b[4]])
    xsheet(wb, "Ten least available",
           [(None, ["Rank", "BEFORE - by item as recorded", "Availability",
                    "AFTER - by molecule, Method 1", "Availability"],
             comp, None, (2, 4))],
           notes=["Combining strengths changes which medicines appear on the "
                  "priority list, because a molecule whose rare strength was "
                  "dragging it down now sits beside its common strength.",
                  "Ranked among medicines with at least 30 eligible "
                  "observations."],
           widths={1: 6, 2: 46, 3: 12, 4: 46, 5: 12})

    # ---- 4. national effect
    e = elig(ROWS)
    nat_item = pct(sum(1 for r in e if M.yes(r[M.AVAIL])), len(e))
    m1s = [f[4] for f in full if f[4] is not None]
    m3s = [f[6] for f in full if f[6] is not None]
    item_means = [x[2] for x in item_rates]
    head = [
        ["Pooled across every observation, items as recorded",
         nat_item, len(e),
         "The current national figure. Unchanged by this grouping - pooling "
         "does not care how items are labelled."],
        ["Mean across items as recorded",
         statistics.mean(item_means) if item_means else None, len(item_rates),
         "Unweighted mean of item-level rates, before combining."],
        ["Mean across molecules, Method 1",
         statistics.mean(m1s) if m1s else None, len(m1s),
         "Each molecule counts once, its strengths averaged first. THE "
         "REQUESTED CALCULATION."],
        ["Mean across molecules, Method 3",
         statistics.mean(m3s) if m3s else None, len(m3s),
         "Each molecule counts once, scored on whether the facility had any "
         "strength. The clinically meaningful version."],
    ]
    xsheet(wb, "National effect",
           [(None, ["Basis", "Availability", "Items or molecules",
                    "What it means"], head, None, (1,))],
           notes=["Combining strengths does not move the pooled national "
                  "figure - that is a property of pooling. What it changes is "
                  "every per-medicine figure, the ranking of the weakest "
                  "medicines, and any mean taken across medicines.",
                  "If a single national number is quoted, say which of these "
                  "four it is."],
           widths={1: 50, 2: 13, 3: 16, 4: 76})

    # ---- 5. method note
    note = [
        ["What was combined",
         f"{len(MULTI)} molecules assessed at more than one strength or "
         f"presentation, covering "
         f"{sum(len(v) for v in MULTI.values())} of the "
         f"{len({s(r[M.MED]) for r in ROWS})} basket items."],
        ["How molecules were identified",
         "Automatically: strength, pack size and dosage form stripped from "
         "the medicine name, the remainder used as the key. Every group was "
         "checked by eye and is listed on the first sheet - no molecule is "
         "grouped silently."],
        ["One grouping bug found and fixed",
         "'Dihydroartemisinin/Piperaquine (DHAP) 20/160 mg' is written with "
         "a matched bracket while its two siblings are written 'DHAP)' with "
         "no opening bracket. The first version of the grouping rule put "
         "20/160 in a group of its own; brackets are now stripped before the "
         "parenthetical rule so all three DHAP strengths group together."],
        ["Which method to use",
         "Method 1 as requested for a medicine-list report; Method 3 when the "
         "question is whether a patient could be treated, because the "
         "strengths are therapeutically interchangeable. Method 2 if the "
         "figure must be consistent with the pooled national rate."],
        ["What this does NOT change",
         "The pooled national availability figure, and any facility-level "
         "basket score computed from observations rather than molecules. "
         "Those are unaffected by relabelling items."],
        ["Interaction with the proposed denominator change",
         "If 'applicable but not stocked' is later counted as unavailable, "
         "combine strengths FIRST: a facility stocking Metformin 500mg would "
         "otherwise be recorded as missing Metformin 850mg, which the "
         "strength-combining rule is precisely designed to prevent."],
    ]
    xsheet(wb, "Method", [(None, ["Item", "Detail"], note, None, ())],
           widths={1: 44, 2: 104})

    wb.active = 0
    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)
    wb.save(OUT)
    print(f"wrote {os.path.basename(OUT)}")
    print(f"  {len({s(r[M.MED]) for r in ROWS})} items -> {len(full)} "
          f"molecules | {len(MULTI)} multi-strength groups")
    print(f"  national, pooled (unchanged)        {nat_item:.1%}")
    print(f"  mean across items as recorded       "
          f"{statistics.mean(item_means):.1%}")
    print(f"  mean across molecules, Method 1     {statistics.mean(m1s):.1%}")
    print(f"  mean across molecules, Method 3     {statistics.mean(m3s):.1%}")
    print("  the nine combined molecules:")
    for row in rows:
        print(f"    {row[0][:34]:36} M1 {row[4]:6.1%}  M2 {row[5]:6.1%}  "
              f"M3 {row[6]:6.1%}")


if __name__ == "__main__":
    build()
