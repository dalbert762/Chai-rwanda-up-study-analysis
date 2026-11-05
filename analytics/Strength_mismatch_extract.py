"""Medicines held at more than one strength where one strength was on the
shelf and another was not.

    python3 Strength_mismatch_extract.py
      -> Strength_mismatch_extract.xlsx

A CASE is one facility x medicine where the facility MANAGES two or more
strengths of the same molecule, answered availability for them, and at
least one strength was available on the visit day while at least one other
was not. Under Method 3 (strengths combined) such a facility counts as
having the medicine; this lists them so the strength-level gap is visible.

Rows are analyze_medicine_assessment's joined rows - the same cleaned base
as the dashboard. Strengths are grouped with molecules.key(), plus one
merge it does not make: "THIAMINE 100MG (VIT B1)TABLET" and "Vitamin B1
tablet(Thiamine) 50mg" are the same molecule written two ways, so they are
grouped together here.

openpyxl + the standard library, like the rest of this project.
"""

import _bootstrap as B   # path setup - must come first
import os
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_medicine_assessment as M
import facility_levels as FL
import molecules as MOL
import molkey

OUT = B.out("Strength_mismatch_extract.xlsx")

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="DDE5F0")
OK_FILL = PatternFill("solid", fgColor="E7F3E7")
BAD_FILL = PatternFill("solid", fgColor="FCE4E4")


def s(v):
    return "" if v is None else str(v).strip()


def num(v):
    return M.num(v)


def mol_key(name):
    return molkey.strength_key(name)


# Cleaning rule S9 recodes the unavailable strength of exactly these cases
# to NOT MANAGED, so the cases are found on the answers AS RECORDED: a row
# S9 changed was managed = Yes, available = No.
# S10 (MFT zoning) also keeps _managed_was; those rows were recorded
# managed = Yes / available = No too, so they read back the same way.
S9_IDS = {id(r) for r in M.STRENGTH_RECODED}
# S9 blanks a recoded row's other answers (2026-11-03); read them back from
# r["_blanked"] so this extract shows the row as recorded.
ROWS = []
for _r in M.MEDROWS:
    if "_blanked" in _r:
        _c = dict(_r, **_r["_blanked"])
        S9_IDS.add(id(_c))
        ROWS.append(_c)
    else:
        ROWS.append(_r)


def recoded(r):
    return bool(r.get("_managed_was"))


def was_managed(r):
    return recoded(r) or M.yes(r[M.MANAGED])


def was_avail(r):
    return "No" if recoded(r) else s(r[M.AVAIL])


def mol_label(names):
    return MOL.label(sorted(names)) if len(names) > 1 else sorted(names)[0]


# A flag is LONG when its unavailable strength was out for more than LONG
# days over June + July + August (the monthly fields summed).
LONG = 60


def so_window(r):
    return sum(num(r[M.SO_MONTH[m]]) or 0 for m in ("June", "July", "August"))


# ---------------------------------------------------------------
# find the cases
# ---------------------------------------------------------------
groups = defaultdict(list)
for r in ROWS:
    groups[(s(r["_p"][M.FAC]), mol_key(r[M.MED]))].append(r)

# Left out of the list at the user's request (2026-11-03): DHAP and AL
# (ARTE + LUME 1x6 / 4x6) strength gaps follow MFT zoning and the
# weight-band packs, not a strength stock-out.
EXCLUDE = {"dihydroartemisinin piperaquine dhap", "arte lume"}

CASES, SINGLE_OK = [], Counter()       # SINGLE_OK: facilities managing 2+
for (fac, mol), rs in groups.items():
    if mol in EXCLUDE:
        continue
    managed = [r for r in rs if was_managed(r) and was_avail(r)]
    if len({s(r[M.MED]) for r in managed}) < 2:
        continue
    SINGLE_OK[mol] += 1
    av = [r for r in managed if was_avail(r) == "Yes"]
    out = [r for r in managed if was_avail(r) == "No"]
    if av and out:
        CASES.append((fac, mol, managed, av, out))

CASES.sort(key=lambda c: (s(c[2][0]["_p"][M.PROV]), s(c[2][0]["_p"][M.DIST]),
                          c[0], c[1]))
NAMES = defaultdict(set)
for (fac, mol), rs in groups.items():
    NAMES[mol] |= {s(r[M.MED]) for r in rs}


# ---------------------------------------------------------------
# workbook
# ---------------------------------------------------------------
wb = Workbook()
wb.remove(wb.active)


def sheet(name, header, rows, widths=None, notes=(), status_col=None):
    ws = wb.create_sheet(name)
    row = 1
    for n in notes:
        c = ws.cell(row, 1, n)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1, end_row=row,
                       end_column=min(len(header), 10))
        ws.row_dimensions[row].height = 15 * max(1, len(n) // 150 + 1)
        row += 1
    if notes:
        row += 1
    for j, h in enumerate(header, 1):
        c = ws.cell(row, j, h)
        c.font = Font(bold=True)
        c.fill = HDR_FILL
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[row].height = 32
    ws.freeze_panes = ws.cell(row + 1, 1)
    first = row + 1
    for vals in rows:
        row += 1
        for j, v in enumerate(vals, 1):
            c = ws.cell(row, j, v)
            c.border = BORDER
            if isinstance(v, float) and 0 <= v <= 1 and "%" in header[j - 1]:
                c.number_format = "0.0%"
            elif isinstance(v, float):
                c.number_format = "#,##0.##"
            elif isinstance(v, int):
                c.number_format = "#,##0"
        if status_col:
            c = ws.cell(row, status_col)
            c.fill = OK_FILL if c.value == "Yes" else BAD_FILL
    ws.auto_filter.ref = (f"A{first - 1}:{get_column_letter(len(header))}"
                          f"{max(row, first)}")
    for j in range(1, len(header) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 14
    for j, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(j)].width = w
    return ws


# 1. one row per case
case_rows = []
for fac, mol, managed, av, out in CASES:
    p = managed[0]["_p"]
    case_rows.append([
        s(p[M.PROV]), s(p[M.DIST]), s(p[M.SECT]), fac,
        FL.group(p[M.LEVEL]), s(p.get(M.MGMT)), s(p.get(M.ENUM)),
        s(p.get("today"))[:10], mol_label(NAMES[mol]), s(managed[0][M.PROGRAM]),
        len(managed),
        "; ".join(f"{s(r[M.MED])} ({num(r[M.QTY_PHYS]) or 0:,.0f} on shelf)"
                  for r in av),
        "; ".join(f"{s(r[M.MED])} (out {num(r[M.SO_NOW]) or 0:,.0f} days now, "
                  f"{so_window(r):,.0f} days Jun-Aug)" for r in out),
        max(so_window(r) for r in out),
        f'=IF(N{len(case_rows) + 4}>{LONG},"FLAG {LONG}+","")'])

# 2. one row per strength in those cases
item_rows = []
for fac, mol, managed, av, out in CASES:
    for r in sorted(managed, key=lambda r: s(r[M.MED])):
        p = r["_p"]
        item_rows.append([
            s(p[M.PROV]), s(p[M.DIST]), fac, FL.group(p[M.LEVEL]),
            mol_label(NAMES[mol]), s(r[M.MED]), s(r[M.PROGRAM]),
            "Yes", was_avail(r), "Not managed (S9)" if id(r) in S9_IDS
            else "Not managed (S10)" if recoded(r) else "unchanged", num(r[M.QTY_PHYS]),
            num(r[M.QTY_SYS]), num(r[M.AMC]), num(r[M.MOS]),
            s(r[M.STOCKCLASS]), num(r[M.SO_NOW]),
            num(r[M.SO_MONTH["June"]]), num(r[M.SO_MONTH["July"]]),
            num(r[M.SO_MONTH["August"]]), num(r[M.REQ_TOT]),
            num(r[M.DEL_TOT]), s(r[M.RMS_REG]), s(r[M.RMS_CEN])])

# 3. by medicine
by_mol = Counter(c[1] for c in CASES)
combo = defaultdict(Counter)
for fac, mol, managed, av, out in CASES:
    combo[mol][(" + ".join(sorted(s(r[M.MED]) for r in av)),
                " + ".join(sorted(s(r[M.MED]) for r in out)))] += 1
sum_rows = []
for mol, n in by_mol.most_common():
    (a, o), k = combo[mol].most_common(1)[0]
    sum_rows.append([mol_label(NAMES[mol]), SINGLE_OK[mol], n,
                     n / SINGLE_OK[mol], a, o, k])
sum_rows.append(["ALL MEDICINES", sum(SINGLE_OK.values()), len(CASES),
                 len(CASES) / sum(SINGLE_OK.values()), "", "", None])

readme = [
    ["What a case is",
     "One facility x medicine where the facility MANAGES two or more "
     "strengths of the same medicine and answered availability for them, "
     "and at least one strength was available on the visit day while at "
     "least one other was not."],
    ["Why it matters",
     "The dashboard combines strengths (Method 3): a facility holding any "
     "strength counts as having the medicine. These are the cases where "
     "that hides a strength-level stock-out."],
    ["Source", "analyze_medicine_assessment.py joined rows - the cleaned "
     "base the dashboard uses (duplicates removed, available-with-zero-"
     "quantity corrected to No, facilities reattributed)."],
    ["Cleaning rule S9", "In the cleaned data the unavailable strength of "
     "every case (all durations, DHAP and AL included - 199 rows) is "
     "recoded to NOT MANAGED, and since 2026-11-03 its other answers are "
     "blanked. This extract shows "
     "the answers as recorded; the 'In the cleaned data' column on Strength "
     "detail marks the rows S9 changed."],
    ["Thiamine", "'THIAMINE 100MG (VIT B1)TABLET' and 'Vitamin B1 tablet"
     "(Thiamine) 50mg' are treated as two strengths of one medicine here. "
     "The dashboard and rule S9 group them the same way."],
    ["Cases found", f"{len(CASES)} at "
     f"{len({c[0] for c in CASES})} facilities, across {len(by_mol)} "
     "medicines."],
    ["Flag", "A facility is FLAGGED when it has at least one case. A "
     "medicine whose strengths are all unavailable is NOT a case and is "
     "not flagged."],
    [f"Out > {LONG} days", "A flag whose unavailable strength was out "
     f"for more than {LONG} days over June + July + August (June, July and "
     "August stock-out days summed; the longest strength when two are "
     "out). Columns H-I of Summary by district, column O of Cases."],
    ["Excluded", "DHAP (all strengths) and AL (ARTE + LUME 1x6 / 4x6) "
     "are left out of every sheet, at the user's request (2026-11-03)."],
    ["Sheets", "Summary by district (flag counts per district, live "
     "COUNTIF formulas on the Cases and Facility flags sheets); Facility "
     "flags (one row per flagged facility); Summary by medicine; Cases (one row per facility x "
     "medicine); Strength detail (one row per strength in those cases, "
     "with stock, consumption, stock-out and requisition fields)."],
    ["Rebuild", "python3 analytics/Strength_mismatch_extract.py"],
]
sheet("Read me", ["Item", "Detail"], readme, widths={1: 18, 2: 120})
for row in wb["Read me"].iter_rows(min_row=2):
    for c in row:
        c.alignment = Alignment(wrap_text=True, vertical="top")

# 4. one row per flagged facility
fac_cases = defaultdict(list)
for c in CASES:
    fac_cases[c[0]].append(c)
fac_rows = []
for fac, cs in fac_cases.items():
    p = cs[0][2][0]["_p"]
    fac_rows.append([s(p[M.PROV]), s(p[M.DIST]), fac, FL.group(p[M.LEVEL]),
                     s(p.get(M.MGMT)), "FLAG", len(cs),
                     sum(len(c[4]) for c in cs),
                     "; ".join(mol_label(NAMES[c[1]]) for c in cs),
                     f'=COUNTIFS(Cases!$D:$D,C{len(fac_rows) + 4},'
                     f'Cases!$O:$O,"FLAG {LONG}+")'])
fac_rows.sort(key=lambda r: (r[0], r[1], r[2]))

# 5. by district - every district assessed, flags counted by formula
FACS = {}
for r in ROWS:
    FACS.setdefault(s(r["_p"][M.FAC]), (s(r["_p"][M.PROV]),
                                        s(r["_p"][M.DIST])))
dist_fac = Counter(FACS.values())
DIST = sorted(dist_fac)
FHDR, CHDR = 3, 3      # header rows on Facility flags / Cases (one note)
dist_rows = []
for i, (prov, dist) in enumerate(DIST):
    n = i + 2
    dist_rows.append([
        prov, dist, dist_fac[(prov, dist)],
        f"=COUNTIFS('Facility flags'!$B:$B,B{n},'Facility flags'!$F:$F,\"FLAG\")",
        f"=IFERROR(D{n}/C{n},0)",
        f"=COUNTIF(Cases!$B:$B,B{n})",
        f"=SUMIF('Facility flags'!$B:$B,B{n},'Facility flags'!$H:$H)",
        f"=COUNTIFS('Facility flags'!$B:$B,B{n},'Facility flags'!$J:$J,\">0\")",
        f"=COUNTIFS(Cases!$B:$B,B{n},Cases!$O:$O,\"FLAG {LONG}+\")"])
last = len(DIST) + 1
dist_rows.append(["ALL DISTRICTS", "", f"=SUM(C2:C{last})",
                  f"=SUM(D2:D{last})", f"=IFERROR(D{last + 1}/C{last + 1},0)",
                  f"=SUM(F2:F{last})", f"=SUM(G2:G{last})",
                  f"=SUM(H2:H{last})", f"=SUM(I2:I{last})"])
ws = sheet("Summary by district",
           ["Province", "District", "Facilities assessed",
            "Facilities flagged", "% of facilities flagged",
            "Flags (facility x medicine)", "Unavailable strength rows",
            f"Facilities flagged, out > {LONG} days",
            f"Flags out > {LONG} days Jun-Aug"],
           dist_rows, widths={1: 16, 2: 16, 3: 12, 4: 12, 5: 12, 6: 14,
                              7: 14})
for row in ws.iter_rows(min_row=2, max_row=last + 1):
    row[4].number_format = "0.0%"
for c in ws[last + 1]:
    c.font = Font(bold=True)
ws.auto_filter.ref = f"A1:I{last}"
for j in (8, 9):
    ws.column_dimensions[get_column_letter(j)].width = 14

sheet("Facility flags",
      ["Province", "District", "Facility", "Level of care", "Management",
       "Flag", "Medicines flagged", "Unavailable strength rows",
       "Medicines", f"Flags out > {LONG} days Jun-Aug"], fac_rows,
      widths={3: 28, 9: 70, 10: 14},
      notes=(f"{len(fac_rows)} facilities flagged: at least one medicine "
             "with one strength available and another not.",))

sheet("Summary by medicine",
      ["Medicine", "Facilities managing 2+ strengths",
       "Facilities with one strength out", "% of those facilities",
       "Most common: strength(s) available", "Most common: strength(s) out",
       "Cases with that combination"], sum_rows,
      widths={1: 40, 5: 46, 6: 46})

sheet("Cases",
      ["Province", "District", "Sector", "Facility", "Level of care",
       "Management", "Data collector", "Visit date", "Medicine", "Programme",
       "Strengths managed", "Available strength(s)",
       "Unavailable strength(s)",
       "Stock-out days Jun-Aug (longest unavailable strength)",
       f"Out > {LONG} days"], case_rows,
      widths={4: 28, 7: 22, 9: 34, 12: 50, 13: 60},
      notes=(f"{len(CASES)} facility x medicine cases. Filter any column "
             "with the header arrows.",))

sheet("Strength detail",
      ["Province", "District", "Facility", "Level of care", "Medicine",
       "Strength as assessed", "Programme", "Managed (as recorded)",
       "Available (as recorded)", "In the cleaned data", "Qty on shelf", "Qty in system", "AMC Jun-Aug", "Months of stock",
       "Stock status", "Current stock-out days", "June stock-out days",
       "July stock-out days", "August stock-out days",
       "Requested Jun-Aug", "Delivered Jun-Aug", "At RMS Regional?",
       "At RMS Central?"], item_rows,
      widths={3: 28, 5: 34, 6: 46}, status_col=9)

wb.move_sheet("Read me", offset=-(len(wb.sheetnames) - 1))
wb.calculation.fullCalcOnLoad = True
wb.active = 0
if os.path.exists(OUT):
    os.chmod(OUT, 0o644)
wb.save(OUT)
print(f"wrote {os.path.basename(OUT)}: {len(CASES)} cases at "
      f"{len({c[0] for c in CASES})} facilities, {len(item_rows)} strength "
      f"rows")
dc = Counter(c[2][0]["_p"][M.DIST] for c in CASES)
LC = [c for c in CASES if max(so_window(r) for r in c[4]) > LONG]
lc = Counter(c[2][0]["_p"][M.DIST] for c in LC)
lf = defaultdict(set)
for c in LC:
    lf[c[2][0]["_p"][M.DIST]].add(c[0])
print(f"LONG: {len(LC)} flags at {len({c[0] for c in LC})} facilities")
for prov, dist in DIST:
    print(f"  L {prov:<12} {dist:<12} {len(lf[dist]):>3} of "
          f"{dist_fac[(prov, dist)]:>3}  {lc[dist]:>3} flags")
print(Counter(mol_label(NAMES[c[1]]) for c in LC).most_common())
df = Counter(r[1] for r in fac_rows)
for prov, dist in DIST:
    print(f"  {prov:<14} {dist:<12} {df[dist]:>3} of {dist_fac[(prov, dist)]:>3}"
          f" facilities  {dc[dist]:>3} flags")
for r in sum_rows:
    print(f"  {r[0][:40]:<40} {r[2]:>3} of {r[1]:>3}  {r[3]:.0%}")


# ---------------------------------------------------------------
# export: only the flags out > LONG days (user request 2026-11-03)
# -> Strength_flags_over60.xlsx
# ---------------------------------------------------------------
OUT2 = B.out(f"Strength_flags_over{LONG}.xlsx")
wb = Workbook()
wb.remove(wb.active)
flag_rows = []
for fac, mol, managed, av, out in sorted(
        LC, key=lambda c: (s(c[2][0]["_p"][M.PROV]), s(c[2][0]["_p"][M.DIST]),
                           c[0], c[1])):
    p = managed[0]["_p"]
    for r in sorted(out, key=lambda r: -so_window(r)):
        if so_window(r) <= LONG:
            continue
        flag_rows.append([
            s(p[M.PROV]), s(p[M.DIST]), s(p[M.SECT]), fac,
            FL.group(p[M.LEVEL]), s(p.get(M.MGMT)), s(p.get("today"))[:10],
            mol_label(NAMES[mol]), s(managed[0][M.PROGRAM]),
            "; ".join(f"{s(a[M.MED])} ({num(a[M.QTY_PHYS]) or 0:,.0f} on "
                      f"shelf)" for a in av),
            s(r[M.MED]), num(r[M.SO_MONTH["June"]]) or 0,
            num(r[M.SO_MONTH["July"]]) or 0,
            num(r[M.SO_MONTH["August"]]) or 0,
            f"=L{len(flag_rows) + 2}+M{len(flag_rows) + 2}"
            f"+N{len(flag_rows) + 2}",
            num(r[M.SO_NOW]),
            "Yes" if all(was_managed(x) for x in managed) else "No",
            "Yes" if av else "No", was_avail(r),
            f'=IF(AND(Q{len(flag_rows) + 2}="Yes",R{len(flag_rows) + 2}="Yes",'
            f'S{len(flag_rows) + 2}="No",O{len(flag_rows) + 2}>{LONG}),'
            f'"FLAG","")'])
n_flag = len(flag_rows) + 1

d_rows = []
for i, (prov, dist) in enumerate(DIST):
    n = i + 2
    d_rows.append([prov, dist, dist_fac[(prov, dist)],
                   len(lf[dist]),
                   f"=IFERROR(D{n}/C{n},0)",
                   f"=COUNTIFS(Flags!$B$2:$B${n_flag},B{n},"
                   f"Flags!$T$2:$T${n_flag},\"FLAG\")"])
last = len(DIST) + 1
d_rows.append(["ALL DISTRICTS", "", f"=SUM(C2:C{last})", f"=SUM(D2:D{last})",
               f"=IFERROR(D{last + 1}/C{last + 1},0)", f"=SUM(F2:F{last})"])
ws = sheet("By district",
           ["Province", "District", "Facilities assessed",
            "Facilities flagged", "% of facilities flagged",
            "Flags (facility x medicine)"], d_rows,
           widths={1: 16, 2: 16, 3: 12, 4: 12, 5: 12, 6: 14})
for row in ws.iter_rows(min_row=2, max_row=last + 1):
    row[4].number_format = "0.0%"
for c in ws[last + 1]:
    c.font = Font(bold=True)
ws.auto_filter.ref = f"A1:F{last}"
ws.cell(last + 3, 1, (
    f"A flag = one facility x medicine where the facility manages two or "
    f"more strengths, one strength was available on the visit day, and "
    f"another was not available AND had more than {LONG} stock-out days "
    f"over June + July + August. Medicines whose strengths were all "
    f"unavailable are not flagged. DHAP and AL (ARTE + LUME) are excluded. "
    f"Facilities flagged is counted in Python (distinct facilities); "
    f"flags are counted from the Flags sheet."))
ws.cell(last + 3, 1).alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(start_row=last + 3, start_column=1, end_row=last + 3,
               end_column=6)
ws.row_dimensions[last + 3].height = 75

ws = sheet("Flags",
           ["Province", "District", "Sector", "Facility", "Level of care",
            "Management", "Visit date", "Medicine", "Programme",
            "Available strength(s)", "Unavailable strength",
            "June stock-out days", "July stock-out days",
            "August stock-out days", "Stock-out days Jun-Aug",
            "Current stock-out days", "All strengths managed",
            "Other strength available on visit day",
            "This strength available on visit day", "Flag"], flag_rows,
           widths={4: 28, 8: 32, 10: 50, 11: 44})
wb.calculation.fullCalcOnLoad = True
if os.path.exists(OUT2):
    os.chmod(OUT2, 0o644)
wb.save(OUT2)
print(f"wrote {os.path.basename(OUT2)}: {len(LC)} flags, "
      f"{len(flag_rows)} strength rows, "
      f"{len({c[0] for c in LC})} facilities")
