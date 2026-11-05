"""Flag facilities and CHWs that manage a first-line antimalarial outside
their district's MFT zone (data-cleaning request, 2026-11-02).

Rwanda allocates one first-line antimalarial per district (multiple first-
line therapies): AL (artemether-lumefantrine), ASPY (artesunate-
pyronaridine) or DHAP (dihydroartemisinin-piperaquine).

The ZONING sheet is the editable input. Its zone column starts as the zone
INFERRED from the data - the antimalarial most facilities in the district
say they manage - because the official list was only received as a cropped
picture. Typing the official zone there (one code, or two as "AL;ASPY" where
a district is allowed both) recalculates every flag: the flag sheets look
the zone up with INDEX/MATCH and test it with SEARCH.

Writes Antimalarial_zoning_flags.xlsx in the project root.
"""
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import Full_indicator_report as R

s, M, C = R.s, R.M, R.C
OUT = "Antimalarial_zoning_flags_v3.xlsx"
GROUPS = {"AL": "AL (artemether-lumefantrine)",
          "ASPY": "ASPY (artesunate-pyronaridine)",
          "DHAP": "DHAP (dihydroartemisinin-piperaquine)"}
# first districts of each row, as visible in the user's cropped picture
PICTURE = {"Nyagatare": "AL", "Gatsibo": "AL", "Kayonza": "AL",
           "Gasabo": "ASPY", "Kicukiro": "ASPY", "Nyarugenge": "ASPY",
           "Muhanga": "DHAP", "Kamonyi": "DHAP", "Ruhango": "DHAP"}


def avail_txt(yes, no, v):
    return "Yes" if yes(v) else ("No" if no(v) else "(blank)")


def group(name):
    n = str(name).upper()
    if "ARTE" in n and "LUME" in n or n.startswith("ALU "):
        return "AL"
    if "PYRONARIDINE" in n or n.startswith("PYR/ART"):
        return "ASPY"
    if "DIHYDROARTEMISININ" in n or n.startswith("DHAP"):
        return "DHAP"
    return None


# rows S10 recoded are read back as recorded; every other cleaning rule
# (S6, S7, S9) stands
S10M = {id(r) for r in getattr(M, "ZONE_RECODED", [])}
S10C = {id(r) for r in getattr(C, "ZONE_RECODED", [])}


def _was(r, key, col, ids):
    """S10 only fires on managed = Yes / available = No (after S7), so a
    recoded row is read back as exactly that."""
    if id(r) not in ids:
        return r[col]
    return "Yes" if key == "_managed_was" else "No"


# ---- facility stock tool: facility x antimalarial group it manages
fac = defaultdict(lambda: {"prods": [], "avail": [], "ids": set()})
meta, dist_fac = {}, defaultdict(set)
frec = []          # one record per managed antimalarial product
for r in M.MEDROWS:
    g = group(s(r[M.MED]))
    if not g:
        continue
    p, f = r["_p"], R.mfac(r)
    d = s(p[M.DIST])
    meta[f] = (d, s(p[M.PROV]), r["_lvl"])
    dist_fac[d].add(f)
    # answers AS RECORDED: S10 recodes the flagged rows to not managed, so
    # read them back from _managed_was / _available_was
    if M.yes(_was(r, "_managed_was", M.MANAGED, S10M)):
        x = fac[(f, g)]
        x["prods"].append(s(r[M.MED]))
        x["avail"].append(M.yes(_was(r, "_available_was", M.AVAIL, S10M)))
        x["ids"].add(str(r.get("_submission__uuid") or p.get("_uuid")))
        frec.append([d, s(p[M.PROV]), f, r["_lvl"], g, s(r[M.MED]),
                     avail_txt(M.yes, M.no,
                               _was(r, "_available_was", M.AVAIL, S10M)),
                     str(r.get("_submission__uuid") or p.get("_uuid"))])

# ---- inferred zone: the group most facilities in the district manage
cnt = defaultdict(lambda: defaultdict(int))
for (f, g) in fac:
    cnt[meta[f][0]][g] += 1
zone, basis = {}, {}
REGION = {"East": "DHAP", "Kigali City": "DHAP", "South": "AL",
          "North": "ASPY", "West": "ASPY"}
dprov = {d: meta[f][1] for f, (d, _, _) in meta.items()}
for d in sorted(dist_fac):
    c = cnt[d]
    top = sorted(c.items(), key=lambda kv: -kv[1])
    if not top:
        zone[d], basis[d] = REGION.get(dprov[d], ""), "no facility manages one"
    elif len(top) > 1 and top[0][1] == top[1][1]:
        zone[d] = REGION.get(dprov[d], top[0][0])
        basis[d] = "inferred - TIE in the data, please confirm"
    else:
        zone[d] = top[0][0]
        second = top[1][1] if len(top) > 1 else 0
        basis[d] = ("inferred from data" if second <= len(dist_fac[d]) / 2
                    else "inferred - split in the data, please confirm")

# ---- the user's official zoning overrides the inference wherever it is
# known (2026-11-02: "ASPY should be the one in zone" for Kigali; the other
# six are the districts visible in the user's zoning picture)
OFFICIAL = dict(PICTURE)
for d, zv in OFFICIAL.items():
    if d in zone:
        zone[d], basis[d] = zv, "OFFICIAL - user's zoning list"
# keep the cleaning rule (S10) and this workbook on one list
import mft_zoning as _ZN
assert all(_ZN.ZONE[d] == zone[d] for d in _ZN.OFFICIAL)

# ---- CHWs: CHW x antimalarial group managed
chw = defaultdict(lambda: {"prods": [], "avail": [], "ids": set()})
cmeta, crec = {}, []
for r in C.MEDROWS:
    g = group(r["_short"])
    if not g or not C.yes(_was(r, "_managed_was", C.MANAGED, S10C)):
        continue
    p = r["_p"]
    k = (C.s(p[C.FAC]), C.s(p[C.CHW]), g)
    cmeta[k[:2]] = (C.s(p[C.DIST]), C.s(p[C.PROV]))
    chw[k]["prods"].append(r["_short"])
    chw[k]["avail"].append(C.yes(_was(r, "_available_was", C.AVAIL, S10C)))
    chw[k]["ids"].add(str(p.get("_uuid")))
    crec.append([C.s(p[C.DIST]), C.s(p[C.PROV]), k[0], k[1], g,
                 r["_short"], avail_txt(C.yes, C.no,
                                        _was(r, "_available_was", C.AVAIL,
                                             S10C)),
                 str(p.get("_uuid"))])

# ---------------------------------------------------------------- workbook
F = "Arial"
HEAD = PatternFill("solid", fgColor="1F4E78")
INPUT = PatternFill("solid", fgColor="FFFF00")
RED = PatternFill("solid", fgColor="F8CBAD")
GREEN = PatternFill("solid", fgColor="C6EFCE")
wb = Workbook()


def header(ws, row, cols, widths):
    for i, (h, w) in enumerate(zip(cols, widths), 1):
        c = ws.cell(row, i, h)
        c.font = Font(name=F, bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(i)].width = w


def body_font(ws):
    for row in ws.iter_rows():
        for c in row:
            if c.font is None or not c.font.bold:
                c.font = Font(name=F, size=10, bold=c.font.bold if c.font
                              else False, color=c.font.color if c.font
                              else None)


# 1. README
ws = wb.active
ws.title = "README"
lines = [
    ("Antimalarial MFT zoning - data-cleaning flags", True),
    ("", False),
    ("Purpose: flag every facility (stock tool) and CHW that says it manages "
     "a first-line antimalarial that is not allocated to its district.", False),
    ("", False),
    ("HOW TO USE", True),
    ("1. Open the ZONING sheet. Column C (yellow) holds each district's zone: "
     "AL, ASPY or DHAP.", False),
    ("2. Type the official zone over the inferred one. A district allowed two "
     "products takes both, e.g. AL;ASPY.", False),
    ("3. Every flag in FACILITY FLAGS, CHW FLAGS and SUMMARY recalculates.",
     False),
    ("", False),
    ("WHY THE ZONES ARE INFERRED", True),
    ("The official list arrived as a cropped picture showing only the first "
     "districts of each row. Column C was therefore filled with the "
     "antimalarial most facilities in the district say they manage.", False),
    ("The picture and the data disagree on all 9 visible districts (column "
     "G). Example: the picture puts Nyagatare in AL, but 9 of its 10 "
     "facilities manage DHAP and none manages AL.", False),
    ("The 9 districts visible in the picture now carry the OFFICIAL zone "
     "(Basis column). The other 21 are still inferred - replace them with "
     "the official list before using their flags.", True),
    ("", False),
    ("Codes: AL = artemether-lumefantrine; ASPY = artesunate-pyronaridine; "
     "DHAP = dihydroartemisinin-piperaquine.", False),
    ("One flag row = one antimalarial product (each strength) a facility or "
     "CHW says it manages; the Kobo uuid locates the record.", False),
    ("Zone flag: OUT OF ZONE when the product's group (AL / ASPY / DHAP) is "
     "not in the district's zone.", False),
    ("Out of zone AND not available: FLAG when the product is out of zone AND "
     "was recorded not available on the visit day - managed by the "
     "facility's own answer, absent from the shelf, and not allocated to "
     "the zone.", False),
]
for i, (t, b) in enumerate(lines, 1):
    ws.cell(i, 1, t).font = Font(name=F, bold=b, size=13 if i == 1 else 10)
ws.column_dimensions["A"].width = 120

# 2. ZONING (input)
z = wb.create_sheet("ZONING")
cols = ["District", "Province", "Zone (EDIT: AL / ASPY / DHAP)", "Basis",
        "Facilities assessed", "Facilities managing AL / ASPY / DHAP",
        "Per the cropped picture"]
header(z, 1, cols, [14, 13, 20, 40, 11, 22, 16])
for i, d in enumerate(sorted(dist_fac), 2):
    c = cnt[d]
    z.cell(i, 1, d)
    z.cell(i, 2, dprov[d])
    zc = z.cell(i, 3, zone[d])
    zc.fill = INPUT
    zc.font = Font(name=F, size=10, bold=True, color="0000FF")
    z.cell(i, 4, basis[d])
    z.cell(i, 5, len(dist_fac[d]))
    z.cell(i, 6, f"{c['AL']} / {c['ASPY']} / {c['DHAP']}")
    pic = PICTURE.get(d, "")
    pc = z.cell(i, 7, pic)
    if pic and pic != zone[d]:
        pc.fill = RED
        pc.comment = Comment(f"Cropped picture: {pic}. Data: most facilities "
                             f"manage {zone[d]}.", "Data check")
ZN = len(dist_fac) + 1
z.freeze_panes = "A2"


def flag_sheet(name, cols, widths, rows):
    """rows: District, Province, <2 id cols>, Antimalarial, Product,
    Available on visit day, uuid - then 3 formula columns: zone, zone flag,
    and the combined flag (out of zone AND not available)."""
    ws = wb.create_sheet(name)
    cols = cols + ["District zone", "Zone flag",
                   "Out of zone AND not available"]
    widths = widths + [12, 14, 30]
    header(ws, 1, cols, widths)
    n = len(cols)
    zl, fl, xl = (get_column_letter(n - 2), get_column_letter(n - 1),
                  get_column_letter(n))
    for i, row in enumerate(rows, 2):
        for j, v in enumerate(row, 1):
            ws.cell(i, j, v)
        ws.cell(i, n - 2,
                f'=IFERROR(INDEX(ZONING!$C$2:$C${ZN},MATCH(A{i},'
                f'ZONING!$A$2:$A${ZN},0)),"(district not in ZONING)")')
        ws.cell(i, n - 1,
                f'=IF(ISNUMBER(SEARCH(E{i},{zl}{i})),"IN ZONE",'
                f'"OUT OF ZONE")')
        ws.cell(i, n, f'=IF(AND({fl}{i}="OUT OF ZONE",G{i}="No"),'
                      f'"FLAG","")')
    last = len(rows) + 1
    ws.conditional_formatting.add(f"{fl}2:{fl}{last}", CellIsRule(
        operator="equal", formula=['"OUT OF ZONE"'], fill=RED))
    ws.conditional_formatting.add(f"{fl}2:{fl}{last}", CellIsRule(
        operator="equal", formula=['"IN ZONE"'], fill=GREEN))
    ws.conditional_formatting.add(f"{xl}2:{xl}{last}", CellIsRule(
        operator="equal", formula=['"FLAG"'], fill=PatternFill(
            "solid", fgColor="C00000"), font=Font(color="FFFFFF", bold=True)))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{xl}{last}"
    return ws, last


# 3. FACILITY FLAGS - one row per managed antimalarial product record
frows = sorted(frec, key=lambda r: (r[0], r[2], r[4], r[5]))
fws, flast = flag_sheet(
    "FACILITY FLAGS",
    ["District", "Province", "Facility", "Level of care", "Antimalarial",
     "Product", "Available on visit day", "Kobo submission uuid"],
    [12, 11, 26, 18, 9, 52, 10, 40], frows)

# 4. CHW FLAGS - one row per managed antimalarial product record
crows = sorted(crec, key=lambda r: (r[0], r[2], r[3], r[4], r[5]))
cws, clast = flag_sheet(
    "CHW FLAGS",
    ["District", "Province", "Health centre", "CHW", "Antimalarial",
     "Product", "Available on visit day", "Kobo submission uuid"],
    [12, 11, 22, 24, 9, 16, 10, 40], crows)

# 5. SUMMARY (formulas). Flag sheets: A district, I zone, J zone flag,
# K out of zone AND not available.
sm = wb.create_sheet("SUMMARY", 1)
header(sm, 1, ["District", "Zone", "Facility records", "Facility OUT OF ZONE",
               "Facility out of zone AND not available", "CHW records",
               "CHW OUT OF ZONE", "CHW out of zone AND not available"],
       [14, 10, 11, 13, 18, 10, 12, 18])
FF, CF = "'FACILITY FLAGS'", "'CHW FLAGS'"
for i, d in enumerate(sorted(dist_fac), 2):
    sm.cell(i, 1, d)
    sm.cell(i, 2, f"=ZONING!C{i}")
    for j, (sh, lst) in enumerate(((FF, flast), (CF, clast))):
        c0 = 3 + j * 3
        sm.cell(i, c0, f"=COUNTIFS({sh}!$A$2:$A${lst},A{i})")
        sm.cell(i, c0 + 1, f'=COUNTIFS({sh}!$A$2:$A${lst},A{i},'
                           f'{sh}!$J$2:$J${lst},"OUT OF ZONE")')
        sm.cell(i, c0 + 2, f'=COUNTIFS({sh}!$A$2:$A${lst},A{i},'
                           f'{sh}!$K$2:$K${lst},"FLAG")')
t_ = ZN + 1
sm.cell(t_, 1, "TOTAL").font = Font(name=F, bold=True)
for j in range(3, 9):
    L = get_column_letter(j)
    sm.cell(t_, j, f"=SUM({L}2:{L}{ZN})").font = Font(name=F, bold=True)
sm.freeze_panes = "A2"

for w in wb.worksheets:
    for row in w.iter_rows():
        for c in row:
            if c.row > 1 or w.title == "README":
                if not (c.font and c.font.bold) and w.title != "README":
                    c.font = Font(name=F, size=10)
wb.calculation.fullCalcOnLoad = True
wb.save(R.os.path.join(R.os.getcwd(), OUT) if hasattr(R, "os") else OUT)

# expected counts (Python check of the formulas, with the inferred zones)
for lab, rows in (("facility", frows), ("CHW", crows)):
    oz = [r for r in rows if r[4] not in zone[r[0]].split(";")]
    na = [r for r in oz if r[6] == "No"]
    print(f"{lab}: records {len(rows)}, out of zone {len(oz)}, out of zone "
          f"AND not available {len(na)} at {len({r[2] for r in na})} "
          f"{'facilities' if lab == 'facility' else 'health centres'}"
          + (f" / {len({(r[2], r[3]) for r in na})} CHWs" if lab == "CHW"
             else ""))
    for r in na:
        print("    ", r[0], "|", r[2], "|", r[3] if lab == "CHW" else "",
              "|", r[5], "| zone", zone[r[0]])
print("zones:", {d: (zone[d], basis[d]) for d in sorted(zone)
                 if "confirm" in basis[d]})
