"""Share of health facilities whose stock-out rate is under 5%, for HIV (ARV),
TB and Malaria (request 2026-10-21).

A facility's stock-out rate for a programme = managed medicines of that
programme with at least one stock-out day, June-August / managed medicines of
that programme - the dashboard's stock-out incidence (FIG 4.1), cut per
facility. A facility enters a programme's denominator only if it manages at
least one of its medicines. "All three" = facilities under 5% in every one of
the three programmes they manage.

Writes Stockout_under5_HIV_TB_Malaria.xlsx in the project root.
"""
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import Full_indicator_report as R

s, M = R.s, R.M
PROGS = [("HIV (ARV)", "ARV"), ("TB", "TB"), ("Malaria", "Malaria")]
CUT = 0.05
OUT = "Stockout_under5_HIV_TB_Malaria.xlsx"

# facility -> programme -> [managed, with a stock-out]
g = defaultdict(lambda: defaultdict(lambda: [0, 0]))
meta = {}
for r in R.MG:
    prog = s(r[M.PROGRAM])
    if prog not in {p for _, p in PROGS}:
        continue
    f = R.mfac(r)
    g[f][prog][0] += 1
    g[f][prog][1] += R._somonths(r)
    p = r["_p"]
    meta.setdefault(f, (s(p[M.PROV]), s(p[M.DIST]), r["_lvl"]))


def rate(f, prog):
    n, so = g[f][prog]
    return so / n if n else None


def summary(facs):
    """[label, managing, under 5%, % under 5%, with no stock-out]"""
    rows = []
    for lab, prog in PROGS:
        m = [f for f in facs if g[f][prog][0]]
        u = [f for f in m if rate(f, prog) < CUT]
        z = [f for f in m if g[f][prog][1] == 0]
        rows.append([lab, len(m), len(u), R.pct(len(u), len(m)), len(z)])
    m = [f for f in facs if any(g[f][p][0] for _, p in PROGS)]
    u = [f for f in m if all(rate(f, p) < CUT for _, p in PROGS
                             if g[f][p][0])]
    z = [f for f in m if all(g[f][p][1] == 0 for _, p in PROGS)]
    rows.append(["All three programmes", len(m), len(u), R.pct(len(u), len(m)),
                 len(z)])
    return rows


HDR = ["Programme", "Facilities managing", "Facilities under 5%",
       "% of facilities under 5%", "Facilities with no stock-out"]
HEAD = PatternFill("solid", fgColor="1F4E78")


def sheet(ws, title, hdr, rows, pct_cols=()):
    ws.append([title])
    ws["A1"].font = Font(bold=True, size=13)
    ws.append([])
    ws.append(hdr)
    for c in ws[3]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for row in rows:
        ws.append(row)
        if row and str(row[0]).startswith("All three"):
            for c in ws[ws.max_row]:
                c.font = Font(bold=True)
    for col in pct_cols:
        for r in range(4, ws.max_row + 1):
            ws.cell(r, col).number_format = "0.0%"
    for i, _ in enumerate(hdr, 1):
        ws.column_dimensions[get_column_letter(i)].width = 30 if i == 1 else 16
    ws.freeze_panes = "A4"


facs = sorted(g)
wb = Workbook()
ws = wb.active
ws.title = "Summary"
sheet(ws, "Health facilities with stock-out rate under 5% - HIV, TB, "
      "Malaria (June-August 2026)", HDR, summary(facs), pct_cols=(4,))
ws.append([])
for line in (
        "Stock-out rate = managed medicines of the programme with at least "
        "one stock-out day, June-August / managed medicines of the programme.",
        "A facility counts in a programme only if it manages at least one of "
        "its medicines.",
        "All three programmes = under 5% in every one of the three it "
        "manages.",
        "With 5-6 medicines per programme, one stocked-out medicine is "
        "already 15-20%, so under 5% means no stock-out in practice."):
    ws.append([line])

for name, keyfn, keys in (
        ("By level of care", lambda f: meta[f][2], R.LEVELS),
        ("By province", lambda f: meta[f][0],
         sorted({m[0] for m in meta.values()}))):
    ws = wb.create_sheet(name)
    rows = []
    for k in keys:
        sub = [f for f in facs if keyfn(f) == k]
        if not sub:
            continue
        rows += [[k] + x for x in summary(sub)] + [[]]
    sheet(ws, f"Facilities under 5% stock-out - {name.lower()}",
          [name.replace("By ", "").capitalize()] + HDR, rows, pct_cols=(5,))

ws = wb.create_sheet("Facility detail")
hdr = ["Facility", "Province", "District", "Level of care"]
for lab, _ in PROGS:
    hdr += [f"{lab} managed", f"{lab} with stock-out", f"{lab} stock-out %",
            f"{lab} under 5%"]
rows = []
for f in facs:
    row = [f, *meta[f]]
    for _, prog in PROGS:
        n, so = g[f][prog]
        v = rate(f, prog)
        row += [n or None, so if n else None, v,
                ("Yes" if v < CUT else "No") if n else "Not managed"]
    rows.append(row)
sheet(ws, "Stock-out rate per facility - HIV, TB, Malaria", hdr, rows,
      pct_cols=(7, 11, 15))
for i in range(1, len(hdr) + 1):
    ws.column_dimensions[get_column_letter(i)].width = 30 if i == 1 else 14

wb.save(R.os.path.join(R.os.getcwd(), OUT) if hasattr(R, "os") else OUT)
for row in summary(facs):
    print(row)
