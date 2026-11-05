"""Where the stock was when the facility did not have it: RMS Regional and
RMS Central, by drug class, by programme and by level of care.

    python3 Upstream_availability.py  ->  Upstream_availability.xlsx

When a medicine is not on the shelf, the form asks whether it was available
at RMS Regional and at RMS Central. That pair of answers is what separates a
facility-level failure from a national one:

    available UPSTREAM      the stock existed and did not reach the facility
                            - ordering, transport, funds or forecasting
    not available UPSTREAM  the warehouse was empty too - a procurement or
                            national supply problem, which no amount of
                            facility-level action fixes

THE DENOMINATOR IS THE ANSWERED ROWS, NOT ALL UNAVAILABLE ROWS
The upstream questions are skipped on roughly a fifth of unavailable rows,
and "Don't know" is a real answer that is not evidence either way. Both are
excluded from the rates and both are reported, so a group's rate is never
computed over people who did not answer.

THIS IS A RECALL ANSWER, NOT A WAREHOUSE RECORD
Nobody checked the RMS stock ledger. A storekeeper was asked whether the
medicine was available upstream, at a facility that had just run out. Read
these as the facility's belief about the warehouse, which is the only
upstream evidence this study collected.
"""

import _bootstrap as B   # path setup - must come first
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_medicine_assessment as M
import explorer as EX
import facility_levels as FL

OUT = B.out("Upstream_availability.xlsx")
MIN_N = 10

HDR_FONT = Font(bold=True, size=9, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="2A68BD")
TITLE_FONT = Font(bold=True, size=11)
TOT_FILL = PatternFill("solid", fgColor="EAEAEA")
BAD_FILL = PatternFill("solid", fgColor="F6C9C4")
MID_FILL = PatternFill("solid", fgColor="FBE3BA")
OK_FILL = PatternFill("solid", fgColor="CDE8DA")
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def s(v):
    return "" if v is None else str(v).strip()


def pct(a, b):
    return a / b if b else None


# ---------------------------------------------------------------
# rows
# ---------------------------------------------------------------
NA = [r for r in M.MEDROWS if M.no(r[M.AVAIL])]
for r in NA:
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])
    r["_mgt"] = s(r["_p"].get(EX.MGMT)) or "(not recorded)"
    r["_cls"] = s(r[M.DCLASS]) or "(unclassified)"
    r["_prog"] = s(r[M.PROGRAM]) or "(none)"


def answered(r, col):
    """A usable answer: Yes or No. 'Don't know' and blank are not."""
    return M.yes(r[col]) or M.no(r[col])


HDR = ["Group", "Unavailable at the facility",
       "REGIONAL: answered", "REGIONAL: available", "REGIONAL %",
       "CENTRAL: answered", "CENTRAL: available", "CENTRAL %",
       "Available at EITHER", "Available at NEITHER",
       "UPSTREAM HAD IT %"]
PCT = {4, 7, 10}


def line(rows, label):
    if not rows:
        return None
    ra = [r for r in rows if answered(r, M.RMS_REG)]
    ca = [r for r in rows if answered(r, M.RMS_CEN)]
    ry = sum(1 for r in ra if M.yes(r[M.RMS_REG]))
    cy = sum(1 for r in ca if M.yes(r[M.RMS_CEN]))
    either = [r for r in rows
              if answered(r, M.RMS_REG) or answered(r, M.RMS_CEN)]
    got = sum(1 for r in either
              if M.yes(r[M.RMS_REG]) or M.yes(r[M.RMS_CEN]))
    return [label, len(rows),
            len(ra), ry, pct(ry, len(ra)) if len(ra) >= MIN_N else None,
            len(ca), cy, pct(cy, len(ca)) if len(ca) >= MIN_N else None,
            got, len(either) - got,
            pct(got, len(either)) if len(either) >= MIN_N else None]


def by(keyfn, keys=None, rows=None, sort=True):
    g = defaultdict(list)
    for r in (rows if rows is not None else NA):
        g[keyfn(r)].append(r)
    out = [line(g[k], k) for k in (keys or sorted(g)) if g.get(k)]
    if sort:
        out.sort(key=lambda x: (x[10] is None, x[10]))
    return out


def sheet(wb, title, blocks, notes=None, widths=None):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    for n in (notes or []):
        c = ws.cell(row=r, column=1, value=n)
        c.font = Font(size=9, italic=True, color="555555")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
        ws.row_dimensions[r].height = 13 * (1 + len(n) // 140)
        r += 1
    if notes:
        r += 1
    for btitle, headers, rows, total, pctcols in blocks:
        if btitle:
            ws.cell(row=r, column=1, value=btitle).font = TITLE_FONT
            r += 1
        for ci, h in enumerate(headers, start=1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
            c.alignment = Alignment(wrap_text=True, vertical="center",
                                    horizontal="center")
        ws.row_dimensions[r].height = 44
        hdr = r
        r += 1
        for row in list(rows) + ([total] if total else []):
            is_tot = total is not None and row is total
            for ci, v in enumerate(row, start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.border = BORDER
                c.alignment = Alignment(vertical="top",
                                        wrap_text=isinstance(v, str)
                                        and len(str(v)) > 30)
                if (ci - 1) in pctcols and isinstance(v, float):
                    c.number_format = "0.0%"
                    # HIGH is bad here: the stock was upstream and did not
                    # reach the facility.
                    c.fill = (BAD_FILL if v >= 0.80
                              else MID_FILL if v >= 0.50 else OK_FILL)
                elif isinstance(v, (int, float)):
                    c.number_format = "#,##0"
                if is_tot:
                    c.font = Font(bold=True)
                    if not ((ci - 1) in pctcols and isinstance(v, float)):
                        c.fill = TOT_FILL
            r += 1
        ws.freeze_panes = f"B{hdr + 1}"
        ws.auto_filter.ref = f"A{hdr}:{get_column_letter(len(headers))}{r - 1}"
        r += 2
    for ci, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    return ws


def build():
    wb = Workbook()
    wb.remove(wb.active)
    nat = line(NA, "ALL UNAVAILABLE MEDICINES")
    wide = {1: 40, 2: 15, 3: 13, 4: 13, 5: 12, 6: 13, 7: 13, 8: 12, 9: 13,
            10: 13, 11: 14}
    NOTES = [
        "Rows where the medicine was NOT on the shelf, and what the facility "
        "said about RMS Regional and RMS Central.",
        "RATES ARE OVER ANSWERED ROWS ONLY. 'Don't know' and unanswered are "
        "excluded from every percentage and shown separately - a group's rate "
        "is never computed over people who did not answer.",
        "UPSTREAM HAD IT % is the share where the stock existed at Regional "
        "or Central: a HIGH figure means the medicine was in the country and "
        "did not reach the facility, which is an ordering, transport or "
        "funding problem rather than a procurement one.",
        "This is a recall answer from the facility, not a warehouse record. "
        "No RMS ledger was consulted.",
    ]

    # --- summary
    ra = [r for r in NA if answered(r, M.RMS_REG)]
    ca = [r for r in NA if answered(r, M.RMS_CEN)]
    sheet(wb, "Summary", [
        ("Where the stock was", ["Measure", "Rows", "Share"],
         [["Medicines not available at the facility", len(NA), ""],
          ["", "", ""],
          ["RMS REGIONAL - usable answer", len(ra), pct(len(ra), len(NA))],
          ["   available at Regional",
           sum(1 for r in ra if M.yes(r[M.RMS_REG])),
           pct(sum(1 for r in ra if M.yes(r[M.RMS_REG])), len(ra))],
          ["   NOT available at Regional",
           sum(1 for r in ra if M.no(r[M.RMS_REG])),
           pct(sum(1 for r in ra if M.no(r[M.RMS_REG])), len(ra))],
          ["   don't know",
           sum(1 for r in NA if s(r[M.RMS_REG]).lower().startswith("don")), ""],
          ["   not asked / blank",
           sum(1 for r in NA if M.blank(r[M.RMS_REG])), ""],
          ["", "", ""],
          ["RMS CENTRAL - usable answer", len(ca), pct(len(ca), len(NA))],
          ["   available at Central",
           sum(1 for r in ca if M.yes(r[M.RMS_CEN])),
           pct(sum(1 for r in ca if M.yes(r[M.RMS_CEN])), len(ca))],
          ["   NOT available at Central",
           sum(1 for r in ca if M.no(r[M.RMS_CEN])),
           pct(sum(1 for r in ca if M.no(r[M.RMS_CEN])), len(ca))],
          ["   don't know",
           sum(1 for r in NA if s(r[M.RMS_CEN]).lower().startswith("don")), ""],
          ["   not asked / blank",
           sum(1 for r in NA if M.blank(r[M.RMS_CEN])), ""],
          ["", "", ""],
          ["ATTRIBUTION - stock existed upstream", nat[8],
           pct(nat[8], nat[8] + nat[9])],
          ["ATTRIBUTION - empty upstream too", nat[9],
           pct(nat[9], nat[8] + nat[9])]],
         None, {2}),
    ], notes=NOTES, widths={1: 44, 2: 14, 3: 12})

    sheet(wb, "By drug class", [
        ("Unavailable medicines by therapeutic drug class", HDR,
         by(lambda r: r["_cls"]), nat, PCT)],
        notes=NOTES + [f"Classes with fewer than {MIN_N} answered rows show "
                       "counts only."],
        widths=wide)

    sheet(wb, "By programme", [
        ("Unavailable medicines by programme", HDR,
         by(lambda r: r["_prog"]), nat, PCT)],
        notes=NOTES, widths=wide)

    sheet(wb, "By facility level", [
        ("Unavailable medicines by level of care", HDR,
         by(lambda r: r["_lvl"], keys=FL.ordered({r["_lvl"] for r in NA}),
            sort=False), nat, PCT),
        ("By management type", HDR, by(lambda r: r["_mgt"]), nat, PCT),
        ("By province", HDR, by(lambda r: s(r["_p"][M.PROV])), nat, PCT)],
        notes=NOTES, widths=wide)

    sheet(wb, "By medicine", [
        ("Every medicine that was unavailable somewhere", HDR,
         by(lambda r: s(r[M.MED])), nat, PCT)],
        notes=NOTES, widths=wide)

    # --- the rows themselves
    det = sorted([[s(r["_p"][M.FAC]), s(r["_p"][M.DIST]), s(r["_p"][M.PROV]),
                   r["_lvl"], r["_mgt"], s(r[M.MED]), r["_prog"], r["_cls"],
                   s(r[M.RMS_REG]) or "(not asked)",
                   s(r[M.RMS_CEN]) or "(not asked)",
                   "upstream had it"
                   if (M.yes(r[M.RMS_REG]) or M.yes(r[M.RMS_CEN]))
                   else ("empty upstream too"
                         if (answered(r, M.RMS_REG) or answered(r, M.RMS_CEN))
                         else "no usable answer")]
                  for r in NA], key=lambda x: (x[10], x[0], x[5]))
    sheet(wb, "Every unavailable row", [
        ("One line per facility x medicine that was not on the shelf",
         ["Facility", "District", "Province", "Level of care", "Management",
          "Medicine", "Programme", "Drug class", "At RMS Regional",
          "At RMS Central", "Attribution"], det, None, set())],
        notes=[f"{len(NA):,} rows."],
        widths={1: 30, 2: 15, 3: 13, 4: 24, 5: 24, 6: 46, 7: 20, 8: 28,
                9: 16, 10: 16, 11: 20})

    wb.save(OUT)
    return nat


if __name__ == "__main__":
    nat = build()
    print(f"wrote {OUT}")
    print(f"  {nat[1]:,} unavailable facility x medicine rows")
    print(f"  Regional: {nat[3]:,} of {nat[2]:,} answered had it "
          f"({nat[4]:.1%})")
    print(f"  Central : {nat[6]:,} of {nat[5]:,} answered had it "
          f"({nat[7]:.1%})")
    print(f"  upstream had it {nat[8]:,} / empty upstream {nat[9]:,} "
          f"= {nat[10]:.1%} attributable to distribution, not procurement")
