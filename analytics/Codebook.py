"""Codebook: every variable in every tool, with what is actually in it.

    python3 Codebook.py            -> Codebook.xlsx, all three tools
    python3 Codebook.py stock      -> print the stock tool to the terminal
    python3 Codebook.py stock qty  -> only columns matching "qty"

Answers "what variables does this tool have, and what is in them" without
opening a 14,000-row export by hand. For each column it reports the sheet,
how much of it is filled, the inferred type, how many distinct values, the
range or the commonest values, and - the part a spreadsheet cannot give you -
the CONSTANT NAME the analysis code uses for it, so a variable seen here can
be found in the scripts immediately.

Reading the export by hand is unreliable for this project in two ways the
codebook removes. Column headers carry the Kobo label, not the field name,
and several run past 100 characters with ${medicine_name} placeholders in
them, so they cannot be eyeballed or typed accurately. And a column that
looks empty in the first screenful is often asked under a skip condition and
filled in lower down - the fill rate tells you which, and against which
denominator.
"""

import _bootstrap as B   # path setup - must come first
import os
import re
import sys
from collections import Counter

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUT = B.out("Codebook.xlsx")

TOOLS = {
    "stock": ("Medicine stock assessment",
              "Medicine_stock_assessment_export.xlsx",
              "analyze_medicine_assessment"),
    "prescription": ("Unserved prescription assessment",
                     "Prescription_kobo_export.xlsx",
                     "analyze_kobo_prescription"),
    "chw": ("Medicine stock assessment for CHWs",
            "Medicine_stock_CHW_export.xlsx",
            "analyze_medicine_assessment_chw"),
}

HDR_FONT = Font(bold=True, size=9, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor="2A68BD")
TITLE_FONT = Font(bold=True, size=11)
EMPTY_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def constants(modname):
    """{column label -> CONSTANT_NAME} from the analysis module.

    The modules hold their column names as module-level strings, and some in
    dicts keyed by month, so both are walked. This is what lets the codebook
    say "this is M.QTY_PHYS" instead of leaving you to grep a 100-character
    header.
    """
    out = {}
    try:
        mod = __import__(modname)
    except Exception:
        return out
    for name in dir(mod):
        if name.startswith("_") or not name.isupper():
            continue
        v = getattr(mod, name)
        if isinstance(v, str) and len(v) > 2:
            out.setdefault(v, name)
        elif isinstance(v, dict):
            for k, vv in v.items():
                if isinstance(vv, str) and len(vv) > 2:
                    out.setdefault(vv, f"{name}[{k!r}]")
    return out


def kind(vals):
    """Inferred type from the non-empty values."""
    if not vals:
        return "empty"
    nums = sum(1 for v in vals if isinstance(v, (int, float)))
    if nums == len(vals):
        return "number"
    if nums:
        return "mixed"
    low = {str(v).strip().lower() for v in vals}
    if low <= {"yes", "no"}:
        return "yes/no"
    if low <= {"yes", "no", "dk", "na", "n/a", "don't know"}:
        return "yes/no + other"
    if len(low) <= 12:
        return "category"
    return "text"


def profile(path, sheet_name):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name]
    it = ws.iter_rows(values_only=True)
    hdr = [str(h) if h is not None else "" for h in next(it)]
    cols = [[] for _ in hdr]
    n = 0
    for row in it:
        n += 1
        for i, v in enumerate(row[:len(hdr)]):
            if v is not None and str(v).strip() != "":
                cols[i].append(v)
    wb.close()
    return hdr, cols, n


def describe(hdr, cols, n, consts):
    out = []
    for i, (h, vals) in enumerate(zip(hdr, cols), start=1):
        k = kind(vals)
        uniq = {str(v).strip() for v in vals}
        detail = ""
        if k == "number":
            nums = [float(v) for v in vals]
            detail = (f"min {min(nums):g} / median "
                      f"{sorted(nums)[len(nums) // 2]:g} / max {max(nums):g}")
        elif vals:
            top = Counter(str(v).strip() for v in vals).most_common(4)
            detail = " | ".join(f"{a} ({b:,})" for a, b in top)
        out.append([i, h, consts.get(h, ""), k, len(vals),
                    len(vals) / n if n else None, len(uniq), detail])
    return out


HEADERS = ["#", "Column label (as exported)", "Constant in the code", "Type",
           "Filled", "% filled", "Distinct", "Range or commonest values"]


def sheet(wb, title, blocks, notes=None):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    for nline in (notes or []):
        c = ws.cell(row=r, column=1, value=nline)
        c.font = Font(size=9, italic=True, color="555555")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 13 * (1 + len(nline) // 150)
        r += 1
    if notes:
        r += 1
    for btitle, rows in blocks:
        ws.cell(row=r, column=1, value=btitle).font = TITLE_FONT
        r += 1
        for ci, h in enumerate(HEADERS, start=1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
            c.alignment = Alignment(wrap_text=True, vertical="center",
                                    horizontal="center")
        hdr = r
        r += 1
        for row in rows:
            for ci, v in enumerate(row, start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.border = BORDER
                c.alignment = Alignment(vertical="top", wrap_text=ci in (2, 8))
                if ci == 6 and isinstance(v, float):
                    c.number_format = "0.0%"
                elif isinstance(v, (int, float)):
                    c.number_format = "#,##0"
                if not row[4]:
                    c.fill = EMPTY_FILL
            r += 1
        ws.freeze_panes = f"C{hdr + 1}"
        ws.auto_filter.ref = f"A{hdr}:H{r - 1}"
        r += 2
    for ci, w in {1: 5, 2: 62, 3: 22, 4: 11, 5: 9, 6: 9, 7: 9, 8: 54}.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    return ws


def build():
    wb = Workbook()
    wb.remove(wb.active)
    index = []
    for key, (label, fname, modname) in TOOLS.items():
        path = os.path.join(B.ROOT, fname)
        if not os.path.exists(path):
            continue
        consts = constants(modname)
        src = openpyxl.load_workbook(path, read_only=True)
        names = src.sheetnames
        src.close()
        blocks = []
        for sn in names:
            hdr, cols, n = profile(path, sn)
            rows = describe(hdr, cols, n, consts)
            blocks.append((f"{sn}  -  {n:,} rows, {len(hdr)} columns", rows))
            index.append([label, sn, n, len(hdr),
                          sum(1 for r in rows if not r[4]),
                          sum(1 for r in rows if r[2])])
        sheet(wb, label, blocks, notes=[
            f"{label}  -  source file {fname}",
            "One line per column. '% filled' is against the rows of that "
            "sheet, so a low figure usually means the question sits behind a "
            "skip condition rather than that the data is missing.",
            "'Constant in the code' is the name the analysis scripts use - "
            f"grep it in {modname}.py to see every place it is read.",
            "Columns with nothing in them are shaded grey.",
        ])
    sheet(wb, "Index", [("Sheets profiled", [])], notes=[
        "Codebook of all variables across the three tools."])
    ws = wb["Index"]
    ws.delete_rows(3, 4)
    ws.cell(row=3, column=1, value="Tool").font = HDR_FONT
    for ci, h in enumerate(["Tool", "Sheet", "Rows", "Columns",
                            "Empty columns", "Named in code"], start=1):
        c = ws.cell(row=3, column=ci, value=h)
        c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
        c.alignment = Alignment(wrap_text=True, horizontal="center")
    for ri, row in enumerate(index, start=4):
        for ci, v in enumerate(row, start=1):
            c = ws.cell(row=ri, column=ci, value=v)
            c.border = BORDER
            if isinstance(v, int):
                c.number_format = "#,##0"
    for ci, w in {1: 36, 2: 34, 3: 10, 4: 10, 5: 15, 6: 15}.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    wb.move_sheet("Index", offset=-len(wb.sheetnames) + 1)
    wb.save(OUT)
    return index


def show(tool, pattern=None):
    """Print one tool to the terminal - the quick lookup."""
    label, fname, modname = TOOLS[tool]
    path = os.path.join(B.ROOT, fname)
    consts = constants(modname)
    src = openpyxl.load_workbook(path, read_only=True)
    names = src.sheetnames
    src.close()
    rx = re.compile(pattern, re.I) if pattern else None
    for sn in names:
        hdr, cols, n = profile(path, sn)
        rows = describe(hdr, cols, n, consts)
        if rx:
            rows = [r for r in rows if rx.search(r[1]) or rx.search(r[2])]
        if not rows:
            continue
        print(f"\n=== {label} / {sn}  ({n:,} rows) "
              + (f"- matching {pattern!r}" if rx else ""))
        print(f"{'#':>4}  {'column':<58}{'constant':<20}{'type':<10}"
              f"{'filled':>9}  values")
        for i, h, cst, k, f, p, u, d in rows:
            print(f"{i:>4}  {h[:56]:<58}{cst[:18]:<20}{k:<10}"
                  f"{(p or 0):>8.0%}  {d[:60]}")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in TOOLS:
        show(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    else:
        idx = build()
        print(f"wrote {OUT}")
        for t, sn, n, c, empty, named in idx:
            print(f"  {t[:34]:<36}{sn[:26]:<28}{n:>7,} rows "
                  f"{c:>4} cols  {named:>3} named in code, {empty} empty")
