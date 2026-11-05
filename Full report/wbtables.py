"""Read indicator tables from the dashboard workbook (Full_indicator_set.xlsx) by id."""
import os
import re
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WB = load_workbook(os.path.join(ROOT, "analytics", "Full_indicator_set.xlsx"), read_only=True)
_T = {}
for ws in WB.worksheets[1:10]:
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    i = 0
    while i < len(rows):
        r = rows[i]
        m = r and r[0] and re.match(r"^(\d+\.\d+[a-z]?)\s{2}(.*)", str(r[0]))
        if m:
            tid, title = m.group(1), m.group(2).strip()
            j = i + 1
            note = []
            while j < len(rows) and (not rows[j] or sum(x is not None for x in rows[j]) <= 1):
                if rows[j] and rows[j][0]:
                    note.append(str(rows[j][0]))
                j += 1
            hdr = [str(x) for x in rows[j] if x is not None] if j < len(rows) else []
            body = []
            k = j + 1
            while k < len(rows) and rows[k] and any(x is not None for x in rows[k]) and not (
                    rows[k][0] and re.match(r"^\d+\.\d+[a-z]?\s{2}", str(rows[k][0]))):
                body.append(rows[k][:len(hdr)])
                k += 1
            _T[tid] = dict(title=title, note=" ".join(note), hdr=hdr, rows=body)
            i = k
        else:
            i += 1


PCT_ROW = ("RATE", "INCIDENCE", "AVAILABILITY", "ACCURACY", "SHARE")


def fmt(v, h, label=""):
    if v is None:
        return ""
    if not isinstance(v, (int, float)):
        return str(v)
    hl = h.lower()
    if "percentage points" in hl:
        return f"{100 * v:+.1f} pts"
    if "ratio" in hl:
        return f"{v:.2f}"
    bad = any(k in hl for k in ("days", "months", "cv", "rwf", "median absolute", "units"))
    pctcol = any(k in hl for k in ("%", "rate", "share", "availab", "accura", "incidence",
                                    "fill", "cumulative"))
    if hl == "value" and any(k in str(label).upper() for k in PCT_ROW):
        pctcol = True
    if pctcol and not bad and -1.5 <= v <= 1.5:
        return f"{100 * v:.1f}%"
    if isinstance(v, float):
        return f"{v:,.1f}"
    return f"{v:,}"


def table(tid, max_rows=None, cols=None, rename=None):
    t = _T[tid]
    hdr = t["hdr"]
    idx = list(range(len(hdr))) if cols is None else cols
    rows = t["rows"][:max_rows] if max_rows else t["rows"]
    out = [[fmt(r[i] if i < len(r) else None, hdr[i], r[0]) for i in idx] for r in rows
           if r and r[0] not in (None, "") and any(x not in (None, "") for x in r)]
    h = [hdr[i] for i in idx]
    if rename:
        h = [rename.get(x, x) for x in h]
    return t["title"], h, out, t["note"]


def ids():
    return list(_T)
