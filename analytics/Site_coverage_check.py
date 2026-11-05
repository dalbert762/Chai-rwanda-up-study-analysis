"""Planned sites vs what each of the three tools actually collected.

The plan and the Kobo tools name the same facility differently, so matching is
done on a normalised (core name, level) pair rather than the raw string:

  * the plan writes "Kagugu HC" where the tools write "Kagugu CS" - HC and CS
    are one level here. Level cannot simply be dropped: the plan lists
    "Muhima DH" and "Muhima HC" as two separate sites.
  * plan names carry inventory codes - "GITUZA CS--389", "NYAKIRIBA H C--1555".
  * a plan row can hold two facilities: "NYARUSIZA HC AND KIREHE HP".
  * spellings differ ("RUHENGELI" / "Ruhengeri"), so anything still unmatched
    is offered to a fuzzy pass INSIDE THE SAME DISTRICT, and every fuzzy match
    is printed for checking rather than silently accepted.

A key therefore carries the DISTRICT as well. Rwanda reuses facility names
freely - RUBONA HC exists in Huye, Ngoma and Rwamagana, CYANIKA HC in Burera
and Nyamagabe - and a name-only key silently merged them, undercounting both
the plan and every tool. The parenthetical the tools use to disambiguate
("Rubona (huye) CS") is stripped, so the district field is what keeps them
apart.
"""

import _bootstrap as B   # path setup - must come first
import difflib, os, re, sys, openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, ".")
PLAN = "Sites_to_be_visited_combined_1.xlsx"
FUZZ = 0.84

# Plan name -> tool name, for pairs the rules cannot reach. Each was confirmed
# by hand to be the same facility IN THE SAME DISTRICT. "KAGARAMA S&A SGHP"
# (Nyamasheke) is deliberately NOT aliased to "AKAGARAMA HP" (Ngoma) - similar
# name, different district, different facility.
ALIASES = {
    "gatsiborwimbogokiburara": "kiburara",   # plan carries the route, not the name
    "cusp": "cuspbutare",
    "rusatira": "rusatirakinazi",
}

LEVELS = {"hc": "hc", "cs": "hc", "dh": "dh", "hospital": "dh", "hd": "dh",
          "l2th": "l2th", "l2t": "l2th", "th": "l2th", "chub": "l2th",
          "ph": "ph", "mhc": "mhc", "hp": "hp", "sghp": "sghp",
          "sg": "sghp", "post": "hp"}
NOISE = {"health", "centre", "center", "hospital", "the", "and", "first",
         "second", "generation", "s", "a"}


# Pairs the automatic rules could only reach loosely, each CONFIRMED by the
# user on 2026-09-25 as the same facility. Keyed (plan key) -> (tool key), so
# they now match exactly and stop being flagged; anything NEW that only matches
# loosely still surfaces on the "Matched loosely" sheet, which is the point of
# keeping that sheet.
#
# Most are a level recorded differently on the two sides (a post upgraded to
# SGHP, a hospital listed as DH in the plan and L2TH in the tool); two are
# spellings (RUHENGELI/Ruhengeri) and two are a qualifier the plan omits
# (Gikomero II, Gacuba II).
CONFIRMED = {
    ("butaro", "l2th"): ("butaro", "dh"),
    ("kibagabaga", "sghp"): ("kibagabaga", "hp"),
    ("gikomero", "hc"): ("gikomeroii", "hc"),
    ("gatuna", "sghp"): ("gatuna", "hp"),
    ("ituze", "hp"): ("ituze", "sghp"),
    ("bwiza", "sghp"): ("bwiza", "hp"),
    ("bethsaida", ""): ("bethsaida", "hc"),
    ("busogo", ""): ("busogo", "hc"),
    ("ruhengeli", "l2th"): ("ruhengeri", "l2th"),
    ("kibungo", "dh"): ("kibungo", "l2th"),
    ("muhororo", "dh"): ("muhororo", "hc"),
    ("gikongoroprison", ""): ("gikongoroprison", "hc"),
    ("kirehe", "hp"): ("kirehe", "hc"),
    ("gisenyi", "dh"): ("gisenyi", "l2th"),
    ("gacubaii", ""): ("gacubaii", "hc"),
    ("ikizeremurehe", "sghp"): ("ikizeremurehe", "hp"),
}


def dkey(name, district):
    """Full key: (core name, level, district). Facility names repeat across
    districts, so the district is part of the identity, not a detail."""
    c, l = key(name)
    return c, l, re.sub(r"\s+", " ", str(district or "")).strip().lower()


def key(name):
    s = str(name or "").lower()
    s = re.sub(r"-{2,}\s*\d+", " ", s)          # "--389" inventory codes
    s = s.replace("__", " ").replace("/", " ").replace("-", " ")
    s = re.sub(r"\(.*?\)", " ", s)              # "(kicukiro)" hints
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    s = re.sub(r"\bh\s+c\b", " hc ", s)         # "H C" written apart
    toks = [t for t in s.split() if t and not t.isdigit()]
    lvl = ""
    while toks and toks[-1] in LEVELS:
        lvl = LEVELS[toks.pop()]
    core = "".join(t for t in toks if t not in LEVELS and t not in NOISE)
    return ALIASES.get(core, core), lvl


def split_row(nm):
    """'NYARUSIZA HC AND KIREHE HP' -> two facilities."""
    parts = re.split(r"\s+AND\s+", str(nm), flags=re.I)
    return [p.strip() for p in parts if p.strip()]


def load_plan():
    wb = openpyxl.load_workbook(PLAN, data_only=True, read_only=True)
    rows = list(wb["All sites"].iter_rows(values_only=True))
    h = [str(x).strip() for x in rows[0]]
    out = {}
    for r in rows[1:]:
        d = dict(zip(h, r))
        raw = str(d.get("Health facility") or "").strip()
        if not raw:
            continue
        for nm in split_row(raw):
            out.setdefault(dkey(nm, d.get("District")), {
                "name": nm, "row": raw,
                "district": str(d.get("District") or "").strip(),
                "province": str(d.get("Province") or "").strip(),
                "date": str(d.get("Date as written") or "").strip()})
    return out


def visited():
    import analyze_kobo_prescription as P
    import analyze_medicine_assessment as M
    import analyze_medicine_assessment_chw as C
    def seen(rows, col, dcol):
        d = {}
        for r in rows:
            n, dist = str(r[col]).strip(), str(r[dcol]).strip()
            d.setdefault(dkey(n, dist), (n, dist))
        return d
    return {"Prescription": seen(P.sub, P.FACIL, P.DISTR),
            "Medicine assessment": seen(M.SUB, M.FAC, M.DIST),
            "CHW": seen(C.SUB, C.FAC, C.DIST)}


def reconcile(plan, seen):
    """-> matched {plan_key: (tool_name, how)}, missing[], extra[]"""
    matched, used = {}, set()
    for k in plan:
        if k in seen:
            matched[k] = (seen[k][0], "exact")
            used.add(k)
    # user-confirmed equivalences count as exact
    for k in [x for x in plan if x not in matched]:
        c = CONFIRMED.get(k[:2])
        t = (c[0], c[1], k[2]) if c else None      # same district, by definition
        if t and t in seen and t not in used:
            matched[k] = (seen[t][0], "exact")
            used.add(t)
    # same core name, different level recorded
    for k in [x for x in plan if x not in matched]:
        for s in seen:
            if s not in used and s[0] == k[0] and s[2] == k[2]:
                matched[k] = (seen[s][0], f"level differs ({k[1] or '-'}->{s[1] or '-'})")
                used.add(s)
                break
    # fuzzy, within district only
    left = [x for x in plan if x not in matched]
    for k in left:
        cands = [s for s in seen if s not in used and s[2] == k[2]]
        best = difflib.get_close_matches(k[0], [c[0] for c in cands], n=1, cutoff=FUZZ)
        if best:
            s = next(c for c in cands if c[0] == best[0])
            matched[k] = (seen[s][0], f"fuzzy {difflib.SequenceMatcher(None, k[0], best[0]).ratio():.2f}")
            used.add(s)
    missing = [k for k in plan if k not in matched]
    extra = [s for s in seen if s not in used and s not in plan]
    return matched, missing, extra


OUT = B.out("Site_coverage_check.xlsx")

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="DDE5F0")
MISS_FILL = PatternFill("solid", fgColor="FCE4E4")   # not yet visited
OK_FILL = PatternFill("solid", fgColor="E7F3E7")     # visited
WARN_FILL = PatternFill("solid", fgColor="FFF4D6")   # matched loosely


def sheet(wb, title, headers, rows, notes=None, fills=None):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    if notes:
        for n in notes:
            c = ws.cell(row=r, column=1, value=n)
            c.font = Font(size=9, italic=True, color="555555")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r,
                           end_column=max(len(headers), 3))
            ws.row_dimensions[r].height = 13 * (1 + len(n) // 110)
            r += 1
        r += 1
    widths = {}
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font, c.fill, c.border = Font(bold=True, size=10), HDR_FILL, BORDER
        c.alignment = Alignment(wrap_text=True, vertical="bottom")
        widths[ci] = min(len(str(h)) + 2, 34)
    hdr_row = r
    for row in rows:
        r += 1
        for ci, v in enumerate(row, start=1):
            c = ws.cell(row=r, column=ci, value=v)
            c.border = BORDER
            c.alignment = Alignment(vertical="top",
                                    wrap_text=isinstance(v, str) and len(str(v)) > 40)
            if fills:
                f = fills(row, ci - 1, v)
                if f:
                    c.fill = f
            widths[ci] = max(widths.get(ci, 9), min(len(str(v)) + 2, 44))
    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)
    ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(len(headers))}{r}"
    for ci, w in widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    return ws


def build():
    plan, vis = load_plan(), visited()
    res = {t: reconcile(plan, s) for t, s in vis.items()}
    tools = list(vis)
    pdist = {p["district"].lower() for p in plan.values()}
    wb = Workbook()
    wb.remove(wb.active)

    # --- 1. summary
    rows = []
    for t in tools:
        matched, missing, extra = res[t]
        ind = [s for s in extra if vis[t][s][1].lower() in pdist]
        rows.append([t, len(vis[t]), len(plan), len(matched),
                     len(matched) / len(plan), len(missing), len(ind),
                     len(extra) - len(ind)])
    sheet(wb, "Summary",
          ["Tool", "Facilities visited", "Planned facilities", "Planned and visited",
           "% of plan", "NOT yet visited", "Visited but not planned (in a plan district)",
           "Visited in districts the plan does not cover"], rows,
          notes=["Planned sites come from Sites_to_be_visited_combined_1.xlsx, sheet "
                 "'All sites' - 170 distinct facilities across 19 districts.",
                 "The plan and the Kobo tools spell facilities differently, so matching "
                 "is on normalised name + level, with HC and CS treated as the same "
                 "level. Anything matched loosely is listed on the 'Matched loosely' "
                 "sheet - check those before relying on a count."])
    for ci, w in ((5, 11),):
        pass
    ws = wb["Summary"]
    for r in range(ws.max_row - len(rows) + 1, ws.max_row + 1):
        ws.cell(row=r, column=5).number_format = "0%"

    # --- 2. the matrix: one row per planned facility
    hdr = ["Province", "District", "Planned facility", "Planned date"]
    for t in tools:
        hdr += [t, f"{t} - recorded as"]
    rows = []
    for k in sorted(plan, key=lambda k: (plan[k]["province"], plan[k]["district"],
                                         plan[k]["name"])):
        p = plan[k]
        row = [p["province"], p["district"], p["name"], p["date"]]
        for t in tools:
            matched = res[t][0]
            if k in matched:
                nm, how = matched[k]
                row += ["Visited", nm if how == "exact" else f"{nm}  [{how}]"]
            else:
                row += ["NOT YET", ""]
        rows.append(row)

    def mfill(row, i, v):
        if v == "NOT YET":
            return MISS_FILL
        if v == "Visited":
            return OK_FILL
        if isinstance(v, str) and "[" in v:
            return WARN_FILL
        return None

    sheet(wb, "Plan coverage", hdr, rows,
          notes=["One row per planned facility. Red = not yet visited by that tool. "
                 "Amber in a 'recorded as' column = the match was loose (level differs, "
                 "or a fuzzy name match inside the same district)."],
          fills=mfill)

    # --- 3. not yet visited, long format
    rows = []
    for t in tools:
        for k in res[t][1]:
            p = plan[k]
            other = [o for o in tools if k in res[o][0]]
            rows.append([t, p["province"], p["district"], p["name"], p["date"],
                         ", ".join(other) if other else "none - no tool visited it"])
    rows.sort(key=lambda r: (r[1], r[2], r[3], r[0]))
    sheet(wb, "Not yet visited", 
          ["Tool", "Province", "District", "Planned facility", "Planned date",
           "Visited by which other tools"], rows,
          notes=["A facility appears once per tool that has not reached it. The last "
                 "column says whether another tool got there, which separates a "
                 "tool-specific gap from a site nobody has visited."])

    # --- 4. visited but not planned
    rows = []
    for t in tools:
        for s in res[t][2]:
            nm, dist = vis[t][s]
            rows.append([t, dist, nm,
                         "yes" if dist.lower() in pdist else
                         "NO - district absent from the plan"])
    rows.sort(key=lambda r: (r[3] != "yes", r[1], r[2], r[0]))
    sheet(wb, "Visited not planned",
          ["Tool", "District", "Facility as recorded", "District is in the plan?"], rows,
          notes=["The plan covers 19 districts; the tools have collected in 30. Rows "
                 "marked NO are in the 11 districts the plan does not mention at all."])

    # --- 5. loose matches to check
    rows = []
    for t in tools:
        for k, (nm, how) in res[t][0].items():
            if how != "exact":
                rows.append([t, plan[k]["district"], plan[k]["name"], nm, how])
    rows.sort(key=lambda r: (r[1], r[2], r[0]))
    sheet(wb, "Matched loosely", 
          ["Tool", "District", "Planned facility", "Matched to (as recorded)",
           "How"], rows,
          notes=["Every non-exact match. If any pair here is actually two different "
                 "facilities, the planned one belongs on 'Not yet visited' instead.",
                 "KAGARAMA S&A SGHP (Nyamasheke) was deliberately NOT matched to "
                 "AKAGARAMA HP (Ngoma) - similar name, different district."])

    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)      # Box resets synced files to read-only
    wb.save(OUT)
    print(f"wrote {OUT}")
    for t in tools:
        matched, missing, extra = res[t]
        print(f"  {t:20} {len(matched):>3}/{len(plan)} of plan "
              f"({100*len(matched)/len(plan):.0f}%) | {len(missing):>2} not yet visited "
              f"| {len(extra)} not planned")


if __name__ == "__main__":
    build()
