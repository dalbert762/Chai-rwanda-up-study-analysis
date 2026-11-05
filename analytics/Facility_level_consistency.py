"""Does each facility's NAME agree with the LEVEL recorded against it?

    python3 Facility_level_consistency.py
      -> Facility_level_consistency.xlsx

Rwandan facility names carry their level: a name ending "CS" is a centre de
sante (health centre), "DH" a district hospital, "HP" a health post, and so
on. The level is also a separate answer on the form, so the two can disagree -
and when they do, every breakdown by level of care is wrong for that facility.

Three checks, as asked:

  1. name ends CS (or HC) but the level is NOT a health centre
  2. name ends in a hospital token (DH, L2TH, PH, RH, CHUB, CHUK, or the word
     "hospital") but the level is NOT any kind of hospital
  3. name ends HP or SGHP but the level is NOT a health post - i.e. not
     "Second Generation Health Post (SGHP)" and not a plain health post

Run across all three tools, because the same facility can be typed one way in
one tool and another way in the next.

A fourth, softer finding comes out of the same comparison: the health-post
level is recorded under THREE different labels that all mean the same thing
("Second Generation Health Post (SGHP)", "Health Post (HP)", and a bare "HP").
That is not a level error, but it splits health posts across three rows in
every by-level table, so it is reported separately.

openpyxl + the standard library, like the rest of this project.
"""

import _bootstrap as B   # path setup - must come first
import os
from collections import Counter, defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import analyze_kobo_prescription as P
import analyze_medicine_assessment as M
import analyze_medicine_assessment_chw as C

OUT = B.out("Facility_level_consistency.xlsx")
OUT_FLAGS = B.out("Facility_level_flags.xlsx")

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(bold=True, size=10, color="FFFFFF")
BAD_FILL = PatternFill("solid", fgColor="FCE4E4")
WARN_FILL = PatternFill("solid", fgColor="FFF4D6")
OK_FILL = PatternFill("solid", fgColor="E7F3E7")
TITLE_FONT = Font(bold=True, size=12, color="1F3864")

# what a name token says the facility is
NAME_CLASS = {
    "CS": "health centre", "HC": "health centre",
    "MHC": "maternity health centre",
    "DH": "hospital", "L2TH": "hospital", "L2T": "hospital", "TH": "hospital",
    "PH": "hospital", "RH": "hospital", "CHUB": "hospital",
    "CHUK": "hospital", "HOSPITAL": "hospital",
    "HP": "health post", "SGHP": "health post",
}
# what a recorded level value means. Both tools' spellings are listed; a value
# that appears here under a different class is what the check is looking for.
LEVEL_CLASS = {
    "health centre (hc)": "health centre", "hc": "health centre",
    "mhc (health-centre level)": "maternity health centre", "mhc":
        "maternity health centre",
    "district hospital": "hospital", "dh": "hospital",
    "level 2 teaching hospital": "hospital", "l2th": "hospital",
    "provincial hospital": "hospital", "ph": "hospital",
    "national referral / teaching hospital": "hospital",
    "national referral/teaching hospital": "hospital",
    "referral hospital": "hospital", "rh": "hospital",
    "second generation health post (sghp)": "health post", "sghp":
        "health post",
    "health post (hp)": "health post", "hp": "health post",
}
# the three interchangeable health-post labels
POST_LABELS = {"second generation health post (sghp)", "health post (hp)",
               "hp", "sghp"}

TOOLS = [
    ("Medicine assessment", M.SUB, M.FAC, M.LEVEL, M.DIST, M.PROV),
    ("Prescription assessment", P.sub, P.FACIL, P.FTYPE, P.DISTR, P.PROV),
    ("CHW assessment", C.SUB, C.FAC, C.LEVEL, C.DIST, C.PROV),
]


def s(v):
    return "" if v is None else str(v).strip()


def name_tokens(n):
    return [t.upper() for t in s(n).replace("__", " ").replace("-", " ").split()]


def expected_class(n):
    """The level the NAME claims. The last recognised token wins, so
    'Gabiro HP__Nyamasheke' reads as HP rather than as the district."""
    hit = [NAME_CLASS[t] for t in name_tokens(n) if t in NAME_CLASS]
    return hit[-1] if hit else ""


def recorded_class(level):
    return LEVEL_CLASS.get(s(level).lower(), "")


def collect():
    """-> flags, label variants, every facility x tool pair."""
    flags, variants, allrows = [], defaultdict(Counter), []
    for tool, rows, fcol, lcol, dcol, pcol in TOOLS:
        seen = {}
        for r in rows:
            n = s(r[fcol])
            seen.setdefault(n, (s(r[lcol]), s(r[dcol]), s(r[pcol]), 0))
            lvl, dist, prov, cnt = seen[n]
            seen[n] = (lvl, dist, prov, cnt + 1)
        for n, (lvl, dist, prov, cnt) in sorted(seen.items()):
            exp, got = expected_class(n), recorded_class(lvl)
            allrows.append([tool, prov, dist, n, lvl, exp or "(no token)",
                            got or "(unrecognised level)",
                            "match" if exp and exp == got else
                            ("MISMATCH" if exp and got else "cannot check"),
                            cnt])
            if exp == "health post":
                variants[tool][lvl] += 1
            if not exp or not got:
                check = ("name carries no level token" if not exp
                         else "level value not recognised")
                flags.append([tool, prov, dist, n, lvl, exp or "-", got or "-",
                              "Cannot check", check, cnt])
                continue
            if exp == got:
                continue
            check = {"health centre": "1. Name ends CS/HC but the level is not "
                                      "a health centre",
                     "hospital": "2. Name is a hospital but the level is not "
                                 "any kind of hospital",
                     "health post": "3. Name is a health post but the level is "
                                    "neither SGHP nor a plain health post",
                     }.get(exp, f"Name says {exp}, level says {got}")
            flags.append([tool, prov, dist, n, lvl, exp, got, "MISMATCH",
                          check, cnt])
    return flags, variants, allrows


FLAGS, VARIANTS, ALLROWS = collect()


def sheet(wb, title, blocks, notes=None, widths=None, fills=None):
    ws = wb.create_sheet(title[:31])
    ws.sheet_view.showGridLines = False
    r = 1
    if notes:
        for n in notes:
            c = ws.cell(row=r, column=1, value=n)
            c.font = Font(size=9, italic=True, color="555555")
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
            ws.row_dimensions[r].height = 13 * (1 + len(n) // 140)
            r += 1
        r += 1
    for btitle, headers, rows in blocks:
        if btitle:
            ws.cell(row=r, column=1, value=btitle).font = TITLE_FONT
            r += 1
        for ci, h in enumerate(headers, start=1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font, c.fill, c.border = HDR_FONT, HDR_FILL, BORDER
            c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 30
        hdr_row = r
        r += 1
        for row in rows:
            for ci, v in enumerate(row, start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.border = BORDER
                c.alignment = Alignment(wrap_text=True, vertical="top")
                if fills:
                    f = fills(row, ci - 1, v)
                    if f:
                        c.fill = f
            r += 1
        ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)
        ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(len(headers))}{r - 1}"
        r += 2
    for ci, w in (widths or {}).items():
        ws.column_dimensions[get_column_letter(ci)].width = w


# ---------------------------------------------------------------
# the actionable flag list
# ---------------------------------------------------------------
# SGHP and a plain health post are different tiers in policy, so a facility
# recorded as one in one tool and the other in the next is a real
# disagreement, not a spelling difference.
TIERS = {"second generation health post (sghp)": "SGHP", "sghp": "SGHP",
         "health post (hp)": "HP", "hp": "HP (unspecified tier)"}


def flag_rows():
    """One row per thing somebody has to act on."""
    per_fac = defaultdict(dict)
    for tool, prov, dist, n, lvl, exp, got, verdict, cnt in ALLROWS:
        per_fac[n][tool] = (lvl, got, prov, dist, cnt)

    out, n_label_only = [], 0
    for n, per in sorted(per_fac.items()):
        lvls = {l for l, _, _, _, _ in per.values()}
        classes = {c for _, c, _, _, _ in per.values()}
        prov = next(p for _, _, p, _, _ in per.values())
        dist = next(d for _, _, _, d, _ in per.values())
        where = "; ".join(f"{t.split()[0]} = {l!r}" for t, (l, _, _, _, _)
                          in sorted(per.items()))

        # 1. name vs level, within a tool
        for tool, (lvl, got, _, _, cnt) in sorted(per.items()):
            exp = expected_class(n)
            if exp and got and exp != got:
                out.append(["F-NAME", "High",
                            "Name and recorded level disagree", tool, prov,
                            dist, n, lvl, exp, cnt,
                            "Every by-level table puts this facility in the "
                            "wrong tier.",
                            "Confirm the level with the collector, then "
                            "correct it in Kobo or add it to the analysis "
                            "overrides."])
        # 2. the tools disagree on the class
        if len(classes) > 1:
            out.append(["F-TOOL", "High",
                        "Tools disagree on the level CLASS", 
                        " vs ".join(sorted(t.split()[0] for t in per)), prov,
                        dist, n, where, expected_class(n) or "-",
                        sum(c for _, _, _, _, c in per.values()),
                        "The same facility is counted in two different tiers "
                        "depending on which tool you read.",
                        "Decide the correct level and align both tools."])
        # 3. the tools disagree on SGHP vs plain health post
        tiers = {TIERS.get(l.lower()) for l in lvls if l.lower() in TIERS}
        if len(tiers) > 1:
            out.append(["F-TIER", "Medium",
                        "Tools disagree on the health-post TIER (SGHP vs HP)",
                        " vs ".join(sorted(t.split()[0] for t in per)), prov,
                        dist, n, where, "health post",
                        sum(c for _, _, _, _, c in per.values()),
                        "SGHP and a plain health post are different tiers in "
                        "policy, with different expected medicine lists.",
                        "Confirm which tier this post is and align both tools."])
        # 4. the bare "HP" value, which names no tier at all
        for tool, (lvl, got, _, _, cnt) in sorted(per.items()):
            if lvl.strip().upper() == "HP":
                out.append(["F-BARE", "Medium",
                            "Level recorded as a bare 'HP' - tier unspecified",
                            tool, prov, dist, n, lvl, "health post", cnt,
                            "Splits health posts into a third tier in every "
                            "by-level table, alongside SGHP and 'Health Post "
                            "(HP)'.",
                            "Assign it to SGHP or to a plain health post, and "
                            "remove the bare option from the form's choice "
                            "list."])
        if len(lvls) > 1 and len(classes) == 1 and len(tiers) <= 1:
            n_label_only += 1
    return out, n_label_only


FLAG_ROWS, LABEL_ONLY = flag_rows()
FLAG_ROWS.sort(key=lambda r: ({"High": 0, "Medium": 1}.get(r[1], 2), r[6]))


def build_flags():
    """The standalone flag export - only what needs action."""
    wb = Workbook()
    wb.remove(wb.active)
    sev = Counter(r[1] for r in FLAG_ROWS)
    kinds = Counter(r[2] for r in FLAG_ROWS)

    summary = [[k, v, {"Name and recorded level disagree":
                       "A facility's name says one tier, the level answer says "
                       "another.",
                       "Tools disagree on the level CLASS":
                       "Two tools record different tiers for the same facility.",
                       "Tools disagree on the health-post TIER (SGHP vs HP)":
                       "Both agree it is a health post but not which kind.",
                       "Level recorded as a bare 'HP' - tier unspecified":
                       "The value names no tier, so it becomes a third "
                       "health-post category."}.get(k, "")]
               for k, v in kinds.most_common()]
    summary.append(["", "", ""])
    summary.append(["Label differences that are NOT flagged", LABEL_ONLY,
                    "The two forms spell the same level differently - "
                    "'Health Centre (HC)' against 'HC', 'MHC (health-centre "
                    "level)' against 'MHC'. Same tier, different choice list, "
                    "so no action beyond normalising before you group."])

    sheet(wb, "Summary",
          [("What was flagged", ["Issue", "Flags", "What it means"], summary),
           ("By severity", ["Severity", "Flags"],
            [[k, v] for k, v in sorted(sev.items())])],
          notes=[f"{len(FLAG_ROWS)} flag(s) across "
                 f"{len({r[6] for r in FLAG_ROWS})} facilities, from "
                 f"{len({r[3] for r in ALLROWS})} distinct facility names "
                 f"checked in three tools.",
                 "High = the facility sits in the wrong tier in at least one "
                 "tool, so by-level figures are wrong for it. Medium = the "
                 "tier is ambiguous rather than wrong."],
          widths={1: 52, 2: 10, 3: 96})

    sheet(wb, "Flags",
          [(None, ["Flag", "Severity", "Issue", "Tool(s)", "Province",
                   "District", "Facility as recorded",
                   "Level recorded", "Level the name implies", "Submissions",
                   "Why it matters", "Action"], FLAG_ROWS)],
          notes=["One row per thing to act on. Sorted by severity then "
                 "facility.",
                 "A flag does not say which side is wrong - the name may be "
                 "right and the level mis-selected, or the facility may have "
                 "been renamed. Each needs the field team to confirm."],
          widths={1: 9, 2: 10, 3: 46, 4: 26, 5: 13, 6: 14, 7: 28, 8: 56,
                  9: 20, 10: 12, 11: 54, 12: 56},
          fills=lambda row, ci, v: (BAD_FILL if ci == 1 and v == "High"
                                    else WARN_FILL if ci == 1 else None))

    # open on the Flags sheet, not on Summary - a reader who lands on a sheet
    # of counts reasonably concludes there are no mismatches to see
    wb.active = wb.sheetnames.index("Flags")
    if os.path.exists(OUT_FLAGS):
        os.chmod(OUT_FLAGS, 0o644)
    wb.save(OUT_FLAGS)
    print(f"wrote {OUT_FLAGS}")
    for k, v in kinds.most_common():
        print(f"    {v:>3}  {k}")
    print(f"    {LABEL_ONLY:>3}  label-only differences, not flagged")


def build():
    wb = Workbook()
    wb.remove(wb.active)
    mism = [f for f in FLAGS if f[7] == "MISMATCH"]
    nocheck = [f for f in FLAGS if f[7] == "Cannot check"]

    # ---------- 1. the flags
    def ffill(row, ci, v):
        if ci == 7:
            return BAD_FILL if v == "MISMATCH" else WARN_FILL
        return None

    sheet(wb, "Flags",
          [("Facilities whose name and recorded level disagree",
            ["Tool", "Province", "District", "Facility as recorded",
             "Level recorded", "Level the NAME implies",
             "Level class recorded", "Verdict", "Which check",
             "Submissions"], FLAGS or [["none", "", "", "", "", "", "",
                                        "all consistent", "", 0]])],
          notes=[f"{len(mism)} mismatch(es) and {len(nocheck)} uncheckable "
                 f"name(s) across the three tools, from "
                 f"{len({r[3] for r in ALLROWS})} distinct facility names.",
                 "Check 1: name ends CS or HC but the level is not a health "
                 "centre. Check 2: name is a hospital (DH, L2TH, PH, RH, "
                 "CHUB, CHUK or the word 'hospital') but the level is not any "
                 "kind of hospital. Check 3: name is a health post (HP or "
                 "SGHP) but the level is neither SGHP nor a plain health post.",
                 "A mismatch does not say which side is wrong - the name may "
                 "be right and the level mis-selected, or the facility may "
                 "have been renamed. Both need the field team to confirm."],
          widths={1: 22, 2: 13, 3: 14, 4: 30, 5: 30, 6: 22, 7: 22, 8: 12,
                  9: 52, 10: 12}, fills=ffill)

    # ---------- 2. the health-post labelling problem
    rows = []
    for tool, cnt in VARIANTS.items():
        for lvl, n in sorted(cnt.items(), key=lambda x: -x[1]):
            rows.append([tool, lvl, n,
                         "yes" if s(lvl).lower() in POST_LABELS else "NO",
                         len({r[3] for r in ALLROWS
                              if r[0] == tool and r[4] == lvl})])
    sheet(wb, "Health post labels",
          [("How the health-post level is recorded, for facilities whose name "
            "says HP or SGHP",
            ["Tool", "Level value as recorded", "Facilities", "Counts as a "
             "health post?", "Distinct facilities"], rows)],
          notes=["All of these pass check 3 - they ARE health-post levels. "
                 "The problem is that the same level is recorded under three "
                 "interchangeable labels, so health posts split across three "
                 "rows in every by-level table and look like three different "
                 "tiers.",
                 "Recommendation: collapse to one label in the form's choice "
                 "list - 'Second Generation Health Post (SGHP)' and 'Health "
                 "Post (HP)' are distinct tiers in policy, so decide which "
                 "facilities are which rather than leaving a bare 'HP' as a "
                 "third option."],
          widths={1: 22, 2: 40, 3: 12, 4: 20, 5: 18},
          fills=lambda row, ci, v: (WARN_FILL if ci == 1 and s(v).lower() == "hp"
                                    else None))

    # ---------- 3. the full crosstab, for reference
    g = defaultdict(Counter)
    for tool, prov, dist, n, lvl, exp, got, verdict, cnt in ALLROWS:
        g[(tool, exp)][lvl] += 1
    rows = []
    for (tool, exp), cnt in sorted(g.items()):
        for lvl, k in sorted(cnt.items(), key=lambda x: -x[1]):
            rows.append([tool, exp, lvl, k,
                         "match" if recorded_class(lvl) == exp else "MISMATCH"])
    sheet(wb, "Name vs level crosstab",
          [("Every combination of name-implied level and recorded level",
            ["Tool", "Level the name implies", "Level recorded", "Facilities",
             "Verdict"], rows)],
          notes=["Read this to see the naming conventions the data actually "
                 "uses, and to confirm the checks are matching the right "
                 "things before acting on a flag."],
          widths={1: 22, 2: 24, 3: 40, 4: 12, 5: 12},
          fills=lambda row, ci, v: (BAD_FILL if ci == 4 and v == "MISMATCH"
                                    else OK_FILL if ci == 4 else None))

    # ---------- 4. every facility, for lookup
    sheet(wb, "All facilities",
          [("One row per facility per tool",
            ["Tool", "Province", "District", "Facility", "Level recorded",
             "Level the name implies", "Level class recorded", "Verdict",
             "Submissions"], ALLROWS)],
          notes=["The full list behind the checks, filterable."],
          widths={1: 22, 2: 13, 3: 14, 4: 30, 5: 30, 6: 22, 7: 22, 8: 12,
                  9: 12},
          fills=lambda row, ci, v: (BAD_FILL if ci == 7 and v == "MISMATCH"
                                    else WARN_FILL if ci == 7 and
                                    v == "cannot check" else None))

    wb.active = wb.sheetnames.index("Flags")
    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)      # Box resets synced files to read-only
    wb.save(OUT)

    print(f"wrote {OUT}")
    print(f"  checked {len(ALLROWS)} facility x tool records, "
          f"{len({r[3] for r in ALLROWS})} distinct facility names")
    print(f"  MISMATCHES: {len(mism)}")
    for f in mism:
        print(f"    [{f[0]}] {f[3]} - level recorded '{f[4]}', name implies "
              f"{f[5]} ({f[8].split('.')[0]})")
    print(f"  uncheckable names: {len(nocheck)}")
    for tool, cnt in VARIANTS.items():
        if len(cnt) > 1:
            print(f"  {tool}: health-post level recorded under "
                  f"{len(cnt)} different labels - "
                  + ", ".join(f"{k!r} ({v})" for k, v in cnt.most_common()))


if __name__ == "__main__":
    build()
    build_flags()
