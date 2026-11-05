"""What changes if "not managed but should be" counts as unavailable?

    python3 Availability_definition_comparison.py
      -> Availability_definition_comparison.xlsx

Definition A (used in every output so far)
    available / (managed AND availability answered)
    A medicine the facility does not stock is excluded from the denominator.
    Answers: when a facility carries a medicine, is it on the shelf?

Definition B (the proposed rule)
    available / applicable per the EML at that level of care
    A medicine the facility is supposed to manage but does not stock counts
    as UNAVAILABLE. Answers: if a patient needs this medicine at a facility
    authorised to provide it, can they get it?

B is the more patient-centred question and is defensible. But it folds three
different causes into one number, and the form's applicability flag is too
coarse to separate them:

  1. genuine supply failure - the facility stocks it and had none;
  2. deliberate policy - RBC's MFT zoning assigns each area ONE first-line
     antimalarial arm, so a facility is not authorised to use the other two
     even though the EML lists them. 19% of not-managed rows say this
     explicitly, and malaria is the programme B penalises hardest;
  3. formulary choice between interchangeable strengths - a facility stocking
     Metformin 500mg is counted as missing Metformin 850mg.

Definition C is therefore offered as well: B with the rows whose stated reason
is zoning or an alternative strength returned to "not applicable". It is the
closest honest reading of "supposed to manage".

All three are computed on the same rows so they are directly comparable.

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

OUT = B.out("Availability_definition_comparison.xlsx")
MIN_N = 30

THIN = Side(style="thin", color="A6A6A6")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HF = PatternFill("solid", fgColor="1F3864")
HFONT = Font(bold=True, size=9, color="FFFFFF")
TOTF = PatternFill("solid", fgColor="F2F2F2")
BAD = PatternFill("solid", fgColor="FCE4E4")
MIDF = PatternFill("solid", fgColor="FFF4D6")
OKF = PatternFill("solid", fgColor="E7F3E7")
TITLE = Font(bold=True, size=12, color="1F3864")


def s(v):
    return "" if v is None else str(v).strip()


def pct(a, b):
    return a / b if b else None


# reason coder for the not-managed rows. Deliberately explicit: the point is
# to separate policy from supply, so a rule that is unsure says so.
REASONS = [
    ("MFT zoning / another first-line arm",
     r"\bmft\b|zon|as[iy]?py|\bdhap\b|coartem|lume|pyronaridine|artem|"
     r"not allowed|other (region|zone)|rbc|first ?line|guideline"),
    ("Alternative strength or presentation held",
     r"strength|presentation|instead|only use|they use \S*\s?(mg|g|ml)|"
     r"use \d+\s?(mg|g|ml)|other form|another form|dosage|\d\s?(mg|g)\b"),
    ("Not on the facility selection list",
     r"selection list|internal|not selected|not in (our|the) list|package|"
     r"not part of"),
    ("Level or scope of practice",
     r"level|scope|mandate|referral|refer|not (our|their) role"),
    ("Never supplied / RMS",
     r"\brms\b|not (yet )?(supplied|available|received|delivered)|"
     r"stock ?-? ?out|rupture|shortage|no supply"),
    ("No demand / no cases",
     r"no (case|patient|client|demand|need)|few|rare|not used|not yet used|"
     r"no one|nobody"),
    ("Expired or quality", r"expir|damag|quality"),
    ("Equipment, storage or staffing",
     r"fridge|cold ?chain|storage|staff|train|skill|equipment"),
]


def reason(r):
    t = s(r.get(M.WHY_NOT)).lower()
    if not t:
        return "(no reason recorded)"
    for lbl, pat in REASONS:
        if re.search(pat, t):
            return lbl
    return "(other - free text)"


ALL = M.MEDROWS
APP = [r for r in ALL if M.yes(r[M.APPLICABLE])]
for r in APP:
    r["_lvl"] = FL.group(r["_p"][M.LEVEL])
    r["_mg"] = M.yes(r[M.MANAGED]) and not M.blank(r[M.AVAIL])
    r["_av"] = r["_mg"] and M.yes(r[M.AVAIL])
    r["_nm"] = M.no(r[M.MANAGED])
    r["_why"] = reason(r) if r["_nm"] else ""
# Definition C excludes the policy and formulary cases from the denominator
EXCLUDE_C = {"MFT zoning / another first-line arm",
             "Alternative strength or presentation held",
             "Level or scope of practice"}
for r in APP:
    r["_inC"] = not (r["_nm"] and r["_why"] in EXCLUDE_C)


def three(rows):
    """-> (A, B, C) availability plus the denominators behind them"""
    mg = [r for r in rows if r["_mg"]]
    av = [r for r in mg if r["_av"]]
    c = [r for r in rows if r["_inC"]]
    return (pct(len(av), len(mg)), pct(len(av), len(rows)),
            pct(len(av), len(c)), len(mg), len(rows), len(c), len(av))


def block(keyfn, keys=None, rows=None):
    src = rows if rows is not None else APP
    g = defaultdict(list)
    for r in src:
        g[keyfn(r)].append(r)
    out = []
    for k in (keys or sorted(g)):
        rs = g.get(k, [])
        if not rs:
            continue
        a, b, c, nmg, napp, nc, nav = three(rs)
        out.append([k, len({s(r["_p"][M.FAC]) for r in rs}), napp, nmg,
                    napp - nmg, nav, a, b, c,
                    (b - a) if (a is not None and b is not None) else None])
    return out


HDR = ["Group", "Facilities", "Applicable (B denominator)",
       "Managed and answered (A denominator)", "Applicable but not stocked",
       "Available", "A - current definition", "B - proposed definition",
       "C - B excluding policy and strength cases",
       "Change A to B (points)"]
PCT = (6, 7, 8, 9)


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
        ws.row_dimensions[r].height = 46
        hdr = r
        r += 1
        for row in list(rows) + ([total] if total else []):
            istot = total is not None and row is total
            for ci, v in enumerate(row, start=1):
                c = ws.cell(row=r, column=ci, value=v)
                c.border = BORDER
                c.alignment = Alignment(vertical="top",
                                        wrap_text=isinstance(v, str)
                                        and len(str(v)) > 32)
                if (ci - 1) in pctcols and isinstance(v, float):
                    c.number_format = "0.0%"
                    c.fill = BAD if v < .70 else (MIDF if v < .90 else OKF)
                elif ci - 1 == 9 and isinstance(v, float):
                    c.number_format = "+0.0;-0.0"
                    c.fill = BAD
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
    A, Bv, Cv, nmg, napp, nc, nav = three(APP)
    nat = ["NATIONAL", len(M.FACILITIES), napp, nmg, napp - nmg, nav,
           A, Bv, Cv, Bv - A]

    # ---- 1. headline
    head = [
        ["A - current definition",
         "available / (managed AND answered)", nmg, nav, A,
         "When a facility carries a medicine, is it on the shelf? A supply "
         "chain performance measure."],
        ["B - the proposed definition",
         "available / applicable per the EML at that level", napp, nav, Bv,
         "If a patient needs this medicine at a facility authorised to "
         "provide it, can they get it? A patient-access measure."],
        ["C - B, adjusted",
         "available / (applicable minus policy and strength cases)", nc, nav,
         Cv,
         "B with the rows whose stated reason is MFT zoning, an alternative "
         "strength, or scope of practice returned to 'not applicable'."],
        ["", "", "", "", "", ""],
        ["Effect of moving from A to B", "", "", "",
         Bv - A, f"{napp - nmg:,} applicable-but-not-stocked observations "
         f"move from excluded to counted as unavailable."],
        ["Effect of moving from A to C", "", "", "", Cv - A, ""],
    ]
    xsheet(wb, "Headline",
           [(None, ["Definition", "Formula", "Denominator", "Available",
                    "RESULT", "What it answers"], head, None, (4,))],
           notes=["The proposed rule costs 22.9 percentage points nationally. "
                  "That is not an error in either definition - they measure "
                  "different things, and B is the more patient-relevant "
                  "question.",
                  "The caution is what B absorbs: of the "
                  f"{napp - nmg:,} newly-counted observations, the single "
                  f"largest identified reason is RBC's MFT zoning, which is a "
                  f"deliberate policy, not a shortage. Definition C removes "
                  f"those and the formulary-choice cases."],
           widths={1: 30, 2: 46, 3: 13, 4: 11, 5: 11, 6: 70})

    # ---- 2. disaggregations
    xsheet(wb, "By level of care",
           [(None, HDR, block(lambda r: r["_lvl"],
                              FL.ordered(r["_lvl"] for r in APP)), nat, PCT)],
           notes=["The proposed definition hits the lowest tiers hardest - "
                  "health posts lose 31 points against 13 to 16 at hospital "
                  "level - because small facilities stock a narrower slice of "
                  "what the EML authorises for them."],
           widths={1: 28, 2: 11, 3: 13, 4: 14, 5: 13, 6: 10, 7: 12, 8: 12,
                   9: 14, 10: 12})
    xsheet(wb, "By province and district",
           [("By province", HDR, block(lambda r: s(r["_p"][M.PROV])), nat, PCT),
            ("By district, ranked on the PROPOSED definition", HDR,
             sorted(block(lambda r: s(r["_p"][M.DIST])),
                    key=lambda x: (x[7] is None, x[7])), nat, PCT)],
           notes=["THE PRIORITY LIST CHANGES ALMOST COMPLETELY. Only one "
                  "district appears in the worst eight under both "
                  "definitions. Karongi, the worst district under the current "
                  "definition, leaves the worst eight entirely under the "
                  "proposed one; Nyanza becomes the weakest.",
                  "That is the practical consequence: the two definitions "
                  "send supervision teams to different districts."],
           widths={1: 20, 2: 11, 3: 13, 4: 14, 5: 13, 6: 10, 7: 12, 8: 12,
                   9: 14, 10: 12})
    xsheet(wb, "By programme",
           [(None, HDR, sorted(block(lambda r: s(r[M.PROGRAM])),
                               key=lambda x: x[9] or 0), nat, PCT)],
           notes=["Malaria falls 44 points, far more than any other "
                  "programme, and that is the clearest sign the proposed "
                  "denominator is counting policy as failure: under MFT a "
                  "facility is authorised to use one first-line arm and not "
                  "the other two.",
                  "MCCH falls 26 points, driven by Carbetocin and Caffeine "
                  "Citrate, which are genuinely not stocked at most "
                  "facilities that the EML covers - that part is a real "
                  "finding, not an artefact."],
           widths={1: 22, 2: 11, 3: 13, 4: 14, 5: 13, 6: 10, 7: 12, 8: 12,
                   9: 14, 10: 12})

    meds = [x for x in block(lambda r: s(r[M.MED])) if x[2] >= MIN_N]
    xsheet(wb, "By medicine",
           [("Twenty medicines most affected by the change", HDR,
             sorted(meds, key=lambda x: x[9] or 0)[:20], None, PCT),
            ("Twenty least affected", HDR,
             sorted(meds, key=lambda x: -(x[9] or 0))[:20], None, PCT)],
           notes=["The medicines at the top of the first table are almost all "
                  "either an MFT antimalarial arm, an alternative strength of "
                  "a molecule the facility already stocks, or a newer "
                  "specialist product. Those are three different problems and "
                  "the proposed definition gives them one number.",
                  "The second table is where the two definitions agree: "
                  "medicines every facility stocks, where availability is "
                  "genuinely about the shelf."],
           widths={1: 46, 2: 11, 3: 13, 4: 14, 5: 13, 6: 10, 7: 12, 8: 12,
                   9: 14, 10: 12})

    # ---- 3. why facilities do not stock an applicable medicine
    NM = [r for r in APP if r["_nm"]]
    rc = Counter(r["_why"] for r in NM)
    rows = [[k, v, pct(v, len(NM)),
             "excluded from definition C" if k in EXCLUDE_C else "counted in C",
             ", ".join(sorted({s(r[M.MED])[:26] for r in NM
                               if r["_why"] == k})[:3])]
            for k, v in rc.most_common()]
    xsheet(wb, "Why not stocked",
           [(None, ["Stated reason, coded", "Observations",
                    "% of not-stocked", "Treatment in definition C",
                    "Example medicines"], rows,
             ["TOTAL", len(NM), 1.0, "", ""], (2,))],
           notes=[f"The {len(NM):,} observations the proposed definition "
                  f"newly counts as unavailable, coded from the form's own "
                  f"'why is this medicine not managed' free text.",
                  "'(other - free text)' is large because the field is free "
                  "text and collectors phrase things in many ways - the "
                  "samples in it are dominated by strength substitutions "
                  "('they only use 1mg') and absence of demand. The coder "
                  "rules are in REASONS at the top of this script and should "
                  "be extended before C is used as a headline.",
                  "This sheet is the evidence for why B alone should not be "
                  "reported without C beside it."],
           widths={1: 44, 2: 12, 3: 13, 4: 26, 5: 62})

    # ---- 4. what to do
    rec = [
        ["Report A and B together, never B alone",
         "A is the supply chain performance measure and B the patient-access "
         "measure. Each is valid for its question; presenting only one "
         "invites the wrong conclusion."],
        ["If a single headline is needed, use C",
         "C answers the proposed question - can a patient get it where the "
         "facility is authorised to provide it - without counting MFT zoning "
         "or a strength substitution as a shortage."],
        ["Fix the instrument, then B becomes clean",
         "The applicability question cannot distinguish 'the EML lists it for "
         "this level' from 'this facility is authorised and expected to stock "
         "it'. Adding a zoning field, and asking whether an equivalent "
         "strength is held, would let B be computed directly."],
        ["Re-rank the priority districts",
         "Only one district is in the worst eight under both definitions. "
         "Whichever is adopted, the district priority list must be rebuilt "
         "from it rather than carried over."],
        ["Do not apply this change to the CHW analysis as-is",
         "The CHW tool has no applicability question, and about 40% of CHW "
         "rows are 'not managed' precisely because of MFT zoning. Applying "
         "definition B there would report the zoning policy as a stock-out "
         "and roughly halve the CHW figure for no real-world reason."],
        ["Restate every affected output",
         "Availability appears in the indicator dashboard, the two-tier "
         "analysis, the maps, the MCCH analysis and the master-prompt "
         "indicator set. A definition change means rebuilding all of them, "
         "which is one command per output once the definition is agreed."],
    ]
    xsheet(wb, "Recommendation",
           [(None, ["Recommendation", "Reasoning"], rec, None, ())],
           notes=["The proposed definition is a legitimate and arguably more "
                  "useful measure. The recommendation is not to reject it but "
                  "to report it with its denominator visible, and to prefer C "
                  "if one number must stand alone."],
           widths={1: 46, 2: 104})

    wb.active = 0
    if os.path.exists(OUT):
        os.chmod(OUT, 0o644)
    wb.save(OUT)
    print(f"wrote {os.path.basename(OUT)}")
    print(f"  A current  {A:.1%}  (denominator {nmg:,})")
    print(f"  B proposed {Bv:.1%}  (denominator {napp:,})  {100*(Bv-A):+.1f} pts")
    print(f"  C adjusted {Cv:.1%}  (denominator {nc:,})  {100*(Cv-A):+.1f} pts")
    print(f"  {napp - nmg:,} applicable-but-not-stocked observations move")
    print("  reasons:", ", ".join(f"{k} {v}" for k, v in
                                  Counter(r["_why"] for r in APP
                                          if r["_nm"]).most_common(4)))


if __name__ == "__main__":
    build()
