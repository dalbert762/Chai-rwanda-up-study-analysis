"""Collapse multi-strength medicines to one molecule per facility (Method 3).

    import molecules as MOL
    rows = MOL.collapse(M.MEDROWS)      # facility x MOLECULE observations
    rows = MOL.eml_base(M.MEDROWS)      # the study's availability denominator

Adopted 2026-10-05. Two strengths of the same molecule are therapeutically
interchangeable, so a facility holding Metformin 500mg can treat a patient who
needs metformin. Counting 500mg and 850mg as two tracer items understates what
the facility can do, and drags a molecule's score down whenever one strength
is rare - hydrochlorothiazide scored 64.7% as an average of its strengths
against 97.5% on the question that matters, because the 50mg form is stocked
by 25 facilities and the 25mg by 234.

THE RULE
For each facility and each molecule, across the strengths assessed there:
    applicable  if ANY strength is applicable at that level
    managed     if ANY strength is managed
    available   if ANY managed strength was on the shelf
The collapsed row keeps the facility, its parent submission and the programme
and drug class of the first contributing strength.

WHAT IS NOT COLLAPSED
Quantities, months of stock, stock-out days, requisitions and financial loss
stay at item level and are read from the original rows. Summing a quantity of
500mg tablets and 850mg tablets gives a number with no meaning, and
stock-out days for two products are not one product's stock-out days. Only
AVAILABILITY and the management/applicability flags collapse.

The groups are derived from the medicine names, never hard-coded, and
`report()` prints every group so a new strength appearing in a later refresh
is visible rather than silently merged.
"""
import re
from collections import defaultdict

import analyze_medicine_assessment as M

# key() is molkey.strength_key: the plain key plus the merge of the two
# thiamine products (100mg "THIAMINE (VIT B1)" and 50mg "Vitamin B1
# (Thiamine)"), agreed 2026-10-15 so Method 3 and cleaning rule S9 group
# strengths the same way.
from molkey import FORMS, strength_key as key   # noqa: F401 - re-exported


def s(v):
    return "" if v is None else str(v).strip()


def groups(rows):
    g = defaultdict(set)
    for r in rows:
        g[key(r[M.MED])].add(s(r[M.MED]))
    return {k: sorted(v) for k, v in g.items()}


def label(names):
    """Display name for a collapsed molecule."""
    if len(names) == 1:
        return names[0]
    stem = names[0].split()[0].strip(",/")
    return f"{stem.title()} (all {len(names)} strengths)"


def collapse(rows=None):
    """-> one synthetic row per facility x molecule, availability by Method 3.

    The returned rows carry the same column names the analyses already read,
    so they are a drop-in replacement for M.MEDROWS wherever the question is
    availability. `_strengths` and `_items` record what went into each.
    """
    src = rows if rows is not None else M.MEDROWS
    g = groups(src)
    buckets = defaultdict(list)
    for r in src:
        buckets[(s(r["_p"][M.FAC]), key(r[M.MED]))].append(r)

    out = []
    for (fac, mol), rs in buckets.items():
        names = g[mol]
        managed = [r for r in rs if M.yes(r[M.MANAGED])]
        answered = [r for r in managed if not M.blank(r[M.AVAIL])]
        any_app = any(M.yes(r[M.APPLICABLE]) for r in rs)
        any_avail = any(M.yes(r[M.AVAIL]) for r in answered)
        base = rs[0]
        row = dict(base)
        row[M.MED] = label(names)
        row[M.APPLICABLE] = "Yes" if any_app else "No"
        row[M.MANAGED] = "Yes" if managed else "No"
        row[M.AVAIL] = ("Yes" if any_avail else "No") if answered else None
        row["_mol"] = mol
        row["_strengths"] = len(names)
        row["_items"] = len(rs)
        row["_contributing"] = sorted({s(r[M.MED]) for r in rs})
        out.append(row)
    return out


def report(rows=None):
    """Lines describing what was merged, for a method note or a console run."""
    src = rows if rows is not None else M.MEDROWS
    g = groups(src)
    multi = {k: v for k, v in g.items() if len(v) > 1}
    lines = [f"{len({s(r[M.MED]) for r in src})} items -> {len(g)} molecules; "
             f"{len(multi)} assessed at more than one strength"]
    for k, v in sorted(multi.items()):
        lines.append(f"   {label(v)}: " + " + ".join(x[:34] for x in v))
    return lines


# ---------------------------------------------------------------
# THE EML DENOMINATOR (adopted 2026-10-05, second definition change)
# ---------------------------------------------------------------
def eml_base(rows=None):
    """Molecule rows on the EML-APPLICABLE denominator.

    The definition the study now uses:

        If a facility does not manage a medicine that it is supposed to
        manage - the medicine is on the EML and authorised at that
        facility's level of care - count that medicine as UNAVAILABLE.

    So the denominator is every molecule APPLICABLE at that level, not
    every molecule the facility chose to carry, and a patient who cannot
    be served because the facility never stocks an authorised medicine is
    counted the same as one who cannot be served because it ran out. The
    two reasons stay separable through `_gap`:

        _gap True   applicable, not managed  -> unavailable by definition
        _gap False  managed and answered     -> unavailable only if empty

    The rows are returned with MANAGED forced to "Yes" and AVAIL forced to
    "No" on the gap rows, so every existing availability filter
    (`yes(MANAGED) and not blank(AVAIL)`) selects exactly the applicable
    set without being rewritten. The facility's REAL management flag is
    preserved as `_managed`, and anything reporting how many medicines a
    facility manages must read that, never MANAGED.

    Rows applicable but with no availability answer and no management
    answer are still counted as unavailable - not managing is itself the
    answer under this definition.
    """
    out = []
    for r in collapse(rows):
        if not M.yes(r[M.APPLICABLE]):
            continue           # not authorised at this level - not expected
        row = dict(r)
        row["_managed"] = M.yes(r[M.MANAGED])
        row["_gap"] = not row["_managed"]
        if row["_gap"]:
            row[M.MANAGED] = "Yes"
            row[M.AVAIL] = "No"
        out.append(row)
    return out


ESSENTIAL = "Essential Medicines"


def two_track_base(rows=None):
    """Molecule rows on the TWO-TRACK denominator - the availability page's
    definition from 2026-10-13.

        Essential Medicines (medicine_program)  -> the EML track: every
            molecule the EML authorises at the facility's level, with
            authorised-but-not-managed counted UNAVAILABLE. Exactly
            eml_base().
        every other programme (Malaria, MCCH, FP, ARV, TB, Consumables)
            -> the managed track: only molecules the facility manages.
            Programme products are allocated by the programme - MFT zoning,
            faith-based FP, level-restricted MCCH - so a facility not
            holding one is usually not supposed to, and counting it
            unavailable would report the allocation as a shortage.

    Rows carry the same flags as eml_base() - `_managed`, `_gap` - plus
    `_track` ("EML" or "managed"). Managed rows with no availability answer
    are dropped on both tracks, so every row returned has a Yes/No and the
    denominator is simply the row count.
    """
    out = []
    for r in collapse(rows):
        managed = M.yes(r[M.MANAGED])
        if s(r[M.PROGRAM]) == ESSENTIAL:
            if not M.yes(r[M.APPLICABLE]):
                continue       # not authorised at this level - not expected
            row = dict(r)
            row["_track"] = "EML"
            row["_managed"] = managed
            row["_gap"] = not managed
            if row["_gap"]:
                row[M.MANAGED] = "Yes"
                row[M.AVAIL] = "No"
        else:
            if not managed:
                continue       # programme product not held - not expected
            row = dict(r)
            row["_track"] = "managed"
            row["_managed"] = True
            row["_gap"] = False
        if M.blank(row[M.AVAIL]):
            continue           # managed, but availability never answered
        out.append(row)
    return out


def managed_base(rows=None):
    """Molecule rows on the MANAGED-ONLY denominator: every molecule the
    facility manages and answered for, whatever its programme and whatever
    the EML says about its level. The comparison method on the availability
    page - it asks "is the shelf stocked where the facility runs the
    product?" and ignores what the facility is supposed to hold."""
    out = []
    for r in collapse(rows):
        if not M.yes(r[M.MANAGED]) or M.blank(r[M.AVAIL]):
            continue
        row = dict(r)
        row["_track"] = "managed"
        row["_managed"] = True
        row["_gap"] = False
        out.append(row)
    return out


def eml_note(rows=None):
    """One sentence stating the denominator, for a method note."""
    b = eml_base(rows)
    g = sum(1 for r in b if r["_gap"])
    av = sum(1 for r in b if M.yes(r[M.AVAIL]))
    return (f"Denominator: {len(b):,} facility x molecule pairs the EML "
            f"authorises at that level of care. {g:,} of them "
            f"({g / len(b):.1%}) are medicines the facility does not manage "
            f"at all; under the study definition those count as unavailable, "
            f"alongside the {len(b) - g - av:,} that were managed but empty "
            f"on the visit day.")
