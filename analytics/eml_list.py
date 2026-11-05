"""Match the assessment's medicine names to the official EML list.

    import eml_list as EML
    by_ing, names = EML.load_eml()
    EML.applicable("Metformin 850mg", "Health Centre (HC)", by_ing, names)

The form asks "Is <medicine> applicable to this facility level as per EML?"
and the collector answers from memory. The EML list itself is the authority,
so this module reads it and says what the list actually holds.

MATCHING IS BY INGREDIENT AND STRENGTH, NOT BY NAME
The two sources name products completely differently - the assessment writes
"HYDROCHLOROTHIAZIDE 50MG TAB", the list writes
"Hydrochlorothiazide - Tab. 25 mg / 50mg" - so only 2 of 80 names match on
text. Each name is reduced to its ingredient words and its numeric strengths,
and matched on those. A match on ingredient AND strength is reported
separately from one on ingredient alone, because an ingredient-only match
pools the levels of every form and strength the list holds for that molecule
and can therefore be too generous.

WHAT IS NOT GUESSED
Fixed-dose combinations, commodities and brand names are left unmatched
rather than mapped by eye. Deciding that "RIFA 150MG+ISO 75MG+PYR 150MG+ETH
275MG TAB B/672" is a particular EML line is a clinical judgement, and a
wrong one silently moves rows into the availability denominator.
"""
import csv
import os
import re
from collections import defaultdict

import _bootstrap as B

EML_CSV = os.path.join(B.ROOT, "Data cleaning", "Final EML list.csv")

# Our recorded level  ->  the EML list's level code.
# FGHP is the first-generation health post. The form does not distinguish it
# from SGHP, and the agreed grouping already reports both as "Health post",
# so a bare "HP" is allowed to match either.
LEVEL_TO_EML = {
    "health centre (hc)": {"HC"},
    "mhc (health-centre level)": {"MHC", "HC"},
    "second generation health post (sghp)": {"SGHP"},
    "health post (hp)": {"SGHP", "FGHP"},
    "hp": {"SGHP", "FGHP"},
    "district hospital": {"DH"},
    "provincial hospital": {"L2TH/PH"},
    "level 2 teaching hospital": {"L2TH/PH"},
    "national referral / teaching hospital": {"UTH/RH"},
}

UNIT = r"(?:mg|g|gr|mcg|µg|ug|ui|iu|ml|l|%)"


def strengths(t):
    """The set of numeric strengths in a name, as plain numbers."""
    out = set()
    for m in re.finditer(r"(\d+(?:[.,]\d+)?)\s*" + UNIT, t, re.I):
        out.add(m.group(1).replace(",", "."))
    return out


STOP = {"tab", "tabs", "tablet", "tablets", "cap", "caps", "capsule",
        "capsules", "inj", "injection", "injectable", "powder", "for", "susp",
        "suspension", "syrup", "solution", "sol", "oint", "ointment", "cream",
        "supp", "drops", "drop", "film", "coated", "granules", "iv", "im",
        "sc", "po", "vial", "amp", "ampoule", "dispersible", "disp", "with",
        "solvent", "infusion", "infu", "and", "the", "of", "eye", "all",
        "strengths", "bottle", "sachet", "rod", "implant", "inserter",
        "cycle", "tests", "test", "dose", "per", "sterile", "isobare"}


# Spellings our form uses that the EML list does not. Only obvious
# single-ingredient slips and abbreviations are listed - a combination product
# is never guessed at, because mapping "RIFA 150+ISO 75+PYR 150+ETH 275" onto
# an EML line by eye is a clinical judgement, not a text-matching one.
ALIAS = {
    "cirpofloxacin": "ciprofloxacin",      # transposed letters on the form
    "carbamezepine": "carbamazepine",      # e for a
    "chlorpherniramine": "chlorpheniramine",
    "desloratadine": "desloratadine",
    "imipramine": "imipramine",
    "amlodipin": "amlodipine",
    "losart": "losartan",
    "htz": "hydrochlorothiazide",
    "clavul": "clavulanic",
    "acid": "acid",
    "arte": "artemether",
    "lume": "lumefantrine",
    "iso": "isoniazid",
    "rifa": "rifampicin",
    "rifamp": "rifampicin",
    "rifampi": "rifampicin",
    "isoniaz": "isoniazid",
    "pyr": "pyrazinamide",
    "pyrazinamid": "pyrazinamide",
    "eth": "ethambutol",
    "mgso": "magnesium",
    "hum": "human",
}


def ingredients(t):
    """Normalised ingredient words, strengths and dosage words removed."""
    t = re.sub(r"\*+", " ", t)
    t = t.split(" - ")[0] if " - " in t else t
    t = re.sub(r"\(.*?\)", " ", t)
    t = re.sub(r"\d+(?:[.,]\d+)?\s*" + UNIT, " ", t, flags=re.I)
    t = re.sub(r"[^A-Za-z]+", " ", t).lower()
    words = [ALIAS.get(w, w) for w in t.split()
             if w not in STOP and len(w) > 2]
    return tuple(sorted(set(words)))


def load_eml():
    """-> {(ingredients, strength): {levels}}, plus a name lookup."""
    by_ing = defaultdict(set)
    names = defaultdict(set)
    with open(EML_CSV, encoding="utf-8-sig", errors="replace") as fh:
        for row in list(csv.reader(fh))[1:]:
            if len(row) < 3 or not row[1].strip():
                continue
            prod, lvl = row[1].strip(), row[2].strip()
            if not lvl:
                continue
            ing = ingredients(prod)
            if not ing:
                continue
            by_ing[ing].add((lvl, prod))
            names[ing].add(prod)
    return by_ing, names


def levels_for(med, by_ing, names):
    """Which EML levels list this medicine. -> (levels, matched products, how)."""
    ing = ingredients(med)
    st = strengths(med)
    if not ing or ing not in by_ing:
        return set(), set(), "no ingredient match"
    entries = by_ing[ing]
    if st:
        exact = {(l, p) for l, p in entries if strengths(p) & st}
        if exact:
            return ({l for l, _ in exact}, {p for _, p in exact},
                    "ingredient + strength")
    return ({l for l, _ in entries}, {p for _, p in entries},
            "ingredient only")


def applicable(med, recorded_level, by_ing, names):
    """-> (verdict, how, eml_levels, products)

    verdict is True / False / None, None meaning the question cannot be
    answered - the medicine did not match the list, or the recorded level has
    no EML equivalent. A starred level code (SGHP*, HC*) is treated as listed;
    the star marks a restriction in the EML's own notes, not an exclusion.
    """
    levels, prods, how = levels_for(med, by_ing, names)
    want = LEVEL_TO_EML.get(str(recorded_level).strip().lower())
    if not levels:
        return None, "no name match", set(), set()
    if not want:
        return None, "level not in the EML scheme", levels, prods
    plain = {x.rstrip("*") for x in levels}
    return bool(plain & want), how, levels, prods
