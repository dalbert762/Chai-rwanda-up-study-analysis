"""Medicine name -> molecule key, with no dependency on the analysis.

    import molkey
    molkey.key("Metformin 850mg")      -> "metformin"

Kept apart from molecules.py because analyze_medicine_assessment needs it
too (cleaning rule S9 groups strengths), and molecules.py imports the
analysis - one importing the other back would be circular.
"""
import re

FORMS = (r"tab|tabs|tablet|tablets|caps|capsule|capsules|inj|injection|"
         r"injectable|powder|for|disp|dispersible|susp|suspension|syrup|"
         r"solution|sol|oint|ointment|cream|supp|suppository|drops|film|"
         r"coated|granules|iv|im|sc|po|bottle|bottleof|vial|amp|ampoule|"
         r"pr|dec|of|with|solvent")


def key(name):
    """Medicine name -> molecule key.

    Brackets are stripped BEFORE the parenthetical rule because one DHAP
    strength is written 'DHAP)' with no opening bracket and would otherwise
    key differently from its own siblings.
    """
    t = str(name).lower().replace("(", " ").replace(")", " ")
    t = re.sub(r"\b\d+[\d.,/]*\s*(mg|g|gr|mcg|µg|ui|iu|ml|l|%|mgs)\b", " ", t)
    t = re.sub(r"\b\d+\s*x\s*\d+\b", " ", t)
    t = re.sub(r"\bb\s*/?\s*\d+\b", " ", t)
    t = re.sub(r"\d+[\d.,/]*", " ", t)
    t = re.sub(r"\b(" + FORMS + r")\b", " ", t)
    t = re.sub(r"[^a-z ]", " ", t)
    return " ".join(t.split())


# The same molecule written two ways in the tool's own medicine list, so the
# plain key splits it. Used where strengths must be compared (cleaning rule
# S9, the strength-mismatch extract) and, since 2026-10-15, for Method 3 -
# molecules.key is strength_key.
SAME_MOLECULE = {"thiamine vit": "thiamine", "vitamin thiamine": "thiamine"}


def strength_key(name):
    k = key(name)
    return SAME_MOLECULE.get(k, k)
