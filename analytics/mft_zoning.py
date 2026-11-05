"""MFT zoning: the one first-line antimalarial allocated to each district
(AL = artemether-lumefantrine, ASPY = artesunate-pyronaridine, DHAP =
dihydroartemisinin-piperaquine). Used by cleaning rule S10 in
analyze_medicine_assessment.py and analyze_medicine_assessment_chw.py, and by
analytics/Antimalarial_zoning_flags.py.

OFFICIAL: from the user's zoning list (2026-11-02). Only these nine districts
were legible. INFERRED: the antimalarial most of the district's assessed
facilities say they manage - replace with the official list when received.
A district allowed two products takes both, e.g. "AL;ASPY".
"""

OFFICIAL = {
    "Gasabo": "ASPY",
    "Gatsibo": "AL",
    "Kamonyi": "DHAP",
    "Kayonza": "AL",
    "Kicukiro": "ASPY",
    "Muhanga": "DHAP",
    "Nyagatare": "AL",
    "Nyarugenge": "ASPY",
    "Ruhango": "DHAP",
}

INFERRED = {
    "Bugesera": "DHAP",
    "Burera": "ASPY",
    "Gakenke": "ASPY",
    "Gicumbi": "DHAP",
    "Gisagara": "AL",
    "Huye": "AL",
    "Karongi": "ASPY",
    "Kirehe": "DHAP",
    "Musanze": "ASPY",
    "Ngoma": "DHAP",
    "Ngororero": "ASPY",
    "Nyabihu": "ASPY",
    "Nyamagabe": "AL",
    "Nyamasheke": "ASPY",
    "Nyanza": "AL",
    "Nyaruguru": "AL",
    "Rubavu": "ASPY",
    "Rulindo": "ASPY",
    "Rusizi": "ASPY",
    "Rutsiro": "ASPY",
    "Rwamagana": "DHAP",
}

ZONE = dict(INFERRED, **OFFICIAL)


def group(name):
    """AL / ASPY / DHAP for a first-line antimalarial product, else None.
    Takes the stock-tool name or the CHW short name."""
    n = str(name).upper()
    if ("ARTE" in n and "LUME" in n) or n.startswith("ALU "):
        return "AL"
    if "PYRONARIDINE" in n or n.startswith("PYR/ART"):
        return "ASPY"
    if "DIHYDROARTEMISININ" in n or n.startswith("DHAP"):
        return "DHAP"
    return None


def out_of_zone(product, district):
    g = group(product)
    z = ZONE.get(str(district).strip())
    return bool(g and z and g not in z.split(";"))
