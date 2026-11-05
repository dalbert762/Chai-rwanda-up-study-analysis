"""RMS hubs: which Rwanda Medical Supply hub serves each district.

    import rms_hubs as HUB
    HUB.HUB_OF["Rubavu"]   -> "Musanze HUB"
    HUB.ORDER              -> the hubs, in the order the user gave them

Supplied by the user on 2026-10-20. Every dashboard page has an "RMS hub"
filter that uses this lookup, keyed on the facility's district.
"""

HUBS = {
    "Musanze HUB": ["Rubavu", "Gakenke", "Nyabihu", "Burera", "Musanze"],
    "Kayonza HUB": ["Nyagatare", "Rwamagana", "Kirehe", "Gatsibo", "Ngoma",
                    "Kayonza"],
    "Kigali HUB": ["Kicukiro", "Nyarugenge", "Gasabo", "Gicumbi", "Rulindo",
                   "Bugesera", "Kamonyi"],
    "Karongi HUB": ["Rutsiro", "Ngororero", "Nyamasheke", "Rusizi", "Karongi"],
    "Huye HUB": ["Gisagara", "Nyaruguru", "Nyamagabe", "Nyanza", "Ruhango",
                 "Muhanga", "Huye"],
}
ORDER = list(HUBS)
HUB_OF = {d: h for h, ds in HUBS.items() for d in ds}


def hub(district):
    """District name -> its RMS hub ("" if the district is not listed)."""
    return HUB_OF.get("" if district is None else str(district).strip(), "")
