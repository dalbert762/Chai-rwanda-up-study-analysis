"""Normalise the recorded name of the electronic recording system.

    import emr_systems as EMR
    EMR.canonical("OPEN MRS")        -> ("OpenMRS", "")
    EMR.canonical("elmis & Medisoft")-> ("eLMIS + Medisoft", "")
    EMR.canonical("1490")            -> ("", "numeric value removed")

The column is free text, and 162 distinct values were typed for what is
really seven systems. "OpenMRS" alone appears as OPENMRS, openmrs, OPEN MRS,
Open MRS, Openmrs, openrms, OMRS, openMRS, opemrs, OpenRMS, operms, OPEMRS,
OPNMRS, OPN MRS, OpnMRS, OPen MRS, open RMS, Open RMS, OpenERMs, OpenERMS and
Open eMR - so any count grouped on the raw value is meaningless.

THREE THINGS HAPPEN
1. A purely numeric entry is removed. Thirty-two rows hold values like 1490
   or 16820 - a quantity typed into the wrong box. A number is not the name
   of a system, and leaving it in makes it look like one.
2. Each name is mapped to one canonical spelling, matched on letters only so
   that case, spacing, hyphens and the common transpositions all collapse
   together.
3. Where a row names more than one system - "OPENMRS &eLMIS &SAP" - each is
   canonicalised and they are joined with " + " in a fixed order, so the same
   combination always reads the same way.

Any qualifier the collector added ("but store manager unavailable to check",
"only used for ordering") is kept in a separate note rather than being
silently dropped, because it says something about how far the system is
actually used.
"""
import re

# canonical name -> the letter-only forms seen in the data
SYSTEMS = {
    "eLMIS": ["elmis", "elms", "elimis", "elims", "elimis", "emlis", "elmls",
              "elimsis", "elmis", "elimis", "emls", "elmiss", "elismi",
              "elmsi", "elis", "elmes", "elimes"],
    "eBuzima": ["ebuzima", "eubuzima", "ebuzuma", "ebzima", "buzima",
                "ebuzim", "eubuzim"],
    "OpenMRS": ["openmrs", "omrs", "openrms", "opemrs", "operms", "opnmrs",
                "opneclinic0", "openemr", "openerms", "openrms", "opnmrs",
                "openmr", "opemrs", "openmsr", "opermrs"],
    "OpenClinic": ["openclinic", "openclinc", "opneclinic", "openclnic",
                   "opencinic", "openclinique"],
    "Medisoft": ["medisoft", "medsoft", "medfoft", "medosoft", "medsof"],
    "IMARA": ["imara", "imra", "imana", "emara", "imaraemr",
              "imaraelectronicmedicalrecord"],
    "SAP": ["sap"],
    "CARE": ["care"],
    "EHRIS": ["ehris"],
    "e-RMS": ["erms"],
    "e-fiche": ["efiche"],
    # Not a named EMR, but it is a system and it is how those facilities
    # record - folding it into "no system" would overstate the gap.
    "In-house spreadsheet": ["inhouseexcelspreadsheet", "excel",
                             "excelspreadsheet", "inhouseexcel"],
}
LOOKUP = {form: name for name, forms in SYSTEMS.items() for form in forms}

# Reporting order, so a combination always reads the same way.
ORDER = ["eLMIS", "eBuzima", "OpenMRS", "OpenClinic", "Medisoft", "IMARA",
         "EHRIS", "SAP", "CARE", "e-RMS", "e-fiche", "In-house spreadsheet"]

# Answers that say there is NO system. Kept distinct from a blank, which
# means the question was not reached.
NONE_WORDS = ("none", "nosystem", "na", "notapplicable", "nil", "no",
              "register", "registers", "paper")

SPLIT = re.compile(r"\s*(?:&|\+|,|/| and | na )\s*", re.I)


def letters(t):
    return re.sub(r"[^a-z]", "", str(t).lower())


def canonical(raw):
    """-> (normalised name, note). Empty name means no system recorded."""
    t = str(raw or "").strip()
    if not t:
        return "", ""
    if re.fullmatch(r"[\d.,\s]+", t):
        return "", "numeric value removed"

    found, leftover = [], []
    for part in SPLIT.split(t):
        part = part.strip(" .()")
        if not part:
            continue
        key = letters(part)
        if not key:
            continue
        if key in LOOKUP:
            found.append(LOOKUP[key])
            continue
        # a token that merely CONTAINS a known system, e.g.
        # "ELMIS BUT STORE MANAGER UNAVAILABLE TO CHECK SYS"
        hit = next((n for form, n in LOOKUP.items()
                    if len(form) > 3 and form in key), None)
        if hit:
            found.append(hit)
            rest = re.sub(r"[^a-z ]", " ", part.lower())
            rest = " ".join(w for w in rest.split()
                            if letters(w) not in LOOKUP)
            if rest:
                leftover.append(rest.strip())
        elif key in NONE_WORDS or any(key.startswith(w) for w in NONE_WORDS):
            leftover.append(part)
        else:
            leftover.append(part)

    name = " + ".join(sorted(set(found), key=lambda x: ORDER.index(x)
                             if x in ORDER else 99))
    note = "; ".join(dict.fromkeys(leftover))
    if not name and note and any(letters(note).startswith(w)
                                 for w in NONE_WORDS):
        return "", f"no system reported: {note}"
    return name, note
