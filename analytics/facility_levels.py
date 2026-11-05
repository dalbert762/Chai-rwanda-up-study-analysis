"""One definition of the facility-level tiers, shared by every analysis.

    import facility_levels as FL
    FL.group(row["Health facility level"])   -> "Health post" | "Health centre" | ...
    FL.ORDER                                  -> the tiers, lowest first

The forms record more tiers than are useful to report, and record some of them
under more than one label. Agreed grouping (2026-10-02):

  Health post    = Second Generation Health Post (SGHP), Health Post (HP), and
                   the bare "HP" value. SGHP and a plain health post are
                   distinct in policy but were splitting one small tier into
                   three thin rows, so they are reported together.
  Health centre  = Health Centre (HC) AND the maternity health centre (MHC).
                   An MHC is a health-centre-level facility - the form's own
                   value says "MHC (health-centre level)" - so it belongs with
                   the health centres rather than as a tier of its own.
  District hospital = district AND provincial hospitals (user decision
                   2026-10-19): a provincial hospital is reported with the
                   district hospitals in every output.
  Other hospitals = level 2 teaching and national referral / teaching are
                   reported separately.

Both forms' spellings map to the same group, so the medicine assessment's
"Health Centre (HC)" and the prescription assessment's "HC" no longer appear
as two tiers.

The ungrouped values are still in the data; nothing is overwritten. Any output
that groups should also show what was merged - see the "Recorded level values"
sheet in Product_availability_by_geography.xlsx.
"""

HEALTH_POST = "Health post"
HEALTH_CENTRE = "Health centre"
DISTRICT_HOSPITAL = "District hospital"
PROVINCIAL_HOSPITAL = "Provincial hospital"
L2_TEACHING = "Level 2 teaching hospital"
NATIONAL_REFERRAL = "National referral hospital"

# every spelling either form uses, lowercased
GROUP = {
    # health post - all three labels, both forms
    "second generation health post (sghp)": HEALTH_POST,
    "sghp": HEALTH_POST,
    "health post (hp)": HEALTH_POST,
    "hp": HEALTH_POST,
    # health centre, with the maternity health centre folded in
    "health centre (hc)": HEALTH_CENTRE,
    "hc": HEALTH_CENTRE,
    "mhc (health-centre level)": HEALTH_CENTRE,
    "mhc": HEALTH_CENTRE,
    # hospitals, kept apart
    "district hospital": DISTRICT_HOSPITAL,
    "dh": DISTRICT_HOSPITAL,
    # provincial hospitals are reported as district hospitals (2026-10-19)
    "provincial hospital": DISTRICT_HOSPITAL,
    "ph": DISTRICT_HOSPITAL,
    "level 2 teaching hospital": L2_TEACHING,
    "l2th": L2_TEACHING,
    "national referral / teaching hospital": NATIONAL_REFERRAL,
    "national referral/teaching hospital": NATIONAL_REFERRAL,
    "chub": NATIONAL_REFERRAL,
    "chuk": NATIONAL_REFERRAL,
}

# lowest tier first, so a gradient reads left to right or top to bottom
ORDER = [HEALTH_POST, HEALTH_CENTRE, DISTRICT_HOSPITAL, L2_TEACHING,
         NATIONAL_REFERRAL]

# what each group absorbed, for the note under a grouped table
MERGED = {
    DISTRICT_HOSPITAL: "District hospital (DH) and provincial hospital (PH)",
    HEALTH_POST: "Second Generation Health Post (SGHP), Health Post (HP), "
                 "and the bare 'HP' value",
    HEALTH_CENTRE: "Health Centre (HC) and MHC (maternity health centre, "
                   "which the form itself marks as health-centre level)",
}


def group(level):
    """Recorded level value -> reporting tier. Unknown values pass through so
    a new form option shows up instead of disappearing into a bucket."""
    v = "" if level is None else str(level).strip()
    return GROUP.get(v.lower(), v or "(not recorded)")


def sort_key(tier):
    return (ORDER.index(tier) if tier in ORDER else len(ORDER), tier)


def ordered(tiers):
    """The tiers present, in reporting order, with anything unrecognised last."""
    return sorted(set(tiers), key=sort_key)
