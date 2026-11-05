"""Code the stock tool's free-text "Why is <medicine> not managed?" answer.

    import not_managed as NM
    NM.theme("They use DHAP")      -> "MFT zoning - another antimalarial ..."

Shared by MCCH_FP_availability.py and Full_indicator_report.py so the two
cannot code the same answer two ways. The rules are regular expressions,
tried in order - FIRST MATCH WINS - so the order is part of the coding:
faith-based before "no FP service", MFT zoning before "uses another
product", another strength before the generic "they use ...". Anything no
rule matches comes back as OTHER; check its count after every refresh and
add a rule until it is small, as with the CHW feedback coder.
"""
import re

NO_REASON = "(no reason recorded)"
OTHER = "Other"

RULES = [
    ("Faith-based: no modern FP by policy",
     r"faith|catholi|religious|church|natur|non natural|artificial"),
    ("Service not offered here (no FP / maternity / deliveries)",
     r"no (planning family|fp|pf|maternit)|planning family service|"
     r"no pf service|don.?t have maternity|not pe?r?forming deliver|"
     r"\bncds?\b|no hiv program"),
    ("MFT zoning - another antimalarial allocated to this zone",
     r"\bmft\b|aspy|\bzon(e|ing)\b|\bblock\b|\bdhap\b|coartem|artemether|"
     r"arthemeter|lumefant|pyronar|\bregion\b|regime|regimen|"
     r"southern province uses|\bal\b|\bbloc\b|rotation of antimalar"),
    ("Holds another strength of the same medicine",
     r"\b(manage[sd]?|use[sd]?|using)\b.*\d+\s*(mg|gr|g|mcg)\b|"
     r"\d+\s*mg (only|instead)|only (have )?\d+\s*(mg|g|gr)\b|"
     r"\d+\s*(mg|g|gr)\b.*\b(managed|is managed)\b|priority is given|"
     r"double the dose|concentration|\b(taken|counted|have|has)\b.*"
     r"\d+\s*(mg|g|gr)\b|\d+\s*(mg|g|gr) tab is managed"),
    ("Rejected by RSSB / cost",
     r"rssb|reimburs|not paid|high cost|expensive|\bcost\b"),
    ("Uses an alternative product",
     r"\buse[sd]?\b.*(oxytocin|carbetocin|vit\.? ?k|sayana|mirena|"
     r"tranexamic)|replaced by|only mirena|alternative|other molecule|"
     r"using carbetocin|facility used|hormonal mostly|moved to|"
     r"\bthey use\b|\busing\b|\buses\b|catgut|substitut|oral forms|"
     r"only (request|use)|serves the same purpose"),
    ("Believed not allowed at this level",
     r"not allowed|allowed only|only allowed|not (applicable|application)|"
     r"no applicable|not (on|at) (their|this|hf|level|hc)|not at hc|"
     r"at (hc|dh|hp) level|level of h|used at dh|dh level|"
     r"only to be used +dh|not used at hc|it is applicable at|"
     r"is not managed at hc|not managed at health|"
     r"product is not managed at|only for refill|for prescription refill|"
     r"no available prescription for refill|for refills|"
     r"product used by hcs"),
    ("New product: not yet introduced, ordered or trained",
     r"(haven.?t|hasnt|havent) started|not yet|new (drug|on|item)|"
     r"first time|not aware|unaware|uninformed|didn.?t know|did not know|"
     r"don.?t know it|no (enough )?information|not know|protocol|train|"
     r"waiting for|planning to order|will (place|make|start|be request)|"
     r"not (yet )?requested|not ordered|no ordered|didn.?t order|"
     r"not decentrali|instruction|mentorship|not yet received|"
     r"not requested by clinicians|pack size|allowed to use|"
     r"information about|never procured|not procured|delay to request"),
    ("No demand / no cases / referred",
     r"no (cases?|clients?|patients?|customers|demand|beneficiary|"
     r"benecifiary)|no client|not moving|refer|transfer|no patient case|"
     r"treated by|preff?err?ed|zero case|"
     r"when case|not prefer|not commonly|didn.?t feel|don.?t have "
     r"(customers|clients)|client|not needed|patient prefers|"
     r"do not have patient|no patient|no +patient|don.?t have patient|"
     r"haven.?t had any case|don.?t need|when there is a case|"
     r"not utili[sz]ed"),
    ("Not on the facility's selection list / formulary",
     r"selection list|formulary|not on the list|usable list|mtc|"
     r"eml list of the hospital|list of needed|\bhf choice\b"),
    ("Reported not available / not supplied",
     r"not available|no available|note available|not acquired|"
     r"haven.?t had it|stock card|lack other material"),
    ("Not used / not managed (no reason)",
     r"not used|no ?t? managing|not managing|not managed|don.?t use|"
     r"don.?t manage|don.?t managed|not  managed|they don.?t use|"
     r"^not$|not used of|not being used|not recommended|manageg"),
]
_RX = [(lab, re.compile(pat, re.I)) for lab, pat in RULES]


def theme(text):
    """Free-text reason -> theme label."""
    t = "" if text is None else str(text).strip()
    low = t.lower()
    if (not t or t in (".", "None", "0") or "no reason" in low
            or "no significant reason" in low or low == "not explained"):
        return NO_REASON
    for lab, rx in _RX:
        if rx.search(t):
            return lab
    return OTHER
