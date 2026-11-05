"""
Pull this study's KoboToolbox exports to local .xlsx files.

    python3 fetch_kobo_export.py                 # every configured dataset
    python3 fetch_kobo_export.py prescription    # just one
    python3 fetch_kobo_export.py --list          # what the token can see
    python3 fetch_kobo_export.py --asset <uid>   # resolve a form's export itself

Two datasets make up the study; see DATASETS below.

  prescription  public export setting, works with no credentials.
  stock         needs an asset uid (KOBO_STOCK_ASSET) or a token - see below.

--------------------------------------------------------------------------
What the API allows, measured against kf.kobotoolbox.org rather than assumed
--------------------------------------------------------------------------
  GET /api/v2/assets/                       anonymously returns count: 0, so a
                                            form cannot be DISCOVERED without a
                                            token, only fetched if you know its uid.
  GET /api/v2/assets/<uid>/                 readable anonymously for a publicly
                                            shared form; 404 otherwise.
  GET /api/v2/assets/<uid>/export-settings/ readable anonymously too - so given
                                            just an asset uid, --asset resolves
                                            the export setting on its own.
  .../export-settings/<es>/data.xlsx        the actual workbook.

Kobo returns 404, never 403, for anything the caller cannot see: a 404 means
"wrong uid" OR "needs a token", and the two are indistinguishable from outside.

A private form therefore needs a token:

    export KOBO_API_TOKEN=<your token>       # or the macOS keychain, see find_token()

from https://kf.kobotoolbox.org/token/?format=json while logged in (Account
Settings -> Security shows the same value), or stored once with:

    security add-generic-password -s kobo -a "$USER" -w <your token>

IMPORTANT: a public export setting REJECTS a bad Authorization header that it
would have served with no header at all, so a token is only ever sent when one
was actually found.

Standard library only (urllib), like the other scripts here.
"""

import argparse
import json
import os
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request

BASE = "https://kf.kobotoolbox.org/api/v2"

# asset          the form's uid (the 'a...' segment of its URL)
# export_setting the saved export whose column selection / language to use;
#                None means "resolve it from the asset at run time"
# out            destination workbook
DATASETS = {
    "prescription": {
        "label": "Unserved prescription assessment",
        "asset": "aj9BEpgxC8i6DCvurnxiLD",
        "export_setting": "esk6wkWnPPXKPMopcMQBKic",
        # A PART, merged into Prescription_kobo_export.xlsx - see MERGES.
        "out": "Prescription_kobo_export_kf.xlsx",
    },
    "prescription_eu": {
        "label": "Unserved prescription assessment (eu server)",
        # Same form, same second team as stock_eu (owner dukuze1609), on
        # Kobo's EU server. A full URL short-circuits uid resolution, which is
        # what lets one script span two hosts.
        "url": (os.environ.get("KOBO_PRESC_EU_URL")
                or "https://eu.kobotoolbox.org/api/v2/assets/"
                   "aPPum7zLxUrzW68z8zWZiM/export-settings/"
                   "eskwfEUfpyuZh6zZbewogyW/data.xlsx"),  # named "prescription2"
        "out": "Prescription_kobo_export_eu.xlsx",
    },
    "stock": {
        "label": "Medicine stock assessment",
        # Publicly shared, like the prescription project, so no token needed.
        # aJgix3JuUqMjx3Gax5Jp3r returns 404 and is NOT the live export.
        # (adtEfgzyyULozJVAxfjRLD is a different form - the CHW one below.)
        # This one answers 302 -> /private-media/AnonymousUser/exports/...,
        # so the redirect must be followed (urllib does that for GET).
        # Env vars still override, KOBO_EXPORT_URL being the same variable
        # analyze_medicine_assessment.py documents.
        "url": os.environ.get("KOBO_EXPORT_URL") or None,
        "asset": os.environ.get("KOBO_STOCK_ASSET") or "ae5rqsV39C7DRscetStjRT",
        "export_setting": (os.environ.get("KOBO_STOCK_EXPORT")
                           or "esfUANo34n6tXBtcyDZLcUc"),
        # One of two PARTS that are merged into the file the analysis reads -
        # see STOCK_PARTS below. Not the analysis input itself.
        "out": "Medicine_stock_assessment_kf.xlsx",
    },
    "stock_eu": {
        "label": "Medicine stock assessment (eu server)",
        # The SAME form, collected by a second team on Kobo's EU server under a
        # different account (owner dukuze1609), so it is a different host, a
        # different asset uid and a different export setting - nothing about it
        # can be resolved from the kf uids. A full URL short-circuits the uid
        # resolution, which is what lets one script span two servers without
        # BASE becoming a per-dataset setting.
        "url": (os.environ.get("KOBO_STOCK_EU_URL")
                or "https://eu.kobotoolbox.org/api/v2/assets/"
                   "aaScwQoxCZNDCfDWgJAVox/export-settings/"
                   "esQNoMCyKsdWMfYn5FcXcXY/data.xlsx"),  # named "Stockassessment1"
        "out": "Medicine_stock_assessment_eu.xlsx",
    },
    "chw": {
        "label": "Medicine stock assessment for CHWs",
        # Separate form from "stock": same medicine-availability questions, but
        # the unit of observation is a CHW attached to a facility, and the
        # medicine list is the CHW package (ACT/ALu, RDTs, Zinc, ORS, Amox,
        # DHAP, Pyronaridine/Artesunate). Publicly shared, no token needed.
        "asset": os.environ.get("KOBO_CHW_ASSET") or "adtEfgzyyULozJVAxfjRLD",
        "export_setting": (os.environ.get("KOBO_CHW_EXPORT")
                           or "estFypyXnbDhuXNnmGmfbiB"),  # named "Stock_CHWs"
        # A PART, merged into Medicine_stock_CHW_export.xlsx - see MERGES.
        "out": "Medicine_stock_CHW_kf.xlsx",
    },
    "chw_eu": {
        "label": "Medicine stock assessment for CHWs (eu server)",
        # Same form, same second team (owner dukuze1609), on Kobo's EU server.
        # This asset has TWO export settings; esULyP7VeUPLgsZJstJEnSE
        # ("MedicineCHW1") is the one to use - the other is named
        # "Export Invalid date".
        "url": (os.environ.get("KOBO_CHW_EU_URL")
                or "https://eu.kobotoolbox.org/api/v2/assets/"
                   "abkZvkGHiLgdDnC7gkXBPx/export-settings/"
                   "esULyP7VeUPLgsZJstJEnSE/data.xlsx"),
        "out": "Medicine_stock_CHW_eu.xlsx",
    },
}

XLSX_MAGIC = b"PK\x03\x04"

# ---------------------------------------------------------------
# the stock assessment spans two servers
# ---------------------------------------------------------------
# Two teams collect the same form into two Kobo accounts on two hosts. Neither
# export alone is the dataset, so the parts are fetched separately and merged
# into the single file analyze_medicine_assessment.py reads - fetch -> analyse
# still works with no copying or renaming in between.
SERVER_COL = "_server"         # provenance, added by the merge

# For each merged dataset:
#   parts    (server tag, part file) in the order they are concatenated
#   merged   the file the analysis reads - rebuilt, never appended to
#   repeats  repeat-group sheets, named by the XLSForm. Sheet 1 is whatever
#            sheet is not one of these, because it is named after the FORM and
#            a rename in Kobo renames it.
#   reindex  offset _index / _parent_index per part. Needed only where rows are
#            joined on _index ALONE; see merge_parts().
#   aliases  {column as the EU part names it: column the analysis expects}.
#            The two servers run different form versions, and a renamed column
#            is worse than a missing one - it looks populated while landing in
#            a column nothing reads.
MERGES = {
    "stock": {
        "parts": [("kf", "Medicine_stock_assessment_kf.xlsx"),
                  ("eu", "Medicine_stock_assessment_eu.xlsx")],
        "merged": "Medicine_stock_assessment_export.xlsx",
        "repeats": ["medicine_group"],
        "reindex": True,
        "aliases": {},
    },
    "prescription": {
        "parts": [("kf", "Prescription_kobo_export_kf.xlsx"),
                  ("eu", "Prescription_kobo_export_eu.xlsx")],
        "merged": "Prescription_kobo_export.xlsx",
        "repeats": ["prescription", "medicine_item"],
        # Rows here join on (_submission__uuid, _index), and _uuid is unique
        # across servers, so colliding _index values cannot cross-match.
        # Offsetting would have to be applied to both sheets in step and buys
        # nothing, so it is not.
        "reindex": False,
        "aliases": {
            "Medicine name (not on the EML list, or not found in the list "
            "above) - please specify":
                "Medicine name (not on the EML list) - please specify",
            # kf's calculated EML flag; the EU version renamed it when it added
            # the explicit "Is this medicine on the EML?" question.
            "on_eml_final": "on_eml",
        },
    },
    "chw": {
        "parts": [("kf", "Medicine_stock_CHW_kf.xlsx"),
                  ("eu", "Medicine_stock_CHW_eu.xlsx")],
        "merged": "Medicine_stock_CHW_export.xlsx",
        "repeats": ["medicine_group"],
        # Joins on _index alone, like the stock form.
        "reindex": True,
        # The two versions' medicine_group sheets are column-identical; only
        # sheet 1 differs, and only by the timing/GPS fields the newer kf
        # version adds.
        "aliases": {},
    },
}


# ---------------------------------------------------------------
# credentials
# ---------------------------------------------------------------
def find_token():
    """Env var first, then the macOS keychain. None if neither has one."""
    token = os.environ.get("KOBO_API_TOKEN") or os.environ.get("KOBO_TOKEN")
    if token:
        return token.strip()
    try:
        out = subprocess.run(["security", "find-generic-password", "-s", "kobo", "-w"],
                             capture_output=True, text=True, timeout=10)
    except (OSError, subprocess.SubprocessError):
        return None
    return out.stdout.strip() or None


# ---------------------------------------------------------------
# http
# ---------------------------------------------------------------
def get(url, token, timeout=300):
    """Anonymous when token is None - never send a header we have no token for."""
    headers = {"Authorization": f"Token {token}"} if token else {}
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def get_json(url, token):
    return json.loads(get(url, token, timeout=60).decode("utf-8"))


def explain(code, token):
    if code == 401:
        return ("the token was rejected - check it was copied whole" if token else
                "no token was sent and this resource is not public")
    if code == 403:
        return "the token is valid but this account may not read that resource"
    if code == 404:
        return ("not visible to this caller. Kobo returns 404 (never 403) for "
                "anything you lack access to, so this is either a wrong uid or a "
                "private form needing KOBO_API_TOKEN"
                + ("" if token else " - no token was sent"))
    return ""


class ApiError(Exception):
    pass


def api(url, token, what):
    try:
        return get_json(url, token)
    except urllib.error.HTTPError as e:
        raise ApiError(f"{what}: HTTP {e.code} {e.reason} - {explain(e.code, token)}")
    except urllib.error.URLError as e:
        raise ApiError(f"{what}: could not reach {url} ({e.reason})")


# ---------------------------------------------------------------
# discovery
# ---------------------------------------------------------------
def list_assets(token):
    """Every form the caller can see, with its export settings."""
    d = api(f"{BASE}/assets/?format=json&limit=200", token, "listing forms")
    results = d.get("results", [])
    if not results:
        print("No forms visible."
              + ("" if token else " No token was sent - an anonymous caller always"
                                  " sees an empty list, even for public forms."))
        return
    print(f"{d.get('count', len(results))} form(s) visible:\n")
    for a in results:
        uid = a.get("uid", "?")
        print(f"  {a.get('name') or '(untitled)'}")
        print(f"    asset uid   : {uid}")
        print(f"    owner       : {a.get('owner__username')}")
        print(f"    submissions : {a.get('deployment__submission_count')}")
        try:
            es = api(f"{BASE}/assets/{uid}/export-settings/?format=json", token,
                     "listing export settings")
            for s in es.get("results", []):
                print(f"    export      : {s['uid']}  "
                      f"name={s.get('name') or '(unnamed)'}")
        except ApiError as e:
            print(f"    export      : unavailable ({e})")
        print()


def resolve_export(asset, token):
    """Pick an export setting for a form. Export settings are listable even
    anonymously, so knowing only the asset uid is enough for a public form."""
    d = api(f"{BASE}/assets/{asset}/export-settings/?format=json", token,
            f"listing export settings for {asset}")
    results = d.get("results", [])
    if not results:
        raise ApiError(f"{asset} has no saved export settings - create one in Kobo "
                       f"(Data -> Downloads), or pass --url for a direct export URL")
    # Prefer an xls/xlsx setting; otherwise take the first.
    for s in results:
        if str(s.get("export_settings", {}).get("type", "")).lower() in ("xls", "xlsx"):
            return s["uid"], s.get("name") or "(unnamed)", len(results)
    return results[0]["uid"], results[0].get("name") or "(unnamed)", len(results)


# ---------------------------------------------------------------
# fetch
# ---------------------------------------------------------------
def data_url(asset, export_setting):
    return f"{BASE}/assets/{asset}/export-settings/{export_setting}/data.xlsx"


def fetch_to(url, out, token, label):
    try:
        body = get(url, token)
    except urllib.error.HTTPError as e:
        detail = e.read()[:300].decode("utf-8", "replace")
        raise ApiError(f"{label}: HTTP {e.code} {e.reason} - {explain(e.code, token)}"
                       + (f"\n    server said: {detail}" if detail.strip() else ""))
    except urllib.error.URLError as e:
        raise ApiError(f"{label}: could not reach {url} ({e.reason})")

    # An export still building, or an error page, arrives as JSON/HTML with a
    # 200 - so check it really is a workbook before overwriting anything.
    if not body.startswith(XLSX_MAGIC):
        raise ApiError(f"{label}: response was not an .xlsx file. First bytes:\n    "
                       + body[:300].decode("utf-8", "replace"))

    with open(out, "wb") as fh:
        fh.write(body)
    print(f"  {label}: wrote {out} ({len(body):,} bytes)")
    return True


def fetch_dataset(name, cfg, token):
    label = cfg.get("label", name)
    asset, es = cfg.get("asset"), cfg.get("export_setting")

    # A full data URL short-circuits uid resolution entirely.
    if cfg.get("url"):
        try:
            return fetch_to(cfg["url"], cfg["out"], token, label)
        except ApiError as e:
            print(f"  {e}")
            return False

    if not asset:
        print(f"  {label}: SKIPPED - no asset uid configured.")
        print(f"      Supply any one of:")
        print(f"        export KOBO_EXPORT_URL=<full .../export-settings/<uid>/"
              f"data.xlsx>")
        print(f"        export KOBO_{name.upper()}_ASSET=<asset uid>")
        print(f"        export KOBO_API_TOKEN=<token>   then: "
              f"fetch_kobo_export.py --list")
        print(f"      This form is not public and anonymous search returns nothing, "
              f"so the uid cannot be discovered - it has to be provided.")
        return False
    if not es:
        try:
            es, es_name, n = resolve_export(asset, token)
        except ApiError as e:
            print(f"  {label}: {e}")
            return False
        print(f"  {label}: using export setting {es} ({es_name})"
              + (f" of {n}" if n > 1 else ""))
    try:
        return fetch_to(data_url(asset, es), cfg["out"], token, label)
    except ApiError as e:
        print(f"  {e}")
        return False


# ---------------------------------------------------------------
# merge
# ---------------------------------------------------------------
def read_sheet(wb, name):
    rows = list(wb[name].iter_rows(values_only=True))
    if not rows:
        return [], []
    hdr = [str(h).strip() if h is not None else "" for h in rows[0]]
    return hdr, [r for r in rows[1:] if any(c is not None for c in r)]


def merge_parts(name, spec):
    """Concatenate one dataset's server parts into the file the analysis reads.

    Rebuilt from the parts every time rather than appended to, so running the
    fetch twice cannot double the data.

    Three things have to be handled or the merge is silently wrong:

      * _index restarts at 1 on each server. Where rows join on _index alone
        (the stock form's medicine_group -> submission), every EU row would
        otherwise attach itself to a kf facility, so those parts are offset
        past the highest _index so far. Where the join also carries
        _submission__uuid (the prescription form), collisions cannot
        cross-match and no offset is applied.
      * the parts run different form VERSIONS, so a column one part has may be
        missing from the other. Columns are unioned in first-seen order and a
        row has nothing under a column its version never had - which is what
        the analysis already expects of a multi-version export.
      * worse, a version can RENAME a column. Then both names survive the
        union, each half-populated, and every rate computed from it is quietly
        halved. spec["aliases"] maps those back onto one name.
    """
    try:
        import openpyxl
    except ImportError:
        print(f"  {spec['merged']}: SKIPPED - openpyxl is not installed")
        return False

    parts = [(tag, f) for tag, f in spec["parts"] if os.path.exists(f)]
    if not parts:
        return False
    repeats, aliases = spec["repeats"], spec["aliases"]

    order, rows_by_sheet, sheet1 = {}, {}, None
    offset = 0
    for tag, path in parts:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        name1 = next(n for n in wb.sheetnames if n not in repeats)
        sheet1 = sheet1 or name1
        hi, counts = 0, []
        for sheet in [name1] + [r for r in repeats if r in wb.sheetnames]:
            key = sheet1 if sheet == name1 else sheet
            hdr, raw = read_sheet(wb, sheet)
            hdr = [aliases.get(c, c) for c in hdr]
            seen = order.setdefault(key, [])
            for c in hdr:
                if c and c not in seen:
                    seen.append(c)
            out = rows_by_sheet.setdefault(key, [])
            for r in raw:
                d = dict(zip(hdr, r))
                if spec["reindex"]:
                    for col in ("_index", "_parent_index"):
                        v = d.get(col)
                        if isinstance(v, (int, float)):
                            if col == "_index" and sheet == name1:
                                hi = max(hi, int(v))
                            d[col] = int(v) + offset
                d[SERVER_COL] = tag
                out.append(d)
            counts.append(f"{len(raw):,} {key}")
        print(f"  merge {name}: {tag} contributes " + ", ".join(counts)
              + (f" (_index offset by {offset:,})" if spec["reindex"] and offset
                 else ""))
        offset += hi

    for seen in order.values():
        if SERVER_COL not in seen:
            seen.append(SERVER_COL)

    # Sheet 1 is the parent of everything; a repeat row that cannot find it is
    # dropped by the analysis, so refuse to write rather than lose rows quietly.
    if spec["reindex"]:
        idx = {r.get("_index") for r in rows_by_sheet[sheet1]}
        orphans = sum(1 for k, rs in rows_by_sheet.items() if k != sheet1
                      for r in rs if r.get("_parent_index") not in idx)
    else:
        uu = {r.get("_uuid") for r in rows_by_sheet[sheet1]}
        orphans = sum(1 for k, rs in rows_by_sheet.items() if k != sheet1
                      for r in rs if r.get("_submission__uuid") not in uu)
    if orphans:
        print(f"  merge {name}: WARNING {orphans:,} repeat rows have no parent "
              f"submission - {spec['merged']} NOT written")
        return False

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)
    for key in [sheet1] + [r for r in repeats if r in rows_by_sheet]:
        ws = out_wb.create_sheet(key[:31])
        ws.append(order[key])
        for d in rows_by_sheet[key]:
            ws.append([d.get(c) for c in order[key]])
    dest = spec["merged"]
    if os.path.exists(dest):
        os.chmod(dest, 0o644)   # Box resets synced files to read-only
    out_wb.save(dest)
    print(f"  merge {name}: wrote {dest} - "
          + ", ".join(f"{len(rows_by_sheet[k]):,} {k}"
                      for k in [sheet1] + [r for r in repeats if r in rows_by_sheet])
          + f" from {len(parts)} server(s)")
    return True


def main():
    ap = argparse.ArgumentParser(
        description="Pull the study's KoboToolbox exports.",
        epilog="Datasets: " + ", ".join(DATASETS))
    # No choices= here: argparse validates its own empty default against it
    # when nargs="*", which rejects the plain "fetch everything" invocation.
    ap.add_argument("datasets", nargs="*", metavar="DATASET",
                    help="which to fetch (default: all)")
    ap.add_argument("--list", action="store_true",
                    help="list the forms and export settings the token can see")
    ap.add_argument("--asset", help="fetch this asset uid, resolving its export setting")
    ap.add_argument("--url", help="fetch a direct export-settings data URL")
    ap.add_argument("--out", help="destination for --asset / --url")
    args = ap.parse_args()

    unknown = [d for d in args.datasets if d not in DATASETS]
    if unknown:
        ap.error(f"unknown dataset(s) {', '.join(unknown)}; "
                 f"choose from {', '.join(DATASETS)}")

    token = find_token()
    print("Kobo API token: " + ("found" if token else "none (anonymous - public "
                                                       "exports only)"))

    try:
        if args.list:
            list_assets(token)
            return 0

        if args.url:
            out = args.out or "kobo_export.xlsx"
            fetch_to(args.url, out, token, "direct URL")
            return 0

        if args.asset:
            es, es_name, n = resolve_export(args.asset, token)
            print(f"  export setting {es} ({es_name})" + (f" of {n}" if n > 1 else ""))
            out = args.out or f"{args.asset}_export.xlsx"
            fetch_to(data_url(args.asset, es), out, token, args.asset)
            return 0
    except ApiError as e:
        print(f"  {e}")
        return 1

    names = args.datasets or list(DATASETS)
    ok = sum(bool(fetch_dataset(n, DATASETS[n], token)) for n in names)
    print(f"{ok} of {len(names)} dataset(s) fetched.")

    # Either stock part landing means the merged file is now out of date.
    # It is rebuilt from whichever parts are on disk, so fetching one part
    # still leaves a consistent file rather than a half-updated one.
    for mname, spec in MERGES.items():
        if any(n in (mname, mname + "_eu") for n in names):
            merge_parts(mname, spec)
    return 0 if ok == len(names) else 1


if __name__ == "__main__":
    sys.exit(main())
