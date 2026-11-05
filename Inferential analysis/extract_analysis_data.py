"""Flat analysis files for the inferential / predictive analysis.

    python3 extract_analysis_data.py  ->  data/*.csv

Everything is taken from the dashboard's own build (Full_indicator_report),
so every base, flag and cleaning rule is the one the dashboard reports.
"""
import contextlib
import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "analytics"))
os.chdir(os.path.join(os.path.dirname(HERE), "analytics"))
with contextlib.redirect_stdout(io.StringIO()):
    import Full_indicator_report as R
    import rms_hubs as HUB
    import Tracer_availability_mapping as MAP
import pandas as pd

s, M, C, P = R.s, R.M, R.C, R.P
OUT = os.path.join(HERE, "data")
os.makedirs(OUT, exist_ok=True)
num = lambda v: M.num(v)
MON = ("June", "July", "August")


def fac_attrs(p):
    return dict(facility=s(p[M.FAC]), district=s(p[M.DIST]),
                province=s(p[M.PROV]), hub=HUB.HUB_OF.get(s(p[M.DIST]), ""),
                level=R.FL.group(p[M.LEVEL]),
                management=s(p.get(R.EX.MGMT)) or "(not recorded)",
                collector=s(p.get(M.ENUM)))


acc_ids = {id(r) for r in R._acc}
so_ids = {id(r) for r in R.SA_OUT}
shelf_ids = {id(r) for r in R.SA_SHELF}
rows = []
for r in R.MG:
    d = fac_attrs(r["_p"])
    d.update(medicine=s(r[M.MED]), molecule=R.MOL.key(r[M.MED]),
             programme=s(r[M.PROGRAM]) or "(none)",
             drug_class=s(r.get("drug_class_name")),
             price=num(r.get("unit_tariff_price")),
             answered=not M.blank(r[M.AVAIL]),
             available=int(M.yes(r[M.AVAIL])),
             qty_phys=num(r[M.QTY_PHYS]), qty_sys=num(r[M.QTY_SYS]),
             system=s(r.get("_emr_clean")),
             in_accuracy=int(id(r) in acc_ids),
             exact=int(R._exact(r)) if id(r) in acc_ids else None,
             amc=num(r[M.AMC]), mos=R._mos(r),
             adequacy=("out" if id(r) in so_ids else
                       ["under", "adequate", "over"][R._band(r)]
                       if id(r) in shelf_ids else ""),
             cur_so_days=num(r[M.SO_NOW]),
             total_so_days=R._tdays(r),
             rms_central=s(r[M.RMS_CEN]), rms_regional=s(r[M.RMS_REG]),
             dist_class=s(r.get("distribution_nationwide")))
    for m in MON:
        k = m[:3].lower()
        d[f"cons_{k}"] = num(r[M.CONS[m]])
        d[f"so_{k}"] = num(r[M.SO_MONTH[m]]) or 0
        d[f"req_{k}"] = num(r[M.REQ[m]]) or 0
        d[f"del_{k}"] = num(r[M.DEL[m]]) or 0
    rows.append(d)
mg = pd.DataFrame(rows)
mg.to_csv(os.path.join(OUT, "stock_managed_rows.csv"), index=False)


def mol_frame(rows_, name):
    out = []
    for r in rows_:
        d = fac_attrs(r["_p"])
        d.update(molecule=R.MOL.key(r[M.MED]), programme=s(r[M.PROGRAM]),
                 track=r.get("_track", ""), available=int(M.yes(r[M.AVAIL])))
        out.append(d)
    pd.DataFrame(out).to_csv(os.path.join(OUT, name), index=False)


mol_frame(R.MOLS, "availability_expected.csv")
mol_frame(R.MOLS_B, "availability_managed.csv")

# facility coordinates: dashboard map points (GIS + form GPS, name-matched)
pts = []
for v in MAP.POINTS:
    pts.append(dict(facility=v["fac"], lat=v["la"], lon=v["lo"], src=v["src"]))
pd.DataFrame(pts).to_csv(os.path.join(OUT, "facility_points.csv"), index=False)

# district centroids from the boundary file (for spatial weights)
dc = []
for d in MAP.DROWS:
    dc.append(dict(district=d["district"], province=d["prov"],
                   cx=d["cx_cy"][0], cy=d["cx_cy"][1]))
pd.DataFrame(dc).to_csv(os.path.join(OUT, "district_centroids.csv"), index=False)

# CHW rows
ch = []
for r in C.MEDROWS:
    p_ = r["_p"]
    ch.append(dict(facility=C.s(p_[C.FAC]), district=C.s(p_[C.DIST]),
                   province=C.s(p_[C.PROV]), chw=C.s(p_[C.CHW]),
                   product=r["_short"], family=r.get("_family", ""),
                   managed=int(C.yes(r[C.MANAGED])),
                   answered=int(not C.blank(r[C.AVAIL])),
                   available=int(C.yes(r[C.AVAIL])),
                   qty=C.num(r[C.QTY])))
pd.DataFrame(ch).to_csv(os.path.join(OUT, "chw_rows.csv"), index=False)
print("wrote", os.listdir(OUT), len(mg), "managed rows")
