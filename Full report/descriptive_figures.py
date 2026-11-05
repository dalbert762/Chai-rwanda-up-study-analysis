"""Descriptive figures for the full report (same bases as the dashboard)."""
import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = os.path.dirname(os.path.abspath(__file__))
D = os.path.join(os.path.dirname(HERE), "Inferential analysis", "data")
OUT = os.path.join(HERE, "figures")
os.makedirs(OUT, exist_ok=True)
NAVY, TEAL, GOLD, MAROON, LIGHT, GREY = "#1F3A68", "#117996", "#F3B71B", "#7C1220", "#BFD7EA", "#9AA5B1"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 8.5,
                     "axes.spines.top": False, "axes.spines.right": False})
LV = ["Health post", "Health centre", "District hospital", "Level 2 teaching hospital",
      "National referral hospital"]
LVS = ["HP", "HC", "DH", "L2TH", "NRH"]

mg = pd.read_csv(os.path.join(D, "stock_managed_rows.csv"))
avm = pd.read_csv(os.path.join(D, "availability_managed.csv"))
ave = pd.read_csv(os.path.join(D, "availability_expected.csv"))
chw = pd.read_csv(os.path.join(D, "chw_rows.csv"))
mg["anyso"] = ((mg.so_jun > 0) | (mg.so_jul > 0) | (mg.so_aug > 0)).astype(int)


def save(fig, fn):
    fig.tight_layout()
    fig.savefig(os.path.join(OUT, fn), dpi=200)
    plt.close(fig)


def labels(ax, bars, fmt="{:.1f}%", inside=False, horiz=False):
    for b in bars:
        v = b.get_width() if horiz else b.get_height()
        if horiz:
            ax.text(v + 0.8, b.get_y() + b.get_height() / 2, fmt.format(v), va="center", fontsize=7)
        else:
            ax.text(b.get_x() + b.get_width() / 2, v + 0.8, fmt.format(v), ha="center", fontsize=7)


# d1 availability by programme and by level (two denominators)
fig, axs = plt.subplots(1, 2, figsize=(10, 3.3))
pg = [p for p in ave.programme.dropna().unique() if p not in ("", "(none)")]
a1 = ave.groupby("programme").available.mean().mul(100).reindex(pg)
a2 = avm.groupby("programme").available.mean().mul(100).reindex(pg)
o = a2.sort_values().index
b2 = axs[0].bar(range(len(o)), a2[o], 0.6, color=NAVY, label="Managed products")
labels(axs[0], b2)
em = list(o).index("Essential Medicines")
axs[0].plot([em - 0.3, em + 0.3], [a1["Essential Medicines"]] * 2, color=GOLD, lw=3,
            label="Essential Medicines on the EML denominator")
axs[0].text(em, a1["Essential Medicines"] - 3, f"{a1['Essential Medicines']:.1f}%", ha="center",
            fontsize=7, color="#7a5a00")
axs[0].set_xticks(range(len(o)))
axs[0].set_xticklabels([s.replace("Essential Medicines", "Essential\nMedicines") for s in o], fontsize=7.5)
axs[0].set_ylim(60, 104); axs[0].set_ylabel("Available on the visit day (%)")
axs[0].set_title("By programme", loc="left", fontsize=9)
axs[0].legend(fontsize=7, frameon=False, loc="lower right")
l1 = ave.groupby("level").available.mean().mul(100).reindex(LV)
l2 = avm.groupby("level").available.mean().mul(100).reindex(LV)
x = np.arange(len(LV))
b1 = axs[1].bar(x - 0.2, l1, 0.4, color=LIGHT, label="Expected (two-track)")
b2 = axs[1].bar(x + 0.2, l2, 0.4, color=NAVY, label="Managed products")
axs[1].legend(fontsize=7, frameon=False, loc="lower right")
labels(axs[1], b1); labels(axs[1], b2)
axs[1].set_xticks(x); axs[1].set_xticklabels(LVS); axs[1].set_ylim(60, 104)
axs[1].set_title("By level of care", loc="left", fontsize=9)
save(fig, "d1_availability.png")

# d2 stock adequacy by level and programme
ad = mg[mg.adequacy.notna() & (mg.adequacy != "")]
cats = ["out", "under", "adequate", "over"]
cols = [MAROON, GOLD, LIGHT, NAVY]
names = ["Stocked out", "Understock (<1 month)", "Adequate (1-2 months; 1-3 at L2TH, NRH)", "Overstock (above the band)"]
fig, axs = plt.subplots(1, 2, figsize=(10, 3.2))
for ax, key, order, lab in ((axs[0], "level", LV, LVS),
                            (axs[1], "programme", None, None)):
    ct = pd.crosstab(ad[key], ad.adequacy, normalize="index").mul(100)[cats]
    if order:
        ct = ct.reindex(order)
    else:
        ct = ct.drop(index=[i for i in ct.index if i in ("", "(none)")], errors="ignore").sort_values("over")
    left = np.zeros(len(ct))
    for c, col, nm in zip(cats, cols, names):
        ax.barh(range(len(ct)), ct[c], left=left, color=col, label=nm)
        for i, (l_, v) in enumerate(zip(left, ct[c])):
            if v >= 7:
                ax.text(l_ + v / 2, i, f"{v:.0f}%", ha="center", va="center", fontsize=6.5,
                        color="white" if col in (MAROON, NAVY) else "black")
        left += ct[c].values
    ax.set_yticks(range(len(ct))); ax.set_yticklabels(lab or ct.index, fontsize=7.5)
    ax.set_xlim(0, 100); ax.set_xlabel("% of stocked-out or stocked rows")
axs[0].set_title("By level of care", loc="left", fontsize=9)
axs[1].set_title("By programme", loc="left", fontsize=9)
axs[1].legend(fontsize=6.5, frameon=False, ncol=2, loc="upper center", bbox_to_anchor=(-0.1, -0.22))
save(fig, "d2_adequacy.png")

# d3 accuracy by level and by system
acc = mg[mg.in_accuracy == 1].copy()
acc["exact"] = acc.exact.fillna(0)
fig, axs = plt.subplots(1, 2, figsize=(10, 2.9))
v = acc.groupby("level").exact.mean().mul(100).reindex(LV)
b = axs[0].bar(LVS, v, color=TEAL); labels(axs[0], b)
axs[0].set_ylim(0, 110); axs[0].set_ylabel("Physical = system count (%)")
axs[0].set_title("Record accuracy by level of care", loc="left", fontsize=9)
top = acc.system.value_counts().index[:6]
v = acc[acc.system.isin(top)].groupby("system").exact.mean().mul(100).sort_values()
n = acc[acc.system.isin(top)].groupby("system").size()
b = axs[1].barh([f"{s} ({n[s]:,})" for s in v.index], v, color=NAVY); labels(axs[1], b, horiz=True)
axs[1].set_xlim(0, 105); axs[1].set_title("By digital system (records)", loc="left", fontsize=9)
save(fig, "d3_accuracy.png")

# d4 stock-outs by programme: incidence and mean days
fig, axs = plt.subplots(1, 2, figsize=(10, 3.0))
pgs = mg[~mg.programme.isin(["", "(none)"])]
inc = pgs.groupby("programme").anyso.mean().mul(100).sort_values()
cols_ = [MAROON if x > 9.7 else TEAL for x in inc]
b = axs[0].barh(inc.index, inc, color=cols_); labels(axs[0], b, horiz=True)
axs[0].axvline(9.7, color=GREY, ls="--", lw=0.8)
axs[0].set_xlabel("Pairs with any stock-out day, Jun-Aug (%)")
axs[0].set_title("Stock-out incidence by programme (red = above 9.7%)", loc="left", fontsize=9)
days = pgs[pgs.total_so_days > 0].groupby("programme").total_so_days.mean().reindex(inc.index)
b = axs[1].barh(days.index, days, color=GOLD); labels(axs[1], b, fmt="{:.0f} d", horiz=True)
axs[1].set_xlabel("Mean total stock-out days per affected pair")
axs[1].set_title("Duration once out", loc="left", fontsize=9)
save(fig, "d4_stockouts.png")

# d5 fill rate by programme and month
mg["rq"] = sum(np.where(mg[f"req_{m}"] >= mg[f"del_{m}"], mg[f"req_{m}"], 0) for m in ("jun", "jul", "aug"))
mg["dv"] = sum(np.where(mg[f"req_{m}"] >= mg[f"del_{m}"], mg[f"del_{m}"], 0) for m in ("jun", "jul", "aug"))
fig, axs = plt.subplots(1, 2, figsize=(10, 2.9))
g = mg[~mg.programme.isin(["", "(none)"])].groupby("programme")[["rq", "dv"]].sum()
fr = (g.dv / g.rq * 100).sort_values()
b = axs[0].barh(fr.index, fr, color=[MAROON if x < 90 else NAVY for x in fr]); labels(axs[0], b, horiz=True)
axs[0].set_xlim(0, 108); axs[0].set_xlabel("Quantity delivered / requested (%)")
axs[0].set_title("Order fill rate by programme", loc="left", fontsize=9)
mm = []
for m in ("jun", "jul", "aug"):
    r_ = np.where(mg[f"req_{m}"] >= mg[f"del_{m}"], mg[f"req_{m}"], 0).sum()
    d_ = np.where(mg[f"req_{m}"] >= mg[f"del_{m}"], mg[f"del_{m}"], 0).sum()
    mm.append(100 * d_ / r_)
b = axs[1].bar(["June", "July", "August"], mm, color=TEAL); labels(axs[1], b)
axs[1].set_ylim(80, 100); axs[1].set_title("Order fill rate by month", loc="left", fontsize=9)
save(fig, "d5_fill.png")

# d6 CHW availability by product and province
cm = chw[(chw.managed == 1) & (chw.answered == 1)]
fig, axs = plt.subplots(1, 2, figsize=(10, 3.2))
v = cm.groupby("product").available.mean().mul(100).sort_values()
b = axs[0].barh(v.index, v, color=[MAROON if x < 75 else TEAL for x in v]); labels(axs[0], b, horiz=True)
axs[0].set_xlim(0, 108); axs[0].set_title("CHW availability by product (managed)", loc="left", fontsize=9)
axs[0].tick_params(axis="y", labelsize=7)
v = cm.groupby("province").available.mean().mul(100).sort_values()
b = axs[1].bar(v.index, v, color=NAVY); labels(axs[1], b)
axs[1].set_ylim(0, 105); axs[1].set_title("CHW availability by province", loc="left", fontsize=9)
save(fig, "d6_chw.png")

# d7 consumption: top AMC
top = mg.groupby("medicine").amc.sum().sort_values().tail(12)
fig, ax = plt.subplots(figsize=(7.5, 3.2))
b = ax.barh([m[:38] for m in top.index], top / 1000, color=TEAL)
labels(ax, b, fmt="{:,.0f}k", horiz=True)
ax.set_xlabel("Total AMC across facilities (thousand units per month)")
ax.set_title("Medicines with the highest total monthly consumption", loc="left", fontsize=9)
ax.tick_params(axis="y", labelsize=7)
save(fig, "d7_consumption.png")
print("descriptive figures:", sorted(f for f in os.listdir(OUT) if f.startswith("d")))
