# make_dashboard.py
# Draws my one page dashboard and saves it as docs/dashboard.png
# It reads the tables from data/dashboard so run export_dashboard_data.py first.

from pathlib import Path
import pandas as pd
import matplotlib
matplotlib.use("Agg")   # draws straight to a file without opening a window
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

ROOT = Path(__file__).resolve().parents[1]
D = ROOT / "data" / "dashboard"
fullyear = pd.read_csv(D / "dash_fullyear.csv")
share = pd.read_csv(D / "dash_market_share.csv")
pricing = pd.read_csv(D / "dash_pricing_vs_market.csv")
sens = pd.read_csv(D / "dash_sensitivity.csv")
breakeven = pd.read_csv(D / "dash_breakeven.csv")

# colours
# note: a $ in matplotlib text has to be written as \$ or two of them turn the text in between into maths
INK, INK2, GRID, BACKGROUND, CARD = "#0b0b0b", "#52514e", "#e4e3df", "#fcfcfb", "#ffffff"
SCENARIO_COLOUR = {"Actual": "#52514e", "Base": "#2a78d6", "Bear": "#eb6834", "Bull": "#1baf7a"}
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9, "axes.edgecolor": GRID,
                     "axes.labelcolor": INK2, "xtick.color": INK2, "ytick.color": INK2})

fig = plt.figure(figsize=(16, 9), facecolor=BACKGROUND)
fig.text(0.03, 0.955, "Card surcharge ban: can Tyro still hit its FY27 guidance?", fontsize=18, weight="bold", color=INK)
fig.text(0.03, 0.925, "Gross profit forecast by scenario  |  Data: Tyro reports FY23 to FY26, RBA tables C1.1 C2.1 C3  |  Reform started 1 Oct 2026",
         fontsize=10, color=INK2)


def tidy(ax, title):
    ax.set_facecolor(CARD)
    for side in ["top", "right"]:
        ax.spines[side].set_visible(False)
    ax.grid(axis="y", color=GRID, lw=0.8)
    ax.set_axisbelow(True)
    ax.set_title(title, loc="left", fontsize=11, weight="bold", color=INK, pad=10)


# the 4 number cards along the top
base27 = fullyear[(fullyear.fy == "FY27") & (fullyear.scenario == "Base")].iloc[0]
need_at_1bp = breakeven.loc[breakeven.squeeze_bps == 1, "core_growth_needed_for_midpoint"].iloc[0]
cards = [("FY27 gross profit (Base)", f"${base27.total_gp_m:.1f}m", "Guidance is \\$240m to \\$255m, so just below"),
         ("Market share, H2 FY26", f"{share.tyro_share_pct.iloc[-1]:.1f}%", f"down from {share.tyro_share_pct.iloc[0]:.1f}% in H1 FY23"),
         ("Price above the market", f"+{pricing.tyro_premium_bps.iloc[-1]:.0f} bps", "Tyro about 97 bps vs market about 62 bps"),
         ("Core growth needed for midpoint", f"{need_at_1bp:.1%}", "with a 1 bp squeeze, recent growth 1.7% to 4.4%")]
for i, (label, value, small) in enumerate(cards):
    ax = fig.add_axes([0.03 + i * 0.2375, 0.76, 0.225, 0.13])
    ax.axis("off")
    ax.add_patch(FancyBboxPatch((0, 0), 1, 1, boxstyle="round,pad=0,rounding_size=0.04", fc=CARD, ec=GRID, transform=ax.transAxes))
    ax.text(0.06, 0.74, label, fontsize=10, color=INK2, transform=ax.transAxes)
    ax.text(0.06, 0.36, value, fontsize=22, weight="bold", color=INK, transform=ax.transAxes)
    ax.text(0.06, 0.12, small, fontsize=8.5, color=INK2, transform=ax.transAxes)


# chart 1: total gross profit, actuals then the 3 scenarios, with the FY27 guidance shaded
ax1 = fig.add_axes([0.045, 0.40, 0.44, 0.30])
tidy(ax1, "Total gross profit ($m): actual and forecast by scenario")
actual = fullyear[fullyear.scenario == "Actual"]
years = ["FY23", "FY24", "FY25", "FY26", "FY27", "FY28", "FY29"]
x = {fy: i for i, fy in enumerate(years)}
ax1.add_patch(plt.Rectangle((3.8, 240), 0.4, 15, color="#2a78d6", alpha=0.12, lw=0))
ax1.text(4, 186, "Shaded = FY27 guidance \\$240m to \\$255m", ha="center", fontsize=8, color=INK2)
ax1.plot([x[f] for f in actual.fy], actual.total_gp_m, color=SCENARIO_COLOUR["Actual"], lw=2, marker="o", ms=5)
for s in ["Bear", "Base", "Bull"]:
    d = fullyear[fullyear.scenario == s]
    # start each scenario line from the FY26 actual so the lines join up
    xs = [x["FY26"]] + [x[f] for f in d.fy]
    ys = [actual.total_gp_m.iloc[-1]] + list(d.total_gp_m)
    ax1.plot(xs, ys, color=SCENARIO_COLOUR[s], lw=2, ls="-" if s == "Base" else "--", marker="o", ms=4)
    ax1.text(x["FY29"] + 0.1, d.total_gp_m.iloc[-1], f" {s} ${d.total_gp_m.iloc[-1]:.0f}m", va="center", fontsize=8.5, color=INK)
ax1.text(x["FY26"], actual.total_gp_m.iloc[-1] - 9, "Actual", ha="center", fontsize=8.5, color=INK)
ax1.set_xticks(range(len(years)))
ax1.set_xticklabels(years)
ax1.set_xlim(-0.3, 7.2)
ax1.set_ylim(180, 315)


# chart 2: market share over time
ax2 = fig.add_axes([0.55, 0.40, 0.42, 0.30])
tidy(ax2, "Tyro's share of Australian card spending (%)")
ax2.plot(range(len(share)), share.tyro_share_pct, color=SCENARIO_COLOUR["Base"], lw=2, marker="o", ms=5)
for i in (0, len(share) - 1):
    ax2.text(i, share.tyro_share_pct.iloc[i] + 0.06, f"{share.tyro_share_pct.iloc[i]:.2f}%", ha="center", fontsize=8.5, color=INK)
ax2.set_xticks(range(len(share)))
ax2.set_xticklabels(share.period, fontsize=8)
ax2.set_ylim(3.6, 5.6)


# chart 3: sensitivity grid, coloured against guidance
ax3 = fig.add_axes([0.045, 0.05, 0.44, 0.27])
ax3.set_title("FY27 gross profit ($m): core volume growth vs margin squeeze", loc="left", fontsize=11, weight="bold", color=INK, pad=10)
growths = sorted(sens.core_growth.unique())
squeezes = sorted(sens.squeeze_bps.unique())
band_colour = {"Below guidance": "#f8cbad", "Below midpoint": "#ffe699", "At or above midpoint": "#c6efce"}
for i, g in enumerate(growths):
    for j, sq in enumerate(squeezes):
        cell = sens[(sens.core_growth == g) & (sens.squeeze_bps == sq)].iloc[0]
        ax3.add_patch(plt.Rectangle((j, i), 0.96, 0.92, color=band_colour[cell.band]))
        ax3.text(j + 0.48, i + 0.46, f"{cell.fy27_total_gp_m:.1f}", ha="center", va="center", fontsize=8, color=INK)
ax3.set_xlim(0, len(squeezes))
ax3.set_ylim(0, len(growths))
ax3.set_xticks([j + 0.48 for j in range(len(squeezes))])
ax3.set_xticklabels([f"{sq:g}bp" for sq in squeezes])
ax3.set_yticks([i + 0.46 for i in range(len(growths))])
ax3.set_yticklabels([f"{g:.1%}" for g in growths])
ax3.set_xlabel("Margin lost to competition")
ax3.set_ylabel("Core TTV growth")
for side in ["top", "right", "left", "bottom"]:
    ax3.spines[side].set_visible(False)
ax3.tick_params(length=0)


# chart 4: Tyro's merchant fee vs the market average each year
ax4 = fig.add_axes([0.55, 0.05, 0.42, 0.27])
tidy(ax4, "Merchant fee: Tyro vs market average (bps of value)")
xs = range(len(pricing))
w = 0.36
ax4.bar([i - w / 2 - 0.01 for i in xs], pricing.tyro_msf_bps, w, color=SCENARIO_COLOUR["Base"], label="Tyro")
ax4.bar([i + w / 2 + 0.01 for i in xs], pricing.rba_blended_bps, w, color="#a8a7a3", label="Market average (RBA)")
for i, tyro, market in zip(xs, pricing.tyro_msf_bps, pricing.rba_blended_bps):
    ax4.text(i - w / 2, tyro + 1.5, f"{tyro:.0f}", ha="center", fontsize=8.5, color=INK)
    ax4.text(i + w / 2, market + 1.5, f"{market:.0f}", ha="center", fontsize=8.5, color=INK)
ax4.set_xticks(list(xs))
ax4.set_xticklabels(pricing.fy)
ax4.set_ylim(0, 115)
ax4.legend(frameon=False, loc="upper left", ncol=2, fontsize=8.5)

(ROOT / "docs").mkdir(exist_ok=True)
fig.savefig(ROOT / "docs" / "dashboard.png", dpi=110, facecolor=BACKGROUND)
print("saved docs/dashboard.png")
