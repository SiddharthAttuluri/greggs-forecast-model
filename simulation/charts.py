"""Charts for the Monte Carlo results (PNG, used in the README).

Usage:
    python simulation/charts.py
"""
import csv
import pathlib

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import numpy as np  # noqa: E402

ROOT = pathlib.Path(__file__).resolve().parents[1]
RES = ROOT / "simulation" / "results"
SURFACE, INK, INK2, GRID = "#ffffff", "#0b0b0b", "#52514e", "#e4e3df"
BLUE, RED = "#2a78d6", "#e34948"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": GRID,
                     "axes.labelcolor": INK2, "xtick.color": INK2, "ytick.color": INK2,
                     "figure.facecolor": SURFACE, "axes.facecolor": SURFACE})


def histogram():
    runs = np.loadtxt(RES / "fy26_upbt_runs.csv", skiprows=1)
    p5, p50, p95 = np.percentile(runs, [5, 50, 95])
    fig, ax = plt.subplots(figsize=(8, 4.2))
    ax.hist(runs, bins=40, color=BLUE, edgecolor=SURFACE, linewidth=2)
    for v, lab, ha in ((p5, "P5", "right"), (p50, "Median", "left"), (p95, "P95", "left")):
        ax.axvline(v, color=INK2, linewidth=1, linestyle=(0, (3, 3)))
        txt = f"{lab} £{v:.0f}m " if ha == "right" else f" {lab} £{v:.0f}m"
        ax.text(v, ax.get_ylim()[1] * 0.97, txt, color=INK, fontsize=9, va="top", ha=ha)
    ax.axvline(171.9, color=RED, linewidth=1.5)
    ax.text(171.9, ax.get_ylim()[1] * 0.97, " FY25 actual £172m", color=RED, fontsize=9, va="top", ha="left")
    ax.set_title("FY26 underlying profit before tax: 10,000 simulated outcomes", color=INK, loc="left", fontsize=12)
    ax.set_xlabel("£m"); ax.set_ylabel("Number of runs")
    ax.grid(axis="y", color=GRID, linewidth=0.8); ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    fig.tight_layout(); fig.savefig(RES / "fy26_upbt_histogram.png", dpi=160, facecolor="#ffffff"); plt.close(fig)


def tornado():
    rows = list(csv.DictReader(open(RES / "tornado_fy28_upbt.csv")))
    rows = [r for r in rows if abs(float(r["swing"])) >= 1.0][::-1]
    base = 217.6
    labels = [r["driver"] for r in rows]
    lo = [float(r["FY28 uPBT at driver min"]) for r in rows]
    hi = [float(r["FY28 uPBT at driver max"]) for r in rows]
    fig, ax = plt.subplots(figsize=(8, 3.8))
    y = np.arange(len(rows))
    ax.barh(y, [l - base for l in lo], left=base, color=RED, height=0.6, edgecolor=SURFACE, linewidth=2,
            label="Driver at its minimum")
    ax.barh(y, [h - base for h in hi], left=base, color=BLUE, height=0.6, edgecolor=SURFACE, linewidth=2,
            label="Driver at its maximum")
    for yi, l, h in zip(y, lo, hi):
        ax.text(min(l, h) - 1, yi, f"£{min(l, h):.0f}m", va="center", ha="right", fontsize=8, color=INK2)
        ax.text(max(l, h) + 1, yi, f"£{max(l, h):.0f}m", va="center", ha="left", fontsize=8, color=INK2)
    ax.axvline(base, color=INK, linewidth=1)
    ax.set_yticks(y, labels, color=INK)
    ax.set_xlabel("FY28 underlying profit before tax, £m (base case £218m)")
    ax.set_title("What moves FY28 profit most: one driver at a time", color=INK, loc="left", fontsize=12)
    ax.legend(frameon=False, loc="lower right", fontsize=9)
    ax.set_xlim(min(lo + hi) - 15, max(lo + hi) + 15)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    fig.tight_layout(); fig.savefig(RES / "tornado_fy28_upbt.png", dpi=160, facecolor="#ffffff"); plt.close(fig)


if __name__ == "__main__":
    histogram(); tornado(); print("charts saved")
