"""Reconciliation checks: rebuild every subtotal from the line items and tie it to the reported total.

Where verify_sources.py proves each number is printed on its cited page, this script proves
the numbers fit together: if a figure were typed into the wrong row, a total would stop tying.

Usage:
    python scripts/reconcile.py
"""
import collections
import csv
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
DATA = ROOT / "data" / "greggs_historicals_fy21_fy25.csv"
YEARS = ["FY21", "FY22", "FY23", "FY24", "FY25"]
TOL = 0.05  # £m; reports round to one decimal


def load():
    d = collections.defaultdict(dict)
    with open(DATA, newline="") as f:
        for r in csv.DictReader(f):
            d[r["key"]][r["year"]] = float(r["value"])
    return d


def main() -> int:
    d = load()
    g = lambda k, y: d.get(k, {}).get(y, 0.0)  # noqa: E731  (missing = reported as nil)

    def op(y):
        return sum(g(k, y) for k in ["rev", "cos", "ds", "adm", "oth"])

    def pbt(y):
        return op(y) + g("finc", y) + g("fexp", y)

    def ta(y):
        return sum(g(k, y) for k in ["intang", "ppe", "rou", "pena", "inv", "rec", "ahfs", "cta", "cash"])

    def tl(y):
        return sum(g(k, y) for k in ["pay", "ctl", "leasec", "provc", "borr", "opay", "leasenc", "dtl", "provnc", "penl"])

    def eq(y):
        return sum(g(k, y) for k in ["sc", "sp", "crr", "re"])

    def dcash(y):
        return sum(g(k, y) for k in ["cgo", "taxp", "intr", "intl", "intb", "capppe", "capint", "dispppe",
                                      "dispahfs", "bdraw", "brep", "shiss", "ownsale", "ownbuy", "div", "leasep"])

    def opening_cash(y):
        i = YEARS.index(y)
        return g("cash0", y) if i == 0 else g("cash", YEARS[i - 1])

    def note_cgo(y):
        return sum(v.get(y, 0.0) for k, v in d.items() if k.startswith("n_"))

    def opening_shops(y):
        i = YEARS.index(y)
        return g("shops_open0", y) if i == 0 else g("shops", YEARS[i - 1])

    checks = {
        "Operating profit ties to report": lambda y: op(y) - g("r_op", y),
        "Profit before tax ties to report": lambda y: pbt(y) - g("r_pbt", y),
        "Profit for the year ties to report": lambda y: pbt(y) + g("tax", y) - g("r_pat", y),
        "Underlying PBT ties to ten-year history": lambda y: pbt(y) - g("exop", y) - g("exfin", y) - g("r_upbt", y),
        "Revenue split ties to total revenue": lambda y: g("rev_cm", y) + g("rev_b2b", y) - g("rev", y),
        "Total assets tie to report": lambda y: ta(y) - g("r_ta", y),
        "Net assets tie to report": lambda y: ta(y) + tl(y) - g("r_na", y),
        "Balance sheet balances": lambda y: ta(y) + tl(y) - eq(y),
        "Cash flow closing cash equals balance sheet cash": lambda y: opening_cash(y) + dcash(y) - g("cash", y),
        "Net change in cash ties to report": lambda y: dcash(y) - g("r_dcash", y),
        "Cash generated: note ties to statement": lambda y: note_cgo(y) - g("cgo", y),
        "Profit in cash flow note equals income statement": lambda y: g("n_pat", y) - g("r_pat", y),
        "Shop openings less closures equal net change": lambda y: (
            0.0 if y not in d["opened"] else g("opened", y) + g("closed", y) - (g("shops", y) - opening_shops(y))),
    }

    failures = 0
    print(f"{'Check':<52}" + "".join(f"{y:>8}" for y in YEARS))
    for name, fn in checks.items():
        diffs = [fn(y) for y in YEARS]
        bad = [abs(x) >= TOL for x in diffs]
        failures += sum(bad)
        print(f"{name:<52}" + "".join(f"{round(x, 1) + 0.0:>8.1f}" for x in diffs) + ("   <-- FAIL" if any(bad) else ""))
    print(f"\n{len(checks)} checks x {len(YEARS)} years: {'all tie' if failures == 0 else f'{failures} failures'}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
