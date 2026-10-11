"""Check that every historical figure appears on the annual-report page it cites.

For each row of data/greggs_historicals_fy21_fy25.csv, the script extracts the text of
the cited PDF page (poppler's `pdftotext -layout`) and looks for the figure in the forms
an annual report prints it: 2,151.2 / (829.1) / 52.4% / 2,181.

Limits (stated plainly):
  * It proves the number is printed on the cited page, not that it sits in the right
    column or row. The reconciliation checks (reconcile.py) cover that from the other side.
  * Zero values print as a dash, so they are reported as "not testable".
  * Rows flagged as derived (e.g. a finance split built from two note lines) are not
    printed as a single number; they are reported as "derived" with the working in `note`.

Usage:
    python scripts/verify_sources.py --pdf-dir sources
PDFs are expected as AR21.pdf ... AR25.pdf (see sources/README.md).
"""
import argparse
import csv
import functools
import pathlib
import re
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
DATA = ROOT / "data" / "greggs_historicals_fy21_fy25.csv"


@functools.lru_cache(maxsize=None)
def page_text(pdf: pathlib.Path, page: int) -> str:
    out = subprocess.run(
        ["pdftotext", "-layout", "-f", str(page), "-l", str(page), str(pdf), "-"],
        capture_output=True, text=True, check=True,
    ).stdout
    return out.replace("−", "-")  # normalise unicode minus


def candidates(value: float, unit: str) -> list[str]:
    if unit == "%":
        return [f"{abs(value) * 100:.1f}%"]
    if unit == "count":
        return [f"{abs(int(value)):,}", str(abs(int(value)))]
    return [f"{abs(value):,.1f}"]


def found(text: str, token: str) -> bool:
    # number must not be glued to other digits (so 1.4 does not match 11.4 or 1.45)
    pattern = r"(?<![\d.,])" + re.escape(token) + r"(?![\d])"
    return re.search(pattern, text) is not None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf-dir", default=str(ROOT / "sources"))
    args = ap.parse_args()
    pdf_dir = pathlib.Path(args.pdf_dir)

    counts = {"verified": 0, "not found": 0, "not testable (zero)": 0, "derived": 0}
    problems = []
    with open(DATA, newline="") as f:
        for row in csv.DictReader(f):
            value = float(row["value"])
            if row["note"] and ("derived" in row["note"].lower() or "+" in row["note"] or "less" in row["note"]):
                counts["derived"] += 1
                continue
            if value == 0:
                counts["not testable (zero)"] += 1
                continue
            text = page_text(pdf_dir / f"{row['source_report']}.pdf", int(row["pdf_page"]))
            if any(found(text, t) for t in candidates(value, row["unit"])):
                counts["verified"] += 1
            else:
                counts["not found"] += 1
                problems.append(row)

    for k, v in counts.items():
        print(f"{k:>22}: {v}")
    for p in problems:
        print(f"  NOT FOUND  {p['key']:<12} {p['year']}  {p['value']:>10}  {p['source_report']} p{p['pdf_page']}")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
