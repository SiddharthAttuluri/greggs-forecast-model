"""Find the primary financial statements and key notes in each Greggs annual report.

A ~170-page annual report has the numbers spread across a few pages. This script scans
every page's text and lists where the income statement, balance sheet, cash flow
statement, cash-generated-from-operations note and ten-year history sit, so each figure
can be taken from (and later checked against) a specific page.

Usage:
    python scripts/locate_statements.py --pdf-dir sources
"""
import argparse
import pathlib
import subprocess

ROOT = pathlib.Path(__file__).resolve().parents[1]

# (label, phrases that must all appear on the page)
TARGETS = [
    ("Income statement", ["CONSOLIDATED INCOME STATEMENT", "Revenue", "Cost of sales"]),
    ("Balance sheet", ["BALANCE SHEET", "Total assets", "Net assets"]),
    ("Cash flow statement", ["STATEMENTS OF CASH FLOWS", "Operating activities", "Financing activities"]),
    ("Cash generated from operations note", ["Cash generated from operations", "Depreciation", "inventories"]),
    ("Ten-year history", ["TEN-YEAR HISTORY", "Number of shops"]),
    ("Tax reconciliation note", ["Reconciliation of effective tax rate"]),
]


def pages(pdf: pathlib.Path) -> list[str]:
    text = subprocess.run(["pdftotext", "-layout", str(pdf), "-"], capture_output=True, text=True, check=True).stdout
    return text.split("\f")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf-dir", default=str(ROOT / "sources"))
    pdf_dir = pathlib.Path(ap.parse_args().pdf_dir)
    for pdf in sorted(pdf_dir.glob("AR*.pdf")):
        texts = pages(pdf)
        print(f"\n{pdf.name} ({len(texts)} pages)")
        for label, phrases in TARGETS:
            hits = [i + 1 for i, t in enumerate(texts)
                    if all(p.lower() in t.lower() for p in phrases)]
            print(f"  {label:<38} PDF page(s): {', '.join(map(str, hits)) or '-'}")


if __name__ == "__main__":
    main()
