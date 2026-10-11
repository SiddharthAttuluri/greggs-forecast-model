"""Build the Greggs model workbook (Excel) from data/greggs_historicals_fy21_fy25.csv.

Creates model/GRG_model_v0.2.xlsx: Notes, Historical (FY21-FY25, with page references),
Forecast (same rows, FY26-FY28 to be built), Drivers, Scenarios, Variance, Tax & Audit,
Checks and Summary tabs. Formulas are written as Excel formulas, so the file recalculates
when opened in Excel.

Usage:
    python scripts/build_workbook.py
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

F = "Arial"
BLUE = Font(name=F, color="0000FF", size=10)
BLACK = Font(name=F, color="000000", size=10)
GREEN = Font(name=F, color="008000", size=10)
BOLD = Font(name=F, bold=True, size=10)
BOLDBLACK = Font(name=F, bold=True, color="000000", size=10)
TITLE = Font(name=F, bold=True, size=14)
SUB = Font(name=F, italic=True, size=9, color="595959")
HDR = Font(name=F, bold=True, color="FFFFFF", size=10)
HDRFILL = PatternFill("solid", fgColor="1F3864")
SECFILL = PatternFill("solid", fgColor="D9E1F2")
YELLOW = PatternFill("solid", fgColor="FFFF00")
FCFILL = PatternFill("solid", fgColor="F2F2F2")
NUM = '#,##0.0;(#,##0.0);"-"'
INT = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
TOP = Border(top=Side(style="thin"))

wb = Workbook()

# ---------------------------------------------------------------- data (read from the CSV)
import csv, pathlib
ROOT = pathlib.Path(__file__).resolve().parents[1]
DATA = {}
with open(ROOT / "data" / "greggs_historicals_fy21_fy25.csv", newline="") as _f:
    for _r in csv.DictReader(_f):
        _v = float(_r["value"])
        DATA[(_r["key"], _r["year"])] = int(_v) if _r["unit"] == "count" else _v
YEARS = ["FY21", "FY22", "FY23", "FY24", "FY25"]
FCY = ["FY26", "FY27", "FY28"]
ENDS = ["1 Jan 2022", "31 Dec 2022", "30 Dec 2023", "28 Dec 2024", "27 Dec 2025"]
WEEKS = [52, 52, 52, 52, 52]

# each row: (key, label, kind, values or formula-template, source, comment)
# kind: 'sec' section header, 'in' input, 'f' formula (template uses {c} for column and {r[key]} for row refs), 'blank'
R = []
def sec(t): R.append(("", t, "sec", None, "", None))
def inp(k, lab, src, cm=None, fmt=NUM): R.append((k, lab, "in", [DATA.get((k, y)) for y in YEARS], src, cm, fmt))
def frm(k, lab, f, src="Formula", cm=None, fmt=NUM, bold=False): R.append((k, lab, "f", f, src, cm, fmt, bold))
def blank(): R.append(("", "", "blank", None, "", None))

S_IS = "IS: FY24–25 AR25 p130; FY23 AR24 p130; FY22 AR23 p121; FY21 AR22 p110"
S_BS = "BS: FY24–25 AR25 p131; FY23 AR24 p131; FY22 AR23 p122; FY21 AR22 p111"
S_CF = "CF: FY24–25 AR25 p134; FY23 AR24 p134; FY22 AR23 p125; FY21 AR22 p114"
S_CGO = "CF note: FY24–25 AR25 p135; FY23 AR24 p135; FY22 AR23 p126; FY21 AR22 p115"

sec("INCOME STATEMENT (statutory, including exceptional items)")
inp("rev", "Revenue", S_IS)
inp("cos", "Cost of sales", S_IS)
frm("gp", "Gross profit", "={c}{rev}+{c}{cos}", bold=True)
inp("ds", "Distribution and selling costs", S_IS,
    "Statutory total, including exceptional items (FY23 +0.3, FY24 +0.3).")
inp("adm", "Administrative expenses", S_IS,
    "Statutory total. FY25 includes a 3.8 exceptional charge.")
inp("oth", "Other income (exceptional)", S_IS,
    "Exceptional other income: FY23 20.3, FY24 13.8. See Note 4 in each report.")
frm("op", "Operating profit", "={c}{gp}+{c}{ds}+{c}{adm}+{c}{oth}", bold=True)
inp("finc", "Finance income", "FY21 AR22 p131 (note 5); FY22 AR23 p142 (note 6); FY23–25 IS",
    "FY21 and FY22 were reported only as a net finance expense on the face of the IS. Split taken from the finance note: FY22 = interest income 1.3 + FX gain 0.1.")
inp("fexp", "Finance expense", "FY21 AR22 p131 (note 5); FY22 AR23 p142 (note 6); FY23–25 IS",
    "Mostly interest on lease liabilities. FY25 includes a 0.7 exceptional charge.")
frm("pbt", "Profit before tax", "={c}{op}+{c}{finc}+{c}{fexp}", bold=True)
inp("tax", "Income tax", S_IS)
frm("pat", "Profit for the year", "={c}{pbt}+{c}{tax}", bold=True)
blank()
sec("Exceptional items and underlying profit")
inp("exop", "Exceptional items within operating profit", "IS exceptional column",
    "FY23 20.6 and FY24 14.1: mainly exceptional other income; read Note 4 in AR24 and AR25 to see what it was. FY25: charge of 3.8 in admin expenses.")
inp("exfin", "Exceptional items within finance expense", "AR25 p130")
inp("extax", "Tax on exceptional items", "IS exceptional column")
frm("uop", "Underlying operating profit", "={c}{op}-{c}{exop}", bold=True)
frm("upbt", "Underlying profit before tax", "={c}{pbt}-{c}{exop}-{c}{exfin}", bold=True)
frm("uopm", "Underlying operating margin", "=IF({c}{rev}=0,0,{c}{uop}/{c}{rev})", fmt=PCT)
frm("etr", "Effective tax rate (statutory)", "=IF({c}{pbt}=0,0,-{c}{tax}/{c}{pbt})", fmt=PCT)
blank()
sec("Reported totals (typed from the reports, for the checks)")
inp("r_op", "Reported operating profit", S_IS)
inp("r_pbt", "Reported profit before tax", S_IS)
inp("r_pat", "Reported profit for the year", S_IS)
inp("r_upbt", "Reported underlying PBT", "AR25 p176 (ten-year history)")
blank()
sec("REVENUE SPLIT AND SHOP KPIs")
inp("rev_cm", "Revenue: company-managed shops", "FY21 AR22 p126; FY22 AR23 p137; FY23 AR24 p175; FY24–25 AR25 p177")
inp("rev_b2b", "Revenue: business to business (franchise, wholesale, grocery)", "FY21 AR22 p126; FY22 AR23 p137; FY23 AR24 p175; FY24–25 AR25 p177")
frm("rev_chk", "Check: split less revenue (should be 0)", "={c}{rev_cm}+{c}{rev_b2b}-{c}{rev}")
inp("shops_open0", "Shops at start of year", "FY20 closing: AR25 p176",
    "Only FY21's opening figure is typed; later years link to the prior year's closing shops.", fmt=INT)
inp("shops", "Shops at year end (total)", "AR25 p176 (ten-year history)", fmt=INT)
inp("fr", "of which franchised", "FY21 AR21 p18; FY22 AR22 p15; FY23 AR23 p17; FY24 AR24 p33; FY25 not found",
    "FY21: AR21 says 'we currently have ... 375 franchise locations' (written in early 2022), so treat as approximate. FY25: not stated in the report text. Find it in the FY25 results announcement (March 2026) and type it in.", fmt=INT)
frm("cm", "of which company-managed", "=IF({c}{fr}=\"\",\"\",{c}{shops}-{c}{fr})", fmt=INT)
inp("opened", "Gross openings", "FY21 AR21 p18; FY22 AR22 p15; FY23 AR23 p17; FY24 AR24 p22; FY25 not found",
    "Includes franchised openings (FY22: 70; FY23: 72; FY24: 63). FY25 gross openings and closures are not stated in the report text; only 121 net. Find them in the FY25 results announcement.", fmt=INT)
inp("closed", "Closures (incl. relocations)", "FY21 AR21 p18; FY22 AR22 p15; FY23 derived; FY24 AR24 p22",
    "FY23 derived: 220 opened less 145 net = 75. Relocations count as a closure plus an opening.", fmt=INT)
frm("net", "Net openings", "={c}{shops}-{prev}{shops}", fmt=INT, cm=None)
frm("net_chk", "Check: openings + closures less net (should be 0)", "=IF({c}{opened}=\"\",0,{c}{opened}+{c}{closed}-{c}{net})", fmt=INT)
inp("lfl", "Company-managed like-for-like sales growth", "AR25 p176 (ten-year history)",
    "FY21 is against Covid-hit 2020, so it is not a useful base. Against 2019, FY21 LFL was -3.3% (AR21 p18).", fmt=PCT)
inp("flfl", "Franchise like-for-like system sales growth", "FY23 AR24 p175; FY24–25 AR25 p177", fmt=PCT)
frm("revshop", "Company-managed revenue per average shop (£m)",
    "=IF({c}{cm}=\"\",\"\",{c}{rev_cm}/AVERAGE({c}{cm},IF({prev}{cm}=\"\",{c}{cm},{prev}{cm})))", fmt='0.000',
    cm="Average of opening and closing company-managed shops. FY21 uses closing only. A simple first driver; refine it in Phase 2.")
blank()
sec("BALANCE SHEET (group)")
inp("intang", "Intangible assets", S_BS)
inp("ppe", "Property, plant and equipment", S_BS)
inp("rou", "Right-of-use assets", S_BS)
inp("pena", "Defined benefit pension asset", S_BS)
frm("nca", "Non-current assets", "=SUM({c}{intang}:{c}{pena})", bold=True)
inp("inv", "Inventories", S_BS)
inp("rec", "Trade and other receivables", S_BS)
inp("ahfs", "Assets held for sale", S_BS)
inp("cta", "Current tax assets", S_BS)
inp("cash", "Cash and cash equivalents", S_BS)
frm("ca", "Current assets", "=SUM({c}{inv}:{c}{cash})", bold=True)
frm("ta", "Total assets", "={c}{nca}+{c}{ca}", bold=True)
inp("pay", "Trade and other payables", S_BS)
inp("ctl", "Current tax liabilities", S_BS)
inp("leasec", "Lease liabilities (current)", S_BS)
inp("provc", "Provisions (current)", S_BS)
frm("cl", "Current liabilities", "=SUM({c}{pay}:{c}{provc})", bold=True)
inp("borr", "Borrowings", S_BS, "First drawn in FY25: 40.0 drawn, 15.0 repaid (AR25 p134).")
inp("opay", "Other payables (non-current)", S_BS)
inp("leasenc", "Lease liabilities (non-current)", S_BS)
inp("dtl", "Deferred tax liability", S_BS)
inp("provnc", "Long-term provisions", S_BS)
inp("penl", "Defined benefit pension liability", S_BS)
frm("ncl", "Non-current liabilities", "=SUM({c}{borr}:{c}{penl})", bold=True)
frm("tl", "Total liabilities", "={c}{cl}+{c}{ncl}", bold=True)
frm("na", "Net assets", "={c}{ta}+{c}{tl}", bold=True)
inp("sc", "Issued capital", S_BS)
inp("sp", "Share premium", S_BS)
inp("crr", "Capital redemption reserve", S_BS)
inp("re", "Retained earnings", S_BS)
frm("eq", "Total equity", "=SUM({c}{sc}:{c}{re})", bold=True)
inp("r_ta", "Reported total assets", S_BS)
inp("r_na", "Reported net assets", S_BS)
blank()
sec("Memo: net debt")
frm("netcash", "Net cash / (debt) excluding leases", "={c}{cash}+{c}{borr}")
frm("leases", "Total lease liabilities", "={c}{leasec}+{c}{leasenc}")
frm("netdebt_l", "Net cash / (debt) including leases", "={c}{netcash}+{c}{leases}")
blank()
sec("CASH FLOW STATEMENT (interest received shown in operating activities, as in FY25)")
inp("cgo", "Cash generated from operations", S_CF)
inp("taxp", "Income tax paid", S_CF)
inp("intr", "Interest received", S_CF,
    "Reported in investing activities up to FY23; moved to operating activities from FY25 (FY24 restated). Shown in operating for every year here, so FY22–23 operating and investing subtotals differ from those reports by this amount. The net change in cash is unaffected.")
inp("intl", "Interest paid on lease liabilities", S_CF)
inp("intb", "Interest paid on borrowings and other charges", S_CF)
frm("cfo", "Net cash from operating activities", "=SUM({c}{cgo}:{c}{intb})", bold=True)
inp("capppe", "Purchase of property, plant and equipment", S_CF)
inp("capint", "Purchase of intangible assets", S_CF)
inp("dispppe", "Proceeds from sale of PPE", S_CF)
inp("dispahfs", "Proceeds from sale of assets held for sale", S_CF)
frm("cfi", "Net cash from investing activities", "=SUM({c}{capppe}:{c}{dispahfs})", bold=True)
inp("bdraw", "Proceeds from borrowing", S_CF)
inp("brep", "Repayment of borrowing", S_CF)
inp("shiss", "Proceeds from issue of share capital", S_CF)
inp("ownsale", "Sale of own shares", S_CF)
inp("ownbuy", "Purchase of own shares", S_CF)
inp("div", "Dividends paid", S_CF,
    "FY22 and FY24 include special dividends (40p per share declared for FY21 and FY23).")
inp("leasep", "Repayment of lease principal", S_CF)
frm("cff", "Net cash from financing activities", "=SUM({c}{bdraw}:{c}{leasep})", bold=True)
frm("dcash", "Net change in cash", "={c}{cfo}+{c}{cfi}+{c}{cff}", bold=True)
inp("cash0", "Cash at start of year", "AR22 p114",
    "Only FY21's opening cash is typed; later years link to the prior year's closing cash.")
frm("cash1", "Cash at end of year", "={c}{cash0}+{c}{dcash}", bold=True)
inp("r_dcash", "Reported net change in cash", S_CF)
blank()
sec("Memo: free cash flow and capex")
frm("capex", "Total capex (PPE + intangibles)", "={c}{capppe}+{c}{capint}")
frm("fcf", "Free cash flow after leases", "={c}{cfo}+{c}{capex}+{c}{leasep}",
    cm="Operating cash flow, less capex, less lease principal. Before dividends and buybacks.")
blank()
sec("Cash generated from operations (reconciliation note)")
inp("n_pat", "Profit for the year", S_CGO)
inp("n_amort", "Amortisation", S_CGO)
inp("n_dppe", "Depreciation: property, plant and equipment", S_CGO)
inp("n_drou", "Depreciation: right-of-use assets", S_CGO)
inp("n_ippe", "Impairment / (reversal): PPE", S_CGO)
inp("n_irou", "Impairment / (reversal): right-of-use assets", S_CGO)
inp("n_disp", "Loss / (profit) on sale of PPE", S_CGO)
inp("n_grant", "Release of government grants", S_CGO)
inp("n_sbp", "Share-based payment expense", S_CGO)
inp("n_finc", "Finance income", S_CGO,
    "FY21 and FY22 notes show only a net finance expense, so finance income is 0 here and the net figure sits in the next row.")
inp("n_fexp", "Finance expense", S_CGO)
inp("n_tax", "Income tax expense", S_CGO)
inp("n_inv", "(Increase) in inventories", S_CGO)
inp("n_rec", "(Increase) / decrease in receivables", S_CGO)
inp("n_pay", "Increase in payables", S_CGO)
inp("n_prov", "Increase / (decrease) in provisions", S_CGO)
inp("n_pen", "Pension contributions above charge", S_CGO)
frm("n_cgo", "Cash generated from operations (from note)", "=SUM({c}{n_pat}:{c}{n_pen})", bold=True)
frm("da", "Memo: total depreciation and amortisation", "={c}{n_amort}+{c}{n_dppe}+{c}{n_drou}")

# ---------------------------------------------------------------- Historical sheet
ws = wb.active
ws.title = "Historical"
ROW0 = 6
rowmap = {}
r = ROW0
for item in R:
    if item[0]:
        rowmap[item[0]] = r
    r += 1

FC = 3  # first year column C
def col(i): return L(FC + i)

ws["A1"] = "Greggs plc – historical financials, FY21–FY25"; ws["A1"].font = TITLE
ws["A2"] = "£m unless stated. Blue = typed from the reports (check each one); black = formula. Page numbers are PDF pages of the annual reports (AR21 = Annual Report 2021, etc.)."; ws["A2"].font = SUB
ws["A3"] = "Status: extracted by Claude on 9 Oct 2026. Not yet checked by Siddharth. Tick 'Checked' on the Checks sheet as you verify each statement."; ws["A3"].font = Font(name=F, italic=True, size=9, color="C00000")
hdr = ["£m", "Source (PDF page)"] + YEARS
for i, h in enumerate(hdr):
    c = ws.cell(row=4, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
    c.alignment = Alignment(horizontal="center" if i > 1 else "left")
ws.cell(row=5, column=1, value="Period end (weeks)").font = SUB
for i in range(5):
    c = ws.cell(row=5, column=FC + i, value=f"{ENDS[i]} ({WEEKS[i]})"); c.font = SUB; c.alignment = Alignment(horizontal="center")

def fill_formula(tmpl, i):
    c = col(i)
    prev = col(i - 1) if i > 0 else None
    out = tmpl
    if "{prev}" in out:
        if prev is None:
            return None
        out = out.replace("{prev}", prev)
    out = out.replace("{c}", c)
    for k, v in rowmap.items():
        out = out.replace("{" + k + "}", str(v))
    return out

r = ROW0
for item in R:
    key, lab, kind = item[0], item[1], item[2]
    a = ws.cell(row=r, column=1, value=lab)
    if kind == "sec":
        a.font = BOLD
        for cc in range(1, FC + 5):
            ws.cell(row=r, column=cc).fill = SECFILL
    elif kind == "in":
        vals, src, cm, fmt = item[3], item[4], item[5], item[6]
        a.font = BLACK
        ws.cell(row=r, column=2, value=src).font = SUB
        for i, v in enumerate(vals):
            c = ws.cell(row=r, column=FC + i)
            if v is None:
                # link openers to prior-year closings
                if key == "shops_open0" and i > 0:
                    c.value = f"={col(i-1)}{rowmap['shops']}"; c.font = BLACK
                elif key == "cash0" and i > 0:
                    c.value = f"={col(i-1)}{rowmap['cash1']}"; c.font = BLACK
                else:
                    c.fill = YELLOW
            else:
                c.value = v; c.font = BLUE
            c.number_format = fmt
        if cm:
            ws.cell(row=r, column=1).comment = Comment(cm, "Claude")
    elif kind == "f":
        tmpl, src, cm, fmt, bold = item[3], item[4], item[5], item[6], item[7]
        a.font = BOLDBLACK if bold else BLACK
        ws.cell(row=r, column=2, value=src).font = SUB
        for i in range(5):
            t = tmpl
            if key == "net":
                t = "={c}{shops}-{c}{shops_open0}"
            f = fill_formula(t, i)
            if key == "revshop" and i == 0:
                f = fill_formula("=IF({c}{cm}=\"\",\"\",{c}{rev_cm}/{c}{cm})", i)
            c = ws.cell(row=r, column=FC + i, value=f)
            c.font = BOLDBLACK if bold else BLACK
            c.number_format = fmt
            if bold:
                c.border = TOP
        if cm:
            ws.cell(row=r, column=1).comment = Comment(cm, "Claude")
    r += 1
LAST_HIST = r - 1
ws.column_dimensions["A"].width = 52
ws.column_dimensions["B"].width = 58
for i in range(5):
    ws.column_dimensions[col(i)].width = 17
ws.freeze_panes = "C6"

# ---------------------------------------------------------------- Forecast sheet (same row layout)
fs = wb.create_sheet("Forecast")
fs["A1"] = "Greggs plc – forecast, FY26–FY28 (to build in Phase 2)"; fs["A1"].font = TITLE
fs["A2"] = "Same rows as the Historical sheet. FY21–FY25 link to Historical (green). Grey FY26–FY28 cells are yours to build from the Drivers sheet; no typed numbers inside formulas."; fs["A2"].font = SUB
fs["A3"] = "FY26 confirmed as a 52-week year to 26 Dec 2026 (Notes, Q3): no extra-week adjustment needed."; fs["A3"].font = SUB
hdr = ["£m", "How to forecast (Phase 2)"] + YEARS + FCY
for i, h in enumerate(hdr):
    c = fs.cell(row=4, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
    c.alignment = Alignment(horizontal="center" if i > 1 else "left")
hints = {
    "rev_cm": "Shop roll-forward x sales per shop, grown by LFL",
    "rev_b2b": "Growth rate on franchise and wholesale",
    "rev": "Sum of company-managed and B2B",
    "uop": "Revenue x underlying operating margin",
    "shops": "Opening shops + net openings",
    "ppe": "Opening + capex - depreciation (schedule)",
    "rou": "Grows with shop numbers (schedule)",
    "leasec": "Lease schedule",
    "leasenc": "Lease schedule",
    "inv": "Days of cost of sales",
    "rec": "Days of revenue",
    "pay": "Days of cost of sales",
    "cash": "From cash flow, never typed",
    "re": "Opening + profit - dividends - buybacks",
    "cfo": "Built from IS and balance sheet movements",
    "div": "Payout ratio x prior-year profit",
    "tax": "PBT x tax rate",
    "exop": "Restructuring cost (c.£60m cash, FY26–28) on its own line",
}
for item in R:
    key, lab, kind = item[0], item[1], item[2]
    rr = rowmap.get(key)
    if kind == "sec" or not key:
        continue
    fs.cell(row=rr, column=1, value=lab).font = BOLDBLACK if (kind == "f" and item[7]) else BLACK
    if key in hints:
        fs.cell(row=rr, column=2, value=hints[key]).font = SUB
    fmt = item[6] if kind == "in" else item[6]
    for i in range(5):
        c = fs.cell(row=rr, column=FC + i, value=f"=Historical!{col(i)}{rr}")
        c.font = GREEN; c.number_format = fmt
    for j in range(3):
        c = fs.cell(row=rr, column=FC + 5 + j); c.fill = FCFILL; c.number_format = fmt
# section headers on forecast
rr = ROW0
for item in R:
    if item[2] == "sec":
        fs.cell(row=rr, column=1, value=item[1]).font = BOLD
        for cc in range(1, FC + 8):
            fs.cell(row=rr, column=cc).fill = SECFILL
    rr += 1
fs.column_dimensions["A"].width = 52
fs.column_dimensions["B"].width = 44
for i in range(8):
    fs.column_dimensions[L(FC + i)].width = 12
fs.freeze_panes = "C6"

# ---------------------------------------------------------------- Drivers sheet
ds = wb.create_sheet("Drivers")
ds["A1"] = "Drivers – historical ratios and forecast assumptions"; ds["A1"].font = TITLE
ds["A2"] = "Phase 2, step 1. Calculate each historical ratio yourself with formulas pointing at the Historical sheet, then set FY26–FY28 values (blue) and log the reason in Notes."; ds["A2"].font = SUB
hdr = ["Driver", "Unit"] + YEARS + FCY + ["Reason / source (log in Notes too)"]
for i, h in enumerate(hdr):
    c = ds.cell(row=4, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
drivers = [
    ("Estate", None),
    ("Gross openings", "shops"), ("Closures", "shops"), ("Net openings", "shops"),
    ("Franchised share of net openings", "%"),
    ("Revenue", None),
    ("Company-managed like-for-like growth", "%"),
    ("Sales per new shop in first year (share of a full year)", "%"),
    ("Company-managed revenue per average shop", "£m"),
    ("B2B revenue growth", "%"),
    ("Profitability", None),
    ("Underlying operating margin", "%"),
    ("Cost inflation (LFL)", "%"),
    ("Restructuring cost (exceptional)", "£m"),
    ("Investment and working capital", None),
    ("Capex", "£m"),
    ("Depreciation as % of opening PPE", "%"),
    ("Right-of-use assets per company-managed shop", "£m"),
    ("Inventory days (of cost of sales)", "days"),
    ("Receivable days (of revenue)", "days"),
    ("Payable days (of cost of sales)", "days"),
    ("Tax and returns", None),
    ("UK statutory corporation tax rate", "%"),
    ("Recurring tax adjustment (percentage points, from Tax & Audit reconciliation)", "%"),
    ("Effective tax rate (statutory rate + adjustment)", "%"),
    ("Dividend per share / payout ratio", "%"),
    ("Share buybacks", "£m"),
]
rr = 5
for name, unit in drivers:
    if unit is None:
        ds.cell(row=rr, column=1, value=name).font = BOLD
        for cc in range(1, 12):
            ds.cell(row=rr, column=cc).fill = SECFILL
    else:
        ds.cell(row=rr, column=1, value=name).font = BLACK
        ds.cell(row=rr, column=2, value=unit).font = SUB
        for j in range(3):
            ds.cell(row=rr, column=FC + 5 + j).fill = YELLOW
    rr += 1
rr += 1
ds.cell(row=rr, column=1, value="Latest company guidance and H1 2026 data (context for your FY26 assumptions)").font = BOLD
for cc in range(1, 12):
    ds.cell(row=rr, column=cc).fill = SECFILL
rr += 1
guide = [
    ("H1 2026 sales (£m), 26 weeks to 27 Jun 2026", 1101.5, "H1 2026 presentation, slides 3 and 6", NUM),
    ("H1 2025 sales (£m)", 1027.7, "H1 2026 presentation, slide 6", NUM),
    ("H1 2026 operating profit (£m)", 86.5, "slide 6", NUM),
    ("H1 2025 operating profit (£m)", 70.4, "slide 6", NUM),
    ("H1 2026 profit before tax (£m)", 76.0, "slide 6", NUM),
    ("H1 2026 company-managed LFL growth", 0.021, "slide 3", PCT),
    ("Shops trading at 27 Jun 2026", 2773, "slide 18", INT),
    ("FY26 net openings guidance (low)", 100, "slide 18: 'expect 100–110 net shop openings in 2026'", INT),
    ("FY26 net openings guidance (high)", 110, "slide 18", INT),
    ("FY26 capex guidance (£m)", 180, "slide 11: 'reduced from £200m to £180m'", NUM),
    ("FY26 LFL cost inflation expectation", 0.02, "slide 10: 'full-year expectation now lower at c.2%'", PCT),
    ("H2 2026 extra Derby site costs vs H2 2025 (£m)", 10, "slide 9: 'H2 costs will step up by c.£10m'", NUM),
]
for lab, v, src, fmt in guide:
    ds.cell(row=rr, column=1, value=lab).font = BLACK
    c = ds.cell(row=rr, column=FC, value=v); c.font = BLUE; c.number_format = fmt
    ds.cell(row=rr, column=FC + 8, value=src).font = SUB
    rr += 1
ds.cell(row=rr + 1, column=1, value="Not in your files, from the September 2026 Q3 trading update (verify on Greggs' investor site): Q3 LFL c.3.4%; proposed supply-chain restructuring, c.740 roles, c.£60m cash cost, c.£20m annual savings by FY28–29.").font = SUB
ds.column_dimensions["A"].width = 55
ds.column_dimensions["B"].width = 8
for i in range(8):
    ds.column_dimensions[L(FC + i)].width = 11
ds.column_dimensions[L(FC + 8)].width = 55
ds.freeze_panes = "C5"

# ---------------------------------------------------------------- Scenarios
sc = wb.create_sheet("Scenarios")
sc["A1"] = "Scenarios"; sc["A1"].font = TITLE
sc["A2"] = "Phase 3. Fill the yellow cells; the Selected column feeds the Drivers sheet once you link it."; sc["A2"].font = SUB
sc["A4"] = "Scenario switch (1 = base, 2 = upside, 3 = downside)"; sc["A4"].font = BOLD
sc["B4"] = 1; sc["B4"].font = BLUE; sc["B4"].fill = YELLOW
sc["C4"] = '=CHOOSE(B4,"Base","Upside","Downside")'; sc["C4"].font = BLACK
for i, h in enumerate(["Driver (FY26)", "Base", "Upside", "Downside", "Selected", "Why these values"]):
    c = sc.cell(row=6, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
for k, (lab, fmt) in enumerate([("Company-managed like-for-like growth", PCT), ("Net openings (shops)", INT),
                                ("Underlying operating margin", PCT), ("Cost inflation (LFL)", PCT)]):
    rr = 7 + k
    sc.cell(row=rr, column=1, value=lab).font = BLACK
    for j in range(3):
        c = sc.cell(row=rr, column=2 + j); c.fill = YELLOW; c.number_format = fmt
    c = sc.cell(row=rr, column=5, value=f"=INDEX(B{rr}:D{rr},1,$B$4)"); c.font = BLACK; c.number_format = fmt
sc.column_dimensions["A"].width = 50
for cl in "BCDE":
    sc.column_dimensions[cl].width = 12
sc.column_dimensions["F"].width = 50

# ---------------------------------------------------------------- Variance
vs = wb.create_sheet("Variance")
vs["A1"] = "Variance analysis"; vs["A1"].font = TITLE
vs["A2"] = "Test 1 (Phase 3): forecast FY25 using data up to FY24 only, then compare with actual FY25. Test 2 (Phase 4): locked FY26 forecast vs the January trading update and March results."; vs["A2"].font = SUB
vs["A4"] = "Test 1: FY25 back-test"; vs["A4"].font = BOLD
for i, h in enumerate(["£m", "Your FY25 forecast", "Actual FY25", "Variance", "Variance %", "Explanation (one sentence each)"]):
    c = vs.cell(row=5, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
items = [("Revenue: company-managed shops", "rev_cm"), ("Revenue: B2B", "rev_b2b"), ("Total revenue", "rev"),
         ("Underlying operating profit", "uop"), ("Underlying profit before tax", "upbt"),
         ("Net cash from operating activities", "cfo"), ("Capex", "capex"), ("Closing cash", "cash")]
for k, (lab, key) in enumerate(items):
    rr = 6 + k
    vs.cell(row=rr, column=1, value=lab).font = BLACK
    c = vs.cell(row=rr, column=2); c.fill = YELLOW; c.number_format = NUM
    c = vs.cell(row=rr, column=3, value=f"=Historical!G{rowmap[key]}"); c.font = GREEN; c.number_format = NUM
    c = vs.cell(row=rr, column=4, value=f'=IF(B{rr}="","",C{rr}-B{rr})'); c.font = BLACK; c.number_format = NUM
    c = vs.cell(row=rr, column=5, value=f'=IF(OR(B{rr}="",B{rr}=0),"",D{rr}/ABS(B{rr}))'); c.font = BLACK; c.number_format = PCT
vs["A16"] = "Test 2: FY26 live test (lock your forecast before the January 2027 trading update)"; vs["A16"].font = BOLD
for i, h in enumerate(["", "Locked FY26 forecast", "Actual (Jan trading update / March results)", "Variance", "Variance %", "Explanation"]):
    c = vs.cell(row=17, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
for k, lab in enumerate(["Total sales (£m)", "Company-managed LFL growth", "Shops at year end", "Underlying PBT (£m, March)"]):
    rr = 18 + k
    vs.cell(row=rr, column=1, value=lab).font = BLACK
    for j in (2, 3):
        vs.cell(row=rr, column=j).fill = YELLOW
    vs.cell(row=rr, column=4, value=f'=IF(OR(B{rr}="",C{rr}=""),"",C{rr}-B{rr})').font = BLACK
vs.column_dimensions["A"].width = 40
for cl in "BCDE":
    vs.column_dimensions[cl].width = 16
vs.column_dimensions["F"].width = 60

# ---------------------------------------------------------------- Checks
ck = wb.create_sheet("Checks")
ck["A1"] = "Checks – every row should show 0 (green). Red means something does not tie."; ck["A1"].font = TITLE
ck["A2"] = "Historical checks work now. FY26–FY28 checks read 0 while the forecast is empty; they only mean something once you have built it."; ck["A2"].font = SUB
for i, h in enumerate(["Check", "Sheet"] + YEARS + FCY):
    c = ck.cell(row=4, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
checks = [
    ("Operating profit ties to report", "H", "={c}{op}-{c}{r_op}"),
    ("Profit before tax ties to report", "H", "={c}{pbt}-{c}{r_pbt}"),
    ("Profit for the year ties to report", "H", "={c}{pat}-{c}{r_pat}"),
    ("Underlying PBT ties to ten-year history", "H", "={c}{upbt}-{c}{r_upbt}"),
    ("Revenue split ties to total revenue", "H", "={c}{rev_cm}+{c}{rev_b2b}-{c}{rev}"),
    ("Total assets tie to report", "H", "={c}{ta}-{c}{r_ta}"),
    ("Net assets tie to report", "H", "={c}{na}-{c}{r_na}"),
    ("Balance sheet balances (net assets less equity)", "HF", "={c}{na}-{c}{eq}"),
    ("Cash flow closing cash equals balance sheet cash", "HF", "={c}{cash1}-{c}{cash}"),
    ("Net change in cash ties to report", "H", "={c}{dcash}-{c}{r_dcash}"),
    ("Cash generated from operations: note ties to statement", "H", "={c}{n_cgo}-{c}{cgo}"),
    ("Profit in cash flow note equals income statement", "H", "={c}{n_pat}-{c}{pat}"),
    ("Shop openings less closures equal net openings", "H", "={c}{net_chk}"),
]
for k, (lab, kind, tmpl) in enumerate(checks):
    rr = 5 + k
    ck.cell(row=rr, column=1, value=lab).font = BLACK
    ck.cell(row=rr, column=2, value="Historical" if kind == "H" else "Historical + Forecast").font = SUB
    ncols = 5 if kind == "H" else 8
    for i in range(ncols):
        f = fill_formula(tmpl, i)
        sheet = "Historical" if i < 5 else "Forecast"
        f = "=ROUND(" + f[1:].replace(L(FC + i), "%s!%s" % (sheet, L(FC + i))) + ",1)"
        c = ck.cell(row=rr, column=FC + i, value=f); c.font = BLACK; c.number_format = '0.0;(0.0);0.0'
last = 5 + len(checks) - 1
ck.cell(row=last + 2, column=1, value="Total of absolute differences (historical)").font = BOLD
c = ck.cell(row=last + 2, column=FC, value=f"=SUMPRODUCT(ABS(C5:G{last}))"); c.font = BOLDBLACK; c.number_format = '0.0'
ck.cell(row=last + 3, column=1, value="Overall status (historical)").font = BOLD
c = ck.cell(row=last + 3, column=FC, value=f'=IF(C{last+2}<0.05,"OK","CHECK")'); c.font = BOLDBLACK
from openpyxl.formatting.rule import CellIsRule, FormulaRule
red = PatternFill("solid", fgColor="F8CBAD"); grn = PatternFill("solid", fgColor="C6EFCE")
rng = f"C5:J{last}"
ck.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(C5<>"",ABS(C5)>=0.05)'], fill=red))
ck.conditional_formatting.add(rng, FormulaRule(formula=[f'AND(C5<>"",ABS(C5)<0.05)'], fill=grn))
ck.conditional_formatting.add(f"C{last+3}", FormulaRule(formula=[f'C{last+3}="OK"'], fill=grn))
ck.conditional_formatting.add(f"C{last+3}", FormulaRule(formula=[f'C{last+3}<>"OK"'], fill=red))
# verification log
vr = last + 11
ck.cell(row=vr, column=1, value="Your verification log (Phase 1): tick when you have checked each block against the PDF").font = BOLD
for i, h in enumerate(["Block", "Checked by you (Y/N)", "Date", "Notes"]):
    c = ck.cell(row=vr + 1, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
for k, b in enumerate(["Income statement FY21–FY25", "Exceptional items", "Revenue split and shop KPIs",
                       "Balance sheet FY21–FY25", "Cash flow FY21–FY25", "Cash generated from operations note"]):
    ck.cell(row=vr + 2 + k, column=1, value=b).font = BLACK
    ck.cell(row=vr + 2 + k, column=2).fill = YELLOW
ck.column_dimensions["A"].width = 52
ck.column_dimensions["B"].width = 22
for i in range(8):
    ck.column_dimensions[L(FC + i)].width = 10

# ---------------------------------------------------------------- Summary
sm = wb.create_sheet("Summary")
sm["A1"] = "Summary (Phase 3) – one page you could show an interviewer"; sm["A1"].font = TITLE
blocks = ["1. What Greggs does and why I chose it (2–3 lines)",
          "2. Key assumptions (base case) and why",
          "3. Base case outputs: revenue, underlying operating profit, PBT, free cash flow, FY26–FY28",
          "4. Scenario range: upside and downside FY26 PBT",
          "5. Sensitivities: +/-1 point LFL growth and +/-1 point operating margin on FY26 PBT",
          "6. Back-test result: FY25 forecast vs actual, and what I learned",
          "7. Main simplifications and where the model would break",
          "8. Tax and audit points: why the tax rate rose, deferred tax, key audit matters and how the model handles them"]
rr = 3
for b in blocks:
    sm.cell(row=rr, column=1, value=b).font = BOLD
    rr += 4
sm.column_dimensions["A"].width = 100

# ---------------------------------------------------------------- Notes (first tab)
ns = wb.create_sheet("Notes", 0)
ns["A1"] = "Greggs three-statement forecast model – Notes"; ns["A1"].font = TITLE
ns["A2"] = "GRG_model_v0.2 · v0.1 shell and historicals by Claude 9 Oct 2026; v0.2 adds the Tax & Audit tab (10 Oct 2026) · save a new version each session"; ns["A2"].font = SUB
ns["A4"] = "NEXT STEP"; ns["A4"].font = BOLD; ns["A4"].fill = YELLOW
ns["B4"] = "Phase 1: check the Historical sheet block by block against the PDFs (pages are given in column B), and tick each block on the Checks sheet."; ns["B4"].font = BLUE; ns["B4"].fill = YELLOW
r0 = 6
ns.cell(row=r0, column=1, value="Colour key").font = BOLD
key = [("Blue text", "Typed input from a report or your own assumption", BLUE, None),
       ("Black text", "Formula on the same sheet", BLACK, None),
       ("Green text", "Link to another sheet", GREEN, None),
       ("Yellow fill", "Cell you still need to fill in", BLACK, YELLOW),
       ("Grey fill", "Forecast cell to build in Phase 2", BLACK, FCFILL)]
for k, (a, b, fnt, fl) in enumerate(key):
    c = ns.cell(row=r0 + 1 + k, column=1, value=a); c.font = fnt
    if fl: c.fill = fl
    ns.cell(row=r0 + 1 + k, column=2, value=b).font = BLACK
r1 = r0 + 8
ns.cell(row=r1, column=1, value="Tabs").font = BOLD
tabs = [("Historical", "FY21–FY25 statements, revenue split, shop KPIs, with page references. Filled."),
        ("Forecast", "Same rows; FY21–FY25 linked, FY26–FY28 for you to build (Phase 2)."),
        ("Drivers", "Driver list (empty for you to calculate) plus H1 2026 data and guidance."),
        ("Scenarios", "Switch cell and base/upside/downside table (Phase 3)."),
        ("Variance", "FY25 back-test and FY26 live-test layouts; actual FY25 already linked."),
        ("Tax & Audit", "Phase 3: tax rate reconciliation, deferred tax roll-forward, key audit matters map, model controls."),
        ("Checks", "Ties every historical total to the reports; balance and cash checks for the forecast."),
        ("Summary", "Headings for the one-page summary (Phase 3).")]
r2_shift = 1
for k, (a, b) in enumerate(tabs):
    ns.cell(row=r1 + 1 + k, column=1, value=a).font = BLACK
    ns.cell(row=r1 + 1 + k, column=2, value=b).font = BLACK
r2 = r1 + 11
ns.cell(row=r2, column=1, value="Source files (put these in 01_Sources)").font = BOLD
srcs = [("AR21", "Greggs Annual Report and Accounts 2021 (52 weeks to 1 Jan 2022)"),
        ("AR22", "Annual Report 2022 (52 weeks to 31 Dec 2022): used for FY21 figures"),
        ("AR23", "Annual Report and Accounts 2023 (52 weeks to 30 Dec 2023): used for FY22 figures"),
        ("AR24", "Annual Report and Accounts 2024 (52 weeks to 28 Dec 2024): used for FY23 figures"),
        ("AR25", "Annual Report and Accounts 2025 (52 weeks to 27 Dec 2025): used for FY24 and FY25 figures"),
        ("H1 2026", "Interim results presentation, 29 Jul 2026 (26 weeks to 27 Jun 2026)"),
        ("Rule", "Each year is taken from the following year's report where possible, so restated comparatives are used. Page numbers are PDF pages (printed page = PDF page minus 2).")]
for k, (a, b) in enumerate(srcs):
    ns.cell(row=r2 + 1 + k, column=1, value=a).font = BLACK
    ns.cell(row=r2 + 1 + k, column=2, value=b).font = BLACK
r3 = r2 + 10
ns.cell(row=r3, column=1, value="Simplifications and presentation choices already made").font = BOLD
simp = ["Interest received is shown in operating cash flow for all years (Greggs' FY25 policy). FY22 and FY23 operating/investing subtotals therefore differ from those reports by 1.4 and 6.1; net change in cash is unchanged.",
        "FY21 and FY22 finance income and expense are split using the finance note; the face of the income statement only showed a net figure.",
        "Distribution, admin and other income are statutory totals (including exceptional items). Underlying profit is derived by stripping the exceptional rows.",
        "Parent company figures are ignored; all figures are Group.",
        "FY21 is distorted by Covid (LFL +52.4% vs 2020). Consider basing ratios on FY22–FY25 and saying so."]
for k, s in enumerate(simp):
    ns.cell(row=r3 + 1 + k, column=1, value=f"{k+1}.").font = BLACK
    ns.cell(row=r3 + 1 + k, column=2, value=s).font = BLACK
r4 = r3 + 8
ns.cell(row=r4, column=1, value="Open questions to resolve").font = BOLD
qs = ["FY25 franchised shop count, gross openings and closures: not in the report text. Look in the FY25 results announcement (early March 2026) on Greggs' investor site.",
      "FY21 franchised count (375) was stated as 'currently' in a report published in early 2022, so it is approximate.",
      "ANSWERED (10 Oct 2026): FY26 is a 52-week year to 26 Dec 2026, confirmed by Siddharth. No 53rd-week adjustment.",
      "Date of the January 2027 trading update (for the locked forecast in Phase 4)."]
for k, s in enumerate(qs):
    ns.cell(row=r4 + 1 + k, column=1, value=f"Q{k+1}").font = BLACK
    ns.cell(row=r4 + 1 + k, column=2, value=s).font = BLACK
r5 = r4 + 7
ns.cell(row=r5, column=1, value="Assumptions log (fill in during Phase 2)").font = BOLD
for i, h in enumerate(["Assumption", "Value", "Years", "Reason and source", "Date set"]):
    c = ns.cell(row=r5 + 1, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
ex = ["Example: company-managed LFL growth", 0.025, "FY26", "H1 2026 LFL 2.1%, Q3 c.3.4%; weighted toward recent trend", "Nov 2026"]
for i, v in enumerate(ex):
    c = ns.cell(row=r5 + 2, column=1 + i, value=v); c.font = Font(name=F, italic=True, size=10, color="808080")
ns.cell(row=r5 + 2, column=2).number_format = PCT
ns.cell(row=r5 + 3, column=1, value="(Example row in grey: replace it with your own)").font = SUB
ns.column_dimensions["A"].width = 38
ns.column_dimensions["B"].width = 120
for cl in "CDE":
    ns.column_dimensions[cl].width = 14
for row in ns.iter_rows():
    for c in row:
        c.alignment = Alignment(wrap_text=False, vertical="top")

for s in wb.worksheets:
    s.sheet_view.showGridLines = False
    s.sheet_view.zoomScale = 100

# ---------------------------------------------------------------- Tax & Audit tab
ta = wb.create_sheet("Tax & Audit", wb.sheetnames.index("Checks"))
ta["A1"] = "Tax & Audit review (Phase 3)"; ta["A1"].font = TITLE
ta["A2"] = ("Fill the yellow cells from the tax note and the auditor's report. Charges are negative, as on the Historical sheet. "
            "Green = linked from Historical. Checks at the bottom of the Checks tab turn on once you start filling.")
ta["A2"].font = SUB
TAX_SRC = "Tax note: FY24–25 AR25 p153; FY23 AR24 p152; FY22 AR23 p143; FY21 AR22 p132"

def hdr_row(r, first="£m", second="Source / note"):
    for i, h in enumerate([first, second] + YEARS):
        c = ta.cell(row=r, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
        c.alignment = Alignment(horizontal="center" if i > 1 else "left")

def section(r, text):
    ta.cell(row=r, column=1, value=text).font = BOLD
    for cc in range(1, FC + 5):
        ta.cell(row=r, column=cc).fill = SECFILL

def line(r, label, src, kind, vals=None, tmpl=None, fmt=NUM, bold=False, cm=None):
    a = ta.cell(row=r, column=1, value=label); a.font = BOLDBLACK if bold else BLACK
    if src:
        ta.cell(row=r, column=2, value=src).font = SUB
    for i in range(5):
        c = ta.cell(row=r, column=FC + i)
        if kind == "input":
            if vals and vals[i] is not None:
                c.value = vals[i]; c.font = BLUE
            else:
                c.fill = YELLOW; c.font = BLUE
        elif kind in ("link", "calc"):
            f = tmpl(i)
            c.value = f
            c.font = GREEN if kind == "link" else (BOLDBLACK if bold else BLACK)
            if bold:
                c.border = TOP
        c.number_format = fmt
    if cm:
        a.comment = Comment(cm, "Claude")

H = rowmap
# ------------------------------------------------ A. Tax rate reconciliation
r = 4
section(r, "A. Tax rate reconciliation: why the effective rate differs from the UK rate")
hdr_row(r + 1)
R_PBT, R_RATE, R_EXP = r + 2, r + 3, r + 4
line(R_PBT, "Profit before tax", "Historical", "link", tmpl=lambda i: f"=Historical!{col(i)}{H['pbt']}")
line(R_RATE, "UK statutory corporation tax rate", "Tax note (FY23 note: 23.5%)", "input",
     vals=[0.19, 0.19, 0.235, 0.25, 0.25], fmt=PCT,
     cm="Main rate rose from 19% to 25% on 1 April 2023, so FY23 uses a blended 23.5% (AR24 p152 states 23.5%). Confirm FY21 and FY22 against their notes.")
line(R_EXP, "Tax at the statutory rate", "Formula", "calc", tmpl=lambda i: f"=-{col(i)}{R_PBT}*{col(i)}{R_RATE}")
ADJ = ["Items not taxable", "Non-tax-deductible depreciation", "Impact of change in deferred tax rate",
       "Adjustment for prior years", "Other (name it in a comment)"]
R_ADJ0 = r + 5
for k, a in enumerate(ADJ):
    line(R_ADJ0 + k, "  " + a, TAX_SRC if k == 0 else "", "input",
         cm=("Line names follow Greggs' reconciliation in AR24 p152. Enter as charges (negative) or credits (positive), same sign as the income tax line." if k == 0 else None))
R_ADJ1 = R_ADJ0 + len(ADJ) - 1
R_TOT = R_ADJ1 + 1
line(R_TOT, "Tax charge per reconciliation", "Formula", "calc",
     tmpl=lambda i: f"=SUM({col(i)}{R_EXP}:{col(i)}{R_ADJ1})", bold=True)
R_IS = R_TOT + 1
line(R_IS, "Tax charge per income statement", "Historical", "link", tmpl=lambda i: f"=Historical!{col(i)}{H['tax']}")
R_DIFF = R_IS + 1
line(R_DIFF, "Difference (should be 0 once filled)", "Formula", "calc",
     tmpl=lambda i: f"={col(i)}{R_TOT}-{col(i)}{R_IS}")
R_ETR = R_DIFF + 1
line(R_ETR, "Effective tax rate", "Formula", "calc",
     tmpl=lambda i: f"=IF({col(i)}{R_PBT}=0,0,-{col(i)}{R_IS}/{col(i)}{R_PBT})", fmt=PCT)
R_GAP = R_ETR + 1
line(R_GAP, "Gap to statutory rate (percentage points)", "Formula: feeds the Drivers adjustment row", "calc",
     tmpl=lambda i: f"={col(i)}{R_ETR}-{col(i)}{R_RATE}", fmt=PCT)

# ------------------------------------------------ B. Current vs deferred
r = R_GAP + 2
section(r, "B. Split of the tax charge: current vs deferred")
hdr_row(r + 1)
R_CUR, R_DEF = r + 2, r + 3
line(R_CUR, "Current tax charge", TAX_SRC, "input")
line(R_DEF, "Deferred tax charge", TAX_SRC, "input",
     cm="Deferred tax is tax charged now but payable later, mainly because capital allowances on new shops and sites are faster than accounting depreciation.")
R_SPL = r + 4
line(R_SPL, "Total (should equal the income statement charge)", "Formula", "calc",
     tmpl=lambda i: f"={col(i)}{R_CUR}+{col(i)}{R_DEF}", bold=True)
R_SPLD = r + 5
line(R_SPLD, "Difference", "Formula", "calc", tmpl=lambda i: f"={col(i)}{R_SPL}-Historical!{col(i)}{H['tax']}")
R_CASH = r + 6
line(R_CASH, "Memo: tax actually paid in cash", "Historical (cash flow)", "link",
     tmpl=lambda i: f"=Historical!{col(i)}{H['taxp']}")

# ------------------------------------------------ C. Deferred tax roll-forward
r = R_CASH + 2
section(r, "C. Deferred tax liability roll-forward (shown as positive balances)")
hdr_row(r + 1)
R_OPEN = r + 2
def open_f(i):
    return None if i == 0 else f"={col(i-1)}{R_OPEN + 4}"
a = ta.cell(row=R_OPEN, column=1, value="Opening deferred tax liability"); a.font = BLACK
ta.cell(row=R_OPEN, column=2, value="FY21: prior-year balance from AR22 p111 (FY20 column of AR21)").font = SUB
c = ta.cell(row=R_OPEN, column=FC, value=None); c.fill = YELLOW; c.font = BLUE; c.number_format = NUM
for i in range(1, 5):
    c = ta.cell(row=R_OPEN, column=FC + i, value=f"={col(i-1)}{R_OPEN + 5}"); c.font = BLACK; c.number_format = NUM
R_DIS, R_DOCI, R_DOTH = R_OPEN + 1, R_OPEN + 2, R_OPEN + 3
line(R_DIS, "Charged to the income statement", "Deferred tax note", "input")
line(R_DOCI, "Charged to other comprehensive income", "Deferred tax note", "input")
line(R_DOTH, "Other movements", "Deferred tax note", "input")
R_DCALC = R_OPEN + 4
# fix: opening links should point at the calculated closing row
for i in range(1, 5):
    ta.cell(row=R_OPEN, column=FC + i, value=f"={col(i-1)}{R_DCALC}")
line(R_DCALC, "Closing liability (calculated)", "Formula", "calc",
     tmpl=lambda i: f"=SUM({col(i)}{R_OPEN}:{col(i)}{R_DOTH})", bold=True)
R_DBS = R_DCALC + 1
line(R_DBS, "Closing liability per balance sheet", "Historical", "link",
     tmpl=lambda i: f"=-Historical!{col(i)}{H['dtl']}")
R_DDIFF = R_DBS + 1
line(R_DDIFF, "Difference (should be 0 once filled)", "Formula", "calc",
     tmpl=lambda i: f"={col(i)}{R_DCALC}-{col(i)}{R_DBS}")
R_CAPX = R_DDIFF + 2
line(R_CAPX, "Memo: capex (PPE + intangibles, cash)", "Historical", "link",
     tmpl=lambda i: f"=-Historical!{col(i)}{H['capex']}")
R_DEPR = R_CAPX + 1
line(R_DEPR, "Memo: depreciation of PPE + amortisation", "Historical", "link",
     tmpl=lambda i: f"=Historical!{col(i)}{H['n_dppe']}+Historical!{col(i)}{H['n_amort']}")
R_XS = R_DEPR + 1
line(R_XS, "Capex in excess of depreciation", "Formula", "calc",
     tmpl=lambda i: f"={col(i)}{R_CAPX}-{col(i)}{R_DEPR}", bold=True,
     cm="When capex runs ahead of depreciation, capital allowances usually run ahead too, so the deferred tax liability grows. Compare this row with the deferred tax charge above.")

# ------------------------------------------------ D. Key audit matters map
r = R_XS + 3
section(r, "D. Key audit matters (AR25 independent auditor's report, PDF pp. 123–125) mapped to the model")
heads = ["Key audit matter", "What the auditor did (in brief)", "Model lines affected and how I treated them"]
for i, h in enumerate(heads):
    c = ta.cell(row=r + 1, column=[1, 2, 3][i], value=h); c.font = HDR; c.fill = HDRFILL
for cc in range(4, FC + 5):
    ta.cell(row=r + 1, column=cc).fill = HDRFILL
ta.merge_cells(start_row=r + 1, start_column=3, end_row=r + 1, end_column=FC + 4)
R_K0 = r + 2
for k in range(5):
    rr = R_K0 + k
    for cc in (1, 2, 3):
        ta.cell(row=rr, column=cc).fill = YELLOW
        ta.cell(row=rr, column=cc).alignment = Alignment(wrap_text=True, vertical="top")
    ta.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=FC + 4)
    ta.row_dimensions[rr].height = 45
ta.cell(row=R_K0 + 5, column=1,
        value="Example of the level of detail (replace with your own, from the report): Impairment of shop assets | Tested the value-in-use model and its assumptions | Right-of-use assets and PPE; impairment kept inside underlying costs, see Notes").font = SUB

# ------------------------------------------------ E. Controls in my model
r = R_K0 + 7
section(r, "E. Controls in my model (how I know the numbers are right)")
for k in range(3):
    rr = r + 1 + k
    ta.cell(row=rr, column=1, value=f"Control {k+1}").font = BLACK
    ta.cell(row=rr, column=2).fill = YELLOW
    ta.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=FC + 4)
ta.cell(row=r + 4, column=1,
        value="Ideas: source page for every typed figure; Checks tab ties totals to the reports; balance and cash checks on every forecast year; colour coding of inputs vs formulas.").font = SUB

ta.column_dimensions["A"].width = 52
ta.column_dimensions["B"].width = 50
for i in range(5):
    ta.column_dimensions[col(i)].width = 17
ta.freeze_panes = "C4"
ta.sheet_view.showGridLines = False

# ------------------------------------------------ Phase 3 checks on the Checks tab
tr = last + 5
ck.cell(row=tr, column=1, value="Phase 3 checks (Tax & Audit tab): blank until you start filling that section").font = BOLD
tchecks = [
    ("Tax reconciliation ties to income statement charge", R_DIFF, (R_ADJ0, R_ADJ1)),
    ("Current + deferred tax equals income statement charge", R_SPLD, (R_CUR, R_DEF)),
    ("Deferred tax roll-forward ties to balance sheet", R_DDIFF, (R_DIS, R_DOTH)),
]
for k, (lab, drow, (a0, a1)) in enumerate(tchecks):
    rr = tr + 1 + k
    ck.cell(row=rr, column=1, value=lab).font = BLACK
    ck.cell(row=rr, column=2, value="Tax & Audit").font = SUB
    for i in range(5):
        c = col(i)
        f = f"=IF(COUNT('Tax & Audit'!{c}{a0}:{c}{a1})=0,\"\",ROUND('Tax & Audit'!{c}{drow},1))"
        cell = ck.cell(row=rr, column=FC + i, value=f); cell.font = BLACK; cell.number_format = '0.0;(0.0);0.0'
rng2 = f"C{tr+1}:G{tr+3}"
ck.conditional_formatting.add(rng2, FormulaRule(formula=[f'AND(C{tr+1}<>"",ABS(C{tr+1})>=0.05)'], fill=red))
ck.conditional_formatting.add(rng2, FormulaRule(formula=[f'AND(C{tr+1}<>"",ABS(C{tr+1})<0.05)'], fill=grn))

out = ROOT / "model" / "GRG_model_v0.2.xlsx"
wb.save(out)
print("saved", out)
