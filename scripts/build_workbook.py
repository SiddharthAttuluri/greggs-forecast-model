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
inp("fr", "of which franchised", "FY21 AR21 p18; FY22 AR22 p15; FY23 AR23 p17; FY24 AR24 p33; FY25 Q4 trading update 8 Jan 2026",
    "FY21: AR21 says 'we currently have ... 375 franchise locations' (written in early 2022), so treat as approximate. FY25: 602 franchised and 2,137 company-managed, from the Q4 trading update (8 Jan 2026).", fmt=INT)
frm("cm", "of which company-managed", "=IF({c}{fr}=\"\",\"\",{c}{shops}-{c}{fr})", fmt=INT)
inp("opened", "Gross openings", "FY21 AR21 p18; FY22 AR22 p15; FY23 AR23 p17; FY24 AR24 p22; FY25 Q4 trading update 8 Jan 2026",
    "Includes franchised openings (FY22: 70; FY23: 72; FY24: 63). FY25: 207 opened and 86 closed (50 relocations, 36 closures), from the Q4 trading update (8 Jan 2026).", fmt=INT)
inp("closed", "Closures (incl. relocations)", "FY21 AR21 p18; FY22 AR22 p15; FY23 derived; FY24 AR24 p22; FY25 Q4 trading update",
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
# Session B decisions (11 Oct 2026) and defaults from FY22-FY25 history.
ds = wb.create_sheet("Drivers")
ds["A1"] = "Drivers – forecast assumptions (Session B, 11 Oct 2026)"; ds["A1"].font = TITLE
ds["A2"] = ("Blue = assumption. Base case in FY26–FY28 columns; min / max columns are the Monte Carlo ranges "
            "(the base value is the most likely point). Reasons are logged in the project plan.")
ds["A2"].font = SUB
DHDR = ["Driver", "Unit", "FY26", "FY27", "FY28", "FY26 min", "FY26 max", "FY27 min", "FY27 max",
        "FY28 min", "FY28 max", "Basis"]
for i, h in enumerate(DHDR):
    c = ds.cell(row=4, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
DRV = {}      # name -> row
DCOL = ["C", "D", "E"]   # FY26, FY27, FY28 base columns


def dsec(r, text):
    ds.cell(row=r, column=1, value=text).font = BOLD
    for cc in range(1, 13):
        ds.cell(row=r, column=cc).fill = SECFILL


def drow(r, name, label, unit, base, rng=None, basis="", fmt=PCT):
    DRV[name] = r
    ds.cell(row=r, column=1, value=label).font = BLACK
    ds.cell(row=r, column=2, value=unit).font = SUB
    for j, v in enumerate(base):
        c = ds.cell(row=r, column=3 + j, value=v)
        c.font = BLACK if isinstance(v, str) and v.startswith("=") else BLUE
        c.number_format = fmt
    if rng:
        for j, v in enumerate(rng):
            c = ds.cell(row=r, column=6 + j, value=v); c.font = BLUE; c.number_format = fmt
    ds.cell(row=r, column=12, value=basis).font = SUB


dsec(5, "Your decisions (Session B)")
drow(6, "open", "Net shop openings", "shops", [105, 110, 110], [85, 112, 80, 140, 80, 140],
     "Guidance 100–110; H1 2026 only 34, so more downside", INT)
drow(7, "cmshare", "Company-managed share of net openings", "%", [0.60] * 3, [0.55, 0.65] * 3,
     "Rising: 55%, 57%, 60% (FY22–FY24)")
drow(8, "lfl", "Company-managed like-for-like growth", "%", [0.026, 0.03, 0.03], [0.015, 0.035, 0.005, 0.05, 0.005, 0.05],
     "H1 2026 2.1%, Q3 c.3.4%")
drow(9, "newshop", "New-shop sales vs average mature shop", "%", [0.85] * 3, [0.70, 0.95] * 3,
     "Newer sites are smaller formats and ramp up over time")
drow(10, "b2b", "B2B revenue growth", "%", [0.10, 0.08, 0.08], [0.06, 0.14, 0.04, 0.12, 0.04, 0.12],
     "Slowing from 24% to 9%; Tesco Express extension")
drow(11, "margin", "Underlying operating margin", "%", [0.088, 0.090, 0.093], [0.082, 0.094, 0.080, 0.100, 0.083, 0.103],
     "H1 cost ratios improved (partly pay-award timing); H2 Derby +c.£10m; FY28 includes savings")
drow(12, "rtotal", "Restructuring total cost", "£m", [60, None, None], [45, 75],
     "Proposal: c.£60m cash cost (Monte Carlo scales the phasing below)", NUM)
drow(13, "rpl", "Restructuring P&L charge phasing (IAS 37, option A)", "% of total", [0.75, 10 / 60, 5 / 60], None,
     "£45m / £10m / £5m at £60m total")
drow(14, "rcash", "Restructuring cash phasing", "% of total", [10 / 60, 35 / 60, 15 / 60], None,
     "£10m / £35m / £15m at £60m total")
drow(15, "capex", "Capex (PPE + intangibles)", "£m", [180, 165, 160], [170, 195, 140, 200, 140, 200],
     "FY26 guidance £180m; capital intensity falling", NUM)
drow(16, "infl", "Inflation (dividend uplift)", "%", [0.02, 0.02, 0.02], None, "Greggs: c.2% LFL cost inflation")
drow(17, "dps", "Dividend per share declared", "£", ["=0.69*(1+C16)", "=C17*(1+D16)", "=D17*(1+E16)"], None,
     "69p held in real terms", '0.000')
drow(18, "tax", "Tax rate (on profit before tax)", "%", [0.26] * 3, [0.255, 0.27] * 3,
     "Underlying rate 25.7–26.8% vs 25% statutory")

dsec(20, "Defaults (from FY22–FY25 history, approved in Session B)")
drow(21, "gm", "Gross margin", "%", [0.615] * 3, None, "FY22–FY25 average 61.5%")
drow(22, "admpct", "Administrative costs (underlying) % of revenue", "%", [0.047] * 3, None, "FY22–FY25 average 4.7%")
drow(23, "invd", "Inventory days (of cost of sales)", "days", [25] * 3, None, "FY22–FY25 average", NUM)
drow(24, "recd", "Receivable days (of revenue)", "days", [11.5] * 3, None, "FY22–FY25 average", NUM)
drow(25, "payd", "Payable days (of revenue)", "days", [45] * 3, None, "FY22–FY25 average", NUM)
drow(26, "deppct", "PPE depreciation % of opening PPE", "%", [0.135] * 3, None, "Fell from 17.1% to 13.6% as long-life sites added")
drow(27, "intshare", "Intangible capex % of total capex", "%", [0.08] * 3, None, "FY25: 7.7%")
drow(28, "amortpct", "Amortisation % of opening intangibles", "%", [0.18] * 3, None, "FY25: 18.9%")
drow(29, "roudep", "Right-of-use depreciation % of opening ROU assets", "%", [0.17] * 3, None, "FY25: 16.8%")
drow(30, "leaseint", "Lease interest % of opening lease liability", "%", [0.039] * 3, None, "FY24–FY25 c.3.9–4.0%")
drow(31, "cashint", "Interest on cash % of opening cash", "%", [0.03] * 3, None, "FY24 4.1%, FY25 1.4% (lower balances)")
drow(32, "borrint", "Interest on borrowings %", "%", [0.05] * 3, None, "FY25: c.5% on average drawn")
drow(33, "dtlpct", "Deferred tax build-up % of (capex − depreciation and amortisation)", "%", [0.11] * 3, None,
     "FY24–FY25: c.11%")
drow(34, "shares", "Shares in issue (m)", "m", [101.9] * 3, None,
     "Derived: FY25 dividends paid £70.3m ÷ 69p", NUM)
drow(35, "curlease", "Current share of lease liabilities", "%", [0.139] * 3, None, "FY25: 62.5 / 449.8")

dsec(37, "Monte Carlo driver links (Session B)")
drow(38, "c_lfl_m", "Correlation: LFL growth and operating margin", "corr", [0.5, None, None], None, "Moderate", '0.0')
drow(39, "c_lfl_b", "Correlation: LFL growth and B2B growth", "corr", [0.3, None, None], None, "Weak", '0.0')

# historical reference ratios (formulas on Historical)
dsec(41, "Historical reference (calculated from the Historical tab)")
for i, h in enumerate(["Ratio", "", "FY22", "FY23", "FY24", "FY25"]):
    c = ds.cell(row=42, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
H = rowmap
REF = [
    ("Revenue growth", lambda c, p: f"=Historical!{c}{H['rev']}/Historical!{p}{H['rev']}-1", PCT),
    ("Like-for-like growth", lambda c, p: f"=Historical!{c}{H['lfl']}", PCT),
    ("B2B revenue growth", lambda c, p: f"=Historical!{c}{H['rev_b2b']}/Historical!{p}{H['rev_b2b']}-1", PCT),
    ("Gross margin", lambda c, p: f"=1+Historical!{c}{H['cos']}/Historical!{c}{H['rev']}", PCT),
    ("Underlying operating margin", lambda c, p: f"=Historical!{c}{H['uopm']}", PCT),
    ("Inventory days", lambda c, p: f"=-Historical!{c}{H['inv']}/Historical!{c}{H['cos']}*365", NUM),
    ("Receivable days", lambda c, p: f"=Historical!{c}{H['rec']}/Historical!{c}{H['rev']}*365", NUM),
    ("Payable days (of revenue)", lambda c, p: f"=-Historical!{c}{H['pay']}/Historical!{c}{H['rev']}*365", NUM),
    ("PPE depreciation % opening PPE", lambda c, p: f"=Historical!{c}{H['n_dppe']}/Historical!{p}{H['ppe']}", PCT),
    ("Capex £m", lambda c, p: f"=-Historical!{c}{H['capex']}", NUM),
]
for k, (lab, f, fmt) in enumerate(REF):
    rr = 43 + k
    ds.cell(row=rr, column=1, value=lab).font = BLACK
    for j, (c, p) in enumerate([("D", "C"), ("E", "D"), ("F", "E"), ("G", "F")]):
        cell = ds.cell(row=rr, column=3 + j, value=f(c, p)); cell.font = GREEN; cell.number_format = fmt

rr = 43 + len(REF) + 1
dsec(rr, "Latest company guidance and H1 2026 data")
rr += 1
guide = [
    ("H1 2026 sales (£m), 26 weeks to 27 Jun 2026 (H1 2025: 1,027.7)", 1101.5, "H1 2026 presentation, slides 3 and 6", NUM),
    ("H1 2026 operating profit (£m) (H1 2025: 70.4)", 86.5, "slide 6", NUM),
    ("H1 2026 company-managed LFL growth", 0.021, "slide 3", PCT),
    ("Shops trading at 27 Jun 2026", 2773, "slide 18", INT),
    ("FY26 net openings guidance", "100–110", "slide 18", None),
    ("FY26 capex guidance (£m)", 180, "slide 11", NUM),
    ("FY26 LFL cost inflation expectation", 0.02, "slide 10", PCT),
    ("H2 2026 extra Derby site costs (£m)", 10, "slide 9", NUM),
]
for lab, v, src, fmt in guide:
    ds.cell(row=rr, column=1, value=lab).font = BLACK
    c = ds.cell(row=rr, column=3, value=v); c.font = BLUE
    if fmt:
        c.number_format = fmt
    ds.cell(row=rr, column=12, value=src).font = SUB
    rr += 1
ds.column_dimensions["A"].width = 58
ds.column_dimensions["B"].width = 10
for cl in "CDEFGHIJK":
    ds.column_dimensions[cl].width = 10
ds.column_dimensions["L"].width = 70
ds.freeze_panes = "C5"


def D(name, j):
    """Absolute reference to a driver for forecast year j (0 = FY26)."""
    return f"Drivers!${DCOL[j]}${DRV[name]}"


def D0(name):
    return f"Drivers!$C${DRV[name]}"


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
ns["A2"] = "GRG_model_v0.3 · v0.1 shell and historicals (9 Oct 2026); v0.2 Tax & Audit tab (10 Oct); v0.3 Session B assumptions and full forecast except revenue (11 Oct) · save a new version each session"; ns["A2"].font = SUB
ns["A4"] = "NEXT STEP"; ns["A4"].font = BOLD; ns["A4"].fill = YELLOW
ns["B4"] = "Session C: build the revenue section on the Revenue tab (yellow cells). Then check the Checks tab: the balance sheet and cash checks should read 0 for FY26–FY28."; ns["B4"].font = BLUE; ns["B4"].fill = YELLOW
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
        ("Revenue", "Session C: company-managed shop roll-forward and B2B revenue, built by Siddharth."),
        ("Forecast", "Same rows; FY21–FY25 linked; FY26–FY28 formulas driven by Drivers and Revenue, with schedules at the bottom."),
        ("Drivers", "Session B assumptions with Monte Carlo ranges, approved defaults, historical reference ratios, H1 2026 guidance."),
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

# ---------------------------------------------------------------- Revenue build tab (built by Siddharth in Session C)
rv = wb.create_sheet("Revenue", wb.sheetnames.index("Forecast"))
rv["A1"] = "Revenue build – company-managed shops and B2B (FY26–FY28)"; rv["A1"].font = TITLE
for i, h in enumerate(["Line", "How to calculate", "FY25", "FY26", "FY27", "FY28"]):
    c = rv.cell(row=4, column=1 + i, value=h); c.font = HDR; c.fill = HDRFILL
H = rowmap
RV = {}
RVCOL = ["D", "E", "F"]          # FY26..FY28 on this tab
REV_ROWS = [
    # key, label, hint, FY25 value/formula, fmt
    ("sec1", "Company-managed shop roll-forward", None, None, None),
    ("start", "Company-managed shops at start of year", "= last year's closing shops", None, INT),
    ("cmnet", "Company-managed net openings", "= net openings × company-managed share (Drivers)",
     f"=Historical!G{H['cm']}-Historical!F{H['cm']}", INT),
    ("end", "Company-managed shops at end of year", "= start + net openings", f"=Historical!G{H['cm']}", INT),
    ("sec2", "Sales per shop", None, None, None),
    ("lfl", "Like-for-like growth", "= link to Drivers", f"=Historical!G{H['lfl']}", PCT),
    ("sps", "Sales per mature shop (£m)", "= last year × (1 + LFL)", f"=Historical!G{H['revshop']}", '0.000'),
    ("newpct", "New-shop sales vs mature shop", "= link to Drivers", None, PCT),
    ("sec3", "Company-managed revenue", None, None, None),
    ("exist", "Existing estate", "= last year's company-managed revenue × (1 + LFL)", None, NUM),
    ("annual", "Full-year effect of last year's openings", "= last year's net openings × 50% × sales per shop × new-shop %", None, NUM),
    ("newrev", "This year's openings (half a year)", "= this year's net openings × 50% × sales per shop × new-shop %", None, NUM),
    ("cmrev", "Company-managed revenue", "= existing + full-year effect + this year's openings",
     f"=Historical!G{H['rev_cm']}", NUM),
    ("b2b", "B2B revenue", "= last year × (1 + B2B growth from Drivers)", f"=Historical!G{H['rev_b2b']}", NUM),
    ("total", "Total revenue", "= company-managed + B2B", f"=Historical!G{H['rev']}", NUM),
    ("growth", "Revenue growth", "= this year ÷ last year − 1", None, PCT),
    ("sec4", "Checks and sense checks", None, None, None),
    ("chk", "Check: shops at end vs Forecast tab (should be 0)", "Filled for you", None, INT),
]
r = 5
for key, lab, hint, fy25, fmt in REV_ROWS:
    if key.startswith("sec"):
        rv.cell(row=r, column=1, value=lab).font = BOLD
        for cc in range(1, 7):
            rv.cell(row=r, column=cc).fill = SECFILL
        r += 1
        continue
    RV[key] = r
    rv.cell(row=r, column=1, value=lab).font = BOLDBLACK if key in ("cmrev", "total") else BLACK
    if hint:
        rv.cell(row=r, column=2, value=("Calculate: " + hint[1:].strip()) if hint.startswith("=") else hint).font = SUB
    if fy25:
        c = rv.cell(row=r, column=3, value=fy25); c.font = GREEN; c.number_format = fmt
    for j in range(3):
        c = rv.cell(row=r, column=4 + j)
        c.number_format = fmt
        if key != "chk":
            c.fill = YELLOW; c.font = BLACK
    r += 1
rv["A2"] = (f"Session C: fill the yellow cells with formulas. FY25 (green) is the starting point. The Forecast tab reads company-managed and B2B revenue from rows {RV['cmrev']} and {RV['b2b']}.")
rv["A2"].font = SUB
# checks filled in for him
for j in range(3):
    fc = L(FC + 5 + j)
    rv.cell(row=RV["chk"], column=4 + j,
            value=f"=IF({RVCOL[j]}{RV['end']}=\"\",\"\",{RVCOL[j]}{RV['end']}-Forecast!{fc}{H['cm']})").font = BLACK
r = RV["chk"] + 2
rv.cell(row=r, column=1, value="Sense checks for FY26").font = BOLD
rv.cell(row=r + 1, column=1, value="H1 2026 sales growth (26 weeks)").font = BLACK
c = rv.cell(row=r + 1, column=3, value=0.072); c.font = BLUE; c.number_format = PCT
rv.cell(row=r + 1, column=2, value="H1 2026 presentation, slide 6").font = SUB
rv.cell(row=r + 2, column=1, value="Edison FY26 revenue estimate, May 2026 (£m)").font = BLACK
c = rv.cell(row=r + 2, column=3, value=2305.8); c.font = BLUE; c.number_format = NUM
rv.cell(row=r + 2, column=2, value="Third-party estimate: use to sense-check, not as an input").font = SUB
rv.cell(row=r + 3, column=1, value="Your FY26 revenue vs Edison").font = BLACK
c = rv.cell(row=r + 3, column=4, value=f"=IF(D{RV['total']}=\"\",\"\",D{RV['total']}/C{r + 2}-1)")
c.font = BLACK; c.number_format = PCT
rv.column_dimensions["A"].width = 46
rv.column_dimensions["B"].width = 72
for cl in "CDEF":
    rv.column_dimensions[cl].width = 12
rv.freeze_panes = "C5"
rv.sheet_view.showGridLines = False

# ---------------------------------------------------------------- Forecast FY26-FY28 formulas
fs["A2"] = ("FY21–FY25 link to Historical (green). FY26–FY28 are formulas driven by the Drivers and Revenue tabs; "
            "supporting schedules are at the bottom of this tab.")
fs["A3"] = "FY26 is a 52-week year to 26 Dec 2026. Revenue comes from the Revenue tab (Session C)."
fs["A3"].font = SUB

SCH0 = LAST_HIST + 3
fs.cell(row=SCH0 - 1, column=1, value="SUPPORTING SCHEDULES (£m, positive numbers)").font = BOLD
for cc in range(1, FC + 8):
    fs.cell(row=SCH0 - 1, column=cc).fill = SECFILL
SCH_LIST = [
    ("s_capint", "Intangible capex"), ("s_capppe", "PPE capex"), ("s_amort", "Amortisation"),
    ("s_dep", "PPE depreciation"), ("s_leaseT", "Total lease liability"), ("s_roudep", "Right-of-use depreciation"),
    ("s_add", "New leases (non-cash additions)"), ("s_princ", "Lease principal repaid"),
    ("s_leaseint", "Lease interest"), ("s_borrint", "Interest on borrowings"), ("s_cashint", "Interest on cash"),
    ("s_admu", "Underlying administrative costs"), ("s_rpl", "Restructuring P&L charge"),
    ("s_rcash", "Restructuring cash spend"), ("s_dtl", "Deferred tax liability increase"),
    ("s_dps", "Dividend per share declared (£)"), ("s_div", "Dividends paid in cash"),
]
S = {}
for k, (key, lab) in enumerate(SCH_LIST):
    S[key] = SCH0 + k
    fs.cell(row=SCH0 + k, column=1, value=lab).font = BLACK
fs.cell(row=S["s_dps"], column=FC + 4, value=0.69).font = BLUE          # FY25 declared DPS (G)
fs.cell(row=S["s_dps"], column=FC + 4).number_format = '0.000'
fs.cell(row=S["s_dps"], column=2, value="FY25: 69p declared (19p interim + 50p final)").font = SUB
fs.cell(row=S["s_div"], column=2, value="Last year's declared DPS × shares (final paid the following year)").font = SUB
fs.cell(row=S["s_dtl"], column=2, value="Drivers % × (capex − PPE depreciation − amortisation)").font = SUB
fs.cell(row=S["s_leaseT"], column=2, value="Grows with company-managed shops").font = SUB
fs.cell(row=S["s_add"], column=2, value="Balancing: ROU movement + ROU depreciation").font = SUB
fs.cell(row=S["s_princ"], column=2, value="Balancing: opening liability + new leases − closing").font = SUB


def fx(key, j):
    return f"{L(FC + 5 + j)}{H[key]}"


def fp(key, j):
    return f"{L(FC + 4 + j)}{H[key]}"


def sx(key, j):
    return f"{L(FC + 5 + j)}{S[key]}"


def sp(key, j):
    return f"{L(FC + 4 + j)}{S[key]}"


def f_for(key, j):
    """Forecast formula for a main-statement row in forecast year j, or None to leave blank."""
    c = L(FC + 5 + j)
    rvc = RVCOL[j]
    F = {
        "rev_cm": f"=Revenue!{rvc}{RV['cmrev']}",
        "rev_b2b": f"=Revenue!{rvc}{RV['b2b']}",
        "rev": f"={fx('rev_cm', j)}+{fx('rev_b2b', j)}",
        "cos": f"=-{fx('rev', j)}*(1-{D('gm', j)})",
        "ds": f"={D('margin', j)}*{fx('rev', j)}-{fx('rev', j)}-{fx('cos', j)}+{sx('s_admu', j)}",
        "adm": f"=-{sx('s_admu', j)}+{fx('exop', j)}",
        "oth": "=0",
        "finc": f"={sx('s_cashint', j)}",
        "fexp": f"=-{sx('s_leaseint', j)}-{sx('s_borrint', j)}",
        "tax": f"=-{D('tax', j)}*{fx('pbt', j)}",
        "exop": f"=-{sx('s_rpl', j)}",
        "exfin": "=0",
        "extax": f"=-{fx('exop', j)}*{D('tax', j)}",
        "shops_open0": f"={fp('shops', j)}",
        "shops": f"={fx('shops_open0', j)}+{D('open', j)}",
        "fr": f"={fp('fr', j)}+(1-{D('cmshare', j)})*{D('open', j)}",
        "lfl": f"={D('lfl', j)}",
        # balance sheet
        "intang": f"={fp('intang', j)}+{sx('s_capint', j)}-{sx('s_amort', j)}",
        "ppe": f"={fp('ppe', j)}+{sx('s_capppe', j)}-{sx('s_dep', j)}",
        "rou": f"={fp('rou', j)}*{fx('cm', j)}/{fp('cm', j)}",
        "pena": "=0", "ahfs": "=0", "cta": "=0",
        "inv": f"=-{fx('cos', j)}*{D('invd', j)}/365",
        "rec": f"={fx('rev', j)}*{D('recd', j)}/365",
        "cash": f"={fx('cash1', j)}",
        "pay": f"=-{fx('rev', j)}*{D('payd', j)}/365",
        "ctl": f"={fp('ctl', j)}",
        "leasec": f"=-{sx('s_leaseT', j)}*{D('curlease', j)}",
        "leasenc": f"=-{sx('s_leaseT', j)}*(1-{D('curlease', j)})",
        "provc": f"={fp('provc', j)}-({sx('s_rpl', j)}-{sx('s_rcash', j)})",
        "borr": f"={fp('borr', j)}",
        "opay": f"={fp('opay', j)}",
        "dtl": f"={fp('dtl', j)}-{sx('s_dtl', j)}",
        "provnc": f"={fp('provnc', j)}",
        "penl": f"={fp('penl', j)}",
        "sc": f"={fp('sc', j)}", "sp": f"={fp('sp', j)}", "crr": f"={fp('crr', j)}",
        "re": f"={fp('re', j)}+{fx('pat', j)}+{fx('div', j)}",
        # cash flow
        "cgo": (f"={fx('op', j)}+{sx('s_dep', j)}+{sx('s_amort', j)}+{sx('s_roudep', j)}"
                f"-({fx('inv', j)}-{fp('inv', j)})-({fx('rec', j)}-{fp('rec', j)})"
                f"-({fx('pay', j)}-{fp('pay', j)})-({fx('provc', j)}-{fp('provc', j)})"),
        "taxp": f"={fx('tax', j)}-({fx('dtl', j)}-{fp('dtl', j)})",
        "intr": f"={fx('finc', j)}",
        "intl": f"=-{sx('s_leaseint', j)}",
        "intb": f"=-{sx('s_borrint', j)}",
        "capppe": f"=-{sx('s_capppe', j)}",
        "capint": f"=-{sx('s_capint', j)}",
        "dispppe": "=0", "dispahfs": "=0", "bdraw": "=0", "brep": "=0", "shiss": "=0",
        "ownsale": "=0", "ownbuy": "=0",
        "div": f"=-{sx('s_div', j)}",
        "leasep": f"=-{sx('s_princ', j)}",
        "cash0": f"={fp('cash1', j)}",
    }
    return F.get(key)


SCHF = {
    "s_capint": lambda j: f"={D('capex', j)}*{D('intshare', j)}",
    "s_capppe": lambda j: f"={D('capex', j)}-{sx('s_capint', j)}",
    "s_amort": lambda j: f"={fp('intang', j)}*{D('amortpct', j)}",
    "s_dep": lambda j: f"={fp('ppe', j)}*{D('deppct', j)}",
    "s_leaseT": lambda j: f"=-({fp('leasec', j)}+{fp('leasenc', j)})*{fx('cm', j)}/{fp('cm', j)}",
    "s_roudep": lambda j: f"={fp('rou', j)}*{D('roudep', j)}",
    "s_add": lambda j: f"={fx('rou', j)}-{fp('rou', j)}+{sx('s_roudep', j)}",
    "s_princ": lambda j: f"=-({fp('leasec', j)}+{fp('leasenc', j)})+{sx('s_add', j)}-{sx('s_leaseT', j)}",
    "s_leaseint": lambda j: f"=-({fp('leasec', j)}+{fp('leasenc', j)})*{D('leaseint', j)}",
    "s_borrint": lambda j: f"=-{fp('borr', j)}*{D('borrint', j)}",
    "s_cashint": lambda j: f"=MAX({fp('cash', j)},0)*{D('cashint', j)}",
    "s_admu": lambda j: f"={fx('rev', j)}*{D('admpct', j)}",
    "s_rpl": lambda j: f"={D0('rtotal')}*{D('rpl', j)}",
    "s_rcash": lambda j: f"={D0('rtotal')}*{D('rcash', j)}",
    "s_dtl": lambda j: f"={D('dtlpct', j)}*({D('capex', j)}-{sx('s_dep', j)}-{sx('s_amort', j)})",
    "s_dps": lambda j: f"={D('dps', j)}",
    "s_div": lambda j: f"={sp('s_dps', j)}*{D('shares', j)}",
}

for item in R:
    key, kind = item[0], item[2]
    if not key or kind == "sec":
        continue
    rr = H[key]
    fmt = item[6]
    for j in range(3):
        cell = fs.cell(row=rr, column=FC + 5 + j)
        if kind == "f":
            tmpl = item[3]
            if key == "net":
                tmpl = "={c}{shops}-{c}{shops_open0}"
            f = fill_formula(tmpl, 5 + j)
        else:
            f = f_for(key, j)
        if f is None:
            cell.fill = PatternFill(fill_type=None)
            continue
        cell.value = f
        cell.font = BOLDBLACK if (kind == "f" and item[7]) else BLACK
        cell.fill = PatternFill(fill_type=None)
        cell.number_format = fmt
for key, rr in S.items():
    for j in range(3):
        c = fs.cell(row=rr, column=FC + 5 + j, value=SCHF[key](j))
        c.font = BLACK
        c.number_format = '0.000' if key == "s_dps" else NUM

# forecast check rows on Checks tab already reference Forecast for FY26-FY28 (balance and cash)

out = ROOT / "model" / "GRG_model_v0.3.xlsx"
wb.save(out)
print("saved", out)
