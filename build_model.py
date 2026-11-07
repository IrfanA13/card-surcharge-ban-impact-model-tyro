# build_model.py
# Makes my Excel model Tyro_RBA_Reform_Model.xlsx
# Python only writes the layout, labels and formulas. Excel does all the calculating when you open it.
# It reads the clean csv files from build_dataset.py so run that first.

from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter as col_letter

ROOT = Path(__file__).resolve().parents[1]
half = pd.read_csv(ROOT / "data/tyro_halfyear.csv")
vert = pd.read_csv(ROOT / "data/tyro_vertical.csv")
annual = pd.read_csv(ROOT / "data/tyro_annual.csv")


# styles
# I follow the normal finance colours: blue = typed in, black = formula, green = comes from another tab
FONT = "Arial"
BLUE, BLACK, GREEN, WHITE, GREY = "0000FF", "000000", "008000", "FFFFFF", "595959"
YELLOW = PatternFill("solid", fgColor="FFFF00")
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
ACTUAL_FILL = PatternFill("solid", fgColor="F2F2F2")
TOP_LINE = Border(top=Side(style="thin", color="000000"))
MONEY = '#,##0.0;(#,##0.0);"-"'
BPS = '0.0;(0.0);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'


def font(color=BLACK, bold=False, italic=False, size=10):
    return Font(name=FONT, color=color, bold=bold, italic=italic, size=size)


def put(ws, ref, value, fmt=None, color=BLACK, bold=False, italic=False, fill=None, note=None, align=None):
    # small helper so I dont repeat the same styling lines everywhere
    cell = ws[ref]
    cell.value = value
    cell.font = font(color, bold, italic)
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if note:
        cell.comment = Comment(note, "Irfan")
    if align:
        cell.alignment = Alignment(horizontal=align)
    return cell


def header_row(ws, row, labels):
    for i, text in enumerate(labels):
        c = ws.cell(row=row, column=1 + i, value=text)
        c.font = font(WHITE, bold=True)
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center" if i else "left", wrap_text=True)


def sheet_title(ws, title, subtitle=None):
    ws["A1"].value = title
    ws["A1"].font = font(BLACK, bold=True, size=14)
    if subtitle:
        put(ws, "A2", subtitle, color=GREY, italic=True)


wb = Workbook()


# Cover tab
cover = wb.active
cover.title = "Cover"
sheet_title(cover, "Tyro Payments and the RBA card reforms",
            "Can Tyro still hit its FY27 guidance of $240m to $255m gross profit after the Oct 2026 changes?")
cover_lines = [
    ("What I'm testing", "Tyro says the RBA changes won't hurt its gross profit because it will pass the lower interchange "
                         "costs on to merchants. I wanted to see what has to go right for Tyro to hit its FY27 guidance."),
    ("How it flows", "Inputs and Data go into Forecast. Summary, Sensitivity and Budget_vs_Actual all read from Forecast."),
    ("The basic idea", "Payments gross profit is volume (TTV) times margin. The margin starts at the FY26 level, adds any "
                       "interchange saving Tyro keeps and takes off any margin lost to competition."),
    ("Periods", "Tyro's year runs July to June. H1 is Jul to Dec and H2 is Jan to Jun. FY27 is Jul 2026 to Jun 2027. "
                "The reform started on 1 Oct 2026."),
    ("Units", "$m unless it says otherwise. bps means basis points, 1 bp is 0.01% of the money processed."),
    ("Sources", "Tyro annual reports FY24 to FY26, Tyro investor presentations H1 FY24 to FY26, RBA tables C1.1 C2.1 C3 "
                "and the RBA Conclusions Paper from March 2026. Page numbers are on the Data tab."),
]
r = 4
for label, text in cover_lines:
    put(cover, f"A{r}", label, bold=True)
    put(cover, f"B{r}", text).alignment = Alignment(wrap_text=True, vertical="top")
    cover.row_dimensions[r].height = 30
    r += 1
r += 1
put(cover, f"A{r}", "Colours", bold=True)
r += 1
for text, color, fill in [("Blue text is a number I typed in, either history or an assumption", BLUE, None),
                          ("Black text is a formula", BLACK, None),
                          ("Green text pulls from another tab", GREEN, None),
                          ("Yellow cells are the main assumptions, or cells to fill in later", BLACK, YELLOW)]:
    put(cover, f"B{r}", text, color=color, fill=fill)
    r += 1
r += 1
put(cover, f"A{r}", "How to use it", bold=True)
r += 1
for text in ["Change the assumptions on the Inputs tab. The Custom column is a copy of Base for trying things out.",
             "Check Summary to see each scenario against guidance and Sensitivity for the growth vs squeeze grid.",
             "When Tyro reports H1 FY27 (around late Feb 2027) type the real numbers into the yellow cells on Budget_vs_Actual."]:
    put(cover, f"B{r}", text)
    r += 1
cover.column_dimensions["A"].width = 16
cover.column_dimensions["B"].width = 120


# Data tab
# history copied in from my csv files, blue because they are typed in numbers
data_ws = wb.create_sheet("Data")
sheet_title(data_ws, "Tyro history (from my Week 1 dataset)",
            "Blue numbers come from Tyro's reports. The margin columns are formulas. Basis and source say where each row came from.")
header_row(data_ws, 4, ["Half year", "Core TTV ($m)", "Bendigo TTV ($m)", "Total TTV ($m)", "Payments GP ($m)",
                        "Banking GP ($m)", "Corporate & other GP ($m)", "Total GP ($m)", "Payments margin (bps)",
                        "Basis", "Source"])
core_ttv = vert[vert.is_core].groupby("period").ttv_m.sum()
bendigo_ttv = vert[~vert.is_core].groupby("period").ttv_m.sum()
DATA_ROW = {}
for i, (_, h) in enumerate(half.iterrows()):
    row = 5 + i
    p = h.period
    DATA_ROW[p] = row
    put(data_ws, f"A{row}", p, bold=True)
    put(data_ws, f"B{row}", float(core_ttv[p]), MONEY, BLUE)
    put(data_ws, f"C{row}", float(bendigo_ttv[p]), MONEY, BLUE)
    put(data_ws, f"D{row}", float(h.ttv_m), MONEY, BLUE,
        note="Total TTV as Tyro reported it (or full year minus H1). Core plus Bendigo can be off by a little because of rounding.")
    put(data_ws, f"E{row}", float(h.payments_gp_m), MONEY, BLUE)
    put(data_ws, f"F{row}", float(h.banking_gp_m), MONEY, BLUE)
    put(data_ws, f"G{row}", f"=H{row}-E{row}-F{row}", MONEY)
    put(data_ws, f"H{row}", float(h.total_gp_m), MONEY, BLUE)
    put(data_ws, f"I{row}", f"=E{row}/D{row}*10000", BPS)
    put(data_ws, f"J{row}", h.basis, color=GREY)
    put(data_ws, f"K{row}", h.source, color=GREY)

annual_start = 5 + len(half) + 2
put(data_ws, f"A{annual_start}", "Yearly payments numbers", bold=True)
header_row(data_ws, annual_start + 1, ["Year", "TTV ($m)", "Payments revenue ($m)", "Direct costs ($m)", "Payments GP ($m)",
                                       "Banking GP ($m)", "Total GP ($m)", "Revenue (bps)", "Direct costs (bps)",
                                       "Other (bps)", "Payments margin (bps)", "Source"])
ANNUAL_ROW = {}
for i, (_, a) in enumerate(annual.iterrows()):
    row = annual_start + 2 + i
    ANNUAL_ROW[a.fy] = row
    put(data_ws, f"A{row}", a.fy, bold=True)
    values = [a.ttv_m, a.payments_revenue_m, a.payments_direct_exp_m, a.payments_gp_m, a.banking_gp_m, a.total_gp_m]
    for col, val in zip("BCDEFG", values):
        put(data_ws, f"{col}{row}", float(val), MONEY, BLUE)
    put(data_ws, f"H{row}", f"=C{row}/B{row}*10000", BPS)
    put(data_ws, f"I{row}", f"=D{row}/B{row}*10000", BPS)
    put(data_ws, f"J{row}", f"=H{row}-I{row}-K{row}", BPS,
        note="Revenue minus direct costs minus margin. Mostly the Bendigo share that Tyro takes off to get normalised GP.")
    put(data_ws, f"K{row}", f"=E{row}/B{row}*10000", BPS)
    put(data_ws, f"L{row}", a.source, color=GREY)
for col, width in zip("ABCDEFGHIJKL", [11, 13, 14, 13, 14, 13, 14, 12, 13, 34, 50, 60]):
    data_ws.column_dimensions[col].width = width
data_ws.freeze_panes = "B5"
FY26_ROW = ANNUAL_ROW["FY26"]


# Inputs tab
inputs = wb.create_sheet("Inputs")
sheet_title(inputs, "Inputs and assumptions", "Blue cells are inputs and yellow ones are the main levers. Change any blue cell and the rest updates.")
SCENARIOS = ["Bear", "Base", "Bull", "Custom"]
SCEN_COL = {"Bear": "C", "Base": "D", "Bull": "E", "Custom": "F"}
header_row(inputs, 4, ["Scenario assumption", "Unit", "Bear", "Base", "Bull", "Custom", "Why I picked these"])

# key, label, unit, format, bear, base, bull, my reasoning
scenario_inputs = [
    ("core_g", "Core TTV growth (vs same half last year)", "%", PCT, 0.01, 0.035, 0.06,
     "Core TTV grew 5.6% in H1 FY26, 4.4% in FY26 and only 1.7% in FY25. The whole card market grows around 8 to 10% a year so Tyro has been losing share (Week 1 Q1)."),
    ("bend_g", "Bendigo TTV growth", "%", PCT, -0.12, -0.10, -0.08,
     "Bendigo volume has dropped 10 to 12% every year for three years. FY26 was -10.5%."),
    ("uplift", "Extra volume from the surcharge ban", "% of TTV", PCT, 0.0, 0.0, 0.005,
     "Could happen if people use cards more or cafes lift prices instead of surcharging. No proof yet so I only put it in Bull."),
    ("drift", "Underlying margin change (pricing before the reform)", "bps / year", BPS, 0.0, 0.5, 1.0,
     "Margin went up 1.8 bps in FY25 and 0.5 bps in FY26. I think it slows down from here."),
    ("pt", "Share of interchange saving passed to merchants", "%", PCT, 1.0, 1.0, 0.9,
     "Tyro said it will pass the lower interchange on to merchants. It keeps the margin the same but revenue drops."),
    ("squeeze", "Margin lost to competition (full amount)", "bps", BPS, 3.0, 1.0, 0.0,
     "Tyro charges about 97 bps vs a market average of about 62 bps (Week 1 Q4). No surcharging plus fees being published from Apr 2027 makes that gap easy to see."),
    ("bank_g", "Banking GP growth (vs same half last year)", "%", PCT, 0.10, 0.18, 0.25,
     "Banking GP grew 23% in FY26 and 6.7% in FY25."),
    ("corp_g", "Corporate & other GP growth", "%", PCT, 0.0, 0.0, 0.05,
     "Small and jumps around, about $12m a year. Mostly interest on Tyro's own cash."),
]
INPUT_ROW = {}
for i, (key, label, unit, fmt, bear, base, bull, why) in enumerate(scenario_inputs):
    row = 5 + i
    INPUT_ROW[key] = row
    put(inputs, f"A{row}", label)
    put(inputs, f"B{row}", unit, color=GREY)
    highlight = YELLOW if key in ("core_g", "squeeze", "pt") else None
    for col, val in zip("CDEF", [bear, base, bull, base]):   # Custom starts as a copy of Base
        put(inputs, f"{col}{row}", val, fmt, BLUE, fill=highlight)
    put(inputs, f"G{row}", why, color=GREY).alignment = Alignment(wrap_text=True, vertical="top")
    inputs.row_dimensions[row].height = 28

reform_start = 5 + len(scenario_inputs) + 1
put(inputs, f"A{reform_start}", "Interchange reform (same for every scenario)", bold=True)
header_row(inputs, reform_start + 1, ["Item", "Unit", "Value", "", "", "", "Why I picked these"])
reform_inputs = [
    ("credit_share", "Credit cards' share of Tyro TTV", "%", PCT, 0.396,
     "Tyro doesnt publish its card mix so I used the national credit share from H2 FY26 (RBA C1.1 and C2.1, Week 1 Q3)."),
    ("consumer", "Consumer cards' share of credit", "%", PCT, 0.85,
     "My assumption. Business credit cards keep the 0.80% cap so only consumer credit gets cut."),
    ("cr_before", "Consumer credit interchange before", "% of value", PCT2, 0.0055,
     "My assumption. The old cap was 0.50% on average and 0.80% per transaction. Small merchants usually pay near the top."),
    ("cr_after", "Consumer credit interchange after", "% of value", PCT2, 0.0030,
     "New cap of 0.30% from 1 Oct 2026 (RBA Conclusions Paper, Mar 2026)."),
    ("db_before", "Debit interchange before", "% of value", PCT2, 0.0010,
     "My assumption. The old cap was 0.20% or 10c but lots of transactions pay less."),
    ("db_after", "Debit interchange after", "% of value", PCT2, 0.0008,
     "New cap is 0.16% or 8c (RBA Conclusions Paper) so I cut my number by about 20%."),
]
for i, (key, label, unit, fmt, val, why) in enumerate(reform_inputs):
    row = reform_start + 2 + i
    INPUT_ROW[key] = row
    put(inputs, f"A{row}", label)
    put(inputs, f"B{row}", unit, color=GREY)
    put(inputs, f"C{row}", val, fmt, BLUE, fill=YELLOW if key == "cr_before" else None)
    put(inputs, f"G{row}", why, color=GREY).alignment = Alignment(wrap_text=True, vertical="top")
    inputs.row_dimensions[row].height = 28

# total saving per $100 = credit part + debit part
row = reform_start + 2 + len(reform_inputs)
INPUT_ROW["saving"] = row
k = INPUT_ROW
put(inputs, f"A{row}", "Interchange saving on Tyro's TTV", bold=True)
put(inputs, f"B{row}", "bps", color=GREY)
put(inputs, f"C{row}", f"=(C{k['credit_share']}*C{k['consumer']}*(C{k['cr_before']}-C{k['cr_after']})"
                       f"+(1-C{k['credit_share']})*(C{k['db_before']}-C{k['db_after']}))*10000", BPS, bold=True)
put(inputs, f"G{row}", "How much Tyro's costs drop per $ processed once the reform is fully in.", color=GREY)
inputs[f"C{row}"].border = TOP_LINE

timing_start = row + 2
put(inputs, f"A{timing_start}", "Timing (same for every scenario)", bold=True)
FORECAST_PERIODS = ["H1FY27", "H2FY27", "H1FY28", "H2FY28", "H1FY29", "H2FY29"]
header_row(inputs, timing_start + 1, ["Item", "Unit"] + FORECAST_PERIODS + ["Why"])
INPUT_ROW["active"] = timing_start + 2
INPUT_ROW["phase"] = timing_start + 3
put(inputs, f"A{timing_start + 2}", "Reform in effect (share of the half)")
put(inputs, f"B{timing_start + 2}", "%", color=GREY)
put(inputs, f"A{timing_start + 3}", "Competition squeeze phase in")
put(inputs, f"B{timing_start + 3}", "%", color=GREY)
for j, (active, phase) in enumerate(zip([0.5, 1, 1, 1, 1, 1], [0.25, 0.75, 1, 1, 1, 1])):
    put(inputs, f"{col_letter(3 + j)}{timing_start + 2}", active, PCT, BLUE)
    put(inputs, f"{col_letter(3 + j)}{timing_start + 3}", phase, PCT, BLUE)
put(inputs, f"I{timing_start + 2}", "Started 1 Oct 2026 so only 3 of the 6 months in H1 FY27.", color=GREY)
put(inputs, f"I{timing_start + 3}", "Merchants feel the full fees from Oct 2026 and fees get published from 1 Apr 2027.", color=GREY)

base_start = timing_start + 5
put(inputs, f"A{base_start}", "FY26 starting point and FY27 guidance", bold=True)
header_row(inputs, base_start + 1, ["Item", "Unit", "Value", "", "", "", "Source"])
base_inputs = [
    ("cost_bps", "FY26 direct costs", "bps", f"=Data!I{FY26_ROW}", "Data tab (Annual Report 2026)", GREEN),
    ("other_bps", "FY26 other deductions (Bendigo share etc)", "bps", f"=Data!J{FY26_ROW}", "Data tab", GREEN),
    ("g_low", "FY27 guidance low", "$m", 240, "Tyro FY26 results release, 25 Aug 2026", BLUE),
    ("g_high", "FY27 guidance high", "$m", 255, "Tyro FY26 results release, 25 Aug 2026", BLUE),
    ("g_mid", "FY27 guidance midpoint", "$m", None, "Formula", BLACK),
]
for i, (key, label, unit, val, src, color) in enumerate(base_inputs):
    row = base_start + 2 + i
    INPUT_ROW[key] = row
    put(inputs, f"A{row}", label)
    put(inputs, f"B{row}", unit, color=GREY)
    if key == "g_mid":
        val = f"=AVERAGE(C{INPUT_ROW['g_low']},C{INPUT_ROW['g_high']})"
    put(inputs, f"C{row}", val, MONEY if unit == "$m" else BPS, color)
    put(inputs, f"G{row}", src, color=GREY)
for col, width in zip("ABCDEFGHI", [46, 11, 11, 11, 11, 11, 90, 11, 60]):
    inputs.column_dimensions[col].width = width
inputs.freeze_panes = "C5"


def inp(key, scenario=None):
    # gives the locked cell address of an input, like Inputs!$D$5
    col = SCEN_COL[scenario] if scenario else "C"
    return f"Inputs!${col}${INPUT_ROW[key]}"


# Forecast tab
# 4 blocks, one per scenario. Columns C and D are the real FY26 halves, E to J are the forecast.
fc = wb.create_sheet("Forecast")
sheet_title(fc, "Half year forecast for each scenario", "Grey columns are actual numbers from the Data tab. H1 is Jul to Dec, H2 is Jan to Jun.")
COLUMNS = ["H1FY26A", "H2FY26A"] + FORECAST_PERIODS
header_row(fc, 3, ["Line item", "Unit"] + COLUMNS)
for j in range(len(COLUMNS)):
    fc.column_dimensions[col_letter(3 + j)].width = 11
fc.column_dimensions["A"].width = 44
fc.column_dimensions["B"].width = 9
put(fc, "A4", "FY26 revenue and costs aren't split by half in Tyro's reports so those rows start in FY27.", color=GREY, italic=True)

LINES = ["Core TTV", "Bendigo TTV", "Extra volume (surcharge ban)", "Total TTV",
         "Reform in effect", "Squeeze phase in", "Underlying payments margin", "Interchange saving",
         "Saving kept by Tyro", "Margin lost to competition", "Payments margin", "Payments gross profit",
         "Direct costs (rate)", "Payments revenue (rate)", "Payments revenue", "Banking gross profit",
         "Corporate & other gross profit", "Total gross profit"]
UNITS = ["$m", "$m", "$m", "$m", "%", "%", "bps", "bps", "bps", "bps", "bps", "$m", "bps", "bps", "$m", "$m", "$m", "$m"]
FORMATS = [MONEY, MONEY, MONEY, MONEY, PCT, PCT, BPS, BPS, BPS, BPS, BPS, MONEY, BPS, BPS, MONEY, MONEY, MONEY, MONEY]
BOLD_LINES = (4, 11, 12, 18)
BLOCK_SIZE = 21
BLOCK_START = {s: 5 + i * BLOCK_SIZE for i, s in enumerate(SCENARIOS)}


def frow(scenario, line):
    # row number of a line item (1 to 18) inside a scenario block
    return BLOCK_START[scenario] + line


for s in SCENARIOS:
    top = BLOCK_START[s]
    title_cell = fc.cell(row=top, column=1, value=f"{s.upper()} SCENARIO")
    title_cell.font = font(bold=True)
    for j in range(1, 11):
        fc.cell(row=top, column=j).fill = SUB_FILL
    for i, (label, unit) in enumerate(zip(LINES, UNITS), start=1):
        put(fc, f"A{top + i}", label, bold=(i in BOLD_LINES))
        put(fc, f"B{top + i}", unit, color=GREY)

    # actual columns, just links to the Data tab
    for col, period in (("C", "H1FY26"), ("D", "H2FY26")):
        d = DATA_ROW[period]
        actual = {1: f"=Data!B{d}", 2: f"=Data!C{d}", 3: 0, 4: f"=SUM({col}{frow(s, 1)}:{col}{frow(s, 3)})",
                  5: 0, 6: 0, 7: f"=Data!I{d}", 8: 0, 9: 0, 10: 0, 11: f"=Data!I{d}", 12: f"=Data!E{d}",
                  16: f"=Data!F{d}", 17: f"=Data!G{d}",
                  18: f"=SUM({col}{frow(s, 12)},{col}{frow(s, 16)},{col}{frow(s, 17)})"}
        for line, f in actual.items():
            if str(f).startswith("=Data"):
                color = GREEN
            elif f == 0:
                color = BLUE
            else:
                color = BLACK
            put(fc, f"{col}{frow(s, line)}", f, FORMATS[line - 1], color).fill = ACTUAL_FILL
        for line in (13, 14, 15):
            fc[f"{col}{frow(s, line)}"].fill = ACTUAL_FILL

    # forecast columns
    # each half grows from the same half last year (2 columns to the left) because H1 has Christmas in it
    for j, period in enumerate(FORECAST_PERIODS):
        col = col_letter(5 + j)
        last_year = col_letter(3 + j)
        timing_col = col_letter(3 + j)
        formulas = {
            1: f"={last_year}{frow(s, 1)}*(1+{inp('core_g', s)})",
            2: f"={last_year}{frow(s, 2)}*(1+{inp('bend_g', s)})",
            3: f"=({col}{frow(s, 1)}+{col}{frow(s, 2)})*{inp('uplift', s)}*{col}{frow(s, 5)}",
            4: f"=SUM({col}{frow(s, 1)}:{col}{frow(s, 3)})",
            5: f"=Inputs!{timing_col}${INPUT_ROW['active']}",
            6: f"=Inputs!{timing_col}${INPUT_ROW['phase']}",
            7: f"={last_year}{frow(s, 7)}+{inp('drift', s)}",
            8: f"={inp('saving')}*{col}{frow(s, 5)}",
            9: f"={col}{frow(s, 8)}*(1-{inp('pt', s)})",
            10: f"=-{inp('squeeze', s)}*{col}{frow(s, 6)}",
            11: f"={col}{frow(s, 7)}+{col}{frow(s, 9)}+{col}{frow(s, 10)}",
            12: f"={col}{frow(s, 4)}*{col}{frow(s, 11)}/10000",
            13: f"={inp('cost_bps')}-{col}{frow(s, 8)}",
            14: f"={col}{frow(s, 11)}+{col}{frow(s, 13)}+{inp('other_bps')}",
            15: f"={col}{frow(s, 4)}*{col}{frow(s, 14)}/10000",
            16: f"={last_year}{frow(s, 16)}*(1+{inp('bank_g', s)})",
            17: f"={last_year}{frow(s, 17)}*(1+{inp('corp_g', s)})",
            18: f"=SUM({col}{frow(s, 12)},{col}{frow(s, 16)},{col}{frow(s, 17)})",
        }
        for line, f in formulas.items():
            put(fc, f"{col}{frow(s, line)}", f, FORMATS[line - 1],
                GREEN if f.startswith("=Inputs") else BLACK, bold=(line in BOLD_LINES))
    for line in BOLD_LINES:
        for j in range(3, 11):
            fc.cell(row=frow(s, line), column=j).border = TOP_LINE
fc.freeze_panes = "C4"


# Summary tab
summary = wb.create_sheet("Summary")
sheet_title(summary, "Summary for each scenario", "Full year is H1 plus H2. Full year margin is gross profit divided by TTV.")
put(summary, "A4", "FY27 vs guidance", bold=True)
header_row(summary, 5, ["Scenario", "FY27 total GP ($m)", "Guidance low", "Guidance high", "Result", "Gap to midpoint ($m)",
                        "FY27 payments margin (bps)", "FY27 TTV ($bn)", "FY27 payments revenue ($m)", "Revenue change vs FY26"])
for i, s in enumerate(SCENARIOS):
    row = 6 + i
    put(summary, f"A{row}", s, bold=True)
    put(summary, f"B{row}", f"=Forecast!E{frow(s, 18)}+Forecast!F{frow(s, 18)}", MONEY, GREEN, bold=True)
    put(summary, f"C{row}", f"={inp('g_low')}", MONEY, GREEN)
    put(summary, f"D{row}", f"={inp('g_high')}", MONEY, GREEN)
    put(summary, f"E{row}", f'=IF(B{row}<C{row},"Below",IF(B{row}>D{row},"Above","Within"))', bold=True, align="center")
    put(summary, f"F{row}", f"=B{row}-{inp('g_mid')}", MONEY)
    put(summary, f"G{row}", f"=(Forecast!E{frow(s, 12)}+Forecast!F{frow(s, 12)})/(Forecast!E{frow(s, 4)}+Forecast!F{frow(s, 4)})*10000", BPS)
    put(summary, f"H{row}", f"=(Forecast!E{frow(s, 4)}+Forecast!F{frow(s, 4)})/1000", "0.0")
    put(summary, f"I{row}", f"=Forecast!E{frow(s, 15)}+Forecast!F{frow(s, 15)}", MONEY)
    put(summary, f"J{row}", f"=I{row}/Data!C{FY26_ROW}-1", PCT)
red = PatternFill("solid", fgColor="F8CBAD")
green = PatternFill("solid", fgColor="C6EFCE")
amber = PatternFill("solid", fgColor="FFE699")
summary.conditional_formatting.add("E6:E9", CellIsRule(operator="equal", formula=['"Below"'], fill=red))
summary.conditional_formatting.add("E6:E9", CellIsRule(operator="equal", formula=['"Within"'], fill=green))
summary.conditional_formatting.add("E6:E9", CellIsRule(operator="equal", formula=['"Above"'], fill=green))

tables = [("Total gross profit ($m)", 18, MONEY, "total"), ("Payments gross profit ($m)", 12, MONEY, "pgp"),
          ("Payments margin (bps)", 11, BPS, "margin"), ("Total TTV ($bn)", 4, "0.0", "ttv")]
r = 12
for name, line, fmt, kind in tables:
    put(summary, f"A{r}", name, bold=True)
    header_row(summary, r + 1, ["Scenario", "FY26A", "FY27", "FY28", "FY29", "Growth a year FY26 to FY29"])
    for i, s in enumerate(SCENARIOS):
        row = r + 2 + i
        put(summary, f"A{row}", s)
        for j, (h1_col, h2_col) in enumerate([("C", "D"), ("E", "F"), ("G", "H"), ("I", "J")]):
            col = col_letter(2 + j)
            if kind == "margin":
                f = f"=(Forecast!{h1_col}{frow(s, 12)}+Forecast!{h2_col}{frow(s, 12)})/(Forecast!{h1_col}{frow(s, 4)}+Forecast!{h2_col}{frow(s, 4)})*10000"
            elif kind == "ttv":
                f = f"=(Forecast!{h1_col}{frow(s, 4)}+Forecast!{h2_col}{frow(s, 4)})/1000"
            else:
                f = f"=Forecast!{h1_col}{frow(s, line)}+Forecast!{h2_col}{frow(s, line)}"
            put(summary, f"{col}{row}", f, fmt, GREEN)
        if kind == "margin":
            put(summary, f"F{row}", f"=E{row}-B{row}", '+0.0;-0.0;"-"', note="Change in bps from FY26 to FY29")
        else:
            put(summary, f"F{row}", f"=(E{row}/B{row})^(1/3)-1", PCT)
    r += 8
put(summary, f"A{r}", "FY26A adds up the two FY26 halves from the Data tab, which matches Tyro's reported FY26 gross profit of $231.8m.", color=GREY, italic=True)
for col, width in zip("ABCDEFGHIJ", [16, 18, 13, 13, 11, 18, 22, 14, 22, 20]):
    summary.column_dimensions[col].width = width


# Sensitivity tab
# FY27 total gross profit for different growth and squeeze numbers, everything else stays at Base
sens = wb.create_sheet("Sensitivity")
sheet_title(sens, "FY27 total gross profit ($m) for different core growth and squeeze",
            "Everything else is the Base scenario. Red is below guidance, amber is below the midpoint, green is at or above the midpoint.")
b = "Base"
squeezes = [0, 0.5, 1, 1.5, 2, 3, 4]
growths = [0, 0.01, 0.02, 0.03, 0.035, 0.04, 0.05, 0.06, 0.08]
put(sens, "A4", "Core TTV growth (down) vs squeeze in bps (across)", bold=True)
for j, sq in enumerate(squeezes):
    put(sens, f"{col_letter(2 + j)}4", sq, "0.0", BLUE, bold=True, fill=SUB_FILL, align="center")


def fy27_gp_formula(growth, squeeze):
    # rebuilds the FY27 Base numbers from the Forecast tab but with a different growth and squeeze
    parts = []
    for this_col, last_col in (("E", "C"), ("F", "D")):
        ttv = (f"(Forecast!${last_col}${frow(b, 1)}*(1+{growth})+Forecast!${this_col}${frow(b, 2)})"
               f"*(1+{inp('uplift', b)}*Forecast!${this_col}${frow(b, 5)})")
        margin = f"(Forecast!${this_col}${frow(b, 7)}+Forecast!${this_col}${frow(b, 9)}-{squeeze}*Forecast!${this_col}${frow(b, 6)})"
        parts.append(f"{ttv}*{margin}/10000")
    other = f"Forecast!$E${frow(b, 16)}+Forecast!$F${frow(b, 16)}+Forecast!$E${frow(b, 17)}+Forecast!$F${frow(b, 17)}"
    return "=" + "+".join(parts) + "+" + other


for i, g in enumerate(growths):
    row = 5 + i
    put(sens, f"A{row}", g, "0.0%", BLUE, bold=True, fill=SUB_FILL)
    for j in range(len(squeezes)):
        put(sens, f"{col_letter(2 + j)}{row}", fy27_gp_formula(f"$A{row}", f"{col_letter(2 + j)}$4"), MONEY)
last_row = 5 + len(growths) - 1
grid = f"B5:{col_letter(1 + len(squeezes))}{last_row}"
sens.conditional_formatting.add(grid, CellIsRule(operator="lessThan", formula=[inp("g_low")], fill=red))
sens.conditional_formatting.add(grid, CellIsRule(operator="between", formula=[inp("g_low"), inp("g_mid")], fill=amber))
sens.conditional_formatting.add(grid, CellIsRule(operator="greaterThan", formula=[inp("g_mid")], fill=green))

# growth needed to hit the guidance midpoint
# gross profit goes up in a straight line with growth so I can solve for it directly
need_row = last_row + 2
put(sens, f"A{need_row}", "Core TTV growth needed to hit the guidance midpoint", bold=True)
for j in range(len(squeezes)):
    sq = f"{col_letter(2 + j)}$4"
    per_dollar = [f"(1+{inp('uplift', b)}*Forecast!${c}${frow(b, 5)})*(Forecast!${c}${frow(b, 7)}+Forecast!${c}${frow(b, 9)}-{sq}*Forecast!${c}${frow(b, 6)})/10000"
                  for c in ("E", "F")]
    other = f"(Forecast!$E${frow(b, 16)}+Forecast!$F${frow(b, 16)}+Forecast!$E${frow(b, 17)}+Forecast!$F${frow(b, 17)})"
    top = (f"({inp('g_mid')}-{other}-(Forecast!$C${frow(b, 1)}+Forecast!$E${frow(b, 2)})*{per_dollar[0]}"
           f"-(Forecast!$D${frow(b, 1)}+Forecast!$F${frow(b, 2)})*{per_dollar[1]})")
    bottom = f"(Forecast!$C${frow(b, 1)}*{per_dollar[0]}+Forecast!$D${frow(b, 1)}*{per_dollar[1]})"
    put(sens, f"{col_letter(2 + j)}{need_row}", f"={top}/{bottom}", PCT, bold=True)
put(sens, f"A{need_row + 1}", "To compare, Tyro's core TTV grew 2.4% in FY24, 1.7% in FY25 and 4.4% in FY26.", color=GREY, italic=True)

# check that the grid gives the same answer as the Forecast tab
put(sens, f"A{need_row + 3}", "Check: grid at Base settings vs Forecast tab", color=GREY)
put(sens, f"B{need_row + 3}", fy27_gp_formula(inp("core_g", b), inp("squeeze", b)), MONEY, GREY)
put(sens, f"C{need_row + 3}", "=Summary!B7", MONEY, GREEN)
put(sens, f"D{need_row + 3}", f'=IF(ABS(B{need_row + 3}-C{need_row + 3})<0.001,"OK","CHECK")', bold=True)
sens.column_dimensions["A"].width = 46
for j in range(len(squeezes)):
    sens.column_dimensions[col_letter(2 + j)].width = 11


# Budget vs Actual tab
# empty for now, I fill it in when Tyro reports H1 FY27
bva = wb.create_sheet("Budget_vs_Actual")
sheet_title(bva, "H1 FY27: my Base forecast vs what actually happened",
            "Fill in the yellow cells when Tyro reports H1 FY27 (around late Feb 2027). Type $m like 23150.0 and bps like 45.1")
header_row(bva, 4, ["Metric", "Unit", "Forecast (Base)", "Actual", "Difference", "Difference %"])
metrics = [("Total TTV", "$m", 4, MONEY), ("Payments margin", "bps", 11, BPS), ("Payments gross profit", "$m", 12, MONEY),
           ("Banking gross profit", "$m", 16, MONEY), ("Corporate & other gross profit", "$m", 17, MONEY),
           ("Total gross profit", "$m", 18, MONEY)]
BVA_ROW = {}
for i, (label, unit, line, fmt) in enumerate(metrics):
    row = 5 + i
    BVA_ROW[line] = row
    put(bva, f"A{row}", label, bold=(line == 18))
    put(bva, f"B{row}", unit, color=GREY)
    put(bva, f"C{row}", f"=Forecast!E{frow('Base', line)}", fmt, GREEN)
    put(bva, f"D{row}", None, fmt, BLUE, fill=YELLOW)
    put(bva, f"E{row}", f'=IF(D{row}="","",D{row}-C{row})', fmt)
    put(bva, f"F{row}", f'=IF(D{row}="","",D{row}/C{row}-1)', PCT)

# splits the payments GP miss into "volume was different" and "margin was different"
split = 5 + len(metrics) + 1
put(bva, f"A{split}", "Why payments gross profit was different", bold=True)
header_row(bva, split + 1, ["Reason", "Unit", "Impact", "", "", "How I work it out"])
put(bva, f"A{split + 2}", "Volume")
put(bva, f"B{split + 2}", "$m", color=GREY)
put(bva, f"C{split + 2}", f'=IF(D{BVA_ROW[4]}="","",(D{BVA_ROW[4]}-C{BVA_ROW[4]})*C{BVA_ROW[11]}/10000)', MONEY)
put(bva, f"F{split + 2}", "(actual TTV minus forecast TTV) times my forecast margin", color=GREY)
put(bva, f"A{split + 3}", "Margin")
put(bva, f"B{split + 3}", "$m", color=GREY)
put(bva, f"C{split + 3}", f'=IF(D{BVA_ROW[11]}="","",(D{BVA_ROW[11]}-C{BVA_ROW[11]})*D{BVA_ROW[4]}/10000)', MONEY)
put(bva, f"F{split + 3}", "(actual margin minus forecast margin) times actual TTV", color=GREY)
put(bva, f"A{split + 4}", "Total", bold=True)
put(bva, f"B{split + 4}", "$m", color=GREY)
put(bva, f"C{split + 4}", f'=IF(D{BVA_ROW[4]}="","",C{split + 2}+C{split + 3})', MONEY, bold=True)
put(bva, f"F{split + 4}", "Should be about the same as the payments GP difference above", color=GREY)
for col, width in zip("ABCDEF", [34, 8, 16, 14, 12, 60]):
    bva.column_dimensions[col].width = width


for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False
out = ROOT / "Tyro_RBA_Reform_Model.xlsx"
wb.save(out)
print("saved", out)
