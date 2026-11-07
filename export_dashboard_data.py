# export_dashboard_data.py
# Makes the tables my dashboard needs and saves them in data/dashboard.
# It takes the history from my csv files, the forecasts from the Excel model and two answers from my SQL file.
# If I change anything in the Excel model I need to open it in Excel and save it first,
# otherwise the forecast numbers in the file won't be updated.

from pathlib import Path
import sqlite3
import pandas as pd
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "dashboard"
OUT.mkdir(parents=True, exist_ok=True)

half = pd.read_csv(ROOT / "data/tyro_halfyear.csv")
FORECAST_PERIODS = ["H1FY27", "H2FY27", "H1FY28", "H2FY28", "H1FY29", "H2FY29"]
ALL_PERIODS = list(half.period) + FORECAST_PERIODS
# a number for each period so charts can put them in the right order
PERIOD_ORDER = {p: i + 1 for i, p in enumerate(ALL_PERIODS)}


# half year actuals and forecasts, one row per period and scenario
rows = []
for _, h in half.iterrows():
    rows.append({"period": h.period, "scenario": "Actual", "type": "Actual", "ttv_m": h.ttv_m,
                 "payments_margin_bps": round(h.calc_margin_bps, 2), "payments_gp_m": h.payments_gp_m,
                 "banking_gp_m": h.banking_gp_m, "corporate_gp_m": h.corporate_other_gp_m,
                 "total_gp_m": h.total_gp_m})

# data_only=True gives me the calculated values instead of the formulas
wb = load_workbook(ROOT / "Tyro_RBA_Reform_Model.xlsx", data_only=True)
fc = wb["Forecast"]

# find where each scenario block starts on the Forecast tab
block_start = {}
for r in range(1, fc.max_row + 1):
    v = fc.cell(r, 1).value
    if isinstance(v, str) and v.endswith("SCENARIO"):
        block_start[v.split()[0].title()] = r

# which row inside a block holds each line item
line_row = {"ttv_m": 4, "payments_margin_bps": 11, "payments_gp_m": 12, "payments_revenue_m": 15,
            "banking_gp_m": 16, "corporate_gp_m": 17, "total_gp_m": 18}

if fc.cell(block_start["Base"] + 18, 5).value is None:
    raise SystemExit("The Excel file has no calculated numbers. Open it in Excel, save it, then run this again.")

for scenario, top in block_start.items():
    if scenario == "Custom":   # Custom is just for playing around so I leave it out
        continue
    for j, period in enumerate(FORECAST_PERIODS):
        rec = {"period": period, "scenario": scenario, "type": "Forecast"}
        for name, line in line_row.items():
            rec[name] = round(fc.cell(top + line, 5 + j).value, 2)
        rows.append(rec)

halfyear = pd.DataFrame(rows)
halfyear["period_order"] = halfyear.period.map(PERIOD_ORDER)
halfyear["fy"] = "FY" + halfyear.period.str[-2:]
halfyear.to_csv(OUT / "dash_halfyear.csv", index=False)


# full year totals, plus the FY27 guidance range to compare against
fullyear = (halfyear.groupby(["fy", "scenario", "type"], as_index=False)
            .agg(ttv_m=("ttv_m", "sum"), payments_gp_m=("payments_gp_m", "sum"), total_gp_m=("total_gp_m", "sum")))
fullyear["payments_margin_bps"] = (fullyear.payments_gp_m / fullyear.ttv_m * 1e4).round(2)
fullyear[["ttv_m", "payments_gp_m", "total_gp_m"]] = fullyear[["ttv_m", "payments_gp_m", "total_gp_m"]].round(2)
fullyear["guidance_low_m"] = fullyear.fy.map({"FY27": 240.0})
fullyear["guidance_high_m"] = fullyear.fy.map({"FY27": 255.0})
fullyear.to_csv(OUT / "dash_fullyear.csv", index=False)


# the sensitivity grid turned into a long list so a dashboard can build a matrix from it
sens = wb["Sensitivity"]
squeezes = [sens.cell(4, c).value for c in range(2, 9)]
grid = []
for r in range(5, 14):
    growth = sens.cell(r, 1).value
    for j, sq in enumerate(squeezes):
        gp = sens.cell(r, 2 + j).value
        if gp < 240:
            band = "Below guidance"
        elif gp < 247.5:
            band = "Below midpoint"
        else:
            band = "At or above midpoint"
        grid.append({"core_growth": growth, "squeeze_bps": sq, "fy27_total_gp_m": round(gp, 1), "band": band})
pd.DataFrame(grid).to_csv(OUT / "dash_sensitivity.csv", index=False)

# growth needed to hit the guidance midpoint for each squeeze
need_row = next(r for r in range(1, sens.max_row + 1) if str(sens.cell(r, 1).value).startswith("Core TTV growth needed"))
pd.DataFrame({"squeeze_bps": squeezes,
              "core_growth_needed_for_midpoint": [round(sens.cell(need_row, 2 + j).value, 4) for j in range(len(squeezes))]}
             ).to_csv(OUT / "dash_breakeven.csv", index=False)


# market share (Q1) and Tyro's price vs the market (Q4) straight from my SQL file
con = sqlite3.connect(ROOT / "data/tyro_project.db")
sql_text = (ROOT / "sql/week1_queries.sql").read_text()

def run_query(tag):
    block = next(b for b in sql_text.split(";") if tag in b)
    query = "\n".join(l for l in block.splitlines() if not l.strip().startswith("--"))
    return pd.read_sql_query(query, con)

share = run_query("-- Q1.")
share["period_order"] = share.period.map(PERIOD_ORDER)
share.to_csv(OUT / "dash_market_share.csv", index=False)
run_query("-- Q4.").to_csv(OUT / "dash_pricing_vs_market.csv", index=False)

for f in sorted(OUT.glob("*.csv")):
    print(f"{f.name}: {len(pd.read_csv(f))} rows")
