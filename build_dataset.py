# build_dataset.py
# This script builds all my clean tables for the Tyro project.
# Tyro only publishes PDFs so I typed their numbers in below and wrote the page I got each one from.
# The RBA numbers get read straight from the Excel files in data/raw.
# At the end it saves 5 csv files and one SQLite database into the data folder.
# All money is in $m.

from pathlib import Path
import sqlite3
import warnings
import pandas as pd

# the RBA files have a weird print area setting that makes pandas show a warning, its harmless
warnings.filterwarnings("ignore", message="Print area cannot be set")

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data" / "raw"
OUT = ROOT / "data"
OUT.mkdir(exist_ok=True)

# short names I use for the source documents
# H1FY24P = H1 FY24 investor presentation, FY26AR = annual report 2026 etc
# page numbers are the page of the PDF file, not the number printed on the page


# Tyro yearly numbers
# FY means Tyro's financial year which ends 30 June
# msf = what merchants pay Tyro, net maf = msf minus interchange and scheme fees
# both are in bps of transaction value (1 bp = 0.01%)

annual = pd.DataFrame([
    ["FY23", 42601.3, 419.2, 234.6, 177.4, 89.0, 33.2, 41.6, 9.7, 193.2, None, 68665,
     "FY24AR p27-28 (FY23 comparatives), total GP from FY26P p19, merchants FY24P p38"],
    ["FY24", 42932.7, 471.5, 251.7, 186.7, 94.1, 35.4, 43.5, 12.6, 210.8, 55.7, 71347,
     "FY24AR p27-28, FY26P p20, merchants FY24P p38"],
    ["FY25", 43012.0, 460.9, 255.6, 195.0, 98.9, 38.3, 45.3, 13.4, 220.1, 61.6, 76195,
     "FY26AR p32-33 (FY25 comparatives, MSF restated from 99.1 to 98.9), FY26P p20, merchants FY25P p39"],
    ["FY26", 44253.3, 459.2, 247.7, 203.0, 97.0, 39.7, 45.9, 16.5, 231.8, 66.9, 77865,
     "FY26AR p32-33, FY26P p16 p20 p27, merchants FY26P p31"],
], columns=["fy", "ttv_m", "payments_revenue_m", "payments_direct_exp_m", "payments_gp_m",
            "msf_bps", "net_maf_bps", "payments_margin_bps", "banking_gp_m", "total_gp_m",
            "ebitda_m", "merchants", "source"])

# corporate and other is whatever is left after payments and banking
annual["corporate_other_gp_m"] = (annual.total_gp_m - annual.payments_gp_m - annual.banking_gp_m).round(1)
# I work the margin out myself as well so I can check it matches what Tyro reported
annual["check_margin_bps"] = (annual.payments_gp_m / annual.ttv_m * 1e4).round(1)
annual["direct_cost_bps"] = (annual.payments_direct_exp_m / annual.ttv_m * 1e4).round(1)


# Tyro half year numbers
# H1 is Jul to Dec and H2 is Jan to Jun
# Tyro reports H1 and the full year but not H2 on its own, so I get H2 by doing full year minus H1

h1 = pd.DataFrame([
    ["H1FY23", 21693.0, 88.2, 4.6, 95.2, 40.7, 32.7,
     "TTV is the sum of the verticals", "H1FY24P p17 (read with OCR) and p31, margins from H1FY25P p17"],
    ["H1FY24", 22173.5, 92.7, 7.1, 105.2, 41.8, 33.7,
     "reported", "H1FY24P p17 (OCR), H1FY25P p28, margins FY25P p24"],
    ["H1FY25", 21971.4, 98.2, 7.8, 112.0, 44.7, 37.6,
     "payments GP worked out as 104.1 / 1.060 (H1FY26P p9 says it grew 6%)", "H1FY25P p6 p18 p28, FY25P p24"],
    ["H1FY26", 22856.6, 104.1, 8.2, 117.6, 45.5, None,
     "banking GP worked out as 7.8 x 1.054 (H1FY26P p11 says it grew 5.4%)", "H1FY26P p9 p11 p22"],
], columns=["period", "ttv_m", "payments_gp_m", "banking_gp_m", "total_gp_m",
            "reported_margin_bps", "net_maf_bps", "basis", "source"])

fy = annual.set_index("fy")

# Tyro does show H2 margins in its charts so I keep them to check my H2 maths against
h2_margin_from_tyro = {"H2FY23": 42.7, "H2FY24": 45.3, "H2FY25": 46.0, "H2FY26": None}
h2_net_maf_from_tyro = {"H2FY23": 33.8, "H2FY24": 37.1, "H2FY25": 38.9, "H2FY26": None}

h2_rows = []
for _, r in h1.iterrows():
    year = "FY" + r.period[-2:]
    a = fy.loc[year]
    h2 = "H2" + year
    h2_rows.append({
        "period": h2,
        "ttv_m": round(a.ttv_m - r.ttv_m, 1),
        "payments_gp_m": round(a.payments_gp_m - r.payments_gp_m, 1),
        "banking_gp_m": round(a.banking_gp_m - r.banking_gp_m, 1),
        "total_gp_m": round(a.total_gp_m - r.total_gp_m, 1),
        "reported_margin_bps": h2_margin_from_tyro[h2],
        "net_maf_bps": h2_net_maf_from_tyro[h2],
        "basis": "full year minus H1",
        "source": f"{year} yearly numbers minus {r.period}, Tyro H2 margins from FY25P p24 and H1FY25P p17",
    })

half = pd.concat([h1, pd.DataFrame(h2_rows)], ignore_index=True)
half["fy"] = "FY" + half.period.str[-2:]
half["half"] = half.period.str[:2]

# start and end dates so I can line Tyro up with the RBA monthly data later
def start_date(r):
    if r.half == "H1":
        return f"20{int(r.fy[2:]) - 1}-07-01"
    return f"20{r.fy[2:]}-01-01"

def end_date(r):
    if r.half == "H1":
        return f"20{int(r.fy[2:]) - 1}-12-31"
    return f"20{r.fy[2:]}-06-30"

half["start_date"] = half.apply(start_date, axis=1)
half["end_date"] = half.apply(end_date, axis=1)
half["corporate_other_gp_m"] = (half.total_gp_m - half.payments_gp_m - half.banking_gp_m).round(1)
half["calc_margin_bps"] = (half.payments_gp_m / half.ttv_m * 1e4).round(1)
half = half.sort_values("start_date").reset_index(drop=True)
half = half[["period", "fy", "half", "start_date", "end_date", "ttv_m", "payments_gp_m", "banking_gp_m",
             "corporate_other_gp_m", "total_gp_m", "reported_margin_bps", "calc_margin_bps",
             "net_maf_bps", "basis", "source"]]


# Tyro numbers by vertical (industry)
# TTV is for the whole period, merchant count is at the end of the period
# order of the lists: Hospitality, Retail, Health, Service/Other, Bendigo

verticals = ["Hospitality", "Retail", "Health", "Service/Other", "Bendigo"]
vert_data = {
    "H1FY23": ([9286, 5508, 2527, 1589, 2783], [17072, 10783, 14306, 8006, 16717], "H1FY24P p31 (OCR)"),
    "H1FY24": ([9455, 5393, 3139, 1704, 2483], [16333, 10850, 18716, 8917, 13964], "H1FY25P p32"),
    "H1FY25": ([9058, 5314, 3648, 1733, 2218], [15955, 10586, 25277, 9206, 12535], "H1FY25P p32"),
    "H1FY26": ([9509, 5558, 3990, 1804, 1996], [15437, 12620, 23182, 13877, 11549], "H1FY26P p26"),
    "FY23":   ([18347, 10360, 5361, 3169, 5364], [17425, 10874, 16165, 8525, 15676], "FY24P p38"),
    "FY24":   ([18049, 10301, 6464, 3311, 4808], [16114, 10639, 22069, 8944, 13581], "FY25P p39"),
    "FY25":   ([17675, 10282, 7433, 3380, 4243], [16111, 10479, 28046, 9365, 12194], "FY26P p31"),
    "FY26":   ([18341, 10684, 7910, 3521, 3796], [15235, 10437, 31666, 9332, 11195], "FY26P p31"),
}

vrows = []
for period, (ttvs, merchants, src) in vert_data.items():
    for v, t, m in zip(verticals, ttvs, merchants):
        vrows.append({"period": period, "vertical": v, "ttv_m": t, "merchants_end": m,
                      "basis": "reported", "source": src})

# H2 again is full year minus H1 for TTV
# for merchants the 30 June count is just the full year count
for year in ["FY23", "FY24", "FY25", "FY26"]:
    ttv_fy, merch_fy, _ = vert_data[year]
    ttv_h1, _, _ = vert_data["H1" + year]
    for v, a, b, m in zip(verticals, ttv_fy, ttv_h1, merch_fy):
        vrows.append({"period": "H2" + year, "vertical": v, "ttv_m": a - b, "merchants_end": m,
                      "basis": "TTV is full year minus H1, merchants are the 30 June count",
                      "source": f"{year} and H1{year} tables"})

vert = pd.DataFrame(vrows)
vert["is_core"] = vert.vertical != "Bendigo"   # Bendigo is a bank partnership, not Tyro's own merchants


# RBA data
# I find each RBA file by part of its name so new downloads still work

def find_raw(name_part):
    hits = sorted(p for p in RAW.glob("*.xlsx") if name_part in p.name.lower())
    if not hits:
        raise FileNotFoundError(f"cant find a file with '{name_part}' in its name in {RAW}")
    return hits[0]

def read_rba(name_part, sheet, series_ids):
    # in the RBA files row 11 has the series ids and the data starts on row 12
    df = pd.read_excel(find_raw(name_part), sheet_name=sheet, header=None)
    ids = df.iloc[10].tolist()
    data = df.iloc[11:].copy()
    data.columns = ["date"] + ids[1:]
    data["date"] = pd.to_datetime(data["date"])
    data = data[["date"] + series_ids]
    for c in series_ids:
        data[c] = pd.to_numeric(data[c], errors="coerce")
    return data

# credit card and debit card spending
# "acquired in Australia" means spending at Australian merchants which is the market Tyro is in
credit = read_rba("c01-1", "Aggregate", ["CCAQTTV", "CCAQTFV", "CCCCSTPV"])
debit = read_rba("c02-1", "Aggregate", ["CDAQTTV", "CDAQTICV", "CDCPTTV"])
cards = credit.merge(debit, on="date", how="outer").sort_values("date")
cards = cards.rename(columns={
    "CCAQTTV": "credit_acquired_au_m", "CCAQTFV": "credit_overseas_cards_m", "CCCCSTPV": "credit_purchases_au_issued_m",
    "CDAQTTV": "debit_acquired_au_m", "CDAQTICV": "debit_overseas_cards_m", "CDCPTTV": "debit_purchases_au_issued_m"})
cards = cards[cards.date >= "2015-07-01"].copy()
cards["total_acquired_au_m"] = cards.credit_acquired_au_m + cards.debit_acquired_au_m
cards["credit_share"] = (cards.credit_acquired_au_m / cards.total_acquired_au_m).round(4)
cards["month"] = cards.date.dt.strftime("%Y-%m")

# tag each month with Tyro's half year so I can join them later
# Jul to Dec is H1 of the next financial year, Jan to Jun is H2 of the same year
def tyro_period(d):
    if d.month >= 7:
        return f"H1FY{(d.year + 1) % 100:02d}"
    return f"H2FY{d.year % 100:02d}"

cards["period"] = cards.date.apply(tyro_period)
cards = cards[["month", "period", "credit_acquired_au_m", "debit_acquired_au_m", "total_acquired_au_m",
               "credit_share", "credit_overseas_cards_m", "debit_overseas_cards_m",
               "credit_purchases_au_issued_m", "debit_purchases_au_issued_m"]]

# average merchant fees, quarterly, as % of value
# the old combined Mastercard and Visa series stopped being updated so I use the ones per scheme
# heads up: from the Dec 2024 quarter the RBA added big payment facilitators which made eftpos jump
fee_ids = {"CMFEPT": "eftpos", "CMFVDDT": "visa_debit_dom", "CMFMDDT": "mc_debit_dom",
           "CMFVCDT": "visa_credit_dom", "CMFMCDT": "mc_credit_dom", "CMFCAET": "amex"}
fees = read_rba("c03", "Data", list(fee_ids))
fees = fees.rename(columns=fee_ids)
fees = fees[fees.date >= "2015-07-01"].copy()
fees["quarter_end"] = fees.date.dt.strftime("%Y-%m-%d")
fees["period"] = fees.date.apply(tyro_period)
fees = fees[["quarter_end", "period"] + list(fee_ids.values())]


# save everything
tables = {"tyro_annual": annual, "tyro_halfyear": half, "tyro_vertical": vert,
          "rba_cards_monthly": cards, "rba_merchant_fees_quarterly": fees}

con = sqlite3.connect(OUT / "tyro_project.db")
for name, df in tables.items():
    df.to_csv(OUT / f"{name}.csv", index=False)
    df.to_sql(name, con, if_exists="replace", index=False)
    print(f"{name}: {len(df)} rows")
con.close()


# checks
# if something doesnt add up I want to know straight away

problems = []

# my margin should match Tyro's reported margin for every year
for _, r in annual.iterrows():
    if abs(r.check_margin_bps - r.payments_margin_bps) > 0.15:
        problems.append(f"{r.fy} margin: mine {r.check_margin_bps}, Tyro {r.payments_margin_bps}")

# same for half years, this also checks my H2 maths because Tyro reported those margins
for _, r in half.dropna(subset=["reported_margin_bps"]).iterrows():
    if abs(r.calc_margin_bps - r.reported_margin_bps) > 0.25:
        problems.append(f"{r.period} margin: mine {r.calc_margin_bps}, Tyro {r.reported_margin_bps}")

# verticals should add up to total TTV (allowing a bit for rounding)
vert_total = vert.groupby("period").ttv_m.sum()
for p in half.period:
    t = half.set_index("period").loc[p, "ttv_m"]
    if abs(vert_total[p] - t) > 3:
        problems.append(f"{p} verticals add to {vert_total[p]} but TTV is {t}")

if problems:
    print("\nCHECKS FAILED:")
    for p in problems:
        print(" ", p)
else:
    print("\nCHECKS: all passed")
