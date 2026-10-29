"""Independent cross-check of the Excel workbook.

Re-implements the Power Query cleaning rules in pandas, then compares the results
with the tables loaded in qld_fuel_price_analysis.xlsm (Cleaning_Log, dim_Sites,
fact_Daily, fact_Stockouts). Any difference is printed; the script exits with code 1
if a check fails.

Usage:  python validate_excel_results.py <folder with the monthly CSV files>
"""
import glob
import os
import sys

import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
WORKBOOK = os.path.join(HERE, "qld_fuel_price_analysis.xlsm")
CORE_FUELS = ["Unleaded", "e10", "Diesel"]
MIN_CPL, MAX_CPL, STOCKOUT = 100, 500, 9999
CBD = (-27.4698, 153.0251)
PERIOD_END = pd.Timestamp("2026-08-31")


def load_raw(folder):
    frames = []
    for f in sorted(glob.glob(os.path.join(folder, "*.csv"))):
        df = pd.read_csv(f, encoding="utf-8-sig", dtype=str)
        df.columns = [c.strip().replace(" ", "_") for c in df.columns]
        frames.append(df)
    return pd.concat(frames, ignore_index=True), len(frames)


def clean(raw):
    df = raw.copy()
    for c in ["Site_Name", "Site_Brand", "Site_Suburb", "Fuel_Type"]:
        df[c] = df[c].str.strip()
    df["Price"] = df["Price"].astype(int)
    df["Site_Latitude"] = df["Site_Latitude"].astype(float)
    df["Site_Longitude"] = df["Site_Longitude"].astype(float)
    df["ts"] = pd.to_datetime(df["TransactionDateutc"], format="%d/%m/%Y %H:%M") + pd.Timedelta(hours=10)
    df = df.drop_duplicates(subset=["SiteId", "Fuel_Type", "Price", "ts"])
    df["day"] = df["ts"].dt.normalize()
    df["cpl"] = df["Price"] / 10
    df["status"] = np.where(df["Price"] == STOCKOUT, "Stock-out",
                   np.where((df["cpl"] < MIN_CPL) | (df["cpl"] > MAX_CPL), "Out of range", "Valid"))
    return df


def sites(df):
    latest = df.sort_values("ts").groupby("SiteId").last()
    out = latest[["Site_Brand", "Site_Latitude", "Site_Longitude"]].rename(columns={"Site_Brand": "Brand"})
    lat1, lon1 = np.radians(CBD[0]), np.radians(CBD[1])
    lat2, lon2 = np.radians(out["Site_Latitude"]), np.radians(out["Site_Longitude"])
    a = np.sin((lat2 - lat1) / 2) ** 2 + np.cos(lat1) * np.cos(lat2) * np.sin((lon2 - lon1) / 2) ** 2
    out["Km_From_CBD"] = (2 * 6371 * np.arcsin(np.sqrt(a))).round(1)
    out["Region"] = np.where(out["Km_From_CBD"] <= 25, "Brisbane",
                    np.where(out["Km_From_CBD"] <= 150, "Rest of SEQ", "Regional QLD"))
    out["Brand_Changed"] = df.groupby("SiteId")["Site_Brand"].nunique() > 1
    return out


def daily(df, site_dim):
    core = df[(df["status"] == "Valid") & df["Fuel_Type"].isin(CORE_FUELS)]
    conflicts = int((core.groupby(["SiteId", "Fuel_Type", "ts"])["cpl"].nunique() > 1).sum())
    # last change of the day; if two prices share that minute, keep the lower one
    last_ts = core.groupby(["SiteId", "Fuel_Type", "day"])["ts"].transform("max")
    last = (core[core["ts"] == last_ts].groupby(["SiteId", "Fuel_Type", "day"])["cpl"].min().reset_index())
    parts = []
    for (sid, fuel), g in last.groupby(["SiteId", "Fuel_Type"]):
        days = pd.date_range(g["day"].min(), PERIOD_END)
        s = g.set_index("day")["cpl"].reindex(days).ffill()
        parts.append(pd.DataFrame({"SiteId": sid, "Fuel_Type": fuel, "Price_Date": days, "cpl": s.values}))
    d = pd.concat(parts).merge(site_dim[["Brand", "Region"]], left_on="SiteId", right_index=True)
    fact = d.groupby(["Price_Date", "Fuel_Type", "Brand", "Region"]).agg(
        Site_Days=("cpl", "size"), Price_Sum=("cpl", "sum")).reset_index()
    return fact, len(core), conflicts


def read_sheet(name):
    return pd.read_excel(WORKBOOK, sheet_name=name, engine="openpyxl")


def main(folder):
    failures = []

    def check(label, expected, actual, tol=0.0):
        ok = abs(expected - actual) <= tol if isinstance(expected, (int, float, np.number)) else expected == actual
        print(f"[{'OK' if ok else 'FAIL'}] {label}: python={expected}  excel={actual}")
        if not ok:
            failures.append(label)

    raw, n_files = load_raw(folder)
    df = clean(raw)
    site_dim = sites(df)
    fact, core_rows, conflicts = daily(df, site_dim)

    log = read_sheet("Cleaning_Log").set_index("Step")
    # pandas skips empty lines on read, Power Query keeps them until step 3
    check("Step 3 rows after removing blanks", len(raw), int(log.loc[3, "Rows out"]))
    check("Step 5 duplicates removed", len(raw) - len(df), int(log.loc[5, "Rows affected"]))
    check("Step 6 stock-outs flagged", int((df["status"] == "Stock-out").sum()), int(log.loc[6, "Rows affected"]))
    check("Step 7 out-of-range flagged", int((df["status"] == "Out of range").sum()), int(log.loc[7, "Rows affected"]))
    check("Step 8 core valid rows", core_rows, int(log.loc[8, "Rows out"]))
    check("Step 9 same-minute conflicts", conflicts, int(log.loc[9, "Rows affected"]))
    check("Step 11 sites", len(site_dim), int(log.loc[11, "Rows out"]))
    check("Step 11 sites with brand change", int(site_dim["Brand_Changed"].sum()), int(log.loc[11, "Rows affected"]))
    check("Step 13 daily site-days", int(fact["Site_Days"].sum()), int(log.loc[13, "Rows out"]))
    check("Step 14 fact_Daily rows", len(fact), int(log.loc[14, "Rows out"]))

    xs = read_sheet("Data_Sites")
    check("Region counts", site_dim["Region"].value_counts().sort_index().to_dict(),
          xs["Region"].value_counts().sort_index().to_dict())

    xf = read_sheet("Data_Daily")
    xf["Price_Date"] = pd.to_datetime(xf["Price_Date"])
    for fuel in CORE_FUELS:
        p = fact[fact["Fuel_Type"] == fuel]
        e = xf[xf["Fuel_Type"] == fuel]
        check(f"{fuel} 12-month average c/L", round(p["Price_Sum"].sum() / p["Site_Days"].sum(), 3),
              round(e["Price_Sum"].sum() / e["Site_Days"].sum(), 3), tol=0.001)

    so = df[df["status"] == "Stock-out"]
    xso = read_sheet("Data_Stockouts")
    check("Stock-out events", len(so), int(xso["Stockout_Events"].sum()))

    print("\nAll checks passed." if not failures else f"\n{len(failures)} check(s) failed: {failures}")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "..", "raw-data-staging", "qld-fuel-prices"))
