# Executive Summary — Queensland Fuel Price Shock 2026

*Workbook: [`qld_fuel_price_analysis.xlsm`](qld_fuel_price_analysis.xlsm). Power Query code: [`power-query/`](power-query/). VBA: [`RefreshReport.bas`](RefreshReport.bas). Independent validation: [`validate_excel_results.py`](validate_excel_results.py).*

## Business Problem

The March–April 2026 Middle East supply shock (strikes on Iran disrupting the Strait of Hormuz) pushed Australian pump prices to records within weeks. This analysis uses Queensland's mandatory fuel price reporting data to answer five questions:

1. How big was the shock, and how fast did it come and go?
2. Did regional Queensland pay more than Brisbane?
3. Which brands were cheapest?
4. When is the cheapest time to fill up?
5. Did the shock cause real shortages?

A fleet cost calculator then turns the answers into dollars.

## Data

**Source:** Queensland Government Open Data Portal, *Fuel Price Reporting* ([2025](https://www.data.qld.gov.au/dataset/fuel-price-reporting-2025), [2026](https://www.data.qld.gov.au/dataset/fuel-price-reporting-2026)), CC BY 4.0. Queensland fuel retailers must report every price change. Each row is one change at one station for one fuel, with a UTC timestamp. Price is in tenths of a cent, and 9999 means "temporarily unavailable".

| Table (Power Query output) | Rows | Content |
|---|---|---|
| Raw (12 monthly CSVs) | 911,299 | Every price change, Sep 2025 – Aug 2026, 1,807 stations, 32 brands, 9 fuel types |
| `fact_Daily` | 71,420 | Date × fuel × brand × region: `Site_Days` (stations reporting) and `Price_Sum` (sum of their prices) |
| `dim_Sites` | 1,807 | One row per station: latest name and brand, distance to Brisbane CBD, region, brand group |
| `fact_Stockouts` | 6,508 | Date × fuel × brand × region: stock-out events and stations affected |

## Cleaning Log (14 steps, recalculated by Power Query on every refresh)

| Step | Action | Rows affected | Why |
|---|---|---|---|
| 1 | Combine 12 monthly files | 911,299 rows read | One table for the full year |
| 2 | Standardise column headers | 1 file | The March 2026 file used `Site Brand` instead of `Site_Brand`; without the fix its columns load empty |
| 3 | Remove blank rows | 10 | Most files end with an empty line |
| 4 | Parse timestamps, convert UTC → Brisbane (UTC+10) | 0 failed | An 11 pm UTC change belongs to the next Brisbane day |
| 5 | Remove duplicate events | 6 | Same station, fuel, price and minute |
| 6 | Flag stock-outs (code 9999) | 8,800 | Not a price: excluded from averages, analysed separately |
| 7 | Flag out-of-range prices (< 100 or > 500 c/L) | 539 | Data-entry errors |
| 8 | Keep core fuels (Unleaded, e10, Diesel) | 451,697 kept | Like-for-like comparison; premium fuels, LPG and E85 excluded |
| 9 | Resolve same-minute conflicts (keep the lower price) | 312 | Two prices in the same minute have no knowable order; usually a typo corrected seconds later (e.g. 495.9 c then 195.9 c) |
| 10 | Keep last price of each day | 347,118 daily prices | The closing price is the one in effect overnight |
| 11 | Use latest station details | 113 stations rebranded | e.g. Caltex → Ampol during the year |
| 12 | Group small brands | 219 stations → "Other brands" | Top 12 brands by station count kept individually; original brand retained |
| 13 | Carry prices forward to every day | 1,251,418 station-days | The source records only changes: each station counts once per day |
| 14 | Aggregate for analysis | 71,420 rows | Averages stay exact as `SUM(Price_Sum) / SUM(Site_Days)` |

**Why step 13 matters.** A station that changes price daily produces 30 rows a month, while one that changes monthly produces one. A simple average of rows would let the busy changers dominate. Carrying each price forward until the next change gives one observation per station per day, so the result is the average price a driver would actually have found.

## Methodology

1. **Power Query** (8 queries, code in `power-query/`). A `Settings` table drives every threshold, so nothing important is hard-coded. The queries were tuned after the first version took more than 30 minutes and exhausted memory:
   - Measuring each query in isolation found the bottleneck: a lazily evaluated list that re-ran the whole pipeline once per station. `List.Buffer` fixed it.
   - Wide rows (names and addresses carried across 900k rows) were replaced with narrow 4-column price tables.
   - The slowest query dropped from more than 30 minutes to about one minute.
2. **Analysis sheet**: five questions answered with dynamic-array formulas. Two dropdowns (fuel, region) re-run everything; "All regions" uses a `*` wildcard inside `SUMIFS`.
   - The pre-shock baseline is the 28 days before 2 March 2026.
   - Shock window: 2 March – 30 April 2026.
   - Price-cycle peaks and lows are days that are the highest or lowest price within ±3 days, found with `MAP` + `LAMBDA`.
3. **Dashboard**: four PivotCharts with a `Price_Sum / Site_Days` calculated field (an exact station-weighted average), slicers, KPI cards and VBA buttons.
4. **Validation**: a separate Python script repeats all cleaning rules and compares 15 results with the workbook, and all 15 match. Issues it caught and fixed:
   - a timezone shift on the `Settings` dates that cut one day from the period;
   - same-minute conflicts that the two tools had resolved differently.

## Results

### Q1. Size and speed of the shock (statewide diesel)

| Measure | Value |
|---|---|
| Pre-shock baseline (28 days before 2 Mar 2026) | 181.2 c/L |
| Peak daily average | **324.4 c/L on 9 Apr 2026 (+79.0%)** |
| Days from shock start to peak | 38 |
| First day back within 10% of baseline | 22 Jun 2026 (74 days after the peak) |
| End of period (31 Aug 2026) | 254.2 c/L (+40.3%) |

Unleaded in Brisbane rose from 178.7 to 257.9 c/L (+44.4%, peak 30 Mar). It took 28 days to peak and 22 days to return within 10%. Diesel was hit harder and stayed high much longer.

### Q2. Regional gap (diesel, average c/L)

| Region | Before shock | Shock (Mar–Apr) | After shock | Shock gap vs Brisbane |
|---|---|---|---|---|
| Brisbane (≤ 25 km of CBD) | 186.5 | 277.7 | 220.7 | — |
| Rest of SEQ (25–150 km) | 183.3 | 277.6 | 220.0 | 0.0 |
| Regional QLD (> 150 km) | 187.2 | 285.3 | 232.6 | +7.6 |

The regional premium went from 0.7 c/L before the shock to 7.6 c/L during it and 11.9 c/L after it.

### Q3. Brands (diesel, all regions, brands averaging ≥ 10 stations a day)

The cheapest brands before the shock were Pearl Energy (175.3 c/L, 10.6 below market), Liberty (178.2) and United (179.2). The dearest were IOR (194.5) and Reddy Express (193.1), a **19.2 c/L spread**. During the shock, IOR rose the most (+105.5 c/L), while Reddy Express (+85.9) and Caltex (+86.8) rose the least. For Brisbane unleaded, Liberty (173.6 c/L) and United (175.6) led and EG Ampol (189.5) was dearest.

### Q4. When to fill up (Brisbane unleaded, before the shock)

| Measure | Value |
|---|---|
| Cheapest / dearest weekday | Tuesday 184.67 / Sunday 185.63 c/L, **0.95 c/L apart ($0.57 per 60 L tank)** |
| Average price-cycle peak / low | 205.4 / 169.2 c/L, **36.2 c/L swing ($21.70 per 60 L tank)** |
| Cycle length | about 38 days between peaks (4 peaks before the shock) |

Diesel shows no meaningful cycle (2.2 c/L swing statewide).

### Q5. Stock-outs (all fuels, events per month)

The average before the shock was about 288 a month. The count rose to **1,917 in March 2026 (6.7×)** and 1,231 in April. It was still 1,119 in July and 1,083 in August, about 3.8× normal. Regional Queensland had the most stock-outs in July (479).

### Fleet calculator example

20 diesel vans in Brisbane using 400 L each per month, all brands:

| Measure | Value |
|---|---|
| Pre-shock price | 180.6 c/L |
| 12-month fuel bill | $204,292 |
| Extra cost Mar–Aug 2026 vs pre-shock price | **$28,089 ($1,404 per vehicle)** |
| Worst month | April 2026, $8,843 extra |

## Recommendations

1. Budget for a slow recovery: diesel took twice as long to come down as to go up and was still +40% six months later.
2. Pick fuel brands on price. The diesel spread of up to 19 c/L outweighs typical fuel-card discounts.
3. For unleaded in SEQ, time purchases to the price cycle (worth up to about 36 c/L), not the weekday (under 1 c/L).
4. Treat regional operations as higher-risk for both price and supply: review on-site storage or contracted supply.
5. Refresh the workbook monthly so the next shock is measured as it happens.

## Limitations

- Daily closing prices are carried forward, so same-day intraday discounts are not counted.
- Regions are distance bands from the Brisbane CBD, not ABS boundaries (editable in `Settings`).
- Averages are station-weighted, because the data contains no sales volumes.
- The fleet calculator assumes constant litres and average prices for the chosen region and brand.
- A full refresh takes about 15–30 minutes on a laptop. That suits a monthly report; multi-year history would belong in Power BI or a database.
