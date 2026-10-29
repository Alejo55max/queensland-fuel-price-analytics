// fact_Stockouts — stock-out events (price code 9999 = "fuel temporarily unavailable").
// These rows are excluded from price averages but kept here as evidence of supply stress.
let
    Source = Prices_Clean,
    KeptStockouts = Table.SelectRows(Source, each [Price_Status] = "Stock-out"),
    MergedSiteAttributes = Table.Join(KeptStockouts, "SiteId",
        Table.RenameColumns(Table.SelectColumns(dim_Sites, {"SiteId", "Brand", "Brand_Group", "Region"}), {{"SiteId", "Dim_SiteId"}}), "Dim_SiteId"),
    Aggregated = Table.Group(MergedSiteAttributes, {"Price_Date", "Fuel_Type", "Brand", "Brand_Group", "Region"}, {
        {"Stockout_Events", each Table.RowCount(_), Int64.Type},
        {"Sites_Affected", each List.Count(List.Distinct([SiteId])), Int64.Type}}),
    AddedMonth = Table.AddColumn(Aggregated, "Month", each Date.StartOfMonth([Price_Date]), type date),
    Sorted = Table.Sort(AddedMonth, {{"Price_Date", Order.Ascending}})
in
    Sorted
