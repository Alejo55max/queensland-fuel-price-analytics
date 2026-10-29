// fact_Daily — daily price table for analysis, one row per Date x Fuel x Brand x Region.
// Each price window is expanded into one row per day it was in effect, then summed.
// Stores Site_Days and Price_Sum (not an average) so any roll-up in Excel
// (=SUM(Price_Sum)/SUM(Site_Days)) is an exact site-weighted average.
let
    Source = Daily_InEffect,
    MergedSiteAttributes = Table.Join(Source, "SiteId",
        Table.RenameColumns(Table.SelectColumns(dim_Sites, {"SiteId", "Brand", "Brand_Group", "Region"}), {{"SiteId", "Dim_SiteId"}}), "Dim_SiteId"),
    AddedDayList = Table.AddColumn(MergedSiteAttributes, "Price_Date", each
        List.Dates([Valid_From], [Days_In_Effect], #duration(1, 0, 0, 0))),
    ExpandedToDays = Table.ExpandListColumn(
        Table.SelectColumns(AddedDayList, {"Price_Date", "Fuel_Type", "Brand", "Brand_Group", "Region", "Price_cpl"}), "Price_Date"),
    Aggregated = Table.Group(ExpandedToDays, {"Price_Date", "Fuel_Type", "Brand", "Brand_Group", "Region"}, {
        {"Site_Days", each Table.RowCount(_), Int64.Type},
        {"Price_Sum", each List.Sum([Price_cpl]), type number}}),
    TypedDate = Table.TransformColumnTypes(Aggregated, {{"Price_Date", type date}}),
    AddedMonth = Table.AddColumn(TypedDate, "Month", each Date.StartOfMonth([Price_Date]), type date),
    AddedWeekday = Table.AddColumn(AddedMonth, "Weekday", each Date.DayOfWeekName([Price_Date], "en-AU"), type text),
    AddedPeriod = Table.AddColumn(AddedWeekday, "Period", each
        if [Price_Date] < Date.From(Settings[Shock_Start]) then "Before shock"
        else if [Price_Date] <= Date.From(Settings[Shock_End]) then "Shock (Mar-Apr 2026)"
        else "After shock", type text),
    Sorted = Table.Sort(AddedPeriod, {{"Price_Date", Order.Ascending}, {"Fuel_Type", Order.Ascending},
        {"Region", Order.Ascending}, {"Brand", Order.Ascending}})
in
    Sorted
