// Daily_InEffect — how long each price stayed in effect at each site.
// The source only records price CHANGES, so averaging raw events would over-weight
// sites that change price often. Steps 8, 9, 10 and 13: keep core fuels and valid prices, resolve
// same-minute conflicts, keep the last price of each site/fuel/day, then give every price a validity window: from its
// day until the day before the next change (or the end of the period).
// One row per window; fact_Daily expands each window into one row per day.
let
    CoreFuels = List.Buffer(Text.Split(Settings[Core_Fuels], ",")),
    PeriodEnd = Date.From(Settings[Period_End]),
    Source = Prices_Clean,
    KeptValidCoreFuels = Table.SelectRows(Source, each [Price_Status] = "Valid" and List.Contains(CoreFuels, [Fuel_Type])),
    KeptNeededColumns = Table.SelectColumns(KeptValidCoreFuels, {"SiteId", "Fuel_Type", "Price_Date", "Transaction_AEST", "Price_cpl"}),
    // Timestamps are to the minute: if two different prices share the last minute of the day,
    // keep the lower one (usually a typo corrected seconds later, e.g. 495.9c then 195.9c)
    LastPriceOfDay = Table.Group(KeptNeededColumns, {"SiteId", "Fuel_Type", "Price_Date"}, {
        {"Price_cpl", each
            let LastTime = List.Max([Transaction_AEST])
            in List.Min(Table.SelectRows(_, each [Transaction_AEST] = LastTime)[Price_cpl]), type number}}),
    GroupedBySeries = Table.Group(LastPriceOfDay, {"SiteId", "Fuel_Type"}, {
        {"Windows", each
            let
                Sorted = Table.Sort(_, {{"Price_Date", Order.Ascending}}),
                Starts = Sorted[Price_Date],
                NextStarts = List.Skip(Starts, 1) & {Date.AddDays(PeriodEnd, 1)}
            in
                Table.FromColumns({Starts, Sorted[Price_cpl], NextStarts}, {"Valid_From", "Price_cpl", "Next_Change"}),
         type table [Valid_From = date, Price_cpl = number, Next_Change = date]}}),
    ExpandedWindows = Table.ExpandTableColumn(GroupedBySeries, "Windows", {"Valid_From", "Price_cpl", "Next_Change"}),
    AddedDaysInEffect = Table.AddColumn(ExpandedWindows, "Days_In_Effect", each
        Duration.Days([Next_Change] - [Valid_From]), Int64.Type),
    RemovedNextChange = Table.RemoveColumns(AddedDaysInEffect, {"Next_Change"})
in
    RemovedNextChange
