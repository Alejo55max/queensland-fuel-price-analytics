// Prices_Clean — typed, de-duplicated price events with a status flag on every row.
// Cleaning steps 3-7: remove blank rows (most monthly files end with an empty line),
// parse UTC timestamps and convert to Brisbane time (AEST, UTC+10, Queensland has no
// daylight saving), remove duplicate events, flag the 9999 "temporarily unavailable"
// code as a stock-out, and flag prices outside the valid range.
// Only the 4 columns needed for pricing are kept: site details live in dim_Sites.
// Keeping rows narrow lets Power Query group and de-duplicate 900k rows in memory.
let
    Source = Raw_Combined,
    RemovedBlankRows = Table.SelectRows(Source, each [SiteId] <> null and Text.Trim([SiteId]) <> ""),
    KeptPriceColumns = Table.SelectColumns(RemovedBlankRows, {"SiteId", "Fuel_Type", "Price", "TransactionDateutc"}),
    TrimmedFuelType = Table.TransformColumns(KeptPriceColumns, {{"Fuel_Type", Text.Trim, type text}}),
    TypedColumns = Table.TransformColumnTypes(TrimmedFuelType, {{"SiteId", type text}, {"Price", Int64.Type}}, "en-US"),
    ParsedUtcTimestamp = Table.AddColumn(TypedColumns, "Transaction_UTC", each
        DateTime.FromText([TransactionDateutc], [Format = "d/M/yyyy H:mm", Culture = "en-AU"]), type datetime),
    ConvertedToAEST = Table.AddColumn(ParsedUtcTimestamp, "Transaction_AEST", each
        [Transaction_UTC] + #duration(0, Settings[UTC_Offset_Hours], 0, 0), type datetime),
    AddedPriceDate = Table.AddColumn(ConvertedToAEST, "Price_Date", each Date.From([Transaction_AEST]), type date),
    RemovedRawTimestamp = Table.RemoveColumns(AddedPriceDate, {"TransactionDateutc", "Transaction_UTC"}),
    RemovedDuplicateEvents = Table.Distinct(RemovedRawTimestamp, {"SiteId", "Fuel_Type", "Price", "Transaction_AEST"}),
    AddedPriceCentsPerLitre = Table.AddColumn(RemovedDuplicateEvents, "Price_cpl", each [Price] / 10, type number),
    FlaggedPriceStatus = Table.AddColumn(AddedPriceCentsPerLitre, "Price_Status", each
        if [Price] = Settings[Stockout_Code] then "Stock-out"
        else if [Price_cpl] < Settings[Min_Valid_cpl] or [Price_cpl] > Settings[Max_Valid_cpl] then "Out of range"
        else "Valid", type text)
in
    FlaggedPriceStatus
