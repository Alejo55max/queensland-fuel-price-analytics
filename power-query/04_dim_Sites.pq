// dim_Sites — one row per fuel station, using its most recent details
// (some sites changed brand or name during the year, e.g. Caltex -> Ampol rebrands).
// Distance to the Brisbane CBD uses the haversine formula; Region bands come from Settings.
let
    Source = Raw_Combined,
    RemovedBlankRows = Table.SelectRows(Source, each [SiteId] <> null and Text.Trim([SiteId]) <> ""),
    KeptSiteColumns = Table.SelectColumns(RemovedBlankRows, {"SiteId", "Site_Name", "Site_Brand", "Site_Suburb",
        "Site_Post_Code", "Site_Latitude", "Site_Longitude", "TransactionDateutc"}),
    TrimmedText = Table.TransformColumns(KeptSiteColumns,
        {{"Site_Name", Text.Trim, type text}, {"Site_Brand", Text.Trim, type text}, {"Site_Suburb", Text.Trim, type text}}),
    TypedColumns = Table.TransformColumnTypes(TrimmedText,
        {{"SiteId", type text}, {"Site_Latitude", type number}, {"Site_Longitude", type number}}, "en-US"),
    ParsedTimestamp = Table.AddColumn(TypedColumns, "Seen_UTC", each
        DateTime.FromText([TransactionDateutc], [Format = "d/M/yyyy H:mm", Culture = "en-AU"]), type datetime),
    // Step A: each distinct version of a site's details, with the last time it was seen (a few thousand rows)
    SiteVersions = Table.Group(ParsedTimestamp,
        {"SiteId", "Site_Name", "Site_Brand", "Site_Suburb", "Site_Post_Code", "Site_Latitude", "Site_Longitude"},
        {{"Last_Seen", each List.Max([Seen_UTC]), type datetime}}),
    // Step B: keep the most recent version per site and count how many brands it traded under
    GroupedBySite = Table.Group(SiteVersions, {"SiteId"}, {
        {"Latest", each Table.Max(_, "Last_Seen"), type record},
        {"Brands_Seen", each List.Count(List.Distinct([Site_Brand])), Int64.Type}}),
    ExpandedLatest = Table.ExpandRecordColumn(GroupedBySite, "Latest",
        {"Site_Name", "Site_Brand", "Site_Suburb", "Site_Post_Code", "Site_Latitude", "Site_Longitude"}),
    RenamedBrand = Table.RenameColumns(ExpandedLatest, {{"Site_Brand", "Brand"}}),
    AddedKmFromCBD = Table.AddColumn(RenamedBrand, "Km_From_CBD", each
        let
            ToRad = Number.PI / 180,
            Lat1 = Settings[CBD_Latitude] * ToRad, Lon1 = Settings[CBD_Longitude] * ToRad,
            Lat2 = [Site_Latitude] * ToRad, Lon2 = [Site_Longitude] * ToRad,
            A = Number.Power(Number.Sin((Lat2 - Lat1) / 2), 2)
                + Number.Cos(Lat1) * Number.Cos(Lat2) * Number.Power(Number.Sin((Lon2 - Lon1) / 2), 2)
        in
            Number.Round(2 * 6371 * Number.Asin(Number.Sqrt(A)), 1), type number),
    AddedRegion = Table.AddColumn(AddedKmFromCBD, "Region", each
        if [Km_From_CBD] <= Settings[Brisbane_Max_Km] then "Brisbane"
        else if [Km_From_CBD] <= Settings[SEQ_Max_Km] then "Rest of SEQ"
        else "Regional QLD", type text),
    FlaggedBrandChange = Table.AddColumn(AddedRegion, "Brand_Changed", each [Brands_Seen] > 1, type logical),
    // Buffered: the site list (about 1,800 rows) is used twice below; without buffering,
    // Power Query would re-read and re-group all 900k source rows for each use
    RemovedHelper = Table.Buffer(Table.RemoveColumns(FlaggedBrandChange, {"Brands_Seen"})),
    // Long tail of small brands grouped as "Other brands" so charts and slicers stay readable
    SitesPerBrand = Table.Sort(Table.Group(RemovedHelper, {"Brand"}, {{"Sites", each Table.RowCount(_), Int64.Type}}),
        {{"Sites", Order.Descending}, {"Brand", Order.Ascending}}),
    // Buffered so the list is computed once, not once per site in the step below
    TopBrands = List.Buffer(List.FirstN(SitesPerBrand[Brand], Settings[Top_Brands])),
    AddedBrandGroup = Table.AddColumn(RemovedHelper, "Brand_Group", each
        if List.Contains(TopBrands, [Brand]) then [Brand] else "Other brands", type text),
    SortedSites = Table.Sort(AddedBrandGroup, {{"SiteId", Order.Ascending}})
in
    SortedSites
