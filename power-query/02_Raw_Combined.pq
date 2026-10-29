// Raw_Combined — combines every monthly CSV in the data folder into one table.
// Cleaning step 1 (combine files) and step 2 (standardise headers: the March 2026
// file uses spaces in column names, e.g. "Site Brand" instead of "Site_Brand").
let
    Source = Folder.Files(Settings[DataFolder]),
    CsvFilesOnly = Table.SelectRows(Source, each Text.Lower([Extension]) = ".csv"),
    ParsedCsv = Table.AddColumn(CsvFilesOnly, "Data", each
        Table.PromoteHeaders(
            Csv.Document([Content], [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
            [PromoteAllScalars = true])),
    FlaggedSpacedHeaders = Table.AddColumn(ParsedCsv, "Had_Spaced_Headers", each
        List.AnyTrue(List.Transform(Table.ColumnNames([Data]), each Text.Contains(_, " "))), type logical),
    StandardisedHeaders = Table.TransformColumns(FlaggedSpacedHeaders, {"Data", each
        Table.TransformColumnNames(_, each Text.Replace(Text.Trim(_), " ", "_"))}),
    KeptColumns = Table.SelectColumns(StandardisedHeaders, {"Name", "Had_Spaced_Headers", "Data"}),
    ExpandedRows = Table.ExpandTableColumn(KeptColumns, "Data",
        {"SiteId", "Site_Name", "Site_Brand", "Sites_Address_Line_1", "Site_Suburb", "Site_State",
         "Site_Post_Code", "Site_Latitude", "Site_Longitude", "Fuel_Type", "Price", "TransactionDateutc"}),
    RenamedSourceFile = Table.RenameColumns(ExpandedRows, {{"Name", "Source_File"}})
in
    RenamedSourceFile
