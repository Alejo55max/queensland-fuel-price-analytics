// Settings — reads the editable tblSettings table on the Settings sheet
// and turns it into a record, so every threshold used in cleaning is visible in Excel.
let
    Source = Excel.CurrentWorkbook(){[Name = "tblSettings"]}[Content],
    AsRecord = Record.FromList(Source[Value], Source[Setting])
in
    AsRecord
