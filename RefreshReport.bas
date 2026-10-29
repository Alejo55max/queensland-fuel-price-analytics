Attribute VB_Name = "RefreshReport"
Option Explicit

' Refreshes every Power Query query and PivotTable in the right order,
' then stamps the refresh time on the Cover sheet.
' Linked to the "Refresh report" button on the Dashboard.
Public Sub RefreshReport()
    Dim conn As WorkbookConnection
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim started As Double

    started = Timer
    Application.ScreenUpdating = False
    Application.StatusBar = "Refreshing Power Query tables..."

    ' Wait for each query to finish before the next one starts,
    ' so the PivotTables never read half-loaded data
    For Each conn In ThisWorkbook.Connections
        If conn.Type = xlConnectionTypeOLEDB Then
            conn.OLEDBConnection.BackgroundQuery = False
            conn.Refresh
        End If
    Next conn

    Application.StatusBar = "Refreshing PivotTables..."
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.PivotCache.Refresh
        Next pt
    Next ws

    Application.CalculateFull
    ThisWorkbook.Worksheets("Cover").Range("C8").Value = Now

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Report refreshed in " & Format(Timer - started, "0") & " seconds.", vbInformation, "Refresh report"
End Sub

' Saves the Dashboard sheet as a PDF next to the workbook, named with today's date.
' Linked to the "Export dashboard to PDF" button on the Dashboard.
Public Sub ExportDashboardPDF()
    Dim target As String

    target = ThisWorkbook.Path & Application.PathSeparator & _
             "QLD_fuel_dashboard_" & Format(Date, "yyyy-mm-dd") & ".pdf"

    With ThisWorkbook.Worksheets("Dashboard").PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    ThisWorkbook.Worksheets("Dashboard").ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=target, Quality:=xlQualityStandard, OpenAfterPublish:=False

    MsgBox "Dashboard saved to:" & vbNewLine & target, vbInformation, "Export dashboard"
End Sub
