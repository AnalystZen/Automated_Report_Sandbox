Attribute VB_Name = "Mod_Exports"
'// List all procedures related to exporting of data and report creation related here.
Option Private Module

Sub Sp_CreateDailyReport()
'// This sub will create a daily cco report and summarizes the data. It also formats the report.
'// Created by "" on 4/15/2024. R-08/25/2024.
    
    '// Declare variables
    Dim CellRow As String
    Dim BlockEnd As String
    Dim Add1 As Long
    Dim Add2 As Long
    Dim LastRow As String
    Dim PrintImg As String
    Dim DailyReport As Workbook
    Dim SheetDate As String
    
    '// Assign values
    PrintImg = Gl_PrintImgPoint
    CellRow = 9
    BlockEnd = 0
    SheetDate = ActiveSheet.Name
    
    '// Obtains last row number.
    Do While BlockEnd <> 1
        If Range("A" & CellRow) = "" Then
            BlockEnd = 1
        Else
            CellRow = CellRow + 1
        End If
    Loop

    '// Enables error handle
    On Error GoTo ErrHandler:
    
    '// Copies the data on the fish report.
    Range("BE3:BO" & CellRow).Select
    Range("BE3:BO" & CellRow).Copy
    Range("A1").Select
    
    '// Adds a new workbook and pastes the copied data.
    Set DailyReport = Workbooks.Add
    ActiveSheet.Paste
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        
    '// Format The Production Report.
    Range("A1:K4").Copy
    Range("A7").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Selection.ClearContents
    Application.CutCopyMode = False
    
    ActiveCell.Offset(-1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "TOTALS  :"
    LastRow = ActiveCell.Row
    '// Installs formulas into report with add variables
    ActiveCell.Offset(0, 4).Range("A1").Select
    ActiveCell = WorksheetFunction.Sum(Range("E7", Range("E7").End(xlDown)))
    Add1 = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell = WorksheetFunction.Sum(Range("F7", Range("F7").End(xlDown)))
    Add2 = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell = WorksheetFunction.Sum(Range("G7", Range("G7").End(xlDown)))
   
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell = WorksheetFunction.Sum(Range("H7", Range("H7").End(xlDown)))
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell = WorksheetFunction.Sum(Range("I7", Range("I7").End(xlDown)))
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell = WorksheetFunction.Sum(Range("J7", Range("J7").End(xlDown)))
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell = Add2 / Add1 * 100 & "%"
    
    Range("A" & LastRow, "K" & LastRow).Font.Bold = True
    Range("A" & LastRow, "K" & LastRow).Font.Size = 12
    
    '// Adjusts the column width of the pasted data and row height.
    Selection.Columns.AutoFit
    Columns("A:A").ColumnWidth = 11.5
    Columns("B:B").ColumnWidth = 15
    Columns("C:C").ColumnWidth = 50
    Columns("D:I").ColumnWidth = 11
    Columns("J:J").ColumnWidth = 12.5
    Columns("K:K").ColumnWidth = 8
    Rows("1:60").Select
    Selection.RowHeight = 25
    Range("A1").Select
        
    '// Printer set up and page setup.
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        '// Picture selection variable for header.
        On Error Resume Next
        .CenterHeaderPicture.Filename = PrintImg
        On Error GoTo 0
        .CenterHeader = "&G"
        .CenterFooter = "&D" & Chr(10) & "&T"
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(1.5)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Application.PrintCommunication = True

    '// Maximizes the report window.
    Application.WindowState = xlMaximized
    Range("K12").Select
    
    '// Create Report Pdf
    CreateReportPdfs "Daily", DailyReport, SheetDate
    
    '// Clean exit
    Exit Sub
    
ErrHandler:
    '// User notification of failure.
    Err.Raise E_ErrorCustom.EecErrExport
End Sub

Sub Sp_CreateWeeklyMixReport()
'// This macro will create a weekly report with Monday being the start of production. It will also import SAP mixes.
'// Created by "" on 6/20/2024. R-08/25/2024

    '// Declare variables.
    Dim DateEntry As Date
    Dim FirstDayinWeek As Variant
    Dim LastDayinWeek As Variant
    Dim FirstSheet As String
    Dim LastSheet As String
    Dim LastRow As Long
    Dim PrintImg As String
    Dim X As Integer
    Dim WsToDelete As Worksheet

    '// Assign variables
    DateEntry = Range("DateEntry").Value
    FirstDayinWeek = DateEntry - Weekday(DateEntry, vbMonday) + 1
    LastDayinWeek = DateEntry
    FirstSheet = Format(FirstDayinWeek, "yyyymmdd")
    LastSheet = Format(LastDayinWeek, "yyyymmdd")
    LastRow = Range("ProcessOrders", Range("ProcessOrders")).End(xlDown).Row
    PrintImg = Gl_PrintImgPoint
    X = 1
    
    '// Conditonal checks to run sub.
    If Range("DateEntry").Value = "" Then
        MsgBox Prompt:="Please insert a valid date and try again.", Buttons:=vbExclamation + vbOKOnly, Title:="Insert Date"
        Exit Sub
    ElseIf MsgBox(Prompt:="This macro will filter mix totals from start date: Monday to the current date within a Week." _
        & vbNewLine & vbNewLine & "Do you want to continue?", Buttons:=vbYesNo) = vbNo Then
        Exit Sub
    End If

    '// Enable error handle.
    On Error GoTo ErrHandler:
    
    '// Unhides data collection sheet and clears old data.
    With Sheets("Weekly Data")
        .Visible = True
        .Cells.ClearContents
    End With
    
    '// Start of data collection loop.
Start:
    Do Until FirstSheet = LastSheet + X
        Sheets(FirstSheet).Activate
        Range("ProcessOrders").Resize(EkrClearRange, 29).Copy
        Range("A1").Select
        Sheets("Weekly Data").Activate
        
        If Range("A9").Value > 0 Then
            Range("A9").End(xlDown).Offset(1).Select
        Else
            Range("A9").Select
        End If
        
        ActiveCell.PasteSpecial (xlPasteValues)
        FirstSheet = FirstSheet + X
    Loop

    '// Start of filter data loop.
    Dim firstcell As Long
    Dim lastcell As Long

    firstcell = Cells(9, "C").Row
    lastcell = Cells(Rows.Count, "C").End(xlUp).Row
    Sheets("Weekly Data").Activate
    
    Do Until lastcell = -50
        If InStr(Range("C" & firstcell).Value, "FISHWIP") < 1 Then
            Rows(firstcell).Delete
        ElseIf InStr(Range("C" & firstcell).Value, "FISHWIP") > 0 Then
            firstcell = firstcell + 1
        End If
            lastcell = lastcell - 1
    Loop

    '// Creates new worksheet from master worksheet and pastes filtered data.
    Sp_UnprotectMaster ConSheetPassword
    ThisWorkbook.Worksheets("Master Worksheet").Copy Before:=Sheets("Master Worksheet")
    Sp_ProtectMaster ConSheetPassword
    
    '// Add new sheet for report.
    Set WsToDelete = ActiveSheet
    WsToDelete.Name = ("Weekly Report") & Worksheets.Count
    
    '// Pastes data into new sheet.
    With Sheets("Weekly Data")
        .Range("A9:G100").Copy Destination:=ActiveSheet.Range("ProcessOrders")
        .Range("AC9:AC100").Copy Destination:=ActiveSheet.Range("AC9")
    End With

    '// Hides data collection sheet.
    With Sheets("Weekly Data")
        .Visible = False
    End With
    
    '// Turn on calculation sap formulas can calculate.
    Application.Calculation = xlCalculationAutomatic
    
    '// Autofill sap formulas for report range.
    WsToDelete.Range("AB9").AutoFill Range("AB9:AB" & LastRow), xlFillDefault
    
    '// Copy po data for sap usage.
    With ActiveSheet.Range("ProcessOrders").Resize(EkrClearRange)
        Application.Intersect(.SpecialCells(xlCellTypeVisible), _
        .SpecialCells(xlCellTypeConstants)).Copy
    End With

    '// Import Sap Mixes.
    Sap_MixImport DateEntry

    '// This clears previous data and pastes new data from SAP.
    With ActiveSheet.Range("SapMixImport")
        .EntireColumn.Resize(, 9).Clear
        .Select
        .PasteSpecial
    End With

    '// Format coid with | delimiter. This installs Coid into columns.
    Range("SapMixImport").EntireColumn.Select
    Selection.TextToColumns Destination:=Range("SapMixImport").EntireColumn.Resize(1), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
    
    '// Format report
    FormatWeeklyReport FirstDayinWeek, Gl_PrintImgPoint
    
    '// Delete base worksheet from project.
    WsToDelete.Delete
        
    '// Clean Exit
    Exit Sub

ErrHandler:
    '// Loop restart after error.
    If Err.Description = "Subscript out of range" Then
        FirstSheet = FirstSheet + 1
        On Error GoTo -1
        GoTo Start:
    End If
    
    '// Custom error message.
    Err.Raise EecErrExport

End Sub

Sub FormatWeeklyReport(ByVal FirstDayinWeek As String, ByVal PrintImg As String)
'// This procedure will format the weekly report is created. It creates report in a new worksheet.
'// Created by "" on 10/08/2024.

    '//Declare variables.
    Dim WbReport As Workbook
    Dim RngHeader As Range
    Dim LastRow As Long

    '// Assign values and create workbook.
    Set RngHeader = ThisWorkbook.ActiveSheet.Range("A5:G8")
    LastRow = Range("ProcessOrders", Range("ProcessOrders")).End(xlDown).Row
    Set WbReport = Workbooks.Add
    
    '// Disable error stops.
     On Error Resume Next
    
    '// Copy header range.
    ThisWorkbook.Activate
    RngHeader.Copy
    WbReport.Worksheets(1).Paste
   
    '// Copy coid information.
    Range("ProcessOrders", Range("ProcessOrders").End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Copy
    
    '// Paste coid information in new book.
    WbReport.Worksheets(1).Activate
    Range("A5").PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    Range("A5").PasteSpecial Paste:=xlPasteValues
    
    '// Copy phase 20 mix information
    ThisWorkbook.Activate
    Range("AB8:AC" & LastRow).Copy
    
    '// Paste mix information in new book.
    WbReport.Worksheets(1).Activate
    Range("H4").PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    Range("H4").PasteSpecial Paste:=xlPasteValues
    
    '// Copy phase 40 mix information
    ThisWorkbook.Activate
    Range("AH8:AH" & LastRow).Copy
    
    '// Paste mix information in new book.
    WbReport.Worksheets(1).Activate
    Range("J4").PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    Range("J4").PasteSpecial Paste:=xlPasteValues
     
    '// Unmerge and remerge new report header.
    WbReport.Worksheets(1).Range("A1:J2").Merge
    WbReport.Worksheets(1).Range("B3:J3").Merge
    
    '// Format the report header titles
    Range("A1") = "IC Weekly Mix Report"
    Range("H4").Value = "SAP20"
    Range("I4").Value = "PRISMA"
    Range("J4").Value = "SAP40"
    Range("A3").Value = "Week Of  :"
    Range("B3").Value = FirstDayinWeek
    Range("E5").EntireColumn.Delete
    Range("A5").End(xlDown).Offset(1).Resize(, 9).Font.Bold = True
    Range("A5").End(xlDown).Offset(1).Resize(, 9).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
    Range("A5").End(xlDown).Offset(1).Resize(, 9).Borders.LineStyle = xlContinuous
    Range("A1:I2").HorizontalAlignment = xlLeft
    Range("B3:I3").HorizontalAlignment = xlLeft
    Range("A1:I2").InsertIndent 25
    Range("B3:I3").InsertIndent 16
    
    '// Install new borders for report.
    With Range("A1:I2").Borders
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With
    
    With Range("B3:I3").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    '// Autofit column widths.
    Columns("A:K").EntireColumn.AutoFit
    
    '// Install formula totals in new report.
    Range("A5").End(xlDown).Offset(1) = "     Totals :"
    Range("A5").End(xlDown).Offset(, 6) = Application.WorksheetFunction.Sum(Range("G5", Range("G5").End(xlDown)))
    Range("A5").End(xlDown).Offset(, 7) = Application.WorksheetFunction.Sum(Range("H5", Range("H5").End(xlDown)))
    Range("A5").End(xlDown).Offset(, 8) = Application.WorksheetFunction.Sum(Range("I5", Range("I5").End(xlDown)))
    
    '// Adjust row height of new report.
    Range("A1", Range("A" & Rows.Count).End(xlUp)).Rows.RowHeight = 20
    
    '// Printer Settings.
    With ActiveSheet.PageSetup
        .CenterHeaderPicture.Filename = PrintImg
        .CenterHeader = "&G"
        .CenterFooter = "&D" & Chr(10) & "&T"
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(1.25)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlPortrait
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Zoom = False
    End With
    
    '// Maximize window of report.
    Application.WindowState = xlMaximized
    Range("A1").Select
    
    '// Create Report Pdf
    CreateReportPdfs "Weekly", WbReport, Format(FirstDayinWeek, "yyyymmdd")
    
    '// Enable normal error handling
    On Error GoTo 0
    
End Sub

Sub CreateReportPdfs(ByVal PdfType As String, ByVal PdfWorkbook As Workbook, Optional ByVal ReportDate As String)
'// This sub will create and save pdfs of the daily or weekly report.
'// Created by "" on 10/25/2024.
    
    '// Verify if admin toggled exports to pdf off.
    If ThisWorkbook.Worksheets("Admin").Range("B4") = "No" Then Exit Sub
    
    '// Declare variables
    Dim PdfPath As String
    
    '// Assign values for daily or weekly report path.
    Select Case PdfType
    
        Case Is = "Daily"
            PdfPath = Gl_PdfDaily
        
        Case Is = "Weekly"
            PdfPath = Gl_PdfWeekly
    
    End Select
    
    '// Export report workbook as pdf.
    PdfWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PdfPath & "GoldFish" & PdfType & ReportDate, Quality:=xlQualityStandard
    
    '// E-mail the created report.
    SendReportEmails PdfType, PdfPath & "GoldFish" & PdfType & ReportDate

End Sub

Sub SendReportEmails(ByVal ReportType As String, ByVal ReportAttachment As String)
'// This sub will send emails of the exported pdf created from daily or weekly reports.
'// Created by "" on 10/25/2024.

    '// Verify if admin allows email of data or  user woul like to emasil reports.
    If ThisWorkbook.Worksheets("Admin").Range("B6") = "No" Then
        Exit Sub
    ElseIf MsgBox("Would you like to e-mail the production report?", vbExclamation + vbYesNo, "E-Mail Report") = vbNo Then
        Exit Sub
    End If
    
    '// Declare variables.
    Dim OutApp As Object
    Dim OutMail As Object
    Dim EmailName As Range
    Dim EmailList As Range

    '// Assign dynamic range
    Set EmailList = ThisWorkbook.Worksheets("Admin").Range("N2").CurrentRegion.Resize(, 1).Offset(1)

    '// Verify what type of report to process
    Select Case ReportType
    
        '// Daily Report
        Case Is = "Daily"
        
        '// For loop through email list range.
        For Each EmailName In EmailList
            If EmailName.Offset(, 1) = "Yes" Then
        
            '// Create Email item
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.Createitem(0)
            
                With OutMail
                    .Display
                    .To = EmailName.Value
                    .CC = ""
                    .BCC = ""
                    .Subject = "Daily GoldFish Report"
                    .Attachments.Add ReportAttachment & ".pdf"
                    .Body = "" & .Body
                    '.Send
                End With
            End If
        Next EmailName
    
        '// Weekly Report
        Case Is = "Weekly"
        
        '// For loop through email list range.
        For Each EmailName In EmailList
            If EmailName.Offset(, 2) = "Yes" Then
        
            '// Create Email item
            Set OutApp = CreateObject("Outlook.Application")
            Set OutMail = OutApp.Createitem(0)
    
                With OutMail
                    .Display
                    .To = EmailName.Value
                    .CC = ""
                    .BCC = ""
                    .Subject = "Weekly GoldFish Report"
                    .Attachments.Add ReportAttachment & ".pdf"
                    .Body = "" & .Body
                    '.Send
                End With
            End If
        Next EmailName
        
    End Select
        
    '// Empty object variables.
    Set OutApp = Nothing
    Set OutMail = Nothing

End Sub
