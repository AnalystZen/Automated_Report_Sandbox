Attribute VB_Name = "Mod_Imports"
'// List all procedures related to importing data for this project in this module.
Option Private Module

Sub Sp_CcoCoidImport()
'// This macro will import sap coid into the hidded coid sheet.
'// Created by "" on 7/26/2024. R-08/26/2024. R- 10/7/2024
    
    '// Declare variable.
    Dim DateEntry As String
    Dim FileDate As String
    Dim CoidImport As Range
    
    '// Assign values.
    DateEntry = Range("DateEntry").Value
    FileDate = Fn_FormatDate(Range("DateEntry"))
    Set CoidImport = ThisWorkbook.Sheets("COID").Range("B4:G100")
    
    '// Conditional check to run macro.
    If FileDate = False Then Exit Sub
    
    '// Error label
    On Error GoTo ErrHandler:

    '// Run Coid Import
    Sap_CoidImport DateEntry
      
    '// Clear old data from active sheet and unhide coid sheet.
    With ThisWorkbook
        .Sheets(FileDate).Range("ProcessOrders").Resize(EkrClearRange, EpoProdProfile).ClearContents
        .Sheets("COID").Visible = True
    End With
    
    '// Clears old data from columns and pastes sap COID.
    With ThisWorkbook.Sheets("COID")
        .Select
        .Cells.ClearContents
        .Range("A1").Select
        .Paste
    End With
    
    '// Format coid with | delimiter. This installs Coid into columns.
    Columns("A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
 
    '// Copy coid data.
    With CoidImport
        Application.Intersect(.SpecialCells(xlCellTypeVisible), _
        .SpecialCells(xlCellTypeConstants)).Copy
    End With

    '// Paste coid data into active sheet.
    With ThisWorkbook.Sheets(FileDate)
        .Range("ProcessOrders").PasteSpecial xlPasteValues
        .Activate
        .Range("A1").Select
    End With

    '// Hide coid paste sheet.
    ThisWorkbook.Sheets("COID").Visible = False
    
    '// Clean exit.
    Exit Sub
     
ErrHandler:
    '// Sap custom error flag
    Err.Raise Number:=E_ErrorCustom.EecErrSapGui
    
End Sub

Sub Sp_ImportCoidCases()
'// Imports the COID Case Commits. COID that is opened is for the biscuit side. It also has cookie information.
'// Created by "" on 4/20/2024. R-08/26/2024. R - 10/7/2024.

    '// Declare variable.
    Dim DateEntry As String
    Dim FileDate As String
    
    '// Assign values.
    DateEntry = Range("DateEntry").Value
    FileDate = Fn_FormatDate(Range("DateEntry"))

    '// Conditional check to run macro.
    If FileDate = False Then Exit Sub
    
    '// Error label.
    On Error GoTo ErrHandler:
    
    '// Copy po data for sap usage.
    With Sheets(FileDate).Range("ProcessOrders").Resize(EkrClearRange)
        Application.Intersect(.SpecialCells(xlCellTypeVisible), _
        .SpecialCells(xlCellTypeConstants)).Copy
    End With
     
    '// Get Sap case information for process orders
    Sap_CaseImport DateEntry

    '// Clears old data from columns and pastes COID
    With Sheets(FileDate).Range("CoidImport")
        .EntireColumn.Resize(, 16).Clear
        .Select
        .PasteSpecial
    End With
    
    '// Format coid with | delimiter. This installs Coid into columns.
    Range("CoidImport").EntireColumn.Select
    Selection.TextToColumns Destination:=Range("CoidImport").EntireColumn.Resize(1), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
        
    '// Selects sheet macro was run on.
    Sheets(FileDate).Select
    
    '// Clean exit
    Exit Sub

ErrHandler:
    '// Raise sap custom error.
    Err.Raise E_ErrorCustom.EecErrImport

End Sub

Sub Sp_ImportPrismaCommits()
'// Sub imports the minimint commits report. Report must be exported from Minimint first.
'// Created by "" on 4/20/2024. R-08/26/2024 . R- 10/7/2024

    '// Declare variable.
    Dim DateEntry As String
    Dim FileDate As String
    Dim WbPrisma As Workbook
    
    '// Assign values.
    DateEntry = Range("DateEntry").Value
    FileDate = Fn_FormatDate(Range("DateEntry"))

    '// Conditional check to run macro.
    If FileDate = False Then Exit Sub
       
    '// Error label.
    On Error GoTo ErrHandler:
  
    '// Clears the columns of old data for new data.
    With Sheets(FileDate).Range("PrismaImport")
        .EntireColumn.Resize(, 42).Clear
    End With
 
    '// Opens the source workbook of the exported MM report.
    Set WbPrisma = Workbooks.Open(Filename:=Gl_PrismaWorkBook + FileDate + ".xls", UpdateLinks:=3, ReadOnly:=True)

    '// Selects data and copies.
    With WbPrisma.Worksheets(1).Range("A1:AP300")
    .Copy
    End With
    
    '// Opens this workboook to paste the data that was retrieved.Data has merged columns.
    ThisWorkbook.Activate
    Range("PrismaImport").Select
    Selection.PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    ActiveSheet.Paste
    
    '// Closes the source data workbook.
    Windows(FileDate + ".xls").Close
    Sheets(FileDate).Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
       
    '// Clean Exit
    Exit Sub

ErrHandler:
    '// User notice of failure.
    Err.Raise E_ErrorCustom.EecErrMmImport
    
End Sub

Sub Sp_ImportSapMixCommits()
'// This macro imports mixes and dumps committed to SAP. It uses the ALMix layout in SAP for the information.
'// Created by "" on 5/1/2024. R-08/26/2024. R- 10/07/2024.

    '// Declare variable.
    Dim DateEntry As String
    Dim FileDate As String
    
    '// Assign values.
    DateEntry = Range("DateEntry").Value
    FileDate = Fn_FormatDate(Range("DateEntry"))

    '// Conditional check to run macro.
    If FileDate = False Then Exit Sub
    
    '// Enable error handle
    On Error GoTo ErrHandler:
     
    '// Copy po data for sap usage.
    With Sheets(FileDate).Range("ProcessOrders").Resize(EkrClearRange)
        Application.Intersect(.SpecialCells(xlCellTypeVisible), _
        .SpecialCells(xlCellTypeConstants)).Copy
    End With
        
    '// Import Sap Mixes
    Sap_MixImport DateEntry
    
    '// Clears old data from columns and pastes COID
    With Sheets(FileDate).Range("SapMixImport")
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
      
    '// Selects the sheet the macro was run on.
     Sheets(FileDate).Select
          
    '// Clean exit
    Exit Sub
    
ErrHandler:
    '// User update of failure
    Err.Raise E_ErrorCustom.EecErrSapGui
    
End Sub

Sub Sp_ImportShiftReport()
'// This macro imports the shift reports from Goldfish. All three shifts information is copied and pasted into this workbook.
'// Created by "" on 5/2/2024. R-08/26/2024. R-10/07/2024
 
    Dim DateEntry As String
    Dim FileDate As String
    Dim SheetDate As String
    Dim CopyRange As String
    Dim Wrkbook As String
    Dim SourceBook As Workbook
   
    DateEntry = Range("DateEntry").Value
    FileDate = Fn_FormatDate(Range("DateEntry"))
    SheetDate = Format(DateEntry, "m-d-yy")
    Wrkbook = Gl_CcoWorkBook + SheetDate + ".xlsx"
    SheetDate = Format(DateEntry, "yyyymmdd")
    CopyRange = "A10:I27"
    
    '// Conditional check to run macro.
    If FileDate = False Then Exit Sub
    
    '// Enable error handle
    On Error GoTo ErrHandler:
           
    '// Clears previous data from worksheet.
    Range("NoShiftImport").EntireColumn.Resize(, 15).Clear
       
    '// Opens the goldfish reports based on the date of the worksheet. Starts with NO.
    Set SourceBook = Workbooks.Open(Filename:=Wrkbook, UpdateLinks:=3, ReadOnly:=True)
   
Start:
    '// Opens the goldfish reports based on the date of the worksheet. Starts with NO.
    Worksheets("NO").Select
    Range(CopyRange).Select
    Selection.Copy
    '// Pastes NO data into this workbook.
    ThisWorkbook.Activate
    Range("NoShiftImport").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
    '// Opens the goldfish reports based on the date of the worksheet. Starts with AM.
    SourceBook.Activate
    Worksheets("AM").Select
    Range(CopyRange).Select
    Selection.Copy
    '// Pastes AM data into this workbook.
    ThisWorkbook.Activate
    Range("AmShiftImport").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False
    
    '// Opens the goldfish reports based on the date of the worksheet. Starts with PM.
    SourceBook.Activate
    Worksheets("PM").Select
    Range(CopyRange).Select
    Selection.Copy
    '// Pastes PM data into this workbook.
    ThisWorkbook.Activate
    Range("PmShiftImport").Select
    Selection.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=False, Transpose:=False

    '// Closes the source data workbook.
    SourceBook.Close
    
    '// Select sheet sub was run from.
    ThisWorkbook.Sheets(FileDate).Select
    Application.GoTo Reference:=Range("A1"), Scroll:=True
    
    '// Get offshift mixes from reports.
    Sp_CalculateShiftMixes
    
    '// Clean exit
    Exit Sub
    
ErrHandler:
    '// If workbook not found let user select or exit on cancel.
    If Err = "1004" Then
        Dim FileToOpen As Variant
        ChDrive "G:"
        ChDir Gl_CcoWorkBook
        FileToOpen = Application.GetOpenFilename(filefilter:="Excel Files(*.xls*),*xls*", Title:="PLEASE SELECT THE CORRECT FILE")
    If FileToOpen <> False Then
        Set SourceBook = Application.Workbooks.Open(FileToOpen, UpdateLinks:=3, ReadOnly:=True)
        GoTo Start:
        ElseIf FileToOpen = False Then
    End If
    End If

    '//User update of error.
    Err.Raise E_ErrorCustom.EecErrImport, , Err.Description
    
End Sub

Sub Sp_ImportProdReport()
'// This sub will import the prod report from sap for NO and Am and PM.
'// Created by "" on 10/05/2024

    '// Variable declarations.
    Dim DateEntry As String
    Dim FileDate As String
    
    '// Assign values.
    DateEntry = Range("DateEntry").Value
    FileDate = Fn_FormatDate(Range("DateEntry"))
    
    '// Conditional check to run macro.
    If FileDate = False Then Exit Sub
    
    '// Enable error handle
     On Error GoTo ErrHandler:
    
    '// Clear old data.
    Range("NoCaseImport").EntireColumn.Resize(, 18).Clear
     
    '// Get prod report for shifts.
    Sap_ProdReport DateEntry, FileDate, "00:00", "07:30", "NoCaseImport"
     
    '// Selects the sheet the macro was run on.
     Sheets(FileDate).Select
           
    '// Clean exit
    Exit Sub
    
ErrHandler:
    '// User update of failure
    Err.Raise E_ErrorCustom.EecErrImport
    
End Sub
