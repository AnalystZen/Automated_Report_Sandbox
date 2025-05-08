Attribute VB_Name = "Mod_ProjectMaint"
'// List all procedures related to project maintenance here.
Option Private Module

Sub Sp_CreateSheets()
'// This macro will create X new master worksheets and date them. Worksheet names are automatically updated and cell B7 is also updated with sheet date.
'// Created by "" on 5/1/2024. R-08/26/2024
    
    '// Declare variables.
    Dim DateEntry As Date
    Dim SheetDate As String
    Dim NextDate As Date
    Dim X As Integer
    Dim ShCount As Integer
    
    '// Assign values
    On Error Resume Next
    DateEntry = Range("DateEntry").Value
    SheetDate = Format(DateEntry, "yyyymmdd")
    X = 1
    ShCount = InputBox("Please Insert The Amount Of New Sheets To Create.", "Create New WorkSheets", "Insert A Number") * 1
    On Error GoTo 0
    
    '// Conditional checks to run macro.
    If Range("DateEntry").Value = "" Then
            MsgBox ("Please insert a date into cell B7 before running Macro.")
            Exit Sub
        ElseIf ShCount = 0 Then
            MsgBox Prompt:="An amount of sheets to create was not selected! Try again.", Buttons:=vbExclamation + vbOKCancel, Title:="Sheet Amount"
            Exit Sub
        ElseIf MsgBox("This action will create additional dated worksheets." & vbNewLine & vbNewLine & "Is this procedure being run on the last dated worksheet?", vbExclamation + vbYesNo) = vbNo Then
            Exit Sub
    End If
    
    '// Enable Error handle.
    On Error GoTo ErrHandler:

    '// Unprotect master worksheet to make copies.
    Sp_UnprotectMaster ConSheetPassword
    
    '// Loop used to create new master worksheets.
Start:
    For X = 1 To ShCount
        NextDate = DateEntry + X
        Sheets("Master Worksheet").Select
        Worksheets("Master Worksheet").Copy Before:=Sheets("Master Worksheet")
        ActiveSheet.Range("DateEntry").Value = NextDate
        ActiveSheet.Name = Format(NextDate, "yyyymmdd")
        ActiveSheet.Range("DateEntry").Select
    Next

    '// Protect master worksheet.
    Sp_ProtectMaster ConSheetPassword

    '// Selects the original worksheet that the macro was ran on.
    Sheets(SheetDate).Select

    '// Clean exit
    Exit Sub

ErrHandler:
    '// Loop restart after error.
    If Err.Description = "That name is already taken. Try a different one." Then
        ActiveSheet.Range("DateEntry").Value = NextDate + 1
        ActiveSheet.Name = Format(NextDate, "yyyymmdd") + Worksheets.Count
        DateEntry = NextDate + 1
        On Error GoTo -1
        GoTo Start:
    End If
    
    '// User update of failure.
    Err.Raise E_ErrorCustom.EecErrGeneral
    
End Sub

Sub Sp_ArchiveSheets()
'// This macro will archive worksheets by saving them to the designated archive workbook. Sheet index numbers are used.
'// Created by "" on 7/13/2024. R-08/24/2024

    '// Variable declarations.
    Dim ShStart As Worksheet
    Dim ShIndexNumber As Long
    Dim ShCycle As Long
    Dim ShCount As Long
    Dim Wrkbook As String
    Dim WrkBookName As String
    Dim MessageResult As VbMsgBoxResult
    
    '// Assign values
    Set ShStart = ThisWorkbook.ActiveSheet
    ShIndexNumber = ThisWorkbook.ActiveSheet.Index - 1
    ShCycle = ThisWorkbook.Sheets(1).Index
    ShCount = 0
    Wrkbook = Gl_ArchiveWorkbook
    WrkBookName = Mid(Wrkbook, InStr(40, Wrkbook, "\") + 1)
    MessageResult = MsgBox(Prompt:="This macro will archive all previous sheets!" & vbNewLine & "Continue?", _
    Buttons:=vbCritical + vbYesNoCancel, Title:="Archive Sheets")
    
    '// Conditional checks to run macro
    If ThisWorkbook.ActiveSheet.Name = "Master Worksheet" Then
            MsgBox Prompt:="No Valid Worksheet Selected!", Buttons:=vbCritical, Title:="Select A Valid Worksheet"
            Exit Sub
        ElseIf MessageResult = vbNo Then
            MsgBox Prompt:="Routine Canceled!", Buttons:=vbInformation, Title:="Canceled"
            Exit Sub
        ElseIf MessageResult = vbCancel Then
            MsgBox Prompt:="Routine Canceled!", Buttons:=vbInformation, Title:="Canceled"
            Exit Sub
    End If
    
    '// Error label.
    On Error GoTo ErrHandler:

    '// Opens archive workbook.
    Workbooks.Open Filename:=Wrkbook, UpdateLinks:=3, ReadOnly:=False
    
    '// Start of loop using sheet index numbers to archive.
Start:
    Do While ShCount <> ShIndexNumber
            If ShCount > ShIndexNumber Then Exit Do
        With ThisWorkbook.Sheets(ShCycle)
            .Activate
            .Cells.Copy
            .Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range("A1").Select
            .Move After:=Workbooks(WrkBookName).Sheets(Workbooks(WrkBookName).Sheets.Count)
        End With
        ShCount = ShCount + 1
    Loop

    '// Selects sheet macro was run on.
    ShStart.Activate
        
    '// Close the archive workbook and saves.
    Workbooks(WrkBookName).Close SaveChanges:=True
    
    '// Clean exit.
    Exit Sub
   
ErrHandler:
    '// Loop restart after error.
    If Err.Description = "Subscript out of range" Then
        ShCycle = ShCycle + 1
        On Error GoTo -1
        GoTo Start:
    End If
    
    '// User notice of failure.
    Err.Raise E_ErrorCustom.EecErrGeneral
End Sub

Sub Sp_RemoveFormulae()
'// This macro will remove formulas from the workrsheets via a loop. Uses sheet index numbers.
'// Created by "" on 7/13/2024. R-08/28/2024

    '// Variable declarations.
    Dim ShStart As Worksheet
    Dim ShIndexNumber As Long
    Dim ShCycle As Long
    Dim MessageResult As VbMsgBoxResult
    
    Set ShStart = ThisWorkbook.ActiveSheet
    ShIndexNumber = ThisWorkbook.ActiveSheet.Index
    ShCycle = ThisWorkbook.Sheets(1).Index
    MessageResult = MsgBox(Prompt:="This macro will remove formulas from all previous sheets!" & vbNewLine & "Continue?", _
    Buttons:=vbCritical + vbYesNoCancel, Title:="Archive Sheets")
    
    '// Conditional checks to run macro
    If ThisWorkbook.ActiveSheet.Name = "Master Worksheet" Then
            MsgBox Prompt:="No Valid Worksheet Selected!", Buttons:=vbCritical, Title:="Select A Valid Worksheet"
            Exit Sub
        ElseIf MessageResult = vbNo Then
            MsgBox Prompt:="Routine Canceled!", Buttons:=vbInformation, Title:="Canceled"
            Exit Sub
        ElseIf MessageResult = vbCancel Then
            MsgBox Prompt:="Routine Canceled!", Buttons:=vbInformation, Title:="Canceled"
            Exit Sub
    End If
    
    '// Enable error trapping.
    On Error GoTo ErrHandler:
        
'// Start of loop using sheet index numbers to archive.
Start:
    Do While ShCycle <> ShIndexNumber
            If ShCycle > ShIndexNumber Then Exit Do
        With ThisWorkbook.Sheets(ShCycle)
            .Activate
            .Cells.Copy
            .Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            .Range("A1").Select
        End With
        ShCycle = ShCycle + 1
    Loop

    '// Selects sheet macro was run on.
    ShStart.Activate

    '// Clean exit.
    Exit Sub
   
ErrHandler:
    '// Loop restart after error.
    If Err.Description = "Subscript out of range" Then
        ShCycle = ShCycle + 1
        On Error GoTo -1
        GoTo Start:
    End If

'// User notice of failure.
    Err.Raise E_ErrorCustom.EecErrGeneral
    
End Sub


