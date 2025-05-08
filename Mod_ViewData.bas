Attribute VB_Name = "Mod_ViewData"
'// List all procedures related to viewing data here.
Option Private Module

Sub Sp_OpenDailyCoid()
'//This macro will open up the biscuit COID in SAP based on the date in Cell B7.
'//Created by "" on 4/10/2024. R-08.26/2024

    '// Assign variables
    Dim DateEntry As String
    
    '// Declare variables.
    DateEntry = Range("DateEntry")
    
    '// Conditional check to run macro.
    If DateEntry = "" Then
        MsgBox Prompt:="Please enter the date.", Buttons:=vbExclamation + vbOKCancel, Title:="Date Entry"
        Exit Sub
    End If
    
    '// Enable error handle.
    On Error GoTo ErrHandler:
    
    '// view coid.
    Sap_ViewCoid DateEntry
    
    '// Clean exit
    Exit Sub

ErrHandler:
    '// User update of failure.
    If Err.Number < "1" Then
        MsgBox Prompt:="Please open a session of SAP", Buttons:=vbCritical + vbOKCancel, Title:="Sap Session"
    Else
        MsgBox Prompt:="Failed to open SAP! Try again.", Buttons:=vbCritical + vbOKCancel, Title:="Sap Session"
    End If
    
End Sub

Sub Sp_OpenShiftReport()
'// This macro will open the fishreport by the date listed in B7 as read only. Sheets "NO" will be selected.
'// Created by "" on 5/12/2024. R-08/26/2024
  
    '// Declare variables.
    Dim DateEntry As String
    Dim FileDate As String
    Dim Wrkbook As String
    Dim SheetDate As String
    Dim Sheetpath As String
    
    '// Assign values.
    DateEntry = Range("DateEntry").Value
    FileDate = Format(DateEntry, "m-d-yy")
    Wrkbook = Gl_CcoWorkBook + (FileDate) + ".xlsx"
    SheetDate = Format(DateEntry, "yyyymmdd")
    Sheetpath = SheetDate + "Data"
      
    '// Conditional check to run macro.
    If DateEntry = "" Then
        MsgBox Prompt:="Please Insert The Date", Buttons:=vbExclamation + vbOKCancel, Title:="Date Entry"
        Exit Sub
    End If
          
    '// Enable error trapping.
    On Error GoTo ErrHandler:
    
    '// This will open the shift report in read only mode so it does not interfere with production.
        Workbooks.Open Filename:=Wrkbook, UpdateLinks:=3, ReadOnly:=True
        
    '// Maximize report and select first sheet.
Start:
        Application.WindowState = xlMaximized
        Worksheets("NO").Select
        
    '// Clean exit
    Exit Sub

ErrHandler:
    '// If file not found let user select or cancel.
    If Err = "1004" Then
        Dim FileToOpen As Variant
        Dim selectedbook As Workbook
        ChDrive "G:"
        ChDir "G:\Crackers\GF Shift reports"
        FileToOpen = Application.GetOpenFilename(filefilter:="Excel Files(*.xls*),*xls*", Title:="PLEASE SELECT THE CORRECT FILE")
    If FileToOpen <> False Then
        Set selectedbook = Application.Workbooks.Open(FileToOpen, UpdateLinks:=3, ReadOnly:=True)
        GoTo Start:
        ElseIf FileToOpen = False Then Exit Sub
    End If
    End If
    
    '// Error status update
    Err.Raise E_ErrorCustom.EecErrGeneral
    
End Sub

Sub Sp_RepositionHome()
'// This macro will scroll to range A1 "home" location from anywhere in the workbook.
'// Created by "" on 7/23/2024.
    Dim Home As Range
    Set Home = Range("A1")
        
    With Application
        .GoTo Home, True
    End With
    
End Sub

Sub Sp_RepositionCoid()
'// This macro will scroll to range DA100 "coid" location from anywhere in the workbook.
'// Created by "" on 7/23/2024.
    Dim Coid As Range
    Set Coid = Range("CoidImport")
       
    If Range("CoidImport").Offset(3, 1) = "" Then
        MsgBox "Please import Coid first!", vbCritical, "Position To Coid"
        Exit Sub
    Else
        With Application
            .GoTo Coid, True
        End With
    End If
End Sub

Sub Sp_RepositionPrisma()
'// This macro will scroll to range DA100 "coid" location from anywhere in the workbook.
'// Created by "" on 7/23/2024.
    Dim Prisma As Range
    Set Prisma = Range("PrismaImport")
       
    If Range("PrismaImport").Offset(1) = "" Then
        MsgBox "Please import Minimint Report first!", vbCritical, "Position To Minimint Report"
        Exit Sub
    Else
        With Application
            .GoTo Prisma, True
        End With
    End If
End Sub

Sub Sp_RepositionShiftReport()
'// This macro will scroll to range CA100 "ShiftReport" location from anywhere in the workbook.
'// Created by "" on 10/23/2024.
    Dim ShiftReport As Range
    Set ShiftReport = Range("NoShiftImport")
       
    If Range("NoShiftImport").Value = "" Then
        MsgBox "Please import the Shift Report first!", vbCritical, "Position To Shift Report"
        Exit Sub
    Else
        With Application
            .GoTo ShiftReport, True
        End With
    End If
End Sub

