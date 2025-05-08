Attribute VB_Name = "Mod_Functions"
'// List all created functions or formula value related procedures for this project here.
Option Private Module

Function Fn_FormatDate(ByVal DateGiven As String) As String
'// Funtion to format input date for user and check if its empty.
    
    '// Format start
    DateGiven = Trim(Format(DateGiven, "YYYYMMDD"))

    '// Check value
    Select Case Len(Trim(DateGiven))
        Case Is < 1 '// Empty String return
            DateGiven = False
            MsgBox Prompt:="Please insert the date into the active sheet.", _
            Buttons:=vbExclamation + vbOKCancel, _
            Title:="Date Entry"
        
        Case Else
            DateGiven = DateGiven '// Check is good.
    End Select
    
    '// Return value
    Fn_FormatDate = DateGiven
    
End Function

Function Fn_LastRow(RngRow As Long, RngColumn As Long) As String
'// Function to find the last row of data for user selected range
    
    '// Declare variables
    Dim WsActive As Worksheet
    Set WsActive = ThisWorkbook.ActiveSheet
    
    '// Find last row of given range.
    With WsActive.Range(Cells(RngRow, 1), Cells(EkrDataRow, RngColumn)).SpecialCells(xlCellTypeConstants)
        '// Return value to function.
        Fn_LastRow = .Cells(.Cells.Count).Row
    End With

End Function

Sub Sp_CalculateShiftMixes()
'// This sub replaced indidrect formulas on the worksheet to increse performance. It sums mixes moved between production dates via a loop.
'// Created by "" on 10/9/2024.
    
    '// Turn off normal error behavior.
    On Error Resume Next
    
    '// Declare variables.
    Dim DateTest As Date
    Dim FileDate As String
    Dim MyRange As Range
    Dim MyCell As Range
    
    '// Assign values
    DateTest = Range("B7") - 1
    FileDate = Format(DateTest, "yyyymmdd")
    Set MyRange = Range("C9:C75")
    
    '// Start of loop.
    For Each MyCell In MyRange
        
        '// Conditional check to insert formula values.
        If LCase(InStr(MyCell.Value, "FISHWIP")) > 1 Then
        
            '// Offshift mixes
            Range("L" & MyCell.Row) = WorksheetFunction.SumIf(Worksheets(FileDate).Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & MyCell.Row), Worksheets(FileDate).Range("$CE$100:$CE$165"))
            
            '// Mixes moved out of orders.
            Range("N" & MyCell.Row) = WorksheetFunction.SumIf(Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & MyCell.Row), Range("$CI$100:$CI$165")) _
            + WorksheetFunction.SumIf(Worksheets(FileDate).Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & MyCell.Row), (Worksheets(FileDate).Range("$CI$100:$CI$165")))
            
            '// Mixes moved into orders.
            Range("O" & MyCell.Row) = WorksheetFunction.SumIf(Range("$CH$100:$CH$165"), ActiveSheet.Range("A" & MyCell.Row), Range("$CI$100:$CI$165")) _
            + WorksheetFunction.SumIf(Worksheets(FileDate).Range("$CH$100:$CH$165"), ActiveSheet.Range("A" & MyCell.Row), (Worksheets(FileDate).Range("$CI$100:$CI$165")))
        End If
   
    Next MyCell
    
    '// Turn on normal error behavior.
    On Error GoTo 0
    
    '// Retrieve mixes moved data from worksheets.
    Sp_CalculateMixesMoved
    
End Sub

Function Fn_RngIsNothing(ByVal RngCheck As Range) As Boolean
'// This function is designed to check the value of a range/cell and return false if there is no value.
'// Created by Antonio Lasslle on 10/11/2024.

    '// Conditional check and user notification.
    If IsEmpty(RngCheck) Then
        MsgBox _
        Prompt:=" Please verify the contents of " & RngCheck.Address & " and try again." _
        , Buttons:=vbOKOnly + vbExclamation _
        , Title:="Verify Value"
            
        '// Assign value to function if false.
        Fn_RngIsNothing = False
        
    Else
        '// Assign value to function if true.
         Fn_RngIsNothing = True
    End If

End Function

Function LinesOfCode()
'// Find how many lines of code are in the project.
'// Created by "" on 10/13/2024.

    '// Declare variables.
    Dim VbeModule As Object
    
    '// 0 Lines of code value for count start.
    LinesOfCode = 0
    
    '// Start loop.
    For Each VbeModule In Application.VBE.ActiveVBProject.VBComponents
        LinesOfCode = LinesOfCode + VbeModule.CodeModule.CountofLines
    Next VbeModule
    
    '// Print line total to immediate window. Includes comments.
    Debug.Print LinesOfCode

End Function

Sub Sp_CalculateMixesMoved(Optional ByVal DateEntry As String)
'// This sub will calculate mixes moved in the ic worksheet. It sums mixes moved between production dates via a loop.
'// Created by "" on 10/27/2024.
    
    '// Turn off normal error behavior.
    On Error Resume Next
    
    '// Declare variables.
    Dim PreviousDate As Date
    Dim NextDate As Date
    Dim FileDate As String
    Dim MyRange As Range
    Dim MyCell As Range
    
    '// Assign values
    PreviousDate = Range("B7") - 1
    FileDate = Format(PreviousDate, "yyyymmdd")
    NextDate = Range("B7") + 1
    NextFileDate = Format(NextDate, "yyyymmdd")
    Set MyRange = Range("C9:C75")
    
    '// Calculations need to be on for this part.
    Application.Calculation = xlCalculationAutomatic
    
    '// Start of loop for previous date of mixes moved.
    For Each MyCell In MyRange
        
        '// Conditional check to insert formula values.
        If LCase(InStr(MyCell.Value, "FISHWIP")) > 1 Then
            
            If Range("AZ" & MyCell.Row) > "" And Range("BC" & MyCell.Row) = "" Then
            
                '// What order the mixes were moved out from from previous and next date.
                Range("BC" & MyCell.Row) = WorksheetFunction.SumIf(Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & MyCell.Row), Range("$CH$100:$CH$165")) _
                + WorksheetFunction.SumIf(Worksheets(FileDate).Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & MyCell.Row), (Worksheets(FileDate).Range("$CH$100:$CH$165"))) _
                + WorksheetFunction.SumIf(Worksheets(NextFileDate).Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & MyCell.Row), (Worksheets(NextFileDate).Range("$CH$100:$CH$165")))
                
                '// Adjust sum of orders if mulitple found.
                If Range("BC" & MyCell.Row).Value > 300000000 Then Range("BC" & MyCell.Row).Value = Range("BC" & MyCell.Row).Value / 2
                
                '// Mixes moved into orders for previous and next date.
                Range("BB" & MyCell.Row) = WorksheetFunction.SumIf(Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & MyCell.Row), Range("$CI$100:$CI$165")) _
                + WorksheetFunction.SumIf(Worksheets(FileDate).Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & MyCell.Row), (Worksheets(FileDate).Range("$CI$100:$CI$165"))) _
                + WorksheetFunction.SumIf(Worksheets(NextFileDate).Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & MyCell.Row), (Worksheets(NextFileDate).Range("$CI$100:$CI$165")))
             
             End If
        
             If Range("BB" & MyCell.Row) > "" And Range("BA" & MyCell.Row) = "" Then

                '// What order the mixes were moved out from from previous and next date.
                Range("BA" & MyCell.Row) = WorksheetFunction.SumIf(Range("$CH$100:$CH$165"), ActiveSheet.Range("A" & MyCell.Row), Range("$CB$100:$CB$165")) _
                + WorksheetFunction.SumIf(Worksheets(FileDate).Range("$CH$100:$CH$165"), ActiveSheet.Range("A" & MyCell.Row), (Worksheets(FileDate).Range("$CB$100:$CB$165"))) _
                + WorksheetFunction.SumIf(Worksheets(NextFileDate).Range("$CH$100:$CH$165"), ActiveSheet.Range("A" & MyCell.Row), (Worksheets(NextFileDate).Range("$CB$100:$CB$165")))

                '// Adjust sum of orders if mulitple found.
                If Range("BA" & MyCell.Row).Value > 300000000 Then Range("BA" & MyCell.Row).Value = Range("BA" & MyCell.Row).Value / 2

                '// Mixes moved into orders for previous and next date.
                Range("AZ" & MyCell.Row) = WorksheetFunction.SumIf(Range("$CH$100:$CH$165"), ActiveSheet.Range("A" & MyCell.Row), Range("$CI$100:$CI$165")) _
                + WorksheetFunction.SumIf(Worksheets(FileDate).Range("$CH$100:$CH$165"), ActiveSheet.Range("A" & MyCell.Row), (Worksheets(FileDate).Range("$CI$100:$CI$165"))) _
                + WorksheetFunction.SumIf(Worksheets(NextFileDate).Range("$CH$100:$CH$165"), ActiveSheet.Range("A" & MyCell.Row), (Worksheets(NextFileDate).Range("$CI$100:$CI$165")))

             End If
        
        End If
    
    Next MyCell
    
    '// Turn on normal error behavior.
    On Error GoTo 0
    
End Sub
