Attribute VB_Name = "Mod_TestSubs"
'// Test any new procedures or functions here.
Option Private Module

Sub Test_DateCheck()
'//funtion test

    DateEntry = Fn_FormatDate(Range("B3").Value)
    Debug.Print DateEntry
    If DateEntry = False Then Exit Sub
        
End Sub

Sub Test_Rng()
'//Range("ProcessOrders").Resize(EkrClearRange).Select
        
        With ActiveSheet
        '//Sheets(FileDate).Select
        '//Range("CoidImport").EntireColumn.Resize(, 16).Clear
        '//Selection.Clear
        '//Range("DA100").Select
        '//ActiveSheet.Paste
        '//Range("CoidImport").EntireColumn.Resize(1).Select
        Range("ProcessOrders").Resize(EkrClearRange, 29).Select
        
    End With
End Sub

Sub Test_Formulas()
'// This sub replaced indirect formulas on the worksheet to increse performance. It sums mixes moved between production dates.
'// Created by "" on 10/9/2024.
    
    '// Declare variables.
    Dim DateTest As Date
    Dim FileDate As String
    Dim MyRange As Range
    Dim MyCell As Range
    Dim X As Long
    
    '// Assign values
    DateTest = Range("B7") - 1
    FileDate = Format(DateTest, "yyyymmdd")
    Set MyRange = Range("L9:L75")
    X = 9
    
    For Each MyCell In MyRange
        '// Conditional check to insert formula values.
        If LCase(InStr(Range("C" & X).Value, "FISHWIP")) > 1 Then
        
            '// Offshift mixes
            Range("L" & X) = WorksheetFunction.SumIf(Worksheets(FileDate).Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & X), Worksheets(FileDate).Range("$CE$100:$CE$165"))
            
            '// Mixes moved out of orders.
            Range("N" & X) = WorksheetFunction.SumIf(Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & X), Range("$CI$100:$CI$165")) _
            + WorksheetFunction.SumIf(Worksheets(FileDate).Range("$CB$100:$CB$165"), ActiveSheet.Range("A" & X), (Worksheets(FileDate).Range("$CI$100:$CI$165")))
            
            '// Mixes moved into orders.
            Range("O" & X) = WorksheetFunction.SumIf(Range("$CH$100:$CH$165"), ActiveSheet.Range("A" & X), Range("$CI$100:$CI$165")) _
            + WorksheetFunction.SumIf(Worksheets(FileDate).Range("$CH$100:$CH$165"), ActiveSheet.Range("A" & X), (Worksheets(FileDate).Range("$CI$100:$CI$165")))
        End If

        X = X + 1
    
    Next MyCell
    
End Sub

Sub Test_FormulasVersion2()
'// This sub replaced indidrect formulas on the worksheet to increse performance. It sums mixes moved between production dates via a loop.
'// Created by "" on 10/9/2024.
    
    '// Trun off normal error behavior.
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
        Debug.Print MyCell.Row
        
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
    
    '// Trun on normal error behavior.
    On Error GoTo 0
    
End Sub

Sub Functiontest()
'// Test for emppty range.
    
    DateEntry = Fn_RngIsNothing(Range("A6:A6"))
    Debug.Print DateEntry
    
End Sub

Sub TestDragFormulas()
'// Sub to drag formulas down range

    '// select cell
    Range("AB9").AutoFill Range("AB9:AB75"), xlFillDefault
    
End Sub



