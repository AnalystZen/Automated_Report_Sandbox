Attribute VB_Name = "Mod_SapLoops"
'// List all Sap/Excel loops for the project here.
Option Private Module

Sub Sp_AutoAdjustDelivery()
'// This sub will adjust the target amount for orders in sap to 1, only orders with no amount delivered will be adjusted.
'// Created by "" on 7/23/2024. R-08/25/2024

    '// Declare variables.
    Dim DateEntry As String
    Dim FileDate As String
    Dim TrackerWb As String
    Dim CellRow As String
    Dim BlockEnd As String
    
    '// Assign values
    DateEntry = Range("B7").Value
    FileDate = Format(DateEntry, "yyyymmdd")
    TrackerWb = ThisWorkbook.Name
    CellRow = 9
    BlockEnd = 0
    
    '// Verifying multiple conditions are met before running macro. Date and delivered amount query.
    If DateEntry = "" Then
            MsgBox "Please enter the date in Range B7!", vbCritical, "Date Entry"
            Exit Sub
        ElseIf MsgBox("This macro will adjust target quantities of the POs on this sheet,Continue?", _
                vbExclamation + vbYesNo, "Adjust PO target quantities") = vbNo Then
                Exit Sub
        ElseIf Workbooks(TrackerWb).Sheets(FileDate).Range("AR" & CellRow) < 1 Or _
                Workbooks(TrackerWb).Sheets(FileDate).Range("DB103") = "" Then
                MsgBox ("Please import the delivered amount from SAP!")
        Exit Sub
    End If
    
    '// Enable error handle.
    On Error GoTo ErrHandler:
    
    '// Establishes SAP connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
    
'// Start of Loop. Loop will end when macro reaches 2 consecutive blank cells.
Loopstart:
On Error GoTo ErrHandler:
    Do While BlockEnd < 2
        If Workbooks(TrackerWb).Sheets(FileDate).Range("A" & CellRow) = "" Then
                BlockEnd = BlockEnd + 1
                CellRow = CellRow + 1
                Else
            If Workbooks(TrackerWb).Sheets(FileDate).Range("AT" & CellRow) > 0 Or _
                Workbooks(TrackerWb).Sheets(FileDate).Range("AS" & CellRow) = 1 Then
                CellRow = CellRow + 1
                Else
                
    '// Start COR2
    session.StartTransaction "COR2"
    '// Selects the appropriate PO from thisworkbook.filedate
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = Workbooks(TrackerWb).Sheets(FileDate).Range("A" & CellRow)
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    '// Set the schedule parameters to not shift PO'//s to a different date.
    session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[1]").Select
    session.findById("wnd[1]/usr/chkTCX00-AUF_SHIFT").Selected = True
    session.findById("wnd[1]/usr/txtTCX00-STVERG").Text = "999"
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    '// Set the target quantity equal to the appropriate Expected Eaches cell
    session.findById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOZE/ssubSUBSCR_5115:SAPLCOKO:5120/txtCAUFVD-GAMNG").Text = "1"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    '// Save the new target amount transaction
    session.findById("wnd[0]/tbar[0]/btn[11]").press
 
    '// Next cell is selected.
    CellRow = CellRow + 1
    BlockEnd = 0
            End If
        End If
    Loop
    
    '// Re-imports COID for PO target amounts.
    With Workbooks(TrackerWb).Sheets(FileDate).Range("A9:A70")
        Application.Intersect(.SpecialCells(xlCellTypeVisible), _
                           .SpecialCells(xlCellTypeConstants)).Copy
    End With
    '// Starts Coid Session for PO'//s.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "COID"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtP_PROFID").Text = "000001"
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = "/AL COID"
    session.findById("wnd[0]/usr/ctxtS_OWERK-LOW").Text = "4014"
    session.findById("wnd[0]/usr/ctxtS_OWERK-LOW").SetFocus
    session.findById("wnd[0]/usr/ctxtS_OWERK-LOW").caretPosition = 4
    session.findById("wnd[0]/usr/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]").sendVKey 8
    '// Sap export to excel.
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    
    '// Clears old data from columns and pastes COID
    Sheets(FileDate).Select
    Columns("DA:DP").Select
    Selection.Clear
    Range("DA100").Select
    ActiveSheet.Paste
    
    '// Format coid with | delimiter. This installs Coid into columns.
    Columns("DA").Select
    Selection.TextToColumns Destination:=Range("DA1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
 
    '// Selects Range A1.
     ThisWorkbook.Sheets(FileDate).Range("A1").Select
    
    '// Clean exit from sub before hitting error handler.
     Exit Sub
 
ErrHandler:
    '// User notification of failure.
    If Err.Number = 614 Then
        MsgBox ("Please open a session of SAP")
    End If
    '// Loop restart after error.
         If Err.Number = 619 Then
            CellRow = CellRow + 1
            session.findById("wnd[0]").resizeWorkingPane 94, 28, False
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            session.findById("wnd[0]/tbar[0]/btn[3]").press
            On Error GoTo -1
            GoTo Loopstart:
         End If
End Sub

Sub Sp_AutoAdjustDiff()
'// This macro will auto adjust the confirmed and delivered difference without confirming the process orders. It will only adjust positve values.
'// Created by "" on 6/20/2024. R-08/25/2024

    '// Declare variables.
    Dim DateEntry As String
    Dim FileDate As String
    Dim TrackerWb As String
    Dim CellRow As String
    Dim BlockEnd As String
    Dim AdjustDiff As String
    
    '// Assign values.
    DateEntry = Range("B7").Value
    FileDate = Format(DateEntry, "yyyymmdd")
    TrackerWb = ThisWorkbook.Name
    CellRow = 9
    BlockEnd = 0

    '// Conditional verifications to run macro.
    If FileDate = "" Then
        MsgBox Prompt:="Please enter the date.", Buttons:=vbExclamation + vbOKCancel, Title:="Date Entry"
        Exit Sub
    End If
    '// Conditional verifications to run macro.
    If MsgBox("This macro will adjust the confirmed amount difference for the selected PO'//s with adj listed in the teco column." & vbNewLine & _
    "Do you want to continue?", vbExclamation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    '// Enable error handle
    On Error GoTo ErrHandler:

    '// Establish SAP Connection
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)

    If IsObject(WScript) Then
     WScript.ConnectObject session, "on"
     WScript.ConnectObject Application, "on"
    End If
 
    '// Multiconditional checks before loop start.
Loopstart:
On Error GoTo ErrHandler:
    Do While BlockEnd < 2

    '// If the macro reaches two adjecent blank cells, it will stop
    If Workbooks(TrackerWb).Sheets(FileDate).Range("A" & CellRow) = "" Then
        BlockEnd = BlockEnd + 1
        CellRow = CellRow + 1
    Else
    
    If Workbooks(TrackerWb).Sheets(FileDate).Range("AP" & CellRow) <> "adj" Then
        CellRow = CellRow + 1
    Else
    
    '// Add to confirmed if a positive value, skip if negative, and skip if equal
    If Workbooks(TrackerWb).Sheets(FileDate).Range("AO" & CellRow) < 0 Then
        CellRow = CellRow + 1
    Else
    If Workbooks(TrackerWb).Sheets(FileDate).Range("AO" & CellRow) = 0 Then
        AdjustDiff = ""
    End If
    
    If Workbooks(TrackerWb).Sheets(FileDate).Range("AO" & CellRow) > 0 Then
        AdjustDiff = Workbooks(TrackerWb).Sheets(FileDate).Range("AO" & CellRow).Value
    End If

    '// Starts Cor6 SAP transaction.
    session.StartTransaction "COR6"
    '// Lists current cellrow according to Loop cell Position.
    session.findById("wnd[0]/usr/ctxtCORUF-AUFNR").Text = Workbooks(TrackerWb).Sheets(FileDate).Range("A" & CellRow)
    '// Adjusts the calculated difference between confirmed and delivered in excel sheet.
    session.findById("wnd[0]/usr/ctxtCORUF-AUFNR").SetFocus
    session.findById("wnd[0]/usr/ctxtCORUF-AUFNR").caretPosition = 9
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtAFRUD-LMNGA").Text = AdjustDiff
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/radCORUF-TEILR").SetFocus
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    
    '// Writes "done" in cell and starts next cell loop.
    Workbooks(TrackerWb).Sheets(FileDate).Range("AP" & CellRow) = "Done"
    CellRow = CellRow + 1
    BlockEnd = 0

    End If
    End If
    End If
    Loop
    
    '// Start SAP COID transaction and copy
       Sheets(FileDate).Select
    With Sheets(FileDate).Range("A9:A70")
        Application.Intersect(.SpecialCells(xlCellTypeVisible), _
                           .SpecialCells(xlCellTypeConstants)).Copy
    End With
    '// Sap t code selection.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "COID"
    
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtP_PROFID").Text = "000001"
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = "/AL COID"
    session.findById("wnd[0]/usr/ctxtS_OWERK-LOW").Text = "4014"
    session.findById("wnd[0]/usr/ctxtS_OWERK-LOW").SetFocus
    session.findById("wnd[0]/usr/ctxtS_OWERK-LOW").caretPosition = 4
    session.findById("wnd[0]/usr/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]").sendVKey 8
    '// Sap export to excel.
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    '// Paste COID
    Sheets(FileDate).Select
    Columns("DA:DN").Select
    Selection.Clear
    Range("DA100").Select
    ActiveSheet.Paste
    
    '// Format coid with | delimiter. This installs Coid into columns.
    Columns("DA").Select
    Selection.TextToColumns Destination:=Range("DA1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
 
    '// Activates initial sheet macro was started on.
    Sheets(FileDate).Select
    Range("A1").Select
 
    '// Clean Exit
    Exit Sub
    
ErrHandler:
    '// User notification.
     If Err.Number = 614 Then
        MsgBox Prompt:="Please open a session of SAP", Buttons:=vbCritical + vbOKCancel, Title:="SAP Session"
    End If
    '// Loop restart after error.
    If Err.Number = 619 Then
        Workbooks(TrackerWb).Sheets(FileDate).Range("AP" & CellRow) = "?"
        CellRow = CellRow + 1
        session.findById("wnd[0]").resizeWorkingPane 133, 41, False
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        On Error GoTo -1
        GoTo Loopstart:
    End If
    
End Sub

Sub Sp_AutoConfirmPOs()
'// This macro auto-confirms the process orders starting from cellrow A9 until three empty spaces.Any PO'//s with a "cnf"
'// typed into the "TECO" column will be adjusted in SAP.
'// Created by NB; Updated by "" on 4/15/2024. R-08/25/2024

    '// Assign values.
    Dim DateEntry As String
    Dim FileDate As String
    Dim TrackerWb As String
    Dim CellRow As String
    Dim BlockEnd As String
    Dim AdjustDiff As String
    
    '// Assign values.
    DateEntry = Range("B7").Value
    FileDate = Format(DateEntry, "yyyymmdd")
    TrackerWb = ThisWorkbook.Name
    CellRow = 9
    BlockEnd = 0
    
    '// Conditional check to run macro.
    If FileDate = "" Then
        MsgBox Prompt:="Please enter the date.", Buttons:=vbInformation + vbOKOnly, Title:="Date Entry"
        Exit Sub
    End If
    
    '// Conditional check to run macro.
    If MsgBox("Do you want to CONFIRM the marked PO'//s on this sheet?", vbExclamation + vbYesNo, "Adjust PO target quantities") = vbNo Then
        Exit Sub
    End If
    
    '// Enable error handle.
    On Error GoTo ErrHandler:
    
    '// Set SAP connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
    
    '// Loop start after conditional checks.
Loopstart:
On Error GoTo ErrHandler:
    Do While BlockEnd < 3
    
    '// If the macro reaches two adjecent blank cells, it will stop
    If Workbooks(TrackerWb).Sheets(FileDate).Range("A" & CellRow) = "" Then
        BlockEnd = BlockEnd + 1
        CellRow = CellRow + 1
    Else
    
    If Workbooks(TrackerWb).Sheets(FileDate).Range("AP" & CellRow) <> "cnf" Then
        CellRow = CellRow + 1
    Else
    
    '// Add to confirmed if a positive value, skip if negative, and skip if equal
    If Workbooks(TrackerWb).Sheets(FileDate).Range("AO" & CellRow) < 0 Then
        CellRow = CellRow + 1
    Else
    If Workbooks(TrackerWb).Sheets(FileDate).Range("AO" & CellRow) = 0 Then
        AdjustDiff = ""
    End If
    If Workbooks(TrackerWb).Sheets(FileDate).Range("AO" & CellRow) > 0 Then
        AdjustDiff = Workbooks(TrackerWb).Sheets(FileDate).Range("AO" & CellRow).Value
    End If
    
    '// Make sure that the final confirmation box is checked
    session.StartTransaction "COR2"
    session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = Workbooks(TrackerWb).Sheets(FileDate).Range("A" & CellRow)
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE").Select
    session.findById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE/ssubSUBSCR_5115:SAPLCOKO:5190/chkAFPOD-ELIKZ").Selected = True
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    
    '// Confirm the PO
    session.StartTransaction "COR6"
    session.findById("wnd[0]/usr/ctxtCORUF-AUFNR").Text = Workbooks(TrackerWb).Sheets(FileDate).Range("A" & CellRow)
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/radCORUF-ENDRU").Select
    session.findById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/chkCORUF-AUSBU").Selected = True
    session.findById("wnd[0]/usr/tabsTABSTRIP_0150/tabpMGLE/ssubVAR_CNF_10:SAPLCORU:5850/txtAFRUD-LMNGA").Text = AdjustDiff
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    
    Workbooks(TrackerWb).Sheets(FileDate).Range("AP" & CellRow) = "Y"
    
    CellRow = CellRow + 1
    BlockEnd = 0
    
    End If
    End If
    End If
    Loop
    
    '// Copies po'//s for sap export.
    Sheets(FileDate).Select
    With Sheets(FileDate).Range("A9:A70")
        Application.Intersect(.SpecialCells(xlCellTypeVisible), _
                           .SpecialCells(xlCellTypeConstants)).Copy
    End With
      
    '// Establishes SAP re connection.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "COID"
    
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtP_PROFID").Text = "000001"
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = "/AL COID"
    session.findById("wnd[0]/usr/ctxtS_OWERK-LOW").Text = "4014"
    session.findById("wnd[0]/usr/ctxtS_OWERK-LOW").SetFocus
    session.findById("wnd[0]/usr/ctxtS_OWERK-LOW").caretPosition = 4
    session.findById("wnd[0]/usr/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]").sendVKey 8
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

    '// Paste COID into active sheet.
    Sheets(FileDate).Select
    Columns("DA:DN").Select
    Selection.Clear
    Range("DA100").Select
    ActiveSheet.Paste
    
    '// Format coid with | delimiter. This installs Coid into columns.
    Columns("DA").Select
    Selection.TextToColumns Destination:=Range("DA1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
        
    '// Selects intital sheet macro was run on.
    Sheets(FileDate).Select
    Range("A1").Select

    '// Clean exit
    Exit Sub

ErrHandler:
    '// User notification of failure.
    If Err.Number = 614 Then
        MsgBox Prompt:="Please open a session of SAP", Buttons:=vbCritical + vbOKCancel, Title:="Sap Session."
    End If
    '// Loop restart after error.
    If Err.Number = 619 Then
        Workbooks(TrackerWb).Sheets(FileDate).Range("AP" & CellRow) = "?"
        CellRow = CellRow + 1
        session.findById("wnd[0]").resizeWorkingPane 133, 41, False
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press
        On Error GoTo -1
        GoTo Loopstart:
    End If

End Sub




