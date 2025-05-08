Attribute VB_Name = "Mod_SapGuiScripts"
'// List all SAP Gui static scripts here.
Option Private Module

Sub Sap_CoidImport(ByVal DateEntry As String)
'// This script will open coid for process order information.
'// Created by "" on 10/07/2024.

    '// Establishes SAP connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
    '// Sap T code selection.
    session.findById("wnd[0]").resizeWorkingPane 94, 27, False
    session.StartTransaction "COID"
     
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 17
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "LASSALAN"
    session.findById("wnd[1]").sendVKey 8
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "3"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").Text = DateEntry
    session.findById("wnd[0]").sendVKey 8
    '// Export Sap data to clipboard.
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

End Sub

Sub Sap_CaseImport(Optional ByVal DateEntry As String)
'// This script will import sap coid for the cases delivered for process orders.
'// Created by "" on 10/17/2024.

'// Establishes SAP connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
    '// Sap T code selection.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "COID"
    '// Sap variant selection
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

End Sub

Sub Sap_MixImport(Optional ByVal DateEntry As String)
'// This script will import sap mixes for process orders.
'// Created by "" on 10/17/2024.

    '// This establishes the SAP connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
    '// Sap t code selection.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "COID"
    '// Sap variant selection.
    session.findById("wnd[0]/usr/radREP_OPER").Select
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/ctxtP_PROFID").Text = "000001"
    session.findById("wnd[0]/usr/ctxtP_LAYOUT").Text = "/ALMIXCOMMIT"
    session.findById("wnd[0]/usr/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    '// Export data from Sap to Excel.
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
    session.findById("wnd[0]/usr/cntlGRID_0100/shellcont/shell").selectContextMenuItem "&PC"
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/tbar[0]/btn[0]").press

End Sub

Sub Sap_ProdReport(ByVal DateEntry As String, ByVal FileDate As String, ByVal TimeStart As String, _
    ByVal TimeEnd As String, ByVal VarRange As String)
'// This script will import the prod report for process orders by shift. It will loop for all 3 shifts.
'// Created by "" on 10/17/2024.

    '// This establishes the SAP connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
       
    '// Sap t code selection.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "ZWMPRODPAL"
    
GetData:
    session.findById("wnd[0]/usr/ctxtP_LGNUM").Text = "407"
    session.findById("wnd[0]/usr/ctxtS_GSTRS-LOW").Text = DateEntry
    '// Erase fields that may contain text.
    session.findById("wnd[0]/usr/ctxtS_AUFNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtS_CHARG-LOW").Text = ""
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    
    ' Filter layout is "TIMEORDERFILTER"
    session.findById("wnd[0]").sendVKey 33
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").contextMenu
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectContextMenuItem "&FIND"
    session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").Text = "/TIMEFILTER"
    session.findById("wnd[2]").sendVKey 0
    session.findById("wnd[2]").Close
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell

    ' Start of filter for time slots.
    session.findById("wnd[0]/tbar[1]/btn[29]").press
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").currentCellRow = 22
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").doubleClickCurrentCell
    session.findById("wnd[1]/usr/subSUB_DYN0500:SAPLSKBH:0600/btn600_BUTTON").press
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = TimeStart
    session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-HIGH").Text = TimeEnd
    session.findById("wnd[2]/tbar[0]/btn[0]").press
    '// Export of data start.
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    '// Back to Main screen.
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    '// This pastes new data from sap.
    Sheets(FileDate).Select
    Range(VarRange).Select
   
    ActiveSheet.Paste
        
    '// Format coid with | delimiter. This installs Coid into columns.
    Range(VarRange).EntireColumn.Select
    Selection.TextToColumns Destination:=Range(VarRange).EntireColumn.Resize(1), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
    1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
    , 1), Array(13, 1), Array(14, 1), Array(15, 1)), TrailingMinusNumbers:=True
 
    '// Get AM and PM data by swapping variable assignment.
    If TimeStart = "00:00" Then
            TimeStart = "07:30"
            TimeEnd = "15:30"
            VarRange = "AmCaseImport"
            GoTo GetData
        ElseIf TimeStart = "07:30" Then
            TimeStart = "15:30"
            TimeEnd = "23:30"
            VarRange = "PmCaseImport"
            GoTo GetData
    End If
            
     '// Go to sap main screen.
     session.StartTransaction "/N"
     
End Sub

Sub Sap_ViewCoid(ByVal DateEntry As String)
'// View Daily coid.
'// Created by "" on 10/8/2024.

    '// Establishes SAP connection.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set Connection = SAPApp.Children(0)
    Set session = Connection.Children(0)
    
    If IsObject(WScript) Then
       WScript.ConnectObject session, "on"
       WScript.ConnectObject Application, "on"
    End If
    
    '// Sap T code selection.
    session.findById("wnd[0]").resizeWorkingPane 94, 28, False
    session.StartTransaction "COID"
    
    '// Sap variant selection.
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = "LASSALAN"
    session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
    session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 7
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 0
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").Text = DateEntry
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").SetFocus
    session.findById("wnd[0]/usr/ctxtS_ECKST-LOW").caretPosition = 9
    session.findById("wnd[0]").sendVKey 8

End Sub
