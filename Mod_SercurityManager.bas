Attribute VB_Name = "Mod_SercurityManager"
'// All sercurity features and subs are to be listed here.
Option Private Module

'// Password for master worksheet protection.
Global Const ConSheetPassword As String =

'// Pasword for admin worksheet access.
Global Const ConAdminAccess As String =

Sub Sp_LogUser(ByVal ActiveUser As String, Optional ByVal EndUser As String)
'// Will log username of person accessing workbook to admin sheet. Acitvated by Workbook open event.
'// Created by "" on 10/11/2024.
    
    '// Force finish.
    On Error Resume Next
    
    '// Variable declarations.
    Dim User As Range
    Dim SameUser As Range
    
    '// Assign values.
    Set User = Worksheets("Admin").Range("J" & Rows.Count).End(xlUp).Offset(1, 0)
    Set SameUser = Worksheets("Admin").Range("K" & Rows.Count).End(xlUp).Offset(1, 0)
    
    '// Conditional check to either end user or not.
    If Trim(Len(EndUser)) > 1 Then
        SameUser.Value = ActiveUser & " " & Now
    Else
    '// Log user and display message box.
        User.Value = ActiveUser & " " & Now
        '// Save Log in.
        ThisWorkbook.Save
    End If
    
    '// Normal error behavior.
    On Error GoTo 0

End Sub

Sub Sp_ShowAdmin()
'// This sub will unhide admin sheet for maintanance.
'// Created by "" on 10/11/2024.
    
    '// Declare variable
    Dim UserInput As Variant
    
    '// Assign values
    UserInput = InputBox( _
        Prompt:="Please input the correct password to access the Admin Sheet.", _
        Title:="Password Required", _
        Default:="***Password***")
                        
    '// Check password and nomalize text.
    If LCase(Trim(UserInput)) = LCase(Trim(ConAdminAccess)) Then
            Worksheets("Admin").Visible = xlSheetVisible
            Worksheets("Log In").Visible = xlSheetVisible
            Worksheets("COID").Visible = xlSheetVisible
            Worksheets("Weekly Data").Visible = xlSheetVisible
            Worksheets("Master Worksheet").Unprotect ConSheetPassword
        Else
    '// User update of failure.
        MsgBox Prompt:="Password Incorrect." & vbNewLine & vbNewLine & "Access Denied.", _
            Buttons:=vbExclamation, _
            Title:="Input Password"
    End If
    
End Sub

Sub Sp_HideAdmin()
'// This sub will unhide admin sheet for maintanance.
'// Created by "" on 10/11/2024.
    
    '// Hide Admin sheet and other sheets,activated by admin deactivate worksheet event.
    Worksheets("Admin").Visible = xlVeryHidden
    Worksheets("Log In").Visible = xlVeryHidden
    Worksheets("COID").Visible = xlHidden
    Worksheets("Weekly Data").Visible = xlHidden
    Worksheets("Master Worksheet").Protect ConSheetPassword

End Sub

Sub Sp_ProtectMaster(ByVal UserPassword As String)
'// Sub that will protect master worksheet.
'// Created by "" on 10/11/2024.

    '// Password protect sheet.
    Dim ShMaster As Worksheet
    
    '// Assign variables.
    Set ShMaster = Worksheets("Master Worksheet")
    
    '// protect sheet with global password and force lower case.
    ShMaster.Protect PassWord:=LCase(Trim(UserPassword))

End Sub

Sub Sp_UnprotectMaster(ByVal UserPassword As String)
'// This sub will unprotect the master sheet.
'// Created by "" on 10/11/2024.

    '// Password protect sheet.
    Dim ShMaster As Worksheet
    
    '// Assign variables.
    Set ShMaster = Worksheets("Master Worksheet")
    
    '// Unprotech sheet with global password.
    ShMaster.Unprotect PassWord:=LCase(Trim(UserPassword))

End Sub

Sub Sp_ToggleProtection()
'// This sub will toggle sheet protection on or off and display status as in msgbox.
'// Created by "" on 7/23/2024. R-08/26/2024

    '// Protect and unprotect sheet with user msgbox update.
    If ThisWorkbook.ActiveSheet.ProtectContents = False Then
        ThisWorkbook.ActiveSheet.Protect
        MsgBox "Worksheet " & Application.ActiveSheet.Name & " has been Protected!", vbExclamation
    
    ElseIf ThisWorkbook.ActiveSheet.ProtectContents = True Then
        ThisWorkbook.ActiveSheet.Unprotect
        MsgBox "Worksheet " & Application.ActiveSheet.Name & " has been Unprotected!", vbExclamation
    End If
    
End Sub

Sub Sp_ViewAllWsHidden(ByVal SheetCheck As Boolean)
'// This sub will call veryhidden function for worksheets,
'// Created by "" on 10/11/2024.

    '// Call function for verification.
    Fn_SheetsVeryHidden (SheetCheck)
    
End Sub

Function Fn_VerifyLogin(ByVal UserName As String, ByVal PassWord As String) As Boolean
'// This function will verify the user has the correct login information.
'// Created by "" on 10/11/2024.

    '// Enable error label.
    On Error GoTo ErrHandler:

    '// Declare variables.
    Dim WsAdmin As Worksheet
    Dim RngLoginInfo As Range
    Dim MsgResult As String
    
    '// Assign values.
    Set WsAdmin = Worksheets("Admin")
    Set RngLoginInfo = WsAdmin.Range("E1").CurrentRegion
   
    '// Check for user and password match.
    If LCase(Trim(WorksheetFunction.VLookup(UserName, RngLoginInfo, 2, False))) = LCase(Trim(PassWord)) Then
        '// Login success
        MsgResult = MsgBox(Prompt:="Success, you are now logged in.", Buttons:=vbInformation, Title:="Success")
        '// Assign function value.
        Fn_VerifyLogin = True
    Else
        '// Login failure
        MsgResult = MsgBox(Prompt:="Login Failed, Username or Password was incorrect.", Buttons:=vbExclamation, Title:="Failed")
        '// Assign function value
        Fn_VerifyLogin = False
    End If
    
    '// Clean exit
    Exit Function
    
ErrHandler:
    '// Error handle for false usernames or passwords.
        MsgResult = MsgBox(Prompt:="Login Failed, Username or Password was incorrect.", Buttons:=vbExclamation, Title:="Failed")
End Function

Function Fn_SheetsVeryHidden(WsStatus As String) As Boolean
'// This sub will place the visibilty of all sheets in the project to xlvery hidden or xlvisible. Activate by Workbook events.
'// Created by "" on 10/11/2024.

    '// Declare variables.
    Dim WsToHide As Worksheet
    
    '// Start case based on boolean bvalue recieved.
     Select Case WsStatus
     
        Case True
            '// Show log in sheet.
            Worksheets("Log In").Visible = xlSheetVisible
            
            '// Start loop through collection.
            For Each WsToHide In Worksheets
                '// Conditional requirement.
                If WsToHide.Name <> Worksheets("Log In").Name Then
                    WsToHide.Visible = xlSheetVeryHidden
                 End If
            Next WsToHide
                
        Case False
            '// Start loop through collection.
            For Each WsToHide In Worksheets
            '// Conditional requirement.
                If WsToHide.Name <> Worksheets("Log In").Name Then
                    WsToHide.Visible = xlSheetVisible
                End If
            Next WsToHide
            '// Rehide raw data sheets.
            Worksheets("Admin").Visible = xlSheetVeryHidden
            Worksheets("Log In").Visible = xlSheetVeryHidden
            Worksheets("COID").Visible = xlHidden
            Worksheets("Weekly Data").Visible = xlHidden
    
    End Select
    
    Fn_SheetsVeryHidden = WsStatus
    
End Function
