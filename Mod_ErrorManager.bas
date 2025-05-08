Attribute VB_Name = "Mod_ErrorManager"
'// List all procedures related to error handling for the project here.
Option Private Module

Sub ErrorManager(ByVal ErrNumber As E_ErrorCustom, Optional ByVal ErrMessage As String)
'// Handle Error Messages and custom raised errors. Activated with action manager.
    
    '// Declare Variables.
    Dim ErrTitle As String
    Dim ErrMsg As String
    
    '// Handle Errors
    On Error GoTo ErrHandler
        
    '// Value passed to select case from sub call.
    Select Case ErrNumber
        
        Case E_ErrorCustom.EecErrGeneral
        '// General error message to user.
            ErrTitle = "Action Falied"
            ErrMsg = "Something went wrong!." & vbNewLine & vbNewLine & _
                     "Please verify actions and try again."
        
        Case E_ErrorCustom.EecErrExport
        '// Error for exporting and report creation.
            ErrTitle = "Export Failed"
            ErrMsg = "Report creation and export failed!" & vbNewLine & vbNewLine & _
                    "Please verify actions and try again."

        Case E_ErrorCustom.EecErrImport
        '// Error for imports.
            ErrTitle = "Import Failed"
            ErrMsg = "Import of data failed." & vbNewLine & vbNewLine & _
                    "Please verify data types and try again."
            
        Case E_ErrorCustom.EecErrSapGui
        '// Error for SAP.
            ErrTitle = "Sap Failed"
            ErrMsg = "Something went wrong!." & vbNewLine & vbNewLine & _
                    "Please verify a session of SAP is open and try again."
        
        Case E_ErrorCustom.EecErrSercurity
            ErrTitle = "Incorrect Information"
            ErrMsg = "Something went wrong!." & vbNewLine & vbNewLine & _
                    "Password or Username is incorrect. Please try again."
                    
        Case E_ErrorCustom.EecErrMmImport
        '// Error for Minimint report.
            ErrTitle = "Incorrect Information"
            ErrMsg = "Please export the Minimint report first!" & vbNewLine & vbNewLine & _
                    "Please verify report from Minimint is valid"
    
        Case Else
        '// General default message.
            ErrTitle = "Error"
            ErrMsg = "Error - something went wrong." & vbNewLine & vbNewLine & _
                     "Please review actions or data and try again."
            
    End Select

    '// Output a message to the user.
    MsgBox _
          Title:=ErrTitle _
        , Prompt:=ErrMsg _
        , Buttons:=vbExclamation + vbOKCancel
        
    '// Clean Exit
    Exit Sub
        
ErrHandler:
    '// General default message.
    ErrTitle = "Error"
    ErrMsg = "Error - something went wrong." & vbNewLine & vbNewLine & _
                     "Please review actions or data and try again."
    
    '// Output a message to the user.
    MsgBox _
          Title:=ErrTitle _
        , Prompt:=ErrMsg _
        , Buttons:=vbExclamation + vbOKCancel
        
    '// Clear any errors.
    On Error GoTo 0
End Sub
