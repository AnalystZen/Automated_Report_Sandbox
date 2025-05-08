Attribute VB_Name = "Mod_ActionManager"
Option Explicit
Option Private Module
'// This is where all subs will be ran for this project. All actions, error handling and variable initialization will happen here.
'// Select case statement is split bewteen procedures and events.

Sub FunnelAction(ByVal WhatAction As E_ActionManager)
'// Manage All Actions that Can be Taken in This Project in one place,

    '// Enable Error trapping.
    On Error GoTo ErrorHandler
    
    '// Turn off Apps for faster subs, effects all subs.
    TurnOffApps
        
    '// This will perform the designated sub based on the value passed to it.
    Select Case WhatAction
            
        Case E_ActionManager.EamArchiveWorkSheets
            
            ActionArchiveWorkSheets
            
        Case E_ActionManager.EamAutoAdjustDiff
            
            ActionAutoAdjustDiff
            
        Case E_ActionManager.EamAutoConfirmPo
            
            ActionAutoConfirmPo
            
        Case E_ActionManager.EamAutoDeliverOne
            
            ActionAutoDeliverOne
            
        Case E_ActionManager.EamCreateDailyReport
            
            ActionCreateDailyReport
            
        Case E_ActionManager.EamCreateWeeklyReport
            
            ActionCreateWeeklyReport
            
        Case E_ActionManager.EamCreateWorksheets
            
            ActionCreateWorksheets
            
        Case E_ActionManager.EamImportCcoCoid
            
            ActionImportCcoCoid
            
        Case E_ActionManager.EamImportCoid
            
            ActionImportCoid
            
        Case E_ActionManager.EamImportMixCommits
            
            ActionImportMixCommits
            
        Case E_ActionManager.EamImportPrismaCommits
            
            ActionImportPrismaCommits
            
        Case E_ActionManager.EamImportShiftReport
            
            ActionImportShiftReport
            
        Case E_ActionManager.EamRemoveFormulae
            
            ActionRemoveFormulae
            
        Case E_ActionManager.EamRepositionCoid
            
            ActionRepositionCoid
            
        Case E_ActionManager.EamRepositionHome
            
            ActionRepositionHome
            
        Case E_ActionManager.EamRepositionPrisma
            
            ActionRepositionPrisma
            
        Case E_ActionManager.EamViewDailyCoid
            
            ActionViewDailyCoid
            
        Case E_ActionManager.EamViewShiftReport
            
            ActionViewShiftReport
            
        Case E_ActionManager.EamProtectWorkSheet
        
            ActionProtectWorksheet
            
        Case E_ActionManager.EamViewAdminSheet
            
            ActionViewAdminSheet
            
        Case E_ActionManager.EamImportProdReport
            
            ActionImportProdReport
            
        Case E_ActionManager.EamRepositionShiftReport
            
            ActionRepositionShiftReport
            
        Case E_ActionManager.EamCalculateMixes
        
            ActionRecalculateMixes
            
    End Select
    
    
    
    '// Workbook and sheet events go here.
    Select Case WhatAction
    
        Case E_ActionManager.EamAddCommandMenu
            
            ActionAddCommandMenu
            
        Case E_ActionManager.EamUnhideAllSheets
            
            ActionViewVeryHiddenSheets
            
        Case E_ActionManager.EamVeryHiddenSheets
            
            ActionAllSheetsVeryHidden
            
    End Select
    
    '// Restore Application Settings
    TurnOnApps
    
    '// Clean exit.
    Exit Sub
    
ErrorHandler:

    '// Call the Error Handler to generate the correct output.
     ErrorManager Err.Number, Err.Description
    
    '// Disable error handling.
    On Error Resume Next
    
    '// Restore Application Settings
    TurnOnApps

    '// Clear Any Errors.Not neccassary.
    On Error GoTo 0
    
End Sub

Sub ActionAddCommandMenu()
'// installs buttons for user interface.
    
    Run "Sp_UserSelectInterface"

End Sub
Sub ActionArchiveWorkSheets()
'// Archives worksheets to selected destination.
    
    '// Verify if admin allows archive of sheets
    If ThisWorkbook.Worksheets("Admin").Range("B5") = "No" Then
        MsgBox "Archive of worksheets currently not allowed.", vbExclamation, "Permission Needed"
        Exit Sub
    End If
    
    Sp_ArchiveSheets

End Sub
Sub ActionAutoConfirmPo()
'// Auto confirms po'//s in sap.
    
    Sp_AutoConfirmPOs
    
End Sub
Sub ActionAutoAdjustDiff()
'// Auto adjust po confirmatin in sap.

    Sp_AutoAdjustDiff
    
End Sub
Sub ActionAutoDeliverOne()
'// Auto changes delivery target to 1 in sap.
    
    Sp_AutoAdjustDelivery
    
End Sub
Sub ActionCreateDailyReport()
'// Creates daily report.
    
    Sp_CreateDailyReport
    
End Sub
Sub ActionCreateWeeklyReport()
'// Creates weekly report
    
    Sp_CreateWeeklyMixReport
    
End Sub
Sub ActionCreateWorksheets()
'// Create new worksheets.

    Sp_CreateSheets
    
End Sub
Sub ActionImportCcoCoid()
'// Import daily process orders.
    
    Sp_CcoCoidImport
    
End Sub
Sub ActionImportShiftReport()
'// Import shift reports.
    
    Sp_ImportShiftReport
    
End Sub
Sub ActionImportMixCommits()
'// Import Sap mixes.

    Sp_ImportSapMixCommits
    
End Sub
Sub ActionImportPrismaCommits()
'// Import prisma mixes.
    
    Sp_ImportPrismaCommits
    
End Sub
Sub ActionImportCoid()
'// Import coid cases.
    
    Sp_ImportCoidCases
End Sub
Sub ActionRemoveFormulae()
'// Removes formulas from sheets.

    '// Verify if admin allows archive of sheets
    If ThisWorkbook.Worksheets("Admin").Range("B7") = "No" Then
        MsgBox "Removal of formulae currently not allowed.", vbExclamation, "Permission Needed"
        Exit Sub
    End If
    
    Sp_RemoveFormulae
    
End Sub
Sub ActionRepositionHome()
'// Position screen to A1.

    Sp_RepositionHome
    
End Sub
Sub ActionRepositionCoid()
'// Position screen to Coid

    Sp_RepositionCoid
    
End Sub
Sub ActionRepositionPrisma()
'// Position screen to prisma report.

    Sp_RepositionPrisma
    
End Sub
Sub ActionRepositionShiftReport()
'// Position screen to Shift Report.

    Sp_RepositionShiftReport
    
End Sub
Sub ActionViewDailyCoid()
'// View Daily coid.
    
    Sp_OpenDailyCoid
    
End Sub
Sub ActionViewShiftReport()
'// View daily shift report.
    
    Sp_OpenShiftReport
    
End Sub

Sub ActionProtectWorksheet()
'// Protect and unprotect worksheet.

  '// Disables toggle interface button from accessing master worksheet.
    If ActiveSheet.Name = "Master Worksheet" Then
        MsgBox _
        Prompt:="This procedure does not have access to unlock Master Sheet." _
        , Buttons:=vbExclamation + vbOKOnly _
        , Title:="Not Authorized"
        Exit Sub
    End If
    
    '// Call sub procedure
    Sp_ToggleProtection

End Sub
Sub ActionViewAdminSheet()
'// This action will show the admin sheet.

    Sp_ShowAdmin
    
End Sub

Sub ActionViewVeryHiddenSheets()
'// This action will unhide all very hidden sheets.

    Sp_ViewAllWsHidden False

End Sub

Sub ActionAllSheetsVeryHidden()
'// This action will hide all sheets.
    
    Sp_ViewAllWsHidden True

End Sub

Sub ActionImportProdReport()
'// This action will import the production case report.
     
    Sp_ImportProdReport
        
End Sub

Sub ActionRecalculateMixes()
'// This action will recalculate mixes from the shift reports

    Sp_CalculateShiftMixes
    
End Sub

