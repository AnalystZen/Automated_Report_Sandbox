Attribute VB_Name = "Mod_InterfaceManager"
'// All click events the user has access to is to be listed here. They will funnel to the action manager.

Private Sub Sp_UserSelectInterface()
'// Establish Menu in Addin Tab.

    '//Declare variables
    Dim cmdbar As CommandBar
    Dim cmdbutton As CommandBarButton
    
    '// Assign Values
    Set cmdbar = Application.CommandBars(1)
    Set cmdbutton = cmdbar.Controls.Add(Type:=msoControlButton, temporary:=True)
    
    '//////////////////////////////////////////////////////////////////////////'//
    With cmdbutton
        .Style = msoButtonIconAndCaption
        .Caption = "Reposition Home"
        .FaceId = 490
        .OnAction = "ClickRepositionHome"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Reposition Reports"
        .FaceId = 491
        .OnAction = "ClickRepositionShiftReport"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Reposition Coid"
        .FaceId = 492
        .OnAction = "ClickRepositionCoid"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Reposition Prisma"
        .FaceId = 493
        .OnAction = "ClickRepositionPrisma"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Open Coid By Date"
        .FaceId = 9718
        .OnAction = "ClickViewDailyCoid"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Import Coid By Date"
        .FaceId = 651
        .OnAction = "ClickImportCcoCoid"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Import Shift Report"
        .FaceId = 6991
        .OnAction = "ClickImportShiftReport"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "View Shift Report"
        .FaceId = 4030
        .OnAction = "ClickViewShiftReport"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Recalculate Mixes"
        .FaceId = 307
        .OnAction = "ClickRecalculateMixes"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Import SAP Mixes"
        .FaceId = 168
        .OnAction = "ClickImportMixCommits"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Import Minmint Mixes"
        .FaceId = 2653
        .OnAction = "ClickImportPrismaCommits"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Import SAP Goods"
        .FaceId = 6173
        .OnAction = "ClickImportCoid"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Import Prod Report"
        .FaceId = 163
        .OnAction = "ClickImportProdReport"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Auto Confirm PO's"
        .FaceId = 7431
        .OnAction = "ClickAutoConfirmPo"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Auto Adjust Conf."
        .FaceId = 62
        .OnAction = "ClickAutoAdjustDiff"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Auto Deliver One"
        .FaceId = 71
        .OnAction = "ClickAutoDeliverOne"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Create Daily Report"
        .FaceId = 422
        .OnAction = "ClickCreateDailyReport"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Create Weekly Report"
        .FaceId = 7800
        .OnAction = "ClickCreateWeeklyReport"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Create Worksheets"
        .FaceId = 3282
        .OnAction = "ClickCreateWorksheets"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Toggle Protection"
        .FaceId = 6243
        .OnAction = "ClickProtectWorkSheet"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Remove Formulae"
        .FaceId = 893
        .OnAction = "ClickRemoveFormulae"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Archive Work Sheets"
        .FaceId = 270
        .OnAction = "ClickArchiveWorkSheets"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
    With Application.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
        .Style = msoButtonIconAndCaption
        .Caption = "Admin Access"
        .FaceId = 351
        .OnAction = "ClickViewAdminAccess"
    End With
    '//////////////////////////////////////////////////////////////////////////'//
End Sub

'// All click events the user has access to are to be listed here. They funnel to the action manager.

Sub ClickArchiveWorkSheets()
    FunnelAction EamArchiveWorkSheets
End Sub
Sub ClickAutoConfirmPo()
    FunnelAction EamAutoConfirmPo
End Sub
Sub ClickAutoAdjustDiff()
    FunnelAction EamAutoAdjustDiff
End Sub
Sub ClickAutoDeliverOne()
    FunnelAction EamAutoDeliverOne
End Sub
Sub ClickCreateDailyReport()
    FunnelAction EamCreateDailyReport
End Sub
Sub ClickCreateWeeklyReport()
    FunnelAction EamCreateWeeklyReport
End Sub
Sub ClickCreateWorksheets()
    FunnelAction EamCreateWorksheets
End Sub
Sub ClickImportCcoCoid()
    FunnelAction EamImportCcoCoid
End Sub
Sub ClickImportShiftReport()
    FunnelAction EamImportShiftReport
End Sub
Sub ClickImportMixCommits()
    FunnelAction EamImportMixCommits
End Sub
Sub ClickImportPrismaCommits()
    FunnelAction EamImportPrismaCommits
End Sub
Sub ClickImportCoid()
    FunnelAction EamImportCoid
End Sub
Sub ClickProtectWorkSheet()
    FunnelAction EamProtectWorkSheet
End Sub
Sub ClickRemoveFormulae()
    FunnelAction EamRemoveFormulae
End Sub
Sub ClickRepositionHome()
    FunnelAction EamRepositionHome
End Sub
Sub ClickRepositionCoid()
    FunnelAction EamRepositionCoid
End Sub
Sub ClickRepositionPrisma()
    FunnelAction EamRepositionPrisma
End Sub
Sub ClickViewDailyCoid()
    FunnelAction EamViewDailyCoid
End Sub
Sub ClickViewShiftReport()
    FunnelAction EamViewShiftReport
End Sub
Sub ClickViewAdminAccess()
    FunnelAction EamViewAdminSheet
End Sub
Sub ClickShowUserForm()
'// Call user form.
    FrmUserName.Show
End Sub
Sub ClickImportProdReport()
'// Import Production case report from sap.
     FunnelAction EamImportProdReport
End Sub
Sub ClickRepositionShiftReport()
'// View shift reports on sheet.
    FunnelAction EamRepositionShiftReport
End Sub
Sub ClickRecalculateMixes()
'// Recalculate mixes for shift reports.
    FunnelAction EamCalculateMixes
End Sub
