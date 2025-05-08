Attribute VB_Name = "Mod_EnumsManager"
'// All Enums are to be listed here for the project
Option Explicit
Option Private Module

'// List of Enums for action manager to run subs.
    Enum E_ActionManager
        EamAddCommandMenu
        EamArchiveWorkSheets
        EamAutoConfirmPo
        EamAutoAdjustDiff
        EamAutoDeliverOne
        EamCreateDailyReport
        EamCreateWeeklyReport
        EamCreateWorksheets
        EamImportCcoCoid
        EamImportShiftReport
        EamImportMixCommits
        EamImportPrismaCommits
        EamImportCoid
        EamImportProdReport
        EamProtectWorkSheet
        EamRemoveFormulae
        EamRepositionHome
        EamRepositionCoid
        EamRepositionPrisma
        EamRepositionShiftReport
        EamViewDailyCoid
        EamViewShiftReport
        EamViewAdminSheet
        EamVeryHiddenSheets
        EamUnhideAllSheets
        EamCalculateMixes
    End Enum
    
'// List of Error Numbers + vbObjectError constant. Start high as to not interfere with defalut values.
    Enum E_ErrorCustom
        EecErrGeneral = vbObjectError + 1000
        EecErrImport
        EecErrExport
        EecErrSapGui
        EecErrSercurity
        EecErrMmImport
    End Enum
    
'// List of SAP Tcodes used in this project
    Enum E_SapTcodes
        EstCoid
        EstCor2
        EstCor6
        EstCor3
    End Enum
    
'// List of key row numbers used in this project.
    Enum E_KeyRows
        EkrStartRow = 9
        EkrClearRange = 82
        EkrDataRow = 100
    End Enum

'// List of key column numbers used in this project.
    Enum E_DataColumns
        EdcShiftImport = 79
        EdcSapCaseImport = 105
        EdcSapMixImport = 122
        EdcPrismaImport = 134
    End Enum
    
'// List of Key prcocess order columns.
    Enum E_ProcOrders
        EpoProcessOder = 1
        EpoMoaNumber
        EpoDescription
        EpoBatchNumber
        EpoUnitMetric
        EpoProdProfile
        EpoOrderVersion
    End Enum
    
