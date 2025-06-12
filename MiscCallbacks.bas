Attribute VB_Name = "MiscCallbacks"
Option Explicit

Private Const MODULE_NAME As String = "MiscCallbacks"

' Ce module contient des callbacks qui n'ont pas encore trouvé leur place
' dans un module métier spécifique. C'est un emplacement temporaire.

Public Sub OnOpex(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnOpex"
    On Error GoTo ErrorHandler
    MsgBox "Fonctionnalité non implémentée.", vbInformation, "Info"
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "callback_error", "Erreur dans OnOpex: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ReloadAllTablesCallback(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ReloadAllTablesCallback"
    On Error GoTo ErrorHandler
    MsgBox "Fonctionnalité non implémentée.", vbInformation, "Info"
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "callback_error", "Erreur dans ReloadAllTablesCallback: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ReloadCurrentTableClick(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ReloadCurrentTableClick"
    On Error GoTo ErrorHandler
    ' TODO: Implémenter la logique
    MsgBox "Logique de rechargement à implémenter pour : " & ActiveCell.ListObject.Name
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "callback_error", "Erreur dans ReloadCurrentTableClick: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub DeleteCurrentTableClick(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "DeleteCurrentTableClick"
    On Error GoTo ErrorHandler
    ' TODO: Implémenter la logique
    MsgBox "Logique de suppression à implémenter pour : " & ActiveCell.ListObject.Name
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "callback_error", "Erreur dans DeleteCurrentTableClick: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnServerFiles(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnServerFiles"
    On Error GoTo ErrorHandler
    MsgBox "Fonctionnalité non implémentée.", vbInformation, "Info"
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "callback_error", "Erreur dans OnServerFiles: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnTechEcoAnalysis(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnTechEcoAnalysis"
    On Error GoTo ErrorHandler
    MsgBox "Fonctionnalité non implémentée.", vbInformation, "Info"
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "callback_error", "Erreur dans OnTechEcoAnalysis: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnBusinessPlan(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnBusinessPlan"
    On Error GoTo ErrorHandler
    MsgBox "Fonctionnalité non implémentée.", vbInformation, "Info"
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "callback_error", "Erreur dans OnBusinessPlan: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnElecOptim(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnElecOptim"
    On Error GoTo ErrorHandler
    MsgBox "Fonctionnalité non implémentée.", vbInformation, "Info"
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "callback_error", "Erreur dans OnElecOptim: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnLCA(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnLCA"
    On Error GoTo ErrorHandler
    MsgBox "Fonctionnalité non implémentée.", vbInformation, "Info"
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "callback_error", "Erreur dans OnLCA: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnSummarySheets(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnSummarySheets"
    On Error GoTo ErrorHandler
    MsgBox "Fonctionnalité non implémentée.", vbInformation, "Info"
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "callback_error", "Erreur dans OnSummarySheets: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME
End Sub 