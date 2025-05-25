' Module: RibbonVisibility
' Gère la visibilité des éléments du ruban
Option Explicit

' ============================================================================
' MODULE DEPENDENCIES
' ============================================================================
' ElyseMain_Orchestrator est accessible via GetInstance()
' Note: SYS_MessageBox et SYS_CoreSystem sont des modules, pas besoin de les instancier

Private Const MODULE_NAME As String = "RibbonVisibility"

' Variable globale pour stocker l'instance du ruban
Public gRibbon As IRibbonUI

' Callback appelé lors du chargement du ruban
Public Sub Ribbon_Load(ByVal ribbon As IRibbonUI)
    Debug.Print "Ribbon_Load appelé"
    Set gRibbon = ribbon
    ' S'assurer que le système est initialisé avant d'utiliser les profils
    ElyseMain_Orchestrator.Initialize
    InitializeDemoProfiles
    Debug.Print "gRibbon initialisé"
End Sub

' Callback pour le sélecteur de profil
Public Sub OnSelectDemoProfile(control As IRibbonControl)
    Select Case control.id
        Case "btnEngineerBasic": SetCurrentProfile Engineer_Basic
        Case "btnProjectManager": SetCurrentProfile Project_Manager
        Case "btnFinanceController": SetCurrentProfile Finance_Controller
        Case "btnTechnicalDirector": SetCurrentProfile Technical_Director
        Case "btnMultiProjectLead": SetCurrentProfile Business_Analyst  ' Changed from Multi_Project_Lead
        Case "btnFullAdmin": SetCurrentProfile Full_Admin
    End Select
    
    InvalidateRibbon
End Sub

' Callback pour l'affichage du profil actuel
Public Sub GetCurrentProfileLabel(control As IRibbonControl, ByRef label)
    label = "Current Profile: " & GetCurrentProfileName()
End Sub

' Callback pour la visibilité du menu Technologies
Public Sub GetTechnologiesVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = HasAccess("Engineering")
End Sub

' Callback pour la visibilité du menu Utilities
Public Sub GetUtilitiesVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = HasAccess("Engineering")
End Sub

' Callback pour la visibilité du menu Files
Public Sub GetServerFilesVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = HasAccess("Tools")
End Sub

' Callback pour la visibilité du menu Outils
Public Sub GetAnalysisToolsVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = HasAccess("Tools")
End Sub

' Callback pour la visibilité du menu Finances
Public Sub GetFinancesVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = HasAccess("Finance")
End Sub

' Callback pour la visibilité des menus de projets
Public Function GetProjectMenuVisibility(projectMenu As String) As Boolean
    ' Extrait le nom du projet du menu (par exemple "Echo" de "summaryEcho")
    Dim projectName As String
    If InStr(1, projectMenu, "GENERIC") > 0 Then
        GetProjectMenuVisibility = HasAccess("Engineering") ' Seuls les ingénieurs voient les génériques
    Else
        projectName = Replace(projectMenu, "summary", "")
        projectName = Replace(projectName, "planning", "")
        projectName = Replace(projectName, "devex", "")
        projectName = Replace(projectName, "capex", "")
        projectName = Replace(projectName, "opex", "")
        projectName = Replace(projectName, "tech", "")
        GetProjectMenuVisibility = HasAccess(projectName)
    End If
End Function

Public Sub GetSummarySheetsVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = GetProjectMenuVisibility(control.id)
End Sub

Public Sub GetPlanningsVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = GetProjectMenuVisibility(control.id)
End Sub

Public Sub GetDevexVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = HasAccess("Finance") Or GetProjectMenuVisibility(control.id)
End Sub

Public Sub GetCapexVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = HasAccess("Finance") Or GetProjectMenuVisibility(control.id)
End Sub

Public Sub GetOpexVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = HasAccess("Finance") Or GetProjectMenuVisibility(control.id)
End Sub

Public Sub GetTechScenariosVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = HasAccess("Engineering") Or GetProjectMenuVisibility(control.id)
End Sub

' Fonction pour forcer le rafraîchissement du ruban
Public Sub InvalidateRibbon()
    Debug.Print "InvalidateRibbon appelé"
    If Not gRibbon Is Nothing Then
        gRibbon.Invalidate
        Debug.Print "Ribbon invalidé"
    Else
        Debug.Print "gRibbon est Nothing"
    End If
End Sub

' Simple test button callback
Public Sub OnTestButton(control As IRibbonControl)
    MsgBox "Test button clicked!"
End Sub

' Callback pour la visibilité du menu Debug (admin only)
Public Sub GetAdminVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = HasAccess("Admin")
End Sub

' Example of a getVisible callback for a ribbon control
Public Sub GetControlVisible(control As IRibbonControl, ByRef returnedVal As Variant)
    Const PROC_NAME As String = "GetControlVisible"
    On Error GoTo ErrorHandler ' Keep error handling light in UI callbacks

    ' ElyseMain_Orchestrator.LogDebug PROC_NAME & "_Start", "Checking visibility for control: " & control.id, PROC_NAME, MODULE_NAME
    ' Using a lighter log for UI callbacks to avoid flooding, or make it conditional
    If ElyseCore_System.IsDebugMode() Then
        ElyseMain_Orchestrator.LogDebug PROC_NAME & "_Callback", "GetControlVisible called for: " & control.id, PROC_NAME, MODULE_NAME
    End If

    Select Case control.id
        Case "customButton1"
            ' returnedVal = CheckUserAccess("someUser") ' Example: original logic
            returnedVal = True ' Placeholder, replace with actual logic
            ' ElyseMain_Orchestrator.LogDebug PROC_NAME & "_Button1Visible", "Visibility for customButton1 set to " & returnedVal, PROC_NAME, MODULE_NAME
        Case "adminToolsGroup"
            ' returnedVal = IsCurrentUserAdmin() ' Example: original logic
            returnedVal = False ' Placeholder
            ' ElyseMain_Orchestrator.LogDebug PROC_NAME & "_AdminGroupVisible", "Visibility for adminToolsGroup set to " & returnedVal, PROC_NAME, MODULE_NAME
        Case Else
            returnedVal = True ' Default to visible
    End Select
    
    Exit Sub

ErrorHandler:
    ' In Ribbon callbacks, avoid complex error handling that might show UI (MsgBox)
    ' Log the error and ensure a default value is returned.
    ElyseMain_Orchestrator.LogError PROC_NAME & "_Error", Err.Number, "Error in GetControlVisible for " & control.id & ": " & Err.Description, PROC_NAME, MODULE_NAME
    returnedVal = False ' Default to not visible/disabled on error to be safe
End Sub

' Example of a getEnabled callback
Public Sub GetControlEnabled(control As IRibbonControl, ByRef returnedVal As Variant)
    Const PROC_NAME As String = "GetControlEnabled"
    On Error GoTo ErrorHandler

    If ElyseCore_System.IsDebugMode() Then
        ElyseMain_Orchestrator.LogDebug PROC_NAME & "_Callback", "GetControlEnabled called for: " & control.id, PROC_NAME, MODULE_NAME
    End If

    Select Case control.id
        Case "customButtonSave"
            ' returnedVal = ActiveWorkbook.Saved ' Example: original logic
            returnedVal = Not ActiveWorkbook.Saved ' Enable if not saved
        Case Else
            returnedVal = True ' Default to enabled
    End Select
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.LogError PROC_NAME & "_Error", Err.Number, "Error in GetControlEnabled for " & control.id & ": " & Err.Description, PROC_NAME, MODULE_NAME
    returnedVal = False ' Default to disabled on error
End Sub

' Example of an onAction callback
Public Sub RibbonButton_OnAction(control As IRibbonControl)
    Const PROC_NAME As String = "RibbonButton_OnAction"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogRibbonAction control.id, "User clicked ribbon button."

    Select Case control.id
        Case "btnRunReport"
            ElyseMain_Orchestrator.LogInfo PROC_NAME & "_RunReport", "User initiated RunReport.", PROC_NAME, MODULE_NAME
            ' Call RunReport_Sub ' Example call
            ElyseMessageBox_System.ShowInfoMessage "Action", "Running Report... (placeholder)"
        Case "btnOpenSettings"
            ElyseMain_Orchestrator.LogInfo PROC_NAME & "_OpenSettings", "User initiated OpenSettings.", PROC_NAME, MODULE_NAME
            ' Call OpenSettings_Form.Show ' Example call
            ElyseMessageBox_System.ShowInfoMessage "Action", "Opening Settings... (placeholder)"
        Case Else
            ElyseMain_Orchestrator.LogWarning PROC_NAME & "_UnknownAction", "Unknown ribbon action for control ID: " & control.id, PROC_NAME, MODULE_NAME
            ElyseMessageBox_System.ShowWarningMessage "Unknown Action", "The action for '" & control.id & "' is not defined."
    End Select
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME, "Error in Ribbon Action: " & control.id
    ' Optionally, show a generic error message to the user via the new system
    ElyseMessageBox_System.ShowErrorMessage "Ribbon Error", "An unexpected error occurred while processing the action for '" & control.id & "'. The error has been logged."
End Sub


