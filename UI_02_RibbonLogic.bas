' Module: RibbonVisibility
' Gère la visibilité des éléments du ruban
Option Explicit

' ============================================================================
' MODULE DEPENDENCIES
' ============================================================================
' ElyseMain_Orchestrator est accessible via GetInstance()
' Note: SYS_MessageBox et SYS_CoreSystem sont des modules, on les utilise directement sans déclaration

Private Const MODULE_NAME As String = "RibbonVisibility"

' Variable globale pour stocker l'instance du ruban
Public gRibbon As IRibbonUI

' Callback appelé lors du chargement du ruban
Public Sub Ribbon_Load(ByVal ribbon As IRibbonUI)
    Debug.Print "Ribbon_Load appelé"
    Set gRibbon = ribbon
    ' S'assurer que le système est initialisé avant d'utiliser les profils
    APP_MainOrchestrator.Initialize
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
Public Function GetControlVisible(ByVal control As IRibbonControl) As Boolean
    Const PROC_NAME As String = "GetControlVisible"
    On Error GoTo ErrorHandler
    
    LogDebug PROC_NAME & "_Callback", "GetControlVisible called for: " & control.id, PROC_NAME, MODULE_NAME
    
    ' Logique de visibilité selon l'ID du contrôle
    Select Case control.id
        Case "customButton1"
            GetControlVisible = True ' Exemple simple
            LogDebug PROC_NAME & "_Button1Visible", "Visibility for customButton1 set to " & GetControlVisible, PROC_NAME, MODULE_NAME
            
        Case "adminToolsGroup"
            GetControlVisible = IsUserAdmin() ' Vérifier les droits admin
            LogDebug PROC_NAME & "_AdminGroupVisible", "Visibility for adminToolsGroup set to " & GetControlVisible, PROC_NAME, MODULE_NAME
            
        Case Else
            GetControlVisible = True ' Par défaut visible
    End Select
    
    Exit Function

ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Error in GetControlVisible for " & control.id & ": " & Err.Description, PROC_NAME, MODULE_NAME
    GetControlVisible = False ' En cas d'erreur, on cache le contrôle
End Function

' Example of a getEnabled callback
Public Function GetControlEnabled(ByVal control As IRibbonControl) As Boolean
    Const PROC_NAME As String = "GetControlEnabled"
    On Error GoTo ErrorHandler
    
    LogDebug PROC_NAME & "_Callback", "GetControlEnabled called for: " & control.id, PROC_NAME, MODULE_NAME
    
    ' Logique d'activation selon l'ID du contrôle
    Select Case control.id
        Case "customButton1"
            GetControlEnabled = True ' Exemple simple
            
        Case "adminToolsGroup"
            GetControlEnabled = IsUserAdmin() ' Vérifier les droits admin
            
        Case Else
            GetControlEnabled = True ' Par défaut activé
    End Select
    
    Exit Function

ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Error in GetControlEnabled for " & control.id & ": " & Err.Description, PROC_NAME, MODULE_NAME
    GetControlEnabled = False ' En cas d'erreur, on désactive le contrôle
End Function

' Example of an onAction callback
Public Sub OnRibbonAction(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnRibbonAction"
    On Error GoTo ErrorHandler
    
    LogRibbonAction control.id, "User clicked ribbon button."
    
    ' Gérer les actions selon l'ID du contrôle
    Select Case control.id
        Case "runReportButton"
            LogInfo PROC_NAME & "_RunReport", "User initiated RunReport.", PROC_NAME, MODULE_NAME
            ' Appeler la fonction de génération de rapport
            
        Case "settingsButton"
            LogInfo PROC_NAME & "_OpenSettings", "User initiated OpenSettings.", PROC_NAME, MODULE_NAME
            ' Ouvrir les paramètres
            
        Case Else
            LogWarning PROC_NAME & "_UnknownAction", "Unknown ribbon action for control ID: " & control.id, PROC_NAME, MODULE_NAME
    End Select
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Error in Ribbon Action: " & control.id
End Sub


