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
    Const PROC_NAME As String = "Ribbon_Load"
    On Error GoTo ErrorHandler
    
    Debug.Print "Ribbon_Load appelé"
    Set gRibbon = ribbon
    
    ' Initialiser le système en mode production avec auto-start
    If Not APP_MainOrchestrator.InitializeElyseSystem(PRODUCTION_MODE, True) Then
        MsgBox "Erreur lors de l'initialisation du système", vbCritical
        Exit Sub
    End If
    
    ' Initialiser le système de messagebox
    Call SYS_MessageBox.InitMessageBoxStatus
    
    InitializeDemoProfiles
    Debug.Print "gRibbon initialisé"
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur lors du chargement du ruban: " & Err.Description, vbCritical
End Sub

' Callback pour le sélecteur de profil
Public Sub OnSelectDemoProfile(control As IRibbonControl)
    Const PROC_NAME As String = "OnSelectDemoProfile"
    On Error GoTo ErrorHandler
    
    Select Case control.id
        Case "btnEngineerBasic": SetCurrentProfile Engineer_Basic
        Case "btnProjectManager": SetCurrentProfile Project_Manager
        Case "btnFinanceController": SetCurrentProfile Finance_Controller
        Case "btnTechnicalDirector": SetCurrentProfile Technical_Director
        Case "btnMultiProjectLead": SetCurrentProfile Business_Analyst
        Case "btnFullAdmin": SetCurrentProfile Full_Admin
    End Select
    
    InvalidateRibbon
    Exit Sub
    
ErrorHandler:
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
End Sub

' Callback pour l'affichage du profil actuel
Public Sub GetCurrentProfileLabel(control As IRibbonControl, ByRef label)
    Const PROC_NAME As String = "GetCurrentProfileLabel"
    On Error GoTo ErrorHandler
    
    label = "Current Profile: " & GetCurrentProfileName()
    Exit Sub
    
ErrorHandler:
    label = "Error loading profile"
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
End Sub

' Callbacks de visibilité avec gestion d'erreurs
Public Sub GetTechnologiesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetTechnologiesVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Engineering")
    Exit Sub
    
ErrorHandler:
    visible = False
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
End Sub

Public Sub GetUtilitiesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetUtilitiesVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Engineering")
    Exit Sub
    
ErrorHandler:
    visible = False
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
End Sub

Public Sub GetServerFilesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetServerFilesVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Tools")
    Exit Sub
    
ErrorHandler:
    visible = False
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
End Sub

Public Sub GetAnalysisToolsVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetAnalysisToolsVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Tools")
    Exit Sub
    
ErrorHandler:
    visible = False
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
End Sub

Public Sub GetFinancesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetFinancesVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Finance")
    Exit Sub
    
ErrorHandler:
    visible = False
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
End Sub

' Fonction pour forcer le rafraîchissement du ruban
Public Sub InvalidateRibbon()
    Const PROC_NAME As String = "InvalidateRibbon"
    On Error GoTo ErrorHandler
    
    If Not gRibbon Is Nothing Then
        gRibbon.Invalidate
    End If
    Exit Sub
    
ErrorHandler:
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
End Sub

' Callback pour la visibilité du menu Debug (admin only)
Public Sub GetAdminVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetAdminVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Admin")
    Exit Sub
    
ErrorHandler:
    visible = False
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
End Sub

' Fonction de visibilité des contrôles
Public Function GetControlVisible(ByVal control As IRibbonControl) As Boolean
    Const PROC_NAME As String = "GetControlVisible"
    On Error GoTo ErrorHandler
    
    Select Case control.id
        Case "adminToolsGroup"
            GetControlVisible = HasAccess("Admin")
        Case Else
            GetControlVisible = True
    End Select
    Exit Function
    
ErrorHandler:
    GetControlVisible = False
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
End Function

' Fonction d'activation des contrôles
Public Function GetControlEnabled(ByVal control As IRibbonControl) As Boolean
    Const PROC_NAME As String = "GetControlEnabled"
    On Error GoTo ErrorHandler
    
    Select Case control.id
        Case "adminToolsGroup"
            GetControlEnabled = HasAccess("Admin")
        Case Else
            GetControlEnabled = True
    End Select
    Exit Function
    
ErrorHandler:
    GetControlEnabled = False
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
End Function

' Callback pour les actions du ruban
Public Sub OnRibbonAction(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnRibbonAction"
    On Error GoTo ErrorHandler
    
    Select Case control.id
        Case "runReportButton"
            ' Appeler la fonction de génération de rapport
            
        Case "settingsButton"
            ' Ouvrir les paramètres
            
        Case Else
            APP_MainOrchestrator.LogWarning "unknown_action", "Unknown ribbon action for control ID: " & control.id
    End Select
    Exit Sub
    
ErrorHandler:
    APP_MainOrchestrator.HandleVBAError PROC_NAME, MODULE_NAME
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

' Simple test button callback
Public Sub OnTestButton(control As IRibbonControl)
    MsgBox "Test button clicked!"
End Sub


