Attribute VB_Name = "RibbonVisibility"
' Module: RibbonVisibility
' Gère la visibilité des éléments du ruban
Option Explicit

' Variable globale pour stocker l'instance du ruban
Public gRibbon As IRibbonUI

' Callback appelé lors du chargement du ruban
Public Sub Ribbon_Load(ByVal ribbon As IRibbonUI)
    Set gRibbon = ribbon
    
    ' Planifie l'initialisation pour s'exécuter dans 1 seconde
    ' Cela permet à l'interface utilisateur de répondre immédiatement.
    Application.OnTime Now + TimeValue("00:00:01"), "DelayedInitialization"
End Sub

' Tâches d'initialisation exécutées après le chargement du ruban
Public Sub DelayedInitialization()
    ' Initialiser le système de logging
    SYS_Logger.InitializeLogger
    
    Log "ribbon", "Ribbon_Load (Delayed) appelé", DEBUG_LEVEL, "DelayedInitialization", "RibbonVisibility"
    
    ' Initialiser les profils de démo
    InitializeDemoProfiles
    Log "ribbon", "gRibbon (Delayed) initialisé", DEBUG_LEVEL, "DelayedInitialization", "RibbonVisibility"
    
    ' Optionnel : rafraîchir le ruban si des états ont changé
    InvalidateRibbon
End Sub

' Callback pour le sélecteur de profil
Public Sub OnSelectDemoProfile(control As IRibbonControl)
    Select Case control.id
        Case "btnEngineerBasic": SetCurrentProfile AccessProfiles.Engineer_Basic
        Case "btnProjectManager": SetCurrentProfile AccessProfiles.Project_Manager
        Case "btnFinanceController": SetCurrentProfile AccessProfiles.Finance_Controller
        Case "btnTechnicalDirector": SetCurrentProfile AccessProfiles.Technical_Director
        Case "btnMultiProjectLead": SetCurrentProfile AccessProfiles.Business_Analyst  ' Changed from Multi_Project_Lead
        Case "btnFullAdmin": SetCurrentProfile AccessProfiles.Full_Admin
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
    Log "ribbon", "InvalidateRibbon appelé", DEBUG_LEVEL, "InvalidateRibbon", "RibbonVisibility"
    If Not gRibbon Is Nothing Then
        gRibbon.Invalidate
        Log "ribbon", "Ribbon invalidé", DEBUG_LEVEL, "InvalidateRibbon", "RibbonVisibility"
    Else
        Log "ribbon", "gRibbon est Nothing", WARNING_LEVEL, "InvalidateRibbon", "RibbonVisibility"
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

