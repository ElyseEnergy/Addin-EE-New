Attribute VB_Name = "RibbonVisibility"
' Module: RibbonVisibility
' Gère la visibilité des éléments du ruban
Option Explicit

' Ajouter en haut du module, après Option Explicit
Private Const MODULE_NAME As String = "RibbonVisibility"

' Ajouter en haut du module, après Option Explicit et MODULE_NAME
Private Const TABLE_PREFIX As String = "EE_"

' Variable globale pour stocker l'instance du ruban
Public gRibbon As IRibbonUI

' Callback appelé lors du chargement du ruban
Public Sub Ribbon_Load(ByVal ribbon As IRibbonUI)
    Const PROC_NAME As String = "Ribbon_Load"
    On Error GoTo ErrorHandler
    
    Set gRibbon = ribbon
    
    ' Planifie l'initialisation pour s'exécuter dans 1 seconde
    ' Cela permet à l'interface utilisateur de répondre immédiatement.
    Application.OnTime Now + TimeValue("00:00:01"), "DelayedInitialization"
    
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Tâches d'initialisation exécutées après le chargement du ruban
Public Sub DelayedInitialization()
    Const PROC_NAME As String = "DelayedInitialization"
    On Error GoTo ErrorHandler
    
    ' Initialiser le système de logging
    SYS_Logger.InitializeLogger
    
    Log "ribbon", "Ribbon_Load (Delayed) appelé", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    ' Initialiser les profils de démo
    InitializeDemoProfiles
    Log "ribbon", "gRibbon (Delayed) initialisé", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    ' Optionnel : rafraîchir le ruban si des états ont changé
    InvalidateRibbon
    
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour le sélecteur de profil
Public Sub OnSelectDemoProfile(control As IRibbonControl)
    Const PROC_NAME As String = "OnSelectDemoProfile"
    On Error GoTo ErrorHandler
    
    Select Case control.id
        Case "btnEngineerBasic": SetCurrentProfile AccessProfiles.Engineer_Basic
        Case "btnProjectManager": SetCurrentProfile AccessProfiles.Project_Manager
        Case "btnFinanceController": SetCurrentProfile AccessProfiles.Finance_Controller
        Case "btnTechnicalDirector": SetCurrentProfile AccessProfiles.Technical_Director
        Case "btnMultiProjectLead": SetCurrentProfile AccessProfiles.Business_Analyst  ' Changed from Multi_Project_Lead
        Case "btnFullAdmin": SetCurrentProfile AccessProfiles.Full_Admin
    End Select
    
    InvalidateRibbon
    
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour l'affichage du profil actuel
Public Sub GetCurrentProfileLabel(control As IRibbonControl, ByRef label)
    Const PROC_NAME As String = "GetCurrentProfileLabel"
    On Error GoTo ErrorHandler
    
    label = "Current Profile: " & GetCurrentProfileName()
    
    Exit Sub
ErrorHandler:
    label = "Error"
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour la visibilité du menu Technologies
Public Sub GetTechnologiesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetTechnologiesVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Engineering")
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour la visibilité du menu Utilities
Public Sub GetUtilitiesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetUtilitiesVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Engineering")
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour la visibilité du menu Files
Public Sub GetServerFilesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetServerFilesVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Tools")
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour la visibilité du menu Outils
Public Sub GetAnalysisToolsVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetAnalysisToolsVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Tools")
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour la visibilité du menu Finances
Public Sub GetFinancesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetFinancesVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Finance")
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour la visibilité des menus de projets
Public Function GetProjectMenuVisibility(projectMenu As String) As Boolean
    Const PROC_NAME As String = "GetProjectMenuVisibility"
    On Error GoTo ErrorHandler
    
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
    
    Exit Function
ErrorHandler:
    GetProjectMenuVisibility = False
    HandleError MODULE_NAME, PROC_NAME
End Function

Public Sub GetSummarySheetsVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetSummarySheetsVisibility"
    On Error GoTo ErrorHandler
    
    visible = GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub GetPlanningsVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetPlanningsVisibility"
    On Error GoTo ErrorHandler
    
    visible = GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub GetDevexVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetDevexVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Finance") Or GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub GetCapexVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetCapexVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Finance") Or GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub GetOpexVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetOpexVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Finance") Or GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub GetTechScenariosVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetTechScenariosVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Engineering") Or GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour la visibilité du bouton d'upload
Public Sub GetUploadButtonVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetUploadButtonVisibility"
    On Error GoTo ErrorHandler
    
    ' Visible uniquement si l'utilisateur a accès aux fichiers serveur
    visible = HasAccess("Files")
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Fonction pour forcer le rafraîchissement du ruban
Public Sub InvalidateRibbon()
    Const PROC_NAME As String = "InvalidateRibbon"
    On Error GoTo ErrorHandler
    
    Log "ribbon", "InvalidateRibbon appelé", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    If Not gRibbon Is Nothing Then
        gRibbon.Invalidate
        Log "ribbon", "Ribbon invalidé", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Else
        Log "ribbon", "gRibbon est Nothing", WARNING_LEVEL, PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Simple test button callback
Public Sub OnTestButton(control As IRibbonControl)
    Const PROC_NAME As String = "OnTestButton"
    On Error GoTo ErrorHandler
    
    MsgBox "Test button clicked!"
    
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour la visibilité du menu Debug (admin only)
Public Sub GetAdminVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetAdminVisibility"
    On Error GoTo ErrorHandler
    
    visible = HasAccess("Admin")
    
    Exit Sub
ErrorHandler:
    visible = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour la visibilité des boutons de rechargement.
' Le ruban appelle cette fonction pour savoir s'il doit afficher les boutons.
'---------------------------------------------------------------------------------------
Public Sub GetReloadButtonsVisible(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetReloadButtonsVisible"
    On Error GoTo ErrorHandler
    ' La logique est simple : les boutons sont toujours visibles.
    visible = True
Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

'---------------------------------------------------------------------------------------
' Callback pour l'état activé du bouton "Recharger le tableau courant".
' Le ruban appelle cette fonction pour savoir si le bouton doit être cliquable.
'---------------------------------------------------------------------------------------
Public Sub GetReloadCurrentEnabled(control As IRibbonControl, ByRef enabled As Variant)
    Const PROC_NAME As String = "GetReloadCurrentEnabled"
    On Error GoTo ErrorHandler
    
    Dim currentTable As ListObject
    enabled = False ' Désactivé par défaut

    On Error Resume Next
    Set currentTable = ActiveCell.ListObject
    On Error GoTo ErrorHandler
    
    If Not currentTable Is Nothing Then
        Dim hasComment As Boolean
        hasComment = False
        On Error Resume Next
        ' Vérifie si la cellule en haut à gauche du tableau a un commentaire non vide.
        hasComment = (Len(currentTable.Range.Cells(1, 1).Comment.Text) > 0)
        On Error GoTo ErrorHandler ' Rétablir la gestion d'erreur

        If currentTable.Name Like TABLE_PREFIX & "*" And hasComment Then
            enabled = True
        End If
    End If

    Exit Sub
ErrorHandler:
    enabled = False
    HandleError MODULE_NAME, PROC_NAME
End Sub

'---------------------------------------------------------------------------------------
' Callback pour l'état activé du bouton "Recharger tous les tableaux".
'---------------------------------------------------------------------------------------
Public Sub GetReloadAllEnabled(control As IRibbonControl, ByRef enabled As Variant)
    Const PROC_NAME As String = "GetReloadAllEnabled"
    On Error GoTo ErrorHandler
    
    ' Activer le bouton s'il y a au moins un tableau géré par l'addin dans le classeur.
    enabled = (TableManager.CountManagedTables(ThisWorkbook) > 0)

Exit Sub
ErrorHandler:
    enabled = False ' En cas d'erreur, le bouton est désactivé
    HandleError MODULE_NAME, PROC_NAME
End Sub

' --- PRIVATE HELPERS ---

'---------------------------------------------------------------------------------------
' Compte le nombre de tableaux gérés par l'addin dans un classeur.
' Un tableau est "géré" s'il a le bon préfixe ET un commentaire non vide sur sa première cellule.
'---------------------------------------------------------------------------------------
Private Function CountManagedTables(ByVal wb As Workbook) As Long
    Const PROC_NAME As String = "CountManagedTables"
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim count As Long
    Dim hasComment As Boolean
    Dim commentText As String
    count = 0

    Log "dataloader", "--- Début de la vérification des tableaux managés ---", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            hasComment = False
            commentText = ""
            On Error Resume Next
            commentText = tbl.Range.Cells(1, 1).Comment.Text
            hasComment = (Len(commentText) > 0)
            On Error GoTo 0

            Log "dataloader", "Tableau: '" & tbl.Name & "' sur la feuille '" & ws.Name & "'. Préfixe OK: " & (tbl.Name Like TABLE_PREFIX & "*") & ". Commentaire OK: " & hasComment & ". Longueur du commentaire: " & Len(commentText), DEBUG_LEVEL, PROC_NAME, MODULE_NAME

            If tbl.Name Like TABLE_PREFIX & "*" And hasComment Then
                count = count + 1
            End If
        Next tbl
    Next ws

    Log "dataloader", "--- Fin de la vérification. Total des tableaux managés: " & count & " ---", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    CountManagedTables = count

Exit Function
ErrorHandler:
    CountManagedTables = 0
    Debug.Print "!!! ERREUR dans CountManagedTables: " & Err.Description
    HandleError MODULE_NAME, PROC_NAME
End Function

