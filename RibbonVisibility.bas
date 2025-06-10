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

' Callback pour la visibilité du bouton d'upload
Public Sub GetUploadButtonVisibility(control As IRibbonControl, ByRef visible As Variant)
    ' Visible uniquement si l'utilisateur a accès aux fichiers serveur
    visible = HasAccess("Files")
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
    enabled = (CountManagedTables(ThisWorkbook) > 0)

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

    Debug.Print "--- Début de la vérification des tableaux managés ---"

    For Each ws In wb.Worksheets
        For Each tbl In ws.ListObjects
            hasComment = False
            commentText = ""
            On Error Resume Next
            ' Essayer de lire le texte du commentaire sur la cellule en haut à gauche du tableau
            commentText = tbl.Range.Cells(1, 1).Comment.Text
            hasComment = (Len(commentText) > 0)
            On Error GoTo 0 ' Réinitialiser la gestion d'erreur

            ' Afficher les informations de diagnostic pour chaque tableau
            Debug.Print "Tableau: '" & tbl.Name & "' sur la feuille '" & ws.Name & "'. " & _
                        "Préfixe OK: " & (tbl.Name Like TABLE_PREFIX & "*") & ". " & _
                        "Commentaire OK: " & hasComment & ". " & _
                        "Longueur du commentaire: " & Len(commentText)

            If tbl.Name Like TABLE_PREFIX & "*" And hasComment Then
                count = count + 1
            End If
        Next tbl
    Next ws

    CountManagedTables = count
    Debug.Print "--- Fin de la vérification. Total des tableaux managés: " & count & " ---"

Exit Function
ErrorHandler:
    CountManagedTables = 0
    Debug.Print "!!! ERREUR dans CountManagedTables: " & Err.Description
    HandleError MODULE_NAME, PROC_NAME
End Function

