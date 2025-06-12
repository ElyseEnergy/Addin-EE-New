Attribute VB_Name = "RibbonVisibility"
Option Explicit

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
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du chargement du ruban."
End Sub

' Tâches d'initialisation exécutées après le chargement du ruban
Public Sub DelayedInitialization()
    Const PROC_NAME As String = "DelayedInitialization"
    On Error GoTo ErrorHandler
    
    ' Initialiser le système de logging
    SYS_Logger.InitializeLogger
    
    SYS_Logger.Log "ribbon", "Ribbon_Load (Delayed) appelé", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    ' Initialiser les profils de démo
    AccessProfiles.InitializeDemoProfiles
    SYS_Logger.Log "ribbon", "gRibbon (Delayed) initialisé", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    ' Optionnel : rafraîchir le ruban si des états ont changé
    InvalidateRibbon
    
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de l'initialisation différée."
End Sub

' Callback pour le sélecteur de profil
Public Sub OnSelectDemoProfile(control As IRibbonControl)
    Const PROC_NAME As String = "OnSelectDemoProfile"
    On Error GoTo ErrorHandler
    
    Dim profileID As DemoProfile
    
    ' Sélectionner le profil en fonction de l'ID du bouton
    Select Case control.id
        Case "btnEngineerBasic": profileID = AccessProfiles.Engineer_Basic
        Case "btnProjectManager": profileID = AccessProfiles.Project_Manager
        Case "btnFinanceController": profileID = AccessProfiles.Finance_Controller
        Case "btnTechnicalDirector": profileID = AccessProfiles.Technical_Director
        Case "btnMultiProjectLead": profileID = AccessProfiles.Business_Analyst  ' Changed from Multi_Project_Lead
        Case "btnFullAdmin": profileID = AccessProfiles.Full_Admin
        Case Else
            Exit Sub ' Ne rien faire si le bouton n'est pas reconnu
    End Select
    
    ' Changer le profil actif
    AccessProfiles.SetCurrentProfile (profileID)
    
    ' Invalider le ruban pour mettre à jour la visibilité
    If Not gRibbon Is Nothing Then
        gRibbon.Invalidate
    End If
    
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "ribbon_error", "Erreur dans OnSelectDemoProfile pour " & control.id & ": " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour l'affichage du profil actuel
Public Sub GetCurrentProfileLabel(ByVal control As IRibbonControl, ByRef label As Variant)
    Const PROC_NAME As String = "GetCurrentProfileLabel"
    On Error GoTo ErrorHandler
    
    label = "Current Profile: " & AccessProfiles.GetCurrentProfileName()
    
    Exit Sub
ErrorHandler:
    label = "Error"
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetCurrentProfileLabel."
End Sub

' Callback pour la visibilité du menu Technologies
Public Sub GetTechnologiesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetTechnologiesVisibility"
    On Error GoTo ErrorHandler
    
    visible = AccessProfiles.HasAccess("Engineering")
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetTechnologiesVisibility."
End Sub

' Callback pour la visibilité du menu Utilities
Public Sub GetUtilitiesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetUtilitiesVisibility"
    On Error GoTo ErrorHandler
    
    visible = AccessProfiles.HasAccess("Engineering")
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetUtilitiesVisibility."
End Sub

' Callback pour la visibilité du menu Files
Public Sub GetServerFilesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetServerFilesVisibility"
    On Error GoTo ErrorHandler
    
    visible = AccessProfiles.HasAccess("Tools")
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetServerFilesVisibility."
End Sub

' Callback pour la visibilité du menu Finances
Public Sub GetFinancesVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetFinancesVisibility"
    On Error GoTo ErrorHandler
    
    visible = AccessProfiles.HasAccess("Finance")
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetFinancesVisibility."
End Sub

' Callback pour la visibilité des menus de projets
Public Function GetProjectMenuVisibility(projectMenu As String) As Boolean
    Const PROC_NAME As String = "GetProjectMenuVisibility"
    On Error GoTo ErrorHandler
    
    ' Extrait le nom du projet du menu (par exemple "Echo" de "summaryEcho")
    Dim projectName As String
    If InStr(1, projectMenu, "GENERIC") > 0 Then
        GetProjectMenuVisibility = AccessProfiles.HasAccess("Engineering") ' Seuls les ingénieurs voient les génériques
    Else
        projectName = Replace(projectMenu, "summary", "")
        projectName = Replace(projectName, "planning", "")
        projectName = Replace(projectName, "devex", "")
        projectName = Replace(projectName, "capex", "")
        projectName = Replace(projectName, "opex", "")
        projectName = Replace(projectName, "tech", "")
        GetProjectMenuVisibility = AccessProfiles.HasAccess(projectName)
    End If
    
    Exit Function
ErrorHandler:
    GetProjectMenuVisibility = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetProjectMenuVisibility."
End Function

Public Sub GetSummarySheetsVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetSummarySheetsVisibility"
    On Error GoTo ErrorHandler
    
    visible = GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetSummarySheetsVisibility."
End Sub

Public Sub GetPlanningsVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetPlanningsVisibility"
    On Error GoTo ErrorHandler
    
    visible = GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetPlanningsVisibility."
End Sub

Public Sub GetDevexVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetDevexVisibility"
    On Error GoTo ErrorHandler
    
    visible = AccessProfiles.HasAccess("Finance") Or GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetDevexVisibility."
End Sub

Public Sub GetCapexVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetCapexVisibility"
    On Error GoTo ErrorHandler
    
    visible = AccessProfiles.HasAccess("Finance") Or GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetCapexVisibility."
End Sub

Public Sub GetOpexVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetOpexVisibility"
    On Error GoTo ErrorHandler
    
    visible = AccessProfiles.HasAccess("Finance") Or GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetOpexVisibility."
End Sub

Public Sub GetTechScenariosVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetTechScenariosVisibility"
    On Error GoTo ErrorHandler
    
    visible = AccessProfiles.HasAccess("Engineering") Or GetProjectMenuVisibility(control.id)
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetTechScenariosVisibility."
End Sub

' Callback pour la visibilité du bouton d'upload
Public Sub GetUploadButtonVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetUploadButtonVisibility"
    On Error GoTo ErrorHandler
    
    ' Visible uniquement si l'utilisateur a accès aux fichiers serveur
    visible = AccessProfiles.HasAccess("Files")
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetUploadButtonVisibility."
End Sub

' Fonction pour forcer le rafraîchissement du ruban
Public Sub InvalidateRibbon()
    Const PROC_NAME As String = "InvalidateRibbon"
    On Error GoTo ErrorHandler
    
    SYS_Logger.Log "ribbon", "InvalidateRibbon appelé", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    If gRibbon Is Nothing Then Exit Sub
    gRibbon.Invalidate
    SYS_Logger.Log "ribbon", "Ribbon invalidé", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur InvalidateRibbon."
End Sub

' Callback pour la visibilité du groupe d'administration
Public Sub GetAdminVisibility(control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetAdminVisibility"
    On Error GoTo ErrorHandler
    
    visible = AccessProfiles.HasAccess("Admin")
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetAdminVisibility."
End Sub

' Visibilité des boutons Recharger/Supprimer
Public Sub GetReloadButtonsVisible(ByVal control As IRibbonControl, ByRef visible As Variant)
    Const PROC_NAME As String = "GetReloadButtonsVisible"
    On Error GoTo ErrorHandler
    
    ' Visible uniquement si la sélection est dans un tableau géré
    visible = IsSelectionInManagedTable()
    
    Exit Sub
ErrorHandler:
    visible = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetReloadButtonsVisible."
End Sub

' État activé des boutons Recharger
Public Sub GetReloadCurrentEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)
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
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetReloadCurrentEnabled."
End Sub

'---------------------------------------------------------------------------------------
' Callback pour l'état activé du bouton "Recharger tous les tableaux".
'---------------------------------------------------------------------------------------
Public Sub GetReloadAllEnabled(control As IRibbonControl, ByRef enabled As Variant)
    Const PROC_NAME As String = "GetReloadAllEnabled"
    On Error GoTo ErrorHandler
    
    enabled = (TableManager.CountManagedTables(ThisWorkbook) > 0)
    
    Exit Sub
ErrorHandler:
    enabled = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetReloadAllEnabled."
End Sub

' Visibilité des boutons Recharger/Supprimer
Public Sub GetDeleteAllEnabled(ByVal control As IRibbonControl, ByRef enabled As Variant)
    Const PROC_NAME As String = "GetDeleteAllEnabled"
    Const MODULE_NAME As String = "RibbonVisibility"
    On Error GoTo ErrorHandler
    
    ' Pour l'instant, on active toujours si visible. Logique à affiner.
    enabled = True
    
    Exit Sub
ErrorHandler:
    enabled = False
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la vérification de l'activation du bouton de suppression."
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
    SYS_Logger.Log "ribbon_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur CountManagedTables."
End Function

Private Function HasAccess(ByVal permission As String) As Boolean
    Const PROC_NAME As String = "HasAccess"
    On Error GoTo ErrorHandler
    HasAccess = AccessProfiles.HasAccess(permission)
    Exit Function
ErrorHandler:
    HasAccess = False
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la vérification de l'accès pour '" & permission & "'."
End Function

