Attribute VB_Name = "DataLoaderManager"

' ==========================================
' Module DataLoaderManager
' ------------------------------------------
' Ce module gère le chargement et l'affichage des données pour toutes les catégories.
' Il orchestre la sélection, le filtrage, le collage et la protection des données importées via PowerQuery.
' Toutes les fonctions sont documentées individuellement ci-dessous.
' ==========================================
Option Explicit

' Variables privées du module
Private wsPQData As Worksheet

Public Enum DataLoadResult
    Success = 1
    Cancelled = 2
    Error = 3
End Enum

' Fonction principale de traitement des chargements de données pour une catégorie.
' Paramètres :
'   loadInfo (DataLoadInfo) : Informations de chargement
'   IsReload (Boolean, optionnel) : Vrai si c'est un rechargement non-interactif
'   TargetTableName (String, optionnel) : Nom du tableau à utiliser lors du rechargement
' Retour :
'   DataLoadResult (Succès, Annulé, Erreur)
Public Function ProcessDataLoad(loadInfo As DataLoadInfo, Optional IsReload As Boolean = False, Optional TargetTableName As String = "") As DataLoadResult
    On Error GoTo ErrorHandler
    Diagnostics.LogTime "Début de ProcessDataLoad pour la catégorie: " & loadInfo.Category.DisplayName
    Log "dataloader", "Début ProcessDataLoad | Catégorie: " & loadInfo.Category.DisplayName & " | URL: " & loadInfo.Category.URL, DEBUG_LEVEL, "ProcessDataLoad", "DataLoaderManager"
    ' Initialiser la feuille PQ_DATA de façon robuste
    Set wsPQData = SheetManager.GetOrCreatePQDataSheet()
    If wsPQData Is Nothing Then
        MsgBox "Erreur lors de l'initialisation de la feuille PQ_DATA", vbExclamation
        ProcessDataLoad = DataLoadResult.Error
        Exit Function
    End If
    ' S'assurer que la requête PowerQuery existe (réinjection si besoin)
    If Not PQQueryManager.EnsurePQQueryExists(loadInfo.Category) Then
        MsgBox "Erreur lors de la création de la requête PowerQuery", vbExclamation
        Log "dataloader", "ERREUR: EnsurePQQueryExists a échoué pour " & loadInfo.Category.DisplayName & " | URL: " & loadInfo.Category.URL, ERROR_LEVEL, "ProcessDataLoad", "DataLoaderManager"
        ProcessDataLoad = DataLoadResult.Error
        Exit Function
    End If
    ' S'assurer que la table PowerQuery existe dans la feuille (chargement si besoin)
    Dim lo As ListObject
    Set lo = Nothing
    On Error Resume Next
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    On Error GoTo 0
    If lo Is Nothing Then
        Log "dataloader", "Table PowerQuery manquante pour " & loadInfo.Category.DisplayName & ". Tentative de (re)chargement via LoadQueries.LoadQuery.", WARNING_LEVEL, "ProcessDataLoad", "DataLoaderManager"
        LoadQueries.LoadQuery loadInfo.Category.PowerQueryName, wsPQData, wsPQData.Cells(1, Utilities.GetLastColumn(wsPQData))
        Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
        If lo Is Nothing Then
            Log "dataloader", "ECHEC: Table PowerQuery toujours absente après LoadQuery. Diagnostics:", ERROR_LEVEL, "ProcessDataLoad", "DataLoaderManager"
            Log "dataloader", "  - QueryExists: " & PQQueryManager.QueryExists(loadInfo.Category.PowerQueryName), ERROR_LEVEL, "ProcessDataLoad", "DataLoaderManager"
            Log "dataloader", "  - Tables PQ_DATA: " & SheetManager.ListAllTableNames(wsPQData), ERROR_LEVEL, "ProcessDataLoad", "DataLoaderManager"
            MsgBox "Impossible de charger la table PowerQuery '" & loadInfo.Category.PowerQueryName & "' dans PQ_DATA. Voir logs pour diagnostic.", vbExclamation
            ProcessDataLoad = DataLoadResult.Error
            Exit Function
        End If
    End If
    Diagnostics.LogTime "Avant EnsurePQQueryExists"
    Log "dataloader", "Avant EnsurePQQueryExists | Catégorie: " & loadInfo.Category.DisplayName, DEBUG_LEVEL, "ProcessDataLoad", "DataLoaderManager"
    If Not PQQueryManager.EnsurePQQueryExists(loadInfo.Category) Then
        MsgBox "Erreur lors de la création de la requête PowerQuery", vbExclamation
        Log "dataloader", "ERREUR: EnsurePQQueryExists a échoué pour " & loadInfo.Category.DisplayName, ERROR_LEVEL, "ProcessDataLoad", "DataLoaderManager"
        ProcessDataLoad = DataLoadResult.Error
        Exit Function
    End If
    Diagnostics.LogTime "Après EnsurePQQueryExists"
    Log "dataloader", "Après EnsurePQQueryExists | Catégorie: " & loadInfo.Category.DisplayName, DEBUG_LEVEL, "ProcessDataLoad", "DataLoaderManager"
    ' --- RETOUR VISUEL PENDANT LE CHARGEMENT ---
    Application.Cursor = xlWait
    Application.StatusBar = "Téléchargement des données pour '" & loadInfo.Category.DisplayName & "' en cours..."
    DoEvents ' Forcer l'affichage du statut

    ' Forcer le chargement des données avant de continuer
    Dim lastCol As Long
    lastCol = Utilities.GetLastColumn(wsPQData)
    Diagnostics.LogTime "Avant LoadQuery (téléchargement des données)"
    Log "dataloader", "Avant LoadQuery | Catégorie: " & loadInfo.Category.DisplayName & " | PowerQuery: " & loadInfo.Category.PowerQueryName, DEBUG_LEVEL, "ProcessDataLoad", "DataLoaderManager"
    LoadQueries.LoadQuery loadInfo.Category.PowerQueryName, wsPQData, wsPQData.Cells(1, lastCol + 1)
    
    ' Vérifier que les données sont bien chargées
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    If lo Is Nothing Or lo.ListRows.Count = 0 Then
        ' Forcer un rafraîchissement de la requête
        Dim qry As WorkbookQuery
        Set qry = ThisWorkbook.Queries(loadInfo.Category.PowerQueryName)
        If Not qry Is Nothing Then
            qry.Refresh
            Application.Wait Now + TimeSerial(0, 0, 1) ' Attendre 1 seconde
        End If
    End If
    
    Diagnostics.LogTime "Après LoadQuery (téléchargement des données)"
    Log "dataloader", "Après LoadQuery | Catégorie: " & loadInfo.Category.DisplayName, DEBUG_LEVEL, "ProcessDataLoad", "DataLoaderManager"

    ' Restaurer le curseur et la barre de statut
    Application.Cursor = xlDefault
    Application.StatusBar = False

    If Not IsReload Then
        ' --- SÉLECTION UTILISATEUR PAR INPUTBOX ---
        Diagnostics.LogTime "Avant sélection des valeurs (InputBox)"
        Set loadInfo.SelectedValues = DataInteraction.GetSelectedValues(loadInfo.Category)
        If loadInfo.SelectedValues Is Nothing Or loadInfo.SelectedValues.Count = 0 Then
            ProcessDataLoad = DataLoadResult.Cancelled
            Exit Function
        End If

        ' --- SÉLECTION DU MODE (NORMAL/TRANSPOSE) SANS PREVIEW ---
        Diagnostics.LogTime "Avant sélection du mode d'affichage (MsgBox)"
        Dim modeChoice As VbMsgBoxResult
        modeChoice = MsgBox("Coller les fiches en mode NORMAL (lignes) ?" & VbCrLf & "Cliquez sur Non pour TRANSPOSE (colonnes).", vbYesNoCancel + vbQuestion, "Mode de collage")
        If modeChoice = vbCancel Then
            ProcessDataLoad = DataLoadResult.Cancelled
            Exit Function
        ElseIf modeChoice = vbNo Then
            loadInfo.ModeTransposed = True
        Else
            loadInfo.ModeTransposed = False
        End If

        Diagnostics.LogTime "Avant sélection de la destination (InputBox)"
        Set loadInfo.FinalDestination = DataInteraction.GetDestination(loadInfo)
        If loadInfo.FinalDestination Is Nothing Then
            ProcessDataLoad = DataLoadResult.Cancelled
            Exit Function
        End If
    End If

    ' Coller les données avec la méthode optimisée
    Diagnostics.LogTime "Avant appel à PasteData (Optimisé)"
    If Not DataPasting.PasteData(loadInfo, TargetTableName) Then
        ProcessDataLoad = DataLoadResult.Error
        Exit Function
    End If
    Diagnostics.LogTime "Après appel à PasteData (Optimisé)"
    ' S'assurer que la destination est visible
    With loadInfo.FinalDestination
        .Parent.Activate
        .Select
        .Parent.Range(.Address).Select
        ActiveWindow.ScrollRow = .Row
        ActiveWindow.ScrollColumn = .Column
    End With
    ' Ne pas nettoyer la requête PowerQuery pour conserver les requêtes dans le classeur
    ProcessDataLoad = DataLoadResult.Success
    Exit Function
ErrorHandler:
    ProcessDataLoad = DataLoadResult.Error
End Function

' Fonction utilitaire pour garantir l'existence de la feuille PQ_DATA et la variable globale
Public Function GetOrCreatePQDataSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreatePQDataSheet = Worksheets("PQ_DATA")
    On Error GoTo 0
    If GetOrCreatePQDataSheet Is Nothing Then
        Utilities.InitializePQData
        Set GetOrCreatePQDataSheet = Worksheets("PQ_DATA")
    End If
    Set wsPQData = GetOrCreatePQDataSheet
End Function

' Nettoie la requête PowerQuery en supprimant son tableau associé et la requête elle-même.
' Paramètres :
'   queryName (String) : Nom de la requête à nettoyer
Public Sub CleanupPowerQuery(queryName As String)
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "CleanupPowerQuery"
    Const MODULE_NAME As String = "DataLoaderManager"
    ' 1. Supprimer la table si elle existe
    Dim lo As ListObject
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(queryName))
    If Not lo Is Nothing Then
        lo.Delete
    End If
    ' 2. Forcer le nettoyage du cache PowerQuery en supprimant la requête
    Dim wb As Workbook
    Set wb = ThisWorkbook
    On Error Resume Next
    wb.Queries(queryName).Delete
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    ' Note : On ignore les erreurs ici car c'est une fonction de nettoyage
    ' qui peut échouer si les éléments n'existent déjà plus
    Log "cleanup_error", "Erreur ignorée lors du nettoyage de la requête " & queryName & ": " & Err.Description, WARNING_LEVEL, PROC_NAME, MODULE_NAME
    Resume Next
End Sub

' Fonction générique pour traiter une catégorie par son nom.
' Paramètres :
'   CategoryName (String) : Nom de la catégorie
'   errorMessage (String, optionnel) : Message d'erreur personnalisé
' Retour :
'   DataLoadResult (Succès, Annulé, Erreur)
Public Function ProcessCategory(CategoryName As String, Optional errorMessage As String = "") As DataLoadResult
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "ProcessCategory"
    Const MODULE_NAME As String = "DataLoaderManager"
    Log "dataloader", "Début ProcessCategory | Catégorie demandée: " & CategoryName, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    If CategoriesCount = 0 Then InitCategories
    Dim loadInfo As DataLoadInfo
    loadInfo.Category = GetCategoryByCategoryName(CategoryName)
    Log "dataloader", "Catégorie trouvée: " & loadInfo.Category.DisplayName & " | URL: " & loadInfo.Category.URL, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    If loadInfo.Category.DisplayName = "" Then
        MsgBox "Catégorie '" & CategoryName & "' non trouvée", vbExclamation
        Log "dataloader", "ERREUR: Catégorie non trouvée: " & CategoryName, ERROR_LEVEL, PROC_NAME, MODULE_NAME
        ProcessCategory = Error ' Utilisation directe de l'énumération
        Exit Function
    End If
    
    loadInfo.PreviewRows = 3
    
    Dim result As DataLoadResult
    result = ProcessDataLoad(loadInfo)
    If result = Cancelled Then ' Utilisation directe de l'énumération
        ProcessCategory = Cancelled ' Utilisation directe de l'énumération
        Exit Function
    ElseIf result = Error Then ' Utilisation directe de l'énumération
        Log "dataloader", "ECHEC: ProcessDataLoad a échoué pour " & CategoryName, ERROR_LEVEL, "ProcessCategory", "DataLoaderManager"
        Log "dataloader", "  - QueryExists: " & PQQueryManager.QueryExists(loadInfo.Category.PowerQueryName), ERROR_LEVEL, "ProcessCategory", "DataLoaderManager"
        Log "dataloader", "  - Tables PQ_DATA: " & SheetManager.ListAllTableNames(wsPQData), ERROR_LEVEL, "ProcessCategory", "DataLoaderManager"
        If errorMessage <> "" Then
            MsgBox errorMessage, vbExclamation
        End If
        ProcessCategory = Error ' Utilisation directe de l'énumération
        Exit Function
    End If
      ProcessCategory = Success ' Utilisation directe de l'énumération
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du traitement de la catégorie " & CategoryName & ": " & Err.Description
    If errorMessage <> "" Then
        MsgBox errorMessage, vbExclamation
    End If
    ProcessCategory = Error ' Utilisation directe de l'énumération
End Function

' =========================================================================================
' NOUVELLES FONCTIONS DE MISE À JOUR
' =========================================================================================

' Met à jour la table de données EE_ actuellement sélectionnée par l'utilisateur.
Public Sub ReloadSelectedTable()
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "ReloadSelectedTable"
    Const MODULE_NAME As String = "UpdateManager"
    
    Dim targetTable As ListObject
    Set targetTable = Nothing
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Veuillez sélectionner une cellule dans un tableau à mettre à jour.", vbInformation, "Sélection invalide"
        Exit Sub
    End If
    
    On Error Resume Next
    Set targetTable = Selection.ListObject
    On Error GoTo ErrorHandler
    
    If targetTable Is Nothing Then
        MsgBox "Veuillez sélectionner une cellule dans un tableau à mettre à jour.", vbInformation, "Aucun tableau trouvé"
        Exit Sub
    End If
    
    ' Vérifier que c'est un tableau EE_ avec un commentaire
    If Left(targetTable.Name, 3) <> "EE_" Then
        MsgBox "Le tableau sélectionné n'est pas un tableau de données géré.", vbExclamation
        Exit Sub
    End If
    
    ' Vérifier le commentaire sur la première cellule du tableau
    Dim tableComment As Comment
    Set tableComment = targetTable.Range.Cells(1, 1).Comment
    If tableComment Is Nothing Then
        MsgBox "Le tableau ne contient pas de métadonnées de rechargement.", vbExclamation
        Exit Sub
    End If
    
    Dim loadInfo As DataLoadInfo
    Set loadInfo = TableMetadata.DeserializeLoadInfo(tableComment.Text)
    
    If loadInfo.Category.CategoryName = "" Then
        MsgBox "Impossible de lire les métadonnées du tableau. Le rechargement a échoué.", vbExclamation, "Erreur de métadonnées"
        Exit Sub
    End If
    
    ' Préparer les informations pour un rechargement non-interactif
    loadInfo.FinalDestination = targetTable.Range.Cells(1, 1)
    
    Dim tableName As String
    tableName = targetTable.Name
    
    Dim ws As Worksheet
    Set ws = targetTable.Parent
    
    Application.ScreenUpdating = False
    
    ' Supprimer l'ancien tableau
    ws.Unprotect
    targetTable.Delete
    
    ' Appeler le processus de chargement en mode rechargement
    Dim result As DataLoadResult
    result = ProcessDataLoad(loadInfo, IsReload:=True, TargetTableName:=tableName)
    
    ' La protection est maintenant gérée à l'intérieur de PasteData

    Application.ScreenUpdating = True

    If result = Success Then
        MsgBox "Le tableau '" & tableName & "' a été mis à jour avec succès.", vbInformation, "Mise à jour réussie"
    Else
        MsgBox "La mise à jour du tableau '" & tableName & "' a échoué.", vbExclamation, "Échec de la mise à jour"
    End If
    
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la mise à jour du tableau."
End Sub

' Met à jour tous les tableaux de données EE_ dans le classeur actif.
Public Sub ReloadAllTables()
    Const PROC_NAME As String = "ReloadAllTables"
    Const MODULE_NAME As String = "DataLoaderManager"
    On Error GoTo ErrorHandler
    
    ' S'assurer que les catégories sont initialisées avant de lister les tableaux.
    If CategoriesCount = 0 Then InitCategories

    Application.ScreenUpdating = False
    
    Dim managedTables As Collection
    Set managedTables = TableManager.CollectManagedTables(ThisWorkbook)
    
    If managedTables.Count = 0 Then
        MsgBox "Aucun tableau géré par l'addin n'a été trouvé dans ce classeur.", vbInformation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Dim updatedCount As Long, failedCount As Long
    updatedCount = 0
    failedCount = 0
    
    Dim item As Variant
    Dim targetTable As ListObject
    
    For Each item In managedTables
        Set targetTable = Nothing ' Réinitialiser pour chaque itération
        On Error Resume Next
        Set targetTable = ThisWorkbook.Worksheets(item("SheetName")).ListObjects(item("Name"))
        On Error GoTo ErrorHandler
        
        If Not targetTable Is Nothing Then
            Dim updateResult As DataLoadResult
            updateResult = ReloadTable(targetTable) ' Utiliser la nouvelle fonction de rechargement
            If updateResult = Success Then
                updatedCount = updatedCount + 1
            Else
            failedCount = failedCount + 1
            End If
        Else
                failedCount = failedCount + 1
            Log "dataloader", "ERREUR: Tableau '" & item("Name") & "' non trouvé sur la feuille '" & item("SheetName") & "' lors de ReloadAllTables.", ERROR_LEVEL, PROC_NAME, MODULE_NAME
            End If
    Next item
    
    Application.ScreenUpdating = True
    
    Dim finalMsg As String
    finalMsg = updatedCount & " tableau(x) mis à jour avec succès."
    If failedCount > 0 Then
        finalMsg = finalMsg & vbCrLf & failedCount & " mise(s) à jour en échec."
    End If
    MsgBox finalMsg, vbInformation, "Rapport de mise à jour"
    
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la mise à jour de tous les tableaux."
End Sub

' Recharge un tableau spécifique de manière non-interactive.
' C'est la fonction logique principale pour le rechargement.
' Retourne un DataLoadResult pour indiquer le succès ou l'échec.
Private Function ReloadTable(ByVal targetTable As ListObject) As DataLoadResult
    Const PROC_NAME As String = "ReloadTable"
    Const MODULE_NAME As String = "DataLoaderManager"
    On Error GoTo ErrorHandler
    
    ' S'assurer que les catégories sont initialisées avant de désérialiser.
    If CategoriesCount = 0 Then InitCategories

    ' Étape 1: Extraire les métadonnées
    Dim tableComment As Comment
    Set tableComment = targetTable.Range.Cells(1, 1).Comment
    If tableComment Is Nothing Then
        HandleError MODULE_NAME, PROC_NAME, "Le tableau '" & targetTable.Name & "' n'a pas de métadonnées."
        ReloadTable = Error
        Exit Function
    End If
    
    Dim loadInfo As DataLoadInfo
    Set loadInfo = TableMetadata.DeserializeLoadInfo(tableComment.Text)
    
    If loadInfo.Category.CategoryName = "" Then
        HandleError MODULE_NAME, PROC_NAME, "Impossible de lire les métadonnées du tableau '" & targetTable.Name & "'."
        ReloadTable = Error
        Exit Function
    End If
    
    ' Étape 2: Préparer les informations pour un rechargement non-interactif
    loadInfo.FinalDestination = targetTable.Range.Cells(1, 1)

    ' Étape 3: Appeler le processus de chargement en mode rechargement non-destructif
                Dim result As DataLoadResult
    result = ProcessDataLoad(loadInfo, IsReload:=True, TargetTableName:=targetTable.Name)
    
    ReloadTable = result ' Retourner le résultat de ProcessDataLoad
    
    Exit Function
ErrorHandler:
    ReloadTable = Error
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du rechargement du tableau " & targetTable.Name
End Function

' Recharge le tableau actuellement sélectionné.
' Cette fonction est l'implémentation interne appelée par le callback du ruban.
Private Function ReloadCurrentTable_Internal() As DataLoadResult
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "ReloadCurrentTable_Internal"
    
    ' Vérifier qu'une cellule est sélectionnée
    If Selection Is Nothing Then
        MsgBox "Veuillez sélectionner une cellule dans le tableau à recharger.", vbExclamation
        ReloadCurrentTable_Internal = DataLoadResult.Cancelled
        Exit Function
                End If
    
    ' Vérifier que la sélection est dans un tableau
    Dim lo As ListObject
    On Error Resume Next
    Set lo = Selection.ListObject
    On Error GoTo ErrorHandler
    
    If lo Is Nothing Then
        MsgBox "Veuillez sélectionner une cellule dans un tableau.", vbExclamation
        ReloadCurrentTable_Internal = DataLoadResult.Cancelled
        Exit Function
            End If
    
    ' Vérifier que c'est un tableau géré par l'addin
    If Left(lo.Name, 3) <> "EE_" Then
        MsgBox "Ce tableau n'est pas géré par l'addin.", vbExclamation
        ReloadCurrentTable_Internal = DataLoadResult.Cancelled
        Exit Function
        End If
    
    ' Récupérer les métadonnées du tableau
    Dim comment As Comment
    Set comment = lo.Range.Cells(1, 1).Comment
    If comment Is Nothing Then
        MsgBox "Ce tableau n'a pas de métadonnées.", vbExclamation
        ReloadCurrentTable_Internal = DataLoadResult.Cancelled
        Exit Function
    End If
    
    ' Désérialiser les métadonnées
    Dim loadInfo As DataLoadInfo
    Set loadInfo = TableMetadata.DeserializeLoadInfo(comment.Text)
    If loadInfo Is Nothing Then
        MsgBox "Les métadonnées du tableau sont corrompues.", vbExclamation
        ReloadCurrentTable_Internal = DataLoadResult.Error
        Exit Function
    End If
    
    ' Recharger le tableau
    ReloadCurrentTable_Internal = ProcessDataLoad(loadInfo, True, lo.Name)
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du rechargement du tableau"
    ReloadCurrentTable_Internal = DataLoadResult.Error
End Function

' Point d'entrée pour le callback du ruban.
Public Sub ReloadCurrentTable_Click(control As IRibbonControl)
    ReloadCurrentTable_Internal
End Sub

' Supprime le tableau géré actuellement sélectionné.
Public Sub DeleteCurrentTable()
    Const PROC_NAME As String = "DeleteCurrentTable"
    Const MODULE_NAME As String = "DataLoaderManager"
    On Error GoTo ErrorHandler
    
    Dim targetTable As ListObject
    
    ' Vérifier la sélection
    If TypeName(Selection) <> "Range" Then
        MsgBox "Veuillez sélectionner une cellule dans un tableau à supprimer.", vbInformation
        Exit Sub
    End If
    
    On Error Resume Next
    Set targetTable = Selection.ListObject
    On Error GoTo 0 ' Rétablir la gestion d'erreur normale
    
    If targetTable Is Nothing Then
        MsgBox "La cellule sélectionnée n'appartient à aucun tableau.", vbInformation
        Exit Sub
    End If
    
    ' Vérifier si c'est un tableau géré
    Dim isManaged As Boolean
    isManaged = (Left(targetTable.Name, 3) = "EE_") And (Not targetTable.Range.Cells(1, 1).Comment Is Nothing)
    
    If Not isManaged Then
        MsgBox "Le tableau '" & targetTable.Name & "' n'est pas géré par l'addin.", vbInformation
        Exit Sub
    End If
    
    ' Demander confirmation
    If MsgBox("Êtes-vous sûr de vouloir supprimer définitivement le tableau '" & targetTable.Name & "' ?" & vbCrLf & _
              "Cette action est irréversible.", vbQuestion + vbYesNo, "Confirmation de suppression") = vbNo Then
        Exit Sub
    End If
    
    ' Supprimer le tableau
    Application.EnableEvents = False
    targetTable.Parent.Unprotect
    targetTable.Delete
    targetTable.Parent.Protect UserInterfaceOnly:=True
    Application.EnableEvents = True
    
    MsgBox "Le tableau a été supprimé.", vbInformation
    
    Exit Sub
ErrorHandler:
    If Not targetTable Is Nothing Then
        On Error Resume Next
        targetTable.Parent.Protect UserInterfaceOnly:=True
        On Error GoTo 0
    End If
    Application.EnableEvents = True
    HandleError MODULE_NAME, PROC_NAME
End Sub

' Callback pour le bouton de suppression du ruban.
Public Sub DeleteCurrentTable_Click(ByVal control As IRibbonControl)
    DeleteCurrentTable
End Sub

' ==========================================
' Fonctions de sérialisation des métadonnées
' ==========================================

' Sérialise les informations de chargement en une chaîne de caractères pour le stockage.
Private Function SerializeLoadInfo(loadInfo As DataLoadInfo) As String
    On Error GoTo ErrorHandler
    Dim parts As Collection
    Set parts = New Collection
    
    parts.Add "CategoryName" & META_KEYVAL_DELIM & loadInfo.Category.CategoryName
    
    Dim sVals As String
    If Not loadInfo.SelectedValues Is Nothing Then
        If loadInfo.SelectedValues.Count > 0 Then
            Dim arrVals() As String
            ReDim arrVals(1 To loadInfo.SelectedValues.Count)
            Dim i As Long: i = 1
            Dim v As Variant
            For Each v In loadInfo.SelectedValues
                arrVals(i) = CStr(v)
                i = i + 1
            Next v
            sVals = Join(arrVals, ",")
        End If
    End If
    parts.Add "SelectedValues" & META_KEYVAL_DELIM & sVals
    
    parts.Add "ModeTransposed" & META_KEYVAL_DELIM & CStr(loadInfo.ModeTransposed)
    
    Dim tempArray() As String
    ReDim tempArray(1 To parts.Count)
    Dim j As Long
    For j = 1 To parts.Count
        tempArray(j) = parts(j)
    Next j

    SerializeLoadInfo = Join(tempArray, META_DELIM)
    Exit Function
    
ErrorHandler:
    HandleError "DataLoaderManager", "SerializeLoadInfo", "Erreur de sérialisation"
    SerializeLoadInfo = ""
End Function

' Désérialise une chaîne de caractères en un objet DataLoadInfo.
Private Sub DeserializeLoadInfo(ByVal metadata As String, ByRef outLoadInfo As DataLoadInfo)
    On Error GoTo ErrorHandler
    
    ' S'assurer que les catégories sont initialisées avant de chercher dedans.
    If CategoriesCount = 0 Then InitCategories
    
    Set outLoadInfo.SelectedValues = New Collection

    Dim parts() As String
    parts = Split(metadata, META_DELIM)
    
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim pair() As String
        pair = Split(parts(i), META_KEYVAL_DELIM)
        
        If UBound(pair) >= 1 Then
            Dim key As String: key = pair(0)
            Dim value As String: value = pair(1)
            
            Select Case key
                Case "CategoryName"
                    outLoadInfo.Category = GetCategoryByCategoryName(value)
                Case "SelectedValues"
                    If value <> "" Then
                        Dim vals() As String
                        vals = Split(value, ",")
                        Dim v As Variant
                        For Each v In vals
                            outLoadInfo.SelectedValues.Add v
                        Next v
                    End If
                Case "ModeTransposed"
                    outLoadInfo.ModeTransposed = (value = "True")
            End Select
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    HandleError "DataLoaderManager", "DeserializeLoadInfo", "Erreur de désérialisation"
    ' outLoadInfo sera partiellement rempli mais la procédure va se terminer
End Sub

' Obtient la plage source pour le mode normal (lignes)
Private Function GetNormalSourceRange(ByVal sourceTable As ListObject, ByVal selectedValues As Collection) As Range
    Const PROC_NAME As String = "GetNormalSourceRange"
    Const MODULE_NAME As String = "DataLoaderManager"
    On Error GoTo ErrorHandler
    
    ' Créer une collection pour stocker les lignes à inclure
    Dim selectedRows As Collection
    Set selectedRows = New Collection
    
    ' Ajouter la ligne d'en-tête
    selectedRows.Add 1
    
    ' Trouver les lignes correspondant aux valeurs sélectionnées
    Dim v As Variant
    Dim rowIndex As Long
    For Each v In selectedValues
        For rowIndex = 1 To sourceTable.DataBodyRange.Rows.Count
            If CStr(sourceTable.DataBodyRange.Cells(rowIndex, 1).Value) = CStr(v) Then
                selectedRows.Add rowIndex + 1 ' +1 car la ligne d'en-tête est 1
                Exit For
    End If
        Next rowIndex
    Next v
    
    ' Créer un tableau pour stocker les adresses des lignes
    Dim rowAddresses() As String
    ReDim rowAddresses(1 To selectedRows.Count)
    
    ' Construire les adresses des lignes
    Dim i As Long
    For i = 1 To selectedRows.Count
        rowAddresses(i) = sourceTable.Range.Rows(selectedRows(i)).Address
    Next i
    
    ' Créer la plage union
    Set GetNormalSourceRange = Application.Union(sourceTable.Range.Worksheet.Range(Join(rowAddresses, ",")))
    
    Exit Function
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la création de la plage source en mode normal"
End Function

' Obtient la plage cible pour le mode normal (lignes)
Private Function GetNormalTargetRange(ByVal targetCell As Range, ByVal sourceRange As Range) As Range
    Const PROC_NAME As String = "GetNormalTargetRange"
    Const MODULE_NAME As String = "DataLoaderManager"
    On Error GoTo ErrorHandler
    
    ' La plage cible a la même taille que la plage source
    Set GetNormalTargetRange = targetCell.Resize(sourceRange.Rows.Count, sourceRange.Columns.Count)
    
    Exit Function
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la création de la plage cible en mode normal"
End Function

' Obtient la plage source pour le mode transposé (colonnes)
Private Function GetTransposedSourceRange(ByVal sourceTable As ListObject, ByVal selectedValues As Collection) As Range
    Const PROC_NAME As String = "GetTransposedSourceRange"
    Const MODULE_NAME As String = "DataLoaderManager"
    On Error GoTo ErrorHandler
    
    ' Créer une collection pour stocker les lignes à inclure
    Dim selectedRows As Collection
    Set selectedRows = New Collection
    
    ' Ajouter la ligne d'en-tête
    selectedRows.Add 1
    
    ' Trouver les lignes correspondant aux valeurs sélectionnées
    Dim v As Variant
    Dim rowIndex As Long
    For Each v In selectedValues
        For rowIndex = 1 To sourceTable.DataBodyRange.Rows.Count
            If CStr(sourceTable.DataBodyRange.Cells(rowIndex, 1).Value) = CStr(v) Then
                selectedRows.Add rowIndex + 1 ' +1 car la ligne d'en-tête est 1
                    Exit For
                End If
        Next rowIndex
    Next v
    
    ' Créer un tableau pour stocker les adresses des lignes
    Dim rowAddresses() As String
    ReDim rowAddresses(1 To selectedRows.Count)
    
    ' Construire les adresses des lignes
    Dim i As Long
    For i = 1 To selectedRows.Count
        rowAddresses(i) = sourceTable.Range.Rows(selectedRows(i)).Address
    Next i
    
    ' Créer la plage union
    Set GetTransposedSourceRange = Application.Union(sourceTable.Range.Worksheet.Range(Join(rowAddresses, ",")))
    
    Exit Function
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la création de la plage source en mode transposé"
End Function

' Obtient la plage cible pour le mode transposé (colonnes)
Private Function GetTransposedTargetRange(ByVal targetCell As Range, ByVal sourceRange As Range) As Range
    Const PROC_NAME As String = "GetTransposedTargetRange"
    Const MODULE_NAME As String = "DataLoaderManager"
    On Error GoTo ErrorHandler
    
    ' En mode transposé, on inverse les dimensions
    Set GetTransposedTargetRange = targetCell.Resize(sourceRange.Columns.Count, sourceRange.Rows.Count)
    
    Exit Function
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la création de la plage cible en mode transposé"
End Function




