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
Private m_tableComment As Comment ' Variable partagée pour la gestion des commentaires

' Constantes pour la sérialisation
Private Const META_DELIM As String = "||"
Private Const META_KEYVAL_DELIM As String = "::"

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
    Set wsPQData = GetOrCreatePQDataSheet()
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
            Log "dataloader", "  - Tables PQ_DATA: " & ListAllTableNames(wsPQData), ERROR_LEVEL, "ProcessDataLoad", "DataLoaderManager"
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
    Dim lastCol As Long
    lastCol = Utilities.GetLastColumn(wsPQData)
    Diagnostics.LogTime "Avant LoadQuery (téléchargement des données)"
    Log "dataloader", "Avant LoadQuery | Catégorie: " & loadInfo.Category.DisplayName & " | PowerQuery: " & loadInfo.Category.PowerQueryName, DEBUG_LEVEL, "ProcessDataLoad", "DataLoaderManager"
    LoadQueries.LoadQuery loadInfo.Category.PowerQueryName, wsPQData, wsPQData.Cells(1, lastCol + 1)
    Diagnostics.LogTime "Après LoadQuery (téléchargement des données)"
    Log "dataloader", "Après LoadQuery | Catégorie: " & loadInfo.Category.DisplayName, DEBUG_LEVEL, "ProcessDataLoad", "DataLoaderManager"
    ' Restaurer le curseur et la barre de statut
    Application.Cursor = xlDefault
    Application.StatusBar = False

    If Not IsReload Then
        ' --- SÉLECTION UTILISATEUR PAR INPUTBOX ---
        Diagnostics.LogTime "Avant sélection des valeurs (InputBox)"
        Set loadInfo.SelectedValues = GetSelectedValues(loadInfo.Category)
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
        Set loadInfo.FinalDestination = GetDestination(loadInfo)
        If loadInfo.FinalDestination Is Nothing Then
            ProcessDataLoad = DataLoadResult.Cancelled
            Exit Function
        End If
    End If


    ' Coller les données avec la méthode optimisée
    Diagnostics.LogTime "Avant appel à PasteData (Optimisé)"
    If Not PasteData(loadInfo, TargetTableName) Then
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

' Récupère les valeurs sélectionnées selon le niveau de filtrage
Private Function GetSelectedValues(Category As CategoryInfo) As Collection
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetSelectedValues"
    Const MODULE_NAME As String = "DataLoaderManager"
    
    Dim lo As ListObject
    Dim dict As Object
    Dim arrValues() As String
    Dim i As Long, j As Long
    Dim cell As Range
    Dim v As Variant
    
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(Category.PowerQueryName))
    ' S'assurer que la table existe, sinon la charger
    If lo Is Nothing Then
        LoadQueries.LoadQuery Category.PowerQueryName, wsPQData, wsPQData.Cells(1, wsPQData.Cells(1, wsPQData.Columns.Count).End(xlToLeft).Column + 1)
        Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(Category.PowerQueryName))
        If lo Is Nothing Then
            MsgBox "Impossible de charger la table PowerQuery '" & Category.PowerQueryName & "'", vbExclamation
            Set GetSelectedValues = Nothing
            Exit Function
        End If
    End If

    ' Si pas de filtrage, permettre à l'utilisateur de choisir directement dans la liste complète
    If Category.FilterLevel = "Pas de filtrage" Then
        On Error Resume Next ' Pour gérer l'annulation de l'InputBox
        
        ' Créer un tableau avec toutes les fiches disponibles
        Dim displayArray() As String
        ReDim displayArray(1 To lo.DataBodyRange.Rows.Count)
        For i = 1 To lo.DataBodyRange.Rows.Count
            ' Utiliser la colonne 2 (nom) comme affichage
            displayArray(i) = lo.DataBodyRange.Rows(i).Columns(2).Value
        Next i
        
        ' Présenter les valeurs à l'utilisateur
        Set GetSelectedValues = LoadQueries.ChooseMultipleValuesFromArrayWithAll(displayArray, _
            "Choisissez une ou plusieurs fiches à charger (ex: 1,3,5 ou *) :")
            
        If Err.Number <> 0 Then
            Set GetSelectedValues = Nothing
            Exit Function
        End If
        On Error GoTo 0
        
        ' Gérer la sélection initiale
        Dim selectedIndices As Collection
        Set selectedIndices = GetSelectedValues
        
        ' Si l'utilisateur a annulé ou n'a rien sélectionné
        If selectedIndices Is Nothing Then
            Set GetSelectedValues = Nothing
            Exit Function
        End If
        
        ' Convertir les valeurs en IDs
        Set GetSelectedValues = New Collection
        For Each v In selectedIndices
            ' v est la valeur affichée, on doit retrouver la ligne correspondante
            For i = 1 To lo.DataBodyRange.Rows.Count
                If lo.DataBodyRange.Rows(i).Columns(2).Value = v Then
                    GetSelectedValues.Add lo.DataBodyRange.Rows(i).Columns(1).Value
                    Exit For
                End If
            Next i
        Next v
        
        ' Vérifier si des IDs ont été ajoutés
        If GetSelectedValues.Count = 0 Then
            MsgBox "Aucune fiche sélectionnée. Opération annulée.", vbExclamation
            Set GetSelectedValues = Nothing
            Exit Function
        End If
    Else
        ' Créer un dictionnaire pour stocker les valeurs uniques
        Set dict = CreateObject("Scripting.Dictionary")

        ' Extraire les valeurs uniques
        For Each cell In lo.ListColumns(Category.FilterLevel).DataBodyRange
            If Not dict.Exists(cell.Value) Then
                dict.Add cell.Value, 1
            End If
        Next cell

        ' Convertir le dictionnaire en tableau et trier
        ReDim arrValues(1 To dict.Count)
        i = 1
        For Each v In dict.Keys
            arrValues(i) = v
            i = i + 1
        Next v

        ' Trier le tableau
        For i = 1 To UBound(arrValues) - 1
            For j = i + 1 To UBound(arrValues)
                If arrValues(i) > arrValues(j) Then
                    Dim temp As String
                    temp = arrValues(i)
                    arrValues(i) = arrValues(j)
                    arrValues(j) = temp
                End If
            Next j
        Next i

        Dim selectedPrimary As Collection
        On Error Resume Next
        Set selectedPrimary = LoadQueries.ChooseMultipleValuesFromArrayWithAll(arrValues, _
            "Choisissez une ou plusieurs " & Category.FilterLevel & " (ex: 1,3,5 ou *) :")
        Dim errorOccurred As Boolean
        errorOccurred = (Err.Number <> 0)
        On Error GoTo 0

        If errorOccurred Or selectedPrimary Is Nothing Then
            MsgBox "Opération annulée", vbInformation
            Set GetSelectedValues = Nothing
            Exit Function
        End If

        If selectedPrimary.Count = 0 Then
            MsgBox "Aucune valeur sélectionnée. Opération annulée.", vbExclamation
            Set GetSelectedValues = Nothing
            Exit Function
        End If

        If Category.SecondaryFilterLevel <> "" Then
            ' Deuxième étape de filtrage
            Set dict = CreateObject("Scripting.Dictionary")
            For i = 1 To lo.DataBodyRange.Rows.Count
                For Each v In selectedPrimary
                    If lo.DataBodyRange.Rows(i).Columns(lo.ListColumns(Category.FilterLevel).Index).Value = v Then
                        Dim secVal As String
                        secVal = lo.DataBodyRange.Rows(i).Columns(lo.ListColumns(Category.SecondaryFilterLevel).Index).Value
                        If Not dict.Exists(secVal) Then dict.Add secVal, 1
                        Exit For
                    End If
                Next v
            Next i

            ReDim arrValues(1 To dict.Count)
            i = 1
            For Each v In dict.Keys
                arrValues(i) = v
                i = i + 1
            Next v

            For i = 1 To UBound(arrValues) - 1
                For j = i + 1 To UBound(arrValues)
                    If arrValues(i) > arrValues(j) Then
                        temp = arrValues(i)
                        arrValues(i) = arrValues(j)
                        arrValues(j) = temp
                    End If
                Next j
            Next i

            Dim selectedSecondary As Collection
            On Error Resume Next
            Set selectedSecondary = LoadQueries.ChooseMultipleValuesFromArrayWithAll(arrValues, _
                "Choisissez une ou plusieurs " & Category.SecondaryFilterLevel & " (ex: 1,3,5 ou *) :")
            errorOccurred = (Err.Number <> 0)
            On Error GoTo 0

            If errorOccurred Or selectedSecondary Is Nothing Then
                MsgBox "Opération annulée", vbInformation
                Set GetSelectedValues = Nothing
                Exit Function
            End If

            If selectedSecondary.Count = 0 Then
                MsgBox "Aucune valeur sélectionnée. Opération annulée.", vbExclamation
                Set GetSelectedValues = Nothing
                Exit Function
            End If

            Set GetSelectedValues = New Collection
            For i = 1 To lo.DataBodyRange.Rows.Count
                Dim matchPrimary As Boolean
                matchPrimary = False
                For Each v In selectedPrimary
                    If lo.DataBodyRange.Rows(i).Columns(lo.ListColumns(Category.FilterLevel).Index).Value = v Then
                        matchPrimary = True
                        Exit For
                    End If
                Next v
                If matchPrimary Then
                    For Each v In selectedSecondary
                        If lo.DataBodyRange.Rows(i).Columns(lo.ListColumns(Category.SecondaryFilterLevel).Index).Value = v Then
                            GetSelectedValues.Add lo.DataBodyRange.Rows(i).Columns(1).Value
                            Exit For
                        End If
                    Next v
                End If
            Next i

            If GetSelectedValues.Count = 0 Then
                MsgBox "Aucune fiche sélectionnée. Opération annulée.", vbExclamation
                Set GetSelectedValues = Nothing
                Exit Function
            End If
        Else
            Set GetSelectedValues = selectedPrimary
        End If
    End If
    Exit Function
    
ErrorHandler:
    If Err.Number = 424 Then  ' "L'objet est requis" - typiquement quand l'utilisateur annule une InputBox
        Set GetSelectedValues = Nothing
    Else
        Dim errorMsg As String
        errorMsg = "Erreur lors de la sélection des valeurs pour la catégorie " & Category.CategoryName & ": " & Err.Description
        HandleError MODULE_NAME, PROC_NAME, errorMsg
        Set GetSelectedValues = Nothing
    End If
    Exit Function
End Function

' Gère la sélection de la destination de collage des données.
' Paramètres :
'   loadInfo (DataLoadInfo) : Informations de chargement
' Retour :
'   Range (cellule de destination)
Private Function GetDestination(loadInfo As DataLoadInfo) As Range
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetDestination"
    Const MODULE_NAME As String = "DataLoaderManager"
    
    Dim lo As ListObject
    Dim nbRows As Long, nbCols As Long
    Dim okPlage As Boolean
    Dim i As Long, j As Long
    
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    
    ' Calculer la taille nécessaire
    If loadInfo.ModeTransposed Then
        nbRows = lo.ListColumns.Count
        nbCols = loadInfo.SelectedValues.Count + 1 ' +1 pour les en-têtes
    Else
        nbRows = loadInfo.SelectedValues.Count + 1 ' +1 pour les en-têtes
        nbCols = lo.ListColumns.Count
    End If
    
    Do
        Dim selectedRange As Range
        
        ' Activer Excel pour la sélection
        Application.Interactive = True
        Application.ScreenUpdating = True
        
        ' Demander à l'utilisateur de sélectionner une cellule
        On Error GoTo ErrorHandler
        Set selectedRange = Application.InputBox( _
            prompt:="Sélectionnez la cellule où charger les fiches (" & nbRows & " x " & nbCols & ")", _
            title:="Destination", _
            Type:=8)
            
        ' Vérifier si une plage valide a été sélectionnée
        If selectedRange Is Nothing Then
            MsgBox "Aucune cellule sélectionnée. Opération annulée.", vbInformation
            Set GetDestination = Nothing
            Exit Function
        End If
        
        ' S'assurer que c'est une seule cellule
        If selectedRange.Cells.Count > 1 Then
            MsgBox "Veuillez sélectionner une seule cellule.", vbExclamation
            GoTo ContinueLoop
        End If
        
        Set GetDestination = selectedRange
        GoTo CheckSpace
        
ErrorHandler:
        If Err.Number = 424 Then  ' Erreur "L'objet est requis" (annulation par l'utilisateur)
            MsgBox "Opération annulée", vbInformation
            Set GetDestination = Nothing
            Exit Function
        ElseIf Err.Number <> 0 Then
            HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la sélection de la destination: " & Err.Description
            Set GetDestination = Nothing
            Exit Function
        End If
        Resume Next
        
ContinueLoop:
        Err.Clear
        Resume Next
        
CheckSpace:
        
        okPlage = True
        For i = 0 To nbRows - 1
            For j = 0 To nbCols - 1
                If Not IsEmpty(GetDestination.Offset(i, j)) Then
                    okPlage = False
                    Exit For
                End If
            Next j
            If Not okPlage Then Exit For
        Next i
        
        If Not okPlage Then
            MsgBox "La plage sélectionnée n'est pas vide. Veuillez choisir un autre emplacement.", vbExclamation
        End If
    Loop Until okPlage
End Function

' Colle les données selon le mode choisi (normal ou transposé).
' Paramètres :
'   loadInfo (DataLoadInfo) : Informations de chargement
'   TargetTableName (String, optionnel) : Nom du tableau à utiliser si fourni
' Retour :
'   Boolean (True si succès)
Private Function PasteData(loadInfo As DataLoadInfo, Optional TargetTableName As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "PasteData"
    Const MODULE_NAME As String = "DataLoaderManager"
    
    ' Log des paramètres d'entrée
    Log "PasteData", "DÉBUT PASTEDATA =====================", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Log "PasteData", "Catégorie: " & loadInfo.Category.DisplayName, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    ' Convertir la Collection en array pour Join
    Dim selectedValuesArray() As String
    Dim idx As Long, v As Variant
    ReDim selectedValuesArray(1 To loadInfo.SelectedValues.Count)
    idx = 1
    For Each v In loadInfo.SelectedValues
        selectedValuesArray(idx) = CStr(v)
        idx = idx + 1
    Next v
    Log "PasteData", "Valeurs sélectionnées: " & Join(selectedValuesArray, ", "), DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    Dim lo As ListObject
    Dim sourceTable As ListObject ' Renamed from tblRange for clarity with PQ source table
    Dim i As Long, j As Long, k As Long ' k for iterating through selected values
    
    Dim destCol As Long, destRow As Long ' Renamed from currentCol, currentRow
    Dim sourceColIndex As Long
    Dim sourceRowIndex As Long

    Dim destSheet As Worksheet
    Set destSheet = loadInfo.FinalDestination.Parent

    ' Performance optimizations
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Log "PasteData", "Début collage. Mode transposé: " & loadInfo.ModeTransposed, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Log "PasteData", "Destination: Feuille=" & destSheet.Name & ", Cellule=" & loadInfo.FinalDestination.Address, DEBUG_LEVEL, PROC_NAME, MODULE_NAME

    ' Log de l'état de la feuille PQ_DATA
    Log "PasteData", "État PQ_DATA: Visible=" & wsPQData.Visible & ", Protected=" & wsPQData.ProtectContents, DEBUG_LEVEL, PROC_NAME, MODULE_NAME

    Set sourceTable = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    If sourceTable Is Nothing Then
        HandleError MODULE_NAME, PROC_NAME, "Table source '" & loadInfo.Category.PowerQueryName & "' non trouvée dans PQ_DATA."
        PasteData = False
        GoTo CleanupAndExit
    End If
    Log "PasteData", "Table source trouvée: " & sourceTable.Name & " (" & sourceTable.ListColumns.Count & " colonnes, " & sourceTable.ListRows.Count & " lignes)", DEBUG_LEVEL, PROC_NAME, MODULE_NAME

    ' Log des en-têtes de colonnes
    Dim headerNames As String
    headerNames = ""
    For i = 1 To sourceTable.ListColumns.Count
        headerNames = headerNames & sourceTable.HeaderRowRange.Cells(1, i).Value & ", "
    Next i
    Log "PasteData", "En-têtes des colonnes: " & headerNames, DEBUG_LEVEL, PROC_NAME, MODULE_NAME

    ' --- Pre-computation of column processing details ---
    On Error Resume Next ' Pour capturer les erreurs potentielles dans le pré-calcul
    Dim columnProcessingDetails As Object ' Scripting.Dictionary
    Set columnProcessingDetails = CreateObject("Scripting.Dictionary")
    Dim sourceHeaderName As String
    Dim ragicType As String
    Dim visibleSourceColIndices As Collection
    Set visibleSourceColIndices = New Collection

    Log "PasteData", "Début pré-calcul des détails de colonnes.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    For sourceColIndex = 1 To sourceTable.ListColumns.Count
        sourceHeaderName = CStr(sourceTable.HeaderRowRange.Cells(1, sourceColIndex).Value)
        Log "PasteData", "Analyse colonne " & sourceColIndex & ": '" & sourceHeaderName & "'", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        
        If Err.Number <> 0 Then
            Log "PasteData", "ERREUR lecture en-tête colonne " & sourceColIndex & ": " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
            Err.Clear
            GoTo NextColumn
        End If
        
        If RagicDictionary.IsFieldHidden(loadInfo.Category.SheetName, sourceHeaderName) Then
            Log "PasteData", "Colonne '" & sourceHeaderName & "' (index " & sourceColIndex & ") est cachée, elle sera ignorée.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        Else
            ragicType = RagicDictionary.GetFieldRagicType(loadInfo.Category.SheetName, sourceHeaderName)
            If Err.Number <> 0 Then
                Log "PasteData", "ERREUR GetFieldRagicType pour " & sourceHeaderName & ": " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
                Err.Clear
                GoTo NextColumn
            End If
            
            Dim colDetail As Object ' Scripting.Dictionary
            Set colDetail = CreateObject("Scripting.Dictionary")
            colDetail("ragicType") = ragicType
            colDetail("headerName") = sourceHeaderName
            colDetail("sourceIndex") = sourceColIndex ' Store original source index
            
            ' Convertir l'index en string pour la clé du dictionnaire
            Set columnProcessingDetails(CStr(visibleSourceColIndices.Count + 1)) = colDetail ' Keyed by visible column order
            visibleSourceColIndices.Add sourceColIndex ' Keep track of the original source index for visible columns
            Log "PasteData", "Colonne visible: '" & sourceHeaderName & "' (source idx " & sourceColIndex & ", visible idx " & visibleSourceColIndices.Count & "), Type: " & ragicType, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        End If
NextColumn:
    Next sourceColIndex
    On Error GoTo ErrorHandler ' Retour au handling d'erreur normal

    If visibleSourceColIndices.Count = 0 Then
        HandleError MODULE_NAME, PROC_NAME, "Aucune colonne visible à coller pour la catégorie '" & loadInfo.Category.DisplayName & "'."
        PasteData = False
        GoTo CleanupAndExit
    End If
    Log "PasteData", "Pré-calcul terminé. Nombre de colonnes visibles: " & visibleSourceColIndices.Count, DEBUG_LEVEL, PROC_NAME, MODULE_NAME

    ' --- Determine base font for section header size calculation ---
    Dim baseFontSize As Single
    Dim baseFontName As String
    With loadInfo.FinalDestination.Font
        baseFontSize = .Size
        baseFontName = .Name
    End With
    Log "PasteData", "Police de base: " & baseFontName & ", Taille: " & baseFontSize, DEBUG_LEVEL, PROC_NAME, MODULE_NAME

    ' --- Start pasting ---
    destRow = loadInfo.FinalDestination.Row
    destCol = loadInfo.FinalDestination.Column
    Log "PasteData", "Début collage à: Ligne=" & destRow & ", Colonne=" & destCol, DEBUG_LEVEL, PROC_NAME, MODULE_NAME

    Dim destCell As Range
    Dim cellInfo As FormattedCellOutput

    If loadInfo.ModeTransposed Then
        Log "PasteData", "Mode TRANSPOSE", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        ' Paste Headers (now as row headers)
        For i = 1 To visibleSourceColIndices.Count ' Iterate through VISIBLE columns
            Set colDetail = columnProcessingDetails.Item(CStr(i))
            sourceHeaderName = colDetail("headerName")
            ragicType = colDetail("ragicType")
            
            Set destCell = destSheet.Cells(destRow + i - 1, destCol)
            
            If ragicType = "Section" Then
                destCell.Value = sourceHeaderName
                destCell.NumberFormat = "@"
                With destCell.Font
                    .Bold = True
                    .Size = baseFontSize + 3
                    .Color = DataFormatter.SECTION_HEADER_DEFAULT_FONT_COLOR ' Use the constant from DataFormatter
                End With
            Else
                destCell.Value = sourceHeaderName
                destCell.NumberFormat = "@"
                ' Apply standard header font if different, or leave as is
            End If
        Next i

        ' Paste Data
        destCol = destCol + 1 ' Move to the first data column
        k = 0 ' Index for selected values
        For Each v In loadInfo.SelectedValues ' v is the ID of the selected record
            k = k + 1
            sourceRowIndex = 0
            Log "PasteData", "Recherche ID " & CStr(v) & " dans la table source...", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
            For j = 1 To sourceTable.DataBodyRange.Rows.Count ' Find the row in source table by ID
                If CStr(sourceTable.DataBodyRange.Cells(j, 1).Value) = CStr(v) Then
                    sourceRowIndex = j
                    Exit For
                End If
            Next j

            Dim originalSourceColIdx as Long
            If sourceRowIndex > 0 Then
                Log "PasteData", "ID " & CStr(v) & " trouvé à la ligne " & sourceRowIndex, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                For i = 1 To visibleSourceColIndices.Count ' Iterate through VISIBLE columns
                    Set colDetail = columnProcessingDetails.Item(CStr(i))
                    sourceHeaderName = colDetail("headerName") ' Field name for GetCellProcessingInfo
                    ragicType = colDetail("ragicType")
                    originalSourceColIdx = colDetail("sourceIndex") ' Get the original source column index

                    Set destCell = destSheet.Cells(destRow + i - 1, destCol + k - 1)
                    Log "PasteData", "Collage cellule: " & destCell.Address & ", Champ: " & sourceHeaderName & ", Type: " & ragicType, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                    
                    If ragicType = "Section" Then
                        destCell.Value = "" ' Blank for section data cells
                        destCell.NumberFormat = "@"
                        ' Ensure default background/no special font for data part of section
                        With destCell.Font
                            .Bold = False
                            .Size = baseFontSize
                            .Color = vbBlack ' Or whatever the default data font color is
                        End With
                        With destCell.Interior
                            .Pattern = xlNone
                        End With
                    Else
                        Dim originalValue As Variant
                        originalValue = sourceTable.DataBodyRange.Cells(sourceRowIndex, originalSourceColIdx).Value
                        Log "PasteData", "Valeur originale pour " & sourceHeaderName & ": " & CStr(originalValue), DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                        
                        On Error Resume Next
                        cellInfo = DataFormatter.GetCellProcessingInfo(originalValue, "", sourceHeaderName, loadInfo.Category.SheetName)
                        If Err.Number <> 0 Then
                            Log "PasteData", "ERREUR GetCellProcessingInfo pour " & sourceHeaderName & ": " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
                            Err.Clear
                        End If
                        On Error GoTo ErrorHandler
                        
                        destCell.Value = cellInfo.FinalValue
                        destCell.NumberFormat = cellInfo.NumberFormatString
                        Log "PasteData", "Valeur finale pour " & sourceHeaderName & ": " & CStr(cellInfo.FinalValue) & ", Format: " & cellInfo.NumberFormatString, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                    End If
                Next i
            Else
                 Log "PasteData", "ID non trouvé dans la table source: " & CStr(v), WARNING_LEVEL, PROC_NAME, MODULE_NAME
            End If
        Next v

    Else ' Mode NORMAL (not transposed)
        Log "PasteData", "Mode NORMAL", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        ' Paste Headers
        For i = 1 To visibleSourceColIndices.Count ' Iterate through VISIBLE columns
            Set colDetail = columnProcessingDetails.Item(CStr(i))
            sourceHeaderName = colDetail("headerName")
            ragicType = colDetail("ragicType")
            
            Set destCell = destSheet.Cells(destRow, destCol + i - 1)

            If ragicType = "Section" Then
                destCell.Value = sourceHeaderName
                destCell.NumberFormat = "@"
                With destCell.Font
                    .Bold = True
                    .Size = baseFontSize + 3 
                    .Color = DataFormatter.SECTION_HEADER_DEFAULT_FONT_COLOR ' Use the constant from DataFormatter
                End With
            Else
                destCell.Value = sourceHeaderName
                destCell.NumberFormat = "@"
                 ' Apply standard header font if different, or leave as is
            End If
        Next i

        ' Paste Data
        destRow = destRow + 1 ' Move to the first data row
        k = 0 ' Index for selected values
        For Each v In loadInfo.SelectedValues ' v is the ID of the selected record
            k = k + 1
            sourceRowIndex = 0
            Log "PasteData", "Recherche ID " & CStr(v) & " dans la table source...", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
            For j = 1 To sourceTable.DataBodyRange.Rows.Count ' Find the row in source table by ID
                If CStr(sourceTable.DataBodyRange.Cells(j, 1).Value) = CStr(v) Then
                    sourceRowIndex = j
                    Exit For
                End If
            Next j

            If sourceRowIndex > 0 Then
                Log "PasteData", "ID " & CStr(v) & " trouvé à la ligne " & sourceRowIndex, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                For i = 1 To visibleSourceColIndices.Count ' Iterate through VISIBLE columns
                    Set colDetail = columnProcessingDetails.Item(CStr(i))
                    sourceHeaderName = colDetail("headerName") ' Field name for GetCellProcessingInfo
                    ragicType = colDetail("ragicType")
                    originalSourceColIdx = colDetail("sourceIndex") ' Get the original source column index
                    
                    Set destCell = destSheet.Cells(destRow + k - 1, destCol + i - 1)

                    If ragicType = "Section" Then
                        destCell.Value = "" ' Blank for section data cells
                        destCell.NumberFormat = "@"
                        ' Ensure default background/no special font for data part of section
                        With destCell.Font
                            .Bold = False
                            .Size = baseFontSize
                            .Color = vbBlack ' Or whatever the default data font color is
                        End With
                        With destCell.Interior
                            .Pattern = xlNone
                        End With
                    Else
                        
                        originalValue = sourceTable.DataBodyRange.Cells(sourceRowIndex, originalSourceColIdx).Value
                        Log "PasteData", "Valeur originale pour " & sourceHeaderName & ": " & CStr(originalValue), DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                        
                        On Error Resume Next
                        cellInfo = DataFormatter.GetCellProcessingInfo(originalValue, "", sourceHeaderName, loadInfo.Category.SheetName)
                        If Err.Number <> 0 Then
                            Log "PasteData", "ERREUR GetCellProcessingInfo pour " & sourceHeaderName & ": " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
                            Err.Clear
                        End If
                        On Error GoTo ErrorHandler
                        
                        destCell.Value = cellInfo.FinalValue
                        destCell.NumberFormat = cellInfo.NumberFormatString
                        Log "PasteData", "Valeur finale pour " & sourceHeaderName & ": " & CStr(cellInfo.FinalValue) & ", Format: " & cellInfo.NumberFormatString, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                    End If
                Next i
            Else
                Log "PasteData", "ID non trouvé dans la table source: " & CStr(v), WARNING_LEVEL, PROC_NAME, MODULE_NAME
            End If
        Next v
    End If

    ' --- CRÉATION OU MISE À JOUR DU LISTOBJECT ---
    Dim targetRange As Range
    Dim nbRows As Long, nbCols As Long

    ' Recalculer la taille finale en fonction des colonnes visibles
    If loadInfo.ModeTransposed Then
        nbRows = visibleSourceColIndices.Count + 1 ' +1 for headers
        nbCols = loadInfo.SelectedValues.Count + 1 ' +1 for headers
    Else
        nbRows = loadInfo.SelectedValues.Count + 1 ' +1 for headers
        nbCols = visibleSourceColIndices.Count
    End If
    
    Set targetRange = destSheet.Range(loadInfo.FinalDestination, loadInfo.FinalDestination.Offset(nbRows - 1, nbCols - 1))

    Dim existingTable As ListObject
    Set existingTable = Nothing
    On Error Resume Next
    If TargetTableName <> "" Then
        Set existingTable = destSheet.ListObjects(TargetTableName)
    End If
    On Error GoTo ErrorHandler

    If Not existingTable Is Nothing Then
        ' --- MISE À JOUR NON-DESTRUCTIVE ---
        Set lo = existingTable
        ' 1. Vider les anciennes données et en-têtes sans supprimer le ListObject
        lo.DataBodyRange.ClearContents
        lo.HeaderRowRange.ClearContents
        ' 2. Redimensionner le tableau pour les nouvelles données
        lo.Resize targetRange
    Else
        ' --- CRÉATION CLASSIQUE ---
        Set lo = destSheet.ListObjects.Add(xlSrcRange, targetRange, , xlYes)
        If TargetTableName <> "" Then
            lo.Name = TargetTableName
        Else
            lo.Name = GetUniqueTableName(loadInfo.Category.CategoryName)
        End If
    End If
    
    ' Le reste du collage des données (qui se fait sur la feuille) est déjà fait avant.
    ' Ici on s'assure juste que le ListObject est bien défini.

    ' Appliquer le style de tableau par défaut
    lo.TableStyle = "TableStyleMedium2"
    
    ' Verrouiller le nouveau tableau
    lo.Range.Locked = True
    
    ' Protéger la feuille avec les paramètres standard
    destSheet.Protect UserInterfaceOnly:=True, _
        AllowFormattingCells:=True, _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, _
        AllowInsertingColumns:=True, _
        AllowInsertingRows:=True, _
        AllowInsertingHyperlinks:=True, _
        AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, _
        AllowSorting:=True, _
        AllowFiltering:=True, _
        AllowUsingPivotTables:=True

    PasteData = True

CleanupAndExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Log "PasteData", "Fin collage. Statut: " & PasteData, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du collage des données: " & Err.Description
    PasteData = False
    GoTo CleanupAndExit
End Function

' Protège uniquement les tableaux EE_ dans la feuille spécifiée.
' Paramètres :
'   ws (Worksheet) : Feuille à protéger
Private Sub ProtectSheetWithTable(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "ProtectSheetWithTable"
    Const MODULE_NAME As String = "DataLoaderManager"
    
    ws.Unprotect
    
    ' 1. Déverrouiller toutes les cellules
    ws.Cells.Locked = False
    
    ' 2. Verrouiller uniquement les cellules des tableaux EE_
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        If Left(tbl.Name, 3) = "EE_" Then
            tbl.Range.Locked = True
        End If
    Next tbl
    
    ' 3. Protéger la feuille avec les permissions standard    ws.Protect UserInterfaceOnly:=True, AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True, _
               AllowInsertingColumns:=True, AllowInsertingRows:=True, _
               AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
               AllowDeletingRows:=True, AllowSorting:=True, _
               AllowFiltering:=True, AllowUsingPivotTables:=True
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la protection de la feuille avec les tableaux"
End Sub

' Génère un nom unique pour un nouveau tableau en incrémentant l'indice.
' Paramètres :
'   CategoryName (String) : Nom de la catégorie
' Retour :
'   String (nom unique)
Private Function GetUniqueTableName(CategoryName As String) As String
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetUniqueTableName"
    Const MODULE_NAME As String = "DataLoaderManager"
    
    Dim baseName As String
    baseName = "EE_" & Utilities.SanitizeTableName(CategoryName)
    Dim maxIndex As Long
    maxIndex = 0
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim currentIndex As Long
    Dim tableName As String

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = baseName Then
                If maxIndex < 1 Then maxIndex = 1
            ElseIf Left(tbl.Name, Len(baseName) + 1) = baseName & "_" Then
                tableName = Mid(tbl.Name, Len(baseName) + 2)
                If IsNumeric(tableName) Then
                    currentIndex = CLng(tableName)
                    If currentIndex > maxIndex Then
                        maxIndex = currentIndex
                    End If
                End If
            End If
        Next tbl
    Next ws    
    GetUniqueTableName = baseName & "_" & (maxIndex + 1)
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la génération du nom unique pour le tableau de la catégorie " & CategoryName
    GetUniqueTableName = baseName & "_ERROR"
End Function

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
        Log "dataloader", "  - Tables PQ_DATA: " & ListAllTableNames(wsPQData), ERROR_LEVEL, "ProcessCategory", "DataLoaderManager"
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
    DeserializeLoadInfo tableComment.Text, loadInfo
    
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
    Set managedTables = CollectManagedTables(ThisWorkbook)
    
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
    DeserializeLoadInfo tableComment.Text, loadInfo
    
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

' Recharge les données du tableau actuellement sélectionné (point d'entrée de la callback)
Public Sub ReloadCurrentTable()
    Const PROC_NAME As String = "ReloadCurrentTable"
    Const MODULE_NAME As String = "DataLoaderManager"
    On Error GoTo ErrorHandler
    
    ' Vérifier qu'une cellule est sélectionnée
    If TypeName(Selection) <> "Range" Then
        MsgBox "Veuillez sélectionner une cellule dans un tableau à mettre à jour.", vbInformation
        Exit Sub
    End If
    
    ' Vérifier que la cellule est dans un tableau
    Dim targetTable As ListObject
    On Error Resume Next
    Set targetTable = Selection.ListObject
    On Error GoTo ErrorHandler
    
    If targetTable Is Nothing Then
        MsgBox "Veuillez sélectionner une cellule dans un tableau à mettre à jour.", vbInformation
        Exit Sub
    End If
    
    ' Vérifier que c'est un tableau géré par l'addin
    Dim hasComment As Boolean
    hasComment = False
    On Error Resume Next
    hasComment = (Len(targetTable.Range.Cells(1, 1).Comment.Text) > 0)
    On Error GoTo ErrorHandler

    If Left(targetTable.Name, 3) = "EE_" And hasComment Then
        ' Appeler la fonction logique de rechargement
        If ReloadTable(targetTable) = Success Then
            MsgBox "Le tableau a été mis à jour avec succès.", vbInformation
        Else
            MsgBox "La mise à jour du tableau a échoué. Consultez les logs pour plus de détails.", vbExclamation
        End If
    Else
        MsgBox "Ce tableau n'est pas géré par l'addin.", vbInformation
    End If
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la mise à jour du tableau"
End Sub

' Callback pour recharger le tableau courant
Public Sub ReloadCurrentTable_Click(control As IRibbonControl)
    ReloadCurrentTable
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




