Attribute VB_Name = "RagicDictionary"

Option Explicit

Public RagicFieldDict As Object
Public wsPQDict As Worksheet

' Constantes pour les noms et propriétés
Private Const BASE_NAME As String = "RagicDictionary"
Private Const RAGIC_PATH As String = "matching-matrix/6.csv"
Private Const PROP_LAST_REFRESH As String = "RagicDictLastRefresh"

'==================================================================================================
' CALLBACKS DU RUBAN
'==================================================================================================

' Callback pour le bouton du ruban pour forcer le rafraîchissement
Public Sub ProcessForceRefreshRagicDictionary(ByVal control As IRibbonControl)
    ForceRefreshRagicDictionary
End Sub

' Callback pour l'info-bulle (supertip) du bouton de rafraîchissement
Public Sub GetRagicDictSupertip(control As IRibbonControl, ByRef supertip)
    Dim lastUpdate As Date
    lastUpdate = GetLastRefreshDate() ' On réutilise la fonction existante
    
    Dim lastUpdateText As String
    If lastUpdate > 0 Then
        lastUpdateText = "Last update: " & Format(lastUpdate, "yyyy-mm-dd")
    Else
        lastUpdateText = "Never updated. Click to download."
    End If
    
    supertip = "Downloads the latest version of the data dictionary from Ragic." & vbCrLf & vbCrLf & lastUpdateText
End Sub

'==================================================================================================
' MÉTHODES PUBLIQUES
'==================================================================================================

' Force le rafraîchissement du dictionnaire depuis Ragic
Public Sub ForceRefreshRagicDictionary()
    Application.StatusBar = "Forçage du rafraîchissement du dictionnaire Ragic..."
    ' Réinitialiser la date dans les propriétés pour forcer le rechargement
    SetLastRefreshDate (0)
    ' Appeler la routine de chargement
    LoadRagicDictionary
    Application.StatusBar = False
    MsgBox "Le dictionnaire Ragic a été mis à jour.", vbInformation
End Sub

' Charge le dictionnaire Ragic, depuis le cache si possible
Public Sub LoadRagicDictionary()
    ' S'assurer que la liste des catégories est initialisée
    If CategoryManager.CategoriesCount = 0 Then
        CategoryManager.InitCategories
    End If
    
    Dim lastRefresh As Date
    lastRefresh = GetLastRefreshDate()

    ' On ne vérifie plus si l'objet RagicFieldDict existe, car il est toujours vide à l'ouverture.
    ' La logique se base sur l'existence de la table dans la feuille de cache et la date.

    Application.StatusBar = "Vérification du dictionnaire Ragic..."

    ' Définir les noms standardisés
    Dim pqName As String
    pqName = "PQ_" & Utilities.SanitizeTableName(BASE_NAME)
    Dim tableName As String
    tableName = "Table_" & pqName

    ' Créer ou récupérer la feuille PQ_DICT
    Set wsPQDict = GetOrCreatePQDictSheet()

    ' Vérifier si la table de cache existe déjà dans la feuille
    Dim tableExists As Boolean
    On Error Resume Next
    tableExists = (wsPQDict.ListObjects(tableName).Name <> "")
    On Error GoTo 0
    Log "load_dict", "Table " & tableName & " existe: " & tableExists, DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"
    
    ' Décider s'il faut rafraîchir depuis le réseau
    Dim needsRefresh As Boolean
    needsRefresh = Not tableExists Or (VBA.Date - lastRefresh >= 1)
    Log "load_dict", "Dernière MàJ: " & lastRefresh & ", Âge (heures): " & ((Now - lastRefresh) * 24) & "h", DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"
    Log "load_dict", "Rafraîchissement nécessaire: " & needsRefresh & " (table existe: " & tableExists & ")", DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"

    If needsRefresh Then
        Application.StatusBar = "Chargement du dictionnaire Ragic depuis le réseau..."
        Log "load_dict", "Rafraîchissement du dictionnaire Ragic depuis le réseau.", INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"

        ' Créer une catégorie pour le dictionnaire
        Dim dictCategory As CategoryInfo
        With dictCategory
            .CategoryName = BASE_NAME
            .DisplayName = BASE_NAME
            .URL = env.RAGIC_BASE_URL & RAGIC_PATH & env.RAGIC_API_PARAMS
            .PowerQueryName = pqName
            .SheetName = BASE_NAME
        End With

        Log "load_dict", "URL de requête dictionnaire : " & dictCategory.URL, INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"

        ' Créer ou mettre à jour la requête
        If QueryExists(pqName) Then
            On Error Resume Next
            ThisWorkbook.Queries(pqName).formula = GenerateDictionaryQuery(dictCategory.URL)
            If Err.Number <> 0 Then
                Log "ragic_dict_err", "Erreur MàJ requête " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
                Application.StatusBar = False
                Exit Sub
            End If
            On Error GoTo 0
        Else
            On Error Resume Next
            ThisWorkbook.Queries.Add pqName, GenerateDictionaryQuery(dictCategory.URL)
            If Err.Number <> 0 Then
                Log "ragic_dict_err", "Erreur ajout requête " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
                Application.StatusBar = False
                Exit Sub
            End If
            On Error GoTo 0
        End If

        ' Rafraîchir la requête (c'est l'étape lente)
        On Error Resume Next
        ThisWorkbook.Queries(pqName).Refresh
        If Err.Number <> 0 Then
            Log "ragic_dict_err", "Erreur Refresh requête " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
            Application.StatusBar = False
            Exit Sub
        End If
        On Error GoTo 0

        ' Charger les données dans la feuille PQ_DICT
        ' Si la table existe déjà, forcer un refresh de la requête associée
        If QueryExists(pqName) Then
            On Error Resume Next
            ThisWorkbook.Queries(pqName).Refresh
            If Err.Number <> 0 Then
                Log "ragic_dict_err", "Erreur Refresh requête existante " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
                Application.StatusBar = False
                Exit Sub
            End If
            On Error GoTo 0
        Else
            LoadQueries.LoadQuery pqName, wsPQDict, wsPQDict.Range("A1")
        End If

        ' Mettre à jour la date du rafraîchissement dans les propriétés du classeur
        SetLastRefreshDate VBA.Date
        
        ' Sauvegarder le classeur pour rendre la date de mise à jour persistante
        Log "load_dict", "Tentative de sauvegarde forcée du classeur pour persistance de la date...", INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"
        On Error Resume Next
        ' On marque explicitement le classeur comme ayant des modifications non enregistrées pour forcer la sauvegarde
        ThisWorkbook.Saved = False
        ThisWorkbook.Save
        If Err.Number <> 0 Then
            Log "ragic_dict_err", "ERREUR lors de la sauvegarde du classeur: " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
        Else
            Log "load_dict", "Classeur sauvegardé avec succès. La date de mise à jour est maintenant persistante.", INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"
        End If
        On Error GoTo 0
        
        ' Invalider le contrôle du ruban pour mettre à jour l'info-bulle
        If Not RibbonVisibility.gRibbon Is Nothing Then
            RibbonVisibility.gRibbon.InvalidateControl "btnForceRefreshRagic"
        End If
    Else
        Log "load_dict", "Chargement du dictionnaire Ragic depuis le cache local (feuille PQ_DICT).", INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"
    End If

    ' Initialiser et charger les données dans le dictionnaire VBA
    Application.StatusBar = "Finalisation du chargement du dictionnaire..."
    If RagicFieldDict Is Nothing Then
        Set RagicFieldDict = CreateObject("Scripting.Dictionary")
    Else
        RagicFieldDict.RemoveAll
    End If

    LoadDictionaryData tableName

    ' Log des 10 premières clés du dictionnaire pour debug
    Dim debugKeys As String, debugCount As Long
    debugKeys = ""
    For debugCount = 0 To Application.Min(9, RagicFieldDict.Count - 1)
        debugKeys = debugKeys & RagicFieldDict.Keys()(debugCount) & "; "
    Next debugCount
    Log "load_dict", "Premières clés du dictionnaire : " & debugKeys, DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"

    ' Forcer la visibilité de la feuille PQ_DICT
    If Not wsPQDict Is Nothing Then
        wsPQDict.Visible = xlSheetVisible
    End If

    ' Réinitialiser la barre de statut
    Application.StatusBar = False
End Sub

' Recherche la meilleure ligne pour un fieldName donné (et SheetName si doublons)
Public Function FindBestRowForField(lo As ListObject, SheetName As String, fieldName As String) As Long
    Log "FindBestRowForField", "Entrée: SheetName='" & SheetName & "', fieldName='" & fieldName & "'", DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary"
    Dim arr As Variant
    arr = lo.DataBodyRange.Value
    Dim colSheetName As Long, colFieldName As Long, i As Long
    For i = 1 To lo.ListColumns.Count
        Select Case lo.ListColumns(i).Name
            Case "SheetName": colSheetName = i
            Case "Field Name": colFieldName = i
        End Select
    Next i
    Dim matches() As Long, matchCount As Long
    matchCount = 0
    For i = 1 To UBound(arr, 1)
        If StrComp(Trim(arr(i, colFieldName)), Trim(fieldName), vbTextCompare) = 0 Then
            matchCount = matchCount + 1
            ReDim Preserve matches(1 To matchCount)
            matches(matchCount) = i
        End If
    Next i
    Log "FindBestRowForField", "Nb matches sur FieldName: " & matchCount, DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary"
    If matchCount = 0 Then Log "FindBestRowForField", "Aucun match trouvé.", DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = 0: Exit Function
    If matchCount = 1 Then Log "FindBestRowForField", "Un seul match, row=" & matches(1), DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = matches(1): Exit Function
    Dim j As Long
    For j = 1 To matchCount
        If StrComp(Trim(arr(matches(j), colSheetName)), Trim(SheetName), vbTextCompare) = 0 Then Log "FindBestRowForField", "Match exact SheetName à row=" & matches(j), DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = matches(j): Exit Function
    Next j
    For j = 1 To matchCount
        If InStr(1, arr(matches(j), colSheetName), SheetName, vbTextCompare) > 0 Then Log "FindBestRowForField", "Match contains SheetName à row=" & matches(j), DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = matches(j): Exit Function
    Next j
    For j = 1 To matchCount
        If Left(arr(matches(j), colSheetName), Len(SheetName)) = SheetName Then Log "FindBestRowForField", "Match startswith SheetName à row=" & matches(j), DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = matches(j): Exit Function
    Next j
    Log "FindBestRowForField", "Fallback premier match à row=" & matches(1), DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = matches(1)
End Function

' Retourne la valeur d'une colonne pour une ligne donnée
Public Function GetValueFromRow(lo As ListObject, arr As Variant, rowIndex As Long, colName As String) As Variant
    Log "GetValueFromRow", "rowIndex=" & rowIndex & ", colName='" & colName & "'", DEBUG_LEVEL, "GetValueFromRow", "RagicDictionary"
    Dim colIdx As Long, i As Long
    For i = 1 To lo.ListColumns.Count
        If lo.ListColumns(i).Name = colName Then colIdx = i: Exit For
    Next i
    If colIdx = 0 Or rowIndex = 0 Then
        Log "GetValueFromRow", "Colonne ou ligne non trouvée.", DEBUG_LEVEL, "GetValueFromRow", "RagicDictionary"
        GetValueFromRow = ""
    Else
        GetValueFromRow = arr(rowIndex, colIdx)
        Log "GetValueFromRow", "Valeur extraite: '" & GetValueFromRow & "'", DEBUG_LEVEL, "GetValueFromRow", "RagicDictionary"
    End If
End Function

' Fonction principale Hidden
Public Function IsFieldHidden(SheetName As String, fieldName As String) As Boolean
    Log "IsFieldHidden", "Entrée: SheetName='" & SheetName & "', fieldName='" & fieldName & "'", DEBUG_LEVEL, "IsFieldHidden", "RagicDictionary"
    IsFieldHidden = False
    If wsPQDict Is Nothing Then Set wsPQDict = GetOrCreatePQDictSheet()
    Dim lo As ListObject
    Dim tableName As String
    tableName = "Table_PQ_" & Utilities.SanitizeTableName(BASE_NAME)
    On Error Resume Next
    Set lo = wsPQDict.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Or lo.ListRows.Count = 0 Then Log "IsFieldHidden", "Table non trouvée ou vide.", DEBUG_LEVEL, "IsFieldHidden", "RagicDictionary": Exit Function
    Dim arr As Variant
    arr = lo.DataBodyRange.Value
    Dim rowIdx As Long
    rowIdx = FindBestRowForField(lo, SheetName, fieldName)
    If rowIdx = 0 Then Log "IsFieldHidden", "Aucune ligne trouvée.", DEBUG_LEVEL, "IsFieldHidden", "RagicDictionary": Exit Function
    Dim memoVal As String
    memoVal = GetValueFromRow(lo, arr, rowIdx, "Memo")
    If InStr(1, memoVal, "Hidden", vbTextCompare) > 0 Then
        Log "IsFieldHidden", "Champ HIDDEN détecté.", DEBUG_LEVEL, "IsFieldHidden", "RagicDictionary"
        IsFieldHidden = True
    Else
        Log "IsFieldHidden", "Champ NON hidden.", DEBUG_LEVEL, "IsFieldHidden", "RagicDictionary"
    End If
End Function

' Normalise le nom de la feuille pour la clé dictionnaire
Public Function NormalizeSheetName(SheetName As String) As String
    Dim i As Long
    Dim c As String
    Dim result As String
    result = ""
    For i = 1 To Len(SheetName)
        c = Mid(SheetName, i, 1)
        If (c >= "A" And c <= "Z") Or _
           (c >= "a" And c <= "z") Or _
           (c >= "0" And c <= "9") Then
            result = result & c
        End If
    Next i
    NormalizeSheetName = result
End Function

'==================================================================================================
' MÉTHODES PRIVÉES
'==================================================================================================

' Gère la persistance de la date de rafraîchissement via les propriétés du document
Private Function GetLastRefreshDate() As Date
    On Error Resume Next
    GetLastRefreshDate = ThisWorkbook.CustomDocumentProperties(PROP_LAST_REFRESH).Value
    If Err.Number <> 0 Then
        GetLastRefreshDate = 0 ' Force le rafraîchissement si la propriété n'existe pas
    End If
    On Error GoTo 0
End Function

Private Sub SetLastRefreshDate(d As Date)
    On Error Resume Next
    Dim prop As Object ' DocumentProperty
    Set prop = ThisWorkbook.CustomDocumentProperties(PROP_LAST_REFRESH)
    
    If Err.Number = 0 Then
        ' La propriété existe, on met juste à jour la valeur
        prop.Value = d
    Else
        ' La propriété n'existe pas, on l'ajoute
        Err.Clear
        ThisWorkbook.CustomDocumentProperties.Add _
            Name:=PROP_LAST_REFRESH, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeDate, _
            Value:=d
    End If
    On Error GoTo 0
End Sub

' Génère la requête PowerQuery spécifique pour le dictionnaire
Private Function GenerateDictionaryQuery(ByVal URL As String) As String

        Dim q As String: q = Chr(34) ' guillemet double
    Dim lines As Collection: Set lines = New Collection

    ' Récupérer les URLs de toutes les catégories actives
    Dim nbUrls As Long
    Dim urlsList As String
    nbUrls = CategoryManager.CategoriesCount

    ' Construire la liste des paths comme une expression M
    urlsList = "{"
    Dim cat As CategoryInfo
    Dim i As Long
    For i = 1 To nbUrls
        cat = CategoryManager.categories(i)
        ' On prend le path relatif (ex: costing/2.csv)
        Dim pathOnly As String
        pathOnly = Mid(cat.URL, InStr(cat.URL, ".energy/") + 8)
        If InStr(pathOnly, "?") > 0 Then
            pathOnly = Left(pathOnly, InStr(pathOnly, "?") - 1)
        End If
        ' Supprimer l'extension .csv pour le matching
        If Right(pathOnly, 4) = ".csv" Then
            pathOnly = Left(pathOnly, Len(pathOnly) - 4)
        End If
        urlsList = urlsList & q & pathOnly & q
        If i < nbUrls Then urlsList = urlsList & ", "
    Next i
    urlsList = urlsList & "}"
    Log "load_dict", "ValidPaths utilisés pour filtrage PowerQuery (Contains, sans .csv): " & urlsList, DEBUG_LEVEL, "GenerateDictionaryQuery", "RagicDictionary"
    
    ' Construction ligne par ligne
    Dim dq As String
    ' Correction: déclaration et affectation séparées pour compatibilité VBA
     dq = Chr(34)
    lines.Add "let"
    lines.Add "    Source = Csv.Document(Web.Contents(" & dq & URL & dq & "),[Delimiter=" & dq & "," & dq & ", Encoding=65001]),"
    lines.Add "    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),"
    lines.Add "    ValidPaths = " & urlsList & ","
    lines.Add "    FilteredByURL = Table.SelectRows(PromotedHeaders, each List.AnyTrue(List.Transform(ValidPaths, (p) => Text.Contains([URL], p)))), "
    lines.Add "    RemovedCols = Table.RemoveColumns(FilteredByURL, {" & dq & "URL" & dq & ", " & dq & "API URL" & dq & "}),"
    lines.Add "    FilteredRows = Table.SelectRows(RemovedCols, each [SheetName] <> null and [Field Name] <> null)"
    lines.Add "in"
    lines.Add "    FilteredRows"
    Dim result As String: result = ""
    Dim l As Variant
    For Each l In lines
        result = result & l & vbCrLf
    Next l
    GenerateDictionaryQuery = result
End Function

' Vérifie si une requête PowerQuery existe
Public Function QueryExists(queryName As String) As Boolean
    On Error Resume Next
    Dim q As Object
    Set q = ThisWorkbook.Queries(queryName)
    QueryExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function GetOrCreatePQDictSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreatePQDictSheet = ThisWorkbook.Worksheets("PQ_DICT")
    On Error GoTo 0
    
    If GetOrCreatePQDictSheet Is Nothing Then
        ' Créer la feuille si elle n'existe pas
        Log "load_dict", "Création de la feuille PQ_DICT...", INFO_LEVEL, "GetOrCreatePQDictSheet", "RagicDictionary"
        Set GetOrCreatePQDictSheet = ThisWorkbook.Worksheets.Add
        GetOrCreatePQDictSheet.Name = "PQ_DICT"
    End If
    
    ' Force la visibilité de la feuille dans tous les cas
    If GetOrCreatePQDictSheet.Visible <> xlSheetVisible Then
        Log "load_dict", "Rendre la feuille PQ_DICT visible...", INFO_LEVEL, "GetOrCreatePQDictSheet", "RagicDictionary"
        GetOrCreatePQDictSheet.Visible = xlSheetVisible
    End If
    
    Exit Function
ErrorHandler:
    Log "load_dict", "Erreur lors de la création/récupération de PQ_DICT : " & Err.Description, ERROR_LEVEL, "GetOrCreatePQDictSheet", "RagicDictionary"
End Function

Private Sub LoadDictionaryData(ByVal tableName As String)
    On Error GoTo ErrorHandler
    Dim lo As ListObject
    Set lo = wsPQDict.ListObjects(tableName)
    
    Dim r As Long
    Dim key As String, value As String
    Dim sheetNameRaw As String, fieldNameRaw As String
    
    For r = 1 To lo.ListRows.Count
        ' Utilisation explicite des noms de colonnes pour éviter les erreurs d'index
        sheetNameRaw = CStr(lo.DataBodyRange(r, lo.ListColumns("SheetName").Index).Value)
        fieldNameRaw = CStr(lo.DataBodyRange(r, lo.ListColumns("Field Name").Index).Value)
        key = NormalizeSheetName(sheetNameRaw) & "|" & fieldNameRaw
        value = lo.DataBodyRange(r, lo.ListColumns("Field Type").Index).Value
        If Not RagicFieldDict.Exists(key) Then
            RagicFieldDict.Add key, value
        Else
            RagicFieldDict(key) = value ' Mettre à jour si la clé existe déjà
        End If
    Next r
    
    Exit Sub
ErrorHandler:
    Log "load_dict_data_err", "Erreur lors du chargement des données depuis la table " & tableName & " : " & Err.Description, ERROR_LEVEL, "LoadDictionaryData", "RagicDictionary"
End Sub

Public Function ListAllTableNames(ws As Worksheet) As String
    Dim tbl As ListObject, names As String
    For Each tbl In ws.ListObjects
        names = names & tbl.Name & ", "
    Next tbl
    If Len(names) > 2 Then names = Left(names, Len(names) - 2)
    ListAllTableNames = names
End Function

Public Sub TestIsFieldHidden_BudgetGroupes()
    If RagicFieldDict Is Nothing Then
        LoadRagicDictionary
    End If
    Log "test_hidden", "Test IsFieldHidden pour Budget Groupes :", DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    Log "test_hidden", "  SheetName = '? Budget Groupes', FieldName = 'Montant Total' => " & IsFieldHidden("? Budget Groupes", "Montant Total"), DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    Log "test_hidden", "  SheetName = '? Budget Groupes', FieldName = 'Année' => " & IsFieldHidden("? Budget Groupes", "Année"), DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
End Sub

' Recherche à deux étages pour le type de champ
Public Function GetFieldRagicType(categorySheetName As String, fieldName As String) As String
    Log "GetFieldRagicType", "Entrée: SheetName='" & categorySheetName & "', fieldName='" & fieldName & "'", DEBUG_LEVEL, "GetFieldRagicType", "RagicDictionary"
    GetFieldRagicType = "Text" ' Default value
    If wsPQDict Is Nothing Then Set wsPQDict = GetOrCreatePQDictSheet()
    Dim lo As ListObject
    Dim tableName As String
    tableName = "Table_PQ_" & Utilities.SanitizeTableName(BASE_NAME)
    On Error Resume Next
    Set lo = wsPQDict.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Or lo.ListRows.Count = 0 Then Log "GetFieldRagicType", "Table non trouvée ou vide.", DEBUG_LEVEL, "GetFieldRagicType", "RagicDictionary": Exit Function
    Dim arr As Variant
    arr = lo.DataBodyRange.Value
    Dim rowIdx As Long
    rowIdx = FindBestRowForField(lo, categorySheetName, fieldName)
    If rowIdx = 0 Then Log "GetFieldRagicType", "Aucune ligne trouvée.", DEBUG_LEVEL, "GetFieldRagicType", "RagicDictionary": Exit Function
    Dim fieldTypeVal As String
    fieldTypeVal = GetValueFromRow(lo, arr, rowIdx, "Field Type")
    If Trim(fieldTypeVal) <> "" Then
        Log "GetFieldRagicType", "Type trouvé: '" & fieldTypeVal & "'", DEBUG_LEVEL, "GetFieldRagicType", "RagicDictionary"
        GetFieldRagicType = Trim(fieldTypeVal)
    Else
        Log "GetFieldRagicType", "Type vide, fallback 'Text'", DEBUG_LEVEL, "GetFieldRagicType", "RagicDictionary"
    End If
End Function