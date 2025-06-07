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
    tableName = "Table_" & Utilities.SanitizeTableName(BASE_NAME)

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
    needsRefresh = Not tableExists Or ((Now - lastRefresh) * 24 >= 24)
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

Public Function IsFieldHidden(SheetName As String, fieldName As String) As Boolean
    If RagicFieldDict Is Nothing Then
        Log "field_hidden", "RagicFieldDict est Nothing, chargement...", WARNING_LEVEL, "IsFieldHidden", "RagicDictionary"
        LoadRagicDictionary
    End If
    Dim key As String
    key = NormalizeSheetName(SheetName) & "|" & fieldName
    
    Dim memoValue As String
    If RagicFieldDict.Exists(key) Then
        memoValue = RagicFieldDict(key)
        IsFieldHidden = InStr(1, memoValue, "Hidden", vbTextCompare) > 0
        Log "IsFieldHidden_Check", "Key: '" & key & "' | Memo: '" & memoValue & "' | Hidden: " & IsFieldHidden, DEBUG_LEVEL, "IsFieldHidden", "RagicDictionary"
    Else
        IsFieldHidden = False
        Log "IsFieldHidden_Check", "Key: '" & key & "' | Key NOT FOUND in RagicFieldDict. Defaulting to Not Hidden.", WARNING_LEVEL, "IsFieldHidden", "RagicDictionary"
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
    Log "load_dict", "Tables présentes dans PQ_DICT : " & ListAllTableNames(wsPQDict), DEBUG_LEVEL, "LoadDictionaryData", "RagicDictionary"
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsPQDict.ListObjects(tableName)
    On Error GoTo 0

    ' Si le tableau n'existe pas, essayer de le trouver par nom partiel
    If lo Is Nothing Then
        Dim tbl As ListObject
        For Each tbl In wsPQDict.ListObjects
            If InStr(tbl.Name, BASE_NAME) > 0 Then
                Set lo = tbl
                Exit For
            End If
        Next tbl
    End If

    If lo Is Nothing Then
        MsgBox "Le tableau '" & tableName & "' n'a pas été trouvé dans la feuille PQ_DICT." & vbCrLf & _
               "Tableaux présents : " & ListAllTableNames(wsPQDict), vbExclamation
        Exit Sub
    End If

    Dim sheetIdx As Long, fieldIdx As Long, memoIdx As Long
    On Error Resume Next
    sheetIdx = lo.ListColumns("SheetName").Index
    fieldIdx = lo.ListColumns("Field Name").Index
    memoIdx = lo.ListColumns("Memo").Index
    On Error GoTo 0
    
    ' Suppression des colonnes URL et API URL : on ne les utilise plus du tout dans le code.
    ' Si elles étaient utilisées ailleurs, il faudrait aussi les retirer.
    
    If sheetIdx = 0 Or fieldIdx = 0 Or memoIdx = 0 Then
        Log "load_dict", "Colonnes non trouvées dans la table du dictionnaire. Le chargement va échouer.", WARNING_LEVEL, "LoadDictionaryData", "RagicDictionary"
        Exit Sub
    End If

    If lo.DataBodyRange Is Nothing Then
        Log "load_dict", "La table '" & tableName & "' ne contient aucune donnée (DataBodyRange is Nothing).", WARNING_LEVEL, "LoadDictionaryData", "RagicDictionary"
        Exit Sub
    End If
    
    Dim i As Long
    Dim nbLignes As Long
    nbLignes = lo.DataBodyRange.Rows.count
    Log "load_dict", "Nombre de lignes dans le dictionnaire : " & nbLignes, DEBUG_LEVEL, "LoadDictionaryData", "RagicDictionary"

    Dim key As Variant
    For i = 1 To nbLignes
        key = NormalizeSheetName(CStr(lo.DataBodyRange.Cells(i, sheetIdx).Value)) & "|" & _
              CStr(lo.DataBodyRange.Cells(i, fieldIdx).Value)
        If Not RagicFieldDict.Exists(key) Then
            RagicFieldDict.Add key, CStr(lo.DataBodyRange.Cells(i, memoIdx).Value)
        End If
    Next i

    Log "load_dict", "Nombre de clés dans le dictionnaire VBA : " & RagicFieldDict.count, DEBUG_LEVEL, "LoadDictionaryData", "RagicDictionary"
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

Public Function GetFieldRagicType(categorySheetName As String, fieldName As String) As String
    ' Ensure dictionary is loaded and wsPQDict is available
    If RagicFieldDict Is Nothing Or RagicFieldDict.Count = 0 Then
        Log "GetFieldRagicType", "RagicFieldDict is not initialized or empty, calling LoadRagicDictionary.", INFO_LEVEL, "GetFieldRagicType", "RagicDictionary"
        LoadRagicDictionary
    End If
    
    If wsPQDict Is Nothing Then
        Log "GetFieldRagicType", "wsPQDict (Ragic Dictionary sheet) is not initialized.", ERROR_LEVEL, "GetFieldRagicType", "RagicDictionary"
        GetFieldRagicType = "Text" ' Default on critical error
        Exit Function
    End If

    Dim dictTable As ListObject
    Dim tableName As String
    tableName = "Table_" & Utilities.SanitizeTableName(BASE_NAME) ' BASE_NAME is "RagicDictionary"

    On Error Resume Next
    Set dictTable = wsPQDict.ListObjects(tableName)
    On Error GoTo 0

    If dictTable Is Nothing Then
        Log "GetFieldRagicType", "Ragic Dictionary table '" & tableName & "' not found on sheet '" & wsPQDict.Name & "'.", ERROR_LEVEL, "GetFieldRagicType", "RagicDictionary"
        GetFieldRagicType = "Text" ' Default if table not found
        Exit Function
    End If

    Dim normalizedCategoryKey As String
    normalizedCategoryKey = NormalizeSheetName(categorySheetName)

    ' Define column names in the Ragic Dictionary table
    Const COL_CATEGORY_IDENTIFIER_IN_TABLE As String = "SheetName"   ' Column storing the category identifier (e.g., "module/1" or sheet name)
    Const COL_FIELD_NAME_IN_TABLE As String = "Field Label"     ' Column storing the field's name/label
    Const COL_FIELD_TYPE_IN_TABLE As String = "Field Type"      ' Column storing the field's type (e.g., "DATE", "NUMBER")

    Dim categoryCol As ListColumn, fieldNameCol As ListColumn, fieldTypeCol As ListColumn
    Dim r As Long
    Dim currentDictFieldName As String, currentDictFieldType As String, currentDictCategoryValue As String

    On Error Resume Next
    Set categoryCol = dictTable.ListColumns(COL_CATEGORY_IDENTIFIER_IN_TABLE)
    Set fieldNameCol = dictTable.ListColumns(COL_FIELD_NAME_IN_TABLE)
    Set fieldTypeCol = dictTable.ListColumns(COL_FIELD_TYPE_IN_TABLE)
    On Error GoTo 0

    If categoryCol Is Nothing Then
        Log "GetFieldRagicType", "Category column '" & COL_CATEGORY_IDENTIFIER_IN_TABLE & "' not found in table '" & tableName & "'.", ERROR_LEVEL, "GetFieldRagicType", "RagicDictionary"
        GetFieldRagicType = "Text"
        Exit Function
    End If
    If fieldNameCol Is Nothing Then
        Log "GetFieldRagicType", "Field Name column '" & COL_FIELD_NAME_IN_TABLE & "' not found in table '" & tableName & "'.", ERROR_LEVEL, "GetFieldRagicType", "RagicDictionary"
        GetFieldRagicType = "Text"
        Exit Function
    End If
    If fieldTypeCol Is Nothing Then
        Log "GetFieldRagicType", "Field Type column '" & COL_FIELD_TYPE_IN_TABLE & "' not found in table '" & tableName & "'.", ERROR_LEVEL, "GetFieldRagicType", "RagicDictionary"
        GetFieldRagicType = "Text"
        Exit Function
    End If

    ' Default return type
    GetFieldRagicType = "Text"

    If dictTable.DataBodyRange Is Nothing Then
        Log "GetFieldRagicType", "Ragic Dictionary table '" & tableName & "' is empty.", WARNING_LEVEL, "GetFieldRagicType", "RagicDictionary"
        Exit Function ' No data to iterate
    End If

    For r = 1 To dictTable.ListRows.Count
        currentDictCategoryValue = NormalizeSheetName(CStr(dictTable.DataBodyRange.Cells(r, categoryCol.Index).Value))
        currentDictFieldName = CStr(dictTable.DataBodyRange.Cells(r, fieldNameCol.Index).Value)

        ' Match category (normalized) and field name (exact)
        If normalizedCategoryKey = currentDictCategoryValue And fieldName = currentDictFieldName Then
            ' Field found. Now determine its type.
            ' Check if the field name from the dictionary matches the "Section" pattern.
            If currentDictFieldName Like "__*__" And Right(currentDictFieldName, 2) = "__" Then
                GetFieldRagicType = "Section"
                Exit Function
            End If

            ' If not a section, check its "Field Type" column.
            currentDictFieldType = UCase(CStr(dictTable.DataBodyRange.Cells(r, fieldTypeCol.Index).Value))

            Select Case currentDictFieldType
                Case "DATE" ' Adapt if Ragic uses a different string e.g. "DATE_INPUT"
                    GetFieldRagicType = "Date"
                Case "NUMBER", "NUMERIC", "INTEGER", "FLOAT" ' Adapt for various numeric types from Ragic
                    GetFieldRagicType = "Number"
                Case Else
                    ' Includes "FREETEXT", "SELECT", "MULTISELECT", "TEXTAREA", etc.
                    GetFieldRagicType = "Text"
            End Select
            Exit Function ' Found and processed, no need to loop further
        End If
    Next r

    ' If loop completes, the specific field was not found in the dictionary for the given category,
    ' or its type didn't match "Date" or "Number" explicitly.
    Log "GetFieldRagicType", "Field '" & fieldName & "' for category '" & categorySheetName & "' (normalized key: '" & normalizedCategoryKey & "') not found in dictionary or type not specifically handled. Defaulting to '" & GetFieldRagicType & "'.", INFO_LEVEL, "GetFieldRagicType", "RagicDictionary"

End Function