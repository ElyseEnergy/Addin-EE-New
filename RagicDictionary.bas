Attribute VB_Name = "RagicDictionary"

Option Explicit

Public RagicFieldDict As Object
Public wsPQDict As Worksheet

' Constantes pour les noms et propri�t�s
Private Const BASE_NAME As String = "RagicDictionary"
Private Const RAGIC_PATH As String = "matching-matrix/6.csv"
Private Const PROP_LAST_REFRESH As String = "RagicDictLastRefresh"

'==================================================================================================
' CALLBACKS DU RUBAN
'==================================================================================================

' Callback pour le bouton du ruban pour forcer le rafra�chissement
Public Sub ProcessForceRefreshRagicDictionary(ByVal control As IRibbonControl)
    ForceRefreshRagicDictionary
End Sub

' Callback pour l'info-bulle (supertip) du bouton de rafra�chissement
Public Sub GetRagicDictSupertip(control As IRibbonControl, ByRef supertip)
    Dim lastUpdate As Date
    lastUpdate = GetLastRefreshDate() ' On r�utilise la fonction existante
    
    Dim lastUpdateText As String
    If lastUpdate > 0 Then
        lastUpdateText = "Last update: " & Format(lastUpdate, "yyyy-mm-dd")
    Else
        lastUpdateText = "Never updated. Click to download."
    End If
    
    supertip = "Downloads the latest version of the data dictionary from Ragic." & vbCrLf & vbCrLf & lastUpdateText
End Sub

'==================================================================================================
' M�THODES PUBLIQUES
'==================================================================================================

' Force le rafra�chissement du dictionnaire depuis Ragic
Public Sub ForceRefreshRagicDictionary()
    Application.StatusBar = "For�age du rafra�chissement du dictionnaire Ragic..."
    ' R�initialiser la date dans les propri�t�s pour forcer le rechargement
    SetLastRefreshDate (0)
    ' Appeler la routine de chargement
    LoadRagicDictionary
    Application.StatusBar = False
    MsgBox "Le dictionnaire Ragic a �t� mis � jour.", vbInformation
End Sub

' Charge le dictionnaire Ragic, depuis le cache si possible
Public Sub LoadRagicDictionary()
    ' S'assurer que la liste des cat�gories est initialis�e
    If CategoryManager.CategoriesCount = 0 Then
        CategoryManager.InitCategories
    End If
    
    Dim lastRefresh As Date
    lastRefresh = GetLastRefreshDate()

    ' On ne v�rifie plus si l'objet RagicFieldDict existe, car il est toujours vide � l'ouverture.
    ' La logique se base sur l'existence de la table dans la feuille de cache et la date.

    Application.StatusBar = "V�rification du dictionnaire Ragic..."

    ' D�finir les noms standardis�s
    Dim pqName As String
    pqName = "PQ_" & Utilities.SanitizeTableName(BASE_NAME)
    Dim tableName As String
    tableName = "Table_" & Utilities.SanitizeTableName(BASE_NAME)

    ' Cr�er ou r�cup�rer la feuille PQ_DICT
    Set wsPQDict = GetOrCreatePQDictSheet()

    ' V�rifier si la table de cache existe d�j� dans la feuille
    Dim tableExists As Boolean
    On Error Resume Next
    tableExists = (wsPQDict.ListObjects(tableName).Name <> "")
    On Error GoTo 0
    Log "load_dict", "Table " & tableName & " existe: " & tableExists, DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"
    
    ' D�cider s'il faut rafra�chir depuis le r�seau
    Dim needsRefresh As Boolean
    needsRefresh = Not tableExists Or ((Now - lastRefresh) * 24 >= 24)
    Log "load_dict", "Derni�re M�J: " & lastRefresh & ", �ge (heures): " & ((Now - lastRefresh) * 24) & "h", DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"
    Log "load_dict", "Rafra�chissement n�cessaire: " & needsRefresh & " (table existe: " & tableExists & ")", DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"

    If needsRefresh Then
        Application.StatusBar = "Chargement du dictionnaire Ragic depuis le r�seau..."
        Log "load_dict", "Rafra�chissement du dictionnaire Ragic depuis le r�seau.", INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"

        ' Cr�er une cat�gorie pour le dictionnaire
        Dim dictCategory As categoryInfo
        With dictCategory
            .categoryName = BASE_NAME
            .DisplayName = BASE_NAME
            .URL = env.RAGIC_BASE_URL & RAGIC_PATH & env.RAGIC_API_PARAMS
            .PowerQueryName = pqName
            .SheetName = BASE_NAME
        End With

        Log "load_dict", "URL de requ�te dictionnaire : " & dictCategory.URL, INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"

        ' Cr�er ou mettre � jour la requ�te
        If QueryExists(pqName) Then
            On Error Resume Next
            ActiveWorkbook.Queries(pqName).formula = GenerateDictionaryQuery(dictCategory.URL)
            If Err.Number <> 0 Then
                Log "ragic_dict_err", "Erreur M�J requ�te " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
                Application.StatusBar = False
                Exit Sub
            End If
            On Error GoTo 0
        Else
            On Error Resume Next
            ActiveWorkbook.Queries.Add pqName, GenerateDictionaryQuery(dictCategory.URL)
            If Err.Number <> 0 Then
                Log "ragic_dict_err", "Erreur ajout requ�te " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
                Application.StatusBar = False
                Exit Sub
            End If
            On Error GoTo 0
        End If

        ' Rafraîchir la requête (c'est l'étape lente)
        On Error Resume Next
        ActiveWorkbook.Queries(pqName).Refresh
        If Err.Number <> 0 Then
            Log "ragic_dict_err", "Erreur Refresh requ�te " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
            Application.StatusBar = False
            Exit Sub
        End If
        On Error GoTo 0

        ' Charger les données dans la feuille PQ_DICT
        ' Si la table existe déjà, forcer un refresh de la requête associée
        If QueryExists(pqName) Then
            On Error Resume Next
            ActiveWorkbook.Queries(pqName).Refresh
            If Err.Number <> 0 Then
                Log "ragic_dict_err", "Erreur Refresh requ�te existante " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
                Application.StatusBar = False
                Exit Sub
            End If
            On Error GoTo 0
        Else
            LoadQueries.LoadQuery pqName, wsPQDict, wsPQDict.Range("A1")
        End If

        ' Mettre � jour la date du rafra�chissement dans les propri�t�s du classeur
        SetLastRefreshDate VBA.Date
        
        ' Sauvegarder le classeur pour rendre la date de mise � jour persistante
        Log "load_dict", "Tentative de sauvegarde forc�e du classeur pour persistance de la date...", INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"
        On Error Resume Next
        ' On marque explicitement le classeur comme ayant des modifications non enregistr�es pour forcer la sauvegarde
        ThisWorkbook.Saved = False
        ThisWorkbook.Save
        If Err.Number <> 0 Then
            Log "ragic_dict_err", "ERREUR lors de la sauvegarde du classeur: " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
        Else
            Log "load_dict", "Classeur sauvegard� avec succ�s. La date de mise � jour est maintenant persistante.", INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"
        End If
        On Error GoTo 0
        
        ' Invalider le contr�le du ruban pour mettre � jour l'info-bulle
        If Not RibbonVisibility.gRibbon Is Nothing Then
            RibbonVisibility.gRibbon.InvalidateControl "btnForceRefreshRagic"
        End If
    Else
        Log "load_dict", "Chargement du dictionnaire Ragic depuis le cache local (feuille PQ_DICT).", INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"
    End If

    ' Initialiser et charger les donn�es dans le dictionnaire VBA
    Application.StatusBar = "Finalisation du chargement du dictionnaire..."
    If RagicFieldDict Is Nothing Then
        Set RagicFieldDict = CreateObject("Scripting.Dictionary")
    Else
        RagicFieldDict.RemoveAll
    End If

    LoadDictionaryData tableName

    ' Log des 10 premi�res cl�s du dictionnaire pour debug
    Dim debugKeys As String, debugCount As Long
    debugKeys = ""
    For debugCount = 0 To Application.Min(9, RagicFieldDict.count - 1)
        debugKeys = debugKeys & RagicFieldDict.Keys()(debugCount) & "; "
    Next debugCount
    Log "load_dict", "Premi�res cl�s du dictionnaire : " & debugKeys, DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"

    ' Forcer la visibilit� de la feuille PQ_DICT
    If Not wsPQDict Is Nothing Then
        wsPQDict.visible = xlSheetVisible
    End If

    ' R�initialiser la barre de statut
    Application.StatusBar = False
End Sub

Public Function IsFieldHidden(SheetName As String, fieldName As String) As Boolean
    If RagicFieldDict Is Nothing Then
        Log "field_hidden", "RagicFieldDict est Nothing, chargement...", WARNING_LEVEL, "IsFieldHidden", "RagicDictionary"
        LoadRagicDictionary
    End If
    Dim key As String
    key = NormalizeSheetName(SheetName) & "|" & fieldName
    If RagicFieldDict.Exists(key) Then
        IsFieldHidden = InStr(1, RagicFieldDict(key), "Hidden", vbTextCompare) > 0
    Else
        IsFieldHidden = False
    End If
End Function

' Normalise le nom de la feuille pour la cl� dictionnaire
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
' M�THODES PRIV�ES
'==================================================================================================

' G�re la persistance de la date de rafra�chissement via les propri�t�s du document
Private Function GetLastRefreshDate() As Date
    On Error Resume Next
    GetLastRefreshDate = ThisWorkbook.CustomDocumentProperties(PROP_LAST_REFRESH).Value
    If Err.Number <> 0 Then
        GetLastRefreshDate = 0 ' Force le rafra�chissement si la propri�t� n'existe pas
    End If
    On Error GoTo 0
End Function

Private Sub SetLastRefreshDate(d As Date)
    On Error Resume Next
    Dim prop As Object ' DocumentProperty
    Set prop = ThisWorkbook.CustomDocumentProperties(PROP_LAST_REFRESH)
    
    If Err.Number = 0 Then
        ' La propri�t� existe, on met juste � jour la valeur
        prop.Value = d
    Else
        ' La propri�t� n'existe pas, on l'ajoute
        Err.Clear
        ThisWorkbook.CustomDocumentProperties.Add _
            Name:=PROP_LAST_REFRESH, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeDate, _
            Value:=d
    End If
    On Error GoTo 0
End Sub

' G�n�re la requ�te PowerQuery sp�cifique pour le dictionnaire
Private Function GenerateDictionaryQuery(ByVal URL As String) As String

        Dim q As String: q = Chr(34) ' guillemet double
    Dim lines As Collection: Set lines = New Collection

    ' R�cup�rer les URLs de toutes les cat�gories actives
    Dim nbUrls As Long
    Dim urlsList As String
    nbUrls = CategoryManager.CategoriesCount

    ' Construire la liste des paths comme une expression M
    urlsList = "{"
    Dim cat As categoryInfo
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
    Log "load_dict", "ValidPaths utilis�s pour filtrage PowerQuery (Contains, sans .csv): " & urlsList, DEBUG_LEVEL, "GenerateDictionaryQuery", "RagicDictionary"
    
    ' Construction ligne par ligne
    Dim dq As String
    ' Correction: d�claration et affectation s�par�es pour compatibilit� VBA
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

' Vérifie si une requête PowerQuery existe dans le classeur actif
Public Function QueryExists(QueryName As String) As Boolean
    On Error Resume Next
    Dim q As Object
    Set q = ActiveWorkbook.Queries(QueryName)
    QueryExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function GetOrCreatePQDictSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreatePQDictSheet = ThisWorkbook.Worksheets("PQ_DICT")
    On Error GoTo 0
    
    If GetOrCreatePQDictSheet Is Nothing Then
        ' Cr�er la feuille si elle n'existe pas
        Log "load_dict", "Cr�ation de la feuille PQ_DICT...", INFO_LEVEL, "GetOrCreatePQDictSheet", "RagicDictionary"
        Set GetOrCreatePQDictSheet = ThisWorkbook.Worksheets.Add
        GetOrCreatePQDictSheet.Name = "PQ_DICT"
    End If
    
    ' Force la visibilit� de la feuille dans tous les cas
    If GetOrCreatePQDictSheet.visible <> xlSheetVisible Then
        Log "load_dict", "Rendre la feuille PQ_DICT visible...", INFO_LEVEL, "GetOrCreatePQDictSheet", "RagicDictionary"
        GetOrCreatePQDictSheet.visible = xlSheetVisible
    End If
    
    Exit Function
ErrorHandler:
    Log "load_dict", "Erreur lors de la cr�ation/r�cup�ration de PQ_DICT : " & Err.Description, ERROR_LEVEL, "GetOrCreatePQDictSheet", "RagicDictionary"
End Function

Private Sub LoadDictionaryData(ByVal tableName As String)
    Log "load_dict", "Tables pr�sentes dans PQ_DICT : " & ListAllTableNames(wsPQDict), DEBUG_LEVEL, "LoadDictionaryData", "RagicDictionary"
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
        MsgBox "Le tableau '" & tableName & "' n'a pas �t� trouv� dans la feuille PQ_DICT." & vbCrLf & _
               "Tableaux pr�sents : " & ListAllTableNames(wsPQDict), vbExclamation
        Exit Sub
    End If

    Dim sheetIdx As Long, fieldIdx As Long, memoIdx As Long
    On Error Resume Next
    sheetIdx = lo.ListColumns("SheetName").index
    fieldIdx = lo.ListColumns("Field Name").index
    memoIdx = lo.ListColumns("Memo").index
    On Error GoTo 0
    
    ' Suppression des colonnes URL et API URL : on ne les utilise plus du tout dans le code.
    ' Si elles �taient utilis�es ailleurs, il faudrait aussi les retirer.
    
    If sheetIdx = 0 Or fieldIdx = 0 Or memoIdx = 0 Then
        Log "load_dict", "Colonnes non trouv�es dans la table du dictionnaire. Le chargement va �chouer.", WARNING_LEVEL, "LoadDictionaryData", "RagicDictionary"
        Exit Sub
    End If

    If lo.DataBodyRange Is Nothing Then
        Log "load_dict", "La table '" & tableName & "' ne contient aucune donn�e (DataBodyRange is Nothing).", WARNING_LEVEL, "LoadDictionaryData", "RagicDictionary"
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

    Log "load_dict", "Nombre de cl�s dans le dictionnaire VBA : " & RagicFieldDict.count, DEBUG_LEVEL, "LoadDictionaryData", "RagicDictionary"
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
    Log "test_hidden", "  SheetName = '? Budget Groupes', FieldName = 'Ann�e' => " & IsFieldHidden("? Budget Groupes", "Ann�e"), DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
End Sub

' Fixed GetFieldRagicType function with correct column names
' Replace the existing GetFieldRagicType function with this version

' Simple fix using the actual column names from the RagicDict
Public Function GetFieldRagicType(categorySheetName As String, fieldName As String) As String
    On Error GoTo ErrorHandler
    
    ' Quick check for Section pattern first
    If fieldName Like "__*__" And Left(fieldName, 2) = "__" And Right(fieldName, 2) = "__" Then
        GetFieldRagicType = "Section"
        Exit Function
    End If
    
    ' Ensure dictionary is loaded
    If RagicFieldDict Is Nothing Or RagicFieldDict.count = 0 Then
        LoadRagicDictionary
    End If
    
    If wsPQDict Is Nothing Then
        GetFieldRagicType = "Text"
        Exit Function
    End If

    ' Get the dictionary table
    Dim dictTable As ListObject
    Dim tableName As String
    tableName = "Table_PQ_" & Utilities.SanitizeTableName(BASE_NAME)

    On Error Resume Next
    Set dictTable = wsPQDict.ListObjects(tableName)
    On Error GoTo ErrorHandler

    If dictTable Is Nothing Then
        GetFieldRagicType = "Text"
        Exit Function
    End If

    ' Use the EXACT column names from the log
    Dim categoryCol As ListColumn, fieldNameCol As ListColumn, fieldTypeCol As ListColumn
    
    On Error Resume Next
    Set categoryCol = dictTable.ListColumns("SheetName")
    Set fieldNameCol = dictTable.ListColumns("Field Name")
    Set fieldTypeCol = dictTable.ListColumns("Field Type")
    On Error GoTo ErrorHandler

    If categoryCol Is Nothing Or fieldNameCol Is Nothing Or fieldTypeCol Is Nothing Then
        GetFieldRagicType = "Text"
        Exit Function
    End If

    If dictTable.DataBodyRange Is Nothing Then
        GetFieldRagicType = "Text"
        Exit Function
    End If

    ' Search in dictionary
    Dim normalizedCategoryKey As String
    normalizedCategoryKey = NormalizeSheetName(categorySheetName)
    
    Dim r As Long
    For r = 1 To dictTable.ListRows.count
        Dim currentCategory As String, currentFieldName As String, currentFieldType As String
        
        On Error Resume Next
        currentCategory = NormalizeSheetName(CStr(dictTable.DataBodyRange.Cells(r, categoryCol.index).Value))
        currentFieldName = CStr(dictTable.DataBodyRange.Cells(r, fieldNameCol.index).Value)
        currentFieldType = UCase(CStr(dictTable.DataBodyRange.Cells(r, fieldTypeCol.index).Value))
        On Error GoTo ErrorHandler

        ' Match found
        If normalizedCategoryKey = currentCategory And fieldName = currentFieldName Then
            ' Check section pattern again
            If currentFieldName Like "__*__" And Left(currentFieldName, 2) = "__" And Right(currentFieldName, 2) = "__" Then
                GetFieldRagicType = "Section"
                Exit Function
            End If

            ' Map field type
            Select Case currentFieldType
                Case "DATE", "DATE_INPUT"
                    GetFieldRagicType = "Date"
                Case "NUMBER", "NUMERIC", "INTEGER", "FLOAT", "DECIMAL"
                    GetFieldRagicType = "Number"
                Case Else
                    GetFieldRagicType = "Text"
            End Select
            Exit Function
        End If
    Next r

    ' Not found in dictionary - default to Text
    GetFieldRagicType = "Text"
    Exit Function

ErrorHandler:
    GetFieldRagicType = "Text"
End Function
