Option Explicit

Public RagicFieldDict As Object
Public wsPQDict As Worksheet

' Constantes pour les noms
Private Const BASE_NAME As String = "RagicDictionary"
Private Const RAGIC_PATH As String = "matching-matrix/6.csv"

Public Sub LoadRagicDictionary()
    ' Afficher un message à l'utilisateur
    Application.StatusBar = "Chargement du dictionnaire Ragic en cours..."
    MsgBox "Le premier chargement du dictionnaire Ragic peut prendre quelques instants." & vbCrLf & _
           "Veuillez patienter...", vbInformation, "Chargement en cours"
    
    ' Créer ou récupérer la feuille PQ_DICT
    Set wsPQDict = GetOrCreatePQDictSheet()
    
    ' Définir les noms standardisés
    Dim pqName As String
    pqName = "PQ_" & Utilities.SanitizeTableName(BASE_NAME)
    Dim tableName As String
    tableName = "Table_" & Utilities.SanitizeTableName(BASE_NAME)
    
    ' Créer une catégorie pour le dictionnaire
    Dim dictCategory As CategoryInfo
    With dictCategory
        .categoryName = BASE_NAME
        .displayName = BASE_NAME
        .URL = env.RAGIC_BASE_URL & RAGIC_PATH & env.RAGIC_API_PARAMS
        .PowerQueryName = pqName
        .SheetName = BASE_NAME
    End With
    
    ' Vérifier si la requête existe déjà
    If QueryExists(pqName) Then
        ' Si la requête existe, mettre à jour sa formule
        On Error Resume Next
        ThisWorkbook.Queries(pqName).Formula = GenerateDictionaryQuery(dictCategory.URL)
        Dim updateError As Long
        updateError = Err.Number
        On Error GoTo 0
        
        If updateError <> 0 Then
            Log "ragic_dict_err", "Erreur lors de la mise à jour de la requête " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
            Application.StatusBar = False
            Exit Sub
        End If
        
        ' Rafraîchir la requête
        ThisWorkbook.Queries(pqName).Refresh
    Else
        ' Créer la requête si elle n'existe pas
        On Error Resume Next
        ThisWorkbook.Queries.Add pqName, GenerateDictionaryQuery(dictCategory.URL)
        Dim addError As Long
        addError = Err.Number
        On Error GoTo 0
          If addError <> 0 Then
            Log "ragic_dict_err", "Erreur lors de l'ajout de la requête " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
            Application.StatusBar = False
            Exit Sub
        End If
    End If
    
    ' Charger les données dans la feuille PQ_DICT
    LoadQueries.LoadQuery pqName, wsPQDict, wsPQDict.Range("A1")
    
    ' Initialiser le dictionnaire
    Set RagicFieldDict = CreateObject("Scripting.Dictionary")
    
    ' Charger les données dans le dictionnaire
    LoadDictionaryData tableName
    
    ' Réinitialiser la barre de statut
    Application.StatusBar = False
End Sub

' Génère la requête PowerQuery spécifique pour le dictionnaire
Private Function GenerateDictionaryQuery(ByVal url As String) As String
    Dim q As String
    q = Chr(34) ' guillemet double
    GenerateDictionaryQuery = _
        "let" & vbCrLf & _
        "    Source = Csv.Document(Web.Contents(" & q & url & q & "),[Delimiter=" & q & "," & q & ", Encoding=65001])," & vbCrLf & _
        "    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & vbCrLf & _
        "    FilteredRows = Table.SelectRows(PromotedHeaders, each [SheetName] <> null and [Field Name] <> null)" & vbCrLf & _
        "in" & vbCrLf & _
        "    FilteredRows"
End Function

' Vérifie si une requête PowerQuery existe
Private Function QueryExists(queryName As String) As Boolean
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
        Set GetOrCreatePQDictSheet = ThisWorkbook.Worksheets.Add
        GetOrCreatePQDictSheet.Name = "PQ_DICT"
        
        ' Masquer la feuille
        GetOrCreatePQDictSheet.Visible = xlSheetVeryHidden
    End If
End Function

' Normalise le nom de la feuille pour la clé dictionnaire
Public Function NormalizeSheetName(sheetName As String) As String
    Dim i As Long, c As String
    ' Supprime tous les caractères non alphanumériques au début
    For i = 1 To Len(sheetName)
        c = Mid(sheetName, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Then
            NormalizeSheetName = Mid(sheetName, i)
            Exit Function
        End If
    Next i
    NormalizeSheetName = sheetName ' fallback
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
    
    If sheetIdx = 0 Or fieldIdx = 0 Or memoIdx = 0 Then
        DataLoaderManager.CleanupPowerQuery "PQ_" & Utilities.SanitizeTableName(BASE_NAME)
        Exit Sub
    End If

    Dim i As Long
    Dim nbLignes As Long    nbLignes = lo.DataBodyRange.Rows.Count
    Log "load_dict", "Nombre de lignes dans le dictionnaire : " & nbLignes, DEBUG_LEVEL, "LoadDictionaryData", "RagicDictionary"

    Dim key As Variant
    For i = 1 To nbLignes
        key = NormalizeSheetName(CStr(lo.DataBodyRange.Cells(i, sheetIdx).Value)) & "|" & _
              CStr(lo.DataBodyRange.Cells(i, fieldIdx).Value)
        If Not RagicFieldDict.Exists(key) Then
            RagicFieldDict.Add key, CStr(lo.DataBodyRange.Cells(i, memoIdx).Value)
        End If
    Next i

    Log "load_dict", "Nombre de clés dans le dictionnaire VBA : " & RagicFieldDict.Count, DEBUG_LEVEL, "LoadDictionaryData", "RagicDictionary"
    Dim c As Long
    c = 0
    For Each key In RagicFieldDict.Keys
        Log "load_dict", "  " & key & " => " & RagicFieldDict(key), DEBUG_LEVEL, "LoadDictionaryData", "RagicDictionary"
        c = c + 1
        If c > 10 Then Exit For
    Next key

    DataLoaderManager.CleanupPowerQuery "PQ_" & Utilities.SanitizeTableName(BASE_NAME)
End Sub

Public Function IsFieldHidden(sheetName As String, fieldName As String) As Boolean
    If RagicFieldDict Is Nothing Then
        Log "field_hidden", "RagicFieldDict est Nothing dans IsFieldHidden", WARNING_LEVEL, "IsFieldHidden", "RagicDictionary"
        Exit Function
    End If
    Dim key As String
    key = NormalizeSheetName(sheetName) & "|" & fieldName
    If RagicFieldDict.Exists(key) Then
        IsFieldHidden = InStr(1, RagicFieldDict(key), "Hidden", vbTextCompare) > 0
    Else
        IsFieldHidden = False
    End If
End Function

Private Function ListAllTableNames(ws As Worksheet) As String
    Dim tbl As ListObject, names As String
    For Each tbl In ws.ListObjects
        names = names & tbl.Name & ", "
    Next tbl
    If Len(names) > 2 Then names = Left(names, Len(names) - 2)
    ListAllTableNames = names
End Function

Public Sub TestIsFieldHidden_BudgetGroupes()
    If RagicFieldDict Is Nothing Then
        Log "test_hidden", "Dictionnaire non initialisé, chargement en cours...", INFO_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
        LoadRagicDictionary
    End If
    Log "test_hidden", "Test IsFieldHidden pour Budget Groupes :", DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    Log "test_hidden", "Champ 1 :", DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    Log "test_hidden", "  SheetName = '↳ Budget Groupes', FieldName = 'Montant Total'", DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    Log "test_hidden", "  Résultat : " & IsFieldHidden("↳ Budget Groupes", "Montant Total"), DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    
    Log "test_hidden", "Champ 2 :", DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    Log "test_hidden", "  SheetName = '↳ Budget Groupes', FieldName = 'Année'", DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    Log "test_hidden", "  Résultat : " & IsFieldHidden("↳ Budget Groupes", "Année"), DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
End Sub


