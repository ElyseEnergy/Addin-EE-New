Attribute VB_Name = "RagicDictionary"
Option Explicit

Public RagicFieldDict As Object
Public wsPQDict As Worksheet

' Constantes pour les noms et propriétés
Private Const BASE_NAME As String = "RagicDictionary"
Private Const RAGIC_PATH As String = "matching-matrix/6.csv"
Private Const PROP_LAST_REFRESH As String = "RagicDictLastRefresh"

'==================================================================================================
' CALLBACK DU RUBAN
'==================================================================================================

' Callback pour le bouton du ruban pour forcer le rafraîchissement
Public Sub ProcessForceRefreshRagicDictionary(ByVal control As IRibbonControl)
    ForceRefreshRagicDictionary
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

    ' Décider s'il faut rafraîchir depuis le réseau
    Dim needsRefresh As Boolean
    needsRefresh = Not tableExists Or (VBA.Date - lastRefresh >= 1)

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
        LoadQueries.LoadQuery pqName, wsPQDict, wsPQDict.Range("A1")

        ' Mettre à jour la date du rafraîchissement dans les propriétés du classeur
        SetLastRefreshDate VBA.Date
        
        ' Sauvegarder le classeur pour rendre la date de mise à jour persistante
        On Error Resume Next
        ThisWorkbook.Save
        If Err.Number <> 0 Then
            Log "ragic_dict_err", "Impossible de sauvegarder le classeur après mise à jour du dictionnaire. La date ne sera pas persistante.", WARNING_LEVEL, "LoadRagicDictionary", "RagicDictionary"
        Else
            Log "load_dict", "Classeur sauvegardé pour persistance de la date de mise à jour.", DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"
        End If
        On Error GoTo 0
        
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
    If RagicFieldDict.Exists(key) Then
        IsFieldHidden = InStr(1, RagicFieldDict(key), "Hidden", vbTextCompare) > 0
    Else
        IsFieldHidden = False
    End If
End Function

' Normalise le nom de la feuille pour la clé dictionnaire
Public Function NormalizeSheetName(SheetName As String) As String
    Dim i As Long, c As String
    ' Supprime tous les caractères non alphanumériques au début
    For i = 1 To Len(SheetName)
        c = Mid(SheetName, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Then
            NormalizeSheetName = Mid(SheetName, i)
            Exit Function
        End If
    Next i
    NormalizeSheetName = SheetName ' fallback
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
    Dim q As String
    q = Chr(34) ' guillemet double
    GenerateDictionaryQuery = _
        "let" & vbCrLf & _
        "    Source = Csv.Document(Web.Contents(" & q & URL & q & "),[Delimiter=" & q & "," & q & ", Encoding=65001])," & vbCrLf & _
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
        GetOrCreatePQDictSheet.visible = xlSheetVeryHidden
    End If
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
        Log "load_dict", "Colonnes non trouvées dans la table du dictionnaire. Le chargement va échouer.", WARNING_LEVEL, "LoadDictionaryData", "RagicDictionary"
        Exit Sub
    End If

    Dim i As Long
    Dim nbLignes As Long
    nbLignes = lo.DataBodyRange.Rows.Count
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
End Sub

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
        LoadRagicDictionary
    End If
    Log "test_hidden", "Test IsFieldHidden pour Budget Groupes :", DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    Log "test_hidden", "  SheetName = '? Budget Groupes', FieldName = 'Montant Total' => " & IsFieldHidden("? Budget Groupes", "Montant Total"), DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    Log "test_hidden", "  SheetName = '? Budget Groupes', FieldName = 'Année' => " & IsFieldHidden("? Budget Groupes", "Année"), DEBUG_LEVEL, "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
End Sub




