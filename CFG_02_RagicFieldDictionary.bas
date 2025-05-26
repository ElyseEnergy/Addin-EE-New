' =============================================================================
' MODULE: CFG_02_RagicFieldDictionary
' Description: Manages the Ragic field dictionary mapping
' =============================================================================
Option Explicit
Private Const MODULE_NAME As String = "CFG_02_RagicFieldDictionary"

' Constants for message dialogs
Private Const MSG_TITLE_INFO As String = "Information"
Private Const MSG_TITLE_ERROR As String = "Erreur"
Private Const MSG_FIRST_LOAD As String = "Le premier chargement du dictionnaire Ragic peut prendre quelques instants." & vbCrLf & _
                                       "Veuillez patienter pendant que nous configurons tout."

' Constantes pour les noms
Private Const BASE_NAME As String = "RagicDictionary"
Private Const RAGIC_PATH As String = "matching-matrix/6.csv"
Private Const QUERY_NAME As String = "PQ_DICT"

' Dépendances
' - SYS_Logger pour le logging
' - SYS_MessageBox pour les messages utilisateur
' - SYS_ErrorHandler pour la gestion des erreurs

Public RagicFieldDict As Object
Public wsPQDict As Worksheet

Public Sub LoadRagicDictionary()
    Const PROC_NAME As String = "LoadRagicDictionary"
    On Error GoTo ErrorHandler
    
    ' Mettre à jour la requête
    On Error Resume Next
    ThisWorkbook.Queries(QUERY_NAME).Formula = GenerateDictionaryQuery(env.RAGIC_DICT_PATH)
    If Err.Number <> 0 Then
        LogError "RagicDict.Load.QueryUpdate", Err.Number, "Erreur lors de la mise à jour de la requête " & QUERY_NAME & ": " & Err.Description, "LoadRagicDictionary", "RagicDictionary"
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Ajouter la requête si elle n'existe pas
    On Error Resume Next
    ThisWorkbook.Queries.Add(QUERY_NAME, GenerateDictionaryQuery(env.RAGIC_DICT_PATH))
    If Err.Number <> 0 Then
        LogError "RagicDict.Load.QueryAdd", Err.Number, "Erreur lors de l'ajout de la requête " & QUERY_NAME & ": " & Err.Description, "LoadRagicDictionary", "RagicDictionary"
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Charger les données
    LoadDictionaryData
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
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

Private Sub LoadDictionaryData()
    Const PROC_NAME As String = "LoadDictionaryData"
    On Error GoTo ErrorHandler
    
    ' Vérifier les tables présentes
    LogDebug "RagicDict.LoadData.Tables", "Tables présentes dans PQ_DICT : " & ListAllTableNames(wsPQDict), "LoadDictionaryData", "RagicDictionary"
    
    ' Vérifier les colonnes requises
    Dim sheetIdx As Long, fieldIdx As Long, memoIdx As Long
    sheetIdx = GetColumnIndex(wsPQDict, "Sheet")
    fieldIdx = GetColumnIndex(wsPQDict, "Field")
    memoIdx = GetColumnIndex(wsPQDict, "Memo")
    
    If sheetIdx = 0 Or fieldIdx = 0 Or memoIdx = 0 Then
        LogWarning "RagicDict.LoadData.ColumnsNotFound", "Colonnes requises non trouvées dans le tableau Ragic. sheetIdx=" & sheetIdx & ", fieldIdx=" & fieldIdx & ", memoIdx=" & memoIdx
        Exit Sub
    End If
    
    ' Compter les lignes
    Dim nbLignes As Long
    nbLignes = wsPQDict.UsedRange.Rows.Count - 1
    LogInfo "RagicDict.LoadData.RowCount", "Nombre de lignes dans le dictionnaire Ragic source: " & nbLignes, "LoadDictionaryData", "RagicDictionary"
    
    ' Charger les données dans le dictionnaire
    Dim i As Long
    For i = 2 To nbLignes + 1
        RagicFieldDict.Add wsPQDict.Cells(i, sheetIdx).Value & "|" & wsPQDict.Cells(i, fieldIdx).Value, wsPQDict.Cells(i, memoIdx).Value
    Next i
    
    LogInfo "RagicDict.LoadData.DictCount", "Nombre de clés chargées dans RagicFieldDict : " & RagicFieldDict.Count, "LoadDictionaryData", "RagicDictionary"
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Function IsFieldHidden(ByVal tableName As String, ByVal fieldName As String) As Boolean
    Const PROC_NAME As String = "IsFieldHidden"
    On Error GoTo ErrorHandler
    
    ' Vérifier si le dictionnaire est initialisé
    If RagicFieldDict Is Nothing Then
        LogWarning "RagicDict.IsFieldHidden.NotInit", "RagicFieldDict non initialisé lors de l'appel à IsFieldHidden.", "IsFieldHidden", "RagicDictionary"
        IsFieldHidden = False
        Exit Function
    End If
    
    ' Vérifier si le champ est caché
    IsFieldHidden = RagicFieldDict.Exists(tableName & "|" & fieldName)
    
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
    IsFieldHidden = False
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
    Const PROC_NAME As String = "TestIsFieldHidden_BudgetGroupes"
    On Error GoTo ErrorHandler
    
    ' Initialiser le dictionnaire si nécessaire
    If RagicFieldDict Is Nothing Then
        LogInfo "RagicDict.Test.Init", "Dictionnaire non initialisé pour TestIsFieldHidden_BudgetGroupes, chargement...", "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
        LoadRagicDictionary
    End If
    
    ' Tester les champs
    LogInfo "RagicDict.Test.Result1", "Test IsFieldHidden ('↳ Budget Groupes', 'Montant Total'): " & IsFieldHidden("↳ Budget Groupes", "Montant Total"), "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    LogInfo "RagicDict.Test.Result2", "Test IsFieldHidden ('↳ Budget Groupes', 'Année'): " & IsFieldHidden("↳ Budget Groupes", "Année"), "TestIsFieldHidden_BudgetGroupes", "RagicDictionary"
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Private Sub CleanupPowerQuery(ByVal queryName As String)
    On Error Resume Next
    ThisWorkbook.Queries(queryName).Delete
    On Error GoTo 0
End Sub


