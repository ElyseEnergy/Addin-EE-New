Attribute VB_Name = "RagicDictionary"

Option Explicit

' ==================================================================================================
' Module: RagicDictionary
' Auteur: 
' Date: 2024-03-01
' Description: Ce module gère la récupération, le cache et l'utilisation du dictionnaire de données
'              provenant de Ragic. Ce dictionnaire fournit des métadonnées sur les champs,
'              comme le type de données ou si un champ doit être masqué.
' ==================================================================================================

' --- CONSTANTES PUBLIQUES ---

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
Public Sub ProcessForceRefreshRagicDictionary(control As IRibbonControl)
    Const PROC_NAME As String = "ProcessForceRefreshRagicDictionary"
    On Error GoTo ErrorHandler
    ForceRefreshRagicDictionary
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & "RagicDictionary" & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, "RagicDictionary"
    SYS_ErrorHandler.HandleError "RagicDictionary", PROC_NAME, "Erreur lors du rafraîchissement du dictionnaire (ProcessForceRefreshRagicDictionary)"
End Sub

' Callback pour l'info-bulle (supertip) du bouton de rafraîchissement
Public Sub GetRagicDictSupertip(ByVal control As IRibbonControl, ByRef supertip As Variant)
    Const PROC_NAME As String = "GetRagicDictSupertip"
    On Error GoTo ErrorHandler
    Dim lastUpdate As Date
    lastUpdate = GetLastRefreshDate() ' On réutilise la fonction existante
    
    Dim lastUpdateText As String
    If lastUpdate > 0 Then
        lastUpdateText = "Last update: " & Format(lastUpdate, "yyyy-mm-dd")
    Else
        lastUpdateText = "Never updated. Click to download."
    End If
    
    supertip = "Downloads the latest version of the data dictionary from Ragic." & vbCrLf & vbCrLf & lastUpdateText
    Exit Sub
ErrorHandler:
    supertip = "Error getting refresh date."
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & "RagicDictionary" & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, "RagicDictionary"
    SYS_ErrorHandler.HandleError "RagicDictionary", PROC_NAME, "Erreur lors de la récupération de la date de rafraîchissement (GetRagicDictSupertip)"
End Sub

'==================================================================================================
' MÉTHODES PUBLIQUES
'==================================================================================================

' Force le rafraîchissement du dictionnaire depuis Ragic
Public Sub ForceRefreshRagicDictionary()
    Const PROC_NAME As String = "ForceRefreshRagicDictionary"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler

    Application.StatusBar = "Forçage du rafraîchissement du dictionnaire Ragic..."
    ' Réinitialiser la date dans les propriétés pour forcer le rechargement
    SetLastRefreshDate (0)
    ' Appeler la routine de chargement
    LoadRagicDictionary
    Application.StatusBar = False
    MsgBox "Le dictionnaire Ragic a été mis à jour.", vbInformation
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Failed to force refresh Ragic dictionary. Some features may not work."
End Sub

' Charge le dictionnaire Ragic, depuis le cache si possible
Public Sub LoadRagicDictionary()
    Const PROC_NAME As String = "LoadRagicDictionary"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler

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
    tableExists = TableExists(wsPQDict, tableName) ' Utilise la fonction helper TableExists
    SYS_Logger.Log "load_dict", "Table " & tableName & " existe (via helper): " & tableExists, DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"
    
    ' Décider s'il faut rafraîchir depuis le réseau
    Dim needsRefresh As Boolean
    needsRefresh = Not tableExists Or (VBA.Date - GetLastRefreshDate() >= 1) ' Ensure GetLastRefreshDate is called
    SYS_Logger.Log "load_dict", "Dernière MàJ: " & lastRefresh & ", Âge (heures): " & ((Now - lastRefresh) * 24) & "h", DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"
    SYS_Logger.Log "load_dict", "Rafraîchissement nécessaire: " & needsRefresh & " (table existe: " & tableExists & ")", DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"

    If needsRefresh Then
        Application.StatusBar = "Chargement du dictionnaire Ragic depuis le réseau..."
        SYS_Logger.Log "load_dict", "Rafraîchissement du dictionnaire Ragic depuis le réseau.", INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"

        ' Créer une catégorie pour le dictionnaire
        Dim dictCategory As CategoryInfo
        With dictCategory
            .CategoryName = BASE_NAME
            .DisplayName = BASE_NAME
            .URL = env.RAGIC_BASE_URL & RAGIC_PATH & env.GetRagicApiParams()
            .PowerQueryName = pqName
            .SheetName = BASE_NAME
        End With

        SYS_Logger.Log "load_dict", "URL de requête dictionnaire : " & dictCategory.URL, INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"        ' Créer ou mettre à jour la requête
        Dim queryFormula As String
        queryFormula = GenerateDictionaryQuery(dictCategory.URL)
        If queryFormula = "" Then
            SYS_Logger.Log "ragic_dict_err", "Échec de la génération de la requête Power Query", ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
            Application.StatusBar = False
            Exit Sub
        End If
        
        If PQQueryManager.QueryExists(pqName) Then
            ' Mise à jour de la requête existante
            On Error GoTo UpdateError
            ThisWorkbook.Queries(pqName).formula = queryFormula
        Else
            ' Création d'une nouvelle requête
            On Error GoTo CreateError
            ThisWorkbook.Queries.Add pqName, queryFormula
        End If
        On Error GoTo 0
        GoTo Continue ' Skip error handlers
        
UpdateError:
        SYS_Logger.Log "ragic_dict_err", "Erreur MàJ requête " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
        Application.StatusBar = False
        Exit Sub
        
CreateError:
        SYS_Logger.Log "ragic_dict_err", "Erreur ajout requête " & pqName & ": " & Err.Description, ERROR_LEVEL, "LoadRagicDictionary", "RagicDictionary"
        Application.StatusBar = False
        Exit Sub
        
Continue:' Rafraîchir la requête (c'est l'étape lente)
        If Not RefreshPowerQuery(pqName) Then
            Application.StatusBar = False
            Exit Sub
        End If

        ' Charger les données dans la feuille PQ_DICT
        If Not PQQueryManager.QueryExists(pqName) Then
            LoadQueries.LoadQuery pqName, wsPQDict, wsPQDict.Range("A1")
        End If

        ' Mettre à jour la date du rafraîchissement dans les propriétés du classeur
        SetLastRefreshDate VBA.Date
        
        ' Proposer la sauvegarde du classeur à l'utilisateur
        If Not SaveWorkbookSafely() Then
            SYS_Logger.Log "load_dict", "Attention: Le classeur n'a pas été sauvegardé. Les modifications pourraient être perdues.", WARNING_LEVEL, "LoadRagicDictionary", "RagicDictionary"
        End If
        
        ' Invalider le contrôle du ruban pour mettre à jour l'info-bulle
        If Not RibbonVisibility.gRibbon Is Nothing Then
            RibbonVisibility.gRibbon.InvalidateControl "btnForceRefreshRagic"
        End If
    Else
        SYS_Logger.Log "load_dict", "Chargement du dictionnaire Ragic depuis le cache local (feuille PQ_DICT).", INFO_LEVEL, "LoadRagicDictionary", "RagicDictionary"
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
    SYS_Logger.Log "load_dict", "Premières clés du dictionnaire : " & debugKeys, DEBUG_LEVEL, "LoadRagicDictionary", "RagicDictionary"

    ' La feuille PQ_DICT est rendue visible par GetOrCreatePQDictSheet si nécessaire.
    ' Il n'est pas utile de la forcer visible ici à nouveau.
    ' If Not wsPQDict Is Nothing Then
    ' wsPQDict.Visible = xlSheetVisible
    ' End If

    ' Réinitialiser la barre de statut
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Critical error while loading Ragic Dictionary. Some features may not work."
End Sub

' Recherche la meilleure ligne pour un fieldName donné (et SheetName si doublons)
Public Function FindBestRowForField(lo As ListObject, SheetName As String, fieldName As String) As Long
    Const PROC_NAME As String = "FindBestRowForField"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler
    
    SYS_Logger.Log "FindBestRowForField", "Entrée: SheetName='" & SheetName & "', fieldName='" & fieldName & "'", DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary"
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
    SYS_Logger.Log "FindBestRowForField", "Nb matches sur FieldName: " & matchCount, DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary"
    If matchCount = 0 Then SYS_Logger.Log "FindBestRowForField", "Aucun match trouvé.", DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = 0: Exit Function
    If matchCount = 1 Then SYS_Logger.Log "FindBestRowForField", "Un seul match, row=" & matches(1), DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = matches(1): Exit Function
    Dim j As Long
    For j = 1 To matchCount
        If StrComp(Trim(arr(matches(j), colSheetName)), Trim(SheetName), vbTextCompare) = 0 Then SYS_Logger.Log "FindBestRowForField", "Match exact SheetName à row=" & matches(j), DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = matches(j): Exit Function
    Next j
    For j = 1 To matchCount
        If InStr(1, arr(matches(j), colSheetName), SheetName, vbTextCompare) > 0 Then SYS_Logger.Log "FindBestRowForField", "Match contains SheetName à row=" & matches(j), DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = matches(j): Exit Function
    Next j
    For j = 1 To matchCount
        If Left(arr(matches(j), colSheetName), Len(SheetName)) = SheetName Then SYS_Logger.Log "FindBestRowForField", "Match startswith SheetName à row=" & matches(j), DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = matches(j): Exit Function
    Next j
    SYS_Logger.Log "FindBestRowForField", "Fallback premier match à row=" & matches(1), DEBUG_LEVEL, "FindBestRowForField", "RagicDictionary": FindBestRowForField = matches(1)
    Exit Function

ErrorHandler:
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la recherche de la meilleure ligne pour le champ " & fieldName
    FindBestRowForField = 0 ' Retourner 0 en cas d'erreur
End Function

' Retourne la valeur d'une colonne pour une ligne donnée
Public Function GetValueFromRow(lo As ListObject, arr As Variant, rowIndex As Long, colName As String) As Variant
    Const PROC_NAME As String = "GetValueFromRow"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler

    SYS_Logger.Log "GetValueFromRow", "rowIndex=" & rowIndex & ", colName='" & colName & "'", DEBUG_LEVEL, "GetValueFromRow", "RagicDictionary"
    Dim colIdx As Long, i As Long
    For i = 1 To lo.ListColumns.Count
        If lo.ListColumns(i).Name = colName Then colIdx = i: Exit For
    Next i
    If colIdx = 0 Or rowIndex = 0 Then        SYS_Logger.Log "GetValueFromRow", "Colonne ou ligne non trouvée.", DEBUG_LEVEL, "GetValueFromRow", "RagicDictionary"
        GetValueFromRow = ""
    Else
        GetValueFromRow = arr(rowIndex, colIdx)
        SYS_Logger.Log "GetValueFromRow", "Valeur extraite: '" & GetValueFromRow & "'", DEBUG_LEVEL, "GetValueFromRow", "RagicDictionary"
    End If
    Exit Function

ErrorHandler:
    GetValueFromRow = CVErr(xlErrValue)
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Failed to get value from row for col '" & colName & "'."
End Function

' Fonction principale Hidden
Public Function IsFieldHidden(SheetName As String, fieldName As String) As Boolean
    Const PROC_NAME As String = "IsFieldHidden"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler    SYS_Logger.Log "IsFieldHidden", "Entrée: SheetName='" & SheetName & "', fieldName='" & fieldName & "'", DEBUG_LEVEL, "IsFieldHidden", "RagicDictionary"
    IsFieldHidden = False
    
    ' Vérification et initialisation de la feuille PQ_DICT
    If wsPQDict Is Nothing Then 
        Set wsPQDict = GetOrCreatePQDictSheet()
        If wsPQDict Is Nothing Then
            SYS_Logger.Log "IsFieldHidden", "Impossible d'accéder à la feuille PQ_DICT", ERROR_LEVEL, "IsFieldHidden", "RagicDictionary"
            Exit Function
        End If
    End If

    ' Vérification et accès à la table de données
    Dim lo As ListObject
    Dim tableName As String
    tableName = "Table_PQ_" & Utilities.SanitizeTableName(BASE_NAME)
    
    If Not TableExists(wsPQDict, tableName) Then
        SYS_Logger.Log "IsFieldHidden", "Table '" & tableName & "' non trouvée. Tables disponibles: " & TableManager.ListAllTableNames(wsPQDict), ERROR_LEVEL, "IsFieldHidden", "RagicDictionary"
        Exit Function
    End If
    
    Set lo = wsPQDict.ListObjects(tableName)If lo Is Nothing Or lo.ListRows.Count = 0 Then SYS_Logger.Log "IsFieldHidden", "Table non trouvée ou vide.", DEBUG_LEVEL, "IsFieldHidden", "RagicDictionary": Exit Function
    Dim arr As Variant
    arr = lo.DataBodyRange.Value
    Dim rowIdx As Long
    rowIdx = FindBestRowForField(lo, SheetName, fieldName)
    If rowIdx = 0 Then SYS_Logger.Log "IsFieldHidden", "Aucune ligne trouvée.", DEBUG_LEVEL, "IsFieldHidden", "RagicDictionary": Exit Function
    Dim memoVal As String
    memoVal = GetValueFromRow(lo, arr, rowIdx, "Memo")
    If InStr(1, memoVal, "Hidden", vbTextCompare) > 0 Then
        SYS_Logger.Log "IsFieldHidden", "Champ HIDDEN détecté.", DEBUG_LEVEL, "IsFieldHidden", "RagicDictionary"
        IsFieldHidden = True
    Else
        SYS_Logger.Log "IsFieldHidden", "Champ NON hidden.", DEBUG_LEVEL, "IsFieldHidden", "RagicDictionary"
    End If

    If lo.ListRows.Count = 0 Then
        SYS_Logger.Log "IsFieldHidden", "ListObject '" & tableName & "' is empty.", WARNING_LEVEL, "IsFieldHidden", "RagicDictionary"
        Exit Function

ErrorHandler:
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & "RagicDictionary" & "." & "IsFieldHidden" & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, "IsFieldHidden", "RagicDictionary"
    SYS_ErrorHandler.HandleError "RagicDictionary", "IsFieldHidden", "Erreur lors de la vérification si le champ " & fieldName & " est masqué."
    IsFieldHidden = False ' Par défaut, ne pas masquer en cas d'erreur
End Function

' Normalise le nom de la feuille pour la clé dictionnaire
Public Function NormalizeSheetName(SheetName As String) As String
    Const PROC_NAME As String = "NormalizeSheetName"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler

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
    Exit Function

ErrorHandler:
    NormalizeSheetName = "ErrorNormalizingName"
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Failed to normalize sheet name '" & SheetName & "'."
End Function

'==================================================================================================
' MÉTHODES PRIVÉES
'==================================================================================================

' Gère la persistance de la date de rafraîchissement via les propriétés du document
Private Function GetLastRefreshDate() As Date
    Const PROC_NAME As String = "GetLastRefreshDate"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler
    
    Dim propValue As Variant
    
    Err.Clear ' Clear any prior error
    On Error Resume Next ' Temporarily resume next to check for property existence
    propValue = ThisWorkbook.CustomDocumentProperties(PROP_LAST_REFRESH).Value
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler ' Restore error handling
        SYS_Logger.Log "GetLastRefreshDate", "CustomProperty '" & PROP_LAST_REFRESH & "' not found or error accessing it. Error: " & Err.Description, WARNING_LEVEL, PROC_NAME, MODULE_NAME
        GetLastRefreshDate = 0 ' Force le rafraîchissement si la propriété n'existe pas ou erreur
    Else
        On Error GoTo ErrorHandler ' Restore error handling
        If IsDate(propValue) Then
            GetLastRefreshDate = CDate(propValue)
        Else
            SYS_Logger.Log "GetLastRefreshDate", "CustomProperty '" & PROP_LAST_REFRESH & "' does not contain a valid date. Value: " & CStr(propValue), WARNING_LEVEL, PROC_NAME, MODULE_NAME
            GetLastRefreshDate = 0 ' Force refresh if not a valid date
        End If
    End If
    Exit Function

ErrorHandler:
    GetLastRefreshDate = 0 ' Retourne 0 (équivalent à une date vide) en cas d'erreur
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Impossible de récupérer la date de dernier rafraîchissement."
End Function

Private Sub SetLastRefreshDate(d As Date)
    Const PROC_NAME As String = "SetLastRefreshDate"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler
    
    Dim prop As Object ' DocumentProperty
    
    Err.Clear
    On Error Resume Next ' Temporarily resume next to check for property existence
    Set prop = ThisWorkbook.CustomDocumentProperties(PROP_LAST_REFRESH)
    
    If Err.Number = 0 And Not prop Is Nothing Then
        On Error GoTo ErrorHandler ' Restore error handling
        ' La propriété existe, on met juste à jour la valeur
        prop.Value = d
    Else
        ' La propriété n'existe pas ou une erreur s'est produite, on l'ajoute
        Err.Clear ' Clear error from trying to Set prop
        On Error GoTo ErrorHandler ' Restore error handling before Add
        ThisWorkbook.CustomDocumentProperties.Add _
            Name:=PROP_LAST_REFRESH, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeDate, _
            Value:=d
    End If
    Exit Sub

ErrorHandler:
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Impossible de définir la date de dernier rafraîchissement."
End Sub

' Génère la requête PowerQuery spécifique pour le dictionnaire
Private Function GenerateDictionaryQuery(ByVal URL As String) As String
    Const PROC_NAME As String = "GenerateDictionaryQuery"
    On Error GoTo ErrorHandler
    Dim template As String
    template = "let" & vbCrLf & _
          "    Source = Csv.Document(Web.Contents(""" & URL & """),[Delimiter="","",Encoding=65001,QuoteStyle=QuoteStyle.Csv])," & vbCrLf & _
          "    PromotedHeaders = Table.PromoteHeaders(Source)" & vbCrLf & _
          "in" & vbCrLf & _
          "    PromotedHeaders"
    GenerateDictionaryQuery = template
    Exit Function
ErrorHandler:
    SYS_Logger.Log "error", "Erreur lors de la génération de la requête dictionnaire: " & Err.Description, ERROR_LEVEL, PROC_NAME, "RagicDictionary"
    SYS_ErrorHandler.HandleError "RagicDictionary", PROC_NAME
    GenerateDictionaryQuery = ""
End Function

Private Function GetOrCreatePQDictSheet() As Worksheet
    Const PROC_NAME As String = "GetOrCreatePQDictSheet"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler
    
    Dim tempSheet As Worksheet
    
    Err.Clear
    On Error Resume Next ' Temporarily resume next to check for sheet existence
    Set tempSheet = ThisWorkbook.Worksheets("PQ_DICT")
    
    If Err.Number <> 0 Or tempSheet Is Nothing Then
        On Error GoTo ErrorHandler ' Restore error handling
        ' Créer la feuille si elle n'existe pas
        SYS_Logger.Log "load_dict", "Création de la feuille PQ_DICT...", INFO_LEVEL, "GetOrCreatePQDictSheet", "RagicDictionary"
        Set tempSheet = ThisWorkbook.Worksheets.Add
        tempSheet.Name = "PQ_DICT"
    Else
        On Error GoTo ErrorHandler ' Restore error handling
    End If
    
    Set GetOrCreatePQDictSheet = tempSheet
    
    ' Force la visibilité de la feuille dans tous les cas
    If GetOrCreatePQDictSheet.Visible <> xlSheetVisible Then
        SYS_Logger.Log "load_dict", "Rendre la feuille PQ_DICT visible...", INFO_LEVEL, "GetOrCreatePQDictSheet", "RagicDictionary"
        GetOrCreatePQDictSheet.Visible = xlSheetVisible
    End If
    
    Exit Function
ErrorHandler:
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Impossible de créer ou de trouver la feuille PQ_DICT."
    Set GetOrCreatePQDictSheet = Nothing
End Function

Private Sub LoadDictionaryData(ByVal tableName As String)
    Const PROC_NAME As String = "LoadDictionaryData"
    Const MODULE_NAME As String = "RagicDictionary"
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
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur lors du chargement des données depuis la table " & tableName & " dans le dictionnaire."
End Sub

' --- Fonctions de Test ---
'---------------------------------------------------------------------------------------
Public Sub TestIsFieldHiddenBudgetGroupes()
    Const PROC_NAME As String = "TestIsFieldHiddenBudgetGroupes"
    On Error GoTo ErrorHandler
    Dim result As Boolean
    result = IsFieldHidden("Budget - Groupes", "ID")
    Debug.Print "Is 'ID' hidden in 'Budget - Groupes'? " & result
    Exit Sub
ErrorHandler:
    SYS_ErrorHandler.HandleError "RagicDictionary", PROC_NAME
End Sub

' Recherche à deux étages pour le type de champ
Public Function GetFieldRagicType(categorySheetName As String, fieldName As String) As String
    Const PROC_NAME As String = "GetFieldRagicType"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler
    
    SYS_Logger.Log "GetFieldRagicType", "Entrée: SheetName='" & categorySheetName & "', fieldName='" & fieldName & "'", DEBUG_LEVEL, "GetFieldRagicType", "RagicDictionary"
    GetFieldRagicType = "Text" ' Default value    ' Vérification et initialisation de la feuille PQ_DICT
    If wsPQDict Is Nothing Then 
        SYS_Logger.Log "GetFieldRagicType", "wsPQDict is Nothing. Tentative de GetOrCreatePQDictSheet.", WARNING_LEVEL, PROC_NAME, MODULE_NAME
        Set wsPQDict = GetOrCreatePQDictSheet()
        If wsPQDict Is Nothing Then
            SYS_Logger.Log "GetFieldRagicType", "Impossible de créer ou d'accéder à PQ_DICT.", ERROR_LEVEL, PROC_NAME, MODULE_NAME
            Exit Function
        End If
    End If

    ' Vérification et accès à la table de données
    Dim lo As ListObject
    Dim tableName As String
    tableName = "Table_PQ_" & Utilities.SanitizeTableName(BASE_NAME)
    
    If Not TableExists(wsPQDict, tableName) Then
        SYS_Logger.Log "GetFieldRagicType", "Table '" & tableName & "' non trouvée. Tables disponibles: " & TableManager.ListAllTableNames(wsPQDict), ERROR_LEVEL, PROC_NAME, MODULE_NAME
        Exit Function
    End If
    
    Set lo = wsPQDict.ListObjects(tableName)If lo Is Nothing Or lo.ListRows.Count = 0 Then SYS_Logger.Log "GetFieldRagicType", "Table non trouvée ou vide.", DEBUG_LEVEL, "GetFieldRagicType", "RagicDictionary": Exit Function
    Dim arr As Variant
    arr = lo.DataBodyRange.Value
    Dim rowIdx As Long
    rowIdx = FindBestRowForField(lo, categorySheetName, fieldName)
    If rowIdx = 0 Then SYS_Logger.Log "GetFieldRagicType", "Aucune ligne trouvée.", DEBUG_LEVEL, "GetFieldRagicType", "RagicDictionary": Exit Function
    Dim fieldTypeVal As String
    fieldTypeVal = GetValueFromRow(lo, arr, rowIdx, "Field Type")
    If Trim(fieldTypeVal) <> "" Then
        SYS_Logger.Log "GetFieldRagicType", "Type trouvé: '" & fieldTypeVal & "'", DEBUG_LEVEL, "GetFieldRagicType", "RagicDictionary"
        GetFieldRagicType = Trim(fieldTypeVal)
    Else
        SYS_Logger.Log "GetFieldRagicType", "Type vide, fallback 'Text'", DEBUG_LEVEL, "GetFieldRagicType", "RagicDictionary"
    End If

    If lo.ListRows.Count = 0 Then
        SYS_Logger.Log "GetFieldRagicType", "ListObject '" & tableName & "' is empty.", WARNING_LEVEL, "GetFieldRagicType", "RagicDictionary"
        Exit Function

ErrorHandler:
    GetFieldRagicType = "Text" ' Fallback de sécurité
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Failed to get Ragic field type for '" & fieldName & "' on sheet '" & categorySheetName & "'."
End Function

Private Function RefreshPowerQuery(ByVal queryName As String) As Boolean
    Const PROC_NAME As String = "RefreshPowerQuery"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler

    ' Si la requête n'existe pas, échec immédiat
    If Not PQQueryManager.QueryExists(queryName) Then
        SYS_Logger.Log "ragic_dict_err", "La requête " & queryName & " n'existe pas.", ERROR_LEVEL, PROC_NAME, MODULE_NAME
        RefreshPowerQuery = False
        Exit Function
    End If

    ' Tenter de rafraîchir la requête
    Application.StatusBar = "Rafraîchissement de la requête " & queryName & "..."

    ThisWorkbook.Queries(queryName).Refresh
    RefreshPowerQuery = True
    Exit Function

ErrorHandler:
    SYS_Logger.Log "ragic_dict_err", "Erreur lors du rafraîchissement de la requête " & queryName & ": " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    Application.StatusBar = False
    RefreshPowerQuery = False
End Function

Private Function SaveWorkbookSafely() As Boolean
    Const PROC_NAME As String = "SaveWorkbookSafely"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler

    ' Demander la confirmation à l'utilisateur
    If MsgBox("Le dictionnaire a été mis à jour. Souhaitez-vous sauvegarder le classeur maintenant ?", vbQuestion + vbYesNo) = vbYes Then
        ThisWorkbook.Save
        SYS_Logger.Log "load_dict", "Classeur sauvegardé avec succès.", INFO_LEVEL, PROC_NAME, MODULE_NAME
        SaveWorkbookSafely = True
    Else
        SYS_Logger.Log "load_dict", "Sauvegarde du classeur annulée par l'utilisateur.", INFO_LEVEL, PROC_NAME, MODULE_NAME
        SaveWorkbookSafely = False
    End If
    Exit Function

ErrorHandler:
    SYS_Logger.Log "ragic_dict_err", "ERREUR lors de la sauvegarde du classeur: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SaveWorkbookSafely = False
End Function

' Vérifie l'existence d'une table dans une feuille de manière sécurisée
Private Function TableExists(ws As Worksheet, tableName As String) As Boolean
    Const PROC_NAME As String = "TableExists"
    Const MODULE_NAME As String = "RagicDictionary"
    On Error GoTo ErrorHandler

    If ws Is Nothing Then 
        TableExists = False
        Exit Function
    End If

    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    Dim errNumber As Long
    errNumber = Err.Number
    On Error GoTo 0

    TableExists = (errNumber = 0 And Not lo Is Nothing)
    Exit Function

ErrorHandler:
    TableExists = False
    SYS_Logger.Log "ragic_dict_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
End Function