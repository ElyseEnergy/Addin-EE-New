Attribute VB_Name = "LoadQueries"
Option Explicit
Sub LoadQuery(queryName As String, ws As Worksheet, DestCell As Range)
    On Error GoTo ErrorHandler
    
    If queryName = "" Then
        SYS_ErrorHandler.HandleError "LoadQueries", "LoadQuery", "Nom de requête vide"
        Exit Sub
    End If
    
    If ws Is Nothing Then
        SYS_ErrorHandler.HandleError "LoadQueries", "LoadQuery", "Feuille de calcul non spécifiée"
        Exit Sub
    End If
    
    If DestCell Is Nothing Then
        SYS_ErrorHandler.HandleError "LoadQueries", "LoadQuery", "Cellule de destination non spécifiée"
        Exit Sub
    End If
    
    Dim lo As ListObject
    Dim sanitizedName As String
    sanitizedName = "Table_" & Utilities.SanitizeTableName(queryName)
    
    ' Log state before
    SYS_Logger.Log "loadquery", "Avant création: QueryExists=" & PQQueryManager.QueryExists(queryName) & ", TableExists=" & TableExists(ws, sanitizedName), DEBUG_LEVEL, "LoadQuery", "LoadQueries"
    
    ' Vérifier si la table existe déjà
    If TableExists(ws, sanitizedName) Then Exit Sub

    With ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & ";Extended Properties=""""", _
        Destination:=DestCell).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & queryName & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .ListObject.DisplayName = sanitizedName
        .Refresh BackgroundQuery:=False
    End With
    
    ' Après le chargement de la requête, s'assurer que le nom est correct
    Set lo = ws.ListObjects(ws.ListObjects.Count) ' Le dernier tableau créé
    If Not lo Is Nothing Then
        lo.Name = sanitizedName
    End If
    ' Log state after
    SYS_Logger.Log "loadquery", "Après création: QueryExists=" & PQQueryManager.QueryExists(queryName) & ", TableExists=" & TableExists(ws, sanitizedName), DEBUG_LEVEL, "LoadQuery", "LoadQueries"
    Exit Sub
    
ErrorHandler:
    SYS_Logger.Log "load_query_error", "Erreur VBA dans LoadQuery. Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, "LoadQuery", "LoadQueries"
    SYS_ErrorHandler.HandleError "LoadQueries", "LoadQuery", "Erreur lors du chargement de la requête: " & queryName
End Sub

Function ChooseMultipleValuesFromListWithAll(idList As Collection, displayList As Collection, prompt As String) As Collection
    Const PROC_NAME As String = "ChooseMultipleValuesFromListWithAll"
    Const MODULE_NAME As String = "LoadQueries"
    On Error GoTo ErrorHandler
    
    If idList Is Nothing Or displayList Is Nothing Then
        SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Listes non initialisées"
        Exit Function
    End If
    
    If idList.Count <> displayList.Count Then
        SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Les listes n'ont pas la même taille"
        Exit Function
    End If
    
    Dim i As Long
    Dim userChoice As String
    Dim selectedIndexes As Variant
    Dim SelectedValues As New Collection
    Dim listPrompt As String
    Dim idx As Long

    listPrompt = prompt & vbCrLf & "* : Toutes" & vbCrLf
    For i = 1 To displayList.Count
        listPrompt = listPrompt & i & ". " & displayList(i) & vbCrLf
    Next i
    
    userChoice = InputBox(listPrompt, "Sélection", "1")
    If StrPtr(userChoice) = 0 Or Len(Trim(userChoice)) = 0 Then
        Exit Function
    End If
    
    If Trim(userChoice) = "*" Then
        For i = 1 To idList.Count
            SelectedValues.Add idList(i)
        Next i
    Else
        selectedIndexes = Split(userChoice, ",")
        For i = LBound(selectedIndexes) To UBound(selectedIndexes)
            idx = val(Trim(selectedIndexes(i)))
            If idx >= 1 And idx <= idList.Count Then
                SelectedValues.Add idList(idx)
            End If
        Next i
    End If
    Set ChooseMultipleValuesFromListWithAll = SelectedValues
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "query_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la sélection des valeurs"
End Function

Function ChooseMultipleValuesFromArrayWithAll(values() As String, prompt As String) As Collection
    Const PROC_NAME As String = "ChooseMultipleValuesFromArrayWithAll"
    Const MODULE_NAME As String = "LoadQueries"
    On Error GoTo ErrorHandler
    
    If Not IsArray(values) Then
        SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Tableau non initialisé"
        Exit Function
    End If
    
    If UBound(values) < 1 Then
        SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Tableau vide"
        Exit Function
    End If
    
    Dim i As Long
    Dim userChoice As String
    Dim listPrompt As String
    Dim idx As Long
    
    ' Construire la liste pour l'InputBox
    listPrompt = prompt & vbCrLf & "* : Toutes" & vbCrLf
    For i = 1 To UBound(values)
        listPrompt = listPrompt & i & ". " & values(i) & vbCrLf
    Next i

    userChoice = InputBox(listPrompt, "Sélection", "1")
    If StrPtr(userChoice) = 0 Or Len(Trim(userChoice)) = 0 Then
        Exit Function
    End If
    
    Dim SelectedValues As New Collection
    userChoice = Trim(userChoice)
    
    ' Cas spécial : sélection de toutes les valeurs avec *
    If userChoice = "*" Then
        For i = 1 To UBound(values)
            SelectedValues.Add values(i)
        Next i
    Else
        ' Sélection de valeurs spécifiques
        Dim selectedIndexes As Variant
        selectedIndexes = Split(userChoice, ",")
        Dim hasValidSelection As Boolean
        hasValidSelection = False
        
        For i = LBound(selectedIndexes) To UBound(selectedIndexes)
            idx = val(Trim(selectedIndexes(i)))
            If idx >= 1 And idx <= UBound(values) Then
                SelectedValues.Add values(idx)
                hasValidSelection = True
            End If
        Next i
        
        ' Si aucune sélection valide n'a été trouvée
        If Not hasValidSelection Then
            SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Aucune sélection valide"
            Exit Function
        End If
    End If
    
    Set ChooseMultipleValuesFromArrayWithAll = SelectedValues
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "query_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la sélection des valeurs"
End Function

' Helper function for table existence
Private Function TableExists(ws As Worksheet, tableName As String) As Boolean
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    TableExists = Not lo Is Nothing
End Function


