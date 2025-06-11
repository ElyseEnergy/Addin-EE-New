Attribute VB_Name = "LoadQueries"
Sub LoadQuery(QueryName As String, ws As Worksheet, destCell As Range)
    On Error GoTo ErrorHandler
    
    If QueryName = "" Then
        HandleError "LoadQueries", "LoadQuery", "Nom de requ�te vide"
        Exit Sub
    End If
    
    If ws Is Nothing Then
        HandleError "LoadQueries", "LoadQuery", "Feuille de calcul non sp�cifi�e"
        Exit Sub
    End If
    
    If destCell Is Nothing Then
        HandleError "LoadQueries", "LoadQuery", "Cellule de destination non sp�cifi�e"
        Exit Sub
    End If
    
    Dim lo As ListObject
    Dim sanitizedName As String
    sanitizedName = "Table_" & Utilities.SanitizeTableName(QueryName)
    
    ' Log state before
    Log "loadquery", "Avant cr�ation: QueryExists=" & PQQueryManager.QueryExists(QueryName) & ", TableExists=" & tableExists(ws, sanitizedName), DEBUG_LEVEL, "LoadQuery", "LoadQueries"
    
    ' V�rifier si la table existe d�j�
    For Each lo In ws.ListObjects
        If lo.Name = sanitizedName Then
            Exit Sub
        End If
    Next lo

    With ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & QueryName & ";Extended Properties=""""", _
        Destination:=destCell).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & QueryName & "]")
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
    
    ' Apr�s le chargement de la requ�te, s'assurer que le nom est correct
    Set lo = ws.ListObjects(ws.ListObjects.count) ' Le dernier tableau cr��
    If Not lo Is Nothing Then
        lo.Name = sanitizedName
    End If
    ' Log state after
    Log "loadquery", "Apr�s cr�ation: QueryExists=" & PQQueryManager.QueryExists(QueryName) & ", TableExists=" & tableExists(ws, sanitizedName), DEBUG_LEVEL, "LoadQuery", "LoadQueries"
    Exit Sub
    
ErrorHandler:
    Log "loadquery", "ERREUR: " & Err.Description, ERROR_LEVEL, "LoadQuery", "LoadQueries"
    Log "loadquery", "Diagnostics: QueryExists=" & PQQueryManager.QueryExists(QueryName) & ", TableExists=" & tableExists(ws, sanitizedName), ERROR_LEVEL, "LoadQuery", "LoadQueries"
    HandleError "LoadQueries", "LoadQuery", "Erreur lors du chargement de la requ�te: " & QueryName
End Sub

Function ChooseMultipleValuesFromListWithAll(idList As Collection, displayList As Collection, prompt As String) As Collection
    On Error GoTo ErrorHandler
    
    If idList Is Nothing Or displayList Is Nothing Then
        HandleError "LoadQueries", "ChooseMultipleValuesFromListWithAll", "Listes non initialis�es"
        Exit Function
    End If
    
    If idList.count <> displayList.count Then
        HandleError "LoadQueries", "ChooseMultipleValuesFromListWithAll", "Les listes n'ont pas la m�me taille"
        Exit Function
    End If
    
    Dim i As Long
    Dim userChoice As String
    Dim selectedIndexes As Variant
    Dim selectedValues As New Collection
    Dim listPrompt As String

    listPrompt = prompt & vbCrLf & "* : Toutes" & vbCrLf
    For i = 1 To displayList.count
        listPrompt = listPrompt & i & ". " & displayList(i) & vbCrLf
    Next i
    
    userChoice = InputBox(listPrompt, "S�lection", "1")
    If StrPtr(userChoice) = 0 Or Len(Trim(userChoice)) = 0 Then
        Exit Function
    End If
    
    If Trim(userChoice) = "*" Then
        For i = 1 To idList.count
            selectedValues.Add idList(i)
        Next i
    Else
        selectedIndexes = Split(userChoice, ",")
        For i = LBound(selectedIndexes) To UBound(selectedIndexes)
            Dim idx As Long
            idx = val(Trim(selectedIndexes(i)))
            If idx >= 1 And idx <= idList.count Then
                selectedValues.Add idList(idx)
            End If
        Next i
    End If
    Set ChooseMultipleValuesFromListWithAll = selectedValues
    Exit Function
    
ErrorHandler:
    HandleError "LoadQueries", "ChooseMultipleValuesFromListWithAll", "Erreur lors de la s�lection des valeurs"
End Function

Function ChooseMultipleValuesFromArrayWithAll(values() As String, prompt As String) As Collection
    On Error GoTo ErrorHandler
    
    If Not IsArray(values) Then
        HandleError "LoadQueries", "ChooseMultipleValuesFromArrayWithAll", "Tableau non initialis�"
        Exit Function
    End If
    
    If UBound(values) < 1 Then
        HandleError "LoadQueries", "ChooseMultipleValuesFromArrayWithAll", "Tableau vide"
        Exit Function
    End If
    
    Dim i As Long
    Dim userChoice As String
    Dim listPrompt As String
    
    ' Construire la liste pour l'InputBox
    listPrompt = prompt & vbCrLf & "* : Toutes" & vbCrLf
    For i = 1 To UBound(values)
        listPrompt = listPrompt & i & ". " & values(i) & vbCrLf
    Next i

    userChoice = InputBox(listPrompt, "S�lection", "1")
    If StrPtr(userChoice) = 0 Or Len(Trim(userChoice)) = 0 Then
        Exit Function
    End If
    
    Dim selectedValues As New Collection
    userChoice = Trim(userChoice)
    
    ' Cas sp�cial : s�lection de toutes les valeurs avec *
    If userChoice = "*" Then
        For i = 1 To UBound(values)
            selectedValues.Add values(i)
        Next i
    Else
        ' S�lection de valeurs sp�cifiques
        Dim selectedIndexes As Variant
        selectedIndexes = Split(userChoice, ",")
        Dim hasValidSelection As Boolean
        hasValidSelection = False
        
        For i = LBound(selectedIndexes) To UBound(selectedIndexes)
            Dim idx As Long
            idx = val(Trim(selectedIndexes(i)))
            If idx >= 1 And idx <= UBound(values) Then
                selectedValues.Add values(idx)
                hasValidSelection = True
            End If
        Next i
        
        ' Si aucune s�lection valide n'a �t� trouv�e
        If Not hasValidSelection Then
            HandleError "LoadQueries", "ChooseMultipleValuesFromArrayWithAll", "Aucune s�lection valide"
            Exit Function
        End If
    End If
    
    Set ChooseMultipleValuesFromArrayWithAll = selectedValues
    Exit Function
    
ErrorHandler:
    HandleError "LoadQueries", "ChooseMultipleValuesFromArrayWithAll", "Erreur lors de la s�lection des valeurs"
End Function

' Helper function for table existence
Private Function tableExists(ws As Worksheet, tableName As String) As Boolean
    On Error Resume Next
    tableExists = Not ws.ListObjects(tableName) Is Nothing
    On Error GoTo 0
End Function


