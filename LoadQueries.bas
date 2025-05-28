Sub LoadQuery(QueryName As String, ws As Worksheet, DestCell As Range)
    On Error GoTo ErrorHandler
    
    If QueryName = "" Then
        HandleError "LoadQueries", "LoadQuery", "Nom de requête vide"
        Exit Sub
    End If
    
    If ws Is Nothing Then
        HandleError "LoadQueries", "LoadQuery", "Feuille de calcul non spécifiée"
        Exit Sub
    End If
    
    If DestCell Is Nothing Then
        HandleError "LoadQueries", "LoadQuery", "Cellule de destination non spécifiée"
        Exit Sub
    End If
    
    Dim lo As ListObject
    Dim sanitizedName As String
    
    ' Nettoyer le nom de la requête pour le nom de tableau
    sanitizedName = "Table_" & Utilities.SanitizeTableName(QueryName)
    
    ' Vérifier si la table existe déjà
    For Each lo In ws.ListObjects
        If lo.Name = sanitizedName Then
            Exit Sub
        End If
    Next lo

    With ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & QueryName & ";Extended Properties=""""", _
        Destination:=DestCell).QueryTable
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
    
    ' Après le chargement de la requête, s'assurer que le nom est correct
    Set lo = ws.ListObjects(ws.ListObjects.Count) ' Le dernier tableau créé
    If Not lo Is Nothing Then
        lo.Name = sanitizedName
    End If
    Exit Sub
    
ErrorHandler:
    HandleError "LoadQueries", "LoadQuery", "Erreur lors du chargement de la requête: " & QueryName
End Sub

Function ChooseMultipleValuesFromListWithAll(idList As Collection, displayList As Collection, prompt As String) As Collection
    On Error GoTo ErrorHandler
    
    If idList Is Nothing Or displayList Is Nothing Then
        HandleError "LoadQueries", "ChooseMultipleValuesFromListWithAll", "Listes non initialisées"
        Exit Function
    End If
    
    If idList.Count <> displayList.Count Then
        HandleError "LoadQueries", "ChooseMultipleValuesFromListWithAll", "Les listes n'ont pas la même taille"
        Exit Function
    End If
    
    Dim i As Long
    Dim userChoice As String
    Dim selectedIndexes As Variant
    Dim selectedValues As New Collection
    Dim listPrompt As String

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
            selectedValues.Add idList(i)
        Next i
    Else
        selectedIndexes = Split(userChoice, ",")
        For i = LBound(selectedIndexes) To UBound(selectedIndexes)
            Dim idx As Long
            idx = Val(Trim(selectedIndexes(i)))
            If idx >= 1 And idx <= idList.Count Then
                selectedValues.Add idList(idx)
            End If
        Next i
    End If
    Set ChooseMultipleValuesFromListWithAll = selectedValues
    Exit Function
    
ErrorHandler:
    HandleError "LoadQueries", "ChooseMultipleValuesFromListWithAll", "Erreur lors de la sélection des valeurs"
End Function

Function ChooseMultipleValuesFromArrayWithAll(values() As String, prompt As String) As Collection
    On Error GoTo ErrorHandler
    
    If Not IsArray(values) Then
        HandleError "LoadQueries", "ChooseMultipleValuesFromArrayWithAll", "Tableau non initialisé"
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

    userChoice = InputBox(listPrompt, "Sélection", "1")
    If StrPtr(userChoice) = 0 Or Len(Trim(userChoice)) = 0 Then
        Exit Function
    End If
    
    Dim selectedValues As New Collection
    userChoice = Trim(userChoice)
    
    ' Cas spécial : sélection de toutes les valeurs avec *
    If userChoice = "*" Then
        For i = 1 To UBound(values)
            selectedValues.Add values(i)
        Next i
    Else
        ' Sélection de valeurs spécifiques
        Dim selectedIndexes As Variant
        selectedIndexes = Split(userChoice, ",")
        Dim hasValidSelection As Boolean
        hasValidSelection = False
        
        For i = LBound(selectedIndexes) To UBound(selectedIndexes)
            Dim idx As Long
            idx = Val(Trim(selectedIndexes(i)))
            If idx >= 1 And idx <= UBound(values) Then
                selectedValues.Add values(idx)
                hasValidSelection = True
            End If
        Next i
        
        ' Si aucune sélection valide n'a été trouvée
        If Not hasValidSelection Then
            HandleError "LoadQueries", "ChooseMultipleValuesFromArrayWithAll", "Aucune sélection valide"
            Exit Function
        End If
    End If
    
    Set ChooseMultipleValuesFromArrayWithAll = selectedValues
    Exit Function
    
ErrorHandler:
    HandleError "LoadQueries", "ChooseMultipleValuesFromArrayWithAll", "Erreur lors de la sélection des valeurs"
End Function
