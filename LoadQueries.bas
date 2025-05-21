Sub LoadQuery(QueryName As String, ws As Worksheet, DestCell As Range)
    Dim lo As ListObject
    ' Vérifier si la table existe déjà
    For Each lo In ws.ListObjects
        If lo.Name = "Table_" & QueryName Then
            Exit Sub
        End If
    Next lo

    On Error Resume Next
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
        .ListObject.DisplayName = "Table_" & QueryName
        .Refresh BackgroundQuery:=False
    End With
    If Err.Number <> 0 Then
        MsgBox "Erreur lors du chargement de la requête " & QueryName & ": " & Err.Description, vbExclamation
    End If
    On Error GoTo 0
End Sub

Sub LoadQueryWithFilter(QueryName As String, ws As Worksheet, DestCell As Range, filterCriteria As String)
    Dim lo As ListObject
    ' Vérifier si la table existe déjà
    For Each lo In ws.ListObjects
        If lo.Name = "Table_" & QueryName Then
            Exit Sub
        End If
    Next lo

    On Error Resume Next
    With ws.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & QueryName & ";Extended Properties=""""", _
        Destination:=DestCell).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & QueryName & "] WHERE " & filterCriteria)
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .ListObject.DisplayName = "Table_" & QueryName
        .Refresh BackgroundQuery:=False
    End With
    If Err.Number <> 0 Then
        MsgBox "Erreur lors du chargement de la requête " & QueryName & ": " & Err.Description, vbExclamation
    End If
    On Error GoTo 0
End Sub

Function ChooseUniqueValueFromTable(ws As Worksheet, tableName As String, colName As String, prompt As String) As String
    Dim lo As ListObject
    Dim cell As Range
    Dim values As Collection
    Dim v As Variant
    Dim userChoice As String
    Set values = New Collection
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    ' Collecter les valeurs uniques
    For Each cell In lo.ListColumns(colName).DataBodyRange
        If cell.Value <> "" Then
            On Error Resume Next
            values.Add cell.Value, CStr(cell.Value)
            On Error GoTo 0
        End If
    Next cell
    If values.Count = 0 Then Exit Function
    ' Construire la liste pour l'InputBox
    Dim listPrompt As String
    listPrompt = prompt & vbCrLf
    For i = 1 To values.Count
        listPrompt = listPrompt & i & ". " & values(i) & vbCrLf
    Next i    userChoice = InputBox(listPrompt, "Sélection", "1")
    
    ' Gestion du bouton Annuler ou entrée vide
    If StrPtr(userChoice) = 0 Or Len(Trim(userChoice)) = 0 Then
        Exit Function
    End If
    
    ' Validation de la saisie
    If IsNumeric(userChoice) Then
        Dim choiceNum As Long
        choiceNum = CLng(userChoice)
        If choiceNum >= 1 And choiceNum <= values.Count Then
            ChooseUniqueValueFromTable = values(choiceNum)
        Else
            MsgBox "Veuillez entrer un numéro entre 1 et " & values.Count, vbExclamation
        End If
    Else
        MsgBox "Veuillez entrer un numéro valide", vbExclamation
    End If
End Function

Function ChooseValueFromTableWithDisplay(ws As Worksheet, tableName As String, valueColumn As String, displayColumn As String, prompt As String) As String
    Dim lo As ListObject
    Dim valueCell As Range, displayCell As Range
    Dim values As Collection, displays As Collection
    Dim userChoice As String
    
    Set values = New Collection
    Set displays = New Collection
    
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    
    ' Collecter les valeurs et les textes à afficher
    Dim i As Long
    i = 1
    For Each valueCell In lo.ListColumns(valueColumn).DataBodyRange
        Set displayCell = lo.ListColumns(displayColumn).DataBodyRange.Cells(i, 1)
        If valueCell.Value <> "" Then
            values.Add valueCell.Value
            displays.Add displayCell.Value
        End If
        i = i + 1
    Next valueCell
    
    If values.Count = 0 Then Exit Function
    
    ' Construire la liste pour l'InputBox
    Dim listPrompt As String
    listPrompt = prompt & vbCrLf
    For i = 1 To displays.Count
        listPrompt = listPrompt & i & ". " & displays(i) & vbCrLf
    Next i
      userChoice = InputBox(listPrompt, "Sélection", "1")
    
    ' Gestion du bouton Annuler ou entrée vide
    If StrPtr(userChoice) = 0 Or Len(Trim(userChoice)) = 0 Then
        Exit Function
    End If
    
    ' Validation de la saisie
    If IsNumeric(userChoice) Then
        Dim choiceNum As Long
        choiceNum = CLng(userChoice)
        If choiceNum >= 1 And choiceNum <= values.Count Then
            ChooseValueFromTableWithDisplay = values(choiceNum)
        Else
            MsgBox "Veuillez entrer un numéro entre 1 et " & values.Count, vbExclamation
        End If
    Else
        MsgBox "Veuillez entrer un numéro valide", vbExclamation
    End If
End Function

Function ChooseMultipleValuesFromTable(ws As Worksheet, tableName As String, colName As String, prompt As String) As Collection
    Dim lo As ListObject
    Dim cell As Range
    Dim values As Collection
    Dim userChoice As String
    Dim i As Long
    Set values = New Collection
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If lo Is Nothing Then Exit Function
    ' Collecter les valeurs uniques
    For Each cell In lo.ListColumns(colName).DataBodyRange
        If cell.Value <> "" Then
            On Error Resume Next
            values.Add cell.Value, CStr(cell.Value)
            On Error GoTo 0
        End If
    Next cell
    If values.Count = 0 Then Exit Function
    ' Construire la liste pour l'InputBox
    Dim listPrompt As String
    listPrompt = prompt & vbCrLf
    For i = 1 To values.Count
        listPrompt = listPrompt & i & ". " & values(i) & vbCrLf
    Next i    userChoice = InputBox(listPrompt, "Sélection", "1")
    
    ' Gestion du bouton Annuler ou entrée vide
    If StrPtr(userChoice) = 0 Or Len(Trim(userChoice)) = 0 Then
        Exit Function
    End If
    
    Dim selectedIndexes As Variant
    selectedIndexes = Split(Trim(userChoice), ",")
    Dim selectedValues As New Collection
    Dim hasValidSelection As Boolean
    hasValidSelection = False
    
    For i = LBound(selectedIndexes) To UBound(selectedIndexes)
        Dim idx As Long
        idx = Val(Trim(selectedIndexes(i)))
        If idx >= 1 And idx <= values.Count Then
            selectedValues.Add values(idx)
            hasValidSelection = True
        End If
    Next i
    
    ' Si aucune sélection valide n'a été trouvée
    If Not hasValidSelection Then
        MsgBox "Veuillez entrer des numéros valides entre 1 et " & values.Count & vbCrLf & _
               "Exemple: 1,2,3", vbExclamation
        Exit Function
    End If
    Set ChooseMultipleValuesFromTable = selectedValues
End Function

Function ChooseMultipleValuesFromList(idList As Collection, displayList As Collection, prompt As String) As Collection
    Dim i As Long
    Dim userChoice As String
    Dim selectedIndexes As Variant
    Dim selectedValues As New Collection
    Dim listPrompt As String

    listPrompt = prompt & vbCrLf
    For i = 1 To displayList.Count
        listPrompt = listPrompt & i & ". " & displayList(i) & vbCrLf
    Next i
    userChoice = InputBox(listPrompt, "Sélection", "1")
    If userChoice = "" Then Exit Function
    selectedIndexes = Split(userChoice, ",")
    For i = LBound(selectedIndexes) To UBound(selectedIndexes)
        Dim idx As Long
        idx = Val(Trim(selectedIndexes(i)))
        If idx >= 1 And idx <= idList.Count Then
            selectedValues.Add idList(idx)
        End If
    Next i
    Set ChooseMultipleValuesFromList = selectedValues
End Function


    Function ChooseMultipleValuesFromTableWithAll(ws As Worksheet, tableName As String, colName As String, prompt As String) As Collection
        Dim lo As ListObject
        Dim cell As Range
        Dim values As Collection
        Dim userChoice As String
        Dim i As Long
        Set values = New Collection
        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo 0
        If lo Is Nothing Then Exit Function
        ' Collecter les valeurs uniques
        For Each cell In lo.ListColumns(colName).DataBodyRange
            If cell.Value <> "" Then
                On Error Resume Next
                values.Add cell.Value, CStr(cell.Value)
                On Error GoTo 0
            End If
        Next cell
        If values.Count = 0 Then Exit Function
        ' Construire la liste pour l'InputBox
        Dim listPrompt As String
        listPrompt = prompt & vbCrLf & "* : Toutes" & vbCrLf
        For i = 1 To values.Count
            listPrompt = listPrompt & i & ". " & values(i) & vbCrLf
        Next i        userChoice = InputBox(listPrompt, "Sélection", "1")
        
        ' Gestion du bouton Annuler ou entrée vide
        If StrPtr(userChoice) = 0 Or Len(Trim(userChoice)) = 0 Then
            Exit Function
        End If
        
        Dim selectedValues As New Collection
        userChoice = Trim(userChoice)
        
        ' Cas spécial : sélection de toutes les valeurs avec *
        If userChoice = "*" Then
            For i = 1 To values.Count
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
                If idx >= 1 And idx <= values.Count Then
                    selectedValues.Add values(idx)
                    hasValidSelection = True
                End If
            Next i
            
            ' Si aucune sélection valide n'a été trouvée
            If Not hasValidSelection Then
                MsgBox "Veuillez entrer des numéros valides entre 1 et " & values.Count & vbCrLf & _
                       "Ou * pour sélectionner toutes les valeurs" & vbCrLf & _
                       "Exemple: 1,2,3", vbExclamation
                Exit Function
            End If
        End If
        Set ChooseMultipleValuesFromTableWithAll = selectedValues
    End Function
    
    Function ChooseMultipleValuesFromListWithAll(idList As Collection, displayList As Collection, prompt As String) As Collection
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
        If userChoice = "" Then Exit Function
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
    End Function
    
    