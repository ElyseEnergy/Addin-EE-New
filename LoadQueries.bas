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
End Function

Function ChooseMultipleValuesFromArrayWithAll(values() As String, prompt As String) As Collection
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
                MsgBox "Veuillez entrer des numéros valides entre 1 et " & UBound(values) & vbCrLf & _
                       "Ou * pour sélectionner toutes les valeurs" & vbCrLf & _
                       "Exemple: 1,2,3", vbExclamation
                Exit Function
            End If
        End If
        
        Set ChooseMultipleValuesFromArrayWithAll = selectedValues
    End Function
