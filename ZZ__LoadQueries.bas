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
    Next i
    userChoice = InputBox(listPrompt, "Sélection", "1")
    If userChoice = "" Then Exit Function
    If IsNumeric(userChoice) Then
        If CInt(userChoice) >= 1 And CInt(userChoice) <= values.Count Then
            ChooseUniqueValueFromTable = values(CInt(userChoice))
        End If
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
    If userChoice = "" Then Exit Function
    
    If IsNumeric(userChoice) Then
        If CInt(userChoice) >= 1 And CInt(userChoice) <= values.Count Then
            ChooseValueFromTableWithDisplay = values(CInt(userChoice))
        End If
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
    Next i
    userChoice = InputBox(listPrompt, "Sélection", "1")
    If userChoice = "" Then Exit Function
    Dim selectedIndexes As Variant
    selectedIndexes = Split(userChoice, ",")
    Dim selectedValues As New Collection
    For i = LBound(selectedIndexes) To UBound(selectedIndexes)
        Dim idx As Long
        idx = Val(Trim(selectedIndexes(i)))
        If idx >= 1 And idx <= values.Count Then
            selectedValues.Add values(idx)
        End If
    Next i
    Set ChooseMultipleValuesFromTable = selectedValues
End Function