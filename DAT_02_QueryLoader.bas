Option Explicit
Private Const MODULE_NAME As String = "LoadQueries"

Sub LoadQuery(QueryName As String, ws As Worksheet, DestCell As Range)
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
        .ListObject.DisplayName = sanitizedName
        .Refresh BackgroundQuery:=False
    End With
    If Err.Number <> 0 Then
        MsgBox "Erreur lors du chargement de la requête " & QueryName & ": " & Err.Description, vbExclamation
    End If
    On Error GoTo 0
    
    ' Après le chargement de la requête, s'assurer que le nom est correct
    Set lo = ws.ListObjects(ws.ListObjects.Count) ' Le dernier tableau créé
    If Not lo Is Nothing Then
        lo.Name = sanitizedName
    End If
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

Public Sub ExecuteQuery(ByVal queryName As String)
    Const PROC_NAME As String = "ExecuteQuery"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Executing query: " & queryName, PROC_NAME, MODULE_NAME

    ' Dim wb As Workbook ' Original
    ' Dim pq As Object ' Power Query connection/query object, original
    ' Set wb = ThisWorkbook ' Original

    ' On Error Resume Next ' Original error handling might have been inline
    ' Set pq = wb.Connections(queryName).OLEDBConnection.ADOConnection ' Example of how one might try to get a query; this is not standard for PQ
    ' On Error GoTo 0 ' Original error handling reset

    ' If pq Is Nothing Then
    '   Debug.Print "Query '" & queryName & "' not found or not a Power Query connection."
    '   MsgBox "Query '" & queryName & "' could not be found or is not a valid Power Query connection.", vbExclamation, "Query Error"
    '   Exit Sub
    ' End If
    
    ' Debug.Print "Refreshing query: " & queryName
    ' pq.Refresh ' This is not how Power Queries are typically refreshed directly in VBA for all cases.
    ' More common: ThisWorkbook.Connections(queryName).Refresh or ActiveWorkbook.Queries(queryName).Refresh

    Dim queryFound As Boolean
    queryFound = False
    
    Dim conn As Object ' WorkbookConnection
    Dim pqQuery As Object ' WorkbookQuery
    
    On Error Resume Next ' Check for connection first
    Set conn = ThisWorkbook.Connections(queryName)
    If Not conn Is Nothing Then
        ElyseMain_Orchestrator.LogDebug PROC_NAME & "_ConnectionFound", "Connection '" & queryName & "' found. Attempting refresh.", PROC_NAME, MODULE_NAME
        conn.Refresh
        queryFound = True
        ElyseMain_Orchestrator.LogInfo PROC_NAME & "_ConnectionRefresh", "Connection '" & queryName & "' refreshed.", PROC_NAME, MODULE_NAME
    Else
        Set pqQuery = ThisWorkbook.Queries(queryName) ' Check for query if not a connection
        If Not pqQuery Is Nothing Then
            ElyseMain_Orchestrator.LogDebug PROC_NAME & "_PQFound", "Power Query '" & queryName & "' found. Attempting refresh.", PROC_NAME, MODULE_NAME
            pqQuery.Refresh
            queryFound = True
            ElyseMain_Orchestrator.LogInfo PROC_NAME & "_PQRefresh", "Power Query '" & queryName & "' refreshed.", PROC_NAME, MODULE_NAME
        End If
    End If
    On Error GoTo ErrorHandler ' Reinstate proper error handling

    If Not queryFound Then
        ElyseMain_Orchestrator.LogWarning PROC_NAME & "_NotFound", "Query or Connection '" & queryName & "' not found.", PROC_NAME, MODULE_NAME
        ElyseMessageBox_System.ShowWarningMessage "Query Error", "Query or Connection '" & queryName & "' could not be found."
    End If
    
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
    ' If Err.Number <> 0 Then ' Check if it was a real error vs. just not found handled above
    '    ElyseMessageBox_System.ShowErrorMessage "Query Execution Error", "An error occurred while executing query '" & queryName & "'. Details: " & Err.Description
    ' End If
End Sub

Public Function ListAllQueries() As Collection
    Const PROC_NAME As String = "ListAllQueries"
    On Error GoTo ErrorHandler
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Listing all Power Queries and Connections.", PROC_NAME, MODULE_NAME
    
    Dim queriesList As Collection
    Set queriesList = New Collection
    
    Dim conn As Object ' WorkbookConnection
    Dim pq As Object ' WorkbookQuery
    
    ' Debug.Print "Available Connections:"
    ElyseMain_Orchestrator.LogDebug PROC_NAME & "_ListConnections", "Listing Workbook Connections...", PROC_NAME, MODULE_NAME
    For Each conn In ThisWorkbook.Connections
        ' Debug.Print " - " & conn.Name
        queriesList.Add conn.Name
        ElyseMain_Orchestrator.LogDebug PROC_NAME & "_ConnItem", "Connection found: " & conn.Name, PROC_NAME, MODULE_NAME
    Next conn
    
    ' Debug.Print "Available Power Queries:"
    ElyseMain_Orchestrator.LogDebug PROC_NAME & "_ListQueries", "Listing Workbook Queries...", PROC_NAME, MODULE_NAME
    For Each pq In ThisWorkbook.Queries
        ' Debug.Print " - " & pq.Name
        queriesList.Add pq.Name
        ElyseMain_Orchestrator.LogDebug PROC_NAME & "_PQItem", "Power Query found: " & pq.Name, PROC_NAME, MODULE_NAME
    Next pq
    
    Set ListAllQueries = queriesList
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Found " & queriesList.Count & " queries/connections.", PROC_NAME, MODULE_NAME
    Exit Function

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
    Set ListAllQueries = New Collection ' Return empty collection on error
End Function

' Add similar transformations for other Subs/Functions in this file.
