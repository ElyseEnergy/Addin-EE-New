Option Explicit
Private Const MODULE_NAME As String = "LoadQueries"
Private Const ERROR_HANDLER_LABEL As String = "ErrorHandler"

Sub LoadQuery(QueryName As String, ws As Worksheet, DestCell As Range)
    Const PROC_NAME As String = "LoadQuery"
    On Error GoTo ErrorHandler
    
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
        ElyseMessageBox_System.ShowErrorMessage "Erreur de chargement", _
            "Erreur lors du chargement de la requête " & QueryName & ": " & Err.Description
    End If
    On Error GoTo 0
    
    ' Après le chargement de la requête, s'assurer que le nom est correct
    Set lo = ws.ListObjects(ws.ListObjects.Count) ' Le dernier tableau créé
    If Not lo Is Nothing Then
        lo.Name = sanitizedName
    End If
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
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

Function ChooseMultipleValuesFromArrayWithAll(idList As Collection, displayList As Collection, prompt As String) As Collection
    Const PROC_NAME As String = "ChooseMultipleValuesFromArrayWithAll"
    On Error GoTo ErrorHandler
    
    Dim selectedValues As New Collection
    
    ' Use custom list selection form with AllowMultiSelect=True
    Dim modeItems As Collection
    Set modeItems = New Collection
    
    ' Add the "All" option at the top
    modeItems.Add "* : Toutes"
    
    ' Add numbered options
    Dim i As Long
    For i = 1 To displayList.Count
        modeItems.Add i & ". " & displayList(i)
    Next i
    
    Dim result As Long
    result = ElyseMain_Orchestrator.SelectFromList( _
        "Sélection des valeurs", _
        prompt, _
        modeItems)
    
    ' Handle selection
    If result = 1 Then ' "All" option selected
        For i = 1 To idList.Count
            selectedValues.Add idList(i)
        Next i
    ElseIf result > 1 Then ' Specific item selected
        selectedValues.Add idList(result - 1) ' -1 because we added "All" at the top
    End If
    
    Set ChooseMultipleValuesFromArrayWithAll = selectedValues
    Exit Function

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
    Set ChooseMultipleValuesFromArrayWithAll = New Collection ' Return empty collection on error
End Function

Public Sub ExecuteQuery(ByVal queryName As String)
    Const PROC_NAME As String = "ExecuteQuery"
    On Error GoTo ErrorHandler
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Executing query: " & queryName, PROC_NAME, MODULE_NAME

    ' Verify if query exists and refresh
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

Public Function GetUserSelection(queries As Collection) As String
    Dim listPrompt As String
    Dim i As Long
    
    ' Build the prompt string
    listPrompt = "Choisissez une requête à charger :" & vbCrLf & vbCrLf
    For i = 1 To queries.Count
        listPrompt = listPrompt & i & ". " & queries(i) & vbCrLf
    Next i
    
    ' Get user input using the message box system
    Dim userChoice As String
    userChoice = ElyseMessageBox_System.ShowInputDialog(MSG_TITLE_SELECT, listPrompt, "1")
    
    ' Validate user input
    If userChoice = "" Then
        GetUserSelection = ""
        Exit Function
    End If
    
    If IsNumeric(userChoice) Then
        Dim choiceNum As Long
        choiceNum = CLng(userChoice)
        If choiceNum >= 1 And choiceNum <= queries.Count Then
            GetUserSelection = queries(choiceNum)
        Else
            ElyseMessageBox_System.ShowErrorMessage MSG_TITLE_ERROR, _
                "Veuillez entrer un numéro entre 1 et " & queries.Count
            GetUserSelection = ""
        End If
    Else
        ElyseMessageBox_System.ShowErrorMessage MSG_TITLE_ERROR, _
            "Veuillez entrer un numéro valide"
        GetUserSelection = ""
    End If
End Function

' Add similar transformations for other Subs/Functions in this file.
