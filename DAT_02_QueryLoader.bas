Option Explicit
Private Const MODULE_NAME As String = "LoadQueries"
Private Const ERROR_HANDLER_LABEL As String = "ErrorHandler"

Sub LoadQuery(QueryName As String, ws As Worksheet, DestCell As Range)
    Const PROC_NAME As String = "LoadQuery"
    On Error GoTo ErrorHandler
    
    LogInfo PROC_NAME & "_Start", "Début du chargement de la requête: " & QueryName, PROC_NAME, MODULE_NAME
    
    ' Validation des paramètres
    If QueryName = "" Then
        LogError PROC_NAME & "_InvalidQuery", 0, "Nom de requête vide", PROC_NAME, MODULE_NAME
        Exit Sub
    End If
    
    If ws Is Nothing Then
        LogError PROC_NAME & "_InvalidSheet", 0, "Feuille de destination invalide", PROC_NAME, MODULE_NAME
        Exit Sub
    End If
    
    If DestCell Is Nothing Then
        LogError PROC_NAME & "_InvalidCell", 0, "Cellule de destination invalide", PROC_NAME, MODULE_NAME
        Exit Sub
    End If
    
    Dim lo As ListObject
    Dim sanitizedName As String
    
    ' Nettoyer le nom de la requête pour le nom de tableau
    sanitizedName = "Table_" & Utilities.SanitizeTableName(QueryName)
    LogDebug PROC_NAME & "_SanitizedName", "Nom nettoyé: " & sanitizedName, PROC_NAME, MODULE_NAME
    
    ' Vérifier si la table existe déjà
    LogDebug PROC_NAME & "_CheckExisting", "Vérification des tables existantes", PROC_NAME, MODULE_NAME
    For Each lo In ws.ListObjects
        If lo.Name = sanitizedName Then
            LogInfo PROC_NAME & "_TableExists", "Table existante trouvée: " & sanitizedName, PROC_NAME, MODULE_NAME
            Exit Sub
        End If
    Next lo

    LogDebug PROC_NAME & "_CreateTable", "Création de la nouvelle table", PROC_NAME, MODULE_NAME
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
        LogError PROC_NAME & "_LoadError", Err.Number, "Error loading query: " & Err.Description, PROC_NAME, MODULE_NAME
        ShowErrorMessage "Load Error", "An error occurred while loading the query. Details: " & Err.Description
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    ' Après le chargement de la requête, s'assurer que le nom est correct
    LogDebug PROC_NAME & "_VerifyName", "Vérification du nom de la table", PROC_NAME, MODULE_NAME
    Set lo = ws.ListObjects(ws.ListObjects.Count) ' Le dernier tableau créé
    If Not lo Is Nothing Then
        lo.Name = sanitizedName
        LogInfo PROC_NAME & "_Success", "Table créée avec succès: " & sanitizedName, PROC_NAME, MODULE_NAME
    Else
        LogWarning PROC_NAME & "_NoTable", "Aucune table n'a été créée", PROC_NAME, MODULE_NAME
    End If
    Exit Sub

ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Erreur lors du chargement de la requête: " & Err.Description, PROC_NAME, MODULE_NAME
End Sub

Function ChooseMultipleValuesFromArrayWithAll(idList As Collection, displayList As Collection, prompt As String) As Collection
    Const PROC_NAME As String = "ChooseMultipleValuesFromArrayWithAll"
    On Error GoTo ErrorHandler
    
    LogInfo PROC_NAME & "_Start", "Début de la sélection multiple", PROC_NAME, MODULE_NAME
    
    ' Validation des paramètres
    If idList Is Nothing Or displayList Is Nothing Then
        LogError PROC_NAME & "_InvalidParams", 0, "Listes d'entrée invalides", PROC_NAME, MODULE_NAME
        Set ChooseMultipleValuesFromArrayWithAll = New Collection
        Exit Function
    End If
    
    If idList.Count <> displayList.Count Then
        LogError PROC_NAME & "_MismatchedLists", 0, "Les listes ont des tailles différentes: " & idList.Count & " vs " & displayList.Count, PROC_NAME, MODULE_NAME
        Set ChooseMultipleValuesFromArrayWithAll = New Collection
        Exit Function
    End If
    
    Dim selectedValues As New Collection
    LogDebug PROC_NAME & "_Init", "Initialisation de la sélection", PROC_NAME, MODULE_NAME
    
    ' Use custom list selection form with AllowMultiSelect=True
    Dim modeItems As Collection
    Set modeItems = New Collection
    
    ' Add the "All" option at the top
    modeItems.Add "* : Toutes"
    LogDebug PROC_NAME & "_AddAllOption", "Option 'Toutes' ajoutée", PROC_NAME, MODULE_NAME
    
    ' Add numbered options
    Dim i As Long
    For i = 1 To displayList.Count
        modeItems.Add i & ". " & displayList(i)
        LogDebug PROC_NAME & "_AddOption", "Option ajoutée: " & i & ". " & displayList(i), PROC_NAME, MODULE_NAME
    Next i
    
    LogDebug PROC_NAME & "_ShowSelector", "Affichage du sélecteur de liste", PROC_NAME, MODULE_NAME
    Dim result As Long
    result = SelectFromList( _
        "Sélection des valeurs", _
        prompt, _
        modeItems)
    
    ' Handle selection
    If result = 1 Then ' "All" option selected
        LogInfo PROC_NAME & "_AllSelected", "Option 'Toutes' sélectionnée", PROC_NAME, MODULE_NAME
        For i = 1 To idList.Count
            selectedValues.Add idList(i)
        Next i
        LogDebug PROC_NAME & "_AllAdded", "Toutes les valeurs ajoutées: " & idList.Count & " éléments", PROC_NAME, MODULE_NAME
    ElseIf result > 1 Then ' Specific item selected
        LogInfo PROC_NAME & "_SingleSelected", "Valeur spécifique sélectionnée: " & displayList(result - 1), PROC_NAME, MODULE_NAME
        selectedValues.Add idList(result - 1) ' -1 because we added "All" at the top
        LogDebug PROC_NAME & "_SingleAdded", "Valeur ajoutée: " & idList(result - 1), PROC_NAME, MODULE_NAME
    Else
        LogWarning PROC_NAME & "_NoSelection", "Aucune sélection effectuée", PROC_NAME, MODULE_NAME
    End If
    
    Set ChooseMultipleValuesFromArrayWithAll = selectedValues
    LogInfo PROC_NAME & "_Complete", "Sélection terminée: " & selectedValues.Count & " éléments", PROC_NAME, MODULE_NAME
    Exit Function

ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Erreur lors de la sélection multiple: " & Err.Description, PROC_NAME, MODULE_NAME
    Set ChooseMultipleValuesFromArrayWithAll = New Collection ' Return empty collection on error
End Function

Public Sub ExecuteQuery(ByVal queryName As String)
    Const PROC_NAME As String = "ExecuteQuery"
    On Error GoTo ErrorHandler
    
    LogInfo PROC_NAME & "_Start", "Executing query: " & queryName, PROC_NAME, MODULE_NAME

    ' Verify if query exists and refresh
    Dim queryFound As Boolean
    queryFound = False
    
    Dim conn As Object ' WorkbookConnection
    Dim pqQuery As Object ' WorkbookQuery
    
    On Error Resume Next ' Check for connection first
    Set conn = ThisWorkbook.Connections(queryName)
    If Not conn Is Nothing Then
        LogDebug PROC_NAME & "_ConnectionFound", "Connection '" & queryName & "' found. Attempting refresh.", PROC_NAME, MODULE_NAME
        conn.Refresh
        queryFound = True
        LogInfo PROC_NAME & "_ConnectionRefresh", "Connection '" & queryName & "' refreshed.", PROC_NAME, MODULE_NAME
    Else
        Set pqQuery = ThisWorkbook.Queries(queryName) ' Check for query if not a connection
        If Not pqQuery Is Nothing Then
            LogDebug PROC_NAME & "_PQFound", "Power Query '" & queryName & "' found. Attempting refresh.", PROC_NAME, MODULE_NAME
            pqQuery.Refresh
            queryFound = True
            LogInfo PROC_NAME & "_PQRefresh", "Power Query '" & queryName & "' refreshed.", PROC_NAME, MODULE_NAME
        End If
    End If
    On Error GoTo ErrorHandler ' Reinstate proper error handling

    If Not queryFound Then
        LogWarning PROC_NAME & "_NotFound", "Query or Connection '" & queryName & "' not found.", PROC_NAME, MODULE_NAME
        ShowWarningMessage "Query Error", "Query or Connection '" & queryName & "' could not be found."
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Function ListAllQueries() As Collection
    Const PROC_NAME As String = "ListAllQueries"
    On Error GoTo ErrorHandler
    
    LogInfo PROC_NAME & "_Start", "Listing all Power Queries and Connections.", PROC_NAME, MODULE_NAME
    
    Dim queriesList As Collection
    Set queriesList = New Collection
    
    Dim conn As Object ' WorkbookConnection
    Dim pq As Object ' WorkbookQuery
    
    ' Debug.Print "Available Connections:"
    LogDebug PROC_NAME & "_ListConnections", "Listing Workbook Connections...", PROC_NAME, MODULE_NAME
    For Each conn In ThisWorkbook.Connections
        ' Debug.Print " - " & conn.Name
        queriesList.Add conn.Name
        LogDebug PROC_NAME & "_ConnItem", "Connection found: " & conn.Name, PROC_NAME, MODULE_NAME
    Next conn
    
    ' Debug.Print "Available Power Queries:"
    LogDebug PROC_NAME & "_ListQueries", "Listing Workbook Queries...", PROC_NAME, MODULE_NAME
    For Each pq In ThisWorkbook.Queries
        ' Debug.Print " - " & pq.Name
        queriesList.Add pq.Name
        LogDebug PROC_NAME & "_PQItem", "Power Query found: " & pq.Name, PROC_NAME, MODULE_NAME
    Next pq
    
    Set ListAllQueries = queriesList
    LogInfo PROC_NAME & "_End", "Found " & queriesList.Count & " queries/connections.", PROC_NAME, MODULE_NAME
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
    Set ListAllQueries = New Collection ' Return empty collection on error
End Function

Public Function GetUserSelection(queries As Collection) As String
    Const PROC_NAME As String = "GetUserSelection"
    On Error GoTo ErrorHandler
    
    LogInfo PROC_NAME & "_Start", "Début de la sélection utilisateur", PROC_NAME, MODULE_NAME
    
    ' Validation des paramètres
    If queries Is Nothing Then
        LogError PROC_NAME & "_InvalidQueries", 0, "Collection de requêtes invalide", PROC_NAME, MODULE_NAME
        GetUserSelection = ""
        Exit Function
    End If
    
    If queries.Count = 0 Then
        LogWarning PROC_NAME & "_EmptyQueries", "Aucune requête disponible", PROC_NAME, MODULE_NAME
        GetUserSelection = ""
        Exit Function
    End If
    
    Dim listPrompt As String
    Dim i As Long
    
    ' Build the prompt string
    LogDebug PROC_NAME & "_BuildPrompt", "Construction du message de sélection", PROC_NAME, MODULE_NAME
    listPrompt = "Choisissez une requête à charger :" & vbCrLf & vbCrLf
    For i = 1 To queries.Count
        listPrompt = listPrompt & i & ". " & queries(i) & vbCrLf
        LogDebug PROC_NAME & "_AddOption", "Option ajoutée: " & i & ". " & queries(i), PROC_NAME, MODULE_NAME
    Next i
    
    ' Get user input using the message box system
    LogDebug PROC_NAME & "_ShowDialog", "Affichage de la boîte de dialogue", PROC_NAME, MODULE_NAME
    Dim userChoice As String
    userChoice = ShowInputDialog(MSG_TITLE_SELECT, listPrompt, "1")
    
    ' Validate user input
    If userChoice = "" Then
        LogWarning PROC_NAME & "_EmptyChoice", "Aucune sélection effectuée par l'utilisateur", PROC_NAME, MODULE_NAME
        GetUserSelection = ""
        Exit Function
    End If
    
    If IsNumeric(userChoice) Then
        Dim choiceNum As Long
        choiceNum = CLng(userChoice)
        If choiceNum >= 1 And choiceNum <= queries.Count Then
            LogInfo PROC_NAME & "_ValidChoice", "Sélection valide: " & queries(choiceNum), PROC_NAME, MODULE_NAME
            GetUserSelection = queries(choiceNum)
        Else
            LogWarning PROC_NAME & "_InvalidRange", "Numéro hors limites: " & choiceNum & " (1-" & queries.Count & ")", PROC_NAME, MODULE_NAME
            ShowErrorMessage MSG_TITLE_ERROR, "Veuillez entrer un numéro entre 1 et " & queries.Count
            GetUserSelection = ""
        End If
    Else
        LogWarning PROC_NAME & "_InvalidInput", "Entrée non numérique: " & userChoice, PROC_NAME, MODULE_NAME
        ShowErrorMessage MSG_TITLE_ERROR, "Veuillez entrer un numéro valide"
        GetUserSelection = ""
    End If
    Exit Function

ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Erreur lors de la sélection utilisateur: " & Err.Description, PROC_NAME, MODULE_NAME
    GetUserSelection = ""
End Function


