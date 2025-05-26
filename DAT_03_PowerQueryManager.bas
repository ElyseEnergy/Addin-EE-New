' Module: PQQueryManager
' Gère la création et la vérification des requêtes PowerQuery
Option Explicit
Private Const MODULE_NAME As String = "PQQueryManager"

Private mColumnTypes As Object ' Dictionnaire pour stocker les types de colonnes

' Vérifie si une requête PowerQuery existe et la crée si nécessaire
Public Function EnsurePQQueryExists(category As CategoryInfo) As Boolean
    Dim query As String
    query = GeneratePQQueryTemplate(category)
    
    If QueryExists(category.PowerQueryName) Then
        ' Si la requête existe, mettre à jour sa formule
        On Error Resume Next
        ThisWorkbook.Queries(category.PowerQueryName).Formula = query
        Dim updateError As Long
        updateError = Err.Number
        On Error GoTo 0
        
        If updateError <> 0 Then
            Debug.Print "Erreur lors de la mise à jour de la requête " & category.PowerQueryName & ": " & Err.Description
            EnsurePQQueryExists = False
            Exit Function
        End If
        
        ' Rafraîchir la requête
        ThisWorkbook.Queries(category.PowerQueryName).Refresh
        EnsurePQQueryExists = True
    Else
        ' Créer la requête si elle n'existe pas
        If Not AddQueryToPowerQuery(category.PowerQueryName, query) Then
            EnsurePQQueryExists = False
            Exit Function
        End If
        EnsurePQQueryExists = True
    End If
    
    ' Stocker les types de colonnes après la création de la requête
    StoreColumnTypes category.PowerQueryName
    
    EnsurePQQueryExists = True
End Function

' Vérifie si une requête PowerQuery existe
Private Function QueryExists(queryName As String) As Boolean
    On Error Resume Next
    Dim query As Object
    Set query = ThisWorkbook.Queries(queryName)
    QueryExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' Ajoute une requête PowerQuery
Private Function AddQueryToPowerQuery(queryName As String, query As String) As Boolean
    On Error Resume Next
    ThisWorkbook.Queries.Add queryName, query
    Dim errNum As Long
    errNum = Err.Number
    If errNum <> 0 Then
        Debug.Print "Erreur lors de l'ajout de la requête " & queryName & ": " & Err.Description
        AddQueryToPowerQuery = False
    Else
        AddQueryToPowerQuery = True
    End If
    On Error GoTo 0
End Function

' Génère le template de requête PowerQuery
Private Function GeneratePQQueryTemplate(category As CategoryInfo) As String
    Dim template As String
    ' Template de base pour charger les données depuis l'API Ragic avec réorganisation des colonnes
    template = "let" & vbCrLf & _
          "    Source = Csv.Document(Web.Contents(""" & category.URL & """),[Delimiter="","",Encoding=65001,QuoteStyle=QuoteStyle.Csv])," & vbCrLf & _
          "    PromotedHeaders = Table.PromoteHeaders(Source)," & vbCrLf & _
          "    // Trouver la colonne ID quelle que soit sa casse" & vbCrLf & _
          "    Colonnes = Table.ColumnNames(PromotedHeaders)," & vbCrLf & _
          "    IdColumn = List.First(List.Select(Colonnes, each Text.Lower(_) = ""id""))," & vbCrLf & _
          "    AutresColonnes = List.Select(Colonnes, each Text.Lower(_) <> ""id"")," & vbCrLf & _
          "    // Réorganiser les colonnes pour avoir ID en premier" & vbCrLf & _
          "    ReorderedColumns = Table.ReorderColumns(PromotedHeaders, {IdColumn} & AutresColonnes)," & vbCrLf & _
          "    // Typer la colonne ID" & vbCrLf & _
          "    TypedTable = Table.TransformColumnTypes(ReorderedColumns,{{IdColumn, Int64.Type}})" & vbCrLf & _
          "in" & vbCrLf & _
          "    TypedTable"
    
    GeneratePQQueryTemplate = template
End Function


' Fonction pour stocker les types de colonnes d'une requête
Private Sub StoreColumnTypes(queryName As String)
    If mColumnTypes Is Nothing Then
        Set mColumnTypes = CreateObject("Scripting.Dictionary")
    End If
    
    On Error Resume Next
    ' Obtenir la référence à la table PowerQuery
    Dim connection As WorkbookConnection
    Set connection = ThisWorkbook.Connections(queryName)
    
    If Not connection Is Nothing Then
        ' Parcourir les colonnes et stocker leurs types
        Dim table As ListObject
        Set table = connection.QueryTable.ResultRange.ListObject
        
        Dim col As ListColumn
        For Each col In table.ListColumns
            ' Stocker le type de données de la colonne
            If Not mColumnTypes.Exists(queryName) Then
                Set mColumnTypes(queryName) = CreateObject("Scripting.Dictionary")
            End If
            mColumnTypes(queryName)(col.Name) = col.Range.Cells(2).NumberFormat
        Next col
    End If
    On Error GoTo 0
End Sub

' Fonction pour récupérer le type d'une colonne
Public Function GetStoredColumnType(queryName As String, columnName As String) As String
    If mColumnTypes Is Nothing Then Exit Function
    If Not mColumnTypes.Exists(queryName) Then Exit Function
    If Not mColumnTypes(queryName).Exists(columnName) Then Exit Function
    
    GetStoredColumnType = mColumnTypes(queryName)(columnName)
End Function

Public Sub RefreshAllPowerQueries(Optional ByVal showErrors As Boolean = True)
    Const PROC_NAME As String = "RefreshAllPowerQueries"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Starting refresh of all Power Queries.", PROC_NAME, MODULE_NAME
    
    ' Dim query As WorkbookQuery ' Original
    ' Dim connection As WorkbookConnection ' Original
    Dim errorCount As Long
    errorCount = 0

    ' On Error Resume Next ' Original blanket error handling
    
    LogDebug PROC_NAME & "_RefreshConnections", "Refreshing all workbook connections.", PROC_NAME, MODULE_NAME
    On Error Resume Next ' Handle errors per connection/query
    ThisWorkbook.RefreshAll
    If Err.Number <> 0 Then
        LogError PROC_NAME & "_RefreshAllError", Err.Number, "Error during ThisWorkbook.RefreshAll: " & Err.Description, PROC_NAME, MODULE_NAME
        If showErrors Then
            ElyseMessageBox_System.ShowErrorMessage "Refresh Error", "An error occurred during the global RefreshAll operation. Some queries or connections might not have refreshed. Details: " & Err.Description
        End If
        errorCount = errorCount + 1 ' Count this as one major error for RefreshAll
        Err.Clear
    Else
        LogInfo PROC_NAME & "_RefreshAllSuccess", "ThisWorkbook.RefreshAll completed.", PROC_NAME, MODULE_NAME
    End If
    On Error GoTo ErrorHandler ' Restore main error handler

    ' The above ThisWorkbook.RefreshAll should handle most Power Queries and connections.
    ' Individual refresh below might be redundant or for specific handling if RefreshAll is not sufficient or too broad.
    ' For simplicity and to avoid double-refreshing, we rely on RefreshAll.
    ' If individual refresh is still desired, the original loop structure can be adapted with new logging.

    ' Example of original individual refresh logic (commented out as RefreshAll is preferred):
    ' For Each query In ThisWorkbook.Queries
    '     Debug.Print "Refreshing query: " & query.Name
    '     query.Refresh
    '     If Err.Number <> 0 Then
    '         Debug.Print "Error refreshing query " & query.Name & ": " & Err.Description
    '         errorCount = errorCount + 1
    '         If showErrors Then MsgBox "Error refreshing query " & query.Name & ": " & vbCrLf & Err.Description, vbCritical
    '         Err.Clear
    '     End If
    ' Next query
    '
    ' For Each connection In ThisWorkbook.Connections
    '     If connection.Type = xlConnectionTypeOLEDB Or InStr(1, connection.OLEDBConnection.Connection, "Provider=Microsoft.Mashup.OleDb.1", vbTextCompare) > 0 Then
    '         Debug.Print "Refreshing connection: " & connection.Name
    '         connection.Refresh
    '         If Err.Number <> 0 Then
    '             Debug.Print "Error refreshing connection " & connection.Name & ": " & Err.Description
    '             errorCount = errorCount + 1
    '             If showErrors Then MsgBox "Error refreshing connection " & connection.Name & ": " & vbCrLf & Err.Description, vbCritical
    '             Err.Clear
    '         End If
    '     End If
    ' Next connection

    If errorCount = 0 Then
        LogInfo PROC_NAME & "_EndSuccess", "All Power Queries and connections refreshed successfully.", PROC_NAME, MODULE_NAME
        If showErrors Then ElyseMessageBox_System.ShowInfoMessage "Refresh Complete", "All Power Queries and connections have been refreshed."
    Else
        LogWarning PROC_NAME & "_EndWithErrors", errorCount & " error(s) occurred during refresh. Check logs for details.", PROC_NAME, MODULE_NAME
        If showErrors Then ElyseMessageBox_System.ShowWarningMessage "Refresh Complete with Errors", errorCount & " error(s) occurred during the refresh process. Please check the logs for more details."
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
    If showErrors Then ElyseMessageBox_System.ShowErrorMessage "Critical Refresh Error", "A critical error occurred in RefreshAllPowerQueries. Process aborted. Details: " & Err.Description
End Sub

Public Function GetQueryLastRefreshDate(ByVal queryName As String) As Date
    Const PROC_NAME As String = "GetQueryLastRefreshDate"
    On Error GoTo ErrorHandler
    
    LogDebug PROC_NAME & "_Start", "Getting last refresh date for query: " & queryName, PROC_NAME, MODULE_NAME
    
    Dim q As WorkbookQuery
    Dim c As WorkbookConnection
    Dim lastRefresh As Date
    lastRefresh = CDate(0) ' Default to a very old date

    On Error Resume Next ' Try to find it as a WorkbookQuery
    Set q = ThisWorkbook.Queries(queryName)
    If Not q Is Nothing Then
        lastRefresh = q.RefreshDate
        LogDebug PROC_NAME & "_QueryFound", "Query '" & queryName & "' found. Refresh Date: " & lastRefresh, PROC_NAME, MODULE_NAME
    Else
        Err.Clear
        Set c = ThisWorkbook.Connections(queryName)
        If Not c Is Nothing Then
            ' WorkbookConnection object does not have a direct RefreshDate property like WorkbookQuery.
            ' It might be available through OLEDBConnection.RefreshDate for some types, but not universally.
            ' For Power Query connections, the RefreshDate is typically on the WorkbookQuery object.
            ' We will log that it's not directly available for connections here.
            LogInfo PROC_NAME & "_ConnectionFoundNoDate", "Connection '" & queryName & "' found, but direct RefreshDate is not available on WorkbookConnection object. Check associated WorkbookQuery if it exists.", PROC_NAME, MODULE_NAME
            ' Attempt to find an associated query if the name matches
            Set q = Nothing ' Reset q
            Set q = ThisWorkbook.Queries(queryName)
            If Not q Is Nothing Then
                 lastRefresh = q.RefreshDate
                 LogDebug PROC_NAME & "_AssociatedQueryFound", "Associated Query '" & queryName & "' found. Refresh Date: " & lastRefresh, PROC_NAME, MODULE_NAME
            End If
        Else
            LogWarning PROC_NAME & "_NotFound", "Query or Connection '" & queryName & "' not found.", PROC_NAME, MODULE_NAME
            ' MsgBox "Query or Connection '" & queryName & "' not found.", vbInformation
            ElyseMessageBox_System.ShowInfoMessage "Query Info", "Query or Connection '" & queryName & "' not found."
        End If
    End If
    On Error GoTo ErrorHandler ' Restore error handling
    
    GetQueryLastRefreshDate = lastRefresh
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
    GetQueryLastRefreshDate = CDate(0) ' Return default on error
End Function

