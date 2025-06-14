Attribute VB_Name = "PQQueryManager"
Option Explicit

Private mColumnTypes As Object ' Dictionnaire pour stocker les types de colonnes

' Vérifie si une requête PowerQuery existe et la crée si nécessaire
Public Function EnsurePQQueryExists(Category As CategoryInfo) As Boolean
    On Error GoTo ErrorHandler
    
    Dim newFormula As String
    newFormula = GeneratePQQueryTemplate(Category)
    
    Dim queryExists As Boolean
    Dim needsUpdate As Boolean
    
    ' Vérifier si la requête existe
    On Error Resume Next
    Dim pq As Object ' WorkbookQuery
    Set pq = ThisWorkbook.Queries(Category.PowerQueryName)
    queryExists = (Err.Number = 0)
    On Error GoTo ErrorHandler

    If queryExists Then
        ' La requête existe, vérifier si la formule a changé
        If pq.formula <> newFormula Then
            needsUpdate = True
            Diagnostics.LogTime "La formule de la requête '" & Category.PowerQueryName & "' a changé. Mise à jour nécessaire."
        Else
            needsUpdate = False
            Diagnostics.LogTime "La requête '" & Category.PowerQueryName & "' est déjà à jour. Pas de modification."
        End If
    Else
        ' La requête n'existe pas, il faut la créer
        needsUpdate = True
        Diagnostics.LogTime "La requête '" & Category.PowerQueryName & "' n'existe pas. Création nécessaire."
    End If
    
    If needsUpdate Then
        If queryExists Then
            ' Mise à jour de la formule
            pq.formula = newFormula
        Else
            ' Ajout de la nouvelle requête
            ThisWorkbook.Queries.Add Category.PowerQueryName, newFormula
        End If
    End If
    
    EnsurePQQueryExists = True
    Exit Function

ErrorHandler:
    Log "pq_error", "Erreur critique dans EnsurePQQueryExists pour " & Category.PowerQueryName & ": " & Err.Description, ERROR_LEVEL, "EnsurePQQueryExists", "PQQueryManager"
    EnsurePQQueryExists = False
End Function

' Vérifie si une requête PowerQuery existe
Public Function QueryExists(queryName As String) As Boolean
    Const PROC_NAME As String = "QueryExists"
    Const MODULE_NAME As String = "PQQueryManager"
    On Error GoTo ErrorHandler
    
    Dim query As Object
    Set query = ThisWorkbook.Queries(queryName)
    QueryExists = (Err.Number = 0)
    Exit Function

ErrorHandler:
    QueryExists = False
    ' Pas de log ici, la fonction est juste une vérification silencieuse
End Function

' Ajoute une requête PowerQuery
Public Function AddQueryToPowerQuery(queryName As String, query As String) As Boolean
    Const PROC_NAME As String = "AddQueryToPowerQuery"
    Const MODULE_NAME As String = "PQQueryManager"
    On Error GoTo ErrorHandler

    ThisWorkbook.Queries.Add queryName, query
    AddQueryToPowerQuery = True
    Exit Function

ErrorHandler:
    Log "pq_add_error", "Erreur lors de l'ajout de la requête " & queryName & ": " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    AddQueryToPowerQuery = False
End Function

' Génère le template de requête PowerQuery
Private Function GeneratePQQueryTemplate(Category As CategoryInfo) As String
    Const PROC_NAME As String = "GeneratePQQueryTemplate"
    Const MODULE_NAME As String = "PQQueryManager"
    On Error GoTo ErrorHandler
    Dim template As String
    ' Template de base pour charger les données depuis l'API Ragic avec réorganisation des colonnes
    template = "let" & vbCrLf & _
          "    Source = Csv.Document(Web.Contents(""" & Category.URL & """),[Delimiter="","",Encoding=65001,QuoteStyle=QuoteStyle.Csv])," & vbCrLf & _
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
    Exit Function
ErrorHandler:
    SYS_Logger.Log "pq_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la génération du template PQ."
    GeneratePQQueryTemplate = ""
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
    Const PROC_NAME As String = "GetStoredColumnType"
    Const MODULE_NAME As String = "PQQueryManager"
    On Error GoTo ErrorHandler

    If mColumnTypes Is Nothing Then Exit Function
    If Not mColumnTypes.Exists(queryName) Then Exit Function
    If Not mColumnTypes(queryName).Exists(columnName) Then Exit Function
    
    GetStoredColumnType = mColumnTypes(queryName)(columnName)
    Exit Function

ErrorHandler:
    GetStoredColumnType = "" ' Retourner une chaîne vide en cas d'erreur
    SYS_Logger.Log "pq_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Failed to get stored column type for " & queryName & ":" & columnName
End Function

