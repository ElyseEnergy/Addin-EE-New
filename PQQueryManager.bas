Attribute VB_Name = "PQQueryManager"
' Module: PQQueryManager
' G�re la cr�ation et la v�rification des requ�tes PowerQuery
Option Explicit

Private mColumnTypes As Object ' Dictionnaire pour stocker les types de colonnes

' V�rifie si une requ�te PowerQuery existe et la cr�e si n�cessaire
Public Function EnsurePQQueryExists(Category As CategoryInfo) As Boolean
    Dim query As String
    query = GeneratePQQueryTemplate(Category)
    
    If QueryExists(Category.PowerQueryName) Then
        ' Si la requ�te existe, mettre � jour sa formule
        On Error Resume Next
        ThisWorkbook.Queries(Category.PowerQueryName).formula = query
        Dim updateError As Long
        updateError = Err.Number
        On Error GoTo 0
          If updateError <> 0 Then
            Log "pq_update", "Erreur lors de la mise � jour de la requ�te " & Category.PowerQueryName & ": " & Err.Description, ERROR_LEVEL, "EnsurePQQueryExists", "PQQueryManager"
            EnsurePQQueryExists = False
            Exit Function
        End If
        
        ' Rafra�chir la requ�te
        ThisWorkbook.Queries(Category.PowerQueryName).Refresh
        EnsurePQQueryExists = True
    Else
        ' Cr�er la requ�te si elle n'existe pas
        If Not AddQueryToPowerQuery(Category.PowerQueryName, query) Then
            EnsurePQQueryExists = False
            Exit Function
        End If
        EnsurePQQueryExists = True
    End If
    
    ' Stocker les types de colonnes apr�s la cr�ation de la requ�te
    StoreColumnTypes Category.PowerQueryName
    
    EnsurePQQueryExists = True
End Function

' V�rifie si une requ�te PowerQuery existe
Private Function QueryExists(queryName As String) As Boolean
    On Error Resume Next
    Dim query As Object
    Set query = ThisWorkbook.Queries(queryName)
    QueryExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' Ajoute une requ�te PowerQuery
Private Function AddQueryToPowerQuery(queryName As String, query As String) As Boolean
    On Error Resume Next
    ThisWorkbook.Queries.Add queryName, query
    Dim errNum As Long
    errNum = Err.Number
    If errNum <> 0 Then
        Log "pq_add", "Erreur lors de l'ajout de la requ�te " & queryName & ": " & Err.Description, ERROR_LEVEL, "AddQueryToPowerQuery", "PQQueryManager"
        AddQueryToPowerQuery = False
    Else
        AddQueryToPowerQuery = True
    End If
    On Error GoTo 0
End Function

' G�n�re le template de requ�te PowerQuery
Private Function GeneratePQQueryTemplate(Category As CategoryInfo) As String
    Dim template As String
    ' Template de base pour charger les donn�es depuis l'API Ragic avec r�organisation des colonnes
    template = "let" & vbCrLf & _
          "    Source = Csv.Document(Web.Contents(""" & Category.URL & """),[Delimiter="","",Encoding=65001,QuoteStyle=QuoteStyle.Csv])," & vbCrLf & _
          "    PromotedHeaders = Table.PromoteHeaders(Source)," & vbCrLf & _
          "    // Trouver la colonne ID quelle que soit sa casse" & vbCrLf & _
          "    Colonnes = Table.ColumnNames(PromotedHeaders)," & vbCrLf & _
          "    IdColumn = List.First(List.Select(Colonnes, each Text.Lower(_) = ""id""))," & vbCrLf & _
          "    AutresColonnes = List.Select(Colonnes, each Text.Lower(_) <> ""id"")," & vbCrLf & _
          "    // R�organiser les colonnes pour avoir ID en premier" & vbCrLf & _
          "    ReorderedColumns = Table.ReorderColumns(PromotedHeaders, {IdColumn} & AutresColonnes)," & vbCrLf & _
          "    // Typer la colonne ID" & vbCrLf & _
          "    TypedTable = Table.TransformColumnTypes(ReorderedColumns,{{IdColumn, Int64.Type}})" & vbCrLf & _
          "in" & vbCrLf & _
          "    TypedTable"
    
    GeneratePQQueryTemplate = template
End Function


' Fonction pour stocker les types de colonnes d'une requ�te
Private Sub StoreColumnTypes(queryName As String)
    If mColumnTypes Is Nothing Then
        Set mColumnTypes = CreateObject("Scripting.Dictionary")
    End If
    
    On Error Resume Next
    ' Obtenir la r�f�rence � la table PowerQuery
    Dim connection As WorkbookConnection
    Set connection = ThisWorkbook.Connections(queryName)
    
    If Not connection Is Nothing Then
        ' Parcourir les colonnes et stocker leurs types
        Dim table As ListObject
        Set table = connection.QueryTable.ResultRange.ListObject
        
        Dim col As ListColumn
        For Each col In table.ListColumns
            ' Stocker le type de donn�es de la colonne
            If Not mColumnTypes.Exists(queryName) Then
                Set mColumnTypes(queryName) = CreateObject("Scripting.Dictionary")
            End If
            mColumnTypes(queryName)(col.Name) = col.Range.Cells(2).NumberFormat
        Next col
    End If
    On Error GoTo 0
End Sub

' Fonction pour r�cup�rer le type d'une colonne
Public Function GetStoredColumnType(queryName As String, columnName As String) As String
    If mColumnTypes Is Nothing Then Exit Function
    If Not mColumnTypes.Exists(queryName) Then Exit Function
    If Not mColumnTypes(queryName).Exists(columnName) Then Exit Function
    
    GetStoredColumnType = mColumnTypes(queryName)(columnName)
End Function

