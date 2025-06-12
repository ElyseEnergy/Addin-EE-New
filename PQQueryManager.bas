Attribute VB_Name = "PQQueryManager"
' Module: PQQueryManager
' G�re la cr�ation et la v�rification des requ�tes PowerQuery
Option Explicit

Private mColumnTypes As Object ' Dictionnaire pour stocker les types de colonnes

' V�rifie si une requ�te PowerQuery existe et la cr�e si n�cessaire
Public Function EnsurePQQueryExists(Category As categoryInfo) As Boolean
    On Error GoTo ErrorHandler
    
    Dim newFormula As String
    newFormula = GeneratePQQueryTemplate(Category)
    
    Dim QueryExists As Boolean
    Dim needsUpdate As Boolean
    
    ' Vérifier si la requête existe dans le classeur actif de l'utilisateur
    On Error Resume Next
    Dim pq As Object ' WorkbookQuery
    Set pq = ActiveWorkbook.Queries(Category.PowerQueryName)
    QueryExists = (Err.Number = 0)
    On Error GoTo ErrorHandler

    If QueryExists Then
        ' La requ�te existe, v�rifier si la formule a chang�
        If pq.formula <> newFormula Then
            needsUpdate = True
            Diagnostics.LogTime "La formule de la requ�te '" & Category.PowerQueryName & "' a chang�. Mise � jour n�cessaire."
        Else
            needsUpdate = False
            Diagnostics.LogTime "La requ�te '" & Category.PowerQueryName & "' est d�j� � jour. Pas de modification."
        End If
    Else
        ' La requ�te n'existe pas, il faut la cr�er
        needsUpdate = True
        Diagnostics.LogTime "La requ�te '" & Category.PowerQueryName & "' n'existe pas. Cr�ation n�cessaire."
    End If
    
    If needsUpdate Then
        If QueryExists Then
            ' Mise à jour de la formule dans le classeur actif
            pq.formula = newFormula
        Else
            ' Ajout de la nouvelle requête dans le classeur actif
            ActiveWorkbook.Queries.Add Category.PowerQueryName, newFormula
        End If
    End If
    
    EnsurePQQueryExists = True
    Exit Function

ErrorHandler:
    Log "pq_error", "Erreur critique dans EnsurePQQueryExists pour " & Category.PowerQueryName & ": " & Err.Description, ERROR_LEVEL, "EnsurePQQueryExists", "PQQueryManager"
    EnsurePQQueryExists = False
End Function

' V�rifie si une requ�te PowerQuery existe
Public Function QueryExists(QueryName As String) As Boolean
    On Error Resume Next
    Dim query As Object
    Set query = ActiveWorkbook.Queries(QueryName)
    QueryExists = (Err.Number = 0)
    On Error GoTo 0
End Function

' Ajoute une requête PowerQuery dans le classeur actif
Public Function AddQueryToPowerQuery(QueryName As String, query As String) As Boolean
    On Error Resume Next
    ActiveWorkbook.Queries.Add QueryName, query
    Dim errNum As Long
    errNum = Err.Number
    If errNum <> 0 Then
        Log "pq_add", "Erreur lors de l'ajout de la requ�te " & QueryName & ": " & Err.Description, ERROR_LEVEL, "AddQueryToPowerQuery", "PQQueryManager"
        AddQueryToPowerQuery = False
    Else
        AddQueryToPowerQuery = True
    End If
    On Error GoTo 0
End Function

' G�n�re le template de requ�te PowerQuery
Private Function GeneratePQQueryTemplate(Category As categoryInfo) As String
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
Private Sub StoreColumnTypes(QueryName As String)
    If mColumnTypes Is Nothing Then
        Set mColumnTypes = CreateObject("Scripting.Dictionary")
    End If
    
    On Error Resume Next
    ' Obtenir la r�f�rence � la table PowerQuery
    Dim connection As WorkbookConnection
    Set connection = ThisWorkbook.Connections(QueryName)
    
    If Not connection Is Nothing Then
        ' Parcourir les colonnes et stocker leurs types
        Dim table As ListObject
        Set table = connection.QueryTable.ResultRange.ListObject
        
        Dim col As ListColumn
        For Each col In table.ListColumns
            ' Stocker le type de donn�es de la colonne
            If Not mColumnTypes.Exists(QueryName) Then
                Set mColumnTypes(QueryName) = CreateObject("Scripting.Dictionary")
            End If
            mColumnTypes(QueryName)(col.Name) = col.Range.Cells(2).NumberFormat
        Next col
    End If
    On Error GoTo 0
End Sub

' Fonction pour r�cup�rer le type d'une colonne
Public Function GetStoredColumnType(QueryName As String, columnName As String) As String
    If mColumnTypes Is Nothing Then Exit Function
    If Not mColumnTypes.Exists(QueryName) Then Exit Function
    If Not mColumnTypes(QueryName).Exists(columnName) Then Exit Function
    
    GetStoredColumnType = mColumnTypes(QueryName)(columnName)
End Function

