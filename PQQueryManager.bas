' Module: PQQueryManager
' Gère la création et la vérification des requêtes PowerQuery
Option Explicit

Private mColumnTypes As Object ' Dictionnaire pour stocker les types de colonnes

' Vérifie si une requête PowerQuery existe et la crée si nécessaire
Public Function EnsurePQQueryExists(category As CategoryInfo) As Boolean
    ' Vérifier si la requête existe déjà
    If QueryExists(category.PowerQueryName) Then
        EnsurePQQueryExists = True
        Exit Function
    End If
    
    ' Créer la requête si elle n'existe pas
    Dim query As String
    query = GeneratePQQueryTemplate(category)
      ' Ajouter la requête à PowerQuery
    If Not AddQueryToPowerQuery(category.PowerQueryName, query) Then
        EnsurePQQueryExists = False
        Exit Function
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
    AddQueryToPowerQuery = (Err.Number = 0)
    On Error GoTo 0
End Function

' Génère le template de requête PowerQuery
Private Function GeneratePQQueryTemplate(category As CategoryInfo) As String
    Dim template As String
    
    ' Template de base pour charger les données depuis l'API Ragic    template = "let" & vbCrLf & _
               "    Source = Csv.Document(Web.Contents(""" & category.URL & """)," & vbCrLf & _
               "        [Delimiter="","", Encoding=65001, QuoteStyle=QuoteStyle.Csv])," & vbCrLf & _
               "    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & vbCrLf & _
               "in" & vbCrLf & _
               "    PromotedHeaders"
    
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