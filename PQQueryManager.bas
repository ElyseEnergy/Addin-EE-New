' Module: PQQueryManager
' Gère la création et la vérification des requêtes PowerQuery
Option Explicit

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
    
    ' Template de base pour toutes les requêtes
    template = "let" & vbCrLf & _
               "    Source = Excel.CurrentWorkbook(){[Name=""Table_02_ELY_List_filtered""]}[Content]," & vbCrLf
    
    ' Ajouter le filtrage selon le niveau de filtrage
    If category.FilterLevel <> "Pas de filtrage" Then
        template = template & _
                  "    FilteredRows = Table.SelectRows(Source, each List.Contains({" & _
                  GetFilterValuesString(category.FilterLevel) & "}, [" & category.FilterLevel & "]))" & vbCrLf & _
                  "in" & vbCrLf & _
                  "    FilteredRows"
    Else
        template = template & _
                  "in" & vbCrLf & _
                  "    Source"
    End If
    
    GeneratePQQueryTemplate = template
End Function

' Génère la chaîne de valeurs de filtrage
Private Function GetFilterValuesString(filterLevel As String) As String
    ' TODO: Implémenter la récupération des valeurs de filtrage
    ' Pour l'instant, retourne une chaîne vide
    GetFilterValuesString = ""
End Function 