Sub LoadQuery(QueryName As String, ws As Worksheet, DestCell As Range)
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "LoadQuery"
    Const MODULE_NAME As String = "LoadQueries"
    
    If QueryName = "" Then
        HandleError MODULE_NAME, PROC_NAME, "Nom de requête vide"
        Exit Sub
    End If
      If ws Is Nothing Then
        HandleError MODULE_NAME, PROC_NAME, "Feuille de calcul non spécifiée"
        Exit Sub
    End If
      If DestCell Is Nothing Then
        HandleError MODULE_NAME, PROC_NAME, "Cellule de destination non spécifiée"
        Exit Sub
    End If
    
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
    
    ' Après le chargement de la requête, s'assurer que le nom est correct
    Set lo = ws.ListObjects(ws.ListObjects.Count) ' Le dernier tableau créé
    If Not lo Is Nothing Then
        lo.Name = sanitizedName
    End If
    Exit Sub
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du chargement de la requête: " & QueryName
End Sub

Function ChooseMultipleValuesFromListWithAll(idList As Collection, displayList As Collection, prompt As String) As Collection
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "ChooseMultipleValuesFromListWithAll"
    Const MODULE_NAME As String = "LoadQueries"
    
    If idList Is Nothing Or displayList Is Nothing Then
        HandleError MODULE_NAME, PROC_NAME, "Listes non initialisées"
        Exit Function
    End If
    If idList.Count <> displayList.Count Then
        HandleError MODULE_NAME, PROC_NAME, "Les listes n'ont pas la même taille"
        Exit Function
    End If
    
    ' Utiliser le nouveau formulaire de sélection
    Dim frm As New FilterSelectionForm
    frm.InitializeWithList displayList, prompt
    frm.Show vbModal
    
    ' Récupérer la sélection
    Dim selectedItems As Collection
    Set selectedItems = frm.GetSelectedItems
    
    ' Mapper les éléments sélectionnés avec leurs IDs
    Dim selectedValues As New Collection
    If Not selectedItems Is Nothing Then
        Dim selectedItem As Variant
        For Each selectedItem In selectedItems
            Dim i As Long
            For i = 1 To displayList.Count
                If displayList(i) = selectedItem Then
                    selectedValues.Add idList(i)
                    Exit For
                End If
            Next i
        Next selectedItem
    End If
    
    Set ChooseMultipleValuesFromListWithAll = selectedValues
    Unload frm
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la sélection des valeurs"
End Function

Function ChooseMultipleValuesFromArrayWithAll(values() As String, prompt As String) As Collection
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "ChooseMultipleValuesFromArrayWithAll"
    Const MODULE_NAME As String = "LoadQueries"
    
    If Not IsArray(values) Then
        HandleError MODULE_NAME, PROC_NAME, "Tableau non initialisé"
        Exit Function
    End If
    If UBound(values) < 1 Then
        HandleError MODULE_NAME, PROC_NAME, "Tableau vide"
        Exit Function
    End If
    
    ' Convertir le tableau en Collection pour l'affichage
    Dim displayList As New Collection
    Dim i As Long
    For i = 1 To UBound(values)
        displayList.Add values(i)
    Next i
    
    ' Utiliser le nouveau formulaire de sélection
    Dim frm As New FilterSelectionForm
    frm.InitializeWithList displayList, prompt
    frm.Show vbModal
    
    ' Récupérer la sélection
    Dim selectedValues As New Collection
    Dim selectedItems As Collection
    Set selectedItems = frm.GetSelectedItems
    
    ' Ajouter les valeurs sélectionnées à la collection de retour
    If Not selectedItems Is Nothing Then
        Dim selectedItem As Variant
        For Each selectedItem In selectedItems
            selectedValues.Add selectedItem
        Next selectedItem
    End If
    
    Set ChooseMultipleValuesFromArrayWithAll = selectedValues
    Unload frm
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la sélection des valeurs"
End Function
