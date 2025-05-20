Sub ELY_Main()
    Dim lastCol As Long
    Dim selectedBrands As Collection
    Dim selectedFicheIds As Collection
    Dim finalDestination As Range
    Dim lo As ListObject
    Dim idList As Collection, nameList As Collection
    Dim i As Long, userChoice As String
    Dim cellId As Range, cellBrand As Range, cellName As Range
    Dim foundRow As Range
    Dim v As Variant

    ' Initialiser la feuille PQ_DATA si besoin
    If wsPQData Is Nothing Then InitializePQData
    
    ' 1. Charger la table des marques et demander le choix
    lastCol = GetLastColumn(wsPQData)
    LoadQuery "01_ELY_Brands", wsPQData, wsPQData.Cells(1, lastCol + 1)
    Set selectedBrands = ChooseMultipleValuesFromTableWithAll(wsPQData, "Table_01_ELY_Brands", "Brand", "Choisissez une ou plusieurs marques (ex: 1,3,5 ou *) :")
    If selectedBrands Is Nothing Or selectedBrands.Count = 0 Then
        MsgBox "Aucune marque sélectionnée. Opération annulée.", vbExclamation
        Exit Sub
    End If
    
    ' 2. Supprimer la table existante si elle existe
    On Error Resume Next
    Set lo = wsPQData.ListObjects("Table_02_ELY_List_filtered")
    If Not lo Is Nothing Then lo.Delete
    On Error GoTo 0

    ' 3. Charger la table des fiches SANS filtre (toutes les fiches)
    lastCol = GetLastColumn(wsPQData)
    LoadQuery "02_ELY_List_filtered", wsPQData, wsPQData.Cells(1, lastCol + 1)

    ' 4. Proposer la sélection à l'utilisateur sur les fiches filtrées par la marque
    Set lo = wsPQData.ListObjects("Table_02_ELY_List_filtered")
    Set idList = New Collection
    Set nameList = New Collection
    i = 1
    For Each cellBrand In lo.ListColumns("Brand").DataBodyRange
        For Each v In selectedBrands
            If cellBrand.Value = v Then
                Set cellId = lo.ListColumns("id").DataBodyRange.Cells(i, 1)
                Set cellName = lo.ListColumns("Name").DataBodyRange.Cells(i, 1)
                idList.Add cellId.Value
                nameList.Add cellName.Value
                Exit For
            End If
        Next v
        i = i + 1
    Next cellBrand

    If idList.Count = 0 Then
        MsgBox "Aucune fiche trouvée pour cette marque.", vbExclamation
        Exit Sub
    End If

    ' 5. Sélection multiple des fiches techniques (avec *)
    Set selectedFicheIds = ChooseMultipleValuesFromListWithAll(idList, nameList, "Choisissez une ou plusieurs fiches (ex: 1,2,5 ou *) :")
    If selectedFicheIds Is Nothing Or selectedFicheIds.Count = 0 Then
        MsgBox "Aucune fiche sélectionnée. Opération annulée.", vbExclamation
        Exit Sub
    End If

    ' 6. Demander où charger la fiche finale
    Set finalDestination = Application.InputBox("Sélectionnez la cellule où charger la fiche finale", "Destination", Type:=8)
    If finalDestination Is Nothing Then
        MsgBox "Aucune destination sélectionnée. Opération annulée.", vbExclamation
        Exit Sub
    End If

    ' 7. Copier les lignes correspondant aux ids choisis (transposé, avec formats)
    Dim firstRow As Boolean
    Dim currentRowOffset As Long
    Dim colCount As Long
    Dim destCell As Range
    Dim srcCell As Range
    Dim destRow As Long, destCol As Long
    firstRow = True
    currentRowOffset = 0
    colCount = lo.ListColumns.Count
    For Each v In selectedFicheIds
        Set foundRow = Nothing
        i = 1
        For Each cellId In lo.ListColumns("id").DataBodyRange
            If cellId.Value = v Then
                Set foundRow = lo.DataBodyRange.Rows(i)
                Exit For
            End If
            i = i + 1
        Next cellId
        If Not foundRow Is Nothing Then
            If firstRow Then
                ' Coller les en-têtes en colonne (transposé)
                For destRow = 1 To colCount
                    finalDestination.Offset(destRow - 1, 0).Value = lo.HeaderRowRange.Cells(1, destRow).Value
                    ' Appliquer le format de la colonne source à la cellule de destination
                    finalDestination.Offset(destRow - 1, 0).NumberFormat = lo.DataBodyRange.Columns(destRow).Cells(1, 1).NumberFormat
                Next destRow
                firstRow = False
            End If
            ' Coller la ligne de données en colonne (transposé)
            For destRow = 1 To colCount
                finalDestination.Offset(destRow - 1, currentRowOffset + 1).Value = foundRow.Cells(1, destRow).Value
                ' Appliquer le format de la colonne source à la cellule de destination
                finalDestination.Offset(destRow - 1, currentRowOffset + 1).NumberFormat = foundRow.Cells(1, destRow).NumberFormat
            Next destRow
            currentRowOffset = currentRowOffset + 1
        End If
    Next v
End Sub

