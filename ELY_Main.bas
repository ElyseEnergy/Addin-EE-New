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
    Dim modeTransposed As Boolean
    Dim previewNormal As String, previewTransposed As String
    Dim nbFiches As Long, nbChamps As Long
    Dim previewRows As Long: previewRows = 3
    Dim previewCols As Long
    Dim okPlage As Boolean
    Dim categoryManager As New CategoryManager

    ' Initialiser la feuille PQ_DATA si besoin
    If wsPQData Is Nothing Then Utilities.InitializePQData

    ' 1. Charger la table des marques et demander le choix
    lastCol = Utilities.GetLastColumn(wsPQData)
    LoadQueries.LoadQuery "01_ELY_Brands", wsPQData, wsPQData.Cells(1, lastCol + 1)
    Set selectedBrands = LoadQueries.ChooseMultipleValuesFromTableWithAll(wsPQData, "Table_01_ELY_Brands", "Brand", "Choisissez une ou plusieurs marques (ex: 1,3,5 ou *) :")
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
    lastCol = Utilities.GetLastColumn(wsPQData)
    LoadQueries.LoadQuery "02_ELY_List_filtered", wsPQData, wsPQData.Cells(1, lastCol + 1)

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
    Set selectedFicheIds = LoadQueries.ChooseMultipleValuesFromListWithAll(idList, nameList, "Choisissez une ou plusieurs fiches (ex: 1,2,5 ou *) :")
    If selectedFicheIds Is Nothing Or selectedFicheIds.Count = 0 Then
        MsgBox "Aucune fiche sélectionnée. Opération annulée.", vbExclamation
        Exit Sub
    End If

    ' 6. Préparer les exemples pour l'inputbox de mode
    nbFiches = selectedFicheIds.Count
    nbChamps = lo.ListColumns.Count
    previewCols = WorksheetFunction.Min(nbFiches, previewRows)
    previewNormal = "Mode NORMAL (tableau classique) :" & vbCrLf
    previewTransposed = "Mode TRANSPOSE (fiches en colonnes) :" & vbCrLf

    ' Extrait les 3 premières lignes pour l'aperçu normal (max 4 colonnes, 10 caractères)
    previewNormal = previewNormal & "| "
    For i = 1 To WorksheetFunction.Min(4, nbChamps)
        previewNormal = previewNormal & Left(lo.HeaderRowRange.Cells(1, i).Value, 10) & " | "
    Next i
    previewNormal = previewNormal & vbCrLf
    Dim idx As Long, j As Long
    idx = 1
    For Each v In selectedFicheIds
        If idx > previewRows Then Exit For
        ' Trouver la ligne correspondante
        For j = 1 To lo.DataBodyRange.Rows.Count
            If lo.DataBodyRange.Rows(j).Columns(1).Value = v Then
                previewNormal = previewNormal & "| "
                For i = 1 To WorksheetFunction.Min(4, nbChamps)
                    previewNormal = previewNormal & Left(lo.DataBodyRange.Rows(j).Cells(1, i).Value, 10) & " | "
                Next i
                previewNormal = previewNormal & vbCrLf
                Exit For
            End If
        Next j
        idx = idx + 1
    Next v

    ' Extrait les 3 premières fiches pour l'aperçu transposé (max 4 champs, 10 caractères)
    previewTransposed = previewTransposed & "(en-têtes en ligne, fiches en colonnes)" & vbCrLf
    For i = 1 To WorksheetFunction.Min(4, nbChamps)
        previewTransposed = previewTransposed & Left(lo.HeaderRowRange.Cells(1, i).Value, 10) & ": "
        idx = 1
        For Each v In selectedFicheIds
            If idx > previewRows Then Exit For
            For j = 1 To lo.DataBodyRange.Rows.Count
                If lo.DataBodyRange.Rows(j).Columns(1).Value = v Then
                    previewTransposed = previewTransposed & Left(lo.DataBodyRange.Rows(j).Cells(1, i).Value, 10) & ", "
                    Exit For
                End If
            Next j
            idx = idx + 1
        Next v
        previewTransposed = previewTransposed & vbCrLf
    Next i

    ' 7. Demander le mode à l'utilisateur
    Dim modePrompt As String
    modePrompt = "Comment souhaitez-vous coller les fiches ?" & vbCrLf & vbCrLf & previewNormal & vbCrLf & previewTransposed & vbCrLf & "Tapez 1 pour NORMAL, 2 pour TRANSPOSE"
    userChoice = InputBox(modePrompt, "Choix du mode de collage", "1")
    If userChoice = "2" Then
        modeTransposed = True
    ElseIf userChoice = "1" Then
        modeTransposed = False
    Else
        MsgBox "Choix invalide. Opération annulée.", vbExclamation
        Exit Sub
    End If

    ' 8. Calculer la taille nécessaire et informer l'utilisateur
    Dim nbRows As Long, nbCols As Long
    If modeTransposed Then
        nbRows = nbChamps
        nbCols = nbFiches + 1 ' +1 pour les en-têtes
    Else
        nbRows = nbFiches + 1 ' +1 pour les en-têtes
        nbCols = nbChamps
    End If
    MsgBox "La plage nécessaire sera de " & nbRows & " lignes x " & nbCols & " colonnes.", vbInformation

    ' 9. Demander la cellule de destination et vérifier la place
    Do
        Set finalDestination = Application.InputBox("Sélectionnez la cellule où charger les fiches (" & nbRows & " x " & nbCols & ")", "Destination", Type:=8)
        If finalDestination Is Nothing Then
            MsgBox "Aucune destination sélectionnée. Opération annulée.", vbExclamation
            Exit Sub
        End If
        okPlage = True
        For i = 0 To nbRows - 1
            For j = 0 To nbCols - 1
                If Not IsEmpty(finalDestination.Offset(i, j)) Then
                    okPlage = False
                    Exit For
                End If
            Next j
            If Not okPlage Then Exit For
        Next i
        If Not okPlage Then
            MsgBox "La plage sélectionnée n'est pas vide. Veuillez choisir un autre emplacement.", vbExclamation
        End If
    Loop Until okPlage

    ' 10. Coller les données selon le mode choisi, appliquer les formats et créer un tableau
    Dim tblRange As Range
    If modeTransposed Then
        ' Coller en transposé
        For i = 1 To nbChamps
            finalDestination.Offset(i - 1, 0).Value = lo.HeaderRowRange.Cells(1, i).Value
            finalDestination.Offset(i - 1, 0).NumberFormat = lo.DataBodyRange.Columns(i).Cells(1, 1).NumberFormat
        Next i
        Dim currentCol As Long: currentCol = 1
        For Each v In selectedFicheIds
            For j = 1 To lo.DataBodyRange.Rows.Count
                If lo.DataBodyRange.Rows(j).Columns(1).Value = v Then
                    For i = 1 To nbChamps
                        finalDestination.Offset(i - 1, currentCol).Value = lo.DataBodyRange.Rows(j).Cells(1, i).Value
                        finalDestination.Offset(i - 1, currentCol).NumberFormat = lo.DataBodyRange.Rows(j).Cells(1, i).NumberFormat
                    Next i
                    Exit For
                End If
            Next j
            currentCol = currentCol + 1
        Next v
        Set tblRange = finalDestination.Resize(nbRows, nbCols)
    Else
        ' Coller en normal
        For i = 1 To nbChamps
            finalDestination.Offset(0, i - 1).Value = lo.HeaderRowRange.Cells(1, i).Value
            finalDestination.Offset(0, i - 1).NumberFormat = lo.DataBodyRange.Columns(i).Cells(1, 1).NumberFormat
        Next i
        Dim currentRow As Long: currentRow = 1
        For Each v In selectedFicheIds
            For j = 1 To lo.DataBodyRange.Rows.Count
                If lo.DataBodyRange.Rows(j).Columns(1).Value = v Then
                    For i = 1 To nbChamps
                        finalDestination.Offset(currentRow, i - 1).Value = lo.DataBodyRange.Rows(j).Cells(1, i).Value
                        finalDestination.Offset(currentRow, i - 1).NumberFormat = lo.DataBodyRange.Rows(j).Cells(1, i).NumberFormat
                    Next i
                    Exit For
                End If
            Next j
            currentRow = currentRow + 1
        Next v
        Set tblRange = finalDestination.Resize(nbRows, nbCols)
    End If

    ' 11. Mettre en forme le tableau final
    Dim ws As Worksheet
    Set ws = finalDestination.Worksheet
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    On Error GoTo 0
    If Not tbl Is Nothing Then
        tbl.TableStyle = "TableStyleMedium9" ' ou un autre style de ton choix
    End If

    ' 12. Protéger la feuille et verrouiller les cellules du tableau
    tblRange.Locked = True
    ws.Protect Password:="elyse", AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
    MsgBox "Collage terminé et protégé !", vbInformation

End Sub