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
    Dim categoryManager As New categoryManager
    Dim dict As Object
    Dim arrBrands() As String
    Dim j As Long

    ' Initialiser la feuille PQ_DATA si besoin
    If wsPQData Is Nothing Then Utilities.InitializePQData

    ' 1. Charger la table des fiches et extraire les marques uniques
    lastCol = Utilities.GetLastColumn(wsPQData)
    LoadQueries.LoadQuery "02_ELY_List_filtered", wsPQData, wsPQData.Cells(1, lastCol + 1)
    
    ' Créer un dictionnaire pour stocker les marques uniques
    Set dict = CreateObject("Scripting.Dictionary")
    Set lo = wsPQData.ListObjects("Table_02_ELY_List_filtered")
    
    ' Extraire les marques uniques
    For Each cellBrand In lo.ListColumns("Brand").DataBodyRange
        If Not dict.exists(cellBrand.Value) Then
            dict.Add cellBrand.Value, 1
        End If
    Next cellBrand
    
    ' Convertir le dictionnaire en tableau et trier
    ReDim arrBrands(1 To dict.Count)
    i = 1
    For Each v In dict.keys
        arrBrands(i) = v
        i = i + 1
    Next v
    
    ' Trier le tableau
    For i = 1 To UBound(arrBrands) - 1
        For j = i + 1 To UBound(arrBrands)
            If arrBrands(i) > arrBrands(j) Then
                Dim temp As String
                temp = arrBrands(i)
                arrBrands(i) = arrBrands(j)
                arrBrands(j) = temp
            End If
        Next j
    Next i
    
    ' Présenter les marques à l'utilisateur
    Set selectedBrands = LoadQueries.ChooseMultipleValuesFromArrayWithAll(arrBrands, "Choisissez une ou plusieurs marques (ex: 1,3,5 ou *) :")
    If (TypeName(selectedBrands) <> "Collection") Then
        MsgBox "Aucune marque sélectionnée. Opération annulée.", vbExclamation
        Exit Sub
    End If
    If selectedBrands.Count = 0 Then
        MsgBox "Aucune marque sélectionnée. Opération annulée.", vbExclamation
        Exit Sub
    End If

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
        MsgBox "No fiche found for this brand.", vbExclamation
        Exit Sub
    End If

    ' 5. Sélection multiple des fiches techniques (avec *)
    Set selectedFicheIds = LoadQueries.ChooseMultipleValuesFromListWithAll(idList, nameList, "Choisissez a fiche or fiches (ex: 1,2,5 or *) :")
    If (TypeName(selectedFicheIds) <> "Collection") Then
        MsgBox "Aucune fiche sélectionnée. Opération annulée.", vbExclamation
        Exit Sub
    End If
    If selectedFicheIds.Count = 0 Then
        MsgBox "Aucune fiche sélectionnée. Opération annulée.", vbExclamation
        Exit Sub
    End If

    ' 6. Préparer les exemples pour l'inputbox de mode
    nbFiches = selectedFicheIds.Count
    nbChamps = lo.ListColumns.Count
    previewCols = WorksheetFunction.Min(nbFiches, previewRows)
    previewNormal = "Mode NORMAL (tableau classique) :" & vbCrLf
    previewTransposed = "Mode TRANSPOSE (fiches en colonnes) :" & vbCrLf



    ' --- Aligned NORMAL preview generation ---
    Dim colWidths(1 To 4) As Integer
    For i = 1 To WorksheetFunction.Min(4, nbChamps)
        colWidths(i) = Len(TruncateWithEllipsis(lo.HeaderRowRange.Cells(1, i).Value, 10))
    Next i
    idx = 1
    For Each v In selectedFicheIds
        If idx > previewRows Then Exit For
        For j = 1 To lo.DataBodyRange.Rows.Count
            If lo.DataBodyRange.Rows(j).Columns(1).Value = v Then
                For i = 1 To WorksheetFunction.Min(4, nbChamps)
                    Dim val As String
                    val = TruncateWithEllipsis(lo.DataBodyRange.Rows(j).Cells(1, i).Value, 10)
                    If Len(val) > colWidths(i) Then colWidths(i) = Len(val)
                Next i
                Exit For
            End If
        Next j
        idx = idx + 1
    Next v
    previewNormal = previewNormal & "| "
    For i = 1 To WorksheetFunction.Min(4, nbChamps)
        Dim head As String
        head = TruncateWithEllipsis(lo.HeaderRowRange.Cells(1, i).Value, 10)
        previewNormal = previewNormal & head & Space(colWidths(i) - Len(head)) & " | "
    Next i
    previewNormal = previewNormal & vbCrLf
    idx = 1
    For Each v In selectedFicheIds
        If idx > previewRows Then Exit For
        For j = 1 To lo.DataBodyRange.Rows.Count
            If lo.DataBodyRange.Rows(j).Columns(1).Value = v Then
                previewNormal = previewNormal & "| "
                For i = 1 To WorksheetFunction.Min(4, nbChamps)
                    val = TruncateWithEllipsis(lo.DataBodyRange.Rows(j).Cells(1, i).Value, 10)
                    previewNormal = previewNormal & val & Space(colWidths(i) - Len(val)) & " | "
                Next i
                previewNormal = previewNormal & vbCrLf
                Exit For
            End If
        Next j
        idx = idx + 1
    Next v

    ' --- Aligned TRANSPOSED preview generation ---
    Dim rowWidths(1 To 4) As Integer
    For i = 1 To WorksheetFunction.Min(4, nbChamps)
        rowWidths(i) = Len(TruncateWithEllipsis(lo.HeaderRowRange.Cells(1, i).Value, 10))
        idx = 1
        For Each v In selectedFicheIds
            If idx > previewRows Then Exit For
            For j = 1 To lo.DataBodyRange.Rows.Count
                If lo.DataBodyRange.Rows(j).Columns(1).Value = v Then
                    val = TruncateWithEllipsis(lo.DataBodyRange.Rows(j).Cells(1, i).Value, 10)
                    If Len(val) > rowWidths(i) Then rowWidths(i) = Len(val)
                    Exit For
                End If
            Next j
            idx = idx + 1
        Next v
    Next i
    previewTransposed = previewTransposed & "(headers in row, sheets in columns)" & vbCrLf
    For i = 1 To WorksheetFunction.Min(4, nbChamps)
        Dim headT As String
        headT = TruncateWithEllipsis(lo.HeaderRowRange.Cells(1, i).Value, 10)
        previewTransposed = previewTransposed & headT & Space(rowWidths(i) - Len(headT)) & ": "
        idx = 1
        For Each v In selectedFicheIds
            If idx > previewRows Then Exit For
            For j = 1 To lo.DataBodyRange.Rows.Count
                If lo.DataBodyRange.Rows(j).Columns(1).Value = v Then
                    Dim valT As String
                    valT = TruncateWithEllipsis(lo.DataBodyRange.Rows(j).Cells(1, i).Value, 10)
                    previewTransposed = previewTransposed & valT & Space(rowWidths(i) - Len(valT)) & ", "
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
        MsgBox "Invalid choice. Operation cancelled.", vbExclamation
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
            MsgBox "No destination selected. Operation cancelled.", vbExclamation
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

    ' 12. Protection modifiée de la feuille
    ' Déprotéger la feuille si besoin
    Set ws = finalDestination.Worksheet
    If ws.ProtectContents Then ws.Unprotect Password:="elyse"

    ' Sauvegarder l'état de verrouillage initial
    Dim lockedCells As Collection
    Set lockedCells = New Collection
    Dim cell As Range
    For Each cell In ws.UsedRange
        If cell.Locked Then
            lockedCells.Add cell.Address
        End If
    Next cell

    ' Déverrouiller toutes les cellules d'abord
    ws.Cells.Locked = False
    
    ' 11. Mettre en forme le tableau final
    Set tbl = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.TableStyle = "TableStyleMedium9" ' ou un autre style de ton choix
    
    ' Restaurer l'état de verrouillage initial
    Dim cellAddress As Variant
    For Each cellAddress In lockedCells
        ws.Range(cellAddress).Locked = True
    Next cellAddress
    
    ' Verrouiller les cellules du tableau
    tblRange.Locked = True
    
    ' Protéger la feuille à la toute fin avec les options souhaitées
    ws.Protect Password:="elyse", _
        AllowFormattingCells:=True, _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, _
        AllowInsertingColumns:=True, _
        AllowInsertingRows:=True, _
        AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, _
        AllowSorting:=True, _
        AllowFiltering:=True, _
        AllowUsingPivotTables:=True

    MsgBox "Collage terminé ! Les données sont protégées mais vous pouvez modifier la mise en forme.", vbInformation

End Sub




