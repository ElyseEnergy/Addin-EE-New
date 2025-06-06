Attribute VB_Name = "DataLoaderManager"
' Module: DataLoaderManager
' G�re le chargement et l'affichage des donn�es pour toutes les cat�gories
Option Explicit

Public Enum DataLoadResult
    Success = 1
    Cancelled = 2
    Error = 3
End Enum

' Fonction principale de traitement
Public Function ProcessDataLoad(loadInfo As DataLoadInfo) As DataLoadResult
    ' Initialiser la feuille PQ_DATA si besoin
    If wsPQData Is Nothing Then Utilities.InitializePQData
    
    ' 1. V�rifier/Cr�er la requ�te PQ
    If Not PQQueryManager.EnsurePQQueryExists(loadInfo.Category) Then
        MsgBox "Erreur lors de la cr�ation de la requ�te PowerQuery", vbExclamation
        ProcessDataLoad = DataLoadResult.Error
        Exit Function
    End If
    
    ' 2. Charger les donn�es
    Dim lastCol As Long
    lastCol = Utilities.GetLastColumn(wsPQData)
    LoadQueries.LoadQuery loadInfo.Category.PowerQueryName, wsPQData, wsPQData.Cells(1, lastCol + 1)
    
    ' 3. G�rer la s�lection des valeurs
    Set loadInfo.SelectedValues = GetSelectedValues(loadInfo.Category)
    If loadInfo.SelectedValues Is Nothing Then
        ' Nettoyer la requ�te avant de sortir
        CleanupPowerQuery loadInfo.Category.PowerQueryName
        ProcessDataLoad = DataLoadResult.Cancelled
        Exit Function
    End If
    
    ' Si un filtre est appliqu� et qu'il n'y a pas de filtre secondaire,
    ' proposer la s�lection des fiches correspondantes
    If loadInfo.Category.FilterLevel <> "Pas de filtrage" And loadInfo.Category.SecondaryFilterLevel = "" Then
        Dim lo As ListObject
        Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
        Dim idList As New Collection
        Dim displayList As New Collection
        Dim i As Long, v As Variant
        Dim displayColIndex As Long
        displayColIndex = 2 ' Afficher la colonne 2 (nom)
        ' Parcourir les lignes et ne garder que celles correspondant au(x) filtre(s) choisi(s)
        For i = 1 To lo.DataBodyRange.Rows.Count
            For Each v In loadInfo.SelectedValues
                If lo.DataBodyRange.Rows(i).Columns(lo.ListColumns(loadInfo.Category.FilterLevel).Index).Value = v Then
                    idList.Add lo.DataBodyRange.Rows(i).Columns(1).Value
                    displayList.Add lo.DataBodyRange.Rows(i).Columns(displayColIndex).Value
                End If
            Next v
        Next i        ' Proposer la s�lection des fiches parmi displayList
        Dim finalSelection As Collection
        On Error Resume Next
        Set finalSelection = LoadQueries.ChooseMultipleValuesFromListWithAll(idList, displayList, "Choisissez les fiches � coller pour la " & loadInfo.Category.FilterLevel & " s�lectionn�e :")
        Dim errorOccurred As Boolean
        errorOccurred = (Err.Number <> 0)
        On Error GoTo 0
        
    ' Si l'utilisateur a annul� ou une erreur s'est produite
    If errorOccurred Or finalSelection Is Nothing Then
        ' Nettoyer la requ�te avant de sortir
        CleanupPowerQuery loadInfo.Category.PowerQueryName
        ProcessDataLoad = DataLoadResult.Cancelled
        Exit Function
    End If
        
        ' Si aucune fiche n'a �t� s�lectionn�e
        If finalSelection.Count = 0 Then
            MsgBox "Aucune fiche s�lectionn�e. Op�ration annul�e.", vbExclamation
            ' Nettoyer la requ�te avant de sortir
            CleanupPowerQuery loadInfo.Category.PowerQueryName
            ProcessDataLoad = DataLoadResult.Cancelled
            Exit Function
        End If
        
        Set loadInfo.SelectedValues = finalSelection
    End If
    
    ' 4. G�rer le mode d'affichage
    Dim displayModeResult As Variant
    displayModeResult = GetDisplayMode(loadInfo)
    If displayModeResult = -999 Then ' Code d'erreur sp�cifique
        ProcessDataLoad = DataLoadResult.Cancelled
        Exit Function
    End If
    loadInfo.ModeTransposed = displayModeResult
    
    ' 5. G�rer la destination
    Set loadInfo.FinalDestination = GetDestination(loadInfo)
    If loadInfo.FinalDestination Is Nothing Then
        ProcessDataLoad = DataLoadResult.Cancelled
        Exit Function
    End If
    
    ' 6. Coller les donn�es
    If Not PasteData(loadInfo) Then
        ProcessDataLoad = DataLoadResult.Error
        Exit Function
    End If
    
    ' 7. S'assurer que la destination est visible
    With loadInfo.FinalDestination
        .Parent.Activate  ' Activer la feuille de destination
        .Select          ' S�lectionner la cellule de d�part
        .Parent.Range(.Address).Select  ' S�lectionner le range complet
        ActiveWindow.ScrollRow = .Row   ' S'assurer que le haut du tableau est visible
        ActiveWindow.ScrollColumn = .Column  ' S'assurer que la gauche du tableau est visible
    End With
    
    ' 8. Nettoyer la requ�te PowerQuery apr�s le collage r�ussi
    CleanupPowerQuery loadInfo.Category.PowerQueryName
    
    ProcessDataLoad = DataLoadResult.Success
End Function

' R�cup�re les valeurs s�lectionn�es selon le niveau de filtrage
Private Function GetSelectedValues(Category As CategoryInfo) As Collection
    Dim lo As ListObject
    Dim dict As Object
    Dim arrValues() As String
    Dim i As Long, j As Long
    Dim cell As Range
    Dim v As Variant
    
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(Category.PowerQueryName))
    ' S'assurer que la table existe, sinon la charger
    If lo Is Nothing Then
        LoadQueries.LoadQuery Category.PowerQueryName, wsPQData, wsPQData.Cells(1, wsPQData.Cells(1, wsPQData.Columns.Count).End(xlToLeft).Column + 1)
        Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(Category.PowerQueryName))
        If lo Is Nothing Then
            MsgBox "Impossible de charger la table PowerQuery '" & Category.PowerQueryName & "'", vbExclamation
            Set GetSelectedValues = Nothing
            Exit Function
        End If
    End If

    ' Si pas de filtrage, permettre � l'utilisateur de choisir directement dans la liste compl�te
    If Category.FilterLevel = "Pas de filtrage" Then
        On Error Resume Next ' Pour g�rer l'annulation de l'InputBox
        
        ' Cr�er un tableau avec toutes les fiches disponibles
        Dim displayArray() As String
        ReDim displayArray(1 To lo.DataBodyRange.Rows.Count)
        For i = 1 To lo.DataBodyRange.Rows.Count
            ' Utiliser la colonne 2 (nom) comme affichage
            displayArray(i) = lo.DataBodyRange.Rows(i).Columns(2).Value
        Next i
        
        ' Pr�senter les valeurs � l'utilisateur
        Set GetSelectedValues = LoadQueries.ChooseMultipleValuesFromArrayWithAll(displayArray, _
            "Choisissez une ou plusieurs fiches � charger (ex: 1,3,5 ou *) :")
            
        If Err.Number <> 0 Then
            Set GetSelectedValues = Nothing
            Exit Function
        End If
        On Error GoTo 0
        
        ' G�rer la s�lection initiale
        Dim selectedIndices As Collection
        Set selectedIndices = GetSelectedValues
        
        ' Si l'utilisateur a annul� ou n'a rien s�lectionn�
        If selectedIndices Is Nothing Then
            Set GetSelectedValues = Nothing
            Exit Function
        End If
        
        ' Convertir les valeurs en IDs
        Set GetSelectedValues = New Collection
        For Each v In selectedIndices
            ' v est la valeur affich�e, on doit retrouver la ligne correspondante
            For i = 1 To lo.DataBodyRange.Rows.Count
                If lo.DataBodyRange.Rows(i).Columns(2).Value = v Then
                    GetSelectedValues.Add lo.DataBodyRange.Rows(i).Columns(1).Value
                    Exit For
                End If
            Next i
        Next v
        
        ' V�rifier si des IDs ont �t� ajout�s
        If GetSelectedValues.Count = 0 Then
            MsgBox "Aucune fiche s�lectionn�e. Op�ration annul�e.", vbExclamation
            Set GetSelectedValues = Nothing
            Exit Function
        End If
    Else
        ' Cr�er un dictionnaire pour stocker les valeurs uniques
        Set dict = CreateObject("Scripting.Dictionary")

        ' Extraire les valeurs uniques
        For Each cell In lo.ListColumns(Category.FilterLevel).DataBodyRange
            If Not dict.Exists(cell.Value) Then
                dict.Add cell.Value, 1
            End If
        Next cell

        ' Convertir le dictionnaire en tableau et trier
        ReDim arrValues(1 To dict.Count)
        i = 1
        For Each v In dict.Keys
            arrValues(i) = v
            i = i + 1
        Next v

        ' Trier le tableau
        For i = 1 To UBound(arrValues) - 1
            For j = i + 1 To UBound(arrValues)
                If arrValues(i) > arrValues(j) Then
                    Dim temp As String
                    temp = arrValues(i)
                    arrValues(i) = arrValues(j)
                    arrValues(j) = temp
                End If
            Next j
        Next i

        Dim selectedPrimary As Collection
        On Error Resume Next
        Set selectedPrimary = LoadQueries.ChooseMultipleValuesFromArrayWithAll(arrValues, _
            "Choisissez une ou plusieurs " & Category.FilterLevel & " (ex: 1,3,5 ou *) :")
        Dim errorOccurred As Boolean
        errorOccurred = (Err.Number <> 0)
        On Error GoTo 0

        If errorOccurred Or selectedPrimary Is Nothing Then
            MsgBox "Op�ration annul�e", vbInformation
            Set GetSelectedValues = Nothing
            Exit Function
        End If

        If selectedPrimary.Count = 0 Then
            MsgBox "Aucune valeur s�lectionn�e. Op�ration annul�e.", vbExclamation
            Set GetSelectedValues = Nothing
            Exit Function
        End If

        If Category.SecondaryFilterLevel <> "" Then
            ' Deuxi�me �tape de filtrage
            Set dict = CreateObject("Scripting.Dictionary")
            For i = 1 To lo.DataBodyRange.Rows.Count
                For Each v In selectedPrimary
                    If lo.DataBodyRange.Rows(i).Columns(lo.ListColumns(Category.FilterLevel).Index).Value = v Then
                        Dim secVal As String
                        secVal = lo.DataBodyRange.Rows(i).Columns(lo.ListColumns(Category.SecondaryFilterLevel).Index).Value
                        If Not dict.Exists(secVal) Then dict.Add secVal, 1
                        Exit For
                    End If
                Next v
            Next i

            ReDim arrValues(1 To dict.Count)
            i = 1
            For Each v In dict.Keys
                arrValues(i) = v
                i = i + 1
            Next v

            For i = 1 To UBound(arrValues) - 1
                For j = i + 1 To UBound(arrValues)
                    If arrValues(i) > arrValues(j) Then
                        temp = arrValues(i)
                        arrValues(i) = arrValues(j)
                        arrValues(j) = temp
                    End If
                Next j
            Next i

            Dim selectedSecondary As Collection
            On Error Resume Next
            Set selectedSecondary = LoadQueries.ChooseMultipleValuesFromArrayWithAll(arrValues, _
                "Choisissez une ou plusieurs " & Category.SecondaryFilterLevel & " (ex: 1,3,5 ou *) :")
            errorOccurred = (Err.Number <> 0)
            On Error GoTo 0

            If errorOccurred Or selectedSecondary Is Nothing Then
                MsgBox "Op�ration annul�e", vbInformation
                Set GetSelectedValues = Nothing
                Exit Function
            End If

            If selectedSecondary.Count = 0 Then
                MsgBox "Aucune valeur s�lectionn�e. Op�ration annul�e.", vbExclamation
                Set GetSelectedValues = Nothing
                Exit Function
            End If

            Set GetSelectedValues = New Collection
            For i = 1 To lo.DataBodyRange.Rows.Count
                Dim matchPrimary As Boolean
                matchPrimary = False
                For Each v In selectedPrimary
                    If lo.DataBodyRange.Rows(i).Columns(lo.ListColumns(Category.FilterLevel).Index).Value = v Then
                        matchPrimary = True
                        Exit For
                    End If
                Next v
                If matchPrimary Then
                    For Each v In selectedSecondary
                        If lo.DataBodyRange.Rows(i).Columns(lo.ListColumns(Category.SecondaryFilterLevel).Index).Value = v Then
                            GetSelectedValues.Add lo.DataBodyRange.Rows(i).Columns(1).Value
                            Exit For
                        End If
                    Next v
                End If
            Next i

            If GetSelectedValues.Count = 0 Then
                MsgBox "Aucune fiche s�lectionn�e. Op�ration annul�e.", vbExclamation
                Set GetSelectedValues = Nothing
                Exit Function
            End If
        Else
            Set GetSelectedValues = selectedPrimary
        End If
    End If
    Exit Function
    
ErrorHandler:
    If Err.Number = 424 Then  ' "L'objet est requis" - typiquement quand l'utilisateur annule une InputBox
        Set GetSelectedValues = Nothing
    Else
        MsgBox "Une erreur s'est produite : " & Err.Description, vbExclamation
        Set GetSelectedValues = Nothing
    End If
    Exit Function
End Function

' G�re le mode d'affichage (normal/transpos�)
Private Function GetDisplayMode(loadInfo As DataLoadInfo) As Variant
    Dim lo As ListObject
    Dim nbFiches As Long, nbChamps As Long
    Dim previewNormal As String, previewTransposed As String
    Dim userChoice As Double
    Dim i As Long, j As Long, idx As Long
    Dim colWidths() As Integer, rowWidths() As Integer
    Dim v As Variant
    
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    nbFiches = loadInfo.SelectedValues.Count
    nbChamps = lo.ListColumns.Count
    
    ' Pr�parer les exemples pour l'inputbox de mode
    previewNormal = "Mode NORMAL (tableau classique) :" & vbCrLf
    previewTransposed = "Mode TRANSPOSE (fiches en colonnes) :" & vbCrLf
      ' G�n�rer les pr�visualisations
    GeneratePreviews lo, loadInfo, previewNormal, previewTransposed
    
    ' Afficher d'abord les pr�visualisations dans une MsgBox
    MsgBox "Pr�visualisations des modes disponibles :" & vbCrLf & vbCrLf & _
           previewNormal & vbCrLf & previewTransposed, _
           vbInformation, "Aper�u des modes"
           
    ' Puis demander le choix avec une InputBox simple
    Dim modePrompt As String
    modePrompt = "Comment souhaitez-vous coller les fiches ?" & vbCrLf & vbCrLf & _
                 "1 pour NORMAL" & vbCrLf & _
                 "2 pour TRANSPOSE"
    userChoice = Application.InputBox(modePrompt, "Choix du mode de collage", "1", Type:=2)
      ' Si l'utilisateur a cliqu� sur Annuler (Type:=2 retourne False pour Annuler)
    If userChoice = 0 Then
        MsgBox "Op�ration annul�e", vbInformation
        GetDisplayMode = -999 ' Code d'erreur sp�cifique
        Exit Function
    End If
    
    ' V�rifier la validit� de la r�ponse
    If userChoice = 2 Then
        GetDisplayMode = True
    ElseIf userChoice = 1 Then
        GetDisplayMode = False
    Else
        MsgBox "Veuillez entrer 1 ou 2", vbExclamation
        GetDisplayMode = -999 ' Code d'erreur sp�cifique
    End If
End Function

' G�n�re les pr�visualisations pour les deux modes
Private Sub GeneratePreviews(lo As ListObject, loadInfo As DataLoadInfo, _
                           ByRef previewNormal As String, ByRef previewTransposed As String)
    Dim i As Long, j As Long, idx As Long
    Dim colWidths() As Integer, rowWidths() As Integer
    Dim v As Variant
    Dim nbChamps As Long: nbChamps = lo.ListColumns.Count
    
    ' --- Aligned NORMAL preview generation ---
    ReDim colWidths(1 To WorksheetFunction.Min(4, nbChamps))
    For i = 1 To WorksheetFunction.Min(4, nbChamps)
        colWidths(i) = Len(TruncateWithEllipsis(lo.HeaderRowRange.Cells(1, i).Value, 10))
    Next i
    
    idx = 1
    For Each v In loadInfo.SelectedValues
        If idx > loadInfo.PreviewRows Then Exit For
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
    
    ' G�n�rer la pr�visualisation normale
    previewNormal = previewNormal & "| "
    For i = 1 To WorksheetFunction.Min(4, nbChamps)
        Dim head As String
        head = TruncateWithEllipsis(lo.HeaderRowRange.Cells(1, i).Value, 10)
        previewNormal = previewNormal & head & Space(colWidths(i) - Len(head)) & " | "
    Next i
    previewNormal = previewNormal & vbCrLf
    
    idx = 1
    For Each v In loadInfo.SelectedValues
        If idx > loadInfo.PreviewRows Then Exit For
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
    ReDim rowWidths(1 To WorksheetFunction.Min(4, nbChamps))
    For i = 1 To WorksheetFunction.Min(4, nbChamps)
        rowWidths(i) = Len(TruncateWithEllipsis(lo.HeaderRowRange.Cells(1, i).Value, 10))
        idx = 1
        For Each v In loadInfo.SelectedValues
            If idx > loadInfo.PreviewRows Then Exit For
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
        For Each v In loadInfo.SelectedValues
            If idx > loadInfo.PreviewRows Then Exit For
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
End Sub

' G�re la s�lection de la destination
Private Function GetDestination(loadInfo As DataLoadInfo) As Range
    Dim lo As ListObject
    Dim nbRows As Long, nbCols As Long
    Dim okPlage As Boolean
    Dim i As Long, j As Long
    
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    
    ' Calculer la taille n�cessaire
    If loadInfo.ModeTransposed Then
        nbRows = lo.ListColumns.Count
        nbCols = loadInfo.SelectedValues.Count + 1 ' +1 pour les en-t�tes
    Else
        nbRows = loadInfo.SelectedValues.Count + 1 ' +1 pour les en-t�tes
        nbCols = lo.ListColumns.Count
    End If
      ' Informer l'utilisateur
    MsgBox "La plage n�cessaire sera de " & nbRows & " lignes x " & nbCols & " colonnes.", vbInformation      ' Demander la cellule de destination et v�rifier la place
    Do
        Dim selectedRange As Range
        
        ' Activer Excel pour la s�lection
        Application.Interactive = True
        Application.ScreenUpdating = True
        
        ' Demander � l'utilisateur de s�lectionner une cellule
        On Error GoTo ErrorHandler
        Set selectedRange = Application.InputBox( _
            prompt:="S�lectionnez la cellule o� charger les fiches (" & nbRows & " x " & nbCols & ")", _
            title:="Destination", _
            Type:=8)
            
        ' V�rifier si une plage valide a �t� s�lectionn�e
        If selectedRange Is Nothing Then
            MsgBox "Aucune cellule s�lectionn�e. Op�ration annul�e.", vbInformation
            Set GetDestination = Nothing
            Exit Function
        End If
        
        ' S'assurer que c'est une seule cellule
        If selectedRange.Cells.Count > 1 Then
            MsgBox "Veuillez s�lectionner une seule cellule.", vbExclamation
            GoTo ContinueLoop
        End If
        
        Set GetDestination = selectedRange
        GoTo CheckSpace
        
ErrorHandler:
        If Err.Number = 424 Then  ' Erreur "L'objet est requis"
            MsgBox "Op�ration annul�e", vbInformation
            Set GetDestination = Nothing
            Exit Function
        ElseIf Err.Number <> 0 Then
            MsgBox "Une erreur s'est produite : " & Err.Description, vbExclamation
            Set GetDestination = Nothing
            Exit Function
        End If
        Resume Next
        
ContinueLoop:
        Err.Clear
        Resume Next
        
CheckSpace:
        
        okPlage = True
        For i = 0 To nbRows - 1
            For j = 0 To nbCols - 1
                If Not IsEmpty(GetDestination.Offset(i, j)) Then
                    okPlage = False
                    Exit For
                End If
            Next j
            If Not okPlage Then Exit For
        Next i
        
        If Not okPlage Then
            MsgBox "La plage s�lectionn�e n'est pas vide. Veuillez choisir un autre emplacement.", vbExclamation
        End If
    Loop Until okPlage
End Function

' Colle les donn�es selon le mode choisi
Private Function PasteData(loadInfo As DataLoadInfo) As Boolean
    Dim lo As ListObject
    Dim tblRange As Range
    Dim i As Long, j As Long
    Dim v As Variant
    Dim currentCol As Long, currentRow As Long
    
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    
    ' D�prot�ger la feuille de destination avant tout collage
    Dim ws As Worksheet
    Set ws = loadInfo.FinalDestination.Worksheet
    ws.Unprotect
      Log "paste_data", "=== D�BUT PASTEDATA ===" & vbCrLf & _
                "Mode Transpos�: " & loadInfo.ModeTransposed & vbCrLf & _
                "Cat�gorie: " & loadInfo.Category.DisplayName & vbCrLf & _
                "Nombre de colonnes: " & lo.ListColumns.Count & vbCrLf & _
                "Nombre de valeurs s�lectionn�es: " & loadInfo.SelectedValues.Count, _
                DEBUG_LEVEL, "PasteData", "DataLoaderManager"

    ' D�terminer les colonnes visibles en fonction du dictionnaire Ragic
    Dim visibleCols As Collection
    Set visibleCols = New Collection
    Dim header As String
    For i = 1 To lo.ListColumns.Count
        header = lo.HeaderRowRange.Cells(1, i).Value
        If Not IsFieldHidden(loadInfo.Category.SheetName, header) Then
            visibleCols.Add i
        End If
    Next i
      If loadInfo.ModeTransposed Then
        Log "paste_data", "--- D�but collage transpos� ---", DEBUG_LEVEL, "PasteData", "DataLoaderManager"
        ' Coller en transpos�
        For i = 1 To visibleCols.Count
            Log "paste_data", "Colonne " & visibleCols(i) & ": " & lo.HeaderRowRange.Cells(1, visibleCols(i)).Value, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
            loadInfo.FinalDestination.Offset(i - 1, 0).Value = lo.HeaderRowRange.Cells(1, visibleCols(i)).Value
            loadInfo.FinalDestination.Offset(i - 1, 0).NumberFormat = lo.DataBodyRange.Columns(visibleCols(i)).Cells(1, 1).NumberFormat
        Next i
        
        currentCol = 1
        For Each v In loadInfo.SelectedValues
            Log "paste_data", "Traitement colonne " & currentCol & ", valeur=" & v, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
            For j = 1 To lo.DataBodyRange.Rows.Count
                If lo.DataBodyRange.Rows(j).Columns(1).Value = v Then
                    Log "paste_data", "  Trouv� � la ligne " & j, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
                    For i = 1 To visibleCols.Count
                        loadInfo.FinalDestination.Offset(i - 1, currentCol).Value = lo.DataBodyRange.Rows(j).Cells(1, visibleCols(i)).Value
                        loadInfo.FinalDestination.Offset(i - 1, currentCol).NumberFormat = lo.DataBodyRange.Rows(j).Cells(1, visibleCols(i)).NumberFormat
                    Next i
                    Exit For
                End If
            Next j
            currentCol = currentCol + 1
        Next v
        Set tblRange = loadInfo.FinalDestination.Resize(visibleCols.Count, loadInfo.SelectedValues.Count + 1)
        Log "paste_data", "Plage transpos�e d�finie: " & tblRange.Address & " (" & tblRange.Rows.Count & " lignes x " & tblRange.Columns.Count & " colonnes)", DEBUG_LEVEL, "PasteData", "DataLoaderManager"
    Else
        Log "paste_data", "--- D�but collage normal ---", DEBUG_LEVEL, "PasteData", "DataLoaderManager"
        ' Coller en normal
        For i = 1 To visibleCols.Count
            Debug.Print "Colonne " & visibleCols(i) & ": " & lo.HeaderRowRange.Cells(1, visibleCols(i)).Value
            loadInfo.FinalDestination.Offset(0, i - 1).Value = lo.HeaderRowRange.Cells(1, visibleCols(i)).Value
            loadInfo.FinalDestination.Offset(0, i - 1).NumberFormat = lo.DataBodyRange.Columns(visibleCols(i)).Cells(1, 1).NumberFormat
        Next i
        
        currentRow = 1
        For Each v In loadInfo.SelectedValues
            Debug.Print "Traitement ligne " & currentRow & ", valeur=" & v
            For j = 1 To lo.DataBodyRange.Rows.Count
                If lo.DataBodyRange.Rows(j).Columns(1).Value = v Then
                    Debug.Print "  Trouv� � la ligne " & j
                    For i = 1 To visibleCols.Count
                        loadInfo.FinalDestination.Offset(currentRow, i - 1).Value = lo.DataBodyRange.Rows(j).Cells(1, visibleCols(i)).Value
                        loadInfo.FinalDestination.Offset(currentRow, i - 1).NumberFormat = lo.DataBodyRange.Rows(j).Cells(1, visibleCols(i)).NumberFormat
                    Next i
                    Exit For
                End If
            Next j
            currentRow = currentRow + 1
        Next v

        Set tblRange = loadInfo.FinalDestination.Resize(loadInfo.SelectedValues.Count + 1, visibleCols.Count)
        Debug.Print "Plage normale d�finie: " & tblRange.Address & " (" & tblRange.Rows.Count & " lignes x " & tblRange.Columns.Count & " colonnes)"
    End If
      ' V�rification de la validit� de la plage
    Log "paste_data", "=== V�RIFICATIONS ===", DEBUG_LEVEL, "PasteData", "DataLoaderManager"
    Log "paste_data", "Dimensions de la plage: " & tblRange.Rows.Count & " x " & tblRange.Columns.Count, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
    Log "paste_data", "Cellules fusionn�es: " & tblRange.MergeCells, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
    Log "paste_data", "Nombre de tableaux existants: " & tblRange.Worksheet.ListObjects.Count, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
      If tblRange.Rows.Count < 2 Or tblRange.Columns.Count < 2 Then
        Log "paste_data", "ERREUR: Plage trop petite", ERROR_LEVEL, "PasteData", "DataLoaderManager"
        MsgBox "Impossible de cr�er un tableau : la plage s�lectionn�e est trop petite (" & tblRange.Rows.Count & " x " & tblRange.Columns.Count & ").", vbExclamation
        PasteData = False
        Exit Function
    End If
    If tblRange.MergeCells Then
        Log "paste_data", "ERREUR: Cellules fusionn�es d�tect�es", ERROR_LEVEL, "PasteData", "DataLoaderManager"
        MsgBox "Impossible de cr�er un tableau : la plage contient des cellules fusionn�es.", vbExclamation
        PasteData = False
        Exit Function
    End If
    If tblRange.Worksheet.ListObjects.Count > 0 Then
        Dim tbl As ListObject
        For Each tbl In tblRange.Worksheet.ListObjects
            If Not Intersect(tblRange, tbl.Range) Is Nothing Then
                Log "paste_data", "ERREUR: Intersection avec tableau existant - " & tbl.Name, ERROR_LEVEL, "PasteData", "DataLoaderManager"
                MsgBox "Impossible de cr�er un tableau : la plage contient d�j� un tableau Excel.", vbExclamation
                PasteData = False
                Exit Function
            End If
        Next tbl
    End If
    
    Log "paste_data", "=== CR�ATION DU TABLEAU ===", DEBUG_LEVEL, "PasteData", "DataLoaderManager"
    ' Mettre en forme le tableau final    On Error Resume Next
    Set tbl = loadInfo.FinalDestination.Worksheet.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    If Err.Number <> 0 Then
        Log "paste_data", "ERREUR lors de la cr�ation du tableau: " & Err.Description & " (Code: " & Err.Number & ")", ERROR_LEVEL, "PasteData", "DataLoaderManager"
        On Error GoTo 0
        PasteData = False
        Exit Function
    End If
    On Error GoTo 0
    
    tbl.Name = GetUniqueTableName(loadInfo.Category.DisplayName)
    tbl.TableStyle = "TableStyleMedium9"
    Log "paste_data", "Tableau cr�� avec succ�s: " & tbl.Name, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
      ' Prot�ger finement la feuille : seules les valeurs des tableaux EE_ sont prot�g�es
    ProtectSheetWithTable tblRange.Worksheet
    Log "paste_data", "=== FIN PASTEDATA ===", DEBUG_LEVEL, "PasteData", "DataLoaderManager"

    PasteData = True
End Function

' Prot�ge uniquement les tableaux EE_ dans la feuille
Private Sub ProtectSheetWithTable(ws As Worksheet)
    ws.Unprotect
    
    ' 1. D�verrouiller toutes les cellules
    ws.Cells.Locked = False
    
    ' 2. Verrouiller uniquement les cellules des tableaux EE_
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        If Left(tbl.Name, 3) = "EE_" Then
            tbl.Range.Locked = True
        End If
    Next tbl
    
    ' 3. Prot�ger la feuille avec les permissions standard
    ws.Protect UserInterfaceOnly:=True, AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True, _
               AllowInsertingColumns:=True, AllowInsertingRows:=True, _
               AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
               AllowDeletingRows:=True, AllowSorting:=True, _
               AllowFiltering:=True, AllowUsingPivotTables:=True
End Sub

' G�n�re un nom unique pour un nouveau tableau en incr�mentant l'indice
Private Function GetUniqueTableName(CategoryName As String) As String
    Dim baseName As String
    baseName = "EE_" & Utilities.SanitizeTableName(CategoryName)
    Dim maxIndex As Long
    maxIndex = 0
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim currentIndex As Long
    Dim tableName As String

    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = baseName Then
                If maxIndex < 1 Then maxIndex = 1
            ElseIf Left(tbl.Name, Len(baseName) + 1) = baseName & "_" Then
                tableName = Mid(tbl.Name, Len(baseName) + 2)
                If IsNumeric(tableName) Then
                    currentIndex = CLng(tableName)
                    If currentIndex > maxIndex Then
                        maxIndex = currentIndex
                    End If
                End If
            End If
        Next tbl
    Next ws
    GetUniqueTableName = baseName & "_" & (maxIndex + 1)
End Function

' Nettoie la requ�te PowerQuery en supprimant son tableau associ� et la requ�te elle-m�me
Public Sub CleanupPowerQuery(queryName As String)
    On Error Resume Next
    
    ' 1. Supprimer la table si elle existe
    Dim lo As ListObject
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(queryName))
    If Not lo Is Nothing Then
        lo.Delete
    End If
    
    ' 2. Forcer le nettoyage du cache PowerQuery en supprimant la requ�te
    Dim wb As Workbook
    Set wb = ThisWorkbook
    With wb.Queries(queryName)
        .Delete
    End With
    
    On Error GoTo 0
End Sub

' Fonction g�n�rique pour traiter une cat�gorie
Public Function ProcessCategory(CategoryName As String, Optional errorMessage As String = "") As DataLoadResult
    If CategoriesCount = 0 Then InitCategories
    
    Dim loadInfo As DataLoadInfo
    loadInfo.Category = GetCategoryByName(CategoryName)
    If loadInfo.Category.DisplayName = "" Then
        MsgBox "Cat�gorie '" & CategoryName & "' non trouv�e", vbExclamation
        ProcessCategory = DataLoadResult.Error
        Exit Function
    End If
    
    loadInfo.PreviewRows = 3
    
    Dim result As DataLoadResult
    result = ProcessDataLoad(loadInfo)
    
    If result = DataLoadResult.Cancelled Then
        ProcessCategory = DataLoadResult.Cancelled
        Exit Function
    ElseIf result = DataLoadResult.Error Then
        If errorMessage <> "" Then
            MsgBox errorMessage, vbExclamation
        End If
        ProcessCategory = DataLoadResult.Error
        Exit Function
    End If
    
    ProcessCategory = DataLoadResult.Success
End Function






