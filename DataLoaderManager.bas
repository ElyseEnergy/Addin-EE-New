Attribute VB_Name = "DataLoaderManager"
' Module: DataLoaderManager
' Gère le chargement et l'affichage des données pour toutes les catégories
Option Explicit

Public Enum DataLoadResult
    Success = 1
    Cancelled = 2
    Error = 3
End Enum

' Fonction principale de traitement
Public Function ProcessDataLoad(loadInfo As DataLoadInfo) As DataLoadResult
    Diagnostics.LogTime "Début de ProcessDataLoad pour la catégorie: " & loadInfo.Category.DisplayName

    ' Initialiser la feuille PQ_DATA si besoin
    If wsPQData Is Nothing Then Utilities.InitializePQData
    
    Diagnostics.LogTime "Avant EnsurePQQueryExists"
    If Not PQQueryManager.EnsurePQQueryExists(loadInfo.Category) Then
        MsgBox "Erreur lors de la création de la requête PowerQuery", vbExclamation
        ProcessDataLoad = DataLoadResult.Error
        Exit Function
    End If
    Diagnostics.LogTime "Après EnsurePQQueryExists"
    
    ' --- RETOUR VISUEL PENDANT LE CHARGEMENT ---
    Application.Cursor = xlWait
    Application.StatusBar = "Téléchargement des données pour '" & loadInfo.Category.DisplayName & "' en cours..."
    DoEvents ' Forcer l'affichage du statut
    
    Dim lastCol As Long
    lastCol = Utilities.GetLastColumn(wsPQData)
    Diagnostics.LogTime "Avant LoadQuery (téléchargement des données)"
    LoadQueries.LoadQuery loadInfo.Category.PowerQueryName, wsPQData, wsPQData.Cells(1, lastCol + 1)
    Diagnostics.LogTime "Après LoadQuery (téléchargement des données)"
    
    ' Restaurer le curseur et la barre de statut
    Application.Cursor = xlDefault
    Application.StatusBar = False
    
    ' --- APPEL AU NOUVEAU FORMULAIRE UNIQUE ---
    Diagnostics.LogTime "Avant affichage du formulaire frmDataSelector"
    Dim selectorForm As New frmDataSelector
    selectorForm.ShowForCategory loadInfo.Category
    
    ' Récupérer les résultats du formulaire
    If selectorForm.IsCancelled Then
        Unload selectorForm
        Set selectorForm = Nothing
        CleanupPowerQuery loadInfo.Category.PowerQueryName
        ProcessDataLoad = DataLoadResult.Cancelled
        Exit Function
    End If
    
    Set loadInfo.SelectedValues = selectorForm.SelectedValues
    loadInfo.ModeTransposed = selectorForm.ModeTransposed
    Set loadInfo.FinalDestination = selectorForm.FinalDestination
    
    Unload selectorForm
    Set selectorForm = Nothing
    Diagnostics.LogTime "Après retour du formulaire frmDataSelector"

    ' Coller les données avec la méthode optimisée
    Diagnostics.LogTime "Avant appel à PasteData (Optimisé)"
    If Not PasteData(loadInfo) Then
        ProcessDataLoad = DataLoadResult.Error
        Exit Function
    End If
    Diagnostics.LogTime "Après appel à PasteData (Optimisé)"
    
    ' Activer la feuille de destination
    With loadInfo.FinalDestination
        .Parent.Activate
        .Select
    End With
    
    ' Nettoyer la requête (si nécessaire)
    Diagnostics.LogTime "Avant CleanupPowerQuery"
    CleanupPowerQuery loadInfo.Category.PowerQueryName
    Diagnostics.LogTime "Après CleanupPowerQuery"
    
    ProcessDataLoad = DataLoadResult.Success
End Function

' Colle les données en utilisant une méthode de tableau (array) ultra-rapide
Private Function PasteData(loadInfo As DataLoadInfo) As Boolean
    Diagnostics.LogTime "Début de PasteData (Méthode Array)"
    
    On Error GoTo ErrorHandler
    
    Dim lo As ListObject
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    
    ' --- DÉBUT SECTION CRITIQUE ---
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Diagnostics.LogTime "Fonctionnalités Excel désactivées"
    
    ' 1. Déterminer les colonnes visibles
    Dim visibleCols As Collection
    Set visibleCols = New Collection
    Dim colIndices() As Long
    Dim i As Long, j As Long
    
    For i = 1 To lo.ListColumns.Count
        If Not IsFieldHidden(loadInfo.Category.SheetName, lo.HeaderRowRange.Cells(1, i).Value) Then
            visibleCols.Add lo.HeaderRowRange.Cells(1, i).Value
            ReDim Preserve colIndices(1 To visibleCols.Count)
            colIndices(visibleCols.Count) = i
        End If
    Next i
    Diagnostics.LogTime "Colonnes visibles identifiées"

    ' 2. Construire le tableau de données en mémoire
    Dim dataArray() As Variant
    Dim numRows As Long, numCols As Long
    numCols = visibleCols.Count
    numRows = loadInfo.SelectedValues.Count + 1 ' +1 pour l'en-tête
    
    If loadInfo.ModeTransposed Then
        ReDim dataArray(1 To numCols, 1 To numRows) ' Inverser pour la transposition
    Else
        ReDim dataArray(1 To numRows, 1 To numCols)
    End If

    ' Remplir la ligne d'en-tête
    For j = 1 To numCols
        If loadInfo.ModeTransposed Then dataArray(j, 1) = visibleCols(j) Else dataArray(1, j) = visibleCols(j)
    Next j

    ' Remplir les lignes de données
    Dim dictSourceData As Object
    Set dictSourceData = CreateObject("Scripting.Dictionary")
    For i = 1 To lo.DataBodyRange.Rows.Count
        dictSourceData(lo.DataBodyRange.Cells(i, 1).Value) = lo.DataBodyRange.Rows(i)
    Next i
    
    Dim rowCounter As Long
    rowCounter = 2
    For i = 1 To loadInfo.SelectedValues.Count
        Dim sourceRow As Range
        Set sourceRow = dictSourceData(loadInfo.SelectedValues(i))
        For j = 1 To numCols
            If loadInfo.ModeTransposed Then
                dataArray(j, rowCounter) = sourceRow.Cells(1, colIndices(j)).Value
            Else
                dataArray(rowCounter, j) = sourceRow.Cells(1, colIndices(j)).Value
            End If
        Next j
        rowCounter = rowCounter + 1
    Next i
    Diagnostics.LogTime "Tableau de données construit en mémoire"

    ' 3. Coller le tableau en une seule opération
    Dim destRange As Range
    If loadInfo.ModeTransposed Then
        Set destRange = loadInfo.FinalDestination.Resize(numCols, numRows)
    Else
        Set destRange = loadInfo.FinalDestination.Resize(numRows, numCols)
    End If
    
    destRange.Value2 = dataArray
    Diagnostics.LogTime "Tableau collé dans la feuille"

    ' 4. Mettre en forme comme un tableau Excel (ListObject)
    Dim tbl As ListObject
    loadInfo.FinalDestination.Worksheet.Unprotect
    Set tbl = loadInfo.FinalDestination.Worksheet.ListObjects.Add(xlSrcRange, destRange, , xlYes)
    tbl.Name = GetUniqueTableName(loadInfo.Category.DisplayName)
    tbl.TableStyle = "TableStyleMedium9"
    Diagnostics.LogTime "Objet Tableau (ListObject) créé et formaté"
    
    ' 5. Protéger la feuille
    ProtectSheetWithTable loadInfo.FinalDestination.Worksheet
    Diagnostics.LogTime "Feuille protégée"
    
    PasteData = True

ErrorHandler:
    ' --- FIN SECTION CRITIQUE ---
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Diagnostics.LogTime "Fonctionnalités Excel réactivées"

    If Err.Number <> 0 Then
        MsgBox "An error occurred during the paste operation: " & vbCrLf & Err.Description, vbCritical
        PasteData = False
    End If
    
    ' Attendre la fin des calculs
    Diagnostics.WaitAndLogCalculation
End Function

' Supprime une requête PowerQuery et le tableau associé pour libérer la mémoire
Public Sub CleanupPowerQuery(queryName As String)
    On Error Resume Next
    Dim lo As ListObject
    Dim tName As String
    tName = "Table_" & Utilities.SanitizeTableName(queryName)
    
    ' Suppression plus robuste, ne plante pas si la feuille ou la table n'existe pas
    Set lo = wsPQData.ListObjects(tName)
    If Not lo Is Nothing Then
        lo.Delete
    End If
    
    ThisWorkbook.Queries(queryName).Delete
    
    Log "cleanup_pq", "Nettoyage de " & queryName & " et " & tName, DEBUG_LEVEL, "CleanupPowerQuery", "DataLoaderManager"
    On Error GoTo 0
End Sub

' Récupère les valeurs sélectionnées selon le niveau de filtrage
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

    ' Si pas de filtrage, permettre à l'utilisateur de choisir directement dans la liste complète
    If Category.FilterLevel = "Pas de filtrage" Then
        On Error Resume Next ' Pour gérer l'annulation de l'InputBox
        
        ' Créer un tableau avec toutes les fiches disponibles
        Dim displayArray() As String
        ReDim displayArray(1 To lo.DataBodyRange.Rows.Count)
        For i = 1 To lo.DataBodyRange.Rows.Count
            ' Utiliser la colonne 2 (nom) comme affichage
            displayArray(i) = lo.DataBodyRange.Rows(i).Columns(2).Value
        Next i
        
        ' Présenter les valeurs à l'utilisateur
        Set GetSelectedValues = LoadQueries.ChooseMultipleValuesFromArrayWithAll(displayArray, _
            "Choisissez une ou plusieurs fiches à charger (ex: 1,3,5 ou *) :")
            
        If Err.Number <> 0 Then
            Set GetSelectedValues = Nothing
            Exit Function
        End If
        On Error GoTo 0
        
        ' Gérer la sélection initiale
        Dim selectedIndices As Collection
        Set selectedIndices = GetSelectedValues
        
        ' Si l'utilisateur a annulé ou n'a rien sélectionné
        If selectedIndices Is Nothing Then
            Set GetSelectedValues = Nothing
            Exit Function
        End If
        
        ' Convertir les valeurs en IDs
        Set GetSelectedValues = New Collection
        For Each v In selectedIndices
            ' v est la valeur affichée, on doit retrouver la ligne correspondante
            For i = 1 To lo.DataBodyRange.Rows.Count
                If lo.DataBodyRange.Rows(i).Columns(2).Value = v Then
                    GetSelectedValues.Add lo.DataBodyRange.Rows(i).Columns(1).Value
                    Exit For
                End If
            Next i
        Next v
        
        ' Vérifier si des IDs ont été ajoutés
        If GetSelectedValues.Count = 0 Then
            MsgBox "Aucune fiche sélectionnée. Opération annulée.", vbExclamation
            Set GetSelectedValues = Nothing
            Exit Function
        End If
    Else
        ' Créer un dictionnaire pour stocker les valeurs uniques
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
            MsgBox "Opération annulée", vbInformation
            Set GetSelectedValues = Nothing
            Exit Function
        End If

        If selectedPrimary.Count = 0 Then
            MsgBox "Aucune valeur sélectionnée. Opération annulée.", vbExclamation
            Set GetSelectedValues = Nothing
            Exit Function
        End If

        If Category.SecondaryFilterLevel <> "" Then
            ' Deuxième étape de filtrage
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
                MsgBox "Opération annulée", vbInformation
                Set GetSelectedValues = Nothing
                Exit Function
            End If

            If selectedSecondary.Count = 0 Then
                MsgBox "Aucune valeur sélectionnée. Opération annulée.", vbExclamation
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
                MsgBox "Aucune fiche sélectionnée. Opération annulée.", vbExclamation
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

' Gère le mode d'affichage (normal/transposé)
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
    
    ' Préparer les exemples pour l'inputbox de mode
    previewNormal = "Mode NORMAL (tableau classique) :" & vbCrLf
    previewTransposed = "Mode TRANSPOSE (fiches en colonnes) :" & vbCrLf
      ' Générer les prévisualisations
    GeneratePreviews lo, loadInfo, previewNormal, previewTransposed
    
    ' Afficher d'abord les prévisualisations dans une MsgBox
    MsgBox "Prévisualisations des modes disponibles :" & vbCrLf & vbCrLf & _
           previewNormal & vbCrLf & previewTransposed, _
           vbInformation, "Aperçu des modes"
           
    ' Puis demander le choix avec une InputBox simple
    Dim modePrompt As String
    modePrompt = "Comment souhaitez-vous coller les fiches ?" & vbCrLf & vbCrLf & _
                 "1 pour NORMAL" & vbCrLf & _
                 "2 pour TRANSPOSE"
    userChoice = Application.InputBox(modePrompt, "Choix du mode de collage", "1", Type:=2)
      ' Si l'utilisateur a cliqué sur Annuler (Type:=2 retourne False pour Annuler)
    If userChoice = 0 Then
        MsgBox "Opération annulée", vbInformation
        GetDisplayMode = -999 ' Code d'erreur spécifique
        Exit Function
    End If
    
    ' Vérifier la validité de la réponse
    If userChoice = 2 Then
        GetDisplayMode = True
    ElseIf userChoice = 1 Then
        GetDisplayMode = False
    Else
        MsgBox "Veuillez entrer 1 ou 2", vbExclamation
        GetDisplayMode = -999 ' Code d'erreur spécifique
    End If
End Function

' Génère les prévisualisations pour les deux modes
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
    
    ' Générer la prévisualisation normale
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

' Gère la sélection de la destination
Private Function GetDestination(loadInfo As DataLoadInfo) As Range
    Dim lo As ListObject
    Dim nbRows As Long, nbCols As Long
    Dim okPlage As Boolean
    Dim i As Long, j As Long
    
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    
    ' Calculer la taille nécessaire
    If loadInfo.ModeTransposed Then
        nbRows = lo.ListColumns.Count
        nbCols = loadInfo.SelectedValues.Count + 1 ' +1 pour les en-têtes
    Else
        nbRows = loadInfo.SelectedValues.Count + 1 ' +1 pour les en-têtes
        nbCols = lo.ListColumns.Count
    End If
      ' Informer l'utilisateur
    MsgBox "La plage nécessaire sera de " & nbRows & " lignes x " & nbCols & " colonnes.", vbInformation      ' Demander la cellule de destination et vérifier la place
    Do
        Dim selectedRange As Range
        
        ' Activer Excel pour la sélection
        Application.Interactive = True
        Application.ScreenUpdating = True
        
        ' Demander à l'utilisateur de sélectionner une cellule
        On Error GoTo ErrorHandler
        Set selectedRange = Application.InputBox( _
            prompt:="Sélectionnez la cellule où charger les fiches (" & nbRows & " x " & nbCols & ")", _
            title:="Destination", _
            Type:=8)
            
        ' Vérifier si une plage valide a été sélectionnée
        If selectedRange Is Nothing Then
            MsgBox "Aucune cellule sélectionnée. Opération annulée.", vbInformation
            Set GetDestination = Nothing
            Exit Function
        End If
        
        ' S'assurer que c'est une seule cellule
        If selectedRange.Cells.Count > 1 Then
            MsgBox "Veuillez sélectionner une seule cellule.", vbExclamation
            GoTo ContinueLoop
        End If
        
        Set GetDestination = selectedRange
        GoTo CheckSpace
        
ErrorHandler:
        If Err.Number = 424 Then  ' Erreur "L'objet est requis"
            MsgBox "Opération annulée", vbInformation
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
            MsgBox "La plage sélectionnée n'est pas vide. Veuillez choisir un autre emplacement.", vbExclamation
        End If
    Loop Until okPlage
End Function

' Protège uniquement les tableaux EE_ dans la feuille
Private Sub ProtectSheetWithTable(ws As Worksheet)
    ws.Unprotect
    
    ' 1. Déverrouiller toutes les cellules
    ws.Cells.Locked = False
    
    ' 2. Verrouiller uniquement les cellules des tableaux EE_
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        If Left(tbl.Name, 3) = "EE_" Then
            tbl.Range.Locked = True
        End If
    Next tbl
    
    ' 3. Protéger la feuille avec les permissions standard
    ws.Protect UserInterfaceOnly:=True, AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True, _
               AllowInsertingColumns:=True, AllowInsertingRows:=True, _
               AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
               AllowDeletingRows:=True, AllowSorting:=True, _
               AllowFiltering:=True, AllowUsingPivotTables:=True
End Sub

' Génère un nom unique pour un nouveau tableau en incrémentant l'indice
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

' Fonction générique pour traiter une catégorie
Public Function ProcessCategory(CategoryName As String, Optional errorMessage As String = "") As DataLoadResult
    If CategoriesCount = 0 Then InitCategories
    
    Dim loadInfo As DataLoadInfo
    loadInfo.Category = GetCategoryByName(CategoryName)
    If loadInfo.Category.DisplayName = "" Then
        MsgBox "Catégorie '" & CategoryName & "' non trouvée", vbExclamation
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






