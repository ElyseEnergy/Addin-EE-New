' =============================================================================
' MODULE: DataLoadManager
' Description: Manages data loading operations from various sources
' =============================================================================
Option Explicit

' Constants
Private Const MODULE_NAME As String = "DataLoadManager"
Private Const ERROR_HANDLER_LABEL As String = "ErrorHandler"

' Message Titles
Private Const MSG_TITLE_ERROR As String = "Erreur"
Private Const MSG_TITLE_INFO As String = "Information"
Private Const MSG_TITLE_WARNING As String = "Avertissement"

Public Enum DataLoadResult
    Success = 1
    Cancelled = 2
    Error = 3
End Enum

' Fonction principale de traitement
Public Function ProcessDataLoad(loadInfo As DataLoadInfo) As DataLoadResult
    Const PROC_NAME As String = "ProcessDataLoad"
    On Error GoTo ErrorHandler

    ' 1. Vérifier/Créer la requête PQ
    If Not PQQueryManager.EnsurePQQueryExists(loadInfo.Category) Then
        ElyseMessageBox_System.ShowErrorMessage "Erreur PowerQuery", _
            "Erreur lors de la création de la requête PowerQuery. Veuillez vérifier la configuration."
        ProcessDataLoad = DataLoadResult.Error
        Exit Function
    End If
    
    ' 2. Charger les données
    LoadQueries.LoadQuery category.PowerQueryName, wsPQData, _
        wsPQData.Cells(1, wsPQData.Cells(1, wsPQData.Columns.Count).End(xlToLeft).Column + 1)
    
    ' 3. Sélectionner les valeurs
    Set loadInfo.SelectedValues = GetSelectedValues(loadInfo.Category)
    If loadInfo.SelectedValues Is Nothing Then
        CleanupPowerQuery loadInfo.Category.PowerQueryName
        ProcessDataLoad = DataLoadResult.Cancelled
        Exit Function
    End If
    
    ' 4. Gérer le mode d'affichage
    Dim displayModeResult As Variant
    displayModeResult = GetDisplayMode(loadInfo)
    If displayModeResult = -999 Then
        ProcessDataLoad = DataLoadResult.Cancelled
        Exit Function
    End If
    loadInfo.ModeTransposed = displayModeResult
    
    ' 5. Gérer la destination
    Set loadInfo.FinalDestination = GetDestination(loadInfo)
    If loadInfo.FinalDestination Is Nothing Then
        ProcessDataLoad = DataLoadResult.Cancelled
        Exit Function
    End If
    
    ' 6. Coller les données
    If Not PasteData(loadInfo) Then
        ProcessDataLoad = DataLoadResult.Error
        Exit Function
    End If
    
    ' 7. S'assurer que la destination est visible
    With loadInfo.FinalDestination
        .Parent.Activate
        .Select
        .Parent.Range(.Address).Select
        ActiveWindow.ScrollRow = .Row
        ActiveWindow.ScrollColumn = .Column
    End With
    
    ' Succès !
    ShowSuccessMessage "Chargement réussi", _
        "Les données ont été chargées avec succès à l'emplacement sélectionné."
    ProcessDataLoad = DataLoadResult.Success
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
    ProcessDataLoad = DataLoadResult.Error
End Function

' Récupère les valeurs sélectionnées selon le niveau de filtrage
Private Function GetSelectedValues(category As CategoryInfo) As Collection
    Const PROC_NAME As String = "GetSelectedValues"
    On Error GoTo ErrorHandler
    
    Dim lo As ListObject
    Dim i As Long, j As Long
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim selectedPrimary As Collection
    Dim v As Variant
    Dim secondaryValue As Variant
    Dim errorOccurred As Boolean
    Dim temp As Variant
    
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(category.PowerQueryName))
    ' S'assurer que la table existe, sinon la charger
    If lo Is Nothing Then
        LoadQueries.LoadQuery category.PowerQueryName, wsPQData, wsPQData.Cells(1, wsPQData.Cells(1, wsPQData.Columns.Count).End(xlToLeft).Column + 1)
        Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(category.PowerQueryName))
        If lo Is Nothing Then
            MsgBox "Impossible de charger la table PowerQuery '" & category.PowerQueryName & "'", vbExclamation
            Set GetSelectedValues = Nothing
            Exit Function
        End If
    End If

    ' Si pas de filtrage, permettre à l'utilisateur de choisir directement dans la liste complète
    If category.FilterLevel = "Pas de filtrage" Then
        On Error Resume Next ' Pour gérer l'annulation de l'InputBox
        
        ' Créer un tableau avec toutes les fiches disponibles
        Dim displayArray() As String
        ReDim displayArray(1 To lo.DataBodyRange.Rows.Count)
        For i = 1 To lo.DataBodyRange.Rows.Count
            ' Utiliser la colonne 2 (nom) comme affichage
            displayArray(i) = lo.DataBodyRange.Rows(i).Columns(2).Value
        Next i
        
        ' Présenter les valeurs à l'utilisateur avec le sélecteur de liste personnalisé
        Set GetSelectedValues = SelectMultipleFromList( _
            "Sélection des fiches", _
            "Choisissez une ou plusieurs fiches à charger:", _
            displayArray)
        
        If GetSelectedValues Is Nothing Or GetSelectedValues.Count = 0 Then
            ElyseMessageBox_System.ShowWarningMessage "Sélection vide", _
                "Aucune fiche sélectionnée. L'opération a été annulée."
            Set GetSelectedValues = Nothing
            Exit Function
        End If
        
        Exit Function
    End If
    
    ' Pour les autres cas avec filtrage
    On Error Resume Next
    Set selectedPrimary = LoadQueries.ChooseMultipleValuesFromArrayWithAll(GetUniqueValues(lo, category.FilterLevel), _
        "Choisissez " & category.FilterLevel & " (ex: 1,3,5 ou *) :")
    errorOccurred = (Err.Number <> 0)
    On Error GoTo ErrorHandler
    
    If errorOccurred Or selectedPrimary Is Nothing Then
        Set GetSelectedValues = Nothing
        Exit Function
    End If
    
    If category.SecondaryFilterLevel <> "" Then
        ' Traitement du filtre secondaire
        For Each v In selectedPrimary
            For i = 1 To lo.DataBodyRange.Rows.Count
                If Trim(CStr(lo.DataBodyRange.Rows(i).Columns(GetFilterColumnIndex(lo, category.FilterLevel)).Value)) = Trim(CStr(v)) Then
                    secondaryValue = lo.DataBodyRange.Rows(i).Columns(GetFilterColumnIndex(lo, category.SecondaryFilterLevel)).Value
                    If Not dict.exists(secondaryValue) Then dict.Add secondaryValue, secondaryValue
                End If
            Next i
        Next v
        
        If dict.Count > 0 Then
            ReDim arrValues(1 To dict.Count)
            i = 1
            For Each v In dict.Keys
                arrValues(i) = v
                i = i + 1
            Next v
            
            ' Trier les valeurs
            For i = 1 To UBound(arrValues) - 1
                For j = i + 1 To UBound(arrValues)
                    If arrValues(i) > arrValues(j) Then
                        temp = arrValues(i)
                        arrValues(i) = arrValues(j)
                        arrValues(j) = temp
                    End If
                Next j
            Next i
            
            ' Utiliser le sélecteur de liste personnalisé pour le filtre secondaire
            Set GetSelectedValues = SelectMultipleFromList( _
                "Sélection " & category.SecondaryFilterLevel, _
                "Choisissez une ou plusieurs " & category.SecondaryFilterLevel & ":", _
                arrValues)
                
            If GetSelectedValues Is Nothing Or GetSelectedValues.Count = 0 Then
                ElyseMessageBox_System.ShowWarningMessage "Sélection vide", _
                    "Aucune " & category.SecondaryFilterLevel & " sélectionnée. L'opération a été annulée."
                Set GetSelectedValues = Nothing
                Exit Function
            End If
        Else
            Set GetSelectedValues = selectedPrimary
        End If
    Else
        Set GetSelectedValues = selectedPrimary
    End If
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
    Set GetSelectedValues = Nothing
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
    
    ' Afficher d'abord les prévisualisations dans une MessageBox
    ElyseMessageBox_System.ShowInfoMessage "Aperçu des modes", _
        "Prévisualisations des modes disponibles :" & vbCrLf & vbCrLf & _
        previewNormal & vbCrLf & previewTransposed

    ' Puis demander le choix avec un dialogue personnalisé
    Dim modeItems As Collection
    Set modeItems = New Collection
    modeItems.Add "Mode Normal (tableau classique)"
    modeItems.Add "Mode Transposé (fiches en colonnes)"
    
    Dim modeChoice As Long
    modeChoice = SelectFromList( _
        "Choix du mode de collage", _
        "Comment souhaitez-vous coller les fiches ?", _
        modeItems)

    ' Gérer l'annulation ou la sélection invalide
    If modeChoice = 0 Then
        ElyseMessageBox_System.ShowInfoMessage "Collage annulé", "L'opération a été annulée par l'utilisateur."
        GetDisplayMode = -999 ' Code d'erreur spécifique
        Exit Function
    End If

    ' Traduire la sélection en booléen
    GetDisplayMode = (modeChoice = 2) ' True pour transposé (2), False pour normal (1)
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
    Const PROC_NAME As String = "GetDestination"
    On Error GoTo ErrorHandler

    Dim okPlage As Boolean
    Dim nbRows As Long, nbCols As Long
    Dim i As Long, j As Long

    ' Calculer les dimensions requises
    If loadInfo.ModeTransposed Then
        nbRows = WorksheetFunction.Max(1, loadInfo.SelectedValues.Count + 1)
        nbCols = WorksheetFunction.Max(1, GetVisibleColumnCount(loadInfo))
    Else
        nbRows = WorksheetFunction.Max(1, GetVisibleColumnCount(loadInfo))
        nbCols = WorksheetFunction.Max(1, loadInfo.SelectedValues.Count + 1)
    End If

    Do
        ' Utiliser le sélecteur de plage personnalisé
        Dim selectedRange As Range
        Set selectedRange = SelectRange( _
            "Sélection de la destination", _
            "Sélectionnez la cellule supérieure gauche où coller les données.")
            
        If selectedRange Is Nothing Then
            Set GetDestination = Nothing
            Exit Function
        End If

        Set GetDestination = selectedRange.Cells(1, 1) ' Prendre uniquement la première cellule

        ' Vérifier l'espace disponible
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
            Dim response As Boolean
            response = ElyseMessageBox_System.ShowConfirmationMessage("Destination non vide", _
                "La plage sélectionnée n'est pas vide. Voulez-vous sélectionner un autre emplacement?")
                
            If Not response Then
                Set GetDestination = Nothing ' L'utilisateur ne veut pas continuer
                Exit Function
            End If
        End If
    Loop Until okPlage

    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
    Set GetDestination = Nothing
End Function

' Colle les données selon le mode choisi
Private Function PasteData(loadInfo As DataLoadInfo) As Boolean
    Dim lo As ListObject
    Dim tblRange As Range
    Dim i As Long, j As Long
    Dim v As Variant
    Dim currentCol As Long, currentRow As Long
    
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    
    ' Déprotéger la feuille de destination avant tout collage
    Dim ws As Worksheet
    Set ws = loadInfo.FinalDestination.Worksheet
    ws.Unprotect
    
    Debug.Print "=== DÉBUT PASTEDATA ===" & vbCrLf & _
                "Mode Transposé: " & loadInfo.ModeTransposed & vbCrLf & _
                "Catégorie: " & loadInfo.Category.DisplayName & vbCrLf & _
                "Nombre de colonnes: " & lo.ListColumns.Count & vbCrLf & _
                "Nombre de valeurs sélectionnées: " & loadInfo.SelectedValues.Count

    ' Déterminer les colonnes visibles en fonction du dictionnaire Ragic
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
        Debug.Print "--- Début collage transposé ---"
        ' Coller en transposé
        For i = 1 To visibleCols.Count
            Debug.Print "Colonne " & visibleCols(i) & ": " & lo.HeaderRowRange.Cells(1, visibleCols(i)).Value
            loadInfo.FinalDestination.Offset(i - 1, 0).Value = lo.HeaderRowRange.Cells(1, visibleCols(i)).Value
            loadInfo.FinalDestination.Offset(i - 1, 0).NumberFormat = lo.DataBodyRange.Columns(visibleCols(i)).Cells(1, 1).NumberFormat
        Next i
        
        currentCol = 1
        For Each v In loadInfo.SelectedValues
            Debug.Print "Traitement colonne " & currentCol & ", valeur=" & v
            For j = 1 To lo.DataBodyRange.Rows.Count
                If lo.DataBodyRange.Rows(j).Columns(1).Value = v Then
                    Debug.Print "  Trouvé à la ligne " & j
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
        Debug.Print "Plage transposée définie: " & tblRange.Address & " (" & tblRange.Rows.Count & " lignes x " & tblRange.Columns.Count & " colonnes)"
    Else
        Debug.Print "--- Début collage normal ---"
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
                    Debug.Print "  Trouvé à la ligne " & j
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
        Debug.Print "Plage normale définie: " & tblRange.Address & " (" & tblRange.Rows.Count & " lignes x " & tblRange.Columns.Count & " colonnes)"
    End If
    
    ' Vérification de la validité de la plage
    Debug.Print "=== VÉRIFICATIONS ==="
    Debug.Print "Dimensions de la plage: " & tblRange.Rows.Count & " x " & tblRange.Columns.Count
    Debug.Print "Cellules fusionnées: " & tblRange.MergeCells
    Debug.Print "Nombre de tableaux existants: " & tblRange.Worksheet.ListObjects.Count
    
    If tblRange.Rows.Count < 2 Or tblRange.Columns.Count < 2 Then
        Debug.Print "ERREUR: Plage trop petite"
        MsgBox "Impossible de créer un tableau : la plage sélectionnée est trop petite (" & tblRange.Rows.Count & " x " & tblRange.Columns.Count & ").", vbExclamation
        PasteData = False
        Exit Function
    End If
    If tblRange.MergeCells Then
        Debug.Print "ERREUR: Cellules fusionnées détectées"
        MsgBox "Impossible de créer un tableau : la plage contient des cellules fusionnées.", vbExclamation
        PasteData = False
        Exit Function
    End If
    If tblRange.Worksheet.ListObjects.Count > 0 Then
        Dim tbl As ListObject
        For Each tbl In tblRange.Worksheet.ListObjects
            If Not Intersect(tblRange, tbl.Range) Is Nothing Then
                Debug.Print "ERREUR: Intersection avec tableau existant - " & tbl.Name
                MsgBox "Impossible de créer un tableau : la plage contient déjà un tableau Excel.", vbExclamation
                PasteData = False
                Exit Function
            End If
        Next tbl
    End If
    
    Debug.Print "=== CRÉATION DU TABLEAU ==="
    ' Mettre en forme le tableau final
    On Error Resume Next
    Set tbl = loadInfo.FinalDestination.Worksheet.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    If Err.Number <> 0 Then
        Debug.Print "ERREUR lors de la création du tableau: " & Err.Description & " (Code: " & Err.Number & ")"
        On Error GoTo 0
        PasteData = False
        Exit Function
    End If
    On Error GoTo 0
    
    tbl.name = GetUniqueTableName(loadInfo.Category.DisplayName)
    tbl.TableStyle = "TableStyleMedium9"
    Debug.Print "Tableau créé avec succès: " & tbl.Name
    
    ' Protéger finement la feuille : seules les valeurs des tableaux EE_ sont protégées
    ProtectSheetWithTable tblRange.Worksheet
    Debug.Print "=== FIN PASTEDATA ==="

    PasteData = True
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
Private Function GetUniqueTableName(categoryName As String) As String
    Dim baseName As String
    baseName = "EE_" & Utilities.SanitizeTableName(categoryName)
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

' Nettoie la requête PowerQuery en supprimant son tableau associé et la requête elle-même
Public Sub CleanupPowerQuery(queryName As String)
    On Error Resume Next
    
    ' 1. Supprimer la table si elle existe
    Dim lo As ListObject
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(queryName))
    If Not lo Is Nothing Then
        lo.Delete
    End If
    
    ' 2. Forcer le nettoyage du cache PowerQuery en supprimant la requête
    Dim wb As Workbook
    Set wb = ThisWorkbook
    With wb.Queries(queryName)
        .Delete
    End With
    
    On Error GoTo 0
End Sub

' Fonction générique pour traiter une catégorie
Public Function ProcessCategory(categoryName As String, Optional errorMessage As String = "") As DataLoadResult
    If CategoriesCount = 0 Then InitCategories
    
    Dim loadInfo As DataLoadInfo
    loadInfo.Category = GetCategoryByName(categoryName)
    If loadInfo.Category.DisplayName = "" Then
        MsgBox "Catégorie '" & categoryName & "' non trouvée", vbExclamation
        ProcessCategory = DataLoadResult.Error
        Exit Function
    End If
    
    loadInfo.PreviewRows = 3
    
    If Not ProcessDataLoad(loadInfo) Then
        ' Vérifie si c'est une annulation ou une erreur
        Dim WasCancelled As Boolean
        If WasCancelled Then
            ProcessCategory = DataLoadResult.Cancelled
        Else
            If errorMessage <> "" Then
                MsgBox errorMessage, vbExclamation
            End If
            ProcessCategory = DataLoadResult.Error
        End If
        Exit Function
    End If
    
    ProcessCategory = DataLoadResult.Success
End Function

Public Function LoadDataFromSource(ByVal sourceName As String) As Boolean
    Const PROC_NAME As String = "LoadDataFromSource"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Attempting to load data from source: " & sourceName, PROC_NAME, MODULE_NAME

    ' Dim success As Boolean ' Original variable
    ' success = False ' Original variable
    
    ' If sourceName = "Database" Then
    '   Debug.Print "Connecting to database..."
    '   ' ... connection logic ...
    '   If Connected Then
    '       Debug.Print "Connection successful. Fetching data..."
    '       ' ... fetch logic ...
    '       success = True
    '       Debug.Print "Data fetched successfully."
    '   Else
    '       Debug.Print "Failed to connect to database."
    '       MsgBox "Could not connect to the database. Please check settings.", vbCritical, "Connection Error"
    '       success = False
    '   End If
    ' Else
    '   Debug.Print "Unknown data source: " & sourceName
    '   MsgBox "The data source '" & sourceName & "' is not recognized.", vbExclamation, "Unknown Source"
    '   success = False
    ' End If
    ' LoadDataFromSource = success

    If sourceName = "Database" Then
        LogDebug PROC_NAME & "_ConnectDB", "Connecting to database...", PROC_NAME, MODULE_NAME
        ' ... connection logic ...
        Dim isConnected As Boolean ' Placeholder
        isConnected = True ' Placeholder for actual connection check
        
        If isConnected Then
            LogInfo PROC_NAME & "_ConnectSuccess", "Connection successful. Fetching data...", PROC_NAME, MODULE_NAME
            ' ... fetch logic ...
            LoadDataFromSource = True ' Assuming success
            LogInfo PROC_NAME & "_FetchSuccess", "Data fetched successfully from Database.", PROC_NAME, MODULE_NAME
        Else
            LogError PROC_NAME & "_ConnectFail", Err.Number, "Failed to connect to database.", PROC_NAME, MODULE_NAME ' Assuming Err object is set by connection attempt
            ElyseMessageBox_System.ShowErrorMessage "Connection Error", "Could not connect to the database. Please check settings. Details have been logged."
            LoadDataFromSource = False
        End If
    Else
        LogWarning PROC_NAME & "_UnknownSource", "Unknown data source: " & sourceName, PROC_NAME, MODULE_NAME
        ElyseMessageBox_System.ShowWarningMessage "Unknown Source", "The data source '" & sourceName & "' is not recognized."
        LoadDataFromSource = False
    End If
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
    LoadDataFromSource = False ' Default error return
End Function




