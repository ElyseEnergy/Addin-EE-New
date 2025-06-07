Attribute VB_Name = "DataLoaderManager"

' ==========================================
' Module DataLoaderManager
' ------------------------------------------
' Ce module gère le chargement et l'affichage des données pour toutes les catégories.
' Il orchestre la sélection, le filtrage, le collage et la protection des données importées via PowerQuery.
' Toutes les fonctions sont documentées individuellement ci-dessous.
' ==========================================
Option Explicit

Public Enum DataLoadResult
    Success = 1
    Cancelled = 2
    Error = 3
End Enum

' Fonction principale de traitement des chargements de données pour une catégorie.
' Paramètres :
'   loadInfo (DataLoadInfo) : Informations de chargement
' Retour :
'   DataLoadResult (Succès, Annulé, Erreur)
Public Function ProcessDataLoad(loadInfo As DataLoadInfo) As DataLoadResult
    On Error GoTo ErrorHandler
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

    ' --- SÉLECTION UTILISATEUR PAR INPUTBOX ---
    Diagnostics.LogTime "Avant sélection des valeurs (InputBox)"
    Set loadInfo.SelectedValues = GetSelectedValues(loadInfo.Category)
    If loadInfo.SelectedValues Is Nothing Or loadInfo.SelectedValues.Count = 0 Then
        ProcessDataLoad = DataLoadResult.Cancelled
        Exit Function
    End If

    Diagnostics.LogTime "Avant sélection du mode d'affichage (InputBox)"
    Dim modeResult As Variant
    modeResult = GetDisplayMode(loadInfo)
    If modeResult = -999 Then
        ProcessDataLoad = DataLoadResult.Cancelled
        Exit Function
    End If
    loadInfo.ModeTransposed = (modeResult = True)

    Diagnostics.LogTime "Avant sélection de la destination (InputBox)"
    Set loadInfo.FinalDestination = GetDestination(loadInfo)
    If loadInfo.FinalDestination Is Nothing Then
        ProcessDataLoad = DataLoadResult.Cancelled
        Exit Function
    End If

    ' Coller les données avec la méthode optimisée
    Diagnostics.LogTime "Avant appel à PasteData (Optimisé)"
    If Not PasteData(loadInfo) Then
        ProcessDataLoad = DataLoadResult.Error
        Exit Function
    End If
    Diagnostics.LogTime "Après appel à PasteData (Optimisé)"
    ' S'assurer que la destination est visible
    With loadInfo.FinalDestination
        .Parent.Activate
        .Select
        .Parent.Range(.Address).Select
        ActiveWindow.ScrollRow = .Row
        ActiveWindow.ScrollColumn = .Column
    End With
    ' Ne pas nettoyer la requête PowerQuery pour conserver les requêtes dans le classeur
    ProcessDataLoad = DataLoadResult.Success
    Exit Function
ErrorHandler:
    ProcessDataLoad = DataLoadResult.Error
End Function

' Nettoie la requête PowerQuery en supprimant son tableau associé et la requête elle-même.
' Paramètres :
'   queryName (String) : Nom de la requête à nettoyer
Public Sub CleanupPowerQuery(queryName As String)
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "CleanupPowerQuery"
    Const MODULE_NAME As String = "DataLoaderManager"
    ' 1. Supprimer la table si elle existe
    Dim lo As ListObject
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(queryName))
    If Not lo Is Nothing Then
        lo.Delete
    End If
    ' 2. Forcer le nettoyage du cache PowerQuery en supprimant la requête
    Dim wb As Workbook
    Set wb = ThisWorkbook
    On Error Resume Next
    wb.Queries(queryName).Delete
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    ' Note : On ignore les erreurs ici car c'est une fonction de nettoyage
    ' qui peut échouer si les éléments n'existent déjà plus
    Log "cleanup_error", "Erreur ignorée lors du nettoyage de la requête " & queryName & ": " & Err.Description, WARNING_LEVEL, PROC_NAME, MODULE_NAME
    Resume Next
End Sub

' Récupère les valeurs sélectionnées selon le niveau de filtrage
Private Function GetSelectedValues(Category As CategoryInfo) As Collection
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetSelectedValues"
    Const MODULE_NAME As String = "DataLoaderManager"
    
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
        Dim errorMsg As String
        errorMsg = "Erreur lors de la sélection des valeurs pour la catégorie " & Category.CategoryName & ": " & Err.Description
        HandleError MODULE_NAME, PROC_NAME, errorMsg
        Set GetSelectedValues = Nothing
    End If
    Exit Function
End Function

' Gère le mode d'affichage (normal ou transposé) pour le collage des données.
' Paramètres :
'   loadInfo (DataLoadInfo) : Informations de chargement
' Retour :
'   Variant (True si transposé, False sinon, -999 si annulé)
Private Function GetDisplayMode(loadInfo As DataLoadInfo) As Variant
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetDisplayMode"
    Const MODULE_NAME As String = "DataLoaderManager"
    
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
    Else        MsgBox "Veuillez entrer 1 ou 2", vbExclamation
        GetDisplayMode = -999 ' Code d'erreur spécifique
    End If
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la sélection du mode d'affichage"
    GetDisplayMode = -999 ' Code d'erreur spécifique
End Function

' Génère les prévisualisations pour les deux modes d'affichage (normal et transposé).
' Paramètres :
'   lo (ListObject) : Tableau source
'   loadInfo (DataLoadInfo) : Informations de chargement
'   previewNormal (ByRef String) : Aperçu mode normal
'   previewTransposed (ByRef String) : Aperçu mode transposé
Private Sub GeneratePreviews(lo As ListObject, loadInfo As DataLoadInfo, _
                           ByRef previewNormal As String, ByRef previewTransposed As String)
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GeneratePreviews"
    Const MODULE_NAME As String = "DataLoaderManager"
    
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
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la génération des prévisualisations"
End Sub

' Gère la sélection de la destination de collage des données.
' Paramètres :
'   loadInfo (DataLoadInfo) : Informations de chargement
' Retour :
'   Range (cellule de destination)
Private Function GetDestination(loadInfo As DataLoadInfo) As Range
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetDestination"
    Const MODULE_NAME As String = "DataLoaderManager"
    
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
        If Err.Number = 424 Then  ' Erreur "L'objet est requis" (annulation par l'utilisateur)
            MsgBox "Opération annulée", vbInformation
            Set GetDestination = Nothing
            Exit Function
        ElseIf Err.Number <> 0 Then
            HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la sélection de la destination: " & Err.Description
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

' Colle les données selon le mode choisi (normal ou transposé).
' Paramètres :
'   loadInfo (DataLoadInfo) : Informations de chargement
' Retour :
'   Boolean (True si succès)
Private Function PasteData(loadInfo As DataLoadInfo) As Boolean
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "PasteData"
    Const MODULE_NAME As String = "DataLoaderManager"
    
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
      Log "paste_data", "=== DÉBUT PASTEDATA ===" & vbCrLf & _
                "Mode Transposé: " & loadInfo.ModeTransposed & vbCrLf & _
                "Catégorie: " & loadInfo.Category.DisplayName & vbCrLf & _
                "Nombre de colonnes: " & lo.ListColumns.Count & vbCrLf & _
                "Nombre de valeurs sélectionnées: " & loadInfo.SelectedValues.Count, _
                DEBUG_LEVEL, "PasteData", "DataLoaderManager"

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
        Log "paste_data", "--- Début collage transposé ---", DEBUG_LEVEL, "PasteData", "DataLoaderManager"
        ' Coller en transposé
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
                    Log "paste_data", "  Trouvé à la ligne " & j, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
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
        Log "paste_data", "Plage transposée définie: " & tblRange.Address & " (" & tblRange.Rows.Count & " lignes x " & tblRange.Columns.Count & " colonnes)", DEBUG_LEVEL, "PasteData", "DataLoaderManager"
    Else
        Log "paste_data", "--- Début collage normal ---", DEBUG_LEVEL, "PasteData", "DataLoaderManager"
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
    Log "paste_data", "=== VÉRIFICATIONS ===", DEBUG_LEVEL, "PasteData", "DataLoaderManager"
    Log "paste_data", "Dimensions de la plage: " & tblRange.Rows.Count & " x " & tblRange.Columns.Count, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
    Log "paste_data", "Cellules fusionnées: " & tblRange.MergeCells, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
    Log "paste_data", "Nombre de tableaux existants: " & tblRange.Worksheet.ListObjects.Count, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
      If tblRange.Rows.Count < 2 Or tblRange.Columns.Count < 2 Then
        Log "paste_data", "ERREUR: Plage trop petite", ERROR_LEVEL, "PasteData", "DataLoaderManager"
        MsgBox "Impossible de créer un tableau : la plage sélectionnée est trop petite (" & tblRange.Rows.Count & " x " & tblRange.Columns.Count & ").", vbExclamation
        PasteData = False
        Exit Function
    End If
    If tblRange.MergeCells Then
        Log "paste_data", "ERREUR: Cellules fusionnées détectées", ERROR_LEVEL, "PasteData", "DataLoaderManager"
        MsgBox "Impossible de créer un tableau : la plage contient des cellules fusionnées.", vbExclamation
        PasteData = False
        Exit Function
    End If
    If tblRange.Worksheet.ListObjects.Count > 0 Then
        Dim tbl As ListObject
        For Each tbl In tblRange.Worksheet.ListObjects
            If Not Intersect(tblRange, tbl.Range) Is Nothing Then
                Log "paste_data", "ERREUR: Intersection avec tableau existant - " & tbl.Name, ERROR_LEVEL, "PasteData", "DataLoaderManager"
                MsgBox "Impossible de créer un tableau : la plage contient déjà un tableau Excel.", vbExclamation
                PasteData = False
                Exit Function
            End If
        Next tbl
    End If
    
    Log "paste_data", "=== CRÉATION DU TABLEAU ===", DEBUG_LEVEL, "PasteData", "DataLoaderManager"
    ' Mettre en forme le tableau final    On Error Resume Next
    Set tbl = loadInfo.FinalDestination.Worksheet.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    If Err.Number <> 0 Then
        Log "paste_data", "ERREUR lors de la création du tableau: " & Err.Description & " (Code: " & Err.Number & ")", ERROR_LEVEL, "PasteData", "DataLoaderManager"
        On Error GoTo 0
        PasteData = False
        Exit Function
    End If
    On Error GoTo 0
    
    tbl.Name = GetUniqueTableName(loadInfo.Category.DisplayName)
    tbl.TableStyle = "TableStyleMedium9"
    Log "paste_data", "Tableau créé avec succès: " & tbl.Name, DEBUG_LEVEL, "PasteData", "DataLoaderManager"
      ' Protéger finement la feuille : seules les valeurs des tableaux EE_ sont protégées
    ProtectSheetWithTable tblRange.Worksheet
    Log "paste_data", "=== FIN PASTEDATA ===", DEBUG_LEVEL, "PasteData", "DataLoaderManager"    
    PasteData = True
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du collage des données: " & Err.Description
    
    ' Cleanup en cas d'erreur
    On Error Resume Next
    ws.Unprotect
    PasteData = False
    Exit Function
End Function

' Protège uniquement les tableaux EE_ dans la feuille spécifiée.
' Paramètres :
'   ws (Worksheet) : Feuille à protéger
Private Sub ProtectSheetWithTable(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "ProtectSheetWithTable"
    Const MODULE_NAME As String = "DataLoaderManager"
    
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
    
    ' 3. Protéger la feuille avec les permissions standard    ws.Protect UserInterfaceOnly:=True, AllowFormattingCells:=True, _
               AllowFormattingColumns:=True, AllowFormattingRows:=True, _
               AllowInsertingColumns:=True, AllowInsertingRows:=True, _
               AllowInsertingHyperlinks:=True, AllowDeletingColumns:=True, _
               AllowDeletingRows:=True, AllowSorting:=True, _
               AllowFiltering:=True, AllowUsingPivotTables:=True
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la protection de la feuille avec les tableaux"
End Sub

' Génère un nom unique pour un nouveau tableau en incrémentant l'indice.
' Paramètres :
'   CategoryName (String) : Nom de la catégorie
' Retour :
'   String (nom unique)
Private Function GetUniqueTableName(CategoryName As String) As String
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetUniqueTableName"
    Const MODULE_NAME As String = "DataLoaderManager"
    
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
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la génération du nom unique pour le tableau de la catégorie " & CategoryName
    GetUniqueTableName = baseName & "_ERROR"
End Function

' Fonction générique pour traiter une catégorie par son nom.
' Paramètres :
'   CategoryName (String) : Nom de la catégorie
'   errorMessage (String, optionnel) : Message d'erreur personnalisé
' Retour :
'   DataLoadResult (Succès, Annulé, Erreur)
Public Function ProcessCategory(CategoryName As String, Optional errorMessage As String = "") As DataLoadResult
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "ProcessCategory"
    Const MODULE_NAME As String = "DataLoaderManager"
    
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
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du traitement de la catégorie " & CategoryName & ": " & Err.Description
    If errorMessage <> "" Then
        MsgBox errorMessage, vbExclamation
    End If
    ProcessCategory = DataLoadResult.Error
End Function




