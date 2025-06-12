Attribute VB_Name = "DataInteraction"
Option Explicit

' ==========================================
' Module DataInteraction
' ------------------------------------------
' Ce module gère toutes les interactions avec l'utilisateur pendant le processus de chargement.
' Il centralise les boîtes de dialogue et les sélections utilisateur.
' ==========================================

Private Const MODULE_NAME As String = "DataInteraction"

' Demande à l'utilisateur de sélectionner les valeurs à charger, en appliquant des filtres si nécessaire.
Public Function GetSelectedValues(Category As CategoryInfo) As Collection
    Const PROC_NAME As String = "GetSelectedValues"
    On Error GoTo ErrorHandler
    SYS_Logger.Log "selection_start", "Début de la sélection pour la catégorie: " & Category.DisplayName, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    ' --- SETUP ---
    Dim sourceTable As ListObject
    Set sourceTable = DataLoaderManager.GetOrCreatePQDataSheet.ListObjects("Table_" & Utilities.SanitizeTableName(Category.PowerQueryName))
    
    If sourceTable Is Nothing Or sourceTable.ListRows.Count = 0 Then
        MsgBox "Aucune donnée source à filtrer pour '" & Category.DisplayName & "'.", vbExclamation
        Set GetSelectedValues = Nothing
        Exit Function
    End If
    
    ' --- LOGIQUE DE FILTRAGE ---
    ' Si pas de filtre défini, on passe en sélection directe
    If Category.FilterLevel = "" Or Category.FilterLevel = "Pas de filtrage" Then
        SYS_Logger.Log "selection_direct_mode", "Aucun filtre défini. Passage en mode de sélection directe.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        Set GetSelectedValues = SelectAllItems(sourceTable, Category)
        Exit Function
    End If

    ' --- PREMIER NIVEAU DE FILTRE ---
    SYS_Logger.Log "selection_filter1", "Application du filtre de premier niveau sur: " & Category.FilterLevel, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Dim primaryFilters As Collection
    Set primaryFilters = GetUniqueValuesForFilter(sourceTable, Category.FilterLevel)
    
    Dim selectedPrimary As Collection
    Set selectedPrimary = AskUserForSelection(primaryFilters, "Choisissez une ou plusieurs valeurs pour : " & Category.FilterLevel)
    If selectedPrimary Is Nothing Or selectedPrimary.Count = 0 Then
        SYS_Logger.Log "selection_cancel_1", "Sélection du filtre primaire annulée par l'utilisateur.", INFO_LEVEL, PROC_NAME, MODULE_NAME
        Set GetSelectedValues = Nothing
        Exit Function ' Annulé par l'utilisateur
    End If
    
    ' --- DEUXIÈME NIVEAU DE FILTRE ---
    Dim finalIDs As Collection
    If Category.SecondaryFilterLevel <> "" Then
        SYS_Logger.Log "selection_filter2", "Application du filtre de second niveau sur: " & Category.SecondaryFilterLevel, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        Dim secondaryFilters As Collection
        Set secondaryFilters = GetUniqueValuesForFilter(sourceTable, Category.SecondaryFilterLevel, Category.FilterLevel, selectedPrimary)
        
        Dim selectedSecondary As Collection
        Set selectedSecondary = AskUserForSelection(secondaryFilters, "Choisissez une ou plusieurs valeurs pour : " & Category.SecondaryFilterLevel)
        If selectedSecondary Is Nothing Or selectedSecondary.Count = 0 Then
            SYS_Logger.Log "selection_cancel_2", "Sélection du filtre secondaire annulée par l'utilisateur.", INFO_LEVEL, PROC_NAME, MODULE_NAME
            Set GetSelectedValues = Nothing
            Exit Function ' Annulé par l'utilisateur
        End If
        
        ' Collecter les IDs finaux basés sur les deux filtres
        Set finalIDs = CollectFinalIDs(sourceTable, Category.FilterLevel, selectedPrimary, Category.SecondaryFilterLevel, selectedSecondary)
    Else
        ' Collecter les IDs finaux basés sur le premier filtre uniquement
        SYS_Logger.Log "selection_no_filter2", "Pas de filtre secondaire. Collecte des IDs basée sur le filtre primaire.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        Set finalIDs = CollectFinalIDs(sourceTable, Category.FilterLevel, selectedPrimary)
    End If
    
    If Not finalIDs Is Nothing Then
        SYS_Logger.Log "selection_complete", finalIDs.Count & " ID(s) final(aux) collecté(s).", INFO_LEVEL, PROC_NAME, MODULE_NAME
    Else
        SYS_Logger.Log "selection_no_ids", "Aucun ID final collecté.", WARNING_LEVEL, PROC_NAME, MODULE_NAME
    End If
    Set GetSelectedValues = finalIDs
    
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "interaction_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la sélection des valeurs"
    Set GetSelectedValues = Nothing
End Function

' Fonction d'aide pour la sélection directe quand aucun filtre n'est appliqué.
Private Function SelectAllItems(ByVal sourceTable As ListObject, Category As CategoryInfo) As Collection
    Const PROC_NAME As String = "SelectAllItems"
    On Error GoTo ErrorHandler
    Dim values() As String
    Dim i As Long
    ReDim values(1 To sourceTable.ListRows.Count)

    If sourceTable.ListColumns.Count > 1 Then
        For i = 1 To sourceTable.ListRows.Count
            values(i) = CStr(sourceTable.ListRows(i).Range(1).Value) & "  -  " & CStr(sourceTable.ListRows(i).Range(2).Value)
        Next i
    Else
        For i = 1 To sourceTable.ListRows.Count
            values(i) = CStr(sourceTable.ListRows(i).Range(1).Value)
        Next i
    End If
    
    Dim availableValuesPrompt As String
    availableValuesPrompt = Join(values, vbCrLf)
    
    Dim selectionStr As String
    selectionStr = InputBox("Entrez les ID (seulement les numéros), séparés par des virgules." & vbCrLf & vbCrLf & _
                            "Valeurs disponibles :" & vbCrLf & availableValuesPrompt, _
                            "Sélection pour " & Category.DisplayName)
    
    SYS_Logger.Log "selection_raw_input", "Sélection brute de l'utilisateur: '" & selectionStr & "'", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    If selectionStr = "" Then Set SelectAllItems = Nothing: Exit Function
    
    Dim idCollection As Collection
    Set idCollection = New Collection
    Dim selectedArray() As String
    selectedArray = Split(selectionStr, ",")
    
    Dim v As Variant
    For Each v In selectedArray
        If Trim(v) <> "" Then idCollection.Add Trim(v)
    Next v
    Set SelectAllItems = idCollection
    Exit Function
ErrorHandler:
    SYS_Logger.Log "interaction_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur dans SelectAllItems."
    Set SelectAllItems = Nothing
End Function

' Fonction d'aide pour extraire les valeurs uniques d'une colonne de table, potentiellement filtrées par une sélection précédente.
Private Function GetUniqueValuesForFilter(ByVal sourceTable As ListObject, ByVal filterColumnName As String, _
                                          Optional ByVal parentFilterColumn As String = "", Optional ByVal parentSelection As Collection = Nothing) As Collection
    Const PROC_NAME As String = "GetUniqueValuesForFilter"
    On Error GoTo ErrorHandler
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim filterColIdx As Long
    filterColIdx = sourceTable.ListColumns(filterColumnName).Index
    
    Dim parentColIdx As Long
    If parentFilterColumn <> "" Then parentColIdx = sourceTable.ListColumns(parentFilterColumn).Index
    
    Dim row As ListRow
    For Each row In sourceTable.ListRows
        Dim shouldAdd As Boolean
        shouldAdd = False
        
        If parentSelection Is Nothing Then
            shouldAdd = True
        Else
            Dim parentValue As String
            parentValue = CStr(row.Range.Cells(1, parentColIdx).Value)
            Dim item As Variant
            For Each item In parentSelection
                If parentValue = CStr(item) Then
                    shouldAdd = True
                    Exit For
                End If
            Next item
        End If
        
        If shouldAdd Then
            Dim cellValue As String
            cellValue = CStr(row.Range.Cells(1, filterColIdx).Value)
            If Not dict.Exists(cellValue) Then
                dict.Add cellValue, 1
            End If
        End If
    Next row
    
    Set GetUniqueValuesForFilter = New Collection
    For Each v In dict.Keys
        GetUniqueValuesForFilter.Add v
    Next v
    SYS_Logger.Log "filter_unique_values", "Filtre sur '" & filterColumnName & "'. Trouvé " & dict.Count & " valeur(s) unique(s).", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Exit Function
ErrorHandler:
    SYS_Logger.Log "interaction_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur dans GetUniqueValuesForFilter pour la colonne " & filterColumnName
    Set GetUniqueValuesForFilter = Nothing
End Function

' Fonction d'aide pour demander à l'utilisateur de faire une sélection à partir d'une liste.
Private Function AskUserForSelection(ByVal items As Collection, ByVal prompt As String) As Collection
    Const PROC_NAME As String = "AskUserForSelection"
    On Error GoTo ErrorHandler
    If items Is Nothing Or items.Count = 0 Then
        MsgBox "Aucune option à sélectionner pour : " & prompt, vbInformation
        Set AskUserForSelection = Nothing
        Exit Function
    End If
    
    ' Convertir la collection en tableau pour utiliser Join
    Dim arr() As String
    ReDim arr(1 To items.Count)
    Dim i As Long
    For i = 1 To items.Count
        arr(i) = CStr(items(i))
    Next i
    
    Dim listPrompt As String
    listPrompt = Join(arr, vbCrLf)
    
    Dim selectionStr As String
    selectionStr = InputBox(prompt & vbCrLf & "Séparez vos choix par une virgule." & vbCrLf & vbCrLf & listPrompt, "Sélection requise")
    
    SYS_Logger.Log "user_selection_raw", "Sélection brute de l'utilisateur pour '" & prompt & "': '" & selectionStr & "'", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    If selectionStr = "" Then Set AskUserForSelection = Nothing: Exit Function
    
    Set AskUserForSelection = New Collection
    Dim selectedArray() As String
    selectedArray = Split(selectionStr, ",")
    
    Dim v As Variant
    For Each v In selectedArray
        If Trim(v) <> "" Then AskUserForSelection.Add Trim(v)
    Next v
    Exit Function
ErrorHandler:
    SYS_Logger.Log "interaction_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur dans AskUserForSelection."
    Set AskUserForSelection = Nothing
End Function

' Fonction d'aide pour collecter les IDs finaux après tous les filtres.
Private Function CollectFinalIDs(ByVal sourceTable As ListObject, ByVal primaryColName As String, ByVal primarySelection As Collection, _
                                 Optional ByVal secondaryColName As String = "", Optional ByVal secondarySelection As Collection = Nothing) As Collection
    Const PROC_NAME As String = "CollectFinalIDs"
    On Error GoTo ErrorHandler
    Set CollectFinalIDs = New Collection
    Dim primaryColIdx As Long
    primaryColIdx = sourceTable.ListColumns(primaryColName).Index
    
    Dim secondaryColIdx As Long
    If secondaryColName <> "" Then secondaryColIdx = sourceTable.ListColumns(secondaryColName).Index
    
    Dim row As ListRow
    For Each row In sourceTable.ListRows
        Dim primaryMatch As Boolean
        primaryMatch = False
        Dim item As Variant
        For Each item In primarySelection
            If CStr(row.Range.Cells(1, primaryColIdx).Value) = CStr(item) Then
                primaryMatch = True
                Exit For
            End If
        Next item
        
        If primaryMatch Then
            If secondaryColName = "" Then
                CollectFinalIDs.Add row.Range.Cells(1, 1).Value ' Ajoute l'ID
            Else
                Dim secondaryMatch As Boolean
                secondaryMatch = False
                For Each item In secondarySelection
                    If CStr(row.Range.Cells(1, secondaryColIdx).Value) = CStr(item) Then
                        secondaryMatch = True
                        Exit For
                    End If
                Next item
                If secondaryMatch Then
                    CollectFinalIDs.Add row.Range.Cells(1, 1).Value ' Ajoute l'ID
                End If
            End If
        End If
    Next row
    Exit Function
ErrorHandler:
    SYS_Logger.Log "interaction_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur dans CollectFinalIDs."
    Set CollectFinalIDs = Nothing
End Function

' Demande à l'utilisateur de sélectionner la destination du collage.
Public Function GetDestination(loadInfo As DataLoadInfo) As Range
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "GetDestination"
    
    ' Afficher la boîte de dialogue de sélection
    Dim message As String
    message = "Sélectionnez la cellule de destination pour " & loadInfo.Category.DisplayName & "." & vbCrLf & _
             "Mode : " & IIf(loadInfo.ModeTransposed, "TRANSPOSÉ (colonnes)", "NORMAL (lignes)")
    
    Dim destination As Range
    On Error Resume Next
    Set destination = Application.InputBox(message, "Destination", Type:=8)
    On Error GoTo ErrorHandler
    
    ' Vérifier si l'utilisateur a annulé
    If destination Is Nothing Then Exit Function
    
    ' Vérifier que la destination est une seule cellule
    If destination.Cells.Count > 1 Then
        MsgBox "Veuillez sélectionner une seule cellule.", vbExclamation
        Set GetDestination = Nothing
        Exit Function
    End If
    
    ' Vérifier que la destination n'est pas dans la feuille PQ_DATA
    If destination.Parent.Name = "PQ_DATA" Then
        MsgBox "La feuille PQ_DATA est réservée aux données brutes. Veuillez choisir une autre destination.", vbExclamation
        Set GetDestination = Nothing
        Exit Function
    End If
    
    Set GetDestination = destination
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "interaction_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la sélection de la destination"
    Set GetDestination = Nothing
End Function

' Demande à l'utilisateur de choisir le mode de collage des données.
' Retourne True si transposé, False si normal.
Public Function GetPasteMode() As Boolean
    ' ... existing code ...
End Function