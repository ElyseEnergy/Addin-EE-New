Attribute VB_Name = "DataPasting"
Option Explicit

' ==========================================
' Module DataPasting
' ------------------------------------------
' Ce module gère la logique de collage des données dans les tableaux Excel.
' Il s'occupe de la mise en forme, du formatage et de la protection des données.
' ==========================================

Private Const MODULE_NAME As String = "DataPasting"

' Colle les données avec la méthode optimisée
Public Function PasteData(loadInfo As DataLoadInfo, Optional TargetTableName As String = "") As Boolean
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "PasteData"
    
    ' --- Déclarations ---
    Dim destSheet As Worksheet
    Dim sourceTable As ListObject
    Dim visibleSourceColIndices As Collection
    Dim headerProcessingInfo As Object ' Dictionary
    Dim destRow As Long, destCol As Long
    Dim colIdx As Long, i As Long, j As Long, k As Long
    Dim v As Variant
    Dim headerName As String
    Dim sourceRowIndex As Long
    Dim cellValue As Variant
    Dim cellInfo As FormattedCellOutput
    Dim finalRange As Range
    Dim numRows As Long, numCols As Long
    Dim lo As ListObject
    Dim m_tableComment As Comment
    
    ' --- SETUP ---
    SYS_Logger.Log PROC_NAME, "Début du collage pour la catégorie: " & loadInfo.Category.DisplayName & ". Mode transposé: " & loadInfo.ModeTransposed, INFO_LEVEL, PROC_NAME, MODULE_NAME
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set destSheet = loadInfo.FinalDestination.Parent
    destSheet.Unprotect
    
    Set sourceTable = DataLoaderManager.GetOrCreatePQDataSheet.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    
    ' --- PRÉ-CALCUL DES COLONNES VISIBLES ---
    Set visibleSourceColIndices = New Collection
    Set headerProcessingInfo = CreateObject("Scripting.Dictionary")
    
    For colIdx = 1 To sourceTable.ListColumns.Count
        headerName = CStr(sourceTable.HeaderRowRange.Cells(1, colIdx).Value)
        
        If Not RagicDictionary.IsFieldHidden(loadInfo.Category.SheetName, headerName) Then
            visibleSourceColIndices.Add colIdx
            headerProcessingInfo(headerName) = RagicDictionary.GetFieldRagicType(loadInfo.Category.SheetName, headerName)
        End If
    Next colIdx
    Log PROC_NAME, "Pré-calcul terminé. " & visibleSourceColIndices.Count & " colonnes visibles à traiter.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    ' --- COLLAGE DES DONNÉES ---
    destRow = loadInfo.FinalDestination.Row
    destCol = loadInfo.FinalDestination.Column
    
    If loadInfo.ModeTransposed Then
        ' --- MODE TRANSPOSÉ ---
        
        ' 1. Coller les en-têtes de ligne
        For i = 1 To visibleSourceColIndices.Count
            colIdx = visibleSourceColIndices(i)
            headerName = sourceTable.HeaderRowRange.Cells(1, colIdx).Value
            Log PROC_NAME, "Collage en-tête transposé: '" & headerName & "' en " & destSheet.Cells(destRow + i - 1, destCol).Address, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
            With destSheet.Cells(destRow + i - 1, destCol)
                .Value = headerName
                .NumberFormat = "@"
                If headerProcessingInfo(headerName) = "Section" Then
                    .Font.Bold = True
                    .Font.Size = .Font.Size + 3
                    .Font.Color = DataFormatter.SECTION_HEADER_DEFAULT_FONT_COLOR
                End If
            End With
        Next i
        
        ' 2. Coller les données
        k = 0
        For Each v In loadInfo.SelectedValues
            k = k + 1
            sourceRowIndex = FindRowIndexInTable(sourceTable, v)
            
            If sourceRowIndex > 0 Then
                Log PROC_NAME, "Traitement ID " & v & " (ligne source " & sourceRowIndex & ")", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                For i = 1 To visibleSourceColIndices.Count
                    colIdx = visibleSourceColIndices(i)
                    headerName = sourceTable.HeaderRowRange.Cells(1, colIdx).Value
                    cellValue = sourceTable.DataBodyRange.Cells(sourceRowIndex, colIdx).Value
                    
                    cellInfo = DataFormatter.GetCellProcessingInfo(cellValue, "", headerName, loadInfo.Category.SheetName)
                    Log PROC_NAME, "  > Champ '" & headerName & "': Valeur='" & cellInfo.FinalValue & "', Format='" & cellInfo.NumberFormatString & "'", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                    
                    With destSheet.Cells(destRow + i - 1, destCol + k)
                        .Value = cellInfo.FinalValue
                        .NumberFormat = cellInfo.NumberFormatString
                    End With
                Next i
            Else
                Log PROC_NAME, "ID non trouvé dans la table source: " & v, WARNING_LEVEL, PROC_NAME, MODULE_NAME
            End If
        Next v
    Else
        ' --- MODE NORMAL ---
        
        ' 1. Coller les en-têtes de colonne
        For i = 1 To visibleSourceColIndices.Count
            colIdx = visibleSourceColIndices(i)
            headerName = sourceTable.HeaderRowRange.Cells(1, colIdx).Value
            Log PROC_NAME, "Collage en-tête normal: '" & headerName & "' en " & destSheet.Cells(destRow, destCol + i - 1).Address, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
            With destSheet.Cells(destRow, destCol + i - 1)
                .Value = headerName
                .NumberFormat = "@"
                If headerProcessingInfo(headerName) = "Section" Then
                    .Font.Bold = True
                    .Font.Size = .Font.Size + 3
                    .Font.Color = DataFormatter.SECTION_HEADER_DEFAULT_FONT_COLOR
                End If
            End With
        Next i
        
        ' 2. Coller les données
        k = 0
        For Each v In loadInfo.SelectedValues
            k = k + 1
            sourceRowIndex = FindRowIndexInTable(sourceTable, v)
            
            If sourceRowIndex > 0 Then
                Log PROC_NAME, "Traitement ID " & v & " (ligne source " & sourceRowIndex & ")", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                For i = 1 To visibleSourceColIndices.Count
                    colIdx = visibleSourceColIndices(i)
                    headerName = sourceTable.HeaderRowRange.Cells(1, colIdx).Value
                    cellValue = sourceTable.DataBodyRange.Cells(sourceRowIndex, colIdx).Value
                    
                    cellInfo = DataFormatter.GetCellProcessingInfo(cellValue, "", headerName, loadInfo.Category.SheetName)
                    Log PROC_NAME, "  > Champ '" & headerName & "': Valeur='" & cellInfo.FinalValue & "', Format='" & cellInfo.NumberFormatString & "'", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                    
                    With destSheet.Cells(destRow + k, destCol + i - 1)
                        .Value = cellInfo.FinalValue
                        .NumberFormat = cellInfo.NumberFormatString
                    End With
                Next i
            Else
                Log PROC_NAME, "ID non trouvé dans la table source: " & v, WARNING_LEVEL, PROC_NAME, MODULE_NAME
            End If
        Next v
    End If
    
    ' --- CRÉATION/MISE À JOUR DU TABLEAU EXCEL ---
    If loadInfo.ModeTransposed Then
        numRows = visibleSourceColIndices.Count
        numCols = loadInfo.SelectedValues.Count + 1
    Else
        numRows = loadInfo.SelectedValues.Count + 1
        numCols = visibleSourceColIndices.Count
    End If
    
    Set finalRange = destSheet.Range(loadInfo.FinalDestination, loadInfo.FinalDestination.Offset(numRows - 1, numCols - 1))
    
    On Error Resume Next
    Set lo = destSheet.ListObjects(TargetTableName)
    On Error GoTo ErrorHandler
    
    If lo Is Nothing Then
        Log PROC_NAME, "Création d'un nouveau tableau Excel dans la plage " & finalRange.Address, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        Set lo = destSheet.ListObjects.Add(xlSrcRange, finalRange, , xlYes)
        If TargetTableName <> "" Then
            lo.Name = TargetTableName
        Else
            lo.Name = TableManager.GetUniqueTableName(loadInfo.Category.CategoryName)
        End If
    Else
        Log PROC_NAME, "Mise à jour du tableau existant '" & lo.Name & "' vers la plage " & finalRange.Address, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        lo.Resize finalRange
    End If

    lo.TableStyle = "TableStyleMedium2"
    lo.Range.Locked = True
    
    ' Sauvegarder les métadonnées
    Set m_tableComment = lo.Range.Cells(1, 1).Comment
    If Not m_tableComment Is Nothing Then m_tableComment.Delete
    Set m_tableComment = lo.Range.Cells(1, 1).AddComment
    m_tableComment.Text TableMetadata.SerializeLoadInfo(loadInfo)
    Log PROC_NAME, "Métadonnées sauvegardées dans le commentaire du tableau '" & lo.Name & "'.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    m_tableComment.Visible = False

    PasteData = True

CleanupAndExit:
    destSheet.Protect UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Log PROC_NAME, "Collage terminé. Statut du succès: " & PasteData, INFO_LEVEL, PROC_NAME, MODULE_NAME
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "pasting_error", "Erreur VBA dans PasteData - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur lors du collage des données"
    PasteData = False
    Resume CleanupAndExit
End Function

' --- Fonctions d'aide privées ---

' Trouve l'index de ligne dans une table source basé sur une valeur ID
Private Function FindRowIndexInTable(ByVal table As ListObject, ByVal idValue As Variant) As Long
    Const PROC_NAME As String = "FindRowIndexInTable"
    On Error GoTo ErrorHandler
    
    FindRowIndexInTable = Application.WorksheetFunction.Match(CLng(idValue), table.ListColumns(1).DataBodyRange, 0)
    
    Exit Function
ErrorHandler:
    ' Si Match ne trouve rien, il déclenche une erreur. C'est le comportement attendu.
    ' On retourne 0 pour indiquer "non trouvé".
    If Err.Number = 1004 Then ' Erreur "Unable to get the Match property of the WorksheetFunction class"
        FindRowIndexInTable = 0
    Else
        ' Pour toute autre erreur, on logue et on la remonte.
        SYS_Logger.Log "pasting_error", "Erreur VBA inattendue dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
        SYS_ErrorHandler.HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la recherche de la ligne pour l'ID " & idValue
        FindRowIndexInTable = 0 ' Retourner 0 en cas d'erreur
    End If
End Function