Attribute VB_Name = "DataPasting"
Option Explicit

' ==========================================
' Module DataPasting
' ------------------------------------------
' Ce module gère la logique de collage des données dans les tableaux Excel.
' Il s'occupe de la mise en forme, du formatage et de la protection des données.
' ==========================================

Private Const MODULE_NAME As String = "DataPasting"
Private m_tableComment As Comment ' Variable partagée pour la gestion des commentaires

' Colle les données avec la méthode optimisée
Public Function PasteData(loadInfo As DataLoadInfo, Optional TargetTableName As String = "") As Boolean
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "PasteData"
    
    ' Désactiver les mises à jour pour la performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Récupérer la feuille de destination
    Dim destSheet As Worksheet
    Set destSheet = loadInfo.FinalDestination.Parent
    
    ' Déprotéger la feuille pour les modifications
    destSheet.Unprotect
    
    ' Récupérer la table source
    Dim sourceTable As ListObject
    Set sourceTable = DataLoaderManager.GetOrCreatePQDataSheet.ListObjects("Table_" & Utilities.SanitizeTableName(loadInfo.Category.PowerQueryName))
    
    ' Préparer les plages source et cible
    Dim sourceRange As Range
    Dim targetRange As Range
    
    ' Déterminer les plages en fonction du mode (normal ou transposé)
    If loadInfo.ModeTransposed Then
        ' Mode transposé : les fiches deviennent des colonnes
        Set sourceRange = GetTransposedSourceRange(sourceTable, loadInfo.SelectedValues)
        Set targetRange = GetTransposedTargetRange(loadInfo.FinalDestination, sourceRange)
    Else
        ' Mode normal : les fiches restent des lignes
        Set sourceRange = GetNormalSourceRange(sourceTable, loadInfo.SelectedValues)
        Set targetRange = GetNormalTargetRange(loadInfo.FinalDestination, sourceRange)
    End If
    
    ' Coller les données
    If loadInfo.ModeTransposed Then
        ' En mode transposé, on doit copier puis transposer
        sourceRange.Copy
        targetRange.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Else
        ' En mode normal, simple copier-coller
        sourceRange.Copy targetRange
    End If
    
    ' Nettoyer le presse-papiers
    Application.CutCopyMode = False
    
    ' Créer ou mettre à jour le ListObject
    Dim lo As ListObject
    Dim existingTable As ListObject
    
    ' Vérifier si le tableau existe déjà
    On Error Resume Next
    If TargetTableName <> "" Then
        Set existingTable = destSheet.ListObjects(TargetTableName)
    End If
    On Error GoTo ErrorHandler

    If Not existingTable Is Nothing Then
        ' --- MISE À JOUR NON-DESTRUCTIVE ---
        Set lo = existingTable
        ' 1. Vider les anciennes données et en-têtes sans supprimer le ListObject
        lo.DataBodyRange.ClearContents
        lo.HeaderRowRange.ClearContents
        ' 2. Redimensionner le tableau pour les nouvelles données
        lo.Resize targetRange
    Else
        ' --- CRÉATION CLASSIQUE ---
        Set lo = destSheet.ListObjects.Add(xlSrcRange, targetRange, , xlYes)
        If TargetTableName <> "" Then
            lo.Name = TargetTableName
        Else
            lo.Name = TableManager.GetUniqueTableName(loadInfo.Category.CategoryName)
        End If
    End If
    
    ' Le reste du collage des données (qui se fait sur la feuille) est déjà fait avant.
    ' Ici on s'assure juste que le ListObject est bien défini.

    ' Appliquer le style de tableau par défaut
    lo.TableStyle = "TableStyleMedium2"
    
    ' Verrouiller le nouveau tableau
    lo.Range.Locked = True
    
    ' Protéger la feuille avec les paramètres standard
    destSheet.Protect UserInterfaceOnly:=True, _
        AllowFormattingCells:=True, _
        AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, _
        AllowInsertingColumns:=True, _
        AllowInsertingRows:=True, _
        AllowInsertingHyperlinks:=True, _
        AllowDeletingColumns:=True, _
        AllowDeletingRows:=True, _
        AllowSorting:=True, _
        AllowFiltering:=True, _
        AllowUsingPivotTables:=True
    
    ' Sauvegarder les métadonnées dans un commentaire
    Set m_tableComment = lo.Range.Cells(1, 1).Comment
    If Not m_tableComment Is Nothing Then m_tableComment.Delete
    Set m_tableComment = lo.Range.Cells(1, 1).AddComment
    m_tableComment.Text TableMetadata.SerializeLoadInfo(loadInfo)
    m_tableComment.Visible = False

    PasteData = True

CleanupAndExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du collage des données"
    PasteData = False
    Resume CleanupAndExit
End Function

' Obtient la plage source pour le mode normal (lignes)
Private Function GetNormalSourceRange(ByVal sourceTable As ListObject, ByVal selectedValues As Collection) As Range
    Const PROC_NAME As String = "GetNormalSourceRange"
    On Error GoTo ErrorHandler
    
    ' Créer une collection pour stocker les lignes à inclure
    Dim selectedRows As Collection
    Set selectedRows = New Collection
    
    ' Ajouter la ligne d'en-tête
    selectedRows.Add 1
    
    ' Trouver les lignes correspondant aux valeurs sélectionnées
    Dim v As Variant
    Dim rowIndex As Long
    For Each v In selectedValues
        For rowIndex = 1 To sourceTable.DataBodyRange.Rows.Count
            If CStr(sourceTable.DataBodyRange.Cells(rowIndex, 1).Value) = CStr(v) Then
                selectedRows.Add rowIndex + 1 ' +1 car la ligne d'en-tête est 1
                Exit For
            End If
        Next rowIndex
    Next v
    
    ' Créer un tableau pour stocker les adresses des lignes
    Dim rowAddresses() As String
    ReDim rowAddresses(1 To selectedRows.Count)
    
    ' Construire les adresses des lignes
    Dim i As Long
    For i = 1 To selectedRows.Count
        rowAddresses(i) = sourceTable.Range.Rows(selectedRows(i)).Address
    Next i
    
    ' Créer la plage union
    Set GetNormalSourceRange = Application.Union(sourceTable.Range.Worksheet.Range(Join(rowAddresses, ",")))
    
    Exit Function
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la création de la plage source en mode normal"
End Function

' Obtient la plage cible pour le mode normal (lignes)
Private Function GetNormalTargetRange(ByVal targetCell As Range, ByVal sourceRange As Range) As Range
    Const PROC_NAME As String = "GetNormalTargetRange"
    On Error GoTo ErrorHandler
    
    ' La plage cible a la même taille que la plage source
    Set GetNormalTargetRange = targetCell.Resize(sourceRange.Rows.Count, sourceRange.Columns.Count)
    
    Exit Function
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la création de la plage cible en mode normal"
End Function

' Obtient la plage source pour le mode transposé (colonnes)
Private Function GetTransposedSourceRange(ByVal sourceTable As ListObject, ByVal selectedValues As Collection) As Range
    Const PROC_NAME As String = "GetTransposedSourceRange"
    On Error GoTo ErrorHandler
    
    ' Créer une collection pour stocker les lignes à inclure
    Dim selectedRows As Collection
    Set selectedRows = New Collection
    
    ' Ajouter la ligne d'en-tête
    selectedRows.Add 1
    
    ' Trouver les lignes correspondant aux valeurs sélectionnées
    Dim v As Variant
    Dim rowIndex As Long
    For Each v In selectedValues
        For rowIndex = 1 To sourceTable.DataBodyRange.Rows.Count
            If CStr(sourceTable.DataBodyRange.Cells(rowIndex, 1).Value) = CStr(v) Then
                selectedRows.Add rowIndex + 1 ' +1 car la ligne d'en-tête est 1
                Exit For
            End If
        Next rowIndex
    Next v
    
    ' Créer un tableau pour stocker les adresses des lignes
    Dim rowAddresses() As String
    ReDim rowAddresses(1 To selectedRows.Count)
    
    ' Construire les adresses des lignes
    Dim i As Long
    For i = 1 To selectedRows.Count
        rowAddresses(i) = sourceTable.Range.Rows(selectedRows(i)).Address
    Next i
    
    ' Créer la plage union
    Set GetTransposedSourceRange = Application.Union(sourceTable.Range.Worksheet.Range(Join(rowAddresses, ",")))
    
    Exit Function
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la création de la plage source en mode transposé"
End Function

' Obtient la plage cible pour le mode transposé (colonnes)
Private Function GetTransposedTargetRange(ByVal targetCell As Range, ByVal sourceRange As Range) As Range
    Const PROC_NAME As String = "GetTransposedTargetRange"
    On Error GoTo ErrorHandler
    
    ' En mode transposé, on inverse les dimensions
    Set GetTransposedTargetRange = targetCell.Resize(sourceRange.Columns.Count, sourceRange.Rows.Count)
    
    Exit Function
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la création de la plage cible en mode transposé"
End Function 