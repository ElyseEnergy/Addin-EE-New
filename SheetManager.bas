Attribute VB_Name = "SheetManager"
Option Explicit

' ==========================================
' Module SheetManager
' ------------------------------------------
' Ce module centralise la logique de manipulation des feuilles Excel.
' Il gère la création, la protection et la gestion des feuilles de l'addin.
' ==========================================

Private Const MODULE_NAME As String = "SheetManager"
Private Const PQ_DATA_SHEET_NAME As String = "PQ_DATA"

' Obtient ou crée la feuille PQ_DATA qui stocke les données brutes des requêtes PowerQuery.
Public Function GetOrCreatePQDataSheet() As Worksheet
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "GetOrCreatePQDataSheet"
    
    ' Essayer d'obtenir la feuille existante
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PQ_DATA_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        ' Créer la feuille si elle n'existe pas
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PQ_DATA_SHEET_NAME
        
        ' Configurer la feuille
        With ws
            .Visible = xlSheetVeryHidden
            .EnableCalculation = True
            .EnableFormatConditionsCalculation = True
            .EnablePivotTable = True
        End With
        
        Log "sheet_manager", "Feuille PQ_DATA créée avec succès", INFO_LEVEL, PROC_NAME, MODULE_NAME
    End If
    
    Set GetOrCreatePQDataSheet = ws
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la création/récupération de la feuille PQ_DATA"
    Set GetOrCreatePQDataSheet = Nothing
End Function

' Protège une feuille avec les paramètres standard de l'addin.
Public Sub ProtectSheetWithStandardSettings(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "ProtectSheetWithStandardSettings"
    
    ws.Protect UserInterfaceOnly:=True, _
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
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la protection de la feuille"
End Sub

' Déprotège une feuille de manière sécurisée.
Public Sub UnprotectSheetSafely(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "UnprotectSheetSafely"
    
    On Error Resume Next
    ws.Unprotect
    On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la déprotection de la feuille"
End Sub

' Liste tous les noms de tableaux dans une feuille donnée.
Public Function ListAllTableNames(ByVal ws As Worksheet) As String
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "ListAllTableNames"
    
    Dim lo As ListObject
    Dim tableNames As String
    tableNames = ""
    
    For Each lo In ws.ListObjects
        If tableNames <> "" Then tableNames = tableNames & ", "
        tableNames = tableNames & lo.Name
    Next lo
    
    ListAllTableNames = tableNames
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la liste des noms de tableaux"
    ListAllTableNames = ""
End Function 