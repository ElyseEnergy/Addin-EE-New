' =============================================================================
' MODULE: DataLoaderManager
' Description: Main module for handling data loading operations
' =============================================================================
Option Explicit
Private Const MODULE_NAME As String = "DataLoaderManager"

' Constants for message dialogs
Private Const MSG_TITLE_ERROR As String = "Erreur"
Private Const MSG_TITLE_INFO As String = "Information"
Private Const MSG_TITLE_WARNING As String = "Avertissement"
Private Const MSG_TITLE_SELECT As String = "Sélection"

Public Sub ProcessCategory(ByVal categoryName As String, Optional errorMessage As String = "")
    On Error GoTo ErrorHandler
    
    Dim category As CategoryInfo
    Set category = CategoryDefinitions_System.GetCategoryByName(categoryName)
    If category Is Nothing Then
        ElyseMessageBox_System.ShowErrorMessage MSG_TITLE_ERROR, _
            "Catégorie '" & categoryName & "' non trouvée"
        Exit Sub
    End If
    
    ' Check if PowerQuery exists and refresh
    If Not PowerQueryManager.ConnectAndRefreshQuery(category.PowerQueryName) Then
        ElyseMessageBox_System.ShowErrorMessage MSG_TITLE_ERROR, _
            "Impossible de charger la table PowerQuery '" & category.PowerQueryName & "'"
        Exit Sub
    End If
    
    ' Get user selection for loading mode
    Dim selectedMode As Long
    selectedMode = GetUserLoadingMode()
    If selectedMode = 0 Then  ' User cancelled
        ElyseMessageBox_System.ShowInfoMessage MSG_TITLE_INFO, "Opération annulée"
        Exit Sub
    End If
    
    ' Get destination range
    Dim destinationRange As Range
    Set destinationRange = GetDestinationRange()
    If destinationRange Is Nothing Then
        ElyseMessageBox_System.ShowErrorMessage MSG_TITLE_ERROR, _
            "Aucune cellule sélectionnée. Opération annulée."
        Exit Sub
    End If
    
    ' ...existing code...
    
    Exit Sub
    
ErrorHandler:
    If Len(errorMessage) > 0 Then
        ElyseMessageBox_System.ShowErrorMessage MSG_TITLE_ERROR, errorMessage
    Else
        ElyseMessageBox_System.ShowErrorMessage MSG_TITLE_ERROR, _
            "Une erreur s'est produite : " & Err.Description
    End If
End Sub

Private Function GetUserLoadingMode() As Long
    Dim modePrompt As String
    modePrompt = "Choisissez le mode de chargement :" & vbCrLf & vbCrLf & _
                 "1. Mode normal" & vbCrLf & _
                 "2. Mode transposé"
    
    Dim userChoice As String
    userChoice = ElyseMessageBox_System.ShowInputDialog(MSG_TITLE_SELECT, modePrompt, "1")
    
    If userChoice = "" Then
        GetUserLoadingMode = 0  ' Cancelled
        Exit Function
    End If
    
    If Not IsNumeric(userChoice) Or userChoice < "1" Or userChoice > "2" Then
        ElyseMessageBox_System.ShowErrorMessage MSG_TITLE_ERROR, _
            "Veuillez entrer 1 ou 2"
        GetUserLoadingMode = 0
        Exit Function
    End If
    
    GetUserLoadingMode = CLng(userChoice)
End Function

Private Function GetDestinationRange() As Range
    Dim prompt As String
    prompt = "Sélectionnez la cellule de destination"
    
    On Error Resume Next
    Set GetDestinationRange = Application.InputBox(prompt, MSG_TITLE_SELECT, Type:=8)
    On Error GoTo 0
    
    If GetDestinationRange Is Nothing Then Exit Function
    
    If GetDestinationRange.Cells.Count > 1 Then
        ElyseMessageBox_System.ShowErrorMessage MSG_TITLE_ERROR, _
            "Veuillez sélectionner une seule cellule."
        Set GetDestinationRange = Nothing
    End If
End Function

Private Sub ValidateDestinationRange(destinationRange As Range, requiredRows As Long, requiredCols As Long)
    ' First show the required size
    ElyseMessageBox_System.ShowInfoMessage MSG_TITLE_INFO, _
        "La plage nécessaire sera de " & requiredRows & " lignes x " & requiredCols & " colonnes."
    
    ' Check if range is empty
    If Not IsRangeEmpty(destinationRange) Then
        ElyseMessageBox_System.ShowErrorMessage MSG_TITLE_ERROR, _
            "La plage sélectionnée n'est pas vide. Veuillez choisir un autre emplacement."
        Exit Sub
    End If
End Function


