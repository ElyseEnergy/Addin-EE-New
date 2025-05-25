Option Explicit
Private Const MODULE_NAME As String = "Utilities"

' Variables globales
Public wsPQData As Worksheet

Sub InitializePQData()
    On Error Resume Next
    Set wsPQData = ActiveWorkbook.Worksheets("PQ_DATA")
    On Error GoTo 0
    
    ' Si la feuille n'existe pas, la créer
    If wsPQData Is Nothing Then
        Set wsPQData = ActiveWorkbook.Worksheets.Add
        wsPQData.Name = "PQ_DATA"
    End If
End Sub

Function GetLastColumn(ws As Worksheet) As Long
    GetLastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column +1
End Function

' --- Utility function for smart truncation ---
Function TruncateWithEllipsis(text As String, maxLen As Integer) As String
    If Len(text) > maxLen Then
        TruncateWithEllipsis = Left(text, maxLen - 3) & "..."
    Else
        TruncateWithEllipsis = text
    End If
End Function

' Nettoie une chaîne pour en faire un nom de tableau valide
Public Function SanitizeTableName(ByVal inputName As String) As String
    Dim result As String
    Dim i As Long
    Dim c As String
    
    ' Remplacer les caractères non valides par des underscores
    result = inputName
    
    ' Remplacer les espaces par des underscores
    result = Replace(result, " ", "_")
    
    ' Remplacer les caractères spéciaux courants
    result = Replace(result, "-", "_")
    result = Replace(result, ".", "_")
    result = Replace(result, "(", "")
    result = Replace(result, ")", "")
    result = Replace(result, "/", "_")
    result = Replace(result, "\", "_")
    
    ' Supprimer les accents
    result = RemoveDiacritics(result)
    
    ' Ne garder que les caractères alphanumériques et underscores
    Dim cleanResult As String
    cleanResult = ""
    For i = 1 To Len(result)
        c = Mid(result, i, 1)
        If (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or _
           (c >= "0" And c <= "9") Or c = "_" Then
            cleanResult = cleanResult & c
        End If
    Next i
    
    ' Limiter la longueur à 250 caractères (garder de la marge pour les suffixes)
    If Len(cleanResult) > 250 Then
        cleanResult = Left(cleanResult, 250)
    End If
    
    SanitizeTableName = cleanResult
End Function

' Fonction auxiliaire pour supprimer les accents
Private Function RemoveDiacritics(ByVal text As String) As String
    Dim i As Long
    Const AccentedChars = "àáâãäçèéêëìíîïñòóôõöùúûüýÿÀÁÂÃÄÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜÝ"
    Const UnaccentedChars = "aaaaaceeeeiiiinooooouuuuyyAAAAACEEEEIIIINOOOOOUUUUY"
    
    For i = 1 To Len(AccentedChars)
        text = Replace(text, Mid(AccentedChars, i, 1), Mid(UnaccentedChars, i, 1))
    Next i
    
    RemoveDiacritics = text
End Function

Public Function GetCurrentFormattedTimestamp() As String
    Const PROC_NAME As String = "GetCurrentFormattedTimestamp"
    ' No On Error GoTo needed for simple, unlikely-to-fail functions unless specific handling is required.
    ' However, for consistency with the new error handling pattern, it can be added.
    On Error GoTo ErrorHandler
    
    ' ElyseMain_Orchestrator.LogDebug PROC_NAME & "_Called", "Generating formatted timestamp.", PROC_NAME, MODULE_NAME
    ' This function might be called very frequently; logging every call could be noisy.
    ' Consider logging only if in debug mode or if it's a critical utility.
    If ElyseCore_System.IsDebugMode() Then
        ElyseMain_Orchestrator.LogDebug PROC_NAME & "_DebugCall", "GetCurrentFormattedTimestamp called.", PROC_NAME, MODULE_NAME
    End If

    GetCurrentFormattedTimestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Exit Function

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
    GetCurrentFormattedTimestamp = "ErrorInTimestamp" ' Default error return
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    Const PROC_NAME As String = "FileExists"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogDebug PROC_NAME & "_Start", "Checking if file exists: " & filePath, PROC_NAME, MODULE_NAME
    
    Dim fso As Object ' FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        FileExists = True
        ElyseMain_Orchestrator.LogDebug PROC_NAME & "_Found", "File found: " & filePath, PROC_NAME, MODULE_NAME
    Else
        FileExists = False
        ElyseMain_Orchestrator.LogDebug PROC_NAME & "_NotFound", "File not found: " & filePath, PROC_NAME, MODULE_NAME
    End If
    
    Set fso = Nothing
    Exit Function

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
    FileExists = False ' Default to false on error (e.g., invalid path characters)
End Function

Public Sub ShowUtilityMessage(ByVal message As String, Optional ByVal title As String = "Utility Message")
    Const PROC_NAME As String = "ShowUtilityMessage"
    On Error GoTo ErrorHandler
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Display", "Displaying utility message. Title: " & title & ", Message: " & Left(message, 100) & "...", PROC_NAME, MODULE_NAME
    
    ' Original: MsgBox message, vbInformation, title
    ElyseMessageBox_System.ShowInfoMessage title, message
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
    ' Fallback or log that message box couldn't be shown, though ElyseMessageBox_System should handle its own errors.
End Sub