Option Explicit
Private Const MODULE_NAME As String = "Utilities"

' Variables globales
Public wsPQData As Worksheet

Sub InitializePQData()
    Const PROC_NAME As String = "InitializePQData"
    On Error GoTo ErrorHandler
    
    LogInfo PROC_NAME & "_Start", "Initialisation de la feuille PQ_DATA", PROC_NAME, MODULE_NAME
    
    ' Essayer de récupérer la feuille existante
    Set wsPQData = ActiveWorkbook.Worksheets("PQ_DATA")
    
    ' Si la feuille n'existe pas, la créer
    If wsPQData Is Nothing Then
        LogInfo PROC_NAME & "_Create", "Création de la feuille PQ_DATA", PROC_NAME, MODULE_NAME
        Set wsPQData = ActiveWorkbook.Worksheets.Add
        wsPQData.Name = "PQ_DATA"
        
        ' Configuration initiale de la feuille
        With wsPQData
            .Visible = xlSheetVeryHidden
            .ProtectContents = True
        End With
        
        LogInfo PROC_NAME & "_Created", "Feuille PQ_DATA créée et configurée", PROC_NAME, MODULE_NAME
    Else
        LogInfo PROC_NAME & "_Found", "Feuille PQ_DATA trouvée", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Erreur lors de l'initialisation de PQ_DATA: " & Err.Description, PROC_NAME, MODULE_NAME
    Set wsPQData = Nothing
End Sub

Function GetLastColumn(ws As Worksheet) As Long
    Const PROC_NAME As String = "GetLastColumn"
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        LogError PROC_NAME & "_InvalidSheet", "Feuille invalide", PROC_NAME, MODULE_NAME
        GetLastColumn = 0
        Exit Function
    End If
    
    GetLastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    LogDebug PROC_NAME & "_Result", "Dernière colonne trouvée: " & GetLastColumn, PROC_NAME, MODULE_NAME
    Exit Function
    
ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Erreur lors de la recherche de la dernière colonne: " & Err.Description, PROC_NAME, MODULE_NAME
    GetLastColumn = 0
End Function

' --- Utility function for smart truncation ---
Function TruncateWithEllipsis(text As String, maxLen As Integer) As String
    Const PROC_NAME As String = "TruncateWithEllipsis"
    On Error GoTo ErrorHandler
    
    ' Validation des paramètres
    If maxLen < 4 Then
        LogWarning PROC_NAME & "_InvalidLength", "Longueur maximale trop courte: " & maxLen, PROC_NAME, MODULE_NAME
        maxLen = 4 ' Minimum pour avoir au moins un caractère + "..."
    End If
    
    If text = "" Then
        TruncateWithEllipsis = ""
        Exit Function
    End If
    
    If Len(text) > maxLen Then
        TruncateWithEllipsis = Left(text, maxLen - 3) & "..."
        LogDebug PROC_NAME & "_Truncated", "Texte tronqué de " & Len(text) & " à " & maxLen & " caractères", PROC_NAME, MODULE_NAME
    Else
        TruncateWithEllipsis = text
    End If
    
    Exit Function
    
ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Erreur lors de la troncature du texte: " & Err.Description, PROC_NAME, MODULE_NAME
    TruncateWithEllipsis = text ' Retourner le texte original en cas d'erreur
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
    
    ' LogDebug PROC_NAME & "_Called", "Generating formatted timestamp.", PROC_NAME, MODULE_NAME
    ' This function might be called very frequently; logging every call could be noisy.
    ' Consider logging only if in debug mode or if it's a critical utility.
    If ElyseCore_System.IsDebugMode() Then
        LogDebug PROC_NAME & "_DebugCall", "GetCurrentFormattedTimestamp called.", PROC_NAME, MODULE_NAME
    End If

    GetCurrentFormattedTimestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
    GetCurrentFormattedTimestamp = "ErrorInTimestamp" ' Default error return
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    Const PROC_NAME As String = "FileExists"
    On Error GoTo ErrorHandler

    LogDebug PROC_NAME & "_Start", "Checking if file exists: " & filePath, PROC_NAME, MODULE_NAME
    
    Dim fso As Object ' FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FileExists(filePath) Then
        FileExists = True
        LogDebug PROC_NAME & "_Found", "File found: " & filePath, PROC_NAME, MODULE_NAME
    Else
        FileExists = False
        LogDebug PROC_NAME & "_NotFound", "File not found: " & filePath, PROC_NAME, MODULE_NAME
    End If
    
    Set fso = Nothing
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
    FileExists = False ' Default to false on error (e.g., invalid path characters)
End Function

Public Sub ShowUtilityMessage(ByVal message As String, Optional ByVal title As String = "Utility Message")
    Const PROC_NAME As String = "ShowUtilityMessage"
    On Error GoTo ErrorHandler
    
    LogInfo PROC_NAME & "_Display", "Displaying utility message. Title: " & title & ", Message: " & Left(message, 100) & "...", PROC_NAME, MODULE_NAME
    
    ' Original: MsgBox message, vbInformation, title
    ElyseMessageBox_System.ShowInfoMessage title, message
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
    ' Fallback or log that message box couldn't be shown, though ElyseMessageBox_System should handle its own errors.
End Sub