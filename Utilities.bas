Attribute VB_Name = "Utilities"

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
    GetLastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
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

