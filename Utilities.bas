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