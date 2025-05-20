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