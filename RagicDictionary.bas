Option Explicit

Public RagicFieldDict As Object

Public Sub LoadRagicDictionary()
    Dim queryName As String
    queryName = "RagicDictionary"

    Dim formula As String
    formula = GenerateDictionaryQuery()

    If QueryExists(queryName) Then
        ThisWorkbook.Queries(queryName).Formula = formula
    Else
        ThisWorkbook.Queries.Add queryName, formula
    End If

    Utilities.InitializePQData
    Dim startCol As Long
    startCol = Utilities.GetLastColumn(wsPQData)
    LoadQueries.LoadQuery queryName, wsPQData, wsPQData.Cells(1, startCol + 1)

    Dim lo As ListObject
    Set lo = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(queryName))
    If lo Is Nothing Then Exit Sub

    Set RagicFieldDict = CreateObject("Scripting.Dictionary")

    Dim sheetIdx As Long, fieldIdx As Long, memoIdx As Long
    On Error Resume Next
    sheetIdx = lo.ListColumns("SheetName").Index
    fieldIdx = lo.ListColumns("Field Name").Index
    memoIdx = lo.ListColumns("Memo").Index
    On Error GoTo 0
    If sheetIdx = 0 Or fieldIdx = 0 Or memoIdx = 0 Then
        DataLoaderManager.CleanupPowerQuery queryName
        Exit Sub
    End If

    Dim i As Long
    Dim key As String
    For i = 1 To lo.DataBodyRange.Rows.Count
        key = CStr(lo.DataBodyRange.Cells(i, sheetIdx).Value) & "|" & _
              CStr(lo.DataBodyRange.Cells(i, fieldIdx).Value)
        If Not RagicFieldDict.Exists(key) Then
            RagicFieldDict.Add key, CStr(lo.DataBodyRange.Cells(i, memoIdx).Value)
        End If
    Next i

    DataLoaderManager.CleanupPowerQuery queryName
End Sub

Private Function GenerateDictionaryQuery() As String
    Dim url As String
    url = "https://ragic.elyse.energy/default/matching-matrix/6.csv"

    GenerateDictionaryQuery = "let" & vbCrLf & _
        "    Source = Csv.Document(Web.Contents(\"" & url & "\"),[Delimiter=\";\", Encoding=65001])," & vbCrLf & _
        "    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])," & vbCrLf & _
        "    Filtered = Table.SelectRows(PromotedHeaders, each [SheetName] <> null and [Field Name] <> null)" & vbCrLf & _
        "in" & vbCrLf & _
        "    Filtered"
End Function

Private Function QueryExists(queryName As String) As Boolean
    On Error Resume Next
    Dim q As Object
    Set q = ThisWorkbook.Queries(queryName)
    QueryExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function IsFieldHidden(sheetName As String, fieldName As String) As Boolean
    If RagicFieldDict Is Nothing Then Exit Function
    Dim key As String
    key = sheetName & "|" & fieldName
    If RagicFieldDict.Exists(key) Then
        IsFieldHidden = InStr(1, RagicFieldDict(key), "Hidden", vbTextCompare) > 0
    Else
        IsFieldHidden = False
    End If
End Function
