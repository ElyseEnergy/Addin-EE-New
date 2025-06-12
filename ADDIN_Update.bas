Option Explicit

Private Const RAGIC_API_KEY As String = "WUJ1UllGWVVyUzRQY3I0Rm0rT2llMWxlOFZxTnQzc092ZlRSU1F0SkJpelFNVFludWQrWHBuamxQMldUVWNTWnJvK1B6RWYzNzR5SDJ6RjdsTTVUcmc9PQ=="
Private Const RAGIC_CSV_URL As String = "https://ragic.elyse.energy/default/simulation-files/1.csv?APIKey=" & RAGIC_API_KEY

Private Function GetTempFilePath(Optional extension As String = "") As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetTempFilePath = fso.GetSpecialFolder(2) & "\" & fso.GetTempName()
    If extension <> "" Then
        GetTempFilePath = GetTempFilePath & "." & extension
    End If
End Function

Private Sub DeleteFileIfExists(filePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If fso.FileExists(filePath) Then
        fso.DeleteFile filePath, True
    End If
    On Error GoTo 0
End Sub

Private Function GetColumnIndex(headerLine As String, columnName As String) As Long
    Dim headers() As String, i As Long
    ' Split en gérant les guillemets
    headers = ParseCSVLine(headerLine)
    For i = 0 To UBound(headers)
        ' Enlever les guillemets si présents
        headers(i) = Replace(Replace(headers(i), """", ""), " ", "")
        If headers(i) = columnName Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = -1
End Function

Private Function ParseCSVLine(line As String) As String()
    Dim result() As String
    Dim field As String
    Dim inQuotes As Boolean
    Dim i As Long, resultCount As Long
    Dim char As String
    
    ReDim result(0)
    field = ""
    inQuotes = False
    
    For i = 1 To Len(line)
        char = Mid(line, i, 1)
        
        If char = """" Then
            inQuotes = Not inQuotes
        ElseIf char = "," And Not inQuotes Then
            ReDim Preserve result(resultCount)
            result(resultCount) = Trim(field)
            field = ""
            resultCount = resultCount + 1
        Else
            field = field & char
        End If
    Next i
    
    ' Ajouter le dernier champ
    ReDim Preserve result(resultCount)
    result(resultCount) = Trim(field)
    
    ParseCSVLine = result
End Function

Sub CheckRagicForUpdate()
    Dim localVersion As String, remoteVersion As String, fileUrl As String
    Dim addinPath As String, tempCsvPath As String, tempXlamPath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Générer un nom unique pour le CSV temporaire
    tempCsvPath = GetTempFilePath("csv")
    
    ' Trouver le fichier addin local
    addinPath = GetLocalAddinPath()
    If addinPath = "" Then
        MsgBox "Impossible de trouver l'addin EE dans le dossier AddIns.", vbCritical
        Exit Sub
    End If
    
    ' Récupérer la version locale depuis le nom du fichier
    localVersion = GetLocalVersion(addinPath)
    If localVersion = "" Then
        MsgBox "Impossible de déterminer la version locale de l'addin.", vbCritical
        Exit Sub
    End If
    
    remoteVersion = ""
    fileUrl = ""
    
    ' 1. Télécharger le CSV dans un fichier temporaire
    If Not DownloadFile(RAGIC_CSV_URL, tempCsvPath) Then
        MsgBox "Impossible de télécharger la base Ragic.", vbCritical
        DeleteFileIfExists tempCsvPath
        Exit Sub
    End If
    
    ' 2. Lire et parser le CSV temporaire
    Dim csvContent As String, lines() As String
    On Error Resume Next
    With fso.OpenTextFile(tempCsvPath, 1, False, -1) ' 1 = ForReading
        csvContent = .ReadAll()
        .Close
    End With
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de la lecture du CSV: " & Err.Description, vbCritical
        DeleteFileIfExists tempCsvPath
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Nettoyer le CSV temporaire, on n'en a plus besoin
    DeleteFileIfExists tempCsvPath
    
    ' Parser le contenu du CSV
    lines = Split(csvContent, vbLf)
    If UBound(lines) < 1 Then
        MsgBox "Format CSV invalide", vbCritical
        Exit Sub
    End If
    
    ' Trouver les indices des colonnes
    Dim nameIdx As Long, versionIdx As Long, fileIdx As Long
    nameIdx = GetColumnIndex(lines(0), "Name")
    versionIdx = GetColumnIndex(lines(0), "Version")
    fileIdx = GetColumnIndex(lines(0), "File")
    
    If nameIdx = -1 Or versionIdx = -1 Or fileIdx = -1 Then
        MsgBox "Format CSV invalide : colonnes manquantes", vbCritical
        Exit Sub
    End If
    
    ' Chercher la dernière version de l'addin
    Dim i As Long, fields() As String
    For i = 1 To UBound(lines)
        fields = ParseCSVLine(lines(i))
        If UBound(fields) >= fileIdx Then
            ' Enlever les guillemets si présents
            fields(nameIdx) = Replace(fields(nameIdx), """", "")
            If Trim(fields(nameIdx)) = "Addin Elyse Energy" Then
                remoteVersion = Replace(fields(versionIdx), """", "")
                fileUrl = Replace(fields(fileIdx), """", "")
                Exit For
            End If
        End If
    Next i
    
    If remoteVersion = "" Or fileUrl = "" Then
        MsgBox "Aucune version trouvée dans la base Ragic.", vbCritical
        Exit Sub
    End If
    
    ' 3. Comparer les versions
    If localVersion <> remoteVersion Then
        If MsgBox("Nouvelle version disponible (" & remoteVersion & "). Mettre à jour ?", vbYesNo + vbQuestion) = vbYes Then
            ' Télécharger dans le dossier temp avec le nom final
            tempXlamPath = fso.BuildPath(fso.GetSpecialFolder(2), "EE Addin_v" & remoteVersion & ".xlam")
            DeleteFileIfExists tempXlamPath ' Au cas où
            
            If DownloadFile(fileUrl, tempXlamPath) Then
                ' Copier vers la destination finale
                On Error Resume Next
                fso.CopyFile tempXlamPath, addinPath, True
                If Err.Number <> 0 Then
                    MsgBox "Erreur lors de la copie du fichier: " & Err.Description, vbCritical
                    DeleteFileIfExists tempXlamPath
                    Exit Sub
                End If
                On Error GoTo 0
                
                ' Nettoyer le fichier temporaire
                DeleteFileIfExists tempXlamPath
                
                MsgBox "Mise à jour effectuée. Excel va se fermer.", vbInformation
                Application.Quit
            Else
                DeleteFileIfExists tempXlamPath
                MsgBox "Erreur lors du téléchargement du fichier.", vbCritical
            End If
        End If
    End If
End Sub

Private Function GetLocalAddinPath() As String
    Dim folder As String, fso As Object, file As Object
    folder = Environ$("APPDATA") & "\Microsoft\AddIns"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folder) Then
        GetLocalAddinPath = ""
        Exit Function
    End If
    
    For Each file In fso.GetFolder(folder).Files
        If file.Name Like "EE Addin_v*.xlam" Then
            GetLocalAddinPath = file.Path
            Exit Function
        End If
    Next
    GetLocalAddinPath = ""
End Function

Private Function GetLocalVersion(addinPath As String) As String
    Dim fso As Object, fileName As String, matches As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = fso.GetFileName(addinPath)
    
    ' Cherche un motif du type vX.Y.Z dans le nom
    With CreateObject("VBScript.RegExp")
        .Pattern = "v(\d+\.\d+\.\d+)"
        .IgnoreCase = True
        .Global = False
        If .Test(fileName) Then
            Set matches = .Execute(fileName)
            GetLocalVersion = matches(0).SubMatches(0)
        Else
            GetLocalVersion = ""
        End If
    End With
End Function

Private Function DownloadText(url As String) As String
    On Error GoTo errh
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send
    If http.Status = 200 Then
        DownloadText = http.responseText
    End If
    Exit Function
errh:
    DownloadText = ""
End Function

Private Function DownloadFile(url As String, destPath As String) As Boolean
    On Error GoTo errh
    Dim http As Object, ado As Object, arr() As Byte
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Ajouter l'API key si pas déjà présente
    If InStr(url, "APIKey=") = 0 Then
        url = url & IIf(InStr(url, "?") > 0, "&", "?") & "APIKey=" & RAGIC_API_KEY
    End If
    
    http.Open "GET", url, False
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send
    If http.Status = 200 Then
        arr = http.responseBody
        Set ado = CreateObject("ADODB.Stream")
        ado.Type = 1
        ado.Open
        ado.Write arr
        ado.SaveToFile destPath, 2
        ado.Close
        DownloadFile = True
    Else
        DownloadFile = False
    End If
    Exit Function
errh:
    DownloadFile = False
End Function

Private Function ChampIndex(nomChamp As String, headerLine As String) As Long
    Dim headers() As String, i As Long
    headers = SplitCSV(headerLine)
    For i = 0 To UBound(headers)
        If Trim(headers(i)) = nomChamp Then
            ChampIndex = i
            Exit Function
        End If
    Next i
    ChampIndex = -1
End Function

Private Function SplitCSV(line As String) As Variant
    ' Simple split, à améliorer si tu as des champs avec des virgules dans des quotes
    SplitCSV = Split(line, ",")
End Function

' ============================================================================
' FONCTIONS DE TEST
' ============================================================================

Private Sub LogTest(testName As String, result As Boolean, Optional details As String = "")
    Debug.Print "TEST - " & testName & ": " & IIf(result, "OK", "ÉCHEC") & IIf(details <> "", " - " & details, "")
End Sub

Public Sub TestCSVParsing()
    ' Test du parsing CSV avec différents cas
    Dim testLine As String, result() As String
    
    Debug.Print "=== DÉBUT TEST PARSING CSV ==="
    
    ' Test 1: Champs simples
    testLine = """id"",""name"",""version"""
    result = ParseCSVLine(testLine)
    LogTest "CSV Simple", result(0) = "id" And result(1) = "name" And result(2) = "version"
    
    ' Test 2: Champs avec virgules
    testLine = """id"",""name, with comma"",""v1.0"""
    result = ParseCSVLine(testLine)
    LogTest "CSV avec virgules", result(1) = "name, with comma"
    
    ' Test 3: Champs avec guillemets échappés
    testLine = """id"",""name """"quoted"""" here"",""version"""
    result = ParseCSVLine(testLine)
    LogTest "CSV avec guillemets", result(1) = "name ""quoted"" here"
    
    Debug.Print "=== FIN TEST PARSING CSV ==="
End Sub

Public Sub TestVersionExtraction()
    Dim testPath As String, version As String
    
    Debug.Print "=== DÉBUT TEST EXTRACTION VERSION ==="
    
    ' Test 1: Format standard
    testPath = "C:\Test\EE Addin_v1.0.0.xlam"
    version = GetLocalVersion(testPath)
    LogTest "Version standard", version = "1.0.0", "Version extraite: " & version
    
    ' Test 2: Format invalide
    testPath = "C:\Test\EE Addin.xlam"
    version = GetLocalVersion(testPath)
    LogTest "Version invalide", version = "", "Devrait retourner vide"
    
    Debug.Print "=== FIN TEST EXTRACTION VERSION ==="
End Sub

Public Sub TestTempFileOperations()
    Dim tempPath As String, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Debug.Print "=== DÉBUT TEST FICHIERS TEMPORAIRES ==="
    
    ' Test 1: Création fichier temp
    tempPath = GetTempFilePath("txt")
    LogTest "Création chemin temp", tempPath <> "", "Chemin: " & tempPath
    
    ' Test 2: Écriture/Suppression
    On Error Resume Next
    Open tempPath For Output As #1
    Print #1, "Test"
    Close #1
    LogTest "Écriture fichier temp", Err.Number = 0, "Erreur: " & Err.Description
    
    DeleteFileIfExists tempPath
    LogTest "Suppression fichier temp", Not fso.FileExists(tempPath)
    
    Debug.Print "=== FIN TEST FICHIERS TEMPORAIRES ==="
End Sub

Public Sub TestRagicConnection()
    Dim content As String
    
    Debug.Print "=== DÉBUT TEST CONNECTION RAGIC ==="
    
    ' Test 1: Téléchargement CSV
    content = DownloadText(RAGIC_CSV_URL)
    LogTest "Download CSV", content <> "", IIf(content = "", "Échec téléchargement", "Taille: " & Len(content))
    
    If content <> "" Then
        ' Test 2: Vérification contenu
        LogTest "Contenu CSV", InStr(content, "Addin Elyse Energy") > 0, "Addin trouvé dans le CSV"
        
        ' Test 3: Parse header
        Dim nameIdx As Long
        nameIdx = GetColumnIndex(Split(content, vbLf)(0), "Name")
        LogTest "Parse Header", nameIdx >= 0, "Index colonne Name: " & nameIdx
    End If
    
    Debug.Print "=== FIN TEST CONNECTION RAGIC ==="
End Sub

Public Sub TestAddinPathDetection()
    Dim path As String
    
    Debug.Print "=== DÉBUT TEST DETECTION ADDIN ==="
    
    path = GetLocalAddinPath()
    LogTest "Détection Addin", path <> "", "Chemin: " & path
    
    If path <> "" Then
        Dim version As String
        version = GetLocalVersion(path)
        LogTest "Version Addin", version <> "", "Version: " & version
    End If
    
    Debug.Print "=== FIN TEST DETECTION ADDIN ==="
End Sub

Public Sub RunAllTests()
    Debug.Print "========================================"
    Debug.Print "DÉBUT DES TESTS - " & Now()
    Debug.Print "========================================"
    
    TestCSVParsing
    TestVersionExtraction
    TestTempFileOperations
    TestRagicConnection
    TestAddinPathDetection
    
    Debug.Print "========================================"
    Debug.Print "FIN DES TESTS - " & Now()
    Debug.Print "========================================"
End Sub