Option Explicit

Private Const RAGIC_API_KEY As String = "WUJ1UllGWVVyUzRQY3I0Rm0rT2llMWxlOFZxTnQzc092ZlRSU1F0SkJpelFNVFludWQrWHBuamxQMldUVWNTWnJvK1B6RWYzNzR5SDJ6RjdsTTVUcmc9PQ=="
Private Const RAGIC_CSV_URL As String = "https://ragic.elyse.energy/default/simulation-files/1.csv?f=all"

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

Private Function GetColumnIndex(ByVal headerLine As String, columnName As String) As Long
    Dim headers() As String, i As Long
    Dim currentProcessedHeader As String
    
    headers = ParseCSVLine(headerLine) ' This should return fields already trimmed and unquoted
    
    For i = 0 To UBound(headers)
        currentProcessedHeader = headers(i) ' Should be trimmed by ParseCSVLine
        
        ' BOMs (Byte Order Marks) only affect the very first field derived from the start of a stream.
        If i = 0 Then
            ' Check for UTF-8 BOM (ï»¿ which are ChrW(&HEF), ChrW(&HBB), ChrW(&HBF))
            If Len(currentProcessedHeader) >= 3 And Left(currentProcessedHeader, 3) = ChrW(&HEF) & ChrW(&HBB) & ChrW(&HBF) Then
                currentProcessedHeader = Mid(currentProcessedHeader, 4)
            ' Check for UTF-16LE BOM (U+FEFF)
            ElseIf Len(currentProcessedHeader) >= 1 And AscW(Left(currentProcessedHeader, 1)) = &HFEFF Then
                currentProcessedHeader = Mid(currentProcessedHeader, 2)
            ' Check for UTF-16BE BOM (U+FFFE)
            ElseIf Len(currentProcessedHeader) >= 1 And AscW(Left(currentProcessedHeader, 1)) = &HFFFE Then
                currentProcessedHeader = Mid(currentProcessedHeader, 2)
            End If
        End If
        
        currentProcessedHeader = Trim(currentProcessedHeader)
        
        If currentProcessedHeader = columnName Then
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
    Dim charIndex As Long, resultCount As Long
    Dim char As String
    Dim lineLen As Long
    
    lineLen = Len(line)
    
    If lineLen = 0 Then
        ReDim result(0 To 0)
        result(0) = ""
        ParseCSVLine = result
        Exit Function
    End If
    
    Dim tempCharIndex As Long, tempChar As String, tempInQuotes As Boolean, numFields As Long
    numFields = 1 ' Always at least one field
    tempInQuotes = False
    For tempCharIndex = 1 To lineLen
        tempChar = Mid(line, tempCharIndex, 1)
        If tempChar = """" Then
            If tempInQuotes And tempCharIndex < lineLen And Mid(line, tempCharIndex + 1, 1) = """" Then
                tempCharIndex = tempCharIndex + 1 ' Skip next quote (escaped)
            Else
                tempInQuotes = Not tempInQuotes
            End If
        ElseIf tempChar = "," And Not tempInQuotes Then
            numFields = numFields + 1
        End If
    Next tempCharIndex
    
    ReDim result(0 To numFields - 1)

    field = ""
    inQuotes = False
    resultCount = 0
    charIndex = 1
    
    While charIndex <= lineLen
        char = Mid(line, charIndex, 1)
        
        If char = """" Then
            ' Check for escaped quote: "" inside quotes
            If inQuotes And charIndex < lineLen And Mid(line, charIndex + 1, 1) = """" Then
                field = field & """" ' Append one quote
                charIndex = charIndex + 1 ' Skip the second quote of the pair
            Else
                inQuotes = Not inQuotes ' Toggle quote state (entering or exiting a quoted field)
            End If
        ElseIf char = "," And Not inQuotes Then
            result(resultCount) = Trim(field) ' Store the trimmed field
            field = "" ' Reset for next field
            resultCount = resultCount + 1
        Else
            field = field & char ' Append character to current field
        End If
        charIndex = charIndex + 1
    Wend
    
    ' Add the last field
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
        MsgBox "Format CSV invalide : colonnes 'Name', 'Version' ou 'File' manquantes.", vbCritical
        Exit Sub
    End If
    
    ' Chercher la dernière version de l'addin
    Dim i As Long, fields() As String
    Dim maxIdx As Long ' To find the highest index needed
    
    maxIdx = nameIdx
    If versionIdx > maxIdx Then maxIdx = versionIdx
    If fileIdx > maxIdx Then maxIdx = fileIdx
            
    For i = 1 To UBound(lines)
        If Trim(lines(i)) = "" Then GoTo NextLineInLoop ' Skip empty lines if any
        
        fields = ParseCSVLine(lines(i))
        
        ' Ensure fields array is large enough for all expected indices
        If UBound(fields) >= maxIdx Then
            ' fields are already unquoted and trimmed by ParseCSVLine
            ' Assuming "Addin Elyse Energy" is the exact string expected in the 'name' column
            If fields(nameIdx) = "Addin Elyse Energy" Then
                remoteVersion = fields(versionIdx)
                fileUrl = fields(fileIdx)
                Exit For ' Found the addin entry
            End If
        End If
NextLineInLoop:
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
            GetLocalAddinPath = file.path
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
    
    ' Si l'URL pointe vers Ragic, ajouter le header d'authentification
    If InStr(1, url, "ragic.elyse.energy", vbTextCompare) > 0 Then
        http.setRequestHeader "Authorization", "Basic " & RAGIC_API_KEY
    End If
    
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
    
    http.Open "GET", url, False
    
    ' Si l'URL pointe vers Ragic, ajouter le header d'authentification
    If InStr(1, url, "ragic.elyse.energy", vbTextCompare) > 0 Then
        http.setRequestHeader "Authorization", "Basic " & RAGIC_API_KEY
    End If
    
    Debug.Print "URL de téléchargement: " & url
    
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
        Debug.Print "Erreur HTTP: " & http.Status & " - " & http.statusText
        DownloadFile = False
    End If
    Exit Function
errh:
    Debug.Print "Erreur de téléchargement: " & Err.Description
    DownloadFile = False
End Function

Private Function URLEncode(ByVal text As String) As String
    Dim length As Long
    Dim pos As Long
    Dim char As String
    Dim asciiVal As Integer
    Dim result As String
    
    length = Len(text)
    
    For pos = 1 To length
        char = Mid(text, pos, 1)
        asciiVal = AscW(char)
        
        ' Ne pas encoder les caractères alphanumériques et certains caractères spéciaux
        If (asciiVal >= 48 And asciiVal <= 57) Or _
           (asciiVal >= 65 And asciiVal <= 90) Or _
           (asciiVal >= 97 And asciiVal <= 122) Or _
           char = "-" Or char = "_" Or char = "." Or char = "~" Then
            result = result & char
        Else
            ' Encoder en %HH où HH est la valeur hexadécimale
            result = result & "%" & Right("0" & Hex(AscB(char)), 2)
        End If
    Next pos
    
    URLEncode = result
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
    Dim tempCsvPath As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim headerLine As String
    Dim csvFileDownloaded As Boolean
    Dim headerLineRead As Boolean
    
    Debug.Print "=== DÉBUT TEST CONNECTION RAGIC ==="
    
    ' Générer un nom unique pour le CSV temporaire
    tempCsvPath = GetTempFilePath("csv")
    headerLine = "" ' Initialize
    csvFileDownloaded = False
    headerLineRead = False ' Initialize

    ' Test 1: Téléchargement CSV via DownloadFile (mimicking main logic)
    If DownloadFile(RAGIC_CSV_URL, tempCsvPath) Then
        LogTest "Download CSV (via DownloadFile)", True, "File downloaded to: " & tempCsvPath
        csvFileDownloaded = True
        
        ' Lire la première ligne du fichier téléchargé (mimicking CheckRagicForUpdate more closely)
        Dim fileNum As Integer, fileContent As String
        
        On Error Resume Next
        fileNum = FreeFile
        Open tempCsvPath For Input As #fileNum
        If Err.Number = 0 Then
            If Not EOF(fileNum) Then
                fileContent = Input$(LOF(fileNum), fileNum) ' Read whole file
            End If
            Close #fileNum
            
            If fileContent <> "" Then
                Dim csvLines() As String
                csvLines = Split(fileContent, vbLf) ' Split by Line Feed
                
                If UBound(csvLines) >= 0 Then
                    headerLine = csvLines(0)
                    ' Remove trailing CR if present (from CRLF line endings)
                    If Len(headerLine) > 0 And Right(headerLine, 1) = vbCr Then
                        headerLine = Left(headerLine, Len(headerLine) - 1)
                    End If
                    headerLineRead = True
                End If ' Else: fileContent was not empty but resulted in no lines (e.g. only LFs)
            End If ' Else: fileContent was empty
            LogTest "Read Header Line from File", headerLineRead And headerLine <> "", "Header: '" & headerLine & "'"
        Else
            LogTest "Read Header Line from File", False, "Error reading file: " & Err.Description
        End If
        On Error GoTo 0
    Else
        LogTest "Download CSV (via DownloadFile)", False, "Failed to download file"
    End If
    
    ' Nettoyer le fichier CSV temporaire
    DeleteFileIfExists tempCsvPath

    If csvFileDownloaded And headerLineRead And headerLine <> "" Then
        ' Test 2: Vérification contenu (simple check on header if it looks like CSV)
        LogTest "Contenu CSV (Header basic check)", InStr(headerLine, ",") > 0 Or InStr(headerLine, ";") > 0, "Header appears to be CSV-like: '" & headerLine & "'"
        
        ' Test 3: Parse header
        Dim nameIdx As Long
        nameIdx = GetColumnIndex(headerLine, "Name") ' Pass the actual header line
        LogTest "Parse Header", nameIdx >= 0, "Index colonne Name: " & nameIdx
    ElseIf csvFileDownloaded And headerLineRead And headerLine = "" Then
        LogTest "Parse Header", False, "Header line was empty after reading file."
    Else
        LogTest "Parse Header", False, "Skipped due to download/read failure or empty header."
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

Public Sub TestFileDownload()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== DÉBUT TEST TÉLÉCHARGEMENT FICHIER ==="
    
    ' 1. D'abord récupérer les informations du fichier via l'API
    Dim jsonResponse As String
    jsonResponse = GetRagicFileInfo()
    
    If jsonResponse = "" Then
        Debug.Print "Erreur : impossible d'obtenir les informations du fichier"
        Exit Sub
    End If
    
    ' 2. Extraire le nom du fichier du JSON
    Dim fileName As String
    fileName = ExtractFileNameFromJson(jsonResponse)
    
    If fileName = "" Then
        Debug.Print "Erreur : impossible de trouver le nom du fichier dans la réponse"
        Exit Sub
    End If
    
    Debug.Print "Nom de fichier trouvé : " & fileName
    
    ' Préparation du test
    Dim tempPath As String
    tempPath = Environ$("TEMP") & "\test_download.xlam"
    
    ' Construction de l'URL selon la doc Ragic
    Dim testFileUrl As String
    testFileUrl = "https://ragic.elyse.energy/sims/file.jsp?a=default&f=" & fileName
    
    Debug.Print "Tentative de téléchargement depuis: " & testFileUrl
    
    ' Test du téléchargement avec l'URL construite
    Dim success As Boolean
    success = DownloadFile(testFileUrl, tempPath)
    
    ' Vérification des résultats
    If success Then
        If Dir(tempPath) <> "" Then
            Debug.Print "Test réussi : fichier téléchargé à " & tempPath
            'Kill tempPath ' Nettoyage
        Else
            Debug.Print "Erreur : fichier non trouvé après téléchargement"
        End If
    Else
        Debug.Print "Erreur : échec du téléchargement"
    End If
    
    Debug.Print "=== FIN TEST TÉLÉCHARGEMENT FICHIER ==="
    Exit Sub

ErrorHandler:
    Debug.Print "Erreur lors du test : " & Err.Description
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
    TestFileDownload
    
    Debug.Print "========================================"
    Debug.Print "FIN DES TESTS - " & Now()
    Debug.Print "========================================"
End Sub

Private Function GetRagicFileInfo() As String
    ' Construction de l'URL pour obtenir les infos du fichier (JSON)
    Dim apiUrl As String
    apiUrl = "https://ragic.elyse.energy/default/simulation-files/1?api"
    
    ' Appel à l'API pour obtenir le JSON
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    On Error GoTo ErrorHandler
    
    http.Open "GET", apiUrl, False
    http.setRequestHeader "Authorization", "Basic " & RAGIC_API_KEY
    http.send
    
    If http.Status = 200 Then
        GetRagicFileInfo = http.responseText
        Debug.Print "Réponse JSON reçue: " & http.responseText
    Else
        Debug.Print "Erreur lors de la requête: " & http.Status & " - " & http.statusText
        GetRagicFileInfo = ""
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors de l'appel API: " & Err.Description
    GetRagicFileInfo = ""
End Function

Private Function ExtractFileNameFromJson(jsonStr As String) As String
    ' Nettoyer la réponse JSON (enlever les sauts de ligne pour faciliter le parsing)
    jsonStr = Replace(Replace(jsonStr, vbCrLf, ""), vbLf, "")
    
    On Error GoTo ErrorHandler
    
    ' On cherche tous les enregistrements qui contiennent "Addin Elyse Energy"
    Dim regex As Object, matches As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Pattern pour trouver un objet JSON qui contient "Name":"Addin Elyse Energy"
    regex.Pattern = "\{[^\}]*""Name""\s*:\s*""Addin Elyse Energy""[^\}]*\}"
    regex.Global = True
    regex.IgnoreCase = True
    
    Set matches = regex.Execute(jsonStr)
    
    If matches.Count = 0 Then
        Debug.Print "Aucun enregistrement avec 'Addin Elyse Energy' trouvé"
        ExtractFileNameFromJson = ""
        Exit Function
    End If
    
    ' Variables pour trouver la version la plus récente
    Dim i As Long
    Dim currentMatch As Object
    Dim latestVersion As String
    Dim latestVersionMatch As String
    Dim latestFileName As String
    
    ' Pour chaque enregistrement trouvé
    For i = 0 To matches.Count - 1
        Set currentMatch = matches(i)
        
        ' Extraire la version
        Dim versionRegex As Object
        Set versionRegex = CreateObject("VBScript.RegExp")
        versionRegex.Pattern = """Version""\s*:\s*""([^""]+)"""
        versionRegex.Global = False
        
        Dim versionMatches As Object
        Set versionMatches = versionRegex.Execute(currentMatch.Value)
        
        If versionMatches.Count > 0 Then
            Dim currentVersion As String
            currentVersion = versionMatches(0).SubMatches(0)
            
            ' Si c'est la première version ou si cette version est plus récente
            If latestVersion = "" Or CompareVersions(currentVersion, latestVersion) > 0 Then
                latestVersion = currentVersion
                latestVersionMatch = currentMatch.Value
            End If
        End If
    Next i
    
    ' Une fois qu'on a trouvé l'enregistrement le plus récent, on extrait le nom du fichier
    If latestVersionMatch <> "" Then
        Dim fileRegex As Object
        Set fileRegex = CreateObject("VBScript.RegExp")
        fileRegex.Pattern = """File""\s*:\s*""([^""]+)"""
        fileRegex.Global = False
        
        Dim fileMatches As Object
        Set fileMatches = fileRegex.Execute(latestVersionMatch)
        
        If fileMatches.Count > 0 Then
            ExtractFileNameFromJson = fileMatches(0).SubMatches(0)
            Debug.Print "Dernière version trouvée : " & latestVersion & ", fichier : " & ExtractFileNameFromJson
        Else
            Debug.Print "Fichier non trouvé dans l'enregistrement le plus récent"
            ExtractFileNameFromJson = ""
        End If
    Else
        ExtractFileNameFromJson = ""
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Erreur lors de l'extraction du nom de fichier: " & Err.Description
    ExtractFileNameFromJson = ""
End Function

' Compare deux versions au format X.Y.Z
' Retourne : 1 si version1 > version2, -1 si version1 < version2, 0 si égales
Private Function CompareVersions(version1 As String, version2 As String) As Integer
    Dim v1Parts() As String, v2Parts() As String
    Dim i As Long, v1Num As Long, v2Num As Long
    
    v1Parts = Split(version1, ".")
    v2Parts = Split(version2, ".")
    
    For i = 0 To 2 ' On compare les 3 parties (X.Y.Z)
        If i < UBound(v1Parts) + 1 And i < UBound(v2Parts) + 1 Then
            v1Num = CLng(v1Parts(i))
            v2Num = CLng(v2Parts(i))
            
            If v1Num > v2Num Then
                CompareVersions = 1
                Exit Function
            ElseIf v1Num < v2Num Then
                CompareVersions = -1
                Exit Function
            End If
        ElseIf i < UBound(v1Parts) + 1 Then
            ' version1 a plus de parties (ex: 1.0.1 vs 1.0)
            CompareVersions = 1
            Exit Function
        ElseIf i < UBound(v2Parts) + 1 Then
            ' version2 a plus de parties
            CompareVersions = -1
            Exit Function
        End If
    Next i
    
    ' Les versions sont identiques
    CompareVersions = 0
End Function
