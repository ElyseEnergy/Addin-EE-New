Attribute VB_Name = "FileManager"

Option Explicit

Private Const PROC_NAME As String = "FileManager"
Private Const MODULE_NAME As String = "FileManager"
Private Const TEMP_FILE_PREFIX As String = "TEMP_UPLOAD_"
Private Const BOUNDARY As String = "------------------------boundary123456789"

' Callback pour le bouton d'upload
Public Sub ProcessFileUpload(ByVal control As IRibbonControl)
    UploadCurrentFile
End Sub

Private Sub UploadCurrentFile()
    Dim tempFilePath As String ' Déclarée au début pour être accessible partout
    tempFilePath = ""
    
    On Error GoTo ErrorHandler
    
    ' Récupérer le fichier actif
    Dim currentFile As String
    currentFile = ActiveWorkbook.FullName
    Log "upload", "Fichier à uploader : " & currentFile, DEBUG_LEVEL, "UploadCurrentFile", MODULE_NAME
    
    ' Vérifier si c'est un fichier SharePoint (commence par http ou https)
    If LCase(Left(currentFile, 4)) = "http" Then
        ' Créer une copie temporaire
        tempFilePath = Environ("TEMP") & "\" & TEMP_FILE_PREFIX & Format(Now, "yyyymmddhhnnss") & ".xlsm"
        ActiveWorkbook.SaveCopyAs tempFilePath
        currentFile = tempFilePath
        Log "upload", "Fichier SharePoint détecté, copie temporaire créée : " & currentFile, DEBUG_LEVEL, "UploadCurrentFile", MODULE_NAME
    End If
    
    ' Demander la version
    Dim version As String
    version = InputBox("Quelle est la version de ce fichier ?", "Version", "1.0")
    If version = "" Then Exit Sub ' L'utilisateur a annulé
    
    ' Préparer les données obligatoires
    Dim data As New Collection
    With data
        .Add Array("1001623", "Real")           ' Fake ? (Real/Fake)
        .Add Array("1001060", GetName())        ' Name
        .Add Array("1001044", version)          ' Version
        .Add Array("1001068", Utilities.GetUserEmail())      ' Author
        .Add Array("1001069", Format(Date, "yyyy-mm-dd")) ' Delivery date
        .Add Array("1001045", "Initial upload") ' Change log
        .Add Array("1001063", "Internal simulation only (expert)") ' Can be use for ?
        .Add Array("1005174", "Planning")       ' Type
        .Add Array("1001066", "methanol")       ' Main molecule/expertise
        .Add Array("1001067", "average per year") ' Main timescale
    End With
    
    Log "upload", "Données préparées pour l'upload :", DEBUG_LEVEL, "UploadCurrentFile", MODULE_NAME
    Dim item As Variant
    For Each item In data
        Log "upload", "  " & item(0) & " = " & item(1), DEBUG_LEVEL, "UploadCurrentFile", MODULE_NAME
    Next item
    
    ' Appeler l'API Ragic
    Dim result As String
    result = UploadToRagic(currentFile, data)
    Log "upload", "Réponse Ragic : " & result, DEBUG_LEVEL, "UploadCurrentFile", MODULE_NAME
    
    ' --- LOGIQUE DE NETTOYAGE ET DE REPONSE AMELIOREE ---
    
    ' 1. Toujours nettoyer le fichier temporaire après l'appel
    If tempFilePath <> "" Then
        Dim killError As Long
        killError = DeleteFile(tempFilePath)
        If killError = 0 Then
            Log "upload", "Fichier temporaire supprimé : " & tempFilePath, DEBUG_LEVEL, "UploadCurrentFile", MODULE_NAME
        End If
    End If
    
    ' 2. Gérer la réponse de Ragic sans utiliser Err.Raise pour les erreurs API
    If InStr(1, Replace(result, " ", ""), """status"":""SUCCESS""") > 0 Then
        Log "upload", "Upload réussi !", INFO_LEVEL, "UploadCurrentFile", MODULE_NAME
        MsgBox "Fichier uploadé avec succès !", vbInformation
    Else
        Log "upload", "Erreur lors de l'upload : " & result, ERROR_LEVEL, "UploadCurrentFile", MODULE_NAME
        MsgBox "L'upload a échoué. La réponse du serveur était :" & vbCrLf & vbCrLf & result, vbCritical, "Echec de l'upload"
    End If
    Exit Sub

ErrorHandler:
    ' Le handler ne gère plus que les erreurs VBA inattendues (ex: réseau coupé)
    ' On tente quand même de nettoyer le fichier temporaire au cas où l'erreur
    ' se serait produite avant le bloc de nettoyage normal.
    If tempFilePath <> "" Then
        If Dir(tempFilePath) <> "" Then ' Vérifie si le fichier existe encore
            Dim killErrorOnError As Long
            killErrorOnError = DeleteFile(tempFilePath)
            If killErrorOnError = 0 Then
                 Log "upload", "Fichier temporaire supprimé après erreur : " & tempFilePath, DEBUG_LEVEL, "UploadCurrentFile", MODULE_NAME
            End If
        End If
    End If
    
    ' Appel au gestionnaire centralisé
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de l'upload du fichier : " & Err.Description
End Sub

Private Function DeleteFile(ByVal filePath As String) As Long
    On Error Resume Next
    Kill filePath
    DeleteFile = Err.Number
    On Error GoTo 0
End Function

Private Function GetName() As String
    GetName = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
End Function

Private Function UploadToRagic(filePath As String, data As Collection) As String
    On Error GoTo ErrorHandler
    
    Log "upload", "Début de l'upload vers Ragic...", DEBUG_LEVEL, "UploadToRagic", MODULE_NAME
    
    ' Créer l'objet HTTP
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo ErrorHandler

    If http Is Nothing Then
        Log "upload", "Impossible de créer l'objet HTTP.", ERROR_LEVEL, "UploadToRagic", MODULE_NAME
        UploadToRagic = "{""status"":""ERROR"",""msg"":""Impossible de créer l'objet HTTP""}"
        Exit Function
    End If

    ' --- AUTHENTIFICATION & URL ---
    ' Utilisation de "Basic" Auth, comme requis pour l'upload de fichiers sur cette API.
    Dim url As String
    url = env.RAGIC_BASE_URL & "simulation-files/1"
    
    Dim apiKey As String
    apiKey = Replace(env.RAGIC_API_KEY, "&", "") ' La clé dans env.bas a un '&' en trop
    
    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Basic " & apiKey
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & BOUNDARY
    Log "upload", "URL : " & url, DEBUG_LEVEL, "UploadToRagic", MODULE_NAME
    Log "upload", "Authorization: Basic ...", DEBUG_LEVEL, "UploadToRagic", MODULE_NAME
    
    ' --- CONSTRUCTION DU CORPS DE LA REQUETE (avec ADODB.Stream pour la fiabilité) ---
    Dim bodyStream As Object, textStream As Object, fileStream As Object
    Set bodyStream = CreateObject("ADODB.Stream")
    bodyStream.Type = 1 ' adTypeBinary
    bodyStream.Open

    Set textStream = CreateObject("ADODB.Stream")
    textStream.Type = 2 ' adTypeText
    textStream.Charset = "iso-8859-1" ' Utiliser un charset simple pour éviter les BOM
    textStream.Open

    Dim crlf As String
    crlf = vbCrLf
    
    ' Ajouter les champs de données
    Dim item As Variant
    For Each item In data
        textStream.SetEOS: textStream.WriteText "--" & BOUNDARY & crlf
        textStream.Position = 0: textStream.CopyTo bodyStream
        
        textStream.SetEOS: textStream.WriteText "Content-Disposition: form-data; name=""" & item(0) & """" & crlf
        textStream.Position = 0: textStream.CopyTo bodyStream
        
        textStream.SetEOS: textStream.WriteText crlf
        textStream.Position = 0: textStream.CopyTo bodyStream
        
        textStream.SetEOS: textStream.WriteText item(1) & crlf
        textStream.Position = 0: textStream.CopyTo bodyStream
    Next item
    
    ' Ajouter les en-têtes de la partie fichier
    textStream.SetEOS: textStream.WriteText "--" & BOUNDARY & crlf
    textStream.Position = 0: textStream.CopyTo bodyStream
    
    textStream.SetEOS: textStream.WriteText "Content-Disposition: form-data; name=""1001040""; filename=""" & GetFileName(filePath) & """" & crlf
    textStream.Position = 0: textStream.CopyTo bodyStream
    
    textStream.SetEOS: textStream.WriteText "Content-Type: application/octet-stream" & crlf
    textStream.Position = 0: textStream.CopyTo bodyStream
    
    textStream.SetEOS: textStream.WriteText crlf
    textStream.Position = 0: textStream.CopyTo bodyStream
    
    ' Ajouter le contenu du fichier
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 1 ' adTypeBinary
    fileStream.Open
    fileStream.LoadFromFile filePath
    Log "upload", "Fichier lu : " & filePath & " (" & fileStream.Size & " octets)", DEBUG_LEVEL, "UploadToRagic", MODULE_NAME
    fileStream.CopyTo bodyStream
    fileStream.Close
    
    ' Ajouter la frontière de fin
    textStream.SetEOS: textStream.WriteText crlf & "--" & BOUNDARY & "--" & crlf
    textStream.Position = 0: textStream.CopyTo bodyStream
    
    textStream.Close

    ' Envoyer le flux binaire complet
    bodyStream.Position = 0
    Log "upload", "Envoi de la requête (" & bodyStream.Size & " octets)...", DEBUG_LEVEL, "UploadToRagic", MODULE_NAME
    http.Send bodyStream.Read
    bodyStream.Close
    
    ' Récupérer et logger la réponse
    Log "upload", "Status : " & http.Status & " " & http.StatusText, DEBUG_LEVEL, "UploadToRagic", MODULE_NAME
    Log "upload", "Headers : " & http.GetAllResponseHeaders(), DEBUG_LEVEL, "UploadToRagic", MODULE_NAME
    Log "upload", "Response : " & http.ResponseText, DEBUG_LEVEL, "UploadToRagic", MODULE_NAME
    
    UploadToRagic = http.ResponseText
    Exit Function

ErrorHandler:
    Log "upload", "Erreur dans UploadToRagic : " & Err.Description, ERROR_LEVEL, "UploadToRagic", MODULE_NAME
    UploadToRagic = "{""status"":""ERROR"",""msg"":""" & Replace(Err.Description, """", "'") & """}"
End Function

Private Function GetFileName(filePath As String) As String
    GetFileName = Mid(filePath, InStrRev(filePath, "\") + 1)
End Function