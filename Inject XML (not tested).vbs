' filepath: c:\Users\JulienFernandez\OneDrive\Coding\_Projets de code\2025.05 New addin EE perso\InjectCustomUI.vbs
Option Explicit

' Script pour injecter le CustomUI (Ribbon) dans un fichier XLAM
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Const ForReading = 1, ForWriting = 2

Dim objFSO, objShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' Chemins des fichiers
Dim strXLAMPath, strTempFolder, strCustomUIXML
strXLAMPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\Addin Elyse Energy.xlam"
strTempFolder = objFSO.GetSpecialFolder(2) & "\XLAMTemp_" & CreateGUID()
strCustomUIXML = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\customUI.xml"

' Créer un dossier temporaire
objFSO.CreateFolder(strTempFolder)

' Copier le XLAM comme ZIP dans le dossier temp
objFSO.CopyFile strXLAMPath, strTempFolder & "\temp.zip"

' Extraire le ZIP
objShell.NameSpace(strTempFolder).CopyHere objShell.NameSpace(strTempFolder & "\temp.zip").Items

' Créer/Mettre à jour le dossier customUI
If Not objFSO.FolderExists(strTempFolder & "\customUI") Then
    objFSO.CreateFolder strTempFolder & "\customUI"
End If

' Copier le fichier XML
objFSO.CopyFile strCustomUIXML, strTempFolder & "\customUI\customUI.xml"

' Mettre à jour [Content_Types].xml
UpdateContentTypes strTempFolder & "\[Content_Types].xml"

' Mettre à jour .rels
UpdateRels strTempFolder & "\_rels\.rels"

' Recréer le ZIP
CreateZipFromFolder strTempFolder, strXLAMPath

' Nettoyage
objFSO.DeleteFolder strTempFolder, True

WScript.Echo "CustomUI injecté avec succès !"

' Fonctions utilitaires
Function CreateGUID()
    Dim TypeLib
    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    CreateGUID = Mid(TypeLib.Guid, 2, 36)
End Function

Sub UpdateContentTypes(strPath)
    Dim xmlContent, objStream
    
    If objFSO.FileExists(strPath) Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Open
        objStream.Type = 2 'Text
        objStream.CharSet = "utf-8"
        objStream.LoadFromFile strPath
        xmlContent = objStream.ReadText
        objStream.Close
        
        If InStr(xmlContent, "customUI/customUI.xml") = 0 Then
            xmlContent = Replace(xmlContent, "</Types>", _
                "<Override PartName=""/customUI/customUI.xml"" ContentType=""application/xml""/>" & vbCrLf & "</Types>")
            
            objStream.Open
            objStream.Type = 2 'Text
            objStream.CharSet = "utf-8"
            objStream.WriteText xmlContent
            objStream.SaveToFile strPath, adSaveCreateOverWrite
            objStream.Close
        End If
    End If
End Sub

Sub UpdateRels(strPath)
    Dim xmlContent, objStream
    
    If objFSO.FileExists(strPath) Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Open
        objStream.Type = 2 'Text
        objStream.CharSet = "utf-8"
        objStream.LoadFromFile strPath
        xmlContent = objStream.ReadText
        objStream.Close
        
        If InStr(xmlContent, "customUI/customUI.xml") = 0 Then
            xmlContent = Replace(xmlContent, "</Relationships>", _
                "<Relationship Target=""customUI/customUI.xml"" Type=""http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"" Id=""Rcd1234""/>" & vbCrLf & "</Relationships>")
            
            objStream.Open
            objStream.Type = 2 'Text
            objStream.CharSet = "utf-8"
            objStream.WriteText xmlContent
            objStream.SaveToFile strPath, adSaveCreateOverWrite
            objStream.Close
        End If
    End If
End Sub

Sub CreateZipFromFolder(strFolderPath, strZipPath)
    Dim objZipFile, objFolder, objApp
    
    If objFSO.FileExists(strZipPath) Then
        objFSO.DeleteFile strZipPath
    End If
    
    ' Créer un fichier ZIP vide
    CreateEmptyZip strZipPath
    
    ' Ajouter les fichiers au ZIP
    Set objZipFile = objShell.NameSpace(strZipPath)
    Set objFolder = objShell.NameSpace(strFolderPath)
    
    objZipFile.CopyHere objFolder.Items
    
    ' Attendre que la compression soit terminée
    Do While objZipFile.Items.Count < objFolder.Items.Count
        WScript.Sleep 200
    Loop
End Sub

Sub CreateEmptyZip(strPath)
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    
    ' En-tête ZIP minimal
    Dim zipHeader
    zipHeader = Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
    
    objStream.Open
    objStream.Type = 1 'Binary
    objStream.Write zipHeader
    objStream.SaveToFile strPath, adSaveCreateOverWrite
    objStream.Close
End Sub