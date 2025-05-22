Option Explicit

' Convertir le fichier XLSM en XLAM
' Script à enregistrer avec l'extension .vbs

Dim objFSO, objExcel, objWorkbook
Dim strSourcePath, strTargetPath

' Créer les objets nécessaires
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objExcel = CreateObject("Excel.Application")

' Configurer Excel
objExcel.Visible = False 
objExcel.DisplayAlerts = False

' Chemins des fichiers
strSourcePath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\Addin Elyse Energy.xlsm"
strTargetPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\Addin Elyse Energy.xlam"

' Ouvrir le fichier source
Set objWorkbook = objExcel.Workbooks.Open(strSourcePath)

' Enregistrer en tant que XLAM
objWorkbook.SaveAs strTargetPath, 55 ' 55 = xlOpenXMLAddIn

' Nettoyage
objWorkbook.Close
objExcel.Quit

Set objWorkbook = Nothing
Set objExcel = Nothing
Set objFSO = Nothing

WScript.Echo "Conversion terminée avec succès !"