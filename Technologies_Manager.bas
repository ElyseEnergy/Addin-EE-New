' Module: H2_Waters_Electrolysis_Manager
' Gère le traitement des données d'électrolyse de l'eau
Option Explicit


' Wrappers sans callback pour permettre l'appel direct
Public Sub ProcessH2ElectrolysisMain()
    ProcessH2Electrolysis Nothing, Nothing
End Sub

Public Sub ProcessCO2CaptureMain()
    ProcessCO2Capture Nothing, Nothing
End Sub

Public Sub ProcessCO2GeneralMain()
    ProcessCO2General Nothing, Nothing
End Sub

Public Sub ProcessCompressionMain()
    ProcessCompression Nothing, Nothing
End Sub

Public Sub ProcessH2GeneralMain()
    ProcessH2General Nothing, Nothing
End Sub

Public Sub ProcessMeOHCO2Main()
    ProcessMeOHCO2 Nothing, Nothing
End Sub

Public Sub ProcessMeOHBiomassMain()
    ProcessMeOHBiomass Nothing, Nothing
End Sub

Public Sub ProcessSAFBtJMain()
    ProcessSAFBtJ Nothing, Nothing
End Sub

Public Sub ProcessSAFMtJMain()
    ProcessSAFMtJ Nothing, Nothing
End Sub

' Fonction générique pour traiter une catégorie
Private Function ProcessCategory(categoryName As String, errorMessage As String) As Boolean
    If CategoriesCount = 0 Then InitCategories
    
    Dim loadInfo As DataLoadInfo
    loadInfo.Category = GetCategoryByName(categoryName)
    If loadInfo.Category.DisplayName = "" Then
        MsgBox "Catégorie '" & categoryName & "' non trouvée", vbExclamation
        ProcessCategory = False
        Exit Function
    End If
    
    loadInfo.PreviewRows = 3
    
    If Not DataLoaderManager.ProcessDataLoad(loadInfo) Then
        MsgBox errorMessage, vbExclamation
        ProcessCategory = False
        Exit Function
    End If
    
    ProcessCategory = True
End Function

Public Sub ProcessH2Electrolysis(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ProcessCategory "H2 waters electrolysis", "Erreur lors du traitement des données d'électrolyse"
End Sub

Public Sub ProcessCO2Capture(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ProcessCategory "CO2 Capture", "Erreur lors du traitement des données CO2 Capture"
End Sub

Public Sub ProcessCO2General(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ProcessCategory "CO2 general parameters", "Erreur lors du traitement des données CO2 General Parameters"
End Sub

Public Sub ProcessCompression(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ProcessCategory "Compression", "Erreur lors du traitement des données de compression"
End Sub

Public Sub ProcessH2General(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ProcessCategory "H2 general parameters", "Erreur lors du traitement des données H2 General Parameters"
End Sub

Public Sub ProcessMeOHCO2(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ProcessCategory "MeOH - CO2-to-Methanol Synthesis", "Erreur lors du traitement des données MeOH CO2"
End Sub

Public Sub ProcessMeOHBiomass(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ProcessCategory "MeOH - Biomass Gasification Synthesis", "Erreur lors du traitement des données MeOH Biomass"
End Sub

Public Sub ProcessSAFBtJ(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ProcessCategory "SAF - BtJ/e-BtJ Synthesis", "Erreur lors du traitement des données SAF BtJ"
End Sub

Public Sub ProcessSAFMtJ(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ProcessCategory "SAF - MtJ Synthesis", "Erreur lors du traitement des données SAF MtJ"
End Sub
