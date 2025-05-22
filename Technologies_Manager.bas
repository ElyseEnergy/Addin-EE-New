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

Public Sub ProcessChillerMain()
    ProcessChiller Nothing, Nothing
End Sub

Public Sub ProcessCoolingWaterMain()
    ProcessCoolingWater Nothing, Nothing
End Sub

Public Sub ProcessHeatProdMain()
    ProcessHeatProd Nothing, Nothing
End Sub

Public Sub ProcessOtherUtilMain()
    ProcessOtherUtil Nothing, Nothing
End Sub

Public Sub ProcessPowerLossMain()
    ProcessPowerLoss Nothing, Nothing
End Sub

Public Sub ProcessWastewaterMain()
    ProcessWastewater Nothing, Nothing
End Sub

Public Sub ProcessWaterTreatMain()
    ProcessWaterTreat Nothing, Nothing
End Sub

Public Sub ProcessH2Electrolysis(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "H2 waters electrolysis", "Erreur lors du traitement des données d'électrolyse"
End Sub

Public Sub ProcessCO2Capture(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "CO2 Capture", "Erreur lors du traitement des données CO2 Capture"
End Sub

Public Sub ProcessCO2General(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "CO2 general parameters", "Erreur lors du traitement des données CO2 General Parameters"
End Sub

Public Sub ProcessCompression(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Compression", "Erreur lors du traitement des données de compression"
End Sub

Public Sub ProcessH2General(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "H2 general parameters", "Erreur lors du traitement des données H2 General Parameters"
End Sub

Public Sub ProcessMeOHCO2(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "MeOH - CO2-to-Methanol Synthesis", "Erreur lors du traitement des données MeOH CO2"
End Sub

Public Sub ProcessMeOHBiomass(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "MeOH - Biomass Gasification Synthesis", "Erreur lors du traitement des données MeOH Biomass"
End Sub

Public Sub ProcessSAFBtJ(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "SAF - BtJ/e-BtJ Synthesis", "Erreur lors du traitement des données SAF BtJ"
End Sub

Public Sub ProcessSAFMtJ(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "SAF - MtJ Synthesis", "Erreur lors du traitement des données SAF MtJ"
End Sub

Public Sub ProcessChiller(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Chiller", "Erreur lors du traitement des données Chiller"
End Sub

Public Sub ProcessCoolingWater(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Cooling Water Production", "Erreur lors du traitement des données Cooling Water Production"
End Sub

Public Sub ProcessHeatProd(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Heat Production", "Erreur lors du traitement des données Heat Production"
End Sub

Public Sub ProcessOtherUtil(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Other utilities", "Erreur lors du traitement des données Other utilities"
End Sub

Public Sub ProcessPowerLoss(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Power losses", "Erreur lors du traitement des données Power losses"
End Sub

Public Sub ProcessWastewater(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "WasteWater Treatment", "Erreur lors du traitement des données WasteWater Treatment"
End Sub

Public Sub ProcessWaterTreat(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Water Treatment", "Erreur lors du traitement des données Water Treatment"
End Sub
