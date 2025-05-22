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

' Nouvelles catégories
Public Sub ProcessMetriquesBaseMain()
    ProcessMetriquesBase Nothing, Nothing
End Sub

Public Sub ProcessMetriquesExpertMain()
    ProcessMetriquesExpert Nothing, Nothing
End Sub

Public Sub ProcessTimingsReferenceMain()
    ProcessTimingsReference Nothing, Nothing
End Sub

Public Sub ProcessMetriquesREDMain()
    ProcessMetriquesRED Nothing, Nothing
End Sub

Public Sub ProcessEmissionsMain()
    ProcessEmissions Nothing, Nothing
End Sub

Public Sub ProcessInfraLogMain()
    ProcessInfraLog Nothing, Nothing
End Sub

Public Sub ProcessBudgetCorpoMain()
    ProcessBudgetCorpo Nothing, Nothing
End Sub

Public Sub ProcessDetailsBudgetsMain()
    ProcessDetailsBudgets Nothing, Nothing
End Sub

Public Sub ProcessDIBMain()
    ProcessDIB Nothing, Nothing
End Sub

Public Sub ProcessDemandesAchatMain()
    ProcessDemandesAchat Nothing, Nothing
End Sub

Public Sub ProcessReceptionsMain()
    ProcessReceptions Nothing, Nothing
End Sub

Public Sub ProcessScenariosMain()
    ProcessScenarios Nothing, Nothing
End Sub

Public Sub ProcessPlanningPhasesMain()
    ProcessPlanningPhases Nothing, Nothing
End Sub

Public Sub ProcessPlanningSousMain()
    ProcessPlanningSous Nothing, Nothing
End Sub

Public Sub ProcessBudgetProjetMain()
    ProcessBudgetProjet Nothing, Nothing
End Sub

Public Sub ProcessDevexMain()
    ProcessDevex Nothing, Nothing
End Sub

Public Sub ProcessCapexMain()
    ProcessCapex Nothing, Nothing
End Sub

Public Sub ProcessCapexEPCMain()
    ProcessCapexEPC Nothing, Nothing
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

Public Sub ProcessMetriquesBase(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Métriques de base", "Erreur lors du traitement des métriques de base"
End Sub

Public Sub ProcessMetriquesExpert(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Métriques expert", "Erreur lors du traitement des métriques expert"
End Sub

Public Sub ProcessTimingsReference(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Timings de référence", "Erreur lors du traitement des timings de référence"
End Sub

Public Sub ProcessMetriquesRED(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Métriques RED III", "Erreur lors du traitement des métriques RED III"
End Sub

Public Sub ProcessEmissions(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Emissions", "Erreur lors du traitement des émissions"
End Sub

Public Sub ProcessInfraLog(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Infra et logistique", "Erreur lors du traitement des données d'infrastructure et logistique"
End Sub

Public Sub ProcessBudgetCorpo(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Budget Corpo", "Erreur lors du traitement du budget corpo"
End Sub

Public Sub ProcessDetailsBudgets(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Détails Budgets", "Erreur lors du traitement des détails budgets"
End Sub

Public Sub ProcessDIB(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "DIB", "Erreur lors du traitement des DIB"
End Sub

Public Sub ProcessDemandesAchat(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Demandes d'achat", "Erreur lors du traitement des demandes d'achat"
End Sub

Public Sub ProcessReceptions(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Réceptions", "Erreur lors du traitement des réceptions"
End Sub

Public Sub ProcessScenarios(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Scénarios techniques", "Erreur lors du traitement des scénarios techniques"
End Sub

Public Sub ProcessPlanningPhases(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Plannings de phases", "Erreur lors du traitement des plannings de phases"
End Sub

Public Sub ProcessPlanningSous(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Plannings de sous phases", "Erreur lors du traitement des plannings de sous phases"
End Sub

Public Sub ProcessBudgetProjet(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Budget Projet", "Erreur lors du traitement du budget projet"
End Sub

Public Sub ProcessDevex(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Devex", "Erreur lors du traitement du Devex"
End Sub

Public Sub ProcessCapex(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Capex", "Erreur lors du traitement du Capex"
End Sub

Public Sub ProcessCapexEPC(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    DataLoaderManager.ProcessCategory "Capex EPC", "Erreur lors du traitement du Capex EPC"
End Sub
