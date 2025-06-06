Attribute VB_Name = "Technologies_Manager"
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
    Diagnostics.StartTimer "H2 waters electrolysis"
    DataLoaderManager.ProcessCategory "H2 waters electrolysis", "Erreur lors du traitement des données d'électrolyse"
    Diagnostics.StopTimer "H2 waters electrolysis"
End Sub

Public Sub ProcessCO2Capture(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "CO2 Capture"
    DataLoaderManager.ProcessCategory "CO2 Capture", "Erreur lors du traitement des données CO2 Capture"
    Diagnostics.StopTimer "CO2 Capture"
End Sub

Public Sub ProcessCO2General(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "CO2 general parameters"
    DataLoaderManager.ProcessCategory "CO2 general parameters", "Erreur lors du traitement des données CO2 General Parameters"
    Diagnostics.StopTimer "CO2 general parameters"
End Sub

Public Sub ProcessCompression(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Compression"
    DataLoaderManager.ProcessCategory "Compression", "Erreur lors du traitement des données de compression"
    Diagnostics.StopTimer "Compression"
End Sub

Public Sub ProcessH2General(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "H2 general parameters"
    DataLoaderManager.ProcessCategory "H2 general parameters", "Erreur lors du traitement des données H2 General Parameters"
    Diagnostics.StopTimer "H2 general parameters"
End Sub

Public Sub ProcessMeOHCO2(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "MeOH - CO2-to-Methanol Synthesis"
    DataLoaderManager.ProcessCategory "MeOH - CO2-to-Methanol Synthesis", "Erreur lors du traitement des données MeOH CO2"
    Diagnostics.StopTimer "MeOH - CO2-to-Methanol Synthesis"
End Sub

Public Sub ProcessMeOHBiomass(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "MeOH - Biomass Gasification Synthesis"
    DataLoaderManager.ProcessCategory "MeOH - Biomass Gasification Synthesis", "Erreur lors du traitement des données MeOH Biomass"
    Diagnostics.StopTimer "MeOH - Biomass Gasification Synthesis"
End Sub

Public Sub ProcessSAFBtJ(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "SAF - BtJ/e-BtJ Synthesis"
    DataLoaderManager.ProcessCategory "SAF - BtJ/e-BtJ Synthesis", "Erreur lors du traitement des données SAF BtJ"
    Diagnostics.StopTimer "SAF - BtJ/e-BtJ Synthesis"
End Sub

Public Sub ProcessSAFMtJ(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "SAF - MtJ Synthesis"
    DataLoaderManager.ProcessCategory "SAF - MtJ Synthesis", "Erreur lors du traitement des données SAF MtJ"
    Diagnostics.StopTimer "SAF - MtJ Synthesis"
End Sub

Public Sub ProcessChiller(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Chiller"
    DataLoaderManager.ProcessCategory "Chiller", "Erreur lors du traitement des données Chiller"
    Diagnostics.StopTimer "Chiller"
End Sub

Public Sub ProcessCoolingWater(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Cooling Water Production"
    DataLoaderManager.ProcessCategory "Cooling Water Production", "Erreur lors du traitement des données Cooling Water Production"
    Diagnostics.StopTimer "Cooling Water Production"
End Sub

Public Sub ProcessHeatProd(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Heat Production"
    DataLoaderManager.ProcessCategory "Heat Production", "Erreur lors du traitement des données Heat Production"
    Diagnostics.StopTimer "Heat Production"
End Sub

Public Sub ProcessOtherUtil(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Other utilities"
    DataLoaderManager.ProcessCategory "Other utilities", "Erreur lors du traitement des données Other utilities"
    Diagnostics.StopTimer "Other utilities"
End Sub

Public Sub ProcessPowerLoss(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Power losses"
    DataLoaderManager.ProcessCategory "Power losses", "Erreur lors du traitement des données Power losses"
    Diagnostics.StopTimer "Power losses"
End Sub

Public Sub ProcessWastewater(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "WasteWater Treatment"
    DataLoaderManager.ProcessCategory "WasteWater Treatment", "Erreur lors du traitement des données WasteWater Treatment"
    Diagnostics.StopTimer "WasteWater Treatment"
End Sub

Public Sub ProcessWaterTreat(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Water Treatment"
    DataLoaderManager.ProcessCategory "Water Treatment", "Erreur lors du traitement des données Water Treatment"
    Diagnostics.StopTimer "Water Treatment"
End Sub

Public Sub ProcessMetriquesBase(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Métriques de base"
    DataLoaderManager.ProcessCategory "Métriques de base", "Erreur lors du traitement des métriques de base"
    Diagnostics.StopTimer "Métriques de base"
End Sub

Public Sub ProcessMetriquesExpert(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Métriques expert"
    DataLoaderManager.ProcessCategory "Métriques expert", "Erreur lors du traitement des métriques expert"
    Diagnostics.StopTimer "Métriques expert"
End Sub

Public Sub ProcessTimingsReference(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Timings de référence"
    DataLoaderManager.ProcessCategory "Timings de référence", "Erreur lors du traitement des timings de référence"
    Diagnostics.StopTimer "Timings de référence"
End Sub

Public Sub ProcessMetriquesRED(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Métriques RED III"
    DataLoaderManager.ProcessCategory "Métriques RED III", "Erreur lors du traitement des métriques RED III"
    Diagnostics.StopTimer "Métriques RED III"
End Sub

Public Sub ProcessEmissions(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Emissions"
    DataLoaderManager.ProcessCategory "Emissions", "Erreur lors du traitement des émissions"
    Diagnostics.StopTimer "Emissions"
End Sub

Public Sub ProcessInfraLog(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Infra et logistique"
    DataLoaderManager.ProcessCategory "Infra et logistique", "Erreur lors du traitement des données d'infrastructure et logistique"
    Diagnostics.StopTimer "Infra et logistique"
End Sub

Public Sub ProcessBudgetCorpo(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Budget Corpo"
    DataLoaderManager.ProcessCategory "Budget Corpo", "Erreur lors du traitement du budget corpo"
    Diagnostics.StopTimer "Budget Corpo"
End Sub

Public Sub ProcessDetailsBudgets(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Détails Budgets"
    DataLoaderManager.ProcessCategory "Détails Budgets", "Erreur lors du traitement des détails budgets"
    Diagnostics.StopTimer "Détails Budgets"
End Sub

Public Sub ProcessDIB(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "DIB"
    DataLoaderManager.ProcessCategory "DIB", "Erreur lors du traitement des DIB"
    Diagnostics.StopTimer "DIB"
End Sub

Public Sub ProcessDemandesAchat(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Demandes d'achat"
    DataLoaderManager.ProcessCategory "Demandes d'achat", "Erreur lors du traitement des demandes d'achat"
    Diagnostics.StopTimer "Demandes d'achat"
End Sub

Public Sub ProcessReceptions(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Réceptions"
    DataLoaderManager.ProcessCategory "Réceptions", "Erreur lors du traitement des réceptions"
    Diagnostics.StopTimer "Réceptions"
End Sub

Public Sub ProcessScenarios(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Scénarios techniques"
    DataLoaderManager.ProcessCategory "Scénarios techniques", "Erreur lors du traitement des scénarios techniques"
    Diagnostics.StopTimer "Scénarios techniques"
End Sub

Public Sub ProcessPlanningPhases(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Plannings de phase"
    DataLoaderManager.ProcessCategory "Plannings de phase", "Erreur lors du traitement des plannings de phase"
    Diagnostics.StopTimer "Plannings de phase"
End Sub

Public Sub ProcessPlanningSous(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Plannings de sous phases"
    DataLoaderManager.ProcessCategory "Plannings de sous phases", "Erreur lors du traitement des plannings de sous-phase"
    Diagnostics.StopTimer "Plannings de sous phases"
End Sub

Public Sub ProcessBudgetProjet(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "Budget Projet"
    DataLoaderManager.ProcessCategory "Budget Projet", "Erreur lors du traitement du budget projet"
    Diagnostics.StopTimer "Budget Projet"
End Sub

Public Sub ProcessDevex(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "DEVEX"
    DataLoaderManager.ProcessCategory "DEVEX", "Erreur lors du traitement des DEVEX"
    Diagnostics.StopTimer "DEVEX"
End Sub

Public Sub ProcessCapex(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "CAPEX"
    DataLoaderManager.ProcessCategory "CAPEX", "Erreur lors du traitement des CAPEX"
    Diagnostics.StopTimer "CAPEX"
End Sub

Public Sub ProcessCapexEPC(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Diagnostics.StartTimer "CAPEX EPC"
    DataLoaderManager.ProcessCategory "CAPEX EPC", "Erreur lors du traitement des CAPEX EPC"
    Diagnostics.StopTimer "CAPEX EPC"
End Sub

