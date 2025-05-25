' Module: H2_Waters_Electrolysis_Manager
' Gère le traitement des données d'électrolyse de l'eau
Option Explicit
Private Const MODULE_NAME As String = "H2_Waters_Electrolysis_Manager"


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
    Const PROC_NAME As String = "ProcessH2Electrolysis"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données d'électrolyse de l'eau.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "H2 waters electrolysis", "Erreur lors du traitement des données d'électrolyse"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données d'électrolyse de l'eau terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCO2Capture(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCO2Capture"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données CO2 Capture.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "CO2 Capture", "Erreur lors du traitement des données CO2 Capture"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données CO2 Capture terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCO2General(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCO2General"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données CO2 General Parameters.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "CO2 general parameters", "Erreur lors du traitement des données CO2 General Parameters"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données CO2 General Parameters terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCompression(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCompression"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données de compression.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Compression", "Erreur lors du traitement des données de compression"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données de compression terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessH2General(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessH2General"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données H2 General Parameters.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "H2 general parameters", "Erreur lors du traitement des données H2 General Parameters"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données H2 General Parameters terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMeOHCO2(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessMeOHCO2"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données MeOH CO2.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "MeOH - CO2-to-Methanol Synthesis", "Erreur lors du traitement des données MeOH CO2"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données MeOH CO2 terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMeOHBiomass(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessMeOHBiomass"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données MeOH Biomass.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "MeOH - Biomass Gasification Synthesis", "Erreur lors du traitement des données MeOH Biomass"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données MeOH Biomass terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessSAFBtJ(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessSAFBtJ"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données SAF BtJ.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "SAF - BtJ/e-BtJ Synthesis", "Erreur lors du traitement des données SAF BtJ"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données SAF BtJ terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessSAFMtJ(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessSAFMtJ"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données SAF MtJ.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "SAF - MtJ Synthesis", "Erreur lors du traitement des données SAF MtJ"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données SAF MtJ terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessChiller(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessChiller"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données Chiller.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Chiller", "Erreur lors du traitement des données Chiller"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données Chiller terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCoolingWater(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCoolingWater"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données Cooling Water Production.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Cooling Water Production", "Erreur lors du traitement des données Cooling Water Production"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données Cooling Water Production terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessHeatProd(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessHeatProd"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données Heat Production.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Heat Production", "Erreur lors du traitement des données Heat Production"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données Heat Production terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessOtherUtil(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessOtherUtil"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données Other utilities.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Other utilities", "Erreur lors du traitement des données Other utilities"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données Other utilities terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessPowerLoss(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessPowerLoss"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données Power losses.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Power losses", "Erreur lors du traitement des données Power losses"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données Power losses terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessWastewater(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessWastewater"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données WasteWater Treatment.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "WasteWater Treatment", "Erreur lors du traitement des données WasteWater Treatment"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données WasteWater Treatment terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessWaterTreat(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessWaterTreat"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données Water Treatment.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Water Treatment", "Erreur lors du traitement des données Water Treatment"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données Water Treatment terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMetriquesBase(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessMetriquesBase"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des métriques de base.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Métriques de base", "Erreur lors du traitement des métriques de base"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des métriques de base terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMetriquesExpert(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessMetriquesExpert"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des métriques expert.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Métriques expert", "Erreur lors du traitement des métriques expert"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des métriques expert terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessTimingsReference(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessTimingsReference"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des timings de référence.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Timings de référence", "Erreur lors du traitement des timings de référence"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des timings de référence terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMetriquesRED(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessMetriquesRED"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des métriques RED III.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Métriques RED III", "Erreur lors du traitement des métriques RED III"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des métriques RED III terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessEmissions(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessEmissions"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des émissions.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Emissions", "Erreur lors du traitement des émissions"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des émissions terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessInfraLog(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessInfraLog"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des données d'infrastructure et logistique.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Infra et logistique", "Erreur lors du traitement des données d'infrastructure et logistique"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des données d'infrastructure et logistique terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessBudgetCorpo(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessBudgetCorpo"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement du budget corpo.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Budget Corpo", "Erreur lors du traitement du budget corpo"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement du budget corpo terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDetailsBudgets(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessDetailsBudgets"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des détails budgets.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Détails Budgets", "Erreur lors du traitement des détails budgets"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des détails budgets terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDIB(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessDIB"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des DIB.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "DIB", "Erreur lors du traitement des DIB"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des DIB terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDemandesAchat(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessDemandesAchat"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des demandes d'achat.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Demandes d'achat", "Erreur lors du traitement des demandes d'achat"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des demandes d'achat terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessReceptions(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessReceptions"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des réceptions.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Réceptions", "Erreur lors du traitement des réceptions"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des réceptions terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessScenarios(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessScenarios"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des scénarios techniques.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Scénarios techniques", "Erreur lors du traitement des scénarios techniques"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des scénarios techniques terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessPlanningPhases(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessPlanningPhases"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des plannings de phases.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Plannings de phases", "Erreur lors du traitement des plannings de phases"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des plannings de phases terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessPlanningSous(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessPlanningSous"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement des plannings de sous phases.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Plannings de sous phases", "Erreur lors du traitement des plannings de sous phases"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement des plannings de sous phases terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessBudgetProjet(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessBudgetProjet"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement du budget projet.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Budget Projet", "Erreur lors du traitement du budget projet"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement du budget projet terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDevex(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessDevex"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement du Devex.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Devex", "Erreur lors du traitement du Devex"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement du Devex terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCapex(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCapex"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement du Capex.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Capex", "Erreur lors du traitement du Capex"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement du Capex terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCapexEPC(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCapexEPC"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Traitement du Capex EPC.", PROC_NAME, MODULE_NAME
    
    DataLoaderManager.ProcessCategory "Capex EPC", "Erreur lors du traitement du Capex EPC"
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Traitement du Capex EPC terminé.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub
