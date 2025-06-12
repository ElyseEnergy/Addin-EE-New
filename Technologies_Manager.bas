Attribute VB_Name = "Technologies_Manager"
Option Explicit

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

Public Sub ProcessDevexMain()
    ProcessDevex Nothing, Nothing
End Sub

Public Sub ProcessCapexMain()
    ProcessCapex Nothing, Nothing
End Sub

Public Sub ProcessH2Electrolysis(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessH2Electrolysis"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler
    
    Diagnostics.StartTimer "H2 waters electrolysis"
    DataLoaderManager.ProcessCategory "H2 waters electrolysis", "Erreur lors du traitement des données d'électrolyse"
    Diagnostics.StopTimer "H2 waters electrolysis"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCO2Capture(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessCO2Capture"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "CO2 Capture"
    DataLoaderManager.ProcessCategory "CO2 Capture", "Erreur lors du traitement des données CO2 Capture"
    Diagnostics.StopTimer "CO2 Capture"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCO2General(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessCO2General"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "CO2 general parameters"
    DataLoaderManager.ProcessCategory "CO2 general parameters", "Erreur lors du traitement des données CO2 General Parameters"
    Diagnostics.StopTimer "CO2 general parameters"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCompression(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessCompression"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler
    
    Diagnostics.StartTimer "Compression"
    DataLoaderManager.ProcessCategory "Compression", "Erreur lors du traitement des données de compression"
    Diagnostics.StopTimer "Compression"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessH2General(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessH2General"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "H2 general parameters"
    DataLoaderManager.ProcessCategory "H2 general parameters", "Erreur lors du traitement des données H2 General Parameters"
    Diagnostics.StopTimer "H2 general parameters"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMeOHCO2(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessMeOHCO2"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "MeOH - CO2-to-Methanol Synthesis"
    DataLoaderManager.ProcessCategory "MeOH - CO2-to-Methanol Synthesis", "Erreur lors du traitement des données MeOH CO2"
    Diagnostics.StopTimer "MeOH - CO2-to-Methanol Synthesis"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMeOHBiomass(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessMeOHBiomass"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "MeOH - Biomass Gasification Synthesis"
    DataLoaderManager.ProcessCategory "MeOH - Biomass Gasification Synthesis", "Erreur lors du traitement des données MeOH Biomass"
    Diagnostics.StopTimer "MeOH - Biomass Gasification Synthesis"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessSAFBtJ(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessSAFBtJ"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "SAF - BtJ/e-BtJ Synthesis"
    DataLoaderManager.ProcessCategory "SAF - BtJ/e-BtJ Synthesis", "Erreur lors du traitement des données SAF BtJ"
    Diagnostics.StopTimer "SAF - BtJ/e-BtJ Synthesis"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessSAFMtJ(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessSAFMtJ"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "SAF - MtJ Synthesis"
    DataLoaderManager.ProcessCategory "SAF - MtJ Synthesis", "Erreur lors du traitement des données SAF MtJ"
    Diagnostics.StopTimer "SAF - MtJ Synthesis"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessChiller(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessChiller"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Chiller"
    DataLoaderManager.ProcessCategory "Chiller", "Erreur lors du traitement des données Chiller"
    Diagnostics.StopTimer "Chiller"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCoolingWater(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessCoolingWater"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Cooling Water Production"
    DataLoaderManager.ProcessCategory "Cooling Water Production", "Erreur lors du traitement des données Cooling Water Production"
    Diagnostics.StopTimer "Cooling Water Production"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessHeatProd(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessHeatProd"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Heat Production"
    DataLoaderManager.ProcessCategory "Heat Production", "Erreur lors du traitement des données Heat Production"
    Diagnostics.StopTimer "Heat Production"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessOtherUtil(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessOtherUtil"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Other utilities"
    DataLoaderManager.ProcessCategory "Other utilities", "Erreur lors du traitement des données Other utilities"
    Diagnostics.StopTimer "Other utilities"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessPowerLoss(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessPowerLoss"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Power losses"
    DataLoaderManager.ProcessCategory "Power losses", "Erreur lors du traitement des données Power losses"
    Diagnostics.StopTimer "Power losses"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessWastewater(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessWastewater"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "WasteWater Treatment"
    DataLoaderManager.ProcessCategory "WasteWater Treatment", "Erreur lors du traitement des données WasteWater Treatment"
    Diagnostics.StopTimer "WasteWater Treatment"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessWaterTreat(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessWaterTreat"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Water Treatment"
    DataLoaderManager.ProcessCategory "Water Treatment", "Erreur lors du traitement des données Water Treatment"
    Diagnostics.StopTimer "Water Treatment"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMetriquesBase(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessMetriquesBase"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Métriques de base"
    DataLoaderManager.ProcessCategory "Métriques de base", "Erreur lors du traitement des métriques de base"
    Diagnostics.StopTimer "Métriques de base"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMetriquesExpert(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessMetriquesExpert"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Métriques expert"
    DataLoaderManager.ProcessCategory "Métriques expert", "Erreur lors du traitement des métriques expert"
    Diagnostics.StopTimer "Métriques expert"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessTimingsReference(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessTimingsReference"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Timings de référence"
    DataLoaderManager.ProcessCategory "Timings de référence", "Erreur lors du traitement des timings de référence"
    Diagnostics.StopTimer "Timings de référence"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMetriquesRED(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessMetriquesRED"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Métriques RED"
    DataLoaderManager.ProcessCategory "Métriques RED", "Erreur lors du traitement des métriques RED"
    Diagnostics.StopTimer "Métriques RED"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessEmissions(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessEmissions"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Emissions"
    DataLoaderManager.ProcessCategory "Emissions", "Erreur lors du traitement des émissions"
    Diagnostics.StopTimer "Emissions"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessInfraLog(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessInfraLog"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Infra & Log"
    DataLoaderManager.ProcessCategory "Infra & Log", "Erreur lors du traitement de Infra & Log"
    Diagnostics.StopTimer "Infra & Log"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessBudgetCorpo(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessBudgetCorpo"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Budget Corpo"
    DataLoaderManager.ProcessCategory "Budget Corpo", "Erreur lors du traitement du budget corpo"
    Diagnostics.StopTimer "Budget Corpo"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDIB(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessDIB"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "DIB"
    DataLoaderManager.ProcessCategory "DIB", "Erreur lors du traitement du DIB"
    Diagnostics.StopTimer "DIB"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDemandesAchat(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessDemandesAchat"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Demandes Achat"
    DataLoaderManager.ProcessCategory "Demandes Achat", "Erreur lors du traitement des demandes d'achat"
    Diagnostics.StopTimer "Demandes Achat"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessReceptions(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessReceptions"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Réceptions"
    DataLoaderManager.ProcessCategory "Réceptions", "Erreur lors du traitement des réceptions"
    Diagnostics.StopTimer "Réceptions"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessScenarios(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessScenarios"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Scénarios"
    DataLoaderManager.ProcessCategory "Scénarios", "Erreur lors du traitement des scénarios"
    Diagnostics.StopTimer "Scénarios"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessPlanningPhases(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessPlanningPhases"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Planning Phases"
    DataLoaderManager.ProcessCategory "Planning Phases", "Erreur lors du traitement du planning des phases"
    Diagnostics.StopTimer "Planning Phases"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessPlanningSous(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessPlanningSous"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Planning Sous-Phases"
    DataLoaderManager.ProcessCategory "Planning Sous-Phases", "Erreur lors du traitement du planning des sous-phases"
    Diagnostics.StopTimer "Planning Sous-Phases"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessBudgetProjet(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessBudgetProjet"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler
    
    DataLoaderManager.ProcessCategory "Budget Projet", "Erreur lors du traitement du budget projet"
    
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDevex(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessDevex"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Devex"
    DataLoaderManager.ProcessCategory "Devex", "Erreur lors du traitement du Devex"
    Diagnostics.StopTimer "Devex"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCapex(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessCapex"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    Diagnostics.StartTimer "Capex"
    DataLoaderManager.ProcessCategory "Capex", "Erreur lors du traitement du Capex"
    Diagnostics.StopTimer "Capex"
    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCapexEPC(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessCapexEPC"
    Const MODULE_NAME As String = "Technologies_Manager"
    On Error GoTo ErrorHandler

    DataLoaderManager.ProcessCategory "CAPEX EPC", "Erreur lors du traitement du CAPEX EPC"

    Exit Sub
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

