Option Explicit
Private Const MODULE_NAME As String = "UI_03_RibbonActions"

Private Function ProcessCategory(ByVal categoryName As String, ByVal errorMessage As String) As DataLoadResult
    On Error GoTo ErrorHandler
    
    Dim category As CategoryInfo
    Set category = CategoryDefinitions_System.GetCategoryByName(categoryName)
    If category Is Nothing Then
        ElyseMessageBox_System.ShowErrorMessage "Erreur", _
            "Catégorie '" & categoryName & "' non trouvée"
        ProcessCategory = DataLoadResult.Error
        Exit Function
    End If
    
    Dim loadInfo As DataLoadInfo
    Set loadInfo = New DataLoadInfo
    Set loadInfo.Category = category
    
    ProcessCategory = DataLoadManager.ProcessDataLoad(loadInfo)
    Exit Function
    
ErrorHandler:
    If Len(errorMessage) > 0 Then
        ElyseMessageBox_System.ShowErrorMessage "Erreur", errorMessage
    Else
        ElyseMessageBox_System.ShowErrorMessage "Erreur", _
            "Une erreur s'est produite : " & Err.Description
    End If
    ProcessCategory = DataLoadResult.Error
End Function

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

    LogInfo PROC_NAME & "_Start", "Traitement des données d'électrolyse de l'eau.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("H2 waters electrolysis", "Erreur lors du traitement des données d'électrolyse")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données d'électrolyse de l'eau terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données d'électrolyse de l'eau annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données d'électrolyse de l'eau.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCO2Capture(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCO2Capture"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données CO2 Capture.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("CO2 Capture", "Erreur lors du traitement des données CO2 Capture")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données CO2 Capture terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données CO2 Capture annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données CO2 Capture.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCO2General(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCO2General"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données CO2 General Parameters.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("CO2 general parameters", "Erreur lors du traitement des données CO2 General Parameters")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données CO2 General Parameters terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données CO2 General Parameters annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données CO2 General Parameters.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCompression(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCompression"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données de compression.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Compression", "Erreur lors du traitement des données de compression")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données de compression terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données de compression annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données de compression.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessH2General(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessH2General"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données H2 General Parameters.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("H2 general parameters", "Erreur lors du traitement des données H2 General Parameters")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données H2 General Parameters terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données H2 General Parameters annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données H2 General Parameters.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMeOHCO2(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessMeOHCO2"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données MeOH CO2.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("MeOH - CO2-to-Methanol Synthesis", "Erreur lors du traitement des données MeOH CO2")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données MeOH CO2 terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données MeOH CO2 annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données MeOH CO2.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMeOHBiomass(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessMeOHBiomass"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données MeOH Biomass.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("MeOH - Biomass Gasification Synthesis", "Erreur lors du traitement des données MeOH Biomass")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données MeOH Biomass terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données MeOH Biomass annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données MeOH Biomass.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessSAFBtJ(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessSAFBtJ"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données SAF BtJ.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("SAF - BtJ/e-BtJ Synthesis", "Erreur lors du traitement des données SAF BtJ")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données SAF BtJ terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données SAF BtJ annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données SAF BtJ.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessSAFMtJ(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessSAFMtJ"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données SAF MtJ.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("SAF - MtJ Synthesis", "Erreur lors du traitement des données SAF MtJ")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données SAF MtJ terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données SAF MtJ annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données SAF MtJ.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessChiller(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessChiller"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données Chiller.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Chiller", "Erreur lors du traitement des données Chiller")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données Chiller terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données Chiller annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données Chiller.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCoolingWater(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCoolingWater"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données Cooling Water Production.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Cooling Water Production", "Erreur lors du traitement des données Cooling Water Production")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données Cooling Water Production terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données Cooling Water Production annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données Cooling Water Production.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessHeatProd(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessHeatProd"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données Heat Production.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Heat Production", "Erreur lors du traitement des données Heat Production")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données Heat Production terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données Heat Production annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données Heat Production.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessOtherUtil(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessOtherUtil"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données Other utilities.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Other utilities", "Erreur lors du traitement des données Other utilities")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données Other utilities terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données Other utilities annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données Other utilities.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessPowerLoss(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessPowerLoss"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données Power losses.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Power losses", "Erreur lors du traitement des données Power losses")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données Power losses terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données Power losses annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données Power losses.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessWastewater(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessWastewater"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données WasteWater Treatment.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("WasteWater Treatment", "Erreur lors du traitement des données WasteWater Treatment")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données WasteWater Treatment terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données WasteWater Treatment annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données WasteWater Treatment.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessWaterTreat(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessWaterTreat"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données Water Treatment.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Water Treatment", "Erreur lors du traitement des données Water Treatment")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données Water Treatment terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données Water Treatment annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données Water Treatment.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMetriquesBase(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessMetriquesBase"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des métriques de base.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Métriques de base", "Erreur lors du traitement des métriques de base")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des métriques de base terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des métriques de base annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des métriques de base.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMetriquesExpert(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessMetriquesExpert"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des métriques expert.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Métriques expert", "Erreur lors du traitement des métriques expert")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des métriques expert terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des métriques expert annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des métriques expert.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessTimingsReference(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessTimingsReference"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des timings de référence.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Timings de référence", "Erreur lors du traitement des timings de référence")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des timings de référence terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des timings de référence annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des timings de référence.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessMetriquesRED(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessMetriquesRED"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des métriques RED III.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Métriques RED III", "Erreur lors du traitement des métriques RED III")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des métriques RED III terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des métriques RED III annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des métriques RED III.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessEmissions(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessEmissions"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des émissions.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Emissions", "Erreur lors du traitement des émissions")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des émissions terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des émissions annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des émissions.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessInfraLog(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessInfraLog"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des données d'infrastructure et logistique.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Infra et logistique", "Erreur lors du traitement des données d'infrastructure et logistique")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des données d'infrastructure et logistique terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des données d'infrastructure et logistique annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des données d'infrastructure et logistique.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessBudgetCorpo(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessBudgetCorpo"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement du budget corpo.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Budget Corpo", "Erreur lors du traitement du budget corpo")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement du budget corpo terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement du budget corpo annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement du budget corpo.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDetailsBudgets(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessDetailsBudgets"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des détails budgets.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Détails Budgets", "Erreur lors du traitement des détails budgets")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des détails budgets terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des détails budgets annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des détails budgets.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDIB(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessDIB"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des DIB.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("DIB", "Erreur lors du traitement des DIB")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des DIB terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des DIB annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des DIB.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDemandesAchat(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessDemandesAchat"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des demandes d'achat.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Demandes d'achat", "Erreur lors du traitement des demandes d'achat")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des demandes d'achat terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des demandes d'achat annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des demandes d'achat.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessReceptions(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessReceptions"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des réceptions.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Réceptions", "Erreur lors du traitement des réceptions")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des réceptions terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des réceptions annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des réceptions.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessScenarios(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessScenarios"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des scénarios techniques.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Scénarios techniques", "Erreur lors du traitement des scénarios techniques")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des scénarios techniques terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des scénarios techniques annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des scénarios techniques.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessPlanningPhases(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessPlanningPhases"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des plannings de phases.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Plannings de phases", "Erreur lors du traitement des plannings de phases")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des plannings de phases terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des plannings de phases annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des plannings de phases.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessPlanningSous(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessPlanningSous"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement des plannings de sous phases.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Plannings de sous phases", "Erreur lors du traitement des plannings de sous phases")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement des plannings de sous phases terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement des plannings de sous phases annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement des plannings de sous phases.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessBudgetProjet(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessBudgetProjet"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement du budget projet.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Budget Projet", "Erreur lors du traitement du budget projet")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement du budget projet terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement du budget projet annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement du budget projet.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDevex(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessDevex"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement du Devex.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Devex", "Erreur lors du traitement du Devex")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement du Devex terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement du Devex annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement du Devex.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCapex(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCapex"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement du Capex.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Capex", "Erreur lors du traitement du Capex")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement du Capex terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement du Capex annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement du Capex.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCapexEPC(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCapexEPC"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Traitement du Capex EPC.", PROC_NAME, MODULE_NAME
    
    Dim result As DataLoadResult
    result = ProcessCategory("Capex EPC", "Erreur lors du traitement du Capex EPC")
    
    If result = DataLoadResult.Success Then
        LogInfo PROC_NAME & "_End", "Traitement du Capex EPC terminé.", PROC_NAME, MODULE_NAME
    ElseIf result = DataLoadResult.Cancelled Then
        LogInfo PROC_NAME & "_Cancelled", "Traitement du Capex EPC annulé.", PROC_NAME, MODULE_NAME
    Else
        LogError PROC_NAME & "_Error", "Erreur lors du traitement du Capex EPC.", PROC_NAME, MODULE_NAME
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnTechEcoAnalysis(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnTechEcoAnalysis"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Opening technical and economic analysis tool.", PROC_NAME, MODULE_NAME
    
    ' TODO: Implement the technical and economic analysis logic
    ShowNotImplementedMessage "technical and economic analysis"
    
    LogInfo PROC_NAME & "_End", "Technical and economic analysis completed.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnBusinessPlan(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnBusinessPlan"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Opening business plan tool.", PROC_NAME, MODULE_NAME
    
    ' TODO: Implement the business plan logic
    ShowNotImplementedMessage "business plan"
    
    LogInfo PROC_NAME & "_End", "Business plan process completed.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnElecOptim(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnElecOptim"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Opening electricity optimization tool.", PROC_NAME, MODULE_NAME
    
    ' TODO: Implement the electricity optimization logic
    ShowNotImplementedMessage "electricity optimization"
    
    LogInfo PROC_NAME & "_End", "Electricity optimization process completed.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnLCA(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnLCA"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Opening LCA tool.", PROC_NAME, MODULE_NAME
    
    ' TODO: Implement the Life Cycle Analysis logic
    ShowNotImplementedMessage "Life Cycle Analysis"
    
    LogInfo PROC_NAME & "_End", "Life Cycle Analysis process completed.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnServerFiles(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnServerFiles"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Opening server files browser.", PROC_NAME, MODULE_NAME
    
    ' TODO: Implement the server files browser logic
    ShowNotImplementedMessage "server files browser"
    
    LogInfo PROC_NAME & "_End", "Server files browser closed.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnSummarySheets(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnSummarySheets"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Opening summary sheets.", PROC_NAME, MODULE_NAME
    
    ' TODO: Implement the summary sheets logic
    ShowNotImplementedMessage "summary sheets"
    
    LogInfo PROC_NAME & "_End", "Summary sheets process completed.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub OnOpex(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "OnOpex"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Opening OPEX management.", PROC_NAME, MODULE_NAME
    
    ' TODO: Implement the OPEX management logic
    ShowNotImplementedMessage "OPEX management"
    
    LogInfo PROC_NAME & "_End", "OPEX management process completed.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessInjectAllPowerQueries(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessInjectAllPowerQueries"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Starting PowerQuery injection process.", PROC_NAME, MODULE_NAME
    
    ' Call the PowerQuery manager to inject all queries
    ' TODO: Implement proper PowerQuery injection logic
    ShowNotImplementedMessage "PowerQuery injection"
    
    LogInfo PROC_NAME & "_End", "PowerQuery injection process completed.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessCleanupAllPowerQueries(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessCleanupAllPowerQueries"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Starting PowerQuery cleanup process.", PROC_NAME, MODULE_NAME
    
    ' TODO: Implement PowerQuery cleanup logic
    ShowNotImplementedMessage "PowerQuery cleanup"
    
    LogInfo PROC_NAME & "_End", "PowerQuery cleanup process completed.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ProcessDebugRagicDictionary(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessDebugRagicDictionary"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Starting Ragic dictionary debug process.", PROC_NAME, MODULE_NAME
    
    ' TODO: Implement Ragic dictionary debug logic
    ShowNotImplementedMessage "Ragic dictionary debug"
    
    LogInfo PROC_NAME & "_End", "Ragic dictionary debug process completed.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

Public Sub ShowNotImplementedMessage(featureName As String)
    ShowInfoMessage "Not Implemented", "The " & featureName & " feature is not yet implemented."
End Sub

Public Sub OnTechnicalAnalysisClick()
    ShowNotImplementedMessage "technical and economic analysis"
End Sub

Public Sub OnBusinessPlanClick()
    ShowNotImplementedMessage "business plan"
End Sub

Public Sub OnElectricityOptimizationClick()
    ShowNotImplementedMessage "electricity optimization"
End Sub

Public Sub OnLifeCycleAnalysisClick()
    ShowNotImplementedMessage "Life Cycle Analysis"
End Sub

Public Sub OnServerFilesBrowserClick()
    ShowNotImplementedMessage "server files browser"
End Sub

Public Sub OnSummarySheetsClick()
    ShowNotImplementedMessage "summary sheets"
End Sub

Public Sub OnOPEXManagementClick()
    ShowNotImplementedMessage "OPEX management"
End Sub

Public Sub OnPowerQueryInjectionClick()
    ShowNotImplementedMessage "PowerQuery injection"
End Sub

Public Sub OnPowerQueryCleanupClick()
    ShowNotImplementedMessage "PowerQuery cleanup"
End Sub

Public Sub OnRagicDictionaryDebugClick()
    ShowNotImplementedMessage "Ragic dictionary debug"
End Sub
