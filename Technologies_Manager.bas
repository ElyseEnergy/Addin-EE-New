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


Public Sub ProcessH2Electrolysis(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ' Initialiser les catégories(à faire une seule fois au démarrage du projet)
    If CategoriesCount = 0 Then InitCategories
    
    ' Créer les informations de chargement
    Dim loadInfo As DataLoadInfo
    loadInfo.Category = GetCategoryByName("H2 waters electrolysis")
    If loadInfo.Category.DisplayName = "" Then
        MsgBox "Catégorie 'H2 waters electrolysis' non trouvée", vbExclamation
        Exit Sub
    End If
    
    loadInfo.PreviewRows = 3 ' Nombre de lignes pour la prévisualisation
    
    ' Traiter les données
    If Not DataLoaderManager.ProcessDataLoad(loadInfo) Then
        MsgBox "Erreur lors du traitement des données d'électrolyse", vbExclamation
        Exit Sub
    End If
    
    ' Plus besoin de protéger la feuille ici, c'est géré dans DataLoaderManager
End Sub

Public Sub ProcessCO2Capture(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ' Initialiser les catégories(à faire une seule fois au démarrage du projet)
    If CategoriesCount = 0 Then InitCategories
    
    ' Créer les informations de chargement
    Dim loadInfo As DataLoadInfo
    loadInfo.Category = GetCategoryByName("CO2 Capture")
    If loadInfo.Category.DisplayName = "" Then
        MsgBox "Catégorie 'CO2 Capture' non trouvée", vbExclamation
        Exit Sub
    End If
    
    loadInfo.PreviewRows = 3 ' Nombre de lignes pour la prévisualisation
    
    ' Traiter les données
    If Not DataLoaderManager.ProcessDataLoad(loadInfo) Then
        MsgBox "Erreur lors du traitement des données CO2 Capture", vbExclamation
        Exit Sub
    End If
    
    ' Plus besoin de protéger la feuille ici, c'est géré dans DataLoaderManager
End Sub

Public Sub ProcessCO2General(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ' Initialiser les catégories(à faire une seule fois au démarrage du projet)
    If CategoriesCount = 0 Then InitCategories
    
    ' Créer les informations de chargement
    Dim loadInfo As DataLoadInfo
    loadInfo.Category = GetCategoryByName("CO2 general parameters")
    If loadInfo.Category.DisplayName = "" Then
        MsgBox "Catégorie 'CO2 general parameters' non trouvée", vbExclamation
        Exit Sub
    End If
    
    loadInfo.PreviewRows = 3 ' Nombre de lignes pour la prévisualisation
    
    ' Traiter les données
    If Not DataLoaderManager.ProcessDataLoad(loadInfo) Then
        MsgBox "Erreur lors du traitement des données CO2 General Parameters", vbExclamation
        Exit Sub
    End If
End Sub
