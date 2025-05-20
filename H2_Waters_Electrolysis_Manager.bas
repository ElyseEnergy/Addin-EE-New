' Module: H2_Waters_Electrolysis_Manager
' Gère le traitement des données d'électrolyse de l'eau
Option Explicit

Public Sub ProcessH2Electrolysis()
    ' Initialiser les catégories (à faire une seule fois au démarrage du projet)
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
