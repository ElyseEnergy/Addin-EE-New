Attribute VB_Name = "CategoryManager"
' Module : CategoryManager.bas
' Gère toutes les catégories et leurs configurations sous forme de module standard
Option Explicit

Public categories() As CategoryInfo
Public CategoriesCount As Long

' Initialise les catégories
Public Sub InitCategories()
    On Error GoTo ErrorHandler
    
    CategoriesCount = 0
    ReDim categories(1 To 1)
    LoadRagicDictionary
    
    ' # Engineering
    ' ## Technologies
    AddCategory "Compression", "Pas de filtrage", "Compression", "fiches-techniques/16/3.csv", "Technologies"
    AddCategory "CO2 general parameters", "Pas de filtrage", "CO2 general parameters", "fiches-techniques/25.csv", "Technologies"
    AddCategory "H2 general parameters", "Pas de filtrage", "H2 general parameters", "fiches-techniques/24.csv", "Technologies"
    AddCategory "CO2 Capture", "Brand", "CO2 Capture", "fiches-techniques/18.csv", "Technologies"
    AddCategory "H2 waters electrolysis", "Brand", "H2 waters electrolysis", "fiches-techniques/26.csv", "Technologies"
    AddCategory "MeOH - CO2-to-Methanol Synthesis", "Brand", "MeOH - CO2-to-Methanol Synthesis", "fiches-techniques/19.csv", "Technologies"
    AddCategory "MeOH - Biomass Gasification Synthesis", "Pas de filtrage", "MeOH - Biomass Gasification Synthesis", "fiches-techniques/20.csv", "Technologies"
    AddCategory "SAF - MtJ Synthesis", "Pas de filtrage", "SAF - MtJ Synthesis", "fiches-techniques/21.csv", "Technologies"
    AddCategory "SAF - BtJ/e-BtJ Synthesis", "Pas de filtrage", "SAF - BtJ/e-BtJ Synthesis", "fiches-techniques/22.csv", "Technologies"
    ' ## Utilities
    AddCategory "Chiller", "Pas de filtrage", "Chiller", "utilities/5.csv", "Utilities"
    AddCategory "Cooling Water Production", "Pas de filtrage", "Cooling Water Production", "utilities/4.csv", "Utilities"
    AddCategory "Heat Production", "Pas de filtrage", "Heat Production", "utilities/3.csv", "Utilities"
    AddCategory "Other utilities", "Pas de filtrage", "Other utilities", "utilities/7.csv", "Utilities"
    AddCategory "Power losses", "Pas de filtrage", "Power losses", "utilities/9.csv", "Utilities"
    AddCategory "WasteWater Treatment", "Pas de filtrage", "WasteWater Treatment", "utilities/6.csv", "Utilities"
    AddCategory "Water Treatment", "Pas de filtrage", "Water Treatment", "utilities/2.csv", "Utilities"
    ' ## Métriques de référence
    AddCategory "Métriques de base", "Pas de filtrage", "Métriques de base", "plant-power-requirement-v2/2.csv", "Engineering Metrics"
    AddCategory "Métriques expert", "Pas de filtrage", "Métriques expert", "plant-power-requirement-v2/5.csv", "Engineering Metrics"
    AddCategory "Timings de référence", "Pas de filtrage", "Timings de référence", "plant-power-requirement-v2/10.csv", "Engineering Metrics"
    ' ## LCA
    AddCategory "Métriques RED III", "Pas de filtrage", "Métriques RED III", "red-ii/7.csv", "LCA"
    AddCategory "Emissions", "Pas de filtrage", "Emissions", "red-ii/8.csv", "LCA"
    ' ## Logistique
    AddCategory "Infra et logistique", "Pas de filtrage", "Infra et logistique", "guilhem-infra/5.csv", "Log"

    ' Finances
    AddCategory "Budget Corpo", "budget Associé", "Budget Corpo", "newbudget/2.csv", "Finances"
    AddCategory "Détails Budgets", "budget Associé", "Détails Budgets", "newbudget/8.csv", "Finances"
    AddCategory "DIB", "Pole & Département", "DIB", "items-budgtaires/1.csv", "Finances"
    AddCategory "Demandes d'achat", "Centre de coût", "Demandes d'achat", "mouvements/15.csv", "Finances"
    AddCategory "Réceptions", "Nom Founisseur", "Réceptions", "mouvements/20.csv", "Finances"
    
    ' Projets
    AddCategory "Scénarios techniques", "Projet", "Scénarios techniques", "scnarios-technico-conomiques/1.csv", "Projets"
    AddCategory "Plannings de phases", "Project", "Plannings de phases", "tests/6.csv", "Projets", "Planning link"
    AddCategory "Plannings de sous phases", "Project", "Plannings de sous phases", "tests/7.csv", "Projets", "Planning link"
    AddCategory "Budget Projet", "budget Associé", "Budget Projet", "newbudget/2.csv", "Projets"
    AddCategory "Devex", "Projet", "Devex", "costing/16.csv", "Projets"
    AddCategory "Capex", "Projet", "Capex", "costing/2.csv", "Projets"
    AddCategory "Capex EPC", "Projet", "Capex EPC", "costing/13.csv", "Projets"
    'TODO : AddCategory "Opex", "Projet", "Opex", "costing/opex.csv", "Projets"
    'TODO : AddCategory "Pricings", "Projet", "Pricings", "costing/pricings.csv", "Projets"
    
    Exit Sub
    
ErrorHandler:
    HandleError "CategoryManager", "InitCategories", "Erreur lors de l'initialisation des catégories"
End Sub

' Ajoute une catégorie au tableau
Public Sub AddCategory(Name As String, FilterLevel As String, DisplayName As String, path As String, CategoryGroup As String, Optional SecondaryFilterLevel As String = "", Optional SheetName As String = "")
    On Error GoTo ErrorHandler
    
    If Name = "" Or DisplayName = "" Or path = "" Or CategoryGroup = "" Then
        HandleError "CategoryManager", "AddCategory", "Paramètres invalides pour l'ajout de catégorie"
        Exit Sub
    End If
    
    Dim idx As Long
    If CategoriesCount = 0 Then
        idx = 1
    Else
        idx = CategoriesCount + 1
    End If
    
    ReDim Preserve categories(1 To idx)
    categories(idx).CategoryName = Name
    categories(idx).FilterLevel = FilterLevel
    categories(idx).SecondaryFilterLevel = SecondaryFilterLevel
    categories(idx).DisplayName = DisplayName
    categories(idx).URL = env.RAGIC_BASE_URL & path & env.RAGIC_API_PARAMS
    categories(idx).PowerQueryName = "PQ_" & Utilities.SanitizeTableName(Name)
    categories(idx).CategoryGroup = CategoryGroup
    If SheetName = "" Then SheetName = DisplayName
    categories(idx).SheetName = SheetName
    CategoriesCount = idx
    Exit Sub
    
ErrorHandler:
    HandleError "CategoryManager", "AddCategory", "Erreur lors de l'ajout de la catégorie: " & Name
End Sub

' Retourne l'index d'une catégorie par son nom d'affichage
Public Function GetCategoryIndexByName(DisplayName As String) As Long
    On Error GoTo ErrorHandler
    
    If DisplayName = "" Then
        HandleError "CategoryManager", "GetCategoryIndexByName", "Nom d'affichage vide"
        GetCategoryIndexByName = 0
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To CategoriesCount
        If categories(i).DisplayName = DisplayName Then
            GetCategoryIndexByName = i
            Exit Function
        End If
    Next i
    GetCategoryIndexByName = 0 ' Non trouvé
    Exit Function
    
ErrorHandler:
    HandleError "CategoryManager", "GetCategoryIndexByName", "Erreur lors de la recherche de l'index de la catégorie: " & DisplayName
    GetCategoryIndexByName = 0
End Function

' Retourne une catégorie par son nom d'affichage
Public Function GetCategoryByName(DisplayName As String) As CategoryInfo
    On Error GoTo ErrorHandler
    
    If DisplayName = "" Then
        HandleError "CategoryManager", "GetCategoryByName", "Nom d'affichage vide"
        Exit Function
    End If
    
    Dim idx As Long
    idx = GetCategoryIndexByName(DisplayName)
    If idx > 0 Then
        GetCategoryByName = categories(idx)
    End If
    Exit Function
    
ErrorHandler:
    HandleError "CategoryManager", "GetCategoryByName", "Erreur lors de la récupération de la catégorie: " & DisplayName
End Function

' Retourne toutes les catégories sous forme de tableau
Public Function GetAllCategories() As Variant
    On Error GoTo ErrorHandler
    
    If CategoriesCount = 0 Then
        HandleError "CategoryManager", "GetAllCategories", "Aucune catégorie n'est définie"
        Exit Function
    End If
    
    GetAllCategories = categories
    Exit Function
    
ErrorHandler:
    HandleError "CategoryManager", "GetAllCategories", "Erreur lors de la récupération de toutes les catégories"
End Function

