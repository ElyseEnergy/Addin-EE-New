' Module : CategoryManager.bas
' Gère toutes les catégories et leurs configurations sous forme de module standard
Option Explicit

Public Categories() As CategoryInfo
Public CategoriesCount As Long

' Initialise les catégories
Public Sub InitCategories()
    CategoriesCount = 0
    ReDim Categories(1 To 1)
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
        AddCategory "Plannings de sous phases", "Level 0 Complete Name", "Plannings de sous phases", "tests/7.csv", "Projets", "Planning link"
        AddCategory "Budget Projet", "budget Associé", "Budget Projet", "newbudget/2.csv", "Projets"
        AddCategory "Devex", "Projet", "Devex", "costing/16.csv", "Projets"
        AddCategory "Capex", "Projet", "Capex", "costing/2.csv", "Projets"
        AddCategory "Capex EPC", "Projet", "Capex EPC", "costing/13.csv", "Projets"
        'TODO : AddCategory "Opex", "Projet", "Opex", "costing/opex.csv", "Projets"
        'TODO : AddCategory "Pricings", "Projet", "Pricings", "costing/pricings.csv", "Projets"

    
End Sub

' Ajoute une catégorie au tableau
Public Sub AddCategory(name As String, filterLevel As String, displayName As String, path As String, categoryGroup As String, Optional secondaryFilterLevel As String = "")
    Dim idx As Long
    If CategoriesCount = 0 Then
        idx = 1
    Else
        idx = CategoriesCount + 1
    End If
    ReDim Preserve Categories(1 To idx)
    Categories(idx).categoryName = name
    Categories(idx).filterLevel = filterLevel
    Categories(idx).SecondaryFilterLevel = secondaryFilterLevel
    Categories(idx).displayName = displayName
    Categories(idx).URL = env.RAGIC_BASE_URL & path & env.RAGIC_API_PARAMS
    Categories(idx).PowerQueryName = "PQ_" & Utilities.SanitizeTableName(name)
    Categories(idx).categoryGroup = categoryGroup
    CategoriesCount = idx
End Sub

' Retourne l'index d'une catégorie par son nom d'affichage
Public Function GetCategoryIndexByName(displayName As String) As Long
    Dim i As Long
    For i = 1 To CategoriesCount
        If Categories(i).displayName = displayName Then
            GetCategoryIndexByName = i
            Exit Function
        End If
    Next i
    GetCategoryIndexByName = 0 ' Non trouvé
End Function

' Retourne une catégorie par son nom d'affichage
Public Function GetCategoryByName(displayName As String) As CategoryInfo
    Dim idx As Long
    idx = GetCategoryIndexByName(displayName)
    If idx > 0 Then
        GetCategoryByName = Categories(idx)
    End If
End Function

' Retourne toutes les catégories sous forme de tableau
Public Function GetAllCategories() As Variant
    GetAllCategories = Categories
End Function
