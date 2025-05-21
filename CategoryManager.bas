' Module : CategoryManager.bas
' Gère toutes les catégories et leurs configurations sous forme de module standard
Option Explicit

Public Categories() As CategoryInfo
Public CategoriesCount As Long

' Initialise les catégories
Public Sub InitCategories()
    CategoriesCount = 0
    ReDim Categories(1 To 1)
    ' Technologies
    AddCategory "Compression", "Molecule Type", "Compression", "fiches-techniques/16/3.csv", "Technologies"
    AddCategory "CO2 general parameters", "Pas de filtrage", "CO2 general parameters", "fiches-techniques/25.csv", "Technologies"
    AddCategory "H2 general parameters", "Pas de filtrage", "H2 general parameters", "fiches-techniques/24.csv", "Technologies"
    AddCategory "CO2 Capture", "Pas de filtrage", "CO2 Capture", "fiches-techniques/18.csv", "Technologies"
    AddCategory "H2 waters electrolysis", "Brand", "H2 waters electrolysis", "fiches-techniques/26.csv", "Technologies"
    AddCategory "MeOH - CO2-to-Methanol Synthesis", "Brand", "MeOH - CO2-to-Methanol Synthesis", "fiches-techniques/19.csv", "Technologies"
    AddCategory "MeOH - Biomass Gasification Synthesis", "Brand", "MeOH - Biomass Gasification Synthesis", "fiches-techniques/20.csv", "Technologies"
    AddCategory "SAF - MtJ Synthesis", "Type", "SAF - MtJ Synthesis", "fiches-techniques/21.csv", "Technologies"
    AddCategory "SAF - BtJ/e-BtJ Synthesis", "Type/Project", "SAF - BtJ/e-BtJ Synthesis", "fiches-techniques/22.csv", "Technologies"
    ' Utilities
    AddCategory "Chiller", "Pas de filtrage", "Chiller", "utilities/5.csv", "Utilities"
    AddCategory "Cooling Water Production", "Pas de filtrage", "Cooling Water Production", "utilities/4.csv", "Utilities"
    AddCategory "Heat Production", "Pas de filtrage", "Heat Production", "utilities/3.csv", "Utilities"
    AddCategory "Other utilities", "Pas de filtrage", "Other utilities", "utilities/7.csv", "Utilities"
    AddCategory "Power losses", "Pas de filtrage", "Power losses", "utilities/9.csv", "Utilities"
    AddCategory "WasteWater Treatment", "Pas de filtrage", "WasteWater Treatment", "utilities/6.csv", "Utilities"
    AddCategory "Water Treatment", "Pas de filtrage", "Water Treatment", "utilities/2.csv", "Utilities"
End Sub

' Ajoute une catégorie au tableau
Public Sub AddCategory(name As String, filterLevel As String, displayName As String, path As String, categoryType As String)
    Dim idx As Long
    If CategoriesCount = 0 Then
        idx = 1
    Else
        idx = CategoriesCount + 1
    End If
    ReDim Preserve Categories(1 To idx)
    Categories(idx).CategoryName = categoryType
    Categories(idx).FilterLevel = filterLevel
    Categories(idx).DisplayName = displayName
    Categories(idx).URL = ENV.RAGIC_BASE_URL & path & ENV.RAGIC_API_PARAMS
    If categoryType = "Utilities" Then
        Categories(idx).PowerQueryName = "PQ_Utility_" & Replace(name, " ", "_") & "_Data"
    Else
        Categories(idx).PowerQueryName = "PQ_" & Replace(name, " ", "_") & "_Data"
    End If
    CategoriesCount = idx
End Sub

' Retourne l'index d'une catégorie par son nom d'affichage
Public Function GetCategoryIndexByName(displayName As String) As Long
    Dim i As Long
    For i = 1 To CategoriesCount
        If Categories(i).DisplayName = displayName Then
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