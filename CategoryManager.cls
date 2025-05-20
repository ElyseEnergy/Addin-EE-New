' Class: CategoryManager
' Gère toutes les catégories et leurs configurations
Option Explicit

' Type pour stocker les informations de catégorie
Private Type CategoryInfo
    CategoryName As String        ' Technologies ou Utilities
    FilterLevel As String         ' Niveau de filtrage (Molecule Type, Brand, Type, etc.)
    DisplayName As String         ' Nom d'affichage pour l'utilisateur
    URL As String                 ' URL complète
End Type

' Collections privées pour stocker les catégories
Private m_technologies As Collection
Private m_utilities As Collection

' Constructeur
Private Sub Class_Initialize()
    Set m_technologies = New Collection
    Set m_utilities = New Collection
    LoadCategories
End Sub

' Chargement des catégories
Private Sub LoadCategories()
    ' Technologies
    AddTechnology "Compression", "Molecule Type", "fiches-techniques/16/3.csv"
    AddTechnology "CO2 general parameters", "Pas de filtrage", "fiches-techniques/25.csv"
    AddTechnology "H2 general parameters", "Pas de filtrage", "fiches-techniques/24.csv"
    AddTechnology "CO2 Capture", "Brand", "fiches-techniques/18.csv"
    AddTechnology "H2 waters electrolysis", "Brand", "fiches-techniques/26.csv"
    AddTechnology "MeOH - CO2-to-Methanol Synthesis", "Brand", "fiches-techniques/19.csv"
    AddTechnology "MeOH - Biomass Gasification Synthesis", "Brand", "fiches-techniques/20.csv"
    AddTechnology "SAF - MtJ Synthesis", "Type", "fiches-techniques/21.csv"
    AddTechnology "SAF - BtJ/e-BtJ Synthesis", "Type/Project", "fiches-techniques/22.csv"
    
    ' Utilities
    AddUtility "Chiller", "utilities/5.csv"
    AddUtility "Cooling Water Production", "utilities/4.csv"
    AddUtility "Heat Production", "utilities/3.csv"
    AddUtility "Other utilities", "utilities/7.csv"
    AddUtility "Power losses", "utilities/9.csv"
    AddUtility "WasteWater Treatment", "utilities/6.csv"
    AddUtility "Water Treatment", "utilities/2.csv"
End Sub

' Méthodes privées pour ajouter des catégories
Private Sub AddTechnology(name As String, filterLevel As String, path As String)
    Dim cat As CategoryInfo
    cat.CategoryName = "Technologies"
    cat.FilterLevel = filterLevel
    cat.DisplayName = name
    cat.URL = ENV.RAGIC_BASE_URL & path & ENV.RAGIC_API_PARAMS
    m_technologies.Add cat, name
End Sub

Private Sub AddUtility(name As String, path As String)
    Dim cat As CategoryInfo
    cat.CategoryName = "Utilities"
    cat.FilterLevel = "Pas de filtrage"
    cat.DisplayName = name
    cat.URL = ENV.RAGIC_BASE_URL & path & ENV.RAGIC_API_PARAMS
    m_utilities.Add cat, name
End Sub

' Interface publique
Public Function GetTechnologies() As Collection
    Set GetTechnologies = m_technologies
End Function

Public Function GetUtilities() As Collection
    Set GetUtilities = m_utilities
End Function


Public Function GetAllCategories() As Collection
    Dim allCategories As New Collection
    Dim i As Long
    
    For i = 1 To m_technologies.Count
        allCategories.Add m_technologies(i)
    Next i
    
    For i = 1 To m_utilities.Count
        allCategories.Add m_utilities(i)
    Next i
    
    Set GetAllCategories = allCategories
End Function 