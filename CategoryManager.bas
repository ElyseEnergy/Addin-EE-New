Attribute VB_Name = "CategoryManager"
' Module : CategoryManager.bas
' ==============================
' G�re toutes les cat�gories et leurs configurations sous forme de module standard.
' Permet l'initialisation, l'ajout, la recherche et la r�cup�ration des cat�gories.
Option Explicit

' Categories
' ----------
' Tableau global contenant toutes les cat�gories charg�es.
Public categories() As categoryInfo
Public CategoriesCount As Long

' InitCategories
' --------------
' Initialise la liste des cat�gories et charge les donn�es associ�es.
' Retourne : Rien
Public Sub InitCategories()
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "InitCategories"
    Const MODULE_NAME As String = "CategoryManager"
    
    CategoriesCount = 0
    ReDim categories(1 To 1)
    
    ' # Engineering
    ' ## Technologies
    AddCategory "Compression", "Pas de filtrage", "Compression", "fiches-techniques/16.csv", "Technologies"
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
    ' ## M�triques de r�f�rence
    AddCategory "M�triques de base", "Pas de filtrage", "M�triques de base", "plant-power-requirement-v2/2.csv", "Engineering Metrics"
    AddCategory "M�triques expert", "Pas de filtrage", "M�triques expert", "plant-power-requirement-v2/5.csv", "Engineering Metrics"
    AddCategory "Timings de r�f�rence", "Pas de filtrage", "Timings de r�f�rence", "plant-power-requirement-v2/10.csv", "Engineering Metrics"
    ' ## LCA
    AddCategory "M�triques RED III", "Pas de filtrage", "M�triques RED III", "red-ii/7.csv", "LCA"
    AddCategory "Emissions", "Pas de filtrage", "Emissions", "red-ii/8.csv", "LCA"
    ' ## Logistique
    AddCategory "Infra et logistique", "Pas de filtrage", "Infra et logistique", "guilhem-infra/5.csv", "Log"

    ' Finances
    AddCategory "Budget Corpo", "budget Associ�", "Budget Corpo", "newbudget/2.csv", "Finances"
    AddCategory "D�tails Budgets", "budget Associ�", "D�tails Budgets", "newbudget/8.csv", "Finances"
    AddCategory "DIB", "Pole & D�partement", "DIB", "items-budgtaires/1.csv", "Finances"
    AddCategory "Demandes d'achat", "Centre de co�t", "Demandes d'achat", "mouvements/15.csv", "Finances"
    AddCategory "R�ceptions", "Nom Founisseur", "R�ceptions", "mouvements/20.csv", "Finances"
    
    ' Projets
    AddCategory "Sc�narios techniques", "Projet", "Sc�narios techniques", "scnarios-technico-conomiques/1.csv", "Projets"
    AddCategory "Plannings de phases", "Project", "Plannings de phases", "tests/6.csv", "Projets", "Planning link"
    AddCategory "Plannings de sous phases", "Project", "Plannings de sous phases", "tests/7.csv", "Projets", "Planning link"
    AddCategory "Budget Projet", "budget Associ�", "Budget Projet", "newbudget/2.csv", "Projets"
    AddCategory "Devex", "Projet", "Devex", "costing/16.csv", "Projets"
    AddCategory "Capex", "Projet", "Capex", "costing/2.csv", "Projets"
    AddCategory "Capex EPC", "Projet", "Capex EPC", "costing/13.csv", "Projets"    'TODO : AddCategory "Opex", "Projet", "Opex", "costing/opex.csv", "Projets"
    'TODO : AddCategory "Pricings", "Projet", "Pricings", "costing/pricings.csv", "Projets"
    
    ' Maintenant que toutes les cat�gories sont initialis�es, on peut charger le dictionnaire
    LoadRagicDictionary
    
    Exit Sub
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de l'initialisation des cat�gories"
End Sub

' Ajoute une cat�gorie au tableau des cat�gories.
' Param�tres :
'   name (String) : Nom de la cat�gorie
'   filterLevel (String) : Niveau de filtrage principal
'   displayName (String) : Nom d'affichage
'   path (String) : Chemin du fichier source
'   categoryGroup (String) : Groupe de la cat�gorie
'   secondaryFilterLevel (String, optionnel) : Niveau de filtrage secondaire
'   sheetName (String, optionnel) : Nom de la feuille associ�e
Public Sub AddCategory(Name As String, FilterLevel As String, DisplayName As String, path As String, CategoryGroup As String, Optional SecondaryFilterLevel As String = "", Optional SheetName As String = "")
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "AddCategory"
    Const MODULE_NAME As String = "CategoryManager"
      If Name = "" Or DisplayName = "" Or path = "" Or CategoryGroup = "" Then
        HandleError MODULE_NAME, PROC_NAME, "Param�tres invalides pour l'ajout de cat�gorie"
        Exit Sub
    End If
    
    Dim idx As Long
    If CategoriesCount = 0 Then
        idx = 1
    Else
        idx = CategoriesCount + 1
    End If
    
    ReDim Preserve categories(1 To idx)
    categories(idx).categoryName = Name
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
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de l'ajout de la cat�gorie: " & Name
End Sub

' Retourne l'index d'une cat�gorie par son nom d'affichage.
' Param�tres :
'   displayName (String) : Nom d'affichage de la cat�gorie
' Retour :
'   Long (index de la cat�gorie ou 0 si non trouv�e)
Public Function GetCategoryIndexByName(DisplayName As String) As Long
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetCategoryIndexByName"
    Const MODULE_NAME As String = "CategoryManager"
      If DisplayName = "" Then
        HandleError MODULE_NAME, PROC_NAME, "Nom d'affichage vide"
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
    GetCategoryIndexByName = 0 ' Non trouv�
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la recherche de l'index de la cat�gorie: " & DisplayName
    GetCategoryIndexByName = 0
End Function

' Retourne une cat�gorie par son nom d'affichage.
' Param�tres :
'   displayName (String) : Nom d'affichage de la cat�gorie
' Retour :
'   CategoryInfo (structure de la cat�gorie)
Public Function GetCategoryByName(DisplayName As String) As categoryInfo
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetCategoryByName"
    Const MODULE_NAME As String = "CategoryManager"
      If DisplayName = "" Then
        HandleError MODULE_NAME, PROC_NAME, "Nom d'affichage vide"
        Exit Function
    End If
    
    Dim idx As Long
    idx = GetCategoryIndexByName(DisplayName)
    If idx > 0 Then
        GetCategoryByName = categories(idx)
    End If
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la r�cup�ration de la cat�gorie: " & DisplayName
End Function

' Retourne toutes les cat�gories sous forme de tableau.
' Retour :
'   CategoryInfo() (tableau de CategoryInfo)
Public Function GetAllCategories() As categoryInfo()
    On Error GoTo ErrorHandler

    Const PROC_NAME As String = "GetAllCategories"
    Const MODULE_NAME As String = "CategoryManager"

    If CategoriesCount = 0 Then
        HandleError MODULE_NAME, PROC_NAME, "Aucune cat�gorie n'est d�finie"
        Exit Function
    End If

    ' Cr�er un nouveau tableau de la bonne taille
    Dim result() As categoryInfo
    ReDim result(LBound(categories) To UBound(categories))

    ' Copier chaque �l�ment (copie profonde)
    Dim i As Long
    For i = LBound(categories) To UBound(categories)
        With categories(i)
            result(i).categoryName = .categoryName
            result(i).FilterLevel = .FilterLevel
            result(i).SecondaryFilterLevel = .SecondaryFilterLevel
            result(i).DisplayName = .DisplayName
            result(i).URL = .URL
            result(i).PowerQueryName = .PowerQueryName
            result(i).CategoryGroup = .CategoryGroup
            result(i).SheetName = .SheetName
        End With
    Next i

    GetAllCategories = result
    Exit Function

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la r�cup�ration de toutes les cat�gories"
End Function
