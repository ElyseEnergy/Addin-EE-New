Attribute VB_Name = "CategoryManager"
Option Explicit

Public Categories() As CategoryInfo
Public CategoriesCount As Long

' InitCategories
' --------------
' Initialise la liste des catégories et charge les données associées.
' Retourne : Rien
Public Sub InitCategories()
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "InitCategories"
    Const MODULE_NAME As String = "CategoryManager"
    
    CategoriesCount = 0
    ' Allouer une taille suffisante dès le départ pour éviter les ReDim Preserve
    ReDim Categories(1 To 50) ' Augmenter si plus de 50 catégories sont prévues
    
    ' # Engineering
    ' ## Technologies
    Call AddCategory("Compression", "Pas de filtrage", "Compression", "fiches-techniques/16.csv", "Technologies")
    Call AddCategory("CO2 general parameters", "Pas de filtrage", "CO2 general parameters", "fiches-techniques/25.csv", "Technologies")
    Call AddCategory("H2 general parameters", "Pas de filtrage", "H2 general parameters", "fiches-techniques/24.csv", "Technologies")
    Call AddCategory("CO2 Capture", "Brand", "CO2 Capture", "fiches-techniques/18.csv", "Technologies")
    Call AddCategory("H2 waters electrolysis", "Brand", "H2 waters electrolysis", "fiches-techniques/26.csv", "Technologies")
    Call AddCategory("MeOH - CO2-to-Methanol Synthesis", "Brand", "MeOH - CO2-to-Methanol Synthesis", "fiches-techniques/19.csv", "Technologies")
    Call AddCategory("MeOH - Biomass Gasification Synthesis", "Pas de filtrage", "MeOH - Biomass Gasification Synthesis", "fiches-techniques/20.csv", "Technologies")
    Call AddCategory("SAF - MtJ Synthesis", "Pas de filtrage", "SAF - MtJ Synthesis", "fiches-techniques/21.csv", "Technologies")
    Call AddCategory("SAF - BtJ/e-BtJ Synthesis", "Pas de filtrage", "SAF - BtJ/e-BtJ Synthesis", "fiches-techniques/22.csv", "Technologies")
    ' ## Utilities
    Call AddCategory("Chiller", "Pas de filtrage", "Chiller", "utilities/5.csv", "Utilities")
    Call AddCategory("Cooling Water Production", "Pas de filtrage", "Cooling Water Production", "utilities/4.csv", "Utilities")
    Call AddCategory("Heat Production", "Pas de filtrage", "Heat Production", "utilities/3.csv", "Utilities")
    Call AddCategory("Other utilities", "Pas de filtrage", "Other utilities", "utilities/7.csv", "Utilities")
    Call AddCategory("Power losses", "Pas de filtrage", "Power losses", "utilities/9.csv", "Utilities")
    Call AddCategory("WasteWater Treatment", "Pas de filtrage", "WasteWater Treatment", "utilities/6.csv", "Utilities")
    Call AddCategory("Water Treatment", "Pas de filtrage", "Water Treatment", "utilities/2.csv", "Utilities")
    ' ## Métriques de référence
    Call AddCategory("Métriques de base", "Pas de filtrage", "Métriques de base", "plant-power-requirement-v2/2.csv", "Engineering Metrics")
    Call AddCategory("Métriques expert", "Pas de filtrage", "Métriques expert", "plant-power-requirement-v2/5.csv", "Engineering Metrics")
    Call AddCategory("Timings de référence", "Pas de filtrage", "Timings de référence", "plant-power-requirement-v2/10.csv", "Engineering Metrics")
    ' ## LCA
    Call AddCategory("Métriques RED III", "Pas de filtrage", "Métriques RED III", "red-ii/7.csv", "LCA")
    Call AddCategory("Emissions", "Pas de filtrage", "Emissions", "red-ii/8.csv", "LCA")
    ' ## Logistique
    Call AddCategory("Infra et logistique", "Pas de filtrage", "Infra et logistique", "guilhem-infra/5.csv", "Log")

    ' Finances
    Call AddCategory("Budget Corpo", "budget Associé", "Budget Corpo", "newbudget/2.csv", "Finances")
    Call AddCategory("Détails Budgets", "budget Associé", "Détails Budgets", "newbudget/8.csv", "Finances")
    Call AddCategory("DIB", "Pole & Département", "DIB", "items-budgtaires/1.csv", "Finances")
    Call AddCategory("Demandes d'achat", "Centre de coût", "Demandes d'achat", "mouvements/15.csv", "Finances")
    Call AddCategory("Réceptions", "Nom Founisseur", "Réceptions", "mouvements/20.csv", "Finances")
    
    ' Projets
    Call AddCategory("Scénarios techniques", "Projet", "Scénarios techniques", "scnarios-technico-conomiques/1.csv", "Projets")
    Call AddCategory("Plannings de phases", "Project", "Plannings de phases", "tests/6.csv", "Projets", "Planning link")
    Call AddCategory("Plannings de sous phases", "Project", "Plannings de sous phases", "tests/7.csv", "Projets", "Planning link")
    Call AddCategory("Budget Projet", "budget Associé", "Budget Projet", "newbudget/2.csv", "Projets")
    Call AddCategory("Devex", "Projet", "Devex", "costing/16.csv", "Projets")
    Call AddCategory("Capex", "Projet", "Capex", "costing/2.csv", "Projets")
    Call AddCategory("Capex EPC", "Projet", "Capex EPC", "costing/13.csv", "Projets")
    
    ' Tronquer le tableau à la taille réelle pour libérer la mémoire
    ReDim Preserve Categories(1 To CategoriesCount)
    
    ' Maintenant que toutes les catégories sont initialisées, on peut charger le dictionnaire
    RagicDictionary.LoadRagicDictionary
    
    Exit Sub
    
ErrorHandler:
    SYS_Logger.Log "category_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de l'initialisation des catégories"
End Sub

' Ajoute une catégorie au tableau des catégories.
' Paramètres :
'   name (String) : Nom de la catégorie
'   filterLevel (String) : Niveau de filtrage principal
'   displayName (String) : Nom d'affichage
'   path (String) : Chemin du fichier source
'   categoryGroup (String) : Groupe de la catégorie
'   secondaryFilterLevel (String, optionnel) : Niveau de filtrage secondaire
'   sheetName (String, optionnel) : Nom de la feuille associée
Public Sub AddCategory(ByVal name As String, ByVal filterLevel As String, ByVal displayName As String, ByVal path As String, ByVal categoryGroup As String, Optional ByVal secondaryFilterLevel As String = "", Optional ByVal sheetName As String = "")
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "AddCategory"
    Const MODULE_NAME As String = "CategoryManager"
    
    If name = "" Or displayName = "" Or path = "" Or categoryGroup = "" Then
        HandleError MODULE_NAME, PROC_NAME, "Paramètres invalides pour l'ajout de catégorie"
        Exit Sub
    End If
    
    CategoriesCount = CategoriesCount + 1
    
    ' Si on dépasse la taille allouée, on logue une erreur car c'est un problème de maintenance
    If CategoriesCount > UBound(Categories) Then
        ReDim Preserve Categories(1 To CategoriesCount) ' Allouer l'espace juste nécessaire en cas de dépassement
        SYS_Logger.Log "category_array_overflow", "Le tableau des catégories a dépassé la taille pré-allouée. Envisager d'augmenter la taille dans InitCategories.", WARNING_LEVEL, PROC_NAME, MODULE_NAME
    End If
    
    Dim idx As Long
    idx = CategoriesCount
    
    With Categories(idx)
        .categoryName = name
        .filterLevel = filterLevel
        .SecondaryFilterLevel = secondaryFilterLevel
        .displayName = displayName
        .URL = env.RAGIC_BASE_URL & path & env.GetRagicApiParams()
        .PowerQueryName = "PQ_" & Utilities.SanitizeTableName(name)
        .categoryGroup = categoryGroup
        If sheetName = "" Then
            .SheetName = displayName
        Else
            .SheetName = sheetName
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    SYS_Logger.Log "category_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de l'ajout de la catégorie: " & name
End Sub

' Retourne l'index d'une catégorie par son nom d'affichage.
' Paramètres :
'   displayName (String) : Nom d'affichage de la catégorie
' Retour :
'   Long (index de la catégorie ou 0 si non trouvée)
Public Function GetCategoryIndexByName(displayName As String) As Long
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetCategoryIndexByName"
    Const MODULE_NAME As String = "CategoryManager"
      If displayName = "" Then
        HandleError MODULE_NAME, PROC_NAME, "Nom d'affichage vide"
        GetCategoryIndexByName = 0
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To CategoriesCount
        If Categories(i).displayName = displayName Then
            GetCategoryIndexByName = i
            Exit Function
        End If
    Next i
    GetCategoryIndexByName = 0 ' Non trouvé
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "category_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la recherche de l'index de la catégorie: " & displayName
    GetCategoryIndexByName = 0
End Function

' Retourne une catégorie par son nom d'affichage.
' Paramètres :
'   displayName (String) : Nom d'affichage de la catégorie
' Retour :
'   CategoryInfo (structure de la catégorie)
Public Function GetCategoryByName(displayName As String) As CategoryInfo
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetCategoryByName"
    Const MODULE_NAME As String = "CategoryManager"
    
    If displayName = "" Then
        HandleError MODULE_NAME, PROC_NAME, "Nom d'affichage vide"
        Exit Function
    End If
    
    ' Assurer l'initialisation
    If CategoriesCount = 0 Then InitCategories
    
    Dim i As Long
    For i = 1 To CategoriesCount
        If LCase(Categories(i).displayName) = LCase(displayName) Then
            GetCategoryByName = Categories(i) ' Pas de 'Set' car CategoryInfo n'est pas un objet
            Exit Function
        End If
    Next i
    
    ' Non trouvé, sortir proprement.
    SYS_Logger.Log "category_not_found", "La catégorie avec le nom d'affichage '" & displayName & "' n'a pas été trouvée.", WARNING_LEVEL, PROC_NAME, MODULE_NAME
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "category_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la récupération de la catégorie: " & displayName
End Function

' Retourne toutes les catégories sous forme de tableau.
' Retour :
'   CategoryInfo() (tableau de CategoryInfo)
Public Function GetAllCategories() As CategoryInfo()
    On Error GoTo ErrorHandler

    Const PROC_NAME As String = "GetAllCategories"
    Const MODULE_NAME As String = "CategoryManager"

    If CategoriesCount = 0 Then
        HandleError MODULE_NAME, PROC_NAME, "Aucune catégorie n'est définie"
        Exit Function
    End If

    ' Créer un nouveau tableau de la bonne taille
    Dim result() As CategoryInfo
    ReDim result(LBound(Categories) To UBound(Categories))

    ' Copier chaque élément (copie profonde)
    Dim i As Long
    For i = LBound(Categories) To UBound(Categories)
        With Categories(i)
            result(i).CategoryName = .CategoryName
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
    SYS_Logger.Log "category_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la récupération de toutes les catégories"
End Function

' Retourne une catégorie par son nom interne (CategoryName).
' Paramètres :
'   name (String) : Nom interne de la catégorie
' Retour :
'   CategoryInfo (structure de la catégorie)
Public Function GetCategoryByCategoryName(name As String) As CategoryInfo
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "GetCategoryByCategoryName"
    Const MODULE_NAME As String = "CategoryManager"
    
    If name = "" Then
        HandleError MODULE_NAME, PROC_NAME, "Nom de catégorie vide"
        Exit Function
    End If
    
    ' Assurer l'initialisation
    If CategoriesCount = 0 Then InitCategories
    
    Dim i As Long
    For i = 1 To CategoriesCount
        If LCase(Categories(i).categoryName) = LCase(name) Then
            GetCategoryByCategoryName = Categories(i) ' Pas de 'Set'
            Exit Function
        End If
    Next i
    
    ' Non trouvé, sortir proprement.
    SYS_Logger.Log "category_not_found", "La catégorie avec le nom interne '" & name & "' n'a pas été trouvée.", WARNING_LEVEL, PROC_NAME, MODULE_NAME
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "category_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la récupération de la catégorie: " & name
End Function
