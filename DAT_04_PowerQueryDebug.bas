' Module : PQDebugTools
' Module de debug pour injecter et tester les requêtes PowerQuery
Option Explicit
Private Const MODULE_NAME As String = "PQDebugTools"

' Module de debug pour tester les requêtes PowerQuery dans l'éditeur

' Force l'injection et le chargement de toutes les requêtes PowerQuery
Public Sub ProcessInjectAllPowerQueries(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessInjectAllPowerQueries"
    On Error GoTo ErrorHandler
    
    LogInfo PROC_NAME & "_Start", "Début de l'injection de toutes les requêtes PowerQuery", PROC_NAME, MODULE_NAME
    
    ' Initialiser les catégories
    LogDebug PROC_NAME & "_InitCategories", "Initialisation des catégories", PROC_NAME, MODULE_NAME
    CategoryManager.InitCategories
    
    ' Récupérer l'accès direct au tableau de catégories
    Dim categories() As CategoryInfo
    categories = CategoryManager.Categories
    
    ' Compteur pour suivre la progression
    Dim totalCount As Long
    Dim successCount As Long
    Dim failureCount As Long
    totalCount = CategoryManager.CategoriesCount
    
    LogInfo PROC_NAME & "_TotalCategories", "Nombre total de catégories à traiter: " & totalCount, PROC_NAME, MODULE_NAME
    
    ' Pour chaque catégorie, injecter la requête
    Dim i As Long
    Dim category As CategoryInfo
    For i = 1 To CategoryManager.CategoriesCount
        category = categories(i)
        
        LogInfo PROC_NAME & "_ProcessingCategory", "Traitement de la catégorie: " & category.DisplayName, PROC_NAME, MODULE_NAME
        LogDebug PROC_NAME & "_CategoryDetails", "URL: " & category.URL & ", Nom de la requête: " & category.PowerQueryName, PROC_NAME, MODULE_NAME
        
        ' Créer/Mettre à jour la requête PowerQuery dans l'éditeur
        If Not PQQueryManager.EnsurePQQueryExists(category) Then
            LogError PROC_NAME & "_QueryCreationFailed", 0, "Échec de la création de la requête PowerQuery pour " & category.DisplayName, PROC_NAME, MODULE_NAME
            failureCount = failureCount + 1
            GoTo NextCategory
        End If
        
        LogInfo PROC_NAME & "_QueryCreated", "Requête créée avec succès dans l'éditeur Power Query: " & category.PowerQueryName, PROC_NAME, MODULE_NAME
        successCount = successCount + 1
        
NextCategory:
        LogDebug PROC_NAME & "_CategoryComplete", "Traitement de la catégorie terminé: " & category.DisplayName, PROC_NAME, MODULE_NAME
    Next i
    
    ' Afficher le résumé
    Dim summaryMsg As String
    summaryMsg = "Traitement terminé" & vbCrLf & _
                 "Total: " & totalCount & vbCrLf & _
                 "Succès: " & successCount & vbCrLf & _
                 "Échecs: " & failureCount
    
    LogInfo PROC_NAME & "_Summary", "Résumé du traitement - Total: " & totalCount & ", Succès: " & successCount & ", Échecs: " & failureCount, PROC_NAME, MODULE_NAME
    ShowInfoMessage "Injection PowerQuery", summaryMsg
    
    LogInfo PROC_NAME & "_End", "Fin de l'injection des requêtes PowerQuery", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Erreur lors de l'injection des requêtes PowerQuery: " & Err.Description, PROC_NAME, MODULE_NAME
    ShowErrorMessage "Erreur d'injection", "Une erreur est survenue lors de l'injection des requêtes PowerQuery: " & vbCrLf & Err.Description
End Sub

' Efface toutes les requêtes PowerQuery et leurs tableaux associés
Public Sub ProcessCleanupAllPowerQueries(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessCleanupAllPowerQueries"
    On Error GoTo ErrorHandler
    
    LogInfo PROC_NAME & "_Start", "Début du nettoyage de toutes les requêtes PowerQuery", PROC_NAME, MODULE_NAME
    
    ' Initialiser les catégories
    LogDebug PROC_NAME & "_InitCategories", "Initialisation des catégories", PROC_NAME, MODULE_NAME
    CategoryManager.InitCategories
    
    ' Récupérer l'accès direct au tableau de catégories
    Dim categories() As CategoryInfo
    categories = CategoryManager.Categories
    
    ' Compteur pour suivre la progression
    Dim totalCount As Long
    Dim successCount As Long
    Dim failureCount As Long
    totalCount = CategoryManager.CategoriesCount
    
    LogInfo PROC_NAME & "_TotalCategories", "Nombre total de catégories à nettoyer: " & totalCount, PROC_NAME, MODULE_NAME
    
    ' Pour chaque catégorie, supprimer la requête et le tableau associé
    Dim i As Long
    Dim category As CategoryInfo
    For i = 1 To CategoryManager.CategoriesCount
        category = categories(i)
        
        LogInfo PROC_NAME & "_ProcessingCategory", "Nettoyage de la catégorie: " & category.DisplayName, PROC_NAME, MODULE_NAME
        LogDebug PROC_NAME & "_CategoryDetails", "Nom de la requête: " & category.PowerQueryName, PROC_NAME, MODULE_NAME
        
        On Error Resume Next
        DAT_01_DataLoadManager.CleanupPowerQuery category.PowerQueryName
        If Err.Number <> 0 Then
            LogError PROC_NAME & "_CleanupFailed", Err.Number, "Échec du nettoyage de la requête " & category.PowerQueryName & ": " & Err.Description, PROC_NAME, MODULE_NAME
            failureCount = failureCount + 1
        Else
            LogInfo PROC_NAME & "_CleanupSuccess", "Nettoyage réussi pour la requête: " & category.PowerQueryName, PROC_NAME, MODULE_NAME
            successCount = successCount + 1
        End If
        On Error GoTo ErrorHandler
        
        LogDebug PROC_NAME & "_CategoryComplete", "Nettoyage de la catégorie terminé: " & category.DisplayName, PROC_NAME, MODULE_NAME
    Next i
    
    ' Afficher le résumé
    Dim summaryMsg As String
    summaryMsg = "Nettoyage terminé" & vbCrLf & _
                 "Total: " & totalCount & vbCrLf & _
                 "Succès: " & successCount & vbCrLf & _
                 "Échecs: " & failureCount
    
    LogInfo PROC_NAME & "_Summary", "Résumé du nettoyage - Total: " & totalCount & ", Succès: " & successCount & ", Échecs: " & failureCount, PROC_NAME, MODULE_NAME
    ShowInfoMessage "Nettoyage PowerQuery", summaryMsg
    
    LogInfo PROC_NAME & "_End", "Fin du nettoyage des requêtes PowerQuery", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Erreur lors du nettoyage des requêtes PowerQuery: " & Err.Description, PROC_NAME, MODULE_NAME
    ShowErrorMessage "Erreur de nettoyage", "Une erreur est survenue lors du nettoyage des requêtes PowerQuery: " & vbCrLf & Err.Description
End Sub

' Test et debug du RagicDictionary
Public Sub ProcessDebugRagicDictionary(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Const PROC_NAME As String = "ProcessDebugRagicDictionary"
    On Error GoTo ErrorHandler
    
    LogInfo PROC_NAME & "_Start", "Début du test du RagicDictionary", PROC_NAME, MODULE_NAME
    
    ' 1. Charger le dictionnaire
    LogDebug PROC_NAME & "_LoadDictionary", "Chargement du dictionnaire...", PROC_NAME, MODULE_NAME
    LoadRagicDictionary
    
    ' 2. Vérifier si le dictionnaire a été chargé
    If RagicFieldDict Is Nothing Then
        LogError PROC_NAME & "_LoadFailed", 0, "Le dictionnaire n'a pas été chargé", PROC_NAME, MODULE_NAME
        ShowErrorMessage "Erreur de test", "Le dictionnaire n'a pas été chargé"
        Exit Sub
    End If
    
    ' 3. Afficher le contenu du dictionnaire
    LogInfo PROC_NAME & "_Content", "Contenu du dictionnaire:", PROC_NAME, MODULE_NAME
    Dim key As Variant
    For Each key In RagicFieldDict.Keys
        LogDebug PROC_NAME & "_DictionaryItem", key & " => " & RagicFieldDict(key), PROC_NAME, MODULE_NAME
    Next key
    
    ' 4. Tester quelques champs
    LogInfo PROC_NAME & "_TestFields", "Test de champs spécifiques", PROC_NAME, MODULE_NAME
    TestField "CO2 Capture", "Brand"
    TestField "H2 waters electrolysis", "Specific Electricity Consumption (SEC) [MWhe/kgH2]"
    TestField "MeOH - CO2-to-Methanol Synthesis", "CO2 Conversion [%]"
    
    LogInfo PROC_NAME & "_End", "Test du RagicDictionary terminé", PROC_NAME, MODULE_NAME
    ShowInfoMessage "Test terminé", "Test du RagicDictionary terminé. Voir les logs pour les détails."
    Exit Sub

ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Erreur lors du test du RagicDictionary: " & Err.Description, PROC_NAME, MODULE_NAME
    ShowErrorMessage "Erreur de test", "Une erreur est survenue lors du test du RagicDictionary: " & vbCrLf & Err.Description
End Sub

' Fonction utilitaire pour tester un champ
Private Sub TestField(sheetName As String, fieldName As String)
    Const PROC_NAME As String = "TestField"
    On Error GoTo ErrorHandler
    
    LogDebug PROC_NAME & "_Start", "Test du champ: " & sheetName & "|" & fieldName, PROC_NAME, MODULE_NAME
    
    Dim isHidden As Boolean
    isHidden = IsFieldHidden(sheetName, fieldName)
    
    LogDebug PROC_NAME & "_Result", "Résultat du test - Hidden = " & isHidden, PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    LogError PROC_NAME & "_Error", Err.Number, "Erreur lors du test du champ " & sheetName & "|" & fieldName & ": " & Err.Description, PROC_NAME, MODULE_NAME
End Sub

Public Sub PrintQueryMCode(ByVal queryName As String)
    Const PROC_NAME As String = "PrintQueryMCode"
    On Error GoTo ErrorHandler

    LogInfo PROC_NAME & "_Start", "Attempting to print M code for query: " & queryName, PROC_NAME, MODULE_NAME

    Dim pq As Object ' WorkbookQuery
    On Error Resume Next ' To check if query exists
    Set pq = ThisWorkbook.Queries(queryName)
    On Error GoTo ErrorHandler ' Reinstate error handling

    If pq Is Nothing Then
        ' Debug.Print "Query '" & queryName & "' not found."
        LogWarning PROC_NAME & "_NotFound", "Query '" & queryName & "' not found.", PROC_NAME, MODULE_NAME
        ShowWarningMessage "M Code Error", "Query '" & queryName & "' not found."
        Exit Sub
    End If

    ' Debug.Print "M Code for query: " & queryName
    ' Debug.Print pq.Formula
    LogDebug PROC_NAME & "_MCodeHeader", "M Code for query: " & queryName, PROC_NAME, MODULE_NAME
    LogDebug PROC_NAME & "_MCodeBody", pq.Formula, PROC_NAME, MODULE_NAME ' Log M code as detail
    
    ' Commented out as it might be too long for a message box
    ' ShowInfoMessage "M Code: " & queryName, pq.Formula
    
    LogInfo PROC_NAME & "_End", "M Code for query '" & queryName & "' logged.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME
End Sub

