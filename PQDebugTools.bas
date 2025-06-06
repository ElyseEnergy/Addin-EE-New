Attribute VB_Name = "PQDebugTools"
' Module : PQDebugTools
' Module de debug pour injecter et tester les requêtes PowerQuery
Option Explicit

' Module de debug pour tester les requêtes PowerQuery dans l'éditeur

' Force l'injection et le chargement de toutes les requêtes PowerQuery
Public Sub ProcessInjectAllPowerQueries(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ' Initialiser les catégories
    CategoryManager.InitCategories
    
    ' Récupérer l'accès direct au tableau de catégories
    Dim categories() As CategoryInfo
    categories = CategoryManager.categories
    
    ' Compteur pour suivre la progression
    Dim totalCount As Long
    Dim successCount As Long
    Dim failureCount As Long
    totalCount = CategoryManager.CategoriesCount
    
    ' Pour chaque catégorie, injecter la requête
    Dim i As Long
    Dim Category As CategoryInfo
    For i = 1 To CategoryManager.CategoriesCount
        Category = categories(i)
          Log "process_category", "=== Traitement de " & Category.DisplayName & " ===", DEBUG_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
        Log "process_category", "URL: " & Category.URL, DEBUG_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
        Log "process_category", "Nom de la requête: " & Category.PowerQueryName, DEBUG_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
        
        ' Créer/Mettre à jour la requête PowerQuery dans l'éditeur
        If Not PQQueryManager.EnsurePQQueryExists(Category) Then
            Log "process_category", "ERREUR: Échec de la création de la requête PowerQuery", ERROR_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
            failureCount = failureCount + 1
            GoTo NextCategory
        End If
        
        Log "process_category", "Succès: Requête créée dans l'éditeur Power Query", DEBUG_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
        successCount = successCount + 1
        
NextCategory:
        Log "process_category", String(50, "-"), DEBUG_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
    Next i
    
    ' Afficher le résumé
    MsgBox "Traitement terminé" & vbCrLf & _
           "Total: " & totalCount & vbCrLf & _
           "Succès: " & successCount & vbCrLf & _
           "Échecs: " & failureCount, _
           vbInformation, "Injection PowerQuery"
End Sub

' Efface toutes les requêtes PowerQuery et leurs tableaux associés
Public Sub ProcessCleanupAllPowerQueries(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ' Initialiser les catégories
    CategoryManager.InitCategories
    
    ' Récupérer l'accès direct au tableau de catégories
    Dim categories() As CategoryInfo
    categories = CategoryManager.categories
    
    ' Pour chaque catégorie, supprimer la requête et le tableau associé
    Dim i As Long
    Dim Category As CategoryInfo
    For i = 1 To CategoryManager.CategoriesCount
        Category = categories(i)
        
        Log "cleanup_pq", "Nettoyage de " & Category.PowerQueryName, DEBUG_LEVEL, "ProcessCleanupAllPowerQueries", "PQDebugTools"
        DataLoaderManager.CleanupPowerQuery Category.PowerQueryName
    Next i
    
    MsgBox "Nettoyage terminé", vbInformation
End Sub

' Test et debug du RagicDictionary
Public Sub ProcessDebugRagicDictionary(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Log "debug_ragic", "=== Test du RagicDictionary ===", DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    
    ' 1. Charger le dictionnaire
    Log "debug_ragic", "1. Chargement du dictionnaire...", DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    LoadRagicDictionary
    
    ' 2. Vérifier si le dictionnaire a été chargé
    If RagicFieldDict Is Nothing Then
        Log "debug_ragic", "ERREUR: Le dictionnaire n'a pas été chargé", ERROR_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
        Exit Sub
    End If
    
    ' 3. Afficher le contenu du dictionnaire
    Log "debug_ragic", "2. Contenu du dictionnaire :", DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    Dim key As Variant
    For Each key In RagicFieldDict.Keys
        Log "debug_ragic", "  " & key & " => " & RagicFieldDict(key), DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    Next key
    
    ' 4. Tester quelques champs
    Log "debug_ragic", "3. Test de quelques champs :", DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    TestField "CO2 Capture", "Brand"
    TestField "H2 waters electrolysis", "Specific Electricity Consumption (SEC) [MWhe/kgH2]"
    TestField "MeOH - CO2-to-Methanol Synthesis", "CO2 Conversion [%]"
    
    Log "debug_ragic", String(50, "-"), DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    MsgBox "Test du RagicDictionary terminé. Voir la fenêtre de debug pour les détails.", vbInformation
End Sub

' Fonction utilitaire pour tester un champ
Private Sub TestField(SheetName As String, fieldName As String)
    Log "test_field", "Test de " & SheetName & "|" & fieldName & " :", DEBUG_LEVEL, "TestField", "PQDebugTools"
    Log "test_field", "  Hidden = " & IsFieldHidden(SheetName, fieldName), DEBUG_LEVEL, "TestField", "PQDebugTools"
End Sub


