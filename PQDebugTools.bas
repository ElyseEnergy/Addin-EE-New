Attribute VB_Name = "PQDebugTools"
' Module : PQDebugTools
' Module de debug pour injecter et tester les requ�tes PowerQuery
Option Explicit

' Module de debug pour tester les requ�tes PowerQuery dans l'�diteur

' Force l'injection et le chargement de toutes les requ�tes PowerQuery
Public Sub ProcessInjectAllPowerQueries(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ' Initialiser les cat�gories
    CategoryManager.InitCategories
    
    ' R�cup�rer l'acc�s direct au tableau de cat�gories
    Dim categories() As CategoryInfo
    categories = CategoryManager.categories
    
    ' Compteur pour suivre la progression
    Dim totalCount As Long
    Dim successCount As Long
    Dim failureCount As Long
    totalCount = CategoryManager.CategoriesCount
    
    ' Pour chaque cat�gorie, injecter la requ�te
    Dim i As Long
    Dim Category As CategoryInfo
    For i = 1 To CategoryManager.CategoriesCount
        Category = categories(i)
          Log "process_category", "=== Traitement de " & Category.DisplayName & " ===", DEBUG_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
        Log "process_category", "URL: " & Category.URL, DEBUG_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
        Log "process_category", "Nom de la requ�te: " & Category.PowerQueryName, DEBUG_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
        
        ' Cr�er/Mettre � jour la requ�te PowerQuery dans l'�diteur
        If Not PQQueryManager.EnsurePQQueryExists(Category) Then
            Log "process_category", "ERREUR: �chec de la cr�ation de la requ�te PowerQuery", ERROR_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
            failureCount = failureCount + 1
            GoTo NextCategory
        End If
        
        Log "process_category", "Succ�s: Requ�te cr��e dans l'�diteur Power Query", DEBUG_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
        successCount = successCount + 1
        
NextCategory:
        Log "process_category", String(50, "-"), DEBUG_LEVEL, "ProcessInjectAllPowerQueries", "PQDebugTools"
    Next i
    
    ' Afficher le r�sum�
    MsgBox "Traitement termin�" & vbCrLf & _
           "Total: " & totalCount & vbCrLf & _
           "Succ�s: " & successCount & vbCrLf & _
           "�checs: " & failureCount, _
           vbInformation, "Injection PowerQuery"
End Sub

' Efface toutes les requ�tes PowerQuery et leurs tableaux associ�s
Public Sub ProcessCleanupAllPowerQueries(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ' Initialiser les cat�gories
    CategoryManager.InitCategories
    
    ' R�cup�rer l'acc�s direct au tableau de cat�gories
    Dim categories() As CategoryInfo
    categories = CategoryManager.categories
    
    ' Pour chaque cat�gorie, supprimer la requ�te et le tableau associ�
    Dim i As Long
    Dim Category As CategoryInfo
    For i = 1 To CategoryManager.CategoriesCount
        Category = categories(i)
        
        Log "cleanup_pq", "Nettoyage de " & Category.PowerQueryName, DEBUG_LEVEL, "ProcessCleanupAllPowerQueries", "PQDebugTools"
        DataLoaderManager.CleanupPowerQuery Category.PowerQueryName
    Next i
    
    MsgBox "Nettoyage termin�", vbInformation
End Sub

' Test et debug du RagicDictionary
Public Sub ProcessDebugRagicDictionary(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Log "debug_ragic", "=== Test du RagicDictionary ===", DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    
    ' 1. Charger le dictionnaire
    Log "debug_ragic", "1. Chargement du dictionnaire...", DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    LoadRagicDictionary
    
    ' 2. V�rifier si le dictionnaire a �t� charg�
    If RagicFieldDict Is Nothing Then
        Log "debug_ragic", "ERREUR: Le dictionnaire n'a pas �t� charg�", ERROR_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
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
    MsgBox "Test du RagicDictionary termin�. Voir la fen�tre de debug pour les d�tails.", vbInformation
End Sub

' Fonction utilitaire pour tester un champ
Private Sub TestField(SheetName As String, fieldName As String)
    Log "test_field", "Test de " & SheetName & "|" & fieldName & " :", DEBUG_LEVEL, "TestField", "PQDebugTools"
    Log "test_field", "  Hidden = " & IsFieldHidden(SheetName, fieldName), DEBUG_LEVEL, "TestField", "PQDebugTools"
End Sub


