Attribute VB_Name = "PQDebugTools"
Option Explicit

Public Sub ProcessInjectAllPowerQueries(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessInjectAllPowerQueries"
    Const MODULE_NAME As String = "PQDebugTools"
    On Error GoTo ErrorHandler

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
    Exit Sub

ErrorHandler:
    SYS_Logger.Log "pq_debug_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Failed during PowerQuery injection process."
End Sub

' Efface toutes les requêtes PowerQuery et leurs tableaux associés
Public Sub ProcessCleanupAllPowerQueries(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessCleanupAllPowerQueries"
    Const MODULE_NAME As String = "PQDebugTools"
    On Error GoTo ErrorHandler

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
        Call DataLoaderManager.CleanupPowerQuery(Category.PowerQueryName)
        ' Nettoyage avancé : supprimer la connexion et les QueryTables orphelins
        Dim conn As WorkbookConnection
        On Error Resume Next
        For Each conn In ThisWorkbook.Connections
            If conn.Name = Category.PowerQueryName Then
                conn.Delete
                Exit For
            End If
        Next conn
        On Error GoTo 0
        Dim ws As Worksheet, qt As QueryTable
        For Each ws In ThisWorkbook.Worksheets
            For Each qt In ws.QueryTables
                If qt.CommandText Like "*" & Category.PowerQueryName & "*" Then
                    qt.Delete
                End If
            Next qt
        Next ws
    Next i
    
    MsgBox "Nettoyage terminé", vbInformation
    Exit Sub

ErrorHandler:
    SYS_Logger.Log "pq_debug_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Failed during PowerQuery cleanup process."
End Sub

' Test et debug du RagicDictionary
Public Sub ProcessDebugRagicDictionary(ByVal control As IRibbonControl)
    Const PROC_NAME As String = "ProcessDebugRagicDictionary"
    Const MODULE_NAME As String = "PQDebugTools"
    On Error GoTo ErrorHandler

    Log "debug_ragic", "=== Test du RagicDictionary ===", DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    
    ' 1. Charger le dictionnaire
    Log "debug_ragic", "1. Chargement du dictionnaire...", DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    RagicDictionary.LoadRagicDictionary
    
    ' 2. Vérifier si le dictionnaire a été chargé
    If RagicDictionary.RagicFieldDict Is Nothing Then
        Log "debug_ragic", "ERREUR: Le dictionnaire n'a pas été chargé", ERROR_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
        Exit Sub
    End If
    
    ' 3. Afficher le contenu du dictionnaire
    Log "debug_ragic", "2. Contenu du dictionnaire :", DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    Dim key As Variant
    For Each key In RagicDictionary.RagicFieldDict.Keys
        Log "debug_ragic", "  " & key & " => " & RagicDictionary.RagicFieldDict(key), DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    Next key
    
    ' 4. Tester quelques champs
    Log "debug_ragic", "3. Test de quelques champs :", DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    Call TestField("CO2 Capture", "Brand")
    Call TestField("H2 waters electrolysis", "Specific Electricity Consumption (SEC) [MWhe/kgH2]")
    Call TestField("MeOH - CO2-to-Methanol Synthesis", "CO2 Conversion [%]")
    
    Log "debug_ragic", String(50, "-"), DEBUG_LEVEL, "ProcessDebugRagicDictionary", "PQDebugTools"
    MsgBox "Test du RagicDictionary terminé. Voir la fenêtre de debug pour les détails.", vbInformation
    Exit Sub

ErrorHandler:
    SYS_Logger.Log "pq_debug_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Failed during RagicDictionary debug process."
End Sub

' Fonction utilitaire pour tester un champ
Private Sub TestField(SheetName As String, fieldName As String)
    Const PROC_NAME As String = "TestField"
    Const MODULE_NAME As String = "PQDebugTools"
    On Error GoTo ErrorHandler
    Log "test_field", "Test de " & SheetName & "|" & fieldName & " :", DEBUG_LEVEL, "TestField", "PQDebugTools"
    Log "test_field", "  Hidden = " & RagicDictionary.IsFieldHidden(SheetName, fieldName), DEBUG_LEVEL, "TestField", "PQDebugTools"
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "pq_debug_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du test du champ " & fieldName
End Sub


