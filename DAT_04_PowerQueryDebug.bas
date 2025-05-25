' Module : PQDebugTools
' Module de debug pour injecter et tester les requêtes PowerQuery
Option Explicit
Private Const MODULE_NAME As String = "PQDebugTools"

' Module de debug pour tester les requêtes PowerQuery dans l'éditeur

' Force l'injection et le chargement de toutes les requêtes PowerQuery
Public Sub ProcessInjectAllPowerQueries(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    ' Initialiser les catégories
    CategoryManager.InitCategories
    
    ' Récupérer l'accès direct au tableau de catégories
    Dim categories() As CategoryInfo
    categories = CategoryManager.Categories
    
    ' Compteur pour suivre la progression
    Dim totalCount As Long
    Dim successCount As Long
    Dim failureCount As Long
    totalCount = CategoryManager.CategoriesCount
    
    ' Pour chaque catégorie, injecter la requête
    Dim i As Long
    Dim category As CategoryInfo
    For i = 1 To CategoryManager.CategoriesCount
        category = categories(i)
        
        Debug.Print "=== Traitement de " & category.DisplayName & " ==="
        Debug.Print "URL: " & category.URL
        Debug.Print "Nom de la requête: " & category.PowerQueryName
        
        ' Créer/Mettre à jour la requête PowerQuery dans l'éditeur
        If Not PQQueryManager.EnsurePQQueryExists(category) Then
            Debug.Print "ERREUR: Échec de la création de la requête PowerQuery"
            failureCount = failureCount + 1
            GoTo NextCategory
        End If
        
        Debug.Print "Succès: Requête créée dans l'éditeur Power Query"
        successCount = successCount + 1
        
NextCategory:
        Debug.Print String(50, "-") & vbCrLf
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
    categories = CategoryManager.Categories
    
    ' Pour chaque catégorie, supprimer la requête et le tableau associé
    Dim i As Long
    Dim category As CategoryInfo
    For i = 1 To CategoryManager.CategoriesCount
        category = categories(i)
          Debug.Print "Nettoyage de " & category.PowerQueryName
        DAT_01_DataLoadManager.CleanupPowerQuery category.PowerQueryName
    Next i
    
    MsgBox "Nettoyage terminé", vbInformation
End Sub

' Test et debug du RagicDictionary
Public Sub ProcessDebugRagicDictionary(ByVal control As IRibbonControl, Optional ByRef returnValue As Variant)
    Debug.Print "=== Test du RagicDictionary ==="
    
    ' 1. Charger le dictionnaire
    Debug.Print "1. Chargement du dictionnaire..."
    LoadRagicDictionary
    
    ' 2. Vérifier si le dictionnaire a été chargé
    If RagicFieldDict Is Nothing Then
        Debug.Print "ERREUR: Le dictionnaire n'a pas été chargé"
        Exit Sub
    End If
    
    ' 3. Afficher le contenu du dictionnaire
    Debug.Print "2. Contenu du dictionnaire :"
    Dim key As Variant
    For Each key In RagicFieldDict.Keys
        Debug.Print "  " & key & " => " & RagicFieldDict(key)
    Next key
    
    ' 4. Tester quelques champs
    Debug.Print "3. Test de quelques champs :"
    TestField "CO2 Capture", "Brand"
    TestField "H2 waters electrolysis", "Specific Electricity Consumption (SEC) [MWhe/kgH2]"
    TestField "MeOH - CO2-to-Methanol Synthesis", "CO2 Conversion [%]"
    
    Debug.Print String(50, "-")
    MsgBox "Test du RagicDictionary terminé. Voir la fenêtre de debug pour les détails.", vbInformation
End Sub

' Fonction utilitaire pour tester un champ
Private Sub TestField(sheetName As String, fieldName As String)
    Debug.Print "  Test de " & sheetName & "|" & fieldName & " :"
    Debug.Print "    Hidden = " & IsFieldHidden(sheetName, fieldName)
End Sub

Public Sub PrintQueryMCode(ByVal queryName As String)
    Const PROC_NAME As String = "PrintQueryMCode"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Attempting to print M code for query: " & queryName, PROC_NAME, MODULE_NAME

    Dim pq As Object ' WorkbookQuery
    On Error Resume Next ' To check if query exists
    Set pq = ThisWorkbook.Queries(queryName)
    On Error GoTo ErrorHandler ' Reinstate error handling

    If pq Is Nothing Then
        ' Debug.Print "Query '" & queryName & "' not found."
        ElyseMain_Orchestrator.LogWarning PROC_NAME & "_NotFound", "Query '" & queryName & "' not found.", PROC_NAME, MODULE_NAME
        ElyseMessageBox_System.ShowWarningMessage "M Code Error", "Query '" & queryName & "' not found."
        Exit Sub
    End If

    ' Debug.Print "M Code for query: " & queryName
    ' Debug.Print pq.Formula
    ElyseMain_Orchestrator.LogDebug PROC_NAME & "_MCodeHeader", "M Code for query: " & queryName, PROC_NAME, MODULE_NAME
    ElyseMain_Orchestrator.LogDebug PROC_NAME & "_MCodeBody", pq.Formula, PROC_NAME, MODULE_NAME ' Log M code as detail
    
    ' Optionally, display it if it was the original intent beyond Debug.Print
    ' ElyseMessageBox_System.ShowInfoMessage "M Code: " & queryName, pq.Formula ' This might be too long for a message box
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "M Code for query '" & queryName & "' logged.", PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

