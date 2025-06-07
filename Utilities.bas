Attribute VB_Name = "Utilities"


Private Declare PtrSafe Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" _
    (ByVal NameFormat As Long, ByVal lpNameBuffer As String, ByRef nSize As Long) As Long

Private Const NameUserPrincipal As Long = 8

Public wsPQData As Worksheet
Public Const ADDIN_VERSION_MAJOR As Integer = 1
Public Const ADDIN_VERSION_MINOR As Integer = 0
Public Const ADDIN_VERSION_PATCH As Integer = 0

' --- NOUVELLES ROUTINES DE DÉMARRAGE ASYNCHRONE ---

' Tâche principale de démarrage, appelée de manière asynchrone.
Public Sub RunStartupTasks()
    ' On Error Resume Next ' Empêche tout crash si une tâche de fond échoue.
    ' Const PROC_NAME As String = "RunStartupTasks"
    ' Const MODULE_NAME_STR As String = "Utilities" ' Explicitly define module name for logger

    ' SYS_Logger.Log "startup", "RunStartupTasks: Démarrage des tâches d'arrière-plan...", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    
    ' 1. Préchauffe le moteur Power Query pour accélérer le premier vrai appel
    ' SYS_Logger.Log "startup", "RunStartupTasks: Appel de WarmUpPowerQueryEngine...", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    ' WarmUpPowerQueryEngine
    ' SYS_Logger.Log "startup", "RunStartupTasks: WarmUpPowerQueryEngine terminé.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    
    ' ' 2. Charge le dictionnaire de données en arrière-plan
    ' SYS_Logger.Log "startup", "RunStartupTasks: Appel de LoadRagicDictionary...", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    ' LoadRagicDictionary ' Assurez-vous que cette fonction existe et est accessible
    ' SYS_Logger.Log "startup", "RunStartupTasks: LoadRagicDictionary terminé.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    
    ' 3. Initialise les profils d'accès
    SYS_Logger.Log "startup", "RunStartupTasks: Appel de InitializeDemoProfiles...", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    InitializeDemoProfiles ' Rétablir l'appel
    SYS_Logger.Log "startup", "RunStartupTasks: InitializeDemoProfiles terminé.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    
    ' SYS_Logger.Log "startup", "RunStartupTasks: Tâches d'arrière-plan terminées.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
End Sub

' Crée, rafraîchit et supprime une requête bidon pour forcer le moteur M à s'initialiser.
Private Sub WarmUpPowerQueryEngine()
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "WarmUpPowerQueryEngine"
    Const MODULE_NAME_STR As String = "Utilities"
    Const WARMUP_QUERY_NAME As String = "Internal_WarmUp"
    Dim formula As String
    Dim qry As Object ' WorkbookQuery
    Dim bQueryRefreshedSuccessfully As Boolean
    bQueryRefreshedSuccessfully = False

    SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: Début du préchauffage.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    formula = "let Source = ""Done"" in Source"

    ' Attempt to delete the query if it already exists
    Dim existingQuery As Object
    On Error Resume Next ' Temporarily ignore error if query doesn't exist for this check
    Set existingQuery = ThisWorkbook.Queries(WARMUP_QUERY_NAME)
    On Error GoTo 0 ' Restore broader error handling immediately

    If Not existingQuery Is Nothing Then
        SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: La requête de préchauffage '" & WARMUP_QUERY_NAME & "' existe déjà. Tentative de suppression...", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
        On Error Resume Next ' Handle error specifically during deletion
        existingQuery.Delete
        If Err.Number <> 0 Then
            SYS_Logger.Log "pq_warmup_warn", "WarmUpPowerQueryEngine: Échec de la suppression de la requête existante '" & WARMUP_QUERY_NAME & "'. Erreur: " & Err.Description, WARNING_LEVEL, PROC_NAME, MODULE_NAME_STR
            Err.Clear
            ' Decide if this is critical; for warmup, maybe we can continue if deletion fails.
        Else
            SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: Requête existante '" & WARMUP_QUERY_NAME & "' supprimée.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
        End If
        On Error GoTo ErrorHandler ' Restore main error handler for the sub
    Else
        SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: La requête de préchauffage '" & WARMUP_QUERY_NAME & "' n'existe pas initialement.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    End If
    Set existingQuery = Nothing
    Set qry = Nothing ' Ensure qry is reset before attempting to Add

    ' Create the query
    SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: Création de la requête de préchauffage '" & WARMUP_QUERY_NAME & "'.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    On Error Resume Next ' Temporarily handle error specifically for Add operation
    Set qry = ThisWorkbook.Queries.Add(WARMUP_QUERY_NAME, formula)
    
    If Err.Number <> 0 Or qry Is Nothing Then
        SYS_Logger.Log "pq_warmup_error", "WarmUpPowerQueryEngine: Échec de l'ajout de la requête '" & WARMUP_QUERY_NAME & "'. Erreur: " & Err.Description & " (Num: " & Err.Number & ")", ERROR_LEVEL, PROC_NAME, MODULE_NAME_STR
        Err.Clear
        On Error GoTo ErrorHandler ' Restore main error handler
        GoTo Cleanup ' Skip refresh if query add failed
    Else
        SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: Requête '" & WARMUP_QUERY_NAME & "' ajoutée avec succès.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    End If
    On Error GoTo ErrorHandler ' Restore main error handler for the sub

    ' Refresh the connection
    Dim connectionName As String
    connectionName = "Query - " & WARMUP_QUERY_NAME
    SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: Tentative de rafraîchissement de la connexion '" & connectionName & "'.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    
    On Error Resume Next ' Specific error handling for Refresh operation
    ThisWorkbook.Connections(connectionName).Refresh
    If Err.Number <> 0 Then
        SYS_Logger.Log "pq_warmup_error", "WarmUpPowerQueryEngine: Erreur lors du rafraîchissement de la connexion '" & connectionName & "'. Erreur: " & Err.Description & " (Num: " & Err.Number & ")", ERROR_LEVEL, PROC_NAME, MODULE_NAME_STR
        Err.Clear
        On Error GoTo ErrorHandler ' Restore main error handler
        GoTo Cleanup ' Error occurred during refresh
    Else
        SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: Connexion '" & connectionName & "' rafraîchie avec succès.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
        bQueryRefreshedSuccessfully = True
    End If
    On Error GoTo ErrorHandler ' Restore main error handler for the sub
    
Cleanup:
    SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: Entrée dans la section Cleanup.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    On Error Resume Next ' Ensure cleanup attempts to run fully, errors here are logged but don't stop cleanup
    
    Dim queryToDeleteByName As Object
    Set queryToDeleteByName = ThisWorkbook.Queries(WARMUP_QUERY_NAME) ' Try to get the query by name for deletion
    
    If Not queryToDeleteByName Is Nothing Then
        SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: Suppression de la requête '" & WARMUP_QUERY_NAME & "' depuis Cleanup.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
        queryToDeleteByName.Delete
        If Err.Number <> 0 Then
             SYS_Logger.Log "pq_warmup_warn", "WarmUpPowerQueryEngine: Échec de la suppression de la requête '" & WARMUP_QUERY_NAME & "' pendant le Cleanup. Erreur: " & Err.Description, WARNING_LEVEL, PROC_NAME, MODULE_NAME_STR
             Err.Clear
        Else
            SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: Requête '" & WARMUP_QUERY_NAME & "' supprimée avec succès depuis Cleanup.", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
        End If
    Else
        SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: Requête '" & WARMUP_QUERY_NAME & "' non trouvée pour suppression dans Cleanup (peut-être échec de l'ajout ou déjà supprimée).", DEBUG_LEVEL, PROC_NAME, MODULE_NAME_STR
    End If
    Set queryToDeleteByName = Nothing
    Set qry = Nothing ' Clear the qry object variable as well
    
    On Error GoTo 0 ' Clear any specific error handling within Cleanup

    If bQueryRefreshedSuccessfully Then
        SYS_Logger.Log "pq_warmup", "WarmUpPowerQueryEngine: Préchauffage du moteur Power Query terminé avec succès.", INFO_LEVEL, PROC_NAME, MODULE_NAME_STR
    Else
        SYS_Logger.Log "pq_warmup_warn", "WarmUpPowerQueryEngine: Préchauffage du moteur Power Query terminé avec des erreurs ou n'a pas pu rafraîchir complètement.", WARNING_LEVEL, PROC_NAME, MODULE_NAME_STR
    End If
    Exit Sub

ErrorHandler:
    SYS_Logger.Log "pq_warmup_error", "WarmUpPowerQueryEngine: Erreur VBA non gérée - Num: " & CStr(Err.Number) & ", Desc: " & Err.Description & ", Src: " & Err.Source, ERROR_LEVEL, PROC_NAME, MODULE_NAME_STR
    Resume Cleanup ' Aller au nettoyage même en cas d'erreur
End Sub


' --- FONCTIONS EXISTANTES ---

Public Sub InitializePQData()
    On Error Resume Next
    Set wsPQData = ActiveWorkbook.Worksheets("PQ_DATA")
    On Error GoTo 0
    
    ' Si la feuille n'existe pas, la créer
    If wsPQData Is Nothing Then
        Set wsPQData = ActiveWorkbook.Worksheets.Add
        wsPQData.Name = "PQ_DATA"
    End If
End Sub

Function GetLastColumn(ws As Worksheet) As Long
    GetLastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
End Function

' --- Utility function for smart truncation ---
Function TruncateWithEllipsis(text As String, maxLen As Integer) As String
    If Len(text) > maxLen Then
        TruncateWithEllipsis = Left(text, maxLen - 3) & "..."
    Else
        TruncateWithEllipsis = text
    End If
End Function

' Nettoie une chaîne pour en faire un nom de tableau valide
Public Function SanitizeTableName(ByVal inputName As String) As String
    Dim result As String
    Dim i As Long
    Dim c As String
    
    ' Remplacer les caractères non valides par des underscores
    result = inputName
    
    ' Remplacer les espaces par des underscores
    result = Replace(result, " ", "_")
    
    ' Remplacer les caractères spéciaux courants
    result = Replace(result, "-", "_")
    result = Replace(result, ".", "_")
    result = Replace(result, "(", "")
    result = Replace(result, ")", "")
    result = Replace(result, "/", "_")
    result = Replace(result, "\", "_")
    
    ' Supprimer les accents
    result = RemoveDiacritics(result)
    
    ' Ne garder que les caractères alphanumériques et underscores
    Dim cleanResult As String
    cleanResult = ""
    For i = 1 To Len(result)
        c = Mid(result, i, 1)
        If (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or _
           (c >= "0" And c <= "9") Or c = "_" Then
            cleanResult = cleanResult & c
        End If
    Next i
    
    ' Limiter la longueur à 250 caractères (garder de la marge pour les suffixes)
    If Len(cleanResult) > 250 Then
        cleanResult = Left(cleanResult, 250)
    End If
    
    SanitizeTableName = cleanResult
End Function

' Fonction auxiliaire pour supprimer les accents
Private Function RemoveDiacritics(ByVal text As String) As String
    Dim i As Long
    Const AccentedChars = "àáâãäçèéêëìíîïñòóôõöùúûüýÿÀÁÂÃÄÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜÝ"
    Const UnaccentedChars = "aaaaaceeeeiiiinooooouuuuyyAAAAACEEEEIIIINOOOOOUUUUY"
    
    For i = 1 To Len(AccentedChars)
        text = Replace(text, Mid(AccentedChars, i, 1), Mid(UnaccentedChars, i, 1))
    Next i
    
    RemoveDiacritics = text
End Function


Function GetUserEmail() As String
    Dim buffer As String * 255
    Dim bufferSize As Long
    bufferSize = 255
    If GetUserNameEx(NameUserPrincipal, buffer, bufferSize) <> 0 Then
        GetUserEmail = Left$(buffer, InStr(buffer, Chr$(0)) - 1)
    Else
        GetUserEmail = "Non disponible"
    End If
End Function

Function GetAddinVersion() As String
    GetAddinVersion = ADDIN_VERSION_MAJOR & "." & ADDIN_VERSION_MINOR & "." & ADDIN_VERSION_PATCH
End Function

' Correction du callback pour getSupertip : doit être une Sub avec ByRef supertip
Public Sub GetAddinVersionSupertip(control As IRibbonControl, ByRef supertip)
    supertip = "Utilisateur : " & GetUserEmail() & Chr(10) & _
        "Version addin : " & GetAddinVersion()
End Sub

