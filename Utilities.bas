Attribute VB_Name = "Utilities"
Option Explicit

' --- Déclarations API Windows ---
#If VBA7 Then
    Private Declare PtrSafe Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" (ByVal NameFormat As Long, ByVal lpNameBuffer As String, ByRef lpnSize As Long) As Long
#Else
    Private Declare Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" (ByVal NameFormat As Long, ByVal lpNameBuffer As String, ByRef lpnSize As Long) As Long
#End If

Private Const NAME_USER_PRINCIPAL As Long = 8

' --- Constantes du module ---
Public Const ADDIN_VERSION_MAJOR As Integer = 1
Public Const ADDIN_VERSION_MINOR As Integer = 0
Public Const ADDIN_VERSION_PATCH As Integer = 0
Private Const MODULE_NAME_STR As String = "Utilities"

' Constantes pour la sanitization des noms
Private Const ACCENTED_CHARS As String = "àáâãäçèéêëìíîïñòóôõöùúûüýÿÀÁÂÃÄÇÈÉÊËÌÍÎÏÑÒÓÔÕÖÙÚÛÜÝ"
Private Const UNACCENTED_CHARS As String = "aaaaaceeeeiiiinooooouuuuyyAAAAACEEEEIIIINOOOOOUUUUY"

' --- Variables de module ---
Public wsPQData As Worksheet

' --- NOUVELLES ROUTINES DE DÉMARRAGE ASYNCHRONE ---

' Tâche principale de démarrage, appelée de manière asynchrone.
Public Sub RunStartupTasks()
    Const PROC_NAME As String = "RunStartupTasks"
    Const MODULE_NAME As String = "Utilities"
    On Error GoTo ErrorHandler
    
    SYS_Logger.Log "startup", "RunStartupTasks: Démarrage des tâches d'arrière-plan...", INFO_LEVEL, PROC_NAME, MODULE_NAME
    
    ' 1. Préchauffe le moteur Power Query pour accélérer le premier vrai appel
    WarmUpPowerQueryEngine
    
    ' 2. Initialise les profils d'accès
    InitializeDemoProfiles
    
    SYS_Logger.Log "startup", "RunStartupTasks: Tâches d'arrière-plan terminées.", INFO_LEVEL, PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Une erreur critique est survenue lors du démarrage des tâches de fond."
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
    Const PROC_NAME As String = "InitializePQData"
    Const MODULE_NAME As String = "Utilities"
    On Error GoTo ErrorHandler

    Set wsPQData = ActiveWorkbook.Worksheets("PQ_DATA")
    
    ' Si la feuille n'existe pas, la créer
    If wsPQData Is Nothing Then
        Set wsPQData = ActiveWorkbook.Worksheets.Add
        wsPQData.Name = "PQ_DATA"
    End If
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Impossible de créer ou de trouver la feuille PQ_DATA."
End Sub

Function GetLastColumn(ws As Worksheet) As Long
    On Error GoTo ErrorHandler
    GetLastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    Exit Function
ErrorHandler:
    GetLastColumn = 1 ' Retourner 1 en cas d'erreur
End Function

' --- Utility function for smart truncation ---
Function TruncateWithEllipsis(text As String, maxLen As Integer) As String
    On Error GoTo ErrorHandler
    If Len(text) > maxLen Then
        TruncateWithEllipsis = Left(text, maxLen - 3) & "..."
    Else
        TruncateWithEllipsis = text
    End If
    Exit Function
ErrorHandler:
    TruncateWithEllipsis = Left(text, maxLen)
End Function

' Fonction auxiliaire pour supprimer les accents
' Source: https://stackoverflow.com/questions/11253628/vba-equivalent-of-excels-clean-and-trim-functions
Private Function RemoveDiacritics(ByVal text As String) As String
    On Error GoTo ErrorHandler
    Dim i As Long
    For i = 1 To Len(ACCENTED_CHARS)
        text = Replace(text, Mid(ACCENTED_CHARS, i, 1), Mid(UNACCENTED_CHARS, i, 1))
    Next i
    RemoveDiacritics = text
    Exit Function
ErrorHandler:
    SYS_Logger.Log "diacritics_error", "Erreur dans RemoveDiacritics", ERROR_LEVEL, "RemoveDiacritics", MODULE_NAME_STR
    RemoveDiacritics = text ' Retourner le texte original en cas d'erreur
End Function

' Nettoie une chaîne pour en faire un nom de tableau valide
Public Function SanitizeTableName(ByVal inputName As String) As String
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "SanitizeTableName"
    
    Dim i As Long
    Dim result As String
    
    ' Le nom de la fonction RemoveDiacritics est déjà explicite
    result = RemoveDiacritics(LCase(inputName))
    
    ' Remplacer les caractères non autorisés par des underscores
    result = Replace(result, " ", "_")
    result = Replace(result, "-", "_")
    result = Replace(result, ".", "_")
    result = Replace(result, "(", "")
    result = Replace(result, ")", "")
    result = Replace(result, "/", "_")
    result = Replace(result, "\", "_")
    
    ' Ne garder que les caractères alphanumériques et underscores
    Dim cleanResult As String
    cleanResult = ""
    For i = 1 To Len(result)
        Dim c As String
        c = Mid(result, i, 1)
        If (c >= "a" And c <= "z") Or (c >= "0" And c <= "9") Or c = "_" Then
            cleanResult = cleanResult & c
        End If
    Next i
    
    ' Supprimer les underscores multiples
    Do While InStr(cleanResult, "__")
        cleanResult = Replace(cleanResult, "__", "_")
    Loop
    
    ' Limiter la longueur à 250 caractères (garder de la marge pour les suffixes)
    If Len(cleanResult) > 250 Then
        cleanResult = Left(cleanResult, 250)
    End If
    
    SanitizeTableName = cleanResult
    Exit Function

ErrorHandler:
    SYS_Logger.Log "sanitize_error", "Erreur lors de la sanitization du nom de table - Num: " & CStr(Err.Number) & ", Desc: " & Err.Description & ", Src: " & Err.Source, ERROR_LEVEL, PROC_NAME, MODULE_NAME_STR
    SanitizeTableName = "SanitizeTableName_Error"
End Function

Function GetUserEmail() As String
    Const PROC_NAME As String = "GetUserEmail"
    Const MODULE_NAME As String = "Utilities"
    On Error GoTo ErrorHandler

    Dim buffer As String * 255
    Dim bufferSize As Long
    bufferSize = 255
    If GetUserNameEx(NAME_USER_PRINCIPAL, buffer, bufferSize) <> 0 Then
        GetUserEmail = Left$(buffer, InStr(buffer, Chr$(0)) - 1)
    Else
        GetUserEmail = "Non disponible"
    End If
    Exit Function

ErrorHandler:
    GetUserEmail = "Erreur"
    HandleError MODULE_NAME, PROC_NAME, "Impossible de récupérer l'email de l'utilisateur."
    GetUserEmail = Environ("USERNAME") ' Fallback
End Function

Public Function GetUserUPN() As String
    On Error GoTo ErrorHandler
    Dim buffer As String
    Dim size As Long
    
    size = 255
    buffer = String(size, vbNullChar)
    
    If GetUserNameEx(NAME_USER_PRINCIPAL, buffer, size) <> 0 Then
        GetUserUPN = Left(buffer, size - 1)
    Else
        GetUserUPN = ""
    End If
    Exit Function
ErrorHandler:
    GetUserUPN = ""
End Function

Public Function GetAddinVersion() As String
    On Error GoTo ErrorHandler
    GetAddinVersion = "v" & ADDIN_VERSION_MAJOR & "." & ADDIN_VERSION_MINOR & "." & ADDIN_VERSION_PATCH
    Exit Function
ErrorHandler:
    GetAddinVersion = "v.Error"
End Function

' Correction du callback pour getSupertip : doit être une Sub avec ByRef supertip
Public Function GetAddinVersionSupertip(ByVal control As IRibbonControl, ByRef supertip As Variant)
    On Error GoTo ErrorHandler
    supertip = "Version " & GetAddinVersion() & " of the Elyse Energy Add-in."
    Exit Function
ErrorHandler:
    supertip = "Could not retrieve version information."
End Function

