Attribute VB_Name = "SYS_Logger"
Option Explicit

' ============================================================================
' SYS_Logger - Système de Logging Simplifié
' ============================================================================

' Chemin du dossier de logs, relatif au classeur de l'addin
Private Function GetLogFolderPath() As String
    On Error Resume Next
    GetLogFolderPath = ThisWorkbook.Path & "\logs"
    If Err.Number <> 0 Then
        ' Fallback sur le dossier temporaire si le chemin n'est pas disponible (par exemple, fichier non sauvegardé)
        GetLogFolderPath = Environ("TEMP") & "\ElyseEnergyAddinLogs"
    End If
End Function

Private Const LOG_FILE_NAME As String = "elyse_energy.log"

' Niveaux de log
Public Enum LogLevel
    DEBUG_LEVEL = 0
    INFO_LEVEL = 1
    WARNING_LEVEL = 2
    ERROR_LEVEL = 3
    CRITICAL_LEVEL = 4
End Enum

' --- NOUVEAU : Paramètres pour le logging vers Ragic ---
Private Const ENABLE_RAGIC_LOGGING As Boolean = True
Private Const RAGIC_LOG_URL As String = "https://ragic.elyse.energy/default/excel-addin/1"
Private Const RAGIC_FIELD_ID_EMAIL As String = "1005232"
Private Const RAGIC_FIELD_ID_LOG As String = "1005233"
' ---------------------------------------------------------

' Variables du module
Private mCurrentLogLevel As LogLevel

' ============================================================================
' INITIALISATION
' ============================================================================

' Nettoyer le log au lancement, puis append pour chaque message
Public Sub InitializeLogger()
    Const PROC_NAME As String = "InitializeLogger"
    Const MODULE_NAME As String = "SYS_Logger"
    On Error GoTo ErrorHandler

    ' Initialisation du niveau de log par défaut
    mCurrentLogLevel = DEBUG_LEVEL
    
    ' S'assurer que le dossier de logs existe
    EnsureLogFolderExists
    
    ' Nettoyer le fichier de log au lancement
    Dim fso As Object, logFilePath As String, logFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    logFilePath = GetLogFolderPath() & "\" & LOG_FILE_NAME
    Set logFile = fso.OpenTextFile(logFilePath, 2, True, -1) ' 2 = ForWriting, efface tout
    logFile.Close
    
    ' Log d'initialisation
    Log "sys_init", "Système de logging initialisé (niveau DEBUG)", INFO_LEVEL, "InitializeLogger", "SYS_Logger"
    Exit Sub

ErrorHandler:
    ' Ne peut pas utiliser le logger ici, il n'est pas encore initialisé
    Debug.Print "CRITICAL ERROR in " & MODULE_NAME & "." & PROC_NAME & ": " & Err.Description
End Sub

' ============================================================================
' FONCTION DE LOGGING
' ============================================================================

' Assure que le dossier de logs existe
Private Sub EnsureLogFolderExists()
    Const PROC_NAME As String = "EnsureLogFolderExists"
    On Error GoTo ErrorHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(GetLogFolderPath()) Then
        fso.CreateFolder GetLogFolderPath()
    End If
    Exit Sub
ErrorHandler:
    Debug.Print "CRITICAL: Could not create log folder in " & PROC_NAME & ": " & Err.Description
End Sub

' Écrit dans le fichier de log (append)
Private Sub WriteToLogFile(logMessage As String)
    Const PROC_NAME As String = "WriteToLogFile"
    On Error GoTo ErrorHandler
    Dim fso As Object
    Dim logFile As Object
    Dim logFilePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    logFilePath = GetLogFolderPath() & "\" & LOG_FILE_NAME
    
    ' S'assurer que le dossier existe
    EnsureLogFolderExists
    
    ' Ouvrir le fichier en mode append (8 = ForAppending, -1 = TristateMixed pour Unicode)
    Set logFile = fso.OpenTextFile(logFilePath, 8, True, -1)
    
    ' Écrire le message (ajoute à la fin du fichier)
    logFile.WriteLine logMessage
    
    ' Fermer le fichier
    logFile.Close
    Exit Sub
ErrorHandler:
    Debug.Print "CRITICAL: Could not write to log file in " & PROC_NAME & ": " & Err.Description
End Sub

Public Sub Log(actionCode As String, message As String, level As LogLevel, _
    Optional ByVal procedureName As String = "", Optional ByVal moduleName As String = "")
    Const PROC_NAME As String = "Log"
    Const MODULE_NAME As String = "SYS_Logger"
    On Error GoTo ErrorHandler
    
    If level < mCurrentLogLevel Then Exit Sub
    
    Dim logMessage As String
    Dim levelString As String
    Dim timeStamp As String
    
    ' Déterminer le niveau de log
    Select Case level
        Case DEBUG_LEVEL
            levelString = "DEBUG"
        Case INFO_LEVEL
            levelString = "INFO "
        Case WARNING_LEVEL
            levelString = "WARN "
        Case ERROR_LEVEL
            levelString = "ERROR"
        Case CRITICAL_LEVEL
            levelString = "CRITICAL"
        Case Else
            levelString = "INFO "
    End Select
    
    ' Format de la date et heure
    timeStamp = "[" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "] "
    
    ' Construire le message
    logMessage = timeStamp & levelString & " [" & actionCode & "] "
    
    ' Ajouter le contexte (module.procedure)
    If moduleName <> "" And procedureName <> "" Then
        logMessage = logMessage & moduleName & "." & procedureName
    ElseIf procedureName <> "" Then
        logMessage = logMessage & procedureName
    End If
    
    ' Ajouter le message
    logMessage = logMessage & ": " & message
    
    ' Afficher dans Immediate Window
    Debug.Print logMessage
    
    ' Écrire dans le fichier de log
    WriteToLogFile logMessage
    
    ' === NOUVEAU: Logging vers Ragic pour les avertissements et erreurs ===
    If ENABLE_RAGIC_LOGGING And level >= WARNING_LEVEL Then
        On Error Resume Next ' "Fire and forget" pour ne pas bloquer l'utilisateur
        LogToRagic logMessage
        On Error GoTo 0
    End If
    Exit Sub

ErrorHandler:
    ' En cas d'erreur dans le logger lui-même, on ne peut que faire un Debug.Print
    Debug.Print "CRITICAL: The logger itself has failed. Error in " & PROC_NAME & ": " & Err.Description
End Sub

' ============================================================================
' NOUVEAU : FONCTIONS POUR LE LOGGING VERS RAGIC
' ============================================================================

' Échappe une chaîne de caractères pour être valide dans un JSON.
Private Function JsonEscape(ByVal text As String) As String
    Const PROC_NAME As String = "JsonEscape"
    On Error GoTo ErrorHandler
    
    ' Remplacer les backslashes
    text = Replace(text, "\", "\\")
    ' Remplacer les guillemets
    text = Replace(text, """", "\""")
    
    JsonEscape = text
    Exit Function
    
ErrorHandler:
    ' IMPORTANT : Ne PAS utiliser Log ou HandleError ici pour éviter une boucle infinie.
    Debug.Print "CRITICAL: Error in " & PROC_NAME & ": " & Err.Description
    JsonEscape = "JSON_ESCAPE_ERROR"
End Function

' Envoie le message de log formaté à la base de données Ragic.
Private Sub LogToRagic(ByVal logMessage As String)
    Dim http As Object
    Dim ragicUrl As String
    Dim jsonPayload As String
    Dim userEmail As String
    
    ' Créer l'objet HTTP. Tente la version 6.0, puis une version de base.
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    If http Is Nothing Then Exit Sub ' Ne pas continuer si l'objet HTTP ne peut être créé
    On Error GoTo 0

    ' Récupérer l'email de l'utilisateur depuis le module Utilities
    userEmail = Utilities.GetUserEmail()

    ' Construire le payload JSON
    jsonPayload = "{" & _
        """" & RAGIC_FIELD_ID_EMAIL & """: """ & JsonEscape(userEmail) & """, " & _
        """" & RAGIC_FIELD_ID_LOG & """: """ & JsonEscape(logMessage) & """" & _
    "}"

    ' Construire l'URL avec la clé API du module env
    ragicUrl = RAGIC_LOG_URL & "?APIKey=" & env.GetRagicApiKey()
    
    ' Envoyer la requête POST de manière asynchrone pour ne pas attendre la réponse
    http.Open "POST", ragicUrl, True ' True = Asynchrone
    http.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    http.send jsonPayload
End Sub

' ============================================================================
' FONCTIONS UTILITAIRES
' ============================================================================

Public Sub SetLogLevel(level As LogLevel)
    Const PROC_NAME As String = "SetLogLevel"
    Const MODULE_NAME As String = "SYS_Logger"
    On Error GoTo ErrorHandler

    mCurrentLogLevel = level
    Log "log_level_changed", "Niveau de log défini à: " & level, INFO_LEVEL
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Failed to set log level."
End Sub

' Purge les anciens fichiers de log
Public Sub PurgeOldLogs(Optional ByVal daysToKeep As Integer = 30)
    Const PROC_NAME As String = "PurgeOldLogs"
    Const MODULE_NAME As String = "SYS_Logger"
    On Error GoTo ErrorHandler

    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim currentDate As Date
    Dim cutoffDate As Date
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' S'assurer que le dossier existe
    EnsureLogFolderExists
    
    Set folder = fso.GetFolder(GetLogFolderPath())
    currentDate = Now
    cutoffDate = DateAdd("d", -daysToKeep, currentDate)
    
    ' Parcourir tous les fichiers du dossier
    For Each file In folder.Files
        ' Si le fichier est plus vieux que la date limite, le supprimer
        If file.DateLastModified < cutoffDate Then
            Log "purge_logs", "Suppression du fichier de log: " & file.Name, INFO_LEVEL, "PurgeOldLogs", "SYS_Logger"
            file.Delete
        End If
    Next file
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Failed to purge old logs."
End Sub

