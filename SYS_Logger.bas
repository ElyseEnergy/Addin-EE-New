Option Explicit

' ============================================================================
' SYS_Logger - Système de Logging Simplifié
' ============================================================================

' Chemin absolu du dossier de logs
Private Const LOG_FOLDER_PATH As String = "c:\Users\JulienFernandez\OneDrive\Coding\_Projets de code\2025.05 New addin EE perso\logs"
Private Const LOG_FILE_NAME As String = "elyse_energy.log"

' Niveaux de log
Public Enum LogLevel
    DEBUG_LEVEL = 0
    INFO_LEVEL = 1
    WARNING_LEVEL = 2
    ERROR_LEVEL = 3
    CRITICAL_LEVEL = 4
End Enum

' Variables du module
Private mCurrentLogLevel As LogLevel

' ============================================================================
' INITIALISATION
' ============================================================================

Public Sub InitializeLogger()
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "InitializeLogger"
    Const MODULE_NAME As String = "SYS_Logger"
    
    ' Initialisation du niveau de log par défaut
    mCurrentLogLevel = INFO_LEVEL
    
    ' S'assurer que le dossier de logs existe
    EnsureLogFolderExists
    
    ' Log d'initialisation
    Log "sys_init", "Système de logging initialisé", INFO_LEVEL, PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de l'initialisation du système de logging"
End Sub

' ============================================================================
' FONCTION DE LOGGING
' ============================================================================

' Assure que le dossier de logs existe
Private Sub EnsureLogFolderExists()
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "EnsureLogFolderExists"
    Const MODULE_NAME As String = "SYS_Logger"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(LOG_FOLDER_PATH) Then
        fso.CreateFolder LOG_FOLDER_PATH
    End If
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la création du dossier de logs: " & LOG_FOLDER_PATH
End Sub

' Écrit dans le fichier de log
Private Sub WriteToLogFile(logMessage As String)
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "WriteToLogFile"
    Const MODULE_NAME As String = "SYS_Logger"
    
    Dim fso As Object
    Dim logFile As Object
    Dim logFilePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    logFilePath = LOG_FOLDER_PATH & "\" & LOG_FILE_NAME
    
    ' S'assurer que le dossier existe
    EnsureLogFolderExists
    
    ' Ouvrir le fichier en mode append (8 = ForAppending, -1 = TristateMixed pour Unicode)
    Set logFile = fso.OpenTextFile(logFilePath, 8, True, -1)
    
    ' Écrire le message
    logFile.WriteLine logMessage
    
    ' Fermer le fichier
    logFile.Close
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de l'écriture dans le fichier de log: " & logFilePath
End Sub

Public Sub Log(actionCode As String, message As String, level As LogLevel, _
    Optional ByVal procedureName As String = "", Optional ByVal moduleName As String = "")
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "Log"
    Const MODULE_NAME As String = "SYS_Logger"
    
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
    Exit Sub

ErrorHandler:
    ' Note: On ne peut pas utiliser HandleError ici car cela créerait une boucle infinie
    Debug.Print "ERREUR CRITIQUE DANS LE SYSTÈME DE LOG: " & Err.Description
End Sub

' ============================================================================
' FONCTIONS UTILITAIRES
' ============================================================================

Public Sub SetLogLevel(level As LogLevel)
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "SetLogLevel"
    Const MODULE_NAME As String = "SYS_Logger"
    
    mCurrentLogLevel = level
    Log "log_level_changed", "Niveau de log défini à: " & level, INFO_LEVEL, PROC_NAME, MODULE_NAME
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du changement de niveau de log"
End Sub

' Purge les anciens fichiers de log
Public Sub PurgeOldLogs(Optional ByVal daysToKeep As Integer = 30)
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "PurgeOldLogs"
    Const MODULE_NAME As String = "SYS_Logger"
    
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim currentDate As Date
    Dim cutoffDate As Date
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' S'assurer que le dossier existe
    EnsureLogFolderExists
    
    Set folder = fso.GetFolder(LOG_FOLDER_PATH)
    currentDate = Now
    cutoffDate = DateAdd("d", -daysToKeep, currentDate)
    
    ' Parcourir tous les fichiers du dossier
    For Each file In folder.Files
        ' Si le fichier est plus vieux que la date limite, le supprimer
        If file.DateLastModified < cutoffDate Then
            Log "purge_logs", "Suppression du fichier de log: " & file.Name, INFO_LEVEL, PROC_NAME, MODULE_NAME
            file.Delete
        End If
    Next file
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la purge des anciens fichiers de log"
End Sub