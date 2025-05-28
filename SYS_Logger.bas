Option Explicit

' ============================================================================
' SYS_Logger - Système de Logging Simplifié
' ============================================================================

' Niveaux de log
Public Enum LogLevel
    DEBUG_LEVEL = 0
    INFO_LEVEL = 1
    WARNING_LEVEL = 2
    ERROR_LEVEL = 3
    CRITICAL_LEVEL = 4
End Enum

' Niveau de log par défaut (INFO par défaut)
Private mCurrentLogLevel As LogLevel = INFO_LEVEL

' ============================================================================
' FONCTION DE LOGGING
' ============================================================================

Public Sub Log(actionCode As String, message As String, level As LogLevel, _
    Optional ByVal procedureName As String = "", Optional ByVal moduleName As String = "")
    
    If level < mCurrentLogLevel Then Exit Sub
    
    Dim logMessage As String
    Dim levelString As String
    
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
    
    ' Construire le message
    logMessage = levelString & " [" & actionCode & "] "
    
    ' Ajouter le contexte (module.procedure)
    If moduleName <> "" And procedureName <> "" Then
        logMessage = logMessage & moduleName & "." & procedureName
    ElseIf procedureName <> "" Then
        logMessage = logMessage & procedureName
    End If
    
    ' Ajouter le message
    logMessage = logMessage & ": " & message
    
    ' Afficher dans Immediate Window
    Debug.Print "[" & Format(Now, "hh:nn:ss") & "] " & logMessage
End Sub

' ============================================================================
' FONCTION UTILITAIRE
' ============================================================================

Public Sub SetLogLevel(level As LogLevel)
    mCurrentLogLevel = level
    Log "log_level_changed", "Niveau de log défini à: " & level, INFO_LEVEL
End Sub