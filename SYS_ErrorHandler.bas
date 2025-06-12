Attribute VB_Name = "SYS_ErrorHandler"
Option Explicit

' ============================================================================
' SYS_ErrorHandler - Gestionnaire d'erreurs simplifié
' ============================================================================

' Types et Enums
Public Type ErrorContext
    errorNumber As Long
    ErrorDescription As String
    ErrorSource As String
    procedureName As String
    moduleName As String
    timeStamp As Date
End Type

Public Enum ErrorSeverity
    LOW_SEVERITY = 1
    MEDIUM_SEVERITY = 2
    HIGH_SEVERITY = 3
    CRITICAL_SEVERITY = 4
End Enum

' Variables d'état
Private mErrorHandlerActive As Boolean

' ============================================================================
' INITIALISATION
' ============================================================================

Public Function InitializeErrorHandler() As Boolean
    On Error GoTo ErrorHandler
    
    mErrorHandlerActive = True
    SYS_Logger.Log "error_handler_init", "Error handler initialized", INFO_LEVEL
    
    InitializeErrorHandler = True
    Exit Function
    
ErrorHandler:
    InitializeErrorHandler = False
End Function

Public Sub ShutdownErrorHandler()
    Const PROC_NAME As String = "ShutdownErrorHandler"
    Const MODULE_NAME As String = "SYS_ErrorHandler"
    On Error GoTo ErrorHandler

    If Not mErrorHandlerActive Then Exit Sub
    
    SYS_Logger.Log "error_handler_shutdown", "Error handler shutdown", INFO_LEVEL
    mErrorHandlerActive = False
    Exit Sub

ErrorHandler:
    ' Ne pas appeler HandleError ici pour éviter une boucle infinie.
    Debug.Print "CRITICAL ERROR in " & MODULE_NAME & "." & PROC_NAME & ": " & Err.Description
End Sub

' ============================================================================
' GESTION DES ERREURS
' ============================================================================

Public Sub HandleError(ByVal moduleName As String, ByVal procedureName As String, Optional ByVal additionalInfo As String = "")
    On Error GoTo ErrorHandler
    
    ' Créer le contexte d'erreur
    Dim errorCtx As ErrorContext
    With errorCtx
        .errorNumber = Err.Number
        .ErrorDescription = Err.Description
        .ErrorSource = Err.Source
        .procedureName = procedureName
        .moduleName = moduleName
        .timeStamp = Now
    End With
    
    ' Déterminer la sévérité
    Dim severity As ErrorSeverity
    severity = DetermineErrorSeverity(errorCtx)
    
    ' Loguer l'erreur
    SYS_Logger.Log "error_occurred", BuildErrorLogMessage(errorCtx), ConvertSeverityToLogLevel(severity)
    
    ' Afficher le message approprié
    ShowErrorMessage errorCtx, severity
    
    Exit Sub
    
ErrorHandler:
    ' En cas d'erreur dans le gestionnaire d'erreurs
    MsgBox "Une erreur est survenue dans le gestionnaire d'erreurs. Détails: " & Err.Description, vbCritical, "Erreur Critique"
End Sub

' ============================================================================
' FONCTIONS UTILITAIRES
' ============================================================================

Private Function ConvertSeverityToLogLevel(ByVal severity As ErrorSeverity) As LogLevel
    Select Case severity
        Case LOW_SEVERITY
            ConvertSeverityToLogLevel = WARNING_LEVEL
        Case MEDIUM_SEVERITY
            ConvertSeverityToLogLevel = ERROR_LEVEL
        Case HIGH_SEVERITY, CRITICAL_SEVERITY
            ConvertSeverityToLogLevel = CRITICAL_LEVEL
        Case Else
            ConvertSeverityToLogLevel = INFO_LEVEL
    End Select
End Function

Private Function DetermineErrorSeverity(errorCtx As ErrorContext) As ErrorSeverity
    Const PROC_NAME As String = "DetermineErrorSeverity"
    On Error GoTo ErrorHandler
    Select Case errorCtx.errorNumber
        Case 1004, 1016 ' Erreurs de plage
            DetermineErrorSeverity = LOW_SEVERITY
            
        Case 9, 13, 91 ' Erreurs de type, objet
            DetermineErrorSeverity = MEDIUM_SEVERITY
            
        Case 7, 11 ' Mémoire, division par zéro
            DetermineErrorSeverity = HIGH_SEVERITY
            
        Case 429, 462 ' Erreurs ActiveX
            DetermineErrorSeverity = HIGH_SEVERITY
            
        Case Else
            If InStr(LCase(errorCtx.ErrorDescription), "critique") > 0 Or _
               InStr(LCase(errorCtx.ErrorDescription), "fatal") > 0 Then
                DetermineErrorSeverity = CRITICAL_SEVERITY
            Else
                DetermineErrorSeverity = MEDIUM_SEVERITY
            End If
    End Select
    Exit Function
ErrorHandler:
    DetermineErrorSeverity = CRITICAL_SEVERITY ' Si la détermination échoue, on est prudent
End Function

Private Sub ShowErrorMessage(errorCtx As ErrorContext, severity As ErrorSeverity)
    Const PROC_NAME As String = "ShowErrorMessage"
    On Error GoTo ErrorHandler
    Dim title As String
    Dim message As String
    
    ' Construire le titre
    Select Case severity
        Case LOW_SEVERITY
            title = "Avertissement"
        Case MEDIUM_SEVERITY
            title = "Erreur"
        Case HIGH_SEVERITY
            title = "Erreur Grave"
        Case CRITICAL_SEVERITY
            title = "Erreur Critique"
    End Select
    
    ' Construire le message
    message = GetUserFriendlyErrorExplanation(errorCtx.errorNumber)
    If message = "" Then
        message = "Une erreur est survenue lors du traitement de votre demande."
    End If
    
    message = message & vbCrLf & vbCrLf & _
              "Détails techniques:" & vbCrLf & _
              "Erreur " & errorCtx.errorNumber & ": " & errorCtx.ErrorDescription & vbCrLf & _
              "Emplacement: " & errorCtx.procedureName
    
    ' Afficher le message
    MsgBox message, IIf(severity >= HIGH_SEVERITY, vbCritical, vbExclamation), title
    Exit Sub
ErrorHandler:
    MsgBox "Erreur lors de l'affichage du message d'erreur. " & Err.Description, vbCritical
End Sub

Private Function BuildErrorLogMessage(errorCtx As ErrorContext) As String
    Const PROC_NAME As String = "BuildErrorLogMessage"
    On Error GoTo ErrorHandler
    BuildErrorLogMessage = "Error " & errorCtx.errorNumber & " in " & _
                          errorCtx.moduleName & "." & errorCtx.procedureName & ": " & _
                          errorCtx.ErrorDescription
    Exit Function
ErrorHandler:
    BuildErrorLogMessage = "Failed to build error log message."
End Function

Private Function GetUserFriendlyErrorExplanation(errorNumber As Long) As String
    Const PROC_NAME As String = "GetUserFriendlyErrorExplanation"
    On Error GoTo ErrorHandler
    Select Case errorNumber
        Case 1004
            GetUserFriendlyErrorExplanation = "La plage ou la cellule spécifiée n'a pas été trouvée. Veuillez vérifier que les données existent et réessayer."
            
        Case 1016
            GetUserFriendlyErrorExplanation = "Aucune donnée correspondante n'a été trouvée. Vous devrez peut-être ajuster vos critères de recherche."
            
        Case 13
            GetUserFriendlyErrorExplanation = "Il y a eu un problème avec le format des données. Veuillez vérifier que toutes les valeurs sont au format attendu."
            
        Case 9
            GetUserFriendlyErrorExplanation = "Un calcul n'a pas pu être effectué car certaines données requises sont manquantes ou invalides."
            
        Case 91
            GetUserFriendlyErrorExplanation = "Un composant requis n'est pas disponible. Le système va tenter de le réinitialiser."
            
        Case 429
            GetUserFriendlyErrorExplanation = "Il y a eu un problème de connexion à un service requis. Veuillez réessayer dans un moment."
            
        Case 462
            GetUserFriendlyErrorExplanation = "Un serveur ou un service distant ne répond pas. Veuillez vérifier votre connexion réseau."
            
        Case 7
            GetUserFriendlyErrorExplanation = "Le système manque de mémoire. Veuillez fermer d'autres applications et réessayer."
            
        Case Else
            GetUserFriendlyErrorExplanation = ""
    End Select
    Exit Function
ErrorHandler:
    GetUserFriendlyErrorExplanation = "Erreur lors de la récupération de l'explication."
End Function

