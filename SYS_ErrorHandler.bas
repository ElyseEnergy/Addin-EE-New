Attribute VB_Name = "SYS_ErrorHandler"
Option Explicit

' ============================================================================
' SYS_ErrorHandler - Gestionnaire d'erreurs simplifi�
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

' Variables d'�tat
Private mErrorHandlerActive As Boolean

' ============================================================================
' INITIALISATION
' ============================================================================

Public Function InitializeErrorHandler() As Boolean
    On Error GoTo ErrorHandler
    
    mErrorHandlerActive = True
    Log "error_handler_init", "Error handler initialized", INFO_LEVEL
    
    InitializeErrorHandler = True
    Exit Function
    
ErrorHandler:
    InitializeErrorHandler = False
End Function

Public Sub ShutdownErrorHandler()
    If Not mErrorHandlerActive Then Exit Sub
    
    Log "error_handler_shutdown", "Error handler shutdown", INFO_LEVEL
    mErrorHandlerActive = False
End Sub

' ============================================================================
' GESTION DES ERREURS
' ============================================================================

Public Sub HandleError(ByVal moduleName As String, ByVal procedureName As String, Optional ByVal additionalInfo As String = "")
    On Error GoTo ErrorHandler
    
    ' Cr�er le contexte d'erreur
    Dim errorCtx As ErrorContext
    With errorCtx
        .errorNumber = Err.Number
        .ErrorDescription = Err.Description
        .ErrorSource = Err.Source
        .procedureName = procedureName
        .moduleName = moduleName
        .timeStamp = Now
    End With
    
    ' D�terminer la s�v�rit�
    Dim severity As ErrorSeverity
    severity = DetermineErrorSeverity(errorCtx)
    
    ' Loguer l'erreur
    Log "error_occurred", BuildErrorLogMessage(errorCtx), severity
    
    ' Afficher le message appropri�
    ShowErrorMessage errorCtx, severity
    
    Exit Sub
    
ErrorHandler:
    ' En cas d'erreur dans le gestionnaire d'erreurs
    MsgBox "Une erreur est survenue dans le gestionnaire d'erreurs. D�tails: " & Err.Description, vbCritical, "Erreur Critique"
End Sub

' ============================================================================
' FONCTIONS UTILITAIRES
' ============================================================================

Private Function DetermineErrorSeverity(errorCtx As ErrorContext) As ErrorSeverity
    Select Case errorCtx.errorNumber
        Case 1004, 1016 ' Erreurs de plage
            DetermineErrorSeverity = LOW_SEVERITY
            
        Case 9, 13, 91 ' Erreurs de type, objet
            DetermineErrorSeverity = MEDIUM_SEVERITY
            
        Case 7, 11 ' M�moire, division par z�ro
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
End Function

Private Sub ShowErrorMessage(errorCtx As ErrorContext, severity As ErrorSeverity)
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
              "D�tails techniques:" & vbCrLf & _
              "Erreur " & errorCtx.errorNumber & ": " & errorCtx.ErrorDescription & vbCrLf & _
              "Emplacement: " & errorCtx.procedureName
    
    ' Afficher le message
    MsgBox message, IIf(severity >= HIGH_SEVERITY, vbCritical, vbExclamation), title
End Sub

Private Function BuildErrorLogMessage(errorCtx As ErrorContext) As String
    BuildErrorLogMessage = "Error " & errorCtx.errorNumber & " in " & _
                          errorCtx.moduleName & "." & errorCtx.procedureName & ": " & _
                          errorCtx.ErrorDescription
End Function

Private Function GetUserFriendlyErrorExplanation(errorNumber As Long) As String
    Select Case errorNumber
        Case 1004
            GetUserFriendlyErrorExplanation = "La plage ou la cellule sp�cifi�e n'a pas �t� trouv�e. Veuillez v�rifier que les donn�es existent et r�essayer."
            
        Case 1016
            GetUserFriendlyErrorExplanation = "Aucune donn�e correspondante n'a �t� trouv�e. Vous devrez peut-�tre ajuster vos crit�res de recherche."
            
        Case 13
            GetUserFriendlyErrorExplanation = "Il y a eu un probl�me avec le format des donn�es. Veuillez v�rifier que toutes les valeurs sont au format attendu."
            
        Case 9
            GetUserFriendlyErrorExplanation = "Un calcul n'a pas pu �tre effectu� car certaines donn�es requises sont manquantes ou invalides."
            
        Case 91
            GetUserFriendlyErrorExplanation = "Un composant requis n'est pas disponible. Le syst�me va tenter de le r�initialiser."
            
        Case 429
            GetUserFriendlyErrorExplanation = "Il y a eu un probl�me de connexion � un service requis. Veuillez r�essayer dans un moment."
            
        Case 462
            GetUserFriendlyErrorExplanation = "Un serveur ou un service distant ne r�pond pas. Veuillez v�rifier votre connexion r�seau."
            
        Case 7
            GetUserFriendlyErrorExplanation = "Le syst�me manque de m�moire. Veuillez fermer d'autres applications et r�essayer."
            
        Case Else
            GetUserFriendlyErrorExplanation = ""
    End Select
End Function

