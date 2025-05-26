' ============================================================================
' SYS_ErrorHandler - Centralized Error Management
' Elyse Energy VBA Ecosystem - Error Handling Component
' Requires: SYS_CoreSystem, SYS_Logger, SYS_MessageBox, SYS_TicketSystem
' ============================================================================

Option Explicit
Private Const MODULE_NAME As String = "SYS_ErrorHandler"

' ============================================================================
' MODULE DEPENDENCIES
' ============================================================================
' This module requires:
' - SYS_CoreSystem (enums, constants, utilities)
' - SYS_Logger (logging functions)
' - SYS_MessageBox (enhanced message boxes)
' - SYS_TicketSystem (ticket creation)

' ============================================================================
' ERROR HANDLING CONFIGURATION
' ============================================================================

Public Type ErrorContext
    ProcedureName As String
    ModuleName As String
    ErrorNumber As Long
    ErrorDescription As String
    ErrorSource As String
    Timestamp As Date
    UserAction As String
    SystemState As String
    RecoveryAttempted As Boolean
    TicketCreated As Boolean
    HelpFile As String
    HelpContext As String
    LineNumber As Long
    Severity As String
    ActionTaken As String
    AdditionalDetails As String
    Context As Dictionary
End Type

Public Enum ErrorSeverity
    LOW_SEVERITY = 1
    MEDIUM_SEVERITY = 2
    HIGH_SEVERITY = 3
    CRITICAL_SEVERITY = 4
End Enum

Public Enum ErrorAction
    LOG_ONLY = 1
    SHOW_MESSAGE = 2
    OFFER_TICKET = 3
    AUTO_RECOVERY = 4
    FORCE_TICKET = 5
End Enum

' ============================================================================
' ERROR HANDLER STATE
' ============================================================================

Private mErrorHandlerActive As Boolean
Private mCurrentErrorContext As ErrorContext
Private mErrorHistory As Collection
Private mAutoRecoveryEnabled As Boolean
Private mSuppressUserMessages As Boolean
Private mErrorHandlerStack As Collection

' Error statistics
Private mTotalErrors As Long
Private mCriticalErrors As Long
Private mRecoveredErrors As Long
Private mTicketsCreated As Long

' ============================================================================
' ERROR HANDLER INITIALIZATION
' ============================================================================

Public Function InitializeErrorHandler() As Boolean
    ' Initialize the centralized error handling system
    
    LogInfo "error_handler_init", "Initializing centralized error handler"
    
    ' Initialize collections
    Set mErrorHistory = New Collection
    Set mErrorHandlerStack = New Collection
    
    ' Set default configuration
    mErrorHandlerActive = True
    mAutoRecoveryEnabled = True
    mSuppressUserMessages = False
    
    ' Reset statistics
    ResetErrorStatistics
    
    ' Install global error hooks
    InstallGlobalErrorHooks
    
    LogInfo "error_handler_ready", "Error handler initialized successfully"
    InitializeErrorHandler = True
End Function

Public Sub ShutdownErrorHandler()
    ' Clean shutdown of error handling system
    
    LogInfo "error_handler_shutdown", "Shutting down error handler"
    
    ' Log final statistics
    LogErrorStatistics
    
    ' Cleanup
    Set mErrorHistory = Nothing
    Set mErrorHandlerStack = Nothing
    mErrorHandlerActive = False
End Sub

Private Sub InstallGlobalErrorHooks()
    ' Install global error handling hooks
    
    ' Enable events for error detection
    Application.EnableEvents = True
    
    ' Set up Excel error event handlers
    LogDebug "error_handler_hooks", "Global error hooks installed"
End Sub

' ============================================================================
' MAIN ERROR HANDLING FUNCTIONS
' ============================================================================

Public Function HandleError(ByVal moduleName As String, ByVal procedureName As String, Optional ByVal additionalInfo As String = "")
    Const PROC_NAME As String = "HandleError"
    On Error GoTo ErrorHandler
    
    ' Créer le contexte d'erreur
    Dim errorCtx As ErrorContext
    With errorCtx
        .ModuleName = moduleName
        .ProcedureName = procedureName
        .ErrorNumber = Err.Number
        .ErrorDescription = Err.Description
        .AdditionalInfo = additionalInfo
        .ErrorSource = Err.Source
        .ErrorLine = Erl
    End With
    
    ' Loguer l'erreur
    LogError "HandleError:" & actionCode, errorCtx.ErrorNumber, combinedDetails, procedureName, moduleName, errorCtx
    
    ' Afficher le message approprié selon le type d'erreur
    If IsCriticalError(errorCtx.ErrorNumber) Then
        ShowCriticalMessage "Critical Error in " & ctx.ModuleName & "." & ctx.ProcedureName, _
            "A critical error has occurred. The application may need to be restarted." & vbCrLf & _
            "Error: " & ctx.ErrorDescription & vbCrLf & _
            "Additional Info: " & ctx.AdditionalInfo
    ElseIf IsWarningError(errorCtx.ErrorNumber) Then
        ShowWarningMessage "Warning in " & ctx.ModuleName & "." & ctx.ProcedureName, _
            "A warning has been generated." & vbCrLf & _
            "Warning: " & ctx.ErrorDescription & vbCrLf & _
            "Additional Info: " & ctx.AdditionalInfo
    Else
        ShowErrorMessage "Error in " & ctx.ModuleName & "." & ctx.ProcedureName, _
            "An error has occurred." & vbCrLf & _
            "Error: " & ctx.ErrorDescription & vbCrLf & _
            "Additional Info: " & ctx.AdditionalInfo
    End If
    
    ' Afficher les informations de débogage si nécessaire
    If IsDebugMode() Then
        ShowInfoMessage "Information", ctx.ErrorDescription, Buttons:=vbOKOnly
    End If
    
    Exit Sub

ErrorHandler:
    ' En cas d'erreur dans le gestionnaire d'erreurs, on affiche un message simple
    MsgBox "Une erreur est survenue dans le gestionnaire d'erreurs. Détails: " & Err.Description, vbCritical, "Erreur Critique"
End Function

' =============================================
' Fonction de logging d'erreur avec contexte
' =============================================
Public Function LogErrorWithContext(ByVal errorNumber As Long, _
                                  ByVal errorDescription As String, _
                                  ByVal errorSource As String, _
                                  ByVal context As Dictionary) As Boolean
    On Error GoTo ErrorHandler
    
    ' Vérifier que le système est initialisé
    If Not mIsInitialized Then
        LogErrorWithContext = False
        Exit Function
    End If
    
    ' Créer un objet d'erreur
    Dim errObj As New ErrorContext
    With errObj
        .ErrorNumber = errorNumber
        .ErrorDescription = errorDescription
        .ErrorSource = errorSource
        .Context = context
        .Timestamp = Now
        .Severity = ErrorSeverity.Error
    End With
    
    ' Logger l'erreur
    LogErrorWithContext = HandleError(errObj)
    Exit Function
    
ErrorHandler:
    LogErrorWithContext = False
End Function

Private Function DetermineErrorSeverityAsString(errorNum As Long, errorDesc As String) As String
    ' TODO: Implement logic to determine severity based on error number or description
    ' Example:
    If errorNum = 0 Then DetermineErrorSeverityAsString = "INFO" ' Not an error
    ElseIf InStr(1, errorDesc, "critical", vbTextCompare) > 0 Then
        DetermineErrorSeverityAsString = "CRITICAL"
    ElseIf errorNum > 0 And errorNum < 100 Then ' Example: Application-defined errors
        DetermineErrorSeverityAsString = "HIGH"
    ElseIf errorNum >= 500 And errorNum <= 600 Then ' Example: File errors
        DetermineErrorSeverityAsString = "MEDIUM"
    Else
        DetermineErrorSeverityAsString = "LOW" ' Default
    End If
    ElyseLogger_Module.LogDebug "DetermineErrorSeverityAsString", "Severity for Err#" & errorNum & " determined as: " & DetermineErrorSeverityAsString, "DetermineErrorSeverity", "ElyseErrorHandler_Module"
End Function

Private Function ExecuteErrorActionAndGetResult(ctx As ErrorContext, Optional additionalInfo As String = "") As String
    ' TODO: Implement logic to decide what action to take (e.g., show message, retry, abort)
    Dim action As String
    action = "LOGGED" ' Default action

    Select Case ctx.Severity
        Case "CRITICAL"
            ElyseMessageBox_System.ShowCriticalMessage "Critical Error in " & ctx.ModuleName & "." & ctx.ProcedureName, _
                "A critical error occurred: " & ctx.ErrorDescription & vbCrLf & _
                "Error Number: " & ctx.ErrorNumber & vbCrLf & _
                "Additional Info: " & additionalInfo & vbCrLf & _
                "The application may be unstable. Please save your work and restart.", _
                Buttons:=vbOKOnly
            action = "SHOW_CRITICAL_ABORT" ' Suggests user should abort
        Case "HIGH"
            ElyseMessageBox_System.ShowErrorMessage "Error in " & ctx.ModuleName & "." & ctx.ProcedureName, _
                "An error occurred: " & ctx.ErrorDescription & vbCrLf & _
                "Error Number: " & ctx.ErrorNumber & vbCrLf & _
                "Additional Info: " & additionalInfo, _
                Buttons:=vbOKOnly
            action = "SHOW_ERROR"
        Case "MEDIUM"
            ElyseMessageBox_System.ShowWarningMessage "Warning in " & ctx.ModuleName & "." & ctx.ProcedureName, _
                "A potential issue arose: " & ctx.ErrorDescription & vbCrLf & _
                "Error Number: " & ctx.ErrorNumber & vbCrLf & _
                "Additional Info: " & additionalInfo, _
                Buttons:=vbOKOnly
            action = "SHOW_WARNING"
        Case "LOW"
            ' For low severity, often just logging is enough, no message box unless specifically needed.
            ' If additionalInfo suggests a user notification, it could be shown.
            If InStr(1, additionalInfo, "notify_user", vbTextCompare) > 0 Then
                 ElyseMessageBox_System.ShowInfoMessage "Information", ctx.ErrorDescription, Buttons:=vbOKOnly
                 action = "SHOW_INFO"
            End If
        Case Else
            ' Default: just logged
    End Select
      ElyseLogger_Module.LogDebug "ExecuteErrorActionAndGetResult", "Action taken for Err#" & ctx.ErrorNumber & " (" & ctx.Severity & "): " & action, "ExecuteErrorActionAndGetResult", "ElyseErrorHandler_Module"
    ExecuteErrorActionAndGetResult = action
End Function

' ============================================================================
' ERROR ANALYSIS AND CLASSIFICATION
' ============================================================================

Private Function DetermineErrorSeverity(errorCtx as Object) As ErrorSeverity
    ' Determine error severity based on error context
    
    Select Case errorCtx.ErrorNumber
        Case 1004, 1016 ' Range errors
            DetermineErrorSeverity = LOW_SEVERITY
            
        Case 9, 13, 91 ' Subscript, type mismatch, object variable errors
            DetermineErrorSeverity = MEDIUM_SEVERITY
            
        Case 7, 11 ' Out of memory, division by zero
            DetermineErrorSeverity = HIGH_SEVERITY
            
        Case 429, 462 ' ActiveX, automation errors
            DetermineErrorSeverity = HIGH_SEVERITY
            
        Case 9999 ' Custom errors
            DetermineErrorSeverity = MEDIUM_SEVERITY
            
        Case Else
            ' Check if error description contains critical keywords
            If InStr(LCase(errorCtx.ErrorDescription), "corrupt") > 0 Or _
               InStr(LCase(errorCtx.ErrorDescription), "fatal") > 0 Or _
               InStr(LCase(errorCtx.ErrorDescription), "cannot save") > 0 Then
                DetermineErrorSeverity = CRITICAL_SEVERITY
            Else
                DetermineErrorSeverity = MEDIUM_SEVERITY
            End If
    End Select
End Function

Private Function DetermineErrorAction(errorCtx as Object, severity As ErrorSeverity) As ErrorAction
    ' Determine what action to take based on error context and severity
    
    ' Check if user messages are suppressed
    If mSuppressUserMessages And severity < CRITICAL_SEVERITY Then
        DetermineErrorAction = LOG_ONLY
        Exit Function
    End If
    
    Select Case severity
        Case LOW_SEVERITY
            DetermineErrorAction = LOG_ONLY
            
        Case MEDIUM_SEVERITY
            If IsProductionMode() Then
                DetermineErrorAction = SHOW_MESSAGE
            Else
                DetermineErrorAction = OFFER_TICKET
            End If
            
        Case HIGH_SEVERITY
            DetermineErrorAction = OFFER_TICKET
            
        Case CRITICAL_SEVERITY
            DetermineErrorAction = FORCE_TICKET
    End Select
    
    ' Override for specific error types that support auto-recovery
    If mAutoRecoveryEnabled And CanAttemptAutoRecovery(errorCtx) Then
        DetermineErrorAction = AUTO_RECOVERY
    End If
End Function

Private Function CanAttemptAutoRecovery(errorCtx as Object) As Boolean
    ' Determine if auto-recovery can be attempted for this error
    
    Select Case errorCtx.ErrorNumber
        Case 1004 ' Range not found - can suggest alternative
            CanAttemptAutoRecovery = True
            
        Case 1016 ' No cells were found - can offer to expand search
            CanAttemptAutoRecovery = True
            
        Case 91 ' Object variable not set - can offer to reinitialize
            CanAttemptAutoRecovery = True
            
        Case Else
            CanAttemptAutoRecovery = False
    End Select
End Function

' ============================================================================
' ERROR ACTION EXECUTION
' ============================================================================

Private Sub ExecuteErrorAction(errorCtx as Object, severity As ErrorSeverity, action As ErrorAction, allowRecovery As Boolean)
    ' Execute the determined error action
    
    Select Case action
        Case LOG_ONLY
            ' Already logged, no further action needed
            
        Case SHOW_MESSAGE
            ShowErrorMessage errorCtx, severity, False
            
        Case OFFER_TICKET
            ShowErrorMessage errorCtx, severity, True
            
        Case AUTO_RECOVERY
            If allowRecovery Then
                AttemptAutoRecovery errorCtx
            Else
                ShowErrorMessage errorCtx, severity, True
            End If
            
        Case FORCE_TICKET
            CreateErrorTicket errorCtx
            ShowErrorMessage errorCtx, severity, False ' Message without ticket option since ticket is forced
    End Select
End Sub

Private Sub ShowErrorMessage(errorCtx as Object, severity As ErrorSeverity, allowTicketCreation As Boolean)
    ' Show error message to user with appropriate options
    
    Dim title As String
    Dim message As String
    
    title = BuildErrorTitle(errorCtx, severity)
    message = BuildUserFriendlyErrorMessage(errorCtx)
    
    Dim result As String
    result = ShowErrorMessage(title, message, allowTicketCreation)
    
    ' Handle user response
    If result = "CREATE_TICKET" Then
        CreateErrorTicket errorCtx
        errorCtx.TicketCreated = True
        mTicketsCreated = mTicketsCreated + 1
    End If
End Sub

Private Sub AttemptAutoRecovery(errorCtx as Object)
    ' Attempt automatic recovery for supported error types
    
    LogInfo "error_auto_recovery", "Attempting auto-recovery for error " & errorCtx.ErrorNumber
    
    Dim recoveryAttempted As Boolean
    recoveryAttempted = False
    
    Select Case errorCtx.ErrorNumber
        Case 1004 ' Range not found
            recoveryAttempted = RecoverRangeNotFound(errorCtx)
            
        Case 91 ' Object variable not set
            recoveryAttempted = RecoverObjectNotSet(errorCtx)
            
        Case Else
            ' No recovery available
    End Select
    
    errorCtx.RecoveryAttempted = recoveryAttempted
    
    If recoveryAttempted Then
        LogInfo "error_recovery_attempted", "Auto-recovery attempted for error " & errorCtx.ErrorNumber
        mRecoveredErrors = mRecoveredErrors + 1
        
        ' Show recovery message to user
        ShowInfoMessage "Auto-Recovery", "An error was detected and automatically resolved: " & vbCrLf & vbCrLf & errorCtx.ErrorDescription
    Else
        ' Recovery failed, fall back to showing error message
        ShowErrorMessage errorCtx, DetermineErrorSeverity(errorCtx), True
    End If
End Sub

' ============================================================================
' AUTO-RECOVERY IMPLEMENTATIONS
' ============================================================================

Private Function RecoverRangeNotFound(errorCtx as Object) As Boolean
    ' Attempt to recover from range not found errors
    On Error Resume Next
    
    ' This is a placeholder for range recovery logic
    ' In a real implementation, you might:
    ' 1. Try to find similar named ranges
    ' 2. Prompt user to select correct range
    ' 3. Use default ranges
    
    RecoverRangeNotFound = False
    On Error GoTo 0
End Function

Private Function RecoverObjectNotSet(errorCtx as Object) As Boolean
    ' Attempt to recover from object not set errors
    On Error Resume Next
    
    ' This is a placeholder for object recovery logic
    ' In a real implementation, you might:
    ' 1. Reinitialize common objects
    ' 2. Check if required add-ins are loaded
    ' 3. Restore default object states
    
    RecoverObjectNotSet = False
    On Error GoTo 0
End Function

' ============================================================================
' ERROR MESSAGE FORMATTING
' ============================================================================

Private Function BuildErrorTitle(errorCtx as Object, severity As ErrorSeverity) As String
    ' Build user-friendly error title
    
    Dim title As String
    
    Select Case severity
        Case LOW_SEVERITY
            title = "Minor Issue"
        Case MEDIUM_SEVERITY
            title = "Error"
        Case HIGH_SEVERITY
            title = "Serious Error"
        Case CRITICAL_SEVERITY
            title = "Critical Error"
    End Select
    
    If errorCtx.ModuleName <> "" Then
        title = title & " in " & errorCtx.ModuleName
    End If
    
    BuildErrorTitle = title
End Function

Private Function BuildUserFriendlyErrorMessage(errorCtx as Object) As String
    ' Build user-friendly error message
    
    Dim message As String
    
    ' Start with a friendly explanation
    message = GetUserFriendlyErrorExplanation(errorCtx.ErrorNumber)
    
    If message = "" Then
        message = "An error occurred while processing your request."
    End If
    
    message = message & vbCrLf & vbCrLf
    
    ' Add technical details in debug mode
    If IsDebugMode() Then
        message = message & "Technical Details:" & vbCrLf
        message = message & "Error " & errorCtx.ErrorNumber & ": " & errorCtx.ErrorDescription & vbCrLf
        message = message & "Location: " & errorCtx.ProcedureName & vbCrLf
        message = message & "Time: " & Format(errorCtx.Timestamp, "yyyy-mm-dd hh:nn:ss")
    Else
        message = message & "If this problem persists, please contact support for assistance."
    End If
    
    BuildUserFriendlyErrorMessage = message
End Function

Private Function GetUserFriendlyErrorExplanation(errorNumber As Long) As String
    ' Get user-friendly explanation for common errors
    
    Select Case errorNumber
        Case 1004
            GetUserFriendlyErrorExplanation = "The specified range or cell could not be found. Please check that the data exists and try again."
            
        Case 1016
            GetUserFriendlyErrorExplanation = "No matching data was found. You may need to adjust your search criteria or check the data source."
            
        Case 13
            GetUserFriendlyErrorExplanation = "There was a problem with the data format. Please check that all values are in the expected format."
            
        Case 9
            GetUserFriendlyErrorExplanation = "A calculation couldn't be completed because some required data is missing or invalid."
            
        Case 91
            GetUserFriendlyErrorExplanation = "A required component is not available. The system will attempt to reinitialize it."
            
        Case 429
            GetUserFriendlyErrorExplanation = "There was a problem connecting to a required service. Please try again in a moment."
            
        Case 462
            GetUserFriendlyErrorExplanation = "A remote server or service is not responding. Please check your network connection."
            
        Case 7
            GetUserFriendlyErrorExplanation = "The system is running low on memory. Please close other applications and try again."
            
        Case Else
            GetUserFriendlyErrorExplanation = ""
    End Select
End Function

Private Function BuildErrorLogMessage(errorCtx as Object) As String
    ' Build detailed error message for logging
    
    Dim logMessage As String
    
    logMessage = "Procedure: " & errorCtx.ProcedureName
    
    If errorCtx.ModuleName <> "" Then
        logMessage = logMessage & " | Module: " & errorCtx.ModuleName
    End If
    
    logMessage = logMessage & " | Error: " & errorCtx.ErrorNumber & " - " & errorCtx.ErrorDescription
    
    If errorCtx.UserAction <> "" Then
        logMessage = logMessage & " | User Action: " & errorCtx.UserAction
    End If
    
    logMessage = logMessage & " | System State: " & errorCtx.SystemState
    
    BuildErrorLogMessage = logMessage
End Function

' ============================================================================
' TICKET CREATION INTEGRATION
' ============================================================================

Private Sub CreateErrorTicket(errorCtx as Object)
    ' Create support ticket for error
    
    LogInfo "error_ticket_creation", "Creating support ticket for error " & errorCtx.ErrorNumber
    
    Dim ticketResult As String
    ticketResult = CreateTicketFromError(errorCtx.ProcedureName, errorCtx.ErrorNumber, errorCtx.ErrorDescription, True)
    
    If ticketResult <> "CANCELLED" And ticketResult <> "SEND_FAILED" Then
        LogInfo "error_ticket_created", "Support ticket created: " & ticketResult
        errorCtx.TicketCreated = True
    Else
        LogWarning "error_ticket_failed", "Failed to create support ticket for error " & errorCtx.ErrorNumber
    End If
End Sub

' ============================================================================
' ERROR HISTORY AND STATISTICS
' ============================================================================

Private Sub AddToErrorHistory(errorCtx as Object)
    ' Add error to history with size limit
    
    mErrorHistory.Add errorCtx
    
    ' Keep only last 100 errors
    Do While mErrorHistory.Count > 100
        mErrorHistory.Remove 1
    Loop
End Sub

Private Sub UpdateErrorStatistics(errorCtx as Object)
    ' Update error statistics
    
    mTotalErrors = mTotalErrors + 1
    
    Dim severity As ErrorSeverity
    severity = DetermineErrorSeverity(errorCtx)
    
    If severity = CRITICAL_SEVERITY Then
        mCriticalErrors = mCriticalErrors + 1
    End If
End Sub

Private Sub ResetErrorStatistics()
    ' Reset error statistics
    
    mTotalErrors = 0
    mCriticalErrors = 0
    mRecoveredErrors = 0
    mTicketsCreated = 0
End Sub

Private Sub LogErrorStatistics()
    ' Log current error statistics
    
    Dim statsMessage As String
    statsMessage = "Total: " & mTotalErrors & " | Critical: " & mCriticalErrors & " | Recovered: " & mRecoveredErrors & " | Tickets: " & mTicketsCreated
    
    LogInfo "error_statistics", statsMessage
End Sub

' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================

Private Function GetSystemState() As String
    ' Get current system state for error context
    
    Dim state As String
    state = "Workbook: " & GetCurrentWorkbookName()
    state = state & " | Sheet: " & GetActiveSheetName()
    state = state & " | Selection: " & GetSelectedRangeAddress()
    
    GetSystemState = state
End Function

' ============================================================================
' CONFIGURATION AND CONTROL
' ============================================================================

Public Sub EnableAutoRecovery()
    ' Enable automatic error recovery
    mAutoRecoveryEnabled = True
    LogInfo "error_auto_recovery_enabled", "Automatic error recovery enabled"
End Sub

Public Sub DisableAutoRecovery()
    ' Disable automatic error recovery
    mAutoRecoveryEnabled = False
    LogInfo "error_auto_recovery_disabled", "Automatic error recovery disabled"
End Sub

Public Sub SuppressUserMessages()
    ' Suppress error messages to users (for batch operations)
    mSuppressUserMessages = True
    LogInfo "error_messages_suppressed", "User error messages suppressed"
End Sub

Public Sub EnableUserMessages()
    ' Re-enable error messages to users
    mSuppressUserMessages = False
    LogInfo "error_messages_enabled", "User error messages enabled"
End Sub

' ============================================================================
' ERROR HANDLER STACK MANAGEMENT
' ============================================================================

Public Sub PushErrorHandler(handlerName As String)
    ' Push error handler onto stack
    mErrorHandlerStack.Add handlerName
End Sub

Public Sub PopErrorHandler()
    ' Pop error handler from stack
    If mErrorHandlerStack.Count > 0 Then
        mErrorHandlerStack.Remove mErrorHandlerStack.Count
    End If
End Sub

Private Function GetCurrentErrorHandler() As String
    ' Get current error handler name
    If mErrorHandlerStack.Count > 0 Then
        GetCurrentErrorHandler = mErrorHandlerStack(mErrorHandlerStack.Count)
    Else
        GetCurrentErrorHandler = "Unknown"
    End If
End Function

' ============================================================================
' PUBLIC CONVENIENCE FUNCTIONS
' ============================================================================

Public Sub LogAndHandleError(procedureName As String, Optional moduleName As String = "")
    ' Convenience function for standard error handling
    HandleError procedureName, moduleName
End Sub

Public Function TryOperation(operationName As String, operationProc As String) As Boolean
    ' Try an operation with automatic error handling
    On Error GoTo ErrorHandler
    
    PushErrorHandler operationName
    
    ' The calling code would execute the operation here
    ' This is a framework for protected operations
    
    PopErrorHandler
    TryOperation = True
    Exit Function
    
ErrorHandler:
    HandleError operationProc, operationName
    PopErrorHandler
    TryOperation = False
End Function

' ============================================================================
' DIAGNOSTICS AND STATUS
' ============================================================================

Public Function GetErrorHandlerStatus() As Object
    ' Get comprehensive error handler status
    
    Dim status As Object
    Set status = CreateObject("Scripting.Dictionary")
    
    status("active") = mErrorHandlerActive
    status("auto_recovery_enabled") = mAutoRecoveryEnabled
    status("suppress_messages") = mSuppressUserMessages
    status("total_errors") = mTotalErrors
    status("critical_errors") = mCriticalErrors
    status("recovered_errors") = mRecoveredErrors
    status("tickets_created") = mTicketsCreated
    status("current_handler") = GetCurrentErrorHandler()
    status("error_history_count") = mErrorHistory.Count
    
    Set GetErrorHandlerStatus = status
End Function

Public Function GetRecentErrors(Optional count As Integer = 10) As Collection
    Dim recentErrors As New Collection
    Dim i As Long
    
    ' Return empty collection if no errors
    If mErrorHistory.Count = 0 Then
        Set GetRecentErrors = recentErrors
        Exit Function
    End If
    
    ' Get the most recent errors
    For i = mErrorHistory.Count To 1 Step -1
        If recentErrors.Count >= count Then Exit For
        recentErrors.Add mErrorHistory(i)
    Next i
    
    Set GetRecentErrors = recentErrors
End Function

' ============================================================================
' TEMPLATE ERROR HANDLERS FOR COMMON SCENARIOS
' ============================================================================

Public Sub HandleCalculationError(calculationType As String, inputData As String)
    ' Template for calculation errors
    
    Dim userAction As String
    userAction = "Calculation: " & calculationType & " | Input: " & inputData
    
    HandleError "Calculation", "ElyseCalculations", userAction
End Sub

Public Sub HandleDataAccessError(dataSource As String, operation As String)
    ' Template for data access errors
    
    Dim userAction As String
    userAction = "Data Access: " & operation & " | Source: " & dataSource
    
    HandleError "DataAccess", "ElyseDataManager", userAction
End Sub

Public Sub HandleUIError(controlName As String, action As String)
    ' Template for UI errors
    
    Dim userAction As String
    userAction = "UI Action: " & action & " | Control: " & controlName
    
    HandleError "UserInterface", "ElyseUI", userAction
End Sub