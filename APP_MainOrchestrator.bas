' ============================================================================
' APP_MainOrchestrator - Main Application Orchestrator
' Elyse Energy VBA Ecosystem - Main Application Logic
' Requires: SYS_CoreSystem, SYS_Logger, SYS_ErrorHandler, SYS_RibbonCallbacks, SYS_WorkbookEvents, SYS_SystemEvents
' ============================================================================

Option Explicit

' ============================================================================
' MODULE DEPENDENCIES
' ============================================================================
' This module requires:
' - SYS_CoreSystem (enums, constants, utilities)
' - SYS_Logger (logging functions)
' - SYS_ErrorHandler (error handling functions)


' ============================================================================
' ORCHESTRATOR STATE AND CONFIGURATION
' ============================================================================

Private mSystemInitialized As Boolean
Private mModulesLoaded As Object
Private mSystemMode As SystemMode
Private mStartupTime As Date
Private mShutdownInProgress As Boolean

' Module status tracking
Private mCoreSystemStatus As Boolean
Private mLoggerStatus As Boolean
Private mMessageBoxStatus As Boolean
Private mTicketSystemStatus As Boolean
Private mSharePointStatus As Boolean
Private mErrorHandlerStatus As Boolean

' System configuration
Private mAutoStartModules As Boolean
Private mEnableHeartbeat As Boolean
Private mSystemHealthChecks As Boolean

' ============================================================================
' SYSTEM INITIALIZATION AND STARTUP
' ============================================================================

Public Function InitializeElyseSystem(Optional systemMode As SystemMode = PRODUCTION_MODE, Optional autoStart As Boolean = True) As Boolean
    ' Main system initialization - call this first
    
    On Error GoTo ErrorHandler
    
    If mSystemInitialized Then
        InitializeElyseSystem = True
        Exit Function
    End If
    
    mStartupTime = Now
    mSystemMode = systemMode
    mAutoStartModules = autoStart
    mShutdownInProgress = False
    
    ' Initialize module status tracking
    Set mModulesLoaded = CreateObject("Scripting.Dictionary")
    
    
    ' Phase 1: Initialize Core System
    If Not InitializeCoreModule() Then
        GoTo ErrorHandler
    End If
    
    ' Phase 2: Initialize Logger (depends on Core)
    If Not InitializeLoggerModule() Then
        GoTo ErrorHandler
    End If
    
    ' Log system startup
    LogInfo "system_startup", "Elyse Energy system starting up - Mode: " & SYS_CoreSystem.GetSystemModeString(systemMode)
    
    ' Phase 3: Initialize remaining modules if auto-start enabled
    If mAutoStartModules Then
        InitializeAllModules
    End If
    
    ' Phase 4: Post-initialization setup
    If Not CompleteSystemInitialization() Then
        GoTo ErrorHandler
    End If
    
    mSystemInitialized = True
    
    LogInfo "system_ready", "Elyse Energy system ready - " & GetLoadedModulesCount() & " modules loaded"
    InitializeElyseSystem = True
    Exit Function
    
ErrorHandler:
    LogError "system_init_failed", Err.Number, "System initialization failed: " & Err.description
    InitializeElyseSystem = False
End Function

Private Function InitializeCoreModule() As Boolean
    ' Initialize core system module
    On Error GoTo ErrorHandler
    
    mCoreSystemStatus = InitializeCoreSystem(mSystemMode)
    mModulesLoaded("core") = mCoreSystemStatus
    
    ' Initialize Core System Extras (for Ragic logging flag, session ID)
    Call SYS_CoreSystem.InitializeCoreSystem(DEBUG_MODE)

    InitializeCoreModule = mCoreSystemStatus
    Exit Function
    
ErrorHandler:
    InitializeCoreModule = False
End Function

Private Function InitializeLoggerModule() As Boolean
    ' Initialize logger module
    On Error GoTo ErrorHandler
    
    If Not mCoreSystemStatus Then
        InitializeLoggerModule = False
        Exit Function
    End If
    
    Dim logLevel As logLevel
    logLevel = IIf(mSystemMode = DEBUG_MODE, DEBUG_LEVEL, INFO_LEVEL)
    
    mLoggerStatus = InitializeLogger(logLevel)
    mModulesLoaded("logger") = mLoggerStatus
    
    InitializeLoggerModule = mLoggerStatus
    Exit Function
    
ErrorHandler:
    InitializeLoggerModule = False
End Function

Private Sub InitializeAllModules()
    ' Initialize all remaining modules
    
    LogInfo "modules_init_start", "Initializing all system modules"
    
    ' Initialize Error Handler (should be early for better error handling)
    InitializeErrorHandlerModule
    
    ' Initialize SharePoint Integration
    InitializeSharePointModule
    
    ' Initialize Message Box System
    InitializeMessageBoxModule
    
    ' Initialize Ticket System
    InitializeTicketModule
    
    LogInfo "modules_init_complete", "Module initialization complete"
End Sub

Private Sub InitializeErrorHandlerModule()
    ' Initialize error handler module
    On Error Resume Next
    
    mErrorHandlerStatus = InitializeErrorHandler()
    mModulesLoaded("error_handler") = mErrorHandlerStatus
    
    If mErrorHandlerStatus Then
        LogInfo "module_loaded", "Error Handler module loaded successfully"
    Else
        LogWarning "module_load_failed", "Error Handler module failed to load"
    End If
    
    On Error GoTo 0
End Sub

Private Sub InitializeSharePointModule()
    ' Initialize SharePoint integration module
    On Error Resume Next
    
    mSharePointStatus = InitializeSharePointIntegration()
    mModulesLoaded("sharepoint") = mSharePointStatus
    
    If mSharePointStatus Then
        LogInfo "module_loaded", "SharePoint Integration module loaded successfully"
    Else
        LogWarning "module_load_failed", "SharePoint Integration module failed to load or not applicable"
    End If
    
    On Error GoTo 0
End Sub

Private Sub InitializeMessageBoxModule()
    ' Initialize message box system module
    On Error Resume Next
    
    ' MessageBox system doesn't have explicit initialization but we track its availability
    mMessageBoxStatus = True ' Always available
    mModulesLoaded("messagebox") = mMessageBoxStatus
    
    LogInfo "module_loaded", "MessageBox System module loaded successfully"
    
    On Error GoTo 0
End Sub

Private Sub InitializeTicketModule()
    ' Initialize ticket system module
    On Error Resume Next
    
    ' Ticket system doesn't have explicit initialization but we track its availability
    mTicketSystemStatus = True ' Always available
    mModulesLoaded("ticket_system") = mTicketSystemStatus
    
    LogInfo "module_loaded", "Ticket System module loaded successfully"
    
    On Error GoTo 0
End Sub

Private Function CompleteSystemInitialization() As Boolean
    ' Complete system initialization with post-setup tasks
    On Error GoTo ErrorHandler
    
    ' Enable automatic error handling if error handler is loaded
    If mErrorHandlerStatus Then
        EnableAutoRecovery
    End If
    
    ' Start heartbeat system if logger is loaded
    If mLoggerStatus And mEnableHeartbeat Then
        EnableAutoFlush
    End If
    
    ' Perform initial system health check
    If mSystemHealthChecks Then
        PerformSystemHealthCheck
    End If
    
    ' Set up integrated error handling across modules
    SetupIntegratedErrorHandling
    
    CompleteSystemInitialization = True
    Exit Function
    
ErrorHandler:
    CompleteSystemInitialization = False
End Function

' ============================================================================
' SYSTEM SHUTDOWN AND CLEANUP
' ============================================================================

Public Sub ShutdownElyseSystem()
    ' Clean shutdown of entire system
    
    If Not mSystemInitialized Or mShutdownInProgress Then Exit Sub
    
    mShutdownInProgress = True
    
    LogInfo "system_shutdown_start", "Elyse Energy system shutdown initiated"
    
    ' Shutdown modules in reverse order
    ShutdownTicketModule
    ShutdownMessageBoxModule
    ShutdownSharePointModule
    ShutdownErrorHandlerModule
    ShutdownLoggerModule
    ShutdownCoreModule
    
    ' Final cleanup
    Set mModulesLoaded = Nothing
    mSystemInitialized = False
    
    ' Note: Cannot log after logger shutdown
End Sub

Private Sub ShutdownTicketModule()
    ' Shutdown ticket system module
    On Error Resume Next
    ' Ticket system cleanup (if needed)
    mModulesLoaded("ticket_system") = False
    On Error GoTo 0
End Sub

Private Sub ShutdownMessageBoxModule()
    ' Shutdown message box system module
    On Error Resume Next
    ' MessageBox system cleanup (if needed)
    mModulesLoaded("messagebox") = False
    On Error GoTo 0
End Sub

Private Sub ShutdownSharePointModule()
    ' Shutdown SharePoint integration module
    On Error Resume Next
    If mSharePointStatus Then
        CleanupSharePointIntegration
    End If
    mModulesLoaded("sharepoint") = False
    On Error GoTo 0
End Sub

Private Sub ShutdownErrorHandlerModule()
    ' Shutdown error handler module
    On Error Resume Next
    If mErrorHandlerStatus Then
        ShutdownErrorHandler
    End If
    mModulesLoaded("error_handler") = False
    On Error GoTo 0
End Sub

Private Sub ShutdownLoggerModule()
    ' Shutdown logger module
    On Error Resume Next
    If mLoggerStatus Then
        LogInfo "system_shutdown_complete", "Elyse Energy system shutdown complete"
        ShutdownLogger
    End If
    mModulesLoaded("logger") = False
    On Error GoTo 0
End Sub

Private Sub ShutdownCoreModule()
    ' Shutdown core system module
    On Error Resume Next
    ShutdownCoreSystem
    mModulesLoaded("core") = False
    On Error GoTo 0
End Sub

' ============================================================================
' UNIFIED PUBLIC API
' ============================================================================

' Logging API - Delegated to SYS_Logger








Public Sub LogUserAction(actionCode As String, description As String, Optional ByVal controlName As String = "")
    If Not mSystemInitialized Then InitializeElyseSystem
    If Not mLoggerStatus Then Exit Sub
    SYS_Logger.LogUserAction actionCode, description, controlName
End Sub

Public Sub LogRibbonAction(buttonId As String, Optional additionalInfo As String = "")
    If Not mSystemInitialized Then InitializeElyseSystem
    If Not mLoggerStatus Then Exit Sub
    SYS_Logger.LogRibbonAction buttonId, additionalInfo
End Sub

' Enhanced MessageBox API


Public Function ShowConfirmation(title As String, message As String) As Boolean
    If mMessageBoxStatus Then
        Dim result As String
        result = ShowEnhancedMessageBox(title, message, CONFIRMATION_MESSAGE, "Yes,No")
        ShowConfirmation = (result = "Yes")
    Else
        ShowConfirmation = (MsgBox(message, vbYesNo + vbQuestion, title) = vbYes)
    End Function
End Function

Public Function SelectFromList(title As String, message As String, items As Collection) As Long
    If mMessageBoxStatus Then
        SelectFromList = ShowListSelectionBox(title, message, items)
    Else
        SelectFromList = 0 ' Fallback not available
    End If
End Function

Public Function SelectRange(title As String, message As String, Optional defaultRange As String = "") As Range
    If mMessageBoxStatus Then
        Set SelectRange = ShowRangeSelectorBox(title, message, defaultRange)
    Else
        Set SelectRange = Nothing
    End If
End Function

Public Function ShowMarkdownInfo(title As String, content As String) As Long
    If mMessageBoxStatus Then
        ShowMarkdownInfo = ShowMarkdownInfoBox(title, content)
    Else
        MsgBox content, vbInformation, title
        ShowMarkdownInfo = 1
    End If
End Function

' Ticket System API
Public Function CreateSupportTicket() As String
    If mTicketSystemStatus Then
        CreateSupportTicket = CreateManualTicket()
    Else
        CreateSupportTicket = "SYSTEM_NOT_AVAILABLE"
    End Function

Public Function CreateErrorTicket(errorMsg As String, Optional errorCode As Long = 0) As String
    If mTicketSystemStatus Then
        CreateErrorTicket = CreateQuickErrorTicket(errorMsg, errorCode)
    Else
        CreateErrorTicket = "SYSTEM_NOT_AVAILABLE"
    End If
End Function

' SharePoint API
Public Function GetDocumentID() As String
    If mSharePointStatus Then
        GetDocumentID = GetSharePointDocumentID()
    Else
        GetDocumentID = "NOT_AVAILABLE"
    End If
End Function

Public Function GetDocumentLocation() As String
    If mSharePointStatus Then
        GetDocumentLocation = GetDisplayablePath()
    Else
        GetDocumentLocation = ActiveWorkbook.FullName
    End If
End Function

Public Function IsOnSharePoint() As Boolean
    If mSharePointStatus Then
        IsOnSharePoint = IsSharePointDocument()
    Else
        IsOnSharePoint = False
    End If
End Function

' Error Handling API
Public Sub HandleVBAError(procedureName As String, Optional moduleName As String = "")
    If mErrorHandlerStatus Then
        HandleError procedureName, moduleName
    Else
        ' Fallback error handling
        LogError "vba_error", Err.Number, "Procedure: " & procedureName & " | Module: " & moduleName & " | Error: " & Err.description
    End If
End Sub

Public Sub HandleCustomError(errorMessage As String, procedureName As String)
    If mErrorHandlerStatus Then
        LogError "custom_error", 9999, "Procedure: " & procedureName & " | Error: " & errorMessage
    Else
        LogError "custom_error", 9999, "Procedure: " & procedureName & " | Error: " & errorMessage
    End If
End Sub

' ============================================================================
' INTEGRATED WORKFLOWS
' ============================================================================

Public Function HandleErrorWithTicketOption(title As String, errorMessage As String, procedureName As String) As String
    ' Integrated workflow: Error handling with automatic ticket option
    
    ' Log the error
    LogError "integrated_error", 0, "Procedure: " & procedureName & " | Message: " & errorMessage
    
    ' Show error message with ticket option
    Dim result As String
    result = SYS_MessageBox.ShowErrorMessage(title, errorMessage, True)
    
    ' Handle ticket creation if requested
    If result = "CREATE_TICKET" Then
        Dim ticketResult As String
        ticketResult = CreateErrorTicket(errorMessage, 0)
        
        If ticketResult <> "SYSTEM_NOT_AVAILABLE" And ticketResult <> "CANCELLED" Then
            SYS_MessageBox.ShowInfoMessage "Ticket Created", "Support ticket " & ticketResult & " has been created and sent to the support team."
        End If
    End If
    
    HandleErrorWithTicketOption = result
End Function

' ============================================================================
' SYSTEM HEALTH AND DIAGNOSTICS
' ============================================================================

Public Function PerformSystemHealthCheck() As Object
    ' Perform comprehensive system health check
    
    LogInfo "health_check_start", "Performing system health check"
    
    Dim healthReport As Object
    Set healthReport = CreateObject("Scripting.Dictionary")
    
    ' Core system health
    healthReport("core_system") = GetCoreSystemHealth()
    
    ' Module health
    healthReport("logger") = GetLoggerHealth()
    healthReport("error_handler") = GetErrorHandlerHealth()
    healthReport("sharepoint") = GetSharePointHealth()
    healthReport("messagebox") = GetMessageBoxHealth()
    healthReport("ticket_system") = GetTicketSystemHealth()
    
    ' Overall system metrics
    healthReport("uptime_minutes") = DateDiff("n", mStartupTime, Now)
    healthReport("modules_loaded") = GetLoadedModulesCount()
    healthReport("system_mode") = SYS_CoreSystem.GetSystemModeString(mSystemMode)
    healthReport("memory_status") = "OK" ' Placeholder
    
    ' Calculate overall health score
    healthReport("overall_health") = CalculateOverallHealth(healthReport)
    
    LogInfo "health_check_complete", "Health check completed - Score: " & healthReport("overall_health")
    
    Set PerformSystemHealthCheck = healthReport
End Function

Private Function GetCoreSystemHealth() As Object
    Dim health As Object
    Set health = CreateObject("Scripting.Dictionary")
    
    ' Example checks - these should be replaced with actual health check logic
    health("status") = "OK"
    health("last_restart") = mStartupTime
    health("uptime_minutes") = DateDiff("n", mStartupTime, Now)
    
    GetCoreSystemHealth = health
End Function

Private Function GetLoggerHealth() As Object
    Dim health As Object
    Set health = CreateObject("Scripting.Dictionary")
    
    ' Logger health check - ensure logger is initialized and functional
    health("status") = IIf(mLoggerStatus, "OK", "Failed")
    health("log_level") = IIf(mLoggerStatus, SYS_Logger.GetCurrentLogLevel(), "N/A")
    
    GetLoggerHealth = health
End Function

Private Function GetErrorHandlerHealth() As Object
    Dim health As Object
    Set health = CreateObject("Scripting.Dictionary")
    
    ' Error handler health check - ensure error handler is initialized
    health("status") = IIf(mErrorHandlerStatus, "OK", "Failed")
    
    GetErrorHandlerHealth = health
End Function

Private Function GetSharePointHealth() As Object
    Dim health As Object
    Set health = CreateObject("Scripting.Dictionary")
    
    ' SharePoint health check - ensure SharePoint module is loaded
    health("status") = IIf(mSharePointStatus, "OK", "Not Available")
    
    GetSharePointHealth = health
End Function

Private Function GetMessageBoxHealth() As Object
    Dim health As Object
    Set health = CreateObject("Scripting.Dictionary")
    
    ' MessageBox health check - always available if module is loaded
    health("status") = IIf(mMessageBoxStatus, "OK", "Not Available")
    
    GetMessageBoxHealth = health
End Function

Private Function GetTicketSystemHealth() As Object
    Dim health As Object
    Set health = CreateObject("Scripting.Dictionary")
    
    ' Ticket system health check - always available
    health("status") = IIf(mTicketSystemStatus, "OK", "Not Available")
    
    GetTicketSystemHealth = health
End Function

Private Function CalculateOverallHealth(healthReport As Object) As Integer
    ' Calculate an overall health score based on individual component health
    Dim score As Integer
    score = 100 ' Start with a perfect score
    
    ' Deduct points for each component that is not OK
    If healthReport("core_system")("status") <> "OK" Then score = score - 20
    If healthReport("logger")("status") <> "OK" Then score = score - 20
    If healthReport("error_handler")("status") <> "OK" Then score = score - 20
    If healthReport("sharepoint")("status") <> "OK" Then score = score - 20
    If healthReport("messagebox")("status") <> "OK" Then score = score - 10
    If healthReport("ticket_system")("status") <> "OK" Then score = score - 10
    
    CalculateOverallHealth = score
End Function

' ============================================================================
' DEBUGGING AND DEVELOPMENT UTILITIES
' ============================================================================

Public Function GetLoadedModulesCount() As Long
    ' Get the count of successfully loaded modules
    GetLoadedModulesCount = mModulesLoaded.count
End Function

' Les fonctions GetSystemModeString et GetCurrentLogLevel ont été supprimées car elles sont dupliquées
' Utiliser SYS_CoreSystem.GetSystemModeString et SYS_Logger.GetCurrentLogLevel à la place

' ============================================================================
' ERROR HANDLING SETUP
' ============================================================================

Private Sub SetupIntegratedErrorHandling()
    ' Configure la gestion d'erreurs intégrée à travers tous les modules
    
    On Error Resume Next
    
    ' Initialiser le gestionnaire d'erreurs
    If Not mErrorHandlerStatus Then
        mErrorHandlerStatus = InitializeErrorHandler()
        If mErrorHandlerStatus Then
            LogInfo "error_handler_setup", "Error handler initialized successfully"
        Else
            LogWarning "error_handler_setup_failed", "Failed to initialize error handler"
            Exit Sub
        End If
    End If
    
    ' Configurer les options de gestion d'erreurs
    EnableAutoRecovery ' Activer la récupération automatique
    EnableUserMessages ' Activer les messages utilisateur
    
    ' Configurer les hooks d'erreur globaux
    Application.EnableEvents = True
    
    ' Configurer les gestionnaires d'erreurs spécifiques aux modules
    SetupModuleErrorHandlers
    
    LogInfo "error_handling_setup", "Integrated error handling setup completed"
    
    On Error GoTo 0
End Sub

Private Sub SetupModuleErrorHandlers()
    ' Configure les gestionnaires d'erreurs spécifiques aux modules
    
    ' Logger
    If mLoggerStatus Then
        PushErrorHandler "SYS_Logger"
        LogInfo "error_handler_setup", "Logger error handler configured"
    End If
    
    ' MessageBox
    If mMessageBoxStatus Then
        PushErrorHandler "SYS_MessageBox"
        LogInfo "error_handler_setup", "MessageBox error handler configured"
    End If
    
    ' SharePoint
    If mSharePointStatus Then
        PushErrorHandler "SYS_SharePointIntegration"
        LogInfo "error_handler_setup", "SharePoint error handler configured"
    End If
    
    ' Ticket System
    If mTicketSystemStatus Then
        PushErrorHandler "SYS_TicketSystem"
        LogInfo "error_handler_setup", "Ticket system error handler configured"
    End If
End Sub

