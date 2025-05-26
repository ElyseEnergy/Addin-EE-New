' ============================================================================
' SYS_Logger - Centralized Logging System
' Elyse Energy VBA Ecosystem - Logging Component
' Requires: SYS_CoreSystem, SYS_ErrorHandler
' ============================================================================

Option Explicit

' ============================================================================
' MODULE INFORMATION
' ============================================================================
Private Const MODULE_NAME As String = "SYS_Logger"
Private Const ERROR_HANDLER_LABEL As String = "ErrorHandler"

' ============================================================================
' MODULE DEPENDENCIES
' ============================================================================
' This module requires:
' - SYS_CoreSystem (enums, constants, utilities)
' - SYS_ErrorHandler (error handling functions)

' ============================================================================
' LOGGING STATE VARIABLES
' ============================================================================

Private mLogBuffer As Collection
Private mLastFlushTime As Date
Private mLoggerInitialized As Boolean
Private mCurrentLogLevel As LogLevel
Private mLoggerActive As Boolean
Private mLogToFile As Boolean
Private Const LOG_FILE_PATH As String = "elyse_log.txt"

' ============================================================================
' LOGGER INITIALIZATION
' ============================================================================

Public Function InitializeLogger(Optional logLevel As LogLevel = INFO_LEVEL) As Boolean
    ' Initialize the logging system
    On Error GoTo ErrorHandler
    
    ' Check core system dependency
    If Not IsSystemInitialized() Then
        InitializeLogger = False
        Exit Function
    End If
    
    ' Set logging configuration
    mCurrentLogLevel = logLevel
    mLastFlushTime = Now
    mLoggerActive = True
    mLogToFile = True ' Can be made configurable
    
    ' Initialize log buffer
    Set mLogBuffer = New Collection
    
    ' Mark as initialized
    mLoggerInitialized = True
    
    ' Log the initialization
    LogEvent "logger_init", "Logger system initialized successfully", INFO_LEVEL
    
    InitializeLogger = True
    Exit Function
    
ErrorHandler:
    InitializeLogger = False
End Function

Public Sub ShutdownLogger()
    ' Clean shutdown of logging system
    
    If Not mLoggerInitialized Then Exit Sub
    
    ' Log shutdown
    LogEvent "logger_shutdown", "Logger system shutting down", INFO_LEVEL
    
    ' Flush remaining logs
    FlushLogBuffer True
    
    ' Cleanup
    Set mLogBuffer = Nothing
    mLoggerInitialized = False
End Sub

' ============================================================================
' CORE LOGGING FUNCTIONS WITH RAGIC INTEGRATION
' ============================================================================

Public Sub LogEvent(actionCode As String, message As String, level As LogLevel, _
    Optional ByVal procedureName As String = "", Optional ByVal moduleName As String = "", _
    Optional ByVal errorCode As Long = 0, Optional ByRef errorCtx As SYS_ErrorHandler.ErrorContext = Nothing)
    ' Central logging function that dispatches to specific log functions based on level
    
    If Not mLoggerInitialized Then Exit Sub
    If Not ShouldLog(level) Then Exit Sub
    
    Select Case level
        Case ERROR_LEVEL
            If Not errorCtx Is Nothing Then
                LogError actionCode, errorCode, message, IIf(Len(procedureName) > 0, procedureName, errorCtx.ProcedureName), _
                    IIf(Len(moduleName) > 0, moduleName, errorCtx.ModuleName), errorCtx
            Else
                LogError actionCode, errorCode, message, procedureName, moduleName
            End If
            
        Case CRITICAL_LEVEL
            LogCritical actionCode, errorCode, message, procedureName, moduleName
            
        Case WARNING_LEVEL
            LogWarning actionCode, message, procedureName, moduleName
            
        Case INFO_LEVEL
            LogInfo actionCode, message, procedureName, moduleName
            
        Case DEBUG_LEVEL
            LogDebug actionCode, message, procedureName, moduleName
            
        Case Else
            LogInfo actionCode, message, procedureName, moduleName ' Default to INFO if level is unknown
    End Select
End Sub

Public Sub LogError(actionCode As String, errorCode As Long, message As String, _
    Optional ByVal procedureName As String = "", Optional ByVal moduleName As String = "", _
    Optional ByRef errorCtx As SYS_ErrorHandler.ErrorContext = Nothing)
    ' Specialized logging for errors with error context support
    
    Dim logMessage As String
    logMessage = "ERROR [" & actionCode & "] "
    If moduleName <> "" And procedureName <> "" Then
        logMessage = logMessage & moduleName & "." & procedureName & " "
    ElseIf procedureName <> "" Then
        logMessage = logMessage & procedureName & " "
    End If
    logMessage = logMessage & "(Code: " & errorCode & "): " & message

    ' Log through main system with context
    If mLoggerActive Then
        PrintToImmediate logMessage
        If mLogToFile Then WriteToLogFile "ERROR", logMessage
    End If
    
    ' Log to Ragic with error context
    Call LogToRagic("ERROR", actionCode, message, errorCode, procedureName, moduleName)
End Sub

Public Sub LogInfo(actionCode As String, message As String, Optional ByVal procedureName As String = "", Optional ByVal moduleName As String = "")
    Dim logMessage As String
    logMessage = "INFO  [" & actionCode & "] "
    If moduleName <> "" And procedureName <> "" Then
        logMessage = logMessage & moduleName & "." & procedureName
    ElseIf procedureName <> "" Then
        logMessage = logMessage & procedureName
    End If
    logMessage = logMessage & ": " & message
    
    If mLoggerActive Then
        Select Case mCurrentLogLevel
            Case DEBUG_LEVEL, INFO_LEVEL
                PrintToImmediate logMessage
                If mLogToFile Then WriteToLogFile "INFO", logMessage
        End Select
    End If
    
    Call LogToRagic("INFO", actionCode, message, , procedureName, moduleName) 
End Sub

Public Sub LogDebug(actionCode As String, message As String, Optional ByVal procedureName As String = "", Optional ByVal moduleName As String = "")
    Dim logMessage As String
    logMessage = "DEBUG [" & actionCode & "] "
     If moduleName <> "" And procedureName <> "" Then
        logMessage = logMessage & moduleName & "." & procedureName
    ElseIf procedureName <> "" Then
        logMessage = logMessage & procedureName
    End If
    logMessage = logMessage & ": " & message
    
    If mLoggerActive Then
        Select Case mCurrentLogLevel
            Case DEBUG_LEVEL
                PrintToImmediate logMessage
                If mLogToFile Then WriteToLogFile "DEBUG", logMessage
        End Select
    End If
    
    Call LogToRagic("DEBUG", actionCode, message, , procedureName, moduleName)
End Sub

Public Sub LogWarning(actionCode As String, message As String, Optional ByVal procedureName As String = "", Optional ByVal moduleName As String = "")
    Dim logMessage As String
    logMessage = "WARN  [" & actionCode & "] "
    If moduleName <> "" And procedureName <> "" Then
        logMessage = logMessage & moduleName & "." & procedureName
    ElseIf procedureName <> "" Then
        logMessage = logMessage & procedureName
    End If
    logMessage = logMessage & ": " & message
    
    If mLoggerActive Then
        Select Case mCurrentLogLevel
            Case DEBUG_LEVEL, INFO_LEVEL, WARNING_LEVEL
                PrintToImmediate logMessage
                If mLogToFile Then WriteToLogFile "WARNING", logMessage
        End Select
    End If
    
    Call LogToRagic("WARNING", actionCode, message, , procedureName, moduleName)
End Sub

Public Sub LogCritical(actionCode As String, errorCode As Long, message As String, Optional ByVal procedureName As String = "", Optional ByVal moduleName As String = "")
    Dim logMessage As String
    logMessage = "CRITICAL [" & actionCode & "] "
    If moduleName <> "" And procedureName <> "" Then
        logMessage = logMessage & moduleName & "." & procedureName
    ElseIf procedureName <> "" Then
        logMessage = logMessage & procedureName
    End If
    logMessage = logMessage & " (Code: " & errorCode & "): " & message
    
    If mLoggerActive Then
        PrintToImmediate logMessage 
        If mLogToFile Then WriteToLogFile "CRITICAL", logMessage
    End If
    
    Dim criticalDetails As String
    criticalDetails = message
    
    Dim tempCtx As ErrorContext ' Create a context for critical errors too
    tempCtx.ErrorNumber = errorCode
    tempCtx.ErrorDescription = message
    tempCtx.ProcedureName = procedureName
    tempCtx.ModuleName = moduleName
    tempCtx.Severity = "CRITICAL" ' Explicitly set
    
    Call LogToRagic("CRITICAL", actionCode, criticalDetails, errorCode, procedureName, moduleName)
End Sub

Public Sub LogFunctionCall(procedureName As String, Optional ByVal moduleName As String = "", Optional params As String = "")
    Dim logMessage As String
    logMessage = "CALL  ["
    If moduleName <> "" Then logMessage = logMessage & moduleName & "."
    logMessage = logMessage & procedureName & "] Called"
    If params <> "" Then logMessage = logMessage & " with params: " & params
    
    If mLoggerActive Then
        If mCurrentLogLevel <= DEBUG_LEVEL Then
            PrintToImmediate logMessage
            If mLogToFile Then WriteToLogFile "CALL", logMessage
        End If
    End If
    
    Dim action As String
    action = "FunctionCall"
    If moduleName <> "" Then action = action & ":" & moduleName
    action = action & ":" & procedureName
    
    Call LogToRagic("DEBUG", action, "Params: " & params, , procedureName, moduleName)
End Sub

Public Sub LogUserAction(actionCode As String, description As String, Optional ByVal controlName As String = "")
    Dim logMessage As String
    logMessage = "USER  [" & actionCode & "] " & description
    If controlName <> "" Then logMessage = logMessage & " (Control: " & controlName & ")"
    
    If mLoggerActive Then
        PrintToImmediate logMessage 
        If mLogToFile Then WriteToLogFile "USER", logMessage
    End If
    
    Dim action As String
    action = "UserAction:" & actionCode
    
    Dim details As String
    details = description
    If controlName <> "" Then details = details & " (Control: " & controlName & ")"
    
    Call LogToRagic("INFO", action, details, , procedureName, moduleName)
End Sub

' ============================================================================
' LOG ENTRY CREATION AND FORMATTING
' ============================================================================

Private Function CreateLogEntry(action As String, details As String, level As LogLevel, includeContext As Boolean) As Object
    ' Create a complete log entry with all metadata
    
    Dim logEntry As Object
    Set logEntry = CreateObject("Scripting.Dictionary")
    
    ' Basic log information
    logEntry("timestamp") = FormatTimestamp(True)
    logEntry("session_id") = GetSessionID()
    logEntry("user") = GetUserIdentity()
    logEntry("action") = action
    logEntry("details") = TruncateString(details, 500) ' Limit detail length
    logEntry("level") = GetLogLevelString(level)
    
    ' Include contextual information if requested
    If includeContext Then
        AddContextualInformation logEntry
    End If
    
    Set CreateLogEntry = logEntry
End Function

Private Sub AddContextualInformation(logEntry As Object)
    ' Add rich contextual information to log entry
    
    ' System context
    logEntry("excel_version") = Application.Version
    logEntry("computer_name") = Environ("COMPUTERNAME")
    logEntry("user_domain") = Environ("USERDOMAIN")
    
    ' Excel context
    logEntry("workbook_name") = GetCurrentWorkbookName()
    logEntry("active_sheet") = GetActiveSheetName()
    logEntry("selected_range") = GetSelectedRangeAddress()
    
    ' File context (will be populated by SharePoint module if available)
    logEntry("file_location") = "pending_sharepoint_check"
    logEntry("sharepoint_doc_id") = "pending_sharepoint_check"
    logEntry("sharepoint_url") = "pending_sharepoint_check"
End Sub

' ============================================================================
' LOG BUFFER MANAGEMENT
' ============================================================================

Private Sub AddToLogBuffer(logEntry As Object)
    ' Add log entry to buffer with overflow protection
    
    ' Initialize buffer if needed
    If mLogBuffer Is Nothing Then
        Set mLogBuffer = New Collection
    End If
    
    ' Add entry
    mLogBuffer.Add logEntry
    
    ' Prevent buffer overflow
    If mLogBuffer.Count > LOG_BUFFER_SIZE * 2 Then
        FlushLogBuffer True ' Force flush with older entries
    End If
End Sub

Public Sub FlushLogBuffer(Optional forceFlush As Boolean = False)
    ' Send buffered log entries to API
    
    If mLogBuffer Is Nothing Then Exit Sub
    If mLogBuffer.Count = 0 Then Exit Sub
    
    ' Check if enough time has passed since last flush (unless forced)
    If Not forceFlush Then
        If DateDiff("s", mLastFlushTime, Now) < 10 Then Exit Sub ' Wait at least 10 seconds
    End If
    
    ' Send the batch
    SendLogBatch mLogBuffer
    
    ' Clear buffer and update flush time
    Set mLogBuffer = New Collection
    mLastFlushTime = Now
End Sub

' ============================================================================
' API COMMUNICATION
' ============================================================================

Private Sub SendLogBatch(logCollection As Collection)
    ' Send batch of log entries to the API
    On Error GoTo ErrorHandler
    
    ' Skip if no logs to send
    If logCollection.Count = 0 Then Exit Sub
    
    ' Create HTTP request
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Configure request
    http.Open "POST", API_BASE_URL & API_LOGS_ENDPOINT, True ' Async
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Authorization", API_TOKEN
    http.SetRequestHeader "User-Agent", "ElyseLogger/1.0"
    
    ' Convert logs to JSON
    Dim jsonPayload As String
    jsonPayload = ConvertLogBatchToJSON(logCollection)
    
    ' Send request (non-blocking)
    http.Send jsonPayload
    
    ' Don't wait for response to avoid blocking Excel
    Exit Sub
    
ErrorHandler:
    ' Silently handle API errors to avoid disrupting user workflow
    ' Could optionally store failed logs for retry later
End Sub

Private Function ConvertLogBatchToJSON(logCollection As Collection) As String
    ' Convert log collection to JSON format
    
    Dim json As String
    json = "{""logs"":["
    
    Dim i As Integer
    For i = 1 To logCollection.Count
        If i > 1 Then json = json & ","
        json = json & ConvertSingleLogToJSON(logCollection(i))
    Next i
    
    json = json & "],""batch_info"":{""count"":" & logCollection.Count & ",""timestamp"":""" & FormatTimestamp() & """}}"
    
    ConvertLogBatchToJSON = json
End Function

Private Function ConvertSingleLogToJSON(logEntry As Object) As String
    ' Convert single log entry to JSON
    
    Dim json As String
    json = "{"
    
    ' Add all log entry fields
    Dim keys As Variant
    keys = logEntry.Keys
    
    Dim i As Integer
    For i = 0 To UBound(keys)
        If i > 0 Then json = json & ","
        json = json & """" & keys(i) & """:""" & EscapeJSON(CStr(logEntry(keys(i)))) & """"
    Next i
    
    json = json & "}"
    
    ConvertSingleLogToJSON = json
End Function

' ============================================================================
' LOG LEVEL MANAGEMENT
' ============================================================================

Public Sub SetLogLevel(level As LogLevel)
    ' Set current logging level
    mCurrentLogLevel = level
    LogEvent "log_level_changed", "Log level set to: " & GetLogLevelString(level), INFO_LEVEL
End Sub

Public Function GetCurrentLogLevel() As LogLevel
    ' Get current logging level
    GetCurrentLogLevel = mCurrentLogLevel
End Function

Public Function ShouldLog(level As LogLevel) As Boolean
    ' Check if a log level should be recorded
    ShouldLog = (level >= mCurrentLogLevel And mLoggerInitialized)
End Function

' ============================================================================
' DEBUGGING AND DIAGNOSTICS
' ============================================================================

Public Function GetLoggerStatus() As Object
    ' Get comprehensive logger status
    Dim status As Object
    Set status = CreateObject("Scripting.Dictionary")
    
    status("initialized") = mLoggerInitialized
    status("current_log_level") = GetLogLevelString(mCurrentLogLevel)
    status("buffer_count") = IIf(mLogBuffer Is Nothing, 0, mLogBuffer.Count)
    status("last_flush_time") = Format(mLastFlushTime, "yyyy-mm-dd hh:nn:ss")
    
    Set GetLoggerStatus = status
End Function

Public Sub DumpBufferToDebug()
    ' Dump current buffer contents to debug window (development only)
    
    If Not IsDebugMode() Then Exit Sub
    If mLogBuffer Is Nothing Then Exit Sub
    
    Debug.Print "=== LOG BUFFER DUMP ==="
    Debug.Print "Buffer Count: " & mLogBuffer.Count
    Debug.Print "========================"
    
    Dim i As Integer
    For i = 1 To mLogBuffer.Count
        Dim logEntry As Object
        Set logEntry = mLogBuffer(i)
        
        Debug.Print i & ": " & logEntry("timestamp") & " | " & logEntry("level") & " | " & logEntry("action") & " | " & logEntry("details")
    Next i
    
    Debug.Print "=== END DUMP ==="
End Sub

' ============================================================================
' PUBLIC CONVENIENCE FUNCTIONS
' ============================================================================

' Note: Removed duplicate simplified logging functions in favor of the complete versions above

' ============================================================================
' TIMER-BASED AUTO FLUSH
' ============================================================================

Public Sub EnableAutoFlush()
    ' Enable automatic buffer flushing every 60 seconds
    Application.OnTime Now + TimeValue("00:01:00"), "AutoFlushCallback"
End Sub

Public Sub AutoFlushCallback()
    ' Auto-flush callback (must be public for OnTime)
    
    If mLoggerInitialized Then
        FlushLogBuffer
        
        ' Schedule next auto-flush
        Application.OnTime Now + TimeValue("00:01:00"), "AutoFlushCallback"
    End If
End Sub

' ============================================================================
' RAGIC LOGGING INTEGRATION
' ============================================================================

Private Sub LogToRagic(logLevel As String, action As String, details As String, Optional errorCode As Long = 0, Optional procedureName As String = "", Optional moduleName As String = "")
    If Not gEnableRagicLogging Then Exit Sub
    If RAGIC_LOG_API_KEY = "" Or RAGIC_LOG_API_KEY = "YOUR_ACTUAL_RAGIC_API_KEY" Then
        Debug.Print "Ragic API Key not configured. Skipping Ragic log."
        Exit Sub
    End If

    Dim http As Object
    Dim payload As String
    Dim fieldData As Object ' Scripting.Dictionary

    On Error GoTo RagicLogErrorHandler

    Set http = CreateObject("MSXML2.XMLHTTP")
    Set fieldData = CreateObject("Scripting.Dictionary")

    ' Populate common fields
    fieldData(RAGIC_FIELD_TIMESTAMP) = Format(Now, "yyyy/MM/dd HH:mm:ss")
    fieldData(RAGIC_FIELD_USER) = GetCurrentUserIdentifier()
    fieldData(RAGIC_FIELD_ACTION) = Left(action, 255)
    fieldData(RAGIC_FIELD_DETAILS) = details
    fieldData(RAGIC_FIELD_LEVEL) = logLevel
    fieldData(RAGIC_FIELD_SESSION_ID) = GetSessionID()

    ' Add system context
    fieldData(RAGIC_FIELD_EXCEL_VERSION) = Application.Version
    fieldData(RAGIC_FIELD_USER_DOMAIN) = Environ("USERDOMAIN")
    fieldData(RAGIC_FIELD_COMPUTER_NAME) = Environ("COMPUTERNAME")

    ' Add Excel context if available
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then
        fieldData(RAGIC_FIELD_WORKBOOK_NAME) = ActiveWorkbook.Name
        fieldData(RAGIC_FIELD_ACTIVE_SHEET) = ActiveSheet.Name
        fieldData(RAGIC_FIELD_FILE_LOCATION) = ActiveWorkbook.FullName
        If TypeName(Selection) = "Range" Then
            fieldData(RAGIC_FIELD_SELECTED_RANGE) = Selection.Address
        End If
    End If
    On Error GoTo RagicLogErrorHandler

    ' If error context is provided, add rich error details
    If errorCode <> 0 Then
        ' Add basic error info
        If RAGIC_FIELD_ERROR_NUMBER <> "" Then fieldData(RAGIC_FIELD_ERROR_NUMBER) = errorCode
        If RAGIC_FIELD_ERROR_SOURCE <> "" Then fieldData(RAGIC_FIELD_ERROR_SOURCE) = Left(Err.Source, 255)
        If RAGIC_FIELD_ERROR_DESCRIPTION <> "" Then fieldData(RAGIC_FIELD_ERROR_DESCRIPTION) = Left(Err.Description, 1000)
        
        ' Add location info
        If RAGIC_FIELD_MODULE_NAME <> "" Then fieldData(RAGIC_FIELD_MODULE_NAME) = Left(moduleName, 255)
        If RAGIC_FIELD_PROCEDURE_NAME <> "" Then fieldData(RAGIC_FIELD_PROCEDURE_NAME) = Left(procedureName, 255)
        If RAGIC_FIELD_LINE_NUMBER <> "" And Err.Erl > 0 Then fieldData(RAGIC_FIELD_LINE_NUMBER) = Err.Erl

        ' Add severity and recovery info if available
        If RAGIC_FIELD_ERROR_SEVERITY <> "" Then fieldData(RAGIC_FIELD_ERROR_SEVERITY) = "CRITICAL"
        If RAGIC_FIELD_RECOVERY_ATTEMPTED <> "" Then fieldData(RAGIC_FIELD_RECOVERY_ATTEMPTED) = "Yes"
        If RAGIC_FIELD_TICKET_CREATED <> "" Then fieldData(RAGIC_FIELD_TICKET_CREATED) = "No"
    End If

    ' Build payload string
    payload = ""
    Dim key As Variant
    For Each key In fieldData.Keys
        If fieldData(key) <> "" Then ' Only send fields with values
            If payload <> "" Then payload = payload & "&"
            payload = payload & key & "=" & EncodeURL(CStr(fieldData(key)))
        End If
    Next key

    ' Add APIKey to the payload for authentication
    If RAGIC_LOG_API_KEY <> "" And RAGIC_LOG_API_KEY <> "YOUR_ACTUAL_RAGIC_API_KEY" Then
        If payload <> "" Then payload = payload & "&"
        payload = payload & "APIKey=" & EncodeURL(RAGIC_LOG_API_KEY)
    End If

    ' Configure and send request
    Dim ragicPostUrl As String
    ragicPostUrl = RAGIC_LOG_API_URL
    If InStr(1, ragicPostUrl, "?") = 0 Then
        ragicPostUrl = ragicPostUrl & "?v=3&api=true" ' MODIFIED: Added v=3 and ensured api=true
    Else
        If InStr(1, ragicPostUrl, "api=") = 0 Then
            ragicPostUrl = ragicPostUrl & "&api=true" ' Ensure api=true if other params exist
        End If
        If InStr(1, ragicPostUrl, "v=") = 0 Then
            ragicPostUrl = ragicPostUrl & "&v=3" ' Ensure v=3 if other params exist
        End If
    End If
    
    http.Open "POST", ragicPostUrl, False ' Synchronous for reliability
    ' http.setRequestHeader "Authorization", "Basic " & EncodeBase64(RAGIC_LOG_API_KEY & ":") ' REMOVED: Using APIKey in payload instead
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send payload

    ' Handle response
    If http.Status >= 200 And http.Status < 300 Then
        Debug.Print "Log entry successfully sent to Ragic. Action: " & action
    Else
        Debug.Print "Failed to send log to Ragic. Status: " & http.Status & " - " & http.statusText
        Debug.Print "Response: " & http.responseText
        Debug.Print "Payload: " & payload
    End If

    Set http = Nothing
    Set fieldData = Nothing
    Exit Sub

RagicLogErrorHandler:
    Debug.Print "Error in LogToRagic: " & Err.Number & " - " & Err.Description
    ' Continue execution - don't let logging failures disrupt main flow
    Resume Next
End Sub

' Helper for URL encoding
Private Function EncodeURL(str As String) As String
    Dim ScriptControl As Object
    Set ScriptControl = CreateObject("ScriptControl")
    ScriptControl.Language = "JScript"
    EncodeURL = ScriptControl.CodeObject.encodeURIComponent(str)
    Set ScriptControl = Nothing
End Function

