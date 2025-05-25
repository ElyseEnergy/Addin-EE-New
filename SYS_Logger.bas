' ============================================================================
' ElyseLogger_Module - Centralized Logging System
' Elyse Energy VBA Ecosystem - Logging Component
' Requires: ElyseCore_System
' ============================================================================

Option Explicit

' ============================================================================
' MODULE DEPENDENCIES
' ============================================================================
' This module requires ElyseCore_System to be loaded first
' Dependencies: ElyseCore_System (enums, constants, utilities)

' ============================================================================
' LOGGING STATE VARIABLES
' ============================================================================

Private mLogBuffer As Collection
Private mLastFlushTime As Date
Private mLoggerInitialized As Boolean
Private mCurrentLogLevel As LogLevel

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
' CORE LOGGING FUNCTIONS
' ============================================================================

Public Sub LogEvent(action As String, details As String, Optional level As LogLevel = INFO_LEVEL, Optional includeContext As Boolean = True)
    ' Main logging function
    
    ' Check if we should log this level
    If level < mCurrentLogLevel Then Exit Sub
    
    ' Ensure logger is initialized
    If Not mLoggerInitialized Then
        If Not InitializeLogger() Then Exit Sub
    End If
    
    ' Create log entry
    Dim logEntry As Object
    Set logEntry = CreateLogEntry(action, details, level, includeContext)
    
    ' Add to buffer
    AddToLogBuffer logEntry
    
    ' Auto-flush if critical or buffer full
    If level >= ERROR_LEVEL Or mLogBuffer.Count >= LOG_BUFFER_SIZE Then
        FlushLogBuffer
    End If
End Sub

Public Sub LogRibbonAction(buttonId As String, Optional additionalInfo As String = "")
    ' Specialized logging for ribbon button clicks
    Dim details As String
    details = "Button ID: " & buttonId
    
    If additionalInfo <> "" Then
        details = details & " | Info: " & additionalInfo
    End If
    
    LogEvent "ribbon_click", details, INFO_LEVEL
End Sub

Public Sub LogFunctionCall(functionName As String, Optional parameters As String = "", Optional executionTime As Double = 0)
    ' Specialized logging for function calls
    Dim details As String
    details = "Function: " & functionName
    
    If parameters <> "" Then
        details = details & " | Parameters: " & parameters
    End If
    
    If executionTime > 0 Then
        details = details & " | Execution time: " & Format(executionTime, "0.000") & "s"
    End If
    
    LogEvent "function_call", details, DEBUG_LEVEL
End Sub

Public Sub LogUserAction(actionType As String, description As String, Optional targetObject As String = "")
    ' Specialized logging for user actions
    Dim details As String
    details = "Action: " & actionType & " | Description: " & description
    
    If targetObject <> "" Then
        details = details & " | Target: " & targetObject
    End If
    
    LogEvent "user_action", details, INFO_LEVEL
End Sub

Public Sub LogError(errorSource As String, errorNumber As Long, errorDescription As String, Optional stackTrace As String = "")
    ' Specialized logging for errors
    Dim details As String
    details = "Source: " & errorSource & " | Error: " & errorNumber & " | Description: " & errorDescription
    
    If stackTrace <> "" Then
        details = details & " | Stack: " & stackTrace
    End If
    
    LogEvent "error", details, ERROR_LEVEL
End Sub

Public Sub LogPerformance(operationName As String, duration As Double, Optional additionalMetrics As String = "")
    ' Specialized logging for performance metrics
    Dim details As String
    details = "Operation: " & operationName & " | Duration: " & Format(duration, "0.000") & "s"
    
    If additionalMetrics <> "" Then
        details = details & " | Metrics: " & additionalMetrics
    End If
    
    LogEvent "performance", details, DEBUG_LEVEL
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

Public Sub LogInfo(action As String, details As String)
    ' Convenience function for info logging
    LogEvent action, details, INFO_LEVEL
End Sub

Public Sub LogDebug(action As String, details As String)
    ' Convenience function for debug logging
    LogEvent action, details, DEBUG_LEVEL
End Sub

Public Sub LogWarning(action As String, details As String)
    ' Convenience function for warning logging
    LogEvent action, details, WARNING_LEVEL
End Sub

Public Sub LogCritical(action As String, details As String)
    ' Convenience function for critical logging
    LogEvent action, details, CRITICAL_LEVEL
End Sub

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

Private Sub LogToRagic(logLevel As String, action As String, details As String, Optional errorCtx As ErrorContext = Nothing)
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
    fieldData(RAGIC_FIELD_TIMESTAMP) = Format(Now, "yyyy/MM/dd HH:mm:ss") ' Ragic often prefers yyyy/MM/dd for dates
    fieldData(RAGIC_FIELD_USER) = GetCurrentUserIdentifier()
    fieldData(RAGIC_FIELD_ACTION) = Left(action, 255) ' Ensure within typical text limits
    fieldData(RAGIC_FIELD_DETAILS) = details
    fieldData(RAGIC_FIELD_LEVEL) = logLevel
    fieldData(RAGIC_FIELD_SESSION_ID) = GetSessionID()
    fieldData(RAGIC_FIELD_EXCEL_VERSION) = Application.Version
    
    On Error Resume Next ' Gracefully handle if workbook/sheet properties are not available
    fieldData(RAGIC_FIELD_WORKBOOK_NAME) = ActiveWorkbook.Name
    fieldData(RAGIC_FIELD_ACTIVE_SHEET) = ActiveSheet.Name
    If TypeName(Selection) = "Range" Then
        fieldData(RAGIC_FIELD_SELECTED_RANGE) = Selection.Address
    End If
    fieldData(RAGIC_FIELD_FILE_LOCATION) = ActiveWorkbook.FullName
    
    ' SharePoint fields integration
    If ElyseCore_System.IsFeatureEnabled("SharePointIntegration") And Not ActiveWorkbook Is Nothing Then
        ' Check if ElyseSharePoint_Integration module is available and function exists
        ' This requires ElyseSharePoint_Integration to be structured to provide these details.
        ' For now, we assume functions GetSharePointDocIDForWorkbook and GetSharePointUrlForWorkbook exist in ElyseSharePoint_Integration
        Dim spDocId As String
        Dim spUrl As String
        
        ' Use Application.Run to safely call functions in another module if it's loaded
        ' This avoids compile errors if the module isn't present, though it should be.
        On Error Resume Next ' For Application.Run calls
        spDocId = Application.Run("ElyseSharePoint_Integration.GetSharePointDocIDForWorkbook", ActiveWorkbook)
        spUrl = Application.Run("ElyseSharePoint_Integration.GetSharePointUrlForWorkbook", ActiveWorkbook)
        On Error GoTo RagicLogErrorHandler ' Restore main error handler after Application.Run attempts
        
        If spDocId <> "" And spDocId <> "Error" Then ' Check for actual value or error string from function
            fieldData(RAGIC_FIELD_SP_DOC_ID) = spDocId
        Else
            fieldData(RAGIC_FIELD_SP_DOC_ID) = "N/A" ' Or some other placeholder
        End If
        
        If spUrl <> "" And spUrl <> "Error" Then
            fieldData(RAGIC_FIELD_SP_URL) = spUrl
        Else
            fieldData(RAGIC_FIELD_SP_URL) = "N/A"
        End If
    Else
        fieldData(RAGIC_FIELD_SP_DOC_ID) = "N/A_SP_Disabled_Or_No_WB"
        fieldData(RAGIC_FIELD_SP_URL) = "N/A_SP_Disabled_Or_No_WB"
    End If
    
    fieldData(RAGIC_FIELD_USER_DOMAIN) = Environ("USERDOMAIN")
    fieldData(RAGIC_FIELD_COMPUTER_NAME) = Environ("COMPUTERNAME")
    On Error GoTo RagicLogErrorHandler ' Reset error handling

    ' If ErrorContext is provided, add more specific error details
    If Not errorCtx Is Nothing Then
        If errorCtx.ModuleName <> "" And errorCtx.ProcedureName <> "" Then
            fieldData(RAGIC_FIELD_ACTION) = Left(errorCtx.ModuleName & "." & errorCtx.ProcedureName, 255)
        ElseIf errorCtx.ProcedureName <> "" Then
             fieldData(RAGIC_FIELD_ACTION) = Left(errorCtx.ProcedureName, 255)
        End If
        
        Dim errorDetailMsg As String
        errorDetailMsg = "Err #" & errorCtx.ErrorNumber & ": " & errorCtx.ErrorDescription
        If errorCtx.ErrorSource <> "" Then errorDetailMsg = errorDetailMsg & " (Source: " & errorCtx.ErrorSource & ")"
        If details <> "" Then errorDetailMsg = errorDetailMsg & "; Original Details: " & details
        
        fieldData(RAGIC_FIELD_DETAILS) = Left(errorDetailMsg, 32000) ' Ragic free text can be large
        
        ' Populate specific error context fields if they exist in Ragic and are defined in ElyseCore_System
        If RAGIC_FIELD_ERROR_NUMBER <> "" Then fieldData(RAGIC_FIELD_ERROR_NUMBER) = errorCtx.ErrorNumber
        If RAGIC_FIELD_ERROR_SOURCE <> "" Then fieldData(RAGIC_FIELD_ERROR_SOURCE) = Left(errorCtx.ErrorSource, 255)
        If RAGIC_FIELD_ERROR_DESCRIPTION <> "" Then fieldData(RAGIC_FIELD_ERROR_DESCRIPTION) = Left(errorCtx.ErrorDescription, 1000) ' Adjust length as needed
        If RAGIC_FIELD_MODULE_NAME <> "" Then fieldData(RAGIC_FIELD_MODULE_NAME) = Left(errorCtx.ModuleName, 255)
        If RAGIC_FIELD_PROCEDURE_NAME <> "" Then fieldData(RAGIC_FIELD_PROCEDURE_NAME) = Left(errorCtx.ProcedureName, 255)
        If RAGIC_FIELD_LINE_NUMBER <> "" And errorCtx.LineNumber > 0 Then fieldData(RAGIC_FIELD_LINE_NUMBER) = errorCtx.LineNumber
        ' Add other RAGIC_FIELD_... from errorCtx as needed
    End If

    ' Build payload string (URL-encoded form data)
    Dim key As Variant
    payload = ""
    For Each key In fieldData.Keys
        If fieldData(key) <> "" Then ' Only send fields with values
            If payload <> "" Then payload = payload & "&"
            payload = payload & key & "=" & EncodeURL(CStr(fieldData(key)))
        End If
    Next key
    
    Dim ragicPostUrl As String
    ragicPostUrl = RAGIC_LOG_API_URL
    If InStr(1, ragicPostUrl, "?") = 0 Then
        ragicPostUrl = ragicPostUrl & "?api=true"
    Else
        ragicPostUrl = ragicPostUrl & "&api=true"
    End If
    
    http.Open "POST", ragicPostUrl, False
    http.setRequestHeader "Authorization", "Basic " & EncodeBase64("APIKEY:" & RAGIC_LOG_API_KEY) ' Basic Auth with API Key
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send payload

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
    Set http = Nothing
    Set fieldData = Nothing
End Sub

' Helper for URL encoding
Private Function EncodeURL(str As String) As String
    Dim ScriptControl As Object
    Set ScriptControl = CreateObject("ScriptControl")
    ScriptControl.Language = "JScript"
    EncodeURL = ScriptControl.CodeObject.encodeURIComponent(str)
    Set ScriptControl = Nothing
End Function

' Helper for Base64 encoding (common for API keys)
Private Function EncodeBase64(text As String) As String
    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing
End Function

' ============================================================================
' MODIFIED LOGGING FUNCTIONS TO INCLUDE RAGIC
' ============================================================================

' Example modification for LogError (apply similar pattern to others)
Public Sub LogError(actionCode As String, errorCode As Long, message As String, Optional ByVal procedureName As String = "", Optional ByVal moduleName As String = "", Optional errorCtx As ErrorContext = Nothing)
    Dim logMessage As String
    logMessage = "ERROR [" & actionCode & "] "
    If moduleName <> "" And procedureName <> "" Then
        logMessage = logMessage & moduleName & "." & procedureName & " "
    ElseIf procedureName <> "" Then
        logMessage = logMessage & procedureName & " "
    End If
    logMessage = logMessage & "(Code: " & errorCode & "): " & message
    
    If mLoggerActive Then ' Assuming mLoggerActive and mCurrentLogLevel are still part of this module's internal state
        Select Case mCurrentLogLevel
            Case DEBUG_LEVEL, INFO_LEVEL, WARNING_LEVEL, ERROR_LEVEL, CRITICAL_LEVEL ' Assuming these constants are defined
                PrintToImmediate logMessage ' Assuming PrintToImmediate is a helper in this module
                If mLogToFile Then WriteToLogFile "ERROR", logMessage ' Assuming mLogToFile and WriteToLogFile exist
        End Select
    End If
    
    ' Ragic Logging Integration
    Dim errorDetails As String
    errorDetails = message ' Base details
    If Not errorCtx Is Nothing Then
        ' If context is provided, LogToRagic will use it for richer details
         Call LogToRagic("ERROR", actionCode, errorDetails, errorCtx)
    Else
        ' Create a minimal context if none provided, or pass Nothing
        Dim tempCtx As ErrorContext
        tempCtx.ErrorNumber = errorCode
        tempCtx.ErrorDescription = message
        tempCtx.ProcedureName = procedureName
        tempCtx.ModuleName = moduleName
        Call LogToRagic("ERROR", actionCode, errorDetails, tempCtx)
    End If
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
    
    Call LogToRagic("INFO", actionCode, message) 
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
    
    Call LogToRagic("DEBUG", actionCode, message)
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
    
    Call LogToRagic("WARNING", actionCode, message)
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
    
    Call LogToRagic("CRITICAL", actionCode, criticalDetails, tempCtx)
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
    
    Call LogToRagic("DEBUG", action, "Params: " & params)
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
    
    Call LogToRagic("INFO", action, details)
End Sub