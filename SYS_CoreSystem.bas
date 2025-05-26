' ============================================================================
' SYS_CoreSystem - Core System Utilities and Definitions
' Elyse Energy VBA Ecosystem - Core Component
' ============================================================================

Option Explicit

' ============================================================================
' MODULE DEPENDENCIES
' ============================================================================
' This module has no direct dependencies on other custom modules.
' It provides foundational enums, constants, and utility functions.
' ============================================================================

' ============================================================================
' GLOBAL CONSTANTS AND CONFIGURATION
' ============================================================================

' API Configuration
Public Const API_BASE_URL As String = "https://api-logs.elyse-energy.com/v1"
Public Const API_LOGS_ENDPOINT As String = "/logs"
Public Const API_TICKETS_ENDPOINT As String = "/tickets"
Public Const API_TOKEN As String = "Bearer elyse_energy_token_2025"
Public Const SUPPORT_EMAIL As String = "support_data@elyse.energy"

' System Configuration
Public Const TICKET_SUBJECT_PREFIX As String = "[ELYSE TICKET] "
Public Const LOG_BUFFER_SIZE As Integer = 50
Public Const SESSION_TIMEOUT_MINUTES As Integer = 480 ' 8 hours

' Color Palette - Elyse Energy Corporate Identity
Public Const COLOR_PRIMARY As Long = &H2E8B57        ' Sea Green (Energy)
Public Const COLOR_SECONDARY As Long = &H4682B4      ' Steel Blue (Technology)
Public Const COLOR_ACCENT As Long = &H228B22         ' Forest Green (Sustainability)
Public Const COLOR_NEUTRAL As Long = &H708090        ' Slate Gray (Professional)
Public Const COLOR_TEXT As Long = &H2F4F4F           ' Dark Slate Gray (Readability)
Public Const COLOR_BACKGROUND As Long = &HF8F8FF     ' Ghost White (Clean)
Public Const COLOR_SUCCESS As Long = &H32CD32        ' Lime Green (Success)
Public Const COLOR_WARNING As Long = &H87CEEB        ' Sky Blue (Warning)
Public Const COLOR_ERROR As Long = &H4169E1          ' Royal Blue (Error)

' File and Path Configuration
Public Const TEMP_FILE_PREFIX As String = "Elyse_"
Public Const LOG_FILE_EXTENSION As String = ".log"
Public Const CONFIG_FILE_NAME As String = "elyse_config.ini"

' ============================================================================
' ENUMERATIONS
' ============================================================================

' Log Levels
Public Enum LogLevel
    DEBUG_LEVEL = 1
    INFO_LEVEL = 2
    WARNING_LEVEL = 3
    ERROR_LEVEL = 4
    CRITICAL_LEVEL = 5
End Enum

' Message Types
Public Enum MessageType
    INFO_MESSAGE = 1
    SUCCESS_MESSAGE = 2
    WARNING_MESSAGE = 3
    ERROR_MESSAGE = 4
    CONFIRMATION_MESSAGE = 5
End Enum

' File Location Types
Public Enum FileLocationType
    UNKNOWN_LOCATION = 0
    LOCAL_DRIVE = 1
    NETWORK_DRIVE = 2
    SHAREPOINT = 3
    ONEDRIVE = 4
End Enum

' System Modes
Public Enum SystemMode
    DEBUG_MODE = 1
    PRODUCTION_MODE = 2
    MAINTENANCE_MODE = 3
End Enum

' ============================================================================
' GLOBAL VARIABLES AND STATE
' ============================================================================

' Core system state
Private mSystemInitialized As Boolean
Private mCurrentMode As SystemMode
Private mSessionID As String
Private mUsername As String
Private mComputerName As String
Private mUserDomain As String
Private mStartupTime As Date

' Configuration cache
Private mConfigCache As Object
Private mColorScheme As Object

' ============================================================================
' RAGIC API CONFIGURATION FOR LOGGING
' ============================================================================
Public Const RAGIC_LOG_API_URL As String = "https://ragic.elyse.energy/default/matching-matrix/8" ' Add ?api=true or other params as needed
Public Const RAGIC_LOG_API_KEY As String = "YOUR_ACTUAL_RAGIC_API_KEY" ' IMPORTANT: Store securely or manage appropriately

' Ragic Form Field IDs for Logging (as strings for dictionary keys)
Public Const RAGIC_FIELD_TIMESTAMP As String = "1005231"
Public Const RAGIC_FIELD_USER As String = "1005232"
Public Const RAGIC_FIELD_ACTION As String = "1005233" ' e.g., ProcedureName or a specific action being logged
Public Const RAGIC_FIELD_DETAILS As String = "1005234" ' e.g., Error message or log details
Public Const RAGIC_FIELD_LEVEL As String = "1005235" ' e.g., "ERROR", "INFO", "DEBUG"
Public Const RAGIC_FIELD_SESSION_ID As String = "1005236"
Public Const RAGIC_FIELD_EXCEL_VERSION As String = "1005237"
Public Const RAGIC_FIELD_WORKBOOK_NAME As String = "1005238"
Public Const RAGIC_FIELD_SP_DOC_ID As String = "1005239"
Public Const RAGIC_FIELD_SP_URL As String = "1005240"
Public Const RAGIC_FIELD_FILE_LOCATION As String = "1005241"
Public Const RAGIC_FIELD_ACTIVE_SHEET As String = "1005242"
Public Const RAGIC_FIELD_SELECTED_RANGE As String = "1005243"
Public Const RAGIC_FIELD_USER_DOMAIN As String = "1005244"
Public Const RAGIC_FIELD_COMPUTER_NAME As String = "1005245"
' Key field for Ragic sheet, usually not sent for new record creation unless specified by Ragic API for updates.
' Public Const RAGIC_FIELD_KEY_ID As String = "1005246"


' Global flag to enable/disable Ragic logging
Public gEnableRagicLogging As Boolean
Private gSessionID As String

' ============================================================================
' SYSTEM INITIALIZATION AND CONFIGURATION
' ============================================================================

Public Function InitializeCoreSystem(Optional mode As SystemMode = PRODUCTION_MODE) As Boolean
    ' Initialize the core Elyse Energy system
    On Error GoTo ErrorHandler
    
    If mSystemInitialized Then
        InitializeCoreSystem = True
        Exit Function
    End If
    
    ' Set system mode
    mCurrentMode = mode
    mStartupTime = Now
    
    ' Initialize core components
    mSessionID = GenerateSessionUUID()
    mUsername = GetUserIdentity()
    mComputerName = Environ("COMPUTERNAME")
    mUserDomain = Environ("USERDOMAIN")
    
    ' Initialize configuration cache
    Set mConfigCache = CreateObject("Scripting.Dictionary")
    Set mColorScheme = CreateObject("Scripting.Dictionary")
    
    ' Load configuration
    LoadSystemConfiguration
    
    ' Initialize color scheme
    InitializeColorScheme
    
    mSystemInitialized = True
    InitializeCoreSystem = True
    Exit Function
    
ErrorHandler:
    InitializeCoreSystem = False
End Function

Public Sub ShutdownCoreSystem()
    ' Clean shutdown of core system
    
    mSystemInitialized = False
    mSessionID = ""
    mUsername = ""
    
    ' Clear caches
    Set mConfigCache = Nothing
    Set mColorScheme = Nothing
End Sub

' ============================================================================
' SESSION AND IDENTITY MANAGEMENT
' ============================================================================

Public Function GenerateSessionUUID() As String
    ' Generate a unique session identifier
    Randomize
    Dim timestamp As String
    Dim randomPart As String
    
    timestamp = Format(Now, "yyyymmdd_hhnnss")
    randomPart = Right("000" & Int(Rnd * 1000), 3)
    
    GenerateSessionUUID = timestamp & "_" & randomPart
End Function

Public Function GetSessionID() As String
    ' Get current session ID
    GetSessionID = mSessionID
End Function

Public Function GetUserIdentity() As String
    ' Get user identity with multiple fallback options
    Dim username As String
    
    ' Try Windows username first
    username = Environ("USERNAME")
    
    ' Fallback to Excel username
    If username = "" Then
        username = Application.UserName
    End If
    
    ' Final fallback
    If username = "" Then
        username = "unknown_user"
    End If
    
    GetUserIdentity = username
End Function

Public Function GetCurrentUserIdentifier() As String
    On Error Resume Next
    GetCurrentUserIdentifier = Application.UserName
    If Err.Number <> 0 Then GetCurrentUserIdentifier = "UnknownUser"
    On Error GoTo 0
End Function

' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================

Public Function GetCurrentWorkbookName() As String
    ' Get current workbook name safely
    On Error Resume Next
    GetCurrentWorkbookName = ActiveWorkbook.Name
    If GetCurrentWorkbookName = "" Then GetCurrentWorkbookName = "unknown_workbook"
    On Error GoTo 0
End Function

Public Function GetActiveSheetName() As String
    ' Get active sheet name safely
    On Error Resume Next
    GetActiveSheetName = ActiveSheet.Name
    If GetActiveSheetName = "" Then GetActiveSheetName = "unknown_sheet"
    On Error GoTo 0
End Function

Public Function GetSelectedRangeAddress() As String
    ' Get selected range address safely
    On Error Resume Next
    GetSelectedRangeAddress = Selection.Address
    If GetSelectedRangeAddress = "" Then GetSelectedRangeAddress = "unknown_range"
    On Error GoTo 0
End Function

Public Function GetTempFilePath(fileName As String) As String
    ' Generate temp file path with Elyse prefix
    GetTempFilePath = Environ("TEMP") & "\" & TEMP_FILE_PREFIX & fileName
End Function

Public Function GetStringHash(text As String) As Long
    ' Generate simple hash for string
    Dim i As Integer
    Dim hashValue As Long
    
    hashValue = 0
    For i = 1 To Len(text)
        hashValue = hashValue + Asc(Mid(text, i, 1)) * i
    Next i
    
    GetStringHash = hashValue
End Function

Public Function EscapeJSON(text As String) As String
    ' Escape special characters for JSON
    EscapeJSON = Replace(text, """", "\""")
    EscapeJSON = Replace(EscapeJSON, "\", "\\")
    EscapeJSON = Replace(EscapeJSON, vbCrLf, "\n")
    EscapeJSON = Replace(EscapeJSON, vbTab, "\t")
    EscapeJSON = Replace(EscapeJSON, vbCr, "\r")
End Function

Public Function StripHTML(htmlText As String) As String
    ' Remove HTML tags from text
    Dim result As String
    Dim i As Integer
    Dim inTag As Boolean
    
    result = ""
    inTag = False
    
    For i = 1 To Len(htmlText)
        Dim char As String
        char = Mid(htmlText, i, 1)
        
        If char = "<" Then
            inTag = True
        ElseIf char = ">" Then
            inTag = False
        ElseIf Not inTag Then
            result = result & char
        End If
    Next i
    
    StripHTML = result
End Function

' ============================================================================
' ENUM TO STRING CONVERSIONS
' ============================================================================

Public Function GetLogLevelString(level As LogLevel) As String
    ' Convert log level enum to string
    Select Case level
        Case DEBUG_LEVEL: GetLogLevelString = "DEBUG"
        Case INFO_LEVEL: GetLogLevelString = "INFO"
        Case WARNING_LEVEL: GetLogLevelString = "WARNING"
        Case ERROR_LEVEL: GetLogLevelString = "ERROR"
        Case CRITICAL_LEVEL: GetLogLevelString = "CRITICAL"
        Case Else: GetLogLevelString = "INFO"
    End Select
End Function

Public Function GetMessageTypeString(msgType As MessageType) As String
    ' Convert message type enum to string
    Select Case msgType
        Case INFO_MESSAGE: GetMessageTypeString = "INFO"
        Case SUCCESS_MESSAGE: GetMessageTypeString = "SUCCESS"
        Case WARNING_MESSAGE: GetMessageTypeString = "WARNING"
        Case ERROR_MESSAGE: GetMessageTypeString = "ERROR"
        Case CONFIRMATION_MESSAGE: GetMessageTypeString = "CONFIRMATION"
        Case Else: GetMessageTypeString = "INFO"
    End Select
End Function

Public Function GetSystemModeString(mode As SystemMode) As String
    ' Convert system mode enum to string
    Select Case mode
        Case DEBUG_MODE: GetSystemModeString = "DEBUG"
        Case PRODUCTION_MODE: GetSystemModeString = "PRODUCTION"
        Case MAINTENANCE_MODE: GetSystemModeString = "MAINTENANCE"
        Case Else: GetSystemModeString = "PRODUCTION"
    End Select
End Function

Public Function GetFileLocationTypeString(locationType As FileLocationType) As String
    ' Convert file location type enum to string
    Select Case locationType
        Case LOCAL_DRIVE: GetFileLocationTypeString = "local_drive"
        Case NETWORK_DRIVE: GetFileLocationTypeString = "network_drive"
        Case SHAREPOINT: GetFileLocationTypeString = "sharepoint"
        Case ONEDRIVE: GetFileLocationTypeString = "onedrive"
        Case Else: GetFileLocationTypeString = "unknown"
    End Select
End Function

' ============================================================================
' COLOR SCHEME MANAGEMENT
' ============================================================================

Private Sub InitializeColorScheme()
    ' Initialize the Elyse Energy color scheme
    
    mColorScheme("primary") = COLOR_PRIMARY
    mColorScheme("secondary") = COLOR_SECONDARY
    mColorScheme("accent") = COLOR_ACCENT
    mColorScheme("neutral") = COLOR_NEUTRAL
    mColorScheme("text") = COLOR_TEXT
    mColorScheme("background") = COLOR_BACKGROUND
    mColorScheme("success") = COLOR_SUCCESS
    mColorScheme("warning") = COLOR_WARNING
    mColorScheme("error") = COLOR_ERROR
End Sub

Public Function GetColorScheme() As Object
    ' Get the complete color scheme
    Set GetColorScheme = mColorScheme
End Function

Public Function GetColor(colorName As String) As Long
    ' Get specific color from scheme
    If mColorScheme.Exists(colorName) Then
        GetColor = mColorScheme(colorName)
    Else
        GetColor = COLOR_NEUTRAL ' Default fallback
    End If
End Function

Public Function GetColorForMessageType(msgType As MessageType) As Long
    ' Get appropriate color for message type
    Select Case msgType
        Case INFO_MESSAGE: GetColorForMessageType = COLOR_PRIMARY
        Case SUCCESS_MESSAGE: GetColorForMessageType = COLOR_SUCCESS
        Case WARNING_MESSAGE: GetColorForMessageType = COLOR_WARNING
        Case ERROR_MESSAGE: GetColorForMessageType = COLOR_ERROR
        Case CONFIRMATION_MESSAGE: GetColorForMessageType = COLOR_SECONDARY
        Case Else: GetColorForMessageType = COLOR_NEUTRAL
    End Select
End Function

' ============================================================================
' CONFIGURATION MANAGEMENT
' ============================================================================

Private Sub LoadSystemConfiguration()
    ' Load system configuration from file or defaults
    
    ' Set default configuration values
    mConfigCache("log_level") = IIf(mCurrentMode = DEBUG_MODE, DEBUG_LEVEL, INFO_LEVEL)
    mConfigCache("auto_flush_logs") = True
    mConfigCache("enable_error_tickets") = True
    mConfigCache("enable_sharepoint") = True
    mConfigCache("api_timeout") = 5000 ' milliseconds
    mConfigCache("max_retry_attempts") = 3
    
    ' Try to load from file (optional enhancement)
    ' LoadConfigurationFromFile()
End Sub

Public Function GetConfigValue(key As String, Optional defaultValue As Variant = "") As Variant
    ' Get configuration value with fallback
    If mConfigCache.Exists(key) Then
        GetConfigValue = mConfigCache(key)
    Else
        GetConfigValue = defaultValue
    End If
End Function

Public Sub SetConfigValue(key As String, value As Variant)
    ' Set configuration value
    mConfigCache(key) = value
End Sub

' ============================================================================
' VALIDATION AND ERROR CHECKING
' ============================================================================

Public Function IsSystemInitialized() As Boolean
    ' Check if core system is initialized
    IsSystemInitialized = mSystemInitialized
End Function

Public Function ValidateEmailAddress(email As String) As Boolean
    ' Basic email validation
    ValidateEmailAddress = (InStr(email, "@") > 0 And InStr(email, ".") > 0)
End Function

Public Function ValidateURL(url As String) As Boolean
    ' Basic URL validation
    ValidateURL = (Left(LCase(url), 4) = "http")
End Function

Public Function IsDebugMode() As Boolean
    ' Check if system is in debug mode
    IsDebugMode = (mCurrentMode = DEBUG_MODE)
End Function

Public Function IsProductionMode() As Boolean
    ' Check if system is in production mode
    IsProductionMode = (mCurrentMode = PRODUCTION_MODE)
End Function

' ============================================================================
' FORMATTING AND DISPLAY HELPERS
' ============================================================================

Public Function FormatTimestamp(Optional includeMilliseconds As Boolean = False) As String
    ' Format current timestamp
    If includeMilliseconds Then
        FormatTimestamp = Format(Now, "yyyy-mm-dd hh:nn:ss") & "." & Right("000" & Timer Mod 1 * 1000, 3)
    Else
        FormatTimestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
    End If
End Function

Public Function FormatFileSize(sizeInBytes As Long) As String
    ' Format file size in human readable format
    If sizeInBytes < 1024 Then
        FormatFileSize = sizeInBytes & " B"
    ElseIf sizeInBytes < 1048576 Then
        FormatFileSize = Format(sizeInBytes / 1024, "0.0") & " KB"
    Else
        FormatFileSize = Format(sizeInBytes / 1048576, "0.0") & " MB"
    End If
End Function

Public Function TruncateString(text As String, maxLength As Integer, Optional suffix As String = "...") As String
    ' Truncate string with optional suffix
    If Len(text) <= maxLength Then
        TruncateString = text
    Else
        TruncateString = Left(text, maxLength - Len(suffix)) & suffix
    End If
End Function

' ============================================================================
' SYSTEM STATUS AND HEALTH
' ============================================================================

Public Sub InitializeCoreSystemExtras()
    ' Called by ElyseMain_Orchestrator.InitializeElyseSystem
    ' Set gEnableRagicLogging based on configuration (e.g., from a settings sheet, environment variable, etc.)
    ' For now, let's default it to True for demonstration.
    gEnableRagicLogging = True 
    Call GetSessionID ' Initialize session ID
End Sub