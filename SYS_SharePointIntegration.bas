' ============================================================================
' ElyseSharePoint_Integration - SharePoint Metadata Management
' Elyse Energy VBA Ecosystem - SharePoint Component
' Requires: ElyseCore_System, ElyseLogger_Module
' ============================================================================

Option Explicit

' ============================================================================
' MODULE DEPENDENCIES
' ============================================================================
' This module requires:
' - ElyseCore_System (enums, constants, utilities)
' - ElyseLogger_Module (logging functions)

' ============================================================================
' SHAREPOINT CONFIGURATION AND TYPES
' ============================================================================

Public Type SharePointDocumentInfo
    DocumentID As String
    ItemGUID As String
    DocumentURL As String
    SiteURL As String
    LibraryName As String
    FolderPath As String
    FileName As String
    FileExtension As String
    LastModified As Date
    FileSize As Long
    CheckoutUser As String
    VersionNumber As String
    ContentType As String
    IsCheckedOut As Boolean
End Type

Public Type SharePointSiteInfo
    SiteURL As String
    SiteName As String
    SiteCollection As String
    WebApplication As String
    ServerName As String
    IsOnline As Boolean
    TenantName As String
End Type

' ============================================================================
' SHAREPOINT STATE VARIABLES
' ============================================================================

Private mCurrentDocumentInfo As SharePointDocumentInfo
Private mCurrentSiteInfo As SharePointSiteInfo
Private mSharePointAvailable As Boolean
Private mLastMetadataCheck As Date
Private mCachedMetadata As Object

' ============================================================================
' SHAREPOINT DETECTION AND INITIALIZATION
' ============================================================================

Public Function InitializeSharePointIntegration() As Boolean
    ' Initialize SharePoint integration and detect if file is on SharePoint
    
    LogInfo "sharepoint_init", "Initializing SharePoint integration"
    
    ' Initialize cache
    Set mCachedMetadata = CreateObject("Scripting.Dictionary")
    
    ' Detect SharePoint availability
    mSharePointAvailable = DetectSharePointEnvironment()
    
    If mSharePointAvailable Then
        ' Get document information
        RefreshDocumentMetadata
        
        LogInfo "sharepoint_detected", "SharePoint environment detected: " & mCurrentSiteInfo.SiteURL
    Else
        LogInfo "sharepoint_not_detected", "File is not on SharePoint or SharePoint not available"
    End If
    
    InitializeSharePointIntegration = mSharePointAvailable
End Function

Private Function DetectSharePointEnvironment() As Boolean
    ' Detect if current workbook is on SharePoint
    On Error GoTo ErrorHandler
    
    Dim workbookPath As String
    workbookPath = ActiveWorkbook.FullName
    
    ' Check if path indicates SharePoint
    If InStr(LCase(workbookPath), "sharepoint") > 0 Or _
       InStr(LCase(workbookPath), ".sharepoint.com") > 0 Or _
       (InStr(LCase(workbookPath), "https://") = 1 And InStr(workbookPath, "/sites/") > 0) Then
        
        ' Extract site information
        ExtractSiteInformation workbookPath
        DetectSharePointEnvironment = True
    Else
        DetectSharePointEnvironment = False
    End If
    
    Exit Function
    
ErrorHandler:
    LogError "sharepoint_detection_error", Err.Number, Err.Description
    DetectSharePointEnvironment = False
End Function

Private Sub ExtractSiteInformation(documentPath As String)
    ' Extract SharePoint site information from document path
    
    mCurrentSiteInfo.SiteURL = documentPath
    
    ' Parse URL components
    If InStr(documentPath, "/sites/") > 0 Then
        Dim parts() As String
        parts = Split(documentPath, "/")
        
        If UBound(parts) >= 4 Then
            mCurrentSiteInfo.ServerName = parts(2)
            mCurrentSiteInfo.SiteName = parts(4)
            mCurrentSiteInfo.IsOnline = (InStr(LCase(parts(2)), ".sharepoint.com") > 0)
            
            If mCurrentSiteInfo.IsOnline Then
                mCurrentSiteInfo.TenantName = Split(parts(2), ".")(0)
            End If
        End If
    End If
End Sub

' ============================================================================
' DOCUMENT METADATA RETRIEVAL
' ============================================================================

Public Function RefreshDocumentMetadata() As Boolean
    ' Refresh document metadata from SharePoint
    On Error GoTo ErrorHandler
    
    If Not mSharePointAvailable Then
        RefreshDocumentMetadata = False
        Exit Function
    End If
    
    LogDebug "sharepoint_metadata_refresh", "Refreshing document metadata"
    
    ' Get document properties
    GetDocumentIDProperties
    GetDocumentURLProperties
    GetDocumentFileProperties
    GetDocumentVersionProperties
    
    ' Cache the metadata
    CacheDocumentMetadata
    
    mLastMetadataCheck = Now
    RefreshDocumentMetadata = True
    
    LogInfo "sharepoint_metadata_refreshed", "Document metadata updated successfully"
    Exit Function
    
ErrorHandler:
    LogError "sharepoint_metadata_error", Err.Number, Err.Description
    RefreshDocumentMetadata = False
End Function

Private Sub GetDocumentIDProperties()
    ' Get SharePoint document ID properties
    On Error Resume Next
    
    ' Try to get SharePoint Document ID
    mCurrentDocumentInfo.DocumentID = ActiveWorkbook.BuiltinDocumentProperties("_dlc_DocId").Value
    
    If mCurrentDocumentInfo.DocumentID = "" Then
        mCurrentDocumentInfo.DocumentID = ActiveWorkbook.CustomDocumentProperties("_dlc_DocId").Value
    End If
    
    ' Try to get Item GUID
    mCurrentDocumentInfo.ItemGUID = ActiveWorkbook.BuiltinDocumentProperties("_dlc_DocIdItemGuid").Value
    
    If mCurrentDocumentInfo.ItemGUID = "" Then
        mCurrentDocumentInfo.ItemGUID = ActiveWorkbook.CustomDocumentProperties("_dlc_DocIdItemGuid").Value
    End If
    
    ' Generate fallback ID if no SharePoint ID found
    If mCurrentDocumentInfo.DocumentID = "" Then
        mCurrentDocumentInfo.DocumentID = "local_" & GenerateFileHashID()
    End If
    
    On Error GoTo 0
End Sub

Private Sub GetDocumentURLProperties()
    ' Get SharePoint URL properties
    On Error Resume Next
    
    ' Get document URL
    mCurrentDocumentInfo.DocumentURL = ActiveWorkbook.FullName
    
    ' Try to get SharePoint-specific URL properties
    Dim sharePointURL As String
    sharePointURL = ActiveWorkbook.BuiltinDocumentProperties("_dlc_DocIdUrl").Value
    
    If sharePointURL <> "" Then
        mCurrentDocumentInfo.DocumentURL = sharePointURL
    End If
    
    ' Extract components from URL
    ExtractURLComponents mCurrentDocumentInfo.DocumentURL
    
    On Error GoTo 0
End Sub

Private Sub ExtractURLComponents(documentURL As String)
    ' Extract components from SharePoint URL
    
    Dim parts() As String
    parts = Split(documentURL, "/")
    
    If UBound(parts) >= 0 Then
        ' Get filename
        mCurrentDocumentInfo.FileName = parts(UBound(parts))
        
        ' Get file extension
        If InStr(mCurrentDocumentInfo.FileName, ".") > 0 Then
            Dim nameParts() As String
            nameParts = Split(mCurrentDocumentInfo.FileName, ".")
            mCurrentDocumentInfo.FileExtension = nameParts(UBound(nameParts))
        End If
        
        ' Extract library and folder path
        If UBound(parts) >= 6 Then
            Dim i As Integer
            Dim folderPath As String
            
            ' Find document library (typically after /sites/sitename/)
            For i = 5 To UBound(parts) - 1
                If i = 5 Then
                    mCurrentDocumentInfo.LibraryName = parts(i)
                Else
                    folderPath = folderPath & "/" & parts(i)
                End If
            Next i
            
            mCurrentDocumentInfo.FolderPath = folderPath
        End If
    End If
End Sub

Private Sub GetDocumentFileProperties()
    ' Get file-related properties
    On Error Resume Next
    
    mCurrentDocumentInfo.LastModified = ActiveWorkbook.BuiltinDocumentProperties("Last Save Time").Value
    
    ' Try to get file size
    Dim filePath As String
    filePath = ActiveWorkbook.FullName
    
    If InStr(filePath, "http") = 0 Then ' Local file
        mCurrentDocumentInfo.FileSize = FileLen(filePath)
    End If
    
    On Error GoTo 0
End Sub

Private Sub GetDocumentVersionProperties()
    ' Get version-related properties
    On Error Resume Next
    
    ' Try to get version information
    mCurrentDocumentInfo.VersionNumber = ActiveWorkbook.BuiltinDocumentProperties("Revision Number").Value
    
    ' Try to get content type
    mCurrentDocumentInfo.ContentType = ActiveWorkbook.CustomDocumentProperties("Content Type").Value
    
    If mCurrentDocumentInfo.ContentType = "" Then
        mCurrentDocumentInfo.ContentType = "Document"
    End If
    
    ' Check if document is checked out
    mCurrentDocumentInfo.IsCheckedOut = False ' Default
    mCurrentDocumentInfo.CheckoutUser = ""
    
    On Error GoTo 0
End Sub

Private Sub CacheDocumentMetadata()
    ' Cache document metadata for quick access
    
    mCachedMetadata("document_id") = mCurrentDocumentInfo.DocumentID
    mCachedMetadata("item_guid") = mCurrentDocumentInfo.ItemGUID
    mCachedMetadata("document_url") = mCurrentDocumentInfo.DocumentURL
    mCachedMetadata("site_url") = mCurrentSiteInfo.SiteURL
    mCachedMetadata("library_name") = mCurrentDocumentInfo.LibraryName
    mCachedMetadata("folder_path") = mCurrentDocumentInfo.FolderPath
    mCachedMetadata("file_name") = mCurrentDocumentInfo.FileName
    mCachedMetadata("file_extension") = mCurrentDocumentInfo.FileExtension
    mCachedMetadata("last_modified") = mCurrentDocumentInfo.LastModified
    mCachedMetadata("file_size") = mCurrentDocumentInfo.FileSize
    mCachedMetadata("version_number") = mCurrentDocumentInfo.VersionNumber
    mCachedMetadata("content_type") = mCurrentDocumentInfo.ContentType
    mCachedMetadata("is_sharepoint") = mSharePointAvailable
    mCachedMetadata("site_name") = mCurrentSiteInfo.SiteName
    mCachedMetadata("server_name") = mCurrentSiteInfo.ServerName
    mCachedMetadata("is_online") = mCurrentSiteInfo.IsOnline
    mCachedMetadata("tenant_name") = mCurrentSiteInfo.TenantName
End Sub

' ============================================================================
' PUBLIC METADATA ACCESS FUNCTIONS
' ============================================================================

Public Function GetSharePointDocumentID() As String
    ' Get SharePoint document ID with caching
    
    If Not IsMetadataCacheValid() Then
        RefreshDocumentMetadata
    End If
    
    GetSharePointDocumentID = mCurrentDocumentInfo.DocumentID
End Function

Public Function GetSharePointDocumentURL() As String
    ' Get SharePoint document URL with caching
    
    If Not IsMetadataCacheValid() Then
        RefreshDocumentMetadata
    End If
    
    GetSharePointDocumentURL = mCurrentDocumentInfo.DocumentURL
End Function

Public Function GetSharePointSiteURL() As String
    ' Get SharePoint site URL
    
    If Not IsMetadataCacheValid() Then
        RefreshDocumentMetadata
    End If
    
    GetSharePointSiteURL = mCurrentSiteInfo.SiteURL
End Function

Public Function GetFileLocationInfo() As Object
    ' Get comprehensive file location information
    
    Dim locationInfo As Object
    Set locationInfo = CreateObject("Scripting.Dictionary")
    
    If Not IsMetadataCacheValid() Then
        RefreshDocumentMetadata
    End If
    
    locationInfo("is_sharepoint") = mSharePointAvailable
    locationInfo("location_type") = GetFileLocationTypeString(GetFileLocationType())
    locationInfo("document_id") = mCurrentDocumentInfo.DocumentID
    locationInfo("document_url") = mCurrentDocumentInfo.DocumentURL
    locationInfo("site_url") = mCurrentSiteInfo.SiteURL
    locationInfo("library_name") = mCurrentDocumentInfo.LibraryName
    locationInfo("folder_path") = mCurrentDocumentInfo.FolderPath
    locationInfo("file_name") = mCurrentDocumentInfo.FileName
    
    Set GetFileLocationInfo = locationInfo
End Function

Public Function GetDocumentMetadata() As Object
    ' Get complete document metadata
    
    If Not IsMetadataCacheValid() Then
        RefreshDocumentMetadata
    End If
    
    Dim metadata As Object
    Set metadata = CreateObject("Scripting.Dictionary")
    
    ' Copy all cached metadata
    Dim keys As Variant
    keys = mCachedMetadata.Keys
    
    Dim i As Integer
    For i = 0 To UBound(keys)
        metadata(keys(i)) = mCachedMetadata(keys(i))
    Next i
    
    Set GetDocumentMetadata = metadata
End Function

Private Function IsMetadataCacheValid() As Boolean
    ' Check if metadata cache is still valid (5 minute timeout)
    
    If mCachedMetadata Is Nothing Then
        IsMetadataCacheValid = False
    ElseIf DateDiff("n", mLastMetadataCheck, Now) > 5 Then
        IsMetadataCacheValid = False
    Else
        IsMetadataCacheValid = True
    End If
End Function

' ============================================================================
' FILE LOCATION DETECTION
' ============================================================================

Public Function GetFileLocationType() As FileLocationType
    ' Determine the type of file location
    
    Dim filePath As String
    filePath = ActiveWorkbook.FullName
    
    If InStr(LCase(filePath), "sharepoint") > 0 Or _
       InStr(LCase(filePath), ".sharepoint.com") > 0 Or _
       (InStr(LCase(filePath), "https://") = 1 And InStr(filePath, "/sites/") > 0) Then
        GetFileLocationType = SHAREPOINT
        
    ElseIf InStr(LCase(filePath), "onedrive") > 0 Or _
           InStr(LCase(filePath), "-my.sharepoint.com") > 0 Then
        GetFileLocationType = ONEDRIVE
        
    ElseIf Left(filePath, 2) = "\\" Then
        GetFileLocationType = NETWORK_DRIVE
        
    ElseIf Mid(filePath, 2, 1) = ":" Then
        GetFileLocationType = LOCAL_DRIVE
        
    Else
        GetFileLocationType = UNKNOWN_LOCATION
    End If
End Function

Public Function IsSharePointDocument() As Boolean
    ' Check if current document is on SharePoint
    IsSharePointDocument = (GetFileLocationType() = SHAREPOINT)
End Function

Public Function IsOneDriveDocument() As Boolean
    ' Check if current document is on OneDrive
    IsOneDriveDocument = (GetFileLocationType() = ONEDRIVE)
End Function

Public Function IsCloudDocument() As Boolean
    ' Check if current document is in the cloud (SharePoint or OneDrive)
    Dim locationType As FileLocationType
    locationType = GetFileLocationType()
    IsCloudDocument = (locationType = SHAREPOINT Or locationType = ONEDRIVE)
End Function

' ============================================================================
' SHAREPOINT OPERATIONS
' ============================================================================

Public Function CheckOutDocument() As Boolean
    ' Check out the current document (if on SharePoint)
    On Error GoTo ErrorHandler
    
    If Not mSharePointAvailable Then
        LogWarning "sharepoint_checkout_failed", "Document is not on SharePoint"
        CheckOutDocument = False
        Exit Function
    End If
    
    LogInfo "sharepoint_checkout_attempt", "Attempting to check out document"
    
    ' Try to check out the document
    ' Note: This requires SharePoint integration APIs which may not be available in all Excel versions
    ActiveWorkbook.CheckOut
    
    ' Update checkout status
    mCurrentDocumentInfo.IsCheckedOut = True
    mCurrentDocumentInfo.CheckoutUser = GetUserIdentity()
    
    LogInfo "sharepoint_checkout_success", "Document checked out successfully"
    CheckOutDocument = True
    Exit Function
    
ErrorHandler:
    LogError "sharepoint_checkout_error", Err.Number, Err.Description
    CheckOutDocument = False
End Function

Public Function CheckInDocument(Optional comment As String = "") As Boolean
    ' Check in the current document (if on SharePoint)
    On Error GoTo ErrorHandler
    
    If Not mSharePointAvailable Then
        LogWarning "sharepoint_checkin_failed", "Document is not on SharePoint"
        CheckInDocument = False
        Exit Function
    End If
    
    LogInfo "sharepoint_checkin_attempt", "Attempting to check in document"
    
    ' Try to check in the document
    If comment = "" Then
        comment = "Checked in via Elyse Energy Excel Add-in"
    End If
    
    ActiveWorkbook.CheckIn True, comment, True ' Save changes, add comment, make major version
    
    ' Update checkout status
    mCurrentDocumentInfo.IsCheckedOut = False
    mCurrentDocumentInfo.CheckoutUser = ""
    
    LogInfo "sharepoint_checkin_success", "Document checked in successfully"
    CheckInDocument = True
    Exit Function
    
ErrorHandler:
    LogError "sharepoint_checkin_error", Err.Number, Err.Description
    CheckInDocument = False
End Function

Public Function DiscardCheckOut() As Boolean
    ' Discard check out of the current document
    On Error GoTo ErrorHandler
    
    If Not mSharePointAvailable Then
        LogWarning "sharepoint_discard_failed", "Document is not on SharePoint"
        DiscardCheckOut = False
        Exit Function
    End If
    
    LogInfo "sharepoint_discard_attempt", "Attempting to discard check out"
    
    ' Discard checkout
    ActiveWorkbook.UndoCheckOut
    
    ' Update checkout status
    mCurrentDocumentInfo.IsCheckedOut = False
    mCurrentDocumentInfo.CheckoutUser = ""
    
    LogInfo "sharepoint_discard_success", "Check out discarded successfully"
    DiscardCheckOut = True
    Exit Function
    
ErrorHandler:
    LogError "sharepoint_discard_error", Err.Number, Err.Description
    DiscardCheckOut = False
End Function

' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================

Private Function GenerateFileHashID() As String
    ' Generate a hash-based ID for local files
    On Error Resume Next
    
    Dim filePath As String
    Dim fileSize As Long
    
    filePath = ActiveWorkbook.FullName
    
    ' Get file size if possible
    If InStr(filePath, "http") = 0 Then
        fileSize = FileLen(filePath)
    Else
        fileSize = 0
    End If
    
    ' Create hash from path and size
    Dim hashValue As Long
    hashValue = GetStringHash(filePath) + fileSize
    
    GenerateFileHashID = Format(Abs(hashValue), "000000000")
    
    On Error GoTo 0
End Function

Public Function FormatSharePointURL(rawURL As String) As String
    ' Format SharePoint URL for display
    
    Dim formattedURL As String
    formattedURL = rawURL
    
    ' Remove query parameters
    If InStr(formattedURL, "?") > 0 Then
        formattedURL = Left(formattedURL, InStr(formattedURL, "?") - 1)
    End If
    
    ' Decode URL if needed
    formattedURL = Replace(formattedURL, "%20", " ")
    
    FormatSharePointURL = formattedURL
End Function

Public Function ExtractSiteNameFromURL(siteURL As String) As String
    ' Extract site name from SharePoint URL
    
    If InStr(siteURL, "/sites/") > 0 Then
        Dim parts() As String
        parts = Split(siteURL, "/")
        
        Dim i As Integer
        For i = 0 To UBound(parts)
            If LCase(parts(i)) = "sites" And i < UBound(parts) Then
                ExtractSiteNameFromURL = parts(i + 1)
                Exit Function
            End If
        Next i
    End If
    
    ExtractSiteNameFromURL = "Unknown"
End Function

' ============================================================================
' INTEGRATION WITH LOGGING SYSTEM
' ============================================================================

Public Sub UpdateLoggerWithSharePointContext()
    ' Update logger with SharePoint context information
    
    If Not mSharePointAvailable Then Exit Sub
    
    ' This function would be called by the logger to enrich log entries
    ' with SharePoint metadata
    
    LogDebug "sharepoint_context_update", "Updating logger with SharePoint context"
    
    ' The logger module would call this function to get enriched context
    ' This creates a two-way integration between modules
End Sub

Public Function GetSharePointContextForLogging() As Object
    ' Get SharePoint context specifically for logging purposes
    
    Dim context As Object
    Set context = CreateObject("Scripting.Dictionary")
    
    If mSharePointAvailable Then
        context("sharepoint_doc_id") = mCurrentDocumentInfo.DocumentID
        context("sharepoint_url") = FormatSharePointURL(mCurrentDocumentInfo.DocumentURL)
        context("sharepoint_site") = mCurrentSiteInfo.SiteName
        context("sharepoint_library") = mCurrentDocumentInfo.LibraryName
        context("file_location") = "sharepoint"
    Else
        context("sharepoint_doc_id") = GenerateFileHashID()
        context("sharepoint_url") = "local_file"
        context("sharepoint_site") = "local"
        context("sharepoint_library") = "local"
        context("file_location") = GetFileLocationTypeString(GetFileLocationType())
    End If
    
    Set GetSharePointContextForLogging = context
End Function

' ============================================================================
' SHAREPOINT HEALTH AND DIAGNOSTICS
' ============================================================================

Public Function TestSharePointConnectivity() As Boolean
    ' Test connectivity to SharePoint
    On Error GoTo ErrorHandler
    
    If Not mSharePointAvailable Then
        TestSharePointConnectivity = False
        Exit Function
    End If
    
    LogInfo "sharepoint_connectivity_test", "Testing SharePoint connectivity"
    
    ' Try to access document properties (this tests connectivity)
    Dim testProperty As String
    testProperty = ActiveWorkbook.BuiltinDocumentProperties("Title").Value
    
    ' If we get here without error, connectivity is OK
    LogInfo "sharepoint_connectivity_success", "SharePoint connectivity test passed"
    TestSharePointConnectivity = True
    Exit Function
    
ErrorHandler:
    LogError "sharepoint_connectivity_failed", Err.Number, Err.Description
    TestSharePointConnectivity = False
End Function

Public Function GetSharePointDiagnostics() As Object
    ' Get comprehensive SharePoint diagnostics
    
    Dim diagnostics As Object
    Set diagnostics = CreateObject("Scripting.Dictionary")
    
    diagnostics("sharepoint_available") = mSharePointAvailable
    diagnostics("last_metadata_check") = Format(mLastMetadataCheck, "yyyy-mm-dd hh:nn:ss")
    diagnostics("metadata_cache_valid") = IsMetadataCacheValid()
    diagnostics("connectivity_ok") = TestSharePointConnectivity()
    diagnostics("document_id") = mCurrentDocumentInfo.DocumentID
    diagnostics("is_checked_out") = mCurrentDocumentInfo.IsCheckedOut
    diagnostics("checkout_user") = mCurrentDocumentInfo.CheckoutUser
    diagnostics("site_url") = mCurrentSiteInfo.SiteURL
    diagnostics("site_name") = mCurrentSiteInfo.SiteName
    diagnostics("is_online") = mCurrentSiteInfo.IsOnline
    diagnostics("tenant_name") = mCurrentSiteInfo.TenantName
    
    Set GetSharePointDiagnostics = diagnostics
End Function

' ============================================================================
' PUBLIC CONVENIENCE FUNCTIONS
' ============================================================================

Public Function GetDocumentIdentifier() As String
    ' Get the best available document identifier
    
    If mSharePointAvailable And mCurrentDocumentInfo.DocumentID <> "" Then
        GetDocumentIdentifier = mCurrentDocumentInfo.DocumentID
    Else
        GetDocumentIdentifier = GenerateFileHashID()
    End If
End Function

Public Function GetDisplayablePath() As String
    ' Get a user-friendly display path
    
    If mSharePointAvailable Then
        Dim displayPath As String
        displayPath = mCurrentSiteInfo.SiteName
        
        If mCurrentDocumentInfo.LibraryName <> "" Then
            displayPath = displayPath & " > " & mCurrentDocumentInfo.LibraryName
        End If
        
        If mCurrentDocumentInfo.FolderPath <> "" Then
            displayPath = displayPath & " > " & Replace(mCurrentDocumentInfo.FolderPath, "/", " > ")
        End If
        
        displayPath = displayPath & " > " & mCurrentDocumentInfo.FileName
        
        GetDisplayablePath = displayPath
    Else
        GetDisplayablePath = ActiveWorkbook.FullName
    End If
End Function

' ============================================================================
' MODULE STATUS AND CLEANUP
' ============================================================================

Public Function GetSharePointModuleStatus() As Object
    ' Get status of SharePoint integration module
    
    Dim status As Object
    Set status = CreateObject("Scripting.Dictionary")
    
    status("module_loaded") = True
    status("sharepoint_available") = mSharePointAvailable
    status("last_metadata_check") = Format(mLastMetadataCheck, "yyyy-mm-dd hh:nn:ss")
    status("cache_valid") = IsMetadataCacheValid()
    status("document_id") = mCurrentDocumentInfo.DocumentID
    status("site_name") = mCurrentSiteInfo.SiteName
    
    Set GetSharePointModuleStatus = status
End Function

Public Sub CleanupSharePointIntegration()
    ' Clean up SharePoint integration resources
    
    LogInfo "sharepoint_cleanup", "Cleaning up SharePoint integration"
    
    ' Clear cached data
    Set mCachedMetadata = Nothing
    
    ' Reset state
    mSharePointAvailable = False
    mLastMetadataCheck = 0
    
    ' Clear structures
    Dim emptyDocInfo As SharePointDocumentInfo
    mCurrentDocumentInfo = emptyDocInfo
    
    Dim emptySiteInfo As SharePointSiteInfo
    mCurrentSiteInfo = emptySiteInfo
End Sub