' ============================================================================
' SYS_TicketSystem - Ticketing System Integration
' Elyse Energy VBA Ecosystem - Ticketing Component
' Requires: SYS_CoreSystem, SYS_Logger, SYS_ErrorHandler
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
' TICKET CONFIGURATION AND TYPES
' ============================================================================

Public Type TicketData
    Subject As String
    Description As String
    Priority As String
    Category As String
    Source As String
    IncludeLogs As Boolean
    IncludeScreenshot As Boolean
    UserEmail As String
    AttachmentPaths As String
End Type

Public Enum TicketPriority
    LOW_PRIORITY = 1
    MEDIUM_PRIORITY = 2
    HIGH_PRIORITY = 3
    CRITICAL_PRIORITY = 4
    URGENT_PRIORITY = 5
End Enum

Public Enum TicketCategory
    TECHNICAL_ERROR = 1
    USER_INTERFACE = 2
    DATA_ISSUE = 3
    FEATURE_REQUEST = 4
    CALCULATION_ERROR = 5
    PERFORMANCE_ISSUE = 6
    OTHER_ISSUE = 7
End Enum

' ============================================================================
' TICKET SYSTEM STATE
' ============================================================================

Private mTicketFormOpen As Boolean
Private mCurrentTicketData As TicketData
Private mTicketCounter As Long
Private mLastTicketID As String

' ============================================================================
' MAIN TICKET CREATION FUNCTIONS
' ============================================================================

Public Function CreateTicketFromError(errorSource As String, errorNumber As Long, errorDescription As String, Optional includeContext As Boolean = True) As String
    ' Create a support ticket from an error
    
    LogInfo "ticket_create_from_error", "Source: " & errorSource & " | Error: " & errorNumber
    
    ' Prepare ticket data
    Dim ticketData As TicketData
    ticketData.Subject = "[ERROR] " & errorSource & " - Error " & errorNumber
    ticketData.Description = BuildErrorTicketDescription(errorSource, errorNumber, errorDescription, includeContext)
    ticketData.Priority = GetPriorityString(HIGH_PRIORITY)
    ticketData.Category = GetCategoryString(TECHNICAL_ERROR)
    ticketData.Source = "error_handler"
    ticketData.IncludeLogs = True
    ticketData.IncludeScreenshot = False
    
    ' Show ticket form
    CreateTicketFromError = ShowTicketCreationForm(ticketData)
End Function

Public Function CreateTicketFromMessage(messageTitle As String, messageContent As String, msgType As MessageType) As String
    ' Create a support ticket from a message box interaction
    
    LogInfo "ticket_create_from_message", "Title: " & messageTitle & " | Type: " & GetMessageTypeString(msgType)
    
    ' Prepare ticket data
    Dim ticketData As TicketData
    ticketData.Subject = "[AUTO] " & messageTitle
    ticketData.Description = BuildMessageTicketDescription(messageTitle, messageContent, msgType)
    ticketData.Priority = GetPriorityFromMessageType(msgType)
    ticketData.Category = GetCategoryString(USER_INTERFACE)
    ticketData.Source = "message_box"
    ticketData.IncludeLogs = True
    ticketData.IncludeScreenshot = False
    
    ' Show ticket form
    CreateTicketFromMessage = ShowTicketCreationForm(ticketData)
End Function

Public Function CreateManualTicket() As String
    ' Create a manual support ticket
    
    LogInfo "ticket_create_manual", "Manual ticket creation initiated"
    
    ' Prepare empty ticket data
    Dim ticketData As TicketData
    ticketData.Subject = ""
    ticketData.Description = ""
    ticketData.Priority = GetPriorityString(MEDIUM_PRIORITY)
    ticketData.Category = GetCategoryString(OTHER_ISSUE)
    ticketData.Source = "manual"
    ticketData.IncludeLogs = False
    ticketData.IncludeScreenshot = False
    
    ' Show ticket form
    CreateManualTicket = ShowTicketCreationForm(ticketData)
End Function

' ============================================================================
' TICKET FORM MANAGEMENT
' ============================================================================

Private Function ShowTicketCreationForm(ticketData As TicketData) As String
    ' Show the ticket creation form with pre-populated data
    Dim frm As Object ' frmTicketInput
    
    If mTicketFormOpen Then
        ShowTicketCreationForm = "FORM_ALREADY_OPEN"
        Exit Function
    End If
    
    mTicketFormOpen = True
    mCurrentTicketData = ticketData ' Store initial data
    
    ' Generate ticket ID if not already set (e.g. for new manual tickets)
    If mLastTicketID = "" Or ticketData.Source = "manual" Then
        mTicketCounter = GetNextTicketCounter() ' Ensure counter is loaded and incremented
        mLastTicketID = "ELY" & Format(Now, "yyyymmdd") & Format(mTicketCounter, "0000") ' Changed to 4 digits for counter
    End If
    
    Set frm = VBA.UserForms.Add("frmTicketInput")
    frm.ShowForm mCurrentTicketData, "Support Ticket - ID: " & mLastTicketID
    
    ' Apply corporate styling (assuming ApplyCorporateStyling is in SYS_MessageBox or accessible)
    ' Need to ensure SYS_MessageBox.ApplyCorporateStyling can handle frmTicketInput
    ' Or call a local ApplyStyling method if defined in SYS_TicketSystem
    On Error Resume Next ' In case styling function is not ready
    SYS_MessageBox.ApplyCorporateStyling frm, INFO_MESSAGE ' Or a new MessageType for tickets
    On Error GoTo 0

    If frm.Submitted Then
        mCurrentTicketData = frm.TicketDetails ' Get updated data from form
        
        ' Send the ticket
        If SendTicketViaOutlook(mCurrentTicketData) Then
            ShowTicketCreationForm = mLastTicketID
            SaveLastTicketCounter mTicketCounter ' Save counter after successful submission
        Else
            ShowTicketCreationForm = "SEND_FAILED"
        End If
    Else
        ShowTicketCreationForm = "CANCELLED"
        ' If cancelled, we might not want to reuse the mLastTicketID unless it's a retry.
        ' For simplicity, current mLastTicketID might be offered again if user reopens form for same initial event.
    End If
    
    Unload frm
    Set frm = Nothing
    mTicketFormOpen = False
End Function

' Remove or comment out ShowTicketFormPlaceholder as it's replaced by frmTicketInput
' Private Function ShowTicketFormPlaceholder(ticketData As TicketData) As String
'     ' Placeholder implementation for ticket form
'     ' In actual implementation, this would be a rich HTML editor form or a UserForm

'     Dim formHtml As String
'     formHtml = BuildTicketFormHTML(ticketData)

'     ' For now, show simplified input
'     Dim userSubject As String
'     Dim userDescription As String

'     ' --- OLD InputBox calls ---
'     ' userSubject = InputBox("Ticket Subject:", "Create Support Ticket", ticketData.Subject)
'     ' If userSubject = "" Then
'     '     ShowTicketFormPlaceholder = "CANCELLED"
'     '     Exit Function
'     ' End If
'     '
'     ' userDescription = InputBox("Ticket Description:" & vbCrLf & vbCrLf & "Current description:" & vbCrLf & ticketData.Description, "Create Support Ticket", "Please describe your issue in detail...")
'     ' If userDescription = "" Or userDescription = "Please describe your issue in detail..." Then
'     '     ShowTicketFormPlaceholder = "CANCELLED"
'     '     Exit Function
'     ' End If
'     ' --- END OLD InputBox calls ---

'     ' --- REPLACEMENT with SYS_MessageBox (Placeholder - UserForm needed for text input) ---
'     ' TODO: Implement a UserForm (e.g., TicketCreationForm) for proper subject and description input.
'     ' The following is a temporary placeholder.
'     SYS_MessageBox.ShowInfoMessage "Ticket Subject Input (Placeholder)", "Please enter the ticket subject in the actual UserForm. For now, using: " & ticketData.Subject
'     userSubject = ticketData.Subject ' Using pre-filled subject as placeholder cannot get input.
'     If userSubject = "" Then userSubject = "[Subject Placeholder - UserForm Needed]


'     SYS_MessageBox.ShowInfoMessage "Ticket Description Input (Placeholder)", "Please enter the ticket description in the actual UserForm. For now, using a generic description."
'     userDescription = ticketData.Description ' Using pre-filled description as placeholder.
'     If userDescription = "" Then userDescription = "[Description Placeholder - UserForm Needed]"
'     ' --- END REPLACEMENT ---
    
'     ' Update ticket data
'     mCurrentTicketData.Subject = userSubject ' Ensure mCurrentTicketData is updated
'     mCurrentTicketData.Description = userDescription

'     ' Send the ticket
'     If SendTicketViaOutlook(mCurrentTicketData) Then ' Use mCurrentTicketData
'         ShowTicketFormPlaceholder = mLastTicketID
'     Else
'         ShowTicketFormPlaceholder = "SEND_FAILED"
'     End If
' End Function

' Helper function to manage ticket counter (example, could be stored in a hidden sheet or setting)
Private Function GetNextTicketCounter() As Long
    ' Placeholder: In a real app, load from a persistent store (e.g., hidden sheet, registry, settings file)
    ' For now, just increment a module-level variable. This will reset each session.
    ' A more robust solution would be needed for production.
    Static sTicketCounter As Long ' Static to persist across calls within a session
    If sTicketCounter = 0 Then
        sTicketCounter = CLng(GetSetting(Application.Name & " - ElyseAddin", "TicketSystem", "LastTicketNumber", "0"))
    End If
    sTicketCounter = sTicketCounter + 1
    GetNextTicketCounter = sTicketCounter
End Function

Private Sub SaveLastTicketCounter(counter As Long)
    ' Placeholder: Save to a persistent store
    SaveSetting Application.Name & " - ElyseAddin", "TicketSystem", "LastTicketNumber", CStr(counter)
End Sub

' --- Add to SYS_TicketSystem for ComboBox population in frmTicketInput ---
Public Function GetPriorityEnumArray() As Variant
    GetPriorityEnumArray = Array("Low", "Medium", "High", "Critical", "Urgent")
    ' Corresponds to TicketPriority enum, but string representation for UI
    ' Could also build this by iterating enum if VBA supported that easily for custom enums
End Function

Public Function GetCategoryEnumArray() As Variant
    GetCategoryEnumArray = Array("Technical Error", "User Interface", "Data Issue", "Feature Request", "Calculation Error", "Performance Issue", "Other")
    ' Corresponds to TicketCategory enum
End Function

' ============================================================================
' TICKET DESCRIPTION BUILDERS
' ============================================================================

Private Function BuildErrorTicketDescription(errorSource As String, errorNumber As Long, errorDescription As String, includeContext As Boolean) As String
    ' Build comprehensive error ticket description
    
    Dim description As String
    description = "## Error Report - Elyse Energy System" & vbCrLf & vbCrLf
    
    ' Error details
    description = description & "### Error Information" & vbCrLf
    description = description & "- **Source:** " & errorSource & vbCrLf
    description = description & "- **Error Number:** " & errorNumber & vbCrLf
    description = description & "- **Description:** " & errorDescription & vbCrLf
    description = description & "- **Timestamp:** " & FormatTimestamp() & vbCrLf & vbCrLf
    
    If includeContext Then
        ' System context
        description = description & "### System Context" & vbCrLf
        description = description & "- **User:** " & GetUserIdentity() & vbCrLf
        description = description & "- **Session ID:** " & GetSessionID() & vbCrLf
        description = description & "- **Computer:** " & Environ("COMPUTERNAME") & vbCrLf
        description = description & "- **Domain:** " & Environ("USERDOMAIN") & vbCrLf
        description = description & "- **Excel Version:** " & Application.Version & vbCrLf & vbCrLf
        
        ' File context
        description = description & "### File Context" & vbCrLf
        description = description & "- **Workbook:** " & GetCurrentWorkbookName() & vbCrLf
        description = description & "- **Active Sheet:** " & GetActiveSheetName() & vbCrLf
        description = description & "- **Selected Range:** " & GetSelectedRangeAddress() & vbCrLf & vbCrLf
    End If
    
    ' Instructions for user
    description = description & "### Additional Information" & vbCrLf
    description = description & "Please provide any additional context about what you were doing when this error occurred:" & vbCrLf & vbCrLf
    description = description & "1. What operation were you performing?" & vbCrLf
    description = description & "2. What data were you working with?" & vbCrLf
    description = description & "3. Had this worked before?" & vbCrLf
    description = description & "4. Any other relevant details?" & vbCrLf
    
    BuildErrorTicketDescription = description
End Function

Private Function BuildMessageTicketDescription(messageTitle As String, messageContent As String, msgType As MessageType) As String
    ' Build ticket description from message box interaction
    
    Dim description As String
    description = "## User Interface Issue - Elyse Energy System" & vbCrLf & vbCrLf
    
    ' Message details
    description = description & "### Original Message" & vbCrLf
    description = description & "- **Title:** " & messageTitle & vbCrLf
    description = description & "- **Type:** " & GetMessageTypeString(msgType) & vbCrLf
    description = description & "- **Content:** " & messageContent & vbCrLf
    description = description & "- **Timestamp:** " & FormatTimestamp() & vbCrLf & vbCrLf
    
    ' Context
    description = description & "### Context Information" & vbCrLf
    description = description & "- **User:** " & GetUserIdentity() & vbCrLf
    description = description & "- **Session:** " & GetSessionID() & vbCrLf
    description = description & "- **Workbook:** " & GetCurrentWorkbookName() & vbCrLf
    description = description & "- **Timestamp:** " & FormatTimestamp() & vbCrLf & vbCrLf
    
    ' User input section
    description = description & "### Issue Description" & vbCrLf
    description = description & "Please describe the issue you encountered:" & vbCrLf & vbCrLf
    description = description & "*[Please provide details about the problem you experienced]*"
    
    BuildMessageTicketDescription = description
End Function

' ============================================================================
' EMAIL INTEGRATION (OUTLOOK)
' ============================================================================

Private Function SendTicketViaOutlook(ticketData As TicketData) As Boolean
    ' Send ticket via Outlook email
    On Error GoTo ErrorHandler
    
    LogInfo "ticket_send_outlook", "Ticket ID: " & mLastTicketID & " | Subject: " & ticketData.Subject
    
    ' Create Outlook application
    Dim outlookApp As Object
    Set outlookApp = CreateObject("Outlook.Application")
    
    ' Create mail item
    Dim mailItem As Object
    Set mailItem = outlookApp.CreateItem(0) ' olMailItem
    
    ' Configure email
    With mailItem
        .To = SUPPORT_EMAIL
        .Subject = TICKET_SUBJECT_PREFIX & ticketData.Subject & " [" & mLastTicketID & "]"
        .Body = BuildTicketEmailBodyText(ticketData)
        .HTMLBody = BuildTicketEmailBodyHTML(ticketData)
        
        ' Add attachments if needed
        If ticketData.IncludeLogs Then
            AttachRecentLogs mailItem
        End If
        
        ' Send email
        .Send
    End With
    
    LogInfo "ticket_sent_successfully", "Ticket ID: " & mLastTicketID
    SendTicketViaOutlook = True
    Exit Function
    
ErrorHandler:
    LogError "ticket_send_failed", Err.Number, Err.Description
    SendTicketViaOutlook = False
End Function

Private Function BuildTicketEmailBodyText(ticketData As TicketData) As String
    ' Build plain text email body
    
    Dim body As String
    body = "ELYSE ENERGY SUPPORT TICKET" & vbCrLf
    body = body & "===========================" & vbCrLf & vbCrLf
    
    body = body & "Ticket ID: " & mLastTicketID & vbCrLf
    body = body & "Date: " & FormatTimestamp() & vbCrLf
    body = body & "User: " & GetUserIdentity() & vbCrLf
    body = body & "Priority: " & ticketData.Priority & vbCrLf
    body = body & "Category: " & ticketData.Category & vbCrLf
    body = body & "Source: " & ticketData.Source & vbCrLf & vbCrLf
    
    body = body & "DESCRIPTION:" & vbCrLf
    body = body & "------------" & vbCrLf
    body = body & ticketData.Description & vbCrLf & vbCrLf
    
    body = body & "SYSTEM INFORMATION:" & vbCrLf
    body = body & "-------------------" & vbCrLf
    body = body & "Computer: " & Environ("COMPUTERNAME") & vbCrLf
    body = body & "Domain: " & Environ("USERDOMAIN") & vbCrLf
    body = body & "Excel Version: " & Application.Version & vbCrLf
    body = body & "Workbook: " & GetCurrentWorkbookName() & vbCrLf & vbCrLf
    
    body = body & "---" & vbCrLf
    body = body & "This ticket was generated automatically by the Elyse Energy Excel Add-in."
    
    BuildTicketEmailBodyText = body
End Function

Private Function BuildTicketEmailBodyHTML(ticketData As TicketData) As String
    ' Build HTML email body with styling
    
    Dim html As String
    html = "<html><head><style>"
    html = html & "body { font-family: 'Segoe UI', sans-serif; line-height: 1.6; color: #333; }"
    html = html & ".header { background: linear-gradient(135deg, #2E8B57, #228B22); color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }"
    html = html & ".ticket-info { background: #f8f9fa; padding: 15px; border-left: 4px solid #2E8B57; margin: 15px 0; }"
    html = html & ".description { background: white; padding: 20px; border: 1px solid #e9ecef; border-radius: 5px; margin: 15px 0; }"
    html = html & ".system-info { background: #f1f3f4; padding: 15px; border-radius: 5px; margin: 15px 0; font-size: 0.9em; }"
    html = html & ".footer { border-top: 1px solid #e9ecef; padding-top: 15px; margin-top: 20px; font-size: 0.8em; color: #666; }"
    html = html & "h2 { margin-top: 0; }"
    html = html & "h3 { color: #2E8B57; border-bottom: 1px solid #e9ecef; padding-bottom: 5px; }"
    html = html & ".priority-high { color: #dc3545; font-weight: bold; }"
    html = html & ".priority-medium { color: #ffc107; font-weight: bold; }"
    html = html & ".priority-low { color: #28a745; font-weight: bold; }"
    html = html & "</style></head><body>"
    
    ' Header
    html = html & "<div class='header'>"
    html = html & "<h2>🎫 Elyse Energy Support Ticket</h2>"
    html = html & "<p>Automated support request from Excel Add-in</p>"
    html = html & "</div>"
    
    ' Ticket information
    html = html & "<div class='ticket-info'>"
    html = html & "<h3>Ticket Information</h3>"
    html = html & "<strong>Ticket ID:</strong> " & mLastTicketID & "<br>"
    html = html & "<strong>Date:</strong> " & FormatTimestamp() & "<br>"
    html = html & "<strong>User:</strong> " & GetUserIdentity() & "<br>"
    html = html & "<strong>Priority:</strong> <span class='priority-" & LCase(ticketData.Priority) & "'>" & ticketData.Priority & "</span><br>"
    html = html & "<strong>Category:</strong> " & ticketData.Category & "<br>"
    html = html & "<strong>Source:</strong> " & ticketData.Source
    html = html & "</div>"
    
    ' Description
    html = html & "<div class='description'>"
    html = html & "<h3>Description</h3>"
    html = html & ConvertMarkdownToHTML(ticketData.Description)
    html = html & "</div>"
    
    ' System information
    html = html & "<div class='system-info'>"
    html = html & "<h3>System Information</h3>"
    html = html & "<strong>Computer:</strong> " & Environ("COMPUTERNAME") & "<br>"
    html = html & "<strong>Domain:</strong> " & Environ("USERDOMAIN") & "<br>"
    html = html & "<strong>Excel Version:</strong> " & Application.Version & "<br>"
    html = html & "<strong>Workbook:</strong> " & GetCurrentWorkbookName() & "<br>"
    html = html & "<strong>Active Sheet:</strong> " & GetActiveSheetName() & "<br>"
    html = html & "<strong>Session ID:</strong> " & GetSessionID()
    html = html & "</div>"
    
    ' Footer
    html = html & "<div class='footer'>"
    html = html & "This ticket was generated automatically by the Elyse Energy Excel Add-in.<br>"
    html = html & "For questions about this system, please contact the IT department."
    html = html & "</div>"
    
    html = html & "</body></html>"
    
    BuildTicketEmailBodyHTML = html
End Function

Private Function ConvertMarkdownToHTML(markdown As String) As String
    ' Convert basic markdown to HTML
    
    Dim html As String
    html = markdown
    
    ' Line breaks
    html = Replace(html, vbCrLf, "<br>")
    
    ' Headers
    html = Replace(html, "### ", "<h4>")
    html = Replace(html, "## ", "<h3>")
    html = Replace(html, "# ", "<h2>")
    
    ' Bold
    html = Replace(html, "**", "<strong>")
    
    ' Italic
    html = Replace(html, "*", "<em>")
    
    ' Lists (basic)
    html = Replace(html, "- ", "<li>")
    
    ConvertMarkdownToHTML = html
End Function

Private Sub AttachRecentLogs(mailItem As Object)
    ' Attach recent log file if available
    On Error Resume Next
    
    ' This would attach actual log files in a real implementation
    ' For now, just log that we would attach logs
    LogInfo "ticket_logs_attached", "Recent logs would be attached to ticket " & mLastTicketID
    
    On Error GoTo 0
End Sub

' ============================================================================
' PUBLIC API FUNCTIONS
' ============================================================================

Public Function CreateQuickErrorTicket(errorMsg As String, Optional errorCode As Long = 0) As String
    ' Quick function to create error ticket
    CreateQuickErrorTicket = CreateTicketFromError("Quick Error", errorCode, errorMsg, True)
End Function

Public Function CreateQuickFeedbackTicket(subject As String, feedback As String) As String
    ' Quick function to create feedback ticket
    
    Dim ticketData As TicketData
    ticketData.Subject = "[FEEDBACK] " & subject
    ticketData.Description = feedback
    ticketData.Priority = GetPriorityString(LOW_PRIORITY)
    ticketData.Category = GetCategoryString(FEATURE_REQUEST)
    ticketData.Source = "feedback"
    
    CreateQuickFeedbackTicket = ShowTicketCreationForm(ticketData)
End Function

' ============================================================================
' RIBBON INTEGRATION
' ============================================================================

Public Sub ShowTicketCreationFromRibbon()
    ' Show ticket creation form from ribbon button
    
    LogRibbonAction "btn_create_ticket"
    CreateManualTicket
End Sub

' ============================================================================
' MODULE STATUS AND DIAGNOSTICS
' ============================================================================

Public Function GetTicketSystemStatus() As Object
    ' Get status of ticket system
    
    Dim status As Object
    Set status = CreateObject("Scripting.Dictionary")
    
    status("module_loaded") = True
    status("ticket_form_open") = mTicketFormOpen
    status("tickets_created_this_session") = mTicketCounter
    status("last_ticket_id") = mLastTicketID
    status("support_email") = SUPPORT_EMAIL
    
    Set GetTicketSystemStatus = status
End Function

Public Function GetPriorityString(priorityValue As Variant) As String
    If IsNumeric(priorityValue) Then
        Dim p As TicketPriority
        p = CLng(priorityValue)
        Select Case p
            Case LOW_PRIORITY: GetPriorityString = "Low"
            Case MEDIUM_PRIORITY: GetPriorityString = "Medium"
            Case HIGH_PRIORITY: GetPriorityString = "High"
            Case CRITICAL_PRIORITY: GetPriorityString = "Critical"
            Case URGENT_PRIORITY: GetPriorityString = "Urgent"
            Case Else: GetPriorityString = "Medium" ' Default
        End Select
    Else
        GetPriorityString = CStr(priorityValue) ' Assume it's already a string
    End If
End Function

Public Function GetCategoryString(categoryValue As Variant) As String
    If IsNumeric(categoryValue) Then
        Dim c As TicketCategory
        c = CLng(categoryValue)
        Select Case c
            Case TECHNICAL_ERROR: GetCategoryString = "Technical Error"
            Case USER_INTERFACE: GetCategoryString = "User Interface"
            Case DATA_ISSUE: GetCategoryString = "Data Issue"
            Case FEATURE_REQUEST: GetCategoryString = "Feature Request"
            Case CALCULATION_ERROR: GetCategoryString = "Calculation Error"
            Case PERFORMANCE_ISSUE: GetCategoryString = "Performance Issue"
            Case OTHER_ISSUE: GetCategoryString = "Other"
            Case Else: GetCategoryString = "Other" ' Default
        End Select
    Else
        GetCategoryString = CStr(categoryValue) ' Assume it's already a string
    End If
End Function

' Ensure GetPriorityFromMessageType returns a string compatible with ComboBox
Public Function GetPriorityFromMessageType(msgType As MessageType) As String
    Select Case msgType
        Case ERROR_MESSAGE, CRITICAL_MESSAGE
            GetPriorityFromMessageType = GetPriorityString(CRITICAL_PRIORITY)
        Case WARNING_MESSAGE
            GetPriorityFromMessageType = GetPriorityString(HIGH_PRIORITY)
        Case INFO_MESSAGE, SUCCESS_MESSAGE, GENERAL_MESSAGE
            GetPriorityFromMessageType = GetPriorityString(MEDIUM_PRIORITY)
        Case Else
            GetPriorityFromMessageType = GetPriorityString(MEDIUM_PRIORITY)
    End Select
End Function