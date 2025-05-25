' ============================================================================
' ElyseMessageBox_System - Enhanced MessageBox Components
' Elyse Energy VBA Ecosystem - MessageBox Component
' Requires: ElyseCore_System, ElyseLogger_Module
' ============================================================================

Option Explicit

' ============================================================================
' MODULE DEPENDENCIES
' ============================================================================
' This module requires:
' - ElyseCore_System (enums, constants, utilities)
' - ElyseLogger_Module (logging functions)

' Button size constants
Private Const STANDARD_BUTTON_WIDTH As Long = 80
Private Const STANDARD_BUTTON_HEIGHT As Long = 24
Private Const STANDARD_BUTTON_PADDING As Long = 10

' ============================================================================
' MESSAGEBOX CONFIGURATION CLASSES
' ============================================================================

' Configuration class for MessageBox buttons
Public Type MessageBoxButton
    Text As String
    ActionCallback As String
    IsDefault As Boolean
    ButtonType As String ' "primary", "secondary", "danger"
End Type

' Configuration class for MessageBox setup
Public Type MessageBoxConfig
    Title As String
    Message As String
    MessageType As MessageType
    Buttons(1 To 3) As MessageBoxButton
    ButtonCount As Integer
    ShowIcon As Boolean
    AllowResize As Boolean
    Width As Long
    Height As Long
End Type

' ============================================================================
' GLOBAL MESSAGEBOX STATE
' ============================================================================

Private mMessageBoxOpen As Boolean
Private mCurrentResult As Long
Private mLastMessageConfig As MessageBoxConfig

' ============================================================================
' TYPE 1: LIST SELECTION MESSAGEBOX
' ============================================================================

Public Function ShowListSelectionBox(title As String, message As String, listItems As Collection, Optional defaultSelection As Integer = 1, Optional allowMultiSelect As Boolean = False) As Variant ' Return type changed to Variant for multi-select
    ' Show numbered list selection dialog using frmListSelection
    
    LogInfo "messagebox_list_show", "Title: " & title & " | Items: " & listItems.Count
    
    ' Validate inputs
    If listItems.Count = 0 Then
        If allowMultiSelect Then
            ShowListSelectionBox = Array() ' Return empty array for multi-select
        Else
            ShowListSelectionBox = 0 ' Return 0 for single-select
        End If
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    Dim frm As frmListSelection
    Set frm = New frmListSelection
    
    ' Setup the form
    frm.SetupListForm Title:=title, _
                      Prompt:=message, _
                      Items:=listItems, _
                      DefaultIndex:=defaultSelection, _
                      AllowMultiSelect:=allowMultiSelect,
                      OkButtonCaption:="Select", _
                      CancelButtonCaption:="Cancel"
                      
    ' Apply corporate styling - Assuming INFO_MESSAGE type for styling list selection for now
    ApplyCorporateStyling frm, INFO_MESSAGE
    
    frm.Show vbModal
    
    If frm.WasCancelled Then
        If allowMultiSelect Then
            ShowListSelectionBox = Array() ' Return empty array for multi-select
        Else
            ShowListSelectionBox = 0 ' Consistent with previous behavior for cancellation
        End If
    Else
        If allowMultiSelect Then
            ShowListSelectionBox = frm.SelectedItemIndices ' Returns an array of indices (1-based)
        Else
            ShowListSelectionBox = frm.SelectedItemIndex ' Returns a single index (1-based)
        End If
    End If
    
    LogInfo "messagebox_list_result", "Selection result processed. Cancelled: " & frm.WasCancelled
    
    Unload frm
    Set frm = Nothing
    Exit Function

ErrorHandler:
    LogError "ShowListSelectionBox", "Error: " & Err.Number & " - " & Err.Description
    If allowMultiSelect Then
        ShowListSelectionBox = Array()
    Else
        ShowListSelectionBox = 0
    End If
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
End Function

' ============================================================================
' TYPE 2: RANGE SELECTOR MESSAGEBOX
' ============================================================================

Public Function ShowRangeSelectorBox(formTitle As String, promptMessage As String, Optional defaultAddress As String = "") As String
    ' Function utilisant la fonction native d'Excel pour sélectionner une plage
    Dim selectedRange As Range
    
    On Error GoTo ErrorHandler
    
    ' Afficher le message à l'utilisateur
    If Len(promptMessage) > 0 Then
        MsgBox promptMessage, vbInformation, formTitle
    End If
    
    ' Utiliser la fonction native d'Excel
    On Error Resume Next
    Set selectedRange = Application.InputBox(prompt:=formTitle, _
                                          Type:=8) ' Type 8 = Plage de cellules
    On Error GoTo ErrorHandler
    
    ' Vérifier si l'utilisateur a annulé
    If selectedRange Is Nothing Then
        ShowRangeSelectorBox = ""
    Else
        ShowRangeSelectorBox = selectedRange.Address
    End If
    
    Exit Function

ErrorHandler:
    LogError "ShowRangeSelectorBox", "Error: " & Err.Number & " - " & Err.Description
    ShowRangeSelectorBox = ""
End Function

' ============================================================================
' TYPE 3: MARKDOWN INFORMATION MESSAGEBOX
' ============================================================================

Public Function ShowMarkdownInfoBox(title As String, markdownContent As String, Optional width As Long = 600, Optional height As Long = 500) As Long
    ' Show scrollable markdown information dialog
    
    LogInfo "messagebox_markdown_show", "Title: " & title & " | Content length: " & Len(markdownContent)
    
    On Error GoTo ErrorHandler
    
    Dim frm As frmMarkdownDisplay
    Set frm = New frmMarkdownDisplay
    
    ' Call the setup method on the form
    frm.ShowContent formTitle:=title, _
                    markdownContent:=markdownContent, _
                    formWidth:=width, _
                    formHeight:=height
    
    ' Apply corporate styling - Assuming INFO_MESSAGE type for styling markdown display
    ApplyCorporateStyling frm, INFO_MESSAGE
                    
    frm.Show vbModal
    
    ' This type of form usually doesn't return a specific result other than being closed
    ShowMarkdownInfoBox = 1 ' Or 0, depending on how you want to interpret "closed"
    
    Unload frm
    Set frm = Nothing
    
    LogInfo "messagebox_markdown_result", "Dialog closed"
    Exit Function

ErrorHandler:
    LogError "ShowMarkdownInfoBox", "Error: " & Err.Number & " - " & Err.Description
    ShowMarkdownInfoBox = 0 ' Indicate failure or closure due to error
    If Not frm Is Nothing Then
        Unload frm
        Set frm = Nothing
    End If
End Function

' ============================================================================
' TYPE 4: STANDARD OK/CANCEL MESSAGEBOX
' ============================================================================

Public Function ShowOKCancelBox(title As String, message As String, Optional defaultButton As String = "OK") As Boolean
    ' Show standard OK/Cancel dialog with corporate styling
    
    LogInfo "messagebox_okcancel_show", "Title: " & title
    
    ' Create configuration
    Dim config As MessageBoxConfig
    config.Title = title
    config.Message = message
    config.MessageType = CONFIRMATION_MESSAGE
    config.ButtonCount = 2
    
    ' Configure buttons
    config.Buttons(1).Text = "OK"
    config.Buttons(1).ButtonType = "primary"
    config.Buttons(1).IsDefault = (defaultButton = "OK")
    
    config.Buttons(2).Text = "Cancel"
    config.Buttons(2).ButtonType = "secondary"
    config.Buttons(2).IsDefault = (defaultButton = "Cancel")
    
    ' Show dialog
    Dim result As Long
    result = ShowCustomMessageBox(config)
    
    Dim success As Boolean
    success = (result = 1) ' 1 = OK, 2 = Cancel
    
    LogInfo "messagebox_okcancel_result", "Result: " & IIf(success, "OK", "Cancel")
    
    ShowOKCancelBox = success
End Function

' ============================================================================
' ENHANCED MESSAGEBOX WITH TICKET INTEGRATION
' ============================================================================

Public Function ShowEnhancedMessageBox(title As String, message As String, msgType As MessageType, Optional buttons As String = "OK", Optional allowTicketCreation As Boolean = False) As String
    ' Show enhanced message box with optional ticket creation
    
    LogInfo "messagebox_enhanced_show", "Title: " & title & " | Type: " & GetMessageTypeString(msgType)
    
    ' Create configuration
    Dim config As MessageBoxConfig
    config.Title = title
    config.Message = message
    config.MessageType = msgType
    config.ShowIcon = True
    
    ' Parse button configuration
    ParseButtonConfiguration config, buttons, allowTicketCreation
    
    ' Show dialog
    Dim result As Long
    result = ShowCustomMessageBox(config)
    
    ' Handle special actions
    Dim resultString As String
    resultString = HandleMessageBoxResult(result, config, allowTicketCreation)
    
    LogInfo "messagebox_enhanced_result", "Result: " & resultString
    
    ShowEnhancedMessageBox = resultString
End Function

Private Sub ParseButtonConfiguration(ByRef config As MessageBoxConfig, buttonString As String, allowTicket As Boolean)
    ' Parse button configuration string
    
    Dim buttonArray() As String
    buttonArray = Split(buttonString, ",")
    
    config.ButtonCount = UBound(buttonArray) + 1
    If config.ButtonCount > 3 Then config.ButtonCount = 3
    
    Dim i As Integer
    For i = 0 To config.ButtonCount - 1
        config.Buttons(i + 1).Text = Trim(buttonArray(i))
        config.Buttons(i + 1).ButtonType = GetButtonTypeFromText(config.Buttons(i + 1).Text)
        config.Buttons(i + 1).IsDefault = (i = 0)
    Next i
    
    ' Add ticket button if allowed and appropriate
    If allowTicket And config.MessageType = ERROR_MESSAGE And config.ButtonCount < 3 Then
        config.ButtonCount = config.ButtonCount + 1
        config.Buttons(config.ButtonCount).Text = "Create Ticket"
        config.Buttons(config.ButtonCount).ButtonType = "secondary"
        config.Buttons(config.ButtonCount).IsDefault = False
    End If
End Sub

Private Function GetButtonTypeFromText(buttonText As String) As String
    ' Determine button type from text
    
    Select Case LCase(Trim(buttonText))
        Case "ok", "yes", "confirm", "save", "submit"
            GetButtonTypeFromText = "primary"
        Case "cancel", "no", "close", "exit"
            GetButtonTypeFromText = "secondary"
        Case "delete", "remove", "reset"
            GetButtonTypeFromText = "danger"
        Case Else
            GetButtonTypeFromText = "secondary"
    End Select
End Function

Private Function HandleMessageBoxResult(result As Long, config As MessageBoxConfig, allowTicket As Boolean) As String
    ' Handle message box result and special actions
    
    If result <= 0 Or result > config.ButtonCount Then
        HandleMessageBoxResult = "CANCELLED"
        Exit Function
    End If
    
    Dim selectedButton As String
    selectedButton = config.Buttons(result).Text
    
    ' Check for special button actions
    If LCase(selectedButton) = "create ticket" And allowTicket Then
        ' Trigger ticket creation (will be handled by ticket module)
        HandleMessageBoxResult = "CREATE_TICKET"
    Else
        HandleMessageBoxResult = selectedButton
    End If
End Function

' ============================================================================
' CORE MESSAGEBOX DISPLAY ENGINE
' ============================================================================

Private Function ShowCustomMessageBox(config As MessageBoxConfig) As Long
    ' Core function to display custom message box
    
    ' For actual implementation, this would create and show a UserForm
    ' This is a placeholder that demonstrates the interface
    
    mMessageBoxOpen = True
    mLastMessageConfig = config
    
    ' Create the form based on configuration
    Dim formResult As Long
    formResult = CreateAndShowMessageBoxForm(config)
    
    mMessageBoxOpen = False
    ShowCustomMessageBox = formResult
End Function

Private Function CreateAndShowMessageBoxForm(config As MessageBoxConfig) As Long
    ' Create and display the actual message box form
    
    ' --- START REFACTOR ---
    ' This function will now instantiate, configure, and show frmCustomMessageBox
    
    On Error GoTo FallbackToMsgBox ' If UserForm fails, fallback to standard MsgBox

    Dim customForm As frmCustomMessageBox ' Assurez-vous que frmCustomMessageBox est créé dans l'éditeur VBA
    Set customForm = New frmCustomMessageBox
    
    ' Initialize the form with the configuration
    customForm.InitializeFromConfig config
    
    ' Apply corporate styling to the UserForm
    ApplyCorporateStyling customForm, config.MessageType
    
    ' Show the form modally
    customForm.Show vbModal
    
    ' Get the result from the form (assuming frmCustomMessageBox has a public property 'ClickedButtonIndex')
    CreateAndShowMessageBoxForm = customForm.ClickedButtonIndex
    
    ' Clean up
    Unload customForm
    Set customForm = Nothing
    
    Exit Function

FallbackToMsgBox:
    ' This is the original MsgBox fallback logic
    LogError "CreateAndShowMessageBoxForm", "Failed to load custom UserForm. Error: " & Err.Number & " - " & Err.Description & ". Falling back to standard VBA MsgBox."
    Err.Clear ' Clear error before attempting fallback

    Dim msgBoxStyle As VbMsgBoxStyle
    msgBoxStyle = ConvertMessageTypeToMsgBoxStyle(config.MessageType)
    
    ' Add buttons based on configuration
    If config.ButtonCount = 0 Then
        msgBoxStyle = msgBoxStyle + vbOKOnly
    ElseIf config.ButtonCount = 1 Then
        msgBoxStyle = msgBoxStyle + vbOKOnly
    ElseIf config.ButtonCount = 2 Then
        If (LCase(config.Buttons(1).Text) = "ok" And LCase(config.Buttons(2).Text) = "cancel") Or _
           (LCase(config.Buttons(1).Text) = "cancel" And LCase(config.Buttons(2).Text) = "ok") Then
            msgBoxStyle = msgBoxStyle + vbOKCancel
        ElseIf (LCase(config.Buttons(1).Text) = "yes" And LCase(config.Buttons(2).Text) = "no") Or _
               (LCase(config.Buttons(1).Text) = "no" And LCase(config.Buttons(2).Text) = "yes") Then
            msgBoxStyle = msgBoxStyle + vbYesNo
        Else
            msgBoxStyle = msgBoxStyle + vbOKCancel ' Fallback
        End If
    ElseIf config.ButtonCount = 3 Then
        msgBoxStyle = msgBoxStyle + vbYesNoCancel
    Else
        msgBoxStyle = msgBoxStyle + vbOKOnly ' Fallback
    End If
    
    Dim result As VbMsgBoxResult
    result = VBA.Interaction.MsgBox(config.Message, msgBoxStyle, config.Title)
    
    ' Map VBA.MsgBox result back to our 1, 2, 3... convention
    Select Case result
        Case vbOK: CreateAndShowMessageBoxForm = GetButtonIndexByText(config, "OK")
                  If CreateAndShowMessageBoxForm = 0 And msgBoxStyle = vbOKOnly Then CreateAndShowMessageBoxForm = 1
        Case vbCancel: CreateAndShowMessageBoxForm = GetButtonIndexByText(config, "Cancel")
                      If CreateAndShowMessageBoxForm = 0 Then
                          ' If "Cancel" was the second button in a 2-button config (OK/Cancel, Yes/No)
                          If config.ButtonCount = 2 Then CreateAndShowMessageBoxForm = 2 Else CreateAndShowMessageBoxForm = 0
                      End If
        Case vbYes: CreateAndShowMessageBoxForm = GetButtonIndexByText(config, "Yes")
                   If CreateAndShowMessageBoxForm = 0 Then CreateAndShowMessageBoxForm = 1 ' Common fallback for Yes
        Case vbNo: CreateAndShowMessageBoxForm = GetButtonIndexByText(config, "No")
                  If CreateAndShowMessageBoxForm = 0 Then
                      ' If "No" was the second button in a Yes/No config
                      If config.ButtonCount >= 2 And (msgBoxStyle And vbYesNo) = vbYesNo Then CreateAndShowMessageBoxForm = 2 Else CreateAndShowMessageBoxForm = 0
                  End If
        ' ... (keep other Case statements for vbAbort, vbRetry, vbIgnore if needed, though less common with custom forms) ...
        Case Else: CreateAndShowMessageBoxForm = 0 ' Unknown or closed via 'X'
    End Select
    
    ' Refined fallback mapping for standard MsgBox results
    If CreateAndShowMessageBoxForm = 0 Then ' If no specific button text matched
        If (msgBoxStyle And vbOKOnly) = vbOKOnly And result = vbOK Then
            CreateAndShowMessageBoxForm = 1
        ElseIf (msgBoxStyle And vbOKCancel) = vbOKCancel Then
            If result = vbOK Then CreateAndShowMessageBoxForm = GetButtonIndexByText(config, "ok") ' Prefer configured text
            If CreateAndShowMessageBoxForm = 0 And result = vbOK Then CreateAndShowMessageBoxForm = 1
            If result = vbCancel Then CreateAndShowMessageBoxForm = GetButtonIndexByText(config, "cancel")
            If CreateAndShowMessageBoxForm = 0 And result = vbCancel Then CreateAndShowMessageBoxForm = 2
        ElseIf (msgBoxStyle And vbYesNo) = vbYesNo Then
            If result = vbYes Then CreateAndShowMessageBoxForm = GetButtonIndexByText(config, "yes")
            If CreateAndShowMessageBoxForm = 0 And result = vbYes Then CreateAndShowMessageBoxForm = 1
            If result = vbNo Then CreateAndShowMessageBoxForm = GetButtonIndexByText(config, "no")
            If CreateAndShowMessageBoxForm = 0 And result = vbNo Then CreateAndShowMessageBoxForm = 2
        ElseIf (msgBoxStyle And vbYesNoCancel) = vbYesNoCancel Then
            If result = vbYes Then CreateAndShowMessageBoxForm = GetButtonIndexByText(config, "yes")
            If CreateAndShowMessageBoxForm = 0 And result = vbYes Then CreateAndShowMessageBoxForm = 1
            If result = vbNo Then CreateAndShowMessageBoxForm = GetButtonIndexByText(config, "no")
            If CreateAndShowMessageBoxForm = 0 And result = vbNo Then CreateAndShowMessageBoxForm = 2
            If result = vbCancel Then CreateAndShowMessageBoxForm = GetButtonIndexByText(config, "cancel")
            If CreateAndShowMessageBoxForm = 0 And result = vbCancel Then CreateAndShowMessageBoxForm = 3
        End If
    End If
    ' --- END REPLACEMENT ---
End Function

Private Function GetButtonIndexByText(config As MessageBoxConfig, buttonText As String) As Long
    Dim i As Integer
    GetButtonIndexByText = 0 ' Default to 0 if not found
    For i = 1 To config.ButtonCount
        If LCase(config.Buttons(i).Text) = LCase(buttonText) Then
            GetButtonIndexByText = i
            Exit Function
        End If
    Next i
End Function

' ============================================================================
' CORPORATE STYLING UTILITIES
' ============================================================================

Public Function GetCorporateColorScheme(msgType As MessageType) As Object
    ' Get corporate color scheme for message type
    
    Dim colors As Object
    Set colors = CreateObject("Scripting.Dictionary")
    
    ' Default colors
    colors("background") = COLOR_BACKGROUND
    colors("text") = COLOR_TEXT
    colors("button_text") = COLOR_BACKGROUND ' Assuming buttons have a solid background, so text is light
    colors("button_border") = COLOR_NEUTRAL

    Select Case msgType
        Case INFO_MESSAGE
            colors("primary") = COLOR_PRIMARY ' Or COLOR_SECONDARY if more appropriate for info
            colors("title_bar") = COLOR_PRIMARY
            colors("icon_area_background") = COLOR_PRIMARY
        Case SUCCESS_MESSAGE
            colors("primary") = COLOR_SUCCESS
            colors("title_bar") = COLOR_SUCCESS
            colors("icon_area_background") = COLOR_SUCCESS
        Case WARNING_MESSAGE
            colors("primary") = COLOR_WARNING
            colors("title_bar") = COLOR_WARNING
            colors("icon_area_background") = COLOR_WARNING
            colors("text") = COLOR_TEXT ' Warnings might need dark text on lighter warning color
            colors("button_text") = COLOR_TEXT
        Case ERROR_MESSAGE
            colors("primary") = COLOR_ERROR
            colors("title_bar") = COLOR_ERROR
            colors("icon_area_background") = COLOR_ERROR
        Case CONFIRMATION_MESSAGE
            colors("primary") = COLOR_SECONDARY ' Or COLOR_NEUTRAL for confirmations
            colors("title_bar") = COLOR_SECONDARY
            colors("icon_area_background") = COLOR_SECONDARY
    End Select
    
    Set GetCorporateColorScheme = colors
End Function

Public Sub ApplyCorporateStyling(frm As Object, messageType As MessageType)
    Dim colors As Object
    Set colors = GetCorporateColorScheme(messageType)

    If colors Is Nothing Then
        Exit Sub
    End If

    On Error Resume Next ' General error handler for styling

    frm.BackColor = colors("background")

    ' Common styling for known form types
    If TypeName(frm) = "frmCustomMessageBox" Then
        StandardizeButtonSizes frm
    ElseIf TypeName(frm) = "frmListSelection" Then
        StandardizeButtonSizes frm
    ElseIf TypeName(frm) = "frmMarkdownDisplay" Then
        StandardizeButtonSizes frm
    ElseIf TypeName(frm) = "frmRangeSelector" Then
        StandardizeButtonSizes frm
    ElseIf TypeName(frm) = "frmTicketInput" Then
        StandardizeButtonSizes frm
    End If

    On Error GoTo 0
End Sub

Private Sub StandardizeButtonSizes(frm As Object)
    Dim ctrl As MSForms.Control
    On Error Resume Next
    
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "CommandButton" Then
            ctrl.Width = STANDARD_BUTTON_WIDTH
            ctrl.Height = STANDARD_BUTTON_HEIGHT
        End If
    Next ctrl
    
    ' Réorganiser les boutons horizontalement si nécessaire
    Dim totalButtons As Long
    Dim currentX As Long
    Dim bottomPadding As Long
    
    totalButtons = 0
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "CommandButton" Then
            totalButtons = totalButtons + 1
        End If
    Next ctrl
    
    If totalButtons > 0 Then
        ' Calculer la position de départ pour centrer les boutons
        bottomPadding = 12 ' Espace entre le bas du formulaire et les boutons
        currentX = (frm.Width - (totalButtons * STANDARD_BUTTON_WIDTH + (totalButtons - 1) * STANDARD_BUTTON_PADDING)) / 2
        
        For Each ctrl In frm.Controls
            If TypeName(ctrl) = "CommandButton" Then
                ctrl.Left = currentX
                ctrl.Top = frm.Height - STANDARD_BUTTON_HEIGHT - bottomPadding
                currentX = currentX + STANDARD_BUTTON_WIDTH + STANDARD_BUTTON_PADDING
            End If
        Next ctrl
    End If
    
    On Error GoTo 0
End Sub

' ============================================================================
' PUBLIC CONVENIENCE FUNCTIONS
' ============================================================================

Public Function ShowInfoMessage(title As String, message As String) As String
    ' Convenience function for info messages
    ShowInfoMessage = ShowEnhancedMessageBox(title, message, INFO_MESSAGE, "OK")
End Function

Public Function ShowSuccessMessage(title As String, message As String) As String
    ' Convenience function for success messages
    ShowSuccessMessage = ShowEnhancedMessageBox(title, message, SUCCESS_MESSAGE, "OK")
End Function

Public Function ShowWarningMessage(title As String, message As String) As String
    ' Convenience function for warning messages
    ShowWarningMessage = ShowEnhancedMessageBox(title, message, WARNING_MESSAGE, "OK")
End Function

Public Function ShowErrorMessage(title As String, message As String, Optional allowTicket As Boolean = True) As String
    ' Convenience function for error messages with ticket option
    ShowErrorMessage = ShowEnhancedMessageBox(title, message, ERROR_MESSAGE, "OK", allowTicket)
End Function

Public Function ShowConfirmationMessage(title As String, message As String) As Boolean
    ' Convenience function for confirmation messages
    Dim result As String
    result = ShowEnhancedMessageBox(title, message, CONFIRMATION_MESSAGE, "Yes,No")
    ShowConfirmationMessage = (LCase(result) = "yes")
End Function

' ============================================================================
' MESSAGEBOX TEMPLATES FOR ELYSE ENERGY
' ============================================================================

Public Function ShowProductionDataConfirmation(operation As String, dataDescription As String) As Boolean
    ' Template for production data confirmations
    
    Dim title As String
    Dim message As String
    
    title = "Production Data " & operation
    message = "Confirm " & LCase(operation) & " for:" & vbCrLf & vbCrLf & _
              dataDescription & vbCrLf & vbCrLf & _
              "This action will affect production tracking data."
    
    ShowProductionDataConfirmation = ShowConfirmationMessage(title, message)
End Function

Public Function ShowCalculationErrorWithSupport(calculationType As String, errorDetails As String) As String
    ' Template for calculation errors with support option
    
    Dim title As String
    Dim message As String
    
    title = "Calculation Error - " & calculationType
    message = "An error occurred during " & LCase(calculationType) & " calculation:" & vbCrLf & vbCrLf & _
              errorDetails & vbCrLf & vbCrLf & _
              "Please verify your input data or contact support for assistance."
    
    ShowCalculationErrorWithSupport = ShowErrorMessage(title, message, True)
End Function

Public Function ShowDataExportSuccess(exportType As String, recordCount As Long, destination As String) As String
    ' Template for successful data export
    
    Dim title As String
    Dim message As String
    
    title = "Export Complete"
    message = exportType & " export completed successfully:" & vbCrLf & vbCrLf & _
              "Records exported: " & Format(recordCount, "#,##0") & vbCrLf & _
              "Destination: " & destination & vbCrLf & vbCrLf & _
              "The data is now available for use."
    
    ShowDataExportSuccess = ShowSuccessMessage(title, message)
End Function

' ============================================================================
' MODULE STATUS AND DIAGNOSTICS
' ============================================================================

Public Function GetMessageBoxSystemStatus() As Object
    ' Get status of message box system
    
    Dim status As Object
    Set status = CreateObject("Scripting.Dictionary")
    
    status("module_loaded") = True
    status("message_box_open") = mMessageBoxOpen
    status("last_message_title") = mLastMessageConfig.Title
    status("supported_types") = "list_selection,range_selector,markdown_info,ok_cancel,enhanced"
    
    Set GetMessageBoxSystemStatus = status
End Function