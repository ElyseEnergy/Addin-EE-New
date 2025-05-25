Option Explicit

Private m_SelectedRangeAddress As String
Private m_WasCancelled As Boolean
Private m_DefaultRange As String
Private m_UserForm As Object ' To hold the UserForm instance

' Public properties to retrieve results
Public Property Get SelectedRangeAddress() As String
    SelectedRangeAddress = m_SelectedRangeAddress
End Property

Public Property Get WasCancelled() As Boolean
    WasCancelled = m_WasCancelled
End Property

' Method to initialize and show the form
Public Sub ShowForm(formTitle As String, promptMessage As String, Optional defaultAddress As String = "")
    m_WasCancelled = True ' Default to cancelled
    m_SelectedRangeAddress = ""
    m_DefaultRange = defaultAddress
    
    ' Set controls
    Me.Caption = formTitle
    Me.lblTitle.Caption = formTitle
    Me.lblPrompt.Caption = promptMessage
    
    If Len(m_DefaultRange) > 0 Then
        Me.refEditRange.Text = m_DefaultRange
    End If
    
    ' Center the form
    Me.StartUpPosition = 0 ' Center on screen
    
    Me.Show vbModal
End Sub

Private Sub UserForm_Initialize()
    ' Set default properties for controls if needed
    Me.lblTitle.Font.Bold = True
    Me.lblTitle.Font.Size = 12
    
    Me.cmdOK.Caption = "OK"
    Me.cmdOK.Default = True
    
    Me.cmdCancel.Caption = "Cancel"
    Me.cmdCancel.Cancel = True
    
    ' Adjust layout (basic example, can be more sophisticated)
    AdjustLayout
End Sub

Private Sub AdjustLayout()
    Const Padding As Long = 10
    Const ButtonHeight As Long = 25
    Const ButtonWidth As Long = 80
    
    ' Adjust prompt height (simple example, could be dynamic based on text lines)
    Me.lblPrompt.AutoSize = True ' Ensure it can resize
    ' If Me.lblPrompt.Height > SomeMaxHeight Then Me.lblPrompt.Height = SomeMaxHeight ' Optional max height
    
    ' Position RefEdit below prompt
    Me.refEditRange.Top = Me.lblPrompt.Top + Me.lblPrompt.Height + Padding
    Me.refEditRange.Width = Me.InsideWidth - (2 * Padding)
    
    ' Position buttons
    Me.cmdCancel.Top = Me.refEditRange.Top + Me.refEditRange.Height + Padding
    Me.cmdCancel.Left = Me.InsideWidth - ButtonWidth - Padding
    Me.cmdCancel.Width = ButtonWidth
    Me.cmdCancel.Height = ButtonHeight
    
    Me.cmdOK.Top = Me.cmdCancel.Top
    Me.cmdOK.Left = Me.cmdCancel.Left - ButtonWidth - (Padding / 2)
    Me.cmdOK.Width = ButtonWidth
    Me.cmdOK.Height = ButtonHeight
    
    ' Adjust form height
    Dim requiredHeight As Long
    requiredHeight = Me.cmdOK.Top + Me.cmdOK.Height + Padding + (Me.Height - Me.InsideHeight) ' Add chrome height
    Me.Height = requiredHeight
    
    ' Adjust form width if necessary (e.g., based on title or prompt width)
    ' Me.Width = Max(Me.lblTitle.Width, Me.lblPrompt.Width, Me.refEditRange.Width) + (2 * Padding) + (Me.Width - Me.InsideWidth)
    ' For now, assuming a fixed width or width set by ShowForm parameters if added
End Sub


Private Sub cmdOK_Click()
    If ValidateRange(Me.refEditRange.Text) Then
        m_SelectedRangeAddress = Me.refEditRange.Text
        m_WasCancelled = False
        Me.Hide
    Else
        ' Use the existing custom message box for error display
        Dim errorConfig As MessageBoxConfig
        errorConfig.Title = "Invalid Range"
        errorConfig.Message = "The selected range is not valid. Please enter or select a valid Excel range."
        errorConfig.MessageType = ERROR_MESSAGE
        errorConfig.ButtonCount = 1
        errorConfig.Buttons(1).Text = "OK"
        errorConfig.Buttons(1).ButtonType = "primary"
        errorConfig.Buttons(1).IsDefault = True
        
        ' Need a way to call ShowCustomMessageBox or a simplified version from SYS_MessageBox
        ' For now, using VBA.MsgBox as a placeholder if direct call is complex
        ' Ideally, SYS_MessageBox.ShowErrorMessage "Invalid Range", "..."
        VBA.Interaction.MsgBox errorConfig.Message, vbExclamation, errorConfig.Title
        Me.refEditRange.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    m_WasCancelled = True
    m_SelectedRangeAddress = ""
    Me.Hide
End Sub

Private Function ValidateRange(rangeAddress As String) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.Range(rangeAddress)
    On Error GoTo 0
    ValidateRange = Not rng Is Nothing
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        ' Handle 'X' button click like Cancel
        cmdCancel_Click
    End If
End Sub

' Public method to be called by ApplyCorporateStyling
Public Sub StyleForm(colors As Object)
    On Error Resume Next ' Ignore errors if a control doesn't exist
    
    Me.BackColor = colors("background")
    
    ' Title Label
    If Not Me.lblTitle Is Nothing Then
        Me.lblTitle.ForeColor = colors("text")
        ' Potentially use colors("title_bar_text") or similar if defined
        ' Me.lblTitle.BackColor = colors("title_bar") ' If title is on a colored bar
    End If
    
    ' Prompt Label
    If Not Me.lblPrompt Is Nothing Then
        Me.lblPrompt.ForeColor = colors("text")
    End If
    
    ' RefEdit - Styling options are limited for RefEdit.
    ' You might style a frame around it if needed.
    ' Me.refEditRange.BackColor = ... (often system controlled)
    ' Me.refEditRange.ForeColor = ...
    
    ' OK Button (Primary)
    If Not Me.cmdOK Is Nothing Then
        Me.cmdOK.BackColor = colors("primary")
        Me.cmdOK.ForeColor = colors("button_text")
        ' Add border styling if SYS_CoreSystem provides constants for it
        ' Me.cmdOK.BorderColor = colors("button_border")
        ' Me.cmdOK.SpecialEffect = fmSpecialEffectFlat ' Example
    End If
    
    ' Cancel Button (Secondary)
    If Not Me.cmdCancel Is Nothing Then
        ' Assuming secondary buttons have a different color or style
        ' If no specific "secondary_button_bg" color, use a neutral or background-like color
        If Not IsEmpty(colors("secondary_button_bg")) Then
             Me.cmdCancel.BackColor = colors("secondary_button_bg")
        Else
             Me.cmdCancel.BackColor = RGB(220, 220, 220) ' Light gray as a fallback
        End If
        If Not IsEmpty(colors("secondary_button_text")) Then
            Me.cmdCancel.ForeColor = colors("secondary_button_text")
        Else
            Me.cmdCancel.ForeColor = colors("text") ' Default text color
        End If
        ' Me.cmdCancel.BorderColor = colors("button_border")
        ' Me.cmdCancel.SpecialEffect = fmSpecialEffectFlat ' Example
    End If
    
    On Error GoTo 0
End Sub
