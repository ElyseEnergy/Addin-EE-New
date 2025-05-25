Option Explicit

Private m_Submitted As Boolean
Private m_TicketData As TicketData

' Public property to get the result
Public Property Get Submitted() As Boolean
    Submitted = m_Submitted
End Property

Public Property Get TicketDetails() As TicketData
    TicketDetails = m_TicketData
End Property

Public Sub ShowForm(initialData As TicketData, Optional formTitle As String = "Support Ticket")
    m_Submitted = False ' Default to not submitted
    m_TicketData = initialData

    Me.Caption = formTitle
    Me.lblFormTitle.Caption = formTitle ' Assuming a label for the title exists

    PopulateFormFields initialData
    PopulateComboBoxes
    
    Me.StartUpPosition = 0 ' CenterScreen
    AdjustLayout ' Call AdjustLayout after populating fields
    Me.Show vbModal
End Sub

Private Sub UserForm_Initialize()
    ' Set up default control properties
    Me.lblFormTitle.Font.Bold = True
    Me.lblFormTitle.Font.Size = 12

    Me.txtDescription.MultiLine = True
    Me.txtDescription.ScrollBars = fmScrollBarsVertical
    Me.txtDescription.EnterKeyBehavior = True ' Allow Enter key for new lines
    Me.txtDescription.WordWrap = True

    Me.cmdSubmit.Caption = "Submit Ticket"
    Me.cmdSubmit.Default = True
    Me.cmdCancel.Caption = "Cancel"
    Me.cmdCancel.Cancel = True
    
    ' Set a default size, AdjustLayout will refine it
    Me.Width = Application.PointsToPixels(450, 0) ' Default width in points
    Me.Height = Application.PointsToPixels(550, 1) ' Default height in points
End Sub

Private Sub PopulateFormFields(data As TicketData)
    Me.txtSubject.Text = data.Subject
    Me.txtDescription.Text = data.Description
    Me.chkIncludeLogs.Value = data.IncludeLogs
    Me.chkIncludeScreenshot.Value = data.IncludeScreenshot
    
    ' Set ComboBoxes - requires helper functions in SYS_TicketSystem or here
    SetComboBoxValue Me.cmbPriority, data.Priority
    SetComboBoxValue Me.cmbCategory, data.Category
End Sub

Private Sub PopulateComboBoxes()
    Dim priorities As Variant
    Dim categories As Variant
    Dim i As Long

    ' These functions need to be available (e.g., from SYS_TicketSystem or defined locally)
    On Error Resume Next ' In case functions are not yet available
    priorities = SYS_TicketSystem.GetPriorityEnumArray() ' Assumes this returns a 0-based array of strings
    categories = SYS_TicketSystem.GetCategoryEnumArray() ' Assumes this returns a 0-based array of strings
    On Error GoTo 0

    If IsArray(priorities) Then
        Me.cmbPriority.Clear
        For i = LBound(priorities) To UBound(priorities)
            Me.cmbPriority.AddItem priorities(i)
        Next i
    Else
        Debug.Print "frmTicketInput.PopulateComboBoxes: Failed to load priorities."
        ' Add default items if loading failed
        Me.cmbPriority.AddItem "Low"
        Me.cmbPriority.AddItem "Medium"
        Me.cmbPriority.AddItem "High"
        Me.cmbPriority.AddItem "Critical"
        Me.cmbPriority.AddItem "Urgent"
    End If

    If IsArray(categories) Then
        Me.cmbCategory.Clear
        For i = LBound(categories) To UBound(categories)
            Me.cmbCategory.AddItem categories(i)
        Next i
    Else
        Debug.Print "frmTicketInput.PopulateComboBoxes: Failed to load categories."
        ' Add default items if loading failed
        Me.cmbCategory.AddItem "Technical Error"
        Me.cmbCategory.AddItem "User Interface"
        Me.cmbCategory.AddItem "Data Issue"
        Me.cmbCategory.AddItem "Feature Request"
        Me.cmbCategory.AddItem "Other"
    End If
End Sub

Private Sub SetComboBoxValue(cmb As MSForms.ComboBox, valueToSet As String)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = valueToSet Then
            cmb.ListIndex = i
            Exit Sub
        End If
    Next i
    ' If value not found, optionally set to a default or leave blank
    If cmb.ListCount > 0 And valueToSet = "" Then
        ' cmb.ListIndex = 0 ' Select first item if value is empty
    ElseIf valueToSet <> "" Then
        ' cmb.AddItem valueToSet ' Or add it if not present and select it
        ' cmb.value = valueToSet
         Debug.Print "SetComboBoxValue: Value '" & valueToSet & "' not found in " & cmb.Name
    End If
End Sub

Private Sub cmdSubmit_Click()
    If ValidateInputs() Then
        m_TicketData.Subject = Me.txtSubject.Text
        m_TicketData.Description = Me.txtDescription.Text
        m_TicketData.Priority = Me.cmbPriority.Text
        m_TicketData.Category = Me.cmbCategory.Text
        m_TicketData.IncludeLogs = Me.chkIncludeLogs.Value
        m_TicketData.IncludeScreenshot = Me.chkIncludeScreenshot.Value
        ' m_TicketData.UserEmail = ... ' If an email field is added
        
        m_Submitted = True
        Me.Hide
    End If
End Sub

Private Function ValidateInputs() As Boolean
    ValidateInputs = True ' Default to true
    If Trim(Me.txtSubject.Text) = "" Then
        SYS_MessageBox.ShowErrorMessage "Input Required", "Ticket subject cannot be empty."
        Me.txtSubject.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    If Trim(Me.txtDescription.Text) = "" Then
        SYS_MessageBox.ShowWarningMessage "Input Recommended", "Ticket description is empty. Are you sure you want to continue?", vbYesNo
        If SYS_MessageBox.LastButtonPressed <> vbYes Then
             Me.txtDescription.SetFocus
             ValidateInputs = False
             Exit Function
        End If
    End If
    If Me.cmbPriority.ListIndex = -1 And Me.cmbPriority.Text <> "" Then
         ' Allow if text is typed but not in list, though ideally it should be from list
    ElseIf Me.cmbPriority.ListIndex = -1 Then
        SYS_MessageBox.ShowErrorMessage "Input Required", "Please select a ticket priority."
        Me.cmbPriority.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    If Me.cmbCategory.ListIndex = -1 And Me.cmbCategory.Text <> "" Then
        ' Allow if text is typed
    ElseIf Me.cmbCategory.ListIndex = -1 Then
        SYS_MessageBox.ShowErrorMessage "Input Required", "Please select a ticket category."
        Me.cmbCategory.SetFocus
        ValidateInputs = False
        Exit Function
    End If
End Function

Private Sub cmdCancel_Click()
    m_Submitted = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        cmdCancel_Click ' Treat 'X' as Cancel
    End If
End Sub

Public Sub StyleForm(colors As Object)
    On Error Resume Next ' Ignore errors if a control doesn't exist
    Me.BackColor = colors("background")

    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        Select Case TypeName(ctrl)
            Case "Label"
                ctrl.ForeColor = colors("text")
                ctrl.BackColor = colors("background") ' Ensure labels are transparent or match bg
                If ctrl.Name = "lblFormTitle" Then ' Specific styling for title
                    ' ctrl.ForeColor = colors("title_text") ' If defined
                    ' ctrl.Font.Size = 12 ' Already set in Initialize
                    ' ctrl.Font.Bold = True
                End If
            Case "TextBox", "ComboBox"
                ctrl.BackColor = colors("input_bg")
                ctrl.ForeColor = colors("input_text")
                ctrl.BorderColor = colors("input_border")
                ' ctrl.SpecialEffect = fmSpecialEffectFlat ' Or another effect
            Case "CheckBox"
                ctrl.ForeColor = colors("text")
                ctrl.BackColor = colors("background")
            Case "CommandButton"
                If ctrl.Name = "cmdSubmit" Then
                    ctrl.BackColor = colors("primary")
                    ctrl.ForeColor = colors("button_text")
                ElseIf ctrl.Name = "cmdCancel" Then
                    If Not IsEmpty(colors("secondary_button_bg")) Then
                        ctrl.BackColor = colors("secondary_button_bg")
                    Else
                        ctrl.BackColor = RGB(220, 220, 220) ' Fallback
                    End If
                    If Not IsEmpty(colors("secondary_button_text")) Then
                        ctrl.ForeColor = colors("secondary_button_text")
                    Else
                        ctrl.ForeColor = colors("text") ' Fallback
                    End If
                Else
                    ctrl.BackColor = colors("neutral_button_bg") ' e.g. for other buttons if any
                    ctrl.ForeColor = colors("text")
                End If
                ' ctrl.BorderColor = colors("button_border")
                ' ctrl.SpecialEffect = fmSpecialEffectFlat
            Case "Frame"
                ctrl.BorderColor = colors("border")
                ctrl.ForeColor = colors("text") ' For frame caption
        End Select
    Next ctrl
    On Error GoTo 0
End Sub

Private Sub AdjustLayout()
    Const BaseMargin As Long = 10 ' Points
    Const ControlSpacing As Long = 6 ' Points
    Const LabelHeight As Long = 15 ' Points
    Const TextBoxHeight As Long = 18 ' Points (single line)
    Const ComboBoxHeight As Long = 20 ' Points
    Const CheckBoxHeight As Long = 18 ' Points
    Const ButtonHeight As Long = 25 ' Points
    Const ButtonWidth As Long = 90 ' Points
    Const MinDescHeight As Long = 60 ' Points
    Const MaxDescHeight As Long = 200 ' Points

    Dim currentY As Single
    Dim controlX As Single
    Dim labelWidth As Single
    Dim controlWidth As Single
    Dim marginPx As Long, spacingPx As Long, lblHeightPx As Long, txtHeightPx As Long
    Dim cmbHeightPx As Long, chkHeightPx As Long, btnHeightPx As Long, btnWidthPx As Long
    Dim minDescHeightPx As Long, maxDescHeightPx As Long
    Dim descHeightPx As Long

    ' Convert points to pixels for positioning
    marginPx = Application.PointsToPixels(BaseMargin, 1)
    spacingPx = Application.PointsToPixels(ControlSpacing, 1)
    lblHeightPx = Application.PointsToPixels(LabelHeight, 1)
    txtHeightPx = Application.PointsToPixels(TextBoxHeight, 1)
    cmbHeightPx = Application.PointsToPixels(ComboBoxHeight, 1)
    chkHeightPx = Application.PointsToPixels(CheckBoxHeight, 1)
    btnHeightPx = Application.PointsToPixels(ButtonHeight, 1)
    btnWidthPx = Application.PointsToPixels(ButtonWidth, 0)
    minDescHeightPx = Application.PointsToPixels(MinDescHeight, 1)
    maxDescHeightPx = Application.PointsToPixels(MaxDescHeight, 1)

    currentY = marginPx
    controlX = marginPx * 2 + Application.PointsToPixels(80, 0) ' X position for input controls (after labels)
    labelWidth = Application.PointsToPixels(75, 0) ' Width for labels
    controlWidth = Me.InsideWidth - controlX - marginPx
    If controlWidth < Application.PointsToPixels(100,0) Then controlWidth = Application.PointsToPixels(100,0)

    ' Form Title Label (lblFormTitle)
    If Not Me.lblFormTitle Is Nothing Then
        Me.lblFormTitle.Left = marginPx
        Me.lblFormTitle.Top = currentY
        Me.lblFormTitle.Width = Me.InsideWidth - (2 * marginPx)
        Me.lblFormTitle.Height = lblHeightPx * 1.5 ' Larger for title
        currentY = currentY + Me.lblFormTitle.Height + spacingPx
    End If

    ' Subject
    PositionControlPair Me.lblSubject, Me.txtSubject, "Subject:", currentY, marginPx, labelWidth, controlX, controlWidth, lblHeightPx, txtHeightPx, spacingPx
    
    ' Priority
    PositionControlPair Me.lblPriority, Me.cmbPriority, "Priority:", currentY, marginPx, labelWidth, controlX, controlWidth, lblHeightPx, cmbHeightPx, spacingPx

    ' Category
    PositionControlPair Me.lblCategory, Me.cmbCategory, "Category:", currentY, marginPx, labelWidth, controlX, controlWidth, lblHeightPx, cmbHeightPx, spacingPx

    ' Description Label
    If Not Me.lblDescription Is Nothing Then
        Me.lblDescription.Caption = "Description:"
        Me.lblDescription.Left = marginPx
        Me.lblDescription.Top = currentY
        Me.lblDescription.Width = labelWidth
        Me.lblDescription.Height = lblHeightPx
        currentY = currentY + Me.lblDescription.Height + Application.PointsToPixels(2,1) ' Small gap before textbox
    End If
    
    ' Description TextBox
    If Not Me.txtDescription Is Nothing Then
        Me.txtDescription.Left = marginPx
        Me.txtDescription.Top = currentY
        Me.txtDescription.Width = Me.InsideWidth - (2 * marginPx)
        ' Estimate height based on content, or use fixed range
        descHeightPx = Application.PointsToPixels(100, 1) ' Default description height
        ' Add logic to estimate based on text lines if needed, or use a scrollable fixed height
        If descHeightPx < minDescHeightPx Then descHeightPx = minDescHeightPx
        If descHeightPx > maxDescHeightPx Then descHeightPx = maxDescHeightPx
        Me.txtDescription.Height = descHeightPx
        currentY = currentY + Me.txtDescription.Height + spacingPx
    End If

    ' Checkboxes (Include Logs, Include Screenshot)
    If Not Me.chkIncludeLogs Is Nothing Then
        Me.chkIncludeLogs.Left = marginPx
        Me.chkIncludeLogs.Top = currentY
        Me.chkIncludeLogs.Width = Application.PointsToPixels(150,0)
        Me.chkIncludeLogs.Height = chkHeightPx
        currentY = currentY + Me.chkIncludeLogs.Height + spacingPx
    End If
    
    If Not Me.chkIncludeScreenshot Is Nothing Then
        Me.chkIncludeScreenshot.Left = marginPx
        Me.chkIncludeScreenshot.Top = currentY
        Me.chkIncludeScreenshot.Width = Application.PointsToPixels(150,0)
        Me.chkIncludeScreenshot.Height = chkHeightPx
        currentY = currentY + Me.chkIncludeScreenshot.Height + spacingPx * 2 ' More space before buttons
    End If

    ' Buttons (Submit, Cancel)
    Dim totalButtonWidth As Long
    totalButtonWidth = (2 * btnWidthPx) + spacingPx ' For two buttons
    Dim buttonStartX As Long
    buttonStartX = (Me.InsideWidth - totalButtonWidth) / 2
    If buttonStartX < marginPx Then buttonStartX = marginPx

    If Not Me.cmdSubmit Is Nothing Then
        Me.cmdSubmit.Top = currentY
        Me.cmdSubmit.Left = buttonStartX
        Me.cmdSubmit.Width = btnWidthPx
        Me.cmdSubmit.Height = btnHeightPx
    End If

    If Not Me.cmdCancel Is Nothing Then
        Me.cmdCancel.Top = currentY
        Me.cmdCancel.Left = buttonStartX + btnWidthPx + spacingPx
        Me.cmdCancel.Width = btnWidthPx
        Me.cmdCancel.Height = btnHeightPx
    End If
    
    If Not Me.cmdSubmit Is Nothing Then
      currentY = currentY + Me.cmdSubmit.Height + marginPx
    ElseIf Not Me.cmdCancel Is Nothing Then
      currentY = currentY + Me.cmdCancel.Height + marginPx
    Else
      currentY = currentY + marginPx ' Just bottom margin if no buttons
    End If

    ' Set Form Height
    Me.Height = currentY + (Me.Height - Me.InsideHeight) ' Add title bar and border height
    
    ' Min/Max form height constraints (similar to frmCustomMessageBox)
    Dim screenHeightPx As Long
    screenHeightPx = Application.PointsToPixels(Application.Height, 1)
    If Me.Height < Application.PointsToPixels(300, 1) Then Me.Height = Application.PointsToPixels(300, 1) ' Min height
    If Me.Height > screenHeightPx * 0.85 Then Me.Height = screenHeightPx * 0.85 ' Max height
End Sub

Private Sub PositionControlPair(lbl As MSForms.Label, ctrl As Object, caption As String, ByRef currentY As Single, ByVal marginPx As Long, ByVal labelWidth As Long, ByVal controlX As Long, ByVal controlWidth As Long, ByVal lblHeightPx As Long, ByVal ctrlHeightPx As Long, ByVal spacingPx As Long)
    If Not lbl Is Nothing Then
        lbl.Caption = caption
        lbl.Left = marginPx
        lbl.Top = currentY + (ctrlHeightPx - lblHeightPx) / 2 ' Vertically align label with center of control
        lbl.Width = labelWidth
        lbl.Height = lblHeightPx
        lbl.TextAlign = fmTextAlignRight ' Align label text to the right for neatness
    End If
    
    If Not ctrl Is Nothing Then
        ctrl.Left = controlX
        ctrl.Top = currentY
        ctrl.Width = controlWidth
        ctrl.Height = ctrlHeightPx
    End If
    currentY = currentY + ctrlHeightPx + spacingPx
End Sub

