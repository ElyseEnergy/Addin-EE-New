VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTicketInput 
   Caption         =   "UserForm1"
   ClientHeight    =   6312
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5064
   OleObjectBlob   =   "frmTicketInput.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTicketInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Submitted As Boolean
Private m_TicketData As ticketData

' Public property to get the result
Public Property Get Submitted() As Boolean
    Submitted = m_Submitted
End Property

Public Property Get TicketDetails() As ticketData
    TicketDetails = m_TicketData
End Property

Public Sub ShowForm(initialData As ticketData, Optional formTitle As String = "Support Ticket")
    m_Submitted = False ' Default to not submitted
    m_TicketData = initialData

    Me.caption = formTitle
    Me.lblFormTitle.caption = formTitle ' Assuming a label for the title exists

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

    Me.cmdSubmit.caption = "Submit Ticket"
    Me.cmdSubmit.Default = True
    Me.cmdCancel.caption = "Cancel"
    Me.cmdCancel.Cancel = True
    
    ' Set a default size, AdjustLayout will refine it
    Me.width = Application.PointsToPixels(450, 0) ' Default width in points
    Me.height = Application.PointsToPixels(550, 1) ' Default height in points
End Sub

Private Sub PopulateFormFields(data As ticketData)
    Me.txtSubject.text = data.subject
    Me.txtDescription.text = data.description
    Me.chkIncludeLogs.value = data.IncludeLogs
    Me.chkIncludeScreenshot.value = data.IncludeScreenshot
    
    ' Set ComboBoxes - requires helper functions in SYS_TicketSystem or here
    SetComboBoxValue Me.cmbPriority, data.Priority
    SetComboBoxValue Me.cmbCategory, data.category
End Sub

Private Sub PopulateComboBoxes()
    Dim priorities As Variant
    Dim Categories As Variant
    Dim i As Long

    ' These functions need to be available (e.g., from SYS_TicketSystem or defined locally)
    On Error Resume Next ' In case functions are not yet available
    priorities = SYS_TicketSystem.GetPriorityEnumArray() ' Assumes this returns a 0-based array of strings
    Categories = SYS_TicketSystem.GetCategoryEnumArray() ' Assumes this returns a 0-based array of strings
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

    If IsArray(Categories) Then
        Me.cmbCategory.Clear
        For i = LBound(Categories) To UBound(Categories)
            Me.cmbCategory.AddItem Categories(i)
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
        m_TicketData.subject = Me.txtSubject.text
        m_TicketData.description = Me.txtDescription.text
        m_TicketData.Priority = Me.cmbPriority.text
        m_TicketData.category = Me.cmbCategory.text
        m_TicketData.IncludeLogs = Me.chkIncludeLogs.value
        m_TicketData.IncludeScreenshot = Me.chkIncludeScreenshot.value
        ' m_TicketData.UserEmail = ... ' If an email field is added
        
        m_Submitted = True
        Me.Hide
    End If
End Sub

Private Function ValidateInputs() As Boolean
    ValidateInputs = True ' Default to true
    If Trim(Me.txtSubject.text) = "" Then
        SYS_MessageBox.ShowErrorMessage "Input Required", "Ticket subject cannot be empty."
        Me.txtSubject.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    If Trim(Me.txtDescription.text) = "" Then
        SYS_MessageBox.ShowWarningMessage "Input Recommended", "Ticket description is empty. Are you sure you want to continue?", vbYesNo
        If SYS_MessageBox.LastButtonPressed <> vbYes Then
             Me.txtDescription.SetFocus
             ValidateInputs = False
             Exit Function
        End If
    End If
    If Me.cmbPriority.ListIndex = -1 And Me.cmbPriority.text <> "" Then
         ' Allow if text is typed but not in list, though ideally it should be from list
    ElseIf Me.cmbPriority.ListIndex = -1 Then
        SYS_MessageBox.ShowErrorMessage "Input Required", "Please select a ticket priority."
        Me.cmbPriority.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    If Me.cmbCategory.ListIndex = -1 And Me.cmbCategory.text <> "" Then
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

    Dim ctrl As MSForms.control
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
    Const LABEL_WIDTH As Long = 80 ' Points
    Const CONTROL_HEIGHT As Long = 20 ' Points
    Const DESCRIPTION_MIN_HEIGHT As Long = 100 ' Points
    
    Dim currentY As Long
    Dim marginPx As Long, spacingPx As Long
    Dim labelWidthPx As Long, controlHeightPx As Long
    Dim availableWidth As Long
    
    ' Convert measurements to pixels
    marginPx = Application.PointsToPixels(10, 1) ' 10 points margin
    spacingPx = Application.PointsToPixels(6, 1) ' 6 points spacing
    labelWidthPx = Application.PointsToPixels(LABEL_WIDTH, 0)
    controlHeightPx = Application.PointsToPixels(CONTROL_HEIGHT, 1)
    
    ' Start position
    currentY = marginPx
    
    ' Calculate available width for controls
    availableWidth = Me.InsideWidth - (2 * marginPx) - labelWidthPx - spacingPx
    
    ' Position title if it exists
    If Not Me.lblFormTitle Is Nothing Then
        With Me.lblFormTitle
            .Left = marginPx
            .Top = currentY
            .width = Me.InsideWidth - (2 * marginPx)
            currentY = currentY + .height + spacingPx
        End With
    End If
    
    ' Position Subject field
    With Me.lblSubject
        .Left = marginPx
        .Top = currentY
        .width = labelWidthPx
    End With
    
    With Me.txtSubject
        .Left = marginPx + labelWidthPx + spacingPx
        .Top = currentY
        .width = availableWidth
        .height = controlHeightPx
    End With
    
    currentY = currentY + controlHeightPx + spacingPx
    
    ' Position Priority field
    With Me.lblPriority
        .Left = marginPx
        .Top = currentY
        .width = labelWidthPx
    End With
    
    With Me.cmbPriority
        .Left = marginPx + labelWidthPx + spacingPx
        .Top = currentY
        .width = availableWidth / 2 - spacingPx
        .height = controlHeightPx
    End With
    
    ' Position Category field (on same line as Priority)
    With Me.lblCategory
        .Left = marginPx + labelWidthPx + spacingPx + (availableWidth / 2)
        .Top = currentY
        .width = labelWidthPx
    End With
    
    With Me.cmbCategory
        .Left = .Left + labelWidthPx + spacingPx
        .Top = currentY
        .width = (availableWidth / 2) - labelWidthPx - spacingPx
        .height = controlHeightPx
    End With
    
    currentY = currentY + controlHeightPx + spacingPx
    
    ' Position Description field
    With Me.lblDescription
        .Left = marginPx
        .Top = currentY
        .width = labelWidthPx
    End With
    
    With Me.txtDescription
        .Left = marginPx + labelWidthPx + spacingPx
        .Top = currentY
        .width = availableWidth
        .height = Application.PointsToPixels(DESCRIPTION_MIN_HEIGHT, 1)
    End With
    
    currentY = currentY + .height + spacingPx
    
    ' Position Checkboxes
    With Me.chkIncludeLogs
        .Left = marginPx + labelWidthPx + spacingPx
        .Top = currentY
        .width = availableWidth / 2
    End With
    
    With Me.chkIncludeScreenshot
        .Left = .Left + .width + spacingPx
        .Top = currentY
        .width = availableWidth / 2
    End With
    
    currentY = currentY + controlHeightPx + spacingPx
    
    ' Position Buttons at bottom
    Dim btnWidth As Long, btnHeight As Long, btnSpacing As Long
    btnWidth = Application.PointsToPixels(80, 0)
    btnHeight = Application.PointsToPixels(25, 1)
    btnSpacing = Application.PointsToPixels(10, 0)
    
    ' Calculate button positions to center them
    Dim totalButtonWidth As Long
    totalButtonWidth = (2 * btnWidth) + btnSpacing
    Dim buttonStartX As Long
    buttonStartX = (Me.InsideWidth - totalButtonWidth) / 2
    
    With Me.cmdSubmit
        .Top = currentY
        .Left = buttonStartX
        .width = btnWidth
        .height = btnHeight
    End With
    
    With Me.cmdCancel
        .Top = currentY
        .Left = buttonStartX + btnWidth + btnSpacing
        .width = btnWidth
        .height = btnHeight
    End With
    
    ' Set final form height
    Me.height = currentY + btnHeight + marginPx + (Me.height - Me.InsideHeight)
    
    ' Ensure the form doesn't exceed screen bounds
    Dim screenHeightPx As Long
    screenHeightPx = Application.PointsToPixels(Application.height, 1)
    If Me.height > screenHeightPx * 0.9 Then
        Me.height = screenHeightPx * 0.9
        ' Adjust description box height to fit
        Me.txtDescription.height = Me.txtDescription.height - _
            (Me.height - (screenHeightPx * 0.9))
    End If
End Sub

Private Sub PositionControlPair(lbl As MSForms.label, ctrl As Object, caption As String, ByRef currentY As Single, ByVal marginPx As Long, ByVal labelWidth As Long, ByVal controlX As Long, ByVal controlWidth As Long, ByVal lblHeightPx As Long, ByVal ctrlHeightPx As Long, ByVal spacingPx As Long)
    If Not lbl Is Nothing Then
        lbl.caption = caption
        lbl.Left = marginPx
        lbl.Top = currentY + (ctrlHeightPx - lblHeightPx) / 2 ' Vertically align label with center of control
        lbl.width = labelWidth
        lbl.height = lblHeightPx
        lbl.TextAlign = fmTextAlignRight ' Align label text to the right for neatness
    End If
    
    If Not ctrl Is Nothing Then
        ctrl.Left = controlX
        ctrl.Top = currentY
        ctrl.width = controlWidth
        ctrl.height = ctrlHeightPx
    End If
    currentY = currentY + ctrlHeightPx + spacingPx
End Sub



