Option Explicit

' Public property to store the result (which button was clicked by its 1-based index)
Public ClickedButtonIndex As Integer
Public mConfig As MessageBoxConfig ' Store the configuration

' --- Constants for Layout ---
Private Const MIN_FORM_WIDTH As Long = 220 ' Minimum width of the form
Private Const DEFAULT_FORM_WIDTH As Long = 350 ' Default width if not specified in config
Private Const MIN_FORM_HEIGHT As Long = 120 ' Minimum height of the form
Private Const MAX_FORM_WIDTH_FACTOR As Single = 0.8 ' Max width as factor of screen width
Private Const MAX_FORM_HEIGHT_FACTOR As Single = 0.7 ' Max height as factor of screen height

Private Const BUTTON_WIDTH As Long = 75
Private Const BUTTON_HEIGHT As Long = 23
Private Const BUTTON_SPACING As Long = 6 ' Horizontal space between buttons
Private Const VERTICAL_SPACING As Long = 10 ' Space between elements like message and buttons
Private Const MARGIN As Long = 12         ' General margin for content from form edges
Private Const ICON_TEXT_SPACING As Long = 8 ' Space between icon and text message
Private Const ICON_WIDTH As Long = 32
Private Const ICON_HEIGHT As Long = 32


' --- Initialize the UserForm based on the configuration passed from SYS_MessageBox ---
Public Sub InitializeFromConfig(config As MessageBoxConfig)
    Dim i As Integer
    Dim btn As MSForms.CommandButton
    Dim defaultButtonSet As Boolean
    defaultButtonSet = False

    Set Me.mConfig = config ' Store the config

    ' 1. Set Form Properties
    Me.Caption = config.Title
    
    ' Determine initial form width
    If config.Width > 0 Then
        Me.Width = Application.PointsToPixels(config.Width, 0)
    Else
        Me.Width = Application.PointsToPixels(DEFAULT_FORM_WIDTH, 0)
    End If
    ' Ensure min/max width
    Dim screenWidthPx As Long
    screenWidthPx = Application.PointsToPixels(Application.Width, 0) ' Application.Width is in points
    If Me.Width < Application.PointsToPixels(MIN_FORM_WIDTH, 0) Then Me.Width = Application.PointsToPixels(MIN_FORM_WIDTH, 0)
    If Me.Width > screenWidthPx * MAX_FORM_WIDTH_FACTOR Then Me.Width = screenWidthPx * MAX_FORM_WIDTH_FACTOR


    ' Height will be adjusted by AdjustFormLayout, but set a preliminary min height
    Me.Height = Application.PointsToPixels(MIN_FORM_HEIGHT, 0)
    
    Me.StartUpPosition = 0 ' CenterScreen

    ' 2. Set Message Text (using Label2 for the main message)
    Me.Label2.Caption = config.Message
    Me.Label2.AutoSize = False ' Crucial for WordWrap to work with fixed width
    Me.Label2.WordWrap = True

    ' 3. Configure Icon (Image1 and Label1 as placeholder)
    Me.Label1.Caption = "" ' Clear Label1 initially
    Me.Image1.Width = Application.PointsToPixels(ICON_WIDTH, 0)
    Me.Image1.Height = Application.PointsToPixels(ICON_HEIGHT, 0)
    Me.Label1.Width = Me.Image1.Width
    Me.Label1.Height = Me.Image1.Height
    Me.Label1.WordWrap = True
    Me.Label1.TextAlign = fmTextAlignCenter
    
    If config.ShowIcon Then
        Me.Image1.Visible = True ' Assume visible, will be hidden if loading fails
        
        ' --- START REPLACEMENT: Load icon from "EE_Image" sheet ---
        Dim wsImages As Worksheet
        Dim shpIcon As Shape
        Dim iconShapeName As String
        Dim tempFilePath As String
        Dim tempChart As ChartObject ' Temporary ChartObject to help with export
        Dim bIconLoadedSuccessfully As Boolean
        bIconLoadedSuccessfully = False

        On Error GoTo LoadIcon_Error

        Set wsImages = ThisWorkbook.Sheets("EE_Image")

        Select Case config.MessageType
            Case INFO_MESSAGE: iconShapeName = "IconInfo"
            Case SUCCESS_MESSAGE: iconShapeName = "IconSuccess"
            Case WARNING_MESSAGE: iconShapeName = "IconWarning"
            Case ERROR_MESSAGE: iconShapeName = "IconError"
            Case CONFIRMATION_MESSAGE: iconShapeName = "IconQuestion"
            Case Else: iconShapeName = ""
        End Select

        If iconShapeName <> "" Then
            Set shpIcon = wsImages.Shapes(iconShapeName)
            shpIcon.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
            tempFilePath = Environ("TEMP") & "\\~elyse_icon_" & Format(Now, "yyyymmddhhmmssfff") & ".gif"
            Set tempChart = wsImages.ChartObjects.Add(Left:=shpIcon.Left, Top:=shpIcon.Top, Width:=shpIcon.Width, Height:=shpIcon.Height)
            With tempChart.Chart
                .ChartArea.Format.Line.Visible = msoFalse
                .Paste
                .Export Filename:=tempFilePath, FilterName:="GIF"
            End With
            Me.Image1.Picture = LoadPicture(tempFilePath)
            Me.Image1.PictureSizeMode = fmPictureSizeModeZoom 

            If Not (Me.Image1.Picture Is Nothing) Then
                bIconLoadedSuccessfully = True
            Else
                Debug.Print "frmCustomMessageBox.InitializeFromConfig: LoadPicture failed for " & tempFilePath
            End If
        Else
            Debug.Print "frmCustomMessageBox.InitializeFromConfig: No iconShapeName defined for MessageType: " & config.MessageType
        End Select

LoadIcon_Cleanup:
        On Error Resume Next
        If Not tempChart Is Nothing Then tempChart.Delete
        If tempFilePath <> "" And Dir(tempFilePath) <> "" Then Kill tempFilePath
        On Error GoTo 0
        GoTo LoadIcon_Finalize

LoadIcon_Error:
        Debug.Print "frmCustomMessageBox.InitializeFromConfig Error (LoadIcon): " & Err.Number & " - " & Err.Description & ". Shape name attempted: '" & iconShapeName & "' on sheet 'EE_Image'."
        Err.Clear
        Resume LoadIcon_Cleanup

LoadIcon_Finalize:
        If bIconLoadedSuccessfully Then
            Me.Image1.Visible = True
            Me.Label1.Visible = False
        Else
            Me.Image1.Visible = False
            Me.Label1.Visible = True
            Select Case config.MessageType
                Case INFO_MESSAGE: Me.Label1.Caption = "[INFO]"
                Case SUCCESS_MESSAGE: Me.Label1.Caption = "[OK]"
                Case WARNING_MESSAGE: Me.Label1.Caption = "[WARN]"
                Case ERROR_MESSAGE: Me.Label1.Caption = "[ERR]"
                Case CONFIRMATION_MESSAGE: Me.Label1.Caption = "[?]"
                Case Else: Me.Label1.Visible = False
            End Select
        End If
    Else
        Me.Image1.Visible = False
        Me.Label1.Visible = False
    End If
    
    ' Initial position of Icon/Fallback Label
    Me.Image1.Top = Application.PointsToPixels(MARGIN, 1)
    Me.Image1.Left = Application.PointsToPixels(MARGIN, 0)
    Me.Label1.Top = Me.Image1.Top
    Me.Label1.Left = Me.Image1.Left

    ' 4. Configure Buttons
    Dim buttonControls(1 To 3) As MSForms.CommandButton
    Set buttonControls(1) = Me.CommandButton1
    Set buttonControls(2) = Me.CommandButton2
    Set buttonControls(3) = Me.CommandButton3

    For i = 1 To 3
        Set btn = buttonControls(i)
        If i <= config.ButtonCount Then
            btn.Caption = config.Buttons(i).Text
            btn.Tag = CStr(i)
            btn.Visible = True
            btn.Enabled = True
            btn.Width = Application.PointsToPixels(BUTTON_WIDTH, 0)
            btn.Height = Application.PointsToPixels(BUTTON_HEIGHT, 1)

            If config.Buttons(i).IsDefault And Not defaultButtonSet Then
                btn.Default = True
                defaultButtonSet = True
            Else
                btn.Default = False
            End If

            If LCase(Trim(config.Buttons(i).Text)) = "cancel" Then
                btn.Cancel = True
            Else
                btn.Cancel = False
            End If
        Else
            btn.Visible = False
            btn.Enabled = False
            btn.Default = False
            btn.Cancel = False
        End If
    Next i

    If Not defaultButtonSet And config.ButtonCount > 0 Then
        For i = 1 To config.ButtonCount
            If buttonControls(i).Visible Then
                buttonControls(i).Default = True
                Exit For
            End If
        Next i
    End If
    
    ' 5. Adjust Layout dynamically
    AdjustFormLayout
End Sub

Private Sub AdjustFormLayout()
    Dim currentX As Long, currentY As Long
    Dim messageAreaWidth As Long
    Dim requiredLabelHeight As Long
    Dim totalButtonWidth As Long
    Dim visibleButtonCount As Integer
    Dim btn As MSForms.CommandButton
    Dim i As Integer
    Dim buttonControls(1 To 3) As MSForms.CommandButton
    
    Set buttonControls(1) = Me.CommandButton1
    Set buttonControls(2) = Me.CommandButton2
    Set buttonControls(3) = Me.CommandButton3

    ' --- Adjust Message Label (Label2) ---
    currentY = Application.PointsToPixels(MARGIN, 1) ' Start Y position for content

    ' Calculate message area width based on icon presence
    If Me.Image1.Visible Or Me.Label1.Visible Then
        Me.Label2.Left = Me.Image1.Left + Me.Image1.Width + Application.PointsToPixels(ICON_TEXT_SPACING, 0)
        messageAreaWidth = Me.InsideWidth - Me.Label2.Left - Application.PointsToPixels(MARGIN, 0)
    Else
        Me.Label2.Left = Application.PointsToPixels(MARGIN, 0)
        messageAreaWidth = Me.InsideWidth - Application.PointsToPixels(MARGIN * 2, 0)
    End If
    
    ' Ensure minimum message width
    If messageAreaWidth < Application.PointsToPixels(MIN_FORM_WIDTH / 2, 0) Then
        messageAreaWidth = Application.PointsToPixels(MIN_FORM_WIDTH / 2, 0)
    End If
    Me.Label2.Width = messageAreaWidth

    ' Calculate required height for message
    Me.Label2.AutoSize = True
    requiredLabelHeight = Me.Label2.Height
    Me.Label2.AutoSize = False
    Me.Label2.Height = requiredLabelHeight

    ' Position message vertically
    Me.Label2.Top = currentY
    
    ' Calculate final content height
    Dim contentHeight As Long
    If Me.Image1.Visible Or Me.Label1.Visible Then
        contentHeight = Application.Max(Me.Image1.Height, Me.Label2.Height)
    Else
        contentHeight = Me.Label2.Height
    End If

    ' Update vertical position for buttons
    currentY = currentY + contentHeight + Application.PointsToPixels(VERTICAL_SPACING, 1)

    ' --- Position Buttons ---
    visibleButtonCount = 0
    For i = 1 To 3
        If buttonControls(i).Visible Then visibleButtonCount = visibleButtonCount + 1
    Next i

    ' Calculate total width needed for buttons
    totalButtonWidth = (visibleButtonCount * BUTTON_WIDTH) + ((visibleButtonCount - 1) * BUTTON_SPACING)
    
    ' Start X position for first button (centered)
    currentX = (Me.InsideWidth - totalButtonWidth) / 2

    ' Position each visible button
    Dim visibleIndex As Integer
    visibleIndex = 0
    For i = 1 To 3
        If buttonControls(i).Visible Then
            With buttonControls(i)
                .Top = currentY
                .Left = currentX + (visibleIndex * (BUTTON_WIDTH + BUTTON_SPACING))
                .Width = BUTTON_WIDTH
                .Height = BUTTON_HEIGHT
            End With
            visibleIndex = visibleIndex + 1
        End If
    Next i

    ' Set final form height
    Me.Height = currentY + BUTTON_HEIGHT + MARGIN + (Me.Height - Me.InsideHeight)
End Sub

' --- Event Handlers for Buttons ---
Private Sub CommandButton1_Click()
    If Me.CommandButton1.Tag <> "" Then
        Me.ClickedButtonIndex = CInt(Me.CommandButton1.Tag)
    Else
        Me.ClickedButtonIndex = 1 ' Fallback if tag is empty
    End If
    Me.Hide
End Sub

Private Sub CommandButton2_Click()
    If Me.CommandButton2.Tag <> "" Then
        Me.ClickedButtonIndex = CInt(Me.CommandButton2.Tag)
    Else
        Me.ClickedButtonIndex = 2 ' Fallback
    End If
    Me.Hide
End Sub

Private Sub CommandButton3_Click()
    If Me.CommandButton3.Tag <> "" Then
        Me.ClickedButtonIndex = CInt(Me.CommandButton3.Tag)
    Else
        Me.ClickedButtonIndex = 3 ' Fallback
    End If
    Me.Hide
End Sub

' --- Handle Form Activation (e.g., to set focus) ---
Private Sub UserForm_Activate()
    ' Set focus to the default button when the form activates
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CommandButton" Then
            Dim btn As MSForms.CommandButton
            Set btn = ctrl
            If btn.Visible And btn.Default Then
                On Error Resume Next ' In case focus cannot be set
                btn.SetFocus
                On Error GoTo 0
                Exit For
            End If
        End If
    Next ctrl
End Sub

' --- Handle Form Closing (e.g., user clicks 'X') ---
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' If the form is closed by the user clicking the 'X' button (vbFormControlMenu)
    If CloseMode = vbFormControlMenu Then
        ' We need to determine what ClickedButtonIndex should be.
        ' Typically, this means "Cancel" or the action of a button marked as .Cancel = True
        Dim i As Integer
        Dim btn As MSForms.CommandButton
        Dim cancelActionIndex As Integer
        cancelActionIndex = 0 ' Default to 0 (generic cancel/close)

        ' Check if any button is explicitly a "Cancel" button
        If Me.CommandButton1.Visible And Me.CommandButton1.Cancel Then
            cancelActionIndex = CInt(Me.CommandButton1.Tag)
        ElseIf Me.CommandButton2.Visible And Me.CommandButton2.Cancel Then
            cancelActionIndex = CInt(Me.CommandButton2.Tag)
        ElseIf Me.CommandButton3.Visible And Me.CommandButton3.Cancel Then
            cancelActionIndex = CInt(Me.CommandButton3.Tag)
        End If
        
        ' If no button has .Cancel=True, check for a button with "Cancel" text
        If cancelActionIndex = 0 Then
            If Me.CommandButton1.Visible And LCase(Trim(Me.CommandButton1.Caption)) = "cancel" Then
                 cancelActionIndex = CInt(Me.CommandButton1.Tag)
            ElseIf Me.CommandButton2.Visible And LCase(Trim(Me.CommandButton2.Caption)) = "cancel" Then
                 cancelActionIndex = CInt(Me.CommandButton2.Tag)
            ElseIf Me.CommandButton3.Visible And LCase(Trim(Me.CommandButton3.Caption)) = "cancel" Then
                 cancelActionIndex = CInt(Me.CommandButton3.Tag)
            End If
        End If
        ' If still 0, and there's a button count of 2, assume the second is cancel if not explicitly set
        If cancelActionIndex = 0 And mConfig.ButtonCount = 2 Then
             If LCase(mConfig.Buttons(2).Text) = "cancel" Or LCase(mConfig.Buttons(2).Text) = "no" Then
                cancelActionIndex = 2
             End If
        End If
        ' If still 0, and only one button, it's not a "cancel" action by closing via X
        ' If more than one button and no clear cancel, result is ambiguous, 0 is fine.

        Me.ClickedButtonIndex = cancelActionIndex
    End If
End Sub
