Option Explicit

' --- Constants for Layout ---
Private Const MIN_FORM_WIDTH_MD As Long = 300 ' Min width for markdown display form
Private Const DEFAULT_FORM_WIDTH_MD As Long = 500 ' Default width for markdown display form
Private Const MIN_FORM_HEIGHT_MD As Long = 200 ' Min height for markdown display form
Private Const MAX_FORM_WIDTH_FACTOR_MD As Single = 0.8 ' Max width as factor of screen width
Private Const MAX_FORM_HEIGHT_FACTOR_MD As Single = 0.8 ' Max height as factor of screen height

Private Const BUTTON_WIDTH_MD As Long = 80
Private Const BUTTON_HEIGHT_MD As Long = 25
Private Const VERTICAL_SPACING_MD As Long = 10
Private Const MARGIN_MD As Long = 12
Private Const MIN_CONTENT_HEIGHT_MD As Long = 100

' --- Private Variables ---
Private m_OriginalContent As String ' Store the original markdown content

' --- Initialize the UserForm ---
Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0 ' Center on screen
    Me.KeyPreview = True ' Allow form-level keyboard handling
    
    ' Set initial dimensions - will be refined by AdjustLayout
    Me.Width = Application.PointsToPixels(DEFAULT_FORM_WIDTH_MD, 0)
    Me.Height = Application.PointsToPixels(MIN_FORM_HEIGHT_MD, 1)
    
    ' Initialize txtContent
    With Me.txtContent ' Assuming a TextBox named txtContent for markdown display
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .EnterKeyBehavior = True ' Allow Enter for new lines
        .WordWrap = True
        .Locked = True ' Read-only for display
    End With
    
    ' Initialize OK button
    With Me.cmdOK
        .Caption = "OK"
        .Default = True ' Responds to Enter key
        .Cancel = True ' Also responds to Escape key (since it's the only button)
    End With
End Sub

' --- Public method to show content ---
Public Sub ShowContent(title As String, content As String)
    Me.Caption = title
    m_OriginalContent = content ' Store original content
    
    ' Convert markdown to rich text and display
    DisplayMarkdownContent content
    
    AdjustLayout
    Me.Show vbModal
End Sub

' --- Private method to convert and display markdown ---
Private Sub DisplayMarkdownContent(markdownText As String)
    Dim formattedText As String
    formattedText = ConvertMarkdownToDisplayText(markdownText)
    Me.txtContent.Text = formattedText
End Sub

Private Function ConvertMarkdownToDisplayText(markdownText As String) As String
    ' This is a basic markdown converter. For now, it just passes through the text.
    ' TODO: Implement markdown conversion if needed, or consider using HTML/RTF
    ' Some ideas for basic formatting:
    ' - Convert # Headers to UPPERCASE or add spacing
    ' - Convert ** or __ to indicate bold text
    ' - Convert * or _ to indicate italic text
    ' - Convert - or * lists to proper bullet points
    ' - Etc.
    
    ConvertMarkdownToDisplayText = markdownText
End Function

' --- Dynamic Layout Management ---
Private Sub AdjustLayout()
    Dim currentY As Long
    Dim marginPx As Long, spacingPx As Long, btnHeightPx As Long, btnWidthPx As Long
    
    ' Convert points to pixels
    marginPx = Application.PointsToPixels(MARGIN_MD, 0)
    spacingPx = Application.PointsToPixels(VERTICAL_SPACING_MD, 1)
    btnHeightPx = Application.PointsToPixels(BUTTON_HEIGHT_MD, 1)
    btnWidthPx = Application.PointsToPixels(BUTTON_WIDTH_MD, 0)
    
    ' Position Content TextBox
    currentY = marginPx
    Me.txtContent.Left = marginPx
    Me.txtContent.Top = currentY
    Me.txtContent.Width = Me.InsideWidth - (2 * marginPx)
    
    ' Calculate height for content (all available space minus margins and button area)
    Dim contentHeight As Long
    contentHeight = Me.InsideHeight - (2 * marginPx) - spacingPx - btnHeightPx
    If contentHeight < Application.PointsToPixels(MIN_CONTENT_HEIGHT_MD, 1) Then
        contentHeight = Application.PointsToPixels(MIN_CONTENT_HEIGHT_MD, 1)
    End If
    Me.txtContent.Height = contentHeight
    
    ' Position OK button at the bottom
    currentY = Me.txtContent.Top + Me.txtContent.Height + spacingPx
    Me.cmdOK.Top = currentY
    Me.cmdOK.Left = (Me.InsideWidth - btnWidthPx) / 2 ' Center horizontally
    Me.cmdOK.Width = btnWidthPx
    Me.cmdOK.Height = btnHeightPx
    
    ' Calculate final form height
    Me.Height = currentY + btnHeightPx + marginPx + (Me.Height - Me.InsideHeight)
    
    ' Ensure min/max dimensions
    Dim screenWidthPx As Long, screenHeightPx As Long
    screenWidthPx = Application.PointsToPixels(Application.Width, 0)
    screenHeightPx = Application.PointsToPixels(Application.Height, 1)
    
    If Me.Width < Application.PointsToPixels(MIN_FORM_WIDTH_MD, 0) Then Me.Width = Application.PointsToPixels(MIN_FORM_WIDTH_MD, 0)
    If Me.Width > screenWidthPx * MAX_FORM_WIDTH_FACTOR_MD Then Me.Width = screenWidthPx * MAX_FORM_WIDTH_FACTOR_MD
    
    If Me.Height < Application.PointsToPixels(MIN_FORM_HEIGHT_MD, 1) Then Me.Height = Application.PointsToPixels(MIN_FORM_HEIGHT_MD, 1)
    If Me.Height > screenHeightPx * MAX_FORM_HEIGHT_FACTOR_MD Then Me.Height = screenHeightPx * MAX_FORM_HEIGHT_FACTOR_MD
    
    ' If form height was capped, adjust content height
    contentHeight = Me.InsideHeight - Me.cmdOK.Height - marginPx * 2 - spacingPx
    If contentHeight >= Application.PointsToPixels(MIN_CONTENT_HEIGHT_MD, 1) Then
        Me.txtContent.Height = contentHeight
    End If
End Sub

' --- Corporate Styling ---
Public Sub StyleForm(colors As Object)
    On Error Resume Next ' Ignore errors if color not defined
    
    Me.BackColor = colors("background")
    
    With Me.txtContent
        .BackColor = colors("input_bg")
        .ForeColor = colors("input_text")
        .BorderColor = colors("input_border")
        ' .SpecialEffect = fmSpecialEffectFlat ' Optional visual effect
    End With
    
    With Me.cmdOK
        .BackColor = colors("primary")
        .ForeColor = colors("button_text")
        ' .BorderColor = colors("button_border")
        ' .SpecialEffect = fmSpecialEffectFlat
    End With
    
    On Error GoTo 0
End Sub

' --- Event Handlers ---
Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        cmdOK_Click ' Handle 'X' button like OK
    End If
End Sub

Private Sub txtContent_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Handle Ctrl+A to select all text
    If Shift = 2 And KeyCode = vbKeyA Then ' Ctrl+A
        txtContent.SelStart = 0
        txtContent.SelLength = Len(txtContent.Text)
        KeyCode = 0 ' Prevent default handling
    End If
End Sub

Private Sub UserForm_Resize()
    If Me.Width > 0 And Me.Height > 0 Then ' Prevent recursion and 0-size
        AdjustLayout
    End If
End Sub