VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListSelection 
   Caption         =   "UserForm1"
   ClientHeight    =   4932
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4848
   OleObjectBlob   =   "frmListSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmListSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Public properties to retrieve selection
Private m_Cancelled As Boolean
Private m_SelectedIndices As Collection ' Stores 0-based indices of selected items
Private m_AllowMultiSelect As Boolean

' --- Properties ---
Public Property Get WasCancelled() As Boolean
    WasCancelled = m_Cancelled
End Property

' Returns the 0-based index of the selected item in single-select mode
' Returns -1 if no selection or in multi-select mode with not exactly one selection
Public Property Get SelectedItemIndex() As Long
    If m_AllowMultiSelect Then
        If m_SelectedIndices.count = 1 Then
            SelectedItemIndex = m_SelectedIndices(1) ' Return the single selected index
        Else
            SelectedItemIndex = -1 ' Not a single selection
        End If
    Else ' Single-select mode
        If m_SelectedIndices.count = 1 Then
            SelectedItemIndex = m_SelectedIndices(1)
        Else
            SelectedItemIndex = -1 ' No selection
        End If
    End If
End Property

' Returns a collection of 0-based indices of all selected items (primarily for multi-select mode)
Public Property Get SelectedItemIndices() As Collection
    Set SelectedItemIndices = m_SelectedIndices
End Property

' --- Constants for Layout (add these at the top of the module) ---
Private Const MIN_FORM_WIDTH_LST As Long = 250       ' Min width for list selection form
Private Const DEFAULT_FORM_WIDTH_LST As Long = 350   ' Default width for list selection form
Private Const MIN_FORM_HEIGHT_LST As Long = 200      ' Min height for list selection form
Private Const MAX_FORM_WIDTH_FACTOR_LST As Single = 0.7 ' Max width as factor of screen width
Private Const MAX_FORM_HEIGHT_FACTOR_LST As Single = 0.8 ' Max height as factor of screen height

Private Const BUTTON_WIDTH_LST As Long = 80
Private Const BUTTON_HEIGHT_LST As Long = 25
Private Const BUTTON_SPACING_LST As Long = 8
Private Const VERTICAL_SPACING_LST As Long = 10
Private Const MARGIN_LST As Long = 10
Private Const PROMPT_HEIGHT_LST As Long = 30 ' Estimated height for a couple of lines for the prompt

' --- Initialization for the UserForm ---
Private Sub UserForm_Initialize()
    Set m_SelectedIndices = New Collection
    m_Cancelled = True ' Default to cancelled state
    
    Me.StartUpPosition = 0 ' Center on screen
    
    ' Set initial size - AdjustLayout will refine this
    Me.width = Application.PointsToPixels(DEFAULT_FORM_WIDTH_LST, 0)
    Me.height = Application.PointsToPixels(MIN_FORM_HEIGHT_LST, 1)
End Sub

' --- Public method to configure and set up the form before showing ---
' Parameters:
'   formTitle: Caption for the UserForm window.
'   promptMessage: Text to display as a prompt or instruction.
'                  (This assumes you might add a Label control, e.g., named lblPrompt, to display this.
'                   If lblPrompt doesn't exist, this part will be skipped without error.)
'   items: A Collection of strings to populate the ListBox.
'   allowMulti: Boolean, True for multi-selection, False for single-selection.
'   defaultItemValues: Optional. For single-select, the string value of the item to select by default.
'                        For multi-select, can be an array of string values to select by default.
'   okButtonCaption: Optional. Caption for the OK button (CommandButton1).
'   cancelButtonCaption: Optional. Caption for the Cancel button (CommandButton2).
Public Sub SetupListForm(ByVal formTitle As String, _
                         ByVal promptMessage As String, _
                         ByVal items As Collection, _
                         Optional ByVal allowMulti As Boolean = False, _
                         Optional ByVal defaultItemValues As Variant, _
                         Optional ByVal okButtonCaption As String = "OK", _
                         Optional ByVal cancelButtonCaption As String = "Annuler")

    Dim i As Long
    Dim currentItemValue As Variant

    Me.caption = formTitle
    m_AllowMultiSelect = allowMulti

    ' Attempt to set prompt message if a Label named "lblPrompt" exists
    On Error Resume Next ' Ignore error if lblPrompt doesn't exist
    Me.Controls("lblPrompt").caption = promptMessage
    On Error GoTo 0     ' Restore error handling

    ' Configure ListBox1
    With Me.ListBox1
        .Clear
        For Each currentItemValue In items
            .AddItem currentItemValue
        Next currentItemValue

        If m_AllowMultiSelect Then
            .MultiSelect = fmMultiSelectMulti ' Standard multi-select (checkboxes)
        Else
            .MultiSelect = fmMultiSelectSingle
        End If

        ' Handle default selection(s)
        If Not IsMissing(defaultItemValues) Then
            If m_AllowMultiSelect Then
                If IsArray(defaultItemValues) Then
                    For Each currentItemValue In defaultItemValues ' Iterate through array of default values
                        For i = 0 To .ListCount - 1
                            If CStr(.List(i)) = CStr(currentItemValue) Then
                                .Selected(i) = True
                                Exit For ' Move to next default value
                            End If
                        Next i
                    Next currentItemValue
                Else ' Single default value provided for a multi-select listbox
                    For i = 0 To .ListCount - 1
                        If CStr(.List(i)) = CStr(defaultItemValues) Then
                            .Selected(i) = True
                            Exit For
                        End If
                    Next i
                End If
            Else ' Single-select mode
                For i = 0 To .ListCount - 1
                    If CStr(.List(i)) = CStr(defaultItemValues) Then
                        .ListIndex = i ' Selects the item
                        Exit For
                    End If
                Next i
                ' If default not found in single-select, and list is not empty, select the first item
                If .ListIndex = -1 And .ListCount > 0 Then .ListIndex = 0
            End If
        Else
            ' If no default specified for single-select, and list is not empty, select the first item
            If .ListCount > 0 And Not m_AllowMultiSelect Then .ListIndex = 0
        End If
    End With

    ' Configure Buttons
    Me.CommandButton1.caption = okButtonCaption
    Me.CommandButton2.caption = cancelButtonCaption
    Me.CommandButton1.Default = True ' Make OK button default (responds to Enter)
    Me.CommandButton2.Cancel = True  ' Make Cancel button respond to Esc key

    AdjustLayout ' Call AdjustLayout after setting up controls and potentially prompt
End Sub

' --- Dynamic Layout Adjustment ---
Private Sub AdjustLayout()
    Dim currentY As Long
    Dim promptHeightPx As Long
    Dim listHeightPx As Long
    Dim buttonAreaHeightPx As Long
    Dim availableWidthForControls As Long
    
    Dim marginPx As Long, spacingPx As Long
    Dim btnHeightPx As Long, btnWidthPx As Long

    ' Convert measurements to pixels
    marginPx = Application.PointsToPixels(MARGIN_LST, 1)
    spacingPx = Application.PointsToPixels(VERTICAL_SPACING_LST, 1)
    btnHeightPx = Application.PointsToPixels(BUTTON_HEIGHT_LST, 1)
    btnWidthPx = Application.PointsToPixels(BUTTON_WIDTH_LST, 0)

    ' Start at top margin
    currentY = marginPx

    ' Calculate available width for controls
    availableWidthForControls = Me.InsideWidth - (2 * marginPx)
    If availableWidthForControls < Application.PointsToPixels(MIN_FORM_WIDTH_LST - 2 * MARGIN_LST, 0) Then
        availableWidthForControls = Application.PointsToPixels(MIN_FORM_WIDTH_LST - 2 * MARGIN_LST, 0)
    End If

    ' Position and size prompt label if it exists
    On Error Resume Next
    If Not Me.Controls("lblPrompt") Is Nothing Then
        With Me.Controls("lblPrompt")
            .Top = currentY
            .Left = marginPx
            .width = availableWidthForControls
            .AutoSize = True
            promptHeightPx = .height
            .AutoSize = False
            currentY = currentY + promptHeightPx + spacingPx
        End With
    End If
    On Error GoTo 0

    ' Position and size ListBox
    With Me.ListBox1
        .Top = currentY
        .Left = marginPx
        .width = availableWidthForControls
    End With

    ' Calculate and position buttons at bottom
    Dim totalButtonWidth As Long
    totalButtonWidth = (2 * btnWidthPx) + Application.PointsToPixels(BUTTON_SPACING_LST, 0)
    
    ' Position buttons
    Dim buttonStartX As Long
    buttonStartX = (Me.InsideWidth - totalButtonWidth) / 2

    Me.CommandButton1.Move buttonStartX, _
                          Me.InsideHeight - btnHeightPx - marginPx, _
                          btnWidthPx, _
                          btnHeightPx

    Me.CommandButton2.Move buttonStartX + btnWidthPx + Application.PointsToPixels(BUTTON_SPACING_LST, 0), _
                          Me.InsideHeight - btnHeightPx - marginPx, _
                          btnWidthPx, _
                          btnHeightPx

    ' Calculate and set ListBox height
    listHeightPx = Me.InsideHeight - currentY - btnHeightPx - (2 * marginPx)
    If listHeightPx < Application.PointsToPixels(100, 1) Then ' Minimum ListBox height
        listHeightPx = Application.PointsToPixels(100, 1)
        Me.height = currentY + listHeightPx + btnHeightPx + (2 * marginPx) + (Me.height - Me.InsideHeight)
    End If
    Me.ListBox1.height = listHeightPx

    ' Ensure the form doesn't exceed screen bounds
    Dim screenHeightPx As Long
    screenHeightPx = Application.PointsToPixels(Application.height, 1)
    If Me.height > screenHeightPx * MAX_FORM_HEIGHT_FACTOR_LST Then
        Me.height = screenHeightPx * MAX_FORM_HEIGHT_FACTOR_LST
        ' Readjust ListBox height
        Me.ListBox1.height = Me.InsideHeight - currentY - btnHeightPx - (2 * marginPx)
    End If
End Sub

' --- Method to apply corporate styling (called from SYS_MessageBox) ---
Public Sub StyleForm(colors As Object)
    On Error Resume Next ' Ignore errors if a control doesn't exist or property not applicable

    Me.BackColor = colors("background")

    ' Prompt Label (lblPrompt)
    Dim lblPromptExists As Boolean
    lblPromptExists = False
    If Not Me.Controls("lblPrompt") Is Nothing Then lblPromptExists = True
    
    If lblPromptExists Then
        With Me.Controls("lblPrompt")
            .ForeColor = colors("text")
            .BackColor = colors("background") ' Ensure transparency or match
        End With
    End If

    ' ListBox (ListBox1)
    With Me.ListBox1
        .BackColor = colors("input_bg")
        .ForeColor = colors("input_text")
        .BorderColor = colors("input_border")
        ' .SpecialEffect = fmSpecialEffectFlat ' Or another effect
    End With

    ' OK Button (CommandButton1 - Primary)
    With Me.CommandButton1
        .BackColor = colors("primary")
        .ForeColor = colors("button_text")
        ' .BorderColor = colors("button_border")
        ' .SpecialEffect = fmSpecialEffectFlat
    End With

    ' Cancel Button (CommandButton2 - Secondary)
    With Me.CommandButton2
        If Not IsEmpty(colors("secondary_button_bg")) Then
             .BackColor = colors("secondary_button_bg")
        Else
             .BackColor = RGB(220, 220, 220) ' Light gray fallback
        End If
        If Not IsEmpty(colors("secondary_button_text")) Then
            .ForeColor = colors("secondary_button_text")
        Else
            .ForeColor = colors("text") ' Default text color fallback
        End If
        ' .BorderColor = colors("button_border")
        ' .SpecialEffect = fmSpecialEffectFlat
    End With
    
    On Error GoTo 0
End Sub

' --- Event Handlers for UserForm Controls ---

' OK Button (CommandButton1)
Private Sub CommandButton1_Click()
    Dim i As Long
    m_Cancelled = False
    Set m_SelectedIndices = New Collection ' Reset collection

    If m_AllowMultiSelect Then
        For i = 0 To Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) Then
                m_SelectedIndices.Add i ' Add 0-based index
            End If
        Next i
    Else ' Single-select mode
        If Me.ListBox1.ListIndex <> -1 Then ' Check if an item is selected
            m_SelectedIndices.Add Me.ListBox1.ListIndex ' Add 0-based index
        End If
    End If
    Me.Hide ' Close the form
End Sub

' Cancel Button (CommandButton2)
Private Sub CommandButton2_Click()
    m_Cancelled = True
    Set m_SelectedIndices = New Collection ' Clear any selections
    Me.Hide ' Close the form
End Sub

' Handles the user clicking the 'X' button to close the form
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        m_Cancelled = True
        Set m_SelectedIndices = New Collection ' Clear any selections
        ' Me.Hide will be called automatically by the system
    End If
End Sub

' Optional: Allow double-click on a ListBox item to act as "OK" in single-select mode
Private Sub ListBox1_DblClick(ByVal CancelLogic As MSForms.ReturnBoolean)
    If Not m_AllowMultiSelect Then
        If Me.ListBox1.ListIndex <> -1 Then ' Ensure an item is actually selected
            Call CommandButton1_Click ' Simulate OK button click
        End If
    End If
End Sub


