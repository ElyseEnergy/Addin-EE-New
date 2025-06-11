VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CategorySelectionForm 
   Caption         =   "Select Items"
   ClientHeight    =   1212
   ClientLeft      =   120
   ClientTop       =   444
   ClientWidth     =   1836
   OleObjectBlob   =   "CategorySelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CategorySelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_selectedValues As Collection
Private m_availableItems As Variant
Private m_userResult As String
Private m_categoryName As String

Private lblHeader As MSForms.label
Private fraContainer As MSForms.Frame
Private checkBoxes As Collection

' UPDATED: Changed OK button to Next button
Private WithEvents btnNext As MSForms.CommandButton
Attribute btnNext.VB_VarHelpID = -1
Private WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1
Private WithEvents btnSelectAll As MSForms.CommandButton
Attribute btnSelectAll.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    Me.caption = "Select Items - Elyse Energy"
    Me.Width = 600
    Me.Height = 550
    Me.BackColor = RGB(248, 249, 250)
    Me.StartUpPosition = 1
    
    m_userResult = "Cancel"
    Set m_selectedValues = New Collection
    Set checkBoxes = New Collection
    
    CreateControls
End Sub

Private Sub CreateControls()
    ' Header Label
    Set lblHeader = Me.Controls.Add("Forms.Label.1", "lblHeader")
    With lblHeader
        .Left = 25
        .Top = 20
        .Width = 540
        .Height = 40
        .caption = "Select Items"
        .Font.Size = 18
        .Font.Bold = True
        .ForeColor = RGB(17, 36, 148)
        .BackStyle = fmBackStyleTransparent
    End With
    
    ' Container Frame
    Set fraContainer = Me.Controls.Add("Forms.Frame.1", "fraContainer")
    With fraContainer
        .Left = 25
        .Top = 75
        .Width = 540
        .Height = 320
        .caption = ""
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
        .ScrollBars = fmScrollBarsVertical
    End With
    
    ' Select All Button
    Set btnSelectAll = Me.Controls.Add("Forms.CommandButton.1", "btnSelectAll")
    With btnSelectAll
        .Left = 25
        .Top = 410
        .Width = 120
        .Height = 35
        .caption = "Select All"
        .Font.Size = 11
        .Font.Bold = True
        .BackColor = RGB(70, 130, 180)
        .ForeColor = RGB(255, 255, 255)
        .TakeFocusOnClick = False
    End With
    
    ' UPDATED: Next Button (was OK button)
    Set btnNext = Me.Controls.Add("Forms.CommandButton.1", "btnNext")
    With btnNext
        .Left = 380
        .Top = 460
        .Width = 90
        .Height = 40
        .caption = "Next"
        .Default = True
        .Font.Bold = True
        .Font.Size = 12
        .BackColor = RGB(50, 231, 185)
        .ForeColor = RGB(255, 255, 255)
        .TakeFocusOnClick = False
    End With
    
    ' Cancel Button
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btnCancel
        .Left = 480
        .Top = 460
        .Width = 90
        .Height = 40
        .caption = "Cancel"
        .Cancel = True
        .Font.Size = 12
        .BackColor = RGB(220, 220, 220)
        .ForeColor = RGB(80, 80, 80)
        .TakeFocusOnClick = False
    End With
End Sub

Public Sub SetupForm(categoryName As String, items As Variant)
    m_categoryName = categoryName
    lblHeader.caption = "Select " & categoryName & " Items"
    m_availableItems = items
    
    ClearCheckboxes
    
    Dim i As Long
    Dim topPosition As Long
    topPosition = 20
    
    For i = LBound(items) To UBound(items)
        Dim chkBox As MSForms.CheckBox
        Set chkBox = fraContainer.Controls.Add("Forms.CheckBox.1", "chk" & i)
        With chkBox
            .Left = 25
            .Top = topPosition
            .Width = 480
            .Height = 28
            .caption = CStr(items(i))
            .Font.Size = 11
            .Value = False
        End With
        checkBoxes.Add chkBox
        topPosition = topPosition + 35
    Next i
    
    If topPosition > fraContainer.Height Then
        fraContainer.ScrollHeight = topPosition + 30
    End If
    
    If checkBoxes.count > 0 Then
        Dim firstBox As MSForms.CheckBox
        Set firstBox = checkBoxes(1)
        firstBox.Value = True
    End If
End Sub

Private Sub ClearCheckboxes()
    Dim ctrl As control
    For Each ctrl In fraContainer.Controls
        fraContainer.Controls.Remove ctrl.Name
    Next ctrl
    Set checkBoxes = New Collection
End Sub

Private Sub btnSelectAll_Click()
    On Error GoTo ErrorHandler
    
    If checkBoxes.count = 0 Then Exit Sub
    
    Dim allSelected As Boolean
    allSelected = True
    
    Dim i As Long
    Dim chkBox As MSForms.CheckBox
    
    For i = 1 To checkBoxes.count
        Set chkBox = checkBoxes(i)
        If chkBox.Value = False Then
            allSelected = False
            Exit For
        End If
    Next i
    
    For i = 1 To checkBoxes.count
        Set chkBox = checkBoxes(i)
        chkBox.Value = Not allSelected
    Next i
    
    If allSelected Then
        btnSelectAll.caption = "Select All"
    Else
        btnSelectAll.caption = "Deselect All"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in Select All function: " & Err.Description, vbExclamation
End Sub

' UPDATED: Next button handler (was btnOK_Click)
Private Sub btnNext_Click()
    Set m_selectedValues = New Collection
    
    Dim i As Long
    For i = 1 To checkBoxes.count
        Dim chkBox As MSForms.CheckBox
        Set chkBox = checkBoxes(i)
        If chkBox.Value = True Then
            m_selectedValues.Add chkBox.caption
        End If
    Next i
    
    If m_selectedValues.count = 0 Then
        MsgBox "Please select at least one item.", vbExclamation
        Exit Sub
    End If
    
    m_userResult = "Next"
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    m_userResult = "Cancel"
    Me.Hide
End Sub

Public Property Get GetSelectedValues() As Collection
    Set GetSelectedValues = m_selectedValues
End Property

Public Property Get WasCancelled() As Boolean
    WasCancelled = (m_userResult = "Cancel")
End Property

' ADDED: Property to check if Next was clicked
Public Property Get WasNext() As Boolean
    WasNext = (m_userResult = "Next")
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        m_userResult = "Cancel"
        Me.Hide
    End If
End Sub

Private Sub UserForm_Terminate()
    Set m_selectedValues = Nothing
    Set checkBoxes = Nothing
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then ' Enter key
        btnNext_Click
    ElseIf KeyCode = 27 Then ' Escape key
        btnCancel_Click
    End If
End Sub
