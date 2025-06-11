VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ModeSelectionForm 
   Caption         =   "UserForm1"
   ClientHeight    =   1212
   ClientLeft      =   336
   ClientTop       =   456
   ClientWidth     =   1836
   OleObjectBlob   =   "ModeSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ModeSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_userResult As String
Private m_modeTransposed As Boolean

Private lblHeader As MSForms.label
Private lblQuestion As MSForms.label
Private fraContainer As MSForms.Frame
Private optNormal As MSForms.OptionButton
Private optTransposed As MSForms.OptionButton

Private WithEvents btnNext As MSForms.CommandButton
Attribute btnNext.VB_VarHelpID = -1
Private WithEvents btnBack As MSForms.CommandButton
Attribute btnBack.VB_VarHelpID = -1
Private WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    Me.caption = "Mode de collage - Elyse Energy"
    Me.Width = 500
    Me.Height = 350
    Me.BackColor = RGB(248, 249, 250)
    Me.StartUpPosition = 1
    
    m_userResult = "Cancel"
    m_modeTransposed = False
    
    CreateControls
End Sub

Private Sub CreateControls()
    ' Header Label
    Set lblHeader = Me.Controls.Add("Forms.Label.1", "lblHeader")
    With lblHeader
        .Left = 25
        .Top = 20
        .Width = 440
        .Height = 40
        .caption = "Mode de collage"
        .Font.Size = 18
        .Font.Bold = True
        .ForeColor = RGB(17, 36, 148)
        .BackStyle = fmBackStyleTransparent
    End With
    
    ' Question Label
    Set lblQuestion = Me.Controls.Add("Forms.Label.1", "lblQuestion")
    With lblQuestion
        .Left = 25
        .Top = 75
        .Width = 440
        .Height = 60
        .caption = "Coller les fiches en mode NORMAL (lignes) ?" & vbCrLf & "ou en mode TRANSPOSE (colonnes) ?"
        .Font.Size = 12
        .ForeColor = RGB(80, 80, 80)
        .BackStyle = fmBackStyleTransparent
        .WordWrap = True
        .TextAlign = fmTextAlignLeft
    End With
    
    ' Container Frame for options
    Set fraContainer = Me.Controls.Add("Forms.Frame.1", "fraContainer")
    With fraContainer
        .Left = 25
        .Top = 150
        .Width = 440
        .Height = 100
        .caption = ""
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle
    End With
    
    ' Normal Mode Option
    Set optNormal = fraContainer.Controls.Add("Forms.OptionButton.1", "optNormal")
    With optNormal
        .Left = 25
        .Top = 20
        .Width = 380
        .Height = 30
        .caption = "Mode NORMAL (lignes) - Données en lignes horizontales"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(50, 150, 50)
        .Value = True  ' Default selection
    End With
    
    ' Transposed Mode Option
    Set optTransposed = fraContainer.Controls.Add("Forms.OptionButton.1", "optTransposed")
    With optTransposed
        .Left = 25
        .Top = 55
        .Width = 380
        .Height = 30
        .caption = "Mode TRANSPOSE (colonnes) - Données en colonnes verticales"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(180, 100, 50)
        .Value = False
    End With
    
    ' Back Button
    Set btnBack = Me.Controls.Add("Forms.CommandButton.1", "btnBack")
    With btnBack
        .Left = 145  ' Centered position
        .Top = 270
        .Width = 90
        .Height = 40
        .caption = "Back"
        .Font.Size = 12
        .Font.Bold = True
        .BackColor = RGB(50, 231, 185)  ' Same color as Next button
        .ForeColor = RGB(255, 255, 255)
        .TakeFocusOnClick = False
    End With
    
    ' Next Button (formerly OK)
    Set btnNext = Me.Controls.Add("Forms.CommandButton.1", "btnNext")
    With btnNext
        .Left = 245  ' Centered position
        .Top = 270
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
        .Left = 380
        .Top = 270
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

Private Sub btnNext_Click()
    ' Get the selected mode
    If optTransposed.Value = True Then
        m_modeTransposed = True
    Else
        m_modeTransposed = False
    End If
    
    m_userResult = "OK"
    Me.Hide
End Sub

Private Sub btnBack_Click()
    m_userResult = "Back"
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    m_userResult = "Cancel"
    Me.Hide
End Sub

Public Property Get WasCancelled() As Boolean
    WasCancelled = (m_userResult = "Cancel")
End Property

Public Property Get WasBack() As Boolean
    WasBack = (m_userResult = "Back")
End Property

Public Property Get isTransposed() As Boolean
    isTransposed = m_modeTransposed
End Property

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        m_userResult = "Cancel"
        Me.Hide
    End If
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then ' Enter key
        btnNext_Click
    ElseIf KeyCode = 27 Then ' Escape key
        btnCancel_Click
    End If
End Sub
