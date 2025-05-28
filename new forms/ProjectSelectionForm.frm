VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProjectSelectionForm 
   Caption         =   "Select Projects - Elyse Energy"
   ClientHeight    =   14265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11355
   OleObjectBlob   =   "ProjectSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProjectSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ========================================
' ORIGINAL PROJECT SELECTION FORM - NO LOGO
' ========================================
' Back to the original code without any logo functionality
' ========================================

Option Explicit

Private selectedProjects As Collection
Private projectData() As String
Private projectCount As Integer
Private WithEvents btnContinue As MSForms.CommandButton
Attribute btnContinue.VB_VarHelpID = -1
Private WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1

Private Sub UserForm_Initialize()
    Set selectedProjects = New Collection
    
    ' Set basic form properties
    Me.Caption = "Select Projects - Elyse Energy"
    Me.BackColor = RGB(248, 249, 250)
    
    ' Load projects from Power Query
    Call LoadProjectsData
    
    ' Create UI elements
    Call CreateProjectUI
End Sub

Private Sub LoadProjectsData()
    Dim tempSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Check if query exists
    If Not QueryExists("00 - List Projet") Then
        MsgBox "The query '00 - List Projet' was not found.", vbExclamation, "Elyse Energy"
        Exit Sub
    End If
    
    ' Create temporary sheet
    Set tempSheet = CreateTempSheet()
    
    On Error GoTo Cleanup
    
    ' Load query data
    LoadQueryToSheet "00 - List Projet", tempSheet
    
    ' Get data range
    lastRow = tempSheet.Cells(tempSheet.Rows.Count, 1).End(xlUp).row
    projectCount = lastRow - 1 ' Skip header
    
    If projectCount <= 0 Then
        MsgBox "No projects found in the query.", vbExclamation, "Elyse Energy"
        GoTo Cleanup
    End If
    
    ' Collect project names
    ReDim projectData(1 To projectCount)
    For i = 2 To lastRow
        projectData(i - 1) = CStr(tempSheet.Cells(i, 1).value)
    Next i
    
Cleanup:
    DeleteTempSheet tempSheet
    On Error GoTo 0
End Sub

Private Sub CreateProjectUI()
    Dim headerLabel As control
    Dim subtitleLabel As control
    Dim scrollFrame As control
    Dim projectCheck As control
    Dim i As Integer
    Dim yPos As Integer
    
    ' Header Label
    Set headerLabel = Me.Controls.Add("Forms.Label.1", "lblHeader")
    With headerLabel
        .Left = 20
        .Top = 20
        .Width = 440
        .Height = 30
        .Caption = "Select Projects"
        .ForeColor = RGB(17, 36, 148)  ' Blue
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    ' Subtitle Label
    Set subtitleLabel = Me.Controls.Add("Forms.Label.1", "lblSubtitle")
    With subtitleLabel
        .Left = 20
        .Top = 55
        .Width = 440
        .Height = 40
        .Caption = "Choose one project to include in your planning analysis"
        .ForeColor = RGB(67, 67, 67)
        .Font.Size = 9
    End With
    
    ' Scrollable Frame for Projects
    Set scrollFrame = Me.Controls.Add("Forms.Frame.1", "frameScroll")
    With scrollFrame
        .Left = 20
        .Top = 100
        .Width = 440
        .Height = 350
        .BackColor = RGB(255, 255, 255)
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = projectCount * 25 + 20
    End With
    
    ' Create project radio buttons (option buttons for single selection)
    yPos = 10
    For i = 1 To projectCount
        ' Option Button (Radio Button)
        Set projectCheck = scrollFrame.Controls.Add("Forms.OptionButton.1", "optProject" & i)
        With projectCheck
            .Left = 15
            .Top = yPos
            .Width = 400
            .Height = 20
            .Caption = projectData(i)
            .Font.Size = 9
            .Tag = projectData(i)  ' Store project name in tag
            .GroupName = "ProjectGroup"  ' All radio buttons in same group
        End With
        
        yPos = yPos + 25
    Next i
    
    ' Create Continue Button with WithEvents
    Set btnContinue = Me.Controls.Add("Forms.CommandButton.1", "btnContinue")
    With btnContinue
        .Left = 330
        .Top = 470
        .Width = 90
        .Height = 35
        .Caption = "Continue"
        .BackColor = RGB(50, 231, 185)  ' Teal
        .Font.Bold = True
    End With
    
    ' Create Cancel Button with WithEvents
    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btnCancel
        .Left = 430
        .Top = 470
        .Width = 80
        .Height = 35
        .Caption = "Cancel"
        .BackColor = RGB(255, 255, 255)
    End With
End Sub

' ========================================
' BUTTON EVENT HANDLERS
' ========================================

Private Sub btnContinue_Click()
    Dim i As Integer
    Dim optButton As control
    Dim selectedProject As String
    
    ' Clear previous selections
    Set selectedProjects = New Collection
    selectedProject = ""
    
    ' Check which project is selected (radio button)
    For i = 1 To projectCount
        On Error Resume Next
        Set optButton = Me.Controls("frameScroll").Controls("optProject" & i)
        If Err.Number = 0 Then
            If optButton.value = True Then
                selectedProject = optButton.Tag
                selectedProjects.Add optButton.Tag
                Exit For
            End If
        End If
        On Error GoTo 0
    Next i
    
    If selectedProject = "" Then
        MsgBox "Please select one project.", vbExclamation, "Elyse Energy"
        Exit Sub
    End If
    
    ' Update the project parameter immediately
    If Not UpdateProjectParameter(selectedProject) Then
        MsgBox "Failed to update project parameter for: " & selectedProject, vbExclamation, "Elyse Energy"
        Exit Sub
    End If
    
    ' Add to global collection for later reference
    If Not g_SelectedProjects Is Nothing Then
        ' Clear and add the selected project
        Set g_SelectedProjects = New Collection
        g_SelectedProjects.Add selectedProject
    End If
    
    ' Hide this form and show planning selection
    Me.Hide
    Call ShowPlanningSelectionForm
End Sub

Private Sub btnCancel_Click()
    MsgBox "Operation cancelled.", vbInformation, "Elyse Energy"
    Unload Me
End Sub

' Alternative event handlers in case the names are different
Private Sub Continue_Click()
    Call btnContinue_Click
End Sub

Private Sub Cancel_Click()
    Call btnCancel_Click
End Sub

' Public function to get selected projects
Public Function GetSelectedProjects() As Collection
    Set GetSelectedProjects = selectedProjects
End Function
