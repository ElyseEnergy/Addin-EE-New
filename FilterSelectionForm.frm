VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FilterSelectionForm 
   Caption         =   "Select Filter - Elyse Energy"
   ClientHeight    =   14265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11355
   OleObjectBlob   =   "FilterSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FilterSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type FilterConfig
    Category As CategoryInfo
    Title As String
    Subtitle As String
End Type

Private filterConfig As FilterConfig
Private selectedItems As Collection
Private itemData() As String
Private itemCount As Integer
Private WithEvents btnContinue As MSForms.CommandButton
Attribute btnContinue.VB_VarHelpID = -1
Private WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1

Public Sub Initialize(ByVal category As CategoryInfo)
    ' Stocker la configuration
    With filterConfig
        Set .Category = category
        .Title = "Select " & category.filterLevel
        .Subtitle = "Choose a " & category.filterLevel & " to filter " & category.displayName
    End With
    
    Set selectedItems = New Collection
    Me.Caption = filterConfig.Title & " - Elyse Energy"
    Me.BackColor = RGB(248, 249, 250)
    
    Call LoadFilterData
    Call CreateFilterUI
End Sub

' Nouvelle initialisation à partir d'une simple liste de valeurs
Public Sub InitializeWithList(values As Collection, prompt As String)
    ' Stocker une configuration minimale
    filterConfig.Title = prompt
    filterConfig.Subtitle = ""
    filterConfig.Category.filterLevel = "value"

    ' Préparer les données à afficher
    Set selectedItems = New Collection
    itemCount = 0
    If Not values Is Nothing Then
        itemCount = values.Count
        ReDim itemData(1 To itemCount)
        Dim i As Long
        For i = 1 To itemCount
            itemData(i) = CStr(values(i))
        Next i
    Else
        ReDim itemData(0)
    End If

    ' Paramétrer l'interface
    Me.Caption = prompt
    Me.BackColor = RGB(248, 249, 250)

    Call CreateFilterUI
End Sub

Private Sub LoadFilterData()
    If filterConfig.Category Is Nothing Then
        MsgBox "Category not initialized.", vbExclamation, "Elyse Energy"
        Exit Sub
    End If
    
    ' Utiliser la catégorie pour obtenir les données
    Dim queryName As String
    queryName = filterConfig.Category.PowerQueryName
    
    If Not QueryExists(queryName) Then
        MsgBox "Query not found: " & queryName, vbExclamation, "Elyse Energy"
        Exit Sub
    End If
    
    ' Charger les données à partir de la query
    Dim filterValues As Collection
    Set filterValues = GetUniqueFilterValues(filterConfig.Category)
    
    If filterValues Is Nothing Or filterValues.Count = 0 Then
        MsgBox "No filter values found.", vbExclamation, "Elyse Energy"
        Exit Sub
    End If
    
    ' Convertir en tableau pour l'affichage
    itemCount = filterValues.Count
    ReDim itemData(1 To itemCount)
    
    Dim i As Integer
    For i = 1 To filterValues.Count
        itemData(i) = filterValues(i)
    Next i
End Sub

Private Sub CreateFilterUI()
    Dim headerLabel As control
    Dim subtitleLabel As control
    Dim scrollFrame As control
    Dim itemRadio As control
    Dim i As Integer
    Dim yPos As Integer
    
    ' Header Label
    Set headerLabel = Me.Controls.Add("Forms.Label.1", "lblHeader")
    With headerLabel
        .Left = 20
        .Top = 20
        .Width = 440
        .Height = 30
        .Caption = filterConfig.Title
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
        .Caption = filterConfig.Subtitle
        .ForeColor = RGB(67, 67, 67)
        .Font.Size = 9
    End With
    
    ' Scrollable Frame for Items
    Set scrollFrame = Me.Controls.Add("Forms.Frame.1", "frameScroll")
    With scrollFrame
        .Left = 20
        .Top = 100
        .Width = 440
        .Height = 350
        .BackColor = RGB(255, 255, 255)
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = itemCount * 25 + 20
    End With
    
    ' Create radio buttons for items
    yPos = 10
    For i = 1 To itemCount
        Set itemRadio = scrollFrame.Controls.Add("Forms.OptionButton.1", "optItem" & i)
        With itemRadio
            .Left = 15
            .Top = yPos
            .Width = 400
            .Height = 20
            .Caption = itemData(i)
            .Font.Size = 9
            .Tag = itemData(i)
            .GroupName = "FilterGroup"
        End With
        
        yPos = yPos + 25
    Next i
    
    ' Create Continue Button
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
    
    ' Create Cancel Button
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

Private Sub btnContinue_Click()
    Dim i As Integer
    Dim optButton As control
    Dim selectedItem As String
    
    ' Clear previous selections
    Set selectedItems = New Collection
    selectedItem = ""
    
    ' Check which item is selected (radio button)
    For i = 1 To itemCount
        On Error Resume Next
        Set optButton = Me.Controls("frameScroll").Controls("optItem" & i)
        If Err.Number = 0 Then
            If optButton.value = True Then
                selectedItem = optButton.Tag
                selectedItems.Add optButton.Tag
                Exit For
            End If
        End If
        On Error GoTo 0
    Next i
    
    If selectedItem = "" Then
        MsgBox "Please select one " & filterConfig.Category.filterLevel & ".", vbExclamation, "Elyse Energy"
        Exit Sub
    End If
    
    ' Update parameter
    MsgBox selectedItem & " sélectionné"
    
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    MsgBox "Operation cancelled.", vbInformation, "Elyse Energy"
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Ne rien faire ici, l'initialisation se fait via Initialize(category)
End Sub

' Public function to get selected items
Public Function GetSelectedItems() As Collection
    Set GetSelectedItems = selectedItems
End Function
