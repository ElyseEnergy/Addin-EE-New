VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataSelector 
   Caption         =   "Select Data to Load"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7515
   OleObjectBlob   =   "frmDataSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==================================================================================================
'               VARIABLES PUBLIQUES POUR RETOURNER LES RÉSULTATS
'==================================================================================================
Public SelectedValues As Collection
Public ModeTransposed As Boolean
Public FinalDestination As Range
Public IsCancelled As Boolean

'==================================================================================================
'               VARIABLES PRIVÉES DU FORMULAIRE
'==================================================================================================
Private m_Category As CategoryInfo
Private m_SourceTable As ListObject
Private WithEvents m_RefEdit As RefEdit.RefEdit
Attribute m_RefEdit.VB_VarHelpID = -1

'==================================================================================================
'               MÉTHODES PUBLIQUES
'==================================================================================================

' Point d'entrée principal pour initialiser et montrer le formulaire
Public Sub ShowForCategory(cat As CategoryInfo)
    Set m_Category = cat
    Set m_SourceTable = wsPQData.ListObjects("Table_" & Utilities.SanitizeTableName(m_Category.PowerQueryName))
    
    If m_SourceTable Is Nothing Then
        MsgBox "Source data table not found for " & m_Category.DisplayName, vbCritical
        IsCancelled = True
        Unload Me
        Exit Sub
    End If
    
    Me.Show
End Sub

'==================================================================================================
'               ÉVÉNEMENTS DU FORMULAIRE ET DES CONTRÔLES
'==================================================================================================

Private Sub UserForm_Initialize()
    ' Initialiser les collections et l'état par défaut
    Set SelectedValues = New Collection
    IsCancelled = True ' Par défaut, on considère l'opération annulée
    
    ' Configurer l'apparence des contrôles
    SetupControls
    
    ' Peuple la liste de filtre primaire si nécessaire
    PopulatePrimaryFilterList
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Si l'utilisateur ferme avec la croix, s'assurer que c'est traité comme une annulation
    If CloseMode = vbFormControlMenu Then
        IsCancelled = True
    End If
End Sub

' L'utilisateur a cliqué sur un élément du premier filtre
Private Sub lstPrimaryFilter_Click()
    PopulateDataSheetList
End Sub

' L'utilisateur a cliqué sur OK
Private Sub cmdOK_Click()
    ' Valider les sélections
    If Not ValidateSelections() Then Exit Sub
    
    ' Récupérer les sélections
    CollectSelections
    
    ' Marquer comme non annulé et fermer
    IsCancelled = False
    Me.Hide
End Sub

' L'utilisateur a cliqué sur Annuler
Private Sub cmdCancel_Click()
    IsCancelled = True
    Me.Hide
End Sub

'==================================================================================================
'               LOGIQUE INTERNE DU FORMULAIRE
'==================================================================================================

' Configure l'apparence et les propriétés initiales des contrôles
Private Sub SetupControls()
    Me.Caption = "Load Data: " & m_Category.DisplayName
    
    ' Boutons OK/Annuler
    cmdOK.Caption = "OK"
    cmdCancel.Caption = "Cancel"
    
    ' Options de collage
    optNormal.Caption = "Paste as Table (Normal)"
    optTransposed.Caption = "Paste as Table (Transposed)"
    optNormal.Value = True ' Valeur par défaut
    
    ' Listes
    lstPrimaryFilter.MultiSelect = fmMultiSelectMulti
    lstDataSheets.MultiSelect = fmMultiSelectMulti
    lblFilter.Caption = m_Category.FilterLevel
    lblSheets.Caption = "Available Data Sheets"
    
    ' Ajouter le contrôle RefEdit dynamiquement
    Set m_RefEdit = Me.Controls.Add("RefEdit.RefEdit", "refeditDestination", True)
    With m_RefEdit
        .top = 4000
        .left = 120
        .Width = 7200
        .Height = 375
    End With
End Sub

' Peuple la liste de filtre primaire
Private Sub PopulatePrimaryFilterList()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    
    Dim i As Long
    Dim filterColIndex As Long
    
    On Error Resume Next
    filterColIndex = m_SourceTable.ListColumns(m_Category.FilterLevel).Index
    If Err.Number <> 0 Then ' Pas de colonne de filtre
        lblFilter.Visible = False
        lstPrimaryFilter.Visible = False
        ' Si pas de filtre, on peuple directement la seconde liste
        PopulateDataSheetList
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Remplir le dictionnaire avec les valeurs uniques
    For i = 1 To m_SourceTable.DataBodyRange.Rows.Count
        Dim val As String
        val = CStr(m_SourceTable.DataBodyRange.Cells(i, filterColIndex).Value)
        If Not dict.Exists(val) Then
            dict.Add val, Nothing
        End If
    Next i
    
    ' Peuple la ListBox
    lstPrimaryFilter.List = dict.Keys
End Sub

' Peuple la liste des fiches de données en fonction du filtre primaire sélectionné
Private Sub PopulateDataSheetList()
    Dim selectedFilters As New Collection
    Dim i As Long
    
    ' Récupérer les filtres sélectionnés
    For i = 0 To lstPrimaryFilter.ListCount - 1
        If lstPrimaryFilter.Selected(i) Then
            selectedFilters.Add lstPrimaryFilter.List(i)
        End If
    Next i
    
    lstDataSheets.Clear
    
    ' Si pas de filtre primaire (ou pas visible), on affiche tout
    If selectedFilters.Count = 0 And lstPrimaryFilter.Visible = False Then
        For i = 1 To m_SourceTable.DataBodyRange.Rows.Count
            lstDataSheets.AddItem m_SourceTable.DataBodyRange.Cells(i, 2).Value, m_SourceTable.DataBodyRange.Cells(i, 1).Value
        Next i
        Exit Sub
    End If
    
    ' Sinon, on filtre
    Dim filterColIndex As Long
    filterColIndex = m_SourceTable.ListColumns(m_Category.FilterLevel).Index
    
    For i = 1 To m_SourceTable.DataBodyRange.Rows.Count
        Dim filterVal As String
        filterVal = CStr(m_SourceTable.DataBodyRange.Cells(i, filterColIndex).Value)
        
        Dim item As Variant
        For Each item In selectedFilters
            If filterVal = item Then
                ' AddItem prend (texte, valeur cachée). On stocke l'ID dans la valeur cachée.
                lstDataSheets.AddItem m_SourceTable.DataBodyRange.Cells(i, 2).Value, m_SourceTable.DataBodyRange.Cells(i, 1).Value
                Exit For
            End If
        Next item
    Next i
End Sub

' Valide que les sélections de l'utilisateur sont cohérentes
Private Function ValidateSelections() As Boolean
    ' Au moins une fiche doit être sélectionnée
    Dim i As Long, count As Long
    For i = 0 To lstDataSheets.ListCount - 1
        If lstDataSheets.Selected(i) Then count = count + 1
    Next i
    If count = 0 Then
        MsgBox "Please select at least one data sheet to load.", vbExclamation
        lstDataSheets.SetFocus
        ValidateSelections = False
        Exit Function
    End If
    
    ' La destination doit être une plage valide
    On Error Resume Next
    Set FinalDestination = Range(m_RefEdit.Value)
    If Err.Number <> 0 Or FinalDestination Is Nothing Then
        MsgBox "Please select a valid destination cell.", vbExclamation
        m_RefEdit.SetFocus
        ValidateSelections = False
        Exit Function
    End If
    On Error GoTo 0
    
    ' S'assurer qu'une seule cellule est sélectionnée
    If FinalDestination.Cells.Count > 1 Then
        Set FinalDestination = FinalDestination.Cells(1, 1)
    End If
    
    ValidateSelections = True
End Function

' Récupère les choix de l'utilisateur et les stocke dans les variables publiques
Private Sub CollectSelections()
    ' Récupérer les ID des fiches sélectionnées
    Dim i As Long
    For i = 0 To lstDataSheets.ListCount - 1
        If lstDataSheets.Selected(i) Then
            ' On récupère la valeur cachée (l'ID) avec .List(i, 1) ou .BoundValue
            SelectedValues.Add lstDataSheets.List(i, lstDataSheets.BoundColumn - 1)
        End If
    Next i
    
    ' Récupérer le mode de collage
    ModeTransposed = optTransposed.Value
    
    ' La destination a déjà été définie dans la fonction de validation
End Sub 