Attribute VB_Name = "DataInteraction"
Option Explicit

' ==========================================
' Module DataInteraction
' ------------------------------------------
' Ce module gère toutes les interactions avec l'utilisateur pendant le processus de chargement.
' Il centralise les boîtes de dialogue et les sélections utilisateur.
' ==========================================

Private Const MODULE_NAME As String = "DataInteraction"

' Demande à l'utilisateur de sélectionner les valeurs à charger.
Public Function GetSelectedValues(Category As CategoryInfo) As Collection
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "GetSelectedValues"
    
    ' Récupérer la table source
    Dim sourceTable As ListObject
    Set sourceTable = DataLoaderManager.GetOrCreatePQDataSheet.ListObjects("Table_" & Utilities.SanitizeTableName(Category.PowerQueryName))
    
    ' Vérifier que la table existe
    If sourceTable Is Nothing Then
        MsgBox "La table source n'existe pas.", vbExclamation
        Exit Function
    End If
    
    ' Vérifier que la table a des données
    If sourceTable.ListRows.Count = 0 Then
        ' Forcer un rafraîchissement de la requête
        On Error Resume Next
        sourceTable.QueryTable.Refresh
        On Error GoTo ErrorHandler
        
        ' Attendre un peu que les données se chargent
        Application.Wait Now + TimeSerial(0, 0, 1)
        
        ' Vérifier à nouveau
        If sourceTable.ListRows.Count = 0 Then
            MsgBox "Aucune donnée disponible dans la table source.", vbExclamation
            Exit Function
        End If
    End If
    
    ' Créer une collection pour stocker les valeurs sélectionnées
    Set GetSelectedValues = New Collection
    
    ' Préparer la liste des valeurs disponibles
    Dim values() As String
    Dim i As Long
    Dim rowCount As Long
    rowCount = sourceTable.ListRows.Count
    
    ReDim values(1 To rowCount)
    
    ' Vérifier si la table a plus d'une colonne pour afficher les intitulés
    If sourceTable.ListColumns.Count > 1 Then
        ' Construire la liste avec "ID - Intitulé" (colonne 2)
        For i = 1 To rowCount
            values(i) = CStr(sourceTable.ListRows(i).Range(1).Value) & "  -  " & CStr(sourceTable.ListRows(i).Range(2).Value)
        Next i
    Else
        ' Comportement par défaut si une seule colonne: ID uniquement
        For i = 1 To rowCount
            values(i) = CStr(sourceTable.ListRows(i).Range(1).Value)
        Next i
    End If
    
    ' Afficher la boîte de dialogue de sélection
    Dim selectedValues As String
    Dim availableValuesPrompt As String
    
    ' Joindre les valeurs avec un retour à la ligne pour une meilleure lisibilité
    availableValuesPrompt = Join(values, vbCrLf)
    
    ' Log pour le débogage
    Log "GetSelectedValues", "Nombre de valeurs disponibles : " & rowCount, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    selectedValues = InputBox("Entrez les ID (seulement les numéros), séparés par des virgules." & vbCrLf & vbCrLf & _
                            "Valeurs disponibles :" & vbCrLf & availableValuesPrompt, _
                            "Sélection pour " & Category.DisplayName)
    
    ' Vérifier si l'utilisateur a annulé
    If selectedValues = "" Then Exit Function
    
    ' Nettoyer et valider les valeurs sélectionnées
    Dim selectedArray() As String
    selectedArray = Split(selectedValues, ",")
    
    ' Ajouter les valeurs à la collection
    Dim v As Variant
    For Each v In selectedArray
        ' Nettoyer la valeur
        v = Trim(v)
        If v <> "" Then
            ' Vérifier si la valeur existe dans la table source
            Dim found As Boolean
            found = False
            For i = 1 To rowCount
                If CStr(sourceTable.ListRows(i).Range(1).Value) = v Then
                    found = True
                    Exit For
                End If
            Next i
            
            If found Then
                GetSelectedValues.Add v
            Else
                MsgBox "La valeur '" & v & "' n'existe pas dans la table source.", vbExclamation
                Set GetSelectedValues = Nothing
                Exit Function
            End If
        End If
    Next v
    
    ' Vérifier qu'au moins une valeur a été sélectionnée
    If GetSelectedValues.Count = 0 Then
        MsgBox "Aucune valeur valide n'a été sélectionnée.", vbExclamation
        Set GetSelectedValues = Nothing
    End If
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la sélection des valeurs"
    Set GetSelectedValues = Nothing
End Function

' Demande à l'utilisateur de sélectionner la destination du collage.
Public Function GetDestination(loadInfo As DataLoadInfo) As Range
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "GetDestination"
    
    ' Afficher la boîte de dialogue de sélection
    Dim message As String
    message = "Sélectionnez la cellule de destination pour " & loadInfo.Category.DisplayName & "." & vbCrLf & _
             "Mode : " & IIf(loadInfo.ModeTransposed, "TRANSPOSÉ (colonnes)", "NORMAL (lignes)")
    
    Dim destination As Range
    On Error Resume Next
    Set destination = Application.InputBox(message, "Destination", Type:=8)
    On Error GoTo ErrorHandler
    
    ' Vérifier si l'utilisateur a annulé
    If destination Is Nothing Then Exit Function
    
    ' Vérifier que la destination est une seule cellule
    If destination.Cells.Count > 1 Then
        MsgBox "Veuillez sélectionner une seule cellule.", vbExclamation
        Set GetDestination = Nothing
        Exit Function
    End If
    
    ' Vérifier que la destination n'est pas dans la feuille PQ_DATA
    If destination.Parent.Name = "PQ_DATA" Then
        MsgBox "La feuille PQ_DATA est réservée aux données brutes. Veuillez choisir une autre destination.", vbExclamation
        Set GetDestination = Nothing
        Exit Function
    End If
    
    Set GetDestination = destination
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la sélection de la destination"
    Set GetDestination = Nothing
End Function 