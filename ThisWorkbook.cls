VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'---------------------------------------------------------------------------------------
' Gère l'événement de changement de sélection sur n'importe quelle feuille du classeur.
'
' Rôle : Invalider le ruban pour mettre à jour l'état des contrôles contextuels,
'        comme le bouton "Recharger le tableau courant".
'---------------------------------------------------------------------------------------
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
  ' On vérifie que le ruban a bien été initialisé
  If Not gRibbon Is Nothing Then
      ' Invalide tout le ruban. C'est rapide et ça garantit que toutes
      ' les callbacks "getEnabled", "getVisible", etc. sont rappelées.
      gRibbon.Invalidate
  End If
End Sub