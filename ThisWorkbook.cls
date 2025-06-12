Option Explicit

'---------------------------------------------------------------------------------------
' Gère l'événement de changement de sélection sur n'importe quelle feuille du classeur.
'
' Rôle : Invalider le ruban pour mettre à jour l'état des contrôles contextuels,
'        comme le bouton "Recharger le tableau courant".
'---------------------------------------------------------------------------------------
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
  On Error GoTo ErrorHandler
  ' Mettre à jour le statut du ruban en fonction de la sélection
  RibbonVisibility.UpdateTableToolsVisibility
  Exit Sub
ErrorHandler:
    ' Gestion d'erreur minimale pour cet événement fréquent
End Sub

Private Sub Workbook_Open()
    ' Initialisation des systèmes critiques de l'addin
    On Error Resume Next ' Gérer le cas où le logging n'est pas prêt
    SYS_Logger.InitializeLogger
    RibbonVisibility.DelayedInitialization
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error GoTo ErrorHandler
    ' Nettoyage avant la fermeture
    SYS_ErrorHandler.ShutdownErrorHandler
    Exit Sub
ErrorHandler:
    ' Gestion d'erreur minimale pour la fermeture
End Sub