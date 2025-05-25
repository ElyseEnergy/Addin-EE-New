'Simulates thisworkbook module

Private Sub Workbook_Open()
    ' Initialisation du système
    ElyseMain_Orchestrator.Initialize
    
    ' Initialisation des profils d'accès
    InitializeDemoProfiles
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Nettoyage du système
    ElyseMain_Orchestrator.Shutdown
    
    ' Nettoyage des profils d'accès
    CleanupProfiles
    
    ' Libération du ruban global
    Set gRibbon = Nothing
End Sub
