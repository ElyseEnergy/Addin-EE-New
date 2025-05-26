'Simulates thisworkbook module

Private Sub Workbook_Open()
    ' Initialisation du système
    APP_MainOrchestrator.Initialize

    ' Initialisation des profils d'accès
    InitializeDemoProfiles
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Nettoyage du système
    APP_MainOrchestrator.Shutdown

    ' Nettoyage des profils d'accès
    CleanupProfiles

    ' Libération du ruban global
    Set gRibbon = Nothing
End Sub
