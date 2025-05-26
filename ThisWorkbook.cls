VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Simulates thisworkbook module

Private Sub Workbook_Open()
    ' Initialisation du système
    Call APP_MainOrchestrator.InitializeElyseSystem(DEBUG_MODE, True)

    ' Initialisation des profils d'accès
    InitializeDemoProfiles
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Nettoyage du système
    

    ' Nettoyage des profils d'accès
    

    ' Libération du ruban global
    Set gRibbon = Nothing
End Sub

