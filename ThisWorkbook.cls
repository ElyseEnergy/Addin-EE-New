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
    ' Initialisation du syst�me
    Call APP_MainOrchestrator.InitializeElyseSystem(DEBUG_MODE, True)

    ' Initialisation des profils d'acc�s
    InitializeDemoProfiles
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Nettoyage du syst�me
    

    ' Nettoyage des profils d'acc�s
    

    ' Lib�ration du ruban global
    Set gRibbon = Nothing
End Sub

