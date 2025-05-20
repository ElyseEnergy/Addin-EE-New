' Module: RibbonVisibility
' Gère la visibilité des éléments du ruban
Option Explicit

' Variable globale pour stocker l'instance du ruban
Public gRibbon As IRibbonUI

' Callback appelé lors du chargement du ruban
Public Sub Ribbon_Load(ByVal ribbon As IRibbonUI)
    Debug.Print "Ribbon_Load appelé"
    Set gRibbon = ribbon
    Debug.Print "gRibbon initialisé"
End Sub

' Callback pour la visibilité du menu Technologies
Public Sub GetTechnologiesVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = True
End Sub

' Callback pour la visibilité du menu Utilities
Public Sub GetUtilitiesVisibility(control As IRibbonControl, ByRef visible As Variant)
    visible = True
End Sub

' Fonction pour forcer le rafraîchissement du ruban
Public Sub InvalidateRibbon()
    Debug.Print "InvalidateRibbon appelé"
    If Not gRibbon Is Nothing Then
        gRibbon.Invalidate
        Debug.Print "Ribbon invalidé"
    Else
        Debug.Print "gRibbon est Nothing"
    End If
End Sub 