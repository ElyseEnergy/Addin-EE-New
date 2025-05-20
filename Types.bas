' Module : Types.bas
' Centralise tous les types personnalisés du projet
Option Explicit

' Type pour stocker les informations de catégorie
Public Type CategoryInfo
    CategoryName As String        ' Technologies ou Utilities
    FilterLevel As String         ' Niveau de filtrage (Molecule Type, Brand, Type, etc.)
    DisplayName As String         ' Nom d'affichage pour l'utilisateur
    URL As String                 ' URL complète
    PowerQueryName As String      ' Nom de la requête PowerQuery
End Type

' Type pour stocker les informations de chargement de données
Public Type DataLoadInfo
    Category As CategoryInfo
    SelectedValues As Collection
    ModeTransposed As Boolean
    FinalDestination As Range
    PreviewRows As Long
End Type

' --- Ajoutez ici d'autres types personnalisés au besoin --- 