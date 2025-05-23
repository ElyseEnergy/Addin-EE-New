' Module : Types.bas
' Centralise tous les types personnalisés du projet
Option Explicit

' Type pour stocker les informations de catégorie
Public Type CategoryInfo
    CategoryName As String        ' Technologies ou Utilities
    FilterLevel As String         ' Niveau de filtrage (Molecule Type, Brand, Type, etc.)
    SecondaryFilterLevel As String ' Niveau de filtrage secondaire (ex: version de planning)
    DisplayName As String         ' Nom d'affichage pour l'utilisateur
    URL As String                 ' URL complète
    PowerQueryName As String      ' Nom de la requête PowerQuery
    CategoryGroup As String       ' Groupe de la catégorie
    SheetName As String           ' Nom de la feuille Ragic associée
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

' Type pour gérer les profils d'accès
Public Type AccessProfile
    Name As String
    Description As String
    Engineering As Boolean
    Finance As Boolean
    Tools As Boolean
    AllProjects As Boolean
    Projects As Collection  ' Projets spécifiques
End Type