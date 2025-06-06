Attribute VB_Name = "Types"
' Module : Types.bas
' Centralise tous les types personnalis�s du projet
Option Explicit

' Type pour stocker les informations de cat�gorie
Public Type CategoryInfo
    CategoryName As String        ' Technologies ou Utilities
    FilterLevel As String         ' Niveau de filtrage (Molecule Type, Brand, Type, etc.)
    SecondaryFilterLevel As String ' Niveau de filtrage secondaire (ex: version de planning)
    DisplayName As String         ' Nom d'affichage pour l'utilisateur
    URL As String                 ' URL compl�te
    PowerQueryName As String      ' Nom de la requ�te PowerQuery
    CategoryGroup As String       ' Groupe de la cat�gorie
    SheetName As String           ' Nom de la feuille Ragic associ�e
End Type

' Type pour stocker les informations de chargement de donn�es
Public Type DataLoadInfo
    Category As CategoryInfo
    SelectedValues As Collection
    ModeTransposed As Boolean
    FinalDestination As Range
    PreviewRows As Long
End Type

' --- Ajoutez ici d'autres types personnalis�s au besoin ---

' Type pour g�rer les profils d'acc�s
Public Type AccessProfile
    Name As String
    Description As String
    Engineering As Boolean
    Finance As Boolean
    Tools As Boolean
    AllProjects As Boolean
    Projects As Collection  ' Projets sp�cifiques
End Type
