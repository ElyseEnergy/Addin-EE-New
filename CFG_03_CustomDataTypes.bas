' Module : Types.bas
' Centralise tous les types personnalisés du projet
Option Explicit
Private Const MODULE_NAME As String = "Types"

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

' This module likely contains Enum definitions or User-Defined Types (UDTs).
' Logging integration here would primarily be in any Subs/Functions that might exist
' for validating or manipulating these types, if any.
' If it only contains declarations, no direct logging changes are needed in this file itself,
' but the types defined here will be used by other modules where logging is applied.

' Example: If there was a function that used a UDT
' Public Type MyCustomType
'     ID As Long
'     Name As String
'     Value As Double
' End Type

' Public Sub ProcessCustomType(data As MyCustomType)
'    Const PROC_NAME As String = "ProcessCustomType"
'    On Error GoTo ErrorHandler
'
'    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Processing custom type: " & data.Name, PROC_NAME, MODULE_NAME
'
'    If data.ID = 0 Then
'        Debug.Print "Custom type ID is zero, potential issue."
'        ElyseMain_Orchestrator.LogWarning PROC_NAME & "_ZeroID", "Custom type ID is zero for " & data.Name, PROC_NAME, MODULE_NAME
'    End If
'
'    ' ... processing logic ...
'    Debug.Print "Finished processing " & data.Name
'    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_End", "Finished processing " & data.Name, PROC_NAME, MODULE_NAME
'    Exit Sub
'
'ErrorHandler:
'    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
' End Sub

' If this file ONLY contains Type definitions and Enums, no executable code, then no changes are needed here.
' For example:
' Public Enum StatusType
'    stPending
'    stInProgress
'    stCompleted
'    stFailed
' End Enum

' Public Type UserProfile
'    UserID As String
'    UserName As String
'    LastLogin As Date
' End Type

' If there are functions, apply logging as in other modules.
' If not, this file remains unchanged regarding logging calls.

' Assuming for now this file is primarily declarations, so no direct logging calls to add *within this file*.
' The types declared here will be used in procedures in other modules where logging is added.
ElyseMain_Orchestrator.LogDebug "ModuleLoad", "Types module (declarations) loaded.", "N/A", MODULE_NAME