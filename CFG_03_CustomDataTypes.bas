' Module : CFG_03_CustomDataTypes.bas
' Centralise tous les types personnalisés du projet
Option Explicit
Private Const MODULE_NAME As String = "CFG_03_CustomDataTypes"

' ============================================================================
' CATEGORY TYPES
' ============================================================================

' Type pour stocker les informations de catégorie
Public Type CategoryInfo
    categoryName As String        ' Technologies ou Utilities
    filterLevel As String         ' Niveau de filtrage (Molecule Type, Brand, Type, etc.)
    SecondaryFilterLevel As String ' Niveau de filtrage secondaire (ex: version de planning)
    displayName As String         ' Nom d'affichage pour l'utilisateur
    URL As String                 ' URL complète
    PowerQueryName As String      ' Nom de la requête PowerQuery
    categoryGroup As String       ' Groupe de la catégorie
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

' ============================================================================
' ACCESS CONTROL TYPES
' ============================================================================

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

' ============================================================================
' APPLICATION STATUS TYPES
' ============================================================================

Public Enum AppStatus
    stReady = 1
    stBusy = 2
    stError = 3
    stMaintenance = 4
End Enum

Public Enum DataRefreshMode
    rmManual = 1
    rmAutomatic = 2
    rmScheduled = 3
End Enum

' ============================================================================
' DATA VALIDATION TYPES
' ============================================================================

Public Type ValidationRule
    FieldName As String
    RuleType As String  ' Required, Range, Format, Custom
    MinValue As Variant
    MaxValue As Variant
    Format As String
    CustomRule As String
End Type

Public Type ValidationResult
    IsValid As Boolean
    ErrorMessage As String
    FieldName As String
    InvalidValue As Variant
End Type