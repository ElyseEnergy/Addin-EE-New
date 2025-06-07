Attribute VB_Name = "AccessProfiles"
' Module: AccessProfiles
' Gère les profils de démonstration pour les droits d'accès
Option Explicit

' Note: Utilise le type AccessProfile défini dans Types.bas

Public Enum DemoProfile
    Engineer_Basic = 0
    Project_Manager = 1
    Finance_Controller = 2
    Technical_Director = 3
    Business_Analyst = 4    ' Renommé de Multi_Project_Lead pour mieux refléter le rôle
    Full_Admin = 5
End Enum

Private mCurrentProfile As DemoProfile
Dim Profiles() As AccessProfile
Dim ProfilesCount As Long
Private mProfilesInitialized As Boolean ' Added module-level flag

' Initialisation des profils de démonstration
Public Sub InitializeDemoProfiles()
    On Error GoTo ErrorHandler
    
    mProfilesInitialized = False ' Initialize to false at the start
    ProfilesCount = 0
    Erase Profiles
    ReDim Profiles(0 To 5)  ' Allouer l'espace pour tous les profils dès le début
    
    ' Ingénieur de base (accès Engineering + Tools basiques)
    AddProfile Engineer_Basic, "Basic Engineer", _
               True, False, True, False, Array()
               
    ' Chef de projet (accès Tools + Projets)
    AddProfile Project_Manager, "Project Manager", _
               False, False, True, True, Array()
               
    ' Contrôleur financier (Finance + tous les budgets)
    AddProfile Finance_Controller, "Finance Controller", _
               False, True, True, False, Array()
               
    ' Directeur technique (Tout Engineering + Tools + All Projects)
    AddProfile Technical_Director, "Technical Director", _
               True, False, True, True, Array()
               
    ' Business Analyst (Tools + Finance partiel)
    AddProfile Business_Analyst, "Business Analyst", _
               False, True, True, False, Array()
               
    ' Admin (accès total)
    AddProfile Full_Admin, "Admin (Full Access)", _
               True, True, True, True, Array()
               
    ' Par défaut, on commence avec le profil Technical Director
    mCurrentProfile = Technical_Director
    mProfilesInitialized = True ' Set to true on successful completion
    Exit Sub
    
ErrorHandler:
    mProfilesInitialized = False ' Ensure it's false if an error occurs
    HandleError "AccessProfiles", "InitializeDemoProfiles", "Erreur lors de l'initialisation des profils de démonstration"
End Sub

Private Sub AddProfile(id As DemoProfile, Name As String, _
                      eng As Boolean, fin As Boolean, Tools As Boolean, _
                      allProj As Boolean, Projects As Variant)
    On Error GoTo ErrorHandler
    
    Profiles(id).Name = Name
    Profiles(id).Description = Name
    Profiles(id).Engineering = eng
    Profiles(id).Finance = fin
    Profiles(id).Tools = Tools
    Profiles(id).AllProjects = allProj
    
    Set Profiles(id).Projects = New Collection
    Dim proj As Variant
    For Each proj In Projects
        Profiles(id).Projects.Add CStr(proj)
    Next
    
    ProfilesCount = id
    Exit Sub
    
ErrorHandler:
    HandleError "AccessProfiles", "AddProfile", "Erreur lors de l'ajout du profil " & Name
End Sub

' Définit le profil actif
Public Sub SetCurrentProfile(profile As DemoProfile)
    On Error GoTo ErrorHandler
    
    If profile < Engineer_Basic Or profile > Full_Admin Then
        HandleError "AccessProfiles", "SetCurrentProfile", "Profil invalide: " & profile
        Exit Sub
    End If
    
    mCurrentProfile = profile
    Exit Sub
    
ErrorHandler:
    HandleError "AccessProfiles", "SetCurrentProfile", "Erreur lors du changement de profil"
End Sub

' Récupère le profil par ID (suppose que l'ID correspond à l'index)
Private Function GetProfileById(id As DemoProfile) As AccessProfile
    On Error GoTo ErrorHandler
    Const PROC_NAME As String = "GetProfileById"
    Const MODULE_NAME_STR As String = "AccessProfiles"
    
    If Not mProfilesInitialized Then
        SYS_Logger.Log "profile_error", "Tentative d'accès au profil ID " & id & " mais les profils ne sont pas initialisés (mProfilesInitialized=False).", WARNING_LEVEL, PROC_NAME, MODULE_NAME_STR
        Exit Function ' Or handle error appropriately, e.g., return an empty/default profile
    End If
    
    ' Combined and clarified boundary checks
    If id < LBound(Profiles) Or id > UBound(Profiles) Then ' Added UBound check
        SYS_Logger.Log "profile_error", "ID de profil " & id & " hors limites (LBound: " & LBound(Profiles) & ", UBound: " & UBound(Profiles) & ").", ERROR_LEVEL, PROC_NAME, MODULE_NAME_STR
        ' Consider returning a default/empty profile or raising a more specific error
        Exit Function
    End If
    
    ' Check if Profiles array has been initialized (basic check)
    ' This check is now largely covered by mProfilesInitialized and the LBound/UBound check above.
    ' The IsEmpty/Name="" check below is still valuable for individual profile validity.
    ' If ProfilesCount = 0 And id <> ProfilesCount Then ' ProfilesCount is last assigned ID, so if 0, only Profiles(0) might be valid if ever assigned directly.
                                                ' More robustly, check if Profiles(id).Name is empty if not all profiles are guaranteed to be filled.
    If IsEmpty(Profiles(id).Name) Or Profiles(id).Name = "" Then
        SYS_Logger.Log "profile_error", "Profil ID " & id & " est dans les limites mais non rempli (nom vide).", WARNING_LEVEL, PROC_NAME, MODULE_NAME_STR
        Exit Function
    End If
    ' End If

    GetProfileById = Profiles(id)
    Exit Function
    
ErrorHandler:
    ' Log the specific error from Err object BEFORE calling HandleError, which might reset it.
    SYS_Logger.Log "profile_error", "Erreur VBA dans " & MODULE_NAME_STR & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description & ", Source: " & Err.Source, ERROR_LEVEL, PROC_NAME, MODULE_NAME_STR
    HandleError MODULE_NAME_STR, PROC_NAME, "Erreur lors de la récupération du profil ID: " & id ' HandleError might be a more generic handler
    ' To prevent returning an uninitialized AccessProfile object, which can cause further errors:
    ' One option is to clear the return object or set it to a known safe state if possible,
    ' but since it's a UDT, direct clearing is tricky. The Exit Function above is safer.
    ' If absolutely necessary, and if AccessProfile had an 'IsValid' flag or similar:
    ' Dim emptyProfile as AccessProfile
    ' GetProfileById = emptyProfile ' Or set a flag within it
End Function

' Vérifie si le profil actuel a accès à une fonctionnalité
Public Function HasAccess(feature As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim prof As AccessProfile
    prof = GetProfileById(mCurrentProfile)
    
    ' Accès total pour l'admin
    If mCurrentProfile = Full_Admin Then
        HasAccess = True
        Exit Function
    End If
    
    Select Case feature
        Case "Engineering"
            HasAccess = prof.Engineering
        Case "Finance"
            HasAccess = prof.Finance
        Case "Tools"
            HasAccess = prof.Tools
        Case "Admin"
            HasAccess = (mCurrentProfile = Full_Admin)
        Case Else
            ' Pour les projets (plus de référence aux projets spécifiques)
            HasAccess = prof.AllProjects
    End Select
    Exit Function
    
ErrorHandler:
    HandleError "AccessProfiles", "HasAccess", "Erreur lors de la vérification des droits d'accès pour: " & feature
    HasAccess = False
End Function

' Récupère le nom du profil actuel
Public Function GetCurrentProfileName() As String
    On Error GoTo ErrorHandler
    
    GetCurrentProfileName = Profiles(mCurrentProfile).Name
    Exit Function
    
ErrorHandler:
    HandleError "AccessProfiles", "GetCurrentProfileName", "Erreur lors de la récupération du nom du profil actuel"
    GetCurrentProfileName = "Profil inconnu"
End Function




