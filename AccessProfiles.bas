Attribute VB_Name = "AccessProfiles"
Option Explicit

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
    Call AddProfile(Engineer_Basic, "Basic Engineer", _
               True, False, True, False, Array())
               
    ' Chef de projet (accès Tools + Projets)
    Call AddProfile(Project_Manager, "Project Manager", _
               False, False, True, True, Array())
               
    ' Contrôleur financier (Finance + tous les budgets)
    Call AddProfile(Finance_Controller, "Finance Controller", _
               False, True, True, False, Array())
               
    ' Directeur technique (Tout Engineering + Tools + All Projects)
    Call AddProfile(Technical_Director, "Technical Director", _
               True, False, True, True, Array())
               
    ' Business Analyst (Tools + Finance partiel)
    Call AddProfile(Business_Analyst, "Business Analyst", _
               False, True, True, False, Array())
               
    ' Admin (accès total)
    Call AddProfile(Full_Admin, "Admin (Full Access)", _
               True, True, True, True, Array())
               
    ' Par défaut, on commence avec le profil Technical Director
    mCurrentProfile = Technical_Director
    mProfilesInitialized = True ' Set to true on successful completion
    Exit Sub
    
ErrorHandler:
    mProfilesInitialized = False ' Ensure it's false if an error occurs
    SYS_Logger.Log "profile_error", "Erreur VBA dans " & "AccessProfiles" & "." & "InitializeDemoProfiles" & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, "InitializeDemoProfiles", "AccessProfiles"
    HandleError "AccessProfiles", "InitializeDemoProfiles", "Erreur lors de l'initialisation des profils de démonstration"
End Sub

Private Sub AddProfile(id As DemoProfile, Name As String, _
                      eng As Boolean, fin As Boolean, Tools As Boolean, _
                      allProj As Boolean, Projects As Variant)
    Const PROC_NAME As String = "AddProfile"
    Const MODULE_NAME As String = "AccessProfiles"
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
    SYS_Logger.Log "profile_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de l'ajout du profil " & Name
End Sub

' Définit le profil actif
Public Sub SetCurrentProfile(profile As DemoProfile)
    Const PROC_NAME As String = "SetCurrentProfile"
    Const MODULE_NAME As String = "AccessProfiles"
    On Error GoTo ErrorHandler
    
    If profile < Engineer_Basic Or profile > Full_Admin Then
        HandleError MODULE_NAME, PROC_NAME, "Profil invalide: " & profile
        Exit Sub
    End If
    
    mCurrentProfile = profile
    Exit Sub
    
ErrorHandler:
    SYS_Logger.Log "profile_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du changement de profil"
End Sub

' Vérifie si le profil actuel a accès à une fonctionnalité
Public Function HasAccess(feature As String) As Boolean
    Const PROC_NAME As String = "HasAccess"
    Const MODULE_NAME As String = "AccessProfiles"
    On Error GoTo ErrorHandler
    
    ' S'assurer que les profils sont initialisés avant toute vérification
    If Not mProfilesInitialized Then
        InitializeDemoProfiles ' Tentative de réinitialisation
        If Not mProfilesInitialized Then ' Si ça échoue encore, on sort en sécurité
            HasAccess = False
            Exit Function
        End If
    End If
    
    ' Accès total pour l'admin, sauf si on demande explicitement l'accès Admin
    If mCurrentProfile = Full_Admin And feature <> "Admin" Then
        HasAccess = True
        Exit Function
    End If
    
    Dim prof As AccessProfile
    prof = Profiles(mCurrentProfile) ' Accès direct au tableau, pas de Set
    
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
    SYS_Logger.Log "profile_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la vérification des droits d'accès pour: " & feature
    HasAccess = False
End Function

' Récupère le nom du profil actuel
Public Function GetCurrentProfileName() As String
    Const PROC_NAME As String = "GetCurrentProfileName"
    Const MODULE_NAME As String = "AccessProfiles"
    On Error GoTo ErrorHandler
    
    GetCurrentProfileName = Profiles(mCurrentProfile).Name
    Exit Function
    
ErrorHandler:
    SYS_Logger.Log "profile_error", "Erreur VBA dans " & MODULE_NAME & "." & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors de la récupération du nom du profil actuel"
    GetCurrentProfileName = "Profil inconnu"
End Function




