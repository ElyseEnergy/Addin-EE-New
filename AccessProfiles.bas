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

' Initialisation des profils de démonstration
Public Sub InitializeDemoProfiles()
    On Error GoTo ErrorHandler
    
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
    Exit Sub
    
ErrorHandler:
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
    
    If id < Engineer_Basic Or id > Full_Admin Then
        HandleError "AccessProfiles", "GetProfileById", "ID de profil invalide: " & id
        Exit Function
    End If
    
    GetProfileById = Profiles(id)
    Exit Function
    
ErrorHandler:
    HandleError "AccessProfiles", "GetProfileById", "Erreur lors de la récupération du profil"
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




