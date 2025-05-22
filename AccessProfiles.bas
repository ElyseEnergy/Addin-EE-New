' Module: AccessProfiles
' Gère les profils de démonstration pour les droits d'accès
Option Explicit

' Note: Utilise le type AccessProfile défini dans Types.bas

Public Enum DemoProfile
    Engineer_Basic = 1
    Project_Manager = 2
    Finance_Controller = 3
    Technical_Director = 4
    Business_Analyst = 5    ' Renommé de Multi_Project_Lead pour mieux refléter le rôle
    Full_Admin = 6 
End Enum

Private mCurrentProfile As DemoProfile
Dim Profiles() As AccessProfile
Dim ProfilesCount As Long

' Initialisation des profils de démonstration
Public Sub InitializeDemoProfiles()
    ProfilesCount = 0
    Erase Profiles
    
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
End Sub

Private Sub AddProfile(id As DemoProfile, Name As String, _
                      eng As Boolean, fin As Boolean, tools As Boolean, _
                      allProj As Boolean, projects As Variant)
    ReDim Preserve Profiles(1 To CInt(id))
    ProfilesCount = CInt(id)
    
    Profiles(id).Name = Name
    Profiles(id).Description = Name
    Profiles(id).Engineering = eng
    Profiles(id).Finance = fin
    Profiles(id).Tools = tools
    Profiles(id).AllProjects = allProj
    
    Set Profiles(id).Projects = New Collection
    Dim proj As Variant
    For Each proj In projects
        Profiles(id).Projects.Add CStr(proj)
    Next
    
    ProfilesCount = id
End Sub

' Définit le profil actif
Public Sub SetCurrentProfile(profile As DemoProfile)
    mCurrentProfile = profile
End Sub

' Récupère le profil par ID (suppose que l'ID correspond à l'index)
Private Function GetProfileById(id As DemoProfile) As AccessProfile
    If id >= 1 And id <= ProfilesCount Then
        GetProfileById = Profiles(id)
    End If
End Function

' Vérifie si le profil actuel a accès à une fonctionnalité
Public Function HasAccess(feature As String) As Boolean
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
End Function

' Récupère le nom du profil actuel
Public Function GetCurrentProfileName() As String
    GetCurrentProfileName = Profiles(mCurrentProfile).Name
End Function
