Attribute VB_Name = "AccessProfiles"
' Module: AccessProfiles
' G�re les profils de d�monstration pour les droits d'acc�s
Option Explicit

' Note: Utilise le type AccessProfile d�fini dans Types.bas

Public Enum DemoProfile
    Engineer_Basic = 0
    Project_Manager = 1
    Finance_Controller = 2
    Technical_Director = 3
    Business_Analyst = 4    ' Renomm� de Multi_Project_Lead pour mieux refl�ter le r�le
    Full_Admin = 5
End Enum

Private mCurrentProfile As DemoProfile
Dim Profiles() As AccessProfile
Dim ProfilesCount As Long

' Initialisation des profils de d�monstration
Public Sub InitializeDemoProfiles()
    On Error GoTo ErrorHandler
    
    ProfilesCount = 0
    Erase Profiles
    ReDim Profiles(0 To 5)  ' Allouer l'espace pour tous les profils d�s le d�but
    
    ' Ing�nieur de base (acc�s Engineering + Tools basiques)
    AddProfile Engineer_Basic, "Basic Engineer", _
               True, False, True, False, Array()
               
    ' Chef de projet (acc�s Tools + Projets)
    AddProfile Project_Manager, "Project Manager", _
               False, False, True, True, Array()
               
    ' Contr�leur financier (Finance + tous les budgets)
    AddProfile Finance_Controller, "Finance Controller", _
               False, True, True, False, Array()
               
    ' Directeur technique (Tout Engineering + Tools + All Projects)
    AddProfile Technical_Director, "Technical Director", _
               True, False, True, True, Array()
               
    ' Business Analyst (Tools + Finance partiel)
    AddProfile Business_Analyst, "Business Analyst", _
               False, True, True, False, Array()
               
    ' Admin (acc�s total)
    AddProfile Full_Admin, "Admin (Full Access)", _
               True, True, True, True, Array()
               
    ' Par d�faut, on commence avec le profil Technical Director
    mCurrentProfile = Technical_Director
    Exit Sub
    
ErrorHandler:
    HandleError "AccessProfiles", "InitializeDemoProfiles", "Erreur lors de l'initialisation des profils de d�monstration"
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

' D�finit le profil actif
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

' R�cup�re le profil par ID (suppose que l'ID correspond � l'index)
Private Function GetProfileById(id As DemoProfile) As AccessProfile
    On Error GoTo ErrorHandler
    
    If id < Engineer_Basic Or id > Full_Admin Then
        HandleError "AccessProfiles", "GetProfileById", "ID de profil invalide: " & id
        Exit Function
    End If
    
    GetProfileById = Profiles(id)
    Exit Function
    
ErrorHandler:
    HandleError "AccessProfiles", "GetProfileById", "Erreur lors de la r�cup�ration du profil"
End Function

' V�rifie si le profil actuel a acc�s � une fonctionnalit�
Public Function HasAccess(feature As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim prof As AccessProfile
    prof = GetProfileById(mCurrentProfile)
    
    ' Acc�s total pour l'admin
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
            ' Pour les projets (plus de r�f�rence aux projets sp�cifiques)
            HasAccess = prof.AllProjects
    End Select
    Exit Function
    
ErrorHandler:
    HandleError "AccessProfiles", "HasAccess", "Erreur lors de la v�rification des droits d'acc�s pour: " & feature
    HasAccess = False
End Function

' R�cup�re le nom du profil actuel
Public Function GetCurrentProfileName() As String
    On Error GoTo ErrorHandler
    
    GetCurrentProfileName = Profiles(mCurrentProfile).Name
    Exit Function
    
ErrorHandler:
    HandleError "AccessProfiles", "GetCurrentProfileName", "Erreur lors de la r�cup�ration du nom du profil actuel"
    GetCurrentProfileName = "Profil inconnu"
End Function




