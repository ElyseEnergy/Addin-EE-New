' Module: AccessProfiles
' Gère les profils de démonstration pour les droits d'accès
Option Explicit

Public Enum DemoProfile
    Engineer_Basic = 1
    Project_Manager = 2
    Finance_Controller = 3
    Technical_Director = 4
    Multi_Project_Lead = 5
End Enum

Private Type AccessProfile
    Name As String
    Description As String
    Engineering As Boolean
    Finance As Boolean
    Tools As Boolean
    AllProjects As Boolean
    Projects As Collection  ' Projets spécifiques
End Type

Private mCurrentProfile As DemoProfile
Private mProfiles As Collection

' Initialisation des profils de démonstration
Public Sub InitializeDemoProfiles()
    Set mProfiles = New Collection
    
    ' Ingénieur de base (accès Engineering + Tools)
    AddProfile Engineer_Basic, "Basic Engineer", _
               True, False, True, False, Array()
               
    ' Chef de projet Echo (accès Tools + Echo specific)
    AddProfile Project_Manager, "Echo Project Manager", _
               False, False, True, False, Array("Echo")
               
    ' Contrôleur financier (Finance + tous les CAPEX/OPEX)
    AddProfile Finance_Controller, "Finance Controller", _
               False, True, True, True, Array()
               
    ' Directeur technique (Tout Engineering + Tools + All Projects)
    AddProfile Technical_Director, "Technical Director", _
               True, False, True, True, Array()
               
    ' Multi-projets (Echo + EmRhone + Tools)
    AddProfile Multi_Project_Lead, "Multi-Project Leader", _
               False, False, True, False, Array("Echo", "EmRhone")
               
    ' Par défaut, on commence avec le profil Technical Director
    mCurrentProfile = Technical_Director
End Sub

Private Sub AddProfile(id As DemoProfile, Name As String, _
                      eng As Boolean, fin As Boolean, tools As Boolean, _
                      allProj As Boolean, projects As Variant)
    Dim prof As AccessProfile
    prof.Name = Name
    prof.Engineering = eng
    prof.Finance = fin
    prof.Tools = tools
    prof.AllProjects = allProj
    
    Set prof.Projects = New Collection
    Dim proj As Variant
    For Each proj In projects
        prof.Projects.Add CStr(proj)
    Next
    
    mProfiles.Add prof, CStr(id)
End Sub

' Définit le profil actif
Public Sub SetCurrentProfile(profile As DemoProfile)
    mCurrentProfile = profile
End Sub

' Vérifie si le profil actuel a accès à une fonctionnalité
Public Function HasAccess(feature As String) As Boolean
    Dim prof As AccessProfile
    prof = mProfiles(CStr(mCurrentProfile))
    
    Select Case feature
        Case "Engineering"
            HasAccess = prof.Engineering
        Case "Finance"
            HasAccess = prof.Finance
        Case "Tools"
            HasAccess = prof.Tools
        Case Else
            ' Pour les projets spécifiques
            If prof.AllProjects Then
                HasAccess = True
            Else
                Dim proj As Variant
                For Each proj In prof.Projects
                    If InStr(1, feature, proj, vbTextCompare) > 0 Then
                        HasAccess = True
                        Exit Function
                    End If
                Next
                HasAccess = False
            End If
    End Select
End Function

' Récupère le nom du profil actuel
Public Function GetCurrentProfileName() As String
    GetCurrentProfileName = mProfiles(CStr(mCurrentProfile)).Name
End Function
