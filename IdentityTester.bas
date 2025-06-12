Attribute VB_Name = "IdentityTester"
Option Explicit

' ========================================================================
' DÉCLARATIONS DES API WINDOWS
' ========================================================================
#If VBA7 Then
    ' API pour obtenir les noms d'utilisateur et d'ordinateur
    Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare PtrSafe Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" _
        (ByVal NameFormat As Long, ByVal lpNameBuffer As String, ByRef nSize As Long) As Long
        
    ' API pour lire le registre
    Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
        (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, _
         ByVal samDesired As Long, ByRef phkResult As LongPtr) As Long
    Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As Long, _
         ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
    Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
    
    ' API pour les informations étendues de l'utilisateur (Groupes locaux, etc.)
    Private Declare PtrSafe Function NetUserGetInfo Lib "netapi32.dll" _
        (ByVal servername As String, ByVal username As String, ByVal level As Long, ByRef bufptr As LongPtr) As Long
    Private Declare PtrSafe Function NetUserGetLocalGroups Lib "netapi32.dll" _
        (ByVal servername As String, ByVal username As String, ByVal level As Long, _
         ByVal flags As Long, ByRef bufptr As LongPtr, ByVal prefmaxlen As Long, _
         ByRef entriesread As Long, ByRef totalentries As Long) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare PtrSafe Function NetApiBufferFree Lib "netapi32.dll" (ByVal Buffer As LongPtr) As Long
#Else
    ' --- Déclarations 32-bit ---
    Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" _
        (ByVal NameFormat As Long, ByVal lpNameBuffer As String, ByRef nSize As Long) As Long
        
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
         ByVal samDesired As Long, ByRef phkResult As Long) As Long
    Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
         ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
    Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
        
    Private Declare Function NetUserGetInfo Lib "netapi32.dll" _
        (ByVal servername As String, ByVal username As String, ByVal level As Long, ByRef bufptr As Long) As Long
    Private Declare Function NetUserGetLocalGroups Lib "netapi32.dll" _
        (ByVal servername As String, ByVal username As String, ByVal level As Long, _
         ByVal flags As Long, ByRef bufptr As Long, ByVal prefmaxlen As Long, _
         ByRef entriesread As Long, ByRef totalentries As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function NetApiBufferFree Lib "netapi32.dll" (ByVal Buffer As Long) As Long
#End If

' Constantes pour les formats de nom
Private Const NAME_UNKNOWN = 0
Private Const NAME_FULLY_QUALIFIED_DN = 1
Private Const NAME_SAM_COMPATIBLE = 2
Private Const NAME_DISPLAY = 3
Private Const NAME_UNIQUE_ID = 6
Private Const NAME_CANONICAL = 7
Private Const NAME_USER_PRINCIPAL = 8

' ========================================================================
' FONCTION PRINCIPALE - Lance tous les tests
' ========================================================================
Sub TestAllIdentityMethods()
    Const PROC_NAME As String = "TestAllIdentityMethods"
    On Error GoTo ErrorHandler
    Dim output As String
    Dim separator As String
    
    separator = String(80, "=")
    
    output = separator & vbCrLf
    output = output & "RAPPORT COMPLET DES MÉTHODES D'IDENTITÉ" & vbCrLf
    output = output & "Généré le: " & Now() & vbCrLf
    output = output & separator & vbCrLf & vbCrLf
    
    ' Test de chaque méthode
    output = output & TestVBANativeMethods() & vbCrLf
    output = output & TestEnvironmentVariables() & vbCrLf
    output = output & TestWindowsAPI() & vbCrLf
    output = output & TestWScriptNetwork() & vbCrLf
    output = output & TestWMIMethods() & vbCrLf
    output = output & TestOfficeInfo() & vbCrLf
    output = output & TestRegistryInfo() & vbCrLf
    output = output & TestActiveDirectory() & vbCrLf
    output = output & TestExtendedUserInfo() & vbCrLf
    
    ' Affichage du rapport complet
    Debug.Print output
    
    ' Optionnel: Sauvegarder dans un fichier
    Call SaveReportToFile(output)
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, "IdentityTester"
    HandleError "IdentityTester", PROC_NAME, "Erreur dans le testeur principal d'identité."
End Sub

' ========================================================================
' MÉTHODES VBA NATIVES
' ========================================================================
Function TestVBANativeMethods() As String
    Dim result As String
    result = "1. MÉTHODES VBA NATIVES" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    On Error Resume Next
    result = result & "Application.UserName: " & Application.UserName & vbCrLf
    result = result & "Application.OperatingSystem: " & Application.OperatingSystem & vbCrLf
    result = result & "Application.Version: " & Application.Version & vbCrLf
    result = result & "Application.Build: " & Application.Build & vbCrLf
    On Error GoTo 0
    
    TestVBANativeMethods = result & vbCrLf
End Function

' ========================================================================
' VARIABLES D'ENVIRONNEMENT
' ========================================================================
Function TestEnvironmentVariables() As String
    Const PROC_NAME As String = "TestEnvironmentVariables"
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim envVars As Variant
    Dim i As Integer
    
    result = "2. VARIABLES D'ENVIRONNEMENT" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    ' Liste des variables d'environnement importantes
    envVars = Array("USERNAME", "USERPROFILE", "USERDOMAIN", "USERDNSDOMAIN", _
                   "LOGONSERVER", "COMPUTERNAME", "SESSIONNAME", "CLIENTNAME", _
                   "APPDATA", "LOCALAPPDATA", "TEMP", "HOMEPATH", "HOMEDRIVE", _
                   "PROCESSOR_ARCHITECTURE", "OS", "NUMBER_OF_PROCESSORS")
    
    For i = 0 To UBound(envVars)
        On Error Resume Next
        result = result & envVars(i) & ": " & Environ(envVars(i)) & vbCrLf
        On Error GoTo 0
    Next i
    
    TestEnvironmentVariables = result & vbCrLf
    Exit Function
    
ErrorHandler:
    TestEnvironmentVariables = "Erreur lors du test des variables d'environnement."
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, "IdentityTester"
    HandleError "IdentityTester", PROC_NAME, "Erreur test variables d'environnement"
End Function

' ========================================================================
' API WINDOWS
' ========================================================================
Function TestWindowsAPI() As String
    Const PROC_NAME As String = "TestWindowsAPI"
    On Error GoTo ErrorHandler
    Dim result As String
    result = "3. API WINDOWS" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    result = result & "GetUserName API: " & GetCurrentUserAPI() & vbCrLf
    result = result & "GetComputerName API: " & GetComputerNameAPI() & vbCrLf
    result = result & "GetUserNameEx (Display): " & GetUserNameExtended(NAME_DISPLAY) & vbCrLf
    result = result & "GetUserNameEx (SAM): " & GetUserNameExtended(NAME_SAM_COMPATIBLE) & vbCrLf
    result = result & "GetUserNameEx (UPN): " & GetUserNameExtended(NAME_USER_PRINCIPAL) & vbCrLf
    result = result & "GetUserNameEx (DN): " & GetUserNameExtended(NAME_FULLY_QUALIFIED_DN) & vbCrLf
    
    TestWindowsAPI = result & vbCrLf
    Exit Function
ErrorHandler:
    TestWindowsAPI = "Erreur lors du test des API Windows."
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, "IdentityTester"
    HandleError "IdentityTester", PROC_NAME, "Erreur test API Windows"
End Function

Function GetCurrentUserAPI() As String
    Const PROC_NAME As String = "GetCurrentUserAPI"
    On Error GoTo ErrorHandler
    Dim buffer As String * 255
    Dim ret As Long
    ret = GetUserName(buffer, 255)
    If ret <> 0 Then
        GetCurrentUserAPI = Left$(buffer, InStr(buffer, Chr$(0)) - 1)
    Else
        GetCurrentUserAPI = "Erreur API"
    End If
    Exit Function
ErrorHandler:
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, "IdentityTester"
    HandleError "IdentityTester", PROC_NAME, "Erreur fatale GetCurrentUserAPI"
End Function

Function GetComputerNameAPI() As String
    Const PROC_NAME As String = "GetComputerNameAPI"
    On Error GoTo ErrorHandler
    Dim buffer As String * 255
    Dim ret As Long
    ret = GetComputerName(buffer, 255)
    If ret <> 0 Then
        GetComputerNameAPI = Left$(buffer, InStr(buffer, Chr$(0)) - 1)
    Else
        GetComputerNameAPI = "Erreur API"
    End If
    Exit Function
ErrorHandler:
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, "IdentityTester"
    HandleError "IdentityTester", PROC_NAME, "Erreur fatale GetComputerNameAPI"
End Function

Function GetUserNameExtended(NameFormat As Long) As String
    Const PROC_NAME As String = "GetUserNameExtended"
    On Error GoTo ErrorHandler
    Dim buffer As String * 255
    Dim bufferSize As Long
    Dim ret As Long
    
    bufferSize = 255
    ret = GetUserNameEx(NameFormat, buffer, bufferSize)
    
    If ret <> 0 Then
        GetUserNameExtended = Left$(buffer, bufferSize)
    Else
        GetUserNameExtended = "Non disponible"
    End If
    Exit Function
ErrorHandler:
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, "IdentityTester"
    HandleError "IdentityTester", PROC_NAME, "Erreur fatale GetUserNameExtended"
End Function

' ========================================================================
' WSCRIPT.NETWORK
' ========================================================================
Function TestWScriptNetwork() As String
    Const PROC_NAME As String = "TestWScriptNetwork"
    On Error GoTo ErrorHandler
    
    Dim result As String
    Dim objNetwork As Object
    
    result = "4. WSCRIPT.NETWORK" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    On Error Resume Next
    Set objNetwork = CreateObject("WScript.Network")
    
    If Not objNetwork Is Nothing Then
        result = result & "UserName: " & objNetwork.UserName & vbCrLf
        result = result & "UserDomain: " & objNetwork.UserDomain & vbCrLf
        result = result & "ComputerName: " & objNetwork.ComputerName & vbCrLf
        
        ' Lecteurs réseau mappés
        Dim objDrives As Object
        Set objDrives = objNetwork.EnumNetworkDrives
        result = result & "Lecteurs mappés: " & objDrives.Count / 2 & vbCrLf
    Else
        result = result & "WScript.Network non disponible" & vbCrLf
    End If
    On Error GoTo 0
    
    TestWScriptNetwork = result & vbCrLf
    Exit Function

ErrorHandler:
    TestWScriptNetwork = "Erreur lors du test WScript.Network."
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, "IdentityTester"
    HandleError "IdentityTester", PROC_NAME, "Erreur test WScript.Network"
End Function

' ========================================================================
' WMI (WINDOWS MANAGEMENT INSTRUMENTATION)
' ========================================================================
Function TestWMIMethods() As String
    Const PROC_NAME As String = "TestWMIMethods"
    On Error GoTo ErrorHandler
    Dim result As String
    result = "5. WMI (WINDOWS MANAGEMENT INSTRUMENTATION)" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    result = result & GetWMIComputerSystem() & vbCrLf
    result = result & GetWMIOperatingSystem() & vbCrLf
    result = result & GetWMILoggedOnUsers() & vbCrLf
    
    TestWMIMethods = result
    Exit Function
ErrorHandler:
    TestWMIMethods = "Erreur lors du test WMI."
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, "IdentityTester"
    HandleError "IdentityTester", PROC_NAME, "Erreur test WMI"
End Function

Function GetWMIComputerSystem() As String
    Const PROC_NAME As String = "GetWMIComputerSystem"
    On Error GoTo ErrorHandler
    
    Dim objWMI As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim result As String
    
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    
    For Each objItem In colItems
        result = result & "WMI User: " & objItem.UserName & vbCrLf
        result = result & "WMI Domain: " & objItem.Domain & vbCrLf
        result = result & "WMI Computer: " & objItem.Name & vbCrLf
        result = result & "WMI Workgroup: " & objItem.Workgroup & vbCrLf
        result = result & "WMI Manufacturer: " & objItem.Manufacturer & vbCrLf
        result = result & "WMI Model: " & objItem.Model & vbCrLf
        Exit For
    Next
    On Error GoTo 0
    
    Set colItems = Nothing
    Set objWMI = Nothing
    
    GetWMIComputerSystem = result
End Function

Function GetWMIOperatingSystem() As String
    Const PROC_NAME As String = "GetWMIOperatingSystem"
    On Error GoTo ErrorHandler
    
    Dim objWMI As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim result As String
    
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    
    For Each objItem In colItems
        result = result & "OS Caption: " & objItem.Caption & vbCrLf
        result = result & "OS Version: " & objItem.Version & vbCrLf
        result = result & "OS Organization: " & objItem.Organization & vbCrLf
        result = result & "OS Registered User: " & objItem.RegisteredUser & vbCrLf
        Exit For
    Next
    On Error GoTo 0
    
    Set colItems = Nothing
    Set objWMI = Nothing
    
    GetWMIOperatingSystem = result
End Function

Function GetWMILoggedOnUsers() As String
    Const PROC_NAME As String = "GetWMILoggedOnUsers"
    On Error GoTo ErrorHandler
    
    Dim objWMI As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim result As String
    Dim count As Integer
    
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_LoggedOnUser")
    
    count = 0
    For Each objItem In colItems
        count = count + 1
        If count <= 5 Then ' Limiter l'affichage
            result = result & "Session " & count & ": " & objItem.Antecedent & vbCrLf
        End If
    Next
    result = result & "Total sessions: " & count & vbCrLf
    On Error GoTo 0
    
    Set colItems = Nothing
    Set objWMI = Nothing
    
    GetWMILoggedOnUsers = result
End Function

' ========================================================================
' INFORMATIONS OFFICE
' ========================================================================
Function TestOfficeInfo() As String
    Dim result As String
    result = "6. INFORMATIONS OFFICE" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    On Error Resume Next
    ' Informations sur l'application courante
    result = result & "Application Name: " & Application.Name & vbCrLf
    result = result & "Application Version: " & Application.Version & vbCrLf
    result = result & "Application Build: " & Application.Build & vbCrLf
    result = result & "Application Path: " & Application.Path & vbCrLf
    
    ' Tentative d'accès aux autres applications Office
    Dim objWord As Object, objOutlook As Object, objPowerPoint As Object
    
    Set objWord = CreateObject("Word.Application")
    If Not objWord Is Nothing Then
        result = result & "Word UserName: " & objWord.UserName & vbCrLf
        objWord.Quit
    End If
    
    Set objOutlook = CreateObject("Outlook.Application")
    If Not objOutlook Is Nothing Then
        result = result & "Outlook disponible: Oui" & vbCrLf
        objOutlook.Quit
    End If
    
    On Error GoTo 0
    
    TestOfficeInfo = result & vbCrLf
End Function

' ========================================================================
' INFORMATIONS DU REGISTRE
' ========================================================================
Function TestRegistryInfo() As String
    Dim result As String
    Dim objShell As Object
    
    result = "7. INFORMATIONS DU REGISTRE" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    On Error Resume Next
    Set objShell = CreateObject("WScript.Shell")
    
    If Not objShell Is Nothing Then
        result = result & "Registered Owner: " & objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\RegisteredOwner") & vbCrLf
        result = result & "Registered Organization: " & objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\RegisteredOrganization") & vbCrLf
        result = result & "Product Name: " & objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName") & vbCrLf
        result = result & "Current Build: " & objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuild") & vbCrLf
        result = result & "Install Date: " & objShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\InstallDate") & vbCrLf
    Else
        result = result & "Accès au registre non disponible" & vbCrLf
    End If
    On Error GoTo 0
    
    TestRegistryInfo = result & vbCrLf
End Function

' ========================================================================
' ACTIVE DIRECTORY
' ========================================================================
Function TestActiveDirectory() As String
    Dim result As String
    Dim objADSysInfo As Object
    Dim objUser As Object
    
    result = "8. ACTIVE DIRECTORY / ENTRA ID" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    On Error Resume Next
    Set objADSysInfo = CreateObject("ADSystemInfo")
    
    If Not objADSysInfo Is Nothing Then
        result = result & "User DN: " & objADSysInfo.UserName & vbCrLf
        result = result & "Computer DN: " & objADSysInfo.ComputerName & vbCrLf
        result = result & "Domain Short: " & objADSysInfo.DomainShortName & vbCrLf
        result = result & "Domain DNS: " & objADSysInfo.DomainDNSName & vbCrLf
        result = result & "Forest DNS: " & objADSysInfo.ForestDNSName & vbCrLf
        result = result & "Site Name: " & objADSysInfo.SiteName & vbCrLf
        
        ' Tentative de récupération des détails utilisateur
        Set objUser = GetObject("LDAP://" & objADSysInfo.UserName)
        If Not objUser Is Nothing Then
            result = result & "Display Name: " & objUser.displayName & vbCrLf
            result = result & "Email: " & objUser.mail & vbCrLf
            result = result & "Title: " & objUser.title & vbCrLf
            result = result & "Department: " & objUser.department & vbCrLf
            result = result & "Company: " & objUser.company & vbCrLf
            result = result & "Phone: " & objUser.telephoneNumber & vbCrLf
        End If
    Else
        result = result & "ADSystemInfo non disponible (pas de domaine AD?)" & vbCrLf
    End If
    On Error GoTo 0
    
    TestActiveDirectory = result & vbCrLf
End Function

' ========================================================================
' INFORMATIONS ÉTENDUES
' ========================================================================
Function TestExtendedUserInfo() As String
    Const PROC_NAME As String = "TestExtendedUserInfo"
    On Error GoTo ErrorHandler
    Dim result As String
    result = "9. INFORMATIONS UTILISATEUR ÉTENDUES" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    ' Informations sur les groupes locaux
    result = result & GetLocalGroups() & vbCrLf
    
    ' Informations sur la session
    result = result & GetSessionInfo() & vbCrLf
    
    result = result & "SID du compte: " & GetUserSid() & vbCrLf
    
    TestExtendedUserInfo = result & vbCrLf
    Exit Function
ErrorHandler:
    TestExtendedUserInfo = "Erreur lors du test des informations étendues."
    HandleError "IdentityTester", PROC_NAME
End Function

Function GetLocalGroups() As String
    Const PROC_NAME As String = "GetLocalGroups"
    On Error GoTo ErrorHandler
    
    Dim username As String
    Dim bufPtr As LongPtr
    Dim count As Long
    Dim result As String
    
    username = Environ("USERNAME")
    If username = "" Then
        GetLocalGroups = "Erreur: USERNAME n'est pas défini"
        Exit Function
    End If
    
    On Error Resume Next
    If NetUserGetLocalGroups("", username, 2, 0, bufPtr, 0, count, 0) = 0 Then
        result = "Groupes locaux (échantillon):" & vbCrLf
        If count > 0 Then
            Dim i As Long
            Dim groupName As String
            For i = 0 To count - 1
                If NetApiBufferFree(bufPtr + i * 256) = 0 Then
                    groupName = Space(255)
                    If NetUserGetInfo("", username, 2, groupName) = 0 Then
                        result = result & "  " & Left$(groupName, InStr(groupName, Chr$(0)) - 1) & vbCrLf
                    End If
                End If
            Next
        End If
    Else
        result = "Erreur lors de la récupération des groupes locaux."
    End If
    On Error GoTo 0
    
    GetLocalGroups = result
End Function

Function GetSessionInfo() As String
    Const PROC_NAME As String = "GetSessionInfo"
    On Error GoTo ErrorHandler
    
    Dim result As String
    result = "Sessions interactives:" & vbCrLf
    result = result & "  Logon ID: " & Environ("LOGONID") & vbCrLf
    result = result & "  Start Time: " & Environ("LOGONSERVER") & vbCrLf
    result = result & "  Session Name: " & Environ("SESSIONNAME") & vbCrLf
    GetSessionInfo = result & vbCrLf
    Exit Function
    
ErrorHandler:
    GetSessionInfo = "  Erreur lors de la récupération des infos de session."
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, "IdentityTester"
    HandleError "IdentityTester", PROC_NAME, "Erreur GetSessionInfo"
End Function

' ========================================================================
' FONCTION UTILITAIRE - Sauvegarde du rapport
' ========================================================================
Sub SaveReportToFile(reportContent As String)
    Const PROC_NAME As String = "SaveReportToFile"
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    Dim fileNum As Integer
    
    filePath = Environ("USERPROFILE") & "\Desktop\IdentityReport_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    fileNum = FreeFile
    
    On Error Resume Next
    Open filePath For Output As #fileNum
    Print #fileNum, reportContent
    Close #fileNum
    
    If Err.Number = 0 Then
        Debug.Print "Rapport sauvegardé: " & filePath
    Else
        Debug.Print "Erreur sauvegarde: " & Err.Description
    End If
    On Error GoTo 0
End Sub

' ========================================================================
' FONCTIONS DE TEST INDIVIDUELLES
' ========================================================================

' Test rapide d'une méthode spécifique
Sub QuickTestUserName()
    Const PROC_NAME As String = "QuickTestUserName"
    On Error GoTo ErrorHandler
    MsgBox "Application.UserName: " & Application.UserName & vbCrLf & _
           "Environ(""USERNAME""): " & Environ("USERNAME") & vbCrLf & _
           "API GetUserName: " & GetCurrentUserAPI(), _
           vbInformation, "Test rapide: Noms d'utilisateur"
    Exit Sub
ErrorHandler:
    HandleError "IdentityTester", PROC_NAME, "Erreur QuickTestUserName"
End Sub

Sub QuickTestEnvironment()
    Const PROC_NAME As String = "QuickTestEnvironment"
    On Error GoTo ErrorHandler
    MsgBox "USERDOMAIN: " & Environ("USERDOMAIN") & vbCrLf & _
           "COMPUTERNAME: " & Environ("COMPUTERNAME") & vbCrLf & _
           "LOGONSERVER: " & Environ("LOGONSERVER"), _
           vbInformation, "Test rapide: Environnement"
    Exit Sub
ErrorHandler:
    HandleError "IdentityTester", PROC_NAME, "Erreur QuickTestEnvironment"
End Sub

Sub QuickTestAPI()
    Const PROC_NAME As String = "QuickTestAPI"
    On Error GoTo ErrorHandler
    MsgBox "GetUserNameEx (UPN): " & GetUserNameExtended(NAME_USER_PRINCIPAL) & vbCrLf & _
           "GetComputerName API: " & GetComputerNameAPI(), _
           vbInformation, "Test rapide: API"
    Exit Sub
ErrorHandler:
    HandleError "IdentityTester", PROC_NAME, "Erreur QuickTestAPI"
End Sub

Sub QuickTestWMI()
    Const PROC_NAME As String = "QuickTestWMI"
    On Error GoTo ErrorHandler
    Dim wmiService As Object
    Dim compSys As Object
    Set wmiService = GetObject("winmgmts:\\.\root\cimv2")
    Set compSys = wmiService.Get("Win32_ComputerSystem")
    MsgBox "WMI ComputerSystem.UserName: " & compSys.UserName, vbInformation, "Test rapide: WMI"
    Exit Sub
ErrorHandler:
    HandleError "IdentityTester", PROC_NAME, "Erreur QuickTestWMI"
End Sub

' ========================================================================
' FONCTION DE DIAGNOSTIC PRINCIPALE
' ========================================================================
Sub DiagnoseIssues()
    Const PROC_NAME As String = "DiagnoseIssues"
    Const MODULE_NAME As String = "IdentityTester"
    On Error GoTo ErrorHandler
    
    Dim output As String
    Dim domain As String
    
    ' Vérifier les permissions
    On Error Resume Next
    Dim objWMI As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    If objWMI Is Nothing Then
        output = output & "⚠️ WMI non accessible"
    Else
        output = output & "✅ WMI accessible"
    End If
    
    ' Vérifier ADSystemInfo
    Dim objAD As Object
    Set objAD = CreateObject("ADSystemInfo")
    If objAD Is Nothing Then
        output = output & vbCrLf & "⚠️ ADSystemInfo non accessible (normal si pas de domaine)"
    Else
        output = output & vbCrLf & "✅ ADSystemInfo accessible"
    End If
    
    ' Vérifier WScript
    Dim objWScript As Object
    Set objWScript = CreateObject("WScript.Network")
    If objWScript Is Nothing Then
        output = output & vbCrLf & "⚠️ WScript.Network non accessible"
    Else
        output = output & vbCrLf & "✅ WScript.Network accessible"
    End If
    
    ' Vérifier le domaine
    domain = objAD.Domain
    If domain = "" Then
        output = output & vbCrLf & "Le domaine n'est pas défini, l'utilisateur est probablement local."
    End If
    
    MsgBox output, vbInformation, "Rapport de Diagnostic"
    Exit Sub
    
ErrorHandler:
    MsgBox "Une erreur fatale est survenue durant le diagnostic: " & Err.Description, vbCritical
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur DiagnoseIssues"
End Sub

' --- Test de la clé API RAGIC ---
Sub TestRagicApiKey()
    Const PROC_NAME As String = "TestRagicApiKey"
    Const MODULE_NAME As String = "IdentityTester"
    On Error GoTo ErrorHandler
    
    Dim key As String
    Dim msg As String
    Dim startTime As Double
    Dim endTime As Double
    
    ' Test 1: Vérifier que la clé API peut être récupérée
    startTime = Timer
    key = env.GetRagicApiKey()
    endTime = Timer
    msg = "Test 1 - Clé API récupérée : " & IIf(Len(key) > 0, "OK", "ÉCHEC")
    
    ' Test 2: Vérifier que les paramètres d'API sont correctement formés
    Dim apiParams As String
    apiParams = env.GetRagicApiParams()
    msg = msg & vbCrLf & "Test 2 - Paramètres API formés : " & IIf(InStr(apiParams, "APIKey=") > 0, "OK", "ÉCHEC")
    
    ' Test 3: Tester une requête réelle vers Ragic
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    Dim testUrl As String
    testUrl = env.RAGIC_BASE_URL & "1" & env.GetRagicApiParams()
    
    http.Open "GET", testUrl, False
    http.send
    
    msg = msg & vbCrLf & "Test 3 - Requête API : " & IIf(http.Status = 200, "OK", "ÉCHEC (Status " & http.Status & ")")
    
    ' Test 4: Vérifier que CategoryManager peut toujours construire des URLs
    Dim categories() As CategoryInfo
    categories = CategoryManager.GetAllCategories()
    msg = msg & vbCrLf & "Test 4 - URLs des catégories : " & IIf(UBound(categories) >= 0, "OK", "ÉCHEC")
    
    ' Test 5: Vérifier que le logging fonctionne toujours
    Log "test_api", "Test de l'API Ragic", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    msg = msg & vbCrLf & "Test 5 - Logging : OK"
    
    msg = msg & vbCrLf & "Temps de récupération: " & Format(endTime - startTime, "0.000s")
    
    MsgBox msg, IIf(InStr(msg, "ERREUR") > 0, vbCritical, vbInformation), "Test Clé API Ragic"
    Exit Sub
    
ErrorHandler:
    MsgBox "Erreur fatale lors du test de la clé API: " & Err.Description, vbCritical
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur TestRagicApiKey"
End Sub

' --- Test des formats de noms ---

Private Function GetUserByFormat(formatCode As Long) As String
    Const PROC_NAME As String = "GetUserByFormat"
    Const MODULE_NAME As String = "IdentityTester"
    On Error GoTo ErrorHandler
    Dim s As String
    s = Space(255)
    Dim size As Long
    Dim success As Boolean
    
    size = 255
    success = GetUserNameEx(formatCode, s, size)
    
    If success Then
        GetUserByFormat = Left(s, size)
    Else
        GetUserByFormat = "N/A"
    End If
    Exit Function
ErrorHandler:
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur GetUserByFormat"
    GetUserByFormat = "ERREUR"
End Function

Sub TestGetUserNameFormats()
    Const PROC_NAME As String = "TestGetUserNameFormats"
    Const MODULE_NAME As String = "IdentityTester"
    On Error GoTo ErrorHandler
    Dim msg As String
    msg = "Test des différents formats de GetUserNameEx:" & vbCrLf & vbCrLf
    
    msg = msg & "NameUnknown: " & GetUserByFormat(NAME_UNKNOWN) & vbCrLf
    msg = msg & "NameFullyQualifiedDN: " & GetUserByFormat(NAME_FULLY_QUALIFIED_DN) & vbCrLf
    msg = msg & "NameSamCompatible: " & GetUserByFormat(NAME_SAM_COMPATIBLE) & vbCrLf
    msg = msg & "NameDisplay: " & GetUserByFormat(NAME_DISPLAY) & vbCrLf
    msg = msg & "NameUniqueId: " & GetUserByFormat(NAME_UNIQUE_ID) & vbCrLf
    msg = msg & "NameCanonical: " & GetUserByFormat(NAME_CANONICAL) & vbCrLf
    msg = msg & "NameUserPrincipal: " & GetUserByFormat(NAME_USER_PRINCIPAL) & vbCrLf
    
    MsgBox msg, vbInformation, "Test GetUserNameEx"
    Exit Sub
ErrorHandler:
    SYS_Logger.Log "identity_test_error", "Erreur VBA dans " & PROC_NAME & " - Numéro: " & CStr(Err.Number) & ", Description: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    HandleError MODULE_NAME, PROC_NAME, "Erreur TestGetUserNameFormats"
End Sub