Attribute VB_Name = "IdentityTester"
'========================================================================
' Module: IdentityTester
' Description: Module complet pour tester toutes les méthodes de récupération
'              d'informations d'identité utilisateur sous VBA
' Contexte: Windows connecté à Entra ID, PC enrollé, sans droits admin
' ========================================================================

Option Explicit

' Déclarations API Windows
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare PtrSafe Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" _
    (ByVal NameFormat As Long, ByVal lpNameBuffer As String, ByRef nSize As Long) As Long

' Constants pour GetUserNameEx
Private Const NameUnknown As Long = 0
Private Const NameFullyQualifiedDN As Long = 1
Private Const NameSamCompatible As Long = 2
Private Const NameDisplay As Long = 3
Private Const NameUniqueId As Long = 6
Private Const NameCanonical As Long = 7
Private Const NameUserPrincipal As Long = 8

' ========================================================================
' FONCTION PRINCIPALE - Lance tous les tests
' ========================================================================
Sub TestAllIdentityMethods()
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
    ' SaveReportToFile output
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
End Function

' ========================================================================
' API WINDOWS
' ========================================================================
Function TestWindowsAPI() As String
    Dim result As String
    result = "3. API WINDOWS" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    result = result & "GetUserName API: " & GetCurrentUserAPI() & vbCrLf
    result = result & "GetComputerName API: " & GetComputerNameAPI() & vbCrLf
    result = result & "GetUserNameEx (Display): " & GetUserNameExtended(NameDisplay) & vbCrLf
    result = result & "GetUserNameEx (SAM): " & GetUserNameExtended(NameSamCompatible) & vbCrLf
    result = result & "GetUserNameEx (UPN): " & GetUserNameExtended(NameUserPrincipal) & vbCrLf
    result = result & "GetUserNameEx (DN): " & GetUserNameExtended(NameFullyQualifiedDN) & vbCrLf
    
    TestWindowsAPI = result & vbCrLf
End Function

Function GetCurrentUserAPI() As String
    Dim buffer As String * 255
    Dim ret As Long
    ret = GetUserName(buffer, 255)
    If ret <> 0 Then
        GetCurrentUserAPI = Left$(buffer, InStr(buffer, Chr$(0)) - 1)
    Else
        GetCurrentUserAPI = "Erreur API"
    End If
End Function

Function GetComputerNameAPI() As String
    Dim buffer As String * 255
    Dim ret As Long
    ret = GetComputerName(buffer, 255)
    If ret <> 0 Then
        GetComputerNameAPI = Left$(buffer, InStr(buffer, Chr$(0)) - 1)
    Else
        GetComputerNameAPI = "Erreur API"
    End If
End Function

Function GetUserNameExtended(NameFormat As Long) As String
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
End Function

' ========================================================================
' WSCRIPT.NETWORK
' ========================================================================
Function TestWScriptNetwork() As String
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
End Function

' ========================================================================
' WMI (WINDOWS MANAGEMENT INSTRUMENTATION)
' ========================================================================
Function TestWMIMethods() As String
    Dim result As String
    result = "5. WMI (WINDOWS MANAGEMENT INSTRUMENTATION)" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    result = result & GetWMIComputerSystem() & vbCrLf
    result = result & GetWMIOperatingSystem() & vbCrLf
    result = result & GetWMILoggedOnUsers() & vbCrLf
    
    TestWMIMethods = result
End Function

Function GetWMIComputerSystem() As String
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
    
    GetWMIComputerSystem = result
End Function

Function GetWMIOperatingSystem() As String
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
    
    GetWMIOperatingSystem = result
End Function

Function GetWMILoggedOnUsers() As String
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
    Dim result As String
    result = "9. INFORMATIONS ÉTENDUES" & vbCrLf
    result = result & String(50, "-") & vbCrLf
    
    ' Informations sur les groupes locaux
    result = result & GetLocalGroups() & vbCrLf
    
    ' Informations sur la session
    result = result & GetSessionInfo() & vbCrLf
    
    TestExtendedUserInfo = result
End Function

Function GetLocalGroups() As String
    Dim objWMI As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim result As String
    Dim count As Integer
    
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_GroupUser WHERE GroupComponent LIKE '%Administrators%'")
    
    result = "Groupes locaux (échantillon):" & vbCrLf
    count = 0
    For Each objItem In colItems
        count = count + 1
        If count <= 3 Then
            result = result & "  " & objItem.PartComponent & vbCrLf
        End If
    Next
    On Error GoTo 0
    
    GetLocalGroups = result
End Function

Function GetSessionInfo() As String
    Dim objWMI As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim result As String
    
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("SELECT * FROM Win32_LogonSession WHERE LogonType=2")
    
    result = "Sessions interactives:" & vbCrLf
    For Each objItem In colItems
        result = result & "  Logon ID: " & objItem.LogonId & vbCrLf
        result = result & "  Start Time: " & objItem.StartTime & vbCrLf
        Exit For ' Prendre seulement la première
    Next
    On Error GoTo 0
    
    GetSessionInfo = result
End Function

' ========================================================================
' FONCTION UTILITAIRE - Sauvegarde du rapport
' ========================================================================
Sub SaveReportToFile(reportContent As String)
    Dim fileName As String
    Dim fileNum As Integer
    
    fileName = Environ("USERPROFILE") & "\Desktop\IdentityReport_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    fileNum = FreeFile
    
    On Error Resume Next
    Open fileName For Output As #fileNum
    Print #fileNum, reportContent
    Close #fileNum
    
    If Err.Number = 0 Then
        Debug.Print "Rapport sauvegardé: " & fileName
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
    Debug.Print "Test rapide - Application.UserName: " & Application.UserName
End Sub

Sub QuickTestEnvironment()
    Debug.Print "Test rapide - USERNAME: " & Environ("USERNAME")
    Debug.Print "Test rapide - USERDOMAIN: " & Environ("USERDOMAIN")
End Sub

Sub QuickTestAPI()
    Debug.Print "Test rapide - API GetUserName: " & GetCurrentUserAPI()
End Sub

Sub QuickTestWMI()
    Debug.Print "Test rapide - WMI: " & GetWMIComputerSystem()
End Sub

' ========================================================================
' FONCTION DE DIAGNOSTIC
' ========================================================================
Sub DiagnoseIssues()
    Debug.Print "=== DIAGNOSTIC DES PROBLÈMES POTENTIELS ==="
    
    ' Vérifier les permissions
    On Error Resume Next
    Dim objWMI As Object
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    If objWMI Is Nothing Then
        Debug.Print "⚠️ WMI non accessible"
    Else
        Debug.Print "✅ WMI accessible"
    End If
    
    ' Vérifier ADSystemInfo
    Dim objAD As Object
    Set objAD = CreateObject("ADSystemInfo")
    If objAD Is Nothing Then
        Debug.Print "⚠️ ADSystemInfo non accessible (normal si pas de domaine)"
    Else
        Debug.Print "✅ ADSystemInfo accessible"
    End If
    
    ' Vérifier WScript
    Dim objWScript As Object
    Set objWScript = CreateObject("WScript.Network")
    If objWScript Is Nothing Then
        Debug.Print "⚠️ WScript.Network non accessible"
    Else
        Debug.Print "✅ WScript.Network accessible"
    End If
    
    On Error GoTo 0
    Debug.Print "=== FIN DU DIAGNOSTIC ==="
End Sub

' --- Constantes et Enums du module ---
Private Const MODULE_NAME As String = "IdentityTester"

' --- Procédures publiques ---
Public Sub TestRagicApiKey()
    Const PROC_NAME As String = "TestRagicApiKey"
    On Error GoTo ErrorHandler
    
    ' Test 1: Vérifier que la clé API peut être récupérée
    Dim apiKey As String
    apiKey = env.GetRagicApiKey()
    Debug.Print "Test 1 - Clé API récupérée : " & IIf(Len(apiKey) > 0, "OK", "ÉCHEC")
    
    ' Test 2: Vérifier que les paramètres d'API sont correctement formés
    Dim apiParams As String
    apiParams = env.GetRagicApiParams()
    Debug.Print "Test 2 - Paramètres API formés : " & IIf(InStr(apiParams, "APIKey=") > 0, "OK", "ÉCHEC")
    
    ' Test 3: Tester une requête réelle vers Ragic
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    Dim testUrl As String
    testUrl = env.RAGIC_BASE_URL & "1" & env.GetRagicApiParams()
    
    http.Open "GET", testUrl, False
    http.send
    
    Debug.Print "Test 3 - Requête API : " & IIf(http.Status = 200, "OK", "ÉCHEC (Status " & http.Status & ")")
    
    ' Test 4: Vérifier que CategoryManager peut toujours construire des URLs
    Dim categories() As CategoryInfo
    categories = CategoryManager.GetAllCategories()
    Debug.Print "Test 4 - URLs des catégories : " & IIf(UBound(categories) >= 0, "OK", "ÉCHEC")
    
    ' Test 5: Vérifier que le logging fonctionne toujours
    Log "test_api", "Test de l'API Ragic", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Debug.Print "Test 5 - Logging : OK"
    
    MsgBox "Tests terminés. Consultez la fenêtre Immediate pour les résultats.", vbInformation
    Exit Sub

ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors des tests de l'API : " & Err.Description
End Sub