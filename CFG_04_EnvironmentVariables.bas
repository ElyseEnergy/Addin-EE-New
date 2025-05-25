' Module: ENV
' Ce module contient uniquement les variables d'environnement sensibles
Option Explicit
Private Const MODULE_NAME As String = "ENV"

' Constantes pour l'API Ragic
Public Const RAGIC_BASE_URL As String = "https://ragic.elyse.energy/default/"
Public Const RAGIC_API_KEY As String = "Njl3OENtYnFnTExxSzNWVXZ6Y2E1Tlg0RWtjcVVBdnFHeVR0cTRCS09OWDMwZHlqRVc3WGx3WFJTNTFXMDRDZlZ2OWdXVElUaEtnPQ==&"
Public Const RAGIC_API_PARAMS As String = "?APIKey=" & RAGIC_API_KEY & "&f=all"

Public Function GetEnvironmentVariable(ByVal varName As String) As String
    Const PROC_NAME As String = "GetEnvironmentVariable"
    On Error GoTo ErrorHandler

    ElyseMain_Orchestrator.LogDebug PROC_NAME & "_Start", "Fetching environment variable: " & varName, PROC_NAME, MODULE_NAME
    
    Dim varValue As String
    varValue = Environ(varName)
    
    If varValue = "" Then
        ElyseMain_Orchestrator.LogWarning PROC_NAME & "_NotFound", "Environment variable '" & varName & "' not found or is empty.", PROC_NAME, MODULE_NAME
        ' MsgBox "Environment variable '" & varName & "' not found.", vbInformation, "Environment Variable"
        ElyseMessageBox_System.ShowInfoMessage "Environment Variable", "Environment variable '" & varName & "' not found or is empty."
    Else
        ElyseMain_Orchestrator.LogDebug PROC_NAME & "_Found", "Environment variable '" & varName & "' found with value (length): " & Len(varValue), PROC_NAME, MODULE_NAME
    End If
    
    GetEnvironmentVariable = varValue
    Exit Function

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
    GetEnvironmentVariable = "" ' Default error return
End Function

Public Sub SetCustomEnvironmentVariable(ByVal varName As String, ByVal varValue As String)
    Const PROC_NAME As String = "SetCustomEnvironmentVariable"
    On Error GoTo ErrorHandler
    
    ElyseMain_Orchestrator.LogInfo PROC_NAME & "_Start", "Setting custom environment variable: " & varName, PROC_NAME, MODULE_NAME
    
    ' This is a conceptual example, as VBA cannot directly set persistent environment variables for the system or other processes.
    ' Typically, this would involve registry edits or calling external processes, which is complex and has security implications.
    ' For the scope of this add-in, we might simulate it with a global variable or a temporary setting.
    ' Let's assume it logs the intent and perhaps stores it in a global collection within the VBA environment.
    
    ' Original: Debug.Print "Attempting to set " & varName & " to " & varValue
    ElyseMain_Orchestrator.LogDebug PROC_NAME & "_Attempt", "Attempting to set (simulated) " & varName & " to " & varValue, PROC_NAME, MODULE_NAME
    
    ' Placeholder for actual logic if it were possible/safe
    ' For now, just log it.
    If varName <> "" Then
        ElyseMain_Orchestrator.LogInfo PROC_NAME & "_SetSuccess", "Custom variable '" & varName & "' conceptually set to '" & varValue & "' within application session.", PROC_NAME, MODULE_NAME
    Else
        ElyseMain_Orchestrator.LogWarning PROC_NAME & "_InvalidName", "Invalid variable name provided for custom environment variable.", PROC_NAME, MODULE_NAME
        ElyseMessageBox_System.ShowWarningMessage "Set Variable", "Cannot set a custom variable with an empty name."
    End If
    
    Exit Sub

ErrorHandler:
    ElyseMain_Orchestrator.HandleError MODULE_NAME, PROC_NAME
End Sub

