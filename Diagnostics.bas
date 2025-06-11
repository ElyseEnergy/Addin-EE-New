Attribute VB_Name = "Diagnostics"
Option Explicit

Private startTime As Double
Private lastTime As Double
Private Const LOG_SHEET_NAME As String = "App_Logs"

' D�marre le chronom�tre global pour un processus donn�
Public Sub StartTimer(processName As String)
    startTime = Timer
    lastTime = startTime
    Log "PERF_LOG", "--- D�BUT: " & processName & " ---", INFO_LEVEL, "StartTimer", "Diagnostics"
End Sub

' Arr�te le chronom�tre global et loggue la dur�e totale
Public Sub StopTimer(processName As String)
    ' On loggue une derni�re �tape avant le calcul total
    LogTime "FIN DU CODE VBA pour " & processName
    
    Dim currentTime As Double
    currentTime = Timer
    Dim elapsedTotal As Double
    
    If currentTime < startTime Then
        elapsedTotal = (86400 - startTime) + currentTime
    Else
        elapsedTotal = currentTime - startTime
    End If
    
    Log "PERF_LOG", "--- FIN: " & processName & ". Dur�e totale du code VBA: " & Format(elapsedTotal, "0.000s") & " ---", INFO_LEVEL, "StopTimer", "Diagnostics"
End Sub

' Loggue le temps �coul� pour une �tape sp�cifique
Public Sub LogTime(stepName As String)
    Dim currentTime As Double
    currentTime = Timer

    Dim elapsedTotal As Double
    Dim elapsedStep As Double
    
    ' G�rer le passage de minuit
    If currentTime < startTime Then elapsedTotal = (86400 - startTime) + currentTime Else elapsedTotal = currentTime - startTime
    If currentTime < lastTime Then elapsedStep = (86400 - lastTime) + currentTime Else elapsedStep = currentTime - lastTime

    Dim logMessage As String
    logMessage = "TIMER | " & stepName & _
                   " | �tape: " & Format(elapsedStep, "0.000s") & _
                   " | Total: " & Format(elapsedTotal, "0.000s")
    
    ' Utiliser le syst�me de logging existant
    Log "PERF_LOG", logMessage, INFO_LEVEL, "LogTime", "Diagnostics"
    
    lastTime = currentTime
End Sub

' Attend la fin des calculs Excel et loggue le temps d'attente
Public Sub WaitAndLogCalculation()
    LogTime "Avant attente des calculs Excel"
    
    Dim calcStartTime As Double
    calcStartTime = Timer
    
    Do While Application.CalculationState <> xlDone
        DoEvents
    Loop
    
    Dim calcEndTime As Double
    calcEndTime = Timer
    
    Dim elapsedCalc As Double
    If calcEndTime < calcStartTime Then elapsedCalc = (86400 - calcStartTime) + calcEndTime Else elapsedCalc = calcEndTime - calcStartTime
    
    Log "PERF_LOG", "Temps de recalcul/rendu Excel: " & Format(elapsedCalc, "0.000s"), INFO_LEVEL, "WaitAndLogCalculation", "Diagnostics"
    
    LogTime "FIN TOTALE (main r�ellement rendue)"
End Sub

' Routine de log qui �crit dans une feuille de calcul
Public Sub LogToSheet(message As String, level As LogLevel, procedureName As String, moduleName As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim nextRow As Long
    
    ' D�sactiver la mise � jour pour acc�l�rer
    Dim initialScreenUpdating As Boolean
    initialScreenUpdating = Application.ScreenUpdating
    If initialScreenUpdating Then Application.ScreenUpdating = False
    
    ' Essayer de r�cup�rer la feuille de log, la cr�er si elle n'existe pas
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo ErrorHandler ' R�activer la gestion d'erreur normale
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = LOG_SHEET_NAME
        ws.Tab.Color = vbRed
        ' Cr�er les en-t�tes
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Level"
        ws.Cells(1, 3).Value = "Module"
        ws.Cells(1, 4).Value = "Procedure"
        ws.Cells(1, 5).Value = "Message"
        ws.Range("A1:E1").Font.Bold = True
    End If
    
    ' Trouver la prochaine ligne vide
    nextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    
    ' �crire les informations de log
    ws.Cells(nextRow, 1).Value = Now()
    ws.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(nextRow, 2).Value = LogLevelToString(level)
    ws.Cells(nextRow, 3).Value = moduleName
    ws.Cells(nextRow, 4).Value = procedureName
    ws.Cells(nextRow, 5).Value = message
    
    ' Ajuster la largeur des colonnes si c'est le premier log
    If nextRow = 2 Then ws.Columns("A:E").AutoFit
    
    ' R�activer la mise � jour de l'�cran si elle l'�tait au d�part
    If initialScreenUpdating Then Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    ' En cas d'erreur de logging (par ex: feuille prot�g�e), on ne fait rien pour ne pas planter le process principal
    Debug.Print "ERREUR DE LOGGING: " & Err.Description
    If initialScreenUpdating Then Application.ScreenUpdating = True
End Sub

' Fonction utilitaire pour convertir le niveau de log en string
Private Function LogLevelToString(level As LogLevel) As String
    Select Case level
        Case 0: LogLevelToString = "DEBUG"
        Case 1: LogLevelToString = "INFO"
        Case 2: LogLevelToString = "WARNING"
        Case 3: LogLevelToString = "ERROR"
        Case Else: LogLevelToString = "UNKNOWN"
    End Select
End Function
