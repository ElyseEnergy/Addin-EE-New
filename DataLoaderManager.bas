Attribute VB_Name = "DataLoaderManager"
' REPLACE the top of DataLoaderManager.bas with this:

' At the top of DataLoaderManager module, add this:
Option Explicit
Private currentProgressForm As OutputRangeSelectionForm
Private m_userCancelled As Boolean  ' Module-level variable to track cancellation


' SIMPLEST FIX: Just don't show error messages when ProcessCategory returns False
' Since the forms already show appropriate messages to the user

Public Function ProcessCategory(categoryName As String, Optional errorMessage As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "ProcessCategory"
    Const MODULE_NAME As String = "DataLoaderManager"
    
    Log "dataloader", "D�but ProcessCategory | Cat�gorie demand�e: " & categoryName, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    ' Initialize categories if needed
    If CategoriesCount = 0 Then InitCategories
    
    ' Get category info
    Dim categoryInfo As categoryInfo
    categoryInfo = GetCategoryByName(categoryName)
    
    If categoryInfo.DisplayName = "" Then
        MsgBox "Cat�gorie '" & categoryName & "' non trouv�e", vbExclamation
        Log "dataloader", "ERREUR: Cat�gorie non trouv�e: " & categoryName, ERROR_LEVEL, PROC_NAME, MODULE_NAME
        ProcessCategory = False
        Exit Function
    End If
    
    ' Process the category with simplified approach
    Dim result As Boolean
    result = ProcessCategorySimplified(categoryInfo)
    
    ProcessCategory = result
    
    ' REMOVED: Error message display when result is False
    ' The forms already handle user feedback appropriately
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE_NAME, PROC_NAME, "Erreur lors du traitement de la cat�gorie " & categoryName & ": " & Err.Description
    ' Only show error message for actual errors, not for False results
    ProcessCategory = False
End Function

' UPDATED: ProcessCategorySimplified without the userCancelled parameter
Private Function ProcessCategorySimplified(categoryInfo As categoryInfo) As Boolean
    On Error GoTo ErrorHandler
    Dim lastCol As Long
    
    Log "debug", "=== ProcessCategorySimplified START ===", DEBUG_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
    
    ' Step 1: Ensure PQ_DATA sheet exists in the active workbook
    Set wsPQData = GetOrCreatePQDataSheet(ActiveWorkbook)
    If wsPQData Is Nothing Then
        Log "debug", "ERROR: GetOrCreatePQDataSheet failed", ERROR_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
        MsgBox "Erreur lors de l'initialisation de la feuille PQ_DATA", vbExclamation
        ProcessCategorySimplified = False
        Exit Function
    End If
    Log "debug", "PQ_DATA sheet OK: " & wsPQData.Name, DEBUG_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
    
    ' Step 2: Ensure PowerQuery exists
    If Not PQQueryManager.EnsurePQQueryExists(categoryInfo) Then
        Log "debug", "ERROR: EnsurePQQueryExists failed", ERROR_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
        MsgBox "Erreur lors de la cr�ation de la requ�te PowerQuery", vbExclamation
        ProcessCategorySimplified = False
        Exit Function
    End If
    Log "debug", "PowerQuery ensured successfully", DEBUG_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
    
    ' Step 3: Set status and load data
    Application.Cursor = xlWait
    Application.StatusBar = "T�l�chargement des donn�es pour '" & categoryInfo.DisplayName & "' en cours..."
    DoEvents
    
    ' Ensure table is loaded in the correct workbook
    Dim tableName As String
    tableName = "Table_" & Utilities.SanitizeTableName(categoryInfo.PowerQueryName)
    
    Dim lo As ListObject
    On Error Resume Next
    ' Utilise wsPQData qui pointe maintenant vers la feuille de l'addin
    Set lo = wsPQData.ListObjects(tableName)
    On Error GoTo ErrorHandler
    
    If lo Is Nothing Then
        Log "debug", "Table not found, loading query...", DEBUG_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
        
        lastCol = Utilities.GetLastColumn(wsPQData)
        LoadQueries.LoadQuery categoryInfo.PowerQueryName, wsPQData, wsPQData.Cells(1, lastCol + 1)
        
        On Error Resume Next
        Set lo = wsPQData.ListObjects(tableName)
        On Error GoTo ErrorHandler
        
        If lo Is Nothing Then
            Log "debug", "ERROR: Table still not found after LoadQuery", ERROR_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
            Application.Cursor = xlDefault
            Application.StatusBar = False
            MsgBox "Impossible de charger la table PowerQuery '" & categoryInfo.PowerQueryName & "' dans PQ_DATA.", vbExclamation
            ProcessCategorySimplified = False
            Exit Function
        End If
    Else
        ' Table exists, but check if it has data
        If lo.DataBodyRange Is Nothing Then
            Log "debug", "Table exists but is empty, forcing refresh...", WARNING_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
            
            ' Force refresh the table
            On Error Resume Next
            lo.QueryTable.Refresh BackgroundQuery:=False
            On Error GoTo ErrorHandler
            
            ' Double-check after refresh
            If lo.DataBodyRange Is Nothing Then
                Log "debug", "Table still empty after refresh, trying to reload completely...", WARNING_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
                
                ' Delete and recreate the table
                lo.Delete
                
                
                lastCol = Utilities.GetLastColumn(wsPQData)
                LoadQueries.LoadQuery categoryInfo.PowerQueryName, wsPQData, wsPQData.Cells(1, lastCol + 1)
                
                On Error Resume Next
                Set lo = wsPQData.ListObjects(tableName)
                On Error GoTo ErrorHandler
                
                If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
                    Log "debug", "CRITICAL: Cannot load table data even after recreating", ERROR_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
                    Application.Cursor = xlDefault
                    Application.StatusBar = False
                    MsgBox "Impossible de charger les données de '" & categoryInfo.PowerQueryName & "'" & vbCrLf & _
                           "La requête PowerQuery fonctionne mais le tableau Excel reste vide.", vbExclamation
                    ProcessCategorySimplified = False
                    Exit Function
                End If
            End If
            Log "debug", "Table refresh successful", INFO_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
        End If
    End If
    
    Application.Cursor = xlDefault
    Application.StatusBar = False
    
    ' Step 4: Get user selections AND mode in one workflow
    Log "debug", "Getting user selection and mode", DEBUG_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
    Dim selectedValues As Collection
    Dim modeTransposed As Boolean
    Set selectedValues = GetSelectedValuesWithMode(categoryInfo, modeTransposed)
    
    If selectedValues Is Nothing Or selectedValues.count = 0 Then
        Log "debug", "User cancelled or no selection", WARNING_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
        m_userCancelled = True ' Set module-level flag to indicate user cancellation
        ProcessCategorySimplified = False ' User cancelled
        Exit Function
    End If
    Log "debug", "User selected " & selectedValues.count & " values, mode: " & IIf(modeTransposed, "Transposed", "Normal"), DEBUG_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
    
    ' Step 5: Get destination and process data with form
    Log "debug", "Getting destination with form", DEBUG_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
    Dim outputRange As Range
    Set outputRange = GetDestinationWithFormSimplified(categoryInfo, selectedValues, modeTransposed)
    
    If outputRange Is Nothing Then
        Log "debug", "User cancelled destination or form failed", WARNING_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
        m_userCancelled = True ' Set module-level flag to indicate user cancellation
        ProcessCategorySimplified = False
    Else
        Log "debug", "SUCCESS: Data processed to " & outputRange.Address, DEBUG_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
        ProcessCategorySimplified = True
        
        ' Navigate to destination
        With outputRange
            .Parent.Activate
            .Select
            ActiveWindow.ScrollRow = .row
            ActiveWindow.ScrollColumn = .Column
        End With
    End If
    
    Log "debug", "=== ProcessCategorySimplified END ===", DEBUG_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
    Exit Function
    
ErrorHandler:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    Log "debug", "ERROR in ProcessCategorySimplified: " & Err.Number & " - " & Err.Description, ERROR_LEVEL, "ProcessCategorySimplified", "DataLoaderManager"
    ProcessCategorySimplified = False
End Function

' FIXED: GetDestinationWithFormSimplified function with back navigation to mode form
Private Function GetDestinationWithFormSimplified(Category As categoryInfo, selectedValues As Collection, modeTransposed As Boolean) As Range
    On Error GoTo ErrorHandler
    
    Log "debug", "=== GetDestinationWithFormSimplified START ===", DEBUG_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
    
    ' Get source table to calculate size
    Dim lo As ListObject
    Dim tableName As String
    tableName = "Table_" & Utilities.SanitizeTableName(Category.PowerQueryName)
    
    On Error Resume Next
    Set lo = wsPQData.ListObjects(tableName)
    On Error GoTo ErrorHandler
    
    If lo Is Nothing Then
        Log "debug", "ERROR: Source table not found", ERROR_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
        Set GetDestinationWithFormSimplified = Nothing
        Exit Function
    End If
    
    ' Calculate required size
    Dim nbRows As Long, nbCols As Long
    If modeTransposed Then
        nbRows = lo.ListColumns.count
        nbCols = selectedValues.count + 1
    Else
        nbRows = selectedValues.count + 1
        nbCols = lo.ListColumns.count
    End If
    
    Log "debug", "Estimated size: " & nbRows & " x " & nbCols, DEBUG_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
    
    Dim continueFlow As Boolean
    continueFlow = True
    
    Do While continueFlow
        ' Create and configure the output range selection form
        Dim outputForm As OutputRangeSelectionForm
        Set outputForm = New OutputRangeSelectionForm
        
        With outputForm
            .categoryName = Category.DisplayName
            .Mode = IIf(modeTransposed, "Transposed Data", "Normal Data")
            Set .selectedItems = selectedValues
            .estimatedRows = nbRows
            .estimatedCols = nbCols
            .SetProcessingData Category.DisplayName, selectedValues, modeTransposed
        End With
        
        Log "debug", "Showing output form modally", DEBUG_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
        
        ' Show the form modally
        outputForm.Show vbModal
        
        ' Check if user went back
        If outputForm.UserWentBack Then
            Log "debug", "User clicked Back in output form", DEBUG_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
            Unload outputForm
            Set outputForm = Nothing
            
            ' Show Mode Selection Form again
            Dim modeForm As ModeSelectionForm
            Set modeForm = New ModeSelectionForm
            
            Log "debug", "Showing mode selection form from back navigation", DEBUG_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
            
            modeForm.Show vbModal
            
            If modeForm.WasCancelled Then
                Log "debug", "User cancelled mode selection after going back", WARNING_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
                Set GetDestinationWithFormSimplified = Nothing
                Unload modeForm
                Set modeForm = Nothing
                Exit Function
            ElseIf modeForm.WasBack Then
                ' This shouldn't happen as we're already at the mode form
                Log "debug", "Unexpected: User clicked Back from mode form", WARNING_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
                Set GetDestinationWithFormSimplified = Nothing
                Unload modeForm
                Set modeForm = Nothing
                Exit Function
            Else
                ' User selected a new mode
                modeTransposed = modeForm.isTransposed
                Log "debug", "New mode selected: " & IIf(modeTransposed, "Transposed", "Normal"), DEBUG_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
                
                ' Recalculate size if mode changed
                If modeTransposed Then
                    nbRows = lo.ListColumns.count
                    nbCols = selectedValues.count + 1
                Else
                    nbRows = selectedValues.count + 1
                    nbCols = lo.ListColumns.count
                End If
                
                Unload modeForm
                Set modeForm = Nothing
                ' Continue the loop to show output form again with new mode
            End If
            
        ElseIf outputForm.UserCancelled Then
            ' User cancelled
            Log "debug", "User cancelled the output form", WARNING_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
            Set GetDestinationWithFormSimplified = Nothing
            Unload outputForm
            Set outputForm = Nothing
            Exit Function
            
        Else
            ' User processed the data or selected a range
            Dim selectedRange As Range
            Set selectedRange = outputForm.GetSelectedRange
            
            ' Check if data was processed successfully
            If outputForm.WasDataProcessed And Not selectedRange Is Nothing Then
                Log "debug", "Form processed data successfully to: " & selectedRange.Address, DEBUG_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
                Set GetDestinationWithFormSimplified = selectedRange
                continueFlow = False
            Else
                Log "debug", "Form was not processed successfully", WARNING_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
                Set GetDestinationWithFormSimplified = Nothing
                continueFlow = False
            End If
            
            ' Clean up
            Unload outputForm
            Set outputForm = Nothing
        End If
    Loop
    
    Log "debug", "=== GetDestinationWithFormSimplified END ===", DEBUG_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
    Exit Function
    
ErrorHandler:
    Log "debug", "ERROR in GetDestinationWithFormSimplified: " & Err.Description, ERROR_LEVEL, "GetDestinationWithFormSimplified", "DataLoaderManager"
    Set GetDestinationWithFormSimplified = Nothing
    If Not outputForm Is Nothing Then
        Unload outputForm
        Set outputForm = Nothing
    End If
    If Not modeForm Is Nothing Then
        Unload modeForm
        Set modeForm = Nothing
    End If
End Function

' KEEP: The ProcessDataToRange and CopyDataToRange functions from previous artifact
' (They should already be working correctly)

' ADD these missing functions to DataLoaderManager.bas

' Fonction utilitaire pour garantir l'existence de la feuille PQ_DATA et la variable globale
Private Function GetOrCreatePQDataSheet(targetWb As Workbook) As Worksheet
    On Error Resume Next
    Dim ws As Worksheet
    ' Cible le classeur pass� en param�tre (qui sera ActiveWorkbook)
    Set ws = targetWb.Worksheets("PQ_DATA")
    
    If ws Is Nothing Then
        Set ws = targetWb.Worksheets.Add
        ws.Name = "PQ_DATA"
        ' Pas besoin de cacher la feuille car elle est dans le classeur de l'utilisateur
    End If
    
    Set GetOrCreatePQDataSheet = ws
End Function

' UPDATED: GetSelectedValues function that now handles both category selection and mode selection
Private Function GetSelectedValuesWithMode(Category As categoryInfo, ByRef modeTransposed As Boolean) As Collection
    On Error GoTo ErrorHandler
    Dim lastCol As Long
    
    Const PROC_NAME As String = "GetSelectedValuesWithMode"
    Const MODULE_NAME As String = "DataLoaderManager"
    
    Log "debug_selection", "=== GetSelectedValuesWithMode START ===", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Log "debug_selection", "Category: " & Category.DisplayName & ", FilterLevel: " & Category.FilterLevel, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    Dim lo As ListObject
    Dim displayArray() As String
    Dim i As Long, j As Long
    Dim cell As Range
    Dim v As Variant
    
    ' Get table name and ensure table exists
    Dim tableName As String
    tableName = "Table_" & Utilities.SanitizeTableName(Category.PowerQueryName)
    Log "debug_selection", "Looking for table: " & tableName, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    ' Ensure wsPQData exists
    If wsPQData Is Nothing Then
        Set wsPQData = GetOrCreatePQDataSheet(ActiveWorkbook)
    End If
    
    ' Find or create the table
    On Error Resume Next
    Set lo = wsPQData.ListObjects(tableName)
    On Error GoTo ErrorHandler
    
    If lo Is Nothing Then
        Log "debug_selection", "Table not found, loading query...", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        
        ' Load the query
        
        lastCol = Utilities.GetLastColumn(wsPQData)
        LoadQueries.LoadQuery Category.PowerQueryName, wsPQData, wsPQData.Cells(1, lastCol + 1)
        
        ' Try to get the table again
        On Error Resume Next
        Set lo = wsPQData.ListObjects(tableName)
        On Error GoTo ErrorHandler
        
        If lo Is Nothing Then
            MsgBox "Impossible de charger la table PowerQuery '" & Category.PowerQueryName & "'" & vbCrLf & _
                   "V�rifiez votre connexion r�seau et r�essayez.", vbExclamation, "Erreur de chargement"
            Log "debug_selection", "CRITICAL: Could not load table after LoadQuery", ERROR_LEVEL, PROC_NAME, MODULE_NAME
            Set GetSelectedValuesWithMode = Nothing
            Exit Function
        End If
    Else
        ' Table exists, but force refresh if empty to sync with PowerQuery
        If lo.DataBodyRange Is Nothing Then
            Log "debug_selection", "Table exists but is empty, forcing refresh...", WARNING_LEVEL, PROC_NAME, MODULE_NAME
            
            On Error Resume Next
            lo.QueryTable.Refresh BackgroundQuery:=False
            On Error GoTo ErrorHandler
            
            ' If still empty after refresh, recreate the table
            If lo.DataBodyRange Is Nothing Then
                Log "debug_selection", "Table still empty, recreating...", WARNING_LEVEL, PROC_NAME, MODULE_NAME
                lo.Delete
                
                lastCol = Utilities.GetLastColumn(wsPQData)
                LoadQueries.LoadQuery Category.PowerQueryName, wsPQData, wsPQData.Cells(1, lastCol + 1)
                
                On Error Resume Next
                Set lo = wsPQData.ListObjects(tableName)
                On Error GoTo ErrorHandler
                
                If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
                    MsgBox "La table '" & Category.DisplayName & "' ne peut pas être chargée." & vbCrLf & _
                           "La requête PowerQuery fonctionne mais le tableau Excel reste vide.", vbExclamation, "Erreur de synchronisation"
                    Log "debug_selection", "CRITICAL: Table recreation failed", ERROR_LEVEL, PROC_NAME, MODULE_NAME
                    Set GetSelectedValuesWithMode = Nothing
                    Exit Function
                End If
            End If
        End If
    End If
    
    ' Note: Table existence and data validation is now handled above in the improved logic
    
    Log "debug_selection", "Table loaded successfully with " & lo.DataBodyRange.Rows.count & " rows", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    ' For "Pas de filtrage" categories, show all items in the UserForm
    If Category.FilterLevel = "Pas de filtrage" Then
        Log "debug_selection", "Using UserForm for 'Pas de filtrage' category", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        
        ' Build display array with descriptive names
        ReDim displayArray(1 To lo.DataBodyRange.Rows.count)
        
        For i = 1 To lo.DataBodyRange.Rows.count
            Dim id As String
            Dim itemType As String
            Dim reference As String
            
            ' Get basic info with safe error handling
            On Error Resume Next
            id = CStr(lo.DataBodyRange.Cells(i, 1).Value)
            If lo.ListColumns.count >= 2 Then
                itemType = CStr(lo.DataBodyRange.Cells(i, 2).Value)
            End If
            If lo.ListColumns.count >= 3 Then
                reference = CStr(lo.DataBodyRange.Cells(i, 3).Value)
            End If
            On Error GoTo ErrorHandler
            
            ' Create descriptive display name
            If itemType <> "" And reference <> "" Then
                displayArray(i) = id & " - " & itemType & " (" & reference & ")"
            ElseIf itemType <> "" Then
                displayArray(i) = id & " - " & itemType
            Else
                displayArray(i) = "Item " & id
            End If
            
            Log "debug_selection", "Prepared item " & i & ": " & displayArray(i), DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        Next i
        
        ' STEP 1 & 2: Show Category Selection Form and Mode Selection Form with Back functionality
        Dim categoryResult As Collection
        Dim continueFlow As Boolean
        continueFlow = True
        
        Do While continueFlow
            ' Show Category Selection Form
            Dim selectionForm As CategorySelectionForm
            Set selectionForm = New CategorySelectionForm
            
            Log "debug_selection", "Created UserForm, setting up with " & UBound(displayArray) & " items", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
            
            ' Setup the form with category data
            On Error Resume Next
            selectionForm.SetupForm Category.DisplayName, displayArray
            If Err.Number <> 0 Then
                Log "debug_selection", "ERROR setting up form: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
                MsgBox "Erreur lors de la configuration du formulaire: " & Err.Description, vbCritical, "Erreur de formulaire"
                Set GetSelectedValuesWithMode = Nothing
                Exit Function
            End If
            On Error GoTo ErrorHandler
            
            Log "debug_selection", "Form setup complete, showing modal dialog", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
            
            ' Show category form modally
            selectionForm.Show vbModal
            
            ' Check if user cancelled category selection
            If selectionForm.WasCancelled Then
                Log "debug_selection", "User cancelled category selection", WARNING_LEVEL, PROC_NAME, MODULE_NAME
                Set GetSelectedValuesWithMode = Nothing
                Unload selectionForm
                Set selectionForm = Nothing
                Exit Function
            End If
            
            If Not selectionForm.WasNext Then
                Log "debug_selection", "Unexpected form result", WARNING_LEVEL, PROC_NAME, MODULE_NAME
                Set GetSelectedValuesWithMode = Nothing
                Unload selectionForm
                Set selectionForm = Nothing
                Exit Function
            End If
            
            ' Get selected values from category form
            Set categoryResult = selectionForm.GetSelectedValues
            
            ' Cleanup category form
            Unload selectionForm
            Set selectionForm = Nothing
            
            If categoryResult Is Nothing Then
                Log "debug_selection", "Category form returned Nothing - user cancelled", WARNING_LEVEL, PROC_NAME, MODULE_NAME
                Set GetSelectedValuesWithMode = Nothing
                Exit Function
            End If
            
            If categoryResult.count = 0 Then
                Log "debug_selection", "No items selected in category form", WARNING_LEVEL, PROC_NAME, MODULE_NAME
                MsgBox "Veuillez s�lectionner au moins un �l�ment.", vbExclamation, "Aucune s�lection"
                ' Loop will continue to show category form again
            Else
                ' We have selections, show Mode Selection Form
                Dim modeForm As ModeSelectionForm
                Set modeForm = New ModeSelectionForm
                
                Log "debug_selection", "Showing mode selection form", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                
                ' Show mode form modally
                modeForm.Show vbModal
                
                ' Check mode form result
                If modeForm.WasCancelled Then
                    Log "debug_selection", "User cancelled mode selection", WARNING_LEVEL, PROC_NAME, MODULE_NAME
                    Set GetSelectedValuesWithMode = Nothing
                    Unload modeForm
                    Set modeForm = Nothing
                    Exit Function
                ElseIf modeForm.WasBack Then
                    Log "debug_selection", "User clicked Back, returning to category selection", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                    Unload modeForm
                    Set modeForm = Nothing
                    ' Loop will continue to show category form again
                Else
                    ' User clicked Next - get mode selection and exit loop
                    modeTransposed = modeForm.isTransposed
                    Log "debug_selection", "Mode selected: " & IIf(modeTransposed, "Transposed", "Normal"), DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                    
                    ' Cleanup mode form
                    Unload modeForm
                    Set modeForm = Nothing
                    
                    ' Exit the loop
                    continueFlow = False
                End If
            End If
        Loop
        
        ' Convert display names back to IDs
        Set GetSelectedValuesWithMode = New Collection
        
        For Each v In categoryResult
            Log "debug_selection", "Processing selection: " & CStr(v), DEBUG_LEVEL, PROC_NAME, MODULE_NAME
            
            ' Find matching display name and extract ID
            For i = 1 To UBound(displayArray)
                If displayArray(i) = CStr(v) Then
                    Dim selectedId As String
                    selectedId = CStr(lo.DataBodyRange.Cells(i, 1).Value)
                    GetSelectedValuesWithMode.Add selectedId
                    Log "debug_selection", "Added ID: " & selectedId & " for display: " & CStr(v), DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                    Exit For
                End If
            Next i
        Next v
        
        Log "debug_selection", "Final selection: " & GetSelectedValuesWithMode.count & " IDs", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        
        ' This check shouldn't happen anymore as we handle it in the loop above
        If GetSelectedValuesWithMode.count = 0 Then
            Log "debug_selection", "ERROR: No IDs found after conversion", ERROR_LEVEL, PROC_NAME, MODULE_NAME
            Set GetSelectedValuesWithMode = Nothing
            Exit Function
        End If
        
    Else
        ' For filtered categories, use InputBox fallback for now
        Log "debug_selection", "Using InputBox for filtered category: " & Category.FilterLevel, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        
        Dim fallbackInput As String
        fallbackInput = InputBox("S�lection simplifi�e pour " & Category.DisplayName & vbCrLf & _
                               "Entrez 'all' pour s�lectionner tous les �l�ments:", _
                               "S�lection " & Category.DisplayName, "all")
        
        If fallbackInput = "" Then
            Log "debug_selection", "User cancelled InputBox", WARNING_LEVEL, PROC_NAME, MODULE_NAME
            Set GetSelectedValuesWithMode = Nothing
            Exit Function
        ElseIf LCase(fallbackInput) = "all" Then
            ' Select all items
            Set GetSelectedValuesWithMode = New Collection
            For i = 1 To lo.DataBodyRange.Rows.count
                GetSelectedValuesWithMode.Add CStr(lo.DataBodyRange.Cells(i, 1).Value)
            Next i
            Log "debug_selection", "Fallback: Selected all " & GetSelectedValuesWithMode.count & " items", WARNING_LEVEL, PROC_NAME, MODULE_NAME
            
            ' Show mode selection for fallback too
            Dim fallbackModeForm As ModeSelectionForm
            Set fallbackModeForm = New ModeSelectionForm
            
            fallbackModeForm.Show vbModal
            
            If fallbackModeForm.WasCancelled Then
                Log "debug_selection", "User cancelled mode selection in fallback", WARNING_LEVEL, PROC_NAME, MODULE_NAME
                Set GetSelectedValuesWithMode = Nothing
                Unload fallbackModeForm
                Set fallbackModeForm = Nothing
                Exit Function
            End If
            
            modeTransposed = fallbackModeForm.isTransposed
            Unload fallbackModeForm
            Set fallbackModeForm = Nothing
        Else
            Log "debug_selection", "Invalid input in fallback", WARNING_LEVEL, PROC_NAME, MODULE_NAME
            Set GetSelectedValuesWithMode = Nothing
            Exit Function
        End If
    End If
    
    Log "debug_selection", "=== GetSelectedValuesWithMode END - Success with " & GetSelectedValuesWithMode.count & " items ===", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Exit Function
    
ErrorHandler:
    Log "debug_selection", "ERROR in GetSelectedValuesWithMode: " & Err.Number & " - " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    
    Set GetSelectedValuesWithMode = Nothing
    
    ' Cleanup on error
    On Error Resume Next
    If Not selectionForm Is Nothing Then
        Unload selectionForm
        Set selectionForm = Nothing
    End If
    If Not modeForm Is Nothing Then
        Unload modeForm
        Set modeForm = Nothing
    End If
    If Not fallbackModeForm Is Nothing Then
        Unload fallbackModeForm
        Set fallbackModeForm = Nothing
    End If
    On Error GoTo 0
End Function

' Add this function to your DataLoaderManager.bas module

' REQUIRED: ProcessDataToRange function that the OutputRangeSelectionForm calls
Public Function ProcessDataToRange(Category As categoryInfo, selectedValues As Collection, targetRange As Range, transposed As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "ProcessDataToRange"
    Const MODULE_NAME As String = "DataLoaderManager"
    
    Log "debug", "=== ProcessDataToRange START ===", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Log "debug", "Category: " & Category.DisplayName & ", Target: " & targetRange.Address & ", Transposed: " & transposed, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    ' Ensure we have the source table
    Dim lo As ListObject
    Dim tableName As String
    tableName = "Table_" & Utilities.SanitizeTableName(Category.PowerQueryName)
    
    ' Make sure wsPQData is available
    If wsPQData Is Nothing Then
        Set wsPQData = GetOrCreatePQDataSheet(ActiveWorkbook)
    End If
    
    On Error Resume Next
    Set lo = wsPQData.ListObjects(tableName)
    On Error GoTo ErrorHandler
    
    If lo Is Nothing Then
        Log "debug", "ERROR: Source table not found: " & tableName, ERROR_LEVEL, PROC_NAME, MODULE_NAME
        ProcessDataToRange = False
        Exit Function
    End If
    
    If lo.DataBodyRange Is Nothing Then
        Log "debug", "ERROR: Source table has no data", ERROR_LEVEL, PROC_NAME, MODULE_NAME
        ProcessDataToRange = False
        Exit Function
    End If
    
    ' Call the actual data copying function
    Dim success As Boolean
    success = CopyDataToRange(lo, selectedValues, targetRange, transposed)
    
    If success Then
        Log "debug", "Data copied successfully to " & targetRange.Address, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Else
        Log "debug", "ERROR: CopyDataToRange failed", ERROR_LEVEL, PROC_NAME, MODULE_NAME
    End If
    
    ProcessDataToRange = success
    
    Log "debug", "=== ProcessDataToRange END ===", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Exit Function
    
ErrorHandler:
    Log "debug", "ERROR in ProcessDataToRange: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    ProcessDataToRange = False
End Function

' UPDATED: CopyDataToRange function with better error handling
Private Function CopyDataToRange(sourceTable As ListObject, selectedValues As Collection, targetRange As Range, transposed As Boolean) As Boolean
    On Error GoTo ErrorHandler
    
    Const PROC_NAME As String = "CopyDataToRange"
    Const MODULE_NAME As String = "DataLoaderManager"
    
    Log "debug_copy", "=== CopyDataToRange START ===", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Log "debug_copy", "Source table: " & sourceTable.Name & " (" & sourceTable.DataBodyRange.Rows.count & " rows)", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Log "debug_copy", "Selected values: " & selectedValues.count & ", Transposed: " & transposed, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    Log "debug_copy", "Target range: " & targetRange.Address, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim targetWs As Worksheet
    Set targetWs = targetRange.Worksheet
    
    Dim currentRow As Long, currentCol As Long
    currentRow = targetRange.row
    currentCol = targetRange.Column
    
    ' Clear the target area first
    Dim clearRows As Long, clearCols As Long
    If transposed Then
        clearRows = sourceTable.ListColumns.count
        clearCols = selectedValues.count + 1
    Else
        clearRows = selectedValues.count + 1
        clearCols = sourceTable.ListColumns.count
    End If
    
    Dim clearRange As Range
    Set clearRange = targetWs.Range(targetWs.Cells(currentRow, currentCol), targetWs.Cells(currentRow + clearRows - 1, currentCol + clearCols - 1))
    clearRange.Clear
    
    If Not transposed Then
        ' NORMAL MODE: Data in rows
        Log "debug_copy", "Processing NORMAL mode (rows)", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        
        ' Write headers
        Dim colIndex As Long
        For colIndex = 1 To sourceTable.ListColumns.count
            targetWs.Cells(currentRow, currentCol + colIndex - 1).Value = sourceTable.ListColumns(colIndex).Name
        Next colIndex
        currentRow = currentRow + 1
        
        ' Write data rows for selected values
        Dim selectedValue As Variant
        Dim sourceRow As Long
        Dim found As Boolean
        
        For Each selectedValue In selectedValues
            found = False
            Log "debug_copy", "Looking for selected value: " & CStr(selectedValue), DEBUG_LEVEL, PROC_NAME, MODULE_NAME
            
            ' Find the row with this ID
            For sourceRow = 1 To sourceTable.DataBodyRange.Rows.count
                If CStr(sourceTable.DataBodyRange.Cells(sourceRow, 1).Value) = CStr(selectedValue) Then
                    found = True
                    Log "debug_copy", "Found value at source row: " & sourceRow, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                    
                    ' Copy entire row
                    For colIndex = 1 To sourceTable.ListColumns.count
                        targetWs.Cells(currentRow, currentCol + colIndex - 1).Value = sourceTable.DataBodyRange.Cells(sourceRow, colIndex).Value
                    Next colIndex
                    
                    currentRow = currentRow + 1
                    Exit For
                End If
            Next sourceRow
            
            If Not found Then
                Log "debug_copy", "WARNING: Selected value not found in source: " & CStr(selectedValue), WARNING_LEVEL, PROC_NAME, MODULE_NAME
            End If
        Next selectedValue
        
    Else
        ' TRANSPOSED MODE: Data in columns
        Log "debug_copy", "Processing TRANSPOSED mode (columns)", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
        
        ' Write column headers (field names) in first column
        For colIndex = 1 To sourceTable.ListColumns.count
            targetWs.Cells(currentRow + colIndex - 1, currentCol).Value = sourceTable.ListColumns(colIndex).Name
        Next colIndex
        currentCol = currentCol + 1
        
        ' Write data columns for selected values
        Dim valueIndex As Long
        valueIndex = 1
        
        For Each selectedValue In selectedValues
            found = False
            Log "debug_copy", "Processing selected value: " & CStr(selectedValue), DEBUG_LEVEL, PROC_NAME, MODULE_NAME
            
            ' Find the row with this ID
            For sourceRow = 1 To sourceTable.DataBodyRange.Rows.count
                If CStr(sourceTable.DataBodyRange.Cells(sourceRow, 1).Value) = CStr(selectedValue) Then
                    found = True
                    Log "debug_copy", "Found value at source row: " & sourceRow, DEBUG_LEVEL, PROC_NAME, MODULE_NAME
                    
                    ' Copy data as a column
                    For colIndex = 1 To sourceTable.ListColumns.count
                        targetWs.Cells(currentRow + colIndex - 1, currentCol).Value = sourceTable.DataBodyRange.Cells(sourceRow, colIndex).Value
                    Next colIndex
                    
                    currentCol = currentCol + 1
                    Exit For
                End If
            Next sourceRow
            
            If Not found Then
                Log "debug_copy", "WARNING: Selected value not found in source: " & CStr(selectedValue), WARNING_LEVEL, PROC_NAME, MODULE_NAME
            End If
            
            valueIndex = valueIndex + 1
        Next selectedValue
    End If
    
    ' Apply formatting
    Dim dataRange As Range
    Set dataRange = targetWs.Range(targetWs.Cells(targetRange.row, targetRange.Column), targetWs.Cells(currentRow - 1, targetRange.Column + clearCols - 1))
    
    With dataRange
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.ColorIndex = xlAutomatic
    End With
    
    ' Format header row/column
    Dim headerRange As Range
    If Not transposed Then
        Set headerRange = targetWs.Range(targetWs.Cells(targetRange.row, targetRange.Column), targetWs.Cells(targetRange.row, targetRange.Column + clearCols - 1))
    Else
        Set headerRange = targetWs.Range(targetWs.Cells(targetRange.row, targetRange.Column), targetWs.Cells(targetRange.row + clearRows - 1, targetRange.Column))
    End If
    
    With headerRange
        .Font.Bold = True
        .Interior.Color = RGB(200, 220, 240)
    End With
    
    ' Auto-fit columns
    dataRange.Columns.AutoFit
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Log "debug_copy", "=== CopyDataToRange SUCCESS ===", DEBUG_LEVEL, PROC_NAME, MODULE_NAME
    CopyDataToRange = True
    Exit Function
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    Log "debug_copy", "ERROR in CopyDataToRange: " & Err.Description, ERROR_LEVEL, PROC_NAME, MODULE_NAME
    CopyDataToRange = False
End Function

