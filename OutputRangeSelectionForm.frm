VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OutputRangeSelectionForm 
   Caption         =   "UserForm1"
   ClientHeight    =   1908
   ClientLeft      =   -36
   ClientTop       =   60
   ClientWidth     =   576
   OleObjectBlob   =   "OutputRangeSelectionForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OutputRangeSelectionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ========================================
' OUTPUT RANGE SELECTION FORM (DATA LOADER) - ELYSE ENERGY TABLE VERSION
' ========================================
' Output range selection with data processing, progress tracking, and Excel table creation
' ========================================

Option Explicit

' Private member variables
Private m_categoryName As String
Private m_dataType As String
Private m_selectedItems As Collection
Private m_estimatedRows As Long
Private m_estimatedCols As Long
Private m_selectedRange As Range
Private m_wasDataProcessed As Boolean
Private m_userCancelled As Boolean
Private m_userWentBack As Boolean

' Progress bar controls
Private progressFrame As MSForms.Frame
Private progressBar As MSForms.label
Private progressLabel As MSForms.label
Private progressBackground As MSForms.label

' Form controls with events
Private WithEvents btnSelectRange As MSForms.CommandButton
Attribute btnSelectRange.VB_VarHelpID = -1
Private WithEvents btnContinue As MSForms.CommandButton
Attribute btnContinue.VB_VarHelpID = -1
Private WithEvents btnCancel As MSForms.CommandButton
Attribute btnCancel.VB_VarHelpID = -1
Private WithEvents btnBack As MSForms.CommandButton
Attribute btnBack.VB_VarHelpID = -1

' ========================================
' PROPERTIES
' ========================================

Public Property Let categoryName(Value As String)
    m_categoryName = Value
    Call UpdateCategoryDisplay
End Property

Public Property Get categoryName() As String
    categoryName = m_categoryName
End Property

' Add Mode property as an alias for DataType
Public Property Let Mode(Value As String)
    m_dataType = Value
    Call UpdateDataTypeDisplay
End Property

Public Property Get Mode() As String
    Mode = m_dataType
End Property

Public Property Let DataType(Value As String)
    m_dataType = Value
    Call UpdateDataTypeDisplay
End Property

Public Property Get DataType() As String
    DataType = m_dataType
End Property

Public Property Set selectedItems(Value As Collection)
    Set m_selectedItems = Value
    Call UpdateSelectedItemsDisplay
End Property

Public Property Get selectedItems() As Collection
    Set selectedItems = m_selectedItems
End Property

Public Property Let estimatedRows(Value As Long)
    m_estimatedRows = Value
    Call UpdateSizeDisplay
End Property

Public Property Get estimatedRows() As Long
    estimatedRows = m_estimatedRows
End Property

Public Property Let estimatedCols(Value As Long)
    m_estimatedCols = Value
    Call UpdateSizeDisplay
End Property

Public Property Get estimatedCols() As Long
    estimatedCols = m_estimatedCols
End Property

Public Property Get GetSelectedRange() As Range
    Set GetSelectedRange = m_selectedRange
End Property

Public Property Get WasDataProcessed() As Boolean
    WasDataProcessed = m_wasDataProcessed
End Property

Public Property Get UserCancelled() As Boolean
    UserCancelled = m_userCancelled
End Property

Public Property Get UserWentBack() As Boolean
    UserWentBack = m_userWentBack
End Property

' ========================================
' FORM INITIALIZATION
' ========================================

Private Sub UserForm_Initialize()
    ' Set form properties with Elyse Energy branding
    Me.caption = "Data Output Configuration - Elyse Energy"
    Me.BackColor = RGB(248, 249, 250)
    Me.Width = 700
    Me.Height = 580
    Me.StartUpPosition = 1
    
    ' Initialize member variables
    m_wasDataProcessed = False
    m_userCancelled = False
    m_userWentBack = False
    Set m_selectedItems = New Collection
    
    ' Set default output range to active cell
    If Not Application.ActiveCell Is Nothing Then
        Set m_selectedRange = Application.ActiveCell
    End If
    
    Call CreateUI
End Sub

Private Sub CreateUI()
    Dim headerLabel As control
    Dim rangeFrame As control
    Dim summaryFrame As control

    ' Header with Elyse Energy branding
    Set headerLabel = Me.Controls.Add("Forms.Label.1", "lblHeader")
    With headerLabel
        .Left = 20
        .Top = 20
        .Width = 640
        .Height = 35
        .caption = "Energy Equipment Data Export"
        .ForeColor = RGB(17, 36, 148)  ' Elyse primary blue
        .Font.Size = 16
        .Font.Bold = True
    End With

    ' Subtitle with energy theme
    Dim subtitleLabel As control
    Set subtitleLabel = Me.Controls.Add("Forms.Label.1", "lblSubtitle")
    With subtitleLabel
        .Left = 20
        .Top = 60
        .Width = 640
        .Height = 25
        .caption = "Configure output location for your energy equipment specifications"
        .ForeColor = RGB(67, 67, 67)
        .Font.Size = 10
    End With

    ' Summary frame (moved above range selection)
    Set summaryFrame = Me.Controls.Add("Forms.Frame.1", "frameSummary")
    With summaryFrame
        .Left = 20
        .Top = 100
        .Width = 640
        .Height = 120
        .BackColor = RGB(255, 255, 255)
        .caption = "Export Summary"
        .Font.Bold = True
        .ForeColor = RGB(17, 36, 148)
    End With

    ' Summary labels with energy equipment theme
    Call CreateSummaryLabels(summaryFrame)

    ' Instructions
    Dim instructionLabel As control
    Set instructionLabel = Me.Controls.Add("Forms.Label.1", "lblInstruction")
    With instructionLabel
        .Left = 20
        .Top = 240
        .Width = 640
        .Height = 40
        .caption = "Select where to create the data table. The output will be formatted as a professional Excel table with Elyse Energy styling."
        .ForeColor = RGB(67, 67, 67)
        .Font.Size = 9
        .WordWrap = True
    End With

    ' Range selection frame
    Set rangeFrame = Me.Controls.Add("Forms.Frame.1", "frameRange")
    With rangeFrame
        .Left = 20
        .Top = 290
        .Width = 640
        .Height = 120
        .BackColor = RGB(255, 255, 255)
        .caption = "Table Output Location"
        .Font.Bold = True
        .ForeColor = RGB(17, 36, 148)
    End With

    ' Select range button with Elyse styling
    Set btnSelectRange = rangeFrame.Controls.Add("Forms.CommandButton.1", "btnSelectRange")
    With btnSelectRange
        .Left = 20
        .Top = 30
        .Width = 140
        .Height = 35
        .caption = "Select Location"
        .BackColor = RGB(17, 36, 148)  ' Elyse primary blue
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
    End With

    ' Selected range display
    Dim txtSelectedRange As MSForms.TextBox
    Set txtSelectedRange = rangeFrame.Controls.Add("Forms.TextBox.1", "txtSelectedRange")
    With txtSelectedRange
        .Left = 180
        .Top = 30
        .Width = 430
        .Height = 35
        .Font.Size = 11
        .text = IIf(m_selectedRange Is Nothing, "No location selected", m_selectedRange.Address(True, True, xlA1, True))
        .Enabled = False
        .BackColor = RGB(248, 249, 250)
        .BorderStyle = fmBorderStyleSingle
    End With

    ' Range info with table details
    Dim lblRangeInfo As MSForms.label
    Set lblRangeInfo = rangeFrame.Controls.Add("Forms.Label.1", "lblRangeInfo")
    With lblRangeInfo
        .Left = 20
        .Top = 75
        .Width = 590
        .Height = 35
        .caption = "Table size: " & m_estimatedRows & " rows × " & m_estimatedCols & " columns | " & _
                   "Will create Excel table with Elyse Energy formatting"
        .ForeColor = RGB(100, 100, 100)
        .Font.Size = 9
        .WordWrap = True
    End With

    ' Buttons with Elyse Energy styling
    Set btnBack = Me.Controls.Add("Forms.CommandButton.1", "btnBack")
    With btnBack
        .Left = 20
        .Top = 430
        .Width = 90
        .Height = 40
        .caption = "? Back"
        .BackColor = RGB(50, 231, 185)  ' Elyse teal
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 10
    End With
    
    Set btnContinue = Me.Controls.Add("Forms.CommandButton.1", "btnContinue")
    With btnContinue
        .Left = 470
        .Top = 430
        .Width = 110
        .Height = 40
        .caption = "Create Table"
        .BackColor = RGB(50, 231, 185)  ' Elyse teal
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 10
        .Enabled = Not (m_selectedRange Is Nothing)
    End With

    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    With btnCancel
        .Left = 590
        .Top = 430
        .Width = 90
        .Height = 40
        .caption = "Cancel"
        .BackColor = RGB(220, 220, 220)
        .ForeColor = RGB(80, 80, 80)
        .Font.Size = 10
    End With
End Sub

Private Sub CreateSummaryLabels(summaryFrame As control)
    Dim lblCategory As control
    Dim lblDataType As control
    Dim lblSelectedCount As control
    Dim lblEquation As control

    ' Equipment category
    Set lblCategory = summaryFrame.Controls.Add("Forms.Label.1", "lblCategory")
    With lblCategory
        .Left = 20
        .Top = 25
        .Width = 300
        .Height = 18
        .caption = "Equipment Category: " & m_categoryName
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(17, 36, 148)
    End With

    ' Data organization mode
    Set lblDataType = summaryFrame.Controls.Add("Forms.Label.1", "lblDataType")
    With lblDataType
        .Left = 340
        .Top = 25
        .Width = 280
        .Height = 18
        .caption = "Layout Mode: " & m_dataType
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(17, 36, 148)
    End With

    ' Selected equipment count
    Set lblSelectedCount = summaryFrame.Controls.Add("Forms.Label.1", "lblSelectedCount")
    With lblSelectedCount
        .Left = 20
        .Top = 50
        .Width = 300
        .Height = 18
        .caption = "Equipment Units: " & IIf(m_selectedItems Is Nothing, 0, m_selectedItems.count) & " selected"
        .Font.Size = 9
        .ForeColor = RGB(80, 80, 80)
    End With

    ' Energy equation info
    Set lblEquation = summaryFrame.Controls.Add("Forms.Label.1", "lblEquation")
    With lblEquation
        .Left = 20
        .Top = 75
        .Width = 580
        .Height = 35
        .caption = "Includes: Energy ID, Type, Reference data, Performance equations (y = a·exp(-b·x)), " & _
                   "Pressure ranges, Stage configurations, and Electrical/Cooling coefficients"
        .Font.Size = 8
        .ForeColor = RGB(100, 100, 100)
        .WordWrap = True
    End With
End Sub

' ========================================
' EVENT HANDLERS
' ========================================

Private Sub btnSelectRange_Click()
    Dim userRange As Range
    Dim prompt As String

    Me.Hide

    prompt = "Select the TOP-LEFT CELL where you want to create the energy equipment table." & vbCrLf & vbCrLf & _
             "Table specifications:" & vbCrLf & _
             "• Size: " & m_estimatedRows & " rows × " & m_estimatedCols & " columns" & vbCrLf & _
             "• Format: Professional Excel table with Elyse Energy styling" & vbCrLf & _
             "• Content: " & m_selectedItems.count & " equipment units with full specifications"

    On Error GoTo SelectionCancelled

    Application.ScreenUpdating = True
    Set userRange = Application.InputBox( _
        prompt:=prompt, _
        title:="Select Table Location - " & m_categoryName, _
        Type:=8)

    Set m_selectedRange = userRange.Cells(1, 1)
    Call UpdateRangeDisplay

    Me.Show
    Exit Sub

SelectionCancelled:
    Me.Show
    On Error GoTo 0
End Sub

Private Sub btnContinue_Click()
    If m_selectedRange Is Nothing Then
        ShowError "Please select an output location first."
        Exit Sub
    End If

    If Not ValidateSelectedRange() Then Exit Sub

    ' Show progress and process data
    Call ShowProgressBar

    Dim success As Boolean
    success = ProcessDataToRange(m_selectedRange, m_categoryName, m_selectedItems, (m_dataType = "Transposed Data"), Me)

    Call HideProgressBar

    If success Then
        m_wasDataProcessed = True
        
        ' Show success message
        ShowMessage "Energy equipment table created successfully!" & vbCrLf & vbCrLf & _
                   "Table features:" & vbCrLf & _
                   "? Professional Elyse Energy formatting" & vbCrLf & _
                   "? Equipment specifications and coefficients" & vbCrLf & _
                   "? Performance equations and ranges" & vbCrLf & _
                   "? Excel table format for easy data management" & vbCrLf & vbCrLf & _
                   "Title cell is now selected to show the full summary.", "Export Complete"
        
        Me.Hide
    Else
        ShowError "Table creation failed. Please check your data and try again."
    End If
End Sub

Private Sub btnCancel_Click()
    m_userCancelled = True
    m_wasDataProcessed = False
    Me.Hide
End Sub

Private Sub btnBack_Click()
    m_userWentBack = True
    m_wasDataProcessed = False
    Me.Hide
End Sub

' ========================================
' HELPER FUNCTIONS
' ========================================

Private Sub UpdateRangeDisplay()
    If Not m_selectedRange Is Nothing Then
        Dim txtRange As MSForms.TextBox
        Dim lblInfo As MSForms.label
        Dim endRange As Range

        Set txtRange = Me.Controls("frameRange").Controls("txtSelectedRange")
        Set lblInfo = Me.Controls("frameRange").Controls("lblRangeInfo")
        Set endRange = m_selectedRange.Offset(m_estimatedRows - 1, m_estimatedCols - 1)

        txtRange.text = m_selectedRange.Address(True, True, xlA1, True)
        lblInfo.caption = "Table will span from " & m_selectedRange.Address & " to " & endRange.Address & _
                         " | Size: " & m_estimatedRows & " rows × " & m_estimatedCols & " columns with Elyse Energy styling"

        btnContinue.Enabled = True
        btnContinue.BackColor = RGB(50, 231, 185)
    End If
End Sub

Private Sub UpdateCategoryDisplay()
    On Error Resume Next
    Dim lblCategory As control
    Set lblCategory = Me.Controls("frameSummary").Controls("lblCategory")
    If Not lblCategory Is Nothing Then
        lblCategory.caption = "Equipment Category: " & m_categoryName
    End If
    On Error GoTo 0
End Sub

Private Sub UpdateDataTypeDisplay()
    On Error Resume Next
    Dim lblDataType As control
    Set lblDataType = Me.Controls("frameSummary").Controls("lblDataType")
    If Not lblDataType Is Nothing Then
        lblDataType.caption = "Layout Mode: " & m_dataType
    End If
    On Error GoTo 0
End Sub

Private Sub UpdateSelectedItemsDisplay()
    On Error Resume Next
    Dim lblSelectedCount As control
    Set lblSelectedCount = Me.Controls("frameSummary").Controls("lblSelectedCount")
    If Not lblSelectedCount Is Nothing Then
        lblSelectedCount.caption = "Equipment Units: " & IIf(m_selectedItems Is Nothing, 0, m_selectedItems.count) & " selected"
    End If
    On Error GoTo 0
End Sub

Private Sub UpdateSizeDisplay()
    On Error Resume Next
    Dim lblInfo As MSForms.label
    Set lblInfo = Me.Controls("frameRange").Controls("lblRangeInfo")
    If Not lblInfo Is Nothing Then
        lblInfo.caption = "Table size: " & m_estimatedRows & " rows × " & m_estimatedCols & " columns | " & _
                         "Will create Excel table with Elyse Energy formatting"
    End If
    On Error GoTo 0
End Sub

Private Function ValidateSelectedRange() As Boolean
    Dim testRange As Range
    Dim cell As Range
    Dim nonEmptyCount As Integer

    ValidateSelectedRange = False
    Set testRange = m_selectedRange.Resize(m_estimatedRows, m_estimatedCols)

    ' Count non-empty cells
    nonEmptyCount = 0
    For Each cell In testRange.Cells
        If Not IsEmpty(cell.Value) And Trim(CStr(cell.Value)) <> "" Then
            nonEmptyCount = nonEmptyCount + 1
            If nonEmptyCount > 20 Then Exit For
        End If
    Next cell

    ' Handle existing data
    If nonEmptyCount > 10 Then
        If Not ShowConfirmationDialog("The selected area contains " & nonEmptyCount & " non-empty cells." & vbCrLf & vbCrLf & _
                                     "Creating the energy equipment table will overwrite this data." & vbCrLf & _
                                     "Continue with table creation?", "Confirm Table Creation") Then
            Exit Function
        End If
    ElseIf nonEmptyCount > 0 Then
        If Not ShowConfirmationDialog("The selected area contains " & nonEmptyCount & " cells with data." & vbCrLf & _
                                     "Continue with creating the table here?", "Confirm Location") Then
            Exit Function
        End If
    End If

    ValidateSelectedRange = True
End Function

' ========================================
' PROGRESS BAR FUNCTIONS
' ========================================

Private Sub ShowProgressBar()
    ' Disable buttons
    btnSelectRange.Enabled = False
    btnContinue.Enabled = False
    btnCancel.Enabled = False
    btnBack.Enabled = False

    ' Create progress frame with Elyse styling
    Set progressFrame = Me.Controls.Add("Forms.Frame.1", "frameProgress")
    With progressFrame
        .Left = 20
        .Top = 430
        .Width = 640
        .Height = 90
        .BackColor = RGB(255, 255, 255)
        .caption = "Creating Energy Equipment Table..."
        .Font.Bold = True
        .ForeColor = RGB(17, 36, 148)
        .visible = True
    End With

    ' Progress label
    Set progressLabel = progressFrame.Controls.Add("Forms.Label.1", "lblProgress")
    With progressLabel
        .Left = 20
        .Top = 25
        .Width = 600
        .Height = 20
        .caption = "Initializing table creation..."
        .Font.Size = 9
        .ForeColor = RGB(67, 67, 67)
    End With

    ' Progress background
    Set progressBackground = progressFrame.Controls.Add("Forms.Label.1", "lblProgressBG")
    With progressBackground
        .Left = 20
        .Top = 50
        .Width = 600
        .Height = 25
        .BackColor = RGB(230, 230, 230)
        .BorderStyle = fmBorderStyleSingle
    End With

    ' Progress bar with Elyse teal
    Set progressBar = progressFrame.Controls.Add("Forms.Label.1", "lblProgressBar")
    With progressBar
        .Left = 21
        .Top = 51
        .Width = 0
        .Height = 23
        .BackColor = RGB(50, 231, 185)  ' Elyse teal
    End With

    DoEvents
End Sub

Public Sub UpdateProgress(progressPercent As Long, statusText As String)
    If Not progressLabel Is Nothing Then
        progressLabel.caption = statusText
    End If

    If Not progressBar Is Nothing Then
        Dim newWidth As Long
        newWidth = (598 * progressPercent) / 100
        progressBar.Width = newWidth
    End If

    DoEvents
End Sub

Private Sub HideProgressBar()
    On Error Resume Next
    If Not progressFrame Is Nothing Then
        Me.Controls.Remove "frameProgress"
        Set progressFrame = Nothing
        Set progressLabel = Nothing
        Set progressBackground = Nothing
        Set progressBar = Nothing
    End If

    ' Re-enable buttons
    btnSelectRange.Enabled = True
    If Not m_selectedRange Is Nothing Then btnContinue.Enabled = True
    btnCancel.Enabled = True
    btnBack.Enabled = True

    On Error GoTo 0
End Sub

' ========================================
' PUBLIC METHODS FOR DATA PROCESSING
' ========================================

Public Sub SetProcessingData(categoryDisplayName As String, selectedValues As Collection, modeTransposed As Boolean)
    m_categoryName = categoryDisplayName
    Set m_selectedItems = selectedValues
    m_dataType = IIf(modeTransposed, "Transposed Data", "Normal Data")
    
    ' Update form display
    Me.categoryName = categoryDisplayName
    Me.DataType = m_dataType
    Set Me.selectedItems = selectedValues
    
    ' Enable continue button if we have a range
    If Not m_selectedRange Is Nothing Then
        btnContinue.Enabled = True
    End If
End Sub

' ========================================
' DATA PROCESSING FUNCTIONS
' ========================================

Public Function ProcessDataToRange(outputRange As Range, categoryDisplayName As String, selectedItems As Collection, modeTransposed As Boolean, progressForm As OutputRangeSelectionForm) As Boolean
    On Error GoTo ErrorHandler
    
    ' Update progress
    If Not progressForm Is Nothing Then
        progressForm.UpdateProgress 15, "Initializing energy equipment table creation..."
    End If
    
    ' Get source table - construct table name from category display name
    Dim sourceTable As ListObject
    Dim tableName As String
    Dim sanitizedName As String
    sanitizedName = Replace(categoryDisplayName, " ", "_")
    sanitizedName = Replace(sanitizedName, "-", "_")
    sanitizedName = Replace(sanitizedName, "(", "")
    sanitizedName = Replace(sanitizedName, ")", "")
    tableName = "Table_" & sanitizedName
    
    ' Update progress
    If Not progressForm Is Nothing Then
        progressForm.UpdateProgress 25, "Loading energy equipment data source..."
    End If
    
    ' Ensure wsPQData exists
    If wsPQData Is Nothing Then
        Set wsPQData = GetOrCreatePQDataSheet()
    End If
    
    ' Try to find the table
    Dim tbl As ListObject
    For Each tbl In wsPQData.ListObjects
        If InStr(1, tbl.Name, sanitizedName, vbTextCompare) > 0 Or _
           InStr(1, tbl.Name, Replace(categoryDisplayName, " ", ""), vbTextCompare) > 0 Then
            Set sourceTable = tbl
            Exit For
        End If
    Next tbl
    
    If sourceTable Is Nothing Then
        ' Try exact match
        On Error Resume Next
        Set sourceTable = wsPQData.ListObjects(tableName)
        On Error GoTo ErrorHandler
    End If
    
    If sourceTable Is Nothing Then
        ProcessDataToRange = False
        Exit Function
    End If
    
    ' Update progress
    If Not progressForm Is Nothing Then
        progressForm.UpdateProgress 40, "Processing energy equipment specifications..."
    End If
    
    ' Copy data and create table using enhanced approach
    Dim success As Boolean
    success = CreateElyseEnergyTable(sourceTable, selectedItems, outputRange, modeTransposed, progressForm)
    
    ' Update progress
    If Not progressForm Is Nothing Then
        progressForm.UpdateProgress 100, "Energy equipment table creation complete!"
    End If
    
    ProcessDataToRange = success
    Exit Function
    
ErrorHandler:
    ProcessDataToRange = False
End Function

' ========================================
' ENHANCED TABLE CREATION WITH ELYSE ENERGY STYLING
' ========================================

Private Function CreateElyseEnergyTable(sourceTable As ListObject, selectedItems As Collection, outputRange As Range, modeTransposed As Boolean, Optional progressForm As OutputRangeSelectionForm = Nothing) As Boolean
    On Error GoTo ErrorHandler
    
    Dim destSheet As Worksheet
    Set destSheet = outputRange.Parent
    
    Dim wasProtected As Boolean
    wasProtected = destSheet.ProtectContents
    If wasProtected Then destSheet.Unprotect
    
    ' Update progress
    If Not progressForm Is Nothing Then
        progressForm.UpdateProgress 50, "Preparing Elyse Energy table structure..."
    End If
    
    ' Calculate table dimensions using filtered columns
    Dim filteredCols As Collection
    Set filteredCols = GetFilteredColumns(sourceTable)
    
    Dim tableRows As Long, tableCols As Long
    If modeTransposed Then
        tableRows = filteredCols.count
        tableCols = selectedItems.count + 1
    Else
        tableRows = selectedItems.count + 1
        tableCols = filteredCols.count
    End If
    
    ' Clear and prepare the area
    Dim tableRange As Range
    Set tableRange = outputRange.Resize(tableRows, tableCols)
    tableRange.Clear
    tableRange.ClearFormats
    
    ' Add summary header above the table (with dynamic content)
    Call CreateSummaryHeader(outputRange, selectedItems.count, modeTransposed, sourceTable, selectedItems)
    
    ' Adjust table range to start below summary (reduced offset)
    Set tableRange = outputRange.Resize(tableRows, tableCols)
    
    ' Update progress
    If Not progressForm Is Nothing Then
        progressForm.UpdateProgress 60, "Copying energy equipment data..."
    End If
    
    ' Copy data based on mode
    If Not modeTransposed Then
        Call CopyNormalModeData(sourceTable, selectedItems, tableRange, progressForm)
    Else
        Call CopyTransposedModeData(sourceTable, selectedItems, tableRange, progressForm)
    End If
    
    ' Update progress
    If Not progressForm Is Nothing Then
        progressForm.UpdateProgress 85, "Creating Excel table with Elyse Energy styling..."
    End If
    
    ' Create Excel Table
    Dim excelTable As ListObject
    Set excelTable = destSheet.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    
    ' Name the table with Elyse Energy convention
    Dim tableNameBase As String
    tableNameBase = "ElyseEnergyTable_" & Format(Now, "hhmmss")
    excelTable.Name = tableNameBase
    
    ' Apply Elyse Energy table styling
    Call ApplyElyseEnergyTableStyle(excelTable, modeTransposed)
    
    ' Update progress
    If Not progressForm Is Nothing Then
        progressForm.UpdateProgress 95, "Finalizing table formatting and styling..."
    End If
    
    ' Auto-fit columns
    tableRange.Columns.AutoFit
    
    ' Restore protection
    If wasProtected Then destSheet.Protect UserInterfaceOnly:=True
    
    ' SELECT THE TITLE CELL AT THE VERY END OF TABLE CREATION
    ' This ensures nothing else overrides our selection
    Application.ScreenUpdating = False
    Dim titleCell As Range
    Set titleCell = outputRange.Offset(-4, 0)
    titleCell.Select
    Application.ScreenUpdating = True
    
    CreateElyseEnergyTable = True
    Exit Function
    
ErrorHandler:
    If wasProtected Then destSheet.Protect UserInterfaceOnly:=True
    CreateElyseEnergyTable = False
End Function

Private Sub CreateSummaryHeader(startRange As Range, itemCount As Long, isTransposed As Boolean, sourceTable As ListObject, selectedItems As Collection)
    Dim ws As Worksheet
    Set ws = startRange.Parent
    
    ' Extract summary data from the first selected item
    Dim summaryData As String
    Dim firstItem As Variant
    If selectedItems.count > 0 Then
        firstItem = selectedItems(1)
        summaryData = ExtractSummaryFromTableData(sourceTable, firstItem)
    End If
    
    ' Calculate the number of columns for merging
    Dim filteredCols As Collection
    Set filteredCols = GetFilteredColumns(sourceTable)
    Dim mergeWidth As Long
    mergeWidth = filteredCols.count
    
    ' Title (merge across table width)
    Dim titleRange As Range
    Set titleRange = ws.Range(startRange.Offset(-4, 0), startRange.Offset(-4, mergeWidth - 1))
    With titleRange
        .Merge
        .Value = "ELYSE ENERGY - EQUIPMENT SPECIFICATIONS"
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(17, 36, 148)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 25
    End With
    
    ' Equipment info (merge across table width)
    Dim infoRange As Range
    Set infoRange = ws.Range(startRange.Offset(-3, 0), startRange.Offset(-3, mergeWidth - 1))
    With infoRange
        .Merge
        .Value = "Equipment Count: " & itemCount & " | Layout: " & IIf(isTransposed, "Transposed", "Standard") & " | Generated: " & Format(Now, "yyyy-mm-dd hh:mm")
        .Font.Size = 10
        .Font.Color = RGB(67, 67, 67)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 20
    End With
    
    ' Dynamic summary from table data (merge across table width with proper height)
    If summaryData <> "" Then
        Dim summaryRange As Range
        Set summaryRange = ws.Range(startRange.Offset(-2, 0), startRange.Offset(-1, mergeWidth - 1))  ' Use 2 rows for height
        With summaryRange
            .Merge
            .Value = summaryData
            .Font.Size = 9
            .Font.Color = RGB(100, 100, 100)
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
            .WrapText = True
            .RowHeight = 50  ' Increased height for better readability
        End With
        
        ' Adjust table start position
        Set startRange = startRange.Offset(1, 0)  ' Move table start down by 1 more row
    End If
End Sub

Private Function ExtractSummaryFromTableData(sourceTable As ListObject, selectedItem As Variant) As String
    On Error Resume Next
    
    Dim summaryText As String
    Dim sourceRow As Long
    
    ' Find the row for the selected item
    For sourceRow = 1 To sourceTable.DataBodyRange.Rows.count
        If CStr(sourceTable.DataBodyRange.Cells(sourceRow, 1).Value) = CStr(selectedItem) Then
            ' Extract summary fields (typically text-heavy columns)
            Dim col As Long
            For col = 1 To sourceTable.ListColumns.count
                Dim columnName As String
                Dim cellValue As String
                
                columnName = LCase(sourceTable.ListColumns(col).Name)
                cellValue = CStr(sourceTable.DataBodyRange.Cells(sourceRow, col).Value)
                
                ' Identify summary fields by name patterns and content length
                If (InStr(columnName, "equation") > 0 Or _
                    InStr(columnName, "information") > 0 Or _
                    InStr(columnName, "change log") > 0 Or _
                    InStr(columnName, "description") > 0 Or _
                    InStr(columnName, "note") > 0 Or _
                    InStr(columnName, "comment") > 0 Or _
                    Len(cellValue) > 50) And _
                   cellValue <> "" Then
                    
                    If summaryText <> "" Then summaryText = summaryText & " | "
                    summaryText = summaryText & sourceTable.ListColumns(col).Name & ": " & cellValue
                End If
            Next col
            Exit For
        End If
    Next sourceRow
    
    ExtractSummaryFromTableData = summaryText
    On Error GoTo 0
End Function

Private Function GetFilteredColumns(sourceTable As ListObject) As Collection
    ' Return only columns that should appear in the table (excluding summary fields)
    Dim filteredCols As New Collection
    Dim col As Long
    
    For col = 1 To sourceTable.ListColumns.count
        Dim columnName As String
        columnName = LCase(sourceTable.ListColumns(col).Name)
        
        ' Include column if it's NOT a summary field
        If Not (InStr(columnName, "equation") > 0 Or _
                InStr(columnName, "information") > 0 Or _
                InStr(columnName, "change log") > 0 Or _
                InStr(columnName, "description") > 0 Or _
                InStr(columnName, "note") > 0 Or _
                InStr(columnName, "comment") > 0) Then
            
            ' Also check if the content is typically summary-like (very long text)
            Dim hasLongContent As Boolean
            hasLongContent = False
            
            ' Sample first few rows to check content length
            Dim checkRow As Long
            For checkRow = 1 To Application.Min(5, sourceTable.DataBodyRange.Rows.count)
                If Len(CStr(sourceTable.DataBodyRange.Cells(checkRow, col).Value)) > 50 Then
                    hasLongContent = True
                    Exit For
                End If
            Next checkRow
            
            If Not hasLongContent Then
                filteredCols.Add col
            End If
        End If
    Next col
    
    Set GetFilteredColumns = filteredCols
End Function

Private Sub CopyNormalModeData(sourceTable As ListObject, selectedItems As Collection, tableRange As Range, Optional progressForm As OutputRangeSelectionForm = Nothing)
    ' Get filtered columns (excluding summary fields)
    Dim filteredCols As Collection
    Set filteredCols = GetFilteredColumns(sourceTable)
    
    ' Copy headers for filtered columns only
    Dim colIndex As Long
    colIndex = 1
    Dim col As Variant
    For Each col In filteredCols
        tableRange.Cells(1, colIndex).Value = sourceTable.ListColumns(col).Name
        colIndex = colIndex + 1
    Next col
    
    ' Copy data rows for filtered columns only
    Dim dataRow As Long: dataRow = 2
    Dim itemIndex As Long: itemIndex = 0
    Dim v As Variant
    
    For Each v In selectedItems
        itemIndex = itemIndex + 1
        
        If Not progressForm Is Nothing Then
            Dim progressPercent As Long
            progressPercent = 60 + (20 * itemIndex / selectedItems.count)
            progressForm.UpdateProgress progressPercent, "Processing equipment " & itemIndex & " of " & selectedItems.count & "..."
        End If
        
        Dim sourceRow As Long
        For sourceRow = 1 To sourceTable.DataBodyRange.Rows.count
            If CStr(sourceTable.DataBodyRange.Cells(sourceRow, 1).Value) = CStr(v) Then
                colIndex = 1
                For Each col In filteredCols
                    tableRange.Cells(dataRow, colIndex).Value = sourceTable.DataBodyRange.Cells(sourceRow, col).Value
                    colIndex = colIndex + 1
                Next col
                dataRow = dataRow + 1
                Exit For
            End If
        Next sourceRow
    Next v
End Sub

Private Sub CopyTransposedModeData(sourceTable As ListObject, selectedItems As Collection, tableRange As Range, Optional progressForm As OutputRangeSelectionForm = Nothing)
    ' Get filtered columns (excluding summary fields)
    Dim filteredCols As Collection
    Set filteredCols = GetFilteredColumns(sourceTable)
    
    ' Copy field names as row headers for filtered columns only
    Dim rowIndex As Long
    rowIndex = 1
    Dim col As Variant
    For Each col In filteredCols
        tableRange.Cells(rowIndex, 1).Value = sourceTable.ListColumns(col).Name
        rowIndex = rowIndex + 1
    Next col
    
    ' Copy data as columns for filtered columns only
    Dim dataCol As Long: dataCol = 2
    Dim itemIndex As Long: itemIndex = 0
    Dim v As Variant
    
    For Each v In selectedItems
        itemIndex = itemIndex + 1
        
        If Not progressForm Is Nothing Then
            Dim progressPercent As Long
            progressPercent = 60 + (20 * itemIndex / selectedItems.count)
            progressForm.UpdateProgress progressPercent, "Processing equipment column " & itemIndex & " of " & selectedItems.count & "..."
        End If
        
        Dim sourceRow As Long
        For sourceRow = 1 To sourceTable.DataBodyRange.Rows.count
            If CStr(sourceTable.DataBodyRange.Cells(sourceRow, 1).Value) = CStr(v) Then
                rowIndex = 1
                For Each col In filteredCols
                    tableRange.Cells(rowIndex, dataCol).Value = sourceTable.DataBodyRange.Cells(sourceRow, col).Value
                    rowIndex = rowIndex + 1
                Next col
                dataCol = dataCol + 1
                Exit For
            End If
        Next sourceRow
    Next v
End Sub

Private Sub ApplyElyseEnergyTableStyle(excelTable As ListObject, isTransposed As Boolean)
    ' Apply Elyse Energy table style
    With excelTable
        ' Set table style to a professional one
        .TableStyle = "TableStyleMedium2"
        
        ' Override with Elyse Energy colors
        With .Range
            .Font.Name = "Calibri"
            .Font.Size = 11
        End With
        
        ' Header styling with Elyse primary blue
        With .HeaderRowRange
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)  ' White text
            .Interior.Color = RGB(17, 36, 148)  ' Elyse primary blue
            .Font.Size = 11
        End With
        
        ' Energy ID column/row highlighting with Elyse teal
        If Not isTransposed Then
            ' Normal mode - highlight first column (Energy ID)
            If .ListColumns.count > 0 Then
                ' Find the ID column (should be first filtered column)
                Dim idColumnFound As Boolean
                idColumnFound = False
                Dim col As Long
                For col = 1 To .ListColumns.count
                    If LCase(.ListColumns(col).Name) = "id" Or InStr(LCase(.ListColumns(col).Name), "energy") > 0 Then
                        With .ListColumns(col).DataBodyRange
                            .Font.Bold = True
                            .Font.Color = RGB(255, 255, 255)
                            .Interior.Color = RGB(50, 231, 185)  ' Elyse teal
                            .HorizontalAlignment = xlCenter
                        End With
                        idColumnFound = True
                        Exit For
                    End If
                Next col
                
                ' If no ID column found specifically, highlight first column
                If Not idColumnFound And .ListColumns.count > 0 Then
                    With .ListColumns(1).DataBodyRange
                        .Font.Bold = True
                        .Font.Color = RGB(255, 255, 255)
                        .Interior.Color = RGB(50, 231, 185)  ' Elyse teal
                        .HorizontalAlignment = xlCenter
                    End With
                End If
            End If
        Else
            ' Transposed mode - highlight first row (Energy ID)
            If .ListRows.count > 0 Then
                ' Find the ID row
                Dim idRowFound As Boolean
                idRowFound = False
                Dim row As Long
                For row = 1 To .ListRows.count
                    If LCase(.Range.Cells(row + 1, 1).Value) = "id" Or InStr(LCase(.Range.Cells(row + 1, 1).Value), "energy") > 0 Then
                        With .ListRows(row).Range
                            .Font.Bold = True
                            .Font.Color = RGB(255, 255, 255)
                            .Interior.Color = RGB(50, 231, 185)  ' Elyse teal
                            .HorizontalAlignment = xlCenter
                        End With
                        idRowFound = True
                        Exit For
                    End If
                Next row
                
                ' If no ID row found specifically, highlight first row
                If Not idRowFound And .ListRows.count > 0 Then
                    With .ListRows(1).Range
                        .Font.Bold = True
                        .Font.Color = RGB(255, 255, 255)
                        .Interior.Color = RGB(50, 231, 185)  ' Elyse teal
                        .HorizontalAlignment = xlCenter
                    End With
                End If
            End If
        End If
        
        ' Apply alternating row colors for better readability
        Dim rowIndex As Long
        For rowIndex = 1 To .ListRows.count
            If rowIndex Mod 2 = 0 Then
                ' Don't override ID highlighting
                Dim currentRow As Range
                Set currentRow = .ListRows(rowIndex).Range
                Dim cell As Range
                For Each cell In currentRow.Cells
                    If cell.Interior.Color <> RGB(50, 231, 185) Then  ' Don't override teal highlighting
                        cell.Interior.Color = RGB(248, 249, 250)  ' Light gray
                    End If
                Next cell
            End If
        Next rowIndex
        
        ' Professional borders
        With .Range.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(200, 200, 200)
        End With
        
        ' Thicker border around the entire table
        With .Range.Borders(xlEdgeTop)
            .Weight = xlMedium
            .Color = RGB(17, 36, 148)
        End With
        With .Range.Borders(xlEdgeBottom)
            .Weight = xlMedium
            .Color = RGB(17, 36, 148)
        End With
        With .Range.Borders(xlEdgeLeft)
            .Weight = xlMedium
            .Color = RGB(17, 36, 148)
        End With
        With .Range.Borders(xlEdgeRight)
            .Weight = xlMedium
            .Color = RGB(17, 36, 148)
        End With
    End With
End Sub

' ========================================
' UTILITY FUNCTIONS
' ========================================

Private Sub ShowMessage(message As String, Optional title As String = "Elyse Energy")
    MsgBox message, vbInformation, title
End Sub

Private Sub ShowError(message As String)
    MsgBox message, vbExclamation, "Elyse Energy - Error"
End Sub

Private Function ShowConfirmationDialog(message As String, title As String) As Boolean
    ShowConfirmationDialog = (MsgBox(message, vbYesNo + vbQuestion, "Elyse Energy - " & title) = vbYes)
End Function

Private Function GetOrCreatePQDataSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreatePQDataSheet = Worksheets("PQ_DATA")
    On Error GoTo 0
    If GetOrCreatePQDataSheet Is Nothing Then
        ' Call the initialization function directly if available
        On Error Resume Next
        Application.Run "Utilities.InitializePQData"
        On Error GoTo 0
        Set GetOrCreatePQDataSheet = Worksheets("PQ_DATA")
    End If
End Function

