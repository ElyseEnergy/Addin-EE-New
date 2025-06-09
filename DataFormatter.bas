Attribute VB_Name = "DataFormatter"
Option Explicit

'==================================================================================================
' TYPES
'==================================================================================================

Public Type FormattedCellOutput
    FinalValue As Variant          ' The value to be written to the cell
    NumberFormatString As String   ' The number format string to apply (e.g., "DD/MM/YYYY", "@", "0.00")
    IsSectionHeader As Boolean     ' True if this cell represents a section header

    ' Styling specific to Section Headers
    SectionFontName As String      ' Font name for section header (empty for default)
    SectionFontSizeOffset As Long  ' e.g., +3 to base font size
    SectionFontColor As Long       ' RGB color for section header font
    SectionFontBold As Boolean     ' True if section header font is bold
End Type

'==================================================================================================
' CONSTANTS
'==================================================================================================

' Default color for section headers. RGB(128, 128, 128) is Medium Gray.
Public Const SECTION_HEADER_DEFAULT_FONT_COLOR As Long = 8421504 ' Medium Gray, adjust as needed

'==================================================================================================
' PUBLIC FUNCTIONS
'==================================================================================================

Public Function GetCellProcessingInfo(originalValue As Variant, sourceNumberFormat As String, fieldName As String, categorySheetName As String) As FormattedCellOutput
    Dim out As FormattedCellOutput
    Dim ragicType As String

    ' Get the Ragic field type
    ragicType = RagicDictionary.GetFieldRagicType(categorySheetName, fieldName)
    Log "GetCellProcessingInfo", "Category: '" & categorySheetName & "', Field: '" & fieldName & "', OriginalValue: '" & CStr(originalValue) & "', RagicType: '" & ragicType & "'", DEBUG_LEVEL, "GetCellProcessingInfo", "DataFormatter"

    ' Initialize default output values
    out.IsSectionHeader = False
    out.FinalValue = originalValue
    out.NumberFormatString = "@" ' Default to Text format

    ' --- Section Specific Styling Defaults ---
    out.SectionFontName = "" ' Use default font unless specified
    out.SectionFontSizeOffset = 0
    out.SectionFontColor = vbBlack ' Default font color
    out.SectionFontBold = False

    ' Apply formatting based on Ragic type
    Select Case ragicType
        Case "Section"
            out.FinalValue = fieldName ' Section header displays the field name
            out.NumberFormatString = "@" ' Text format for header
            out.IsSectionHeader = True
            
            ' Section specific styling
            out.SectionFontBold = True
            out.SectionFontSizeOffset = 3 ' Font size +3 relative to base
            out.SectionFontColor = SECTION_HEADER_DEFAULT_FONT_COLOR ' Lighter font color

        Case "Date"
            out.FinalValue = originalValue ' Assume originalValue is a valid date or can be coerced by Excel
            out.NumberFormatString = "DD/MM/YYYY"
            ' If originalValue is text, Excel might need help. Consider CDate conversion if issues arise.
            ' For example: If IsDate(originalValue) Then out.FinalValue = CDate(originalValue) Else out.FinalValue = originalValue

        Case "Number"
            ' Convertir explicitement en nombre en gérant le séparateur décimal
            If IsNumeric(Replace(CStr(originalValue), ".", Application.DecimalSeparator)) Then
                out.FinalValue = CDbl(Replace(CStr(originalValue), ".", Application.DecimalSeparator))
            Else
                out.FinalValue = originalValue
            End If
            out.NumberFormatString = "General" ' Permet d'utiliser les paramètres locaux pour les séparateurs
            Log "GetCellProcessingInfo", "Conversion numérique: Original='" & CStr(originalValue) & "' -> Final='" & CStr(out.FinalValue) & "'", DEBUG_LEVEL, "GetCellProcessingInfo", "DataFormatter"

        Case "Text"
            ' Default is already Text format ("@") and original value
            out.FinalValue = CStr(originalValue) ' Ensure it's a string
            out.NumberFormatString = "@"
            
        Case Else ' Includes any unknown types, treat as Text
            Log "GetCellProcessingInfo", "Unknown RagicType: '" & ragicType & "' for field '" & fieldName & "'. Defaulting to Text.", WARNING_LEVEL, "GetCellProcessingInfo", "DataFormatter"
            out.FinalValue = CStr(originalValue) ' Ensure it's a string
            out.NumberFormatString = "@"
    End Select
    
    Log "GetCellProcessingInfo", "Output for '" & fieldName & "': FinalValue='" & CStr(out.FinalValue) & "', NumberFormat='" & out.NumberFormatString & "', IsSection=" & out.IsSectionHeader, DEBUG_LEVEL, "GetCellProcessingInfo", "DataFormatter"
    GetCellProcessingInfo = out
End Function
