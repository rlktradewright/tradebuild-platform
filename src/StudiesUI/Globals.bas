Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023
Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const RegionNameCustom As String = "$custom"
Public Const RegionNameDefault As String = "$default"
Public Const RegionNamePrice As String = "Price"
Public Const RegionNameVolume As String = "Volume"

Public Const LB_SETHORZEXTENT = &H194

Public Const TaskTypeStartStudy As Long = 1
Public Const TaskTypeReplayBars As Long = 2
Public Const TaskTypeAddValueListener As Long = 3

Public Const ErroredFieldColor As Long = &HD0CAFA

Public Const PositiveChangeBackColor As Long = &HB7E43
Public Const NegativeChangebackColor As Long = &H4444EB

Public Const PositiveProfitColor As Long = &HB7E43
Public Const NegativeProfitColor As Long = &H4444EB

Public Const IncreasedValueColor As Long = &HB7E43
Public Const DecreasedValueColor As Long = &H4444EB

Public Const NullColor As Long = SystemColorConstants.vbApplicationWorkspace

Public Const BarModeBar As String = "Bars"
Public Const BarModeCandle As String = "Candles"
Public Const BarModeSolidCandle As String = "Solid candles"
Public Const BarModeLine As String = "Line"

Public Const BarStyleNarrow As String = "Narrow"
Public Const BarStyleMedium As String = "Medium"
Public Const BarStyleWide As String = "Wide"

Public Const BarWidthNarrow As Single = 0.3
Public Const BarWidthMedium As Single = 0.6
Public Const BarWidthWide As Single = 0.9

Public Const HistogramStyleNarrow As String = "Narrow"
Public Const HistogramStyleMedium As String = "Medium"
Public Const HistogramStyleWide As String = "Wide"

Public Const HistogramWidthNarrow As Single = 0.3
Public Const HistogramWidthMedium As Single = 0.6
Public Const HistogramWidthWide As Single = 0.9

Public Const LineDisplayModePlain As String = "Plain"
Public Const LineDisplayModeArrowEnd As String = "End arrow"
Public Const LineDisplayModeArrowStart As String = "Start arrow"
Public Const LineDisplayModeArrowBoth As String = "Both arrows"

Public Const LineStyleSolid As String = "Solid"
Public Const LineStyleDash As String = "Dash"
Public Const LineStyleDot As String = "Dot"
Public Const LineStyleDashDot As String = "Dash dot"
Public Const LineStyleDashDotDot As String = "Dash dot dot"
Public Const LineStyleInsideSolid As String = "Inside solid"
Public Const LineStyleInvisible As String = "Invisible"

Public Const PointDisplayModeLine As String = "Line"
Public Const PointDisplayModePoint As String = "Point"
Public Const PointDisplayModeSteppedLine As String = "Stepped line"
Public Const PointDisplayModeHistogram As String = "Histogram"

Public Const PointStyleRound As String = "Round"
Public Const PointStyleSquare As String = "Square"

Public Const TextDisplayModePlain As String = "Plain"
Public Const TextDisplayModeWIthBackground As String = "With background"
Public Const TextDisplayModeWithBox As String = "With box"
Public Const TextDisplayModeWithFilledBox As String = "With filled box"

Public Const CustomStyle As String = "(Custom)"
Public Const CustomDisplayMode As String = "(Custom)"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mDefaultStudyConfigurations As Collection

Public gCustColors(15) As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Function gLineStyleToString( _
                ByVal value As LineStyles) As String
Select Case value
Case LineSolid
    gLineStyleToString = LineStyleSolid
Case LineDash
    gLineStyleToString = LineStyleDash
Case LineDot
    gLineStyleToString = LineStyleDot
Case LineDashDot
    gLineStyleToString = LineStyleDashDot
Case LineDashDotDot
    gLineStyleToString = LineStyleDashDotDot
Case LineInvisible
    gLineStyleToString = LineStyleInvisible
Case LineInsideSolid
    gLineStyleToString = LineStyleInsideSolid
End Select
End Function

Public Function gPointStyleToString( _
                ByVal value As PointStyles) As String
Select Case value
Case PointRound
    gPointStyleToString = PointStyleRound
Case PointSquare
    gPointStyleToString = PointStyleSquare
End Select
End Function

Public Sub filterNonNumericKeyPress(ByRef KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) Then
    KeyAscii = 0
End If
End Sub

Public Function isInteger( _
                ByVal value As String, _
                Optional ByVal minValue As Long = 0, _
                Optional ByVal maxValue As Long = &H7FFFFFFF) As Boolean
Dim quantity As Long

On Error GoTo err

If IsNumeric(value) Then
    quantity = CLng(value)
    If CDbl(value) - quantity = 0 Then
        If quantity >= minValue And quantity <= maxValue Then
            isInteger = True
        End If
    End If
End If
                
Exit Function

err:
If err.Number <> VBErrorCodes.VbErrOverflow Then err.Raise err.Number
End Function

Public Function isPrice( _
                ByVal value As String, _
                ByVal ticksize As Double) As Boolean
Dim theVal As Double

On Error GoTo err

If IsNumeric(value) Then
    theVal = value
    If theVal > 0 And _
        Int(theVal / ticksize) * ticksize = theVal _
    Then
        isPrice = True
    End If
End If

Exit Function

err:
If err.Number <> VBErrorCodes.VbErrOverflow Then err.Raise err.Number
End Function

Public Function loadDefaultStudyConfiguration( _
                ByVal name As String, _
                ByVal spName As String) As studyConfiguration
Dim sc As studyConfiguration
If mDefaultStudyConfigurations Is Nothing Then
    Set loadDefaultStudyConfiguration = Nothing
Else
    On Error Resume Next
    Set sc = mDefaultStudyConfigurations.item(calcDefaultStudyKey(name, spName))
    On Error GoTo 0
    If Not sc Is Nothing Then Set loadDefaultStudyConfiguration = sc.Clone
End If
End Function

Public Sub notImplemented()
MsgBox "This facility has not yet been implemented", , "Sorry"
End Sub

Public Sub updateDefaultStudyConfiguration( _
                ByVal value As studyConfiguration)
Dim sc As studyConfiguration

If mDefaultStudyConfigurations Is Nothing Then
    Set mDefaultStudyConfigurations = New Collection
End If
On Error Resume Next
mDefaultStudyConfigurations.Remove calcDefaultStudyKey(value.name, value.StudyLibraryName)
On Error GoTo 0

Set sc = value.Clone
sc.underlyingStudy = Nothing
mDefaultStudyConfigurations.Add sc, calcDefaultStudyKey(value.name, value.StudyLibraryName)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcDefaultStudyKey( _
                ByVal studyName As String, _
                ByVal StudyLibraryName As String) As String
calcDefaultStudyKey = "$$" & studyName & "$$" & StudyLibraryName & "$$"
End Function



