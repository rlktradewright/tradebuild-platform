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

Public Const ProjectName                                    As String = "StudiesUI27"
Private Const ModuleName                As String = "Globals"


Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023
Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023

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

Public Function gChooseAColor( _
                ByVal pInitialColor As Long, _
                ByVal pAllowNull As Boolean, _
                ByVal pParent As Form) As Long
Static lSimpleColorPicker As fSimpleColorPicker
Dim cursorpos As GDI_POINT

Const ProcName As String = "gChooseAColor"
On Error GoTo Err

GetCursorPos cursorpos

If lSimpleColorPicker Is Nothing Then Set lSimpleColorPicker = New fSimpleColorPicker

lSimpleColorPicker.Top = cursorpos.Y * Screen.TwipsPerPixelY
lSimpleColorPicker.Left = cursorpos.X * Screen.TwipsPerPixelX
lSimpleColorPicker.initialColor = pInitialColor
If pAllowNull Then lSimpleColorPicker.NoColorButton.Enabled = True
lSimpleColorPicker.ZOrder 0
lSimpleColorPicker.Show vbModal, pParent
gChooseAColor = lSimpleColorPicker.selectedColor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

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

Public Function gLogger() As Logger
Dim lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("log")
Set gLogger = lLogger
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

Const ProcName As String = "isInteger"
On Error GoTo Err

If IsNumeric(value) Then
    quantity = CLng(value)
    If CDbl(value) - quantity = 0 Then
        If quantity >= minValue And quantity <= maxValue Then
            isInteger = True
        End If
    End If
End If
                
Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrOverflow Then Exit Function
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function isPrice( _
                ByVal value As String, _
                ByVal ticksize As Double) As Boolean
Dim theVal As Double

Const ProcName As String = "isPrice"
On Error GoTo Err

If IsNumeric(value) Then
    theVal = value
    If theVal > 0 And _
        Int(theVal / ticksize) * ticksize = theVal _
    Then
        isPrice = True
    End If
End If

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrOverflow Then Exit Function
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub notImplemented()
MsgBox "This facility has not yet been implemented", , "Sorry"
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



