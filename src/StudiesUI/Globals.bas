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

Public Sub gApplyTheme(ByVal pTheme As ITheme, ByVal pControls As Object)
Const ProcName As String = "gApplyTheme"
On Error GoTo Err

If pTheme Is Nothing Then Exit Sub

Dim lControl As Control
For Each lControl In pControls
    If TypeOf lControl Is Label Then
        lControl.Appearance = pTheme.Appearance
        lControl.BackColor = pTheme.BackColor
        lControl.ForeColor = pTheme.ForeColor
    ElseIf TypeOf lControl Is CheckBox Or _
        TypeOf lControl Is Frame Or _
        TypeOf lControl Is OptionButton _
    Then
        SetWindowThemeOff lControl.hWnd
        lControl.Appearance = pTheme.Appearance
        lControl.BackColor = pTheme.BackColor
        lControl.ForeColor = pTheme.ForeColor
    ElseIf TypeOf lControl Is PictureBox Then
        lControl.Appearance = pTheme.Appearance
        lControl.BorderStyle = pTheme.BorderStyle
        lControl.BackColor = pTheme.BackColor
        lControl.ForeColor = pTheme.ForeColor
    ElseIf TypeOf lControl Is TextBox Then
        lControl.Appearance = pTheme.Appearance
        lControl.BorderStyle = pTheme.BorderStyle
        lControl.BackColor = pTheme.TextBackColor
        lControl.ForeColor = pTheme.TextForeColor
        If Not pTheme.TextFont Is Nothing Then
            Set lControl.Font = pTheme.TextFont
        ElseIf Not pTheme.BaseFont Is Nothing Then
            Set lControl.Font = pTheme.BaseFont
        End If
    ElseIf TypeOf lControl Is ComboBox Or _
        TypeOf lControl Is ListBox _
    Then
        lControl.Appearance = pTheme.Appearance
        lControl.BackColor = pTheme.TextBackColor
        lControl.ForeColor = pTheme.TextForeColor
        If Not pTheme.ComboFont Is Nothing Then
            Set lControl.Font = pTheme.ComboFont
        ElseIf Not pTheme.BaseFont Is Nothing Then
            Set lControl.Font = pTheme.BaseFont
        End If
    ElseIf TypeOf lControl Is CommandButton Or _
        TypeOf lControl Is Shape _
    Then
        ' nothing for these
    ElseIf TypeOf lControl Is Object  Then
        On Error Resume Next
        If TypeOf lControl.object Is IThemeable Then
            If Err.Number = 0 Then
                On Error GoTo Err
                Dim lThemeable As IThemeable
                Set lThemeable = lControl.object
                lThemeable.Theme = pTheme
            Else
                On Error GoTo Err
            End If
        Else
            On Error GoTo Err
        End If
    End If
Next
        
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

Public Sub gFilterNonNumericKeyPress(ByRef KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) Then
    KeyAscii = 0
End If
End Sub

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

Public Sub gNotImplemented()
MsgBox "This facility has not yet been implemented", , "Sorry"
End Sub

Public Function gPointStyleToString( _
                ByVal value As PointStyles) As String
Select Case value
Case PointRound
    gPointStyleToString = PointStyleRound
Case PointSquare
    gPointStyleToString = PointStyleSquare
End Select
End Function

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub SetWindowThemeOff(ByVal phWnd As Long)
Dim result As Long
result = SetWindowTheme(phWnd, vbNullString, "")
End Sub


