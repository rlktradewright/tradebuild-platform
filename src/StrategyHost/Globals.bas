Attribute VB_Name = "Globals"
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                            As String = "StrategyHost27"

Private Const ModuleName                            As String = "Globals"

'@================================================================================
' Member variables
'@================================================================================

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
    gApplyThemeToControl pTheme, lControl
Next
        
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gApplyThemeToControl(ByVal pTheme As ITheme, ByVal pControl As Control)
Const ProcName As String = "gApplyThemeToControl"
On Error GoTo Err

If TypeOf pControl Is Label Then
    pControl.Appearance = pTheme.Appearance
    pControl.BackColor = pTheme.BackColor
    pControl.ForeColor = pTheme.ForeColor
ElseIf TypeOf pControl Is CheckBox Or _
    TypeOf pControl Is Frame Or _
    TypeOf pControl Is OptionButton _
Then
    SetWindowThemeOff pControl.hWnd
    pControl.Appearance = pTheme.Appearance
    pControl.BackColor = pTheme.BackColor
    pControl.ForeColor = pTheme.ForeColor
ElseIf TypeOf pControl Is PictureBox Then
    pControl.Appearance = pTheme.Appearance
    pControl.BorderStyle = pTheme.BorderStyle
    pControl.BackColor = pTheme.BackColor
    pControl.ForeColor = pTheme.ForeColor
ElseIf TypeOf pControl Is TextBox Then
    pControl.Appearance = pTheme.Appearance
    pControl.BorderStyle = pTheme.BorderStyle
    pControl.BackColor = pTheme.TextBackColor
    pControl.ForeColor = pTheme.TextForeColor
    If Not pTheme.TextFont Is Nothing Then
        Set pControl.Font = pTheme.TextFont
    ElseIf Not pTheme.BaseFont Is Nothing Then
        Set pControl.Font = pTheme.BaseFont
    End If
ElseIf TypeOf pControl Is ComboBox Or _
    TypeOf pControl Is ListBox _
Then
    pControl.Appearance = pTheme.Appearance
    pControl.BackColor = pTheme.TextBackColor
    pControl.ForeColor = pTheme.TextForeColor
    If Not pTheme.ComboFont Is Nothing Then
        Set pControl.Font = pTheme.ComboFont
    ElseIf Not pTheme.BaseFont Is Nothing Then
        Set pControl.Font = pTheme.BaseFont
    End If
ElseIf TypeOf pControl Is CommandButton Or _
    TypeOf pControl Is Shape _
Then
    ' nothing for these
ElseIf TypeOf pControl Is Toolbar Then
    pControl.Appearance = pTheme.Appearance
    pControl.BorderStyle = pTheme.BorderStyle
    
    If pControl.Style = tbrStandard Then
        Dim lDoneFirstStandardToolbar As Boolean
        If Not lDoneFirstStandardToolbar Then
            lDoneFirstStandardToolbar = True
            SetToolbarColor pControl, pTheme.ToolbarBackColor
        End If
    Else
        Dim lDoneFirstFlatToolbar As Boolean
        If Not lDoneFirstFlatToolbar Then
            lDoneFirstFlatToolbar = True
            SetToolbarColor pControl, pTheme.ToolbarBackColor
        End If
    End If
    pControl.Refresh
ElseIf TypeOf pControl Is SSTab Then
    pControl.BackColor = pTheme.TabstripBackColor
    pControl.ForeColor = pTheme.TabstripForeColor
ElseIf TypeOf pControl Is Object  Then
    On Error Resume Next
    If TypeOf pControl.object Is IThemeable Then
        If Err.Number = 0 Then
            On Error GoTo Err
            Dim lThemeable As IThemeable
            Set lThemeable = pControl.object
            lThemeable.Theme = pTheme
        Else
            On Error GoTo Err
        End If
    Else
        On Error GoTo Err
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gCreateChartStyle() As ChartStyle
Const ProcName As String = "gCreateChartStyle"
On Error GoTo Err

LogMessage "Creating chart style""Black"""

Dim lCrosshairsLineStyle As LineStyle
Set lCrosshairsLineStyle = New LineStyle
lCrosshairsLineStyle.Color = 128
lCrosshairsLineStyle.LineStyle = LineSolid
lCrosshairsLineStyle.Thickness = 1

Dim lDefaultRegionStyle As ChartRegionStyle
Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone
lDefaultRegionStyle.IntegerYScale = False
lDefaultRegionStyle.YScaleQuantum = 0.015625
lDefaultRegionStyle.YGridlineSpacing = 1.8
lDefaultRegionStyle.minimumHeight = 0.015625
lDefaultRegionStyle.CursorSnapsToTickBoundaries = True
lDefaultRegionStyle.BackGradientFillColors = gCreateColorArray(2105376, 2105376)

lDefaultRegionStyle.XGridLineStyle = GetDefaultLineStyle.Clone
lDefaultRegionStyle.XGridLineStyle.Color = 3158064
lDefaultRegionStyle.YGridLineStyle = lDefaultRegionStyle.XGridLineStyle
lDefaultRegionStyle.SessionEndGridLineStyle = lDefaultRegionStyle.XGridLineStyle
lDefaultRegionStyle.SessionStartGridLineStyle = lDefaultRegionStyle.SessionEndGridLineStyle.Clone
lDefaultRegionStyle.SessionStartGridLineStyle.Thickness = 3

Dim lXAxisStyle As ChartRegionStyle
Set lXAxisStyle = GetDefaultChartXAxisRegionStyle.Clone
lXAxisStyle.CursorTextPosition = CursorTextPositionBelowLeftCursor
lXAxisStyle.XGridTextPosition = XGridTextPositionBottom
lXAxisStyle.XGridTextStyle.Box = True
lXAxisStyle.XGridTextStyle.BoxFillWithBackgroundColor = True
lXAxisStyle.XGridTextStyle.BoxStyle = LineInvisible
lXAxisStyle.XGridTextStyle.Color = 13684944
lXAxisStyle.YScaleQuantum = 0.0001
lXAxisStyle.BackGradientFillColors = gCreateColorArray(0, 0)

lXAxisStyle.XCursorTextStyle.BoxFillWithBackgroundColor = True
lXAxisStyle.XCursorTextStyle.BoxStyle = LineInvisible
lXAxisStyle.XCursorTextStyle.BoxThickness = 0
lXAxisStyle.XCursorTextStyle.Color = 255

Dim lFont As New StdFont
lFont.Name = "Courier New"
lFont.Bold = True
lFont.Italic = False
lFont.Size = 8.25
lFont.Strikethrough = False
lFont.Underline = False
lXAxisStyle.XCursorTextStyle.Font = lFont

Dim lYAxisStyle As ChartRegionStyle
Set lYAxisStyle = GetDefaultChartYAxisRegionStyle.Clone
lYAxisStyle.BackGradientFillColors = gCreateColorArray(0, 0)
lYAxisStyle.YCursorTextStyle.Font = lFont
lYAxisStyle.YGridTextPosition = YGridTextPositionLeft
lYAxisStyle.YGridTextStyle.Box = True
lYAxisStyle.YGridTextStyle.BoxFillWithBackgroundColor = True
lYAxisStyle.YGridTextStyle.BoxStyle = LineInvisible
lYAxisStyle.YGridTextStyle.Color = 13684944
lYAxisStyle.YCursorTextStyle.BoxFillWithBackgroundColor = True
lYAxisStyle.YCursorTextStyle.BoxStyle = LineInvisible
lYAxisStyle.YCursorTextStyle.BoxThickness = 0
lYAxisStyle.YCursorTextStyle.Color = 255

Set gCreateChartStyle = ChartStylesManager.Add("Black", _
                        ChartStylesManager.DefaultStyle, _
                        lDefaultRegionStyle, _
                        lXAxisStyle, _
                        lYAxisStyle.Clone, _
                        lCrosshairsLineStyle, _
                        pTemporary:=True)

gCreateChartStyle.Autoscrolling = True
gCreateChartStyle.ChartBackColor = 2105376
gCreateChartStyle.HorizontalMouseScrollingAllowed = True
gCreateChartStyle.HorizontalScrollBarVisible = False
gCreateChartStyle.VerticalMouseScrollingAllowed = True
gCreateChartStyle.XAxisVisible = True
gCreateChartStyle.YAxisVisible = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateColorArray(ParamArray pColors()) As Long()
ReDim lColors(UBound(pColors)) As Long
Dim i As Long
For i = 0 To UBound(lColors)
    lColors(i) = CLng(pColors(i))
Next
gCreateColorArray = lColors
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

Public Sub gLog(ByRef pMsg As String, _
                ByRef pModName As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLog"
On Error GoTo Err

Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("strategyhost", ProjectName)

sLogger.Log pMsg, pProcName, pModName, pLogLevel, pMsgQualifier

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Public Sub SetToolbarColor(ByVal pToolbar As Toolbar, ByVal pColor As Long)
Dim lBrush As Long
lBrush = CreateSolidBrush(NormalizeColor(pColor))

Dim lhWnd As Long
Select Case pToolbar.Style
Case ToolbarStyleConstants.tbrFlat
    lhWnd = pToolbar.hWnd
Case ToolbarStyleConstants.tbrStandard
    lhWnd = FindWindowEx(pToolbar.hWnd, 0, "msvb_lib_toolbar", vbNullString)
End Select

Dim lResult As Long
lResult = SetClassLong(lhWnd, GCLP_HBRBACKGROUND, lBrush)
End Sub

Public Sub SetWindowThemeOff(ByVal phWnd As Long)
Const ProcName As String = "SetWindowThemeOff"
On Error GoTo Err

Dim result As Long
result = SetWindowTheme(phWnd, vbNullString, "")
If result <> 0 Then LogMessage "Error " & result & " setting window theme off"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




