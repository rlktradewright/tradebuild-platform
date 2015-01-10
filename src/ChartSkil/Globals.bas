Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ProjectName                As String = "ChartSkil27"
Public Const ModuleName                 As String = "Globals"

Public Const Pi As Double = 3.14159265358979

Public Const MinusInfinityDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023
Public Const PlusInfinityDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const MinusInfinitySingle As Single = -(2 - 2 ^ -23) * 2 ^ 127
Public Const PlusInfinitySingle As Single = (2 - 2 ^ -23) * 2 ^ 127

Public Const PlusInfinityLong As Long = &H7FFFFFFF
Public Const MinusInfinityLong As Long = &H80000000

Public Const Log10 As Double = 2.30258509299405

Public Const OneMicroSecond As Double = 1.15740740740741E-11

Public Const MaxSystemColor As Long = &H80000018

Public Const GridlineSpacingCm As Double = 1.8

Public Const HitTestTolerancePixels As Long = 3

Public Const TwipsPerCm As Double = 1440 / 2.54

Public Const ConfigSettingName                  As String = "&Name"
Public Const ConfigSettingStyleType             As String = "&StyleType"

Public Const ToolbarCommandAutoScale As String = "autoscale"
Public Const ToolbarCommandAutoscroll As String = "autoscroll"

Public Const ToolbarCommandIncreaseSpacing As String = "increasespacing"
Public Const ToolbarCommandReduceSpacing As String = "reducespacing"

Public Const ToolbarCommandScaleDown As String = "scaledown"
Public Const ToolbarCommandScaleUp As String = "scaleup"

Public Const ToolbarCommandScrollDown As String = "scrolldown"
Public Const ToolbarCommandScrollEnd As String = "scrollend"
Public Const ToolbarCommandScrollLeft As String = "scrollleft"
Public Const ToolbarCommandScrollRight As String = "scrollright"
Public Const ToolbarCommandScrollUp As String = "scrollup"

Public Const ToolbarCommandShowBars As String = "showbars"
Public Const ToolbarCommandShowCandlesticks As String = "showcandlesticks"
Public Const ToolbarCommandShowLine As String = "showline"
Public Const ToolbarCommandShowCrosshair As String = "showcrosshair"
Public Const ToolbarCommandShowPlainCursor As String = "showplaincursor"
Public Const ToolbarCommandShowDiscCursor As String = "showdisccursor"

Public Const ToolbarCommandThickerBars As String = "thickerbars"
Public Const ToolbarCommandThinnerBars As String = "thinnerbars"

'================================================================================
' Enums
'================================================================================

Public Type PeriodLabels
    Label As String
    CoarseLabel As String
End Type

Public Enum RegionTypes
    RegionTypeData
    RegionTypeXAxis
    RegionTypeYAxis
    RegionTypeBackground
End Enum

Public Enum VerticalGridlineTypes
    VerticalGridlineTypeNone
    VerticalGridlineTypePeriodBoundary
    VerticalGridlineTypeSessionEnd
    VerticalGridlineTypeSessionStart
    VerticalGridlineTypeDateChange
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mIsInDev As Boolean

Public gBlankCursor As IPictureDisp
Public gSelectorCursor As IPictureDisp

'================================================================================
' Properties
'================================================================================

Public Property Get gIsInDev() As Boolean
gIsInDev = mIsInDev
End Property

Public Property Get gDefaultBarStyle() As BarStyle
Static lStyle As BarStyle
If lStyle Is Nothing Then
    Set lStyle = New BarStyle
    lStyle.Color = -1
    lStyle.DisplayMode = BarDisplayModeCandlestick
    lStyle.DownColor = vbBlack
    lStyle.IncludeInAutoscale = True
    lStyle.layer = LayerLowestUser
    lStyle.OutlineThickness = 1
    lStyle.SolidUpBody = False
    lStyle.TailThickness = 1
    lStyle.Thickness = 2
    lStyle.UpColor = vbBlack
    lStyle.Width = 0.6
End If
Set gDefaultBarStyle = lStyle
End Property

Public Property Get gDefaultChartDataRegionStyle() As ChartRegionStyle
Static lStyle As ChartRegionStyle

Const DefaultGridlineColor As Long = &HF0F0F0

If lStyle Is Nothing Then
    Set lStyle = New ChartRegionStyle
    lStyle.Autoscaling = True
    lStyle.BackGradientFillColors = gCreateColorArray(vbWhite, vbWhite)
    lStyle.CursorSnapsToTickBoundaries = True
    lStyle.CursorTextMode = CursorTextModeNone
    lStyle.CursorTextPosition = CursorTextPositionBelowLeftCursor
    lStyle.CursorTextStyle = Nothing
    lStyle.HasXGrid = True
    lStyle.HasYGrid = True
    lStyle.HasXGridText = False
    lStyle.HasYGridText = False
    lStyle.IntegerYScale = False
    lStyle.MinimumHeight = 0.015625
    
    lStyle.SessionEndGridLineStyle = New LineStyle
    lStyle.SessionEndGridLineStyle.Color = DefaultGridlineColor
    lStyle.SessionEndGridLineStyle.LineStyle = LineDash
    
    lStyle.SessionStartGridLineStyle = New LineStyle
    lStyle.SessionStartGridLineStyle.Color = DefaultGridlineColor
    lStyle.SessionStartGridLineStyle.Thickness = 3
    
    lStyle.XCursorTextPosition = CursorTextPositionBottom
    lStyle.XCursorTextStyle = Nothing
    
    lStyle.XGridLineStyle = New LineStyle
    lStyle.XGridLineStyle.Color = DefaultGridlineColor
    
    lStyle.XGridTextPosition = XGridTextPositionBottom
    lStyle.XGridTextStyle = Nothing
    
    lStyle.YCursorTextPosition = CursorTextPositionLeft
    lStyle.YCursorTextStyle = Nothing
    lStyle.YGridlineSpacing = 1.8
    
    lStyle.YGridLineStyle = New LineStyle
    lStyle.YGridLineStyle.Color = DefaultGridlineColor
    
    lStyle.YGridTextPosition = YGridTextPositionLeft
    lStyle.YGridTextStyle = Nothing
    
    lStyle.YScaleQuantum = 0.015625
    
End If
Set gDefaultChartDataRegionStyle = lStyle
End Property

Public Property Get gDefaultChartXAxisRegionStyle() As ChartRegionStyle
Static lStyle As ChartRegionStyle

Dim afont As StdFont

If lStyle Is Nothing Then
    Set lStyle = New ChartRegionStyle
    lStyle.Autoscaling = True
    lStyle.BackGradientFillColors = gCreateColorArray(vbWhite, vbWhite)
    lStyle.CursorSnapsToTickBoundaries = False
    lStyle.CursorTextMode = CursorTextModeXOnly
    lStyle.CursorTextPosition = CursorTextPositionBelowLeftCursor
    lStyle.CursorTextStyle = Nothing
    lStyle.HasXGrid = False
    lStyle.HasYGrid = False
    lStyle.HasXGridText = True
    lStyle.HasYGridText = False
    lStyle.IntegerYScale = False
    
    lStyle.SessionEndGridLineStyle = Nothing
    
    lStyle.SessionStartGridLineStyle = Nothing
    
    lStyle.XCursorTextPosition = CursorTextPositionTop
    lStyle.XCursorTextStyle = New TextStyle
    lStyle.XCursorTextStyle.Box = True
    lStyle.XCursorTextStyle.BoxFillColor = vbWhite
    'lStyle.XCursorTextStyle.Offset = gNewSize(0, -0.1, CoordsDistance, CoordsDistance)
    
    lStyle.XGridLineStyle = Nothing
    
    lStyle.XGridTextPosition = XGridTextPositionBottom
    lStyle.XGridTextStyle = New TextStyle
    lStyle.XGridTextStyle.Box = True
    lStyle.XGridTextStyle.BoxFillWithBackgroundColor = True
    lStyle.XGridTextStyle.BoxStyle = LineInvisible
    'lStyle.XGridTextStyle.Offset = gNewSize(0#, 0.1, CoordsDistance, CoordsDistance)
    
    lStyle.YCursorTextPosition = CursorTextPositionLeft
    lStyle.YCursorTextStyle = Nothing
    
    lStyle.YGridlineSpacing = 1.8
    
    lStyle.YGridLineStyle = Nothing
    
    lStyle.YGridTextPosition = YGridTextPositionLeft
    lStyle.YGridTextStyle = Nothing
    
End If
Set gDefaultChartXAxisRegionStyle = lStyle
End Property

Public Property Get gDefaultChartYAxisRegionStyle() As ChartRegionStyle
Static lStyle As ChartRegionStyle

Dim afont As StdFont

If lStyle Is Nothing Then
    Set lStyle = New ChartRegionStyle
    lStyle.Autoscaling = False
    lStyle.BackGradientFillColors = gCreateColorArray(vbWhite, vbWhite)
    lStyle.CursorSnapsToTickBoundaries = True
    lStyle.CursorTextMode = CursorTextModeYOnly
    lStyle.CursorTextPosition = CursorTextPositionBelowLeftCursor
    lStyle.CursorTextStyle = Nothing
    lStyle.HasXGrid = False
    lStyle.HasYGrid = False
    lStyle.HasXGridText = False
    lStyle.HasYGridText = True
    lStyle.IntegerYScale = False
    
    lStyle.SessionEndGridLineStyle = Nothing
    
    lStyle.SessionStartGridLineStyle = Nothing
    
    lStyle.XCursorTextPosition = CursorTextPositionTop
    lStyle.XCursorTextStyle = Nothing
    
    lStyle.XGridLineStyle = Nothing
    
    lStyle.XGridTextPosition = XGridTextPositionBottom
    lStyle.XGridTextStyle = Nothing
    
    lStyle.YCursorTextPosition = CursorTextPositionLeft
    lStyle.YCursorTextStyle = New TextStyle
    lStyle.YCursorTextStyle.Box = True
    lStyle.YCursorTextStyle.BoxFillColor = vbWhite
    lStyle.YCursorTextStyle.Offset = gNewSize(0.1, 0#, CoordsDistance, CoordsDistance)
    
    lStyle.YGridlineSpacing = 1.8
    
    lStyle.YGridLineStyle = Nothing
    
    lStyle.YGridTextPosition = YGridTextPositionLeft
    lStyle.YGridTextStyle = New TextStyle
    lStyle.YGridTextStyle.Box = True
    lStyle.YGridTextStyle.BoxFillWithBackgroundColor = True
    lStyle.YGridTextStyle.BoxStyle = LineInvisible
    lStyle.YGridTextStyle.Offset = gNewSize(0.2, 0#, CoordsDistance, CoordsDistance)
    
End If
Set gDefaultChartYAxisRegionStyle = lStyle
End Property

Public Property Get gDefaultChartStudyRegionStyle() As ChartRegionStyle
Static lStyle As ChartRegionStyle

Dim afont As StdFont

If lStyle Is Nothing Then
    Set lStyle = New ChartRegionStyle
    lStyle.Autoscaling = True
    lStyle.BackGradientFillColors = gCreateColorArray(vbWhite, vbWhite)
    lStyle.CursorSnapsToTickBoundaries = False
    lStyle.CursorTextMode = CursorTextModeNone
    lStyle.CursorTextPosition = CursorTextPositionBelowLeftCursor
    lStyle.CursorTextStyle = Nothing
    lStyle.HasXGrid = True
    lStyle.HasYGrid = True
    lStyle.HasXGridText = False
    lStyle.HasYGridText = False
    lStyle.IntegerYScale = False
    lStyle.MinimumHeight = 0.0001
    
    lStyle.SessionEndGridLineStyle = New LineStyle
    lStyle.SessionEndGridLineStyle.Color = &HC0C0C0
    lStyle.SessionEndGridLineStyle.LineStyle = LineDash
    
    lStyle.SessionStartGridLineStyle = New LineStyle
    lStyle.SessionStartGridLineStyle.Color = &HC0C0C0
    lStyle.SessionStartGridLineStyle.Thickness = 3
    
    lStyle.XCursorTextPosition = CursorTextPositionBottom
    lStyle.XCursorTextStyle = Nothing
    
    lStyle.XGridLineStyle = New LineStyle
    lStyle.XGridLineStyle.Color = &HC0C0C0
    
    lStyle.XGridTextPosition = XGridTextPositionBottom
    lStyle.XGridTextStyle = Nothing
    
    lStyle.YCursorTextPosition = CursorTextPositionLeft
    lStyle.YCursorTextStyle = Nothing
    lStyle.YGridlineSpacing = 0.9
    
    lStyle.YGridLineStyle = New LineStyle
    lStyle.YGridLineStyle.Color = &HC0C0C0
    
    lStyle.YGridTextPosition = YGridTextPositionLeft
    lStyle.YGridTextStyle = Nothing
    
    lStyle.YScaleQuantum = 0.0001
    
End If
Set gDefaultChartStudyRegionStyle = lStyle
End Property

Public Property Get gDefaultDataPointStyle() As DataPointStyle
Static lStyle As DataPointStyle
If lStyle Is Nothing Then
    Set lStyle = New DataPointStyle
    lStyle.Color = vbBlack
    lStyle.DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine
    lStyle.DownColor = -1
    lStyle.HistogramBarWidth = 0.6
    lStyle.IncludeInAutoscale = True
    lStyle.layer = LayerLowestUser + 1
    lStyle.LineStyle = LineStyles.LineSolid
    lStyle.LineThickness = 1
    lStyle.PointStyle = PointRound
    lStyle.UpColor = -1
End If
Set gDefaultDataPointStyle = lStyle
End Property

Public Property Get gDefaultLineStyle() As LineStyle
Static lStyle As LineStyle
If lStyle Is Nothing Then
    Set lStyle = New LineStyle
    lStyle.ArrowEndColor = vbBlack
    lStyle.ArrowEndFillColor = vbBlack
    lStyle.ArrowEndFillStyle = FillStyles.FillSolid
    lStyle.ArrowEndLength = 10
    lStyle.ArrowEndStyle = ArrowStyles.ArrowNone
    lStyle.ArrowEndWidth = 10
    lStyle.ArrowStartColor = vbBlack
    lStyle.ArrowStartFillColor = vbBlack
    lStyle.ArrowStartFillStyle = FillStyles.FillSolid
    lStyle.ArrowStartLength = 10
    lStyle.ArrowStartStyle = ArrowStyles.ArrowNone
    lStyle.ArrowStartWidth = 10
    lStyle.Color = vbBlack
    lStyle.ExtendAfter = False
    lStyle.ExtendBefore = False
    lStyle.Extended = False
    lStyle.FixedX = False
    lStyle.FixedY = False
    lStyle.IncludeInAutoscale = False
    lStyle.layer = LayerHighestUser
    lStyle.LineStyle = LineStyles.LineSolid
    lStyle.Thickness = 1
End If
Set gDefaultLineStyle = lStyle
End Property

Public Property Get gDefaultTextStyle() As TextStyle
Static lStyle As TextStyle
If lStyle Is Nothing Then
    Set lStyle = New TextStyle
    
    Dim afont As New StdFont
    afont.Bold = False
    afont.Italic = False
    afont.Name = "Arial"
    afont.Size = 8
    afont.Strikethrough = False
    afont.Underline = False
    
    lStyle.Angle = 0
    lStyle.Align = TextAlignModes.AlignTopLeft
    lStyle.Box = False
    lStyle.BoxColor = vbBlack
    lStyle.BoxStyle = LineStyles.LineSolid
    lStyle.BoxThickness = 1
    lStyle.BoxFillColor = vbWhite
    lStyle.BoxFillStyle = FillStyles.FillSolid
    lStyle.BoxFillWithBackgroundColor = False
    lStyle.Color = vbBlack
    lStyle.Font = afont
    lStyle.Ellipsis = EllipsisModes.EllipsisNone
    lStyle.ExpandTabs = True
    lStyle.Extended = False
    lStyle.FixedX = False
    lStyle.FixedY = False
    lStyle.HideIfBlank = True
    lStyle.IncludeInAutoscale = False
    lStyle.Justification = TextJustifyModes.JustifyLeft
    lStyle.layer = LayerHighestUser
    lStyle.MultiLine = False
    lStyle.PaddingX = 0.3
    lStyle.PaddingY = 0#
    lStyle.TabWidth = 8
    lStyle.WordWrap = True
End If
Set gDefaultTextStyle = lStyle
End Property

Public Property Get gErrorLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("error")
Set gErrorLogger = lLogger
End Property

Public Property Get gGraphicObjectStyleManager() As GraphicObjectStyleManager
Static gosm As GraphicObjectStyleManager
If gosm Is Nothing Then Set gosm = New GraphicObjectStyleManager
Set gGraphicObjectStyleManager = gosm
End Property

Public Property Get gLogger() As FormattingLogger
Static lLogger As FormattingLogger
If lLogger Is Nothing Then Set lLogger = CreateFormattingLogger("chartskil.log", ProjectName)
Set gLogger = lLogger
End Property

Public Property Get gTracer() As Tracer
Static lTracer As Tracer
If lTracer Is Nothing Then Set lTracer = GetTracer("chartskil")
Set gTracer = lTracer
End Property

'================================================================================
' Methods
'================================================================================

Public Sub gApplyTheme(ByVal pTheme As ITheme, ByVal pControls As Object)
Const ProcName As String = "gApplyTheme"
On Error GoTo Err

If pTheme Is Nothing Then Exit Sub

Dim lControl As Control
For Each lControl In pControls
    If TypeOf lControl Is Label Or _
        TypeOf lControl Is CheckBox Or _
        TypeOf lControl Is Frame Or _
        TypeOf lControl Is OptionButton _
    Then
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
    ElseIf TypeOf lControl Is ComboBox Or _
        TypeOf lControl Is ListBox _
    Then
        lControl.Appearance = pTheme.Appearance
        lControl.BackColor = pTheme.TextBackColor
        lControl.ForeColor = pTheme.TextForeColor
    ElseIf TypeOf lControl Is CommandButton Or _
        TypeOf lControl Is Shape _
    Then
        ' nothing for these
    ElseIf TypeOf lControl Is Toolbar Then
        lControl.Appearance = pTheme.Appearance
        lControl.BorderStyle = pTheme.BorderStyle
        
        If lControl.Style = tbrStandard Then
            Dim lDoneFirstStandardToolbar As Boolean
            If Not lDoneFirstStandardToolbar Then
                lDoneFirstStandardToolbar = True
                SetToolbarColor lControl, pTheme.ToolbarBackColor
            End If
        Else
            Dim lDoneFirstFlatToolbar As Boolean
            If Not lDoneFirstFlatToolbar Then
                lDoneFirstFlatToolbar = True
                SetToolbarColor lControl, pTheme.ToolbarBackColor
            End If
        End If
        lControl.Refresh
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

Public Function gCloneFont( _
                ByVal pFont As StdFont) As StdFont
Dim afont As New StdFont
afont.Bold = pFont.Bold
afont.Charset = pFont.Charset
afont.Italic = pFont.Italic
afont.Name = pFont.Name
afont.Size = pFont.Size
afont.Strikethrough = pFont.Strikethrough
afont.Underline = pFont.Underline
afont.Weight = pFont.Weight
Set gCloneFont = afont
End Function

Public Function gCreateColorArray(ParamArray pColors()) As Long()
ReDim lcolors(UBound(pColors)) As Long
Dim i As Long
For i = 0 To UBound(lcolors)
    lcolors(i) = CLng(pColors(i))
Next
gCreateColorArray = lcolors
End Function

Public Function gDegreesToRadians( _
                ByVal degrees As Double) As Double
gDegreesToRadians = degrees * Pi / 180
End Function

Public Function gGraphicObjectsOverlap(ByVal pObj1 As IGraphicObject, ByVal pObj2 As IGraphicObject) As Boolean
If pObj1 Is Nothing Or pObj2 Is Nothing Then Exit Function
gGraphicObjectsOverlap = RectOverlaps(pObj1.BoundingRectangle, pObj2.BoundingRectangle)
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

Public Sub gIsPositiveLong(ByVal pThis As Object, ByVal pValue As Variant)
Dim lValue As Long
lValue = CLng(pValue)
If lValue <= 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value must be greater than 0"
End If
End Sub

Public Sub gIsPositiveSingle(ByVal pThis As Object, ByVal pValue As Variant)
Dim lValue As Single
lValue = CSng(pValue)
If lValue <= 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value must be greater than 0"
End If
End Sub

Public Function gIsValidColor( _
                ByVal Value As Long) As Boolean
                
If Value > &HFFFFFF Then Exit Function
If Value < 0 And Value > MaxSystemColor Then Exit Function
gIsValidColor = True
End Function

Public Sub gIsValidColorObj(ByVal pThis As Object, ByVal pValue As Variant)
If Not gIsValidColor(pValue) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid RGB color or a system color"
End If
End Sub

Public Function gLongToHexString(ByVal Value As Long) As String
gLongToHexString = "&h" & Hex$(Value)
End Function

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

Public Function gRadiansToDegrees( _
                ByVal radians As Double) As Double
gRadiansToDegrees = radians * 180 / Pi
End Function

Public Function gRegionTypeToString( _
                ByVal Value As RegionTypes) As String
Select Case Value
Case RegionTypeData
    gRegionTypeToString = "data"
Case RegionTypeXAxis
    gRegionTypeToString = "x-axis"
Case RegionTypeYAxis
    gRegionTypeToString = "y-axis"
Case RegionTypeBackground
    gRegionTypeToString = "background"
End Select
End Function

Public Function gSetProperty( _
                ByVal pExtHost As ExtendedPropertyHost, _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pNewValue As Variant, _
                Optional ByRef pPrevValue As Variant) As Boolean
Const ProcName As String = "gSetProperty"
On Error GoTo Err

If Not IsMissing(pPrevValue) Then gSetVariant pPrevValue, pExtHost.getLocalValue(pExtProp)

gSetProperty = pExtHost.SetValue(pExtProp, pNewValue)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Sub Main()
Debug.Print "ChartSkil running in development environment: " & CStr(inDev)
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function inDev() As Boolean
mIsInDev = True
inDev = True
End Function

Private Sub SetToolbarColor(ByVal pToolbar As Toolbar, ByVal pColor As Long)
Dim lBrush As Long
lBrush = CreateSolidBrush(NormalizeColor(pColor))

Dim lResult As Long
Select Case pToolbar.Style
Case ToolbarStyleConstants.tbrFlat
    lResult = SetClassLong(pToolbar.hWnd, GCLP_HBRBACKGROUND, lBrush)
Case ToolbarStyleConstants.tbrStandard
    Dim lhWnd As Long
    lhWnd = FindWindowEx(pToolbar.hWnd, 0, "msvb_lib_toolbar", vbNullString)
    lResult = SetClassLong(lhWnd, GCLP_HBRBACKGROUND, lBrush)
End Select

End Sub


