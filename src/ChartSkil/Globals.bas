Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ProjectName                As String = "ChartSkil26"

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

Public Enum LayerNumberRange
    MinLayer = 0
    MaxLayer = 255
End Enum

Public Enum RegionTypes
    RegionTypeData
    RegionTypeXAxis
    RegionTypeYAxis
    RegionTypeBackground
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
    lStyle.Layer = LayerLowestUser
    lStyle.OutlineThickness = 1
    lStyle.SolidUpBody = False
    lStyle.TailThickness = 1
    lStyle.Thickness = 2
    lStyle.UpColor = vbBlack
    lStyle.Width = 0.6
End If
Set gDefaultBarStyle = lStyle
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
    lStyle.Layer = LayerLowestUser + 1
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
    lStyle.Layer = LayerHighestUser
    lStyle.LineStyle = LineStyles.LineSolid
    lStyle.Thickness = 1
End If
Set gDefaultLineStyle = lStyle
End Property

Public Property Get gDefaultTextStyle() As TextStyle
Static lStyle As TextStyle
If lStyle Is Nothing Then
    Set lStyle = New TextStyle
    
    Dim aFont As New StdFont
    aFont.Bold = False
    aFont.Italic = False
    aFont.Name = "Arial"
    aFont.Size = 8
    aFont.Strikethrough = False
    aFont.Underline = False
    
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
    lStyle.Font = aFont
    lStyle.Ellipsis = EllipsisModes.EllipsisNone
    lStyle.ExpandTabs = True
    lStyle.Extended = False
    lStyle.FixedX = False
    lStyle.FixedY = False
    lStyle.HideIfBlank = True
    lStyle.IncludeInAutoscale = False
    lStyle.Justification = TextJustifyModes.JustifyLeft
    lStyle.Layer = LayerHighestUser
    lStyle.MultiLine = False
    lStyle.PaddingX = 1
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

Public Property Get gLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("chartskil.log")
Set gLogger = lLogger
End Property

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Property Get gTracer() As Tracer
Static lTracer As Tracer
If lTracer Is Nothing Then Set lTracer = GetTracer("chartskil")
Set gTracer = lTracer
End Property

'================================================================================
' Methods
'================================================================================

Public Function gCloneFont( _
                ByVal pFont As StdFont) As StdFont
Dim aFont As StdFont
Set aFont = New StdFont
aFont.Bold = pFont.Bold
aFont.Charset = pFont.Charset
aFont.Italic = pFont.Italic
aFont.Name = pFont.Name
aFont.Size = pFont.Size
aFont.Strikethrough = pFont.Strikethrough
aFont.Underline = pFont.Underline
aFont.Weight = pFont.Weight
Set gCloneFont = aFont
End Function

Public Function gDegreesToRadians( _
                ByVal degrees As Double) As Double
gDegreesToRadians = degrees * Pi / 180
End Function

Public Function gIsValidColor( _
                ByVal Value As Long) As Boolean
                
If Value > &HFFFFFF Then Exit Function
If Value < 0 And Value > MaxSystemColor Then Exit Function
gIsValidColor = True
End Function

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

