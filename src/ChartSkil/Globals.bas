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
    lStyle.TailThickness = 1
    lStyle.OutlineThickness = 1
    lStyle.UpColor = vbBlack
    lStyle.DownColor = vbBlack
    lStyle.DisplayMode = BarDisplayModeCandlestick
    lStyle.SolidUpBody = False
    lStyle.Thickness = 2
    lStyle.Width = 0.6
    lStyle.Color = -1
    lStyle.IncludeInAutoscale = True
End If
Set gDefaultBarStyle = lStyle
End Property

Public Property Get gDefaultDataPointStyle() As DataPointStyle
Static lStyle As DataPointStyle
If lStyle Is Nothing Then
    Set lStyle = New DataPointStyle
    lStyle.LineThickness = 1
    lStyle.Color = vbBlack
    lStyle.LineStyle = LineStyles.LineSolid
    lStyle.PointStyle = PointRound
    lStyle.DisplayMode = DataPointDisplayModes.DataPointDisplayModeLine
    lStyle.HistogramBarWidth = 0.6
    lStyle.DownColor = -1
    lStyle.UpColor = -1
    lStyle.IncludeInAutoscale = True
End If
Set gDefaultDataPointStyle = lStyle
End Property

Public Property Get gDefaultLineStyle() As LineStyle
Static lStyle As LineStyle
If lStyle Is Nothing Then
    Set lStyle = New LineStyle
    lStyle.Color = vbBlack
    lStyle.Thickness = 1
    lStyle.LineStyle = LineStyles.LineSolid
    lStyle.ExtendBefore = False
    lStyle.ExtendAfter = False
    lStyle.ArrowStartStyle = ArrowStyles.ArrowNone
    lStyle.ArrowStartLength = 10
    lStyle.ArrowStartWidth = 10
    lStyle.ArrowStartColor = vbBlack
    lStyle.ArrowStartFillColor = vbBlack
    lStyle.ArrowStartFillStyle = FillStyles.FillSolid
    lStyle.ArrowEndStyle = ArrowStyles.ArrowNone
    lStyle.ArrowEndLength = 10
    lStyle.ArrowEndWidth = 10
    lStyle.ArrowEndColor = vbBlack
    lStyle.ArrowEndFillColor = vbBlack
    lStyle.ArrowEndFillStyle = FillStyles.FillSolid
    lStyle.IncludeInAutoscale = False
    lStyle.FixedX = False
    lStyle.FixedY = False
    lStyle.Extended = False
End If
Set gDefaultLineStyle = lStyle
End Property

Public Property Get gDefaultTextStyle() As TextStyle
Static lStyle As TextStyle
If lStyle Is Nothing Then
    Set lStyle = New TextStyle
    lStyle.Font = New StdFont
    lStyle.Font.Bold = False
    lStyle.Font.Italic = False
    lStyle.Font.Name = "Arial"
    lStyle.Font.Size = 8
    lStyle.Font.Strikethrough = False
    lStyle.Font.Underline = False
    
    lStyle.Color = vbBlack
    lStyle.Box = False
    lStyle.BoxColor = vbBlack
    lStyle.BoxStyle = LineStyles.LineSolid
    lStyle.BoxThickness = 1
    lStyle.BoxFillColor = vbWhite
    lStyle.BoxFillStyle = FillStyles.FillSolid
    lStyle.BoxFillWithBackgroundColor = False
    lStyle.Align = TextAlignModes.AlignTopLeft
    lStyle.PaddingX = 1
    lStyle.PaddingY = 0#
    lStyle.IncludeInAutoscale = False
    lStyle.FixedX = False
    lStyle.FixedY = False
    lStyle.Extended = False
    
    lStyle.Angle = 0
    lStyle.Justification = TextJustifyModes.JustifyLeft
    lStyle.MultiLine = False
    lStyle.Ellipsis = EllipsisModes.EllipsisNone
    lStyle.ExpandTabs = True
    lStyle.TabWidth = 8
    lStyle.WordWrap = True
    lStyle.HideIfBlank = True
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

Public Property Get gLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = GetLogger("chartskil.log")
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

