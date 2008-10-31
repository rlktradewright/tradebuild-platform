Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const Pi As Double = 3.14159265358979

Public Const MinusInfinityDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023
Public Const PlusInfinityDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const MinusInfinitySingle As Single = -(2 - 2 ^ -23) * 2 ^ 127
Public Const PlusInfinitySingle As Single = (2 - 2 ^ -23) * 2 ^ 127

Public Const OneMicroSecond As Double = 1.15740740740741E-11

Public Const MaxSystemColor As Long = &H80000018

Public Const GridlineSpacingCm As Double = 1.8

Public Const HitTestTolerancePixels As Long = 3

Public Const TwipsPerCm As Double = 1440 / 2.54

Public Const ToolbarCommandAutoScale As String = "autoscale"
Public Const ToolbarCommandAutoScroll As String = "autoscroll"

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

Public Enum PointerModes
    PointerModeDefault
    PointerModeTool
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mLogger As Logger

Private mIsInDev As Boolean

Public gBlankMouseIcon As IPictureDisp

'================================================================================
' Properties
'================================================================================

Public Property Get gIsInDev() As Boolean
gIsInDev = mIsInDev
End Property

Public Property Get gLogger() As Logger
If mLogger Is Nothing Then Set mLogger = GetLogger("log.chartskil")
Set gLogger = mLogger
End Property

'================================================================================
' Methods
'================================================================================

Public Function gCalculateX( _
                ByVal timestamp As Date, _
                ByVal pController As ChartController, _
                Optional ByVal forceNewPeriod As Boolean, _
                Optional ByVal duplicateNumber As Long) As Double
Dim lPeriod As period
Dim periodEndtime As Date

Select Case pController.BarTimePeriod.units
Case TimePeriodNone, _
        TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear
    
    On Error Resume Next
    Set lPeriod = pController.Periods.Item(timestamp)
    On Error GoTo 0
    
    If lPeriod Is Nothing Then
        If pController.Periods.Count = 0 Then
            Set lPeriod = pController.Periods.addPeriod(timestamp)
        ElseIf timestamp < pController.Periods.Item(1).timestamp Then
            Set lPeriod = pController.Periods.Item(1)
            timestamp = lPeriod.timestamp
        Else
            Set lPeriod = pController.Periods.addPeriod(timestamp)
        End If
    End If
    
    periodEndtime = BarEndTime(lPeriod.timestamp, _
                            pController.BarTimePeriod, _
                            pController.SessionStartTime)
    gCalculateX = lPeriod.PeriodNumber + (timestamp - lPeriod.timestamp) / (periodEndtime - lPeriod.timestamp)
    
Case TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    If Not forceNewPeriod Then
        On Error Resume Next
        Set lPeriod = pController.Periods.itemDup(timestamp, duplicateNumber)
        On Error GoTo 0
        
        If lPeriod Is Nothing Then
            Set lPeriod = pController.Periods.addPeriod(timestamp, True)
        End If
        gCalculateX = lPeriod.PeriodNumber
    Else
        Set lPeriod = pController.Periods.addPeriod(timestamp, True)
        gCalculateX = lPeriod.PeriodNumber
    End If
End Select

End Function

Public Function gCloneFont( _
                ByVal aFont As StdFont) As StdFont
Set gCloneFont = New StdFont
With gCloneFont
    .Bold = aFont.Bold
    .Charset = aFont.Charset
    .Italic = aFont.Italic
    .name = aFont.name
    .size = aFont.size
    .Strikethrough = aFont.Strikethrough
    .Underline = aFont.Underline
    .Weight = aFont.Weight
End With
End Function

Public Function gCreateBarStyle() As BarStyle
Set gCreateBarStyle = New BarStyle
With gCreateBarStyle
    .includeInAutoscale = True
    .tailThickness = 1
    .outlineThickness = 1
    .upColor = &H1D9311
    .downColor = &H43FC2
    .displayMode = BarDisplayModeBar
    .solidUpBody = True
    .barThickness = 2
    .barWidth = 0.6
    .barColor = -1
End With
End Function

Public Function gCreateDataPointStyle() As DataPointStyle
Set gCreateDataPointStyle = New DataPointStyle
With gCreateDataPointStyle
    .lineThickness = 1
    .Color = vbBlack
    .linestyle = LineStyles.LineSolid
    .pointStyle = PointRound
    .displayMode = DataPointDisplayModes.DataPointDisplayModeLine
    .histBarWidth = 0.6
    .includeInAutoscale = True
    .downColor = -1
    .upColor = -1
End With
End Function

Public Function gCreateChartRegionStyle() As ChartRegionStyle
Set gCreateChartRegionStyle = New ChartRegionStyle
With gCreateChartRegionStyle
    .Autoscale = True
    .BackColor = vbWhite
    .GridColor = &HC0C0C0
    .GridlineSpacingY = GridlineSpacingCm
    .GridTextColor = vbBlack
    .HasGrid = True
    .IntegerYScale = False
    .HasGridText = False
    '.pointerStyle = PointerStyles.PointerCrosshairs
    .MinimumHeight = 0
    .YScaleQuantum = 0
End With
End Function

Public Function gCreateLineStyle() As linestyle
Set gCreateLineStyle = New linestyle
With gCreateLineStyle
    .Color = vbBlack
    .thickness = 1
    .linestyle = LineStyles.LineSolid
    .extendBefore = False
    .extendAfter = False
    .arrowStartStyle = ArrowStyles.ArrowNone
    .arrowStartLength = 10
    .arrowStartWidth = 10
    .arrowStartColor = vbBlack
    .arrowStartFillColor = vbBlack
    .arrowStartfillstyle = FillStyles.FillSolid
    .arrowEndStyle = ArrowStyles.ArrowNone
    .arrowEndLength = 10
    .arrowEndWidth = 10
    .arrowEndColor = vbBlack
    .arrowEndFillColor = vbBlack
    .arrowEndFillStyle = FillStyles.FillSolid
    .fixedX = False
    .fixedY = False
    .includeInAutoscale = False
    .extended = False
End With
End Function

Public Function gCreateTextStyle() As TextStyle
Dim aFont As StdFont

Set aFont = New StdFont
aFont.Bold = False
aFont.Italic = False
aFont.name = "Arial"
aFont.size = 8
aFont.Strikethrough = False
aFont.Underline = False

Set gCreateTextStyle = New TextStyle
With gCreateTextStyle
    .font = aFont
    .Color = vbBlack
    .box = False
    .boxColor = vbBlack
    .boxStyle = LineStyles.LineSolid
    .boxThickness = 1
    .boxFillColor = vbWhite
    .boxFillStyle = FillStyles.FillSolid
    .align = TextAlignModes.AlignBottomRight
    .includeInAutoscale = False
    .extended = False
    .paddingX = 1#
    .paddingY = 0.5
End With
End Function

Public Function gIsValidColor( _
                ByVal value As Long) As Boolean
                
If value > &HFFFFFF Then Exit Function
If value < 0 And value > MaxSystemColor Then Exit Function
gIsValidColor = True
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

