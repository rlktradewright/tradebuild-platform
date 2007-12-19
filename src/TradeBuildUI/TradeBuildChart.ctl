VERSION 5.00
Object = "{74951842-2BEF-4829-A34F-DC7795A37167}#16.0#0"; "ChartSkil2-6.ocx"
Begin VB.UserControl TradeBuildChart 
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   ScaleHeight     =   4965
   ScaleWidth      =   7275
   ToolboxBitmap   =   "TradeBuildChart.ctx":0000
   Begin ChartSkil26.Chart Chart1 
      Align           =   1  'Align Top
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   8070
   End
End
Attribute VB_Name = "TradeBuildChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements TaskCompletionListener

'@================================================================================
' Events
'@================================================================================

Event StateChange(ev As StateChangeEvent)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                               As String = "TradeBuildUI25"
Private Const ModuleName                                As String = "TradeBuildChart"

Private Const PropNameAllowHorizontalMouseScrolling     As String = "AllowHorizontalMouseScrolling"
Private Const PropNameAllowVerticalMouseScrolling       As String = "AllowVerticalMouseScrolling"
Private Const PropNameAutoscroll                        As String = "Autoscroll"
Private Const PropNameChartBackColor                    As String = "ChartBackColor"
Private Const PropNameDfltRegnStyleAutoscale            As String = "DfltRegnStyleAutoscale"
Private Const PropNameDfltRegnStyleBackColor            As String = "DfltRegnStyleBackColor"
Private Const PropNameDfltRegnStyleGridColor            As String = "DfltRegnStyleGridColor"
Private Const PropNameDfltRegnStyleGridlineSpacingY     As String = "DfltRegnStyleGridlineSpacingY"
Private Const PropNameDfltRegnStyleGridTextColor        As String = "DfltRegnStyleGridTextColor"
Private Const PropNameDfltRegnStyleHasGrid              As String = "DfltRegnStyleHasGrid"
Private Const PropNameDfltRegnStyleHasGridtext          As String = "DfltRegnStyleHasGridtext"
Private Const PropNameDfltRegnStylePointerStyle         As String = "DfltRegnStylePointerStyle"
Private Const PropNamePointerDiscColor                  As String = "PointerDiscColor"
Private Const PropNamePointerCrosshairsColor            As String = "PointerCrosshairsColor"
Private Const PropNameShowHorizontalScrollBar           As String = "ShowHorizontalScrollBar"
Private Const PropNameShowToolbar                       As String = "ShowToobar"
Private Const PropNameTwipsPerBar                       As String = "TwipsPerBar"
Private Const PropNameYAxisWidthCm                      As String = "YAxisWidthCm"

Private Const PropDfltAllowHorizontalMouseScrolling     As Boolean = True
Private Const PropDfltAllowVerticalMouseScrolling       As Boolean = True
Private Const PropDfltAutoscroll                        As Boolean = True
Private Const PropDfltChartBackColor                    As Long = vbWhite
Private Const PropDfltDfltRegnStyleAutoscale            As Boolean = True
Private Const PropDfltDfltRegnStyleBackColor            As Long = vbWhite
Private Const PropDfltDfltRegnStyleGridColor            As Long = &HC0C0C0
Private Const PropDfltDfltRegnStyleGridlineSpacingY     As Double = 1.8
Private Const PropDfltDfltRegnStyleGridTextColor        As Long = vbBlack
Private Const PropDfltDfltRegnStyleHasGrid              As Boolean = True
Private Const PropDfltDfltRegnStyleHasGridtext          As Boolean = False
Private Const PropDfltDfltRegnStylePointerStyle         As Long = PointerStyles.PointerCrosshairs
Private Const PropDfltPointerDiscColor                  As Long = &H89FFFF
Private Const PropDfltPointerCrosshairsColor            As Long = &HC1DFE
Private Const PropDfltShowHorizontalScrollBar           As Boolean = True
Private Const PropDfltShowToolbar                       As Boolean = True
Private Const PropDfltTwipsPerBar                       As Long = 150
Private Const PropDfltYAxisWidthCm                      As Single = 1.3

'@================================================================================
' Member variables
'@================================================================================

Private mChartManager As chartManager
Private mChartController As chartController

Private WithEvents mTicker As ticker
Attribute mTicker.VB_VarHelpID = -1
Private mTimeframes As Timeframes
Private WithEvents mTimeframe As timeframe
Attribute mTimeframe.VB_VarHelpID = -1

Private mBarsStudyConfig As StudyConfiguration

Private mUpdatePerTick As Boolean

Private mIsHistoricChart As Boolean
Private mInitialNumberOfBars As Long
Private mFromTime As Date
Private mToTime As Date
Private mIncludeBarsOutsideSession As Boolean
Private mMinimumTicksHeight As Long

Private mContract As Contract

Private mPeriodLength As Long
Private mPeriodUnits As TimePeriodUnits

Private mPriceRegion As ChartRegion
Private mPriceRegionStyle As ChartRegionStyle

Private mVolumeRegion As ChartRegion
Private mVolumeRegionStyle As ChartRegionStyle

Private mBarsStyle As barStyle
Private mVolumeStyle As dataPointStyle

Private mPrevWidth As Single
Private mPrevHeight As Single

Private mNumberOfOutstandingTasks As Long
Private mHistDataLoaded As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()

mPrevWidth = UserControl.Width
mPrevHeight = UserControl.Height

mUpdatePerTick = True

End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next

AllowHorizontalMouseScrolling = PropDfltAllowHorizontalMouseScrolling
AllowVerticalMouseScrolling = PropDfltAllowVerticalMouseScrolling
Autoscroll = PropDfltAutoscroll
chartBackColor = PropDfltChartBackColor
RegionDefaultAutoscale = PropDfltDfltRegnStyleAutoscale
RegionDefaultBackColor = PropDfltDfltRegnStyleBackColor
RegionDefaultGridColor = PropDfltDfltRegnStyleGridColor
RegionDefaultGridlineSpacingY = PropDfltDfltRegnStyleGridlineSpacingY
RegionDefaultGridTextColor = PropDfltDfltRegnStyleGridTextColor
RegionDefaultHasGrid = PropDfltDfltRegnStyleHasGrid
RegionDefaultHasGridText = PropDfltDfltRegnStyleHasGridtext
RegionDefaultPointerStyle = PropDfltDfltRegnStylePointerStyle
PointerCrosshairsColor = PropDfltPointerCrosshairsColor
PointerDiscColor = PropDfltPointerDiscColor
showHorizontalScrollBar = PropDfltShowHorizontalScrollBar
showToolbar = PropDfltShowToolbar
twipsPerBar = PropDfltTwipsPerBar
YAxisWidthCm = PropDfltYAxisWidthCm

initialiseChart
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

AllowHorizontalMouseScrolling = PropBag.ReadProperty(PropNameAllowHorizontalMouseScrolling, PropDfltAllowHorizontalMouseScrolling)
If Err.Number <> 0 Then
    AllowHorizontalMouseScrolling = PropDfltAllowHorizontalMouseScrolling
    Err.clear
End If

AllowVerticalMouseScrolling = PropBag.ReadProperty(PropNameAllowVerticalMouseScrolling, PropDfltAllowVerticalMouseScrolling)
If Err.Number <> 0 Then
    AllowVerticalMouseScrolling = PropDfltAllowVerticalMouseScrolling
    Err.clear
End If

Autoscroll = PropBag.ReadProperty(PropNameAutoscroll, PropDfltAutoscroll)
If Err.Number <> 0 Then
    Autoscroll = PropDfltAutoscroll
    Err.clear
End If

chartBackColor = PropBag.ReadProperty(PropNameChartBackColor, PropDfltChartBackColor)
If Err.Number <> 0 Then
    chartBackColor = PropDfltChartBackColor
    Err.clear
End If

RegionDefaultAutoscale = PropBag.ReadProperty(PropNameDfltRegnStyleAutoscale, PropDfltDfltRegnStyleAutoscale)
If Err.Number <> 0 Then
    RegionDefaultAutoscale = PropDfltDfltRegnStyleAutoscale
    Err.clear
End If

RegionDefaultBackColor = PropBag.ReadProperty(PropNameDfltRegnStyleBackColor, PropDfltDfltRegnStyleBackColor)
If Err.Number <> 0 Then
    RegionDefaultBackColor = PropDfltDfltRegnStyleBackColor
    Err.clear
End If

RegionDefaultGridColor = PropBag.ReadProperty(PropNameDfltRegnStyleGridColor, PropDfltDfltRegnStyleGridColor)
If Err.Number <> 0 Then
    RegionDefaultGridColor = PropDfltDfltRegnStyleGridColor
    Err.clear
End If

RegionDefaultGridlineSpacingY = PropBag.ReadProperty(PropNameDfltRegnStyleGridlineSpacingY, PropDfltDfltRegnStyleGridlineSpacingY)
If Err.Number <> 0 Then
    RegionDefaultGridlineSpacingY = PropDfltDfltRegnStyleGridlineSpacingY
    Err.clear
End If

RegionDefaultGridTextColor = PropBag.ReadProperty(PropNameDfltRegnStyleGridTextColor, PropDfltDfltRegnStyleGridTextColor)
If Err.Number <> 0 Then
    RegionDefaultGridTextColor = PropDfltDfltRegnStyleGridTextColor
    Err.clear
End If

RegionDefaultHasGrid = PropBag.ReadProperty(PropNameDfltRegnStyleHasGrid, PropDfltDfltRegnStyleHasGrid)
If Err.Number <> 0 Then
    RegionDefaultHasGrid = PropDfltDfltRegnStyleHasGrid
    Err.clear
End If

RegionDefaultHasGridText = PropBag.ReadProperty(PropNameDfltRegnStyleHasGridtext, PropDfltDfltRegnStyleHasGridtext)
If Err.Number <> 0 Then
    RegionDefaultHasGridText = PropDfltDfltRegnStyleHasGridtext
    Err.clear
End If

RegionDefaultPointerStyle = PropBag.ReadProperty(PropNameDfltRegnStylePointerStyle, PropDfltDfltRegnStylePointerStyle)
If Err.Number <> 0 Then
    RegionDefaultPointerStyle = PropDfltDfltRegnStylePointerStyle
    Err.clear
End If

PointerCrosshairsColor = PropBag.ReadProperty(PropNamePointerCrosshairsColor, PropDfltPointerCrosshairsColor)
If Err.Number <> 0 Then
    PointerCrosshairsColor = PropDfltPointerCrosshairsColor
    Err.clear
End If

PointerDiscColor = PropBag.ReadProperty(PropNamePointerDiscColor, PropDfltPointerDiscColor)
If Err.Number <> 0 Then
    PointerDiscColor = PropDfltPointerDiscColor
    Err.clear
End If

showHorizontalScrollBar = PropBag.ReadProperty(PropNameShowHorizontalScrollBar, PropDfltShowHorizontalScrollBar)
If Err.Number <> 0 Then
    showHorizontalScrollBar = PropDfltShowHorizontalScrollBar
    Err.clear
End If

showToolbar = PropBag.ReadProperty(PropNameShowToolbar, PropDfltShowToolbar)
If Err.Number <> 0 Then
    showToolbar = PropDfltShowToolbar
    Err.clear
End If

twipsPerBar = PropBag.ReadProperty(PropNameTwipsPerBar, PropDfltTwipsPerBar)
If Err.Number <> 0 Then
    twipsPerBar = PropDfltTwipsPerBar
    Err.clear
End If

YAxisWidthCm = PropBag.ReadProperty(PropNameYAxisWidthCm, PropDfltYAxisWidthCm)
If Err.Number <> 0 Then
    YAxisWidthCm = PropDfltYAxisWidthCm
    Err.clear
End If

initialiseChart

End Sub

Private Sub UserControl_Resize()
If UserControl.Width <> mPrevWidth Then
    mPrevWidth = UserControl.Width
End If
If UserControl.Height <> mPrevHeight Then
    Chart1.Height = UserControl.Height
    mPrevHeight = UserControl.Height
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty PropNameAllowHorizontalMouseScrolling, AllowHorizontalMouseScrolling, PropDfltAllowHorizontalMouseScrolling
PropBag.WriteProperty PropNameAllowVerticalMouseScrolling, AllowVerticalMouseScrolling, PropDfltAllowVerticalMouseScrolling
PropBag.WriteProperty PropNameAutoscroll, Autoscroll, PropDfltAutoscroll
PropBag.WriteProperty PropNameChartBackColor, chartBackColor, PropDfltChartBackColor
PropBag.WriteProperty PropNameDfltRegnStyleAutoscale, RegionDefaultAutoscale, PropDfltDfltRegnStyleAutoscale
PropBag.WriteProperty PropNameDfltRegnStyleBackColor, RegionDefaultBackColor, PropDfltDfltRegnStyleBackColor
PropBag.WriteProperty PropNameDfltRegnStyleGridColor, RegionDefaultGridColor, PropDfltDfltRegnStyleGridColor
PropBag.WriteProperty PropNameDfltRegnStyleGridlineSpacingY, RegionDefaultGridlineSpacingY, PropDfltDfltRegnStyleGridlineSpacingY
PropBag.WriteProperty PropNameDfltRegnStyleGridTextColor, RegionDefaultGridTextColor, PropDfltDfltRegnStyleGridTextColor
PropBag.WriteProperty PropNameDfltRegnStyleHasGrid, RegionDefaultHasGrid, PropDfltDfltRegnStyleHasGrid
PropBag.WriteProperty PropNameDfltRegnStyleHasGridtext, RegionDefaultHasGridText, PropDfltDfltRegnStyleHasGridtext
PropBag.WriteProperty PropNameDfltRegnStylePointerStyle, RegionDefaultPointerStyle, PropDfltDfltRegnStylePointerStyle
PropBag.WriteProperty PropNamePointerCrosshairsColor, PointerCrosshairsColor, PropDfltPointerCrosshairsColor
PropBag.WriteProperty PropNamePointerDiscColor, PointerDiscColor, PropDfltPointerDiscColor
PropBag.WriteProperty PropNameShowHorizontalScrollBar, showHorizontalScrollBar, PropDfltShowHorizontalScrollBar
PropBag.WriteProperty PropNameShowToolbar, showToolbar, PropDfltShowToolbar
PropBag.WriteProperty PropNameTwipsPerBar, twipsPerBar, PropDfltTwipsPerBar
PropBag.WriteProperty PropNameYAxisWidthCm, YAxisWidthCm, PropDfltYAxisWidthCm
End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_taskCompleted(ev As TaskCompletionEvent)
Dim stateEv As StateChangeEvent

mNumberOfOutstandingTasks = mNumberOfOutstandingTasks - 1
If mNumberOfOutstandingTasks = 0 And mHistDataLoaded Then
    stateEv.state = ChartStates.ChartStateLoaded
    Set stateEv.Source = Me
    RaiseEvent StateChange(stateEv)
End If
End Sub

'@================================================================================
' mTicker Event Handlers
'@================================================================================

Private Sub mTicker_StateChange(ev As StateChangeEvent)
Dim stateEv As StateChangeEvent
If ev.state = TickerStates.TickerStateReady Then
    ' this means that the ticker object has retrieved the contract info, so we can
    ' now start the chart
    loadchart
    stateEv.state = ChartStates.ChartStateInitialised
    Set stateEv.Source = Me
    RaiseEvent StateChange(stateEv)
End If
End Sub

'@================================================================================
' mTimeframe Event Handlers
'@================================================================================

Private Sub mTimeframe_BarsLoaded()
Dim stateEv As StateChangeEvent

Chart1.suppressDrawing = False

mHistDataLoaded = True

If mNumberOfOutstandingTasks = 0 Then
    stateEv.state = ChartStates.ChartStateLoaded
    Set stateEv.Source = Me
    RaiseEvent StateChange(stateEv)
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let AllowHorizontalMouseScrolling( _
                ByVal value As Boolean)
Chart1.AllowHorizontalMouseScrolling = value
End Property

Public Property Get AllowHorizontalMouseScrolling() As Boolean
Attribute AllowHorizontalMouseScrolling.VB_ProcData.VB_Invoke_Property = ";Behavior"
AllowHorizontalMouseScrolling = Chart1.AllowHorizontalMouseScrolling
End Property

Public Property Let AllowVerticalMouseScrolling( _
                ByVal value As Boolean)
Chart1.AllowVerticalMouseScrolling = value
End Property

Public Property Get AllowVerticalMouseScrolling() As Boolean
Attribute AllowVerticalMouseScrolling.VB_ProcData.VB_Invoke_Property = ";Behavior"
AllowVerticalMouseScrolling = Chart1.AllowVerticalMouseScrolling
End Property

Public Property Let Autoscroll( _
                ByVal value As Boolean)
Chart1.Autoscroll = value
End Property

Public Property Get Autoscroll() As Boolean
Attribute Autoscroll.VB_ProcData.VB_Invoke_Property = ";Behavior"
Autoscroll = Chart1.Autoscroll
End Property

Public Property Let barsStyle(ByVal value As barStyle)
Set mBarsStyle = value
End Property

Public Property Get chartBackColor() As OLE_COLOR
Attribute chartBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
chartBackColor = Chart1.chartBackColor
End Property

Public Property Let chartBackColor(ByVal val As OLE_COLOR)
Chart1.chartBackColor = val
End Property

Public Property Get chartController() As chartController
Set chartController = Chart1.controller
End Property

Public Property Get chartManager() As chartManager
Set chartManager = mChartManager
End Property

Public Property Get initialNumberOfBars() As Long
Attribute initialNumberOfBars.VB_ProcData.VB_Invoke_Property = ";Behavior"
initialNumberOfBars = mInitialNumberOfBars
End Property

Public Property Get minimumTicksHeight() As Double
Attribute minimumTicksHeight.VB_ProcData.VB_Invoke_Property = ";Behavior"
minimumTicksHeight = mMinimumTicksHeight
End Property

Public Property Get PointerCrosshairsColor() As OLE_COLOR
Attribute PointerCrosshairsColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
PointerCrosshairsColor = Chart1.PointerCrosshairsColor
End Property

Public Property Let PointerCrosshairsColor(ByVal value As OLE_COLOR)
Chart1.PointerCrosshairsColor = value
End Property

Public Property Get PointerDiscColor() As OLE_COLOR
Attribute PointerDiscColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
PointerDiscColor = Chart1.PointerDiscColor
End Property

Public Property Let PointerDiscColor(ByVal value As OLE_COLOR)
Chart1.PointerDiscColor = value
End Property

Public Property Get priceRegion() As ChartRegion
Set priceRegion = mPriceRegion
End Property

Public Property Let priceRegionStyle(ByVal value As ChartRegionStyle)
Set mPriceRegionStyle = value
If Not mPriceRegion Is Nothing Then mPriceRegion.Style = value
End Property

Public Property Get priceRegionStyle() As ChartRegionStyle
Set priceRegionStyle = mPriceRegionStyle
End Property

Public Property Get RegionDefaultAutoscale() As Boolean
Attribute RegionDefaultAutoscale.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultAutoscale = Chart1.RegionDefaultAutoscale
End Property

Public Property Let RegionDefaultAutoscale(ByVal value As Boolean)
Chart1.RegionDefaultAutoscale = value
End Property

Public Property Get RegionDefaultBackColor() As OLE_COLOR
Attribute RegionDefaultBackColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultBackColor = Chart1.RegionDefaultBackColor
End Property

Public Property Let RegionDefaultBackColor(ByVal val As OLE_COLOR)
Chart1.RegionDefaultBackColor = val
End Property

Public Property Get RegionDefaultGridColor() As OLE_COLOR
Attribute RegionDefaultGridColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultGridColor = Chart1.RegionDefaultGridColor
End Property

Public Property Let RegionDefaultGridColor(ByVal val As OLE_COLOR)
Chart1.RegionDefaultGridColor = val
End Property

Public Property Get RegionDefaultGridlineSpacingY() As Double
Attribute RegionDefaultGridlineSpacingY.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultGridlineSpacingY = Chart1.RegionDefaultGridlineSpacingY
End Property

Public Property Let RegionDefaultGridlineSpacingY(ByVal value As Double)
Chart1.RegionDefaultGridlineSpacingY = value
End Property

Public Property Get RegionDefaultGridTextColor() As OLE_COLOR
Attribute RegionDefaultGridTextColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultGridTextColor = Chart1.RegionDefaultGridTextColor
End Property

Public Property Let RegionDefaultGridTextColor(ByVal val As OLE_COLOR)
Chart1.RegionDefaultGridTextColor = val
End Property

Public Property Get RegionDefaultHasGrid() As Boolean
Attribute RegionDefaultHasGrid.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultHasGrid = Chart1.RegionDefaultHasGrid
End Property

Public Property Let RegionDefaultHasGrid(ByVal val As Boolean)
Chart1.RegionDefaultHasGrid = val
End Property

Public Property Get RegionDefaultHasGridText() As Boolean
Attribute RegionDefaultHasGridText.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultHasGridText = Chart1.RegionDefaultHasGridText
End Property

Public Property Let RegionDefaultHasGridText(ByVal val As Boolean)
Chart1.RegionDefaultHasGridText = val
End Property

Public Property Get RegionDefaultPointerStyle() As PointerStyles
Attribute RegionDefaultPointerStyle.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultPointerStyle = Chart1.RegionDefaultPointerStyle
End Property

Public Property Let RegionDefaultPointerStyle(ByVal value As PointerStyles)
Chart1.RegionDefaultPointerStyle = value
End Property

Public Property Get regionNames() As String()
regionNames = mChartManager.regionNames
End Property

Public Property Get showHorizontalScrollBar() As Boolean
Attribute showHorizontalScrollBar.VB_ProcData.VB_Invoke_Property = ";Appearance"
showHorizontalScrollBar = Chart1.showHorizontalScrollBar
End Property

Public Property Let showHorizontalScrollBar(ByVal val As Boolean)
Chart1.showHorizontalScrollBar = val
End Property

Public Property Get showToolbar() As Boolean
Attribute showToolbar.VB_ProcData.VB_Invoke_Property = ";Appearance"
showToolbar = Chart1.showToolbar
End Property

Public Property Let showToolbar(ByVal val As Boolean)
Chart1.showToolbar = val
End Property

Public Property Get timeframeCaption() As String

Select Case mPeriodUnits
Case TimePeriodUnits.TimePeriodSecond
    timeframeCaption = IIf(mPeriodLength = 1, "1 Sec", mPeriodLength & " Secs")
Case TimePeriodUnits.TimePeriodMinute
    timeframeCaption = IIf(mPeriodLength = 1, "1 Min", mPeriodLength & " Mins")
Case TimePeriodUnits.TimePeriodHour
    timeframeCaption = IIf(mPeriodLength = 1, "1 Hour", mPeriodLength & " Hrs")
Case TimePeriodUnits.TimePeriodDay
    timeframeCaption = IIf(mPeriodLength = 1, "Daily", mPeriodLength & " Days")
Case TimePeriodUnits.TimePeriodWeek
    timeframeCaption = IIf(mPeriodLength = 1, "Weekly", mPeriodLength & " Wks")
Case TimePeriodUnits.TimePeriodMonth
    timeframeCaption = IIf(mPeriodLength = 1, "Monthly", mPeriodLength & " Mths")
Case TimePeriodUnits.TimePeriodYear
    timeframeCaption = IIf(mPeriodLength = 1, "Yearly", mPeriodLength & " Yrs")
End Select
End Property

Public Property Get timeframe() As timeframe
Set timeframe = mTimeframe
End Property

Public Property Get twipsPerBar() As Long
Attribute twipsPerBar.VB_ProcData.VB_Invoke_Property = ";Appearance"
twipsPerBar = Chart1.twipsPerBar
End Property

Public Property Let twipsPerBar(ByVal val As Long)
Chart1.twipsPerBar = val
End Property

Public Property Let updatePerTick(ByVal value As Boolean)
Attribute updatePerTick.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
mUpdatePerTick = value
End Property

Public Property Get volumeRegion() As ChartRegion
Set volumeRegion = mVolumeRegion
End Property

Public Property Let volumeRegionStyle(ByVal value As ChartRegionStyle)
Set mVolumeRegionStyle = value
If Not mVolumeRegion Is Nothing Then mVolumeRegion.Style = value
End Property

Public Property Get volumeRegionStyle() As ChartRegionStyle
Set volumeRegionStyle = mVolumeRegionStyle
End Property

Public Property Let volumeStyle(ByVal value As dataPointStyle)
Set mVolumeStyle = value
End Property

Public Property Get YAxisWidthCm() As Single
Attribute YAxisWidthCm.VB_ProcData.VB_Invoke_Property = ";Appearance"
YAxisWidthCm = Chart1.YAxisWidthCm
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
Chart1.YAxisWidthCm = value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub clearChart()
initialiseChart
End Sub

Public Sub finish()
mChartManager.finish

Set mTimeframes = Nothing
Set mTimeframe = Nothing
Set mTicker = Nothing
End Sub

Public Sub scrollToTime(ByVal pTime As Date)
mChartManager.scrollToTime pTime
End Sub

Public Sub showChart( _
                ByVal pTicker As ticker, _
                ByVal initialNumberOfBars As Long, _
                ByVal includeBarsOutsideSession As Boolean, _
                ByVal minimumTicksHeight As Long, _
                ByVal periodlength As Long, _
                ByVal periodUnits As TimePeriodUnits, _
                Optional ByVal priceRegionStyle As ChartRegionStyle, _
                Optional ByVal volumeRegionStyle As ChartRegionStyle)

Select Case periodUnits
Case TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear, _
        TimePeriodVolume, _
        TimePeriodTickMovement
Case Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "showChart", _
                "Time period units not supported"
    
End Select

Set mTicker = pTicker
mInitialNumberOfBars = initialNumberOfBars
mIncludeBarsOutsideSession = includeBarsOutsideSession
mMinimumTicksHeight = minimumTicksHeight
mPeriodLength = periodlength
mPeriodUnits = periodUnits
Set mPriceRegionStyle = priceRegionStyle
Set mVolumeRegionStyle = volumeRegionStyle

Set mChartManager = createChartManager(mTicker.StudyManager, Chart1.controller)

If mTicker.state = TickerStateRunning Then
    loadchart
End If
End Sub

Public Sub showHistoricChart( _
                ByVal pTicker As ticker, _
                ByVal initialNumberOfBars As Long, _
                ByVal fromTime As Date, _
                ByVal toTime As Date, _
                ByVal includeBarsOutsideSession As Boolean, _
                ByVal minimumTicksHeight As Long, _
                ByVal periodlength As Long, _
                ByVal periodUnits As TimePeriodUnits, _
                Optional ByVal priceRegionStyle As ChartRegionStyle, _
                Optional ByVal volumeRegionStyle As ChartRegionStyle)

Select Case periodUnits
Case TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear, _
        TimePeriodVolume, _
        TimePeriodTickMovement
Case Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "showHistoricChart", _
                "Time period units not supported"
    
End Select

mIsHistoricChart = True
Set mTicker = pTicker
mInitialNumberOfBars = initialNumberOfBars
mFromTime = fromTime
mToTime = toTime
mIncludeBarsOutsideSession = includeBarsOutsideSession
mMinimumTicksHeight = minimumTicksHeight
mPeriodLength = periodlength
mPeriodUnits = periodUnits
Set mPriceRegionStyle = priceRegionStyle
Set mVolumeRegionStyle = volumeRegionStyle

Set mChartManager = createChartManager(mTicker.StudyManager, Chart1.controller)

If mTicker.state = TickerStateReady Then
    loadchart
End If
End Sub

Public Sub showStudyPickerForm()
If mTicker.state = TickerStateRunning Then showStudyPicker mChartManager
End Sub

Public Sub syncStudyPickerForm()
If mTicker.state = TickerStateRunning Then syncStudyPicker mChartManager
End Sub

Public Sub unsyncStudyPickerForm()
unsyncStudyPicker
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createBarsStudyConfig() As StudyConfiguration
Dim lStudy As study
Dim studyDef As StudyDefinition
ReDim inputValueNames(1) As String
Dim params As New Parameters
Dim studyValueConfig As StudyValueConfiguration
Dim barsStyle As barStyle
Dim volumeStyle As dataPointStyle

Set createBarsStudyConfig = New StudyConfiguration

createBarsStudyConfig.underlyingStudy = mTicker.InputStudy

Set lStudy = mTimeframe.tradeStudy
createBarsStudyConfig.study = mTimeframe.tradeStudy
Set studyDef = lStudy.StudyDefinition

createBarsStudyConfig.chartRegionName = RegionNamePrice
inputValueNames(0) = mTicker.InputNameTrade
inputValueNames(1) = mTicker.InputNameVolume
createBarsStudyConfig.inputValueNames = inputValueNames
createBarsStudyConfig.name = studyDef.name
params.setParameterValue "Bar length", mPeriodLength
params.setParameterValue "Time units", TimePeriodUnitsToString(mPeriodUnits)
createBarsStudyConfig.Parameters = params
'createBarsStudyConfig.studyDefinition = studyDef

Set studyValueConfig = createBarsStudyConfig.StudyValueConfigurations.add("Bar")
studyValueConfig.chartRegionName = RegionNamePrice
studyValueConfig.includeInChart = True
studyValueConfig.layer = 200
If Not mBarsStyle Is Nothing Then
    Set barsStyle = mBarsStyle
Else
    Set barsStyle = Chart1.defaultBarStyle
    barsStyle.outlineThickness = 1
    barsStyle.barThickness = 2
    barsStyle.barWidth = 0.6
    barsStyle.displayMode = BarDisplayModeCandlestick
    barsStyle.downColor = &H43FC2
    barsStyle.solidUpBody = True
    barsStyle.tailThickness = 2
    barsStyle.upColor = &H1D9311
End If
studyValueConfig.barStyle = barsStyle

Set studyValueConfig = createBarsStudyConfig.StudyValueConfigurations.add("Volume")
studyValueConfig.chartRegionName = RegionNameVolume
studyValueConfig.includeInChart = True
If Not mVolumeStyle Is Nothing Then
    Set volumeStyle = mVolumeStyle
Else
    Set volumeStyle = Chart1.defaultDataPointStyle
    volumeStyle.upColor = vbGreen
    volumeStyle.downColor = vbRed
    volumeStyle.displayMode = DataPointDisplayModeHistogram
    volumeStyle.histBarWidth = 0.5
    volumeStyle.includeInAutoscale = True
    volumeStyle.lineThickness = 1
End If
studyValueConfig.dataPointStyle = volumeStyle
End Function

Private Sub initialiseChart()
Dim regionStyle As ChartRegionStyle

Set mChartController = Chart1.controller

Chart1.suppressDrawing = True

Chart1.clearChart

If Not mPriceRegionStyle Is Nothing Then
    Set regionStyle = mPriceRegionStyle
Else
    Set regionStyle = Chart1.defaultRegionStyle
    regionStyle.gridlineSpacingY = 2
End If

Set mPriceRegion = mChartController.addChartRegion(100, 25, regionStyle, RegionNamePrice)

If Not mVolumeRegionStyle Is Nothing Then
    Set regionStyle = mVolumeRegionStyle
Else
    Set regionStyle = Chart1.defaultRegionStyle
    regionStyle.gridlineSpacingY = 0.8
    regionStyle.minimumHeight = 10
    regionStyle.integerYScale = True
End If

Set mVolumeRegion = mChartController.addChartRegion(20, , regionStyle, RegionNameVolume)

Chart1.suppressDrawing = False

End Sub

Private Sub loadchart()

Set mContract = mTicker.Contract

Chart1.suppressDrawing = True

Chart1.setPeriodParameters mPeriodLength, mPeriodUnits

Chart1.sessionStartTime = mContract.sessionStartTime
Chart1.sessionEndTime = mContract.sessionEndTime

mPriceRegion.YScaleQuantum = mContract.ticksize
If mMinimumTicksHeight * mContract.ticksize <> 0 Then
    mPriceRegion.minimumHeight = mMinimumTicksHeight * mContract.ticksize
End If

mPriceRegion.setTitle mContract.specifier.localSymbol & _
                " (" & mContract.specifier.exchange & ") " & _
                timeframeCaption, _
                vbBlue, _
                Nothing

mVolumeRegion.setTitle "Volume", vbBlue, Nothing

Chart1.suppressDrawing = False

Set mTimeframes = mTicker.Timeframes

If mIsHistoricChart Then
    Set mTimeframe = mTimeframes.addHistorical(mPeriodLength, _
                                mPeriodUnits, _
                                "", _
                                mInitialNumberOfBars, _
                                mFromTime, _
                                mToTime, _
                                mIncludeBarsOutsideSession)
Else
    Set mTimeframe = mTimeframes.add(mPeriodLength, _
                                mPeriodUnits, _
                                "", _
                                mInitialNumberOfBars, _
                                mIncludeBarsOutsideSession, _
                                IIf(mTicker.replayingTickfile, True, False))
End If

If Not mTimeframe.historicDataLoaded Then
    Chart1.suppressDrawing = True
End If

showStudies

End Sub

Private Sub showStudies()
Dim tcs() As TaskController
Dim tc As TaskController
Dim tcMaxIndex As Long
Dim i As Long

Set mBarsStudyConfig = createBarsStudyConfig

tcs = mChartManager.setupStudyValueListeners(mBarsStudyConfig)

tcMaxIndex = -1
On Error Resume Next
tcMaxIndex = UBound(tcs)
On Error GoTo 0

If tcMaxIndex <> -1 Then
    For i = 0 To tcMaxIndex
        Set tc = tcs(i)
        tc.addTaskCompletionListener Me
        mNumberOfOutstandingTasks = mNumberOfOutstandingTasks + 1
    Next
End If

End Sub
