VERSION 5.00
Object = "{015212C3-04F2-4693-B20B-0BEB304EFC1B}#8.5#0"; "ChartSkil2-5.ocx"
Begin VB.UserControl TradeBuildChart 
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   ScaleHeight     =   5745
   ScaleWidth      =   7365
   ToolboxBitmap   =   "TradeBuildChart.ctx":0000
   Begin ChartSkil25.Chart Chart1 
      Align           =   1  'Align Top
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   9128
      PointerDiscColor=   794110
   End
End
Attribute VB_Name = "TradeBuildChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TaskCompletionListener

'================================================================================
' Events
'================================================================================

Event StateChange(ev As StateChangeEvent)

'================================================================================
' Enums
'================================================================================

Public Enum ChartStates
    ''
    ' The chart is ready to have studies added to it
    '
    '@/
    ChartStateInitialised
    
    ''
    ' All historic data has been added to the chart
    '
    '@/
    ChartStateLoaded
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Constants
'================================================================================

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

'================================================================================
' Member variables
'================================================================================

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

Private mHighPrice As Double
Private mLowPrice As Double
Private mPrevClosePrice As Double

Private mPrevWidth As Single
Private mPrevHeight As Single

Private mNumberOfOutstandingTasks As Long
Private mHistDataLoaded As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

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
DefaultRegionAutoscale = PropDfltDfltRegnStyleAutoscale
DefaultRegionBackColor = PropDfltDfltRegnStyleBackColor
DefaultRegionGridColor = PropDfltDfltRegnStyleGridColor
DefaultRegionGridlineSpacingY = PropDfltDfltRegnStyleGridlineSpacingY
DefaultRegionGridTextColor = PropDfltDfltRegnStyleGridTextColor
DefaultRegionHasGrid = PropDfltDfltRegnStyleHasGrid
DefaultRegionHasGridText = PropDfltDfltRegnStyleHasGridtext
DefaultRegionPointerStyle = PropDfltDfltRegnStylePointerStyle
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

DefaultRegionAutoscale = PropBag.ReadProperty(PropNameDfltRegnStyleAutoscale, PropDfltDfltRegnStyleAutoscale)
If Err.Number <> 0 Then
    DefaultRegionAutoscale = PropDfltDfltRegnStyleAutoscale
    Err.clear
End If

DefaultRegionBackColor = PropBag.ReadProperty(PropNameDfltRegnStyleBackColor, PropDfltDfltRegnStyleBackColor)
If Err.Number <> 0 Then
    DefaultRegionBackColor = PropDfltDfltRegnStyleBackColor
    Err.clear
End If

DefaultRegionGridColor = PropBag.ReadProperty(PropNameDfltRegnStyleGridColor, PropDfltDfltRegnStyleGridColor)
If Err.Number <> 0 Then
    DefaultRegionGridColor = PropDfltDfltRegnStyleGridColor
    Err.clear
End If

DefaultRegionGridlineSpacingY = PropBag.ReadProperty(PropNameDfltRegnStyleGridlineSpacingY, PropDfltDfltRegnStyleGridlineSpacingY)
If Err.Number <> 0 Then
    DefaultRegionGridlineSpacingY = PropDfltDfltRegnStyleGridlineSpacingY
    Err.clear
End If

DefaultRegionGridTextColor = PropBag.ReadProperty(PropNameDfltRegnStyleGridTextColor, PropDfltDfltRegnStyleGridTextColor)
If Err.Number <> 0 Then
    DefaultRegionGridTextColor = PropDfltDfltRegnStyleGridTextColor
    Err.clear
End If

DefaultRegionHasGrid = PropBag.ReadProperty(PropNameDfltRegnStyleHasGrid, PropDfltDfltRegnStyleHasGrid)
If Err.Number <> 0 Then
    DefaultRegionHasGrid = PropDfltDfltRegnStyleHasGrid
    Err.clear
End If

DefaultRegionHasGridText = PropBag.ReadProperty(PropNameDfltRegnStyleHasGridtext, PropDfltDfltRegnStyleHasGridtext)
If Err.Number <> 0 Then
    DefaultRegionHasGridText = PropDfltDfltRegnStyleHasGridtext
    Err.clear
End If

DefaultRegionPointerStyle = PropBag.ReadProperty(PropNameDfltRegnStylePointerStyle, PropDfltDfltRegnStylePointerStyle)
If Err.Number <> 0 Then
    DefaultRegionPointerStyle = PropDfltDfltRegnStylePointerStyle
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
PropBag.WriteProperty PropNameDfltRegnStyleAutoscale, DefaultRegionAutoscale, PropDfltDfltRegnStyleAutoscale
PropBag.WriteProperty PropNameDfltRegnStyleBackColor, DefaultRegionBackColor, PropDfltDfltRegnStyleBackColor
PropBag.WriteProperty PropNameDfltRegnStyleGridColor, DefaultRegionGridColor, PropDfltDfltRegnStyleGridColor
PropBag.WriteProperty PropNameDfltRegnStyleGridlineSpacingY, DefaultRegionGridlineSpacingY, PropDfltDfltRegnStyleGridlineSpacingY
PropBag.WriteProperty PropNameDfltRegnStyleGridTextColor, DefaultRegionGridTextColor, PropDfltDfltRegnStyleGridTextColor
PropBag.WriteProperty PropNameDfltRegnStyleHasGrid, DefaultRegionHasGrid, PropDfltDfltRegnStyleHasGrid
PropBag.WriteProperty PropNameDfltRegnStyleHasGridtext, DefaultRegionHasGridText, PropDfltDfltRegnStyleHasGridtext
PropBag.WriteProperty PropNameDfltRegnStylePointerStyle, DefaultRegionPointerStyle, PropDfltDfltRegnStylePointerStyle
PropBag.WriteProperty PropNamePointerCrosshairsColor, PointerCrosshairsColor, PropDfltPointerCrosshairsColor
PropBag.WriteProperty PropNamePointerDiscColor, PointerDiscColor, PropDfltPointerDiscColor
PropBag.WriteProperty PropNameShowHorizontalScrollBar, showHorizontalScrollBar, PropDfltShowHorizontalScrollBar
PropBag.WriteProperty PropNameShowToolbar, showToolbar, PropDfltShowToolbar
PropBag.WriteProperty PropNameTwipsPerBar, twipsPerBar, PropDfltTwipsPerBar
PropBag.WriteProperty PropNameYAxisWidthCm, YAxisWidthCm, PropDfltYAxisWidthCm
End Sub

'================================================================================
' TaskCompletionListener Interface Members
'================================================================================

Private Sub TaskCompletionListener_taskCompleted(ev As Tasks.TaskCompletionEvent)
Dim stateEv As StateChangeEvent

mNumberOfOutstandingTasks = mNumberOfOutstandingTasks - 1
If mNumberOfOutstandingTasks = 0 And mHistDataLoaded Then
    stateEv.State = ChartStates.ChartStateLoaded
    Set stateEv.Source = Me
    RaiseEvent StateChange(stateEv)
End If
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_StateChange(ev As StateChangeEvent)
Dim stateEv As StateChangeEvent
If ev.State = TickerStates.TickerStateReady Then
    ' this means that the ticker object has retrieved the contract info, so we can
    ' now start the chart
    loadchart
    stateEv.State = ChartStates.ChartStateInitialised
    Set stateEv.Source = Me
    RaiseEvent StateChange(stateEv)
End If
End Sub

'================================================================================
' mTimeframe Event Handlers
'================================================================================

Private Sub mTimeframe_BarsLoaded()
Dim stateEv As StateChangeEvent

Chart1.suppressDrawing = False

mHistDataLoaded = True

If mNumberOfOutstandingTasks = 0 Then
    stateEv.State = ChartStates.ChartStateLoaded
    Set stateEv.Source = Me
    RaiseEvent StateChange(stateEv)
End If
End Sub

'================================================================================
' Properties
'================================================================================

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

Public Property Get DefaultRegionAutoscale() As Boolean
Attribute DefaultRegionAutoscale.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
DefaultRegionAutoscale = Chart1.DefaultRegionAutoscale
End Property

Public Property Let DefaultRegionAutoscale(ByVal value As Boolean)
Chart1.DefaultRegionAutoscale = value
End Property

Public Property Get DefaultRegionBackColor() As OLE_COLOR
Attribute DefaultRegionBackColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
DefaultRegionBackColor = Chart1.DefaultRegionBackColor
End Property

Public Property Let DefaultRegionBackColor(ByVal val As OLE_COLOR)
Chart1.DefaultRegionBackColor = val
End Property

Public Property Get DefaultRegionGridColor() As OLE_COLOR
Attribute DefaultRegionGridColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
DefaultRegionGridColor = Chart1.DefaultRegionGridColor
End Property

Public Property Let DefaultRegionGridColor(ByVal val As OLE_COLOR)
Chart1.DefaultRegionGridColor = val
End Property

Public Property Get DefaultRegionGridlineSpacingY() As Double
Attribute DefaultRegionGridlineSpacingY.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
DefaultRegionGridlineSpacingY = Chart1.DefaultRegionGridlineSpacingY
End Property

Public Property Let DefaultRegionGridlineSpacingY(ByVal value As Double)
Chart1.DefaultRegionGridlineSpacingY = value
End Property

Public Property Get DefaultRegionGridTextColor() As OLE_COLOR
Attribute DefaultRegionGridTextColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
DefaultRegionGridTextColor = Chart1.DefaultRegionGridTextColor
End Property

Public Property Let DefaultRegionGridTextColor(ByVal val As OLE_COLOR)
Chart1.DefaultRegionGridTextColor = val
End Property

Public Property Get DefaultRegionHasGrid() As Boolean
Attribute DefaultRegionHasGrid.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
DefaultRegionHasGrid = Chart1.DefaultRegionHasGrid
End Property

Public Property Let DefaultRegionHasGrid(ByVal val As Boolean)
Chart1.DefaultRegionHasGrid = val
End Property

Public Property Get DefaultRegionHasGridText() As Boolean
Attribute DefaultRegionHasGridText.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
DefaultRegionHasGridText = Chart1.DefaultRegionHasGridText
End Property

Public Property Let DefaultRegionHasGridText(ByVal val As Boolean)
Chart1.DefaultRegionHasGridText = val
End Property

Public Property Get DefaultRegionPointerStyle() As PointerStyles
Attribute DefaultRegionPointerStyle.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
DefaultRegionPointerStyle = Chart1.DefaultRegionPointerStyle
End Property

Public Property Let DefaultRegionPointerStyle(ByVal value As PointerStyles)
Chart1.DefaultRegionPointerStyle = value
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
Dim units As String
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

Public Property Get YAxisWidthCm() As Single
Attribute YAxisWidthCm.VB_ProcData.VB_Invoke_Property = ";Appearance"
YAxisWidthCm = Chart1.YAxisWidthCm
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
Chart1.YAxisWidthCm = value
End Property

'================================================================================
' Methods
'================================================================================

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
Set mTicker = pTicker
mInitialNumberOfBars = initialNumberOfBars
mIncludeBarsOutsideSession = includeBarsOutsideSession
mMinimumTicksHeight = minimumTicksHeight
mPeriodLength = periodlength
mPeriodUnits = periodUnits
Set mPriceRegionStyle = priceRegionStyle
Set mVolumeRegionStyle = volumeRegionStyle

Set mChartManager = createChartManager(mTicker.studyManager, Chart1.controller)

If mTicker.State = TickerStateRunning Then
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

Set mChartManager = createChartManager(mTicker.studyManager, Chart1.controller)

If mTicker.State = TickerStateReady Then
    loadchart
End If
End Sub

Public Sub showStudyPickerForm()
If mTicker.State = TickerStateRunning Then showStudyPicker mChartManager
End Sub

Public Sub syncStudyPickerForm()
If mTicker.State = TickerStateRunning Then syncStudyPicker mChartManager
End Sub

Public Sub unsyncStudyPickerForm()
unsyncStudyPicker
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function createBarsStudyConfig() As StudyConfiguration
Dim lStudy As study
Dim studyDef As StudyDefinition
ReDim inputValueNames(1) As String
Dim params As New Parameters2.Parameters
Dim studyValueConfig As StudyValueConfiguration

Set createBarsStudyConfig = New StudyConfiguration

createBarsStudyConfig.underlyingStudy = mTicker.inputStudy

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
studyValueConfig.outlineThickness = 1
studyValueConfig.barThickness = 2
studyValueConfig.barWidth = 0.6
studyValueConfig.chartRegionName = RegionNamePrice
studyValueConfig.barDisplayMode = BarDisplayModeCandlestick
studyValueConfig.downColor = &H43FC2
studyValueConfig.includeInAutoscale = True
studyValueConfig.includeInChart = True
studyValueConfig.layer = 200
studyValueConfig.solidUpBody = True
studyValueConfig.tailThickness = 2
studyValueConfig.upColor = &H1D9311

Set studyValueConfig = createBarsStudyConfig.StudyValueConfigurations.add("Volume")
studyValueConfig.chartRegionName = RegionNameVolume
studyValueConfig.Color = vbBlack
studyValueConfig.dataPointDisplayMode = DataPointDisplayModeHistogram
studyValueConfig.histogramBarWidth = 0.5
studyValueConfig.includeInAutoscale = True
studyValueConfig.includeInChart = True
studyValueConfig.lineThickness = 1
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
Dim tcs() As TaskCompletion
Dim tc As TaskCompletion
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
