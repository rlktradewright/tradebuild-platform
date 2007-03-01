VERSION 5.00
Object = "{015212C3-04F2-4693-B20B-0BEB304EFC1B}#5.1#0"; "ChartSkil2-5.ocx"
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

''
'
' Raised when the initial bar study listeners have completed loading historical
' data to the chart,
'@/

'================================================================================
' Constants
'================================================================================

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

Private mInitialNumberOfBars As Long
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

Private Sub UserControl_Resize()
If UserControl.Width <> mPrevWidth Then
    mPrevWidth = UserControl.Width
End If
If UserControl.Height <> mPrevHeight Then
    Chart1.Height = UserControl.Height
    mPrevHeight = UserControl.Height
End If
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
If ev.State = TickerStates.TickerStateRunning Then
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

Public Property Get chartController() As chartController
Set chartController = Chart1.controller
End Property

Public Property Get chartManager() As chartManager
Set chartManager = mChartManager
End Property

Public Property Get initialNumberOfBars() As Long
initialNumberOfBars = mInitialNumberOfBars
End Property

Public Property Get minimumTicksHeight() As Double
minimumTicksHeight = mMinimumTicksHeight
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

Public Property Let updatePerTick(ByVal value As Boolean)
mUpdatePerTick = value
End Property

Public Property Let volumeRegionStyle(ByVal value As ChartRegionStyle)
Set mVolumeRegionStyle = value
If Not mVolumeRegion Is Nothing Then mVolumeRegion.Style = value
End Property

Public Property Get volumeRegionStyle() As ChartRegionStyle
Set volumeRegionStyle = mVolumeRegionStyle
End Property

'================================================================================
' Methods
'================================================================================

Public Sub clearChart()
Chart1.clearChart
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
Set mChartController = mChartManager.chartController

If mTicker.State = TickerStateRunning Then
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
studyValueConfig.tailThickness = 1
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

Chart1.suppressDrawing = True

Chart1.clearChart
Chart1.twipsPerBar = 120
Chart1.showHorizontalScrollBar = True

Chart1.sessionStartTime = mContract.sessionStartTime
Chart1.sessionEndTime = mContract.sessionEndTime

Chart1.setPeriodParameters mPeriodLength, mPeriodUnits

Set regionStyle = Chart1.defaultRegionStyle
regionStyle.backColor = vbWhite
regionStyle.autoscale = True
regionStyle.hasGrid = True
regionStyle.pointerStyle = PointerCrosshairs

Chart1.defaultRegionStyle = regionStyle

If Not mPriceRegionStyle Is Nothing Then
    Set regionStyle = mPriceRegionStyle
Else
    Set regionStyle = Chart1.defaultRegionStyle
    regionStyle.gridlineSpacingY = 2
    regionStyle.YScaleQuantum = mContract.ticksize
    If mMinimumTicksHeight * mContract.ticksize <> 0 Then
        regionStyle.minimumHeight = mMinimumTicksHeight * mContract.ticksize
    End If

End If

Set mPriceRegion = mChartController.addChartRegion(100, 25, regionStyle, RegionNamePrice)

mPriceRegion.setTitle mContract.specifier.localSymbol & _
                " (" & mContract.specifier.exchange & ") " & _
                timeframeCaption, _
                vbBlue, _
                Nothing

If Not mVolumeRegionStyle Is Nothing Then
    Set regionStyle = mVolumeRegionStyle
Else
    Set regionStyle = Chart1.defaultRegionStyle
    regionStyle.gridlineSpacingY = 0.8
    regionStyle.minimumHeight = 10
    regionStyle.integerYScale = True
End If

Set mVolumeRegion = mChartController.addChartRegion(20, , regionStyle, RegionNameVolume)

mVolumeRegion.setTitle "Volume", vbBlue, Nothing

Chart1.suppressDrawing = False

End Sub

Private Sub loadchart()

Set mContract = mTicker.Contract

initialiseChart

Set mTimeframes = mTicker.Timeframes

Set mTimeframe = mTimeframes.add(mPeriodLength, _
                            mPeriodUnits, _
                            "", _
                            mInitialNumberOfBars, _
                            mIncludeBarsOutsideSession, _
                            IIf(mTicker.replayingTickfile, True, False))
                            
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
