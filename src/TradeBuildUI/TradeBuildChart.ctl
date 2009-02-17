VERSION 5.00
Object = "{74951842-2BEF-4829-A34F-DC7795A37167}#77.0#0"; "ChartSkil2-6.ocx"
Begin VB.UserControl TradeBuildChart 
   Alignable       =   -1  'True
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   ScaleHeight     =   5475
   ScaleWidth      =   7740
   ToolboxBitmap   =   "TradeBuildChart.ctx":0000
   Begin ChartSkil26.Chart Chart1 
      Align           =   1  'Align Top
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   8705
      ChartBackColor  =   6566450
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

Event KeyDown(KeyCode As Integer, Shift As Integer)

Event KeyPress(KeyAscii As Integer)

Event KeyUp(KeyCode As Integer, Shift As Integer)

Event StateChange(ev As StateChangeEvent)

Event TimeframeChange()

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
Private Const PropNameDefaultBarDisplayMode             As String = "DefaultBarDisplayMode"
Private Const PropNameDfltRegnStyleAutoscale            As String = "DfltRegnStyleAutoscale"
Private Const PropNameDfltRegnStyleBackColor            As String = "DfltRegnStyleBackColor"
Private Const PropNameDfltRegnStyleGridColor            As String = "DfltRegnStyleGridColor"
Private Const PropNameDfltRegnStyleGridlineSpacingY     As String = "DfltRegnStyleGridlineSpacingY"
Private Const PropNameDfltRegnStyleGridTextColor        As String = "DfltRegnStyleGridTextColor"
Private Const PropNameDfltRegnStyleHasGrid              As String = "DfltRegnStyleHasGrid"
Private Const PropNameDfltRegnStyleHasGridtext          As String = "DfltRegnStyleHasGridtext"
Private Const PropNamePointerDiscColor                  As String = "PointerDiscColor"
Private Const PropNamePointerCrosshairsColor            As String = "PointerCrosshairsColor"
Private Const PropNamePointerStyle                      As String = "PointerStyle"
Private Const PropNameShowHorizontalScrollBar           As String = "ShowHorizontalScrollBar"
Private Const PropNameShowToolbar                       As String = "ShowToobar"
Private Const PropNameTwipsPerBar                       As String = "TwipsPerBar"
Private Const PropNameYAxisWidthCm                      As String = "YAxisWidthCm"

Private Const PropDfltAllowHorizontalMouseScrolling     As Boolean = True
Private Const PropDfltAllowVerticalMouseScrolling       As Boolean = True
Private Const PropDfltAutoscroll                        As Boolean = True
Private Const PropDfltChartBackColor                    As Long = vbWhite
Private Const PropDfltDefaultBarDisplayMode             As Long = BarDisplayModes.BarDisplayModeBar
Private Const PropDfltDfltRegnStyleAutoscale            As Boolean = True
Private Const PropDfltDfltRegnStyleBackColor            As Long = vbWhite
Private Const PropDfltDfltRegnStyleGridColor            As Long = &HC0C0C0
Private Const PropDfltDfltRegnStyleGridlineSpacingY     As Double = 1.8
Private Const PropDfltDfltRegnStyleGridTextColor        As Long = vbBlack
Private Const PropDfltDfltRegnStyleHasGrid              As Boolean = True
Private Const PropDfltDfltRegnStyleHasGridtext          As Boolean = False
Private Const PropDfltPointerDiscColor                  As Long = &H89FFFF
Private Const PropDfltPointerCrosshairsColor            As Long = &HC1DFE
Private Const PropDfltPointerStyle                      As Long = PointerStyles.PointerCrosshairs
Private Const PropDfltShowHorizontalScrollBar           As Boolean = True
Private Const PropDfltShowToolbar                       As Boolean = True
Private Const PropDfltTwipsPerBar                       As Long = 150
Private Const PropDfltYAxisWidthCm                      As Single = 1.3

'@================================================================================
' Member variables
'@================================================================================

Private mManager                                        As chartManager
Private mController                                     As chartController

Private WithEvents mTicker                              As ticker
Attribute mTicker.VB_VarHelpID = -1
Private mTimeframes                                     As Timeframes
Private WithEvents mTimeframe                           As Timeframe
Attribute mTimeframe.VB_VarHelpID = -1

Private mUpdatePerTick                                  As Boolean

Private mState                                          As ChartStates

Private mBarsStudyConfig                                As StudyConfiguration

Private mIsHistoricChart                                As Boolean

Private mChartSpec                                      As ChartSpecifier

Private mFromTime                                       As Date
Private mToTime                                         As Date

Private mContract                                       As Contract

Private mPriceRegion                                    As ChartRegion

Private mVolumeRegion                                   As ChartRegion

Private mPrevWidth                                      As Single
Private mPrevHeight                                     As Single

Private mNumberOfOutstandingTasks                       As Long
Private mHistDataLoaded                                 As Boolean

Private mLoadingText                                    As Text

Private mBarFormatterFactory                            As BarFormatterFactory

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
ChartBackColor = PropDfltChartBackColor
DefaultBarDisplayMode = PropDfltDefaultBarDisplayMode
RegionDefaultAutoscale = PropDfltDfltRegnStyleAutoscale
RegionDefaultBackColor = PropDfltDfltRegnStyleBackColor
RegionDefaultGridColor = PropDfltDfltRegnStyleGridColor
RegionDefaultGridlineSpacingY = PropDfltDfltRegnStyleGridlineSpacingY
RegionDefaultGridTextColor = PropDfltDfltRegnStyleGridTextColor
RegionDefaultHasGrid = PropDfltDfltRegnStyleHasGrid
RegionDefaultHasGridText = PropDfltDfltRegnStyleHasGridtext
PointerStyle = PropDfltPointerStyle
PointerCrosshairsColor = PropDfltPointerCrosshairsColor
PointerDiscColor = PropDfltPointerDiscColor
showHorizontalScrollBar = PropDfltShowHorizontalScrollBar
showToolbar = PropDfltShowToolbar
twipsPerBar = PropDfltTwipsPerBar
YAxisWidthCm = PropDfltYAxisWidthCm

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

ChartBackColor = PropBag.ReadProperty(PropNameChartBackColor)
' if no ChartBackColor has been set, we'll just use the ChartSkil default
If Err.Number <> 0 Then Err.clear

DefaultBarDisplayMode = PropBag.ReadProperty(PropNameDefaultBarDisplayMode, PropDfltDefaultBarDisplayMode)
If Err.Number <> 0 Then
    DefaultBarDisplayMode = PropDfltDefaultBarDisplayMode
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

PointerStyle = PropBag.ReadProperty(PropNamePointerStyle, PropDfltPointerStyle)
If Err.Number <> 0 Then
    PointerStyle = PropDfltPointerStyle
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

End Sub

Private Sub UserControl_Resize()
'If UserControl.Width <> mPrevWidth Then
    mPrevWidth = UserControl.Width
'End If
'If UserControl.Height <> mPrevHeight Then
    Chart1.Height = UserControl.Height
    mPrevHeight = UserControl.Height
'End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty PropNameAllowHorizontalMouseScrolling, AllowHorizontalMouseScrolling, PropDfltAllowHorizontalMouseScrolling
PropBag.WriteProperty PropNameAllowVerticalMouseScrolling, AllowVerticalMouseScrolling, PropDfltAllowVerticalMouseScrolling
PropBag.WriteProperty PropNameAutoscroll, Autoscroll, PropDfltAutoscroll
PropBag.WriteProperty PropNameChartBackColor, ChartBackColor, PropDfltChartBackColor
PropBag.WriteProperty PropNameDefaultBarDisplayMode, DefaultBarDisplayMode, PropDfltDefaultBarDisplayMode
PropBag.WriteProperty PropNameDfltRegnStyleAutoscale, RegionDefaultAutoscale, PropDfltDfltRegnStyleAutoscale
PropBag.WriteProperty PropNameDfltRegnStyleBackColor, RegionDefaultBackColor, PropDfltDfltRegnStyleBackColor
PropBag.WriteProperty PropNameDfltRegnStyleGridColor, RegionDefaultGridColor, PropDfltDfltRegnStyleGridColor
PropBag.WriteProperty PropNameDfltRegnStyleGridlineSpacingY, RegionDefaultGridlineSpacingY, PropDfltDfltRegnStyleGridlineSpacingY
PropBag.WriteProperty PropNameDfltRegnStyleGridTextColor, RegionDefaultGridTextColor, PropDfltDfltRegnStyleGridTextColor
PropBag.WriteProperty PropNameDfltRegnStyleHasGrid, RegionDefaultHasGrid, PropDfltDfltRegnStyleHasGrid
PropBag.WriteProperty PropNameDfltRegnStyleHasGridtext, RegionDefaultHasGridText, PropDfltDfltRegnStyleHasGridtext
PropBag.WriteProperty PropNamePointerStyle, PointerStyle, PropDfltPointerStyle
PropBag.WriteProperty PropNamePointerCrosshairsColor, PointerCrosshairsColor, PropDfltPointerCrosshairsColor
PropBag.WriteProperty PropNamePointerDiscColor, PointerDiscColor, PropDfltPointerDiscColor
PropBag.WriteProperty PropNameShowHorizontalScrollBar, showHorizontalScrollBar, PropDfltShowHorizontalScrollBar
PropBag.WriteProperty PropNameShowToolbar, showToolbar, PropDfltShowToolbar
PropBag.WriteProperty PropNameTwipsPerBar, twipsPerBar, PropDfltTwipsPerBar
PropBag.WriteProperty PropNameYAxisWidthCm, YAxisWidthCm, PropDfltYAxisWidthCm
End Sub

'@================================================================================
' Chart1 Event Handlers
'@================================================================================

Private Sub Chart1_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Chart1_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Chart1_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_taskCompleted(ev As TaskCompletionEvent)
mNumberOfOutstandingTasks = mNumberOfOutstandingTasks - 1
If mNumberOfOutstandingTasks = 0 And mHistDataLoaded Then
    setState ChartStates.ChartStateLoaded
End If
End Sub

'@================================================================================
' mTicker Event Handlers
'@================================================================================

Private Sub mTicker_StateChange(ev As StateChangeEvent)
If ev.State = TickerStates.TickerStateReady Then
    ' this means that the ticker object has retrieved the contract info, so we can
    ' now start the chart
    loadchart
    setState ChartStates.ChartStateInitialised
End If
End Sub

'@================================================================================
' mTimeframe Event Handlers
'@================================================================================

Private Sub mTimeframe_BarsLoaded()
Chart1.SuppressDrawing = False

mHistDataLoaded = True

If mNumberOfOutstandingTasks = 0 Then
    setState ChartStates.ChartStateLoaded
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
AllowHorizontalMouseScrolling = Chart1.AllowHorizontalMouseScrolling
End Property

Public Property Let AllowVerticalMouseScrolling( _
                ByVal value As Boolean)
Chart1.AllowVerticalMouseScrolling = value
End Property

Public Property Get AllowVerticalMouseScrolling() As Boolean
AllowVerticalMouseScrolling = Chart1.AllowVerticalMouseScrolling
End Property

Public Property Let Autoscroll( _
                ByVal value As Boolean)
Chart1.Autoscroll = value
End Property

Public Property Get Autoscroll() As Boolean
Autoscroll = Chart1.Autoscroll
End Property

Public Property Get ChartBackColor() As OLE_COLOR
ChartBackColor = Chart1.ChartBackColor
End Property

Public Property Let ChartBackColor(ByVal val As OLE_COLOR)
Chart1.ChartBackColor = val
End Property

Public Property Get ChartBackGradientFillColors() As Long()
ChartBackGradientFillColors = Chart1.ChartBackGradientFillColors
End Property

Public Property Let ChartBackGradientFillColors(ByRef value() As Long)
Dim ar() As Long
ar = value
Chart1.controller.ChartBackGradientFillColors = ar
End Property

Public Property Get chartController() As chartController
Set chartController = Chart1.controller
End Property

Public Property Get chartManager() As chartManager
Set chartManager = mManager
End Property

Public Property Get DefaultBarDisplayMode() As BarDisplayModes
DefaultBarDisplayMode = Chart1.DefaultBarDisplayMode
End Property

Public Property Let DefaultBarDisplayMode( _
                ByVal value As BarDisplayModes)
Chart1.DefaultBarDisplayMode = value
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
UserControl.Enabled = value
PropertyChanged "Enabled"
End Property

Public Property Get InitialNumberOfBars() As Long
Attribute InitialNumberOfBars.VB_ProcData.VB_Invoke_Property = ";Behavior"
InitialNumberOfBars = mChartSpec.InitialNumberOfBars
End Property

Public Property Get MinimumTicksHeight() As Double
Attribute MinimumTicksHeight.VB_ProcData.VB_Invoke_Property = ";Behavior"
MinimumTicksHeight = mChartSpec.MinimumTicksHeight
End Property

Public Property Get PointerCrosshairsColor() As OLE_COLOR
PointerCrosshairsColor = Chart1.PointerCrosshairsColor
End Property

Public Property Let PointerCrosshairsColor(ByVal value As OLE_COLOR)
Chart1.PointerCrosshairsColor = value
End Property

Public Property Get PointerDiscColor() As OLE_COLOR
PointerDiscColor = Chart1.PointerDiscColor
End Property

Public Property Let PointerDiscColor(ByVal value As OLE_COLOR)
Chart1.PointerDiscColor = value
End Property

Public Property Get PointerStyle() As PointerStyles
PointerStyle = Chart1.PointerStyle
End Property

Public Property Let PointerStyle(ByVal value As PointerStyles)
Chart1.PointerStyle = value
End Property

Public Property Get priceRegion() As ChartRegion
Set priceRegion = mPriceRegion
End Property

Public Property Get RegionDefaultAutoscale() As Boolean
RegionDefaultAutoscale = Chart1.RegionDefaultAutoscale
End Property

Public Property Let RegionDefaultAutoscale(ByVal value As Boolean)
Chart1.RegionDefaultAutoscale = value
End Property

Public Property Get RegionDefaultBackColor() As OLE_COLOR
RegionDefaultBackColor = Chart1.RegionDefaultBackColor
End Property

Public Property Let RegionDefaultBackColor(ByVal val As OLE_COLOR)
Chart1.RegionDefaultBackColor = val
End Property

Public Property Get RegionDefaultGridColor() As OLE_COLOR
RegionDefaultGridColor = Chart1.RegionDefaultGridColor
End Property

Public Property Let RegionDefaultGridColor(ByVal val As OLE_COLOR)
Chart1.RegionDefaultGridColor = val
End Property

Public Property Get RegionDefaultGridlineSpacingY() As Double
RegionDefaultGridlineSpacingY = Chart1.RegionDefaultGridlineSpacingY
End Property

Public Property Let RegionDefaultGridlineSpacingY(ByVal value As Double)
Chart1.RegionDefaultGridlineSpacingY = value
End Property

Public Property Get RegionDefaultGridTextColor() As OLE_COLOR
RegionDefaultGridTextColor = Chart1.RegionDefaultGridTextColor
End Property

Public Property Let RegionDefaultGridTextColor(ByVal val As OLE_COLOR)
Chart1.RegionDefaultGridTextColor = val
End Property

Public Property Get RegionDefaultHasGrid() As Boolean
RegionDefaultHasGrid = Chart1.RegionDefaultHasGrid
End Property

Public Property Let RegionDefaultHasGrid(ByVal val As Boolean)
Chart1.RegionDefaultHasGrid = val
End Property

Public Property Get RegionDefaultHasGridText() As Boolean
RegionDefaultHasGridText = Chart1.RegionDefaultHasGridText
End Property

Public Property Let RegionDefaultHasGridText(ByVal val As Boolean)
Chart1.RegionDefaultHasGridText = val
End Property

Public Property Get regionNames() As String()
regionNames = mManager.regionNames
End Property

Public Property Get showHorizontalScrollBar() As Boolean
showHorizontalScrollBar = Chart1.showHorizontalScrollBar
End Property

Public Property Let showHorizontalScrollBar(ByVal val As Boolean)
Chart1.showHorizontalScrollBar = val
End Property

Public Property Get showToolbar() As Boolean
showToolbar = Chart1.showToolbar
End Property

Public Property Let showToolbar(ByVal val As Boolean)
Chart1.showToolbar = val
End Property

Public Property Get State() As ChartStates
State = mState
End Property

Public Property Get timeframeCaption() As String
timeframeCaption = mChartSpec.Timeframe.toString
End Property

Public Property Get timeframeShortCaption() As String
timeframeShortCaption = mChartSpec.Timeframe.toShortString
End Property

Public Property Get Timeframe() As Timeframe
Set Timeframe = mTimeframe
End Property

Public Property Get TimePeriod() As TimePeriod
Set TimePeriod = mChartSpec.Timeframe
End Property

Public Property Get twipsPerBar() As Long
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

'Public Property Let VolumeRegionStyle(ByVal value As ChartRegionStyle)
'Set mVolumeRegionStyle = value
'If Not mVolumeRegion Is Nothing Then mVolumeRegion.Style = value
'End Property

Public Property Get VolumeRegionStyle() As ChartRegionStyle
Set VolumeRegionStyle = mChartSpec.VolumeRegionStyle
End Property

'Public Property Let VolumeStyle(ByVal value As dataPointStyle)
'Set mVolumeStyle = value
'End Property

Public Property Get YAxisWidthCm() As Single
YAxisWidthCm = Chart1.YAxisWidthCm
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
Chart1.YAxisWidthCm = value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ChangeTimeframe(ByVal Timeframe As TimePeriod)
Dim oldBaseStudyConfig As StudyConfiguration: Set oldBaseStudyConfig = mBarsStudyConfig
Dim oldStudyConfigs As StudyConfigurations: Set oldStudyConfigs = mManager.StudyConfigurations

finish

mChartSpec.Timeframe = Timeframe

initialiseChart

setState ChartStates.ChartStateCreated

If mTicker.State = TickerStates.TickerStateReady Or _
    mTicker.State = TickerStates.TickerStateRunning _
Then
    loadchart
    setState ChartStates.ChartStateInitialised
End If

reconfigureDependingStudies oldStudyConfigs, oldBaseStudyConfig, mBarsStudyConfig

RaiseEvent TimeframeChange
End Sub

Public Sub finish()
Chart1.clearChart
mManager.finish

Set mManager = Nothing
Set mController = Nothing

Set mTimeframes = Nothing
Set mTimeframe = Nothing

Set mBarsStudyConfig = Nothing

Set mContract = Nothing

Set mPriceRegion = Nothing

Set mVolumeRegion = Nothing

Set mLoadingText = Nothing

End Sub

Public Sub scrollToTime(ByVal pTime As Date)
mManager.scrollToTime pTime
End Sub

Public Sub showChart( _
                ByVal pTicker As ticker, _
                ByVal chartSpec As ChartSpecifier, _
                Optional ByVal BarFormatterFactory As BarFormatterFactory)

Select Case chartSpec.Timeframe.units
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

prepareChart pTicker, chartSpec, BarFormatterFactory

End Sub

Public Sub showHistoricChart( _
                ByVal pTicker As ticker, _
                ByVal chartSpec As ChartSpecifier, _
                ByVal fromTime As Date, _
                ByVal toTime As Date, _
                Optional ByVal BarFormatterFactory As BarFormatterFactory)

Select Case chartSpec.Timeframe.units
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
mFromTime = fromTime
mToTime = toTime
prepareChart pTicker, chartSpec, BarFormatterFactory

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createBarsStudyConfig() As StudyConfiguration
Dim lStudy As Study
Dim studyDef As StudyDefinition

ReDim inputValueNames(3) As String
Dim params As New Parameters

Dim studyValueConfig As StudyValueConfiguration
Dim BarsStyle As barStyle
Dim VolumeStyle As dataPointStyle

Set createBarsStudyConfig = New StudyConfiguration

createBarsStudyConfig.underlyingStudy = mTicker.InputStudy

Set lStudy = mTimeframe.tradeStudy
createBarsStudyConfig.Study = lStudy
Set studyDef = lStudy.StudyDefinition

createBarsStudyConfig.chartRegionName = RegionNamePrice
inputValueNames(0) = mTicker.InputNameTrade
inputValueNames(1) = mTicker.InputNameVolume
inputValueNames(2) = mTicker.InputNameTickVolume
inputValueNames(3) = mTicker.InputNameOpenInterest
createBarsStudyConfig.inputValueNames = inputValueNames
createBarsStudyConfig.name = studyDef.name
params.setParameterValue "Bar length", mChartSpec.Timeframe.length
params.setParameterValue "Time units", TimePeriodUnitsToString(mChartSpec.Timeframe.units)
createBarsStudyConfig.Parameters = params

Set studyValueConfig = createBarsStudyConfig.StudyValueConfigurations.add("Bar")
studyValueConfig.chartRegionName = RegionNamePrice
studyValueConfig.includeInChart = True
studyValueConfig.layer = 200
studyValueConfig.SetBarFormatterFactory mBarFormatterFactory, mTimeframe.tradeBars

If Not mChartSpec.BarsStyle Is Nothing Then
    Set BarsStyle = mChartSpec.BarsStyle
Else
    Set BarsStyle = mPriceRegion.DefaultBarStyle
    BarsStyle.displayMode = BarDisplayModes.BarDisplayModeCandlestick
    BarsStyle.outlineThickness = 1
    BarsStyle.tailThickness = 1
    BarsStyle.upColor = &HA0A0A0
    BarsStyle.solidUpBody = False
End If
studyValueConfig.barStyle = BarsStyle

If mContract.specifier.sectype <> SecurityTypes.SecTypeCash And _
    mContract.specifier.sectype <> SecurityTypes.SecTypeIndex _
Then
    Set studyValueConfig = createBarsStudyConfig.StudyValueConfigurations.add("Volume")
    studyValueConfig.chartRegionName = RegionNameVolume
    studyValueConfig.includeInChart = True
    If Not mChartSpec.VolumeStyle Is Nothing Then
        Set VolumeStyle = mChartSpec.VolumeStyle
    Else
        Set VolumeStyle = Chart1.DefaultDataPointStyle
        VolumeStyle.upColor = vbGreen
        VolumeStyle.downColor = vbRed
        VolumeStyle.displayMode = DataPointDisplayModeHistogram
        VolumeStyle.histBarWidth = 0.5
        VolumeStyle.includeInAutoscale = True
        VolumeStyle.lineThickness = 1
    End If
    studyValueConfig.dataPointStyle = VolumeStyle
End If
End Function

Private Sub initialiseChart()
Static notFirstTime As Boolean

Set mManager = CreateChartManager(mTicker.StudyManager, Chart1.controller)

setState ChartStates.ChartStateCreated

Set mController = Chart1.controller

Chart1.SuppressDrawing = True

If Not notFirstTime Then
    Chart1.twipsPerBar = mChartSpec.twipsPerBar
    Chart1.controller.ChartBackGradientFillColors = mChartSpec.ChartBackGradientFillColors
    notFirstTime = True
End If

If Not mChartSpec.XAxisRegionStyle Is Nothing Then Chart1.XAxisRegion.Style = mChartSpec.XAxisRegionStyle
If Not mChartSpec.DefaultYAxisRegionStyle Is Nothing Then Chart1.DefaultYAxisStyle = mChartSpec.DefaultYAxisRegionStyle

If Not mChartSpec.DefaultRegionStyle Is Nothing Then Chart1.DefaultRegionStyle = mChartSpec.DefaultRegionStyle

Set mPriceRegion = mController.AddChartRegion(100, 25, , , RegionNamePrice)

Chart1.SuppressDrawing = False

End Sub

Private Sub loadchart()
Dim volRegionStyle As ChartRegionStyle

Set mContract = mTicker.Contract

Chart1.SuppressDrawing = True

Chart1.barTimePeriod = mChartSpec.Timeframe

Chart1.sessionStartTime = mContract.sessionStartTime
Chart1.sessionEndTime = mContract.sessionEndTime

mPriceRegion.YScaleQuantum = mContract.tickSize
If mChartSpec.MinimumTicksHeight * mContract.tickSize <> 0 Then
    mPriceRegion.MinimumHeight = mChartSpec.MinimumTicksHeight * mContract.tickSize
End If

mPriceRegion.Title.Text = mContract.specifier.localSymbol & _
                " (" & mContract.specifier.exchange & ") " & _
                timeframeCaption
mPriceRegion.Title.Color = vbBlue

If mContract.specifier.sectype <> SecurityTypes.SecTypeCash _
    And mContract.specifier.sectype <> SecurityTypes.SecTypeIndex _
Then
    If Not mChartSpec.VolumeRegionStyle Is Nothing Then
        Set volRegionStyle = mChartSpec.VolumeRegionStyle
    Else
        Set volRegionStyle = Chart1.DefaultRegionStyle
        volRegionStyle.GridlineSpacingY = 0.8
        volRegionStyle.MinimumHeight = 10
        volRegionStyle.IntegerYScale = True
    End If
    
    Set mVolumeRegion = mController.AddChartRegion(20, , volRegionStyle, , RegionNameVolume)
    
    mVolumeRegion.Title.Text = "Volume"
    mVolumeRegion.Title.Color = vbBlue
End If

Set mLoadingText = mPriceRegion.AddText(ChartSkil26.LayerNumbers.LayerHighestUser)
mLoadingText.Text = "Loading historical data"
Dim Font As New stdole.StdFont
Font.size = 18
mLoadingText.Font = Font
mLoadingText.Color = vbBlack
mLoadingText.box = True
mLoadingText.boxFillColor = vbWhite
mLoadingText.boxFillStyle = FillStyles.FillSolid
mLoadingText.position = mPriceRegion.newPoint(50, 50, CoordinateSystems.CoordsRelative, CoordinateSystems.CoordsRelative)
mLoadingText.align = TextAlignModes.AlignBoxCentreCentre
mLoadingText.fixedX = True
mLoadingText.fixedY = True

Chart1.SuppressDrawing = False  ' causes the loading text to appear
Chart1.SuppressDrawing = True

Set mTimeframes = mTicker.Timeframes

If mIsHistoricChart Then
    Set mTimeframe = mTimeframes.addHistorical(mChartSpec.Timeframe, _
                                "", _
                                mChartSpec.InitialNumberOfBars, _
                                mFromTime, _
                                mToTime, _
                                mChartSpec.includeBarsOutsideSession)
Else
    Set mTimeframe = mTimeframes.add(mChartSpec.Timeframe, _
                                "", _
                                mChartSpec.InitialNumberOfBars, _
                                mChartSpec.includeBarsOutsideSession, _
                                IIf(mTicker.replayingTickfile, True, False))
End If

If mTimeframe.historicDataLoaded Then
    mHistDataLoaded = True
    Chart1.SuppressDrawing = False
End If

showStudies

End Sub

Private Sub prepareChart( _
                ByVal pTicker As ticker, _
                ByVal chartSpec As ChartSpecifier, _
                Optional ByVal BarFormatterFactory As BarFormatterFactory)
Set mTicker = pTicker
Set mChartSpec = chartSpec.Clone

Set mBarFormatterFactory = BarFormatterFactory

initialiseChart

setState (ChartStates.ChartStateCreated)

If mTicker.State = TickerStates.TickerStateReady Or _
    mTicker.State = TickerStates.TickerStateRunning _
Then
    loadchart
    setState ChartStates.ChartStateInitialised
End If

End Sub

Private Sub reconfigureDependingStudies(ByVal studyConfigs As StudyConfigurations, ByVal oldBaseStudyConfig As StudyConfiguration, ByVal newBaseStudyConfig As StudyConfiguration)
Dim sc As StudyConfiguration
For Each sc In studyConfigs
    If sc.underlyingStudy Is oldBaseStudyConfig.Study Then
        Dim newSc As StudyConfiguration
        Set newSc = sc.Clone
        newSc.underlyingStudy = newBaseStudyConfig.Study
        If sc.chartRegionName = oldBaseStudyConfig.chartRegionName Then
            newSc.chartRegionName = newBaseStudyConfig.chartRegionName
        End If
        mManager.addStudy newSc
        mManager.startStudy newSc.Study
        reconfigureDependingStudies studyConfigs, sc, newSc
    End If
Next
End Sub

Private Sub setState(ByVal value As ChartStates)
Dim stateEv As StateChangeEvent
mState = value
If mState = ChartStates.ChartStateLoaded Then mLoadingText.Text = ""
stateEv.State = mState
Set stateEv.Source = Me
RaiseEvent StateChange(stateEv)
End Sub

Private Sub showStudies()
Dim tcs() As TaskController
Dim tc As TaskController
Dim tcMaxIndex As Long
Dim i As Long

Set mBarsStudyConfig = createBarsStudyConfig

tcs = mManager.setupStudyValueListeners(mBarsStudyConfig)

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
