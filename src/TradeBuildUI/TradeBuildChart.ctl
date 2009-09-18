VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{74951842-2BEF-4829-A34F-DC7795A37167}#116.0#0"; "ChartSkil2-6.ocx"
Begin VB.UserControl TradeBuildChart 
   Alignable       =   -1  'True
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   ScaleHeight     =   5475
   ScaleWidth      =   7740
   ToolboxBitmap   =   "TradeBuildChart.ctx":0000
   Begin MSComctlLib.ProgressBar LoadingProgressBar 
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
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

Private Const ModuleName                                As String = "TradeBuildChart"

Private Const ConfigSectionBarFormatterFactory          As String = "BarFormatterFactory"
Private Const ConfigSectionChartSpecifier               As String = "ChartSpecifier"
Private Const ConfigSectionStudies                      As String = "Studies"

Private Const ConfigSettingFromTime                     As String = ".FromTime"
Private Const ConfigSettingIsHistoricChart              As String = ".IsHistoricChart"
Private Const ConfigSettingProgId                       As String = "&ProgId"
Private Const ConfigSettingToTime                       As String = ".ToTime"
Private Const ConfigSettingTickerKey                    As String = ".TickerKey"
Private Const ConfigSettingWorkspace                    As String = ".Workspace"

Private Const PropNameHorizontalMouseScrollingAllowed     As String = "HorizontalMouseScrollingAllowed"
Private Const PropNameVerticalMouseScrollingAllowed       As String = "VerticalMouseScrollingAllowed"
Private Const PropNameAutoscroll                        As String = "Autoscrolling"
Private Const PropNameChartBackColor                    As String = "ChartBackColor"
Private Const PropNamePointerDiscColor                  As String = "PointerDiscColor"
Private Const PropNamePointerCrosshairsColor            As String = "PointerCrosshairsColor"
Private Const PropNamePointerStyle                      As String = "PointerStyle"
Private Const PropNameShowHorizontalScrollBar           As String = "HorizontalScrollBarVisible"
Private Const PropNameTwipsPerBar                       As String = "TwipsPerBar"
Private Const PropNameYAxisWidthCm                      As String = "YAxisWidthCm"

Private Const PropDfltHorizontalMouseScrollingAllowed     As Boolean = True
Private Const PropDfltVerticalMouseScrollingAllowed       As Boolean = True
Private Const PropDfltAutoscroll                        As Boolean = True
Private Const PropDfltChartBackColor                    As Long = vbWhite
Private Const PropDfltPointerDiscColor                  As Long = &H89FFFF
Private Const PropDfltPointerCrosshairsColor            As Long = &HC1DFE
Private Const PropDfltPointerStyle                      As Long = PointerStyles.PointerCrosshairs
Private Const PropDfltShowHorizontalScrollBar           As Boolean = True
Private Const PropDfltTwipsPerBar                       As Long = 150
Private Const PropDfltYAxisWidthCm                      As Single = 1.3

'@================================================================================
' Member variables
'@================================================================================

Private mManager                                        As ChartManager

Private WithEvents mTicker                              As Ticker
Attribute mTicker.VB_VarHelpID = -1
Private mTimeframes                                     As Timeframes
Private WithEvents mTimeframe                           As Timeframe
Attribute mTimeframe.VB_VarHelpID = -1

Private mUpdatePerTick                                  As Boolean

Private mState                                          As ChartStates

Private mIsHistoricChart                                As Boolean

Private mChartSpec                                      As ChartSpecifier

Private mFromTime                                       As Date
Private mToTime                                         As Date

Private mContract                                       As Contract

Private mPriceRegion                                    As ChartRegion

Private mVolumeRegion                                   As ChartRegion

Private mPrevWidth                                      As Single
Private mPrevHeight                                     As Single

Private mLoadingText                                    As Text

Private mBarFormatterFactory                            As BarFormatterFactory

Private mConfig                                         As ConfigurationSection
Private mLoadedFromConfig                               As Boolean

Private mTradeBarSeries                                 As BarSeries

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

HorizontalMouseScrollingAllowed = PropDfltHorizontalMouseScrollingAllowed
VerticalMouseScrollingAllowed = PropDfltVerticalMouseScrollingAllowed
Autoscrolling = PropDfltAutoscroll
ChartBackColor = PropDfltChartBackColor
PointerStyle = PropDfltPointerStyle
PointerCrosshairsColor = PropDfltPointerCrosshairsColor
PointerDiscColor = PropDfltPointerDiscColor
HorizontalScrollBarVisible = PropDfltShowHorizontalScrollBar
TwipsPerBar = PropDfltTwipsPerBar
YAxisWidthCm = PropDfltYAxisWidthCm

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

HorizontalMouseScrollingAllowed = PropBag.ReadProperty(PropNameHorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed)
If Err.Number <> 0 Then
    HorizontalMouseScrollingAllowed = PropDfltHorizontalMouseScrollingAllowed
    Err.Clear
End If

VerticalMouseScrollingAllowed = PropBag.ReadProperty(PropNameVerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed)
If Err.Number <> 0 Then
    VerticalMouseScrollingAllowed = PropDfltVerticalMouseScrollingAllowed
    Err.Clear
End If

Autoscrolling = PropBag.ReadProperty(PropNameAutoscroll, PropDfltAutoscroll)
If Err.Number <> 0 Then
    Autoscrolling = PropDfltAutoscroll
    Err.Clear
End If

ChartBackColor = PropBag.ReadProperty(PropNameChartBackColor)
' if no ChartBackColor has been set, we'll just use the ChartSkil default
If Err.Number <> 0 Then Err.Clear

PointerStyle = PropBag.ReadProperty(PropNamePointerStyle, PropDfltPointerStyle)
If Err.Number <> 0 Then
    PointerStyle = PropDfltPointerStyle
    Err.Clear
End If

PointerCrosshairsColor = PropBag.ReadProperty(PropNamePointerCrosshairsColor, PropDfltPointerCrosshairsColor)
If Err.Number <> 0 Then
    PointerCrosshairsColor = PropDfltPointerCrosshairsColor
    Err.Clear
End If

PointerDiscColor = PropBag.ReadProperty(PropNamePointerDiscColor, PropDfltPointerDiscColor)
If Err.Number <> 0 Then
    PointerDiscColor = PropDfltPointerDiscColor
    Err.Clear
End If

HorizontalScrollBarVisible = PropBag.ReadProperty(PropNameShowHorizontalScrollBar, PropDfltShowHorizontalScrollBar)
If Err.Number <> 0 Then
    HorizontalScrollBarVisible = PropDfltShowHorizontalScrollBar
    Err.Clear
End If

TwipsPerBar = PropBag.ReadProperty(PropNameTwipsPerBar, PropDfltTwipsPerBar)
If Err.Number <> 0 Then
    TwipsPerBar = PropDfltTwipsPerBar
    Err.Clear
End If

YAxisWidthCm = PropBag.ReadProperty(PropNameYAxisWidthCm, PropDfltYAxisWidthCm)
If Err.Number <> 0 Then
    YAxisWidthCm = PropDfltYAxisWidthCm
    Err.Clear
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

Private Sub UserControl_Terminate()
gLogger.Log LogLevelDetail, "TradeBuildChart terminated"
Debug.Print "TradeBuildChart terminated"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty PropNameHorizontalMouseScrollingAllowed, HorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed
PropBag.WriteProperty PropNameVerticalMouseScrollingAllowed, VerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed
PropBag.WriteProperty PropNameAutoscroll, Autoscrolling, PropDfltAutoscroll
PropBag.WriteProperty PropNameChartBackColor, ChartBackColor, PropDfltChartBackColor
PropBag.WriteProperty PropNamePointerStyle, PointerStyle, PropDfltPointerStyle
PropBag.WriteProperty PropNamePointerCrosshairsColor, PointerCrosshairsColor, PropDfltPointerCrosshairsColor
PropBag.WriteProperty PropNamePointerDiscColor, PointerDiscColor, PropDfltPointerDiscColor
PropBag.WriteProperty PropNameShowHorizontalScrollBar, HorizontalScrollBarVisible, PropDfltShowHorizontalScrollBar
PropBag.WriteProperty PropNameTwipsPerBar, TwipsPerBar, PropDfltTwipsPerBar
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
' mTicker Event Handlers
'@================================================================================

Private Sub mTicker_StateChange(ev As StateChangeEvent)
If ev.State = TickerStates.TickerStateReady Then
    ' this means that the Ticker object has retrieved the contract info, so we can
    ' now start the chart
    loadchart
    If mLoadedFromConfig Then
        loadStudiesFromConfig
    Else
        showStudies createBarsStudyConfig
    End If
End If
End Sub

'@================================================================================
' mTimeframe Event Handlers
'@================================================================================

Private Sub mTimeframe_BarsLoaded()
LoadingProgressBar.Visible = False
mLoadingText.Text = ""
Chart1.EnableDrawing

setState ChartStates.ChartStateLoaded
End Sub

Private Sub mTimeframe_BarLoadProgress(ByVal barsRetrieved As Long, ByVal percentComplete As Single)
If Not LoadingProgressBar.Visible Then
    LoadingProgressBar.Top = UserControl.Height - LoadingProgressBar.Height
    LoadingProgressBar.Width = UserControl.Width
    LoadingProgressBar.Left = 0
    LoadingProgressBar.Visible = True
    
    mLoadingText.Text = "Loading historical data"
    setState ChartStateLoading
    Chart1.EnableDrawing
    Chart1.DisableDrawing
End If
LoadingProgressBar.value = percentComplete
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Autoscrolling( _
                ByVal value As Boolean)
Chart1.Autoscrolling = value
End Property

Public Property Get Autoscrolling() As Boolean
Autoscrolling = Chart1.Autoscrolling
End Property

Public Property Get BaseChartController() As ChartController
Set BaseChartController = Chart1.Controller
End Property

Public Property Get ChartBackColor() As OLE_COLOR
ChartBackColor = Chart1.ChartBackColor
End Property

Public Property Let ChartBackColor(ByVal val As OLE_COLOR)
Chart1.ChartBackColor = val
End Property

Public Property Get ChartManager() As ChartManager
Set ChartManager = mManager
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings

If Not mManager Is Nothing Then mManager.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudies)
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

Public Property Let HorizontalMouseScrollingAllowed( _
                ByVal value As Boolean)
Chart1.HorizontalMouseScrollingAllowed = value
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
HorizontalMouseScrollingAllowed = Chart1.HorizontalMouseScrollingAllowed
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
HorizontalScrollBarVisible = Chart1.HorizontalScrollBarVisible
End Property

Public Property Let HorizontalScrollBarVisible(ByVal val As Boolean)
Chart1.HorizontalScrollBarVisible = val
End Property

Public Property Get InitialNumberOfBars() As Long
Attribute InitialNumberOfBars.VB_ProcData.VB_Invoke_Property = ";Behavior"
InitialNumberOfBars = mChartSpec.InitialNumberOfBars
End Property

Public Property Get LoadingText() As Text
Set LoadingText = mLoadingText
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

Public Property Get PriceRegion() As ChartRegion
Set PriceRegion = mPriceRegion
End Property

Public Property Get RegionNames() As String()
RegionNames = mManager.RegionNames
End Property

Public Property Get State() As ChartStates
State = mState
End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
End Property

Public Property Get TimeframeCaption() As String
TimeframeCaption = mChartSpec.Timeframe.toString
End Property

Public Property Get TimeframeShortCaption() As String
TimeframeShortCaption = mChartSpec.Timeframe.ToShortString
End Property

Public Property Get Timeframe() As Timeframe
Set Timeframe = mTimeframe
End Property

Public Property Get TimePeriod() As TimePeriod
Set TimePeriod = mChartSpec.Timeframe
End Property

Friend Property Get TradeBarSeries() As BarSeries
Set TradeBarSeries = mTradeBarSeries
End Property

Public Property Get TwipsPerBar() As Long
TwipsPerBar = Chart1.TwipsPerBar
End Property

Public Property Let TwipsPerBar(ByVal val As Long)
Chart1.TwipsPerBar = val
End Property

Public Property Let UpdatePerTick(ByVal value As Boolean)
Attribute UpdatePerTick.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
mUpdatePerTick = value
End Property

Public Property Let VerticalMouseScrollingAllowed( _
                ByVal value As Boolean)
Chart1.VerticalMouseScrollingAllowed = value
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
VerticalMouseScrollingAllowed = Chart1.VerticalMouseScrollingAllowed
End Property

Public Property Get VolumeRegion() As ChartRegion
Set VolumeRegion = mVolumeRegion
End Property

Public Property Get VolumeRegionStyle() As ChartRegionStyle
Set VolumeRegionStyle = mChartSpec.VolumeRegionStyle
End Property

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
Dim baseStudyConfig As StudyConfiguration

Dim failpoint As Long
On Error GoTo Err

If State <> ChartStateLoaded Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                            ProjectName & "." & ModuleName & ":" & "ChangeTimeframe", _
                                            "Can't change timeframe until chart is loaded"

mLoadedFromConfig = False

Set baseStudyConfig = mManager.BaseStudyConfiguration

Set mPriceRegion = Nothing
Set mVolumeRegion = Nothing
Set mTradeBarSeries = Nothing

mManager.ClearChart

setState ChartStateBlank

mChartSpec.Timeframe = Timeframe

createTimeframe
baseStudyConfig.Study = mTimeframe.tradeStudy
baseStudyConfig.StudyValueConfigurations.item("Bar").SetBarFormatterFactory mBarFormatterFactory, mTimeframe.tradeBars
Dim lStudy As Study
Set lStudy = mTimeframe.tradeStudy
baseStudyConfig.Parameters = lStudy.Parameters

initialiseChart

loadchart
showStudies baseStudyConfig

RaiseEvent TimeframeChange

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ChangeTimeframe" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub DisableDrawing()
Chart1.DisableDrawing
End Sub

Public Sub EnableDrawing()
Chart1.EnableDrawing
End Sub

Public Sub Finish()
Dim failpoint As Long
On Error GoTo Err

' update the number of bars in case this chart is reloaded from the config
If Not mChartSpec Is Nothing Then
    If mChartSpec.InitialNumberOfBars < Chart1.Periods.Count Then mChartSpec.InitialNumberOfBars = Chart1.Periods.Count
End If

If Not mManager Is Nothing Then mManager.Finish

Set mManager = Nothing

Set mTimeframes = Nothing
Set mTimeframe = Nothing

Set mContract = Nothing

Set mPriceRegion = Nothing
Set mVolumeRegion = Nothing
Set mTradeBarSeries = Nothing

Set mLoadingText = Nothing

mLoadedFromConfig = False

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Finish" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Dim cs As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

Set mConfig = config
mLoadedFromConfig = True

Set mTicker = TradeBuildAPI.WorkSpaces(mConfig.GetSetting(ConfigSettingWorkspace)).Tickers(mConfig.GetSetting(ConfigSettingTickerKey))
Set mChartSpec = LoadChartSpecifierFromConfig(mConfig.GetConfigurationSection(ConfigSectionChartSpecifier))
mIsHistoricChart = CBool(mConfig.GetSetting(ConfigSettingIsHistoricChart, "False"))
mFromTime = CDate(mConfig.GetSetting(ConfigSettingFromTime, "0"))
mToTime = CDate(mConfig.GetSetting(ConfigSettingToTime, "0"))

Set cs = mConfig.GetConfigurationSection(ConfigSectionBarFormatterFactory)
If Not cs Is Nothing Then
    Set mBarFormatterFactory = CreateObject(cs.GetSetting(ConfigSettingProgId))
    mBarFormatterFactory.LoadFromConfig cs
End If

prepareChart

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "LoadFromConfig" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Sub RemoveFromConfig()
Dim failpoint As Long
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove
Set mConfig = Nothing

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "RemoveFromConfig" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub ScrollToTime(ByVal pTime As Date)
Dim failpoint As Long
On Error GoTo Err

mManager.ScrollToTime pTime

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ScrollToTime" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub showChart( _
                ByVal pTicker As Ticker, _
                ByVal chartSpec As ChartSpecifier, _
                Optional ByVal BarFormatterFactory As BarFormatterFactory)

Dim failpoint As Long
On Error GoTo Err

Select Case chartSpec.Timeframe.Units
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
Set mChartSpec = chartSpec.Clone
Set mBarFormatterFactory = BarFormatterFactory

storeSettings

prepareChart

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "showChart" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Public Sub showHistoricChart( _
                ByVal pTicker As Ticker, _
                ByVal chartSpec As ChartSpecifier, _
                ByVal fromTime As Date, _
                ByVal toTime As Date, _
                Optional ByVal BarFormatterFactory As BarFormatterFactory)

Dim failpoint As Long
On Error GoTo Err

Select Case chartSpec.Timeframe.Units
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

Set mTicker = pTicker
Set mChartSpec = chartSpec.Clone
Set mBarFormatterFactory = BarFormatterFactory
mIsHistoricChart = True
mFromTime = fromTime
mToTime = toTime

storeSettings

prepareChart

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "showHistoricChart" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

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
Dim BarsStyle As BarStyle
Dim VolumeStyle As DataPointStyle

Dim studyConfig As StudyConfiguration

Set studyConfig = New StudyConfiguration

studyConfig.UnderlyingStudy = mTicker.InputStudy

Set lStudy = mTimeframe.tradeStudy
studyConfig.Study = lStudy
Set studyDef = lStudy.StudyDefinition

studyConfig.ChartRegionName = ChartRegionNamePrice

inputValueNames(0) = mTicker.InputNameTrade
inputValueNames(1) = mTicker.InputNameVolume
inputValueNames(2) = mTicker.InputNameTickVolume
inputValueNames(3) = mTicker.InputNameOpenInterest
studyConfig.inputValueNames = inputValueNames
studyConfig.name = studyDef.name
params.SetParameterValue "Bar length", mChartSpec.Timeframe.length
params.SetParameterValue "Time units", TimePeriodUnitsToString(mChartSpec.Timeframe.Units)
studyConfig.Parameters = params

Set studyValueConfig = studyConfig.StudyValueConfigurations.Add("Bar")
studyValueConfig.ChartRegionName = ChartRegionNamePrice
studyValueConfig.IncludeInChart = True
studyValueConfig.Layer = 200
studyValueConfig.SetBarFormatterFactory mBarFormatterFactory, mTimeframe.tradeBars

If Not mChartSpec.BarsStyle Is Nothing Then
    Set BarsStyle = mChartSpec.BarsStyle
Else
    Set BarsStyle = New BarStyle
    BarsStyle.DisplayMode = BarDisplayModes.BarDisplayModeCandlestick
    BarsStyle.OutlineThickness = 1
    BarsStyle.TailThickness = 1
    BarsStyle.UpColor = &HA0A0A0
    BarsStyle.SolidUpBody = False
End If
studyValueConfig.BarStyle = BarsStyle

If mContract.specifier.secType <> SecurityTypes.SecTypeCash And _
    mContract.specifier.secType <> SecurityTypes.SecTypeIndex _
Then
    Set studyValueConfig = studyConfig.StudyValueConfigurations.Add("Volume")
    studyValueConfig.ChartRegionName = ChartRegionNameVolume
    studyValueConfig.IncludeInChart = True
    If Not mChartSpec.VolumeStyle Is Nothing Then
        Set VolumeStyle = mChartSpec.VolumeStyle
    Else
        Set VolumeStyle = New DataPointStyle
        VolumeStyle.UpColor = vbGreen
        VolumeStyle.DownColor = vbRed
        VolumeStyle.DisplayMode = DataPointDisplayModeHistogram
        VolumeStyle.HistogramBarWidth = 0.5
        VolumeStyle.IncludeInAutoscale = True
        VolumeStyle.LineThickness = 1
    End If
    studyValueConfig.DataPointStyle = VolumeStyle
End If

Set createBarsStudyConfig = studyConfig
End Function

Private Sub createTimeframe()
Set mTimeframes = mTicker.Timeframes

If mIsHistoricChart Then
    Set mTimeframe = mTimeframes.AddHistorical(mChartSpec.Timeframe, _
                                "", _
                                mChartSpec.InitialNumberOfBars, _
                                mFromTime, _
                                mToTime, _
                                mChartSpec.IncludeBarsOutsideSession)
Else
    Set mTimeframe = mTimeframes.Add(mChartSpec.Timeframe, _
                                "", _
                                mChartSpec.InitialNumberOfBars, _
                                mChartSpec.IncludeBarsOutsideSession, _
                                IIf(mTicker.ReplayingTickfile, True, False))
End If

End Sub

Private Sub initialiseChart()
Static notFirstTime As Boolean

Chart1.DisableDrawing

If Not notFirstTime Then
    Set mManager = CreateChartManager(mTicker.StudyManager, Chart1.Controller)
    If Not mConfig Is Nothing Then mManager.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudies)
    
    Chart1.Regions.DefaultDataRegionStyle = mChartSpec.DefaultRegionStyle
    Chart1.Regions.DefaultYAxisRegionStyle = mChartSpec.DefaultYAxisRegionStyle
    Chart1.TwipsPerBar = mChartSpec.TwipsPerBar
    Chart1.ChartBackColor = mChartSpec.ChartBackColor
    notFirstTime = True
End If

If Not mChartSpec.XAxisRegionStyle Is Nothing Then Chart1.XAxisRegion.Style = mChartSpec.XAxisRegionStyle

Set mPriceRegion = Chart1.Regions.Add(100, _
                                        25, _
                                        , _
                                        , _
                                        ChartRegionNamePrice)
setLoadingText
Chart1.EnableDrawing

setState ChartStates.ChartStateCreated
End Sub

Private Sub loadchart()
Dim volRegionStyle As ChartRegionStyle

Set mContract = mTicker.Contract

Chart1.DisableDrawing

Chart1.BarTimePeriod = mChartSpec.Timeframe

Chart1.SessionStartTime = mContract.SessionStartTime
Chart1.SessionEndTime = mContract.SessionEndTime

mPriceRegion.YScaleQuantum = mContract.tickSize
If mChartSpec.MinimumTicksHeight * mContract.tickSize <> 0 Then
    mPriceRegion.MinimumHeight = mChartSpec.MinimumTicksHeight * mContract.tickSize
End If

mPriceRegion.Title.Text = mContract.specifier.localSymbol & _
                " (" & mContract.specifier.exchange & ") " & _
                TimeframeCaption
mPriceRegion.Title.Color = vbBlue

If mContract.specifier.secType <> SecurityTypes.SecTypeCash _
    And mContract.specifier.secType <> SecurityTypes.SecTypeIndex _
Then
    If Not mChartSpec.VolumeRegionStyle Is Nothing Then
        Set volRegionStyle = mChartSpec.VolumeRegionStyle
    Else
        Set volRegionStyle = mChartSpec.DefaultRegionStyle
        volRegionStyle.GridlineSpacingY = 0.8
        volRegionStyle.MinimumHeight = 10
        volRegionStyle.IntegerYScale = True
    End If
    
    On Error Resume Next
    Set mVolumeRegion = Chart1.Regions.item(ChartRegionNameVolume)
    On Error GoTo 0
    
    If mVolumeRegion Is Nothing Then
        Set mVolumeRegion = Chart1.Regions.Add(20 _
                                                , _
                                                , _
                                                volRegionStyle, _
                                                , _
                                                ChartRegionNameVolume)
    
    End If
    
    mVolumeRegion.Title.Text = "Volume"
    mVolumeRegion.Title.Color = vbBlue
End If

If Not mTimeframe.historicDataLoaded Then
    mLoadingText.Text = "Fetching historical data"
    setState ChartStates.ChartStateInitialised
    Chart1.EnableDrawing    ' causes the loading text to appear
    Chart1.DisableDrawing
Else
    Chart1.EnableDrawing
    setState ChartStates.ChartStateInitialised
    setState ChartStates.ChartStateLoaded
End If

End Sub

Private Sub loadStudiesFromConfig()
mManager.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionStudies), mTimeframe.tradeStudy
setTradeBarSeries
End Sub

Private Sub prepareChart()

createTimeframe
initialiseChart

If mTicker.State = TickerStates.TickerStateReady Or _
    mTicker.State = TickerStates.TickerStateRunning _
Then
    loadchart
    If mLoadedFromConfig Then
        loadStudiesFromConfig
    Else
        showStudies createBarsStudyConfig
    End If
End If

End Sub

Private Sub setLoadingText()
Set mLoadingText = mPriceRegion.AddText(, ChartSkil26.LayerNumbers.LayerHighestUser)
Dim Font As New stdole.StdFont
Font.size = 18
mLoadingText.Font = Font
mLoadingText.Color = vbBlack
mLoadingText.Box = True
mLoadingText.BoxFillColor = vbWhite
mLoadingText.BoxFillStyle = FillStyles.FillSolid
mLoadingText.Position = mPriceRegion.NewPoint(50, 50, CoordinateSystems.CoordsRelative, CoordinateSystems.CoordsRelative)
mLoadingText.align = TextAlignModes.AlignBoxCentreCentre
mLoadingText.FixedX = True
mLoadingText.FixedY = True
End Sub

Private Sub setState(ByVal value As ChartStates)
Dim stateEv As StateChangeEvent
mState = value
stateEv.State = mState
Set stateEv.Source = Me
RaiseEvent StateChange(stateEv)
End Sub

Private Sub setTradeBarSeries()
Set mTradeBarSeries = mManager.BaseStudyConfiguration.ValueSeries("Bar")
End Sub

Private Sub showStudies( _
                ByVal studyConfig As StudyConfiguration)
mManager.BaseStudyConfiguration = studyConfig
setTradeBarSeries
End Sub

Private Sub storeSettings()
Dim cs As ConfigurationSection

If mConfig Is Nothing Then Exit Sub
    
If mTicker Is Nothing Then Exit Sub

mConfig.SetSetting ConfigSettingWorkspace, mTicker.Workspace.name
mConfig.SetSetting ConfigSettingTickerKey, mTicker.Key
mChartSpec.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionChartSpecifier)
mConfig.SetSetting ConfigSettingIsHistoricChart, CStr(mIsHistoricChart)
mConfig.SetSetting ConfigSettingFromTime, CStr(CDbl(mFromTime))
mConfig.SetSetting ConfigSettingToTime, CStr(CDbl(mToTime))

If Not mBarFormatterFactory Is Nothing Then
    Set cs = mConfig.AddConfigurationSection(ConfigSectionBarFormatterFactory)
    cs.SetSetting ConfigSettingProgId, GetProgIdFromObject(mBarFormatterFactory)
    mBarFormatterFactory.ConfigurationSection = cs
End If
End Sub
