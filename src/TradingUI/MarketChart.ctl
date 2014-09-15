VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5EF6A0B6-9E1F-426C-B84A-601F4CBF70C4}#214.0#0"; "ChartSkil27.ocx"
Begin VB.UserControl MarketChart 
   Alignable       =   -1  'True
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   ScaleHeight     =   5475
   ScaleWidth      =   7740
   ToolboxBitmap   =   "MarketChart.ctx":0000
   Begin ChartSkil27.Chart Chart1 
      Align           =   1  'Align Top
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   2990
   End
   Begin MSComctlLib.ProgressBar LoadingProgressBar 
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "MarketChart"
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
Attribute KeyDown.VB_UserMemId = -602

Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603

Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604

Event StateChange(ev As StateChangeEventData)

Event TimePeriodChange()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                As String = "MarketChart"

Private Const ConfigSectionChartControl                 As String = "ChartControl"
Private Const ConfigSectionChartSpecifier               As String = "ChartSpecifier"
Private Const ConfigSectionStudies                      As String = "Studies"

Private Const ConfigSettingBarFormatterFactoryName      As String = "&BarFormatterFactoryName"
Private Const ConfigSettingBarFormatterLibraryName      As String = "&BarFormatterLibraryName"
Private Const ConfigSettingIsHistoricChart              As String = "&IsHistoricChart"
Private Const ConfigSettingTimePeriod                   As String = "&TimePeriod"
Private Const ConfigSettingDataSourceKey                As String = "&DataSourceKey"

Private Const PropNameHorizontalMouseScrollingAllowed   As String = "HorizontalMouseScrollingAllowed"
Private Const PropNameVerticalMouseScrollingAllowed     As String = "VerticalMouseScrollingAllowed"
Private Const PropNameAutoscroll                        As String = "Autoscrolling"
Private Const PropNameChartBackColor                    As String = "ChartBackColor"
Private Const PropNamePointerDiscColor                  As String = "PointerDiscColor"
Private Const PropNamePointerCrosshairsColor            As String = "PointerCrosshairsColor"
Private Const PropNamePointerStyle                      As String = "PointerStyle"
Private Const PropNameShowHorizontalScrollBar           As String = "HorizontalScrollBarVisible"
Private Const PropNamePeriodWidth                       As String = "PeriodWidth"
Private Const PropNameYAxisWidthCm                      As String = "YAxisWidthCm"

Private Const PropDfltHorizontalMouseScrollingAllowed   As Boolean = True
Private Const PropDfltVerticalMouseScrollingAllowed     As Boolean = True
Private Const PropDfltAutoscroll                        As Boolean = True
Private Const PropDfltChartBackColor                    As Long = vbWhite
Private Const PropDfltPointerDiscColor                  As Long = &H89FFFF
Private Const PropDfltPointerCrosshairsColor            As Long = &HC1DFE
Private Const PropDfltPointerStyle                      As Long = PointerStyles.PointerCrosshairs
Private Const PropDfltShowHorizontalScrollBar           As Boolean = True
Private Const PropDfltPeriodWidth                       As Long = 100
Private Const PropDfltYAxisWidthCm                      As Single = 1.8

Private Const StudyValueConfigNameBar                   As String = "Bar"
Private Const StudyValueConfigNameVolume                As String = "Volume"

'@================================================================================
' Member variables
'@================================================================================

Private mManager                                        As ChartManager

Private mTimeframes                                     As Timeframes
Private WithEvents mTimeframe                           As Timeframe
Attribute mTimeframe.VB_VarHelpID = -1

Private mTimePeriod                                     As TimePeriod

Private mUpdatePerTick                                  As Boolean

Private mState                                          As ChartStates

Private mIsHistoricChart                                As Boolean

Private mChartSpec                                      As ChartSpecifier
Private mChartStyle                                     As ChartStyle

Private mContract                                       As Contract

Private mPriceRegion                                    As ChartRegion

Private mVolumeRegion                                   As ChartRegion

Private mPrevWidth                                      As Single
Private mPrevHeight                                     As Single

Private mLoadingText                                    As Text

Private mStudyManager                                   As StudyManager
Private mBarFormatterLibManager                         As BarFormatterLibManager

Private mBarFormatterFactoryName                        As String
Private mBarFormatterLibraryName                        As String

Private mConfig                                         As ConfigurationSection
Private mLoadedFromConfig                               As Boolean

Private mDeferStart                                     As Boolean

Private mMinimumTicksHeight                             As Long

'' this is a temporary style that is used initially to apply property bag settings.
'' This prevents the property bag settings from overriding any style that is later
'' applied.
'Private mInitialStyle                                   As ChartStyle

Private mInitialised                                    As Boolean

Private mExcludeCurrentBar                              As Boolean

Private WithEvents mFutureWaiter                        As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTitle                                          As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()

mPrevWidth = UserControl.Width
mPrevHeight = UserControl.Height

mUpdatePerTick = True

'Set mInitialStyle = ChartStylesManager.Add(GenerateGUIDString, ChartStylesManager.DefaultStyle, pTemporary:=True)

End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next

'mInitialStyle.HorizontalMouseScrollingAllowed = PropDfltHorizontalMouseScrollingAllowed
'mInitialStyle.VerticalMouseScrollingAllowed = PropDfltVerticalMouseScrollingAllowed
'mInitialStyle.Autoscrolling = PropDfltAutoscroll
'mInitialStyle.ChartBackColor = PropDfltChartBackColor
'PointerStyle = PropDfltPointerStyle
'PointerCrosshairsColor = PropDfltPointerCrosshairsColor
'PointerDiscColor = PropDfltPointerDiscColor
'mInitialStyle.HorizontalScrollBarVisible = PropDfltShowHorizontalScrollBar
'mInitialStyle.PeriodWidth = PropDfltPeriodWidth
'mInitialStyle.YAxisWidthCm = PropDfltYAxisWidthCm

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next

'mInitialStyle.HorizontalMouseScrollingAllowed = PropBag.ReadProperty(PropNameHorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed)
'mInitialStyle.VerticalMouseScrollingAllowed = PropBag.ReadProperty(PropNameVerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed)
'mInitialStyle.Autoscrolling = PropBag.ReadProperty(PropNameAutoscroll, PropDfltAutoscroll)
'mInitialStyle.ChartBackColor = PropBag.ReadProperty(PropNameChartBackColor)
'' if no ChartBackColor has been set, we'll just use the ChartSkil default
'
'PointerStyle = PropBag.ReadProperty(PropNamePointerStyle, PropDfltPointerStyle)
'PointerCrosshairsColor = PropBag.ReadProperty(PropNamePointerCrosshairsColor, PropDfltPointerCrosshairsColor)
'PointerDiscColor = PropBag.ReadProperty(PropNamePointerDiscColor, PropDfltPointerDiscColor)
'mInitialStyle.HorizontalScrollBarVisible = PropBag.ReadProperty(PropNameShowHorizontalScrollBar, PropDfltShowHorizontalScrollBar)
'mInitialStyle.PeriodWidth = PropBag.ReadProperty(PropNamePeriodWidth, PropDfltPeriodWidth)
'mInitialStyle.YAxisWidthCm = PropBag.ReadProperty(PropNameYAxisWidthCm, PropDfltYAxisWidthCm)
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
Const ProcName As String = "UserControl_Terminate"
gLogger.Log "MarketChart terminated", ProcName, ModuleName, LogLevelDetail
Debug.Print "MarketChart terminated"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
'PropBag.WriteProperty PropNameHorizontalMouseScrollingAllowed, mInitialStyle.HorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed
'PropBag.WriteProperty PropNameVerticalMouseScrollingAllowed, mInitialStyle.VerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed
'PropBag.WriteProperty PropNameAutoscroll, mInitialStyle.Autoscrolling, PropDfltAutoscroll
'PropBag.WriteProperty PropNameChartBackColor, mInitialStyle.ChartBackColor, PropDfltChartBackColor
'PropBag.WriteProperty PropNamePointerStyle, PointerStyle, PropDfltPointerStyle
'PropBag.WriteProperty PropNamePointerCrosshairsColor, PointerCrosshairsColor, PropDfltPointerCrosshairsColor
'PropBag.WriteProperty PropNamePointerDiscColor, PointerDiscColor, PropDfltPointerDiscColor
'PropBag.WriteProperty PropNameShowHorizontalScrollBar, mInitialStyle.HorizontalScrollBarVisible, PropDfltShowHorizontalScrollBar
'PropBag.WriteProperty PropNamePeriodWidth, mInitialStyle.PeriodWidth, PropDfltPeriodWidth
'PropBag.WriteProperty PropNameYAxisWidthCm, mInitialStyle.YAxisWidthCm, PropDfltYAxisWidthCm
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
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsAvailable Then
    ' this means that the contract info is available, so we can
    ' now start the chart

    Set mContract = mTimeframes.ContractFuture.value

    If mDeferStart Then Exit Sub

    setupStudies
    loadchart
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTimeframe Event Handlers
'@================================================================================

Private Sub mTimeframe_BarsLoaded()
Const ProcName As String = "mTimeframe_BarsLoaded"
On Error GoTo Err

LoadingProgressBar.Visible = False
mLoadingText.Text = ""
Chart1.EnableDrawing

setState ChartStates.ChartStateLoaded

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mTimeframe_BarLoadProgress(ByVal barsRetrieved As Long, ByVal percentComplete As Single)
Const ProcName As String = "mTimeframe_BarLoadProgress"
On Error GoTo Err

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

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Autoscrolling( _
                ByVal value As Boolean)
Const ProcName As String = "Autoscrolling"
On Error GoTo Err

Chart1.Autoscrolling = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Autoscrolling() As Boolean
Attribute Autoscrolling.VB_MemberFlags = "400"
Const ProcName As String = "Autoscrolling"
On Error GoTo Err

Autoscrolling = Chart1.Autoscrolling

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BaseChartController() As ChartController
Attribute BaseChartController.VB_MemberFlags = "400"
Const ProcName As String = "BaseChartController"
On Error GoTo Err

Set BaseChartController = Chart1.Controller

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Attribute ChartBackColor.VB_MemberFlags = "400"
Const ProcName As String = "ChartBackColor"
On Error GoTo Err

ChartBackColor = Chart1.ChartBackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ChartBackColor(ByVal val As OLE_COLOR)
Const ProcName As String = "ChartBackColor"
On Error GoTo Err

Chart1.ChartBackColor = val

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartManager() As ChartManager
Attribute ChartManager.VB_MemberFlags = "400"
Const ProcName As String = "ChartManager"
On Error GoTo Err

Set ChartManager = mManager

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
Attribute ConfigurationSection.VB_MemberFlags = "400"
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If mConfig Is value Then Exit Property
If Not mConfig Is Nothing Then mConfig.Remove
Set mConfig = Nothing
If value Is Nothing Then Exit Property

Set mConfig = value

gLogger.Log "Chart added to config at: " & mConfig.Path, ProcName, ModuleName

storeSettings

Chart1.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionChartControl)
If Not mManager Is Nothing Then mManager.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudies)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Const ProcName As String = "Enabled"
On Error GoTo Err

Enabled = UserControl.Enabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled = value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HorizontalMouseScrollingAllowed( _
                ByVal value As Boolean)
Const ProcName As String = "HorizontalMouseScrollingAllowed"
On Error GoTo Err

Chart1.HorizontalMouseScrollingAllowed = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
Attribute HorizontalMouseScrollingAllowed.VB_MemberFlags = "400"
Const ProcName As String = "HorizontalMouseScrollingAllowed"
On Error GoTo Err

HorizontalMouseScrollingAllowed = Chart1.HorizontalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
Attribute HorizontalScrollBarVisible.VB_MemberFlags = "400"
Const ProcName As String = "HorizontalScrollBarVisible"
On Error GoTo Err

HorizontalScrollBarVisible = Chart1.HorizontalScrollBarVisible

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HorizontalScrollBarVisible(ByVal val As Boolean)
Const ProcName As String = "HorizontalScrollBarVisible"
On Error GoTo Err

Chart1.HorizontalScrollBarVisible = val

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InitialNumberOfBars() As Long
Attribute InitialNumberOfBars.VB_MemberFlags = "400"
Const ProcName As String = "InitialNumberOfBars"
On Error GoTo Err

InitialNumberOfBars = mChartSpec.InitialNumberOfBars

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LoadingText() As Text
Const ProcName As String = "LoadingText"
On Error GoTo Err

Set LoadingText = mLoadingText

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let MinimumTicksHeight(ByVal value As Double)
Const ProcName As String = "MinimumTicksHeight"
On Error GoTo Err

mMinimumTicksHeight = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MinimumTicksHeight() As Double
Const ProcName As String = "MinimumTicksHeight"
On Error GoTo Err

MinimumTicksHeight = mMinimumTicksHeight

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PeriodWidth() As Long
Const ProcName As String = "PeriodWidth"
On Error GoTo Err

PeriodWidth = Chart1.PeriodWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PeriodWidth(ByVal value As Long)
Const ProcName As String = "PeriodWidth"
On Error GoTo Err

Chart1.PeriodWidth = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerCrosshairsColor() As OLE_COLOR
Attribute PointerCrosshairsColor.VB_MemberFlags = "400"
Const ProcName As String = "PointerCrosshairsColor"
On Error GoTo Err

PointerCrosshairsColor = Chart1.PointerCrosshairsColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PointerCrosshairsColor(ByVal value As OLE_COLOR)
Const ProcName As String = "PointerCrosshairsColor"
On Error GoTo Err

Chart1.PointerCrosshairsColor = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerDiscColor() As OLE_COLOR
Attribute PointerDiscColor.VB_MemberFlags = "400"
Const ProcName As String = "PointerDiscColor"
On Error GoTo Err

PointerDiscColor = Chart1.PointerDiscColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PointerDiscColor(ByVal value As OLE_COLOR)
Const ProcName As String = "PointerDiscColor"
On Error GoTo Err

Chart1.PointerDiscColor = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PointerStyle() As PointerStyles
Attribute PointerStyle.VB_MemberFlags = "400"
Const ProcName As String = "PointerStyle"
On Error GoTo Err

PointerStyle = Chart1.PointerStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PointerStyle(ByVal value As PointerStyles)
Const ProcName As String = "PointerStyle"
On Error GoTo Err

Chart1.PointerStyle = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PriceRegion() As ChartRegion
Attribute PriceRegion.VB_MemberFlags = "400"
Const ProcName As String = "PriceRegion"
On Error GoTo Err

Set PriceRegion = mPriceRegion

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RegionNames() As String()
Attribute RegionNames.VB_MemberFlags = "400"
Const ProcName As String = "RegionNames"
On Error GoTo Err

RegionNames = mManager.RegionNames

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get State() As ChartStates
Attribute State.VB_MemberFlags = "400"
Const ProcName As String = "State"
On Error GoTo Err

State = mState

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'Public Property Get DataSource() As IMarketDataSource
'Const ProcName As String = "DataSource"
'On Error GoTo Err
'
'Set DataSource = mDataSource
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

Public Property Get TimeframeCaption() As String
Attribute TimeframeCaption.VB_MemberFlags = "400"
Const ProcName As String = "TimeframeCaption"
On Error GoTo Err

TimeframeCaption = mTimePeriod.ToString

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimeframeShortCaption() As String
Attribute TimeframeShortCaption.VB_MemberFlags = "400"
Const ProcName As String = "TimeframeShortCaption"
On Error GoTo Err

TimeframeShortCaption = mTimePeriod.ToShortString

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Timeframe() As Timeframe
Attribute Timeframe.VB_MemberFlags = "400"
Const ProcName As String = "Timeframe"
On Error GoTo Err

Set Timeframe = mTimeframe

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimePeriod() As TimePeriod
Attribute TimePeriod.VB_MemberFlags = "400"
Const ProcName As String = "TimePeriod"
On Error GoTo Err

Set TimePeriod = mTimePeriod

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get TradeBarSeries() As BarSeries
Const ProcName As String = "TradeBarSeries"
On Error GoTo Err

Set TradeBarSeries = mManager.BaseStudyConfiguration.ValueSeries(StudyValueConfigNameBar)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let UpdatePerTick(ByVal value As Boolean)
Const ProcName As String = "UpdatePerTick"
On Error GoTo Err

mUpdatePerTick = value
If Not mManager Is Nothing Then mManager.UpdatePerTick = mUpdatePerTick

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let VerticalMouseScrollingAllowed( _
                ByVal value As Boolean)
Const ProcName As String = "VerticalMouseScrollingAllowed"
On Error GoTo Err

Chart1.VerticalMouseScrollingAllowed = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
Attribute VerticalMouseScrollingAllowed.VB_MemberFlags = "400"
Const ProcName As String = "VerticalMouseScrollingAllowed"
On Error GoTo Err

VerticalMouseScrollingAllowed = Chart1.VerticalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VolumeRegion() As ChartRegion
Attribute VolumeRegion.VB_MemberFlags = "400"
Const ProcName As String = "VolumeRegion"
On Error GoTo Err

Set VolumeRegion = mVolumeRegion

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get YAxisWidthCm() As Single
Attribute YAxisWidthCm.VB_MemberFlags = "400"
Const ProcName As String = "YAxisWidthCm"
On Error GoTo Err

YAxisWidthCm = Chart1.YAxisWidthCm

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
Const ProcName As String = "YAxisWidthCm"
On Error GoTo Err

Chart1.YAxisWidthCm = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ChangeTimePeriod(ByVal pNewTimePeriod As TimePeriod)
Const ProcName As String = "ChangeTimePeriod"
On Error GoTo Err

Assert State = ChartStateLoaded, "Can't change timeframe until chart is loaded"

gLogger.Log "Changing timeframe to", ProcName, ModuleName, , pNewTimePeriod.ToString

mLoadedFromConfig = False

Dim baseStudyConfig As StudyConfiguration
Set baseStudyConfig = mManager.BaseStudyConfiguration

Set mPriceRegion = Nothing
Set mVolumeRegion = Nothing

mManager.ClearChart

setState ChartStateBlank

Set mTimePeriod = pNewTimePeriod
storeSettings

createTimeframe

baseStudyConfig.Study = mTimeframe.BarStudy
baseStudyConfig.StudyValueConfigurations.Item(StudyValueConfigNameBar).BarFormatterFactoryName = mBarFormatterFactoryName
baseStudyConfig.StudyValueConfigurations.Item(StudyValueConfigNameBar).BarFormatterLibraryName = mBarFormatterLibraryName

Dim lStudy As IStudy
Set lStudy = mTimeframe.BarStudy
baseStudyConfig.Parameters = lStudy.Parameters

initialiseChart
mManager.BaseStudyConfiguration = baseStudyConfig

loadchart

RaiseEvent TimePeriodChange

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DisableDrawing()
Const ProcName As String = "DisableDrawing"
On Error GoTo Err

Chart1.DisableDrawing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub EnableDrawing()
Const ProcName As String = "EnableDrawing"
On Error GoTo Err

Chart1.EnableDrawing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

'' update the number of bars in case this chart is reloaded from the config
'If Not mChartSpec Is Nothing Then
'    If mChartSpec.InitialNumberOfBars < Chart1.Periods.Count Then
'        Set mChartSpec = CreateChartSpecifier(Chart1.Periods.Count, mChartSpec.IncludeBarsOutsideSession)
'        storeSettings
'    End If
'End If

If Not mManager Is Nothing Then mManager.Finish

Set mManager = Nothing

Set mTimeframes = Nothing
Set mTimeframe = Nothing

Set mContract = Nothing

Set mPriceRegion = Nothing
Set mVolumeRegion = Nothing

Set mLoadingText = Nothing

mLoadedFromConfig = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LoadFromConfig( _
                ByVal pTimeframes As Timeframes, _
                ByVal config As ConfigurationSection, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal deferStart As Boolean)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config

gLogger.Log "Loading chart from config at: " & mConfig.Path, ProcName, ModuleName

mLoadedFromConfig = True

mDeferStart = deferStart

Set mTimeframes = pTimeframes
Set mStudyManager = mTimeframes.StudyBase.StudyManager
Set mBarFormatterLibManager = pBarFormatterLibManager

Set mTimePeriod = TimePeriodFromString(mConfig.GetSetting(ConfigSettingTimePeriod))
Set mChartSpec = LoadChartSpecifierFromConfig(mConfig.GetConfigurationSection(ConfigSectionChartSpecifier))

Chart1.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionChartControl)

mIsHistoricChart = CBool(mConfig.GetSetting(ConfigSettingIsHistoricChart, "False"))
mBarFormatterFactoryName = mConfig.GetSetting(ConfigSettingBarFormatterFactoryName, "")
mBarFormatterLibraryName = mConfig.GetSetting(ConfigSettingBarFormatterLibraryName, "")

If Not mDeferStart Then prepareChart

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

gLogger.Log "Chart removed from config at: " & mConfig.Path, ProcName, ModuleName

mConfig.Remove
Set mConfig = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScrollToTime(ByVal pTime As Date)
Const ProcName As String = "ScrollToTime"
On Error GoTo Err

mManager.ScrollToTime pTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ShowChart( _
                ByVal pTimeframes As Timeframes, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pChartSpec As ChartSpecifier, _
                ByVal pChartStyle As ChartStyle, _
                Optional ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                Optional ByVal pBarFormatterFactoryName As String, _
                Optional ByVal pBarFormatterLibraryName As String, _
                Optional ByVal pExcludeCurrentBar As Boolean, _
                Optional ByVal pTitle As String)
Const ProcName As String = "ShowChart"
On Error GoTo Err

AssertArgument pBarFormatterFactoryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterFactoryName is not blank then pBarFormatterLibManagermust be supplied"
AssertArgument pBarFormatterLibraryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterLibraryName is not blank then pBarFormatterLibManagermust be supplied"
AssertArgument (pBarFormatterLibraryName = "" And pBarFormatterFactoryName = "") Or (pBarFormatterLibraryName <> "" And pBarFormatterFactoryName <> ""), "If pBarFormatterLibraryName is not blank then pBarFormatterLibManagermust be supplied"

setState ChartStateBlank

If Not mTimeframes Is Nothing Then
    mInitialised = False
    Chart1.ClearChart
End If

Set mTimeframes = pTimeframes
Set mStudyManager = mTimeframes.StudyBase.StudyManager
Set mBarFormatterLibManager = pBarFormatterLibManager

Set mTimePeriod = pTimePeriod
Set mChartSpec = pChartSpec
Set mChartStyle = pChartStyle
mBarFormatterFactoryName = pBarFormatterFactoryName
mBarFormatterLibraryName = pBarFormatterLibraryName
mExcludeCurrentBar = pExcludeCurrentBar
mTitle = pTitle

storeSettings

prepareChart

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Start()
Const ProcName As String = "Start"
On Error GoTo Err

Assert mLoadedFromConfig And mState = ChartStates.ChartStateBlank, "Start method only permitted for charts loaded from configuration and with state ChartStateBlank"

mDeferStart = False
prepareChart

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createBarsStudyConfig() As StudyConfiguration
Const ProcName As String = "createBarsStudyConfig"
On Error GoTo Err

Dim studyConfig As New StudyConfiguration
studyConfig.UnderlyingStudy = mTimeframes.StudyBase.BaseStudy

Dim lStudy As IStudy
Set lStudy = mTimeframe.BarStudy
studyConfig.Study = lStudy

Dim studyDef As StudyDefinition
Set studyDef = lStudy.StudyDefinition

studyConfig.ChartRegionName = ChartRegionNamePrice

ReDim InputValueNames(3) As String
InputValueNames(0) = InputNameTrade
InputValueNames(1) = InputNameVolume
InputValueNames(2) = InputNameTickVolume
InputValueNames(3) = InputNameOpenInterest

studyConfig.InputValueNames = InputValueNames
studyConfig.Name = studyDef.Name

Dim params As New Parameters
params.SetParameterValue "Bar length", mTimePeriod.Length
params.SetParameterValue "Time units", TimePeriodUnitsToString(mTimePeriod.Units)
studyConfig.Parameters = params

Dim studyValueConfig As StudyValueConfiguration
Set studyValueConfig = studyConfig.StudyValueConfigurations.Add(StudyValueConfigNameBar)
studyValueConfig.ChartRegionName = ChartRegionNamePrice
studyValueConfig.IncludeInChart = True
'studyValueConfig.Layer = 200
studyValueConfig.BarFormatterFactoryName = mBarFormatterFactoryName
studyValueConfig.BarFormatterLibraryName = mBarFormatterLibraryName

If mContract Is Nothing Then
ElseIf mContract.Specifier.secType <> SecurityTypes.SecTypeCash And _
    mContract.Specifier.secType <> SecurityTypes.SecTypeIndex _
Then
    Set studyValueConfig = studyConfig.StudyValueConfigurations.Add(StudyValueConfigNameVolume)
    studyValueConfig.ChartRegionName = ChartRegionNameVolume
    studyValueConfig.IncludeInChart = True
End If

Set createBarsStudyConfig = studyConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createPriceFormatter() As PriceFormatter
Const ProcName As String = "createPriceFormatter"
On Error GoTo Err

Set createPriceFormatter = New PriceFormatter
If mContract Is Nothing Then
    createPriceFormatter.Initialise SecTypeNone, 0.0001
Else
    createPriceFormatter.Initialise mContract.Specifier.secType, mContract.TickSize
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub createTimeframe()
Const ProcName As String = "createTimeframe"
On Error GoTo Err

gLogger.Log "Creating timeframe", ProcName, ModuleName

If mChartSpec.toTime <> CDate(0) Then
    Set mTimeframe = mTimeframes.AddHistorical(mTimePeriod, _
                                "", _
                                mChartSpec.InitialNumberOfBars, _
                                mChartSpec.FromTime, _
                                mChartSpec.toTime, _
                                , _
                                mChartSpec.IncludeBarsOutsideSession)
Else
    Set mTimeframe = mTimeframes.Add(mTimePeriod, _
                                "", _
                                mChartSpec.InitialNumberOfBars, _
                                mChartSpec.FromTime, _
                                , _
                                mChartSpec.IncludeBarsOutsideSession, _
                                mExcludeCurrentBar)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseChart()
Const ProcName As String = "initialiseChart"
On Error GoTo Err

gLogger.Log "Initialising chart", ProcName, ModuleName

Chart1.DisableDrawing

If Not mInitialised Then
    Set mManager = CreateChartManager(Chart1.Controller, mStudyManager, mBarFormatterLibManager)
    mManager.UpdatePerTick = mUpdatePerTick
    If Not mConfig Is Nothing Then mManager.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudies)

    If mChartStyle Is Nothing Then
        gLogger.Log "No chart style is defined", ProcName, ModuleName
    Else
        gLogger.Log "Setting chart style to", ProcName, ModuleName, , mChartStyle.Name
    End If

    If Not mChartStyle Is Nothing Then Chart1.Style = mChartStyle
    mInitialised = True
End If

Set mPriceRegion = Chart1.Regions.Add(100, 25, , , ChartRegionNamePrice)
setLoadingText
Chart1.EnableDrawing

setState ChartStates.ChartStateCreated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadchart()
Const ProcName As String = "loadchart"
On Error GoTo Err

gLogger.Log "Loading chart", ProcName, ModuleName

Chart1.DisableDrawing

Chart1.TimePeriod = mTimePeriod

If mContract Is Nothing Then
    Chart1.SessionStartTime = 0#
    Chart1.SessionEndTime = 0#
Else
    Chart1.SessionStartTime = mContract.SessionStartTime
    Chart1.SessionEndTime = mContract.SessionEndTime
End If

If mContract Is Nothing Then
    mPriceRegion.YScaleQuantum = 0.001
Else
    mPriceRegion.YScaleQuantum = mContract.TickSize
    If mMinimumTicksHeight * mContract.TickSize <> 0 Then
        mPriceRegion.MinimumHeight = mMinimumTicksHeight * mContract.TickSize
    End If
End If

mPriceRegion.PriceFormatter = createPriceFormatter

If mTitle <> "" Then
    mPriceRegion.Title.Text = mTitle
ElseIf Not mContract Is Nothing Then
    mPriceRegion.Title.Text = mContract.Specifier.LocalSymbol & _
                    " (" & mContract.Specifier.Exchange & ") " & _
                    TimeframeCaption
End If
mPriceRegion.Title.Color = vbBlue

If mContract Is Nothing Then
ElseIf mContract.Specifier.secType <> SecurityTypes.SecTypeCash _
    And mContract.Specifier.secType <> SecurityTypes.SecTypeIndex _
Then
    On Error Resume Next
    Set mVolumeRegion = Chart1.Regions.Item(ChartRegionNameVolume)
    On Error GoTo Err

    If mVolumeRegion Is Nothing Then Set mVolumeRegion = Chart1.Regions.Add(20, , , , ChartRegionNameVolume)

    mVolumeRegion.MinimumHeight = 10
    mVolumeRegion.IntegerYScale = True
    mVolumeRegion.Title.Text = "Volume"
    mVolumeRegion.Title.Color = vbBlue
End If

If mTimeframe.State <> TimeframeStateLoaded Then
    mLoadingText.Text = "Fetching historical data"
    setState ChartStates.ChartStateInitialised
    Chart1.EnableDrawing    ' causes the loading text to appear
    Chart1.DisableDrawing
Else
    Chart1.EnableDrawing
    setState ChartStates.ChartStateInitialised
    mLoadingText.Text = ""
    setState ChartStates.ChartStateLoaded
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub loadStudiesFromConfig()
Const ProcName As String = "loadStudiesFromConfig"
On Error GoTo Err

mManager.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionStudies), mTimeframe.BarStudy

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub prepareChart()
Const ProcName As String = "prepareChart"
On Error GoTo Err

createTimeframe
initialiseChart

If mTimeframes.ContractFuture Is Nothing Then
    setupStudies
    loadchart
ElseIf mTimeframes.ContractFuture.IsAvailable Then
    Set mContract = mTimeframes.ContractFuture.value

    setupStudies
    loadchart
Else
    Set mFutureWaiter = New FutureWaiter
    mFutureWaiter.Add mTimeframes.ContractFuture
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setLoadingText()
Const ProcName As String = "setLoadingText"
On Error GoTo Err

Set mLoadingText = mPriceRegion.AddText(, ChartSkil27.LayerNumbers.LayerHighestUser)
Dim Font As New stdole.StdFont
Font.Size = 18
mLoadingText.Font = Font
mLoadingText.Color = vbBlack
mLoadingText.Box = True
mLoadingText.BoxFillColor = vbWhite
mLoadingText.BoxFillStyle = FillStyles.FillSolid
mLoadingText.Position = NewPoint(50, 50, CoordinateSystems.CoordsRelative, CoordinateSystems.CoordsRelative)
mLoadingText.align = TextAlignModes.AlignBoxCentreCentre
mLoadingText.FixedX = True
mLoadingText.FixedY = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setState(ByVal value As ChartStates)
Const ProcName As String = "setState"
On Error GoTo Err

Dim stateEv As StateChangeEventData

mState = value
stateEv.State = mState
Set stateEv.Source = Me
RaiseEvent StateChange(stateEv)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupStudies()
Const ProcName As String = "setupStudies"
On Error GoTo Err

gLogger.Log "Setting up studies", ProcName, ModuleName

If mLoadedFromConfig Then
    loadStudiesFromConfig
Else
    mManager.BaseStudyConfiguration = createBarsStudyConfig
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeSettings()
Const ProcName As String = "storeSettings"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

'If mDataSource Is Nothing Then Exit Sub

'mConfig.SetSetting ConfigSettingDataSourceKey, mKey
mConfig.SetSetting ConfigSettingTimePeriod, mTimePeriod.ToString
If Not mChartSpec Is Nothing Then mChartSpec.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionChartSpecifier)
mConfig.SetSetting ConfigSettingIsHistoricChart, CStr(mIsHistoricChart)
mConfig.SetSetting ConfigSettingBarFormatterFactoryName, mBarFormatterFactoryName
mConfig.SetSetting ConfigSettingBarFormatterLibraryName, mBarFormatterLibraryName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
