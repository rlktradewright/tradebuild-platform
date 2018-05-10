VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5EF6A0B6-9E1F-426C-B84A-601F4CBF70C4}#276.0#0"; "ChartSkil27.ocx"
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

Implements IThemeable

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

Event StyleChanged(ByVal pNewStyle As ChartStyle)

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

Private mLocalSymbol                                    As String
Private mSecType                                        As SecurityTypes
Private mExchange                                       As String
Private mTickSize                                       As Double
Private mSessionEndTime                                 As Date
Private mSessionStartTime                               As Date

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
Private mReadyForDeferredStart                          As Boolean
Private mDeferredStartRequested                         As Boolean

Private mMinimumTicksHeight                             As Long

Private mInitialised                                    As Boolean

Private mExcludeCurrentBar                              As Boolean

Private WithEvents mFutureWaiter                        As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTitle                                          As String

Private mTheme                                          As ITheme

Private mIsRaw                                          As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Set mFutureWaiter = New FutureWaiter
mPrevWidth = UserControl.Width
mPrevHeight = UserControl.Height

mUpdatePerTick = True
mMinimumTicksHeight = 10

End Sub

Private Sub UserControl_Resize()
mPrevWidth = UserControl.Width
Chart1.Height = UserControl.Height
mPrevHeight = UserControl.Height
End Sub

Private Sub UserControl_Terminate()
Const ProcName As String = "UserControl_Terminate"
gLogger.Log "MarketChart terminated", ProcName, ModuleName, LogLevelDetail
Debug.Print "MarketChart terminated"
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

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

Private Sub Chart1_StyleChanged(ByVal pNewStyle As ChartStyle)
RaiseEvent StyleChanged(pNewStyle)
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub
If TypeOf ev.Future.Value Is IContract Then
    setContractProperties mTimeframes.ContractFuture.Value
    If mDeferStart Then
        mReadyForDeferredStart = True
        If mDeferredStartRequested Then Start
    Else
        initialiseChart mChartSpec.IncludeBarsOutsideSession
        prepareChart
    End If
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
LoadingProgressBar.Value = percentComplete

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Autoscrolling( _
                ByVal Value As Boolean)
Const ProcName As String = "Autoscrolling"
On Error GoTo Err

Chart1.Autoscrolling = Value

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
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If mConfig Is Value Then Exit Property
If Not mConfig Is Nothing Then mConfig.Remove
Set mConfig = Nothing
If Value Is Nothing Then Exit Property

Set mConfig = Value

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
                ByVal Value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled = Value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HorizontalMouseScrollingAllowed( _
                ByVal Value As Boolean)
Const ProcName As String = "HorizontalMouseScrollingAllowed"
On Error GoTo Err

Chart1.HorizontalMouseScrollingAllowed = Value

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

Public Property Let MinimumTicksHeight(ByVal Value As Double)
Const ProcName As String = "MinimumTicksHeight"
On Error GoTo Err

mMinimumTicksHeight = Value
If mMinimumTicksHeight * mTickSize <> 0 Then
    mPriceRegion.MinimumHeight = mMinimumTicksHeight * mTickSize
End If

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

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Get PeriodWidth() As Long
Const ProcName As String = "PeriodWidth"
On Error GoTo Err

PeriodWidth = Chart1.PeriodWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PeriodWidth(ByVal Value As Long)
Const ProcName As String = "PeriodWidth"
On Error GoTo Err

Chart1.PeriodWidth = Value

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

Public Property Let PointerCrosshairsColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "PointerCrosshairsColor"
On Error GoTo Err

Chart1.PointerCrosshairsColor = Value

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

Public Property Let PointerDiscColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "PointerDiscColor"
On Error GoTo Err

Chart1.PointerDiscColor = Value

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

Public Property Let PointerStyle(ByVal Value As PointerStyles)
Const ProcName As String = "PointerStyle"
On Error GoTo Err

Chart1.PointerStyle = Value

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

Public Property Let StudyManager(ByVal Value As StudyManager)
Const ProcName As String = "StudyManager"
On Error GoTo Err

Set mStudyManager = Value
mManager.Finish
Set mManager = CreateChartManager(Chart1.Controller, mStudyManager, mBarFormatterLibManager, False)
mManager.UpdatePerTick = mUpdatePerTick
initialiseChart mChartSpec.IncludeBarsOutsideSession

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

UserControl.BackColor = mTheme.BackColor
gApplyTheme mTheme, UserControl.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

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

Set TradeBarSeries = mManager.BaseStudyConfiguration.ValueSeries(BarStudyValueBar)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let VerticalMouseScrollingAllowed( _
                ByVal Value As Boolean)
Const ProcName As String = "VerticalMouseScrollingAllowed"
On Error GoTo Err

Chart1.VerticalMouseScrollingAllowed = Value

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

Public Property Let YAxisWidthCm(ByVal Value As Single)
Const ProcName As String = "YAxisWidthCm"
On Error GoTo Err

Chart1.YAxisWidthCm = Value

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

Set mTimeframe = createTimeframe(mTimeframes, mTimePeriod, mChartSpec, mExcludeCurrentBar)

baseStudyConfig.Study = mTimeframe.BarStudy
baseStudyConfig.StudyValueConfigurations.Item(BarStudyValueBar).BarFormatterFactoryName = mBarFormatterFactoryName
baseStudyConfig.StudyValueConfigurations.Item(BarStudyValueBar).BarFormatterLibraryName = mBarFormatterLibraryName

Dim lStudy As IStudy
Set lStudy = mTimeframe.BarStudy
baseStudyConfig.Parameters = lStudy.Parameters

initialiseChart mChartSpec.IncludeBarsOutsideSession
mManager.SetBaseStudyConfiguration baseStudyConfig

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

If Not mManager Is Nothing Then mManager.Finish

Set mManager = Nothing

Set mTimeframes = Nothing
Set mTimeframe = Nothing

mLocalSymbol = ""
mSecType = SecurityTypes.SecTypeNone
mExchange = ""
mTickSize = 0#
mSessionEndTime = 0#
mSessionStartTime = 0#

Set mPriceRegion = Nothing
Set mVolumeRegion = Nothing

Set mLoadingText = Nothing

mLoadedFromConfig = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pTimeframes As Timeframes, _
                ByVal pUpdatePerTick As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Assert Not mIsRaw, "Already initialised as raw"
Set mTimeframes = pTimeframes
Set mStudyManager = mTimeframes.StudyBase.StudyManager
mFutureWaiter.Add mTimeframes.ContractFuture
mUpdatePerTick = pUpdatePerTick

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub InitialiseRaw( _
                ByVal pStudyManager As StudyManager, _
                ByVal pUpdatePerTick As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

mIsRaw = True
Set mStudyManager = pStudyManager
mUpdatePerTick = pUpdatePerTick

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
mFutureWaiter.Add mTimeframes.ContractFuture

Set mBarFormatterLibManager = pBarFormatterLibManager

Set mTimePeriod = TimePeriodFromString(mConfig.GetSetting(ConfigSettingTimePeriod))
Set mChartSpec = LoadChartSpecifierFromConfig(mConfig.GetConfigurationSection(ConfigSectionChartSpecifier))

Chart1.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionChartControl)

mIsHistoricChart = CBool(mConfig.GetSetting(ConfigSettingIsHistoricChart, "False"))
mBarFormatterFactoryName = mConfig.GetSetting(ConfigSettingBarFormatterFactoryName, "")
mBarFormatterLibraryName = mConfig.GetSetting(ConfigSettingBarFormatterLibraryName, "")

If Not mDeferStart Then Set mTimeframe = createTimeframe(mTimeframes, mTimePeriod, mChartSpec, False)

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

AssertArgument pBarFormatterFactoryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterFactoryName is not blank then pBarFormatterLibManager must be supplied"
AssertArgument pBarFormatterLibraryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterLibraryName is not blank then pBarFormatterLibManager must be supplied"
AssertArgument (pBarFormatterLibraryName = "" And pBarFormatterFactoryName = "") Or (pBarFormatterLibraryName <> "" And pBarFormatterFactoryName <> ""), "pBarFormatterLibraryName and pBarFormatterFactoryName must both be blank or non-blank"

setState ChartStateBlank

If Not mTimeframes Is Nothing Then
    mInitialised = False
    Chart1.ClearChart
End If

Set mBarFormatterLibManager = pBarFormatterLibManager

Set mTimePeriod = pTimePeriod
Set mChartSpec = pChartSpec
If Not pChartStyle Is Nothing Then Set mChartStyle = pChartStyle
mBarFormatterFactoryName = pBarFormatterFactoryName
mBarFormatterLibraryName = pBarFormatterLibraryName
mExcludeCurrentBar = pExcludeCurrentBar
mTitle = pTitle

storeSettings
Set mTimeframe = createTimeframe(mTimeframes, mTimePeriod, mChartSpec, mExcludeCurrentBar)

setState ChartStates.ChartStateCreated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ShowChartRaw( _
                ByVal pTimeframe As Timeframe, _
                ByVal pChartStyle As ChartStyle, _
                Optional ByVal pLocalSymbol As String, _
                Optional ByVal pSecType As SecurityTypes, _
                Optional ByVal pExchange As String, _
                Optional ByVal pTickSize As Double, _
                Optional ByVal pSessionStartTime As Date, _
                Optional ByVal pSessionEndTime As Date, _
                Optional ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                Optional ByVal pBarFormatterFactoryName As String, _
                Optional ByVal pBarFormatterLibraryName As String, _
                Optional ByVal pTitle As String)
Const ProcName As String = "ShowChartRaw"
On Error GoTo Err

AssertArgument pBarFormatterFactoryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterFactoryName is not blank then pBarFormatterLibManager must be supplied"
AssertArgument pBarFormatterLibraryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterLibraryName is not blank then pBarFormatterLibManager must be supplied"
AssertArgument (pBarFormatterLibraryName = "" And pBarFormatterFactoryName = "") Or (pBarFormatterLibraryName <> "" And pBarFormatterFactoryName <> ""), "pBarFormatterLibraryName and pBarFormatterFactoryName must both be blank or non-blank"

setState ChartStateBlank

Set mBarFormatterLibManager = pBarFormatterLibManager

Set mTimeframe = pTimeframe
Set mTimePeriod = mTimeframe.TimePeriod
If Not pChartStyle Is Nothing Then Set mChartStyle = pChartStyle
mLocalSymbol = pLocalSymbol
mSecType = pSecType
mExchange = pExchange
mTickSize = pTickSize
mSessionEndTime = pSessionEndTime
Chart1.SessionEndTime = mSessionEndTime
mSessionStartTime = pSessionStartTime
Chart1.SessionStartTime = mSessionStartTime
mBarFormatterFactoryName = pBarFormatterFactoryName
mBarFormatterLibraryName = pBarFormatterLibraryName
mTitle = pTitle

initialiseChart False
prepareChart

setState ChartStates.ChartStateCreated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Start()
Const ProcName As String = "Start"
On Error GoTo Err

Assert mLoadedFromConfig And mState = ChartStates.ChartStateBlank, "Start method only permitted for charts loaded from configuration and with state ChartStateBlank"

If Not mReadyForDeferredStart Then
    mDeferredStartRequested = True
    Exit Sub
End If

setState ChartStates.ChartStateCreated

mReadyForDeferredStart = False
mDeferStart = False
Set mTimeframe = createTimeframe(mTimeframes, mTimePeriod, mChartSpec, False)
initialiseChart mChartSpec.IncludeBarsOutsideSession
prepareChart

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub UpdateLastBar()
Const ProcName As String = "UpdateLastBar"
On Error GoTo Err

mManager.UpdateLastBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createPriceFormatter() As PriceFormatter
Const ProcName As String = "createPriceFormatter"
On Error GoTo Err

Set createPriceFormatter = New PriceFormatter
If mTickSize = 0# Then
    createPriceFormatter.Initialise SecTypeNone, 0.0001
Else
    createPriceFormatter.Initialise mSecType, mTickSize
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createTimeframe( _
                ByVal pTimeframes As Timeframes, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pChartSpec As ChartSpecifier, _
                ByVal pExcludeCurrentBar As Boolean) As Timeframe
Const ProcName As String = "createTimeframe"
On Error GoTo Err

gLogger.Log "Creating timeframe", ProcName, ModuleName

If pChartSpec.toTime <> CDate(0) Then
    Set createTimeframe = pTimeframes.AddHistorical(pTimePeriod, _
                                "", _
                                pChartSpec.InitialNumberOfBars, _
                                pChartSpec.FromTime, _
                                pChartSpec.toTime, _
                                , _
                                pChartSpec.IncludeBarsOutsideSession)
Else
    Set createTimeframe = pTimeframes.Add(pTimePeriod, _
                                "", _
                                pChartSpec.InitialNumberOfBars, _
                                pChartSpec.FromTime, _
                                , _
                                pChartSpec.IncludeBarsOutsideSession, _
                                pExcludeCurrentBar)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub initialiseChart(ByVal pIncludeBarsOutsideSession As Boolean)
Const ProcName As String = "initialiseChart"
On Error GoTo Err

gLogger.Log "Initialising chart", ProcName, ModuleName

Chart1.DisableDrawing

If Not mInitialised Then
    Set mManager = CreateChartManager(Chart1.Controller, mStudyManager, mBarFormatterLibManager, pIncludeBarsOutsideSession)
    mManager.UpdatePerTick = mUpdatePerTick
    If Not mConfig Is Nothing Then mManager.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudies)

    If mChartStyle Is Nothing Then
        gLogger.Log "No chart style is defined", ProcName, ModuleName
    Else
        gLogger.Log "Setting chart style to", ProcName, ModuleName, , mChartStyle.Name
        Chart1.Style = mChartStyle
    End If

    mInitialised = True
End If

Set mPriceRegion = Chart1.Regions.Add(100, 25, , , ChartRegionNamePrice)
setLoadingText
Chart1.EnableDrawing

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

Chart1.SessionStartTime = mSessionStartTime
Chart1.SessionEndTime = mSessionEndTime

If mTickSize = 0# Then
    mPriceRegion.YScaleQuantum = 0.001
Else
    mPriceRegion.YScaleQuantum = mTickSize
    If mMinimumTicksHeight * mTickSize <> 0 Then
        mPriceRegion.MinimumHeight = mMinimumTicksHeight * mTickSize
    End If
End If

mPriceRegion.PriceFormatter = createPriceFormatter

If mTitle <> "" Then
    mPriceRegion.Title.Text = mTitle
ElseIf mLocalSymbol <> "" Then
    mPriceRegion.Title.Text = mLocalSymbol & _
                    " (" & mExchange & ") " & _
                    TimeframeCaption
End If
mPriceRegion.Title.Color = vbBlue

If mSecType = SecTypeNone Then
ElseIf mSecType <> SecurityTypes.SecTypeCash _
    And mSecType <> SecurityTypes.SecTypeIndex _
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

If mIsRaw Then
    Chart1.EnableDrawing
    setState ChartStates.ChartStateInitialised
    mLoadingText.Text = ""
    setState ChartStates.ChartStateLoaded
ElseIf mTimeframe.State <> TimeframeStateLoaded Then
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

mManager.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionStudies), mTimeframe.BarStudy, mChartSpec.IncludeBarsOutsideSession

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub prepareChart()
Const ProcName As String = "prepareChart"
On Error GoTo Err

If mTimeframes Is Nothing Then Assert Not mTimeframe Is Nothing, "mTimeframe Is Nothing"
    
setupStudies
loadchart

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setContractProperties(ByVal pContract As IContract)
mLocalSymbol = pContract.Specifier.LocalSymbol
mSecType = pContract.Specifier.secType
mExchange = pContract.Specifier.Exchange
mTickSize = pContract.TickSize
mSessionEndTime = pContract.SessionEndTime
Chart1.SessionEndTime = mSessionEndTime
mSessionStartTime = pContract.SessionStartTime
Chart1.SessionStartTime = mSessionStartTime
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

Private Sub setState(ByVal Value As ChartStates)
Const ProcName As String = "setState"
On Error GoTo Err

Dim stateEv As StateChangeEventData

mState = Value
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
    mManager.SetBaseStudyConfiguration CreateBarsStudyConfig( _
                                                        mTimeframe, _
                                                        mSecType, _
                                                        mStudyManager.StudyLibraryManager, _
                                                        mBarFormatterFactoryName, _
                                                        mBarFormatterLibraryName)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeSettings()
Const ProcName As String = "storeSettings"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

mConfig.SetSetting ConfigSettingTimePeriod, mTimePeriod.ToString
If Not mChartSpec Is Nothing Then mChartSpec.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionChartSpecifier)
mConfig.SetSetting ConfigSettingIsHistoricChart, CStr(mIsHistoricChart)
mConfig.SetSetting ConfigSettingBarFormatterFactoryName, mBarFormatterFactoryName
mConfig.SetSetting ConfigSettingBarFormatterLibraryName, mBarFormatterLibraryName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
