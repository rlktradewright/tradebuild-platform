VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{74951842-2BEF-4829-A34F-DC7795A37167}#208.0#0"; "ChartSkil2-6.ocx"
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

Event StateChange(ev As StateChangeEventData)

Event PeriodLengthChange()

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

Private Const ConfigSectionChartControl                 As String = "ChartControl"
Private Const ConfigSectionChartSpecifier               As String = "ChartSpecifier"
Private Const ConfigSectionStudies                      As String = "Studies"

Private Const ConfigSettingBarFormatterFactoryName      As String = "&BarFormatterFactoryName"
Private Const ConfigSettingBarFormatterLibraryName      As String = "&BarFormatterLibraryName"
Private Const ConfigSettingIsHistoricChart              As String = "&IsHistoricChart"
Private Const ConfigSettingPeriodLength                 As String = "&PeriodLength"
Private Const ConfigSettingTickerKey                    As String = "&TickerKey"
Private Const ConfigSettingWorkspace                    As String = "&Workspace"

Private Const PropNameHorizontalMouseScrollingAllowed   As String = "HorizontalMouseScrollingAllowed"
Private Const PropNameVerticalMouseScrollingAllowed     As String = "VerticalMouseScrollingAllowed"
Private Const PropNameAutoscroll                        As String = "Autoscrolling"
Private Const PropNameChartBackColor                    As String = "ChartBackColor"
Private Const PropNamePointerDiscColor                  As String = "PointerDiscColor"
Private Const PropNamePointerCrosshairsColor            As String = "PointerCrosshairsColor"
Private Const PropNamePointerStyle                      As String = "PointerStyle"
Private Const PropNameShowHorizontalScrollBar           As String = "HorizontalScrollBarVisible"
Private Const PropNameTwipsPerPeriod                    As String = "TwipsPerPeriod"
Private Const PropNameYAxisWidthCm                      As String = "YAxisWidthCm"

Private Const PropDfltHorizontalMouseScrollingAllowed   As Boolean = True
Private Const PropDfltVerticalMouseScrollingAllowed     As Boolean = True
Private Const PropDfltAutoscroll                        As Boolean = True
Private Const PropDfltChartBackColor                    As Long = vbWhite
Private Const PropDfltPointerDiscColor                  As Long = &H89FFFF
Private Const PropDfltPointerCrosshairsColor            As Long = &HC1DFE
Private Const PropDfltPointerStyle                      As Long = PointerStyles.PointerCrosshairs
Private Const PropDfltShowHorizontalScrollBar           As Boolean = True
Private Const PropDfltTwipsPerPeriod                    As Long = 150
Private Const PropDfltYAxisWidthCm                      As Single = 1.5

'@================================================================================
' Member variables
'@================================================================================

Private mManager                                        As ChartManager

Private WithEvents mTicker                              As Ticker
Attribute mTicker.VB_VarHelpID = -1
Private mTimeframes                                     As Timeframes
Private WithEvents mTimeframe                           As Timeframe
Attribute mTimeframe.VB_VarHelpID = -1

Private mPeriodLength                                   As TimePeriod

Private mUpdatePerTick                                  As Boolean

Private mState                                          As ChartStates

Private mIsHistoricChart                                As Boolean

Private mChartSpec                                      As ChartSpecifier
Private mChartStyle                                     As ChartStyle

Private mcontract                                       As Contract

Private mPriceRegion                                    As ChartRegion

Private mVolumeRegion                                   As ChartRegion

Private mPrevWidth                                      As Single
Private mPrevHeight                                     As Single

Private mLoadingText                                    As Text

Private mBarFormatterFactoryName                        As String
Private mBarFormatterLibraryName                        As String

Private mConfig                                         As ConfigurationSection
Private mLoadedFromConfig                               As Boolean

Private mDeferStart                                     As Boolean

Private mMinimumTicksHeight                             As Long

' this is a temporary style that is used initially to apply property bag settings.
' This prevents the property bag settings from overriding any style that is later
' applied.
Private mInitialStyle                                   As ChartStyle

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()

mPrevWidth = UserControl.Width
mPrevHeight = UserControl.Height

mUpdatePerTick = True

Set mInitialStyle = ChartStylesManager.Add(GenerateGUIDString, ChartStylesManager.DefaultStyle, pTemporary:=True)

End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next

mInitialStyle.HorizontalMouseScrollingAllowed = PropDfltHorizontalMouseScrollingAllowed
mInitialStyle.VerticalMouseScrollingAllowed = PropDfltVerticalMouseScrollingAllowed
mInitialStyle.Autoscrolling = PropDfltAutoscroll
mInitialStyle.ChartBackColor = PropDfltChartBackColor
PointerStyle = PropDfltPointerStyle
PointerCrosshairsColor = PropDfltPointerCrosshairsColor
PointerDiscColor = PropDfltPointerDiscColor
mInitialStyle.HorizontalScrollBarVisible = PropDfltShowHorizontalScrollBar
mInitialStyle.TwipsPerPeriod = PropDfltTwipsPerPeriod
mInitialStyle.YAxisWidthCm = PropDfltYAxisWidthCm

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

mInitialStyle.HorizontalMouseScrollingAllowed = PropBag.ReadProperty(PropNameHorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed)
mInitialStyle.VerticalMouseScrollingAllowed = PropBag.ReadProperty(PropNameVerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed)
mInitialStyle.Autoscrolling = PropBag.ReadProperty(PropNameAutoscroll, PropDfltAutoscroll)
mInitialStyle.ChartBackColor = PropBag.ReadProperty(PropNameChartBackColor)
' if no ChartBackColor has been set, we'll just use the ChartSkil default

PointerStyle = PropBag.ReadProperty(PropNamePointerStyle, PropDfltPointerStyle)
PointerCrosshairsColor = PropBag.ReadProperty(PropNamePointerCrosshairsColor, PropDfltPointerCrosshairsColor)
PointerDiscColor = PropBag.ReadProperty(PropNamePointerDiscColor, PropDfltPointerDiscColor)
mInitialStyle.HorizontalScrollBarVisible = PropBag.ReadProperty(PropNameShowHorizontalScrollBar, PropDfltShowHorizontalScrollBar)
mInitialStyle.TwipsPerPeriod = PropBag.ReadProperty(PropNameTwipsPerPeriod, PropDfltTwipsPerPeriod)
mInitialStyle.YAxisWidthCm = PropBag.ReadProperty(PropNameYAxisWidthCm, PropDfltYAxisWidthCm)
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
gLogger.Log "TradeBuildChart terminated", ProcName, ModuleName, LogLevelDetail
Debug.Print "TradeBuildChart terminated"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
PropBag.WriteProperty PropNameHorizontalMouseScrollingAllowed, mInitialStyle.HorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed
PropBag.WriteProperty PropNameVerticalMouseScrollingAllowed, mInitialStyle.VerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed
PropBag.WriteProperty PropNameAutoscroll, mInitialStyle.Autoscrolling, PropDfltAutoscroll
PropBag.WriteProperty PropNameChartBackColor, mInitialStyle.ChartBackColor, PropDfltChartBackColor
PropBag.WriteProperty PropNamePointerStyle, PointerStyle, PropDfltPointerStyle
PropBag.WriteProperty PropNamePointerCrosshairsColor, PointerCrosshairsColor, PropDfltPointerCrosshairsColor
PropBag.WriteProperty PropNamePointerDiscColor, PointerDiscColor, PropDfltPointerDiscColor
PropBag.WriteProperty PropNameShowHorizontalScrollBar, mInitialStyle.HorizontalScrollBarVisible, PropDfltShowHorizontalScrollBar
PropBag.WriteProperty PropNameTwipsPerPeriod, mInitialStyle.TwipsPerPeriod, PropDfltTwipsPerPeriod
PropBag.WriteProperty PropNameYAxisWidthCm, mInitialStyle.YAxisWidthCm, PropDfltYAxisWidthCm
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

Private Sub mTicker_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mTicker_StateChange"

On Error GoTo Err

If ev.State = TickerStates.TickerStateReady Then
    ' this means that the Ticker object has retrieved the contract info, so we can
    ' now start the chart
    
    Set mcontract = mTicker.Contract
    
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Autoscrolling() As Boolean
Const ProcName As String = "Autoscrolling"

On Error GoTo Err

Autoscrolling = Chart1.Autoscrolling

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BaseChartController() As ChartController
Const ProcName As String = "BaseChartController"

On Error GoTo Err

Set BaseChartController = Chart1.Controller

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Const ProcName As String = "ChartBackColor"

On Error GoTo Err

ChartBackColor = Chart1.ChartBackColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ChartBackColor(ByVal val As OLE_COLOR)
Const ProcName As String = "ChartBackColor"

On Error GoTo Err

Chart1.ChartBackColor = val

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ChartManager() As ChartManager
Const ProcName As String = "ChartManager"

On Error GoTo Err

Set ChartManager = mManager

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"

On Error GoTo Err

If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings

Chart1.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionChartControl)
If Not mManager Is Nothing Then mManager.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudies)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Const ProcName As String = "Enabled"

On Error GoTo Err

Enabled = UserControl.Enabled

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
Const ProcName As String = "Enabled"

On Error GoTo Err

UserControl.Enabled = value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HorizontalMouseScrollingAllowed( _
                ByVal value As Boolean)
Const ProcName As String = "HorizontalMouseScrollingAllowed"

On Error GoTo Err

Chart1.HorizontalMouseScrollingAllowed = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
Const ProcName As String = "HorizontalMouseScrollingAllowed"

On Error GoTo Err

HorizontalMouseScrollingAllowed = Chart1.HorizontalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
Const ProcName As String = "HorizontalScrollBarVisible"

On Error GoTo Err

HorizontalScrollBarVisible = Chart1.HorizontalScrollBarVisible

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HorizontalScrollBarVisible(ByVal val As Boolean)
Const ProcName As String = "HorizontalScrollBarVisible"

On Error GoTo Err

Chart1.HorizontalScrollBarVisible = val

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get InitialNumberOfBars() As Long
Attribute InitialNumberOfBars.VB_ProcData.VB_Invoke_Property = ";Behavior"
Const ProcName As String = "InitialNumberOfBars"

On Error GoTo Err

InitialNumberOfBars = mChartSpec.InitialNumberOfBars

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LoadingText() As Text
Const ProcName As String = "LoadingText"

On Error GoTo Err

Set LoadingText = mLoadingText

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let MinimumTicksHeight(ByVal value As Double)
Const ProcName As String = "MinimumTicksHeight"

On Error GoTo Err

mMinimumTicksHeight = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MinimumTicksHeight() As Double
Attribute MinimumTicksHeight.VB_ProcData.VB_Invoke_Property = ";Behavior"
Const ProcName As String = "MinimumTicksHeight"

On Error GoTo Err

MinimumTicksHeight = mMinimumTicksHeight

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerCrosshairsColor() As OLE_COLOR
Const ProcName As String = "PointerCrosshairsColor"

On Error GoTo Err

PointerCrosshairsColor = Chart1.PointerCrosshairsColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PointerCrosshairsColor(ByVal value As OLE_COLOR)
Const ProcName As String = "PointerCrosshairsColor"

On Error GoTo Err

Chart1.PointerCrosshairsColor = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerDiscColor() As OLE_COLOR
Const ProcName As String = "PointerDiscColor"

On Error GoTo Err

PointerDiscColor = Chart1.PointerDiscColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PointerDiscColor(ByVal value As OLE_COLOR)
Const ProcName As String = "PointerDiscColor"

On Error GoTo Err

Chart1.PointerDiscColor = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PointerStyle() As PointerStyles
Const ProcName As String = "PointerStyle"

On Error GoTo Err

PointerStyle = Chart1.PointerStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PointerStyle(ByVal value As PointerStyles)
Const ProcName As String = "PointerStyle"

On Error GoTo Err

Chart1.PointerStyle = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PriceRegion() As ChartRegion
Const ProcName As String = "PriceRegion"

On Error GoTo Err

Set PriceRegion = mPriceRegion

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get RegionNames() As String()
Const ProcName As String = "RegionNames"

On Error GoTo Err

RegionNames = mManager.RegionNames

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get State() As ChartStates
Const ProcName As String = "State"

On Error GoTo Err

State = mState

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Ticker() As Ticker
Const ProcName As String = "Ticker"

On Error GoTo Err

Set Ticker = mTicker

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TimeframeCaption() As String
Const ProcName As String = "TimeframeCaption"

On Error GoTo Err

TimeframeCaption = mPeriodLength.ToString

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TimeframeShortCaption() As String
Const ProcName As String = "TimeframeShortCaption"

On Error GoTo Err

TimeframeShortCaption = mPeriodLength.ToShortString

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Timeframe() As Timeframe
Const ProcName As String = "Timeframe"

On Error GoTo Err

Set Timeframe = mTimeframe

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PeriodLength() As TimePeriod
Const ProcName As String = "TimePeriod"

On Error GoTo Err

Set PeriodLength = mPeriodLength

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get TradeBarSeries() As BarSeries
Const ProcName As String = "TradeBarSeries"

On Error GoTo Err

Set TradeBarSeries = mManager.BaseStudyConfiguration.ValueSeries("Bar")

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TwipsPerPeriod() As Long
Const ProcName As String = "TwipsPerPeriod"

On Error GoTo Err

TwipsPerPeriod = Chart1.TwipsPerPeriod

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let TwipsPerPeriod(ByVal value As Long)
Const ProcName As String = "TwipsPerPeriod"

On Error GoTo Err

Chart1.TwipsPerPeriod = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let UpdatePerTick(ByVal value As Boolean)
Attribute UpdatePerTick.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
Const ProcName As String = "UpdatePerTick"

On Error GoTo Err

mUpdatePerTick = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let VerticalMouseScrollingAllowed( _
                ByVal value As Boolean)
Const ProcName As String = "VerticalMouseScrollingAllowed"

On Error GoTo Err

Chart1.VerticalMouseScrollingAllowed = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
Const ProcName As String = "VerticalMouseScrollingAllowed"

On Error GoTo Err

VerticalMouseScrollingAllowed = Chart1.VerticalMouseScrollingAllowed

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get VolumeRegion() As ChartRegion
Const ProcName As String = "VolumeRegion"

On Error GoTo Err

Set VolumeRegion = mVolumeRegion

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get YAxisWidthCm() As Single
Const ProcName As String = "YAxisWidthCm"

On Error GoTo Err

YAxisWidthCm = Chart1.YAxisWidthCm

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
Const ProcName As String = "YAxisWidthCm"

On Error GoTo Err

Chart1.YAxisWidthCm = value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ChangePeriodLength(ByVal pNewPeriodLength As TimePeriod)
Const ProcName As String = "ChangePeriodLength"
On Error GoTo Err

Dim baseStudyConfig As StudyConfiguration

If State <> ChartStateLoaded Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                            ProjectName & "." & ModuleName & ":" & ProcName, _
                                            "Can't change timeframe until chart is loaded"

gLogger.Log "Changing timeframe to", ProcName, ModuleName, , pNewPeriodLength.ToString

mLoadedFromConfig = False

Set baseStudyConfig = mManager.BaseStudyConfiguration

Set mPriceRegion = Nothing
Set mVolumeRegion = Nothing

mManager.ClearChart

setState ChartStateBlank

Set mPeriodLength = pNewPeriodLength

createTimeframe

baseStudyConfig.Study = mTimeframe.TradeStudy
baseStudyConfig.StudyValueConfigurations.item("Bar").BarFormatterFactoryName = mBarFormatterFactoryName
baseStudyConfig.StudyValueConfigurations.item("Bar").BarFormatterLibraryName = mBarFormatterLibraryName
Dim lStudy As Study
Set lStudy = mTimeframe.TradeStudy
baseStudyConfig.Parameters = lStudy.Parameters

initialiseChart
mManager.BaseStudyConfiguration = baseStudyConfig
'setupStudies

loadchart

RaiseEvent PeriodLengthChange

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub DisableDrawing()
Const ProcName As String = "DisableDrawing"

On Error GoTo Err

Chart1.DisableDrawing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub EnableDrawing()
Const ProcName As String = "EnableDrawing"

On Error GoTo Err

Chart1.EnableDrawing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"

On Error GoTo Err

' update the number of bars in case this chart is reloaded from the config
If Not mChartSpec Is Nothing Then
    If mChartSpec.InitialNumberOfBars < Chart1.Periods.Count Then
        Set mChartSpec = CreateChartSpecifier(Chart1.Periods.Count, mChartSpec.IncludeBarsOutsideSession)
        storeSettings
    End If
End If

If Not mManager Is Nothing Then mManager.Finish

Set mManager = Nothing

Set mTimeframes = Nothing
Set mTimeframe = Nothing

Set mcontract = Nothing

Set mPriceRegion = Nothing
Set mVolumeRegion = Nothing

Set mLoadingText = Nothing

mLoadedFromConfig = False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection, _
                ByVal deferStart As Boolean)
Dim cs As ConfigurationSection

Const ProcName As String = "LoadFromConfig"

On Error GoTo Err

Set mConfig = config
mLoadedFromConfig = True

mDeferStart = deferStart

Set mTicker = TradeBuildAPI.WorkSpaces(mConfig.GetSetting(ConfigSettingWorkspace)).Tickers(mConfig.GetSetting(ConfigSettingTickerKey))
Set mPeriodLength = TimePeriodFromString(mConfig.GetSetting(ConfigSettingPeriodLength))
Set mChartSpec = LoadChartSpecifierFromConfig(mConfig.GetConfigurationSection(ConfigSectionChartSpecifier))

Chart1.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionChartControl)

mIsHistoricChart = CBool(mConfig.GetSetting(ConfigSettingIsHistoricChart, "False"))
mBarFormatterFactoryName = mConfig.GetSetting(ConfigSettingBarFormatterFactoryName, "")
mBarFormatterLibraryName = mConfig.GetSetting(ConfigSettingBarFormatterLibraryName, "")

If Not mDeferStart Then prepareChart

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"

On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove
Set mConfig = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ScrollToTime(ByVal pTime As Date)
Const ProcName As String = "ScrollToTime"

On Error GoTo Err

mManager.ScrollToTime pTime

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ShowChart( _
                ByVal pTicker As Ticker, _
                ByVal pTimeframe As TimePeriod, _
                ByVal pChartSpec As ChartSpecifier, _
                ByVal pChartStyle As ChartStyle, _
                Optional ByVal pBarFormatterFactoryName As String, _
                Optional ByVal pBarFormatterLibraryName As String)
Const ProcName As String = "ShowChart"

On Error GoTo Err

Set mTicker = pTicker
Set mPeriodLength = pTimeframe
Set mChartSpec = pChartSpec
Set mChartStyle = pChartStyle
mBarFormatterFactoryName = pBarFormatterFactoryName
mBarFormatterLibraryName = pBarFormatterLibraryName

storeSettings

prepareChart

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Start()
Const ProcName As String = "Start"
On Error GoTo Err

If Not (mLoadedFromConfig And mState = ChartStates.ChartStateBlank) Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Start method only permitted for charts loaded from configuration and with state ChartStateBlank"
End If

mDeferStart = False
prepareChart

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createBarsStudyConfig() As StudyConfiguration
Dim lStudy As Study
Dim studyDef As StudyDefinition

Const ProcName As String = "createBarsStudyConfig"

On Error GoTo Err

ReDim inputValueNames(3) As String
Dim params As New Parameters

Dim studyValueConfig As StudyValueConfiguration

Dim studyConfig As StudyConfiguration

Set studyConfig = New StudyConfiguration

studyConfig.UnderlyingStudy = mTicker.InputStudy

Set lStudy = mTimeframe.TradeStudy
studyConfig.Study = lStudy
Set studyDef = lStudy.StudyDefinition

studyConfig.ChartRegionName = ChartRegionNamePrice

inputValueNames(0) = mTicker.InputNameTrade
inputValueNames(1) = mTicker.InputNameVolume
inputValueNames(2) = mTicker.InputNameTickVolume
inputValueNames(3) = mTicker.InputNameOpenInterest
studyConfig.inputValueNames = inputValueNames
studyConfig.name = studyDef.name
params.SetParameterValue "Bar length", mPeriodLength.Length
params.SetParameterValue "Time units", TimePeriodUnitsToString(mPeriodLength.Units)
studyConfig.Parameters = params

Set studyValueConfig = studyConfig.StudyValueConfigurations.Add("Bar")
studyValueConfig.ChartRegionName = ChartRegionNamePrice
studyValueConfig.IncludeInChart = True
'studyValueConfig.Layer = 200
studyValueConfig.BarFormatterFactoryName = mBarFormatterFactoryName
studyValueConfig.BarFormatterLibraryName = mBarFormatterLibraryName

If mcontract.Specifier.secType <> SecurityTypes.SecTypeCash And _
    mcontract.Specifier.secType <> SecurityTypes.SecTypeIndex _
Then
    Set studyValueConfig = studyConfig.StudyValueConfigurations.Add("Volume")
    studyValueConfig.ChartRegionName = ChartRegionNameVolume
    studyValueConfig.IncludeInChart = True
End If

Set createBarsStudyConfig = studyConfig

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function createPriceFormatter() As PriceFormatter
Set createPriceFormatter = New PriceFormatter
createPriceFormatter.Contract = mcontract
End Function

Private Sub createTimeframe()
Const ProcName As String = "createTimeframe"
On Error GoTo Err

gLogger.Log "Creating timeframe", ProcName, ModuleName

Set mTimeframes = mTicker.Timeframes

If mChartSpec.toTime <> CDate(0) Then
    Set mTimeframe = mTimeframes.AddHistorical(mPeriodLength, _
                                "", _
                                mChartSpec.InitialNumberOfBars, _
                                mChartSpec.FromTime, _
                                mChartSpec.toTime, _
                                mChartSpec.IncludeBarsOutsideSession)
Else
    Set mTimeframe = mTimeframes.Add(mPeriodLength, _
                                "", _
                                mChartSpec.InitialNumberOfBars, _
                                mChartSpec.IncludeBarsOutsideSession, _
                                IIf(mTicker.ReplayingTickfile, True, False))
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub initialiseChart()
Const ProcName As String = "initialiseChart"
On Error GoTo Err

Static notFirstTime As Boolean

gLogger.Log "Initialising chart", ProcName, ModuleName

Chart1.DisableDrawing

If Not notFirstTime Then
    Set mManager = CreateChartManager(mTicker.StudyManager, Chart1.Controller)
    If Not mConfig Is Nothing Then mManager.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudies)
    
    If mChartStyle Is Nothing Then
        gLogger.Log "No chart style is defined", ProcName, ModuleName
    Else
        gLogger.Log "Setting chart style to", ProcName, ModuleName, , mChartStyle.name
    End If

    If Not mChartStyle Is Nothing Then Chart1.Style = mChartStyle
    notFirstTime = True
End If

Set mPriceRegion = Chart1.Regions.Add(100, 25, , , ChartRegionNamePrice)
setLoadingText
Chart1.EnableDrawing

setState ChartStates.ChartStateCreated

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub loadchart()
Const ProcName As String = "loadchart"
On Error GoTo Err

gLogger.Log "Loading chart", ProcName, ModuleName

Chart1.DisableDrawing

Chart1.PeriodLength = mPeriodLength

Chart1.SessionStartTime = mcontract.SessionStartTime
Chart1.SessionEndTime = mcontract.SessionEndTime

mPriceRegion.YScaleQuantum = mcontract.tickSize
If mMinimumTicksHeight * mcontract.tickSize <> 0 Then
    mPriceRegion.MinimumHeight = mMinimumTicksHeight * mcontract.tickSize
End If
mPriceRegion.PriceFormatter = createPriceFormatter

mPriceRegion.Title.Text = mcontract.Specifier.localSymbol & _
                " (" & mcontract.Specifier.exchange & ") " & _
                TimeframeCaption
mPriceRegion.Title.Color = vbBlue

If mcontract.Specifier.secType <> SecurityTypes.SecTypeCash _
    And mcontract.Specifier.secType <> SecurityTypes.SecTypeIndex _
Then
    On Error Resume Next
    Set mVolumeRegion = Chart1.Regions.item(ChartRegionNameVolume)
    On Error GoTo Err
    
    If mVolumeRegion Is Nothing Then Set mVolumeRegion = Chart1.Regions.Add(20, , , , ChartRegionNameVolume)
    
    mVolumeRegion.MinimumHeight = 10
    mVolumeRegion.IntegerYScale = True
    mVolumeRegion.Title.Text = "Volume"
    mVolumeRegion.Title.Color = vbBlue
End If

If Not mTimeframe.HistoricDataLoaded Then
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub loadStudiesFromConfig()
Const ProcName As String = "loadStudiesFromConfig"

On Error GoTo Err

mManager.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionStudies), mTimeframe.TradeStudy

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub prepareChart()

Const ProcName As String = "prepareChart"

On Error GoTo Err

createTimeframe
initialiseChart

If mTicker.State = TickerStates.TickerStateReady Or _
    mTicker.State = TickerStates.TickerStateRunning _
Then
    Set mcontract = mTicker.Contract

    setupStudies
    loadchart
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub setLoadingText()
Const ProcName As String = "setLoadingText"

On Error GoTo Err

Set mLoadingText = mPriceRegion.AddText(, ChartSkil26.LayerNumbers.LayerHighestUser)
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setState(ByVal value As ChartStates)
Dim stateEv As StateChangeEventData
Const ProcName As String = "setState"

On Error GoTo Err

mState = value
stateEv.State = mState
Set stateEv.Source = Me
RaiseEvent StateChange(stateEv)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub storeSettings()
Dim cs As ConfigurationSection

Const ProcName As String = "storeSettings"

On Error GoTo Err

If mConfig Is Nothing Then Exit Sub
    
If mTicker Is Nothing Then Exit Sub

mConfig.SetSetting ConfigSettingWorkspace, mTicker.Workspace.name
mConfig.SetSetting ConfigSettingTickerKey, mTicker.Key
mConfig.SetSetting ConfigSettingPeriodLength, mPeriodLength.ToString
mChartSpec.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionChartSpecifier)
mConfig.SetSetting ConfigSettingIsHistoricChart, CStr(mIsHistoricChart)
mConfig.SetSetting ConfigSettingBarFormatterFactoryName, mBarFormatterFactoryName
mConfig.SetSetting ConfigSettingBarFormatterLibraryName, mBarFormatterLibraryName

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


