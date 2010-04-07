VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{74951842-2BEF-4829-A34F-DC7795A37167}#146.0#0"; "ChartSkil2-6.ocx"
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

Private mUpdatePerTick                                  As Boolean

Private mState                                          As ChartStates

Private mIsHistoricChart                                As Boolean

Private mChartSpec                                      As ChartSpecifier

Private mFromTime                                       As Date
Private mToTime                                         As Date

Private mcontract                                       As Contract

Private mPriceRegion                                    As ChartRegion

Private mVolumeRegion                                   As ChartRegion

Private mPrevWidth                                      As Single
Private mPrevHeight                                     As Single

Private mLoadingText                                    As Text

Private mBarFormatterFactory                            As BarFormatterFactory

Private mConfig                                         As ConfigurationSection
Private mLoadedFromConfig                               As Boolean

Private mTradeBarSeries                                 As BarSeries

Private mDeferStart                                     As Boolean

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
Const ProcName As String = "UserControl_Terminate"
gLogger.Log "TradeBuildChart terminated", ProcName, ModuleName, LogLevelDetail
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
Const ProcName As String = "mTicker_StateChange"
Dim failpoint As Long
On Error GoTo Err

If ev.State = TickerStates.TickerStateReady Then
    ' this means that the Ticker object has retrieved the contract info, so we can
    ' now start the chart
    
    If mDeferStart Then Exit Sub

    loadchart
    If mLoadedFromConfig Then
        loadStudiesFromConfig
    Else
        showStudies createBarsStudyConfig
    End If
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mTimeframe Event Handlers
'@================================================================================

Private Sub mTimeframe_BarsLoaded()
Const ProcName As String = "mTimeframe_BarsLoaded"
Dim failpoint As Long
On Error GoTo Err

LoadingProgressBar.Visible = False
mLoadingText.Text = ""
Chart1.EnableDrawing

setState ChartStates.ChartStateLoaded

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub mTimeframe_BarLoadProgress(ByVal barsRetrieved As Long, ByVal percentComplete As Single)
Const ProcName As String = "mTimeframe_BarLoadProgress"
Dim failpoint As Long
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
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Autoscrolling( _
                ByVal value As Boolean)
Const ProcName As String = "Autoscrolling"
Dim failpoint As Long
On Error GoTo Err

Chart1.Autoscrolling = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get Autoscrolling() As Boolean
Const ProcName As String = "Autoscrolling"
Dim failpoint As Long
On Error GoTo Err

Autoscrolling = Chart1.Autoscrolling

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get BaseChartController() As ChartController
Const ProcName As String = "BaseChartController"
Dim failpoint As Long
On Error GoTo Err

Set BaseChartController = Chart1.Controller

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Const ProcName As String = "ChartBackColor"
Dim failpoint As Long
On Error GoTo Err

ChartBackColor = Chart1.ChartBackColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let ChartBackColor(ByVal val As OLE_COLOR)
Const ProcName As String = "ChartBackColor"
Dim failpoint As Long
On Error GoTo Err

Chart1.ChartBackColor = val

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get ChartManager() As ChartManager
Const ProcName As String = "ChartManager"
Dim failpoint As Long
On Error GoTo Err

Set ChartManager = mManager

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
Dim failpoint As Long
On Error GoTo Err

If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings

If Not mManager Is Nothing Then mManager.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudies)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Const ProcName As String = "Enabled"
Dim failpoint As Long
On Error GoTo Err

Enabled = UserControl.Enabled

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
Const ProcName As String = "Enabled"
Dim failpoint As Long
On Error GoTo Err

UserControl.Enabled = value
PropertyChanged "Enabled"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let HorizontalMouseScrollingAllowed( _
                ByVal value As Boolean)
Const ProcName As String = "HorizontalMouseScrollingAllowed"
Dim failpoint As Long
On Error GoTo Err

Chart1.HorizontalMouseScrollingAllowed = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
Const ProcName As String = "HorizontalMouseScrollingAllowed"
Dim failpoint As Long
On Error GoTo Err

HorizontalMouseScrollingAllowed = Chart1.HorizontalMouseScrollingAllowed

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
Const ProcName As String = "HorizontalScrollBarVisible"
Dim failpoint As Long
On Error GoTo Err

HorizontalScrollBarVisible = Chart1.HorizontalScrollBarVisible

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let HorizontalScrollBarVisible(ByVal val As Boolean)
Const ProcName As String = "HorizontalScrollBarVisible"
Dim failpoint As Long
On Error GoTo Err

Chart1.HorizontalScrollBarVisible = val

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get InitialNumberOfBars() As Long
Attribute InitialNumberOfBars.VB_ProcData.VB_Invoke_Property = ";Behavior"
Const ProcName As String = "InitialNumberOfBars"
Dim failpoint As Long
On Error GoTo Err

InitialNumberOfBars = mChartSpec.InitialNumberOfBars

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get LoadingText() As Text
Const ProcName As String = "LoadingText"
Dim failpoint As Long
On Error GoTo Err

Set LoadingText = mLoadingText

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get MinimumTicksHeight() As Double
Attribute MinimumTicksHeight.VB_ProcData.VB_Invoke_Property = ";Behavior"
Const ProcName As String = "MinimumTicksHeight"
Dim failpoint As Long
On Error GoTo Err

MinimumTicksHeight = mChartSpec.MinimumTicksHeight

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get PointerCrosshairsColor() As OLE_COLOR
Const ProcName As String = "PointerCrosshairsColor"
Dim failpoint As Long
On Error GoTo Err

PointerCrosshairsColor = Chart1.PointerCrosshairsColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let PointerCrosshairsColor(ByVal value As OLE_COLOR)
Const ProcName As String = "PointerCrosshairsColor"
Dim failpoint As Long
On Error GoTo Err

Chart1.PointerCrosshairsColor = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get PointerDiscColor() As OLE_COLOR
Const ProcName As String = "PointerDiscColor"
Dim failpoint As Long
On Error GoTo Err

PointerDiscColor = Chart1.PointerDiscColor

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let PointerDiscColor(ByVal value As OLE_COLOR)
Const ProcName As String = "PointerDiscColor"
Dim failpoint As Long
On Error GoTo Err

Chart1.PointerDiscColor = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get PointerStyle() As PointerStyles
Const ProcName As String = "PointerStyle"
Dim failpoint As Long
On Error GoTo Err

PointerStyle = Chart1.PointerStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let PointerStyle(ByVal value As PointerStyles)
Const ProcName As String = "PointerStyle"
Dim failpoint As Long
On Error GoTo Err

Chart1.PointerStyle = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get PriceRegion() As ChartRegion
Const ProcName As String = "PriceRegion"
Dim failpoint As Long
On Error GoTo Err

Set PriceRegion = mPriceRegion

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get RegionNames() As String()
Const ProcName As String = "RegionNames"
Dim failpoint As Long
On Error GoTo Err

RegionNames = mManager.RegionNames

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get State() As ChartStates
Const ProcName As String = "State"
Dim failpoint As Long
On Error GoTo Err

State = mState

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get Ticker() As Ticker
Const ProcName As String = "Ticker"
Dim failpoint As Long
On Error GoTo Err

Set Ticker = mTicker

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get TimeframeCaption() As String
Const ProcName As String = "TimeframeCaption"
Dim failpoint As Long
On Error GoTo Err

TimeframeCaption = mChartSpec.Timeframe.ToString

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get TimeframeShortCaption() As String
Const ProcName As String = "TimeframeShortCaption"
Dim failpoint As Long
On Error GoTo Err

TimeframeShortCaption = mChartSpec.Timeframe.ToShortString

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get Timeframe() As Timeframe
Const ProcName As String = "Timeframe"
Dim failpoint As Long
On Error GoTo Err

Set Timeframe = mTimeframe

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get TimePeriod() As TimePeriod
Const ProcName As String = "TimePeriod"
Dim failpoint As Long
On Error GoTo Err

Set TimePeriod = mChartSpec.Timeframe

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Friend Property Get TradeBarSeries() As BarSeries
Const ProcName As String = "TradeBarSeries"
Dim failpoint As Long
On Error GoTo Err

Set TradeBarSeries = mTradeBarSeries

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get TwipsPerBar() As Long
Const ProcName As String = "TwipsPerBar"
Dim failpoint As Long
On Error GoTo Err

TwipsPerBar = Chart1.TwipsPerBar

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let TwipsPerBar(ByVal val As Long)
Const ProcName As String = "TwipsPerBar"
Dim failpoint As Long
On Error GoTo Err

Chart1.TwipsPerBar = val

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let UpdatePerTick(ByVal value As Boolean)
Attribute UpdatePerTick.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
Const ProcName As String = "UpdatePerTick"
Dim failpoint As Long
On Error GoTo Err

mUpdatePerTick = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let VerticalMouseScrollingAllowed( _
                ByVal value As Boolean)
Const ProcName As String = "VerticalMouseScrollingAllowed"
Dim failpoint As Long
On Error GoTo Err

Chart1.VerticalMouseScrollingAllowed = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
Const ProcName As String = "VerticalMouseScrollingAllowed"
Dim failpoint As Long
On Error GoTo Err

VerticalMouseScrollingAllowed = Chart1.VerticalMouseScrollingAllowed

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get VolumeRegion() As ChartRegion
Const ProcName As String = "VolumeRegion"
Dim failpoint As Long
On Error GoTo Err

Set VolumeRegion = mVolumeRegion

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get VolumeRegionStyle() As ChartRegionStyle
Const ProcName As String = "VolumeRegionStyle"
Dim failpoint As Long
On Error GoTo Err

Set VolumeRegionStyle = mChartSpec.VolumeRegionStyle

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get YAxisWidthCm() As Single
Const ProcName As String = "YAxisWidthCm"
Dim failpoint As Long
On Error GoTo Err

YAxisWidthCm = Chart1.YAxisWidthCm

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
Const ProcName As String = "YAxisWidthCm"
Dim failpoint As Long
On Error GoTo Err

Chart1.YAxisWidthCm = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ChangeTimeframe(ByVal Timeframe As TimePeriod)
Dim baseStudyConfig As StudyConfiguration

Const ProcName As String = "ChangeTimeframe"
Dim failpoint As Long
On Error GoTo Err

If State <> ChartStateLoaded Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                            ProjectName & "." & ModuleName & ":" & ProcName, _
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
baseStudyConfig.Study = mTimeframe.TradeStudy
baseStudyConfig.StudyValueConfigurations.item("Bar").SetBarFormatterFactory mBarFormatterFactory, mTimeframe.TradeBars
Dim lStudy As Study
Set lStudy = mTimeframe.TradeStudy
baseStudyConfig.Parameters = lStudy.Parameters

initialiseChart

loadchart
showStudies baseStudyConfig

RaiseEvent TimeframeChange

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Public Sub DisableDrawing()
Const ProcName As String = "DisableDrawing"
Dim failpoint As Long
On Error GoTo Err

Chart1.DisableDrawing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Public Sub EnableDrawing()
Const ProcName As String = "EnableDrawing"
Dim failpoint As Long
On Error GoTo Err

Chart1.EnableDrawing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Public Sub Finish()
' update the number of bars in case this chart is reloaded from the config
Const ProcName As String = "Finish"
Dim failpoint As Long
On Error GoTo Err

If Not mChartSpec Is Nothing Then
    If mChartSpec.InitialNumberOfBars < Chart1.Periods.Count Then
        mChartSpec.InitialNumberOfBars = Chart1.Periods.Count
    End If
End If

If Not mManager Is Nothing Then mManager.Finish

Set mManager = Nothing

Set mTimeframes = Nothing
Set mTimeframe = Nothing

Set mcontract = Nothing

Set mPriceRegion = Nothing
Set mVolumeRegion = Nothing
Set mTradeBarSeries = Nothing

Set mLoadingText = Nothing

mLoadedFromConfig = False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection, _
                ByVal deferStart As Boolean)
Dim cs As ConfigurationSection

Const ProcName As String = "LoadFromConfig"
Dim failpoint As Long
On Error GoTo Err

Set mConfig = config
mLoadedFromConfig = True

mDeferStart = deferStart

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

If Not deferStart Then prepareChart

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
Dim failpoint As Long
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove
Set mConfig = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Public Sub ScrollToTime(ByVal pTime As Date)
Const ProcName As String = "ScrollToTime"
Dim failpoint As Long
On Error GoTo Err

mManager.ScrollToTime pTime

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Public Sub ShowChart( _
                ByVal pTicker As Ticker, _
                ByVal chartSpec As ChartSpecifier, _
                Optional ByVal BarFormatterFactory As BarFormatterFactory)
Const ProcName As String = "ShowChart"
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
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Time period units not supported"
    
End Select

Set mTicker = pTicker
Set mChartSpec = chartSpec.Clone
Set mBarFormatterFactory = BarFormatterFactory

storeSettings

prepareChart

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Public Sub ShowHistoricChart( _
                ByVal pTicker As Ticker, _
                ByVal chartSpec As ChartSpecifier, _
                ByVal fromTime As Date, _
                ByVal toTime As Date, _
                Optional ByVal BarFormatterFactory As BarFormatterFactory)
Const ProcName As String = "ShowHistoricChart"
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
                ProjectName & "." & ModuleName & ":" & ProcName, _
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createBarsStudyConfig() As StudyConfiguration
Dim lStudy As Study
Dim studyDef As StudyDefinition

Const ProcName As String = "createBarsStudyConfig"
Dim failpoint As Long
On Error GoTo Err

ReDim inputValueNames(3) As String
Dim params As New Parameters

Dim studyValueConfig As StudyValueConfiguration
Dim BarsStyle As BarStyle
Dim VolumeStyle As DataPointStyle

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
params.SetParameterValue "Bar length", mChartSpec.Timeframe.length
params.SetParameterValue "Time units", TimePeriodUnitsToString(mChartSpec.Timeframe.Units)
studyConfig.Parameters = params

Set studyValueConfig = studyConfig.StudyValueConfigurations.Add("Bar")
studyValueConfig.ChartRegionName = ChartRegionNamePrice
studyValueConfig.IncludeInChart = True
studyValueConfig.Layer = 200
studyValueConfig.SetBarFormatterFactory mBarFormatterFactory, mTimeframe.TradeBars

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

If mcontract.Specifier.secType <> SecurityTypes.SecTypeCash And _
    mcontract.Specifier.secType <> SecurityTypes.SecTypeIndex _
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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Function createPriceFormatter() As PriceFormatter
Set createPriceFormatter = New PriceFormatter
createPriceFormatter.Contract = mcontract
End Function

Private Sub createTimeframe()
Const ProcName As String = "createTimeframe"
Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub

Private Sub initialiseChart()
Static notFirstTime As Boolean

Const ProcName As String = "initialiseChart"
Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub loadchart()
Dim volRegionStyle As ChartRegionStyle

Const ProcName As String = "loadchart"
Dim failpoint As Long
On Error GoTo Err

Set mcontract = mTicker.Contract

Chart1.DisableDrawing

Chart1.BarTimePeriod = mChartSpec.Timeframe

Chart1.SessionStartTime = mcontract.SessionStartTime
Chart1.SessionEndTime = mcontract.SessionEndTime

mPriceRegion.YScaleQuantum = mcontract.tickSize
If mChartSpec.MinimumTicksHeight * mcontract.tickSize <> 0 Then
    mPriceRegion.MinimumHeight = mChartSpec.MinimumTicksHeight * mcontract.tickSize
End If
mPriceRegion.PriceFormatter = createPriceFormatter

mPriceRegion.Title.Text = mcontract.Specifier.localSymbol & _
                " (" & mcontract.Specifier.exchange & ") " & _
                TimeframeCaption
mPriceRegion.Title.Color = vbBlue

If mcontract.Specifier.secType <> SecurityTypes.SecTypeCash _
    And mcontract.Specifier.secType <> SecurityTypes.SecTypeIndex _
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
    On Error GoTo Err
    
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

If Not mTimeframe.HistoricDataLoaded Then
    mLoadingText.Text = "Fetching historical data"
    setState ChartStates.ChartStateInitialised
    Chart1.EnableDrawing    ' causes the loading text to appear
    Chart1.DisableDrawing
Else
    Chart1.EnableDrawing
    setState ChartStates.ChartStateInitialised
    setState ChartStates.ChartStateLoaded
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub

Private Sub loadStudiesFromConfig()
Const ProcName As String = "loadStudiesFromConfig"
Dim failpoint As Long
On Error GoTo Err

mManager.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionStudies), mTimeframe.TradeStudy
setTradeBarSeries

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub prepareChart()

Const ProcName As String = "prepareChart"
Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub

Private Sub setLoadingText()
Const ProcName As String = "setLoadingText"
Dim failpoint As Long
On Error GoTo Err

Set mLoadingText = mPriceRegion.AddText(, ChartSkil26.LayerNumbers.LayerHighestUser)
Dim Font As New stdole.StdFont
Font.Size = 18
mLoadingText.Font = Font
mLoadingText.Color = vbBlack
mLoadingText.Box = True
mLoadingText.BoxFillColor = vbWhite
mLoadingText.BoxFillStyle = FillStyles.FillSolid
mLoadingText.Position = mPriceRegion.NewPoint(50, 50, CoordinateSystems.CoordsRelative, CoordinateSystems.CoordsRelative)
mLoadingText.align = TextAlignModes.AlignBoxCentreCentre
mLoadingText.FixedX = True
mLoadingText.FixedY = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub setState(ByVal value As ChartStates)
Dim stateEv As StateChangeEvent
Const ProcName As String = "setState"
Dim failpoint As Long
On Error GoTo Err

mState = value
stateEv.State = mState
Set stateEv.Source = Me
RaiseEvent StateChange(stateEv)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub setTradeBarSeries()
Const ProcName As String = "setTradeBarSeries"
Dim failpoint As Long
On Error GoTo Err

Set mTradeBarSeries = mManager.BaseStudyConfiguration.ValueSeries("Bar")

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub showStudies( _
                ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "showStudies"
Dim failpoint As Long
On Error GoTo Err

mManager.BaseStudyConfiguration = studyConfig
setTradeBarSeries

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub storeSettings()
Dim cs As ConfigurationSection

Const ProcName As String = "storeSettings"
Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub
