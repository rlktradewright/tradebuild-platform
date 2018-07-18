VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form fChart 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   Begin TradingUI27.MultiChart MultiChart1 
      Align           =   1  'Align Top
      Height          =   5415
      Left            =   0
      TabIndex        =   3
      Top             =   330
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   9551
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fChart.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fChart.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fChart.frx":05AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fChart.frx":0706
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      _ExtentX        =   19897
      _ExtentY        =   582
      BandCount       =   4
      BandBorders     =   0   'False
      _CBWidth        =   11280
      _CBHeight       =   330
      _Version        =   "6.7.9816"
      BandBackColor1  =   -2147483638
      Child1          =   "ChartNavToolbar1"
      MinWidth1       =   5865
      MinHeight1      =   330
      Width1          =   5865
      UseCoolbarColors1=   0   'False
      NewRow1         =   0   'False
      BandBackColor2  =   -2147483638
      Child2          =   "BarFormatterPicker"
      MinWidth2       =   1185
      MinHeight2      =   270
      Width2          =   1185
      UseCoolbarColors2=   0   'False
      NewRow2         =   0   'False
      BandBackColor3  =   -2147483638
      Child3          =   "ChartStylePicker"
      MinWidth3       =   1185
      MinHeight3      =   330
      Width3          =   1185
      UseCoolbarColors3=   0   'False
      NewRow3         =   0   'False
      BandBackColor4  =   -2147483638
      Child4          =   "ChartToolsToolbar"
      MinWidth4       =   1410
      MinHeight4      =   330
      Width4          =   1410
      UseCoolbarColors4=   0   'False
      NewRow4         =   0   'False
      Begin TradingUI27.ChartStylePicker ChartStylePicker 
         Height          =   330
         Left            =   8400
         TabIndex        =   5
         ToolTipText     =   "Change the chart style"
         Top             =   0
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListWidth       =   3000
      End
      Begin TradingUI27.BarFormatterPicker BarFormatterPicker 
         Height          =   270
         Left            =   6990
         TabIndex        =   4
         ToolTipText     =   "Change the bar formatting"
         Top             =   30
         Width           =   1185
         _ExtentX        =   12594
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListWidth       =   3000
      End
      Begin TradingUI27.ChartNavToolbar ChartNavToolbar1 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   0
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   582
      End
      Begin MSComctlLib.Toolbar ChartToolsToolbar 
         Height          =   330
         Left            =   9810
         TabIndex        =   1
         Top             =   0
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "studies"
               Object.ToolTipText     =   "Manage the studies displayed on the chart"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "selection"
               Description     =   "Select a chart object"
               Object.ToolTipText     =   "Show selection pointer"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "lines"
               Object.ToolTipText     =   "Draw lines"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "fib"
               Object.ToolTipText     =   "Draw Fibonacci retracement lines"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "fChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements IDeferredAction
Implements IGenericTickListener
Implements IThemeable
Implements IStateChangeListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                            As String = "fChart"

Private Const ChartToolsCommandStudies              As String = "studies"
Private Const ChartToolsCommandSelection            As String = "selection"
Private Const ChartToolsCommandLines                As String = "lines"
Private Const ChartToolsCommandFib                  As String = "fib"

Private Const ConfigSectionMultiChart               As String = "MultiChart"

Private Const ConfigSettingHeight                   As String = "&Height"
Private Const ConfigSettingLeft                     As String = "&Left"
Private Const ConfigSettingTop                      As String = "&Top"
Private Const ConfigSettingWidth                    As String = "&Width"
Private Const ConfigSettingWindowstate              As String = "&Windowstate"


'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mDataSource                             As IMarketDataSource
Attribute mDataSource.VB_VarHelpID = -1

Private mSymbol                                 As String
Private mSecType                                As SecurityTypes
Private mTickSize                               As Double

Private mCurrentBid                             As String
Private mCurrentAsk                             As String
Private mCurrentTrade                           As String
Private mCurrentVolume                          As String
Private mCurrentHigh                            As String
Private mCurrentLow                             As String
Private mPreviousClose                          As String

Private mIsHistorical                           As Boolean

Private WithEvents mChartController             As ChartController
Attribute mChartController.VB_VarHelpID = -1

Private mCurrentTool                            As IChartTool

Private mConfig                                 As ConfigurationSection

Private mBarFormatterLibManager                 As BarFormatterLibManager

Private WithEvents mFutureWaiter                As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mOwner                                  As Variant

Private mIsInitialised                          As Boolean

Private mTheme                                  As ITheme

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Form_Activate()
Const ProcName As String = "Form_Activate"
On Error GoTo Err

syncStudyPicker

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Initialize()
Const ProcName As String = "Form_Initialize"
On Error GoTo Err

Set mFutureWaiter = New FutureWaiter

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Const ProcName As String = "Form_QueryUnload"
On Error GoTo Err

If Not mDataSource Is Nothing Then
    mDataSource.RemoveGenericTickListener Me
    mDataSource.RemoveStateChangeListener Me
End If

MultiChart1.Finish
If mIsHistorical Then
    If Not mDataSource Is Nothing Then mDataSource.Finish
    Set mDataSource = Nothing
End If
gUnsyncStudyPicker

Select Case UnloadMode
Case QueryUnloadConstants.vbFormControlMenu
    ' the chart has been closed by the user so remove it from the config
    If Not mConfig Is Nothing Then mConfig.Remove
Case QueryUnloadConstants.vbFormCode, _
        QueryUnloadConstants.vbAppWindows, _
        QueryUnloadConstants.vbAppTaskManager, _
        QueryUnloadConstants.vbFormMDIForm, _
        QueryUnloadConstants.vbFormOwner
    If Not mConfig Is Nothing Then updateSettings
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Resize()
Const ProcName As String = "Form_Resize"
On Error GoTo Err

resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Terminate()
Const ProcName As String = "Form_Terminate"
LogMessage "Chart form terminated", LogLevelDetail
End Sub

'================================================================================
' IDeferredAction Interface Members
'================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

Set mTheme = Data
Me.BackColor = mTheme.BackColor
gApplyTheme mTheme, Me.Controls

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' IGenericTickListener Interface Members
'================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Select Case ev.Tick.TickType
Case TickTypeBid
    mCurrentBid = getFormattedPrice(ev.Tick.Price)
Case TickTypeAsk
    mCurrentAsk = getFormattedPrice(ev.Tick.Price)
Case TickTypeClosePrice
    mPreviousClose = getFormattedPrice(ev.Tick.Price)
Case TickTypeHighPrice
    mCurrentHigh = getFormattedPrice(ev.Tick.Price)
Case TickTypeLowPrice
    mCurrentLow = getFormattedPrice(ev.Tick.Price)
Case TickTypeTrade
    mCurrentTrade = getFormattedPrice(ev.Tick.Price)
Case TickTypeVolume
    mCurrentVolume = CStr(ev.Tick.Size)
Case Else
    Exit Sub
End Select

setCaption

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

'================================================================================
' IStateChangeListener Interface Members
'================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

If ev.State = MarketDataSourceStates.MarketDataSourceStateRunning Then
    getInitialTickerValues
ElseIf ev.State = MarketDataSourceStates.MarketDataSourceStateStopped Or _
    ev.State = MarketDataSourceStates.MarketDataSourceStateFinished _
Then
    ' the ticker has been stopped before the chart has been closed,
    ' so remove the chart from the config and close it
    MultiChart1.Finish
    If Not mConfig Is Nothing Then mConfig.Remove
    Set mConfig = Nothing
    Unload Me
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub ChartToolsToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Const ProcName As String = "ChartToolsToolbar_ButtonClick"
On Error GoTo Err

If MultiChart1.Count = 0 Then Exit Sub

Select Case Button.Key
Case ChartToolsCommandStudies
    gShowStudyPicker MultiChart1.ChartManager, _
                    mSymbol & _
                    " (" & MultiChart1.TimePeriod.ToString & ")", _
                    mOwner, _
                    mTheme
Case ChartToolsCommandSelection
    setSelectionMode
Case ChartToolsCommandLines
    createLineChartTool
Case ChartToolsCommandFib
    createFibChartTool
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
Const ProcName As String = "CoolBar1_HeightChanged"
On Error GoTo Err

resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub MultiChart1_Change(ev As ChangeEventData)
Const ProcName As String = "MultiChart1_Change"
On Error GoTo Err

Dim changeType As MultiChartChangeTypes
changeType = ev.changeType

Select Case changeType
Case MultiChartSelectionChanged
    If MultiChart1.Count > 0 Then
        ChartToolsToolbar.Enabled = True
        Set mChartController = MultiChart1.BaseChartController
        
        setCaption
        setSelectionButton
        syncStudyPicker
    Else
        setCaption
        ChartToolsToolbar.Enabled = False
        Set mChartController = Nothing
    End If
    Set mCurrentTool = Nothing
Case MultiChartAdd
    Dim lTitle As Text
    Set lTitle = MultiChart1.BaseChartController(MultiChart1.Count).XAxisRegion.Title
    lTitle.Box = False
    lTitle.Position = NewPoint(0.1, 0.1, CoordsDistance, CoordsCounterDistance)
    lTitle.FixedX = True
    lTitle.FixedY = True
    lTitle.align = TextAlignModes.AlignTopLeft
    lTitle.IncludeInAutoscale = False
    lTitle.PaddingX = 0#
    lTitle.Color = &H808080
    lTitle.Layer = LayerBackground + 1
    lTitle.Text = "© " & Year(Now) & " Copyright TradeWright Software Systems"
    Dim lFont As New StdFont
    lFont.Name = "Arial"
    lFont.Size = 7
    lTitle.Font = lFont
Case MultiChartRemove
    gUnsyncStudyPicker
Case MultiChartChangeTypes.MultiChartPeriodLengthChanged
    If MultiChart1.Count > 0 Then
        Set mChartController = MultiChart1.BaseChartController
    End If
    setCaption
    setSelectionButton
    syncStudyPicker
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub MultiChart1_ChartStateChanged(ByVal index As Long, ev As StateChangeEventData)
Const ProcName As String = "MultiChart1_ChartStateChanged"
On Error GoTo Err

Dim lChart As MarketChart
Set lChart = ev.Source

Dim LoadingText As Text

Select Case ev.State
Case ChartStateBlank

Case ChartStateCreated

Case ChartStateFetching
    Set LoadingText = lChart.LoadingText
    LoadingText.Box = True
    LoadingText.BoxFillStyle = FillTransparent
    LoadingText.BoxFillWithBackgroundColor = True
    LoadingText.BoxThickness = 1
    LoadingText.BoxStyle = LineInvisible
    LoadingText.Color = vbYellow
    LoadingText.Font.Size = 16
    LoadingText.Font.Italic = True
    LoadingText.align = AlignBottomCentre
    LoadingText.Position = NewPoint(50, 0.2, CoordsRelative, CoordsDistance)
    LoadingText.Text = "Fetching historical data"
Case ChartStateLoading
    Set LoadingText = lChart.LoadingText
    LoadingText.Color = vbGreen
    LoadingText.Text = "Loading historical data"
Case ChartStateRunning

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mChartController Event Handlers
'================================================================================

Private Sub mChartController_PointerModeChanged()
Const ProcName As String = "mChartController_PointerModeChanged"
On Error GoTo Err

setSelectionButton

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName

End Sub

'================================================================================
' mFutureWaiter Event Handlers
'================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub
Dim lContract As IContract
Set lContract = ev.Future.Value
mSecType = lContract.Specifier.secType
mSymbol = lContract.Specifier.LocalSymbol
mTickSize = lContract.TickSize
setCaption
If mIsHistorical And Not mConfig Is Nothing Then
    SaveContractToConfig lContract, mConfig.AddConfigurationSection(ConfigSectionContract)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get IsHistorical() As Boolean
IsHistorical = mIsHistorical
End Property

Public Property Let Owner(ByVal Value As Variant)
gSetVariant mOwner, Value
End Property

Public Property Let Style(ByVal Value As ChartStyle)
Const ProcName As String = "Style"
On Error GoTo Err

MultiChart1.Style = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

If mTheme Is Nothing Then
    If Value Is Nothing Then Exit Property
    DeferAction Me, Value, 1, ExpiryTimeUnitSeconds
Else
    Set mTheme = Value
    If mTheme Is Nothing Then Exit Property
        
    Me.BackColor = mTheme.BackColor
    gApplyTheme mTheme, Me.Controls
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pPeriodLength As TimePeriod, _
                ByVal pTimeframes As Timeframes, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle, _
                ByVal pOwner As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

Assert Not pPeriodLength Is Nothing, "pPeriodLength is nothing"
Assert Not pDataSource Is Nothing, "pDataSource is nothing"
Assert Not pTimeframes Is Nothing, "pTimeframes is nothing"
Assert Not pSpec Is Nothing, "pSpec is nothing"

gSetVariant mOwner, pOwner

Set mDataSource = pDataSource
getInitialTickerValues
mFutureWaiter.Add pDataSource.ContractFuture

mDataSource.AddGenericTickListener Me
mDataSource.AddStateChangeListener Me

Set mBarFormatterLibManager = pBarFormatterLibManager

mIsHistorical = False

Dim lExcludeLastBar As Boolean
lExcludeLastBar = mDataSource.IsTickReplay

MultiChart1.Initialise pTimeframes, pTimePeriodValidator, pSpec, pStyle, pBarFormatterLibManager, , , lExcludeLastBar
If Not pConfig Is Nothing Then setConfig pConfig

' we have to do something to cause Form_Load to run, otherwise MultiChart1 is
' not created for use below

MultiChart1.Enabled = True

ChartNavToolbar1.Initialise , MultiChart1.object
BarFormatterPicker.Initialise mBarFormatterLibManager, , MultiChart1.object
ChartStylePicker.Initialise , MultiChart1.object

MultiChart1.Add pPeriodLength

setCaption
mIsInitialised = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseHistoric( _
                ByVal pPeriodLength As TimePeriod, _
                ByVal pContractFuture As IFuture, _
                ByVal pStudyManager As StudyManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle, _
                ByVal pOwner As Variant)
Const ProcName As String = "InitialiseHistoric"
On Error GoTo Err

gSetVariant mOwner, pOwner

mFutureWaiter.Add pContractFuture

Set mBarFormatterLibManager = pBarFormatterLibManager

mIsHistorical = True

MultiChart1.Initialise createNewTimeframes(pStudyManager, pContractFuture, pHistDataStore), pHistDataStore.TimePeriodValidator, pSpec, pStyle, pBarFormatterLibManager
setConfig pConfig

' we have to do something to cause Form_Load to run, otherwise MultiChart1 is
' not created for use below

MultiChart1.Enabled = True

ChartNavToolbar1.Initialise , MultiChart1.object
BarFormatterPicker.Initialise mBarFormatterLibManager, , MultiChart1.object
ChartStylePicker.Initialise , MultiChart1.object

MultiChart1.Add pPeriodLength

setCaption
mIsInitialised = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function LoadFromConfig( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTimeframes As Timeframes, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pOwner As Variant) As Boolean
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Assert Not pDataSource Is Nothing, "pDataSource is nothing"
Assert Not pTimeframes Is Nothing, "pTimeframes is nothing"

gSetVariant mOwner, pOwner

Set mDataSource = pDataSource
getInitialTickerValues
mFutureWaiter.Add pDataSource.ContractFuture

mDataSource.AddGenericTickListener Me
mDataSource.AddStateChangeListener Me

Set mBarFormatterLibManager = pBarFormatterLibManager

' we have to do something to cause Form_Load to run, otherwise MultiChart1 is
' not created for use below

MultiChart1.Enabled = True

ChartNavToolbar1.Initialise , MultiChart1.object
BarFormatterPicker.Initialise mBarFormatterLibManager, , MultiChart1.object
ChartStylePicker.Initialise , MultiChart1.object

Set mConfig = pConfig
mIsHistorical = False
If Not MultiChart1.LoadFromConfig(mConfig.GetConfigurationSection(ConfigSectionMultiChart), pTimeframes, pTimePeriodValidator, pBarFormatterLibManager) Then
    LoadFromConfig = False
    Exit Function
End If

setWindow
mIsInitialised = True

LoadFromConfig = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function LoadHistoricFromConfig( _
                ByVal pContractFuture As IFuture, _
                ByVal pStudyManager As StudyManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pConfig As ConfigurationSection) As Boolean
Const ProcName As String = "LoadHistoricFromConfig"
On Error GoTo Err

mFutureWaiter.Add pContractFuture

Set mBarFormatterLibManager = pBarFormatterLibManager

' we have to do something to cause Form_Load to run, otherwise MultiChart1 is
' not created for use below

MultiChart1.Enabled = True

ChartNavToolbar1.Initialise , MultiChart1.object
BarFormatterPicker.Initialise mBarFormatterLibManager, , MultiChart1.object
ChartStylePicker.Initialise , MultiChart1.object

Set mConfig = pConfig
mIsHistorical = True
If Not MultiChart1.LoadFromConfig(mConfig.GetConfigurationSection(ConfigSectionMultiChart), createNewTimeframes(pStudyManager, pContractFuture, pHistDataStore), pHistDataStore.TimePeriodValidator, pBarFormatterLibManager) Then
    LoadHistoricFromConfig = False
    Exit Function
End If

setWindow
mIsInitialised = True

LoadHistoricFromConfig = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub createFibChartTool()
Const ProcName As String = "createFibChartTool"
On Error GoTo Err

Dim ls As LineStyle
Set ls = New LineStyle
ls.Extended = True
ls.IncludeInAutoscale = False
ls.Color = &H808080

Dim lineSpecs(4) As FibLineSpecifier
Set lineSpecs(0).Style = ls.Clone
lineSpecs(0).Percentage = 0

ls.Color = vbRed
Set lineSpecs(1).Style = ls.Clone
lineSpecs(1).Percentage = 100

ls.Color = &H8000&   ' dark green
Set lineSpecs(2).Style = ls.Clone
lineSpecs(2).Percentage = 50

ls.Color = vbBlue
Set lineSpecs(3).Style = ls.Clone
lineSpecs(3).Percentage = 38.2

ls.Color = vbMagenta
Set lineSpecs(4).Style = ls.Clone
lineSpecs(4).Percentage = 61.8

Set mCurrentTool = CreateFibRetracementTool(mChartController, lineSpecs, LayerNumbers.LayerHighestUser)
MultiChart1.SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createLineChartTool()
Const ProcName As String = "createLineChartTool"
On Error GoTo Err

Dim ls As LineStyle
Set ls = New LineStyle
ls.Extended = True
ls.ExtendAfter = True
ls.IncludeInAutoscale = False

Set mCurrentTool = CreateLineTool(mChartController, ls, LayerNumbers.LayerHighestUser)
MultiChart1.SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function createNewTimeframes( _
                ByVal pStudyManager As StudyManager, _
                ByVal pContractFuture As IFuture, _
                ByVal pHistDataStore As IHistoricalDataStore) As Timeframes
Const ProcName As String = "createNewTimeframes"
On Error GoTo Err

Dim lStudyBase As IStudyBase
Set lStudyBase = CreateStudyBaseForTickDataInput(pStudyManager, Nothing, pContractFuture)

Dim lTimeframes As Timeframes
Set lTimeframes = CreateTimeframes(lStudyBase, pContractFuture, pHistDataStore)

Set createNewTimeframes = lTimeframes

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFormattedPrice(ByVal pPrice As Double) As String
getFormattedPrice = FormatPrice(pPrice, mSecType, mTickSize)
End Function

Private Sub getInitialTickerValues()
Const ProcName As String = "getInitialTickerValues"
On Error GoTo Err

If mDataSource.State <> MarketDataSourceStates.MarketDataSourceStateRunning Then Exit Sub

mCurrentBid = getFormattedPrice(mDataSource.CurrentQuote(TickTypeBid).Price)
mCurrentTrade = getFormattedPrice(mDataSource.CurrentQuote(TickTypeTrade).Price)
mCurrentAsk = getFormattedPrice(mDataSource.CurrentQuote(TickTypeAsk).Price)
mCurrentVolume = CStr(mDataSource.CurrentQuote(TickTypeVolume).Size)
mCurrentHigh = getFormattedPrice(mDataSource.CurrentQuote(TickTypeHighPrice).Price)
mCurrentLow = getFormattedPrice(mDataSource.CurrentQuote(TickTypeLowPrice).Price)
mPreviousClose = getFormattedPrice(mDataSource.CurrentQuote(TickTypeClosePrice).Price)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub resize()
Const ProcName As String = "Resize"
On Error GoTo Err

If Me.WindowState = FormWindowStateConstants.vbMinimized Then Exit Sub

If Me.ScaleHeight >= CoolBar1.Height Then
    MultiChart1.Height = Me.ScaleHeight - CoolBar1.Height
    MultiChart1.Top = CoolBar1.Height
End If

updateSettings

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCaption()
Const ProcName As String = "setCaption"
On Error GoTo Err

Dim s As String
If MultiChart1.Count = 0 Then
    s = mSymbol
Else
    s = mSymbol & " (" & MultiChart1.TimePeriod.ToString & ")"
End If
    
If mIsHistorical Then
    s = s & _
        "    (historical)"
Else
    s = s & _
        "    B=" & mCurrentBid & _
        "  T=" & mCurrentTrade & _
        "  A=" & mCurrentAsk & _
        "  V=" & mCurrentVolume & _
        "  H=" & mCurrentHigh & _
        "  L=" & mCurrentLow & _
        "  C=" & mPreviousClose
End If
Me.Caption = s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "setConfig"
On Error GoTo Err

Set mConfig = pConfig
updateSettings
MultiChart1.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionMultiChart)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSelectionMode()
Const ProcName As String = "setSelectionMode"
On Error GoTo Err

If mChartController.PointerMode <> PointerModeSelection Then
    mChartController.SetPointerModeSelection
    ChartToolsToolbar.Buttons("selection").Value = tbrPressed
Else
    mChartController.SetPointerModeDefault
    ChartToolsToolbar.Buttons("selection").Value = tbrUnpressed
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSelectionButton()
Const ProcName As String = "setSelectionButton"
On Error GoTo Err

If mChartController.PointerMode = PointerModeSelection Then
    ChartToolsToolbar.Buttons("selection").Value = tbrPressed
Else
    ChartToolsToolbar.Buttons("selection").Value = tbrUnpressed
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setWindow()
Const ProcName As String = "setWindow"
On Error GoTo Err

Me.Width = CLng(mConfig.GetSetting(ConfigSettingWidth, Me.Width / Screen.TwipsPerPixelX)) * Screen.TwipsPerPixelX
Me.Height = CLng(mConfig.GetSetting(ConfigSettingHeight, Me.Height / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
Me.Left = CLng(mConfig.GetSetting(ConfigSettingLeft, Rnd * (Screen.Width - Me.Width) / Screen.TwipsPerPixelX)) * Screen.TwipsPerPixelX
Me.Top = CLng(mConfig.GetSetting(ConfigSettingTop, Rnd * (Screen.Height - Me.Height) / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY

Select Case mConfig.GetSetting(ConfigSettingWindowstate, WindowStateNormal)
Case WindowStateMaximized
    Me.WindowState = FormWindowStateConstants.vbMaximized
Case WindowStateMinimized
    Me.WindowState = FormWindowStateConstants.vbMinimized
Case WindowStateNormal
    Me.WindowState = FormWindowStateConstants.vbNormal
End Select

resize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub syncStudyPicker()
Const ProcName As String = "syncStudyPicker"
On Error GoTo Err

If MultiChart1.Count = 0 Then Exit Sub
gSyncStudyPicker MultiChart1.ChartManager, _
                "Study picker for " & mSymbol & _
                " (" & MultiChart1.TimePeriod.ToString & ")", _
                mOwner

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateSettings()
Const ProcName As String = "updateSettings"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub
If Not mIsInitialised Then Exit Sub

Select Case Me.WindowState
Case FormWindowStateConstants.vbMaximized
    mConfig.SetSetting ConfigSettingWindowstate, WindowStateMaximized
Case FormWindowStateConstants.vbMinimized
    mConfig.SetSetting ConfigSettingWindowstate, WindowStateMinimized
Case FormWindowStateConstants.vbNormal
    mConfig.SetSetting ConfigSettingWindowstate, WindowStateNormal
    mConfig.SetSetting ConfigSettingWidth, Me.Width / Screen.TwipsPerPixelX
    mConfig.SetSetting ConfigSettingHeight, Me.Height / Screen.TwipsPerPixelY
    mConfig.SetSetting ConfigSettingLeft, Me.Left / Screen.TwipsPerPixelX
    mConfig.SetSetting ConfigSettingTop, Me.Top / Screen.TwipsPerPixelY
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


