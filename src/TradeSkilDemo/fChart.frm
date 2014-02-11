VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#203.0#0"; "TradingUI27.ocx"
Begin VB.Form fChart 
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin TradingUI27.MultiChart MultiChart1 
      Align           =   1  'Align Top
      Height          =   5415
      Left            =   0
      TabIndex        =   3
      Top             =   330
      Width           =   12525
      _ExtentX        =   22093
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
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   582
      BandCount       =   4
      BackColor       =   -2147483638
      _CBWidth        =   12525
      _CBHeight       =   330
      _Version        =   "6.7.9816"
      Child1          =   "ChartToolsToolbar"
      MinWidth1       =   1890
      MinHeight1      =   330
      Width1          =   1890
      NewRow1         =   0   'False
      Child2          =   "BarFormatterPicker"
      MinWidth2       =   1185
      MinHeight2      =   330
      Width2          =   1185
      NewRow2         =   0   'False
      Child3          =   "ChartStylePicker"
      MinWidth3       =   1185
      MinHeight3      =   330
      Width3          =   1185
      NewRow3         =   0   'False
      Child4          =   "ChartNavToolbar1"
      MinWidth4       =   6465
      MinHeight4      =   330
      Width4          =   6465
      NewRow4         =   0   'False
      Begin TradingUI27.ChartStylePicker ChartStylePicker 
         Height          =   330
         Left            =   3795
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
         Height          =   330
         Left            =   2340
         TabIndex        =   4
         ToolTipText     =   "Change the bar formatting"
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
      Begin TradingUI27.ChartNavToolbar ChartNavToolbar1 
         Height          =   330
         Left            =   5250
         TabIndex        =   2
         Top             =   0
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   582
      End
      Begin MSComctlLib.Toolbar ChartToolsToolbar 
         Height          =   330
         Left            =   180
         TabIndex        =   1
         Top             =   0
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "studies"
               Object.ToolTipText     =   "Manage the studies displayed on the chart"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "selection"
               Description     =   "Select a chart object"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "lines"
               Object.ToolTipText     =   "Draw lines"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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

Implements IGenericTickListener
Implements StateChangeListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                        As String = "fChart"

Private Const ChartToolsCommandStudies          As String = "studies"
Private Const ChartToolsCommandSelection        As String = "selection"
Private Const ChartToolsCommandLines            As String = "lines"
Private Const ChartToolsCommandFib              As String = "fib"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTicker                                 As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private mSymbol                                 As String
Private mSecType                                As SecurityTypes
Private mTicksize                               As Double

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

Private mTimePeriodValidator                    As ITimePeriodValidator

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

Resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
Const ProcName As String = "Form_QueryUnload"
On Error GoTo Err

If Not mTicker Is Nothing Then
    mTicker.RemoveGenericTickListener Me
    mTicker.RemoveStateChangeListener Me
End If

MultiChart1.Finish
If mIsHistorical Then
    If Not mTicker Is Nothing Then mTicker.Finish
    Set mTicker = Nothing
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

Resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Form_Terminate()
Const ProcName As String = "Form_Terminate"
LogMessage "Chart form terminated", LogLevelDetail
End Sub

'================================================================================
' IGenericTickListener Interface Members
'================================================================================

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

'================================================================================
' StateChangeListener Interface Members
'================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

If ev.State = MarketDataSourceStates.MarketDataSourceStateRunning Then
    getInitialTickerValues
ElseIf ev.State = MarketDataSourceStates.MarketDataSourceStateStopped Then
    ' the ticker has been stopped before the chart has been closed,
    ' so remove the chart from the config and close it
    MultiChart1.Finish
    mConfig.Remove
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
                    " (" & MultiChart1.TimePeriod.ToString & ")"
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

Resize

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
    Set lTitle = MultiChart1.BaseChartController(MultiChart1.Count).XAxisRegion.title
    lTitle.Box = False
    lTitle.Position = NewPoint(0.1, 0.1, CoordsDistance, CoordsCounterDistance)
    lTitle.FixedX = True
    lTitle.FixedY = True
    lTitle.Align = TextAlignModes.AlignTopLeft
    lTitle.IncludeInAutoscale = False
    lTitle.PaddingX = 0#
    lTitle.Color = &H808080
    lTitle.layer = LayerBackground + 1
    lTitle.Text = "© " & Year(Now) & " Copyright TradeWright Software Systems"
    Dim lFont As New StdFont
    lFont.name = "Arial"
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

Dim loadingText As Text
Select Case ev.State
Case ChartStateBlank

Case ChartStateCreated

Case ChartStateInitialised
    Set loadingText = MultiChart1.loadingText(index)
    loadingText.Box = True
    loadingText.BoxFillWithBackgroundColor = True
    loadingText.BoxThickness = 1
    loadingText.BoxStyle = LineInvisible
    loadingText.Color = vbYellow
    loadingText.Font.Size = 16
    loadingText.Font.Italic = True
    loadingText.Align = AlignBottomCentre
    loadingText.Position = NewPoint(50, 0.2, CoordsRelative, CoordsDistance)
    loadingText.Text = "Fetching historical data"
Case ChartStateLoading
    Set loadingText = MultiChart1.loadingText(index)
    loadingText.Color = vbGreen
    loadingText.Text = "Loading historical data"
Case ChartStateLoaded

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
Set lContract = ev.Future.value
mSecType = lContract.Specifier.SecType
mSymbol = lContract.Specifier.LocalSymbol
mTicksize = lContract.TickSize
setCaption

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

Public Property Let Style(ByVal value As ChartStyle)
MultiChart1.Style = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise( _
                ByVal pTicker As Ticker, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pAppInstanceConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTicker = pTicker
mTicker.AddStateChangeListener Me
mTicker.AddGenericTickListener Me
mFutureWaiter.Add mTicker.ContractFuture

Set mBarFormatterLibManager = pBarFormatterLibManager

mIsHistorical = (pSpec.toTime <> CDate(0))

MultiChart1.Initialise mTicker.Timeframes, pTimePeriodValidator, pSpec, pStyle, pBarFormatterLibManager, , , mTicker.IsTickReplay
If Not mTicker.IsTickReplay Then setConfig pAppInstanceConfig

ChartNavToolbar1.Initialise , MultiChart1
BarFormatterPicker.Initialise mBarFormatterLibManager, , MultiChart1
ChartStylePicker.Initialise , MultiChart1

getInitialTickerValues

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function LoadFromConfig( _
                ByVal pTicker As Ticker, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal config As ConfigurationSection) As Boolean
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mTicker = pTicker
mTicker.AddStateChangeListener Me
mTicker.AddGenericTickListener Me
mFutureWaiter.Add mTicker.ContractFuture

Set mBarFormatterLibManager = pBarFormatterLibManager

ChartNavToolbar1.Initialise , MultiChart1
BarFormatterPicker.Initialise mBarFormatterLibManager, , MultiChart1
ChartStylePicker.Initialise , MultiChart1

Set mConfig = config
mIsHistorical = CBool(mConfig.GetSetting(ConfigSettingHistorical, "False"))
If Not MultiChart1.LoadFromConfig(mConfig.GetConfigurationSection(ConfigSectionMultiChart), mTicker.Timeframes, pTimePeriodValidator, pBarFormatterLibManager) Then
    LoadFromConfig = False
    Exit Function
End If

setWindow

LoadFromConfig = True

Exit Function

Err:
Set mTicker = Nothing
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ShowChart( _
                ByVal pPeriodLength As TimePeriod)
Const ProcName As String = "showChart"
On Error GoTo Err

MultiChart1.Add pPeriodLength

setCaption

Exit Sub

Err:
Set mTicker = Nothing
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

Private Function getFormattedPrice(ByVal pPrice As Double) As String
getFormattedPrice = FormatPrice(pPrice, mSecType, mTicksize)
End Function

Private Sub getInitialTickerValues()
Const ProcName As String = "getInitialTickerValues"
On Error GoTo Err

If mTicker.State <> MarketDataSourceStates.MarketDataSourceStateRunning Then Exit Sub

mCurrentBid = getFormattedPrice(mTicker.CurrentQuote(TickTypeBid).Price)
mCurrentTrade = getFormattedPrice(mTicker.CurrentQuote(TickTypeTrade).Price)
mCurrentAsk = getFormattedPrice(mTicker.CurrentQuote(TickTypeAsk).Price)
mCurrentVolume = CStr(mTicker.CurrentQuote(TickTypeVolume).Size)
mCurrentHigh = getFormattedPrice(mTicker.CurrentQuote(TickTypeHighPrice).Price)
mCurrentLow = getFormattedPrice(mTicker.CurrentQuote(TickTypeLowPrice).Price)
mPreviousClose = getFormattedPrice(mTicker.CurrentQuote(TickTypeClosePrice).Price)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Resize()
Const ProcName As String = "Resize"
On Error GoTo Err

If Me.WindowState = FormWindowStateConstants.vbMinimized Then Exit Sub

If Me.ScaleHeight >= CoolBar1.Height Then
    MultiChart1.Height = Me.ScaleHeight - CoolBar1.Height
    MultiChart1.Top = CoolBar1.Height
End If

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
Me.caption = s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setConfig(ByVal pAppInstanceConfig As ConfigurationSection)
Const ProcName As String = "setConfig"
On Error GoTo Err

Set mConfig = pAppInstanceConfig.GetConfigurationSection(ConfigSectionCharts).AddConfigurationSection(ConfigSectionChart & "(" & mTicker.Key & ")")
mConfig.SetSetting ConfigSettingHistorical, CStr(mIsHistorical)
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
    ChartToolsToolbar.buttons("selection").value = tbrPressed
Else
    mChartController.SetPointerModeDefault
    ChartToolsToolbar.buttons("selection").value = tbrUnpressed
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSelectionButton()
Const ProcName As String = "setSelectionButton"
On Error GoTo Err

If mChartController.PointerMode = PointerModeSelection Then
    ChartToolsToolbar.buttons("selection").value = tbrPressed
Else
    ChartToolsToolbar.buttons("selection").value = tbrUnpressed
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
Me.left = CLng(mConfig.GetSetting(ConfigSettingLeft, Rnd * (Screen.Width - Me.Width) / Screen.TwipsPerPixelX)) * Screen.TwipsPerPixelX
Me.Top = CLng(mConfig.GetSetting(ConfigSettingTop, Rnd * (Screen.Height - Me.Height) / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY

Select Case mConfig.GetSetting(ConfigSettingWindowstate, WindowStateNormal)
Case WindowStateMaximized
    Me.WindowState = FormWindowStateConstants.vbMaximized
Case WindowStateMinimized
    Me.WindowState = FormWindowStateConstants.vbMinimized
Case WindowStateNormal
    Me.WindowState = FormWindowStateConstants.vbNormal
End Select

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
                " (" & MultiChart1.TimePeriod.ToString & ")"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateSettings()
Const ProcName As String = "updateSettings"
On Error GoTo Err

Select Case Me.WindowState
Case FormWindowStateConstants.vbMaximized
    mConfig.SetSetting ConfigSettingWindowstate, WindowStateMaximized
Case FormWindowStateConstants.vbMinimized
    mConfig.SetSetting ConfigSettingWindowstate, WindowStateMinimized
Case FormWindowStateConstants.vbNormal
    mConfig.SetSetting ConfigSettingWindowstate, WindowStateNormal
    mConfig.SetSetting ConfigSettingWidth, Me.Width / Screen.TwipsPerPixelX
    mConfig.SetSetting ConfigSettingHeight, Me.Height / Screen.TwipsPerPixelY
    mConfig.SetSetting ConfigSettingLeft, Me.left / Screen.TwipsPerPixelX
    mConfig.SetSetting ConfigSettingTop, Me.Top / Screen.TwipsPerPixelY
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


