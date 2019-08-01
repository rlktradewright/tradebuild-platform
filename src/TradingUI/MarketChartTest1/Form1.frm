VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#345.0#0"; "TradingUI27.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#32.0#0"; "TWControls40.ocx"
Begin VB.Form Form1 
   Caption         =   "Market Chart Test1"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   14415
   StartUpPosition =   3  'Windows Default
   Begin TWControls40.TWImageCombo ChartStylesCombo 
      Height          =   330
      Left            =   1680
      TabIndex        =   11
      Top             =   6720
      Width           =   2295
      _ExtentX        =   4048
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
      MouseIcon       =   "Form1.frx":0000
      Text            =   ""
   End
   Begin VB.TextBox NumHistoryBarsText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Text            =   "500"
      Top             =   6120
      Width           =   975
   End
   Begin VB.CheckBox SessionOnlyCheck 
      Caption         =   "Session only"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   6360
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin TradingUI27.TimeframeSelector TimeframeSelector1 
      Height          =   270
      Left            =   1680
      TabIndex        =   7
      Top             =   5700
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   476
   End
   Begin TradingUI27.ChartNavToolbar ChartNavToolbar1 
      Height          =   330
      Left            =   4080
      TabIndex        =   6
      Top             =   480
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   582
   End
   Begin TradingUI27.ChartStylePicker ChartStylePicker1 
      Height          =   330
      Left            =   7680
      TabIndex        =   5
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
   End
   Begin TradingUI27.BarFormatterPicker BarFormatterPicker1 
      Height          =   270
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
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
   End
   Begin TradingUI27.MarketChart MarketChart1 
      Height          =   6375
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11245
   End
   Begin VB.TextBox LogText 
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7320
      Width           =   14175
   End
   Begin TradingUI27.ContractSearch ContractSearch 
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9657
      AllowMultipleSelection=   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "Form1.frx":001C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0176
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":05C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0722
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ChartToolsToolbar 
      Height          =   330
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
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
   Begin VB.Label Label22 
      Caption         =   "Number of history bars"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IDeferredAction
Implements ITwsConnectionStateListener
Implements ILogListener
Implements IStateChangeListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Enum DeferredActions
    DeferredActionSaveConfig = 1
End Enum

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "Form1"

Private Const ChartStyleNameAppDefault              As String = "Application default"
Private Const ChartStyleNameBlack                   As String = "Black"
Private Const ChartStyleNameDarkBlueFade            As String = "Dark blue fade"
Private Const ChartStyleNameGoldFade                As String = "Gold fade"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mUnhandledErrorHandler           As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1
Private mIsInDev                                    As Boolean

Private mClientId                                   As Long

Private mDataClient                                 As Client

Private mMarketDataManager                          As IMarketDataManager
Private mContractStore                              As IContractStore
Private mHistDataStore                              As IHistoricalDataStore

Private mNoLogfile                                  As Boolean

Private mPreferredGridRow                           As Long

Private WithEvents mConfigStore                     As ConfigurationStore
Attribute mConfigStore.VB_VarHelpID = -1
Private mMarketDataManagerConfig                    As ConfigurationSection

Private mTicker                                     As IMarketDataSource
Private WithEvents mTickers                         As EnumerableCollection
Attribute mTickers.VB_VarHelpID = -1

Private mTimePeriod                                 As TimePeriod

Private mStudyLibraryManager                        As StudyLibraryManager

Private mInputHandleBid                             As Long
Private mInputHandleAsk                             As Long
Private mInputHandleOpenInterest                    As Long
Private mInputHandleTickVolume                      As Long
Private mInputHandleTrade                           As Long
Private mInputHandleVolume                          As Long

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mBarFormatterLibManager                     As New BarFormatterLibManager

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Initialize()
Const ProcName As String = "Form_Initialize"
On Error GoTo Err

Debug.Print "Running in development environment: " & CStr(inDev)

Set mUnhandledErrorHandler = UnhandledErrorHandler

InitialiseTWUtilities

ApplicationGroupName = "TradeWright"
ApplicationName = "MarketChartTest1"
SetupDefaultLogging Command

Set mFutureWaiter = New FutureWaiter

Set mStudyLibraryManager = New StudyLibraryManager
mStudyLibraryManager.AddBuiltInStudyLibrary
mBarFormatterLibManager.AddBarFormatterLibrary "BarFormatters27.BarFormattersLib", True, "Built-in"
 
Set mConfigStore = getConfigStore
Set mMarketDataManagerConfig = mConfigStore.AddPrivateConfigurationSection("/MarketDataManager")
If mConfigStore.Dirty Then mConfigStore.Save

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrSecurityException Then
    mNoLogfile = True
    Resume Next
End If
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

GetLogger("log").AddLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box
If mNoLogfile Then
    LogMessage "Can't open the log file: " & DefaultLogFileName(Command)
Else
    LogMessage "Logging to file: " & DefaultLogFileName(Command)
End If

mClientId = 1413860445
Set mDataClient = GetClient("Essy", 7497, mClientId, , , ApiMessageLoggingOptionDefault, ApiMessageLoggingOptionNone, False, , Me)
mDataClient.SetTwsLogLevel TwsLogLevelDetail

Set mContractStore = mDataClient.GetContractStore
Set mMarketDataManager = CreateRealtimeDataManager(mDataClient.GetMarketDataFactory, mStudyLibraryManager)
Set mTickers = mMarketDataManager.DataSources

Set mHistDataStore = mDataClient.GetHistoricalDataStore

ContractSearch.Initialise mContractStore, Nothing

TimeframeSelector1.Initialise mHistDataStore.TimePeriodValidator
TimeframeSelector1.SelectTimeframe GetTimePeriod(5, TimePeriodMinute)

setupChartStyles ChartStylesCombo.ComboItems
ChartStylesCombo.ComboItems.Item(ChartStyleNameAppDefault).Selected = True

mMarketDataManager.LoadFromConfig mMarketDataManagerConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If Not mDataClient Is Nothing Then mDataClient.Finish
TerminateTWUtilities
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

Select Case CLng(Data)
Case DeferredActions.DeferredActionSaveConfig
    If mConfigStore.Dirty Then mConfigStore.Save
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

If mTicker.State = MarketDataSourceStateRunning Then
    mTicker.RemoveStateChangeListener Me
    showTheChart mTicker, _
            CreateChartSpecifier(CLng(NumHistoryBarsText.Text), Not (SessionOnlyCheck = vbChecked)), _
            ChartStylesManager.Item(ChartStylesCombo.SelectedItem.Text)
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As ApiConnectionStates, ByVal pMessage As String)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionStateChange"
On Error GoTo Err

Select Case pState
Case ApiConnNotConnected
    LogMessage "Disconnected from TWS: " & pMessage
Case ApiConnConnecting
    LogMessage "Connecting to TWS: " & pMessage
Case ApiConnConnected
    LogMessage "Connected to TWS: " & pMessage
Case ApiConnFailed
    LogMessage "Failed to connect to TWS: " & pMessage
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)

End Sub

'@================================================================================
' ILogListener Interface Members
'@================================================================================

Private Sub ILogListener_Finish()

End Sub

Private Sub ILogListener_Notify(ByVal Logrec As LogRecord)
Const ProcName As String = "ILogListener_Notify"
On Error GoTo Err

If Len(LogText.Text) >= 32767 Then
    ' clear some space at the start of the textbox
    LogText.SelStart = 0
    LogText.SelLength = 16384
    LogText.SelText = ""
End If

LogText.SelStart = Len(LogText.Text)
LogText.SelLength = 0
If Len(LogText.Text) > 0 Then LogText.SelText = vbCrLf
LogText.SelText = formatLogRecord(Logrec)
LogText.SelStart = InStrRev(LogText.Text, vbCrLf) + 2

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ContractSearch_Action()
Const ProcName As String = "ContractSearch_Action"
On Error GoTo Err

mMarketDataManager.CreateMarketDataSource CreateFuture(ContractSearch.SelectedContracts.ItemAtIndex(1)), True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mConfigStore Event Handlers
'@================================================================================

Private Sub mConfigStore_Change(ev As ChangeEventData)
Const ProcName As String = "mConfigStore_Change"
On Error GoTo Err

If ev.ChangeType = ConfigChangeTypes.ConfigDirty Then DeferAction Me, DeferredActions.DeferredActionSaveConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

setCaption ev.Future.Value

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTickers Event Handlers
'@================================================================================

Private Sub mTickers_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mTickers_CollectionChanged"
On Error GoTo Err

If ev.ChangeType <> CollItemAdded Then Exit Sub

If Not mTicker Is Nothing Then
    mTicker.Finish
    mTicker.RemoveFromConfig
End If

Set mTicker = ev.AffectedItem
mTicker.AddStateChangeListener Me
mTicker.StartMarketData

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mUnhandledErrorHandler Event Handlers
'@================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As ErrorEventData)

If Not mDataClient Is Nothing Then mDataClient.Finish

handleFatalError

' Tell TWUtilities that we've now handled this unhandled error. Not actually
' needed here because HandleFatalError never returns anyway
UnhandledErrorHandler.Handled = True
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Function formatLogRecord(ByVal Logrec As LogRecord) As String
Const ProcName As String = "formatLogRecord"
Static formatter As ILogFormatter

On Error GoTo Err

If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter(TimestampFormats.TimestampTimeOnlyLocal)
formatLogRecord = formatter.FormatRecord(Logrec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getConfigStore() As ConfigurationStore
Const ProcName As String = "getConfigStore"
On Error GoTo Err

Set getConfigStore = GetDefaultConfigurationStore(Command, "1.0", False, ConfigFileOptionFirstArg)
If getConfigStore Is Nothing Then Set getConfigStore = GetDefaultConfigurationStore(Command, "1.0", True, ConfigFileOptionFirstArg)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

If Not mDataClient Is Nothing Then mDataClient.Finish

MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please email the log file located at" & vbCrLf & vbCrLf & _
        "     " & DefaultLogFileName(Command) & vbCrLf & vbCrLf & _
        "to support@tradewright.com", _
        vbCritical, _
        "Fatal error"

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly.
'
' Note that normally one would use the End statement to terminate a VB6 program abruptly. But
' the TWUtilities component interferes with the End statement's processing and may prevent
' proper shutdown, so we use the TWUtilities component's EndProcess method instead.
'
' However if we are running in the development environment, then we call End because the
' EndProcess method kills the entire development environment as well which can have undesirable
' side effects if other components are also loaded.

If mIsInDev Then
    End
Else
    EndProcess
End If

End Sub

Private Function inDev() As Boolean
mIsInDev = True
inDev = True
End Function

Private Sub setCaption(ByVal pContract As IContract)
Const ProcName As String = "setCaption"
On Error GoTo Err

Me.Caption = pContract.Specifier.LocalSymbol & " (" & mTimePeriod.ToString & ")"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupChartStyles(ByVal pComboItems As ComboItems)
Const ProcName As String = "setupChartStyles"
On Error GoTo Err

setupChartStyleAppDefault
pComboItems.Add , ChartStyleNameAppDefault, ChartStyleNameAppDefault

setupChartStyleBlack
pComboItems.Add , ChartStyleNameBlack, ChartStyleNameBlack

setupChartStyleDarkBlueFade
pComboItems.Add , ChartStyleNameDarkBlueFade, ChartStyleNameDarkBlueFade

setupChartStyleGoldFade
pComboItems.Add , ChartStyleNameGoldFade, ChartStyleNameGoldFade

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupChartStyleAppDefault()
Const ProcName As String = "setupChartStyleAppDefault"
Dim lDefaultRegionStyle As ChartRegionStyle
Dim lxAxisRegionStyle As ChartRegionStyle
Dim lDefaultYAxisRegionStyle As ChartRegionStyle
Dim lCrosshairLineStyle As LineStyle
Dim lCursorTextStyle As TextStyle
Dim lFont As StdFont

On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameAppDefault) Then Exit Sub

ReDim GradientFillColors(1) As Long

Set lCursorTextStyle = New TextStyle
lCursorTextStyle.Align = AlignBoxTopCentre
lCursorTextStyle.Box = True
lCursorTextStyle.BoxFillWithBackgroundColor = True
lCursorTextStyle.BoxStyle = LineInvisible
lCursorTextStyle.BoxThickness = 0
lCursorTextStyle.Color = &H80&
lCursorTextStyle.PaddingX = 2
lCursorTextStyle.PaddingY = 0
Set lFont = New StdFont
lFont.Name = "Courier New"
lFont.Bold = True
lFont.Size = 8
lCursorTextStyle.Font = lFont

Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone
GradientFillColors(0) = RGB(192, 192, 192)
GradientFillColors(1) = RGB(248, 248, 248)
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors
    
Set lxAxisRegionStyle = GetDefaultChartXAxisRegionStyle.Clone
lxAxisRegionStyle.XCursorTextStyle = lCursorTextStyle
GradientFillColors(0) = RGB(230, 236, 207)
GradientFillColors(1) = RGB(222, 236, 215)
lxAxisRegionStyle.BackGradientFillColors = GradientFillColors
    
Set lDefaultYAxisRegionStyle = GetDefaultChartYAxisRegionStyle.Clone
lDefaultYAxisRegionStyle.YCursorTextStyle = lCursorTextStyle
GradientFillColors(0) = RGB(234, 246, 254)
GradientFillColors(1) = RGB(226, 246, 255)
lDefaultYAxisRegionStyle.BackGradientFillColors = GradientFillColors
    
Set lCrosshairLineStyle = New LineStyle
lCrosshairLineStyle.Color = &H7F

ChartStylesManager.Add ChartStyleNameAppDefault, _
                        ChartStylesManager.DefaultStyle, _
                        lDefaultRegionStyle, _
                        lxAxisRegionStyle, _
                        lDefaultYAxisRegionStyle, _
                        lCrosshairLineStyle


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupChartStyleBlack()
Const ProcName As String = "setupChartStyleBlack"
Dim lDefaultRegionStyle As ChartRegionStyle
Dim lxAxisRegionStyle As ChartRegionStyle
Dim lDefaultYAxisRegionStyle As ChartRegionStyle
Dim lCrosshairLineStyle As LineStyle
Dim lCursorTextStyle As TextStyle
Dim lFont As StdFont
Dim lGridLineStyle As LineStyle
Dim lGridTextStyle As TextStyle

On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameBlack) Then Exit Sub

ReDim GradientFillColors(1) As Long

Set lCursorTextStyle = New TextStyle
lCursorTextStyle.Align = AlignBoxTopCentre
lCursorTextStyle.Box = True
lCursorTextStyle.BoxFillWithBackgroundColor = True
lCursorTextStyle.BoxStyle = LineInvisible
lCursorTextStyle.BoxThickness = 0
lCursorTextStyle.Color = vbRed
lCursorTextStyle.PaddingX = 2
lCursorTextStyle.PaddingY = 0
Set lFont = New StdFont
lFont.Name = "Courier New"
lFont.Bold = True
lFont.Size = 8
lCursorTextStyle.Font = lFont

Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone
GradientFillColors(0) = RGB(0, 0, 0)
GradientFillColors(1) = RGB(0, 0, 0)
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors

Set lGridLineStyle = New LineStyle
lGridLineStyle.Color = RGB(64, 64, 64)
lDefaultRegionStyle.XGridLineStyle = lGridLineStyle
lDefaultRegionStyle.YGridLineStyle = lGridLineStyle
    
Set lGridLineStyle = New LineStyle
lGridLineStyle.Color = RGB(64, 64, 64)
lGridLineStyle.LineStyle = LineDash
lDefaultRegionStyle.SessionEndGridLineStyle = lGridLineStyle
    
Set lGridLineStyle = New LineStyle
lGridLineStyle.Color = RGB(64, 64, 64)
lGridLineStyle.Thickness = 3
lDefaultRegionStyle.SessionStartGridLineStyle = lGridLineStyle

Set lxAxisRegionStyle = GetDefaultChartXAxisRegionStyle.Clone
GradientFillColors(0) = RGB(0, 0, 0)
GradientFillColors(1) = RGB(0, 0, 0)
lxAxisRegionStyle.BackGradientFillColors = GradientFillColors
lxAxisRegionStyle.XCursorTextStyle = lCursorTextStyle

Set lGridTextStyle = New TextStyle
lGridTextStyle.Box = True
lGridTextStyle.BoxFillWithBackgroundColor = True
lGridTextStyle.BoxStyle = LineInvisible
lGridTextStyle.Color = &HD0D0D0
lxAxisRegionStyle.XGridTextStyle = lGridTextStyle
    
Set lDefaultYAxisRegionStyle = GetDefaultChartYAxisRegionStyle.Clone
GradientFillColors(0) = RGB(0, 0, 0)
GradientFillColors(1) = RGB(0, 0, 0)
lDefaultYAxisRegionStyle.BackGradientFillColors = GradientFillColors
lDefaultYAxisRegionStyle.YCursorTextStyle = lCursorTextStyle
lDefaultYAxisRegionStyle.YGridTextStyle = lGridTextStyle
    
Set lCrosshairLineStyle = New LineStyle
lCrosshairLineStyle.Color = vbRed

ChartStylesManager.Add ChartStyleNameBlack, _
                        ChartStylesManager.Item(ChartStyleNameAppDefault), _
                        lDefaultRegionStyle, _
                        lxAxisRegionStyle, _
                        lDefaultYAxisRegionStyle, _
                        lCrosshairLineStyle


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupChartStyleDarkBlueFade()
Const ProcName As String = "setupChartStyleDarkBlueFade"
Dim lDefaultRegionStyle As ChartRegionStyle
Dim lCrosshairLineStyle As LineStyle
Dim lFont As StdFont
Dim lGridLineStyle As LineStyle

On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameDarkBlueFade) Then Exit Sub

ReDim GradientFillColors(1) As Long

Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone
GradientFillColors(0) = &H643232
GradientFillColors(1) = &HF8F8F8
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors
    
Set lGridLineStyle = New LineStyle
lGridLineStyle.Color = &HC0C0C0
lDefaultRegionStyle.XGridLineStyle = lGridLineStyle
lDefaultRegionStyle.YGridLineStyle = lGridLineStyle
    
Set lGridLineStyle = New LineStyle
lGridLineStyle.Color = &HC0C0C0
lGridLineStyle.LineStyle = LineDash
lDefaultRegionStyle.SessionEndGridLineStyle = lGridLineStyle
    
Set lGridLineStyle = New LineStyle
lGridLineStyle.Color = &HC0C0C0
lGridLineStyle.Thickness = 3
lDefaultRegionStyle.SessionStartGridLineStyle = lGridLineStyle

Set lCrosshairLineStyle = New LineStyle
lCrosshairLineStyle.Color = vbRed

ChartStylesManager.Add ChartStyleNameDarkBlueFade, _
                        ChartStylesManager.Item(ChartStyleNameAppDefault), _
                        lDefaultRegionStyle, _
                        , _
                        , _
                        lCrosshairLineStyle


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupChartStyleGoldFade()
Const ProcName As String = "setupChartStyleGoldFade"
Dim lDefaultRegionStyle As ChartRegionStyle
Dim lCrosshairLineStyle As LineStyle
Dim lCursorTextStyle As TextStyle
Dim lFont As StdFont
Dim lGridLineStyle As LineStyle

On Error GoTo Err

If ChartStylesManager.Contains(ChartStyleNameGoldFade) Then Exit Sub

ReDim GradientFillColors(1) As Long

Set lCursorTextStyle = New TextStyle
lCursorTextStyle.Align = AlignBoxTopCentre
lCursorTextStyle.Box = True
lCursorTextStyle.BoxFillWithBackgroundColor = True
lCursorTextStyle.BoxStyle = LineInvisible
lCursorTextStyle.BoxThickness = 0
lCursorTextStyle.Color = &H80&
lCursorTextStyle.PaddingX = 2
lCursorTextStyle.PaddingY = 0
Set lFont = New StdFont
lFont.Name = "Courier New"
lFont.Bold = True
lFont.Size = 8
lCursorTextStyle.Font = lFont

Set lDefaultRegionStyle = GetDefaultChartDataRegionStyle.Clone
GradientFillColors(0) = &H82DFE6
GradientFillColors(1) = &HEBFAFB
lDefaultRegionStyle.BackGradientFillColors = GradientFillColors
    
Set lGridLineStyle = New LineStyle
lGridLineStyle.Color = &HC0C0C0
lDefaultRegionStyle.XGridLineStyle = lGridLineStyle
lDefaultRegionStyle.YGridLineStyle = lGridLineStyle
    
Set lGridLineStyle = New LineStyle
lGridLineStyle.Color = &HC0C0C0
lGridLineStyle.LineStyle = LineDash
lDefaultRegionStyle.SessionEndGridLineStyle = lGridLineStyle
    
Set lGridLineStyle = New LineStyle
lGridLineStyle.Color = &HC0C0C0
lGridLineStyle.Thickness = 3
lDefaultRegionStyle.SessionStartGridLineStyle = lGridLineStyle

Set lCrosshairLineStyle = New LineStyle
lCrosshairLineStyle.Color = 127

ChartStylesManager.Add ChartStyleNameGoldFade, _
                        ChartStylesManager.Item(ChartStyleNameAppDefault), _
                        lDefaultRegionStyle, _
                        , _
                        , _
                        lCrosshairLineStyle


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showTheChart( _
                ByVal pTicker As IMarketDataSource, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle)
Const ProcName As String = "showTheChart"
On Error GoTo Err

Set mTimePeriod = TimeframeSelector1.TimePeriod
MarketChart1.ShowChart CreateTimeframes(mTicker.StudyBase, mTicker.ContractFuture, mHistDataStore, mTicker.ClockFuture), mTimePeriod, pSpec, pStyle, True, mBarFormatterLibManager

ChartNavToolbar1.Initialise MarketChart1
BarFormatterPicker1.Initialise mBarFormatterLibManager, MarketChart1
ChartStylePicker1.Initialise MarketChart1

Exit Sub

Err:
Set mTicker = Nothing
gHandleUnexpectedError ProcName, ModuleName
End Sub





