VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#292.0#0"; "TradingUI27.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ClosePriceText 
      Height          =   285
      Index           =   1
      Left            =   10920
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox OpenPriceText 
      Height          =   285
      Index           =   1
      Left            =   9720
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox ClosePriceText 
      Height          =   285
      Index           =   0
      Left            =   10920
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox OpenPriceText 
      Height          =   285
      Index           =   0
      Left            =   9720
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CheckBox AddListenerCheck 
      Caption         =   "Add listener"
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   25
      Top             =   600
      Width           =   1815
   End
   Begin VB.CheckBox AddListenerCheck 
      Caption         =   "Add listener"
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   24
      Top             =   120
      Width           =   1815
   End
   Begin TradingUI27.ContractSpecBuilder ContractSpecBuilder1 
      Height          =   3690
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5556
      ForeColor       =   -2147483640
      ModeAdvanced    =   -1  'True
   End
   Begin VB.TextBox BidSizeText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox BidPriceText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox TradePriceText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox AskPriceText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox AskSizeText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox TradeSizeText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox VolumeText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox HighText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox LowText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox LowText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox HighText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox VolumeText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox TradeSizeText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox AskSizeText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox AskPriceText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox TradePriceText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox BidPriceText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox BidSizeText 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton StopTickerButton 
      Caption         =   "Stop ticker 2"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton StopTickerButton 
      Caption         =   "Stop ticker 1"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton StartTickerButton 
      Caption         =   "Start ticker &2"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton StartTickerButton 
      Caption         =   "Start ticker &1"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox LogText 
      Height          =   4095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5400
      Width           =   9495
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

Implements IErrorListener
Implements IGenericTickListener
Implements ILogListener
Implements INotificationListener
Implements ITwsConnectionStateListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "Form1"

Private Const ClientId                              As Long = 1132256741
Private Const Server                                As String = "Sven"

Private Const DataSourceKey0                        As String = "Source0"
Private Const DataSourceKey1                        As String = "Source1"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mUnhandledErrorHandler           As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1
Private mIsInDev                                    As Boolean

Private mResultCount                                As Long

Private mClient                                     As Client
Attribute mClient.VB_VarHelpID = -1
Private mDataManager                                As IMarketDataManager

Private mMarketDataSource(1)                        As IMarketDataSource

Private mContractStore                              As IContractStore

Private mMarketDataListener(1)                      As MarketDataListener

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Initialize()
Debug.Print "Running in development environment: " & CStr(inDev)
InitialiseTWUtilities
Set mUnhandledErrorHandler = UnhandledErrorHandler

DefaultLogLevel = LogLevelHighDetail

ApplicationGroupName = "TradeWright"
ApplicationName = "MarketDataTest1"
SetupDefaultLogging Command
GetLogger("log").AddLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box

Set mFutureWaiter = New FutureWaiter

Set mClient = GetClient(Server, 7497, ClientId, , , , Me, , Me, Me)

Set mDataManager = CreateRealtimeDataManager(mClient.GetMarketDataFactory)
Set mContractStore = mClient.GetContractStore

End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

Private Sub Form_Unload(Cancel As Integer)
mClient.Finish
End Sub

'@================================================================================
' IErrorListener Interface Members
'@================================================================================

Private Sub IErrorListener_Notify(ev As ErrorEventData)
Dim lIndex As Long: lIndex = getDataSourceIndex(ev.Source)

StartTickerButton(lIndex).Enabled = True
StopTickerButton(lIndex).Enabled = False
LogMessage "Error " & ev.ErrorCode & ": " & ev.ErrorMessage & vbCrLf & _
            "At: " & ev.ErrorSource
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Dim lDataSource         As IMarketDataSource:   Set lDataSource = ev.Source
Dim lIndex              As Long:                lIndex = getDataSourceIndex(lDataSource)
Dim lContract           As IContract:           Set lContract = lDataSource.ContractFuture.Value
Dim lSecType            As SecurityTypes:       lSecType = lContract.Specifier.SecType
Dim lTicksize           As Double:              lTicksize = lContract.TickSize

With ev.Tick
    Select Case ev.Tick.TickType
    Case TickTypeBid
        BidPriceText(lIndex) = FormatPrice(ev.Tick.Price, lSecType, lTicksize)
        BidSizeText(lIndex) = ev.Tick.Size
    Case TickTypeAsk
        AskPriceText(lIndex) = FormatPrice(ev.Tick.Price, lSecType, lTicksize)
        AskSizeText(lIndex) = ev.Tick.Size
    Case TickTypeClosePrice
        ClosePriceText(lIndex) = FormatPrice(ev.Tick.Price, lSecType, lTicksize)
    Case TickTypeHighPrice
        HighText(lIndex) = FormatPrice(ev.Tick.Price, lSecType, lTicksize)
    Case TickTypeLowPrice
        LowText(lIndex) = FormatPrice(ev.Tick.Price, lSecType, lTicksize)
    Case TickTypeMarketDepth
    
    Case TickTypeMarketDepthReset
    
    Case TickTypeTrade
        TradePriceText(lIndex) = FormatPrice(ev.Tick.Price, lSecType, lTicksize)
        TradeSizeText(lIndex) = ev.Tick.Size
    Case TickTypeVolume
        VolumeText(lIndex) = ev.Tick.Size
    Case TickTypeOpenInterest
    
    Case TickTypeOpenPrice
        OpenPriceText(lIndex) = FormatPrice(ev.Tick.Price, lSecType, lTicksize)
    End Select
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionAwaitRetry(ByVal pSource As Object, ByVal pRetryInterval As Long)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionAwaitRetry"
On Error GoTo Err

LogMessage "Reconnecting to TWS in " & pRetryInterval & " seconds"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
' INotificationListener Interface Members
'@================================================================================

Private Sub INotificationListener_Notify(ev As NotificationEventData)
LogMessage "Notification " & ev.EventCode & ": " & ev.EventMessage
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AddListenerCheck_Click(pIndex As Integer)
Const ProcName As String = "AddListenerCheck_Click"
On Error GoTo Err

If mMarketDataSource(pIndex) Is Nothing Then Exit Sub
If mMarketDataSource(pIndex).State = MarketDataSourceStateStopped Then Exit Sub

If AddListenerCheck(pIndex).Value = vbChecked Then
    Set mMarketDataListener(pIndex) = New MarketDataListener
    mMarketDataListener(pIndex).Listen mMarketDataSource(pIndex)
Else
    mMarketDataListener(pIndex).UnListen
    Set mMarketDataListener(pIndex) = Nothing
    If mMarketDataSource(pIndex).State = MarketDataSourceStateStopped And ContractSpecBuilder1.IsReady Then StartTickerButton(pIndex).Enabled = True
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ContractSpecBuilder1_NotReady()
mFutureWaiter.Cancel
StartTickerButton(0).Enabled = False
StartTickerButton(1).Enabled = False
End Sub

Private Sub ContractSpecBuilder1_Ready()
Const ProcName As String = "ContractSpecBuilder1_Ready"
On Error GoTo Err

mFutureWaiter.Cancel
mFutureWaiter.Add FetchContract(ContractSpecBuilder1.ContractSpecifier, mContractStore, pCookie:=ContractSpecBuilder1.ContractSpecifier)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub StartTickerButton_Click(pIndex As Integer)
Const ProcName As String = "StartTickerButton_Click"
On Error GoTo Err

Dim lContractFuture As IFuture
Set lContractFuture = FetchContract(ContractSpecBuilder1.ContractSpecifier, mContractStore, , pCookie:=ContractSpecBuilder1.ContractSpecifier)
mFutureWaiter.Add lContractFuture

Dim lDataSource As IMarketDataSource
Set lDataSource = mDataManager.CreateMarketDataSource(lContractFuture, False, IIf(pIndex = 0, DataSourceKey0, DataSourceKey1))
lDataSource.AddErrorListener Me
lDataSource.AddGenericTickListener Me
lDataSource.StartMarketData
If AddListenerCheck(pIndex).Value = vbChecked Then
    Set mMarketDataListener(pIndex) = New MarketDataListener
    mMarketDataListener(pIndex).Listen lDataSource
End If

Set mMarketDataSource(pIndex) = lDataSource

StopTickerButton(pIndex).Enabled = True
StartTickerButton(pIndex).Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StopTickerButton_Click(pIndex As Integer)
Const ProcName As String = "StopTickerButton_Click"
On Error GoTo Err

mMarketDataSource(pIndex).StopMarketData
mMarketDataSource(pIndex).RemoveGenericTickListener Me
mMarketDataSource(pIndex).Finish
clearTickerFields pIndex
StopTickerButton(pIndex).Enabled = False
If ContractSpecBuilder1.IsReady Then StartTickerButton(pIndex).Enabled = True
Set mMarketDataSource(pIndex) = Nothing

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

If ev.Future.IsPending Then Exit Sub

Dim lContractSpec As IContractSpecifier
Set lContractSpec = ev.Future.Cookie

If ev.Future.IsFaulted Then
    If ev.Future.ErrorNumber = ErrorCodes.ErrIllegalArgumentException Then
        LogMessage ev.Future.ErrorMessage & ": " & lContractSpec.ToString
    Else
        LogMessage "Error " & ev.Future.ErrorNumber & " retrieving contracts for: " & lContractSpec.ToString & vbCrLf & _
                            ev.Future.ErrorMessage & vbCrLf & _
                            ev.Future.ErrorSource
    End If
    StartTickerButton(0).Enabled = False
    StartTickerButton(1).Enabled = False
ElseIf ev.Future.IsCancelled Then
    'LogMessage "Contract data fetch cancelled for: " & lContractSpec.ToString
    StartTickerButton(0).Enabled = False
    StartTickerButton(1).Enabled = False
Else
    If mMarketDataSource(0) Is Nothing Then StartTickerButton(0).Enabled = True
    If mMarketDataSource(1) Is Nothing Then StartTickerButton(1).Enabled = True
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mUnhandledErrorHandler Event Handlers
'@================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As ErrorEventData)

mClient.Finish

MsgBox "Unhandled error", vbCritical, "Ooops!"
'handleFatalError
'
'' Tell TWUtilities that we've now handled this unhandled error. Not actually
'' needed here because HandleFatalError never returns anyway
'UnhandledErrorHandler.Handled = True
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

Private Sub clearTickerFields(ByVal pIndex As Integer)
BidSizeText(pIndex) = ""
BidPriceText(pIndex) = ""
AskSizeText(pIndex) = ""
AskPriceText(pIndex) = ""
TradeSizeText(pIndex) = ""
TradePriceText(pIndex) = ""
VolumeText(pIndex) = ""
HighText(pIndex) = ""
LowText(pIndex) = ""
OpenPriceText(pIndex) = ""
ClosePriceText(pIndex) = ""
End Sub

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

Private Function getDataSourceIndex(ByVal pDataSource As IMarketDataSource) As Long
Const ProcName As String = "getDataSourceIndex"
On Error GoTo Err

getDataSourceIndex = IIf(pDataSource.Key = DataSourceKey0, 0, 1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

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






