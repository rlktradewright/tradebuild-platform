VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                    As String = "TwsAPI"

'================================================================================
' Enums
'================================================================================

Private Enum TwsHistDataDateFormats
    DateFormatString = 1
    DateFormatInteger = 2
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mServer                                     As String
Private mPort                                       As String

Private mClientID                                   As Long

Private mProgramErrorHandler                        As IProgramErrorListener
Private mErrorAndNotificationConsumer               As IErrorAndNotificationConsumer
Private mConnectionStatusConsumer                   As IConnectionStatusConsumer

Private WithEvents mSocketHandler                   As SocketHandler
Attribute mSocketHandler.VB_VarHelpID = -1
Private WithEvents mConnectionEstablishedTLI        As TimerListItem
Attribute mConnectionEstablishedTLI.VB_VarHelpID = -1
Private mConnectionNegotiator                       As ApiConnectionNegotiator

Private mWriter                                     As BufferedWriter

Private WithEvents mInMessageHandler                As InputMessageHandler
Attribute mInMessageHandler.VB_VarHelpID = -1

Private mServerVersion                              As Long

Private mConnectionState                            As TwsConnectionStates

Private WithEvents mPerformanceTLISecond            As TimerListItem
Attribute mPerformanceTLISecond.VB_VarHelpID = -1
Private WithEvents mPerformanceTLIPeriod            As TimerListItem
Attribute mPerformanceTLIPeriod.VB_VarHelpID = -1

Private mPerformanceLogger                          As Logger

Private mStatsRecorder                              As PerformanceStatsRecorder

Private mConnectionRetryIntervalSecs                As Long

Private mLogLevel                                   As TwsLogLevels

Private mLogTwsMessages                             As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mInMessageHandler = New InputMessageHandler

mConnectionState = TwsConnectionStates.TwsConnNotConnected

mLogLevel = 0

Set mStatsRecorder = New PerformanceStatsRecorder
Set mPerformanceLogger = GetLogger("tradebuild.log.ibapi.performance")

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub Class_Terminate()
Debug.Print "TwsAPI terminated"
End Sub

'================================================================================
' mConnectionEstablishedTLI Event Handlers
'================================================================================

Private Sub mConnectionEstablishedTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mConnectionEstablishedTLI_StateChange"
On Error GoTo Err

If ev.State = TimerListItemStates.TimerListItemStateExpired Then
    If Not mSocketHandler Is Nothing Then mSocketHandler.Disconnect "API connection has not been completed", False
    Set mConnectionEstablishedTLI = Nothing
ElseIf ev.State = TimerListItemStates.TimerListItemStateCancelled Then
    Set mConnectionEstablishedTLI = Nothing
    startPerformanceLogging
End If

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'================================================================================
' mInMessageHandler Event Handlers
'================================================================================

Private Sub mInMessageHandler_ApiConnectionEstablished(ByVal pServerVersion As Variant)
Const ProcName As String = "mInMessageHandler_ApiConnectionEstablished"
On Error GoTo Err

If Not mConnectionEstablishedTLI Is Nothing Then
    mConnectionEstablishedTLI.Cancel
    Set mConnectionEstablishedTLI = Nothing
End If

mServerVersion = pServerVersion
If mLogLevel <> 0 Then sendLogLevel

setConnectionState TwsConnConnected, ConnectionString

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mInMessageHandler_PaperTradingDisclaimerNotAcknowledged()
mSocketHandler.TempConnectionRetryIntervalSecs = 30
End Sub

'================================================================================
' mPerformanceTLIPeriod Event Handlers
'================================================================================

Private Sub mPerformanceTLIPeriod_StateChange(ev As TWUtilities40.StateChangeEventData)
Const ProcName As String = "mPerformanceTLIPeriod_StateChange"
On Error GoTo Err

If ev.State <> TimerListItemStates.TimerListItemStateExpired Then Exit Sub
If Not mPerformanceLogger.IsLoggable(LogLevelMediumDetail) Then Exit Sub

mPerformanceLogger.Log LogLevelMediumDetail, "Socket message statistics:" & vbCrLf & mStatsRecorder.GenerateStats
startPerformanceTLIPeriod

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'================================================================================
' mPerformanceTLISecond Event Handlers
'================================================================================

Private Sub mPerformanceTLISecond_StateChange(ev As TWUtilities40.StateChangeEventData)
Const ProcName As String = "mPerformanceTLISecond_StateChange"
On Error GoTo Err

If ev.State = TimerListItemStates.TimerListItemStateExpired Then mStatsRecorder.AccumulateStats
startPerformanceTLISecond

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'================================================================================
' mSocketHandler Event Handlers
'================================================================================

Private Sub mSocketHandler_Connected()
Const ProcName As String = "mSocketHandler_Connected"
On Error GoTo Err

Set mConnectionEstablishedTLI = GetGlobalTimerList.Add(Nothing, 30, ExpiryTimeUnitSeconds)
mInMessageHandler.Reset

mConnectionNegotiator.NegotiateApiConnection mClientID

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_ConnectFailed( _
                ByVal pDescription As String, _
                ByVal pRetryInterval As Long)
Const ProcName As String = "mSocketHandler_ConnectFailed"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnFailed, pDescription & ": " & ConnectionString

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_Connecting()
Const ProcName As String = "mSocketHandler_Connecting"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnConnecting, ConnectionString

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_ConnectionClosed( _
                ByVal pDescription As String)
Const ProcName As String = "mSocketHandler_ConnectionClosed"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnNotConnected, pDescription & ": " & ConnectionString

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_Disconnected(ByVal pReason As String)
Const ProcName As String = "mSocketHandler_Disconnected"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnNotConnected, pReason
stopPerformanceLogging

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let AccountDataConsumer(ByVal Value As IAccountDataConsumer)
mInMessageHandler.AccountDataConsumer = Value
End Property

Public Property Get AccountDataConsumer() As IAccountDataConsumer
Set AccountDataConsumer = mInMessageHandler.AccountDataConsumer
End Property

Public Property Get ClientID() As Long
ClientID = mClientID
End Property

Public Property Let ConnectionRetryIntervalSecs(ByVal Value As Long)
Const ProcName As String = "ConnectionRetryIntervalSecs"
On Error GoTo Err

AssertArgument Value >= 0, "Value cannot be negative"
mConnectionRetryIntervalSecs = Value
If Not mSocketHandler Is Nothing Then mSocketHandler.ConnectionRetryIntervalSecs = Value

Exit Property

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Property

Public Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Public Property Get ConnectionState() As TwsConnectionStates
ConnectionState = mConnectionState
End Property

Public Property Let ConnectionStatusConsumer(ByVal Value As IConnectionStatusConsumer)
Set mConnectionStatusConsumer = Value
mInMessageHandler.ConnectionStatusConsumer = Value
End Property

Public Property Get ConnectionStatusConsumer() As IConnectionStatusConsumer
Set ConnectionStatusConsumer = mConnectionStatusConsumer
End Property

Public Property Get ConnectionString() As String
ConnectionString = "server=" & mServer & _
                        " port=" & mPort & _
                        " client Id=" & mClientID
End Property

Public Property Let ContractDetailsConsumer(ByVal Value As IContractDetailsConsumer)
mInMessageHandler.ContractDetailsConsumer = Value
End Property

Public Property Get ContractDetailsConsumer() As IContractDetailsConsumer
Set ContractDetailsConsumer = mInMessageHandler.ContractDetailsConsumer
End Property

Public Property Let ErrorAndNotificationConsumer(ByVal Value As IErrorAndNotificationConsumer)
Set mErrorAndNotificationConsumer = Value
End Property

Public Property Get ErrorAndNotificationConsumer() As IErrorAndNotificationConsumer
Set ErrorAndNotificationConsumer = mErrorAndNotificationConsumer
End Property

Public Property Let HistDataConsumer(ByVal Value As IHistDataConsumer)
mInMessageHandler.HistDataConsumer = Value
End Property

Public Property Get HistDataConsumer() As IHistDataConsumer
Set HistDataConsumer = mInMessageHandler.HistDataConsumer
End Property

Public Property Let MarketDataConsumer(ByVal Value As IMarketDataConsumer)
mInMessageHandler.MarketDataConsumer = Value
End Property

Public Property Get MarketDataConsumer() As IMarketDataConsumer
Set MarketDataConsumer = mInMessageHandler.MarketDataConsumer
End Property

Public Property Let MarketDepthConsumer(ByVal Value As IMarketDepthConsumer)
mInMessageHandler.MarketDepthConsumer = Value
End Property

Public Property Get MarketDepthConsumer() As IMarketDepthConsumer
Set MarketDepthConsumer = mInMessageHandler.MarketDepthConsumer
End Property

Public Property Get MaxContractRequestId() As Long
MaxContractRequestId = BaseOrderId - BaseContractRequestId - 1
End Property

Public Property Get MaxExecutionsRequestId() As Long
MaxExecutionsRequestId = BaseContractRequestId - BaseExecutionsRequestId - 1
End Property

Public Property Get MaxHistoricalDataRequestId() As Long
MaxHistoricalDataRequestId = BaseExecutionsRequestId - BaseHistoricalDataRequestId - 1
End Property

Public Property Get MaxMarketDataRequestId() As Long
MaxMarketDataRequestId = BaseMarketDepthRequestId - 1
End Property

Public Property Get MaxMarketDepthRequestId() As Long
MaxMarketDepthRequestId = BaseHistoricalDataRequestId - BaseMarketDepthRequestId - 1
End Property

Public Property Get MaxOrderId() As Long
MaxOrderId = &H7FFFFFFF
End Property

Public Property Get NextOrderId() As Long
NextOrderId = mInMessageHandler.NextOrderId
End Property

Public Property Let OrderInfoConsumer(ByVal Value As IOrderInfoConsumer)
mInMessageHandler.OrderInfoConsumer = Value
End Property

Public Property Get OrderInfoConsumer() As IOrderInfoConsumer
Set OrderInfoConsumer = mInMessageHandler.OrderInfoConsumer
End Property

Public Property Get Port() As String
Port = mPort
End Property

Public Property Let ProgramErrorHandler(ByVal Value As IProgramErrorListener)
Set mProgramErrorHandler = Value
End Property

Public Property Get ProgramErrorHandler() As IProgramErrorListener
Set ProgramErrorHandler = mProgramErrorHandler
End Property

Public Property Let ScannerDataConsumer(ByVal Value As IScannerDataConsumer)
mInMessageHandler.ScannerDataConsumer = Value
End Property

Public Property Get ScannerDataConsumer() As IScannerDataConsumer
Set ScannerDataConsumer = mInMessageHandler.ScannerDataConsumer
End Property

'Public Property Get SerialiseContractDetailsRequests() As Boolean
'SerialiseContractDetailsRequests = mInMessageHandler.SerialiseContractDetailsRequests
'End Property

Public Property Get Server() As String
Server = mServer
End Property

Public Property Get ServerVersion() As Long
ServerVersion = mServerVersion
End Property

'================================================================================
' Methods
'================================================================================

Public Sub CalculateImpliedVolatility( _
        ByVal pReqId As Long, _
        ByVal pContract As TwsContract, _
        ByVal pOptionPrice As Double, _
        ByVal pUnderPrice As Double)
Const ProcName As String = "CalculateImpliedVolatility"
On Error GoTo Err

AssertArgument mServerVersion < MIN_SERVER_VER_TRADING_CLASS Or pContract.TradingClass = "", "TradingClass property in CalculateImpliedVolatility not supported"

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 2

mWriter.AddMessageId REQ_CALC_IMPLIED_VOLAT
mWriter.AddString VERSION, "Version"
mWriter.AddString pReqId, "ReqId"

With pContract
    mWriter.AddString .ConId, "Contract id"
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    mWriter.AddString IIf(.Multiplier = 1, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString IIf(.Exchange = "", "*", .Exchange), "Exchange"
    mWriter.AddString .PrimaryExch, "PrimaryExch"
    mWriter.AddString .CurrencyCode, "Currency"
    mWriter.AddString UCase$(.LocalSymbol), "Local Symbol"
    If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString .TradingClass, "Trading Class"
End With

mWriter.AddString pOptionPrice, "OptionPrice"
mWriter.AddString pUnderPrice, "UnderPrice"

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CalculateOptionPrice( _
        ByVal pReqId As Long, _
        ByVal pContract As TwsContract, _
        ByVal pVolatility As Double, _
        ByVal pUnderPrice As Double)
Const ProcName As String = "CalculateOptionPrice"
On Error GoTo Err

AssertArgument mServerVersion < MIN_SERVER_VER_TRADING_CLASS Or pContract.TradingClass = "", "TradingClass property in CalculateImpliedVolatility not supported"

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 2

mWriter.AddMessageId REQ_CALC_OPTION_PRICE
mWriter.AddString VERSION, "Version"
mWriter.AddString pReqId, "ReqId"

With pContract
    mWriter.AddString .ConId, "Contract id"
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    mWriter.AddString IIf(.Multiplier = 1, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString IIf(.Exchange = "", "*", .Exchange), "Exchange"
    mWriter.AddString .PrimaryExch, "PrimaryExch"
    mWriter.AddString .CurrencyCode, "Currency"
    mWriter.AddString UCase$(.LocalSymbol), "Local Symbol"
    If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString .TradingClass, "Trading Class"
End With

mWriter.AddString pVolatility, "Volatility"
mWriter.AddString pUnderPrice, "UnderPrice"

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelAccountSummary( _
                ByVal pReqId As Long)
Const ProcName As String = "CancelAccountSummary"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

AssertArgument mServerVersion >= MIN_SERVER_VER_ACCT_SUMMARY, "Account summary request cancellation not supported"

Const VERSION = 1

mWriter.AddMessageId CANCEL_ACCOUNT_SUMMARY
mWriter.AddString VERSION, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelCalculateImpliedVolatility(ByVal pReqId As Long)
Const ProcName As String = "CancelCalculateImpliedVolatility"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 1

mWriter.AddMessageId CANCEL_CALC_IMPLIED_VOLAT
mWriter.AddString VERSION, "Version"
mWriter.AddString pReqId, "ReqId"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelCalculateOptionPrice(ByVal pReqId As Long)
Const ProcName As String = "CancelCalculateOptionPrice"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 1

mWriter.AddMessageId CANCEL_CALC_OPTION_PRICE
mWriter.AddString VERSION, "Version"
mWriter.AddString pReqId, "ReqId"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelHistoricalData( _
                ByVal pRequestId As Long)
Const ProcName As String = "CancelHistoricalData"
On Error GoTo Err

cancelHistDataRequest pRequestId

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelMarketData(ByVal pTickerId As Long)
Const ProcName As String = "CancelMarketData"
On Error GoTo Err

cancelMarketDataEx pTickerId

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelFundamentalData(ByVal pReqId As Long)
Const ProcName As String = "CancelFundamentalData"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 1

mWriter.AddMessageId CANCEL_FUNDAMENTAL_DATA
mWriter.AddString VERSION, "Version"
mWriter.AddString pReqId, "ReqId"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelMarketDepth(ByVal pTickerId As Long)
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 1

mWriter.AddMessageId CANCEL_MKT_DEPTH
mWriter.AddString VERSION, "Version"
mWriter.AddString getTwsMarketDepthRequestIdFromCallersRequestId(pTickerId), "Ticker id"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mErrorAndNotificationConsumer, pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub CancelNewsBulletins()
Const ProcName As String = "CancelNewsBulletins"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 1

mWriter.AddMessageId CANCEL_NEWS_BULLETINS
mWriter.AddString VERSION, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelOrder(ByVal pOrderId As Long)
Const ProcName As String = "CancelOrder"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1
mWriter.AddMessageId CANCEL_ORDER
mWriter.AddString VERSION, "Version"
mWriter.AddString pOrderId, "Order id"
mWriter.Send True

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelPositions()
Const ProcName As String = "CancelPositions"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

AssertArgument mServerVersion >= MIN_SERVER_VER_ACCT_SUMMARY, "Cancel position requests not supported"

Const VERSION = 1

mWriter.AddMessageId CANCEL_POSITIONS
mWriter.AddString VERSION, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelRealTimeBars(ByVal pTickerId As Long)
Const ProcName As String = "CancelRealTimeBars"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId CANCEL_REAL_TIME_BARS
mWriter.AddString VERSION, "Version"
mWriter.AddString pTickerId, "Ticker id"
mWriter.Send True

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelScannerSubscription(ByVal pTickerId As Long)
Const ProcName As String = "CancelScannerSubscription"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId CANCEL_SCANNER_SUBSCRIPTION
mWriter.AddString VERSION, "Version"
mWriter.AddString pTickerId, "Ticker Id"
mWriter.Send True

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub Connect()
Const ProcName As String = "Connect"
On Error GoTo Err

Dim lReader As New BufferedReader
lReader.Initialise mInMessageHandler, mLogTwsMessages

Set mSocketHandler = New SocketHandler
mSocketHandler.Initialise mServer, mPort, lReader, mProgramErrorHandler
mSocketHandler.ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs

Set mWriter = New BufferedWriter
mWriter.Initialise mSocketHandler

Set mConnectionNegotiator = New ApiConnectionNegotiator
mConnectionNegotiator.Initialise lReader, mWriter

mInMessageHandler.Initialise mConnectionNegotiator, _
                            lReader, _
                            mWriter, _
                            mStatsRecorder, _
                            mErrorAndNotificationConsumer, _
                            mProgramErrorHandler, _
                            mLogTwsMessages

mSocketHandler.Connect

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub Disconnect( _
                ByVal pReason As String)
Const ProcName As String = "Disconnect"
On Error GoTo Err

mSocketHandler.Disconnect pReason, True
Set mSocketHandler = Nothing

stopPerformanceLogging

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub ExerciseOptions( _
                ByVal pTickerId As Long, _
                ByVal pContract As TwsContract, _
                ByVal pExerciseAction As Long, _
                ByVal pExerciseQuantity As Long, _
                ByVal pAccount As String, _
                ByVal pOverride As Long)
Const ProcName As String = "ExerciseOptions"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 2

AssertArgument mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Or (pContract.TradingClass = "" And pContract.ConId = 0), "ConId and TradingClass properties in ExerciseOptions not supported"

mWriter.AddMessageId EXERCISE_OPTIONS
mWriter.AddString VERSION, "Version"
mWriter.AddString pTickerId, "TickerId"

If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString pContract.ConId, "ConId"

mWriter.AddString pContract.Symbol, "Symbol"
mWriter.AddString gTwsSecTypeToShortString(pContract.Sectype), "Sectype"
mWriter.AddString pContract.Expiry, "Expiry"
mWriter.AddString pContract.Strike, "Strike"
mWriter.AddString gTwsOptionRightToString(pContract.OptRight), "Right"
mWriter.AddString IIf(pContract.Multiplier = 1, "", pContract.Multiplier), "Multiplier"
mWriter.AddString pContract.Exchange, "Exchange"
mWriter.AddString pContract.CurrencyCode, "Currency"
mWriter.AddString UCase$(pContract.LocalSymbol), "LocalSymbol"
If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString pContract.TradingClass, "TradingClass"

mWriter.AddString pExerciseAction, "ExerciseAction"
mWriter.AddString pExerciseQuantity, "ExerciseQuantity"
mWriter.AddString pAccount, "Account"
mWriter.AddString pOverride, "Override"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Function GetRawTwsMessage() As Byte()
Const ProcName As String = "GetRawMessage"
On Error GoTo Err

GetRawTwsMessage = mInMessageHandler.BufferedReader.GetRawMessage

Exit Function

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pClientId As Long, _
                Optional ByVal pLogTwsMessages As Boolean = True)
AssertArgument pPort > 0, "Port must be > 0"
AssertArgument pClientId >= 0, "ClientID must be >= 0"

If pServer = "" Then pServer = "127.0.0.1"
mServer = pServer
mPort = pPort
mClientID = pClientId
mLogTwsMessages = pLogTwsMessages
End Sub

Public Sub PlaceOrder( _
                ByVal pOrder As TwsOrder, _
                ByVal pContract As TwsContract)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

Dim i As Long
Dim lOrderComboLegsCount As Long

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

AssertArgument mServerVersion >= MIN_SERVER_VER_DELTA_NEUTRAL_OPEN_CLOSE Or _
    (pOrder.DeltaNeutralOpenClose = "" And _
        Not pOrder.DeltaNeutralShortSale And _
        pOrder.DeltaNeutralShortSaleSlot = 0 And _
        pOrder.DeltaNeutralDesignatedLocation = ""), _
    "deltaNeutral parameters: OpenClose, ShortSale, ShortSaleSlot, DesignatedLocation not supported"

AssertArgument mServerVersion >= MIN_SERVER_VER_SCALE_ORDERS3 Or _
    ((pOrder.ScalePriceIncrement = 0 And pOrder.ScalePriceIncrement = MaxDouble) And _
        pOrder.ScalePriceAdjustValue = MaxDouble And _
        pOrder.ScalePriceAdjustInterval = MaxLong And _
        pOrder.ScaleProfitOffset = MaxDouble And _
        Not pOrder.ScaleAutoReset And _
        pOrder.ScaleInitPosition = MaxLong And _
        pOrder.ScaleInitFillQty = MaxLong And _
        Not pOrder.ScaleRandomPercent), _
    "Scale order parameters: PriceAdjustValue, PriceAdjustInterval, ProfitOffset, AutoReset, InitPosition, InitFillQty and RandomPercent not supported"

If mServerVersion < MIN_SERVER_VER_ORDER_COMBO_LEGS_PRICE And pContract.Sectype = TwsSecTypes.TwsSecTypeCombo Then
    On Error Resume Next
    lOrderComboLegsCount = UBound(pOrder.OrderComboLegs) + 1
    On Error GoTo Err
    
    If lOrderComboLegsCount > 0 Then
        Dim lEntry As Variant
        For Each lEntry In pOrder.OrderComboLegs
            Dim lComboLeg As TwsOrderComboLeg
            Set lComboLeg = lEntry
            AssertArgument lComboLeg.Price = MaxDouble, "Per-leg prices for order combo legs not supported"
        Next
    End If
End If

AssertArgument mServerVersion >= MIN_SERVER_VER_TRAILING_PERCENT Or pOrder.TrailingPercent = MaxDouble, "Trailing percent property not supported"

AssertArgument mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Or pContract.TradingClass = "", "TradingClass property not suppoerted"

AssertArgument mServerVersion >= MIN_SERVER_VER_SCALE_TABLE Or (pOrder.ScaleTable = "" And pOrder.ActiveStartTime = "" And pOrder.ActiveStopTime = ""), "ScaleTable, ActiveStartTime and ActiveStopTime properties not supported"

Const VERSION = 41

mWriter.AddMessageId PLACE_ORDER
mWriter.AddString VERSION, "Version"

If pOrder.OrderId = 0 Then pOrder.OrderId = NextOrderId
Assert pOrder.OrderId >= BaseOrderId, "Order id must not be less than " & BaseOrderId
If pOrder.OrderId > mInMessageHandler.NextOrderId Then mInMessageHandler.NextOrderId = pOrder.OrderId + 1

mWriter.AddString pOrder.OrderId, "Order id"

' mwriter.send contract fields
With pContract
    mWriter.AddString .ConId, "Con id"
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sectype"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    mWriter.AddString IIf(.Multiplier = 1, "", CStr(.Multiplier)), "Multiplier"  ' multiplier - this is the number of units of the underlying
                                                            ' delivered per contract - see http://www.interactivebrokers.com/discus/messages/2/27413.html
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString .PrimaryExch, "Primary Exchange"    ' only relevant when
                                                        ' Exchange is SMART and there are SMART routers in more
                                                        ' than one country (eg try IBM in Tws)
    mWriter.AddString .CurrencyCode, "Currency"
    mWriter.AddString UCase$(.LocalSymbol), "Local Symbol"
    If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString .TradingClass, "Trading Class"
    mWriter.AddString .SecIdType, "Sec id type"
    mWriter.AddString .SecId, "Sec id"
End With

With pOrder
    
    ' mwriter.send main Order fields
    mWriter.AddString gTwsOrderActionToString(.Action), "Action"
    mWriter.AddString .TotalQuantity, "Quantity"
    mWriter.AddString gTwsOrderTypeToString(.OrderType), "Order type"
    
    If mServerVersion < MIN_SERVER_VER_ORDER_COMBO_LEGS_PRICE Then
        mWriter.AddDouble IIf(.LmtPrice = MaxDouble, 0, .LmtPrice), "Price"
    Else
        mWriter.AddDouble .LmtPrice, "Price"
    End If
    If mServerVersion < MIN_SERVER_VER_TRAILING_PERCENT Then
        mWriter.AddDouble IIf(.AuxPrice = MaxDouble, 0, .AuxPrice), "Aux price"
    Else
        mWriter.AddDouble .AuxPrice, "Aux price"
    End If
    
    ' mwriter.send extended Order fields
    mWriter.AddString gTwsOrderTIFToString(.Tif), "TIF"
    mWriter.AddString .OcaGroup, "Oca Group"
    mWriter.AddString .Account, "Account"
    mWriter.AddString .OpenClose, "OpenClose"
    mWriter.AddString .Origin, "Origin"
    mWriter.AddString .OrderRef, "Order ref"
    mWriter.AddBoolean .Transmit, "Transmit"
    mWriter.AddString .ParentId, "Parent id"
    mWriter.AddBoolean .BlockOrder, "Block Order"
    mWriter.AddBoolean .SweepToFill, "Sweep to fill"
    mWriter.AddString .DisplaySize, "Display Size"
    mWriter.AddString stopTriggerMethodToString(.TriggerMethod), "Trigger method"
    mWriter.AddBoolean .OutsideRth, "Outside RTH"
    mWriter.AddBoolean .Hidden, "Hidden"
End With

' send combo legs for BAG requests
With pContract
    If .Sectype = TwsSecTypeCombo Then
        If .ComboLegs.Count = 0 Then
            mWriter.AddString 0, "Combo legs count"
        Else
            mWriter.AddString .ComboLegs.Count, "Combo legs count"
            Dim comboLeg As TwsComboLeg
            For Each comboLeg In .ComboLegs
                With comboLeg
                    mWriter.AddString .ConId, "Leg " & i & " Con id"
                    mWriter.AddString .Ratio, "Leg " & i & " Ration"
                    mWriter.AddString gTwsOrderActionToString(.Action), "Leg " & i & " Action"
                    mWriter.AddString .Exchange, "Leg " & i & " Exchange"
                    mWriter.AddString legOpenCloseToString(.OpenClose), "Leg " & i & " Open/close"
                
                    mWriter.AddString comboLeg.ShortSaleSlot, "Short Sale Slot"
                    mWriter.AddString comboLeg.DesignatedLocation, "Designated Location"
                    mWriter.AddString comboLeg.ExemptCode, "Exempt Code"
                End With
            Next
        End If
    End If
End With

' Send order combo legs for BAG requests
If mServerVersion >= MIN_SERVER_VER_ORDER_COMBO_LEGS_PRICE And pContract.Sectype = TwsSecTypeCombo Then
    If lOrderComboLegsCount = 0 Then
        mWriter.AddString 0, "Order Combo Legs Count"
    Else
        Dim lOrderComboLegs() As TwsOrderComboLeg
        lOrderComboLegs = pOrder.OrderComboLegs
        mWriter.AddString lOrderComboLegsCount, "Order Combo Legs Count"

        For i = 0 To lOrderComboLegsCount - 1
            Dim lOrderComboLeg  As TwsOrderComboLeg
            Set lOrderComboLeg = lOrderComboLegs(i)
            mWriter.AddString lOrderComboLeg.Price, "Leg" & i & " Price"
        Next
    End If
End If

If pContract.Sectype = TwsSecTypeCombo Then
    Dim lSmartComboRoutingParamsCount As Long
    On Error Resume Next
    lSmartComboRoutingParamsCount = UBound(pOrder.SmartComboRoutingParams) + 1
    On Error GoTo Err
    
    If lSmartComboRoutingParamsCount = 0 Then
        mWriter.AddString 0, "Smart Combo Routing Params Count"
    Else
        mWriter.AddString lSmartComboRoutingParamsCount, "Smart Combo Routing Params Count"
        
        Dim lSmartComboRoutingParams() As TwsTagValue
        lSmartComboRoutingParams = pOrder.SmartComboRoutingParams
        For i = 0 To lSmartComboRoutingParamsCount - 1
            Dim l As TwsTagValue
            l = lSmartComboRoutingParams(i)
            mWriter.AddString l.Tag, "Param" & i & " Tag"
            mWriter.AddString l.Value, "Param" & i & " Value"
        Next
    End If
End If

With pOrder
    mWriter.AddString "", "Shares Allocation"
    mWriter.AddString .DiscretionaryAmt, "Discretionary amount"
    mWriter.AddString .GoodAfterTime, "Good after Time"
    mWriter.AddString .GoodTillDate, "Good till date"
    
    mWriter.AddString .FaGroup, "FAGroup"
    mWriter.AddString .FaMethod, "FAMethod"
    mWriter.AddString .FaPercentage, "FAPercentage"
    mWriter.AddString .FaProfile, "FAProfile"
    
    'institutional short sale slot fields.
    mWriter.AddString .ShortSaleSlot, " Short Sale Slot"           ' 0 only for retail, 1 or 2 only for institution.
    mWriter.AddString .DesignatedLocation, "Designated Location"       ' only populate when shortSaleSlot = 2.
    
    mWriter.AddString .ExemptCode, "Exempt Code"
    
    mWriter.AddString .OcaType, "Oca type"
    mWriter.AddString .Rule80A, "Rule 80A"
    mWriter.AddString .SettlingFirm, "Settling firm"
    mWriter.AddBoolean .AllOrNone, "All or none"
    mWriter.AddLongMax .MinQty, "Minimum quantity"
    mWriter.AddDoubleMax .PercentOffset, "Percent Offset"
    mWriter.AddBoolean .ETradeOnly, "E-trade only"
    mWriter.AddBoolean .FirmQuoteOnly, "Firm quote only"
    mWriter.AddDoubleMax .NbboPriceCap, "NBBO price cap"
    mWriter.AddLongMax .AuctionStrategy, "Auction strategy"
    mWriter.AddDoubleMax .StartingPrice, "Starting price"
    mWriter.AddDoubleMax .StockRefPrice, "Stock ref price"
    mWriter.AddDoubleMax .Delta, "Delta"
    
    mWriter.AddDoubleMax .StockRangeLower, "Stock range lower"
    mWriter.AddDoubleMax .StockRangeUpper, "Stock range upper"
    
    mWriter.AddBoolean .OverridePercentageConstraints, "Override percentage constraints"

    ' Volatility orders
    mWriter.AddDoubleMax .Volatility, "Volatility"
    mWriter.AddLongMax .VolatilityType, "Volatility type"
    mWriter.AddString gTwsOrderTypeToString(.DeltaNeutralOrderType), "Delta neutral Order type"
    mWriter.AddDoubleMax .DeltaNeutralAuxPrice, "Delta neutral aux price"
    If .DeltaNeutralOrderType <> TwsOrderTypeNone Then
        mWriter.AddString .DeltaNeutralConId, "Delta Neutral Con Id"
        mWriter.AddString .DeltaNeutralSettlingFirm, "Delta Neutral Settling Firm"
        mWriter.AddString .DeltaNeutralClearingAccount, "Delta Neutral Clearing Account"
        mWriter.AddString .DeltaNeutralClearingIntent, "Delta Neutral Clearing Intent"
    End If

    If mServerVersion >= MIN_SERVER_VER_DELTA_NEUTRAL_OPEN_CLOSE And .DeltaNeutralOrderType <> TwsOrderTypeNone Then
        mWriter.AddString .DeltaNeutralOpenClose, "Delta Neutral Open Close"
        mWriter.AddString .DeltaNeutralShortSale, "Delta Neutral Short Sale"
        mWriter.AddString .DeltaNeutralShortSaleSlot, "Delta Neutral Short Sale Slot"
        mWriter.AddString .DeltaNeutralDesignatedLocation, "Delta Neutral Designated Location"
    End If
    
    mWriter.AddBoolean .ContinuousUpdate, "Continuous update"
    
    mWriter.AddLongMax .ReferencePriceType, "Reference price type"
    
    mWriter.AddDoubleMax .TrailStopPrice, "Trail stop price"
    If mServerVersion >= MIN_SERVER_VER_TRAILING_PERCENT Then mWriter.AddDoubleMax .TrailingPercent, "Trailing Percent"
    
    mWriter.AddLongMax .ScaleInitLevelSize, "ScaleInitLevelSize"
    mWriter.AddLongMax .ScaleSubsLevelSize, "ScaleSubsLevelSize"
    mWriter.AddDoubleMax .ScalePriceIncrement, "ScalePriceIncrement"
    
    If mServerVersion >= MIN_SERVER_VER_SCALE_ORDERS3 And .ScalePriceIncrement > 0# And .ScalePriceIncrement <> MaxDouble Then
        mWriter.AddDoubleMax .ScalePriceAdjustValue, "Scale Price Adjust Value"
        mWriter.AddLongMax .ScalePriceAdjustInterval, "Scale Price Adjust Interval"
        mWriter.AddDoubleMax .ScaleProfitOffset, "Scale Profit Offset"
        mWriter.AddBoolean .ScaleAutoReset, "Scale Auto Reset"
        mWriter.AddLongMax .ScaleInitPosition, "Scale Init POSITION"
        mWriter.AddLongMax .ScaleInitFillQty, "Scale Init Fill Qty"
        mWriter.AddBoolean .ScaleRandomPercent, "Scale Random Percent"
    End If

    If mServerVersion >= MIN_SERVER_VER_SCALE_TABLE Then
        mWriter.AddString .ScaleTable, "Scale Table"
        mWriter.AddString .ActiveStartTime, "Active Start Time"
        mWriter.AddString .ActiveStopTime, "Active Stop Time"
    End If

    mWriter.AddString gTwsHedgeTypeToString(.HedgeType), "Hedge Type"
    If .HedgeType <> TwsHedgeTypeNone Then mWriter.AddString .HedgeParam, "Hedge Param"

    mWriter.AddBoolean .OptOutSmartRouting, "Opt Out Smart Routing"

    mWriter.AddString .ClearingAccount, "Clearing account"
    mWriter.AddString .ClearingIntent, "Clearing intent"

    mWriter.AddBoolean .NotHeld, "Not held"

    If pContract.UnderComp.ConId <> 0 Then
        mWriter.AddBoolean True, "Under comp"
        mWriter.AddString pContract.UnderComp.ConId, "Under comp conid"
        mWriter.AddDouble pContract.UnderComp.Delta, "Under comp delta"
        mWriter.AddDouble pContract.UnderComp.Price, "Under comp price"
    Else
        mWriter.AddBoolean False, "Under comp"
    End If

    mWriter.AddString .AlgoStrategy, "Algo strategy"
    If .AlgoStrategy <> "" Then
        Dim algoParamsCount  As Long
        
        On Error Resume Next
        algoParamsCount = UBound(.AlgoParams) + 1
        On Error GoTo Err
         
        mWriter.AddString algoParamsCount, "Algo params count"
        If algoParamsCount > 0 Then
            Dim lAlgoParams() As TwsTagValue
            lAlgoParams = .AlgoParams
            
            For i = 0 To algoParamsCount - 1
                mWriter.AddString lAlgoParams(i).Tag, "Tag" & i
                mWriter.AddString lAlgoParams(i).Value, "value" & i
            Next
        End If
    End If

    mWriter.AddBoolean .WhatIf, "WhatIf"
End With

mWriter.Send True

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub
                    
Public Sub ReplaceFA(ByVal DataType As TwsFADataTypes, ByVal xml As String)
Const ProcName As String = "ReplaceFA"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

Assert mServerVersion >= 13, "Server version does not support ReplaceFA", ErrorCodes.ErrUnsupportedOperationException

mWriter.AddMessageId REPLACE_FA
mWriter.AddString VERSION, "Version"
mWriter.AddString DataType, "Data type"
mWriter.AddString xml, "XML"

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestAccountSummary( _
                ByVal pReqId As Long, _
                ByVal pGroup As String, _
                ByVal pTags As String)
Const ProcName As String = "RequestAccountSummary"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

AssertArgument mServerVersion >= MIN_SERVER_VER_ACCT_SUMMARY, "Account summary requests not supported"

Const VERSION = 1

mWriter.AddMessageId REQ_ACCOUNT_SUMMARY
mWriter.AddString VERSION, "Version"
mWriter.AddString pReqId, "ReqId"
mWriter.AddString pGroup, "Group"
mWriter.AddString pTags, "Tags"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestAccountUpdates(ByVal subscribe As Boolean, ByVal acctCode As String)
Const ProcName As String = "RequestAccountUpdates"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 2

mWriter.AddMessageId REQ_ACCT_DATA
mWriter.AddString VERSION, "Version"
mWriter.AddBoolean subscribe, "Subscribe"

mWriter.AddString acctCode, "Account code"

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestAllOpenOrders()
Const ProcName As String = "RequestAllOpenOrders"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
                            

Const VERSION = 1

mWriter.AddMessageId REQ_ALL_OPEN_ORDERS
mWriter.AddString VERSION, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestAutoOpenOrders(ByVal autoBind As Boolean)
Const ProcName As String = "RequestAutoOpenOrders"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_AUTO_OPEN_ORDERS
mWriter.AddString VERSION, "Version"
mWriter.AddBoolean autoBind, "Autobind"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestContractDetails( _
                ByVal pRequestId As Long, _
                ByVal pContract As TwsContract)
Const ProcName As String = "RequestContractDetails"
On Error GoTo Err

reqContractDetails pRequestId, pContract

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestCurrentTime()
Const ProcName As String = "RequestCurrentTime"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 3

mWriter.AddMessageId REQ_CURRENT_TIME
mWriter.AddString VERSION, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestExecutions( _
                ByVal pRequestId As Long, _
                ByVal filter As TwsExecutionFilter)
Const ProcName As String = "RequestExecutions"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 3

mWriter.AddMessageId REQ_EXECUTIONS
mWriter.AddString VERSION, "Version"

mWriter.AddString getTwsExecutionsRequestIdFromCallersRequestId(pRequestId), "ReqId"

Dim theFilter As TwsExecutionFilter
If filter Is Nothing Then
    Set theFilter = New TwsExecutionFilter
Else
    Set theFilter = filter
End If

With theFilter
    mWriter.AddString .ClientID, "Client id"
    mWriter.AddString .AccountCode, "Account code"
    If .Time = Int(.Time) Then
        mWriter.AddString IIf(CLng(.Time) <> 0, Format(.Time, "yyyymmdd"), ""), "Fill Time"
    Else
        mWriter.AddString IIf(CLng(.Time) <> 0, Format(.Time, "yyyymmdd-hh\:nn\:ss"), ""), "Fill Time"
    End If
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString gTwsOrderActionToString(.Action), "Action"
End With

mWriter.Send True

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestFA(ByVal DataType As TwsFADataTypes)
Const ProcName As String = "RequestFA"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_FA
mWriter.AddString VERSION, "Version"
mWriter.AddString DataType, "Data type"

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestFundamentalData( _
                ByVal pReqId As Long, _
                ByVal pContract As TwsContract, _
                ByVal pReportType As String)
Const ProcName As String = "RequestFundamentalData"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 2

mWriter.AddMessageId REQ_FUNDAMENTAL_DATA
mWriter.AddString VERSION, "Version"
mWriter.AddString pReqId, "ReqId"

If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString pContract.ConId, "ConId"
mWriter.AddString pContract.Symbol, "Symbol"
mWriter.AddString gTwsSecTypeToShortString(pContract.Sectype), "Sectype"
mWriter.AddString pContract.Exchange, "Exchange"
mWriter.AddString pContract.PrimaryExch, "PrimaryExch"
mWriter.AddString pContract.CurrencyCode, "Currency"
mWriter.AddString UCase$(pContract.LocalSymbol), "LocalSymbol"
mWriter.AddString UCase$(pReportType), "ReportType"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestGlobalCancel()
Const ProcName As String = "RequestGlobalCancel"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_GLOBAL_CANCEL
mWriter.AddString VERSION, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestHistoricalData( _
                ByVal pRequestId As Long, _
                ByRef pRequest As TwsHistoricalDataRequest)
Const ProcName As String = "RequestHistoricalData"
On Error GoTo Err

reqHistoricalData pRequestId, pRequest

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestIds(ByVal pNumIds As Long)
Const ProcName As String = "RequestIds"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_IDS
mWriter.AddString VERSION, "Version"
mWriter.AddString pNumIds, "NumIds"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestManagedAccounts()
Const ProcName As String = "RequestManagedAccounts"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_MANAGED_ACCTS
mWriter.AddString VERSION, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestMarketData( _
                ByVal pTickerId As Long, _
                ByVal pContract As TwsContract, _
                ByVal pGenericTicks As String, _
                ByVal pSnapshot As Boolean)
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

reqMarketData pTickerId, pContract, pGenericTicks, pSnapshot

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestMarketDataType(ByVal pMarketDataType As Long)
Const ProcName As String = "RequestMarketDataType"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_MARKET_DATA_TYPE
mWriter.AddString VERSION, "Version"
mWriter.AddString pMarketDataType, "MarketDataType"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestMarketDepth( _
                ByVal pTickerId As Long, _
                ByVal pContract As TwsContract, _
                Optional ByVal pNumberOfRows As Long = 20)
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

reqMarketDepth pTickerId, pContract, pNumberOfRows

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestNewsBulletins(allMsgs As Boolean)
Const ProcName As String = "RequestNewsBulletins"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_NEWS_BULLETINS
mWriter.AddString VERSION, "Version"
mWriter.AddString allMsgs, "All messages"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestOpenOrders()
Const ProcName As String = "RequestOpenOrders"
On Error GoTo Err

Const VERSION = 1

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

mWriter.AddMessageId REQ_OPEN_ORDERS
mWriter.AddString VERSION, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestPositions()
Const ProcName As String = "RequestPositions"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

AssertArgument mServerVersion >= MIN_SERVER_VER_ACCT_SUMMARY, "Position requests not supported"

Const VERSION = 1

mWriter.AddMessageId REQ_POSITIONS
mWriter.AddString VERSION, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestRealTimeBars( _
                ByVal pTickerId As Long, _
                ByVal pContract As TwsContract, _
                ByVal pBarSize As Long, _
                ByVal pWhatToShow As String, _
                ByVal pUseRTH As Boolean)
Const ProcName As String = "RequestRealTimeBars"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

AssertArgument mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Or (pContract.TradingClass = "" And pContract.ConId = 0), "ConId and TradingClass properties in RequestRealTimeBars not supported"

Const VERSION = 2

mWriter.AddMessageId REQ_REAL_TIME_BARS
mWriter.AddString VERSION, "Version"
mWriter.AddString pTickerId, "TickerId"

' send contract fields
If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString pContract.ConId, "ConId"
mWriter.AddString pContract.Symbol, "Symbol"
mWriter.AddString gTwsSecTypeToShortString(pContract.Sectype), "Sectype"
mWriter.AddString pContract.Expiry, "Expiry"
mWriter.AddString pContract.Strike, "Strike"
mWriter.AddString gTwsOptionRightToString(pContract.OptRight), "Right"
mWriter.AddString IIf(pContract.Multiplier = 1, "", pContract.Multiplier), "Multiplier"
mWriter.AddString pContract.Exchange, "Exchange"
mWriter.AddString pContract.PrimaryExch, "PrimaryExch"
mWriter.AddString pContract.CurrencyCode, "Currency"
mWriter.AddString UCase$(pContract.LocalSymbol), "LocalSymbol"
If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString pContract.TradingClass, "TradingClass"
mWriter.AddString pBarSize, "BarSize"   ' this parameter is not currently used
mWriter.AddString pWhatToShow, "WhatToShow"
mWriter.AddString pUseRTH, "UseRTH"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestScannerParameters()
Const ProcName As String = "RequestScannerParameters"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_SCANNER_PARAMETERS
mWriter.AddString VERSION, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestScannerSubscription(ByVal pTickerId As Long, ByVal pSubscription As TwsScannerSubscription)
Const ProcName As String = "RequestScannerSubscription"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 3

mWriter.AddMessageId REQ_SCANNER_SUBSCRIPTION
mWriter.AddString VERSION, "Version"
mWriter.AddString pTickerId, "Ticker Id"
mWriter.AddLongMax pSubscription.NumberOfRows, "NumberOfRows"
mWriter.AddString pSubscription.Instrument, "Instrument"
mWriter.AddString pSubscription.LocationCode, "LocationCode"
mWriter.AddString pSubscription.ScanCode, "ScanCode"
mWriter.AddDoubleMax pSubscription.AbovePrice, "AbovePrice"
mWriter.AddDoubleMax pSubscription.BelowPrice, "BelowPrice"
mWriter.AddLongMax pSubscription.AboveVolume, "AboveVolume"
mWriter.AddDoubleMax pSubscription.MarketCapAbove, "MarketCapAbove"
mWriter.AddDoubleMax pSubscription.MarketCapBelow, "MarketCapBelow"
mWriter.AddString pSubscription.MoodyRatingAbove, "MoodyRatingAbove"
mWriter.AddString pSubscription.MoodyRatingBelow, "MoodyRatingBelow"
mWriter.AddString pSubscription.SpRatingAbove, "SpRatingAbove"
mWriter.AddString pSubscription.SpRatingBelow, "SpRatingBelow"
mWriter.AddString pSubscription.MaturityDateAbove, ""
mWriter.AddString pSubscription.MaturityDateBelow, "MaturityDateBelow"
mWriter.AddDoubleMax pSubscription.CouponRateAbove, "CouponRateAbove"
mWriter.AddDoubleMax pSubscription.CouponRateBelow, "CouponRateBelow"
mWriter.AddString pSubscription.ExcludeConvertible, "ExcludeConvertible"
mWriter.AddLongMax pSubscription.AverageOptionVolumeAbove, "AverageOptionVolumeAbove"
mWriter.AddString pSubscription.ScannerSettingPairs, "ScannerSettingPairs"
mWriter.AddString pSubscription.StockTypeFilter, "StockTypeFilter"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub SetTwsLogLevel(ByVal pLogLevel As TwsLogLevels)
mLogLevel = pLogLevel
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub cancelHistDataRequest( _
                ByVal pRequestId As Long)
Const ProcName As String = "cancelHistDataRequest"
On Error GoTo Err

If mConnectionState = TwsConnConnected Then
    Const VERSION As Long = 1

    mWriter.AddMessageId CANCEL_HISTORICAL_DATA
    mWriter.AddString VERSION, "Version"
    mWriter.AddString getTwsHistRequestIdFromCallersRequestId(pRequestId), "Request id"
    mWriter.Send
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub cancelMarketDataEx(ByVal pTickerId As Long)
Const ProcName As String = "cancelMarketDataEx"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 2

mWriter.AddMessageId CANCEL_MKT_DATA
mWriter.AddString VERSION, "Version"
mWriter.AddString getTwsMarketDataRequestIdFromCallersRequestId(pTickerId), "Ticker id"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Function getTwsContractRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersContractRequestId, "Max request id is " & MaxCallersContractRequestId
getTwsContractRequestIdFromCallersRequestId = pId + BaseContractRequestId
End Function

Private Function getTwsExecutionsRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersExecutionsRequestId, "Max request id is " & MaxCallersExecutionsRequestId
getTwsExecutionsRequestIdFromCallersRequestId = pId + BaseExecutionsRequestId
End Function

Private Function getTwsHistRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersHistoricalDataRequestId, "Max request id is " & MaxCallersHistoricalDataRequestId
getTwsHistRequestIdFromCallersRequestId = pId + BaseHistoricalDataRequestId
End Function

Private Function getTwsMarketDataRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersMarketDataRequestId, "Max request id is " & MaxCallersMarketDataRequestId
getTwsMarketDataRequestIdFromCallersRequestId = pId + BaseMarketDataRequestId
End Function

Private Function getTwsMarketDepthRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersMarketDepthRequestId, "Max request id is " & MaxCallersMarketDepthRequestId
getTwsMarketDepthRequestIdFromCallersRequestId = pId + BaseMarketDepthRequestId
End Function

Private Function legOpenCloseToString(ByVal Value As TwsLegOpenCloseCodes) As String
Select Case Value
Case TwsLegOpenCloseCodes.TwsLegOpenCloseSame, _
        TwsLegOpenCloseCodes.TwsLegOpenCloseOpen, _
        TwsLegOpenCloseCodes.TwsLegOpenCloseClose, _
        TwsLegOpenCloseCodes.TwsLegOpenCloseUnknown
    legOpenCloseToString = CStr(Value)
Case Else
    Assert False, "Invalid leg open/close value"
End Select
End Function

Private Sub reqContractDetails( _
                ByVal pRequestId As Long, _
                ByVal pContract As TwsContract)
Const ProcName As String = "reqContractDetails"
On Error GoTo Err

Const VERSION = 7

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
AssertArgument pContract.Sectype <> TwsSecTypeCombo, "Combo contracts are not supported"
AssertArgument mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Or pContract.TradingClass = "", "TradingClass property not supported"

gLog "Requesting contract details: request id=" & pRequestId & "; contract=" & pContract.ToString, ModuleName, ProcName

mWriter.AddMessageId REQ_CONTRACT_DATA
mWriter.AddString VERSION, "Version"

mWriter.AddString getTwsContractRequestIdFromCallersRequestId(pRequestId), "Request id"

With pContract
    mWriter.AddString .ConId, "Contract id"
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    mWriter.AddString IIf(.Multiplier = 1, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString IIf(.Exchange = "", "*", .Exchange), "Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    
    mWriter.AddString UCase$(.LocalSymbol), "Local Symbol"
    If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString .TradingClass, "Trading Class"
    mWriter.AddBoolean True, "Include expired"
    mWriter.AddString .SecIdType, "SecIdType"
    mWriter.AddString .SecId, "SecId"
End With
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub reqHistoricalData( _
                ByVal pRequestId As Long, _
                ByRef pRequest As TwsHistoricalDataRequest)
Const ProcName As String = "reqHistoricalData"
On Error GoTo Err

Const VERSION = 5

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
Assert mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Or (pRequest.contract.TradingClass = "" And pRequest.contract.ConId = 0), "ConId and TradingClass properties not supported"

Dim lTwsId As Long
lTwsId = getTwsHistRequestIdFromCallersRequestId(pRequestId)

gLog "Requesting historical data for: " & UCase$(pRequest.contract.LocalSymbol) & _
            "; id=" & lTwsId & _
            "; barsize=" & pRequest.BarSizeSetting & _
            "; endTime=" & pRequest.EndDateTime & _
            "; duration=" & pRequest.Duration, _
            ModuleName, ProcName

mWriter.AddMessageId REQ_HISTORICAL_DATA
mWriter.AddString VERSION, "Version"
mWriter.AddString lTwsId, "Request id"

With pRequest.contract
    If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString .ConId, "Contract id"
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    mWriter.AddString IIf(.Multiplier = 1, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString .PrimaryExch, "Primary Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    
    Dim lExpired  As Boolean
    lExpired = gContractHasExpired(pRequest.contract)
    mWriter.AddString IIf(lExpired, "", UCase$(.LocalSymbol)), "Local Symbol"
    mWriter.AddString .TradingClass, "Trading Class"
    mWriter.AddString IIf(lExpired, 1, 0), "Include expired"   ' can't include expired for non-expiring contracts
    
End With

mWriter.AddString pRequest.EndDateTime, "End date"
mWriter.AddString pRequest.BarSizeSetting, "Bar Size"

mWriter.AddString pRequest.Duration, "Duration"
mWriter.AddString 0, "Use RTH"
mWriter.AddString pRequest.WhatToShow, "What to show"

mWriter.AddString TwsHistDataDateFormats.DateFormatString, "Date format"

If pRequest.contract.Sectype = TwsSecTypeCombo Then
    mWriter.AddString pRequest.contract.ComboLegs.Count, "Combo legs count"
    Dim lComboLeg As TwsComboLeg
    Dim i As Long
    For Each lComboLeg In pRequest.contract.ComboLegs
        With lComboLeg
            i = i + 1
            mWriter.AddString .ConId, "ConId" & i
            mWriter.AddString .Ratio, "Ratio" & i
            mWriter.AddString gTwsOrderActionToString(.Action), "Action" & i
            mWriter.AddString .Exchange, "Exchange" & i
        End With
    Next
End If

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub reqMarketData( _
                ByVal pTickerId As Long, _
                ByVal pContract As TwsContract, _
                ByVal pGenericTicks As String, _
                ByVal pSnapshot As Boolean)
Const ProcName As String = "reqMarketData"
On Error GoTo Err

Const VERSION = 10

gLog "Requesting market data: ticker id=" & pTickerId & "; contract=" & pContract.ToString, ModuleName, ProcName

Assert mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Or pContract.TradingClass = "", "TradingClass property not supported"

mWriter.AddMessageId REQ_MKT_DATA
mWriter.AddString VERSION, "Version"
mWriter.AddString getTwsMarketDataRequestIdFromCallersRequestId(pTickerId), "Ticker id"
With pContract
    mWriter.AddString .ConId, "Con id"
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    mWriter.AddString IIf(.Multiplier = 1, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString .PrimaryExch, "Primary Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    mWriter.AddString UCase$(.LocalSymbol), "Local Symbol"
    If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString .TradingClass, "Trading Class"
    
    ' Add combo legs for BAG requests
    If .Sectype = TwsSecTypes.TwsSecTypeCombo Then
        mWriter.AddString .ComboLegs.Count, "Combolegs count"
        Dim comboLeg As TwsComboLeg
        Dim i As Long
        For Each comboLeg In .ComboLegs
            With comboLeg
                i = i + 1
                mWriter.AddString .ConId, "ConId" & i
                mWriter.AddString .Ratio, "Ratio" & i
                mWriter.AddString gTwsOrderActionToString(.Action), "Action" & i
                mWriter.AddString .Exchange, "Exchange" & i
            End With
        Next
    End If
    
    If .UnderComp.ConId <> 0 Then
        mWriter.AddBoolean True, "Under comp"
        mWriter.AddString .UnderComp.ConId, "Under comp conid"
        mWriter.AddString .UnderComp.Delta, "Under comp delta"
        mWriter.AddString .UnderComp.Price, "Under comp price"
    Else
        mWriter.AddBoolean False, "Under comp"
    End If
    
    mWriter.AddString pGenericTicks, "Generic tick list"
    
    mWriter.AddBoolean pSnapshot, "Snapshot"
    
    mWriter.Send
End With

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName

End Sub

Private Sub reqMarketDepth(ByVal TickerId As Long, _
                        ByVal pContract As TwsContract, _
                        Optional ByVal pNumberOfRows As Long = 20)
Const ProcName As String = "reqMarketDepth"
On Error GoTo Err

AssertArgument mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Or (pContract.TradingClass = "" And pContract.ConId = 0), "ContractId and TradingClass properties are not supported"

Const VERSION = 4

gLog "Requesting market depth: ticker id=" & TickerId & "; contract=" & pContract.ToString, ModuleName, ProcName

mWriter.AddMessageId REQ_MKT_DEPTH
mWriter.AddString VERSION, "Version"
mWriter.AddString getTwsMarketDepthRequestIdFromCallersRequestId(TickerId), "Request id"

If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString pContract.ConId, "Contract Id"

With pContract
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    mWriter.AddString IIf(.Multiplier = 1, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    mWriter.AddString UCase$(.LocalSymbol), "Local Symbol"
    If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString pContract.TradingClass, "Trading Class"
    mWriter.AddString pNumberOfRows, "Num rows"
End With
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub sendLogLevel()
Const ProcName As String = "sendLogLevel"
On Error GoTo Err

Const VERSION = 1

mWriter.AddMessageId SET_SERVER_LOGLEVEL
mWriter.AddString VERSION, "Version"
mWriter.AddString mLogLevel, "Loglevel"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub setConnectionState(ByVal pState As TwsConnectionStates, ByVal pMessage As String)
Const ProcName As String = "setConnectionState"
On Error GoTo Err

mConnectionState = pState
If Not mConnectionStatusConsumer Is Nothing Then mConnectionStatusConsumer.NotifyAPIConnectionStateChange mConnectionState, pMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub startPerformanceTLIPeriod()
Const ProcName As String = "startPerformanceTLIPeriod"
On Error GoTo Err

Set mPerformanceTLIPeriod = GetGlobalTimerList.Add(Nothing, _
                                        CDate(Int((GetTimestamp + OneMinute + OneSecond) / OneMinute) * OneMinute), _
                                        ExpiryTimeUnitDateTime)
Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub startPerformanceTLISecond()
Const ProcName As String = "startPerformanceTLISecond"
On Error GoTo Err

Set mPerformanceTLISecond = GetGlobalTimerList.Add(Nothing, _
                                        CDate(Int((GetTimestamp + 2 * OneSecond) / OneSecond) * OneSecond), _
                                        ExpiryTimeUnitDateTime)
    
Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub startPerformanceLogging()
Const ProcName As String = "startPerformanceLogging"
On Error GoTo Err

If Not gLogger.IsLoggable(LogLevelDetail) Then Exit Sub
startPerformanceTLISecond
startPerformanceTLIPeriod

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub stopPerformanceLogging()
Set mStatsRecorder = Nothing

If mPerformanceTLIPeriod Is Nothing Then Exit Sub

mPerformanceTLIPeriod.Cancel
Set mPerformanceTLIPeriod = Nothing

mPerformanceTLISecond.Cancel
Set mPerformanceTLISecond = Nothing
End Sub

Private Function stopTriggerMethodToString(ByVal Value As TwsStopTriggerMethods) As String
Select Case Value
Case TwsStopTriggerMethods.TwsStopTriggerBidAsk, _
        TwsStopTriggerMethods.TwsStopTriggerDefault, _
        TwsStopTriggerMethods.TwsStopTriggerDoubleBidAsk, _
        TwsStopTriggerMethods.TwsStopTriggerDoubleLast, _
        TwsStopTriggerMethods.TwsStopTriggerLast, _
        TwsStopTriggerMethods.TwsStopTriggerLastOrBidAsk, _
        TwsStopTriggerMethods.TwsStopTriggerMidPoint
    stopTriggerMethodToString = CStr(Value)
Case Else
    AssertArgument False, "Invalid stop trigger method"
End Select
End Function

