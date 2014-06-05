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

' Client version history
'
'  6 = Added parentId to orderStatus
'  7 = The new execDetails event returned for an order filled status and reqExecDetails
'     Also market depth is available.
'  8 = Added lastFillPrice to orderStatus() event and permId to execution details
'  9 = Added 'averageCost', 'unrealizedPNL', and 'unrealizedPNL' to updatePortfolio event
' 10 = Added 'serverId' to the 'open order' & 'order status' events.
'      We send back all the API open orders upon connection.
'      Added new methods reqAllOpenOrders, reqAutoOpenOrders()
'      Added FA support - reqExecution has filter.
'                       - reqAccountUpdates takes acct code.
' 11 = Added permId to openOrder event.
' 12 = requsting open order attributes ignoreRth, hidden, and discretionary
' 13 = added goodAfterTime
' 14 = always send size on bid/ask/last tick
' 15 = send allocation description string on openOrder
' 16 = can receive account name in account and portfolio updates, and fa params in openOrder
' 17 = can receive liquidation field in exec reports, and notAutoAvailable field in mkt data
' 18 = can receive good till date field in open order messages, and request intraday backfill
' 19 = can receive rthOnly flag in ORDER_STATUS
' 20 = expects TWS time string on connection after server version >= 20.
' 21 = can receive bond contract details.
' 22 = can receive price magnifier in version 2 contract details message
' 23 = support for scanner
' 24 = can receive volatility order parameters in open order messages
' 25 = can receive HMDS query start and end times
' 26 = can receive option vols in option market data messages
' 27 = can receive delta neutral order type and delta neutral aux price in place order version 20: API 8.85
' 28 = can receive option model computation ticks: API 8.9
' 29 = can receive trail stop limit price in open order and can place them: API 8.91
' 30 = can receive extended bond contract def, new ticks, and trade count in bars
' 31 = can receive EFP extensions to scanner and market data, and combo legs on open orders
'    ; can receive RT bars
' 32 = can receive TickType.LAST_TIMESTAMP
'    ; can receive "whyHeld" in order status messages
' 33 = can receive ScaleNumComponents and ScaleComponentSize is open order messages
' 34 = can receive whatIf orders / order state
' 35 = can receive contId field for Contract objects
' 36 = can receive outsideRth field for Order objects
' 37 = can receive clearingAccount and clearingIntent for Order objects
' 38 = can receive multiplier and primaryExchange in portfolio updates
'    ; can receive cumQty and avgPrice in execution
'    ; can receive fundamental data
'    ; can receive underComp for Contract objects
'    ; can receive reqId and end marker in contractDetails/bondContractDetails
'    ; can receive ScaleInitComponentSize and ScaleSubsComponentSize for Order objects
' 39 = can receive underConId in contractDetails
' 40 = can receive algoStrategy/algoParams in openOrder
' 41 = can receive end marker for openOrder
'    ; can receive end marker for account download
'    ; can receive end marker for executions download
' 42 = can receive deltaNeutralValidation
' 43 = can receive longName(companyName)
'    ; can receive listingExchange
'    ; can receive RTVolume tick
' 44 = can receive end market for ticker snapshot
' 45 = can receive notHeld field in openOrder
' 46 = can receive contractMonth, industry, category, subcategory fields in contractDetails
'    ; can receive timeZoneId, tradingHours, liquidHours fields in contractDetails
' 47 = can receive gamma, vega, theta, undPrice fields in TICK_OPTION_COMPUTATION
' 48 = can receive exemptCode in openOrder
' 49 = can receive hedgeType and hedgeParam in openOrder
' 50 = can receive optOutSmartRouting field in openOrder
' 51 = can receive smartComboRoutingParams in openOrder
' 52 = can receive deltaNeutralConId, deltaNeutralSettlingFirm, deltaNeutralClearingAccount and deltaNeutralClearingIntent in openOrder
' 53 = can receive orderRef in execution
' 54 = can receive scale order fields (PriceAdjustValue, PriceAdjustInterval, ProfitOffset, AutoReset,
'      InitPosition, InitFillQty and RandomPercent) in openOrder
' 55 = can receive orderComboLegs (price) in openOrder
' 56 = can receive trailingPercent in openOrder
' 57 = can receive commissionReport message
' 58 = can receive CUSIP/ISIN/etc. in contractDescription/bondContractDescription
' 59 = can receive evRule, evMultiplier in contractDescription/bondContractDescription/executionDetails
'      can receive multiplier in executionDetails
' 60 = can receive deltaNeutralOpenClose, deltaNeutralShortSale, deltaNeutralShortSaleSlot and deltaNeutralDesignatedLocation in openOrder
' 61 = can receive multiplier in openOrder
'      can receive tradingClass in openOrder, updatePortfolio, execDetails and POSITION
' 62 = can receive avgCost in POSITION message

Private Const ClientVersion                 As Long = 62

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

Private mProgramErrorHandler                        As ProgramErrorListener
Private mErrorAndNotificationConsumer               As IErrorAndNotificationConsumer
Private mConnectionStatusConsumer                   As IConnectionStatusConsumer

Private WithEvents mSocketHandler                   As SocketHandler
Attribute mSocketHandler.VB_VarHelpID = -1

Private mWriter                                     As BufferedWriter

Private WithEvents mInMessageHandler                As InputMessageHandler
Attribute mInMessageHandler.VB_VarHelpID = -1

Private mServerVersion                              As Long

Private mConnectionState                            As TwsConnectionStates

Private WithEvents mPerformanceTimerSecond          As IntervalTimer
Attribute mPerformanceTimerSecond.VB_VarHelpID = -1
Private WithEvents mPerformanceTimerPeriod          As IntervalTimer
Attribute mPerformanceTimerPeriod.VB_VarHelpID = -1

Private mPerformanceLogger                          As Logger

Private mStatsRecorder                              As PerformanceStatsRecorder

Private mConnectionRetryIntervalSecs                As Long

Private mLogLevel                                   As TwsLogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

createLoggers

Set mInMessageHandler = New InputMessageHandler
Set mStatsRecorder = New PerformanceStatsRecorder

mConnectionState = TwsConnectionStates.TwsConnNotConnected

mLogLevel = 0

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub Class_Terminate()
Debug.Print "TwsAPI terminated"
End Sub

'================================================================================
' mInMessageHandler Event Handlers
'================================================================================

Private Sub mInMessageHandler_ApiConnectionEstablished(ByVal pServerVersion As Variant)
Const ProcName As String = "mInMessageHandler_ApiConnectionEstablished"
On Error GoTo Err

mServerVersion = pServerVersion
If mLogLevel <> 0 Then sendLogLevel

setConnectionState TwsConnConnected, ConnectionString

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'================================================================================
' mPerformanceTimerSecond Event Handlers
'================================================================================

Private Sub mPerformanceTimerPeriod_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mPerformanceTimerPeriod_TimerExpired"
On Error GoTo Err

mPerformanceLogger.Log LogLevelDetail, "Socket message statistics:" & vbCrLf & mStatsRecorder.GenerateStats

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mPerformanceTimerSecond_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mPerformanceTimerSecond_TimerExpired"
On Error GoTo Err

mStatsRecorder.AccumulateStats

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

gLog "Connecting to Tws: negotiating versions and client id", ModuleName, ProcName
mWriter.AddString ClientVersion, "Client version"
mWriter.Send
mWriter.AddString mClientID, "Client id"
mWriter.Send

mInMessageHandler.Reset

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

Public Property Let ProgramErrorHandler(ByVal Value As ProgramErrorListener)
Set mProgramErrorHandler = Value
End Property

Public Property Get ProgramErrorHandler() As ProgramErrorListener
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
    mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString IIf(.Exchange = "", "*", .Exchange), "Exchange"
    mWriter.AddString .PrimaryExch, "PrimaryExch"
    mWriter.AddString .CurrencyCode, "Currency"
    mWriter.AddString .LocalSymbol, "Local Symbol"
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
    mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString IIf(.Exchange = "", "*", .Exchange), "Exchange"
    mWriter.AddString .PrimaryExch, "PrimaryExch"
    mWriter.AddString .CurrencyCode, "Currency"
    mWriter.AddString .LocalSymbol, "Local Symbol"
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
                ByVal pReqId As Long, _
                ByVal pGroup As String, _
                ByVal pTags As String)
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

Set mSocketHandler = New SocketHandler
mSocketHandler.Initialise mServer, mPort, lReader, mProgramErrorHandler
mSocketHandler.ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs

Set mWriter = New BufferedWriter
mWriter.Initialise mSocketHandler

mInMessageHandler.Initialise lReader, mWriter, mStatsRecorder, mErrorAndNotificationConsumer, mProgramErrorHandler

mSocketHandler.Connect

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub Disconnect( _
                ByVal pReason As String)
Const ProcName As String = "Disconnect"
On Error GoTo Err

mSocketHandler.Disconnect pReason
Set mSocketHandler = Nothing

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
mWriter.AddString IIf(pContract.Multiplier = 0, "", pContract.Multiplier), "Multiplier"
mWriter.AddString pContract.Exchange, "Exchange"
mWriter.AddString pContract.CurrencyCode, "Currency"
mWriter.AddString pContract.LocalSymbol, "LocalSymbol"
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
                ByVal pClientId As Long)
AssertArgument pPort > 0, "Port must be > 0"
AssertArgument pClientId >= 0, "ClientID must be >= 0"

If pServer = "" Then pServer = "127.0.0.1"
mServer = pServer
mPort = pPort
mClientID = pClientId
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
    mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier" ' multiplier - this is the number of units of the underlying
                                                            ' delivered per contract - see http://www.interactivebrokers.com/discus/messages/2/27413.html
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString .PrimaryExch, "Primary Exchange"    ' only relevant when
                                                        ' Exchange is SMART and there are SMART routers in more
                                                        ' than one country (eg try IBM in Tws)
    mWriter.AddString .CurrencyCode, "Currency"
    mWriter.AddString .LocalSymbol, "Local Symbol"
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
                    
Public Sub ProcessRawTwsMessage(ByRef pMessage() As Byte, ByVal pLog As Boolean)
Const ProcName As String = "ProcessRawTwsMessage"
On Error GoTo Err

Dim lReader As New BufferedReader
Dim lInMessageHandler As New InputMessageHandler

lReader.AddData pMessage
lInMessageHandler.Initialise lReader, mWriter, mStatsRecorder, mErrorAndNotificationConsumer, mProgramErrorHandler
lInMessageHandler.AccountDataConsumer = mInMessageHandler.AccountDataConsumer
lInMessageHandler.ConnectionStatusConsumer = mInMessageHandler.ConnectionStatusConsumer
lInMessageHandler.ContractDetailsConsumer = mInMessageHandler.ContractDetailsConsumer
lInMessageHandler.HistDataConsumer = mInMessageHandler.HistDataConsumer
lInMessageHandler.MarketDataConsumer = mInMessageHandler.MarketDataConsumer
lInMessageHandler.MarketDepthConsumer = mInMessageHandler.MarketDepthConsumer
lInMessageHandler.OrderInfoConsumer = mInMessageHandler.OrderInfoConsumer
lInMessageHandler.ScannerDataConsumer = mInMessageHandler.ScannerDataConsumer
lInMessageHandler.ProcessBuffer pLog

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
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

If mServerVersion >= 9 Then mWriter.AddString acctCode, "Account code"

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

If mServerVersion >= 9 Then
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
End If

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestFA(ByVal DataType As TwsFADataTypes)
Const ProcName As String = "RequestFA"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

Assert mServerVersion >= 13, "Server version does not support RequestFA", ErrorCodes.ErrUnsupportedOperationException

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
mWriter.AddString pContract.LocalSymbol, "LocalSymbol"
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
mWriter.AddString IIf(pContract.Multiplier = 0, "", pContract.Multiplier), "Multiplier"
mWriter.AddString pContract.Exchange, "Exchange"
mWriter.AddString pContract.PrimaryExch, "PrimaryExch"
mWriter.AddString pContract.CurrencyCode, "Currency"
mWriter.AddString pContract.LocalSymbol, "LocalSymbol"
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

Private Sub createLoggers()

Const ProcName As String = "createLoggers"
On Error GoTo Err

If gLogger.IsLoggable(LogLevelDetail) Then
    If mPerformanceTimerSecond Is Nothing Then
        Set mPerformanceTimerSecond = CreateIntervalTimer(CDate(Int((Now + 2 * OneSecond) / OneSecond) * OneSecond), _
                                                ExpiryTimeUnitDateTime, _
                                                1000)
        Set mPerformanceTimerPeriod = CreateIntervalTimer(CDate(Int((Now + OneMinute) / OneMinute) * OneMinute), _
                                                ExpiryTimeUnitDateTime, _
                                                60000)
        mPerformanceTimerSecond.StartTimer
        mPerformanceTimerPeriod.StartTimer
        
        Set mPerformanceLogger = GetLogger("tradebuild.log.ibapi.performance")
    End If
End If

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
    mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString IIf(.Exchange = "", "*", .Exchange), "Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    
    mWriter.AddString .LocalSymbol, "Local Symbol"
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

gLog "Requesting historical data for: " & pRequest.contract.LocalSymbol & _
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
    mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString .PrimaryExch, "Primary Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    
    Dim lExpired  As Boolean
    lExpired = gContractHasExpired(pRequest.contract)
    mWriter.AddString IIf(lExpired, "", .LocalSymbol), "Local Symbol"
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

gLog "Requesting market data for: " & pContract.ToString, ModuleName, ProcName

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
    mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString .PrimaryExch, "Primary Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    mWriter.AddString .LocalSymbol, "Local Symbol"
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
    
    If mServerVersion >= 31 Then
         ' Note: Even though SHORTABLE tick type supported only
         '       starting server version 33 it would be relatively
         '       expensive to expose this restriction here.
         '
         '       Therefore we are relying on Tws doing validation.
         '
        mWriter.AddString pGenericTicks, "Generic tick list"
    End If
    
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
    If mServerVersion >= 15 Then mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    mWriter.AddString .LocalSymbol, "Local Symbol"
    If mServerVersion >= MIN_SERVER_VER_TRADING_CLASS Then mWriter.AddString pContract.TradingClass, "Trading Class"
    If mServerVersion >= 19 Then mWriter.AddString pNumberOfRows, "Num rows"
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

