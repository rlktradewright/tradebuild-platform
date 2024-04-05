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

Private CompeteAgainstBestOffsetUpToMid             As Variant

Private mServer                                     As String
Private mPort                                       As String

Private mClientID                                   As Long
Private mOptionalCapabilities                       As String

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

Private mLogLevel                                   As TwsLoggingLevels

Private mLogApiMessages                             As TwsApiMessageLoggingOptions
Private mLogRawApiMessages                          As TwsApiMessageLoggingOptions
Private mLogApiMessageStats                         As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

CompeteAgainstBestOffsetUpToMid = MinDouble

Set mInMessageHandler = New InputMessageHandler

mConnectionState = TwsConnectionStates.TwsConnNotConnected

mLogLevel = 0

Set mStatsRecorder = New PerformanceStatsRecorder
Set mPerformanceLogger = GetLogger("tradebuild.log.ibapi.performance")

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
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
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
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
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
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
If Not mPerformanceLogger.IsLoggable(LogLevelNormal) Then Exit Sub

mPerformanceLogger.Log LogLevelNormal, "Socket message statistics:" & vbCrLf & mStatsRecorder.GenerateStats
startPerformanceTLIPeriod

Exit Sub

Err:
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
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
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'================================================================================
' mSocketHandler Event Handlers
'================================================================================

Private Sub mSocketHandler_Connected()
Const ProcName As String = "mSocketHandler_Connected"
On Error GoTo Err

Set mConnectionEstablishedTLI = GetGlobalTimerList.Add(Nothing, 30, ExpiryTimeUnitSeconds)
mInMessageHandler.Reset

mConnectionNegotiator.NegotiateApiConnection

Exit Sub

Err:
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_ConnectFailed( _
                ByVal pDescription As String, _
                ByVal pRetryInterval As Long)
Const ProcName As String = "mSocketHandler_ConnectFailed"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnFailed, pDescription & ": " & ConnectionString

Exit Sub

Err:
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_Connecting()
Const ProcName As String = "mSocketHandler_Connecting"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnConnecting, ConnectionString

Exit Sub

Err:
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_ConnectionClosed( _
                ByVal pDescription As String)
Const ProcName As String = "mSocketHandler_ConnectionClosed"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnNotConnected, pDescription & ": " & ConnectionString

Exit Sub

Err:
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_Disconnected(ByVal pReason As String)
Const ProcName As String = "mSocketHandler_Disconnected"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnNotConnected, pReason
stopPerformanceLogging

Exit Sub

Err:
GIB.NotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
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
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
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

Public Property Get IsTwsConnectedToIBServers() As Boolean
IsTwsConnectedToIBServers = mInMessageHandler.IsTwsConnectedToIBServers
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

Public Property Get MaxAccountRequestId() As Long
MaxAccountRequestId = GIdManager.MaxCallersAccountRequestId
End Property

Public Property Get MaxContractRequestId() As Long
MaxContractRequestId = GIdManager.MaxCallersContractRequestId
End Property

Public Property Get MaxExecutionsRequestId() As Long
MaxExecutionsRequestId = GIdManager.MaxCallersExecutionsRequestId
End Property

Public Property Get MaxHistoricalDataRequestId() As Long
MaxHistoricalDataRequestId = GIdManager.MaxCallersHistoricalDataRequestId
End Property

Public Property Get MaxMarketDataRequestId() As Long
MaxMarketDataRequestId = GIdManager.MaxCallersMarketDataRequestId
End Property

Public Property Get MaxMarketDepthRequestId() As Long
MaxMarketDepthRequestId = GIdManager.MaxCallersMarketDepthRequestId
End Property

Public Property Get MaxOrderId() As Long
MaxOrderId = &H7FFFFFFF
End Property

Public Property Get NextOrderId() As Long
NextOrderId = GIdManager.GetNextOrderId
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
        ByVal pContractSpec As TwsContractSpecifier, _
        ByVal pOptionPrice As Double, _
        ByVal pUnderPrice As Double)
Const ProcName As String = "CalculateImpliedVolatility"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 3

mWriter.AddMessageId REQ_CALC_IMPLIED_VOLAT
mWriter.AddLong VERSION, "Version"
mWriter.AddLong pReqId, "ReqId"
mWriter.AddContract pContractSpec
mWriter.AddDouble pOptionPrice, "OptionPrice"
mWriter.AddDouble pUnderPrice, "UnderPrice"

mWriter.AddString "0", "Options Count"
mWriter.AddString "", "Options"

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CalculateOptionPrice( _
        ByVal pReqId As Long, _
        ByVal pContractSpec As TwsContractSpecifier, _
        ByVal pVolatility As Double, _
        ByVal pUnderPrice As Double)
Const ProcName As String = "CalculateOptionPrice"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 3

mWriter.AddMessageId REQ_CALC_OPTION_PRICE
mWriter.AddLong VERSION, "Version"
mWriter.AddLong pReqId, "ReqId"
mWriter.AddContract pContractSpec

mWriter.AddDouble pVolatility, "Volatility"
mWriter.AddDouble pUnderPrice, "UnderPrice"

mWriter.AddLong 0, "Options Count"
mWriter.AddString "", "Options"

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelAccountSummary( _
                ByVal pReqId As Long)
Const ProcName As String = "CancelAccountSummary"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId CANCEL_ACCOUNT_SUMMARY
mWriter.AddLong VERSION, "Version"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelAccountUpdatesMulti(ByVal pRequestId As Long)
Const ProcName As String = "CancelAccountUpdatesMulti"
On Error GoTo Err

Const VERSION = 1

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

mWriter.AddMessageId TwsSocketOutMsgTypes.CancelAccountUpdatesMulti
mWriter.AddLong VERSION, "Version"

mWriter.AddLong pRequestId, "RequestId"

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelCalculateImpliedVolatility(ByVal pReqId As Long)
Const ProcName As String = "CancelCalculateImpliedVolatility"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 1

mWriter.AddMessageId CANCEL_CALC_IMPLIED_VOLAT
mWriter.AddLong VERSION, "Version"
mWriter.AddLong pReqId, "ReqId"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelCalculateOptionPrice(ByVal pReqId As Long)
Const ProcName As String = "CancelCalculateOptionPrice"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 1

mWriter.AddMessageId CANCEL_CALC_OPTION_PRICE
mWriter.AddLong VERSION, "Version"
mWriter.AddLong pReqId, "ReqId"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelFundamentalData(ByVal pReqId As Long)
Const ProcName As String = "CancelFundamentalData"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 1

mWriter.AddMessageId CANCEL_FUNDAMENTAL_DATA
mWriter.AddLong VERSION, "Version"
mWriter.AddLong pReqId, "ReqId"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelHeadTimestamp(ByVal pRequestId As Long)
Const ProcName As String = "CancelHeadTimestamp"
On Error GoTo Err

If ConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub
AssertArgument ServerVersion > ApiServerVersions.CANCEL_HEADTIMESTAMP, "CancelHeadTimestamp not supported"

mWriter.AddMessageId TwsSocketOutMsgTypes.CancelHeadTimestamp

mWriter.AddLong pRequestId, "Request Id"

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelHistogramData(ByVal pRequestId As Long)
Const ProcName As String = "CancelHistogramData"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub
AssertArgument ServerVersion > ApiServerVersions.REQ_HISTOGRAM_DATA, "CancelHistogramData not supported"

mWriter.AddMessageId TwsSocketOutMsgTypes.CancelHistogramData
mWriter.AddLong pRequestId, "Request Id"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelHistoricalData( _
                ByVal pRequestId As Long)
Const ProcName As String = "CancelHistoricalData"
On Error GoTo Err

If mConnectionState = TwsConnConnected Then
    Const VERSION As Long = 1

    mWriter.AddMessageId CANCEL_HISTORICAL_DATA
    mWriter.AddLong VERSION, "Version"
    mWriter.AddString GIdManager.GetTwsId(pRequestId, IdTypeHistoricalData), "RequestID"
    mWriter.Send
End If

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelMarketData(ByVal pTickerId As Long)
Const ProcName As String = "CancelMarketData"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 2

mWriter.AddMessageId CANCEL_MKT_DATA
mWriter.AddLong VERSION, "Version"
mWriter.AddString GIdManager.GetTwsId(pTickerId, IdTypeMarketData), "TickerId"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelMarketDepth(ByVal pTickerId As Long, pIsSmartDepth As Boolean)
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub
AssertArgument Not (pIsSmartDepth And ServerVersion < ApiServerVersions.SMART_DEPTH), "SMART depth cancel is not supported"

Const VERSION = 1

mWriter.AddMessageId CANCEL_MKT_DEPTH
mWriter.AddLong VERSION, "Version"
mWriter.AddString GIdManager.GetTwsId(pTickerId, IdTypeMarketDepth), "Ticker id"
If ServerVersion >= ApiServerVersions.SMART_DEPTH Then mWriter.AddBoolean pIsSmartDepth, "IsSmartDepth"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub CancelNewsBulletins()
Const ProcName As String = "CancelNewsBulletins"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const VERSION = 1

mWriter.AddMessageId CANCEL_NEWS_BULLETINS
mWriter.AddLong VERSION, "Version"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelOrder(ByVal pOrderId As Long, pManualOrderCancelTime As String)
Const ProcName As String = "CancelOrder"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1
mWriter.AddMessageId CANCEL_ORDER
mWriter.AddLong VERSION, "Version"
mWriter.AddLong pOrderId, "Order id"
If ServerVersion >= ApiServerVersions.MANUAL_ORDER_TIME Then mWriter.AddString pManualOrderCancelTime, "manualOrderCancelTime"
mWriter.Send True

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelPnL(ByVal pRequestId As Long)
Const ProcName As String = "CancelPnL"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
AssertArgument ServerVersion >= ApiServerVersions.PNL, "CancelPnL not supported"

mWriter.AddMessageId TwsSocketOutMsgTypes.CancelPnL
mWriter.AddLong pRequestId, "Request Id"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelPnLSingle(ByVal pRequestId As Long)
Const ProcName As String = "CancelPnLSingle"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
AssertArgument ServerVersion >= ApiServerVersions.PNL, "CancelPnLSingle not supported"

mWriter.AddMessageId TwsSocketOutMsgTypes.CancelPnLSingle
mWriter.AddLong pRequestId, "Request Id"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelPositions()
Const ProcName As String = "CancelPositions"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId CANCEL_POSITIONS
mWriter.AddLong VERSION, "Version"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelPositionsMulti(ByVal pRequestId As Long)
Const ProcName As String = "CancelPositionsMulti"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
AssertArgument ServerVersion >= ApiServerVersions.MODELS_SUPPORT, "CancelPositionsMulti not supported"

Const VERSION = 1

mWriter.AddMessageId TwsSocketOutMsgTypes.CancelPositionsMulti
mWriter.AddLong VERSION, "Version"
mWriter.AddLong pRequestId, "Request Id)"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelRealtimeBars(ByVal pTickerId As Long)
Const ProcName As String = "CancelRealTimeBars"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId CANCEL_REAL_TIME_BARS
mWriter.AddLong VERSION, "Version"
mWriter.AddString GIdManager.GetTwsId(pTickerId, IdTypeHistoricalData), "Ticker id"
mWriter.Send True

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelScannerSubscription(ByVal pRequestId As Long)
Const ProcName As String = "CancelScannerSubscription"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId CANCEL_SCANNER_SUBSCRIPTION
mWriter.AddLong VERSION, "Version"

mWriter.AddLong GIdManager.GetTwsId(pRequestId, IdTypeScanner), "Request Id"
mWriter.Send True

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelTickByTickData(ByVal pRequestId As Long)
Const ProcName As String = "CancelTickByTickData"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
AssertArgument ServerVersion >= ApiServerVersions.TICK_BY_TICK, "CancelTickByTickData not supported"

mWriter.AddMessageId TwsSocketOutMsgTypes.CancelTickByTickData
mWriter.AddLong pRequestId, "Request Id"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub Connect()
Const ProcName As String = "Connect"
On Error GoTo Err

Dim lReader As New BufferedReader
lReader.Initialise mInMessageHandler, mLogApiMessages, mLogRawApiMessages

Set mSocketHandler = New SocketHandler
mSocketHandler.Initialise mServer, mPort, lReader, mProgramErrorHandler, mLogRawApiMessages
mSocketHandler.ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs

Set mWriter = New BufferedWriter
mWriter.Initialise mSocketHandler, mLogApiMessages, mLogRawApiMessages

Set mConnectionNegotiator = New ApiConnectionNegotiator
mConnectionNegotiator.Initialise lReader, mWriter

mInMessageHandler.Initialise mClientID, _
                            mOptionalCapabilities, _
                            mConnectionNegotiator, _
                            lReader, _
                            mWriter, _
                            mStatsRecorder, _
                            mErrorAndNotificationConsumer, _
                            mProgramErrorHandler, _
                            mLogApiMessageStats

mSocketHandler.Connect

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
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
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub ExerciseOptions( _
                ByVal pTickerId As Long, _
                ByVal pContractSpec As TwsContractSpecifier, _
                ByVal pExerciseAction As Long, _
                ByVal pExerciseQuantity As Long, _
                ByVal pAccount As String, _
                ByVal pOverride As Boolean)
Const ProcName As String = "ExerciseOptions"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 2

mWriter.AddMessageId EXERCISE_OPTIONS
mWriter.AddLong VERSION, "Version"
mWriter.AddLong pTickerId, "TickerId"

mWriter.AddContract pContractSpec, True

mWriter.AddLong pExerciseAction, "ExerciseAction"
mWriter.AddLong pExerciseQuantity, "ExerciseQuantity"
mWriter.AddString pAccount, "Account"
mWriter.AddBoolean pOverride, "Override"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Function GetRawTwsMessage() As Byte()
Const ProcName As String = "GetRawMessage"
On Error GoTo Err

GetRawTwsMessage = mInMessageHandler.BufferedReader.GetRawMessage

Exit Function

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pClientId As Long, _
                ByVal pLogApiMessages As TwsApiMessageLoggingOptions, _
                ByVal pLogRawApiMessages As TwsApiMessageLoggingOptions, _
                ByVal pLogApiMessageStats As Boolean)
AssertArgument pPort > 0, "Port must be > 0"

AssertArgument pClientId >= 0 And pClientId <= &H7FFFFFFF, "ClientID must be >= 0 and <= &h7fffffff"

If pServer = "" Then pServer = "127.0.0.1"
mServer = pServer
mPort = pPort
mClientID = pClientId
mLogApiMessages = pLogApiMessages
mLogRawApiMessages = pLogRawApiMessages
mLogApiMessageStats = pLogApiMessageStats
End Sub

Public Sub PlaceOrder( _
                ByVal pOrder As TwsOrder, _
                ByVal pContractSpec As TwsContractSpecifier, _
                Optional ByVal pSecIdType As String, _
                Optional ByVal pSecId As String)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

Dim i As Long

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

AssertArgument ServerVersion >= ApiServerVersions.EXT_OPERATOR Or pOrder.ExtOperator = "", "ExtOperator parameter not supported"
AssertArgument ServerVersion >= ApiServerVersions.CASH_QTY Or pOrder.CashQty = MaxDouble, "CashQty parameter not supported"

AssertArgument ServerVersion >= ApiServerVersions.ORDER_CONTAINER Or Not pOrder.IsOmsContainer, "IsOmsContainer attribute not supported"
AssertArgument ServerVersion >= ApiServerVersions.D_PEG_ORDERS Or Not pOrder.DiscretionaryUpToLimitPrice, "DiscretionaryUpToLimitPrice attribute not supported"
AssertArgument ServerVersion >= ApiServerVersions.PRICE_MGMT_ALGO Or pOrder.UsePriceMgmtAlgo, "UsePriceMgmtAlgo attribute not supported"
AssertArgument ServerVersion >= ApiServerVersions.Duration Or Not pOrder.Duration = GIB.MaxLong, "Duration attribute not supported"
AssertArgument ServerVersion >= ApiServerVersions.POST_TO_ATS Or Not pOrder.PostToAts = GIB.MaxLong, "PostToAts attribute not supported"
AssertArgument ServerVersion >= ApiServerVersions.AUTO_CANCEL_PARENT Or Not pOrder.AutoCancelParent, "AutoCancelParent attribute not supported"
AssertArgument ServerVersion >= ApiServerVersions.ADVANCED_ORDER_REJECT Or pOrder.AdvancedErrorOverride = "", "AdvancedErrorOverride attribute not supported"
AssertArgument ServerVersion >= ApiServerVersions.MANUAL_ORDER_TIME Or pOrder.ManualOrderTime = "", "AdvancedErrorOverride attribute not supported"
AssertArgument ServerVersion >= ApiServerVersions.PEGBEST_PEGMID_OFFSETS Or (pOrder.MinimumTradeQuantity = GIB.MaxLong And _
                                                                        pOrder.MinimumCompeteSize = GIB.MaxLong And _
                                                                        pOrder.CompeteAgainstBestOffset = MaxDouble And _
                                                                        pOrder.MidOffsetAtWhole = MaxDouble And _
                                                                        pOrder.MidOffsetAtHalf = MaxDouble), "PEGBEST/PEGMID-related  attributes not supported"
Const VERSION = 45

mWriter.AddMessageId PLACE_ORDER
If ServerVersion < ApiServerVersions.ORDER_CONTAINER Then mWriter.AddLong VERSION, "Version"

If pOrder.OrderId = 0 Then pOrder.OrderId = GIdManager.GetNextOrderId
AssertArgument pOrder.OrderId >= GIdManager.BaseOrderId, "Order id must not be less than " & GIdManager.BaseOrderId
If pOrder.OrderId > GIdManager.GetNextOrderId Then GIdManager.SetNextOrderId pOrder.OrderId + 1

mWriter.AddLong pOrder.OrderId, "Order id"

mWriter.AddContract pContractSpec

mWriter.AddString pSecIdType, "Sec id type"
mWriter.AddString pSecId, "Sec id"

With pOrder
    
    ' mwriter.send main Order fields
    mWriter.AddString TwsOrderActionToString(.Action), "Action"
    mWriter.AddDecimal .TotalQuantity, "Quantity"
    mWriter.AddString TwsOrderTypeToString(.OrderType), "Order type"
    
    mWriter.AddDoubleMax .LmtPrice, "Price"
    mWriter.AddDoubleMax .AuxPrice, "Aux price"
    
    ' mwriter.send extended Order fields
    If .OrderType = TwsOrderTypes.TwsOrderTypeMarketOnOpen Or .OrderType = TwsOrderTypes.TwsOrderTypeLimitOnOpen Then
        mWriter.AddString "OPG", "TIF"
    Else
        mWriter.AddString TwsOrderTIFToString(.Tif), "TIF"
    End If
    mWriter.AddString .OcaGroup, "Oca Group"
    mWriter.AddString .Account, "Account"
    mWriter.AddString .OpenClose, "OpenClose"
    mWriter.AddLong .Origin, "Origin"
    mWriter.AddString .OrderRef, "Order ref"
    mWriter.AddBoolean .Transmit, "Transmit"
    mWriter.AddLong .ParentId, "Parent id"
    mWriter.AddBoolean .BlockOrder, "Block Order"
    mWriter.AddBoolean .SweepToFill, "Sweep to fill"
    mWriter.AddLong .DisplaySize, "Display Size"
    mWriter.AddString stopTriggerMethodToString(.TriggerMethod), "Trigger method"
    mWriter.AddBoolean .OutsideRth, "Outside RTH"
    mWriter.AddBoolean .Hidden, "Hidden"

    ' send contract combo legs for BAG requests
    If pContractSpec.Sectype = TwsSecTypeCombo Then
        With pContractSpec
            If .ComboLegs.Count = 0 Then
                mWriter.AddLong 0, "Combo legs count"
            Else
                mWriter.AddLong .ComboLegs.Count, "Combo legs count"
                Dim comboLeg As TwsComboLeg
                For Each comboLeg In .ComboLegs
                    With comboLeg
                        mWriter.AddLong .ConId, "Leg " & i & " Con id"
                        mWriter.AddLong .Ratio, "Leg " & i & " Ratio"
                        mWriter.AddString TwsOrderActionToString(.Action), "Leg " & i & " Action"
                        mWriter.AddString .Exchange, "Leg " & i & " Exchange"
                        mWriter.AddString legOpenCloseToString(.OpenClose), "Leg " & i & " Open/close"
                    
                        mWriter.AddLong comboLeg.ShortSaleSlot, "Short Sale Slot"
                        mWriter.AddString comboLeg.DesignatedLocation, "Designated Location"
                        mWriter.AddLong comboLeg.ExemptCode, "Exempt Code"
                    End With
                Next
            End If
        End With

        ' Send order combo legs for BAG requests
        Dim lOrderComboLegsCount As Long
        On Error Resume Next
        lOrderComboLegsCount = UBound(pOrder.ComboLegs) + 1
        On Error GoTo Err
        
        If lOrderComboLegsCount = 0 Then
            mWriter.AddLong 0, "Order Combo Legs Count"
        Else
            Dim lOrderComboLegs() As TwsOrderComboLeg
            lOrderComboLegs = pOrder.ComboLegs
            mWriter.AddLong lOrderComboLegsCount, "Order Combo Legs Count"
    
            For i = 0 To lOrderComboLegsCount - 1
                Dim lOrderComboLeg  As TwsOrderComboLeg
                Set lOrderComboLeg = lOrderComboLegs(i)
                mWriter.AddDoubleMax lOrderComboLeg.Price, "Leg" & i & " Price"
            Next
        End If
        
        mWriter.AddTagValues pOrder.SmartComboRoutingParams, "Smart Combo Routing Params"
    End If

    mWriter.AddString "", "Shares Allocation"
    mWriter.AddDouble .DiscretionaryAmt, "Discretionary amount"
    mWriter.AddString .GoodAfterTime, "Good after Time"
    mWriter.AddString .GoodTillDate, "Good till date"
    
    mWriter.AddString .FaGroup, "FAGroup"
    mWriter.AddString .FaMethod, "FAMethod"
    mWriter.AddString .FaPercentage, "FAPercentage"
    If ServerVersion < ApiServerVersions.FA_PROFILE_DESUPPORT Then mWriter.AddString "", "FAProfile"
    
    If ServerVersion >= ApiServerVersions.MODELS_SUPPORT Then mWriter.AddString .ModelCode, "Model Code"

    'institutional short sale slot fields.
    mWriter.AddLong .ShortSaleSlot, " Short Sale Slot"           ' 0 only for retail, 1 or 2 only for institution.
    mWriter.AddString .DesignatedLocation, "Designated Location"       ' only populate when shortSaleSlot = 2.
    
    mWriter.AddLong .ExemptCode, "Exempt Code"
    
    mWriter.AddLong .OcaType, "Oca type"
    mWriter.AddString .Rule80A, "Rule 80A"
    mWriter.AddString .SettlingFirm, "Settling firm"
    mWriter.AddBoolean .AllOrNone, "All or none"
    mWriter.AddLongMax .MinQty, "Minimum quantity"
    mWriter.AddDoubleMax .PercentOffset, "Percent Offset"
    mWriter.AddBoolean False, "E-trade only"
    mWriter.AddBoolean False, "Firm quote only"
    mWriter.AddDoubleMax MaxDouble, "NBBO price cap"
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
    mWriter.AddString TwsOrderTypeToString(.DeltaNeutralOrderType), "Delta neutral Order type"
    mWriter.AddDoubleMax .DeltaNeutralAuxPrice, "Delta neutral aux price"
    If .DeltaNeutralOrderType <> TwsOrderTypeNone Then
        mWriter.AddLong .DeltaNeutralConId, "Delta Neutral Con Id"
        mWriter.AddString .DeltaNeutralSettlingFirm, "Delta Neutral Settling Firm"
        mWriter.AddString .DeltaNeutralClearingAccount, "Delta Neutral Clearing Account"
        mWriter.AddString .DeltaNeutralClearingIntent, "Delta Neutral Clearing Intent"
        mWriter.AddString .DeltaNeutralOpenClose, "Delta Neutral Open Close"
        mWriter.AddBoolean .DeltaNeutralShortSale, "Delta Neutral Short Sale"
        mWriter.AddLong .DeltaNeutralShortSaleSlot, "Delta Neutral Short Sale Slot"
        mWriter.AddString .DeltaNeutralDesignatedLocation, "Delta Neutral Designated Location"
    End If
    
    mWriter.AddLong .ContinuousUpdate, "Continuous update"
    
    mWriter.AddLongMax .ReferencePriceType, "Reference price type"
    
    mWriter.AddDoubleMax .TrailStopPrice, "Trail stop price"
    mWriter.AddDoubleMax .TrailingPercent, "Trailing Percent"
    
    mWriter.AddLongMax .ScaleInitLevelSize, "ScaleInitLevelSize"
    mWriter.AddLongMax .ScaleSubsLevelSize, "ScaleSubsLevelSize"
    mWriter.AddDoubleMax .ScalePriceIncrement, "ScalePriceIncrement"
    
    If .ScalePriceIncrement > 0# And .ScalePriceIncrement <> MaxDouble Then
        mWriter.AddDoubleMax .ScalePriceAdjustValue, "Scale Price Adjust Value"
        mWriter.AddLongMax .ScalePriceAdjustInterval, "Scale Price Adjust Interval"
        mWriter.AddDoubleMax .ScaleProfitOffset, "Scale Profit Offset"
        mWriter.AddBoolean .ScaleAutoReset, "Scale Auto Reset"
        mWriter.AddLongMax .ScaleInitPosition, "Scale Init POSITION"
        mWriter.AddLongMax .ScaleInitFillQty, "Scale Init Fill Qty"
        mWriter.AddBoolean .ScaleRandomPercent, "Scale Random Percent"
    End If

    mWriter.AddString .ScaleTable, "Scale Table"
    mWriter.AddString .ActiveStartTime, "Active Start Time"
    mWriter.AddString .ActiveStopTime, "Active Stop Time"

    mWriter.AddString TwsHedgeTypeToString(.HedgeType), "Hedge Type"
    If .HedgeType <> TwsHedgeTypeNone Then mWriter.AddString .HedgeParam, "Hedge Param"

    mWriter.AddBoolean .OptOutSmartRouting, "Opt Out Smart Routing"

    mWriter.AddString .ClearingAccount, "Clearing account"
    mWriter.AddString .ClearingIntent, "Clearing intent"

    If .OrderType = TwsOrderTypes.TwsOrderTypePeggedToBest Or .OrderType = TwsOrderTypes.TwsOrderTypePeggedToMidpoint Then
        mWriter.AddBoolean True, "Not held"
    Else
        mWriter.AddBoolean .NotHeld, "Not held"
    End If
    
    If pContractSpec.DeltaNeutralContract.ConId <> 0 Then
        mWriter.AddBoolean True, "Delta Neutral"
        mWriter.AddLong pContractSpec.DeltaNeutralContract.ConId, "Delta Neutral conid"
        mWriter.AddDouble pContractSpec.DeltaNeutralContract.Delta, "Delta Neutral delta"
        mWriter.AddDouble pContractSpec.DeltaNeutralContract.Price, "Delta Neutral price"
    Else
        mWriter.AddBoolean False, "Delta Neutral"
    End If

    mWriter.AddString .AlgoStrategy, "Algo strategy"
    If .AlgoStrategy <> "" Then
        mWriter.AddTagValues .AlgoParams, "Algo params"
    End If

    mWriter.AddString .AlgoId, "Algo Id"
    
    mWriter.AddBoolean .WhatIf, "WhatIf"
    mWriter.AddTagValues .Options, "Misc options"
    mWriter.AddBoolean .Solicited, "Solicited"
    mWriter.AddBoolean .RandomizeSize, "Randomize size"
    mWriter.AddBoolean .RandomizePrice, "Randomize price"

    If (ServerVersion >= ApiServerVersions.PEGGED_TO_BENCHMARK) Then
        If .OrderType = TwsOrderTypes.TwsOrderTypePeggedToBenchmark Then
            mWriter.AddLong .ReferenceContractId, "ReferenceContractId"
            mWriter.AddBoolean .IsPeggedChangeAmountDecrease, "IsPeggedChangeAmountDecrease"
            mWriter.AddDoubleMax .PeggedChangeAmount, "PeggedChangeAmount"
            mWriter.AddDoubleMax .ReferenceChangeAmount, "ReferenceChangeAmount"
            mWriter.AddString .ReferenceExchange, "ReferenceExchange"
        End If

        mWriter.AddLong 0, "Conditions Count" '  .Conditions.Count

'       Conditions not implemented
'        If .Conditions.Count > 0 Then
'            For Each Item In .Conditions
'                mWriter.AddLong Item.Type, "Type"
'                Item.Serialize lWriter
'            Next
'
'            mWriter.AddBoolean .ConditionsIgnoreRth, "ConditionsIgnoreRth"
'            mWriter.AddBoolean .ConditionsCancelOrder, "ConditionsCancelOrder"
'        End If

        mWriter.AddString TwsOrderTypeToString(.AdjustedOrderType), "AdjustedOrderType"
        mWriter.AddDoubleMax .TriggerPrice, "TriggerPrice"
        mWriter.AddDoubleMax .LmtPriceOffset, "LmtPriceOffset"
        mWriter.AddDoubleMax .AdjustedStopPrice, "AdjustedStopPrice"
        mWriter.AddDoubleMax .AdjustedStopLimitPrice, "AdjustedStopLimitPrice"
        mWriter.AddDoubleMax .AdjustedTrailingAmount, "AdjustedTrailingAmount"
        mWriter.AddLong .AdjustableTrailingUnit, "AdjustableTrailingUnit"
    End If

    If ServerVersion >= ApiServerVersions.EXT_OPERATOR Then
        mWriter.AddString .ExtOperator, "ExtOperator"
    End If

    If ServerVersion >= ApiServerVersions.SOFT_DOLLAR_TIER Then
        mWriter.AddString .Tier.Name, "Tier Name"
        mWriter.AddString .Tier.Value, "Tier Value"
    End If

    If ServerVersion >= ApiServerVersions.CASH_QTY Then
        mWriter.AddDoubleMax .CashQty, "Cash Qty"
    End If

    If ServerVersion >= ApiServerVersions.DECISION_MAKER Then
        mWriter.AddString .Mifid2DecisionMaker, "Mifid2DecisionMaker"
        mWriter.AddString .Mifid2DecisionAlgo, "Mifid2DecisionAlgo"
    End If

    If ServerVersion >= ApiServerVersions.MIFID_EXECUTION Then
        mWriter.AddString .Mifid2ExecutionTrader, "Mifid2ExecutionTrader"
        mWriter.AddString .Mifid2ExecutionAlgo, "Mifid2ExecutionAlgo"
    End If

    If ServerVersion >= ApiServerVersions.AUTO_PRICE_FOR_HEDGE Then
        mWriter.AddBoolean .DontUseAutoPriceForHedge, "DontUseAutoPriceForHedge"
    End If

    If ServerVersion >= ApiServerVersions.ORDER_CONTAINER Then
        mWriter.AddBoolean .IsOmsContainer, "IsOmsContainer"
    End If

    If ServerVersion >= ApiServerVersions.D_PEG_ORDERS Then
        mWriter.AddBoolean .DiscretionaryUpToLimitPrice, "DiscretionaryUpToLimitPrice"
    End If

    If ServerVersion >= ApiServerVersions.PRICE_MGMT_ALGO Then
        mWriter.AddBoolean .UsePriceMgmtAlgo, "UsePriceMgmtAlgo"
    End If

    If ServerVersion >= ApiServerVersions.Duration Then
        mWriter.AddLongMax .Duration, "Duration"
    End If

    If ServerVersion >= ApiServerVersions.POST_TO_ATS Then
        mWriter.AddLongMax .PostToAts, "PostToAts"
    End If

    If ServerVersion >= ApiServerVersions.AUTO_CANCEL_PARENT Then
        mWriter.AddBoolean .AutoCancelParent, "AutoCancelParent"
    End If

    If ServerVersion >= ApiServerVersions.ADVANCED_ORDER_REJECT Then
        mWriter.AddString .AdvancedErrorOverride, "AdvancedErrorOverride"
    End If

    If ServerVersion >= ApiServerVersions.MANUAL_ORDER_TIME Then
        mWriter.AddString .ManualOrderTime, "ManualOrderTime"
    End If

    If ServerVersion >= ApiServerVersions.PEGBEST_PEGMID_OFFSETS Then
        If UCase$(pContractSpec.Exchange) = "IBKRATS" Then
            mWriter.AddDoubleMax .MinimumTradeQuantity, "MinimumTradeQuantity"
        End If
        If .OrderType = TwsOrderTypes.TwsOrderTypePeggedToBest Then
            mWriter.AddLongMax .MinimumCompeteSize, "MinimumCompeteSize"
            mWriter.AddDoubleMax .CompeteAgainstBestOffset, "CompeteAgainstBestOffset"
            If .CompeteAgainstBestOffset = CompeteAgainstBestOffsetUpToMid Then
                mWriter.AddDoubleMax .MidOffsetAtWhole, "MidOffsetAtWhole"
                mWriter.AddDoubleMax .MidOffsetAtHalf, "MidOffsetAtHalf"
            End If
        ElseIf .OrderType = TwsOrderTypes.TwsOrderTypePeggedToMidpoint Then
            mWriter.AddDoubleMax .MidOffsetAtWhole, "MidOffsetAtWhole"
            mWriter.AddDoubleMax .MidOffsetAtHalf, "MidOffsetAtHalf"
        End If
    End If

End With

mWriter.Send True

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub
                    
Public Sub ReplaceFA(requestId As Integer, ByVal DataType As TwsFADataTypes, ByVal xml As String)
Const ProcName As String = "ReplaceFA"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REPLACE_FA
mWriter.AddLong VERSION, "Version"
mWriter.AddLong DataType, "Data type"
mWriter.AddString xml, "XML"
If ServerVersion >= ApiServerVersions.REPLACE_FA_END Then
    mWriter.AddLong GIdManager.GetTwsId(requestId, IdTypes.IdTypeAccount), "Request id"
End If

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestAccountSummary( _
                ByVal pReqId As Long, _
                ByVal pGroup As String, _
                ByVal pTags As String)
Const ProcName As String = "RequestAccountSummary"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

GIB.Log "Requesting account summary: requester id=" & pReqId, ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_ACCOUNT_SUMMARY
mWriter.AddLong VERSION, "Version"
mWriter.AddLong pReqId, "ReqId"
mWriter.AddString pGroup, "Group"
mWriter.AddString pTags, "Tags"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestAccountUpdates(ByVal subscribe As Boolean, ByVal acctCode As String)
Const ProcName As String = "RequestAccountUpdates"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 2

GIB.Log "Requesting account updates", ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_ACCT_DATA
mWriter.AddLong VERSION, "Version"
mWriter.AddBoolean subscribe, "Subscribe"

mWriter.AddString acctCode, "Account code"

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestAllOpenOrders()
Const ProcName As String = "RequestAllOpenOrders"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

GIB.Log "Requesting all open orders", ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_ALL_OPEN_ORDERS
mWriter.AddLong VERSION, "Version"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestAutoOpenOrders(ByVal autoBind As Boolean)
Const ProcName As String = "RequestAutoOpenOrders"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

GIB.Log "Requesting auto open orders", ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_AUTO_OPEN_ORDERS
mWriter.AddLong VERSION, "Version"
mWriter.AddBoolean autoBind, "Autobind"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestContractDetails( _
                ByVal pRequestId As Long, _
                ByVal pContractSpec As TwsContractSpecifier, _
                Optional ByVal pSecIdType As String, _
                Optional ByVal pSecId As String)
Const ProcName As String = "RequestContractDetails"
On Error GoTo Err

Const VERSION = 8

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
AssertArgument pContractSpec.Sectype <> TwsSecTypeCombo, "Combo contracts are not supported"

Dim lApiId As Long
lApiId = GIdManager.GetTwsId(pRequestId, IdTypeContractData)

GIB.Log "Requesting contract details: requester id=" & pRequestId & "; api id=" & lApiId & "; contract=" & pContractSpec.ToString, ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_CONTRACT_DATA
mWriter.AddLong VERSION, "Version"

mWriter.AddLong lApiId, "Request id"
mWriter.AddContract pContractSpec
mWriter.AddBoolean True, "Include expired"
mWriter.AddString pSecIdType, "SecIdType"
mWriter.AddString pSecId, "SecId"

If ServerVersion >= ApiServerVersions.BOND_ISSUERID Then mWriter.AddString "", "IssuerId"

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestCurrentTime()
Const ProcName As String = "RequestCurrentTime"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 3

mWriter.AddMessageId REQ_CURRENT_TIME
mWriter.AddLong VERSION, "Version"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestExecutions( _
                ByVal pRequestId As Long, _
                ByVal filter As TwsExecutionFilter)
Const ProcName As String = "RequestExecutions"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 3

Dim lApiId As Long
lApiId = GIdManager.GetTwsId(pRequestId, IdTypeExecution)

GIB.Log "Requesting executions: requester id=" & pRequestId & "; api id=" & lApiId, ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_EXECUTIONS
mWriter.AddLong VERSION, "Version"

mWriter.AddLong lApiId, "ReqId"

Dim theFilter As TwsExecutionFilter
If filter Is Nothing Then
    Set theFilter = New TwsExecutionFilter
Else
    Set theFilter = filter
End If

With theFilter
    mWriter.AddString .ClientID, "Client id"
    mWriter.AddString .AccountCode, "Account code"
    Dim lFromTime As Date: lFromTime = ConvertDateLocalToUTC(.Time)
    If .Timezone = "" Then
        mWriter.AddString IIf(.Time <> MaxDate, Format(lFromTime, "yyyymmdd-hh\:nn\:ss"), ""), "Fill Time"
    Else
        mWriter.AddString IIf(.Time <> MaxDate, Format(lFromTime, "yyyymmdd hh\:nn\:ss") & .Timezone, ""), "Fill Time"
    End If
    mWriter.AddString .Symbol, UCase$("Symbol")
    mWriter.AddString TwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString TwsOrderActionToString(.Action), "Action"
End With

mWriter.Send True

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestFA(ByVal DataType As TwsFADataTypes)
Const ProcName As String = "RequestFA"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_FA
mWriter.AddLong VERSION, "Version"
mWriter.AddLong DataType, "Data type"

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestFundamentalData( _
                ByVal pReqId As Long, _
                ByVal pContractSpec As TwsContractSpecifier, _
                ByVal pReportType As String)
Const ProcName As String = "RequestFundamentalData"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 3

GIB.Log "Requesting fundamental data: requester id=" & pReqId & "; contract=" & pContractSpec.ToString, ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_FUNDAMENTAL_DATA
mWriter.AddLong VERSION, "Version"
mWriter.AddLong pReqId, "ReqId"

mWriter.AddString pContractSpec.ConId, "ConId"
mWriter.AddString pContractSpec.Symbol, "Symbol"
mWriter.AddString TwsSecTypeToShortString(pContractSpec.Sectype), "SecType"
mWriter.AddString pContractSpec.Exchange, "Exchange"
mWriter.AddString pContractSpec.PrimaryExch, "PrimaryExch"
mWriter.AddString pContractSpec.CurrencyCode, "Currency"
mWriter.AddString UCase$(pContractSpec.LocalSymbol), "LocalSymbol"
mWriter.AddString UCase$(pReportType), "ReportType"
mWriter.AddString "0", "Options count"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestGlobalCancel()
Const ProcName As String = "RequestGlobalCancel"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

GIB.Log "Requesting global cancel", ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_GLOBAL_CANCEL
mWriter.AddLong VERSION, "Version"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestHistoricalData( _
                ByVal pRequestId As Long, _
                ByRef pRequest As TwsHistoricalDataRequest, _
                Optional ByVal pKeepUpToDate As Boolean)
Const ProcName As String = "RequestHistoricalData"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
AssertArgument ServerVersion >= ApiServerVersions.HISTORICAL_SCHEDULE Or UCase$(pRequest.WhatToShow) <> "SCHEDULE", "Historical schedule requests not supported"

Dim lApiId As Long
lApiId = GIdManager.GetTwsId(pRequestId, IdTypeHistoricalData)

GIB.Log "Requesting historical data for: " & UCase$(pRequest.ContractSpec.LocalSymbol) & _
            "; requester id=" & pRequestId & _
            "; api id=" & lApiId & _
            "; barsize=" & pRequest.BarSizeSetting & _
            "; endTime=" & pRequest.EndDateTime & _
            "; duration=" & pRequest.Duration & _
            "; whattoshow=" & pRequest.WhatToShow & _
            "; use rth=" & IIf(pRequest.MainSessionOnly, 1, 0), _
            ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_HISTORICAL_DATA
mWriter.AddLong lApiId, "Request id"
mWriter.AddContract pRequest.ContractSpec

Dim lExpired  As Boolean
lExpired = ContractHasExpired(pRequest.ContractSpec)
mWriter.AddBoolean lExpired, "Include expired"   ' can't include expired for non-expiring contracts

mWriter.AddString pRequest.EndDateTime, "End date"
mWriter.AddString pRequest.BarSizeSetting, "Bar Size"

mWriter.AddString pRequest.Duration, "Duration"
mWriter.AddBoolean pRequest.MainSessionOnly, "Use RTH"
mWriter.AddString pRequest.WhatToShow, "What to show"

mWriter.AddString TwsHistDataDateFormats.DateFormatString, "Date format"

If pRequest.ContractSpec.Sectype = TwsSecTypeCombo Then
    mWriter.AddLong pRequest.ContractSpec.ComboLegs.Count, "Combo legs count"
    Dim lComboLeg As TwsComboLeg
    Dim i As Long
    For Each lComboLeg In pRequest.ContractSpec.ComboLegs
        With lComboLeg
            i = i + 1
            mWriter.AddString .ConId, "ConId" & i
            mWriter.AddString .Ratio, "Ratio" & i
            mWriter.AddString TwsOrderActionToString(.Action), "Action" & i
            mWriter.AddString .Exchange, "Exchange" & i
        End With
    Next
End If

If ServerVersion >= ApiServerVersions.SYNT_REALTIME_BARS Then mWriter.AddBoolean pKeepUpToDate, "KeepUpToDate"
mWriter.AddString "", "Options"

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestManagedAccounts()
Const ProcName As String = "RequestManagedAccounts"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_MANAGED_ACCTS
mWriter.AddLong VERSION, "Version"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestMarketData( _
                ByVal pTickerId As Long, _
                ByVal pContractSpec As TwsContractSpecifier, _
                ByVal pGenericTicks As String, _
                ByVal pSnapshot As Boolean, _
                ByVal pRegulatorySnapshot As Boolean)
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 11

Dim lApiId As Long
lApiId = GIdManager.GetTwsId(pTickerId, IdTypeMarketData)

GIB.Log "Requesting market data" & _
        IIf(pSnapshot, " (snapshot)", "") & _
        ": requester id=" & pTickerId & _
        "; api id=" & lApiId & _
        "; contract=" & pContractSpec.ToString, _
        ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_MKT_DATA
mWriter.AddLong VERSION, "Version"
mWriter.AddLong lApiId, "Ticker id"
mWriter.AddContract pContractSpec
With pContractSpec
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
                mWriter.AddString TwsOrderActionToString(.Action), "Action" & i
                mWriter.AddString .Exchange, "Exchange" & i
            End With
        Next
    End If
    
    If .DeltaNeutralContract.ConId <> 0 Then
        mWriter.AddBoolean True, "Delta neutral"
        mWriter.AddLong .DeltaNeutralContract.ConId, "Delta neutral conid"
        mWriter.AddDouble .DeltaNeutralContract.Delta, "Delta neutral delta"
        mWriter.AddDouble .DeltaNeutralContract.Price, "Delta neutral price"
    Else
        mWriter.AddBoolean False, "Delta neutral"
    End If
    
    mWriter.AddString pGenericTicks, "Generic tick list"
    
    mWriter.AddBoolean pSnapshot, "Snapshot"
    
    If ServerVersion >= ApiServerVersions.SMART_COMPONENTS Then mWriter.AddBoolean pRegulatorySnapshot, "Regulatory Snapshot"

    mWriter.AddString "", "Options"
    
    mWriter.Send
End With

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestMarketDataType(ByVal pMarketDataType As Long)
Const ProcName As String = "RequestMarketDataType"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_MARKET_DATA_TYPE
mWriter.AddLong VERSION, "Version"
mWriter.AddLong pMarketDataType, "MarketDataType"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestMarketDepth( _
                ByVal pTickerId As Long, _
                ByVal pContractSpec As TwsContractSpecifier, _
                Optional ByVal pIsSmartDepth As Boolean, _
                Optional ByVal pNumberOfRows As Long = 20)
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 5

AssertArgument ServerVersion >= ApiServerVersions.SMART_DEPTH Or (Not pIsSmartDepth), "SMART depth request is not supported"
AssertArgument ServerVersion >= ApiServerVersions.MKT_DEPTH_PRIM_EXCHANGE Or pContractSpec.PrimaryExch = "", "PrimaryExch parameter is not supported"

Dim lApiId As Long
lApiId = GIdManager.GetTwsId(pTickerId, IdTypeMarketDepth)

GIB.Log "Requesting market depth: requester id=" & pTickerId & "; api id=" & lApiId & "; contract=" & pContractSpec.ToString, ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_MKT_DEPTH
mWriter.AddLong VERSION, "Version"
mWriter.AddLong lApiId, "Request id"

mWriter.AddContract pContractSpec, pIgnorePrimaryExchange:=(ServerVersion < ApiServerVersions.MKT_DEPTH_PRIM_EXCHANGE)
mWriter.AddLong pNumberOfRows, "Num rows"
If ServerVersion >= ApiServerVersions.SMART_DEPTH Then mWriter.AddBoolean pIsSmartDepth, "IsSmartDepth"
mWriter.AddString "", "Options"

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestMarketRule(ByVal pMarketRuleId As Long)
Const ProcName As String = "RequestMarketRule"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
Assert ServerVersion >= ApiServerVersions.MARKET_RULES, "Market rule requests not supported"

GIB.Log "Requesting market rule: id=" & pMarketRuleId, ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId TwsSocketOutMsgTypes.RequestMarketRule
mWriter.AddLong pMarketRuleId, "Market Rule ID"

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestNewsBulletins(allMsgs As Boolean)
Const ProcName As String = "RequestNewsBulletins"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

mWriter.AddMessageId REQ_NEWS_BULLETINS
mWriter.AddLong VERSION, "Version"
mWriter.AddBoolean allMsgs, "All messages"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestOpenOrders()
Const ProcName As String = "RequestOpenOrders"
On Error GoTo Err

Const VERSION = 1

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

GIB.Log "Requesting open orders", ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_OPEN_ORDERS
mWriter.AddLong VERSION, "Version"
mWriter.Send True

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestOptionParameters( _
                ByVal pRequestId As Long, _
                ByVal pUnderlyingSymbol As String, _
                ByVal pExchange As String, _
                ByVal pUnderlyingSecType As TwsSecTypes, _
                ByVal pUnderlyingConId As Long)
Const ProcName As String = "RequestOptionParameters"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Assert ServerVersion >= ApiServerVersions.SEC_DEF_OPT_PARAMS_REQ, "Option parameters requests not supported"

mWriter.AddMessageId TwsSocketOutMsgTypes.RequestOptionParameters

Dim lApiId As Long
lApiId = GIdManager.GetTwsId(pRequestId, IdTypeContractData)

GIB.Log "Requesting option parameters: " & _
        "requester id=" & pRequestId & _
        "; api id=" & lApiId & _
        "; symbol=" & pUnderlyingSymbol & _
        "; exchange=" & pExchange & _
        "; underlying sectype=" & TwsSecTypeToShortString(pUnderlyingSecType) & _
        "; underlying conid=" & pUnderlyingConId, _
        ModuleName, ProcName, , LogLevelDetail

mWriter.AddLong lApiId, "Request Id"
mWriter.AddString pUnderlyingSymbol, "Underlying Symbol"
mWriter.AddString pExchange, "Exchange"
mWriter.AddString TwsSecTypeToShortString(pUnderlyingSecType), "Underlying SecType"
mWriter.AddLong pUnderlyingConId, "Underlying ConId"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestPositions()
Const ProcName As String = "RequestPositions"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

GIB.Log "Requesting positions", ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_POSITIONS
mWriter.AddLong VERSION, "Version"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestRealtimeBars( _
                ByVal pTickerId As Long, _
                ByVal pContractSpec As TwsContractSpecifier, _
                ByVal pBarSize As Long, _
                ByVal pWhatToShow As String, _
                ByVal pUseRTH As Boolean)
Const ProcName As String = "RequestRealTimeBars"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 3

Dim lApiId As Long
lApiId = GIdManager.GetTwsId(pTickerId, IdTypeHistoricalData)

GIB.Log "Requesting realtime bars: requester id=" & pTickerId & "; api id=" & lApiId & "; contract=" & pContractSpec.ToString, ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_REAL_TIME_BARS
mWriter.AddLong VERSION, "Version"
mWriter.AddLong lApiId, "TickerId"

mWriter.AddContract pContractSpec
mWriter.AddLong pBarSize, "BarSize"   ' this parameter is not currently used
mWriter.AddString pWhatToShow, "WhatToShow"
mWriter.AddBoolean pUseRTH, "UseRTH"
mWriter.AddString "", "Options"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestScannerParameters()
Const ProcName As String = "RequestScannerParameters"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 1

GIB.Log "Requesting scanner parameters", ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_SCANNER_PARAMETERS
mWriter.AddLong VERSION, "Version"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestScannerSubscription( _
                ByVal pRequestId As Long, _
                ByVal pSubscription As TwsScannerSubscription, _
                ByVal pSubscriptionOptions As Parameters, _
                ByVal pFilterOptions As Parameters)
Const ProcName As String = "RequestScannerSubscription"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const VERSION = 4

Dim lApiId As Long
lApiId = GIdManager.GetTwsId(pRequestId, IdTypeScanner)

GIB.Log "Requesting scanner subscription: requester id=" & pRequestId & _
        "; api id=" & lApiId, _
        ModuleName, ProcName, , LogLevelDetail

mWriter.AddMessageId REQ_SCANNER_SUBSCRIPTION
If mServerVersion < SCANNER_GENERIC_OPTS Then mWriter.AddLong VERSION, "Version"
mWriter.AddLong lApiId, "Request Id"
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
mWriter.AddString pSubscription.MaturityDateAbove, "MaturityDateAbove"
mWriter.AddString pSubscription.MaturityDateBelow, "MaturityDateBelow"
mWriter.AddDoubleMax pSubscription.CouponRateAbove, "CouponRateAbove"
mWriter.AddDoubleMax pSubscription.CouponRateBelow, "CouponRateBelow"
mWriter.AddBoolean pSubscription.ExcludeConvertible, "ExcludeConvertible"
mWriter.AddLongMax pSubscription.AverageOptionVolumeAbove, "AverageOptionVolumeAbove"
mWriter.AddString pSubscription.ScannerSettingPairs, "SettingPairs"
mWriter.AddString pSubscription.StockTypeFilter, "StockTypeFilter"

If ServerVersion >= ApiServerVersions.SCANNER_GENERIC_OPTS Then
    mWriter.AddTagValues pFilterOptions, "FilterOptions"
End If

mWriter.AddTagValues pSubscriptionOptions, "SubscriptionOptions"

mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub SetTwsLoggingLevel(ByVal pLogLevel As TwsLoggingLevels)
mLogLevel = pLogLevel
End Sub

'================================================================================
' Helper Functions
'================================================================================

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

Private Sub sendLogLevel()
Const ProcName As String = "sendLogLevel"
On Error GoTo Err

Const VERSION = 1

mWriter.AddMessageId SET_SERVER_LOGLEVEL
mWriter.AddLong VERSION, "Version"
mWriter.AddLong mLogLevel, "Loglevel"
mWriter.Send

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub setConnectionState(ByVal pState As TwsConnectionStates, ByVal pMessage As String)
Const ProcName As String = "setConnectionState"
On Error GoTo Err

mConnectionState = pState

If Not mConnectionStatusConsumer Is Nothing Then mConnectionStatusConsumer.NotifyAPIConnectionStateChange mConnectionState, pMessage

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub startPerformanceTLIPeriod()
Const ProcName As String = "startPerformanceTLIPeriod"
On Error GoTo Err

Set mPerformanceTLIPeriod = GetGlobalTimerList.Add(Nothing, _
                                        CDate(Int((GetTimestamp + GIB.OneMinute + GIB.OneSecond) / GIB.OneMinute) * GIB.OneMinute), _
                                        ExpiryTimeUnitDateTime)
Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub startPerformanceTLISecond()
Const ProcName As String = "startPerformanceTLISecond"
On Error GoTo Err

Set mPerformanceTLISecond = GetGlobalTimerList.Add(Nothing, _
                                        CDate(Int((GetTimestamp + 2 * GIB.OneSecond) / GIB.OneSecond) * GIB.OneSecond), _
                                        ExpiryTimeUnitDateTime)
    
Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub startPerformanceLogging()
Const ProcName As String = "startPerformanceLogging"
On Error GoTo Err

If Not mLogApiMessageStats Then Exit Sub
If Not GIB.Logger.IsLoggable(LogLevelNormal) Then Exit Sub
startPerformanceTLISecond
startPerformanceTLIPeriod

Exit Sub

Err:
GIB.HandleUnexpectedError Nothing, ProcName, ModuleName
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

