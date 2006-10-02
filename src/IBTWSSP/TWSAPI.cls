VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWSAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

' Raised when a successful connection to TWS has been achieved.
Event Connected()

' Raised when connection to TWS has failed and the connection
' retry interval is zero.
Event ConnectFailed(ByVal Description As String, _
                ByVal retrying As Boolean)

' Raised when an attempt to connect to TWS is initiated.
Event Connecting()

' Raised when the connection to TWS is disconnected.
Event ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)

Event ConnectionToIBClosed()

Event ConnectionToIBRecovered()

Event Disconnected(ByVal reason As String)
    ' Raised when we disconnect from TWS

Event FAData( _
                ByVal DataType As FADataTypes, _
                ByVal XMLData As String)
                
Event managedAccounts( _
                ByVal timestamp As Date, _
                ByVal accountsList As String)

Event updateAccountTime( _
                ByVal accountTime As String)

Event updateAccountValue( _
                ByVal key As String, _
                ByVal value As String, _
                ByVal currencyCode As String, _
                ByVal accountName As String)

Event updateNewsBulletin( _
                ByVal msgID As Integer, _
                ByVal msgType As BulletinTypes, _
                ByVal message As String, _
                ByVal origExchange As String)
                
Event updatePortfolioValue( _
                ByVal pContractSpecifier As TradeBuildSP.IContractSpecifier, _
                ByVal position As Long, _
                ByVal marketPrice As Double, _
                ByVal marketValue As Double, _
                ByVal averageCost As Double, _
                ByVal unrealizedPNL As Double, _
                ByVal realizedPNL As Double, _
                ByVal accountName As String)


'================================================================================
' Constants
'================================================================================

Private Const CLIENT_VERSION As Long = 23
Private Const INCARNATION_INCREMENT As Long = &H1000&

Private Const InitialMaxTickers As Long = 100&

Private Const MARKET_DEPTH_TICKER_BASE As Long = &H400000

Private Const HistoricalDataTickerBase As Long = &H410000
Private Const HistoricalDataTickerIncrement As Long = &H100

Private Const OrderIdBase As Long = &H800000

Private Const MaxConcurrentHistDataRequests As Long = 50    ' Limit imposed by IB

Private Const SpName As String = "TradeWright IB TWS Service Provider"

Private Const TWSWhatToShowAsk As String = "ASK"
Private Const TWSWhatToShowBid As String = "BID"
Private Const TWSWhatToShowBidAsk As String = "BID_ASK"
Private Const TWSWhatToShowMidpoint As String = "MIDPOINT"
Private Const TWSWhatToShowTrades As String = "TRADES"

'================================================================================
' Enums
'================================================================================

Private Enum InternalErrorCodes
    DataIncomplete = vbObjectError + 512
End Enum

Private Enum TWSHistDataDateFormats
    DateFormatString = 1
    DateFormatInteger = 2
End Enum
'================================================================================
' Types
'================================================================================

Private Type HistoricalDataRequestTableEntry
    inUse                   As Boolean
    dataRetrievedFromTWS    As Boolean
    cancelling              As Boolean
    barSpecifier            As TradeBuildSP.BarDataSpecifier
    earliestReuseTime       As Date
    numberOfTwsRequests     As Long         ' the number of historical data requests
                                            ' made to TWS
    barType                 As TradeBuildSP.BarTypes
    barLength               As Long
    historicalDataReader    As HistDataReader
    nextBarIndex            As Long
    bars                    As LinkedList
    requestKeys()           As String
End Type

Private Type TickerTableEntry
    incarnation             As Long
    contractSpecifier       As TradeBuildSP.IContractSpecifier
    prevBid                 As Double
    prevBidSize             As Long
    prevAsk                 As Double
    prevAskSize             As Long
    prevTrade               As Double
    prevTradeSize           As Long
    inUse                   As Boolean
    receivingMarketDepth    As Boolean
    dataConsumer            As IStreamingDataConsumer
    dataReader              As RealtimeDataReader
End Type

'================================================================================
' Member variables
'================================================================================

Private WithEvents mSocket As Winsock
Attribute mSocket.VB_VarHelpID = -1
Private WithEvents mConnectionTimer As IntervalTimer
Attribute mConnectionTimer.VB_VarHelpID = -1

Private mClientID As Long
Private mServer As String
Private mPort As Long
Private mProviderKey As String
Private mConnectionRetryIntervalSecs As Long
Private mDisconnectedByApplication As Boolean

Private mSocketInputData() As Byte
Private mBufferOut() As Byte
Private mIndexOut As Long
Private mBufferIn() As Byte
Private mIndexIn As Long
Private mInputIndex As Long
Private mInputIndexChk As Long

Private mServerVersion As Long
Private mCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer

Private mConnectionState As ConnectionStates

Private mTickerTable() As TickerTableEntry

Private mContractDetailsRequest As TWSContractDetailsRequestQueueEntry
Private mFirstContractDetailsReceived As Boolean

Private mExchangeCodes() As String
Private mMaxExchangeCodesIndex As Long

Private mNextOrderID  As Long

Private mLogLevel As LogLevels
Private mTWSLogLevel As TWSLogLevels

Private mTwsHistDataRequestQueue As Queue
Private mConcurrentHistDataRequests As Long

Private mHistoricalDataTable() As HistoricalDataRequestTableEntry

Private mOrderIDMap As Collection

Private mTwsContractDetailsRequestQueue As Queue

Private mExecutions As Collection

'================================================================================
' External function declarations
'================================================================================

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                            Destination As Any, _
                            source As Any, _
                            ByVal length As Long)
                            
Private Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                            Destination As Any, _
                            source As Any, _
                            ByVal length As Long)
                            
'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

ReDim mTickerTable(InitialMaxTickers) As TickerTableEntry
ReDim mBufferIn(1023) As Byte
ReDim mBufferOut(1023) As Byte
mConnectionState = ConnectionStates.ConnNotConnected

Set mOrderIDMap = New Collection
Set mExecutions = New Collection

mNextOrderID = 1

mLogLevel = LogLevelLow
mTWSLogLevel = TWSLogLevelError ' this is the TWS default

Set mTwsHistDataRequestQueue = New Queue
mTwsHistDataRequestQueue.initialCapacity = 20

ReDim mHistoricalDataTable(MaxConcurrentHistDataRequests) As HistoricalDataRequestTableEntry

Set mTwsContractDetailsRequestQueue = New Queue
mTwsContractDetailsRequestQueue.initialCapacity = 20

ReDim mExchangeCodes(100) As String
mMaxExchangeCodesIndex = -1

addExchangeCode "ACE"
addExchangeCode "AEB"
addExchangeCode "AMEX"
addExchangeCode "ARCA"

addExchangeCode "BELFOX"
addExchangeCode "BOX"
addExchangeCode "BRUT"
addExchangeCode "BTRADE"
addExchangeCode "BVME"

addExchangeCode "CAES"
addExchangeCode "CBOE"
addExchangeCode "CDE"
addExchangeCode "CFE"

addExchangeCode "DTB"

addExchangeCode "EBS"
addExchangeCode "ECBOT"
addExchangeCode "EUREX"
addExchangeCode "EUREXUS"

addExchangeCode "FTA"
addExchangeCode "FWB"

addExchangeCode "GLOBEX"

addExchangeCode "HKFE"

addExchangeCode "IBIS"
addExchangeCode "IDEAL"
addExchangeCode "IDEALPRO"
addExchangeCode "IDEM"
addExchangeCode "INET"
addExchangeCode "INSTINET"
addExchangeCode "ISE"
addExchangeCode "ISLAND"

addExchangeCode "LIFFE"
addExchangeCode "LIFFE_NF"
addExchangeCode "LSE"

addExchangeCode "MATIF"
addExchangeCode "MEFF"
addExchangeCode "MEFFRV"
addExchangeCode "MONEP"
addExchangeCode "MXT"

addExchangeCode "NASDAQ"
addExchangeCode "NQLX"
addExchangeCode "NYMEX"
addExchangeCode "NYSE"

addExchangeCode "OMS"
addExchangeCode "ONE"
addExchangeCode "OSE.JPN"

addExchangeCode "PHLX"
addExchangeCode "PINK"
addExchangeCode "PSE"

addExchangeCode "RDBK"

addExchangeCode "SBF"
addExchangeCode "SFB"
addExchangeCode "SGX"
addExchangeCode "SMART"
addExchangeCode "SNFE"
addExchangeCode "SOFFEX"
addExchangeCode "SUPERSOES"
addExchangeCode "SWB"
addExchangeCode "SWX"

addExchangeCode "TSE"
addExchangeCode "TSE.JPN"

addExchangeCode "VENTURE"
addExchangeCode "VIRTX"
addExchangeCode "VWAP"

ReDim Preserve mExchangeCodes(mMaxExchangeCodesIndex) As String
End Sub


Private Sub Class_Terminate()
Debug.Print "TWSAPI terminated"
End Sub

'================================================================================
' mConnectionTimer Event Handlers
'================================================================================

Private Sub mConnectionTimer_TimerExpired()
On Error GoTo err
Connect
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mConnectionTimer_TimerExpired", _
                err.Description
End Sub

'================================================================================
' Socket Event Handlers
'================================================================================

Private Sub mSocket_Close()
On Error GoTo err
handleTwsDisconnection
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mSocket_Close", _
                err.Description
End Sub

Private Sub mSocket_Connect()
On Error GoTo err
mConnectionState = ConnectionStates.ConnConnecting
RaiseEvent Connecting
logMessage "Connecting to TWS: negotiating versions and client id"
addData CLIENT_VERSION
send
addData mClientID
send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mSocket_Connect", _
                err.Description
End Sub

Private Sub mSocket_DataArrival(ByVal bytesTotal As Long)

On Error GoTo err

If bytesTotal = 0 Then
    checkmessages
    Exit Sub
End If

ReDim mSocketInputData(bytesTotal - 1) As Byte
mSocket.GetData mSocketInputData, vbArray + vbByte
If mLogLevel >= LogLevelHigh Then
    mCommonServiceConsumer.NotifyListeners showBuffer(mSocketInputData), _
                                        ListenValueTypes.ValueTypeRawSocketData, _
                                        Me
End If

'Debug.Print "Socket: " & showBuffer(mSocketInputData)

'if spaceUsed + newSpaceRequired > lengthOfBuffer then
If mIndexIn + bytesTotal > UBound(mBufferIn) + 1 Then
    'if newSpaceRequired <= lengthOfBuffer - unprocessedLength then
    If bytesTotal <= (UBound(mBufferIn) + 1 - mIndexIn + mInputIndex) Then
        'if unprocessedLength <> 0 then
        If mInputIndex < mIndexIn Then
            ' copy the existing contents down to the start of the buffer
            MoveMemory mBufferIn(0), mBufferIn(mInputIndex), mIndexIn - mInputIndex
            mIndexIn = mIndexIn - mInputIndex
            mInputIndex = 0
        Else
            mIndexIn = 0
            mInputIndex = 0
        End If
    Else
        ReDim Preserve mBufferIn(UBound(mBufferIn) + IIf(bytesTotal > 1024, bytesTotal, 1024)) As Byte
    End If
End If
CopyMemory mBufferIn(mIndexIn), mSocketInputData(0), bytesTotal
mIndexIn = mIndexIn + bytesTotal

checkmessages

Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mSocket_DataArrival", _
                err.Description

End Sub

Private Sub mSocket_Error(ByVal number As Integer, _
                        Description As String, _
                        ByVal Scode As Long, _
                        ByVal source As String, _
                        ByVal HelpFile As String, _
                        ByVal HelpContext As Long, _
                        CancelDisplay As Boolean)
Dim errorNum As MSWinsockLib.ErrorConstants

On Error GoTo err
errorNum = number
Select Case errorNum
'Case sckInvalidPropertyValue

'Case sckGetNotSupported

'Case sckSetNotSupported

'Case sckOutOfMemory

'Case sckBadState

'Case sckInvalidArg

'Case sckSuccess

'Case sckUnsupported

'Case sckInvalidOp

'Case sckOutOfRange

'Case sckWrongProtocol

'Case sckOpCanceled

'Case sckInvalidArgument

'Case sckWouldBlock

'Case sckInProgress

'Case sckAlreadyComplete

'Case sckNotSocket

'Case sckMsgTooBig

'Case sckPortNotSupported

'Case sckAddressInUse

Case sckAddressNotAvailable, _
    sckNetworkSubsystemFailed, _
    sckNetworkUnreachable, _
    sckNetReset, _
    sckConnectAborted, _
    sckConnectionRefused, _
    sckHostNotFound, _
    sckHostNotFoundTryAgain, _
    sckTimedout
    
    If mConnectionState = ConnectionStates.ConnNotConnected Then
        ReleaseSocket
        If mConnectionRetryIntervalSecs = 0 Then
            logMessage "Failed to connect to TWS: " & _
                        Description & ": " & _
                        calcConnectionString, _
                        LogLevelLow
            
            cancelContractRequests "Failed to connect to TWS"
            RaiseEvent ConnectFailed(Description, False)
        Else
            logMessage "Failed to connect to TWS - retrying: " & _
                        Description & ": " & _
                        calcConnectionString, _
                        LogLevelLow
            
            RaiseEvent ConnectFailed(Description, True)
            retryConnection
        End If
    Else
        logMessage "Socket error " & number & ": " & _
                    Description & ": " & _
                    calcConnectionString, _
                    LogLevelLow
        handleTwsDisconnection
    End If
'Case sckConnectionReset

'Case sckNoBufferSpace

'Case sckAlreadyConnected

'Case sckNotConnected

'Case sckSocketShutdown

'Case sckNotInitialized

'Case sckNonRecoverableError

'Case sckNoData
Case Else
    cancelContractRequests "Failed to connect to TWS"
    logMessage "Socket error " & number & ": " & _
                Description & ": " & _
                calcConnectionString, _
                LogLevelLow
    RaiseEvent ConnectFailed(Description, False)
End Select

Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mSocket_Error", _
                err.Description

End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Friend Property Let commonServiceConsumer(ByVal value As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

Friend Property Let connectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
If mConnectionRetryIntervalSecs = 0 Then Set mConnectionTimer = Nothing
End Property

Friend Property Get connectionState() As ConnectionStates
connectionState = mConnectionState
End Property

Friend Property Get InBuffer() As String
InBuffer = showBuffer(mBufferIn)
End Property

Friend Property Get initialOrderId() As Long
initialOrderId = mNextOrderID
End Property

Friend Property Let logLevel(ByVal value As LogLevels)
mLogLevel = value
End Property

Friend Property Let port(ByVal value As Long)
mPort = value
End Property

Friend Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Friend Property Get serverVersion() As Long
serverVersion = mServerVersion
End Property

Friend Property Let TWSLogLevel(ByVal value As TWSLogLevels)
mTWSLogLevel = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub CancelHistoricalData(ByVal requestHandle As Long)
Dim i As Long
Dim key As String

On Error GoTo err
If Not mHistoricalDataTable(requestHandle).dataRetrievedFromTWS Then
    ' all data has already been got from TWS
    
    '   more here????
    
    releaseHistoricalDataTableEntry requestHandle

Else
    If mServerVersion < 24 Then
        ' cancel not supported by TWS - we'll complete the cancel when
        ' the data arrives
        mHistoricalDataTable(i).cancelling = True
    Else
        For i = 0 To UBound(mHistoricalDataTable(requestHandle).requestKeys)
            key = mHistoricalDataTable(requestHandle).requestKeys(i)
            If key = "" Then
                ' data has alreay been retrieved for this TWS request, so
                ' nothing to do
                
            Else
                On Error Resume Next
                mTwsHistDataRequestQueue.remove key
                On Error GoTo 0
                If err.number = 0 Then
                    ' request was still in queue - no further action required
                ElseIf err.number <> ErrorCodes.ErrIllegalArgumentException Then
                    err.Raise err.number
                Else
                    ' request is with TWS so cancel it
                    
                    If mConnectionState = ConnConnected Then
                        Const version As Long = 1
                
                        addData CANCEL_HISTORICAL_DATA
                        addData version
                        addData requestHandle + HistoricalDataTickerBase + i * HistoricalDataTickerIncrement
                        send
                    End If
                    
                    mHistoricalDataTable(requestHandle).cancelling = True
                    mConcurrentHistDataRequests = mConcurrentHistDataRequests - 1
                    
                    processNextHistDataRequest
                End If
            End If
        Next

        If mHistoricalDataTable(requestHandle).cancelling Then
            ' it's possible that TWS has just sent the data for one or more of the cancelled
            ' requests but we haven't received it yet, so we set this entry to indicate
            ' the cancellation and prevent it being reused for 'long enough' to ensure
            ' that any pending data has been received - ie 60 seconds
            releaseHistoricalDataTableEntry requestHandle
            mHistoricalDataTable(requestHandle).cancelling = True
            mHistoricalDataTable(requestHandle).earliestReuseTime = Now + 60 / 86400
        Else
            releaseHistoricalDataTableEntry requestHandle
        End If
    End If
End If


Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "CancelHistoricalData", _
                err.Description
        
End Sub

Friend Sub CancelMarketData(ByVal tickerid As Long)
On Error GoTo err
cancelMarketDataEx tickerid, True
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "CancelMarketData", _
                err.Description
End Sub

Friend Sub CancelMarketDepth(ByVal tickerid As Long)
On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

Const version = 1

With mTickerTable(tickerid)
    If Not .inUse Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "IBTWSSP.TWSAPI::cancelMarketDepth", _
                                    "Invalid ticker id " & tickerid
    If Not .receivingMarketDepth Then
        err.Raise ErrorCodes.ErrIllegalStateException, _
                    "IBTWSSP.TWSAPI::cancelMarketDepth", _
                    "Ticker id " & tickerid & " not currently receiving market depth"
    End If
    mTickerTable(tickerid).receivingMarketDepth = False
End With

addData CANCEL_MKT_DEPTH
addData version
addData MARKET_DEPTH_TICKER_BASE + tickerid
send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "CancelMarketDepth", _
                err.Description

End Sub

Friend Sub cancelNewsBulletins()

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

Const version = 1

addData CANCEL_NEWS_BULLETINS
addData version
send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "cancelNewsBulletins", _
                err.Description
End Sub

Friend Sub cancelOrder(ByVal orderId As Long)

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::cancelOrder", _
                "Not connected"
End If

Const version = 1
addData CANCEL_ORDER
addData version
addData orderId
send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "cancelOrder", _
                err.Description

End Sub

Friend Sub Connect()

On Error GoTo err
mDisconnectedByApplication = False

If mConnectionState <> ConnectionStates.ConnNotConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::connect", _
                "Already connected"
End If

createsocket


logMessage "Connecting to TWS: " & calcConnectionString
mSocket.Connect mServer, mPort

Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "Connect", _
                err.Description
End Sub

Friend Sub disconnect( _
                ByVal reason As String)
On Error GoTo err
mDisconnectedByApplication = True
If Not mConnectionTimer Is Nothing Then mConnectionTimer.StopTimer
mConnectionState = ConnectionStates.ConnNotConnected
releaseAllTickerTableEntries
ReleaseSocket
mIndexIn = 0
mIndexOut = 0
mInputIndex = 0
RaiseEvent Disconnected(reason)
logMessage "Disconnected from: " & calcConnectionString, LogLevelLow
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "disconnect", _
                err.Description
End Sub

Friend Function IsValidExchangeCode(ByVal code As String) As Boolean
Dim bottom As Long
Dim top As Long
Dim middle As Long

code = UCase$(code)
bottom = 0
top = mMaxExchangeCodesIndex
middle = Fix((bottom + top) / 2)

Do
    If code < mExchangeCodes(middle) Then
        top = middle
    ElseIf code > mExchangeCodes(middle) Then
        bottom = middle
    Else
        IsValidExchangeCode = True
        Exit Function
    End If
    middle = Fix((bottom + top) / 2)
Loop Until bottom = middle

If code = mExchangeCodes(middle) Then IsValidExchangeCode = True
End Function

Friend Sub placeOrder( _
                ByVal orderSubmitter As orderSubmitter, _
                ByVal pOrder As TradeBuildSP.IOrder, _
                ByVal transmit As Boolean, _
                ByVal parentId As String, _
                ByVal ocaGroup As String, _
                ByVal routingHint As String)

Dim mapEntry As OrderIdMapEntry

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::placeOrder", _
                "Already connected"
End If

If pOrder.brokerId = "" Then
    pOrder.brokerId = CStr(mNextOrderID)
    mNextOrderID = mNextOrderID + 1
    Set mapEntry.order = pOrder
    Set mapEntry.orderSubmitter = orderSubmitter
    mOrderIDMap.add mapEntry, generateTwsOrderKey(pOrder.brokerId, mClientID)
End If

sendOrder pOrder, transmit, parentId, ocaGroup, routingHint
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "placeOrder", _
                err.Description

End Sub
                    
Public Function ReadHistoricalBar( _
                        ByVal requestHandle As Long, _
                        ByRef timestamp As Date, _
                        ByRef barType As Long, _
                        ByRef barLength As Long, _
                        ByRef openPrice As Double, _
                        ByRef highPrice As Double, _
                        ByRef lowPrice As Double, _
                        ByRef closePrice As Double, _
                        ByRef volume As Long, _
                        ByRef tickVolume As Long, _
                        ByRef openInterest As Long) As Boolean

Dim Bar As HistoricBar

On Error GoTo err
With mHistoricalDataTable(requestHandle)
    If .nextBarIndex = -1 Then
        releaseHistoricalDataTableEntry requestHandle
    Else
        ReadHistoricalBar = True
        barType = .barType
        barLength = .barLength
        Bar = .bars.itemAt(.nextBarIndex)
        With Bar
            timestamp = .timestamp
            openPrice = .openValue
            highPrice = .highValue
            lowPrice = .lowValue
            closePrice = .closeValue
            volume = .volume
            tickVolume = 0
            openInterest = 0
        End With
        .nextBarIndex = .bars.nextIndex(.nextBarIndex)
    End If
End With
Exit Function
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "ReadHistoricalBar", _
                err.Description

End Function

Friend Sub ReplaceFA(ByVal DataType As FADataTypes, ByVal xml As String)


On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::ReplaceFA", _
                "Already connected"
End If

Const version = 1

If mServerVersion < 13 Then
    err.Raise ErrorCodes.ErrUnsupportedOperationException, _
                "IBTWSSP.TWSAPI::ReplaceFA", _
                "Server version does not support ReplaceFA"
End If

addData REPLACE_FA
addData version
addData DataType
addData xml

send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "ReplaceFA", _
                err.Description
End Sub

Friend Sub requestAccountUpdates(ByVal subscribe As Boolean, ByVal acctCode As String)

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::requestAccountUpdates", _
                "Already connected"
End If

Const version = 2

addData REQ_ACCT_DATA
addData version
addBoolean subscribe

If mServerVersion >= 9 Then addData acctCode

send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "requestAccountUpdates", _
                err.Description
End Sub

Friend Sub requestAllOpenOrders()

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::requestAllOpenOrders", _
                "Not connected"
End If
                            

Const version = 1

addData REQ_ALL_OPEN_ORDERS
addData version
send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "requestAllOpenOrders", _
                err.Description
End Sub

Friend Sub requestAutoOpenOrders(ByVal autoBind As Boolean)

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::requestAutoOpenOrders", _
                "Not connected"
End If

Const version = 1

addData REQ_AUTO_OPEN_ORDERS
addData version
addBoolean autoBind
send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "requestAutoOpenOrders", _
                err.Description
End Sub

Friend Sub RequestContract( _
                ByVal pContracts As TradeBuildSP.IContracts, _
                ByVal requester As ContractDetailsRequester)

Dim contractSpecifier As TradeBuildSP.IContractSpecifier
Dim requestQueueEntry As TWSContractDetailsRequestQueueEntry

On Error GoTo err
Set requestQueueEntry.tradeBuildContracts = pContracts
Set requestQueueEntry.requester = requester

Set contractSpecifier = pContracts.contractSpecifier
Select Case contractSpecifier.sectype
Case SecurityTypes.SecTypeBag
Case SecurityTypes.SecTypeCash
Case SecurityTypes.SecTypeFuture
Case SecurityTypes.SecTypeFuturesOption
Case SecurityTypes.SecTypeIndex
Case SecurityTypes.SecTypeOption
Case SecurityTypes.SecTypeStock
Case Else
    requester.contractSpecifierInvalid requestQueueEntry, "security type must be supplied"
    Exit Sub
End Select

If contractSpecifier.localSymbol = "" And _
    contractSpecifier.symbol = "" _
Then
    requester.contractSpecifierInvalid requestQueueEntry, "either symbol or local symbol must be supplied"
    Exit Sub
End If
    
mTwsContractDetailsRequestQueue.addToEnd requestQueueEntry

submitNextContractDetailsRequest
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestContract", _
                err.Description
End Sub

Friend Sub RequestExecutions(ByVal filter As ExecutionFilter)
Dim theFilter As ExecutionFilter

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::RequestExecutions", _
                "Not connected"
End If

Const version = 2

addData REQ_EXECUTIONS
addData version

If mServerVersion >= 9 Then
    If filter Is Nothing Then
        Set theFilter = New ExecutionFilter
    Else
        Set theFilter = filter
    End If
    
    With theFilter
        addData .clientID
        addData .accountCode
        addData IIf(CLng(.time) <> 0, Format(.time, "yyyymmdd-hh:nn:ss"), "")
        addData .symbol
        addData secTypeToString(.sectype)
        addData .exchange
        addData orderActionToString(.action)
    End With
End If

send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestExecutions", _
                err.Description
End Sub

Friend Sub requestFA(ByVal DataType As FADataTypes)

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::requestFA", _
                "Not connected"
End If

Const version = 1

If mServerVersion < 13 Then
    err.Raise ErrorCodes.ErrUnsupportedOperationException, _
                "IBTWSSP.TWSAPI::requestFA", _
                "Server version does not support requestFA"
End If

addData REQ_FA
addData version
addData DataType

send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "requestFA", _
                err.Description
End Sub

Friend Sub requestManagedAccounts()

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::requestManagedAccounts", _
                "Not connected"
End If

Const version = 1

addData REQ_MANAGED_ACCTS
addData version
send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "requestManagedAccounts", _
                err.Description
End Sub

Friend Function RequestMarketData(ByVal pContractSpecifier As TradeBuildSP.IContractSpecifier, _
                            ByVal dataReader As RealtimeDataReader, _
                            ByVal dataConsumer As IStreamingDataConsumer, _
                            Optional ByVal includeMarketDepth As Boolean = False) As Long

Dim lTickerID As Long

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::RequestMarketDataEx", _
                "Not connected"
End If

pContractSpecifier.Locked = True

lTickerID = AllocateTickerTableEntry(dataConsumer)
RequestMarketData = lTickerID
Set mTickerTable(lTickerID).contractSpecifier = pContractSpecifier
Set mTickerTable(lTickerID).dataReader = dataReader
reqMarketData lTickerID

If includeMarketDepth Then
    mTickerTable(lTickerID).receivingMarketDepth = True
    reqMarketDepth lTickerID, pContractSpecifier
End If
Exit Function
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestMarketData", _
                err.Description

End Function

Friend Sub RequestMarketDepth(ByVal pTickerID As Long)
                            
On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::RequestMarketDepth", _
                "Not connected"
End If

With mTickerTable(pTickerID)
    If .contractSpecifier Is Nothing Then err.Raise ErrorCodes.ErrIllegalArgumentException, "IBTWSSP.TWSAPI::requestMarketDepth", "Invalid ticker id: " & pTickerID
    .receivingMarketDepth = True
    reqMarketDepth pTickerID, .contractSpecifier
End With
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestMarketDepth", _
                err.Description

End Sub

Friend Function RequestHistoricalData( _
                barSpecifier As TradeBuildSP.BarDataSpecifier, _
                ByVal historicalDataReader As HistDataReader) As Long

Dim requestID As Long
Dim i As Long

On Error GoTo err
requestID = -1

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::RequestHistoricalData", _
                "Not connected"
End If

If mServerVersion < 16 Then
    err.Raise ErrorCodes.ErrUnsupportedOperationException, _
                "IBTWSSP.TWSAPI::RequestHistoricalData", _
                "Server version does not support historical data retrieval"
End If

Select Case barSpecifier.BarLengthMinutes
Case 1
Case 2
Case 5
Case 15
Case 30
Case 60
Case 1440
Case Else
    historicalDataReader.historicalDataRequestInvalid barSpecifier.NumberOfBars & "-minute bars not supported"
    Exit Function
End Select

For i = 0 To UBound(mHistoricalDataTable)
    If Not mHistoricalDataTable(i).inUse And _
        Not (mHistoricalDataTable(i).cancelling And Now < mHistoricalDataTable(i).earliestReuseTime) _
    Then
        requestID = i
        Exit For
    End If
Next

If requestID = -1 Then
    requestID = UBound(mHistoricalDataTable) + 1
    ReDim Preserve mHistoricalDataTable(UBound(mHistoricalDataTable) + 10) As HistoricalDataRequestTableEntry
End If

mHistoricalDataTable(i).inUse = True
mHistoricalDataTable(i).cancelling = False
mHistoricalDataTable(i).barSpecifier = barSpecifier
mHistoricalDataTable(i).nextBarIndex = 0
mHistoricalDataTable(i).barType = barSpecifier.barType
mHistoricalDataTable(i).barLength = barSpecifier.BarLengthMinutes
mHistoricalDataTable(i).nextBarIndex = -1
Set mHistoricalDataTable(i).historicalDataReader = historicalDataReader
Set mHistoricalDataTable(i).bars = New LinkedList
mHistoricalDataTable(i).bars.initialCapacity = IIf(barSpecifier.NumberOfBars <> 0, barSpecifier.NumberOfBars, 50)

generateTwsHistoricalDataRequest requestID
RequestHistoricalData = requestID
Exit Function
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestHistoricalData", _
                err.Description
End Function

Friend Sub RequestNewsBulletins(allMsgs As Boolean)

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::RequestNewsBulletins", _
                "Not connected"
End If

Const version = 1

addData REQ_NEWS_BULLETINS
addData version
addData allMsgs
send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestNewsBulletins", _
                err.Description
End Sub

Friend Sub RequestOpenOrders()

On Error GoTo err
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::RequestOpenOrders", _
                "Not connected"
End If

Const version = 1

addData REQ_OPEN_ORDERS
addData version
send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestOpenOrders", _
                err.Description
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub addBoolean(ByVal val As Boolean)
addData IIf(val, 1, 0)
End Sub

Private Sub addData(ByVal data As String)
Dim i As Long

If mIndexOut + Len(data) + 1 > UBound(mBufferOut) Then
    ReDim Preserve mBufferOut(UBound(mBufferOut) + 1024) As Byte
End If

For i = 1 To Len(data)
    mBufferOut(mIndexOut) = Asc(Mid$(data, i, 1))
    mIndexOut = mIndexOut + 1
Next
mBufferOut(mIndexOut) = 0
mIndexOut = mIndexOut + 1
End Sub

Private Sub addExchangeCode(ByVal code As String)
mMaxExchangeCodesIndex = mMaxExchangeCodesIndex + 1
If mMaxExchangeCodesIndex > UBound(mExchangeCodes) Then
    ReDim Preserve mExchangeCodes(UBound(mExchangeCodes) + 10) As String
End If
mExchangeCodes(mMaxExchangeCodesIndex) = UCase$(code)
End Sub

Private Function adjustPrice( _
                ByVal price As Double, _
                ByVal pContractSpecifier As IContractSpecifier) As Double
If pContractSpecifier.currencyCode = "GBP" And _
    mServerVersion < 23 _
Then
    adjustPrice = price * 100
Else
    adjustPrice = price
End If
End Function

Private Function AllocateTickerTableEntry(ByVal pDataConsumer As IStreamingDataConsumer) As Long
Dim i As Long

AllocateTickerTableEntry = -1

For i = 0 To UBound(mTickerTable)
    If Not mTickerTable(i).inUse Then
        AllocateTickerTableEntry = i
        Exit For
    End If
Next

If AllocateTickerTableEntry = -1 Then
    AllocateTickerTableEntry = UBound(mTickerTable) + 1
    ReDim Preserve mTickerTable(UBound(mTickerTable) + 100) As TickerTableEntry
'    mOrderSimulator.MaxTickers = UBound(mTickerTable) + 1
'    mTickers.MaxTickers = UBound(mTickerTable) + 1
End If

mTickerTable(AllocateTickerTableEntry).inUse = True
Set mTickerTable(AllocateTickerTableEntry).dataConsumer = pDataConsumer

End Function

Private Function calcConnectionString() As String
calcConnectionString = "server=" & mServer & _
                        "; port=" & mPort & _
                        "; client Id=" & mClientID
End Function

Private Sub cancelMarketDataEx(ByVal tickerid As Long, _
                ByVal releaseEntry As Boolean)

If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

Dim id As Long
Const version = 2

If Not mTickerTable(tickerid).inUse Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "IBTWSSP.TWSAPI::CancelMarketDataEx", _
    "Invalid ticker id " & tickerid
End If

id = tickerid + mTickerTable(tickerid).incarnation

If mTickerTable(tickerid).receivingMarketDepth Then CancelMarketDepth tickerid

If releaseEntry Then releaseTickerTableEntry tickerid

addData CANCEL_MKT_DATA
addData version
addData id
send

End Sub

Private Sub cancelContractRequests(ByVal reason As String)
Dim requestEntry As TWSContractDetailsRequestQueueEntry
Do While mTwsContractDetailsRequestQueue.size <> 0
    requestEntry = mTwsContractDetailsRequestQueue.removeFromFront
    requestEntry.requester.cancelRequest requestEntry, reason
Loop
End Sub
Private Sub checkmessages()
Static checkPointNumber As Long
Static msgID As TWSSocketInMsgTypes
Static version As Long
Static retryCount As Long
Static eventCount As Long
Static mElapsedTimer As TimerUtils.ElapsedTimer

If mElapsedTimer Is Nothing Then
    Set mElapsedTimer = New TimerUtils.ElapsedTimer
    mElapsedTimer.StartTiming
End If

On Error GoTo err

If mConnectionState <> ConnectionStates.ConnConnected Then
    ' first data received is plain server version number
    If checkPointNumber = 0 Then
        mServerVersion = getLong
        checkpointBuffer
        checkPointNumber = 100
        Debug.Print "Set checkpoint number=" & 100
    End If
    
    If checkPointNumber = 100 Then
        If mServerVersion >= 20 Then getString  ' TWS time
        checkpointBuffer
        checkPointNumber = 200
        Debug.Print "Set checkpoint number=" & 200
    End If
    
    mConnectionState = ConnectionStates.ConnConnected
    logMessage "Connection to TWS successful"
    
    ' set the log level in TWS
    setTWSLogLevel mTWSLogLevel
    
    ' now restart any tickers that were previously in operation
    reEstablishMarketData True
    
    ' and re-request any market depth
    resetMarketDepth True
    
    ' check for pending contract details requests
    submitNextContractDetailsRequest
        
    
    RaiseEvent Connected
    
    checkPointNumber = 1000
    Debug.Print "Set checkpoint number=" & 1000
End If

Do While True
    
    If checkPointNumber = 1000 Then
        msgID = getLong
        checkpointBuffer
        checkPointNumber = 1100
        Debug.Print "Set checkpoint number=" & 1100
    End If
    If checkPointNumber = 1100 Then
        version = getLong
        checkpointBuffer
        checkPointNumber = 1200
        Debug.Print "Set checkpoint number=" & 1200
    End If
    
    Select Case msgID
    Case TWSSocketInMsgTypes.TICK_PRICE
        Debug.Print "getTickPrice"
        getTickPrice version
    Case TWSSocketInMsgTypes.TICK_SIZE
        Debug.Print "getTickSize "
        getTickSize version
    Case TWSSocketInMsgTypes.ORDER_STATUS
        Debug.Print "getOrderStatus"
        getOrderStatus version
    Case TWSSocketInMsgTypes.ERR_MSG
        Debug.Print "getErrorMsg "
        getErrorMsg version
    Case TWSSocketInMsgTypes.OPEN_ORDER
        Debug.Print "getOpenOrder "
        getOpenOrder version
    Case TWSSocketInMsgTypes.ACCT_VALUE
        Debug.Print "getAccountValue "
        getAccountValue version
    Case TWSSocketInMsgTypes.PORTFOLIO_VALUE
        Debug.Print "getPortfolioValue "
        getPortfolioValue version
    Case TWSSocketInMsgTypes.ACCT_UPDATE_TIME
        Debug.Print "getAccountTime "
        getAccountTime version
    Case TWSSocketInMsgTypes.NEXT_VALID_ID
        Debug.Print "getNextValidId"
        getNextValidId version
    Case TWSSocketInMsgTypes.CONTRACT_DATA
        Debug.Print "getContractData"
        getContractData version
    Case TWSSocketInMsgTypes.EXECUTION_DATA
        Debug.Print "getExecutionData "
        getExecutionData version
    Case TWSSocketInMsgTypes.MARKET_DEPTH
        Debug.Print "getMarketDepth"
        getMarketDepth version
    Case TWSSocketInMsgTypes.MARKET_DEPTH_L2
        Debug.Print "getMarketDepthL2 "
        getMarketDepthL2 version
    Case TWSSocketInMsgTypes.NEWS_BULLETINS
        Debug.Print "getNewsBulletins "
        getNewsBulletins version
    Case TWSSocketInMsgTypes.MANAGED_ACCTS
        Debug.Print "getManagedAccounts "
        getManagedAccounts version
    Case TWSSocketInMsgTypes.RECEIVE_FA
        Debug.Print "getFAData "
        getFAData version
    Case TWSSocketInMsgTypes.HISTORICAL_DATA
        Debug.Print "getHistoricalData "
        getHistoricalData version
    Case TWSSocketInMsgTypes.BOND_CONTRACT_DATA
        Debug.Print "getBondContractData "
        getBondContractData version
    Case TWSSocketInMsgTypes.SCANNER_PARAMETERS
        Debug.Print "getScannerParameters "
        getScannerParameters version
    Case TWSSocketInMsgTypes.SCANNER_DATA
        Debug.Print "getScannerData "
        getScannerData version
    Case Else
        Dim s As String
        Dim i As Long
        Dim j As Long
        s = "Invalid message code from TWS: " & msgID & vbCrLf & _
            "Buffer contents to failure point: " & vbCrLf
        Do While i < mInputIndex
            s = s & Format(i, "0000") & "  "
            For j = i To i + 50
                If j = mInputIndex Then Exit For
                s = s & IIf(mBufferIn(j) <> 0, Chr$(mBufferIn(j)), "_")
            Next
            i = i + 50
            If j < mInputIndex Then s = s & vbCrLf
        Loop
        disconnect s
        Debug.Print s
        Exit Sub
    End Select
    retryCount = 0
    
    eventCount = eventCount + 1
    
    If mElapsedTimer.ElapsedTimeMicroseconds >= 10000000 Then
        logMessage "Event rate per second = " & Format(eventCount / 10, "0.0"), LogLevelMedium
        Debug.Print "Event rate per second = " & Format(eventCount / 10, "0.0")
        eventCount = 0
        mElapsedTimer.StartTiming
    End If
    
    checkpointBuffer
    checkPointNumber = 1000
    Debug.Print "Set checkpoint number=" & 1000
Loop

Exit Sub

err:
If err.number = DataIncomplete Then
    rollbackBuffer
    retryCount = retryCount + 1
    Debug.Print "Data incomplete count: " & retryCount & "; CheckPointNumber: " & checkPointNumber
    Exit Sub
End If

mCommonServiceConsumer.ServiceProviderError err.number, err.Description, "IBTWSSP"
disconnect "Error " & err.number & ": " & err.Description
End Sub

Private Sub checkpointBuffer()
mInputIndexChk = mInputIndex
End Sub

Private Sub createsocket()
If mSocket Is Nothing Then
    Set mSocket = New MSWinsockLib.Winsock
End If
End Sub

Private Sub generateTwsHistoricalDataRequest(ByVal index As Long)
Dim TwsRequest As TWSHistoricalDataRequestQueueEntry
Dim reqNum As Long
Dim totalMinsReqd As Long

reqNum = mHistoricalDataTable(index).numberOfTwsRequests + 1
mHistoricalDataTable(index).numberOfTwsRequests = reqNum

TwsRequest.id = index + HistoricalDataTickerBase + (mHistoricalDataTable(index).numberOfTwsRequests - 1) * HistoricalDataTickerIncrement

ReDim mHistoricalDataTable(index).requestKeys(reqNum - 1) As String

Set TwsRequest.contractSpec = mHistoricalDataTable(index).barSpecifier.contract.specifier

If mServerVersion >= 20 Then
    TwsRequest.endDateTime = Format(mHistoricalDataTable(index).barSpecifier.To, "yyyymmdd hh:nn:ss")
    
    Select Case mHistoricalDataTable(index).barSpecifier.BarLengthMinutes
    Case 1
        TwsRequest.barSizeSetting = TWSBarSizes.BarSize1Minute
    Case 2
        TwsRequest.barSizeSetting = TWSBarSizes.BarSize2Minutes
    Case 5
        TwsRequest.barSizeSetting = TWSBarSizes.BarSize5Minutes
    Case 15
        TwsRequest.barSizeSetting = TWSBarSizes.BarSize15Minutes
    Case 30
        TwsRequest.barSizeSetting = TWSBarSizes.BarSize30Minutes
    Case 60
        TwsRequest.barSizeSetting = TWSBarSizes.BarSize1Hour
    Case 1440
        TwsRequest.barSizeSetting = TWSBarSizes.BarSize1Day
    End Select
End If

' durationStr

totalMinsReqd = mHistoricalDataTable(index).barSpecifier.NumberOfBars * mHistoricalDataTable(index).barSpecifier.BarLengthMinutes

If totalMinsReqd <= 1440 Then
    TwsRequest.duration = totalMinsReqd * 60 & " S"
ElseIf mHistoricalDataTable(index).barSpecifier.BarLengthMinutes < 15 Then
    TwsRequest.duration = "86400 S"
ElseIf mHistoricalDataTable(index).barSpecifier.BarLengthMinutes < 1440 Then
    TwsRequest.duration = "5 D"
ElseIf totalMinsReqd <= 7200 Then
    TwsRequest.duration = Int((totalMinsReqd + 1439) / 1440) & " D"
Else
    TwsRequest.duration = "5 D"
End If

Select Case mHistoricalDataTable(index).barSpecifier.barType
Case TradeBuildSP.BarTypes.Ask
    TwsRequest.whatToShow = TWSWhatToShowAsk
Case TradeBuildSP.BarTypes.Bid
    TwsRequest.whatToShow = TWSWhatToShowBid
Case TradeBuildSP.BarTypes.Trade
    TwsRequest.whatToShow = TWSWhatToShowTrades
Case Else
End Select

mHistoricalDataTable(index).requestKeys(reqNum - 1) = queueTWSHistoricalDataRequest(TwsRequest)

End Sub

Private Function generateTwsOrderKey( _
                ByVal twsOrderId As Long, _
                ByVal clientID As Long) As String
generateTwsOrderKey = clientID & "/" & twsOrderId
End Function

Private Function getBoolean() As Boolean
Dim s As String
s = getString
If s = "1" Then
    getBoolean = True
Else
    getBoolean = False
End If
End Function

Private Function getDouble() As Double
Dim s As String
s = getString
If s = "" Then
    getDouble = 0#
Else
    getDouble = CDbl(s)
End If
End Function

Private Function getInteger() As Integer
Dim s As String
s = getString
If s = "" Then
    getInteger = 0
Else
    getInteger = CInt(s)
End If
End Function

Private Function getLong() As Long
Dim s As String
s = getString
If s = "" Then
    getLong = 0&
Else
    getLong = CLng(s)
End If
End Function

Private Function getSingle() As Single
Dim s As String
s = getString
If s = "" Then
    getSingle = 0!
Else
    getSingle = CSng(s)
End If
End Function

Private Function getString() As String
Dim s As String
Dim i As Long

For i = mInputIndex To mIndexIn - 1
    If mBufferIn(i) = 0 Then
        ' we've found our next string
        mInputIndex = i + 1
        getString = s
        'debug.Print "Got data: " & s
        'debug.Print "mIndexIn=" & mIndexIn & "  mInputIndex=" & mInputIndex
        Exit Function
    End If
    s = s & Chr$(mBufferIn(i))
Next

' we've got to the end of the current data, so raise an error to abort
' this attempt to assemble a message
err.Raise DataIncomplete
End Function

Private Sub getAccountTime(ByVal version As Long)
Dim listendata As TWSSocketMsgAcctUpdateTime
Dim timestamp As Date: timestamp = gCurrentTime
Dim accountTime As String: accountTime = getString

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.ACCT_UPDATE_TIME
    listendata.accountTime = accountTime
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgAcctUpdateTime, _
                    Me
End If

End Sub

Private Sub getAccountValue(ByVal version As Long)
Dim listendata As TWSSocketMsgAcctValue
Dim timestamp As Date: timestamp = gCurrentTime
Dim key As String: key = getString
Dim val As String: val = getString
Dim cur As String: cur = getString
Dim accountName As String: If version >= 2 Then accountName = getString

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.ACCT_VALUE
    listendata.accountName = accountName
    listendata.currencyCode = cur
    listendata.key = key
    listendata.value = val
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgAcctValue, _
                    Me
End If

End Sub

Private Sub getBondContractData(ByVal version As Long)
' not currently supported
getString 'symbol
getString 'sectype
getString 'cusip
getDouble 'coupon
getString 'maturity
getString 'issueDate
getString 'ratings
getString 'bondType
getString 'couponType
getBoolean 'convertible
getBoolean 'callable
getBoolean 'putable
getString 'descAppend
getString 'exchange
getString 'currency
getString 'marketName
getString 'tradingClass
getLong 'conid
getDouble 'minTick
getString 'OrderTypes
getString 'validExchanges

mContractDetailsRequest.requester.contractSpecifierInvalid mContractDetailsRequest, _
                                                    "bonds not yet supported"
                                                    

End Sub

Private Sub getContractData(ByVal version As Long)
Dim listendata As TWSSocketMsgContractData
Dim timestamp As Date
Dim lContract As contract
Dim minimumTick As Double
Dim multiplier As Long
Dim lOrderType As TradeBuildSP.OrderTypes
Dim OrderTypes As String
Dim orderTypesStr() As String
Dim permittedOrderTypes As Long
Dim validExchanges As String
Dim lValidExchanges() As String
Dim permittedOrderTifs As Long
Dim permittedOrderAttributes As Long
Dim expiryString As String
Dim i As Long

timestamp = gCurrentTime

Set lContract = New contract
lContract.tradeBuildContract = mCommonServiceConsumer.NewContract
With lContract.tradeBuildContract
    .specifier = mCommonServiceConsumer.NewContractSpecifier
    .specifier.symbol = getString
    .specifier.sectype = secTypeFromString(getString)
    expiryString = getString
    .specifier.expiry = expiryString
    .specifier.strike = getDouble
    .specifier.right = optionRightFromString(getString)
    .specifier.exchange = getString
    .specifier.currencyCode = getString
    .specifier.localSymbol = getString
    If expiryString <> "" Then
        .ExpiryDate = CDate(Left$(expiryString, 4) & "/" & Mid$(expiryString, 5, 2) & "/" & right$(expiryString, 2))
    End If
    .marketName = getString
    .tradingClass = getString
    .contractID = getLong
    minimumTick = getDouble
    multiplier = getLong
    
    OrderTypes = getString
    orderTypesStr = Split(OrderTypes, ",")
    
    permittedOrderAttributes = OrderAttributes.OrderAttBlockOrder
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttMinimumQuantity
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttNBBOPriceCap
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttOriginatorRef
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttOverrideConstraints
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttPercentOffset
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttRule80A
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttAction
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttLimitPrice
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttOrderType
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttQuantity
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttTimeInForce
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttTriggerPrice
    permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttStopTriggerMethod
    
    For i = 0 To UBound(orderTypesStr)
        lOrderType = orderTypeFromString(orderTypesStr(i))
        
        If Not lOrderType = TradeBuildSP.OrderTypes.OrderTypeNone Then
            permittedOrderTypes = permittedOrderTypes Or lOrderType
        Else
            Select Case orderTypesStr(i)
            Case "AON"  ' all or none
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttAllOrNone
            Case "DAY"  ' today only
                permittedOrderTifs = permittedOrderTifs Or OrderTifs.TIFDay
            Case "DIS"  ' discretionary amount
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttDiscretionaryAmount
            Case "ELECONLY"  ' electronic trade only
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttETradeOnly
            Case "FIRMONLY"  ' firm quote only
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttFirmQuoteOnly
            Case "GAT"  ' good after time
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttGoodAfterTime
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttGoodAfterTimeTZ
            Case "GTC"  ' good till cancelled
                permittedOrderTifs = permittedOrderTifs Or OrderTifs.TIFGoodTillCancelled
            Case "GTD"  ' good till date
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDate
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDateTZ
            Case "GTT"  ' good till time
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDate
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttGoodTillDateTZ
            Case "HID"  ' hidden?
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttHidden
            Case "ICE"  ' iceberg
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttDisplaySize
            Case "IOC"  ' immediate or cancel
                permittedOrderTifs = permittedOrderTifs Or OrderTifs.TIFImmediateOrCancel
            Case "RTH"
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttIgnoreRTH
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttRTHOnly
            Case "SWEEP"
                permittedOrderAttributes = permittedOrderAttributes Or OrderAttributes.OrderAttSweepToFill
            Case "BASKET", "COND", "OCA", "OPG", "OPGREROUT", "TIMEPRIO"
                ' not sure what these signify
            Case Else
                If orderTypesStr(i) <> "" Then
                    logMessage "Unknown order type " & orderTypesStr(i) & " received from TWS", LogLevelLow
                End If
            End Select
        End If
    Next
    
    lContract.permittedOrderAttributes = permittedOrderAttributes
    lContract.permittedOrderTypes = permittedOrderTypes
    
    lContract.permittedOrderTifs = permittedOrderTifs
    
    lContract.permittedStopTriggerMethods = TradeBuildSP.StopTriggerMethods.StopTriggerDefault Or _
                                            TradeBuildSP.StopTriggerMethods.StopTriggerDoubleBidAsk Or _
                                            TradeBuildSP.StopTriggerMethods.StopTriggerDoubleLast Or _
                                            TradeBuildSP.StopTriggerMethods.StopTriggerLast

    validExchanges = getString
    lValidExchanges = Split(validExchanges, ",")
    .validExchanges = lValidExchanges
    
    If version >= 2 Then
        lContract.priceMagnifier = getLong
    Else
        If .specifier.currencyCode = "GBP" Then
            lContract.priceMagnifier = 100
        Else
            lContract.priceMagnifier = 1
        End If
    End If
    .multiplier = multiplier / lContract.priceMagnifier
    .minimumTick = minimumTick * lContract.priceMagnifier
End With

processContract lContract

If mLogLevel >= LogLevelHigh Then
    With lContract.tradeBuildContract
        listendata.msgType = TWSSocketInMsgTypes.CONTRACT_DATA
        listendata.contractID = .contractID
        listendata.currencyCode = .specifier.currencyCode
        listendata.exchange = .specifier.exchange
        listendata.expiry = .specifier.expiry
        listendata.localSymbol = .specifier.localSymbol
        listendata.marketName = .marketName
        listendata.minimumTick = minimumTick
        listendata.multiplier = multiplier
        listendata.OrderTypes = OrderTypes
        listendata.right = .specifier.right
        listendata.sectype = .specifier.sectype
        listendata.strike = .specifier.strike
        listendata.symbol = .specifier.symbol
        listendata.tradingClass = .tradingClass
        listendata.validExchanges = validExchanges
        listendata.version = version
    End With
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgContractData, _
                    Me
End If

End Sub

Private Sub getErrorMsg(ByVal version As Long)
Dim listendata As TWSSocketMsgErrMsg
Dim timestamp As Date
Dim id As Long
Dim errorCode As Long
Dim errorMsg As String
Dim execFilter As ExecutionFilter
Dim index As Long
Dim mapEntry As OrderIdMapEntry
Dim failpoint As Long

On Error GoTo err

failpoint = 100

id = getLong
Debug.Print "Error id: " & id

failpoint = 200

errorCode = getLong
Debug.Print "Error code: " & errorCode

failpoint = 300

errorMsg = getString
Debug.Print "Error msg: " & errorMsg
timestamp = gCurrentTime


Select Case errorCode
Case 101
    ' max tickers reached
    failpoint = 350
    mTickerTable(id - MARKET_DEPTH_TICKER_BASE).dataReader.marketDataRequestFailed "Maximum number of tickers reached"
Case 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, _
    121, 125, 126, 127, 128, 129, 130, 132, 137, 140, 141, 144, _
    152, 153, 154, 157, 160, 201, 312, 313, 314, 315, 325, 328, _
    334, 337, 339, 343
    
    failpoint = 400

    mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
    mapEntry.orderSubmitter.updateStatus mapEntry.order.tradeBuildId, _
                                        OrderStatuses.OrderStatusRejected
    mapEntry.orderSubmitter.invalidOrder mapEntry.order.tradeBuildId, errorMsg
Case 162    ' historical data request problem
    failpoint = 500

    index = (id - HistoricalDataTickerBase) Mod HistoricalDataTickerIncrement

    mHistoricalDataTable(index).historicalDataReader.historicalDataRequestFailed errorMsg
    releaseHistoricalDataTableEntry index
    processNextHistDataRequest
Case 165
    failpoint = 600

    logMessage "Connected to IB Historical Market Data Service"
Case 200    ' security definition not known
    If id >= 0 Then
        ' this must relate to a reqMktData - this can occur even though
        ' we make sure the contract is valid before starting the ticker because
        ' for example the contract may have expired
        
        failpoint = 700

        mTickerTable(id Mod INCARNATION_INCREMENT).dataReader.marketDataRequestFailed errorMsg
        releaseTickerTableEntry (id Mod INCARNATION_INCREMENT)
    Else
        ' this must relate to a RequestContract
        failpoint = 750

        processContractError errorMsg
    End If
Case 202
    failpoint = 800

    mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
    mapEntry.orderSubmitter.updateStatus mapEntry.order.tradeBuildId, _
                                        OrderStatuses.OrderStatusCancelled
    mOrderIDMap.remove generateTwsOrderKey(id, mClientID)
Case 309
    ' max market depth requests exceeded
    failpoint = 900

    mTickerTable(id - MARKET_DEPTH_TICKER_BASE).dataReader.marketDepthRequestFailed "Max market depth requests exceeded"
Case 316
    failpoint = 1000

    resetMarketDepth reEstablish:=True
Case 317
    failpoint = 1100

    resetMarketDepth reEstablish:=False
Case 321    ' invalid request
    failpoint = 1200

    If id >= HistoricalDataTickerBase And id < OrderIdBase Then
        index = (id - HistoricalDataTickerBase) Mod HistoricalDataTickerIncrement
    
        If index > 0 Then
            mHistoricalDataTable(id).historicalDataReader.historicalDataRequestInvalid errorMsg
            releaseHistoricalDataTableEntry index
            processNextHistDataRequest
        End If
    End If
Case 1100
    failpoint = 1300

    ' connectivity between TWS and IB has been lost
    RaiseEvent ConnectionToIBClosed
    logMessage "Connection to IB has been lost"
    
Case 1101
    ' connectivity between TWS and IB has been restored, but data has been
    ' lost, so need to reestablish market data and market depth requests
    failpoint = 1400

    
    reEstablishMarketData False
    resetMarketDepth True
    
    ' Now need to reassociate order ids with tws
    RequestOpenOrders
    
    ' Now make sure we're aware of any executions that occurred during
    ' the disconnection
    Set execFilter = New ExecutionFilter
    execFilter.clientID = mClientID
    RequestExecutions execFilter

    RaiseEvent ConnectionToIBRecovered
    
    logMessage "Connection to IB recovered: market data re-established"
    
Case 1102
    ' connectivity between TWS and IB has been restored without loss of data
    ' Now need to reassociate order ids with tws
    failpoint = 1500

    RequestOpenOrders
    
    ' Now make sure we're aware of any executions that occurred during
    ' the disconnection
    Set execFilter = New ExecutionFilter
    execFilter.clientID = mClientID
    RequestExecutions execFilter

    RaiseEvent ConnectionToIBRecovered
    
    logMessage "Connection to IB recovered: no loss of data"
    
Case 2103, 2104, 2105, 2106, 2107, 2108
    failpoint = 1600

    logMessage errorMsg
Case Else
    failpoint = 1700

    mCommonServiceConsumer.ServiceProviderError errorCode, errorMsg, SpName
End Select

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.ERR_MSG
    listendata.errorCode = errorCode
    listendata.errorMsg = errorMsg
    listendata.id = id
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgErrMsg, _
                    Me
End If

Exit Sub

err:
err.Raise err.number, , "TWSAPI::getErrorMsg(failpoint=" & failpoint & "): " & err.Description
End Sub

Private Sub getExecutionData(ByVal version As Long)
Dim listendata As TWSSocketMsgExecutionData
Dim timestamp As Date: timestamp = gCurrentTime
Dim timeString  As String
Dim sectype As String
Dim rightString As String
Dim side As String
Dim price As Double
Dim mapEntry As OrderIdMapEntry
Dim id As Long: id = getLong
Dim permId As Long
Dim clientID As Long
Dim lContractSpecifier As TradeBuildSP.IContractSpecifier

Set lContractSpecifier = mCommonServiceConsumer.NewContractSpecifier
With lContractSpecifier
    .Locked = True
    .symbol = getString
    sectype = getString
    .sectype = secTypeFromString(sectype)
    .expiry = getString
    .strike = getDouble
    rightString = getString
    .right = optionRightFromString(rightString)
    .exchange = getString
    .currencyCode = getString
    .localSymbol = getString
End With

On Error Resume Next
mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
On Error GoTo 0

Dim exec As TradeBuildSP.IExecution
Set exec = mCommonServiceConsumer.NewExecution
exec.contractSpecifier = lContractSpecifier
With exec
    .orderBrokerId = id
    If Not mapEntry.order Is Nothing Then .orderTradeBuildId = mapEntry.order.tradeBuildId
    .execId = getString
    timeString = getString
    .time = CDate(Left$(timeString, 4) & "/" & _
                    Mid$(timeString, 5, 2) & "/" & _
                    Mid$(timeString, 7, 2) & " " & _
                    right$(timeString, 8))
    .accountIdentifier = getString
    .fillingExchange = getString
    side = getString
    .action = IIf(side = "BUY", OrderActions.ActionBuy, OrderActions.ActionSell)
    .quantity = getLong
    price = getDouble
    .price = adjustPrice(price, lContractSpecifier)
End With

If version >= 2 Then permId = getLong  ' permId

If version >= 3 Then clientID = getLong  ' clientID
   
If version >= 4 Then exec.isLiquidation = getBoolean

Dim existingExec As TradeBuildSP.IExecution
On Error Resume Next
Set existingExec = mExecutions(exec.execId)
On Error GoTo 0
If Not existingExec Is Nothing Then
    ' we've had this execution before so nothing to do
Else
    If Not mapEntry.orderSubmitter Is Nothing Then
        mapEntry.orderSubmitter.executionDetails lContractSpecifier, _
                                                    exec
        ' notify a status update - this is in case this execution occurred during a period when
        ' TWS had lost its connection to the IB servers and this exec details is being notified
        ' in response to the reqExecutions call made when the connection is recovered. In these
        ' circumstances TWS will not itself issue an orderStatus.
        If exec.quantity >= mapEntry.order.quantityRemaining Then
            mapEntry.orderSubmitter.updateStatus mapEntry.order.tradeBuildId, _
                                                OrderStatuses.OrderStatusFilled
        Else
            mapEntry.orderSubmitter.updateStatus mapEntry.order.tradeBuildId, _
                                                OrderStatuses.OrderStatusSubmitted
        End If
    End If
End If

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.EXECUTION_DATA
    With lContractSpecifier
        listendata.currencyCode = .currencyCode
        listendata.exchange = .exchange
        listendata.expiry = .expiry
        listendata.localSymbol = .localSymbol
        listendata.right = rightString
        listendata.sectype = sectype
        listendata.strike = .strike
        listendata.symbol = .symbol
    End With
    With exec
        listendata.accountNumber = .accountIdentifier
        listendata.clientID = clientID
        listendata.execExchange = .fillingExchange
        listendata.execId = .execId
        listendata.orderId = id
        listendata.permId = permId
        listendata.price = price
        listendata.quantity = .quantity
        listendata.side = side
        listendata.time = timeString
    End With
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgExecutionData, _
                    Me
End If

End Sub

Private Sub getFAData(ByVal version As Long)
Dim listendata As TWSSocketMsgReceiveFa
Dim timestamp As Date: timestamp = gCurrentTime
Dim DataType As FADataTypes: DataType = getLong
Dim XMLData As String: XMLData = getString

RaiseEvent FAData(DataType, XMLData)

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.MANAGED_ACCTS
    listendata.DataType = DataType
    listendata.XMLData = XMLData
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgManagedAccts, _
                    Me
End If

End Sub

Private Sub getHistoricalData(ByVal version As Long)
Static checkPointNumber As Long
Static requestID As Long
Static index As Long
Static reqNo As Long
Static itemCount As Long
Static Bar As HistoricBar
Static i As Long
Static hint As Long

Dim bardate As String
Dim timestamp As Date

If checkPointNumber = 0 Then
    hint = 0
    i = 0
    requestID = getLong
    index = (requestID - HistoricalDataTickerBase) Mod HistoricalDataTickerIncrement
    reqNo = Int((requestID - HistoricalDataTickerBase) / HistoricalDataTickerIncrement)
    
    checkpointBuffer
    checkPointNumber = 100
End If

If checkPointNumber = 100 Then
    itemCount = getLong
    Debug.Print "Bars retrieved: " & itemCount
    
    checkpointBuffer
    checkPointNumber = 200
End If

If checkPointNumber >= 200 Then
    For i = i To itemCount - 1
        If checkPointNumber = 200 Then
            bardate = getString
            timestamp = CDate(Left$(bardate, 4) & "/" & _
                        Mid$(bardate, 5, 2) & "/" & _
                        Mid$(bardate, 7, 2) & " " & _
                        right$(bardate, 8))
            Bar.timestamp = timestamp
            
            If i = 0 Then
                mHistoricalDataTable(index).barSpecifier.To = timestamp - mHistoricalDataTable(index).barSpecifier.BarLengthMinutes / 1440
            End If
            
            checkpointBuffer
            checkPointNumber = 210
        End If
        If checkPointNumber = 210 Then
            Bar.openValue = getDouble
            checkpointBuffer
            checkPointNumber = 220
        End If
        If checkPointNumber = 220 Then
            Bar.highValue = getDouble
            checkpointBuffer
            checkPointNumber = 230
        End If
        If checkPointNumber = 230 Then
            Bar.lowValue = getDouble
            checkpointBuffer
            checkPointNumber = 240
        End If
        If checkPointNumber = 240 Then
            Bar.closeValue = getDouble
            checkpointBuffer
            checkPointNumber = 250
        End If
        If checkPointNumber = 250 Then
            Bar.volume = getLong
            checkpointBuffer
            checkPointNumber = 260
        End If
        If checkPointNumber = 260 Then
            Bar.WAP = getDouble
            checkpointBuffer
            checkPointNumber = 270
        End If
        If checkPointNumber = 270 Then
            Bar.hasGaps = getBoolean
            If Not mHistoricalDataTable(index).cancelling Then
                hint = mHistoricalDataTable(index).bars.add(Bar, Bar.timestamp, hint)
            End If
            checkpointBuffer
            checkPointNumber = 280
        End If
        
        checkPointNumber = 200
    Next
End If
checkPointNumber = 0


If mHistoricalDataTable(index).cancelling Then Exit Sub

mHistoricalDataTable(index).requestKeys(reqNo) = ""

mConcurrentHistDataRequests = mConcurrentHistDataRequests - 1
processNextHistDataRequest

mHistoricalDataTable(index).barSpecifier.NumberOfBars = mHistoricalDataTable(index).barSpecifier.NumberOfBars - itemCount
If mHistoricalDataTable(index).barSpecifier.NumberOfBars <= 0 Then
    Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
    mHistoricalDataTable(index).dataRetrievedFromTWS = True
    mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
    mHistoricalDataTable(index).historicalDataReader.historicalDataAvailable
Else
    ' request the next lot of bars
    generateTwsHistoricalDataRequest index
End If


End Sub

Private Sub getManagedAccounts(ByVal version As Long)
Dim listendata As TWSSocketMsgManagedAccts
Dim timestamp As Date: timestamp = gCurrentTime
Dim accountsList As String: accountsList = getString

RaiseEvent managedAccounts(timestamp, accountsList)

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.MANAGED_ACCTS
    listendata.accountsList = accountsList
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgManagedAccts, _
                    Me
End If

End Sub

Private Sub getMarketDepth(ByVal version As Long)
Dim listendata As TWSSocketMsgMarketDepth
Dim timestamp As Date
Dim id As Long: id = getLong
Dim position As Long: position = getLong
Dim operation As Long: operation = getLong
Dim side As Long: side = getLong
Dim price As Double: price = getDouble
Dim size As Long: size = getLong

timestamp = gCurrentTime

processMarketDepth timestamp, _
                id - MARKET_DEPTH_TICKER_BASE, _
                position, _
                operation, _
                side, _
                price, _
                size

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.MARKET_DEPTH
    listendata.operation = operation
    listendata.position = position
    listendata.price = price
    listendata.side = side
    listendata.size = size
    listendata.tickerid = id
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgMarketDepth, _
                    Me
End If

End Sub

Private Sub getMarketDepthL2(ByVal version As Long)
Dim listendata As TWSSocketMsgMarketDepthL2
Dim timestamp As Date: timestamp = gCurrentTime
Dim id As Long: id = getLong
Dim position As Long: position = getLong
Dim marketmaker As String: marketmaker = getString
Dim operation As Long: operation = getLong
Dim side As Long: side = getLong
Dim price As Double: price = getDouble
Dim size As Long: size = getLong

With mTickerTable(id - MARKET_DEPTH_TICKER_BASE)
    If .inUse Then
        
        .dataConsumer.UpdateMktDepthL2 timestamp, _
                        position, _
                        marketmaker, _
                        operation, _
                        side, _
                        price, _
                        size
                        
    End If
End With

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.MARKET_DEPTH_L2
    listendata.marketmaker = marketmaker
    listendata.operation = operation
    listendata.position = position
    listendata.price = price
    listendata.side = side
    listendata.size = size
    listendata.tickerid = id
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgMarketDepthL2, _
                    Me
End If

End Sub

Private Sub getNewsBulletins(ByVal version As Long)
Dim listendata As TWSSocketMsgNewsBulletins
Dim timestamp As Date: timestamp = gCurrentTime
Dim msgID As Long: msgID = getLong
Dim msgType As Long: msgType = getLong
Dim newsMessage As String: newsMessage = getString
Dim originatingExch As String: originatingExch = getString

RaiseEvent updateNewsBulletin(msgID, msgType, newsMessage, originatingExch)

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.NEWS_BULLETINS
    listendata.newsID = msgID
    listendata.newsMessage = newsMessage
    listendata.newsType = msgType
    listendata.originatingExch = originatingExch
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgNewsBulletins, _
                    Me
End If

End Sub

Private Sub getNextValidId(ByVal version As Long)
Dim listendata As TWSSocketMsgNextValidId
Dim timestamp As Date: timestamp = gCurrentTime

mNextOrderID = getLong

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.NEXT_VALID_ID
    listendata.id = mNextOrderID
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgNextValidId, _
                    Me
End If

If mNextOrderID < OrderIdBase Then mNextOrderID = OrderIdBase

End Sub

Private Sub getOpenOrder(ByVal version As Long)
Dim listendata As TWSSocketMsgOpenOrder
Dim timestamp As Date: timestamp = gCurrentTime
'Dim lOrder As Order
Dim id As Long
Dim sectype As String
Dim right As String
Dim action As String
Dim orderType As String
Dim timeInForce As String

id = getLong

'On Error Resume Next
'Set lOrder = mAllOrders(CStr(id))
'On Error GoTo 0
'
'If lOrder Is Nothing Then
'    ' this must be an open order from the previous API session, but
'    ' for which we haven't yet had the openOrder notification
    'Set lOrder = New Order
    'lOrder.id = id
'    mAllOrders.add lOrder, CStr(id)
'End If

' read contract fields
Dim lContractSpecifier As TradeBuildSP.IContractSpecifier
Set lContractSpecifier = mCommonServiceConsumer.NewContractSpecifier
With lContractSpecifier
    .Locked = True
    .symbol = getString
    sectype = getString
    .sectype = secTypeFromString(sectype)
    .expiry = getString
    .strike = getDouble
    right = getString
    .right = optionRightFromString(right)
    .exchange = getString
    .currencyCode = getString
End With

If version >= 2 Then lContractSpecifier.localSymbol = getString

' read order fields
'With lOrder
    action = getString
    '.action = orderActionFromString(action)
     getLong    ' .quantity
    orderType = getString
    '.orderType = orderTypeFromString(orderType)
    getDouble   ' .limitPrice
    getDouble   ' .auxPrice
    timeInForce = getString
    '.timeInForce = orderTIFFromString(timeInForce)
    getString   ' .ocaGroup
    getString   ' .account
    getString   ' .openClose
    getLong     ' .origin
    getString   ' .orderRef
    If version >= 3 Then getLong    ' .clientID
    
    If version >= 4 Then
        getLong     ' .permId =
        getBoolean  ' .ignoreRTH
        getBoolean  ' .hidden
        getDouble   ' .discretionaryAmt
    End If

    If version >= 5 Then getString  ' .goodAfterTime
    
    If version >= 6 Then getString  ' .sharesAllocation
    
    If version >= 7 Then
        getString   ' .FAGroup
        getString   ' .FAMethod
        getString   ' .FAPercentage
        getString   ' .FAProfile
    End If
    
    If version >= 8 Then getString  ' goodTillDate

    If version >= 9 Then
        getString           ' rule80A
        getDouble           ' percentOffset
        getString           ' settlingFirm
        getLong             ' shortSaleSlot
        getString           ' designatedLocation
        getLong             ' auctionStrategy
        getDouble           ' startingPrice
        getDouble           ' stockRefPrice
        getDouble           ' delta
        getDouble           ' stockRangeLower
        getDouble           ' stockRangeUpper
        getLong             ' displaySize
        getBoolean          ' rthOnly
        getBoolean          ' blockOrder
        getBoolean          ' sweepToFill
        getBoolean          ' allOrNone
        getLong             ' minQty
        getLong             ' ocaType
        getBoolean          ' eTradeOnly
        getBoolean          ' firmQuoteOnly
        getDouble           ' nbboPriceCap
    End If

    If version >= 10 Then
        getLong             ' parentId
        getLong             ' triggerMethod
    End If
'End With

'If mLogLevel >= LogLevelLow Then
'    mCommonServiceConsumer.NotifyListeners _
'            "Open order:" & _
'            " ID: " & lOrder.id & _
'            "; Act: " & orderActionToString(lOrder.action) & _
'            "; Qty: " & lOrder.quantity & _
'            "; Type: " & orderTypeToString(lOrder.orderType) & _
'            "; Lmt: " & lOrder.limitPrice & _
'            "; Trig: " & lOrder.auxPrice & _
'            "; TIF: " & orderTIFToString(lOrder.timeInForce) & _
'            "; OCA: " & lOrder.ocaGroup, _
'            TradeBuildSP.StandardListenValueTypes.Log, _
'            timestamp
'End If

'If mLogLevel >= LogLevelHigh Then
'    listendata.msgType = TWSSocketInMsgTypes.OPEN_ORDER
'    With lContractSpecifier
'        listendata.currencyCode = .currencyCode
'        listendata.exchange = .exchange
'        listendata.expiry = .expiry
'        listendata.localSymbol = .localSymbol
'        listendata.right = .right
'        listendata.sectype = .sectype
'        listendata.strike = .strike
'        listendata.symbol = .symbol
'    End With
'    With lOrder
'        listendata.account = .account
'        listendata.action = action
'        listendata.auxPrice = .auxPrice
'        listendata.clientID = .clientID
'        listendata.discretionaryAmt = .discretionaryAmt
'        listendata.FAGroup = .FAGroup
'        listendata.FAMethod = .FAMethod
'        listendata.FAPercentage = .FAPercentage
'        listendata.FAProfile = .FAProfile
'        listendata.goodAfterTime = .goodAfterTime
'        listendata.hidden = .hidden
'        listendata.ignoreRTH = .ignoreRTH
'        listendata.limitPrice = .limitPrice
'        listendata.ocaGroup = .ocaGroup
'        listendata.openClose = .openClose
'        listendata.orderId = id
'        listendata.orderRef = .orderRef
'        listendata.orderType = orderType
'        listendata.origin = .origin
'        listendata.permId = .permId
'        listendata.quantity = .quantity
'        listendata.sharesAllocation = .sharesAllocation
'        listendata.timeInForce = timeInForce
'    End With
'    listendata.version = version
'
'    mCommonServiceConsumer.NotifyListeners listendata, _
'                    ListenValueTypes.ValueTypeTWSSocketMsgOpenOrder, _
'                    timestamp
'End If

End Sub

Private Sub getOrderStatus(ByVal version As Long)
Dim mapEntry As OrderIdMapEntry
Dim listendata As TWSSocketMsgOrderStatus

Dim timestamp As Date: timestamp = gCurrentTime
Dim id As Long: id = getLong
Dim statusString As String: statusString = getString
Dim status As TradeBuildSP.OrderStatuses: status = orderStatusFromString(statusString)
Dim filled As Long: filled = getLong
Dim remaining As Long: remaining = getLong
Dim avgFillPrice As Double:  avgFillPrice = getDouble
Dim lastFillPrice As Double
Dim clientID As Long
Dim permId As Long
Dim parentId As Long

If version >= 2 Then permId = getLong

If version >= 3 Then parentId = getLong

If version >= 4 Then lastFillPrice = getDouble

If version >= 5 Then clientID = getLong

On Error Resume Next
mapEntry = mOrderIDMap(generateTwsOrderKey(id, clientID))
On Error GoTo 0
If Not mapEntry.order Is Nothing Then
    avgFillPrice = adjustPrice(avgFillPrice, mapEntry.order.contract.specifier)
    lastFillPrice = adjustPrice(lastFillPrice, mapEntry.order.contract.specifier)
    
    mapEntry.orderSubmitter.updateStatus mapEntry.order.tradeBuildId, _
                            status
    If status = OrderStatusCancelled Or status = OrderStatusFilled Then
        mOrderIDMap.remove generateTwsOrderKey(id, clientID)
    End If
Else
    ' this must be an open order from the previous API session, but
    ' for which we haven't yet had the openOrder notification
    ' don't notify the service consumer because it won't know about this
    ' order yet
    '
    ' or it's a cancelled status for an order which has already been
    ' notified as cancelled because of an error 202
    '
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! need to improve this!!!!!!!!!!!!!!!
End If

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.ORDER_STATUS
    listendata.avgFillPrice = avgFillPrice
    listendata.clientID = clientID
    listendata.filled = filled
    listendata.lastFillPrice = lastFillPrice
    listendata.orderId = id
    listendata.parentId = parentId
    listendata.permId = permId
    listendata.remaining = remaining
    listendata.status = statusString
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgOrderStatus, _
                    Me
End If

End Sub

Private Sub getPortfolioValue(ByVal version As Long)
Dim listendata As TWSSocketMsgPortfolioValue
Dim timestamp As Date: timestamp = gCurrentTime
Dim sectype As String
Dim right As String
Dim lContractSpecifier As TradeBuildSP.IContractSpecifier

Set lContractSpecifier = mCommonServiceConsumer.NewContractSpecifier
With lContractSpecifier
    .Locked = True
    .symbol = getString
    sectype = getString
    .sectype = secTypeFromString(sectype)
    .expiry = getString
    .strike = getDouble
    right = getString
    .right = optionRightFromString(right)
    .currencyCode = getString
    If version >= 2 Then .localSymbol = getString
End With

Dim position As Long: position = getLong
Dim marketPrice As Double: marketPrice = getDouble
Dim marketValue As Double: marketValue = getDouble
Dim averageCost As Double: averageCost = 0#
Dim unrealizedPNL As Double: unrealizedPNL = 0#
Dim realizedPNL As Double: realizedPNL = 0#
If version >= 3 Then
    averageCost = getDouble
    unrealizedPNL = getDouble
    realizedPNL = getDouble
End If
Dim accountName As String: If version >= 4 Then accountName = getString
                         
RaiseEvent updatePortfolioValue(lContractSpecifier, position, marketPrice, marketValue, averageCost, unrealizedPNL, realizedPNL, accountName)

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.PORTFOLIO_VALUE
    With lContractSpecifier
        listendata.currencyCode = .currencyCode
        listendata.expiry = .expiry
        listendata.localSymbol = .localSymbol
        listendata.right = right
        listendata.sectype = sectype
        listendata.strike = .strike
        listendata.symbol = .symbol
    End With
    listendata.accountName = accountName
    listendata.averageCost = averageCost
    listendata.marketPrice = marketPrice
    listendata.marketValue = marketValue
    listendata.position = position
    listendata.realizedPNL = realizedPNL
    listendata.unrealizedPNL = unrealizedPNL
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgPortfolioValue, _
                    Me
End If

End Sub

Private Sub getScannerData(ByVal version As Long)
Dim contract As TradeBuildSP.IContract
Dim tickerid As Long: tickerid = getLong
Dim numberOfElements As Long: numberOfElements = getLong
Dim rank As Long
Dim distance As String
Dim benchmark As String
Dim projection As String
Dim i As Long

For i = 0 To numberOfElements - 1
    Set contract = mCommonServiceConsumer.NewContract
    rank = getLong
    With contract.specifier
        .symbol = getString
        .sectype = secTypeFromString(getString)
        .expiry = getString
        .strike = getDouble
        .right = optRightFromString(getString) 'right
        .exchange = getString
        .currencyCode = getString
        .localSymbol = getString
    End With
    contract.marketName = getString
    contract.tradingClass = getString
    distance = getString
    benchmark = getString
    projection = getString
Next
End Sub

Private Sub getScannerParameters(ByVal version As Long)
Dim xml  As String
xml = getString
End Sub

Private Sub getTickPrice(ByVal version As Long)
Dim listendata As TWSSocketMsgTickPrice
Dim timestamp As Date
Dim price As Double
Dim size As Long
Dim tickType As Long
Dim tickerid As Long
Dim canAutoExecute As Boolean
Dim failpoint As Long

On Error GoTo err

failpoint = 100

tickerid = getLong

failpoint = 200

tickType = getLong

failpoint = 300

price = getDouble

failpoint = 400

If version >= 2 Then size = getLong
If version >= 3 Then canAutoExecute = getBoolean

timestamp = gCurrentTime

failpoint = 500

processTickPrice timestamp, _
                tickerid Mod INCARNATION_INCREMENT, _
                tickType, _
                price, _
                size

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.TICK_PRICE
    listendata.price = price
    listendata.size = size
    listendata.tickerid = tickerid
    listendata.tickType = tickType
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgTickPrice, _
                    Me
End If

Exit Sub

err:
err.Raise err.number, , "TWSAPI::getTickPrice(failpoint=" & failpoint & "): " & err.Description
End Sub

Private Sub getTickSize(ByVal version As Long)
Dim listendata As TWSSocketMsgTickSize
Dim timestamp As Date
Dim size As Long
Dim tickType As Long
Dim tickerid As Long
Dim failpoint As Long

On Error GoTo err

failpoint = 100

tickerid = getLong

failpoint = 200

tickType = getLong

failpoint = 300

size = getLong
timestamp = gCurrentTime

failpoint = 400

processTickSize timestamp, _
                tickerid Mod INCARNATION_INCREMENT, _
                tickType, _
                size

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.TICK_PRICE
    listendata.size = size
    listendata.tickerid = tickerid
    listendata.tickType = tickType
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgTickSize, _
                    Me
End If

Exit Sub

err:
err.Raise err.number, , "TWSAPI::getTickSize(failpoint=" & failpoint & "): " & err.Description
End Sub

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError number, source, Description, -1

If mConnectionState = ConnConnected Then disconnect "Fatal error"
End Sub

Private Sub handleTwsDisconnection()
mConnectionState = ConnectionStates.ConnNotConnected
ReleaseSocket
If mDisconnectedByApplication Then
    RaiseEvent ConnectionToTWSClosed(False)
    logMessage "Connection to TWS closed: " & calcConnectionString, LogLevelLow
Else
    If mConnectionRetryIntervalSecs = 0 Then
        cancelContractRequests "Lost connection to TWS"
        RaiseEvent ConnectionToTWSClosed(False)
        logMessage "Connection to TWS closed" & _
                    calcConnectionString, _
                    LogLevelLow
    Else
        retryConnection
        RaiseEvent ConnectionToTWSClosed(True)
        logMessage "Connection to TWS closed - reconnecting: " & _
                    calcConnectionString, _
                    LogLevelLow
    End If
End If
End Sub

Private Sub logMessage( _
                ByVal message As String, _
                Optional ByVal logLevel As LogLevels = LogLevelLow)
If mLogLevel < logLevel Or _
    mLogLevel = LogLevelNone Then Exit Sub
mCommonServiceConsumer.NotifyListeners "IBTWSSP: " & message, _
                                        TradeBuildSP.StandardListenValueTypes.Log, _
                                        Me
End Sub

Private Function optionRightFromString(ByVal value As String) As OptionRights
Select Case UCase$(value)
Case "C"
    optionRightFromString = OptCall
Case "P"
    optionRightFromString = OptPut
Case Else
    optionRightFromString = OptNone
End Select
End Function

Private Function optionRightToString(ByVal value As OptionRights) As String
Select Case value
Case OptCall
    optionRightToString = "CALL"
Case OptPut
    optionRightToString = "PUT"
End Select
End Function

Private Function orderStatusFromString(ByVal value As String) As OrderStatuses
Select Case UCase$(value)
Case "CREATED"
    orderStatusFromString = OrderStatusCreated
Case "REJECTED"
    orderStatusFromString = OrderStatusRejected
Case "PENDINGSUBMIT"
    orderStatusFromString = OrderStatusPendingSubmit
Case "PRESUBMITTED"
    orderStatusFromString = OrderStatusPreSubmitted
Case "SUBMITTED"
    orderStatusFromString = OrderStatusSubmitted
Case "PENDINGCANCEL"
    orderStatusFromString = OrderStatusCancelling
Case "CANCELLED"
    orderStatusFromString = OrderStatusCancelled
Case "FILLED"
    orderStatusFromString = OrderStatusFilled
End Select
End Function

Private Function orderStopTriggerMethodToString(ByVal value As StopTriggerMethods) As String
Select Case value
Case StopTriggerMethods.StopTriggerDefault
    orderStopTriggerMethodToString = "0"
Case StopTriggerMethods.StopTriggerDoubleBidAsk
    orderStopTriggerMethodToString = "1"
Case StopTriggerMethods.StopTriggerDoubleLast
    orderStopTriggerMethodToString = "3"
Case StopTriggerMethods.StopTriggerLast
    orderStopTriggerMethodToString = "2"
End Select
End Function

Private Function orderTIFFromString(ByVal value As String) As OrderTifs
Select Case UCase$(value)
Case "DAY"
    orderTIFFromString = TIFDay
Case "GTC"
    orderTIFFromString = TIFGoodTillCancelled
Case "IOC"
    orderTIFFromString = TIFImmediateOrCancel
End Select
End Function

Private Function orderTIFToString(ByVal value As OrderTifs) As String
Select Case value
Case TIFDay
    orderTIFToString = "DAY"
Case TIFGoodTillCancelled
    orderTIFToString = "GTC"
Case TIFImmediateOrCancel
    orderTIFToString = "IOC"
End Select
End Function

Private Function orderTypeFromString(ByVal value As String) As TradeBuildSP.OrderTypes
Select Case UCase$(value)
Case "MKT"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeMarket
Case "MKTCLS"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeMarketOnClose
Case "LMT"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeLimit
Case "LMTCLS"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeLimitOnClose
Case "PEGMKT"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypePeggedToMarket
Case "STP"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeStop
Case "STPLMT"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeStopLimit
Case "TRAIL"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeTrail
Case "REL"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeRelative
Case "VWAP"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeVWAP
Case "MTL"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeMarketToLimit
Case "RFQ"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeQuote
Case "ADJUST"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeAdjust
Case "ALERT"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeAlert
Case "LIT"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeLimitIfTouched
Case "MIT"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeMarketIfTouched
Case "TRAILLMT"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeTrailLimit
Case "MKTPROT"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeMarketWithProtection
Case "MOO"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeMarketOnOpen
Case "MOC"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeMarketOnClose
Case "LOO"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeLimitOnOpen
Case "LOC"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeLimitOnClose
Case "PEGPRI"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypePeggedToPrimary
Case Else
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeNone
End Select
End Function

Private Function orderTypeToString(ByVal value As OrderTypes) As String
Select Case value
Case TradeBuildSP.OrderTypes.OrderTypeMarket
    orderTypeToString = "MKT"
Case TradeBuildSP.OrderTypes.OrderTypeMarketOnClose
    orderTypeToString = "MKTCLS"
Case TradeBuildSP.OrderTypes.OrderTypeLimit
    orderTypeToString = "LMT"
Case TradeBuildSP.OrderTypes.OrderTypeLimitOnClose
    orderTypeToString = "LMTCLS"
Case TradeBuildSP.OrderTypes.OrderTypePeggedToMarket
    orderTypeToString = "PEGMKT"
Case TradeBuildSP.OrderTypes.OrderTypeStop
    orderTypeToString = "STP"
Case TradeBuildSP.OrderTypes.OrderTypeStopLimit
    orderTypeToString = "STPLMT"
Case TradeBuildSP.OrderTypes.OrderTypeTrail
    orderTypeToString = "TRAIL"
Case TradeBuildSP.OrderTypes.OrderTypeRelative
    orderTypeToString = "REL"
Case TradeBuildSP.OrderTypes.OrderTypeVWAP
    orderTypeToString = "VWAP"
Case TradeBuildSP.OrderTypes.OrderTypeMarketToLimit
    orderTypeToString = "MTL"
Case TradeBuildSP.OrderTypes.OrderTypeQuote
    orderTypeToString = "QUOTE"
Case TradeBuildSP.OrderTypes.OrderTypeAdjust
    orderTypeToString = "ADJUST"
Case TradeBuildSP.OrderTypes.OrderTypeAlert
    orderTypeToString = "ALERT"
Case TradeBuildSP.OrderTypes.OrderTypeLimitIfTouched
    orderTypeToString = "LIT"
Case TradeBuildSP.OrderTypes.OrderTypeMarketIfTouched
    orderTypeToString = "MIT"
Case TradeBuildSP.OrderTypes.OrderTypeTrailLimit
    orderTypeToString = "TRAILLMT"
Case TradeBuildSP.OrderTypes.OrderTypeMarketWithProtection
    orderTypeToString = "MKTPROT"
Case TradeBuildSP.OrderTypes.OrderTypeMarketOnOpen
    orderTypeToString = "MOO"
Case TradeBuildSP.OrderTypes.OrderTypeLimitOnOpen
    orderTypeToString = "LOO"
Case TradeBuildSP.OrderTypes.OrderTypePeggedToPrimary
    orderTypeToString = "PEGPRI"
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException
End Select
End Function

Private Sub ProcessAsk(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal price As Double, _
                            ByVal size As Long)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    .dataConsumer.Ask timestamp, _
                price, _
                size
End With
End Sub

Private Sub ProcessBid(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal price As Double, _
                            ByVal size As Long)
With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    .dataConsumer.Bid timestamp, _
                price, _
                size
End With
End Sub

Private Sub ProcessClose(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal price As Double)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    .dataConsumer.PreviousClose timestamp, price
End With
End Sub

Private Sub processContract(ByVal pContract As contract)

If Not mFirstContractDetailsReceived Then
    mFirstContractDetailsReceived = True
    
    ' now send a reqContract that will give a 'no such security
    ' exists' error. This error will act as a delimiter for all
    ' the contract details returned by the orginal request. Note that
    ' tests indicate that contract details for different reqContracts
    ' are not interleaved
    
    Const version = 1
    addData REQ_CONTRACT_DATA
    addData version
    addData "ES"
    addData "FUT"
    addData "200501"
    addData "0"
    addData ""
    addData "LIFFE"
    addData "GBP"
    addData ""
    send
End If

mContractDetailsRequest.localContracts.add pContract

End Sub

Private Sub processContractError(ByVal reason As String)
Dim failpoint As Long

On Error GoTo err

If mContractDetailsRequest.localContracts.Count > 0 Then
    ' we are processing the forced invalid contract details request
    ' for a request for which we have received one or more contract details
    failpoint = 100
    mContractDetailsRequest.requester.contractsLoaded mContractDetailsRequest
Else
    ' this is a genuine error response to a contract details request
    failpoint = 200
    mContractDetailsRequest.requester.contractSpecifierInvalid mContractDetailsRequest, reason
End If
Set mContractDetailsRequest.localContracts = Nothing
Set mContractDetailsRequest.tradeBuildContracts = Nothing
Set mContractDetailsRequest.requester = Nothing
mFirstContractDetailsReceived = False

failpoint = 300
submitNextContractDetailsRequest

Exit Sub

err:
err.Raise err.number, , "TWSAPI::processContractError(failpoint=" & failpoint & "): " & err.Description
End Sub

Private Sub ProcessHigh(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal price As Double)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    .dataConsumer.High timestamp, price
    
End With
End Sub

Private Sub ProcessLast(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal price As Double, _
                            ByVal size As Long)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    .dataConsumer.Trade timestamp, _
            price, _
            size
End With
End Sub

Private Sub ProcessLow(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal price As Double)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    .dataConsumer.Low timestamp, price

End With
End Sub

Private Sub processMarketDepth(ByVal timestamp As Date, _
                                ByVal id As Long, _
                                ByVal position As Long, _
                                ByVal operation As DOMOperations, _
                                ByVal side As DOMSides, _
                                ByVal price As Double, _
                                ByVal size As Long)

With mTickerTable(id)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    .dataConsumer.UpdateMktDepth timestamp, _
                position, _
                operation, _
                side, _
                price, _
                size
    
End With
End Sub

Private Sub processNextHistDataRequest()
If mTwsHistDataRequestQueue.size <> 0 Then
    Dim request As TWSHistoricalDataRequestQueueEntry
    Dim requestKey As String
    request = mTwsHistDataRequestQueue.removeFromFront
    reqHistoricalData request
End If
End Sub

Private Sub processTickPrice(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal tickType As TWSSocketTickTypes, _
                            ByVal price As Double, _
                            ByVal size As Long)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    If price < (0.5 * .prevTrade) Then
        Exit Sub        ' ignore invalid prices
    End If
    
    Select Case tickType
    Case TICK_BID
        size = IIf(size <> 0, size, .prevBidSize)
        ProcessBid timestamp, _
                    tickerid, _
                    price, _
                    size
        .prevBid = price
        If size <> 0 Then .prevBidSize = size
    
    Case TICK_ASK
        size = IIf(size <> 0, size, .prevAskSize)
        ProcessAsk timestamp, _
                    tickerid, _
                    price, _
                    size
        
        .prevAsk = price
        If size <> 0 Then .prevAskSize = size
    
    Case TICK_LAST
        size = IIf(size <> 0, size, .prevTradeSize)
        ProcessLast timestamp, _
                    tickerid, _
                    price, _
                    size
        .prevTrade = price
        If size <> 0 Then .prevTradeSize = size
    
    Case TICK_HIGH
        ProcessHigh timestamp, _
                    tickerid, _
                    price
    Case TICK_LOW
        ProcessLow timestamp, _
                    tickerid, _
                    price
    Case TICK_CLOSE
        ProcessClose timestamp, _
                    tickerid, _
                    price
    End Select
End With

End Sub

Private Sub processTickSize(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal tickType As TWSSocketTickTypes, _
                            ByVal size As Long)

With mTickerTable(tickerid)
    
    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    Select Case tickType
    Case TICK_BID_SIZE
        If size = .prevBidSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        .prevBidSize = size
        ProcessBid timestamp, _
                    tickerid, _
                    .prevBid, _
                    size
    Case TICK_ASK_SIZE
        If size = .prevAskSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        ProcessAsk timestamp, _
                    tickerid, _
                    .prevAsk, _
                    size
    Case TICK_LAST_SIZE
        If size = .prevTradeSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        .prevTradeSize = size
        ProcessLast timestamp, _
                    tickerid, _
                    .prevTrade, _
                    size
    Case TICK_VOLUME
        processVolume timestamp, _
                    tickerid, _
                    size
        
    End Select
End With

End Sub

Private Sub processVolume(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal size As Long)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    If size = 0 Then Exit Sub
    
    .dataConsumer.volume timestamp, _
                size
End With
End Sub

Private Function queueTWSHistoricalDataRequest(ByRef TwsRequest As TWSHistoricalDataRequestQueueEntry) As String
queueTWSHistoricalDataRequest = mTwsHistDataRequestQueue.addToEnd(TwsRequest)

If mConcurrentHistDataRequests < MaxConcurrentHistDataRequests Then
    mTwsHistDataRequestQueue.removeFromFront
    reqHistoricalData TwsRequest
End If
End Function

Private Sub reEstablishMarketData(ByVal newConnection As Boolean)
Dim i As Long
For i = 0 To UBound(mTickerTable)
    With mTickerTable(i)
        If Not .contractSpecifier Is Nothing Then
            If Not newConnection Then cancelMarketDataEx i, False
            .incarnation = .incarnation + INCARNATION_INCREMENT
            reqMarketData i
        End If
    End With
Next
End Sub

Private Sub releaseAllTickerTableEntries()
Dim i As Long
For i = 0 To UBound(mTickerTable)
    releaseTickerTableEntry i
Next
End Sub

Private Sub releaseHistoricalDataTableEntry( _
                ByVal index As Long)
With mHistoricalDataTable(index)
    Set .bars = Nothing
    Erase .requestKeys
    Set .historicalDataReader = Nothing
    .numberOfTwsRequests = 0
    .nextBarIndex = -1
    .inUse = False
    .cancelling = False
    .dataRetrievedFromTWS = False
    .barLength = 0
    .barType = 0
    .earliestReuseTime = 0
End With
End Sub

Private Sub ReleaseSocket()
If Not mSocket Is Nothing Then
    mSocket.Close
    Set mSocket = Nothing
End If
End Sub

Private Sub releaseTickerTableEntry(ByVal tickerid As Long)
With mTickerTable(tickerid)
    .inUse = False
    .incarnation = .incarnation + INCARNATION_INCREMENT
    Set .contractSpecifier = Nothing
    .prevBid = 0#
    .prevBidSize = 0&
    .prevAsk = 0#
    .prevAskSize = 0&
    .prevTrade = 0#
    .prevTradeSize = 0&
    .receivingMarketDepth = False
    Set .dataConsumer = Nothing
    Set .dataReader = Nothing
End With
End Sub

Private Sub reqContract()
Dim contractSpecifier As TradeBuildSP.IContractSpecifier
Set contractSpecifier = mContractDetailsRequest.tradeBuildContracts.contractSpecifier
contractSpecifier.Locked = True

If contractSpecifier.exchange <> "" And Not IsValidExchangeCode(contractSpecifier.exchange) Then
    processContractError "invalid exchange"
    Exit Sub
End If

Const version = 1

addData REQ_CONTRACT_DATA
addData version
With contractSpecifier
    addData .symbol
    addData secTypeToString(.sectype)
    addData Left$(.expiry, 6)
    addData .strike
    addData optionRightToString(.right)
    addData .exchange
    addData .currencyCode
    addData .localSymbol
End With
send
End Sub

Private Sub reqHistoricalData( _
                request As TWSHistoricalDataRequestQueueEntry)

Const version = 3

If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

addData REQ_HISTORICAL_DATA
addData version
addData request.id
addData request.contractSpec.symbol
addData secTypeToString(request.contractSpec.sectype)
addData request.contractSpec.expiry
addData request.contractSpec.strike
addData optRightToString(request.contractSpec.right)
addData ""      ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>> multiplier ??? why?
addData request.contractSpec.exchange
addData ""      ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>> primaryExch ??? not in ActiveX interface
addData request.contractSpec.currencyCode
addData request.contractSpec.localSymbol
If mServerVersion >= 20 Then
    addData request.endDateTime
    
    addData request.barSizeSetting
End If

' durationStr

addData request.duration

' useRTH
addData 0

addData request.whatToShow

' formatDate
If mServerVersion > 16 Then
    addData TWSHistDataDateFormats.DateFormatString
End If

If request.contractSpec.ComboLegs Is Nothing Then
    addData 0
Else
    addData request.contractSpec.ComboLegs.Count
    
    Dim comboLeg As TradeBuildSP.IComboLeg
    For Each comboLeg In request.contractSpec.ComboLegs
        addData comboLeg.contractID
        addData comboLeg.Ratio
        addData orderActionToString(comboLeg.action)
        addData comboLeg.exchange
    Next
End If

send

mConcurrentHistDataRequests = mConcurrentHistDataRequests + 1
      
End Sub

Private Sub reqMarketData(ByVal tickerTableIndex As Long)

Const version = 3

addData REQ_MKT_DATA
addData version
addData tickerTableIndex + mTickerTable(tickerTableIndex).incarnation
With mTickerTable(tickerTableIndex).contractSpecifier
    addData .symbol
    addData secTypeToString(.sectype)
    addData Left$(.expiry, 6)
    addData .strike
    addData optionRightToString(.right)
    addData .exchange
    addData .currencyCode
    If mServerVersion >= 2 Then addData .localSymbol
    ' Add combo legs for BAG requests
    If mServerVersion >= 8 And .sectype = SecurityTypes.SecTypeBag Then
        addData .ComboLegs.Count
        Dim comboLeg As TradeBuildSP.IComboLeg
        For Each comboLeg In .ComboLegs
            With comboLeg
                addData .contractID
                addData .Ratio
                addData orderActionToString(.action)
                addData .exchange
            End With
        Next
    End If
    
    send
End With

End Sub

Private Sub reqMarketDepth(ByVal tickerid As Long, _
                        ByVal contractSpec As TradeBuildSP.IContractSpecifier)

Const version = 1

addData REQ_MKT_DEPTH
addData version
addData tickerid + MARKET_DEPTH_TICKER_BASE
With contractSpec
    addData .symbol
    addData secTypeToString(.sectype)
    addData Left$(.expiry, 6)
    addData .strike
    addData optionRightToString(.right)
    addData .exchange
    addData .currencyCode
    addData .localSymbol
End With
send
End Sub

Private Sub resetMarketDepth(ByVal reEstablish As Boolean)
Dim i As Long
For i = 0 To UBound(mTickerTable)
    With mTickerTable(i)
        If .receivingMarketDepth Then
            If reEstablish Then reqMarketDepth i, .contractSpecifier
            .dataConsumer.resetMarketDepth gCurrentTime, reEstablish
        End If
    End With
Next
End Sub

Private Sub retryConnection()
If mConnectionRetryIntervalSecs <> 0 Then
    If mConnectionTimer Is Nothing Then
        Set mConnectionTimer = New IntervalTimer
        mConnectionTimer.RepeatNotifications = False
        mConnectionTimer.TimerIntervalSecs = mConnectionRetryIntervalSecs
    End If
    mConnectionTimer.StartTimer
End If
End Sub

Private Sub rollbackBuffer()
'mIndexIn = mIndexInChk
mInputIndex = mInputIndexChk
End Sub

Private Sub send()
ReDim Preserve mBufferOut(mIndexOut - 1) As Byte
mSocket.SendData mBufferOut
mIndexOut = 0
End Sub

Private Sub sendOrder( _
                ByVal pOrder As TradeBuildSP.IOrder, _
                ByVal transmit As Boolean, _
                ByVal parentId As String, _
                ByVal ocaGroup As String, _
                ByVal routingHint As String)

Dim specifier As TradeBuildSP.IContractSpecifier

Set specifier = pOrder.contract.specifier

Const version = 18
addData PLACE_ORDER
addData version
addData pOrder.brokerId

' send contract fields
With specifier
    addData .symbol
    addData secTypeToString(.sectype)
    addData Left$(.expiry, 6)
    addData .strike
    addData optionRightToString(.right)
    If mServerVersion >= 15 Then addData ""         ' multiplier - this is the number of units of the underlying
                                                  ' delivered per contract - see http://www.interactivebrokers.com/discus/messages/2/27413.html
    addData .exchange
    If mServerVersion >= 14 Then addData routingHint     ' aka primaryExchange in TWS - only relevant when
                                                        ' exchange is SMART and there are SMART routers in more
                                                        ' than one country (eg try IBM in TWS)
    addData .currencyCode
    If mServerVersion >= 2 Then addData .localSymbol
End With

With pOrder
    ' send main order fields
    addData orderActionToString(.action)
    addData .quantity
    addData orderTypeToString(.orderType)
    If mServerVersion <= 10 And _
        (UCase$(specifier.exchange) = "LIFFE" Or _
        UCase$(specifier.exchange) = "LSE") _
    Then
        addData .limitPrice / 100
        addData .triggerPrice / 100
    Else
        addData .limitPrice
        addData .triggerPrice
    End If
    ' send extended order fields
    addData orderTIFToString(.timeInForce)
    addData ocaGroup              ' ocaGroup
    addData ""              ' account
    addData ""              ' openClose
    addData .origin     ' origin
    addData .originatorRef
    addBoolean transmit
    If mServerVersion >= 4 Then addData parentId
    If mServerVersion >= 5 Then
        addBoolean .blockOrder
        addBoolean .sweepToFill
        addData .displaySize
        addData orderStopTriggerMethodToString(.stopTriggerMethod)
        addBoolean .ignoreRegularTradingHours
    End If
    If mServerVersion >= 7 Then addBoolean .hidden
End With

' Send combo legs for BAG requests
If mServerVersion >= 8 And specifier.sectype = SecurityTypes.SecTypeBag Then
    With specifier
        If .ComboLegs.Count = 0 Then
            addData 0
        Else
            addData .ComboLegs.Count
            Dim comboLeg As TradeBuildSP.IComboLeg
            For Each comboLeg In .ComboLegs
                With comboLeg
                    addData .contractID
                    addData .Ratio
                    addData orderActionToString(.action)
                    addData .exchange
                    addData LegOpenCloseToString(.openClose)
                End With
            Next
        End If
    End With
End If

With pOrder
    If mServerVersion >= 9 Then addData ""          ' sharesAllocation (deprecated)
    
    If mServerVersion >= 10 Then addData .discretionaryAmount
    
    If mServerVersion >= 11 Then
        If .goodAfterTime = 0 Then
            addData ""
        Else
            addData Format(.goodAfterTime, "yyyymmdd hh:nn:ss ") & .goodAfterTimeTZ
        End If
    End If
    
    If mServerVersion >= 12 Then
        If .goodTillDate = 0 Then
            addData ""
        Else
            addData Format(.goodTillDate, "yyyymmdd hh:nn:ss ") & .goodTillDateTZ
        End If
    End If
    
    If mServerVersion >= 13 Then
        addData ""      ' FAGroup
        addData ""      ' FAMethod
        addData ""      ' FAPercentage
        addData ""      ' FAProfile
    End If
    
    If mServerVersion >= 18 Then 'institutional short sale slot fields.
        addData 0       ' shortSaleSlot - 0 only for retail, 1 or 2 only for institution.
        addData ""      ' designatedLocation - only populate when shortSaleSlot = 2.
    End If
    
    If mServerVersion >= 19 Then
        addData 0           ' ocaType
        addBoolean .regularTradingHoursOnly
        addBoolean False    ' rule80A
        addData .settlingFirm
        addBoolean .allOrNone
        addData .minimumQuantity
        addData ""          ' percentOffset
        addBoolean .eTradeOnly
        addBoolean .firmQuoteOnly
        addData .NbboPriceCap
        addData ""          ' auctionStrategy
        addData ""          ' startingPrice
        addData ""          ' stockRefPrice
        addData ""          ' delta
        addData ""          ' stockRangeLower
        addData ""          ' stockRangeUpper
    End If
    
    If mServerVersion >= 22 Then addBoolean .overrideConstraints

    If mServerVersion >= 26 Then ' Volatility orders
        addData ""          ' volatility
        addData ""          ' volatilityType);
        If mServerVersion < 28 Then
            addData "0"     ' send( order.m_deltaNeutralOrderType.equalsIgnoreCase("MKT"));
        Else
            addData ""          ' deltaNeutralOrderType);
            addData ""          ' deltaNeutralAuxPrice
        End If
        
        addData 0          ' continuousUpdate
        
        If mServerVersion = 26 Then
            ' Volatility orders had specific watermark price attribs in server version 26
            addData ""          ' stockRangeLower
            addData ""          ' stockRangeUpper
        End If
        
        addData ""          ' referencePriceType
    End If
    
End With

send
End Sub
                    
Private Sub setTWSLogLevel(logLevel As TWSLogLevels)

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::SetServerLogLevel", _
                "Not connected"
End If

Const version = 1

addData SET_SERVER_LOGLEVEL
addData version
addData logLevel
send
End Sub

Private Function showBuffer(buffer() As Byte) As String
Dim i As Long
Dim char As Byte

showBuffer = String(UBound(buffer) + 1, "_")

For i = 0 To UBound(buffer)
    char = buffer(i)
    If char <> 0 Then Mid$(showBuffer, i + 1, 1) = Chr$(char)
Next
End Function

Private Sub submitNextContractDetailsRequest()
Dim failpoint As Long

On Error GoTo err

If Not mConnectionState = ConnConnected Then Exit Sub
If Not mContractDetailsRequest.tradeBuildContracts Is Nothing Then Exit Sub
If mTwsContractDetailsRequestQueue.size = 0 Then Exit Sub
    
failpoint = 100
On Error Resume Next
mContractDetailsRequest = mTwsContractDetailsRequestQueue.removeFromFront
If err.number <> 0 Then Stop
On Error GoTo err

failpoint = 200
Set mContractDetailsRequest.localContracts = New Collection
reqContract

Exit Sub

err:
err.Raise err.number, , "TWSAPI::submitNextContractDetailsRequest(failpoint=" & failpoint & "): " & err.Description
End Sub


