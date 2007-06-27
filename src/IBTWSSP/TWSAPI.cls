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
                ByVal pContractSpecifier As ContractSpecifier, _
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

Private Const ClientVersion As Long = 23

Private Const ErrHistGenerationFailed = vbObjectError

Private Const HistoricalDataTickerBase As Long = &H410000
Private Const HistoricalDataTickerIncrement As Long = &H100

Private Const IncarnationIncrement As Long = &H1000&

Private Const InitialMaxTickers As Long = 100&

Private Const MarketDepthTickerBase As Long = &H400000

Private Const MaxConcurrentHistDataRequests As Long = 50    ' Limit imposed by IB

Private Const NumDaysInWeek As Long = 5
Private Const NumDaysInMonth  As Long = 22
Private Const NumDaysInYear As Long = 260
Private Const NumMonthsInYear As Long = 12

Private Const OrderIdBase As Long = &H800000

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
    ignoreTimestamps        As Boolean
    barSpecifier            As TradeBuildSP.IBarDataSpecifier
    SessionBuilder          As SessionBuilder
    Session                 As Session
    earliestReuseTime       As Date
    numberOfTwsRequests     As Long         ' the number of historical data requests
                                            ' made to TWS
    barType                 As TradeBuildSP.BarTypes
    barLength               As Double       ' days
    historicalDataReader    As HistDataReader
    nextBarIndex            As Long
    bars                    As LinkedList
    requestKeys()           As String
End Type

Private Type TickerTableEntry
    incarnation             As Long
    Contract                As Contract
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
Private mWinsockForm As WinsockForm

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

mIndexIn = 0
mIndexOut = 0
mInputIndex = 0

logMessage "Connecting to TWS: negotiating versions and client id"
addData ClientVersion
send
addData mClientID
send
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mSocket_Connect", _
                err.Description
End Sub

Private Sub mSocket_DataArrival( _
                ByVal bytesTotal As Long)
Dim bytesReceived As Long
On Error GoTo err

mSocket.GetData mSocketInputData, vbArray + vbByte
bytesReceived = UBound(mSocketInputData) + 1
'Debug.Print "Received " & bytesReceived & " from socket: " & showBuffer(mSocketInputData, bytesReceived)

If mLogLevel >= LogLevelHigh Then
    mCommonServiceConsumer.NotifyListeners showBuffer(mSocketInputData, bytesReceived), _
                                        ListenValueTypes.ValueTypeRawSocketData, _
                                        Me
End If

If mIndexIn = mInputIndex Then
    ' no unprocessed data so start at the beginning of the buffer again
    'Debug.Print "Reverting to start of buffer (no unprocessed data)"
    mIndexIn = 0
    mInputIndex = 0
    mInputIndexChk = 0
End If

'if spaceUsed + newSpaceRequired > lengthOfBuffer then
If mIndexIn + bytesReceived > UBound(mBufferIn) + 1 Then
    'if newSpaceRequired <= lengthOfBuffer - unprocessedLength then
    If bytesReceived <= (UBound(mBufferIn) + 1 - mIndexIn + mInputIndex) Then
        ' copy the existing contents down to the start of the buffer
        'Debug.Print "Moving unprocessed data (" & mIndexIn - mInputIndex & " bytes) from " & mInputIndex & " to start of buffer"
        MoveMemory mBufferIn(0), mBufferIn(mInputIndex), mIndexIn - mInputIndex
        mIndexIn = mIndexIn - mInputIndex
        mInputIndex = 0
        mInputIndexChk = 0
    Else
        'Debug.Print "Extending buffer by " & IIf(bytesReceived > 1024, bytesReceived, 1024) & " bytes"
        ReDim Preserve mBufferIn(UBound(mBufferIn) + IIf(bytesReceived > 1024, bytesReceived, 1024)) As Byte
    End If
End If
'Debug.Print "Appending socket message to buffer at " & mIndexIn
CopyMemory mBufferIn(mIndexIn), mSocketInputData(0), bytesReceived
mIndexIn = mIndexIn + bytesReceived

If mIndexIn <> mInputIndex Then
    checkmessages
End If

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
                        connectionString, _
                        LogLevelLow
            
            cancelContractRequests "Failed to connect to TWS"
            RaiseEvent ConnectFailed(Description, False)
        Else
            logMessage "Failed to connect to TWS - retrying: " & _
                        Description & ": " & _
                        connectionString, _
                        LogLevelLow
            
            RaiseEvent ConnectFailed(Description, True)
            retryConnection
        End If
    Else
        logMessage "Socket error " & number & ": " & _
                    Description & ": " & _
                    connectionString, _
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
                connectionString, _
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

Friend Property Get connectionString() As String
connectionString = "server=" & mServer & _
                        "; port=" & mPort & _
                        "; client Id=" & mClientID
End Property

Friend Property Get InBuffer() As String
InBuffer = showBuffer(mBufferIn, mIndexIn - 1)
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
addData MarketDepthTickerBase + tickerid
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


logMessage "Connecting to TWS: " & connectionString
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
RaiseEvent Disconnected(reason)
logMessage "Disconnected from: " & connectionString, LogLevelLow
gReleaseTWSAPIInstance Me, True
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
                        ByRef pBar As TradeBuildSP.Bar, _
                        ByRef barLength As Long, _
                        ByRef barUnit As TimePeriodUnits) As Boolean

Dim Bar As HistoricBar

On Error GoTo err
With mHistoricalDataTable(requestHandle)
    If .nextBarIndex = -1 Then
        releaseHistoricalDataTableEntry requestHandle
    Else
        ReadHistoricalBar = True
        Bar = .bars.itemAt(.nextBarIndex)
        barLength = .barSpecifier.barLength
        barUnit = .barSpecifier.BarLengthUnits
        With Bar
            pBar.timestamp = .timestamp
            pBar.OpenPrice = .openValue
            pBar.HighPrice = .highValue
            pBar.LowPrice = .lowValue
            pBar.ClosePrice = .closeValue
            pBar.volume = .volume
            pBar.TickVolume = 0
            pBar.OpenInterest = 0
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
                ByVal contractSpec As ContractSpecifier, _
                ByVal pContractsBuilder As ContractsBuilder, _
                ByVal requester As ContractDetailsRequester)

Dim requestQueueEntry As TWSContractDetailsRequestQueueEntry

On Error GoTo err
Set requestQueueEntry.contractSpec = contractSpec
Set requestQueueEntry.tradeBuildContractsBuilder = pContractsBuilder
Set requestQueueEntry.requester = requester

Select Case contractSpec.secType
Case SecurityTypes.SecTypeCombo
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

If contractSpec.localSymbol = "" And _
    contractSpec.symbol = "" _
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
        addData IIf(CLng(.Time) <> 0, Format(.Time, "yyyymmdd-hh\:nn\:ss"), "")
        addData .symbol
        addData secTypeToString(.secType)
        addData .exchange
        addData orderActionToString(.Action)
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

Friend Function RequestMarketData(ByVal pContract As Contract, _
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

lTickerID = AllocateTickerTableEntry(dataConsumer)
RequestMarketData = lTickerID
Set mTickerTable(lTickerID).Contract = pContract
Set mTickerTable(lTickerID).dataReader = dataReader
reqMarketData lTickerID

If includeMarketDepth Then
    mTickerTable(lTickerID).receivingMarketDepth = True
    reqMarketDepth lTickerID, pContract.specifier
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
    If .Contract Is Nothing Then err.Raise ErrorCodes.ErrIllegalArgumentException, "IBTWSSP.TWSAPI::requestMarketDepth", "Invalid ticker id: " & pTickerID
    .receivingMarketDepth = True
    reqMarketDepth pTickerID, .Contract.specifier
End With
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestMarketDepth", _
                err.Description

End Sub

Friend Function RequestHistoricalData( _
                barSpecifier As TradeBuildSP.IBarDataSpecifier, _
                ByVal historicalDataReader As HistDataReader) As Long

Dim requestID As Long
Dim i As Long
Dim ignoreTimestamps As Boolean

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

' now adjust the bar specifier to indicate the number of bars of
' appropriate length supported by IB
Select Case barSpecifier.BarLengthUnits
Case TimePeriodUnits.TimePeriodDay
    ignoreTimestamps = True
    If barSpecifier.barLength > 1 Then
        barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * barSpecifier.barLength
        barSpecifier.barLength = 1
    End If
Case TimePeriodUnits.TimePeriodHour
    If barSpecifier.barLength > 1 Then
        barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * barSpecifier.barLength
        barSpecifier.barLength = 1
    End If
Case TimePeriodUnits.TimePeriodMinute
    If barSpecifier.barLength Mod 30 = 0 Then
        barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * (barSpecifier.barLength / 30)
        barSpecifier.barLength = 30
    ElseIf barSpecifier.barLength Mod 15 = 0 Then
        barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * (barSpecifier.barLength / 15)
        barSpecifier.barLength = 15
    ElseIf barSpecifier.barLength Mod 5 = 0 Then
        barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * (barSpecifier.barLength / 5)
        barSpecifier.barLength = 5
    ElseIf barSpecifier.barLength Mod 2 = 0 Then
        barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * (barSpecifier.barLength / 2)
        barSpecifier.barLength = 2
    ElseIf barSpecifier.barLength > 1 Then
        barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * barSpecifier.barLength
        barSpecifier.barLength = 1
    End If
Case TimePeriodUnits.TimePeriodMonth
    ignoreTimestamps = True
    barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * barSpecifier.barLength
    barSpecifier.barLength = 1
Case TimePeriodUnits.TimePeriodSecond
    If barSpecifier.barLength Mod 30 = 0 Then
        barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * (barSpecifier.barLength / 30)
        barSpecifier.barLength = 30
    ElseIf barSpecifier.barLength Mod 15 = 0 Then
        barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * (barSpecifier.barLength / 15)
        barSpecifier.barLength = 15
    ElseIf barSpecifier.barLength Mod 5 = 0 Then
        barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * (barSpecifier.barLength / 5)
        barSpecifier.barLength = 5
    ElseIf barSpecifier.barLength > 1 Then
        barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * barSpecifier.barLength
        barSpecifier.barLength = 1
    End If
Case TimePeriodUnits.TimePeriodWeek
    ignoreTimestamps = True
    barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * barSpecifier.barLength
    barSpecifier.barLength = 1
Case TimePeriodUnits.TimePeriodYear
    ignoreTimestamps = True
    barSpecifier.NumberOfBars = barSpecifier.NumberOfBars * barSpecifier.barLength * NumMonthsInYear
    barSpecifier.barLength = 1
    barSpecifier.BarLengthUnits = TimePeriodMonth
Case TimePeriodUnits.TimePeriodVolume, TimePeriodUnits.TimePeriodTickMovement
    ' nothing to do
Case Else
    historicalDataReader.historicalDataRequestInvalid barSpecifier.NumberOfBars & "specified timeframe not supported"
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

mHistoricalDataTable(requestID).inUse = True
mHistoricalDataTable(requestID).cancelling = False
Set mHistoricalDataTable(requestID).barSpecifier = barSpecifier
mHistoricalDataTable(requestID).ignoreTimestamps = ignoreTimestamps
Set mHistoricalDataTable(requestID).SessionBuilder = New SessionBuilder
mHistoricalDataTable(requestID).SessionBuilder.sessionStartTime = barSpecifier.Contract.sessionStartTime
mHistoricalDataTable(requestID).SessionBuilder.sessionEndTime = barSpecifier.Contract.sessionEndTime
Set mHistoricalDataTable(requestID).Session = mHistoricalDataTable(requestID).SessionBuilder.Session
mHistoricalDataTable(requestID).nextBarIndex = 0
mHistoricalDataTable(requestID).barType = barSpecifier.barType
'mHistoricalDataTable(requestid).barLength = barLength( _
'                                            barSpecifier.barLength, _
'                                            barSpecifier.BarLengthUnits)
mHistoricalDataTable(requestID).nextBarIndex = -1
Set mHistoricalDataTable(requestID).historicalDataReader = historicalDataReader
Set mHistoricalDataTable(requestID).bars = New LinkedList
mHistoricalDataTable(requestID).bars.initialCapacity = IIf(barSpecifier.NumberOfBars <> 0, barSpecifier.NumberOfBars, 50)

generateTwsHistoricalDataRequest requestID
RequestHistoricalData = requestID
Exit Function
err:

If err.number = ErrHistGenerationFailed Then
    releaseHistoricalDataTableEntry requestID
    RequestHistoricalData = -1
    Exit Function
End If
    
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

Private Sub addDouble(ByVal data As String)
addData DoubleToString(data)
End Sub

Private Sub addExchangeCode(ByVal code As String)
mMaxExchangeCodesIndex = mMaxExchangeCodesIndex + 1
If mMaxExchangeCodesIndex > UBound(mExchangeCodes) Then
    ReDim Preserve mExchangeCodes(UBound(mExchangeCodes) + 10) As String
End If
mExchangeCodes(mMaxExchangeCodesIndex) = UCase$(code)
End Sub

Private Function adjustPrice( _
                ByVal Price As Double, _
                ByVal pContractSpecifier As ContractSpecifier) As Double
If pContractSpecifier.currencyCode = "GBP" And _
    mServerVersion < 23 _
Then
    adjustPrice = Price * 100
Else
    adjustPrice = Price
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
Static mElapsedTimer As ElapsedTimer

If mElapsedTimer Is Nothing Then
    Set mElapsedTimer = New ElapsedTimer
    mElapsedTimer.StartTiming
End If

On Error GoTo err

If mConnectionState <> ConnectionStates.ConnConnected Then
    ' first data received is plain server version number
    If checkPointNumber = 0 Then
        mServerVersion = getLong
        checkpointBuffer
        checkPointNumber = 100
        'Debug.Print "Set checkpoint number=" & 100
    End If
    
    If checkPointNumber = 100 Then
        If mServerVersion >= 20 Then getString  ' TWS time
        checkpointBuffer
        checkPointNumber = 200
        'Debug.Print "Set checkpoint number=" & 200
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
    'Debug.Print "Set checkpoint number=" & 1000
End If

Do While mInputIndex < mIndexIn
    
    If checkPointNumber = 1000 Then
        msgID = getLong
        checkpointBuffer
        checkPointNumber = 1100
        'Debug.Print "Set checkpoint number=" & 1100 & ": msgID=" & msgID
    End If
    If checkPointNumber = 1100 Then
        version = getLong
        checkpointBuffer
        checkPointNumber = 1200
        'Debug.Print "Set checkpoint number=" & 1200 & ": version=" & version
    End If
    
    Select Case msgID
    Case TWSSocketInMsgTypes.TICK_PRICE
        'Debug.Print "getTickPrice"
        getTickPrice version
    Case TWSSocketInMsgTypes.TICK_SIZE
        'Debug.Print "getTickSize "
        getTickSize version
    Case TWSSocketInMsgTypes.ORDER_STATUS
        'Debug.Print "getOrderStatus"
        getOrderStatus version
    Case TWSSocketInMsgTypes.ERR_MSG
        'Debug.Print "getErrorMsg "
        getErrorMsg version
    Case TWSSocketInMsgTypes.OPEN_ORDER
        'Debug.Print "getOpenOrder "
        getOpenOrder version
    Case TWSSocketInMsgTypes.ACCT_VALUE
        'Debug.Print "getAccountValue "
        getAccountValue version
    Case TWSSocketInMsgTypes.PORTFOLIO_VALUE
        'Debug.Print "getPortfolioValue "
        getPortfolioValue version
    Case TWSSocketInMsgTypes.ACCT_UPDATE_TIME
        'Debug.Print "getAccountTime "
        getAccountTime version
    Case TWSSocketInMsgTypes.NEXT_VALID_ID
        'Debug.Print "getNextValidId"
        getNextValidId version
    Case TWSSocketInMsgTypes.CONTRACT_DATA
        'Debug.Print "getContractData"
        getContractData version
    Case TWSSocketInMsgTypes.EXECUTION_DATA
        'Debug.Print "getExecutionData "
        getExecutionData version
    Case TWSSocketInMsgTypes.MARKET_DEPTH
        'Debug.Print "getMarketDepth"
        getMarketDepth version
    Case TWSSocketInMsgTypes.MARKET_DEPTH_L2
        'Debug.Print "getMarketDepthL2 "
        getMarketDepthL2 version
    Case TWSSocketInMsgTypes.NEWS_BULLETINS
        'Debug.Print "getNewsBulletins "
        getNewsBulletins version
    Case TWSSocketInMsgTypes.MANAGED_ACCTS
        'Debug.Print "getManagedAccounts "
        getManagedAccounts version
    Case TWSSocketInMsgTypes.RECEIVE_FA
        'Debug.Print "getFAData "
        getFAData version
    Case TWSSocketInMsgTypes.HISTORICAL_DATA
        'Debug.Print "getHistoricalData "
        getHistoricalData version
    Case TWSSocketInMsgTypes.BOND_CONTRACT_DATA
        'Debug.Print "getBondContractData "
        getBondContractData version
    Case TWSSocketInMsgTypes.SCANNER_PARAMETERS
        'Debug.Print "getScannerParameters "
        getScannerParameters version
    Case TWSSocketInMsgTypes.SCANNER_DATA
        'Debug.Print "getScannerData "
        getScannerData version
    Case Else
        Dim s As String
        Dim i As Long
        Dim j As Long
        s = "Invalid message code from TWS: msgID=" & msgID & " version=" & version & vbCrLf & _
            "mIndexIn=" & mIndexIn & " mInputIndex=" & mInputIndex & vbCrLf & _
            "Buffer contents: " & vbCrLf
        Do While i < mIndexIn
            s = s & Format(i, "0000") & "  "
            For j = i To i + 50
                If j = mIndexIn Then Exit For
                s = s & IIf(mBufferIn(j) <> 0, Chr$(mBufferIn(j)), "_")
            Next
            i = i + 50
            If j < mIndexIn Then s = s & vbCrLf
        Loop
        Debug.Print s
        disconnect s
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
    'Debug.Print "Set checkpoint number=" & 1000 & "; minputindex=" & mInputIndex
Loop

Exit Sub

err:
If err.number = DataIncomplete Then
    rollbackBuffer
    retryCount = retryCount + 1
    Debug.Print "Data incomplete count: " & retryCount & "; CheckPointNumber: " & checkPointNumber
    Exit Sub
End If

mCommonServiceConsumer.ServiceProviderError err.number, err.Description, -1
disconnect "Error " & err.number & ": " & err.Description
End Sub

Private Sub checkpointBuffer()
mInputIndexChk = mInputIndex
End Sub

Private Sub createsocket()
If mSocket Is Nothing Then
    Set mWinsockForm = New WinsockForm
    Set mSocket = mWinsockForm.Winsock1
End If
End Sub

' Raises an ErrHistGenerationFailed error if IB cannot satisfy the request because
' it would start more than a year and a day ago.
Private Sub generateTwsHistoricalDataRequest(ByVal index As Long)
Dim TwsRequest As TWSHistoricalDataRequestQueueEntry
Dim reqNum As Long
Dim numBarsInsession As Long

reqNum = mHistoricalDataTable(index).numberOfTwsRequests + 1
mHistoricalDataTable(index).numberOfTwsRequests = reqNum

TwsRequest.id = index + HistoricalDataTickerBase + (mHistoricalDataTable(index).numberOfTwsRequests - 1) * HistoricalDataTickerIncrement

ReDim mHistoricalDataTable(index).requestKeys(reqNum - 1) As String

Set TwsRequest.contractSpec = mHistoricalDataTable(index).barSpecifier.Contract.specifier

' Set the TWS bar size and duration to retrieve as many bars as possible in one
' go, taking account of the following max durations per bar size.
'
' NB: using the D duration only returns bars in whole days, so requesting "1 D" for
' Z contract ending at 08:05 will only return 1 bar, for 08:00 on that day. But
' requesting "86400 S" gives 86400/barlengthsecs bars before the end time.
'
' Note also that the duration for any request must be such that the start time is not
' more than one year before the CURRENT-time-less-one-day (not 1 year before the end
' time in the request)
'
'   Bar size        Max duration
'   --------        ------------
'
'   1 sec           2000 S
'   5 sec           10000 S
'   15 sec          30000 S
'   30 sec          86400 S
'   1 minute        86400 S
'                   6 D
'   2 minutes       86400 S
'                   6 D
'   5 minutes       86400 S
'                   6 D
'   15 minutes      86400 S
'                   20 D
'                   2 W
'   30 minutes      86400 S
'                   34 D
'                   4 W
'                   1 M
'   1 hour          86400 S
'                   34 D
'                   4 w
'                   1 M
'   1 day           60 D
'                   12 M
'                   52 W
'                   1 Y
 
If mServerVersion >= 20 Then
    TwsRequest.endDateTime = Format(mHistoricalDataTable(index).barSpecifier.ToDate, "yyyymmdd hh\:nn\:ss")
    
    With mHistoricalDataTable(index).barSpecifier
        
        Select Case .BarLengthUnits
        Case TimePeriodUnits.TimePeriodDay
            TwsRequest.barSizeSetting = TWSBarSizes.BarSize1Day
            If .NumberOfBars >= NumDaysInYear Then
                If DateAdd("d", 1, .ToDate) > Now Then
                    TwsRequest.duration = "1 Y"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 60) & " D"
                End If
            ElseIf .NumberOfBars <= 60 Then
                TwsRequest.duration = permittedHistDays(.ToDate, .NumberOfBars) & " D"
            ElseIf .NumberOfBars <= 52 * NumDaysInWeek Then
                TwsRequest.duration = permittedHistWeeks(.ToDate, Int(.NumberOfBars / NumDaysInWeek)) & " W"
            Else
                TwsRequest.duration = permittedHistDays(.ToDate, 60) & " D"
            End If
        Case TimePeriodUnits.TimePeriodHour
            TwsRequest.barSizeSetting = TWSBarSizes.BarSize1Hour
            
            numBarsInsession = mHistoricalDataTable(index).Session.NumberOfBarsInSession( _
                                                                    .barLength, _
                                                                    .BarLengthUnits)
            
            If .NumberOfBars <= 24 Then
                TwsRequest.duration = permittedHistSeconds(.ToDate, .NumberOfBars * 3600) & " S"
            ElseIf -Int(-.NumberOfBars / numBarsInsession) <= 34 Then
                TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.NumberOfBars / numBarsInsession)) & " D"
            Else
                TwsRequest.duration = permittedHistDays(.ToDate, 34) & " D"
            End If
        Case TimePeriodUnits.TimePeriodMinute
            
            numBarsInsession = mHistoricalDataTable(index).Session.NumberOfBarsInSession( _
                                                                    .barLength, _
                                                                    .BarLengthUnits)
            
            Select Case .barLength
            Case 1
                TwsRequest.barSizeSetting = TWSBarSizes.BarSize1Minute
                If .NumberOfBars <= 1440 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .NumberOfBars * 60) & " S"
                ElseIf -Int(-.NumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.NumberOfBars / numBarsInsession) <= 6 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.NumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 6) & " D"
                End If
            Case 2
                TwsRequest.barSizeSetting = TWSBarSizes.BarSize2Minutes
                If .NumberOfBars <= 720 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .NumberOfBars * 120) & " S"
                ElseIf -Int(-.NumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.NumberOfBars / numBarsInsession) <= 6 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.NumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 6) & " D"
                End If
            Case 5
                TwsRequest.barSizeSetting = TWSBarSizes.BarSize5Minutes
                If .NumberOfBars <= 288 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .NumberOfBars * 300) & " S"
                ElseIf -Int(-.NumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.NumberOfBars / numBarsInsession) <= 6 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.NumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 6) & " D"
                End If
            Case 15
                TwsRequest.barSizeSetting = TWSBarSizes.BarSize15Minutes
                If .NumberOfBars <= 96 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .NumberOfBars * 900) & " S"
                ElseIf -Int(-.NumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.NumberOfBars / numBarsInsession) <= 20 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.NumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 20) & " D"
                End If
            Case 30
                TwsRequest.barSizeSetting = TWSBarSizes.BarSize30Minutes
                If .NumberOfBars <= 48 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .NumberOfBars * 1800) & " S"
                ElseIf -Int(-.NumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.NumberOfBars / numBarsInsession) <= 34 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.NumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 34) & " D"
                End If
            End Select
        Case TimePeriodUnits.TimePeriodMonth
            TwsRequest.barSizeSetting = TWSBarSizes.BarSize1Day
            If .NumberOfBars <= 12 Then
                TwsRequest.duration = permittedHistMonths(.ToDate, .NumberOfBars) & " M"
            Else
                If DateAdd("d", 1, .ToDate) > Now Then
                    TwsRequest.duration = "12 M"
                Else
                    TwsRequest.duration = permittedHistMonths(.ToDate, .NumberOfBars) & " M"
                End If
            End If
        Case TimePeriodUnits.TimePeriodSecond
            
            numBarsInsession = mHistoricalDataTable(index).Session.NumberOfBarsInSession( _
                                                                    .barLength, _
                                                                    .BarLengthUnits)
            
            Select Case .barLength
            Case 1
                
                If .NumberOfBars <= 2000 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .NumberOfBars) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 2000) & " S"
                End If
            Case 5
                TwsRequest.barSizeSetting = TWSBarSizes.BarSize5Seconds
                If .NumberOfBars <= 2000 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .NumberOfBars * 5) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 10000) & " S"
                End If
            Case 15
                TwsRequest.barSizeSetting = TWSBarSizes.BarSize15Seconds
                If .NumberOfBars <= 2000 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .NumberOfBars * 15) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 30000) & "S"
                End If
            Case 30
                TwsRequest.barSizeSetting = TWSBarSizes.BarSize15Seconds
                If .NumberOfBars <= 2880 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .NumberOfBars * 30) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                End If
            End Select
        Case TimePeriodUnits.TimePeriodWeek
            TwsRequest.barSizeSetting = TWSBarSizes.BarSize1Day
            If .NumberOfBars <= 52 Then
                TwsRequest.duration = permittedHistWeeks(.ToDate, .NumberOfBars) & " W"
            Else
                If DateAdd("d", 1, .ToDate) > Now Then
                    TwsRequest.duration = "52 W"
                Else
                    TwsRequest.duration = permittedHistWeeks(.ToDate, .NumberOfBars) & " W"
                End If
            End If
        Case TimePeriodUnits.TimePeriodVolume, TimePeriodUnits.TimePeriodTickMovement
            TwsRequest.barSizeSetting = TWSBarSizes.BarSize15Seconds
            TwsRequest.duration = "30000 S"
        Case Else
            Debug.Assert False
        End Select
    End With
End If

Select Case mHistoricalDataTable(index).barSpecifier.barType
Case TradeBuildSP.BarTypes.BarTypeAsk
    TwsRequest.whatToShow = TWSWhatToShowAsk
Case TradeBuildSP.BarTypes.BarTypeBid
    TwsRequest.whatToShow = TWSWhatToShowBid
Case TradeBuildSP.BarTypes.BarTypeTrade
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
    getDouble = DoubleFromString(s)
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
        'Debug.Print "Got string: mInputIndex=" & mInputIndex & " value=" & s
        mInputIndex = i + 1
        getString = s
        Exit Function
    End If
    s = s & Chr$(mBufferIn(i))
Next

' we've got to the end of the current data, so raise an error to abort
' this attempt to assemble a message
'Debug.Print "Failed to get string"
err.Raise DataIncomplete
End Function

Private Sub getAccountTime(ByVal version As Long)
Dim listendata As TWSSocketMsgAcctUpdateTime
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
Dim lContractWrapper As ContractWrapper
Dim lContractSpec As ContractSpecifier

Dim symbol As String
Dim localSymbol As String
Dim secType As String ' keep the socket string for logging
Dim expiry As String
Dim strike As Double
Dim optRight As String ' keep the socket string for logging
Dim exchange As String
Dim currencyCode As String
Dim marketName As String
Dim tradingClass As String

Dim lContractBuilder As ContractBuilder

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
Dim i As Long

Set lContractWrapper = New ContractWrapper

symbol = getString
secType = getString
expiry = getString
strike = getDouble
optRight = optionRightFromString(getString)
exchange = getString
currencyCode = getString
localSymbol = getString
marketName = getString
tradingClass = getString
lContractWrapper.contractId = getLong
minimumTick = getDouble
multiplier = getLong
    
Set lContractSpec = CreateContractSpecifier(localSymbol, _
                                            symbol, _
                                            exchange, _
                                            secTypeFromString(secType), _
                                            currencyCode, _
                                            expiry, _
                                            strike, _
                                            optionRightFromString(optRight))
    
Set lContractBuilder = CreateContractBuilder(lContractSpec)
lContractWrapper.tradeBuildContract = lContractBuilder.Contract
    

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
        Case "ALGO", "AVGCOST", "BASKET", "COND", "CONDORDER", "NONALGO", "OCA", "OPG", "OPGREROUT", "TIMEPRIO"
            ' not sure what these signify
        Case Else
            If orderTypesStr(i) <> "" Then
                logMessage "Unknown order type " & orderTypesStr(i) & " received from TWS", LogLevelLow
            End If
        End Select
    End If
Next

lContractWrapper.permittedOrderAttributes = permittedOrderAttributes
lContractWrapper.permittedOrderTypes = permittedOrderTypes

lContractWrapper.permittedOrderTifs = permittedOrderTifs

lContractWrapper.permittedStopTriggerMethods = TradeBuildSP.StopTriggerMethods.StopTriggerDefault Or _
                                        TradeBuildSP.StopTriggerMethods.StopTriggerDoubleBidAsk Or _
                                        TradeBuildSP.StopTriggerMethods.StopTriggerDoubleLast Or _
                                        TradeBuildSP.StopTriggerMethods.StopTriggerLast

validExchanges = getString
lValidExchanges = Split(validExchanges, ",")
    
    
lContractBuilder.Description = localSymbol & " on " & exchange
If expiry <> "" Then
    lContractBuilder.expiryDate = CDate(Left$(expiry, 4) & "/" & _
                                        Mid$(expiry, 5, 2) & "/" & _
                                        right$(expiry, 2))
End If

If version >= 2 Then
    lContractWrapper.priceMagnifier = getLong
Else
    If currencyCode = "GBP" Then
        lContractWrapper.priceMagnifier = 100
    Else
        lContractWrapper.priceMagnifier = 1
    End If
End If
lContractBuilder.multiplier = multiplier / lContractWrapper.priceMagnifier
lContractBuilder.TickSize = minimumTick * lContractWrapper.priceMagnifier

processContract lContractWrapper

If mLogLevel >= LogLevelHigh Then
    With lContractWrapper.tradeBuildContract
        listendata.msgType = TWSSocketInMsgTypes.CONTRACT_DATA
        listendata.contractId = lContractWrapper.contractId
        listendata.currencyCode = currencyCode
        listendata.exchange = exchange
        listendata.expiry = expiry
        listendata.localSymbol = localSymbol
        listendata.marketName = marketName
        listendata.minimumTick = minimumTick
        listendata.multiplier = multiplier
        listendata.OrderTypes = OrderTypes
        listendata.right = optRight
        listendata.secType = secType
        listendata.strike = strike
        listendata.symbol = symbol
        listendata.tradingClass = tradingClass
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

Select Case errorCode
Case 101
    ' max tickers reached
    failpoint = 350
    mTickerTable(id - MarketDepthTickerBase).dataReader.marketDataRequestFailed "Maximum number of tickers reached"
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

    If mHistoricalDataTable(index).bars.size <> 0 Then
        Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
        mHistoricalDataTable(index).dataRetrievedFromTWS = True
        mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
        mHistoricalDataTable(index).historicalDataReader.historicalDataAvailable
    Else
        logMessage "Error " & errorCode & " from TWS: " & errorMsg
        mHistoricalDataTable(index).historicalDataReader.historicalDataRequestFailed errorMsg
        releaseHistoricalDataTableEntry index
    End If
    
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

        mTickerTable(id Mod IncarnationIncrement).dataReader.marketDataRequestFailed errorMsg
        releaseTickerTableEntry (id Mod IncarnationIncrement)
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

    mTickerTable(id - MarketDepthTickerBase).dataReader.marketDepthRequestFailed "Max market depth requests exceeded"
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
    
        If mHistoricalDataTable(index).bars.size <> 0 Then
            Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
            mHistoricalDataTable(index).dataRetrievedFromTWS = True
            mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
            mHistoricalDataTable(index).historicalDataReader.historicalDataAvailable
        Else
            logMessage "Error " & errorCode & " from TWS: " & errorMsg
            mHistoricalDataTable(index).historicalDataReader.historicalDataRequestFailed errorMsg
            releaseHistoricalDataTableEntry index
        End If
        processNextHistDataRequest
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

    mCommonServiceConsumer.ServiceProviderError errorCode, errorMsg, -1
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
Dim timeString  As String

Dim symbol As String
Dim localSymbol As String
Dim secType As String ' keep the socket string for logging
Dim expiry As String
Dim strike As Double
Dim rightString As String
Dim exchange As String
Dim currencyCode As String

Dim execId As String
Dim AccountIdentifier  As String
Dim FillingExchange  As String
Dim side As String
Dim Price As Double
Dim Quantity As Long
Dim mapEntry As OrderIdMapEntry
Dim id As Long
Dim permId As Long
Dim clientID As Long
Dim lContractSpecifier As ContractSpecifier

Dim exec As Execution

id = getLong

On Error Resume Next
mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
On Error GoTo 0

symbol = getString
secType = getString
expiry = getString
strike = getDouble
rightString = getString
exchange = getString
currencyCode = getString
localSymbol = getString

Set lContractSpecifier = CreateContractSpecifier(localSymbol, _
                                                symbol, _
                                                exchange, _
                                                secTypeFromString(secType), _
                                                currencyCode, _
                                                expiry, _
                                                strike, _
                                                optionRightFromString(rightString))

Set exec.ContractSpecifier = lContractSpecifier
exec.OrderBrokerId = id
If Not mapEntry.order Is Nothing Then exec.OrderTradeBuildId = mapEntry.order.tradeBuildId
exec.execId = getString
timeString = getString
exec.Time = CDate(Left$(timeString, 4) & "/" & _
                    Mid$(timeString, 5, 2) & "/" & _
                    Mid$(timeString, 7, 2) & " " & _
                    right$(timeString, 8))
exec.AccountIdentifier = getString
exec.FillingExchange = getString
side = getString
exec.Action = IIf(side = "BUY", OrderActions.ActionBuy, OrderActions.ActionSell)
Quantity = getLong
Price = getDouble
exec.Price = adjustPrice(Price, lContractSpecifier)
    
If version >= 2 Then permId = getLong  ' permId

If version >= 3 Then clientID = getLong  ' clientID
   
If version >= 4 Then exec.IsLiquidation = getBoolean

Dim existingExec As Execution
On Error Resume Next
existingExec = mExecutions(execId)
If err.number = 0 Then
    ' we've had this execution before so nothing to do
    On Error GoTo 0
Else
    On Error GoTo 0
    If Not mapEntry.orderSubmitter Is Nothing Then
        mapEntry.orderSubmitter.executionDetails lContractSpecifier, _
                                                    exec.AccountIdentifier, _
                                                    exec.Action, _
                                                    exec.execId, _
                                                    exec.FillingExchange, _
                                                    exec.IsLiquidation, _
                                                    exec.OrderBrokerId, _
                                                    exec.OrderTradeBuildId, _
                                                    exec.Price, _
                                                    exec.Quantity, _
                                                    exec.Time
        ' notify a status update - this is in case this execution occurred during a period when
        ' TWS had lost its connection to the IB servers and this exec details is being notified
        ' in response to the reqExecutions call made when the connection is recovered. In these
        ' circumstances TWS will not itself issue an orderStatus.
        If exec.Quantity >= mapEntry.order.quantityRemaining Then
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
        listendata.secType = secType
        listendata.strike = .strike
        listendata.symbol = .symbol
    End With
    With exec
        listendata.accountNumber = .AccountIdentifier
        listendata.clientID = clientID
        listendata.execExchange = .FillingExchange
        listendata.execId = .execId
        listendata.orderId = id
        listendata.permId = permId
        listendata.Price = Price
        listendata.Quantity = .Quantity
        listendata.side = side
        listendata.Time = timeString
    End With
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgExecutionData, _
                    Me
End If

End Sub

Private Sub getFAData(ByVal version As Long)
Dim listendata As TWSSocketMsgReceiveFa
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
Static totalVolume As Long

Dim bardate As String
Dim timestamp As Date

Dim failpoint As Long
On Error GoTo err

If checkPointNumber = 0 Then
    hint = 0
    i = 0
    totalVolume = 0
    
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
            If mHistoricalDataTable(index).ignoreTimestamps Then
                Bar.timestamp = Bar.timestamp + mHistoricalDataTable(index).barSpecifier.Contract.sessionStartTime
            End If
            
            If i = 0 Then
                If mHistoricalDataTable(index).barSpecifier.BarLengthUnits = TimePeriodVolume Or _
                    mHistoricalDataTable(index).barSpecifier.BarLengthUnits = TimePeriodTickMovement _
                Then
                    mHistoricalDataTable(index).barSpecifier.ToDate = mHistoricalDataTable(index).Session.OffsetBarStartTime(timestamp, 15, TimePeriodSecond, -1)
                Else
                    mHistoricalDataTable(index).barSpecifier.ToDate = mHistoricalDataTable(index).Session.OffsetBarStartTime(timestamp, mHistoricalDataTable(index).barSpecifier.barLength, mHistoricalDataTable(index).barSpecifier.BarLengthUnits, -1)
                End If
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
            totalVolume = totalVolume + Bar.volume
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
                mHistoricalDataTable(index).SessionBuilder.setSessionCurrentTime Bar.timestamp
                If mHistoricalDataTable(index).barSpecifier.includeBarsOutsideSession Or _
                    mHistoricalDataTable(index).Session.isTimeInSession(Bar.timestamp) Or _
                    mHistoricalDataTable(index).ignoreTimestamps _
                Then
                    hint = mHistoricalDataTable(index).bars.add(Bar, Bar.timestamp, hint)
                End If
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

If mHistoricalDataTable(index).barSpecifier.BarLengthUnits = TimePeriodVolume Then
    mHistoricalDataTable(index).barSpecifier.NumberOfBars = mHistoricalDataTable(index).barSpecifier.NumberOfBars - (totalVolume / mHistoricalDataTable(index).barSpecifier.barLength)
Else
    mHistoricalDataTable(index).barSpecifier.NumberOfBars = mHistoricalDataTable(index).barSpecifier.NumberOfBars - itemCount
End If
If mHistoricalDataTable(index).barSpecifier.NumberOfBars <= 0 Then
    Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
    mHistoricalDataTable(index).dataRetrievedFromTWS = True
    mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
    mHistoricalDataTable(index).historicalDataReader.historicalDataAvailable
Else
    ' request the next lot of bars
    generateTwsHistoricalDataRequest index
End If


Exit Sub

err:

If err.number = ErrHistGenerationFailed Then
    mHistoricalDataTable(index).barSpecifier.NumberOfBars = 0
    Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
    mHistoricalDataTable(index).dataRetrievedFromTWS = True
    mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
    mHistoricalDataTable(index).historicalDataReader.historicalDataAvailable
    Exit Sub
End If

err.Raise err.number, _
        "IBTWSSP" & "." & "TWSAPI" & "::" & "getHistoricalData" & "." & failpoint & _
        IIf(err.source <> "", vbCrLf & err.source, ""), _
        err.Description


End Sub

Private Sub getManagedAccounts(ByVal version As Long)
Dim listendata As TWSSocketMsgManagedAccts
Dim accountsList As String: accountsList = getString

RaiseEvent managedAccounts(accountsList)

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
Dim id As Long: id = getLong
Dim position As Long: position = getLong
Dim operation As Long: operation = getLong
Dim side As Long: side = getLong
Dim Price As Double: Price = getDouble
Dim size As Long: size = getLong

processMarketDepth id - MarketDepthTickerBase, _
                position, _
                operation, _
                side, _
                Price, _
                size

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.MARKET_DEPTH
    listendata.operation = operation
    listendata.position = position
    listendata.Price = Price
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
Dim timestamp As Date
Dim id As Long: id = getLong
Dim position As Long: position = getLong
Dim marketmaker As String: marketmaker = getString
Dim operation As Long: operation = getLong
Dim side As Long: side = getLong
Dim Price As Double: Price = getDouble
Dim size As Long: size = getLong

With mTickerTable(id - MarketDepthTickerBase)
    If .inUse Then
    
        timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.TimeZone)
        
        .dataConsumer.UpdateMktDepthL2 timestamp, _
                        position, _
                        marketmaker, _
                        operation, _
                        side, _
                        Price, _
                        size
                        
    End If
End With

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.MARKET_DEPTH_L2
    listendata.marketmaker = marketmaker
    listendata.operation = operation
    listendata.position = position
    listendata.Price = Price
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
'Dim lOrder As Order
Dim id As Long

Dim symbol As String
Dim localSymbol As String
Dim secType As String ' keep the socket string for logging
Dim expiry As String
Dim strike As Double
Dim rightString As String
Dim exchange As String
Dim currencyCode As String

Dim Action As String
Dim orderType As String
Dim timeInForce As String

Dim lContractSpecifier As ContractSpecifier

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
symbol = getString
secType = getString
expiry = getString
strike = getDouble
rightString = getString
exchange = getString
currencyCode = getString
If version >= 2 Then localSymbol = getString

Set lContractSpecifier = CreateContractSpecifier(localSymbol, _
                                                symbol, _
                                                exchange, _
                                                secTypeFromString(secType), _
                                                currencyCode, _
                                                expiry, _
                                                strike, _
                                                optionRightFromString(rightString))


' read order fields
'With lOrder
    Action = getString
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
    avgFillPrice = adjustPrice(avgFillPrice, mapEntry.order.Contract.specifier)
    lastFillPrice = adjustPrice(lastFillPrice, mapEntry.order.Contract.specifier)
    
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

Dim symbol As String
Dim localSymbol As String
Dim secType As String ' keep the socket string for logging
Dim expiry As String
Dim strike As Double
Dim rightString As String
Dim exchange As String
Dim currencyCode As String

Dim lContractSpecifier As ContractSpecifier

symbol = getString
secType = getString
expiry = getString
strike = getDouble
rightString = getString
currencyCode = getString
If version >= 2 Then localSymbol = getString

Set lContractSpecifier = CreateContractSpecifier(localSymbol, _
                                                symbol, _
                                                exchange, _
                                                secTypeFromString(secType), _
                                                currencyCode, _
                                                expiry, _
                                                strike, _
                                                optionRightFromString(rightString))

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
        listendata.right = rightString
        listendata.secType = secType
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
Dim lContractSpecifier As ContractSpecifier
Dim lContractBuilder As ContractBuilder
Dim tickerid As Long
Dim numberOfElements As Long
Dim rank As Long
Dim distance As String
Dim benchmark As String
Dim projection As String
Dim i As Long

Dim symbol As String
Dim localSymbol As String
Dim secType As SecurityTypes
Dim expiry As String
Dim strike As Double
Dim optRight As OptionRights
Dim exchange As String
Dim currencyCode As String

tickerid = getLong
numberOfElements = getLong

For i = 0 To numberOfElements - 1
    rank = getLong

    symbol = getString
    secType = secTypeFromString(getString)
    expiry = getString
    strike = getDouble
    optRight = optionRightFromString(getString)
    exchange = getString
    currencyCode = getString
    localSymbol = getString
    
    Set lContractSpecifier = CreateContractSpecifier(localSymbol, _
                                                    symbol, _
                                                    exchange, _
                                                    secType, _
                                                    currencyCode, _
                                                    expiry, _
                                                    strike, _
                                                    optRight)
    Set lContractBuilder = CreateContractBuilder(lContractSpecifier)
    getString 'marketName
    getString 'tradingClass
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
Dim Price As Double
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

Price = getDouble

failpoint = 400

If version >= 2 Then size = getLong
If version >= 3 Then canAutoExecute = getBoolean

failpoint = 500

processTickPrice tickerid Mod IncarnationIncrement, _
                tickType, _
                Price, _
                size

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.TICK_PRICE
    listendata.Price = Price
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

failpoint = 400

processTickSize tickerid Mod IncarnationIncrement, _
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
    logMessage "Connection to TWS closed: " & connectionString, LogLevelLow
Else
    If mConnectionRetryIntervalSecs = 0 Then
        cancelContractRequests "Lost connection to TWS"
        RaiseEvent ConnectionToTWSClosed(False)
        logMessage "Connection to TWS closed" & _
                    connectionString, _
                    LogLevelLow
    Else
        retryConnection
        RaiseEvent ConnectionToTWSClosed(True)
        logMessage "Connection to TWS closed - reconnecting: " & _
                    connectionString, _
                    LogLevelLow
    End If
End If
End Sub

Private Function LegOpenCloseFromString(ByVal value As String) As LegOpenClose
Select Case UCase$(value)
Case ""
    LegOpenCloseFromString = LegUnknownPos
Case "SAME"
    LegOpenCloseFromString = LegSamePos
Case "OPEN"
    LegOpenCloseFromString = LegOpenPos
Case "CLOSE"
    LegOpenCloseFromString = LegClosePos
End Select
End Function

Private Function LegOpenCloseToString(ByVal value As LegOpenClose) As String
Select Case value
Case LegSamePos
    LegOpenCloseToString = "SAME"
Case LegOpenPos
    LegOpenCloseToString = "OPEN"
Case LegClosePos
    LegOpenCloseToString = "CLOSE"
End Select
End Function

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

Private Function orderActionFromString(ByVal value As String) As OrderActions
Select Case UCase$(value)
Case "BUY"
    orderActionFromString = OrderActions.ActionBuy
Case "SELL"
    orderActionFromString = OrderActions.ActionSell
End Select
End Function

Private Function orderActionToString(ByVal value As OrderActions) As String
Select Case value
Case OrderActions.ActionBuy
    orderActionToString = "BUY"
Case OrderActions.ActionSell
    orderActionToString = "SELL"
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

Private Function permittedHistDays( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
permittedHistDays = DateDiff("d", DateAdd("yyyy", -1, Now), reqEndDate)
If permittedHistDays = 0 Then err.Raise ErrHistGenerationFailed
If permittedHistDays > desiredBars Then permittedHistDays = desiredBars
End Function

Private Function permittedHistMonths( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
permittedHistMonths = DateDiff("m", DateAdd("m", 1, DateAdd("yyyy", -1, Now)), reqEndDate)
If permittedHistMonths = 0 Then err.Raise ErrHistGenerationFailed
If permittedHistMonths > desiredBars Then permittedHistMonths = desiredBars
End Function

Private Function permittedHistSeconds( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
' the 86398 should really be 86400 but allows for clocks not being
' properly synchronised
permittedHistSeconds = DateDiff("s", DateAdd("yyyy", -1, Now), reqEndDate) + 86398
If permittedHistSeconds = 0 Then err.Raise ErrHistGenerationFailed
If permittedHistSeconds > desiredBars Then permittedHistSeconds = desiredBars
End Function

Private Function permittedHistWeeks( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
permittedHistWeeks = DateDiff("w", DateAdd("w", 1, DateAdd("yyyy", -1, Now)), reqEndDate)
If permittedHistWeeks = 0 Then err.Raise ErrHistGenerationFailed
If permittedHistWeeks > desiredBars Then permittedHistWeeks = desiredBars
End Function

Private Sub ProcessAsk(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal Price As Double, _
                            ByVal size As Long)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    .dataConsumer.Ask timestamp, _
                Price, _
                size
End With
End Sub

Private Sub ProcessBid(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal Price As Double, _
                            ByVal size As Long)
With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    .dataConsumer.Bid timestamp, _
                Price, _
                size
End With
End Sub

Private Sub ProcessClose(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal Price As Double)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    .dataConsumer.PreviousClose timestamp, Price
End With
End Sub

Private Sub processContract(ByVal pContractWrapper As ContractWrapper)

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

mContractDetailsRequest.contractWrappers.add pContractWrapper

End Sub

Private Sub processContractError(ByVal reason As String)
Dim failpoint As Long

On Error GoTo err

If mContractDetailsRequest.contractWrappers.Count > 0 Then
    ' we are processing the forced invalid contract details request
    ' for a request for which we have received one or more contract details
    failpoint = 100
    mContractDetailsRequest.requester.contractsLoaded mContractDetailsRequest
Else
    ' this is a genuine error response to a contract details request
    failpoint = 200
    mContractDetailsRequest.requester.contractSpecifierInvalid mContractDetailsRequest, reason
End If
Set mContractDetailsRequest.contractSpec = Nothing
Set mContractDetailsRequest.contractWrappers = Nothing
Set mContractDetailsRequest.tradeBuildContractsBuilder = Nothing
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
                            ByVal Price As Double)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    .dataConsumer.High timestamp, Price
    
End With
End Sub

Private Sub ProcessLast(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal Price As Double, _
                            ByVal size As Long)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    .dataConsumer.Trade timestamp, _
            Price, _
            size
End With
End Sub

Private Sub ProcessLow(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal Price As Double)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    .dataConsumer.Low timestamp, Price

End With
End Sub

Private Sub processMarketDepth(ByVal id As Long, _
                                ByVal position As Long, _
                                ByVal operation As DOMOperations, _
                                ByVal side As DOMSides, _
                                ByVal Price As Double, _
                                ByVal size As Long)

Dim timestamp As Date

With mTickerTable(id)

    timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.TimeZone)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    .dataConsumer.UpdateMktDepth timestamp, _
                position, _
                operation, _
                side, _
                Price, _
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

Private Sub processTickPrice(ByVal tickerid As Long, _
                            ByVal tickType As TWSSocketTickTypes, _
                            ByVal Price As Double, _
                            ByVal size As Long)
Dim timestamp As Date

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    If Price < (0.5 * .prevTrade) Then
        Exit Sub        ' ignore invalid prices
    End If
    
    timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.TimeZone)
    
    Select Case tickType
    Case TICK_BID
        size = IIf(size <> 0, size, .prevBidSize)
        ProcessBid timestamp, _
                    tickerid, _
                    Price, _
                    size
        .prevBid = Price
        If size <> 0 Then .prevBidSize = size
    
    Case TICK_ASK
        size = IIf(size <> 0, size, .prevAskSize)
        ProcessAsk timestamp, _
                    tickerid, _
                    Price, _
                    size
        
        .prevAsk = Price
        If size <> 0 Then .prevAskSize = size
    
    Case TICK_LAST
        size = IIf(size <> 0, size, .prevTradeSize)
        ProcessLast timestamp, _
                    tickerid, _
                    Price, _
                    size
        .prevTrade = Price
        If size <> 0 Then .prevTradeSize = size
    
    Case TICK_HIGH
        ProcessHigh timestamp, _
                    tickerid, _
                    Price
    Case TICK_LOW
        ProcessLow timestamp, _
                    tickerid, _
                    Price
    Case TICK_CLOSE
        ProcessClose timestamp, _
                    tickerid, _
                    Price
    End Select
End With

End Sub

Private Sub processTickSize(ByVal tickerid As Long, _
                            ByVal tickType As TWSSocketTickTypes, _
                            ByVal size As Long)
Dim timestamp As Date

With mTickerTable(tickerid)
    
    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.TimeZone)
    
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
        
        .prevAskSize = size
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
        If Not .Contract Is Nothing Then
            If Not newConnection Then cancelMarketDataEx i, False
            .incarnation = .incarnation + IncarnationIncrement
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
    .ignoreTimestamps = False
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
    .incarnation = .incarnation + IncarnationIncrement
    Set .Contract = Nothing
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
Dim ContractSpecifier As ContractSpecifier
Set ContractSpecifier = mContractDetailsRequest.contractSpec

If ContractSpecifier.exchange <> "" And Not IsValidExchangeCode(ContractSpecifier.exchange) Then
    processContractError "invalid exchange"
    Exit Sub
End If

Const version = 1

addData REQ_CONTRACT_DATA
addData version
With ContractSpecifier
    addData .symbol
    addData secTypeToString(.secType)
    addData Left$(.expiry, 6)
    addDouble .strike
    addData optionRightToString(.right)
    addData .exchange
    addData .currencyCode
    addData .localSymbol
End With
send
End Sub

Private Sub reqHistoricalData( _
                request As TWSHistoricalDataRequestQueueEntry)
Debug.Print "ReqHistData from TWS: barsize=" & request.barSizeSetting & "; duration=" & request.duration
Const version = 3

If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

addData REQ_HISTORICAL_DATA
addData version
addData request.id
addData request.contractSpec.symbol
addData secTypeToString(request.contractSpec.secType)
addData request.contractSpec.expiry
addDouble request.contractSpec.strike
addData optionRightToString(request.contractSpec.right)
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
    ' not supported
    
'    addData request.contractSpec.ComboLegs.Count
'
'    Dim comboLeg As TradeBuildSP.IComboLeg
'    For Each comboLeg In request.contractSpec.ComboLegs
'        addData comboLeg.contractId
'        addData comboLeg.Ratio
'        addData orderActionToString(comboLeg.Action)
'        addData comboLeg.exchange
'    Next
End If

send

mConcurrentHistDataRequests = mConcurrentHistDataRequests + 1
      
End Sub

Private Sub reqMarketData(ByVal tickerTableIndex As Long)

Const version = 3

addData REQ_MKT_DATA
addData version
addData tickerTableIndex + mTickerTable(tickerTableIndex).incarnation
With mTickerTable(tickerTableIndex).Contract.specifier
    addData .symbol
    addData secTypeToString(.secType)
    addData Left$(.expiry, 6)
    addDouble .strike
    addData optionRightToString(.right)
    addData .exchange
    addData .currencyCode
    If mServerVersion >= 2 Then addData .localSymbol
    ' Add combo legs for BAG requests
    If mServerVersion >= 8 And .secType = SecurityTypes.SecTypeCombo Then
        ' not currently aupported
        
'        addData .ComboLegs.Count
'        Dim comboLeg As TradeBuildSP.IComboLeg
'        For Each comboLeg In .ComboLegs
'            With comboLeg
'                addData .contractId
'                addData .Ratio
'                addData orderActionToString(.Action)
'                addData .exchange
'            End With
'        Next
    End If
    
    send
End With

End Sub

Private Sub reqMarketDepth(ByVal tickerid As Long, _
                        ByVal contractSpec As ContractSpecifier)

Const version = 1

addData REQ_MKT_DEPTH
addData version
addData tickerid + MarketDepthTickerBase
With contractSpec
    addData .symbol
    addData secTypeToString(.secType)
    addData Left$(.expiry, 6)
    addDouble .strike
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
            If reEstablish Then reqMarketDepth i, .Contract.specifier
            .dataConsumer.resetMarketDepth ConvertDateUTCToTZ(GetTimestampUTC, .Contract.TimeZone), _
                                            reEstablish
        End If
    End With
Next
End Sub

Private Sub retryConnection()
If mConnectionRetryIntervalSecs <> 0 Then
    If mConnectionTimer Is Nothing Then
        Set mConnectionTimer = CreateIntervalTimer(mConnectionRetryIntervalSecs, ExpiryTimeUnitSeconds)
    End If
    mConnectionTimer.StartTimer
End If
End Sub

Private Sub rollbackBuffer()
'Debug.Print "Rollback: mInputIndex=" & mInputIndexChk
mInputIndex = mInputIndexChk
End Sub

Private Function secTypeFromString(ByVal value As String) As SecurityTypes
Select Case UCase$(value)
Case "STK"
    secTypeFromString = SecTypeStock
Case "FUT"
    secTypeFromString = SecTypeFuture
Case "OPT"
    secTypeFromString = SecTypeOption
Case "FOP"
    secTypeFromString = SecTypeFuturesOption
Case "CASH"
    secTypeFromString = SecTypeCash
Case "IND"
    secTypeFromString = SecTypeIndex
End Select
End Function

Private Function secTypeToString(ByVal value As SecurityTypes) As String
Select Case value
Case SecTypeStock
    secTypeToString = "STK"
Case SecTypeFuture
    secTypeToString = "FUT"
Case SecTypeOption
    secTypeToString = "OPT"
Case SecTypeFuturesOption
    secTypeToString = "FOP"
Case SecTypeCash
    secTypeToString = "CASH"
Case SecTypeIndex
    secTypeToString = "IND"
End Select
End Function

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

Dim specifier As ContractSpecifier

Set specifier = pOrder.Contract.specifier

Const version = 18
addData PLACE_ORDER
addData version
addData pOrder.brokerId

' send contract fields
With specifier
    addData .symbol
    addData secTypeToString(.secType)
    addData Left$(.expiry, 6)
    addDouble .strike
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
    addData orderActionToString(.Action)
    addData .Quantity
    addData orderTypeToString(.orderType)
    If mServerVersion <= 10 And _
        (UCase$(specifier.exchange) = "LIFFE" Or _
        UCase$(specifier.exchange) = "LSE") _
    Then
        addDouble .limitPrice / 100
        addDouble .triggerPrice / 100
    Else
        addDouble .limitPrice
        addDouble .triggerPrice
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
If mServerVersion >= 8 And specifier.secType = SecurityTypes.SecTypeCombo Then
'    With specifier
'        If .ComboLegs.Count = 0 Then
'            addData 0
'        Else
'            addData .ComboLegs.Count
'            Dim comboLeg As TradeBuildSP.IComboLeg
'            For Each comboLeg In .ComboLegs
'                With comboLeg
'                    addData .contractId
'                    addData .Ratio
'                    addData orderActionToString(.Action)
'                    addData .exchange
'                    addData LegOpenCloseToString(.openClose)
'                End With
'            Next
'        End If
'    End With
End If

With pOrder
    If mServerVersion >= 9 Then addData ""          ' sharesAllocation (deprecated)
    
    If mServerVersion >= 10 Then addData .discretionaryAmount
    
    If mServerVersion >= 11 Then
        If .goodAfterTime = 0 Then
            addData ""
        Else
            addData Format(.goodAfterTime, "yyyymmdd hh\:nn\:ss ") & .goodAfterTimeTZ
        End If
    End If
    
    If mServerVersion >= 12 Then
        If .goodTillDate = 0 Then
            addData ""
        Else
            addData Format(.goodTillDate, "yyyymmdd hh\:nn\:ss ") & .goodTillDateTZ
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

Private Function showBuffer( _
                ByRef buffer() As Byte, _
                ByVal length As Long) As String
Dim i As Long
Dim char As Byte

showBuffer = String(UBound(buffer) + 1, "_")

For i = 0 To length - 1
    char = buffer(i)
    If char <> 0 Then Mid$(showBuffer, i + 1, 1) = Chr$(char)
Next
End Function

Private Sub submitNextContractDetailsRequest()
Dim failpoint As Long

On Error GoTo err

If Not mConnectionState = ConnConnected Then Exit Sub
If Not mContractDetailsRequest.contractSpec Is Nothing Then Exit Sub
If mTwsContractDetailsRequestQueue.size = 0 Then Exit Sub
    
failpoint = 100
On Error Resume Next
mContractDetailsRequest = mTwsContractDetailsRequestQueue.removeFromFront
If err.number <> 0 Then Stop
On Error GoTo err

failpoint = 200
Set mContractDetailsRequest.contractWrappers = New Collection
reqContract

Exit Sub

err:
err.Raise err.number, , "TWSAPI::submitNextContractDetailsRequest(failpoint=" & failpoint & "): " & err.Description
End Sub


