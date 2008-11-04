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

'Event FAData( _
'                ByVal DataType As FADataTypes, _
'                ByVal XMLData As String)
'
'Event managedAccounts( _
'                ByVal accountsList As String)
'
'Event updateAccountTime( _
'                ByVal accountTime As String)
'
'Event updateAccountValue( _
'                ByVal key As String, _
'                ByVal value As String, _
'                ByVal currencyCode As String, _
'                ByVal accountName As String)
'
'Event updateNewsBulletin( _
'                ByVal msgID As Integer, _
'                ByVal msgType As BulletinTypes, _
'                ByVal message As String, _
'                ByVal origExchange As String)
'
'Event updatePortfolioValue( _
'                ByVal pContractSpecifier As ContractSpecifier, _
'                ByVal position As Long, _
'                ByVal marketPrice As Double, _
'                ByVal marketValue As Double, _
'                ByVal averageCost As Double, _
'                ByVal unrealizedPNL As Double, _
'                ByVal realizedPNL As Double, _
'                ByVal accountName As String)


'================================================================================
' Constants
'================================================================================

Private Const ProjectName                As String = "IBTWSSP26"
Private Const ModuleName                As String = "TWSAPI"

Private Const ClientVersion As Long = 32

Private Const ErrHistGenerationFailed = vbObjectError

Private Const HistoricalDataTickerBase As Long = &H410000
Private Const HistoricalDataTickerIncrement As Long = &H100

Private Const IncarnationIncrement As Long = &H1000&

Private Const MarketDepthTickerBase As Long = &H400000

Private Const MaxConcurrentHistDataRequests As Long = 50 ' limit imposed by IB

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

Private Enum IdTypes
    IdTypeNone
    IdTypeRealtimeData
    IdTypeMarketDepth
    IdTypeHistoricalData
    IdTypeOrder
End Enum

Private Enum InternalErrorCodes
    DataIncomplete = vbObjectError + 4327   ' let's hope nothing else uses this number!
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
    barSpecifier            As BarDataSpecifier
    SessionBuilder          As SessionBuilder
    Session                 As Session
    earliestReuseTime       As Date
    numberOfTwsRequests     As Long         ' the number of historical data requests
                                            ' made to TWS
    barType                 As BarTypes
    barLength               As Double       ' days
    historicalDataReader    As HistDataReader
    nextBarIndex            As Long
    bars                    As LinkedList
    requestKeys()           As String
End Type

Private Type StatsEntry
    lastSecondCount         As Long
    lastSecondTime          As Single
    lastPeriodCount         As Long
    lastPeriodTime          As Single
    totalCount              As Long
    totalTime               As Single
    maxSecondCount          As Long
    longestTime             As Single
    shortestTime            As Single
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
Private mCheckPointNumber As Long

Private mServerVersion As Long
Private mCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer

Private mConnectionState As ConnectionStates

Private mTickerTable() As TickerTableEntry

Private mContractDetailsRequest As TWSContractDetailsRequestQueueEntry
Private mFirstContractDetailsReceived As Boolean

Private mNextOrderID  As Long

Private mTWSLogLevel As TWSLogLevels

Private mTwsHistDataRequestQueue As Queue
Private mConcurrentHistDataRequests As Long

Private mHistoricalDataTable() As HistoricalDataRequestTableEntry

Private mOrderIDMap As Collection

Private mTwsContractDetailsRequestQueue As Queue

Private mExecutions As Collection

Private mPerformanceStats(MAX_SOCKET_INMSG) As StatsEntry

Private WithEvents mPerformanceTimerSecond As IntervalTimer
Attribute mPerformanceTimerSecond.VB_VarHelpID = -1
Private WithEvents mPerformanceTimerPeriod As IntervalTimer
Attribute mPerformanceTimerPeriod.VB_VarHelpID = -1

Private mPerformanceElapsedTimer As ElapsedTimer

Private mPerformanceLogger As Logger
Private mSocketLogger As Logger

Private mMessageBuilder As StringBuilder

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

ReDim mTickerTable(3) As TickerTableEntry
ReDim mBufferIn(1023) As Byte
ReDim mBufferOut(15) As Byte

createLoggers

mConnectionState = ConnectionStates.ConnNotConnected

Set mOrderIDMap = New Collection
Set mExecutions = New Collection

mNextOrderID = 1

mTWSLogLevel = TWSLogLevelError ' this is the TWS default

Set mTwsHistDataRequestQueue = New Queue
mTwsHistDataRequestQueue.initialCapacity = 20

ReDim mHistoricalDataTable(3) As HistoricalDataRequestTableEntry

Set mTwsContractDetailsRequestQueue = New Queue
mTwsContractDetailsRequestQueue.initialCapacity = 20

Set mPerformanceElapsedTimer = New ElapsedTimer
End Sub


Private Sub Class_Terminate()
Debug.Print "TWSAPI terminated"
End Sub

'================================================================================
' mConnectionTimer Event Handlers
'================================================================================

Private Sub mConnectionTimer_TimerExpired()
On Error GoTo Err
Connect
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mConnectionTimer_TimerExpired", _
                Err.Description
End Sub

'================================================================================
' mPerformanceTimerSecond Event Handlers
'================================================================================

Private Sub mPerformanceTimerPeriod_TimerExpired()
Dim i As Long
Dim s As String

On Error GoTo Err

For i = 0 To TWSSocketInMsgTypes.MAX_SOCKET_INMSG
    With mPerformanceStats(i)
        .lastPeriodCount = .lastPeriodCount + .lastSecondCount
        .lastPeriodTime = .lastPeriodTime + .lastSecondTime
        .totalCount = .totalCount + .lastSecondCount
        .totalTime = .totalTime + .lastSecondTime
        
        If .totalCount <> 0 Then
            s = gSocketInMsgTypeToString(i)
            s = s & Format(.lastPeriodCount, " 00000")
            If .lastPeriodCount <> 0 Then
                s = s & Format(.lastPeriodTime / .lastPeriodCount, " 00000.0")
            Else
                s = s & "     0.0"
            End If
            s = s & Format(.totalCount, " 00000000")
            If .totalCount <> 0 Then
                s = s & Format(.totalTime / .totalCount, " 00000.0")
            Else
                s = s & "     0.0"
            End If
            s = s & Format(.maxSecondCount, " 0000")
            s = s & Format(.longestTime, " 00000.0")
            s = s & Format(.shortestTime, " 00000.0")
            mPerformanceLogger.Log LogLevelMediumDetail, s
        End If
        
        .lastPeriodCount = 0
        .lastPeriodTime = 0
        .lastSecondCount = 0
        .lastSecondTime = 0
    End With
Next
Exit Sub

Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mPerformanceTimerPeriod_TimerExpired", _
                Err.Description
End Sub

Private Sub mPerformanceTimerSecond_TimerExpired()
Dim i As Long

On Error GoTo Err

For i = 0 To TWSSocketInMsgTypes.MAX_SOCKET_INMSG
    With mPerformanceStats(i)
        .lastPeriodCount = .lastPeriodCount + .lastSecondCount
        .lastPeriodTime = .lastPeriodTime + .lastSecondTime
        .totalCount = .totalCount + .lastSecondCount
        .totalTime = .totalTime + .lastSecondTime
        .lastSecondCount = 0
        .lastSecondTime = 0
    End With
Next
Exit Sub

Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mPerformanceTimerSecond_TimerExpired", _
                Err.Description
End Sub

'================================================================================
' Socket Event Handlers
'================================================================================

Private Sub mSocket_Close()
On Error GoTo Err
handleTwsDisconnection
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mSocket_Close", _
                Err.Description
End Sub

Private Sub mSocket_Connect()
On Error GoTo Err
mConnectionState = ConnectionStates.ConnConnecting
RaiseEvent Connecting

mIndexIn = 0
mIndexOut = 0
mInputIndex = 0
mCheckPointNumber = 0

logMessage "Connecting to TWS: negotiating versions and client id", LogLevelNormal
addData ClientVersion
send
addData mClientID
send
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mSocket_Connect", _
                Err.Description
End Sub

Private Sub mSocket_DataArrival( _
                ByVal bytesTotal As Long)
Dim bytesReceived As Long
Static processingMessages As Boolean

On Error GoTo Err

mSocket.GetData mSocketInputData, vbArray + vbByte
bytesReceived = UBound(mSocketInputData) + 1
'Debug.Print "Received " & bytesReceived & " from socket: " & showBuffer(mSocketInputData, bytesReceived)

If mSocketLogger.isLoggable(LogLevelMediumDetail) Then
    mSocketLogger.Log LogLevels.LogLevelMediumDetail, _
                        showBuffer(mSocketInputData, bytesReceived)
End If

If mIndexIn = mInputIndex Then
    ' no unprocessed data so start at the beginning of the buffer again
    'Debug.Print "Reverting to start of buffer (no unprocessed data)"
    mIndexIn = 0
    mInputIndex = 0
    mInputIndexChk = 0
End If

If mIndexIn + bytesReceived > UBound(mBufferIn) + 1 Then
    If bytesReceived <= (UBound(mBufferIn) + 1 - mIndexIn + mInputIndex) Then
        ' copy the existing contents down to the start of the buffer
        'Debug.Print "Moving unprocessed data (" & mIndexIn - mInputIndex & " bytes) from " & mInputIndex & " to start of buffer"
        MoveMemory mBufferIn(0), mBufferIn(mInputIndex), mIndexIn - mInputIndex
        mIndexIn = mIndexIn - mInputIndex
        mInputIndex = 0
        mInputIndexChk = 0
    Else
        If mSocketLogger.isLoggable(LogLevelHighDetail) Then
            mSocketLogger.Log LogLevels.LogLevelHighDetail, _
                            "Extending socket buffer by " & IIf(bytesReceived > 1024, bytesReceived, 1024) & " bytes"
        End If
        ReDim Preserve mBufferIn(UBound(mBufferIn) + IIf(bytesReceived > 1024, bytesReceived, 1024)) As Byte
    End If
End If
'Debug.Print "Appending socket message to buffer at " & mIndexIn
CopyMemory mBufferIn(mIndexIn), mSocketInputData(0), bytesReceived
mIndexIn = mIndexIn + bytesReceived

If mIndexIn <> mInputIndex Then
    If processingMessages Then
        ' this is a re-entrant call to DataArrival. We don't want to call checkMessages
        ' here, because the current call to it will process the data we've just
        ' added to the buffer
        mSocketLogger.Log LogLevels.LogLevelDetail, _
                        "Re-entrant call to mSocket_DataArrival"
    Else
        processingMessages = True
        checkmessages
        processingMessages = False
    End If
End If

Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mSocket_DataArrival", _
                Err.Description

End Sub

Private Sub mSocket_Error(ByVal number As Integer, _
                        Description As String, _
                        ByVal Scode As Long, _
                        ByVal source As String, _
                        ByVal HelpFile As String, _
                        ByVal HelpContext As Long, _
                        CancelDisplay As Boolean)
Dim errorNum As MSWinsockLib.ErrorConstants

On Error GoTo Err
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
                        LogLevelNormal
            
            cancelContractRequests "Failed to connect to TWS"
            RaiseEvent ConnectFailed(Description, False)
        Else
            logMessage "Failed to connect to TWS - retrying: " & _
                        Description & ": " & _
                        connectionString, _
                        LogLevelNormal
            
            RaiseEvent ConnectFailed(Description, True)
            retryConnection
        End If
    Else
        logMessage "Socket error " & number & ": " & _
                    Description & ": " & _
                    connectionString, _
                    LogLevelSevere
                    
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
                LogLevelSevere
    RaiseEvent ConnectFailed(Description, False)
End Select

Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "mSocket_Error", _
                Err.Description

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

Friend Property Get contractDetailsRequestQueueLength() As Long
contractDetailsRequestQueueLength = mTwsContractDetailsRequestQueue.size
End Property

Friend Property Get InBuffer() As String
InBuffer = showBuffer(mBufferIn, mIndexIn - 1)
End Property

Friend Property Get initialOrderId() As Long
initialOrderId = mNextOrderID
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

On Error GoTo Err
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
                If Err.number = 0 Then
                    ' request was still in queue - no further action required
                ElseIf Err.number <> ErrorCodes.ErrIllegalArgumentException Then
                    Err.Raise Err.number
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
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "CancelHistoricalData", _
                Err.Description
        
End Sub

Friend Sub CancelMarketData(ByVal tickerid As Long)
On Error GoTo Err
cancelMarketDataEx tickerid, True
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "CancelMarketData", _
                Err.Description
End Sub

Friend Sub CancelMarketDepth(ByVal tickerid As Long)
On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

Const version = 1

With mTickerTable(tickerid)
    If Not .inUse Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "IBTWSSP.TWSAPI::cancelMarketDepth", _
                                    "Invalid ticker id " & tickerid
    If Not .receivingMarketDepth Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
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
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "CancelMarketDepth", _
                Err.Description

End Sub

Friend Sub cancelNewsBulletins()

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

Const version = 1

addData CANCEL_NEWS_BULLETINS
addData version
send
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "cancelNewsBulletins", _
                Err.Description
End Sub

Friend Sub cancelOrder(ByVal orderId As Long)

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::cancelOrder", _
                "Not connected"
End If

Const version = 1
addData CANCEL_ORDER
addData version
addData orderId
send
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "cancelOrder", _
                Err.Description

End Sub

Friend Sub Connect()

On Error GoTo Err
mDisconnectedByApplication = False

If mConnectionState <> ConnectionStates.ConnNotConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::connect", _
                "Already connected"
End If

createsocket


logMessage "Connecting to TWS: " & connectionString, LogLevelNormal
mSocket.Connect mServer, mPort

Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "Connect", _
                Err.Description
End Sub

Friend Sub disconnect( _
                ByVal reason As String, _
                ByVal forced As Boolean)
On Error GoTo Err

If mConnectionState = ConnectionStates.ConnNotConnected Then Exit Sub

mDisconnectedByApplication = True
If Not mConnectionTimer Is Nothing Then mConnectionTimer.StopTimer
mConnectionState = ConnectionStates.ConnNotConnected
releaseAllTickerTableEntries
ReleaseSocket
If forced Then RaiseEvent Disconnected(reason)
logMessage "IBTWSSP:TwsAPI:disconnect: Disconnected from: " & connectionString & ": " & reason, LogLevelNormal
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "disconnect", _
                Err.Description
End Sub

Friend Sub placeOrder( _
                ByVal OrderSubmitter As OrderSubmitter, _
                ByVal pOrder As TradeBuildSP.IOrder, _
                ByVal transmit As Boolean, _
                ByVal parentId As String, _
                ByVal ocaGroup As String, _
                ByVal routingHint As String)

Dim mapEntry As OrderIdMapEntry

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::placeOrder", _
                "Already connected"
End If

If pOrder.brokerId = "" Then
    pOrder.brokerId = CStr(mNextOrderID)
    mNextOrderID = mNextOrderID + 1
    Set mapEntry.order = pOrder
    Set mapEntry.OrderSubmitter = OrderSubmitter
    mOrderIDMap.add mapEntry, generateTwsOrderKey(pOrder.brokerId, mClientID)
End If

sendOrder pOrder, transmit, parentId, ocaGroup, routingHint
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "placeOrder", _
                Err.Description

End Sub
                    
Friend Function ReadHistoricalBar( _
                        ByVal requestHandle As Long, _
                        ByRef pBar As Bar, _
                        ByRef barTimePeriod As TimePeriod) As Boolean

Dim Bar As HistoricBar

On Error GoTo Err
With mHistoricalDataTable(requestHandle)
    If .nextBarIndex = -1 Then
        releaseHistoricalDataTableEntry requestHandle
    Else
        ReadHistoricalBar = True
        Bar = .bars.itemAt(.nextBarIndex)
        Set barTimePeriod = .barSpecifier.barTimePeriod
        With Bar
            Set pBar = CreateBar(.timestamp, _
                            .openValue, _
                            .highValue, _
                            .lowValue, _
                            .closeValue, _
                            .volume, _
                            .tickVolume)
        End With
        .nextBarIndex = .bars.nextIndex(.nextBarIndex)
    End If
End With
Exit Function
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "ReadHistoricalBar", _
                Err.Description

End Function

Friend Sub ReplaceFA(ByVal DataType As FADataTypes, ByVal xml As String)


On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::ReplaceFA", _
                "Already connected"
End If

Const version = 1

If mServerVersion < 13 Then
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
                "IBTWSSP.TWSAPI::ReplaceFA", _
                "Server version does not support ReplaceFA"
End If

addData REPLACE_FA
addData version
addData DataType
addData xml

send
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "ReplaceFA", _
                Err.Description
End Sub

Friend Sub requestAccountUpdates(ByVal subscribe As Boolean, ByVal acctCode As String)

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
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
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "requestAccountUpdates", _
                Err.Description
End Sub

Friend Sub requestAllOpenOrders()

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::requestAllOpenOrders", _
                "Not connected"
End If
                            

Const version = 1

addData REQ_ALL_OPEN_ORDERS
addData version
send
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "requestAllOpenOrders", _
                Err.Description
End Sub

Friend Sub requestAutoOpenOrders(ByVal autoBind As Boolean)

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::requestAutoOpenOrders", _
                "Not connected"
End If

Const version = 1

addData REQ_AUTO_OPEN_ORDERS
addData version
addBoolean autoBind
send
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "requestAutoOpenOrders", _
                Err.Description
End Sub

Friend Sub RequestContract( _
                ByVal contractSpec As ContractSpecifier, _
                ByVal pContractsBuilder As ContractsBuilder, _
                ByVal requester As ContractDetailsRequester)

Dim requestQueueEntry As TWSContractDetailsRequestQueueEntry

On Error GoTo Err
Set requestQueueEntry.contractSpec = contractSpec
Set requestQueueEntry.tradeBuildContractsBuilder = pContractsBuilder
Set requestQueueEntry.requester = requester

Select Case contractSpec.sectype
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
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestContract", _
                Err.Description
End Sub

Friend Sub RequestExecutions(ByVal filter As ExecutionFilter)
Dim theFilter As ExecutionFilter

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
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
        addData IIf(CLng(.time) <> 0, Format(.time, "yyyymmdd-hh\:nn\:ss"), "")
        addData .symbol
        addData secTypeToString(.sectype)
        addData .exchange
        addData orderActionToString(.Action)
    End With
End If

send
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestExecutions", _
                Err.Description
End Sub

Friend Sub requestFA(ByVal DataType As FADataTypes)

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::requestFA", _
                "Not connected"
End If

Const version = 1

If mServerVersion < 13 Then
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
                "IBTWSSP.TWSAPI::requestFA", _
                "Server version does not support requestFA"
End If

addData REQ_FA
addData version
addData DataType

send
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "requestFA", _
                Err.Description
End Sub

Friend Sub requestManagedAccounts()

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::requestManagedAccounts", _
                "Not connected"
End If

Const version = 1

addData REQ_MANAGED_ACCTS
addData version
send
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "requestManagedAccounts", _
                Err.Description
End Sub

Friend Function RequestMarketData(ByVal pContract As Contract, _
                            ByVal dataReader As RealtimeDataReader, _
                            ByVal dataConsumer As IStreamingDataConsumer, _
                            Optional ByVal includeMarketDepth As Boolean = False) As Long

Dim lTickerID As Long

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
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
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestMarketData", _
                Err.Description

End Function

Friend Sub RequestMarketDepth(ByVal pTickerID As Long)
                            
On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::RequestMarketDepth", _
                "Not connected"
End If

With mTickerTable(pTickerID)
    If .Contract Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, "IBTWSSP.TWSAPI::requestMarketDepth", "Invalid ticker id: " & pTickerID
    .receivingMarketDepth = True
    reqMarketDepth pTickerID, .Contract.specifier
End With
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestMarketDepth", _
                Err.Description

End Sub

Friend Function RequestHistoricalData( _
                barSpecifier As BarDataSpecifier, _
                ByVal historicalDataReader As HistDataReader) As Long

Dim requestID As Long
Dim i As Long
Dim ignoreTimestamps As Boolean

On Error GoTo Err
requestID = -1

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::RequestHistoricalData", _
                "Not connected"
End If

If mServerVersion < 16 Then
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
                "IBTWSSP.TWSAPI::RequestHistoricalData", _
                "Server version does not support historical data retrieval"
End If

' now adjust the bar specifier to indicate the number of bars of
' appropriate length supported by IB
Select Case barSpecifier.barTimePeriod.units
Case TimePeriodUnits.TimePeriodDay
    ignoreTimestamps = True
    If barSpecifier.barTimePeriod.length > 1 Then
        barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * barSpecifier.barTimePeriod.length
        barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.units)
    End If
Case TimePeriodUnits.TimePeriodHour
    If barSpecifier.barTimePeriod.length > 1 Then
        barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * barSpecifier.barTimePeriod.length
        barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.units)
    End If
Case TimePeriodUnits.TimePeriodMinute
    If barSpecifier.barTimePeriod.length Mod 30 = 0 Then
        barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * (barSpecifier.barTimePeriod.length / 30)
        barSpecifier.barTimePeriod = GetTimePeriod(30, barSpecifier.barTimePeriod.units)
    ElseIf barSpecifier.barTimePeriod.length Mod 15 = 0 Then
        barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * (barSpecifier.barTimePeriod.length / 15)
        barSpecifier.barTimePeriod = GetTimePeriod(15, barSpecifier.barTimePeriod.units)
    ElseIf barSpecifier.barTimePeriod.length Mod 5 = 0 Then
        barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * (barSpecifier.barTimePeriod.length / 5)
        barSpecifier.barTimePeriod = GetTimePeriod(5, barSpecifier.barTimePeriod.units)
    ElseIf barSpecifier.barTimePeriod.length Mod 2 = 0 Then
        barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * (barSpecifier.barTimePeriod.length / 2)
        barSpecifier.barTimePeriod = GetTimePeriod(2, barSpecifier.barTimePeriod.units)
    ElseIf barSpecifier.barTimePeriod.length > 1 Then
        barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * barSpecifier.barTimePeriod.length
        barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.units)
    End If
Case TimePeriodUnits.TimePeriodMonth
    ignoreTimestamps = True
    barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * barSpecifier.barTimePeriod.length
    barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.units)
Case TimePeriodUnits.TimePeriodSecond
    If barSpecifier.barTimePeriod.length Mod 30 = 0 Then
        barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * (barSpecifier.barTimePeriod.length / 30)
        barSpecifier.barTimePeriod = GetTimePeriod(30, barSpecifier.barTimePeriod.units)
    ElseIf barSpecifier.barTimePeriod.length Mod 15 = 0 Then
        barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * (barSpecifier.barTimePeriod.length / 15)
        barSpecifier.barTimePeriod = GetTimePeriod(15, barSpecifier.barTimePeriod.units)
    ElseIf barSpecifier.barTimePeriod.length Mod 5 = 0 Then
        barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * (barSpecifier.barTimePeriod.length / 5)
        barSpecifier.barTimePeriod = GetTimePeriod(5, barSpecifier.barTimePeriod.units)
    ElseIf barSpecifier.barTimePeriod.length > 1 Then
        barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * barSpecifier.barTimePeriod.length
        barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.units)
    End If
Case TimePeriodUnits.TimePeriodWeek
    ignoreTimestamps = True
    barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * barSpecifier.barTimePeriod.length
    barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.units)
Case TimePeriodUnits.TimePeriodYear
    ignoreTimestamps = True
    barSpecifier.maxNumberOfBars = barSpecifier.maxNumberOfBars * barSpecifier.barTimePeriod.length * NumMonthsInYear
    barSpecifier.barTimePeriod = GetTimePeriod(1, TimePeriodMonth)
Case TimePeriodUnits.TimePeriodVolume, TimePeriodUnits.TimePeriodTickMovement
    ' nothing to do
Case Else
    historicalDataReader.historicalDataRequestInvalid barSpecifier.maxNumberOfBars & "specified timeframe not supported"
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
    ReDim Preserve mHistoricalDataTable(2 * (UBound(mHistoricalDataTable) + 1) - 1) As HistoricalDataRequestTableEntry
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
mHistoricalDataTable(requestID).bars.initialCapacity = IIf(barSpecifier.maxNumberOfBars <> 0, barSpecifier.maxNumberOfBars, 50)

generateTwsHistoricalDataRequest requestID
RequestHistoricalData = requestID
Exit Function
Err:

If Err.number = ErrHistGenerationFailed Then
    releaseHistoricalDataTableEntry requestID
    RequestHistoricalData = -1
    Exit Function
End If
    
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestHistoricalData", _
                Err.Description
End Function

Friend Sub RequestNewsBulletins(allMsgs As Boolean)

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::RequestNewsBulletins", _
                "Not connected"
End If

Const version = 1

addData REQ_NEWS_BULLETINS
addData version
addData allMsgs
send
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestNewsBulletins", _
                Err.Description
End Sub

Friend Sub RequestOpenOrders()

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "IBTWSSP.TWSAPI::RequestOpenOrders", _
                "Not connected"
End If

Const version = 1

addData REQ_OPEN_ORDERS
addData version
send
Exit Sub
Err:
handleFatalError Err.number, _
                "IBTWSSP" & "." & "TWSAPI" & "::" & "RequestOpenOrders", _
                Err.Description
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub addBoolean(ByVal val As Boolean)
addData IIf(val, 1, 0)
End Sub

Private Sub addByte(ByVal value As Integer)
If mIndexOut > UBound(mBufferOut) Then
    ReDim Preserve mBufferOut(2 * (UBound(mBufferOut) + 1) - 1) As Byte
End If
mBufferOut(mIndexOut) = value
mIndexOut = mIndexOut + 1
End Sub

Private Sub addData(ByVal data As String)
Dim i As Long

For i = 1 To Len(data)
    addByte Asc(Mid$(data, i, 1))
Next
addByte 0
End Sub

Private Sub addDouble(ByVal data As String)
addData DoubleToString(data)
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
    ReDim Preserve mTickerTable(2 * (UBound(mTickerTable) + 1) - 1) As TickerTableEntry
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
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
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
Static msgID As TWSSocketInMsgTypes
Static version As Long
Static startIndex As Long

Static prevMsgID As TWSSocketInMsgTypes
Static prevVersion As Long
Static prevStartIndex As Long

Static retryCount As Long
Static eventCount As Long
Static mElapsedTimer As ElapsedTimer
Dim s As String

If mElapsedTimer Is Nothing Then
    Set mElapsedTimer = New ElapsedTimer
    mElapsedTimer.StartTiming
End If

On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    ' first data received is plain server version number
    If mCheckPointNumber = 0 Then
        createMessageBuilder
        mServerVersion = getLong("Server version")
        checkpointBuffer
        mCheckPointNumber = 100
        'Debug.Print "Set checkpoint number=" & 100
    End If
    
    ' in some circumstances, some versions of TWS send malformed
    ' ERR_MSG messages before sending the version number, so we
    ' need to skip these if they occur. (They are malformed in that
    ' they don't include the id or error code.)
    
    Do While mServerVersion = TWSSocketInMsgTypes.ERR_MSG
        If mCheckPointNumber = 100 Then
            getLong ("Invalid version")
            checkpointBuffer
            mCheckPointNumber = 101
        End If
        If mCheckPointNumber = 101 Then
            getString "Invalid errormsg"
            checkpointBuffer
            mCheckPointNumber = 102
        End If
        
        logSocketMessage
        createMessageBuilder
        
        If mCheckPointNumber = 102 Then
            mServerVersion = getLong("Server version")
            checkpointBuffer
            mCheckPointNumber = 100
        End If
    Loop
    
    logSocketMessage
    createMessageBuilder
    
    If mCheckPointNumber = 100 Then
        If mServerVersion >= 20 Then getString "TWS time"  ' TWS time
        checkpointBuffer
        mCheckPointNumber = 200
        'Debug.Print "Set checkpoint number=" & 200
    End If
    
    logSocketMessage
    
    mConnectionState = ConnectionStates.ConnConnected
    logMessage "Connection to TWS successful", LogLevelNormal
    
    ' set the log level in TWS
    setTWSLogLevel mTWSLogLevel
    
    ' now restart any tickers that were previously in operation
    reEstablishMarketData True
    
    ' and re-request any market depth
    resetMarketDepth True
    
    ' check for pending contract details requests
    submitNextContractDetailsRequest
        
    
    RaiseEvent Connected
    
    mCheckPointNumber = 1000
    'Debug.Print "Set checkpoint number=" & 1000
End If

Do While mInputIndex < mIndexIn
    
    If mCheckPointNumber = 1000 Then
        Dim tmpStartIndex As Long
        tmpStartIndex = startIndex
        startIndex = mInputIndex
        prevMsgID = msgID
        
        createMessageBuilder
    
        msgID = getLong("Msg id")
        prevStartIndex = tmpStartIndex
        checkpointBuffer
        mCheckPointNumber = 1100
        'Debug.Print "Set checkpoint number=" & 1100 & ": msgID=" & msgID
    End If
    If mCheckPointNumber = 1100 Then
        prevVersion = version
        version = getLong("Version")
        checkpointBuffer
        mCheckPointNumber = 1200
        'Debug.Print "Set checkpoint number=" & 1200 & ": version=" & version
    End If
    
    mPerformanceElapsedTimer.StartTiming
    
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
    Case TWSSocketInMsgTypes.TICK_STRING
        getTickString version
    Case Else
        s = "Invalid message code from TWS: " & vbCrLf & _
            "msgID=" & msgID & _
            " version=" & version & _
            " startindex=" & startIndex & vbCrLf & _
            "prevMsgID=" & prevMsgID & _
            " prevVersion=" & prevVersion & _
            " prevStartindex=" & prevStartIndex & vbCrLf & _
            "mIndexIn=" & mIndexIn & _
            " mInputIndex=" & mInputIndex & vbCrLf & _
            "Buffer contents: " & vbCrLf & _
            formatBuffer
        
        Debug.Print s
        logMessage s, LogLevelSevere
        gReleaseTWSAPIInstance Me, True
        Exit Sub
    End Select
    
    retryCount = 0
    
    Dim et As Single
    et = mPerformanceElapsedTimer.ElapsedTimeMicroseconds
    With mPerformanceStats(msgID)
        .lastSecondCount = .lastSecondCount + 1
        .lastSecondTime = .lastSecondTime + et
        If et > .longestTime Then .longestTime = et
        If et < .shortestTime Or .shortestTime = 0 Then .shortestTime = et
    End With
    
    eventCount = eventCount + 1
    
    If mElapsedTimer.ElapsedTimeMicroseconds >= 10000000 Then
        logMessage "Event rate per second = " & Format(eventCount / 10, "0.0"), LogLevelDetail
        Debug.Print "Event rate per second = " & Format(eventCount / 10, "0.0")
        eventCount = 0
        mElapsedTimer.StartTiming
    End If
    
    checkpointBuffer
    mCheckPointNumber = 1000
    'Debug.Print "Set checkpoint number=" & 1000 & "; minputindex=" & mInputIndex
Loop

Exit Sub

Err:
Dim errNum As Long
Dim errDesc As String

errNum = Err.number
errDesc = Err.Description

If errNum = DataIncomplete Then
    rollbackBuffer
    retryCount = retryCount + 1
    Debug.Print "Data incomplete count: " & retryCount & "; mCheckPointNumber: " & mCheckPointNumber
    Exit Sub
End If

s = "msgID=" & msgID & _
    " version=" & version & _
    " startindex=" & startIndex & vbCrLf & _
    "prevMsgID=" & prevMsgID & _
    " prevVersion=" & prevVersion & _
    " prevStartindex=" & prevStartIndex & vbCrLf & _
    "mIndexIn=" & mIndexIn & _
    " mInputIndex=" & mInputIndex & vbCrLf & _
    "Buffer contents: " & vbCrLf & _
    formatBuffer

Debug.Print "Error " & errNum & ": " & errDesc & vbCrLf & s

logMessage "IBTWSSP:TwsAPI:checkMessages Error (" & errNum & "): " & errDesc, LogLevelSevere
logMessage s, LogLevelDetail
mCommonServiceConsumer.ServiceProviderError errNum, errDesc, -1
gReleaseTWSAPIInstance Me, True
End Sub

Private Sub checkpointBuffer()
mInputIndexChk = mInputIndex
End Sub

Private Sub createLoggers()

If gLogger.isLoggable(LogLevelMediumDetail) Then
    If mPerformanceTimerSecond Is Nothing Then
        Set mPerformanceTimerSecond = CreateIntervalTimer(CDate(Int((Now + 2 * OneSecond) / OneSecond) * OneSecond), _
                                                ExpiryTimeUnitDateTime, _
                                                1000)
        Set mPerformanceTimerPeriod = CreateIntervalTimer(CDate(Int((Now + OneMinute) / OneMinute) * OneMinute), _
                                                ExpiryTimeUnitDateTime, _
                                                60000)
        mPerformanceTimerSecond.StartTimer
        mPerformanceTimerPeriod.StartTimer
        
        Set mPerformanceLogger = GetLogger("tradebuild.serviceprovider.ibtwssp.performance")
    End If
End If

Set mSocketLogger = GetLogger("tradebuild.serviceprovider.ibtwssp.socket")
End Sub

Private Sub createMessageBuilder()
If mMessageBuilder Is Nothing And _
    mSocketLogger.isLoggable(LogLevelHighDetail) _
Then
    Set mMessageBuilder = CreateStringBuilder
End If
End Sub

Private Sub createsocket()
If mSocket Is Nothing Then
    Set mWinsockForm = New WinsockForm
    Set mSocket = mWinsockForm.Winsock1
End If
End Sub

Private Function expiryToDate( _
                ByVal expiry As String) As Date
If Len(expiry) = 8 Then
    expiryToDate = CDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2))
ElseIf Len(expiry) = 6 Then
    expiryToDate = CDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & "01")
End If
End Function

Private Function formatBuffer() As String
Dim s As StringBuilder
Dim i As Long
Dim j As Long

Set s = CreateStringBuilder
Do While i < mIndexIn
    s.append Format(i, "0000  ")
    For j = i To i + 49
        If j = mIndexIn Then Exit For
        s.append IIf(mBufferIn(j) <> 0, Chr$(mBufferIn(j)), "_")
    Next
    i = i + 50
    If j < mIndexIn Then s.appendLine ""
Loop
formatBuffer = s.toString
End Function

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

Set TwsRequest.Contract = mHistoricalDataTable(index).barSpecifier.Contract

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
    TwsRequest.endDateTime = Format(ConvertDateUTCToLocal(ConvertDateTzToUTC(mHistoricalDataTable(index).barSpecifier.ToDate, _
                                                                            mHistoricalDataTable(index).barSpecifier.Contract.TimeZone)), _
                                    "yyyymmdd hh\:nn\:ss")
    
    With mHistoricalDataTable(index).barSpecifier
        
        Select Case .barTimePeriod.units
        Case TimePeriodUnits.TimePeriodDay
            TwsRequest.barSizeSetting = "1 day"
            If .maxNumberOfBars >= NumDaysInYear Then
                If DateAdd("d", 1, .ToDate) > Now Then
                    TwsRequest.duration = "1 Y"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 60) & " D"
                End If
            ElseIf .maxNumberOfBars <= 60 Then
                TwsRequest.duration = permittedHistDays(.ToDate, .maxNumberOfBars) & " D"
            ElseIf .maxNumberOfBars <= 52 * NumDaysInWeek Then
                TwsRequest.duration = permittedHistWeeks(.ToDate, Int(.maxNumberOfBars / NumDaysInWeek)) & " W"
            Else
                TwsRequest.duration = permittedHistDays(.ToDate, 60) & " D"
            End If
        Case TimePeriodUnits.TimePeriodHour
            TwsRequest.barSizeSetting = "1 hour"
            
            numBarsInsession = mHistoricalDataTable(index).Session.NumberOfBarsInSession( _
                                                                    .barTimePeriod)
            
            If .maxNumberOfBars <= 24 Then
                TwsRequest.duration = permittedHistSeconds(.ToDate, .maxNumberOfBars * 3600) & " S"
            ElseIf -Int(-.maxNumberOfBars / numBarsInsession) <= 34 Then
                TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.maxNumberOfBars / numBarsInsession)) & " D"
            Else
                TwsRequest.duration = permittedHistDays(.ToDate, 34) & " D"
            End If
        Case TimePeriodUnits.TimePeriodMinute
            
            numBarsInsession = mHistoricalDataTable(index).Session.NumberOfBarsInSession( _
                                                                    .barTimePeriod)
            
            Select Case .barTimePeriod.length
            Case 1
                TwsRequest.barSizeSetting = "1 min"
                If .maxNumberOfBars <= 1440 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .maxNumberOfBars * 60) & " S"
                ElseIf -Int(-.maxNumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.maxNumberOfBars / numBarsInsession) <= 6 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.maxNumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 6) & " D"
                End If
            Case 2
                TwsRequest.barSizeSetting = "2 mins"
                If .maxNumberOfBars <= 720 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .maxNumberOfBars * 120) & " S"
                ElseIf -Int(-.maxNumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.maxNumberOfBars / numBarsInsession) <= 6 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.maxNumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 6) & " D"
                End If
            Case 5
                TwsRequest.barSizeSetting = "5 mins"
                If .maxNumberOfBars <= 288 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .maxNumberOfBars * 300) & " S"
                ElseIf -Int(-.maxNumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.maxNumberOfBars / numBarsInsession) <= 6 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.maxNumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 6) & " D"
                End If
            Case 15
                TwsRequest.barSizeSetting = "15 mins"
                If .maxNumberOfBars <= 96 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .maxNumberOfBars * 900) & " S"
                ElseIf -Int(-.maxNumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.maxNumberOfBars / numBarsInsession) <= 20 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.maxNumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 20) & " D"
                End If
            Case 30
                TwsRequest.barSizeSetting = "30 mins"
                If .maxNumberOfBars <= 48 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .maxNumberOfBars * 1800) & " S"
                ElseIf -Int(-.maxNumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.maxNumberOfBars / numBarsInsession) <= 34 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.maxNumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 34) & " D"
                End If
            End Select
        Case TimePeriodUnits.TimePeriodMonth
            TwsRequest.barSizeSetting = "1 day"
            If .maxNumberOfBars <= 12 Then
                TwsRequest.duration = permittedHistMonths(.ToDate, .maxNumberOfBars) & " M"
            Else
                If DateAdd("d", 1, .ToDate) > Now Then
                    TwsRequest.duration = "12 M"
                Else
                    TwsRequest.duration = permittedHistMonths(.ToDate, .maxNumberOfBars) & " M"
                End If
            End If
        Case TimePeriodUnits.TimePeriodSecond
            
            numBarsInsession = mHistoricalDataTable(index).Session.NumberOfBarsInSession( _
                                                                    .barTimePeriod)
            
            Select Case .barTimePeriod.length
            Case 1
                
                If .maxNumberOfBars <= 2000 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .maxNumberOfBars) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 2000) & " S"
                End If
            Case 5
                TwsRequest.barSizeSetting = "5 secs"
                If .maxNumberOfBars <= 2000 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .maxNumberOfBars * 5) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 10000) & " S"
                End If
            Case 15
                TwsRequest.barSizeSetting = "15 secs"
                If .maxNumberOfBars <= 2000 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .maxNumberOfBars * 15) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 30000) & "S"
                End If
            Case 30
                TwsRequest.barSizeSetting = "30 secs"
                If .maxNumberOfBars <= 2880 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .maxNumberOfBars * 30) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                End If
            End Select
        Case TimePeriodUnits.TimePeriodWeek
            TwsRequest.barSizeSetting = "1 day"
            If .maxNumberOfBars <= 52 Then
                TwsRequest.duration = permittedHistWeeks(.ToDate, .maxNumberOfBars) & " W"
            Else
                If DateAdd("d", 1, .ToDate) > Now Then
                    TwsRequest.duration = "52 W"
                Else
                    TwsRequest.duration = permittedHistWeeks(.ToDate, .maxNumberOfBars) & " W"
                End If
            End If
        Case TimePeriodUnits.TimePeriodVolume, TimePeriodUnits.TimePeriodTickMovement
            TwsRequest.barSizeSetting = "15 secs"
            TwsRequest.duration = "30000 S"
        Case Else
            Debug.Assert False
        End Select
    End With
End If

Select Case mHistoricalDataTable(index).barSpecifier.barType
Case BarTypes.BarTypeAsk
    TwsRequest.whatToShow = TWSWhatToShowAsk
Case BarTypes.BarTypeBid
    TwsRequest.whatToShow = TWSWhatToShowBid
Case BarTypes.BarTypeTrade
    If mHistoricalDataTable(index).barSpecifier.Contract.specifier.sectype = SecTypeCash Then
        TwsRequest.whatToShow = TWSWhatToShowMidpoint
    Else
        TwsRequest.whatToShow = TWSWhatToShowTrades
    End If
Case Else
End Select

mHistoricalDataTable(index).requestKeys(reqNum - 1) = queueTWSHistoricalDataRequest(TwsRequest)

End Sub

Private Function generateTwsOrderKey( _
                ByVal twsOrderId As Long, _
                ByVal clientID As Long) As String
generateTwsOrderKey = clientID & "/" & twsOrderId
End Function

Private Function getBoolean( _
                ByRef fieldName As String) As Boolean
Dim s As String
s = getString(fieldName)
If s = "1" Then
    getBoolean = True
Else
    getBoolean = False
End If
End Function

Private Function getDouble( _
                ByRef fieldName As String) As Double
Dim s As String
s = getString(fieldName)
If s = "" Then
    getDouble = 0#
Else
    getDouble = DoubleFromString(s)
End If
End Function

Private Function getInteger( _
                ByRef fieldName As String) As Integer
Dim s As String
s = getString(fieldName)
If s = "" Then
    getInteger = 0
Else
    getInteger = CInt(s)
End If
End Function

Private Function getLong( _
                ByRef fieldName As String) As Long
Dim s As String
s = getString(fieldName)
If s = "" Then
    getLong = 0&
Else
    getLong = CLng(s)
End If
End Function

Private Function getSingle( _
                ByRef fieldName As String) As Single
Dim s As String
s = getString(fieldName)
If s = "" Then
    getSingle = 0!
Else
    getSingle = CSng(s)
End If
End Function

Private Function getString( _
                ByRef fieldName As String) As String
Dim s As String
Dim i As Long

For i = mInputIndex To mIndexIn - 1
    If mBufferIn(i) = 0 Then
        ' we've found our next string
        'Debug.Print "Got string: mInputIndex=" & mInputIndex & " value=" & s
        mInputIndex = i + 1
        getString = s
        If Not mMessageBuilder Is Nothing Then
            mMessageBuilder.append fieldName
            mMessageBuilder.append "="
            mMessageBuilder.append s
            mMessageBuilder.append "; "
        End If
        Exit Function
    End If
    s = s & Chr$(mBufferIn(i))
Next

' we've got to the end of the current data, so raise an error to abort
' this attempt to assemble a message
'Debug.Print "Failed to get string"
Err.Raise DataIncomplete
End Function

Private Sub getAccountTime(ByVal version As Long)
Dim accountTime As String: accountTime = getString("Account time")

logSocketMessage

End Sub

Private Sub getAccountValue(ByVal version As Long)
Dim key As String: key = getString("Key")
Dim val As String: val = getString("Value")
Dim cur As String: cur = getString("Currency")
Dim accountName As String: If version >= 2 Then accountName = getString("Account name")

logSocketMessage
End Sub

Private Sub getBondContractData(ByVal version As Long)
' not currently supported
getString "symbol"
getString "sectype"
getString "cusip"
getDouble "coupon"
getString "maturity"
getString "issueDate"
getString "ratings"
getString "bondType"
getString "couponType"
getBoolean "convertible"
getBoolean "callable"
getBoolean "putable"
getString "descAppend"
getString "exchange"
getString "currency"
getString "marketName"
getString "tradingClass"
getLong "conid"
getDouble "minTick"
getString "OrderTypes"
getString "validExchanges"

logSocketMessage

mContractDetailsRequest.requester.contractSpecifierInvalid mContractDetailsRequest, _
                                                    "bonds not yet supported"
                                                    

End Sub

Private Sub getContractData(ByVal version As Long)
Dim lContractWrapper As ContractWrapper
Dim lContractSpec As ContractSpecifier

Dim symbol As String
Dim localSymbol As String
Dim sectypeStr As String ' keep the socket string for logging
Dim sectype As SecurityTypes
Dim expiry As String
Dim strike As Double
Dim optRightStr As String ' keep the socket string for logging
Dim optRight As OptionRights
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

Dim ignore As Boolean

Set lContractWrapper = New ContractWrapper

symbol = getString("Symbol")
sectypeStr = getString("Sec type")
sectype = secTypeFromString(sectypeStr)
expiry = getString("Expiry")
strike = getDouble("Strike")
optRightStr = getString("Right")
optRight = optionRightFromString(optRightStr)
exchange = getString("Exchange")
currencyCode = getString("Currency")
localSymbol = getString("Local symbol")
marketName = getString("Market name")
tradingClass = getString("Trading class")
lContractWrapper.contractId = getLong("Contract id")
minimumTick = getDouble("Minimum tick")
multiplier = getLong("Multiplier")
If multiplier = 0 Then multiplier = 1
    
If IsValidExchangeCode(exchange) Then

    Set lContractSpec = CreateContractSpecifier(localSymbol, _
                                                symbol, _
                                                exchange, _
                                                sectype, _
                                                currencyCode, _
                                                expiry, _
                                                strike, _
                                                optRight)
        
    Set lContractBuilder = CreateContractBuilder(lContractSpec)
    lContractWrapper.tradeBuildContract = lContractBuilder.Contract
Else
    logMessage "Unknown exchange code " & exchange & " received from TWS", LogLevelNormal
    ignore = True
End If

OrderTypes = getString("Order types")
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
        Case "ALGO", _
            "AVGCOST", _
            "AUC", _
            "BASKET", _
            "COND", _
            "CONDORDER", _
            "CONSCOST", _
            "DEACTEOD", _
            "FOK", _
            "HPENNY", _
            "LARGE", _
            "LTH", _
            "NONALGO", _
            "OCA", _
            "OPG", _
            "OPGREROUT", _
            "PAON", _
            "PEGMID", _
            "PREOPGRTH", _
            "RELSTK", _
            "SCALE", _
            "SCALERST", _
            "SOI", _
            "TIMEPRIO", _
            "VOLAT"
            ' not sure what these signify
        Case Else
            If orderTypesStr(i) <> "" Then
                logMessage "Unknown order type " & orderTypesStr(i) & " received from TWS", LogLevelNormal
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

validExchanges = getString("Valid exchanges")
lValidExchanges = Split(validExchanges, ",")
    
If version >= 2 Then
    lContractWrapper.priceMagnifier = getLong("Price magnifier")
Else
    If currencyCode = "GBP" Then
        lContractWrapper.priceMagnifier = 100
    Else
        lContractWrapper.priceMagnifier = 1
    End If
End If

logSocketMessage

If Not ignore Then
    Dim expiryDate As Date
    If expiry <> "" Then
        expiryDate = CDate(Left$(expiry, 4) & "/" & _
                                            Mid$(expiry, 5, 2) & "/" & _
                                            Right$(expiry, 2))
    End If
    
    lContractBuilder.expiryDate = expiryDate
    
    lContractBuilder.Description = symbol & _
                                " (" & SecTypeToShortString(sectype) & ")" & _
                                IIf(expiryDate <> 0, Format(expiryDate, " dd mmm yy"), "") & _
                                IIf(optRight <> OptNone, " " & optionRightToString(optRight), "") & _
                                IIf(strike <> 0#, " " & strike, "")

    lContractBuilder.multiplier = multiplier / lContractWrapper.priceMagnifier
    lContractBuilder.TickSize = minimumTick * lContractWrapper.priceMagnifier

    mContractDetailsRequest.contractWrappers.add lContractWrapper
End If

If Not mFirstContractDetailsReceived Then
    mFirstContractDetailsReceived = True
    
    ' now send a reqContract that will give a 'no such security
    ' exists' error. This error will act as a delimiter for all
    ' the contract details returned by the orginal request. Note that
    ' tests indicate that contract details for different reqContracts
    ' are not interleaved
    
    addData REQ_CONTRACT_DATA
    addData 1   ' version
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

End Sub

Private Sub getErrorMsg(ByVal version As Long)
Dim id As Long
Dim errorCode As Long
Dim errorMsg As String
Dim execFilter As ExecutionFilter
Dim index As Long
Dim mapEntry As OrderIdMapEntry
Dim failpoint As Long

On Error GoTo Err

failpoint = 100

id = getLong("Id")
Debug.Print "Error id: " & id

failpoint = 200

errorCode = getLong("Error code")
Debug.Print "Error code: " & errorCode

failpoint = 300

errorMsg = getString("Error msg")
Debug.Print "Error msg: " & errorMsg

logSocketMessage

Select Case errorCode
Case 101
    ' max tickers reached
    failpoint = 350
    mTickerTable(id Mod IncarnationIncrement).dataReader.marketDataRequestFailed "Maximum number of tickers reached"
Case 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, _
    121, 125, 126, 127, 128, 129, 130, 132, 137, 140, 141, 144, _
    152, 153, 154, 157, 160, 201, 312, 313, 314, 315, 325, 328, _
    334, 337, 339, 343
    
    failpoint = 400

    mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
    mapEntry.OrderSubmitter.updateStatus mapEntry.order.tradeBuildId, _
                                        OrderStatuses.OrderStatusRejected
    mapEntry.OrderSubmitter.invalidOrder mapEntry.order.tradeBuildId, errorMsg
Case 162    ' historical data request problem
    failpoint = 500

    index = (id - HistoricalDataTickerBase) Mod HistoricalDataTickerIncrement

    If mHistoricalDataTable(index).bars.size <> 0 Then
        Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
        mHistoricalDataTable(index).dataRetrievedFromTWS = True
        mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
        mHistoricalDataTable(index).historicalDataReader.historicalDataAvailable
    Else
        logMessage "Error " & errorCode & " from TWS: " & errorMsg, LogLevelNormal
        mHistoricalDataTable(index).historicalDataReader.historicalDataRequestFailed errorMsg
        releaseHistoricalDataTableEntry index
    End If
    
    processNextHistDataRequest
Case 165
    failpoint = 600

    logMessage "Connected to IB Historical Market Data Service", LogLevelNormal
Case 200    ' security definition not known
    Select Case getIdType(id)
    Case IdTypeNone
        ' this must relate to a RequestContract
        failpoint = 750

        processContractError errorMsg
    Case IdTypeRealtimeData
        ' this must relate to a reqMktData - this can occur even though
        ' we make sure the contract is valid before starting the ticker because
        ' for example the contract may have expired
        
        failpoint = 700

        mTickerTable(id Mod IncarnationIncrement).dataReader.marketDataRequestFailed errorMsg
        releaseTickerTableEntry (id Mod IncarnationIncrement)
    Case IdTypeMarketDepth
        mTickerTable(id - MarketDepthTickerBase).dataReader.marketDepthRequestFailed errorMsg
    Case IdTypeHistoricalData
        index = (id - HistoricalDataTickerBase) Mod HistoricalDataTickerIncrement
    
        logMessage "Error " & errorCode & " from TWS: " & errorMsg, LogLevelNormal
        mHistoricalDataTable(index).historicalDataReader.historicalDataRequestFailed errorMsg
        releaseHistoricalDataTableEntry index
        
        processNextHistDataRequest
    Case IdTypeOrder
        mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
        mapEntry.OrderSubmitter.updateStatus mapEntry.order.tradeBuildId, _
                                            OrderStatuses.OrderStatusRejected
        mapEntry.OrderSubmitter.invalidOrder mapEntry.order.tradeBuildId, errorMsg
    End Select
    
Case 202
    failpoint = 800

    mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
    mapEntry.OrderSubmitter.updateStatus mapEntry.order.tradeBuildId, _
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

    Select Case getIdType(id)
    Case IdTypeNone
    Case IdTypeRealtimeData
    Case IdTypeMarketDepth
    Case IdTypeHistoricalData
        index = (id - HistoricalDataTickerBase) Mod HistoricalDataTickerIncrement
    
        If mHistoricalDataTable(index).bars.size <> 0 Then
            Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
            mHistoricalDataTable(index).dataRetrievedFromTWS = True
            mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
            mHistoricalDataTable(index).historicalDataReader.historicalDataAvailable
        Else
            logMessage "Error " & errorCode & " from TWS: " & errorMsg, LogLevelNormal
            mHistoricalDataTable(index).historicalDataReader.historicalDataRequestFailed errorMsg
            releaseHistoricalDataTableEntry index
        End If
        processNextHistDataRequest
    Case IdTypeOrder
    End Select
Case 326    ' clientID already in use

Case 354    ' no market data subscription
    failpoint = 1250
    
    Select Case getIdType(id)
    Case IdTypeNone

    Case IdTypeRealtimeData
        mTickerTable(id Mod IncarnationIncrement).dataReader.marketDataRequestFailed errorMsg
        releaseTickerTableEntry (id Mod IncarnationIncrement)
    Case IdTypeMarketDepth

    Case IdTypeHistoricalData

    Case IdTypeOrder

    End Select
    

Case 1100
    failpoint = 1300

    ' connectivity between TWS and IB has been lost
    RaiseEvent ConnectionToIBClosed
    logMessage "Connection to IB has been lost", LogLevelNormal
    
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
    
    logMessage "Connection to IB recovered: market data re-established", LogLevelNormal
    
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
    
    logMessage "Connection to IB recovered: no loss of data", LogLevelNormal
    
Case 2103, 2104, 2105, 2106, 2107, 2108
    failpoint = 1600

    logMessage errorMsg, LogLevelNormal
Case Else
    failpoint = 1700

    logMessage "IBTWSSP:TwsAPI:getErrorMsg: Error (" & errorCode & ") from TWS: " & errorMsg, LogLevelSevere
    mCommonServiceConsumer.ServiceProviderError errorCode, errorMsg, -1
End Select

Exit Sub

Err:
Err.Raise Err.number, , "TWSAPI::getErrorMsg(failpoint=" & failpoint & "): " & Err.Description
End Sub

Private Sub getExecutionData(ByVal version As Long)
Dim timeString  As String

Dim symbol As String
Dim localSymbol As String
Dim sectype As String ' keep the socket string for logging
Dim expiry As String
Dim strike As Double
Dim rightString As String
Dim exchange As String
Dim currencyCode As String

Dim execId As String
Dim side As String
Dim Price As Double
Dim mapEntry As OrderIdMapEntry
Dim id As Long
Dim permId As Long
Dim clientID As Long
Dim lContractSpecifier As ContractSpecifier

Dim exec As Execution

id = getLong("Id")

On Error Resume Next
mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
On Error GoTo 0

symbol = getString("Symbol")
sectype = getString("Sec type")
expiry = getString("Expiry")
strike = getDouble("Strike")
rightString = getString("Right")
exchange = getString("Exchange")
currencyCode = getString("Currency")
localSymbol = getString("Local symbol")

Set lContractSpecifier = CreateContractSpecifier(localSymbol, _
                                                symbol, _
                                                exchange, _
                                                secTypeFromString(sectype), _
                                                currencyCode, _
                                                expiry, _
                                                strike, _
                                                optionRightFromString(rightString))

Set exec.ContractSpecifier = lContractSpecifier
exec.OrderBrokerId = id
If Not mapEntry.order Is Nothing Then exec.OrderTradeBuildId = mapEntry.order.tradeBuildId
exec.execId = getString("Exec id")
timeString = getString("Exec time")
exec.time = CDate(Left$(timeString, 4) & "/" & _
                    Mid$(timeString, 5, 2) & "/" & _
                    Mid$(timeString, 7, 2) & " " & _
                    Right$(timeString, 8))
exec.AccountIdentifier = getString("Account id")
exec.FillingExchange = getString("Filling exchange")
side = getString("Side")
exec.Action = IIf(side = "BUY", OrderActions.ActionBuy, OrderActions.ActionSell)
exec.Quantity = getLong("Quantity")
Price = getDouble("Price")
exec.Price = adjustPrice(Price, lContractSpecifier)
    
If version >= 2 Then permId = getLong("Perm Id")

If version >= 3 Then clientID = getLong("Client id")  ' clientID
   
If version >= 4 Then exec.IsLiquidation = getBoolean("Is liquidation")

logSocketMessage

Dim existingExec As Execution
On Error Resume Next
existingExec = mExecutions(execId)
If Err.number = 0 Then
    ' we've had this execution before so nothing to do
    On Error GoTo 0
Else
    On Error GoTo 0
    If Not mapEntry.OrderSubmitter Is Nothing Then
        Dim qtyRemaining As Long
        qtyRemaining = mapEntry.order.quantityRemaining
        
        mapEntry.OrderSubmitter.executionDetails lContractSpecifier, _
                                                    exec.AccountIdentifier, _
                                                    exec.Action, _
                                                    exec.execId, _
                                                    exec.FillingExchange, _
                                                    exec.IsLiquidation, _
                                                    exec.OrderBrokerId, _
                                                    exec.OrderTradeBuildId, _
                                                    exec.Price, _
                                                    exec.Quantity, _
                                                    exec.time
        
        ' notify a status update - this is in case this execution occurred during a period when
        ' TWS had lost its connection to the IB servers and this exec details is being notified
        ' in response to the reqExecutions call made when the connection is recovered. In these
        ' circumstances TWS will not itself issue an orderStatus.
        If exec.Quantity >= qtyRemaining Then
            mapEntry.OrderSubmitter.updateStatus mapEntry.order.tradeBuildId, _
                                                OrderStatuses.OrderStatusFilled
        Else
            mapEntry.OrderSubmitter.updateStatus mapEntry.order.tradeBuildId, _
                                                OrderStatuses.OrderStatusSubmitted
        End If
    End If
End If

End Sub

Private Sub getFAData(ByVal version As Long)
Dim DataType As FADataTypes: DataType = getLong("Data type")
Dim XMLData As String: XMLData = getString("XML data")

logSocketMessage

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
Static totalVolume As Currency  ' can get too large for a Long
Static lCustomSessionBuilder As SessionBuilder

Dim bardate As String
Dim timestamp As Date

Dim failpoint As Long

On Error GoTo Err

If checkPointNumber = 0 Then
    hint = 0
    i = 0
    totalVolume = 0
    
    requestID = getLong("Request id")
    index = (requestID - HistoricalDataTickerBase) Mod HistoricalDataTickerIncrement
    reqNo = Int((requestID - HistoricalDataTickerBase) / HistoricalDataTickerIncrement)
    
    If mHistoricalDataTable(index).barSpecifier.customSessionStartTime <> 0 Or _
        mHistoricalDataTable(index).barSpecifier.customSessionEndTime <> 0 _
    Then
        Set lCustomSessionBuilder = New SessionBuilder
        lCustomSessionBuilder.sessionStartTime = mHistoricalDataTable(index).barSpecifier.customSessionStartTime
        lCustomSessionBuilder.sessionEndTime = mHistoricalDataTable(index).barSpecifier.customSessionEndTime
        lCustomSessionBuilder.TimeZone = mHistoricalDataTable(index).barSpecifier.Contract.TimeZone
    End If
    
    checkpointBuffer
    checkPointNumber = 100
End If

If checkPointNumber = 100 Then
    If version >= 2 Then
        getString ("Start date")      ' startDateStr
    End If
    checkpointBuffer
    checkPointNumber = 110
End If

If checkPointNumber = 110 Then
    If version >= 2 Then
        getString ("End date")      ' endDateStr
    End If
    checkpointBuffer
    checkPointNumber = 120
End If

If checkPointNumber = 120 Then
    itemCount = getLong("Item count")
    Debug.Print "Bars retrieved: " & itemCount
    
    checkpointBuffer
    checkPointNumber = 200
End If

If checkPointNumber >= 200 Then
    For i = i To itemCount - 1
        If checkPointNumber = 200 Then
            bardate = getString("Bar date")
            If Len(bardate) = 8 Then
                timestamp = CDate(Left$(bardate, 4) & "/" & _
                                                Mid$(bardate, 5, 2) & "/" & _
                                                Mid$(bardate, 7, 2))
            Else
                timestamp = gRoundTimeToSecond(CDate(Left$(bardate, 4) & "/" & _
                                                Mid$(bardate, 5, 2) & "/" & _
                                                Mid$(bardate, 7, 2) & " " & _
                                                Right$(bardate, 8)))
            End If
            Bar.timestamp = ConvertDateUTCToTZ( _
                                    ConvertDateLocalToUTC(timestamp), _
                                    mHistoricalDataTable(index).barSpecifier.Contract.TimeZone)
            If mHistoricalDataTable(index).ignoreTimestamps Then
                Bar.timestamp = Bar.timestamp + mHistoricalDataTable(index).barSpecifier.Contract.sessionStartTime
            End If
            
            If i = 0 Then
                If mHistoricalDataTable(index).barSpecifier.barTimePeriod.units = TimePeriodVolume Or _
                    mHistoricalDataTable(index).barSpecifier.barTimePeriod.units = TimePeriodTickMovement _
                Then
                    mHistoricalDataTable(index).barSpecifier.ToDate = mHistoricalDataTable(index).Session.OffsetBarStartTime(timestamp, GetTimePeriod(15, TimePeriodSecond), -1)
                Else
                    mHistoricalDataTable(index).barSpecifier.ToDate = mHistoricalDataTable(index).Session.OffsetBarStartTime(timestamp, mHistoricalDataTable(index).barSpecifier.barTimePeriod, -1)
                End If
            End If
            
            checkpointBuffer
            checkPointNumber = 210
        End If
        If checkPointNumber = 210 Then
            Bar.openValue = getDouble("Open")
            checkpointBuffer
            checkPointNumber = 220
        End If
        If checkPointNumber = 220 Then
            Bar.highValue = getDouble("High")
            checkpointBuffer
            checkPointNumber = 230
        End If
        If checkPointNumber = 230 Then
            Bar.lowValue = getDouble("Low")
            checkpointBuffer
            checkPointNumber = 240
        End If
        If checkPointNumber = 240 Then
            Bar.closeValue = getDouble("Close")
            checkpointBuffer
            checkPointNumber = 250
        End If
        If checkPointNumber = 250 Then
            Bar.volume = getLong("Volume")
            totalVolume = totalVolume + Bar.volume
            checkpointBuffer
            checkPointNumber = 260
        End If
        If checkPointNumber = 260 Then
            Bar.WAP = getDouble("WAP")
            checkpointBuffer
            checkPointNumber = 270
        End If
        If checkPointNumber = 270 Then
            Bar.hasGaps = getBoolean("Has gaps")
            checkpointBuffer
            checkPointNumber = 280
        End If
        If checkPointNumber = 280 Then
            If version >= 3 Then
                Bar.tickVolume = getLong("Tick volume")
                checkpointBuffer
                checkPointNumber = 280
            End If
        End If
        
        If Not mHistoricalDataTable(index).cancelling Then
            mHistoricalDataTable(index).SessionBuilder.setSessionCurrentTime Bar.timestamp
            If Not lCustomSessionBuilder Is Nothing Then
                lCustomSessionBuilder.setSessionCurrentTime Bar.timestamp
                If lCustomSessionBuilder.Session.isTimeInSession(Bar.timestamp) Then
                    If mHistoricalDataTable(index).barSpecifier.includeBarsOutsideSession Or _
                        mHistoricalDataTable(index).Session.isTimeInSession(Bar.timestamp) Or _
                        mHistoricalDataTable(index).ignoreTimestamps _
                    Then
                        hint = mHistoricalDataTable(index).bars.add(Bar, Bar.timestamp, hint)
                    End If
                End If
            Else
                If mHistoricalDataTable(index).barSpecifier.includeBarsOutsideSession Or _
                    mHistoricalDataTable(index).Session.isTimeInSession(Bar.timestamp) Or _
                    mHistoricalDataTable(index).ignoreTimestamps _
                Then
                    hint = mHistoricalDataTable(index).bars.add(Bar, Bar.timestamp, hint)
                End If
            End If
        End If
        
        checkPointNumber = 200
    Next
End If
checkPointNumber = 0

logSocketMessage

If mHistoricalDataTable(index).cancelling Then Exit Sub

mHistoricalDataTable(index).requestKeys(reqNo) = ""

mConcurrentHistDataRequests = mConcurrentHistDataRequests - 1
processNextHistDataRequest

failpoint = 500

If mHistoricalDataTable(index).barSpecifier.barTimePeriod.units = TimePeriodVolume Then
    mHistoricalDataTable(index).barSpecifier.maxNumberOfBars = mHistoricalDataTable(index).barSpecifier.maxNumberOfBars - Int(totalVolume / mHistoricalDataTable(index).barSpecifier.barTimePeriod.length)
Else
    mHistoricalDataTable(index).barSpecifier.maxNumberOfBars = mHistoricalDataTable(index).barSpecifier.maxNumberOfBars - itemCount
End If
If mHistoricalDataTable(index).barSpecifier.maxNumberOfBars <= 0 Then
    
    failpoint = 600
    
    Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
    mHistoricalDataTable(index).dataRetrievedFromTWS = True
    mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
    mHistoricalDataTable(index).historicalDataReader.historicalDataAvailable
Else
    
    failpoint = 700
    
    ' request the next lot of bars
    generateTwsHistoricalDataRequest index
End If


Exit Sub

Err:

If Err.number = ErrHistGenerationFailed Then
    mHistoricalDataTable(index).barSpecifier.maxNumberOfBars = 0
    Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
    mHistoricalDataTable(index).dataRetrievedFromTWS = True
    mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
    mHistoricalDataTable(index).historicalDataReader.historicalDataAvailable
    Exit Sub
End If

If Err.number <> DataIncomplete Then
    gLogger.Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "getHistoricalData" & "." & IIf(failpoint <> 0, failpoint, checkPointNumber) & _
                                IIf(Err.source <> "", vbCrLf & Err.source, "") & vbCrLf & _
                                Err.Description
End If

Err.Raise Err.number, _
        ProjectName & "." & ModuleName & ":" & "getHistoricalData" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description


End Sub

Private Function getIdType( _
                ByVal id As Long) As IdTypes
If id >= OrderIdBase Then
    getIdType = IdTypeOrder
ElseIf id >= HistoricalDataTickerBase Then
    getIdType = IdTypeHistoricalData
ElseIf id >= MarketDepthTickerBase Then
    getIdType = IdTypeMarketDepth
ElseIf id >= 0 Then
    getIdType = IdTypeRealtimeData
Else
    getIdType = IdTypeNone
End If
End Function

Private Sub getManagedAccounts(ByVal version As Long)
Dim accountsList As String: accountsList = getString("Accounts list")

logSocketMessage

End Sub

Private Sub getMarketDepth(ByVal version As Long)
Dim id As Long: id = getLong("Id")
Dim tick As GenericTick
tick.tickType = TickTypeMarketDepth
tick.position = getLong("Position")
tick.operation = getLong("Operation")
tick.side = getLong("Side")
tick.Price = getDouble("Price")
tick.size = getLong("Size")

logSocketMessage

With mTickerTable(id - MarketDepthTickerBase)
    
    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    tick.timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.TimeZone)
    
    .dataConsumer.tick tick
    
End With

End Sub

Private Sub getMarketDepthL2(ByVal version As Long)
Dim id As Long: id = getLong("Id")
Dim tick As GenericTick
tick.tickType = TickTypeMarketDepth
tick.position = getLong("Position")
tick.marketmaker = getString("Marketmaker")
tick.operation = getLong("Operation")
tick.side = getLong("Side")
tick.Price = getDouble("Price")
tick.size = getLong("Size")

logSocketMessage

With mTickerTable(id - MarketDepthTickerBase)
    
    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    tick.timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.TimeZone)
    
    .dataConsumer.tick tick
    
End With

End Sub

Private Sub getNewsBulletins(ByVal version As Long)
Dim msgID As Long: msgID = getLong("Id")
Dim msgType As Long: msgType = getLong("Type")
Dim newsMessage As String: newsMessage = getString("Message")
Dim originatingExch As String: originatingExch = getString("Orig exchange")

logSocketMessage

End Sub

Private Sub getNextValidId(ByVal version As Long)

mNextOrderID = getLong("Next id")

logSocketMessage

If mNextOrderID < OrderIdBase Then mNextOrderID = OrderIdBase

End Sub

Private Sub getOpenOrder(ByVal version As Long)
Dim id As Long

Dim symbol As String
Dim localSymbol As String
Dim sectype As String ' keep the socket string for logging
Dim expiry As String
Dim strike As Double
Dim rightString As String
Dim exchange As String
Dim currencyCode As String

Dim Action As String
Dim orderType As String
Dim timeInForce As String

Dim lContractSpecifier As ContractSpecifier

id = getLong("Id")

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
symbol = getString("Symbol")
sectype = getString("Sec type")
expiry = getString("Expiry")
strike = getDouble("Strike")
rightString = getString("Right")
exchange = getString("Exchange")
currencyCode = getString("Currency")
If version >= 2 Then localSymbol = getString("Local symbol")

Set lContractSpecifier = CreateContractSpecifier(localSymbol, _
                                                symbol, _
                                                exchange, _
                                                secTypeFromString(sectype), _
                                                currencyCode, _
                                                expiry, _
                                                strike, _
                                                optionRightFromString(rightString))


' read order fields
'With lOrder
    Action = getString("Action")
    '.action = orderActionFromString(action)
     getLong "Quantity"   ' .quantity
    orderType = getString("Order type")
    '.orderType = orderTypeFromString(orderType)
    getDouble "Limit price"  ' .limitPrice
    getDouble "Aux price"   ' .auxPrice
    timeInForce = getString("Time in force")
    '.timeInForce = orderTIFFromString(timeInForce)
    getString "OCA group"   ' .ocaGroup
    getString "Account"   ' .account
    getString "Open/close"   ' .openClose
    getLong "Origin"     ' .origin
    getString "Order ref"  ' .orderRef
    If version >= 3 Then getLong "Client id"   ' .clientID
    
    If version >= 4 Then
        getLong "Perm id"    ' .permId =
        getBoolean "Ignore RTH"  ' .ignoreRTH
        getBoolean "Hidden" ' .hidden
        getDouble "Discr amt"  ' .discretionaryAmt
    End If

    If version >= 5 Then getString "Good after time"  ' .goodAfterTime
    
    If version >= 6 Then getString "Shares allocation"  ' .sharesAllocation
    
    If version >= 7 Then
        getString "FA Group"   ' .FAGroup
        getString "FA method"  ' .FAMethod
        getString "FA Percentage"  ' .FAPercentage
        getString "FA Profile"  ' .FAProfile
    End If
    
    If version >= 8 Then getString "Good till date"  ' goodTillDate

    If version >= 9 Then
        getString "Rule80A"          ' rule80A
        getDouble "Percent Offset"
        getString "Settling Firm"
        getLong "shortSaleSlot"
        getString "designatedLocation"
        getLong "Auction Strategy"
        getDouble "starting Price"
        getDouble "stockRefPrice"
        getDouble "delta"
        getDouble "stockRangeLower"
        getDouble " stockRangeUpper"
        getLong "displaySize"
        getBoolean "rthOnly"
        getBoolean "blockOrder"
        getBoolean "sweepToFill"
        getBoolean "allOrNone"
        getLong "minQty"
        getLong "ocaType"
        getBoolean "eTradeOnly"
        getBoolean "firmQuoteOnly"
        getDouble "nbboPriceCap"
    End If

    If version >= 10 Then
        getLong "parentId"
        getLong "triggerMethod"
    End If

    If version >= 11 Then
        getDouble "volatility"
        getLong "volatilityType"
        If version = 11 Then
            getLong "receivedInt"
        Else ' version 12 and up
            getString "deltaNeutralOrderType"
            getDouble "deltaNeutralAuxPrice"
        End If
        getLong "continuousUpdate"
        If mServerVersion = 26 Then
            getDouble "stockRangeLower"
            getDouble "stockRangeUpper"
        End If
        getLong "referencePriceType"
    End If
    
    If version >= 13 Then
        getDouble "trailStopPrice"
    End If

    If version >= 14 Then
        getDouble "basisPoints"
        getLong "basisPointsType"
        getString "comboLegsDescrip"
    End If
'End With

logSocketMessage

End Sub

Private Sub getOrderStatus(ByVal version As Long)
Dim mapEntry As OrderIdMapEntry

Dim id As Long: id = getLong("id")
Dim statusString As String: statusString = getString("Status")
Dim status As TradeBuildSP.OrderStatuses: status = orderStatusFromString(statusString)
Dim filled As Long: filled = getLong("Filled")
Dim remaining As Long: remaining = getLong("Remaining")
Dim avgFillPrice As Double:  avgFillPrice = getDouble("Avg fill price")
Dim lastFillPrice As Double
Dim clientID As Long
Dim permId As Long
Dim parentId As Long
Dim whyHeld As String

If version >= 2 Then permId = getLong("Perm id")

If version >= 3 Then parentId = getLong("Parent id")

If version >= 4 Then lastFillPrice = getDouble("Last fill price")

If version >= 5 Then clientID = getLong("Client id")

If version >= 6 Then whyHeld = getString("Why held")

logSocketMessage

On Error Resume Next
mapEntry = mOrderIDMap(generateTwsOrderKey(id, clientID))
On Error GoTo 0
If Not mapEntry.order Is Nothing Then
    avgFillPrice = adjustPrice(avgFillPrice, mapEntry.order.Contract.specifier)
    lastFillPrice = adjustPrice(lastFillPrice, mapEntry.order.Contract.specifier)
    
    mapEntry.OrderSubmitter.updateStatus mapEntry.order.tradeBuildId, _
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

End Sub

Private Sub getPortfolioValue(ByVal version As Long)

Dim symbol As String
Dim localSymbol As String
Dim sectype As String ' keep the socket string for logging
Dim expiry As String
Dim strike As Double
Dim rightString As String
Dim exchange As String
Dim currencyCode As String

Dim lContractSpecifier As ContractSpecifier

symbol = getString("Symbol")
sectype = getString("Sec type")
expiry = getString("Expiry")
strike = getDouble("Strike")
rightString = getString("Right")
currencyCode = getString("Currency")
If version >= 2 Then localSymbol = getString("Locale symbol")

Set lContractSpecifier = CreateContractSpecifier(localSymbol, _
                                                symbol, _
                                                exchange, _
                                                secTypeFromString(sectype), _
                                                currencyCode, _
                                                expiry, _
                                                strike, _
                                                optionRightFromString(rightString))

Dim position As Long: position = getLong("Position")
Dim marketPrice As Double: marketPrice = getDouble("Market price")
Dim marketValue As Double: marketValue = getDouble("Market value")
Dim averageCost As Double: averageCost = 0#
Dim unrealizedPNL As Double: unrealizedPNL = 0#
Dim realizedPNL As Double: realizedPNL = 0#
If version >= 3 Then
    averageCost = getDouble("Avg cost")
    unrealizedPNL = getDouble("Unrealized PNL")
    realizedPNL = getDouble("Realized PNL")
End If
Dim accountName As String: If version >= 4 Then accountName = getString("Account name")
                         
logSocketMessage
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
Dim sectype As SecurityTypes
Dim expiry As String
Dim strike As Double
Dim optRight As OptionRights
Dim exchange As String
Dim currencyCode As String

tickerid = getLong("Id")
numberOfElements = getLong("Number of elements")

For i = 0 To numberOfElements - 1
    rank = getLong("Rank")

    symbol = getString("Symbol")
    sectype = secTypeFromString(getString("Sec type"))
    expiry = getString("Expiry")
    strike = getDouble("Stirke")
    optRight = optionRightFromString(getString("Right"))
    exchange = getString("Exchange")
    currencyCode = getString("Currency")
    localSymbol = getString("Local symbol")
    
    Set lContractSpecifier = CreateContractSpecifier(localSymbol, _
                                                    symbol, _
                                                    exchange, _
                                                    sectype, _
                                                    currencyCode, _
                                                    expiry, _
                                                    strike, _
                                                    optRight)
    Set lContractBuilder = CreateContractBuilder(lContractSpecifier)
    getString "marketName"
    getString "tradingClass"
    distance = getString("Distance")
    benchmark = getString("Benchmark")
    projection = getString("Projection")
Next

logSocketMessage
End Sub

Private Sub getScannerParameters(ByVal version As Long)
Dim xml  As String
xml = getString("XML")
End Sub

Private Sub getTickPrice(ByVal version As Long)
Dim tick As GenericTick
Dim tickType As Long
Dim tickerid As Long
Dim canAutoExecute As Boolean
Dim failpoint As Long

On Error GoTo Err

failpoint = 100

tickerid = getLong("Ticker id")

failpoint = 200

tickType = getLong("Tick type")

failpoint = 300

tick.Price = getDouble("Price")

failpoint = 400

If version = 1 Then Debug.Print "****************** Type 1 tick price message received *********************"
If version >= 2 Then tick.size = getLong("Size")
If version >= 3 Then canAutoExecute = getBoolean("Can autoexecute")

logSocketMessage

failpoint = 500

With mTickerTable(tickerid Mod IncarnationIncrement)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    If tick.Price < (0.5 * .prevTrade) Then
        Exit Sub        ' ignore invalid prices
    End If
    
    tick.timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.TimeZone)
    
    Select Case tickType
    Case TICK_BID
        tick.tickType = TickTypeBid
        If tick.size = 0 Then tick.size = .prevBidSize
        .prevBid = tick.Price
        If tick.size <> 0 Then .prevBidSize = tick.size
    
    Case TICK_ASK
        tick.tickType = TickTypeAsk
        If tick.size = 0 Then tick.size = .prevAskSize
        .prevAsk = tick.Price
        If tick.size <> 0 Then .prevAskSize = tick.size
    
    Case TICK_LAST
        tick.tickType = TickTypeTrade
        If tick.size = 0 Then tick.size = .prevTradeSize
        .prevTrade = tick.Price
        If tick.size <> 0 Then .prevTradeSize = tick.size
    
    Case TICK_HIGH
        tick.tickType = TickTypeHighPrice
    
    Case TICK_LOW
        tick.tickType = TickTypeLowPrice
    
    Case TICK_OPEN
        tick.tickType = TickTypeOpenPrice
    
    Case TICK_CLOSE
        tick.tickType = TickTypeClosePrice
    
    End Select
        
    .dataConsumer.tick tick
End With


Exit Sub

Err:
Err.Raise Err.number, , "TWSAPI::getTickPrice(failpoint=" & failpoint & "): " & Err.Description
End Sub

Private Sub getTickSize(ByVal version As Long)
Dim tick As GenericTick
Dim tickType As Long
Dim tickerid As Long
Dim failpoint As Long

On Error GoTo Err

failpoint = 100

tickerid = getLong("Ticker id")

failpoint = 200

tickType = getLong("Tick type")

failpoint = 300

tick.size = getLong("Size")

logSocketMessage

failpoint = 400

With mTickerTable(tickerid Mod IncarnationIncrement)
    
    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    tick.timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.TimeZone)
    
    Select Case tickType
    Case TICK_BID_SIZE
        tick.tickType = TickTypeBid
        If tick.size = .prevBidSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        tick.Price = .prevBid
        .prevBidSize = tick.size
    
    Case TICK_ASK_SIZE
        tick.tickType = TickTypeAsk
        If tick.size = .prevAskSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        tick.Price = .prevAsk
        .prevAskSize = tick.size
    
    Case TICK_LAST_SIZE
        tick.tickType = TickTypeTrade
        If tick.size = .prevTradeSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        tick.Price = .prevTrade
        .prevTradeSize = tick.size
    
    Case TICK_VOLUME
        tick.tickType = TickTypeVolume
    
    Case TICK_OPEN_INTEREST
        tick.tickType = TickTypeOpenInterest
    
    End Select

    .dataConsumer.tick tick
End With

Exit Sub

Err:
Err.Raise Err.number, , "TWSAPI::getTickSize(failpoint=" & failpoint & "): " & Err.Description
End Sub

Private Sub getTickString(ByVal version As Long)
getLong "tickerId"
getLong "tickType"
getString "value"

logSocketMessage
End Sub

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)
logMessage source & " Error (" & number & "): " & Description, LogLevelSevere

mCommonServiceConsumer.FatalServiceProviderError number, source, Description, -1

If mConnectionState = ConnConnected Then gReleaseTWSAPIInstance Me, True
End Sub

Private Sub handleTwsDisconnection()
mConnectionState = ConnectionStates.ConnNotConnected
ReleaseSocket
If mDisconnectedByApplication Then
    RaiseEvent ConnectionToTWSClosed(False)
    logMessage "Connection to TWS closed: " & connectionString, LogLevelNormal
Else
    If mConnectionRetryIntervalSecs = 0 Then
        cancelContractRequests "Lost connection to TWS"
        RaiseEvent ConnectionToTWSClosed(False)
        logMessage "Connection to TWS closed" & _
                    connectionString
    Else
        retryConnection
        RaiseEvent ConnectionToTWSClosed(True)
        logMessage "Connection to TWS closed - reconnecting: " & _
                    connectionString, _
                    LogLevelNormal
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
                Optional ByVal logLevel As LogLevels = LogLevelNormal)
If Not gLogger.isLoggable(logLevel) Then Exit Sub
gLogger.Log logLevel, _
            message
End Sub

Private Sub logSocketMessage()
If mSocketLogger.isLoggable(LogLevelHighDetail) Then
    mSocketLogger.Log LogLevelHighDetail, mMessageBuilder.toString
    Set mMessageBuilder = Nothing
End If
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
    Err.Raise ErrorCodes.ErrIllegalArgumentException
End Select
End Function

Private Function permittedHistDays( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
permittedHistDays = DateDiff("d", DateAdd("yyyy", -1, Now), reqEndDate)
If permittedHistDays = 0 Then Err.Raise ErrHistGenerationFailed
If permittedHistDays > desiredBars Then permittedHistDays = desiredBars
End Function

Private Function permittedHistMonths( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
permittedHistMonths = DateDiff("m", DateAdd("m", 1, DateAdd("yyyy", -1, Now)), reqEndDate)
If permittedHistMonths = 0 Then Err.Raise ErrHistGenerationFailed
If permittedHistMonths > desiredBars Then permittedHistMonths = desiredBars
End Function

Private Function permittedHistSeconds( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
' the 86398 should really be 86400 but allows for clocks not being
' properly synchronised
permittedHistSeconds = DateDiff("s", DateAdd("yyyy", -1, Now), reqEndDate) + 86398
If permittedHistSeconds = 0 Then Err.Raise ErrHistGenerationFailed
If permittedHistSeconds > desiredBars Then permittedHistSeconds = desiredBars
End Function

Private Function permittedHistWeeks( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
permittedHistWeeks = DateDiff("w", DateAdd("w", 1, DateAdd("yyyy", -1, Now)), reqEndDate)
If permittedHistWeeks = 0 Then Err.Raise ErrHistGenerationFailed
If permittedHistWeeks > desiredBars Then permittedHistWeeks = desiredBars
End Function

Private Sub processContractError(ByVal reason As String)
Dim failpoint As Long

On Error GoTo Err

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

Err:
Err.Raise Err.number, , "TWSAPI::processContractError(failpoint=" & failpoint & "): " & Err.Description
End Sub

Private Sub processNextHistDataRequest()
Dim failpoint As Long
On Error GoTo Err

If mTwsHistDataRequestQueue.size <> 0 Then
    Dim request As TWSHistoricalDataRequestQueueEntry
    Dim requestKey As String
    request = mTwsHistDataRequestQueue.removeFromFront
    reqHistoricalData request
End If

Exit Sub

Err:
Err.Raise Err.number, _
        ProjectName & "." & ModuleName & ":" & "processNextHistDataRequest" & "." & failpoint & _
        IIf(Err.source <> "", vbCrLf & Err.source, ""), _
        Err.Description

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

Const version = 3

addData REQ_CONTRACT_DATA
addData version
With ContractSpecifier
    addData .symbol
    addData secTypeToString(.sectype)
    addData Left$(.expiry, 6)
    addDouble .strike
    addData optionRightToString(.Right)
    If mServerVersion >= 15 Then addData "" ' multiplier
    addData .exchange
    addData .currencyCode
    If (.sectype = SecTypeFuture Or _
        .sectype = SecTypeFuturesOption Or _
        .sectype = SecTypeOption) _
    Then
        If expiryToDate(.expiry) >= Now Or .expiry = "" Then
            addData .localSymbol
        Else
            addData ""
        End If
    Else
        addData .localSymbol
    End If
    If mServerVersion >= 31 Then
        If .sectype = SecTypeCash Or .sectype = SecTypeIndex Or .sectype = SecTypeStock Then
            addData 0  ' can't include expired for non-expiring contracts
        Else
            addData 1  ' include expired
        End If
    End If
End With
send
End Sub

Private Sub reqHistoricalData( _
                request As TWSHistoricalDataRequestQueueEntry)
Debug.Print "ReqHistData from TWS: barsize=" & request.barSizeSetting & "; duration=" & request.duration
Const version = 4

If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

addData REQ_HISTORICAL_DATA
addData version
addData request.id
addData request.Contract.specifier.symbol
addData secTypeToString(request.Contract.specifier.sectype)
addData Left$(request.Contract.specifier.expiry, 6)
addDouble request.Contract.specifier.strike
addData optionRightToString(request.Contract.specifier.Right)
addData ""      ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>> multiplier ??? why?
addData request.Contract.specifier.exchange
addData ""      ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>> primaryExch ??? not in ActiveX interface
addData request.Contract.specifier.currencyCode
If expiryToDate(request.Contract.specifier.expiry) < Now Then
    addData ""
Else
    addData request.Contract.specifier.localSymbol
End If
If mServerVersion >= 31 Then
    If request.Contract.specifier.sectype = SecTypeCash Or _
        request.Contract.specifier.sectype = SecTypeIndex Or _
        request.Contract.specifier.sectype = SecTypeStock _
    Then
        addData 0  ' can't include expired for non-expiring contracts
    Else
        addData 1  ' include expired
    End If
End If
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

If request.Contract.specifier.ComboLegs Is Nothing Then
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

Const version = 6

addData REQ_MKT_DATA
addData version
addData tickerTableIndex + mTickerTable(tickerTableIndex).incarnation
With mTickerTable(tickerTableIndex).Contract.specifier
    addData .symbol
    addData secTypeToString(.sectype)
    addData Left$(.expiry, 6)
    addDouble .strike
    addData optionRightToString(.Right)
    If mServerVersion >= 15 Then addData "" ' multiplier
    addData .exchange
    If mServerVersion >= 15 Then addData "" ' primary exchange
    addData .currencyCode
    If mServerVersion >= 2 Then addData .localSymbol
    ' Add combo legs for BAG requests
    If mServerVersion >= 8 And .sectype = SecurityTypes.SecTypeCombo Then
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
    If mServerVersion >= 31 Then addData "" ' generic tick list
    
    send
End With

End Sub

Private Sub reqMarketDepth(ByVal tickerid As Long, _
                        ByVal contractSpec As ContractSpecifier)

Const version = 3

addData REQ_MKT_DEPTH
addData version
addData tickerid + MarketDepthTickerBase
With contractSpec
    addData .symbol
    addData secTypeToString(.sectype)
    addData Left$(.expiry, 6)
    addDouble .strike
    addData optionRightToString(.Right)
    If mServerVersion >= 15 Then addData "" ' multiplier
    addData .exchange
    addData .currencyCode
    addData .localSymbol
    If mServerVersion >= 19 Then addData 20 ' number of rows
End With
send
End Sub

Private Sub resetMarketDepth(ByVal reEstablish As Boolean)
Dim i As Long
Dim tick As GenericTick

tick.tickType = TickTypeMarketDepthReset

For i = 0 To UBound(mTickerTable)
    With mTickerTable(i)
        If .receivingMarketDepth Then
            tick.timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.TimeZone)
            If reEstablish Then reqMarketDepth i, .Contract.specifier
            .dataConsumer.tick tick
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
ReDim mBufferOut(15) As Byte
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

Const version = 21
addData PLACE_ORDER
addData version
addData pOrder.brokerId

' send contract fields
With specifier
    addData .symbol
    addData secTypeToString(.sectype)
    addData Left$(.expiry, 6)
    addDouble .strike
    addData optionRightToString(.Right)
    If mServerVersion >= 15 Then addData ""     ' multiplier - this is the number of units of the underlying
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
    If mServerVersion >= 7 Then addBoolean .Hidden
End With

' Send combo legs for BAG requests
If mServerVersion >= 8 And specifier.sectype = SecurityTypes.SecTypeCombo Then
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
    
    If mServerVersion >= 30 Then addData 0 ' TRAIL_STOP_LIMIT stop price
End With

send
End Sub
                    
Private Sub setTWSLogLevel(logLevel As TWSLogLevels)

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
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

On Error GoTo Err

If Not mConnectionState = ConnConnected Then Exit Sub
If Not mContractDetailsRequest.contractSpec Is Nothing Then Exit Sub
If mTwsContractDetailsRequestQueue.size = 0 Then Exit Sub
    
failpoint = 100
On Error Resume Next
mContractDetailsRequest = mTwsContractDetailsRequestQueue.removeFromFront
If Err.number <> 0 Then Stop
On Error GoTo Err

failpoint = 200
Set mContractDetailsRequest.contractWrappers = New Collection
reqContract

Exit Sub

Err:
Err.Raise Err.number, , "TWSAPI::submitNextContractDetailsRequest(failpoint=" & failpoint & "): " & Err.Description
End Sub


