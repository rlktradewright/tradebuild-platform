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
    ' Raised when we Disconnect from TWS

'================================================================================
' Constants
'================================================================================

Private Const ProjectName                   As String = "IBTWSSP26"
Private Const ModuleName                    As String = "TWSAPI"

Private Const ClientVersion As Long = 45
Private Const ReqdServerVersion As Long = 38

Private Const ErrHistGenerationFailed = vbObjectError + &H7FFF


Private Const IncarnationIncrement As Long = &H1000&

Private Const IdBaseContractDataReq As Long = &H1000000
Private Const IdBaseMarketDepthReq As Long = &H400000
Private Const IdBaseReqHistoricalData As Long = &H410000
Private Const HistoricalDataReqIdIncrement As Long = &H100
Private Const IdBaseExecutionsReq As Long = &H420000
Private Const IdBaseOrderPlex As Long = &H10000000
Private Const OrderPlexIdIncrement As Long = &H100

'Private Const MaxConcurrentHistDataRequests As Long = 50 ' limit imposed by IB
Private Const MaxConcurrentHistDataRequests As Long = 1 ' try serialising them

Private Const MIN_SERVER_VER_REAL_TIME_BARS As Long = 34
Private Const MIN_SERVER_VER_SCALE_ORDERS As Long = 35
Private Const MIN_SERVER_VER_SNAPSHOT_MKT_DATA As Long = 35
Private Const MIN_SERVER_VER_SSHORT_COMBO_LEGS As Long = 35
Private Const MIN_SERVER_VER_WHAT_IF_ORDERS As Long = 36
Private Const MIN_SERVER_VER_CONTRACT_CONID As Long = 37
Private Const MIN_SERVER_VER_PTA_ORDERS As Long = 39
Private Const MIN_SERVER_VER_FUNDAMENTAL_DATA As Long = 40
Private Const MIN_SERVER_VER_UNDER_COMP As Long = 40
Private Const MIN_SERVER_VER_CONTRACT_DATA_CHAIN As Long = 40
Private Const MIN_SERVER_VER_SCALE_ORDERS2 As Long = 40
Private Const MIN_SERVER_VER_ALGO_ORDERS As Long = 41
Private Const MIN_SERVER_VER_EXECUTION_DATA_CHAIN As Long = 42
Private Const MIN_SERVER_VER_NOT_HELD As Long = 44

Private Const NumDaysInWeek As Long = 5
Private Const NumDaysInMonth  As Long = 22
Private Const NumDaysInYear As Long = 260
Private Const NumMonthsInYear As Long = 12

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
    IdTypeContractData
    IdTypeExecution
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
    requestKey              As String
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
    prevOpen                As Double
    prevHigh                As Double
    prevLow                 As Double
    prevClose               As Double
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

Private mOutputBuffer() As Byte
Private mOutputBufferIndex As Long

Private mInputBuffer() As Byte
Private mInputBufferNextFreeIndex As Long

Private mInputParseIndex As Long
Private mInputParseIndexChk As Long
Private mCheckPointNumber As Long

Private mServerVersion As Long
Private mCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer

Private mConnectionState As ConnectionStates

Private mTickerTable() As TickerTableEntry

Private mContractDetailsRequest As TWSContractDetailsRequestQueueEntry
Private mFirstContractDetailsReceived As Boolean

Private mNextOrderPlexID  As Long

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

Private mInMessageBuilder As StringBuilder

Private mOutMessageBuilder As StringBuilder

Private mRetryingConnection As Boolean

Private mOrderRecoveryServiceConsumer As IOrderRecoveryServiceConsumer

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
ReDim mInputBuffer(1023) As Byte
ReDim mOutputBuffer(15) As Byte

createLoggers

mConnectionState = ConnectionStates.ConnNotConnected

Set mOrderIDMap = New Collection
Set mExecutions = New Collection

mNextOrderPlexID = 1

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
Const ProcName As String = "mConnectionTimer_TimerExpired"
On Error GoTo Err
Set mConnectionTimer = Nothing
Connect
Exit Sub
Err:
handleFatalError ProcName
End Sub

'================================================================================
' mPerformanceTimerSecond Event Handlers
'================================================================================

Private Sub mPerformanceTimerPeriod_TimerExpired()
Const ProcName As String = "mPerformanceTimerPeriod_TimerExpired"
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
        .maxSecondCount = 0
    End With
Next
Exit Sub

Err:
handleFatalError ProcName
End Sub

Private Sub mPerformanceTimerSecond_TimerExpired()
Const ProcName As String = "mPerformanceTimerSecond_TimerExpired"
Dim i As Long

On Error GoTo Err

For i = 0 To TWSSocketInMsgTypes.MAX_SOCKET_INMSG
    With mPerformanceStats(i)
        .lastPeriodCount = .lastPeriodCount + .lastSecondCount
        .lastPeriodTime = .lastPeriodTime + .lastSecondTime
        If .lastSecondTime > .maxSecondCount Then .maxSecondCount = .lastSecondCount
        .totalCount = .totalCount + .lastSecondCount
        .totalTime = .totalTime + .lastSecondTime
        .lastSecondCount = 0
        .lastSecondTime = 0
    End With
Next
Exit Sub

Err:
handleFatalError ProcName
End Sub

'================================================================================
' Socket Event Handlers
'================================================================================

Private Sub mSocket_Close()
Const ProcName As String = "mSocket_Close"
On Error GoTo Err
mSocket.Close
handleTwsDisconnection
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mSocket_Connect()
Const ProcName As String = "mSocket_Connect"
On Error GoTo Err

If mRetryingConnection Then
    ' this means we have been waiting for TWS to start. If we try to start tickers
    ' too soon after TWS accepts socket connections, TWS gives a 'max number of
    ' tickers reached' error. So we'll cancel this connection and connect again in
    ' 10 seconds
    logMessage "TWS is now available - connecting again in 10 seconds to allow TWS time to initialise", "mSocket_Connect"
    mSocket.Close
    Set mConnectionTimer = CreateIntervalTimer(10, ExpiryTimeUnitSeconds)
    mConnectionTimer.StartTimer
    mRetryingConnection = False
    Exit Sub
End If

mConnectionState = ConnectionStates.ConnConnecting
RaiseEvent Connecting

mInputBufferNextFreeIndex = 0
mOutputBufferIndex = 0
mInputParseIndex = 0
mCheckPointNumber = 0

logMessage "Connecting to TWS: negotiating versions and client id", pProcName:="mSocket_Connect"
addData ClientVersion, "Client version"
send
addData mClientID, "Client id"
send
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub mSocket_DataArrival( _
                ByVal bytesTotal As Long)
Const ProcName As String = "mSocket_DataArrival"
Dim bytesReceived As Long
Static processingMessages As Boolean

On Error GoTo Err

If mSocket.State <> StateConstants.sckConnected Then Exit Sub

mSocket.GetData mSocketInputData, vbArray + vbByte
bytesReceived = UBound(mSocketInputData) + 1
'Debug.Print "Received " & bytesReceived & " from socket: " & showBuffer(mSocketInputData, bytesReceived)

If mSocketLogger.IsLoggable(LogLevelMediumDetail) Then
    mSocketLogger.Log LogLevels.LogLevelMediumDetail, _
                        showBuffer("In buf: ", mSocketInputData)
End If

If mInputBufferNextFreeIndex = mInputParseIndex Then
    ' no unprocessed data so start at the beginning of the buffer again
    'Debug.Print "Reverting to start of buffer (no unprocessed data)"
    mInputBufferNextFreeIndex = 0
    mInputParseIndex = 0
    mInputParseIndexChk = 0
End If

If mInputBufferNextFreeIndex + bytesReceived > UBound(mInputBuffer) + 1 Then
    If bytesReceived <= (UBound(mInputBuffer) + 1 - mInputBufferNextFreeIndex + mInputParseIndex) Then
        ' copy the existing contents down to the start of the buffer
        'Debug.Print "Moving unprocessed data (" & mInputBufferNextFreeIndex - mInputParseIndex & " bytes) from " & mInputParseIndex & " to start of buffer"
        MoveMemory mInputBuffer(0), mInputBuffer(mInputParseIndex), mInputBufferNextFreeIndex - mInputParseIndex
        mInputBufferNextFreeIndex = mInputBufferNextFreeIndex - mInputParseIndex
        mInputParseIndex = 0
        mInputParseIndexChk = 0
    Else
        If mSocketLogger.IsLoggable(LogLevelHighDetail) Then
            mSocketLogger.Log LogLevels.LogLevelHighDetail, _
                            "Extending socket buffer by " & IIf(bytesReceived > 1024, bytesReceived, 1024) & " bytes"
        End If
        ReDim Preserve mInputBuffer(UBound(mInputBuffer) + IIf(bytesReceived > 1024, bytesReceived, 1024)) As Byte
    End If
End If
'Debug.Print "Appending socket message to buffer at " & mInputBufferNextFreeIndex
CopyMemory mInputBuffer(mInputBufferNextFreeIndex), mSocketInputData(0), bytesReceived
mInputBufferNextFreeIndex = mInputBufferNextFreeIndex + bytesReceived

If mInputBufferNextFreeIndex <> mInputParseIndex Then
    If processingMessages Then
        ' this is a re-entrant call to DataArrival. We don't want to call checkMessages
        ' here, because the current call to it will process the data we've just
        ' added to the buffer
        mSocketLogger.Log LogLevels.LogLevelNormal, _
                        "Re-entrant call to mSocket_DataArrival"
    Else
        processingMessages = True
        checkmessages
        processingMessages = False
    End If
End If

Exit Sub
Err:
handleFatalError ProcName

End Sub

Private Sub mSocket_Error(ByVal number As Integer, _
                        Description As String, _
                        ByVal Scode As Long, _
                        ByVal source As String, _
                        ByVal HelpFile As String, _
                        ByVal HelpContext As Long, _
                        CancelDisplay As Boolean)
Const ProcName As String = "mSocket_Error"
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
    
    ReleaseSocket
    If mConnectionState = ConnectionStates.ConnNotConnected Then
        If mConnectionRetryIntervalSecs = 0 Then
            logMessage "Failed to connect to TWS", _
                        pMsgQualifier:=Description & ": " & _
                                        ConnectionString, _
                        pProcName:="mSocket_Error"
            
            cancelContractRequests "Failed to connect to TWS"
            RaiseEvent ConnectFailed(Description, False)
        Else
            logMessage "Failed to connect to TWS - retrying in " & mConnectionRetryIntervalSecs & " seconds", _
                        pMsgQualifier:=Description & ": " & _
                                        ConnectionString, _
                        pProcName:="mSocket_Error"
            
            RaiseEvent ConnectFailed(Description, True)
            retryConnection
        End If
    Else
        logMessage "Socket error " & number & ": " & _
                    Description & ": " & _
                    ConnectionString, _
                    "mSocket_Error"
                    
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
                ConnectionString, _
                "mSocket_Error", _
                pLogLevel:=LogLevelSevere
    RaiseEvent ConnectFailed(Description, False)
End Select

Exit Sub
Err:
handleFatalError ProcName
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Friend Property Get clientID() As Long
clientID = mClientID
End Property

Friend Property Let CommonServiceConsumer(ByVal value As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

Friend Property Let ConnectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
If mConnectionRetryIntervalSecs = 0 Then Set mConnectionTimer = Nothing
End Property

Friend Property Get ConnectionState() As ConnectionStates
ConnectionState = mConnectionState
End Property

Friend Property Get ConnectionString() As String
ConnectionString = "server=" & mServer & _
                        " port=" & mPort & _
                        " client Id=" & mClientID
End Property

Friend Property Get ContractDetailsRequestQueueLength() As Long
ContractDetailsRequestQueueLength = mTwsContractDetailsRequestQueue.size
End Property

Friend Property Get InitialOrderId() As Long
InitialOrderId = mNextOrderPlexID
End Property

Friend Property Let Port(ByVal value As Long)
mPort = value
End Property

Friend Property Let ProviderKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let Server(ByVal value As String)
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
Const ProcName As String = "CancelHistoricalData"
Dim key As String

On Error GoTo Err
If mHistoricalDataTable(requestHandle).dataRetrievedFromTWS Then
    ' all data has already been got from TWS
    
    '   more here????
    
    releaseHistoricalDataTableEntry requestHandle

Else
    If mServerVersion < 24 Then
        ' cancel not supported by TWS - we'll complete the cancel when
        ' the data arrives
        mHistoricalDataTable(requestHandle).cancelling = True
    Else
        key = mHistoricalDataTable(requestHandle).requestKey
        If key = "" Then
            ' data has already been retrieved for this TWS request, so
            ' nothing to do
            
        Else
            On Error Resume Next
            mTwsHistDataRequestQueue.remove key
            If Err.number = 0 Then
                ' request was still in queue - no further action required
                On Error GoTo Err
            ElseIf Err.number <> ErrorCodes.ErrIllegalArgumentException Then
                Dim errNum As Long: errNum = Err.number
                Dim errSource As String: errSource = Err.source
                Dim errDesc As String: errDesc = Err.Description
                On Error GoTo Err
                Err.Raise errNum, errSource, errDesc
            Else
                On Error GoTo Err
                
                ' request is with TWS so cancel it
                
                If mConnectionState = ConnConnected Then
                    Const version As Long = 1
            
                    addMsgId CANCEL_HISTORICAL_DATA
                    addData version, "Version"
                    addData generateHistDataRequestId(requestHandle), "Request id"
                    send
                End If
                
                mHistoricalDataTable(requestHandle).cancelling = True
                
                processNextHistDataRequest
            End If
        End If

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
handleFatalError ProcName
End Sub

Friend Sub CancelMarketData(ByVal tickerId As Long)
Const ProcName As String = "CancelMarketData"
On Error GoTo Err
cancelMarketDataEx tickerId, True
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub CancelMarketDepth(ByVal tickerId As Long)
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

Const version = 1

With mTickerTable(tickerId)
    If Not .inUse Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    ProjectName & "." & ModuleName & ":" & ProcName, _
                                    "Invalid ticker id " & tickerId
    If Not .receivingMarketDepth Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Ticker id " & tickerId & " not currently receiving market depth"
    End If
    mTickerTable(tickerId).receivingMarketDepth = False
End With

addMsgId CANCEL_MKT_DEPTH
addData version, "Version"
addData IdBaseMarketDepthReq + tickerId, "Ticker id"
send
Exit Sub
Err:
handleFatalError ProcName

End Sub

Friend Sub cancelNewsBulletins()
Const ProcName As String = "cancelNewsBulletins"


On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

Const version = 1

addMsgId CANCEL_NEWS_BULLETINS
addData version, "Version"
send
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub cancelOrder(ByVal orderID As Long)
Const ProcName As String = "cancelOrder"


On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const version = 1
addMsgId CANCEL_ORDER
addData version, "Version"
addData orderID, "Order id"
send True
Exit Sub
Err:
handleFatalError ProcName

End Sub

Friend Sub Connect()
Const ProcName As String = "Connect"


On Error GoTo Err
mDisconnectedByApplication = False

If mConnectionState <> ConnectionStates.ConnNotConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Already connected"
End If

createsocket


logMessage "Connecting to TWS: " & ConnectionString, "Connect"
mSocket.Connect mServer, mPort

Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub Disconnect( _
                ByVal reason As String, _
                ByVal forced As Boolean)
Const ProcName As String = "Disconnect"
On Error GoTo Err

If mConnectionState = ConnectionStates.ConnNotConnected Then Exit Sub

mDisconnectedByApplication = True
If Not mConnectionTimer Is Nothing Then mConnectionTimer.StopTimer
mConnectionState = ConnectionStates.ConnNotConnected
releaseAllTickerTableEntries
ReleaseSocket
logMessage "IBTWSSP:TwsAPI:Disconnect: Disconnected from: " & ConnectionString & ": " & reason, "Disconnect"
handleTwsDisconnection
If Not forced Then RaiseEvent Disconnected(reason)
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub GenerateDeferredOrderCancelNotification( _
                ByVal orderID As Long, _
                ByVal clientID As Long)
Dim mapEntry As OrderIdMapEntry

Const ProcName As String = "GenerateDeferredOrderCancelNotification"
Dim failpoint As Long
On Error GoTo Err

On Error Resume Next
mapEntry = mOrderIDMap(generateTwsOrderKey(orderID, clientID))
On Error GoTo Err

If Not mapEntry.order Is Nothing Then DeferAction New DeferredCancelStatus, mapEntry, 500

Exit Sub

Err:
handleFatalError ProcName

End Sub

Friend Function GetNextOrderPlexId() As Long
GetNextOrderPlexId = CStr(mNextOrderPlexID)
mNextOrderPlexID = mNextOrderPlexID + OrderPlexIdIncrement
End Function

Friend Sub placeOrder( _
                ByVal OrderSubmitter As OrderSubmitter, _
                ByVal pOrder As TradeBuildSP.IOrder, _
                ByVal transmit As Boolean, _
                ByVal parentId As String, _
                ByVal ocaGroup As String, _
                ByVal routingHint As String)
Const ProcName As String = "placeOrder"


Dim mapEntry As OrderIdMapEntry

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Already connected"
End If

On Error Resume Next
Set mapEntry.order = pOrder
Set mapEntry.OrderSubmitter = OrderSubmitter
mOrderIDMap.add mapEntry, generateTwsOrderKey(pOrder.BrokerId, mClientID)
On Error GoTo Err

sendOrder pOrder, transmit, parentId, ocaGroup, routingHint
Exit Sub
Err:
handleFatalError ProcName

End Sub
                    
Friend Function ReadHistoricalBar( _
                        ByVal requestHandle As Long, _
                        ByRef pBar As Bar, _
                        ByRef barTimePeriod As TimePeriod) As Boolean
Const ProcName As String = "ReadHistoricalBar"


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
handleFatalError ProcName

End Function

Friend Sub ReplaceFA(ByVal DataType As FADataTypes, ByVal xml As String)
Const ProcName As String = "ReplaceFA"



On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Already connected"
End If

Const version = 1

If mServerVersion < 13 Then
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Server version does not support ReplaceFA"
End If

addMsgId REPLACE_FA
addData version, "Version"
addData DataType, "Data type"
addData xml, "XML"

send
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub requestAccountUpdates(ByVal subscribe As Boolean, ByVal acctCode As String)
Const ProcName As String = "requestAccountUpdates"


On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Already connected"
End If

Const version = 2

addMsgId REQ_ACCT_DATA
addData version, "Version"
addBoolean subscribe, "Subscribe"

If mServerVersion >= 9 Then addData acctCode, "Account code"

send
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub requestAllOpenOrders()
Const ProcName As String = "requestAllOpenOrders"


On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Not connected"
End If
                            

Const version = 1

addMsgId REQ_ALL_OPEN_ORDERS
addData version, "Version"
send
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub requestAutoOpenOrders(ByVal autoBind As Boolean)
Const ProcName As String = "requestAutoOpenOrders"


On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const version = 1

addMsgId REQ_AUTO_OPEN_ORDERS
addData version, "Version"
addBoolean autoBind, "Autobind"
send
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub RequestContract( _
                ByVal contractSpec As ContractSpecifier, _
                ByVal requester As ContractDetailsRequester)
Const ProcName As String = "RequestContract"


Dim requestQueueEntry As TWSContractDetailsRequestQueueEntry

On Error GoTo Err
Set requestQueueEntry.contractSpec = contractSpec
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

If mConnectionState = ConnConnected Then submitNextContractDetailsRequest
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub RequestExecutions(ByVal filter As ExecutionFilter)
Const ProcName As String = "RequestExecutions"

Static reqId As Long
Dim theFilter As ExecutionFilter

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const version = 3

addMsgId REQ_EXECUTIONS
addData version, "Version"

If mServerVersion >= MIN_SERVER_VER_EXECUTION_DATA_CHAIN Then
    addData reqId + IdBaseExecutionsReq, "ReqId"
End If

If mServerVersion >= 9 Then
    If filter Is Nothing Then
        Set theFilter = New ExecutionFilter
    Else
        Set theFilter = filter
    End If
    
    With theFilter
        addData .clientID, "Client id"
        addData .accountCode, "Account code"
        If .time = Int(.time) Then
            addData IIf(CLng(.time) <> 0, Format(.time, "yyyymmdd"), ""), "Fill time"
        Else
            addData IIf(CLng(.time) <> 0, Format(.time, "yyyymmdd-hh\:nn\:ss"), ""), "Fill time"
        End If
        addData .symbol, "Symbol"
        addData secTypeToString(.sectype), "Sec type"
        addData .exchange, "Exchange"
        addData orderActionToString(.action), "Action"
    End With
End If

send
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub requestFA(ByVal DataType As FADataTypes)
Const ProcName As String = "requestFA"


On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const version = 1

If mServerVersion < 13 Then
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Server version does not support requestFA"
End If

addMsgId REQ_FA
addData version, "Version"
addData DataType, "Data type"

send
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub requestManagedAccounts()
Const ProcName As String = "requestManagedAccounts"


On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const version = 1

addMsgId REQ_MANAGED_ACCTS
addData version, "Version"
send
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Function RequestMarketData(ByVal pContract As Contract, _
                            ByVal dataReader As RealtimeDataReader, _
                            ByVal dataConsumer As IStreamingDataConsumer, _
                            Optional ByVal includeMarketDepth As Boolean = False) As Long
Const ProcName As String = "RequestMarketData"


Dim lTickerID As Long

On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
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
handleFatalError ProcName

End Function

Friend Sub RequestMarketDepth(ByVal pTickerID As Long)
Const ProcName As String = "RequestMarketDepth"
                            
On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

With mTickerTable(pTickerID)
    If .Contract Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, "IBTWSSP.TWSAPI::requestMarketDepth", "Invalid ticker id: " & pTickerID
    .receivingMarketDepth = True
    reqMarketDepth pTickerID, .Contract.specifier
End With
Exit Sub
Err:
handleFatalError ProcName

End Sub

Friend Function RequestHistoricalData( _
                barSpecifier As BarDataSpecifier, _
                ByVal historicalDataReader As HistDataReader) As Long
Const ProcName As String = "RequestHistoricalData"


Dim index As Long
Dim i As Long
Dim ignoreTimestamps As Boolean

On Error GoTo Err
index = -1

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

If mServerVersion < 16 Then
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Server version does not support historical data retrieval"
End If

' now adjust the bar specifier to indicate the number of bars of
' appropriate length supported by IB
Select Case barSpecifier.barTimePeriod.Units
Case TimePeriodUnits.TimePeriodDay
    ignoreTimestamps = True
    If barSpecifier.barTimePeriod.length > 1 Then
        barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * barSpecifier.barTimePeriod.length
        barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.Units)
    End If
Case TimePeriodUnits.TimePeriodHour
    If barSpecifier.barTimePeriod.length > 1 Then
        barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * barSpecifier.barTimePeriod.length
        barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.Units)
    End If
Case TimePeriodUnits.TimePeriodMinute
    If barSpecifier.barTimePeriod.length Mod 30 = 0 Then
        barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * (barSpecifier.barTimePeriod.length / 30)
        barSpecifier.barTimePeriod = GetTimePeriod(30, barSpecifier.barTimePeriod.Units)
    ElseIf barSpecifier.barTimePeriod.length Mod 15 = 0 Then
        barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * (barSpecifier.barTimePeriod.length / 15)
        barSpecifier.barTimePeriod = GetTimePeriod(15, barSpecifier.barTimePeriod.Units)
    ElseIf barSpecifier.barTimePeriod.length Mod 5 = 0 Then
        barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * (barSpecifier.barTimePeriod.length / 5)
        barSpecifier.barTimePeriod = GetTimePeriod(5, barSpecifier.barTimePeriod.Units)
    ElseIf barSpecifier.barTimePeriod.length Mod 2 = 0 Then
        barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * (barSpecifier.barTimePeriod.length / 2)
        barSpecifier.barTimePeriod = GetTimePeriod(2, barSpecifier.barTimePeriod.Units)
    ElseIf barSpecifier.barTimePeriod.length > 1 Then
        barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * barSpecifier.barTimePeriod.length
        barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.Units)
    End If
Case TimePeriodUnits.TimePeriodMonth
    ignoreTimestamps = True
    barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * barSpecifier.barTimePeriod.length
    barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.Units)
Case TimePeriodUnits.TimePeriodSecond
    If barSpecifier.barTimePeriod.length Mod 30 = 0 Then
        barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * (barSpecifier.barTimePeriod.length / 30)
        barSpecifier.barTimePeriod = GetTimePeriod(30, barSpecifier.barTimePeriod.Units)
    ElseIf barSpecifier.barTimePeriod.length Mod 15 = 0 Then
        barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * (barSpecifier.barTimePeriod.length / 15)
        barSpecifier.barTimePeriod = GetTimePeriod(15, barSpecifier.barTimePeriod.Units)
    ElseIf barSpecifier.barTimePeriod.length Mod 5 = 0 Then
        barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * (barSpecifier.barTimePeriod.length / 5)
        barSpecifier.barTimePeriod = GetTimePeriod(5, barSpecifier.barTimePeriod.Units)
    ElseIf barSpecifier.barTimePeriod.length > 1 Then
        barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * barSpecifier.barTimePeriod.length
        barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.Units)
    End If
Case TimePeriodUnits.TimePeriodWeek
    ignoreTimestamps = True
    barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * barSpecifier.barTimePeriod.length
    barSpecifier.barTimePeriod = GetTimePeriod(1, barSpecifier.barTimePeriod.Units)
Case TimePeriodUnits.TimePeriodYear
    ignoreTimestamps = True
    barSpecifier.MaxNumberOfBars = barSpecifier.MaxNumberOfBars * barSpecifier.barTimePeriod.length * NumMonthsInYear
    barSpecifier.barTimePeriod = GetTimePeriod(1, TimePeriodMonth)
Case TimePeriodUnits.TimePeriodVolume, TimePeriodUnits.TimePeriodTickMovement
    ' nothing to do
Case Else
    historicalDataReader.HistoricalDataRequestInvalid barSpecifier.MaxNumberOfBars & "specified timeframe not supported"
    Exit Function
End Select

For i = 0 To UBound(mHistoricalDataTable)
    If Not mHistoricalDataTable(i).inUse And _
        Not (mHistoricalDataTable(i).cancelling And Now < mHistoricalDataTable(i).earliestReuseTime) _
    Then
        index = i
        Exit For
    End If
Next

If index = -1 Then
    index = UBound(mHistoricalDataTable) + 1
    ReDim Preserve mHistoricalDataTable(2 * (UBound(mHistoricalDataTable) + 1) - 1) As HistoricalDataRequestTableEntry
End If

mHistoricalDataTable(index).inUse = True
mHistoricalDataTable(index).cancelling = False
Set mHistoricalDataTable(index).barSpecifier = barSpecifier
mHistoricalDataTable(index).ignoreTimestamps = ignoreTimestamps
Set mHistoricalDataTable(index).SessionBuilder = New SessionBuilder
mHistoricalDataTable(index).SessionBuilder.SessionStartTime = barSpecifier.Contract.SessionStartTime
mHistoricalDataTable(index).SessionBuilder.SessionEndTime = barSpecifier.Contract.SessionEndTime
Set mHistoricalDataTable(index).Session = mHistoricalDataTable(index).SessionBuilder.Session
mHistoricalDataTable(index).barType = barSpecifier.barType
mHistoricalDataTable(index).nextBarIndex = -1
Set mHistoricalDataTable(index).historicalDataReader = historicalDataReader

Set mHistoricalDataTable(index).bars = New LinkedList
If barSpecifier.MaxNumberOfBars = 0 Then
    mHistoricalDataTable(index).bars.initialCapacity = 128
ElseIf barSpecifier.MaxNumberOfBars > 512 Then
    mHistoricalDataTable(index).bars.initialCapacity = 512
Else
    mHistoricalDataTable(index).bars.initialCapacity = barSpecifier.MaxNumberOfBars
End If

generateTwsHistoricalDataRequest index
RequestHistoricalData = index
Exit Function
Err:

If Err.number = ErrHistGenerationFailed Then
    releaseHistoricalDataTableEntry index
    RequestHistoricalData = -1
    Exit Function
End If
    
handleFatalError ProcName
End Function

Friend Sub RequestNewsBulletins(allMsgs As Boolean)
Const ProcName As String = "RequestNewsBulletins"


On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const version = 1

addMsgId REQ_NEWS_BULLETINS
addData version, "Version"
addData allMsgs, "All messages"
send
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub RequestOpenOrders()
Const ProcName As String = "RequestOpenOrders"


On Error GoTo Err
If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const version = 1

addMsgId REQ_OPEN_ORDERS
addData version, "Version"
send
Exit Sub
Err:
handleFatalError ProcName
End Sub

Friend Sub StartOrderRecovery( _
                ByVal pOrderRecoveryServiceConsumer As TradeBuildSP.IOrderRecoveryServiceConsumer, _
                ByVal pStartFrom As Date)
Const ProcName As String = "StartOrderRecovery"
Dim execFilter As New ExecutionFilter

Set mOrderRecoveryServiceConsumer = pOrderRecoveryServiceConsumer
RequestOpenOrders

execFilter.time = pStartFrom
RequestExecutions execFilter
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub addBoolean( _
                ByVal val As Boolean, _
                ByRef fieldName As String)
addData IIf(val, 1, 0), fieldName
End Sub

Private Sub addByte(ByVal value As Integer)
If mOutputBufferIndex > UBound(mOutputBuffer) Then
    ReDim Preserve mOutputBuffer(2 * (UBound(mOutputBuffer) + 1) - 1) As Byte
End If
mOutputBuffer(mOutputBufferIndex) = value
mOutputBufferIndex = mOutputBufferIndex + 1
End Sub

Private Sub addData( _
                ByVal data As String, _
                ByRef fieldName As String)
Dim i As Long

If mOutMessageBuilder Is Nothing Then
    Set mOutMessageBuilder = CreateStringBuilder
    mOutMessageBuilder.Append "OUT: "
End If

mOutMessageBuilder.Append fieldName
mOutMessageBuilder.Append "="
mOutMessageBuilder.Append data
mOutMessageBuilder.Append ";"

For i = 1 To Len(data)
    addByte Asc(Mid$(data, i, 1))
Next
addByte 0
End Sub

Private Sub addDouble( _
                ByVal data As String, _
                ByRef fieldName As String)
addData DoubleToString(data), fieldName
End Sub

Private Function adjustPrice( _
                ByVal price As Double, _
                ByVal pContractSpecifier As ContractSpecifier) As Double
If pContractSpecifier.currencyCode = "GBP" And _
    mServerVersion < 23 _
Then
    adjustPrice = price * 100
Else
    adjustPrice = price
End If
End Function

Private Sub addMsgId( _
                ByVal id As TWSSocketOutMsgTypes)
addData id, "Msg Id (" & gOutputMessageIdToString(id) & ")"
End Sub

Private Function AllocateTickerTableEntry(ByVal pDataConsumer As IStreamingDataConsumer) As Long
Dim i As Long

Const ProcName As String = "AllocateTickerTableEntry"
On Error GoTo Err

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
End If

' NB: don't use 'With' in the following as it can lead to an Error 10 table is temporarily locked
mTickerTable(AllocateTickerTableEntry).inUse = True
Set mTickerTable(AllocateTickerTableEntry).dataConsumer = pDataConsumer
mTickerTable(AllocateTickerTableEntry).prevAsk = MaxDouble
mTickerTable(AllocateTickerTableEntry).prevBid = MaxDouble
mTickerTable(AllocateTickerTableEntry).prevClose = MaxDouble
mTickerTable(AllocateTickerTableEntry).prevHigh = MaxDouble
mTickerTable(AllocateTickerTableEntry).prevLow = MaxDouble
mTickerTable(AllocateTickerTableEntry).prevOpen = MaxDouble
mTickerTable(AllocateTickerTableEntry).prevTrade = MaxDouble

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Function

Private Sub cancelMarketDataEx(ByVal index As Long, _
                ByVal releaseEntry As Boolean)

Const ProcName As String = "cancelMarketDataEx"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

Const version = 2

If Not mTickerTable(index).inUse Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid ticker id " & index
End If

If mTickerTable(index).receivingMarketDepth Then CancelMarketDepth index

addMsgId CANCEL_MKT_DATA
addData version, "Version"
addData index + mTickerTable(index).incarnation, "Ticker id"
send

If releaseEntry Then releaseTickerTableEntry index

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub cancelContractRequests(ByVal reason As String)
Dim requestEntry As TWSContractDetailsRequestQueueEntry
Const ProcName As String = "cancelContractRequests"
On Error GoTo Err

Do While mTwsContractDetailsRequestQueue.size <> 0
    requestEntry = mTwsContractDetailsRequestQueue.removeFromFront
    requestEntry.requester.cancelRequest requestEntry, reason
Loop

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub
Private Sub checkmessages()
Const ProcName As String = "checkmessages"
Static msgId As TWSSocketInMsgTypes
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

Dim failpoint As Long
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    ' first data received is plain Server version number
    If mCheckPointNumber = 0 Then
        createInMessageBuilder
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
        
        logSocketInputMessage
        createInMessageBuilder
        
        If mCheckPointNumber = 102 Then
            mServerVersion = getLong("Server version")
            checkpointBuffer
            mCheckPointNumber = 100
        End If
    Loop
    
    logSocketInputMessage
    createInMessageBuilder
    
    If mCheckPointNumber = 100 Then
        If mServerVersion >= 20 Then
            logMessage "TWS time at connection", pMsgQualifier:=getString("TWS time"), pProcName:="checkmessages"
        End If
        checkpointBuffer
        mCheckPointNumber = 200
        'Debug.Print "Set checkpoint number=" & 200
    End If
    
    If mServerVersion < ReqdServerVersion Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "TWS is out of date and needs to be upgraded to at least version 884"
    End If
    
    logSocketInputMessage
    
    mConnectionState = ConnectionStates.ConnReSynching
    logMessage "Connection to TWS successful", pProcName:="checkmessages"
    
    ' set the log level in TWS
    setTWSLogLevel mTWSLogLevel
    
    ' now restart any tickers that were previously in operation
    reEstablishMarketData True
    
    ' and re-request any market depth
    resetMarketDepth True
    
    ' check for pending contract details requests
    submitNextContractDetailsRequest
    
    mConnectionState = ConnectionStates.ConnConnected
    RaiseEvent Connected
    
    mCheckPointNumber = 1000
    'Debug.Print "Set checkpoint number=" & 1000
End If

Do While mInputParseIndex < mInputBufferNextFreeIndex
    
    If mCheckPointNumber = 1000 Then
        Dim tmpStartIndex As Long
        tmpStartIndex = startIndex
        startIndex = mInputParseIndex
        prevMsgID = msgId
        
        createInMessageBuilder
    
        msgId = getLong("Msg id")
        mInMessageBuilder.Append "("
        mInMessageBuilder.Append gInputMessageIdToString(msgId)
        mInMessageBuilder.Append ") "
        
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
    
    Select Case msgId
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
    Case TWSSocketInMsgTypes.TICK_OPTION_COMPUTATION
        getTickOptionComputation version
    Case TWSSocketInMsgTypes.TICK_GENERIC
        getTickGeneric version
    Case TWSSocketInMsgTypes.TICK_STRING
        getTickString version
    Case TWSSocketInMsgTypes.CONTRACT_DATA_END
        getContractEnd version
    Case TWSSocketInMsgTypes.OPEN_ORDER_END
        getOpenOrderEnd version
    Case TWSSocketInMsgTypes.ACCT_DOWNLOAD_END
        getAccountDownloadEnd version
    Case TWSSocketInMsgTypes.EXECUTION_DATA_END
        getExecutionDataEnd version
    Case TWSSocketInMsgTypes.DELTA_NEUTRAL_VALIDATION
        getDeltaNeutralValidation version
    Case TWSSocketInMsgTypes.TICK_SNAPSHOT_END
        getTickSnapshotEnd version
    Case Else
        s = "Invalid message code from TWS: " & vbCrLf & _
            "msgID=" & msgId & _
            " version=" & version & _
            " startindex=" & startIndex & vbCrLf & _
            "prevMsgID=" & prevMsgID & _
            " prevVersion=" & prevVersion & _
            " prevStartindex=" & prevStartIndex & vbCrLf & _
            "mIndexIn=" & mInputBufferNextFreeIndex & _
            " mInputIndex=" & mInputParseIndex & vbCrLf & _
            "Buffer contents: " & vbCrLf & _
            formatBuffer
        
        Debug.Print s
        logMessage s, pLogLevel:=LogLevelSevere, pProcName:="checkmessages"
        gReleaseTWSAPIInstance Me, True
        Exit Sub
    End Select
    
    retryCount = 0
    
    Dim et As Single
    et = mPerformanceElapsedTimer.ElapsedTimeMicroseconds
    With mPerformanceStats(msgId)
        .lastSecondCount = .lastSecondCount + 1
        .lastSecondTime = .lastSecondTime + et
        If et > .longestTime Then .longestTime = et
        If et < .shortestTime Or .shortestTime = 0 Then .shortestTime = et
    End With
    
    eventCount = eventCount + 1
    
    If mElapsedTimer.ElapsedTimeMicroseconds >= 10000000 Then
        logMessage "Event rate per second", pMsgQualifier:=Format(eventCount / 10, "0.0"), pLogLevel:=LogLevelDetail, pProcName:="checkmessages"
        Debug.Print "Event rate per second = " & Format(eventCount / 10, "0.0")
        eventCount = 0
        mElapsedTimer.StartTiming
    End If
    
    checkpointBuffer
    mCheckPointNumber = 1000
    'Debug.Print "Set checkpoint number=" & 1000 & " minputindex=" & mInputParseIndex
Loop

Exit Sub

Err:
Dim errNum As Long
Dim errDesc As String
Dim errSource As String

errNum = Err.number
errDesc = Err.Description
errSource = Err.source

If errNum = DataIncomplete Then
    rollbackBuffer
    retryCount = retryCount + 1
    Debug.Print "Data incomplete count: " & retryCount & " mCheckPointNumber: " & mCheckPointNumber
    Exit Sub
End If

s = "msgID=" & msgId & _
    " version=" & version & _
    " startindex=" & startIndex & vbCrLf & _
    "prevMsgID=" & prevMsgID & _
    " prevVersion=" & prevVersion & _
    " prevStartindex=" & prevStartIndex & vbCrLf & _
    "mInputBufferNextFreeIndex=" & mInputBufferNextFreeIndex & _
    " mInputParseIndex=" & mInputParseIndex & vbCrLf & _
    "Buffer contents: " & vbCrLf & _
    formatBuffer

Debug.Print "Error " & errNum & ": " & errDesc & vbCrLf & s

logMessage s, pLogLevel:=LogLevelDetail, pProcName:="checkmessages"
logSocketInputMessage True
handleFatalErrorEx errNum, errSource, errDesc, "checkMessages"

End Sub

Private Sub checkpointBuffer()
mInputParseIndexChk = mInputParseIndex
End Sub

Private Sub clearPrevPrices( _
                ByVal index As Long)
Const ProcName As String = "clearPrevPrices"
On Error GoTo Err

With mTickerTable(index)
    .prevBid = 0#
    .prevBidSize = 0&
    .prevAsk = 0#
    .prevAskSize = 0&
    .prevTrade = 0#
    .prevTradeSize = 0&
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub createLoggers()

Const ProcName As String = "createLoggers"
On Error GoTo Err

If gLogger.IsLoggable(LogLevelMediumDetail) Then
    If mPerformanceTimerSecond Is Nothing Then
        Set mPerformanceTimerSecond = CreateIntervalTimer(CDate(Int((Now + 2 * OneSecond) / OneSecond) * OneSecond), _
                                                ExpiryTimeUnitDateTime, _
                                                1000)
        Set mPerformanceTimerPeriod = CreateIntervalTimer(CDate(Int((Now + OneMinute) / OneMinute) * OneMinute), _
                                                ExpiryTimeUnitDateTime, _
                                                60000)
        mPerformanceTimerSecond.StartTimer
        mPerformanceTimerPeriod.StartTimer
        
        Set mPerformanceLogger = GetLogger("tradebuild.log.serviceprovider.ibtwssp.performance")
    End If
End If

Set mSocketLogger = GetLogger("tradebuild.log.serviceprovider.ibtwssp.socket")

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub createInMessageBuilder()
'If mInMessageBuilder Is Nothing And mSocketLogger.isLoggable(LogLevelHighDetail) then
Const ProcName As String = "createInMessageBuilder"
On Error GoTo Err

If mInMessageBuilder Is Nothing Then
    Set mInMessageBuilder = CreateStringBuilder
    mInMessageBuilder.Append "IN: "
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub createsocket()
Const ProcName As String = "createsocket"
On Error GoTo Err

If mWinsockForm Is Nothing Then Set mWinsockForm = New WinsockForm
Set mSocket = mWinsockForm.Winsock1

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Function expiryToDate( _
                ByVal expiry As String) As Date
Const ProcName As String = "expiryToDate"
On Error GoTo Err

If Len(expiry) = 8 Then
    expiryToDate = CDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2))
ElseIf Len(expiry) = 6 Then
    expiryToDate = CDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & "01")
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

Private Function formatBuffer() As String
Dim s As StringBuilder
Dim i As Long
Dim j As Long

Const ProcName As String = "formatBuffer"
On Error GoTo Err

Set s = CreateStringBuilder
Do While i < mInputBufferNextFreeIndex
    s.Append Format(i, "0000  ")
    For j = i To i + 49
        If j = mInputBufferNextFreeIndex Then Exit For
        s.Append IIf(mInputBuffer(j) <> 0, Chr$(mInputBuffer(j)), "_")
    Next
    i = i + 50
    If j < mInputBufferNextFreeIndex Then s.AppendLine ""
Loop
formatBuffer = s.ToString

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

Private Function generateHistDataRequestId(ByVal index As Long) As Long
Const ProcName As String = "generateHistDataRequestId"
On Error GoTo Err

generateHistDataRequestId = index + IdBaseReqHistoricalData + (mHistoricalDataTable(index).numberOfTwsRequests - 1) * HistoricalDataReqIdIncrement

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

' Raises an ErrHistGenerationFailed error if IB cannot satisfy the request because
' it would start more than a year and a day ago.
Private Sub generateTwsHistoricalDataRequest(ByVal index As Long)
Dim TwsRequest As TWSHistoricalDataRequestQueueEntry
Dim reqNum As Long
Dim numBarsInsession As Long

Const ProcName As String = "generateTwsHistoricalDataRequest"
On Error GoTo Err

reqNum = mHistoricalDataTable(index).numberOfTwsRequests + 1
mHistoricalDataTable(index).numberOfTwsRequests = reqNum

TwsRequest.id = generateHistDataRequestId(index)

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
                                                                            mHistoricalDataTable(index).barSpecifier.Contract.Timezone)), _
                                    "yyyymmdd hh\:nn\:ss")
    
    With mHistoricalDataTable(index).barSpecifier
        
        Select Case .barTimePeriod.Units
        Case TimePeriodUnits.TimePeriodDay
            TwsRequest.barSizeSetting = "1 day"
            If .MaxNumberOfBars >= NumDaysInYear Then
                If DateAdd("d", 1, .ToDate) > Now Then
                    TwsRequest.duration = "1 Y"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 60) & " D"
                End If
            ElseIf .MaxNumberOfBars <= 60 Then
                TwsRequest.duration = permittedHistDays(.ToDate, .MaxNumberOfBars) & " D"
            ElseIf .MaxNumberOfBars <= 52 * NumDaysInWeek Then
                TwsRequest.duration = permittedHistWeeks(.ToDate, Int(.MaxNumberOfBars / NumDaysInWeek)) & " W"
            Else
                TwsRequest.duration = permittedHistDays(.ToDate, 60) & " D"
            End If
        Case TimePeriodUnits.TimePeriodHour
            TwsRequest.barSizeSetting = "1 hour"
            
            numBarsInsession = mHistoricalDataTable(index).Session.NumberOfBarsInSession( _
                                                                    .barTimePeriod)
            
            If .MaxNumberOfBars <= 24 Then
                TwsRequest.duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 3600) & " S"
            ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 34 Then
                TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
            Else
                TwsRequest.duration = permittedHistDays(.ToDate, 34) & " D"
            End If
        Case TimePeriodUnits.TimePeriodMinute
            
            numBarsInsession = mHistoricalDataTable(index).Session.NumberOfBarsInSession( _
                                                                    .barTimePeriod)
            
            Select Case .barTimePeriod.length
            Case 1
                TwsRequest.barSizeSetting = "1 min"
                If .MaxNumberOfBars <= 1440 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 60) & " S"
                ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 6 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 6) & " D"
                End If
            Case 2
                TwsRequest.barSizeSetting = "2 mins"
                If .MaxNumberOfBars <= 720 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 120) & " S"
                ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 6 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 6) & " D"
                End If
            Case 5
                TwsRequest.barSizeSetting = "5 mins"
                If .MaxNumberOfBars <= 288 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 300) & " S"
                ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 6 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 6) & " D"
                End If
            Case 15
                TwsRequest.barSizeSetting = "15 mins"
                If .MaxNumberOfBars <= 96 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 900) & " S"
                ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 20 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 20) & " D"
                End If
            Case 30
                TwsRequest.barSizeSetting = "30 mins"
                If .MaxNumberOfBars <= 48 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 1800) & " S"
                ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 3 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                ElseIf -Int(-.MaxNumberOfBars / numBarsInsession) <= 34 Then
                    TwsRequest.duration = permittedHistDays(.ToDate, -Int(-.MaxNumberOfBars / numBarsInsession)) & " D"
                Else
                    TwsRequest.duration = permittedHistDays(.ToDate, 34) & " D"
                End If
            End Select
        Case TimePeriodUnits.TimePeriodMonth
            TwsRequest.barSizeSetting = "1 day"
            If .MaxNumberOfBars <= 12 Then
                TwsRequest.duration = permittedHistMonths(.ToDate, .MaxNumberOfBars) & " M"
            Else
                If DateAdd("d", 1, .ToDate) > Now Then
                    TwsRequest.duration = "12 M"
                Else
                    TwsRequest.duration = permittedHistMonths(.ToDate, .MaxNumberOfBars) & " M"
                End If
            End If
        Case TimePeriodUnits.TimePeriodSecond
            
            numBarsInsession = mHistoricalDataTable(index).Session.NumberOfBarsInSession( _
                                                                    .barTimePeriod)
            
            Select Case .barTimePeriod.length
            Case 1
                
                If .MaxNumberOfBars <= 2000 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 2000) & " S"
                End If
            Case 5
                TwsRequest.barSizeSetting = "5 secs"
                If .MaxNumberOfBars <= 2000 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 5) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 10000) & " S"
                End If
            Case 15
                TwsRequest.barSizeSetting = "15 secs"
                If .MaxNumberOfBars <= 2000 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 15) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 30000) & "S"
                End If
            Case 30
                TwsRequest.barSizeSetting = "30 secs"
                If .MaxNumberOfBars <= 2880 Then
                    TwsRequest.duration = permittedHistSeconds(.ToDate, .MaxNumberOfBars * 30) & " S"
                Else
                    TwsRequest.duration = permittedHistSeconds(.ToDate, 86400) & " S"
                End If
            End Select
        Case TimePeriodUnits.TimePeriodWeek
            TwsRequest.barSizeSetting = "1 day"
            If .MaxNumberOfBars <= 52 Then
                TwsRequest.duration = permittedHistWeeks(.ToDate, .MaxNumberOfBars) & " W"
            Else
                If DateAdd("d", 1, .ToDate) > Now Then
                    TwsRequest.duration = "52 W"
                Else
                    TwsRequest.duration = permittedHistWeeks(.ToDate, .MaxNumberOfBars) & " W"
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

mHistoricalDataTable(index).requestKey = queueTWSHistoricalDataRequest(TwsRequest)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
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
    getDouble = val(s)
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

For i = mInputParseIndex To mInputBufferNextFreeIndex - 1
    If mInputBuffer(i) = 0 Then
        ' we've found our next string
        'Debug.Print "Got string: mInputIndex=" & mInputParseIndex & " value=" & s
        mInputParseIndex = i + 1
        getString = s
        If Not mInMessageBuilder Is Nothing Then
            mInMessageBuilder.Append fieldName
            mInMessageBuilder.Append "="
            mInMessageBuilder.Append s
            mInMessageBuilder.Append ";"
        End If
        Exit Function
    End If
    ' NB: I've tried various other means of generating the string
    ' without using string concatenation, but for the lengths of the
    ' fields in the TWSAPI protocol (except for error messages) they are all much slower
    ' than string concatenation
    s = s & Chr$(mInputBuffer(i))
Next

' we've got to the end of the current data, so raise an error to abort
' this attempt to assemble a message
'Debug.Print "Failed to get string"
Err.Raise DataIncomplete
End Function

Private Sub getAccountDownloadEnd(ByVal version As Long)
Dim accountName As String
Const ProcName As String = "getAccountDownloadEnd"
Dim failpoint As String
On Error GoTo Err

accountName = getString("Account name")
logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getAccountTime(ByVal version As Long)
Const ProcName As String = "getAccountTime"
Dim accountTime As String: accountTime = getString("Account time")

Dim failpoint As Long
On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getAccountValue(ByVal version As Long)
Const ProcName As String = "getAccountValue"
Dim key As String: key = getString("Key")
Dim val As String: val = getString("Value")
Dim cur As String: cur = getString("Currency")
Dim accountName As String: If version >= 2 Then accountName = getString("Account Name")

Dim failpoint As Long
On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getBondContractData(ByVal version As Long)
' not currently supported
Const ProcName As String = "getBondContractData"
Dim failpoint As Long
On Error GoTo Err

If (version >= 3) Then getLong "Req Id"
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
If (version >= 2) Then
    getString "Next Option Date"
    getString "Next Option Type"
    getBoolean "Next Option Partial"
    getString "Notes"
End If
If version >= 4 Then getString "longName"


logSocketInputMessage

mContractDetailsRequest.requester.contractSpecifierInvalid mContractDetailsRequest, _
                                                    "bonds not yet supported"

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getContractData(ByVal version As Long)
Const ProcName As String = "getContractData"
Dim lContractWrapper As ContractWrapper
Dim lContractSpec As ContractSpecifier

Dim reqId As Long

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

Dim failpoint As Long
On Error GoTo Err

If (version >= 3) Then reqId = getLong("Req Id")

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
marketName = getString("Market Name")
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
    logMessage "Unknown exchange code received from TWS", pMsgQualifier:=exchange, pProcName:="getContractData"
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
        Case "ACTIVETIM", _
            "ALLOC", _
            "ALGO", _
            "AVGCOST", _
            "AUC", _
            "BASKET", _
            "COND", _
            "CONDORDER", _
            "CONSCOST", _
            "DARKPOLL", _
            "DEACT", _
            "DEACTDIS", _
            "DEACTEOD", _
            "FOK", _
            "HPENNY", _
            "LARGE", _
            "LTH"
        Case "NONALGO", _
            "OCA", _
            "OPG", _
            "OPGREROUT", _
            "PAON", _
            "PEGMID", _
            "POSTONLY", _
            "PREOPGRTH", _
            "RELSTK", _
            "RTHIGNOPG", _
            "SCALE", _
            "SCALERST", _
            "SMARTSTG"
        Case "SOI", _
            "TIMEPRIO", _
            "TRAILLIT", _
            "TRAILMIT", _
            "VOLAT", _
            "WHATIF"
            ' not sure what these signify
        Case Else
            If orderTypesStr(i) <> "" Then
                logMessage "Unknown order type received from TWS", pMsgQualifier:=orderTypesStr(i), pProcName:="getContractData"
            End If
        End Select
    End If
Next

lContractWrapper.permittedOrderAttributes = permittedOrderAttributes
lContractWrapper.permittedOrderTypes = permittedOrderTypes

lContractWrapper.permittedOrderTifs = permittedOrderTifs

lContractWrapper.permittedStopTriggerMethods = TradeBuildSP.StopTriggerMethods.StopTriggerBidAsk Or _
                                        TradeBuildSP.StopTriggerMethods.StopTriggerDefault Or _
                                        TradeBuildSP.StopTriggerMethods.StopTriggerDoubleBidAsk Or _
                                        TradeBuildSP.StopTriggerMethods.StopTriggerDoubleLast Or _
                                        TradeBuildSP.StopTriggerMethods.StopTriggerLast Or _
                                        TradeBuildSP.StopTriggerMethods.StopTriggerLastOrBidAsk Or _
                                        TradeBuildSP.StopTriggerMethods.StopTriggerMidPoint

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

If version >= 4 Then getLong "UnderConId"

If version >= 5 Then
   lContractBuilder.Description = getString("Long name")
   getString "Primary exchange"
End If

logSocketInputMessage

If Not ignore Then
    Dim expiryDate As Date
    If expiry <> "" Then
        expiryDate = CDate(Left$(expiry, 4) & "/" & _
                                            Mid$(expiry, 5, 2) & "/" & _
                                            Right$(expiry, 2))
    End If
    
    lContractBuilder.expiryDate = expiryDate
    
    If lContractBuilder.Contract.Description = "" Then
        lContractBuilder.Description = symbol & _
                                    " (" & SecTypeToShortString(sectype) & ")" & _
                                    IIf(expiryDate <> 0, Format(expiryDate, " dd mmm yy"), "") & _
                                    IIf(optRight <> OptNone, " " & optionRightToString(optRight), "") & _
                                    IIf(strike <> 0#, " " & strike, "")
    End If
    
    lContractBuilder.multiplier = multiplier / lContractWrapper.priceMagnifier
    lContractBuilder.TickSize = minimumTick * lContractWrapper.priceMagnifier

    mContractDetailsRequest.contractWrappers.add lContractWrapper
End If

If mServerVersion < MIN_SERVER_VER_CONTRACT_DATA_CHAIN Then
    If Not mFirstContractDetailsReceived Then
        mFirstContractDetailsReceived = True
        
        ' now send a reqContract that will give a 'no such security
        ' exists' error. This error will act as a delimiter for all
        ' the contract details returned by the orginal request. Note that
        ' tests indicate that contract details for different reqContracts
        ' are not interleaved
        
        addMsgId REQ_CONTRACT_DATA
        addData 1, "Version"
        addData "ES", "Symbol"
        addData "FUT", "Sec type"
        addData "200501", "Expiry"
        addData "0", "Strike"
        addData "", "Right"
        addData "LIFFE", "Exchange"
        addData "GBP", "Currency"
        addData "", "Local symbol"
        send
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getContractEnd(ByVal version As Long)
Const ProcName As String = "getContractEnd"
Dim reqId As Long

Dim failpoint As Long
On Error GoTo Err

reqId = getLong("Req Id")

logSocketInputMessage

mContractDetailsRequest.requester.contractsLoaded mContractDetailsRequest

Set mContractDetailsRequest.contractSpec = Nothing
Set mContractDetailsRequest.contractWrappers = Nothing
Set mContractDetailsRequest.requester = Nothing
mFirstContractDetailsReceived = False

If mConnectionState = ConnConnected Then submitNextContractDetailsRequest

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getDeltaNeutralValidation(ByVal version As Long)
Dim reqId As Long: reqId = getLong("reqID")

Dim conId As Long: conId = getLong("conID")
Dim delta As Double: delta = getDouble("delta")
Dim price As Double: price = getDouble("price")

Const ProcName As String = "getDeltaNeutralValidation"
Dim failpoint As String
On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getErrorMsg(ByVal version As Long)
Dim id As Long
Dim errorCode As Long
Dim errorMsg As String
Dim execFilter As ExecutionFilter
Dim index As Long
Dim mapEntry As OrderIdMapEntry

Dim failpoint As Long
Const ProcName As String = "getErrorMsg"

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

logSocketInputMessage True

Select Case errorCode
Case 101
    ' max tickers reached
    failpoint = 350
    mTickerTable(id Mod IncarnationIncrement).dataReader.marketDataRequestFailed "Maximum number of tickers reached"
Case 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, _
    121, 125, 126, 127, 128, 129, 130, 132, 137, 140, 141, 144, _
    152, 153, 154, 157, 160, 201, 312, 313, 314, 315, 325, 328, _
    334, 337, 338, 339, 343, 355, 361, 382, 383, 387, 388, 391, 395, 396, 397, 451
    
    failpoint = 400

    On Error Resume Next
    mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
    On Error GoTo Err
    If Not mapEntry.order Is Nothing Then
        mapEntry.OrderSubmitter.InvalidOrder mapEntry.order.tradebuildId, errorMsg
    End If
Case 162    ' historical data request problem
    failpoint = 500

    index = getHistDataTableIndexFromRequestId(id)
    
    If mHistoricalDataTable(index).bars Is Nothing Then
        ' might get here if the request has already been cancelled
    ElseIf mHistoricalDataTable(index).bars.size <> 0 Then
        
        failpoint = 505
        
        Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
        mHistoricalDataTable(index).dataRetrievedFromTWS = True
        mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
        
        failpoint = 510
        
        mHistoricalDataTable(index).historicalDataReader.HistoricalDataAvailable mHistoricalDataTable(index).bars.size
    Else
        
        failpoint = 515
        
        logMessage "Error " & errorCode & " from TWS: " & errorMsg, pProcName:="getErrorMsg"
        
        failpoint = 520
        
        mHistoricalDataTable(index).historicalDataReader.HistoricalDataRequestFailed errorMsg
        releaseHistoricalDataTableEntry index
    End If
    
    processNextHistDataRequest
Case 165
    failpoint = 600

    logMessage "Connected to IB Historical Market Data Service - " & errorMsg, pProcName:="getErrorMsg"
Case 200    ' security definition not known
    Select Case getIdType(id)
    Case IdTypeNone, IdTypeContractData
        ' this must relate to a RequestContract
        failpoint = 750

        processContractError errorMsg
    Case IdTypeRealtimeData
        ' this must relate to a reqMktData - this can occur even though
        ' we make sure the contract is valid before starting the ticker:
        ' for example the contract may have expired while waiting for TWS
        ' to be restarted
        
        failpoint = 700

        mTickerTable(id Mod IncarnationIncrement).dataReader.marketDataRequestFailed errorMsg
        releaseTickerTableEntry (id Mod IncarnationIncrement)
    Case IdTypeMarketDepth
        failpoint = 710

        mTickerTable(id - IdBaseMarketDepthReq).dataReader.marketDepthRequestFailed errorMsg
    Case IdTypeHistoricalData
        failpoint = 720

        index = getHistDataTableIndexFromRequestId(id)
    
        logMessage "Error " & errorCode & " from TWS: " & errorMsg, pProcName:="getErrorMsg"
        mHistoricalDataTable(index).historicalDataReader.HistoricalDataRequestFailed errorMsg
        releaseHistoricalDataTableEntry index
        
        processNextHistDataRequest
    Case IdTypeOrder
        failpoint = 730

        On Error Resume Next
        mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
        On Error GoTo Err
        If Not mapEntry.order Is Nothing Then
            mapEntry.OrderSubmitter.InvalidOrder mapEntry.order.tradebuildId, errorMsg
        End If
    End Select
    failpoint = 800

    On Error Resume Next
    mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
    On Error GoTo Err
    If Not mapEntry.order Is Nothing Then
        Dim osp As New OrderStatusReport
        osp.Initialise mapEntry.order.tradebuildId, OrderStatuses.OrderStatusCancelled
        mapEntry.OrderSubmitter.OrderStatusReport osp
    End If
Case 309
    ' max market depth requests exceeded
    failpoint = 900

    mTickerTable(id - IdBaseMarketDepthReq).dataReader.marketDepthRequestFailed "Max market depth requests exceeded"
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
        If mServerVersion < MIN_SERVER_VER_CONTRACT_DATA_CHAIN And _
            Not mContractDetailsRequest.contractSpec Is Nothing _
        Then
            ' this probably relates to a RequestContract
            processContractError errorMsg
        Else
            logMessage "Error (" & errorCode & "; id=" & id & ") from TWS: " & errorMsg, pProcName:="getErrorMsg"
            mCommonServiceConsumer.NotifyRecoverableError errorCode, errorMsg, -1
        End If
    Case IdTypeContractData
        processContractError errorMsg
    Case IdTypeRealtimeData
    Case IdTypeMarketDepth
    Case IdTypeHistoricalData
        index = getHistDataTableIndexFromRequestId(id)
    
        If mHistoricalDataTable(index).bars.size <> 0 Then
            Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
            mHistoricalDataTable(index).dataRetrievedFromTWS = True
            mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
            mHistoricalDataTable(index).historicalDataReader.HistoricalDataAvailable mHistoricalDataTable(index).bars.size
        Else
            logMessage "Error " & errorCode & " from TWS: " & errorMsg, pProcName:="getErrorMsg"
            mHistoricalDataTable(index).historicalDataReader.HistoricalDataRequestFailed errorMsg
            releaseHistoricalDataTableEntry index
        End If
        processNextHistDataRequest
    Case IdTypeOrder
    End Select
Case 326    ' ClientID already in use

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
    
Case 399
    
    failpoint = 100

    On Error Resume Next
    mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
    On Error GoTo Err
    If Not mapEntry.order Is Nothing Then
        mapEntry.OrderSubmitter.OrderDeferred mapEntry.order.tradebuildId, errorMsg
    End If

Case 1100
    failpoint = 1300

    ' connectivity between TWS and IB has been lost
    RaiseEvent ConnectionToIBClosed
    logMessage "Connection to IB has been lost", pProcName:="getErrorMsg"
    
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
    
    logMessage "Connection to IB recovered: market data re-established", pProcName:="getErrorMsg"
    
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
    
    logMessage "Connection to IB recovered: no loss of data", pProcName:="getErrorMsg"
    
Case 2103, 2104, 2105, 2106, 2107, 2108
    failpoint = 1600

    logMessage errorMsg, pProcName:="getErrorMsg"
Case Else
    failpoint = 1700

    logMessage "Error (" & errorCode & "; id=" & id & ") from TWS: " & errorMsg, pProcName:="getErrorMsg"
    mCommonServiceConsumer.NotifyRecoverableError errorCode, errorMsg, -1
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pFailpoint:=failpoint
End Sub

Private Sub getExecutionData(ByVal version As Long)
Const ProcName As String = "getExecutionData"
Dim timeString  As String

Dim reqId As Long

Dim symbol As String
Dim localSymbol As String
Dim sectype As String ' keep the socket string for logging
Dim expiry As String
Dim strike As Double
Dim rightString As String
Dim exchange As String
Dim currencyCode As String

Dim ExecId As String
Dim side As String
Dim price As Double
Dim mapEntry As OrderIdMapEntry
Dim id As Long
Dim permId As Long
Dim clientID As Long
Dim lContractSpecifier As ContractSpecifier
Dim execTime As String
Dim accountId As String
Dim fillingExchg  As String
Dim qty As Long
Dim IsLiquidation As Boolean
Dim tradebuildId As String

Dim execRpt As New ExecutionReport

Dim failpoint As Long

On Error GoTo Err

If (version >= 7) Then reqId = getLong("ReqId")

id = getLong("Id")

On Error Resume Next
mapEntry = mOrderIDMap(generateTwsOrderKey(id, mClientID))
On Error GoTo Err

If Not mapEntry.order Is Nothing Then tradebuildId = mapEntry.order.tradebuildId

If (version >= 5) Then getLong "Con Id"
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

ExecId = getString("Exec id")
timeString = getString("Exec time")
execTime = CDate(Left$(timeString, 4) & "/" & _
                    Mid$(timeString, 5, 2) & "/" & _
                    Mid$(timeString, 7, 2) & " " & _
                    Right$(timeString, 8))
accountId = getString("Account id")
fillingExchg = getString("Filling exchange")
side = getString("Side")
qty = getLong("Quantity")
price = getDouble("Price")
price = adjustPrice(price, lContractSpecifier)
    
If version >= 2 Then permId = getLong("Perm Id")

If version >= 3 Then clientID = getLong("Client id")  ' ClientID
   
If version >= 4 Then IsLiquidation = getBoolean("Is liquidation")

If (version >= 6) Then
    getLong "Cum Qty"
    getDouble "Avg Price"
End If

logSocketInputMessage True

execRpt.Initialise lContractSpecifier.localSymbol & "." & lContractSpecifier.exchange, _
                    accountId, _
                    IIf(UCase$(side) = "BOT", OrderActions.ActionBuy, OrderActions.ActionSell), _
                    ExecId, _
                    fillingExchg, _
                    IsLiquidation, _
                    id, _
                    tradebuildId, _
                    price, _
                    qty, _
                    execTime, _
                    GetTimeZone("").DisplayName

Dim existingExecRpt As ExecutionReport
On Error Resume Next
Set existingExecRpt = mExecutions(ExecId)
On Error GoTo Err

If Not existingExecRpt Is Nothing Then
    ' we've had this execution before so nothing to do
Else
    mExecutions.add execRpt, ExecId
    If Not mapEntry.OrderSubmitter Is Nothing Then
        Dim qtyRemaining As Long
        Dim statusRpt As New OrderStatusReport
        
        qtyRemaining = mapEntry.order.QuantityRemaining
        
        mapEntry.OrderSubmitter.ExecutionReport execRpt
        
        ' notify a status update - this is in case this execution occurred during a period when
        ' TWS had lost its connection to the IB servers and this exec details is being notified
        ' in response to the reqExecutions call made when the connection is recovered. In these
        ' circumstances TWS will not itself issue an orderStatus.
        If qty >= qtyRemaining Then
            statusRpt.Initialise mapEntry.order.tradebuildId, _
                                OrderStatuses.OrderStatusFilled
            mapEntry.OrderSubmitter.OrderStatusReport statusRpt
        Else
            statusRpt.Initialise mapEntry.order.tradebuildId, _
                                OrderStatuses.OrderStatusSubmitted
            mapEntry.OrderSubmitter.OrderStatusReport statusRpt
        End If
    Else
        ' this must be an execution from a previous session
        If Not mOrderRecoveryServiceConsumer Is Nothing Then
            mOrderRecoveryServiceConsumer.ExecutionReport execRpt
        End If
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pFailpoint:=failpoint
End Sub

Private Sub getExecutionDataEnd(ByVal version As Long)
Dim reqId  As Long
Const ProcName As String = "getExecutionDataEnd"
Dim failpoint As String
On Error GoTo Err

reqId = getLong("reqId")
logSocketInputMessage

If Not mOrderRecoveryServiceConsumer Is Nothing Then mOrderRecoveryServiceConsumer.ExecutionReportsComplete
Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getFAData(ByVal version As Long)
Const ProcName As String = "getFAData"
Dim DataType As FADataTypes: DataType = getLong("Data type")
Dim XMLData As String: XMLData = getString("XML data")

Dim failpoint As Long
On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getHistoricalData(ByVal version As Long)
Const ProcName As String = "getHistoricalData"
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
    index = (requestID - IdBaseReqHistoricalData) Mod HistoricalDataReqIdIncrement
    reqNo = Int((requestID - IdBaseReqHistoricalData) / HistoricalDataReqIdIncrement)
    
    If mHistoricalDataTable(index).barSpecifier.CustomSessionStartTime <> 0 Or _
        mHistoricalDataTable(index).barSpecifier.CustomSessionEndTime <> 0 _
    Then
        Set lCustomSessionBuilder = New SessionBuilder
        lCustomSessionBuilder.SessionStartTime = mHistoricalDataTable(index).barSpecifier.CustomSessionStartTime
        lCustomSessionBuilder.SessionEndTime = mHistoricalDataTable(index).barSpecifier.CustomSessionEndTime
        lCustomSessionBuilder.Timezone = mHistoricalDataTable(index).barSpecifier.Contract.Timezone
    Else
        Set lCustomSessionBuilder = Nothing
    End If
    
    checkpointBuffer
    checkPointNumber = 100
End If

If checkPointNumber = 100 Then
    If version >= 2 Then
        getString "Start date"      ' startDateStr
    End If
    checkpointBuffer
    checkPointNumber = 110
End If

If checkPointNumber = 110 Then
    If version >= 2 Then
        getString "End date"      ' endDateStr
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
                                    mHistoricalDataTable(index).barSpecifier.Contract.Timezone)
            If mHistoricalDataTable(index).ignoreTimestamps Then
                Bar.timestamp = Bar.timestamp + mHistoricalDataTable(index).barSpecifier.Contract.SessionStartTime
            End If
            
            If i = 0 Then
                If mHistoricalDataTable(index).barSpecifier.barTimePeriod.Units = TimePeriodVolume Or _
                    mHistoricalDataTable(index).barSpecifier.barTimePeriod.Units = TimePeriodTickMovement _
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
            mHistoricalDataTable(index).SessionBuilder.SetSessionCurrentTime Bar.timestamp
            If Bar.timestamp < mHistoricalDataTable(index).barSpecifier.FromDate Then
                ' ignore this bar
            ElseIf Not lCustomSessionBuilder Is Nothing Then
                lCustomSessionBuilder.SetSessionCurrentTime Bar.timestamp
                If lCustomSessionBuilder.Session.IsTimeInSession(Bar.timestamp) Then
                    If mHistoricalDataTable(index).barSpecifier.IncludeBarsOutsideSession Or _
                        mHistoricalDataTable(index).Session.IsTimeInSession(Bar.timestamp) Or _
                        mHistoricalDataTable(index).ignoreTimestamps _
                    Then
                        hint = mHistoricalDataTable(index).bars.add(Bar, Bar.timestamp, hint)
                    End If
                End If
            Else
                If mHistoricalDataTable(index).barSpecifier.IncludeBarsOutsideSession Or _
                    mHistoricalDataTable(index).Session.IsTimeInSession(Bar.timestamp) Or _
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

logSocketInputMessage

If mHistoricalDataTable(index).cancelling Then Exit Sub

mHistoricalDataTable(index).requestKey = ""

processNextHistDataRequest

failpoint = 500

If mHistoricalDataTable(index).barSpecifier.barTimePeriod.Units = TimePeriodVolume Then
    mHistoricalDataTable(index).barSpecifier.MaxNumberOfBars = mHistoricalDataTable(index).barSpecifier.MaxNumberOfBars - Int(totalVolume / mHistoricalDataTable(index).barSpecifier.barTimePeriod.length)
Else
    mHistoricalDataTable(index).barSpecifier.MaxNumberOfBars = mHistoricalDataTable(index).barSpecifier.MaxNumberOfBars - itemCount
End If
If mHistoricalDataTable(index).barSpecifier.MaxNumberOfBars <= 0 Then
    
    failpoint = 600
    
    Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
    mHistoricalDataTable(index).dataRetrievedFromTWS = True
    mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
    mHistoricalDataTable(index).historicalDataReader.HistoricalDataAvailable mHistoricalDataTable(index).bars.size
Else
    
    failpoint = 700
    
    ' request the next lot of bars
    generateTwsHistoricalDataRequest index
End If


Exit Sub

Err:
If Err.number = ErrHistGenerationFailed Then
    mHistoricalDataTable(index).barSpecifier.MaxNumberOfBars = 0
    Debug.Print "Historical data retrieved: bars in list: " & mHistoricalDataTable(index).bars.size
    mHistoricalDataTable(index).dataRetrievedFromTWS = True
    mHistoricalDataTable(index).nextBarIndex = mHistoricalDataTable(index).bars.firstIndex
    mHistoricalDataTable(index).historicalDataReader.HistoricalDataAvailable mHistoricalDataTable(index).bars.size
    Exit Sub
End If

If Err.number = DataIncomplete Then Err.Raise DataIncomplete

HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Function getHistDataTableIndexFromRequestId(ByVal id As Long) As Long
Const ProcName As String = "getHistDataTableIndexFromRequestId"
On Error GoTo Err

getHistDataTableIndexFromRequestId = (id - IdBaseReqHistoricalData) Mod HistoricalDataReqIdIncrement

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

Private Function getIdType( _
                ByVal id As Long) As IdTypes
Const ProcName As String = "getIdType"
On Error GoTo Err

If id >= IdBaseOrderPlex Then
    getIdType = IdTypeOrder
ElseIf id >= IdBaseContractDataReq Then
    getIdType = IdTypeContractData
ElseIf id >= IdBaseExecutionsReq Then
    getIdType = IdTypeExecution
ElseIf id >= IdBaseReqHistoricalData Then
    getIdType = IdTypeHistoricalData
ElseIf id >= IdBaseMarketDepthReq Then
    getIdType = IdTypeMarketDepth
ElseIf id >= 0 Then
    getIdType = IdTypeRealtimeData
Else
    getIdType = IdTypeNone
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

Private Sub getManagedAccounts(ByVal version As Long)
Const ProcName As String = "getManagedAccounts"
Dim accountsList As String: accountsList = getString("Accounts list")

Dim failpoint As Long
On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getMarketDepth(ByVal version As Long)
Const ProcName As String = "getMarketDepth"
Dim id As Long: id = getLong("Id")
Dim tick As GenericTick

Dim failpoint As Long
On Error GoTo Err

tick.tickType = TickTypeMarketDepth
tick.position = getLong("Position")
tick.operation = getLong("Operation")
tick.side = getLong("Side")
tick.price = getDouble("Price")
tick.size = getLong("Size")

logSocketInputMessage

With mTickerTable(id - IdBaseMarketDepthReq)
    
    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    tick.timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.Timezone)
    
    .dataConsumer.tick tick
    
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getMarketDepthL2(ByVal version As Long)
Const ProcName As String = "getMarketDepthL2"
Dim id As Long: id = getLong("Id")
Dim tick As GenericTick

Dim failpoint As Long
On Error GoTo Err

tick.tickType = TickTypeMarketDepth
tick.position = getLong("Position")
tick.marketmaker = getString("Marketmaker")
tick.operation = getLong("Operation")
tick.side = getLong("Side")
tick.price = getDouble("Price")
tick.size = getLong("Size")

logSocketInputMessage

With mTickerTable(id - IdBaseMarketDepthReq)
    
    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    tick.timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.Timezone)
    
    .dataConsumer.tick tick
    
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getNewsBulletins(ByVal version As Long)
Const ProcName As String = "getNewsBulletins"
Dim msgId As Long: msgId = getLong("Id")
Dim msgType As Long: msgType = getLong("Type")
Dim newsMessage As String: newsMessage = getString("Message")
Dim originatingExch As String: originatingExch = getString("Orig exchange")

Dim failpoint As Long
On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getNextValidId(ByVal version As Long)
Const ProcName As String = "getNextValidId"


Dim failpoint As Long
On Error GoTo Err

mNextOrderPlexID = OrderPlexIdIncrement * Int((getLong("Next id") + OrderPlexIdIncrement - 1) / OrderPlexIdIncrement)

logSocketInputMessage

If mNextOrderPlexID < IdBaseOrderPlex Then mNextOrderPlexID = IdBaseOrderPlex

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getOpenOrder(ByVal version As Long)
Const ProcName As String = "getOpenOrder"
Dim mapEntry As OrderIdMapEntry

Dim id As Long
Dim clientID As Long

Dim symbol As String
Dim localSymbol As String
Dim sectype As String ' keep the socket string for logging
Dim expiry As String
Dim strike As Double
Dim rightString As String
Dim exchange As String
Dim currencyCode As String

Dim dateString As String

Dim lContractSpecifier As ContractSpecifier

Dim lOrderReport As IOrderReport

Dim failpoint As Long
On Error GoTo Err

Set lOrderReport = New OrderReport

id = getLong("Id")

' read contract fields
If version >= 17 Then setproperty lOrderReport.ProviderProperties, "conId"
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
With lOrderReport
    .BrokerId = id
    .ContractSpecifier = lContractSpecifier
    .action = orderActionFromString(getString("Action"))
    .Quantity = getLong("Quantity")
    .OrderType = orderTypeFromString(getString("Order type"))
    .LimitPrice = getDouble("Limit price")
    .TriggerPrice = getDouble("Aux price")
    .TimeInForce = orderTIFFromString(getString("Time in force"))
    setproperty .ProviderProperties, "OCA group"
    setproperty .ProviderProperties, "Account"   ' .account
    setproperty .ProviderProperties, "Open/close"   ' .openClose
    .Origin = getLong("Origin")
    .OriginatorRef = getString("Order ref")
    If version >= 3 Then clientID = getLong("Client id")
    
    If version >= 4 Then
        setproperty .ProviderProperties, "Perm id"    ' .permId =
        .IgnoreRegularTradingHours = getBoolean("Outside RTH")
        .Hidden = getBoolean("Hidden")
        .DiscretionaryAmount = getDouble("Discr amt")
    End If

    If version >= 5 Then
        dateString = getString("Good after time")
        If dateString <> "" Then .GoodAfterTime = CDate(Left$(dateString, 4) & "/" & _
                                                        Mid$(dateString, 5, 2) & "/" & _
                                                        Mid$(dateString, 7, 2) & " " & _
                                                        Right$(dateString, 8))
    End If
    
    If version >= 6 Then setproperty .ProviderProperties, "Shares allocation"  ' .sharesAllocation
    
    If version >= 7 Then
        setproperty .ProviderProperties, "FA Group"   ' .FAGroup
        setproperty .ProviderProperties, "FA method"  ' .FAMethod
        setproperty .ProviderProperties, "FA Percentage"  ' .FAPercentage
        setproperty .ProviderProperties, "FA Profile"  ' .FAProfile
    End If
    
    If version >= 8 Then
        dateString = getString("Good till date")
        If dateString <> "" Then .GoodTillDate = CDate(Left$(dateString, 4) & "/" & _
                                                        Mid$(dateString, 5, 2) & "/" & _
                                                        Mid$(dateString, 7, 2) & " " & _
                                                        Right$(dateString, 8))
    End If
    
    If version >= 9 Then
        setproperty .ProviderProperties, "Rule80A"          ' rule80A
        setproperty .ProviderProperties, "Percent Offset"
        .SettlingFirm = getString("Settling Firm")
        setproperty .ProviderProperties, "shortSaleSlot"
        setproperty .ProviderProperties, "designatedLocation"
        setproperty .ProviderProperties, "Auction Strategy"
        setproperty .ProviderProperties, "starting Price"
        setproperty .ProviderProperties, "stockRefPrice"
        setproperty .ProviderProperties, "delta"
        setproperty .ProviderProperties, "stockRangeLower"
        setproperty .ProviderProperties, "stockRangeUpper"
        .DisplaySize = getLong("displaySize")
        'setproperty .providerProperties, "rthOnly"
        .BlockOrder = getBoolean("blockOrder")
        .SweepToFill = getBoolean("sweepToFill")
        .AllOrNone = getBoolean("allOrNone")
        .MinimumQuantity = getLong("minQty")
        setproperty .ProviderProperties, "ocaType"
        .ETradeOnly = getBoolean("eTradeOnly")
        .FirmQuoteOnly = getBoolean("firmQuoteOnly")
        .NbboPriceCap = getDouble("nbboPriceCap")
    End If

    If version >= 10 Then
        setproperty .ProviderProperties, "parentId"
        .StopTriggerMethod = orderStopTriggerMethodFromString(getString("triggerMethod"))
    End If

    If version >= 11 Then
        setproperty .ProviderProperties, "volatility"
        setproperty .ProviderProperties, "volatilityType"
        If version = 11 Then
            setproperty .ProviderProperties, "receivedInt"
        Else ' version 12 and up
            setproperty .ProviderProperties, "deltaNeutralOrderType"
            setproperty .ProviderProperties, "deltaNeutralAuxPrice"
        End If
        setproperty .ProviderProperties, "continuousUpdate"
        If mServerVersion = 26 Then
            setproperty .ProviderProperties, "stockRangeLower"
            setproperty .ProviderProperties, "stockRangeUpper"
        End If
        setproperty .ProviderProperties, "referencePriceType"
    End If
    
    If version >= 13 Then
        setproperty .ProviderProperties, "trailStopPrice"
    End If

    If version >= 14 Then
        setproperty .ProviderProperties, "basisPoints"
        setproperty .ProviderProperties, "basisPointsType"
        setproperty .ProviderProperties, "comboLegsDescrip"
    End If

    If (version >= 15) Then
        If (version >= 20) Then
            setproperty .ProviderProperties, "Scale Init Level Size"
            setproperty .ProviderProperties, "Scale Subs Level Size"
        Else
            setproperty .ProviderProperties, "Scale Init Level Size"
        End If
        setproperty .ProviderProperties, "Scale Price Increment"
    End If
    
    If (version >= 19) Then
        setproperty .ProviderProperties, "Clearing Account"
        setproperty .ProviderProperties, "Clearing Intent"
    End If
    
    If version >= 22 Then
        setproperty .ProviderProperties, "Not held"
    End If
    
    If (version >= 20) Then
        setproperty .ProviderProperties, "UnderComp"
        If CBool(.ProviderProperties.GetParameterValue("UnderComp")) Then
            'UnderComp underComp = new UnderComp()
            setproperty .ProviderProperties, "UnderComp ConId"
            setproperty .ProviderProperties, "UnderComp Delta"
            setproperty .ProviderProperties, "UnderComp Price"
            'Contract.underComp = underComp
        End If
    End If
    
    If version >= 21 Then
        Dim algoStrategy As String
        algoStrategy = getString("Algo strategy")
        If algoStrategy <> "" Then
            Dim algoParamsCount As Long
            algoParamsCount = getLong("Algo params count")
            If algoParamsCount > 0 Then
                Dim i As Long
                For i = 1 To algoParamsCount
                    getString "Tag"
                    getString "Value"
                Next
            End If
        End If
    End If
    
    If (version >= 16) Then

        setproperty .ProviderProperties, "What If"

        .status = orderStatusFromString(getString("Status"))
        setproperty .ProviderProperties, "Init Margin"
        setproperty .ProviderProperties, "Maint Margin"
        setproperty .ProviderProperties, "Equity With Loan"
        setproperty .ProviderProperties, "Commission"
        setproperty .ProviderProperties, "Min Commission"
        setproperty .ProviderProperties, "Max Commission"
        setproperty .ProviderProperties, "Commission Currency"
        setproperty .ProviderProperties, "Warning Text"
    End If
End With

logSocketInputMessage True

On Error Resume Next
mapEntry = mOrderIDMap(generateTwsOrderKey(id, clientID))
On Error GoTo Err

If Not mapEntry.order Is Nothing Then
    lOrderReport.tradebuildId = mapEntry.order.tradebuildId
    mapEntry.OrderSubmitter.OrderReport lOrderReport
Else
    ' this must be an open order from the previous API session, or this is
    ' a duplicate openOrder message for an order that has already finished
    
    If Not mOrderRecoveryServiceConsumer Is Nothing Then
        mOrderRecoveryServiceConsumer.OrderReport lOrderReport
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getOpenOrderEnd(ByVal version As Long)
Const ProcName As String = "getOpenOrderEnd"
Dim failpoint As String
On Error GoTo Err

logSocketInputMessage

If Not mOrderRecoveryServiceConsumer Is Nothing Then mOrderRecoveryServiceConsumer.OrderReportsComplete
Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getOrderStatus(ByVal version As Long)
Const ProcName As String = "getOrderStatus"
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
Dim statusRpt As New OrderStatusReport

Dim failpoint As Long
On Error GoTo Err

If version >= 2 Then permId = getLong("Perm id")

If version >= 3 Then parentId = getLong("Parent id")

If version >= 4 Then lastFillPrice = getDouble("Last fill price")

If version >= 5 Then clientID = getLong("Client id")

If version >= 6 Then whyHeld = getString("Why held")

logSocketInputMessage True

On Error Resume Next
mapEntry = mOrderIDMap(generateTwsOrderKey(id, clientID))
On Error GoTo Err

If Not mapEntry.order Is Nothing Then
    avgFillPrice = adjustPrice(avgFillPrice, mapEntry.order.Contract.specifier)
    lastFillPrice = adjustPrice(lastFillPrice, mapEntry.order.Contract.specifier)
    
    statusRpt.Initialise mapEntry.order.tradebuildId, status
    mapEntry.OrderSubmitter.OrderStatusReport statusRpt
    If status = OrderStatusCancelling And parentId <> 0 Then
        ' for some reason IB never sends a cancelled status when a target order
        ' is cancelled after the entry order has been filled. It just sits there at
        ' 'cancelling' status until TWS is closed (note that it doesn't reappear if
        ' TWS is restarted, so it must have been cancelled).
        ' So we'll treat it as cancelled if we haven't received a fill in say 500 millisecs
        GenerateDeferredOrderCancelNotification id, clientID
    End If
    
    If status = OrderStatusCancelled Or status = OrderStatusFilled Then
        'mOrderIDMap.remove generateTwsOrderKey(id, clientID)
    End If
Else
    ' this must be an open order from the previous API session, but
    ' for which we haven't yet had the openOrder notification
    ' don't notify the service consumer because it won't know about this
    ' order yet
    '
    ' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! need to improve this!!!!!!!!!!!!!!!
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getPortfolioValue(ByVal version As Long)
Const ProcName As String = "getPortfolioValue"


Dim symbol As String
Dim localSymbol As String
Dim sectype As String ' keep the socket string for logging
Dim expiry As String
Dim strike As Double
Dim rightString As String
Dim exchange As String
Dim currencyCode As String

Dim lContractSpecifier As ContractSpecifier

Dim failpoint As Long
On Error GoTo Err

If version >= 6 Then getLong "Contract id"

symbol = getString("Symbol")
sectype = getString("Sec type")
expiry = getString("Expiry")
strike = getDouble("Strike")
rightString = getString("Right")

If version >= 7 Then
    getString "Multiplier"
    getString "Primary exchg"
End If

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

Dim accountName As String: If version >= 4 Then accountName = getString("Account Name")

If version = 6 And mServerVersion = 39 Then getString "Primary exchange"

logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getScannerData(ByVal version As Long)
Const ProcName As String = "getScannerData"
Dim lContractSpecifier As ContractSpecifier
Dim lContractBuilder As ContractBuilder
Dim tickerId As Long
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

Dim failpoint As Long
On Error GoTo Err

tickerId = getLong("Id")
numberOfElements = getLong("Number of elements")

For i = 0 To numberOfElements - 1
    rank = getLong("Rank")

    If (version >= 3) Then getLong "ConId"
    
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
    If (version >= 2) Then getString "Legs Str "
Next

logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getScannerParameters(ByVal version As Long)
Const ProcName As String = "getScannerParameters"
Dim xml  As String

Dim failpoint As Long
On Error GoTo Err

xml = getString("XML")

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getTickGeneric(ByVal version As Long)
Const ProcName As String = "getTickGeneric"
Dim tickerId As Long
Dim tickType As TWSSocketTickTypes
Dim value As Double

Dim failpoint As Long
On Error GoTo Err

tickerId = getLong("tickerId")
tickType = getLong("tickType")
value = getDouble("value")

logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getTickOptionComputation(ByVal version As Long)
Dim tickerId As Long: tickerId = getLong("tickerId")
Dim tickType As TWSSocketTickTypes: tickType = getLong("tickType")
Dim impliedVol As Double: impliedVol = getDouble("impliedVol")

Const ProcName As String = "getTickOptionComputation"
Dim failpoint As String
On Error GoTo Err

If impliedVol < 0 Then
    ' -1 is the "not yet computed" indicator
    impliedVol = MaxDouble
End If

Dim delta As Double: delta = getDouble("delta")
If Abs(delta) > 1 Then
    ' -2 is the "not yet computed" indicator
    delta = MaxDouble
End If

Dim modelPrice As Double
Dim pvDividend As Double
If tickType = TWSSocketTickTypes.TICK_MODEL_OPTION Then
    modelPrice = getDouble("modelPrice")
    pvDividend = getDouble("pvDividend")
Else
    modelPrice = MaxDouble
    pvDividend = MaxDouble
End If

logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getTickPrice(ByVal version As Long)
Const ProcName As String = "getTickPrice"
Dim tick As GenericTick
Dim tickType As Long
Dim tickerId As Long
Dim canAutoExecute As Boolean

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

tickerId = getLong("Ticker id")

failpoint = 200

tickType = getLong("Tick type")

failpoint = 300

tick.price = getDouble("Price")

failpoint = 400

If version = 1 Then Debug.Print "****************** Type 1 tick price message received *********************"
If version >= 2 Then tick.size = getLong("Size")
If version >= 3 Then canAutoExecute = getBoolean("Can autoexecute")

logSocketInputMessage

failpoint = 500

With mTickerTable(tickerId Mod IncarnationIncrement)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    tick.timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.Timezone)
    
    Select Case tickType
    Case TICK_BID
        If Not validPrice(tick.price, .prevBid, "bid", .Contract) Then Exit Sub
        tick.tickType = TickTypeBid
        If tick.size = 0 Then tick.size = .prevBidSize
        .prevBid = tick.price
        If tick.size <> 0 Then .prevBidSize = tick.size
    Case TICK_ASK
        If Not validPrice(tick.price, .prevAsk, "ask", .Contract) Then Exit Sub
        tick.tickType = TickTypeAsk
        If tick.size = 0 Then tick.size = .prevAskSize
        .prevAsk = tick.price
        If tick.size <> 0 Then .prevAskSize = tick.size
    Case TICK_LAST
        If Not validPrice(tick.price, .prevTrade, "trade", .Contract) Then Exit Sub
        tick.tickType = TickTypeTrade
        If tick.size = 0 Then tick.size = .prevTradeSize
        .prevTrade = tick.price
        If tick.size <> 0 Then .prevTradeSize = tick.size
    Case TICK_HIGH
        If Not validPrice(tick.price, .prevHigh, "high", .Contract) Then Exit Sub
        tick.tickType = TickTypeHighPrice
        .prevHigh = tick.price
    Case TICK_LOW
        If Not validPrice(tick.price, .prevLow, "low", .Contract) Then Exit Sub
        tick.tickType = TickTypeLowPrice
        .prevLow = tick.price
    Case TICK_OPEN
        If Not validPrice(tick.price, .prevOpen, "open", .Contract) Then Exit Sub
        tick.tickType = TickTypeOpenPrice
        .prevOpen = tick.price
    Case TICK_CLOSE
        If Not validPrice(tick.price, .prevClose, "close", .Contract) Then Exit Sub
        tick.tickType = TickTypeClosePrice
        .prevClose = tick.price
    End Select
        
    .dataConsumer.tick tick
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getTickSize(ByVal version As Long)
Const ProcName As String = "getTickSize"
Dim tick As GenericTick
Dim tickType As Long
Dim tickerId As Long

Dim failpoint As Long
On Error GoTo Err

failpoint = 100

tickerId = getLong("Ticker id")

failpoint = 200

tickType = getLong("Tick type")

failpoint = 300

tick.size = getLong("Size")

logSocketInputMessage

failpoint = 400

With mTickerTable(tickerId Mod IncarnationIncrement)
    
    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    tick.timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.Timezone)
    
    Select Case tickType
    Case TICK_BID_SIZE
        tick.tickType = TickTypeBid
        If tick.size = .prevBidSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        tick.price = .prevBid
        .prevBidSize = tick.size
    
    Case TICK_ASK_SIZE
        tick.tickType = TickTypeAsk
        If tick.size = .prevAskSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        tick.price = .prevAsk
        .prevAskSize = tick.size
    
    Case TICK_LAST_SIZE
        tick.tickType = TickTypeTrade
        If tick.size = .prevTradeSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        tick.price = .prevTrade
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub getTickSnapshotEnd(ByVal version As Long)
Const ProcName As String = "getTickSnapshotEnd"
Dim failpoint As String
On Error GoTo Err

Dim reqId As Long: reqId = getLong("reqID")
logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub getTickString(ByVal version As Long)
Const ProcName As String = "getTickString"
Dim failpoint As Long
On Error GoTo Err

getLong "tickerId"
getLong "tickType"
getString "value"

logSocketInputMessage

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
handleFatalErrorEx Err.number, Err.source, Err.Description, pProcName, pFailpoint
End Sub

Private Sub handleFatalErrorEx( _
                ByVal errNum As Long, _
                ByRef errSource As String, _
                ByRef errDesc As String, _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
On Error GoTo Err

errSource = errSource & vbCrLf & _
            ProjectName & "." & _
            ModuleName & ":" & _
            pProcName & _
            IIf(pFailpoint <> "", " At " & pFailpoint, "")

Err.Raise errNum, errSource, errDesc

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.number, Err.source, Err.Description, -1
End Sub

Private Sub HandleUnexpectedError( _
                ByRef pProcedureName As String, _
                Optional ByVal pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False)
TWUtilities30.HandleUnexpectedError pProcedureName, ProjectName, ModuleName, pFailpoint, pReRaise, pLog
End Sub

Private Sub handleTwsDisconnection()
Const ProcName As String = "handleTwsDisconnection"
Dim failpoint As Long
On Error GoTo Err

mConnectionState = ConnectionStates.ConnNotConnected
Set mSocket = Nothing

If mDisconnectedByApplication Then
    logMessage "Connection to TWS closed: " & ConnectionString, pProcName:="handleTwsDisconnection"
Else
    If mConnectionRetryIntervalSecs = 0 Then
        cancelContractRequests "Lost connection to TWS"
        RaiseEvent ConnectionToTWSClosed(False)
        logMessage "Connection to TWS closed" & _
                    ConnectionString, _
                    "handleTwsDisconnection"
    Else
        retryConnection
        RaiseEvent ConnectionToTWSClosed(True)
        logMessage "Connection to TWS closed - reconnecting: " & _
                    ConnectionString, _
                    "handleTwsDisconnection"
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
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
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub logSocketInputMessage(Optional ByVal ignoreLogLevel As Boolean)
Const ProcName As String = "logSocketInputMessage"
Static prevMessageBuilder As StringBuilder

Dim failpoint As Long
On Error GoTo Err

If ignoreLogLevel Then
    If Not mInMessageBuilder Is Nothing Then
        mSocketLogger.Log LogLevelNormal, mInMessageBuilder.ToString
        Set prevMessageBuilder = mInMessageBuilder
    ElseIf Not prevMessageBuilder Is Nothing Then
        mSocketLogger.Log LogLevelNormal, prevMessageBuilder.ToString
        Set prevMessageBuilder = Nothing
    End If
ElseIf mSocketLogger.IsLoggable(LogLevelHighDetail) Then
    mSocketLogger.Log LogLevelHighDetail, mInMessageBuilder.ToString
    Set prevMessageBuilder = mInMessageBuilder
Else
    Set prevMessageBuilder = mInMessageBuilder
End If
Set mInMessageBuilder = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub logSocketOutputMessage(Optional ByVal ignoreLogLevel As Boolean)
Const ProcName As String = "logSocketOutputMessage"
Dim failpoint As Long
On Error GoTo Err

If ignoreLogLevel Then
    mSocketLogger.Log LogLevelNormal, mOutMessageBuilder.ToString
ElseIf mSocketLogger.IsLoggable(LogLevelHighDetail) Then
    mSocketLogger.Log LogLevelHighDetail, mOutMessageBuilder.ToString
End If
Set mOutMessageBuilder = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
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
Case "REJECTED", "INACTIVE"
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

Private Function orderStopTriggerMethodFromString(ByVal value As String) As StopTriggerMethods
Select Case value
Case "0"
    orderStopTriggerMethodFromString = StopTriggerMethods.StopTriggerDefault
Case "1"
    orderStopTriggerMethodFromString = StopTriggerMethods.StopTriggerDoubleBidAsk
Case "2"
    orderStopTriggerMethodFromString = StopTriggerMethods.StopTriggerLast
Case "3"
    orderStopTriggerMethodFromString = StopTriggerMethods.StopTriggerDoubleLast
Case "4"
    orderStopTriggerMethodFromString = StopTriggerMethods.StopTriggerBidAsk
Case "7"
    orderStopTriggerMethodFromString = StopTriggerMethods.StopTriggerLastOrBidAsk
Case "8"
    orderStopTriggerMethodFromString = StopTriggerMethods.StopTriggerMidPoint
End Select
End Function

Private Function orderStopTriggerMethodToString(ByVal value As StopTriggerMethods) As String
Select Case value
Case StopTriggerMethods.StopTriggerBidAsk
    orderStopTriggerMethodToString = "1"
Case StopTriggerMethods.StopTriggerDefault
    orderStopTriggerMethodToString = "0"
Case StopTriggerMethods.StopTriggerDoubleBidAsk
    orderStopTriggerMethodToString = "1"
Case StopTriggerMethods.StopTriggerDoubleLast
    orderStopTriggerMethodToString = "3"
Case StopTriggerMethods.StopTriggerLast
    orderStopTriggerMethodToString = "2"
Case StopTriggerMethods.StopTriggerLastOrBidAsk
    orderStopTriggerMethodToString = "7"
Case StopTriggerMethods.StopTriggerMidPoint
    orderStopTriggerMethodToString = "8"
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
Const ProcName As String = "orderTypeToString"
On Error GoTo Err

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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

Private Function permittedHistDays( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
Const ProcName As String = "permittedHistDays"
On Error GoTo Err

permittedHistDays = DateDiff("d", DateAdd("yyyy", -1, Now), reqEndDate)
If permittedHistDays = 0 Then Err.Raise ErrHistGenerationFailed
If permittedHistDays > desiredBars Then permittedHistDays = desiredBars

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

Private Function permittedHistMonths( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
Const ProcName As String = "permittedHistMonths"
On Error GoTo Err

permittedHistMonths = DateDiff("m", DateAdd("m", 1, DateAdd("yyyy", -1, Now)), reqEndDate)
If permittedHistMonths = 0 Then Err.Raise ErrHistGenerationFailed
If permittedHistMonths > desiredBars Then permittedHistMonths = desiredBars

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

Private Function permittedHistSeconds( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
' the 86398 should really be 86400 but allows for clocks not being
' properly synchronised
Const ProcName As String = "permittedHistSeconds"
On Error GoTo Err

permittedHistSeconds = DateDiff("s", DateAdd("yyyy", -1, Now), reqEndDate) + 86398
If permittedHistSeconds = 0 Then Err.Raise ErrHistGenerationFailed
If permittedHistSeconds > desiredBars Then permittedHistSeconds = desiredBars

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

Private Function permittedHistWeeks( _
                ByVal reqEndDate As Date, _
                ByVal desiredBars As Long) As Long
Const ProcName As String = "permittedHistWeeks"
On Error GoTo Err

permittedHistWeeks = DateDiff("w", DateAdd("w", 1, DateAdd("yyyy", -1, Now)), reqEndDate)
If permittedHistWeeks = 0 Then Err.Raise ErrHistGenerationFailed
If permittedHistWeeks > desiredBars Then permittedHistWeeks = desiredBars

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

Private Sub processContractError(ByVal reason As String)
Const ProcName As String = "processContractError"
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
Set mContractDetailsRequest.requester = Nothing
mFirstContractDetailsReceived = False

failpoint = 300
If mConnectionState = ConnConnected Then submitNextContractDetailsRequest

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub processNextHistDataRequest()
Const ProcName As String = "processNextHistDataRequest"
Dim failpoint As Long
On Error GoTo Err

If mConcurrentHistDataRequests > 0 Then mConcurrentHistDataRequests = mConcurrentHistDataRequests - 1
If mTwsHistDataRequestQueue.size <> 0 Then
    Dim request As TWSHistoricalDataRequestQueueEntry
    request = mTwsHistDataRequestQueue.removeFromFront
    reqHistoricalData request
    mConcurrentHistDataRequests = mConcurrentHistDataRequests + 1
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Function queueTWSHistoricalDataRequest(ByRef TwsRequest As TWSHistoricalDataRequestQueueEntry) As String
Const ProcName As String = "queueTWSHistoricalDataRequest"
Dim failpoint As Long
On Error GoTo Err

queueTWSHistoricalDataRequest = mTwsHistDataRequestQueue.addToEnd(TwsRequest)

If mConcurrentHistDataRequests < MaxConcurrentHistDataRequests Then
    mTwsHistDataRequestQueue.removeFromFront
    reqHistoricalData TwsRequest
    mConcurrentHistDataRequests = mConcurrentHistDataRequests + 1
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

Private Sub reEstablishMarketData(ByVal newConnection As Boolean)
Const ProcName As String = "reEstablishMarketData"
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

logMessage "Re-establishing market data: newConnection = " & CStr(newConnection), "reEstablishMarketData"
For i = 0 To UBound(mTickerTable)
    With mTickerTable(i)
        If Not .Contract Is Nothing Then
            If Not newConnection Then cancelMarketDataEx i, False
            .incarnation = .incarnation + IncarnationIncrement
            reqMarketData i
        End If
    End With
    clearPrevPrices i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub releaseAllTickerTableEntries()
Const ProcName As String = "releaseAllTickerTableEntries"
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

For i = 0 To UBound(mTickerTable)
    releaseTickerTableEntry i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub releaseHistoricalDataTableEntry( _
                ByVal index As Long)
Const ProcName As String = "releaseHistoricalDataTableEntry"
Dim failpoint As Long
On Error GoTo Err

With mHistoricalDataTable(index)
    Set .bars = Nothing
    .requestKey = ""
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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub ReleaseSocket()
Const ProcName As String = "ReleaseSocket"
Dim failpoint As Long
On Error GoTo Err

If Not mSocket Is Nothing Then
    mSocket.Close
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub releaseTickerTableEntry(ByVal index As Long)
Const ProcName As String = "releaseTickerTableEntry"
Dim failpoint As Long
On Error GoTo Err

With mTickerTable(index)
    .inUse = False
    .incarnation = .incarnation + IncarnationIncrement
    Set .Contract = Nothing
    .receivingMarketDepth = False
    Set .dataConsumer = Nothing
    Set .dataReader = Nothing
End With
clearPrevPrices index

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub reqContract()
Const ProcName As String = "reqContract"
Static reqId As Long
Dim ContractSpecifier As ContractSpecifier
Dim failpoint As Long
On Error GoTo Err

Set ContractSpecifier = mContractDetailsRequest.contractSpec

If ContractSpecifier.exchange <> "" And Not IsValidExchangeCode(ContractSpecifier.exchange) Then
    processContractError "invalid exchange"
    Exit Sub
End If

Const version = 5

addMsgId REQ_CONTRACT_DATA
addData version, "Version"

If mServerVersion >= MIN_SERVER_VER_CONTRACT_DATA_CHAIN Then
    reqId = reqId + 1
    addData IdBaseContractDataReq + reqId, "Request id"
End If

With ContractSpecifier
    If mServerVersion >= MIN_SERVER_VER_CONTRACT_CONID Then addData 0, "Contract id"
    addData .symbol, "Symbol"
    addData secTypeToString(.sectype), "Sec type"
    'addData Left$(.expiry, 6), "Expiry"
    addData .expiry, "Expiry"
    addDouble .strike, "Strike"
    addData optionRightToString(.Right), "Right"
    If mServerVersion >= 15 Then addData "", "Multiplier"
    addData .exchange, "Exchange"
    addData .currencyCode, "Currency"
    If (.sectype = SecTypeFuture Or _
        .sectype = SecTypeFuturesOption Or _
        .sectype = SecTypeOption) _
    Then
        If expiryToDate(.expiry) >= Now Or .expiry = "" Then
            addData .localSymbol, "Local symbol"
        Else
            addData "", "Local Symbol"
        End If
    Else
        addData .localSymbol, "Local Symbol"
    End If
    If mServerVersion >= 31 Then
        If .sectype = SecTypeCash Or .sectype = SecTypeIndex Or .sectype = SecTypeStock Then
            addData 0, "Include expired" ' can't include expired for non-expiring contracts
        Else
            addData 1, "Include expired"
        End If
    End If
End With
send

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub reqHistoricalData( _
                request As TWSHistoricalDataRequestQueueEntry)
Const ProcName As String = "reqHistoricalData"
Dim failpoint As Long
On Error GoTo Err

logMessage "Requesting historical data for: " & request.Contract.specifier.localSymbol & _
            "; id=" & request.id & _
            "; barsize=" & request.barSizeSetting & _
            "; endTime=" & request.endDateTime & _
            "; duration=" & request.duration, _
            "reqHistoricalData"
Const version = 4

If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

addMsgId REQ_HISTORICAL_DATA
addData version, "Version"
addData request.id, "Request id"
addData request.Contract.specifier.symbol, "Symbol"
addData secTypeToString(request.Contract.specifier.sectype), "Sec type"
'addData Left$(request.Contract.specifier.expiry, 6), "Expiry"
addData request.Contract.specifier.expiry, "Expiry"
addDouble request.Contract.specifier.strike, "Strike"
addData optionRightToString(request.Contract.specifier.Right), "Right"
addData "", "Multiplier"     ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>> multiplier ??? why?
addData request.Contract.specifier.exchange, "Exchange"
addData "", "Primary exchange"
addData request.Contract.specifier.currencyCode, "Currency"
If expiryToDate(request.Contract.specifier.expiry) < Now Then
    addData "", "Local symbol"
Else
    addData request.Contract.specifier.localSymbol, "Local symbol"
End If
If mServerVersion >= 31 Then
    If request.Contract.specifier.sectype = SecTypeCash Or _
        request.Contract.specifier.sectype = SecTypeIndex Or _
        request.Contract.specifier.sectype = SecTypeStock _
    Then
        addData 0, "Include expired" ' can't include expired for non-expiring contracts
    Else
        addData 1, "Include expired"
    End If
End If
If mServerVersion >= 20 Then
    addData request.endDateTime, "End date"
    addData request.barSizeSetting, "Bar size"
End If

' durationStr

addData request.duration, "Duration"

' useRTH
addData 0, "Use RTH"

addData request.whatToShow, "What to show"

' formatDate
If mServerVersion > 16 Then
    addData TWSHistDataDateFormats.DateFormatString, "Date format"
End If

If request.Contract.specifier.ComboLegs Is Nothing Then
    addData 0, "Combo legs count"
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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub reqMarketData(ByVal tickerTableIndex As Long)
Const ProcName As String = "reqMarketData"


Const version = 8

Dim failpoint As Long
On Error GoTo Err

logMessage "Requesting market data for: " & mTickerTable(tickerTableIndex).Contract.specifier.ToString, "reqMarketData"

addMsgId REQ_MKT_DATA
addData version, "Version"
addData tickerTableIndex + mTickerTable(tickerTableIndex).incarnation, "Ticker id"
With mTickerTable(tickerTableIndex).Contract.specifier
    addData .symbol, "Symbol"
    addData secTypeToString(.sectype), "Sec type"
    'addData Left$(.expiry, 6), "Expiry"
    addData .expiry, "Expiry"
    addDouble .strike, "Strike"
    addData optionRightToString(.Right), "Right"
    If mServerVersion >= 15 Then addData "", "Multiplier"
    addData .exchange, "Exchange"
    If mServerVersion >= 15 Then addData "", "Primary exchange"
    addData .currencyCode, "Currency"
    If mServerVersion >= 2 Then addData .localSymbol, "Local symbol"
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
    
    If (mServerVersion >= MIN_SERVER_VER_UNDER_COMP) Then
        ' not currently aupported
'       if (contract.m_underComp != null) then
'           UnderComp underComp = contract.m_underComp;
'           send( true);
'           send( underComp.m_conId);
'           send( underComp.m_delta);
'           send( underComp.m_price);
'       else then
           addBoolean False, "Under comp"
'       End If
    End If
    
    If mServerVersion >= 31 Then addData "", "Generic tick list"
            
    If mServerVersion >= MIN_SERVER_VER_SNAPSHOT_MKT_DATA Then addBoolean False, "Snapshot"
    
    send
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName

End Sub

Private Sub reqMarketDepth(ByVal tickerId As Long, _
                        ByVal contractSpec As ContractSpecifier)
Const ProcName As String = "reqMarketDepth"


Const version = 3

Dim failpoint As Long
On Error GoTo Err

addMsgId REQ_MKT_DEPTH
addData version, "Version"
addData tickerId + IdBaseMarketDepthReq, "Request id"
With contractSpec
    addData .symbol, "Symbol"
    addData secTypeToString(.sectype), "Sec type"
    'addData Left$(.expiry, 6), "Expiry"
    addData .expiry, "Expiry"
    addDouble .strike, "Strike"
    addData optionRightToString(.Right), "Right"
    If mServerVersion >= 15 Then addData "", "Multiplier"
    addData .exchange, "Exchange"
    addData .currencyCode, "Currency"
    addData .localSymbol, "Local symbol"
    If mServerVersion >= 19 Then addData 20, "Num rows"
End With
send

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub resetMarketDepth(ByVal reEstablish As Boolean)
Const ProcName As String = "resetMarketDepth"
Dim i As Long
Dim tick As GenericTick

Dim failpoint As Long
On Error GoTo Err

tick.tickType = TickTypeMarketDepthReset

For i = 0 To UBound(mTickerTable)
    With mTickerTable(i)
        If .receivingMarketDepth Then
            tick.timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .Contract.Timezone)
            If reEstablish Then reqMarketDepth i, .Contract.specifier
            .dataConsumer.tick tick
        End If
    End With
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub retryConnection()
Const ProcName As String = "retryConnection"
Dim failpoint As Long
On Error GoTo Err

If mConnectionRetryIntervalSecs <> 0 Then
    Set mConnectionTimer = CreateIntervalTimer(mConnectionRetryIntervalSecs, ExpiryTimeUnitSeconds)
    mConnectionTimer.StartTimer
    mRetryingConnection = True
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub rollbackBuffer()
'Debug.Print "Rollback: mInputIndex=" & mInputParseIndexChk
mInputParseIndex = mInputParseIndexChk
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

Private Sub send(Optional ByVal ignoreLogLevel As Boolean)
Const ProcName As String = "send"
Dim failpoint As Long
On Error GoTo Err

logSocketOutputMessage ignoreLogLevel

ReDim Preserve mOutputBuffer(mOutputBufferIndex - 1) As Byte
If mSocketLogger.IsLoggable(LogLevelMediumDetail) Then
    mSocketLogger.Log LogLevels.LogLevelMediumDetail, _
                        showBuffer("Out buf: ", mOutputBuffer)
End If
mSocket.SendData mOutputBuffer
ReDim mOutputBuffer(15) As Byte
mOutputBufferIndex = 0

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub sendOrder( _
                ByVal pOrder As TradeBuildSP.IOrder, _
                ByVal transmit As Boolean, _
                ByVal parentId As String, _
                ByVal ocaGroup As String, _
                ByVal routingHint As String)
Const ProcName As String = "sendOrder"


Dim specifier As ContractSpecifier

Dim failpoint As Long
On Error GoTo Err

Set specifier = pOrder.Contract.specifier

Const version = 26
addMsgId PLACE_ORDER
addData version, "Version"
addData pOrder.BrokerId, "Order id"

' send contract fields
With specifier
    addData .symbol, "Symbol"
    addData secTypeToString(.sectype), "Sectype"
    'addData Left$(.expiry, 6), "Expiry"
    addData .expiry, "Expiry"
    addDouble .strike, "Strike"
    addData optionRightToString(.Right), "Right"
    If mServerVersion >= 15 Then addData "", "Multiplier"   ' multiplier - this is the number of units of the underlying
                                                            ' delivered per contract - see http://www.interactivebrokers.com/discus/messages/2/27413.html
    addData .exchange, "Exchange"
    If mServerVersion >= 14 Then addData routingHint, "Primary exchange"    ' aka primaryExchange in TWS - only relevant when
                                                        ' exchange is SMART and there are SMART routers in more
                                                        ' than one country (eg try IBM in TWS)
    addData .currencyCode, "Currency"
    If mServerVersion >= 2 Then addData .localSymbol, "Local symbol"
End With

With pOrder
    ' send main order fields
    addData orderActionToString(.action), "Action"
    addData .Quantity, "Quantity"
    addData orderTypeToString(.OrderType), "Order type"
    If mServerVersion <= 10 And _
        (UCase$(specifier.exchange) = "LIFFE" Or _
        UCase$(specifier.exchange) = "LSE") _
    Then
        addDouble .LimitPrice / 100, "Price"
        addDouble .TriggerPrice / 100, "Aux price"
    Else
        addDouble .LimitPrice, "Price"
        addDouble .TriggerPrice, "Aux price"
    End If
    ' send extended order fields
    addData orderTIFToString(.TimeInForce), "TIF"
    addData ocaGroup, "Oca Group"
    addData "", "Account"
    addData "", "OpenClose"
    addData .Origin, "Origin"
    addData .OriginatorRef, "Originator ref"
    addBoolean transmit, "Transmit"
    If mServerVersion >= 4 Then addData parentId, "Parent id"
    If mServerVersion >= 5 Then
        addBoolean .BlockOrder, "Block order"
        addBoolean .SweepToFill, "Sweep to fill"
        addData .DisplaySize, "Display size"
        addData orderStopTriggerMethodToString(.StopTriggerMethod), "Stop trigger method"
        If mServerVersion < 38 Then
            addBoolean False, "Ignore RTH"
        Else
            addBoolean .IgnoreRegularTradingHours, "Ignore RTH"
        End If
    End If
    If mServerVersion >= 7 Then addBoolean .Hidden, "Hidden"
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
    If mServerVersion >= 9 Then addData "", "Shares Allocation"
    
    If mServerVersion >= 10 Then addData .DiscretionaryAmount, "Discretionary amount"
    
    If mServerVersion >= 11 Then
        If .GoodAfterTime = 0 Then
            addData "", "Good after time"
        Else
            addData Format(.GoodAfterTime, "yyyymmdd hh\:nn\:ss ") & .GoodAfterTimeTZ, "Good after time"
        End If
    End If
    
    If mServerVersion >= 12 Then
        If .GoodTillDate = 0 Then
            addData "", "Good till date"
        Else
            addData Format(.GoodTillDate, "yyyymmdd hh\:nn\:ss ") & .GoodTillDateTZ, "Good till date"
        End If
    End If
    
    If mServerVersion >= 13 Then
        addData "", "FAGroup"
        addData "", "FAMethod"
        addData "", "FAPercentage"
        addData "", "FAProfile"
    End If
    
    If mServerVersion >= 18 Then 'institutional short sale slot fields.
        addData 0, " Short Sale Slot"           ' 0 only for retail, 1 or 2 only for institution.
        addData "", "Designated Location"       ' only populate when shortSaleSlot = 2.
    End If
    
    If mServerVersion >= 19 Then
        addData 0, "Oca type"
        If mServerVersion < 38 Then
            addBoolean False, "RTH only" ' (now combined with ignoreRTH)
        End If
        addBoolean False, "Rule 80A"
        addData "", "Settling firm"
        addBoolean .AllOrNone, "All or none"
        addData .MinimumQuantity, "Minimum quantity"
        addData "", "Percent Offset"
        addBoolean .ETradeOnly, "E-trade only"
        addBoolean .FirmQuoteOnly, "Firm quote only"
        addData .NbboPriceCap, "NBBO price cap"
        addData "", "Auction strategy"
        addData "", "Starting price"
        addData "", "Stock ref price"
        addData "", "Delta"
        addData "", "Stock range lower"
        addData "", "Stock range upper"
    End If
    
    If mServerVersion >= 22 Then addBoolean .OverrideConstraints, "Override constraints"

    If mServerVersion >= 26 Then ' Volatility orders
        addData "", "Volatility"
        addData "", "Volatility type"
        If mServerVersion < 28 Then
            addData "0", "Delta neutral order type"
        Else
            addData "", "Delta neutral order type)"
            addData "", "Delta neutral aux price"
        End If
        
        addData 0, "Continuous update"
        
        If mServerVersion = 26 Then
            ' Volatility orders had specific watermark price attribs in Server version 26
            addData "", "Stock range lower"
            addData "", "Stock range upper"
        End If
        
        addData "", "Reference price type"
    End If
    
    If mServerVersion >= 30 Then addData 0, "TRAIL_STOP_LIMIT stop price"

    If (mServerVersion >= MIN_SERVER_VER_SCALE_ORDERS) Then
        If (mServerVersion >= MIN_SERVER_VER_SCALE_ORDERS2) Then
            addData "", "ScaleInitLevelSize"
            addData "", "ScaleSubsLevelSize"
        
        Else
            addData "", "ScaleInitLevelSize"
            addData "", "ScaleSubsLevelSize"
            
        End If
        addData "", "ScalePriceIncrement"
    End If
    
    If (mServerVersion >= MIN_SERVER_VER_PTA_ORDERS) Then
        addData "", "Clearing account"
        addData "", "Clearing intent"
    End If

    If (mServerVersion >= MIN_SERVER_VER_NOT_HELD) Then
        addBoolean False, "Not held"
    End If

    If (mServerVersion >= MIN_SERVER_VER_UNDER_COMP) Then
'        if (contract.m_underComp != null) then
'            underComp underComp = Contract.m_underComp
'            send (True)
'            send (underComp.m_conId)
'            send (underComp.m_delta)
'            send (underComp.m_price)
'
'        Else
            addBoolean False, "under comp"
'        End If
    End If

    If (mServerVersion >= MIN_SERVER_VER_ALGO_ORDERS) Then
        addData "", "Algo strategy"
'        if( !IsEmpty(order.m_algoStrategy)) {
'            java.util.Vector algoParams = order.m_algoParams;
'            int algoParamsCount = algoParams == null ? 0 : algoParams.size();
'            send( algoParamsCount);
'            if( algoParamsCount > 0) {
'                for( int i = 0; i < algoParamsCount; ++i) {
'                    TagValue tagValue = (TagValue)algoParams.get(i);
'                    send( tagValue.m_tag);
'                    send( tagValue.m_value);
'                }
'            }
'        }
    End If

    If (mServerVersion >= MIN_SERVER_VER_WHAT_IF_ORDERS) Then
        addBoolean False, "WhatIf"
    End If
End With

send True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Sub setproperty( _
                ByVal props As Parameters, _
                ByVal propName As String)
Const ProcName As String = "setproperty"
Dim failpoint As Long
On Error GoTo Err

props.SetParameterValue propName, getString(propName)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub
                    
Private Sub setTWSLogLevel(logLevel As TWSLogLevels)
Const ProcName As String = "setTWSLogLevel"


Const version = 1

Dim failpoint As Long
On Error GoTo Err

addMsgId SET_SERVER_LOGLEVEL
addData version, "Version"
addData logLevel, "Loglevel"
send

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Function showBuffer( _
                ByRef header As String, _
                ByRef buffer() As Byte) As String
Const ProcName As String = "showBuffer"
Dim failpoint As Long
On Error GoTo Err

showBuffer = header & Replace(StrConv(buffer, vbUnicode), Chr$(0), "_")

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

Private Sub submitNextContractDetailsRequest()
Const ProcName As String = "submitNextContractDetailsRequest"
Dim failpoint As Long
On Error GoTo Err

If Not (mConnectionState = ConnReSynching Or mConnectionState = ConnConnected) Then Exit Sub
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Sub

Private Function validPrice( _
                ByVal price As Double, _
                ByVal prevPrice As Double, _
                ByVal priceType As String, _
                ByVal pContract As Contract) As Boolean
Const ProcName As String = "validPrice"
Dim failpoint As Long
On Error GoTo Err

If (prevPrice <> MaxDouble Or (price <> 0 And price <> -1)) And _
    pContract.IsValidPrice(price, prevPrice) _
Then
    validPrice = True
Else
    gLogger.Log LogLevels.LogLevelNormal, "Invalid price rejected (" & priceType & "): " & _
                    "price=" & price & _
                    "; contract={" & pContract.specifier.ToString & "}"
    Exit Function
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName
End Function

