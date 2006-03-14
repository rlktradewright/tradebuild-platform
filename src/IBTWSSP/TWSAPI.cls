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

Event Connected()
    ' Raised when a successful connection to TWS has been achieved.

Event ConnectFailed(ByVal Description As String, _
                ByVal retrying As Boolean)
    ' Raised when connection to TWS has failed and the connection
    ' retry interval is zero.

Event Connecting()
    ' Raised when an attempt to connect to TWS is initiated.

Event ConnectionToTWSClosed( _
                ByVal reconnecting As Boolean)
    ' Raised when the connection to TWS is disconnected.

Event ConnectionToIBClosed()

Event ConnectionToIBRecovered()

Event ContractSpecifierInvalid( _
                ByVal reason As String, _
                ByVal pContracts As TradeBuildSP.IContracts)

Event ContractsLoaded()

Event dataReceived()
    ' Raised when a new block of data is received from the socket connecting to TWS.
    
Event ExecutionDetails( _
                ByVal id As Long, _
                ByVal pContractSpecifier As TradeBuildSP.IContractSpecifier, _
                ByVal exec As Execution)

Event FAData( _
                ByVal DataType As FADataTypes, _
                ByVal XMLData As String)
                
Event managedAccounts( _
                ByVal timestamp As Date, _
                ByVal accountsList As String)

Event openOrder( _
                ByVal pContractSpecifier As TradeBuildSP.IContractSpecifier, _
                ByVal pOrder As Order)

Event orderStatus( _
                ByVal id As Long, _
                ByVal status As OrderStatuses, _
                ByVal filled As Long, _
                ByVal remaining As Long, _
                ByVal avgFillPrice As Double, _
                ByVal permId As Long, _
                ByVal parentId As Long, _
                ByVal lastFillPrice As Double, _
                ByVal clientID As Long)

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

Private Const SERVER_VERSION As Long = 1

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
Private mConnectionRetryIntervalSecs As Long
Private mDisconnectedByApplication As Boolean

Private mSocketInputData() As Byte
Private mBufferOut() As Byte
Private mIndexOut As Long
Private mBufferIn() As Byte
Private mIndexIn As Long
Private mInputIndex As Long
Private mIndexInChk As Long
Private mInputIndexChk As Long

Private mServerVersion As Long
Private mCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer

Private mConnectionState As ConnectionStates

Private mTickerTable() As TickerTableEntry

Private mRequestedContracts As TradeBuildSP.IContracts
Private mFirstContractDetailsReceived As Boolean

Private mExchangeCodes() As String
Private mMaxExchangeCodesIndex As Long

Private mNextOrderID  As Long

Private mLogLevel As LogLevels
Private mTWSLogLevel As TWSLogLevels

Private mAllOrders As New Collection

Private mTwsHistDataRequestQueue As Queue
Private mConcurrentHistDataRequests As Long

Private mHistoricalDataTable() As HistoricalDataRequestTableEntry

'================================================================================
' External function declarations
'================================================================================

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                            Destination As Any, _
                            Source As Any, _
                            ByVal length As Long)
                            
Private Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                            Destination As Any, _
                            Source As Any, _
                            ByVal length As Long)
                            
'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

ReDim mTickerTable(InitialMaxTickers) As TickerTableEntry
ReDim mBufferIn(1023) As Byte
ReDim mBufferOut(1023) As Byte
mConnectionState = ConnectionStates.ConnNotConnected

Set mAllOrders = New Collection

mNextOrderID = 1

mLogLevel = LogLevelLow
mTWSLogLevel = TWSLogLevelError ' this is the TWS default

Set mTwsHistDataRequestQueue = New Queue
mTwsHistDataRequestQueue.initialCapacity = 20

ReDim mHistoricalDataTable(MaxConcurrentHistDataRequests) As HistoricalDataRequestTableEntry

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
addExchangeCode "PSE"

addExchangeCode "RDBK"

addExchangeCode "SBF"
addExchangeCode "SFB"
addExchangeCode "SGX"
addExchangeCode "SMART"
addExchangeCode "SNFE"
addExchangeCode "SOFFEX"
addExchangeCode "SWB"
addExchangeCode "SWX"

addExchangeCode "TSE"
addExchangeCode "TSE.JPN"

addExchangeCode "VENTURE"
addExchangeCode "VIRTX"
addExchangeCode "VWAP"

ReDim Preserve mExchangeCodes(mMaxExchangeCodesIndex) As String
End Sub


'================================================================================
' mConnectionTimer Event Handlers
'================================================================================

Private Sub mConnectionTimer_TimerExpired()
Connect
End Sub

'================================================================================
' Socket Event Handlers
'================================================================================

Private Sub mSocket_Close()
mConnectionState = ConnectionStates.ConnNotConnected
ReleaseSocket
If mDisconnectedByApplication Then
    RaiseEvent ConnectionToTWSClosed(False)
    logMessage "Connection to TWS closed: " & calcConnectionString, LogLevelLow
Else
    retryConnection
    RaiseEvent ConnectionToTWSClosed( _
                            IIf(mConnectionRetryIntervalSecs = 0, False, True))
    logMessage "Connection to TWS closed" & _
                IIf(mConnectionRetryIntervalSecs = 0, "", " (retrying)") & ": " & _
                calcConnectionString, _
                LogLevelLow
End If
End Sub

Private Sub mSocket_Connect()
mConnectionState = ConnectionStates.ConnConnecting
RaiseEvent Connecting
logMessage "Connecting to TWS: " & calcConnectionString, LogLevelLow
addData CLIENT_VERSION
send
addData mClientID
send
End Sub

Private Sub mSocket_DataArrival(ByVal bytesTotal As Long)

If bytesTotal = 0 Then
    checkmessages
    Exit Sub
End If

ReDim mSocketInputData(bytesTotal - 1) As Byte
mSocket.GetData mSocketInputData, vbArray + vbByte
RaiseEvent dataReceived

'Debug.Print "Socket: " & showBuffer(mSocketInputData)

If mIndexIn + bytesTotal > UBound(mBufferIn) Then
    If bytesTotal <= (UBound(mBufferIn) + 1 - mIndexIn + mInputIndex) Then
        If mInputIndex <= UBound(mBufferIn) Then
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

End Sub

Private Sub mSocket_Error(ByVal Number As Integer, _
                        Description As String, _
                        ByVal Scode As Long, _
                        ByVal Source As String, _
                        ByVal HelpFile As String, _
                        ByVal HelpContext As Long, _
                        CancelDisplay As Boolean)
Dim errorNum As MSWinsockLib.ErrorConstants

errorNum = Number
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
    sckHostNotFoundTryAgain
    
    If mConnectionState = ConnectionStates.ConnNotConnected Then
        ReleaseSocket
        logMessage "Failed to connect to TWS" & _
                    IIf(mConnectionRetryIntervalSecs = 0, "", " (retrying)") & ": " & _
                    Description & ": " & _
                    calcConnectionString, _
                    LogLevelLow
        RaiseEvent ConnectFailed(Description, _
                                IIf(mConnectionRetryIntervalSecs = 0, False, True))
        retryConnection
    Else
        mConnectionState = ConnectionStates.ConnNotConnected
        ReleaseSocket
        retryConnection
        logMessage "Connection to TWS closed: " & _
                    IIf(mConnectionRetryIntervalSecs = 0, "", " (retrying)") & ": " & _
                    Description & ": " & _
                    calcConnectionString, _
                    LogLevelLow
        RaiseEvent ConnectionToTWSClosed(IIf(mConnectionRetryIntervalSecs = 0, False, True))
    End If
'Case sckConnectionReset

'Case sckNoBufferSpace

'Case sckAlreadyConnected

'Case sckNotConnected

'Case sckSocketShutdown

'Case sckTimedout

'Case sckNotInitialized

'Case sckNonRecoverableError

'Case sckNoData
Case Else
    logMessage "RuntimeException in IBTWSSP.TWSAPI::mSocket_error" & vbCrLf & _
                "Socket error:" & vbCrLf & _
                "Number: " & Number & vbCrLf & _
                "Description: " & Description & vbCrLf & _
                "Scode: " & Scode, _
                LogLevelLow
    err.Raise ErrorCodes.RuntimeException, _
                "IBTWSSP.TWSAPI::mSocket_error", _
                "Socket error:" & vbCrLf & _
                "Number: " & Number & vbCrLf & _
                "Description: " & Description & vbCrLf & _
                "Scode: " & Scode
End Select

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

Friend Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Friend Property Get serverVersion() As Long
serverVersion = mServerVersion
End Property

Friend Property Get socketData() As String
socketData = showBuffer(mSocketInputData)
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
                If err.Number = 0 Then
                    ' request was still in queue - no further action required
                ElseIf err.Number <> ErrorCodes.IllegalArgumentException Then
                    err.Raise err.Number
                Else
                    ' request is with TWS so cancel it
                    
                    Const version As Long = 1
            
                    addData CANCEL_HISTORICAL_DATA
                    addData version
                    addData requestHandle + HistoricalDataTickerBase + i * HistoricalDataTickerIncrement
                    send
                    
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


        
End Sub

Friend Sub CancelMarketData(ByVal tickerid As Long)
cancelMarketDataEx tickerid, True
End Sub

Friend Sub CancelMarketDepth(ByVal tickerid As Long)
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::CancelMarketDepth", _
                "Not connected"
End If

Const version = 1

With mTickerTable(tickerid)
    If Not .inUse Then err.Raise ErrorCodes.InvalidTickerID, _
                                    "IBTWSSP.TWSAPI::cancelMarketDepth", _
                                    "Invalid ticker id " & tickerid
    If Not .receivingMarketDepth Then
        err.Raise ErrorCodes.NotReceivingMarketDepth, _
                    "IBTWSSP.TWSAPI::cancelMarketDepth", _
                    "Ticker id " & tickerid & " not currently receiving market depth"
    End If
    mTickerTable(tickerid).receivingMarketDepth = False
End With

addData CANCEL_MKT_DEPTH
addData version
addData MARKET_DEPTH_TICKER_BASE + tickerid
send

End Sub

Friend Sub cancelNewsBulletins()

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::cancelNewsBulletins", _
                "Not connected"
End If

Const version = 1

addData CANCEL_NEWS_BULLETINS
addData version
send
End Sub

Friend Sub cancelOrder(ByVal orderId As Long)
Dim lOrder As Order

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::cancelOrder", _
                "Not connected"
End If

Const version = 1
addData CANCEL_ORDER
addData version
addData orderId
send

End Sub

Friend Sub Connect()
On Error GoTo errNoSocket

mDisconnectedByApplication = False
createsocket

If mConnectionState <> ConnectionStates.ConnNotConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::connect", _
                "Already connected"
End If

On Error GoTo errNoConnect
mSocket.Connect mServer, mPort
Exit Sub

errNoSocket:
    err.Raise ErrorCodes.RuntimeException, _
                "IBTWSSP.TWSAPI::connect", _
                "Can't create socket"
Exit Sub

errNoConnect:
    err.Raise ErrorCodes.RuntimeException, _
                "IBTWSSP.TWSAPI::connect", _
                "Can't connect"
End Sub

Friend Sub disconnect()
mDisconnectedByApplication = True
If Not mConnectionTimer Is Nothing Then mConnectionTimer.StopTimer
mConnectionState = ConnectionStates.ConnNotConnected
releaseAllTickerTableEntries
ReleaseSocket
mIndexIn = 0
mIndexOut = 0
mInputIndex = 0
RaiseEvent ConnectionToTWSClosed(False)
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

'Friend Function newContractSpecifier( _
'                Optional ByVal localSymbol As String, _
'                Optional ByVal symbol As String, _
'                Optional ByVal exchange As String, _
'                Optional ByVal sectype As SecurityTypes, _
'                Optional ByVal currencyCode As String, _
'                Optional ByVal expiry As String, _
'                Optional ByVal strike As Double, _
'                Optional ByVal right As OptionRights = OptNone) As TradeBuildSP.IContractSpecifier
'
'Dim Source As String: Source = "IBTWSSP.TWSAPI::newContractSpecifier"
'
'If localSymbol = "" And symbol = "" Then
'    err.Raise ErrorCodes.IllegalArgumentException, _
'                Source, _
'                "Symbol must be supplied if localSymbol is not supplied"
'End If
'
'If exchange <> "" And _
'    Not IsValidExchangeCode(exchange) _
'Then
'    err.Raise ErrorCodes.IllegalArgumentException, _
'                Source, _
'                "'" & exchange & "' is not a valid exchange code"
'End If
'
'Select Case sectype
'Case 0  ' ie not supplied
'Case SecTypeStock
'Case SecTypeFuture
'    If expiry <> "" Then
'' need to write a function to validate the expiry format
''        err.Raise ErrorCodes.IllegalArgumentException, _
''                    source, _
''                    "expiry must be supplied"
'    End If
'Case SecTypeOption, SecTypeFuturesOption
'    If expiry <> "" Then
'' need to write a function to validate the expiry format
''        err.Raise ErrorCodes.IllegalArgumentException, _
''                    source, _
''                    "expiry must be supplied"
'    End If
'    If strike < 0 Then
'        err.Raise ErrorCodes.IllegalArgumentException, _
'                    Source, _
'                    "strike must be > 0"
'    End If
'    Select Case right
'    Case OptCall
'    Case OptPut
'    Case OptNone
'    Case Else
'        err.Raise ErrorCodes.IllegalArgumentException, _
'                    Source, _
'                    "'" & right & "' is not a valid option right"
'    End Select
'Case SecTypeCash
'Case SecTypeBag
'Case SecTypeIndex
'Case Else
'    err.Raise ErrorCodes.IllegalArgumentException, _
'                Source, _
'                "'" & sectype & "' is not a valid secType"
'End Select
'
'Set newContractSpecifier = New TradeBuildSP.IContractSpecifier
'With newContractSpecifier
'    .localSymbol = localSymbol
'    .symbol = symbol
'    .exchange = exchange
'    .sectype = sectype
'    .currencyCode = currencyCode
'    .expiry = expiry
'    .strike = strike
'    .right = right
'End With
'End Function

Friend Sub placeOrder(ByVal pOrder As Order)
Dim specifier As TradeBuildSP.IContractSpecifier

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::placeOrder", _
                "Already connected"
End If

pOrder.clientID = mClientID

Set specifier = pOrder.Contract.specifier

Const version = 11
addData PLACE_ORDER
addData version
addData pOrder.id

' send contract fields
With specifier
    addData .symbol
    addData secTypeToString(.sectype)
    addData Left$(.expiry, 6)
    addData .strike
    addData optionRightToString(.right)
    addData .exchange
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
        addData .auxPrice / 100
    Else
        addData .limitPrice
        addData .auxPrice
    End If
    ' send extended order fields
    addData orderTIFToString(.timeInForce)
    addData .ocaGroup
    addData .account
    addData .openClose
    addData .origin
    addData .orderRef
    addBoolean .transmit
    If mServerVersion >= 4 Then addData .parentId
    If mServerVersion >= 5 Then
        addBoolean .blockOrder
        addBoolean .sweepToFill
        addData .displaySize
        addData .triggerMethod
        addBoolean .ignoreRTH
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

If mServerVersion >= 9 Then addData pOrder.sharesAllocation

If mServerVersion >= 10 Then addData pOrder.discretionaryAmt

If mServerVersion >= 11 Then addData pOrder.goodAfterTime

If mServerVersion >= 12 Then addData pOrder.goodTillDate

If mServerVersion >= 13 Then
    addData pOrder.FAGroup
    addData pOrder.FAMethod
    addData pOrder.FAPercentage
    addData pOrder.FAProfile
End If
send
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

End Function

Friend Sub ReplaceFA(ByVal DataType As FADataTypes, ByVal xml As String)


If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::ReplaceFA", _
                "Already connected"
End If

Const version = 1

If mServerVersion < 13 Then
    err.Raise ErrorCodes.UnsupportedOperationException, _
                "IBTWSSP.TWSAPI::ReplaceFA", _
                "Server version does not support ReplaceFA"
End If

addData REPLACE_FA
addData version
addData DataType
addData xml

send
End Sub

Friend Sub requestAccountUpdates(ByVal subscribe As Boolean, ByVal acctCode As String)

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::requestAccountUpdates", _
                "Already connected"
End If

Const version = 2

addData REQ_ACCT_DATA
addData version
addBoolean subscribe

If mServerVersion >= 9 Then addData acctCode

send
End Sub

Friend Sub requestAllOpenOrders()

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::requestAllOpenOrders", _
                "Not connected"
End If
                            

Const version = 1

addData REQ_ALL_OPEN_ORDERS
addData version
send
End Sub

Friend Sub requestAutoOpenOrders(ByVal autoBind As Boolean)

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::requestAutoOpenOrders", _
                "Not connected"
End If

Const version = 1

addData REQ_AUTO_OPEN_ORDERS
addData version
addBoolean autoBind
send
End Sub

Friend Sub RequestContract(ByVal pContracts As TradeBuildSP.IContracts)

Dim contractSpecifier As TradeBuildSP.IContractSpecifier

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::RequestContract", _
                "Not connected"
End If

' don't send another request if there's already one outstanding
If Not mRequestedContracts Is Nothing Then
    err.Raise ContractDetailsReqNotAllowed, _
                "TradeBuildAPI::RequestContract", _
                "An existing request has not yet completed"
End If

Set contractSpecifier = pContracts.contractSpecifier
contractSpecifier.Locked = True

Select Case contractSpecifier.sectype
Case SecurityTypes.SecTypeBag
Case SecurityTypes.SecTypeCash
Case SecurityTypes.SecTypeFuture
Case SecurityTypes.SecTypeFuturesOption
Case SecurityTypes.SecTypeIndex
Case SecurityTypes.SecTypeOption
Case SecurityTypes.SecTypeStock
Case Else
    RaiseEvent ContractSpecifierInvalid("security type must be supplied", pContracts)
    Exit Sub
End Select

If contractSpecifier.localSymbol = "" And _
    contractSpecifier.symbol = "" _
Then
    RaiseEvent ContractSpecifierInvalid("either symbol or local symbol must be supplied", pContracts)
    Exit Sub
End If
    

Set mRequestedContracts = pContracts

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

Friend Sub RequestExecutions(ByVal filter As ExecutionFilter)
Dim theFilter As ExecutionFilter

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
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
End Sub

Friend Sub requestFA(ByVal DataType As FADataTypes)

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::requestFA", _
                "Not connected"
End If

Const version = 1

If mServerVersion < 13 Then
    err.Raise ErrorCodes.UnsupportedOperationException, _
                "IBTWSSP.TWSAPI::requestFA", _
                "Server version does not support requestFA"
End If

addData REQ_FA
addData version
addData DataType

send
End Sub

Friend Sub requestManagedAccounts()

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::requestManagedAccounts", _
                "Not connected"
End If

Const version = 1

addData REQ_MANAGED_ACCTS
addData version
send
End Sub

Friend Function RequestMarketDataEx(ByVal pContractSpecifier As TradeBuildSP.IContractSpecifier, _
                            ByVal dataConsumer As IStreamingDataConsumer, _
                            Optional ByVal includeMarketDepth As Boolean = False) As Long

Dim lTickerID As Long

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::RequestMarketDataEx", _
                "Not connected"
End If

pContractSpecifier.Locked = True

lTickerID = AllocateTickerTableEntry(dataConsumer)
RequestMarketDataEx = lTickerID
Set mTickerTable(lTickerID).contractSpecifier = pContractSpecifier
reqMarketData lTickerID

If includeMarketDepth Then
    mTickerTable(lTickerID).receivingMarketDepth = True
    reqMarketDepth lTickerID, pContractSpecifier
End If

End Function

Friend Sub RequestMarketDepth(ByVal pTickerID As Long)
                            
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::RequestMarketDepth", _
                "Not connected"
End If

With mTickerTable(pTickerID)
    If .contractSpecifier Is Nothing Then err.Raise ErrorCodes.InvalidTickerID, "IBTWSSP.TWSAPI::requestMarketDepth", "Invalid ticker id: " & pTickerID
    .receivingMarketDepth = True
    reqMarketDepth pTickerID, .contractSpecifier
End With

End Sub

Friend Function RequestHistoricalData( _
                barSpecifier As TradeBuildSP.BarDataSpecifier, _
                ByVal historicalDataReader As HistDataReader) As Long

Dim requestID As Long
Dim i As Long

requestID = -1

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::RequestHistoricalData", _
                "Not connected"
End If

If mServerVersion < 16 Then
    err.Raise ErrorCodes.UnsupportedOperationException, _
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
End Function

Friend Sub RequestNewsBulletins(allMsgs As Boolean)

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::RequestNewsBulletins", _
                "Not connected"
End If

Const version = 1

addData REQ_NEWS_BULLETINS
addData version
addData allMsgs
send
End Sub

Friend Sub RequestOpenOrders()

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::RequestOpenOrders", _
                "Not connected"
End If

Const version = 1

addData REQ_OPEN_ORDERS
addData version
send
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
If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::cancelMarketDataEx", _
                "Not connected"
End If
Dim id As Long
Const version = 2

If Not mTickerTable(tickerid).inUse Then
    err.Raise InvalidTickerID, _
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

Private Sub checkmessages()
Static checkPointNumber As Long
Static alreadyInvoked As Boolean
Static msgID As TWSSocketInMsgTypes
Static version As Long
Dim timeStampStr As String
Static retryCount As Long

If alreadyInvoked Then Exit Sub
alreadyInvoked = True

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
    
    ' set the log level in TWS
    setTWSLogLevel mTWSLogLevel
    
    ' now restart any tickers that were previously in operation
    reEstablishMarketData True
    
    ' and re-request any market depth
    resetMarketDepth True
    
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
        disconnect
        Debug.Print s
        err.Raise ErrorCodes.RuntimeException, _
                    "IBTWSSP.TWSAPI::checkmessages", _
                    s
    End Select
    retryCount = 0
    checkpointBuffer
    checkPointNumber = 1000
    Debug.Print "Set checkpoint number=" & 1000
Loop

alreadyInvoked = False

Exit Sub

err:
If err.Number = DataIncomplete Then
    rollbackBuffer
    alreadyInvoked = False
    retryCount = retryCount + 1
    Debug.Print "Data incomplete count: " & retryCount & "; CheckPointNumber: " & checkPointNumber
    Exit Sub
End If
alreadyInvoked = False
err.Raise err.Number
End Sub

Private Sub checkpointBuffer()
'mIndexInChk = mIndexIn
mInputIndexChk = mInputIndex
End Sub

Private Sub createsocket()
If mSocket Is Nothing Then
    Set mSocket = New MSWinsockLib.Winsock
End If
End Sub

Private Function execSideFromString(ByVal value As String) As ExecSides
Select Case UCase$(value)
Case "BOT"
    execSideFromString = SideBuy
Case "SLD"
    execSideFromString = SideSell
End Select
End Function

Private Function execSideToString(ByVal value As ExecSides) As String
Select Case value
Case SideBuy
    execSideToString = "BOT"
Case SideSell
    execSideToString = "SLD"
End Select
End Function

Private Sub generateTwsHistoricalDataRequest(ByVal index As Long)
Dim TwsRequest As TWSHistoricalDataRequestQueueEntry
Dim reqNum As Long
Dim totalMinsReqd As Long

reqNum = mHistoricalDataTable(index).numberOfTwsRequests + 1
mHistoricalDataTable(index).numberOfTwsRequests = reqNum

TwsRequest.id = index + HistoricalDataTickerBase + (mHistoricalDataTable(index).numberOfTwsRequests - 1) * HistoricalDataTickerIncrement

ReDim mHistoricalDataTable(index).requestKeys(reqNum - 1) As String

Set TwsRequest.contractSpec = mHistoricalDataTable(index).barSpecifier.Contract.specifier

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
'debug.Print "Data incomplete"
'If i >= 487 Then Stop
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
                    timestamp
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
                    timestamp
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

RaiseEvent ContractSpecifierInvalid("bonds not yet supported", mRequestedContracts)

End Sub

Private Sub getContractData(ByVal version As Long)
Dim listendata As TWSSocketMsgContractData
Dim timestamp As Date
Dim lContract As TradeBuildSP.IContract
Dim minimumTick As Double
Dim multiplier As Long
Dim OrderTypes As String
Dim orderTypesStr() As String
Dim lOrderTypes() As TradeBuildSP.OrderTypes
Dim validExchanges As String
Dim lValidExchanges() As String
Dim i As Long
Dim j As Long

timestamp = gCurrentTime

Set lContract = mCommonServiceConsumer.NewContract
With lContract
    .specifier = mCommonServiceConsumer.NewContractSpecifier
    .specifier.symbol = getString
    .specifier.sectype = secTypeFromString(getString)
    .specifier.expiry = getString
    .specifier.strike = getDouble
    .specifier.right = optionRightFromString(getString)
    .specifier.exchange = getString
    .specifier.currencyCode = getString
    .specifier.localSymbol = getString
    .marketName = getString
    .tradingClass = getString
    .contractID = getLong
    minimumTick = getDouble
    multiplier = getLong
    If .specifier.currencyCode = "GBP" Then
        .multiplier = multiplier / 100
        .minimumTick = minimumTick * 100
    Else
        .minimumTick = minimumTick
        .multiplier = multiplier
    End If
    
    OrderTypes = getString
    orderTypesStr = Split(OrderTypes, ",")
    
    ReDim lOrderTypes(UBound(orderTypesStr)) As TradeBuildSP.OrderTypes
    
    For i = 0 To UBound(orderTypesStr)
        lOrderTypes(j) = orderTypeFromString(orderTypesStr(i))
        
        If Not lOrderTypes(j) = TradeBuildSP.OrderTypes.OrderTypeNone Then
            j = j + 1
        Else
            Select Case orderTypesStr(i)
            Case "DAY"  ' today only
            Case "GAT"  ' good after time
            Case "GTC"  'good till cancelled
            Case "GTD"  ' good till date
            Case "GTT"  ' good till time
            Case "ICE"  ' iceberg
            Case "IOC"  ' immediate or cancel
            Case "RFQ"  ' request for quote
            Case Else
                If orderTypesStr(i) <> "" Then
                    mCommonServiceConsumer.ServiceProviderError ErrorCodes.UnknownOrderTypeFromTWS, _
                                            "Unknown order type " & orderTypesStr(i) & " received from TWS", _
                                            SpName
                End If
            End Select
        End If
    Next
    
    If j <> (UBound(orderTypesStr) - 1) Then
        ReDim Preserve lOrderTypes(j - 1) As TradeBuildSP.OrderTypes
    End If
    
    .OrderTypes = lOrderTypes
    validExchanges = getString
    lValidExchanges = Split(validExchanges, ",")
    .validExchanges = lValidExchanges
    
    If version >= 2 Then
        .PriceMagnifier = getLong
    End If
End With

processContract lContract

If mLogLevel >= LogLevelHigh Then
    With lContract
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
                    timestamp
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

id = getLong
Debug.Print "Error id: " & id
errorCode = getLong
Debug.Print "Error code: " & errorCode
errorMsg = getString
Debug.Print "Error msg: " & errorMsg
timestamp = gCurrentTime


Select Case errorCode
Case 162    ' historical data request problem
    
    index = (id - HistoricalDataTickerBase) Mod HistoricalDataTickerIncrement

    mHistoricalDataTable(index).historicalDataReader.historicalDataRequestFailed errorMsg
    releaseHistoricalDataTableEntry index
    processNextHistDataRequest
Case 200    ' security definition not known
    If id >= 0 Then
        ' this must relate to a reqMktData - this should never occur as
        ' we make sure the contract is valid before starting the ticker
        
        '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        ' the following line is commented out because we need to do something here
        ' that will be picked up by the realtime data provider, not the contract
        ' info provider
        'RaiseEvent ContractInvalid
        releaseTickerTableEntry (id Mod INCARNATION_INCREMENT)
    Else
        ' this must relate to a RequestContract
        processContractError errorMsg
    End If
Case 202
    RaiseEvent orderStatus(id, OrderStatuses.OrderStatusCancelled, 0, 0, 0, 0, 0, 0, mClientID)
Case 309
    ' max market depth requests exceeded
    
    ' need to do something about the following !!!!!!!!!!!!!!!!!!!!!!!!!
    'mTickerTable(id - MARKET_DEPTH_TICKER_BASE).tickerObj.MarketDepthNotAvailable errorMsg
    'mTickers.MarketDepthNotAvailable gCurrentTime, id - MARKET_DEPTH_TICKER_BASE, errorMsg
Case 316
    resetMarketDepth reEstablish:=True
Case 317
    resetMarketDepth reEstablish:=False
Case 321    ' invalid request
    If id >= HistoricalDataTickerBase And id < OrderIdBase Then
        index = (id - HistoricalDataTickerBase) Mod HistoricalDataTickerIncrement
    
        If index > 0 Then
            mHistoricalDataTable(id).historicalDataReader.historicalDataRequestInvalid errorMsg
            releaseHistoricalDataTableEntry index
            processNextHistDataRequest
        End If
    End If
Case 1100
    ' connectivity between TWS and IB has been lost
    RaiseEvent ConnectionToIBClosed
Case 1101
    ' connectivity between TWS and IB has been restored, but data has been
    ' lost, so need to reestablish market data and market depth requests
    
    reEstablishMarketData False
    resetMarketDepth True
    
    ' Now need to reassociate order ids with tws
    RequestOpenOrders
    
    ' Now make sure we're aware of any executions that occurred during
    ' the disconnection
    Set execFilter = New ExecutionFilter
    execFilter.clientID = mClientID
    RequestExecutions execFilter
Case 1102
    ' connectivity between TWS and IB has been restored without loss of data
    ' Now need to reassociate order ids with tws
    RequestOpenOrders
    
    ' Now make sure we're aware of any executions that occurred during
    ' the disconnection
    Set execFilter = New ExecutionFilter
    execFilter.clientID = mClientID
    RequestExecutions execFilter

    RaiseEvent ConnectionToIBRecovered
    
Case Else
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
                    timestamp
End If

End Sub

Private Sub getExecutionData(ByVal version As Long)
Dim listendata As TWSSocketMsgExecutionData
Dim timestamp As Date: timestamp = gCurrentTime
Dim sectype As String
Dim right As String
Dim side As String
Dim price As Double

Dim id As Long: id = getLong

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
    .localSymbol = getString
End With

Dim exec As Execution
Set exec = New Execution
With exec
    .orderId = id
    .execId = getString
    .time = getString
    .accountNumber = getString
    .exchange = getString
    side = getString
    .side = execSideFromString(side)
    .quantity = getLong
    price = getDouble
    If lContractSpecifier.currencyCode = "GBP" Then
        .price = price * 100
    Else
        .price = price
    End If
End With

If version >= 2 Then exec.permId = getLong

If version >= 3 Then exec.clientID = getLong
   
RaiseEvent ExecutionDetails(id, lContractSpecifier, exec)

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.EXECUTION_DATA
    With lContractSpecifier
        listendata.currencyCode = .currencyCode
        listendata.exchange = .exchange
        listendata.expiry = .expiry
        listendata.localSymbol = .localSymbol
        listendata.right = right
        listendata.sectype = sectype
        listendata.strike = .strike
        listendata.symbol = .symbol
    End With
    With exec
        listendata.accountNumber = .accountNumber
        listendata.clientID = .clientID
        listendata.execExchange = .exchange
        listendata.execId = .execId
        listendata.orderId = id
        listendata.permId = .permId
        listendata.price = price
        listendata.quantity = .quantity
        listendata.side = side
        listendata.time = .time
    End With
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgExecutionData, _
                    timestamp
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
                    timestamp
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
                    timestamp
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
                    timestamp
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
                    timestamp
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
                    timestamp
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
                    timestamp
End If

If mNextOrderID < OrderIdBase Then mNextOrderID = OrderIdBase

End Sub

Private Sub getOpenOrder(ByVal version As Long)
Dim listendata As TWSSocketMsgOpenOrder
Dim timestamp As Date: timestamp = gCurrentTime
Dim lOrder As Order
Dim id As Long
Dim sectype As String
Dim right As String
Dim action As String
Dim orderType As String
Dim timeInForce As String

id = getLong

On Error Resume Next
Set lOrder = mAllOrders(CStr(id))
On Error GoTo 0

If lOrder Is Nothing Then
    ' this must be an open order from the previous API session, but
    ' for which we haven't yet had the openOrder notification
    Set lOrder = New Order
    lOrder.id = id
    mAllOrders.add lOrder, CStr(id)
End If

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
With lOrder
    action = getString
    .action = orderActionFromString(action)
    .quantity = getLong
    orderType = getString
    .orderType = orderTypeFromString(orderType)
    .limitPrice = getDouble
    .auxPrice = getDouble
    timeInForce = getString
    .timeInForce = orderTIFFromString(timeInForce)
    .ocaGroup = getString
    .account = getString
    .openClose = getString
    .origin = getLong
    .orderRef = getString
    If version >= 3 Then .clientID = getLong
    
    If version >= 4 Then
        .permId = getLong
        .ignoreRTH = getBoolean
        .hidden = getBoolean
        .discretionaryAmt = getDouble
    End If

    If version >= 5 Then .goodAfterTime = getString
    
    If version >= 6 Then .sharesAllocation = getString
    
    If version >= 7 Then
        .FAGroup = getString
        .FAMethod = getString
        .FAPercentage = getString
        .FAProfile = getString
    End If
    
End With

RaiseEvent openOrder(lContractSpecifier, lOrder.clone)

If mLogLevel >= LogLevelHigh Then
    listendata.msgType = TWSSocketInMsgTypes.OPEN_ORDER
    With lContractSpecifier
        listendata.currencyCode = .currencyCode
        listendata.exchange = .exchange
        listendata.expiry = .expiry
        listendata.localSymbol = .localSymbol
        listendata.right = .right
        listendata.sectype = .sectype
        listendata.strike = .strike
        listendata.symbol = .symbol
    End With
    With lOrder
        listendata.account = .account
        listendata.action = action
        listendata.auxPrice = .auxPrice
        listendata.clientID = .clientID
        listendata.discretionaryAmt = .discretionaryAmt
        listendata.FAGroup = .FAGroup
        listendata.FAMethod = .FAMethod
        listendata.FAPercentage = .FAPercentage
        listendata.FAProfile = .FAProfile
        listendata.goodAfterTime = .goodAfterTime
        listendata.hidden = .hidden
        listendata.ignoreRTH = .ignoreRTH
        listendata.limitPrice = .limitPrice
        listendata.ocaGroup = .ocaGroup
        listendata.openClose = .openClose
        listendata.orderId = id
        listendata.orderRef = .orderRef
        listendata.orderType = orderType
        listendata.origin = .origin
        listendata.permId = .permId
        listendata.quantity = .quantity
        listendata.sharesAllocation = .sharesAllocation
        listendata.timeInForce = timeInForce
    End With
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgOpenOrder, _
                    timestamp
End If

End Sub

Private Sub getOrderStatus(ByVal version As Long)
Dim listendata As TWSSocketMsgOrderStatus
Dim timestamp As Date: timestamp = gCurrentTime
Dim id As Long: id = getLong
Dim status As String: status = getString
Dim filled As Long: filled = getLong
Dim remaining As Long: remaining = getLong
Dim avgFillPrice As Double:  avgFillPrice = getDouble
Dim lastFillPrice As Double
Dim clientID As Long
Dim permId As Long
Dim parentId As Long
Dim lOrder As Order
Dim dontNotify As Boolean

If version >= 2 Then permId = getLong

If version >= 3 Then parentId = getLong

If version >= 4 Then lastFillPrice = getDouble

If version >= 5 Then clientID = getLong

On Error Resume Next
Set lOrder = mAllOrders(CStr(id))
On Error GoTo 0
If lOrder Is Nothing Then
    ' this must be an open order from the previous API session, but
    ' for which we haven't yet had the openOrder notification
    Set lOrder = New Order
    lOrder.id = id
    mAllOrders.add lOrder, CStr(id)
    ' don't notify the user program because it won't know about this
    ' order yet
    dontNotify = False
End If

lOrder.status = orderStatusFromString(status)
lOrder.quantityFilled = filled
lOrder.quantity = remaining
lOrder.averagePrice = avgFillPrice
lOrder.permId = permId
lOrder.lastFillPrice = lastFillPrice
lOrder.clientID = clientID

If Not dontNotify Then
    RaiseEvent orderStatus(id, _
                            orderStatusFromString(status), _
                            filled, _
                            remaining, _
                            avgFillPrice, _
                            permId, _
                            parentId, _
                            lastFillPrice, _
                            clientID)
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
    listendata.status = status
    listendata.version = version
    
    mCommonServiceConsumer.NotifyListeners listendata, _
                    ListenValueTypes.ValueTypeTWSSocketMsgOrderStatus, _
                    timestamp
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
                    timestamp
End If

End Sub

Private Sub getScannerData(ByVal version As Long)
Dim Contract As TradeBuildSP.IContract
Dim tickerid As Long: tickerid = getLong
Dim numberOfElements As Long: numberOfElements = getLong
Dim rank As Long
Dim distance As String
Dim benchmark As String
Dim projection As String
Dim i As Long

For i = 0 To numberOfElements - 1
    Set Contract = mCommonServiceConsumer.NewContract
    rank = getLong
    With Contract.specifier
        .symbol = getString
        .sectype = secTypeFromString(getString)
        .expiry = getString
        .strike = getDouble
        .right = optRightFromString(getString) 'right
        .exchange = getString
        .currencyCode = getString
        .localSymbol = getString
    End With
    Contract.marketName = getString
    Contract.tradingClass = getString
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
tickerid = getLong
tickType = getLong
price = getDouble
If version >= 2 Then size = getLong
timestamp = gCurrentTime

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
                    timestamp
End If

End Sub

Private Sub getTickSize(ByVal version As Long)
Dim listendata As TWSSocketMsgTickSize
Dim timestamp As Date
Dim size As Long
Dim tickType As Long
Dim tickerid As Long
tickerid = getLong
tickType = getLong
size = getLong
timestamp = gCurrentTime

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
                    timestamp
End If

End Sub

Private Sub logMessage( _
                ByVal message As String, _
                ByVal logLevel As LogLevels)
If mLogLevel < logLevel Or _
    mLogLevel = LogLevelNone Then Exit Sub
mCommonServiceConsumer.NotifyListeners message, _
                                        TradeBuildSP.StandardListenValueTypes.Log, _
                                        gCurrentTime
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
Case "QUOTE"
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
Case "LOO"
    orderTypeFromString = TradeBuildSP.OrderTypes.OrderTypeLimitOnOpen
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
    err.Raise ErrorCodes.InvalidOrderType
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

Private Sub processContract(ByVal pContract As TradeBuildSP.IContract)

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

mRequestedContracts.AddContract pContract

End Sub

Private Sub processContractError(ByVal reason As String)
Dim tickerid As Long
Dim lContract As TradeBuildSP.IContract

If mRequestedContracts.Count > 0 Then
    ' we are processing the forced invalid contract details request
    ' for a request for which we have received one or more contract details
    RaiseEvent ContractsLoaded
Else
    ' this is a genuine error response to a contract details request
    RaiseEvent ContractSpecifierInvalid(reason, mRequestedContracts)
End If
Set mRequestedContracts = Nothing
        
mFirstContractDetailsReceived = False

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
End With
End Sub

Private Sub reqHistoricalData( _
                request As TWSHistoricalDataRequestQueueEntry)
Dim i As Long

Const version = 3

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
                "IBTWSSP.TWSAPI::RequestHistoricalData", _
                "Not connected"
End If

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

Private Sub setTWSLogLevel(logLevel As TWSLogLevels)

If mConnectionState <> ConnectionStates.ConnConnected Then
    err.Raise ErrorCodes.IllegalStateException, _
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




