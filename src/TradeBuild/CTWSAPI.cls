VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradeBuildAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

'Implements TradeBuildSP.IStreamingDataConsumer

'================================================================================
' Events
'================================================================================

Event connected(ByVal timestamp As Date)
    ' Raised when a successful connection to TWS has been achieved.

Event connectFailed(ByVal timestamp As Date, _
                ByVal Description As String, _
                ByVal retrying As Boolean)
    ' Raised when connection to TWS has failed and the connection
    ' retry interval is zero.

Event connecting(ByVal timestamp As Date)
    ' Raised when an attempt to connect to TWS is initiated.

Event connectionToTWSClosed( _
                ByVal timestamp As Date, _
                ByVal reconnecting As Boolean)
    ' Raised when the connection to TWS is disconnected.

Event connectionToIBClosed(ByVal timestamp As Date)

Event connectionToIBRecovered(ByVal timestamp As Date)

Event Contract( _
                ByVal timestamp As Date, _
                ByVal pContract As Contract)
    ' Raised when details for a security are received from TWS.

Event ContractInvalid( _
                ByVal timestamp As Date, _
                ByVal pContractSpecifier As contractSpecifier)

Event dataReceived(ByVal timestamp As Date)
    ' Raised when a new block of data is received from the socket connecting to TWS.
    
Event errorMessage( _
                ByVal timestamp As Date, _
                ByVal id As Long, _
                ByVal errorCode As ApiErrorCodes, _
                ByVal errorMsg As String)

Event ExecutionDetails( _
                ByVal timestamp As Date, _
                ByVal id As Long, _
                ByVal pContractSpecifier As contractSpecifier, _
                ByVal exec As Execution)

Event FAData( _
                ByVal timestamp As Date, _
                ByVal DataType As FADataTypes, _
                ByVal XMLData As String)
                
Event managedAccounts( _
                ByVal timestamp As Date, _
                ByVal accountsList As String)

Event openOrder( _
                ByVal timestamp As Date, _
                ByVal pContractSpecifier As contractSpecifier, _
                ByVal pOrder As Order)

Event orderStatus( _
                ByVal timestamp As Date, _
                ByVal id As Long, _
                ByVal status As OrderStatuses, _
                ByVal filled As Long, _
                ByVal remaining As Long, _
                ByVal avgFillPrice As Double, _
                ByVal permId As Long, _
                ByVal parentId As Long, _
                ByVal lastFillPrice As Double, _
                ByVal clientID As Long)

Event preFill( _
                ByVal timestamp As Date, _
                ByVal Order As Order, _
                ByVal pContractSpecifier As contractSpecifier, _
                ByRef fillPrice As Double, _
                ByRef fillSize As Long)

Event updateAccountTime(ByVal timestamp As Date, _
                ByVal accountTime As String)

Event updateAccountValue( _
                ByVal timestamp As Date, _
                ByVal key As String, _
                ByVal value As String, _
                ByVal currencyCode As String, _
                ByVal accountName As String)

Event updateMarketDepth( _
                ByVal timestamp As Date, _
                ByVal id As Long, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)
                
Event UpdateMarketDepthL2( _
                ByVal timestamp As Date, _
                ByVal id As Long, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal Size As Long)

Event updateNewsBulletin( _
                ByVal timestamp As Date, _
                ByVal msgID As Integer, _
                ByVal msgType As BulletinTypes, _
                ByVal message As String, _
                ByVal origExchange As String)
                
Event updatePortfolioValue( _
                ByVal timestamp As Date, _
                ByVal pContractSpecifier As contractSpecifier, _
                ByVal position As Long, _
                ByVal marketPrice As Double, _
                ByVal marketValue As Double, _
                ByVal averageCost As Double, _
                ByVal unrealizedPNL As Double, _
                ByVal realizedPNL As Double, _
                ByVal accountName As String)

'================================================================================
' Types
'================================================================================

Private Type TickerTableEntry
    incarnation             As Long
    contractSpecifier       As contractSpecifier
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
' Constants
'================================================================================
Private Const CLIENT_VERSION As Long = 16
Private Const SERVER_VERSION As Long = 1

Private Const INCARNATION_INCREMENT As Long = 1024&

Private Const MARKET_DEPTH_TICKER_BASE As Long = 10000&

'================================================================================
' Member variables
'================================================================================

Private mfWinsock As fWinsock
Private WithEvents mSocket As Winsock
Attribute mSocket.VB_VarHelpID = -1
Private WithEvents mConnectionTimer As IntervalTimer
Attribute mConnectionTimer.VB_VarHelpID = -1
Private mConnectionRetryIntervalSecs As Long
Private mServer As String
Private mPort As Long
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

Private mClientID As Long

Private mConnectionState As ConnectionStates

Private mTickerTable() As TickerTableEntry

Private mContractDetailsRequest As ContractDtlsReq
Private mFirstContractDetailsReceived As Boolean

Private WithEvents mOrderSimulator As AdvancedOrderSimulator
Attribute mOrderSimulator.VB_VarHelpID = -1
Private mSimulateOrders As Boolean

Private mServiceProviderError As ServiceProviderError

Private mListeners As listeners
Private mEnableListeners As Boolean

Private mTickers As Tickers

Private mExchangeCodes() As String
Private mMaxExchangeCodesIndex As Long

'================================================================================
' Enums
'================================================================================

Private Enum InternalErrorCodes
    DataIncomplete = vbObjectError + 512
End Enum


'================================================================================
' Class Event Handlers
'================================================================================
Private Sub Class_Initialize()

Set gServiceProviders = New ServiceProviders
Set gTradeBuildAPI = Me

ReDim mTickerTable(InitialMaxTickers) As TickerTableEntry
ReDim mBufferIn(1023) As Byte
ReDim mBufferOut(1023) As Byte
Set gAllOrders = New Collection
mConnectionState = ConnectionStates.ConnNotConnected
Set mOrderSimulator = New AdvancedOrderSimulator
Set gOrderSimulator = mOrderSimulator
Set mListeners = New listeners
Set gListeners = mListeners
Set mTickers = New Tickers
Set gTickers = mTickers

gNextOrderID = 1

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
Connect mServer, mPort, mClientID
End Sub

'================================================================================
' Socket Event Handlers
'================================================================================

Private Sub mSocket_Close()
mConnectionState = ConnectionStates.ConnNotConnected
ReleaseSocket
If mDisconnectedByApplication Then
    RaiseEvent connectionToTWSClosed(gCurrentTime, False)
Else
    retryConnection
    RaiseEvent connectionToTWSClosed( _
                            gCurrentTime, _
                            IIf(mConnectionRetryIntervalSecs = 0, False, True))
End If
End Sub

Private Sub mSocket_Connect()
Dim timestamp As Date: timestamp = gCurrentTime
mConnectionState = ConnectionStates.ConnConnecting
RaiseEvent connecting(timestamp)
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
RaiseEvent dataReceived(gCurrentTime)

If mIndexIn + bytesTotal > UBound(mBufferIn) Then
    If mInputIndex > (mIndexIn + bytesTotal - UBound(mBufferIn) - 1) Then
        If mInputIndex < UBound(mBufferIn) Then
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
Dim timestamp As Date: timestamp = gCurrentTime

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
        RaiseEvent connectFailed(timestamp, _
                                Description, _
                                IIf(mConnectionRetryIntervalSecs = 0, False, True))
        retryConnection
    Else
        SetServiceProviderError errorNum, Description, "IB TWS"
        mConnectionState = ConnectionStates.ConnNotConnected
        ReleaseSocket
        retryConnection
        RaiseEvent connectionToTWSClosed( _
                                gCurrentTime, _
                                IIf(mConnectionRetryIntervalSecs = 0, False, True))
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
    SetServiceProviderError errorNum, Description, "IB TWS"
End Select

End Sub

'================================================================================
' Order Simulator Event Handlers
'================================================================================

Private Sub mOrderSimulator_errorMessage(ByVal timestamp As Date, _
                                        ByVal id As Long, _
                                        ByVal errorCode As Long, _
                                        ByVal errorMsg As String)
RaiseEvent errorMessage(timestamp, _
                        id, _
                        errorCode, _
                        errorMsg)
End Sub

Private Sub mOrderSimulator_executionDetails(ByVal timestamp As Date, _
                                            ByVal id As Long, _
                                            ByVal pContractSpecifier As contractSpecifier, _
                                            ByVal exec As Execution)
RaiseEvent ExecutionDetails(timestamp, _
                            id, _
                            pContractSpecifier, _
                            exec)
End Sub

Private Sub mOrderSimulator_openOrder(ByVal timestamp As Date, _
                                    ByVal pContractSpecifier As contractSpecifier, _
                                    ByVal pOrder As Order)
RaiseEvent openOrder(timestamp, _
                    pContractSpecifier, _
                    pOrder)
End Sub

Private Sub mOrderSimulator_orderStatus(ByVal timestamp As Date, _
                                        ByVal id As Long, _
                                        ByVal status As OrderStatuses, _
                                        ByVal filled As Long, _
                                        ByVal remaining As Long, _
                                        ByVal avgFillPrice As Double, _
                                        ByVal permId As Long, _
                                        ByVal parentId As Long, _
                                        ByVal lastFillPrice As Double, _
                                        ByVal clientID As Long)
RaiseEvent orderStatus(timestamp, _
                        id, _
                        status, _
                        filled, _
                        remaining, _
                        avgFillPrice, _
                        permId, _
                        parentId, _
                        lastFillPrice, _
                        clientID)
End Sub

Private Sub mOrderSimulator_preFill(ByVal timestamp As Date, _
                                    ByVal Order As Order, _
                                    ByVal pContractSpecifier As contractSpecifier, _
                                    fillPrice As Double, _
                                    fillSize As Long)
RaiseEvent preFill(timestamp, _
                    Order, _
                    pContractSpecifier, _
                    fillPrice, _
                    fillSize)
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Public Property Let ConnectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
If mConnectionRetryIntervalSecs = 0 Then Set mConnectionTimer = Nothing
End Property

Public Property Get connectionState() As ConnectionStates
connectionState = mConnectionState
End Property

Public Property Let EnableListeners(ByVal value As Boolean)
mEnableListeners = value
End Property

Public Property Get EnableListeners() As Boolean
EnableListeners = mEnableListeners
End Property

Friend Property Get InBuffer() As String
InBuffer = showBuffer(mBufferIn)
End Property

Public Property Get serverVersion() As Long
serverVersion = mServerVersion
End Property

Public Property Get ServiceProviders() As ServiceProviders
Set ServiceProviders = gServiceProviders
End Property

Public Property Let simulateOrders(ByVal value As Boolean)
mSimulateOrders = value
End Property

Public Property Get simulateOrders() As Boolean
simulateOrders = mSimulateOrders
End Property

Public Property Get socketData() As String
socketData = showBuffer(mSocketInputData)
End Property

Public Property Get SupportedInputTickfileFormats() As TradeBuild.TickfileFormatSpecifier()
SupportedInputTickfileFormats = gServiceProviders.SupportedInputTickfileFormats
End Property

Public Property Get SupportedOutputTickfileFormats() As TradeBuild.TickfileFormatSpecifier()
SupportedOutputTickfileFormats = gServiceProviders.SupportedOutputTickfileFormats
End Property

Public Property Get Tickers() As Tickers
Set Tickers = mTickers
End Property

Friend Property Get TickerTableSize() As Long
TickerTableSize = UBound(mTickerTable) + 1
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddListener(ByVal newListener As IListener, _
                            Optional ByRef valueType As Long = ListenerValueTypes.ValueTypesAll) As String
mListeners.add newListener, valueType
End Function

Public Sub CancelMarketData(ByVal tickerid As Long)
If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            tickerid, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If
Dim id As Long
Const Version = 2

If Not mTickerTable(tickerid).inUse Then
    err.Raise InvalidTickerID, _
    "TradeBuild.TradeBuildAPI::CancelMarketData", _
    "Invalid ticker id " & tickerid
End If

id = tickerid + mTickerTable(tickerid).incarnation

If mTickerTable(tickerid).receivingMarketDepth Then CancelMarketDepth tickerid
releaseTickerTableEntry tickerid
addData CANCEL_MKT_DATA
addData Version
addData id
send

End Sub

Public Sub CancelMarketDepth(ByVal tickerid As Long)
If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                tickerid, _
                ApiErrorCodes.NotConnected, _
                LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 1

With mTickerTable(tickerid)
    If Not .inUse Then err.Raise ErrorCodes.InvalidTickerID, _
                                    "TradeBuild.TradeBuildAPI::cancelMarketDepth", _
                                    "Invalid ticker id " & tickerid
    If Not .receivingMarketDepth Then
        err.Raise ErrorCodes.NotReceivingMarketDepth, _
                    "TradeBuild.TradeBuildAPI::cancelMarketDepth", _
                    "Ticker id " & tickerid & " not currently receiving market depth"
    End If
    mTickerTable(tickerid).receivingMarketDepth = False
End With

addData CANCEL_MKT_DEPTH
addData Version
addData MARKET_DEPTH_TICKER_BASE + tickerid
send

End Sub

Public Sub cancelNewsBulletins()

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
            NoValidID, _
            ApiErrorCodes.NotConnected, _
            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 1

addData CANCEL_NEWS_BULLETINS
addData Version
send
End Sub

Public Sub cancelOrder(ByVal orderId As Long)
Dim lOrder As Order

If mConnectionState <> ConnectionStates.ConnConnected And _
    Not mSimulateOrders _
Then
    RaiseEvent errorMessage(gCurrentTime, _
                        orderId, _
                        ApiErrorCodes.NotConnected, _
                        LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Set lOrder = gAllOrders(CStr(orderId))
If mSimulateOrders Or lOrder.simulated Then
    mOrderSimulator.cancelOrder orderId
    Exit Sub
End If

Const Version = 1
addData CANCEL_ORDER
addData Version
addData orderId
send

End Sub

Public Sub Connect(ByVal server As String, ByVal port As Long, ByVal clientID As Long)
On Error GoTo errNoSocket

mDisconnectedByApplication = False
createsocket

If mConnectionState <> ConnectionStates.ConnNotConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.AlreadyConnected, _
                            LoadResString(ApiErrorCodes.AlreadyConnected))
    Exit Sub
End If

mClientID = clientID
mServer = IIf(server = "", "127.0.0.1", server)
mPort = port

On Error GoTo errNoConnect
mSocket.Connect mServer, mPort
Exit Sub

errNoSocket:
RaiseEvent errorMessage(gCurrentTime, _
                NoValidID, _
                ApiErrorCodes.FailCreateSocket, _
                LoadResString(ApiErrorCodes.FailCreateSocket))
Exit Sub

errNoConnect:
RaiseEvent errorMessage(gCurrentTime, _
                NoValidID, _
                ApiErrorCodes.ConnectFail, _
                LoadResString(ApiErrorCodes.ConnectFail))
End Sub

Public Sub disconnect()
mDisconnectedByApplication = True
If Not mConnectionTimer Is Nothing Then mConnectionTimer.StopTimer
mConnectionState = ConnectionStates.ConnNotConnected
releaseAllTickerTableEntries
ReleaseSocket
mIndexIn = 0
mIndexOut = 0
mInputIndex = 0
RaiseEvent connectionToTWSClosed(Now, False)
End Sub

Public Function GetExchangeCodes() As String()
GetExchangeCodes = mExchangeCodes
End Function


Public Function getServiceProviderError() As ServiceProviderError
Set getServiceProviderError = mServiceProviderError
Set mServiceProviderError = Nothing
End Function

Public Function IsValidExchangeCode(ByVal code As String) As Boolean
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

Public Function NewContracts() As contracts
Set NewContracts = New contracts
End Function

Public Function newContractSpecifier( _
                Optional ByVal localSymbol As String, _
                Optional ByVal symbol As String, _
                Optional ByVal exchange As String, _
                Optional ByVal secType As SecurityTypes = SecTypeNone, _
                Optional ByVal currencyCode As String, _
                Optional ByVal expiry As String, _
                Optional ByVal strike As Double, _
                Optional ByVal right As OptionRights = OptNone) As TradeBuild.contractSpecifier

Dim Source As String: Source = "TradeBuild.TradeBuildAPI::newContractSpecifier"

If localSymbol = "" And symbol = "" Then
    err.Raise ErrorCodes.IllegalArgumentException, _
                Source, _
                "Symbol must be supplied if localSymbol is not supplied"
End If

If exchange <> "" And _
    Not IsValidExchangeCode(exchange) _
Then
    err.Raise ErrorCodes.IllegalArgumentException, _
                Source, _
                "'" & exchange & "' is not a valid exchange code"
End If

Select Case secType
Case 0  ' ie not supplied
Case SecTypeStock
Case SecTypeFuture
    If expiry <> "" Then
' need to write a function to validate the expiry format
'        err.Raise ErrorCodes.IllegalArgumentException, _
'                    source, _
'                    "expiry must be supplied"
    End If
Case SecTypeOption, SecTypeFuturesOption
    If expiry <> "" Then
' need to write a function to validate the expiry format
'        err.Raise ErrorCodes.IllegalArgumentException, _
'                    source, _
'                    "expiry must be supplied"
    End If
    If strike < 0 Then
        err.Raise ErrorCodes.IllegalArgumentException, _
                    Source, _
                    "strike must be > 0"
    End If
    Select Case right
    Case OptCall
    Case OptPut
    Case OptNone
    Case Else
        err.Raise ErrorCodes.IllegalArgumentException, _
                    Source, _
                    "'" & right & "' is not a valid option right"
    End Select
Case SecTypeCash
Case SecTypeBag
Case SecTypeIndex
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
                Source, _
                "'" & secType & "' is not a valid secType"
End Select

Set newContractSpecifier = New TradeBuild.contractSpecifier
With newContractSpecifier
    .localSymbol = localSymbol
    .symbol = symbol
    .exchange = exchange
    .secType = secType
    .currencyCode = currencyCode
    .expiry = expiry
    .strike = strike
    .right = right
End With
End Function

Public Function NewTicker() As Ticker
Static keySeq As Long
keySeq = keySeq + 1
Set NewTicker = mTickers.add("$" & keySeq)
End Function

Public Sub placeOrder(ByVal pOrder As Order)
Dim specifier As TradeBuild.contractSpecifier

If mConnectionState <> ConnectionStates.ConnConnected And _
   Not mSimulateOrders _
Then
    RaiseEvent errorMessage(gCurrentTime, _
                            pOrder.id, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

pOrder.clientID = mClientID

If mSimulateOrders Or pOrder.simulated Then
    mOrderSimulator.placeOrder pOrder
    Exit Sub
End If

Set specifier = pOrder.Contract.specifier

Const Version = 11
addData PLACE_ORDER
addData Version
addData pOrder.id

' send contract fields
With specifier
    addData .symbol
    addData secTypeToString(.secType)
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
If mServerVersion >= 8 And specifier.secType = SecurityTypes.SecTypeBag Then
    With specifier
        If .ComboLegs.Count = 0 Then
            addData 0
        Else
            addData .ComboLegs.Count
            Dim ComboLeg As ComboLeg
            For Each ComboLeg In .ComboLegs
                With ComboLeg
                    addData .contractID
                    addData .ratio
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
                    
Public Sub RemoveListener(ByVal key As String)
mListeners.Remove key
End Sub

Public Sub ReplaceFA(ByVal DataType As FADataTypes, ByVal xml As String)


If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 1

If mServerVersion < 13 Then
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.UpdateTWS, _
                    LoadResString(ApiErrorCodes.UpdateTWS))
End If

addData REPLACE_FA
addData Version
addData DataType
addData xml

send
End Sub

Public Sub requestAccountUpdates(ByVal subscribe As Boolean, ByVal acctCode As String)

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 2

addData REQ_ACCT_DATA
addData Version
addBoolean subscribe

If mServerVersion >= 9 Then addData acctCode

send
End Sub

Public Sub requestAllOpenOrders()

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If
                            

Const Version = 1

addData REQ_ALL_OPEN_ORDERS
addData Version
send
End Sub

Public Sub requestAutoOpenOrders(ByVal autoBind As Boolean)

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 1

addData REQ_AUTO_OPEN_ORDERS
addData Version
addBoolean autoBind
send
End Sub

Public Function RequestContract(ByVal pContractSpecifier As contractSpecifier) As contracts

On Error GoTo err

'If mConnectionState <> ConnectionStates.ConnConnected Then
'    RaiseEvent errorMessage(gCurrentTime, _
'                            NoValidID, _
'                            ApiErrorCodes.NotConnected, _
'                            LoadResString(ApiErrorCodes.NotConnected))
'    Exit Function
'End If

pContractSpecifier.locked = True

Set RequestContract = gServiceProviders.RequestContract(pContractSpecifier)
Exit Function

' the following code will be used when the TWS Contract Info Service Provider
' is developed
Set RequestContract = New contracts
RequestContract.contractSpecifier = pContractSpecifier

reqContract -1, RequestContract, pContractSpecifier, False

Exit Function

err:
err.Raise err.Number, _
            "TradeBuild.TradeBuildAPI::RequestContract"
End Function

Friend Function RequestContractEx( _
                            ByVal pContracts As contracts, _
                            ByVal pContractSpecifier As contractSpecifier) As contracts

On Error GoTo err

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Function
End If

pContractSpecifier.locked = True
reqContract -1, pContracts, pContractSpecifier, False

Exit Function
err:
err.Raise err.Number, _
            "TradeBuild.TradeBuildAPI::RequestContractEx"
End Function

Public Sub RequestExecutions(ByVal filter As ExecutionFilter)
Dim theFilter As ExecutionFilter

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 2

addData REQ_EXECUTIONS
addData Version

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
        addData secTypeToString(.secType)
        addData .exchange
        addData orderActionToString(.action)
    End With
End If

send
End Sub

Public Sub requestFA(ByVal DataType As FADataTypes)

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 1

If mServerVersion < 13 Then
    RaiseEvent errorMessage(gCurrentTime, _
                    NoValidID, _
                    ApiErrorCodes.UpdateTWS, _
                    LoadResString(ApiErrorCodes.UpdateTWS))
End If

addData REQ_FA
addData Version
addData DataType

send
End Sub

Public Sub requestManagedAccounts()

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 1

addData REQ_MANAGED_ACCTS
addData Version
send
End Sub

Friend Function RequestMarketDataEx(ByVal pContractSpecifier As contractSpecifier, _
                            ByVal dataConsumer As IStreamingDataConsumer, _
                            Optional ByVal includeMarketDepth As Boolean = False) As Long

Dim lTickerID As Long

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Function
End If

pContractSpecifier.locked = True

lTickerID = AllocateTickerTableEntry(dataConsumer)
RequestMarketDataEx = lTickerID
Set mTickerTable(lTickerID).contractSpecifier = pContractSpecifier
reqMarketData lTickerID

If includeMarketDepth Then reqMarketDepth lTickerID, pContractSpecifier

End Function

Friend Sub RequestMarketDepth(ByVal pTicker As Ticker)
                            
If mConnectionState <> ConnectionStates.ConnConnected And _
    mConnectionState <> ConnectionStates.ConnReplaying _
Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

With mTickerTable(pTicker.id)
    If pTicker.Contract Is Nothing Then err.Raise ErrorCodes.InvalidTickerID, "TradeBuild.TradeBuildAPI::requestMarketDepth", "Invalid ticker: key=" & pTicker.key
    .receivingMarketDepth = True
End With

reqMarketDepth pTicker.id, pTicker.Contract.specifier

End Sub

Public Sub RequestNewsBulletins(allMsgs As Boolean)

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 1

addData REQ_NEWS_BULLETINS
addData Version
addData allMsgs
send
End Sub

Public Sub RequestOpenOrders()

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 1

addData REQ_OPEN_ORDERS
addData Version
send
End Sub

Public Sub SetServerLogLevel(logLevel As Long)

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 1

addData SET_SERVER_LOGLEVEL
addData Version
addData logLevel
send
End Sub

Friend Sub SetServiceProviderError(errorCode, _
                    errorMessage, _
                    serviceProviderName)
Set mServiceProviderError = New ServiceProviderError
mServiceProviderError.errorCode = errorCode
mServiceProviderError.message = errorMessage
mServiceProviderError.serviceProviderName = serviceProviderName
mServiceProviderError.timestamp = gCurrentTime
RaiseEvent errorMessage(gCurrentTime, _
                        NoValidID, _
                        ApiErrorCodes.ServiceProviderErrorNotification, _
                        "A Service Provider error has occurred")
                    
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

Private Sub checkmessages()
Static alreadyInvoked As Boolean
Dim msgID As TWSSocketInMsgTypes
Dim Version As Long
Dim serverVersion As Long
Dim expectTimestamp As Boolean
Dim timeStampStr As String
Dim prevTimeStampStr As String

If alreadyInvoked Then Exit Sub
alreadyInvoked = True

On Error GoTo err

If mConnectionState <> ConnectionStates.ConnConnected Then
    ' first data received is plain server version number
    prevTimeStampStr = gFormatTimestamp(gCurrentTime, TimestampDateAndTime)
    checkpointBuffer
    serverVersion = getLong
    mServerVersion = serverVersion And &H3FF&
    expectTimestamp = serverVersion And &H400&
    If mServerVersion < SERVER_VERSION Then
        RaiseEvent errorMessage(gCurrentTime, _
                        NoValidID, _
                        ApiErrorCodes.UpdateTWS, _
                        LoadResString(ApiErrorCodes.UpdateTWS))
        disconnect
        Exit Sub
    End If
    mConnectionState = ConnectionStates.ConnConnected
    Set gAllOrders = New Collection
    
    ' now restart any tickers that were previously in operation
    reEstablishMarketData True
    
    ' and re-request any market depth
    resetMarketDepth gCurrentTime, True
    
    RaiseEvent connected(gCurrentTime)
End If

Do While True
    checkpointBuffer
    If expectTimestamp Then
        ' expected timestamp format is 'yyyymmddhhmmss.f'. The fractional part .f
        ' represents a fraction of 1 second, to any desired resolution. The
        ' timestamp may be truncated from the beginning up to but not
        ' including the decimal point, in which case the missing characters
        ' are substituted from the last received timestamp
        timeStampStr = getString
        timeStampStr = Left$(prevTimeStampStr, 15 - InStr(1, timeStampStr, ".")) * timeStampStr
        prevTimeStampStr = timeStampStr
'        gSetTime CDbl(DateSerial(Left$(timeStampStr, 4), _
'                                Mid$(timeStampStr, 5, 2), _
'                                Mid$(timeStampStr, 7, 2))) + _
'                CDbl(TimeSerial(Mid$(timeStampStr, 9, 2), _
'                                Mid$(timeStampStr, 11, 2), _
'                                Mid$(timeStampStr, 13, 2))) + _
'                CDbl(right$(timeStampStr, Len(timeStampStr) - InStr(1, timeStampStr, ".") + 1) & "0") / 86400#
                                
    End If
    msgID = getLong
    Version = getLong
    
    Select Case msgID
    Case TICK_PRICE
        getTickPrice Version
    Case TICK_SIZE
        getTickSize Version
    Case ORDER_STATUS
        getOrderStatus Version
    Case ERR_MSG
        getErrorMsg Version
    Case OPEN_ORDER
        getOpenOrder Version
    Case ACCT_VALUE
        getAccountValue Version
    Case PORTFOLIO_VALUE
        getPortfolioValue Version
    Case ACCT_UPDATE_TIME
        getAccountTime Version
    Case NEXT_VALID_ID
        getNextValidId Version
    Case CONTRACT_DATA
        getContractData Version
    Case EXECUTION_DATA
        getExecutionData Version
    Case MARKET_DEPTH
        getMarketDepth Version
    Case MARKET_DEPTH_L2
        getMarketDepthL2 Version
    Case NEWS_BULLETINS
        getNewsBulletins Version
    Case MANAGED_ACCTS
        getManagedAccounts Version
    Case RECEIVE_FA
        getFAData Version
    Case Else
        Dim s As String
        Dim i As Long
        Dim j As Long
        s = "Invalid message code: " & msgID & vbCrLf & _
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
        RaiseEvent errorMessage(gCurrentTime, _
                        NoValidID, _
                        ApiErrorCodes.InvalidMessageCode, _
                        s)
        disconnect
    End Select
Loop

alreadyInvoked = False

Exit Sub

err:
If err.Number = DataIncomplete Then
    rollbackBuffer
    alreadyInvoked = False
    Exit Sub
End If
alreadyInvoked = False
err.Raise err.Number
End Sub

Private Sub checkpointBuffer()
mIndexInChk = mIndexIn
mInputIndexChk = mInputIndex
End Sub

Private Sub createsocket()
If mfWinsock Is Nothing Then
    Set mfWinsock = New fWinsock
    Set mSocket = mfWinsock.Socket
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
err.Raise DataIncomplete
End Function

Private Sub getAccountTime(ByVal Version As Long)
Dim listendata As TWSSocketMsgAcctUpdateTime
Dim timestamp As Date: timestamp = gCurrentTime
Dim accountTime As String: accountTime = getString

RaiseEvent updateAccountTime(timestamp, accountTime)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.ACCT_UPDATE_TIME
    listendata.accountTime = accountTime
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgAcctUpdateTime, _
                    timestamp
End If

End Sub

Private Sub getAccountValue(ByVal Version As Long)
Dim listendata As TWSSocketMsgAcctValue
Dim timestamp As Date: timestamp = gCurrentTime
Dim key As String: key = getString
Dim val As String: val = getString
Dim cur As String: cur = getString
Dim accountName As String: If Version >= 2 Then accountName = getString

RaiseEvent updateAccountValue(timestamp, key, val, cur, accountName)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.ACCT_VALUE
    listendata.accountName = accountName
    listendata.currencyCode = cur
    listendata.key = key
    listendata.value = val
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgAcctValue, _
                    timestamp
End If

End Sub

Private Sub getContractData(ByVal Version As Long)
Dim listendata As TWSSocketMsgContractData
Dim timestamp As Date
Dim lContract As Contract
Dim minimumTick As Double
Dim multiplier As Long
Dim OrderTypes As String
Dim orderTypesStr() As String
Dim lOrderTypes() As TradeBuild.OrderTypes
Dim validExchanges As String
Dim lValidExchanges() As String
Dim i As Long
Dim j As Long

timestamp = gCurrentTime

Set lContract = New Contract
With lContract
    .specifier = New contractSpecifier
    .specifier.locked = True
    .specifier.symbol = getString
    .specifier.secType = secTypeFromString(getString)
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
    
    ReDim lOrderTypes(UBound(orderTypesStr)) As TradeBuild.OrderTypes
    
    For i = 0 To UBound(orderTypesStr)
        lOrderTypes(j) = orderTypeFromString(orderTypesStr(i))
        
        If Not lOrderTypes(j) = TradeBuild.OrderTypes.OrderTypeNone Then
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
                    SetServiceProviderError ErrorCodes.UnknownOrderTypeFromTWS, _
                                            "Unknown order type " & orderTypesStr(i) & " received from TWS", _
                                            "IB TWS"
                End If
            End Select
        End If
    Next
    
    If j <> (UBound(orderTypesStr) - 1) Then
        ReDim Preserve lOrderTypes(j - 1) As TradeBuild.OrderTypes
    End If
    
    .OrderTypes = lOrderTypes
    validExchanges = getString
    lValidExchanges = Split(validExchanges, ",")
    .validExchanges = lValidExchanges
End With

processContract lContract

If mEnableListeners Then
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
        listendata.secType = .specifier.secType
        listendata.strike = .specifier.strike
        listendata.symbol = .specifier.symbol
        listendata.tradingClass = .tradingClass
        listendata.validExchanges = validExchanges
        listendata.Version = Version
    End With
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgContractData, _
                    timestamp
End If

End Sub

Private Sub getErrorMsg(ByVal Version As Long)
Dim listendata As TWSSocketMsgErrMsg
Dim timestamp As Date
Dim id As Long
Dim errorCode As Long
Dim errorMsg As String
Dim execFilter As ExecutionFilter
id = getLong
errorCode = getLong
errorMsg = getString
timestamp = gCurrentTime


Select Case errorCode
Case 200    ' security definition not known
    If id >= 0 Then
        ' this must relate to a reqMktData - this should never occur as
        ' we make sure the contract is valid before starting the ticker
        With mTickerTable(id Mod INCARNATION_INCREMENT)
            RaiseEvent ContractInvalid(timestamp, _
                                        .contractSpecifier)
            
            ' need to do something about the following two lines!!!!!!!!!!!!!!!!!!
            '.tickerObj.ContractInvalid
            'mTickers.ContractInvalid .tickerObj, .ContractSpecifier
            releaseTickerTableEntry (id Mod INCARNATION_INCREMENT)
        End With
    Else
        ' this must relate to a reqContract
        processContractError
    End If
Case 202
    RaiseEvent orderStatus(timestamp, id, OrderStatuses.OrderStatusCancelled, 0, 0, 0, 0, 0, 0, mClientID)
Case 309
    ' max market depth requests exceeded
    
    ' need to do something about the following !!!!!!!!!!!!!!!!!!!!!!!!!
    'mTickerTable(id - MARKET_DEPTH_TICKER_BASE).tickerObj.MarketDepthNotAvailable errorMsg
    'mTickers.MarketDepthNotAvailable gCurrentTime, id - MARKET_DEPTH_TICKER_BASE, errorMsg
Case 316
    resetMarketDepth timestamp, reEstablish:=True
Case 317
    resetMarketDepth timestamp, reEstablish:=False
Case 1100
    ' connectivity between TWS and IB has been lost
    RaiseEvent connectionToIBClosed(timestamp)
Case 1101
    ' connectivity between TWS and IB has been restored, but data has been
    ' lost, so need to reestablish market data requests
    
    reEstablishMarketData False
    
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

    RaiseEvent connectionToIBRecovered(timestamp)
    
Case Else
    SetServiceProviderError errorCode, errorMsg, "IB TWS"
End Select

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.ERR_MSG
    listendata.errorCode = errorCode
    listendata.errorMsg = errorMsg
    listendata.id = id
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgErrMsg, _
                    timestamp
End If

End Sub

Private Sub getExecutionData(ByVal Version As Long)
Dim listendata As TWSSocketMsgExecutionData
Dim timestamp As Date: timestamp = gCurrentTime
Dim secType As String
Dim right As String
Dim side As String
Dim price As Double

Dim id As Long: id = getLong

Dim lContractSpecifier As contractSpecifier
Set lContractSpecifier = New contractSpecifier
With lContractSpecifier
    .locked = True
    .symbol = getString
    secType = getString
    .secType = secTypeFromString(secType)
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

If Version >= 2 Then exec.permId = getLong

If Version >= 3 Then exec.clientID = getLong
   
RaiseEvent ExecutionDetails(timestamp, id, lContractSpecifier, exec)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.EXECUTION_DATA
    With lContractSpecifier
        listendata.currencyCode = .currencyCode
        listendata.exchange = .exchange
        listendata.expiry = .expiry
        listendata.localSymbol = .localSymbol
        listendata.right = right
        listendata.secType = secType
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
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgExecutionData, _
                    timestamp
End If

End Sub

Private Sub getFAData(ByVal Version As Long)
Dim listendata As TWSSocketMsgReceiveFa
Dim timestamp As Date: timestamp = gCurrentTime
Dim DataType As FADataTypes: DataType = getLong
Dim XMLData As String: XMLData = getString

RaiseEvent FAData(timestamp, DataType, XMLData)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.MANAGED_ACCTS
    listendata.DataType = DataType
    listendata.XMLData = XMLData
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgManagedAccts, _
                    timestamp
End If

End Sub

Private Sub getManagedAccounts(ByVal Version As Long)
Dim listendata As TWSSocketMsgManagedAccts
Dim timestamp As Date: timestamp = gCurrentTime
Dim accountsList As String: accountsList = getString

RaiseEvent managedAccounts(timestamp, accountsList)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.MANAGED_ACCTS
    listendata.accountsList = accountsList
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgManagedAccts, _
                    timestamp
End If

End Sub

Private Sub getMarketDepth(ByVal Version As Long)
Dim listendata As TWSSocketMsgMarketDepth
Dim timestamp As Date
Dim id As Long: id = getLong
Dim position As Long: position = getLong
Dim operation As Long: operation = getLong
Dim side As Long: side = getLong
Dim price As Double: price = getDouble
Dim Size As Long: Size = getLong

timestamp = gCurrentTime

processMarketDepth timestamp, _
                id - MARKET_DEPTH_TICKER_BASE, _
                position, _
                operation, _
                side, _
                price, _
                Size

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.MARKET_DEPTH
    listendata.operation = operation
    listendata.position = position
    listendata.price = price
    listendata.side = side
    listendata.Size = Size
    listendata.tickerid = id
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgMarketDepth, _
                    timestamp
End If

End Sub

Private Sub getMarketDepthL2(ByVal Version As Long)
Dim listendata As TWSSocketMsgMarketDepthL2
Dim timestamp As Date: timestamp = gCurrentTime
Dim id As Long: id = getLong
Dim position As Long: position = getLong
Dim marketMaker As String: marketMaker = getString
Dim operation As Long: operation = getLong
Dim side As Long: side = getLong
Dim price As Double: price = getDouble
Dim Size As Long: Size = getLong

With mTickerTable(id - MARKET_DEPTH_TICKER_BASE)
    If .inUse Then
        
        .dataConsumer.UpdateMktDepthL2 timestamp, _
                        position, _
                        marketMaker, _
                        operation, _
                        side, _
                        price, _
                        Size
                        
    End If
End With

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.MARKET_DEPTH_L2
    listendata.marketMaker = marketMaker
    listendata.operation = operation
    listendata.position = position
    listendata.price = price
    listendata.side = side
    listendata.Size = Size
    listendata.tickerid = id
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgMarketDepthL2, _
                    timestamp
End If

End Sub

Private Sub getNewsBulletins(ByVal Version As Long)
Dim listendata As TWSSocketMsgNewsBulletins
Dim timestamp As Date: timestamp = gCurrentTime
Dim msgID As Long: msgID = getLong
Dim msgType As Long: msgType = getLong
Dim newsMessage As String: newsMessage = getString
Dim originatingExch As String: originatingExch = getString

RaiseEvent updateNewsBulletin(timestamp, msgID, msgType, newsMessage, originatingExch)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.NEWS_BULLETINS
    listendata.newsID = msgID
    listendata.newsMessage = newsMessage
    listendata.newsType = msgType
    listendata.originatingExch = originatingExch
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgNewsBulletins, _
                    timestamp
End If

End Sub

Private Sub getNextValidId(ByVal Version As Long)
Dim listendata As TWSSocketMsgNextValidId
Dim timestamp As Date: timestamp = gCurrentTime

gNextOrderID = getLong

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.NEXT_VALID_ID
    listendata.id = gNextOrderID
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgNextValidId, _
                    timestamp
End If

End Sub

Private Sub getOpenOrder(ByVal Version As Long)
Dim listendata As TWSSocketMsgOpenOrder
Dim timestamp As Date: timestamp = gCurrentTime
Dim lOrder As Order
Dim id As Long
Dim secType As String
Dim right As String
Dim action As String
Dim orderType As String
Dim timeInForce As String

id = getLong

On Error Resume Next
Set lOrder = gAllOrders(CStr(id))
On Error GoTo 0

If lOrder Is Nothing Then
    ' this must be an open order from the previous API session, but
    ' for which we haven't yet had the openOrder notification
    Set lOrder = New Order
    lOrder.id = id
    gAllOrders.add lOrder, CStr(id)
End If

' read contract fields
Dim lContractSpecifier As contractSpecifier
Set lContractSpecifier = New contractSpecifier
With lContractSpecifier
    .locked = True
    .symbol = getString
    secType = getString
    .secType = secTypeFromString(secType)
    .expiry = getString
    .strike = getDouble
    right = getString
    .right = optionRightFromString(right)
    .exchange = getString
    .currencyCode = getString
End With

If Version >= 2 Then lContractSpecifier.localSymbol = getString

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
    If Version >= 3 Then .clientID = getLong
    
    If Version >= 4 Then
        .permId = getLong
        .ignoreRTH = getBoolean
        .hidden = getBoolean
        .discretionaryAmt = getDouble
    End If

    If Version >= 5 Then .goodAfterTime = getString
    
    If Version >= 6 Then .sharesAllocation = getString
    
    If Version >= 7 Then
        .FAGroup = getString
        .FAMethod = getString
        .FAPercentage = getString
        .FAProfile = getString
    End If
    
End With

RaiseEvent openOrder(timestamp, lContractSpecifier, lOrder.clone)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.OPEN_ORDER
    With lContractSpecifier
        listendata.currencyCode = .currencyCode
        listendata.exchange = .exchange
        listendata.expiry = .expiry
        listendata.localSymbol = .localSymbol
        listendata.right = .right
        listendata.secType = .secType
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
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgOpenOrder, _
                    timestamp
End If

End Sub

Private Sub getOrderStatus(ByVal Version As Long)
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

If Version >= 2 Then permId = getLong

If Version >= 3 Then parentId = getLong

If Version >= 4 Then lastFillPrice = getDouble

If Version >= 5 Then clientID = getLong

On Error Resume Next
Set lOrder = gAllOrders(CStr(id))
On Error GoTo 0
If lOrder Is Nothing Then
    ' this must be an open order from the previous API session, but
    ' for which we haven't yet had the openOrder notification
    Set lOrder = New Order
    lOrder.id = id
    gAllOrders.add lOrder, CStr(id)
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
    RaiseEvent orderStatus(timestamp, _
                            id, _
                            orderStatusFromString(status), _
                            filled, _
                            remaining, _
                            avgFillPrice, _
                            permId, _
                            parentId, _
                            lastFillPrice, _
                            clientID)
End If

If mEnableListeners Then
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
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgOrderStatus, _
                    timestamp
End If

End Sub

Private Sub getPortfolioValue(ByVal Version As Long)
Dim listendata As TWSSocketMsgPortfolioValue
Dim timestamp As Date: timestamp = gCurrentTime
Dim secType As String
Dim right As String
Dim lContractSpecifier As contractSpecifier

Set lContractSpecifier = New contractSpecifier
With lContractSpecifier
    .locked = True
    .symbol = getString
    secType = getString
    .secType = secTypeFromString(secType)
    .expiry = getString
    .strike = getDouble
    right = getString
    .right = optionRightFromString(right)
    .currencyCode = getString
    If Version >= 2 Then .localSymbol = getString
End With

Dim position As Long: position = getLong
Dim marketPrice As Double: marketPrice = getDouble
Dim marketValue As Double: marketValue = getDouble
Dim averageCost As Double: averageCost = 0#
Dim unrealizedPNL As Double: unrealizedPNL = 0#
Dim realizedPNL As Double: realizedPNL = 0#
If Version >= 3 Then
    averageCost = getDouble
    unrealizedPNL = getDouble
    realizedPNL = getDouble
End If
Dim accountName As String: If Version >= 4 Then accountName = getString
                         
RaiseEvent updatePortfolioValue(timestamp, lContractSpecifier, position, marketPrice, marketValue, averageCost, unrealizedPNL, realizedPNL, accountName)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.PORTFOLIO_VALUE
    With lContractSpecifier
        listendata.currencyCode = .currencyCode
        listendata.expiry = .expiry
        listendata.localSymbol = .localSymbol
        listendata.right = right
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
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgPortfolioValue, _
                    timestamp
End If

End Sub

Private Sub getTickPrice(ByVal Version As Long)
Dim listendata As TWSSocketMsgTickPrice
Dim timestamp As Date
Dim price As Double
Dim Size As Long
Dim TickType As Long
Dim tickerid As Long
tickerid = getLong
TickType = getLong
price = getDouble
If Version >= 2 Then Size = getLong
timestamp = gCurrentTime

processTickPrice timestamp, _
                tickerid Mod INCARNATION_INCREMENT, _
                TickType, _
                price, _
                Size

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.TICK_PRICE
    listendata.price = price
    listendata.Size = Size
    listendata.tickerid = tickerid
    listendata.TickType = TickType
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgTickPrice, _
                    timestamp
End If

End Sub

Private Sub getTickSize(ByVal Version As Long)
Dim listendata As TWSSocketMsgTickSize
Dim timestamp As Date
Dim Size As Long
Dim TickType As Long
Dim tickerid As Long
tickerid = getLong
TickType = getLong
Size = getLong
timestamp = gCurrentTime

processTickSize timestamp, _
                tickerid Mod INCARNATION_INCREMENT, _
                TickType, _
                Size

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.TICK_PRICE
    listendata.Size = Size
    listendata.tickerid = tickerid
    listendata.TickType = TickType
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgTickSize, _
                    timestamp
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

Private Function orderTypeFromString(ByVal value As String) As TradeBuild.OrderTypes
Select Case UCase$(value)
Case "MKT"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeMarket
Case "MKTCLS"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeMarketClose
Case "LMT"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeLimit
Case "LMTCLS"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeLimitClose
Case "PEGMKT"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypePegMarket
Case "STP"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeStop
Case "STPLMT"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeStopLimit
Case "TRAIL"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeTrail
Case "REL"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeRelative
Case "VWAP"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeVWAP
Case "MTL"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeMarketToLimit
Case "QUOTE"
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeQuote
Case Else
    orderTypeFromString = TradeBuild.OrderTypes.OrderTypeNone
End Select
End Function

Private Function orderTypeToString(ByVal value As OrderTypes) As String
Select Case value
Case TradeBuild.OrderTypes.OrderTypeMarket
    orderTypeToString = "MKT"
Case TradeBuild.OrderTypes.OrderTypeMarketClose
    orderTypeToString = "MKTCLS"
Case TradeBuild.OrderTypes.OrderTypeLimit
    orderTypeToString = "LMT"
Case TradeBuild.OrderTypes.OrderTypeLimitClose
    orderTypeToString = "LMTCLS"
Case TradeBuild.OrderTypes.OrderTypePegMarket
    orderTypeToString = "PEGMKT"
Case TradeBuild.OrderTypes.OrderTypeStop
    orderTypeToString = "STP"
Case TradeBuild.OrderTypes.OrderTypeStopLimit
    orderTypeToString = "STPLMT"
Case TradeBuild.OrderTypes.OrderTypeTrail
    orderTypeToString = "TRAIL"
Case TradeBuild.OrderTypes.OrderTypeRelative
    orderTypeToString = "REL"
Case TradeBuild.OrderTypes.OrderTypeVWAP
    orderTypeToString = "VWAP"
Case TradeBuild.OrderTypes.OrderTypeMarketToLimit
    orderTypeToString = "MTL"
Case TradeBuild.OrderTypes.OrderTypeQuote
    orderTypeToString = "QUOTE"
Case Else
    err.Raise ErrorCodes.InvalidOrderType
End Select
End Function

Private Sub ProcessAsk(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal price As Double, _
                            ByVal Size As Long)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    .dataConsumer.ask timestamp, _
                price, _
                Size
End With
End Sub

Private Sub ProcessBid(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal price As Double, _
                            ByVal Size As Long)
With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    .dataConsumer.bid timestamp, _
                price, _
                Size
End With
End Sub

Private Sub ProcessClose(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal price As Double)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    .dataConsumer.previousClose timestamp, price
End With
End Sub

Private Sub processContract(ByVal pContract As Contract)

If Not mFirstContractDetailsReceived Then
    mFirstContractDetailsReceived = True
    
    ' now send a reqContract that will give a 'no such security
    ' exists' error. This error will act as a delimiter for all
    ' the contract details returned by the orginal request. Note that
    ' tests indicate that contract details for different reqContracts
    ' are not interleaved
    
    Const Version = 1
    addData REQ_CONTRACT_DATA
    addData Version
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

With mContractDetailsRequest
    .contracts.AddContract pContract
    If Not .InternalRequest Then
        RaiseEvent Contract(gCurrentTime, pContract)
    End If
End With

End Sub

Private Sub processContractError()
Dim tickerid As Long
Dim lContract As Contract

With mContractDetailsRequest
    If .contracts.Count > 0 Then
        ' we are processing the forced invalid contract details request
        ' for a request for which we have received one or more contract details
        If .InternalRequest Then
            ' we no longer do the following - should we????????????????????????????
'            TickerID = .TickerID
'            If .Contracts.Count = 1 Then
'                Set lContract = .Contracts(1)
'                With mTickerTable(TickerID)
'                    Set .ContractSpecifier = lContract.specifier
'                    If Not .tickerObj Is Nothing Then
'                        .tickerObj.Contract = lContract
'                    End If
'                    If .writeMktDataToTickfile Or _
'                        .writeMktDepthToTickfile _
'                    Then
'                        .TickfileWriter.Contract = lContract
'                    End If
'                    If findTickerTableIndex(.ContractSpecifier, TickerID) <> -1 Then
'                        ' a ticker for this contract is already running
'                        .tickerObj.Contract = Nothing
'                        .tickerObj.id = -1
'                        .tickerObj.DuplicateTickerRequest .ContractSpecifier
'                        mTickers.DuplicateTickerRequest .tickerObj, .ContractSpecifier
'                        RaiseEvent errorMessage(gCurrentTime, _
'                                                TickerID, _
'                                                ApiErrorCodes.TickerAlreadyRunningForThisContract, _
'                                                "A ticker has already been started for the contract specified in this ticker")
'                        releaseTickerTableEntry TickerID
'                    Else
'                        reqMarketData TickerID
'                    End If
'                End With
'            Else
'                ' the requested contract is not unique
'                With mTickerTable(TickerID)
'                    .tickerObj.ContractInvalid
'                    mTickers.ContractInvalid .tickerObj, .ContractSpecifier
'                End With
'                releaseTickerTableEntry TickerID
'            End If
        Else
            .contracts.LoadComplete
        End If
    Else
        ' this is a genuine error response to a contract details request
        RaiseEvent ContractInvalid(gCurrentTime, .contractSpecifier)
        If .InternalRequest Then
'            TickerID = .TickerID
'            With mTickerTable(TickerID)
'                .tickerObj.ContractInvalid
'                mTickers.ContractInvalid .tickerObj, .ContractSpecifier
'            End With
'            releaseTickerTableEntry TickerID
        Else
            .contracts.ContractSpecifierInvalid
            .contracts.LoadComplete
        End If
    End If
End With
        
Set mContractDetailsRequest = Nothing
mFirstContractDetailsReceived = False

End Sub

Private Sub ProcessHigh(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal price As Double)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    .dataConsumer.high timestamp, price
    
End With
End Sub

Private Sub ProcessLast(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal price As Double, _
                            ByVal Size As Long)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    .dataConsumer.trade timestamp, _
            price, _
            Size
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
                                ByVal Size As Long)

With mTickerTable(id)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    .dataConsumer.UpdateMktDepth timestamp, _
                position, _
                operation, _
                side, _
                price, _
                Size
    
End With
End Sub

Private Sub processTickPrice(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal TickType As TWSSocketTickTypes, _
                            ByVal price As Double, _
                            ByVal Size As Long)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    If price < (0.5 * .prevTrade) Then
        Exit Sub        ' ignore invalid prices
    End If
    
    Select Case TickType
    Case TICK_BID
        Size = IIf(Size <> 0, Size, .prevBidSize)
        ProcessBid timestamp, _
                    tickerid, _
                    price, _
                    Size
        .prevBid = price
        If Size <> 0 Then .prevBidSize = Size
    
    Case TICK_ASK
        Size = IIf(Size <> 0, Size, .prevAskSize)
        ProcessAsk timestamp, _
                    tickerid, _
                    price, _
                    Size
        
        .prevAsk = price
        If Size <> 0 Then .prevAskSize = Size
    
    Case TICK_LAST
        Size = IIf(Size <> 0, Size, .prevTradeSize)
        ProcessLast timestamp, _
                    tickerid, _
                    price, _
                    Size
        .prevTrade = price
        If Size <> 0 Then .prevTradeSize = Size
    
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
                            ByVal TickType As TWSSocketTickTypes, _
                            ByVal Size As Long)

With mTickerTable(tickerid)
    
    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    Select Case TickType
    Case TICK_BID_SIZE
        If Size = .prevBidSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        .prevBidSize = Size
        ProcessBid timestamp, _
                    tickerid, _
                    .prevBid, _
                    Size
    Case TICK_ASK_SIZE
        If Size = .prevAskSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        ProcessAsk timestamp, _
                    tickerid, _
                    .prevAsk, _
                    Size
    Case TICK_LAST_SIZE
        If Size = .prevTradeSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        .prevTradeSize = Size
        ProcessLast timestamp, _
                    tickerid, _
                    .prevTrade, _
                    Size
    Case TICK_VOLUME
        processVolume timestamp, _
                    tickerid, _
                    Size
        
    End Select
End With

End Sub

Private Sub processVolume(ByVal timestamp As Date, _
                            ByVal tickerid As Long, _
                            ByVal Size As Long)

With mTickerTable(tickerid)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    If Size = 0 Then Exit Sub
    
    .dataConsumer.volume timestamp, _
                Size
End With
End Sub

Private Sub reEstablishMarketData(ByVal newConnection As Boolean)
Dim i As Long
For i = 0 To UBound(mTickerTable)
    With mTickerTable(i)
        If Not .contractSpecifier Is Nothing Then
            If Not newConnection Then CancelMarketData i
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

Private Sub ReleaseSocket()
If Not mSocket Is Nothing Then
    mSocket.Close
    Set mSocket = Nothing
    Unload mfWinsock
    Set mfWinsock = Nothing
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

Private Sub reqContract(ByVal tickerid As Long, _
                        ByVal contracts As contracts, _
                        ByVal contractSpecifier As contractSpecifier, _
                        ByVal internal As Boolean)

' don't send another request if there's already one outstanding
If Not mContractDetailsRequest Is Nothing Then
    err.Raise ContractDetailsReqNotAllowed, _
                "TradeBuildAPI::reqContract", _
                "An existing request has not yet completed"
End If

contractSpecifier.locked = True

Set mContractDetailsRequest = New ContractDtlsReq
mContractDetailsRequest.tickerid = tickerid
mContractDetailsRequest.contracts = contracts
mContractDetailsRequest.contractSpecifier = contractSpecifier
mContractDetailsRequest.InternalRequest = internal

If Not IsValidExchangeCode(contractSpecifier.exchange) Then
    processContractError
    Exit Sub
End If

Const Version = 1

addData REQ_CONTRACT_DATA
addData Version
With contractSpecifier
    addData .symbol
    addData secTypeToString(.secType)
    addData Left$(.expiry, 6)
    addData .strike
    addData optionRightToString(.right)
    addData .exchange
    addData .currencyCode
    addData .localSymbol
End With
send

End Sub

Private Sub reqMarketData(ByVal tickerTableIndex As Long)

Const Version = 3

addData REQ_MKT_DATA
addData Version
addData tickerTableIndex + mTickerTable(tickerTableIndex).incarnation
With mTickerTable(tickerTableIndex).contractSpecifier
    addData .symbol
    addData secTypeToString(.secType)
    addData Left$(.expiry, 6)
    addData .strike
    addData optionRightToString(.right)
    addData .exchange
    addData .currencyCode
    If mServerVersion >= 2 Then addData .localSymbol
    ' Add combo legs for BAG requests
    If mServerVersion >= 8 And .secType = SecurityTypes.SecTypeBag Then
        addData .ComboLegs.Count
        Dim ComboLeg As ComboLeg
        For Each ComboLeg In .ComboLegs
            With ComboLeg
                addData .contractID
                addData .ratio
                addData orderActionToString(.action)
                addData .exchange
            End With
        Next
    End If
    
    send
End With

End Sub

Private Sub reqMarketDepth(ByVal tickerid As Long, _
                        ByVal contractSpec As contractSpecifier)

Const Version = 1

addData REQ_MKT_DEPTH
addData Version
addData tickerid + MARKET_DEPTH_TICKER_BASE
With contractSpec
    addData .symbol
    addData secTypeToString(.secType)
    addData Left$(.expiry, 6)
    addData .strike
    addData optionRightToString(.right)
    addData .exchange
    addData .currencyCode
    addData .localSymbol
End With
send
End Sub

Private Sub resetMarketDepth(ByVal timestamp As Date, _
                            ByVal reEstablish As Boolean)
Dim i As Long
For i = 0 To UBound(mTickerTable)
    With mTickerTable(i)
        If .receivingMarketDepth Then
            If reEstablish Then reqMarketDepth i, .contractSpecifier
            .dataConsumer.resetMarketDepth timestamp, reEstablish
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
mIndexIn = mIndexInChk
mInputIndex = mInputIndexChk
End Sub

Private Sub send()
ReDim Preserve mBufferOut(mIndexOut - 1) As Byte
mSocket.SendData mBufferOut
mIndexOut = 0
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


