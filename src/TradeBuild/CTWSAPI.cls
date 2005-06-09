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

Implements TradeBuildSP.IStreamingDataConsumer

'================================================================================
' Events
'================================================================================

Event ask( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal Price As Double, _
                ByVal Size As Long)

Event bid( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal Price As Double, _
                ByVal Size As Long)

Event connected(ByVal Timestamp As Date)
    ' Raised when a successful connection to TWS has been achieved.

Event connectFailed(ByVal Timestamp As Date, _
                ByVal Description As String, _
                ByVal retrying As Boolean)
    ' Raised when connection to TWS has failed and the connection
    ' retry interval is zero.

Event connecting(ByVal Timestamp As Date)
    ' Raised when an attempt to connect to TWS is initiated.

Event connectionToTWSClosed(ByVal Timestamp As Date)
    ' Raised when the connection to TWS is disconnected.

Event connectionToIBClosed(ByVal Timestamp As Date)

Event connectionToIBRecovered(ByVal Timestamp As Date)

Event Contract( _
                ByVal Timestamp As Date, _
                ByVal pContract As Contract)
    ' Raised when details for a security are received from TWS.

Event ContractInvalid( _
                ByVal Timestamp As Date, _
                ByVal pContractSpecifier As contractSpecifier)

Event dataReceived(ByVal Timestamp As Date)
    ' Raised when a new block of data is received from the socket connecting to TWS.
    
Event DOMClearCell( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal side As DOMSides, _
                ByVal Price As Double)
                
Event DOMSetCell( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal side As DOMSides, _
                ByVal Price As Double, _
                ByVal Size As Long)
                
Event DOMReset( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal marketDataReRequested As Boolean)
                
Event errorMessage( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal errorCode As ApiErrorCodes, _
                ByVal errorMsg As String)

Event ExecutionDetails( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal pContractSpecifier As contractSpecifier, _
                ByVal exec As Execution)

Event FAData( _
                ByVal Timestamp As Date, _
                ByVal DataType As FADataTypes, _
                ByVal XMLData As String)
                
Event High( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal Price As Double)

Event Low( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal Price As Double)

Event managedAccounts( _
                ByVal Timestamp As Date, _
                ByVal accountsList As String)

Event openOrder( _
                ByVal Timestamp As Date, _
                ByVal pContractSpecifier As contractSpecifier, _
                ByVal pOrder As Order)

Event orderStatus( _
                ByVal Timestamp As Date, _
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
                ByVal Timestamp As Date, _
                ByVal Order As Order, _
                ByVal pContractSpecifier As contractSpecifier, _
                ByRef fillPrice As Double, _
                ByRef fillSize As Long)

Event PreviousClose( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal Price As Double)

Event trade( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal Price As Double, _
                ByVal Size As Long)

Event updateAccountTime(ByVal Timestamp As Date, _
                ByVal accountTime As String)

Event updateAccountValue( _
                ByVal Timestamp As Date, _
                ByVal key As String, _
                ByVal value As String, _
                ByVal currencyCode As String, _
                ByVal accountName As String)

Event UpdateMarketDepth( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal position As Long, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal Price As Double, _
                ByVal Size As Long)
                
Event UpdateMarketDepthL2( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal position As Long, _
                ByVal marketMaker As String, _
                ByVal operation As DOMOperations, _
                ByVal side As DOMSides, _
                ByVal Price As Double, _
                ByVal Size As Long)

Event updateNewsBulletin( _
                ByVal Timestamp As Date, _
                ByVal msgID As Integer, _
                ByVal msgType As BulletinTypes, _
                ByVal message As String, _
                ByVal origExchange As String)
                
Event updatePortfolioValue( _
                ByVal Timestamp As Date, _
                ByVal pContractSpecifier As contractSpecifier, _
                ByVal position As Long, _
                ByVal marketPrice As Double, _
                ByVal marketValue As Double, _
                ByVal averageCost As Double, _
                ByVal unrealizedPNL As Double, _
                ByVal realizedPNL As Double, _
                ByVal accountName As String)

Event Volume( _
                ByVal Timestamp As Date, _
                ByVal id As Long, _
                ByVal Size As Long)

'================================================================================
' Types
'================================================================================

Private Type PriceTableEntry
    Price                   As Double
    Size                    As Long
End Type

Private Type TickerTableEntry
    incarnation             As Long
    contractSpecifier       As contractSpecifier
    prevBid                 As Double
    prevBidSize             As Long
    prevAsk                 As Double
    prevAskSize             As Long
    prevTrade               As Double
    prevTradeSize           As Long
    accumulatedVolume       As Long
    volumeCorrectionIncrement   As Long
    inUse                   As Boolean
    noVolumeAdjustments     As Boolean
    noImpliedTrades         As Boolean
    firstVolumeReceived     As Boolean
    writeMktDataToTickfile  As Boolean
    receivingMarketDepth    As Boolean
    writeMktDepthToTickfile As Boolean
    DOMEventsRequired       As DOMEvents
    TickfileWriter          As TickfileWriter
    tickerObj               As ticker
    askPrices()             As PriceTableEntry
    maxAskPricesIndex       As Long
    bidPrices()             As PriceTableEntry
    maxBidPricesIndex       As Long
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

Private mNextOrderID As Long
Private mAllOrders As Collection

Private WithEvents mOrderSimulator As AdvancedOrderSimulator
Attribute mOrderSimulator.VB_VarHelpID = -1
Private mSimulateOrders As Boolean

Private mServiceProviderError As ServiceProviderError

Private mListeners As listeners
Private mEnableListeners As Boolean

Private mTickers As Tickers

Private mPendingContractDetailsReqIndex As Long
    ' when set, this references the contract details request for which
    ' contract details are currently being received
    
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
Set mAllOrders = New Collection
mConnectionState = ConnectionStates.ConnNotConnected
Set mOrderSimulator = New AdvancedOrderSimulator
Set mListeners = New listeners
Set gListeners = mListeners
Set mTickers = New Tickers
Set gTickers = mTickers

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
' TradeBuildSP.IStreamingDataConsumer Interface Members
'================================================================================

Private Sub IStreamingDataConsumer_Ask( _
                            ByVal TickerID As Long, _
                            ByVal Timestamp As Date, _
                            ByVal Price As Double, _
                            ByVal Size As Long)
ProcessAsk Timestamp, _
            TickerID, _
            Price, _
            Size
End Sub

Private Sub IStreamingDataConsumer_Bid( _
                            ByVal TickerID As Long, _
                            ByVal Timestamp As Date, _
                            ByVal Price As Double, _
                            ByVal Size As Long)
ProcessBid Timestamp, _
            TickerID, _
            Price, _
            Size
End Sub

Private Sub IStreamingDataConsumer_High( _
                            ByVal TickerID As Long, _
                            ByVal Timestamp As Date, _
                            ByVal Price As Double)
ProcessHigh Timestamp, _
            TickerID, _
            Price
End Sub

Private Sub IStreamingDataConsumer_Low( _
                            ByVal TickerID As Long, _
                            ByVal Timestamp As Date, _
                            ByVal Price As Double)
ProcessLow Timestamp, _
            TickerID, _
            Price
End Sub

Private Sub IStreamingDataConsumer_OpenInterest( _
                            ByVal TickerID As Long, _
                            ByVal Timestamp As Date, _
                            ByVal Size As Long)

End Sub

Private Sub IStreamingDataConsumer_PreviousClose( _
                            ByVal TickerID As Long, _
                            ByVal Timestamp As Date, _
                            ByVal Price As Double)
ProcessClose Timestamp, _
            TickerID, _
            Price
End Sub

Private Sub IStreamingDataConsumer_ResetMarketDepth( _
                            ByVal TickerID As Long, _
                            ByVal Timestamp As Date)
DOMres Timestamp, _
        mTickerTable(TickerID).DOMEventsRequired, _
        TickerID, _
        False
End Sub

Private Sub IStreamingDataConsumer_Trade( _
                            ByVal TickerID As Long, _
                            ByVal Timestamp As Date, _
                            ByVal Price As Double, _
                            ByVal Size As Long)
ProcessLast Timestamp, _
            TickerID, _
            Price, _
            Size
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepth( _
                            ByVal TickerID As Long, _
                            ByVal Timestamp As Date, _
                            ByVal position As Long, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal Price As Double, _
                            ByVal Size As Long)
processMarketDepth Timestamp, _
                    TickerID, _
                    position, _
                    operation, _
                    side, _
                    Price, _
                    Size
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepthL2( _
                            ByVal TickerID As Long, _
                            ByVal Timestamp As Date, _
                            ByVal position As Long, _
                            ByVal marketMaker As String, _
                            ByVal operation As TradeBuildSP.DOMOperations, _
                            ByVal side As TradeBuildSP.DOMSides, _
                            ByVal Price As Double, _
                            ByVal Size As Long)

With mTickerTable(TickerID)
    If .inUse Then
        
        DOMRawUpdateL2 Timestamp, _
                        .DOMEventsRequired, _
                        TickerID, _
                        position, _
                        marketMaker, _
                        operation, _
                        side, _
                        Price, _
                        Size
        
        If .writeMktDepthToTickfile Then
            .TickfileWriter.writeMarketDepthData Timestamp, _
                                        position, _
                                        marketMaker, _
                                        operation, _
                                        side, _
                                        Price, _
                                        Size
        End If
    End If
End With

End Sub

Private Sub IStreamingDataConsumer_Volume( _
                            ByVal TickerID As Long, _
                            ByVal Timestamp As Date, _
                            ByVal Size As Long)
processVolume Timestamp, _
            TickerID, _
            Size
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
RaiseEvent connectionToTWSClosed(gCurrentTime)
mSocket.Close   ' VB Documentation indicates this should be done
Set mSocket = Nothing
End Sub

Private Sub mSocket_Connect()
Dim Timestamp As Date: Timestamp = gCurrentTime
mConnectionState = ConnectionStates.ConnConnecting
' On Error Resume Next
RaiseEvent connecting(Timestamp)
' If Err.Number <> 0 Then handleUserProgramError Err.Number, Err.description, "connecting"
' On Error GoTo 0
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
' On Error Resume Next
RaiseEvent dataReceived(gCurrentTime)
' If Err.Number <> 0 Then handleUserProgramError Err.Number, Err.description, "dataReceived"
' On Error GoTo 0

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
Dim Timestamp As Date: Timestamp = gCurrentTime

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
        RaiseEvent connectFailed(Timestamp, _
                                Description, _
                                IIf(mConnectionRetryIntervalSecs = 0, False, True))
        If mConnectionRetryIntervalSecs <> 0 Then
            If mConnectionTimer Is Nothing Then
                Set mConnectionTimer = New IntervalTimer
                mConnectionTimer.RepeatNotifications = False
                mConnectionTimer.TimerIntervalSecs = mConnectionRetryIntervalSecs
            End If
            mConnectionTimer.StartTimer
        End If
    Else
        SetServiceProviderError errorNum, Description, "IB TWS"
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

Private Sub mOrderSimulator_errorMessage(ByVal Timestamp As Date, _
                                        ByVal id As Long, _
                                        ByVal errorCode As Long, _
                                        ByVal errorMsg As String)
RaiseEvent errorMessage(Timestamp, _
                        id, _
                        errorCode, _
                        errorMsg)
End Sub

Private Sub mOrderSimulator_executionDetails(ByVal Timestamp As Date, _
                                            ByVal id As Long, _
                                            ByVal pContractSpecifier As contractSpecifier, _
                                            ByVal exec As Execution)
RaiseEvent ExecutionDetails(Timestamp, _
                            id, _
                            pContractSpecifier.clone, _
                            exec)
End Sub

Private Sub mOrderSimulator_openOrder(ByVal Timestamp As Date, _
                                    ByVal pContractSpecifier As contractSpecifier, _
                                    ByVal pOrder As Order)
RaiseEvent openOrder(Timestamp, _
                    pContractSpecifier.clone, _
                    pOrder)
End Sub

Private Sub mOrderSimulator_orderStatus(ByVal Timestamp As Date, _
                                        ByVal id As Long, _
                                        ByVal status As OrderStatuses, _
                                        ByVal filled As Long, _
                                        ByVal remaining As Long, _
                                        ByVal avgFillPrice As Double, _
                                        ByVal permId As Long, _
                                        ByVal parentId As Long, _
                                        ByVal lastFillPrice As Double, _
                                        ByVal clientID As Long)
RaiseEvent orderStatus(Timestamp, _
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

Private Sub mOrderSimulator_preFill(ByVal Timestamp As Date, _
                                    ByVal Order As Order, _
                                    ByVal pContractSpecifier As contractSpecifier, _
                                    fillPrice As Double, _
                                    fillSize As Long)
RaiseEvent preFill(Timestamp, _
                    Order, _
                    pContractSpecifier.clone, _
                    fillPrice, _
                    fillSize)
End Sub

'================================================================================
' Properties
'================================================================================

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
'If Not mTickFileProxy Is Nothing And value = False Then
'    err.Raise ErrorCodes.TickfileReplayProhibitsLiveOrders, _
'            "TradeBuild::(let)simulateOrders", _
'            "Live orders are not permitted during tickfile replay"
'End If
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

Friend Function AllocateTickerTableEntry(ByVal pTicker As ticker) As Long
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
    mOrderSimulator.MaxTickers = UBound(mTickerTable) + 1
    mTickers.MaxTickers = UBound(mTickerTable) + 1
End If

mTickerTable(AllocateTickerTableEntry).inUse = True
Set mTickerTable(AllocateTickerTableEntry).tickerObj = pTicker

End Function

Public Sub CancelMarketData(ByVal TickerID As Long)
If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            TickerID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If
Dim id As Long
Const Version = 2

If Not mTickerTable(TickerID).inUse Then
    err.Raise InvalidTickerID, _
    "TradeBuild.TradeBuildAPI::CancelMarketData", _
    "Invalid ticker id " & TickerID
End If

id = TickerID + mTickerTable(TickerID).incarnation

If mTickerTable(TickerID).receivingMarketDepth Then CancelMarketDepth TickerID, True
releaseTickerTableEntry TickerID
mOrderSimulator.stopMarket TickerID
addData CANCEL_MKT_DATA
addData Version
addData id
send

End Sub

Public Sub CancelMarketDepth(ByVal TickerID As Long, _
                            Optional ByVal forceCancel As Boolean = False)
If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                TickerID, _
                ApiErrorCodes.NotConnected, _
                LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Const Version = 1

With mTickerTable(TickerID)
    If Not .inUse Then err.Raise ErrorCodes.InvalidTickerID, _
                                    "TradeBuild.TradeBuildAPI::cancelMarketDepth", _
                                    "Invalid ticker id " & TickerID
    If Not .receivingMarketDepth Then
        err.Raise ErrorCodes.NotReceivingMarketDepth, _
                    "TradeBuild.TradeBuildAPI::cancelMarketDepth", _
                    "Ticker id " & TickerID & " not currently receiving market depth"
    End If
    .DOMEventsRequired = DOMEvents.DOMNoEvents
    ReDim .askPrices(0) As PriceTableEntry
    ReDim .bidPrices(0) As PriceTableEntry
    .maxAskPricesIndex = 0&
    .maxBidPricesIndex = 0&
    If .writeMktDataToTickfile And Not forceCancel Then Exit Sub
'    If Not mTickFileProxy Is Nothing Then Exit Sub
    .receivingMarketDepth = False
End With

addData CANCEL_MKT_DEPTH
addData Version
addData MARKET_DEPTH_TICKER_BASE + TickerID
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
    mConnectionState <> ConnectionStates.ConnReplaying _
Then
    RaiseEvent errorMessage(gCurrentTime, _
                        orderId, _
                        ApiErrorCodes.NotConnected, _
                        LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

Set lOrder = mAllOrders(CStr(orderId))
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

Public Function createOrder() As Order
Dim lOrder As Order
Set lOrder = New Order
lOrder.id = mNextOrderID
mAllOrders.add lOrder, CStr(mNextOrderID)
mNextOrderID = mNextOrderID + 1
Set createOrder = lOrder.clone   ' return a copy of the order to the app
End Function

Public Sub disconnect()
If Not mConnectionTimer Is Nothing Then mConnectionTimer.StopTimer
mConnectionState = ConnectionStates.ConnNotConnected
releaseAllTickerTableEntries
ReleaseSocket
mIndexIn = 0
mIndexOut = 0
mInputIndex = 0
RaiseEvent connectionToTWSClosed(Now)
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

Public Function NewContracts() As Contracts
Set NewContracts = New Contracts
End Function

Public Function NewTicker() As ticker
Static keySeq As Long
keySeq = keySeq + 1
Set NewTicker = mTickers.add("$" & keySeq)
End Function


Public Sub placeOrder(ByVal pContractSpecifier As contractSpecifier, _
                    ByVal pOrder As Order)

If mConnectionState <> ConnectionStates.ConnConnected And _
   mConnectionState <> ConnectionStates.ConnReplaying _
Then
    RaiseEvent errorMessage(gCurrentTime, _
                            pOrder.id, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

pOrder.clientID = mClientID

If mSimulateOrders Or pOrder.simulated Then
    mOrderSimulator.placeOrder pContractSpecifier, pOrder
    Exit Sub
End If

Const Version = 11
addData PLACE_ORDER
addData Version
addData pOrder.id

' send contract fields
With pContractSpecifier
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
        (UCase$(pContractSpecifier.exchange) = "LIFFE" Or _
        UCase$(pContractSpecifier.exchange) = "LSE") _
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
If mServerVersion >= 8 And pContractSpecifier.secType = SecurityTypes.SecTypeBag Then
    With pContractSpecifier
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
                    addData .openClose
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

Public Function RequestContract(ByVal pContractSpecifier As contractSpecifier) As Contracts

On Error GoTo err

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Function
End If

Set RequestContract = New Contracts
RequestContract.contractSpecifier = pContractSpecifier

reqContract -1, RequestContract, pContractSpecifier, False

Exit Function

err:
err.Raise err.Number, _
            "TradeBuild.TradeBuildAPI::RequestContract"
End Function

Friend Function RequestContractEx( _
                            ByVal pContracts As Contracts, _
                            ByVal pContractSpecifier As contractSpecifier) As Contracts

On Error GoTo err

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Function
End If


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

Public Function RequestMarketData(ByVal pContractSpecifier As contractSpecifier, _
                            Optional ByVal DOMEventsRequired As DOMEvents = DOMNoEvents, _
                            Optional ByVal writeToTickFile As Boolean = False, _
                            Optional ByVal includeMarketDepthInTickfile As Boolean = False, _
                            Optional ByVal noVolumeAdjustments As Boolean = False, _
                            Optional ByVal noImpliedTrades As Boolean = False) As Long
Dim ticker As ticker

Set ticker = mTickers.add(pContractSpecifier.key & "|" & Format(CLng(1000000000 * Rnd)), "0")
RequestMarketData = ticker.id
 
ticker.StartTicker pContractSpecifier, _
                            DOMEventsRequired, _
                            writeToTickFile, _
                            includeMarketDepthInTickfile, _
                            noVolumeAdjustments, _
                            noImpliedTrades
End Function

Friend Sub RequestMarketDataEx(ByVal pContractSpecifier As contractSpecifier, _
                            ByVal pTicker As ticker, _
                            Optional ByVal DOMEventsRequired As DOMEvents = DOMNoEvents, _
                            Optional ByVal pTickfileWriter As TickfileWriter = Nothing, _
                            Optional ByVal includeMarketDepthInTickfile As Boolean = False, _
                            Optional ByVal noVolumeAdjustments As Boolean = False, _
                            Optional ByVal noImpliedTrades As Boolean = False)

'Dim lContracts As Contracts

Dim lTickerID As Long

If mConnectionState <> ConnectionStates.ConnConnected Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

lTickerID = pTicker.id
UpdateTickerTableEntry lTickerID, _
                        pContractSpecifier, _
                        DOMEventsRequired, _
                        pTickfileWriter, _
                        includeMarketDepthInTickfile, _
                        noVolumeAdjustments, _
                        noImpliedTrades

If findTickerTableIndex(pContractSpecifier, lTickerID) <> -1 Then
    ' a ticker for this contract is already running
    With mTickerTable(lTickerID)
        .tickerObj.Contract = Nothing
        .tickerObj.id = -1
        .tickerObj.DuplicateTickerRequest .contractSpecifier
        mTickers.DuplicateTickerRequest .tickerObj, .contractSpecifier
'        RaiseEvent errorMessage(gCurrentTime, _
'                                lTickerID, _
'                                ApiErrorCodes.TickerAlreadyRunningForThisContract, _
'                                "A ticker has already been started for the contract specified in this ticker")
        releaseTickerTableEntry lTickerID
    End With
Else
    reqMarketData lTickerID
End If

'Set lContracts = New Contracts
'reqContract pTicker.id, _
'            lContracts, _
'            pContractSpecifier, _
'            True

End Sub

Public Sub RequestMarketDepth(ByVal TickerID As Long, _
                            ByVal DOMEventsRequired As DOMEvents, _
                            ByVal writeToTickFile As Boolean)
                            
Dim i As Long
Dim listendata As TickerListenMDData

If mConnectionState <> ConnectionStates.ConnConnected And _
    mConnectionState <> ConnectionStates.ConnReplaying _
Then
    RaiseEvent errorMessage(gCurrentTime, _
                            NoValidID, _
                            ApiErrorCodes.NotConnected, _
                            LoadResString(ApiErrorCodes.NotConnected))
    Exit Sub
End If

With mTickerTable(TickerID)
    If .contractSpecifier Is Nothing Then err.Raise ErrorCodes.InvalidTickerID, "TradeBuild.TradeBuildAPI::requestMarketDepth", "Invalid ticker id " & TickerID
    If Not .receivingMarketDepth Then
        .receivingMarketDepth = True
        .DOMEventsRequired = DOMEventsRequired
        ReDim .askPrices(20) As PriceTableEntry
        ReDim .bidPrices(20) As PriceTableEntry
        If writeToTickFile Then
            .writeMktDepthToTickfile = True
            If .TickfileWriter Is Nothing Then
                Set .TickfileWriter = New TickFileManager
'                .tickfile.OutputTickfileFormat = mOutputTickfileFormat
'                .tickfile.Path = mTickFilePath
            End If
        Else
            .writeMktDepthToTickfile = False
            Set .TickfileWriter = Nothing
        End If
    Else
        If writeToTickFile Then
            .writeMktDepthToTickfile = True
            If .TickfileWriter Is Nothing Then
                Set .TickfileWriter = New TickFileManager
'                .tickfile.OutputTickfileFormat = mOutputTickfileFormat
'                .tickfile.Path = mTickFilePath
            End If
        Else
            .writeMktDepthToTickfile = False
            Set .TickfileWriter = Nothing
        End If
        
        DOMres .tickerObj.Timestamp, _
                .DOMEventsRequired, _
                TickerID, _
                False
        
        If DOMEventsRequired <> DOMNoEvents Then
            For i = 0 To UBound(.bidPrices)
                If .bidPrices(i).Size <> 0 Then
                    
                    DOMRawUpdate .tickerObj.Timestamp, _
                                DOMEventsRequired, _
                                TickerID, _
                                i, _
                                DOMOperations.DOMInsert, _
                                DOMSides.DOMBid, _
                                .bidPrices(i).Price, _
                                .bidPrices(i).Size
                    
                    DOMset .tickerObj.Timestamp, _
                            DOMEventsRequired, _
                            TickerID, _
                            DOMSides.DOMBid, _
                            .bidPrices(i).Price, _
                            .bidPrices(i).Size

                End If
            Next
        
            For i = 0 To UBound(.askPrices)
                If .askPrices(i).Size <> 0 Then
                    
                    DOMRawUpdate .tickerObj.Timestamp, _
                                DOMEventsRequired, _
                                TickerID, _
                                i, _
                                DOMOperations.DOMInsert, _
                                DOMSides.DOMAsk, _
                                .askPrices(i).Price, _
                                .askPrices(i).Size
                    
                    DOMset .tickerObj.Timestamp, _
                            DOMEventsRequired, _
                            TickerID, _
                            DOMSides.DOMAsk, _
                            .askPrices(i).Price, _
                            .askPrices(i).Size

                End If
            Next
        End If
        
        .DOMEventsRequired = DOMEventsRequired
        
        Exit Sub
    End If
End With

reqMarketDepth TickerID

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
mServiceProviderError.Timestamp = gCurrentTime
RaiseEvent errorMessage(gCurrentTime, _
                        NoValidID, _
                        ApiErrorCodes.ServiceProviderErrorNotification, _
                        "A Service Provider error has occurred")
                    
End Sub

Friend Sub StartSimulatorMarket( _
                            ByVal TickerID As Long, _
                            ByVal pContractSpec As contractSpecifier, _
                            ByVal simulateExchangeDelays As Boolean)
mOrderSimulator.simulateExchangeDelays = simulateExchangeDelays
mOrderSimulator.startMarket TickerID, pContractSpec

End Sub

Friend Sub UpdateTickerTableEntry(ByVal TickerID As Long, _
                                ByVal pContractSpec As contractSpecifier, _
                                ByVal DOMEventsRequired As DOMEvents, _
                                ByVal pTickfileWriter As TickfileWriter, _
                                ByVal includeMarketDepthInTickfile As Boolean, _
                                ByVal noVolumeAdjustments As Boolean, _
                                ByVal noImpliedTrades As Boolean)
With mTickerTable(TickerID)
    Set .contractSpecifier = pContractSpec
    .DOMEventsRequired = DOMEventsRequired
    .noVolumeAdjustments = noVolumeAdjustments
    .noImpliedTrades = noImpliedTrades
    If Not pTickfileWriter Is Nothing Then
        .writeMktDataToTickfile = True
        .writeMktDepthToTickfile = includeMarketDepthInTickfile
        Set .TickfileWriter = pTickfileWriter
    ElseIf Not .TickfileWriter Is Nothing Then
        .writeMktDataToTickfile = False
        .writeMktDepthToTickfile = False
        Set .TickfileWriter = Nothing
    End If
    If (.writeMktDataToTickfile And .writeMktDepthToTickfile) Or _
        .DOMEventsRequired <> DOMEvents.DOMNoEvents _
    Then
        If Not .receivingMarketDepth Then
            .receivingMarketDepth = True
            ReDim .askPrices(20) As PriceTableEntry
            ReDim .bidPrices(20) As PriceTableEntry
        End If
    End If
End With
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
    Set mAllOrders = New Collection
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

Private Function findTickerTableIndex(ByVal pContractSpecifier As contractSpecifier, _
                                    Optional ByVal excludeIndex As Long = -1) As Long
Dim i As Long

findTickerTableIndex = -1

For i = 0 To UBound(mTickerTable)
    If i <> excludeIndex Then
        If Not mTickerTable(i).contractSpecifier Is Nothing Then
            If mTickerTable(i).contractSpecifier.Equals(pContractSpecifier) Then
                findTickerTableIndex = i
                Exit For
            End If
        End If
    End If
Next

End Function

Private Sub DOMclear( _
                        ByVal Timestamp As Date, _
                        ByVal DOMEventsRequired As DOMEvents, _
                        ByVal TickerID As Long, _
                        ByVal side As DOMSides, _
                        ByVal Price As Double)
Dim listendata As TickerListenMDData
If DOMEventsRequired = DOMBothEvents Or DOMEventsRequired = DOMProcessedEvents Then
    With mTickerTable(TickerID)
        If Not .tickerObj Is Nothing Then
            .tickerObj.DOMClearCell Timestamp, _
                                side, _
                                Price
            mTickers.DOMClearCell TickerID, _
                                Timestamp, _
                                side, _
                                Price
                                
        End If
    End With
    RaiseEvent DOMClearCell(Timestamp, _
                            TickerID, _
                            side, _
                            Price)
    If mEnableListeners Then
        listendata.TickerID = TickerID
        listendata.side = side
        listendata.Price = Price
        listendata.Size = 0
        mListeners.notify listendata, _
                        TradeBuildListenValueTypes.ValueTypeTradeBuildMarketdepth, _
                        Timestamp
    End If
End If
End Sub

Private Sub DOMset( _
                        ByVal Timestamp As Date, _
                        ByVal DOMEventsRequired As DOMEvents, _
                        ByVal TickerID As Long, _
                        ByVal side As DOMSides, _
                        ByVal Price As Double, _
                        ByVal Size As Long)
Dim listendata As TickerListenMDData
If DOMEventsRequired = DOMBothEvents Or DOMEventsRequired = DOMProcessedEvents Then
    With mTickerTable(TickerID)
        If Not .tickerObj Is Nothing Then
            .tickerObj.DOMSetCell Timestamp, _
                                side, _
                                Price, _
                                Size
            mTickers.DOMSetCell TickerID, _
                                Timestamp, _
                                side, _
                                Price, _
                                Size
        End If
    End With
    RaiseEvent DOMSetCell(Timestamp, _
                            TickerID, _
                            side, _
                            Price, _
                            Size)
    If mEnableListeners Then
        listendata.TickerID = TickerID
        listendata.side = side
        listendata.Price = Price
        listendata.Size = Size
        mListeners.notify listendata, _
                        TradeBuildListenValueTypes.ValueTypeTradeBuildMarketdepth, _
                        Timestamp
    End If
End If
End Sub


Private Sub DOMRawUpdate( _
                        ByVal Timestamp As Date, _
                        ByVal DOMEventsRequired As DOMEvents, _
                        ByVal TickerID As Long, _
                        ByVal position As Long, _
                        ByVal operation As DOMOperations, _
                        ByVal side As DOMSides, _
                        ByVal Price As Double, _
                        ByVal Size As Long)
Dim listendata As TickerListenRawMDData
If DOMEventsRequired = DOMRawEvents Or _
    DOMEventsRequired = DOMBothEvents _
Then
    With mTickerTable(TickerID)
        If Not .tickerObj Is Nothing Then
            .tickerObj.UpdateMarketDepth Timestamp, _
                                        position, _
                                        operation, _
                                        side, _
                                        Price, _
                                        Size
        End If
    End With
    RaiseEvent UpdateMarketDepth(Timestamp, _
                                TickerID, _
                                position, _
                                operation, _
                                side, _
                                Price, _
                                Size)
    
    If mEnableListeners Then
        listendata.TickerID = TickerID
        listendata.position = position
        listendata.operation = operation
        listendata.side = side
        listendata.Price = Price
        listendata.Size = Size
        mListeners.notify listendata, _
                            TradeBuildListenValueTypes.ValueTypeTradeBuildRawMarketdepth, _
                            Timestamp
    End If
End If

End Sub

Private Sub DOMRawUpdateL2( _
                        ByVal Timestamp As Date, _
                        ByVal DOMEventsRequired As DOMEvents, _
                        ByVal TickerID As Long, _
                        ByVal position As Long, _
                        ByVal marketMaker As String, _
                        ByVal operation As DOMOperations, _
                        ByVal side As DOMSides, _
                        ByVal Price As Double, _
                        ByVal Size As Long)
Dim listendata As TickerListenRawMDData
If DOMEventsRequired = DOMRawEvents Or _
    DOMEventsRequired = DOMBothEvents _
Then
    With mTickerTable(TickerID)
        If Not .tickerObj Is Nothing Then
            .tickerObj.UpdateMarketDepthL2 Timestamp, _
                                        position, _
                                        marketMaker, _
                                        operation, _
                                        side, _
                                        Price, _
                                        Size
        End If
    End With
    RaiseEvent UpdateMarketDepthL2(Timestamp, _
                                TickerID, _
                                position, _
                                marketMaker, _
                                operation, _
                                side, _
                                Price, _
                                Size)
    
    If mEnableListeners Then
        listendata.TickerID = TickerID
        listendata.position = position
        listendata.marketMaker = marketMaker
        listendata.operation = operation
        listendata.side = side
        listendata.Price = Price
        listendata.Size = Size
        mListeners.notify listendata, _
                            TradeBuildListenValueTypes.ValueTypeTradeBuildRawMarketdepthL2, _
                            Timestamp
    End If
End If

End Sub

Private Sub DOMres( _
                        ByVal Timestamp As Date, _
                        ByVal DOMEventsRequired As DOMEvents, _
                    ByVal TickerID As Long, _
                    ByVal marketDataReRequested As Boolean)
Dim listendata As TickerListenMDData
If DOMEventsRequired <> DOMNoEvents Then
    With mTickerTable(TickerID)
        If Not .tickerObj Is Nothing Then
            .tickerObj.DOMReset Timestamp, _
                                marketDataReRequested
            mTickers.DOMReset TickerID, _
                                Timestamp, _
                                marketDataReRequested
        End If
        RaiseEvent DOMReset(Timestamp, _
                            TickerID, _
                            marketDataReRequested)
    End With
    
    If mEnableListeners Then
        listendata.TickerID = TickerID
        listendata.marketDataReRequested = marketDataReRequested
        mListeners.notify listendata, _
                        TradeBuildListenValueTypes.ValueTypeTradeBuildMarketdepth, _
                        Timestamp
    End If
End If
    

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
err.Raise DataIncomplete
End Function

Private Sub getAccountTime(ByVal Version As Long)
Dim listendata As TWSSocketMsgAcctUpdateTime
Dim Timestamp As Date: Timestamp = gCurrentTime
Dim accountTime As String: accountTime = getString

RaiseEvent updateAccountTime(Timestamp, accountTime)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.ACCT_UPDATE_TIME
    listendata.accountTime = accountTime
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgAcctUpdateTime, _
                    Timestamp
End If

End Sub

Private Sub getAccountValue(ByVal Version As Long)
Dim listendata As TWSSocketMsgAcctValue
Dim Timestamp As Date: Timestamp = gCurrentTime
Dim key As String: key = getString
Dim val As String: val = getString
Dim cur As String: cur = getString
Dim accountName As String: If Version >= 2 Then accountName = getString

RaiseEvent updateAccountValue(Timestamp, key, val, cur, accountName)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.ACCT_VALUE
    listendata.accountName = accountName
    listendata.currencyCode = cur
    listendata.key = key
    listendata.value = val
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgAcctValue, _
                    Timestamp
End If

End Sub

Private Sub getContractData(ByVal Version As Long)
Dim listendata As TWSSocketMsgContractData
Dim Timestamp As Date
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

Timestamp = gCurrentTime

Set lContract = New Contract
With lContract
    .Timestamp = Timestamp
    .specifier = New contractSpecifier
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
                    Timestamp
End If

End Sub

Private Sub getErrorMsg(ByVal Version As Long)
Dim listendata As TWSSocketMsgErrMsg
Dim Timestamp As Date
Dim id As Long
Dim errorCode As Long
Dim errorMsg As String
Dim execFilter As ExecutionFilter
id = getLong
errorCode = getLong
errorMsg = getString
Timestamp = gCurrentTime


Select Case errorCode
Case 200    ' security definition not known
    If id >= 0 Then
        ' this must relate to a reqMktData - this should never occur as
        ' we make sure the contract is valid before starting the ticker
        With mTickerTable(id Mod INCARNATION_INCREMENT)
            RaiseEvent ContractInvalid(Timestamp, _
                                        .contractSpecifier)
            .tickerObj.ContractInvalid
            mTickers.ContractInvalid .tickerObj, .contractSpecifier
            releaseTickerTableEntry (id Mod INCARNATION_INCREMENT)
        End With
    Else
        ' this must relate to a reqContract
        processContractError
    End If
Case 202
    RaiseEvent orderStatus(Timestamp, id, OrderStatuses.OrderStatusCancelled, 0, 0, 0, 0, 0, 0, mClientID)
Case 309
    ' max market depth requests exceeded
    mTickerTable(id - MARKET_DEPTH_TICKER_BASE).tickerObj.MarketDepthNotAvailable errorMsg
    mTickers.MarketDepthNotAvailable gCurrentTime, id - MARKET_DEPTH_TICKER_BASE, errorMsg
Case 316
    ResetMarketDepth Timestamp, reEstablish:=True
Case 317
    ResetMarketDepth Timestamp, reEstablish:=False
Case 1100
    ' connectivity between TWS and IB has been lost
    RaiseEvent connectionToIBClosed(Timestamp)
Case 1101
    ' connectivity between TWS and IB has been restored, but data has been
    ' lost, so need to reestablish market data requests
    
    reEstablishMarketData
    
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

    RaiseEvent connectionToIBRecovered(Timestamp)
    
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
                    Timestamp
End If

End Sub

Private Sub getExecutionData(ByVal Version As Long)
Dim listendata As TWSSocketMsgExecutionData
Dim Timestamp As Date: Timestamp = gCurrentTime
Dim secType As String
Dim right As String
Dim side As String
Dim Price As Double

Dim id As Long: id = getLong

Dim lContractSpecifier As contractSpecifier
Set lContractSpecifier = New contractSpecifier
With lContractSpecifier
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
    Price = getDouble
    If lContractSpecifier.currencyCode = "GBP" Then
        .Price = Price * 100
    Else
        .Price = Price
    End If
End With

If Version >= 2 Then exec.permId = getLong

If Version >= 3 Then exec.clientID = getLong
   
RaiseEvent ExecutionDetails(Timestamp, id, lContractSpecifier.clone, exec)

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
        listendata.Price = Price
        listendata.quantity = .quantity
        listendata.side = side
        listendata.time = .time
    End With
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgExecutionData, _
                    Timestamp
End If

End Sub

Private Sub getFAData(ByVal Version As Long)
Dim listendata As TWSSocketMsgReceiveFa
Dim Timestamp As Date: Timestamp = gCurrentTime
Dim DataType As FADataTypes: DataType = getLong
Dim XMLData As String: XMLData = getString

RaiseEvent FAData(Timestamp, DataType, XMLData)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.MANAGED_ACCTS
    listendata.DataType = DataType
    listendata.XMLData = XMLData
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgManagedAccts, _
                    Timestamp
End If

End Sub

Private Sub getManagedAccounts(ByVal Version As Long)
Dim listendata As TWSSocketMsgManagedAccts
Dim Timestamp As Date: Timestamp = gCurrentTime
Dim accountsList As String: accountsList = getString

RaiseEvent managedAccounts(Timestamp, accountsList)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.MANAGED_ACCTS
    listendata.accountsList = accountsList
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgManagedAccts, _
                    Timestamp
End If

End Sub

Private Sub getMarketDepth(ByVal Version As Long)
Dim listendata As TWSSocketMsgMarketDepth
Dim Timestamp As Date
Dim id As Long: id = getLong
Dim position As Long: position = getLong
Dim operation As Long: operation = getLong
Dim side As Long: side = getLong
Dim Price As Double: Price = getDouble
Dim Size As Long: Size = getLong

Timestamp = gCurrentTime

processMarketDepth Timestamp, _
                id - MARKET_DEPTH_TICKER_BASE, _
                position, _
                operation, _
                side, _
                Price, _
                Size

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.MARKET_DEPTH
    listendata.operation = operation
    listendata.position = position
    listendata.Price = Price
    listendata.side = side
    listendata.Size = Size
    listendata.TickerID = id
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgMarketDepth, _
                    Timestamp
End If

End Sub

Private Sub getMarketDepthL2(ByVal Version As Long)
Dim listendata As TWSSocketMsgMarketDepthL2
Dim Timestamp As Date: Timestamp = gCurrentTime
Dim id As Long: id = getLong
Dim position As Long: position = getLong
Dim marketMaker As String: marketMaker = getString
Dim operation As Long: operation = getLong
Dim side As Long: side = getLong
Dim Price As Double: Price = getDouble
Dim Size As Long: Size = getLong

With mTickerTable(id - MARKET_DEPTH_TICKER_BASE)
    If .inUse Then
        
        DOMRawUpdateL2 Timestamp, _
                        .DOMEventsRequired, _
                        id - MARKET_DEPTH_TICKER_BASE, _
                        position, _
                        marketMaker, _
                        operation, _
                        side, _
                        Price, _
                        Size
        
        If .writeMktDepthToTickfile Then
            .TickfileWriter.writeMarketDepthData Timestamp, _
                                        position, _
                                        marketMaker, _
                                        operation, _
                                        side, _
                                        Price, _
                                        Size
        End If
    End If
End With

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.MARKET_DEPTH_L2
    listendata.marketMaker = marketMaker
    listendata.operation = operation
    listendata.position = position
    listendata.Price = Price
    listendata.side = side
    listendata.Size = Size
    listendata.TickerID = id
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgMarketDepthL2, _
                    Timestamp
End If

End Sub

Private Sub getNewsBulletins(ByVal Version As Long)
Dim listendata As TWSSocketMsgNewsBulletins
Dim Timestamp As Date: Timestamp = gCurrentTime
Dim msgID As Long: msgID = getLong
Dim msgType As Long: msgType = getLong
Dim newsMessage As String: newsMessage = getString
Dim originatingExch As String: originatingExch = getString

RaiseEvent updateNewsBulletin(Timestamp, msgID, msgType, newsMessage, originatingExch)

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.NEWS_BULLETINS
    listendata.newsID = msgID
    listendata.newsMessage = newsMessage
    listendata.newsType = msgType
    listendata.originatingExch = originatingExch
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgNewsBulletins, _
                    Timestamp
End If

End Sub

Private Sub getNextValidId(ByVal Version As Long)
Dim listendata As TWSSocketMsgNextValidId
Dim Timestamp As Date: Timestamp = gCurrentTime

mNextOrderID = getLong

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.NEXT_VALID_ID
    listendata.id = mNextOrderID
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgNextValidId, _
                    Timestamp
End If

End Sub

Private Sub getOpenOrder(ByVal Version As Long)
Dim listendata As TWSSocketMsgOpenOrder
Dim Timestamp As Date: Timestamp = gCurrentTime
Dim lOrder As Order
Dim id As Long
Dim secType As String
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
Dim lContractSpecifier As contractSpecifier
Set lContractSpecifier = New contractSpecifier
With lContractSpecifier
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

RaiseEvent openOrder(Timestamp, lContractSpecifier.clone, lOrder.clone)

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
                    Timestamp
End If

End Sub

Private Sub getOrderStatus(ByVal Version As Long)
Dim listendata As TWSSocketMsgOrderStatus
Dim Timestamp As Date: Timestamp = gCurrentTime
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
    RaiseEvent orderStatus(Timestamp, _
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
                    Timestamp
End If

End Sub

Private Sub getPortfolioValue(ByVal Version As Long)
Dim listendata As TWSSocketMsgPortfolioValue
Dim Timestamp As Date: Timestamp = gCurrentTime
Dim secType As String
Dim right As String
Dim lContractSpecifier As contractSpecifier

Set lContractSpecifier = New contractSpecifier
With lContractSpecifier
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
                         
RaiseEvent updatePortfolioValue(Timestamp, lContractSpecifier.clone, position, marketPrice, marketValue, averageCost, unrealizedPNL, realizedPNL, accountName)

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
                    Timestamp
End If

End Sub

Private Sub getTickPrice(ByVal Version As Long)
Dim listendata As TWSSocketMsgTickPrice
Dim Timestamp As Date
Dim Price As Double
Dim Size As Long
Dim TickType As Long
Dim TickerID As Long
TickerID = getLong
TickType = getLong
Price = getDouble
If Version >= 2 Then Size = getLong
Timestamp = gCurrentTime

processTickPrice Timestamp, _
                TickerID Mod INCARNATION_INCREMENT, _
                TickType, _
                Price, _
                Size

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.TICK_PRICE
    listendata.Price = Price
    listendata.Size = Size
    listendata.TickerID = TickerID
    listendata.TickType = TickType
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgTickPrice, _
                    Timestamp
End If

End Sub

Private Sub getTickSize(ByVal Version As Long)
Dim listendata As TWSSocketMsgTickSize
Dim Timestamp As Date
Dim Size As Long
Dim TickType As Long
Dim TickerID As Long
TickerID = getLong
TickType = getLong
Size = getLong
Timestamp = gCurrentTime

processTickSize Timestamp, _
                TickerID Mod INCARNATION_INCREMENT, _
                TickType, _
                Size

If mEnableListeners Then
    listendata.msgType = TWSSocketInMsgTypes.TICK_PRICE
    listendata.Size = Size
    listendata.TickerID = TickerID
    listendata.TickType = TickType
    listendata.Version = Version
    
    mListeners.notify listendata, _
                    TradeBuildListenValueTypes.ValueTypeTWSSocketMsgTickSize, _
                    Timestamp
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

Private Sub ProcessAsk(ByVal Timestamp As Date, _
                            ByVal TickerID As Long, _
                            ByVal Price As Double, _
                            ByVal Size As Long)
Dim tickListendata As TickerListenData

With mTickerTable(TickerID)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    If Not .tickerObj Is Nothing Then
        .tickerObj.setAsk Timestamp, _
                    Price, _
                    Size
        mTickers.ask Timestamp, _
                    TickerID, _
                    Price, _
                    Size
    End If
    RaiseEvent ask(Timestamp, _
                    TickerID, _
                    Price, _
                    Size)
    
    If mSimulateOrders Then
        mOrderSimulator.ask Timestamp, _
                            TickerID Mod INCARNATION_INCREMENT, _
                            Price
    End If
    
    If mEnableListeners Then
        tickListendata.TickerID = TickerID
        tickListendata.Price = Price
        tickListendata.Size = Size
        mListeners.notify tickListendata, _
                        TradeBuildListenValueTypes.ValueTypeTradeBuildAsk, _
                        Timestamp
    End If
    
    If .prevAsk <> 0 And Price > .prevAsk Then
        DOMclear Timestamp, _
                .DOMEventsRequired, _
                TickerID, _
                DOMSides.DOMAsk, _
                .prevAsk
    End If
    
    If Size <> 0 Then
        DOMset Timestamp, _
                .DOMEventsRequired, _
                TickerID, _
                DOMSides.DOMAsk, _
                Price, _
                Size
    End If
    
    If .writeMktDataToTickfile Then _
        .TickfileWriter.writeAsk Timestamp, Price, Size

    .prevAsk = Price
    If Size <> 0 Then .prevAskSize = Size
End With
End Sub

Private Sub ProcessBid(ByVal Timestamp As Date, _
                            ByVal TickerID As Long, _
                            ByVal Price As Double, _
                            ByVal Size As Long)
Dim tickListendata As TickerListenData

With mTickerTable(TickerID)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    If Not .tickerObj Is Nothing Then
        .tickerObj.setBid Timestamp, _
                    Price, _
                    Size
        mTickers.bid Timestamp, _
                    TickerID, _
                    Price, _
                    Size
    End If
    RaiseEvent bid(Timestamp, _
                    TickerID, _
                    Price, _
                    Size)
    
    If mSimulateOrders Then
        mOrderSimulator.bid Timestamp, _
                            TickerID Mod INCARNATION_INCREMENT, _
                            Price
    End If
    
    If mEnableListeners Then
        tickListendata.TickerID = TickerID
        tickListendata.Price = Price
        tickListendata.Size = Size
        mListeners.notify tickListendata, _
                        TradeBuildListenValueTypes.ValueTypeTradeBuildBid, _
                        Timestamp
    End If
    
    If .prevBid <> 0 And Price < .prevBid Then
        DOMclear Timestamp, _
                .DOMEventsRequired, _
                TickerID, _
                DOMSides.DOMBid, _
                .prevBid
    End If
    
    If Size <> 0 Then
        DOMset Timestamp, _
                .DOMEventsRequired, _
                TickerID, _
                DOMSides.DOMBid, _
                Price, _
                Size
    End If
    
    If .writeMktDataToTickfile Then _
        .TickfileWriter.writeBid Timestamp, Price, Size
       
    .prevBid = Price
    If Size <> 0 Then .prevBidSize = Size
End With
End Sub

Private Sub ProcessClose(ByVal Timestamp As Date, _
                            ByVal TickerID As Long, _
                            ByVal Price As Double)
Dim tickListendata As TickerListenData

With mTickerTable(TickerID)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    If Not .tickerObj Is Nothing Then
        With .tickerObj
            .setClosePrice Timestamp, Price
            mTickers.PrevClose Timestamp, TickerID, Price, .Change, .ChangePercent
        End With
    End If
    RaiseEvent PreviousClose(Timestamp, TickerID, Price)
    
    If mEnableListeners Then
        tickListendata.TickerID = TickerID
        tickListendata.Price = Price
        tickListendata.Size = 0
        mListeners.notify tickListendata, _
                        TradeBuildListenValueTypes.ValueTypeTradeBuildClose, _
                        Timestamp
    End If
    
    If .writeMktDataToTickfile Then _
        .TickfileWriter.writeClose Timestamp, Price

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
    .Contracts.AddContract pContract
    If Not .InternalRequest Then
        RaiseEvent Contract(pContract.Timestamp, pContract.clone)
    End If
End With

End Sub

Private Sub processContractError()
Dim TickerID As Long
Dim lContract As Contract

With mContractDetailsRequest
    If .Contracts.Count > 0 Then
        ' we are processing the forced invalid contract details request
        ' for a request for which we have received one or more contract details
        If .InternalRequest Then
            TickerID = .TickerID
            If .Contracts.Count = 1 Then
                Set lContract = .Contracts(1)
                With mTickerTable(TickerID)
                    Set .contractSpecifier = lContract.specifier
                    If Not .tickerObj Is Nothing Then
                        .tickerObj.Contract = lContract
                    End If
                    If .writeMktDataToTickfile Or _
                        .writeMktDepthToTickfile _
                    Then
                        .TickfileWriter.Contract = lContract
                    End If
                    If findTickerTableIndex(.contractSpecifier, TickerID) <> -1 Then
                        ' a ticker for this contract is already running
                        .tickerObj.Contract = Nothing
                        .tickerObj.id = -1
                        .tickerObj.DuplicateTickerRequest .contractSpecifier
                        mTickers.DuplicateTickerRequest .tickerObj, .contractSpecifier
                        RaiseEvent errorMessage(gCurrentTime, _
                                                TickerID, _
                                                ApiErrorCodes.TickerAlreadyRunningForThisContract, _
                                                "A ticker has already been started for the contract specified in this ticker")
                        releaseTickerTableEntry TickerID
                    Else
                        reqMarketData TickerID
                    End If
                End With
            Else
                ' the requested contract is not unique
                With mTickerTable(TickerID)
                    .tickerObj.ContractInvalid
                    mTickers.ContractInvalid .tickerObj, .contractSpecifier
                End With
                releaseTickerTableEntry TickerID
            End If
        Else
            .Contracts.LoadComplete
        End If
    Else
        ' this is a genuine error response to a contract details request
        RaiseEvent ContractInvalid(gCurrentTime, .contractSpecifier)
        If .InternalRequest Then
            TickerID = .TickerID
            With mTickerTable(TickerID)
                .tickerObj.ContractInvalid
                mTickers.ContractInvalid .tickerObj, .contractSpecifier
            End With
            releaseTickerTableEntry TickerID
        Else
            .Contracts.ContractSpecifierInvalid .contractSpecifier
            .Contracts.LoadComplete
        End If
    End If
End With
        
Set mContractDetailsRequest = Nothing
mFirstContractDetailsReceived = False

End Sub

Private Sub ProcessHigh(ByVal Timestamp As Date, _
                            ByVal TickerID As Long, _
                            ByVal Price As Double)
Dim tickListendata As TickerListenData

With mTickerTable(TickerID)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    If Not .tickerObj Is Nothing Then
       .tickerObj.setHighPrice Timestamp, Price
       mTickers.High Timestamp, TickerID, Price
    End If
    RaiseEvent High(Timestamp, TickerID, Price)
    
    If mEnableListeners Then
        tickListendata.TickerID = TickerID
        tickListendata.Price = Price
        tickListendata.Size = 0
        mListeners.notify tickListendata, _
                        TradeBuildListenValueTypes.ValueTypeTradeBuildHigh, _
                        Timestamp
    End If
    
    If .writeMktDataToTickfile Then _
        .TickfileWriter.writeHigh Timestamp, Price

End With
End Sub

Private Sub ProcessLast(ByVal Timestamp As Date, _
                            ByVal TickerID As Long, _
                            ByVal Price As Double, _
                            ByVal Size As Long)
Dim tickListendata As TickerListenData

With mTickerTable(TickerID)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
        
    If Not .tickerObj Is Nothing Then
        With .tickerObj
            .setTrade Timestamp, _
                    Price, _
                    Size
            mTickers.Last Timestamp, _
                        TickerID, _
                        Price, _
                        Size, _
                        .Change, _
                        .ChangePercent
        End With
    End If
    RaiseEvent trade(Timestamp, _
                    TickerID, _
                    Price, _
                    Size)
                    
    If mSimulateOrders Then
        mOrderSimulator.trade Timestamp, _
                            TickerID Mod INCARNATION_INCREMENT, _
                            Price
    End If
    
    If mEnableListeners Then
        tickListendata.TickerID = TickerID
        tickListendata.Price = Price
        tickListendata.Size = Size
        mListeners.notify tickListendata, _
                        TradeBuildListenValueTypes.ValueTypeTradeBuildTrade, _
                        Timestamp
    End If
    
    If .prevTrade <> 0 And Price <> .prevTrade Then
        DOMclear Timestamp, _
                .DOMEventsRequired, _
                TickerID, _
                DOMSides.DOMLast, _
                .prevTrade
    End If
    
    If Size <> 0 Then
        DOMset Timestamp, _
                .DOMEventsRequired, _
                TickerID, _
                DOMSides.DOMLast, _
                Price, _
                Size
    End If
   
    If .writeMktDataToTickfile Then _
        .TickfileWriter.writeTrade Timestamp, Price, Size

    .accumulatedVolume = .accumulatedVolume + .prevTradeSize

    .prevTrade = Price
    If Size <> 0 Then .prevTradeSize = Size
                            
End With
End Sub

Private Sub ProcessLow(ByVal Timestamp As Date, _
                            ByVal TickerID As Long, _
                            ByVal Price As Double)
Dim tickListendata As TickerListenData

With mTickerTable(TickerID)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    If Not .tickerObj Is Nothing Then
       .tickerObj.setLowPrice Timestamp, Price
       mTickers.Low Timestamp, TickerID, Price
    End If
    RaiseEvent Low(Timestamp, TickerID, Price)
    
    If mEnableListeners Then
        tickListendata.TickerID = TickerID
        tickListendata.Price = Price
        tickListendata.Size = 0
        mListeners.notify tickListendata, _
                        TradeBuildListenValueTypes.ValueTypeTradeBuildLow, _
                        Timestamp
    End If
    
    If .writeMktDataToTickfile Then _
        .TickfileWriter.writeLow Timestamp, Price

End With
End Sub

Private Sub processMarketDepth(ByVal Timestamp As Date, _
                                ByVal id As Long, _
                                ByVal position As Long, _
                                ByVal operation As DOMOperations, _
                                ByVal side As DOMSides, _
                                ByVal Price As Double, _
                                ByVal Size As Long)

With mTickerTable(id)

    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    DOMRawUpdate Timestamp, _
                .DOMEventsRequired, _
                id, _
                position, _
                operation, _
                side, _
                Price, _
                Size
    
    If side = DOMSides.DOMAsk Then
        Select Case operation
        Case DOMOperations.DOMInsert
            If position > .maxAskPricesIndex Then
                .maxAskPricesIndex = position
            Else
                .maxAskPricesIndex = .maxAskPricesIndex + 1
            End If
        Case DOMOperations.DOMUpdate
            If position > .maxAskPricesIndex Then .maxAskPricesIndex = position
        Case DOMOperations.DOMDelete
            If position > .maxAskPricesIndex Then
                .maxAskPricesIndex = position
            End If
        End Select
        If .maxAskPricesIndex >= UBound(.askPrices) Then
            ReDim Preserve .askPrices(UBound(.askPrices) + 10) As PriceTableEntry
        End If
        updatePrices Timestamp, _
                    id, _
                    .askPrices, _
                    .maxAskPricesIndex, _
                    DOMSides.DOMAsk, _
                    position, _
                    operation, _
                    Price, _
                    Size, _
                    .prevBid, _
                    .prevAsk, _
                    .DOMEventsRequired = DOMProcessedEvents Or .DOMEventsRequired = DOMBothEvents
        If operation = DOMOperations.DOMDelete Then .maxAskPricesIndex = .maxAskPricesIndex - 1
    Else
        Select Case operation
        Case DOMOperations.DOMInsert
            If position > .maxBidPricesIndex Then
                .maxBidPricesIndex = position
            Else
                .maxBidPricesIndex = .maxBidPricesIndex + 1
            End If
        Case DOMOperations.DOMUpdate
            If position > .maxBidPricesIndex Then .maxBidPricesIndex = position
        Case DOMOperations.DOMDelete
            If position > .maxBidPricesIndex Then
                .maxBidPricesIndex = position
            End If
        End Select
        If .maxBidPricesIndex >= UBound(.bidPrices) Then
            ReDim Preserve .bidPrices(UBound(.bidPrices) + 10) As PriceTableEntry
        End If
        updatePrices Timestamp, _
                    id, _
                    .bidPrices, _
                    .maxBidPricesIndex, _
                    DOMSides.DOMBid, _
                    position, _
                    operation, _
                    Price, _
                    Size, _
                    .prevBid, _
                    .prevAsk, _
                    .DOMEventsRequired = DOMProcessedEvents Or .DOMEventsRequired = DOMBothEvents
        If operation = DOMOperations.DOMDelete Then .maxBidPricesIndex = .maxBidPricesIndex - 1
    End If
    
    If .writeMktDepthToTickfile Then
        .TickfileWriter.writeMarketDepthData Timestamp, _
                                        position, _
                                        "", _
                                        operation, _
                                        side, _
                                        Price, _
                                        Size
    End If
End With
End Sub

Private Sub processTickPrice(ByVal Timestamp As Date, _
                            ByVal TickerID As Long, _
                            ByVal TickType As TWSSocketTickTypes, _
                            ByVal Price As Double, _
                            ByVal Size As Long)
Dim tickListendata As TickerListenData

With mTickerTable(TickerID)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    If Price < (0.5 * .prevTrade) Then
        Exit Sub        ' ignore invalid prices
    End If
    
    Select Case TickType
    Case TICK_BID
        Size = IIf(Size <> 0, Size, .prevBidSize)
        ProcessBid Timestamp, _
                    TickerID, _
                    Price, _
                    Size
        .prevBid = Price
        If Size <> 0 Then .prevBidSize = Size
    
    Case TICK_ASK
        Size = IIf(Size <> 0, Size, .prevAskSize)
        ProcessAsk Timestamp, _
                    TickerID, _
                    Price, _
                    Size
        
        .prevAsk = Price
        If Size <> 0 Then .prevAskSize = Size
    
    Case TICK_LAST
        Size = IIf(Size <> 0, Size, .prevTradeSize)
        ProcessLast Timestamp, _
                    TickerID, _
                    Price, _
                    Size
        .prevTrade = Price
        If Size <> 0 Then .prevTradeSize = Size
    
    Case TICK_HIGH
        ProcessHigh Timestamp, _
                    TickerID, _
                    Price
    Case TICK_LOW
        ProcessLow Timestamp, _
                    TickerID, _
                    Price
    Case TICK_CLOSE
        ProcessClose Timestamp, _
                    TickerID, _
                    Price
    End Select
End With

End Sub

Private Sub processTickSize(ByVal Timestamp As Date, _
                            ByVal TickerID As Long, _
                            ByVal TickType As TWSSocketTickTypes, _
                            ByVal Size As Long)
Dim tickListendata As TickerListenMDData

With mTickerTable(TickerID)
    
    If Not .inUse Then Exit Sub     ' ticker has been stopped
    
    Select Case TickType
    Case TICK_BID_SIZE
        If Size = .prevBidSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        .prevBidSize = Size
        ProcessBid Timestamp, _
                    TickerID, _
                    .prevBid, _
                    Size
    Case TICK_ASK_SIZE
        If Size = .prevAskSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        ProcessAsk Timestamp, _
                    TickerID, _
                    .prevAsk, _
                    Size
    Case TICK_LAST_SIZE
        If Size = .prevTradeSize Then
            ' this happens when TWS has sent a tick price including the
            ' size. It then sends the size again as a ticksize
            Exit Sub
        End If
        
        .prevTradeSize = Size
        ProcessLast Timestamp, _
                    TickerID, _
                    .prevTrade, _
                    Size
    Case TICK_VOLUME
        processVolume Timestamp, _
                    TickerID, _
                    Size
        
    End Select
End With

End Sub

Private Sub processVolume(ByVal Timestamp As Date, _
                            ByVal TickerID As Long, _
                            ByVal Size As Long)
Dim tickListendata As TickerListenData

With mTickerTable(TickerID)

    If Not .inUse Then Exit Sub     ' ticker has been stopped

    If Size = 0 Then Exit Sub
    
    If Not .firstVolumeReceived Then
        If Not .tickerObj Is Nothing Then
            .tickerObj.setVolume Timestamp, _
                        Size
            mTickers.Volume Timestamp, _
                        TickerID, _
                        Size
        End If
        RaiseEvent Volume(Timestamp, _
                        TickerID, _
                        Size)
        
        .accumulatedVolume = Size
        .firstVolumeReceived = True
        
        If mEnableListeners Then
            tickListendata.TickerID = TickerID
            tickListendata.Price = 0
            tickListendata.Size = Size
            mListeners.notify tickListendata, _
                            TradeBuildListenValueTypes.ValueTypeTradeBuildVolume, _
                            Timestamp
        End If
        
    ElseIf Size + .volumeCorrectionIncrement < .accumulatedVolume Then
        If .noVolumeAdjustments Then
            If Not .tickerObj Is Nothing Then
                .tickerObj.setVolume Timestamp, _
                            Size
                mTickers.Volume Timestamp, _
                            TickerID, _
                            Size
            End If
            RaiseEvent Volume(Timestamp, _
                            TickerID, _
                            Size)
            
            If mEnableListeners Then
                tickListendata.TickerID = TickerID
                tickListendata.Price = 0
                tickListendata.Size = Size
                mListeners.notify tickListendata, _
                                TradeBuildListenValueTypes.ValueTypeTradeBuildVolume, _
                                Timestamp
            End If
        Else
            If Not .tickerObj Is Nothing Then
                .tickerObj.setVolume Timestamp, _
                            .accumulatedVolume
                mTickers.Volume Timestamp, _
                            TickerID, _
                            .accumulatedVolume
            End If
            RaiseEvent Volume(Timestamp, _
                                TickerID, _
                                .accumulatedVolume)
            
            If mEnableListeners Then
                tickListendata.TickerID = TickerID
                tickListendata.Price = 0
                tickListendata.Size = .accumulatedVolume
                mListeners.notify tickListendata, _
                                TradeBuildListenValueTypes.ValueTypeTradeBuildVolume, _
                                Timestamp
            End If
            
        End If
        .volumeCorrectionIncrement = .accumulatedVolume - Size
    Else
        If Size + .volumeCorrectionIncrement > .accumulatedVolume Then
            If Not .noImpliedTrades Then
                If Not .tickerObj Is Nothing Then
                    .tickerObj.setTrade Timestamp, _
                                .prevTrade, _
                                Size + .volumeCorrectionIncrement - .accumulatedVolume
                    mTickers.Volume Timestamp, _
                                TickerID, _
                                Size + .volumeCorrectionIncrement - .accumulatedVolume
                End If
                RaiseEvent trade(Timestamp, _
                                TickerID, _
                                .prevTrade, _
                                Size + .volumeCorrectionIncrement - .accumulatedVolume)
                
                If mEnableListeners Then
                    tickListendata.TickerID = TickerID
                    tickListendata.Price = .prevTrade
                    tickListendata.Size = Size + .volumeCorrectionIncrement - .accumulatedVolume
                    mListeners.notify tickListendata, _
                                    TradeBuildListenValueTypes.ValueTypeTradeBuildTrade, _
                                    Timestamp
                End If
                
            End If
            
            .accumulatedVolume = Size + .volumeCorrectionIncrement
            
            If Not .tickerObj Is Nothing Then
                .tickerObj.setVolume Timestamp, _
                                    .accumulatedVolume
                mTickers.Volume Timestamp, _
                                TickerID, _
                                .accumulatedVolume
            End If
            RaiseEvent Volume(Timestamp, _
                            TickerID, _
                            .accumulatedVolume)
            
            If mEnableListeners Then
                tickListendata.TickerID = TickerID
                tickListendata.Price = 0
                tickListendata.Size = .accumulatedVolume
                mListeners.notify tickListendata, _
                                TradeBuildListenValueTypes.ValueTypeTradeBuildVolume, _
                                Timestamp
            End If
            
        Else
            If Not .tickerObj Is Nothing Then
                .tickerObj.setVolume Timestamp, _
                                    .accumulatedVolume
                mTickers.Volume Timestamp, TickerID, .accumulatedVolume
            End If
            RaiseEvent Volume(Timestamp, _
                            TickerID, _
                            .accumulatedVolume)
            
            If mEnableListeners Then
                tickListendata.TickerID = TickerID
                tickListendata.Price = 0
                tickListendata.Size = .accumulatedVolume
                mListeners.notify tickListendata, _
                                TradeBuildListenValueTypes.ValueTypeTradeBuildVolume, _
                                Timestamp
            End If
        End If
    End If
    
    
    ' write the uncorrected size to the tickfile
    If .writeMktDataToTickfile Then _
        .TickfileWriter.writeVolume Timestamp, _
                            Size
End With
End Sub

Private Sub reEstablishMarketData()
Dim i As Long
For i = 0 To UBound(mTickerTable)
    With mTickerTable(i)
        If Not .contractSpecifier Is Nothing Then
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

Private Sub releaseTickerTableEntry(ByVal TickerID As Long)
With mTickerTable(TickerID)
    .inUse = False
    .incarnation = .incarnation + INCARNATION_INCREMENT
    Set .contractSpecifier = Nothing
    .prevBid = 0#
    .prevBidSize = 0&
    .prevAsk = 0#
    .prevAskSize = 0&
    .prevTrade = 0#
    .prevTradeSize = 0&
    .accumulatedVolume = 0&
    .volumeCorrectionIncrement = 0&
    .noImpliedTrades = False
    .noVolumeAdjustments = False
    .firstVolumeReceived = False
    .writeMktDataToTickfile = False
    .receivingMarketDepth = False
    .writeMktDepthToTickfile = False
    .DOMEventsRequired = DOMEvents.DOMNoEvents
    If Not .tickerObj Is Nothing Then
        mTickers.Remove .tickerObj
        Set .tickerObj = Nothing
    End If
    If Not .TickfileWriter Is Nothing Then
        .TickfileWriter.closeOutputFile
        Set .TickfileWriter = Nothing
    End If
    ReDim .askPrices(0) As PriceTableEntry
    .maxAskPricesIndex = 0&
    ReDim .bidPrices(0) As PriceTableEntry
    .maxBidPricesIndex = 0&
End With
End Sub

Private Sub reqContract(ByVal TickerID As Long, _
                        ByVal Contracts As Contracts, _
                        ByVal contractSpecifier As contractSpecifier, _
                        ByVal internal As Boolean)

' don't send another request if there's already one outstanding
If Not mContractDetailsRequest Is Nothing Then
    err.Raise ContractDetailsReqNotAllowed, _
                "TradeBuildAPI::reqContract", _
                "An existing request has not yet completed"
End If

Set mContractDetailsRequest = New ContractDtlsReq
mContractDetailsRequest.TickerID = TickerID
mContractDetailsRequest.Contracts = Contracts
mContractDetailsRequest.contractSpecifier = contractSpecifier.clone
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

With mTickerTable(tickerTableIndex)
    If (.writeMktDataToTickfile And .writeMktDepthToTickfile) Or _
        .DOMEventsRequired <> DOMEvents.DOMNoEvents _
    Then
        If Not .receivingMarketDepth Then
            .receivingMarketDepth = True
            ReDim .askPrices(20) As PriceTableEntry
            ReDim .bidPrices(20) As PriceTableEntry
            reqMarketDepth tickerTableIndex
        End If
    End If
    mTickers.StartTicker .tickerObj
    mOrderSimulator.startMarket tickerTableIndex, .contractSpecifier
End With

End Sub

Private Sub reqMarketDepth(ByVal TickerID As Long)

Const Version = 1

addData REQ_MKT_DEPTH
addData Version
addData TickerID + MARKET_DEPTH_TICKER_BASE
With mTickerTable(TickerID).contractSpecifier
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

Private Sub ResetMarketDepth(ByVal Timestamp As Date, _
                            ByVal reEstablish As Boolean)
Dim i As Long
For i = 0 To UBound(mTickerTable)
    With mTickerTable(i)
        If .receivingMarketDepth Then
            ReDim .askPrices(20) As PriceTableEntry
            ReDim .bidPrices(20) As PriceTableEntry
            .maxAskPricesIndex = 0&
            .maxBidPricesIndex = 0&
            If reEstablish Then reqMarketDepth i
            DOMres Timestamp, _
                    .DOMEventsRequired, _
                    i, _
                    reEstablish
            If .writeMktDepthToTickfile Then
                .TickfileWriter.writeMarketDepthReset Timestamp
            End If
        End If
    End With
Next
End Sub

Private Sub rollbackBuffer()
mIndexIn = mIndexInChk
mInputIndex = mInputIndexChk
'debug.Print "mIndexIn=" & mIndexIn & "  mInputIndex=" & mInputIndex
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

Private Sub updatePrices(ByVal Timestamp As Date, _
                    ByVal TickerID As Long, _
                    prices() As PriceTableEntry, _
                    ByVal maxIndex As Long, _
                    ByVal side As DOMSides, _
                    ByVal position As Long, _
                    ByVal operation As DOMOperations, _
                    ByVal Price As Double, _
                    ByVal Size As Long, _
                    ByVal currentBid As Double, _
                    ByVal currentAsk As Double, _
                    ByVal notifyApp As Boolean)
Dim listendata As TickerListenMDData
Dim i As Long

' NB: we don't update the screen if position is 0, since these are
' passed earlier via the notifyBid and notifyAsk methods. Except for deletions
' (for example if the bid drops, the current entry 0 must be deleted - though
' it could just be updated, TWS doesn't seem to work this way)

Select Case operation
Case DOMOperations.DOMInsert
    For i = maxIndex - 1 To position Step -1
        prices(i + 1) = prices(i)
    Next
    prices(position).Price = Price
    prices(position).Size = Size
    If position <> 0 Then
        If notifyApp Then
            DOMset Timestamp, _
                    DOMProcessedEvents, _
                    TickerID, _
                    side, _
                    Price, _
                    Size
        End If
    End If
Case DOMOperations.DOMUpdate
    If Price <> prices(position).Price And prices(position).Price <> 0 Then
        ' if need be we must blank the size display for the price currently
        ' in this slot. Note that sometimes the market depth updates teomporarily
        ' result in two or more adjacent slots for the same price - therefore we
        ' only blank the size if the adjacent slot(s) are for a different price.
        If position = 0 Then
            If prices(1).Price <> prices(0).Price Then
                If notifyApp Then
                    DOMclear Timestamp, _
                            DOMProcessedEvents, _
                            TickerID, _
                            side, _
                            prices(0).Price
                End If
            End If
        Else
            If prices(position + 1).Price <> prices(position).Price And _
                prices(position - 1).Price <> prices(position).Price _
            Then
                If notifyApp Then
                    DOMclear Timestamp, _
                            DOMProcessedEvents, _
                            TickerID, _
                            side, _
                            prices(position).Price
                End If
            End If
        End If
    End If
    prices(position).Price = Price
    prices(position).Size = Size
    If position <> 0 Then
        If notifyApp Then
            DOMset Timestamp, _
                    DOMProcessedEvents, _
                    TickerID, _
                    side, _
                    Price, _
                    Size
        End If
    End If
Case DOMOperations.DOMDelete
    If (position = 0 And side = DOMSides.DOMBid And prices(position).Price = currentBid) _
        Or _
        (position = 0 And side = DOMSides.DOMAsk And prices(position).Price = currentAsk) _
    Then
        ' we get here when we've already set the bid or ask in the notifyBid/notifyAsk
        ' methods, but the marketdepth updates are deleting some entries prior to
        ' inserting the bid/ask (why it deletes an entry then immediately reinserts
        ' it is a mystery to me, but it does!).
        ' So in this scenario, we don't want to blank the size display for this price
    Else
        ' otherwise we do want to blank the size entry for this price
        If notifyApp Then
            DOMclear Timestamp, _
                    DOMProcessedEvents, _
                    TickerID, _
                    side, _
                    prices(position).Price
        End If
    End If
    For i = position To maxIndex - 1
        prices(i) = prices(i + 1)
    Next
    prices(maxIndex).Price = 0
    prices(maxIndex).Size = 0
End Select
End Sub


