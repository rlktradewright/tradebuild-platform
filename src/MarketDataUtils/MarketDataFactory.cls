VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarketDataFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IMarketDataConsumer
Implements IMarketDepthConsumer
Implements IMarketDataFactory
Implements ITwsConnectionStateListener
Implements StateChangeListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type TickerTableEntry
    Next                    As Long
    InUse                   As Boolean
    TickStreamBuilder       As TickStreamBuilder
    Incarnation             As Long
    ContractDetails         As TwsContractDetails
    ConsumerContract        As IContract
    ConsumerTimeZone        As TimeZone
    ErrorListener           As IMarketDataErrorListener
    DataHasBeenReceived     As Boolean
    PrevBid                 As Double
    PrevBidValid            As Boolean
    PrevBidSize             As Long
    PrevAsk                 As Double
    PrevAskValid            As Boolean
    PrevAskSize             As Long
    PrevTrade               As Double
    PrevTradeValid          As Boolean
    PrevTradeSize           As Long
    PrevOpen                As Double
    PrevHigh                As Double
    PrevLow                 As Double
    PrevClose               As Double
    PrevVolume              As Long
    PrevOpenInterest        As Long
    NumberOfRows            As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MarketDataFactory"

Private Const IncarnationIncrement                  As Long = &H1000&

Private Const MarketDepthIdOffset                   As Long = &H40000000

'@================================================================================
' Member variables
'@================================================================================

Private mTwsAPI                                     As TwsAPI
Attribute mTwsAPI.VB_VarHelpID = -1

Private mMarketDataTable()                          As TickerTableEntry
Private mMarketDataTableFirstFreeIndex              As Long
Private mMarketDataTableNextIndex                   As Long
Private mMarketDataTableMaxUsedIndex                As Long

Private mMarketDepthTable()                         As TickerTableEntry
Private mMarketDepthTableFirstFreeIndex             As Long
Private mMarketDepthTableNextIndex                  As Long
Private mMarketDepthTableMaxUsedIndex               As Long

Private mContractRequester                          As ContractDetailsRequester

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mContractCache                              As TwsContractCache

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mMarketDataTable(3) As TickerTableEntry
mMarketDataTableFirstFreeIndex = NullIndex
mMarketDataTableMaxUsedIndex = NullIndex

ReDim mMarketDepthTable(3) As TickerTableEntry
mMarketDepthTableFirstFreeIndex = NullIndex
mMarketDepthTableMaxUsedIndex = NullIndex

Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' IMarketDataConsumer Interface Members
'@================================================================================

Private Sub IMarketDataConsumer_EndTickSnapshot(ByVal pReqId As Long)
Const ProcName As String = "IMarketDataConsumer_EndTickSnapshot"
On Error GoTo Err

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataConsumer_NotifyError(ByVal pTickerId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IMarketDataConsumer_NotifyError"
On Error GoTo Err

Dim lIndex As Long

If tickerHasBeenStopped(pTickerId) Then Exit Sub

lIndex = getIndexFromTickerId(pTickerId)

Select Case pErrorCode
Case 200    ' security definition not known
    ' make sure we are not already successfully receiving data
    ' for this ticker. Occasionally Tws will send a 200 spuriously
    ' while the data stream is running
    
    If mMarketDataTable(lIndex).DataHasBeenReceived Then Exit Sub
End Select

mMarketDataTable(lIndex).ErrorListener.NotifyMarketDataError lIndex, pErrorCode, pErrorMsg
releaseTableEntry mMarketDataTable, mMarketDataTableFirstFreeIndex, lIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataConsumer_NotifyTickPrice( _
                ByVal pTickerId As Long, _
                ByVal pTickType As TwsTickTypes, _
                ByVal pPrice As Double, _
                ByVal pSize As Long, _
                ByVal pCanAutoExecute As Boolean)
Const ProcName As String = "IMarketDataConsumer_NotifyTickPrice"
On Error GoTo Err

Dim lTick As GenericTick
Dim lIndex As Long

If tickerHasBeenStopped(pTickerId) Then Exit Sub

lIndex = getIndexFromTickerId(pTickerId)

lTick.Price = pPrice
lTick.Size = pSize

With mMarketDataTable(lIndex)

    .DataHasBeenReceived = True

    lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .ConsumerTimeZone)
    
    Select Case pTickType
    Case TwsTickTypes.TwsTickBid
        .PrevBidValid = validPrice(lTick.Price, .PrevBid, "bid", .ConsumerContract)
        If Not .PrevBidValid Then Exit Sub
        lTick.TickType = TickTypeBid
        If lTick.Size = 0 And .PrevBidSize <> MaxLong Then lTick.Size = .PrevBidSize
        .PrevBid = lTick.Price
        If lTick.Size <> 0 Then .PrevBidSize = lTick.Size
    Case TwsTickTypes.TwsTickAsk
        .PrevAskValid = validPrice(lTick.Price, .PrevAsk, "ask", .ConsumerContract)
        If Not .PrevAskValid Then Exit Sub
        lTick.TickType = TickTypeAsk
        If lTick.Size = 0 And .PrevAskSize <> MaxLong Then lTick.Size = .PrevAskSize
        .PrevAsk = lTick.Price
        If lTick.Size <> 0 Then .PrevAskSize = lTick.Size
    Case TwsTickTypes.TwsTickLast
        .PrevTradeValid = validPrice(lTick.Price, .PrevTrade, "trade", .ConsumerContract)
        If Not .PrevTradeValid Then Exit Sub
        lTick.TickType = TickTypeTrade
        If lTick.Size = 0 And .PrevTradeSize <> MaxLong Then lTick.Size = .PrevTradeSize
        .PrevTrade = lTick.Price
        If lTick.Size <> 0 Then .PrevTradeSize = lTick.Size
    Case TwsTickTypes.TwsTickHigh
        If Not validPrice(lTick.Price, .PrevHigh, "high", .ConsumerContract) Then Exit Sub
        lTick.TickType = TickTypeHighPrice
        .PrevHigh = lTick.Price
    Case TwsTickTypes.TwsTickLow
        If Not validPrice(lTick.Price, .PrevLow, "low", .ConsumerContract) Then Exit Sub
        lTick.TickType = TickTypeLowPrice
        .PrevLow = lTick.Price
    Case TwsTickTypes.TwsTickOpen
        If Not validPrice(lTick.Price, .PrevOpen, "open", .ConsumerContract) Then Exit Sub
        lTick.TickType = TickTypeOpenPrice
        .PrevOpen = lTick.Price
    Case TwsTickTypes.TwsTickClose
        If Not validPrice(lTick.Price, .PrevClose, "close", .ConsumerContract) Then Exit Sub
        lTick.TickType = TickTypeClosePrice
        .PrevClose = lTick.Price
    End Select
        
End With

FireTick lIndex, lTick

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataConsumer_NotifyTickEFP(ByVal pTickerId As Long, ByVal pTickType As TwsTickTypes, ByVal pBasisPoints As Double, ByVal pFormattedBasisPoints As String, ByVal pTotalDividends As Double, ByVal pHoldDays As Long, ByVal pFutureExpiry As String, ByVal pDividendImpact As Double, ByVal pDividendsToExpiry As Double)

End Sub

Private Sub IMarketDataConsumer_NotifyTickGeneric(ByVal pTickerId As Long, ByVal pTickType As TwsTickTypes, ByVal pValue As Double)

End Sub

Private Sub IMarketDataConsumer_NotifyTickOptionComputation(ByVal pTickerId As Long, ByVal pTickType As TwsTickTypes, ByVal pImpliedVol As Double, ByVal pDelta As Double, ByVal pOptPrice As Double, ByVal pPvDividend As Double, ByVal pGamma As Double, ByVal pVega As Double, ByVal pTheta As Double, ByVal pUndPrice As Double)

End Sub

Private Sub IMarketDataConsumer_NotifyTickSize( _
                ByVal pTickerId As Long, _
                ByVal pTickType As Long, _
                ByVal pSize As Long)
Const ProcName As String = "IMarketDataConsumer_NotifyTickSize"
On Error GoTo Err

Dim lTick As GenericTick
Dim lIndex As Long

If tickerHasBeenStopped(pTickerId) Then Exit Sub

lIndex = getIndexFromTickerId(pTickerId)
With mMarketDataTable(lIndex)
    
    .DataHasBeenReceived = True
    
    lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .ConsumerTimeZone)
    
    Select Case pTickType
    Case TwsTickTypes.TwsTickBidSize
        If pSize = .PrevBidSize Then
            ' this happens when Tws has sent a tick price including the
            ' Size. It then sends the Size again as a ticksize
            Exit Sub
        End If
        
        If Not .PrevBidValid Then Exit Sub
        lTick.TickType = TickTypeBid
        lTick.Price = .PrevBid
        lTick.Size = pSize
        .PrevBidSize = pSize
    
    Case TwsTickTypes.TwsTickAskSize
        If lTick.Size = .PrevAskSize Then
            ' this happens when Tws has sent a tick price including the
            ' Size. It then sends the Size again as a ticksize
            Exit Sub
        End If
        
        If Not .PrevAskValid Then Exit Sub
        lTick.TickType = TickTypeAsk
        lTick.Price = .PrevAsk
        lTick.Size = pSize
        .PrevAskSize = pSize
    
    Case TwsTickTypes.TwsTickLastSize
        If lTick.Size = .PrevTradeSize Then
            ' this happens when Tws has sent a tick price including the
            ' Size. It then sends the Size again as a ticksize
            Exit Sub
        End If
        
        If Not .PrevTradeValid Then Exit Sub
        lTick.TickType = TickTypeTrade
        lTick.Price = .PrevTrade
        lTick.Size = pSize
        .PrevTradeSize = pSize
    
    Case TwsTickTypes.TwsTickVolume
        lTick.TickType = TickTypeVolume
        lTick.Size = pSize
        .PrevVolume = pSize
    
    Case TwsTickTypes.TwsTickOpenInterest
        lTick.TickType = TickTypeOpenInterest
        lTick.Size = pSize
        .PrevOpenInterest = pSize
    
    End Select

End With

FireTick lIndex, lTick

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataConsumer_NotifyTickString(ByVal pTickerId As Long, ByVal pTickType As TwsTickTypes, ByVal pValue As String)

End Sub

'@================================================================================
' IMarketDataFactory Interface Members
'@================================================================================

Private Function IMarketDataFactory_CreateMarketDataStream(ByVal pContractFuture As IFuture, ByVal pErrorListener As IMarketDataErrorListener, Optional ByVal pStreamId As Long = -1&) As ITickStream
Const ProcName As String = "IMarketDataFactory_CreateMarketDataStream"
On Error GoTo Err

Set IMarketDataFactory_CreateMarketDataStream = CreateMarketDataStream(pContractFuture, pErrorListener, pStreamId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IMarketDataFactory_CreateMarketDepthStream(ByVal pContractFuture As IFuture, ByVal pNumberOfRows As Long, ByVal pErrorListener As IMarketDataErrorListener, Optional ByVal pStreamId As Long = -1) As ITickStream
Const ProcName As String = "IMarketDataFactory_CreateMarketDepthStream"
On Error GoTo Err

Set IMarketDataFactory_CreateMarketDepthStream = CreateMarketDepthStream(pContractFuture, pNumberOfRows, pErrorListener, pStreamId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IMarketDataFactory_MaxStreamId() As Long
IMarketDataFactory_MaxStreamId = IncarnationIncrement
End Property

'@================================================================================
' IMarketDepthConsumer Interface Members
'@================================================================================

Private Sub IMarketDepthConsumer_NotifyError(ByVal pMarketDepthId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IMarketDepthConsumer_NotifyError"
On Error GoTo Err

Dim lIndex As Long

If tickerMarketDepthHasBeenStopped(pMarketDepthId) Then Exit Sub

lIndex = getIndexFromMarketDepthId(pMarketDepthId)
Select Case pErrorCode
Case 316
    ResetMarketDepth pReEstablish:=True
    Exit Sub
Case 317
    ResetMarketDepth pReEstablish:=False
    Exit Sub
End Select

mMarketDepthTable(lIndex).ErrorListener.NotifyMarketDepthError lIndex, pErrorCode, pErrorMsg
releaseTableEntry mMarketDepthTable, mMarketDepthTableFirstFreeIndex, lIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDepthConsumer_NotifyMarketDepth( _
                ByVal pMarketDepthId As Long, _
                ByVal pPosition As Long, _
                ByVal pMarketMaker As String, _
                ByVal pOperation As TwsDOMOperations, _
                ByVal pSide As TwsDOMSides, _
                ByVal pPrice As Double, _
                ByVal pSize As Long)
Const ProcName As String = "IMarketDepthConsumer_NotifyMarketDepth"
On Error GoTo Err

Dim lTick As GenericTick
Dim lIndex As Long

If tickerMarketDepthHasBeenStopped(pMarketDepthId) Then Exit Sub

lIndex = getIndexFromMarketDepthId(pMarketDepthId)

lTick.TickType = TickTypeMarketDepth
lTick.MarketMaker = pMarketMaker
lTick.Operation = pOperation
lTick.Position = pPosition
lTick.Price = pPrice
lTick.Side = pSide
lTick.Size = pSize

With mMarketDataTable(lIndex)

    .DataHasBeenReceived = True

    lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .ConsumerTimeZone)
    
End With

FireTick lIndex, lTick

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDepthConsumer_ResetMarketDepth(ByVal pReEstablish As Boolean)
Const ProcName As String = "IMarketDepthConsumer_ResetMarketDepth"
On Error GoTo Err

ResetMarketDepth pReEstablish

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As ApiConnectionStates, ByVal pMessage As String)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionStateChange"
On Error GoTo Err

Select Case pState
Case ApiConnNotConnected
    
Case ApiConnConnecting

Case ApiConnConnected
    ' restart any tickers that were previously in operation
    ReEstablishMarketData True
    
    ' and re-Request any market depth
    ResetMarketDepth True
Case ApiConnFailed

End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)

End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)

End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

Dim lBuilder As TickStreamBuilder
Set lBuilder = ev.Source

Dim lStreamId As Long
lStreamId = lBuilder.TickStream.StreamId

If ev.State = TickStreamStates.TickStreamStateFinished Then
    If lStreamId <= mMarketDataTableMaxUsedIndex Then
        If mMarketDataTable(lStreamId).TickStreamBuilder Is lBuilder Then
            CancelMarketData lStreamId
            Exit Sub
        End If
    ElseIf lStreamId <= mMarketDepthTableMaxUsedIndex Then
        If mMarketDepthTable(lStreamId).TickStreamBuilder Is lBuilder Then
            CancelMarketDepth lStreamId
            Exit Sub
        End If
    End If
ElseIf ev.State = TickStreamStates.TickStreamStateRunning Then
    If lStreamId <= mMarketDataTableMaxUsedIndex Then
        If isReadyForData(mMarketDataTable, lStreamId) Then reqMarketData lStreamId
    ElseIf lStreamId <= mMarketDepthTableMaxUsedIndex Then
        If isReadyForData(mMarketDepthTable, lStreamId) Then reqMarketDepth lStreamId
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

Dim lStreamId As Long
lStreamId = ev.ContinuationData

If lStreamId < MarketDepthIdOffset Then
    If handleCompletedFuture(ev.Future, mMarketDataTable, lStreamId) Then reqMarketData lStreamId
Else
    If handleCompletedFuture(ev.Future, mMarketDepthTable, lStreamId - MarketDepthIdOffset) Then reqMarketDepth lStreamId
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get CurrentAskPrice(ByVal pTickerId As Long) As Double
Const ProcName As String = "CurrentAskPrice"
On Error GoTo Err

Assert Not tickerHasBeenStopped(pTickerId), "Ticker has been stopped"

CurrentAskPrice = mMarketDataTable(getIndexFromTickerId(pTickerId)).PrevAsk

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentAskSize(ByVal pTickerId As Long) As Long
Const ProcName As String = "CurrentAskSize"
On Error GoTo Err

Assert Not tickerHasBeenStopped(pTickerId), "Ticker has been stopped"

CurrentAskSize = mMarketDataTable(getIndexFromTickerId(pTickerId)).PrevAskSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentBidPrice(ByVal pTickerId As Long) As Double
Const ProcName As String = "CurrentBidPrice"
On Error GoTo Err

Assert Not tickerHasBeenStopped(pTickerId), "Ticker has been stopped"

CurrentBidPrice = mMarketDataTable(getIndexFromTickerId(pTickerId)).PrevBid

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentBidSize(ByVal pTickerId As Long) As Long
CurrentBidSize = mMarketDataTable(pTickerId).PrevBidSize
End Property

Public Property Get CurrentOpenInterest(ByVal pTickerId As Long) As Long
CurrentOpenInterest = mMarketDataTable(pTickerId).PrevOpenInterest
End Property

Public Property Get CurrentTradePrice(ByVal pTickerId As Long) As Double
CurrentTradePrice = mMarketDataTable(pTickerId).PrevTrade
End Property

Public Property Get CurrentTradeSize(ByVal pTickerId As Long) As Long
CurrentTradeSize = mMarketDataTable(pTickerId).PrevTradeSize
End Property

Public Property Get CurrentVolume(ByVal pTickerId As Long) As Long
CurrentVolume = mMarketDataTable(pTickerId).PrevVolume
End Property

Public Property Get MaxTickerId() As Long
MaxTickerId = IncarnationIncrement
End Property

Public Property Get PreviousSessionClosePrice(ByVal pTickerId As Long) As Double
PreviousSessionClosePrice = mMarketDataTable(pTickerId).PrevClose
End Property

Public Property Get SessionHighPrice(ByVal pTickerId As Long) As Double
SessionHighPrice = mMarketDataTable(pTickerId).PrevHigh
End Property

Public Property Get SessionLowPrice(ByVal pTickerId As Long) As Double
SessionLowPrice = mMarketDataTable(pTickerId).PrevLow
End Property

Public Property Get Timestamp(ByVal pTickerId As Long) As Date
Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, mMarketDataTable(pTickerId).ConsumerTimeZone)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub CancelMarketData( _
                ByVal pStreamId As Long)
Const ProcName As String = "CancelMarketData"
On Error GoTo Err

cancelMarketDataEx pStreamId
releaseTableEntry mMarketDataTable, mMarketDataTableFirstFreeIndex, pStreamId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub CancelMarketDepth( _
                ByVal pStreamId As Long)
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err

cancelMarketDepthEx pStreamId
releaseTableEntry mMarketDepthTable, mMarketDepthTableFirstFreeIndex, pStreamId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function CreateMarketDataStream(ByVal pContractFuture As IFuture, ByVal pErrorListener As IMarketDataErrorListener, Optional ByVal pStreamId As Long = -1&) As ITickStream
Const ProcName As String = "CreateMarketDataStream"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture is Nothing"
AssertArgument IsInteger(pStreamId, -1, MaxTickerId), "Invalid stream id"

pStreamId = getTableEntry(mMarketDataTable, mMarketDataTableNextIndex, mMarketDataTableFirstFreeIndex, mMarketDataTableMaxUsedIndex, pStreamId, pContractFuture, pErrorListener)

gLogger.Log "Creating market data stream: stream id=" & pStreamId, ProcName, ModuleName, LogLevelDetail

mFutureWaiter.Add pContractFuture, pStreamId
mFutureWaiter.Add gContractFutureToTwsContractDetails(mContractRequester, pContractFuture, mContractCache), pStreamId
mFutureWaiter.Add mMarketDataTable(pStreamId).TickStreamBuilder.TickStream.ClockFuture, pStreamId

Set CreateMarketDataStream = mMarketDataTable(pStreamId).TickStreamBuilder.TickStream

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateMarketDepthStream(ByVal pContractFuture As IFuture, ByVal pNumberOfRows As Long, ByVal pErrorListener As IMarketDataErrorListener, Optional ByVal pStreamId As Long = -1&) As ITickStream
Const ProcName As String = "CreateMarketDepthStream"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture is Nothing"
AssertArgument IsInteger(pStreamId, -1, MaxTickerId), "Invalid stream id"

pStreamId = getTableEntry(mMarketDepthTable, mMarketDepthTableNextIndex, mMarketDepthTableFirstFreeIndex, mMarketDepthTableMaxUsedIndex, pStreamId, pContractFuture, pErrorListener)
mMarketDepthTable(pStreamId).NumberOfRows = pNumberOfRows

gLogger.Log "Creating market depth stream: stream id=" & pStreamId + MarketDepthIdOffset, ProcName, ModuleName, LogLevelDetail

mFutureWaiter.Add pContractFuture, pStreamId + MarketDepthIdOffset
mFutureWaiter.Add gContractFutureToTwsContractDetails(mContractRequester, pContractFuture, mContractCache), pStreamId + MarketDepthIdOffset
mFutureWaiter.Add mMarketDepthTable(pStreamId).TickStreamBuilder.TickStream.ClockFuture, pStreamId + MarketDepthIdOffset

Set CreateMarketDepthStream = mMarketDepthTable(pStreamId).TickStreamBuilder.TickStream

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pTwsAPI As TwsAPI, _
                ByVal pContractRequester As ContractDetailsRequester, _
                ByVal pContractCache As TwsContractCache)
Set mTwsAPI = pTwsAPI
Set mContractRequester = pContractRequester
Set mContractCache = pContractCache
End Sub

Friend Sub ReEstablishMarketData(ByVal pNewConnection As Boolean)
Const ProcName As String = "ReEstablishMarketData"
On Error GoTo Err

gLog "Re-establishing market data: newConnection = " & CStr(pNewConnection), ModuleName, ProcName

Dim i As Long
For i = 0 To UBound(mMarketDataTable)
    With mMarketDataTable(i)
        If .InUse Then
            If Not pNewConnection Then cancelMarketDataEx i
            .Incarnation = updateIncarnation(i, .Incarnation)
            reqMarketData i
        End If
        clearPrevPrices i
    End With
Next

For i = 0 To UBound(mMarketDepthTable)
    With mMarketDepthTable(i)
        If .InUse Then
            If Not pNewConnection Then cancelMarketDepthEx i
            .Incarnation = updateIncarnation(i, .Incarnation)
            reqMarketDepth i
        End If
    End With
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub cancelMarketDataEx(ByVal pIndex As Long)
Const ProcName As String = "cancelMarketDataEx"
On Error GoTo Err

If mTwsAPI.ConnectionState = TwsConnConnected Then mTwsAPI.CancelMarketData getTickerIdFromIndex(pIndex)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub cancelMarketDepthEx(ByVal pIndex As Long)
Const ProcName As String = "cancelMarketDepthEx"
On Error GoTo Err

If mTwsAPI.ConnectionState = TwsConnConnected Then mTwsAPI.CancelMarketDepth getMarketDepthIdFromIndex(pIndex)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearPrevPrices( _
                ByVal pIndex As Long)
Const ProcName As String = "clearPrevPrices"
On Error GoTo Err

With mMarketDataTable(pIndex)
    .PrevAsk = MaxDouble
    .PrevAskSize = MaxLong
    .PrevBid = MaxDouble
    .PrevBidSize = MaxLong
    .PrevClose = MaxDouble
    .PrevHigh = MaxDouble
    .PrevLow = MaxDouble
    .PrevOpen = MaxDouble
    .PrevTrade = MaxDouble
    .PrevTradeSize = MaxLong
    .PrevVolume = MaxLong
    .PrevOpenInterest = MaxLong
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FireTick( _
                ByVal pStreamId As Long, _
                ByRef pTick As GenericTick)
Const ProcName As String = "fireTick"
On Error GoTo Err

If pTick.TickType = TickTypeMarketDepth Or _
    pTick.TickType = TickTypeMarketDepthReset _
Then
    mMarketDepthTable(pStreamId).TickStreamBuilder.NotifyTick pTick
Else
    mMarketDataTable(pStreamId).TickStreamBuilder.NotifyTick pTick
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getIndexFromTickerId(ByVal pTickerId As Long) As Long
getIndexFromTickerId = pTickerId Mod IncarnationIncrement
End Function

Private Function getIndexFromMarketDepthId(ByVal pMarketDepthId As Long) As Long
getIndexFromMarketDepthId = pMarketDepthId Mod IncarnationIncrement
End Function

Private Function getTheTimezone(ByVal pTimeZoneId As String) As TimeZone
Const ProcName As String = "getTheTimezone"
On Error GoTo Err

Set getTheTimezone = GetTimeZone(gTwsTimezoneNameToStandardTimeZoneName(pTimeZoneId))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getMarketDepthIdFromIndex(ByVal pIndex As Long) As Long
getMarketDepthIdFromIndex = pIndex + mMarketDepthTable(pIndex).Incarnation
End Function

Private Function getTickerIdFromIndex(ByVal pIndex As Long) As Long
getTickerIdFromIndex = pIndex + mMarketDataTable(pIndex).Incarnation
End Function

Private Function getTableEntry( _
                ByRef pTable() As TickerTableEntry, _
                ByRef pNextIndex As Long, _
                ByRef pFirstFreeIndex As Long, _
                ByRef pMaxUsedIndex As Long, _
                ByVal pIndex As Long, _
                ByVal pContractFuture As IFuture, _
                ByVal pErrorListener As IMarketDataErrorListener) As Long
Const ProcName As String = "getTableEntry"
On Error GoTo Err

Dim lIndex As Long
lIndex = pIndex

If lIndex = -1 Then
    If pFirstFreeIndex <> NullIndex Then
        lIndex = pFirstFreeIndex
        pFirstFreeIndex = pTable(pFirstFreeIndex).Next
    Else
        Do
            If pNextIndex > UBound(pTable) Then ReDim Preserve pTable(2 * (UBound(pTable) + 1) - 1) As TickerTableEntry
            lIndex = pNextIndex
            pNextIndex = pNextIndex + 1
        Loop Until Not pTable(lIndex).InUse And pTable(lIndex).Next <> 0
    End If
Else
    If lIndex <= UBound(pTable) Then AssertArgument Not pTable(lIndex).InUse, "Index already in use"
    
    Do While lIndex > UBound(pTable)
        ReDim Preserve pTable(2 * (UBound(pTable) + 1) - 1) As TickerTableEntry
    Loop
End If

If lIndex > pMaxUsedIndex Then pMaxUsedIndex = lIndex
    
With pTable(lIndex)
    .InUse = True
    Set .ErrorListener = pErrorListener
    Set .TickStreamBuilder = CreateTickStreamBuilder(lIndex, pContractFuture, CreateClockFuture(pContractFuture))
    .TickStreamBuilder.AddStateChangeListener Me
End With

clearPrevPrices lIndex

getTableEntry = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function handleCompletedFuture(ByVal pFuture As IFuture, ByRef pTickerTable() As TickerTableEntry, ByVal pStreamId As Long) As Boolean
Const ProcName As String = "handleCompletedFuture"
On Error GoTo Err

If TypeOf pFuture.Value Is IContract Then
    gLogger.Log "Got Contract", ProcName, ModuleName, LogLevelDetail
    Dim lContract As IContract
    Set lContract = pFuture.Value
    Set pTickerTable(pStreamId).ConsumerContract = lContract
    Set pTickerTable(pStreamId).ConsumerTimeZone = GetTimeZone(lContract.TimezoneName)
ElseIf TypeOf pFuture.Value Is TwsContractDetails Then
    gLogger.Log "Got TwsContractDetails", ProcName, ModuleName, LogLevelDetail
    Set pTickerTable(pStreamId).ContractDetails = pFuture.Value
ElseIf TypeOf pFuture.Value Is Clock Then
    gLogger.Log "Got Clock", ProcName, ModuleName, LogLevelDetail
End If

handleCompletedFuture = isReadyForData(pTickerTable, pStreamId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


Private Function isReadyForData(ByRef pTickerTable() As TickerTableEntry, ByVal pStreamId As Long) As Boolean
Const ProcName As String = "isReadyForData"
On Error GoTo Err

With pTickerTable(pStreamId)
    isReadyForData = Not .ContractDetails Is Nothing And _
                    Not .ConsumerContract Is Nothing And _
                    .TickStreamBuilder.TickStream.ClockFuture.IsAvailable And _
                    .TickStreamBuilder.TickStream.State = TickStreamStateRunning
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub releaseTableEntry(ByRef pTable() As TickerTableEntry, ByRef pFirstFreeTableIndex As Long, ByVal pIndex As Long)
Const ProcName As String = "releaseTableEntry"
On Error GoTo Err

With pTable(pIndex)
    .InUse = False
    .Incarnation = updateIncarnation(pIndex, .Incarnation)
    Set .ContractDetails = Nothing
    Set .ContractDetails = Nothing
    Set .ConsumerTimeZone = Nothing
    .DataHasBeenReceived = False
    Set .TickStreamBuilder = Nothing
    .Next = pFirstFreeTableIndex
    pFirstFreeTableIndex = pIndex
End With
clearPrevPrices pIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub reqMarketData( _
                ByVal pIndex As Long)
Const ProcName As String = "reqMarketData"
On Error GoTo Err

If mTwsAPI.ConnectionState <> TwsConnConnected Then Exit Sub
If mMarketDataTable(pIndex).ContractDetails Is Nothing Then Exit Sub

mTwsAPI.RequestMarketData getTickerIdFromIndex(pIndex), mMarketDataTable(pIndex).ContractDetails.Summary, "", False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub reqMarketDepth( _
                ByVal pIndex As Long)
Const ProcName As String = "reqMarketDepth"
On Error GoTo Err

If mTwsAPI.ConnectionState <> TwsConnConnected Then Exit Sub
If mMarketDepthTable(pIndex).ContractDetails Is Nothing Then Exit Sub

mTwsAPI.RequestMarketDepth getMarketDepthIdFromIndex(pIndex), mMarketDepthTable(pIndex).ContractDetails.Summary, mMarketDepthTable(pIndex).NumberOfRows

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ResetMarketDepth(ByVal pReEstablish As Boolean)
Const ProcName As String = "ResetMarketDepth"
On Error GoTo Err

Dim i As Long
Dim lTick As GenericTick
Dim var As Variant


lTick.TickType = TickTypeMarketDepthReset

For i = 0 To UBound(mMarketDepthTable)
    With mMarketDepthTable(i)
        If .InUse Then
            lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .ConsumerTimeZone)
            If pReEstablish Then mTwsAPI.RequestMarketDepth i, .ContractDetails.Summary, .NumberOfRows
            FireTick i, lTick
        End If
    End With
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function tickerHasBeenStopped(ByVal pTickerId As Long) As Boolean
Dim lIndex As Long
Const ProcName As String = "tickerHasBeenStopped"
On Error GoTo Err

lIndex = getIndexFromTickerId(pTickerId)
tickerHasBeenStopped = (pTickerId <> getTickerIdFromIndex(lIndex))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function tickerMarketDepthHasBeenStopped(ByVal pMarketDepthId As Long) As Boolean
Dim lIndex As Long
Const ProcName As String = "tickerMarketDepthHasBeenStopped"
On Error GoTo Err

lIndex = getIndexFromMarketDepthId(pMarketDepthId)
tickerMarketDepthHasBeenStopped = (pMarketDepthId <> getMarketDepthIdFromIndex(lIndex))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function validPrice( _
                ByVal pPrice As Double, _
                ByVal pPrevPrice As Double, _
                ByVal pPriceType As String, _
                ByVal pContract As IContract) As Boolean
Const ProcName As String = "validPrice"

On Error GoTo Err

If (pPrevPrice <> MaxDouble Or (pPrice <> 0# And pPrice <> -1# And pPrice <> -100#)) And _
    IsValidPrice(pPrice, pPrevPrice, pContract.Specifier.Sectype, pContract.TickSize) _
Then
    validPrice = True
Else
    gLogger.Log "Invalid pPrice rejected (" & pPriceType & "): " & _
                    "price=" & pPrice & _
                    "; contract={" & pContract.Specifier.ToString & "}", _
                 ProcName, ModuleName
    Exit Function
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function updateIncarnation(ByVal pIndex As Long, ByVal pIncarnation As Long) As Long
Dim lIncarnation As Long
lIncarnation = pIncarnation + IncarnationIncrement
If pIndex + lIncarnation > mTwsAPI.MaxMarketDataRequestId Or _
    pIndex + lIncarnation > mTwsAPI.MaxMarketDepthRequestId _
Then lIncarnation = 0
updateIncarnation = lIncarnation
End Function



