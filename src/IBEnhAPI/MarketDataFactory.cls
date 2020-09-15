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
Implements IStateChangeListener

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
    IsInUse                 As Boolean
    IsSnapShot              As Boolean
    TickStreamBuilder       As TickStreamBuilder
    Incarnation             As Long
    ContractFuture          As IFuture
    TwsContract         As TwsContract
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

Private mMarketDepthTable()                         As TickerTableEntry
Private mMarketDepthTableFirstFreeIndex             As Long
Private mMarketDepthTableNextIndex                  As Long

Private mContractRequester                          As ContractsTwsRequester
Private mHistoricalDataRequester                    As HistDataRequester

Private WithEvents mTicksFutureWaiter               As FutureWaiter
Attribute mTicksFutureWaiter.VB_VarHelpID = -1
Private WithEvents mDepthFutureWaiter               As FutureWaiter
Attribute mDepthFutureWaiter.VB_VarHelpID = -1

Private mContractCache                              As ContractCache

Private mIsTwsConnectedToIBServers                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mMarketDataTable(3) As TickerTableEntry
mMarketDataTableFirstFreeIndex = NullIndex

ReDim mMarketDepthTable(3) As TickerTableEntry
mMarketDepthTableFirstFreeIndex = NullIndex

Set mTicksFutureWaiter = New FutureWaiter
Set mDepthFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' IMarketDataConsumer Interface Members
'@================================================================================

Private Sub IMarketDataConsumer_EndTickSnapshot(ByVal pTickerId As Long)
Const ProcName As String = "IMarketDataConsumer_EndTickSnapshot"
On Error GoTo Err

If tickerHasBeenStopped(pTickerId) Then Exit Sub

Dim lIndex As Long
lIndex = getIndexFromTickerId(pTickerId)
releaseTableEntry mMarketDataTable, mMarketDataTableFirstFreeIndex, lIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataConsumer_NotifyError(ByVal pTickerId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IMarketDataConsumer_NotifyError"
On Error GoTo Err

If tickerHasBeenStopped(pTickerId) Then Exit Sub

Dim lIndex As Long
lIndex = getIndexFromTickerId(pTickerId)

Select Case pErrorCode
Case 200    ' security definition not known
    ' make sure we are not already successfully receiving data
    ' for this ticker. Occasionally Tws will send a 200 spuriously
    ' while the data stream is running
    
    If mMarketDataTable(lIndex).DataHasBeenReceived Then Exit Sub
    
    ' Check whether we've been notified that connection between TWS and IB servers
    ' has been lost. If so, just ignore this error because we'll retry the
    ' market data request when connection is re-established.
    If Not mIsTwsConnectedToIBServers Then Exit Sub
    
Case 10197  ' No markety data lumps during competing live session
    gLog "Spurious message from IB: " & pErrorMsg, ModuleName, ProcName
    ' ignore this
    Exit Sub
End Select

mMarketDataTable(lIndex).ErrorListener.NotifyMarketDataError lIndex, pErrorCode, pErrorMsg
releaseTableEntry mMarketDataTable, mMarketDataTableFirstFreeIndex, lIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataConsumer_NotifyTickEFP(ByVal pTickerId As Long, ByVal pTickType As TwsTickTypes, ByVal pBasisPoints As Double, ByVal pFormattedBasisPoints As String, ByVal pTotalDividends As Double, ByVal pHoldDays As Long, ByVal pFutureExpiry As String, ByVal pDividendImpact As Double, ByVal pDividendsToExpiry As Double)

End Sub

Private Sub IMarketDataConsumer_NotifyTickGeneric(ByVal pTickerId As Long, ByVal pTickType As TwsTickTypes, ByVal pValue As Double)

End Sub

Private Sub IMarketDataConsumer_NotifyTickOptionComputation( _
                ByVal pTickerId As Long, _
                ByVal pTickType As TwsTickTypes, _
                ByVal pImpliedVolatility As Double, _
                ByVal pDelta As Double, _
                ByVal pOptPrice As Double, _
                ByVal pPvDividend As Double, _
                ByVal pGamma As Double, _
                ByVal pVega As Double, _
                ByVal pTheta As Double, _
                ByVal pUnderlyingPrice As Double)
Const ProcName As String = "IMarketDataConsumer_NotifyTickOptionComputation"
On Error GoTo Err

If tickerHasBeenStopped(pTickerId) Then Exit Sub

Dim lIndex As Long
lIndex = getIndexFromTickerId(pTickerId)

Dim lComputation As OptionComputation

With mMarketDataTable(lIndex)

    .DataHasBeenReceived = True

    lComputation.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .ConsumerTimeZone)
    lComputation.Delta = pDelta
    lComputation.Gamma = pGamma
    lComputation.ImpliedVolatility = pImpliedVolatility
    lComputation.OptionPrice = pOptPrice
    lComputation.PvDividend = pPvDividend
    lComputation.Theta = pTheta
    lComputation.UnderlyingPrice = pUnderlyingPrice
    lComputation.Vega = pVega
    
    Select Case pTickType
    Case TwsTickTypes.TwsTickAskOption
        lComputation.ComputationType = OptionComputationTypeAsk
    Case TwsTickTypes.TwsTickBidOption
        lComputation.ComputationType = OptionComputationTypeBid
    Case TwsTickTypes.TwsTickLastOption
        lComputation.ComputationType = OptionComputationTypeTrade
    Case TwsTickTypes.TwsTickModelOption
        lComputation.ComputationType = OptionComputationTypeModel
    Case Else
        Assert False, "Unexpected option computation type"
    End Select
End With

FireOptionComputation lIndex, lComputation

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataConsumer_NotifyTickPrice( _
                ByVal pTickerId As Long, _
                ByVal pTickType As TwsTickTypes, _
                ByVal pPrice As Double, _
                ByVal pSize As Long, _
                ByRef pAttributes As TwsTickAttributes)
Const ProcName As String = "IMarketDataConsumer_NotifyTickPrice"
On Error GoTo Err

If tickerHasBeenStopped(pTickerId) Then Exit Sub

Dim lIndex As Long
lIndex = getIndexFromTickerId(pTickerId)

Dim lTick As GenericTick
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

Private Sub IMarketDataConsumer_NotifyTickRequestParams( _
                ByVal pTickerId As Long, _
                ByVal pMinTick As Double, _
                ByVal pBboExchange As String, _
                ByVal pSnapshotPermissions As Long)

End Sub

Private Sub IMarketDataConsumer_NotifyTickSize( _
                ByVal pTickerId As Long, _
                ByVal pTickType As Long, _
                ByVal pSize As Long)
Const ProcName As String = "IMarketDataConsumer_NotifyTickSize"
On Error GoTo Err

If tickerHasBeenStopped(pTickerId) Then Exit Sub

Dim lIndex As Long
lIndex = getIndexFromTickerId(pTickerId)

Dim lTick As GenericTick
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
        If pSize = .PrevAskSize Then
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
        If pSize = .PrevTradeSize Then
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

Private Function IMarketDataFactory_CreateSnapshotMarketDataStream(ByVal pContractFuture As TWUtilities40.IFuture, ByVal pErrorListener As TickUtils27.IMarketDataErrorListener, Optional ByVal pStreamId As Long = -1&) As TickUtils27.ITickStream
Const ProcName As String = "IMarketDataFactory_CreateSnapshotMarketDataStream"
On Error GoTo Err

Set IMarketDataFactory_CreateSnapshotMarketDataStream = CreateSnapshotMarketDataStream(pContractFuture, pErrorListener, pStreamId)

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

If tickerMarketDepthHasBeenStopped(pMarketDepthId) Then Exit Sub

Dim lIndex As Long
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

If tickerMarketDepthHasBeenStopped(pMarketDepthId) Then Exit Sub

Dim lIndex As Long
lIndex = getIndexFromMarketDepthId(pMarketDepthId)

Dim lTick As GenericTick
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

' if this is not set, it means that the API connection has never yet been
' fully established, so we shouldn't try to 're-establish' market data or depth
' that has already been requested
Static sHasBeenConnected As Boolean
Static sWasDisconnected As Boolean

Select Case pState
Case ApiConnNotConnected
    If sHasBeenConnected Then sWasDisconnected = True
Case ApiConnConnecting

Case ApiConnConnected
    sHasBeenConnected = True
    
    ' we'll assume that TWS has a server connection: if not,
    ' it will soon tell us
    mIsTwsConnectedToIBServers = True
    
    If sWasDisconnected Then
        ' reset any market depth
        ResetMarketDepth False
        
        ' restart any market data and market depth tickers that were
        ' previously in operation
        reEstablishTickers MarketDataReestablishmentModes.All
    End If
Case ApiConnFailed

End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed( _
                ByVal pSource As Object)
mIsTwsConnectedToIBServers = False
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered( _
                ByVal pSource As Object, _
                ByVal pDataLost As Boolean)
Const ProcName As String = "ITwsConnectionStateListener_NotifyIBServerConnectionRecovered"
On Error GoTo Err

mIsTwsConnectedToIBServers = True

If pDataLost Then
    reEstablishTickers MarketDataReestablishmentModes.Cancel
Else
    reEstablishTickers MarketDataReestablishmentModes.NewTickersOnly
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

Dim lStreamId As Long
Dim lIsMarketData As Boolean
lIsMarketData = isStreamBuilderForMarketData(ev.Source, lStreamId)

If lStreamId = -1 Then
    ' already stopped - nothing to do
ElseIf ev.State = TickStreamStates.TickStreamStateFinished Then
    If lIsMarketData Then
        CancelMarketData lStreamId
    Else
        CancelMarketDepth lStreamId
    End If
ElseIf ev.State = TickStreamStates.TickStreamStateRunning Then
    If lIsMarketData Then
        If isReadyForData(mMarketDataTable, lStreamId) Then reqMarketData lStreamId
    Else
        If isReadyForData(mMarketDepthTable, lStreamId) Then reqMarketDepth lStreamId
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mTicksFutureWaiter Event Handlers
'@================================================================================

Private Sub mTicksFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mTicksFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then Exit Sub

Dim lStreamId As Long
lStreamId = ev.ContinuationData

If Not mMarketDataTable(lStreamId).IsInUse Then Exit Sub
If ev.Future.IsFaulted Then
    mMarketDataTable(lStreamId).ErrorListener.NotifyMarketDataError lStreamId, ev.Future.ErrorNumber, ev.Future.ErrorMessage
    releaseTableEntry mMarketDataTable, mMarketDataTableFirstFreeIndex, lStreamId
ElseIf handleCompletedFuture(ev.Future, mMarketDataTable, lStreamId) Then
    reqMarketData lStreamId
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mDepthFutureWaiter Event Handlers
'@================================================================================

Private Sub mDepthFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mDepthFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then Exit Sub

Dim lStreamId As Long
lStreamId = ev.ContinuationData

lStreamId = lStreamId - MarketDepthIdOffset
If Not mMarketDepthTable(lStreamId).IsInUse Then Exit Sub
If ev.Future.IsFaulted Then
    mMarketDepthTable(lStreamId).ErrorListener.NotifyMarketDataError lStreamId, ev.Future.ErrorNumber, ev.Future.ErrorMessage
    releaseTableEntry mMarketDepthTable, mMarketDepthTableFirstFreeIndex, lStreamId
ElseIf handleCompletedFuture(ev.Future, mMarketDepthTable, lStreamId) Then
    reqMarketDepth lStreamId
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

Set CreateMarketDataStream = doCreateMarketDataStream( _
                                pContractFuture, _
                                False, _
                                pErrorListener, _
                                pStreamId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateMarketDepthStream(ByVal pContractFuture As IFuture, ByVal pNumberOfRows As Long, ByVal pErrorListener As IMarketDataErrorListener, Optional ByVal pStreamId As Long = -1&) As ITickStream
Const ProcName As String = "CreateMarketDepthStream"
On Error GoTo Err

Set CreateMarketDepthStream = doCreateMarketDepthStream( _
                                    pContractFuture, _
                                    pNumberOfRows, _
                                    pErrorListener, _
                                    pStreamId)


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSnapshotMarketDataStream(ByVal pContractFuture As IFuture, ByVal pErrorListener As IMarketDataErrorListener, Optional ByVal pStreamId As Long = -1&) As ITickStream
Const ProcName As String = "CreateSnapshotMarketDataStream"
On Error GoTo Err

Set CreateSnapshotMarketDataStream = doCreateMarketDataStream( _
                                        pContractFuture, _
                                        True, _
                                        pErrorListener, _
                                        pStreamId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pTwsAPI As TwsAPI, _
                ByVal pContractRequester As ContractsTwsRequester, _
                ByVal pHistoricalDataRequester As HistDataRequester, _
                ByVal pContractCache As ContractCache)
Set mTwsAPI = pTwsAPI
Set mContractRequester = pContractRequester
Set mHistoricalDataRequester = pHistoricalDataRequester
Set mContractCache = pContractCache
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub cancelMarketDataEx(ByVal pIndex As Long)
Const ProcName As String = "cancelMarketDataEx"
On Error GoTo Err

If Not (mTwsAPI.ConnectionState = TwsConnConnected And mIsTwsConnectedToIBServers) Then Exit Sub

If mMarketDataTable(pIndex).TwsContract Is Nothing Then
    ' this can happen if the client cancels market data before handleCompletedFuture
    ' has been called
    Exit Sub
End If

If isFTSE100Index(mMarketDataTable(pIndex).TwsContract.Specifier) Then
    mHistoricalDataRequester.CancelRealtimeBars getTickerIdFromIndex(pIndex)
Else
    If Not mMarketDataTable(pIndex).IsSnapShot Then
        mTwsAPI.CancelMarketData getTickerIdFromIndex(pIndex)
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub cancelMarketDepthEx(ByVal pIndex As Long)
Const ProcName As String = "cancelMarketDepthEx"
On Error GoTo Err

If Not (mTwsAPI.ConnectionState = TwsConnConnected And mIsTwsConnectedToIBServers) Then Exit Sub
mTwsAPI.CancelMarketDepth getMarketDepthIdFromIndex(pIndex)

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

Private Function doCreateMarketDataStream( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsSnapshot As Boolean, _
                ByVal pErrorListener As IMarketDataErrorListener, _
                Optional ByVal pStreamId As Long = -1&) As ITickStream
Const ProcName As String = "doCreateMarketDataStream"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture is Nothing"
AssertArgument IsInteger(pStreamId, -1, MaxTickerId), "Invalid stream id"

pStreamId = getTableEntry( _
                mMarketDataTable, _
                mMarketDataTableNextIndex, _
                mMarketDataTableFirstFreeIndex, _
                pStreamId, _
                pContractFuture, _
                pIsSnapshot, _
                pErrorListener)

gLogger.Log "Creating market data stream: stream id=" & pStreamId, ProcName, ModuleName, LogLevelDetail

prepareTickStream mMarketDataTable, pContractFuture, pStreamId, False
Set doCreateMarketDataStream = mMarketDataTable(pStreamId).TickStreamBuilder.TickStream

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function doCreateMarketDepthStream( _
                ByVal pContractFuture As IFuture, _
                ByVal pNumberOfRows As Long, _
                ByVal pErrorListener As IMarketDataErrorListener, _
                Optional ByVal pStreamId As Long = -1&) As ITickStream
Const ProcName As String = "doCreateMarketDepthStream"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture is Nothing"
AssertArgument IsInteger(pStreamId, -1, MaxTickerId), "Invalid stream id"

pStreamId = getTableEntry( _
                mMarketDepthTable, _
                mMarketDepthTableNextIndex, _
                mMarketDepthTableFirstFreeIndex, _
                pStreamId, _
                pContractFuture, _
                False, _
                pErrorListener)
mMarketDepthTable(pStreamId).NumberOfRows = pNumberOfRows

gLogger.Log "Creating market depth stream: stream id=" & pStreamId + MarketDepthIdOffset, ProcName, ModuleName, LogLevelDetail

prepareTickStream mMarketDepthTable, pContractFuture, pStreamId, True
Set doCreateMarketDepthStream = mMarketDepthTable(pStreamId).TickStreamBuilder.TickStream


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub FireOptionComputation( _
                ByVal pStreamId As Long, _
                ByRef pComputation As OptionComputation)
Const ProcName As String = "FireOptionComputation"
On Error GoTo Err

If mMarketDataTable(pStreamId).IsInUse Then
    mMarketDataTable(pStreamId).TickStreamBuilder.NotifyOptionComputation pComputation
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FireTick( _
                ByVal pStreamId As Long, _
                ByRef pTick As GenericTick)
Const ProcName As String = "FireTick"
On Error GoTo Err

If pTick.TickType = TickTypeMarketDepth Or _
    pTick.TickType = TickTypeMarketDepthReset _
Then
    If mMarketDepthTable(pStreamId).IsInUse Then
        mMarketDepthTable(pStreamId).TickStreamBuilder.NotifyTick pTick
    End If
Else
    If mMarketDataTable(pStreamId).IsInUse Then
        mMarketDataTable(pStreamId).TickStreamBuilder.NotifyTick pTick
    End If
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
                ByVal pIndex As Long, _
                ByVal pContractFuture As IFuture, _
                ByVal pIsSnapshot As Boolean, _
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
        Loop Until Not pTable(lIndex).IsInUse And pTable(lIndex).Next <> 0
    End If
Else
    If lIndex <= UBound(pTable) Then AssertArgument Not pTable(lIndex).IsInUse, "Index already in use"
    
    Do While lIndex > UBound(pTable)
        ReDim Preserve pTable(2 * (UBound(pTable) + 1) - 1) As TickerTableEntry
    Loop
End If

With pTable(lIndex)
    .IsInUse = True
    .IsSnapShot = pIsSnapshot
    Set .ErrorListener = pErrorListener
    Set .ContractFuture = pContractFuture
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
ElseIf TypeOf pFuture.Value Is TwsContract Then
    gLogger.Log "Got TwsContract", ProcName, ModuleName, LogLevelDetail
    Set pTickerTable(pStreamId).TwsContract = pFuture.Value
ElseIf TypeOf pFuture.Value Is Clock Then
    gLogger.Log "Got Clock", ProcName, ModuleName, LogLevelDetail
End If

handleCompletedFuture = isReadyForData(pTickerTable, pStreamId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


Private Function incrementIncarnation(ByVal pIndex As Long, ByVal pIncarnation As Long) As Long
Dim lIncarnation As Long
lIncarnation = pIncarnation + IncarnationIncrement
If pIndex + lIncarnation > mTwsAPI.MaxMarketDataRequestId Or _
    pIndex + lIncarnation > mTwsAPI.MaxMarketDepthRequestId _
Then lIncarnation = 0
incrementIncarnation = lIncarnation
End Function

Private Function isFTSE100Index(ByVal pContract As TwsContractSpecifier) As Boolean
isFTSE100Index = pContract.CurrencyCode = "GBP" And _
                pContract.Exchange = "ICEEU" And _
                pContract.LocalSymbol = "Z" And _
                pContract.SecType = TwsSecTypeIndex And _
                pContract.Symbol = "Z"
End Function

Private Function isReadyForData(ByRef pTickerTable() As TickerTableEntry, ByVal pStreamId As Long) As Boolean
Const ProcName As String = "isReadyForData"
On Error GoTo Err

With pTickerTable(pStreamId)
    isReadyForData = Not .TwsContract Is Nothing And _
                    Not .ConsumerContract Is Nothing And _
                    .TickStreamBuilder.TickStream.ClockFuture.IsAvailable And _
                    .TickStreamBuilder.TickStream.State = TickStreamStateRunning
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isStreamBuilderForMarketData( _
                ByVal pBuilder As TickStreamBuilder, _
                ByRef pStreamId As Long) As Boolean
Const ProcName As String = "isStreamBuilderForMarketData"
On Error GoTo Err

pStreamId = pBuilder.TickStream.StreamId

If pStreamId <= UBound(mMarketDataTable) Then
    If mMarketDataTable(pStreamId).TickStreamBuilder Is pBuilder Then
        isStreamBuilderForMarketData = True
        Exit Function
    End If
End If

If pStreamId <= UBound(mMarketDepthTable) Then
    If mMarketDepthTable(pStreamId).TickStreamBuilder Is pBuilder Then
        isStreamBuilderForMarketData = False
        Exit Function
    End If
End If

' the stream has already been stopped (eg due to an error being notified
' via IMarketDataConsumer or IMarketDepthConsumer)
isStreamBuilderForMarketData = False
pStreamId = -1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub prepareTickStream( _
                ByRef pTickers() As TickerTableEntry, _
                ByVal pContractFuture As IFuture, _
                ByVal pStreamId As Long, _
                ByVal pIsMarketDepth As Boolean)
Dim lContinuationData As Long
lContinuationData = pStreamId + IIf(pIsMarketDepth, MarketDepthIdOffset, 0)
Dim lFutureWaiter As FutureWaiter
If pIsMarketDepth Then
    Set lFutureWaiter = mDepthFutureWaiter
Else
    Set lFutureWaiter = mTicksFutureWaiter
End If
lFutureWaiter.Add pContractFuture, lContinuationData
lFutureWaiter.Add gContractFutureToTwsContractFuture(mContractRequester, pContractFuture, mContractCache), lContinuationData
lFutureWaiter.Add pTickers(pStreamId).TickStreamBuilder.TickStream.ClockFuture, lContinuationData
End Sub

Private Sub reEstablishData( _
                ByRef pTickers() As TickerTableEntry, _
                ByVal pIsMarketDepth As Boolean, _
                ByVal pMode As MarketDataReestablishmentModes)
Const ProcName As String = "reEstablishData"
On Error GoTo Err

Dim s As String: s = IIf(pMode = MarketDataReestablishmentModes.All, "All", IIf(pMode = MarketDataReestablishmentModes.Cancel, "Cancel", "New tickers only"))
gLog "Re-establishing market " & IIf(pIsMarketDepth, "depth ", "") & "data: mode = " & s, ModuleName, ProcName

Dim i As Long
For i = 0 To UBound(pTickers)
    With pTickers(i)
        If .IsInUse Then
            ' in the following, incrementing the ticker id will avoid potential
            ' problems with duplicate or unknown ticker id errors
            If pMode = MarketDataReestablishmentModes.NewTickersOnly And Not .DataHasBeenReceived Then
                .Incarnation = incrementIncarnation(i, .Incarnation)
                reqData i, pIsMarketDepth
            ElseIf pMode = MarketDataReestablishmentModes.Cancel Then
                cancelMarketDataEx i
                .Incarnation = incrementIncarnation(i, .Incarnation)
                reqData i, pIsMarketDepth
            ElseIf pMode = MarketDataReestablishmentModes.All Then
                .Incarnation = incrementIncarnation(i, .Incarnation)
                reqData i, pIsMarketDepth
            End If
        End If
        clearPrevPrices i
    End With
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub reEstablishTickers(ByVal pMode As MarketDataReestablishmentModes)
Const ProcName As String = "reEstablishTickers"
On Error GoTo Err

reEstablishData mMarketDataTable, False, pMode
reEstablishData mMarketDepthTable, True, pMode

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub releaseTableEntry(ByRef pTable() As TickerTableEntry, ByRef pFirstFreeTableIndex As Long, ByVal pIndex As Long)
Const ProcName As String = "releaseTableEntry"
On Error GoTo Err

With pTable(pIndex)
    .IsInUse = False
    .Incarnation = incrementIncarnation(pIndex, .Incarnation)
    .IsSnapShot = False
    Set .TwsContract = Nothing
    Set .ContractFuture = Nothing
    Set .ConsumerContract = Nothing
    Set .ConsumerTimeZone = Nothing
    .DataHasBeenReceived = False
    .TickStreamBuilder.Finish
    ' don't do this as it causes a failure if this is during a FireTick call
    'Set .TickStreamBuilder = Nothing
    
    .Next = pFirstFreeTableIndex
    pFirstFreeTableIndex = pIndex
End With
clearPrevPrices pIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub reqData(ByVal pIndex As Long, ByVal pIsMarketDepth As Boolean)
Const ProcName As String = "reqData"
On Error GoTo Err

If pIsMarketDepth Then
    reqMarketDepth pIndex
Else
    reqMarketData pIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub reqMarketData( _
                ByVal pIndex As Long)
Const ProcName As String = "reqMarketData"
On Error GoTo Err

If Not (mTwsAPI.ConnectionState = TwsConnConnected And mIsTwsConnectedToIBServers) Then Exit Sub
If mMarketDataTable(pIndex).TwsContract Is Nothing Then
    Assert Not mMarketDataTable(pIndex).ContractFuture Is Nothing, "Contract future is not available"
    prepareTickStream mMarketDataTable, mMarketDataTable(pIndex).ContractFuture, pIndex, False
ElseIf isFTSE100Index(mMarketDataTable(pIndex).TwsContract.Specifier) Then
    ' because IB no longer offer realtime data for the FTSE 100 index, but do still
    ' allow RequestRealTimeBars, we call that and use the returned bars to deduce
    ' the relevant ticks
    Dim lRTBRequester As New RealtimeBarRequester
    lRTBRequester.Initialise Me
    mHistoricalDataRequester.RequestRealtimeBars _
                                    getTickerIdFromIndex(pIndex), _
                                    mMarketDataTable(pIndex).TwsContract.Specifier, _
                                    lRTBRequester
Else
    mTwsAPI.RequestMarketData getTickerIdFromIndex(pIndex), _
                                mMarketDataTable(pIndex).TwsContract.Specifier, _
                                "", _
                                mMarketDataTable(pIndex).IsSnapShot, _
                                False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub reqMarketDepth( _
                ByVal pIndex As Long)
Const ProcName As String = "reqMarketDepth"
On Error GoTo Err

If Not (mTwsAPI.ConnectionState = TwsConnConnected And mIsTwsConnectedToIBServers) Then Exit Sub
If mMarketDepthTable(pIndex).TwsContract Is Nothing Then
    Assert Not mMarketDepthTable(pIndex).ContractFuture Is Nothing, "Contract future is not available"
    prepareTickStream mMarketDepthTable, mMarketDepthTable(pIndex).ContractFuture, pIndex, True
Else
    mTwsAPI.RequestMarketDepth getMarketDepthIdFromIndex(pIndex), mMarketDepthTable(pIndex).TwsContract.Specifier, mMarketDepthTable(pIndex).NumberOfRows
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ResetMarketDepth(ByVal pReEstablish As Boolean)
Const ProcName As String = "ResetMarketDepth"
On Error GoTo Err

Dim lTick As GenericTick
lTick.TickType = TickTypeMarketDepthReset

Dim i As Long
For i = 0 To UBound(mMarketDepthTable)
    With mMarketDepthTable(i)
        If .IsInUse Then
            lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .ConsumerTimeZone)
            If pReEstablish Then mTwsAPI.RequestMarketDepth i, .TwsContract.Specifier, .NumberOfRows
            FireTick i, lTick
        End If
    End With
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function tickerHasBeenStopped(ByVal pTickerId As Long) As Boolean
Const ProcName As String = "tickerHasBeenStopped"
On Error GoTo Err

Dim lIndex As Long
lIndex = getIndexFromTickerId(pTickerId)
tickerHasBeenStopped = (pTickerId <> getTickerIdFromIndex(lIndex))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function tickerMarketDepthHasBeenStopped(ByVal pMarketDepthId As Long) As Boolean
Const ProcName As String = "tickerMarketDepthHasBeenStopped"
On Error GoTo Err

Dim lIndex As Long
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
    IsValidPrice(pPrice, pPrevPrice, pContract.Specifier.SecType, pContract.TickSize) _
Then
    validPrice = True
Else
    gLogger.Log "Invalid price rejected (" & pPriceType & "): " & _
                    "price=" & pPrice & _
                    "; contract={" & pContract.Specifier.ToString & "}", _
                 ProcName, ModuleName
    Exit Function
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



