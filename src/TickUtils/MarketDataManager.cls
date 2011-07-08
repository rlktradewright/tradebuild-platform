VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarketDataRequester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Implements ITwsConnectionStateListener

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
    Incarnation             As Long
    ContractDetails         As TwsContractDetails
    TimeZone                As TimeZone
    Next                    As Long
    NumberOfRows            As Long
    DataHasBeenReceived     As Boolean
    Subscriptions           As DataSourceSubscriptions
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
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MarketDataRequester"

Private Const IncarnationIncrement                  As Long = &H1000&

'@================================================================================
' Member variables
'@================================================================================

Private mTwsAPI                                     As TwsAPI
Attribute mTwsAPI.VB_VarHelpID = -1

Private mTickerTable()                              As TickerTableEntry
Private mNextTickerTableIndex                       As Long
Private mFirstFreeTickerTableIndex                  As Long
Private mLastFreeTickerTableIndex                   As Long

Private mTickerMap                                  As ContractTickerMap

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mTickerTable(3) As TickerTableEntry
mFirstFreeTickerTableIndex = NullIndex
mLastFreeTickerTableIndex = NullIndex
Set mTickerMap = New ContractTickerMap
End Sub

'@================================================================================
' IMarketDataConsumer Interface Members
'@================================================================================

Private Sub IMarketDataConsumer_EndTickSnapshot(ByVal pReqId As Long)
Const ProcName As String = "IMarketDataConsumer_EndTickSnapshot"
On Error GoTo Err

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
    
    If mTickerTable(lIndex).DataHasBeenReceived Then Exit Sub
End Select

mTickerTable(lIndex).Subscriptions.FireMarketDataError pErrorCode, pErrorMsg
mTickerTable(lIndex).Subscriptions.RemoveMarketDataSubscriptions

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

lIndex = getIndexFromTickerId(pTickerId)

lTick.Price = pPrice
lTick.Size = pSize

With mTickerTable(lIndex)

    .DataHasBeenReceived = True

    lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .TimeZone)
    
    Select Case pTickType
    Case TwsTickTypes.TwsTickBid
        .PrevBidValid = validPrice(lTick.Price, .PrevBid, "bid", .ContractDetails)
        If Not .PrevBidValid Then Exit Sub
        lTick.TickType = TickTypeBid
        If lTick.Size = 0 Then lTick.Size = .PrevBidSize
        .PrevBid = lTick.Price
        If lTick.Size <> 0 Then .PrevBidSize = lTick.Size
    Case TwsTickTypes.TwsTickAsk
        .PrevAskValid = validPrice(lTick.Price, .PrevAsk, "ask", .ContractDetails)
        If Not .PrevAskValid Then Exit Sub
        lTick.TickType = TickTypeAsk
        If lTick.Size = 0 Then lTick.Size = .PrevAskSize
        .PrevAsk = lTick.Price
        If lTick.Size <> 0 Then .PrevAskSize = lTick.Size
    Case TwsTickTypes.TwsTickLast
        .PrevTradeValid = validPrice(lTick.Price, .PrevTrade, "trade", .ContractDetails)
        If Not .PrevTradeValid Then Exit Sub
        lTick.TickType = TickTypeTrade
        If lTick.Size = 0 Then lTick.Size = .PrevTradeSize
        .PrevTrade = lTick.Price
        If lTick.Size <> 0 Then .PrevTradeSize = lTick.Size
    Case TwsTickTypes.TwsTickHigh
        If Not validPrice(lTick.Price, .PrevHigh, "high", .ContractDetails) Then Exit Sub
        lTick.TickType = TickTypeHighPrice
        .PrevHigh = lTick.Price
    Case TwsTickTypes.TwsTickLow
        If Not validPrice(lTick.Price, .PrevLow, "low", .ContractDetails) Then Exit Sub
        lTick.TickType = TickTypeLowPrice
        .PrevLow = lTick.Price
    Case TwsTickTypes.TwsTickOpen
        If Not validPrice(lTick.Price, .PrevOpen, "open", .ContractDetails) Then Exit Sub
        lTick.TickType = TickTypeOpenPrice
        .PrevOpen = lTick.Price
    Case TwsTickTypes.TwsTickClose
        If Not validPrice(lTick.Price, .PrevClose, "close", .ContractDetails) Then Exit Sub
        lTick.TickType = TickTypeClosePrice
        .PrevClose = lTick.Price
    End Select
        
    .Subscriptions.FireTick lTick
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
With mTickerTable(lIndex)
    
    .DataHasBeenReceived = True
    
    lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .TimeZone)
    
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

    .Subscriptions.FireTick lTick
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IMarketDataConsumer_NotifyTickString(ByVal pTickerId As Long, ByVal pTickType As TwsTickTypes, ByVal pValue As String)

End Sub

'@================================================================================
' IMarketDepthConsumer Interface Members
'@================================================================================

Private Sub IMarketDepthConsumer_NotifyError(ByVal pTickerId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "IMarketDepthConsumer_NotifyError"
On Error GoTo Err

Dim lIndex As Long

If tickerHasBeenStopped(pTickerId) Then Exit Sub

lIndex = getIndexFromTickerId(pTickerId)
Select Case pErrorCode
Case 316
    ResetMarketDepth pReEstablish:=True
    Exit Sub
Case 317
    ResetMarketDepth pReEstablish:=False
    Exit Sub
End Select

mTickerTable(lIndex).Subscriptions.FireMarketDepthError pErrorCode, pErrorMsg
mTickerTable(lIndex).Subscriptions.RemoveMarketDepthSubscriptions

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IMarketDepthConsumer_NotifyMarketDepth( _
                ByVal pTickerId As Long, _
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

If tickerHasBeenStopped(pTickerId) Then Exit Sub

lIndex = getIndexFromTickerId(pTickerId)

lTick.TickType = TickTypeMarketDepth
lTick.MarketMaker = pMarketMaker
lTick.Operation = pOperation
lTick.Position = pPosition
lTick.Price = pPrice
lTick.Size = pSize

With mTickerTable(lIndex)

    .DataHasBeenReceived = True

    lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .TimeZone)
    
    .Subscriptions.FireTick lIndex, lTick
End With
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IMarketDepthConsumer_ResetMarketDepth(ByVal pReEstablish As Boolean)
Const ProcName As String = "IMarketDepthConsumer_ResetMarketDepth"
On Error GoTo Err

If tickerHasBeenStopped(pTickerId) Then Exit Sub

ResetMarketDepth pReEstablish

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As ApiConnectionStates, ByVal pMessage As String)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionStateChange"
On Error GoTo Err

Select Case pState
Case ApiConnNotConnected
    releaseAllTickerTableEntries
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)

End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)

End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get CurrentAskPrice(ByVal pTickerId As Long) As Double
Const ProcName As String = "CurrentAskPrice"
On Error GoTo Err

If tickerHasBeenStopped(pTickerId) Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Ticker has been stopped"

CurrentAskPrice = mTickerTable(getIndexFromTickerId(pTickerId)).PrevAsk

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get CurrentAskSize(ByVal pTickerId As Long) As Long
Const ProcName As String = "CurrentAskSize"
On Error GoTo Err

If tickerHasBeenStopped(pTickerId) Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Ticker has been stopped"

CurrentAskSize = mTickerTable(getIndexFromTickerId(pTickerId)).PrevAskSize

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get CurrentBidPrice(ByVal pTickerId As Long) As Double
Const ProcName As String = "CurrentBidPrice"
On Error GoTo Err

If tickerHasBeenStopped(pTickerId) Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Ticker has been stopped"

CurrentBidPrice = mTickerTable(getIndexFromTickerId(pTickerId)).PrevBid

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get CurrentBidSize(ByVal pTickerId As Long) As Long
CurrentBidSize = mTickerTable(pTickerId).PrevBidSize
End Property

Friend Property Get CurrentOpenInterest(ByVal pTickerId As Long) As Long
CurrentOpenInterest = mTickerTable(pTickerId).PrevOpenInterest
End Property

Friend Property Get CurrentTradePrice(ByVal pTickerId As Long) As Double
CurrentTradePrice = mTickerTable(pTickerId).PrevTrade
End Property

Friend Property Get CurrentTradeSize(ByVal pTickerId As Long) As Long
CurrentTradeSize = mTickerTable(pTickerId).PrevTradeSize
End Property

Friend Property Get CurrentVolume(ByVal pTickerId As Long) As Long
CurrentVolume = mTickerTable(pTickerId).PrevVolume
End Property

Friend Property Get PreviousSessionClosePrice(ByVal pTickerId As Long) As Double
PreviousSessionClosePrice = mTickerTable(pTickerId).PrevClose
End Property

Friend Property Get SessionHighPrice(ByVal pTickerId As Long) As Double
SessionHighPrice = mTickerTable(pTickerId).PrevHigh
End Property

Friend Property Get SessionLowPrice(ByVal pTickerId As Long) As Double
SessionLowPrice = mTickerTable(pTickerId).PrevLow
End Property

Friend Property Get Timestamp(ByVal pTickerId As Long) As Date
Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, mTickerTable(pTickerId).TimeZone)
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddMarketDataListener( _
                    ByVal pTickerId As Long, _
                    ByVal pDataSource As MarketDataSource, _
                    ByVal pListener As IGenericTickListener)
Const ProcName As String = "AddMarketDataListener"
On Error GoTo Err

notifyCurrentMarketDataValues pTickerId, pListener, pDataSource
mTickerTable(pTickerId).Subscriptions.AddMarketDataSubscription pDataSource, pListener

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub CancelMarketDepth( _
                ByVal pTickerId As Long, _
                ByVal pDataSource As MarketDataSource, _
                ByVal pListener As IGenericTickListener)
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err

If tickerHasBeenStopped(pTickerId) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid ticker id " & pTickerId

mTickerTable(pTickerId).Subscriptions.RemoveMarketDepthSubscription pDataSource, pListener
If mTickerTable(pTickerId).Subscriptions.NumberOfMarketDepthSubscriptions = 0 Then
    cancelMarketDepthEx pTickerId
    If mTickerTable(pTickerId).Subscriptions.NumberOfMarketDataSubscriptions = 0 Then
        mTickerMap.Remove mTickerTable(pTickerId).ContractDetails.Summary
        releaseTickerTableEntry pTickerId
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTwsAPI As TwsAPI)
Set mTwsAPI = pTwsAPI
End Sub

Friend Sub ReEstablishMarketData(ByVal pNewConnection As Boolean)
Const ProcName As String = "ReEstablishMarketData"
Dim i As Long
Dim var As Variant

On Error GoTo Err

logMessage "Re-establishing market data: newConnection = " & CStr(pNewConnection), ProcName
For Each var In mTickerMap
    i = CLng(var)
    With mTickerTable(i)
        If Not pNewConnection Then
            If .Subscriptions.NumberOfMarketDataSubscriptions <> 0 Then cancelMarketDataEx i
            If .Subscriptions.NumberOfMarketDepthSubscriptions <> 0 Then cancelMarketDepthEx i
        End If
        
        .Incarnation = updateIncarnation(i, .Incarnation)
        
        If .Subscriptions.NumberOfMarketDataSubscriptions <> 0 Then reqMarketData i
        If .Subscriptions.NumberOfMarketDepthSubscriptions <> 0 Then reqMarketDepth i
    End With
    clearPrevPrices i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RemoveMarketDataListener( _
                ByVal pTickerId As Long, _
                ByVal pDataSource As MarketDataSource, _
                ByVal pListener As IGenericTickListener)
Const ProcName As String = "RemoveMarketDataListener"
On Error GoTo Err

mTickerTable(pTickerId).Subscriptions.RemoveMarketDataSubscription pDataSource, pListener
If mTickerTable(pTickerId).Subscriptions.NumberOfMarketDataSubscriptions = 0 Then
    cancelMarketDataEx pTickerId
    If mTickerTable(pTickerId).Subscriptions.NumberOfMarketDepthSubscriptions = 0 Then
        mTickerMap.Remove mTickerTable(pTickerId).ContractDetails.Summary
        releaseTickerTableEntry pTickerId
    End If
End If

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Function RequestMarketData( _
                ByVal pContractDetails As TwsContractDetails, _
                ByVal pDataSource As MarketDataSource, _
                ByVal pListeners As Enumerable) As Long
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

Dim lIndex As Long
Dim lListener As IGenericTickListener
Dim en As Enumerator

If mTickerMap.Contains(pContractDetails.Summary) Then
    lIndex = mTickerMap.Item(pContractDetails.Summary)
Else
    lIndex = allocateTickerTableEntry
    Set mTickerTable(lIndex).ContractDetails = pContractDetails
    Set mTickerTable(lIndex).TimeZone = getTheTimezone(pContractDetails.TimeZoneId)
    
    mTickerMap.Add pContractDetails.Summary, lIndex
End If

If mTickerTable(lIndex).Subscriptions.NumberOfMarketDataSubscriptions = 0 Then
    reqMarketData lIndex
Else
    Set en = pListeners.Enumerator
    Do While en.MoveNext
        Set lListener = en.Current
        notifyCurrentMarketDataValues lIndex, lListener, pDataSource
    Loop
End If

Set en = pListeners.Enumerator
Do While en.MoveNext
    Set lListener = en.Current
    mTickerTable(lIndex).Subscriptions.AddMarketDataSubscription pDataSource, lListener
Loop

RequestMarketData = getTickerIdFromIndex(lIndex)
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function RequestMarketDepth( _
                ByVal pContractDetails As TwsContractDetails, _
                ByVal pDataSource As MarketDataSource, _
                ByVal pListeners As Enumerable, _
                ByVal pNumberOfRows As Long) As Long
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

Dim lIndex As Long
Dim lListener As IGenericTickListener
Dim en As Enumerator

If mTickerMap.Contains(pContractDetails.Summary) Then
    lIndex = mTickerMap.Item(pContractDetails.Summary)
Else
    lIndex = allocateTickerTableEntry
    Set mTickerTable(lIndex).ContractDetails = pContractDetails
    Set mTickerTable(lIndex).TimeZone = getTheTimezone(pContractDetails.TimeZoneId)
    
    mTickerMap.Add pContractDetails.Summary, lIndex
End If

If pNumberOfRows > mTickerTable(lIndex).NumberOfRows Then mTickerTable(lIndex).NumberOfRows = pNumberOfRows

If mTickerTable(lIndex).Subscriptions.NumberOfMarketDepthSubscriptions = 0 Then reqMarketDepth lIndex

Set en = pListeners.Enumerator
Do While en.MoveNext
    Set lListener = en.Current
    mTickerTable(lIndex).Subscriptions.AddMarketDepthSubscription pDataSource, lListener
Loop

RequestMarketDepth = getTickerIdFromIndex(lIndex)
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocateTickerTableEntry() As Long
Const ProcName As String = "allocateTickerTableEntry"
On Error GoTo Err

If mNextTickerTableIndex <= UBound(mTickerTable) Then
    allocateTickerTableEntry = mNextTickerTableIndex
    mNextTickerTableIndex = mNextTickerTableIndex + 1
ElseIf mFirstFreeTickerTableIndex <> NullIndex Then
    allocateTickerTableEntry = mFirstFreeTickerTableIndex
    mFirstFreeTickerTableIndex = mTickerTable(mFirstFreeTickerTableIndex).Next
    If mFirstFreeTickerTableIndex = NullIndex Then mLastFreeTickerTableIndex = NullIndex
Else
    Dim lNewBound As Long
    lNewBound = 2 * (UBound(mTickerTable) + 1) - 1
    If lNewBound >= IncarnationIncrement Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Cannot allocate ticker table entry - table is full"
    ReDim Preserve mTickerTable(2 * (UBound(mTickerTable) + 1) - 1) As TickerTableEntry
    allocateTickerTableEntry = mNextTickerTableIndex
    mNextTickerTableIndex = mNextTickerTableIndex + 1
End If

' NB: don't use 'With' in the following as it can lead to an Error 10 table is temporarily locked
mTickerTable(allocateTickerTableEntry).Next = NullIndex
Set mTickerTable(allocateTickerTableEntry).Subscriptions = New DataSourceSubscriptions
mTickerTable(allocateTickerTableEntry).PrevAsk = MaxDouble
mTickerTable(allocateTickerTableEntry).PrevBid = MaxDouble
mTickerTable(allocateTickerTableEntry).PrevClose = MaxDouble
mTickerTable(allocateTickerTableEntry).PrevHigh = MaxDouble
mTickerTable(allocateTickerTableEntry).PrevLow = MaxDouble
mTickerTable(allocateTickerTableEntry).PrevOpen = MaxDouble
mTickerTable(allocateTickerTableEntry).PrevTrade = MaxDouble
mTickerTable(allocateTickerTableEntry).PrevVolume = MaxLong
mTickerTable(allocateTickerTableEntry).PrevOpenInterest = MaxLong

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub cancelMarketDataEx(ByVal pTickerId As Long)
Const ProcName As String = "cancelMarketDataEx"
On Error GoTo Err

If mTwsAPI.ConnectionState = TwsConnConnected Then mTwsAPI.CancelMarketData pTickerId + mTickerTable(pTickerId).Incarnation

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub cancelMarketDepthEx(ByVal pTickerId As Long)
Const ProcName As String = "cancelMarketDepthEx"
On Error GoTo Err

If mTwsAPI.ConnectionState = TwsConnConnected Then mTwsAPI.CancelMarketDepth pTickerId + mTickerTable(pTickerId).Incarnation

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub clearPrevPrices( _
                ByVal index As Long)
Const ProcName As String = "clearPrevPrices"
On Error GoTo Err

With mTickerTable(index)
    .PrevBid = 0#
    .PrevBidSize = 0&
    .PrevAsk = 0#
    .PrevAskSize = 0&
    .PrevTrade = 0#
    .PrevTradeSize = 0&
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function getIndexFromTickerId(ByVal pTickerId As Long) As Long
getIndexFromTickerId = pTickerId Mod IncarnationIncrement
End Function

Private Function getTheTimezone(ByVal pTimeZoneId As String) As TimeZone
Const ProcName As String = "getTheTimezone"
On Error GoTo Err

Set getTheTimezone = GetTimeZone(gTwsTimezoneNameToStandardTimeZoneName(pTimeZoneId))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getTickerIdFromIndex(ByVal pIndex As Long) As Long
getTickerIdFromIndex = pIndex + mTickerTable(pIndex).Incarnation
End Function

Private Sub notifyCurrentMarketDataValues(ByVal pTickerId As Long, ByVal pListener As IGenericTickListener, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "notifyCurrentMarketDataValues"
On Error GoTo Err

Dim ev As GenericTickEventData

With mTickerTable(pTickerId)
    ev.Tick.Timestamp = .TimeZone.ConvertDateLocalToTz(GetTimestamp)
    Set ev.Source = pDataSource
    
    If .PrevAsk <> MaxDouble Then
        ev.Tick.TickType = TickTypeAsk
        ev.Tick.Price = .PrevAsk
        ev.Tick.Size = .PrevAskSize
        pListener.NotifyTick ev
    End If
    
    If .PrevBid <> MaxDouble Then
        ev.Tick.TickType = TickTypeBid
        ev.Tick.Price = .PrevBid
        ev.Tick.Size = .PrevBidSize
        pListener.NotifyTick ev
    End If
    
    If .PrevTrade <> MaxDouble Then
        ev.Tick.TickType = TickTypeTrade
        ev.Tick.Price = .PrevTrade
        ev.Tick.Size = .PrevTradeSize
        pListener.NotifyTick ev
    End If
    
    If .PrevVolume <> MaxLong Then
        ev.Tick.TickType = TickTypeVolume
        ev.Tick.Size = .PrevVolume
        pListener.NotifyTick ev
    End If
    
    If .PrevHigh <> MaxDouble Then
        ev.Tick.TickType = TickTypeHighPrice
        ev.Tick.Price = .PrevHigh
        pListener.NotifyTick ev
    End If
    
    If .PrevLow <> MaxDouble Then
        ev.Tick.TickType = TickTypeLowPrice
        ev.Tick.Price = .PrevLow
        pListener.NotifyTick ev
    End If
    
    If .PrevOpen <> MaxDouble Then
        ev.Tick.TickType = TickTypeOpenPrice
        ev.Tick.Price = .PrevOpen
        pListener.NotifyTick ev
    End If
    
    If .PrevClose <> MaxDouble Then
        ev.Tick.TickType = TickTypeClosePrice
        ev.Tick.Price = .PrevClose
        pListener.NotifyTick ev
    End If
    
    If .PrevOpenInterest <> MaxLong Then
        ev.Tick.TickType = TickTypeOpenInterest
        ev.Tick.Price = .PrevOpenInterest
        pListener.NotifyTick ev
    End If

End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub releaseAllTickerTableEntries()
Const ProcName As String = "releaseAllTickerTableEntries"
On Error GoTo Err

Dim i As Long
Dim var As Variant

For Each var In mTickerMap
    i = CLng(var)
    releaseTickerTableEntry i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub releaseTickerTableEntry(ByVal pIndex As Long)
Const ProcName As String = "releaseTickerTableEntry"

On Error GoTo Err

If mLastFreeTickerTableIndex <> NullIndex Then
    mTickerTable(mLastFreeTickerTableIndex).Next = pIndex
Else
    mFirstFreeTickerTableIndex = pIndex
End If
mLastFreeTickerTableIndex = pIndex

With mTickerTable(pIndex)
    
    .Next = NullIndex
    .Incarnation = updateIncarnation(pIndex, .Incarnation)
    Set .ContractDetails = Nothing
    .DataHasBeenReceived = False
    Set .Subscriptions = Nothing
End With
clearPrevPrices pIndex

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub reqMarketData( _
                ByVal pIndex As Long)
Const ProcName As String = "reqMarketData"
On Error GoTo Err

If mTwsAPI.ConnectionState = TwsConnConnected Then mTwsAPI.RequestMarketData getTickerIdFromIndex(pIndex), mTickerTable(pIndex).ContractDetails.Summary, "", False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub reqMarketDepth( _
                ByVal pIndex As Long)
Const ProcName As String = "reqMarketDepth"
On Error GoTo Err

If mTwsAPI.ConnectionState = TwsConnConnected Then mTwsAPI.RequestMarketDepth getTickerIdFromIndex(pIndex), mTickerTable(pIndex).ContractDetails.Summary, mTickerTable(pIndex).NumberOfRows

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub ResetMarketDepth(ByVal pReEstablish As Boolean)
Const ProcName As String = "ResetMarketDepth"
On Error GoTo Err

Dim i As Long
Dim lTick As GenericTick
Dim var As Variant


lTick.TickType = TickTypeMarketDepthReset

For Each var In mTickerMap
    i = CLng(var)

    With mTickerTable(i)
        If .Subscriptions.NumberOfMarketDepthSubscriptions <> 0 Then
            lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .TimeZone)
            If pReEstablish Then mTwsAPI.RequestMarketDepth i, .ContractDetails.Summary, .NumberOfRows
            .Subscriptions.FireTick lTick
        End If
    End With
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function tickerHasBeenStopped(ByVal pTickerId As Long) As Boolean
Dim lIndex As Long
Const ProcName As String = "tickerHasBeenStopped"
On Error GoTo Err

lIndex = getIndexFromTickerId(pTickerId)
tickerHasBeenStopped = (pTickerId <> getTickerIdFromIndex(lIndex))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function validPrice( _
                ByVal pPrice As Double, _
                ByVal pPrevPrice As Double, _
                ByVal pPriceType As String, _
                ByVal pContractDetails As TwsContractDetails) As Boolean
Const ProcName As String = "validPrice"

On Error GoTo Err

If (pPrevPrice <> MaxDouble Or (pPrice <> 0 And pPrice <> -1)) And _
    IsValidPrice(pPrice, pPrevPrice, pContractDetails.Summary.SecType, pContractDetails.MinTick) _
Then
    validPrice = True
Else
    gLogger.Log "Invalid pPrice rejected (" & pPriceType & "): " & _
                    "price=" & pPrice & _
                    "; contract={" & pContractDetails.Summary.ToString & "}", _
                 ProcName, ModuleName
    Exit Function
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function updateIncarnation(ByVal pIndex As Long, ByVal pIncarnation As Long) As Long
Dim lIncarnation As Long
lIncarnation = pIncarnation + IncarnationIncrement
If pIndex + lIncarnation > mTwsAPI.MaxMarketDataRequestId Or _
    pIndex + lIncarnation > mTwsAPI.MaxMarketDepthRequestId _
Then lIncarnation = 0
updateIncarnation = lIncarnation
End Function



