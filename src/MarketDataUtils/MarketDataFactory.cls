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
Implements TaskCompletionListener

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
    DataListener            As IGenericTickListener
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
End Type

Private Type MarketDepthTableEntry
    Incarnation             As Long
    ContractDetails         As TwsContractDetails
    TimeZone                As TimeZone
    DataListener            As IGenericTickListener
    ErrorListener           As IMarketDataErrorListener
    NumberOfRows            As Long
    DataHasBeenReceived     As Boolean
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MarketDataFactory"

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

Private mMarketDepthTable()                         As MarketDepthTableEntry

Private mContractRequester                          As ContractDetailsRequester

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mTickerTable(3) As TickerTableEntry
mFirstFreeTickerTableIndex = NullIndex
mLastFreeTickerTableIndex = NullIndex

ReDim mMarketDepthTable(3) As MarketDepthTableEntry

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
    
    If mTickerTable(lIndex).DataHasBeenReceived Then Exit Sub
End Select

mTickerTable(lIndex).ErrorListener.NotifyMarketDataError lIndex, pErrorCode, pErrorMsg
releaseTickerTableEntry lIndex

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

With mTickerTable(lIndex)

    .DataHasBeenReceived = True

    lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .TimeZone)
    
    Select Case pTickType
    Case TwsTickTypes.TwsTickBid
        .PrevBidValid = validPrice(lTick.Price, .PrevBid, "bid", .ContractDetails)
        If Not .PrevBidValid Then Exit Sub
        lTick.TickType = TickTypeBid
        If lTick.Size = 0 And .PrevBidSize <> 0 Then lTick.Size = .PrevBidSize
        .PrevBid = lTick.Price
        If lTick.Size <> 0 Then .PrevBidSize = lTick.Size
    Case TwsTickTypes.TwsTickAsk
        .PrevAskValid = validPrice(lTick.Price, .PrevAsk, "ask", .ContractDetails)
        If Not .PrevAskValid Then Exit Sub
        lTick.TickType = TickTypeAsk
        If lTick.Size = 0 And .PrevAskSize <> 0 Then lTick.Size = .PrevAskSize
        .PrevAsk = lTick.Price
        If lTick.Size <> 0 Then .PrevAskSize = lTick.Size
    Case TwsTickTypes.TwsTickLast
        .PrevTradeValid = validPrice(lTick.Price, .PrevTrade, "trade", .ContractDetails)
        If Not .PrevTradeValid Then Exit Sub
        lTick.TickType = TickTypeTrade
        If lTick.Size = 0 And .PrevTradeSize <> 0 Then lTick.Size = .PrevTradeSize
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

Private Sub IMarketDataFactory_CancelMarketData(ByVal pStreamId As Long)
Const ProcName As String = "IMarketDataFactory_CancelMarketData"
On Error GoTo Err

CancelMarketData pStreamId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataFactory_CancelMarketDepth(ByVal pStreamId As Long)
Const ProcName As String = "IMarketDataFactory_CancelMarketDepth"
On Error GoTo Err

CancelMarketDepth pStreamId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IMarketDataFactory_CreateMarketDataStream(ByVal pContract As ContractUtils27.IContract, ByVal pDataListener As TickUtils27.IGenericTickListener, ByVal pErrorListener As TickUtils27.IMarketDataErrorListener, Optional ByVal pStreamId As Long) As TickUtils27.ITickStream

End Function

Private Function IMarketDataFactory_CreateMarketDepthStream(ByVal pContract As ContractUtils27.IContract, ByVal pNumberOfRows As Long, ByVal pDataListener As TickUtils27.IGenericTickListener, ByVal pErrorListener As TickUtils27.IMarketDataErrorListener, Optional ByVal pStreamId As Long) As TickUtils27.ITickStream

End Function

Private Property Get IMarketDataFactory_MaxStreamId() As Long
IMarketDataFactory_MaxStreamId = IncarnationIncrement
End Property

Private Sub IMarketDataFactory_RequestMarketData(ByVal pStreamId As Long, ByVal pContract As IContract, ByVal pDataListener As IGenericTickListener, ByVal pErrorListener As IMarketDataErrorListener)
Const ProcName As String = "IMarketDataFactory_RequestMarketData"
On Error GoTo Err

RequestMarketData pStreamId, pContract, pDataListener, pErrorListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataFactory_RequestMarketDepth(ByVal pStreamId As Long, ByVal pContract As IContract, ByVal pNumberOfRows As Long, ByVal pDataListener As IGenericTickListener, ByVal pErrorListener As IMarketDataErrorListener)
Const ProcName As String = "IMarketDataFactory_RequestMarketDepth"
On Error GoTo Err

RequestMarketDepth pStreamId, pContract, pNumberOfRows, pDataListener, pErrorListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
releaseMarketDepthTableEntry lIndex

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

With mTickerTable(lIndex)

    .DataHasBeenReceived = True

    lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .TimeZone)
    
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
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_TaskCompleted"
On Error GoTo Err

Dim lContractDetails As TwsContractDetails
Dim lStreamId As Long

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
Else
    Set lContractDetails = ev.Result
    lStreamId = ev.Cookie
    Set mTickerTable(lStreamId).ContractDetails = lContractDetails
    Set mTickerTable(lStreamId).TimeZone = getTheTimezone(lContractDetails.TimeZoneId)
    reqMarketData lStreamId
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get CurrentAskPrice(ByVal pTickerId As Long) As Double
Const ProcName As String = "CurrentAskPrice"
On Error GoTo Err

Assert Not tickerHasBeenStopped(pTickerId), "Ticker has been stopped"

CurrentAskPrice = mTickerTable(getIndexFromTickerId(pTickerId)).PrevAsk

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentAskSize(ByVal pTickerId As Long) As Long
Const ProcName As String = "CurrentAskSize"
On Error GoTo Err

Assert Not tickerHasBeenStopped(pTickerId), "Ticker has been stopped"

CurrentAskSize = mTickerTable(getIndexFromTickerId(pTickerId)).PrevAskSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentBidPrice(ByVal pTickerId As Long) As Double
Const ProcName As String = "CurrentBidPrice"
On Error GoTo Err

Assert Not tickerHasBeenStopped(pTickerId), "Ticker has been stopped"

CurrentBidPrice = mTickerTable(getIndexFromTickerId(pTickerId)).PrevBid

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CurrentBidSize(ByVal pTickerId As Long) As Long
CurrentBidSize = mTickerTable(pTickerId).PrevBidSize
End Property

Public Property Get CurrentOpenInterest(ByVal pTickerId As Long) As Long
CurrentOpenInterest = mTickerTable(pTickerId).PrevOpenInterest
End Property

Public Property Get CurrentTradePrice(ByVal pTickerId As Long) As Double
CurrentTradePrice = mTickerTable(pTickerId).PrevTrade
End Property

Public Property Get CurrentTradeSize(ByVal pTickerId As Long) As Long
CurrentTradeSize = mTickerTable(pTickerId).PrevTradeSize
End Property

Public Property Get CurrentVolume(ByVal pTickerId As Long) As Long
CurrentVolume = mTickerTable(pTickerId).PrevVolume
End Property

Public Property Get MaxTickerId() As Long
MaxTickerId = IncarnationIncrement
End Property

Public Property Get PreviousSessionClosePrice(ByVal pTickerId As Long) As Double
PreviousSessionClosePrice = mTickerTable(pTickerId).PrevClose
End Property

Public Property Get SessionHighPrice(ByVal pTickerId As Long) As Double
SessionHighPrice = mTickerTable(pTickerId).PrevHigh
End Property

Public Property Get SessionLowPrice(ByVal pTickerId As Long) As Double
SessionLowPrice = mTickerTable(pTickerId).PrevLow
End Property

Public Property Get Timestamp(ByVal pTickerId As Long) As Date
Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, mTickerTable(pTickerId).TimeZone)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub CancelMarketData( _
                ByVal pStreamId As Long)
Const ProcName As String = "CancelMarketData"
On Error GoTo Err

cancelMarketDataEx pStreamId
releaseTickerTableEntry pStreamId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub CancelMarketDepth( _
                ByVal pStreamId As Long)
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err

cancelMarketDepthEx pStreamId
releaseMarketDepthTableEntry pStreamId

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pTwsAPI As TwsAPI, _
                ByVal pContractRequester As ContractDetailsRequester)
Set mTwsAPI = pTwsAPI
Set mContractRequester = pContractRequester
End Sub

Friend Sub ReEstablishMarketData(ByVal pNewConnection As Boolean)
Const ProcName As String = "ReEstablishMarketData"
Dim i As Long
Dim var As Variant

On Error GoTo Err

gLog "Re-establishing market data: newConnection = " & CStr(pNewConnection), ModuleName, ProcName

For i = 0 To UBound(mTickerTable)
    With mTickerTable(i)
        If Not .ContractDetails Is Nothing Then
            If Not pNewConnection Then cancelMarketDataEx i
            .Incarnation = updateIncarnation(i, .Incarnation)
            reqMarketData i
        End If
        clearPrevPrices i
    End With
Next

For i = 0 To UBound(mMarketDepthTable)
    With mMarketDepthTable(i)
        If Not .ContractDetails Is Nothing Then
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

Public Sub RequestMarketData( _
                ByVal pStreamId As Long, _
                ByVal pContract As Contract, _
                ByVal pDataListener As IGenericTickListener, _
                ByVal pErrorListener As IMarketDataErrorListener)
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

Dim lTask As ComboContractConversionTask
Dim lContractDetails As TwsContractDetails

initialiseTickerTableEntry pStreamId, pDataListener, pErrorListener

If pContract.Specifier.Sectype = SecTypeCombo Then
    Set lTask = New ComboContractConversionTask
    lTask.Initialise mContractRequester, pContract.Specifier
    StartTask(lTask, PriorityNormal, , pStreamId).AddTaskCompletionListener Me
Else
    Set lContractDetails = gContractToTwsContractDetails(pContract)
    Set mTickerTable(pStreamId).ContractDetails = lContractDetails
    Set mTickerTable(pStreamId).TimeZone = getTheTimezone(lContractDetails.TimeZoneId)
    reqMarketData pStreamId
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RequestMarketDepth( _
                ByVal pStreamId As Long, _
                ByVal pContract As Contract, _
                ByVal pNumberOfRows As Long, _
                ByVal pDataListener As IGenericTickListener, _
                ByVal pErrorListener As IMarketDataErrorListener)
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

Dim lContractDetails As TwsContractDetails

Assert pContract.Specifier.Sectype <> SecTypeCombo, "Market depth not supported for Combo contracts", ErrorCodes.ErrUnsupportedOperationException

initialiseMarketDepthTableEntry pStreamId, pDataListener, pErrorListener
Set lContractDetails = gContractToTwsContractDetails(pContract)
Set mMarketDepthTable(pStreamId).ContractDetails = lContractDetails
Set mMarketDepthTable(pStreamId).TimeZone = getTheTimezone(lContractDetails.TimeZoneId)
mMarketDepthTable(pStreamId).NumberOfRows = pNumberOfRows

reqMarketDepth pStreamId

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

With mTickerTable(pIndex)
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

Dim ev As GenericTickEventData

ev.StreamId = pStreamId
ev.Tick = pTick

Set ev.Source = Me
If pTick.TickType = TickTypeMarketDepth Or _
    pTick.TickType = TickTypeMarketDepthReset _
Then
    mMarketDepthTable(pStreamId).DataListener.NotifyTick ev
Else
    mTickerTable(pStreamId).DataListener.NotifyTick ev
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
getTickerIdFromIndex = pIndex + mTickerTable(pIndex).Incarnation
End Function

Private Sub initialiseMarketDepthTableEntry( _
                ByVal pIndex As Long, _
                ByVal pDataListener As IGenericTickListener, _
                ByVal pErrorListener As IMarketDataErrorListener)
Const ProcName As String = "initialiseMarketDepthTableEntry"
On Error GoTo Err

Do While pIndex > UBound(mMarketDepthTable)
    ReDim Preserve mMarketDepthTable(2 * (UBound(mMarketDepthTable) + 1) - 1) As MarketDepthTableEntry
Loop

With mMarketDepthTable(pIndex)
    Set .DataListener = pDataListener
    Set .ErrorListener = pErrorListener
End With
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initialiseTickerTableEntry( _
                ByVal pIndex As Long, _
                ByVal pDataListener As IGenericTickListener, _
                ByVal pErrorListener As IMarketDataErrorListener)
Const ProcName As String = "initialiseTickerTableEntry"
On Error GoTo Err

Do While pIndex > UBound(mTickerTable)
    ReDim Preserve mTickerTable(2 * (UBound(mTickerTable) + 1) - 1) As TickerTableEntry
Loop

With mTickerTable(pIndex)
    Set .DataListener = pDataListener
    Set .ErrorListener = pErrorListener
End With
clearPrevPrices pIndex
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub releaseMarketDepthTableEntry(ByVal pIndex As Long)
Const ProcName As String = "releaseMarketDepthTableEntry"

On Error GoTo Err

With mMarketDepthTable(pIndex)
    .Incarnation = updateIncarnation(pIndex, .Incarnation)
    Set .ContractDetails = Nothing
    .DataHasBeenReceived = False
End With
clearPrevPrices pIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub releaseTickerTableEntry(ByVal pIndex As Long)
Const ProcName As String = "releaseTickerTableEntry"

On Error GoTo Err

With mTickerTable(pIndex)
    .Incarnation = updateIncarnation(pIndex, .Incarnation)
    Set .ContractDetails = Nothing
    .DataHasBeenReceived = False
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

If mTwsAPI.ConnectionState = TwsConnConnected Then mTwsAPI.RequestMarketData getTickerIdFromIndex(pIndex), mTickerTable(pIndex).ContractDetails.Summary, "", False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub reqMarketDepth( _
                ByVal pIndex As Long)
Const ProcName As String = "reqMarketDepth"
On Error GoTo Err

If mTwsAPI.ConnectionState = TwsConnConnected Then mTwsAPI.RequestMarketDepth getMarketDepthIdFromIndex(pIndex), mMarketDepthTable(pIndex).ContractDetails.Summary, mMarketDepthTable(pIndex).NumberOfRows

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
        If Not .ContractDetails Is Nothing Then
            lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .TimeZone)
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
                ByVal pContractDetails As TwsContractDetails) As Boolean
Const ProcName As String = "validPrice"

On Error GoTo Err

If (pPrevPrice <> MaxDouble Or (pPrice <> 0# And pPrice <> -1# And pPrice <> -100#)) And _
    IsValidPrice(pPrice, pPrevPrice, pContractDetails.Summary.Sectype, pContractDetails.MinTick) _
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



