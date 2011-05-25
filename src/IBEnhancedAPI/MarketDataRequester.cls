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
    PrevBidSize             As Long
    PrevAsk                 As Double
    PrevAskSize             As Long
    PrevTrade               As Double
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

Private WithEvents mTwsAPI                          As TwsAPI
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

Select Case pErrorCode
Case 200    ' security definition not known
    ' make sure we are not already successfully receiving data
    ' for this ticker. Occasionally Tws will send a 200 spuriously
    ' while the data stream is running
    
    If mTickerTable(pTickerId Mod IncarnationIncrement).DataHasBeenReceived Then Exit Sub
Case 316
    ResetMarketDepth pReEstablish:=True
    Exit Sub
Case 317
    ResetMarketDepth pReEstablish:=False
    Exit Sub
End Select

mTickerTable(pTickerId Mod IncarnationIncrement).Subscriptions.FireMarketDataError pErrorCode, pErrorMsg

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

If tickerHasBeenStopped(pTickerId) Then Exit Sub

lTick.Price = pPrice
lTick.Size = pSize

With mTickerTable(pTickerId Mod IncarnationIncrement)

    .DataHasBeenReceived = True

    lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .TimeZone)
    
    Select Case pTickType
    Case TwsTickTypes.TwsTickBid
        If Not validPrice(lTick.Price, .PrevBid, "bid", .ContractDetails) Then Exit Sub
        lTick.TickType = TickTypeBid
        If lTick.Size = 0 Then lTick.Size = .PrevBidSize
        .PrevBid = lTick.Price
        If lTick.Size <> 0 Then .PrevBidSize = lTick.Size
    Case TwsTickTypes.TwsTickAsk
        If Not validPrice(lTick.Price, .PrevAsk, "ask", .ContractDetails) Then Exit Sub
        lTick.TickType = TickTypeAsk
        If lTick.Size = 0 Then lTick.Size = .PrevAskSize
        .PrevAsk = lTick.Price
        If lTick.Size <> 0 Then .PrevAskSize = lTick.Size
    Case TwsTickTypes.TwsTickLast
        If Not validPrice(lTick.Price, .PrevTrade, "trade", .ContractDetails) Then Exit Sub
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

If tickerHasBeenStopped(pTickerId) Then Exit Sub

With mTickerTable(pTickerId Mod IncarnationIncrement)
    
    .DataHasBeenReceived = True
    
    lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .TimeZone)
    
    Select Case pTickType
    Case TwsTickTypes.TwsTickBidSize
        If pSize = .PrevBidSize Then
            ' this happens when Tws has sent a tick price including the
            ' Size. It then sends the Size again as a ticksize
            Exit Sub
        End If
        
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

Select Case pErrorCode
Case 316
    ResetMarketDepth pReEstablish:=True
    Exit Sub
Case 317
    ResetMarketDepth pReEstablish:=False
    Exit Sub
End Select

mTickerTable(pTickerId Mod IncarnationIncrement).Subscriptions.FireMarketDepthError pErrorCode, pErrorMsg

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IMarketDepthConsumer_NotifyMarketDepth( _
                ByVal pTickerId As Long, _
                ByVal pPosition As Long, _
                ByVal pMarketMaker As String, _
                ByVal pOperation As DOMOperations, _
                ByVal pSide As DOMSides, _
                ByVal pPrice As Double, _
                ByVal pSize As Long)
Const ProcName As String = "IMarketDepthConsumer_NotifyMarketDepth"
On Error GoTo Err

Dim lTick As GenericTick

If tickerHasBeenStopped(pTickerId) Then Exit Sub

lTick.TickType = TickTypeMarketDepth
lTick.MarketMaker = pMarketMaker
lTick.Operation = pOperation
lTick.position = pPosition
lTick.Price = pPrice
lTick.Size = pSize

With mTickerTable(pTickerId Mod IncarnationIncrement)

    .DataHasBeenReceived = True

    lTick.Timestamp = ConvertDateUTCToTZ(GetTimestampUTC, .TimeZone)
    
    .Subscriptions.FireTick lTick
End With
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IMarketDepthConsumer_ResetMarketDepth(ByVal pReEstablish As Boolean)
Const ProcName As String = "IMarketDepthConsumer_ResetMarketDepth"
On Error GoTo Err

ResetMarketDepth pReEstablish

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mTwsAPI Event Handlers
'@================================================================================

Private Sub mTwsAPI_Connected()
' now restart any tickers that were previously in operation
Const ProcName As String = "mTwsAPI_Connected"
On Error GoTo Err

reEstablishMarketData True

' and re-Request any market depth
ResetMarketDepth True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mTwsAPI_Disconnected(ByVal pReason As String, ByVal pForced As Boolean)
Const ProcName As String = "mTwsAPI_Disconnected"
On Error GoTo Err

releaseAllTickerTableEntries

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub CancelMarketData( _
                ByVal pTickerId As Long, _
                ByVal pDataSource As MarketDataSource)
Const ProcName As String = "CancelMarketData"
On Error GoTo Err

mTickerTable(pTickerId).Subscriptions.RemoveMarketDataSubscription pDataSource
If mTickerTable(pTickerId).Subscriptions.NumberOfMarketDataSubscriptions = 0 Then
    cancelMarketDataEx pTickerId
    If mTickerTable(pTickerId).Subscriptions.NumberOfMarketDepthSubscriptions = 0 Then
        mTickerMap.Remove mTickerTable(pTickerId).ContractDetails.Summary
        releaseTickerTableEntry pTickerId
    End If
End If

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub CancelMarketDepth( _
                ByVal pTickerId As Long, _
                ByVal pDataSource As MarketDataSource)
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err

If tickerHasBeenStopped(pTickerId) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid ticker id " & pTickerId

mTickerTable(pTickerId).Subscriptions.RemoveMarketDepthSubscription pDataSource
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

Friend Function RequestMarketData( _
                ByVal pContractDetails As TwsContractDetails, _
                ByVal pDataSource As MarketDataSource) As Long
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

Dim lTickerId As Long

If mTickerMap.Contains(pContractDetails.Summary) Then
    lTickerId = mTickerMap.Item(pContractDetails.Summary)
Else
    lTickerId = allocateTickerTableEntry
    Set mTickerTable(lTickerId).ContractDetails = pContractDetails
    Set mTickerTable(lTickerId).TimeZone = getTheTimezone(pContractDetails.TimeZoneId)
    
    mTickerMap.Add pContractDetails.Summary, lTickerId
End If

If mTickerTable(lTickerId).Subscriptions.NumberOfMarketDataSubscriptions = 0 Then
    reqMarketData lTickerId
Else
    notifyCurrentMarketDataValues lTickerId, pDataSource
End If
mTickerTable(lTickerId).Subscriptions.AddMarketDataSubscription pDataSource

RequestMarketData = lTickerId
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Function RequestMarketDepth( _
                ByVal pContractDetails As TwsContractDetails, _
                ByVal pDataSource As MarketDataSource, _
                ByVal pNumberOfRows As Long) As Long
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

Dim lTickerId As Long

If mTickerMap.Contains(pContractDetails.Summary) Then
    lTickerId = mTickerMap.Item(pContractDetails.Summary)
Else
    lTickerId = allocateTickerTableEntry
    Set mTickerTable(lTickerId).ContractDetails = pContractDetails
    Set mTickerTable(lTickerId).TimeZone = getTheTimezone(pContractDetails.TimeZoneId)
    
    mTickerMap.Add pContractDetails.Summary, lTickerId
End If

If pNumberOfRows > mTickerTable(lTickerId).NumberOfRows Then mTickerTable(lTickerId).NumberOfRows = pNumberOfRows

If mTickerTable(lTickerId).Subscriptions.NumberOfMarketDepthSubscriptions = 0 Then reqMarketDepth lTickerId

mTickerTable(lTickerId).Subscriptions.AddMarketDepthSubscription pDataSource

RequestMarketDepth = lTickerId
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

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub cancelMarketDataEx(ByVal pTickerId As Long)
Const ProcName As String = "cancelMarketDataEx"
On Error GoTo Err

mTwsAPI.CancelMarketData pTickerId + mTickerTable(pTickerId).Incarnation

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub cancelMarketDepthEx(ByVal pTickerId As Long)
Const ProcName As String = "cancelMarketDepthEx"
On Error GoTo Err

mTwsAPI.CancelMarketDepth pTickerId + mTickerTable(pTickerId).Incarnation

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

Private Function getTheTimezone(ByVal pTimeZoneId As String) As TimeZone
Const ProcName As String = "getTheTimezone"
On Error GoTo Err

Set getTheTimezone = GetTimeZone(gTwsTimezoneNameToStandardTimeZoneName(pTimeZoneId))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub notifyCurrentMarketDataValues(ByVal pTickerId As Long, ByVal pDataSource As MarketDataSource)
Const ProcName As String = "notifyCurrentMarketDataValues"
On Error GoTo Err

Dim lTick As GenericTick

With mTickerTable(pTickerId)
    lTick.Timestamp = .TimeZone.ConvertDateLocalToTz(GetTimestamp)
    
    lTick.TickType = TickTypeAsk
    lTick.Price = .PrevAsk
    lTick.Size = .PrevAskSize
    pDataSource.NotifyTick lTick
    
    lTick.TickType = TickTypeBid
    lTick.Price = .PrevBid
    lTick.Size = .PrevBidSize
    pDataSource.NotifyTick lTick
    
    lTick.TickType = TickTypeTrade
    lTick.Price = .PrevTrade
    lTick.Size = .PrevTradeSize
    pDataSource.NotifyTick lTick
    
    lTick.TickType = TickTypeVolume
    lTick.Size = .PrevVolume
    pDataSource.NotifyTick lTick
    
    lTick.TickType = TickTypeHighPrice
    lTick.Price = .PrevHigh
    pDataSource.NotifyTick lTick
    
    lTick.TickType = TickTypeLowPrice
    lTick.Price = .PrevLow
    pDataSource.NotifyTick lTick
    
    lTick.TickType = TickTypeOpenPrice
    lTick.Price = .PrevOpen
    pDataSource.NotifyTick lTick
    
    lTick.TickType = TickTypeClosePrice
    lTick.Price = .PrevClose
    pDataSource.NotifyTick lTick
    
    lTick.TickType = TickTypeOpenInterest
    lTick.Price = .PrevOpenInterest
    pDataSource.NotifyTick lTick
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub reEstablishMarketData(ByVal pNewConnection As Boolean)
Const ProcName As String = "reEstablishMarketData"
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

mTwsAPI.RequestMarketData pIndex + mTickerTable(pIndex).Incarnation, mTickerTable(pIndex).ContractDetails.Summary, "", False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub reqMarketDepth( _
                ByVal pIndex As Long)
Const ProcName As String = "reqMarketDepth"
On Error GoTo Err

mTwsAPI.RequestMarketDepth pIndex + mTickerTable(pIndex).Incarnation, mTickerTable(pIndex).ContractDetails.Summary, mTickerTable(pIndex).NumberOfRows

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

lIndex = (pTickerId Mod IncarnationIncrement)
tickerHasBeenStopped = (pTickerId <> (lIndex + mTickerTable(lIndex).Incarnation))

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



