VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdvancedOrderSimulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event errorMessage( _
                ByVal timestamp As Date, _
                ByVal id As Long, _
                ByVal errorCode As Long, _
                ByVal errorMsg As String)

Event ExecutionDetails( _
                ByVal timestamp As Date, _
                ByVal id As Long, _
                ByVal pContractSpecifier As contractSpecifier, _
                ByVal exec As Execution)

Event openOrder( _
                ByVal timestamp As Date, _
                ByVal pContractSpecifier As contractSpecifier, _
                ByVal pOrder As Order)

Event orderStatus(ByVal timestamp As Date, _
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
                ByVal contractSpecifier As contractSpecifier, _
                ByRef fillPrice As Double, _
                ByRef fillSize As Long)


Private Type MarketTableEntry
    inUse                   As Boolean
    contractSpecifier       As contractSpecifier
    bid                     As Double
    bidSize                 As Long
    ask                     As Double
    AskSize                 As Long
    activeOrders            As Collection
    pendingOrders           As Collection
    clock                   As clock
End Type

Private mMarketTable() As MarketTableEntry

Private mAllOrders As Collection

Private mNextPermId As Long

Private mSimulateExchangeDelays As Boolean
Private WithEvents mTimerList As TimerList
Attribute mTimerList.VB_VarHelpID = -1
    ' This is used to simulate delays in submitting orders to the
    ' exchange for execution. When it is time to change an order's status
    ' from pendingsubmit to submitted, the order is added to this list
    ' with a random wait of between 200 and 500 millisecs.

Private Sub Class_Initialize()
ReDim mMarketTable(InitialMaxTickers) As MarketTableEntry
Set mAllOrders = New Collection

Randomize
mNextPermId = Fix(Rnd() * 1999999999 + 1)
Set mTimerList = New TimerUtils.TimerList
mTimerList.TimerIntervalMillisecs = 200
mTimerList.StartTiming
mSimulateExchangeDelays = True
End Sub

Private Sub mTimerList_expired(ByVal data As Object)
Dim lOrder As Order

Set lOrder = data

Select Case lOrder.status
Case OrderStatuses.OrderStatusSubmitted
    lOrder.atExchange = True
    matchOrder lOrder
Case OrderStatuses.OrderStatusCancelling
    setCancelOrder lOrder
End Select
End Sub

Friend Property Let simulateExchangeDelays(ByVal value As Boolean)
mSimulateExchangeDelays = value
End Property

'Friend Property Let MaxTickers(ByVal value As Long)
'ReDim Preserve mMarketTable(value - 1) As MarketTableEntry
'End Property

Friend Sub cancelOrder(ByVal id As Long)
Dim lOrder As Order
Dim pendingOrder As Order
Dim relatedOrder As Order

On Error Resume Next
Set lOrder = mAllOrders(CStr(id))
On Error GoTo 0

If lOrder Is Nothing Then
    err.Raise ErrorCodes.UnknownOrderID, _
                "TradeBuild.CAdvancedOrderSimulator::cancelOrder", _
                "Unknown order id " & id
End If

With mMarketTable(lOrder.marketIndex)
    
    On Error Resume Next
    Set pendingOrder = .pendingOrders(CStr(id))
    On Error GoTo 0
    
    If Not pendingOrder Is Nothing Then
        If lOrder.parentId <> 0 Then
            Set relatedOrder = .pendingOrders(CStr(lOrder.parentId))
            setCancelOrder relatedOrder
            
            For Each relatedOrder In .pendingOrders
                If relatedOrder.parentId = lOrder.parentId And _
                    relatedOrder.id <> lOrder.id _
                Then
                    setCancelOrder relatedOrder
                End If
            Next
        Else
            For Each relatedOrder In .pendingOrders
                If relatedOrder.parentId = lOrder.id Or _
                    (relatedOrder.ocaGroup = lOrder.ocaGroup And _
                    lOrder.ocaGroup <> "" And _
                    relatedOrder.id <> lOrder.id) _
                Then
                    setCancelOrder relatedOrder
                End If
            Next
        End If
    Else
        On Error Resume Next
        Set lOrder = .activeOrders(CStr(id))
        On Error GoTo 0
        
        If lOrder Is Nothing Then Exit Sub
        
        If lOrder.parentId <> 0 Then
            On Error Resume Next
            Set relatedOrder = .activeOrders(CStr(lOrder.parentId))
            On Error GoTo 0
            If Not relatedOrder Is Nothing Then
                setCancelOrder relatedOrder
            End If
            For Each relatedOrder In .activeOrders
                If relatedOrder.parentId = lOrder.parentId And _
                    relatedOrder.id <> lOrder.id _
                Then
                    setCancelOrder relatedOrder
                End If
            Next
        Else
            For Each relatedOrder In .activeOrders
                If relatedOrder.parentId = lOrder.id Or _
                    (relatedOrder.ocaGroup = lOrder.ocaGroup And _
                    lOrder.ocaGroup <> "" And _
                    relatedOrder.id <> lOrder.id) _
                  Then
                    setCancelOrder relatedOrder
                End If
            Next
        End If
    End If

    setCancelOrder lOrder

End With
End Sub

Friend Sub placeOrder(ByVal pContractSpecifier As contractSpecifier, _
                    ByVal pOrder As Order)

Dim lOrder As Order
Dim pendingOrder As Order
Dim newOrder As Boolean
Dim marketIndex As Long

Select Case pOrder.orderType
Case TradeBuild.OrderTypes.OrderTypeMarket, _
    TradeBuild.OrderTypes.OrderTypeLimit, _
    TradeBuild.OrderTypes.OrderTypeStop, _
    TradeBuild.OrderTypes.OrderTypeStopLimit
Case Else
    setRejectOrder pOrder, "Order type " & pOrder.orderType & " not supported by simulator"
    Exit Sub
End Select

For marketIndex = 0 To UBound(mMarketTable)
    If pContractSpecifier.Equals(mMarketTable(marketIndex).contractSpecifier) Then Exit For
Next

If marketIndex > UBound(mMarketTable) Then
    err.Raise ErrorCodes.UnexpectedContract, _
            "TradeBuild.CAdvancedOrderSimulator::placeOrder", _
            "Unexpected contract: " & vbCrLf & _
            pContractSpecifier.ToString
End If

On Error Resume Next
Set lOrder = mAllOrders(CStr(pOrder.id))
On Error GoTo 0

If lOrder Is Nothing Then
    newOrder = True
    Set lOrder = pOrder.Clone
    lOrder.marketIndex = marketIndex
    lOrder.permId = mNextPermId
    mNextPermId = mNextPermId + 1
    mAllOrders.add lOrder, CStr(lOrder.id)
Else
    lOrder.action = pOrder.action
    lOrder.auxPrice = pOrder.auxPrice
    lOrder.discretionaryAmt = pOrder.discretionaryAmt
    lOrder.displaySize = pOrder.displaySize
    lOrder.goodAfterTime = pOrder.goodAfterTime
    lOrder.hidden = pOrder.hidden
    lOrder.ignoreRTH = pOrder.ignoreRTH
    lOrder.limitPrice = pOrder.limitPrice
    lOrder.ocaGroup = pOrder.ocaGroup
    lOrder.openClose = pOrder.openClose
    lOrder.orderRef = pOrder.orderRef
    lOrder.orderType = pOrder.orderType
    lOrder.origin = pOrder.origin
    lOrder.parentId = pOrder.parentId
    lOrder.quantity = pOrder.quantity
    lOrder.sharesAllocation = pOrder.sharesAllocation
    lOrder.sweepToFill = pOrder.sweepToFill
    lOrder.timeInForce = pOrder.timeInForce
    lOrder.triggerMethod = pOrder.triggerMethod
    Exit Sub
End If

With mMarketTable(marketIndex)
    If Not pOrder.transmit Then
        .pendingOrders.add lOrder, CStr(lOrder.id)
    Else
        ' if this order was previously in the pending list, remove
        ' it now
        On Error Resume Next
        .pendingOrders.Remove CStr(lOrder.id)
        On Error GoTo 0
        
        ' now see if there are any other pending orders related
        ' to this one via parentid or oca group
        If lOrder.parentId <> 0 Then
            Set pendingOrder = .pendingOrders(CStr(lOrder.parentId))
            .activeOrders.add pendingOrder, CStr(pendingOrder.id)
            .pendingOrders.Remove CStr(pendingOrder.id)
            setSubmitOrder pendingOrder
            
            ' find any other child orders
            For Each pendingOrder In .pendingOrders
                If pendingOrder.parentId = lOrder.parentId Then
                    .activeOrders.add pendingOrder, CStr(pendingOrder.id)
                    .pendingOrders.Remove CStr(pendingOrder.id)
                    setSubmitOrder pendingOrder
                End If
            Next
        Else
            ' find any child orders or related OCA orders
            For Each pendingOrder In .pendingOrders
                If pendingOrder.parentId = lOrder.id Or _
                    (pendingOrder.ocaGroup = lOrder.ocaGroup And _
                    lOrder.ocaGroup <> "") _
                Then
                    .activeOrders.add pendingOrder, CStr(pendingOrder.id)
                    .pendingOrders.Remove CStr(pendingOrder.id)
                    setSubmitOrder pendingOrder
                End If
            Next
                        
        End If
        
        On Error Resume Next
        .activeOrders.add lOrder, CStr(lOrder.id)
        On Error GoTo 0
        setSubmitOrder lOrder
    End If
End With
End Sub
                    
Friend Sub removeCancelledOrder(ByVal id As Long)
' this function is called by the instrumentProcessor when a real order
' has been canceled, in case the simulator hasn't also cancelled it (because
' the simulator isn't an exact reflection of reality)


Dim lOrder As Order
Dim relatedOrder As Order
Dim theOrder As Order

On Error Resume Next
Set theOrder = mAllOrders(CStr(id))
On Error GoTo 0
If theOrder Is Nothing Then
    err.Raise ErrorCodes.UnknownOrderID, _
                "TradeBuild.CAdvancedOrderSimulator::removeCancelledOrder", _
                "Unknown order id " & id
End If

With mMarketTable(lOrder.marketIndex)
    
    On Error Resume Next
    
    Set lOrder = .activeOrders(CStr(id))
    If lOrder Is Nothing Then Exit Sub
    
    On Error GoTo 0
    
    .activeOrders.Remove CStr(id)
    
    If lOrder.parentId <> 0 Then
        For Each relatedOrder In .activeOrders
            If relatedOrder.parentId = lOrder.parentId And _
                relatedOrder.id <> lOrder.id _
            Then
                setCancelOrder relatedOrder
            End If
        Next
    Else
        For Each relatedOrder In .activeOrders
            If relatedOrder.parentId = lOrder.id Or _
                (relatedOrder.ocaGroup = lOrder.ocaGroup And _
                lOrder.ocaGroup <> "" And _
                relatedOrder.id <> lOrder.id) _
            Then
                setCancelOrder relatedOrder
            End If
        Next
    End If
    
End With

End Sub

Friend Sub removeFilledOrder(ByVal id As Long)
' this function is called by the instrumentProcessor when a real order
' has been filled, in case the simulator hasn't also filled it (because
' the simulator isn't an exact reflection of reality)
Dim theOrder As Order
Dim lOrder As Order

On Error Resume Next
Set theOrder = mAllOrders(CStr(id))
On Error GoTo 0
If theOrder Is Nothing Then
    err.Raise ErrorCodes.UnknownOrderID, _
                "TradeBuild.CAdvancedOrderSimulator::removeFilledOrder", _
                "Unknown order id " & id
End If

With mMarketTable(lOrder.marketIndex)
    
    On Error Resume Next
    
    Set lOrder = .activeOrders(CStr(id))
    If lOrder Is Nothing Then Exit Sub
    
    On Error GoTo 0
    
    handleRelatedOrders lOrder
    .activeOrders.Remove CStr(id)
End With
End Sub

Friend Sub removeOrder(ByVal id As Long)
Dim lOrder As Order

On Error Resume Next
Set lOrder = mAllOrders(CStr(id))
On Error GoTo 0
If lOrder Is Nothing Then
    err.Raise ErrorCodes.UnknownOrderID, _
                "TradeBuild.CAdvancedOrderSimulator::removeOrder", _
                "Unknown order id " & id
End If

With mMarketTable(lOrder.marketIndex)
    On Error Resume Next
    .activeOrders.Remove (CStr(id))
End With
End Sub

Friend Function startMarket( _
                    ByVal pContractSpecifier As contractSpecifier) As Long
Dim i As Long

startMarket = -1

For i = 0 To UBound(mMarketTable)
    If (Not mMarketTable(i).inUse) Then
        startMarket = i
        Exit For
    End If
Next

If startMarket = -1 Then
    startMarket = UBound(mMarketTable) + 1
    ReDim Preserve mMarketTable(UBound(mMarketTable) + 20) As MarketTableEntry
End If

mMarketTable(startMarket).inUse = True
With mMarketTable(startMarket)
    Set .contractSpecifier = pContractSpecifier.Clone
    Set .activeOrders = New Collection
    Set .pendingOrders = New Collection
    Set .clock = New clock
    .clock.SimulateTime True
End With

End Function

Friend Sub stopMarket(ByVal handle As Long)
With mMarketTable(handle)
    .inUse = False
    Set .contractSpecifier = Nothing
    Set .activeOrders = Nothing
    Set .pendingOrders = Nothing
    Set .clock = Nothing
End With
End Sub

Friend Sub bid(ByVal timestamp As Date, _
                ByVal TickerID As Long, _
                ByVal bid As Double)
                
Dim lOrder As Order

If Not mSimulateExchangeDelays Then handlePendingSubmitOrders TickerID

With mMarketTable(TickerID)
    .clock.SetTime timestamp
    .bid = bid
    For Each lOrder In .activeOrders
        If lOrder.status = OrderStatuses.OrderStatusSubmitted Then
            If lOrder.atExchange Then
                matchOrder lOrder
            End If
        ElseIf lOrder.status = OrderStatuses.OrderStatusCancelling Then
            If lOrder.atExchange Then
                matchOrder lOrder
                If lOrder.status = OrderStatuses.OrderStatusCancelling Then
                    setCancelOrder lOrder
                End If
            ElseIf Not mSimulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

End Sub

Friend Sub ask(ByVal timestamp As Date, _
                ByVal TickerID As Long, _
                ByVal ask As Double)
                
Dim lOrder As Order

If Not mSimulateExchangeDelays Then handlePendingSubmitOrders TickerID

With mMarketTable(TickerID)
    .clock.SetTime timestamp
    .ask = ask
    For Each lOrder In .activeOrders
        If lOrder.status = OrderStatuses.OrderStatusSubmitted Then
            If lOrder.atExchange Then
                matchOrder lOrder
            End If
        ElseIf lOrder.status = OrderStatuses.OrderStatusCancelling Then
            If lOrder.atExchange Then
                matchOrder lOrder
                If lOrder.status = OrderStatuses.OrderStatusCancelling Then
                    setCancelOrder lOrder
                End If
            ElseIf Not mSimulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

End Sub

Friend Sub trade(ByVal timestamp As Date, _
                ByVal TickerID As Long, _
                ByVal trade As Double)
                
Dim lOrder As Order


If Not mSimulateExchangeDelays Then handlePendingSubmitOrders TickerID

With mMarketTable(TickerID)
    .clock.SetTime timestamp
    For Each lOrder In .activeOrders
        If lOrder.status = OrderStatuses.OrderStatusSubmitted Then
            If lOrder.atExchange Then
                matchOrderTrade lOrder, trade
            End If
        ElseIf lOrder.status = OrderStatuses.OrderStatusPreSubmitted Then
            Select Case lOrder.orderType
                Case TradeBuild.OrderTypes.OrderTypeStop
                    If lOrder.action = OrderActions.ActionBuy And _
                        trade >= lOrder.auxPrice Then
                        'lOrder.orderType = TradeBuild.OrderTypes.OrderTypeMarket
                        'lOrder.auxPrice = 0
                        setSubmitOrder lOrder
                    ElseIf lOrder.action = OrderActions.ActionSell And _
                        trade <= lOrder.auxPrice Then
                        'lOrder.orderType = TradeBuild.OrderTypes.OrderTypeMarket
                        'lOrder.auxPrice = 0
                        setSubmitOrder lOrder
                    End If
                Case TradeBuild.OrderTypes.OrderTypeStopLimit
                    If lOrder.action = OrderActions.ActionBuy And _
                        trade >= lOrder.auxPrice Then
                        'lOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
                        'lOrder.auxPrice = 0
                        setSubmitOrder lOrder
                    ElseIf lOrder.action = OrderActions.ActionSell And _
                        trade <= lOrder.auxPrice Then
                        'lOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
                        'lOrder.auxPrice = 0
                        setSubmitOrder lOrder
                    End If
            End Select
        ElseIf lOrder.status = OrderStatuses.OrderStatusCancelling Then
            If lOrder.atExchange Then
                matchOrder lOrder
                If lOrder.status = OrderStatuses.OrderStatusCancelling Then
                    setCancelOrder lOrder
                End If
            ElseIf Not mSimulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

End Sub

Private Function delay() As Long
delay = 200 + Fix(Rnd() * 300 + 1)
End Function

Private Sub handlePendingSubmitOrders(ByVal marketID As Long)
Dim lOrder As Order
With mMarketTable(marketID)
    For Each lOrder In .activeOrders
        If lOrder.status = OrderStatuses.OrderStatusPendingSubmit And lOrder.parentId = 0 Then
            setSubmitOrder lOrder
        End If
    Next
End With
End Sub

Private Sub handleRelatedOrders(ByVal pOrder As Order)
Dim lOrder As Order

With mMarketTable(pOrder.marketIndex)
    For Each lOrder In .activeOrders
        If lOrder.parentId = pOrder.id Then
            setSubmitOrder lOrder
        End If
        If lOrder.parentId = pOrder.parentId And _
            lOrder.id <> pOrder.id And _
            lOrder.parentId <> 0 _
        Then
            setCancelOrder lOrder
        End If
        If lOrder.ocaGroup = pOrder.ocaGroup And _
            pOrder.ocaGroup <> "" And _
            lOrder.id <> pOrder.id _
        Then
            setCancelOrder lOrder
        End If
        
    Next
End With
End Sub

Private Sub matchOrder(ByVal pOrder As Order)
With mMarketTable(pOrder.marketIndex)
    Select Case pOrder.orderType
    Case TradeBuild.OrderTypes.OrderTypeMarket, _
        TradeBuild.OrderTypes.OrderTypeStop
        setFillOrder pOrder, IIf(pOrder.action = OrderActions.ActionBuy, .ask, .bid), pOrder.quantity
    Case TradeBuild.OrderTypes.OrderTypeLimit, _
        TradeBuild.OrderTypes.OrderTypeStopLimit
        If pOrder.action = OrderActions.ActionBuy And _
            pOrder.limitPrice >= .ask _
        Then
            setFillOrder pOrder, .ask, pOrder.quantity
        ElseIf pOrder.action = OrderActions.ActionSell And _
            pOrder.limitPrice <= .bid _
        Then
            setFillOrder pOrder, .bid, pOrder.quantity
        End If
    End Select
End With
End Sub

Private Sub matchOrderTrade(ByVal pOrder As Order, _
                            ByVal trade As Double)
With mMarketTable(pOrder.marketIndex)
    Select Case pOrder.orderType
    Case TradeBuild.OrderTypes.OrderTypeMarket, _
        TradeBuild.OrderTypes.OrderTypeStop
        setFillOrder pOrder, IIf(pOrder.action = OrderActions.ActionBuy, .ask, .bid), pOrder.quantity
    Case TradeBuild.OrderTypes.OrderTypeLimit, _
        TradeBuild.OrderTypes.OrderTypeStop
        If pOrder.action = OrderActions.ActionBuy And _
            pOrder.limitPrice >= trade _
        Then
            setFillOrder pOrder, pOrder.limitPrice, pOrder.quantity
        ElseIf pOrder.action = OrderActions.ActionSell And _
            pOrder.limitPrice <= trade _
        Then
            setFillOrder pOrder, pOrder.limitPrice, pOrder.quantity
        End If
    End Select
End With
End Sub

Private Sub setCancelOrder(ByVal pOrder As Order)
If pOrder.status = OrderStatuses.OrderStatusCancelled Then Exit Sub

Select Case pOrder.status
Case OrderStatuses.OrderStatusSubmitted, OrderStatuses.OrderStatusPreSubmitted
    pOrder.status = OrderStatuses.OrderStatusCancelling
    RaiseEvent orderStatus(mMarketTable(pOrder.marketIndex).clock.timestamp, _
                            pOrder.id, _
                            OrderStatuses.OrderStatusCancelling, _
                            0, _
                            pOrder.quantity, _
                            0, _
                            pOrder.permId, _
                            pOrder.parentId, _
                            0, _
                            pOrder.clientID)
    If mSimulateExchangeDelays Then
        mTimerList.AddTimerListEntry delay, _
                            ExpiryTimeUnits.RelativeMilliseconds, _
                            pOrder
    End If
Case OrderStatuses.OrderStatusCancelling, OrderStatuses.OrderStatusPendingSubmit
    pOrder.status = OrderStatuses.OrderStatusCancelled
    RaiseEvent orderStatus(mMarketTable(pOrder.marketIndex).clock.timestamp, _
                            pOrder.id, _
                            OrderStatuses.OrderStatusCancelled, _
                            0, _
                            pOrder.quantity, _
                            0, _
                            pOrder.permId, _
                            pOrder.parentId, _
                            0, _
                            pOrder.clientID)
    mMarketTable(pOrder.marketIndex).activeOrders.Remove CStr(pOrder.id)
Case OrderStatuses.OrderStatusCreated
    If Not mSimulateExchangeDelays Then
        pOrder.status = OrderStatuses.OrderStatusCancelled
        RaiseEvent orderStatus(mMarketTable(pOrder.marketIndex).clock.timestamp, _
                                pOrder.id, _
                                OrderStatuses.OrderStatusCancelled, _
                                0, _
                                pOrder.quantity, _
                                0, _
                                pOrder.permId, _
                                pOrder.parentId, _
                                0, _
                                pOrder.clientID)
        mMarketTable(pOrder.marketIndex).pendingOrders.Remove CStr(pOrder.id)
    End If
End Select
End Sub

Private Sub setFillOrder(pOrder As Order, _
                        ByVal fillPrice As Double, _
                        ByVal fillSize As Long)
' This needs enhancing to cater for the case
' where the quantity ordered is greater than the
' current bid/offer size
Dim exec As Execution
Dim lContractSpecifier As contractSpecifier
Dim execId As String
Dim dayHex As String
Dim millisecsHex As String
Dim millisecs As Single
Static lastMillisecs As Single
Dim lFillPrice As Double
Dim lFillSize As Long

If fillPrice = 0 Then Exit Sub ' can occur at start of session before first bid or offer

lFillPrice = fillPrice
lFillSize = fillSize

Set lContractSpecifier = mMarketTable(pOrder.marketIndex).contractSpecifier

Set exec = New Execution
exec.accountNumber = 0
exec.clientID = pOrder.clientID
exec.exchange = lContractSpecifier.exchange
execId = String(23, "0")
dayHex = Hex(CDbl(Date))
millisecs = Int(1000 * Timer)
If millisecs <= lastMillisecs Then
    ' this can happen if more than one order is filled in the same tick
    millisecs = lastMillisecs + 1
    lastMillisecs = millisecs
End If
millisecsHex = Hex(millisecs)
Mid(execId, 9 - Len(dayHex), Len(dayHex)) = dayHex
Mid(execId, 18 - Len(millisecsHex), Len(millisecsHex)) = millisecsHex
Mid(execId, 9, 1) = "."
Mid(execId, 18, 6) = ".01.01"
exec.execId = execId
exec.orderId = pOrder.id
exec.permId = pOrder.permId
exec.price = fillPrice
exec.quantity = pOrder.quantity
exec.side = IIf(pOrder.action = OrderActions.ActionBuy, ExecSides.SideBuy, ExecSides.SideSell)
exec.time = Format(mMarketTable(pOrder.marketIndex).clock.timestamp, "yyyymmdd hh:nn:ss")

RaiseEvent preFill(mMarketTable(pOrder.marketIndex).clock.timestamp, _
                    pOrder.Clone, _
                    lContractSpecifier, _
                    lFillPrice, _
                    lFillSize)
                    
If lFillSize > pOrder.quantity Then lFillSize = pOrder.quantity
If lFillSize < 1 Then lFillSize = 1
If lFillPrice = 0 Then lFillPrice = fillPrice

pOrder.fill lFillSize, lFillPrice

RaiseEvent ExecutionDetails(mMarketTable(pOrder.marketIndex).clock.timestamp, _
                        pOrder.id, _
                        lContractSpecifier, _
                        exec)

RaiseEvent openOrder(mMarketTable(pOrder.marketIndex).clock.timestamp, _
                    lContractSpecifier, _
                    pOrder.Clone)
                    
RaiseEvent orderStatus(mMarketTable(pOrder.marketIndex).clock.timestamp, _
                        pOrder.id, _
                        OrderStatuses.OrderStatusFilled, _
                        pOrder.quantityFilled, _
                        pOrder.quantity, _
                        pOrder.averagePrice, _
                        pOrder.permId, _
                        pOrder.parentId, _
                        pOrder.lastFillPrice, _
                        pOrder.clientID)

handleRelatedOrders pOrder

On Error Resume Next
If pOrder.quantity = 0 Then
    mMarketTable(pOrder.marketIndex).activeOrders.Remove CStr(pOrder.id)
End If
On Error GoTo 0
End Sub

'Private Sub setPendingCancelOrder(order as order)
'If order.status = "pendingcancel" Then Exit Sub
'order.status = "pendingcancel"
'RaiseEvent orderStatus(gettime, _
'                        order.id, _
'                        "pendingcancel", _
'                        0, _
'                        order.quantity, _
'                        0, _
'                        order.permId, _
'                        order.parentId, _
'                        0, _
'                        order.clientId)
'If mSimulateExchangeDelays Then
'    mTimerList.addTimer delay, ExpiryTimeUnits.RelativeMilliseconds, order
'End If
'End Sub

'Private Sub setPendingSubmitOrder(order as order)
'order.status = "pendingsubmit"
'RaiseEvent orderStatus(gettime, _
'                        order.id, _
'                        "pendingsubmit", _
'                        0, _
'                        order.quantity, _
'                        0, _
'                        0, _
'                        order.parentId, _
'                        0, _
'                        order.clientId)
'End Sub

'Private Sub setPresubmitOrder(order as order)
'order.status = "presubmitted"
'RaiseEvent orderStatus(gettime, _
'                        order.id, _
'                        "presubmitted", _
'                        0, _
'                        order.quantity, _
'                        0, _
'                        0, _
'                        order.parentId, _
'                        0, _
'                        order.clientId)
'End Sub

Private Sub setRejectOrder(ByVal pOrder As Order, _
                            message As String)
RaiseEvent errorMessage(mMarketTable(pOrder.marketIndex).clock.timestamp, _
                    pOrder.id, _
                    201, _
                    "Order rejected - reason: " & message)
End Sub

Private Sub setSubmitOrder(pOrder As Order)
Select Case pOrder.orderType
Case TradeBuild.OrderTypes.OrderTypeLimit, TradeBuild.OrderTypes.OrderTypeMarket
    Select Case pOrder.status
    Case OrderStatuses.OrderStatusCreated
        If pOrder.parentId <> 0 Then
            pOrder.status = OrderStatuses.OrderStatusPendingSubmit
        ElseIf mSimulateExchangeDelays Then
            pOrder.status = OrderStatuses.OrderStatusSubmitted
            pOrder.permId = mNextPermId
            mNextPermId = mNextPermId + 1
            mTimerList.AddTimerListEntry delay, ExpiryTimeUnits.RelativeMilliseconds, pOrder
        Else
            pOrder.status = OrderStatuses.OrderStatusPendingSubmit
        End If
    Case OrderStatuses.OrderStatusPendingSubmit
        pOrder.status = OrderStatuses.OrderStatusSubmitted
        pOrder.atExchange = True
        pOrder.permId = mNextPermId
        mNextPermId = mNextPermId + 1
    Case Else
        err.Raise ErrorCodes.InvalidOrderTypeInThisContext, _
                    "TradeBuild.CAdvancedOrderSimulator::setSubmitOrder", _
                    "Order type " & pOrder.orderType & " invalid for submission"
    End Select
Case TradeBuild.OrderTypes.OrderTypeStop, TradeBuild.OrderTypes.OrderTypeStopLimit
    Select Case pOrder.status
    Case OrderStatuses.OrderStatusCreated
        If pOrder.parentId <> 0 Then
            pOrder.status = OrderStatuses.OrderStatusPendingSubmit
        Else
            pOrder.status = OrderStatuses.OrderStatusPreSubmitted
        End If
    Case OrderStatuses.OrderStatusPendingSubmit
        pOrder.status = OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusPreSubmitted
        pOrder.status = OrderStatuses.OrderStatusSubmitted
        pOrder.atExchange = True
        pOrder.permId = mNextPermId
        mNextPermId = mNextPermId + 1
    Case Else
        err.Raise ErrorCodes.InvalidOrderTypeInThisContext, _
                    "TradeBuild.CAdvancedOrderSimulator::setSubmitOrder", _
                    "Order type " & pOrder.orderType & " invalid for submission"
    End Select
        
Case Else
    err.Raise ErrorCodes.InvalidOrderTypeInThisContext, _
                "TradeBuild.CAdvancedOrderSimulator::setSubmitOrder", _
                "Order type " & pOrder.orderType & " invalid for submission"
End Select

RaiseEvent orderStatus(mMarketTable(pOrder.marketIndex).clock.timestamp, _
                        pOrder.id, _
                        pOrder.status, _
                        0, _
                        pOrder.quantity, _
                        0, _
                        pOrder.permId, _
                        pOrder.parentId, _
                        0, _
                        pOrder.clientID)

End Sub

