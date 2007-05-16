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

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements StateChangeListener

'@================================================================================
' Events
'@================================================================================

'Event openOrder( _
'                ByVal Timestamp As Date, _
'                ByVal pContractSpecifier As contractSpecifier, _
'                ByVal pOrder as simulatororder)

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName               As String = "TradeBuild26"
Private Const ModuleName                As String = "AdvancedOrderSimulator"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type MarketTableEntry
    inUse                   As Boolean
    contractSpecifier       As contractSpecifier
    bid                     As Double
    bidSize                 As Long
    ask                     As Double
    AskSize                 As Long
    activeOrders            As Collection
    pendingOrders           As Collection
    Clock                   As Clock
    serviceConsumer         As TradeBuildSP.IOrderSubmissionSrvcConsumer
    simulateExchangeDelays  As Boolean
    readyNotified           As Boolean
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer

Private mMarketTable() As MarketTableEntry

Private mAllOrders As Collection

Private mTimerList As TimerList
Attribute mTimerList.VB_VarHelpID = -1
    ' This is used to simulate delays in submitting orders to the
    ' exchange for execution. When it is time to change an order's status
    ' from pendingsubmit to submitted, the order is added to this list
    ' with a random wait of between 200 and 500 millisecs.

Private mNextBrokerOrderId As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mMarketTable(InitialMaxTickers) As MarketTableEntry
Set mAllOrders = New Collection

Set mTimerList = GetGlobalTimerList

mNextBrokerOrderId = 1
End Sub

Private Sub Class_Terminate()
Debug.Print "AdvancedOrderSimulator terminated"
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change( _
                ev As TWUtilities.StateChangeEvent)
Dim tli As TimerListItem
Dim lOrder As SimulatorOrder

Set tli = ev.source
Set lOrder = tli.data

If ev.State = TimerListItemStates.TimerListItemStateExpired Then
    Select Case lOrder.order.status
    Case OrderStatuses.OrderStatusSubmitted
        lOrder.atExchange = True
        matchOrder lOrder
    Case OrderStatuses.OrderStatusCancelling
        setCancelOrder lOrder
    End Select
Else
    err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "StateChangeListener_Change", _
            "Unexpected TimerListItem status"
End If

End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let clockRate( _
                ByVal marketIndex As Long, _
                ByVal value As Single)
With mMarketTable(marketIndex)
    .Clock.Rate = value
    .simulateExchangeDelays = (value <> 0)
End With
End Property

Friend Property Let commonServiceConsumer(ByVal value As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub bid(ByVal timestamp As Date, _
                ByVal marketIndex As Long, _
                ByVal bid As Double)
                
Dim lOrder As SimulatorOrder

With mMarketTable(marketIndex)
    .Clock.SetTime timestamp
    .bid = bid
    
    If .ask <> 0 And .bid <> 0 And Not .readyNotified Then
        .readyNotified = True
        .serviceConsumer.Ready
    End If
    
    handleCreatedOrders marketIndex
    If Not .simulateExchangeDelays Then handlePendingSubmitOrders marketIndex

    For Each lOrder In .activeOrders
        If lOrder.order.status = OrderStatuses.OrderStatusSubmitted Then
            If lOrder.atExchange Then
                matchOrder lOrder
            End If
        ElseIf lOrder.order.status = OrderStatuses.OrderStatusCancelling Then
            If lOrder.atExchange Then
                matchOrder lOrder
                If lOrder.order.status = OrderStatuses.OrderStatusCancelling Then
                    setCancelOrder lOrder
                End If
            ElseIf Not .simulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

End Sub

Friend Sub ask(ByVal timestamp As Date, _
                ByVal marketIndex As Long, _
                ByVal ask As Double)
                
Dim lOrder As SimulatorOrder

With mMarketTable(marketIndex)
    .Clock.SetTime timestamp
    .ask = ask
    
    If .ask <> 0 And .bid <> 0 And Not .readyNotified Then
        .readyNotified = True
        .serviceConsumer.Ready
    End If
    
    handleCreatedOrders marketIndex
    If Not .simulateExchangeDelays Then handlePendingSubmitOrders marketIndex

    For Each lOrder In .activeOrders
        If lOrder.order.status = OrderStatuses.OrderStatusSubmitted Then
            If lOrder.atExchange Then
                matchOrder lOrder
            End If
        ElseIf lOrder.order.status = OrderStatuses.OrderStatusCancelling Then
            If lOrder.atExchange Then
                matchOrder lOrder
                If lOrder.order.status = OrderStatuses.OrderStatusCancelling Then
                    setCancelOrder lOrder
                End If
            ElseIf Not .simulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

End Sub

Friend Sub cancelOrder(ByVal id As Long)
Dim lOrder As SimulatorOrder
Dim pendingOrder As SimulatorOrder
Dim relatedOrder As SimulatorOrder

On Error Resume Next
Set lOrder = mAllOrders(id)
On Error GoTo 0

If lOrder Is Nothing Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.AdvancedOrderSimulator::cancelOrder", _
                "Unknown order id " & id
End If

With mMarketTable(lOrder.marketIndex)
    
    On Error Resume Next
    Set pendingOrder = .pendingOrders(id)
    On Error GoTo 0
    
    If Not pendingOrder Is Nothing Then
        If lOrder.parentId <> "" Then
            Set relatedOrder = .pendingOrders(lOrder.parentId)
            setCancelOrder relatedOrder
            
            For Each relatedOrder In .pendingOrders
                If relatedOrder.parentId = lOrder.parentId And _
                    relatedOrder.order.id <> lOrder.order.id _
                Then
                    setCancelOrder relatedOrder
                End If
            Next
        Else
            For Each relatedOrder In .pendingOrders
                If relatedOrder.parentId = lOrder.order.id Or _
                    (relatedOrder.ocaGroup = lOrder.ocaGroup And _
                    lOrder.ocaGroup <> "" And _
                    relatedOrder.order.id <> lOrder.order.id) _
                Then
                    setCancelOrder relatedOrder
                End If
            Next
        End If
    Else
        On Error Resume Next
        Set lOrder = .activeOrders(id)
        On Error GoTo 0
        
        If lOrder Is Nothing Then Exit Sub
        
        If lOrder.parentId <> "" Then
            On Error Resume Next
            Set relatedOrder = .activeOrders(lOrder.parentId)
            On Error GoTo 0
            If Not relatedOrder Is Nothing Then
                setCancelOrder relatedOrder
            End If
            For Each relatedOrder In .activeOrders
                If relatedOrder.parentId = lOrder.parentId And _
                    relatedOrder.order.id <> lOrder.order.id _
                Then
                    setCancelOrder relatedOrder
                End If
            Next
        Else
            For Each relatedOrder In .activeOrders
                If relatedOrder.parentId = lOrder.order.id Or _
                    (relatedOrder.ocaGroup = lOrder.ocaGroup And _
                    lOrder.ocaGroup <> "" And _
                    relatedOrder.order.id <> lOrder.order.id) _
                  Then
                    setCancelOrder relatedOrder
                End If
            Next
        End If
    End If

    setCancelOrder lOrder

End With
End Sub

''
' Returns the <code>simulatorOrder</code> with the specified id.
'
' @return       rhe <code>simulatorOrder</code> whose encapsulated order has
'               the id specified
'
' @param pId    the id of the <code>order</code> who's corresponding
'               <code>simulatedOrder</code> is required
'
'@/
Friend Function getOrder( _
                ByVal pId As String) As SimulatorOrder
Set getOrder = mAllOrders(pId)
End Function

Friend Sub placeOrder( _
                ByVal pOrder As order, _
                ByVal parentId As String, _
                ByVal ocaGroup As String, _
                ByVal marketIndex As Long, _
                ByVal transmit As Boolean)

Dim lOrder As SimulatorOrder
Dim specifier As TradeBuild26.contractSpecifier
Dim pendingOrder As SimulatorOrder

Select Case pOrder.orderType
Case TradeBuild26.OrderTypes.OrderTypeMarket, _
    TradeBuild26.OrderTypes.OrderTypeLimit, _
    TradeBuild26.OrderTypes.OrderTypeStop, _
    TradeBuild26.OrderTypes.OrderTypeStopLimit
Case Else
    setRejectOrder pOrder, "Order type " & pOrder.orderType & " not supported by simulator"
    Exit Sub
End Select

If marketIndex < 0 Or marketIndex > UBound(mMarketTable) Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuild26.AdvancedOrderSimulator::placeOrder", _
            "Invalid marketIndex " & marketIndex & " in order " & pOrder.id
End If

Set specifier = pOrder.Contract.specifier
If Not specifier.Equals(mMarketTable(marketIndex).contractSpecifier) Then
    err.Raise ErrorCodes.ErrIllegalStateException, _
            "TradeBuild26.AdvancedOrderSimulator::placeOrder", _
            "Incorrect contract in order " & pOrder.id & ": " & _
            specifier.ToString
End If

On Error Resume Next
Set lOrder = mAllOrders(pOrder.id)
On Error GoTo 0

If lOrder Is Nothing Then
    ' this is a new order
    pOrder.brokerId = mNextBrokerOrderId
    mNextBrokerOrderId = mNextBrokerOrderId + 1
    Set lOrder = New SimulatorOrder
    lOrder.order = pOrder.clone
    lOrder.marketIndex = marketIndex
    lOrder.ocaGroup = ocaGroup
    lOrder.parentId = parentId
    mAllOrders.add lOrder, pOrder.id
ElseIf pOrder.brokerId <> "" Then
    ' the order is being updated
    pOrder.syncToMe lOrder.order
    Exit Sub
Else
    ' the order is being re-placed (typically re-placing a stop or target order
    ' after it was cancelled while trying to cancel an order plex, but the entry
    ' order was subsequently filled)
    mAllOrders.remove (pOrder.id)
    pOrder.status = OrderStatuses.OrderStatusCreated
    pOrder.brokerId = mNextBrokerOrderId
    mNextBrokerOrderId = mNextBrokerOrderId + 1
    Set lOrder = New SimulatorOrder
    lOrder.order = pOrder.clone
    lOrder.marketIndex = marketIndex
    lOrder.ocaGroup = ocaGroup
    lOrder.parentId = parentId
    mAllOrders.add lOrder, pOrder.id
    
End If

With mMarketTable(marketIndex)
    If Not transmit Then
        .pendingOrders.add lOrder, lOrder.order.id
    Else
        ' if this order was previously in the pending list, remove
        ' it now
        On Error Resume Next
        .pendingOrders.remove lOrder.order.id
        On Error GoTo 0
        
        ' now see if there are any other pending orders related
        ' to this one via parentid or oca group
        If lOrder.parentId <> "" Then
            Set pendingOrder = .pendingOrders(lOrder.parentId)
            .activeOrders.add pendingOrder, pendingOrder.order.id
            .pendingOrders.remove pendingOrder.order.id
            setSubmitOrder pendingOrder
            
            ' find any other child orders
            For Each pendingOrder In .pendingOrders
                If pendingOrder.parentId = lOrder.parentId Then
                    .activeOrders.add pendingOrder, pendingOrder.order.id
                    .pendingOrders.remove pendingOrder.order.id
                    setSubmitOrder pendingOrder
                End If
            Next
        Else
            ' find any child orders or related OCA orders
            For Each pendingOrder In .pendingOrders
                If pendingOrder.parentId = lOrder.order.id Or _
                    (pendingOrder.ocaGroup = lOrder.ocaGroup And _
                    lOrder.ocaGroup <> "") _
                Then
                    .activeOrders.add pendingOrder, pendingOrder.order.id
                    .pendingOrders.remove pendingOrder.order.id
                    setSubmitOrder pendingOrder
                End If
            Next
                        
        End If
        
        On Error Resume Next
        .activeOrders.add lOrder, lOrder.order.id
        On Error GoTo 0
        setSubmitOrder lOrder
    End If
End With
End Sub
                    
Friend Function startMarket( _
                ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                ByVal contractSpec As TradeBuildSP.IContractSpecifier) As Long
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

With mMarketTable(startMarket)
    .inUse = True
    Set .contractSpecifier = contractSpec
    Set .serviceConsumer = serviceConsumer
    Set .activeOrders = New Collection
    Set .pendingOrders = New Collection
    Set .Clock = New Clock
    .Clock.SimulateTime True
End With

End Function

Friend Sub stopMarket(ByVal handle As Long)
With mMarketTable(handle)
    .inUse = False
    Set .contractSpecifier = Nothing
    Set .activeOrders = Nothing
    Set .pendingOrders = Nothing
    Set .Clock = Nothing
    Set .serviceConsumer = Nothing
    .simulateExchangeDelays = False
    .ask = 0
    .AskSize = 0
    .bid = 0
    .bidSize = 0
End With
End Sub

Friend Sub trade(ByVal timestamp As Date, _
                ByVal marketIndex As Long, _
                ByVal trade As Double)
                
Dim lOrder As SimulatorOrder


With mMarketTable(marketIndex)
    .Clock.SetTime timestamp
    
    handleCreatedOrders marketIndex
    If Not .simulateExchangeDelays Then handlePendingSubmitOrders marketIndex

    For Each lOrder In .activeOrders
        If lOrder.order.status = OrderStatuses.OrderStatusSubmitted Then
            If lOrder.atExchange Then
                matchOrderTrade lOrder, trade
            End If
        ElseIf lOrder.order.status = OrderStatuses.OrderStatusPreSubmitted And _
            lOrder.parentId = "" _
        Then
            Select Case lOrder.order.orderType
                Case TradeBuild26.OrderTypes.OrderTypeStop
                    If lOrder.order.action = OrderActions.ActionBuy And _
                        trade >= lOrder.order.triggerPrice Then
                        setSubmitOrder lOrder
                    ElseIf lOrder.order.action = OrderActions.ActionSell And _
                        trade <= lOrder.order.triggerPrice Then
                        setSubmitOrder lOrder
                    End If
                Case TradeBuild26.OrderTypes.OrderTypeStopLimit
                    If lOrder.order.action = OrderActions.ActionBuy And _
                        trade >= lOrder.order.triggerPrice Then
                        setSubmitOrder lOrder
                    ElseIf lOrder.order.action = OrderActions.ActionSell And _
                        trade <= lOrder.order.triggerPrice Then
                        setSubmitOrder lOrder
                    End If
            End Select
        ElseIf lOrder.order.status = OrderStatuses.OrderStatusCancelling Then
            If lOrder.atExchange Then
                matchOrder lOrder
                If lOrder.order.status = OrderStatuses.OrderStatusCancelling Then
                    setCancelOrder lOrder
                End If
            ElseIf Not .simulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function delay() As Long
delay = 200 + Fix(Rnd() * 300 + 1)
End Function

Private Sub handleCreatedOrders(ByVal marketid As Long)
' submits any orders which are still in the created state
Dim lOrder As SimulatorOrder
With mMarketTable(marketid)
    For Each lOrder In .activeOrders
        If lOrder.order.status = OrderStatuses.OrderStatusCreated Then
            setSubmitOrder lOrder
        End If
    Next
End With
End Sub

Private Sub handlePendingSubmitOrders(ByVal marketid As Long)
Dim lOrder As SimulatorOrder
With mMarketTable(marketid)
    For Each lOrder In .activeOrders
        If lOrder.order.status = OrderStatuses.OrderStatusPendingSubmit And lOrder.parentId = "" Then
            setSubmitOrder lOrder
        End If
    Next
End With
End Sub

Private Sub handleRelatedOrders(ByVal pOrder As SimulatorOrder)
Dim lOrder As SimulatorOrder

With mMarketTable(pOrder.marketIndex)
    For Each lOrder In .activeOrders
        If lOrder.parentId = pOrder.order.id Then
            lOrder.parentId = ""    ' detach the order from the parent
            
            ' now set the status if necessary
            Select Case lOrder.order.orderType
            Case OrderTypes.OrderTypeLimit, OrderTypes.OrderTypeMarket
                If mMarketTable(pOrder.marketIndex).simulateExchangeDelays Then
                    pOrder.order.status = OrderStatuses.OrderStatusSubmitted
                    mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                                            pOrder.order.status
                    mTimerList.add(pOrder, _
                                    delay, _
                                    expiryTimeUnits.ExpiryTimeUnitMilliseconds).addStateChangeListener Me
                Else
                    pOrder.order.status = OrderStatuses.OrderStatusSubmitted
                    pOrder.atExchange = True
                    mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                                            pOrder.order.status
                End If
            Case Else
                ' the order status must remain unchanged
            End Select
        End If
        If lOrder.parentId = pOrder.parentId And _
            lOrder.order.id <> pOrder.order.id And _
            lOrder.parentId <> "" _
        Then
            setCancelOrder lOrder
        End If
        If lOrder.ocaGroup = pOrder.ocaGroup And _
            pOrder.ocaGroup <> "" And _
            lOrder.order.id <> pOrder.order.id _
        Then
            setCancelOrder lOrder
        End If
        
    Next
End With
End Sub

Private Sub matchOrder(ByVal pOrder As SimulatorOrder)
With mMarketTable(pOrder.marketIndex)
    Select Case pOrder.order.orderType
    Case TradeBuild26.OrderTypes.OrderTypeMarket, _
        TradeBuild26.OrderTypes.OrderTypeStop
        setFillOrder pOrder, IIf(pOrder.order.action = OrderActions.ActionBuy, .ask, .bid), pOrder.order.quantity
    Case TradeBuild26.OrderTypes.OrderTypeLimit, _
        TradeBuild26.OrderTypes.OrderTypeStopLimit
        If pOrder.order.action = OrderActions.ActionBuy And _
            pOrder.order.limitPrice >= .ask _
        Then
            setFillOrder pOrder, .ask, pOrder.order.quantity
        ElseIf pOrder.order.action = OrderActions.ActionSell And _
            pOrder.order.limitPrice <= .bid _
        Then
            setFillOrder pOrder, .bid, pOrder.order.quantity
        End If
    End Select
End With
End Sub

Private Sub matchOrderTrade(ByVal pOrder As SimulatorOrder, _
                            ByVal trade As Double)
With mMarketTable(pOrder.marketIndex)
    Select Case pOrder.order.orderType
    Case TradeBuild26.OrderTypes.OrderTypeMarket, _
        TradeBuild26.OrderTypes.OrderTypeStop
        setFillOrder pOrder, IIf(pOrder.order.action = OrderActions.ActionBuy, .ask, .bid), pOrder.order.quantity
    Case TradeBuild26.OrderTypes.OrderTypeLimit, _
        TradeBuild26.OrderTypes.OrderTypeStop
        If pOrder.order.action = OrderActions.ActionBuy And _
            pOrder.order.limitPrice >= trade _
        Then
            setFillOrder pOrder, pOrder.order.limitPrice, pOrder.order.quantity
        ElseIf pOrder.order.action = OrderActions.ActionSell And _
            pOrder.order.limitPrice <= trade _
        Then
            setFillOrder pOrder, pOrder.order.limitPrice, pOrder.order.quantity
        End If
    End Select
End With
End Sub


' commented out on 22/5/6
'Private Sub removeCancelledOrder(ByVal id As Long)
'' this function is called by the instrumentProcessor when a real order
'' has been canceled, in case the simulator hasn't also cancelled it (because
'' the simulator isn't an exact reflection of reality)
'
'
'Dim lOrder as simulatororder
'Dim relatedOrder as simulatororder
'Dim theOrder as simulatororder
'
'On Error Resume Next
'Set theOrder = mAllOrders(id)
'On Error GoTo 0
'If theOrder Is Nothing Then
'    err.Raise ErrorCodes.UnknownOrderID, _
'                "TradeBuild26.AdvancedOrderSimulator::removeCancelledOrder", _
'                "Unknown order id " & id
'End If
'
'With mMarketTable(lOrder.marketIndex)
'
'    On Error Resume Next
'
'    Set lOrder = .activeOrders(id)
'    If lOrder Is Nothing Then Exit Sub
'
'    On Error GoTo 0
'
'    .activeOrders.Remove id
'
'    If lOrder.parentId <> "" Then
'        For Each relatedOrder In .activeOrders
'            If relatedOrder.parentId = lOrder.parentId And _
'                relatedOrder.id <> lOrder.id _
'            Then
'                setCancelOrder relatedOrder
'            End If
'        Next
'    Else
'        For Each relatedOrder In .activeOrders
'            If relatedOrder.parentId = lOrder.id Or _
'                (relatedOrder.ocaGroup = lOrder.ocaGroup And _
'                lOrder.ocaGroup <> "" And _
'                relatedOrder.id <> lOrder.id) _
'            Then
'                setCancelOrder relatedOrder
'            End If
'        Next
'    End If
'
'End With
'
'End Sub
'
'Private Sub removeFilledOrder(ByVal id As Long)
'' this function is called by the instrumentProcessor when a real order
'' has been filled, in case the simulator hasn't also filled it (because
'' the simulator isn't an exact reflection of reality)
'Dim theOrder as simulatororder
'Dim lOrder as simulatororder
'
'On Error Resume Next
'Set theOrder = mAllOrders(id)
'On Error GoTo 0
'If theOrder Is Nothing Then
'    err.Raise ErrorCodes.UnknownOrderID, _
'                "TradeBuild26.AdvancedOrderSimulator::removeFilledOrder", _
'                "Unknown order id " & id
'End If
'
'With mMarketTable(lOrder.marketIndex)
'
'    On Error Resume Next
'
'    Set lOrder = .activeOrders(id)
'    If lOrder Is Nothing Then Exit Sub
'
'    On Error GoTo 0
'
'    handleRelatedOrders lOrder
'    .activeOrders.Remove id
'End With
'End Sub
'
'Private Sub removeOrder(ByVal id As Long)
'Dim lOrder as simulatororder
'
'On Error Resume Next
'Set lOrder = mAllOrders(id)
'On Error GoTo 0
'If lOrder Is Nothing Then
'    err.Raise ErrorCodes.UnknownOrderID, _
'                "TradeBuild26.AdvancedOrderSimulator::removeOrder", _
'                "Unknown order id " & id
'End If
'
'With mMarketTable(lOrder.marketIndex)
'    On Error Resume Next
'    .activeOrders.Remove id
'End With
'End Sub

Private Sub setCancelOrder(ByVal pOrder As SimulatorOrder)
If pOrder.order.status = OrderStatuses.OrderStatusCancelled Then Exit Sub

Select Case pOrder.order.status
Case OrderStatuses.OrderStatusSubmitted, OrderStatuses.OrderStatusPreSubmitted
    pOrder.order.status = OrderStatuses.OrderStatusCancelling
    mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                            OrderStatuses.OrderStatusCancelling
    If mMarketTable(pOrder.marketIndex).simulateExchangeDelays Then
        mTimerList.add(pOrder, _
                            delay, _
                            expiryTimeUnits.ExpiryTimeUnitMilliseconds).addStateChangeListener Me
    End If
Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusPendingSubmit, _
        OrderStatuses.OrderStatusRejected
    pOrder.order.status = OrderStatuses.OrderStatusCancelled
    mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                            OrderStatuses.OrderStatusCancelled
    mMarketTable(pOrder.marketIndex).activeOrders.remove pOrder.order.id
End Select
End Sub

Private Sub setFillOrder(pOrder As SimulatorOrder, _
                        ByVal fillPrice As Double, _
                        ByVal fillSize As Long)
' This needs enhancing to cater for the case
' where the quantity ordered is greater than the
' current bid/offer size
Dim exec As TradeBuildSP.IExecution
Dim lContractSpecifier As contractSpecifier
Dim ExecId As String
Dim dayHex As String
Dim millisecsHex As String
Dim millisecs As Long
Static lastMillisecs As Long
Dim lFillPrice As Double
Dim lFillSize As Long

If fillPrice = 0 Then Exit Sub ' can occur at start of session before first bid or offer

lFillPrice = fillPrice
lFillSize = fillSize

Set lContractSpecifier = mMarketTable(pOrder.marketIndex).contractSpecifier

Set exec = mCommonServiceConsumer.NewExecution
exec.contractSpecifier = lContractSpecifier
exec.accountIdentifier = ""
exec.fillingExchange = lContractSpecifier.exchange
ExecId = String(23, "0")
dayHex = Hex(CDbl(Date))
millisecs = Int(1000 * Timer)
If millisecs <= lastMillisecs Then
    ' this can happen if more than one order is filled in the same tick
    millisecs = lastMillisecs + 1
End If
lastMillisecs = millisecs
millisecsHex = Hex(millisecs)
Mid(ExecId, 9 - Len(dayHex), Len(dayHex)) = dayHex
Mid(ExecId, 18 - Len(millisecsHex), Len(millisecsHex)) = millisecsHex
Mid(ExecId, 9, 1) = "."
Mid(ExecId, 18, 6) = ".01.01"
exec.ExecId = ExecId
exec.isLiquidation = False
exec.orderBrokerId = pOrder.order.brokerId
exec.orderTradeBuildId = pOrder.order.id
exec.action = pOrder.order.action
exec.time = mMarketTable(pOrder.marketIndex).Clock.timestamp

mMarketTable(pOrder.marketIndex).serviceConsumer.preFill pOrder.order.id, _
                    lFillPrice, _
                    lFillSize
                    
If lFillSize > pOrder.order.quantityRemaining Then lFillSize = pOrder.order.quantityRemaining
If lFillSize < 1 Then lFillSize = fillSize
If lFillPrice = 0 Then lFillPrice = fillPrice

exec.price = lFillPrice
exec.quantity = lFillSize

pOrder.order.quantityRemaining = pOrder.order.quantityRemaining - lFillSize
If pOrder.order.quantityRemaining = 0 Then pOrder.order.status = OrderStatuses.OrderStatusFilled

mMarketTable(pOrder.marketIndex).serviceConsumer.executionDetails _
                        lContractSpecifier, _
                        exec

mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                        pOrder.order.status

handleRelatedOrders pOrder

On Error Resume Next
If pOrder.order.quantityRemaining = 0 Then
    mMarketTable(pOrder.marketIndex).activeOrders.remove pOrder.order.id
End If
On Error GoTo 0
End Sub

'Private Sub setPendingCancelOrder(order as simulatororder)
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

'Private Sub setPendingSubmitOrder(order as simulatororder)
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

'Private Sub setPresubmitOrder(order as simulatororder)
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

Private Sub setRejectOrder(ByVal pOrder As SimulatorOrder, _
                            reason As String)
mMarketTable(pOrder.marketIndex).serviceConsumer.rejectOrder pOrder.order.id, _
                                        "Order rejected - reason: " & reason
End Sub

Private Sub setSubmitOrder(pOrder As SimulatorOrder)
Select Case pOrder.order.orderType
Case TradeBuild26.OrderTypes.OrderTypeLimit, TradeBuild26.OrderTypes.OrderTypeMarket
    Select Case pOrder.order.status
    Case OrderStatuses.OrderStatusCreated
        If pOrder.parentId <> "" Then
            ' this is an attached order. Set the status to pending submit and
            ' then move it on at the next tick
            pOrder.order.status = OrderStatuses.OrderStatusPendingSubmit
            mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                                    pOrder.order.status
        Else
            ' this is a parent order, a free standing order, or part of an oca group.
            ' We want to move it to status submitted, either at the next tick or
            ' using exchange submission delay
            If mMarketTable(pOrder.marketIndex).simulateExchangeDelays Then
                pOrder.order.status = OrderStatuses.OrderStatusSubmitted
                mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                                        pOrder.order.status
                mTimerList.add(pOrder, _
                                delay, _
                                expiryTimeUnits.ExpiryTimeUnitMilliseconds).addStateChangeListener Me
            Else
                pOrder.order.status = OrderStatuses.OrderStatusPendingSubmit
                mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                                        pOrder.order.status
            End If
        End If
    Case OrderStatuses.OrderStatusPendingSubmit
        If pOrder.parentId <> "" Then
            ' this is an attached order. Set the status to presubmitted
            ' then move it on at the next tick
            pOrder.order.status = OrderStatuses.OrderStatusPreSubmitted
            mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                                    pOrder.order.status
        Else
            pOrder.order.status = OrderStatuses.OrderStatusSubmitted
            pOrder.atExchange = True
            mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                                    pOrder.order.status
        End If
    Case Else
        err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "TradeBuild26.AdvancedOrderSimulator::setSubmitOrder", _
                    "Order type " & pOrder.order.orderType & " invalid for submission"
    End Select
Case TradeBuild26.OrderTypes.OrderTypeStop, TradeBuild26.OrderTypes.OrderTypeStopLimit
    Select Case pOrder.order.status
    Case OrderStatuses.OrderStatusCreated
        ' in all cases, we want the order to move to pending submit on this
        ' tick, and to presubmitted on the next
        pOrder.order.status = OrderStatuses.OrderStatusPendingSubmit
        mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                                pOrder.order.status
    Case OrderStatuses.OrderStatusPendingSubmit
        pOrder.order.status = OrderStatuses.OrderStatusPreSubmitted
        mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                                pOrder.order.status
    Case OrderStatuses.OrderStatusPreSubmitted
        ' this case occurs when the condition for the order to be triggered
        ' has been reached
        If mMarketTable(pOrder.marketIndex).simulateExchangeDelays Then
            pOrder.order.status = OrderStatuses.OrderStatusSubmitted
            mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                                    pOrder.order.status
            mTimerList.add(pOrder, _
                            delay, _
                            expiryTimeUnits.ExpiryTimeUnitMilliseconds).addStateChangeListener Me
        Else
            pOrder.order.status = OrderStatuses.OrderStatusSubmitted
            pOrder.atExchange = True
            mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                                    pOrder.order.status
        End If
    Case Else
        err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "TradeBuild26.AdvancedOrderSimulator::setSubmitOrder", _
                    "Order type " & pOrder.order.orderType & " invalid for submission"
    End Select
        
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.AdvancedOrderSimulator::setSubmitOrder", _
                "Order type " & pOrder.order.orderType & " invalid for submission"
End Select

mMarketTable(pOrder.marketIndex).serviceConsumer.orderStatus pOrder.order.id, _
                        pOrder.order.status

End Sub

