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

'Event OpenOrder( _
'                ByVal Timestamp As Date, _
'                ByVal pContractSpecifier As contractSpecifier, _
'                ByVal pOrder as simulatororder)

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                As String = "AdvancedOrderSimulator"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type MarketTableEntry
    inUse                   As Boolean
    Contract                As Contract
    Bid                     As Double
    BidSize                 As Long
    Ask                     As Double
    AskSize                 As Long
    activeOrders            As Collection
    pendingOrders           As Collection
    Clock                   As Clock
    serviceConsumer         As TradeBuildSP.IOrderSubmissionSrvcConsumer
    simulateExchangeDelays  As Boolean
'    ReadyNotified           As Boolean
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
    ' from pendingsubmit to submitted, the order is Added to this list
    ' with a random wait of between 200 and 500 millisecs.

Private mNextBrokerOrderId As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mMarketTable(3) As MarketTableEntry
Set mAllOrders = New Collection

Set mTimerList = GetGlobalTimerList

mNextBrokerOrderId = 1

Randomize
End Sub

Private Sub Class_Terminate()
Debug.Print "AdvancedOrderSimulator terminated"
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change( _
                ev As StateChangeEvent)
Dim tli As TimerListItem
Dim lOrder As SimulatorOrder

Const ProcName As String = "StateChangeListener_Change"

On Error GoTo Err

Set tli = ev.source
Set lOrder = tli.Data

If ev.State = TimerListItemStates.TimerListItemStateExpired Then
    Select Case lOrder.Order.Status
    Case OrderStatuses.OrderStatusSubmitted
        lOrder.AtExchange = True
        matchOrder lOrder
    Case OrderStatuses.OrderStatusCancelling
        setCancelOrder lOrder
    End Select
Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Unexpected TimerListItem status"
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let ClockRate( _
                ByVal MarketIndex As Long, _
                ByVal value As Single)
Dim latestTimestamp As Date
Const ProcName As String = "ClockRate"

On Error GoTo Err

With mMarketTable(MarketIndex)
    If Not .Clock Is Nothing Then
        latestTimestamp = .Clock.Timestamp
    End If
    Set .Clock = CreateSimulatedClock(value, .Contract.Timezone.StandardName)
    .simulateExchangeDelays = (value <> 0)
End With

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let CommonServiceConsumer(ByVal value As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Bid(ByVal Timestamp As Date, _
                ByVal MarketIndex As Long, _
                ByVal Bid As Double)
                
Dim lOrder As SimulatorOrder

Const ProcName As String = "Bid"

On Error GoTo Err

With mMarketTable(MarketIndex)
    .Clock.setTime Timestamp
    .Bid = Bid
    
    handleCreatedOrders MarketIndex
    If Not .simulateExchangeDelays Then handlePendingSubmitOrders MarketIndex

    For Each lOrder In .activeOrders
        If lOrder.Order.Status = OrderStatuses.OrderStatusSubmitted Then
            If lOrder.AtExchange Then
                matchOrder lOrder
            End If
        ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
            If lOrder.AtExchange Then
                matchOrder lOrder
                If lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
                    setCancelOrder lOrder
                End If
            ElseIf Not .simulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Friend Sub Ask(ByVal Timestamp As Date, _
                ByVal MarketIndex As Long, _
                ByVal Ask As Double)
                
Dim lOrder As SimulatorOrder

Const ProcName As String = "Ask"

On Error GoTo Err

With mMarketTable(MarketIndex)
    .Clock.setTime Timestamp
    .Ask = Ask
    
    handleCreatedOrders MarketIndex
    If Not .simulateExchangeDelays Then handlePendingSubmitOrders MarketIndex

    For Each lOrder In .activeOrders
        If lOrder.Order.Status = OrderStatuses.OrderStatusSubmitted Then
            If lOrder.AtExchange Then
                matchOrder lOrder
            End If
        ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
            If lOrder.AtExchange Then
                matchOrder lOrder
                If lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
                    setCancelOrder lOrder
                End If
            ElseIf Not .simulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Friend Sub CancelOrder(ByVal Id As Long)
Dim lOrder As SimulatorOrder
Dim pendingOrder As SimulatorOrder
Dim relatedOrder As SimulatorOrder

Const ProcName As String = "CancelOrder"

On Error GoTo Err

On Error Resume Next
Set lOrder = mAllOrders(Id)
On Error GoTo Err

If lOrder Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Unknown order id " & Id
End If

With mMarketTable(lOrder.MarketIndex)
    
    On Error Resume Next
    Set pendingOrder = .pendingOrders(Id)
    On Error GoTo Err
    
    If Not pendingOrder Is Nothing Then
        If lOrder.ParentId <> "" Then
            Set relatedOrder = .pendingOrders(lOrder.ParentId)
            setCancelOrder relatedOrder
            
            For Each relatedOrder In .pendingOrders
                If relatedOrder.ParentId = lOrder.ParentId And _
                    relatedOrder.Order.Id <> lOrder.Order.Id _
                Then
                    setCancelOrder relatedOrder
                End If
            Next
        Else
            For Each relatedOrder In .pendingOrders
                If relatedOrder.ParentId = lOrder.Order.Id Or _
                    (relatedOrder.OcaGroup = lOrder.OcaGroup And _
                    lOrder.OcaGroup <> "" And _
                    relatedOrder.Order.Id <> lOrder.Order.Id) _
                Then
                    setCancelOrder relatedOrder
                End If
            Next
        End If
    Else
        On Error Resume Next
        Set lOrder = .activeOrders(Id)
        On Error GoTo Err
        
        If lOrder Is Nothing Then Exit Sub
        
        If lOrder.ParentId <> "" Then
            On Error Resume Next
            Set relatedOrder = .activeOrders(lOrder.ParentId)
            On Error GoTo Err
            If Not relatedOrder Is Nothing Then
                setCancelOrder relatedOrder
            End If
            For Each relatedOrder In .activeOrders
                If relatedOrder.ParentId = lOrder.ParentId And _
                    relatedOrder.Order.Id <> lOrder.Order.Id _
                Then
                    setCancelOrder relatedOrder
                End If
            Next
        Else
            For Each relatedOrder In .activeOrders
                If relatedOrder.ParentId = lOrder.Order.Id Or _
                    (relatedOrder.OcaGroup = lOrder.OcaGroup And _
                    lOrder.OcaGroup <> "" And _
                    relatedOrder.Order.Id <> lOrder.Order.Id) _
                  Then
                    setCancelOrder relatedOrder
                End If
            Next
        End If
    End If

    setCancelOrder lOrder

End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
Friend Function GetOrder( _
                ByVal pId As String) As SimulatorOrder
Const ProcName As String = "GetOrder"

On Error GoTo Err

Set GetOrder = mAllOrders(pId)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub placeOrder( _
                ByVal pOrder As Order, _
                ByVal ParentId As String, _
                ByVal OcaGroup As String, _
                ByVal MarketIndex As Long, _
                ByVal transmit As Boolean)

Dim lOrder As SimulatorOrder
Dim specifier As ContractSpecifier
Dim pendingOrder As SimulatorOrder

Const ProcName As String = "placeOrder"

On Error GoTo Err

If MarketIndex < 0 Or MarketIndex > UBound(mMarketTable) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid marketIndex " & MarketIndex & " in order " & pOrder.Id
End If

Select Case pOrder.OrderType
Case OrderTypes.OrderTypeMarket, _
    OrderTypes.OrderTypeLimit, _
    OrderTypes.OrderTypeStop, _
    OrderTypes.OrderTypeStopLimit
Case Else
    setRejectOrder pOrder, _
                    "Order type " & pOrder.OrderType & " not supported by simulator", _
                    MarketIndex
    Exit Sub
End Select

Set specifier = pOrder.Contract.specifier
If Not specifier.Equals(mMarketTable(MarketIndex).Contract.specifier) Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Incorrect contract in order " & pOrder.Id & ": " & _
            specifier.ToString
End If

On Error Resume Next
Set lOrder = mAllOrders(pOrder.Id)
On Error GoTo Err

If lOrder Is Nothing Then
    ' this is a new order
    pOrder.BrokerId = mNextBrokerOrderId
    mNextBrokerOrderId = mNextBrokerOrderId + 1
    Set lOrder = New SimulatorOrder
    lOrder.Order = pOrder.Clone
    lOrder.MarketIndex = MarketIndex
    lOrder.OcaGroup = OcaGroup
    lOrder.ParentId = ParentId
    mAllOrders.Add lOrder, pOrder.Id
ElseIf pOrder.BrokerId <> "" Then
    ' the order is being updated
    pOrder.syncToMe lOrder.Order
    Exit Sub
Else
    ' the order is being re-placed (typically re-placing a stop or target order
    ' after it was cancelled while trying to cancel an order plex, but the entry
    ' order was subsequently filled)
    mAllOrders.Remove (pOrder.Id)
    pOrder.Status = OrderStatuses.OrderStatusCreated
    pOrder.BrokerId = mNextBrokerOrderId
    mNextBrokerOrderId = mNextBrokerOrderId + 1
    Set lOrder = New SimulatorOrder
    lOrder.Order = pOrder.Clone
    lOrder.MarketIndex = MarketIndex
    lOrder.OcaGroup = OcaGroup
    lOrder.ParentId = ParentId
    mAllOrders.Add lOrder, pOrder.Id
    
End If

With mMarketTable(MarketIndex)
    If Not transmit Then
        .pendingOrders.Add lOrder, lOrder.Order.Id
    Else
        ' if this order was previously in the pending list, Remove
        ' it now
        On Error Resume Next
        .pendingOrders.Remove lOrder.Order.Id
        On Error GoTo Err
        
        ' now see if there are any other pending orders related
        ' to this one via parentid or oca group
        If lOrder.ParentId <> "" Then
            Set pendingOrder = .pendingOrders(lOrder.ParentId)
            .activeOrders.Add pendingOrder, pendingOrder.Order.Id
            .pendingOrders.Remove pendingOrder.Order.Id
            setSubmitOrder pendingOrder
            
            ' find any other child orders
            For Each pendingOrder In .pendingOrders
                If pendingOrder.ParentId = lOrder.ParentId Then
                    .activeOrders.Add pendingOrder, pendingOrder.Order.Id
                    .pendingOrders.Remove pendingOrder.Order.Id
                    setSubmitOrder pendingOrder
                End If
            Next
        Else
            ' find any child orders or related OCA orders
            For Each pendingOrder In .pendingOrders
                If pendingOrder.ParentId = lOrder.Order.Id Or _
                    (pendingOrder.OcaGroup = lOrder.OcaGroup And _
                    lOrder.OcaGroup <> "") _
                Then
                    .activeOrders.Add pendingOrder, pendingOrder.Order.Id
                    .pendingOrders.Remove pendingOrder.Order.Id
                    setSubmitOrder pendingOrder
                End If
            Next
                        
        End If
        
        On Error Resume Next
        .activeOrders.Add lOrder, lOrder.Order.Id
        On Error GoTo Err
        setSubmitOrder lOrder
    End If
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub
                    
Friend Function startMarket( _
                ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                ByVal pContract As Contract) As Long
Dim i As Long

Const ProcName As String = "startMarket"

On Error GoTo Err

startMarket = -1

For i = 0 To UBound(mMarketTable)
    If (Not mMarketTable(i).inUse) Then
        startMarket = i
        Exit For
    End If
Next

If startMarket = -1 Then
    startMarket = UBound(mMarketTable) + 1
    ReDim Preserve mMarketTable(2 * (UBound(mMarketTable) + 1) - 1) As MarketTableEntry
End If

With mMarketTable(startMarket)
    .inUse = True
    Set .Contract = pContract
    Set .serviceConsumer = serviceConsumer
    Set .activeOrders = New Collection
    Set .pendingOrders = New Collection
    Set .Clock = CreateSimulatedClock(1, pContract.Timezone.StandardName)
End With

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Friend Sub stopMarket(ByVal Handle As Long)
Const ProcName As String = "stopMarket"

On Error GoTo Err

With mMarketTable(Handle)
    .inUse = False
    Set .Contract = Nothing
    Set .activeOrders = Nothing
    Set .pendingOrders = Nothing
    Set .Clock = Nothing
    Set .serviceConsumer = Nothing
    .simulateExchangeDelays = False
    .Ask = 0
    .AskSize = 0
    .Bid = 0
    .BidSize = 0
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Trade(ByVal Timestamp As Date, _
                ByVal MarketIndex As Long, _
                ByVal Trade As Double)
                
Dim lOrder As SimulatorOrder


Const ProcName As String = "Trade"

On Error GoTo Err

With mMarketTable(MarketIndex)
    .Clock.setTime Timestamp
    
    handleCreatedOrders MarketIndex
    If Not .simulateExchangeDelays Then handlePendingSubmitOrders MarketIndex

    For Each lOrder In .activeOrders
        If lOrder.Order.Status = OrderStatuses.OrderStatusSubmitted Then
            If lOrder.AtExchange Then
                matchOrderTrade lOrder, Trade
            End If
        ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusPreSubmitted And _
            lOrder.ParentId = "" _
        Then
            Select Case lOrder.Order.OrderType
                Case OrderTypes.OrderTypeStop
                    If lOrder.Order.action = OrderActions.ActionBuy And _
                        Trade >= lOrder.Order.TriggerPrice Then
                        setSubmitOrder lOrder
                    ElseIf lOrder.Order.action = OrderActions.ActionSell And _
                        Trade <= lOrder.Order.TriggerPrice Then
                        setSubmitOrder lOrder
                    End If
                Case OrderTypes.OrderTypeStopLimit
                    If lOrder.Order.action = OrderActions.ActionBuy And _
                        Trade >= lOrder.Order.TriggerPrice Then
                        setSubmitOrder lOrder
                    ElseIf lOrder.Order.action = OrderActions.ActionSell And _
                        Trade <= lOrder.Order.TriggerPrice Then
                        setSubmitOrder lOrder
                    End If
            End Select
        ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
            If lOrder.AtExchange Then
                matchOrder lOrder
                If lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
                    setCancelOrder lOrder
                End If
            ElseIf Not .simulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

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
Const ProcName As String = "handleCreatedOrders"

On Error GoTo Err

With mMarketTable(marketid)
    For Each lOrder In .activeOrders
        If lOrder.Order.Status = OrderStatuses.OrderStatusCreated Then
            setSubmitOrder lOrder
        End If
    Next
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub handlePendingSubmitOrders(ByVal marketid As Long)
Dim lOrder As SimulatorOrder
Const ProcName As String = "handlePendingSubmitOrders"

On Error GoTo Err

With mMarketTable(marketid)
    For Each lOrder In .activeOrders
        If lOrder.Order.Status = OrderStatuses.OrderStatusPendingSubmit And lOrder.ParentId = "" Then
            setSubmitOrder lOrder
        End If
    Next
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub handleRelatedOrders(ByVal pOrder As SimulatorOrder)
Dim lOrder As SimulatorOrder

Const ProcName As String = "handleRelatedOrders"

On Error GoTo Err

With mMarketTable(pOrder.MarketIndex)
    For Each lOrder In .activeOrders
        If lOrder.ParentId = pOrder.Order.Id Then
            lOrder.ParentId = ""    ' detach the order from the parent
            
            ' now set the status if necessary
            Select Case lOrder.Order.OrderType
            Case OrderTypes.OrderTypeLimit, OrderTypes.OrderTypeMarket
                If mMarketTable(pOrder.MarketIndex).simulateExchangeDelays Then
                    setOrderStatus pOrder, OrderStatuses.OrderStatusSubmitted
                    mTimerList.Add(pOrder, _
                                    delay, _
                                    expiryTimeUnits.ExpiryTimeUnitMilliseconds).AddStateChangeListener Me
                Else
                    pOrder.AtExchange = True
                    setOrderStatus pOrder, OrderStatuses.OrderStatusSubmitted
                End If
            Case Else
                ' the order status must remain unchanged
            End Select
        End If
        If lOrder.ParentId = pOrder.ParentId And _
            lOrder.Order.Id <> pOrder.Order.Id And _
            lOrder.ParentId <> "" _
        Then
            setCancelOrder lOrder
        End If
        If lOrder.OcaGroup = pOrder.OcaGroup And _
            pOrder.OcaGroup <> "" And _
            lOrder.Order.Id <> pOrder.Order.Id _
        Then
            setCancelOrder lOrder
        End If
        
    Next
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub matchOrder(ByVal pOrder As SimulatorOrder)
Const ProcName As String = "matchOrder"

On Error GoTo Err

With mMarketTable(pOrder.MarketIndex)
    Select Case pOrder.Order.OrderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeStop
        setFillOrder pOrder, IIf(pOrder.Order.action = OrderActions.ActionBuy, .Ask, .Bid), pOrder.Order.Quantity
    Case OrderTypes.OrderTypeLimit, _
        OrderTypes.OrderTypeStopLimit
        If pOrder.Order.action = OrderActions.ActionBuy And _
            pOrder.Order.LimitPrice >= .Ask _
        Then
            setFillOrder pOrder, .Ask, pOrder.Order.Quantity
        ElseIf pOrder.Order.action = OrderActions.ActionSell And _
            pOrder.Order.LimitPrice <= .Bid _
        Then
            setFillOrder pOrder, .Bid, pOrder.Order.Quantity
        End If
    End Select
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub matchOrderTrade(ByVal pOrder As SimulatorOrder, _
                            ByVal Trade As Double)
Const ProcName As String = "matchOrderTrade"

On Error GoTo Err

With mMarketTable(pOrder.MarketIndex)
    Select Case pOrder.Order.OrderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeStop
        setFillOrder pOrder, IIf(pOrder.Order.action = OrderActions.ActionBuy, .Ask, .Bid), pOrder.Order.Quantity
    Case OrderTypes.OrderTypeLimit, _
        OrderTypes.OrderTypeStop
        If pOrder.Order.action = OrderActions.ActionBuy And _
            pOrder.Order.LimitPrice >= Trade _
        Then
            setFillOrder pOrder, pOrder.Order.LimitPrice, pOrder.Order.Quantity
        ElseIf pOrder.Order.action = OrderActions.ActionSell And _
            pOrder.Order.LimitPrice <= Trade _
        Then
            setFillOrder pOrder, pOrder.Order.LimitPrice, pOrder.Order.Quantity
        End If
    End Select
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub notifyOrderStatus( _
                ByVal pOrder As SimulatorOrder)
mMarketTable(pOrder.MarketIndex).serviceConsumer.OrderStatusReport gCreateOrderStatusReport(pOrder.Order.Id, pOrder.Order.Status)
End Sub


Private Sub setCancelOrder(ByVal pOrder As SimulatorOrder)
Const ProcName As String = "setCancelOrder"

On Error GoTo Err

If pOrder.Order.Status = OrderStatuses.OrderStatusCancelled Then Exit Sub

Select Case pOrder.Order.Status
Case OrderStatuses.OrderStatusSubmitted, OrderStatuses.OrderStatusPreSubmitted
    setOrderStatus pOrder, OrderStatuses.OrderStatusCancelling
    If mMarketTable(pOrder.MarketIndex).simulateExchangeDelays Then
        mTimerList.Add(pOrder, _
                            delay, _
                            expiryTimeUnits.ExpiryTimeUnitMilliseconds).AddStateChangeListener Me
    End If
Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusPendingSubmit, _
        OrderStatuses.OrderStatusRejected
    setOrderStatus pOrder, OrderStatuses.OrderStatusCancelled
    mMarketTable(pOrder.MarketIndex).activeOrders.Remove pOrder.Order.Id
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setFillOrder(pOrder As SimulatorOrder, _
                        ByVal fillPrice As Double, _
                        ByVal fillSize As Long)
' This needs enhancing to cater for the case
' where the quantity ordered is greater than the
' current Bid/offer Size
'Dim exec As TradeBuildSP.IExecution
Dim lContractSpecifier As ContractSpecifier
Dim ExecId As String
Dim dayHex As String
Dim millisecsHex As String
Dim millisecs As Long
Static lastMillisecs As Long
Dim lFillPrice As Double
Dim lFillSize As Long
Dim exec As Execution

Const ProcName As String = "setFillOrder"

On Error GoTo Err

If fillPrice = 0 Then Exit Sub ' can occur at start of Session before first Bid or offer

lFillPrice = fillPrice
lFillSize = fillSize

Set lContractSpecifier = mMarketTable(pOrder.MarketIndex).Contract.specifier

'Set exec = mCommonServiceConsumer.NewExecution
'exec.contractSpecifier = lContractSpecifier
'exec.accountIdentifier = ""
'exec.fillingExchange = lContractSpecifier.exchange
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
'exec.execId = execId
'exec.isLiquidation = False
'exec.orderBrokerId = pOrder.order.brokerId
'exec.orderTradeBuildId = pOrder.order.id
'exec.Action = pOrder.order.Action
'exec.time = mMarketTable(pOrder.marketIndex).Clock.Timestamp

mMarketTable(pOrder.MarketIndex).serviceConsumer.PreFill pOrder.Order.Id, _
                    lFillPrice, _
                    lFillSize
                    
If lFillSize > pOrder.Order.QuantityRemaining Then lFillSize = pOrder.Order.QuantityRemaining
If lFillSize < 1 Then lFillSize = fillSize
If lFillPrice = 0 Then lFillPrice = fillPrice

'exec.price = lFillPrice
'exec.quantity = lFillSize

pOrder.Order.QuantityRemaining = pOrder.Order.QuantityRemaining - lFillSize
If pOrder.Order.QuantityRemaining = 0 Then pOrder.Order.Status = OrderStatuses.OrderStatusFilled

Set exec = New Execution
exec.Initialise lContractSpecifier, _
                "", _
                pOrder.Order.action, _
                ExecId, _
                lContractSpecifier.Exchange, _
                False, _
                pOrder.Order.BrokerId, _
                pOrder.Order.Id, _
                lFillPrice, _
                lFillSize, _
                mMarketTable(pOrder.MarketIndex).Clock.Timestamp, _
                mMarketTable(pOrder.MarketIndex).Clock.Timezone.DisplayName
exec.IsSimulated = True
mMarketTable(pOrder.MarketIndex).serviceConsumer.ExecutionReport exec


notifyOrderStatus pOrder

handleRelatedOrders pOrder

On Error Resume Next
If pOrder.Order.QuantityRemaining = 0 Then
    mMarketTable(pOrder.MarketIndex).activeOrders.Remove pOrder.Order.Id
End If
On Error GoTo Err

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setOrderStatus( _
                ByVal pOrder As SimulatorOrder, _
                ByVal pStatus As OrderStatuses)

pOrder.Order.Status = pStatus
notifyOrderStatus pOrder
End Sub

'Private Sub setPendingCancelOrder(order as simulatororder)
'If order.status = "pendingcancel" Then Exit Sub
'order.status = "pendingcancel"
'RaiseEvent OrderStatus(gettime, _
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
'    mTimerList.AddTimer delay, ExpiryTimeUnits.RelativeMilliseconds, order
'End If
'End Sub

'Private Sub setPendingSubmitOrder(order as simulatororder)
'order.status = "pendingsubmit"
'RaiseEvent OrderStatus(gettime, _
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
'RaiseEvent OrderStatus(gettime, _
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
                            reason As String, _
                            ByVal MarketIndex As Long)
Const ProcName As String = "setRejectOrder"

On Error GoTo Err

mMarketTable(MarketIndex).serviceConsumer.RejectOrder pOrder.Id, _
                                        "Order rejected - reason: " & reason

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setSubmitOrder(pOrder As SimulatorOrder)
Const ProcName As String = "setSubmitOrder"

On Error GoTo Err

Select Case pOrder.Order.OrderType
Case OrderTypes.OrderTypeLimit, OrderTypes.OrderTypeMarket
    Select Case pOrder.Order.Status
    Case OrderStatuses.OrderStatusCreated
        If pOrder.ParentId <> "" Then
            ' this is an attached order. Set the status to pending submit and
            ' then move it on at the next tick
            setOrderStatus pOrder, OrderStatuses.OrderStatusPendingSubmit
        Else
            ' this is a parent order, a free standing order, or part of an oca group.
            ' We want to move it to status submitted, either at the next tick or
            ' using exchange submission delay
            If mMarketTable(pOrder.MarketIndex).simulateExchangeDelays Then
                setOrderStatus pOrder, OrderStatuses.OrderStatusSubmitted
                mTimerList.Add(pOrder, _
                                delay, _
                                expiryTimeUnits.ExpiryTimeUnitMilliseconds).AddStateChangeListener Me
            Else
                setOrderStatus pOrder, OrderStatuses.OrderStatusPendingSubmit
            End If
        End If
    Case OrderStatuses.OrderStatusPendingSubmit
        If pOrder.ParentId <> "" Then
            ' this is an attached order. Set the status to presubmitted
            ' then move it on at the next tick
            setOrderStatus pOrder, OrderStatuses.OrderStatusPreSubmitted
        Else
            pOrder.AtExchange = True
            setOrderStatus pOrder, OrderStatuses.OrderStatusSubmitted
        End If
    Case Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Order type " & pOrder.Order.OrderType & " invalid for submission"
    End Select
Case OrderTypes.OrderTypeStop, OrderTypes.OrderTypeStopLimit
    Select Case pOrder.Order.Status
    Case OrderStatuses.OrderStatusCreated
        ' in all cases, we want the order to move to pending submit on this
        ' tick, and to presubmitted on the next
        setOrderStatus pOrder, OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusPendingSubmit
        setOrderStatus pOrder, OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusPreSubmitted
        ' this case occurs when the condition for the order to be triggered
        ' has been reached
        If mMarketTable(pOrder.MarketIndex).simulateExchangeDelays Then
            setOrderStatus pOrder, OrderStatuses.OrderStatusSubmitted
            mTimerList.Add(pOrder, _
                            delay, _
                            expiryTimeUnits.ExpiryTimeUnitMilliseconds).AddStateChangeListener Me
        Else
            pOrder.AtExchange = True
            setOrderStatus pOrder, OrderStatuses.OrderStatusSubmitted
        End If
    Case Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Order type " & pOrder.Order.OrderType & " invalid for submission"
    End Select
        
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Order type " & pOrder.Order.OrderType & " invalid for submission"
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

