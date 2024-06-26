VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarketSimulator"
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

Implements IGenericTickListener
Implements IDeferredAction

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MarketSimulator"

'@================================================================================
' Member variables
'@================================================================================

Private mBid                                        As Double
Private mAsk                                        As Double
Private mActiveOrders                               As New EnumerableCollection
Private mPendingOrders                              As New EnumerableCollection
Private mSimulateExchangeDelays                     As Boolean
Private mDataSource                                 As IMarketDataSource

Private mAllOrders                                  As New EnumerableCollection

Private mUsageCount                                 As Long

Private mKey                                        As String

' Holds orders that are placed by listeners while processing a
' notification. When the event handler that caused the notification
' completes, the orders in here are queued. This is to avoid mActiveOrders
' being modified while it is being enumerated.
Private mOrdersToBeQueued                           As New EnumerableCollection

Private mEnumerationInProgess                       As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Randomize
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)

End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Select Case ev.Tick.TickType
Case TickTypes.TickTypeAsk
    processAsk ev.Tick.Timestamp, ev.Tick.Price
Case TickTypes.TickTypeBid
    processBid ev.Tick.Timestamp, ev.Tick.Price
Case TickTypes.TickTypeClosePrice
Case TickTypes.TickTypeHighPrice
Case TickTypes.TickTypeLowPrice
Case TickTypes.TickTypeMarketDepth
Case TickTypes.TickTypeMarketDepthReset
Case TickTypes.TickTypeOpenInterest
Case TickTypes.TickTypeOpenPrice
Case TickTypes.TickTypeTrade
    processTrade ev.Tick.Timestamp, ev.Tick.Price
Case TickTypes.TickTypeVolume
End Select

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal pData As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

mEnumerationInProgess = True

Dim lOrder As SimulatorOrder
Set lOrder = pData

Select Case lOrder.Order.Status
Case OrderStatuses.OrderStatusSubmitted
    lOrder.AtExchange = True
    matchOrder lOrder
Case OrderStatuses.OrderStatusCancelling
    setCancelOrder lOrder
End Select

mEnumerationInProgess = False
queueOrders

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Key() As String
Key = mKey
End Property

Public Property Get PermittedOrderProperties() As IPermittedOrderProperties
Const ProcName As String = "PermittedOrderProperties"
On Error GoTo Err

Dim lPermittedOrderProperties As New PermittedOrderProperties
lPermittedOrderProperties.PermittedOrderAttributes = _
                            OrderAttributes.OrderAttAction + _
                            OrderAttributes.OrderAttLimitPrice + _
                            OrderAttributes.OrderAttOrderType + _
                            OrderAttributes.OrderAttOriginatorRef + _
                            OrderAttributes.OrderAttQuantity + _
                            OrderAttributes.OrderAttStopTriggerMethod + _
                            OrderAttributes.OrderAttTimeInForce + _
                            OrderAttributes.OrderAttTriggerPrice

lPermittedOrderProperties.PermittedOrderTifs = OrderTIFDay + OrderTIFGoodTillCancelled

lPermittedOrderProperties.PermittedOrderTypes = _
                        OrderTypes.OrderTypeMarket + _
                        OrderTypes.OrderTypeLimit + _
                        OrderTypes.OrderTypeStop + _
                        OrderTypes.OrderTypeStopLimit

lPermittedOrderProperties.PermittedStopTriggerMethods = OrderStopTriggerDefault + OrderStopTriggerLast

Set PermittedOrderProperties = lPermittedOrderProperties

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let UsageCount(ByVal Value As Long)
mUsageCount = Value
End Property

Friend Property Get UsageCount() As Long
UsageCount = mUsageCount
End Property


'@================================================================================
' Methods
'@================================================================================

Public Sub CancelOrder(ByVal Id As String)
Const ProcName As String = "CancelOrder"
On Error GoTo Err

AssertArgument mAllOrders.Contains(Id), "Unknown order id " & Id

Dim lOrder As SimulatorOrder
Set lOrder = mAllOrders(Id)

cancelRelatedOrders mPendingOrders.Enumerator, lOrder
cancelRelatedOrders mActiveOrders.Enumerator, lOrder
setCancelOrder lOrder

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetOcaGroup( _
                ByVal pId As String) As String
Const ProcName As String = "GetOcaGroup"
On Error GoTo Err

Dim lOrder As SimulatorOrder
Set lOrder = mAllOrders(pId)
GetOcaGroup = lOrder.OcaGroup

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pKey As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mDataSource = pDataSource
mDataSource.PriorityTickListener = Me
mKey = pKey

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PlaceOrder( _
                ByVal pOrder As IOrder, _
                ByVal pParentId As String, _
                ByVal pOcaGroup As String, _
                ByVal pTransmit As Boolean, _
                ByVal pListener As IOrderSubmissionListener)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

Select Case pOrder.OrderType
Case OrderTypes.OrderTypeMarket, _
    OrderTypes.OrderTypeLimit, _
    OrderTypes.OrderTypeStop, _
    OrderTypes.OrderTypeStopLimit
Case Else
    setRejectOrder createSimulatorOrder(pOrder, pParentId, pOcaGroup, pTransmit, pListener), _
                    ErrorCodes.ErrUnsupportedOperationException, _
                    "Order type " & pOrder.OrderType & " not supported by simulator"
    Exit Sub
End Select

Dim lSimulatorOrder As SimulatorOrder

If Not mAllOrders.Contains(pOrder.Id) Then
    ' this is a new order
    Set lSimulatorOrder = createSimulatorOrder(pOrder, pParentId, pOcaGroup, pTransmit, pListener)
ElseIf pOrder.BrokerId = "" Then
    ' the order is being re-placed (typically re-placing a stoploss or target order
    ' after it was cancelled while trying to cancel a bracket order, but the entry
    ' order was subsequently filled)
    mAllOrders.Remove pOrder.Id
    Set lSimulatorOrder = createSimulatorOrder(pOrder, pParentId, pOcaGroup, pTransmit, pListener)
Else
    ' the order is being updated
    Set lSimulatorOrder = mAllOrders(pOrder.Id)
    SyncToOrder lSimulatorOrder.Order, pOrder
    lSimulatorOrder.Transmit = pTransmit
End If

queueOrder lSimulatorOrder

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveOrdersForListener(ByVal pListener As IOrderSubmissionListener)
Const ProcName As String = "RemoveOrdersForListener"
On Error GoTo Err

If mAllOrders.Count = 0 Then Exit Sub

Dim en As Enumerator
Set en = mAllOrders.Enumerator

Do While en.MoveNext
    Dim lOrder As SimulatorOrder
    Set lOrder = en.Current
    If lOrder.Listener Is pListener Then en.Remove
Loop

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub
                    
'@================================================================================
' Helper Functions
'@================================================================================

Private Sub cancelRelatedOrders(ByVal en As Enumerator, ByVal pOrder As SimulatorOrder)
Const ProcName As String = "cancelRelatedOrders"
On Error GoTo Err

Do While en.MoveNext
    Dim lOrder As SimulatorOrder
    Set lOrder = en.Current
    If lOrder.IsFinished Then
        en.Remove
    ElseIf pOrder.IsRelatedTo(lOrder) Then
        setCancelOrder lOrder
    End If
Loop

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function createExecId() As String
Static sLastMillisecs As Long

Dim lExecId As String: lExecId = String(23, "0")
Dim dayHex As String: dayHex = Hex(CDbl(Date))
Dim millisecs As Long: millisecs = Int(1000 * Timer)
If millisecs <= sLastMillisecs Then
    ' this can happen if more than one order is filled in the same tick
    millisecs = sLastMillisecs + 1
End If
sLastMillisecs = millisecs

Dim millisecsHex As String: millisecsHex = Hex(millisecs)
Mid(lExecId, 9 - Len(dayHex), Len(dayHex)) = dayHex
Mid(lExecId, 18 - Len(millisecsHex), Len(millisecsHex)) = millisecsHex
Mid(lExecId, 9, 1) = "."
Mid(lExecId, 18, 6) = ".01.01"

createExecId = lExecId
End Function

Private Function createSimulatorOrder( _
                ByVal pOrder As IOrder, _
                ByVal pParentId As String, _
                ByVal pOcaGroup As String, _
                ByVal pTransmit As Boolean, _
                ByVal pListener As IOrderSubmissionListener) As SimulatorOrder
Const ProcName As String = "createSimulatorOrder"
On Error GoTo Err

Static sNextBrokerOrderId As Long

sNextBrokerOrderId = sNextBrokerOrderId + 1
pOrder.BrokerId = CStr(sNextBrokerOrderId)

Dim lOrder As New Order
lOrder.DontCheckAttributeModificationPermitted = True
lOrder.Initialise pOrder.GroupName, pOrder.ContractSpecifier, pOrder.OrderContext
lOrder.Id = pOrder.Id
SyncToOrder lOrder, pOrder

Dim lSimulatorOrder As New SimulatorOrder
lSimulatorOrder.Initialise lOrder, pParentId, pOcaGroup, pTransmit, pListener
mAllOrders.Add lSimulatorOrder, pOrder.Id

If pOrder.Status = OrderStatusCreated Then setOrderStatus lSimulatorOrder, OrderStatusPendingSubmit

Set createSimulatorOrder = lSimulatorOrder

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function delay() As Long
delay = 200 + Fix(Rnd() * 300 + 1)
End Function

Private Sub handleNewOrder(ByVal pOrder As SimulatorOrder)
Const ProcName As String = "handleNewOrder"
On Error GoTo Err

If pOrder.ParentId <> "" Then
    ' this is an attached order. We'll move it on at the next tick
Else
    ' this is a parent order, a free standing order, or part of an oca group.
    ' We want to move it to status submitted, either at the next tick or
    ' using exchange submission delay
    If mSimulateExchangeDelays Then
        setOrderStatus pOrder, OrderStatuses.OrderStatusSubmitted
        DeferAction Me, pOrder, delay, ExpiryTimeUnits.ExpiryTimeUnitMilliseconds
    End If
End If

pOrder.Handled = True

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub handleNewOrders()
' submits any orders which are still in the created state
Const ProcName As String = "handleNewOrders"
On Error GoTo Err

If mActiveOrders.Count = 0 Then Exit Sub
Dim lOrder As SimulatorOrder
For Each lOrder In mActiveOrders
    If Not lOrder.Handled Then
        handleNewOrder lOrder
    End If
Next

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub handlePendingSubmitOrders()
Const ProcName As String = "handlePendingSubmitOrders"
On Error GoTo Err

If mActiveOrders.Count = 0 Then Exit Sub
Dim lOrder As SimulatorOrder
For Each lOrder In mActiveOrders
    If lOrder.Order.Status = OrderStatuses.OrderStatusPendingSubmit And lOrder.ParentId = "" Then
        setSubmitOrder lOrder
    End If
Next

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub handleRelatedOrders(ByVal pOrder As SimulatorOrder)
Const ProcName As String = "handleRelatedOrders"
On Error GoTo Err

If mActiveOrders.Count = 0 Then Exit Sub
Dim lOrder As SimulatorOrder
For Each lOrder In mActiveOrders
    If pOrder.IsParentOf(lOrder) Then
        lOrder.ParentId = ""    ' detach the order from the parent
        
        ' now set the status if necessary
        Select Case lOrder.Order.OrderType
        Case OrderTypes.OrderTypeLimit, OrderTypes.OrderTypeMarket
            If mSimulateExchangeDelays Then
                setOrderStatus lOrder, OrderStatuses.OrderStatusSubmitted
                DeferAction Me, lOrder, delay, ExpiryTimeUnits.ExpiryTimeUnitMilliseconds
            Else
                pOrder.AtExchange = True
                setOrderStatus lOrder, OrderStatuses.OrderStatusSubmitted
            End If
        Case Else
            ' the order status must remain unchanged
        End Select
    End If
    If pOrder.IsSiblingOf(lOrder) Then setCancelOrder lOrder
    If pOrder.IsInSameOcaGroupAs(lOrder) Then setCancelOrder lOrder
Next

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub matchOrder(ByVal pOrder As SimulatorOrder)
Const ProcName As String = "matchOrder"
On Error GoTo Err

Select Case pOrder.Order.OrderType
Case OrderTypes.OrderTypeMarket, _
    OrderTypes.OrderTypeStop
    setFillOrder pOrder, IIf(pOrder.Order.Action = OrderActions.OrderActionBuy, mAsk, mBid), pOrder.Order.Quantity
Case OrderTypes.OrderTypeLimit, _
    OrderTypes.OrderTypeStopLimit
    If pOrder.Order.Action = OrderActions.OrderActionBuy And _
        pOrder.Order.LimitPrice >= mAsk _
    Then
        setFillOrder pOrder, mAsk, pOrder.Order.Quantity
    ElseIf pOrder.Order.Action = OrderActions.OrderActionSell And _
        pOrder.Order.LimitPrice <= mBid _
    Then
        setFillOrder pOrder, mBid, pOrder.Order.Quantity
    End If
End Select

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub matchOrderTrade(ByVal pOrder As SimulatorOrder, _
                            ByVal Trade As Double)
Const ProcName As String = "matchOrderTrade"
On Error GoTo Err

Select Case pOrder.Order.OrderType
Case OrderTypes.OrderTypeMarket, _
    OrderTypes.OrderTypeStop
    setFillOrder pOrder, IIf(pOrder.Order.Action = OrderActions.OrderActionBuy, mAsk, mBid), pOrder.Order.Quantity
Case OrderTypes.OrderTypeLimit, _
    OrderTypes.OrderTypeStop
    If pOrder.Order.Action = OrderActions.OrderActionBuy And _
        pOrder.Order.LimitPrice >= Trade _
    Then
        setFillOrder pOrder, pOrder.Order.LimitPrice, pOrder.Order.Quantity
    ElseIf pOrder.Order.Action = OrderActions.OrderActionSell And _
        pOrder.Order.LimitPrice <= Trade _
    Then
        setFillOrder pOrder, pOrder.Order.LimitPrice, pOrder.Order.Quantity
    End If
End Select

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub NotifyOrderReport( _
                ByVal pOrder As SimulatorOrder)
Const ProcName As String = "notifyOrderReport"
On Error GoTo Err

Dim lOrder As IOrder
Set lOrder = pOrder.Order

Dim lOrderReport As New OrderReport
lOrderReport.Action = lOrder.Action
lOrderReport.AllOrNone = lOrder.AllOrNone
lOrderReport.BlockOrder = lOrder.BlockOrder
lOrderReport.BrokerId = lOrder.BrokerId
lOrderReport.DiscretionaryAmount = lOrder.DiscretionaryAmount
lOrderReport.DisplaySize = lOrder.DisplaySize
lOrderReport.GoodAfterTime = lOrder.GoodAfterTime
lOrderReport.GoodAfterTimeTZ = lOrder.GoodAfterTimeTZ
lOrderReport.GoodTillDate = lOrder.GoodTillDate
lOrderReport.GoodTillDateTZ = lOrder.GoodTillDateTZ
lOrderReport.Hidden = lOrder.Hidden
lOrderReport.Id = lOrder.Id
lOrderReport.IgnoreRegularTradingHours = lOrder.IgnoreRegularTradingHours
lOrderReport.IsSimulated = lOrder.IsSimulated
lOrderReport.LimitPrice = lOrder.LimitPrice
lOrderReport.MinimumQuantity = lOrder.MinimumQuantity
lOrderReport.OrderType = lOrder.OrderType
lOrderReport.Origin = lOrder.Origin
lOrderReport.OriginatorRef = lOrder.OriginatorRef
lOrderReport.OverrideConstraints = lOrder.OverrideConstraints
If Not lOrder.ProviderProperties Is Nothing Then lOrderReport.ProviderProperties = lOrder.ProviderProperties
lOrderReport.Quantity = lOrder.Quantity
lOrderReport.SettlingFirm = lOrder.SettlingFirm
lOrderReport.Status = lOrder.Status
lOrderReport.StopTriggerMethod = lOrder.StopTriggerMethod
lOrderReport.SweepToFill = lOrder.SweepToFill
lOrderReport.TimeInForce = lOrder.TimeInForce
lOrderReport.TriggerPrice = lOrder.TriggerPrice

pOrder.Listener.NotifyOrderReport lOrderReport

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyOrderStatus( _
                ByVal pOrder As SimulatorOrder)
Const ProcName As String = "notifyOrderStatus"
On Error GoTo Err

Dim lOrderStatusReport As New OrderStatusReport
lOrderStatusReport.OrderId = pOrder.Order.Id
lOrderStatusReport.Status = pOrder.Order.Status
lOrderStatusReport.BrokerId = pOrder.Order.BrokerId
pOrder.Listener.NotifyOrderStatusReport lOrderStatusReport

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processAsk(ByVal Timestamp As Date, _
                ByVal Ask As Double)
Const ProcName As String = "processAsk"
On Error GoTo Err

' filter out occasional zeros in tick data
If Ask = 0# Then Exit Sub

mAsk = Ask

If mActiveOrders.Count = 0 Then Exit Sub

mEnumerationInProgess = True

handleNewOrders
If Not mSimulateExchangeDelays Then handlePendingSubmitOrders

Dim en As Enumerator
Set en = mActiveOrders.Enumerator
Do While en.MoveNext
    Dim lOrder As SimulatorOrder
    Set lOrder = en.Current
    If lOrder.IsFinished Then
        en.Remove
    ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusSubmitted Then
        If lOrder.AtExchange Then
            matchOrder lOrder
        End If
    ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
        If lOrder.AtExchange Then
            matchOrder lOrder
            If lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
                setCancelOrder lOrder
            End If
        ElseIf Not mSimulateExchangeDelays Then
            setCancelOrder lOrder
        End If
    End If
Loop

mEnumerationInProgess = False
queueOrders

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processBid(ByVal Timestamp As Date, _
                ByVal Bid As Double)
Const ProcName As String = "processBid"
On Error GoTo Err

' filter out occasional zeros in tick data
If Bid = 0# Then Exit Sub

mBid = Bid

If mActiveOrders.Count = 0 Then Exit Sub

mEnumerationInProgess = True

handleNewOrders
If Not mSimulateExchangeDelays Then handlePendingSubmitOrders

Dim en As Enumerator
Set en = mActiveOrders.Enumerator
Do While en.MoveNext
    Dim lOrder As SimulatorOrder
    Set lOrder = en.Current
    If lOrder.IsFinished Then
        en.Remove
    ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusSubmitted Then
        If lOrder.AtExchange Then
            matchOrder lOrder
        End If
    ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
        If lOrder.AtExchange Then
            matchOrder lOrder
            If lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
                setCancelOrder lOrder
            End If
        ElseIf Not mSimulateExchangeDelays Then
            setCancelOrder lOrder
        End If
    End If
Loop

mEnumerationInProgess = False
queueOrders

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTrade(ByVal pTimestamp As Date, _
                ByVal pTrade As Double)
Const ProcName As String = "processTrade"
On Error GoTo Err

' filter out occasional zeros in tick data
If pTrade = 0# Then Exit Sub

If mActiveOrders.Count = 0 Then Exit Sub

mEnumerationInProgess = True

handleNewOrders
If Not mSimulateExchangeDelays Then handlePendingSubmitOrders

Dim en As Enumerator
Set en = mActiveOrders.Enumerator
Do While en.MoveNext
    Dim lOrder As SimulatorOrder
    Set lOrder = en.Current
    If lOrder.IsFinished Then
        en.Remove
    ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusSubmitted Then
        If lOrder.AtExchange Then
            matchOrderTrade lOrder, pTrade
        End If
    ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusPreSubmitted And _
        lOrder.ParentId = "" _
    Then
        Select Case lOrder.Order.OrderType
        Case OrderTypes.OrderTypeStop, OrderTypes.OrderTypeStopLimit
            If (lOrder.Order.Action = OrderActions.OrderActionBuy And _
                    pTrade >= lOrder.Order.TriggerPrice) Or _
                (lOrder.Order.Action = OrderActions.OrderActionSell And _
                pTrade <= lOrder.Order.TriggerPrice) Then
                If mSimulateExchangeDelays Then
                    setOrderStatus lOrder, OrderStatuses.OrderStatusSubmitted
                    DeferAction Me, lOrder, delay, ExpiryTimeUnits.ExpiryTimeUnitMilliseconds
                Else
                    lOrder.AtExchange = True
                    setOrderStatus lOrder, OrderStatuses.OrderStatusSubmitted
                End If
            End If
        End Select
    ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
        If lOrder.AtExchange Then
            matchOrder lOrder
            If lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
                setCancelOrder lOrder
            End If
        ElseIf Not mSimulateExchangeDelays Then
            setCancelOrder lOrder
        End If
    End If
Loop

mEnumerationInProgess = False
queueOrders

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub queueOrder(ByVal pSimulatorOrder As SimulatorOrder)
Const ProcName As String = "queueOrder"
On Error GoTo Err

If mEnumerationInProgess Then
    mOrdersToBeQueued.Add pSimulatorOrder
    Exit Sub
End If

' if this order was previously queued, remove
' it now
If mPendingOrders.Contains(pSimulatorOrder.Order.Id) Then mPendingOrders.Remove pSimulatorOrder.Order.Id
If mActiveOrders.Contains(pSimulatorOrder.Order.Id) Then mActiveOrders.Remove pSimulatorOrder.Order.Id

If Not pSimulatorOrder.Transmit Then
    mPendingOrders.Add pSimulatorOrder, pSimulatorOrder.Order.Id
Else
    ' submit any other pending orders related
    ' to this one via pParentId or oca group
    If mPendingOrders.Count <> 0 Then
        Dim en As Enumerator
        Set en = mPendingOrders.Enumerator
        Do While en.MoveNext
            Dim lPendingOrder As SimulatorOrder
            Set lPendingOrder = en.Current
            If lPendingOrder.IsFinished Then
                en.Remove
            ElseIf pSimulatorOrder.IsParentOf(lPendingOrder) Or _
                    lPendingOrder.IsRelatedTo(pSimulatorOrder) Then
                mActiveOrders.Add lPendingOrder, lPendingOrder.Order.Id
                en.Remove
                setSubmitOrder lPendingOrder
            End If
        Loop
    End If
    
    mActiveOrders.Add pSimulatorOrder, pSimulatorOrder.Order.Id
    setSubmitOrder pSimulatorOrder
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub queueOrders()
Const ProcName As String = "queueOrders"
On Error GoTo Err

If mOrdersToBeQueued.Count = 0 Then Exit Sub

Dim lSimulatorOrder As SimulatorOrder
For Each lSimulatorOrder In mOrdersToBeQueued
    queueOrder lSimulatorOrder
Next

mOrdersToBeQueued.Clear

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub
                    
Private Sub setCancelOrder(ByVal pOrder As SimulatorOrder)
Const ProcName As String = "setCancelOrder"
On Error GoTo Err

If pOrder.Order.Status = OrderStatuses.OrderStatusCancelled Then Exit Sub

Select Case pOrder.Order.Status
Case OrderStatuses.OrderStatusSubmitted, OrderStatuses.OrderStatusPreSubmitted
    setOrderStatus pOrder, OrderStatuses.OrderStatusCancelling
    If mSimulateExchangeDelays Then
        DeferAction Me, pOrder, delay, ExpiryTimeUnits.ExpiryTimeUnitMilliseconds
    End If
Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusPendingSubmit, _
        OrderStatuses.OrderStatusRejected
    setOrderStatus pOrder, OrderStatuses.OrderStatusCancelled
    pOrder.IsFinished = True
End Select

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setFillOrder(ByVal pOrder As SimulatorOrder, _
                        ByVal pFillPrice As Double, _
                        ByVal pFillSize As BoxedDecimal)
' This needs enhancing to cater for the case
' where the quantity ordered is greater than the
' current Bid/offer Size
Const ProcName As String = "setFillOrder"
On Error GoTo Err

If pFillPrice = 0 Then Exit Sub ' can occur at start of Session before first Bid or offer

Dim lFillPrice As Double
lFillPrice = pFillPrice

Dim lFillSize As BoxedDecimal
Set lFillSize = pFillSize

Dim lContract As IContract
Set lContract = mDataSource.ContractFuture.Value

Dim lContractSpecifier As IContractSpecifier
Set lContractSpecifier = lContract.Specifier

Dim lExec As Execution
Set lExec = New Execution
lExec.IsSimulated = True
lExec.AccountIdentifier = ""
lExec.FillingExchange = lContractSpecifier.Exchange

lExec.Id = createExecId

lExec.IsLiquidation = False
lExec.BrokerId = pOrder.Order.BrokerId
lExec.OrderId = pOrder.Order.Id
lExec.Action = pOrder.Order.Action

Dim lClock As Clock
Set lClock = mDataSource.ClockFuture.Value
lExec.FillTime = lClock.Timestamp

If lFillSize > pOrder.Order.QuantityRemaining Then Set lFillSize = pOrder.Order.QuantityRemaining
If lFillSize < 1 Then Set lFillSize = pFillSize
If lFillPrice = 0 Then lFillPrice = pFillPrice

lExec.Price = lFillPrice
lExec.Quantity = lFillSize
lExec.SecurityName = lContractSpecifier.LocalSymbol
lExec.TimezoneName = lClock.TimeZone.StandardName

pOrder.Order.QuantityRemaining = pOrder.Order.QuantityRemaining.Subtr(lFillSize)
If pOrder.Order.QuantityRemaining = 0 Then pOrder.Order.Status = OrderStatuses.OrderStatusFilled

pOrder.Listener.NotifyExecutionReport lExec

NotifyOrderReport pOrder
notifyOrderStatus pOrder

handleRelatedOrders pOrder

If pOrder.Order.QuantityRemaining = 0 Then pOrder.IsFinished = True

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrderStatus( _
                ByVal pOrder As SimulatorOrder, _
                ByVal pStatus As OrderStatuses)
Const ProcName As String = "setOrderStatus"
On Error GoTo Err

pOrder.Order.Status = pStatus
NotifyOrderReport pOrder
notifyOrderStatus pOrder

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setRejectOrder( _
                ByVal pOrder As SimulatorOrder, _
                ByVal pErrorCode As Long, _
                ByVal pReason As String)
Const ProcName As String = "setRejectOrder"
On Error GoTo Err

pOrder.Order.Status = OrderStatuses.OrderStatusRejected

pOrder.Listener.NotifyError pOrder.Order.Id, pErrorCode, "Order rejected - reason: " & pReason
notifyOrderStatus pOrder

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSubmitOrder(pOrder As SimulatorOrder)
Const ProcName As String = "setSubmitOrder"
On Error GoTo Err

Select Case pOrder.Order.OrderType
Case OrderTypes.OrderTypeLimit, OrderTypes.OrderTypeMarket
    Select Case pOrder.Order.Status
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
        AssertArgument True, "Order type " & pOrder.Order.OrderType & " invalid for submission"
    End Select
Case OrderTypes.OrderTypeStop, OrderTypes.OrderTypeStopLimit
    Select Case pOrder.Order.Status
    Case OrderStatuses.OrderStatusPendingSubmit
        setOrderStatus pOrder, OrderStatuses.OrderStatusPreSubmitted
    Case Else
        AssertArgument True, "Order type " & pOrder.Order.OrderType & " invalid for submission"
    End Select
        
Case Else
    AssertArgument True, "Order type " & pOrder.Order.OrderType & " invalid for submission"
End Select

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub







