VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InProcessExchangeSimulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements IGenericTickListener
Implements StateChangeListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "InProcessExchangeSimulator"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type MarketTableEntry
    ContractSpec            As IContractSpecifier
    Bid                     As Double
    BidSize                 As Long
    Ask                     As Double
    AskSize                 As Long
    ActiveOrders            As EnumerableCollection
    PendingOrders           As EnumerableCollection
    SimulateExchangeDelays  As Boolean
    DataSource              As IMarketDataSource
End Type

Private Type ClientTableEntry
    InUse                   As Boolean
    Listener                As IOrderSubmissionListener
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mMarketTable()                  As MarketTableEntry
Private mClientTable()                  As ClientTableEntry

Private mAllOrders                      As EnumerableCollection

Private mTimerList                      As TimerList
Attribute mTimerList.VB_VarHelpID = -1
    ' This is used to simulate delays in submitting orders to the
    ' exchange for execution. When it is time to change an order's status
    ' from pendingsubmit to submitted, the order is Added to this list
    ' with a random wait of between 200 and 500 millisecs.

Private mMarketDataManager              As IMarketDataManager
Private mPrimaryContractStore           As IContractStore
Private mSecondaryContractStore         As IContractStore

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mMarketTable(3) As MarketTableEntry
ReDim mClientTable(3) As ClientTableEntry

Set mAllOrders = New EnumerableCollection

Set mTimerList = GetGlobalTimerList

Randomize
End Sub

Private Sub Class_Terminate()
Debug.Print "AdvancedOrderSimulator terminated"
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Dim lDataSource As IMarketDataSource

Set lDataSource = ev.Source

Select Case ev.Tick.TickType
Case TickTypes.TickTypeAsk
    processAsk ev.Tick.Timestamp, lDataSource.Handle, ev.Tick.Price
Case TickTypes.TickTypeBid
    processBid ev.Tick.Timestamp, lDataSource.Handle, ev.Tick.Price
Case TickTypes.TickTypeClosePrice
Case TickTypes.TickTypeHighPrice
Case TickTypes.TickTypeLowPrice
Case TickTypes.TickTypeMarketDepth
Case TickTypes.TickTypeMarketDepthReset
Case TickTypes.TickTypeOpenInterest
Case TickTypes.TickTypeOpenPrice
Case TickTypes.TickTypeTrade
    processTrade ev.Tick.Timestamp, lDataSource.Handle, ev.Tick.Price
Case TickTypes.TickTypeVolume
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change( _
                ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err


Dim tli As TimerListItem
Set tli = ev.Source

Dim lOrder As SimulatorOrder
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
    Assert True, "Unexpected TimerListItem status"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get PermittedOrderProperties() As IPermittedOrderProperties
Const ProcName As String = "PermittedOrderProperties"
On Error GoTo Err

Dim lPermittedOrderProperties As New PermittedOrderProperties

Dim lPermittedOrderattributes As OrderAttributes
lPermittedOrderattributes = OrderAttributes.OrderAttAction + _
                            OrderAttributes.OrderAttLimitPrice + _
                            OrderAttributes.OrderAttOrderType + _
                            OrderAttributes.OrderAttOriginatorRef + _
                            OrderAttributes.OrderAttQuantity + _
                            OrderAttributes.OrderAttStopTriggerMethod + _
                            OrderAttributes.OrderAttTriggerPrice
lPermittedOrderProperties.PermittedOrderAttributes = lPermittedOrderattributes

Dim lPermittedOrderTifs As OrderTIFs
lPermittedOrderTifs = OrderTIFDay
lPermittedOrderProperties.PermittedOrderTifs = lPermittedOrderTifs

Dim lPermittedOrderTypes As OrderTypes
lPermittedOrderTypes = OrderTypes.OrderTypeMarket + _
                        OrderTypes.OrderTypeLimit + _
                        OrderTypes.OrderTypeStop + _
                        OrderTypes.OrderTypeStopLimit
lPermittedOrderProperties.PermittedOrderTypes = lPermittedOrderTypes

Dim lPermittedOrderStopTriggerMethods As OrderStopTriggerMethods
lPermittedOrderStopTriggerMethods = OrderStopTriggerDefault + OrderStopTriggerLast
lPermittedOrderProperties.PermittedStopTriggerMethods = lPermittedOrderStopTriggerMethods

Set PermittedOrderProperties = lPermittedOrderProperties

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property


'@================================================================================
' Methods
'@================================================================================

Public Sub CancelOrder(ByVal Id As String)
Const ProcName As String = "CancelOrder"
On Error GoTo Err

Dim relatedOrder As SimulatorOrder

AssertArgument mAllOrders.Contains(Id), "Unknown order id " & Id

Dim lOrder As SimulatorOrder
Set lOrder = mAllOrders(Id)

With mMarketTable(lOrder.MarketId)
    If .PendingOrders.Contains(Id) Then
        Dim lPendingOrder As SimulatorOrder
        Set lPendingOrder = .PendingOrders(Id)
        If lOrder.ParentId <> "" Then
            Set relatedOrder = .PendingOrders(lOrder.ParentId)
            setCancelOrder relatedOrder
            
            For Each relatedOrder In .PendingOrders
                If relatedOrder.ParentId = lOrder.ParentId And _
                    relatedOrder.Order.Id <> lOrder.Order.Id _
                Then
                    setCancelOrder relatedOrder
                End If
            Next
        Else
            For Each relatedOrder In .PendingOrders
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
        Set lOrder = .ActiveOrders(Id)
        On Error GoTo Err
        
        If lOrder Is Nothing Then Exit Sub
        
        If lOrder.ParentId <> "" Then
            On Error Resume Next
            Set relatedOrder = .ActiveOrders(lOrder.ParentId)
            On Error GoTo Err
            If Not relatedOrder Is Nothing Then
                setCancelOrder relatedOrder
            End If
            For Each relatedOrder In .ActiveOrders
                If relatedOrder.ParentId = lOrder.ParentId And _
                    relatedOrder.Order.Id <> lOrder.Order.Id _
                Then
                    setCancelOrder relatedOrder
                End If
            Next
        Else
            For Each relatedOrder In .ActiveOrders
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise(ByVal pMarketDataManager As IMarketDataManager, ByVal pPrimaryContractStore As IContractStore, ByVal pSecondaryContractStore As IContractStore)
Set mMarketDataManager = pMarketDataManager
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore
End Sub

Public Sub PlaceOrder( _
                ByVal pOrder As IOrder, _
                ByVal pParentId As String, _
                ByVal pOcaGroup As String, _
                ByVal pClientId As Long, _
                ByVal pTransmit As Boolean)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

AssertArgument pClientId >= 0 And pClientId <= UBound(mClientTable), "Invalid ClientId " & pClientId & " in order " & pOrder.Id
AssertArgument mClientTable(pClientId).InUse, "Invalid ClientId " & pClientId & " in order " & pOrder.Id

Select Case pOrder.OrderType
Case OrderTypes.OrderTypeMarket, _
    OrderTypes.OrderTypeLimit, _
    OrderTypes.OrderTypeStop, _
    OrderTypes.OrderTypeStopLimit
Case Else
    setRejectOrder pOrder, _
                    ErrorCodes.ErrUnsupportedOperationException, _
                    "Order type " & pOrder.OrderType & " not supported by simulator", _
                    pClientId
    Exit Sub
End Select

Dim lSimulatorOrder As SimulatorOrder

If mAllOrders.Contains(pOrder.Id) And pOrder.BrokerId = "" Then
    ' the order is being re-placed (typically re-placing a stoploss or target order
    ' after it was cancelled while trying to cancel a bracket order, but the entry
    ' order was subsequently filled)
    mAllOrders.Remove pOrder.Id
    Set lSimulatorOrder = createSimulatorOrder(pOrder, pClientId, pParentId, pOcaGroup, pTransmit)
ElseIf mAllOrders.Contains(pOrder.Id) Then
    Set lSimulatorOrder = mAllOrders(pOrder.Id)
    ' the order is being updated
    gSyncToOrder lSimulatorOrder.Order, pOrder
    lSimulatorOrder.Transmit = pTransmit
Else
    Set lSimulatorOrder = createSimulatorOrder(pOrder, pClientId, pParentId, pOcaGroup, pTransmit)
End If

queueOrder lSimulatorOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                    
Public Function RegisterClient( _
                ByVal pListener As IOrderSubmissionListener) As Long
Const ProcName As String = "RegisterClient"
On Error GoTo Err

Dim lIndex As Long
lIndex = -1

Dim i As Long
For i = 0 To UBound(mClientTable)
    If (Not mClientTable(i).InUse) Then
        lIndex = i
        Exit For
    End If
Next

If lIndex = -1 Then
    lIndex = UBound(mClientTable) + 1
    ReDim Preserve mClientTable(2 * (UBound(mClientTable) + 1) - 1) As ClientTableEntry
End If

With mClientTable(lIndex)
    .InUse = True
    Set .Listener = pListener
End With

RegisterClient = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

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
                ByVal pClientId As Long, _
                ByVal pParentId As String, _
                ByVal pOcaGroup As String, _
                ByVal pTransmit As Boolean) As SimulatorOrder
Const ProcName As String = "createSimulatorOrder"
On Error GoTo Err

Static sNextBrokerOrderId As Long

sNextBrokerOrderId = sNextBrokerOrderId + 1
pOrder.BrokerId = CStr(sNextBrokerOrderId)

Dim lOrder As New SimulatorOrder
lOrder.Initialise Me, New Order, pClientId, pParentId, pOcaGroup, pTransmit, getMarketId(pOrder.ContractSpecifier)
gSyncToOrder lOrder.Order, pOrder
mAllOrders.Add lOrder, pOrder.Id

Set createSimulatorOrder = lOrder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function delay() As Long
delay = 200 + Fix(Rnd() * 300 + 1)
End Function

Private Function findMarketId( _
                ByVal pContractSpec As IContractSpecifier) As Long
Const ProcName As String = "findMarketId"
On Error GoTo Err

Dim lIndex As Long
lIndex = -1

Dim i As Long
For i = 0 To UBound(mMarketTable)
    If Not mMarketTable(i).ContractSpec Is Nothing Then
        If pContractSpec.Equals(mMarketTable(i).ContractSpec) Then
            lIndex = i
            Exit For
        End If
    End If
Next

findMarketId = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getMarketDataSource(ByVal pContractSpec As IContractSpecifier) As IMarketDataSource
Const ProcName As String = "getMarketDataSource"
On Error GoTo Err

Dim lDataSource As IMarketDataSource

Set lDataSource = mMarketDataManager.CreateMarketDataSource(FetchContract(pContractSpec, mPrimaryContractStore, mSecondaryContractStore), False)
Assert Not lDataSource Is Nothing, "Can't create market data source"

Set getMarketDataSource = lDataSource

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getMarketId( _
                ByVal pContractSpec As IContractSpecifier) As Long
Const ProcName As String = "getMarketId"
On Error GoTo Err

Dim lIndex As Long
lIndex = findMarketId(pContractSpec)

If lIndex = -1 Then lIndex = startMarket(pContractSpec)

getMarketId = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleCreatedOrders(ByVal MarketId As Long)
' submits any orders which are still in the created state
Const ProcName As String = "handleCreatedOrders"
On Error GoTo Err

With mMarketTable(MarketId)
    Dim lOrder As SimulatorOrder
    For Each lOrder In .ActiveOrders
        If lOrder.Order.Status = OrderStatuses.OrderStatusCreated Then
            setSubmitOrder lOrder
        End If
    Next
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub handlePendingSubmitOrders(ByVal MarketId As Long)
Const ProcName As String = "handlePendingSubmitOrders"
On Error GoTo Err

With mMarketTable(MarketId)
    Dim lOrder As SimulatorOrder
    For Each lOrder In .ActiveOrders
        If lOrder.Order.Status = OrderStatuses.OrderStatusPendingSubmit And lOrder.ParentId = "" Then
            setSubmitOrder lOrder
        End If
    Next
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub handleRelatedOrders(ByVal pOrder As SimulatorOrder)
Const ProcName As String = "handleRelatedOrders"
On Error GoTo Err

With mMarketTable(pOrder.MarketId)
    Dim lOrder As SimulatorOrder
    For Each lOrder In .ActiveOrders
        If lOrder.ParentId = pOrder.Order.Id Then
            lOrder.ParentId = ""    ' detach the order from the parent
            
            ' now set the status if necessary
            Select Case lOrder.Order.OrderType
            Case OrderTypes.OrderTypeLimit, OrderTypes.OrderTypeMarket
                If mMarketTable(pOrder.ClientId).SimulateExchangeDelays Then
                    setOrderStatus pOrder, OrderStatuses.OrderStatusSubmitted
                    mTimerList.Add(pOrder, _
                                    delay, _
                                    ExpiryTimeUnits.ExpiryTimeUnitMilliseconds).AddStateChangeListener Me
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub matchOrder(ByVal pOrder As SimulatorOrder)
Const ProcName As String = "matchOrder"
On Error GoTo Err

With mMarketTable(pOrder.ClientId)
    Select Case pOrder.Order.OrderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeStop
        setFillOrder pOrder, IIf(pOrder.Order.Action = OrderActions.OrderActionBuy, .Ask, .Bid), pOrder.Order.Quantity
    Case OrderTypes.OrderTypeLimit, _
        OrderTypes.OrderTypeStopLimit
        If pOrder.Order.Action = OrderActions.OrderActionBuy And _
            pOrder.Order.LimitPrice >= .Ask _
        Then
            setFillOrder pOrder, .Ask, pOrder.Order.Quantity
        ElseIf pOrder.Order.Action = OrderActions.OrderActionSell And _
            pOrder.Order.LimitPrice <= .Bid _
        Then
            setFillOrder pOrder, .Bid, pOrder.Order.Quantity
        End If
    End Select
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub matchOrderTrade(ByVal pOrder As SimulatorOrder, _
                            ByVal Trade As Double)
Const ProcName As String = "matchOrderTrade"
On Error GoTo Err

With mMarketTable(pOrder.ClientId)
    Select Case pOrder.Order.OrderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeStop
        setFillOrder pOrder, IIf(pOrder.Order.Action = OrderActions.OrderActionBuy, .Ask, .Bid), pOrder.Order.Quantity
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
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyOrderStatus( _
                ByVal pOrder As SimulatorOrder)
Const ProcName As String = "notifyOrderStatus"
On Error GoTo Err

Dim lOrderStatusReport As New OrderStatusReport
lOrderStatusReport.OrderId = pOrder.Order.Id
lOrderStatusReport.Status = pOrder.Order.Status
lOrderStatusReport.BrokerId = pOrder.Order.BrokerId
mClientTable(pOrder.ClientId).Listener.NotifyOrderStatusReport lOrderStatusReport

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub pauseMarket(ByVal pHandle As Long)
Const ProcName As String = "pauseMarket"
On Error GoTo Err

With mMarketTable(pHandle)
    .Ask = 0
    .AskSize = 0
    .Bid = 0
    .BidSize = 0
    .DataSource.StopMarketData
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processAsk(ByVal Timestamp As Date, _
                ByVal mMarketId As Long, _
                ByVal Ask As Double)
Const ProcName As String = "processAsk"
On Error GoTo Err

With mMarketTable(mMarketId)
    .Ask = Ask
    
    handleCreatedOrders mMarketId
    If Not .SimulateExchangeDelays Then handlePendingSubmitOrders mMarketId

    Dim lOrder As SimulatorOrder
    For Each lOrder In .ActiveOrders
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
            ElseIf Not .SimulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processBid(ByVal Timestamp As Date, _
                ByVal mMarketId As Long, _
                ByVal Bid As Double)
Const ProcName As String = "processBid"
On Error GoTo Err

With mMarketTable(mMarketId)
    .Bid = Bid
    
    handleCreatedOrders mMarketId
    If Not .SimulateExchangeDelays Then handlePendingSubmitOrders mMarketId

    Dim lOrder As SimulatorOrder
    For Each lOrder In .ActiveOrders
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
            ElseIf Not .SimulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub queueOrder(ByVal pSimulatedOrder As SimulatorOrder)
Const ProcName As String = "queueOrder"
On Error GoTo Err

Dim lPendingOrder As SimulatorOrder

With mMarketTable(pSimulatedOrder.MarketId)
    If Not pSimulatedOrder.Transmit Then
        .PendingOrders.Add pSimulatedOrder, pSimulatedOrder.Order.Id
    Else
        ' if this order was previously in the pending list, Remove
        ' it now
        If .PendingOrders.Contains(pSimulatedOrder.Order.Id) Then .PendingOrders.Remove pSimulatedOrder.Order.Id
        
        ' now see if there are any other pending orders related
        ' to this one via pParentId or oca group
        If pSimulatedOrder.ParentId <> "" Then
            Set lPendingOrder = .PendingOrders(pSimulatedOrder.ParentId)
            .ActiveOrders.Add lPendingOrder, lPendingOrder.Order.Id
            .PendingOrders.Remove lPendingOrder.Order.Id
            setSubmitOrder lPendingOrder
            
            ' find any other child orders
            For Each lPendingOrder In .PendingOrders
                If lPendingOrder.ParentId = pSimulatedOrder.ParentId Then
                    .ActiveOrders.Add lPendingOrder, lPendingOrder.Order.Id
                    .PendingOrders.Remove lPendingOrder.Order.Id
                    setSubmitOrder lPendingOrder
                End If
            Next
        Else
            ' find any child orders or related OCA orders
            For Each lPendingOrder In .PendingOrders
                If lPendingOrder.ParentId = pSimulatedOrder.Order.Id Or _
                    (lPendingOrder.OcaGroup = pSimulatedOrder.OcaGroup And _
                    pSimulatedOrder.OcaGroup <> "") _
                Then
                    .ActiveOrders.Add lPendingOrder, lPendingOrder.Order.Id
                    .PendingOrders.Remove lPendingOrder.Order.Id
                    setSubmitOrder lPendingOrder
                End If
            Next
                        
        End If
        
        .ActiveOrders.Add pSimulatedOrder, pSimulatedOrder.Order.Id
        setSubmitOrder pSimulatedOrder
    End If
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                    
Private Sub restartMarket(ByVal pHandle As Long)
Const ProcName As String = "restartMarket"
On Error GoTo Err

With mMarketTable(pHandle)
    .DataSource.StartMarketData
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCancelOrder(ByVal pOrder As SimulatorOrder)
Const ProcName As String = "setCancelOrder"
On Error GoTo Err

If pOrder.Order.Status = OrderStatuses.OrderStatusCancelled Then Exit Sub

Select Case pOrder.Order.Status
Case OrderStatuses.OrderStatusSubmitted, OrderStatuses.OrderStatusPreSubmitted
    setOrderStatus pOrder, OrderStatuses.OrderStatusCancelling
    If mMarketTable(pOrder.ClientId).SimulateExchangeDelays Then
        mTimerList.Add(pOrder, _
                            delay, _
                            ExpiryTimeUnits.ExpiryTimeUnitMilliseconds).AddStateChangeListener Me
    End If
Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusPendingSubmit, _
        OrderStatuses.OrderStatusRejected
    setOrderStatus pOrder, OrderStatuses.OrderStatusCancelled
    mMarketTable(pOrder.ClientId).ActiveOrders.Remove pOrder.Order.Id
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setFillOrder(ByVal pOrder As SimulatorOrder, _
                        ByVal pFillPrice As Double, _
                        ByVal pFillSize As Long)
' This needs enhancing to cater for the case
' where the quantity ordered is greater than the
' current Bid/offer Size
Const ProcName As String = "setFillOrder"
On Error GoTo Err

If pFillPrice = 0 Then Exit Sub ' can occur at start of Session before first Bid or offer

Dim lFillPrice As Double
lFillPrice = pFillPrice

Dim lFillSize As Long
lFillSize = pFillSize

Dim lContractSpecifier As ContractSpecifier
Set lContractSpecifier = mMarketTable(pOrder.ClientId).ContractSpec

Dim lExec As Execution
Set lExec = New Execution
'lExec.contractSpecifier = lContractSpecifier
lExec.AccountIdentifier = ""
lExec.FillingExchange = lContractSpecifier.Exchange

lExec.Id = createExecId

lExec.IsLiquidation = False
lExec.BrokerId = pOrder.Order.BrokerId
lExec.OrderId = pOrder.Order.Id
lExec.Action = pOrder.Order.Action

Dim lClock As Clock
Set lClock = mMarketTable(pOrder.ClientId).DataSource.ClockFuture.Value
lExec.FillTime = lClock.Timestamp

'mMarketTable(pOrder.ClientId).OrderSubmitter.PreFill pOrder.Order.Id, _
'                    lFillPrice, _
'                    lFillSize
                    
If lFillSize > pOrder.Order.QuantityRemaining Then lFillSize = pOrder.Order.QuantityRemaining
If lFillSize < 1 Then lFillSize = pFillSize
If lFillPrice = 0 Then lFillPrice = pFillPrice

lExec.Price = lFillPrice
lExec.Quantity = lFillSize
lExec.TimezoneName = lClock.TimeZone.DisplayName

pOrder.Order.QuantityRemaining = pOrder.Order.QuantityRemaining - lFillSize
If pOrder.Order.QuantityRemaining = 0 Then pOrder.Order.Status = OrderStatuses.OrderStatusFilled

lExec.IsSimulated = True
mClientTable(pOrder.ClientId).Listener.NotifyExecutionReport lExec


notifyOrderStatus pOrder

handleRelatedOrders pOrder

If pOrder.Order.QuantityRemaining = 0 Then
    mMarketTable(pOrder.MarketId).ActiveOrders.Remove pOrder.Order.Id
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Sub setRejectOrder( _
                ByVal pOrder As Order, _
                ByVal pErrorCode As Long, _
                ByVal pReason As String, _
                ByVal pClientId As Long)
Const ProcName As String = "setRejectOrder"

On Error GoTo Err

pOrder.Status = OrderStatuses.OrderStatusRejected

mClientTable(pClientId).Listener.NotifyError pOrder.Id, pErrorCode, "Order rejected - reason: " & pReason
notifyOrderStatus pOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
            If mMarketTable(pOrder.ClientId).SimulateExchangeDelays Then
                setOrderStatus pOrder, OrderStatuses.OrderStatusSubmitted
                mTimerList.Add(pOrder, _
                                delay, _
                                ExpiryTimeUnits.ExpiryTimeUnitMilliseconds).AddStateChangeListener Me
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
        AssertArgument True, "Order type " & pOrder.Order.OrderType & " invalid for submission"
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
        If mMarketTable(pOrder.ClientId).SimulateExchangeDelays Then
            setOrderStatus pOrder, OrderStatuses.OrderStatusSubmitted
            mTimerList.Add(pOrder, _
                            delay, _
                            ExpiryTimeUnits.ExpiryTimeUnitMilliseconds).AddStateChangeListener Me
        Else
            pOrder.AtExchange = True
            setOrderStatus pOrder, OrderStatuses.OrderStatusSubmitted
        End If
    Case Else
        AssertArgument True, "Order type " & pOrder.Order.OrderType & " invalid for submission"
    End Select
        
Case Else
    AssertArgument True, "Order type " & pOrder.Order.OrderType & " invalid for submission"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function startMarket( _
                ByVal pContractSpec As IContractSpecifier) As Long
Const ProcName As String = "startMarket"
On Error GoTo Err

Dim lDataSource As IMarketDataSource
Set lDataSource = getMarketDataSource(pContractSpec)

Dim lIndex As Long
lIndex = lDataSource.Handle

If lIndex > UBound(mMarketTable) Then
    ReDim Preserve mMarketTable(2 * (lIndex + 1) - 1) As MarketTableEntry
End If

With mMarketTable(lIndex)
    Set .ActiveOrders = New EnumerableCollection
    Set .ContractSpec = pContractSpec
    Set .PendingOrders = New EnumerableCollection
    Set .DataSource = lDataSource
    .DataSource.AddGenericTickListener Me
End With

lDataSource.StartMarketData

startMarket = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub stopMarket(ByVal pHandle As Long)
Const ProcName As String = "StopMarket"
On Error GoTo Err

With mMarketTable(pHandle)
    Set .ContractSpec = Nothing
    Set .ActiveOrders = Nothing
    Set .PendingOrders = Nothing
    .SimulateExchangeDelays = False
    .Ask = 0
    .AskSize = 0
    .Bid = 0
    .BidSize = 0
    .DataSource.StopMarketData
    Set .DataSource = Nothing
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTrade(ByVal pTimestamp As Date, _
                ByVal pMarketId As Long, _
                ByVal pTrade As Double)
Const ProcName As String = "processTrade"
On Error GoTo Err

Dim lOrder As SimulatorOrder

With mMarketTable(pMarketId)
    handleCreatedOrders pMarketId
    If Not .SimulateExchangeDelays Then handlePendingSubmitOrders pMarketId

    For Each lOrder In .ActiveOrders
        If lOrder.Order.Status = OrderStatuses.OrderStatusSubmitted Then
            If lOrder.AtExchange Then
                matchOrderTrade lOrder, pTrade
            End If
        ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusPreSubmitted And _
            lOrder.ParentId = "" _
        Then
            Select Case lOrder.Order.OrderType
                Case OrderTypes.OrderTypeStop
                    If lOrder.Order.Action = OrderActions.OrderActionBuy And _
                        pTrade >= lOrder.Order.TriggerPrice Then
                        setSubmitOrder lOrder
                    ElseIf lOrder.Order.Action = OrderActions.OrderActionSell And _
                        pTrade <= lOrder.Order.TriggerPrice Then
                        setSubmitOrder lOrder
                    End If
                Case OrderTypes.OrderTypeStopLimit
                    If lOrder.Order.Action = OrderActions.OrderActionBuy And _
                        pTrade >= lOrder.Order.TriggerPrice Then
                        setSubmitOrder lOrder
                    ElseIf lOrder.Order.Action = OrderActions.OrderActionSell And _
                        pTrade <= lOrder.Order.TriggerPrice Then
                        setSubmitOrder lOrder
                    End If
            End Select
        ElseIf lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
            If lOrder.AtExchange Then
                matchOrder lOrder
                If lOrder.Order.Status = OrderStatuses.OrderStatusCancelling Then
                    setCancelOrder lOrder
                End If
            ElseIf Not .SimulateExchangeDelays Then
                setCancelOrder lOrder
            End If
        End If
    Next
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



