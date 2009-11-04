VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimulatedOrderSubmitter"
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
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements TradeBuildSP.IOrderSubmitter
Implements TradeBuildSP.IStreamingDataConsumer

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mServiceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer

Private mSimulator As AdvancedOrderSimulator

Private mMarketIndex As Long

Private mContract As Contract

Private mFirstAskReceived       As Boolean
Private mFirstBidReceived       As Boolean
Private mReadyNotified          As Boolean


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "SimulatedOrderSubmitter terminated"
End Sub

'@================================================================================
' IOrderSubmitter Interface Members
'@================================================================================

Private Function IOrderSubmitter_cancelOrderPlex(ByVal pOrderPlex As TradeBuildSP.IOrderPlex) As Boolean
Dim EntryOrder As Order
Dim stoporder As Order
Dim TargetOrder As Order

Set EntryOrder = pOrderPlex.EntryOrder
Select Case EntryOrder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    mSimulator.CancelOrder EntryOrder.BrokerId  ' should automatically cancel the other orders
                                                    ' if they have parentid set
    IOrderSubmitter_cancelOrderPlex = True
End Select

Set stoporder = pOrderPlex.stoporder
If Not stoporder Is Nothing Then
    Select Case stoporder.Status
    Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case EntryOrder.Status
        Case OrderStatuses.OrderStatusFilled, _
            OrderStatuses.OrderStatusCancelling, _
            OrderStatuses.OrderStatusCancelled
            mSimulator.CancelOrder stoporder.BrokerId
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If

Set TargetOrder = pOrderPlex.TargetOrder
If Not TargetOrder Is Nothing Then
    Select Case TargetOrder.Status
    Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case EntryOrder.Status
        Case OrderStatuses.OrderStatusFilled, _
            OrderStatuses.OrderStatusCancelling, _
            OrderStatuses.OrderStatusCancelled
            mSimulator.CancelOrder TargetOrder.BrokerId
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If
    
' need some stuff here to cancel if it's an oca group

End Function

Private Sub IOrderSubmitter_cancelStopOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim stoporder As Order

Set stoporder = pOrderPlex.stoporder
If stoporder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelStopOrder", _
                                    "Order plex has no stop order"

Select Case stoporder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelStopOrder", _
                                    "Stop order state invalid for cancellation"
Case Else
    mSimulator.CancelOrder stoporder.BrokerId
End Select
End Sub

Private Sub IOrderSubmitter_cancelTargetOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim TargetOrder As Order

Set TargetOrder = pOrderPlex.TargetOrder
If TargetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelTargetOrder", _
                                    "Order plex has no target order"


Select Case TargetOrder.Status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelTargetOrder", _
                                    "Target order state invalid for cancellation"
Case Else
    mSimulator.CancelOrder TargetOrder.BrokerId
End Select
End Sub

Private Property Let IOrderSubmitter_ClockRate(ByVal RHS As Single)
mSimulator.ClockRate(mMarketIndex) = RHS
End Property

Private Sub IOrderSubmitter_executeOrderPlex(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim EntryOrder As Order
Dim stoporder As Order
Dim TargetOrder As Order
Dim transmit As Boolean
Dim ocaGroup As String

Set EntryOrder = pOrderPlex.EntryOrder
Set stoporder = pOrderPlex.stoporder
Set TargetOrder = pOrderPlex.TargetOrder

If Not EntryOrder Is Nothing Then
    transmit = True
    If Not stoporder Is Nothing Then
        If stoporder.OrderType <> OrderTypes.OrderTypeAutoStop Then transmit = False
    End If
    If Not TargetOrder Is Nothing Then
        If TargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then transmit = False
    End If
        
    placeOrder EntryOrder, "", "", transmit
    
    ocaGroup = GenerateTextID

    If Not stoporder Is Nothing Then
        transmit = True
        If Not TargetOrder Is Nothing Then
            If TargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then transmit = False
        End If
        
        If stoporder.OrderType <> OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry order is filled
            placeOrder stoporder, EntryOrder.Id, ocaGroup, transmit
        End If
    End If
    
    If Not TargetOrder Is Nothing Then
        If TargetOrder.OrderType <> OrderTypes.OrderTypeAutoLimit Then
            placeOrder TargetOrder, EntryOrder.Id, ocaGroup, True
        End If
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If
End Sub

Private Sub IOrderSubmitter_Finish()
mSimulator.stopMarket mMarketIndex
Set mServiceConsumer = Nothing
Set mSimulator = Nothing
Set mContract = Nothing
End Sub

Private Sub IOrderSubmitter_modifyOrderPlex( _
                ByVal pOrderPlex As TradeBuildSP.IOrderPlex, _
                ByVal entryOrderChanged As Boolean, _
                ByVal stopOrderChanged As Boolean, _
                ByVal targetOrderChanged As Boolean, _
                ByVal closeoutOrderChanged As Boolean)
Dim lOrder As Order

If entryOrderChanged Then
    placeOrder pOrderPlex.EntryOrder, "", "", True
End If
If stopOrderChanged Then
    Set lOrder = pOrderPlex.stoporder
    placeOrder lOrder, "", "", True
End If
If targetOrderChanged Then
    Set lOrder = pOrderPlex.stoporder
    placeOrder lOrder, "", "", True
End If
If closeoutOrderChanged Then
    placeOrder pOrderPlex.CloseoutOrder, "", "", True
End If
    
End Sub

Private Property Get IOrderSubmitter_permittedOrderAttributes() As Long
Dim attributes As Long
attributes = OrderAttributeIds.OrderAttAction
attributes = attributes Or OrderAttributeIds.OrderAttLimitPrice
attributes = attributes Or OrderAttributeIds.OrderAttOrderType
attributes = attributes Or OrderAttributeIds.OrderAttOriginatorRef
attributes = attributes Or OrderAttributeIds.OrderAttQuantity
attributes = attributes Or OrderAttributeIds.OrderAttTimeInForce
attributes = attributes Or OrderAttributeIds.OrderAttTriggerPrice
IOrderSubmitter_permittedOrderAttributes = attributes
End Property

Private Property Get IOrderSubmitter_permittedOrderTifs() As Long
Dim permittedTifs As Long
permittedTifs = TIFDay
IOrderSubmitter_permittedOrderTifs = permittedTifs
End Property

Private Property Get IOrderSubmitter_permittedOrderTypes() As Long
Dim permittedOrderTypes As Long
permittedOrderTypes = OrderTypeLimit Or _
                        OrderTypeMarket Or _
                        OrderTypeStop Or _
                        OrderTypeStopLimit
IOrderSubmitter_permittedOrderTypes = permittedOrderTypes
End Property

Private Property Get IOrderSubmitter_permittedStopTriggerMethods() As Long
Dim permittedStopTriggerMethods As Long
permittedStopTriggerMethods = StopTriggerLast
IOrderSubmitter_permittedStopTriggerMethods = permittedStopTriggerMethods
End Property

Private Sub IOrderSubmitter_resubmitStopAndTargetOrders(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim ocaGroup As String
Dim stoporder As Order
Dim TargetOrder As Order

ocaGroup = GenerateTextID

Set stoporder = pOrderPlex.stoporder
stoporder.BrokerId = ""    ' force a new id to be allocated
placeOrder stoporder, "", ocaGroup, False

Set TargetOrder = pOrderPlex.TargetOrder
TargetOrder.BrokerId = "" ' force a new id to be allocated
placeOrder TargetOrder, "", ocaGroup, True
End Sub

Private Sub IOrderSubmitter_resubmitStopOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim stoporder As Order
Dim TargetOrder As Order

Set stoporder = pOrderPlex.stoporder
Set TargetOrder = pOrderPlex.TargetOrder
stoporder.BrokerId = ""     ' force a new id to be allocated
If Not pOrderPlex.TargetOrder Is Nothing Then
    placeOrder stoporder, "", mSimulator.GetOrder(TargetOrder.Id), True
Else
    placeOrder stoporder, "", "", True
End If
End Sub

Private Sub IOrderSubmitter_resubmitTargetOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim stoporder As Order
Dim TargetOrder As Order

Set stoporder = pOrderPlex.stoporder
Set TargetOrder = pOrderPlex.TargetOrder
TargetOrder.BrokerId = ""   ' force a new id to be allocated
If Not pOrderPlex.stoporder Is Nothing Then
    placeOrder TargetOrder, "", mSimulator.GetOrder(stoporder.Id), True
Else
    placeOrder TargetOrder, "", "", True
End If
End Sub

'@================================================================================
' IStreamingDataConsumer Interface Members
'@================================================================================

Private Sub IStreamingDataConsumer_Tick(Tick As TickUtils26.GenericTick)
Select Case Tick.tickType
Case TickTypes.TickTypeAsk
    mFirstAskReceived = True
    If mFirstBidReceived And Not mReadyNotified Then
        mReadyNotified = True
        mServiceConsumer.ready
    End If
        
    mSimulator.Ask Tick.Timestamp, mMarketIndex, Tick.Price
Case TickTypes.TickTypeBid
    mFirstBidReceived = True
    If mFirstAskReceived And Not mReadyNotified Then
        mReadyNotified = True
        mServiceConsumer.ready
    End If
        
    mSimulator.Bid Tick.Timestamp, mMarketIndex, Tick.Price
Case TickTypes.TickTypeClosePrice
Case TickTypes.TickTypeHighPrice
Case TickTypes.TickTypeLowPrice
Case TickTypes.TickTypeMarketDepth
Case TickTypes.TickTypeMarketDepthReset
Case TickTypes.TickTypeOpenInterest
Case TickTypes.TickTypeOpenPrice
Case TickTypes.TickTypeTrade
    mSimulator.Trade Tick.Timestamp, mMarketIndex, Tick.Price
Case TickTypes.TickTypeVolume
End Select

End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Contract(ByVal value As Contract)
Set mContract = value
End Property

Friend Property Let serviceConsumer(ByVal value As TradeBuildSP.IOrderSubmissionSrvcConsumer)
Set mServiceConsumer = value
mServiceConsumer.RequestTickData Me, False
End Property

Friend Property Let simulator(ByVal value As AdvancedOrderSimulator)
Set mSimulator = value
mMarketIndex = mSimulator.startMarket(mServiceConsumer, _
                                    mContract)
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub placeOrder( _
                ByVal pOrder As Order, _
                ByVal parentId As String, _
                ByVal ocaGroup As String, _
                ByVal transmit As Boolean)
mSimulator.placeOrder pOrder, parentId, ocaGroup, mMarketIndex, transmit
If pOrder.Status = OrderStatusCreated Then mServiceConsumer.OrderStatus pOrder.Id, OrderStatusPendingSubmit
End Sub

