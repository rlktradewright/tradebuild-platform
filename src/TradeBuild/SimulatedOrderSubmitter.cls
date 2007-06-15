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
Dim entryOrder As order
Dim stoporder As order
Dim targetOrder As order

Set entryOrder = pOrderPlex.entryOrder
Select Case entryOrder.status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    mSimulator.cancelOrder entryOrder.brokerId  ' should automatically cancel the other orders
                                                    ' if they have parentid set
    IOrderSubmitter_cancelOrderPlex = True
End Select

Set stoporder = pOrderPlex.stoporder
If Not stoporder Is Nothing Then
    Select Case stoporder.status
    Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case entryOrder.status
        Case OrderStatuses.OrderStatusFilled, _
            OrderStatuses.OrderStatusCancelling, _
            OrderStatuses.OrderStatusCancelled
            mSimulator.cancelOrder stoporder.brokerId
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If

Set targetOrder = pOrderPlex.targetOrder
If Not targetOrder Is Nothing Then
    Select Case targetOrder.status
    Case OrderStatuses.OrderStatusCreated, _
        OrderStatuses.OrderStatusFilled, _
        OrderStatuses.OrderStatusCancelling, _
        OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case entryOrder.status
        Case OrderStatuses.OrderStatusFilled, _
            OrderStatuses.OrderStatusCancelling, _
            OrderStatuses.OrderStatusCancelled
            mSimulator.cancelOrder targetOrder.brokerId
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If
    
' need some stuff here to cancel if it's an oca group

End Function

Private Sub IOrderSubmitter_cancelStopOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim stoporder As order

Set stoporder = pOrderPlex.stoporder
If stoporder Is Nothing Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelStopOrder", _
                                    "Order plex has no stop order"

Select Case stoporder.status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelStopOrder", _
                                    "Stop order state invalid for cancellation"
Case Else
    mSimulator.cancelOrder stoporder.brokerId
End Select
End Sub

Private Sub IOrderSubmitter_cancelTargetOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim targetOrder As order

Set targetOrder = pOrderPlex.targetOrder
If targetOrder Is Nothing Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelTargetOrder", _
                                    "Order plex has no target order"


Select Case targetOrder.status
Case OrderStatuses.OrderStatusCreated, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
    err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelTargetOrder", _
                                    "Target order state invalid for cancellation"
Case Else
    mSimulator.cancelOrder targetOrder.brokerId
End Select
End Sub

Private Property Let IOrderSubmitter_clockRate(ByVal RHS As Single)
mSimulator.clockRate(mMarketIndex) = RHS
End Property

Private Sub IOrderSubmitter_executeOrderPlex(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim entryOrder As order
Dim stoporder As order
Dim targetOrder As order
Dim transmit As Boolean
Dim ocaGroup As String

Set entryOrder = pOrderPlex.entryOrder
Set stoporder = pOrderPlex.stoporder
Set targetOrder = pOrderPlex.targetOrder

If Not entryOrder Is Nothing Then
    transmit = True
    If Not stoporder Is Nothing Then
        If stoporder.orderType <> OrderTypes.OrderTypeAutoStop Then transmit = False
    End If
    If Not targetOrder Is Nothing Then
        If targetOrder.orderType <> OrderTypes.OrderTypeAutoLimit Then transmit = False
    End If
        
    placeOrder entryOrder, "", "", transmit
    
    ocaGroup = GenerateTextID

    If Not stoporder Is Nothing Then
        transmit = True
        If Not targetOrder Is Nothing Then
            If targetOrder.orderType <> OrderTypes.OrderTypeAutoLimit Then transmit = False
        End If
        
        If stoporder.orderType <> OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry order is filled
            placeOrder stoporder, entryOrder.id, ocaGroup, transmit
        End If
    End If
    
    If Not targetOrder Is Nothing Then
        If targetOrder.orderType <> OrderTypes.OrderTypeAutoLimit Then
            placeOrder targetOrder, entryOrder.id, ocaGroup, True
        End If
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If
End Sub

Private Sub IOrderSubmitter_finish()
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
Dim lOrder As order

If entryOrderChanged Then
    placeOrder pOrderPlex.entryOrder, "", "", True
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
    placeOrder pOrderPlex.closeoutOrder, "", "", True
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
Dim stoporder As order
Dim targetOrder As order

ocaGroup = GenerateTextID

Set stoporder = pOrderPlex.stoporder
stoporder.brokerId = ""    ' force a new id to be allocated
placeOrder stoporder, "", ocaGroup, False

Set targetOrder = pOrderPlex.targetOrder
targetOrder.brokerId = "" ' force a new id to be allocated
placeOrder targetOrder, "", ocaGroup, True
End Sub

Private Sub IOrderSubmitter_resubmitStopOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim stoporder As order
Dim targetOrder As order

Set stoporder = pOrderPlex.stoporder
Set targetOrder = pOrderPlex.targetOrder
stoporder.brokerId = ""     ' force a new id to be allocated
If Not pOrderPlex.targetOrder Is Nothing Then
    placeOrder stoporder, "", mSimulator.getOrder(targetOrder.id), True
Else
    placeOrder stoporder, "", "", True
End If
End Sub

Private Sub IOrderSubmitter_resubmitTargetOrder(ByVal pOrderPlex As TradeBuildSP.IOrderPlex)
Dim stoporder As order
Dim targetOrder As order

Set stoporder = pOrderPlex.stoporder
Set targetOrder = pOrderPlex.targetOrder
targetOrder.brokerId = ""   ' force a new id to be allocated
If Not pOrderPlex.stoporder Is Nothing Then
    placeOrder targetOrder, "", mSimulator.getOrder(stoporder.id), True
Else
    placeOrder targetOrder, "", "", True
End If
End Sub

'@================================================================================
' IStreamingDataConsumer Interface Members
'@================================================================================

Private Sub IStreamingDataConsumer_Ask(ByVal timestamp As Date, ByVal price As Double, ByVal size As Long)
mSimulator.ask timestamp, mMarketIndex, price
End Sub

Private Sub IStreamingDataConsumer_Bid(ByVal timestamp As Date, ByVal price As Double, ByVal size As Long)
mSimulator.bid timestamp, mMarketIndex, price
End Sub

Private Sub IStreamingDataConsumer_High(ByVal timestamp As Date, ByVal price As Double)

End Sub

Private Sub IStreamingDataConsumer_Low(ByVal timestamp As Date, ByVal price As Double)

End Sub

Private Sub IStreamingDataConsumer_OpenInterest(ByVal timestamp As Date, ByVal size As Long)

End Sub

Private Sub IStreamingDataConsumer_PreviousClose(ByVal timestamp As Date, ByVal price As Double)

End Sub

Private Sub IStreamingDataConsumer_ResetMarketDepth(ByVal timestamp As Date, ByVal marketDepthReRequested As Boolean)

End Sub

Private Sub IStreamingDataConsumer_Trade(ByVal timestamp As Date, ByVal price As Double, ByVal size As Long)
mSimulator.trade timestamp, mMarketIndex, price
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepth(ByVal timestamp As Date, ByVal position As Long, ByVal operation As TradeBuildSP.DOMOperations, ByVal side As TradeBuildSP.DOMSides, ByVal price As Double, ByVal size As Long)

End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepthL2(ByVal timestamp As Date, ByVal position As Long, ByVal marketMaker As String, ByVal operation As TradeBuildSP.DOMOperations, ByVal side As TradeBuildSP.DOMSides, ByVal price As Double, ByVal size As Long)

End Sub

Private Sub IStreamingDataConsumer_Volume(ByVal timestamp As Date, ByVal size As Long)

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
mServiceConsumer.requestTickData Me, False
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
                ByVal pOrder As order, _
                ByVal parentId As String, _
                ByVal ocaGroup As String, _
                ByVal transmit As Boolean)
mSimulator.placeOrder pOrder, parentId, ocaGroup, mMarketIndex, transmit
End Sub

