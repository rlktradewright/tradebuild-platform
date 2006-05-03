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
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements IOrderSubmitter
Implements TradeBuildSP.IStreamingDataConsumer

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mServiceConsumer As IOrderSubmissionSrvcConsumer

Private mSimulator As TradeBuild.AdvancedOrderSimulator

Private mMarketIndex As Long

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IOrderSubmitter Interface Members
'================================================================================

Private Function IOrderSubmitter_cancelOrderPlex(ByVal pOrderPlex As orderPlex) As Boolean
Dim entryOrder As order
Dim stoporder As order
Dim targetOrder As order

Set entryOrder = pOrderPlex.entryOrder
Select Case entryOrder.status
Case TradeBuild.OrderStatuses.OrderStatusCreated, _
    TradeBuild.OrderStatuses.OrderStatusFilled, _
    TradeBuild.OrderStatuses.OrderStatusCancelling, _
    TradeBuild.OrderStatuses.OrderStatusCancelled
Case Else
    mSimulator.cancelOrder entryOrder.brokerId  ' should automatically cancel the other orders
                                                    ' if they have parentid set
    IOrderSubmitter_cancelOrderPlex = True
End Select

Set stoporder = pOrderPlex.stoporder
If Not stoporder Is Nothing Then
    Select Case stoporder.status
    Case TradeBuild.OrderStatuses.OrderStatusCreated, _
        TradeBuild.OrderStatuses.OrderStatusFilled, _
        TradeBuild.OrderStatuses.OrderStatusCancelling, _
        TradeBuild.OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case entryOrder.status
        Case TradeBuild.OrderStatuses.OrderStatusFilled, _
            TradeBuild.OrderStatuses.OrderStatusCancelling, _
            TradeBuild.OrderStatuses.OrderStatusCancelled
            mSimulator.cancelOrder stoporder.brokerId
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If

Set targetOrder = pOrderPlex.targetOrder
If Not targetOrder Is Nothing Then
    Select Case targetOrder.status
    Case TradeBuild.OrderStatuses.OrderStatusCreated, _
        TradeBuild.OrderStatuses.OrderStatusFilled, _
        TradeBuild.OrderStatuses.OrderStatusCancelling, _
        TradeBuild.OrderStatuses.OrderStatusCancelled
    Case Else
        Select Case entryOrder.status
        Case TradeBuild.OrderStatuses.OrderStatusFilled, _
            TradeBuild.OrderStatuses.OrderStatusCancelling, _
            TradeBuild.OrderStatuses.OrderStatusCancelled
            mSimulator.cancelOrder targetOrder.brokerId
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If
    
' need some stuff here to cancel if it's an oca group

End Function

Private Sub IOrderSubmitter_cancelStopOrder(ByVal pOrderPlex As orderPlex)
Dim stoporder As order

Set stoporder = pOrderPlex.stoporder
If stoporder Is Nothing Then err.Raise ErrorCodes.IllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelStopOrder", _
                                    "Order plex has no stop order"

Select Case stoporder.status
Case TradeBuild.OrderStatuses.OrderStatusCreated, _
    TradeBuild.OrderStatuses.OrderStatusFilled, _
    TradeBuild.OrderStatuses.OrderStatusCancelling, _
    TradeBuild.OrderStatuses.OrderStatusCancelled
    err.Raise ErrorCodes.IllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelStopOrder", _
                                    "Stop order state invalid for cancellation"
Case Else
    mSimulator.cancelOrder stoporder.brokerId
End Select
End Sub

Private Sub IOrderSubmitter_cancelTargetOrder(ByVal pOrderPlex As orderPlex)
Dim targetOrder As order

Set targetOrder = pOrderPlex.targetOrder
If targetOrder Is Nothing Then err.Raise ErrorCodes.IllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelTargetOrder", _
                                    "Order plex has no target order"


Select Case targetOrder.status
Case TradeBuild.OrderStatuses.OrderStatusCreated, _
    TradeBuild.OrderStatuses.OrderStatusCancelling, _
    TradeBuild.OrderStatuses.OrderStatusCancelled
    err.Raise ErrorCodes.IllegalStateException, _
                                    "OrderSubmissionServiceProvider::cancelTargetOrder", _
                                    "Target order state invalid for cancellation"
Case Else
    mSimulator.cancelOrder targetOrder.brokerId
End Select
End Sub

Private Sub IOrderSubmitter_executeOrderPlex(ByVal pOrderPlex As orderPlex)
Dim entryOrder As order
Dim stoporder As order
Dim targetOrder As order
Dim transmit As Boolean

Set entryOrder = pOrderPlex.entryOrder
Set stoporder = pOrderPlex.stoporder
Set targetOrder = pOrderPlex.targetOrder

If Not entryOrder Is Nothing Then
    transmit = True
    If Not stoporder Is Nothing Then
        If stoporder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoStop Then transmit = False
    End If
    If Not targetOrder Is Nothing Then
        If targetOrder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoLimit Then transmit = False
    End If
        
    placeOrder entryOrder, 0, transmit
    
    If Not stoporder Is Nothing Then
        transmit = True
        If Not targetOrder Is Nothing Then
            If targetOrder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoLimit Then transmit = False
        End If
        
        If stoporder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry order is filled
            placeOrder stoporder, entryOrder.brokerId, transmit
        End If
    End If
    
    If Not targetOrder Is Nothing Then
        If targetOrder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoLimit Then
            placeOrder targetOrder, entryOrder.brokerId, True
        End If
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If
End Sub

Private Sub IOrderSubmitter_modifyOrderPlex( _
                ByVal pOrderPlex As orderPlex, _
                ByVal entryOrderChanged As Boolean, _
                ByVal stopOrderChanged As Boolean, _
                ByVal targetOrderChanged As Boolean, _
                ByVal closeoutOrderChanged As Boolean)
If entryOrderChanged Then
    placeOrder pOrderPlex.entryOrder, 0, True
End If
If stopOrderChanged Then
    placeOrder pOrderPlex.stoporder, pOrderPlex.entryOrder.brokerId, True
End If
If targetOrderChanged Then
    placeOrder pOrderPlex.targetOrder, pOrderPlex.entryOrder.brokerId, True
End If
If closeoutOrderChanged Then
    placeOrder pOrderPlex.closeoutOrder, 0, True
End If
    
End Sub

Private Sub IOrderSubmitter_resubmitStopAndTargetOrders(ByVal pOrderPlex As orderPlex)
Dim ocaGroup As String
Dim stoporder As order
Dim targetOrder As order

ocaGroup = GenerateTextID

Set stoporder = pOrderPlex.stoporder
stoporder.ocaGroup = ocaGroup
stoporder.id = ""   ' force a new id to be allocated
placeOrder stoporder, 0, False

Set targetOrder = pOrderPlex.targetOrder
targetOrder.ocaGroup = ocaGroup
targetOrder.id = "" ' force a new id to be allocated
placeOrder targetOrder, 0, True
End Sub

Private Sub IOrderSubmitter_resubmitStopOrder(ByVal pOrderPlex As orderPlex)
Dim stoporder As order

Set stoporder = pOrderPlex.stoporder
stoporder.brokerId = ""     ' force a new id to be allocated
If Not pOrderPlex.targetOrder Is Nothing Then
    stoporder.ocaGroup = pOrderPlex.targetOrder.ocaGroup
End If
placeOrder stoporder, 0, True
End Sub

Private Sub IOrderSubmitter_resubmitTargetOrder(ByVal pOrderPlex As orderPlex)
Dim targetOrder As order

Set targetOrder = pOrderPlex.targetOrder
targetOrder.brokerId = ""   ' force a new id to be allocated
If Not pOrderPlex.stoporder Is Nothing Then
    targetOrder.ocaGroup = pOrderPlex.stoporder.ocaGroup
End If
placeOrder targetOrder, 0, True
End Sub

'================================================================================
' IStreamingDataConsumer Interface Members
'================================================================================

Private Sub IStreamingDataConsumer_Ask(ByVal Timestamp As Date, ByVal price As Double, ByVal size As Long)
mSimulator.ask Timestamp, mMarketIndex, price
End Sub

Private Sub IStreamingDataConsumer_Bid(ByVal Timestamp As Date, ByVal price As Double, ByVal size As Long)
mSimulator.bid Timestamp, mMarketIndex, price
End Sub

Private Sub IStreamingDataConsumer_High(ByVal Timestamp As Date, ByVal price As Double)

End Sub

Private Sub IStreamingDataConsumer_Low(ByVal Timestamp As Date, ByVal price As Double)

End Sub

Private Sub IStreamingDataConsumer_OpenInterest(ByVal Timestamp As Date, ByVal size As Long)

End Sub

Private Sub IStreamingDataConsumer_PreviousClose(ByVal Timestamp As Date, ByVal price As Double)

End Sub

Private Sub IStreamingDataConsumer_ResetMarketDepth(ByVal Timestamp As Date, ByVal marketDepthReRequested As Boolean)

End Sub

Private Sub IStreamingDataConsumer_Trade(ByVal Timestamp As Date, ByVal price As Double, ByVal size As Long)
mSimulator.trade Timestamp, mMarketIndex, price
End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepth(ByVal Timestamp As Date, ByVal Position As Long, ByVal operation As TradeBuildSP.DOMOperations, ByVal side As TradeBuildSP.DOMSides, ByVal price As Double, ByVal size As Long)

End Sub

Private Sub IStreamingDataConsumer_UpdateMktDepthL2(ByVal Timestamp As Date, ByVal Position As Long, ByVal marketMaker As String, ByVal operation As TradeBuildSP.DOMOperations, ByVal side As TradeBuildSP.DOMSides, ByVal price As Double, ByVal size As Long)

End Sub

Private Sub IStreamingDataConsumer_Volume(ByVal Timestamp As Date, ByVal size As Long)

End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let serviceConsumer(ByVal value As IOrderSubmissionSrvcConsumer)
Set mServiceConsumer = value
mServiceConsumer.requestTickData Me, False
End Property

Friend Property Let simulator(ByVal value As TradeBuild.AdvancedOrderSimulator)
Set mSimulator = value
mMarketIndex = mSimulator.startMarket(mServiceConsumer, mServiceConsumer.clockRate)
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub placeOrder( _
                ByVal pOrder As order, _
                ByVal parentId As String, _
                ByVal transmit As Boolean)
pOrder.marketIndex = mMarketIndex
mSimulator.placeOrder pOrder, parentId, transmit
End Sub

