VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderSubmissionSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Private Function IOrderSubmitter_cancelOrderPlex(ByVal pOrderPlex As OrderPlex) As Boolean
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
    gTradeBuildAPI.cancelOrder entryOrder.brokerId  ' should automatically cancel the other orders
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
            gTradeBuildAPI.cancelOrder stoporder.brokerId
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
            gTradeBuildAPI.cancelOrder targetOrder.brokerId
            IOrderSubmitter_cancelOrderPlex = True
        End Select
    End Select
End If
    
' need some stuff here to cancel if it's an oca group

End Function

Private Sub IOrderSubmitter_cancelStopOrder(ByVal pOrderPlex As OrderPlex)
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
    gTradeBuildAPI.cancelOrder stoporder.brokerId
End Select
End Sub

Private Sub IOrderSubmitter_cancelTargetOrder(ByVal pOrderPlex As OrderPlex)
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
    gTradeBuildAPI.cancelOrder targetOrder.brokerId
End Select
End Sub

Private Sub IOrderSubmitter_executeOrderPlex(ByVal pOrderPlex As OrderPlex)
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
        
    gTradeBuildAPI.placeOrder entryOrder, 0, transmit
    
    If Not stoporder Is Nothing Then
        transmit = True
        If Not targetOrder Is Nothing Then
            If targetOrder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoLimit Then transmit = False
        End If
        
        If stoporder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry order is filled
            gTradeBuildAPI.placeOrder stoporder, entryOrder.brokerId, transmit
        End If
    End If
    
    If Not targetOrder Is Nothing Then
        If targetOrder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoLimit Then
            gTradeBuildAPI.placeOrder targetOrder, entryOrder.brokerId, True
        End If
    End If
    
Else
    ' treat the other orders as an OCA group - still to be implemented
End If
End Sub

Private Sub IOrderSubmitter_modifyOrderPlex( _
                ByVal pOrderPlex As OrderPlex, _
                ByVal entryOrderChanged As Boolean, _
                ByVal stopOrderChanged As Boolean, _
                ByVal targetOrderChanged As Boolean, _
                ByVal closeoutOrderChanged As Boolean)
If entryOrderChanged Then
    gTradeBuildAPI.placeOrder pOrderPlex.entryOrder, 0, True
End If
If stopOrderChanged Then
    gTradeBuildAPI.placeOrder pOrderPlex.stoporder, pOrderPlex.entryOrder.brokerId, True
End If
If targetOrderChanged Then
    gTradeBuildAPI.placeOrder pOrderPlex.targetOrder, pOrderPlex.entryOrder.brokerId, True
End If
If closeoutOrderChanged Then
    gTradeBuildAPI.placeOrder pOrderPlex.closeoutOrder, 0, True
End If
    
End Sub

Private Sub IOrderSubmitter_resubmitStopAndTargetOrders(ByVal pOrderPlex As OrderPlex)
Dim ocaGroup As String
Dim stoporder As order
Dim targetOrder As order

ocaGroup = GenerateTextID

Set stoporder = pOrderPlex.stoporder
stoporder.ocaGroup = ocaGroup
stoporder.id = ""   ' force a new id to be allocated
gTradeBuildAPI.placeOrder stoporder, 0, False

Set targetOrder = pOrderPlex.targetOrder
targetOrder.ocaGroup = ocaGroup
targetOrder.id = "" ' force a new id to be allocated
gTradeBuildAPI.placeOrder targetOrder, 0, True
End Sub

Private Sub IOrderSubmitter_resubmitStopOrder(ByVal pOrderPlex As OrderPlex)
Dim stoporder As order

Set stoporder = pOrderPlex.stoporder
stoporder.brokerId = ""     ' force a new id to be allocated
If Not pOrderPlex.targetOrder Is Nothing Then
    stoporder.ocaGroup = pOrderPlex.targetOrder.ocaGroup
End If
gTradeBuildAPI.placeOrder stoporder, 0, True
End Sub

Private Sub IOrderSubmitter_resubmitTargetOrder(ByVal pOrderPlex As OrderPlex)
Dim targetOrder As order

Set targetOrder = pOrderPlex.targetOrder
targetOrder.brokerId = ""   ' force a new id to be allocated
If Not pOrderPlex.stoporder Is Nothing Then
    targetOrder.ocaGroup = pOrderPlex.stoporder.ocaGroup
End If
gTradeBuildAPI.placeOrder targetOrder, 0, True
End Sub

'================================================================================
' Helper Functions
'================================================================================

