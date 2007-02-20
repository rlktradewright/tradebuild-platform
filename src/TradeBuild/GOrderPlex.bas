Attribute VB_Name = "GOrderPlex"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

' This condition indicates that the cancellation of the order plex
' has been requested, provided that the entry order has not been filled
Public Const COND_NO_FILL_CANCELLATION As Long = &H1&

' This condition indicates that a notification that the stop order has been
' cancelled has been received via the API. This can be in the form of either
' an orderStatus message with status 'cancelled', or an errorMessage with
' errorCode = 202, or an errorMessage with errorCode = 201 (indicating
' that the order has been rejected for some reason).
Public Const COND_STOP_ORDER_CANCELLED As Long = &H2&

' This condition indicates that a notification that the target order has been
' cancelled has been received via the API. This can be in the form of either
' an orderStatus message with status 'cancelled', or an errorMessage with
' errorCode = 202, or an errorMessage with errorCode = 201 (indicating
' that the order has been rejected for some reason).
Public Const COND_TARGET_ORDER_CANCELLED As Long = &H4&

' This condition indicates that the stop order exists.
Public Const COND_STOP_ORDER_EXISTS As Long = &H8&

' This condition indicates that the target order exists.
Public Const COND_TARGET_ORDER_EXISTS As Long = &H10&

' This condition indicates that the entry order has been partially or
' completely filled.
Public Const COND_SIZE_NON_ZERO As Long = &H20&

Public Const DummyOffset As Long = &H7FFFFFFF

'@================================================================================
' Enums
'@================================================================================

Public Enum StateTransitionStimuli
    
    ' This stimulus indicates that the application has requested that
    ' the order plex be executed
    STIM_EXECUTE = 1
    
    ' This stimulus indicates that the application has requested that
    ' the order plex be cancelled provided the entry order has not already
    ' been fully or partially filled. If the entry order is filled during
    ' cancelling, then the stop and target orders (if they exist) must
    ' remain in place
    STIM_CANCEL_IF_NO_FILL
    
    ' This stimulus indicates that the application has requested that
    ' the order plex be cancelled even if the entry order has already been
    ' fully or partially filled. If the entry order is filled during
    ' cancelling, then the stop and target orders (if they exist) must
    ' nevertheless be cancelled.
    STIM_CANCEL_EVEN_IF_FILL
    
    ' This stimulus indicates that the application has requested that
    ' the order plex be closed out, ie that any outstanding orders be
    ' cancelled and that if the order plex then has a non-zero size, then
    ' a closeout order be submitted to reduce the size to zero.
    STIM_CLOSEOUT
    
    ' This stimulus indicates that the all the orders in the order plex
    ' have been completed (ie either fully filled, or cancelled). Note
    ' that this includes the closeout order where appropriate.
    STIM_ALL_ORDERS_COMPLETE
    
    ' This stimulus indicates that the API has generated a notification
    ' that the entry order has been cancelled. This can be in the
    ' form of either an orderStatus message with status 'cancelled', or
    ' an errorMessage with errorCode = 202, or an errorMessage with
    ' errorCode = 201 (indicating that the order has been rejected for
    ' some reason).
    STIM_ENTRY_ORDER_CANCELLED
    
    ' This stimulus indicates that the API has generated a notification
    ' that the stop order has been cancelled. This can be in the
    ' form of either an orderStatus message with status 'cancelled', or
    ' an errorMessage with errorCode = 202, or an errorMessage with
    ' errorCode = 201 (indicating that the order has been rejected for
    ' some reason).
    STIM_STOP_ORDER_CANCELLED
    
    ' This stimulus indicates that the API has generated a notification
    ' that the closeout order has been cancelled. This can be in the
    ' form of either an orderStatus message with status 'cancelled', or
    ' an errorMessage with errorCode = 202, or an errorMessage with
    ' errorCode = 201 (indicating that the order has been rejected for
    ' some reason). Note that this is a very unpleasant situation, since
    ' it only occurs when attempting to closeout a position and it leaves
    ' us with an unprotected position.
    STIM_CLOSEOUT_ORDER_CANCELLED
    
    ' This stimulus indicates that the API has generated a notification
    ' that the target order has been cancelled. This can be in the
    ' form of either an orderStatus message with status 'cancelled', or
    ' an errorMessage with errorCode = 202, or an errorMessage with
    ' errorCode = 201 (indicating that the order has been rejected for
    ' some reason).
    
    STIM_TARGET_ORDER_CANCELLED
    ' This stimulus indicates that the API has generated a notification
    ' that the entry order has been filled.
    STIM_ENTRY_ORDER_FILL
End Enum

Public Enum Actions
    
    ' This action places all orders defined in the order plex.
    ACT_PLACE_ORDERS = 1
    
    ' This action cancels all outstanding orders whose current status
    ' indicates that they are not already either filled, cancelled or
    ' cancelling. Note that where an order has not yet been placed, there
    ' may still be work to do, for example logging or notifying listeners.
    ACT_CANCEL_ORDERS
    
    ' This action cancels the stop order if it exists and its current
    ' status indicates that it is not already either filled, cancelled or
    ' cancelling. Note that where the order has not yet been placed,
    ' there may still be work to do, for example logging or notifying
    ' listeners.
    ACT_CANCEL_STOP_ORDER
    
    ' This action cancels the target order if it exists and its current
    ' status indicates that it is not already either filled, cancelled or
    ' cancelling. Note that where the order has not yet been placed,
    ' there may still be work to do, for example logging or notifying
    ' listeners.
    ACT_CANCEL_TARGET_ORDER
    
    ' This action resubmits the stop order (with a new order id). If a
    ' target order exists, then the ocaGroup of the stop order is set to
    ' the ocaGroup of the target order
    ACT_RESUBMIT_STOP_ORDER
    
    ' This action resubmits the target order (with a new order id). If a
    ' stop order exists, then the ocaGroup of the target order is set to
    ' the ocaGroup of the stop order
    ACT_RESUBMIT_TARGET_ORDER
    
    ' This action resubmits the both the stop and target orders (with new
    ' order ids and a new ocaGroup).
    ACT_RESUBMIT_STOP_AND_TARGET_ORDERS
    
    ' This action creates and places an orders whose effect is to cancel
    ' any existing size already acquired by this order plex. For example,
    ' if the order plex is currently long 1 contract, the closeout order
    ' must sell 1 contract.
    ACT_PLACE_CLOSEOUT_ORDER
    
    ' This action causes an alarm to be generated (for example, audible
    ' sound, on-screen alert, email, SMS etc).
    ACT_ALARM
    
    ' This action performs any tidying up needed when an order plex is
    ' completed.
    ACT_COMPLETION_ACTIONS
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Global object references
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Variables
'@================================================================================

Private mTableBuilder As StateTableBuilder

'@================================================================================
' Properties
'@================================================================================

Public Property Get tableBuilder() As StateTableBuilder
If mTableBuilder Is Nothing Then
    Set mTableBuilder = New StateTableBuilder
    buildStateTable
End If
Set tableBuilder = mTableBuilder
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gNextApplicationIndex() As Long
Static lNextApplicationIndex As Long

gNextApplicationIndex = lNextApplicationIndex
lNextApplicationIndex = lNextApplicationIndex + 1
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub buildStateTable()

'=======================================================================
'                       State:      OrderPlexStateCodes.OrderPlexStateCreated
'=======================================================================

' The application requests that the order plex be cancelled provided no
' fills have occurred. Since the orders have not yet been placed, we
' merely cancel the orders. do any tidying up and set the state to closed.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCreated, _
            STIM_CANCEL_IF_NO_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            ACT_CANCEL_ORDERS, ACT_COMPLETION_ACTIONS
            
' The application requests that the order plex be cancelled even if
' fills have occurred. Since the orders have not yet been placed, we
' merely cancel the orders. do any tidying up and set the state to closed.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCreated, _
            STIM_CANCEL_EVEN_IF_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            ACT_CANCEL_ORDERS, ACT_COMPLETION_ACTIONS
            
' The application requests that the order plex be executed. We do that and
' go to submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCreated, _
            STIM_EXECUTE, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            ACT_PLACE_ORDERS

'=======================================================================
'                       State:      OrderPlexStateCodes.OrderPlexStateSubmitted
'=======================================================================

' TWS tells us that the entry order has been filled. Nothing to do here.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            STIM_ENTRY_ORDER_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateSubmitted

' All orders have been completed, so we set the state to closed and do any
' tidying up.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            STIM_ALL_ORDERS_COMPLETE, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            ACT_COMPLETION_ACTIONS

' The application requests that the order plex be cancelled provided no fill
' have occurred. But a fill has already occurred, so we do nothing.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            STIM_CANCEL_IF_NO_FILL, _
            COND_SIZE_NON_ZERO, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateSubmitted

' The application requests that the order plex be cancelled provided no fill
' have occurred. No fills have already occurred, so we cancel all the orders
' and enter the cancelling state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            STIM_CANCEL_IF_NO_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            COND_SIZE_NON_ZERO, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            ACT_CANCEL_ORDERS

' The application requests that the order plex be cancelled even if fills
' have occurred. We cancel all the orders and enter the cancelling state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            STIM_CANCEL_EVEN_IF_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            ACT_CANCEL_ORDERS

' We are notified that the entry order has been cancelled (for example it
' may have been rejected by TWS or the user may have cancelled it at TWS).
' There has been no fill, so we cancel the stop and target orders (not
' really necessary, since TWS should do this, but just in case...).
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            STIM_ENTRY_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            COND_SIZE_NON_ZERO, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            ACT_CANCEL_STOP_ORDER, ACT_CANCEL_TARGET_ORDER

' We are notified that the entry order has been cancelled (for example the
' user may have cancelled it at TWS). Note that it can't be the application
' that cancelled it because it has no way of cancelling individual orders.
' The cancellation will have caused the stop and/or target orders to be
' cancelled as well (though we haven't been notified of this yet). Therefore
' we'll be left with an unprotected position, so we cancel the stop and target
' orders (just in case) and go into closing out state to negate the
' unprotected position.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            STIM_ENTRY_ORDER_CANCELLED, _
            COND_SIZE_NON_ZERO, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            ACT_CANCEL_STOP_ORDER, ACT_CANCEL_TARGET_ORDER

' We are notified that the stop order has been cancelled. This could be because
' it has been rejected by TWS, or because the user has cancelled it at TWS. We
' can't tell which of these is the case, so we cancel all orders and go into
' closing out state, because the entry order could be filled before being
' cancelled, and closing out will prevent an unprotected position.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            STIM_STOP_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            ACT_CANCEL_ORDERS

' We are notified that the target order has been cancelled. This could be because
' it has been rejected by TWS, or because the user has cancelled it at TWS. We
' can't tell which of these is the case, so we cancel all orders and go into
' closing out state, because the entry order could be filled before being
' cancelled, and closing out will prevent an unprotected position.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            STIM_TARGET_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            ACT_CANCEL_ORDERS

' The application has requested that the order plex be closed out. So cancel any
' outstanding orders and go to closing out state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            STIM_CLOSEOUT, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            ACT_CANCEL_ORDERS
            
'=======================================================================
'                       State:      OrderPlexStateCodes.OrderPlexStateCancelling
'=======================================================================

' The application has requested that the order plex be cancelled, provided
' there have been no fills. Since it is already being cancelled, there is
' nothing to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_CANCEL_IF_NO_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateCancelling

' The application has requested that the order plex be cancelled, even if
' there have already been fills. Since it is already being cancelled, there
' is nothing to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_CANCEL_EVEN_IF_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateCancelling

' All orders have now been completed, so do any tidying up and go to the
' closed state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ALL_ORDERS_COMPLETE, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            ACT_COMPLETION_ACTIONS

' We are notified that the entry order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ENTRY_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateCancelling

' We are notified that the stop order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_STOP_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateCancelling

' We are notified that the target order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_TARGET_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateCancelling

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). Since the original
' cancellation request from the application was to cancel even if there have
' been some fills, we just continue with the cancellation by re-requesting
' cancellation of any outstanding orders.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ENTRY_ORDER_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            COND_NO_FILL_CANCELLATION, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            ACT_CANCEL_ORDERS

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There are no stop or target orders,
' so we just return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION, _
            COND_STOP_ORDER_EXISTS + COND_TARGET_ORDER_EXISTS, _
            OrderPlexStateCodes.OrderPlexStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order but no target
' order, and the stop order has not been cancelled, so we just return to the
' submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_EXISTS, _
            COND_STOP_ORDER_CANCELLED + COND_TARGET_ORDER_EXISTS, _
            OrderPlexStateCodes.OrderPlexStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order but no target
' order, and the stop order has been cancelled, so we resubmit the stop order
' and return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_CANCELLED, _
            COND_TARGET_ORDER_EXISTS, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            ACT_RESUBMIT_STOP_ORDER

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a target order but no stop
' order, and the tartget order has not been cancelled, so we return to the
' submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_TARGET_ORDER_EXISTS, _
            COND_TARGET_ORDER_CANCELLED + COND_STOP_ORDER_EXISTS, _
            OrderPlexStateCodes.OrderPlexStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, but neither has been cancelled, so we return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_EXISTS + COND_TARGET_ORDER_EXISTS, _
            COND_STOP_ORDER_CANCELLED + COND_TARGET_ORDER_CANCELLED, _
            OrderPlexStateCodes.OrderPlexStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, and the stop order has been cancelled but not the target order, so we
' resubmit the stop order and return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_CANCELLED + COND_TARGET_ORDER_EXISTS, _
            COND_TARGET_ORDER_CANCELLED, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            ACT_RESUBMIT_STOP_ORDER

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a target order but no stop
' order, and the target order has been cancelled, so we resubmit the target
' order and return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_TARGET_ORDER_CANCELLED, _
            COND_STOP_ORDER_EXISTS, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            ACT_RESUBMIT_TARGET_ORDER

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, and the target order has been cancelled but not the stop order, so we
' resubmit the target order and return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_EXISTS + COND_TARGET_ORDER_CANCELLED, _
            COND_STOP_ORDER_CANCELLED, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            ACT_RESUBMIT_TARGET_ORDER

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, and both have been cancelled, so we resubmit both the stop order and
' the target order, and return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_CANCELLED + COND_TARGET_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            ACT_RESUBMIT_STOP_AND_TARGET_ORDERS
            
            
'=======================================================================
'                       State:      OrderPlexStateCodes.OrderPlexStateClosingOut
'=======================================================================

' The entry order has been cancelled, nothing for us to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            STIM_ENTRY_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosingOut

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the orders and TWS's cancellation
' request arriving at the IB servers or the exchange). There is nothing for
' us to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            STIM_ENTRY_ORDER_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosingOut

' The stop order has been cancelled, nothing for us to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            STIM_STOP_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosingOut

' The target order has been cancelled, nothing for us to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            STIM_TARGET_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosingOut

' All orders have completed, and we are left with a non-zero size. So submit
' a closeout order to reduce the size to zero. Stay in this state awaiting the
' next 'all orders complete' stimulus.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            STIM_ALL_ORDERS_COMPLETE, _
            COND_SIZE_NON_ZERO, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            ACT_PLACE_CLOSEOUT_ORDER

' All orders have completed, and we are left with a zero size. We are done,
' so go to the closed state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            STIM_ALL_ORDERS_COMPLETE, _
            SpecialConditions.NO_CONDITIONS, _
            COND_SIZE_NON_ZERO, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            ACT_COMPLETION_ACTIONS

' The closeout order has been cancelled (presumably it has been rejected
' by TWS). This is a serious situation since we are left with an unprotected
' position, so raise an alarm.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            STIM_CLOSEOUT_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            ACT_ALARM, ACT_COMPLETION_ACTIONS

mTableBuilder.stateTableComplete
End Sub

