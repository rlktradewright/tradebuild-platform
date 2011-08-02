Attribute VB_Name = "GBracketOrder"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const DummyOffset As Long = &H7FFFFFFF

'@================================================================================
' Enums
'@================================================================================

Public Enum OpActions
    
    ' This Action places all orders defined in the bracket order.
    ActPlaceOrders = 1
    
    ' This Action cancels all outstanding orders whose current status
    ' indicates that they are not already either filled, cancelled or
    ' cancelling. Note that where an order has not yet been placed, there
    ' may still be work to do, for example logging or notifying listeners.
    ActCancelOrders
    
    ' This Action cancels the stop order if it exists and its current
    ' status indicates that it is not already either filled, cancelled or
    ' cancelling. Note that where the order has not yet been placed,
    ' there may still be work to do, for example logging or notifying
    ' listeners.
    ActCancelStopOrder
    
    ' This Action cancels the target order if it exists and its current
    ' status indicates that it is not already either filled, cancelled or
    ' cancelling. Note that where the order has not yet been placed,
    ' there may still be work to do, for example logging or notifying
    ' listeners.
    ActCancelTargetOrder
    
    ' This Action resubmits the stop order (with a new order id). If a
    ' target order exists, then the ocaGroup of the stop order is set to
    ' the ocaGroup of the target order
    ActResubmitStopOrder
    
    ' This Action resubmits the target order (with a new order id). If a
    ' stop order exists, then the ocaGroup of the target order is set to
    ' the ocaGroup of the stop order
    ActResubmitTargetOrder
    
    ' This Action resubmits the both the stop and target orders (with new
    ' order ids and a new ocaGroup).
    ActResubmitStopAndTargetOrders
    
    ' This Action creates and places an orders whose effect is to cancel
    ' any existing Size already acquired by this bracket order. For example,
    ' if the bracket order is currently long 1 contract, the closeout order
    ' must sell 1 contract.
    ActPlaceCloseoutOrder
    
    ' This Action causes an alarm to be generated (for example, audible
    ' sound, on-screen alert, email, SMS etc).
    ActAlarm
    
    ' This Action performs any tidying up needed when an bracket order is
    ' completed.
    ActCompletionActions
    
    ' This Action causes a timeout stimulus to occur after a short time.
    ActSetTimeout

    ' This Action cancels a previously set timeout.
    ActCancelTimeout

End Enum

Public Enum OpConditions
    ' This condition indicates that the cancellation of the bracket order
    ' has been requested, provided that the entry order has not been filled
    CondNoFillCancellation = &H1&

    ' This condition indicates that a notification that the stop order has been
    ' cancelled has been received via the API. This can be in the form of either
    ' an orderStatus message with status 'cancelled', or an errorMessage with
    ' errorCode = 202, or an errorMessage with errorCode = 201 (indicating
    ' that the order has been rejected for some reason).
    CondStopOrderCancelled = &H2&

    ' This condition indicates that a notification that the target order has been
    ' cancelled has been received via the API. This can be in the form of either
    ' an orderStatus message with status 'cancelled', or an errorMessage with
    ' errorCode = 202, or an errorMessage with errorCode = 201 (indicating
    ' that the order has been rejected for some reason).
    CondTargetOrderCancelled = &H4&

    ' This condition indicates that the stop order exists.
    CondStopOrderExists = &H8&

    ' This condition indicates that the target order exists.
    CondTargetOrderExists = &H10&

    ' This condition indicates that the entry order has been partially or
    ' completely filled.
    CondSizeNonZero = &H20&
    
    ' This condition indicates that this bracket order is to prevent unprotected
    ' positions as far as possible.
    CondProtected = &H40&

End Enum

Public Enum OpStimuli
    
    ' This stimulus indicates that the application has requested that
    ' the bracket order be executed
    StimExecute = 1
    
    ' This stimulus indicates that the application has requested that
    ' the bracket order be cancelled provided the entry order has not already
    ' been fully or partially filled. If the entry order is filled during
    ' cancelling, then the stop and target orders (if they exist) must
    ' remain in place
    StimCancelIfNoFill
    
    ' This stimulus indicates that the application has requested that
    ' the bracket order be cancelled even if the entry order has already been
    ' fully or partially filled. If the entry order is filled during
    ' cancelling, then the stop and target orders (if they exist) must
    ' nevertheless be cancelled.
    StimCancelEvenIfFill
    
    ' This stimulus indicates that the application has requested that
    ' the bracket order be closed out, ie that any outstanding orders be
    ' cancelled and that if the bracket order then has a non-zero Size, then
    ' a closeout order be submitted to reduce the Size to zero.
    StimCloseout
    
    ' This stimulus indicates that all the orders in the bracket order
    ' have been completed (ie either fully filled, or cancelled). Note
    ' that this includes the closeout order where appropriate.
    StimAllOrdersComplete
    
    ' This stimulus indicates that the API has generated a notification
    ' that the entry order has been cancelled. This can be in the
    ' form of either an orderStatus message with status 'cancelled', or
    ' an errorMessage with errorCode = 202, or an errorMessage with
    ' errorCode = 201 (indicating that the order has been rejected for
    ' some reason).
    StimEntryOrderCancelled
    
    ' This stimulus indicates that the API has generated a notification
    ' that the stop order has been cancelled. This can be in the
    ' form of either an orderStatus message with status 'cancelled', or
    ' an errorMessage with errorCode = 202, or an errorMessage with
    ' errorCode = 201 (indicating that the order has been rejected for
    ' some reason).
    StimStopOrderCancelled
    
    ' This stimulus indicates that the API has generated a notification
    ' that the closeout order has been cancelled. This can be in the
    ' form of either an orderStatus message with status 'cancelled', or
    ' an errorMessage with errorCode = 202, or an errorMessage with
    ' errorCode = 201 (indicating that the order has been rejected for
    ' some reason). Note that this is a very unpleasant situation, since
    ' it only occurs when attempting to closeout a position and it leaves
    ' us with an unprotected position.
    StimCloseoutOrderCancelled
    
    ' This stimulus indicates that the API has generated a notification
    ' that the target order has been cancelled. This can be in the
    ' form of either an orderStatus message with status 'cancelled', or
    ' an errorMessage with errorCode = 202, or an errorMessage with
    ' errorCode = 201 (indicating that the order has been rejected for
    ' some reason).
    StimTargetOrderCancelled
    
    ' This stimulus indicates that the API has generated a notification
    ' that the entry order has been filled.
    StimEntryOrderFill

    ' This stimulus indicates that a state timeout has expired.
    StimTimeoutExpired

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

Public Property Get TableBuilder() As StateTableBuilder
If mTableBuilder Is Nothing Then
    Set mTableBuilder = New StateTableBuilder
    buildStateTable
End If
Set TableBuilder = mTableBuilder
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
'                       State:      BracketOrderStateCreated
'=======================================================================

' The application requests that the bracket order be cancelled provided no
' fills have occurred. Since the orders have not yet been placed, we
' merely cancel the orders. do any tidying up and set the state to closed.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCreated, _
            OpStimuli.StimCancelIfNoFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosed, _
            OpActions.ActCancelOrders, OpActions.ActCompletionActions
            
' The application requests that the bracket order be cancelled even if
' fills have occurred. Since the orders have not yet been placed, we
' merely cancel the BracketOrderStateCodesorders. do any tidying up and set the state to closed.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCreated, _
            OpStimuli.StimCancelEvenIfFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosed, _
            OpActions.ActCancelOrders, OpActions.ActCompletionActions
            
' The application requests that the bracket order be executed, and it is not
' protected. We do that and go to submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCreated, _
            OpStimuli.StimExecute, _
            SpecialConditions.NoConditions, _
            OpConditions.CondProtected, _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpActions.ActPlaceOrders

' The application requests that the bracket order be executed: it is
' protected and there is a stop order. We do that and go to submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCreated, _
            OpStimuli.StimExecute, _
            OpConditions.CondProtected Or OpConditions.CondStopOrderExists, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpActions.ActPlaceOrders

' The application requests that the bracket order be executed: it is
' protected and there is NO stop order. This is a programming error!
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCreated, _
            OpStimuli.StimExecute, _
            OpConditions.CondProtected, _
            OpConditions.CondStopOrderExists, _
            SpecialStates.StateError, _
            SpecialActions.NoAction


'=======================================================================
'                       State:      BracketOrderStateSubmitted
'=======================================================================

' TWS tells us that the entry order has been filled. Nothing to do here.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimEntryOrderFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateSubmitted

' All orders have been completed, so we set the state to closed and do any
' tidying up.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosed, _
            OpActions.ActCompletionActions

' The application requests that the bracket order be cancelled provided no fills
' have occurred. But a fill has already occurred, so we do nothing.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimCancelIfNoFill, _
            OpConditions.CondSizeNonZero, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateSubmitted

' The application requests that the bracket order be cancelled provided no fills
' have occurred. No fills have already occurred, so we cancel all the orders
' and enter the cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimCancelIfNoFill, _
            SpecialConditions.NoConditions, _
            OpConditions.CondSizeNonZero, _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpActions.ActCancelOrders

' The application requests that the bracket order be cancelled even if fills
' have occurred. We cancel all the orders and enter the cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimCancelEvenIfFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpActions.ActCancelOrders

' We are notified that the entry order has been cancelled (for example it
' may have been rejected by TWS or the user may have cancelled it at TWS).
' There has been no fill, so we cancel the stop and target orders (not
' really necessary, since TWS should do this, but just in case...).
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimEntryOrderCancelled, _
            SpecialConditions.NoConditions, _
            OpConditions.CondSizeNonZero, _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpActions.ActCancelStopOrder, OpActions.ActCancelTargetOrder

' We are notified that the entry order has been cancelled (for example the
' user may have cancelled it at TWS). Note that it can't be the application
' that cancelled it because it has no way of cancelling individual orders.
' There has been a fill. The cancellation will have caused the stop and/or
' target orders to be cancelled as well (though we haven't been notified of
' this yet), but we cancel them anyway just in case. We'll be left with an
' unprotected position, so as this is a protected bracket order, go into
' closing out state to negate the unprotected position.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimEntryOrderCancelled, _
            OpConditions.CondSizeNonZero Or OpConditions.CondProtected, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpActions.ActCancelStopOrder, OpActions.ActCancelTargetOrder

' We are notified that the entry order has been cancelled (for example the
' user may have cancelled it at TWS). Note that it can't be the application
' that cancelled it because it has no way of cancelling individual orders.
' There has been a fill. The cancellation will have caused the stop and/or
' target orders to be cancelled as well (though we haven't been notified of
' this yet), but we cancel them anyway just in case. We'll be left with an
' unprotected position, but since this is NOT a protected bracket order
' plex, go into Cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimEntryOrderCancelled, _
            OpConditions.CondSizeNonZero, _
            OpConditions.CondProtected, _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpActions.ActCancelStopOrder, OpActions.ActCancelTargetOrder

' We are notified that the stop order has been cancelled, and there is no target
' order. This could be because it has been rejected by TWS, or because the user
' has cancelled it at TWS. We can't tell which of these is the case, so we
' cancel all orders and, as this is a protected bracket order, go into closing out state,
' because the entry order could be filled before being cancelled, and closing out
' will prevent an unprotected position.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimStopOrderCancelled, _
            OpConditions.CondProtected, _
            OpConditions.CondTargetOrderExists, _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpActions.ActCancelOrders

' We are notified that the stop order has been cancelled, and there is no target
' order. This could be because it has been rejected by TWS, or because the user
' has cancelled it at TWS. We can't tell which of these is the case, so we
' cancel all orders and, as this is NOT a protected bracket order, go into cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            OpConditions.CondTargetOrderExists Or OpConditions.CondProtected, _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpActions.ActCancelOrders

' We are notified that the stop order has been cancelled, and there IS a target
' order. This could be because it has been rejected by TWS, or because the user
' has cancelled it at TWS, or because the target order has been filled. We can't
' tell which of these is the case, so, as this is a protected bracket order, we enter
' the 'awaiting other order cancel' state and set a timeout.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimStopOrderCancelled, _
            OpConditions.CondTargetOrderExists Or OpConditions.CondProtected, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateAwaitingOtherOrderCancel, _
            OpActions.ActSetTimeout

' We are notified that the stop order has been cancelled, and there IS a target
' order. This could be because it has been rejected by TWS, or because the user
' has cancelled it at TWS, or because the target order has been filled. We can't
' tell which of these is the case, but, as this is NOT a protected bracket order, we
' don't care so we cancel the target order and enter the cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimStopOrderCancelled, _
            OpConditions.CondTargetOrderExists, _
            OpConditions.CondProtected, _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpActions.ActCancelTargetOrder

' We are notified that the target order has been cancelled, and this is NOT a
' protected bracket order. Cancel all orders and go into cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            OpConditions.CondProtected, _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpActions.ActCancelOrders

' We are notified that the target order has been cancelled, and there IS a
' stop order. This could be because it has been rejected by TWS, or because the
' user has cancelled it at TWS, or because the stop order has been filled and
' not yet notified.  We can't tell which of these is the case, so, as this is
' a protected bracket order, we enter the 'awaiting other order cancel' state
' and set a timeout.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimTargetOrderCancelled, _
            OpConditions.CondStopOrderExists Or OpConditions.CondProtected, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateAwaitingOtherOrderCancel, _
            OpActions.ActSetTimeout

' The application has requested that the bracket order be closed out. So cancel any
' outstanding orders and go to closing out state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpStimuli.StimCloseout, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpActions.ActCancelOrders
            
'=======================================================================
'                       State:      BracketOrderStateCancelling
'=======================================================================

' The application has requested that the bracket order be cancelled, provided
' there have been no fills. Since it is already being cancelled, there is
' nothing to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimCancelIfNoFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateCancelling

' The application has requested that the bracket order be cancelled, even if
' there have already been fills. Since it is already being cancelled, there
' is nothing to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimCancelEvenIfFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateCancelling

' All orders have now been completed, so do any tidying up and go to the
' closed state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosed, _
            OpActions.ActCompletionActions

' We are notified that the entry order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateCancelling

' We are notified that the stop order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateCancelling

' We are notified that the target order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateCancelling

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). Since the original
' cancellation request from the application was to cancel even if there have
' been some fills, we just continue with the cancellation by re-requesting
' cancellation of any outstanding orders.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            SpecialConditions.NoConditions, _
            OpConditions.CondNoFillCancellation, _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpActions.ActCancelOrders

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There are no stop or target orders,
' so we just return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation, _
            OpConditions.CondStopOrderExists + OpConditions.CondTargetOrderExists, _
            BracketOrderStateCodes.BracketOrderStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order but no target
' order, and the stop order has not been cancelled, so we just return to the
' submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderExists, _
            OpConditions.CondStopOrderCancelled + OpConditions.CondTargetOrderExists, _
            BracketOrderStateCodes.BracketOrderStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order but no target
' order, and the stop order has been cancelled, so we resubmit the stop order
' and return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderCancelled, _
            OpConditions.CondTargetOrderExists, _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpActions.ActResubmitStopOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a target order but no stop
' order, and the tartget order has not been cancelled, so we return to the
' submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondTargetOrderExists, _
            OpConditions.CondTargetOrderCancelled + OpConditions.CondStopOrderExists, _
            BracketOrderStateCodes.BracketOrderStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, but neither has been cancelled, so we return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderExists + OpConditions.CondTargetOrderExists, _
            OpConditions.CondStopOrderCancelled + OpConditions.CondTargetOrderCancelled, _
            BracketOrderStateCodes.BracketOrderStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, and the stop order has been cancelled but not the target order, so we
' resubmit the stop order and return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderCancelled + OpConditions.CondTargetOrderExists, _
            OpConditions.CondTargetOrderCancelled, _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpActions.ActResubmitStopOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a target order but no stop
' order, and the target order has been cancelled, so we resubmit the target
' order and return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondTargetOrderCancelled, _
            OpConditions.CondStopOrderExists, _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpActions.ActResubmitTargetOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, and the target order has been cancelled but not the stop order, so we
' resubmit the target order and return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderExists + OpConditions.CondTargetOrderCancelled, _
            OpConditions.CondStopOrderCancelled, _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpActions.ActResubmitTargetOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, and both have been cancelled, so we resubmit both the stop order and
' the target order, and return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderCancelled + OpConditions.CondTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateSubmitted, _
            OpActions.ActResubmitStopAndTargetOrders
            
            
'=======================================================================
'                       State:      BracketOrderStateAwaitingOtherOrderCancel
'=======================================================================

' A state timeout has occurred. This means that neither a cancellation nor
' a fill notification has arrived, and we take the view that no such will
' arrive. Closeout the bracket order.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateAwaitingOtherOrderCancel, _
            OpStimuli.StimTimeoutExpired, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpActions.ActPlaceCloseoutOrder

' The application has requested that the bracket order be closed out. Place the
' closeut order and go to closing out state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateAwaitingOtherOrderCancel, _
            OpStimuli.StimCloseout, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpActions.ActCancelOrders
            
' A stop order cancellation has occurred. Enter closing out state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateAwaitingOtherOrderCancel, _
            OpStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            SpecialActions.NoAction

' A target order cancellation has occurred. Enter closing out state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateAwaitingOtherOrderCancel, _
            OpStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            SpecialActions.NoAction

' All orders have completed. We are done, so go to the closed state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateAwaitingOtherOrderCancel, _
            OpStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosed, _
            OpActions.ActCompletionActions


'=======================================================================
'                       State:      BracketOrderStateClosingOut
'=======================================================================

' A state timeout has occurred. This can simply be ignored.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateAwaitingOtherOrderCancel, _
            OpStimuli.StimTimeoutExpired, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpActions.ActPlaceCloseoutOrder

' The entry order has been cancelled, nothing for us to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpStimuli.StimEntryOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the orders and TWS's cancellation
' request arriving at the IB servers or the exchange). There is nothing for
' us to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpStimuli.StimEntryOrderFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut

' The stop order has been cancelled, nothing for us to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut

' The target order has been cancelled, nothing for us to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut

' All orders have completed, and we are left with a non-zero Size. So submit
' a closeout order to reduce the Size to zero. Stay in this state awaiting the
' next 'all orders complete' stimulus.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpStimuli.StimAllOrdersComplete, _
            OpConditions.CondSizeNonZero, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpActions.ActPlaceCloseoutOrder

' All orders have completed, and we are left with a zero Size. We are done,
' so go to the closed state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            OpConditions.CondSizeNonZero, _
            BracketOrderStateCodes.BracketOrderStateClosed, _
            OpActions.ActCompletionActions

' The closeout order has been cancelled (presumably it has been rejected
' by TWS). This is a serious situation since we are left with an unprotected
' position, so raise an alarm.
mTableBuilder.AddStateTableEntry _
            BracketOrderStateCodes.BracketOrderStateClosingOut, _
            OpStimuli.StimCloseoutOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStateCodes.BracketOrderStateClosed, _
            OpActions.ActAlarm, OpActions.ActCompletionActions

mTableBuilder.StateTableComplete
End Sub

