Attribute VB_Name = "GOrderPlex"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const DummyOffset As Long = &H7FFFFFFF

'@================================================================================
' Enums
'@================================================================================

Public Enum Actions
    
    ' This action places all orders defined in the order plex.
    ActPlaceOrders = 1
    
    ' This action cancels all outstanding orders whose current status
    ' indicates that they are not alReady either filled, cancelled or
    ' cancelling. Note that where an order has not yet been placed, there
    ' may still be work to do, for example logging or notifying listeners.
    ActCancelOrders
    
    ' This action cancels the stop order if it exists and its current
    ' status indicates that it is not alReady either filled, cancelled or
    ' cancelling. Note that where the order has not yet been placed,
    ' there may still be work to do, for example logging or notifying
    ' listeners.
    ActCancelStopOrder
    
    ' This action cancels the target order if it exists and its current
    ' status indicates that it is not alReady either filled, cancelled or
    ' cancelling. Note that where the order has not yet been placed,
    ' there may still be work to do, for example logging or notifying
    ' listeners.
    ActCancelTargetOrder
    
    ' This action resubmits the stop order (with a new order id). If a
    ' target order exists, then the ocaGroup of the stop order is set to
    ' the ocaGroup of the target order
    ActResubmitStopOrder
    
    ' This action resubmits the target order (with a new order id). If a
    ' stop order exists, then the ocaGroup of the target order is set to
    ' the ocaGroup of the stop order
    ActResubmitTargetOrder
    
    ' This action resubmits the both the stop and target orders (with new
    ' order ids and a new ocaGroup).
    ActResubmitStopAndTargetOrders
    
    ' This action creates and places an orders whose effect is to cancel
    ' any existing Size alReady acquired by this order plex. For example,
    ' if the order plex is currently long 1 contract, the closeout order
    ' must sell 1 contract.
    ActPlaceCloseoutOrder
    
    ' This action causes an alarm to be generated (for example, audible
    ' sound, on-screen alert, email, SMS etc).
    ActAlarm
    
    ' This action performs any tidying up needed when an order plex is
    ' completed.
    ActCompletionActions
    
    ' This action causes a timeout stimulus to occur after a short time.
    ActSetTimeout

    ' This action cancels a previously set timeout.
    ActCancelTimeout

End Enum

Public Enum Conditions
    ' This condition indicates that the cancellation of the order plex
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
    
    ' This condition indicates that this order plex is to prevent unprotected
    ' positions as far as possible.
    CondProtected = &H40&

End Enum

Public Enum StateTransitionStimuli
    
    ' This stimulus indicates that the application has requested that
    ' the order plex be executed
    StimExecute = 1
    
    ' This stimulus indicates that the application has requested that
    ' the order plex be cancelled provided the entry order has not alReady
    ' been fully or partially filled. If the entry order is filled during
    ' cancelling, then the stop and target orders (if they exist) must
    ' remain in place
    StimCancelIfNoFill
    
    ' This stimulus indicates that the application has requested that
    ' the order plex be cancelled even if the entry order has alReady been
    ' fully or partially filled. If the entry order is filled during
    ' cancelling, then the stop and target orders (if they exist) must
    ' nevertheless be cancelled.
    StimCancelEvenIfFill
    
    ' This stimulus indicates that the application has requested that
    ' the order plex be closed out, ie that any outstanding orders be
    ' cancelled and that if the order plex then has a non-zero Size, then
    ' a closeout order be submitted to reduce the Size to zero.
    StimCloseout
    
    ' This stimulus indicates that all the orders in the order plex
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
'                       State:      OrderPlexStateCreated
'=======================================================================

' The application requests that the order plex be cancelled provided no
' fills have occurred. Since the orders have not yet been placed, we
' merely cancel the orders. do any tidying up and set the state to closed.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCreated, _
            StateTransitionStimuli.StimCancelIfNoFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            Actions.ActCancelOrders, Actions.ActCompletionActions
            
' The application requests that the order plex be cancelled even if
' fills have occurred. Since the orders have not yet been placed, we
' merely cancel the OrderPlexStateCodesorders. do any tidying up and set the state to closed.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCreated, _
            StateTransitionStimuli.StimCancelEvenIfFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            Actions.ActCancelOrders, Actions.ActCompletionActions
            
' The application requests that the order plex be executed, and it is not
' protected. We do that and go to submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCreated, _
            StateTransitionStimuli.StimExecute, _
            SpecialConditions.NoConditions, _
            Conditions.CondProtected, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            Actions.ActPlaceOrders

' The application requests that the order plex be executed: it is
' protected and there is a stop order. We do that and go to submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCreated, _
            StateTransitionStimuli.StimExecute, _
            Conditions.CondProtected Or Conditions.CondStopOrderExists, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            Actions.ActPlaceOrders

' The application requests that the order plex be executed: it is
' protected and there is NO stop order. This is a programming error!
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCreated, _
            StateTransitionStimuli.StimExecute, _
            Conditions.CondProtected, _
            Conditions.CondStopOrderExists, _
            SpecialStates.StateError, _
            SpecialActions.NoAction


'=======================================================================
'                       State:      OrderPlexStateSubmitted
'=======================================================================

' TWS tells us that the entry order has been filled. Nothing to do here.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimEntryOrderFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateSubmitted

' All orders have been completed, so we set the state to closed and do any
' tidying up.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            Actions.ActCompletionActions

' The application requests that the order plex be cancelled provided no fills
' have occurred. But a fill has alReady occurred, so we do nothing.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimCancelIfNoFill, _
            Conditions.CondSizeNonZero, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateSubmitted

' The application requests that the order plex be cancelled provided no fills
' have occurred. No fills have alReady occurred, so we cancel all the orders
' and enter the cancelling state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimCancelIfNoFill, _
            SpecialConditions.NoConditions, _
            Conditions.CondSizeNonZero, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            Actions.ActCancelOrders

' The application requests that the order plex be cancelled even if fills
' have occurred. We cancel all the orders and enter the cancelling state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimCancelEvenIfFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            Actions.ActCancelOrders

' We are notified that the entry order has been cancelled (for example it
' may have been rejected by TWS or the user may have cancelled it at TWS).
' There has been no fill, so we cancel the stop and target orders (not
' really necessary, since TWS should do this, but just in case...).
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimEntryOrderCancelled, _
            SpecialConditions.NoConditions, _
            Conditions.CondSizeNonZero, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            Actions.ActCancelStopOrder, Actions.ActCancelTargetOrder

' We are notified that the entry order has been cancelled (for example the
' user may have cancelled it at TWS). Note that it can't be the application
' that cancelled it because it has no way of cancelling individual orders.
' There has been a fill. The cancellation will have caused the stop and/or
' target orders to be cancelled as well (though we haven't been notified of
' this yet), but we cancel them anyway just in case. We'll be left with an
' unprotected position, so as this is a protected order plex, go into
' closing out state to negate the unprotected position.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimEntryOrderCancelled, _
            Conditions.CondSizeNonZero Or Conditions.CondProtected, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            Actions.ActCancelStopOrder, Actions.ActCancelTargetOrder

' We are notified that the entry order has been cancelled (for example the
' user may have cancelled it at TWS). Note that it can't be the application
' that cancelled it because it has no way of cancelling individual orders.
' There has been a fill. The cancellation will have caused the stop and/or
' target orders to be cancelled as well (though we haven't been notified of
' this yet), but we cancel them anyway just in case. We'll be left with an
' unprotected position, but since this is NOT a protected order plex
' plex, go into Cancelling state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimEntryOrderCancelled, _
            Conditions.CondSizeNonZero, _
            Conditions.CondProtected, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            Actions.ActCancelStopOrder, Actions.ActCancelTargetOrder

' We are notified that the stop order has been cancelled, and there is no target
' order. This could be because it has been rejected by TWS, or because the user
' has cancelled it at TWS. We can't tell which of these is the case, so we
' cancel all orders and, as this is a protected order plex, go into closing out state,
' because the entry order could be filled before being cancelled, and closing out
' will prevent an unprotected position.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimStopOrderCancelled, _
            Conditions.CondProtected, _
            Conditions.CondTargetOrderExists, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            Actions.ActCancelOrders

' We are notified that the stop order has been cancelled, and there is no target
' order. This could be because it has been rejected by TWS, or because the user
' has cancelled it at TWS. We can't tell which of these is the case, so we
' cancel all orders and, as this is NOT a protected order plex, go into cancelling state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            Conditions.CondTargetOrderExists Or Conditions.CondProtected, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            Actions.ActCancelOrders

' We are notified that the stop order has been cancelled, and there IS a target
' order. This could be because it has been rejected by TWS, or because the user
' has cancelled it at TWS, or because the target order has been filled. We can't
' tell which of these is the case, so, as this is a protected order plex, we enter
' the 'awaiting other order cancel' state and set a timeout.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimStopOrderCancelled, _
            Conditions.CondTargetOrderExists Or Conditions.CondProtected, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateAwaitingOtherOrderCancel, _
            Actions.ActSetTimeout

' We are notified that the stop order has been cancelled, and there IS a target
' order. This could be because it has been rejected by TWS, or because the user
' has cancelled it at TWS, or because the target order has been filled. We can't
' tell which of these is the case, but, as this is NOT a protected order plex, we
' don't care so we cancel the target order and enter the cancelling state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimStopOrderCancelled, _
            Conditions.CondTargetOrderExists, _
            Conditions.CondProtected, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            Actions.ActCancelTargetOrder

' We are notified that the target order has been cancelled, and this is NOT a
' protected order plex. Cancel all orders and go into cancelling state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            Conditions.CondProtected, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            Actions.ActCancelOrders

' We are notified that the target order has been cancelled, and there IS a
' stop order. This could be because it has been rejected by TWS, or because the
' user has cancelled it at TWS, or because the stop order has been filled and
' not yet notified.  We can't tell which of these is the case, so, as this is
' a protected order plex, we enter the 'awaiting other order cancel' state
' and set a timeout.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimTargetOrderCancelled, _
            Conditions.CondStopOrderExists Or Conditions.CondProtected, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateAwaitingOtherOrderCancel, _
            Actions.ActSetTimeout

' The application has requested that the order plex be closed out. So cancel any
' outstanding orders and go to closing out state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            StateTransitionStimuli.StimCloseout, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            Actions.ActCancelOrders
            
'=======================================================================
'                       State:      OrderPlexStateCancelling
'=======================================================================

' The application has requested that the order plex be cancelled, provided
' there have been no fills. Since it is alReady being cancelled, there is
' nothing to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimCancelIfNoFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateCancelling

' The application has requested that the order plex be cancelled, even if
' there have alReady been fills. Since it is alReady being cancelled, there
' is nothing to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimCancelEvenIfFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateCancelling

' All orders have now been completed, so do any tidying up and go to the
' closed state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            Actions.ActCompletionActions

' We are notified that the entry order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimEntryOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateCancelling

' We are notified that the stop order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateCancelling

' We are notified that the target order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateCancelling

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). Since the original
' cancellation request from the application was to cancel even if there have
' been some fills, we just continue with the cancellation by re-requesting
' cancellation of any outstanding orders.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimEntryOrderFill, _
            SpecialConditions.NoConditions, _
            Conditions.CondNoFillCancellation, _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            Actions.ActCancelOrders

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There are no stop or target orders,
' so we just return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimEntryOrderFill, _
            Conditions.CondNoFillCancellation, _
            Conditions.CondStopOrderExists + Conditions.CondTargetOrderExists, _
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
            StateTransitionStimuli.StimEntryOrderFill, _
            Conditions.CondNoFillCancellation + Conditions.CondStopOrderExists, _
            Conditions.CondStopOrderCancelled + Conditions.CondTargetOrderExists, _
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
            StateTransitionStimuli.StimEntryOrderFill, _
            Conditions.CondNoFillCancellation + Conditions.CondStopOrderCancelled, _
            Conditions.CondTargetOrderExists, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            Actions.ActResubmitStopOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a target order but no stop
' order, and the tartget order has not been cancelled, so we return to the
' submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimEntryOrderFill, _
            Conditions.CondNoFillCancellation + Conditions.CondTargetOrderExists, _
            Conditions.CondTargetOrderCancelled + Conditions.CondStopOrderExists, _
            OrderPlexStateCodes.OrderPlexStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, but neither has been cancelled, so we return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimEntryOrderFill, _
            Conditions.CondNoFillCancellation + Conditions.CondStopOrderExists + Conditions.CondTargetOrderExists, _
            Conditions.CondStopOrderCancelled + Conditions.CondTargetOrderCancelled, _
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
            StateTransitionStimuli.StimEntryOrderFill, _
            Conditions.CondNoFillCancellation + Conditions.CondStopOrderCancelled + Conditions.CondTargetOrderExists, _
            Conditions.CondTargetOrderCancelled, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            Actions.ActResubmitStopOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a target order but no stop
' order, and the target order has been cancelled, so we resubmit the target
' order and return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimEntryOrderFill, _
            Conditions.CondNoFillCancellation + Conditions.CondTargetOrderCancelled, _
            Conditions.CondStopOrderExists, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            Actions.ActResubmitTargetOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, and the target order has been cancelled but not the stop order, so we
' resubmit the target order and return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimEntryOrderFill, _
            Conditions.CondNoFillCancellation + Conditions.CondStopOrderExists + Conditions.CondTargetOrderCancelled, _
            Conditions.CondStopOrderCancelled, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            Actions.ActResubmitTargetOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, and both have been cancelled, so we resubmit both the stop order and
' the target order, and return to the submitted state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateCancelling, _
            StateTransitionStimuli.StimEntryOrderFill, _
            Conditions.CondNoFillCancellation + Conditions.CondStopOrderCancelled + Conditions.CondTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateSubmitted, _
            Actions.ActResubmitStopAndTargetOrders
            
            
'=======================================================================
'                       State:      OrderPlexStateAwaitingOtherOrderCancel
'=======================================================================

' A state timeout has occurred. This means that neither a cancellation nor
' a fill notification has arrived, and we take the view that no such will
' arrive. Closeout the order plex.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateAwaitingOtherOrderCancel, _
            StateTransitionStimuli.StimTimeoutExpired, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            Actions.ActPlaceCloseoutOrder

' The application has requested that the order plex be closed out. Place the
' closeut order and go to closing out state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateAwaitingOtherOrderCancel, _
            StateTransitionStimuli.StimCloseout, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            Actions.ActCancelOrders
            
' A stop order cancellation has occurred. Enter closing out state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateAwaitingOtherOrderCancel, _
            StateTransitionStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            SpecialActions.NoAction

' A target order cancellation has occurred. Enter closing out state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateAwaitingOtherOrderCancel, _
            StateTransitionStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            SpecialActions.NoAction

' All orders have completed. We are done, so go to the closed state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateAwaitingOtherOrderCancel, _
            StateTransitionStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            Actions.ActCompletionActions


'=======================================================================
'                       State:      OrderPlexStateClosingOut
'=======================================================================

' A state timeout has occurred. This can simply be ignored.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateAwaitingOtherOrderCancel, _
            StateTransitionStimuli.StimTimeoutExpired, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            Actions.ActPlaceCloseoutOrder

' The entry order has been cancelled, nothing for us to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            StateTransitionStimuli.StimEntryOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the orders and TWS's cancellation
' request arriving at the IB servers or the exchange). There is nothing for
' us to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            StateTransitionStimuli.StimEntryOrderFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut

' The stop order has been cancelled, nothing for us to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            StateTransitionStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut

' The target order has been cancelled, nothing for us to do.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            StateTransitionStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut

' All orders have completed, and we are left with a non-zero Size. So submit
' a closeout order to reduce the Size to zero. Stay in this state awaiting the
' next 'all orders complete' stimulus.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            StateTransitionStimuli.StimAllOrdersComplete, _
            Conditions.CondSizeNonZero, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            Actions.ActPlaceCloseoutOrder

' All orders have completed, and we are left with a zero Size. We are done,
' so go to the closed state.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            StateTransitionStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            Conditions.CondSizeNonZero, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            Actions.ActCompletionActions

' The closeout order has been cancelled (presumably it has been rejected
' by TWS). This is a serious situation since we are left with an unprotected
' position, so raise an alarm.
mTableBuilder.addStateTableEntry _
            OrderPlexStateCodes.OrderPlexStateClosingOut, _
            StateTransitionStimuli.StimCloseoutOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            OrderPlexStateCodes.OrderPlexStateClosed, _
            Actions.ActAlarm, Actions.ActCompletionActions

mTableBuilder.stateTableComplete
End Sub

