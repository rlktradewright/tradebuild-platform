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
    
    ' This Action cancels the stop-loss order if it exists and its current
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
    
    ' This Action resubmits the stop-loss order (with a new order id). If a
    ' target order exists, then the ocaGroup of the stop-loss order is set to
    ' the ocaGroup of the target order
    ActResubmitStopOrder
    
    ' This Action resubmits the target order (with a new order id). If a
    ' stop-loss order exists, then the ocaGroup of the target order is set to
    ' the ocaGroup of the stop-loss order
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

    ' This Action logs the state transition.
    ActLog

End Enum

Public Enum OpConditions
    ' This condition indicates that the cancellation of the bracket order
    ' has been requested, provided that the entry order has not been filled
    CondNoFillCancellation = &H1&

    ' This condition indicates that a notification that the stop-loss order has been
    ' cancelled has been received.
    CondStopOrderCancelled = &H2&

    ' This condition indicates that a notification that the target order has been
    ' cancelled has been received.
    CondTargetOrderCancelled = &H4&

    ' This condition indicates that the stop-loss order exists.
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
    
    ' This stimulus indicates that the entry order has been cancelled.
    StimEntryOrderCancelled
    
    ' This stimulus indicates that the stop-loss order has been cancelled.
    StimStopOrderCancelled
    
    ' This stimulus indicates that the closeout order has been cancelled.
    ' Note that this is a very unpleasant situation, since it only occurs
    ' when attempting to closeout a position and it leaves us with an
    ' unprotected position.
    StimCloseoutOrderCancelled
    
    ' This stimulus indicates that the target order has been cancelled.
    StimTargetOrderCancelled
    
    ' This stimulus indicates that the entry order has been filled.
    StimEntryOrderFill

    ' This stimulus indicates that a state timeout has expired.
    StimTimeoutExpired
    
    ' This stimulus indicates that a message (potentially an error) has
    ' been notified with regard to an order
    StimOrderError

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

Public Function gBracketOrderStatesToString(ByVal pState As BracketOrderStates) As String
If pState = SpecialStates.StateError Then
    gBracketOrderStatesToString = "$ERROR$"
    Exit Function
End If
Select Case pState
Case BracketOrderStateCreated
    gBracketOrderStatesToString = "Created"
Case BracketOrderStateSubmitted
    gBracketOrderStatesToString = "Submitted"
Case BracketOrderStateCancelling
    gBracketOrderStatesToString = "Cancelling"
Case BracketOrderStateClosingOut
    gBracketOrderStatesToString = "Closing out"
Case BracketOrderStateClosed
    gBracketOrderStatesToString = "Closed"
Case BracketOrderStateAwaitingOtherOrderCancel
    gBracketOrderStatesToString = "Awaiting other order cancel"
Case Else
    AssertArgument False, "Invalid state"
End Select
End Function

Public Function gNextApplicationIndex() As Long
Static sNextApplicationIndex As Long

gNextApplicationIndex = sNextApplicationIndex
sNextApplicationIndex = sNextApplicationIndex + 1
End Function

Public Function gOpActionsToString(ByVal pAction As OpActions) As String
If pAction = SpecialActions.NoAction Then
    gOpActionsToString = "None"
    Exit Function
End If

Select Case pAction
Case ActPlaceOrders
    gOpActionsToString = "Place orders"
Case ActCancelOrders
    gOpActionsToString = "Cancel orders"
Case ActCancelStopOrder
    gOpActionsToString = "Cancel stop-loss order"
Case ActCancelTargetOrder
    gOpActionsToString = "Cancel target order"
Case ActResubmitStopOrder
    gOpActionsToString = "Resubmit stop-loss order"
Case ActResubmitTargetOrder
    gOpActionsToString = "Resubmit target order"
Case ActResubmitStopAndTargetOrders
    gOpActionsToString = "Resubmit stop-loss and target orders"
Case ActPlaceCloseoutOrder
    gOpActionsToString = "Place closeout order"
Case ActAlarm
    gOpActionsToString = "Invoke alarm"
Case ActCompletionActions
    gOpActionsToString = "Do completion actions"
Case ActSetTimeout
    gOpActionsToString = "Set timeout"
Case ActCancelTimeout
    gOpActionsToString = "Cancel timeout"
Case ActLog
    gOpActionsToString = "Log"
Case Else
    AssertArgument False, "Invalid action " & CStr(pAction)
End Select
End Function

Public Function gOpConditionsToString(ByVal pCondition As OpConditions) As String
If pCondition = SpecialConditions.NoConditions Then
    gOpConditionsToString = "None"
    Exit Function
End If

If pCondition = SpecialConditions.AllConditions Then
    gOpConditionsToString = "None"
    Exit Function
End If

Dim s As String
If pCondition And CondNoFillCancellation Then s = "{No-fill cancellation} "
If pCondition And CondStopOrderCancelled Then s = s & "{Stop-loss order cancelled} "
If pCondition And CondTargetOrderCancelled Then s = s & "{Target order cancelled} "
If pCondition And CondStopOrderExists Then s = s & "{Stop-loss order exists} "
If pCondition And CondTargetOrderExists Then s = s & "{Target order exists} "
If pCondition And CondSizeNonZero Then s = s & "{Size non-zero} "
If pCondition And CondProtected Then s = s & "{Protected} "

gOpConditionsToString = s
End Function

Public Function gOpStimuliToString(ByVal pStimulus As OpStimuli) As String
Select Case pStimulus
Case StimExecute
    gOpStimuliToString = "Execute"
Case StimCancelIfNoFill
    gOpStimuliToString = "Cancel if no fill"
Case StimCancelEvenIfFill
    gOpStimuliToString = "Cancel even if filled"
Case StimCloseout
    gOpStimuliToString = "Closeout"
Case StimAllOrdersComplete
    gOpStimuliToString = "All orders complete"
Case StimEntryOrderCancelled
    gOpStimuliToString = "Entry order cancelled"
Case StimStopOrderCancelled
    gOpStimuliToString = "Stop-loss order cancelled"
Case StimCloseoutOrderCancelled
    gOpStimuliToString = "Closeout order cancelled"
Case StimTargetOrderCancelled
    gOpStimuliToString = "Target order cancelled"
Case StimEntryOrderFill
    gOpStimuliToString = "Entry order fill"
Case StimTimeoutExpired
    gOpStimuliToString = "Timeout expired"
Case StimOrderError
    gOpStimuliToString = "Order error"
Case Else
    AssertArgument False, "Invalid stimulus"
End Select
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
            BracketOrderStates.BracketOrderStateCreated, _
            OpStimuli.StimCancelIfNoFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosed, _
            OpActions.ActCancelOrders, OpActions.ActCompletionActions
            
' The application requests that the bracket order be cancelled even if
' fills have occurred. Since the orders have not yet been placed, we
' merely cancel the BracketOrderStatesorders. do any tidying up and set the state to closed.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCreated, _
            OpStimuli.StimCancelEvenIfFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosed, _
            OpActions.ActCancelOrders, OpActions.ActCompletionActions
            
' The application requests that the bracket order be executed, and it is not
' protected. We do that and go to submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCreated, _
            OpStimuli.StimExecute, _
            SpecialConditions.NoConditions, _
            OpConditions.CondProtected, _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpActions.ActPlaceOrders

' The application requests that the bracket order be executed: it is
' protected and there is a stop-loss order. We do that and go to submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCreated, _
            OpStimuli.StimExecute, _
            OpConditions.CondProtected Or OpConditions.CondStopOrderExists, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpActions.ActPlaceOrders

' The application requests that the bracket order be executed: it is
' protected and there is NO stop-loss order. This is a programming error!
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCreated, _
            OpStimuli.StimExecute, _
            OpConditions.CondProtected, _
            OpConditions.CondStopOrderExists, _
            SpecialStates.StateError, _
            SpecialActions.NoAction


'=======================================================================
'                       State:      BracketOrderStateSubmitted
'=======================================================================

' An error has occurred regarding an order. As this is not a protected
' bracket order, we take no action.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimOrderError, _
            SpecialConditions.NoConditions, _
            OpConditions.CondProtected, _
            BracketOrderStates.BracketOrderStateSubmitted

' An error has occurred regarding an order. As this is a protected
' bracket order, we kill all orders.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimOrderError, _
            CondProtected, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpActions.ActCancelOrders

' The entry order has been filled. Nothing to do here.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimEntryOrderFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateSubmitted

' All orders have been completed, so we set the state to closed and do any
' tidying up.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosed, _
            OpActions.ActCompletionActions

' The application requests that the bracket order be cancelled provided no fills
' have occurred. But a fill has already occurred, so we do nothing.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimCancelIfNoFill, _
            OpConditions.CondSizeNonZero, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateSubmitted

' The application requests that the bracket order be cancelled provided no fills
' have occurred. No fills have already occurred, so we cancel all the orders
' and enter the cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimCancelIfNoFill, _
            SpecialConditions.NoConditions, _
            OpConditions.CondSizeNonZero, _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpActions.ActCancelOrders

' The application requests that the bracket order be cancelled even if fills
' have occurred. We cancel all the orders and enter the cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimCancelEvenIfFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpActions.ActCancelOrders

' The entry order has been cancelled (for example it may have been rejected
' by the broker or the user may have cancelled it externally to the
' application, or the bracket order may have been cancelled due to time or
' price constraints being violated).
' There has been no fill, so we cancel the stop and target orders.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimEntryOrderCancelled, _
            SpecialConditions.NoConditions, _
            OpConditions.CondSizeNonZero, _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpActions.ActCancelStopOrder, OpActions.ActCancelTargetOrder

' The entry order has been cancelled after there has been a fill (for example
' the user may have cancelled it externally to the application).
' We cancel the stop and/or target orders. We'll be left with an unprotected
' position, so as this is a protected bracket order, go into closing out
' state to negate the unprotected position.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimEntryOrderCancelled, _
            OpConditions.CondSizeNonZero Or OpConditions.CondProtected, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpActions.ActCancelStopOrder, OpActions.ActCancelTargetOrder

' The entry order has been cancelled after there has been a fill (for example
' the user may have cancelled it externally to the application).
' We cancel the stop and/or target orders. We'll be left with an unprotected
' position, but since this is NOT a protected bracket order, go into
' Cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimEntryOrderCancelled, _
            OpConditions.CondSizeNonZero, _
            OpConditions.CondProtected, _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpActions.ActCancelStopOrder, OpActions.ActCancelTargetOrder

' The stop-loss order has been cancelled, and there is no target
' order. This could be because it has been rejected by the broker, or because
' the user has cancelled it externally to the application. We can't tell which
' of these is the case, so we cancel all orders and, as this is a protected
' bracket order, go into closing out state, because the entry order could be
' filled before being cancelled, and closing out will prevent an unprotected
' position.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimStopOrderCancelled, _
            OpConditions.CondProtected, _
            OpConditions.CondTargetOrderExists, _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpActions.ActCancelOrders

' The stop-loss order has been cancelled, and there is no target
' order. This could be because it has been rejected by the broker, or because
' the user has cancelled it externally to the application. We can't tell which
' of these is the case, so we cancel all orders and, as this is NOT a protected
' bracket order, go into cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            OpConditions.CondTargetOrderExists Or OpConditions.CondProtected, _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpActions.ActCancelOrders

' The stop-loss order has been cancelled, and there IS a target order. This
' could be because it has been rejected by the broker, or because the user has
' cancelled it externally to the application. We can't tell which of these is
' the case, so, as this is a protected bracket order, we enter the 'awaiting
' other order cancel' state and set a timeout.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimStopOrderCancelled, _
            OpConditions.CondTargetOrderExists Or OpConditions.CondProtected, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateAwaitingOtherOrderCancel, _
            OpActions.ActSetTimeout

' The stop-loss order has been cancelled. This could be because it has been
' rejected by the broker, or because the user has cancelled it externally to
' the application. As this is NOT a protected bracket order, we don't care so
' we do nothing and enter the cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            OpConditions.CondProtected, _
            BracketOrderStates.BracketOrderStateCancelling, _
            SpecialActions.NoAction

' The target order has been cancelled, and there IS a stop loss order. This
' could be because it has been rejected by the broker, or because the user has
' cancelled it externally to the application. We can't tell which of these is
' the case, so, as this is a protected bracket order, we enter the 'awaiting
' other order cancel' state and set a timeout.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimTargetOrderCancelled, _
            OpConditions.CondStopOrderExists Or OpConditions.CondProtected, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateAwaitingOtherOrderCancel, _
            OpActions.ActSetTimeout

' The target order has been cancelled.This could be because it has been
' rejected by the broker, or because the user has cancelled it externally to
' the application. As this is NOT a protected bracket order, we don't care so
' we do nothing and enter the cancelling state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            OpConditions.CondProtected, _
            BracketOrderStates.BracketOrderStateCancelling, _
            SpecialActions.NoAction

' The application has requested that the bracket order be closed out. So cancel any
' outstanding orders and go to closing out state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpStimuli.StimCloseout, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpActions.ActCancelOrders
            
'=======================================================================
'                       State:      BracketOrderStateCancelling
'=======================================================================

' The application has requested that the bracket order be cancelled, provided
' there have been no fills. Since it is already being cancelled, there is
' nothing to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimCancelIfNoFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateCancelling

' The application has requested that the bracket order be cancelled, even if
' there have already been fills. Since it is already being cancelled, there
' is nothing to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimCancelEvenIfFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateCancelling

' All orders have now been completed, so do any tidying up and go to the
' closed state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosed, _
            OpActions.ActCompletionActions

' We are notified that the entry order has been cancelled. There has
' been no fill, so we know that no other orders will have been
' active, so we can just go to closed state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderCancelled, _
            SpecialConditions.NoConditions, _
            OpConditions.CondSizeNonZero, _
            BracketOrderStates.BracketOrderStateClosed, _
            OpActions.ActCompletionActions

' We are notified that the entry order has been cancelled. There has
' been a fill, so we need to wait for any other orders to be cancelled.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpConditions.CondSizeNonZero, _
            SpecialConditions.NoConditions, _
            OpConditions.CondSizeNonZero, _
            BracketOrderStates.BracketOrderStateCancelling

' We are notified that the stop-loss order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateCancelling

' We are notified that the target order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateCancelling

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested cancellation and the cancellation request
' being actioned). Since the original cancellation request from the application
' was to cancel even if there have been some fills, we just continue with
' the cancellation by re-requesting cancellation of any outstanding orders.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            SpecialConditions.NoConditions, _
            OpConditions.CondNoFillCancellation, _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpActions.ActCancelOrders

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested cancellation and the cancellation request
' being actioned). The original cancellation request from the application was
' to cancel only if there have been no fills. There now has been a fill.
' There are no stop or target orders, so we just return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation, _
            OpConditions.CondStopOrderExists + OpConditions.CondTargetOrderExists, _
            BracketOrderStates.BracketOrderStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested cancellation and the cancellation request
' being actioned). The original cancellation request from the application was
' to cancel only if there have been no fills. There now has been a fill. There
' is a stop-loss order but no target order, and the stop-loss order has not
' been cancelled, so we just return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderExists, _
            OpConditions.CondStopOrderCancelled + OpConditions.CondTargetOrderExists, _
            BracketOrderStates.BracketOrderStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested cancellation and the cancellation request
' being actioned). The original cancellation request from the application was
' to cancel only if there have been no fills. There now has been a fill. There
' is a stop-loss order but no target order, and the stop-loss order has been
' cancelled, so we resubmit the stop-loss order and return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderCancelled, _
            OpConditions.CondTargetOrderExists, _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpActions.ActResubmitStopOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested cancellation and the cancellation request
' being actioned). The original cancellation request from the application was
' to cancel only if there have been no fills. There now has been a fill. There
' is a target order but no stop-loss order, and the target order has not been
' cancelled, so we return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondTargetOrderExists, _
            OpConditions.CondTargetOrderCancelled + OpConditions.CondStopOrderExists, _
            BracketOrderStates.BracketOrderStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested cancellation and the cancellation request
' being actioned). The original cancellation request from the application was
' to cancel only if there have been no fills. There now has been a fill. There
' is a stop-loss order and a target order, but neither has been cancelled, so
' we return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderExists + OpConditions.CondTargetOrderExists, _
            OpConditions.CondStopOrderCancelled + OpConditions.CondTargetOrderCancelled, _
            BracketOrderStates.BracketOrderStateSubmitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested cancellation and the cancellation request
' being actioned). The original cancellation request from the application was
' to cancel only if there have been no fills. There now has been a fill. There
' is a stop-loss order and a target order, and the stop-loss order has been
' cancelled but not the target order, so we resubmit the stop-loss order and
' return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderCancelled + OpConditions.CondTargetOrderExists, _
            OpConditions.CondTargetOrderCancelled, _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpActions.ActResubmitStopOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested cancellation and the cancellation request
' being actioned). The original cancellation request from the application was
' to cancel only if there have been no fills. There now has been a fill. There
' is a target order but no stop loss order, and the target order has been
' cancelled, so we resubmit the target order and return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondTargetOrderCancelled, _
            OpConditions.CondStopOrderExists, _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpActions.ActResubmitTargetOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested cancellation and the cancellation request
' being actioned). The original cancellation request from the application was
' to cancel only if there have been no fills. There now has been a fill. There
' is a stop-loss order and a target order, and the target order has been
' cancelled but not the stop-loss order, so we resubmit the target order and
' return to the submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderExists + OpConditions.CondTargetOrderCancelled, _
            OpConditions.CondStopOrderCancelled, _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpActions.ActResubmitTargetOrder

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested cancellation and the cancellation request
' being actioned). The original cancellation request from the application was
' to cancel only if there have been no fills. There now has been a fill. There
' is a stop-loss order and a target order, and both have been cancelled, so we
' resubmit both the stop-loss order and the target order, and return to the
' submitted state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateCancelling, _
            OpStimuli.StimEntryOrderFill, _
            OpConditions.CondNoFillCancellation + OpConditions.CondStopOrderCancelled + OpConditions.CondTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateSubmitted, _
            OpActions.ActResubmitStopAndTargetOrders
            
            
'=======================================================================
'                       State:      BracketOrderStateClosed
'=======================================================================

' The bracket order has been completed but the application has requested that
' the bracket order be closed out. So go to closing out state and place the
' closeout order.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateClosed, _
            OpStimuli.StimCloseout, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpActions.ActPlaceCloseoutOrder
            
' The bracket order has been completed but something unexpected happens. Just
' swallow it! An example of this is when an order has been rejected by TWS
' but not removed by the user: a cancellation notification may arrive up
' to several hours later (possibly when the market closes?).
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateClosed, _
            SpecialStimuli.StimulusAll, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosed, _
            OpActions.ActLog
            

'=======================================================================
'                       State:      BracketOrderStateAwaitingOtherOrderCancel
'=======================================================================

' A state timeout has occurred. This means that neither a cancellation nor
' a fill notification has arrived, and we take the view that no such will
' arrive. Closeout the bracket order.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateAwaitingOtherOrderCancel, _
            OpStimuli.StimTimeoutExpired, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpActions.ActPlaceCloseoutOrder

' The application has requested that the bracket order be closed out. Place the
' closeut order and go to closing out state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateAwaitingOtherOrderCancel, _
            OpStimuli.StimCloseout, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpActions.ActPlaceCloseoutOrder
            
' A stop-loss order cancellation has occurred. Enter closing out state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateAwaitingOtherOrderCancel, _
            OpStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut, _
            SpecialActions.NoAction

' A target order cancellation has occurred. Enter closing out state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateAwaitingOtherOrderCancel, _
            OpStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut, _
            SpecialActions.NoAction

' All orders have completed. We are done, so go to the closed state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateAwaitingOtherOrderCancel, _
            OpStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosed, _
            OpActions.ActCompletionActions


'=======================================================================
'                       State:      BracketOrderStateClosingOut
'=======================================================================

' A state timeout has occurred. This can simply be ignored.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpStimuli.StimTimeoutExpired, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpActions.ActPlaceCloseoutOrder

' The entry order has been cancelled, nothing for us to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpStimuli.StimEntryOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested cancelling the orders and the cancellation request
' being actioned). There is nothing for us to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpStimuli.StimEntryOrderFill, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut

' The stop-loss order has been cancelled, nothing for us to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpStimuli.StimStopOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut

' The target order has been cancelled, nothing for us to do.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpStimuli.StimTargetOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut

' All orders have completed, and we are left with a non-zero Size. So submit
' a closeout order to reduce the Size to zero. Stay in this state awaiting the
' next 'all orders complete' stimulus.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpStimuli.StimAllOrdersComplete, _
            OpConditions.CondSizeNonZero, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpActions.ActPlaceCloseoutOrder

' All orders have completed, and we are left with a zero Size. We are done,
' so go to the closed state.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpStimuli.StimAllOrdersComplete, _
            SpecialConditions.NoConditions, _
            OpConditions.CondSizeNonZero, _
            BracketOrderStates.BracketOrderStateClosed, _
            OpActions.ActCompletionActions

' The closeout order has been cancelled. This is a serious situation
' since we are left with an unprotected position, so raise an alarm.
mTableBuilder.AddStateTableEntry _
            BracketOrderStates.BracketOrderStateClosingOut, _
            OpStimuli.StimCloseoutOrderCancelled, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            BracketOrderStates.BracketOrderStateClosed, _
            OpActions.ActAlarm, OpActions.ActCompletionActions

mTableBuilder.StateTableComplete
End Sub

