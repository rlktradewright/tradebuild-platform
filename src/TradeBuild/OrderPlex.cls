VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPlex"
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
' Interfaces
'================================================================================

Implements QuoteListener

'================================================================================
' Events
'================================================================================

Event ChangesCancelled()
Event Clean()
Event Dirty()
Event Completed()
Event SelfCancelled()
Event EntryOrderFilled()
Event EntryOrderPartiallyFilled()
Event StopOrderFilled()
Event StopOrderPartiallyFilled()
Event TargetOrderFilled()
Event TargetOrderPartiallyFilled()
Event CloseoutOrderPartiallyFilled()
Event StopAdjusted()
Event TargetAdjusted()
Event ProfitThresholdExceeded()
Event LossThresholdExceeded()
Event DrawdownThresholdExceeded()

'================================================================================
' Constants
'================================================================================

' This condition indicates that the cancellation of the order plex
' has been requested, provided that the entry order has not been filled
Private Const COND_NO_FILL_CANCELLATION As Long = &H1

' This condition indicates that a notification that the stop order has been
' cancelled has been received via the API. This can be in the form of either
' an orderStatus message with status 'cancelled', or an errorMessage with
' errorCode = 202, or an errorMessage with errorCode = 201 (indicating
' that the order has been rejected for some reason).
Private Const COND_STOP_ORDER_CANCELLED As Long = &H2

' This condition indicates that a notification that the target order has been
' cancelled has been received via the API. This can be in the form of either
' an orderStatus message with status 'cancelled', or an errorMessage with
' errorCode = 202, or an errorMessage with errorCode = 201 (indicating
' that the order has been rejected for some reason).
Private Const COND_TARGET_ORDER_CANCELLED As Long = &H4

' This condition indicates that the stop order exists.
Private Const COND_STOP_ORDER_EXISTS As Long = &H8

' This condition indicates that the target order exists.
Private Const COND_TARGET_ORDER_EXISTS As Long = &H10

' This condition indicates that the entry order has been partially or
' completely filled.
Private Const COND_SIZE_NON_ZERO As Long = &H20

Private Const DummyOffset As Long = &H7FFFFFFF

'================================================================================
' Enums
'================================================================================

Private Enum ErrCodes
    errNotImplemented = vbObjectError + 513
    errCloseoutIncomplete
    errPriceChangeNotAllowedForOrderType
    errTriggerPriceChangeNotAllowedForOrderType
    errOffsetChangeNotAllowedForOrderType
    errCantChangeFilledOrder
    errNoStopOrder
    errNoTargetOrder
    errCantChangeUnplacedOrder
    errCantChangeSubmittedOrder
    errInvalidActionReturnedFromStateEngine
    errActionNotAllowedInThisState
    errWidenStopNotAllowed
    errCantChangeStopForOrderType
    errCantChangeTargetForOrderType
End Enum

Public Enum OrderPlexStateCodes
    
    ' This state indicates that the order plex has been created but none of
    ' the orders have yet been placed.
    Created = 1
    
    ' This state indicates that all the orders in the order plex have been placed.
    Submitted
    
    ' This state indicates that the order plex is in the process of being
    ' cancelled.
    Cancelling
    
    ' This state indicates that the order plex is being closed out.
    ' (When an order plex has been closed out, there are no outstanding
    ' orders and no net size - note that this may require a closeout
    ' order to be placed to negate any existing size.)
    ClosingOut
    
    ' This state indicates that all orders have been completed.
    Closed
End Enum

Private Enum StateTransitionStimuli
    
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

Private Enum Actions
    
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

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mPositionManager As positionManager
Private mTicker As ticker

Private mTickSize As Double
Private mTickValue As Currency
Private mcurrencyCode As String

Private mStateEngine As cStateEngine
Private mOrderContext As OrderContext
Private mKey As String
Private mHandle As Long
Private WithEvents mTimerEntry As TimerListEntry
Attribute mTimerEntry.VB_VarHelpID = -1

Private mDescription As String
Private mEntryReason As String
Private mTargetReason As String
Private mStopReason As String

Private mSize As Long
Private mLastListenerKey As String
Private mCancelPrice As Double

Private WithEvents mEntryOrder As order
Attribute mEntryOrder.VB_VarHelpID = -1
Private mEntryOrderFinished As Boolean
Private WithEvents mStopOrder As order
Attribute mStopOrder.VB_VarHelpID = -1
Private mStopOrderFinished As Boolean
Private WithEvents mTargetOrder As order
Attribute mTargetOrder.VB_VarHelpID = -1
Private mTargetOrderFinished As Boolean
Private WithEvents mCloseoutOrder As order
Attribute mCloseoutOrder.VB_VarHelpID = -1

Private mStopStrategy As StopManagementStrategy
Private mScalingStrategy As PositionManagementStrategy

Private mNewQuantity As Long
Private mNewEntryPrice As Double
Private mNewEntryTriggerPrice As Double
Private mNewStopPrice As Double
Private mNewStopOffset As Long
Private mNewStopTriggerPrice As Double
Private mNewTargetPrice As Double
Private mNewTargetOffset As Long
Private mNewTargetTriggerPrice As Double

Private mDummy As Boolean
    ' if set, this is a dummy orderplex, which means that it will be handled by
    ' the simulator like any other order, but its size will be considered 0. It
    ' will not be submitted as a real order

Private mValueFactor As Double
Private mProfit As Currency
Private mMaxProfit As Currency
Private mDrawdown As Currency
Private mMaxLoss As Currency
Private mProfitNotificationThreshold As Currency
Private mLossNotificationThreshold As Currency
Private mDrawdownNotificationThreshold As Currency

Private mChangeListeners As Collection

Private mCreationTime As Date

Private mProfitListeners As Collection

Private mIndexApplication As Long

Private mDirty As Boolean

Private mOrders As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mHandle = -1
mNewQuantity = -1
mNewEntryPrice = -1
mNewEntryTriggerPrice = -1
mNewStopPrice = -1
mNewStopOffset = DummyOffset
mNewStopTriggerPrice = -1
mNewTargetPrice = -1
mNewTargetOffset = DummyOffset
mNewTargetTriggerPrice = -1

mIndexApplication = OrderPlexGlobals.gNextApplicationIndex

Set mOrders = New Collection

Set mChangeListeners = New Collection

Set mProfitListeners = New Collection

Set mStateEngine = New cStateEngine
mStateEngine.initialState = OrderPlexStateCodes.Created

'set up the state table entries

'=======================================================================
'                       State:      OrderPlexStateCodes.Created
'=======================================================================

' The application requests that the order plex be cancelled provided no
' fills have occurred. Since the orders have not yet been placed, we
' merely cancel the orders. do any tidying up and set the state to closed.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Created, _
            STIM_CANCEL_IF_NO_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Closed, _
            ACT_CANCEL_ORDERS, ACT_COMPLETION_ACTIONS
            
' The application requests that the order plex be cancelled even if
' fills have occurred. Since the orders have not yet been placed, we
' merely cancel the orders. do any tidying up and set the state to closed.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Created, _
            STIM_CANCEL_EVEN_IF_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Closed, _
            ACT_CANCEL_ORDERS, ACT_COMPLETION_ACTIONS
            
' The application requests that the order plex be executed. We do that and
' go to submitted state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Created, _
            STIM_EXECUTE, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Submitted, _
            ACT_PLACE_ORDERS

'=======================================================================
'                       State:      OrderPlexStateCodes.Submitted
'=======================================================================

' TWS tells us that the entry order has been filled. Nothing to do here.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Submitted, _
            STIM_ENTRY_ORDER_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Submitted

' All orders have been completed, so we set the state to closed and do any
' tidying up.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Submitted, _
            STIM_ALL_ORDERS_COMPLETE, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Closed, _
            ACT_COMPLETION_ACTIONS

' The application requests that the order plex be cancelled provided no fill
' have occurred. But a fill has already occurred, so we do nothing.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Submitted, _
            STIM_CANCEL_IF_NO_FILL, _
            COND_SIZE_NON_ZERO, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Submitted

' The application requests that the order plex be cancelled provided no fill
' have occurred. No fills have already occurred, so we cancel all the orders
' and enter the cancelling state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Submitted, _
            STIM_CANCEL_IF_NO_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            COND_SIZE_NON_ZERO, _
            OrderPlexStateCodes.Cancelling, _
            ACT_CANCEL_ORDERS

' The application requests that the order plex be cancelled even if fills
' have occurred. We cancel all the orders and enter the cancelling state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Submitted, _
            STIM_CANCEL_EVEN_IF_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Cancelling, _
            ACT_CANCEL_ORDERS

' We are notified that the entry order has been cancelled (for example it
' may have been rejected by TWS or the user may have cancelled it at TWS).
' There has been no fill, so we cancel the stop and target orders (not
' really necessary, since TWS should do this, but just in case...).
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Submitted, _
            STIM_ENTRY_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            COND_SIZE_NON_ZERO, _
            OrderPlexStateCodes.Cancelling, _
            ACT_CANCEL_STOP_ORDER, ACT_CANCEL_TARGET_ORDER

' We are notified that the entry order has been cancelled (for example the
' user may have cancelled it at TWS). Note that it can't be the application
' that cancelled it because it has no way of cancelling individual orders.
' The cancellation will have caused the stop and/or target orders to be
' cancelled as well (though we haven't been notified of this yet). Therefore
' we'll be left with an unprotected position, so we cancel the stop and target
' orders (just in case) and go into closing out state to negate the
' unprotected position.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Submitted, _
            STIM_ENTRY_ORDER_CANCELLED, _
            COND_SIZE_NON_ZERO, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.ClosingOut, _
            ACT_CANCEL_STOP_ORDER, ACT_CANCEL_TARGET_ORDER

' We are notified that the stop order has been cancelled. This could be because
' it has been rejected by TWS, or because the user has cancelled it at TWS. We
' can't tell which of these is the case, so we cancel all orders and go into
' closing out state, because the entry order could be filled before being
' cancelled, and closing out will prevent an unprotected position.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Submitted, _
            STIM_STOP_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.ClosingOut, _
            ACT_CANCEL_ORDERS

' We are notified that the target order has been cancelled. This could be because
' it has been rejected by TWS, or because the user has cancelled it at TWS. We
' can't tell which of these is the case, so we cancel all orders and go into
' closing out state, because the entry order could be filled before being
' cancelled, and closing out will prevent an unprotected position.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Submitted, _
            STIM_TARGET_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.ClosingOut, _
            ACT_CANCEL_ORDERS

' The application has requested that the order plex be closed out. So cancel any
' outstanding orders and go to closing out state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Submitted, _
            STIM_CLOSEOUT, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.ClosingOut, _
            ACT_CANCEL_ORDERS
            
'=======================================================================
'                       State:      OrderPlexStateCodes.Cancelling
'=======================================================================

' The application has requested that the order plex be cancelled, provided
' there have been no fills. Since it is already being cancelled, there is
' nothing to do.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_CANCEL_IF_NO_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Cancelling

' The application has requested that the order plex be cancelled, even if
' there have already been fills. Since it is already being cancelled, there
' is nothing to do.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_CANCEL_EVEN_IF_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Cancelling

' All orders have now been completed, so do any tidying up and go to the
' closed state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ALL_ORDERS_COMPLETE, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Closed, _
            ACT_COMPLETION_ACTIONS

' We are notified that the entry order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ENTRY_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Cancelling

' We are notified that the stop order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_STOP_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Cancelling

' We are notified that the target order has been cancelled. Now we just need
' to wait for any other orders to be cancelled.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_TARGET_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Cancelling

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). Since the original
' cancellation request from the application was to cancel even if there have
' been some fills, we just continue with the cancellation by re-requesting
' cancellation of any outstanding orders.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ENTRY_ORDER_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            COND_NO_FILL_CANCELLATION, _
            OrderPlexStateCodes.Cancelling, _
            ACT_CANCEL_ORDERS

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There are no stop or target orders,
' so we just return to the submitted state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION, _
            COND_STOP_ORDER_EXISTS + COND_TARGET_ORDER_EXISTS, _
            OrderPlexStateCodes.Submitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order but no target
' order, and the stop order has not been cancelled, so we just return to the
' submitted state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_EXISTS, _
            COND_STOP_ORDER_CANCELLED + COND_TARGET_ORDER_EXISTS, _
            OrderPlexStateCodes.Submitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order but no target
' order, and the stop order has been cancelled, so we resubmit the stop order
' and return to the submitted state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_CANCELLED, _
            COND_TARGET_ORDER_EXISTS, _
            OrderPlexStateCodes.Submitted, _
            ACT_RESUBMIT_STOP_ORDER

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a target order but no stop
' order, and the tartget order has not been cancelled, so we return to the
' submitted state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_TARGET_ORDER_EXISTS, _
            COND_TARGET_ORDER_CANCELLED + COND_STOP_ORDER_EXISTS, _
            OrderPlexStateCodes.Submitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, but neither has been cancelled, so we return to the submitted state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_EXISTS + COND_TARGET_ORDER_EXISTS, _
            COND_STOP_ORDER_CANCELLED + COND_TARGET_ORDER_CANCELLED, _
            OrderPlexStateCodes.Submitted

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, and the stop order has been cancelled but not the target order, so we
' resubmit the stop order and return to the submitted state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_CANCELLED + COND_TARGET_ORDER_EXISTS, _
            COND_TARGET_ORDER_CANCELLED, _
            OrderPlexStateCodes.Submitted, _
            ACT_RESUBMIT_STOP_ORDER

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a target order but no stop
' order, and the target order has been cancelled, so we resubmit the target
' order and return to the submitted state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_TARGET_ORDER_CANCELLED, _
            COND_STOP_ORDER_EXISTS, _
            OrderPlexStateCodes.Submitted, _
            ACT_RESUBMIT_TARGET_ORDER

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, and the target order has been cancelled but not the stop order, so we
' resubmit the target order and return to the submitted state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_EXISTS + COND_TARGET_ORDER_CANCELLED, _
            COND_STOP_ORDER_CANCELLED, _
            OrderPlexStateCodes.Submitted, _
            ACT_RESUBMIT_TARGET_ORDER

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the order and TWS's cancellation
' request arriving at the IB servers or the exchange). The original
' cancellation request from the application was to cancel only if there have
' been no fills. There now has been a fill. There is a stop order and a target
' order, and both have been cancelled, so we resubmit both the stop order and
' the target order, and return to the submitted state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.Cancelling, _
            STIM_ENTRY_ORDER_FILL, _
            COND_NO_FILL_CANCELLATION + COND_STOP_ORDER_CANCELLED + COND_TARGET_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Submitted, _
            ACT_RESUBMIT_STOP_AND_TARGET_ORDERS
            
            
'=======================================================================
'                       State:      OrderPlexStateCodes.ClosingOut
'=======================================================================

' The entry order has been cancelled, nothing for us to do.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.ClosingOut, _
            STIM_ENTRY_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.ClosingOut

' The entry order has been unexpectedly filled (this occurred between the
' time that we requested TWS to cancel the orders and TWS's cancellation
' request arriving at the IB servers or the exchange). There is nothing for
' us to do.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.ClosingOut, _
            STIM_ENTRY_ORDER_FILL, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.ClosingOut

' The stop order has been cancelled, nothing for us to do.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.ClosingOut, _
            STIM_STOP_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.ClosingOut

' The target order has been cancelled, nothing for us to do.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.ClosingOut, _
            STIM_TARGET_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.ClosingOut

' All orders have completed, and we are left with a non-zero size. So submit
' a closeout order to reduce the size to zero. Stay in this state awaiting the
' next 'all orders complete' stimulus.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.ClosingOut, _
            STIM_ALL_ORDERS_COMPLETE, _
            COND_SIZE_NON_ZERO, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.ClosingOut, _
            ACT_PLACE_CLOSEOUT_ORDER

' All orders have completed, and we are left with a zero size. We are done,
' so go to the closed state.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.ClosingOut, _
            STIM_ALL_ORDERS_COMPLETE, _
            SpecialConditions.NO_CONDITIONS, _
            COND_SIZE_NON_ZERO, _
            OrderPlexStateCodes.Closed, _
            ACT_COMPLETION_ACTIONS

' The closeout order has been cancelled (presumably it has been rejected
' by TWS). This is a serious situation since we are left with an unprotected
' position, so raise an alarm.
mStateEngine.addStateTableEntry _
            OrderPlexStateCodes.ClosingOut, _
            STIM_CLOSEOUT_ORDER_CANCELLED, _
            SpecialConditions.NO_CONDITIONS, _
            SpecialConditions.NO_CONDITIONS, _
            OrderPlexStateCodes.Closed, _
            ACT_ALARM, ACT_COMPLETION_ACTIONS

End Sub

Private Sub Class_Terminate()
Debug.Print "OrderPlex terminated"
End Sub

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As QuoteEvent)
If Not longPosition Then calcProfitLoss ev.price
End Sub

Private Sub QuoteListener_bid(ev As QuoteEvent)
If longPosition Then calcProfitLoss ev.price
End Sub

Private Sub QuoteListener_High(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_PreviousClose(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_trade(ev As QuoteEvent)
Dim cancelPriceHit As Boolean

If mCancelPrice = 0 Then Exit Sub

If longPosition Then
    Select Case mEntryOrder.orderType
    Case TradeBuild.OrderTypes.OrderTypeLimit
        If ev.price >= mCancelPrice Then cancelPriceHit = True
    Case TradeBuild.OrderTypes.OrderTypeStop, _
        TradeBuild.OrderTypes.OrderTypeStopLimit
        If ev.price <= mCancelPrice Then cancelPriceHit = True
    End Select
Else
    Select Case mEntryOrder.orderType
    Case TradeBuild.OrderTypes.OrderTypeLimit
        If ev.price <= mCancelPrice Then cancelPriceHit = True
    Case TradeBuild.OrderTypes.OrderTypeStop, _
        TradeBuild.OrderTypes.OrderTypeStopLimit
        If ev.price >= mCancelPrice Then cancelPriceHit = True
    End Select
End If
If cancelPriceHit Then
    mTicker.notifyListeners _
                        "Cancel price hit (" & mTicker.TradePrice & ")", _
                        IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
    Cancel
    fireChange SelfCancelled
    RaiseEvent SelfCancelled
End If
End Sub

Private Sub QuoteListener_volume(ev As QuoteEvent)

End Sub

'================================================================================
' mCloseoutOrder Event Handlers
'================================================================================

Private Sub mCloseoutOrder_StatusChanged()
fireChange CloseoutOrderChanged
Select Case mCloseoutOrder.status
    Case TradeBuild.OrderStatuses.OrderStatusFilled
        mTicker.notifyListeners _
                            "Closeout order completed (" & _
                            mCloseoutOrder.id & "): " & _
                            orderActionToString(mCloseoutOrder.action) & " " & mCloseoutOrder.quantityFilled & " at " & _
                            mCloseoutOrder.averagePrice, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
        checkAllOrdersComplete
        
        ' I think there is a possibility that when the first closeout
        ' order is filled, it might be necessary to submit another closeout
        ' order because there may have been another fill on the entry, stop
        ' or target orders just after submitting the cancel. The following
        ' test checks for this
        ' NB: I no longer think the above is true - the closeout order is now
        ' not submitted until the entry, stop and target orders are all either
        ' filled or cancelled
        If size <> 0 Then err.Raise errCloseoutIncomplete, "TradeBuild.OrderPlex::mCloseoutOrder_StatusChanged", "Closeout incomplete: size remaining is " & size
        
'        Raiseevent Closedout
    Case TradeBuild.OrderStatuses.OrderStatusPendingSubmit
    Case TradeBuild.OrderStatuses.OrderStatusCancelling
    Case TradeBuild.OrderStatuses.OrderStatusPreSubmitted
    Case TradeBuild.OrderStatuses.OrderStatusSubmitted
    Case TradeBuild.OrderStatuses.OrderStatusCancelled
        mTicker.notifyListeners _
                            "Closeout order cancelled (" & _
                            mCloseoutOrder.id & ")", _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
End Select
End Sub

Private Sub mCloseoutOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

calcProfitLoss fillPrice

fireChange CloseoutOrderChanged
thisSize = IIf(mCloseoutOrder.action = ActionBuy, quantityFilled, -quantityFilled)
setSize size + thisSize
mTicker.notifyListeners _
                    "Closeout order fill (" & _
                    mCloseoutOrder.id & "): " & _
                    orderActionToString(mCloseoutOrder.action) & " " & quantityFilled & " at " & _
                    fillPrice, _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)

mPositionManager.notifyExecution thisSize, fillPrice, mHandle, True

RaiseEvent CloseoutOrderPartiallyFilled
End Sub

'================================================================================
' mEntryOrder Event Handlers
'================================================================================

Private Sub mEntryOrder_StatusChanged()
fireChange EntryOrderChanged
Select Case mEntryOrder.status
    Case TradeBuild.OrderStatuses.OrderStatusFilled
        mCancelPrice = 0
        
        mTicker.notifyListeners _
                            "Entry order completed (" & _
                            mEntryOrder.id & "): " & _
                            orderActionToString(mEntryOrder.action) & " " & mEntryOrder.quantityFilled & " at " & _
                            mEntryOrder.averagePrice, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
        
        RaiseEvent EntryOrderFilled
        
        mEntryOrderFinished = True
        checkAllOrdersComplete
    
    Case TradeBuild.OrderStatuses.OrderStatusPendingSubmit
    Case TradeBuild.OrderStatuses.OrderStatusCancelling
    Case TradeBuild.OrderStatuses.OrderStatusPreSubmitted
    Case TradeBuild.OrderStatuses.OrderStatusSubmitted
        
'        If mEntryOrder.orderType = TradeBuild.OrderTypes.OrderTypeStopLimit Then
'            ' once a stop limit order becomes submitted, it changes to
'            ' a limit order
'            mEntryOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
'            mEntryOrder.triggerPrice = 0
'        End If
        
        If mEntryOrder.quantityFilled > 0 Then
            placePeggedOrders
        
            If Not mStopOrder Is Nothing Then
                If mEntryOrder.quantityFilled <> mStopOrder.Quantity And _
                    mStopOrder.parentId = 0 Then
                    mStopOrder.Quantity = mEntryOrder.quantityFilled
                    mStopOrder.transmit = True
                    gTradeBuildAPI.placeOrder mStopOrder
                End If
            End If
            If Not mTargetOrder Is Nothing Then
                If mEntryOrder.quantityFilled <> mTargetOrder.Quantity And _
                    mTargetOrder.parentId = 0 Then
                    mTargetOrder.Quantity = mEntryOrder.quantityFilled
                    mTargetOrder.transmit = True
                    gTradeBuildAPI.placeOrder mTargetOrder
                End If
            End If
        End If
    Case TradeBuild.OrderStatuses.OrderStatusCancelled
        mEntryOrderFinished = True
        mTicker.notifyListeners _
                            "Entry order cancelled (" & _
                            mEntryOrder.id & ") ", _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
        handleStimulus STIM_ENTRY_ORDER_CANCELLED
        checkAllOrdersComplete
End Select
End Sub

Private Sub mEntryOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long
Dim risk As Currency

fireChange EntryOrderChanged

cancelAfter = 0

mCancelPrice = 0

thisSize = IIf(longPosition, quantityFilled, -quantityFilled)
setSize size + thisSize


mTicker.notifyListeners _
                    "Entry order fill (" & _
                    mEntryOrder.id & "): " & _
                    orderActionToString(mEntryOrder.action) & " " & quantityFilled & " at " & _
                    fillPrice, _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)

If mHandle < 0 And Not mStopOrder Is Nothing Then
    If longPosition Then
        If mStopOrder.orderType = TradeBuild.OrderTypes.OrderTypeAutoStop Then
            risk = (mEntryOrder.Quantity + mEntryOrder.quantityFilled) * mStopOrder.offset * mTickValue
        Else
            risk = (mEntryOrder.Quantity + mEntryOrder.quantityFilled) * (fillPrice - mStopOrder.triggerPrice) * mValueFactor
        End If
    Else
        If mStopOrder.orderType = TradeBuild.OrderTypes.OrderTypeAutoStop Then
            risk = -(mEntryOrder.Quantity + mEntryOrder.quantityFilled) * mStopOrder.offset * mTickValue
        Else
            risk = (mEntryOrder.Quantity + mEntryOrder.quantityFilled) * (mStopOrder.triggerPrice - fillPrice) * mValueFactor
        End If
    End If
End If

If mHandle < 0 Then
    mHandle = mPositionManager.registerPlex(mKey, mDescription, mEntryOrder.action, risk)
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonEntry, mEntryReason
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonTarget, mTargetReason
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonStop, mStopReason
End If
mPositionManager.notifyExecution thisSize, fillPrice, mHandle, False

placePeggedOrders
        
If Not mStopOrder Is Nothing Then
    If mStopOrder.Quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        adjustStop mStopOrder.triggerPrice, Abs(size)
    End If
End If

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        adjustTarget mTargetOrder.limitPrice, Abs(size)
    End If
End If

RaiseEvent EntryOrderPartiallyFilled

handleStimulus STIM_ENTRY_ORDER_FILL
End Sub

'================================================================================
' mStopOrder Event Handlers
'================================================================================

Private Sub mStopOrder_StatusChanged()
fireChange StopOrderChanged
Select Case mStopOrder.status
    Case TradeBuild.OrderStatuses.OrderStatusFilled
        mTicker.notifyListeners _
                            "Stop order completed (" & _
                            mStopOrder.id & "): " & _
                            orderActionToString(mStopOrder.action) & " " & mStopOrder.quantityFilled & " at " & _
                            mStopOrder.averagePrice, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
        
        If Not mTargetOrder Is Nothing Then
            If mStopOrder.parentId = 0 Or mTargetOrder.parentId = 0 Then
                ' this happens when either the target order
                ' or the stop order is AUTO
                gTradeBuildAPI.cancelOrder mTargetOrder.id
            End If
        End If
        
        RaiseEvent StopOrderFilled
        
        mStopOrderFinished = True
        checkAllOrdersComplete
    
    Case TradeBuild.OrderStatuses.OrderStatusPendingSubmit
    Case TradeBuild.OrderStatuses.OrderStatusCancelling
    Case TradeBuild.OrderStatuses.OrderStatusPreSubmitted
    Case TradeBuild.OrderStatuses.OrderStatusSubmitted
        If Not mTargetOrder Is Nothing Then
            If mStopOrder.Quantity <> mTargetOrder.Quantity And _
                mTargetOrder.parentId = 0 _
            Then
                mTargetOrder.Quantity = mStopOrder.Quantity
                mStopOrder.transmit = True
                gTradeBuildAPI.placeOrder mTargetOrder
            End If
        End If
    Case TradeBuild.OrderStatuses.OrderStatusCancelled
        mStopOrderFinished = True
        mTicker.notifyListeners _
                            "Stop order cancelled (" & _
                            mStopOrder.id & ")", _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
        mStateEngine.setCondition COND_STOP_ORDER_CANCELLED
        handleStimulus STIM_STOP_ORDER_CANCELLED
        checkAllOrdersComplete
End Select
End Sub

Private Sub mStopOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

calcProfitLoss fillPrice

fireChange StopOrderChanged

thisSize = IIf(mStopOrder.action = ActionBuy, quantityFilled, -quantityFilled)
setSize size + thisSize
mTicker.notifyListeners _
                        "Stop order fill (" & _
                        mStopOrder.id & "): " & _
                        orderActionToString(mStopOrder.action) & " " & quantityFilled & " at " & _
                        fillPrice, _
                        IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)

mPositionManager.notifyExecution thisSize, fillPrice, mHandle, False

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        adjustTarget mTargetOrder.limitPrice, Abs(size)
    End If
End If

RaiseEvent StopOrderPartiallyFilled
End Sub

'================================================================================
' mTargetOrder Event Handlers
'================================================================================

Private Sub mTargetOrder_StatusChanged()
fireChange TargetOrderChanged
Select Case mTargetOrder.status
    Case TradeBuild.OrderStatuses.OrderStatusFilled
        mTicker.notifyListeners _
                            "Target order completed (" & _
                            mTargetOrder.id & "): " & _
                            orderActionToString(mTargetOrder.action) & " " & mTargetOrder.quantityFilled & " at " & _
                            mTargetOrder.averagePrice, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
        
        If Not mStopOrder Is Nothing Then
            If mTargetOrder.parentId = 0 Or mStopOrder.parentId = 0 Then
                ' this happens when either the target order or
                ' the stop order is AUTO
                gTradeBuildAPI.cancelOrder mStopOrder.id
            End If
        End If
        
        RaiseEvent TargetOrderFilled
        
        mTargetOrderFinished = True
        checkAllOrdersComplete
    
    Case TradeBuild.OrderStatuses.OrderStatusPendingSubmit
    Case TradeBuild.OrderStatuses.OrderStatusCancelling
    Case TradeBuild.OrderStatuses.OrderStatusPendingSubmit
    Case TradeBuild.OrderStatuses.OrderStatusSubmitted
        If Not mStopOrder Is Nothing Then
            If mStopOrder.Quantity <> mTargetOrder.Quantity And _
                mStopOrder.parentId = 0 Then
                mStopOrder.Quantity = mTargetOrder.Quantity
                mStopOrder.transmit = True
                gTradeBuildAPI.placeOrder mStopOrder
            End If
        End If
    Case TradeBuild.OrderStatuses.OrderStatusCancelled
        mTargetOrderFinished = True
        mTicker.notifyListeners _
                            "Target order cancelled (" & _
                            mTargetOrder.id & ")", _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
        mStateEngine.setCondition COND_TARGET_ORDER_CANCELLED
        handleStimulus STIM_TARGET_ORDER_CANCELLED
        checkAllOrdersComplete
    
End Select
End Sub

Private Sub mTargetOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

calcProfitLoss fillPrice

fireChange TargetOrderChanged
thisSize = IIf(mTargetOrder.action = ActionBuy, quantityFilled, -quantityFilled)
setSize size + thisSize
mTicker.notifyListeners _
                        "Target order fill (" & _
                        mTargetOrder.id & "): " & _
                        orderActionToString(mTargetOrder.action) & " " & quantityFilled & " at " & _
                        fillPrice, _
                        IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)

mPositionManager.notifyExecution thisSize, fillPrice, mHandle, False

If Not mStopOrder Is Nothing Then
    If mStopOrder.Quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        adjustStop mStopOrder.triggerPrice, Abs(size)
    End If
End If

RaiseEvent TargetOrderPartiallyFilled
End Sub

'================================================================================
' mTimerEntry Event Handlers
'================================================================================

Private Sub mTimerEntry_Expired()
mTicker.notifyListeners _
                    "Cancel time expired (" & mTicker.TradePrice & ")", _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
Cancel
fireChange SelfCancelled
RaiseEvent SelfCancelled
cancelAfter = 0
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let cancelAfter(ByVal value As Long)
If value <> 0 Then
    Set mTimerEntry = mTicker.addTimer(value, expiryTimeUnits.RelativeSeconds, Nothing)
ElseIf Not mTimerEntry Is Nothing Then
    mTicker.removeTimer mTimerEntry
    Set mTimerEntry = Nothing
End If
End Property

Public Property Get closedOut() As Boolean
closedOut = (mStateEngine.state = OrderPlexStateCodes.Closed)
End Property

Public Property Get closeoutOrder() As order
Set closeoutOrder = mCloseoutOrder
End Property

Public Property Get contract() As contract
Set contract = mTicker.contract
End Property

Public Property Get creationTime() As Date
creationTime = mCreationTime
End Property

Public Property Let Description(ByVal value As String)
mDescription = value
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Get Dirty() As Boolean
Dirty = mDirty
End Property

Public Property Let drawdownNotificationThreshold(ByVal value As Currency)
mDrawdownNotificationThreshold = value
End Property

Friend Property Let dummy(ByVal value As Boolean)
mDummy = value
End Property

Public Property Get dummy() As Boolean
dummy = mDummy
End Property

Public Property Get entryOrder() As order
Set entryOrder = mEntryOrder
End Property

Public Property Let entryReason(ByVal value As String)
mEntryReason = value
If mHandle > 0 Then
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonEntry, mEntryReason
End If
End Property

Public Property Get entryReason() As String
entryReason = mEntryReason
End Property

Public Property Get indexApplication() As Long
indexApplication = mIndexApplication
End Property

Friend Property Let key(ByVal value As String)
mKey = value
End Property

Public Property Get key() As String
key = mKey
End Property

Public Property Get longPosition() As Boolean
longPosition = (mEntryOrder.action = ActionBuy)
End Property

Public Property Let lossNotificationThreshold(ByVal value As Currency)
mLossNotificationThreshold = value
End Property

Public Property Let newEntryPrice(ByVal value As Double)
Select Case mStateEngine.state
Case OrderPlexStateCodes.Cancelling
    Exit Property
Case OrderPlexStateCodes.Submitted
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::(let)newEntryPrice", _
                "Action not allowed in state " & mStateEngine.state
End Select

If mEntryOrder.status = TradeBuild.OrderStatuses.OrderStatusFilled Then err.Raise errCantChangeFilledOrder, _
                                                "TradeBuild.OrderPlex::(let)newEntryPrice", _
                                                "Can't change entry price - order " & entryOrder.id & " already filled"

Select Case mEntryOrder.orderType
    Case TradeBuild.OrderTypes.OrderTypeMarket, _
        TradeBuild.OrderTypes.OrderTypeTrail, _
        TradeBuild.OrderTypes.OrderTypeAutoStop, _
        TradeBuild.OrderTypes.OrderTypeAutoLimit, _
        TradeBuild.OrderTypes.OrderTypeStop
        err.Raise errPriceChangeNotAllowedForOrderType, _
                    "TradeBuild.OrderPlex::(let)newEntryPrice", _
                    "Can't change entry price for order type " & orderTypeToString(mEntryOrder.orderType)
End Select

If value <> mEntryOrder.limitPrice Then
    mNewEntryPrice = value
    mDirty = True
    RaiseEvent Dirty
End If
End Property

Public Property Let newEntryTriggerPrice(ByVal value As Double)
Select Case mStateEngine.state
Case OrderPlexStateCodes.Cancelling
    Exit Property
Case OrderPlexStateCodes.Submitted
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::(let)newEntryTriggerPrice", _
                "Action not allowed in state " & mStateEngine.state
End Select

If mEntryOrder.status = TradeBuild.OrderStatuses.OrderStatusFilled Then err.Raise errCantChangeFilledOrder, _
                                                "TradeBuild.OrderPlex::(let)newEntryTriggerPrice", _
                                                "Can't change entry trigger price - order " & entryOrder.id & " already filled"

Select Case mEntryOrder.orderType
    Case TradeBuild.OrderTypes.OrderTypeMarket, _
        TradeBuild.OrderTypes.OrderTypeLimit, _
        TradeBuild.OrderTypes.OrderTypeTrail, _
        TradeBuild.OrderTypes.OrderTypeAutoLimit
        err.Raise errTriggerPriceChangeNotAllowedForOrderType, _
                    "TradeBuild.OrderPlex::(let)newEntryTriggerPrice", _
                    "Can't change trigger price for order type " & orderTypeToString(mEntryOrder.orderType)
End Select

If value <> mEntryOrder.triggerPrice Then
    mNewEntryTriggerPrice = value
    mDirty = True
    RaiseEvent Dirty
End If
End Property

Public Property Let newOrderPrice( _
                ByVal orderIndex As Long, _
                ByVal value As Double)
Dim lOrder As order

Set lOrder = mOrders(orderIndex)

If Not mEntryOrder Is Nothing Then
    If lOrder Is mEntryOrder Then
        newEntryPrice = value
        Exit Property
    End If
End If

If Not mStopOrder Is Nothing Then
    If lOrder Is mStopOrder Then
        newStopPrice = value
        Exit Property
    End If
End If

If Not mTargetOrder Is Nothing Then
    If lOrder Is mTargetOrder Then
        newTargetPrice = value
        Exit Property
    End If
End If
End Property

Public Property Let newOrderTriggerPrice( _
                ByVal orderIndex As Long, _
                ByVal value As Double)
Dim lOrder As order

Set lOrder = mOrders(orderIndex)

If Not mEntryOrder Is Nothing Then
    If lOrder Is mEntryOrder Then
        newEntryTriggerPrice = value
        Exit Property
    End If
End If

If Not mStopOrder Is Nothing Then
    If lOrder Is mStopOrder Then
        newStopTriggerPrice = value
        Exit Property
    End If
End If

If Not mTargetOrder Is Nothing Then
    If lOrder Is mTargetOrder Then
        newTargetTriggerPrice = value
        Exit Property
    End If
End If
End Property

Public Property Let newQuantity(ByVal value As Long)
Select Case mStateEngine.state
Case OrderPlexStateCodes.Cancelling
    Exit Property
Case OrderPlexStateCodes.Submitted
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::(let)newQuantity", _
                "Action not allowed in state " & mStateEngine.state
End Select

If mEntryOrder.status = TradeBuild.OrderStatuses.OrderStatusFilled Then err.Raise errCantChangeFilledOrder, _
                                                "TradeBuild.OrderPlex::(let)newQuantity", _
                                                "Can't change quantity - order " & entryOrder.id & " already filled"

If value <> mEntryOrder.Quantity Then
    mNewQuantity = value
    mDirty = True
    RaiseEvent Dirty
End If
End Property

Public Property Let newStopOffset(ByVal value As Long)
Select Case mStateEngine.state
Case OrderPlexStateCodes.Submitted
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::(let)newStopOffset", _
                "Action not allowed in state " & mStateEngine.state
End Select

If mStopOrder Is Nothing Then err.Raise errNoStopOrder, _
                                        "TradeBuild.OrderPlex::(let)newStopOffset", _
                                        "No stop order exists"

If Not mStopOrder.orderType = TradeBuild.OrderTypes.OrderTypeAutoStop Then
        err.Raise errOffsetChangeNotAllowedForOrderType, _
                    "TradeBuild.OrderPlex::(let)newStopOffset", _
                    "Can't change offset for order type " & orderTypeToString(mEntryOrder.orderType)
End If

If mStopOrder.status <> TradeBuild.OrderStatuses.OrderStatusCreated Then err.Raise errCantChangeUnplacedOrder, _
                                        "TradeBuild.OrderPlex::(let)newStopOffset", _
                                        "Can't change stop offset - order " & mStopOrder.id & " already placed"

If value <> mStopOrder.offset Then
    mNewStopOffset = value
    mDirty = True
    RaiseEvent Dirty
End If
End Property

Public Property Let newStopPrice(ByVal value As Double)
Select Case mStateEngine.state
Case OrderPlexStateCodes.Cancelling
    Exit Property
Case OrderPlexStateCodes.Submitted
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::(let)newStopPrice", _
                "Action not allowed in state " & mStateEngine.state
End Select

If mStopOrder Is Nothing Then err.Raise errNoStopOrder, _
                                        "TradeBuild.OrderPlex::(let)newStopPrice", _
                                        "No stop order exists"

Select Case mStopOrder.orderType
Case TradeBuild.OrderTypes.OrderTypeStopLimit
Case Else
    err.Raise errTriggerPriceChangeNotAllowedForOrderType, _
                "TradeBuild.OrderPlex::(let)newStopPrice", _
                "Can't change trigger price for order type " & orderTypeToString(mEntryOrder.orderType)
End Select

If value <> mStopOrder.limitPrice Then
    mNewStopPrice = value
    mDirty = True
    RaiseEvent Dirty
End If
End Property

Public Property Let newStopTriggerPrice(ByVal value As Double)
Select Case mStateEngine.state
Case OrderPlexStateCodes.Cancelling
    Exit Property
Case OrderPlexStateCodes.Submitted
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::(let)newStopTriggerPrice", _
                "Action not allowed in state " & mStateEngine.state
End Select

If mStopOrder Is Nothing Then err.Raise errNoStopOrder, _
                                        "TradeBuild.OrderPlex::(let)newStopTriggerPrice", _
                                        "No stop order exists"

Select Case mStopOrder.orderType
    Case TradeBuild.OrderTypes.OrderTypeStop, _
        TradeBuild.OrderTypes.OrderTypeStopLimit
    Case Else
        err.Raise errTriggerPriceChangeNotAllowedForOrderType, _
                    "TradeBuild.OrderPlex::(let)newStopTriggerPrice", _
                    "Can't change trigger price for order type " & orderTypeToString(mEntryOrder.orderType)
End Select

If value <> mStopOrder.triggerPrice Then
    mNewStopTriggerPrice = value
    mDirty = True
    RaiseEvent Dirty
End If
End Property

Public Property Let newTargetPrice(ByVal value As Double)
Select Case mStateEngine.state
Case OrderPlexStateCodes.Submitted
Case OrderPlexStateCodes.Cancelling
    Exit Property
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::(let)newTargetPrice", _
                "Action not allowed in state " & mStateEngine.state
End Select

If mTargetOrder Is Nothing Then err.Raise errNoTargetOrder, _
                                        "TradeBuild.OrderPlex::(let)newTargetPrice", _
                                        "No target order exists"

Select Case mTargetOrder.orderType
Case TradeBuild.OrderTypes.OrderTypeLimitIfTouched, _
    TradeBuild.OrderTypes.OrderTypeLimit
Case Else
    err.Raise errTriggerPriceChangeNotAllowedForOrderType, _
                "TradeBuild.OrderPlex::(let)newTargetPrice", _
                "Can't change price for order type " & orderTypeToString(mEntryOrder.orderType)
End Select

If value <> mTargetOrder.limitPrice Then
    mNewTargetPrice = value
    mDirty = True
    RaiseEvent Dirty
End If
End Property

Public Property Let newTargetOffset(ByVal value As Long)
Select Case mStateEngine.state
Case OrderPlexStateCodes.Submitted
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::(let)newTargetOffset", _
                "Action not allowed in state " & mStateEngine.state
End Select

If mTargetOrder Is Nothing Then err.Raise errNoTargetOrder, _
                                        "TradeBuild.OrderPlex::(let)newTargetOffset", _
                                        "No target order exists"

If Not mTargetOrder.orderType = TradeBuild.OrderTypes.OrderTypeAutoLimit Then
        err.Raise errOffsetChangeNotAllowedForOrderType, _
                    "TradeBuild.OrderPlex::(let)newTargetOffset", _
                    "Can't change offset for order type " & orderTypeToString(mEntryOrder.orderType)
End If

If mTargetOrder.status <> TradeBuild.OrderStatuses.OrderStatusCreated Then err.Raise errCantChangeUnplacedOrder, _
                                        "TradeBuild.OrderPlex::(let)newTargetOffset", _
                                        "Can't change target offset - order " & targetOrder.id & " already placed"

If value <> mTargetOrder.offset Then
    mNewTargetOffset = value
    mDirty = True
    RaiseEvent Dirty
End If
End Property

Public Property Let newTargetTriggerPrice(ByVal value As Double)
Select Case mStateEngine.state
Case OrderPlexStateCodes.Cancelling
    Exit Property
Case OrderPlexStateCodes.Submitted
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::(let)newTargetTriggerPrice", _
                "Action not allowed in state " & mStateEngine.state
End Select

If mTargetOrder Is Nothing Then err.Raise errNoStopOrder, _
                                        "TradeBuild.OrderPlex::(let)newTargetTriggerPrice", _
                                        "No target order exists"

Select Case mTargetOrder.orderType
Case TradeBuild.OrderTypes.OrderTypeLimitIfTouched, _
    TradeBuild.OrderTypes.OrderTypeMarketIfTouched
Case Else
    err.Raise errTriggerPriceChangeNotAllowedForOrderType, _
                "TradeBuild.OrderPlex::(let)newTargetTriggerPrice", _
                "Can't change trigger price for order type " & orderTypeToString(mEntryOrder.orderType)
End Select

If value <> mTargetOrder.triggerPrice Then
    mNewTargetTriggerPrice = value
    mDirty = True
    RaiseEvent Dirty
End If
End Property

Public Property Get pendingSize() As Long
If mEntryOrder Is Nothing Then pendingSize = 0: Exit Property
If mStateEngine.state = OrderPlexStateCodes.Closed Then pendingSize = 0: Exit Property
pendingSize = IIf(longPosition, mEntryOrder.Quantity, -mEntryOrder.Quantity)
End Property

Public Property Get order(ByVal index As Long) As order
Set order = mOrders(index)
End Property

Friend Property Let OrderContext(ByVal value As OrderContext)
Set mOrderContext = value
End Property

Public Property Get numberOfOrders() As Long
numberOfOrders = mOrders.Count
End Property

Friend Property Let positionManager(ByVal value As positionManager)
Set mPositionManager = value
End Property

Public Property Let profitNotificationThreshold(ByVal value As Currency)
mProfitNotificationThreshold = value
End Property

Public Property Get risk() As Currency
Dim stopTotalPrice As Double
Dim pendingTotalPrice As Double
Dim filledTotalPrice As Double

If mStopOrder Is Nothing Then Exit Property

If mStateEngine.state = OrderPlexStateCodes.Closed Then Exit Property

If mStopOrder.status = TradeBuild.OrderStatuses.OrderStatusFilled Then Exit Property

' the amount at risk has two parts - that due to fills already obtained, and
' that due to fills pending.

filledTotalPrice = mEntryOrder.quantityFilled * mEntryOrder.averagePrice
Select Case entryOrder.orderType
    Case TradeBuild.OrderTypes.OrderTypeMarket
        pendingTotalPrice = mEntryOrder.Quantity * _
                IIf(longPosition, mTicker.AskPrice, mTicker.BidPrice)
    Case TradeBuild.OrderTypes.OrderTypeLimit, TradeBuild.OrderTypes.OrderTypeStopLimit
        pendingTotalPrice = mEntryOrder.Quantity * mEntryOrder.limitPrice
    Case TradeBuild.OrderTypes.OrderTypeStop
        pendingTotalPrice = mEntryOrder.Quantity * mEntryOrder.triggerPrice
End Select
stopTotalPrice = mStopOrder.triggerPrice * _
                        (mEntryOrder.Quantity + mEntryOrder.quantityFilled)
If longPosition Then
    risk = mTickSize * mTickValue * (filledTotalPrice + pendingTotalPrice - stopTotalPrice) / mTickSize
Else
    risk = mTickSize * mTickValue * (stopTotalPrice - filledTotalPrice - pendingTotalPrice) / mTickSize
End If

End Property

Friend Property Let scalingStrategy( _
                ByVal value As PositionManagementStrategy)
Set mScalingStrategy = value
End Property

Public Property Get size() As Long
size = mSize
End Property

Public Property Get state() As OrderPlexStateCodes
state = mStateEngine.state
End Property

Public Property Get stopOrder() As order
Set stopOrder = mStopOrder
End Property

Public Property Let stopReason(ByVal value As String)
mStopReason = value
If mHandle > 0 Then
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonStop, mStopReason
End If
End Property

Public Property Get stopReason() As String
stopReason = mStopReason
End Property

Friend Property Let stopStrategy( _
                ByVal value As StopManagementStrategy)
Set mStopStrategy = value
End Property

Public Property Get targetOrder() As order
Set targetOrder = mTargetOrder
End Property

Public Property Let targetReason(ByVal value As String)
mTargetReason = value
If mHandle > 0 Then
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonTarget, mTargetReason
End If
End Property

Public Property Get targetReason() As String
targetReason = mTargetReason
End Property

Friend Property Let ticker(ByVal value As ticker)
Dim lContract As contract
Set mTicker = value
Set lContract = mTicker.contract
mTickSize = lContract.tickSize
mTickValue = lContract.tickValue
mValueFactor = lContract.multiplier
mcurrencyCode = lContract.specifier.currencyCode
End Property

Public Property Get ticker() As ticker
Set ticker = mTicker
End Property

'================================================================================
' Methods
'================================================================================

Public Sub addChangeListener(ByVal value As ChangeListener)
mChangeListeners.add value

' notify this new listenener that we are already created
Dim ev As ChangeEvent
Set ev.Source = Me
ev.ChangeType = OrderPlexChangeTypes.Created
value.Change ev
End Sub

Public Sub addProfitListener(ByVal value As ProfitListener)
mProfitListeners.add value
End Sub

Public Sub adjustStop(ByVal stopPrice As Double, _
                    Optional ByVal size As Long = -1, _
                    Optional ByVal allowWider As Boolean = False)
Select Case mStateEngine.state
Case OrderPlexStateCodes.Cancelling
    Exit Sub
Case OrderPlexStateCodes.Submitted
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::adjustStop", _
                "Action not allowed in state " & mStateEngine.state
End Select

If mStopOrder Is Nothing Then err.Raise errNoStopOrder, _
                                        "TradeBuild.OrderPlex::adjustStop", _
                                        "No stop order exists"

If mStopOrder.orderType = TradeBuild.OrderTypes.OrderTypeAutoStop Then err.Raise errCantChangeStopForOrderType, _
                                        "TradeBuild.OrderPlex::adjustStop", _
                                        "Can't change the stop for an AUTOSTOP order"""

If Not allowWider Then
    If (longPosition And _
        stopPrice < mStopOrder.triggerPrice) Or _
        ((Not longPosition) And _
        stopPrice > mStopOrder.triggerPrice) _
    Then
        Exit Sub
'        err.Raise errWidenStopNotAllowed, _
'                    "TradeBuild.OrderPlex::adjustStop", _
'                    "Attempt to widen the stop"
    End If
End If


mStopOrder.triggerPrice = stopPrice
If size <> -1 Then mStopOrder.Quantity = size
mTicker.notifyListeners _
                            "Adjust stop: (" & mTicker.TradePrice & ") " & _
                            " stop " & stopPrice & _
                            ", size " & mStopOrder.Quantity, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
mStopOrder.transmit = True
gTradeBuildAPI.placeOrder mStopOrder

fireChange StopOrderChanged
RaiseEvent StopAdjusted
End Sub

Public Sub adjustTarget(ByVal targetPrice As Double, _
                    Optional ByVal size As Long = -1)
Select Case mStateEngine.state
Case OrderPlexStateCodes.Cancelling
    Exit Sub
Case OrderPlexStateCodes.Submitted
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::adjustTarget", _
                "Action not allowed in state " & mStateEngine.state
End Select

If mTargetOrder Is Nothing Then err.Raise errNoTargetOrder, _
                                        "TradeBuild.OrderPlex::adjustTarget", _
                                        "No target order exists"

If mTargetOrder.orderType = TradeBuild.OrderTypes.OrderTypeAutoLimit Then err.Raise errCantChangeTargetForOrderType, _
                                        "TradeBuild.OrderPlex::adjustTarget", _
                                        "Can't change the target for an AUTOLMT order"""

mTargetOrder.limitPrice = targetPrice
If size <> -1 Then mTargetOrder.Quantity = size
mTicker.notifyListeners _
                            "Adjust target: (" & mTicker.TradePrice & ") " & _
                            " target " & targetPrice & _
                            ", size " & mTargetOrder.Quantity, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
mTargetOrder.transmit = True
gTradeBuildAPI.placeOrder mTargetOrder

fireChange TargetOrderChanged
RaiseEvent TargetAdjusted
End Sub

Public Function Cancel(Optional ByVal evenIfFilled As Boolean = False) As Boolean

If evenIfFilled Then
    mStateEngine.clearCondition COND_NO_FILL_CANCELLATION
    Cancel = doActions(mStateEngine.stimulus(STIM_CANCEL_EVEN_IF_FILL))
Else
    mStateEngine.setCondition COND_NO_FILL_CANCELLATION
    Cancel = doActions(mStateEngine.stimulus(STIM_CANCEL_IF_NO_FILL))
End If
End Function

Public Sub cancelChanges()
mNewQuantity = -1
mNewEntryPrice = -1
mNewEntryTriggerPrice = -1
mNewStopPrice = -1
mNewStopOffset = DummyOffset
mNewStopTriggerPrice = -1
mNewTargetPrice = -1
mNewTargetOffset = DummyOffset
mNewTargetTriggerPrice = -1
mDirty = False
RaiseEvent ChangesCancelled
End Sub

Public Sub CloseOut(ByVal method As String)

If UCase(method) <> "MKT" Then
    err.Raise errNotImplemented, "TradeBuild.OrderPlex::closeout", "method " & method & " not implemented"
End If

mTicker.notifyListeners _
                    "Closing out: (" & mTicker.TradePrice & ") " & _
                    "size " & size & _
                    ", pending size " & pendingSize, _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)

handleStimulus STIM_CLOSEOUT

End Sub

Friend Function createEntryOrder() As order
If Not mEntryOrder Is Nothing Then _
    err.Raise ErrorCodes.IllegalStateException, _
                "TradeBuild.OrderPlex::createEntryOrder", _
                "Entry order already exists"

Set mEntryOrder = mTicker.newOrder
If mDummy Then mEntryOrder.dummy = True
Set createEntryOrder = mEntryOrder

If mOrders.Count >= 1 Then
    mOrders.add mEntryOrder, , 1
Else
    mOrders.add mEntryOrder
End If
End Function

Friend Function createStopOrder() As order
If Not mStopOrder Is Nothing Then _
    err.Raise ErrorCodes.IllegalStateException, _
                "TradeBuild.OrderPlex::createStopOrder", _
                "Stop order already exists"

Set mStopOrder = mTicker.newOrder
If dummy Then mStopOrder.dummy = True
Set createStopOrder = mStopOrder
mStateEngine.setCondition COND_STOP_ORDER_EXISTS
mStateEngine.clearCondition COND_STOP_ORDER_CANCELLED

If Not mEntryOrder Is Nothing Then
    mOrders.add mStopOrder, , , 1
ElseIf Not mTargetOrder Is Nothing Then
    mOrders.add mStopOrder, , 1
Else
    mOrders.add mStopOrder
End If
End Function

Friend Function createTargetOrder() As order
If Not mTargetOrder Is Nothing Then _
    err.Raise ErrorCodes.IllegalStateException, _
                "TradeBuild.OrderPlex::createTargetOrder", _
                "Target order already exists"

Set mTargetOrder = mTicker.newOrder
If mDummy Then mTargetOrder.dummy = True
Set createTargetOrder = mTargetOrder
mStateEngine.setCondition COND_TARGET_ORDER_EXISTS
mStateEngine.clearCondition COND_TARGET_ORDER_CANCELLED

If mEntryOrder Is Nothing And _
    mStopOrder Is Nothing _
Then
    mOrders.add mTargetOrder
ElseIf mEntryOrder Is Nothing Then
    mOrders.add mTargetOrder, , , 1
ElseIf mStopOrder Is Nothing Then
    mOrders.add mTargetOrder, , , 1
Else
    mOrders.add mTargetOrder, , , 2
End If
End Function

Friend Sub execute(Optional ByVal cancelPrice As Double = 0)
mCancelPrice = cancelPrice
fireChange OrderPlexChangeTypes.Created
handleStimulus STIM_EXECUTE
End Sub

Friend Sub finish()
Set mTicker = Nothing
Set mPositionManager = Nothing
Set mOrderContext = Nothing
If Not mEntryOrder Is Nothing Then mEntryOrder.finish
If Not mStopOrder Is Nothing Then mStopOrder.finish
If Not mTargetOrder Is Nothing Then mTargetOrder.finish
If Not mCloseoutOrder Is Nothing Then mCloseoutOrder.finish
End Sub

Public Sub removeChangeListener(ByVal value As ChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next
End Sub

Public Sub removeProfitListener(ByVal value As ProfitListener)
Dim i As Long
For i = mProfitListeners.Count To 1 Step -1
    If mProfitListeners.Item(i) Is value Then mProfitListeners.Remove i
Next
End Sub

Friend Sub terminate()
Set mStopStrategy = Nothing
Set mScalingStrategy = Nothing
End Sub

Public Sub update()
Dim updateEntry As Boolean
Dim updateStop As Boolean
Dim updateTarget As Boolean

Select Case mStateEngine.state
Case OrderPlexStateCodes.Submitted
Case OrderPlexStateCodes.Cancelling
    Exit Sub
Case Else
    err.Raise errActionNotAllowedInThisState, _
                "TradeBuild.OrderPlex::update", _
                "Action not allowed in state " & mStateEngine.state
End Select

If Not mEntryOrder Is Nothing Then
    If mNewQuantity >= 0 And _
        mNewQuantity <> mEntryOrder.Quantity _
    Then
        If mEntryOrder.status = TradeBuild.OrderStatuses.OrderStatusFilled Then err.Raise errCantChangeFilledOrder, _
                                                        "TradeBuild.OrderPlex::update", _
                                                        "Can't change size - order " & entryOrder.id & " already filled"
        mEntryOrder.Quantity = mNewQuantity
'       I don't think the following are necessary, because should update them
'       appropriately. Besides the code is too simplistic - if the order has
'       already been partially filled, it wouldn't work
'        If Not mStopOrder Is Nothing Then
'            mStopOrder.Quantity = mNewQuantity
'            updateStop = True
'        End If
'        If Not mTargetOrder Is Nothing Then
'            mTargetOrder.Quantity = mNewQuantity
'            updateTarget = True
'        End If
        updateEntry = True
    End If
    mNewQuantity = -1
    
    If mNewEntryPrice >= 0 And _
        mNewEntryPrice <> mEntryOrder.limitPrice _
    Then
        If mEntryOrder.status = TradeBuild.OrderStatuses.OrderStatusFilled Then err.Raise errCantChangeFilledOrder, _
                                                        "TradeBuild.OrderPlex::update", _
                                                        "Can't change price - order " & mEntryOrder.id & " already filled"
        mEntryOrder.limitPrice = mNewEntryPrice
        updateEntry = True
    End If
    mNewEntryPrice = -1
    
    If mNewEntryTriggerPrice >= 0 And _
        mNewEntryTriggerPrice <> mEntryOrder.triggerPrice _
    Then
        If mEntryOrder.status = TradeBuild.OrderStatuses.OrderStatusFilled Then err.Raise errCantChangeFilledOrder, _
                                                        "TradeBuild.OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.id & " already filled"
        mEntryOrder.triggerPrice = mNewEntryTriggerPrice
        updateEntry = True
    End If
    mNewEntryTriggerPrice = -1
End If

If Not mStopOrder Is Nothing Then
    If mNewStopPrice >= 0 And _
        mNewStopPrice <> mStopOrder.triggerPrice _
    Then
        If mStopOrder.status = TradeBuild.OrderStatuses.OrderStatusFilled Then err.Raise errCantChangeFilledOrder, _
                                                        "TradeBuild.OrderPlex::update", _
                                                        "Can't change trigger price - order " & mStopOrder.id & " already filled"
        If mStopOrder.status = TradeBuild.OrderStatuses.OrderStatusSubmitted Then err.Raise errCantChangeSubmittedOrder, _
                                                        "TradeBuild.OrderPlex::update", _
                                                        "Can't change price - order " & mStopOrder.id & " already submitted"
        mStopOrder.triggerPrice = mNewStopPrice
        updateStop = True
    End If
    mNewStopPrice = -1
    
    If mNewStopOffset <> DummyOffset And _
        mNewStopOffset <> mStopOrder.offset _
    Then
        If mStopOrder.status <> TradeBuild.OrderStatuses.OrderStatusCreated Then err.Raise errCantChangeUnplacedOrder, _
                                                            "TradeBuild.OrderPlex::update", _
                                                            "Can't change stop offset - order " & mStopOrder.id & " already placed"
        mStopOrder.offset = mNewStopOffset
    End If
    mNewStopOffset = DummyOffset

    If mNewStopTriggerPrice >= 0 And _
        mNewStopTriggerPrice <> mStopOrder.triggerPrice _
    Then
        If mStopOrder.status = TradeBuild.OrderStatuses.OrderStatusFilled Then err.Raise errCantChangeFilledOrder, _
                                                        "TradeBuild.OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.id & " already filled"
        mStopOrder.triggerPrice = mNewStopTriggerPrice
        updateStop = True
    End If
    mNewStopTriggerPrice = -1
End If

If Not mTargetOrder Is Nothing Then
    If mNewTargetPrice >= 0 And _
        mNewTargetPrice <> mTargetOrder.limitPrice _
    Then
        If mTargetOrder.status = TradeBuild.OrderStatuses.OrderStatusFilled Then err.Raise errCantChangeFilledOrder, _
                                                        "TradeBuild.OrderPlex::update", _
                                                        "Can't change price - order " & mTargetOrder.id & " already filled"
        mTargetOrder.limitPrice = mNewTargetPrice
        updateTarget = True
    End If
    mNewTargetPrice = -1
    
    If mNewTargetOffset <> DummyOffset And _
        mNewTargetOffset <> mTargetOrder.offset _
    Then
        If mTargetOrder.status <> TradeBuild.OrderStatuses.OrderStatusCreated Then err.Raise errCantChangeUnplacedOrder, _
                                                            "TradeBuild.OrderPlex::update", _
                                                            "Can't change target offset - order " & targetOrder.id & " already placed"
        mTargetOrder.offset = mNewTargetOffset
    End If
    mNewTargetOffset = DummyOffset

    If mNewTargetTriggerPrice >= 0 And _
        mNewTargetTriggerPrice <> mTargetOrder.triggerPrice _
    Then
        If mTargetOrder.status = TradeBuild.OrderStatuses.OrderStatusFilled Then err.Raise errCantChangeFilledOrder, _
                                                        "TradeBuild.OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.id & " already filled"
        mTargetOrder.triggerPrice = mNewTargetTriggerPrice
        updateTarget = True
    End If
    mNewTargetTriggerPrice = -1
End If

If updateEntry Then
    mEntryOrder.transmit = True
    gTradeBuildAPI.placeOrder mEntryOrder
    fireChange EntryOrderChanged
End If
If updateStop Then
    mStopOrder.transmit = True
    gTradeBuildAPI.placeOrder mStopOrder
    fireChange StopOrderChanged
    RaiseEvent StopAdjusted
End If
If updateTarget Then
    mTargetOrder.transmit = True
    gTradeBuildAPI.placeOrder mTargetOrder
    fireChange TargetOrderChanged
    RaiseEvent TargetAdjusted
End If
RaiseEvent Clean
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function calcProfitLoss(ByVal value As Double)
Dim profit As Currency
Dim notifyProfit As Boolean
Dim notifyMaxProfit As Boolean
Dim notifyDrawdown As Boolean

profit = size * (value - mEntryOrder.averagePrice) * mValueFactor

If profit <> mProfit Then
    mProfit = profit
    notifyProfit = True
End If

If profit > mMaxProfit Then
    mMaxProfit = profit
    notifyMaxProfit = True
End If

If profit < mMaxProfit Then
    mDrawdown = mMaxProfit - profit
    notifyDrawdown = True
ElseIf mDrawdown <> 0 Then
    mDrawdown = 0
    notifyDrawdown = True
End If

If notifyProfit Then fireProfit mProfit, OrderPlexProfitTypes.profit
If notifyMaxProfit Then fireProfit mMaxProfit, OrderPlexProfitTypes.MaxProfit
If notifyDrawdown Then fireProfit mDrawdown, OrderPlexProfitTypes.drawdown

If notifyProfit Then
    If mProfitNotificationThreshold <> 0 And _
        mProfit >= mProfitNotificationThreshold _
    Then
        fireChange ProfitThresholdExceeded
        RaiseEvent ProfitThresholdExceeded
        mProfitNotificationThreshold = 0
    End If
    
    If mLossNotificationThreshold <> 0 And _
        mProfit <= mLossNotificationThreshold _
    Then
        fireChange LossThresholdExceeded
        RaiseEvent LossThresholdExceeded
        mLossNotificationThreshold = 0
    End If

    If mDrawdownNotificationThreshold <> 0 And _
        (mMaxProfit - profit >= mDrawdownNotificationThreshold) _
    Then
        fireChange DrawdownThresholdExceeded
        RaiseEvent DrawdownThresholdExceeded
        mDrawdownNotificationThreshold = 0
    End If
End If
End Function

Private Function cancelOrders() As Boolean

Dim entryOrderID As Long
Dim stopOrderID As Long
Dim targetOrderID As Long

cancelAfter = 0

mCancelPrice = 0

If Not mEntryOrder Is Nothing Then
    Select Case mEntryOrder.status
    Case TradeBuild.OrderStatuses.OrderStatusCreated
        mEntryOrder.updateStatus TradeBuild.OrderStatuses.OrderStatusCancelled, _
                                0, _
                                0, _
                                0, _
                                0, _
                                0, _
                                0, _
                                mTicker.timestamp
                                
        cancelOrders = True
                                
    Case TradeBuild.OrderStatuses.OrderStatusFilled, _
        TradeBuild.OrderStatuses.OrderStatusCancelling, _
        TradeBuild.OrderStatuses.OrderStatusCancelled
    Case Else
        entryOrderID = mEntryOrder.id
        gTradeBuildAPI.cancelOrder entryOrderID  ' should automatically cancel the other orders
                                                        ' if they have parentid set
        cancelOrders = True
    End Select
End If
If Not mStopOrder Is Nothing Then
    Select Case mStopOrder.status
    Case TradeBuild.OrderStatuses.OrderStatusCreated
        mStopOrder.updateStatus TradeBuild.OrderStatuses.OrderStatusCancelled, _
                                0, _
                                0, _
                                0, _
                                0, _
                                0, _
                                0, _
                                mTicker.timestamp
                                
        cancelOrders = True
                                
    Case TradeBuild.OrderStatuses.OrderStatusFilled, _
        TradeBuild.OrderStatuses.OrderStatusCancelling, _
        TradeBuild.OrderStatuses.OrderStatusCancelled
    Case Else
        stopOrderID = mStopOrder.id
        If mStopOrder.parentId = 0 Then
            gTradeBuildAPI.cancelOrder stopOrderID
            cancelOrders = True
        Else
            Select Case mEntryOrder.status
            Case TradeBuild.OrderStatuses.OrderStatusFilled, _
                TradeBuild.OrderStatuses.OrderStatusCancelling, _
                TradeBuild.OrderStatuses.OrderStatusCancelled
                gTradeBuildAPI.cancelOrder stopOrderID
                cancelOrders = True
            End Select
        End If
    End Select
End If
If Not mTargetOrder Is Nothing Then
    Select Case mTargetOrder.status
    Case TradeBuild.OrderStatuses.OrderStatusCreated
        mTargetOrder.updateStatus TradeBuild.OrderStatuses.OrderStatusCancelled, _
                                0, _
                                0, _
                                0, _
                                0, _
                                0, _
                                0, _
                                mTicker.timestamp
                                
        cancelOrders = True
                                
    Case TradeBuild.OrderStatuses.OrderStatusFilled, _
        TradeBuild.OrderStatuses.OrderStatusCancelling, _
        TradeBuild.OrderStatuses.OrderStatusCancelled
    Case Else
        targetOrderID = mTargetOrder.id
        If mTargetOrder.parentId = 0 Then
            gTradeBuildAPI.cancelOrder targetOrderID
            cancelOrders = True
        Else
            Select Case mEntryOrder.status
            Case TradeBuild.OrderStatuses.OrderStatusFilled, _
                TradeBuild.OrderStatuses.OrderStatusCancelling, _
                TradeBuild.OrderStatuses.OrderStatusCancelled
                gTradeBuildAPI.cancelOrder targetOrderID
                cancelOrders = True
            End Select
        End If
    End Select
End If
    
' need some stuff here to cancel if it's an oca group

If cancelOrders Then
    mTicker.notifyListeners _
                    "Cancel orders: (" & mTicker.TradePrice & ") " & _
                    IIf(entryOrderID <> 0, entryOrderID, "") & _
                    IIf(stopOrderID <> 0, ", " & stopOrderID, ", ") & _
                    IIf(targetOrderID <> 0, ", " & targetOrderID, ", "), _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
End If
End Function

Private Function cancelStopOrder()
If mStopOrder Is Nothing Then Exit Function
Select Case mStopOrder.status
Case TradeBuild.OrderStatuses.OrderStatusCreated
    mStopOrder.updateStatus TradeBuild.OrderStatuses.OrderStatusCancelled, _
                            0, _
                            0, _
                            0, _
                            0, _
                            0, _
                            0, _
                            mTicker.timestamp
                            
                            
Case TradeBuild.OrderStatuses.OrderStatusFilled, _
    TradeBuild.OrderStatuses.OrderStatusCancelling, _
    TradeBuild.OrderStatuses.OrderStatusCancelled
Case Else
    mTicker.notifyListeners _
                "Cancel order: (" & mTicker.TradePrice & ") " & _
                mStopOrder.id, _
                IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
    gTradeBuildAPI.cancelOrder mStopOrder.id
End Select
End Function

Private Function cancelTargetOrder()
If mTargetOrder Is Nothing Then Exit Function
Select Case mTargetOrder.status
Case TradeBuild.OrderStatuses.OrderStatusCreated
    mTargetOrder.updateStatus TradeBuild.OrderStatuses.OrderStatusCancelled, _
                            0, _
                            0, _
                            0, _
                            0, _
                            0, _
                            0, _
                            mTicker.timestamp
                            
                            
Case TradeBuild.OrderStatuses.OrderStatusFilled, _
    TradeBuild.OrderStatuses.OrderStatusCancelling, _
    TradeBuild.OrderStatuses.OrderStatusCancelled
Case Else
    mTicker.notifyListeners _
                "Cancel order: (" & mTicker.TradePrice & ") " & _
                mTargetOrder.id, _
                IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
    gTradeBuildAPI.cancelOrder mTargetOrder.id
End Select
End Function

Private Sub checkAllOrdersComplete()
If mStateEngine.state = OrderPlexStateCodes.Closed Then Exit Sub
If mEntryOrderFinished And _
    mStopOrderFinished And _
    mTargetOrderFinished _
Then
    handleStimulus STIM_ALL_ORDERS_COMPLETE
End If

End Sub

Private Function completionActions()
mOrderContext.orderPlexClosed Me
fireChange Completed
RaiseEvent Completed
Set mStopStrategy = Nothing
Set mScalingStrategy = Nothing
If size = 0 Then mTicker.removeQuoteListener Me
End Function

Private Function doActions(ByRef Actions() As Long) As Variant
Dim action As Variant

For Each action In Actions
    Select Case CLng(action)
    Case SpecialActions.NO_ACTION
    Case ACT_PLACE_ORDERS
        doActions = placeOrders(simulatorOnly:=False)
    Case ACT_CANCEL_ORDERS
        doActions = cancelOrders
    Case ACT_CANCEL_STOP_ORDER
        doActions = cancelStopOrder
    Case ACT_CANCEL_TARGET_ORDER
        doActions = cancelTargetOrder
    Case ACT_RESUBMIT_STOP_ORDER
        doActions = resubmitStopOrder
    Case ACT_RESUBMIT_TARGET_ORDER
        doActions = resubmitTargetOrder
    Case ACT_RESUBMIT_STOP_AND_TARGET_ORDERS
        doActions = resubmitStopAndTargetOrders
    Case ACT_PLACE_CLOSEOUT_ORDER
        doActions = placeCloseoutOrder
    Case ACT_COMPLETION_ACTIONS
        doActions = completionActions
    Case ACT_ALARM
        ' need to put something in here
    Case Else
        err.Raise errInvalidActionReturnedFromStateEngine, _
                    "TradeBuild.OrderPlex::doActions", _
                    "Invalid action " & action
    End Select
Next
End Function

Private Sub fireChange( _
                ByVal ChangeType As OrderPlexChangeTypes)
Dim listener As ChangeListener
Dim Change As ChangeEvent
Set Change.Source = Me
Change.ChangeType = ChangeType
For Each listener In mChangeListeners
    listener.Change Change
Next
End Sub

Private Sub fireProfit( _
                ByVal profitAmount As Double, _
                ByVal profitType As OrderPlexProfitTypes)
Dim listener As ProfitListener
Dim profit As ProfitEvent
Set profit.Source = Me
profit.profitAmount = profitAmount
profit.profitType = profitType
profit.currencyCode = mcurrencyCode
For Each listener In mProfitListeners
    listener.profitAmount profit
Next
End Sub

Private Sub handleStimulus(ByVal stimulus As StateTransitionStimuli)
Dim currState As OrderPlexStateCodes

currState = mStateEngine.state
doActions mStateEngine.stimulus(stimulus)
If mStateEngine.state <> currState Then fireChange OrderPlexChangeTypes.StateChanged
End Sub

Private Function placeCloseoutOrder()
' place an order to closeout the filled part
Set mCloseoutOrder = mTicker.newOrder
mCloseoutOrder.action = IIf(longPosition, TradeBuild.OrderActions.ActionSell, TradeBuild.OrderActions.ActionBuy)
mCloseoutOrder.Quantity = Abs(size)
mCloseoutOrder.orderType = TradeBuild.OrderTypes.OrderTypeMarket
mCloseoutOrder.timeInForce = TradeBuild.OrderTifs.TIFGoodTillCancelled
mCloseoutOrder.transmit = True
fireChange CloseoutOrderCreated
gTradeBuildAPI.placeOrder mCloseoutOrder
                                

End Function

Private Function placeOrders(Optional ByVal simulatorOnly As Boolean = False)

Dim transmit As Boolean

mCreationTime = mTicker.timestamp

If mEntryOrder Is Nothing Then mEntryOrderFinished = True
If mStopOrder Is Nothing Then mStopOrderFinished = True
If mTargetOrder Is Nothing Then mTargetOrderFinished = True

If Not mEntryOrder Is Nothing Then
    transmit = True
    If Not mStopOrder Is Nothing Then
        If mStopOrder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoStop Then transmit = False
    End If
    If Not mTargetOrder Is Nothing Then
        If mTargetOrder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoLimit Then transmit = False
    End If
        
    mEntryOrder.transmit = transmit
    gTradeBuildAPI.placeOrder mEntryOrder
    
    If Not mStopOrder Is Nothing Then
        transmit = True
        If Not mTargetOrder Is Nothing Then
            If mTargetOrder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoLimit Then transmit = False
        End If
        
        If mStopOrder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoStop Then
            ' note that AUTOSTP orders will be sent when the entry order is filled
            ' and that they don't have parentID set since the would-be parent has
            ' already been transmitted
            mStopOrder.parentId = mEntryOrder.id
            mStopOrder.transmit = transmit
            gTradeBuildAPI.placeOrder mStopOrder
        End If
    End If
    
    If Not mTargetOrder Is Nothing Then
        If mTargetOrder.orderType <> TradeBuild.OrderTypes.OrderTypeAutoLimit Then
            mTargetOrder.parentId = mEntryOrder.id
            mTargetOrder.transmit = True
            gTradeBuildAPI.placeOrder mTargetOrder
        End If
    End If
    
    mTicker.addQuoteListener Me
Else
    ' treat the other orders as an OCA group - still to be implemented
End If
End Function

Private Sub placePeggedOrders()
If Not mStopOrder Is Nothing Then
    If mStopOrder.status = TradeBuild.OrderStatuses.OrderStatusCreated Then
        If mStopOrder.orderType = TradeBuild.OrderTypes.OrderTypeAutoStop Then
            ' the stop price is to be pegged to the entry price - the
            ' order's offset property will already contain any required
            ' offset from that price
            mStopOrder.triggerPrice = (mStopOrder.offset * mTickSize) + entryOrder.lastFillPrice
            mStopOrder.Quantity = mEntryOrder.quantityFilled
            mStopOrder.orderType = TradeBuild.OrderTypes.OrderTypeStop
            mStopOrder.transmit = True
            gTradeBuildAPI.placeOrder mStopOrder
        End If
    End If
End If
If Not mTargetOrder Is Nothing Then
    If mTargetOrder.status = TradeBuild.OrderStatuses.OrderStatusCreated Then
        If mTargetOrder.orderType = TradeBuild.OrderTypes.OrderTypeAutoLimit Then
            ' the limit price is to be pegged to the entry price - the
            ' order's offset property will already contain any required
            ' offset from that price
            mTargetOrder.limitPrice = (mTargetOrder.offset * mTickSize) + entryOrder.lastFillPrice
            mTargetOrder.Quantity = mEntryOrder.quantityFilled
            mTargetOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
            mTargetOrder.transmit = True
            gTradeBuildAPI.placeOrder mTargetOrder
        End If
    End If
End If
End Sub

Private Function resubmitStopOrder()
mStopOrder.id = 0 ' force a new id to be allocated
mStopOrder.parentId = 0
If Not mTargetOrder Is Nothing Then
    mStopOrder.ocaGroup = mTargetOrder.ocaGroup
End If
mTicker.notifyListeners _
            "Resubmit stop order: (" & mTicker.TradePrice & ") new id " & _
            mStopOrder.id, _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
mStopOrder.transmit = True
gTradeBuildAPI.placeOrder mStopOrder
mStateEngine.clearCondition COND_STOP_ORDER_CANCELLED
mStopOrderFinished = False
End Function

Private Function resubmitStopAndTargetOrders()
Dim ocaGroup As String
ocaGroup = GenerateTextID

mStopOrder.ocaGroup = ocaGroup
mStopOrder.id = 0 ' force a new id to be allocated
mStopOrder.parentId = 0
mTicker.notifyListeners _
            "Resubmit stop order: (" & mTicker.TradePrice & ") new id " & _
            mStopOrder.id, _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
mStopOrder.transmit = False
gTradeBuildAPI.placeOrder mStopOrder
mStateEngine.clearCondition COND_STOP_ORDER_CANCELLED
mStopOrderFinished = False

mTargetOrder.ocaGroup = ocaGroup
mTargetOrder.id = 0 ' force a new id to be allocated
mTargetOrder.parentId = 0
mTicker.notifyListeners _
            "Resubmit target order: (" & mTicker.TradePrice & ") new id " & _
            mStopOrder.id, _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
mTargetOrder.transmit = True
gTradeBuildAPI.placeOrder mTargetOrder
mStateEngine.clearCondition COND_TARGET_ORDER_CANCELLED
mTargetOrderFinished = False

End Function

Private Function resubmitTargetOrder()
mTargetOrder.id = 0 ' force a new id to be allocated
mTargetOrder.parentId = 0
If Not mStopOrder Is Nothing Then
    mTargetOrder.ocaGroup = mStopOrder.ocaGroup
End If
mTicker.notifyListeners _
            "Resubmit target order: (" & mTicker.TradePrice & ") new id " & _
            mStopOrder.id, _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)
mTargetOrder.transmit = True
gTradeBuildAPI.placeOrder mTargetOrder
mStateEngine.clearCondition COND_TARGET_ORDER_CANCELLED
mTargetOrderFinished = False
End Function

'Private Sub resubmitToSimulator()
'If Not mEntryOrder Is Nothing Then
'    gInstrumentProcessor.removeSimulatedOrder mEntryOrder.ID
'End If
'If Not mStopOrder Is Nothing Then
'    gInstrumentProcessor.removeSimulatedOrder mStopOrder.ID
'End If
'If Not mTargetOrder Is Nothing Then
'    gInstrumentProcessor.removeSimulatedOrder mTargetOrder.ID
'End If
'
'placeOrders True
'End Sub

Private Sub setSize(ByVal value As Long)
mSize = value
If mSize <> 0 Then
    mStateEngine.setCondition COND_SIZE_NON_ZERO
Else
    mStateEngine.clearCondition COND_SIZE_NON_ZERO
End If
fireChange OrderPlexChangeTypes.SizeChanged
End Sub



