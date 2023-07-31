VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderUtilsEnums"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum BracketAutosizeTypes
    BracketAutosizeNone
    BracketAutosizePercentOfAccount
    BracketAutosizeMonetaryAmount
End Enum

Public Enum BracketOrderChangeTypes
    BracketOrderCreated = 1
    BracketOrderCompleted
    BracketOrderSelfCancelled
    BracketOrderEntryOrderChanged
    BracketOrderEntryOrderFilled
    BracketOrderStopLossOrderChanged
    BracketOrderStopLossOrderFilled
    BracketOrderTargetOrderChanged
    BracketOrderTargetOrderFilled
    BracketOrderCloseoutOrderCreated
    BracketOrderCloseoutOrderChanged
    BracketOrderCloseoutOrderFilled
    BracketOrderSizeChanged
    BracketOrderStateChanged
    BracketOrderChangesPending
    BracketOrderChangesApplied
    BracketOrderChangesCancelled
    BracketOrderPositionCloseoutNotified
    BracketOrderRolloverInitiated
    BracketOrderRolloverGotContract
    BracketOrderRolloverCompleted
    BracketOrderRolloverPending
End Enum

Public Enum BracketOrderRoles
    BracketOrderRoleNone = 0
    BracketOrderRoleEntry = 1
    BracketOrderRoleStopLoss
    BracketOrderRoleTarget
    BracketOrderRoleCloseout
End Enum

Public Enum BracketOrderStates

    ' This state indicates that the bracket order has been created but none of
    ' the orders have yet been placed.
    BracketOrderStateCreated = 1

    ' This state indicates that all the orders in the bracket order have been placed.
    BracketOrderStateSubmitted

    ' This state indicates that the bracket order is in the process of being
    ' cancelled.
    BracketOrderStateCancelling

    ' This state indicates that the bracket order is being closed out.
    ' (When a bracket order has been closed out, there are no outstanding
    ' orders and no net Size - note that this may require a closeout
    ' order to be placed to negate any existing Size.)
    BracketOrderStateClosingOut

    ' This state indicates that all orders have been completed.
    BracketOrderStateClosed

    ' This state indicates that a cancellation notification for either
    ' the stop-loss order or the target order, when both exist, has been received.
    ' We don't know whether the order was cancelled because the user/broker
    ' cancelled it, or because the other order was filled. So in this
    ' state we wait for the other order to be either cancelled or
    ' filled or for a timeout to elapse.
    BracketOrderStateAwaitingOtherOrderCancel

End Enum

Public Enum CloseoutModes
    CloseoutModeMarket
    CloseoutModeLimit
End Enum

Public Enum OrderActions
    OrderActionNone = 0
    OrderActionBuy = 1
    OrderActionSell = 2
End Enum

Public Enum OptionContractResolverStates
    OptionContractResolverStateNone
    OptionContractResolverStateFetchingUnderlyingContract
    OptionContractResolverStateWaitingForUnderlyingMarketPrice
    OptionContractResolverStateFetchingInitialContract
    OptionContractResolverStateFetchingCandidateContracts
    OptionContractResolverStateFindingRequiredContract
    OptionContractResolverStateGotTargetContract
    OptionContractResolverStateFetchingStrikes
    OptionContractResolverStateFetchingExpiries
End Enum

Public Enum OptionStrikeSelectionModes
    OptionStrikeSelectionModeNone
    OptionStrikeSelectionModeIncrement
    OptionStrikeSelectionModeExpenditure
    OptionStrikeSelectionModeDelta
End Enum

Public Enum OptionStrikeSelectionOperators
    OptionStrikeSelectionOperatorNone
    OptionStrikeSelectionOperatorLT
    OptionStrikeSelectionOperatorLE
    OptionStrikeSelectionOperatorGT
    OptionStrikeSelectionOperatorGE
End Enum

Public Enum OrderAttributes
    OrderAttOpenClose = 1
    OrderAttOrigin = 2
    OrderAttOriginatorRef = 4
    OrderAttBlockOrder = 8
    OrderAttSweepToFill = &H10&
    OrderAttDisplaySize = &H20&
    OrderAttIgnoreRTH = &H40&
    OrderAttHidden = &H80&
    OrderAttDiscretionaryAmount = &H100&
    OrderAttGoodAfterTime = &H200&
    OrderAttGoodTillDate = &H400&
    'OrderAttRTHOnly = &H800&
    OrderAttRule80A = &H1000&
    OrderAttSettlingFirm = &H2000&
    OrderAttAllOrNone = &H4000&
    OrderAttMinimumQuantity = &H8000&
    OrderAttPercentOffset = &H10000
    'OrderAttETradeOnly = &H20000
    'OrderAttFirmQuoteOnly = &H40000
    'OrderAttNBBOPriceCap = &H80000
    OrderAttOverrideConstraints = &H100000
    OrderAttAction = &H200000
    OrderAttLimitPrice = &H400000
    OrderAttOrderType = &H800000
    OrderAttQuantity = &H1000000
    OrderAttTimeInForce = &H2000000
    OrderAttTriggerPrice = &H4000000
    OrderAttGoodAfterTimeTZ = &H8000000
    OrderAttGoodTillDateTZ = &H10000000
    OrderAttStopTriggerMethod = &H20000000
End Enum

Public Enum OrderOrigins
    OriginCustomer = 0
    OriginFirm = 1
End Enum

''
'
'
' @param OrderStatusCreated
'   The order has been created but not yet transmitted to the broker
'
' @param OrderStatusRejected
'   The order has been transmitted to the broker but has been rejected
'   by the broker for some reason.
'
' @param OrderStatusPendingSubmit
'   The order has been transmitted to the broker, but no confirmation has
'   yet been received that it has been accepted.
'
' @param OrderStatusPreSubmitted
'   The order has been accepted by the broker, but the conditions for it
'   to be activated have not yet been met.
'
' @param OrderStatusSubmitted
'   The order has been activated at the order destination.
'
' @param OrderStatusFilled
'   The order has been completely filled.
'
' @param OrderStatusCancelling
'   A request to cancel the order has been sent, but confirmation of
'   the cancel has not yet been received.
'
' @param OrderStatusCancelled
'   The order has been cancelled.
'
'@/
'
Public Enum OrderStatuses
    OrderStatusCreated = 0
    OrderStatusRejected = 1
    OrderStatusPendingSubmit = 2
    OrderStatusPreSubmitted = 3
    OrderStatusSubmitted = 4
    OrderStatusFilled = 5
    OrderStatusCancelling = 6
    OrderStatusCancelled = 7
End Enum

Public Enum OrderContextChangeTypes
    OrderContextReadyStateChanged = 1
    OrderContextActiveStateChanged
    OrderContextFinished
End Enum

Public Enum OrderStopTriggerMethods
    OrderStopTriggerNone = 0
    OrderStopTriggerDefault = 1
    OrderStopTriggerDoubleBidAsk = 2
    OrderStopTriggerLast = 4
    OrderStopTriggerDoubleLast = 8
    OrderStopTriggerBidAsk = 16
    OrderStopTriggerLastOrBidAsk = 32
    OrderStopTriggerMidPoint = 64
End Enum

Public Enum OrderSubmitterCapabilities
    OrderSubmitterCapabilityNone = 0
    OrderSubmitterCapabilityCanStageOrders = 1
End Enum

Public Enum OrderSubmitterStates
    OrderSubmitterStateReady = 1
    OrderSubmitterStateNotReady
End Enum

Public Enum OrderTIFs
    OrderTIFNone = 0
    OrderTIFDay = 1
    OrderTIFGoodTillCancelled = 2
    OrderTIFImmediateOrCancel = 4
End Enum

Public Enum OrderTypes
    OrderTypeNone = 0
    OrderTypeMarket = 1
    OrderTypeMarketIfTouched = 2
    OrderTypeMarketOnClose = 4
    OrderTypeMarketOnOpen = 8
    OrderTypeMarketToLimit = &H10&
    OrderTypeLimit = &H20&
    OrderTypeLimitIfTouched = &H40&
    OrderTypeLimitOnClose = &H80&
    OrderTypeLimitOnOpen = &H100&
    OrderTypeStop = &H200&
    OrderTypeStopLimit = &H400&
    OrderTypeTrail = &H8000&
    OrderTypeTrailLimit = &H10000
    OrderTypePeggedToMarket = &H20000
    OrderTypeMarketWithProtection = &H40000
    OrderTypePeggedToPrimary = &H80000
    OrderTypeRelative = &H100000
End Enum

Public Enum PositionManagerChangeTypes
    PositionSizeChanged
    ProviderReadinessChanged
    PositionClosed
    PositionCloseFailed
End Enum

Public Enum PositionTypes
    PositionTypeLive = 1
    PositionTypeSimulated = 2
    PositionTypeAll = 3
End Enum

Public Enum PriceOffsetTypes
    PriceOffsetTypeNone = 0
    PriceOffsetTypeIncrement
    PriceOffsetTypeNumberOfTicks
    PriceOffsetTypeBidAskPercent
    PriceOffsetTypePercent
End Enum

Public Enum PriceValueTypes
    PriceValueTypeNone = 0
    PriceValueTypeValue
    PriceValueTypeAsk
    PriceValueTypeBid
    PriceValueTypeLast
    PriceValueTypeEntry
    PriceValueTypeMid
    PriceValueTypeBidOrAsk
    PriceValueTypeModel
End Enum

Public Enum ProfitChangeTypes
    ProfitChangeProfitThresholdExceeded
    ProfitChangeLossThresholdExceeded
    ProfitChangeDrawdownThresholdExceeded
End Enum

Public Enum ProfitTypes
    ProfitTypeProfit = 1&
    ProfitTypeMaxProfit = 2&
    ProfitTypeMaxLoss = 4&
    ProfitTypeDrawdown = 8&
    ProfitTypeSessionProfit = &H10&
    ProfitTypeSessionMaxProfit = &H20&
    ProfitTypeSessionMaxLoss = &H40&
    ProfitTypeSessionDrawdown = &H80&
    ProfitTypeTradeProfit = &H100&
    ProfitTypeTradeMaxProfit = &H200&
    ProfitTypeTradeMaxLoss = &H400&
    ProfitTypeTradeDrawdown = &H800&
End Enum

Public Enum RolloverControllerStates
    RolloverControllerStateInitial = 0
    RolloverControllerStateClosingCurrentPosition
    RolloverControllerStateFindingNextContract
    RolloverControllerStateCreatingNextPositionManager
    RolloverControllerStateWaitingForOrderContext
    RolloverControllerStateCreatingNextBracketOrder
    RolloverControllerStateExecutingNextBracketOrder
    RolloverControllerStateCompleted
    RolloverControllerStateCancellingNextBracketOrder
    RolloverControllerStateFailed
End Enum

Public Enum RolloverQuantityModes
'    - number                                       45
'    - previous number                              <
'    - currency amount                              1000$
'    - previous currency amount                     <$
'    - percentage of funds                          5.5%
'    - previous percentage of funds                 <%
'    - current value                                =$
'    - current value less percentage of profits     =45%P$
    RolloverQuantityModeNone
    RolloverQuantityModeAsPrevious
    RolloverQuantityModeNumber
    RolloverQuantityModePreviousNumber
    RolloverQuantityModeMonetaryAmount
    RolloverQuantityModePreviousMonetaryAmount
    RolloverQuantityModePercentageOfAccount
    RolloverQuantityModePreviousPercentageOfAccount
    RolloverQuantityModeCurrentValue
    RolloverQuantityModeCurrentProfit

End Enum

Public Enum RolloverStrikeModes
    RolloverStrikeModeNone
    RolloverStrikeModeAsPrevious
    RolloverStrikeModeCurrentValue
    RolloverStrikeModeCurrentProfit
    RolloverStrikeModeMonetaryAmount
    RolloverStrikeModePreviousMonetaryAmount
    RolloverStrikeModeIncrement
    RolloverStrikeModeDelta
    RolloverStrikeModePreviousDelta
    RolloverStrikeModeDeltaIncrement
End Enum
    
Public Enum TickRoundingModes
    TickRoundingModeNearest
    TickRoundingModeDown
    TickRoundingModeUp
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "OrderUtilsEnums"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




