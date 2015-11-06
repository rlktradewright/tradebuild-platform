VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderUtils"
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

Public Enum BracketEntryTypes
    BracketEntryTypeNone = 0
    BracketEntryTypeMarket = 1
    BracketEntryTypeMarketOnOpen = &H100000
    BracketEntryTypeMarketOnClose = 2
    BracketEntryTypeMarketIfTouched = &H20000
    BracketEntryTypeMarketToLimit = &H400&
    BracketEntryTypeBid = &H1000000
    BracketEntryTypeAsk = &H2000000
    BracketEntryTypeLast = &H4000000
    BracketEntryTypeLimit = 4
    BracketEntryTypeLimitOnOpen = &H200000
    BracketEntryTypeLimitOnClose = 8
    BracketEntryTypeLimitIfTouched = &H10000
    BracketEntryTypeStop = &H20&
    BracketEntryTypeStopLimit = &H40&
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
End Enum

Public Enum BracketOrderRoles
    BracketOrderRoleNone = 0
    BracketOrderRoleEntry = 1
    BracketOrderRoleStopLoss
    BracketOrderRoleTarget
    BracketOrderRoleCloseout
End Enum

Public Enum BracketOrderStates
    BracketOrderStateCreated = 1   ' indicates that the bracket order has been created but none of the orders have yet been placed
    BracketOrderStateSubmitted     ' indicates that all the orders in the bracket order have been placed
    BracketOrderStateCancelling    ' indicates that the bracket order is in the process of being cancelled
    BracketOrderStateClosingOut    ' indicates that the bracket order is being closed out. (When a bracket order has been closed out, there are no outstanding orders and no net Size - note that this may require a closeout order to be placed to negate any existing Size
    BracketOrderStateClosed        ' indicates that all orders have been completed
    BracketOrderStateAwaitingOtherOrderCancel  ' indicates that a cancellation notification for either the stop-loss order or the target order, when both exist, has been received
End Enum

Public Enum BracketStopLossTypes
    BracketStopLossTypeNone = 0
    BracketStopLossTypeStop = &H20&
    BracketStopLossTypeStopLimit = &H40&
    BracketStopLossTypeBid = &H1000000
    BracketStopLossTypeAsk = &H2000000
    BracketStopLossTypeLast = &H4000000
    BracketStopLossTypeAuto = &H8000000
End Enum

Public Enum BracketTargetTypes
    BracketTargetTypeNone = 0
    BracketTargetTypeLimit = 4
    BracketTargetTypeLimitIfTouched = &H10000
    BracketTargetTypeMarketIfTouched = &H20000
    BracketTargetTypeBid = &H1000000
    BracketTargetTypeAsk = &H2000000
    BracketTargetTypeLast = &H4000000
    BracketTargetTypeAuto = &H8000000
End Enum

Public Enum ClosePositionsFlags
    ClosePositionsCancelOrders = 1
    ClosePositionsWaitForCancel = 2
End Enum

Public Enum OrderActions
    OrderActionNone = 0
    OrderActionBuy = 1
    OrderActionSell = 2
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
    OrderAttETradeOnly = &H20000
    OrderAttFirmQuoteOnly = &H40000
    OrderAttNBBOPriceCap = &H80000
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
    OrderTypeMarketOnClose = 2
    OrderTypeLimit = 4
    OrderTypeLimitOnClose = 8
    OrderTypePeggedToMarket = &H10&
    OrderTypeStop = &H20&
    OrderTypeStopLimit = &H40&
    OrderTypeTrail = &H80&
    OrderTypeRelative = &H100&
    OrderTypeVWAP = &H200&
    OrderTypeMarketToLimit = &H400&
    OrderTypeQuote = &H800&
    OrderTypeAutoStop = &H1000&
    OrderTypeAutoLimit = &H2000&
    OrderTypeAdjust = &H4000&
    OrderTypeAlert = &H8000&
    OrderTypeLimitIfTouched = &H10000
    OrderTypeMarketIfTouched = &H20000
    OrderTypeTrailLimit = &H40000
    OrderTypeMarketWithProtection = &H80000
    OrderTypeMarketOnOpen = &H100000
    OrderTypeLimitOnOpen = &H200000
    OrderTypePeggedToPrimary = &H400000
End Enum

Public Enum PositionManagerChangeTypes
    PositionSizeChanged
    ProviderReadinessChanged
    PositionClosed
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

'@================================================================================
' Types
'@================================================================================

Public Type BracketOrderProfile
    Key                 As String
    Action              As OrderActions
    Quantity            As Long
    EntryPrice          As Double
    ExitPrice           As Double
    Profit              As Currency
    MaxProfit           As Currency
    MaxLoss             As Currency
    Risk                As Currency
    StartTime           As Date
    EndTime             As Date
    Description         As String
    EntryReason         As String
    StopReason          As String
    TargetReason        As String
End Type

Public Type BracketOrderErrorEventData
    Source              As Object
    AffectedOrder       As IOrder
    ErrorCode           As Long
    ErrorMsg            As String
End Type

Public Type ExecutionEventData
    Source              As Object
    ExecutionReport     As IExecutionReport
End Type

Public Type ProfitEventData
    Source          As Object
    ProfitTypes     As ProfitTypes
    CurrencyCode    As String
End Type

Public Type RecoveryDataChangedEventData
    Source              As Object
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "OrderUtils"

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

Public Function BracketEntryTypeToOrderType( _
                ByVal Value As BracketEntryTypes) As OrderTypes
Const ProcName As String = "BracketEntryTypeToOrderType"
On Error GoTo Err

BracketEntryTypeToOrderType = gBracketEntryTypeToOrderType(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BracketEntryTypeToString( _
                ByVal Value As BracketEntryTypes) As String
Const ProcName As String = "BracketEntryTypeToString"
On Error GoTo Err

BracketEntryTypeToString = gBracketEntryTypeToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BracketEntryTypeToShortString( _
                ByVal Value As BracketEntryTypes) As String
Const ProcName As String = "BracketEntryTypeToShortString"
On Error GoTo Err

BracketEntryTypeToShortString = gBracketEntryTypeToShortString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BracketOrderRoleToString(ByVal pOrderRole As BracketOrderRoles) As String
Const ProcName As String = "BracketOrderRoleToString"
On Error GoTo Err

BracketOrderRoleToString = gBracketOrderRoleToString(pOrderRole)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a member of the BracketStopLossTypes enumeration to the equivalent OrderTypes Value.
'
' @return           the OrderTypes Value corresponding to the parameter
' @param Value the BracketStopLossTypes Value to be converted
' @see
'
'@/
Public Function BracketStopLossTypeToOrderType( _
                ByVal Value As BracketStopLossTypes) As OrderTypes
Const ProcName As String = "BracketStopLossTypeToOrderType"
On Error GoTo Err

BracketStopLossTypeToOrderType = gBracketStopLossTypeToOrderType(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BracketStopLossTypeToString( _
                ByVal Value As BracketStopLossTypes)
Const ProcName As String = "BracketStopLossTypeToString"
On Error GoTo Err

BracketStopLossTypeToString = gBracketStopLossTypeToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a member of the <code>BracketTargetTypes</code> enumeration to the equivalent <code>OrderTypes</code> Value.
'
' @return           the <code>OrderTypes</code> Value corresponding to the parameter
' @param Value the <code>BracketTargetTypes</code> Value to be converted
' @see
'
'@/
Public Function BracketTargetTypeToOrderType( _
                ByVal Value As BracketTargetTypes) As OrderTypes
Const ProcName As String = "BracketTargetTypeToOrderType"
On Error GoTo Err

BracketTargetTypeToOrderType = gBracketTargetTypeToOrderType(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BracketTargetTypeToString( _
                ByVal Value As BracketTargetTypes) As String
Const ProcName As String = "BracketTargetTypeToString"
On Error GoTo Err

BracketTargetTypeToString = gBracketTargetTypeToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBracketProfitCalculator( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pDataSource As IMarketDataSource) As BracketProfitCalculator
Const ProcName As String = "CreateBracketProfitCalculator"
On Error GoTo Err

Set CreateBracketProfitCalculator = gCreateBracketProfitCalculator(pBracketOrder, pDataSource)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateOrderPersistenceDataStore(ByVal pRecoveryFilePath As String) As IOrderPersistenceDataStore
Const ProcName As String = "CreateOrderPersistenceDataStore"
On Error GoTo Err

Dim lDataStore As OrderPersistenceDataStore
Set lDataStore = New OrderPersistenceDataStore
lDataStore.Initialise pRecoveryFilePath
Set CreateOrderPersistenceDataStore = lDataStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderActionFromString(ByVal Value As String) As OrderActions
Const ProcName As String = "OrderActionFromString"
On Error GoTo Err

OrderActionFromString = gOrderActionFromString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderActionToString(ByVal Value As OrderActions) As String
Const ProcName As String = "OrderActionToString"
On Error GoTo Err

OrderActionToString = gOrderActionToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTIFFromString(ByVal Value As String) As OrderTIFs
Const ProcName As String = "OrderTIFFromString"
On Error GoTo Err

OrderTIFFromString = gOrderTIFFromString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTIFToString(ByVal Value As OrderTIFs) As String
Const ProcName As String = "OrderTIFToString"
On Error GoTo Err

OrderTIFToString = gOrderTIFToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTypeFromString(ByVal Value As String) As OrderTypes
Const ProcName As String = "OrderTypeFromString"
On Error GoTo Err

OrderTypeFromString = gOrderTypeFromString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTypeToString(ByVal Value As OrderTypes) As String
Const ProcName As String = "OrderTypeToString"
On Error GoTo Err

OrderTypeToString = gOrderTypeToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTypeToShortString(ByVal Value As OrderTypes) As String
Const ProcName As String = "OrderTypeToShortString"
On Error GoTo Err

OrderTypeToShortString = gOrderTypeToShortString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderStatusToString(ByVal pOrderStatus As OrderStatuses) As String
Const ProcName As String = "OrderStatusToString"
On Error GoTo Err

OrderStatusToString = gOrderStatusToString(pOrderStatus)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderStopTriggerMethodToString(ByVal Value As OrderStopTriggerMethods) As String
Const ProcName As String = "gOrderStopTriggerMethodToString"
On Error GoTo Err

OrderStopTriggerMethodToString = gOrderStopTriggerMethodToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




