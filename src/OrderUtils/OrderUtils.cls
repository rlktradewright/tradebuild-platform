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

'@================================================================================
' Types
'@================================================================================

Public Type BracketOrderErrorEventData
    Source              As Object
    AffectedOrder       As IOrder
    ErrorCode           As Long
    ErrorMsg            As String
End Type

Public Type BracketOrderMessageEventData
    Source              As Object
    AffectedOrder       As IOrder
    Message             As String
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

Public Type RecoveryDataCommentEventData
    Source              As Object
    Comment            As String
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

Public Property Get EntryOrderTypes() As OrderTypes()
EntryOrderTypes = GOrderUtils.EntryOrderTypes
End Property

Public Property Get StopLossOrderTypes() As OrderTypes()
StopLossOrderTypes = GOrderUtils.StopLossOrderTypes
End Property

Public Property Get TargetOrderTypes() As OrderTypes()
TargetOrderTypes = GOrderUtils.TargetOrderTypes
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function BracketOrderRoleToString(ByVal pOrderRole As BracketOrderRoles) As String
Const ProcName As String = "BracketOrderRoleToString"
On Error GoTo Err

BracketOrderRoleToString = GOrderUtils.BracketOrderRoleToString(pOrderRole)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function BracketOrderStateToString(ByVal pState As BracketOrderStates) As String
Const ProcName As String = "BracketOrderStateToString"
On Error GoTo Err

BracketOrderStateToString = GOrderUtils.BracketOrderStateToString(pState)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBracketProfitCalculator( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pDataSource As IMarketDataSource) As BracketProfitCalculator
Const ProcName As String = "CreateBracketProfitCalculator"
On Error GoTo Err

Set CreateBracketProfitCalculator = GOrderUtils.CreateBracketProfitCalculator(pBracketOrder, pDataSource)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateOrderPersistenceDataStore(ByVal pRecoveryFilePath As String) As IOrderPersistenceDataStore
Const ProcName As String = "CreateOrderPersistenceDataStore"
On Error GoTo Err

Set CreateOrderPersistenceDataStore = GOrderUtils.CreateOrderPersistenceDataStore(pRecoveryFilePath)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsEntryOrderType(ByVal pOrderType As OrderTypes) As Boolean
IsEntryOrderType = GOrderUtils.IsEntryOrderType(pOrderType)
End Function

Public Function IsNullPriceSpecifier(ByVal pPriceSpec As PriceSpecifier) As Boolean
IsNullPriceSpecifier = GOrderUtils.IsNullPriceSpecifier(pPriceSpec)
End Function

Public Function IsStopLossOrderType(ByVal pOrderType As OrderTypes) As Boolean
IsStopLossOrderType = GOrderUtils.IsStopLossOrderType(pOrderType)
End Function

Public Function IsTargetOrderType(ByVal pOrderType As OrderTypes) As Boolean
IsTargetOrderType = GOrderUtils.IsTargetOrderType(pOrderType)
End Function

Public Function NewPriceSpecifier( _
                Optional ByVal pPrice As Double = GOrderUtils.MaxDoubleValue, _
                Optional ByVal pPriceString As String = "", _
                Optional ByVal pPriceType As PriceValueTypes = PriceValueTypeNone, _
                Optional ByVal pOffset As Double = 0#, _
                Optional ByVal pOffsetType As PriceOffsetTypes = PriceOffsetTypeNone) As PriceSpecifier
Set NewPriceSpecifier = GOrderUtils.NewPriceSpecifier(pPrice, pPriceType, pOffset, pOffsetType)
End Function

Public Function OptionStrikeSelectionModeFromString( _
                ByVal Value As String) As OptionStrikeSelectionModes
OptionStrikeSelectionModeFromString = GOrderUtils.OptionStrikeSelectionModeFromString(Value)
End Function

Public Function OptionStrikeSelectionModeToString( _
                ByVal Value As OptionStrikeSelectionModes) As String
OptionStrikeSelectionModeToString = GOrderUtils.OptionStrikeSelectionModeToString(Value)
End Function

Public Function OptionStrikeSelectionOperatorFromString(ByVal Value As String) As OptionStrikeSelectionOperators
OptionStrikeSelectionOperatorFromString = GOrderUtils.OptionStrikeSelectionOperatorFromString(Value)
End Function

Public Function OptionStrikeSelectionOperatorToString(ByVal Value As OptionStrikeSelectionOperators) As String
OptionStrikeSelectionOperatorToString = GOrderUtils.OptionStrikeSelectionOperatorToString(Value)
End Function

Public Function OrderActionFromString(ByVal Value As String) As OrderActions
Const ProcName As String = "OrderActionFromString"
On Error GoTo Err

OrderActionFromString = GOrderUtils.OrderActionFromString(Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderActionToString(ByVal Value As OrderActions) As String
Const ProcName As String = "OrderActionToString"
On Error GoTo Err

OrderActionToString = GOrderUtils.OrderActionToString(Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTIFFromString(ByVal Value As String) As OrderTIFs
Const ProcName As String = "OrderTIFFromString"
On Error GoTo Err

OrderTIFFromString = GOrderUtils.OrderTIFFromString(Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTIFToString(ByVal Value As OrderTIFs) As String
Const ProcName As String = "OrderTIFToString"
On Error GoTo Err

OrderTIFToString = GOrderUtils.OrderTIFToString(Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateOptionRolloverSpecification( _
                ByVal pDays As Long, _
                ByVal pTime As Date, _
                ByVal pInitialStrikeSelectionMode As OptionStrikeSelectionModes, _
                ByVal pInitialStrikeParameter As Double, _
                ByVal pInitialStrikeOperator As OptionStrikeSelectionOperators, _
                ByVal pRolloverStrikeSelectionMode As RolloverStrikeModes, _
                ByVal pRolloverStrikeValue As Double, _
                ByVal pRolloverStrikeOperator As OptionStrikeSelectionOperators, _
                ByVal pRolloverQuantityMode As RolloverQuantityModes, _
                ByVal pRolloverQuantityParameter As BoxedDecimal, _
                ByVal pRolloverQuantityLotSize As Long, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pCloseOrderType As OrderTypes, _
                ByVal pCloseTimeoutSecs As Long, _
                ByVal pCloseLimitPriceSpec As PriceSpecifier, _
                ByVal pCloseTriggerPriceSpec As PriceSpecifier, _
                ByVal pEntryOrderType As OrderTypes, _
                ByVal pEntryTimeoutSecs As Long, _
                ByVal pEntryLimitPriceSpec As PriceSpecifier, _
                ByVal pEntryTriggerPriceSpec As PriceSpecifier) As RolloverSpecification
Const ProcName As String = "CreateOptionRolloverSpecification"
On Error GoTo Err

Set CreateOptionRolloverSpecification = GOrderUtils.CreateOptionRolloverSpecification( _
                                                pDays, _
                                                pTime, _
                                                pInitialStrikeSelectionMode, _
                                                pInitialStrikeParameter, _
                                                pInitialStrikeOperator, _
                                                pRolloverStrikeSelectionMode, _
                                                pRolloverStrikeValue, _
                                                pRolloverStrikeOperator, _
                                                pRolloverQuantityMode, _
                                                pRolloverQuantityParameter, _
                                                pRolloverQuantityLotSize, _
                                                pUnderlyingExchangeName, _
                                                pCloseOrderType, _
                                                pCloseTimeoutSecs, _
                                                pCloseLimitPriceSpec, _
                                                pCloseTriggerPriceSpec, _
                                                pEntryOrderType, _
                                                pEntryTimeoutSecs, _
                                                pEntryLimitPriceSpec, _
                                                pEntryTriggerPriceSpec)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateRolloverSpecification( _
                ByVal pDays As Long, _
                ByVal pTime As Date, _
                ByVal pCloseOrderType As OrderTypes, _
                ByVal pCloseTimeoutSecs As Long, _
                ByVal pCloseLimitPriceSpec As PriceSpecifier, _
                ByVal pCloseTriggerPriceSpec As PriceSpecifier, _
                ByVal pEntryOrderType As OrderTypes, _
                ByVal pEntryTimeoutSecs As Long, _
                ByVal pEntryLimitPriceSpec As PriceSpecifier, _
                ByVal pEntryTriggerPriceSpec As PriceSpecifier) As RolloverSpecification
Const ProcName As String = "CreateRolloverSpecification"
On Error GoTo Err

Set CreateRolloverSpecification = GOrderUtils.CreateRolloverSpecification( _
                                                pDays, _
                                                pTime, _
                                                pCloseOrderType, _
                                                pCloseTimeoutSecs, _
                                                pCloseLimitPriceSpec, _
                                                pCloseTriggerPriceSpec, _
                                                pEntryOrderType, _
                                                pEntryTimeoutSecs, _
                                                pEntryLimitPriceSpec, _
                                                pEntryTriggerPriceSpec)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GenerateBracketOrderId() As String
GenerateBracketOrderId = GOrderUtils.GenerateBracketOrderId
End Function

Public Function GetOptionContract( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pAction As OrderActions, _
                ByVal pContractStore As IContractStore, _
                ByVal pSelectionMode As OptionStrikeSelectionModes, _
                ByVal pParameter As Long, _
                ByVal pOperator As OptionStrikeSelectionOperators, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pMarketDataManager As IMarketDataManager, _
                Optional ByVal pListener As IStateChangeListener, _
                Optional ByVal pReferenceDate As Date = GOrderUtils.MinDate) As IFuture
Const ProcName As String = "GetOptionContract"
On Error GoTo Err

Set GetOptionContract = GOrderUtils.GetOptionContract( _
                                pContractSpec, _
                                pAction, _
                                pContractStore, _
                                pSelectionMode, _
                                pParameter, _
                                pOperator, _
                                pUnderlyingExchangeName, _
                                pMarketDataManager, _
                                pListener, _
                                pReferenceDate)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTypeFromString(ByVal Value As String) As OrderTypes
Const ProcName As String = "OrderTypeFromString"
On Error GoTo Err

OrderTypeFromString = GOrderUtils.OrderTypeFromString(Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTypeToString(ByVal Value As OrderTypes) As String
Const ProcName As String = "OrderTypeToString"
On Error GoTo Err

OrderTypeToString = GOrderUtils.OrderTypeToString(Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderTypeToShortString(ByVal Value As OrderTypes) As String
Const ProcName As String = "OrderTypeToShortString"
On Error GoTo Err

OrderTypeToShortString = GOrderUtils.OrderTypeToShortString(Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderStatusToString(ByVal pOrderStatus As OrderStatuses) As String
Const ProcName As String = "OrderStatusToString"
On Error GoTo Err

OrderStatusToString = GOrderUtils.OrderStatusToString(pOrderStatus)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function OrderStopTriggerMethodToString(ByVal Value As OrderStopTriggerMethods) As String
Const ProcName As String = "OrderStopTriggerMethodToString"
On Error GoTo Err

OrderStopTriggerMethodToString = GOrderUtils.OrderStopTriggerMethodToString(Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function ParsePriceAndOffset( _
                ByRef pPriceSpec As PriceSpecifier, _
                ByVal pValue As String, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double, _
                ByRef pMessage As String, _
                Optional ByVal pUseCloseoutSemantics As Boolean = False) As Boolean
Const ProcName As String = "ParsePriceAndOffset"
On Error GoTo Err

ParsePriceAndOffset = GOrderUtils.ParsePriceAndOffset(pPriceSpec, pValue, pSecType, pTickSize, pMessage, pUseCloseoutSemantics)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function


Public Function PriceOffsetToString( _
                ByVal pOffset As Double, _
                ByVal pOffsetType As PriceOffsetTypes)
Const ProcName As String = "PriceOffsetToString"
On Error GoTo Err

PriceOffsetToString = GOrderUtils.PriceOffsetToString(pOffset, pOffsetType)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function PriceOffsetTypeToString( _
                ByVal pOffsetType As PriceOffsetTypes)
Const ProcName As String = "PriceOffsetTypeToString"
On Error GoTo Err

PriceOffsetTypeToString = GOrderUtils.PriceOffsetTypeToString(pOffsetType)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function PriceSpecifierToString( _
                ByVal pPriceSpec As PriceSpecifier, _
                ByVal pContract As IContract)
Const ProcName As String = "PriceSpecifierToString"
On Error GoTo Err

PriceSpecifierToString = GOrderUtils.PriceSpecifierToString(pPriceSpec, _
                            pContract)
                            

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function TypedPriceToString( _
                ByVal pPrice As Double, _
                ByVal pPriceType As PriceValueTypes, _
                ByVal pContract As IContract) As String
Const ProcName As String = "TypedPriceToString"
On Error GoTo Err

TypedPriceToString = GOrderUtils.TypedPriceToString(pPrice, pPriceType, pContract)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================





