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
EntryOrderTypes = gEntryOrderTypes
End Property

Public Property Get StopLossOrderTypes() As OrderTypes()
StopLossOrderTypes = gStopLossOrderTypes
End Property

Public Property Get TargetOrderTypes() As OrderTypes()
TargetOrderTypes = gTargetOrderTypes
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function BracketOrderRoleToString(ByVal pOrderRole As BracketOrderRoles) As String
Const ProcName As String = "BracketOrderRoleToString"
On Error GoTo Err

BracketOrderRoleToString = gBracketOrderRoleToString(pOrderRole)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BracketOrderStateToString(ByVal pState As BracketOrderStates) As String
Const ProcName As String = "BracketOrderStateToString"
On Error GoTo Err

BracketOrderStateToString = gBracketOrderStateToString(pState)

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

Public Function IsEntryOrderType(ByVal pOrderType As OrderTypes) As Boolean
IsEntryOrderType = gIsEntryOrderType(pOrderType)
End Function

Public Function IsStopLossOrderType(ByVal pOrderType As OrderTypes) As Boolean
IsStopLossOrderType = gIsStopLossOrderType(pOrderType)
End Function

Public Function IsTargetOrderType(ByVal pOrderType As OrderTypes) As Boolean
IsTargetOrderType = gIsTargetOrderType(pOrderType)
End Function

Public Function NewPriceSpecifier( _
                Optional ByVal pPrice As Double = MaxDoubleValue, _
                Optional ByVal pPriceType As PriceValueTypes = PriceValueTypeNone, _
                Optional ByVal pOffset As Double = 0#, _
                Optional ByVal pOffsetType As PriceOffsetTypes = PriceOffsetTypeNone) As PriceSpecifier
Set NewPriceSpecifier = gNewPriceSpecifier(pPrice, pPriceType, pOffset, pOffsetType)
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

Public Function CreateRolloverSpecification( _
                ByVal pDays As Long, _
                ByVal pTime As Date, _
                ByVal pCloseOrderType As OrderTypes, _
                ByVal pCloseLimitPriceSpec As PriceSpecifier, _
                ByVal pCloseTriggerPriceSpec As PriceSpecifier, _
                ByVal pEntryOrderType As OrderTypes, _
                ByVal pEntryLimitPriceSpec As PriceSpecifier, _
                ByVal pEntryTriggerPriceSpec As PriceSpecifier) As RolloverSpecification
Const ProcName As String = "CreateRolloverSpecification"
On Error GoTo Err

Set CreateRolloverSpecification = gCreateRolloverSpecification(pDays, pTime, pCloseOrderType, pCloseLimitPriceSpec, pCloseTriggerPriceSpec, pEntryOrderType, pEntryLimitPriceSpec, pEntryTriggerPriceSpec)

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

Public Function PriceOffsetToString( _
                ByVal pOffset As Double, _
                ByVal pOffsetType As PriceOffsetTypes)
Const ProcName As String = "PriceOffsetToString"
On Error GoTo Err

PriceOffsetToString = gPriceOffsetToString(pOffset, pOffsetType)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function PriceOffsetTypeToString( _
                ByVal pOffsetType As PriceOffsetTypes)
Const ProcName As String = "PriceOffsetTypeToString"
On Error GoTo Err

PriceOffsetTypeToString = gPriceOffsetTypeToString(pOffsetType)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function PriceSpecifierToString( _
                ByVal pPriceSpec As PriceSpecifier, _
                ByVal pContract As IContract)
Const ProcName As String = "PriceSpecifierToString"
On Error GoTo Err

PriceSpecifierToString = gPriceSpecifierToString(pPriceSpec, _
                            pContract)
                            

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function TypedPriceToString( _
                ByVal pPrice As Double, _
                ByVal pPriceType As PriceValueTypes, _
                ByVal pContract As IContract) As String
Const ProcName As String = "gTypedPriceToString"
On Error GoTo Err

TypedPriceToString = gTypedPriceToString(pPrice, pPriceType, pContract)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================





