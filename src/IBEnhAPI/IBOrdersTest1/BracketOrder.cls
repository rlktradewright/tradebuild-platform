VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements OrderUtils27.IBracketOrder

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "BracketOrder"

'@================================================================================
' Member variables
'@================================================================================

Private mBrokerData                                 As Variant
Private mCloseoutOrder                              As IOrder
Private mContract                                   As IContract
Private mEntryOrder                                 As IOrder
Private mKey                                        As String
Private mStopLossOrder                              As IOrder
Private mTargetOrder                                As IOrder
Private mCreationTime                               As Date



'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mKey = GenerateGUIDString
mCreationTime = Now
End Sub

'@================================================================================
' IBracketOrder Interface Members
'@================================================================================

Private Sub IBracketOrder_AddBracketOrderErrorListener(ByVal pListener As IBracketOrderErrorListener)
Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "AddBracketOrderErrorListener not supported"
End Sub

Private Sub IBracketOrder_AddBracketOrderMessageListener(ByVal pListener As OrderUtils27.IBracketOrderMsgListener)
Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "AddBracketOrderMessageListener not supported"
End Sub

Private Sub IBracketOrder_AdjustStop(ByVal pStopTriggerPriceSpec As PriceSpecifier, Optional ByVal pQuantity As Long = -1&)
End Sub

Private Sub IBracketOrder_AdjustTarget(ByVal pTargetPriceSpec As PriceSpecifier, Optional ByVal pQuantity As Long = -1&)
End Sub

Private Property Get IBracketOrder_ApplicationIndex() As Long
IBracketOrder_ApplicationIndex = gNextAplicationIndex
End Property

Private Property Let IBracketOrder_CancelAfter(ByVal Value As Long)
End Property

Private Property Let IBracketOrder_CancelPrice(ByVal Value As Double)
End Property

Private Property Get IBracketOrder_CancelPrice() As Double
End Property

Private Property Let IBracketOrder_CancelTime(ByVal Value As Date)
End Property

Private Property Get IBracketOrder_CancelTime() As Date
End Property

Private Property Get IBracketOrder_CloseoutOrder() As IOrder
Set IBracketOrder_CloseoutOrder = CloseoutOrder
End Property

Private Property Get IBracketOrder_ContextName() As String
End Property

Private Property Get IBracketOrder_Contract() As IContract
Set IBracketOrder_Contract = Contract
End Property

Private Property Get IBracketOrder_CreationTime() As Date
IBracketOrder_CreationTime = mCreationTime
End Property

Private Property Get IBracketOrder_CumBuyPrice() As Double
End Property

Private Property Get IBracketOrder_CumSellPrice() As Double
End Property

Private Property Let IBracketOrder_Description(ByVal Value As String)
End Property

Private Property Get IBracketOrder_Description() As String
End Property

Private Property Get IBracketOrder_EndTime() As Date
End Property

Private Property Get IBracketOrder_EntryOrder() As IOrder
Set IBracketOrder_EntryOrder = EntryOrder
End Property

Private Property Get IBracketOrder_EntryPrice() As Double
End Property

Private Property Let IBracketOrder_EntryReason(ByVal Value As String)
End Property

Private Property Get IBracketOrder_EntryReason() As String
End Property

Private Property Get IBracketOrder_ExitPrice() As Double
End Property

Private Property Get IBracketOrder_GroupName() As String
End Property

Private Property Get IBracketOrder_IsComplete() As Boolean
End Property

Private Property Get IBracketOrder_IsDirty() As Boolean
End Property

Private Property Get IBracketOrder_IsSimulated() As Boolean
End Property

Private Property Get IBracketOrder_Key() As String
IBracketOrder_Key = Key
End Property

Private Property Get IBracketOrder_LongPosition() As Boolean
End Property

Private Sub IBracketOrder_NotifyCloseout(ByVal pExec As IExecutionReport)
End Sub

Private Property Get IBracketOrder_Order(ByVal pIndex As Long) As IOrder

End Property

Private Property Get IBracketOrder_OrderContext() As OrderContext

End Property

Private Property Get IBracketOrder_PendingSize() As Long
End Property

Private Property Let IBracketOrder_PreventUnprotectedPosition(ByVal Value As Boolean)
End Property

Private Property Get IBracketOrder_PreventUnprotectedPosition() As Boolean
End Property

Private Sub IBracketOrder_RemoveBracketOrderErrorListener(ByVal Value As IBracketOrderErrorListener)
End Sub

Private Sub IBracketOrder_RemoveBracketOrderMessageListener(ByVal pListener As IBracketOrderMsgListener)
End Sub

Private Property Get IBracketOrder_Risk() As Currency
End Property

Private Property Let IBracketOrder_RolloverSpecification(ByVal Value As RolloverSpecification)
End Property

Private Property Get IBracketOrder_RolloverSpecification() As RolloverSpecification
End Property

Private Sub IBracketOrder_SetNewEntryLimitPrice(ByVal pPriceSpec As PriceSpecifier)
End Sub

Private Sub IBracketOrder_SetNewEntryQuantity(ByVal pNewQuantity As Long)
End Sub

Private Sub IBracketOrder_SetNewEntryTriggerPrice(ByVal pPriceSpec As PriceSpecifier)
End Sub

Private Sub IBracketOrder_SetNewOrderLimitPrice(ByVal pRole As BracketOrderRoles, ByVal pPriceSpec As PriceSpecifier)
End Sub

Private Sub IBracketOrder_SetNewOrderTriggerPrice(ByVal pRole As BracketOrderRoles, ByVal pPriceSpec As PriceSpecifier)
End Sub

Private Sub IBracketOrder_SetNewStopLossLimitPrice(ByVal pPriceSpec As PriceSpecifier)
End Sub

Private Sub IBracketOrder_SetNewStopLossTriggerPrice(ByVal pPriceSpec As PriceSpecifier)
End Sub

Private Sub IBracketOrder_SetNewTargetLimitPrice(ByVal pPriceSpec As PriceSpecifier)
End Sub

Private Sub IBracketOrder_SetNewTargetTriggerPrice(ByVal pPriceSpec As PriceSpecifier)
End Sub

Private Property Get IBracketOrder_Size() As Long
End Property

Private Property Get IBracketOrder_StartTime() As Date
End Property

Private Property Get IBracketOrder_State() As BracketOrderStates
End Property

Private Property Get IBracketOrder_StopLossOrder() As IOrder
Set IBracketOrder_StopLossOrder = StopLossOrder
End Property

Private Property Let IBracketOrder_StopReason(ByVal Value As String)
End Property

Private Property Get IBracketOrder_StopReason() As String
End Property

Private Property Get IBracketOrder_TargetOrder() As IOrder
Set IBracketOrder_TargetOrder = TargetOrder
End Property

Private Property Let IBracketOrder_TargetReason(ByVal Value As String)
End Property

Private Property Get IBracketOrder_TargetReason() As String
End Property

Private Sub IBracketOrder_AddChangeListener(ByVal Value As IChangeListener)
End Sub

Private Sub IBracketOrder_AddExecutionListener(ByVal Value As IExecutionListener)
End Sub

Private Sub IBracketOrder_AddStateChangeListener(ByVal Value As IStateChangeListener)
End Sub

Private Function IBracketOrder_Cancel(Optional ByVal pEvenIfFilled As Boolean = False) As Boolean
End Function

Private Sub IBracketOrder_CancelChanges()
End Sub

Private Sub IBracketOrder_Checkpoint()
End Sub

Private Sub IBracketOrder_CloseOut()
End Sub

Private Function IBracketOrder_ContainsOrder(ByVal pOrder As IOrder) As Boolean
End Function

Private Sub IBracketOrder_Execute()
End Sub

Private Sub IBracketOrder_RemoveChangeListener(ByVal pListener As IChangeListener)
End Sub

Private Sub IBracketOrder_RemoveExecutionListener(ByVal pListener As IExecutionListener)
End Sub

Private Sub IBracketOrder_RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
End Sub

Private Sub IBracketOrder_SetNewOrderPrice(ByVal pOrderIndex As Long, ByVal Value As Double)
End Sub

Private Sub IBracketOrder_SetNewQuantity(ByVal Value As Long)
End Sub

Private Sub IBracketOrder_SetNewStopLossOffset(ByVal Value As Long)
End Sub

Private Sub IBracketOrder_SetNewStopLossPrice(ByVal Value As Double)
End Sub

Private Sub IBracketOrder_SetNewStopLossQuantity(ByVal Value As Long)
End Sub

Private Sub IBracketOrder_SetNewTargetPrice(ByVal Value As Double)
End Sub

Private Sub IBracketOrder_SetNewTargetOffset(ByVal Value As Long)
End Sub

Private Sub IBracketOrder_SetNewTargetQuantity(ByVal Value As Long)
End Sub

Private Function IBracketOrder_ToString() As String
End Function

Private Sub IBracketOrder_Update()
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let CloseoutOrder(ByVal Value As IOrder)
Set mCloseoutOrder = Value
End Property

Public Property Get CloseoutOrder() As IOrder
Set CloseoutOrder = mCloseoutOrder
End Property

Public Property Let Contract(ByVal Value As IContract)
Set mContract = Value
End Property

Public Property Get Contract() As IContract
Set Contract = mContract
End Property

Public Property Let EntryOrder(ByVal Value As IOrder)
Set mEntryOrder = Value
End Property

Public Property Get EntryOrder() As IOrder
Set EntryOrder = mEntryOrder
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Let StopLossOrder(ByVal Value As IOrder)
Set mStopLossOrder = Value
End Property

Public Property Get StopLossOrder() As IOrder
Set StopLossOrder = mStopLossOrder
End Property

Public Property Let TargetOrder(ByVal Value As IOrder)
Set mTargetOrder = Value
End Property

Public Property Get TargetOrder() As IOrder
Set TargetOrder = mTargetOrder
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

