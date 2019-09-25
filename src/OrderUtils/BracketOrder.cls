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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements IBracketOrder
Implements IRecoverable
Implements IStateChangeListener
Implements IStringable

'@================================================================================
' Events
'@================================================================================

Event Change(ByRef ev As ChangeEventData)
Event ChangesApplied()
Event ChangesCancelled()
Event ChangesPending()
Event CloseoutOrderChanged()
Event CloseoutOrderCreated()
Event CloseoutOrderFilled()
Event Completed()
Event EntryOrderChanged()
Event EntryOrderFilled()
Event Execution(ByRef ev As ExecutionEventData)
Event OrderError(ByRef ev As BracketOrderErrorEventData)
Event OrderMessage(ByRef ev As BracketOrderMessageEventData)
Event PositionCloseoutNotified()
Event SelfCancelled()
Event SizeChanged()
Event StateChange(ByRef ef As StateChangeEventData)
Event StopLossOrderChanged()
Event StopLossOrderFilled()
Event TargetOrderChanged()
Event TargetOrderFilled()

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "BracketOrder"

'@================================================================================
' Enums
'@================================================================================

'Public Enum BracketOrderStates
'
'    ' This state indicates that the bracket order has been created but none of
'    ' the orders have yet been placed.
'    BracketOrderStateCreated = 1
'
'    ' This state indicates that all the orders in the bracket order have been placed.
'    BracketOrderStateSubmitted
'
'    ' This state indicates that the bracket order is in the process of being
'    ' cancelled.
'    BracketOrderStateCancelling
'
'    ' This state indicates that the bracket order is being closed out.
'    ' (When an bracket order has been closed out, there are no outstanding
'    ' orders and no net Size - note that this may require a closeout
'    ' order to be placed to negate any existing Size.)
'    BracketOrderStateClosingOut
'
'    ' This state indicates that all orders have been completed.
'    BracketOrderStateClosed
'
'    ' This state indicates that a cancellation notification for either
'    ' the stop-loss order or the target order, when both exist, has been received.
'    ' We don't know whether the order was cancelled because the user/broker
'    ' cancelled it, or because the other order was filled. So in this
'    ' state we wait for the other order to be either cancelled or
'    ' filled or for a timeout to elapse.
'    BracketOrderStateAwaitingOtherOrderCancel
'
'End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSummary                                    As BOS

Private mRecoveryContextName                        As String

Private mOrderContext                               As OrderContext
Private mContract                                   As IContract
Private mDataSource                                 As IMarketDataSource
Private mOrderSubmitter                             As IOrderSubmitter
Private mTimerList                                  As TimerList

Private mStateEngine                                As StateEngine

Private mCancelAfter                                As Long
Private WithEvents mCancelTLI                       As TimerListItem
Attribute mCancelTLI.VB_VarHelpID = -1

Private WithEvents mEntryOrder                      As Order
Attribute mEntryOrder.VB_VarHelpID = -1
Private mEntryOrderCopy                             As Order

Private WithEvents mStopLossOrder                   As Order
Attribute mStopLossOrder.VB_VarHelpID = -1
Private mStopOrderCopy                              As Order
Attribute mStopOrderCopy.VB_VarHelpID = -1

Private WithEvents mTargetOrder                     As Order
Attribute mTargetOrder.VB_VarHelpID = -1
Private mTargetOrderCopy                            As Order
Attribute mTargetOrderCopy.VB_VarHelpID = -1

Private WithEvents mCloseoutOrder                   As Order
Attribute mCloseoutOrder.VB_VarHelpID = -1

Private mNewQuantity                                As Long
Private mNewEntryPrice                              As Double
Private mNewEntryTriggerPrice                       As Double
Private mNewStopLossPrice                           As Double
Private mNewStopLossOffset                          As Long
Private mNewStopLossTriggerPrice                    As Double
Private mNewStopLossQuantity                        As Long
Private mNewTargetPrice                             As Double
Private mNewTargetOffset                            As Long
Private mNewTargetTriggerPrice                      As Double
Private mNewTargetQuantity                          As Long

Private mIsSimulated                                As Boolean

Private mChangeListeners                            As New Listeners
Private mStateChangeListeners                       As New Listeners
Private mRecoverableObjectListeners                 As New Listeners
Private mExecutionListeners                         As New Listeners
Private mBracketOrderErrorListeners                 As New Listeners
Private mBracketOrderMessageListeners               As New Listeners

Private mApplicationIndex                           As Long

Private mIsDirty                                    As Boolean

Private mNumberOfOrders                             As Long

Private WithEvents mStateTimeoutTLI                 As TimerListItem
Attribute mStateTimeoutTLI.VB_VarHelpID = -1

Private mStartTime                                  As Date
Private mEndTime                                    As Date
Private mTimezoneName                               As String

Private WithEvents mRolloverTLI                     As TimerListItem
Attribute mRolloverTLI.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

mNewQuantity = -1
mNewEntryPrice = -1
mNewEntryTriggerPrice = -1
mNewStopLossPrice = -1
mNewStopLossOffset = DummyOffset
mNewStopLossTriggerPrice = -1
mNewStopLossQuantity = -1
mNewTargetPrice = -1
mNewTargetOffset = DummyOffset
mNewTargetTriggerPrice = -1
mNewTargetQuantity = -1

Set mSummary = New BOS

mApplicationIndex = GBracketOrder.gNextApplicationIndex

Set mStateEngine = CreateStateEngine(GBracketOrder.TableBuilder, _
                                    BracketOrderStates.BracketOrderStateCreated)

mSummary.State = State

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Class_Terminate()
Debug.Print "BracketOrder terminated"
End Sub

'@================================================================================
' IBracketOrder Interface Members
'@================================================================================

Private Sub IBracketOrder_AddBracketOrderErrorListener(ByVal pListener As IBracketOrderErrorListener)
AddBracketOrderErrorListener pListener
End Sub

Private Sub IBracketOrder_AddBracketOrderMessageListener(ByVal pListener As IBracketOrderMsgListener)
AddBracketOrderMessageListener pListener
End Sub

Private Sub IBracketOrder_AddChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "IBracketOrder_AddChangeListener"
On Error GoTo Err

AddChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_AddExecutionListener(ByVal pListener As IExecutionListener)
Const ProcName As String = "IBracketOrder_AddExecutionListener"
On Error GoTo Err

AddExecutionListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_AddStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "IBracketOrder_AddStateChangeListener"
On Error GoTo Err

AddStateChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_AdjustStop(ByVal pStopTriggerPrice As Double, Optional ByVal pQuantity As Long = -1&)
Const ProcName As String = "IBracketOrder_AdjustStop"
On Error GoTo Err

AdjustStopLoss pStopTriggerPrice, pQuantity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_AdjustTarget(ByVal pTargetPrice As Double, Optional ByVal pQuantity As Long = -1&)
Const ProcName As String = "IBracketOrder_AdjustTarget"
On Error GoTo Err

AdjustTarget pTargetPrice, pQuantity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IBracketOrder_ApplicationIndex() As Long
IBracketOrder_ApplicationIndex = ApplicationIndex
End Property

Private Function IBracketOrder_Cancel(Optional ByVal pEvenIfFilled As Boolean = False) As Boolean
Const ProcName As String = "IBracketOrder_Cancel"
On Error GoTo Err

IBracketOrder_Cancel = Cancel(pEvenIfFilled)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Let IBracketOrder_CancelAfter(ByVal Value As Long)
Const ProcName As String = "IBracketOrder_CancelAfter"
On Error GoTo Err

CancelAfter = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IBracketOrder_CancelChanges()
Const ProcName As String = "IBracketOrder_CancelChanges"
On Error GoTo Err

CancelChanges

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let IBracketOrder_CancelPrice(ByVal Value As Double)
Const ProcName As String = "IBracketOrder_CancelPrice"
On Error GoTo Err

CancelPrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IBracketOrder_CancelPrice() As Double
IBracketOrder_CancelPrice = CancelPrice
End Property

Private Property Let IBracketOrder_CancelTime(ByVal Value As Date)
Const ProcName As String = "IBracketOrder_CancelTime"
On Error GoTo Err

CancelTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IBracketOrder_CancelTime() As Date
IBracketOrder_CancelTime = CancelTime
End Property

Private Sub IBracketOrder_Checkpoint()
Const ProcName As String = "IBracketOrder_Checkpoint"
On Error GoTo Err

Checkpoint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_CloseOut()
Const ProcName As String = "IBracketOrder_CloseOut"
On Error GoTo Err

CloseOut

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IBracketOrder_CloseoutOrder() As IOrder
Set IBracketOrder_CloseoutOrder = mCloseoutOrder
End Property

Private Function IBracketOrder_ContainsOrder(ByVal pOrder As IOrder) As Boolean
Const ProcName As String = "IBracketOrder_ContainsOrder"
On Error GoTo Err

IBracketOrder_ContainsOrder = ContainsOrder(pOrder)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IBracketOrder_ContextName() As String
IBracketOrder_ContextName = mSummary.ContextName
End Property

Private Property Get IBracketOrder_Contract() As IContract
Set IBracketOrder_Contract = mContract
End Property

Private Property Get IBracketOrder_CreationTime() As Date
IBracketOrder_CreationTime = CreationTime
End Property

Private Property Get IBracketOrder_CumBuyPrice() As Double
IBracketOrder_CumBuyPrice = CumBuyPrice
End Property

Private Property Get IBracketOrder_CumSellPrice() As Double
IBracketOrder_CumSellPrice = CumSellPrice
End Property

Private Property Let IBracketOrder_Description(ByVal RHS As String)
Description = RHS
End Property

Private Property Get IBracketOrder_Description() As String
IBracketOrder_Description = Description
End Property

Private Property Get IBracketOrder_EndTime() As Date
IBracketOrder_EndTime = EndTime
End Property

Private Property Get IBracketOrder_EntryOrder() As IOrder
Set IBracketOrder_EntryOrder = mEntryOrder
End Property

Private Property Get IBracketOrder_EntryPrice() As Double
IBracketOrder_EntryPrice = EntryPrice
End Property

Private Property Let IBracketOrder_EntryReason(ByVal RHS As String)
EntryReason = RHS
End Property

Private Property Get IBracketOrder_EntryReason() As String
IBracketOrder_EntryReason = EntryReason
End Property

Private Sub IBracketOrder_Execute()
Const ProcName As String = "IBracketOrder_Execute"
On Error GoTo Err

Execute

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IBracketOrder_ExitPrice() As Double
IBracketOrder_ExitPrice = ExitPrice
End Property

Private Property Get IBracketOrder_GroupName() As String
IBracketOrder_GroupName = GroupName
End Property

Private Property Get IBracketOrder_IsComplete() As Boolean
Const ProcName As String = "IBracketOrder_IsComplete"
On Error GoTo Err

IBracketOrder_IsComplete = IsComplete

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IBracketOrder_IsDirty() As Boolean
IBracketOrder_IsDirty = IsDirty
End Property

Private Property Get IBracketOrder_IsSimulated() As Boolean
IBracketOrder_IsSimulated = IsSimulated
End Property

Private Property Get IBracketOrder_Key() As String
IBracketOrder_Key = mSummary.Key
End Property

Private Property Get IBracketOrder_LongPosition() As Boolean
IBracketOrder_LongPosition = LongPosition
End Property

Private Sub IBracketOrder_NotifyCloseout(ByVal pExec As IExecutionReport)
Const ProcName As String = "IBracketOrder_NotifyCloseout"
On Error GoTo Err

NotifyCloseout pExec

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IBracketOrder_Order(ByVal pRole As BracketOrderRoles) As IOrder
Const ProcName As String = "IBracketOrder_Order"
On Error GoTo Err

Set IBracketOrder_Order = Order(pRole)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IBracketOrder_OrderContext() As OrderContext
Set IBracketOrder_OrderContext = OrderContext
End Property

Private Property Get IBracketOrder_PendingSize() As Long
Const ProcName As String = "IBracketOrder_PendingSize"
On Error GoTo Err

IBracketOrder_PendingSize = PendingSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Let IBracketOrder_PreventUnprotectedPosition(ByVal Value As Boolean)
Const ProcName As String = "IBracketOrder_PreventUnprotectedPosition"
On Error GoTo Err

PreventUnprotectedPosition = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IBracketOrder_PreventUnprotectedPosition() As Boolean
IBracketOrder_PreventUnprotectedPosition = PreventUnprotectedPosition
End Property

Private Sub IBracketOrder_RemoveBracketOrderErrorListener(ByVal pListener As IBracketOrderErrorListener)
RemoveBracketOrderErrorListener pListener
End Sub

Private Sub IBracketOrder_RemoveBracketOrderMessageListener(ByVal pListener As IBracketOrderMsgListener)
RemoveBracketOrderMessageListener pListener
End Sub

Private Sub IBracketOrder_RemoveChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "IBracketOrder_RemoveChangeListener"
On Error GoTo Err

RemoveChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_RemoveExecutionListener(ByVal pListener As IExecutionListener)
Const ProcName As String = "IBracketOrder_RemoveExecutionListener"
On Error GoTo Err

RemoveExecutionListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "IBracketOrder_RemoveStateChangeListener"
On Error GoTo Err

RemoveStateChangeListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IBracketOrder_Risk() As Currency
IBracketOrder_Risk = Risk
End Property

Private Property Let IBracketOrder_RolloverSpecification(ByVal Value As RolloverSpecification)
RolloverSpecification = Value
End Property

Private Property Get IBracketOrder_RolloverSpecification() As RolloverSpecification
Set IBracketOrder_RolloverSpecification = RolloverSpecification
End Property

Private Sub IBracketOrder_SetNewEntryPrice(ByVal Value As Double)
Const ProcName As String = "IBracketOrder_SetNewEntryPrice"
On Error GoTo Err

SetNewEntryPrice Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewEntryTriggerPrice(ByVal Value As Double)
Const ProcName As String = "IBracketOrder_SetNewEntryTriggerPrice"
On Error GoTo Err

SetNewEntryTriggerPrice Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewOrderPrice(ByVal orderIndex As Long, ByVal Value As Double)
Const ProcName As String = "IBracketOrder_SetNewOrderPrice"
On Error GoTo Err

SetNewOrderPrice orderIndex, Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewOrderTriggerPrice(ByVal orderIndex As Long, ByVal Value As Double)
Const ProcName As String = "IBracketOrder_SetNewOrderTriggerPrice"
On Error GoTo Err

SetNewOrderTriggerPrice orderIndex, Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewQuantity(ByVal Value As Long)
Const ProcName As String = "IBracketOrder_SetNewQuantity"
On Error GoTo Err

SetNewQuantity Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewStopLossOffset(ByVal Value As Long)
Const ProcName As String = "IBracketOrder_SetNewStopLossOffset"
On Error GoTo Err

SetNewStopLossOffset Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewStopLossPrice(ByVal Value As Double)
Const ProcName As String = "IBracketOrder_SetNewStopLossPrice"
On Error GoTo Err

SetNewStopLossPrice Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewStopLossQuantity(ByVal Value As Long)
Const ProcName As String = "IBracketOrder_SetNewStopLossQuantity"
On Error GoTo Err

SetNewStopLossQuantity Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewStopLossTriggerPrice(ByVal Value As Double)
Const ProcName As String = "IBracketOrder_SetNewStopLossTriggerPrice"
On Error GoTo Err

SetNewStopLossTriggerPrice Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewTargetOffset(ByVal Value As Long)
Const ProcName As String = "IBracketOrder_SetNewTargetOffset"
On Error GoTo Err

SetNewTargetOffset Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewTargetPrice(ByVal Value As Double)
Const ProcName As String = "IBracketOrder_SetNewTargetPrice"
On Error GoTo Err

SetNewTargetPrice Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewTargetQuantity(ByVal Value As Long)
Const ProcName As String = "IBracketOrder_SetNewTargetQuantity"
On Error GoTo Err

SetNewTargetQuantity Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBracketOrder_SetNewTargetTriggerPrice(ByVal Value As Double)
Const ProcName As String = "IBracketOrder_SetNewTargetTriggerPrice"
On Error GoTo Err

SetNewTargetTriggerPrice Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IBracketOrder_Size() As Long
IBracketOrder_Size = Size
End Property

Private Property Get IBracketOrder_StartTime() As Date
IBracketOrder_StartTime = StartTime
End Property

Private Property Get IBracketOrder_State() As BracketOrderStates
IBracketOrder_State = State
End Property

Private Property Get IBracketOrder_StopLossOrder() As IOrder
Set IBracketOrder_StopLossOrder = mStopLossOrder
End Property

Private Property Let IBracketOrder_StopReason(ByVal RHS As String)
StopReason = RHS
End Property

Private Property Get IBracketOrder_StopReason() As String
IBracketOrder_StopReason = StopReason
End Property

Private Property Get IBracketOrder_TargetOrder() As IOrder
Set IBracketOrder_TargetOrder = mTargetOrder
End Property

Private Property Let IBracketOrder_TargetReason(ByVal RHS As String)
TargetReason = RHS
End Property

Private Property Get IBracketOrder_TargetReason() As String
IBracketOrder_TargetReason = TargetReason
End Property

Private Function IBracketOrder_ToString() As String
Const ProcName As String = "IBracketOrder_ToString"
On Error GoTo Err

IBracketOrder_ToString = ToString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IBracketOrder_Update()
Const ProcName As String = "IBracketOrder_Update"
On Error GoTo Err

Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IRecoverable Interface Members
'@================================================================================

Private Sub IRecoverable_AddRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "IRecoverable_AddRecoverableObjectListener"
On Error GoTo Err

mRecoverableObjectListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IRecoverable_FinishProcessingOrderPersistenceData()

End Sub

Private Sub IRecoverable_FinishRecovery()
Const ProcName As String = "IRecoverable_FinishRecovery"
On Error GoTo Err

If mSummary.CancelPrice <> 0# Then
    Dim lPriceCanceller As New BracketPriceCanceller
    lPriceCanceller.Initialise Me, mSummary.CancelPrice, mDataSource
End If

If mSummary.CancelTime <> 0 And mSummary.CancelTime > Now Then Set mCancelTLI = mTimerList.Add(Nothing, mSummary.CancelTime, ExpiryTimeUnits.ExpiryTimeUnitDateTime)

If mSummary.RolloverDate <> 0 Then
    If mSummary.RolloverDate > Now Then
        Set mRolloverTLI = mTimerList.Add(Nothing, mSummary.RolloverDate, ExpiryTimeUnitDateTime)
    Else
        Dim lExpiryTime As Date
        lExpiryTime = mContract.ExpiryDate + 10# / 24#
        Dim lTz As TimeZone: Set lTz = GetTimeZone(mContract.TimezoneName)
        lExpiryTime = ConvertDateTzToLocal(lExpiryTime, lTz)
        If lExpiryTime > Now Then
            initiateRollover
        End If
    End If
End If

CheckComplete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IRecoverable_LoadRecoveryData( _
                ByVal pRecoveryContextName As String, _
                ByVal pData As Variant)
Const ProcName As String = "IRecoverable_LoadRecoveryData"
On Error GoTo Err

If mRecoveryContextName = "" Then mRecoveryContextName = pRecoveryContextName

Dim lSummary As BOS: Set lSummary = pData
If mSummary.Key = "" Then
    mSummary.Key = lSummary.Key
    GIdProvider.gAddExistingId lSummary.Key
End If
lSummary.Diff.UpdateHost mSummary

setState mSummary.State

If mSummary.CloseoutOrderId <> "" And mCloseoutOrder Is Nothing And Not mOrderContext Is Nothing Then CloseoutOrder = CreateAnOrder(mSummary.CloseoutOrderId)
If mSummary.EntryOrderId <> "" And mEntryOrder Is Nothing And Not mOrderContext Is Nothing Then EntryOrder = CreateAnOrder(mSummary.EntryOrderId)
If mSummary.StopLossOrderId <> "" And mStopLossOrder Is Nothing And Not mOrderContext Is Nothing Then StopLossOrder = CreateAnOrder(mSummary.StopLossOrderId)
If mSummary.TargetOrderId <> "" And mTargetOrder Is Nothing And Not mOrderContext Is Nothing Then TargetOrder = CreateAnOrder(mSummary.TargetOrderId)

mSummary.setClean

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IRecoverable_ProgId() As String
IRecoverable_ProgId = ProjectName & "." & ModuleName
End Property

Private Property Get IRecoverable_RecoveryContextName() As String
IRecoverable_RecoveryContextName = mRecoveryContextName
End Property

Private Property Get IRecoverable_RecoveryData() As Variant
Set IRecoverable_RecoveryData = mSummary
End Property

Private Sub IRecoverable_RemoveRecoverableObjectListener(ByVal pListener As IRecoverableObjectListener)
Const ProcName As String = "IRecoverable_RemoveRecoverableObjectListener"
On Error GoTo Err

mRecoverableObjectListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

mSummary.RolloverController = ev.Source
SaveRecoveryInfo

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IStringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
IStringable_ToString = ToString
End Function

'@================================================================================
' mCloseoutOrder Event Handlers
'@================================================================================

Private Sub mCloseoutOrder_Error(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "mCloseoutOrder_Error"
On Error GoTo Err

handleStimulus OpStimuli.StimCloseoutOrderCancelled
fireBracketOrderError mCloseoutOrder, pErrorCode, pErrorMsg

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mCloseoutOrder_Fill(ByVal pExec As IExecutionReport)
Const ProcName As String = "mCloseoutOrder_Fill"
On Error GoTo Err

gLogOrderMessage "Closeout order fill: " & _
                    gOrderActionToString(mCloseoutOrder.Action) & " " & pExec.Quantity & " at " & _
                    pExec.Price & "; remaining: " & mCloseoutOrder.QuantityRemaining, _
                mCloseoutOrder, mDataSource, mContract, Key, mIsSimulated, Me

processExecution pExec

fireChange BracketOrderCloseoutOrderFilled
RaiseEvent CloseoutOrderFilled

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mCloseoutOrder_Message(ByVal pMessage As String)
Const ProcName As String = "mCloseoutOrder_Message"
On Error GoTo Err

fireBracketOrderMessage mCloseoutOrder, pMessage

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mCloseoutOrder_PropertyChanged()
Const ProcName As String = "mCloseoutOrder_PropertyChanged"
On Error GoTo Err

fireChange BracketOrderCloseoutOrderChanged
RaiseEvent CloseoutOrderChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mCloseoutOrder_StatusChanged()
Const ProcName As String = "mCloseoutOrder_StatusChanged"
On Error GoTo Err

Select Case mCloseoutOrder.Status
Case OrderStatuses.OrderStatusFilled
    gLogOrderMessage "Closeout order completed: " & _
                        gOrderActionToString(mCloseoutOrder.Action) & " " & mCloseoutOrder.QuantityFilled & " at " & _
                        mCloseoutOrder.AveragePrice, _
                    mCloseoutOrder, mDataSource, mContract, Key, mIsSimulated, Me
Case OrderStatuses.OrderStatusPendingSubmit
Case OrderStatuses.OrderStatusCancelling
Case OrderStatuses.OrderStatusPreSubmitted
Case OrderStatuses.OrderStatusSubmitted
Case OrderStatuses.OrderStatusCancelled
    gLogOrderMessage "Closeout order cancelled", _
                    mCloseoutOrder, mDataSource, mContract, Key, mIsSimulated, Me
Case OrderStatuses.OrderStatusRejected
    gLogOrderMessage "Closeout order rejected", _
                    mCloseoutOrder, mDataSource, mContract, Key, mIsSimulated, Me
End Select

fireChange BracketOrderCloseoutOrderChanged
RaiseEvent CloseoutOrderChanged


Select Case mCloseoutOrder.Status
Case OrderStatuses.OrderStatusFilled
    Assert Size = 0, "Closeout incomplete: Size remaining is " & Size
    checkAllOrdersComplete
Case OrderStatuses.OrderStatusCancelled
    handleStimulus OpStimuli.StimCloseoutOrderCancelled
    checkAllOrdersComplete
End Select

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mEntryOrder Event Handlers
'@================================================================================

Private Sub mEntryOrder_Dirty()
Const ProcName As String = "mEntryOrder_Dirty"
On Error GoTo Err

setDirty

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mEntryOrder_Error(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "mEntryOrder_Error"
On Error GoTo Err

handleStimulus OpStimuli.StimOrderError
fireBracketOrderError mEntryOrder, pErrorCode, pErrorMsg

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mEntryOrder_Fill(ByVal pExec As IExecutionReport)
Const ProcName As String = "mEntryOrder_Fill"
On Error GoTo Err

gLogOrderMessage "Entry order fill: " & _
                    gOrderActionToString(mEntryOrder.Action) & " " & pExec.Quantity & " at " & _
                    pExec.Price & "; remaining: " & mEntryOrder.QuantityRemaining, _
                mEntryOrder, mDataSource, mContract, Key, mIsSimulated, Me

placePeggedOrders pExec.Quantity
        
' we do this after placing any pegged orders. Then, in
' a recovery situation, the pegged orders will not be
' resubmitted.
processExecution pExec

If Not mCancelTLI Is Nothing Then
    mCancelTLI.Cancel
    Set mCancelTLI = Nothing
End If

mSummary.CancelPrice = 0

'If Not mStopLossOrder Is Nothing Then
'    Dim Risk As Currency
'    If LongPosition Then
'        If mStopLossOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
'            Risk = (mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * mStopLossOrder.Offset * mContract.TickValue
'        Else
'            Risk = (mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * (pExec.Price - mStopLossOrder.TriggerPrice) * mContract.Multiplier
'        End If
'    Else
'        If mStopLossOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
'            Risk = -(mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * mStopLossOrder.Offset * mContract.TickValue
'        Else
'            Risk = (mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * (mStopLossOrder.TriggerPrice - pExec.Price) * mContract.Multiplier
'        End If
'    End If
'End If

If Not mStopLossOrder Is Nothing Then
    If mStopLossOrder.Quantity <> Abs(Size) And Abs(Size) <> 0 Then
        AdjustStopLoss mStopLossOrder.TriggerPrice, Abs(Size)
    End If
End If

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Quantity <> Abs(Size) And Abs(Size) <> 0 Then
        AdjustTarget mTargetOrder.LimitPrice, Abs(Size)
    End If
End If

handleStimulus OpStimuli.StimEntryOrderFill

If State <> BracketOrderStateCancelling Then
    fireChange BracketOrderEntryOrderFilled
    RaiseEvent EntryOrderFilled
End If

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mEntryOrder_Message(ByVal pMessage As String)
Const ProcName As String = "mEntryOrder_Message"
On Error GoTo Err

fireBracketOrderMessage mEntryOrder, pMessage

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mEntryOrder_PropertyChanged()
Const ProcName As String = "mEntryOrder_PropertyChanged"
On Error GoTo Err

fireChange BracketOrderEntryOrderChanged
RaiseEvent EntryOrderChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mEntryOrder_StatusChanged()
Const ProcName As String = "mEntryOrder_StatusChanged"
On Error GoTo Err

Select Case mEntryOrder.Status
Case OrderStatuses.OrderStatusFilled
    gLogOrderMessage "Entry order completed: " & _
                        gOrderActionToString(mEntryOrder.Action) & " " & mEntryOrder.QuantityFilled & " at " & _
                        mEntryOrder.AveragePrice, _
                    mEntryOrder, mDataSource, mContract, Key, mIsSimulated, Me
Case OrderStatuses.OrderStatusCancelled
    gLogOrderMessage "Entry order cancelled", _
                    mEntryOrder, mDataSource, mContract, Key, mIsSimulated, Me
Case OrderStatuses.OrderStatusRejected
    gLogOrderMessage "Entry order rejected", _
                    mEntryOrder, mDataSource, mContract, Key, mIsSimulated, Me
End Select

fireChange BracketOrderEntryOrderChanged
RaiseEvent EntryOrderChanged

Select Case mEntryOrder.Status
Case OrderStatuses.OrderStatusFilled
    mSummary.CancelPrice = 0
    RaiseEvent EntryOrderFilled
    checkAllOrdersComplete
Case OrderStatuses.OrderStatusPendingSubmit
Case OrderStatuses.OrderStatusCancelling
Case OrderStatuses.OrderStatusPreSubmitted
Case OrderStatuses.OrderStatusSubmitted
'    If mEntryOrder.QuantityFilled > 0 Then
'        placePeggedOrders
'    End If
Case OrderStatuses.OrderStatusCancelled
    handleStimulus OpStimuli.StimEntryOrderCancelled
    checkAllOrdersComplete
Case OrderStatuses.OrderStatusRejected
    Cancel
    checkAllOrdersComplete
End Select

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mRolloverTLI Event Handlers
'@================================================================================

Private Sub mRolloverTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mRolloverTLI_StateChange"
On Error GoTo Err

initiateRollover

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mStateTimeoutTLI Event Handlers
'@================================================================================

Private Sub mStateTimeoutTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mStateTimeoutTLI_TimerExpired"
On Error GoTo Err

If ev.State <> TimerListItemStates.TimerListItemStateExpired Then Exit Sub

handleStimulus OpStimuli.StimTimeoutExpired

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mStopLossOrder Event Handlers
'@================================================================================

Private Sub mStopLossOrder_Dirty()
Const ProcName As String = "mStopLossOrder_Dirty"
On Error GoTo Err

setDirty

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mStopLossOrder_Error(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "mStopLossOrder_Error"
On Error GoTo Err

handleStimulus OpStimuli.StimOrderError
fireBracketOrderError mStopLossOrder, pErrorCode, pErrorMsg

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mStopLossOrder_Fill(ByVal pExec As IExecutionReport)
Const ProcName As String = "mStopLossOrder_Fill"
On Error GoTo Err

gLogOrderMessage "Stop-loss order fill: " & _
                    gOrderActionToString(mStopLossOrder.Action) & " " & pExec.Quantity & " at " & _
                    pExec.Price & "; remaining: " & mStopLossOrder.QuantityRemaining, _
                mStopLossOrder, mDataSource, mContract, Key, mIsSimulated, Me

processExecution pExec

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Quantity <> Abs(Size) And Abs(Size) <> 0 Then
        AdjustTarget mTargetOrder.LimitPrice, Abs(Size)
    End If
End If

fireChange BracketOrderStopLossOrderFilled
RaiseEvent StopLossOrderFilled

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mStopLossOrder_Message(ByVal pMessage As String)
Const ProcName As String = "mStopLossOrder_Message"
On Error GoTo Err

fireBracketOrderMessage mStopLossOrder, pMessage

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mStopLossOrder_PropertyChanged()
Const ProcName As String = "mStopLossOrder_PropertyChanged"
On Error GoTo Err

fireChange BracketOrderStopLossOrderChanged
RaiseEvent StopLossOrderChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mStopLossOrder_StatusChanged()
Const ProcName As String = "mStopLossOrder_StatusChanged"
On Error GoTo Err

Select Case mStopLossOrder.Status
Case OrderStatuses.OrderStatusFilled
    gLogOrderMessage "Stop-loss order completed: " & _
                        gOrderActionToString(mStopLossOrder.Action) & " " & mStopLossOrder.QuantityFilled & " at " & _
                        mStopLossOrder.AveragePrice, _
                    mStopLossOrder, mDataSource, mContract, Key, mIsSimulated, Me
Case OrderStatuses.OrderStatusCancelled
    gLogOrderMessage "Stop-loss order cancelled", _
                    mStopLossOrder, mDataSource, mContract, Key, mIsSimulated, Me
Case OrderStatuses.OrderStatusRejected
    gLogOrderMessage "Stop-loss order rejected", _
                    mStopLossOrder, mDataSource, mContract, Key, mIsSimulated, Me
End Select

fireChange BracketOrderStopLossOrderChanged
RaiseEvent StopLossOrderChanged

Select Case mStopLossOrder.Status
Case OrderStatuses.OrderStatusFilled
    RaiseEvent StopLossOrderFilled
    checkAllOrdersComplete
Case OrderStatuses.OrderStatusPendingSubmit
Case OrderStatuses.OrderStatusCancelling
Case OrderStatuses.OrderStatusPreSubmitted
Case OrderStatuses.OrderStatusSubmitted
Case OrderStatuses.OrderStatusCancelled
    handleStimulus OpStimuli.StimStopOrderCancelled
    checkAllOrdersComplete
Case OrderStatuses.OrderStatusRejected
    handleStimulus OpStimuli.StimStopOrderCancelled
    checkAllOrdersComplete
End Select

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTargetOrder Event Handlers
'@================================================================================

Private Sub mTargetOrder_Dirty()
Const ProcName As String = "mTargetOrder_Dirty"
On Error GoTo Err

setDirty

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mTargetOrder_Error(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
Const ProcName As String = "mTargetOrder_Error"
On Error GoTo Err

handleStimulus OpStimuli.StimOrderError
fireBracketOrderError mTargetOrder, pErrorCode, pErrorMsg

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mTargetOrder_Fill(ByVal pExec As IExecutionReport)
Const ProcName As String = "mTargetOrder_Fill"
On Error GoTo Err

gLogOrderMessage "Target order fill: " & _
                    gOrderActionToString(mTargetOrder.Action) & " " & pExec.Quantity & " at " & _
                    pExec.Price & "; remaining: " & mTargetOrder.QuantityRemaining, _
                mTargetOrder, mDataSource, mContract, Key, mIsSimulated, Me

processExecution pExec

If Not mStopLossOrder Is Nothing Then
    If mStopLossOrder.Quantity <> Abs(Size) And Abs(Size) <> 0 Then
        AdjustStopLoss mStopLossOrder.TriggerPrice, Abs(Size)
    End If
End If

fireChange BracketOrderTargetOrderFilled
RaiseEvent TargetOrderFilled

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mTargetOrder_Message(ByVal pMessage As String)
Const ProcName As String = "mTargetOrder_Message"
On Error GoTo Err

fireBracketOrderMessage mTargetOrder, pMessage

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mTargetOrder_PropertyChanged()
Const ProcName As String = "mTargetOrder_PropertyChanged"
On Error GoTo Err

fireChange BracketOrderTargetOrderChanged
RaiseEvent TargetOrderChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mTargetOrder_StatusChanged()
Const ProcName As String = "mTargetOrder_StatusChanged"
On Error GoTo Err

Select Case mTargetOrder.Status
Case OrderStatuses.OrderStatusFilled
    gLogOrderMessage "Target order completed: " & _
                        gOrderActionToString(mTargetOrder.Action) & " " & mTargetOrder.QuantityFilled & " at " & _
                        mTargetOrder.AveragePrice, _
                    mTargetOrder, mDataSource, mContract, Key, mIsSimulated, Me
Case OrderStatuses.OrderStatusCancelled
    gLogOrderMessage "Target order cancelled", _
                    mTargetOrder, mDataSource, mContract, Key, mIsSimulated, Me
Case OrderStatuses.OrderStatusRejected
    gLogOrderMessage "Target order rejected", _
                    mTargetOrder, mDataSource, mContract, Key, mIsSimulated, Me
End Select

fireChange BracketOrderTargetOrderChanged
RaiseEvent TargetOrderChanged

Select Case mTargetOrder.Status
Case OrderStatuses.OrderStatusFilled
    RaiseEvent TargetOrderFilled
    checkAllOrdersComplete
Case OrderStatuses.OrderStatusCancelling
Case OrderStatuses.OrderStatusPendingSubmit
Case OrderStatuses.OrderStatusSubmitted
Case OrderStatuses.OrderStatusCancelled
    handleStimulus OpStimuli.StimTargetOrderCancelled
    checkAllOrdersComplete
Case OrderStatuses.OrderStatusRejected
    handleStimulus OpStimuli.StimTargetOrderCancelled
    checkAllOrdersComplete
End Select

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mCancelTLI Event Handlers
'@================================================================================

Private Sub mCancelTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mCancelTLI_StateChange"
On Error GoTo Err

If ev.State <> TimerListItemStates.TimerListItemStateExpired Then Exit Sub

SelfCancel "Cancel time expired"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ApplicationIndex() As Long
ApplicationIndex = mApplicationIndex
End Property

Public Property Let CancelAfter(ByVal Value As Long)
Const ProcName As String = "CancelAfter"
On Error GoTo Err

Assert State = BracketOrderStateCreated, "CancelAfter cannot be set after execution"
Assert mSummary.CancelTime = 0, "CancelTime is already set"

mCancelAfter = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let CancelPrice(ByVal Value As Double)
Const ProcName As String = "CancelPrice"
On Error GoTo Err

Assert State = BracketOrderStateCreated, "CancelPrice cannot be set after execution"

mSummary.CancelPrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CancelPrice() As Double
CancelPrice = mSummary.CancelPrice
End Property

Public Property Let CancelTime(ByVal Value As Date)
Const ProcName As String = "CancelTime"
On Error GoTo Err

Assert State = BracketOrderStateCreated, "CancelTime cannot be set after execution"
Assert mCancelAfter = 0, "CancelAfter is already set"

mSummary.CancelTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CancelTime() As Date
CancelTime = mSummary.CancelTime
End Property

Friend Property Let CloseoutOrder(ByVal Value As IOrder)
Const ProcName As String = "CloseoutOrder"
On Error GoTo Err

Set mCloseoutOrder = Value
mSummary.CloseoutOrderId = mCloseoutOrder.Id

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CloseoutOrder() As IOrder
Set CloseoutOrder = mCloseoutOrder
End Property

Public Property Get Contract() As IContract
Const ProcName As String = "Contract"
On Error GoTo Err

Set Contract = mContract

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let CreationTime(ByVal Value As Date)
Const ProcName As String = "CreationTime"
On Error GoTo Err

mSummary.CreationTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CreationTime() As Date
CreationTime = mSummary.CreationTime
End Property

Friend Property Let CumBuyPrice(ByVal Value As Double)
mSummary.CumBuyPrice = Value
End Property

Public Property Get CumBuyPrice() As Double
CumBuyPrice = mSummary.CumBuyPrice
End Property

Friend Property Let CumSellPrice(ByVal Value As Double)
Const ProcName As String = "CumSellPrice"
On Error GoTo Err

mSummary.CumSellPrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CumSellPrice() As Double
CumSellPrice = mSummary.CumSellPrice
End Property

Public Property Let Description(ByVal Value As String)
Const ProcName As String = "Description"
On Error GoTo Err

mSummary.Description = Value
If State <> BracketOrderStateCreated Then SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Description() As String
Description = mSummary.Description
End Property

Friend Property Let EntryOrder(ByVal Value As IOrder)
Const ProcName As String = "EntryOrder"
On Error GoTo Err

Assert mEntryOrder Is Nothing, "Entry order already exists"

Set mEntryOrder = Value
mSummary.EntryOrderId = mEntryOrder.Id

mNumberOfOrders = mNumberOfOrders + 1

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get EntryOrder() As IOrder
Set EntryOrder = mEntryOrder
End Property

Public Property Get EntryPrice() As Double
Const ProcName As String = "EntryPrice"
On Error GoTo Err

If LongPosition Then
    EntryPrice = mSummary.CumBuyPrice / mEntryOrder.Quantity
Else
    EntryPrice = mSummary.CumSellPrice / mEntryOrder.Quantity
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let EntryReason(ByVal Value As String)
Const ProcName As String = "EntryReason"
On Error GoTo Err

mSummary.EntryReason = Value
If State <> BracketOrderStateCreated Then SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get EntryReason() As String
EntryReason = mSummary.EntryReason
End Property

Public Property Get ExitPrice() As Double
Const ProcName As String = "ExitPrice"
On Error GoTo Err

If LongPosition Then
    ExitPrice = mSummary.CumSellPrice / mEntryOrder.Quantity
Else
    ExitPrice = mSummary.CumBuyPrice / mEntryOrder.Quantity
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get GroupName() As String
GroupName = mSummary.GroupName
End Property

Public Property Get IsComplete() As Boolean
Const ProcName As String = "IsComplete"
On Error GoTo Err

IsComplete = (mStateEngine.State = BracketOrderStates.BracketOrderStateClosed)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsDirty() As Boolean
IsDirty = mIsDirty
End Property

Public Property Get IsRiskUnlimited() As Boolean
Const ProcName As String = "IsRiskUnlimited"
On Error GoTo Err

IsRiskUnlimited = (Risk = MaxCurrency)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Get Key() As String
Key = mSummary.Key
End Property

Public Property Get LongPosition() As Boolean
Const ProcName As String = "LongPosition"
On Error GoTo Err

If Not mEntryOrder Is Nothing Then LongPosition = (mEntryOrder.Action = OrderActionBuy)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Order(ByVal pRole As BracketOrderRoles) As IOrder
Const ProcName As String = "Order"
On Error GoTo Err

Select Case pRole
Case BracketOrderRoleEntry
    Set Order = mEntryOrder
Case BracketOrderRoleStopLoss
    Set Order = mStopLossOrder
Case BracketOrderRoleTarget
    Set Order = mTargetOrder
Case BracketOrderRoleCloseout
    Set Order = mCloseoutOrder
Case Else
    AssertArgument False, "Invalid pRole"
End Select

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OrderContext() As OrderContext
Const ProcName As String = "OrderContext"
On Error GoTo Err

Set OrderContext = mOrderContext

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get NumberOfOrders() As Long
Const ProcName As String = "NumberOfOrders"
On Error GoTo Err

NumberOfOrders = mNumberOfOrders

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PendingSize() As Long
Const ProcName As String = "PendingSize"
On Error GoTo Err

If mEntryOrder Is Nothing Then PendingSize = 0: Exit Property
If mStateEngine.State = BracketOrderStates.BracketOrderStateClosed Then PendingSize = 0: Exit Property
PendingSize = IIf(LongPosition, mEntryOrder.QuantityRemaining, -mEntryOrder.QuantityRemaining)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PreventUnprotectedPosition( _
                ByVal Value As Boolean)
Const ProcName As String = "PreventUnprotectedPosition"
On Error GoTo Err

Assert (Not Value) Or (State = BracketOrderStateCreated) Or (Not mStopLossOrder Is Nothing), "No stop-loss order exists"

mSummary.PreventUnprotectedPosition = Value
If State <> BracketOrderStateCreated Then SaveRecoveryInfo

If mSummary.PreventUnprotectedPosition Then
    mStateEngine.SetConditions OpConditions.CondProtected
Else
    mStateEngine.ClearConditions OpConditions.CondProtected
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PreventUnprotectedPosition() As Boolean
Const ProcName As String = "PreventUnprotectedPosition"
On Error GoTo Err

PreventUnprotectedPosition = mSummary.PreventUnprotectedPosition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Risk() As Currency
Const ProcName As String = "Risk"
On Error GoTo Err

If mStopLossOrder Is Nothing Then
    Risk = MaxCurrency
    Exit Property
End If

If mStateEngine.State = BracketOrderStates.BracketOrderStateClosed Then Exit Property

If mStopLossOrder.Status = OrderStatuses.OrderStatusFilled Then Exit Property

' the amount at Risk has two parts - that due to fills already obtained, and
' that due to fills pending.

Dim filledTotalPrice As Double
filledTotalPrice = mEntryOrder.QuantityFilled * mEntryOrder.AveragePrice

Dim pendingTotalPrice As Double
Select Case EntryOrder.OrderType
    Case OrderTypes.OrderTypeMarket
        pendingTotalPrice = mEntryOrder.QuantityRemaining * _
                IIf(LongPosition, mDataSource.CurrentTick(TickTypeAsk).Price, mDataSource.CurrentTick(TickTypeBid).Price)
    Case OrderTypes.OrderTypeLimit, _
            OrderTypes.OrderTypeStopLimit, _
            OrderTypeLimitIfTouched, _
            OrderTypeMarketIfTouched
        pendingTotalPrice = mEntryOrder.QuantityRemaining * mEntryOrder.LimitPrice
    Case OrderTypes.OrderTypeStop
        pendingTotalPrice = mEntryOrder.QuantityRemaining * mEntryOrder.TriggerPrice
    
    
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
' need to work out what to do with the rest of these
' !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Case OrderTypeMarketOnClose
    Case OrderTypeLimitOnClose
    Case OrderTypePeggedToMarket
    Case OrderTypeRelative
    Case OrderTypeVWAP
    Case OrderTypeMarketToLimit
    Case OrderTypeQuote
    Case OrderTypeAutoStop
    Case OrderTypeAutoLimit
    Case OrderTypeAdjust
    Case OrderTypeAlert
    Case OrderTypeTrailLimit
    Case OrderTypeMarketWithProtection
    Case OrderTypeMarketOnOpen
    Case OrderTypeLimitOnOpen
    Case OrderTypePeggedToPrimary
End Select

Dim stopTotalPrice As Double
stopTotalPrice = mStopLossOrder.TriggerPrice * _
                        (mEntryOrder.Quantity + mEntryOrder.QuantityFilled)
If LongPosition Then
    Risk = mContract.TickValue * (filledTotalPrice + pendingTotalPrice - stopTotalPrice) / mContract.TickSize
Else
    Risk = mContract.TickValue * (stopTotalPrice - filledTotalPrice - pendingTotalPrice) / mContract.TickSize
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get RolloverSpecification() As RolloverSpecification
Set RolloverSpecification = mSummary.RolloverSpecification
End Property

Friend Property Let RolloverSpecification(ByVal Value As RolloverSpecification)
Const ProcName As String = "RolloverSpecification"
On Error GoTo Err

Assert mSummary.State = BracketOrderStateCreated Or _
        mSummary.State = BracketOrderStateSubmitted, _
        "Bracket order state does not allow rollover"

mSummary.RolloverSpecification = Value

If Not mRolloverTLI Is Nothing Then mRolloverTLI.Cancel

Dim lRolloverTime As Date
lRolloverTime = GetOffsetSessionTimes(mContract.ExpiryDate, _
                                    -1 * mSummary.RolloverSpecification.Days).StartTime + mSummary.RolloverSpecification.Time

Dim lTz As TimeZone: Set lTz = GetTimeZone(mContract.TimezoneName)
lRolloverTime = ConvertDateTzToLocal(lRolloverTime, lTz)
Set mRolloverTLI = mTimerList.Add(Nothing, lRolloverTime, ExpiryTimeUnitDateTime)
mSummary.RolloverDate = lRolloverTime

If mSummary.State = BracketOrderStateSubmitted Then SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Size() As Long
Size = mSummary.Size
End Property

Public Property Get StartTime() As Date
StartTime = mStartTime
End Property

Public Property Get EndTime() As Date
EndTime = mEndTime
End Property

Private Sub setState(ByVal Value As BracketOrderStates)
Const ProcName As String = "setState"
On Error GoTo Err

Dim lConditions As OpConditions
If Not mStateEngine Is Nothing Then lConditions = mStateEngine.Conditions

Set mStateEngine = CreateStateEngine(GBracketOrder.TableBuilder, _
                                    Value)
mStateEngine.SetConditions lConditions

mSummary.State = Value
If IsComplete Then doCompletionActions

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Property Get State() As BracketOrderStates
Const ProcName As String = "State"
On Error GoTo Err

State = mStateEngine.State

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let StopLossOrder(ByVal Value As IOrder)
Const ProcName As String = "StopLossOrder"
On Error GoTo Err

Assert mStopLossOrder Is Nothing, "stop-loss order already exists"

mNumberOfOrders = mNumberOfOrders + 1
Set mStopLossOrder = Value
mSummary.StopLossOrderId = mStopLossOrder.Id

mStateEngine.SetConditions OpConditions.CondStopOrderExists
mStateEngine.ClearConditions OpConditions.CondStopOrderCancelled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StopLossOrder() As IOrder
Set StopLossOrder = mStopLossOrder
End Property

Public Property Let StopReason(ByVal Value As String)
Const ProcName As String = "StopReason"
On Error GoTo Err

mSummary.StopReason = Value
If State <> BracketOrderStateCreated Then SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StopReason() As String
StopReason = mSummary.StopReason
End Property

Friend Property Let TargetOrder(ByVal Value As IOrder)
Const ProcName As String = "TargetOrder"
On Error GoTo Err

Assert mTargetOrder Is Nothing, "Target order already exists"

mNumberOfOrders = mNumberOfOrders + 1
Set mTargetOrder = Value
mSummary.TargetOrderId = mTargetOrder.Id

mStateEngine.SetConditions OpConditions.CondTargetOrderExists
mStateEngine.ClearConditions OpConditions.CondTargetOrderCancelled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TargetOrder() As IOrder
Set TargetOrder = mTargetOrder
End Property

Public Property Let TargetReason(ByVal Value As String)
Const ProcName As String = "TargetReason"
On Error GoTo Err

mSummary.TargetReason = Value
If State <> BracketOrderStateCreated Then SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TargetReason() As String
TargetReason = mSummary.TargetReason
End Property

Public Property Get DataSource() As IMarketDataSource
Set DataSource = mDataSource
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Activate()
Const ProcName As String = "Activate"
On Error GoTo Err

fireChange BracketOrderCreated
If Not mEntryOrder Is Nothing Then
    fireChange BracketOrderEntryOrderChanged
    RaiseEvent EntryOrderChanged
End If
If Not mStopLossOrder Is Nothing Then
    fireChange BracketOrderStopLossOrderChanged
    RaiseEvent StopLossOrderChanged
End If
If Not mTargetOrder Is Nothing Then
    fireChange BracketOrderTargetOrderChanged
    RaiseEvent TargetOrderChanged
End If
If Not mCloseoutOrder Is Nothing Then
    fireChange BracketOrderCloseoutOrderChanged
    RaiseEvent CloseoutOrderChanged
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddBracketOrderErrorListener(ByVal pListener As IBracketOrderErrorListener)
Const ProcName As String = "AddBracketOrderErrorListener"
On Error GoTo Err

mBracketOrderErrorListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddBracketOrderMessageListener(ByVal pListener As IBracketOrderMsgListener)
Const ProcName As String = "AddBracketOrderMessageListener"
On Error GoTo Err

mBracketOrderMessageListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "AddChangeListener"
On Error GoTo Err

mChangeListeners.Add pListener

' notify this new listenener that we are already created
Dim ev As ChangeEventData
Set ev.Source = Me
ev.changeType = BracketOrderChangeTypes.BracketOrderCreated
pListener.Change ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddExecutionListener(ByVal pListener As IExecutionListener)
Const ProcName As String = "AddExecutionListener"
On Error GoTo Err

mExecutionListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AdjustStopLoss(ByVal stopTriggerPrice As Double, _
                    Optional ByVal Quantity As Long = -1)
Const ProcName As String = "AdjustStopLoss"
On Error GoTo Err

SetNewStopLossTriggerPrice stopTriggerPrice
If Quantity <> -1 Then SetNewStopLossQuantity Quantity
Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Public Sub AdjustTarget(ByVal targetPrice As Double, _
                    Optional ByVal Quantity As Long = -1)
Const ProcName As String = "AdjustTarget"
On Error GoTo Err

SetNewTargetPrice targetPrice
If Quantity <> -1 Then SetNewTargetQuantity Quantity
Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Public Function Cancel(Optional ByVal evenIfFilled As Boolean = False) As Boolean
Const ProcName As String = "Cancel"
On Error GoTo Err

mSummary.CancelPrice = 0
If Not mCancelTLI Is Nothing Then
    If mCancelTLI.State = TimerListItemStatePending Then mCancelTLI.Cancel
    Set mCancelTLI = Nothing
End If

If evenIfFilled Then
    mStateEngine.ClearConditions OpConditions.CondNoFillCancellation
    Cancel = handleStimulus(OpStimuli.StimCancelEvenIfFill)
Else
    mStateEngine.SetConditions OpConditions.CondNoFillCancellation
    Cancel = handleStimulus(OpStimuli.StimCancelIfNoFill)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Cancels any changes made to this <code>BracketOrder</code> that have not yet
' been actioned.
'@/
Public Sub CancelChanges()
Const ProcName As String = "CancelChanges"
On Error GoTo Err

mNewQuantity = -1
mNewEntryPrice = -1
mNewEntryTriggerPrice = -1
mNewStopLossPrice = -1
mNewStopLossOffset = DummyOffset
mNewStopLossTriggerPrice = -1
mNewStopLossQuantity = -1
mNewTargetPrice = -1
mNewTargetOffset = DummyOffset
mNewTargetTriggerPrice = -1
mNewTargetQuantity = -1

''
' note that we can't just use the copies because external
' code may have references to the member orders which must
' remain valid
'@/
gSyncToOrder mEntryOrder, mEntryOrderCopy
mEntryOrder.setClean
gSyncToOrder mStopLossOrder, mStopOrderCopy
mStopLossOrder.setClean
gSyncToOrder mTargetOrder, mTargetOrderCopy
mTargetOrder.setClean

mIsDirty = False
fireChange BracketOrderChangeTypes.BracketOrderChangesCancelled
RaiseEvent ChangesCancelled

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub CheckComplete()
Const ProcName As String = "CheckComplete"
On Error GoTo Err

checkAllOrdersComplete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Copies the current state of the member orders for use by cancelChanges
'@/
Public Sub Checkpoint()
Const ProcName As String = "Checkpoint"
On Error GoTo Err

Set mEntryOrderCopy = mEntryOrder.Clone
mEntryOrder.setClean

If Not mStopLossOrder Is Nothing Then
    Set mStopOrderCopy = mStopLossOrder.Clone
    mStopLossOrder.setClean
End If

If Not mTargetOrder Is Nothing Then
    Set mTargetOrderCopy = mTargetOrder.Clone
    mTargetOrder.setClean
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub CloseOut()
Const ProcName As String = "CloseOut"
On Error GoTo Err

gLogBracketOrderMessage "Closing out: " & _
                    "size=" & Size & _
                    "; pending size=" & PendingSize, _
                mDataSource, mContract, Key, mIsSimulated, Me

handleStimulus OpStimuli.StimCloseout

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ContainsOrder( _
                ByVal pOrder As IOrder) As Boolean
Const ProcName As String = "ContainsOrder"
On Error GoTo Err

AssertArgument Not pOrder Is Nothing, "pOrder is Nothing"

ContainsOrder = True
If mEntryOrder Is pOrder Then Exit Function
If mStopLossOrder Is pOrder Then Exit Function
If mTargetOrder Is pOrder Then Exit Function
If mCloseoutOrder Is pOrder Then Exit Function
ContainsOrder = False

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Execute()
Const ProcName As String = "Execute"
On Error GoTo Err

Assert (Not mSummary.PreventUnprotectedPosition) Or (Not mStopLossOrder Is Nothing), "No stop-loss order exists"

mSummary.CreationTime = GetTimestamp

If mCancelAfter <> 0 Then
    mSummary.CancelTime = mTimerList.Clock.Timestamp + CDbl(mCancelAfter) / 86400#
    mCancelAfter = 0
End If

If mSummary.CancelTime <> 0 Then Set mCancelTLI = mTimerList.Add(Nothing, mSummary.CancelTime, ExpiryTimeUnits.ExpiryTimeUnitDateTime)
    
If mSummary.CancelPrice <> 0# Then
    Dim lPriceCanceller As New BracketPriceCanceller
    lPriceCanceller.Initialise Me, mSummary.CancelPrice, mDataSource
End If

fireChange BracketOrderChangeTypes.BracketOrderCreated

SaveRecoveryInfo    ' ensure the BracketOrder is recorded before
                    ' doing the state change, which will cause
                    ' the orders to be recorded
mEntryOrder.SaveRecoveryInfo
If Not mStopLossOrder Is Nothing Then mStopLossOrder.SaveRecoveryInfo
If Not mTargetOrder Is Nothing Then mTargetOrder.SaveRecoveryInfo

handleStimulus OpStimuli.StimExecute

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Set mDataSource = Nothing
Set mOrderContext = Nothing
Set mOrderSubmitter = Nothing
If Not mEntryOrder Is Nothing Then mEntryOrder.Finish
If Not mStopLossOrder Is Nothing Then mStopLossOrder.Finish
If Not mTargetOrder Is Nothing Then mTargetOrder.Finish
If Not mCloseoutOrder Is Nothing Then mCloseoutOrder.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pKey As String, _
                ByVal pRecoveryContextName As String, _
                ByVal pGroupName As String, _
                ByVal pOrderContext As OrderContext, _
                ByVal pContract As IContract, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pIsSimulated As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

mSummary.Key = pKey
mRecoveryContextName = pRecoveryContextName
mSummary.GroupName = pGroupName
Set mOrderContext = pOrderContext
mSummary.ContextName = mOrderContext.Name
Set mContract = pContract
Set mDataSource = pDataSource
Set mOrderSubmitter = pOrderSubmitter
mIsSimulated = pIsSimulated

Set mTimerList = mOrderContext.TimerList

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyCloseout( _
                ByVal pExec As IExecutionReport)
Const ProcName As String = "NotifyCloseout"
On Error GoTo Err

If Size = 0 Then Exit Sub

createCloseoutOrder

pExec.OrderId = mCloseoutOrder.Id
mCloseoutOrder.NotifyFillEx pExec, False

fireChange BracketOrderPositionCloseoutNotified
RaiseEvent PositionCloseoutNotified
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveBracketOrderErrorListener(ByVal pListener As IBracketOrderErrorListener)
Const ProcName As String = "RemoveBracketOrderErrorListener"
On Error GoTo Err

mBracketOrderErrorListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveBracketOrderMessageListener(ByVal pListener As IBracketOrderMsgListener)
Const ProcName As String = "RemoveBracketOrderMessageListener"
On Error GoTo Err

mBracketOrderMessageListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "RemoveChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveExecutionListener(ByVal pListener As IExecutionListener)
Const ProcName As String = "RemoveExecutionListener"
On Error GoTo Err

mExecutionListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SelfCancel(ByVal pReason As String)
gLogBracketOrderMessage "Cancel bracket order: " & pReason, _
                mDataSource, mContract, Key, mIsSimulated, Me
Cancel
fireChange BracketOrderSelfCancelled
RaiseEvent SelfCancelled
End Sub

Public Sub SetNewEntryPrice(ByVal Value As Double)
Const ProcName As String = "SetNewEntryPrice"
On Error GoTo Err

Assert Not mEntryOrder Is Nothing, "No entry order exists"

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateCancelling
    Exit Sub
Case BracketOrderStates.BracketOrderStateSubmitted
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert mEntryOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change entry Price - order " & EntryOrder.Id & " already filled"

Select Case mEntryOrder.OrderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeTrail, _
        OrderTypes.OrderTypeAutoStop, _
        OrderTypes.OrderTypeAutoLimit, _
        OrderTypes.OrderTypeStop
        Assert False, "Can't change entry Price for order type " & gOrderTypeToString(mEntryOrder.OrderType)
End Select

If Value <> mEntryOrder.LimitPrice Then
    gLogBracketOrderMessage "Set entry price=" & gPriceToString(Value, mContract), _
                    mDataSource, mContract, Key, mIsSimulated, Me
    mNewEntryPrice = Value
    setDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewEntryTriggerPrice(ByVal Value As Double)
Const ProcName As String = "SetNewEntryTriggerPrice"
On Error GoTo Err

Assert Not mEntryOrder Is Nothing, "No entry order exists"

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateCancelling
    Exit Sub
Case BracketOrderStates.BracketOrderStateSubmitted
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert mEntryOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change entry trigger Price - order " & EntryOrder.Id & " already filled"

Select Case mEntryOrder.OrderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeLimit, _
        OrderTypes.OrderTypeTrail, _
        OrderTypes.OrderTypeAutoLimit
        Assert False, "Can't change trigger Price for order type " & gOrderTypeToString(mEntryOrder.OrderType)
End Select

If Value <> mEntryOrder.TriggerPrice Then
    gLogBracketOrderMessage "Set entry trigger price=" & gPriceToString(Value, mContract), _
                    mDataSource, mContract, Key, mIsSimulated, Me
    mNewEntryTriggerPrice = Value
    setDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewOrderPrice( _
                ByVal pRole As BracketOrderRoles, _
                ByVal Value As Double)
Const ProcName As String = "SetNewOrderPrice"
On Error GoTo Err

Select Case pRole
Case BracketOrderRoleEntry
    SetNewEntryPrice Value
Case BracketOrderRoleStopLoss
    SetNewStopLossPrice Value
Case BracketOrderRoleTarget
    SetNewTargetPrice Value
Case Else
    AssertArgument False, "Invalid pRole"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewOrderTriggerPrice( _
                ByVal pRole As BracketOrderRoles, _
                ByVal Value As Double)
Const ProcName As String = "SetNewOrderTriggerPrice"
On Error GoTo Err

Select Case pRole
Case BracketOrderRoleEntry
    SetNewEntryTriggerPrice Value
Case BracketOrderRoleStopLoss
    SetNewStopLossTriggerPrice Value
Case BracketOrderRoleTarget
    SetNewTargetTriggerPrice Value
Case Else
    AssertArgument False, "Invalid pRole"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewQuantity(ByVal Value As Long)
Const ProcName As String = "SetNewQuantity"
On Error GoTo Err

Assert Not mEntryOrder Is Nothing, "No entry order"

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateCancelling
    Exit Sub
Case BracketOrderStates.BracketOrderStateSubmitted
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert mEntryOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change quantity - order " & EntryOrder.Id & " already filled"

If Value <> mEntryOrder.Quantity Then
    gLogBracketOrderMessage "Set quantity=" & Value, _
                    mDataSource, mContract, Key, mIsSimulated, Me
    mNewQuantity = Value
    setDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewStopLossOffset(ByVal Value As Long)
Const ProcName As String = "SetNewStopLossOffset"
On Error GoTo Err

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateSubmitted
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert Not mStopLossOrder Is Nothing, "No stop-loss order exists"
Assert mStopLossOrder.OrderType = OrderTypes.OrderTypeAutoStop, "Can't change offset for order type " & gOrderTypeToString(mEntryOrder.OrderType)
Assert mStopLossOrder.Status = OrderStatuses.OrderStatusCreated, "Can't change stop offset - order " & mStopLossOrder.Id & " already placed"

If Value <> mStopLossOrder.Offset Then
    gLogBracketOrderMessage "Set stop-loss offset=" & Value, _
                    mDataSource, mContract, Key, mIsSimulated, Me
    mNewStopLossOffset = Value
    setDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewStopLossPrice(ByVal Value As Double)
Const ProcName As String = "SetNewStopLossPrice"
On Error GoTo Err

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateCancelling
    Exit Sub
Case BracketOrderStates.BracketOrderStateSubmitted
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert Not mStopLossOrder Is Nothing, "No stop-loss order exists"

Select Case mStopLossOrder.OrderType
Case OrderTypes.OrderTypeStopLimit
Case Else
    Assert False, "Can't change Price for order type " & gOrderTypeToString(mStopLossOrder.OrderType)
End Select

If Value <> mStopLossOrder.LimitPrice Then
    gLogBracketOrderMessage "Set stop-loss price=" & gPriceToString(Value, mContract), _
                    mDataSource, mContract, Key, mIsSimulated, Me
    mNewStopLossPrice = Value
    setDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewStopLossQuantity(ByVal Value As Long)
Const ProcName As String = "SetNewStopLossQuantity"
On Error GoTo Err

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateCancelling
    Exit Sub
Case BracketOrderStates.BracketOrderStateSubmitted
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert Not mStopLossOrder Is Nothing, "No stop-loss order exists"

If Value <> mStopLossOrder.Quantity Then
    gLogBracketOrderMessage "Set stop-loss quantity=" & Value, _
                    mDataSource, mContract, Key, mIsSimulated, Me
    mNewStopLossQuantity = Value
    setDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewStopLossTriggerPrice(ByVal Value As Double)
Const ProcName As String = "SetNewStopLossTriggerPrice"
On Error GoTo Err

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateCancelling
    Exit Sub
Case BracketOrderStates.BracketOrderStateSubmitted
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert Not mStopLossOrder Is Nothing, "No stop-loss order exists"

Select Case mStopLossOrder.OrderType
    Case OrderTypes.OrderTypeStop, _
        OrderTypes.OrderTypeStopLimit
    Case Else
        Assert False, "Can't change trigger Price for order type " & gOrderTypeToString(mStopLossOrder.OrderType)
End Select

If Value <> mStopLossOrder.TriggerPrice Then
    gLogBracketOrderMessage "Set stop-loss trigger price=" & gPriceToString(Value, mContract), _
                    mDataSource, mContract, Key, mIsSimulated, Me
    mNewStopLossTriggerPrice = Value
    setDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewTargetPrice(ByVal Value As Double)
Const ProcName As String = "SetNewTargetPrice"
On Error GoTo Err

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateSubmitted
Case BracketOrderStates.BracketOrderStateCancelling
    Exit Sub
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert Not mTargetOrder Is Nothing, "No target order exists"

Select Case mTargetOrder.OrderType
Case OrderTypes.OrderTypeLimitIfTouched, _
    OrderTypes.OrderTypeLimit
Case Else
    Assert False, "Can't change Price for order type " & gOrderTypeToString(mTargetOrder.OrderType)
End Select

If Value <> mTargetOrder.LimitPrice Then
    gLogBracketOrderMessage "Set target price=" & gPriceToString(Value, mContract), _
                    mDataSource, mContract, Key, mIsSimulated, Me
    mNewTargetPrice = Value
    setDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewTargetOffset(ByVal Value As Long)
Const ProcName As String = "SetNewTargetOffset"
On Error GoTo Err

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateSubmitted
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert Not mTargetOrder Is Nothing, "No target order exists"
Assert mTargetOrder.OrderType = OrderTypes.OrderTypeAutoLimit, "Can't change offset for order type " & gOrderTypeToString(mTargetOrder.OrderType)
Assert mTargetOrder.Status = OrderStatuses.OrderStatusCreated, "Can't change target offset - order " & mTargetOrder.Id & " already placed"

If Value <> mTargetOrder.Offset Then
    gLogBracketOrderMessage "Set target offset=" & Value, _
                    mDataSource, mContract, Key, mIsSimulated, Me
    mNewTargetOffset = Value
    setDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewTargetQuantity(ByVal Value As Long)
Const ProcName As String = "SetNewTargetQuantity"
On Error GoTo Err

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateCancelling
    Exit Sub
Case BracketOrderStates.BracketOrderStateSubmitted
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert Not mTargetOrder Is Nothing, "No target order exists"

If Value <> mTargetOrder.Quantity Then
    gLogBracketOrderMessage "Set target quantity=" & Value, _
                    mDataSource, mContract, Key, mIsSimulated, Me
    mNewTargetQuantity = Value
    setDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewTargetTriggerPrice(ByVal Value As Double)
Const ProcName As String = "SetNewTargetTriggerPrice"
On Error GoTo Err

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateCancelling
    Exit Sub
Case BracketOrderStates.BracketOrderStateSubmitted
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert Not mTargetOrder Is Nothing, "No target order exists"

Select Case mTargetOrder.OrderType
Case OrderTypes.OrderTypeLimitIfTouched, _
    OrderTypes.OrderTypeMarketIfTouched
Case Else
    Assert False, "Can't change trigger Price for order type " & gOrderTypeToString(mTargetOrder.OrderType)
End Select

If Value <> mTargetOrder.TriggerPrice Then
    gLogBracketOrderMessage "Set target trigger price=" & gPriceToString(Value, mContract), _
                    mDataSource, mContract, Key, mIsSimulated, Me
    mNewTargetTriggerPrice = Value
    setDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetSize(ByVal Value As Long)
Const ProcName As String = "SetSize"
On Error GoTo Err

mSummary.Size = Value

If mSummary.Size <> 0 Then
    mStateEngine.SetConditions OpConditions.CondSizeNonZero
Else
    mStateEngine.ClearConditions OpConditions.CondSizeNonZero
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ToString() As String
Const ProcName As String = "ToString"
On Error GoTo Err

ToString = gOrderActionToString(mEntryOrder.Action) & " " & _
            mEntryOrder.Quantity & " " & _
            gGetOrderTypeAndPricesString(mEntryOrder, mContract)

ToString = ToString & "; "
If Not mStopLossOrder Is Nothing Then
    ToString = ToString & _
            gGetOrderTypeAndPricesString(mStopLossOrder, mContract)
End If

ToString = ToString & "; "
If Not mTargetOrder Is Nothing Then
    ToString = ToString & _
        gGetOrderTypeAndPricesString(mTargetOrder, mContract)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Update()
Const ProcName As String = "Update"
On Error GoTo Err

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateCreated
Case BracketOrderStates.BracketOrderStateSubmitted
Case BracketOrderStates.BracketOrderStateCancelling
    Exit Sub
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Dim updateEntry As Boolean
Dim updateStop As Boolean
Dim updateTarget As Boolean
checkForPendingUpdates updateEntry, updateStop, updateTarget

If (Not updateEntry) And (Not updateStop) And (Not updateTarget) Then
    ' no updates to Action
    Exit Sub
End If

gLogBracketOrderMessage "Update bracket order: " & ToString, _
                mDataSource, mContract, Key, mIsSimulated, Me

If State <> BracketOrderStateCreated Then
    mOrderSubmitter.ModifyBracketOrder Me, updateEntry, updateStop, updateTarget, False
End If

Checkpoint

If updateEntry Then
    fireChange BracketOrderEntryOrderChanged
    RaiseEvent EntryOrderChanged
End If
If updateStop Then
    fireChange BracketOrderStopLossOrderChanged
    RaiseEvent StopLossOrderChanged
End If
If updateTarget Then
    fireChange BracketOrderTargetOrderChanged
    RaiseEvent TargetOrderChanged
End If
fireChange BracketOrderChangeTypes.BracketOrderChangesApplied
RaiseEvent ChangesApplied

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function cancelOrders() As Boolean
Const ProcName As String = "cancelOrders"
On Error GoTo Err

mSummary.CancelPrice = 0
If Not mCancelTLI Is Nothing Then
    mCancelTLI.Cancel
    Set mCancelTLI = Nothing
End If

Dim lNeedCancel As Boolean

Select Case mEntryOrder.Status
Case OrderStatuses.OrderStatusCreated
    CancelAfter = 0
    CancelTime = 0
    mEntryOrder.Status = OrderStatuses.OrderStatusCancelled
Case Else
    lNeedCancel = True
End Select

If Not mStopLossOrder Is Nothing Then
    Select Case mStopLossOrder.Status
    Case OrderStatuses.OrderStatusCreated
        mStopLossOrder.Status = OrderStatuses.OrderStatusCancelled
    Case Else
        lNeedCancel = True
    End Select
End If

If Not mTargetOrder Is Nothing Then
    Select Case mTargetOrder.Status
    Case OrderStatuses.OrderStatusCreated
        mTargetOrder.Status = OrderStatuses.OrderStatusCancelled
    Case Else
        lNeedCancel = True
    End Select
End If

If Not lNeedCancel Then Exit Function
If Not mOrderSubmitter.CancelBracketOrder(Me) Then Exit Function

cancelOrders = True

Dim StopLossOrderId As String
If Not mStopLossOrder Is Nothing Then StopLossOrderId = mStopLossOrder.Id

Dim TargetOrderId As String
If Not mTargetOrder Is Nothing Then TargetOrderId = mTargetOrder.Id
gLogBracketOrderMessage "Cancel orders: " & _
                    EntryOrder.Id & _
                    ", " & StopLossOrderId & _
                    ", " & TargetOrderId, _
                mDataSource, mContract, Key, mIsSimulated, Me

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Function cancelStateTimeout()
Const ProcName As String = "cancelStateTimeout"
On Error GoTo Err

mStateTimeoutTLI.Cancel

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function CancelStopOrder()
Const ProcName As String = "cancelStopOrder"
On Error GoTo Err

If mStopLossOrder Is Nothing Then Exit Function
Select Case mStopLossOrder.Status
Case OrderStatuses.OrderStatusCreated
    mStopLossOrder.Status = OrderStatuses.OrderStatusCancelled
Case OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    gLogOrderMessage "Cancel order", _
                    mStopLossOrder, mDataSource, mContract, Key, mIsSimulated, Me
    mOrderSubmitter.CancelStopOrder Me
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function CancelTargetOrder()
Const ProcName As String = "cancelTargetOrder"
On Error GoTo Err

If mTargetOrder Is Nothing Then Exit Function
Select Case mTargetOrder.Status
Case OrderStatuses.OrderStatusCreated
    mTargetOrder.Status = OrderStatuses.OrderStatusCancelled
Case OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    gLogOrderMessage "Cancel order", _
                    mTargetOrder, mDataSource, mContract, Key, mIsSimulated, Me
    mOrderSubmitter.CancelTargetOrder Me
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub checkAllOrdersComplete()
Const ProcName As String = "checkAllOrdersComplete"
On Error GoTo Err

If IsComplete Then Exit Sub

If Not mEntryOrder Is Nothing Then
    If mEntryOrder.IsActive Then Exit Sub
End If

If Not mStopLossOrder Is Nothing Then
    If mStopLossOrder.IsActive Then Exit Sub
End If

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.IsActive Then Exit Sub
End If

If Not mCloseoutOrder Is Nothing Then
    If mCloseoutOrder.IsActive Then Exit Sub
End If

handleStimulus OpStimuli.StimAllOrdersComplete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub checkForPendingUpdates( _
                ByRef updateEntry As Boolean, _
                ByRef updateStop As Boolean, _
                ByRef updateTarget As Boolean)
Const ProcName As String = "checkForPendingUpdates"
On Error GoTo Err

If Not mEntryOrder Is Nothing Then
    If mEntryOrder.IsDirty Then updateEntry = True
    If mNewQuantity >= 0 And _
        mNewQuantity <> mEntryOrder.Quantity _
    Then
        Assert mEntryOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change Size - order " & EntryOrder.Id & " already filled"
        mEntryOrder.Quantity = mNewQuantity
        updateEntry = True
    End If
    mNewQuantity = -1
    
    If mNewEntryPrice >= 0 And _
        mNewEntryPrice <> mEntryOrder.LimitPrice _
    Then
        Assert mEntryOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change price - order " & mEntryOrder.Id & " already filled"
        mEntryOrder.LimitPrice = mNewEntryPrice
        updateEntry = True
    End If
    mNewEntryPrice = -1
    
    If mNewEntryTriggerPrice >= 0 And _
        mNewEntryTriggerPrice <> mEntryOrder.TriggerPrice _
    Then
        Assert mEntryOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change trigger Price - order " & mEntryOrder.Id & " already filled"
        mEntryOrder.TriggerPrice = mNewEntryTriggerPrice
        updateEntry = True
    End If
    mNewEntryTriggerPrice = -1
End If

If Not mStopLossOrder Is Nothing Then
    If mStopLossOrder.IsDirty Then updateStop = True
    If mNewStopLossPrice >= 0 And _
        mNewStopLossPrice <> mStopLossOrder.LimitPrice _
    Then
        Assert mStopLossOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change trigger Price - order " & mStopLossOrder.Id & " already filled"
        Assert mStopLossOrder.Status <> OrderStatuses.OrderStatusSubmitted, "Can't change Price - order " & mStopLossOrder.Id & " already submitted"
        mStopLossOrder.LimitPrice = mNewStopLossPrice
        updateStop = True
    End If
    mNewStopLossPrice = -1
    
    If mNewStopLossOffset <> DummyOffset And _
        mNewStopLossOffset <> mStopLossOrder.Offset _
    Then
        Assert mStopLossOrder.Status = OrderStatuses.OrderStatusCreated, "Can't change stop offset - order " & mStopLossOrder.Id & " already placed"
        mStopLossOrder.Offset = mNewStopLossOffset
        updateStop = True
    End If
    mNewStopLossOffset = DummyOffset

    If mNewStopLossQuantity >= 0 And _
        mNewStopLossQuantity <> mStopLossOrder.Quantity _
    Then
        Assert mStopLossOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change stop quantity - order " & mStopLossOrder.Id & " already filled"
        mStopLossOrder.Quantity = mNewStopLossQuantity
        updateStop = True
    End If
    mNewStopLossQuantity = -1

    If mNewStopLossTriggerPrice >= 0 And _
        mNewStopLossTriggerPrice <> mStopLossOrder.TriggerPrice _
    Then
        Assert mStopLossOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change trigger Price - order " & mEntryOrder.Id & " already filled"
        mStopLossOrder.TriggerPrice = mNewStopLossTriggerPrice
        updateStop = True
    End If
    mNewStopLossTriggerPrice = -1
End If

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.IsDirty Then updateTarget = True
    If mNewTargetPrice >= 0 And _
        mNewTargetPrice <> mTargetOrder.LimitPrice _
    Then
        Assert mTargetOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change Price - order " & mTargetOrder.Id & " already filled"
        mTargetOrder.LimitPrice = mNewTargetPrice
        updateTarget = True
    End If
    mNewTargetPrice = -1
    
    If mNewTargetOffset <> DummyOffset And _
        mNewTargetOffset <> mTargetOrder.Offset _
    Then
        Assert mTargetOrder.Status = OrderStatuses.OrderStatusCreated, "Can't change target offset - order " & TargetOrder.Id & " already placed"
        mTargetOrder.Offset = mNewTargetOffset
        updateTarget = True
    End If
    mNewTargetOffset = DummyOffset

    If mNewTargetQuantity >= 0 And _
        mNewTargetQuantity <> mTargetOrder.Quantity _
    Then
        Assert mTargetOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change target quantity - order " & mTargetOrder.Id & " already filled"
        mTargetOrder.Quantity = mNewTargetQuantity
        updateTarget = True
    End If
    mNewTargetQuantity = -1

    If mNewTargetTriggerPrice >= 0 And _
        mNewTargetTriggerPrice <> mTargetOrder.TriggerPrice _
    Then
        Assert mTargetOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change trigger Price - order " & mEntryOrder.Id & " already filled"
        mTargetOrder.TriggerPrice = mNewTargetTriggerPrice
        updateTarget = True
    End If
    mNewTargetTriggerPrice = -1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function CreateAnOrder(ByVal pId As String) As IOrder
Const ProcName As String = "CreateAnOrder"
On Error GoTo Err

Dim lOrder As IOrder
Set lOrder = mOrderContext.CreateRawOrder(pId, True, mRecoveryContextName)


lOrder.Initialise mSummary.GroupName, mContract.Specifier, mOrderContext
Set CreateAnOrder = lOrder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub createCloseoutOrder()
Const ProcName As String = "createCloseoutOrder"
On Error GoTo Err

CloseoutOrder = mOrderContext.NewOrder(True)

mCloseoutOrder.Action = IIf(LongPosition, OrderActions.OrderActionSell, OrderActions.OrderActionBuy)
mCloseoutOrder.Quantity = Abs(Size)
mCloseoutOrder.OrderType = OrderTypes.OrderTypeMarket
mCloseoutOrder.TimeInForce = OrderTIFs.OrderTIFGoodTillCancelled

SaveRecoveryInfo

fireChange BracketOrderCloseoutOrderCreated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function doActions( _
                ByRef pActions() As OpActions, _
                ByVal pPreviousState As BracketOrderStates, _
                ByVal pStimulus As OpStimuli) As Variant
Const ProcName As String = "doActions"
On Error GoTo Err

Dim lAction As Variant
For Each lAction In pActions
    Select Case CLng(lAction)
    Case SpecialActions.NoAction
    Case OpActions.ActPlaceOrders
        doActions = placeOrders
    Case OpActions.ActCancelOrders
        doActions = cancelOrders
    Case OpActions.ActCancelStopOrder
        doActions = CancelStopOrder
    Case OpActions.ActCancelTargetOrder
        doActions = CancelTargetOrder
    Case OpActions.ActResubmitStopOrder
        doActions = ResubmitStopOrder
    Case OpActions.ActResubmitTargetOrder
        doActions = ResubmitTargetOrder
    Case OpActions.ActResubmitStopAndTargetOrders
        doActions = ResubmitStopAndTargetOrders
    Case OpActions.ActPlaceCloseoutOrder
        doActions = placeCloseoutOrder
    Case OpActions.ActCompletionActions
        doActions = doCompletionActions
    Case OpActions.ActAlarm
        ' need to put something in here
    Case OpActions.ActSetTimeout
        doActions = setStateTimeout
    Case OpActions.ActCancelTimeout
        doActions = cancelStateTimeout
    Case OpActions.ActLog
        gLog "Stimulus " & gOpStimuliToString(pStimulus) & " occurred in state " & gBracketOrderStatesToString(pPreviousState) & " (" & mSummary.Key & ")", ProcName, ModuleName, , LogLevelSevere
        doActions = Empty
    Case Else
        Assert False, "Invalid Action " & lAction
    End Select
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function doCompletionActions()
Const ProcName As String = "doCompletionActions"
On Error GoTo Err

If Not mOrderContext Is Nothing Then
    If mOrderContext.Name <> BalancingOrderContextName Then
        mOrderContext.NotifyBracketOrderClosed Me
    End If
End If

fireChange BracketOrderCompleted
RaiseEvent Completed

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fireBracketOrderError( _
                ByVal pOrder As IOrder, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "fireBracketOrderError"
On Error GoTo Err

Dim ev As BracketOrderErrorEventData
Set ev.Source = Me
Set ev.AffectedOrder = pOrder
ev.ErrorCode = pErrorCode
ev.ErrorMsg = pErrorMsg

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mBracketOrderErrorListeners.Valid Then
    sInit = True
    sSomeListeners = mBracketOrderErrorListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IBracketOrderErrorListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyBracketOrderError ev
    Next
End If

RaiseEvent OrderError(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireBracketOrderMessage( _
                ByVal pOrder As IOrder, _
                ByVal pMessage As String)
Const ProcName As String = "fireBracketOrderMessage"
On Error GoTo Err

Dim ev As BracketOrderMessageEventData
Set ev.Source = Me
Set ev.AffectedOrder = pOrder
ev.Message = pMessage

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mBracketOrderMessageListeners.Valid Then
    sInit = True
    sSomeListeners = mBracketOrderMessageListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IBracketOrderMsgListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyBracketOrderMessage ev
    Next
End If

RaiseEvent OrderMessage(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireChange( _
                ByVal pChangeType As BracketOrderChangeTypes)
Const ProcName As String = "fireChange"
On Error GoTo Err

Dim ev As ChangeEventData
Set ev.Source = Me
ev.changeType = pChangeType

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If

RaiseEvent Change(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireExecution( _
                ByVal pExec As IExecutionReport)
Const ProcName As String = "fireExecution"
On Error GoTo Err

Dim ev As ExecutionEventData
Set ev.Source = Me
Set ev.ExecutionReport = pExec

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mExecutionListeners.Valid Then
    sInit = True
    sSomeListeners = mExecutionListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IExecutionListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyExecution ev
    Next
End If

RaiseEvent Execution(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireStateChange( _
                ByVal pState As BracketOrderStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = pState

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mStateChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mStateChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IStateChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If

RaiseEvent StateChange(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireRecoveryDataChanged()
Const ProcName As String = "fireRecoveryDataChanged"
On Error GoTo Err

Dim ev As RecoveryDataChangedEventData
Set ev.Source = Me

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mRecoverableObjectListeners.Valid Then
    sInit = True
    sSomeListeners = mRecoverableObjectListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IRecoverableObjectListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.NotifyRecoveryDataChanged ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function handleStimulus(ByVal pStimulus As OpStimuli) As Variant
Const ProcName As String = "handleStimulus"
On Error GoTo Err

Dim lPrevState As BracketOrderStates
lPrevState = mStateEngine.State

Dim lConditions As OpConditions
lConditions = mStateEngine.Conditions

Dim lActions() As OpActions
lActions = mStateEngine.NotifyStimulus(pStimulus)
mSummary.State = State
SaveRecoveryInfo

Dim lActionsString As String
Dim i As Long
For i = 0 To UBound(lActions)
    lActionsString = lActionsString & gOpActionsToString(lActions(i)) & IIf(i = UBound(lActions), "", ",")
Next

gLog "Handling bracket order stimulus (" & mSummary.Key & "): " & vbCrLf & _
                "    prior state=" & gBracketOrderStatesToString(lPrevState) & _
                "; stimulus=" & gOpStimuliToString(pStimulus) & _
                "; conditions=" & gOpConditionsToString(lConditions) & _
                "; new state=" & gBracketOrderStatesToString(mStateEngine.State) & _
                "; actions=" & lActionsString, _
            ProcName, ModuleName

handleStimulus = doActions(lActions, lPrevState, pStimulus)
If State <> lPrevState Then
    fireChange BracketOrderChangeTypes.BracketOrderStateChanged
    fireStateChange mStateEngine.State
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub initiateRollover()
Const ProcName As String = "initiateRollover"
On Error GoTo Err

mSummary.RolloverController = New RolloverController
mSummary.RolloverController.Initialise Me, mOrderContext.OrderManager, mOrderContext.ScopeName
StartTask mSummary.RolloverController, PriorityNormal
mSummary.RolloverController.AddStateChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function placeCloseoutOrder()
Const ProcName As String = "placeCloseoutOrder"
On Error GoTo Err

createCloseoutOrder

RaiseEvent CloseoutOrderCreated
mOrderSubmitter.ModifyBracketOrder Me, False, False, False, True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function placeOrders()
Const ProcName As String = "placeOrders"
On Error GoTo Err

mOrderSubmitter.ExecuteBracketOrder Me

Checkpoint

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub placePeggedOrders(ByVal pFillSize As Long)
Const ProcName As String = "placePeggedOrders"
On Error GoTo Err

If mSummary.PeggedOrdersProcessed Then Exit Sub

Dim updated As Boolean

If Not mStopLossOrder Is Nothing Then
    If mStopLossOrder.Status = OrderStatuses.OrderStatusCreated Then
        If mStopLossOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
            ' the stop Price is to be pegged to the entry Price - the
            ' order's offset property will already contain any required
            ' offset from that Price
            mStopLossOrder.OrderType = OrderTypes.OrderTypeStop
            SetNewStopLossTriggerPrice (mStopLossOrder.Offset * mContract.TickSize) + EntryOrder.LastFillPrice
            SetNewStopLossQuantity pFillSize
            updated = True
        End If
    End If
End If
If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Status = OrderStatuses.OrderStatusCreated Then
        If mTargetOrder.OrderType = OrderTypes.OrderTypeAutoLimit Then
            ' the limit Price is to be pegged to the entry Price - the
            ' order's offset property will already contain any required
            ' offset from that Price
            mTargetOrder.OrderType = OrderTypes.OrderTypeLimit
            SetNewTargetPrice (mTargetOrder.Offset * mContract.TickSize) + EntryOrder.LastFillPrice
            SetNewTargetQuantity pFillSize
            updated = True
        End If
    End If
End If

If updated Then
    mSummary.PeggedOrdersProcessed = True
    Update
    SaveRecoveryInfo
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processExecution( _
                ByVal pExec As Execution)
Const ProcName As String = "processExecution"
On Error GoTo Err

' Note that the Order object ensures that pExec will always
' be an Execution object, not just some other object that
' implements IExecutionReport

If mSummary.Size = 0 Then
    mStartTime = pExec.FillTime
    mTimezoneName = pExec.TimezoneName
End If

If pExec.Action = OrderActionBuy Then
    CumBuyPrice = mSummary.CumBuyPrice + pExec.Quantity * pExec.Price
    SetSize Size + pExec.Quantity
Else
    CumSellPrice = mSummary.CumSellPrice + pExec.Quantity * pExec.Price
    SetSize Size - pExec.Quantity
End If
    
If Size = 0 Then mEndTime = pExec.FillTime

fireExecution pExec
fireChange BracketOrderChangeTypes.BracketOrderSizeChanged
RaiseEvent SizeChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function ResubmitStopOrder()
Const ProcName As String = "resubmitStopOrder"
On Error GoTo Err

gLogOrderMessage "Resubmit stop-loss order", _
                mStopLossOrder, mDataSource, mContract, Key, mIsSimulated, Me
mOrderSubmitter.ResubmitStopOrder Me
mStateEngine.ClearConditions OpConditions.CondStopOrderCancelled

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function ResubmitStopAndTargetOrders()
Const ProcName As String = "resubmitStopAndTargetOrders"
On Error GoTo Err

gLogOrderMessage "Resubmit stop-loss order", _
                mStopLossOrder, mDataSource, mContract, Key, mIsSimulated, Me
gLogOrderMessage "Resubmit target order", _
                mTargetOrder, mDataSource, mContract, Key, mIsSimulated, Me
mOrderSubmitter.ResubmitStopAndTargetOrders Me
mStateEngine.ClearConditions OpConditions.CondStopOrderCancelled
mStateEngine.ClearConditions OpConditions.CondTargetOrderCancelled

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Function ResubmitTargetOrder()
Const ProcName As String = "resubmitTargetOrder"
On Error GoTo Err

gLogOrderMessage "Resubmit target order", _
                mTargetOrder, mDataSource, mContract, Key, mIsSimulated, Me
mOrderSubmitter.ResubmitTargetOrder Me
mStateEngine.ClearConditions OpConditions.CondTargetOrderCancelled

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub SaveRecoveryInfo()
Const ProcName As String = "SaveRecoveryInfo"
On Error GoTo Err

If mIsSimulated Then Exit Sub
If Not mSummary.IsDirty Then Exit Sub

If mRecoverableObjectListeners Is Nothing Or mRecoverableObjectListeners.Count = 0 Then Exit Sub

fireRecoveryDataChanged

mSummary.setClean

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Set the 'dirty' flag, and raise a ChangesPending event if one has not already
' been raised.
'@/
Private Sub setDirty()
Const ProcName As String = "setDirty"
On Error GoTo Err

If mIsDirty Then Exit Sub

mIsDirty = True
fireChange BracketOrderChangeTypes.BracketOrderChangesPending
RaiseEvent ChangesPending

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setStateTimeout()
Const ProcName As String = "setStateTimeout"
On Error GoTo Err

Set mStateTimeoutTLI = mTimerList.Add(Empty, 1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

