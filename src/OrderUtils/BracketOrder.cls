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
Implements IPositionMember
Implements IRecoverable
Implements Stringable

'@================================================================================
' Events
'@================================================================================

Event Change(ByRef ev As ChangeEventData)
Event ChangesCancelled()
Event Clean()
Event CloseoutOrderFilled()
Event CloseoutOrderPartiallyFilled()
Event Completed()
Event Dirty()
Event EntryOrderFilled()
Event EntryOrderPartiallyFilled()
Event Execution(ByRef ev As ExecutionEventData)
Event SelfCancelled()
Event StateChange(ByRef ef As StateChangeEventData)
Event StopAdjusted()
Event StopOrderFilled()
Event StopOrderPartiallyFilled()
Event TargetAdjusted()
Event TargetOrderFilled()
Event TargetOrderPartiallyFilled()

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
'    ' the stop order or the target order, when both exist, has been received.
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

Private mOrderContext                               As OrderContext
Private mContract                                   As IContract
Private mDataSource                                 As IMarketDataSource
Private mOrderSubmitter                             As IOrderSubmitter
Private mTimerList                                  As TimerList

'Private mPositionManagerRef                         As WeakReference

Private mGroupName                                  As String
Private mContextsName                               As String
Private mContextName                                As String
Private mKey                                        As String

Private mStateEngine                                As StateEngine

Private mDescription                                As String
Private mEntryReason                                As String
Private mTargetReason                               As String
Private mStopReason                                 As String

Private mSize                                       As Long
Private mCancelPrice                                As Double
Private mCancelAfter                                As Long
Private mCancelTime                                 As Date
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

Private mCreationTime                               As Date

Private mApplicationIndex                           As Long

Private mIsDirty                                    As Boolean

Private mOrders                                     As EnumerableCollection

Private mPreventUnprotectedPosition                 As Boolean

Private WithEvents mStateTimeoutTLI                 As TimerListItem
Attribute mStateTimeoutTLI.VB_VarHelpID = -1

Private mSummary                                    As BracketOrderSummary

Private mStartTime                                  As Date
Private mEndTime                                    As Date
Private mTimezoneName                               As String

Private mCumBuyPrice                                As Double
Private mCumSellPrice                               As Double

Private mBrokerData                                 As Variant

Private mRecoverableObjectListeners                 As New Listeners

Private mExecutionListeners                         As New Listeners

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

Set mSummary = New BracketOrderSummary

mApplicationIndex = GBracketOrder.gNextApplicationIndex

Set mOrders = New EnumerableCollection

Set mStateEngine = CreateStateEngine(GBracketOrder.TableBuilder, _
                                    BracketOrderStates.BracketOrderStateCreated)

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

Private Sub IBracketOrder_AddChangeListener(ByVal pListener As ChangeListener)
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

Private Sub IBracketOrder_AddStateChangeListener(ByVal pListener As StateChangeListener)
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

AdjustStop pStopTriggerPrice, pQuantity

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

Cancel pEvenIfFilled

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
IBracketOrder_ContextName = mContextName
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
IBracketOrder_Key = mKey
End Property

Private Property Get IBracketOrder_LongPosition() As Boolean
IBracketOrder_LongPosition = LongPosition
End Property

Private Property Get IBracketOrder_Order(ByVal pIndex As Long) As IOrder
Const ProcName As String = "IBracketOrder_Order"
On Error GoTo Err

Set IBracketOrder_Order = Order(pIndex)

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

Private Sub IBracketOrder_RemoveChangeListener(ByVal pListener As ChangeListener)
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

Private Sub IBracketOrder_RemoveStateChangeListener(ByVal pListener As TWUtilities40.StateChangeListener)
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

Private Property Get IBracketOrder_StopReason() As String
IBracketOrder_StopReason = StopReason
End Property

Private Property Get IBracketOrder_TargetOrder() As IOrder
Set IBracketOrder_TargetOrder = mTargetOrder
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
' IPositionMember Interface Members
'@================================================================================

Private Property Get IPositionMember_PositionManagerName() As String
IPositionMember_PositionManagerName = mContextsName
End Property

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

Private Sub IRecoverable_FinishRecovery()
Const ProcName As String = "IRecoverable_FinishRecovery"
On Error GoTo Err

If mCancelPrice <> 0# Then
    Dim lPriceCanceller As New BracketPriceCanceller
    lPriceCanceller.Initialise Me, mCancelPrice, mDataSource
End If

If mCancelTime <> 0 And mCancelTime > Now Then Set mCancelTLI = mTimerList.Add(Nothing, mCancelTime, ExpiryTimeUnits.ExpiryTimeUnitDateTime)
    
CheckComplete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IRecoverable_LoadRecoveryData(ByVal pData As Variant)
Const ProcName As String = "IRecoverable_LoadRecoveryData"
On Error GoTo Err

Set mSummary = pData

mCancelPrice = mSummary.CancelPrice
mCancelTime = mSummary.CancelTime
mCreationTime = mSummary.CreationTime
CumBuyPrice = mSummary.CumBuyPrice
CumSellPrice = mSummary.CumSellPrice
mContextName = mSummary.ContextName
mDescription = mSummary.Description
mEntryReason = mSummary.EntryReason
mGroupName = mSummary.GroupName
mKey = mSummary.Key
mPreventUnprotectedPosition = mSummary.PreventUnprotectedPosition
mSize = mSummary.Size
State = mSummary.State
mStopReason = mSummary.StopReason
mTargetReason = mSummary.TargetReason

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
' Stringable Interface Members
'@================================================================================

Private Function Stringable_ToString() As String
Stringable_ToString = ToString
End Function

'@================================================================================
' mCloseoutOrder Event Handlers
'@================================================================================

Private Sub mCloseoutOrder_Fill(ByVal pExec As IExecutionReport)
Const ProcName As String = "mCloseoutOrder_Fill"
On Error GoTo Err

fireChange BracketOrderCloseoutOrderChanged

processExecution pExec

logOrderMessage "Closeout order fill: (id=" & _
                    mCloseoutOrder.Id & "): " & _
                    gOrderActionToString(mCloseoutOrder.Action) & " " & pExec.Quantity & " at " & _
                    pExec.Price

RaiseEvent CloseoutOrderPartiallyFilled

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mCloseoutOrder_PropertyChanged()
Const ProcName As String = "mCloseoutOrder_PropertyChanged"
On Error GoTo Err

fireChange BracketOrderCloseoutOrderChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mCloseoutOrder_StatusChanged()
Const ProcName As String = "mCloseoutOrder_StatusChanged"
On Error GoTo Err

fireChange BracketOrderCloseoutOrderChanged
Select Case mCloseoutOrder.Status
    Case OrderStatuses.OrderStatusFilled
        logOrderMessage "Closeout order completed (id=" & _
                            mCloseoutOrder.Id & "): " & _
                            gOrderActionToString(mCloseoutOrder.Action) & " " & mCloseoutOrder.QuantityFilled & " at " & _
                            mCloseoutOrder.AveragePrice
                            
        RaiseEvent CloseoutOrderFilled

        checkAllOrdersComplete
        
        Assert Size = 0, "Closeout incomplete: Size remaining is " & Size
        
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusSubmitted
    Case OrderStatuses.OrderStatusCancelled
        logOrderMessage "Closeout order cancelled (id=" & _
                            mCloseoutOrder.Id & ")"
    Case OrderStatuses.OrderStatusRejected
        logOrderMessage "Closeout order rejected (id=" & _
                            mCloseoutOrder.Id & ")"
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

Private Sub mEntryOrder_Fill(ByVal pExec As IExecutionReport)
Const ProcName As String = "mEntryOrder_Fill"
On Error GoTo Err

fireChange BracketOrderEntryOrderChanged

processExecution pExec

mCancelPrice = 0

logOrderMessage "Entry order fill (id=" & _
                    mEntryOrder.Id & "): " & _
                    gOrderActionToString(mEntryOrder.Action) & " " & pExec.Quantity & " at " & _
                    pExec.Price

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

placePeggedOrders
        
If Not mStopLossOrder Is Nothing Then
    If mStopLossOrder.Quantity <> Abs(Size) And _
        Abs(Size) <> 0 _
    Then
        AdjustStop mStopLossOrder.TriggerPrice, Abs(Size)
    End If
End If

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Quantity <> Abs(Size) And _
        Abs(Size) <> 0 _
    Then
        AdjustTarget mTargetOrder.LimitPrice, Abs(Size)
    End If
End If

RaiseEvent EntryOrderPartiallyFilled

handleStimulus OpStimuli.StimEntryOrderFill

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mEntryOrder_PropertyChanged()
Const ProcName As String = "mEntryOrder_PropertyChanged"
On Error GoTo Err

fireChange BracketOrderEntryOrderChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mEntryOrder_StatusChanged()
Const ProcName As String = "mEntryOrder_StatusChanged"
On Error GoTo Err

fireChange BracketOrderEntryOrderChanged
Select Case mEntryOrder.Status
    Case OrderStatuses.OrderStatusFilled
        mCancelPrice = 0
        
        logOrderMessage "Entry order completed (id=" & _
                            mEntryOrder.Id & "): " & _
                            gOrderActionToString(mEntryOrder.Action) & " " & mEntryOrder.QuantityFilled & " at " & _
                            mEntryOrder.AveragePrice
        
        RaiseEvent EntryOrderFilled
        
        checkAllOrdersComplete
    
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusSubmitted
        
        If mEntryOrder.QuantityFilled > 0 Then
            placePeggedOrders
        End If
    Case OrderStatuses.OrderStatusCancelled
        logOrderMessage "Entry order cancelled (id=" & mEntryOrder.Id & ")"
        handleStimulus OpStimuli.StimEntryOrderCancelled
        checkAllOrdersComplete
    Case OrderStatuses.OrderStatusRejected
        logOrderMessage "Entry order rejected (id=" & mEntryOrder.Id & ")"
        handleStimulus OpStimuli.StimEntryOrderCancelled
        checkAllOrdersComplete
End Select

SaveRecoveryInfo
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

SaveRecoveryInfo

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

Private Sub mStopLossOrder_Fill(ByVal pExec As IExecutionReport)
Const ProcName As String = "mStopLossOrder_Fill"
On Error GoTo Err

fireChange BracketOrderStopOrderChanged

processExecution pExec

logOrderMessage "Stop order fill (id=" & _
                        mStopLossOrder.Id & "): " & _
                        gOrderActionToString(mStopLossOrder.Action) & " " & pExec.Quantity & " at " & _
                        pExec.Price

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Quantity <> Abs(Size) And _
        Abs(Size) <> 0 _
    Then
        AdjustTarget mTargetOrder.LimitPrice, Abs(Size)
    End If
End If

RaiseEvent StopOrderPartiallyFilled

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mStopLossOrder_PropertyChanged()
Const ProcName As String = "mStopLossOrder_PropertyChanged"
On Error GoTo Err

fireChange BracketOrderStopOrderChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mStopLossOrder_StatusChanged()
Const ProcName As String = "mStopLossOrder_StatusChanged"
On Error GoTo Err

fireChange BracketOrderStopOrderChanged
Select Case mStopLossOrder.Status
    Case OrderStatuses.OrderStatusFilled
        logOrderMessage "Stop order completed (id=" & _
                            mStopLossOrder.Id & "): " & _
                            gOrderActionToString(mStopLossOrder.Action) & " " & mStopLossOrder.QuantityFilled & " at " & _
                            mStopLossOrder.AveragePrice
        
        RaiseEvent StopOrderFilled
        
        checkAllOrdersComplete
    
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusSubmitted
    Case OrderStatuses.OrderStatusCancelled
        logOrderMessage "Stop order cancelled (id=" & mStopLossOrder.Id & ")"
        handleStimulus OpStimuli.StimStopOrderCancelled
        checkAllOrdersComplete
    Case OrderStatuses.OrderStatusRejected
        logOrderMessage "Stop order rejected (id=" & mStopLossOrder.Id & ")"
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

Private Sub mTargetOrder_Fill(ByVal pExec As IExecutionReport)
Const ProcName As String = "mTargetOrder_Fill"
On Error GoTo Err

fireChange BracketOrderTargetOrderChanged

processExecution pExec

logOrderMessage "Target order fill (id=" & _
                        mTargetOrder.Id & "): " & _
                        gOrderActionToString(mTargetOrder.Action) & " " & pExec.Quantity & " at " & _
                        pExec.Price

If Not mStopLossOrder Is Nothing Then
    If mStopLossOrder.Quantity <> Abs(Size) And _
        Abs(Size) <> 0 _
    Then
        AdjustStop mStopLossOrder.TriggerPrice, Abs(Size)
    End If
End If

RaiseEvent TargetOrderPartiallyFilled

SaveRecoveryInfo
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mTargetOrder_PropertyChanged()
Const ProcName As String = "mTargetOrder_PropertyChanged"
On Error GoTo Err

fireChange BracketOrderTargetOrderChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mTargetOrder_StatusChanged()
Const ProcName As String = "mTargetOrder_StatusChanged"
On Error GoTo Err

fireChange BracketOrderTargetOrderChanged
Select Case mTargetOrder.Status
    Case OrderStatuses.OrderStatusFilled
        logOrderMessage "Target order completed (id=" & _
                            mTargetOrder.Id & "): " & _
                            gOrderActionToString(mTargetOrder.Action) & " " & mTargetOrder.QuantityFilled & " at " & _
                            mTargetOrder.AveragePrice
        
        RaiseEvent TargetOrderFilled
        
        checkAllOrdersComplete
    
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusSubmitted
    Case OrderStatuses.OrderStatusCancelled
        logOrderMessage "Target order cancelled (id=" & mTargetOrder.Id & ")"
        handleStimulus OpStimuli.StimTargetOrderCancelled
        checkAllOrdersComplete
    Case OrderStatuses.OrderStatusRejected
        logOrderMessage "Target order rejected (id=" & mTargetOrder.Id & ")"
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

logOrderMessage "Cancel time expired"
Cancel
fireChange BracketOrderSelfCancelled
RaiseEvent SelfCancelled

SaveRecoveryInfo
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

mCancelAfter = Value
SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let CancelPrice(ByVal Value As Double)
Const ProcName As String = "CancelPrice"
On Error GoTo Err

Assert State = BracketOrderStateCreated, "CancelPrice cannot be set after execution"

mCancelPrice = Value
mSummary.CancelPrice = mCancelPrice
SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CancelPrice() As Double
CancelPrice = mCancelPrice
End Property

Public Property Let CancelTime(ByVal Value As Date)
Const ProcName As String = "CancelTime"
On Error GoTo Err

Assert State = BracketOrderStateCreated, "CancelTime cannot be set after execution"

mCancelTime = Value
mSummary.CancelTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CancelTime() As Date
CancelTime = mCancelTime
End Property

Friend Property Let CloseoutOrder(ByVal Value As IOrder)
Const ProcName As String = "CloseoutOrder"
On Error GoTo Err

Set mCloseoutOrder = Value
mSummary.CloseoutOrderId = mCloseoutOrder.Id
mOrders.Add mCloseoutOrder

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

mCreationTime = Value
mSummary.CreationTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CreationTime() As Date
CreationTime = mCreationTime
End Property

Friend Property Let CumBuyPrice(ByVal Value As Double)
mCumBuyPrice = Value
mSummary.CumBuyPrice = Value
End Property

Public Property Get CumBuyPrice() As Double
CumBuyPrice = mCumBuyPrice
End Property

Friend Property Let CumSellPrice(ByVal Value As Double)
Const ProcName As String = "CumSellPrice"
On Error GoTo Err

mCumSellPrice = Value
mSummary.CumSellPrice = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CumSellPrice() As Double
CumSellPrice = mCumSellPrice
End Property

Public Property Let Description(ByVal Value As String)
Const ProcName As String = "Description"
On Error GoTo Err

If mDescription = Value Then Exit Property
mDescription = Value
mSummary.Description = Value
SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Friend Property Let EntryOrder(ByVal Value As IOrder)
Const ProcName As String = "EntryOrder"
On Error GoTo Err

Assert mEntryOrder Is Nothing, "Entry order already exists"

Set mEntryOrder = Value
mSummary.EntryOrderId = mEntryOrder.Id

If mOrders.Count >= 1 Then
    mOrders.Add mEntryOrder, , 1
Else
    mOrders.Add mEntryOrder
End If

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
    EntryPrice = mCumBuyPrice / mEntryOrder.Quantity
Else
    EntryPrice = mCumSellPrice / mEntryOrder.Quantity
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let EntryReason(ByVal Value As String)
Const ProcName As String = "EntryReason"
On Error GoTo Err

If mEntryReason = Value Then Exit Property
mEntryReason = Value
mSummary.EntryReason = Value
SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get EntryReason() As String
EntryReason = mEntryReason
End Property

Public Property Get ExitPrice() As Double
Const ProcName As String = "ExitPrice"
On Error GoTo Err

If LongPosition Then
    ExitPrice = mCumSellPrice / mEntryOrder.Quantity
Else
    ExitPrice = mCumBuyPrice / mEntryOrder.Quantity
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
Key = mKey
End Property

Public Property Get LongPosition() As Boolean
Const ProcName As String = "LongPosition"
On Error GoTo Err

If Not mEntryOrder Is Nothing Then LongPosition = (mEntryOrder.Action = OrderActionBuy)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Order(ByVal pIndex As Long) As IOrder
Const ProcName As String = "Order"
On Error GoTo Err

Set Order = mOrders(pIndex)

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

NumberOfOrders = mOrders.Count

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

'Public Property Get PositionManager() As PositionManager
'Set PositionManager = mPositionManagerRef.Target
'End Property

Public Property Let PreventUnprotectedPosition( _
                ByVal Value As Boolean)
Const ProcName As String = "PreventUnprotectedPosition"
On Error GoTo Err

If mPreventUnprotectedPosition = Value Then Exit Property

Assert (Not Value) Or (State = BracketOrderStateCreated) Or (Not mStopLossOrder Is Nothing), "No stop loss order exists"

mPreventUnprotectedPosition = Value
mSummary.PreventUnprotectedPosition = Value
SaveRecoveryInfo

If mPreventUnprotectedPosition Then
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

PreventUnprotectedPosition = mPreventUnprotectedPosition

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

Public Property Get Size() As Long
Size = mSize
End Property

Public Property Get StartTime() As Date
StartTime = mStartTime
End Property

Public Property Get EndTime() As Date
EndTime = mEndTime
End Property

Friend Property Let State(ByVal Value As BracketOrderStates)
Const ProcName As String = "State"
On Error GoTo Err

Dim lConditions As OpConditions
If Not mStateEngine Is Nothing Then lConditions = mStateEngine.Conditions

Set mStateEngine = CreateStateEngine(GBracketOrder.TableBuilder, _
                                    Value)
mStateEngine.SetConditions lConditions

mSummary.State = Value
If IsComplete Then doCompletionActions

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

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

Assert mStopLossOrder Is Nothing, "Stop loss order already exists"

Set mStopLossOrder = Value
mSummary.StopLossOrderId = mStopLossOrder.Id

mStateEngine.SetConditions OpConditions.CondStopOrderExists
mStateEngine.ClearConditions OpConditions.CondStopOrderCancelled

If Not mEntryOrder Is Nothing Then
    mOrders.Add mStopLossOrder, , , 1
ElseIf Not mTargetOrder Is Nothing Then
    mOrders.Add mStopLossOrder, , 1
Else
    mOrders.Add mStopLossOrder
End If

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

If mStopReason = Value Then Exit Property
mStopReason = Value
mSummary.StopReason = Value
SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StopReason() As String
StopReason = mStopReason
End Property

Friend Property Let TargetOrder(ByVal Value As IOrder)
Const ProcName As String = "TargetOrder"
On Error GoTo Err

Assert mTargetOrder Is Nothing, "Target order already exists"

Set mTargetOrder = Value
mSummary.TargetOrderId = mTargetOrder.Id

mStateEngine.SetConditions OpConditions.CondTargetOrderExists
mStateEngine.ClearConditions OpConditions.CondTargetOrderCancelled

If mEntryOrder Is Nothing And _
    mStopLossOrder Is Nothing _
Then
    mOrders.Add mTargetOrder
ElseIf mEntryOrder Is Nothing Then
    mOrders.Add mTargetOrder, , , 1
ElseIf mStopLossOrder Is Nothing Then
    mOrders.Add mTargetOrder, , , 1
Else
    mOrders.Add mTargetOrder, , , 2
End If

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

If mTargetReason = Value Then Exit Property
mTargetReason = Value
mSummary.TargetReason = Value
SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TargetReason() As String
TargetReason = mTargetReason
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
If Not mEntryOrder Is Nothing Then fireChange BracketOrderEntryOrderChanged
If Not mStopLossOrder Is Nothing Then fireChange BracketOrderStopOrderChanged
If Not mTargetOrder Is Nothing Then fireChange BracketOrderTargetOrderChanged
If Not mCloseoutOrder Is Nothing Then fireChange BracketOrderCloseoutOrderChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddChangeListener(ByVal pListener As ChangeListener)
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

Public Sub AddStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrElementAlreadyExists Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AdjustStop(ByVal stopTriggerPrice As Double, _
                    Optional ByVal Quantity As Long = -1)
Const ProcName As String = "AdjustStop"
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

mCancelPrice = 0
If Not mCancelTLI Is Nothing Then
    If mCancelTLI.State = TimerListItemStatePending Then
        mCancelTLI.Cancel
        Set mCancelTLI = Nothing
    End If
End If

If evenIfFilled Then
    mStateEngine.ClearConditions OpConditions.CondNoFillCancellation
    Cancel = handleStimulus(OpStimuli.StimCancelEvenIfFill)
Else
    mStateEngine.SetConditions OpConditions.CondNoFillCancellation
    Cancel = handleStimulus(OpStimuli.StimCancelIfNoFill)
End If

SaveRecoveryInfo
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
RaiseEvent Clean

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub CheckComplete()
Const ProcName As String = "CheckComplete"
On Error GoTo Err

checkAllOrdersComplete
SaveRecoveryInfo

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

logOrderMessage "Closing out: " & _
                    "size=" & Size & _
                    "; pending size=" & PendingSize

handleStimulus OpStimuli.StimCloseout

SaveRecoveryInfo
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ContainsOrder( _
                ByVal pOrder As IOrder) As Boolean
Const ProcName As String = "ContainsOrder"
On Error GoTo Err

ContainsOrder = mOrders.Contains(pOrder)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Execute()
Const ProcName As String = "Execute"
On Error GoTo Err

Assert (Not mPreventUnprotectedPosition) Or (Not mStopLossOrder Is Nothing), "No stop loss order exists"

mCreationTime = GetTimestamp
mSummary.CreationTime = mCreationTime

If mCancelAfter <> 0 Then
    mCancelTime = mTimerList.Clock.Timestamp + CDbl(mCancelAfter) / 86400#
    mCancelAfter = 0
    mSummary.CancelTime = mCancelTime
End If

If mCancelTime <> 0 Then Set mCancelTLI = mTimerList.Add(Nothing, mCancelTime, ExpiryTimeUnits.ExpiryTimeUnitDateTime)
    
If mCancelPrice <> 0# Then
    Dim lPriceCanceller As New BracketPriceCanceller
    lPriceCanceller.Initialise Me, mCancelPrice, mDataSource
End If

fireChange BracketOrderChangeTypes.BracketOrderCreated

SaveRecoveryInfo    ' ensure the BracketOrder is recorded before
                    ' doing the state change, which will cause
                    ' the orders to be recorded
                    
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
                ByVal pContextsName As String, _
                ByVal pGroupName As String, _
                ByVal pOrderContext As OrderContext, _
                ByVal pContract As IContract, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pIsSimulated As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

mKey = pKey
mContextsName = pContextsName
mGroupName = pGroupName
Set mOrderContext = pOrderContext
mContextName = mOrderContext.Name
mSummary.ContextName = mContextName
Set mContract = pContract
Set mDataSource = pDataSource
Set mOrderSubmitter = pOrderSubmitter
mIsSimulated = pIsSimulated

Set mTimerList = mOrderContext.TimerList

mSummary.GroupName = mGroupName
mSummary.Key = mKey
mSummary.ContextsName = mContextsName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Friend Sub NotifyCloseout( _
'                ByVal pPrice As Double, _
'                ByVal pTimestamp As Date)
'Const ProcName As String = "NotifyCloseout"
'On Error GoTo Err
'
'If Size <> 0 Then processExecution -Size, pPrice, pTimestamp, mTimezoneName, True
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError  ProcName, ModuleName
'End Sub

Public Sub RemoveChangeListener(ByVal pListener As ChangeListener)
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

Public Sub RemoveStateChangeListener(ByVal pListener As StateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SelfCancel(ByVal pReason As String)
logOrderMessage pReason
Cancel
fireChange BracketOrderSelfCancelled
RaiseEvent SelfCancelled
End Sub

Public Sub SetNewEntryPrice(ByVal Value As Double)
Const ProcName As String = "SetNewEntryPrice"
On Error GoTo Err

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
    logOrderMessage "Set entry price=" & priceToString(Value)
    mNewEntryPrice = Value
    setDirty
    fireChange BracketOrderChangeTypes.BracketOrderChangesPending
    RaiseEvent Dirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewEntryTriggerPrice(ByVal Value As Double)
Const ProcName As String = "SetNewEntryTriggerPrice"
On Error GoTo Err

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
    logOrderMessage "Set entry trigger price=" & priceToString(Value)
    mNewEntryTriggerPrice = Value
    setDirty
    fireChange BracketOrderChangeTypes.BracketOrderChangesPending
    RaiseEvent Dirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewOrderPrice( _
                ByVal orderIndex As Long, _
                ByVal Value As Double)
Const ProcName As String = "SetNewOrderPrice"
On Error GoTo Err

Dim lOrder As IOrder
Set lOrder = mOrders(orderIndex)

If Not mEntryOrder Is Nothing Then
    If lOrder Is mEntryOrder Then
        SetNewEntryPrice Value
        Exit Sub
    End If
End If

If Not mStopLossOrder Is Nothing Then
    If lOrder Is mStopLossOrder Then
        SetNewStopLossPrice Value
        Exit Sub
    End If
End If

If Not mTargetOrder Is Nothing Then
    If lOrder Is mTargetOrder Then
        SetNewTargetPrice Value
        Exit Sub
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewOrderTriggerPrice( _
                ByVal orderIndex As Long, _
                ByVal Value As Double)
Const ProcName As String = "SetNewOrderTriggerPrice"
On Error GoTo Err

Dim lOrder As IOrder
Set lOrder = mOrders(orderIndex)

If Not mEntryOrder Is Nothing Then
    If lOrder Is mEntryOrder Then
        SetNewEntryTriggerPrice Value
        Exit Sub
    End If
End If

If Not mStopLossOrder Is Nothing Then
    If lOrder Is mStopLossOrder Then
        SetNewStopLossTriggerPrice Value
        Exit Sub
    End If
End If

If Not mTargetOrder Is Nothing Then
    If lOrder Is mTargetOrder Then
        SetNewTargetTriggerPrice Value
        Exit Sub
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewQuantity(ByVal Value As Long)
Const ProcName As String = "SetNewQuantity"
On Error GoTo Err

Select Case mStateEngine.State
Case BracketOrderStates.BracketOrderStateCancelling
    Exit Sub
Case BracketOrderStates.BracketOrderStateSubmitted
Case Else
    Assert False, "Action not allowed in state " & mStateEngine.State
End Select

Assert mEntryOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change quantity - order " & EntryOrder.Id & " already filled"

If Value <> mEntryOrder.Quantity Then
    logOrderMessage "Set quantity=" & Value
    mNewQuantity = Value
    setDirty
    fireChange BracketOrderChangeTypes.BracketOrderChangesPending
    RaiseEvent Dirty
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

Assert Not mStopLossOrder Is Nothing, "No stop loss order exists"
Assert mStopLossOrder.OrderType = OrderTypes.OrderTypeAutoStop, "Can't change offset for order type " & gOrderTypeToString(mEntryOrder.OrderType)
Assert mStopLossOrder.Status = OrderStatuses.OrderStatusCreated, "Can't change stop offset - order " & mStopLossOrder.Id & " already placed"

If Value <> mStopLossOrder.Offset Then
    logOrderMessage "Set stop offset=" & Value
    mNewStopLossOffset = Value
    setDirty
    fireChange BracketOrderChangeTypes.BracketOrderChangesPending
    RaiseEvent Dirty
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

Assert Not mStopLossOrder Is Nothing, "No stop loss order exists"

Select Case mStopLossOrder.OrderType
Case OrderTypes.OrderTypeStopLimit
Case Else
    Assert False, "Can't change Price for order type " & gOrderTypeToString(mStopLossOrder.OrderType)
End Select

If Value <> mStopLossOrder.LimitPrice Then
    logOrderMessage "Set stop price=" & priceToString(Value)
    mNewStopLossPrice = Value
    setDirty
    fireChange BracketOrderChangeTypes.BracketOrderChangesPending
    RaiseEvent Dirty
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

Assert Not mStopLossOrder Is Nothing, "No stop loss order exists"

If Value <> mStopLossOrder.Quantity Then
    logOrderMessage "Set stop quantity=" & Value
    mNewStopLossQuantity = Value
    setDirty
    fireChange BracketOrderChangeTypes.BracketOrderChangesPending
    RaiseEvent Dirty
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

Assert Not mStopLossOrder Is Nothing, "No stop loss order exists"

Select Case mStopLossOrder.OrderType
    Case OrderTypes.OrderTypeStop, _
        OrderTypes.OrderTypeStopLimit
    Case Else
        Assert False, "Can't change trigger Price for order type " & gOrderTypeToString(mStopLossOrder.OrderType)
End Select

If Value <> mStopLossOrder.TriggerPrice Then
    logOrderMessage "Set stop trigger price=" & priceToString(Value)
    mNewStopLossTriggerPrice = Value
    setDirty
    fireChange BracketOrderChangeTypes.BracketOrderChangesPending
    RaiseEvent Dirty
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
    logOrderMessage "Set target price=" & priceToString(Value)
    mNewTargetPrice = Value
    setDirty
    fireChange BracketOrderChangeTypes.BracketOrderChangesPending
    RaiseEvent Dirty
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

Assert Not mTargetOrder Is Nothing, "BracketOrder::(let)NewTargetOffset", "No target order exists"
Assert mTargetOrder.OrderType = OrderTypes.OrderTypeAutoLimit, "Can't change offset for order type " & gOrderTypeToString(mTargetOrder.OrderType)
Assert mTargetOrder.Status = OrderStatuses.OrderStatusCreated, "Can't change target offset - order " & mTargetOrder.Id & " already placed"

If Value <> mTargetOrder.Offset Then
    logOrderMessage "Set target offset=" & Value
    mNewTargetOffset = Value
    setDirty
    fireChange BracketOrderChangeTypes.BracketOrderChangesPending
    RaiseEvent Dirty
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

Assert Not mStopLossOrder Is Nothing, "No target order exists"

If Value <> mTargetOrder.Quantity Then
    logOrderMessage "Set target quantity=" & Value
    mNewTargetQuantity = Value
    setDirty
    fireChange BracketOrderChangeTypes.BracketOrderChangesPending
    RaiseEvent Dirty
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
    logOrderMessage "Set target trigger price=" & priceToString(Value)
    mNewTargetTriggerPrice = Value
    setDirty
    fireChange BracketOrderChangeTypes.BracketOrderChangesPending
    RaiseEvent Dirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ToString() As String
Const ProcName As String = "ToString"
On Error GoTo Err

ToString = gOrderActionToString(mEntryOrder.Action) & " " & mEntryOrder.Quantity & " at " & _
        gOrderTypeToString(mEntryOrder.OrderType) & " " & _
        IIf(mEntryOrder.TriggerPrice <> 0, "(" & mEntryOrder.TriggerPrice & ")", "")
If Not mStopLossOrder Is Nothing Then
    ToString = ToString & _
        ", stop " & IIf(mStopLossOrder.OrderType = OrderTypeAutoStop, "AUTO(" & mStopLossOrder.Offset * mContract.TickSize & ")", mStopLossOrder.TriggerPrice)
End If
If Not mTargetOrder Is Nothing Then
    ToString = ToString & _
        ", target " & IIf(mTargetOrder.OrderType = OrderTypeAutoLimit, "AUTO(" & mTargetOrder.Offset * mContract.TickSize & ")", mTargetOrder.LimitPrice)
End If
ToString = ToString & "(id=" & mKey & ")"

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

logOrderMessage "Update bracket order"

If State <> BracketOrderStateCreated Then
    mOrderSubmitter.ModifyBracketOrder Me, updateEntry, updateStop, updateTarget, False
End If

Checkpoint

If updateEntry Then
    fireChange BracketOrderEntryOrderChanged
End If
If updateStop Then
    fireChange BracketOrderStopOrderChanged
    RaiseEvent StopAdjusted
End If
If updateTarget Then
    fireChange BracketOrderTargetOrderChanged
    RaiseEvent TargetAdjusted
End If
fireChange BracketOrderChangeTypes.BracketOrderChangesApplied
RaiseEvent Clean

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

mCancelPrice = 0

Select Case mEntryOrder.Status
Case OrderStatuses.OrderStatusCreated
    CancelAfter = 0
    CancelTime = 0
    mEntryOrder.Status = OrderStatuses.OrderStatusCancelled

    cancelOrders = True
End Select

If Not mStopLossOrder Is Nothing Then
    Select Case mStopLossOrder.Status
    Case OrderStatuses.OrderStatusCreated
        mStopLossOrder.Status = OrderStatuses.OrderStatusCancelled
        cancelOrders = True
    End Select
End If

If Not mTargetOrder Is Nothing Then
    Select Case mTargetOrder.Status
    Case OrderStatuses.OrderStatusCreated
        mTargetOrder.Status = OrderStatuses.OrderStatusCancelled
        cancelOrders = True
    End Select
End If

If mOrderSubmitter.CancelBracketOrder(Me) Then
    cancelOrders = True
    
    Dim StopLossOrderId As String
    If Not mStopLossOrder Is Nothing Then StopLossOrderId = mStopLossOrder.Id
    
    Dim TargetOrderId As String
    If Not mTargetOrder Is Nothing Then TargetOrderId = mTargetOrder.Id
    logOrderMessage "Cancel orders: " & _
                    EntryOrder.Id & _
                    ", " & StopLossOrderId & _
                    ", " & TargetOrderId
End If

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
    logOrderMessage "Cancel order: id=" & mStopLossOrder.Id
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
    logOrderMessage "Cancel order: id=" & mTargetOrder.Id
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
        Assert mEntryOrder.Status <> OrderStatuses.OrderStatusFilled, "Can't change Price - order " & mEntryOrder.Id & " already filled"
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
Set lOrder = mOrderContext.CreateRawOrder(pId)

lOrder.Initialise mGroupName, mContextsName, mContract.Specifier, mOrderContext
Set CreateAnOrder = lOrder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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
        gLog "Stimulus " & pStimulus & " occurred in state " & pPreviousState, ProcName, ModuleName, , LogLevelSevere
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

If Not mOrderContext Is Nothing Then mOrderContext.NotifyBracketOrderClosed Me
fireChange BracketOrderCompleted
RaiseEvent Completed

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fireChange( _
                ByVal pChangeType As BracketOrderChangeTypes)
Const ProcName As String = "fireChange"
On Error GoTo Err

Dim ev As ChangeEventData
Set ev.Source = Me
ev.changeType = pChangeType

mChangeListeners.SetCurrentListeners

Dim i As Long
For i = 1 To mChangeListeners.Count
    Dim lListener As ChangeListener
    Set lListener = mChangeListeners.GetListener(i)
    lListener.Change ev
Next

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

mExecutionListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mExecutionListeners.Count
    Dim lListener As IExecutionListener
    Set lListener = mExecutionListeners.GetListener(i)
    lListener.NotifyExecution ev
Next

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

mStateChangeListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mStateChangeListeners.Count
    Dim lListener As StateChangeListener
    Set lListener = mStateChangeListeners.GetListener(i)
    lListener.Change ev
Next

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

mRecoverableObjectListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mRecoverableObjectListeners.Count
    Dim lListener As IRecoverableObjectListener
    Set lListener = mRecoverableObjectListeners.GetListener(i)
    lListener.NotifyRecoveryDataChanged ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function handleStimulus(ByVal pStimulus As OpStimuli) As Variant
Const ProcName As String = "handleStimulus"
On Error GoTo Err

Dim lPrevState As BracketOrderStates
lPrevState = mStateEngine.State

Dim lActions() As OpActions
lActions = mStateEngine.NotifyStimulus(pStimulus)
mSummary.State = State
SaveRecoveryInfo

Dim lConditions As OpConditions
lConditions = mStateEngine.Conditions

Dim lActionsString As String
Dim i As Long
For i = 0 To UBound(lActions)
    lActionsString = lActionsString & gOpActionsToString(lActions(i)) & IIf(i = UBound(lActions), "", ",")
Next

gLog "Handling bracket order stimulus: prior state=" & gBracketOrderStatesToString(lPrevState) & _
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

Private Sub logOrderMessage( _
                ByVal pMessage As String)
Const ProcName As String = "logOrderMessage"
On Error GoTo Err

Dim lTickPart As String

If mDataSource Is Nothing Then
ElseIf mDataSource.State <> MarketDataSourceStateRunning Then
ElseIf mDataSource.CurrentTick(TickTypeTrade).Timestamp = 0# Then
Else
    lTickPart = "Curr price=" & priceToString(mDataSource.CurrentTick(TickTypeTrade).Price) & "; "
End If

gLogOrder pMessage & " (" & lTickPart & "id=" & Key & ")", _
        mIsSimulated, _
        Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function placeCloseoutOrder()
Const ProcName As String = "placeCloseoutOrder"
On Error GoTo Err

Dim lOrder As IOrder
Set lOrder = mOrderContext.NewOrder
CloseoutOrder = lOrder

mCloseoutOrder.Action = IIf(LongPosition, OrderActions.OrderActionSell, OrderActions.OrderActionBuy)
mCloseoutOrder.Quantity = Abs(Size)
mCloseoutOrder.OrderType = OrderTypes.OrderTypeMarket
mCloseoutOrder.TimeInForce = OrderTIFs.OrderTIFGoodTillCancelled

fireChange BracketOrderCloseoutOrderCreated
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

Private Sub placePeggedOrders()
Const ProcName As String = "placePeggedOrders"
On Error GoTo Err

Dim updated As Boolean

If Not mStopLossOrder Is Nothing Then
    If mStopLossOrder.Status = OrderStatuses.OrderStatusCreated Then
        If mStopLossOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
            ' the stop Price is to be pegged to the entry Price - the
            ' order's offset property will already contain any required
            ' offset from that Price
            mStopLossOrder.OrderType = OrderTypes.OrderTypeStop
            SetNewStopLossTriggerPrice (mStopLossOrder.Offset * mContract.TickSize) + EntryOrder.LastFillPrice
            SetNewStopLossQuantity mEntryOrder.QuantityFilled
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
            SetNewTargetQuantity mEntryOrder.QuantityFilled
            updated = True
        End If
    End If
End If

If updated Then Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function priceToString(ByVal pPrice As Double) As String
Const ProcName As String = "priceToString"
On Error GoTo Err

priceToString = FormatPrice(pPrice, mContract.Specifier.SecType, mContract.TickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processExecution( _
                ByVal pExec As IExecutionReport)
Const ProcName As String = "processExecution"
On Error GoTo Err

If mSize = 0 Then
    mStartTime = pExec.FillTime
    mTimezoneName = pExec.TimezoneName
End If

If pExec.Action = OrderActionBuy Then
    CumBuyPrice = mCumBuyPrice + pExec.Quantity * pExec.Price
    setSize Size + pExec.Quantity
Else
    CumSellPrice = mCumSellPrice + pExec.Quantity * pExec.Price
    setSize Size - pExec.Quantity
End If
    
If Size = 0 Then mEndTime = pExec.FillTime

fireExecution pExec
fireChange BracketOrderChangeTypes.BracketOrderSizeChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function ResubmitStopOrder()
Const ProcName As String = "resubmitStopOrder"
On Error GoTo Err

mOrderSubmitter.ResubmitStopOrder Me
logOrderMessage "Resubmit stop order: id=" & mStopLossOrder.Id & _
            "; new broker id=" & mStopLossOrder.BrokerId
mStateEngine.ClearConditions OpConditions.CondStopOrderCancelled

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function ResubmitStopAndTargetOrders()
Const ProcName As String = "resubmitStopAndTargetOrders"
On Error GoTo Err

mOrderSubmitter.ResubmitStopAndTargetOrders Me
logOrderMessage "Resubmit stop order: id=" & mStopLossOrder.Id & _
            "; new broker id=" & mStopLossOrder.BrokerId
logOrderMessage "Resubmit target order: id=" & mTargetOrder.Id & _
            "; new broker id=" & mTargetOrder.BrokerId
mStateEngine.ClearConditions OpConditions.CondStopOrderCancelled
mStateEngine.ClearConditions OpConditions.CondTargetOrderCancelled

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Function ResubmitTargetOrder()
Const ProcName As String = "resubmitTargetOrder"
On Error GoTo Err

mOrderSubmitter.ResubmitTargetOrder Me
logOrderMessage "Resubmit target order: id=" & mTargetOrder.Id & _
            "; new broker id=" & mTargetOrder.BrokerId
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
' Set the 'dirty' flag, and raise a Dirty event if one has not already
' been raised.
'@/
Private Sub setDirty()
Const ProcName As String = "setDirty"
On Error GoTo Err

If Not mIsDirty Then
    mIsDirty = True
    RaiseEvent Dirty
Else
    mIsDirty = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSize(ByVal Value As Long)
Const ProcName As String = "setSize"
On Error GoTo Err

mSize = Value
mSummary.Size = Value

If mSize <> 0 Then
    mStateEngine.SetConditions OpConditions.CondSizeNonZero
Else
    mStateEngine.ClearConditions OpConditions.CondSizeNonZero
End If

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

