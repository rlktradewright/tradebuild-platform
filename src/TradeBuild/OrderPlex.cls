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

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements QuoteListener
Implements IOrderPlex

'@================================================================================
' Events
'@================================================================================

Event Change(ByRef ev As ChangeEventData)
Event StateChange(ByRef ef As StateChangeEventData)
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

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "OrderPlex"

'@================================================================================
' Enums
'@================================================================================

Public Enum OrderPlexStateCodes
    
    ' This state indicates that the order plex has been created but none of
    ' the orders have yet been placed.
    OrderPlexStateCreated = 1
    
    ' This state indicates that all the orders in the order plex have been placed.
    OrderPlexStateSubmitted
    
    ' This state indicates that the order plex is in the process of being
    ' cancelled.
    OrderPlexStateCancelling
    
    ' This state indicates that the order plex is being closed out.
    ' (When an order plex has been closed out, there are no outstanding
    ' orders and no net Size - note that this may require a closeout
    ' order to be placed to negate any existing Size.)
    OrderPlexStateClosingOut
        
    ' This state indicates that all orders have been completed.
    OrderPlexStateClosed

    ' This state indicates that a cancellation notification for either
    ' the stop order or the target order, when both exist, has been received.
    ' We don't know whether the order was cancelled because the user/broker
    ' cancelled it, or because the other order was filled. So in this
    ' state we wait for the other order to be either cancelled or
    ' filled or for a timeout to elapse.
    OrderPlexStateAwaitingOtherOrderCancel
    
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type OrderPlexProfile
    Key                 As String
    action              As OrderActions
    Quantity            As Long
    EntryPrice          As Double
    ExitPrice           As Double
    Profit              As Currency
    MaxProfit           As Currency
    MaxLoss             As Currency
    Risk                As Currency
    StartTime           As Date
    EndTime             As Date
    Timezone            As String
    Description         As String
    EntryReason         As String
    StopReason          As String
    TargetReason        As String
    ClosedOut           As Boolean
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mPositionManager As PositionManager
Private mOrderContext As OrderContext
Private mTicker As Ticker

Private mTickSize As Double
Private mTickValue As Currency
Private mcurrencyCode As String

Private mStateEngine As StateEngine
Private mKey As String
Private WithEvents mTimerEntry As TimerListItem
Attribute mTimerEntry.VB_VarHelpID = -1

Private mDescription As String
Private mEntryReason As String
Private mTargetReason As String
Private mStopReason As String

Private mSize As Long
Private mCancelPrice As Double
Private mCancelTime As Date

Private WithEvents mEntryOrder As Order
Attribute mEntryOrder.VB_VarHelpID = -1
Private mEntryOrderCopy As Order

Private WithEvents mStopOrder As Order
Attribute mStopOrder.VB_VarHelpID = -1
Private mStopOrderCopy As Order
Attribute mStopOrderCopy.VB_VarHelpID = -1

Private WithEvents mTargetOrder As Order
Attribute mTargetOrder.VB_VarHelpID = -1
Private mTargetOrderCopy As Order
Attribute mTargetOrderCopy.VB_VarHelpID = -1

Private WithEvents mCloseoutOrder As Order
Attribute mCloseoutOrder.VB_VarHelpID = -1

' the following references are only to ensure that
' something has a reference to these objects, otherwise they
' may disappear prematurely!
Private mStopStrategy As StopManagementStrategy
Private mScalingStrategy As PositionManagementStrategy
Private mExitStrategy As PositionManagementStrategy

Private mNewQuantity As Long
Private mNewEntryPrice As Double
Private mNewEntryTriggerPrice As Double
Private mNewStopPrice As Double
Private mNewStopOffset As Long
Private mNewStopTriggerPrice As Double
Private mNewStopQuantity As Long
Private mNewTargetPrice As Double
Private mNewTargetOffset As Long
Private mNewTargetTriggerPrice As Double
Private mNewTargetQuantity As Long

Private mIsSimulated As Boolean

Private mValueFactor As Double
Private mProfit As Currency
Private mMaxProfit As Currency
Private mMaxLoss As Currency
Private mDrawdown As Currency
Private mProfitNotificationThreshold As Currency
Private mLossNotificationThreshold As Currency
Private mDrawdownNotificationThreshold As Currency

Private mChangeListeners As Collection
Private mStateChangeListeners As Collection

Private mCreationTime As Date

Private mProfitListeners As Collection

Private mIndexApplication As Long

Private mDirty As Boolean

Private mOrders As Collection

Private mOrderSubmitter As IOrderSubmitter

Private mPreventUnprotectedPosition As Boolean

Private WithEvents mStateTimeoutTimer As IntervalTimer
Attribute mStateTimeoutTimer.VB_VarHelpID = -1

Private mSummary As OrderPlexSummary

Private mStartTime As Date
Private mEndTime As Date
Private mTimezone As String

Private mCumBuyPrice As Double
Private mCumSellPrice As Double

Private mEntryPrice As Double
Private mExitPrice As Double

Private mLastBid As Double
Private mLastAsk As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mNewQuantity = -1
mNewEntryPrice = -1
mNewEntryTriggerPrice = -1
mNewStopPrice = -1
mNewStopOffset = DummyOffset
mNewStopTriggerPrice = -1
mNewStopQuantity = -1
mNewTargetPrice = -1
mNewTargetOffset = DummyOffset
mNewTargetTriggerPrice = -1
mNewTargetQuantity = -1

Set mSummary = New OrderPlexSummary

mIndexApplication = GOrderPlex.gNextApplicationIndex

Set mOrders = New Collection

Set mChangeListeners = New Collection
Set mStateChangeListeners = New Collection

Set mProfitListeners = New Collection

Set mStateEngine = CreateStateEngine(GOrderPlex.TableBuilder, _
                                    OrderPlexStateCodes.OrderPlexStateCreated)
                                    
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderPlex terminated"
End Sub

'@================================================================================
' IOrderPlex Interface Members
'@================================================================================

Private Property Get IOrderPlex_closeoutOrder() As IOrder
Set IOrderPlex_closeoutOrder = mCloseoutOrder
End Property

Private Property Get IOrderPlex_EntryOrder() As IOrder
Set IOrderPlex_EntryOrder = mEntryOrder
End Property

Private Property Get IOrderPlex_Key() As String
IOrderPlex_Key = mKey
End Property

Private Property Get IOrderPlex_numberOfOrders() As Long
Const ProcName As String = "IOrderPlex_numberOfOrders"

On Error GoTo Err

IOrderPlex_numberOfOrders = mOrders.Count

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IOrderPlex_Order(ByVal index As Long) As IOrder
Const ProcName As String = "IOrderPlex_Order"

On Error GoTo Err

Set IOrderPlex_Order = Order(index)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get IOrderPlex_State() As OrderPlexStateCodes
Const ProcName As String = "IOrderPlex_State"
IOrderPlex_State = State
End Property

Private Property Get IOrderPlex_Stoporder() As IOrder
Const ProcName As String = "IOrderPlex_Stoporder"
Set IOrderPlex_Stoporder = mStopOrder
End Property

Private Property Get IOrderPlex_targetOrder() As IOrder
Const ProcName As String = "IOrderPlex_targetOrder"
Set IOrderPlex_targetOrder = mTargetOrder
End Property

'@================================================================================
' QuoteListener Interface Members
'@================================================================================

Private Sub QuoteListener_Ask(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_Ask"

On Error GoTo Err

If ev.price = mLastAsk Then Exit Sub

If Not LongPosition Then calcProfitLoss ev.price

mLastAsk = ev.price

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub QuoteListener_Bid(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_Bid"

On Error GoTo Err

If ev.price = mLastBid Then Exit Sub

If LongPosition Then calcProfitLoss ev.price

mLastBid = ev.price

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub QuoteListener_High(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_High"


End Sub

Private Sub QuoteListener_Low(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_Low"


End Sub

Private Sub QuoteListener_OpenInterest(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_OpenInterest"


End Sub

Private Sub QuoteListener_PreviousClose(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_PreviousClose"


End Sub

Private Sub QuoteListener_sessionOpen(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_sessionOpen"


End Sub

Private Sub QuoteListener_Trade(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_Trade"
Dim cancelPriceHit As Boolean


On Error GoTo Err

If ev.price > mLastAsk Then QuoteListener_Ask ev    ' because the Ask price must have gone at least that high
If ev.price < mLastBid Then QuoteListener_Bid ev    ' because the Bid price must have gone at least that low

If mCancelPrice = 0 Then Exit Sub

If LongPosition Then
    Select Case mEntryOrder.OrderType
    Case OrderTypes.OrderTypeLimit, _
            OrderTypes.OrderTypeLimitIfTouched, _
            OrderTypes.OrderTypeMarketIfTouched
        If ev.price >= mCancelPrice Then cancelPriceHit = True
    Case OrderTypes.OrderTypeStop, _
        OrderTypes.OrderTypeStopLimit
        If ev.price <= mCancelPrice Then cancelPriceHit = True
    End Select
Else
    Select Case mEntryOrder.OrderType
    Case OrderTypes.OrderTypeLimit, _
            OrderTypes.OrderTypeLimitIfTouched, _
            OrderTypes.OrderTypeMarketIfTouched
        If ev.price <= mCancelPrice Then cancelPriceHit = True
    Case OrderTypes.OrderTypeStop, _
        OrderTypes.OrderTypeStopLimit
        If ev.price >= mCancelPrice Then cancelPriceHit = True
    End Select
End If
If cancelPriceHit Then
    logOrderMessage "Cancel price hit"
    Cancel
    fireChange OrderPlexSelfCancelled
    RaiseEvent SelfCancelled
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub QuoteListener_Volume(ev As QuoteEventData)
Const ProcName As String = "QuoteListener_Volume"


End Sub

'@================================================================================
' mCloseoutOrder Event Handlers
'@================================================================================

Private Sub mCloseoutOrder_Fill(ByVal pExec As Execution)
Const ProcName As String = "mCloseoutOrder_Fill"

On Error GoTo Err

fireChange OrderPlexCloseoutOrderChanged

NotifyExecution pExec

logOrderMessage "Closeout order fill: (id=" & _
                    mCloseoutOrder.Id & "): " & _
                    GOrderActionToString(mCloseoutOrder.action) & " " & pExec.Quantity & " at " & _
                    pExec.price

RaiseEvent CloseoutOrderPartiallyFilled

SaveRecoveryInfo
Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mCloseoutOrder_PropertyChanged()
Const ProcName As String = "mCloseoutOrder_PropertyChanged"

On Error GoTo Err

fireChange OrderPlexCloseoutOrderChanged

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mCloseoutOrder_StatusChanged()
Const ProcName As String = "mCloseoutOrder_StatusChanged"

On Error GoTo Err

fireChange OrderPlexCloseoutOrderChanged
Select Case mCloseoutOrder.Status
    Case OrderStatuses.OrderStatusFilled
        logOrderMessage "Closeout order completed (id=" & _
                            mCloseoutOrder.Id & "): " & _
                            GOrderActionToString(mCloseoutOrder.action) & " " & mCloseoutOrder.QuantityFilled & " at " & _
                            mCloseoutOrder.AveragePrice
                            
        checkAllOrdersComplete
        
        If size <> 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                        ProjectName & "." & ModuleName & ":" & ProcName, _
                        "Closeout incomplete: Size remaining is " & size
        
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
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
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
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mEntryOrder_Fill(ByVal pExec As Execution)
Const ProcName As String = "mEntryOrder_Fill"
Dim Risk As Currency


On Error GoTo Err

fireChange OrderPlexEntryOrderChanged

NotifyExecution pExec

CancelAfter = 0

mCancelPrice = 0

logOrderMessage "Entry order fill (id=" & _
                    mEntryOrder.Id & "): " & _
                    GOrderActionToString(mEntryOrder.action) & " " & pExec.Quantity & " at " & _
                    pExec.price

If Not mStopOrder Is Nothing Then
    If LongPosition Then
        If mStopOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
            Risk = (mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * mStopOrder.Offset * mTickValue
        Else
            Risk = (mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * (pExec.price - mStopOrder.TriggerPrice) * mValueFactor
        End If
    Else
        If mStopOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
            Risk = -(mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * mStopOrder.Offset * mTickValue
        Else
            Risk = (mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * (mStopOrder.TriggerPrice - pExec.price) * mValueFactor
        End If
    End If
End If

placePeggedOrders
        
If Not mStopOrder Is Nothing Then
    If mStopOrder.Quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        AdjustStop mStopOrder.TriggerPrice, Abs(size)
    End If
End If

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        AdjustTarget mTargetOrder.LimitPrice, Abs(size)
    End If
End If

RaiseEvent EntryOrderPartiallyFilled

handleStimulus OpStimuli.StimEntryOrderFill

SaveRecoveryInfo
Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mEntryOrder_PropertyChanged()
Const ProcName As String = "mEntryOrder_PropertyChanged"

On Error GoTo Err

fireChange OrderPlexEntryOrderChanged

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mEntryOrder_StatusChanged()
Const ProcName As String = "mEntryOrder_StatusChanged"

On Error GoTo Err

fireChange OrderPlexEntryOrderChanged
Select Case mEntryOrder.Status
    Case OrderStatuses.OrderStatusFilled
        mCancelPrice = 0
        
        logOrderMessage "Entry order completed (id=" & _
                            mEntryOrder.Id & "): " & _
                            GOrderActionToString(mEntryOrder.action) & " " & mEntryOrder.QuantityFilled & " at " & _
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
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mStateTimeoutTimer Event Handlers
'@================================================================================

Private Sub mStateTimeoutTimer_TimerExpired()
Const ProcName As String = "mStateTimeoutTimer_TimerExpired"

On Error GoTo Err

handleStimulus OpStimuli.StimTimeoutExpired

SaveRecoveryInfo
Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mStopOrder Event Handlers
'@================================================================================

Private Sub mStopOrder_Dirty()
Const ProcName As String = "mStopOrder_Dirty"

On Error GoTo Err

setDirty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub mStopOrder_Fill(ByVal pExec As Execution)
Const ProcName As String = "mStopOrder_Fill"

On Error GoTo Err

fireChange OrderPlexStopOrderChanged

NotifyExecution pExec

logOrderMessage "Stop order fill (id=" & _
                        mStopOrder.Id & "): " & _
                        GOrderActionToString(mStopOrder.action) & " " & pExec.Quantity & " at " & _
                        pExec.price

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        AdjustTarget mTargetOrder.LimitPrice, Abs(size)
    End If
End If

RaiseEvent StopOrderPartiallyFilled

SaveRecoveryInfo
Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mStopOrder_PropertyChanged()
Const ProcName As String = "mStopOrder_PropertyChanged"

On Error GoTo Err

fireChange OrderPlexStopOrderChanged

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mStopOrder_StatusChanged()
Const ProcName As String = "mStopOrder_StatusChanged"

On Error GoTo Err

fireChange OrderPlexStopOrderChanged
Select Case mStopOrder.Status
    Case OrderStatuses.OrderStatusFilled
        logOrderMessage "Stop order completed (id=" & _
                            mStopOrder.Id & "): " & _
                            GOrderActionToString(mStopOrder.action) & " " & mStopOrder.QuantityFilled & " at " & _
                            mStopOrder.AveragePrice
        
        RaiseEvent StopOrderFilled
        
        checkAllOrdersComplete
    
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusSubmitted
    Case OrderStatuses.OrderStatusCancelled
        logOrderMessage "Stop order cancelled (id=" & mStopOrder.Id & ")"
        handleStimulus OpStimuli.StimStopOrderCancelled
        checkAllOrdersComplete
    Case OrderStatuses.OrderStatusRejected
        logOrderMessage "Stop order rejected (id=" & mStopOrder.Id & ")"
        handleStimulus OpStimuli.StimStopOrderCancelled
        checkAllOrdersComplete
End Select

SaveRecoveryInfo
Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
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
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mTargetOrder_Fill(ByVal pExec As Execution)
Const ProcName As String = "mTargetOrder_Fill"

On Error GoTo Err

fireChange OrderPlexTargetOrderChanged

NotifyExecution pExec

logOrderMessage "Target order fill (id=" & _
                        mTargetOrder.Id & "): " & _
                        GOrderActionToString(mTargetOrder.action) & " " & pExec.Quantity & " at " & _
                        pExec.price

If Not mStopOrder Is Nothing Then
    If mStopOrder.Quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        AdjustStop mStopOrder.TriggerPrice, Abs(size)
    End If
End If

RaiseEvent TargetOrderPartiallyFilled

SaveRecoveryInfo
Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mTargetOrder_PropertyChanged()
Const ProcName As String = "mTargetOrder_PropertyChanged"

On Error GoTo Err

fireChange OrderPlexTargetOrderChanged

Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mTargetOrder_StatusChanged()
Const ProcName As String = "mTargetOrder_StatusChanged"

On Error GoTo Err

fireChange OrderPlexTargetOrderChanged
Select Case mTargetOrder.Status
    Case OrderStatuses.OrderStatusFilled
        logOrderMessage "Target order completed (id=" & _
                            mTargetOrder.Id & "): " & _
                            GOrderActionToString(mTargetOrder.action) & " " & mTargetOrder.QuantityFilled & " at " & _
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
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mTimerEntry Event Handlers
'@================================================================================

Private Sub mTimerEntry_Expired()
Const ProcName As String = "mTimerEntry_Expired"

On Error GoTo Err

logOrderMessage "Cancel time expired"
Cancel
fireChange OrderPlexSelfCancelled
RaiseEvent SelfCancelled
CancelAfter = 0

SaveRecoveryInfo
Exit Sub

Err:
GHandleFatalError pProcName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let CancelAfter(ByVal value As Long)
Const ProcName As String = "CancelAfter"

On Error GoTo Err

If value <> 0 Then
    mCancelTime = Now + CDbl(value) / 86400#
    Set mTimerEntry = mTicker.AddTimer(value, expiryTimeUnits.ExpiryTimeUnitSeconds, Nothing)

    mSummary.CancelTime = mCancelTime
ElseIf Not mTimerEntry Is Nothing Then
    mCancelTime = 0
    mTicker.RemoveTimer mTimerEntry
    Set mTimerEntry = Nothing

    mSummary.CancelTime = mCancelTime
End If

SaveRecoveryInfo
Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let CancelTime(ByVal value As Date)
Const ProcName As String = "CancelTime"
mCancelTime = value
mSummary.CancelTime = value
End Property

Friend Property Let CloseoutOrder(ByVal value As Order)
Const ProcName As String = "CloseoutOrder"
Set mCloseoutOrder = value
mSummary.CloseoutOrderId = mCloseoutOrder.Id
mOrders.Add mCloseoutOrder
End Property

Public Property Get CloseoutOrder() As Order
Const ProcName As String = "CloseoutOrder"
Set CloseoutOrder = mCloseoutOrder
End Property

Public Property Get Contract() As Contract
Const ProcName As String = "Contract"

On Error GoTo Err

Set Contract = mTicker.Contract

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let CreationTime(ByVal value As Date)
Const ProcName As String = "CreationTime"
mCreationTime = value
mSummary.CreationTime = value
End Property

Public Property Get CreationTime() As Date
Const ProcName As String = "CreationTime"
CreationTime = mCreationTime
End Property

Friend Property Let CumBuyPrice(ByVal value As Double)
mCumBuyPrice = value
mSummary.CumBuyPrice = value
End Property

Friend Property Let CumSellPrice(ByVal value As Double)
mCumSellPrice = value
mSummary.CumSellPrice = value
End Property

Public Property Let Description(ByVal value As String)
Const ProcName As String = "Description"
If mDescription = value Then Exit Property
mDescription = value
mSummary.Description = value
SaveRecoveryInfo
End Property

Public Property Get Description() As String
Const ProcName As String = "Description"
Description = mDescription
End Property

Public Property Get Dirty() As Boolean
Const ProcName As String = "Dirty"
Dirty = mDirty
End Property

Public Property Let DrawdownNotificationThreshold(ByVal value As Currency)
Const ProcName As String = "DrawdownNotificationThreshold"
If mDrawdownNotificationThreshold = value Then Exit Property
mDrawdownNotificationThreshold = value
mSummary.DrawdownNotificationThreshold = value
SaveRecoveryInfo
End Property

Friend Property Let EntryOrder(ByVal value As Order)
Const ProcName As String = "EntryOrder"
Set mEntryOrder = value
mSummary.EntryOrderId = mEntryOrder.Id

If mOrders.Count >= 1 Then
    mOrders.Add mEntryOrder, , 1
Else
    mOrders.Add mEntryOrder
End If

End Property

Public Property Get EntryOrder() As Order
Const ProcName As String = "EntryOrder"
Set EntryOrder = mEntryOrder
End Property

Public Property Get EntryPrice() As Double
Const ProcName As String = "EntryPrice"
EntryPrice = mEntryPrice
End Property

Public Property Get ExitPrice() As Double
Const ProcName As String = "ExitPrice"
ExitPrice = mExitPrice
End Property

Public Property Let EntryReason(ByVal value As String)
Const ProcName As String = "EntryReason"

On Error GoTo Err

If mEntryReason = value Then Exit Property
mEntryReason = value
mSummary.EntryReason = value
SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get EntryReason() As String
Const ProcName As String = "EntryReason"
EntryReason = mEntryReason
End Property

Friend Property Let ExitStrategy( _
                ByVal value As PositionManagementStrategy)
Const ProcName As String = "ExitStrategy"
Set mExitStrategy = value
End Property

Public Property Get IndexApplication() As Long
Const ProcName As String = "IndexApplication"
IndexApplication = mIndexApplication
End Property

Public Property Get IsComplete() As Boolean
Const ProcName As String = "IsComplete"

On Error GoTo Err

IsComplete = (mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IsRiskUnlimited() As Boolean
Const ProcName As String = "IsRiskUnlimited"

On Error GoTo Err

IsRiskUnlimited = (Risk = MaxCurrency)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let IsSimulated(ByVal value As Boolean)
Const ProcName As String = "IsSimulated"
mIsSimulated = value
End Property

Public Property Get IsSimulated() As Boolean
Const ProcName As String = "IsSimulated"
IsSimulated = mIsSimulated
End Property

Friend Property Let Key(ByVal value As String)
Const ProcName As String = "Key"
mKey = value
mSummary.Key = mKey
End Property

Public Property Get Key() As String
Const ProcName As String = "Key"
Key = mKey
End Property

Public Property Get LongPosition() As Boolean
Const ProcName As String = "LongPosition"

On Error GoTo Err

LongPosition = (mEntryOrder.action = ActionBuy)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let LossNotificationThreshold(ByVal value As Currency)
Const ProcName As String = "LossNotificationThreshold"
If mLossNotificationThreshold = value Then Exit Property
mLossNotificationThreshold = value
mSummary.LossNotificationThreshold = value
SaveRecoveryInfo
End Property

Public Property Let NewEntryPrice(ByVal value As Double)
Const ProcName As String = "NewEntryPrice"

On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                ProjectName & "." & ModuleName & ":" & ProcName, _
                                                "Can't change entry price - order " & EntryOrder.Id & " already filled"

Select Case mEntryOrder.OrderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeTrail, _
        OrderTypes.OrderTypeAutoStop, _
        OrderTypes.OrderTypeAutoLimit, _
        OrderTypes.OrderTypeStop
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Can't change entry price for order type " & GOrderTypeToString(mEntryOrder.OrderType)
End Select

If value <> mEntryOrder.LimitPrice Then
    logOrderMessage "Set entry price=" & mTicker.FormatPrice(value)
    mNewEntryPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewEntryTriggerPrice(ByVal value As Double)
Const ProcName As String = "NewEntryTriggerPrice"

On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                ProjectName & "." & ModuleName & ":" & ProcName, _
                                                "Can't change entry trigger price - order " & EntryOrder.Id & " already filled"

Select Case mEntryOrder.OrderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeLimit, _
        OrderTypes.OrderTypeTrail, _
        OrderTypes.OrderTypeAutoLimit
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Can't change trigger price for order type " & GOrderTypeToString(mEntryOrder.OrderType)
End Select

If value <> mEntryOrder.TriggerPrice Then
    logOrderMessage "Set entry trigger price=" & mTicker.FormatPrice(value)
    mNewEntryTriggerPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewOrderPrice( _
                ByVal orderIndex As Long, _
                ByVal value As Double)
Const ProcName As String = "NewOrderPrice"
Dim lOrder As Order


On Error GoTo Err

Set lOrder = mOrders(orderIndex)

If Not mEntryOrder Is Nothing Then
    If lOrder Is mEntryOrder Then
        NewEntryPrice = value
        Exit Property
    End If
End If

If Not mStopOrder Is Nothing Then
    If lOrder Is mStopOrder Then
        NewStopPrice = value
        Exit Property
    End If
End If

If Not mTargetOrder Is Nothing Then
    If lOrder Is mTargetOrder Then
        NewTargetPrice = value
        Exit Property
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewOrderTriggerPrice( _
                ByVal orderIndex As Long, _
                ByVal value As Double)
Const ProcName As String = "NewOrderTriggerPrice"
Dim lOrder As Order


On Error GoTo Err

Set lOrder = mOrders(orderIndex)

If Not mEntryOrder Is Nothing Then
    If lOrder Is mEntryOrder Then
        NewEntryTriggerPrice = value
        Exit Property
    End If
End If

If Not mStopOrder Is Nothing Then
    If lOrder Is mStopOrder Then
        NewStopTriggerPrice = value
        Exit Property
    End If
End If

If Not mTargetOrder Is Nothing Then
    If lOrder Is mTargetOrder Then
        NewTargetTriggerPrice = value
        Exit Property
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewQuantity(ByVal value As Long)
Const ProcName As String = "NewQuantity"

On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                ProjectName & "." & ModuleName & ":" & ProcName, _
                                                "Can't change quantity - order " & EntryOrder.Id & " already filled"

If value <> mEntryOrder.Quantity Then
    logOrderMessage "Set quantity=" & value
    mNewQuantity = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewStopOffset(ByVal value As Long)
Const ProcName As String = "NewStopOffset"

On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewStopOffset", _
                                        "No stop order exists"

If Not mStopOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Can't change offset for order type " & GOrderTypeToString(mEntryOrder.OrderType)
End If

If mStopOrder.Status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                        "Can't change stop offset - order " & mStopOrder.Id & " already placed"

If value <> mStopOrder.Offset Then
    logOrderMessage "Set stop offset=" & value
    mNewStopOffset = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewStopPrice(ByVal value As Double)
Const ProcName As String = "NewStopPrice"

On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                        "No stop order exists"

Select Case mStopOrder.OrderType
Case OrderTypes.OrderTypeStopLimit
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Can't change price for order type " & GOrderTypeToString(mStopOrder.OrderType)
End Select

If value <> mStopOrder.LimitPrice Then
    logOrderMessage "Set stop price=" & mTicker.FormatPrice(value)
    mNewStopPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewStopQuantity(ByVal value As Long)
Const ProcName As String = "NewStopQuantity"

On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                        "No stop order exists"

If value <> mStopOrder.Quantity Then
    logOrderMessage "Set stop quantity=" & value
    mNewStopQuantity = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewStopTriggerPrice(ByVal value As Double)
Const ProcName As String = "NewStopTriggerPrice"

On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                        "No stop order exists"

Select Case mStopOrder.OrderType
    Case OrderTypes.OrderTypeStop, _
        OrderTypes.OrderTypeStopLimit
    Case Else
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Can't change trigger price for order type " & GOrderTypeToString(mStopOrder.OrderType)
End Select

If value <> mStopOrder.TriggerPrice Then
    logOrderMessage "Set stop trigger price=" & mTicker.FormatPrice(value)
    mNewStopTriggerPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewTargetPrice(ByVal value As Double)
Const ProcName As String = "NewTargetPrice"

On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                        "No target order exists"

Select Case mTargetOrder.OrderType
Case OrderTypes.OrderTypeLimitIfTouched, _
    OrderTypes.OrderTypeLimit
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Can't change price for order type " & GOrderTypeToString(mTargetOrder.OrderType)
End Select

If value <> mTargetOrder.LimitPrice Then
    logOrderMessage "Set target price=" & mTicker.FormatPrice(value)
    mNewTargetPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewTargetOffset(ByVal value As Long)
Const ProcName As String = "NewTargetOffset"

On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewTargetOffset", _
                                        "No target order exists"

If Not mTargetOrder.OrderType = OrderTypes.OrderTypeAutoLimit Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Can't change offset for order type " & GOrderTypeToString(mTargetOrder.OrderType)
End If

If mTargetOrder.Status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                        "Can't change target offset - order " & mTargetOrder.Id & " already placed"

If value <> mTargetOrder.Offset Then
    logOrderMessage "Set target offset=" & value
    mNewTargetOffset = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewTargetQuantity(ByVal value As Long)
Const ProcName As String = "NewTargetQuantity"

On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                        "No target order exists"

If value <> mTargetOrder.Quantity Then
    logOrderMessage "Set target quantity=" & value
    mNewTargetQuantity = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let NewTargetTriggerPrice(ByVal value As Double)
Const ProcName As String = "NewTargetTriggerPrice"

On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                        "No target order exists"

Select Case mTargetOrder.OrderType
Case OrderTypes.OrderTypeLimitIfTouched, _
    OrderTypes.OrderTypeMarketIfTouched
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Can't change trigger price for order type " & GOrderTypeToString(mTargetOrder.OrderType)
End Select

If value <> mTargetOrder.TriggerPrice Then
    logOrderMessage "Set target trigger price=" & mTicker.FormatPrice(value)
    mNewTargetTriggerPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PendingSize() As Long
Const ProcName As String = "PendingSize"

On Error GoTo Err

If mEntryOrder Is Nothing Then PendingSize = 0: Exit Property
If mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed Then PendingSize = 0: Exit Property
PendingSize = IIf(LongPosition, mEntryOrder.QuantityRemaining, -mEntryOrder.QuantityRemaining)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Order(ByVal index As Long) As Order
Const ProcName As String = "Order"

On Error GoTo Err

Set Order = mOrders(index)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get OrderContext() As OrderContext
Const ProcName As String = "OrderContext"

On Error GoTo Err

Set OrderContext = mOrderContext

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let OrderContext(ByVal value As OrderContext)
Const ProcName As String = "OrderContext"
Set mOrderContext = value
End Property

Friend Property Let orderSubmitter( _
                ByVal value As IOrderSubmitter)
Const ProcName As String = "orderSubmitter"
Set mOrderSubmitter = value
End Property

Public Property Get NumberOfOrders() As Long
Const ProcName As String = "NumberOfOrders"

On Error GoTo Err

NumberOfOrders = mOrders.Count

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PositionManager() As PositionManager
Const ProcName As String = "PositionManager"
Set PositionManager = mPositionManager
End Property

Friend Property Let PositionManager(ByVal value As PositionManager)
Const ProcName As String = "PositionManager"
Set mPositionManager = value
mSummary.TickerKey = mPositionManager.Key
mSummary.WorkspaceName = mPositionManager.WorkSpace.name
End Property

Public Property Let PreventUnprotectedPosition( _
                ByVal value As Boolean)
Const ProcName As String = "PreventUnprotectedPosition"

On Error GoTo Err

If mPreventUnprotectedPosition = value Then Exit Property
If value And _
    State <> OrderPlexStateCreated And _
    mStopOrder Is Nothing _
Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "No stop loss order exists"
End If

mPreventUnprotectedPosition = value
mSummary.PreventUnprotectedPosition = value
SaveRecoveryInfo

If mPreventUnprotectedPosition Then
    mStateEngine.SetConditions OpConditions.CondProtected
Else
    mStateEngine.ClearConditions OpConditions.CondProtected
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PreventUnprotectedPosition() As Boolean
Const ProcName As String = "PreventUnprotectedPosition"

On Error GoTo Err

PreventUnprotectedPosition = mPreventUnprotectedPosition

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Let ProfitNotificationThreshold(ByVal value As Currency)
Const ProcName As String = "ProfitNotificationThreshold"
If mProfitNotificationThreshold = value Then Exit Property
mProfitNotificationThreshold = value
mSummary.ProfitNotificationThreshold = value
SaveRecoveryInfo
End Property

Public Property Get Risk() As Currency
Const ProcName As String = "Risk"
Dim stopTotalPrice As Double
Dim pendingTotalPrice As Double
Dim filledTotalPrice As Double


On Error GoTo Err

If mStopOrder Is Nothing Then
    Risk = MaxCurrency
    Exit Property
End If

If mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed Then Exit Property

If mStopOrder.Status = OrderStatuses.OrderStatusFilled Then Exit Property

' the amount at Risk has two parts - that due to fills already obtained, and
' that due to fills pending.

filledTotalPrice = mEntryOrder.QuantityFilled * mEntryOrder.AveragePrice
Select Case EntryOrder.OrderType
    Case OrderTypes.OrderTypeMarket
        pendingTotalPrice = mEntryOrder.QuantityRemaining * _
                IIf(LongPosition, mTicker.AskPrice, mTicker.BidPrice)
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
stopTotalPrice = mStopOrder.TriggerPrice * _
                        (mEntryOrder.Quantity + mEntryOrder.QuantityFilled)
If LongPosition Then
    Risk = mTickSize * mTickValue * (filledTotalPrice + pendingTotalPrice - stopTotalPrice) / mTickSize
Else
    Risk = mTickSize * mTickValue * (stopTotalPrice - filledTotalPrice - pendingTotalPrice) / mTickSize
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Friend Property Let ScalingStrategy( _
                ByVal value As PositionManagementStrategy)
Const ProcName As String = "ScalingStrategy"
Set mScalingStrategy = value
End Property

Friend Property Let size(ByVal value As Long)
Const ProcName As String = "size"
setSize value
End Property

Public Property Get size() As Long
Const ProcName As String = "size"
size = mSize
End Property

Public Property Get StartTime() As Date
Const ProcName As String = "StartTime"
StartTime = mStartTime
End Property

Public Property Get EndTime() As Date
Const ProcName As String = "EndTime"
EndTime = mEndTime
End Property

Friend Property Let State(ByVal value As OrderPlexStateCodes)
Const ProcName As String = "State"
Set mStateEngine = CreateStateEngine(GOrderPlex.TableBuilder, _
                                    value)
mSummary.State = value
If IsComplete Then doCompletionActions
End Property

Public Property Get State() As OrderPlexStateCodes
Const ProcName As String = "State"

On Error GoTo Err

State = mStateEngine.State

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let Stoporder(ByVal value As Order)
Const ProcName As String = "Stoporder"
Set mStopOrder = value
mSummary.StopOrderId = mStopOrder.Id

mStateEngine.SetConditions OpConditions.CondStopOrderExists
mStateEngine.ClearConditions OpConditions.CondStopOrderCancelled

If Not mEntryOrder Is Nothing Then
    mOrders.Add mStopOrder, , , 1
ElseIf Not mTargetOrder Is Nothing Then
    mOrders.Add mStopOrder, , 1
Else
    mOrders.Add mStopOrder
End If

End Property

Public Property Get Stoporder() As Order
Const ProcName As String = "Stoporder"
Set Stoporder = mStopOrder
End Property

Public Property Let StopReason(ByVal value As String)
Const ProcName As String = "StopReason"

On Error GoTo Err

If mStopReason = value Then Exit Property
mStopReason = value
mSummary.StopReason = value
SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get StopReason() As String
Const ProcName As String = "StopReason"
StopReason = mStopReason
End Property

Friend Property Let StopStrategy( _
                ByVal value As StopManagementStrategy)
Const ProcName As String = "StopStrategy"
Set mStopStrategy = value
End Property

Friend Property Let TargetOrder(ByVal value As Order)
Const ProcName As String = "TargetOrder"
Set mTargetOrder = value
mSummary.TargetOrderId = mTargetOrder.Id

mStateEngine.SetConditions OpConditions.CondTargetOrderExists
mStateEngine.ClearConditions OpConditions.CondTargetOrderCancelled

If mEntryOrder Is Nothing And _
    mStopOrder Is Nothing _
Then
    mOrders.Add mTargetOrder
ElseIf mEntryOrder Is Nothing Then
    mOrders.Add mTargetOrder, , , 1
ElseIf mStopOrder Is Nothing Then
    mOrders.Add mTargetOrder, , , 1
Else
    mOrders.Add mTargetOrder, , , 2
End If

End Property

Public Property Get TargetOrder() As Order
Const ProcName As String = "TargetOrder"
Set TargetOrder = mTargetOrder
End Property

Public Property Let TargetReason(ByVal value As String)
Const ProcName As String = "TargetReason"

On Error GoTo Err

If mTargetReason = value Then Exit Property
mTargetReason = value
mSummary.TargetReason = value
SaveRecoveryInfo

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TargetReason() As String
Const ProcName As String = "TargetReason"
TargetReason = mTargetReason
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Const ProcName As String = "Ticker"
Dim lContract As Contract

On Error GoTo Err

Set mTicker = value
Set lContract = mTicker.Contract
mTickSize = lContract.TickSize
mTickValue = lContract.TickValue
mValueFactor = lContract.Multiplier
mcurrencyCode = lContract.specifier.currencyCode

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Ticker() As Ticker
Const ProcName As String = "Ticker"
Set Ticker = mTicker
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Activate()
Const ProcName As String = "Activate"
mTicker.AddQuoteListener Me
fireChange OrderPlexCreated
If Not mEntryOrder Is Nothing Then fireChange OrderPlexEntryOrderChanged
If Not mStopOrder Is Nothing Then fireChange OrderPlexStopOrderChanged
If Not mTargetOrder Is Nothing Then fireChange OrderPlexTargetOrderChanged
If Not mCloseoutOrder Is Nothing Then fireChange OrderPlexCloseoutOrderChanged
End Sub

Public Sub AddChangeListener(ByVal value As ChangeListener)
Const ProcName As String = "AddChangeListener"

On Error GoTo Err

mChangeListeners.Add value

' notify this new listenener that we are already created
Dim ev As ChangeEventData
Set ev.source = Me
ev.changeType = OrderPlexChangeTypes.OrderPlexCreated
value.Change ev

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub AddStateChangeListener(ByVal value As StateChangeListener)
Const ProcName As String = "AddStateChangeListener"

On Error GoTo Err

mStateChangeListeners.Add value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub AddProfitListener(ByVal value As ProfitListener)
Const ProcName As String = "AddProfitListener"

On Error GoTo Err

mProfitListeners.Add value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub AdjustStop(ByVal stopTriggerPrice As Double, _
                    Optional ByVal Quantity As Long = -1)
Const ProcName As String = "AdjustStop"

On Error GoTo Err

NewStopTriggerPrice = stopTriggerPrice
If Quantity <> -1 Then NewStopQuantity = Quantity
Update

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Public Sub AdjustTarget(ByVal targetPrice As Double, _
                    Optional ByVal Quantity As Long = -1)
Const ProcName As String = "AdjustTarget"

On Error GoTo Err

NewTargetPrice = targetPrice
If Quantity <> -1 Then NewTargetQuantity = Quantity
Update

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Public Function Cancel(Optional ByVal evenIfFilled As Boolean = False) As Boolean
Const ProcName As String = "Cancel"



On Error GoTo Err

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

''
' Cancels any changes made to this <code>orderPlex</code> that have not yet
' been actioned.
'@/
Public Sub CancelChanges()
Const ProcName As String = "CancelChanges"

On Error GoTo Err

mNewQuantity = -1
mNewEntryPrice = -1
mNewEntryTriggerPrice = -1
mNewStopPrice = -1
mNewStopOffset = DummyOffset
mNewStopTriggerPrice = -1
mNewTargetPrice = -1
mNewTargetOffset = DummyOffset
mNewTargetTriggerPrice = -1

''
' note that we can't just use the copies because external
' code may have references to the member orders which must
' remain valid
'@/
mEntryOrderCopy.syncToMe mEntryOrder
mEntryOrder.setClean
mStopOrderCopy.syncToMe mStopOrder
mStopOrder.setClean
mTargetOrderCopy.syncToMe mTargetOrder
mTargetOrder.setClean

mDirty = False
fireChange OrderPlexChangeTypes.OrderPlexChangesCancelled
RaiseEvent ChangesCancelled
RaiseEvent Clean

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub CheckComplete()

Const ProcName As String = "CheckComplete"
On Error GoTo Err

checkAllOrdersComplete
SaveRecoveryInfo

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

''
' Copies the current state of the member orders for use by cancelChanges
'@/
Friend Sub Checkpoint()
Const ProcName As String = "Checkpoint"

On Error GoTo Err

Set mEntryOrderCopy = mEntryOrder.Clone
mEntryOrder.setClean

If Not mStopOrder Is Nothing Then
    Set mStopOrderCopy = mStopOrder.Clone
    mStopOrder.setClean
End If

If Not mTargetOrder Is Nothing Then
    Set mTargetOrderCopy = mTargetOrder.Clone
    mTargetOrder.setClean
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub CloseOut(ByVal method As String)
Const ProcName As String = "CloseOut"



On Error GoTo Err

If UCase(method) <> "MKT" Then
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "method " & method & " not implemented"
End If

logOrderMessage "Closing out: " & _
                    "size=" & size & _
                    "; pending size=" & PendingSize

handleStimulus OpStimuli.StimCloseout

SaveRecoveryInfo
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Friend Function ContainsOrder( _
                ByVal pOrder As Order)
Const ProcName As String = "ContainsOrder"
Dim lOrder As Order

For Each lOrder In mOrders
    If lOrder.Id = pOrder.Id Then
        ContainsOrder = True
        Exit For
    End If
Next
End Function

Friend Function CreateEntryOrder() As Order
Const ProcName As String = "CreateEntryOrder"
Dim lOrder As Order


On Error GoTo Err

If Not mEntryOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Entry order already exists"

Set lOrder = mOrderContext.NewOrder
EntryOrder = lOrder

Set CreateEntryOrder = mEntryOrder

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Function CreateStopOrder() As Order
Const ProcName As String = "CreateStopOrder"

Dim lOrder As Order


On Error GoTo Err

If Not mStopOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Stop order already exists"

Set lOrder = mOrderContext.NewOrder
Stoporder = lOrder
Set CreateStopOrder = mStopOrder

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Function CreateTargetOrder() As Order
Const ProcName As String = "CreateTargetOrder"

Dim lOrder As Order


On Error GoTo Err

If Not mTargetOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Target order already exists"

Set lOrder = mOrderContext.NewOrder
TargetOrder = lOrder
Set CreateTargetOrder = mTargetOrder
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Execute(Optional ByVal cancelPrice As Double = 0)
Const ProcName As String = "Execute"

On Error GoTo Err

If mPreventUnprotectedPosition And mStopOrder Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "No stop loss order exists"
End If

mCancelPrice = cancelPrice
mSummary.cancelPrice = mCancelPrice

mCreationTime = mTicker.Timestamp
mSummary.CreationTime = mCreationTime

fireChange OrderPlexChangeTypes.OrderPlexCreated

SaveRecoveryInfo    ' ensure the OrderPlex is recorded before
                    ' doing the state change, which will cause
                    ' the orders to be recorded
                    
handleStimulus OpStimuli.StimExecute

SaveRecoveryInfo
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"

On Error GoTo Err

Set mTicker = Nothing
Set mPositionManager = Nothing
Set mOrderContext = Nothing
Set mOrderSubmitter = Nothing
If Not mEntryOrder Is Nothing Then mEntryOrder.Finish
If Not mStopOrder Is Nothing Then mStopOrder.Finish
If Not mTargetOrder Is Nothing Then mTargetOrder.Finish
If Not mCloseoutOrder Is Nothing Then mCloseoutOrder.Finish
Set mStopStrategy = Nothing
Set mScalingStrategy = Nothing
Set mExitStrategy = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub NotifyCloseout( _
                ByVal pPrice As Double, _
                ByVal pTimestamp As Date)
Const ProcName As String = "NotifyCloseout"

On Error GoTo Err

If size <> 0 Then processExecution -size, pPrice, pTimestamp, mTimezone, True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub NotifyEndOfTradingSession( _
                ByVal pTimestamp As Date)
Const ProcName As String = "NotifyEndOfTradingSession"

On Error GoTo Err

mEndTime = pTimestamp
logPlexProfile False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub NotifyExecution( _
                ByVal pExec As IExecutionReport, _
                Optional ByVal pCloseout As Boolean)
Const ProcName As String = "NotifyExecution"

On Error GoTo Err

processExecution IIf(pExec.action = ActionBuy, pExec.Quantity, -1 * pExec.Quantity), pExec.price, pExec.FillTime, pExec.Timezone, pCloseout

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveChangeListener(ByVal value As ChangeListener)
Const ProcName As String = "RemoveChangeListener"
Dim i As Long

On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveStateChangeListener(ByVal value As StateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
Dim i As Long

On Error GoTo Err

For i = mStateChangeListeners.Count To 1 Step -1
    If mStateChangeListeners.Item(i) Is value Then mStateChangeListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveProfitListener(ByVal value As ProfitListener)
Const ProcName As String = "RemoveProfitListener"
Dim i As Long

On Error GoTo Err

For i = mProfitListeners.Count To 1 Step -1
    If mProfitListeners.Item(i) Is value Then mProfitListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub ResyncCloseoutOrder()
Const ProcName As String = "ResyncCloseoutOrder"
On Error GoTo Err

If mOrderSubmitter Is Nothing Then Exit Sub
If mCloseoutOrder Is Nothing Then Exit Sub
If Not mCloseoutOrder.IsActive Then Exit Sub
mOrderSubmitter.ModifyOrderPlex Me, False, False, False, True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub ResyncEntryOrder()
Const ProcName As String = "ResyncEntryOrder"
On Error GoTo Err

If mOrderSubmitter Is Nothing Then Exit Sub
If mEntryOrder Is Nothing Then Exit Sub
If Not mEntryOrder.IsActive Then Exit Sub

mOrderSubmitter.ModifyOrderPlex Me, True, False, False, False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub ResyncStopOrder()
Const ProcName As String = "ResyncStopOrder"
On Error GoTo Err

If mOrderSubmitter Is Nothing Then Exit Sub
If mStopOrder Is Nothing Then Exit Sub
If Not mStopOrder.IsActive Then Exit Sub

mOrderSubmitter.ModifyOrderPlex Me, False, True, False, False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub ResyncTargetOrder()
Const ProcName As String = "ResyncTargetOrder"
On Error GoTo Err

If mOrderSubmitter Is Nothing Then Exit Sub
If mTargetOrder Is Nothing Then Exit Sub
If Not mTargetOrder.IsActive Then Exit Sub

mOrderSubmitter.ModifyOrderPlex Me, False, False, True, False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub SaveRecoveryInfo()
Const ProcName As String = "SaveRecoveryInfo"
On Error GoTo Err

If mIsSimulated Then Exit Sub
If Not mSummary.IsDirty Then Exit Sub

gTB.OrderPlexRecoveryController.SaveRecoveryInfo mPositionManager.Key, mSummary
mSummary.setClean

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub SyncToSummary(ByVal ops As OrderPlexSummary)
Const ProcName As String = "UpdateOrderPlex"

On Error GoTo Err

Set mSummary = ops

mCancelPrice = mSummary.cancelPrice
mCancelTime = mSummary.CancelTime
If mSummary.CloseoutOrderId <> "" And mCloseoutOrder Is Nothing Then CloseoutOrder = createRawOrder(mSummary.CloseoutOrderId)
mCreationTime = mSummary.CreationTime
mCumBuyPrice = mSummary.CumBuyPrice
mCumSellPrice = mSummary.CumSellPrice
mDescription = mSummary.Description
mDrawdownNotificationThreshold = mSummary.DrawdownNotificationThreshold
If mSummary.EntryOrderId <> "" And mEntryOrder Is Nothing Then EntryOrder = createRawOrder(mSummary.EntryOrderId)
mEntryReason = mSummary.EntryReason
mLossNotificationThreshold = mSummary.LossNotificationThreshold
mPreventUnprotectedPosition = mSummary.PreventUnprotectedPosition
mProfitNotificationThreshold = mSummary.ProfitNotificationThreshold
mSize = mSummary.size
State = mSummary.State
If mSummary.StopOrderId <> "" And mStopOrder Is Nothing Then Stoporder = createRawOrder(mSummary.StopOrderId)
mStopReason = mSummary.StopReason
If mSummary.TargetOrderId <> "" And mTargetOrder Is Nothing Then TargetOrder = createRawOrder(mSummary.TargetOrderId)
mTargetReason = mSummary.TargetReason

mSummary.setClean

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Public Sub Update()
Const ProcName As String = "Update"
Dim updateEntry As Boolean
Dim updateStop As Boolean
Dim updateTarget As Boolean


On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCreated
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Sub
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Action not allowed in state " & mStateEngine.State
End Select

checkForPendingUpdates updateEntry, updateStop, updateTarget

If (Not updateEntry) And (Not updateStop) And (Not updateTarget) Then
    ' no updates to Action
    Exit Sub
End If

logOrderMessage "Update order plex"

If State <> OrderPlexStateCreated Then
    mOrderSubmitter.ModifyOrderPlex Me, updateEntry, updateStop, updateTarget, False
End If

Checkpoint

If updateEntry Then
    fireChange OrderPlexEntryOrderChanged
End If
If updateStop Then
    fireChange OrderPlexStopOrderChanged
    RaiseEvent StopAdjusted
End If
If updateTarget Then
    fireChange OrderPlexTargetOrderChanged
    RaiseEvent TargetAdjusted
End If
fireChange OrderPlexChangeTypes.OrderPlexChangesApplied
RaiseEvent Clean

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcProfitLoss(ByVal value As Double)
Const ProcName As String = "calcProfitLoss"
Dim Profit As Currency
Dim notifyProfit As Boolean
Dim notifyMaxProfit As Boolean
Dim notifyMaxLoss As Boolean
Dim notifyDrawdown As Boolean


On Error GoTo Err

If value = 0 Then Exit Function

Profit = (mCumSellPrice + size * value - mCumBuyPrice) * mValueFactor

If Profit <> mProfit Then
    mProfit = Profit
    notifyProfit = True
End If

If Profit > mMaxProfit Then
    mMaxProfit = Profit
    notifyMaxProfit = True
End If

If Profit < mMaxLoss Then
    mMaxLoss = Profit
    notifyMaxLoss = True
End If

If Profit < mMaxProfit Then
    mDrawdown = mMaxProfit - Profit
    notifyDrawdown = True
ElseIf mDrawdown <> 0 Then
    mDrawdown = 0
    notifyDrawdown = True
End If

If notifyProfit Then fireProfit mProfit, ProfitTypes.ProfitTypeProfit
If notifyMaxProfit Then fireProfit mMaxProfit, ProfitTypes.ProfitTypeMaxProfit
If notifyMaxLoss Then fireProfit mMaxLoss, ProfitTypes.ProfitTypeMaxLoss
If notifyDrawdown Then fireProfit mDrawdown, ProfitTypes.ProfitTypeDrawdown

If notifyProfit Then
    If mProfitNotificationThreshold <> 0 And _
        mProfit >= mProfitNotificationThreshold _
    Then
        fireChange OrderPlexProfitThresholdExceeded
        RaiseEvent ProfitThresholdExceeded
        mProfitNotificationThreshold = 0
        mSummary.ProfitNotificationThreshold = 0
    End If
    
    If mLossNotificationThreshold <> 0 And _
        mProfit <= mLossNotificationThreshold _
    Then
        fireChange OrderPlexLossThresholdExceeded
        RaiseEvent LossThresholdExceeded
        mLossNotificationThreshold = 0
        mSummary.LossNotificationThreshold = 0
End If

    If mDrawdownNotificationThreshold <> 0 And _
        (mMaxProfit - Profit >= mDrawdownNotificationThreshold) _
    Then
        fireChange OrderPlexDrawdownThresholdExceeded
        RaiseEvent DrawdownThresholdExceeded
        mDrawdownNotificationThreshold = 0
        mSummary.DrawdownNotificationThreshold = 0
    End If
End If

SaveRecoveryInfo
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function cancelOrders() As Boolean
Const ProcName As String = "cancelOrders"
Dim StopOrderId As String
Dim TargetOrderId As String


On Error GoTo Err

CancelAfter = 0
mCancelPrice = 0

Select Case mEntryOrder.Status
Case OrderStatuses.OrderStatusCreated
    mEntryOrder.UpdateStatus OrderStatuses.OrderStatusCancelled, mPositionManager.Key

    cancelOrders = True
End Select

If Not mStopOrder Is Nothing Then
    Select Case mStopOrder.Status
    Case OrderStatuses.OrderStatusCreated
        mStopOrder.UpdateStatus OrderStatuses.OrderStatusCancelled, mPositionManager.Key
        cancelOrders = True
    End Select
End If

If Not mTargetOrder Is Nothing Then
    Select Case mTargetOrder.Status
    Case OrderStatuses.OrderStatusCreated
        mTargetOrder.UpdateStatus OrderStatuses.OrderStatusCancelled, mPositionManager.Key
        cancelOrders = True
    End Select
End If

If mOrderSubmitter.CancelOrderPlex(Me) Then
    cancelOrders = True
    
    If Not mStopOrder Is Nothing Then StopOrderId = mStopOrder.Id
    If Not mTargetOrder Is Nothing Then TargetOrderId = mTargetOrder.Id
    logOrderMessage "Cancel orders: " & _
                    EntryOrder.Id & _
                    IIf(StopOrderId <> "", ", " & StopOrderId, ", ") & _
                    IIf(TargetOrderId <> "", ", " & TargetOrderId, ", ")
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Private Function cancelStateTimeout()
Const ProcName As String = "cancelStateTimeout"

On Error GoTo Err

mStateTimeoutTimer.StopTimer

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function cancelStopOrder()
Const ProcName As String = "cancelStopOrder"

On Error GoTo Err

If mStopOrder Is Nothing Then Exit Function
Select Case mStopOrder.Status
Case OrderStatuses.OrderStatusCreated
    mStopOrder.UpdateStatus OrderStatuses.OrderStatusCancelled, mPositionManager.Key
Case OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    logOrderMessage "Cancel order: id=" & mStopOrder.Id
    mOrderSubmitter.cancelStopOrder Me
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function cancelTargetOrder()
Const ProcName As String = "cancelTargetOrder"

On Error GoTo Err

If mTargetOrder Is Nothing Then Exit Function
Select Case mTargetOrder.Status
Case OrderStatuses.OrderStatusCreated
    mTargetOrder.UpdateStatus OrderStatuses.OrderStatusCancelled, mPositionManager.Key
Case OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    logOrderMessage "Cancel order: id=" & mTargetOrder.Id
    mOrderSubmitter.cancelTargetOrder Me
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub checkAllOrdersComplete()
Const ProcName As String = "checkAllOrdersComplete"

On Error GoTo Err

If IsComplete Then Exit Sub

If Not mEntryOrder Is Nothing Then
    If mEntryOrder.IsActive Then Exit Sub
End If

If Not mStopOrder Is Nothing Then
    If mStopOrder.IsActive Then Exit Sub
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

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
        If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                                        "Can't change Size - order " & EntryOrder.Id & " already filled"
        mEntryOrder.Quantity = mNewQuantity
        updateEntry = True
    End If
    mNewQuantity = -1
    
    If mNewEntryPrice >= 0 And _
        mNewEntryPrice <> mEntryOrder.LimitPrice _
    Then
        If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                                        "Can't change price - order " & mEntryOrder.Id & " already filled"
        mEntryOrder.LimitPrice = mNewEntryPrice
        updateEntry = True
    End If
    mNewEntryPrice = -1
    
    If mNewEntryTriggerPrice >= 0 And _
        mNewEntryTriggerPrice <> mEntryOrder.TriggerPrice _
    Then
        If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                                        "Can't change trigger price - order " & mEntryOrder.Id & " already filled"
        mEntryOrder.TriggerPrice = mNewEntryTriggerPrice
        updateEntry = True
    End If
    mNewEntryTriggerPrice = -1
End If

If Not mStopOrder Is Nothing Then
    If mStopOrder.IsDirty Then updateStop = True
    If mNewStopPrice >= 0 And _
        mNewStopPrice <> mStopOrder.LimitPrice _
    Then
        If mStopOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                                        "Can't change trigger price - order " & mStopOrder.Id & " already filled"
        If mStopOrder.Status = OrderStatuses.OrderStatusSubmitted Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                                        "Can't change price - order " & mStopOrder.Id & " already submitted"
        mStopOrder.LimitPrice = mNewStopPrice
        updateStop = True
    End If
    mNewStopPrice = -1
    
    If mNewStopOffset <> DummyOffset And _
        mNewStopOffset <> mStopOrder.Offset _
    Then
        If mStopOrder.Status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            ProjectName & "." & ModuleName & ":" & ProcName, _
                                                            "Can't change stop offset - order " & mStopOrder.Id & " already placed"
        mStopOrder.Offset = mNewStopOffset
    End If
    mNewStopOffset = DummyOffset

    If mNewStopQuantity >= 0 And _
        mNewStopQuantity <> mStopOrder.Quantity _
    Then
        If mStopOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            ProjectName & "." & ModuleName & ":" & ProcName, _
                                                            "Can't change stop quantity - order " & mStopOrder.Id & " already placed"
        mStopOrder.Quantity = mNewStopQuantity
    End If
    mNewStopQuantity = -1

    If mNewStopTriggerPrice >= 0 And _
        mNewStopTriggerPrice <> mStopOrder.TriggerPrice _
    Then
        If mStopOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                                        "Can't change trigger price - order " & mEntryOrder.Id & " already filled"
        mStopOrder.TriggerPrice = mNewStopTriggerPrice
        updateStop = True
    End If
    mNewStopTriggerPrice = -1
End If

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.IsDirty Then updateTarget = True
    If mNewTargetPrice >= 0 And _
        mNewTargetPrice <> mTargetOrder.LimitPrice _
    Then
        If mTargetOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                                        "Can't change price - order " & mTargetOrder.Id & " already filled"
        mTargetOrder.LimitPrice = mNewTargetPrice
        updateTarget = True
    End If
    mNewTargetPrice = -1
    
    If mNewTargetOffset <> DummyOffset And _
        mNewTargetOffset <> mTargetOrder.Offset _
    Then
        If mTargetOrder.Status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            ProjectName & "." & ModuleName & ":" & ProcName, _
                                                            "Can't change target offset - order " & TargetOrder.Id & " already placed"
        mTargetOrder.Offset = mNewTargetOffset
    End If
    mNewTargetOffset = DummyOffset

    If mNewTargetQuantity >= 0 And _
        mNewTargetQuantity <> mTargetOrder.Quantity _
    Then
        If mTargetOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            ProjectName & "." & ModuleName & ":" & ProcName, _
                                                            "Can't change target quantity - order " & mTargetOrder.Id & " already placed"
        mTargetOrder.Quantity = mNewTargetQuantity
    End If
    mNewTargetQuantity = -1

    If mNewTargetTriggerPrice >= 0 And _
        mNewTargetTriggerPrice <> mTargetOrder.TriggerPrice _
    Then
        If mTargetOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        ProjectName & "." & ModuleName & ":" & ProcName, _
                                                        "Can't change trigger price - order " & mEntryOrder.Id & " already filled"
        mTargetOrder.TriggerPrice = mNewTargetTriggerPrice
        updateTarget = True
    End If
    mNewTargetTriggerPrice = -1
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function createRawOrder(ByVal Id As String) As Order
Dim lOrder As Order
Const ProcName As String = "createRawOrder"

On Error GoTo Err

Set lOrder = New Order
lOrder.Id = Id
Set createRawOrder = lOrder

mPositionManager.Orders.Add lOrder

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function doActions(ByRef pActions() As OpActions) As Variant
Const ProcName As String = "doActions"
Dim action As Variant


On Error GoTo Err

For Each action In pActions
    Select Case CLng(action)
    Case SpecialActions.NoAction
    Case OpActions.ActPlaceOrders
        doActions = placeOrders
    Case OpActions.ActCancelOrders
        doActions = cancelOrders
    Case OpActions.ActCancelStopOrder
        doActions = cancelStopOrder
    Case OpActions.ActCancelTargetOrder
        doActions = cancelTargetOrder
    Case OpActions.ActResubmitStopOrder
        doActions = resubmitStopOrder
    Case OpActions.ActResubmitTargetOrder
        doActions = resubmitTargetOrder
    Case OpActions.ActResubmitStopAndTargetOrders
        doActions = resubmitStopAndTargetOrders
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
    Case Else
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Invalid Action " & action
    End Select
Next

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function doCompletionActions()
Const ProcName As String = "doCompletionActions"

On Error GoTo Err

mOrderContext.NotifyOrderPlexClosed Me
fireChange OrderPlexCompleted
RaiseEvent Completed
Set mStopStrategy = Nothing
Set mScalingStrategy = Nothing
Set mExitStrategy = Nothing
If size = 0 Then If Not mTicker Is Nothing Then mTicker.RemoveQuoteListener Me

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub fireChange( _
                ByVal changeType As OrderPlexChangeTypes)
Const ProcName As String = "fireChange"
Dim listener As ChangeListener
Dim ev As ChangeEventData

On Error GoTo Err

Set ev.source = Me
ev.changeType = changeType
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent Change(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireStateChange( _
                ByVal State As OrderPlexStateCodes)
Const ProcName As String = "fireStateChange"
Dim listener As StateChangeListener
Dim ev As StateChangeEventData

On Error GoTo Err

Set ev.source = Me
ev.State = State
For Each listener In mStateChangeListeners
    listener.Change ev
Next
RaiseEvent StateChange(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub fireProfit( _
                ByVal ProfitAmount As Double, _
                ByVal profitType As ProfitTypes)
Const ProcName As String = "fireProfit"
Dim listener As ProfitListener
Dim Profit As ProfitEventData

On Error GoTo Err

Set Profit.source = Me
Profit.ProfitAmount = ProfitAmount
Profit.profitType = profitType
Profit.currencyCode = mcurrencyCode
For Each listener In mProfitListeners
    listener.ProfitAmount Profit
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function handleStimulus(ByVal stimulus As OpStimuli) As Variant
Const ProcName As String = "handleStimulus"
Dim currState As OrderPlexStateCodes
Dim lActions() As OpActions

On Error GoTo Err

currState = mStateEngine.State
lActions = mStateEngine.NotifyStimulus(stimulus)
mSummary.State = State

handleStimulus = doActions(lActions)
If mStateEngine.State <> currState Then
    fireChange OrderPlexChangeTypes.OrderPlexStateChanged
    fireStateChange mStateEngine.State
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub logOrderMessage( _
                ByRef msg As String)
Const ProcName As String = "logOrderMessage"
If mTicker Is Nothing Then Exit Sub
GLogOrder msg & " (Curr price=" & mTicker.FormatPrice(mTicker.TradePrice) & "; id=" & Key & ")", _
        mIsSimulated, _
        Me
End Sub

Private Sub logPlexProfile(ByVal pClosedOut As Boolean)
Const ProcName As String = "logPlexProfile"
Dim profile As OrderPlexProfile

On Error GoTo Err

With profile
    .action = mEntryOrder.action
    .ClosedOut = pClosedOut
    .Description = mDescription
    .EndTime = mEndTime
    .EntryReason = mEntryReason
    If LongPosition Then
        .EntryPrice = mCumBuyPrice / mEntryOrder.Quantity
        .ExitPrice = mCumSellPrice / mEntryOrder.Quantity
    Else
        .EntryPrice = mCumSellPrice / mEntryOrder.Quantity
        .ExitPrice = mCumBuyPrice / mEntryOrder.Quantity
    End If
    .Key = mKey
    .MaxLoss = mMaxLoss
    .MaxProfit = mMaxProfit
    .Profit = mProfit
    .Quantity = mEntryOrder.Quantity
    .Risk = Risk
    .StartTime = mStartTime
    .StopReason = mStopReason
    .TargetReason = mTargetReason
    .Timezone = mTimezone
End With
GLogOrderPlexProfileStruct profile, mIsSimulated, Me

GLogOrderPlexProfileString FormatDateTime(mStartTime) & "," & _
            FormatDateTime(mEndTime) & "," & _
            mDescription & "," & _
            mKey & "," & _
            GOrderActionToString(mEntryOrder.action) & "," & _
            mEntryOrder.Quantity & "," & _
            mEntryPrice & "," & _
            mExitPrice & "," & _
            mEntryReason & "," & _
            mTargetReason & "," & _
            mStopReason & "," & _
            IIf(pClosedOut, "1,", "0,") & _
            mProfit & "," & _
            mMaxProfit & "," & _
            mMaxLoss & "," & _
            Risk, _
            mIsSimulated, _
            Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function placeCloseoutOrder()
Const ProcName As String = "placeCloseoutOrder"

Dim lOrder As Order


On Error GoTo Err

Set lOrder = mOrderContext.NewOrder
CloseoutOrder = lOrder

mCloseoutOrder.action = IIf(LongPosition, OrderActions.ActionSell, OrderActions.ActionBuy)
mCloseoutOrder.Quantity = Abs(size)
mCloseoutOrder.OrderType = OrderTypes.OrderTypeMarket
mCloseoutOrder.TimeInForce = OrderTifs.TIFGoodTillCancelled

fireChange OrderPlexCloseoutOrderCreated
mOrderSubmitter.ModifyOrderPlex Me, False, False, False, True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Private Function placeOrders()
Const ProcName As String = "placeOrders"



On Error GoTo Err

mOrderSubmitter.ExecuteOrderPlex Me
mTicker.AddQuoteListener Me

Checkpoint

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub placePeggedOrders()
Const ProcName As String = "placePeggedOrders"
Dim updated As Boolean

On Error GoTo Err

If Not mStopOrder Is Nothing Then
    If mStopOrder.Status = OrderStatuses.OrderStatusCreated Then
        If mStopOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
            ' the stop price is to be pegged to the entry price - the
            ' order's offset property will already contain any required
            ' offset from that price
            mStopOrder.OrderType = OrderTypes.OrderTypeStop
            NewStopTriggerPrice = (mStopOrder.Offset * mTickSize) + EntryOrder.LastFillPrice
            NewStopQuantity = mEntryOrder.QuantityFilled
            updated = True
        End If
    End If
End If
If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Status = OrderStatuses.OrderStatusCreated Then
        If mTargetOrder.OrderType = OrderTypes.OrderTypeAutoLimit Then
            ' the limit price is to be pegged to the entry price - the
            ' order's offset property will already contain any required
            ' offset from that price
            mTargetOrder.OrderType = OrderTypes.OrderTypeLimit
            NewTargetPrice = (mTargetOrder.Offset * mTickSize) + EntryOrder.LastFillPrice
            NewTargetQuantity = mEntryOrder.QuantityFilled
            updated = False
        End If
    End If
End If

If updated Then Update

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub processExecution( _
                ByVal pSignedQuantity As Long, _
                ByVal pPrice As Double, _
                ByVal pTimestamp As Date, _
                ByVal pTimezone As String, _
                ByVal pCloseout As Boolean)
Const ProcName As String = "processExecution"

On Error GoTo Err

If mSize = 0 Then
    mStartTime = pTimestamp
    mTimezone = pTimezone
End If

setSize size + pSignedQuantity

If pSignedQuantity > 0 Then
    CumBuyPrice = mCumBuyPrice + pSignedQuantity * pPrice
Else
    CumSellPrice = mCumSellPrice - pSignedQuantity * pPrice
End If
    
If size > 0 Then
    calcProfitLoss mLastBid
Else
    calcProfitLoss mLastAsk
End If

If size = 0 Then
    mEndTime = pTimestamp
    logPlexProfile pCloseout
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Function resubmitStopOrder()
Const ProcName As String = "resubmitStopOrder"

On Error GoTo Err

mOrderSubmitter.resubmitStopOrder Me
logOrderMessage "Resubmit stop order: id=" & mStopOrder.Id & _
            "; new broker id=" & mStopOrder.BrokerId
mStateEngine.ClearConditions OpConditions.CondStopOrderCancelled

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function resubmitStopAndTargetOrders()
Const ProcName As String = "resubmitStopAndTargetOrders"

On Error GoTo Err

mOrderSubmitter.resubmitStopAndTargetOrders Me
logOrderMessage "Resubmit stop order: id=" & mStopOrder.Id & _
            "; new broker id=" & mStopOrder.BrokerId
logOrderMessage "Resubmit target order: id=" & mTargetOrder.Id & _
            "; new broker id=" & mTargetOrder.BrokerId
mStateEngine.ClearConditions OpConditions.CondStopOrderCancelled
mStateEngine.ClearConditions OpConditions.CondTargetOrderCancelled

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Private Function resubmitTargetOrder()
Const ProcName As String = "resubmitTargetOrder"

On Error GoTo Err

mOrderSubmitter.resubmitTargetOrder Me
logOrderMessage "Resubmit target order: id=" & mTargetOrder.Id & _
            "; new broker id=" & mTargetOrder.BrokerId
mStateEngine.ClearConditions OpConditions.CondTargetOrderCancelled

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

''
' Set the 'dirty' flag, and raise a Dirty event if one has not already
' been raised.
'@/
Private Sub setDirty()
Const ProcName As String = "setDirty"

On Error GoTo Err

If Not mDirty Then
    mDirty = True
    RaiseEvent Dirty
Else
    mDirty = True
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setSize(ByVal value As Long)
Const ProcName As String = "setSize"

On Error GoTo Err

mSize = value
mSummary.size = value

If mSize <> 0 Then
    mStateEngine.SetConditions OpConditions.CondSizeNonZero
Else
    mStateEngine.ClearConditions OpConditions.CondSizeNonZero
End If
fireChange OrderPlexChangeTypes.OrderPlexSizeChanged

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function setStateTimeout()
Const ProcName As String = "setStateTimeout"

On Error GoTo Err

Set mStateTimeoutTimer = CreateIntervalTimer(1000)
mStateTimeoutTimer.StartTimer

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

