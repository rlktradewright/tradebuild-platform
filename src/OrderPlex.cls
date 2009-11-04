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
Implements TradeBuildSP.IOrderPlex

'@================================================================================
' Events
'@================================================================================

Event Change(ByRef ev As ChangeEvent)
Event StateChange(ByRef ef As StateChangeEvent)
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
Private mHandle As Long
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
Private mEntryOrderFinished As Boolean

Private WithEvents mStopOrder As Order
Attribute mStopOrder.VB_VarHelpID = -1
Private mStopOrderCopy As Order
Attribute mStopOrderCopy.VB_VarHelpID = -1
Private mStopOrderFinished As Boolean

Private WithEvents mTargetOrder As Order
Attribute mTargetOrder.VB_VarHelpID = -1
Private mTargetOrderCopy As Order
Attribute mTargetOrderCopy.VB_VarHelpID = -1
Private mTargetOrderFinished As Boolean

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

Private mOrderSubmitter As TradeBuildSP.IOrderSubmitter

Private mPreventUnprotectedPosition As Boolean

Private WithEvents mStateTimeoutTimer As IntervalTimer
Attribute mStateTimeoutTimer.VB_VarHelpID = -1

Private mSummary As OrderPlexSummary

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mHandle = -1
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

Private Property Get IOrderPlex_closeoutOrder() As TradeBuildSP.IOrder
Set IOrderPlex_closeoutOrder = mCloseoutOrder
End Property

Private Property Get IOrderPlex_EntryOrder() As TradeBuildSP.IOrder
Set IOrderPlex_EntryOrder = mEntryOrder
End Property

Private Property Get IOrderPlex_Key() As String
IOrderPlex_Key = mKey
End Property

Private Property Get IOrderPlex_numberOfOrders() As Long
Dim failpoint As Long
On Error GoTo Err

IOrderPlex_numberOfOrders = mOrders.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IOrderPlex_numberOfOrders", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Private Property Get IOrderPlex_Order(ByVal index As Long) As TradeBuildSP.IOrder
Dim failpoint As Long
On Error GoTo Err

Set IOrderPlex_Order = Order(index)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IOrderPlex_Order", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Private Property Get IOrderPlex_State() As TradeBuildSP.OrderPlexStateCodes
IOrderPlex_State = State
End Property

Private Property Get IOrderPlex_Stoporder() As TradeBuildSP.IOrder
Set IOrderPlex_Stoporder = mStopOrder
End Property

Private Property Get IOrderPlex_targetOrder() As TradeBuildSP.IOrder
Set IOrderPlex_targetOrder = mTargetOrder
End Property

'@================================================================================
' QuoteListener Interface Members
'@================================================================================

Private Sub QuoteListener_Ask(ev As QuoteEvent)
Dim failpoint As Long
On Error GoTo Err

If Not LongPosition Then calcProfitLoss ev.price

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_Ask", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub QuoteListener_Bid(ev As QuoteEvent)
Dim failpoint As Long
On Error GoTo Err

If LongPosition Then calcProfitLoss ev.price

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_Bid", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub QuoteListener_High(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_OpenInterest(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_PreviousClose(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_sessionOpen(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_Trade(ev As QuoteEvent)
Dim cancelPriceHit As Boolean

Dim failpoint As Long
On Error GoTo Err

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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_Trade", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub QuoteListener_Volume(ev As QuoteEvent)

End Sub

'@================================================================================
' mCloseoutOrder Event Handlers
'@================================================================================

Private Sub mCloseoutOrder_Fill(ByVal QuantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

Dim failpoint As Long
On Error GoTo Err

calcProfitLoss fillPrice

fireChange OrderPlexCloseoutOrderChanged
thisSize = IIf(mCloseoutOrder.Action = ActionBuy, QuantityFilled, -QuantityFilled)
setSize size + thisSize
logOrderMessage "Closeout order fill: (id=" & _
                    mCloseoutOrder.Id & "): " & _
                    GOrderActionToString(mCloseoutOrder.Action) & " " & QuantityFilled & " at " & _
                    fillPrice

mPositionManager.NotifyExecution thisSize, fillPrice, mHandle, True

RaiseEvent CloseoutOrderPartiallyFilled

Exit Sub

Err:
GHandleFatalError pProcName:="mCloseoutOrder_Fill", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mCloseoutOrder_PropertyChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexCloseoutOrderChanged

Exit Sub

Err:
GHandleFatalError pProcName:="mCloseoutOrder_PropertyChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mCloseoutOrder_StatusChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexCloseoutOrderChanged
Select Case mCloseoutOrder.Status
    Case OrderStatuses.OrderStatusFilled
        logOrderMessage "Closeout order completed (id=" & _
                            mCloseoutOrder.Id & "): " & _
                            GOrderActionToString(mCloseoutOrder.Action) & " " & mCloseoutOrder.QuantityFilled & " at " & _
                            mCloseoutOrder.AveragePrice
                            
        checkAllOrdersComplete
        
        If size <> 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, "OrderPlex::mCloseoutOrder_StatusChanged", "Closeout incomplete: Size remaining is " & size
        
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

Exit Sub

Err:
GHandleFatalError pProcName:="mCloseoutOrder_StatusChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mEntryOrder Event Handlers
'@================================================================================

Private Sub mEntryOrder_Dirty()
Dim failpoint As Long
On Error GoTo Err

setDirty

Exit Sub

Err:
GHandleFatalError pProcName:="mEntryOrder_Dirty", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mEntryOrder_Fill(ByVal QuantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long
Dim Risk As Currency

Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexEntryOrderChanged

CancelAfter = 0

mCancelPrice = 0

thisSize = IIf(LongPosition, QuantityFilled, -QuantityFilled)
setSize size + thisSize


logOrderMessage "Entry order fill (id=" & _
                    mEntryOrder.Id & "): " & _
                    GOrderActionToString(mEntryOrder.Action) & " " & QuantityFilled & " at " & _
                    fillPrice

If mHandle < 0 And Not mStopOrder Is Nothing Then
    If LongPosition Then
        If mStopOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
            Risk = (mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * mStopOrder.Offset * mTickValue
        Else
            Risk = (mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * (fillPrice - mStopOrder.TriggerPrice) * mValueFactor
        End If
    Else
        If mStopOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
            Risk = -(mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * mStopOrder.Offset * mTickValue
        Else
            Risk = (mEntryOrder.Quantity + mEntryOrder.QuantityFilled) * (mStopOrder.TriggerPrice - fillPrice) * mValueFactor
        End If
    End If
End If

If mHandle < 0 Then
    mHandle = mPositionManager.RegisterPlex(mKey, mDescription, mEntryOrder.Action, Risk)
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonEntry, mEntryReason
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonTarget, mTargetReason
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonStop, mStopReason
End If
mPositionManager.NotifyExecution thisSize, fillPrice, mHandle, False

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

handleStimulus StateTransitionStimuli.StimEntryOrderFill

Exit Sub

Err:
GHandleFatalError pProcName:="mEntryOrder_Fill", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mEntryOrder_PropertyChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexEntryOrderChanged

Exit Sub

Err:
GHandleFatalError pProcName:="mEntryOrder_PropertyChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mEntryOrder_StatusChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexEntryOrderChanged
Select Case mEntryOrder.Status
    Case OrderStatuses.OrderStatusFilled
        mCancelPrice = 0
        
        logOrderMessage "Entry order completed (id=" & _
                            mEntryOrder.Id & "): " & _
                            GOrderActionToString(mEntryOrder.Action) & " " & mEntryOrder.QuantityFilled & " at " & _
                            mEntryOrder.AveragePrice
        
        RaiseEvent EntryOrderFilled
        
        mEntryOrderFinished = True
        checkAllOrdersComplete
    
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusSubmitted
        
        If mEntryOrder.QuantityFilled > 0 Then
            placePeggedOrders
        End If
    Case OrderStatuses.OrderStatusCancelled
        mEntryOrderFinished = True
        logOrderMessage "Entry order cancelled (id=" & mEntryOrder.Id & ")"
        handleStimulus StateTransitionStimuli.StimEntryOrderCancelled
        checkAllOrdersComplete
    Case OrderStatuses.OrderStatusRejected
        mEntryOrderFinished = True
        logOrderMessage "Entry order rejected (id=" & mEntryOrder.Id & ")"
        handleStimulus StateTransitionStimuli.StimEntryOrderCancelled
        checkAllOrdersComplete
End Select

Exit Sub

Err:
GHandleFatalError pProcName:="mEntryOrder_StatusChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mStateTimeoutTimer Event Handlers
'@================================================================================

Private Sub mStateTimeoutTimer_TimerExpired()
Dim failpoint As Long
On Error GoTo Err

handleStimulus StateTransitionStimuli.StimTimeoutExpired

Exit Sub

Err:
GHandleFatalError pProcName:="mStateTimeoutTimer_TimerExpired", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mStopOrder Event Handlers
'@================================================================================

Private Sub mStopOrder_Dirty()
Dim failpoint As Long
On Error GoTo Err

setDirty

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="mStopOrder_Dirty", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mStopOrder_Fill(ByVal QuantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

Dim failpoint As Long
On Error GoTo Err

calcProfitLoss fillPrice

fireChange OrderPlexStopOrderChanged

thisSize = IIf(mStopOrder.Action = ActionBuy, QuantityFilled, -QuantityFilled)
setSize size + thisSize
logOrderMessage "Stop order fill (id=" & _
                        mStopOrder.Id & "): " & _
                        GOrderActionToString(mStopOrder.Action) & " " & QuantityFilled & " at " & _
                        fillPrice

mPositionManager.NotifyExecution thisSize, fillPrice, mHandle, False

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        AdjustTarget mTargetOrder.LimitPrice, Abs(size)
    End If
End If

RaiseEvent StopOrderPartiallyFilled

Exit Sub

Err:
GHandleFatalError pProcName:="mStopOrder_Fill", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mStopOrder_PropertyChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexStopOrderChanged

Exit Sub

Err:
GHandleFatalError pProcName:="mStopOrder_PropertyChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mStopOrder_StatusChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexStopOrderChanged
Select Case mStopOrder.Status
    Case OrderStatuses.OrderStatusFilled
        logOrderMessage "Stop order completed (id=" & _
                            mStopOrder.Id & "): " & _
                            GOrderActionToString(mStopOrder.Action) & " " & mStopOrder.QuantityFilled & " at " & _
                            mStopOrder.AveragePrice
        
        RaiseEvent StopOrderFilled
        
        mStopOrderFinished = True
        checkAllOrdersComplete
    
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusSubmitted
    Case OrderStatuses.OrderStatusCancelled
        mStopOrderFinished = True
        logOrderMessage "Stop order cancelled (id=" & mStopOrder.Id & ")"
        mStateEngine.SetConditions Conditions.CondStopOrderCancelled
        handleStimulus StateTransitionStimuli.StimStopOrderCancelled
        checkAllOrdersComplete
    Case OrderStatuses.OrderStatusRejected
        mStopOrderFinished = True
        logOrderMessage "Stop order rejected (id=" & mStopOrder.Id & ")"
        mStateEngine.SetConditions Conditions.CondStopOrderCancelled
        handleStimulus StateTransitionStimuli.StimStopOrderCancelled
        checkAllOrdersComplete
End Select

Exit Sub

Err:
GHandleFatalError pProcName:="mStopOrder_StatusChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mTargetOrder Event Handlers
'@================================================================================

Private Sub mTargetOrder_Dirty()
Dim failpoint As Long
On Error GoTo Err

setDirty

Exit Sub

Err:
GHandleFatalError pProcName:="mTargetOrder_Dirty", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mTargetOrder_Fill(ByVal QuantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

Dim failpoint As Long
On Error GoTo Err

calcProfitLoss fillPrice

fireChange OrderPlexTargetOrderChanged
thisSize = IIf(mTargetOrder.Action = ActionBuy, QuantityFilled, -QuantityFilled)
setSize size + thisSize
logOrderMessage "Target order fill (id=" & _
                        mTargetOrder.Id & "): " & _
                        GOrderActionToString(mTargetOrder.Action) & " " & QuantityFilled & " at " & _
                        fillPrice

mPositionManager.NotifyExecution thisSize, fillPrice, mHandle, False

If Not mStopOrder Is Nothing Then
    If mStopOrder.Quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        AdjustStop mStopOrder.TriggerPrice, Abs(size)
    End If
End If

RaiseEvent TargetOrderPartiallyFilled

Exit Sub

Err:
GHandleFatalError pProcName:="mTargetOrder_Fill", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mTargetOrder_PropertyChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexTargetOrderChanged

Exit Sub

Err:
GHandleFatalError pProcName:="mTargetOrder_PropertyChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mTargetOrder_StatusChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexTargetOrderChanged
Select Case mTargetOrder.Status
    Case OrderStatuses.OrderStatusFilled
        logOrderMessage "Target order completed (id=" & _
                            mTargetOrder.Id & "): " & _
                            GOrderActionToString(mTargetOrder.Action) & " " & mTargetOrder.QuantityFilled & " at " & _
                            mTargetOrder.AveragePrice
        
        RaiseEvent TargetOrderFilled
        
        mTargetOrderFinished = True
        checkAllOrdersComplete
    
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusSubmitted
    Case OrderStatuses.OrderStatusCancelled
        mTargetOrderFinished = True
        logOrderMessage "Target order cancelled (id=" & mTargetOrder.Id & ")"
        mStateEngine.SetConditions Conditions.CondTargetOrderCancelled
        handleStimulus StateTransitionStimuli.StimTargetOrderCancelled
        checkAllOrdersComplete
    Case OrderStatuses.OrderStatusRejected
        mTargetOrderFinished = True
        logOrderMessage "Target order rejected (id=" & mTargetOrder.Id & ")"
        mStateEngine.SetConditions Conditions.CondTargetOrderCancelled
        handleStimulus StateTransitionStimuli.StimTargetOrderCancelled
        checkAllOrdersComplete
    
End Select

Exit Sub

Err:
GHandleFatalError pProcName:="mTargetOrder_StatusChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mTimerEntry Event Handlers
'@================================================================================

Private Sub mTimerEntry_Expired()
Dim failpoint As Long
On Error GoTo Err

logOrderMessage "Cancel time expired"
Cancel
fireChange OrderPlexSelfCancelled
RaiseEvent SelfCancelled
CancelAfter = 0

Exit Sub

Err:
GHandleFatalError pProcName:="mTimerEntry_Expired", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let CancelAfter(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

If value <> 0 Then
    mCancelTime = Now + CDbl(value) / 86400#
    Set mTimerEntry = mTicker.AddTimer(value, expiryTimeUnits.ExpiryTimeUnitSeconds, Nothing)
ElseIf Not mTimerEntry Is Nothing Then
    mCancelTime = 0
    mTicker.RemoveTimer mTimerEntry
    Set mTimerEntry = Nothing
End If

mSummary.Canceltime = mCancelTime
saveRecoveryInfo

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="CancelAfter", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let Cancelprice(ByVal value As Double)
mCancelPrice = value
End Property

Friend Property Let Canceltime(ByVal value As Date)
mCancelTime = value
End Property

Public Property Get ClosedOut() As Boolean
Dim failpoint As Long
On Error GoTo Err

ClosedOut = (mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="closedOut", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let CloseoutOrder(ByVal value As Order)
Set mCloseoutOrder = value
End Property

Public Property Get CloseoutOrder() As Order
Set CloseoutOrder = mCloseoutOrder
End Property

Public Property Get Contract() As Contract
Dim failpoint As Long
On Error GoTo Err

Set Contract = mTicker.Contract

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Contract", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let CreationTime(ByVal value As Date)
mCreationTime = value
End Property

Public Property Get CreationTime() As Date
CreationTime = mCreationTime
End Property

Public Property Let Description(ByVal value As String)
mDescription = value
mSummary.Description = value
saveRecoveryInfo
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Get Dirty() As Boolean
Dirty = mDirty
End Property

Public Property Let DrawdownNotificationThreshold(ByVal value As Currency)
mDrawdownNotificationThreshold = value
mSummary.DrawdownNotificationThreshold = value
saveRecoveryInfo
End Property

Friend Property Let EntryOrder(ByVal value As Order)
Set mEntryOrder = value
End Property

Public Property Get EntryOrder() As Order
Set EntryOrder = mEntryOrder
End Property

Public Property Let EntryReason(ByVal value As String)
Dim failpoint As Long
On Error GoTo Err

mEntryReason = value
mSummary.EntryReason = value
saveRecoveryInfo

If mHandle > 0 Then
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonEntry, mEntryReason
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="entryReason", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get EntryReason() As String
EntryReason = mEntryReason
End Property

Friend Property Let ExitStrategy( _
                ByVal value As PositionManagementStrategy)
Set mExitStrategy = value
End Property

Public Property Get IndexApplication() As Long
IndexApplication = mIndexApplication
End Property

Public Property Get IsRiskUnlimited() As Boolean
Dim failpoint As Long
On Error GoTo Err

IsRiskUnlimited = (Risk = MaxCurrency)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IsRiskUnlimited", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let IsSimulated(ByVal value As Boolean)
mIsSimulated = value
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Let Key(ByVal value As String)
mKey = value
mSummary.Key = mKey
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Get LongPosition() As Boolean
Dim failpoint As Long
On Error GoTo Err

LongPosition = (mEntryOrder.Action = ActionBuy)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="longPosition", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let LossNotificationThreshold(ByVal value As Currency)
mLossNotificationThreshold = value
mSummary.LossNotificationThreshold = value
saveRecoveryInfo
End Property

Public Property Let NewEntryPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewEntryPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                "OrderPlex::(let)NewEntryPrice", _
                                                "Can't change entry price - order " & EntryOrder.Id & " alReady filled"

Select Case mEntryOrder.OrderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeTrail, _
        OrderTypes.OrderTypeAutoStop, _
        OrderTypes.OrderTypeAutoLimit, _
        OrderTypes.OrderTypeStop
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "OrderPlex::(let)NewEntryPrice", _
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewEntryPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewEntryTriggerPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewEntryTriggerPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                "OrderPlex::(let)NewEntryTriggerPrice", _
                                                "Can't change entry trigger price - order " & EntryOrder.Id & " alReady filled"

Select Case mEntryOrder.OrderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeLimit, _
        OrderTypes.OrderTypeTrail, _
        OrderTypes.OrderTypeAutoLimit
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "OrderPlex::(let)NewEntryTriggerPrice", _
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewEntryTriggerPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewOrderPrice( _
                ByVal orderIndex As Long, _
                ByVal value As Double)
Dim lOrder As Order

Dim failpoint As Long
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewOrderPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewOrderTriggerPrice( _
                ByVal orderIndex As Long, _
                ByVal value As Double)
Dim lOrder As Order

Dim failpoint As Long
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewOrderTriggerPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewQuantity(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewQuantity", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                "OrderPlex::(let)NewQuantity", _
                                                "Can't change quantity - order " & EntryOrder.Id & " alReady filled"

If value <> mEntryOrder.Quantity Then
    logOrderMessage "Set quantity=" & value
    mNewQuantity = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewQuantity", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewStopOffset(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewStopOffset", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewStopOffset", _
                                        "No stop order exists"

If Not mStopOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "OrderPlex::(let)NewStopOffset", _
                    "Can't change offset for order type " & GOrderTypeToString(mEntryOrder.OrderType)
End If

If mStopOrder.Status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewStopOffset", _
                                        "Can't change stop offset - order " & mStopOrder.Id & " alReady placed"

If value <> mStopOrder.Offset Then
    logOrderMessage "Set stop offset=" & value
    mNewStopOffset = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewStopOffset", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewStopPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewStopPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewStopPrice", _
                                        "No stop order exists"

Select Case mStopOrder.OrderType
Case OrderTypes.OrderTypeStopLimit
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewStopPrice", _
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewStopPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewStopQuantity(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewStopQuantity", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewStopQuantity", _
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewStopQuantity", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewStopTriggerPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewStopTriggerPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewStopTriggerPrice", _
                                        "No stop order exists"

Select Case mStopOrder.OrderType
    Case OrderTypes.OrderTypeStop, _
        OrderTypes.OrderTypeStopLimit
    Case Else
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "OrderPlex::(let)NewStopTriggerPrice", _
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewStopTriggerPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewTargetPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewTargetPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewTargetPrice", _
                                        "No target order exists"

Select Case mTargetOrder.OrderType
Case OrderTypes.OrderTypeLimitIfTouched, _
    OrderTypes.OrderTypeLimit
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewTargetPrice", _
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewTargetPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewTargetOffset(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewTargetOffset", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewTargetOffset", _
                                        "No target order exists"

If Not mTargetOrder.OrderType = OrderTypes.OrderTypeAutoLimit Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "OrderPlex::(let)NewTargetOffset", _
                    "Can't change offset for order type " & GOrderTypeToString(mTargetOrder.OrderType)
End If

If mTargetOrder.Status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewTargetOffset", _
                                        "Can't change target offset - order " & mTargetOrder.Id & " alReady placed"

If value <> mTargetOrder.Offset Then
    logOrderMessage "Set target offset=" & value
    mNewTargetOffset = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewTargetOffset", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewTargetQuantity(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewTargetQuantity", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewTargetQuantity", _
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewTargetQuantity", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let NewTargetTriggerPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewTargetTriggerPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)NewTargetTriggerPrice", _
                                        "No target order exists"

Select Case mTargetOrder.OrderType
Case OrderTypes.OrderTypeLimitIfTouched, _
    OrderTypes.OrderTypeMarketIfTouched
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)NewTargetTriggerPrice", _
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="NewTargetTriggerPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get PendingSize() As Long
Dim failpoint As Long
On Error GoTo Err

If mEntryOrder Is Nothing Then PendingSize = 0: Exit Property
If mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed Then PendingSize = 0: Exit Property
PendingSize = IIf(LongPosition, mEntryOrder.QuantityRemaining, -mEntryOrder.QuantityRemaining)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="pendingSize", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Order(ByVal index As Long) As Order
Dim failpoint As Long
On Error GoTo Err

Set Order = mOrders(index)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="order", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get OrderContext() As OrderContext
Dim failpoint As Long
On Error GoTo Err

Set OrderContext = mOrderContext

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="OrderContext", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let OrderContext(ByVal value As OrderContext)
Set mOrderContext = value
End Property

Friend Property Let orderSubmitter( _
                ByVal value As TradeBuildSP.IOrderSubmitter)
Set mOrderSubmitter = value
End Property

Public Property Get NumberOfOrders() As Long
Dim failpoint As Long
On Error GoTo Err

NumberOfOrders = mOrders.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="numberOfOrders", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get PositionManager() As PositionManager
Set PositionManager = mPositionManager
End Property

Friend Property Let PositionManager(ByVal value As PositionManager)
Set mPositionManager = value
End Property

Public Property Let PreventUnprotectedPosition( _
                ByVal value As Boolean)
Dim failpoint As Long
On Error GoTo Err

If State <> OrderPlexStateCreated And _
    mStopOrder Is Nothing _
Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "PreventUnprotectedPosition", _
            "No stop loss order exists"
End If

mPreventUnprotectedPosition = value
mSummary.PreventUnprotectedPosition = value
saveRecoveryInfo

If mPreventUnprotectedPosition Then
    mStateEngine.SetConditions Conditions.CondProtected
Else
    mStateEngine.ClearConditions Conditions.CondProtected
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="PreventUnprotectedPosition", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get PreventUnprotectedPosition() As Boolean
Dim failpoint As Long
On Error GoTo Err

PreventUnprotectedPosition = mPreventUnprotectedPosition

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="PreventUnprotectedPosition", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let ProfitNotificationThreshold(ByVal value As Currency)
mProfitNotificationThreshold = value
mSummary.ProfitNotificationThreshold = value
saveRecoveryInfo
End Property

Public Property Get Risk() As Currency
Dim stopTotalPrice As Double
Dim pendingTotalPrice As Double
Dim filledTotalPrice As Double

Dim failpoint As Long
On Error GoTo Err

If mStopOrder Is Nothing Then
    Risk = MaxCurrency
    Exit Property
End If

If mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed Then Exit Property

If mStopOrder.Status = OrderStatuses.OrderStatusFilled Then Exit Property

' the amount at Risk has two parts - that due to fills alReady obtained, and
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Risk", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Property

Friend Property Let ScalingStrategy( _
                ByVal value As PositionManagementStrategy)
Set mScalingStrategy = value
End Property

Friend Property Let size(ByVal value As Long)
mSize = value
End Property

Public Property Get size() As Long
size = mSize
End Property

Friend Property Let State(ByVal value As OrderPlexStateCodes)
Set mStateEngine = CreateStateEngine(GOrderPlex.TableBuilder, _
                                    value)
End Property

Public Property Get State() As OrderPlexStateCodes
Dim failpoint As Long
On Error GoTo Err

State = mStateEngine.State

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="State", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let StopOrder(ByVal value As Order)
Set mStopOrder = value
End Property

Public Property Get StopOrder() As Order
Set StopOrder = mStopOrder
End Property

Public Property Let StopReason(ByVal value As String)
Dim failpoint As Long
On Error GoTo Err

mStopReason = value
mSummary.StopReason = value
saveRecoveryInfo

If mHandle > 0 Then
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonStop, mStopReason
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="stopReason", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get StopReason() As String
StopReason = mStopReason
End Property

Friend Property Let StopStrategy( _
                ByVal value As StopManagementStrategy)
Set mStopStrategy = value
End Property

Friend Property Let TargetOrder(ByVal value As Order)
Set mTargetOrder = value
End Property

Public Property Get TargetOrder() As Order
Set TargetOrder = mTargetOrder
End Property

Public Property Let TargetReason(ByVal value As String)
Dim failpoint As Long
On Error GoTo Err

mTargetReason = value
mSummary.TargetReason = value
saveRecoveryInfo

If mHandle > 0 Then
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonTarget, mTargetReason
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="targetReason", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get TargetReason() As String
TargetReason = mTargetReason
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Dim lContract As Contract
Dim failpoint As Long
On Error GoTo Err

Set mTicker = value
Set lContract = mTicker.Contract
mTickSize = lContract.TickSize
mTickValue = lContract.tickValue
mValueFactor = lContract.multiplier
mcurrencyCode = lContract.specifier.currencyCode

mSummary.WorkspaceName = mTicker.Workspace.name
mSummary.TickerKey = mTicker.Key

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Ticker", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddChangeListener(ByVal value As ChangeListener)
Dim failpoint As Long
On Error GoTo Err

mChangeListeners.Add value

' notify this new listenener that we are alReady created
Dim ev As ChangeEvent
Set ev.source = Me
ev.changeType = OrderPlexChangeTypes.OrderPlexCreated
value.Change ev

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddChangeListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub AddStateChangeListener(ByVal value As StateChangeListener)
Dim failpoint As Long
On Error GoTo Err

mStateChangeListeners.Add value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddStateChangeListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub AddProfitListener(ByVal value As ProfitListener)
Dim failpoint As Long
On Error GoTo Err

mProfitListeners.Add value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="AddProfitListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub AdjustStop(ByVal stopTriggerPrice As Double, _
                    Optional ByVal Quantity As Long = -1)
Dim failpoint As Long
On Error GoTo Err

NewStopTriggerPrice = stopTriggerPrice
If Quantity <> -1 Then NewStopQuantity = Quantity
Update

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="adjustStop", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Public Sub AdjustTarget(ByVal targetPrice As Double, _
                    Optional ByVal Quantity As Long = -1)
Dim failpoint As Long
On Error GoTo Err

NewTargetPrice = targetPrice
If Quantity <> -1 Then NewTargetQuantity = Quantity
Update

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="adjustTarget", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Public Function Cancel(Optional ByVal evenIfFilled As Boolean = False) As Boolean

Dim failpoint As Long
On Error GoTo Err

If evenIfFilled Then
    mStateEngine.ClearConditions Conditions.CondNoFillCancellation
    Cancel = handleStimulus(StateTransitionStimuli.StimCancelEvenIfFill)
Else
    mStateEngine.SetConditions Conditions.CondNoFillCancellation
    Cancel = handleStimulus(StateTransitionStimuli.StimCancelIfNoFill)
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Cancel", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

''
' Cancels any changes made to this <code>orderPlex</code> that have not yet
' been actioned.
'@/
Public Sub CancelChanges()
Dim failpoint As Long
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="cancelChanges", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

''
' Copies the current state of the member orders for use by cancelChanges
'@/
Friend Sub Checkpoint()
Dim failpoint As Long
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="checkpoint", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub CloseOut(ByVal method As String)

Dim failpoint As Long
On Error GoTo Err

If UCase(method) <> "MKT" Then
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, "OrderPlex::closeout", "method " & method & " not implemented"
End If

logOrderMessage "Closing out: " & _
                    "size=" & size & _
                    "; pending size=" & PendingSize

handleStimulus StateTransitionStimuli.StimCloseout

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="CloseOut", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Friend Function createEntryOrder() As Order
Dim failpoint As Long
On Error GoTo Err

If Not mEntryOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::createEntryOrder", _
                "Entry order alReady exists"

Set mEntryOrder = mOrderContext.newOrder
mSummary.EntryOrderId = mEntryOrder.Id
saveRecoveryInfo

Set createEntryOrder = mEntryOrder

If mOrders.Count >= 1 Then
    mOrders.Add mEntryOrder, , 1
Else
    mOrders.Add mEntryOrder
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="createEntryOrder", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Friend Function createStopOrder() As Order
Dim failpoint As Long
On Error GoTo Err

If Not mStopOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::createStopOrder", _
                "Stop order alReady exists"

Set mStopOrder = mOrderContext.newOrder
mSummary.StopOrderId = mStopOrder.Id
saveRecoveryInfo

Set createStopOrder = mStopOrder
mStateEngine.SetConditions Conditions.CondStopOrderExists
mStateEngine.ClearConditions Conditions.CondStopOrderCancelled

If Not mEntryOrder Is Nothing Then
    mOrders.Add mStopOrder, , , 1
ElseIf Not mTargetOrder Is Nothing Then
    mOrders.Add mStopOrder, , 1
Else
    mOrders.Add mStopOrder
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="createStopOrder", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Friend Function createTargetOrder() As Order
Dim failpoint As Long
On Error GoTo Err

If Not mTargetOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::createTargetOrder", _
                "Target order alReady exists"

Set mTargetOrder = mOrderContext.newOrder
mSummary.TargetOrderId = mTargetOrder.Id
saveRecoveryInfo

Set createTargetOrder = mTargetOrder
mStateEngine.SetConditions Conditions.CondTargetOrderExists
mStateEngine.ClearConditions Conditions.CondTargetOrderCancelled

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

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="createTargetOrder", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Friend Sub Execute(Optional ByVal Cancelprice As Double = 0)
Dim failpoint As Long
On Error GoTo Err

If mPreventUnprotectedPosition And mStopOrder Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "execute", _
            "No stop loss order exists"
End If

mCancelPrice = Cancelprice
mSummary.Cancelprice = mCancelPrice

mCreationTime = mTicker.Timestamp
mSummary.CreationTime = mCreationTime

saveRecoveryInfo

fireChange OrderPlexChangeTypes.OrderPlexCreated
handleStimulus StateTransitionStimuli.StimExecute

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="execute", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub finish()
Dim failpoint As Long
On Error GoTo Err

Set mTicker = Nothing
Set mPositionManager = Nothing
Set mOrderContext = Nothing
Set mOrderSubmitter = Nothing
If Not mEntryOrder Is Nothing Then mEntryOrder.finish
If Not mStopOrder Is Nothing Then mStopOrder.finish
If Not mTargetOrder Is Nothing Then mTargetOrder.finish
If Not mCloseoutOrder Is Nothing Then mCloseoutOrder.finish
Set mStopStrategy = Nothing
Set mScalingStrategy = Nothing
Set mExitStrategy = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Finish", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveChangeListener(ByVal value As ChangeListener)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveChangeListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveStateChangeListener(ByVal value As StateChangeListener)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

For i = mStateChangeListeners.Count To 1 Step -1
    If mStateChangeListeners.Item(i) Is value Then mStateChangeListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveStateChangeListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveProfitListener(ByVal value As ProfitListener)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

For i = mProfitListeners.Count To 1 Step -1
    If mProfitListeners.Item(i) Is value Then mProfitListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveProfitListener", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub Update()
Dim updateEntry As Boolean
Dim updateStop As Boolean
Dim updateTarget As Boolean

Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCreated
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Sub
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::update", _
                "Action not allowed in state " & mStateEngine.State
End Select

checkForPendingUpdates updateEntry, updateStop, updateTarget

If (Not updateEntry) And (Not updateStop) And (Not updateTarget) Then
    ' no updates to action
    Exit Sub
End If

logOrderMessage "Update order plex"

If State <> OrderPlexStateCreated Then
    mOrderSubmitter.modifyOrderPlex Me, updateEntry, updateStop, updateTarget, False
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="update", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcProfitLoss(ByVal value As Double)
Dim Profit As Currency
Dim notifyProfit As Boolean
Dim notifyMaxProfit As Boolean
Dim notifyDrawdown As Boolean

Dim failpoint As Long
On Error GoTo Err

Profit = size * (value - mEntryOrder.AveragePrice) * mValueFactor

If Profit <> mProfit Then
    mProfit = Profit
    notifyProfit = True
End If

If Profit > mMaxProfit Then
    mMaxProfit = Profit
    notifyMaxProfit = True
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
If notifyDrawdown Then fireProfit mDrawdown, ProfitTypes.ProfitTypeDrawdown

If notifyProfit Then
    If mProfitNotificationThreshold <> 0 And _
        mProfit >= mProfitNotificationThreshold _
    Then
        fireChange OrderPlexProfitThresholdExceeded
        RaiseEvent ProfitThresholdExceeded
        mProfitNotificationThreshold = 0
        mSummary.ProfitNotificationThreshold = 0
        saveRecoveryInfo
    End If
    
    If mLossNotificationThreshold <> 0 And _
        mProfit <= mLossNotificationThreshold _
    Then
        fireChange OrderPlexLossThresholdExceeded
        RaiseEvent LossThresholdExceeded
        mLossNotificationThreshold = 0
        mSummary.LossNotificationThreshold = 0
        saveRecoveryInfo
End If

    If mDrawdownNotificationThreshold <> 0 And _
        (mMaxProfit - Profit >= mDrawdownNotificationThreshold) _
    Then
        fireChange OrderPlexDrawdownThresholdExceeded
        RaiseEvent DrawdownThresholdExceeded
        mDrawdownNotificationThreshold = 0
        mSummary.DrawdownNotificationThreshold = 0
        saveRecoveryInfo
    End If
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="calcProfitLoss", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function cancelOrders() As Boolean
Dim StopOrderId As String
Dim TargetOrderId As String

Dim failpoint As Long
On Error GoTo Err

CancelAfter = 0
mCancelPrice = 0

Select Case mEntryOrder.Status
Case OrderStatuses.OrderStatusCreated
    mEntryOrder.updateStatus OrderStatuses.OrderStatusCancelled

    cancelOrders = True
End Select

If Not mStopOrder Is Nothing Then
    Select Case mStopOrder.Status
    Case OrderStatuses.OrderStatusCreated
        mStopOrder.updateStatus OrderStatuses.OrderStatusCancelled
        cancelOrders = True
    End Select
End If

If Not mTargetOrder Is Nothing Then
    Select Case mTargetOrder.Status
    Case OrderStatuses.OrderStatusCreated
        mTargetOrder.updateStatus OrderStatuses.OrderStatusCancelled
        cancelOrders = True
    End Select
End If

If mOrderSubmitter.cancelOrderPlex(Me) Then
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="cancelOrders", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function

Private Function cancelStateTimeout()
Dim failpoint As Long
On Error GoTo Err

mStateTimeoutTimer.StopTimer

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="cancelStateTimeout", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function cancelStopOrder()
Dim failpoint As Long
On Error GoTo Err

If mStopOrder Is Nothing Then Exit Function
Select Case mStopOrder.Status
Case OrderStatuses.OrderStatusCreated
    mStopOrder.updateStatus OrderStatuses.OrderStatusCancelled
Case OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    logOrderMessage "Cancel order: id=" & mStopOrder.Id
    mOrderSubmitter.cancelStopOrder Me
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="cancelStopOrder", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function cancelTargetOrder()
Dim failpoint As Long
On Error GoTo Err

If mTargetOrder Is Nothing Then Exit Function
Select Case mTargetOrder.Status
Case OrderStatuses.OrderStatusCreated
    mTargetOrder.updateStatus OrderStatuses.OrderStatusCancelled
Case OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    logOrderMessage "Cancel order: id=" & mTargetOrder.Id
    mOrderSubmitter.cancelTargetOrder Me
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="cancelTargetOrder", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub checkAllOrdersComplete()
Dim failpoint As Long
On Error GoTo Err

If mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed Then Exit Sub
If mEntryOrderFinished And _
    mStopOrderFinished And _
    mTargetOrderFinished _
Then
    handleStimulus StateTransitionStimuli.StimAllOrdersComplete
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="checkAllOrdersComplete", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub checkForPendingUpdates( _
                ByRef updateEntry As Boolean, _
                ByRef updateStop As Boolean, _
                ByRef updateTarget As Boolean)
Dim failpoint As Long
On Error GoTo Err

If Not mEntryOrder Is Nothing Then
    If mEntryOrder.IsDirty Then updateEntry = True
    If mNewQuantity >= 0 And _
        mNewQuantity <> mEntryOrder.Quantity _
    Then
        If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change Size - order " & EntryOrder.Id & " alReady filled"
        mEntryOrder.Quantity = mNewQuantity
        updateEntry = True
    End If
    mNewQuantity = -1
    
    If mNewEntryPrice >= 0 And _
        mNewEntryPrice <> mEntryOrder.LimitPrice _
    Then
        If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change price - order " & mEntryOrder.Id & " alReady filled"
        mEntryOrder.LimitPrice = mNewEntryPrice
        updateEntry = True
    End If
    mNewEntryPrice = -1
    
    If mNewEntryTriggerPrice >= 0 And _
        mNewEntryTriggerPrice <> mEntryOrder.TriggerPrice _
    Then
        If mEntryOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.Id & " alReady filled"
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
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mStopOrder.Id & " alReady filled"
        If mStopOrder.Status = OrderStatuses.OrderStatusSubmitted Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change price - order " & mStopOrder.Id & " alReady submitted"
        mStopOrder.LimitPrice = mNewStopPrice
        updateStop = True
    End If
    mNewStopPrice = -1
    
    If mNewStopOffset <> DummyOffset And _
        mNewStopOffset <> mStopOrder.Offset _
    Then
        If mStopOrder.Status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            "OrderPlex::update", _
                                                            "Can't change stop offset - order " & mStopOrder.Id & " alReady placed"
        mStopOrder.Offset = mNewStopOffset
    End If
    mNewStopOffset = DummyOffset

    If mNewStopQuantity >= 0 And _
        mNewStopQuantity <> mStopOrder.Quantity _
    Then
        If mStopOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            "OrderPlex::update", _
                                                            "Can't change stop quantity - order " & mStopOrder.Id & " alReady placed"
        mStopOrder.Quantity = mNewStopQuantity
    End If
    mNewStopQuantity = -1

    If mNewStopTriggerPrice >= 0 And _
        mNewStopTriggerPrice <> mStopOrder.TriggerPrice _
    Then
        If mStopOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.Id & " alReady filled"
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
                                                        "OrderPlex::update", _
                                                        "Can't change price - order " & mTargetOrder.Id & " alReady filled"
        mTargetOrder.LimitPrice = mNewTargetPrice
        updateTarget = True
    End If
    mNewTargetPrice = -1
    
    If mNewTargetOffset <> DummyOffset And _
        mNewTargetOffset <> mTargetOrder.Offset _
    Then
        If mTargetOrder.Status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            "OrderPlex::update", _
                                                            "Can't change target offset - order " & TargetOrder.Id & " alReady placed"
        mTargetOrder.Offset = mNewTargetOffset
    End If
    mNewTargetOffset = DummyOffset

    If mNewTargetQuantity >= 0 And _
        mNewTargetQuantity <> mTargetOrder.Quantity _
    Then
        If mTargetOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            "OrderPlex::update", _
                                                            "Can't change target quantity - order " & mTargetOrder.Id & " alReady placed"
        mTargetOrder.Quantity = mNewTargetQuantity
    End If
    mNewTargetQuantity = -1

    If mNewTargetTriggerPrice >= 0 And _
        mNewTargetTriggerPrice <> mTargetOrder.TriggerPrice _
    Then
        If mTargetOrder.Status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.Id & " alReady filled"
        mTargetOrder.TriggerPrice = mNewTargetTriggerPrice
        updateTarget = True
    End If
    mNewTargetTriggerPrice = -1
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="checkForPendingUpdates", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Function completionActions()
Dim failpoint As Long
On Error GoTo Err

mOrderContext.orderPlexClosed Me
fireChange OrderPlexCompleted
RaiseEvent Completed
Set mStopStrategy = Nothing
Set mScalingStrategy = Nothing
Set mExitStrategy = Nothing
If size = 0 Then mTicker.RemoveQuoteListener Me

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="completionActions", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function doActions(ByRef pActions() As Long) As Variant
Dim Action As Variant

Dim failpoint As Long
On Error GoTo Err

For Each Action In pActions
    Select Case CLng(Action)
    Case SpecialActions.NoAction
    Case Actions.ActPlaceOrders
        doActions = placeOrders
    Case Actions.ActCancelOrders
        doActions = cancelOrders
    Case Actions.ActCancelStopOrder
        doActions = cancelStopOrder
    Case Actions.ActCancelTargetOrder
        doActions = cancelTargetOrder
    Case Actions.ActResubmitStopOrder
        doActions = resubmitStopOrder
    Case Actions.ActResubmitTargetOrder
        doActions = resubmitTargetOrder
    Case Actions.ActResubmitStopAndTargetOrders
        doActions = resubmitStopAndTargetOrders
    Case Actions.ActPlaceCloseoutOrder
        doActions = placeCloseoutOrder
    Case Actions.ActCompletionActions
        doActions = completionActions
    Case Actions.ActAlarm
        ' need to put something in here
    Case Actions.ActSetTimeout
        doActions = setStateTimeout
    Case Actions.ActCancelTimeout
        doActions = cancelStateTimeout
    Case Else
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "OrderPlex::doActions", _
                    "Invalid action " & Action
    End Select
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="doActions", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub fireChange( _
                ByVal changeType As OrderPlexChangeTypes)
Dim listener As ChangeListener
Dim ev As ChangeEvent
Dim failpoint As Long
On Error GoTo Err

Set ev.source = Me
ev.changeType = changeType
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent Change(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireChange", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireStateChange( _
                ByVal State As OrderPlexStateCodes)
Dim listener As StateChangeListener
Dim ev As StateChangeEvent
Dim failpoint As Long
On Error GoTo Err

mSummary.State = State
saveRecoveryInfo

Set ev.source = Me
ev.State = State
For Each listener In mStateChangeListeners
    listener.Change ev
Next
RaiseEvent StateChange(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireStateChange", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub fireProfit( _
                ByVal profitAmount As Double, _
                ByVal profitType As ProfitTypes)
Dim listener As ProfitListener
Dim Profit As ProfitEvent
Dim failpoint As Long
On Error GoTo Err

Set Profit.source = Me
Profit.profitAmount = profitAmount
Profit.profitType = profitType
Profit.currencyCode = mcurrencyCode
For Each listener In mProfitListeners
    listener.profitAmount Profit
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="fireProfit", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Function handleStimulus(ByVal stimulus As StateTransitionStimuli) As Variant
Dim currState As OrderPlexStateCodes

Dim failpoint As Long
On Error GoTo Err

currState = mStateEngine.State
handleStimulus = doActions(mStateEngine.NotifyStimulus(stimulus))
If mStateEngine.State <> currState Then
    fireChange OrderPlexChangeTypes.OrderPlexStateChanged
    fireStateChange mStateEngine.State
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="handleStimulus", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub logOrderMessage( _
                ByRef msg As String)
GLogData _
                    msg & " (Curr price=" & mTicker.FormatPrice(mTicker.TradePrice) & "; id=" & Key & ")", _
                    IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                    Me
End Sub

Private Function placeCloseoutOrder()
' place an order to closeout the filled part
Dim failpoint As Long
On Error GoTo Err

Set mCloseoutOrder = mOrderContext.newOrder
mSummary.CloseoutOrderId = mCloseoutOrder.Id
saveRecoveryInfo

mCloseoutOrder.Action = IIf(LongPosition, OrderActions.ActionSell, OrderActions.ActionBuy)
mCloseoutOrder.Quantity = Abs(size)
mCloseoutOrder.OrderType = OrderTypes.OrderTypeMarket
mCloseoutOrder.TimeInForce = OrderTifs.TIFGoodTillCancelled

mOrders.Add mCloseoutOrder

fireChange OrderPlexCloseoutOrderCreated
mOrderSubmitter.modifyOrderPlex Me, False, False, False, True
                                

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="placeCloseoutOrder", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function

Private Function placeOrders()

Dim failpoint As Long
On Error GoTo Err

If mEntryOrder Is Nothing Then mEntryOrderFinished = True
If mStopOrder Is Nothing Then mStopOrderFinished = True
If mTargetOrder Is Nothing Then mTargetOrderFinished = True

mOrderSubmitter.ExecuteOrderPlex Me
mTicker.AddQuoteListener Me

Checkpoint

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="placeOrders", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub placePeggedOrders()
Dim failpoint As Long
On Error GoTo Err

If Not mStopOrder Is Nothing Then
    If mStopOrder.Status = OrderStatuses.OrderStatusCreated Then
        If mStopOrder.OrderType = OrderTypes.OrderTypeAutoStop Then
            ' the stop price is to be pegged to the entry price - the
            ' order's offset property will alReady contain any required
            ' offset from that price
            mStopOrder.OrderType = OrderTypes.OrderTypeStop
            NewStopTriggerPrice = (mStopOrder.Offset * mTickSize) + EntryOrder.LastFillPrice
            NewStopQuantity = mEntryOrder.QuantityFilled
        End If
    End If
End If
If Not mTargetOrder Is Nothing Then
    If mTargetOrder.Status = OrderStatuses.OrderStatusCreated Then
        If mTargetOrder.OrderType = OrderTypes.OrderTypeAutoLimit Then
            ' the limit price is to be pegged to the entry price - the
            ' order's offset property will alReady contain any required
            ' offset from that price
            mTargetOrder.OrderType = OrderTypes.OrderTypeLimit
            NewTargetPrice = (mTargetOrder.Offset * mTickSize) + EntryOrder.LastFillPrice
            NewTargetQuantity = mEntryOrder.QuantityFilled
        End If
    End If
End If
Update

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="placePeggedOrders", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Function resubmitStopOrder()
Dim failpoint As Long
On Error GoTo Err

mOrderSubmitter.resubmitStopOrder Me
logOrderMessage "Resubmit stop order: id=" & mStopOrder.Id & _
            "; new broker id=" & mStopOrder.BrokerId
mStateEngine.ClearConditions Conditions.CondStopOrderCancelled
mStopOrderFinished = False

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="resubmitStopOrder", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function resubmitStopAndTargetOrders()
Dim failpoint As Long
On Error GoTo Err

mOrderSubmitter.resubmitStopAndTargetOrders Me
logOrderMessage "Resubmit stop order: id=" & mStopOrder.Id & _
            "; new broker id=" & mStopOrder.BrokerId
logOrderMessage "Resubmit target order: id=" & mTargetOrder.Id & _
            "; new broker id=" & mTargetOrder.BrokerId
mStateEngine.ClearConditions Conditions.CondStopOrderCancelled
mStopOrderFinished = False
mStateEngine.ClearConditions Conditions.CondTargetOrderCancelled
mTargetOrderFinished = False

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="resubmitStopAndTargetOrders", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function

Private Function resubmitTargetOrder()
Dim failpoint As Long
On Error GoTo Err

mOrderSubmitter.resubmitTargetOrder Me
logOrderMessage "Resubmit target order: id=" & mTargetOrder.Id & _
            "; new broker id=" & mTargetOrder.BrokerId
mStateEngine.ClearConditions Conditions.CondTargetOrderCancelled
mTargetOrderFinished = False

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="resubmitTargetOrder", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub saveRecoveryInfo()
If mStateEngine.State = OrderPlexStateCodes.OrderPlexStateCreated Then Exit Sub
If gTB.OrderPlexRecoverer.OrderPersistenceServiceProvider Is Nothing Then Exit Sub

gTB.OrderPlexRecoverer.OrderPersistenceServiceProvider.SaveOrderInfo mSummary.ToJSON
End Sub

''
' Set the 'dirty' flag, and raise a Dirty event if one has not alReady
' been raised.
'@/
Private Sub setDirty()
Dim failpoint As Long
On Error GoTo Err

If Not mDirty Then
    mDirty = True
    RaiseEvent Dirty
Else
    mDirty = True
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setDirty", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub setSize(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

mSize = value
mSummary.size = value
saveRecoveryInfo

If mSize <> 0 Then
    mStateEngine.SetConditions Conditions.CondSizeNonZero
Else
    mStateEngine.ClearConditions Conditions.CondSizeNonZero
End If
fireChange OrderPlexChangeTypes.OrderPlexSizeChanged

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setSize", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Function setStateTimeout()
Dim failpoint As Long
On Error GoTo Err

Set mStateTimeoutTimer = CreateIntervalTimer(1000)
mStateTimeoutTimer.StartTimer

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setStateTimeout", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

