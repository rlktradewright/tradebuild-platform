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

Private Const ProjectName                As String = "TradeBuild26"
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
    ' orders and no net size - note that this may require a closeout
    ' order to be placed to negate any existing size.)
    OrderPlexStateClosingOut
        
    ' This state indicates that all orders have been completed.
    OrderPlexStateClosed

    ' This state indicates that a cancellation notification for either
    ' the stop order or the target order, when both exis, has been received.
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

Private WithEvents mEntryOrder As order
Attribute mEntryOrder.VB_VarHelpID = -1
Private mEntryOrderCopy As order
Private mEntryOrderFinished As Boolean

Private WithEvents mStopOrder As order
Attribute mStopOrder.VB_VarHelpID = -1
Private mStopOrderCopy As order
Attribute mStopOrderCopy.VB_VarHelpID = -1
Private mStopOrderFinished As Boolean

Private WithEvents mTargetOrder As order
Attribute mTargetOrder.VB_VarHelpID = -1
Private mTargetOrderCopy As order
Attribute mTargetOrderCopy.VB_VarHelpID = -1
Private mTargetOrderFinished As Boolean

Private WithEvents mCloseoutOrder As order
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

mIndexApplication = GOrderPlex.gNextApplicationIndex

Set mOrders = New Collection

Set mChangeListeners = New Collection
Set mStateChangeListeners = New Collection

Set mProfitListeners = New Collection

Set mStateEngine = CreateStateEngine(GOrderPlex.tableBuilder, _
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

Private Property Get IOrderPlex_entryOrder() As TradeBuildSP.IOrder
Set IOrderPlex_entryOrder = mEntryOrder
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

Set IOrderPlex_Order = order(index)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="IOrderPlex_Order", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Private Property Get IOrderPlex_stoporder() As TradeBuildSP.IOrder
Set IOrderPlex_stoporder = mStopOrder
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

If Not longPosition Then calcProfitLoss ev.price

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_Ask", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub QuoteListener_Bid(ev As QuoteEvent)
Dim failpoint As Long
On Error GoTo Err

If longPosition Then calcProfitLoss ev.price

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

If longPosition Then
    Select Case mEntryOrder.orderType
    Case OrderTypes.OrderTypeLimit, _
            OrderTypes.OrderTypeLimitIfTouched, _
            OrderTypes.OrderTypeMarketIfTouched
        If ev.price >= mCancelPrice Then cancelPriceHit = True
    Case OrderTypes.OrderTypeStop, _
        OrderTypes.OrderTypeStopLimit
        If ev.price <= mCancelPrice Then cancelPriceHit = True
    End Select
Else
    Select Case mEntryOrder.orderType
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
    gTB.notifyListeners _
                        "Cancel price hit (" & mTicker.TradePrice & ")", _
                        IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                        Me
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

Private Sub mCloseoutOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

Dim failpoint As Long
On Error GoTo Err

calcProfitLoss fillPrice

fireChange OrderPlexCloseoutOrderChanged
thisSize = IIf(mCloseoutOrder.action = ActionBuy, quantityFilled, -quantityFilled)
setSize size + thisSize
gTB.notifyListeners _
                    "Closeout order fill (" & _
                    mCloseoutOrder.id & "): " & _
                    gOrderActionToString(mCloseoutOrder.action) & " " & quantityFilled & " at " & _
                    fillPrice, _
                    IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                    Me

mPositionManager.NotifyExecution thisSize, fillPrice, mHandle, True

RaiseEvent CloseoutOrderPartiallyFilled

Exit Sub

Err:
gHandleFatalError pProcName:="mCloseoutOrder_Fill", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mCloseoutOrder_PropertyChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexCloseoutOrderChanged

Exit Sub

Err:
gHandleFatalError pProcName:="mCloseoutOrder_PropertyChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mCloseoutOrder_StatusChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexCloseoutOrderChanged
Select Case mCloseoutOrder.status
    Case OrderStatuses.OrderStatusFilled
        gTB.notifyListeners _
                            "Closeout order completed (" & _
                            mCloseoutOrder.id & "): " & _
                            gOrderActionToString(mCloseoutOrder.action) & " " & mCloseoutOrder.quantityFilled & " at " & _
                            mCloseoutOrder.averagePrice, _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
                            
        checkAllOrdersComplete
        
        If size <> 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, "OrderPlex::mCloseoutOrder_StatusChanged", "Closeout incomplete: size remaining is " & size
        
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusSubmitted
    Case OrderStatuses.OrderStatusCancelled
        gTB.notifyListeners _
                            "Closeout order cancelled (" & _
                            mCloseoutOrder.id & ")", _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
    Case OrderStatuses.OrderStatusRejected
        gTB.notifyListeners _
                            "Closeout order rejected (" & _
                            mCloseoutOrder.id & ")", _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
End Select

Exit Sub

Err:
gHandleFatalError pProcName:="mCloseoutOrder_StatusChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
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
gHandleFatalError pProcName:="mEntryOrder_Dirty", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mEntryOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long
Dim Risk As Currency

Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexEntryOrderChanged

cancelAfter = 0

mCancelPrice = 0

thisSize = IIf(longPosition, quantityFilled, -quantityFilled)
setSize size + thisSize


gTB.notifyListeners _
                    "Entry order fill (" & _
                    mEntryOrder.id & "): " & _
                    gOrderActionToString(mEntryOrder.action) & " " & quantityFilled & " at " & _
                    fillPrice, _
                    IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                    Me

If mHandle < 0 And Not mStopOrder Is Nothing Then
    If longPosition Then
        If mStopOrder.orderType = OrderTypes.OrderTypeAutoStop Then
            Risk = (mEntryOrder.quantity + mEntryOrder.quantityFilled) * mStopOrder.offset * mTickValue
        Else
            Risk = (mEntryOrder.quantity + mEntryOrder.quantityFilled) * (fillPrice - mStopOrder.triggerPrice) * mValueFactor
        End If
    Else
        If mStopOrder.orderType = OrderTypes.OrderTypeAutoStop Then
            Risk = -(mEntryOrder.quantity + mEntryOrder.quantityFilled) * mStopOrder.offset * mTickValue
        Else
            Risk = (mEntryOrder.quantity + mEntryOrder.quantityFilled) * (mStopOrder.triggerPrice - fillPrice) * mValueFactor
        End If
    End If
End If

If mHandle < 0 Then
    mHandle = mPositionManager.RegisterPlex(mKey, mDescription, mEntryOrder.action, Risk)
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonEntry, mEntryReason
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonTarget, mTargetReason
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonStop, mStopReason
End If
mPositionManager.NotifyExecution thisSize, fillPrice, mHandle, False

placePeggedOrders
        
If Not mStopOrder Is Nothing Then
    If mStopOrder.quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        adjustStop mStopOrder.triggerPrice, Abs(size)
    End If
End If

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        adjustTarget mTargetOrder.limitPrice, Abs(size)
    End If
End If

RaiseEvent EntryOrderPartiallyFilled

handleStimulus StateTransitionStimuli.StimEntryOrderFill

Exit Sub

Err:
gHandleFatalError pProcName:="mEntryOrder_Fill", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mEntryOrder_PropertyChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexEntryOrderChanged

Exit Sub

Err:
gHandleFatalError pProcName:="mEntryOrder_PropertyChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mEntryOrder_StatusChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexEntryOrderChanged
Select Case mEntryOrder.status
    Case OrderStatuses.OrderStatusFilled
        mCancelPrice = 0
        
        gTB.notifyListeners _
                            "Entry order completed (" & _
                            mEntryOrder.id & "): " & _
                            gOrderActionToString(mEntryOrder.action) & " " & mEntryOrder.quantityFilled & " at " & _
                            mEntryOrder.averagePrice, _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        
        RaiseEvent EntryOrderFilled
        
        mEntryOrderFinished = True
        checkAllOrdersComplete
    
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusSubmitted
        
        If mEntryOrder.quantityFilled > 0 Then
            placePeggedOrders
        End If
    Case OrderStatuses.OrderStatusCancelled
        mEntryOrderFinished = True
        gTB.notifyListeners _
                            "Entry order cancelled (" & _
                            mEntryOrder.id & ") ", _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        handleStimulus StateTransitionStimuli.StimEntryOrderCancelled
        checkAllOrdersComplete
    Case OrderStatuses.OrderStatusRejected
        mEntryOrderFinished = True
        gTB.notifyListeners _
                            "Entry order rejected (" & _
                            mEntryOrder.id & ") ", _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        handleStimulus StateTransitionStimuli.StimEntryOrderCancelled
        checkAllOrdersComplete
End Select

Exit Sub

Err:
gHandleFatalError pProcName:="mEntryOrder_StatusChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
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
gHandleFatalError pProcName:="mStateTimeoutTimer_TimerExpired", pModuleName:=ModuleName, pFailpoint:=failpoint
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

Private Sub mStopOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

Dim failpoint As Long
On Error GoTo Err

calcProfitLoss fillPrice

fireChange OrderPlexStopOrderChanged

thisSize = IIf(mStopOrder.action = ActionBuy, quantityFilled, -quantityFilled)
setSize size + thisSize
gTB.notifyListeners _
                        "Stop order fill (" & _
                        mStopOrder.id & "): " & _
                        gOrderActionToString(mStopOrder.action) & " " & quantityFilled & " at " & _
                        fillPrice, _
                        IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                        Me

mPositionManager.NotifyExecution thisSize, fillPrice, mHandle, False

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        adjustTarget mTargetOrder.limitPrice, Abs(size)
    End If
End If

RaiseEvent StopOrderPartiallyFilled

Exit Sub

Err:
gHandleFatalError pProcName:="mStopOrder_Fill", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mStopOrder_PropertyChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexStopOrderChanged

Exit Sub

Err:
gHandleFatalError pProcName:="mStopOrder_PropertyChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mStopOrder_StatusChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexStopOrderChanged
Select Case mStopOrder.status
    Case OrderStatuses.OrderStatusFilled
        gTB.notifyListeners _
                            "Stop order completed (" & _
                            mStopOrder.id & "): " & _
                            gOrderActionToString(mStopOrder.action) & " " & mStopOrder.quantityFilled & " at " & _
                            mStopOrder.averagePrice, _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        
        RaiseEvent StopOrderFilled
        
        mStopOrderFinished = True
        checkAllOrdersComplete
    
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusSubmitted
    Case OrderStatuses.OrderStatusCancelled
        mStopOrderFinished = True
        gTB.notifyListeners _
                            "Stop order cancelled (" & _
                            mStopOrder.id & ")", _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        mStateEngine.setConditions Conditions.CondStopOrderCancelled
        handleStimulus StateTransitionStimuli.StimStopOrderCancelled
        checkAllOrdersComplete
    Case OrderStatuses.OrderStatusRejected
        mStopOrderFinished = True
        gTB.notifyListeners _
                            "Stop order rejected (" & _
                            mStopOrder.id & ")", _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        mStateEngine.setConditions Conditions.CondStopOrderCancelled
        handleStimulus StateTransitionStimuli.StimStopOrderCancelled
        checkAllOrdersComplete
End Select

Exit Sub

Err:
gHandleFatalError pProcName:="mStopOrder_StatusChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
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
gHandleFatalError pProcName:="mTargetOrder_Dirty", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mTargetOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

Dim failpoint As Long
On Error GoTo Err

calcProfitLoss fillPrice

fireChange OrderPlexTargetOrderChanged
thisSize = IIf(mTargetOrder.action = ActionBuy, quantityFilled, -quantityFilled)
setSize size + thisSize
gTB.notifyListeners _
                        "Target order fill (" & _
                        mTargetOrder.id & "): " & _
                        gOrderActionToString(mTargetOrder.action) & " " & quantityFilled & " at " & _
                        fillPrice, _
                        IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                        Me

mPositionManager.NotifyExecution thisSize, fillPrice, mHandle, False

If Not mStopOrder Is Nothing Then
    If mStopOrder.quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        adjustStop mStopOrder.triggerPrice, Abs(size)
    End If
End If

RaiseEvent TargetOrderPartiallyFilled

Exit Sub

Err:
gHandleFatalError pProcName:="mTargetOrder_Fill", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mTargetOrder_PropertyChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexTargetOrderChanged

Exit Sub

Err:
gHandleFatalError pProcName:="mTargetOrder_PropertyChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mTargetOrder_StatusChanged()
Dim failpoint As Long
On Error GoTo Err

fireChange OrderPlexTargetOrderChanged
Select Case mTargetOrder.status
    Case OrderStatuses.OrderStatusFilled
        gTB.notifyListeners _
                            "Target order completed (" & _
                            mTargetOrder.id & "): " & _
                            gOrderActionToString(mTargetOrder.action) & " " & mTargetOrder.quantityFilled & " at " & _
                            mTargetOrder.averagePrice, _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        
        RaiseEvent TargetOrderFilled
        
        mTargetOrderFinished = True
        checkAllOrdersComplete
    
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusSubmitted
    Case OrderStatuses.OrderStatusCancelled
        mTargetOrderFinished = True
        gTB.notifyListeners _
                            "Target order cancelled (" & _
                            mTargetOrder.id & ")", _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        mStateEngine.setConditions Conditions.CondTargetOrderCancelled
        handleStimulus StateTransitionStimuli.StimTargetOrderCancelled
        checkAllOrdersComplete
    Case OrderStatuses.OrderStatusRejected
        mTargetOrderFinished = True
        gTB.notifyListeners _
                            "Target order rejected (" & _
                            mTargetOrder.id & ")", _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        mStateEngine.setConditions Conditions.CondTargetOrderCancelled
        handleStimulus StateTransitionStimuli.StimTargetOrderCancelled
        checkAllOrdersComplete
    
End Select

Exit Sub

Err:
gHandleFatalError pProcName:="mTargetOrder_StatusChanged", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mTimerEntry Event Handlers
'@================================================================================

Private Sub mTimerEntry_Expired()
Dim failpoint As Long
On Error GoTo Err

gTB.notifyListeners _
                    "Cancel time expired (" & mTicker.TradePrice & ")", _
                    IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                    Me
Cancel
fireChange OrderPlexSelfCancelled
RaiseEvent SelfCancelled
cancelAfter = 0

Exit Sub

Err:
gHandleFatalError pProcName:="mTimerEntry_Expired", pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let cancelAfter(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

If value <> 0 Then
    Set mTimerEntry = mTicker.AddTimer(value, expiryTimeUnits.ExpiryTimeUnitSeconds, Nothing)
ElseIf Not mTimerEntry Is Nothing Then
    mTicker.RemoveTimer mTimerEntry
    Set mTimerEntry = Nothing
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="cancelAfter", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get closedOut() As Boolean
Dim failpoint As Long
On Error GoTo Err

closedOut = (mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="closedOut", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get closeoutOrder() As order
Set closeoutOrder = mCloseoutOrder
End Property

Public Property Get Contract() As Contract
Dim failpoint As Long
On Error GoTo Err

Set Contract = mTicker.Contract

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Contract", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
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

Public Property Get entryOrder() As order
Set entryOrder = mEntryOrder
End Property

Public Property Let entryReason(ByVal value As String)
Dim failpoint As Long
On Error GoTo Err

mEntryReason = value
If mHandle > 0 Then
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonEntry, mEntryReason
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="entryReason", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get entryReason() As String
entryReason = mEntryReason
End Property

Friend Property Let exitStrategy( _
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
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Get longPosition() As Boolean
Dim failpoint As Long
On Error GoTo Err

longPosition = (mEntryOrder.action = ActionBuy)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="longPosition", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let lossNotificationThreshold(ByVal value As Currency)
mLossNotificationThreshold = value
End Property

Public Property Let newEntryPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newEntryPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                "OrderPlex::(let)newEntryPrice", _
                                                "Can't change entry price - order " & entryOrder.id & " alReady filled"

Select Case mEntryOrder.orderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeTrail, _
        OrderTypes.OrderTypeAutoStop, _
        OrderTypes.OrderTypeAutoLimit, _
        OrderTypes.OrderTypeStop
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "OrderPlex::(let)newEntryPrice", _
                    "Can't change entry price for order type " & gOrderTypeToString(mEntryOrder.orderType)
End Select

If value <> mEntryOrder.limitPrice Then
    mNewEntryPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newEntryPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newEntryTriggerPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newEntryTriggerPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                "OrderPlex::(let)newEntryTriggerPrice", _
                                                "Can't change entry trigger price - order " & entryOrder.id & " alReady filled"

Select Case mEntryOrder.orderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeLimit, _
        OrderTypes.OrderTypeTrail, _
        OrderTypes.OrderTypeAutoLimit
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "OrderPlex::(let)newEntryTriggerPrice", _
                    "Can't change trigger price for order type " & gOrderTypeToString(mEntryOrder.orderType)
End Select

If value <> mEntryOrder.triggerPrice Then
    mNewEntryTriggerPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newEntryTriggerPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newOrderPrice( _
                ByVal orderIndex As Long, _
                ByVal value As Double)
Dim lOrder As order

Dim failpoint As Long
On Error GoTo Err

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

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newOrderPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newOrderTriggerPrice( _
                ByVal orderIndex As Long, _
                ByVal value As Double)
Dim lOrder As order

Dim failpoint As Long
On Error GoTo Err

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

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newOrderTriggerPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newQuantity(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newQuantity", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                "OrderPlex::(let)newQuantity", _
                                                "Can't change quantity - order " & entryOrder.id & " alReady filled"

If value <> mEntryOrder.quantity Then
    mNewQuantity = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newQuantity", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newStopOffset(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newStopOffset", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)newStopOffset", _
                                        "No stop order exists"

If Not mStopOrder.orderType = OrderTypes.OrderTypeAutoStop Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "OrderPlex::(let)newStopOffset", _
                    "Can't change offset for order type " & gOrderTypeToString(mEntryOrder.orderType)
End If

If mStopOrder.status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)newStopOffset", _
                                        "Can't change stop offset - order " & mStopOrder.id & " alReady placed"

If value <> mStopOrder.offset Then
    mNewStopOffset = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newStopOffset", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newStopPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newStopPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)newStopPrice", _
                                        "No stop order exists"

Select Case mStopOrder.orderType
Case OrderTypes.OrderTypeStopLimit
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newStopPrice", _
                "Can't change price for order type " & gOrderTypeToString(mStopOrder.orderType)
End Select

If value <> mStopOrder.limitPrice Then
    mNewStopPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newStopPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newStopQuantity(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newStopQuantity", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)newStopQuantity", _
                                        "No stop order exists"

If value <> mStopOrder.quantity Then
    mNewStopQuantity = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newStopQuantity", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newStopTriggerPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newStopTriggerPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)newStopTriggerPrice", _
                                        "No stop order exists"

Select Case mStopOrder.orderType
    Case OrderTypes.OrderTypeStop, _
        OrderTypes.OrderTypeStopLimit
    Case Else
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "OrderPlex::(let)newStopTriggerPrice", _
                    "Can't change trigger price for order type " & gOrderTypeToString(mStopOrder.orderType)
End Select

If value <> mStopOrder.triggerPrice Then
    mNewStopTriggerPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newStopTriggerPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newTargetPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newTargetPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)newTargetPrice", _
                                        "No target order exists"

Select Case mTargetOrder.orderType
Case OrderTypes.OrderTypeLimitIfTouched, _
    OrderTypes.OrderTypeLimit
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newTargetPrice", _
                "Can't change price for order type " & gOrderTypeToString(mTargetOrder.orderType)
End Select

If value <> mTargetOrder.limitPrice Then
    mNewTargetPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newTargetPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newTargetOffset(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newTargetOffset", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)newTargetOffset", _
                                        "No target order exists"

If Not mTargetOrder.orderType = OrderTypes.OrderTypeAutoLimit Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                    "OrderPlex::(let)newTargetOffset", _
                    "Can't change offset for order type " & gOrderTypeToString(mTargetOrder.orderType)
End If

If mTargetOrder.status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)newTargetOffset", _
                                        "Can't change target offset - order " & mTargetOrder.id & " alReady placed"

If value <> mTargetOrder.offset Then
    mNewTargetOffset = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newTargetOffset", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newTargetQuantity(ByVal value As Long)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newTargetQuantity", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)newTargetQuantity", _
                                        "No target order exists"

If value <> mTargetOrder.quantity Then
    mNewTargetQuantity = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newTargetQuantity", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Let newTargetTriggerPrice(ByVal value As Double)
Dim failpoint As Long
On Error GoTo Err

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newTargetTriggerPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                        "OrderPlex::(let)newTargetTriggerPrice", _
                                        "No target order exists"

Select Case mTargetOrder.orderType
Case OrderTypes.OrderTypeLimitIfTouched, _
    OrderTypes.OrderTypeMarketIfTouched
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::(let)newTargetTriggerPrice", _
                "Can't change trigger price for order type " & gOrderTypeToString(mTargetOrder.orderType)
End Select

If value <> mTargetOrder.triggerPrice Then
    mNewTargetTriggerPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="newTargetTriggerPrice", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get pendingSize() As Long
Dim failpoint As Long
On Error GoTo Err

If mEntryOrder Is Nothing Then pendingSize = 0: Exit Property
If mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed Then pendingSize = 0: Exit Property
pendingSize = IIf(longPosition, mEntryOrder.quantityRemaining, -mEntryOrder.quantityRemaining)

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="pendingSize", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get order(ByVal index As Long) As order
Dim failpoint As Long
On Error GoTo Err

Set order = mOrders(index)

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

Public Property Get numberOfOrders() As Long
Dim failpoint As Long
On Error GoTo Err

numberOfOrders = mOrders.Count

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
If mPreventUnprotectedPosition Then
    mStateEngine.setConditions Conditions.CondProtected
Else
    mStateEngine.clearConditions Conditions.CondProtected
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

Public Property Let profitNotificationThreshold(ByVal value As Currency)
mProfitNotificationThreshold = value
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

If mStopOrder.status = OrderStatuses.OrderStatusFilled Then Exit Property

' the amount at Risk has two parts - that due to fills alReady obtained, and
' that due to fills pending.

filledTotalPrice = mEntryOrder.quantityFilled * mEntryOrder.averagePrice
Select Case entryOrder.orderType
    Case OrderTypes.OrderTypeMarket
        pendingTotalPrice = mEntryOrder.quantityRemaining * _
                IIf(longPosition, mTicker.AskPrice, mTicker.BidPrice)
    Case OrderTypes.OrderTypeLimit, _
            OrderTypes.OrderTypeStopLimit, _
            OrderTypeLimitIfTouched, _
            OrderTypeMarketIfTouched
        pendingTotalPrice = mEntryOrder.quantityRemaining * mEntryOrder.limitPrice
    Case OrderTypes.OrderTypeStop
        pendingTotalPrice = mEntryOrder.quantityRemaining * mEntryOrder.triggerPrice
    
    
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
stopTotalPrice = mStopOrder.triggerPrice * _
                        (mEntryOrder.quantity + mEntryOrder.quantityFilled)
If longPosition Then
    Risk = mTickSize * mTickValue * (filledTotalPrice + pendingTotalPrice - stopTotalPrice) / mTickSize
Else
    Risk = mTickSize * mTickValue * (stopTotalPrice - filledTotalPrice - pendingTotalPrice) / mTickSize
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="Risk", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Property

Friend Property Let scalingStrategy( _
                ByVal value As PositionManagementStrategy)
Set mScalingStrategy = value
End Property

Public Property Get size() As Long
size = mSize
End Property

Public Property Get State() As OrderPlexStateCodes
Dim failpoint As Long
On Error GoTo Err

State = mStateEngine.State

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="State", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get stoporder() As order
Set stoporder = mStopOrder
End Property

Public Property Let stopReason(ByVal value As String)
Dim failpoint As Long
On Error GoTo Err

mStopReason = value
If mHandle > 0 Then
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonStop, mStopReason
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="stopReason", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
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
Dim failpoint As Long
On Error GoTo Err

mTargetReason = value
If mHandle > 0 Then
    mPositionManager.SetPlexReason mHandle, ReasonTypes.ReasonTarget, mTargetReason
End If

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="targetReason", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get targetReason() As String
targetReason = mTargetReason
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

Public Sub adjustStop(ByVal stopPrice As Double, _
                    Optional ByVal quantity As Long = -1)
Dim failpoint As Long
On Error GoTo Err

gTB.notifyListeners _
                            "Adjust stop: (" & mTicker.TradePrice & ") " & _
                            " stop " & stopPrice & _
                            ", quantity " & quantity, _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me

newStopTriggerPrice = stopPrice
If quantity <> -1 Then newStopQuantity = quantity
update

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="adjustStop", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Public Sub adjustTarget(ByVal targetPrice As Double, _
                    Optional ByVal quantity As Long = -1)
Dim failpoint As Long
On Error GoTo Err

gTB.notifyListeners _
                            "Adjust target: (" & mTicker.TradePrice & ") " & _
                            " target " & targetPrice & _
                            ", quantity " & quantity, _
                            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me

newTargetPrice = targetPrice
If quantity <> -1 Then newTargetQuantity = quantity
update

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="adjustTarget", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Public Function Cancel(Optional ByVal evenIfFilled As Boolean = False) As Boolean

Dim failpoint As Long
On Error GoTo Err

If evenIfFilled Then
    mStateEngine.clearConditions Conditions.CondNoFillCancellation
    Cancel = handleStimulus(StateTransitionStimuli.StimCancelEvenIfFill)
Else
    mStateEngine.setConditions Conditions.CondNoFillCancellation
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
Public Sub cancelChanges()
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
Friend Sub checkpoint()
Dim failpoint As Long
On Error GoTo Err

Set mEntryOrderCopy = mEntryOrder.clone
mEntryOrder.setClean

If Not mStopOrder Is Nothing Then
    Set mStopOrderCopy = mStopOrder.clone
    mStopOrder.setClean
End If

If Not mTargetOrder Is Nothing Then
    Set mTargetOrderCopy = mTargetOrder.clone
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

gTB.notifyListeners _
                    "Closing out: (" & mTicker.TradePrice & ") " & _
                    "size " & size & _
                    ", pending size " & pendingSize, _
                    IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                    Me

handleStimulus StateTransitionStimuli.StimCloseout

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="CloseOut", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Friend Function createEntryOrder() As order
Dim failpoint As Long
On Error GoTo Err

If Not mEntryOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::createEntryOrder", _
                "Entry order alReady exists"

Set mEntryOrder = mOrderContext.newOrder

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

Friend Function createStopOrder() As order
Dim failpoint As Long
On Error GoTo Err

If Not mStopOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::createStopOrder", _
                "Stop order alReady exists"

Set mStopOrder = mOrderContext.newOrder

Set createStopOrder = mStopOrder
mStateEngine.setConditions Conditions.CondStopOrderExists
mStateEngine.clearConditions Conditions.CondStopOrderCancelled

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

Friend Function createTargetOrder() As order
Dim failpoint As Long
On Error GoTo Err

If Not mTargetOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::createTargetOrder", _
                "Target order alReady exists"

Set mTargetOrder = mOrderContext.newOrder

Set createTargetOrder = mTargetOrder
mStateEngine.setConditions Conditions.CondTargetOrderExists
mStateEngine.clearConditions Conditions.CondTargetOrderCancelled

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

Friend Sub execute(Optional ByVal cancelPrice As Double = 0)
Dim failpoint As Long
On Error GoTo Err

If mPreventUnprotectedPosition And mStopOrder Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "execute", _
            "No stop loss order exists"
End If

mCancelPrice = cancelPrice
fireChange OrderPlexChangeTypes.OrderPlexCreated
handleStimulus StateTransitionStimuli.StimExecute

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="execute", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub Finish()
Dim failpoint As Long
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

Public Sub update()
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

If State <> OrderPlexStateCreated Then
    mOrderSubmitter.modifyOrderPlex Me, updateEntry, updateStop, updateTarget, False
End If

checkpoint

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

Profit = size * (value - mEntryOrder.averagePrice) * mValueFactor

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
    End If
    
    If mLossNotificationThreshold <> 0 And _
        mProfit <= mLossNotificationThreshold _
    Then
        fireChange OrderPlexLossThresholdExceeded
        RaiseEvent LossThresholdExceeded
        mLossNotificationThreshold = 0
    End If

    If mDrawdownNotificationThreshold <> 0 And _
        (mMaxProfit - Profit >= mDrawdownNotificationThreshold) _
    Then
        fireChange OrderPlexDrawdownThresholdExceeded
        RaiseEvent DrawdownThresholdExceeded
        mDrawdownNotificationThreshold = 0
    End If
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="calcProfitLoss", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function cancelOrders() As Boolean
Dim stopOrderID As String
Dim targetOrderID As String

Dim failpoint As Long
On Error GoTo Err

cancelAfter = 0
mCancelPrice = 0

Select Case mEntryOrder.status
Case OrderStatuses.OrderStatusCreated
    mEntryOrder.updateStatus OrderStatuses.OrderStatusCancelled

    cancelOrders = True
End Select

If Not mStopOrder Is Nothing Then
    Select Case mStopOrder.status
    Case OrderStatuses.OrderStatusCreated
        mStopOrder.updateStatus OrderStatuses.OrderStatusCancelled
        cancelOrders = True
    End Select
End If

If Not mTargetOrder Is Nothing Then
    Select Case mTargetOrder.status
    Case OrderStatuses.OrderStatusCreated
        mTargetOrder.updateStatus OrderStatuses.OrderStatusCancelled
        cancelOrders = True
    End Select
End If

If mOrderSubmitter.cancelOrderPlex(Me) Then
    cancelOrders = True
    
    If Not mStopOrder Is Nothing Then stopOrderID = mStopOrder.id
    If Not mTargetOrder Is Nothing Then targetOrderID = mTargetOrder.id
    gTB.notifyListeners _
                    "Cancel orders: (" & mTicker.TradePrice & ") " & _
                    entryOrder.id & _
                    IIf(stopOrderID <> "", ", " & stopOrderID, ", ") & _
                    IIf(targetOrderID <> "", ", " & targetOrderID, ", "), _
                    IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                    Me
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
Select Case mStopOrder.status
Case OrderStatuses.OrderStatusCreated
    mStopOrder.updateStatus OrderStatuses.OrderStatusCancelled
Case OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    gTB.notifyListeners _
                "Cancel order: (" & mTicker.TradePrice & ") " & _
                mStopOrder.id, _
                IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                Me
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
Select Case mTargetOrder.status
Case OrderStatuses.OrderStatusCreated
    mTargetOrder.updateStatus OrderStatuses.OrderStatusCancelled
Case OrderStatuses.OrderStatusFilled, _
    OrderStatuses.OrderStatusCancelling, _
    OrderStatuses.OrderStatusCancelled
Case Else
    gTB.notifyListeners _
                "Cancel order: (" & mTicker.TradePrice & ") " & _
                mTargetOrder.id, _
                IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
                Me
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
    If mEntryOrder.isDirty Then updateEntry = True
    If mNewQuantity >= 0 And _
        mNewQuantity <> mEntryOrder.quantity _
    Then
        If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change size - order " & entryOrder.id & " alReady filled"
        mEntryOrder.quantity = mNewQuantity
        updateEntry = True
    End If
    mNewQuantity = -1
    
    If mNewEntryPrice >= 0 And _
        mNewEntryPrice <> mEntryOrder.limitPrice _
    Then
        If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change price - order " & mEntryOrder.id & " alReady filled"
        mEntryOrder.limitPrice = mNewEntryPrice
        updateEntry = True
    End If
    mNewEntryPrice = -1
    
    If mNewEntryTriggerPrice >= 0 And _
        mNewEntryTriggerPrice <> mEntryOrder.triggerPrice _
    Then
        If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.id & " alReady filled"
        mEntryOrder.triggerPrice = mNewEntryTriggerPrice
        updateEntry = True
    End If
    mNewEntryTriggerPrice = -1
End If

If Not mStopOrder Is Nothing Then
    If mStopOrder.isDirty Then updateStop = True
    If mNewStopPrice >= 0 And _
        mNewStopPrice <> mStopOrder.limitPrice _
    Then
        If mStopOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mStopOrder.id & " alReady filled"
        If mStopOrder.status = OrderStatuses.OrderStatusSubmitted Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change price - order " & mStopOrder.id & " alReady submitted"
        mStopOrder.limitPrice = mNewStopPrice
        updateStop = True
    End If
    mNewStopPrice = -1
    
    If mNewStopOffset <> DummyOffset And _
        mNewStopOffset <> mStopOrder.offset _
    Then
        If mStopOrder.status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            "OrderPlex::update", _
                                                            "Can't change stop offset - order " & mStopOrder.id & " alReady placed"
        mStopOrder.offset = mNewStopOffset
    End If
    mNewStopOffset = DummyOffset

    If mNewStopQuantity >= 0 And _
        mNewStopQuantity <> mStopOrder.quantity _
    Then
        If mStopOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            "OrderPlex::update", _
                                                            "Can't change stop quantity - order " & mStopOrder.id & " alReady placed"
        mStopOrder.quantity = mNewStopQuantity
    End If
    mNewStopQuantity = -1

    If mNewStopTriggerPrice >= 0 And _
        mNewStopTriggerPrice <> mStopOrder.triggerPrice _
    Then
        If mStopOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.id & " alReady filled"
        mStopOrder.triggerPrice = mNewStopTriggerPrice
        updateStop = True
    End If
    mNewStopTriggerPrice = -1
End If

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.isDirty Then updateTarget = True
    If mNewTargetPrice >= 0 And _
        mNewTargetPrice <> mTargetOrder.limitPrice _
    Then
        If mTargetOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change price - order " & mTargetOrder.id & " alReady filled"
        mTargetOrder.limitPrice = mNewTargetPrice
        updateTarget = True
    End If
    mNewTargetPrice = -1
    
    If mNewTargetOffset <> DummyOffset And _
        mNewTargetOffset <> mTargetOrder.offset _
    Then
        If mTargetOrder.status <> OrderStatuses.OrderStatusCreated Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            "OrderPlex::update", _
                                                            "Can't change target offset - order " & targetOrder.id & " alReady placed"
        mTargetOrder.offset = mNewTargetOffset
    End If
    mNewTargetOffset = DummyOffset

    If mNewTargetQuantity >= 0 And _
        mNewTargetQuantity <> mTargetOrder.quantity _
    Then
        If mTargetOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                            "OrderPlex::update", _
                                                            "Can't change target quantity - order " & mTargetOrder.id & " alReady placed"
        mTargetOrder.quantity = mNewTargetQuantity
    End If
    mNewTargetQuantity = -1

    If mNewTargetTriggerPrice >= 0 And _
        mNewTargetTriggerPrice <> mTargetOrder.triggerPrice _
    Then
        If mTargetOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.id & " alReady filled"
        mTargetOrder.triggerPrice = mNewTargetTriggerPrice
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
Dim action As Variant

Dim failpoint As Long
On Error GoTo Err

For Each action In pActions
    Select Case CLng(action)
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
                    "Invalid action " & action
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
handleStimulus = doActions(mStateEngine.notifyStimulus(stimulus))
If mStateEngine.State <> currState Then
    fireChange OrderPlexChangeTypes.OrderPlexStateChanged
    fireStateChange mStateEngine.State
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="handleStimulus", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function placeCloseoutOrder()
' place an order to closeout the filled part
Dim failpoint As Long
On Error GoTo Err

Set mCloseoutOrder = mOrderContext.newOrder
mCloseoutOrder.action = IIf(longPosition, OrderActions.ActionSell, OrderActions.ActionBuy)
mCloseoutOrder.quantity = Abs(size)
mCloseoutOrder.orderType = OrderTypes.OrderTypeMarket
mCloseoutOrder.timeInForce = OrderTifs.TIFGoodTillCancelled

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

mCreationTime = mTicker.Timestamp

If mEntryOrder Is Nothing Then mEntryOrderFinished = True
If mStopOrder Is Nothing Then mStopOrderFinished = True
If mTargetOrder Is Nothing Then mTargetOrderFinished = True

mOrderSubmitter.executeOrderPlex Me
mTicker.AddQuoteListener Me

checkpoint

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="placeOrders", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub placePeggedOrders()
Dim failpoint As Long
On Error GoTo Err

If Not mStopOrder Is Nothing Then
    If mStopOrder.status = OrderStatuses.OrderStatusCreated Then
        If mStopOrder.orderType = OrderTypes.OrderTypeAutoStop Then
            ' the stop price is to be pegged to the entry price - the
            ' order's offset property will alReady contain any required
            ' offset from that price
            mStopOrder.orderType = OrderTypes.OrderTypeStop
            newStopTriggerPrice = (mStopOrder.offset * mTickSize) + entryOrder.lastFillPrice
            newStopQuantity = mEntryOrder.quantityFilled
        End If
    End If
End If
If Not mTargetOrder Is Nothing Then
    If mTargetOrder.status = OrderStatuses.OrderStatusCreated Then
        If mTargetOrder.orderType = OrderTypes.OrderTypeAutoLimit Then
            ' the limit price is to be pegged to the entry price - the
            ' order's offset property will alReady contain any required
            ' offset from that price
            mTargetOrder.orderType = OrderTypes.OrderTypeLimit
            newTargetPrice = (mTargetOrder.offset * mTickSize) + entryOrder.lastFillPrice
            newTargetQuantity = mEntryOrder.quantityFilled
        End If
    End If
End If
update

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="placePeggedOrders", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Function resubmitStopOrder()
Dim failpoint As Long
On Error GoTo Err

mOrderSubmitter.resubmitStopOrder Me
gTB.notifyListeners _
            "Resubmit stop order: (" & mTicker.TradePrice & ") id " & mStopOrder.id & _
            "new broker id " & mStopOrder.brokerId, _
            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
            Me
mStateEngine.clearConditions Conditions.CondStopOrderCancelled
mStopOrderFinished = False

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="resubmitStopOrder", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function resubmitStopAndTargetOrders()
Dim failpoint As Long
On Error GoTo Err

mOrderSubmitter.resubmitStopAndTargetOrders Me
gTB.notifyListeners _
            "Resubmit stop order: (" & mTicker.TradePrice & ") id " & mStopOrder.id & _
            " new broker id " & mStopOrder.brokerId, _
            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
            Me
gTB.notifyListeners _
            "Resubmit target order: (" & mTicker.TradePrice & ") id " & mTargetOrder.id & _
            " new broker id " & mTargetOrder.brokerId, _
            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
            Me
mStateEngine.clearConditions Conditions.CondStopOrderCancelled
mStopOrderFinished = False
mStateEngine.clearConditions Conditions.CondTargetOrderCancelled
mTargetOrderFinished = False

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="resubmitStopAndTargetOrders", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function

Private Function resubmitTargetOrder()
Dim failpoint As Long
On Error GoTo Err

mOrderSubmitter.resubmitTargetOrder Me
gTB.notifyListeners _
            "Resubmit target order: (" & mTicker.TradePrice & ") id " & mTargetOrder.id & _
            " new broker id " & mTargetOrder.brokerId, _
            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
            Me
mStateEngine.clearConditions Conditions.CondTargetOrderCancelled
mTargetOrderFinished = False

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="resubmitTargetOrder", pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

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
If mSize <> 0 Then
    mStateEngine.setConditions Conditions.CondSizeNonZero
Else
    mStateEngine.clearConditions Conditions.CondSizeNonZero
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

