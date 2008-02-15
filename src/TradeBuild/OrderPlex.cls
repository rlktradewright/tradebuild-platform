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

'@================================================================================
' Enums
'@================================================================================

Private Enum ErrCodes
    ErrNotImplemented = vbObjectError + 513
    errCloseoutIncomplete
    errPriceChangeNotAllowedForOrderType
    errTriggerPriceChangeNotAllowedForOrderType
    errOffsetChangeNotAllowedForOrderType
    errCantChangeFilledOrder
    errNoStopOrder
    errNoTargetOrder
    errCantChangeUnplacedOrder
    errCantChangeSubmittedOrder
    errInvalidActionReturnedFromStateEngine
    errActionNotAllowedInThisState
    errWidenStopNotAllowed
    errCantChangeStopForOrderType
    errCantChangeTargetForOrderType
End Enum

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

Private mDummy As Boolean
    ' if set, this is a dummy orderplex, which means that it will be handled by
    ' the simulator like any other order, but its size will be considered 0. It
    ' will not be submitted as a real order

Private mSimulated As Boolean

Private mValueFactor As Double
Private mProfit As Currency
Private mMaxProfit As Currency
Private mDrawdown As Currency
Private mProfitNotificationThreshold As Currency
Private mLossNotificationThreshold As Currency
Private mDrawdownNotificationThreshold As Currency

Private mChangeListeners As Collection

Private mCreationTime As Date

Private mProfitListeners As Collection

Private mIndexApplication As Long

Private mDirty As Boolean

Private mOrders As Collection

Private mOrderSubmitter As TradeBuildSP.IOrderSubmitter

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
IOrderPlex_numberOfOrders = mOrders.Count
End Property

Private Property Get IOrderPlex_Order(ByVal index As Long) As TradeBuildSP.IOrder
Set IOrderPlex_Order = order(index)
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

Private Sub QuoteListener_ask(ev As QuoteEvent)
If Not longPosition Then calcProfitLoss ev.price
End Sub

Private Sub QuoteListener_bid(ev As QuoteEvent)
If longPosition Then calcProfitLoss ev.price
End Sub

Private Sub QuoteListener_High(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_PreviousClose(ev As QuoteEvent)

End Sub

Private Sub QuoteListener_trade(ev As QuoteEvent)
Dim cancelPriceHit As Boolean

If mCancelPrice = 0 Then Exit Sub

If longPosition Then
    Select Case mEntryOrder.orderType
    Case OrderTypes.OrderTypeLimit
        If ev.price >= mCancelPrice Then cancelPriceHit = True
    Case OrderTypes.OrderTypeStop, _
        OrderTypes.OrderTypeStopLimit
        If ev.price <= mCancelPrice Then cancelPriceHit = True
    End Select
Else
    Select Case mEntryOrder.orderType
    Case OrderTypes.OrderTypeLimit
        If ev.price <= mCancelPrice Then cancelPriceHit = True
    Case OrderTypes.OrderTypeStop, _
        OrderTypes.OrderTypeStopLimit
        If ev.price >= mCancelPrice Then cancelPriceHit = True
    End Select
End If
If cancelPriceHit Then
    gTB.notifyListeners _
                        "Cancel price hit (" & mTicker.TradePrice & ")", _
                        IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                        Me
    Cancel
    fireChange OrderPlexSelfCancelled
    RaiseEvent SelfCancelled
End If
End Sub

Private Sub QuoteListener_volume(ev As QuoteEvent)

End Sub

'@================================================================================
' mCloseoutOrder Event Handlers
'@================================================================================

Private Sub mCloseoutOrder_StatusChanged()
fireChange OrderPlexCloseoutOrderChanged
Select Case mCloseoutOrder.status
    Case OrderStatuses.OrderStatusFilled
        gTB.notifyListeners _
                            "Closeout order completed (" & _
                            mCloseoutOrder.id & "): " & _
                            gOrderActionToString(mCloseoutOrder.action) & " " & mCloseoutOrder.quantityFilled & " at " & _
                            mCloseoutOrder.averagePrice, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
                            
        checkAllOrdersComplete
        
        If size <> 0 Then Err.Raise errCloseoutIncomplete, "OrderPlex::mCloseoutOrder_StatusChanged", "Closeout incomplete: size remaining is " & size
        
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPreSubmitted
    Case OrderStatuses.OrderStatusSubmitted
    Case OrderStatuses.OrderStatusCancelled
        gTB.notifyListeners _
                            "Closeout order cancelled (" & _
                            mCloseoutOrder.id & ")", _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
End Select
End Sub

Private Sub mCloseoutOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

calcProfitLoss fillPrice

fireChange OrderPlexCloseoutOrderChanged
thisSize = IIf(mCloseoutOrder.action = ActionBuy, quantityFilled, -quantityFilled)
setSize size + thisSize
gTB.notifyListeners _
                    "Closeout order fill (" & _
                    mCloseoutOrder.id & "): " & _
                    gOrderActionToString(mCloseoutOrder.action) & " " & quantityFilled & " at " & _
                    fillPrice, _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                    Me

mPositionManager.notifyExecution thisSize, fillPrice, mHandle, True

RaiseEvent CloseoutOrderPartiallyFilled
End Sub

'@================================================================================
' mEntryOrder Event Handlers
'@================================================================================

Private Sub mEntryOrder_Dirty()
setDirty
End Sub

Private Sub mEntryOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long
Dim risk As Currency

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
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                    Me

If mHandle < 0 And Not mStopOrder Is Nothing Then
    If longPosition Then
        If mStopOrder.orderType = OrderTypes.OrderTypeAutoStop Then
            risk = (mEntryOrder.quantity + mEntryOrder.quantityFilled) * mStopOrder.offset * mTickValue
        Else
            risk = (mEntryOrder.quantity + mEntryOrder.quantityFilled) * (fillPrice - mStopOrder.triggerPrice) * mValueFactor
        End If
    Else
        If mStopOrder.orderType = OrderTypes.OrderTypeAutoStop Then
            risk = -(mEntryOrder.quantity + mEntryOrder.quantityFilled) * mStopOrder.offset * mTickValue
        Else
            risk = (mEntryOrder.quantity + mEntryOrder.quantityFilled) * (mStopOrder.triggerPrice - fillPrice) * mValueFactor
        End If
    End If
End If

If mHandle < 0 Then
    mHandle = mPositionManager.registerPlex(mKey, mDescription, mEntryOrder.action, risk)
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonEntry, mEntryReason
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonTarget, mTargetReason
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonStop, mStopReason
End If
mPositionManager.notifyExecution thisSize, fillPrice, mHandle, False

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

handleStimulus STIM_ENTRY_ORDER_FILL
End Sub

Private Sub mEntryOrder_StatusChanged()
fireChange OrderPlexEntryOrderChanged
Select Case mEntryOrder.status
    Case OrderStatuses.OrderStatusFilled
        mCancelPrice = 0
        
        gTB.notifyListeners _
                            "Entry order completed (" & _
                            mEntryOrder.id & "): " & _
                            gOrderActionToString(mEntryOrder.action) & " " & mEntryOrder.quantityFilled & " at " & _
                            mEntryOrder.averagePrice, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
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
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        handleStimulus STIM_ENTRY_ORDER_CANCELLED
        checkAllOrdersComplete
End Select
End Sub

'@================================================================================
' mStopOrder Event Handlers
'@================================================================================

Private Sub mStopOrder_Dirty()
setDirty
End Sub

Private Sub mStopOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

calcProfitLoss fillPrice

fireChange OrderPlexStopOrderChanged

thisSize = IIf(mStopOrder.action = ActionBuy, quantityFilled, -quantityFilled)
setSize size + thisSize
gTB.notifyListeners _
                        "Stop order fill (" & _
                        mStopOrder.id & "): " & _
                        gOrderActionToString(mStopOrder.action) & " " & quantityFilled & " at " & _
                        fillPrice, _
                        IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                        Me

mPositionManager.notifyExecution thisSize, fillPrice, mHandle, False

If Not mTargetOrder Is Nothing Then
    If mTargetOrder.quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        adjustTarget mTargetOrder.limitPrice, Abs(size)
    End If
End If

RaiseEvent StopOrderPartiallyFilled
End Sub

Private Sub mStopOrder_StatusChanged()
fireChange OrderPlexStopOrderChanged
Select Case mStopOrder.status
    Case OrderStatuses.OrderStatusFilled
        gTB.notifyListeners _
                            "Stop order completed (" & _
                            mStopOrder.id & "): " & _
                            gOrderActionToString(mStopOrder.action) & " " & mStopOrder.quantityFilled & " at " & _
                            mStopOrder.averagePrice, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
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
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        mStateEngine.setConditions COND_STOP_ORDER_CANCELLED
        handleStimulus STIM_STOP_ORDER_CANCELLED
        checkAllOrdersComplete
End Select
End Sub

'@================================================================================
' mTargetOrder Event Handlers
'@================================================================================

Private Sub mTargetOrder_Dirty()
setDirty
End Sub

Private Sub mTargetOrder_Fill(ByVal quantityFilled As Long, _
                                        ByVal fillPrice As Double)
Dim thisSize As Long

calcProfitLoss fillPrice

fireChange OrderPlexTargetOrderChanged
thisSize = IIf(mTargetOrder.action = ActionBuy, quantityFilled, -quantityFilled)
setSize size + thisSize
gTB.notifyListeners _
                        "Target order fill (" & _
                        mTargetOrder.id & "): " & _
                        gOrderActionToString(mTargetOrder.action) & " " & quantityFilled & " at " & _
                        fillPrice, _
                        IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                        Me

mPositionManager.notifyExecution thisSize, fillPrice, mHandle, False

If Not mStopOrder Is Nothing Then
    If mStopOrder.quantity <> Abs(size) And _
        Abs(size) <> 0 _
    Then
        adjustStop mStopOrder.triggerPrice, Abs(size)
    End If
End If

RaiseEvent TargetOrderPartiallyFilled
End Sub

Private Sub mTargetOrder_StatusChanged()
fireChange OrderPlexTargetOrderChanged
Select Case mTargetOrder.status
    Case OrderStatuses.OrderStatusFilled
        gTB.notifyListeners _
                            "Target order completed (" & _
                            mTargetOrder.id & "): " & _
                            gOrderActionToString(mTargetOrder.action) & " " & mTargetOrder.quantityFilled & " at " & _
                            mTargetOrder.averagePrice, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        
        RaiseEvent TargetOrderFilled
        
        mTargetOrderFinished = True
        checkAllOrdersComplete
    
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusCancelling
    Case OrderStatuses.OrderStatusPendingSubmit
    Case OrderStatuses.OrderStatusSubmitted
    Case OrderStatuses.OrderStatusCancelled
        mTargetOrderFinished = True
        gTB.notifyListeners _
                            "Target order cancelled (" & _
                            mTargetOrder.id & ")", _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me
        mStateEngine.setConditions COND_TARGET_ORDER_CANCELLED
        handleStimulus STIM_TARGET_ORDER_CANCELLED
        checkAllOrdersComplete
    
End Select
End Sub

'@================================================================================
' mTimerEntry Event Handlers
'@================================================================================

Private Sub mTimerEntry_Expired()
gTB.notifyListeners _
                    "Cancel time expired (" & mTicker.TradePrice & ")", _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                    Me
Cancel
fireChange OrderPlexSelfCancelled
RaiseEvent SelfCancelled
cancelAfter = 0
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let cancelAfter(ByVal value As Long)
If value <> 0 Then
    Set mTimerEntry = mTicker.addTimer(value, expiryTimeUnits.ExpiryTimeUnitSeconds, Nothing)
ElseIf Not mTimerEntry Is Nothing Then
    mTicker.removeTimer mTimerEntry
    Set mTimerEntry = Nothing
End If
End Property

Public Property Get closedOut() As Boolean
closedOut = (mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed)
End Property

Public Property Get closeoutOrder() As order
Set closeoutOrder = mCloseoutOrder
End Property

Public Property Get Contract() As Contract
Set Contract = mTicker.Contract
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

Friend Property Let dummy(ByVal value As Boolean)
mDummy = value
End Property

Public Property Get dummy() As Boolean
dummy = mDummy
End Property

Public Property Get entryOrder() As order
Set entryOrder = mEntryOrder
End Property

Public Property Let entryReason(ByVal value As String)
mEntryReason = value
If mHandle > 0 Then
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonEntry, mEntryReason
End If
End Property

Public Property Get entryReason() As String
entryReason = mEntryReason
End Property

Friend Property Let exitStrategy( _
                ByVal value As PositionManagementStrategy)
Set mExitStrategy = value
End Property

Public Property Get indexApplication() As Long
indexApplication = mIndexApplication
End Property

Friend Property Let key(ByVal value As String)
mKey = value
End Property

Public Property Get key() As String
key = mKey
End Property

Public Property Get longPosition() As Boolean
longPosition = (mEntryOrder.action = ActionBuy)
End Property

Public Property Let lossNotificationThreshold(ByVal value As Currency)
mLossNotificationThreshold = value
End Property

Public Property Let newEntryPrice(ByVal value As Double)
Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise errActionNotAllowedInThisState, _
                "OrderPlex::(let)newEntryPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeFilledOrder, _
                                                "OrderPlex::(let)newEntryPrice", _
                                                "Can't change entry price - order " & entryOrder.id & " already filled"

Select Case mEntryOrder.orderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeTrail, _
        OrderTypes.OrderTypeAutoStop, _
        OrderTypes.OrderTypeAutoLimit, _
        OrderTypes.OrderTypeStop
        Err.Raise errPriceChangeNotAllowedForOrderType, _
                    "OrderPlex::(let)newEntryPrice", _
                    "Can't change entry price for order type " & gOrderTypeToString(mEntryOrder.orderType)
End Select

If value <> mEntryOrder.limitPrice Then
    mNewEntryPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If
End Property

Public Property Let newEntryTriggerPrice(ByVal value As Double)
Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise errActionNotAllowedInThisState, _
                "OrderPlex::(let)newEntryTriggerPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeFilledOrder, _
                                                "OrderPlex::(let)newEntryTriggerPrice", _
                                                "Can't change entry trigger price - order " & entryOrder.id & " already filled"

Select Case mEntryOrder.orderType
    Case OrderTypes.OrderTypeMarket, _
        OrderTypes.OrderTypeLimit, _
        OrderTypes.OrderTypeTrail, _
        OrderTypes.OrderTypeAutoLimit
        Err.Raise errTriggerPriceChangeNotAllowedForOrderType, _
                    "OrderPlex::(let)newEntryTriggerPrice", _
                    "Can't change trigger price for order type " & gOrderTypeToString(mEntryOrder.orderType)
End Select

If value <> mEntryOrder.triggerPrice Then
    mNewEntryTriggerPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If
End Property

Public Property Let newOrderPrice( _
                ByVal orderIndex As Long, _
                ByVal value As Double)
Dim lOrder As order

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
End Property

Public Property Let newOrderTriggerPrice( _
                ByVal orderIndex As Long, _
                ByVal value As Double)
Dim lOrder As order

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
End Property

Public Property Let newQuantity(ByVal value As Long)
Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise errActionNotAllowedInThisState, _
                "OrderPlex::(let)newQuantity", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeFilledOrder, _
                                                "OrderPlex::(let)newQuantity", _
                                                "Can't change quantity - order " & entryOrder.id & " already filled"

If value <> mEntryOrder.quantity Then
    mNewQuantity = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If
End Property

Public Property Let newStopOffset(ByVal value As Long)
Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise errActionNotAllowedInThisState, _
                "OrderPlex::(let)newStopOffset", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise errNoStopOrder, _
                                        "OrderPlex::(let)newStopOffset", _
                                        "No stop order exists"

If Not mStopOrder.orderType = OrderTypes.OrderTypeAutoStop Then
        Err.Raise errOffsetChangeNotAllowedForOrderType, _
                    "OrderPlex::(let)newStopOffset", _
                    "Can't change offset for order type " & gOrderTypeToString(mEntryOrder.orderType)
End If

If mStopOrder.status <> OrderStatuses.OrderStatusCreated Then Err.Raise errCantChangeUnplacedOrder, _
                                        "OrderPlex::(let)newStopOffset", _
                                        "Can't change stop offset - order " & mStopOrder.id & " already placed"

If value <> mStopOrder.offset Then
    mNewStopOffset = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If
End Property

Public Property Let newStopPrice(ByVal value As Double)
Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise errActionNotAllowedInThisState, _
                "OrderPlex::(let)newStopPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise errNoStopOrder, _
                                        "OrderPlex::(let)newStopPrice", _
                                        "No stop order exists"

Select Case mStopOrder.orderType
Case OrderTypes.OrderTypeStopLimit
Case Else
    Err.Raise errTriggerPriceChangeNotAllowedForOrderType, _
                "OrderPlex::(let)newStopPrice", _
                "Can't change trigger price for order type " & gOrderTypeToString(mEntryOrder.orderType)
End Select

If value <> mStopOrder.limitPrice Then
    mNewStopPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If
End Property

Public Property Let newStopQuantity(ByVal value As Long)
Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise errActionNotAllowedInThisState, _
                "OrderPlex::(let)newStopQuantity", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise errNoStopOrder, _
                                        "OrderPlex::(let)newStopQuantity", _
                                        "No stop order exists"

If value <> mStopOrder.quantity Then
    mNewStopQuantity = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If
End Property

Public Property Let newStopTriggerPrice(ByVal value As Double)
Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise errActionNotAllowedInThisState, _
                "OrderPlex::(let)newStopTriggerPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise errNoStopOrder, _
                                        "OrderPlex::(let)newStopTriggerPrice", _
                                        "No stop order exists"

Select Case mStopOrder.orderType
    Case OrderTypes.OrderTypeStop, _
        OrderTypes.OrderTypeStopLimit
    Case Else
        Err.Raise errTriggerPriceChangeNotAllowedForOrderType, _
                    "OrderPlex::(let)newStopTriggerPrice", _
                    "Can't change trigger price for order type " & gOrderTypeToString(mEntryOrder.orderType)
End Select

If value <> mStopOrder.triggerPrice Then
    mNewStopTriggerPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If
End Property

Public Property Let newTargetPrice(ByVal value As Double)
Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case Else
    Err.Raise errActionNotAllowedInThisState, _
                "OrderPlex::(let)newTargetPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise errNoTargetOrder, _
                                        "OrderPlex::(let)newTargetPrice", _
                                        "No target order exists"

Select Case mTargetOrder.orderType
Case OrderTypes.OrderTypeLimitIfTouched, _
    OrderTypes.OrderTypeLimit
Case Else
    Err.Raise errTriggerPriceChangeNotAllowedForOrderType, _
                "OrderPlex::(let)newTargetPrice", _
                "Can't change price for order type " & gOrderTypeToString(mEntryOrder.orderType)
End Select

If value <> mTargetOrder.limitPrice Then
    mNewTargetPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If
End Property

Public Property Let newTargetOffset(ByVal value As Long)
Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise errActionNotAllowedInThisState, _
                "OrderPlex::(let)newTargetOffset", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise errNoTargetOrder, _
                                        "OrderPlex::(let)newTargetOffset", _
                                        "No target order exists"

If Not mTargetOrder.orderType = OrderTypes.OrderTypeAutoLimit Then
        Err.Raise errOffsetChangeNotAllowedForOrderType, _
                    "OrderPlex::(let)newTargetOffset", _
                    "Can't change offset for order type " & gOrderTypeToString(mEntryOrder.orderType)
End If

If mTargetOrder.status <> OrderStatuses.OrderStatusCreated Then Err.Raise errCantChangeUnplacedOrder, _
                                        "OrderPlex::(let)newTargetOffset", _
                                        "Can't change target offset - order " & targetOrder.id & " already placed"

If value <> mTargetOrder.offset Then
    mNewTargetOffset = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If
End Property

Public Property Let newTargetQuantity(ByVal value As Long)
Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise errActionNotAllowedInThisState, _
                "OrderPlex::(let)newTargetQuantity", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mStopOrder Is Nothing Then Err.Raise errNoTargetOrder, _
                                        "OrderPlex::(let)newTargetQuantity", _
                                        "No target order exists"

If value <> mTargetOrder.quantity Then
    mNewTargetQuantity = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If
End Property

Public Property Let newTargetTriggerPrice(ByVal value As Double)
Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Property
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case Else
    Err.Raise errActionNotAllowedInThisState, _
                "OrderPlex::(let)newTargetTriggerPrice", _
                "Action not allowed in state " & mStateEngine.State
End Select

If mTargetOrder Is Nothing Then Err.Raise errNoStopOrder, _
                                        "OrderPlex::(let)newTargetTriggerPrice", _
                                        "No target order exists"

Select Case mTargetOrder.orderType
Case OrderTypes.OrderTypeLimitIfTouched, _
    OrderTypes.OrderTypeMarketIfTouched
Case Else
    Err.Raise errTriggerPriceChangeNotAllowedForOrderType, _
                "OrderPlex::(let)newTargetTriggerPrice", _
                "Can't change trigger price for order type " & gOrderTypeToString(mEntryOrder.orderType)
End Select

If value <> mTargetOrder.triggerPrice Then
    mNewTargetTriggerPrice = value
    setDirty
    fireChange OrderPlexChangeTypes.OrderPlexChangesPending
    RaiseEvent Dirty
End If
End Property

Public Property Get pendingSize() As Long
If mEntryOrder Is Nothing Then pendingSize = 0: Exit Property
If mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed Then pendingSize = 0: Exit Property
pendingSize = IIf(longPosition, mEntryOrder.quantity, -mEntryOrder.quantity)
End Property

Public Property Get order(ByVal index As Long) As order
Set order = mOrders(index)
End Property

Friend Property Let OrderContext(ByVal value As OrderContext)
Set mOrderContext = value
End Property

Friend Property Let orderSubmitter( _
                ByVal value As TradeBuildSP.IOrderSubmitter)
Set mOrderSubmitter = value
End Property

Public Property Get numberOfOrders() As Long
numberOfOrders = mOrders.Count
End Property

Friend Property Let PositionManager(ByVal value As PositionManager)
Set mPositionManager = value
End Property

Public Property Let profitNotificationThreshold(ByVal value As Currency)
mProfitNotificationThreshold = value
End Property

Public Property Get risk() As Currency
Dim stopTotalPrice As Double
Dim pendingTotalPrice As Double
Dim filledTotalPrice As Double

If mStopOrder Is Nothing Then Exit Property

If mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed Then Exit Property

If mStopOrder.status = OrderStatuses.OrderStatusFilled Then Exit Property

' the amount at risk has two parts - that due to fills already obtained, and
' that due to fills pending.

filledTotalPrice = mEntryOrder.quantityFilled * mEntryOrder.averagePrice
Select Case entryOrder.orderType
    Case OrderTypes.OrderTypeMarket
        pendingTotalPrice = mEntryOrder.quantity * _
                IIf(longPosition, mTicker.AskPrice, mTicker.BidPrice)
    Case OrderTypes.OrderTypeLimit, OrderTypes.OrderTypeStopLimit
        pendingTotalPrice = mEntryOrder.quantity * mEntryOrder.limitPrice
    Case OrderTypes.OrderTypeStop
        pendingTotalPrice = mEntryOrder.quantity * mEntryOrder.triggerPrice
End Select
stopTotalPrice = mStopOrder.triggerPrice * _
                        (mEntryOrder.quantity + mEntryOrder.quantityFilled)
If longPosition Then
    risk = mTickSize * mTickValue * (filledTotalPrice + pendingTotalPrice - stopTotalPrice) / mTickSize
Else
    risk = mTickSize * mTickValue * (stopTotalPrice - filledTotalPrice - pendingTotalPrice) / mTickSize
End If

End Property

Friend Property Let scalingStrategy( _
                ByVal value As PositionManagementStrategy)
Set mScalingStrategy = value
End Property

Friend Property Let simulated(ByVal value As Boolean)
mSimulated = value
End Property

Public Property Get simulated() As Boolean
simulated = mSimulated
End Property

Public Property Get size() As Long
size = mSize
End Property

Public Property Get State() As OrderPlexStateCodes
State = mStateEngine.State
End Property

Public Property Get stoporder() As order
Set stoporder = mStopOrder
End Property

Public Property Let stopReason(ByVal value As String)
mStopReason = value
If mHandle > 0 Then
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonStop, mStopReason
End If
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
mTargetReason = value
If mHandle > 0 Then
    mPositionManager.setPlexReason mHandle, ReasonTypes.ReasonTarget, mTargetReason
End If
End Property

Public Property Get targetReason() As String
targetReason = mTargetReason
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Dim lContract As Contract
Set mTicker = value
Set lContract = mTicker.Contract
mTickSize = lContract.TickSize
mTickValue = lContract.tickValue
mValueFactor = lContract.multiplier
mcurrencyCode = lContract.specifier.currencyCode
End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub addChangeListener(ByVal value As ChangeListener)
mChangeListeners.add value

' notify this new listenener that we are already created
Dim ev As ChangeEvent
Set ev.source = Me
ev.changeType = OrderPlexChangeTypes.OrderPlexCreated
value.Change ev
End Sub

Public Sub addProfitListener(ByVal value As ProfitListener)
mProfitListeners.add value
End Sub

Public Sub adjustStop(ByVal stopPrice As Double, _
                    Optional ByVal quantity As Long = -1)
gTB.notifyListeners _
                            "Adjust stop: (" & mTicker.TradePrice & ") " & _
                            " stop " & stopPrice & _
                            ", quantity " & quantity, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me

newStopTriggerPrice = stopPrice
If quantity <> -1 Then newStopQuantity = quantity
update

End Sub

Public Sub adjustTarget(ByVal targetPrice As Double, _
                    Optional ByVal quantity As Long = -1)
gTB.notifyListeners _
                            "Adjust target: (" & mTicker.TradePrice & ") " & _
                            " target " & targetPrice & _
                            ", quantity " & quantity, _
                            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                            Me

newTargetPrice = targetPrice
If quantity <> -1 Then newTargetQuantity = quantity
update

End Sub

Public Function Cancel(Optional ByVal evenIfFilled As Boolean = False) As Boolean

If evenIfFilled Then
    mStateEngine.clearConditions COND_NO_FILL_CANCELLATION
    Cancel = doActions(mStateEngine.notifyStimulus(STIM_CANCEL_EVEN_IF_FILL))
Else
    mStateEngine.setConditions COND_NO_FILL_CANCELLATION
    Cancel = doActions(mStateEngine.notifyStimulus(STIM_CANCEL_IF_NO_FILL))
End If
End Function

''
' Cancels any changes made to this <code>orderPlex</code> that have not yet
' been actioned.
'@/
Public Sub cancelChanges()
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
End Sub

''
' Copies the current state of the member orders for use by cancelChanges
'@/
Friend Sub checkpoint()
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
End Sub

Public Sub CloseOut(ByVal method As String)

If UCase(method) <> "MKT" Then
    Err.Raise ErrNotImplemented, "OrderPlex::closeout", "method " & method & " not implemented"
End If

gTB.notifyListeners _
                    "Closing out: (" & mTicker.TradePrice & ") " & _
                    "size " & size & _
                    ", pending size " & pendingSize, _
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                    Me

handleStimulus STIM_CLOSEOUT

End Sub

Friend Function createEntryOrder() As order
If Not mEntryOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::createEntryOrder", _
                "Entry order already exists"

Set mEntryOrder = mOrderContext.newOrder

Set createEntryOrder = mEntryOrder

If mOrders.Count >= 1 Then
    mOrders.add mEntryOrder, , 1
Else
    mOrders.add mEntryOrder
End If
End Function

Friend Function createStopOrder() As order
If Not mStopOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::createStopOrder", _
                "Stop order already exists"

Set mStopOrder = mOrderContext.newOrder

Set createStopOrder = mStopOrder
mStateEngine.setConditions COND_STOP_ORDER_EXISTS
mStateEngine.clearConditions COND_STOP_ORDER_CANCELLED

If Not mEntryOrder Is Nothing Then
    mOrders.add mStopOrder, , , 1
ElseIf Not mTargetOrder Is Nothing Then
    mOrders.add mStopOrder, , 1
Else
    mOrders.add mStopOrder
End If
End Function

Friend Function createTargetOrder() As order
If Not mTargetOrder Is Nothing Then _
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "OrderPlex::createTargetOrder", _
                "Target order already exists"

Set mTargetOrder = mOrderContext.newOrder

Set createTargetOrder = mTargetOrder
mStateEngine.setConditions COND_TARGET_ORDER_EXISTS
mStateEngine.clearConditions COND_TARGET_ORDER_CANCELLED

If mEntryOrder Is Nothing And _
    mStopOrder Is Nothing _
Then
    mOrders.add mTargetOrder
ElseIf mEntryOrder Is Nothing Then
    mOrders.add mTargetOrder, , , 1
ElseIf mStopOrder Is Nothing Then
    mOrders.add mTargetOrder, , , 1
Else
    mOrders.add mTargetOrder, , , 2
End If
End Function

Friend Sub execute(Optional ByVal cancelPrice As Double = 0)
mCancelPrice = cancelPrice
fireChange OrderPlexChangeTypes.OrderPlexCreated
handleStimulus STIM_EXECUTE
End Sub

Friend Sub finish()
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
End Sub

Public Sub removeChangeListener(ByVal value As ChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.item(i) Is value Then mChangeListeners.remove i
Next
End Sub

Public Sub removeProfitListener(ByVal value As ProfitListener)
Dim i As Long
For i = mProfitListeners.Count To 1 Step -1
    If mProfitListeners.item(i) Is value Then mProfitListeners.remove i
Next
End Sub

Public Sub update()
Dim updateEntry As Boolean
Dim updateStop As Boolean
Dim updateTarget As Boolean

Select Case mStateEngine.State
Case OrderPlexStateCodes.OrderPlexStateCreated
Case OrderPlexStateCodes.OrderPlexStateSubmitted
Case OrderPlexStateCodes.OrderPlexStateCancelling
    Exit Sub
Case Else
    Err.Raise errActionNotAllowedInThisState, _
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
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcProfitLoss(ByVal value As Double)
Dim profit As Currency
Dim notifyProfit As Boolean
Dim notifyMaxProfit As Boolean
Dim notifyDrawdown As Boolean

profit = size * (value - mEntryOrder.averagePrice) * mValueFactor

If profit <> mProfit Then
    mProfit = profit
    notifyProfit = True
End If

If profit > mMaxProfit Then
    mMaxProfit = profit
    notifyMaxProfit = True
End If

If profit < mMaxProfit Then
    mDrawdown = mMaxProfit - profit
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
        (mMaxProfit - profit >= mDrawdownNotificationThreshold) _
    Then
        fireChange OrderPlexDrawdownThresholdExceeded
        RaiseEvent DrawdownThresholdExceeded
        mDrawdownNotificationThreshold = 0
    End If
End If
End Function

Private Function cancelOrders() As Boolean
Dim stopOrderID As String
Dim targetOrderID As String

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
                    IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                    Me
End If

End Function

Private Function cancelStopOrder()
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
                IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                Me
    mOrderSubmitter.cancelStopOrder Me
End Select
End Function

Private Function cancelTargetOrder()
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
                IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
                Me
    mOrderSubmitter.cancelTargetOrder Me
End Select
End Function

Private Sub checkAllOrdersComplete()
If mStateEngine.State = OrderPlexStateCodes.OrderPlexStateClosed Then Exit Sub
If mEntryOrderFinished And _
    mStopOrderFinished And _
    mTargetOrderFinished _
Then
    handleStimulus STIM_ALL_ORDERS_COMPLETE
End If

End Sub

Private Sub checkForPendingUpdates( _
                ByRef updateEntry As Boolean, _
                ByRef updateStop As Boolean, _
                ByRef updateTarget As Boolean)
If Not mEntryOrder Is Nothing Then
    If mEntryOrder.isDirty Then updateEntry = True
    If mNewQuantity >= 0 And _
        mNewQuantity <> mEntryOrder.quantity _
    Then
        If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeFilledOrder, _
                                                        "OrderPlex::update", _
                                                        "Can't change size - order " & entryOrder.id & " already filled"
        mEntryOrder.quantity = mNewQuantity
        updateEntry = True
    End If
    mNewQuantity = -1
    
    If mNewEntryPrice >= 0 And _
        mNewEntryPrice <> mEntryOrder.limitPrice _
    Then
        If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeFilledOrder, _
                                                        "OrderPlex::update", _
                                                        "Can't change price - order " & mEntryOrder.id & " already filled"
        mEntryOrder.limitPrice = mNewEntryPrice
        updateEntry = True
    End If
    mNewEntryPrice = -1
    
    If mNewEntryTriggerPrice >= 0 And _
        mNewEntryTriggerPrice <> mEntryOrder.triggerPrice _
    Then
        If mEntryOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeFilledOrder, _
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.id & " already filled"
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
        If mStopOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeFilledOrder, _
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mStopOrder.id & " already filled"
        If mStopOrder.status = OrderStatuses.OrderStatusSubmitted Then Err.Raise errCantChangeSubmittedOrder, _
                                                        "OrderPlex::update", _
                                                        "Can't change price - order " & mStopOrder.id & " already submitted"
        mStopOrder.limitPrice = mNewStopPrice
        updateStop = True
    End If
    mNewStopPrice = -1
    
    If mNewStopOffset <> DummyOffset And _
        mNewStopOffset <> mStopOrder.offset _
    Then
        If mStopOrder.status <> OrderStatuses.OrderStatusCreated Then Err.Raise errCantChangeUnplacedOrder, _
                                                            "OrderPlex::update", _
                                                            "Can't change stop offset - order " & mStopOrder.id & " already placed"
        mStopOrder.offset = mNewStopOffset
    End If
    mNewStopOffset = DummyOffset

    If mNewStopQuantity >= 0 And _
        mNewStopQuantity <> mStopOrder.quantity _
    Then
        If mStopOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeUnplacedOrder, _
                                                            "OrderPlex::update", _
                                                            "Can't change stop quantity - order " & mStopOrder.id & " already placed"
        mStopOrder.quantity = mNewStopQuantity
    End If
    mNewStopQuantity = -1

    If mNewStopTriggerPrice >= 0 And _
        mNewStopTriggerPrice <> mStopOrder.triggerPrice _
    Then
        If mStopOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeFilledOrder, _
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.id & " already filled"
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
        If mTargetOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeFilledOrder, _
                                                        "OrderPlex::update", _
                                                        "Can't change price - order " & mTargetOrder.id & " already filled"
        mTargetOrder.limitPrice = mNewTargetPrice
        updateTarget = True
    End If
    mNewTargetPrice = -1
    
    If mNewTargetOffset <> DummyOffset And _
        mNewTargetOffset <> mTargetOrder.offset _
    Then
        If mTargetOrder.status <> OrderStatuses.OrderStatusCreated Then Err.Raise errCantChangeUnplacedOrder, _
                                                            "OrderPlex::update", _
                                                            "Can't change target offset - order " & targetOrder.id & " already placed"
        mTargetOrder.offset = mNewTargetOffset
    End If
    mNewTargetOffset = DummyOffset

    If mNewTargetQuantity >= 0 And _
        mNewTargetQuantity <> mTargetOrder.quantity _
    Then
        If mTargetOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeUnplacedOrder, _
                                                            "OrderPlex::update", _
                                                            "Can't change target quantity - order " & mTargetOrder.id & " already placed"
        mTargetOrder.quantity = mNewTargetQuantity
    End If
    mNewTargetQuantity = -1

    If mNewTargetTriggerPrice >= 0 And _
        mNewTargetTriggerPrice <> mTargetOrder.triggerPrice _
    Then
        If mTargetOrder.status = OrderStatuses.OrderStatusFilled Then Err.Raise errCantChangeFilledOrder, _
                                                        "OrderPlex::update", _
                                                        "Can't change trigger price - order " & mEntryOrder.id & " already filled"
        mTargetOrder.triggerPrice = mNewTargetTriggerPrice
        updateTarget = True
    End If
    mNewTargetTriggerPrice = -1
End If
End Sub

Private Function completionActions()
mOrderContext.orderPlexClosed Me
fireChange OrderPlexCompleted
RaiseEvent Completed
Set mStopStrategy = Nothing
Set mScalingStrategy = Nothing
Set mExitStrategy = Nothing
If size = 0 Then mTicker.removeQuoteListener Me
End Function

Private Function doActions(ByRef Actions() As Long) As Variant
Dim action As Variant

For Each action In Actions
    Select Case CLng(action)
    Case SpecialActions.NoAction
    Case ACT_PLACE_ORDERS
        doActions = placeOrders
    Case ACT_CANCEL_ORDERS
        doActions = cancelOrders
    Case ACT_CANCEL_STOP_ORDER
        doActions = cancelStopOrder
    Case ACT_CANCEL_TARGET_ORDER
        doActions = cancelTargetOrder
    Case ACT_RESUBMIT_STOP_ORDER
        doActions = resubmitStopOrder
    Case ACT_RESUBMIT_TARGET_ORDER
        doActions = resubmitTargetOrder
    Case ACT_RESUBMIT_STOP_AND_TARGET_ORDERS
        doActions = resubmitStopAndTargetOrders
    Case ACT_PLACE_CLOSEOUT_ORDER
        doActions = placeCloseoutOrder
    Case ACT_COMPLETION_ACTIONS
        doActions = completionActions
    Case ACT_ALARM
        ' need to put something in here
    Case Else
        Err.Raise errInvalidActionReturnedFromStateEngine, _
                    "OrderPlex::doActions", _
                    "Invalid action " & action
    End Select
Next
End Function

Private Sub fireChange( _
                ByVal changeType As OrderPlexChangeTypes)
Dim listener As ChangeListener
Dim i As Long
Dim Change As ChangeEvent
Set Change.source = Me
Change.changeType = changeType
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change Change
Next
End Sub

Private Sub fireProfit( _
                ByVal profitAmount As Double, _
                ByVal profitType As ProfitTypes)
Dim listener As ProfitListener
Dim i As Long
Dim profit As ProfitEvent
Set profit.source = Me
profit.profitAmount = profitAmount
profit.profitType = profitType
profit.currencyCode = mcurrencyCode
For i = 1 To mProfitListeners.Count
    Set listener = mProfitListeners(i)
    listener.profitAmount profit
Next
End Sub

Private Sub handleStimulus(ByVal stimulus As StateTransitionStimuli)
Dim currState As OrderPlexStateCodes

currState = mStateEngine.State
doActions mStateEngine.notifyStimulus(stimulus)
If mStateEngine.State <> currState Then fireChange OrderPlexChangeTypes.OrderPlexStateChanged
End Sub

Private Function placeCloseoutOrder()
' place an order to closeout the filled part
Set mCloseoutOrder = mOrderContext.newOrder
mCloseoutOrder.action = IIf(longPosition, OrderActions.ActionSell, OrderActions.ActionBuy)
mCloseoutOrder.quantity = Abs(size)
mCloseoutOrder.orderType = OrderTypes.OrderTypeMarket
mCloseoutOrder.timeInForce = OrderTifs.TIFGoodTillCancelled

mOrderSubmitter.modifyOrderPlex Me, False, False, False, True

mOrders.add mCloseoutOrder

fireChange OrderPlexCloseoutOrderCreated
mOrderSubmitter.modifyOrderPlex Me, False, False, False, True
                                

End Function

Private Function placeOrders()

mCreationTime = mTicker.timestamp

If mEntryOrder Is Nothing Then mEntryOrderFinished = True
If mStopOrder Is Nothing Then mStopOrderFinished = True
If mTargetOrder Is Nothing Then mTargetOrderFinished = True

mOrderSubmitter.executeOrderPlex Me
mTicker.addQuoteListener Me

checkpoint
End Function

Private Sub placePeggedOrders()
If Not mStopOrder Is Nothing Then
    If mStopOrder.status = OrderStatuses.OrderStatusCreated Then
        If mStopOrder.orderType = OrderTypes.OrderTypeAutoStop Then
            ' the stop price is to be pegged to the entry price - the
            ' order's offset property will already contain any required
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
            ' order's offset property will already contain any required
            ' offset from that price
            mTargetOrder.orderType = OrderTypes.OrderTypeLimit
            newTargetPrice = (mTargetOrder.offset * mTickSize) + entryOrder.lastFillPrice
            newTargetQuantity = mEntryOrder.quantityFilled
        End If
    End If
End If
update
End Sub

Private Function resubmitStopOrder()
mOrderSubmitter.resubmitStopOrder Me
gTB.notifyListeners _
            "Resubmit stop order: (" & mTicker.TradePrice & ") id " & mStopOrder.id & _
            "new broker id " & mStopOrder.brokerId, _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
            Me
mStateEngine.clearConditions COND_STOP_ORDER_CANCELLED
mStopOrderFinished = False
End Function

Private Function resubmitStopAndTargetOrders()
mOrderSubmitter.resubmitStopAndTargetOrders Me
gTB.notifyListeners _
            "Resubmit stop order: (" & mTicker.TradePrice & ") id " & mStopOrder.id & _
            " new broker id " & mStopOrder.brokerId, _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
            Me
gTB.notifyListeners _
            "Resubmit target order: (" & mTicker.TradePrice & ") id " & mTargetOrder.id & _
            " new broker id " & mTargetOrder.brokerId, _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
            Me
mStateEngine.clearConditions COND_STOP_ORDER_CANCELLED
mStopOrderFinished = False
mStateEngine.clearConditions COND_TARGET_ORDER_CANCELLED
mTargetOrderFinished = False

End Function

Private Function resubmitTargetOrder()
mOrderSubmitter.resubmitTargetOrder Me
gTB.notifyListeners _
            "Resubmit target order: (" & mTicker.TradePrice & ") id " & mTargetOrder.id & _
            " new broker id " & mTargetOrder.brokerId, _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
            Me
mStateEngine.clearConditions COND_TARGET_ORDER_CANCELLED
mTargetOrderFinished = False
End Function

''
' Set the 'dirty' flag, and raise a Dirty event if one has not already
' been raised.
'@/
Private Sub setDirty()
If Not mDirty Then
    mDirty = True
    RaiseEvent Dirty
Else
    mDirty = True
End If
End Sub

Private Sub setSize(ByVal value As Long)
mSize = value
If mSize <> 0 Then
    mStateEngine.setConditions COND_SIZE_NON_ZERO
Else
    mStateEngine.clearConditions COND_SIZE_NON_ZERO
End If
fireChange OrderPlexChangeTypes.OrderPlexSizeChanged
End Sub



