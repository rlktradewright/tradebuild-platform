VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderContext"
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

Implements IStateChangeListener

'@================================================================================
' Events
'@================================================================================

Event Change(ev As ChangeEventData)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "OrderContext"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mOrderManager                           As OrderManager
Private mGroupName                              As String
Private mRecoveryContextName                    As String
Private mName                                   As String

Private mSessionName                            As String

Private mContractFuture                         As IFuture
Private mContract                               As IContract
Private mDataSource                             As IMarketDataSource
Attribute mDataSource.VB_VarHelpID = -1
Private mOrderSubmitter                         As IOrderSubmitter
Private mOrderAuthoriser                        As IOrderAuthoriser
Private mAccumulatedBracketOrders               As BracketOrders
Private mAccumulatedOrders                      As Orders


Private mIsSimulated                            As Boolean

Private mActiveBracketOrders                    As BracketOrders
    ' this contains all bracket orders for which one or more orders are
    ' not yet complete (or not yet cancelled)
    
Private mClosedBracketOrders                    As BracketOrders
    ' this contains all bracket orders for which all orders have completed (or
    ' been cancelled)
    
Private mCompletedSize                          As BoxedDecimal
    ' the total Size for all completed bracket orders (may be non-zero if one
    ' or more completed bracket orders has no target or stop-loss orders)
    
Private mPermittedOrderProperties               As IPermittedOrderProperties

Private mEnabled                                As Boolean
Private mProviderIsReady                        As Boolean
Private mDataSourceIsReady                      As Boolean

Private mTemplateOrder                          As Order

Private mPreventUnprotectedPositions            As Boolean

Private mChangeListeners                        As New Listeners

Private mTimerList                              As TimerList
Private mSimulatedClockFuture                   As IFuture
Private mClock                                  As Clock

Private WithEvents mFutureWaiter                As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mBracketOrderRecoveryController         As BracketOrderRecoveryCtlr

Private mOrderContextsRef                       As WeakReference

Private mTickSize                               As Double

Private mNotReadyMessage                        As String

Private mAccountDataProvider                    As IAccountDataProvider
Private mCurrencyConverter                      As ICurrencyConverter

Private mFixedAccountBalance                    As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mActiveBracketOrders = New BracketOrders
Set mClosedBracketOrders = New BracketOrders
Set mFutureWaiter = New FutureWaiter
Set mCompletedSize = DecimalZero
mTickSize = -1
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderContext terminated"
End Sub

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

If TypeOf ev.Source Is ICurrencyConverter Then fireChange OrderContextReadyStateChanged
    

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is IContract Then
    setContract ev.Future.Value
    setTimerList
    IsReadyTickSize
    fireChange OrderContextReadyStateChanged
ElseIf TypeOf ev.Future.Value Is Clock Then
    Set mClock = ev.Future.Value
    setTimerList
    fireChange OrderContextReadyStateChanged
End If

Exit Sub

Err:
GOrders.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get AccountDataProvider() As IAccountDataProvider
Set AccountDataProvider = mAccountDataProvider
End Property

Public Property Get ActiveBracketOrders() As BracketOrders
Set ActiveBracketOrders = mActiveBracketOrders
End Property

Public Property Get ActiveSize() As BoxedDecimal
Const ProcName As String = "ActiveSize"
On Error GoTo Err

Set ActiveSize = DecimalZero

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    Set ActiveSize = ActiveSize.Add(lBracketOrder.Size)
Next

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CancelPending() As Boolean
Const ProcName As String = "CancelPending"
On Error GoTo Err

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    If lBracketOrder.Cancel(pEvenIfFilled:=False) Then CancelPending = True
Next

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Clock() As Clock
Const ProcName As String = "Clock"
On Error GoTo Err

Assert Not mClock Is Nothing, "clock is not yet available"
Set Clock = mClock

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ClosedBracketOrders() As BracketOrders
Set ClosedBracketOrders = mClosedBracketOrders
End Property

Public Property Get CompletedSize() As BoxedDecimal
Set CompletedSize = mCompletedSize
End Property

Public Property Get CurrencyConverter() As ICurrencyConverter
Set CurrencyConverter = mCurrencyConverter
End Property

Public Property Get RecoveryContextName() As String
RecoveryContextName = mRecoveryContextName
End Property

Public Property Get ContractFuture() As IFuture
Set ContractFuture = mContractFuture
End Property

Public Property Get DataSource() As IMarketDataSource
Set DataSource = mDataSource
End Property

Public Property Let Enabled(ByVal Value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

If mEnabled = Value Then Exit Property

mEnabled = Value
fireChange OrderContextReadyStateChanged

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let FixedAccountBalance(ByVal Value As Double)
mFixedAccountBalance = Value
End Property

Friend Property Get FixedAccountBalance() As Double
FixedAccountBalance = mFixedAccountBalance
End Property

Friend Property Get GroupName() As String
GroupName = mGroupName
End Property

Friend Property Let IsAccountProviderReady(ByVal Value As Boolean)
Const ProcName As String = "isAccountProviderReady"
On Error GoTo Err

fireChange OrderContextReadyStateChanged

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsActive() As Boolean
IsActive = (mActiveBracketOrders.Count <> 0)
End Property

Friend Property Let IsDataSourceReady(ByVal Value As Boolean)
Const ProcName As String = "IsDataSourceReady"
On Error GoTo Err

If mDataSourceIsReady = Value Then Exit Property

mDataSourceIsReady = Value
fireChange OrderContextReadyStateChanged

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsDataSourceReady() As Boolean
IsDataSourceReady = mDataSourceIsReady
End Property

Friend Property Let IsProviderReady(ByVal Value As Boolean)
Const ProcName As String = "IsProviderReady"
On Error GoTo Err

If mProviderIsReady = Value Then Exit Property

mProviderIsReady = Value
If mProviderIsReady Then
    IsReadyTickSize
End If

fireChange OrderContextReadyStateChanged

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsProviderReady() As Boolean
IsProviderReady = mProviderIsReady
End Property

Public Property Get IsReady() As Boolean
If Not mProviderIsReady Then
    mNotReadyMessage = "Provider not ready to submit orders"
    Exit Function
End If

If Not mEnabled Then
    mNotReadyMessage = "Not ready - order submission disabled"
    Exit Function
End If

If mContract Is Nothing Then
    mNotReadyMessage = "Not ready - contract not yet resolved"
    Exit Function
End If

If mSimulatedClockFuture Is Nothing Then
ElseIf Not mSimulatedClockFuture.IsAvailable Then
    mNotReadyMessage = "Not ready - awaiting clock"
End If

mNotReadyMessage = ""
IsReady = True
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get NotReadyMessage() As String
NotReadyMessage = mNotReadyMessage
End Property

Friend Property Get OrderManager() As OrderManager
Set OrderManager = mOrderManager
End Property

Public Property Get OrderSubmitter() As IOrderSubmitter
Set OrderSubmitter = mOrderSubmitter
End Property

Public Property Get PendingBuySize() As BoxedDecimal
Const ProcName As String = "PendingBuySize"
On Error GoTo Err

Dim lSize As BoxedDecimal
Set lSize = DecimalZero

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    If lBracketOrder.LongPosition Then Set lSize = lSize.Add(lBracketOrder.PendingSize)
Next

Set PendingBuySize = lSize

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PendingSellSize() As BoxedDecimal
Const ProcName As String = "PendingSellSize"
On Error GoTo Err

Dim lSize As BoxedDecimal
Set lSize = DecimalZero

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    If Not lBracketOrder.LongPosition Then Set lSize = lSize.Add(lBracketOrder.PendingSize)
Next

Set PendingSellSize = lSize

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PendingSize() As BoxedDecimal
Const ProcName As String = "PendingSize"
On Error GoTo Err

Dim lSize As BoxedDecimal
Set lSize = DecimalZero

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    Set lSize = lSize.Add(lBracketOrder.PendingSize)
Next

Set PendingSize = lSize

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let PermittedOrderProperties(ByVal Value As IPermittedOrderProperties)
Set mPermittedOrderProperties = Value
End Property

Public Property Get PermittedOrderTifs() As Long
PermittedOrderTifs = mPermittedOrderProperties.PermittedOrderTifs
End Property

Public Property Get PermittedOrderTypes() As Long
PermittedOrderTypes = mPermittedOrderProperties.PermittedOrderTypes
End Property

Public Property Get PermittedStopTriggerMethods() As Long
PermittedStopTriggerMethods = mPermittedOrderProperties.PermittedStopTriggerMethods
End Property

Public Property Let PreventUnprotectedPositions( _
                ByVal Value As Boolean)
mPreventUnprotectedPositions = Value
End Property

Public Property Get PreventUnprotectedPositions() As Boolean
PreventUnprotectedPositions = mPreventUnprotectedPositions
End Property

Public Property Get Risk() As Currency
Const ProcName As String = "Risk"
On Error GoTo Err

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    If lBracketOrder.Risk = MaxCurrency Then
        Risk = MaxCurrency
        Exit For
    End If
    Risk = Risk + lBracketOrder.Risk
Next

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ScopeName() As String
ScopeName = mBracketOrderRecoveryController.ScopeName
End Property

Public Property Get Size() As BoxedDecimal
Const ProcName As String = "Size"
On Error GoTo Err

Set Size = mCompletedSize

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    Set Size = Size.Add(lBracketOrder.Size)
Next

Exit Property

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TemplateOrder() As IOrder
Set TemplateOrder = mTemplateOrder
End Property

Public Property Get TickSize() As Double
Assert mTickSize > 0#, "TickSize not yet available"
TickSize = mTickSize
End Property

Friend Property Get TimerList() As TimerList
Set TimerList = mTimerList
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub ActivateBracketOrder( _
                ByVal pBracketOrder As BracketOrder)
Const ProcName As String = "ActivateBracketOrder"
On Error GoTo Err

If pBracketOrder.IsComplete And pBracketOrder.RolloverDate = 0 Then
    NotifyBracketOrderClosed pBracketOrder
Else
    mActiveBracketOrders.Add pBracketOrder
    fireChange OrderContextActiveStateChanged
End If
mOrderManager.AddBracketOrder pBracketOrder
pBracketOrder.Activate

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "AddChangeListener"
On Error GoTo Err

mChangeListeners.Add pListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

''
' Creates a buy <code>BracketOrder</code>
'
' @return the <code>BracketOrder</code> object created
'
' @param quantity       quantity to buy
' @param EntryOrder     the order to enter the position
' @param StopLossOrder  the order to act as a stop-loss
' @param TargetOrder    the order to act as the profit target
' @param EntryTIF       specifies the time in force for the entry order
' @param StopTIF        specifies the time in force for the stop-loss order
' @param TargetTIF      specifies the time in force for the target order
' @param CancelPrice    if the entry order is a limit order and market Price rises to this
'                       Price before the entry order is filled, then the order is cancelled.
'                       If the entry order is a stop or stop limit order and market Price
'                       falls to this Price before the entry order is filled, then the order
'                       is cancelled
' @param CancelAfter    a time in seconds after which the bracket order is to be
'                       cancelled if tne entry order has not been filled
'
' @see
'
'@/
Public Function Buy( _
    ByVal Quantity As BoxedDecimal, _
    ByVal EntryOrder As IOrder, _
    Optional ByVal StopLossOrder As IOrder, _
    Optional ByVal TargetOrder As IOrder, _
    Optional ByVal EntryTIF As OrderTIFs = OrderTIFDay, _
    Optional ByVal StopTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
    Optional ByVal TargetTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
    Optional ByVal CancelPrice As Double, _
    Optional ByVal CancelAfter As Long) As IBracketOrder
Const ProcName As String = "Buy"
On Error GoTo Err

Set Buy = CreateBracketOrder(OrderActions. _
                            OrderActionBuy, _
                            Quantity, _
                            EntryOrder, _
                            StopLossOrder, _
                            TargetOrder)

EntryOrder.TimeInForce = EntryTIF
If Not StopLossOrder Is Nothing Then StopLossOrder.TimeInForce = StopTIF
If Not TargetOrder Is Nothing Then TargetOrder.TimeInForce = TargetTIF
Buy.CancelPrice = CancelPrice
Buy.CancelAfter = CancelAfter

ExecuteBracketOrder Buy, True

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CalculateOffsettedPrice( _
                ByVal pPriceSpec As PriceSpecifier, _
                ByVal pSecType As SecurityTypes, _
                ByVal pOrderAction As OrderActions) As Double
Const ProcName As String = "CalculateOffsettedPrice"
On Error GoTo Err

Dim lBidPrice As Double: lBidPrice = MaxDouble
If mDataSource.HasCurrentTick(TickTypeBid) Then lBidPrice = getCurrentPrice(TickTypeBid)

Dim lAskPrice As Double: lAskPrice = MaxDouble
If mDataSource.HasCurrentTick(TickTypeAsk) Then lAskPrice = getCurrentPrice(TickTypeAsk)

Dim lTradePrice As Double: lTradePrice = MaxDouble
If mDataSource.HasCurrentTick(TickTypeTrade) Then lTradePrice = getCurrentPrice(TickTypeTrade)

Dim lOptionModelPrice As Double: lOptionModelPrice = MaxDouble
If mDataSource.HasCurrentTick(TickTypeOptionModelPrice) Then lOptionModelPrice = getCurrentPrice(TickTypeOptionModelPrice)

If pPriceSpec.PriceType = PriceValueTypeEntry Or _
        pPriceSpec.PriceType = PriceValueTypeNone Then
    CalculateOffsettedPrice = GOrderUtils.MaxDoubleValue
    Exit Function
End If

If pPriceSpec.PriceType = PriceValueTypeValue Then
    pPriceSpec.CheckPriceValid mTickSize, pSecType
End If

Dim lRoundingMode As TickRoundingModes
If pOrderAction = OrderActionBuy Then
    lRoundingMode = TickRoundingModeUp
Else
    lRoundingMode = TickRoundingModeDown
End If

Dim lPrice As Double

Select Case pPriceSpec.PriceType
Case PriceValueTypeNone
    Assert False
Case PriceValueTypeValue
    lPrice = pPriceSpec.Price
Case PriceValueTypeAsk
    AssertArgument lAskPrice <> MaxDouble, "Ask price not available"
    lPrice = lAskPrice
Case PriceValueTypeBid
    AssertArgument lBidPrice <> MaxDouble, "Bid price not available"
    lPrice = lBidPrice
Case PriceValueTypeLast
    AssertArgument lTradePrice <> MaxDouble, "Trade price not available"
    lPrice = lTradePrice
Case PriceValueTypeModel
    AssertArgument lOptionModelPrice <> MaxDouble, "Model price not available"
    lPrice = lOptionModelPrice
Case PriceValueTypeEntry
    Assert False
Case PriceValueTypeMid
    AssertArgument lAskPrice <> MaxDouble, "Ask price not available"
    AssertArgument lBidPrice <> MaxDouble, "Bid price not available"
    lPrice = (lAskPrice + lBidPrice) / 2#
Case PriceValueTypeBidOrAsk
    If pOrderAction = OrderActionBuy Then
        AssertArgument lBidPrice <> MaxDouble, "Bid price not available"
        lPrice = lBidPrice
    Else
        AssertArgument lAskPrice <> MaxDouble, "Ask price not available"
        lPrice = lAskPrice
    End If
End Select

Dim lResult As Double

Select Case pPriceSpec.OffsetType
Dim lOffset As Double
Case PriceOffsetTypeNone
    lResult = lPrice
Case PriceOffsetTypeIncrement
    lOffset = pPriceSpec.Offset
Case PriceOffsetTypeNumberOfTicks
    lOffset = pPriceSpec.Offset * mTickSize
Case PriceOffsetTypeBidAskPercent
    AssertArgument lAskPrice <> MaxDouble, "Ask price not available"
    AssertArgument lBidPrice <> MaxDouble, "Bid price not available"
    lOffset = pPriceSpec.Offset * (lAskPrice - lBidPrice) / 100#
Case PriceOffsetTypePercent
    lOffset = pPriceSpec.Offset * lPrice / 100#
End Select

If Not pPriceSpec.UseCloseoutSemantics And _
    pPriceSpec.PriceType <> PriceValueTypeBidOrAsk _
Then
    lResult = lPrice + lOffset
ElseIf pOrderAction = OrderActionBuy Then
    lResult = lPrice + lOffset
Else
    lResult = lPrice - lOffset
End If

CalculateOffsettedPrice = gRoundToTickBoundary( _
                                        lResult, _
                                        mTickSize, _
                                        lRoundingMode)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function Cancel() As Boolean
Const ProcName As String = "Cancel"
On Error GoTo Err

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    If lBracketOrder.Cancel(pEvenIfFilled:=True) Then Cancel = True
Next

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CloseOut() As TaskController
Const ProcName As String = "CloseOut"
On Error GoTo Err

Dim t As New CloseoutOrderContextTask

Set CloseOut = StartTask(New CloseoutOrderContextTask, PriorityNormal, , , Me)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBuyBracketOrder( _
    ByVal Quantity As BoxedDecimal, _
    ByVal EntryOrder As IOrder, _
    Optional ByVal StopLossOrder As IOrder, _
    Optional ByVal TargetOrder As IOrder) As IBracketOrder
Const ProcName As String = "CreateBuyBracketOrder"
On Error GoTo Err

Set CreateBuyBracketOrder = CreateBracketOrder( _
                                        OrderActions.OrderActionBuy, _
                                        Quantity, _
                                        EntryOrder, _
                                        StopLossOrder, _
                                        TargetOrder)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateEntryOrder( _
                    ByVal pEntryType As OrderTypes, _
                    Optional ByVal pLimitPriceSpec As PriceSpecifier, _
                    Optional ByVal pTriggerPriceSpec As PriceSpecifier, _
                    Optional ByVal pOutsideRTH As Boolean = False) As IOrder
Const ProcName As String = "CreateEntryOrder"
On Error GoTo Err

AssertArgument IsEntryOrderType(pEntryType), "Invalid entry order type"

Dim lOrder As IOrder
Set lOrder = NewOrder(pEntryType)

lOrder.LimitPriceSpec = pLimitPriceSpec
lOrder.TriggerPriceSpec = pTriggerPriceSpec
lOrder.IgnoreRegularTradingHours = pOutsideRTH

Select Case pEntryType
Case OrderTypeMarket
Case OrderTypeMidprice
Case OrderTypeMarketOnOpen
Case OrderTypeMarketOnClose
Case OrderTypeMarketIfTouched
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
Case OrderTypeMarketToLimit
Case OrderTypeLimit
    AssertArgument pLimitPriceSpec.IsValid, "The limit price specification must be valid for this order type"
Case OrderTypeLimitOnOpen
    AssertArgument pLimitPriceSpec.IsValid, "The limit price specification must be valid for this order type"
Case OrderTypeLimitOnClose
    AssertArgument pLimitPriceSpec.IsValid, "The limit price specification must be valid for this order type"
Case OrderTypeLimitIfTouched
    AssertArgument pLimitPriceSpec.IsValid, "The limit price specification must be valid for this order type"
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
Case OrderTypeStop
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
Case OrderTypeStopLimit
    AssertArgument pLimitPriceSpec.IsValid, "The limit price specification must be valid for this order type"
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
Case OrderTypeTrail
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
Case OrderTypeTrailLimit
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
    AssertArgument pLimitPriceSpec.IsValid, "The limit price specification must be valid for this order type"
Case Else
    AssertArgument False, "Invalid entry order type"
End Select

Set CreateEntryOrder = lOrder

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStopLossOrder( _
                    ByVal pStopLossType As OrderTypes, _
                    Optional ByVal pLimitPriceSpec As PriceSpecifier, _
                    Optional ByVal pTriggerPriceSpec As PriceSpecifier, _
                    Optional ByVal pOutsideRTH As Boolean = False) As IOrder
Const ProcName As String = "CreateStopLossOrder"
On Error GoTo Err

AssertArgument IsStopLossOrderType(pStopLossType), "Invalid stop-loss order type"

Dim lOrder As IOrder
Set lOrder = NewOrder(pStopLossType)

lOrder.LimitPriceSpec = pLimitPriceSpec
lOrder.TriggerPriceSpec = pTriggerPriceSpec
lOrder.IgnoreRegularTradingHours = pOutsideRTH

Select Case pStopLossType
Case OrderTypeStop
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
Case OrderTypeStopLimit
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
    AssertArgument pLimitPriceSpec.IsValid, "The limit price specification must be valid for this order type"
Case OrderTypeTrail
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
Case OrderTypeTrailLimit
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
    AssertArgument pLimitPriceSpec.IsValid, "The limit price specification must be valid for this order type"
End Select

Set CreateStopLossOrder = lOrder

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTargetOrder( _
                    ByVal pTargetType As OrderTypes, _
                    Optional ByVal pLimitPriceSpec As PriceSpecifier, _
                    Optional ByVal pTriggerPriceSpec As PriceSpecifier, _
                    Optional ByVal pOutsideRTH As Boolean = False) As IOrder
Const ProcName As String = "CreateTargetOrder"
On Error GoTo Err

AssertArgument IsTargetOrderType(pTargetType), "Invalid target order type"

Dim lOrder As IOrder
Set lOrder = NewOrder(pTargetType)

lOrder.LimitPriceSpec = pLimitPriceSpec
lOrder.TriggerPriceSpec = pTriggerPriceSpec
lOrder.IgnoreRegularTradingHours = pOutsideRTH

Select Case pTargetType
Case OrderTypeLimit
    AssertArgument pLimitPriceSpec.IsValid, "The limit price specification must be valid for this order type"
Case OrderTypeLimitIfTouched
    AssertArgument pLimitPriceSpec.IsValid, "The limit price specification must be valid for this order type"
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
Case OrderTypeMarketIfTouched
    AssertArgument pTriggerPriceSpec.IsValid, "The trigger price specification must be valid for this order type"
End Select

Set CreateTargetOrder = lOrder

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBracketOrder( _
                ByVal pAction As OrderActions, _
                ByVal pQuantity As BoxedDecimal, _
                ByVal pEntryOrder As IOrder, _
                Optional ByVal pStopLossOrder As IOrder, _
                Optional ByVal pTargetOrder As IOrder) As IBracketOrder
Const ProcName As String = "CreateBracketOrder"
On Error GoTo Err



Set CreateBracketOrder = CreateBracketOrderWithId( _
                            pAction, _
                            pQuantity, _
                            pEntryOrder, _
                            GIdProvider.gNextId, _
                            pStopLossOrder, _
                            pTargetOrder)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBracketOrderWithId( _
                ByVal pAction As OrderActions, _
                ByVal pQuantity As BoxedDecimal, _
                ByVal pEntryOrder As IOrder, _
                ByVal pId As String, _
                Optional ByVal pStopLossOrder As IOrder, _
                Optional ByVal pTargetOrder As IOrder) As IBracketOrder
Const ProcName As String = "CreateBracketOrderWithId"
On Error GoTo Err

AssertArgument pQuantity > 0, "Quantity must be greater than zero"
AssertArgument Not pEntryOrder Is Nothing, "Entry order cannot be Nothing"
AssertArgument pId <> "", "Id must be supplied"

AssertArgument IsEntryOrderType(pEntryOrder.OrderType), "Invalid entry order type"
If Not pStopLossOrder Is Nothing Then AssertArgument IsStopLossOrderType(pStopLossOrder.OrderType), "Invalid stop-loss order type"
If Not pTargetOrder Is Nothing Then AssertArgument IsTargetOrderType(pTargetOrder.OrderType), "Invalid target order type"

Dim lBracketOrder As New BracketOrder
lBracketOrder.Initialise pId, mRecoveryContextName, mGroupName, Me, mContract, mOrderSubmitter, mDataSource, mIsSimulated
lBracketOrder.PreventUnprotectedPosition = mPreventUnprotectedPositions

lBracketOrder.SetEntryOrder pEntryOrder
pEntryOrder.Action = pAction
lBracketOrder.EntryOrder.Quantity = pQuantity

If Not pStopLossOrder Is Nothing Then
    lBracketOrder.SetStopLossOrder pStopLossOrder
    lBracketOrder.StopLossOrder.Action = IIf(pAction = OrderActionBuy, OrderActionSell, OrderActionBuy)
    lBracketOrder.StopLossOrder.Quantity = pQuantity
End If

If Not pTargetOrder Is Nothing Then
    lBracketOrder.SetTargetOrder pTargetOrder
    lBracketOrder.TargetOrder.Action = IIf(pAction = OrderActionBuy, OrderActionSell, OrderActionBuy)
    lBracketOrder.TargetOrder.Quantity = pQuantity
End If

lBracketOrder.Checkpoint

Set CreateBracketOrderWithId = lBracketOrder

LogBracketOrderMessage "Created bracket order: " & vbCrLf & _
                        lBracketOrder.ToString, _
                        mDataSource, _
                        mContract, _
                        lBracketOrder.Key, _
                        mIsSimulated, _
                        Me

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function CreateRawOrder( _
                ByVal Id As String, _
                Optional ByVal pRecoveryContextName As String) As IOrder
Const ProcName As String = "CreateRawOrder"
On Error GoTo Err

Assert IsProviderReady, "The service provider is not Ready"

Dim lOrder As New Order
lOrder.Id = Id
Set CreateRawOrder = lOrder

If Not mAccumulatedOrders Is Nothing Then mAccumulatedOrders.Add lOrder

If pRecoveryContextName <> "" And Not mBracketOrderRecoveryController Is Nothing Then
    lOrder.RecoveryContextName = pRecoveryContextName
    Dim lObj As IRecoverable: Set lObj = lOrder
    lObj.AddRecoverableObjectListener mBracketOrderRecoveryController
End If
Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSellBracketOrder( _
    ByVal Quantity As BoxedDecimal, _
    ByVal EntryOrder As IOrder, _
    Optional ByVal StopLossOrder As IOrder, _
    Optional ByVal TargetOrder As IOrder) As IBracketOrder
Const ProcName As String = "CreateSellBracketOrder"
On Error GoTo Err

Set CreateSellBracketOrder = CreateBracketOrder( _
                                        OrderActions.OrderActionSell, _
                                        Quantity, _
                                        EntryOrder, _
                                        StopLossOrder, _
                                        TargetOrder)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ExecuteBracketOrderEx(ByVal pBracketOrder As IBracketOrder, _
                    ByVal pCheckMoneyManagement As Boolean, _
                    ByVal pAddToActiveBracketOrders As Boolean)
Const ProcName As String = "ExecuteBracketOrderEx"
On Error GoTo Err

AssertArgument Not pBracketOrder Is Nothing, "No bracket order supplied"
Assert IsReadyForBracketOrder(pBracketOrder), _
        "The service provider is not Ready"

FixupEntryOrder pBracketOrder.EntryOrder
fixupStopLossOrder pBracketOrder.StopLossOrder
fixupTargetOrder pBracketOrder.TargetOrder

Dim msg As String
If TypeOf pBracketOrder Is IAutoSizingBracket Then
    Dim lBracketOrder As IAutoSizingBracket: Set lBracketOrder = pBracketOrder
    Dim lMessage As String
    
    If lBracketOrder.AutosizeType = BracketAutosizeNone Then
    ElseIf mAccountDataProvider Is Nothing Then
        Assert False, "No account data provider has been configured"
    ElseIf lBracketOrder.AutosizeValue = 0 Then
        Assert False, "Invalid autosize value"
    ElseIf Not autoSize(lBracketOrder, lMessage) Then
        pBracketOrder.EntryOrder.NotifyError ErrorCodes.ErrIllegalArgumentException, lMessage
        Exit Sub
    End If
End If

If pCheckMoneyManagement And Not mIsSimulated Then
    Dim lOrigQuantity As BoxedDecimal
    Set lOrigQuantity = pBracketOrder.EntryOrder.Quantity
    
    Dim lNewQuantity As BoxedDecimal
    Set lNewQuantity = lOrigQuantity
    If Not mOrderAuthoriser Is Nothing Then
        If Not mOrderAuthoriser.AllowOrder( _
                                pBracketOrder, _
                                mTimerList.Clock.Timestamp, _
                                lNewQuantity, _
                                msg) _
        Then
            LogBracketOrderMessage "Rejected by money management: " & msg, mDataSource, mContract, pBracketOrder.Key, mIsSimulated, Me
            pBracketOrder.Cancel
            Set pBracketOrder = Nothing
            Exit Sub
        End If
    End If
    If lNewQuantity <> lOrigQuantity Then
        LogBracketOrderMessage "Order quantity adjusted by money management: " & _
                                "was " & lOrigQuantity & _
                                "; now " & lNewQuantity, _
                                mDataSource, _
                                mContract, _
                                pBracketOrder.Key, _
                                mIsSimulated, _
                                Me
        pBracketOrder.SetNewEntryQuantity lNewQuantity
        pBracketOrder.Update
    End If
End If

LogBracketOrderMessage "Execute bracket order: " & vbCrLf & _
                        BracketOrderToString(pBracketOrder), _
                        mDataSource, _
                        mContract, _
                        pBracketOrder.Key, _
                        mIsSimulated, _
                        Me

If pAddToActiveBracketOrders Then
    If Not mActiveBracketOrders.Contains(pBracketOrder) Then
        mActiveBracketOrders.Add pBracketOrder
        GOrderUtils.Log "Order context " & mName & ": active orders count: " & mActiveBracketOrders.Count, ProcName, ModuleName, , LogLevelDetail
    End If
End If

If Not mAccumulatedBracketOrders Is Nothing Then
    If Not mAccumulatedBracketOrders.Contains(pBracketOrder) Then
        mAccumulatedBracketOrders.Add pBracketOrder
    End If
End If

setOrderRecoveryForBracketOrder pBracketOrder

fireChange OrderContextActiveStateChanged
pBracketOrder.Execute

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ExecuteBracketOrder(ByVal pBracketOrder As IBracketOrder, _
                    Optional pCheckMoneyManagement As Boolean)
Const ProcName As String = "ExecuteBracketOrder"
On Error GoTo Err

ExecuteBracketOrderEx pBracketOrder, pCheckMoneyManagement, True

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim lBracketOrder As BracketOrder

Set mDataSource = Nothing
For Each lBracketOrder In mActiveBracketOrders
    lBracketOrder.Finish
Next
Set mActiveBracketOrders = Nothing

For Each lBracketOrder In mClosedBracketOrders
    lBracketOrder.Finish
Next
Set mClosedBracketOrders = Nothing

Set mOrderSubmitter = Nothing

mNotReadyMessage = "Provider terminated"
fireChange OrderContextFinished

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub FixupEntryOrder(pOrder As Order)
Const ProcName As String = "fixupEntryOrder"
On Error GoTo Err

If pOrder Is Nothing Then Exit Sub

Dim lPriceDesc As String

Select Case pOrder.OrderType
Case OrderTypeMarket
Case OrderTypeMidprice
Case OrderTypeMarketOnOpen
Case OrderTypeMarketOnClose
Case OrderTypeMarketIfTouched
    lPriceDesc = "trigger"
    pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                pOrder.TriggerPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeMarketToLimit
Case OrderTypeLimit
    lPriceDesc = "limit"
    pOrder.LimitPrice = CalculateOffsettedPrice( _
                                pOrder.LimitPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeLimitOnOpen
    lPriceDesc = "limit"
    pOrder.LimitPrice = CalculateOffsettedPrice( _
                                pOrder.LimitPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeLimitOnClose
    lPriceDesc = "limit"
    pOrder.LimitPrice = CalculateOffsettedPrice( _
                                pOrder.LimitPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeLimitIfTouched
    lPriceDesc = "limit"
    pOrder.LimitPrice = CalculateOffsettedPrice( _
                                pOrder.LimitPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
    lPriceDesc = "trigger"
    pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                pOrder.TriggerPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeStop, _
        OrderTypeTrail
    lPriceDesc = "trigger"
    pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                pOrder.TriggerPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeStopLimit, _
        OrderTypeTrailLimit
    lPriceDesc = "limit"
    pOrder.LimitPrice = CalculateOffsettedPrice( _
                                pOrder.LimitPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
    lPriceDesc = "trigger"
    pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                pOrder.TriggerPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
End Select

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                getSourceDesignator(ModuleName, ProcName), _
                "Entry " & lPriceDesc & " price: " & Err.Description
End If
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pOrderContexts As OrderContexts, _
                ByVal pName As String, _
                ByVal pOrderManager As OrderManager, _
                ByVal pRecoveryContextName As String, _
                ByVal pGroupName As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pContractFuture As IFuture, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pOrderAuthoriser As IOrderAuthoriser, _
                ByVal pAccumulatedBracketOrders As BracketOrders, _
                ByVal pAccumulatedOrders As Orders, _
                ByVal pSimulatedClockFuture As IFuture, _
                ByVal pBracketOrderRecoveryController As BracketOrderRecoveryCtlr, _
                ByVal pAccountDataProvider As IAccountDataProvider, _
                ByVal pCurrencyConverter As ICurrencyConverter)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mOrderContextsRef = CreateWeakReference(pOrderContexts)
mName = pName
Set mOrderManager = pOrderManager
mSessionName = mOrderManager.SessionName
mRecoveryContextName = pRecoveryContextName
mGroupName = pGroupName
mIsSimulated = pIsSimulated
Set mDataSource = pDataSource
Set mOrderSubmitter = pOrderSubmitter
Set mOrderAuthoriser = pOrderAuthoriser
Set mAccumulatedBracketOrders = pAccumulatedBracketOrders
Set mAccumulatedOrders = pAccumulatedOrders

Set mBracketOrderRecoveryController = pBracketOrderRecoveryController
Set mAccountDataProvider = pAccountDataProvider
Set mCurrencyConverter = pCurrencyConverter

Set mContractFuture = pContractFuture
If mContractFuture.IsAvailable Then
    setContract mContractFuture.Value
    setTimerList
Else
    mFutureWaiter.Add mContractFuture
End If

If pSimulatedClockFuture Is Nothing Then
Else
    Set mSimulatedClockFuture = pSimulatedClockFuture
    If mSimulatedClockFuture.IsAvailable Then
        Set mClock = pSimulatedClockFuture.Value
        setTimerList
    Else
        mFutureWaiter.Add mSimulatedClockFuture
    End If
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub


Public Function IsOrderAttributeSupported( _
                ByVal Value As OrderAttributes) As Boolean
Const ProcName As String = "IsOrderAttributeSupported"
On Error GoTo Err

Assert IsProviderReady, "OrderContext is not ready"
IsOrderAttributeSupported = ((mPermittedOrderProperties.PermittedOrderAttributes And Value) = Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsOrderTifSupported( _
                ByVal Value As OrderTIFs) As Boolean
Const ProcName As String = "IsOrderTifSupported"
On Error GoTo Err

Assert IsProviderReady, "OrderContext is not ready"
IsOrderTifSupported = ((mPermittedOrderProperties.PermittedOrderTifs And Value) = Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsOrderTypeSupported( _
                ByVal Value As OrderTypes) As Boolean
Const ProcName As String = "IsOrderTypeSupported"
On Error GoTo Err

Assert IsProviderReady, "OrderContext is not ready"
IsOrderTypeSupported = ((mPermittedOrderProperties.PermittedOrderTypes And Value) = Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsReadyForAutoSize( _
                ByVal pAutosizeType As BracketAutosizeTypes, _
                ByVal pCurrency As String) As Boolean
Const ProcName As String = "IsReadyForAutoSize"
On Error GoTo Err

IsReadyForAutoSize = True

If pAutosizeType = BracketAutosizeMonetaryAmount Or pAutosizeType = BracketAutosizePercentOfAccount Then
    If mDataSource.HasCurrentTick(TickTypeAsk) Then
    ElseIf mDataSource.HasCurrentTick(TickTypeBid) Then
    ElseIf mDataSource.HasCurrentTick(TickTypeTrade) Then
    Else
        IsReadyForAutoSize = False
    End If
End If

If pAutosizeType = BracketAutosizePercentOfAccount Then
    If Not IsAccountProviderReady Then
        IsReadyForAutoSize = False
    Else
        If Not mCurrencyConverter.IsReadyCurrencies(mAccountDataProvider.BaseCurrency, "USD") Then
            mCurrencyConverter.AddStateChangeListener mAccountDataProvider.BaseCurrency, "USD", Me
            IsReadyForAutoSize = False
        End If
        If Not mCurrencyConverter.IsReadyCurrencies(mContract.Specifier.CurrencyCode, "USD") Then
            mCurrencyConverter.AddStateChangeListener mContract.Specifier.CurrencyCode, "USD", Me
            IsReadyForAutoSize = False
        End If
    End If
End If

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsReadyForBracketOrder( _
                ByVal pBracketOrder As IBracketOrder) As Boolean
Const ProcName As String = "IsReadyForBracketOrder"
On Error GoTo Err

If Not IsReady Then Exit Function

If Not IsReadyTickSize() Then Exit Function

If TypeOf pBracketOrder Is IAutoSizingBracket Then
    Dim lBracket As IAutoSizingBracket: Set lBracket = pBracketOrder
    If Not IsReadyForAutoSize(lBracket.AutosizeType, pBracketOrder.Contract.Specifier.CurrencyCode) Then Exit Function
End If

If Not isReadyForOrder(pBracketOrder.EntryOrder) Then Exit Function

If Not pBracketOrder.StopLossOrder Is Nothing Then
    If Not isReadyForOrder(pBracketOrder.StopLossOrder) Then Exit Function
End If

If Not pBracketOrder.TargetOrder Is Nothing Then
    If Not isReadyForOrder(pBracketOrder.TargetOrder) Then Exit Function
End If

IsReadyForBracketOrder = True

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsReadyForOrderType( _
                ByVal pOrderType As OrderTypes) As Boolean
Const ProcName As String = "IsReadyForOrderType"
On Error GoTo Err

IsReadyForOrderType = mOrderSubmitter.IsReadyForOrderType(pOrderType)
If IsReadyForOrderType Then
    mNotReadyMessage = ""
Else
    mNotReadyMessage = "Not ready for order type : " & OrderTypeToShortString(pOrderType)
End If

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsReadyForPriceSpecifier(ByVal pPriceSpec As PriceSpecifier) As Boolean
Const ProcName As String = "IsReadyForPriceSpecifier"
On Error GoTo Err

mNotReadyMessage = ""

If pPriceSpec Is Nothing Then
    IsReadyForPriceSpecifier = True
    Exit Function
End If

IsReadyForPriceSpecifier = False

Select Case pPriceSpec.PriceType
Case PriceValueTypeAsk
    If Not mDataSource.HasCurrentTick(TickTypeAsk) Then
        mNotReadyMessage = "Not ready for ask price"
        Exit Function
    End If
Case PriceValueTypeBid
    If Not mDataSource.HasCurrentTick(TickTypeBid) Then
        mNotReadyMessage = "Not ready for bid price"
        Exit Function
    End If
Case PriceValueTypeLast
    If Not mDataSource.HasCurrentTick(TickTypeTrade) Then
        mNotReadyMessage = "Not ready for last price"
        Exit Function
    End If
Case PriceValueTypeModel
    If Not mDataSource.HasCurrentTick(TickTypeOptionModelPrice) Then
        mNotReadyMessage = "Not ready for option model price"
        Exit Function
    End If
Case PriceValueTypeMid, _
        PriceValueTypeBidOrAsk
    If Not mDataSource.HasCurrentTick(TickTypeAsk) Then
        mNotReadyMessage = "Not ready for ask price"
        Exit Function
    ElseIf Not mDataSource.HasCurrentTick(TickTypeBid) Then
        mNotReadyMessage = "Not ready for ask price"
        Exit Function
    End If
Case PriceValueTypeEntry
End Select

If pPriceSpec.OffsetType = PriceOffsetTypeBidAskPercent Then
    If Not mDataSource.HasCurrentTick(TickTypeAsk) Then
        mNotReadyMessage = "Not ready for ask price"
        Exit Function
    End If
    If Not mDataSource.HasCurrentTick(TickTypeBid) Then
        mNotReadyMessage = "Not ready for bid price"
        Exit Function
    End If
End If

IsReadyForPriceSpecifier = True

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsReadyTickSize() As Boolean
Const ProcName As String = "IsReadyTickSize"
On Error GoTo Err

If mTickSize <> -1# Then
    mNotReadyMessage = ""
    IsReadyTickSize = True
ElseIf Not IsProviderReady Then
    mNotReadyMessage = "Not ready to submit orders"
    IsReadyTickSize = False
ElseIf getOrderContexts.IsGettingTickSize Then
    mNotReadyMessage = "Not ready: determining minimum price variation"
    IsReadyTickSize = False
ElseIf Not getOrderContexts.IsReadyTickSize Then
    mNotReadyMessage = "Not ready: determining minimum price variation"
    IsReadyTickSize = False
Else
    Assert False, "unexpected order context state"
End If

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsStopTriggerMethodSupported( _
                ByVal Value As OrderStopTriggerMethods) As Boolean
Const ProcName As String = "IsStopTriggerMethodSupported"
On Error GoTo Err

Assert IsProviderReady, "OrderContext is not ready"
IsStopTriggerMethodSupported = ((mPermittedOrderProperties.PermittedStopTriggerMethods And Value) = Value)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function NewExecution(ByVal pSetOrderRecovery As Boolean) As IExecutionReport
Const ProcName As String = "NewExecution"
On Error GoTo Err

Dim lExec As New Execution
If pSetOrderRecovery And Not mBracketOrderRecoveryController Is Nothing Then
    lExec.AddRecoverableObjectListener mBracketOrderRecoveryController
End If
lExec.GroupName = mGroupName

Set NewExecution = lExec

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function NewOrder( _
                ByVal pOrderType As OrderTypes, _
                Optional ByVal pRecoveryContextName As String, _
                Optional ByVal pBasedOn As IOrder) As IOrder
Const ProcName As String = "NewOrder"
On Error GoTo Err

Dim lOrder As Order

If pBasedOn Is Nothing Then
    Set lOrder = mTemplateOrder.Clone
    lOrder.OrderType = pOrderType
    lOrder.Id = gNextId
Else
    Set lOrder = New Order
    lOrder.Initialise pBasedOn.GroupName, pBasedOn.ContractSpecifier, pBasedOn.OrderContext
    
    lOrder.DontCheckAttributeModificationPermitted = True
    lOrder.OrderType = pOrderType
    lOrder.Id = gNextId
    
    SyncToOrder lOrder, pBasedOn
    
    lOrder.BrokerId = ""
    lOrder.ErrorCode = 0
    lOrder.ErrorMessage = ""
    lOrder.LimitPrice = GOrderUtils.MaxDoubleValue
    lOrder.Message = ""
    lOrder.FillTime = 0
    lOrder.LastFillPrice = 0#
    lOrder.QuantityFilled = DecimalZero
    lOrder.QuantityRemaining = lOrder.Quantity
    lOrder.Status = OrderStatusCreated
    lOrder.TriggerPrice = GOrderUtils.MaxDoubleValue
    lOrder.DontCheckAttributeModificationPermitted = False
End If


If Not mAccumulatedOrders Is Nothing Then mAccumulatedOrders.Add lOrder

If pRecoveryContextName <> "" And Not mBracketOrderRecoveryController Is Nothing Then
    lOrder.RecoveryContextName = pRecoveryContextName
    Dim lObj As IRecoverable: Set lObj = lOrder
    lObj.AddRecoverableObjectListener mBracketOrderRecoveryController
End If

Set NewOrder = lOrder

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Function NotifyBracketOrderClosed( _
                ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "NotifyBracketOrderClosed"
On Error GoTo Err

If mActiveBracketOrders.Contains(pBracketOrder) Then
    mActiveBracketOrders.Remove pBracketOrder
    GOrderUtils.Log "Order context " & mName & ": active orders count: " & mActiveBracketOrders.Count, ProcName, ModuleName, , LogLevelDetail
End If

If Not mClosedBracketOrders.Contains(pBracketOrder.Key) Then
    mClosedBracketOrders.Add pBracketOrder
    Set mCompletedSize = mCompletedSize.Add(pBracketOrder.Size)
    ' now check if there are any active bracket orders left
    If mActiveBracketOrders.Count = 0 Then fireChange OrderContextActiveStateChanged
End If
Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub NotifyDataReadinessChanged()
Const ProcName As String = "NotifyDataReadinessChanged"
On Error GoTo Err

IsReadyTickSize
fireChange OrderContextReadyStateChanged

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ParsePriceAndOffset( _
                ByRef pPriceSpec As PriceSpecifier, _
                ByVal pPriceString As String, _
                ByRef pMessage As String, _
                Optional ByVal pUseCloseoutSemantics As Boolean = False) As Boolean
Const ProcName As String = "ParsePriceAndOffset"
On Error GoTo Err

ParsePriceAndOffset = gParsePriceAndOffset( _
                                pPriceSpec, _
                                pPriceString, _
                                mContract.Specifier.SecType, _
                                mTickSize, _
                                pUseCloseoutSemantics, _
                                pMessage)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Public Sub RemoveChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "RemoveChangeListenerChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Sell( _
    ByVal Quantity As BoxedDecimal, _
    ByVal EntryOrder As IOrder, _
    Optional ByVal StopLossOrder As IOrder, _
    Optional ByVal TargetOrder As IOrder, _
    Optional ByVal EntryTIF As OrderTIFs = OrderTIFDay, _
    Optional ByVal StopTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
    Optional ByVal TargetTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
    Optional ByVal CancelPrice As Double, _
    Optional ByVal CancelAfter As Long) As IBracketOrder
Const ProcName As String = "Sell"
On Error GoTo Err

Set Sell = CreateBracketOrder(OrderActions.OrderActionSell, _
                                Quantity, _
                                EntryOrder, _
                                StopLossOrder, _
                                TargetOrder)

EntryOrder.TimeInForce = EntryTIF
If Not StopLossOrder Is Nothing Then StopLossOrder.TimeInForce = StopTIF
If Not TargetOrder Is Nothing Then TargetOrder.TimeInForce = TargetTIF
Sell.CancelPrice = CancelPrice
Sell.CancelAfter = CancelAfter

ExecuteBracketOrder Sell, True

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub SetTickSize(ByVal pTickSize As Double)
Const ProcName As String = "SetTickSize"
On Error GoTo Err

mTickSize = pTickSize
If mTickSize > 0# Then fireChange OrderContextReadyStateChanged

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function autoSize( _
                ByVal pBracketOrder As IAutoSizingBracket, _
                ByRef pMessage As String) As Boolean
Const ProcName As String = "autoSize"
On Error GoTo Err

Dim lBracketOrder As IBracketOrder: Set lBracketOrder = pBracketOrder

Dim lQuantity As BoxedDecimal

If pBracketOrder.AutosizeType = BracketAutosizeMonetaryAmount Then
    Set lQuantity = autoSizeAmount(pBracketOrder.AutosizeValue, _
                                pBracketOrder.LotSize, _
                                lBracketOrder.LongPosition, _
                                pMessage)
Else
    Set lQuantity = autoSizePercent(pBracketOrder.AutosizeValue, _
                                pBracketOrder.LotSize, _
                                lBracketOrder.LongPosition, _
                                pMessage)
End If

LogBracketOrderMessage pMessage, _
                        mDataSource, _
                        mContract, _
                        lBracketOrder.Key, _
                        mIsSimulated, _
                        Me

If lQuantity <> 0 Then
    lBracketOrder.SetNewEntryQuantity lQuantity
    lBracketOrder.Update
    autoSize = True
Else
    lBracketOrder.Cancel
    autoSize = False
End If


Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function autoSizeAmount( _
                ByVal pAmount As Double, _
                ByVal pLotSize As Long, _
                ByVal pIsLong As Boolean, _
                ByRef pMessage As String) As BoxedDecimal
Const ProcName As String = "autoSizeAmount"
On Error GoTo Err

Dim lPrice As Double

If pIsLong Then
    If mDataSource.HasCurrentTick(TickTypeAsk) Then
        lPrice = mDataSource.CurrentTick(TickTypeAsk).Price
    ElseIf mDataSource.HasCurrentTick(TickTypeTrade) Then
        lPrice = mDataSource.CurrentTick(TickTypeTrade).Price
    ElseIf mDataSource.HasCurrentTick(TickTypeBid) Then
        lPrice = mDataSource.CurrentTick(TickTypeBid).Price
    End If
Else
    If mDataSource.HasCurrentTick(TickTypeBid) Then
        lPrice = mDataSource.CurrentTick(TickTypeBid).Price
    ElseIf mDataSource.HasCurrentTick(TickTypeTrade) Then
        lPrice = mDataSource.CurrentTick(TickTypeTrade).Price
    ElseIf mDataSource.HasCurrentTick(TickTypeAsk) Then
        lPrice = mDataSource.CurrentTick(TickTypeAsk).Price
    End If
End If

Dim lQuantity As Long
If lPrice = 0 Then
    pMessage = "No price currently available to size order"
Else
    lQuantity = Int(pAmount / (mContract.Specifier.Multiplier * lPrice))
    If pLotSize <> 0 Then lQuantity = pLotSize * Int(lQuantity / pLotSize)
    
    If lQuantity <> 0 Then
        pMessage = "Order quantity set to " & lQuantity
    Else
        pMessage = "Order quantity reduced to 0 due to lack of funds"
    End If
End If

Set autoSizeAmount = CreateBoxedDecimal(lQuantity)

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function autoSizePercent( _
                ByVal pPercent As Double, _
                ByVal pLotSize As Long, _
                ByVal pIsLong As Boolean, _
                ByRef pMessage As String) As BoxedDecimal
Const ProcName As String = "autoSizePercent"
On Error GoTo Err

Dim lAccountBalance As Double
If mFixedAccountBalance <> 0# Then
    lAccountBalance = mFixedAccountBalance
Else
    lAccountBalance = CDbl(mAccountDataProvider.GetAccountValue("NetLiquidation", mAccountDataProvider.BaseCurrency).Value)
End If

If mContract.Specifier.CurrencyCode <> mAccountDataProvider.BaseCurrency Then
    If mAccountDataProvider.BaseCurrency <> "USD" Then
        lAccountBalance = mCurrencyConverter.Convert(lAccountBalance, mAccountDataProvider.BaseCurrency, "USD")
    End If
    If mContract.Specifier.CurrencyCode <> "USD" Then
        lAccountBalance = mCurrencyConverter.Convert(lAccountBalance, "USD", mContract.Specifier.CurrencyCode)
    End If
End If

Dim lAmount As Double
lAmount = Int(pPercent * lAccountBalance / 100#)
Set autoSizePercent = CreateBoxedDecimal(CDec(autoSizeAmount(lAmount, pLotSize, pIsLong, pMessage)))

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fireChange(ByVal pChangeType As OrderContextChangeTypes)
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
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fixupStopLossOrder(pOrder As Order)
Const ProcName As String = "fixupStopLossOrder"
On Error GoTo Err

If pOrder Is Nothing Then Exit Sub

Dim lPriceDesc As String

Select Case pOrder.OrderType
Case OrderTypeStop, _
        OrderTypeTrail
    lPriceDesc = "trigger"
    If pOrder.TriggerPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                    pOrder.TriggerPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
Case OrderTypeStopLimit, _
        OrderTypeTrailLimit
    lPriceDesc = "limit"
    If pOrder.LimitPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.LimitPrice = CalculateOffsettedPrice( _
                                    pOrder.LimitPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
    
    lPriceDesc = "trigger"
    If pOrder.TriggerPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                    pOrder.TriggerPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
End Select

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                getSourceDesignator(ModuleName, ProcName), _
                "Entry " & lPriceDesc & " price: " & Err.Description
End If
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fixupTargetOrder(pOrder As Order)
Const ProcName As String = "fixupTargetOrder"
On Error GoTo Err

If pOrder Is Nothing Then Exit Sub

Dim lPriceDesc As String

Select Case pOrder.OrderType
Case OrderTypeLimit, _
        OrderTypeLimitOnClose, _
        OrderTypeLimitOnOpen
    lPriceDesc = "limit"
    If pOrder.LimitPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.LimitPrice = CalculateOffsettedPrice( _
                                    pOrder.LimitPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
Case OrderTypeLimitIfTouched
    lPriceDesc = "limit"
    If pOrder.LimitPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.LimitPrice = CalculateOffsettedPrice( _
                                    pOrder.LimitPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
    lPriceDesc = "trigger"
    If pOrder.TriggerPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                    pOrder.TriggerPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
Case OrderTypeMarketIfTouched
    lPriceDesc = "trigger"
    If pOrder.TriggerPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                    pOrder.TriggerPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
Case OrderTypeMarketOnClose, _
        OrderTypeMarketOnOpen
End Select

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                getSourceDesignator(ModuleName, ProcName), _
                "Entry " & lPriceDesc & " price: " & Err.Description
End If
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getCurrentPrice(ByVal pTickType As TickTypes) As Double
Const ProcName As String = "getCurrentPrice"
On Error GoTo Err

If isDataSourceRunning Then getCurrentPrice = mDataSource.CurrentTick(pTickType).Price

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function getOrderContexts() As OrderContexts
Set getOrderContexts = mOrderContextsRef.Target
End Function

Private Function getSourceDesignator( _
                ByRef pModuleName As String, _
                ByRef pProcedureName As String) As String
getSourceDesignator = "[" & ProjectName & "." & _
            pModuleName & ":" & _
            pProcedureName & "]"
End Function

Private Property Get IsAccountProviderReady() As Boolean
IsAccountProviderReady = (mAccountDataProvider.State = AccountProviderReady)
End Property

Private Function isDataSourceRunning()
If mDataSource Is Nothing Then
    isDataSourceRunning = False
Else
    isDataSourceRunning = (mDataSource.State = MarketDataSourceStateRunning)
End If
End Function

Private Function isReadyForOrder( _
                ByVal pOrder As IOrder) As Boolean
Const ProcName As String = "isReadyForOrder"
On Error GoTo Err

If Not IsReadyForOrderType(pOrder.OrderType) Then Exit Function
If Not IsReadyForPriceSpecifier(pOrder.LimitPriceSpec) Then Exit Function
If Not IsReadyForPriceSpecifier(pOrder.TriggerPriceSpec) Then Exit Function

isReadyForOrder = True

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setContract(ByVal pContract As IContract)
Const ProcName As String = "setContract"
On Error GoTo Err

Set mContract = pContract
    
Set mTemplateOrder = New Order
mTemplateOrder.RecoveryContextName = mRecoveryContextName
mTemplateOrder.Initialise mGroupName, mContract.Specifier, Me
mTemplateOrder.IsSimulated = mIsSimulated

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrderRecoveryForBracketOrder(ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "setOrderRecoveryForBracketOrder"
On Error GoTo Err

If mBracketOrderRecoveryController Is Nothing Then Exit Sub

Dim lObj As IRecoverable
Set lObj = pBracketOrder
lObj.AddRecoverableObjectListener mBracketOrderRecoveryController

Set lObj = pBracketOrder.EntryOrder
If Not lObj Is Nothing Then lObj.AddRecoverableObjectListener mBracketOrderRecoveryController

Set lObj = pBracketOrder.StopLossOrder
If Not lObj Is Nothing Then lObj.AddRecoverableObjectListener mBracketOrderRecoveryController

Set lObj = pBracketOrder.TargetOrder
If Not lObj Is Nothing Then lObj.AddRecoverableObjectListener mBracketOrderRecoveryController

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setTimerList()
Const ProcName As String = "setTimerList"
On Error GoTo Err

If mContract Is Nothing Then Exit Sub

If mClock Is Nothing Then
    Set mClock = GetClock(mContract.TimezoneName)
    Set mTimerList = GetGlobalTimerList
ElseIf mClock.IsSimulated Then
    Set mTimerList = GetSimulatedTimerList(mClock)
Else
    Set mTimerList = GetGlobalTimerList
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

