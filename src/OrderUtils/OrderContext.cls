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

Private mOrderManager                   As OrderManager
Private mGroupName                      As String
Private mRecoveryContextName            As String
Private mName                           As String

Private mSessionName                    As String

Private mContractFuture                 As IFuture
Private mContract                       As IContract
Private mDataSource                     As IMarketDataSource
Attribute mDataSource.VB_VarHelpID = -1
Private mOrderSubmitter                 As IOrderSubmitter
Private mOrderAuthoriser                As IOrderAuthoriser
Private mAccumulatedBracketOrders       As BracketOrders
Private mAccumulatedOrders              As Orders


Private mIsSimulated                    As Boolean

Private mActiveBracketOrders            As BracketOrders
    ' this contains all bracket orders for which one or more orders are
    ' not yet complete (or not yet cancelled)
    
Private mClosedBracketOrders            As BracketOrders
    ' this contains all bracket orders for which all orders have completed (or
    ' been cancelled)
    
Private mCompletedSize                  As Long
    ' the total Size for all completed bracket orders (may be non-zero if one
    ' or more completed bracket orders has no target or stop-loss orders)
    
Private mPermittedOrderProperties       As IPermittedOrderProperties

Private mEnabled                        As Boolean
Private mProviderIsReady                As Boolean
Private mDataSourceIsReady              As Boolean

Private mTemplateOrder                  As Order

Private mPreventUnprotectedPositions    As Boolean

Private mChangeListeners                As New Listeners

Private mTimerList                      As TimerList

Private WithEvents mFutureWaiter        As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mBracketOrderRecoveryController As BracketOrderRecoveryCtlr

Private mOrderContextsRef               As WeakReference

Private mTickSizeFuture                 As IFuture
Private mTickSize                       As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mActiveBracketOrders = New BracketOrders
Set mClosedBracketOrders = New BracketOrders
Set mFutureWaiter = New FutureWaiter
mTickSize = -1
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderContext terminated"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is IContract Then
    setContract ev.Future.Value
    IsReadyTickSize
    fireChange OrderContextReadyStateChanged
ElseIf TypeOf ev.Future.Value Is Clock Then
    setTimerListFromSimulatedClock ev.Future.Value
    fireChange OrderContextReadyStateChanged
ElseIf TypeOf ev.Future.Value Is BoxedValue Then
    Dim lBoxedValue As BoxedValue: Set lBoxedValue = ev.Future.Value
    mTickSize = lBoxedValue.Value
    If mTickSize <> 0# Then fireChange OrderContextReadyStateChanged
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ActiveBracketOrders() As BracketOrders
Set ActiveBracketOrders = mActiveBracketOrders
End Property

Public Property Get ActiveSize() As Long
Const ProcName As String = "ActiveSize"
On Error GoTo Err

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    ActiveSize = ActiveSize + lBracketOrder.Size
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ClosedBracketOrders() As BracketOrders
Set ClosedBracketOrders = mClosedBracketOrders
End Property

Public Property Get CompletedSize() As Long
CompletedSize = mCompletedSize
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get GroupName() As String
GroupName = mGroupName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsDataSourceReady() As Boolean
IsDataSourceReady = mDataSourceIsReady
End Property

Friend Property Let IsProviderReady(ByVal Value As Boolean)
Const ProcName As String = "IsProviderReady"
On Error GoTo Err

If mProviderIsReady = Value Then Exit Property

mProviderIsReady = Value

fireChange OrderContextReadyStateChanged

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsProviderReady() As Boolean
IsProviderReady = mProviderIsReady
End Property

Public Property Get IsReady() As Boolean
IsReady = (mProviderIsReady And mEnabled And Not mContract Is Nothing And Not mTimerList Is Nothing)
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Get Name() As String
Name = mName
End Property

Friend Property Get OrderManager() As OrderManager
Set OrderManager = mOrderManager
End Property

Public Property Get OrderSubmitter() As IOrderSubmitter
Set OrderSubmitter = mOrderSubmitter
End Property

Public Property Get PendingBuySize() As Long
Const ProcName As String = "PendingBuySize"
On Error GoTo Err

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    If lBracketOrder.LongPosition Then PendingBuySize = PendingBuySize + lBracketOrder.PendingSize
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PendingSellSize() As Long
Const ProcName As String = "PendingSellSize"
On Error GoTo Err

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    If Not lBracketOrder.LongPosition Then PendingSellSize = PendingSellSize + lBracketOrder.PendingSize
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PendingSize() As Long
Const ProcName As String = "PendingSize"
On Error GoTo Err

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    PendingSize = PendingSize + lBracketOrder.PendingSize
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ScopeName() As String
ScopeName = mBracketOrderRecoveryController.ScopeName
End Property

Public Property Get Size() As Long
Const ProcName As String = "Size"
On Error GoTo Err

Size = mCompletedSize

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    Size = Size + lBracketOrder.Size
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TemplateOrder() As IOrder
Set TemplateOrder = mTemplateOrder
End Property

Public Property Get TickSize() As Double
Assert mTickSize <> 0#, "TickSize not yet available"
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

If pBracketOrder.IsComplete And pBracketOrder.RolloverSpecification Is Nothing Then
    NotifyBracketOrderClosed pBracketOrder
Else
    mActiveBracketOrders.Add pBracketOrder
    fireChange OrderContextActiveStateChanged
End If
pBracketOrder.Activate

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "AddChangeListener"
On Error GoTo Err

mChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Creates a buy <code>BracketOrder</code>
'
' @return the <code>BracketOrder</code> object created
'
' @param number         quantity to buy
'       EntryType  the type of order to enter the position:
'           BracketEntryTypeMarket            a market order
'           BracketEntryTypeMarketOnOpen      a market on open order
'           BracketEntryTypeMarketOnClose     a market on close order
'           BracketEntryTypeMarketIfTouched   a market if touched order
'           BracketEntryTypeMarketToLimit     a market to limit order
'           BracketEntryTypeBid               a limit order at the current Bid
'           BracketEntryTypeAsk               a limit order at the current Ask
'           BracketEntryTypeLast              a limit order at the last Trade Price
'           BracketEntryTypeLimit             a limit order
'           BracketEntryTypeLimitOnOpen       a limit on open order
'           BracketEntryTypeLimitOnClose      a limit on close order
'           BracketEntryTypeLimitIfTouched    a limit if touched order
'           BracketEntryTypeStop              a stop order
'           BracketEntryTypeStopLimit         a stop limit order
' @param BuyPrice       the limit Price for a limit order (ignored for BracketEntryTypeBid,
'                       BracketEntryTypeAsk and BracketEntryTypeLast)
' @param BuyOffset      number of ticks to modify the BuyPrice. For example, to place an order
'                       one tick below the current Ask, use BracketEntryTypeAsk and set
'                       BuyOffset to -1
' @param TriggerPrice   used with BracketEntryTypeMarketIfTouched, BracketEntryTypeLimitIfTouched,
'                       BracketEntryTypeStop and BracketEntryTypeStopLimit to specify the Price
'                       at which the order is triggered
' @param StopLossType  the type of stop-loss order to place:
'           BracketStopLossTypeNone               no stop-loss order
'           OrderTypestop               use a stop order
'           OrderTypestopLimit          use a stop limit order
'           BracketStopLossTypeBid                use a stop order at the current Bid (see StopOffset)
'           BracketStopLossTypeAsk                use a stop order at the current Ask (see StopOffset)
'           BracketStopLossTypeLast               use a stop order at the last Trade Price (see StopOffset)
'           BracketStopLossTypeAuto               use a stop order at the fill Price for the
'                                           entry order (see StopOffset)
' @param StopPrice      the trigger Price for OrderTypestop and OrderTypestopLimit
' @param StopOffset     number of ticks to modify the StopPrice. For example, for a stop-loss
'                       order 4 ticks below the entry order fill Price, use BracketStopLossTypeAuto
'                       and set StopOffset to -4
' @param StopLimitPrice the limit Price for OrderTypestopLimit
' @param TargetType the type of target order to place:
'           BracketTargetTypeNone             no target order
'           BracketTargetTypeLimit            use a limit order
'           BracketTargetTypeLimitIfTouched   use a limit if touched order
'           BracketTargetTypeMarketIfTouched  use a market if touched order
'           BracketTargetTypeBid              use a limit order at the current Bid (see TargetOffset)
'           BracketTargetTypeAsk              use a limit order at the current Ask (see TargetOffset)
'           BracketTargetTypeLast             use a limit order at the last Trade Price (see TargetOffset)
'           BracketTargetTypeAuto             use a limit order at the fill Price for the
'                                           entry order (see TargetOffset)
' @param TargetPrice    the limit Price for the target order
' @param TargetOffset   number of ticks to modify the TargetPrice. For example, for a target
'                       order 10 ticks above the entry order fill Price, use BracketTargetTypeAuto
'                       and set TargetOffset to 10
' @param TargetTriggerPrice  the trigger Price for BracketTargetTypeLimitIfTouched and BracketTargetTypeMarketIfTouched
' @param CancelPrice    if the entry order is a limit order and market Price rises to this
'                       Price before the entry order is filled, then the order is cancelled.
'                       If the entry order is a stop or stop limit order and market Price
'                       falls to this Price before the entry order is filled, then the order
'                       is cancelled
' @param EntryTIF       specifies the time in force for the entry order
' @param StopTIF        specifies the time in force for the stop-loss order
' @param TargetTIF      specifies the time in force for the target order
'
' @see
'
'@/
Public Function Buy( _
    ByVal Number As Long, _
    ByVal pEntryOrder As IOrder, _
    Optional ByVal pStopLossOrder As IOrder, _
    Optional ByVal pTargetOrder As IOrder, _
    Optional ByVal EntryTIF As OrderTIFs = OrderTIFDay, _
    Optional ByVal StopTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
    Optional ByVal TargetTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
    Optional ByVal CancelPrice As Double, _
    Optional ByVal CancelAfter As Long) As IBracketOrder
Const ProcName As String = "Buy"
On Error GoTo Err

'Dim s As String
's = "Buy" & _
'                IIf(mIsSimulated, " (simulated): ", ": ") & mContract.Specifier.LocalSymbol & _
'                "; Number=" & Number & vbCrLf & _
'                "; EntryType=" & gBracketEntryTypeToShortString(EntryType) & _
'                "; EntryLimitPrice=" & EntryLimitPrice & _
'                "; EntryTriggerPrice=" & EntryTriggerPrice & vbCrLf & _
'                "; EntryLimitPriceOffset=" & EntryLimitPriceOffset & _
'                "; EntryLimitPriceOffsetType=" & gPriceOffsetTypeToString(EntryLimitPriceOffsetType) & _
'                "; EntryTriggerPriceOffset=" & EntryTriggerPriceOffset & _
'                "; EntryTriggerPriceOffsetType=" & gPriceOffsetTypeToString(EntryTriggerPriceOffsetType) & vbCrLf & _
'                "; StopLossType=" & gBracketStopLossTypeToShortString(StopLossType) & _
'                "; StopLossLimitPrice=" & StopLossLimitPrice & _
'                "; StopLossTriggerPrice=" & StopLossTriggerPrice & vbCrLf & _
'                "; StopLossLimitPriceOffset=" & StopLossLimitPriceOffset & _
'                "; StopLossLimitPriceOffsetType=" & gPriceOffsetTypeToString(StopLossLimitPriceOffsetType) & _
'                "; StopLossTriggertPriceOffset=" & StopLossTriggerPriceOffset & _
'                "; StopLossTriggerPriceOffsetType=" & gPriceOffsetTypeToString(StopLossTriggerPriceOffsetType) & vbCrLf & _
'                "; TargetType=" & gBracketTargetTypeToShortString(TargetType) & _
'                "; TargetLimitPrice=" & TargetLimitPrice & _
'                "; TargetTriggerPrice=" & TargetTriggerPrice & vbCrLf & _
'                "; TargetLimitPriceOffset=" & TargetLimitPriceOffset & _
'                "; TargetLimitPriceOffsetType=" & gPriceOffsetTypeToString(TargetLimitPriceOffsetType)
's = s & _
'                "; TargetTriggerPriceOffset=" & TargetTriggerPriceOffset & _
'                "; TargetTriggerPriceOffsetType=" & gPriceOffsetTypeToString(TargetTriggerPriceOffsetType) & vbCrLf & _
'                "; EntryTIF=" & gOrderTIFToString(EntryTIF) & _
'                "; StopTIF=" & gOrderTIFToString(StopTIF) & _
'                "; TargetTIF=" & gOrderTIFToString(TargetTIF)
'gLog s, ProcName, ModuleName
                
Set Buy = CreateBracketOrder(OrderActions. _
                            OrderActionBuy, _
                            Number, _
                            pEntryOrder, _
                            pStopLossOrder, _
                            pTargetOrder)

pEntryOrder.TimeInForce = EntryTIF
If Not pStopLossOrder Is Nothing Then pStopLossOrder.TimeInForce = StopTIF
If Not pTargetOrder Is Nothing Then pTargetOrder.TimeInForce = TargetTIF
Buy.CancelPrice = CancelPrice
Buy.CancelAfter = CancelAfter

ExecuteBracketOrder Buy, True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CalculateOffsettedPrice( _
                ByVal pPriceSpec As PriceSpecifier, _
                ByVal pSectype As SecurityTypes, _
                ByVal pOrderAction As OrderActions) As Double
Const ProcName As String = "CalculateOffsettedPrice"
On Error GoTo Err

If pPriceSpec.PriceType = PriceValueTypeEntry Or _
        pPriceSpec.PriceType = PriceValueTypeNone Then
    CalculateOffsettedPrice = MaxDoubleValue
    Exit Function
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
    lPrice = getCurrentPrice(TickTypeAsk)
Case PriceValueTypeBid
    lPrice = getCurrentPrice(TickTypeBid)
Case PriceValueTypeLast
    lPrice = getCurrentPrice(TickTypeTrade)
Case PriceValueTypeEntry
    Assert False
Case PriceValueTypeMid
    lPrice = gRoundToTickBoundary((getCurrentPrice(TickTypeAsk) + getCurrentPrice(TickTypeBid)) / 2#, mTickSize, TickRoundingModeNearest)
Case PriceValueTypeBidOrAsk
    If pOrderAction = OrderActionBuy Then
        lPrice = getCurrentPrice(TickTypeBid)
    Else
        lPrice = getCurrentPrice(TickTypeAsk)
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
    Dim lBid As Double: lBid = getCurrentPrice(TickTypeBid)
    Dim lAsk As Double: lAsk = getCurrentPrice(TickTypeAsk)
    
    lOffset = pPriceSpec.Offset * (lAsk - lBid) / 100#
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

CalculateOffsettedPrice = gRoundToTickBoundary(lResult, mTickSize, lRoundingMode)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CloseOut() As TaskController
Const ProcName As String = "CloseOut"
On Error GoTo Err

Dim t As New CloseoutOrderContextTask

Set CloseOut = StartTask(New CloseoutOrderContextTask, PriorityNormal, , , Me)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBuyBracketOrder( _
    ByVal Number As Long, _
    ByVal pEntryOrder As IOrder, _
    Optional ByVal pStopLossOrder As IOrder, _
    Optional ByVal pTargetOrder As IOrder) As IBracketOrder
Const ProcName As String = "CreateBuyBracketOrder"
On Error GoTo Err

Set CreateBuyBracketOrder = CreateBracketOrder( _
                                        OrderActions.OrderActionBuy, _
                                        Number, _
                                        pEntryOrder, _
                                        pStopLossOrder, _
                                        pTargetOrder)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateEntryOrder( _
                    ByVal pEntryType As OrderTypes, _
                    Optional ByVal pLimitPriceSpec As PriceSpecifier, _
                    Optional ByVal pTriggerPriceSpec As PriceSpecifier, _
                    Optional ByVal pOutsideRTH As Boolean = False) As IOrder
Const ProcName As String = "CreateEntryOrder"
On Error GoTo Err

AssertArgument gIsEntryOrderType(pEntryType), "Invalid entry order type"

Dim lOrder As IOrder
Set lOrder = NewOrder(pEntryType)

lOrder.LimitPriceSpec = pLimitPriceSpec
lOrder.TriggerPriceSpec = pTriggerPriceSpec
lOrder.IgnoreRegularTradingHours = pOutsideRTH

Select Case pEntryType
Case OrderTypeMarket
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateStopLossOrder( _
                    ByVal pStopLossType As OrderTypes, _
                    Optional ByVal pLimitPriceSpec As PriceSpecifier, _
                    Optional ByVal pTriggerPriceSpec As PriceSpecifier, _
                    Optional ByVal pOutsideRTH As Boolean = False) As IOrder
Const ProcName As String = "CreateStopLossOrder"
On Error GoTo Err

AssertArgument gIsStopLossOrderType(pStopLossType), "Invalid stop-loss order type"

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
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTargetOrder( _
                    ByVal pTargetType As OrderTypes, _
                    Optional ByVal pLimitPriceSpec As PriceSpecifier, _
                    Optional ByVal pTriggerPriceSpec As PriceSpecifier, _
                    Optional ByVal pOutsideRTH As Boolean = False) As IOrder
Const ProcName As String = "CreateTargetOrder"
On Error GoTo Err

AssertArgument gIsTargetOrderType(pTargetType), "Invalid target order type"

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
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateBracketOrder( _
                ByVal pAction As OrderActions, _
                ByVal pQuantity As Long, _
                ByVal pEntryOrder As IOrder, _
                Optional ByVal pStopLossOrder As IOrder, _
                Optional ByVal pTargetOrder As IOrder) As IBracketOrder
Const ProcName As String = "CreateBracketOrder"
On Error GoTo Err

AssertArgument pQuantity > 0, "Quantity must be greater than zero"
AssertArgument Not pEntryOrder Is Nothing, "Entry order cannot be Nothing"

AssertArgument gIsEntryOrderType(pEntryOrder.OrderType), "Invalid entry order type"
If Not pStopLossOrder Is Nothing Then AssertArgument gIsStopLossOrderType(pStopLossOrder.OrderType), "Invalid stop-loss order type"
If Not pTargetOrder Is Nothing Then AssertArgument gIsTargetOrderType(pTargetOrder.OrderType), "Invalid target order type"

Dim lBracketOrder As New BracketOrder
lBracketOrder.Initialise GIdProvider.gNextId, mRecoveryContextName, mGroupName, Me, mContract, mOrderSubmitter, mDataSource, mIsSimulated
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

Set CreateBracketOrder = lBracketOrder

gLogBracketOrderMessage "Created bracket order: " & vbCrLf & _
                        lBracketOrder.ToString, _
                        mDataSource, _
                        mContract, _
                        lBracketOrder.Key, _
                        mIsSimulated, _
                        Me

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSellBracketOrder( _
    ByVal Number As Long, _
    ByVal pEntryOrder As IOrder, _
    Optional ByVal pStopLossOrder As IOrder, _
    Optional ByVal pTargetOrder As IOrder) As IBracketOrder
Const ProcName As String = "CreateSellBracketOrder"
On Error GoTo Err

Set CreateSellBracketOrder = CreateBracketOrder( _
                                        OrderActions.OrderActionSell, _
                                        Number, _
                                        pEntryOrder, _
                                        pStopLossOrder, _
                                        pTargetOrder)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ExecuteBracketOrderEx(ByVal pBracketOrder As BracketOrder, _
                    ByVal pCheckMoneyManagement As Boolean, _
                    ByVal pAddToActiveBracketOrders As Boolean)
Const ProcName As String = "ExecuteBracketOrderEx"
On Error GoTo Err

AssertArgument Not pBracketOrder Is Nothing, "No bracket order supplied"
Assert IsReadyForBracketOrder(pBracketOrder), _
        "The service provider is not Ready"

fixupEntryOrder pBracketOrder.EntryOrder
fixupStopLossOrder pBracketOrder.StopLossOrder
fixupTargetOrder pBracketOrder.TargetOrder

If pCheckMoneyManagement And Not mIsSimulated Then
    Dim lOrigQuantity As Long
    lOrigQuantity = pBracketOrder.EntryOrder.Quantity
    
    Dim lNewQuantity As Long
    lNewQuantity = lOrigQuantity
    If Not mOrderAuthoriser Is Nothing Then
        Dim msg As String
        If Not mOrderAuthoriser.AllowOrder( _
                                IIf(pBracketOrder.EntryOrder.Action = OrderActionBuy, True, False), _
                                lNewQuantity, _
                                pBracketOrder.Risk, _
                                msg) _
        Then
            gLogBracketOrderMessage "Rejected by money management: " & msg, mDataSource, mContract, pBracketOrder.Key, mIsSimulated, Me
            Set pBracketOrder = Nothing
            Exit Sub
        End If
    End If
    If lNewQuantity <> lOrigQuantity Then
        pBracketOrder.SetNewEntryQuantity lNewQuantity
        pBracketOrder.Update
    End If
End If

gLogBracketOrderMessage "Execute bracket order: " & vbCrLf & _
                        gBracketOrderToString(pBracketOrder), _
                        mDataSource, _
                        mContract, _
                        pBracketOrder.Key, _
                        mIsSimulated, _
                        Me

If pAddToActiveBracketOrders Then
    mActiveBracketOrders.Add pBracketOrder
End If

If Not mAccumulatedBracketOrders Is Nothing Then mAccumulatedBracketOrders.Add pBracketOrder
setOrderRecoveryForBracketOrder pBracketOrder

fireChange OrderContextActiveStateChanged
pBracketOrder.Execute

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ExecuteBracketOrder(ByVal pBracketOrder As BracketOrder, _
                    Optional pCheckMoneyManagement As Boolean)
Const ProcName As String = "ExecuteBracketOrder"
On Error GoTo Err

ExecuteBracketOrderEx pBracketOrder, pCheckMoneyManagement, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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
                ByVal pBracketOrderRecoveryController As BracketOrderRecoveryCtlr)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mOrderContextsRef = CreateWeakReference(pOrderContexts)
mName = pName
Set mOrderManager = pOrderManager
mSessionName = mOrderManager.SessionName
mRecoveryContextName = pRecoveryContextName
mGroupName = pGroupName
mIsSimulated = pIsSimulated

Set mContractFuture = pContractFuture
If mContractFuture.IsAvailable Then
    setContract mContractFuture.Value
    IsReadyTickSize
Else
    mFutureWaiter.Add mContractFuture
End If

If pSimulatedClockFuture Is Nothing Then
    Set mTimerList = GetGlobalTimerList
ElseIf pSimulatedClockFuture.IsAvailable Then
    setTimerListFromSimulatedClock pSimulatedClockFuture.Value
Else
    mFutureWaiter.Add pSimulatedClockFuture
End If

Set mDataSource = pDataSource
Set mOrderSubmitter = pOrderSubmitter
Set pOrderAuthoriser = pOrderAuthoriser
Set mAccumulatedBracketOrders = pAccumulatedBracketOrders
Set mAccumulatedOrders = pAccumulatedOrders

Set mBracketOrderRecoveryController = pBracketOrderRecoveryController

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


Public Function IsOrderAttributeSupported( _
                ByVal Value As OrderAttributes) As Boolean
Const ProcName As String = "IsOrderAttributeSupported"
On Error GoTo Err

Assert IsProviderReady, "OrderContext is not ready"
IsOrderAttributeSupported = ((mPermittedOrderProperties.PermittedOrderAttributes And Value) = Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsOrderTifSupported( _
                ByVal Value As OrderTIFs) As Boolean
Const ProcName As String = "IsOrderTifSupported"
On Error GoTo Err

Assert IsProviderReady, "OrderContext is not ready"
IsOrderTifSupported = ((mPermittedOrderProperties.PermittedOrderTifs And Value) = Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsOrderTypeSupported( _
                ByVal Value As OrderTypes) As Boolean
Const ProcName As String = "IsOrderTypeSupported"
On Error GoTo Err

Assert IsProviderReady, "OrderContext is not ready"
IsOrderTypeSupported = ((mPermittedOrderProperties.PermittedOrderTypes And Value) = Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsReadyForBracketOrder( _
                ByVal pBracketOrder As IBracketOrder) As Boolean
Const ProcName As String = "IsReadyForBracketOrder"
On Error GoTo Err

If Not IsReady Then Exit Function

If Not IsReadyTickSize() Then Exit Function

If Not IsReadyForOrderType(pBracketOrder.EntryOrder.OrderType) Then Exit Function
If Not IsReadyForPriceSpecifier(pBracketOrder.EntryOrder.LimitPriceSpec) Then Exit Function
If Not IsReadyForPriceSpecifier(pBracketOrder.EntryOrder.TriggerPriceSpec) Then Exit Function

If Not pBracketOrder.StopLossOrder Is Nothing Then
    If Not IsReadyForOrderType(pBracketOrder.StopLossOrder.OrderType) Then Exit Function
    If Not IsReadyForPriceSpecifier(pBracketOrder.StopLossOrder.LimitPriceSpec) Then Exit Function
    If Not IsReadyForPriceSpecifier(pBracketOrder.StopLossOrder.TriggerPriceSpec) Then Exit Function
End If

If Not pBracketOrder.TargetOrder Is Nothing Then
    If Not IsReadyForOrderType(pBracketOrder.TargetOrder.OrderType) Then Exit Function
    If Not IsReadyForPriceSpecifier(pBracketOrder.TargetOrder.LimitPriceSpec) Then Exit Function
    If Not IsReadyForPriceSpecifier(pBracketOrder.TargetOrder.TriggerPriceSpec) Then Exit Function
End If

IsReadyForBracketOrder = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsReadyForOrderType(ByVal pOrderType As OrderTypes) As Boolean
Const ProcName As String = "IsReadyForOrderType"
On Error GoTo Err

IsReadyForOrderType = mOrderSubmitter.IsReadyForOrderType(pOrderType)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsReadyForPriceSpecifier(ByVal pPriceSpec As PriceSpecifier) As Boolean
Const ProcName As String = "IsReadyForPriceSpecifier"
On Error GoTo Err

If pPriceSpec Is Nothing Then
    IsReadyForPriceSpecifier = True
    Exit Function
End If

IsReadyForPriceSpecifier = False

Select Case pPriceSpec.PriceType
Case PriceValueTypeAsk
    If Not mDataSource.HasCurrentTick(TickTypeAsk) Then Exit Function
Case PriceValueTypeBid
    If Not mDataSource.HasCurrentTick(TickTypeBid) Then Exit Function
Case PriceValueTypeLast
    If Not mDataSource.HasCurrentTick(TickTypeTrade) Then Exit Function
Case PriceValueTypeMid, _
        PriceValueTypeBidOrAsk
    If Not mDataSource.HasCurrentTick(TickTypeAsk) Then Exit Function
    If Not mDataSource.HasCurrentTick(TickTypeBid) Then Exit Function
Case PriceValueTypeEntry
End Select

IsReadyForPriceSpecifier = pPriceSpec.OffsetType <> PriceOffsetTypeBidAskPercent Or _
                    (mDataSource.HasCurrentTick(TickTypeAsk) And _
                    mDataSource.HasCurrentTick(TickTypeBid))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsReadyTickSize() As Boolean
Const ProcName As String = "IsReadyTickSize"
On Error GoTo Err

If mTickSize <> -1# Then
    IsReadyTickSize = True
Else
    If Not mTickSizeFuture Is Nothing Then
        IsReadyTickSize = False
    Else
        Select Case mContract.Specifier.SecType
        Case SecTypeStock
            If dataSourceHasPrice Then
                Set mTickSizeFuture = mOrderSubmitter.GetTickSize(Nothing)
                mFutureWaiter.Add mTickSizeFuture
            End If
            IsReadyTickSize = False
        Case SecTypeFuture
            mTickSize = mContract.TickSize
            IsReadyTickSize = True
        Case SecTypeOption, _
                SecTypeFuturesOption
            If dataSourceHasPrice Then
                Dim lDummyOrder As IOrder
                Set lDummyOrder = CreateEntryOrder(OrderTypeLimit, gNewPriceSpecifier(, PriceValueTypeAsk), gNewPriceSpecifier, True)
                Set mTickSizeFuture = mOrderSubmitter.GetTickSize(lDummyOrder)
                mFutureWaiter.Add mTickSizeFuture
            End If
            IsReadyTickSize = False
        Case SecTypeCash
            mTickSize = mContract.TickSize
            IsReadyTickSize = True
        Case SecTypeCombo
            Assert False, "Combo orders not supported"
        Case SecTypeIndex
            mTickSize = mContract.TickSize
            IsReadyTickSize = True
        End Select
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsStopTriggerMethodSupported( _
                ByVal Value As OrderStopTriggerMethods) As Boolean
Const ProcName As String = "IsStopTriggerMethodSupported"
On Error GoTo Err

Assert IsProviderReady, "OrderContext is not ready"
IsStopTriggerMethodSupported = ((mPermittedOrderProperties.PermittedStopTriggerMethods And Value) = Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
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
    
    gSyncToOrder lOrder, pBasedOn
    
    lOrder.BrokerId = ""
    lOrder.ErrorCode = 0
    lOrder.ErrorMessage = ""
    lOrder.LimitPrice = MaxDoubleValue
    lOrder.Message = ""
    lOrder.FillTime = 0
    lOrder.LastFillPrice = 0#
    lOrder.QuantityFilled = 0
    lOrder.QuantityRemaining = lOrder.Quantity
    lOrder.Status = OrderStatusCreated
    lOrder.TriggerPrice = MaxDoubleValue
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
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function NotifyBracketOrderClosed( _
                ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "NotifyBracketOrderClosed"
On Error GoTo Err

If mActiveBracketOrders.Contains(pBracketOrder) Then mActiveBracketOrders.Remove pBracketOrder

If Not mClosedBracketOrders.Contains(pBracketOrder.Key) Then
    mClosedBracketOrders.Add pBracketOrder
    mCompletedSize = mCompletedSize + pBracketOrder.Size
    ' now check if there are any active bracket orders left -
    If mActiveBracketOrders.Count = 0 Then fireChange OrderContextActiveStateChanged
End If
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub NotifyDataReadinessChanged()
Const ProcName As String = "NotifyDataReadinessChanged"
On Error GoTo Err

fireChange OrderContextReadyStateChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ParsePriceAndOffset( _
                ByRef pPriceSpec As PriceSpecifier, _
                ByVal pPriceString As String, _
                Optional ByVal pUseCloseoutSemantics As Boolean = False) As Boolean
Const ProcName As String = "ParsePriceAndOffset"
On Error GoTo Err

ParsePriceAndOffset = gParsePriceAndOffset( _
                                pPriceSpec, _
                                pPriceString, _
                                mContract.Specifier.SecType, _
                                mTickSize, _
                                pUseCloseoutSemantics)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub RemoveChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "RemoveChangeListenerChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Sell( _
    ByVal Number As Long, _
    ByVal pEntryOrder As IOrder, _
    Optional ByVal pStopLossOrder As IOrder, _
    Optional ByVal pTargetOrder As IOrder, _
    Optional ByVal EntryTIF As OrderTIFs = OrderTIFDay, _
    Optional ByVal StopTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
    Optional ByVal TargetTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
    Optional ByVal CancelPrice As Double, _
    Optional ByVal CancelAfter As Long) As IBracketOrder
Const ProcName As String = "Sell"
On Error GoTo Err

'Dim s As String
's = "Sell" & _
'                IIf(mIsSimulated, " (simulated): ", ": ") & mContract.Specifier.LocalSymbol & _
'                "; Number=" & Number & vbCrLf & _
'                "; EntryType=" & gBracketEntryTypeToShortString(EntryType) & _
'                "; EntryLimitPrice=" & EntryLimitPrice & _
'                "; EntryTriggerPrice=" & EntryTriggerPrice & vbCrLf & _
'                "; EntryLimitPriceOffset=" & EntryLimitPriceOffset & _
'                "; EntryLimitPriceOffsetType=" & gPriceOffsetTypeToString(EntryLimitPriceOffsetType) & _
'                "; EntryTriggerPriceOffset=" & EntryTriggerPriceOffset & _
'                "; EntryTriggerPriceOffsetType=" & gPriceOffsetTypeToString(EntryTriggerPriceOffsetType) & vbCrLf & _
'                "; StopLossType=" & gBracketStopLossTypeToShortString(StopLossType) & _
'                "; StopLossLimitPrice=" & StopLossLimitPrice & _
'                "; StopLossTriggerPrice=" & StopLossTriggerPrice & vbCrLf & _
'                "; StopLossLimitPriceOffset=" & StopLossLimitPriceOffset & _
'                "; StopLossLimitPriceOffsetType=" & gPriceOffsetTypeToString(StopLossLimitPriceOffsetType) & _
'                "; StopLossTriggertPriceOffset=" & StopLossTriggerPriceOffset & _
'                "; StopLossTriggerPriceOffsetType=" & gPriceOffsetTypeToString(StopLossTriggerPriceOffsetType) & vbCrLf & _
'                "; TargetType=" & gBracketTargetTypeToShortString(TargetType) & _
'                "; TargetLimitPrice=" & TargetLimitPrice & _
'                "; TargetTriggerPrice=" & TargetTriggerPrice & vbCrLf & _
'                "; TargetLimitPriceOffset=" & TargetLimitPriceOffset & _
'                "; TargetLimitPriceOffsetType=" & gPriceOffsetTypeToString(TargetLimitPriceOffsetType)
's = s & _
'                "; TargetTriggerPriceOffset=" & TargetTriggerPriceOffset & _
'                "; TargetTriggerPriceOffsetType=" & gPriceOffsetTypeToString(TargetTriggerPriceOffsetType) & vbCrLf & _
'                "; EntryTIF=" & gOrderTIFToString(EntryTIF) & _
'                "; StopTIF=" & gOrderTIFToString(StopTIF) & _
'                "; TargetTIF=" & gOrderTIFToString(TargetTIF)
'gLog s, ProcName, ModuleName
                
Set Sell = CreateBracketOrder(OrderActions.OrderActionSell, _
                                Number, _
                                pEntryOrder, _
                                pStopLossOrder, _
                                pTargetOrder)

pEntryOrder.TimeInForce = EntryTIF
If Not pStopLossOrder Is Nothing Then pStopLossOrder.TimeInForce = StopTIF
If Not pTargetOrder Is Nothing Then pTargetOrder.TimeInForce = TargetTIF
Sell.CancelPrice = CancelPrice
Sell.CancelAfter = CancelAfter

ExecuteBracketOrder Sell, True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function dataSourceHasPrice() As Boolean
dataSourceHasPrice = (mDataSource.HasCurrentTick(TickTypeBid) Or _
                        mDataSource.HasCurrentTick(TickTypeAsk) Or _
                        mDataSource.HasCurrentTick(TickTypeTrade))
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fixupEntryOrder(pOrder As Order)
Const ProcName As String = "fixupEntryOrder"
On Error GoTo Err

If pOrder Is Nothing Then Exit Sub

Select Case pOrder.OrderType
Case OrderTypeMarket
Case OrderTypeMarketOnOpen
Case OrderTypeMarketOnClose
Case OrderTypeMarketIfTouched
    pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                pOrder.TriggerPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeMarketToLimit
Case OrderTypeLimit
    pOrder.LimitPrice = CalculateOffsettedPrice( _
                                pOrder.LimitPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeLimitOnOpen
    pOrder.LimitPrice = CalculateOffsettedPrice( _
                                pOrder.LimitPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeLimitOnClose
    pOrder.LimitPrice = CalculateOffsettedPrice( _
                                pOrder.LimitPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeLimitIfTouched
    pOrder.LimitPrice = CalculateOffsettedPrice( _
                                pOrder.LimitPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
    pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                pOrder.TriggerPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeStop, _
        OrderTypeTrail
    pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                pOrder.TriggerPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
Case OrderTypeStopLimit, _
        OrderTypeTrailLimit
    pOrder.LimitPrice = CalculateOffsettedPrice( _
                                pOrder.LimitPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
    pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                pOrder.TriggerPriceSpec, _
                                mContract.Specifier.SecType, _
                                pOrder.Action)
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fixupStopLossOrder(pOrder As Order)
Const ProcName As String = "fixupStopLossOrder"
On Error GoTo Err

If pOrder Is Nothing Then Exit Sub

Select Case pOrder.OrderType
Case OrderTypeStop, _
        OrderTypeTrail
    If pOrder.TriggerPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                    pOrder.TriggerPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
Case OrderTypeStopLimit, _
        OrderTypeTrailLimit
    If pOrder.LimitPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.LimitPrice = CalculateOffsettedPrice( _
                                    pOrder.LimitPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
    
    If pOrder.TriggerPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                    pOrder.TriggerPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fixupTargetOrder(pOrder As Order)
Const ProcName As String = "fixupTargetOrder"
On Error GoTo Err

If pOrder Is Nothing Then Exit Sub

Select Case pOrder.OrderType
Case OrderTypeLimit, _
        OrderTypeLimitOnClose, _
        OrderTypeLimitOnOpen
    If pOrder.LimitPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.LimitPrice = CalculateOffsettedPrice( _
                                    pOrder.LimitPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
Case OrderTypeLimitIfTouched
    If pOrder.LimitPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.LimitPrice = CalculateOffsettedPrice( _
                                    pOrder.LimitPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
    If pOrder.TriggerPriceSpec.PriceType <> PriceValueTypeEntry Then
        pOrder.TriggerPrice = CalculateOffsettedPrice( _
                                    pOrder.TriggerPriceSpec, _
                                    mContract.Specifier.SecType, _
                                    pOrder.Action)
    End If
Case OrderTypeMarketIfTouched
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getCurrentPrice(ByVal pTickType As TickTypes) As Double
Const ProcName As String = "getCurrentPrice"
On Error GoTo Err

If isDataSourceRunning Then getCurrentPrice = mDataSource.CurrentTick(pTickType).Price

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getOrderContexts() As OrderContexts
Set getOrderContexts = mOrderContextsRef.Target
End Function

Private Function isDataSourceRunning()
If mDataSource Is Nothing Then
    isDataSourceRunning = False
Else
    isDataSourceRunning = (mDataSource.State = MarketDataSourceStateRunning)
End If
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
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setTimerListFromSimulatedClock(ByVal pClock As Clock)
Const ProcName As String = "setTimerListFromSimulatedClock"
On Error GoTo Err

Assert pClock.IsSimulated, "Not a simulated clock"
Set mTimerList = GetSimulatedTimerList(pClock)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

