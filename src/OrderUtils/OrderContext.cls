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

Private mGroupName                      As String
Private mContextsName                   As String
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

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mActiveBracketOrders = New BracketOrders
Set mClosedBracketOrders = New BracketOrders
Set mFutureWaiter = New FutureWaiter
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
    fireChange OrderContextReadyStateChanged
ElseIf TypeOf ev.Future.Value Is Clock Then
    setTimerListFromSimulatedClock ev.Future.Value
    fireChange OrderContextReadyStateChanged
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

Public Property Get ContextsName() As String
ContextsName = mContextsName
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

Friend Property Get IsProviderReady() As Boolean
IsProviderReady = mProviderIsReady
End Property

Public Property Get IsReady() As Boolean
IsReady = (mDataSourceIsReady And mProviderIsReady And mEnabled And Not mContract Is Nothing And Not mTimerList Is Nothing)
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Get Name() As String
Name = mName
End Property

Public Property Get OrderSubmitter() As IOrderSubmitter
Set OrderSubmitter = mOrderSubmitter
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

If pBracketOrder.IsComplete Then
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
' Creates a buy<code>BracketOrder</code>
'
' @return the <code>BracketOrder</code> object created
'
' @param number         quantity to buy
'       BracketEntryType  the type of order to enter the position:
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
' @param BracketStopLossType  the type of stop-loss order to place:
'           BracketStopLossTypeNone               no stop-loss order
'           BracketStopLossTypeStop               use a stop order
'           BracketStopLossTypeStopLimit          use a stop limit order
'           BracketStopLossTypeBid                use a stop order at the current Bid (see StopOffset)
'           BracketStopLossTypeAsk                use a stop order at the current Ask (see StopOffset)
'           BracketStopLossTypeLast               use a stop order at the last Trade Price (see StopOffset)
'           BracketStopLossTypeAuto               use a stop order at the fill Price for the
'                                           entry order (see StopOffset)
' @param StopPrice      the trigger Price for BracketStopLossTypeStop and BracketStopLossTypeStopLimit
' @param StopOffset     number of ticks to modify the StopPrice. For example, for a stop-loss
'                       order 4 ticks below the entry order fill Price, use BracketStopLossTypeAuto
'                       and set StopOffset to -4
' @param StopLimitPrice the limit Price for BracketStopLossTypeStopLimit
' @param BracketTargetType the type of target order to place:
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
Public Function Buy(ByVal Number As Long, _
                    ByVal BracketEntryType As BracketEntryTypes, _
                    ByVal EntryPrice As Double, _
                    ByVal EntryOffset As Long, _
                    ByVal EntryTriggerPrice As Double, _
                    ByVal BracketStopLossType As BracketStopLossTypes, _
                    ByVal StopLossPrice As Double, _
                    ByVal StopLossOffset As Long, _
                    ByVal StopLossLimitPrice As Double, _
                    ByVal BracketTargetType As BracketTargetTypes, _
                    ByVal targetPrice As Double, _
                    ByVal TargetOffset As Long, _
                    ByVal TargetTriggerPrice As Double, _
                    Optional ByVal EntryTIF As OrderTIFs = OrderTIFDay, _
                    Optional ByVal StopTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
                    Optional ByVal TargetTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
                    Optional ByVal CancelPrice As Double, _
                    Optional ByVal CancelAfter As Long) As IBracketOrder
Const ProcName As String = "Buy"
On Error GoTo Err

gLog "Buy" & IIf(mIsSimulated, " (simulated): ", ": ") & mContract.Specifier.LocalSymbol & _
                " number=" & Number & _
                "; BracketEntryType=" & gBracketEntryTypeToShortString(BracketEntryType) & _
                "; EntryPrice=" & EntryPrice & _
                "; EntryOffset=" & EntryOffset & _
                "; EntryTriggerPrice=" & EntryTriggerPrice & _
                "; BracketStopLossType=" & gBracketStopLossTypeToShortString(BracketStopLossType) & _
                "; StopLossPrice=" & StopLossPrice & _
                "; StopLossOffset=" & StopLossOffset & _
                "; StopLossLimitPrice=" & StopLossLimitPrice & _
                "; BracketTargetType=" & gBracketTargetTypeToShortString(BracketTargetType) & _
                "; TargetPrice=" & targetPrice & _
                "; TargetOffset=" & TargetOffset & _
                "; TargetTriggerPrice=" & TargetTriggerPrice & _
                "; EntryTIF=" & gOrderTIFToString(EntryTIF) & _
                "; StopTIF=" & gOrderTIFToString(StopTIF) & _
                "; TargetTIF=" & gOrderTIFToString(TargetTIF), _
    ProcName, ModuleName
                
Set Buy = CreateBracketOrder(OrderActions.OrderActionBuy, _
                                Number, _
                                BracketEntryType, _
                                EntryPrice, _
                                EntryOffset, _
                                EntryTriggerPrice, _
                                BracketStopLossType, _
                                StopLossPrice, _
                                StopLossOffset, _
                                StopLossLimitPrice, _
                                BracketTargetType, _
                                targetPrice, _
                                TargetOffset, _
                                TargetTriggerPrice)
                                
Buy.EntryOrder.TimeInForce = EntryTIF
If Not Buy.StopLossOrder Is Nothing Then Buy.StopLossOrder.TimeInForce = StopTIF
If Not Buy.TargetOrder Is Nothing Then Buy.TargetOrder.TimeInForce = TargetTIF
Buy.CancelPrice = CancelPrice
Buy.CancelAfter = CancelAfter

ExecuteBracketOrder Buy, True

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

Public Sub CloseOut()
Const ProcName As String = "CloseOut"
On Error GoTo Err

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mActiveBracketOrders
    lBracketOrder.CloseOut
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Public Function CreateBuyBracketOrder( _
                ByVal Number As Long, _
                ByVal BracketEntryType As BracketEntryTypes, _
                ByVal EntryPrice As Double, _
                ByVal EntryOffset As Long, _
                ByVal EntryTriggerPrice As Double, _
                ByVal BracketStopLossType As BracketStopLossTypes, _
                ByVal StopLossPrice As Double, _
                ByVal StopLossOffset As Long, _
                ByVal StopLossLimitPrice As Double, _
                ByVal BracketTargetType As BracketTargetTypes, _
                ByVal targetPrice As Double, _
                ByVal TargetOffset As Long, _
                ByVal TargetTriggerPrice As Double) As IBracketOrder
Const ProcName As String = "CreateBuyBracketOrder"

On Error GoTo Err

Set CreateBuyBracketOrder = CreateBracketOrder(OrderActions.OrderActionBuy, _
                                        Number, _
                                        BracketEntryType, _
                                        EntryPrice, _
                                        EntryOffset, _
                                        EntryTriggerPrice, _
                                        BracketStopLossType, _
                                        StopLossPrice, _
                                        StopLossOffset, _
                                        StopLossLimitPrice, _
                                        BracketTargetType, _
                                        targetPrice, _
                                        TargetOffset, _
                                        TargetTriggerPrice)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function CreateBracketOrder(ByVal pAction As OrderActions, _
                    ByVal pNumber As Long, _
                    ByVal pBracketEntryType As BracketEntryTypes, _
                    ByVal pEntryPrice As Double, _
                    ByVal pEntryOffset As Long, _
                    ByVal pEntryTriggerPrice As Double, _
                    ByVal pBracketStopLossType As BracketStopLossTypes, _
                    ByVal pStopLossTriggerPrice As Double, _
                    ByVal pStopLossOffset As Long, _
                    ByVal pStopLossLimitPrice As Double, _
                    ByVal pBracketTargetType As BracketTargetTypes, _
                    ByVal pTargetPrice As Double, _
                    ByVal pTargetOffset As Long, _
                    ByVal pTargetTriggerPrice As Double) As BracketOrder
Const ProcName As String = "CreateBracketOrder"
On Error GoTo Err

Assert IsReady, "The service provider is not Ready"
Assert pNumber > 0, "The quantity must be greater than zero"

validateBracketEntryType pBracketEntryType
validateBracketStopLossType pBracketStopLossType
validateBracketTargetType pBracketTargetType

Dim lBracketOrder As BracketOrder
Set lBracketOrder = New BracketOrder
lBracketOrder.Initialise nextKey, mContextsName, mGroupName, Me, mContract, mOrderSubmitter, mDataSource, mIsSimulated
lBracketOrder.PreventUnprotectedPosition = mPreventUnprotectedPositions

lBracketOrder.EntryOrder = createEntryOrder(pAction, pNumber, pBracketEntryType, pEntryPrice, pEntryTriggerPrice, pEntryOffset)

If pBracketStopLossType <> BracketStopLossTypeNone Then lBracketOrder.StopLossOrder = createStopLossOrder(IIf(pAction = OrderActionBuy, OrderActions.OrderActionSell, OrderActions.OrderActionBuy), pNumber, pBracketStopLossType, pStopLossLimitPrice, pStopLossTriggerPrice, pStopLossOffset)

If pBracketTargetType <> BracketTargetTypeNone Then lBracketOrder.TargetOrder = createTargetOrder(IIf(pAction = OrderActionBuy, OrderActions.OrderActionSell, OrderActions.OrderActionBuy), pNumber, pBracketTargetType, pTargetPrice, pTargetTriggerPrice, pTargetOffset)

lBracketOrder.Checkpoint

Set CreateBracketOrder = lBracketOrder

gLogBracketOrderMessage "Created bracket order: " & lBracketOrder.ToString, mDataSource, mContract, lBracketOrder.Key, mIsSimulated, Me

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function CreateRawOrder(ByVal Id As String) As IOrder
Const ProcName As String = "CreateRawOrder"
On Error GoTo Err

Assert IsProviderReady, "The service provider is not Ready"

Dim lOrder As New Order
lOrder.Id = Id
Set CreateRawOrder = lOrder

If Not mAccumulatedOrders Is Nothing Then mAccumulatedOrders.Add lOrder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateSellBracketOrder( _
                ByVal Number As Long, _
                ByVal BracketEntryType As BracketEntryTypes, _
                ByVal EntryPrice As Double, _
                ByVal EntryOffset As Long, _
                ByVal EntryTriggerPrice As Double, _
                ByVal BracketStopLossType As BracketStopLossTypes, _
                ByVal StopLossPrice As Double, _
                ByVal StopLossOffset As Long, _
                ByVal StopLossLimitPrice As Double, _
                ByVal BracketTargetType As BracketTargetTypes, _
                ByVal targetPrice As Double, _
                ByVal TargetOffset As Long, _
                ByVal TargetTriggerPrice As Double) As IBracketOrder
Const ProcName As String = "CreateSellBracketOrder"

On Error GoTo Err

Set CreateSellBracketOrder = CreateBracketOrder(OrderActions.OrderActionSell, _
                                    Number, _
                                    BracketEntryType, _
                                    EntryPrice, _
                                    EntryOffset, _
                                    EntryTriggerPrice, _
                                    BracketStopLossType, _
                                    StopLossPrice, _
                                    StopLossOffset, _
                                    StopLossLimitPrice, _
                                    BracketTargetType, _
                                    targetPrice, _
                                    TargetOffset, _
                                    TargetTriggerPrice)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ExecuteBracketOrderEx(ByVal pBracketOrder As IBracketOrder, _
                    ByVal pCheckMoneyManagement As Boolean, _
                    ByVal pAddToActiveBracketOrders As Boolean)
Const ProcName As String = "ExecuteBracketOrderEx"
On Error GoTo Err

AssertArgument Not pBracketOrder Is Nothing, "No bracket order supplied"

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
        pBracketOrder.SetNewQuantity lNewQuantity
        pBracketOrder.Update
    End If
End If

gLogBracketOrderMessage "Execute bracket order: " & pBracketOrder.ToString, mDataSource, mContract, pBracketOrder.Key, mIsSimulated, Me

If pAddToActiveBracketOrders Then
    mActiveBracketOrders.Add pBracketOrder
    If Not mAccumulatedBracketOrders Is Nothing Then mAccumulatedBracketOrders.Add pBracketOrder
End If
fireChange OrderContextActiveStateChanged
pBracketOrder.Execute

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ExecuteBracketOrder(ByVal pBracketOrder As IBracketOrder, _
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
                ByVal pName As String, _
                ByVal pSessionName As String, _
                ByVal pContextsName As String, _
                ByVal pGroupName As String, _
                ByVal pIsSimulated As Boolean, _
                ByVal pContractFuture As IFuture, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pOrderAuthoriser As IOrderAuthoriser, _
                ByVal pAccumulatedBracketOrders As BracketOrders, _
                ByVal pAccumulatedOrders As Orders, _
                ByVal pSimulatedClockFuture As IFuture)
Const ProcName As String = "Initialise"
On Error GoTo Err

mGroupName = pGroupName
mName = pName
mSessionName = pSessionName
mContextsName = pContextsName
mGroupName = pGroupName
mIsSimulated = pIsSimulated

Set mContractFuture = pContractFuture
If mContractFuture.IsAvailable Then
    setContract mContractFuture.Value
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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function IsBracketEntryTypeSupported( _
                ByVal Value As BracketEntryTypes) As Boolean
Const ProcName As String = "IsBracketEntryTypeSupported"
On Error GoTo Err

Assert IsProviderReady, "OrderContext is not ready"

Dim lOrderType As OrderTypes
lOrderType = gBracketEntryTypeToOrderType(Value)
IsBracketEntryTypeSupported = ((mPermittedOrderProperties.PermittedOrderTypes And lOrderType) = lOrderType)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsBracketStopLossTypeSupported( _
                ByVal Value As BracketStopLossTypes) As Boolean
Const ProcName As String = "IsBracketStopLossTypeSupported"
On Error GoTo Err

Assert IsProviderReady, "OrderContext is not ready"

If Value = BracketStopLossTypeAuto Then
    IsBracketStopLossTypeSupported = True
    Exit Function
End If

If Value = BracketStopLossTypeAsk Or Value = BracketStopLossTypeBid Or Value = BracketStopLossTypeLast Then
    If Not isDataSourceRunning Then
        IsBracketStopLossTypeSupported = False
        Exit Function
    End If
End If

Dim lOrderType As OrderTypes
lOrderType = gBracketStopLossTypeToOrderType(Value)
IsBracketStopLossTypeSupported = ((mPermittedOrderProperties.PermittedOrderTypes And lOrderType) = lOrderType)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsBracketTargetTypeSupported( _
                ByVal Value As BracketTargetTypes) As Boolean
Const ProcName As String = "IsBracketTargetTypeSupported"
On Error GoTo Err

Assert IsProviderReady, "OrderContext is not ready"

If Value = BracketTargetTypeAuto Then
    IsBracketTargetTypeSupported = True
    Exit Function
End If

If Value = BracketTargetTypeAsk Or Value = BracketTargetTypeBid Or Value = BracketTargetTypeLast Then
    If Not isDataSourceRunning Then
        IsBracketTargetTypeSupported = False
        Exit Function
    End If
End If

Dim lOrderType As OrderTypes
lOrderType = gBracketTargetTypeToOrderType(Value)
IsBracketTargetTypeSupported = ((mPermittedOrderProperties.PermittedOrderTypes And lOrderType) = lOrderType)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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
If Value = OrderTypeAutoLimit Or _
    Value = OrderTypeAutoStop _
Then
    IsOrderTypeSupported = True
Else
    IsOrderTypeSupported = ((mPermittedOrderProperties.PermittedOrderTypes And Value) = Value)
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

Friend Function NewOrder() As IOrder
Const ProcName As String = "NewOrder"
On Error GoTo Err

Dim lOrder As IOrder
Set lOrder = mTemplateOrder.Clone

lOrder.Id = GenerateGUIDString

If Not mAccumulatedOrders Is Nothing Then mAccumulatedOrders.Add lOrder

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

Friend Sub NotifyPositionCloseout( _
                ByVal pExec As IExecutionReport)
Const ProcName As String = "NotifyPositionCloseout"
On Error GoTo Err

Dim lBracketOrder As IBracketOrder
For Each lBracketOrder In mClosedBracketOrders
    lBracketOrder.NotifyCloseout pExec
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "RemoveChangeListenerChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Sell(ByVal Number As Long, _
                    ByVal BracketEntryType As BracketEntryTypes, _
                    ByVal EntryPrice As Double, _
                    ByVal EntryOffset As Long, _
                    ByVal EntryTriggerPrice As Double, _
                    ByVal BracketStopLossType As BracketStopLossTypes, _
                    ByVal StopLossPrice As Double, _
                    ByVal StopLossOffset As Long, _
                    ByVal StopLossLimitPrice As Double, _
                    ByVal BracketTargetType As BracketTargetTypes, _
                    ByVal targetPrice As Double, _
                    ByVal TargetOffset As Long, _
                    ByVal TargetTriggerPrice As Double, _
                    Optional ByVal EntryTIF As OrderTIFs = OrderTIFDay, _
                    Optional ByVal StopTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
                    Optional ByVal TargetTIF As OrderTIFs = OrderTIFGoodTillCancelled, _
                    Optional ByVal CancelPrice As Double, _
                    Optional ByVal CancelAfter As Long) As IBracketOrder
Const ProcName As String = "Sell"
On Error GoTo Err

gLog "Sell" & IIf(mIsSimulated, " (simulated): ", ": ") & mContract.Specifier.LocalSymbol & _
                " number=" & Number & _
                "; BracketEntryType=" & gBracketEntryTypeToShortString(BracketEntryType) & _
                "; EntryPrice=" & EntryPrice & _
                "; EntryOffset=" & EntryOffset & _
                "; EntryTriggerPrice=" & EntryTriggerPrice & _
                "; BracketStopLossType=" & gBracketStopLossTypeToShortString(BracketStopLossType) & _
                "; StopLossPrice=" & StopLossPrice & _
                "; StopLossOffset=" & StopLossOffset & _
                "; StopLossLimitPrice=" & StopLossLimitPrice & _
                "; BracketTargetType=" & gBracketTargetTypeToShortString(BracketTargetType) & _
                "; TargetPrice=" & targetPrice & _
                "; TargetOffset=" & TargetOffset & _
                "; TargetTriggerPrice=" & TargetTriggerPrice & _
                "; EntryTIF=" & gOrderTIFToString(EntryTIF) & _
                "; EtopTIF=" & gOrderTIFToString(StopTIF) & _
                "; TargetTIF=" & gOrderTIFToString(TargetTIF), _
    ProcName, ModuleName
                
Set Sell = CreateBracketOrder(OrderActions.OrderActionSell, _
                                Number, _
                                BracketEntryType, _
                                EntryPrice, _
                                EntryOffset, _
                                EntryTriggerPrice, _
                                BracketStopLossType, _
                                StopLossPrice, _
                                StopLossOffset, _
                                StopLossLimitPrice, _
                                BracketTargetType, _
                                targetPrice, _
                                TargetOffset, _
                                TargetTriggerPrice)

Sell.EntryOrder.TimeInForce = EntryTIF
If Not Sell.StopLossOrder Is Nothing Then Sell.StopLossOrder.TimeInForce = StopTIF
If Not Sell.TargetOrder Is Nothing Then Sell.TargetOrder.TimeInForce = TargetTIF
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

Private Function createEntryOrder(ByVal pAction As OrderActions, _
                    ByRef pNumber As Long, _
                    ByVal pBracketEntryType As BracketEntryTypes, _
                    ByVal pLimitPrice As Double, _
                    ByVal pTriggerPrice As Double, _
                    ByVal pOffset As Long) As IOrder
Const ProcName As String = "createEntryOrder"
On Error GoTo Err

Dim lOrder As IOrder
Set lOrder = NewOrder

lOrder.Action = pAction
lOrder.Quantity = pNumber

lOrder.OrderType = gBracketEntryTypeToOrderType(pBracketEntryType)

Select Case pBracketEntryType
Case BracketEntryTypeMarket
Case BracketEntryTypeMarketOnOpen
Case BracketEntryTypeMarketOnClose
Case BracketEntryTypeMarketIfTouched
    AssertArgument pTriggerPrice <> 0, "The trigger price must not be zero for this order type"
    
    lOrder.TriggerPrice = pTriggerPrice
Case BracketEntryTypeMarketToLimit
Case BracketEntryTypeBid
    lOrder.LimitPrice = getCurrentPrice(TickTypeBid) + pOffset * mContract.TickSize
Case BracketEntryTypeAsk
    lOrder.LimitPrice = getCurrentPrice(TickTypeAsk) + pOffset * mContract.TickSize
Case BracketEntryTypeLast
    lOrder.LimitPrice = getCurrentPrice(TickTypeTrade) + pOffset * mContract.TickSize
Case BracketEntryTypeLimit
    AssertArgument pLimitPrice <> 0, "The entry price must not be zero for this order type"
    
    lOrder.LimitPrice = pLimitPrice + pOffset * mContract.TickSize
Case BracketEntryTypeLimitOnOpen
    AssertArgument pLimitPrice <> 0, "The entry price must not be zero for this order type"
    
    lOrder.LimitPrice = pLimitPrice + pOffset * mContract.TickSize
Case BracketEntryTypeLimitOnClose
    AssertArgument pLimitPrice <> 0, "The entry price must not be zero for this order type"
    
    lOrder.LimitPrice = pLimitPrice + pOffset * mContract.TickSize
Case BracketEntryTypeLimitIfTouched
    AssertArgument pLimitPrice <> 0, "The entry price must not be zero for this order type"
    AssertArgument pTriggerPrice <> 0, "The trigger price must not be zero for this order type"

    lOrder.LimitPrice = pLimitPrice + pOffset * mContract.TickSize
    lOrder.TriggerPrice = pTriggerPrice
Case BracketEntryTypeStop
    AssertArgument pTriggerPrice <> 0, "The trigger price must not be zero for this order type"
    
    lOrder.TriggerPrice = pTriggerPrice
Case BracketEntryTypeStopLimit
    AssertArgument pTriggerPrice <> 0, "The trigger price must not be zero for this order type"
    
    lOrder.LimitPrice = pLimitPrice + pOffset * mContract.TickSize
    lOrder.TriggerPrice = pTriggerPrice
End Select

lOrder.Offset = pOffset

Set createEntryOrder = lOrder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createStopLossOrder( _
                    ByVal pAction As OrderActions, _
                    ByVal pNumber As Long, _
                    ByVal pBracketStopLossType As BracketStopLossTypes, _
                    ByVal pLimitPrice As Double, _
                    ByVal pTriggerPrice As Double, _
                    ByVal pOffset As Long) As IOrder
Const ProcName As String = "createStopLossOrder"
On Error GoTo Err

Dim lOrder As IOrder
Set lOrder = NewOrder

lOrder.Action = pAction
lOrder.Quantity = pNumber

lOrder.OrderType = gBracketStopLossTypeToOrderType(pBracketStopLossType)
Select Case pBracketStopLossType
Case BracketStopLossTypeStop
    AssertArgument pTriggerPrice <> 0, "The stopTrigger price must not be zero for this order type"
    
    lOrder.TriggerPrice = pTriggerPrice + pOffset * mContract.TickSize
Case BracketStopLossTypeStopLimit
    AssertArgument pTriggerPrice <> 0, "The stopTrigger price must not be zero for this order type"
    AssertArgument pLimitPrice <> 0, "The stopLimit price must not be zero for this order type"
    
    lOrder.TriggerPrice = pTriggerPrice + pOffset * mContract.TickSize
    lOrder.LimitPrice = pLimitPrice
Case BracketStopLossTypeBid
    lOrder.TriggerPrice = getCurrentPrice(TickTypeBid) + pOffset * mContract.TickSize
    
Case BracketStopLossTypeAsk
    lOrder.TriggerPrice = getCurrentPrice(TickTypeAsk) + pOffset * mContract.TickSize
Case BracketStopLossTypeLast
    lOrder.TriggerPrice = getCurrentPrice(TickTypeTrade) + pOffset * mContract.TickSize
Case BracketStopLossTypeAuto
    lOrder.Offset = pOffset
End Select

Set createStopLossOrder = lOrder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createTargetOrder( _
                    ByVal pAction As OrderActions, _
                    ByVal pNumber As Long, _
                    ByVal pBracketTargetType As BracketTargetTypes, _
                    ByVal pLimitPrice As Double, _
                    ByVal pTriggerPrice As Double, _
                    ByVal pOffset As Long) As IOrder
Const ProcName As String = "createTargetOrder"
On Error GoTo Err

Dim lOrder As IOrder
Set lOrder = NewOrder
lOrder.Action = pAction
lOrder.Quantity = pNumber

lOrder.OrderType = gBracketTargetTypeToOrderType(pBracketTargetType)
Select Case pBracketTargetType
Case BracketTargetTypeLimit
    AssertArgument pLimitPrice <> 0, "The target price must not be zero for this order type"
    
    lOrder.LimitPrice = pLimitPrice + pOffset * mContract.TickSize
Case BracketTargetTypeLimitIfTouched
    AssertArgument pLimitPrice <> 0, "The target price must not be zero for this order type"
    AssertArgument pTriggerPrice <> 0, "The target trigger price must not be zero for this order type"

    lOrder.LimitPrice = pLimitPrice + pOffset * mContract.TickSize
    lOrder.TriggerPrice = pTriggerPrice
Case BracketTargetTypeMarketIfTouched
    AssertArgument pTriggerPrice <> 0, "The target trigger price must not be zero for this order type"
    
    lOrder.TriggerPrice = pTriggerPrice
Case BracketTargetTypeBid
    lOrder.LimitPrice = getCurrentPrice(TickTypeBid) + pOffset * mContract.TickSize
Case BracketTargetTypeAsk
    lOrder.LimitPrice = getCurrentPrice(TickTypeAsk) + pOffset * mContract.TickSize
Case BracketTargetTypeLast
    lOrder.LimitPrice = getCurrentPrice(TickTypeTrade) + pOffset * mContract.TickSize
Case BracketTargetTypeAuto
    lOrder.Offset = pOffset
End Select

Set createTargetOrder = lOrder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Function getCurrentPrice(ByVal pTickType As TickTypes) As Double
Const ProcName As String = "getCurrentPrice"
On Error GoTo Err

If isDataSourceRunning Then getCurrentPrice = mDataSource.CurrentTick(pTickType).Price

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isDataSourceRunning()
If mDataSource Is Nothing Then
    isDataSourceRunning = False
Else
    isDataSourceRunning = (mDataSource.State = MarketDataSourceStateRunning)
End If
End Function

Private Function nextKey() As String
Const ProcName As String = "nextKey"
On Error GoTo Err

Static currID As Long
currID = currID + 1
nextKey = mSessionName & "/" & mContextsName & "/" & mName & "/" & currID

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function priceToString(ByVal pPrice As Double) As String
Const ProcName As String = "priceToString"
On Error GoTo Err

priceToString = FormatPrice(pPrice, mContract.Specifier.SecType, mContract.TickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setContract(ByVal pContract As IContract)
Const ProcName As String = "setContract"
On Error GoTo Err

Set mContract = pContract
Set mTemplateOrder = New Order
mTemplateOrder.Initialise mGroupName, mContextsName, mContract.Specifier, Me
mTemplateOrder.IsSimulated = mIsSimulated

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

Private Sub validateBracketEntryType(ByVal pBracketEntryType As BracketEntryTypes)
Select Case pBracketEntryType
Case BracketEntryTypeMarket
Case BracketEntryTypeMarketOnOpen
Case BracketEntryTypeMarketOnClose
Case BracketEntryTypeMarketIfTouched
Case BracketEntryTypeMarketToLimit
Case BracketEntryTypeBid
Case BracketEntryTypeAsk
Case BracketEntryTypeLast
Case BracketEntryTypeLimit
Case BracketEntryTypeLimitOnOpen
Case BracketEntryTypeLimitOnClose
Case BracketEntryTypeLimitIfTouched
Case BracketEntryTypeStop
Case BracketEntryTypeStopLimit
Case Else
    AssertArgument False, "Invalid entry type"
End Select
AssertArgument IsBracketEntryTypeSupported(pBracketEntryType), "Invalid entry type"
End Sub

Private Sub validateBracketStopLossType(ByVal pBracketStopLossType As BracketStopLossTypes)
Select Case pBracketStopLossType
Case BracketStopLossTypeNone
Case BracketStopLossTypeStop
Case BracketStopLossTypeStopLimit
Case BracketStopLossTypeBid
Case BracketStopLossTypeAsk
Case BracketStopLossTypeLast
Case BracketStopLossTypeAuto
Case Else
    AssertArgument False, "Invalid stop-loss type"
End Select
AssertArgument IsBracketStopLossTypeSupported(pBracketStopLossType), "Invalid stop-loss type"
End Sub

Private Sub validateBracketTargetType(ByVal pBracketTargetType As BracketTargetTypes)
Select Case pBracketTargetType
Case BracketTargetTypeNone
Case BracketTargetTypeLimit
Case BracketTargetTypeLimitIfTouched
Case BracketTargetTypeMarketIfTouched
Case BracketTargetTypeBid
Case BracketTargetTypeAsk
Case BracketTargetTypeLast
Case BracketTargetTypeAuto
Case Else
    AssertArgument False, "Invalid target type"
End Select
AssertArgument IsBracketTargetTypeSupported(pBracketTargetType), "Invalid stop-loss type"
End Sub

