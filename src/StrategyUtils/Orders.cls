VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Orders"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

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

Private Const ModuleName                            As String = "Orders"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get AccountBalance() As Double
' dummy value at present, awaiting development of an Account Service Provider
AccountBalance = 10000
End Property

Public Property Get ActiveBuySize() As Long
Const ProcName As String = "ActiveBuySize"
On Error GoTo Err

ActiveBuySize = getDefaultOrderContext(True, False).ActiveSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ActiveSellSize() As Long
Const ProcName As String = "ActiveSellSize"
On Error GoTo Err

ActiveSellSize = getDefaultOrderContext(False, False).ActiveSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ActiveSimulatedBuySize() As Long
Const ProcName As String = "ActiveSimulatedBuySize"
On Error GoTo Err

ActiveSimulatedBuySize = getDefaultOrderContext(True, True).ActiveSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ActiveSimulatedSellSize() As Long
Const ProcName As String = "ActiveSimulatedSellSize"
On Error GoTo Err

ActiveSimulatedSellSize = getDefaultOrderContext(False, True).ActiveSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ActiveSize(ByVal OrderContext As ResourceIdentifier) As Long
Const ProcName As String = "ActiveSize"
On Error GoTo Err

ActiveSize = getOrderContext(OrderContext).ActiveSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get AveragePrice(Optional ByVal BracketOrder As ResourceIdentifier) As Double
Const ProcName As String = "AveragePrice"
On Error GoTo Err

AveragePrice = getBracketOrder(BracketOrder).EntryOrder.AveragePrice

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get EntryPrice(Optional ByVal BracketOrder As ResourceIdentifier) As Double
Const ProcName As String = "EntryPrice"
On Error GoTo Err

EntryPrice = getBracketOrder(BracketOrder).EntryOrder.LimitPrice

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get EntryOffset(Optional ByVal BracketOrder As ResourceIdentifier) As Long
Const ProcName As String = "EntryOffset"
On Error GoTo Err

EntryOffset = getBracketOrder(BracketOrder).EntryOrder.Offset

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get EntryTriggerPrice(Optional ByVal BracketOrder As ResourceIdentifier) As Double
Const ProcName As String = "EntryTriggerPrice"
On Error GoTo Err

EntryTriggerPrice = getBracketOrder(BracketOrder).EntryOrder.TriggerPrice

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsBuy(Optional ByVal BracketOrder As ResourceIdentifier) As Boolean
Const ProcName As String = "IsBuy"
On Error GoTo Err

IsBuy = getBracketOrder(BracketOrder).LongPosition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsCompleted(Optional ByVal BracketOrder As ResourceIdentifier) As Boolean
Const ProcName As String = "IsCompleted"
On Error GoTo Err

IsCompleted = (getBracketOrder(BracketOrder).State = BracketOrderStateClosed)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsSell(Optional ByVal BracketOrder As ResourceIdentifier) As Boolean
Const ProcName As String = "IsSell"
On Error GoTo Err

IsSell = Not getBracketOrder(BracketOrder).LongPosition

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PendingBuySize() As Long
Const ProcName As String = "PendingBuySize"
On Error GoTo Err

PendingBuySize = getDefaultOrderContext(True, False).PendingSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PendingSellSize() As Long
Const ProcName As String = "PendingSellSize"
On Error GoTo Err

PendingSellSize = getDefaultOrderContext(False, False).PendingSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PendingSimulatedBuySize() As Long
Const ProcName As String = "PendingSimulatedBuySize"
On Error GoTo Err

PendingSimulatedBuySize = getDefaultOrderContext(True, True).PendingSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PendingSimulatedSellSize() As Long
Const ProcName As String = "PendingSimulatedSellSize"
On Error GoTo Err

PendingSimulatedSellSize = getDefaultOrderContext(False, True).PendingSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PendingSize(ByVal OrderContext As ResourceIdentifier) As Long
Const ProcName As String = "PendingSize"
On Error GoTo Err

PendingSize = getOrderContext(OrderContext).PendingSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PrimaryBracketOrder() As ResourceIdentifier
Const ProcName As String = "PrimaryBracketOrder"
On Error GoTo Err

Set PrimaryBracketOrder = gResourceContext.PrimaryBracketOrder

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get QuantityFilled(Optional ByVal BracketOrder As ResourceIdentifier) As Long
Const ProcName As String = "QuantityFilled"
On Error GoTo Err

QuantityFilled = getBracketOrder(BracketOrder).EntryOrder.QuantityFilled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Size(ByVal OrderContext As ResourceIdentifier) As Long
Const ProcName As String = "Size"
On Error GoTo Err

Size = getOrderContext(OrderContext).Size

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StopLossPrice(Optional ByVal BracketOrder As ResourceIdentifier) As Double
Const ProcName As String = "StopLossPrice"
On Error GoTo Err

StopLossPrice = getBracketOrder(BracketOrder).StopLossOrder.LimitPrice

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StopLossOffset(Optional ByVal BracketOrder As ResourceIdentifier) As Long
Const ProcName As String = "StopLossOffset"
On Error GoTo Err

StopLossOffset = getBracketOrder(BracketOrder).StopLossOrder.Offset

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StopLossTriggerPrice(Optional ByVal BracketOrder As ResourceIdentifier) As Double
Const ProcName As String = "StopLossTriggerPrice"
On Error GoTo Err

StopLossTriggerPrice = getBracketOrder(BracketOrder).StopLossOrder.TriggerPrice

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TargetPrice(Optional ByVal BracketOrder As ResourceIdentifier) As Double
Const ProcName As String = "TargetPrice"
On Error GoTo Err

TargetPrice = getBracketOrder(BracketOrder).TargetOrder.LimitPrice

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TargetOffset(Optional ByVal BracketOrder As ResourceIdentifier) As Long
Const ProcName As String = "TargetOffset"
On Error GoTo Err

TargetOffset = getBracketOrder(BracketOrder).TargetOrder.Offset

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TargetTriggerPrice(Optional ByVal BracketOrder As ResourceIdentifier) As Double
Const ProcName As String = "TargetTriggerPrice"
On Error GoTo Err

TargetTriggerPrice = getBracketOrder(BracketOrder).TargetOrder.TriggerPrice

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TotalPendingPositionSize() As Long
Const ProcName As String = "TotalPendingPositionSize"
On Error GoTo Err

TotalPendingPositionSize = gTradingContext.PositionManager.PendingPositionSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TotalPositionSize() As Long
Const ProcName As String = "TotalPositionSize"
On Error GoTo Err

TotalPositionSize = gTradingContext.PositionManager.PositionSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TotalSimulatedPendingPositionSize() As Long
Const ProcName As String = "TotalSimulatedPendingPositionSize"
On Error GoTo Err

TotalSimulatedPendingPositionSize = gTradingContext.PositionManagerSimulated.PendingPositionSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TotalSimulatedPositionSize() As Long
Const ProcName As String = "TotalSimulatedPositionSize"
On Error GoTo Err

TotalSimulatedPositionSize = gTradingContext.PositionManagerSimulated.PositionSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AdjustStop( _
                ByVal StopTriggerPrice As Double, _
                Optional ByVal Quantity As Long = -1, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "AdjustStop"
On Error GoTo Err

getBracketOrder(BracketOrder).AdjustStop StopTriggerPrice, Quantity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Buy( _
                ByVal Number As Long, _
                ByVal EntryOrder As ResourceIdentifier, _
                Optional ByVal StopLossOrder As ResourceIdentifier, _
                Optional ByVal TargetOrder As ResourceIdentifier, _
                Optional ByVal CancelPrice As Double, _
                Optional ByVal CancelAfter As Long, _
                Optional ByVal NotifyCompletion As Boolean, _
                Optional ByVal OrderContext As ResourceIdentifier) As ResourceIdentifier
Const ProcName As String = "Buy"
On Error GoTo Err

Set Buy = buyOrSell(True, OrderContext, False, Number, EntryOrder, StopLossOrder, TargetOrder, CancelPrice, CancelAfter, NotifyCompletion)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function BuySimulated( _
                ByVal Number As Long, _
                ByVal EntryOrder As ResourceIdentifier, _
                Optional ByVal StopLossOrder As ResourceIdentifier, _
                Optional ByVal TargetOrder As ResourceIdentifier, _
                Optional ByVal CancelPrice As Double, _
                Optional ByVal CancelAfter As Long, _
                Optional ByVal NotifyCompletion As Boolean, _
                Optional ByVal OrderContext As ResourceIdentifier) As ResourceIdentifier
Const ProcName As String = "BuySimulated"
On Error GoTo Err

Set BuySimulated = buyOrSell(True, OrderContext, True, Number, EntryOrder, StopLossOrder, TargetOrder, CancelPrice, CancelAfter, NotifyCompletion)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub CancelBracketOrder( _
                Optional ByVal EvenIfFilled As Boolean, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "CancelBracketOrder"
On Error GoTo Err

getBracketOrder(BracketOrder).Cancel EvenIfFilled

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function CanTrade( _
                Optional ByVal OrderContext As ResourceIdentifier) As Boolean
Const ProcName As String = "CanTrade"
On Error GoTo Err

If OrderContext Is Nothing Then
    CanTrade = canTradeContext(Nothing, True, False) And canTradeContext(Nothing, False, False)
Else
    ' note the value of pIsBuy in this call is irrelevant
    CanTrade = canTradeContext(OrderContext, False, False)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CanTradeSimulated( _
                Optional ByVal OrderContext As ResourceIdentifier) As Boolean
Const ProcName As String = "CanTradeSimulated"
On Error GoTo Err

If OrderContext Is Nothing Then
    CanTradeSimulated = canTradeContext(Nothing, True, True) And canTradeContext(Nothing, False, True)
Else
    ' note the value of pIsBuy in this call is irrelevant
    CanTradeSimulated = canTradeContext(OrderContext, False, True)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub ClearPrimaryBracketOrder()
gResourceContext.ClearPrimaryBracketOrder
End Sub

Public Sub CloseAllPositions()
Const ProcName As String = "CloseAllPositions"
On Error GoTo Err

gTradingContext.PositionManager.ClosePositions

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub CloseAllSimulatedPositions()
Const ProcName As String = "CloseAllSimulatedPositions"
On Error GoTo Err

gTradingContext.PositionManagerSimulated.ClosePositions

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AllowUnprotectedPositions()
Const ProcName As String = "AllowUnprotectedPositions"
On Error GoTo Err

Assert Not gInitialisationContext Is Nothing, "Method can only be called during strategy initialisation"

gInitialisationContext.AllowUnprotectedPositions = True

Dim oc As OrderContext
For Each oc In gInitialisationContext.PositionManager.OrderContexts
    oc.PreventUnprotectedPositions = False
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Sell( _
                ByVal Number As Long, _
                ByVal EntryOrder As ResourceIdentifier, _
                Optional ByVal StopLossOrder As ResourceIdentifier, _
                Optional ByVal TargetOrder As ResourceIdentifier, _
                Optional ByVal CancelPrice As Double, _
                Optional ByVal CancelAfter As Long, _
                Optional ByVal NotifyCompletion As Boolean, _
                Optional ByVal OrderContext As ResourceIdentifier) As ResourceIdentifier
Const ProcName As String = "Sell"
On Error GoTo Err

Set Sell = buyOrSell(False, OrderContext, False, Number, EntryOrder, StopLossOrder, TargetOrder, CancelPrice, CancelAfter, NotifyCompletion)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SellSimulated( _
                ByVal Number As Long, _
                ByVal EntryOrder As ResourceIdentifier, _
                Optional ByVal StopLossOrder As ResourceIdentifier, _
                Optional ByVal TargetOrder As ResourceIdentifier, _
                Optional ByVal CancelPrice As Double, _
                Optional ByVal CancelAfter As Long, _
                Optional ByVal NotifyCompletion As Boolean, _
                Optional ByVal OrderContext As ResourceIdentifier) As ResourceIdentifier
Const ProcName As String = "SellSimulated"
On Error GoTo Err

Set SellSimulated = buyOrSell(False, OrderContext, True, Number, EntryOrder, StopLossOrder, TargetOrder, CancelPrice, CancelAfter, NotifyCompletion)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub SetEntryReason( _
                ByVal Reason As String, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetEntryReason"
On Error GoTo Err

getBracketOrder(BracketOrder).EntryReason = Reason

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewEntryPrice( _
                ByVal Price As Double, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetNewEntryPrice"
On Error GoTo Err

getBracketOrder(BracketOrder).SetNewEntryPrice Price

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewEntryTriggerPrice( _
                ByVal Price As Double, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetNewEntryTriggerPrice"
On Error GoTo Err

getBracketOrder(BracketOrder).SetNewEntryTriggerPrice Price

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewQuantity( _
                ByVal Quantity As Double, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetNewQuantity"
On Error GoTo Err

getBracketOrder(BracketOrder).SetNewQuantity Quantity

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewStopLossOffset( _
                ByVal Offset As Long, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetNewStopLossOffset"
On Error GoTo Err

getBracketOrder(BracketOrder).SetNewStopLossOffset Offset

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewStopLossPrice( _
                ByVal Price As Double, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetNewStopLossPrice"
On Error GoTo Err

getBracketOrder(BracketOrder).SetNewStopLossPrice Price

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewStopLossTriggerPrice( _
                ByVal Price As Double, _
                Optional ByVal BracketOrder As ResourceIdentifier)

Const ProcName As String = "SetNewStopLossTriggerPrice"
On Error GoTo Err

getBracketOrder(BracketOrder).SetNewStopLossTriggerPrice Price

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewTargetOffset( _
                ByVal Offset As Long, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetNewTargetOffset"
On Error GoTo Err

getBracketOrder(BracketOrder).SetNewTargetOffset Offset

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewTargetPrice( _
                ByVal Price As Double, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetNewEntryTriggerPrice"
On Error GoTo Err

getBracketOrder(BracketOrder).SetNewTargetPrice Price

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetNewTargetTriggerPrice( _
                ByVal Price As Double, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetNewTargetTriggerPrice "
On Error GoTo Err

getBracketOrder(BracketOrder).SetNewTargetTriggerPrice Price

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetPrimaryBracketOrder(ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetPrimaryBracketOrder"
On Error GoTo Err

AssertArgument TypeOf BracketOrder.Resource Is IBracketOrder, "BracketOrder is not a BracketOrder object"
gResourceContext.SetPrimaryBracketOrder BracketOrder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetStopReason( _
                ByVal Reason As String, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetStopReason "
On Error GoTo Err

getBracketOrder(BracketOrder).StopReason = Reason

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetTargetReason( _
                ByVal Reason As String, _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "SetTargetReason"
On Error GoTo Err

getBracketOrder(BracketOrder).TargetReason = Reason

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Update( _
                Optional ByVal BracketOrder As ResourceIdentifier)
Const ProcName As String = "Update"
On Error GoTo Err

getBracketOrder(BracketOrder).Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Public Function buyOrSell( _
                ByVal pIsBuy As Boolean, _
                ByVal pOrderContext As ResourceIdentifier, _
                ByVal pIsSimulated As Boolean, _
                ByVal pNumber As Long, _
                ByVal pEntryOrder As ResourceIdentifier, _
                ByVal pStopLossOrder As ResourceIdentifier, _
                ByVal pTargetOrder As ResourceIdentifier, _
                ByVal pCancelPrice As Double, _
                ByVal pCancelAfter As Long, _
                ByVal pNotifyCompletion As Boolean) As ResourceIdentifier
Const ProcName As String = "buyOrSell"
On Error GoTo Err

Assert Not gTradingContext Is Nothing, "Method can only be called during strategy execution"
AssertArgument pNumber > 0, "Number must be greater than 0"
AssertArgument Not pEntryOrder Is Nothing, "pEntryOrder must be supplied"

Dim lOrderContext As OrderContext
If pOrderContext Is Nothing Then
    Set lOrderContext = getDefaultOrderContext(pIsBuy, pIsSimulated)
Else
    Set lOrderContext = getOrderContext(pOrderContext)
    Assert lOrderContext.IsSimulated = pIsSimulated, "Order context has incorrect simulated property"
End If

Dim lEntryOrderSpec As OrderSpecifier
Set lEntryOrderSpec = getOrderSpecifier(pEntryOrder, OrderRoleEntry)

Dim lStopLossOrderSpec As New OrderSpecifier
If Not pStopLossOrder Is Nothing Then Set lStopLossOrderSpec = getOrderSpecifier(pStopLossOrder, OrderRoleStopLoss)

Dim lTargetOrderSpec As New OrderSpecifier
If Not pTargetOrder Is Nothing Then Set lTargetOrderSpec = getOrderSpecifier(pTargetOrder, OrderRoleTarget)

Dim lBracketOrder As IBracketOrder
If pIsBuy Then
    Set lBracketOrder = doBuy( _
                            lOrderContext, _
                            pNumber, _
                            lEntryOrderSpec, _
                            lStopLossOrderSpec, _
                            lTargetOrderSpec, _
                            pCancelPrice, _
                            pCancelAfter)
Else
    Set lBracketOrder = doSell( _
                            lOrderContext, _
                            pNumber, _
                            lEntryOrderSpec, _
                            lStopLossOrderSpec, _
                            lTargetOrderSpec, _
                            pCancelPrice, _
                            pCancelAfter)
End If

Set buyOrSell = gCreateResourceIdentifier(lBracketOrder)
gStrategyRunner.MapBracketOrderToResourceId lBracketOrder, buyOrSell

requestNotification lBracketOrder, pNotifyCompletion
createAttachedStrategies lBracketOrder

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function canTradeContext( _
                ByVal pOrderContext As ResourceIdentifier, _
                ByVal pIsBuy As Boolean, _
                ByVal pIsSimulated As Boolean) As Boolean
Const ProcName As String = "canTradeContext"
On Error GoTo Err

Dim lOrderContext As OrderContext
If pOrderContext Is Nothing Then
    Set lOrderContext = getDefaultOrderContext(pIsBuy, pIsSimulated)
Else
    Set lOrderContext = getOrderContext(pOrderContext)
    Assert lOrderContext.IsSimulated = pIsSimulated, "Order context has incorrect simulated property"
End If

canTradeContext = lOrderContext.IsReady

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub createAttachedStrategies( _
                ByVal pBracketOrder As IBracketOrder)
Const ProcName As String = "createAttachedStrategies"
On Error GoTo Err

If TypeOf gStrategy Is IStrategy Then
    Dim lFactory As IPosnMgmtStrtgyFactory
    For Each lFactory In gStrategyRunner.PositionManagementStrategyFactories
        gTradingContext.ApplyPositionManagementStrategy _
                    pBracketOrder, _
                    lFactory.CreateStrategy(gTradingContext), _
                    gStrategyRunner.GetPositionManagementStrategyResourceContext(lFactory).Clone
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function doBuy( _
                ByVal pOrderContext As OrderContext, _
                ByVal pNumber As Long, _
                ByVal pEntryOrderSpec As OrderSpecifier, _
                ByVal pStopLossOrderSpec As OrderSpecifier, _
                ByVal pTargetOrderSpec As OrderSpecifier, _
                ByVal pCancelPrice As Double, _
                ByVal pCancelAfter As Long) As IBracketOrder
Const ProcName As String = "doBuy"
On Error GoTo Err

Set doBuy = pOrderContext.Buy( _
                        pNumber, _
                        pEntryOrderSpec.OrderType, _
                        pEntryOrderSpec.Price, _
                        pEntryOrderSpec.Offset, _
                        pEntryOrderSpec.TriggerPrice, _
                        pStopLossOrderSpec.OrderType, _
                        pStopLossOrderSpec.TriggerPrice, _
                        pStopLossOrderSpec.Offset, _
                        pStopLossOrderSpec.Price, _
                        pTargetOrderSpec.OrderType, _
                        pTargetOrderSpec.Price, _
                        pTargetOrderSpec.Offset, _
                        pTargetOrderSpec.TriggerPrice, _
                        pEntryOrderSpec.TIF, _
                        pStopLossOrderSpec.TIF, _
                        pTargetOrderSpec.TIF, _
                        pCancelPrice, _
                        pCancelAfter)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function doSell( _
                ByVal pOrderContext As OrderContext, _
                ByVal pNumber As Long, _
                ByVal pEntryOrderSpec As OrderSpecifier, _
                ByVal pStopLossOrderSpec As OrderSpecifier, _
                ByVal pTargetOrderSpec As OrderSpecifier, _
                ByVal pCancelPrice As Double, _
                ByVal pCancelAfter As Long) As IBracketOrder
Const ProcName As String = "doSell"
On Error GoTo Err

Set doSell = pOrderContext.Sell( _
                        pNumber, _
                        pEntryOrderSpec.OrderType, _
                        pEntryOrderSpec.Price, _
                        pEntryOrderSpec.Offset, _
                        pEntryOrderSpec.TriggerPrice, _
                        pStopLossOrderSpec.OrderType, _
                        pStopLossOrderSpec.TriggerPrice, _
                        pStopLossOrderSpec.Offset, _
                        pStopLossOrderSpec.Price, _
                        pTargetOrderSpec.OrderType, _
                        pTargetOrderSpec.Price, _
                        pTargetOrderSpec.Offset, _
                        pTargetOrderSpec.TriggerPrice, _
                        pEntryOrderSpec.TIF, _
                        pStopLossOrderSpec.TIF, _
                        pTargetOrderSpec.TIF, _
                        pCancelPrice, _
                        pCancelAfter)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getBracketOrder( _
                ByVal pBracketOrderIdentifier As ResourceIdentifier) As IBracketOrder
If pBracketOrderIdentifier Is Nothing Then Set pBracketOrderIdentifier = gResourceContext.PrimaryBracketOrder

Dim lObj As Object
Set lObj = pBracketOrderIdentifier.Resource
AssertArgument TypeOf lObj Is IBracketOrder, "ResourceIdentifier does not refer to a bracket order"

Set getBracketOrder = lObj
End Function

Private Function getDefaultOrderContext( _
                ByVal pIsBuy As Boolean, _
                ByVal pIsSimulated As Boolean) As OrderContext
Dim lOrderContext As OrderContext

If pIsBuy Then
    If pIsSimulated Then
        Set lOrderContext = gTradingContext.DefaultBuyOrderContextSimulated
    Else
        Set lOrderContext = gTradingContext.DefaultBuyOrderContext
    End If
Else
    If pIsSimulated Then
        Set lOrderContext = gTradingContext.DefaultSellOrderContextSimulated
    Else
        Set lOrderContext = gTradingContext.DefaultSellOrderContext
    End If
End If

Set getDefaultOrderContext = lOrderContext
End Function

Private Function getOrderContext( _
                ByVal pOrderContextIdentifier As ResourceIdentifier) As OrderContext
AssertArgument Not pOrderContextIdentifier Is Nothing, "OrderContext must be supplied"

Dim lObj As Object
Set lObj = pOrderContextIdentifier.Resource
AssertArgument TypeOf lObj Is OrderContext, "ResourceIdentifier does not refer to an order context"

Set getOrderContext = lObj
End Function

Private Function getOrderSpecifier( _
                ByVal pOrderSpecIdentifier As ResourceIdentifier, _
                ByVal pRole As OrderRoles) As OrderSpecifier
Dim lObj As Object
Set lObj = pOrderSpecIdentifier.Resource
AssertArgument TypeOf lObj Is OrderSpecifier, "ResourceIdentifier does not refer to an order specifier"

Dim lOrderSpec As OrderSpecifier
Set lOrderSpec = lObj
AssertArgument lOrderSpec.OrderRole = pRole, "Order specifier not correct role (entry, stop-loss or target)"

Set getOrderSpecifier = lOrderSpec
End Function

Private Sub requestNotification( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pNotifyCompletion As Boolean)
Const ProcName As String = "requestNotification"
On Error GoTo Err

If pNotifyCompletion Then gStrategyRunner.RequestBracketOrderNotification pBracketOrder, gStrategy, gResourceContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




