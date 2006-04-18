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

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mName As String

Private mTicker As ticker
Private mPositionManager As positionManager

Private mTickSize As Double

Private mDummy As Boolean

Private mActiveOrderPlexes As Collection
    ' this contains all order plexes for which one or more orders are
    ' not yet complete (or not yet cancelled)
    
Private mClosedOrderPlexes As Collection
    ' this contains all order plexes for which all orders have completed (or
    ' been cancelled)
    
Private mCompletedSize As Long
    ' the total size for all completed order plexes (may be non-zero if one
    ' or more completed order plexes has no target or stop orders)
    
'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mActiveOrderPlexes = New Collection
Set mClosedOrderPlexes = New Collection
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderContext terminated"
End Sub
'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get ActiveSize() As Long
Dim orderPlex As orderPlex

For Each orderPlex In mActiveOrderPlexes
    ActiveSize = ActiveSize + orderPlex.size
Next
End Property

Public Property Get Complete() As Boolean
Dim op As orderPlex
For Each op In mActiveOrderPlexes
    If Not op.dummy Then Complete = False: Exit Property
Next
Complete = True
End Property

Public Property Get CompletedSize() As Long

CompletedSize = mCompletedSize

End Property

Friend Property Let dummy(ByVal value As Boolean)
mDummy = value
End Property

Friend Property Get dummy() As Boolean
dummy = mDummy
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Friend Property Get name() As String
name = mName
End Property

Public Property Get pendingSize() As Long
Dim orderPlex As orderPlex
For Each orderPlex In mActiveOrderPlexes
    pendingSize = pendingSize + orderPlex.pendingSize
Next
End Property

Friend Property Let positionManager(ByVal value As positionManager)
Set mPositionManager = value
End Property

Public Property Get risk() As Currency
Dim orderPlex As orderPlex
For Each orderPlex In mActiveOrderPlexes
    risk = risk + orderPlex.risk
Next
End Property

Public Property Get size() As Long
Dim orderPlex As orderPlex

size = mCompletedSize

For Each orderPlex In mActiveOrderPlexes
    size = size + orderPlex.size
Next
End Property

Friend Property Let ticker(ByVal value As ticker)
Set mTicker = value
mTickSize = mTicker.contract.TickSize
End Property

Public Property Get ticker() As ticker
Set ticker = mTicker
End Property

'================================================================================
' Methods
'================================================================================

Public Function Buy(ByRef number As Long, _
                    ByVal entryType As EntryTypes, _
                    ByVal buyPrice As Double, _
                    ByVal buyOffset As Long, _
                    ByVal triggerPrice As Double, _
                    ByVal stopType As StopTypes, _
                    ByVal stopPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Long, _
                    ByVal targetType As TargetTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double, _
                    ByVal cancelPrice As Double, _
                    stopStrategy As StopManagementStrategy, _
                    Optional exitStrategy As PositionManagementStrategy, _
                    Optional scalingStrategy As PositionManagementStrategy, _
                    Optional ByVal entryTIF As OrderTifs = TIFDay, _
                    Optional ByVal stopTIF As OrderTifs = TIFGoodTillCancelled, _
                    Optional ByVal targetTIF As OrderTifs = TIFGoodTillCancelled) As orderPlex

Select Case entryType
Case EntryTypeMarket
Case EntryTypeMarketOnOpen
Case EntryTypeMarketOnClose
Case EntryTypeMarketIfTouched
Case EntryTypeMarketToLimit
Case EntryTypeBid
Case EntryTypeAsk
Case EntryTypeLast
Case EntryTypeLimit
Case EntryTypeLimitOnOpen
Case EntryTypeLimitOnClose
Case EntryTypeLimitIfTouched
Case EntryTypeStop
Case EntryTypeStopLimit
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::buy", _
    "Invalid entry type"
End Select

Select Case stopType
Case StopTypeNone
Case StopTypeStop
Case StopTypeStopLimit
Case StopTypeBid
Case StopTypeAsk
Case StopTypeLast
Case StopTypeAuto
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::buy", _
    "Invalid stop type"
End Select

Select Case targetType
Case TargetTypeNone
Case TargetTypeLimit
Case TargetTypeLimitIfTouched
Case TargetTypeMarketIfTouched
Case TargetTypeBid
Case TargetTypeAsk
Case TargetTypeLast
Case TargetTypeAuto
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::buy", _
    "Invalid target type"
End Select

Select Case entryTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case TIFGoodTillDate
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::buy", _
    "Invalid entry TIF"
End Select

Select Case stopTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case TIFGoodTillDate
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::buy", _
    "Invalid stop TIF"
End Select

Select Case targetTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case TIFGoodTillDate
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::buy", _
    "Invalid target TIF"
End Select

Set Buy = placeOrder(OrderActions.ActionBuy, _
                                True, _
                                number, _
                                entryType, _
                                buyPrice, _
                                buyOffset, _
                                triggerPrice, _
                                entryTIF, _
                                stopType, _
                                stopPrice, _
                                stopOffset, _
                                stopLimitPrice, _
                                stopTIF, _
                                targetType, _
                                targetPrice, _
                                targetOffset, _
                                targetTriggerPrice, _
                                targetTIF, _
                                cancelPrice, _
                                stopStrategy, _
                                exitStrategy, _
                                scalingStrategy)
End Function

Public Function Cancel() As Boolean
Dim orderPlex As orderPlex

For Each orderPlex In mActiveOrderPlexes
    If orderPlex.Cancel(evenIfFilled:=True) Then Cancel = True
Next

End Function

Public Function CancelPending() As Boolean
Dim orderPlex As orderPlex

For Each orderPlex In mActiveOrderPlexes
    If orderPlex.Cancel(evenIfFilled:=False) Then CancelPending = True
Next

End Function

Public Sub CloseOut(ByVal method As String)
Dim orderPlex As orderPlex

For Each orderPlex In mActiveOrderPlexes
    orderPlex.CloseOut method
Next

End Sub

Friend Function finish()
Dim op As orderPlex

Set mTicker = Nothing
Set mPositionManager = Nothing
For Each op In mActiveOrderPlexes
    op.finish
Next
For Each op In mClosedOrderPlexes
    op.finish
Next
End Function

Public Function getAccountBalance() As Double
getAccountBalance = 10000                       '!!!!!!!!!!!!!!!!!!!!!!!!! need to fix this up
End Function

Friend Function orderPlexClosed( _
                ByVal orderPlex As orderPlex)
Dim key As String

key = orderPlex.key
mClosedOrderPlexes.add orderPlex, key
mActiveOrderPlexes.Remove key
mCompletedSize = mCompletedSize + orderPlex.size
' now check if there are any active order plexes left -
' if not, need to tell the Position Manager
If mActiveOrderPlexes.Count = 0 Then
    mPositionManager.OrderContextInactive Me
End If
End Function

Friend Function placeOrder(ByVal action As OrderActions, _
                    ByVal check As Boolean, _
                    ByRef number As Long, _
                    ByVal entryType As EntryTypes, _
                    ByVal price As Double, _
                    ByVal offset As Long, _
                    ByVal triggerPrice As Double, _
                    ByVal entryTIF As OrderTifs, _
                    ByVal stopType As StopTypes, _
                    ByVal stopTriggerPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Double, _
                    ByVal stopTIF As OrderTifs, _
                    ByVal targetType As TargetTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double, _
                    ByVal targetTIF As OrderTifs, _
                    ByVal cancelPrice As Double, _
                    ByVal stopStrategy As StopManagementStrategy, _
                    ByVal exitStrategy As PositionManagementStrategy, _
                    ByVal scalingStrategy As PositionManagementStrategy) As orderPlex
                    
Dim orderPlex As orderPlex
Dim entryOrder As order
Dim stopOrder As order
Dim targetOrder As order
Dim lEntryPrice As Double
Dim lTriggerPrice As Double
Dim lStopPrice As Double
Dim lTargetPrice As Double
Dim msg As String
Dim ocaGroup As String

If number <= 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                            "TradeBuild.OrderContext::placeOrder", _
                            "The quantity must be greater than zero"

ocaGroup = GenerateTextID

Set orderPlex = New orderPlex
orderPlex.key = nextPlexKey
orderPlex.positionManager = mPositionManager
orderPlex.ticker = mTicker
orderPlex.dummy = mDummy

Set entryOrder = orderPlex.createEntryOrder
entryOrder.action = action
entryOrder.Quantity = number
entryOrder.timeInForce = entryTIF
Select Case entryType
Case EntryTypeMarket
'    If price <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The entry price must be zero for this order type"
'    If offset <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The entry offset must be zero for this order type"
'    If triggerPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The trigger price must be zero for this order type"
    
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeMarket

Case EntryTypeMarketOnOpen
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeMarketOnOpen

Case EntryTypeMarketOnClose
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeMarketOnClose

Case EntryTypeMarketIfTouched
    If triggerPrice = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                "TradeBuild.OrderContext::placeOrder", _
                                "The trigger price must not be zero for this order type"
    
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeMarketIfTouched
    entryOrder.triggerPrice = triggerPrice

Case EntryTypeMarketToLimit
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeMarketToLimit

Case EntryTypeBid
'    If price <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The entry price must be zero for this order type"
'    If triggerPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The trigger price must be zero for this order type"
    
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
    entryOrder.limitPrice = mTicker.BidPrice + offset * mTickSize

Case EntryTypeAsk
'    If price <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The entry price must be zero for this order type"
'    If triggerPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The trigger price must be zero for this order type"
    
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
    entryOrder.limitPrice = mTicker.AskPrice + offset * mTickSize

Case EntryTypeLast
'    If price <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The entry price must be zero for this order type"
'    If triggerPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The trigger price must be zero for this order type"
    
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
    entryOrder.limitPrice = mTicker.TradePrice + offset * mTickSize

Case EntryTypeLimit
'    If triggerPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The trigger price must be zero for this order type"
    If price = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                "TradeBuild.OrderContext::placeOrder", _
                                "The entry price must not be zero for this order type"
    
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
    entryOrder.limitPrice = price + offset * mTickSize

Case EntryTypeLimitOnOpen
    If price = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                "TradeBuild.OrderContext::placeOrder", _
                                "The entry price must not be zero for this order type"
    
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimitOnOpen
    entryOrder.limitPrice = price + offset * mTickSize

Case EntryTypeLimitOnClose
    If price = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                "TradeBuild.OrderContext::placeOrder", _
                                "The entry price must not be zero for this order type"
    
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimitOnClose
    entryOrder.limitPrice = price + offset * mTickSize

Case EntryTypeLimitIfTouched
    If price = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                "TradeBuild.OrderContext::placeOrder", _
                                "The entry price must not be zero for this order type"
    If triggerPrice = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                "TradeBuild.OrderContext::placeOrder", _
                                "The trigger price must not be zero for this order type"

    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimitIfTouched
    entryOrder.limitPrice = price + offset * mTickSize
    entryOrder.triggerPrice = triggerPrice

Case EntryTypeStop
'    If price <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The entry price must be zero for this order type"
'    If offset <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                "TradeBuild.OrderContext::placeOrder", _
'                                "The entry offset must be zero for this order type"
    If triggerPrice = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                "TradeBuild.OrderContext::placeOrder", _
                                "The trigger price must not be zero for this order type"
    
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeStop
    entryOrder.triggerPrice = triggerPrice

Case EntryTypeStopLimit
    If triggerPrice = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                "TradeBuild.OrderContext::placeOrder", _
                                "The trigger price must not be zero for this order type"
    
    entryOrder.orderType = TradeBuild.OrderTypes.OrderTypeStopLimit
    entryOrder.limitPrice = price + offset * mTickSize
    entryOrder.triggerPrice = triggerPrice

End Select

entryOrder.offset = offset
lEntryPrice = entryOrder.limitPrice
lTriggerPrice = entryOrder.triggerPrice

If stopType <> StopTypeNone Then
    Set stopOrder = orderPlex.createStopOrder
    stopOrder.action = IIf(action = ActionBuy, TradeBuild.OrderActions.ActionSell, TradeBuild.OrderActions.ActionBuy)
    stopOrder.Quantity = number
    stopOrder.timeInForce = stopTIF
    
    Select Case stopType
    Case StopTypeStop
        If stopTriggerPrice = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                    "TradeBuild.OrderContext::placeOrder", _
                                    "The stopTrigger price must not be zero for this order type"
        
        stopOrder.ocaGroup = ocaGroup
        stopOrder.orderType = TradeBuild.OrderTypes.OrderTypeStop
        stopOrder.triggerPrice = stopTriggerPrice + stopOffset * mTickSize
    
    Case StopTypeStopLimit
        If stopTriggerPrice = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                    "TradeBuild.OrderContext::placeOrder", _
                                    "The stopTrigger price must not be zero for this order type"
        If stopLimitPrice = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                    "TradeBuild.OrderContext::placeOrder", _
                                    "The stopLimit price must not be zero for this order type"
        
        stopOrder.ocaGroup = ocaGroup
        stopOrder.orderType = TradeBuild.OrderTypes.OrderTypeStopLimit
        stopOrder.triggerPrice = stopTriggerPrice + stopOffset * mTickSize
        stopOrder.limitPrice = stopLimitPrice
    
    Case StopTypeBid
'        If stopTriggerPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                    "TradeBuild.OrderContext::placeOrder", _
'                                    "The stopTrigger price must be zero for this order type"
        
        stopOrder.ocaGroup = ocaGroup
        stopOrder.orderType = TradeBuild.OrderTypes.OrderTypeStop
        stopOrder.triggerPrice = mTicker.BidPrice + stopOffset * mTickSize
        
    Case StopTypeAsk
'        If stopTriggerPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                    "TradeBuild.OrderContext::placeOrder", _
'                                    "The stopTrigger price must be zero for this order type"
        
        stopOrder.ocaGroup = ocaGroup
        stopOrder.orderType = TradeBuild.OrderTypes.OrderTypeStop
        stopOrder.triggerPrice = mTicker.AskPrice + stopOffset * mTickSize
    
    Case StopTypeLast
'        If stopTriggerPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                    "TradeBuild.OrderContext::placeOrder", _
'                                    "The stopTrigger price must be zero for this order type"
        
        stopOrder.ocaGroup = ocaGroup
        stopOrder.orderType = TradeBuild.OrderTypes.OrderTypeStop
        stopOrder.triggerPrice = mTicker.TradePrice + stopOffset * mTickSize

    Case StopTypeAuto
'        If stopTriggerPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                    "TradeBuild.OrderContext::placeOrder", _
'                                    "The stopTrigger price must be zero for this order type"
        
        stopOrder.orderType = TradeBuild.OrderTypes.OrderTypeAutoStop
        stopOrder.offset = stopOffset
    
    End Select
    lStopPrice = stopOrder.triggerPrice
    stopOrder.offset = stopOffset
End If

If targetType <> TargetTypeNone Then
    Set targetOrder = orderPlex.createTargetOrder
    targetOrder.action = IIf(action = ActionBuy, TradeBuild.OrderActions.ActionSell, TradeBuild.OrderActions.ActionBuy)
    targetOrder.Quantity = number
    targetOrder.timeInForce = targetTIF
    
    Select Case targetType
    Case TargetTypeLimit
        If targetPrice = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                    "TradeBuild.OrderContext::placeOrder", _
                                    "The target price must not be zero for this order type"
        
        targetOrder.ocaGroup = ocaGroup
        targetOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
        targetOrder.limitPrice = targetPrice + targetOffset * mTickSize
    
    Case TargetTypeLimitIfTouched
        If targetPrice = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                    "TradeBuild.OrderContext::placeOrder", _
                                    "The target price must not be zero for this order type"
        If targetTriggerPrice = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                    "TradeBuild.OrderContext::placeOrder", _
                                    "The target trigger price must not be zero for this order type"
    
        targetOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimitIfTouched
        targetOrder.limitPrice = targetPrice + offset * mTickSize
        targetOrder.triggerPrice = targetTriggerPrice
    
    Case TargetTypeMarketIfTouched
        If targetTriggerPrice = 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                    "TradeBuild.OrderContext::placeOrder", _
                                    "The target trigger price must not be zero for this order type"
        
        targetOrder.orderType = TradeBuild.OrderTypes.OrderTypeMarketIfTouched
        targetOrder.triggerPrice = targetTriggerPrice
    
    Case TargetTypeBid
'        If targetPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                    "TradeBuild.OrderContext::placeOrder", _
'                                    "The target price must be zero for this order type"
        
        targetOrder.ocaGroup = ocaGroup
        targetOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
        targetOrder.limitPrice = mTicker.BidPrice + targetOffset * mTickSize
    
    Case TargetTypeAsk
'        If targetPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
'                                    "TradeBuild.OrderContext::placeOrder", _
'                                    "The target price must be zero for this order type"
        
        targetOrder.ocaGroup = ocaGroup
        targetOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
        targetOrder.limitPrice = mTicker.AskPrice + targetOffset * mTickSize
    
    Case TargetTypeLast
        targetOrder.ocaGroup = ocaGroup
        targetOrder.orderType = TradeBuild.OrderTypes.OrderTypeLimit
        targetOrder.limitPrice = mTicker.TradePrice + targetOffset * mTickSize
    
    Case TargetTypeAuto
        If targetPrice <> 0 Then err.Raise ErrorCodes.IllegalArgumentException, _
                                    "TradeBuild.OrderContext::placeOrder", _
                                    "The target price must be zero for this order type"
        
        
        targetOrder.orderType = TradeBuild.OrderTypes.OrderTypeAutoLimit
        targetOrder.offset = targetOffset
    
    End Select

    targetOrder.offset = targetOffset
    lTargetPrice = targetOrder.limitPrice
End If

If check And Not mDummy Then
    If Not mPositionManager.allowOrder( _
                            IIf(action = ActionBuy, True, False), _
                            number, _
                            orderPlex.risk, _
                            msg) _
    Then
        mTicker.notifyListeners "Rejected: " & msg, _
                            TradeBuildListenValueTypes.VTOrder
        Set orderPlex = Nothing
        Exit Function
    End If
End If

If Not stopStrategy Is Nothing Then
    stopStrategy.ticker = mTicker
    stopStrategy.orderPlex = orderPlex
    orderPlex.stopStrategy = stopStrategy
End If

If Not exitStrategy Is Nothing Then exitStrategy.orderPlex = orderPlex

If Not scalingStrategy Is Nothing Then
    scalingStrategy.ticker = mTicker
    scalingStrategy.orderPlex = orderPlex
    scalingStrategy.OrderContext = Me
    orderPlex.scalingStrategy = scalingStrategy
End If

mTicker.notifyListeners _
            "Create order plex: (" & mTicker.TradePrice & ") " & _
            action & " " & number & " at " & _
            orderTypeToString(entryOrder.orderType) & " " & lEntryPrice & _
            IIf(lTriggerPrice <> 0, "(" & lTriggerPrice & ")", "") & _
            ", stop " & IIf(stopType = StopTypeAuto, "AUTO(" & stopOffset * mTickSize & ")", lStopPrice) & _
            ", target " & IIf(targetType = TargetTypeAuto, "AUTO(" & targetOffset * mTickSize & ")", lTargetPrice), _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder)

mActiveOrderPlexes.add orderPlex, orderPlex.key
orderPlex.OrderContext = Me
orderPlex.execute cancelPrice
Set placeOrder = orderPlex

End Function

Public Function Sell(ByRef number As Long, _
                    ByVal entryType As EntryTypes, _
                    ByVal sellPrice As Double, _
                    ByVal sellOffset As Long, _
                    ByVal triggerPrice As Double, _
                    ByVal stopType As StopTypes, _
                    ByVal stopPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Long, _
                    ByVal targetType As TargetTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double, _
                    ByVal cancelPrice As Double, _
                    stopStrategy As StopManagementStrategy, _
                    Optional exitStrategy As PositionManagementStrategy, _
                    Optional scalingStrategy As PositionManagementStrategy, _
                    Optional ByVal entryTIF As OrderTifs = TIFDay, _
                    Optional ByVal stopTIF As OrderTifs = TIFGoodTillCancelled, _
                    Optional ByVal targetTIF As OrderTifs = TIFGoodTillCancelled) As orderPlex

Select Case entryType
Case EntryTypeMarket
Case EntryTypeMarketOnOpen
Case EntryTypeMarketOnClose
Case EntryTypeMarketIfTouched
Case EntryTypeMarketToLimit
Case EntryTypeBid
Case EntryTypeAsk
Case EntryTypeLast
Case EntryTypeLimit
Case EntryTypeLimitOnOpen
Case EntryTypeLimitOnClose
Case EntryTypeLimitIfTouched
Case EntryTypeStop
Case EntryTypeStopLimit
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::sell", _
    "Invalid entry type"
End Select

Select Case stopType
Case StopTypeNone
Case StopTypeStop
Case StopTypeStopLimit
Case StopTypeBid
Case StopTypeAsk
Case StopTypeLast
Case StopTypeAuto
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::sell", _
    "Invalid stop type"
End Select

Select Case targetType
Case TargetTypeNone
Case TargetTypeLimit
Case TargetTypeLimitIfTouched
Case TargetTypeMarketIfTouched
Case TargetTypeBid
Case TargetTypeAsk
Case TargetTypeLast
Case TargetTypeAuto
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::sell", _
    "Invalid target type"
End Select

Select Case entryTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case TIFGoodTillDate
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::sell", _
    "Invalid entry TIF"
End Select

Select Case stopTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case TIFGoodTillDate
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::sell", _
    "Invalid stop TIF"
End Select

Select Case targetTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case TIFGoodTillDate
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
    "TradeBuild.OrderContext::sell", _
    "Invalid target TIF"
End Select

Set Sell = placeOrder(OrderActions.ActionSell, _
                                True, _
                                number, _
                                entryType, _
                                sellPrice, _
                                sellOffset, _
                                triggerPrice, _
                                entryTIF, _
                                stopType, _
                                stopPrice, _
                                stopOffset, _
                                stopLimitPrice, _
                                stopTIF, _
                                targetType, _
                                targetPrice, _
                                targetOffset, _
                                targetTriggerPrice, _
                                targetTIF, _
                                cancelPrice, _
                                stopStrategy, _
                                exitStrategy, _
                                scalingStrategy)
End Function

Friend Sub terminate()
Dim op As orderPlex
For Each op In mActiveOrderPlexes
    op.terminate
Next
Set mActiveOrderPlexes = Nothing
For Each op In mClosedOrderPlexes
    op.terminate
Next
Set mClosedOrderPlexes = Nothing
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function nextPlexKey() As String
Static currID As Long
currID = currID + 1
nextPlexKey = mName & currID
End Function





