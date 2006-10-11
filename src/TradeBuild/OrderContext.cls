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

Event Ready()
Event NotReady()

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

Private mTicker As Ticker
Private mPositionManager As PositionManager

Private mTickSize As Double

Private mDummy As Boolean

Private mSimulated As Boolean

Private mActiveOrderPlexes As orderPlexes
    ' this contains all order plexes for which one or more orders are
    ' not yet complete (or not yet cancelled)
    
Private mClosedOrderPlexes As orderPlexes
    ' this contains all order plexes for which all orders have completed (or
    ' been cancelled)
    
Private mCompletedSize As Long
    ' the total size for all completed order plexes (may be non-zero if one
    ' or more completed order plexes has no target or stop orders)
    
Private mOrderSubmitter As TradeBuildSP.IOrderSubmitter
Private mPermittedOrderAttributes As Long
Private mPermittedOrderTifs As Long
Private mPermittedOrderTypes As Long
Private mPermittedStopTriggerMethods As Long

Private mIsReady As Boolean

Private mOrigin As TradeBuild.OrderOrigins
Private mBlockOrder As Boolean
Private mSweepToFill As Boolean
Private mDisplaySize As Long
Private mIgnoreRegularTradingHours As Boolean
Private mHidden As Boolean
Private mDiscretionaryAmount As Double
Private mGoodAfterTime As Date
Private mGoodAfterTimeTZ    As String
Private mGoodTillDate   As Date
Private mGoodTillDateTZ As String
Private mRegularTradingHoursOnly As Boolean
'Private mRule80A               ' not implemented
Private mSettlingFirm As String
Private mAllOrNone As Boolean
Private mMinimumQuantity As Long
Private mPercentOffset As Double
Private mETradeOnly As Boolean
Private mFirmQuoteOnly  As Boolean
Private mNBBOPriceCap As Double
Private mOverrideConstraints  As Boolean
Private mLimitPrice As Double
Private mOrderType As TradeBuild.OrderTypes
Private mQuantity As Long
Private mTimeInForce As TradeBuild.OrderTifs
Private mTriggerPrice As Double
Private mStopTriggerMethod As TradeBuild.StopTriggerMethods

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mActiveOrderPlexes = New orderPlexes
Set mClosedOrderPlexes = New orderPlexes
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

Public Property Get activeOrderPlexes() As orderPlexes
Set activeOrderPlexes = mActiveOrderPlexes
End Property

Public Property Get ActiveSize() As Long
Dim orderPlex As orderPlex

For Each orderPlex In mActiveOrderPlexes
    ActiveSize = ActiveSize + orderPlex.Size
Next
End Property

Public Property Let allOrNone(ByVal value As Boolean)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttAllOrNone) Then attributeNotSupported
mAllOrNone = value
End Property

Public Property Get allOrNone() As Boolean
allOrNone = mAllOrNone
End Property

Public Property Let blockOrder(ByVal value As Boolean)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttBlockOrder) Then attributeNotSupported
mBlockOrder = value
End Property

Public Property Get blockOrder() As Boolean
blockOrder = mBlockOrder
End Property

Public Property Get closedOrderPlexes() As orderPlexes
Set closedOrderPlexes = mClosedOrderPlexes
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

Public Property Let displaySize(ByVal value As Long)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttDisplaySize) Then attributeNotSupported
mDisplaySize = value
End Property

Public Property Get displaySize() As Long
displaySize = mDisplaySize
End Property

Friend Property Let dummy(ByVal value As Boolean)
mDummy = value
End Property

Friend Property Get dummy() As Boolean
dummy = mDummy
End Property

Public Property Let eTradeOnly(ByVal value As Boolean)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttETradeOnly) Then attributeNotSupported
mETradeOnly = value
End Property

Public Property Get eTradeOnly() As Boolean
eTradeOnly = mETradeOnly
End Property

Public Property Let firmQuoteOnly(ByVal value As Boolean)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttFirmQuoteOnly) Then attributeNotSupported
mFirmQuoteOnly = value
End Property

Public Property Get firmQuoteOnly() As Boolean
firmQuoteOnly = mFirmQuoteOnly
End Property

Public Property Let goodAfterTime(ByVal value As Date)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttGoodAfterTime) Then attributeNotSupported
mGoodAfterTime = value
End Property

Public Property Get goodAfterTime() As Date
goodAfterTime = mGoodAfterTime
End Property

Public Property Let goodAfterTimeTZ(ByVal value As String)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttGoodAfterTimeTZ) Then attributeNotSupported
mGoodAfterTimeTZ = value
End Property

Public Property Get goodAfterTimeTZ() As String
goodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Public Property Let goodTillDate(ByVal value As Date)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttGoodTillDate) Then attributeNotSupported
mGoodTillDate = value
End Property

Public Property Get goodTillDate() As Date
goodTillDate = mGoodTillDate
End Property

Public Property Let goodTillDateTZ(ByVal value As String)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttGoodTillDateTZ) Then attributeNotSupported
mGoodTillDateTZ = value
End Property

Public Property Get goodTillDateTZ() As String
goodTillDateTZ = mGoodTillDateTZ
End Property

Public Property Let hidden(ByVal value As Boolean)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttHidden) Then attributeNotSupported
mHidden = value
End Property

Public Property Get hidden() As Boolean
hidden = mHidden
End Property

Public Property Let ignoreRegularTradingHours(ByVal value As Boolean)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttIgnoreRTH) Then attributeNotSupported
mIgnoreRegularTradingHours = value
End Property

Public Property Get ignoreRegularTradingHours() As Boolean
ignoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Friend Property Let isReady(ByVal value As Boolean)
mIsReady = value
If mIsReady Then
    mPermittedOrderAttributes = mOrderSubmitter.permittedOrderAttributes
    mPermittedOrderTifs = mOrderSubmitter.permittedOrderTifs
    mPermittedOrderTypes = mOrderSubmitter.permittedOrderTypes
    mPermittedStopTriggerMethods = mOrderSubmitter.permittedStopTriggerMethods
    RaiseEvent Ready
Else
    RaiseEvent NotReady
End If
End Property

Public Property Get isReady() As Boolean
isReady = mIsReady
End Property

Public Property Let limitPrice(ByVal value As Double)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttLimitPrice) Then attributeNotSupported
mLimitPrice = value
End Property

Public Property Get limitPrice() As Double
limitPrice = mLimitPrice
End Property

Public Property Let minimumQuantity(ByVal value As Long)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttMinimumQuantity) Then attributeNotSupported
mMinimumQuantity = value
End Property

Public Property Get minimumQuantity() As Long
minimumQuantity = mMinimumQuantity
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Friend Property Get name() As String
name = mName
End Property

Public Property Let NbboPriceCap(ByVal value As Double)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttNBBOPriceCap) Then attributeNotSupported
mNBBOPriceCap = value
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mNBBOPriceCap
End Property

Public Property Let orderType(ByVal value As TradeBuild.OrderTypes)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttOrderType) Then attributeNotSupported
If Not isOrderTypeValueSupported(value) Then attributeValueNotSupported
mOrderType = value
End Property

Public Property Get orderType() As TradeBuild.OrderTypes
orderType = mOrderType
End Property

Public Property Let origin(ByVal value As TradeBuild.OrderOrigins)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttOrigin) Then attributeNotSupported
mOrigin = value
End Property

Public Property Get origin() As TradeBuild.OrderOrigins
origin = mOrigin
End Property

Friend Property Let orderSubmitter( _
                ByVal value As TradeBuildSP.IOrderSubmitter)
Set mOrderSubmitter = value
End Property

Public Property Let overrideConstraints(ByVal value As Boolean)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttOverrideConstraints) Then attributeNotSupported
mOverrideConstraints = value
End Property

Public Property Get overrideConstraints() As Boolean
overrideConstraints = mOverrideConstraints
End Property

Public Property Get pendingSize() As Long
Dim orderPlex As orderPlex
For Each orderPlex In mActiveOrderPlexes
    pendingSize = pendingSize + orderPlex.pendingSize
Next
End Property

Public Property Let percentOffset(ByVal value As Double)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttPercentOffset) Then attributeNotSupported
mPercentOffset = value
End Property

Public Property Get percentOffset() As Double
percentOffset = mPercentOffset
End Property

Public Property Get permittedOrderTifs() As Long
permittedOrderTifs = mPermittedOrderTifs
End Property

Public Property Get permittedOrderTypes() As Long
permittedOrderTypes = mPermittedOrderTypes
End Property

Public Property Get permittedStopTriggerMethods() As Long
permittedStopTriggerMethods = mPermittedStopTriggerMethods
End Property

Friend Property Let PositionManager(ByVal value As PositionManager)
Set mPositionManager = value
End Property

Public Property Let quantity(ByVal value As Long)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttQuantity) Then attributeNotSupported
mQuantity = value
End Property

Public Property Get quantity() As Long
quantity = mQuantity
End Property

Public Property Get risk() As Currency
Dim orderPlex As orderPlex
For Each orderPlex In mActiveOrderPlexes
    risk = risk + orderPlex.risk
Next
End Property

Public Property Let regularTradingHoursOnly(ByVal value As Boolean)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttRTHOnly) Then attributeNotSupported
mRegularTradingHoursOnly = value
End Property

Public Property Get regularTradingHoursOnly() As Boolean
regularTradingHoursOnly = mRegularTradingHoursOnly
End Property

Public Property Let settlingFirm(ByVal value As String)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttSettlingFirm) Then attributeNotSupported
mSettlingFirm = value
End Property

Public Property Get settlingFirm() As String
settlingFirm = mSettlingFirm
End Property

Friend Property Let simulated(ByVal value As Boolean)
mSimulated = value
End Property

Public Property Get simulated() As Boolean
simulated = mSimulated
End Property

Public Property Get Size() As Long
Dim orderPlex As orderPlex

Size = mCompletedSize

For Each orderPlex In mActiveOrderPlexes
    Size = Size + orderPlex.Size
Next
End Property

Public Property Let StopTriggerMethod(ByVal value As TradeBuild.StopTriggerMethods)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttStopTriggerMethod) Then attributeNotSupported
If Not isStopTriggerMethodValueSupported(value) Then attributeValueNotSupported
mStopTriggerMethod = value
End Property

Public Property Get StopTriggerMethod() As TradeBuild.StopTriggerMethods
StopTriggerMethod = mStopTriggerMethod
End Property

Public Property Let SweepToFill(ByVal value As Boolean)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttSweepToFill) Then attributeNotSupported
mSweepToFill = value
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSweepToFill
End Property

Friend Property Let Ticker(ByVal value As Ticker)
If value Is Nothing Then Exit Property
Set mTicker = value
mTickSize = mTicker.Contract.TickSize
End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
End Property

Public Property Let timeInForce(ByVal value As TradeBuild.OrderTifs)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttTimeInForce) Then attributeNotSupported
If Not isOrderTifValueSupported(value) Then attributeValueNotSupported
mTimeInForce = value
End Property

Public Property Get timeInForce() As TradeBuild.OrderTifs
timeInForce = mTimeInForce
End Property

Public Property Let triggerPrice(ByVal value As Double)
If Not isAttributeSupported(TradeBuild.OrderAttributeIds.OrderAttTriggerPrice) Then attributeNotSupported
mTriggerPrice = value
End Property

Public Property Get triggerPrice() As Double
triggerPrice = mTriggerPrice
End Property

'================================================================================
' Methods
'================================================================================

Public Function Buy(ByRef number As Long, _
                    ByVal EntryOrderType As EntryOrderTypes, _
                    ByVal buyPrice As Double, _
                    ByVal buyOffset As Long, _
                    ByVal triggerPrice As Double, _
                    ByVal StopOrderType As StopOrderTypes, _
                    ByVal stopPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Long, _
                    ByVal TargetOrderType As TargetOrderTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double, _
                    Optional ByVal cancelPrice As Double, _
                    Optional ByVal entryTIF As OrderTifs = TIFDay, _
                    Optional ByVal stopTIF As OrderTifs = TIFGoodTillCancelled, _
                    Optional ByVal targetTIF As OrderTifs = TIFGoodTillCancelled, _
                    Optional ByVal stopStrategy As StopManagementStrategy, _
                    Optional ByVal exitStrategy As PositionManagementStrategy, _
                    Optional ByVal scalingStrategy As PositionManagementStrategy) As orderPlex

Select Case entryTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "TradeBuild.OrderContext::buy", _
    "Invalid entry TIF"
End Select

Select Case stopTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "TradeBuild.OrderContext::buy", _
    "Invalid stop TIF"
End Select

Select Case targetTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "TradeBuild.OrderContext::buy", _
    "Invalid target TIF"
End Select

Set Buy = CreateOrderPlex(OrderActions.ActionBuy, _
                                number, _
                                EntryOrderType, _
                                buyPrice, _
                                buyOffset, _
                                triggerPrice, _
                                StopOrderType, _
                                stopPrice, _
                                stopOffset, _
                                stopLimitPrice, _
                                TargetOrderType, _
                                targetPrice, _
                                targetOffset, _
                                targetTriggerPrice)
                                
executeOrderPlex Buy, True, cancelPrice, stopStrategy, exitStrategy, scalingStrategy
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

Public Function CreateBuyOrderPlex( _
                ByRef number As Long, _
                ByVal EntryOrderType As EntryOrderTypes, _
                ByVal buyPrice As Double, _
                ByVal buyOffset As Long, _
                ByVal triggerPrice As Double, _
                ByVal StopOrderType As StopOrderTypes, _
                ByVal stopPrice As Double, _
                ByVal stopOffset As Long, _
                ByVal stopLimitPrice As Long, _
                ByVal TargetOrderType As TargetOrderTypes, _
                ByVal targetPrice As Double, _
                ByVal targetOffset As Long, _
                ByVal targetTriggerPrice As Double) As orderPlex
Set CreateBuyOrderPlex = CreateOrderPlex(OrderActions.ActionBuy, _
                                        number, _
                                        EntryOrderType, _
                                        buyPrice, _
                                        buyOffset, _
                                        triggerPrice, _
                                        StopOrderType, _
                                        stopPrice, _
                                        stopOffset, _
                                        stopLimitPrice, _
                                        TargetOrderType, _
                                        targetPrice, _
                                        targetOffset, _
                                        targetTriggerPrice)
End Function

Friend Function CreateOrderPlex(ByVal action As OrderActions, _
                    ByRef number As Long, _
                    ByVal EntryOrderType As EntryOrderTypes, _
                    ByVal price As Double, _
                    ByVal offset As Long, _
                    ByVal triggerPrice As Double, _
                    ByVal StopOrderType As StopOrderTypes, _
                    ByVal stopTriggerPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Double, _
                    ByVal TargetOrderType As TargetOrderTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double) As orderPlex
                    
Dim orderPlex As orderPlex
Dim entryOrder As order
Dim stoporder As order
Dim targetOrder As order
Dim lEntryPrice As Double
Dim lTriggerPrice As Double
Dim lStopPrice As Double
Dim lTargetPrice As Double

If Not mIsReady Then err.Raise ErrorCodes.ErrIllegalStateException, _
                            "TradeBuild.OrderContext::CreateOrderPlex", _
                            "The service provider is not ready"

If number <= 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                            "TradeBuild.OrderContext::CreateOrderPlex", _
                            "The quantity must be greater than zero"

Select Case EntryOrderType
Case EntryOrderTypeMarket
Case EntryOrderTypeMarketOnOpen
Case EntryOrderTypeMarketOnClose
Case EntryOrderTypeMarketIfTouched
Case EntryOrderTypeMarketToLimit
Case EntryOrderTypeBid
Case EntryOrderTypeAsk
Case EntryOrderTypeLast
Case EntryOrderTypeLimit
Case EntryOrderTypeLimitOnOpen
Case EntryOrderTypeLimitOnClose
Case EntryOrderTypeLimitIfTouched
Case EntryOrderTypeStop
Case EntryOrderTypeStopLimit
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "TradeBuild.OrderContext::CreateOderPlex", _
    "Invalid entry type"
End Select

Select Case StopOrderType
Case StopOrderTypeNone
Case StopOrderTypeStop
Case StopOrderTypeStopLimit
Case StopOrderTypeBid
Case StopOrderTypeAsk
Case StopOrderTypeLast
Case StopOrderTypeAuto
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "TradeBuild.OrderContext::CreateOderPlex", _
    "Invalid stop type"
End Select

Select Case TargetOrderType
Case TargetOrderTypeNone
Case TargetOrderTypeLimit
Case TargetOrderTypeLimitIfTouched
Case TargetOrderTypeMarketIfTouched
Case TargetOrderTypeBid
Case TargetOrderTypeAsk
Case TargetOrderTypeLast
Case TargetOrderTypeAuto
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "TradeBuild.OrderContext::CreateOderPlex", _
    "Invalid target type"
End Select

Set orderPlex = New orderPlex
orderPlex.OrderContext = Me
orderPlex.key = nextPlexKey
orderPlex.PositionManager = mPositionManager
orderPlex.Ticker = mTicker
orderPlex.orderSubmitter = mOrderSubmitter
orderPlex.dummy = mDummy
orderPlex.simulated = mSimulated

Set entryOrder = orderPlex.createEntryOrder
entryOrder.action = action
entryOrder.quantity = number

entryOrder.orderType = gEntryOrderTypeToOrderType(EntryOrderType)
Select Case EntryOrderType
Case EntryOrderTypeMarket
Case EntryOrderTypeMarketOnOpen
Case EntryOrderTypeMarketOnClose
Case EntryOrderTypeMarketIfTouched
    If triggerPrice = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "TradeBuild.OrderContext::CreateOrderPlex", _
                                "The trigger price must not be zero for this order type"
    
    entryOrder.triggerPrice = triggerPrice
Case EntryOrderTypeMarketToLimit
Case EntryOrderTypeBid
    entryOrder.limitPrice = mTicker.BidPrice + offset * mTickSize
Case EntryOrderTypeAsk
    entryOrder.limitPrice = mTicker.AskPrice + offset * mTickSize
Case EntryOrderTypeLast
    entryOrder.limitPrice = mTicker.TradePrice + offset * mTickSize
Case EntryOrderTypeLimit
    If price = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "TradeBuild.OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    
    entryOrder.limitPrice = price + offset * mTickSize
Case EntryOrderTypeLimitOnOpen
    If price = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "TradeBuild.OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    
    entryOrder.limitPrice = price + offset * mTickSize
Case EntryOrderTypeLimitOnClose
    If price = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "TradeBuild.OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    
    entryOrder.limitPrice = price + offset * mTickSize
Case EntryOrderTypeLimitIfTouched
    If price = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "TradeBuild.OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    If triggerPrice = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "TradeBuild.OrderContext::CreateOrderPlex", _
                                "The trigger price must not be zero for this order type"

    entryOrder.limitPrice = price + offset * mTickSize
    entryOrder.triggerPrice = triggerPrice
Case EntryOrderTypeStop
    If triggerPrice = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "TradeBuild.OrderContext::CreateOrderPlex", _
                                "The trigger price must not be zero for this order type"
    
    entryOrder.triggerPrice = triggerPrice
Case EntryOrderTypeStopLimit
    If triggerPrice = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "TradeBuild.OrderContext::CreateOrderPlex", _
                                "The trigger price must not be zero for this order type"
    
    entryOrder.limitPrice = price + offset * mTickSize
    entryOrder.triggerPrice = triggerPrice
End Select

entryOrder.offset = offset
lEntryPrice = entryOrder.limitPrice
lTriggerPrice = entryOrder.triggerPrice

If StopOrderType <> StopOrderTypeNone Then
    Set stoporder = orderPlex.createStopOrder
    stoporder.action = IIf(action = ActionBuy, TradeBuild.OrderActions.ActionSell, TradeBuild.OrderActions.ActionBuy)
    stoporder.quantity = number
    
    stoporder.orderType = gStopOrderTypeToOrderType(StopOrderType)
    Select Case StopOrderType
    Case StopOrderTypeStop
        If stopTriggerPrice = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "TradeBuild.OrderContext::CreateOrderPlex", _
                                    "The stopTrigger price must not be zero for this order type"
        
        stoporder.triggerPrice = stopTriggerPrice + stopOffset * mTickSize
    Case StopOrderTypeStopLimit
        If stopTriggerPrice = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "TradeBuild.OrderContext::CreateOrderPlex", _
                                    "The stopTrigger price must not be zero for this order type"
        If stopLimitPrice = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "TradeBuild.OrderContext::CreateOrderPlex", _
                                    "The stopLimit price must not be zero for this order type"
        
        stoporder.triggerPrice = stopTriggerPrice + stopOffset * mTickSize
        stoporder.limitPrice = stopLimitPrice
    Case StopOrderTypeBid
        stoporder.triggerPrice = mTicker.BidPrice + stopOffset * mTickSize
        
    Case StopOrderTypeAsk
        stoporder.triggerPrice = mTicker.AskPrice + stopOffset * mTickSize
    Case StopOrderTypeLast
        stoporder.triggerPrice = mTicker.TradePrice + stopOffset * mTickSize
    Case StopOrderTypeAuto
        stoporder.offset = stopOffset
    End Select
    lStopPrice = stoporder.triggerPrice
    stoporder.offset = stopOffset
End If

If TargetOrderType <> TargetOrderTypeNone Then
    Set targetOrder = orderPlex.createTargetOrder
    targetOrder.action = IIf(action = ActionBuy, TradeBuild.OrderActions.ActionSell, TradeBuild.OrderActions.ActionBuy)
    targetOrder.quantity = number
    
    targetOrder.orderType = gTargetOrderTypeToOrderType(TargetOrderType)
    Select Case TargetOrderType
    Case TargetOrderTypeLimit
        If targetPrice = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "TradeBuild.OrderContext::CreateOrderPlex", _
                                    "The target price must not be zero for this order type"
        
        targetOrder.limitPrice = targetPrice + targetOffset * mTickSize
    Case TargetOrderTypeLimitIfTouched
        If targetPrice = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "TradeBuild.OrderContext::CreateOrderPlex", _
                                    "The target price must not be zero for this order type"
        If targetTriggerPrice = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "TradeBuild.OrderContext::CreateOrderPlex", _
                                    "The target trigger price must not be zero for this order type"
    
        targetOrder.limitPrice = targetPrice + offset * mTickSize
        targetOrder.triggerPrice = targetTriggerPrice
    Case TargetOrderTypeMarketIfTouched
        If targetTriggerPrice = 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "TradeBuild.OrderContext::CreateOrderPlex", _
                                    "The target trigger price must not be zero for this order type"
        
        targetOrder.triggerPrice = targetTriggerPrice
    Case TargetOrderTypeBid
        targetOrder.limitPrice = mTicker.BidPrice + targetOffset * mTickSize
    Case TargetOrderTypeAsk
        targetOrder.limitPrice = mTicker.AskPrice + targetOffset * mTickSize
    Case TargetOrderTypeLast
        targetOrder.limitPrice = mTicker.TradePrice + targetOffset * mTickSize
    Case TargetOrderTypeAuto
        If targetPrice <> 0 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "TradeBuild.OrderContext::CreateOrderPlex", _
                                    "The target price must be zero for this order type"
        
        targetOrder.offset = targetOffset
    End Select

    targetOrder.offset = targetOffset
    lTargetPrice = targetOrder.limitPrice
End If

orderPlex.checkpoint

Set CreateOrderPlex = orderPlex

mTicker.notifyListeners _
            "Create order plex: (" & mTicker.TradePrice & ") " & _
            orderPlex.key & ": " & _
            gOrderActionToString(action) & " " & number & " at " & _
            gOrderTypeToString(entryOrder.orderType) & " " & lEntryPrice & _
            IIf(lTriggerPrice <> 0, "(" & lTriggerPrice & ")", "") & _
            ", stop " & IIf(StopOrderType = StopOrderTypeAuto, "AUTO(" & stopOffset * mTickSize & ")", lStopPrice) & _
            ", target " & IIf(TargetOrderType = TargetOrderTypeAuto, "AUTO(" & targetOffset * mTickSize & ")", lTargetPrice), _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
            Me

End Function

Public Function CreateSellOrderPlex( _
                ByRef number As Long, _
                ByVal EntryOrderType As EntryOrderTypes, _
                ByVal buyPrice As Double, _
                ByVal buyOffset As Long, _
                ByVal triggerPrice As Double, _
                ByVal StopOrderType As StopOrderTypes, _
                ByVal stopPrice As Double, _
                ByVal stopOffset As Long, _
                ByVal stopLimitPrice As Long, _
                ByVal TargetOrderType As TargetOrderTypes, _
                ByVal targetPrice As Double, _
                ByVal targetOffset As Long, _
                ByVal targetTriggerPrice As Double) As orderPlex
Set CreateSellOrderPlex = CreateOrderPlex(OrderActions.ActionSell, _
                                    number, _
                                    EntryOrderType, _
                                    buyPrice, _
                                    buyOffset, _
                                    triggerPrice, _
                                    StopOrderType, _
                                    stopPrice, _
                                    stopOffset, _
                                    stopLimitPrice, _
                                    TargetOrderType, _
                                    targetPrice, _
                                    targetOffset, _
                                    targetTriggerPrice)
End Function

Public Sub executeOrderPlex(ByVal orderPlex As orderPlex, _
                    Optional checkMoneyManagement As Boolean, _
                    Optional cancelPrice As Double, _
                    Optional stopStrategy As StopManagementStrategy, _
                    Optional exitStrategy As PositionManagementStrategy, _
                    Optional scalingStrategy As PositionManagementStrategy)
                    
Dim msg As String
Dim origQuantity As Long
Dim newQuantity As Long

If orderPlex Is Nothing Then err.Raise ErrorCodes.ErrInvalidObjectReference, _
                            "TradeBuild.OrderContext::executeOrderPlex", _
                            "No order plex supplied"

If checkMoneyManagement And Not mDummy Then
    origQuantity = orderPlex.entryOrder.quantity
    newQuantity = origQuantity
    If Not mPositionManager.allowOrder( _
                            IIf(orderPlex.entryOrder.action = ActionBuy, True, False), _
                            newQuantity, _
                            orderPlex.risk, _
                            msg) _
    Then
        mTicker.notifyListeners "Rejected: " & msg, _
                            TradeBuildListenValueTypes.VTOrder, _
                            Me
        Set orderPlex = Nothing
        Exit Sub
    End If
    If newQuantity <> origQuantity Then
        orderPlex.newQuantity = newQuantity
        orderPlex.update
    End If
End If

If Not stopStrategy Is Nothing Then
    stopStrategy.Ticker = mTicker
    stopStrategy.orderPlex = orderPlex
    orderPlex.stopStrategy = stopStrategy
End If

If Not exitStrategy Is Nothing Then
    exitStrategy.Ticker = mTicker
    exitStrategy.orderPlex = orderPlex
    exitStrategy.OrderContext = Me
    exitStrategy.OrderContext = Me
End If

If Not scalingStrategy Is Nothing Then
    scalingStrategy.Ticker = mTicker
    scalingStrategy.orderPlex = orderPlex
    scalingStrategy.OrderContext = Me
    orderPlex.scalingStrategy = scalingStrategy
End If

mTicker.notifyListeners _
            "Execute order plex: (" & mTicker.TradePrice & ") " & _
            orderPlex.key, _
            IIf(mDummy, TradeBuildListenValueTypes.VTDummyOrder, TradeBuildListenValueTypes.VTOrder), _
            Me

mActiveOrderPlexes.add orderPlex
orderPlex.execute cancelPrice

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
Set mOrderSubmitter = Nothing
End Function

Public Function getAccountBalance() As Double
getAccountBalance = 10000                       '!!!!!!!!!!!!!!!!!!!!!!!!! need to fix this up
End Function

Public Function isAttributeSupported( _
                ByVal attributeId As TradeBuild.OrderAttributeIds) As Boolean
isAttributeSupported = (mPermittedOrderAttributes And attributeId)
End Function

Public Function isOrderTifValueSupported( _
                ByVal value As TradeBuild.OrderTifs) As Boolean
isOrderTifValueSupported = (mPermittedOrderTifs And value)
End Function

Public Function isOrderTypeValueSupported( _
                ByVal value As TradeBuild.OrderTypes) As Boolean
isOrderTypeValueSupported = (mPermittedOrderTypes And value)
End Function

Public Function isStopTriggerMethodValueSupported( _
                ByVal value As TradeBuild.StopTriggerMethods) As Boolean
isStopTriggerMethodValueSupported = (mPermittedStopTriggerMethods And value)
End Function

Friend Function newOrder() As order
Dim lOrder As order
Set lOrder = New order
With lOrder
    .id = gGenerateGUIDString
    .dummy = mDummy
    .simulated = mSimulated
    .OrderContext = Me
    .Ticker = mTicker
    If isAttributeSupported(OrderAttributeIds.OrderAttAllOrNone) Then .allOrNone = mAllOrNone
    If isAttributeSupported(OrderAttributeIds.OrderAttBlockOrder) Then .blockOrder = mBlockOrder
    If isAttributeSupported(OrderAttributeIds.OrderAttDiscretionaryAmount) Then .discretionaryAmount = mDiscretionaryAmount
    If isAttributeSupported(OrderAttributeIds.OrderAttDisplaySize) Then .displaySize = mDisplaySize
    If isAttributeSupported(OrderAttributeIds.OrderAttETradeOnly) Then .eTradeOnly = mETradeOnly
    If isAttributeSupported(OrderAttributeIds.OrderAttFirmQuoteOnly) Then .firmQuoteOnly = mFirmQuoteOnly
    If isAttributeSupported(OrderAttributeIds.OrderAttGoodAfterTime) Then .goodAfterTime = mGoodAfterTime
    If isAttributeSupported(OrderAttributeIds.OrderAttGoodAfterTimeTZ) Then .goodAfterTimeTZ = mGoodAfterTimeTZ
    If isAttributeSupported(OrderAttributeIds.OrderAttGoodTillDate) Then .goodTillDate = mGoodTillDate
    If isAttributeSupported(OrderAttributeIds.OrderAttGoodTillDateTZ) Then .goodTillDateTZ = mGoodTillDateTZ
    If isAttributeSupported(OrderAttributeIds.OrderAttHidden) Then .hidden = mHidden
    If isAttributeSupported(OrderAttributeIds.OrderAttIgnoreRTH) Then .ignoreRegularTradingHours = mIgnoreRegularTradingHours
    If isAttributeSupported(OrderAttributeIds.OrderAttLimitPrice) Then .limitPrice = mLimitPrice
    If isAttributeSupported(OrderAttributeIds.OrderAttMinimumQuantity) Then .minimumQuantity = mMinimumQuantity
    If isAttributeSupported(OrderAttributeIds.OrderAttNBBOPriceCap) Then .NbboPriceCap = mNBBOPriceCap
    If mOrderType <> OrderTypes.OrderTypeNone Then
        If isAttributeSupported(OrderAttributeIds.OrderAttOrderType) Then .orderType = mOrderType
    End If
    If isAttributeSupported(OrderAttributeIds.OrderAttOrigin) Then .origin = mOrigin
    If isAttributeSupported(OrderAttributeIds.OrderAttOverrideConstraints) Then .overrideConstraints = mOverrideConstraints
    If isAttributeSupported(OrderAttributeIds.OrderAttPercentOffset) Then .percentOffset = mPercentOffset
    If isAttributeSupported(OrderAttributeIds.OrderAttQuantity) Then .quantity = mQuantity
    If isAttributeSupported(OrderAttributeIds.OrderAttRTHOnly) Then .regularTradingHoursOnly = mRegularTradingHoursOnly
    If isAttributeSupported(OrderAttributeIds.OrderAttSettlingFirm) Then .settlingFirm = mSettlingFirm
    .simulated = mSimulated
    If mStopTriggerMethod <> StopTriggerMethods.StopTriggerNone Then
        If isAttributeSupported(OrderAttributeIds.OrderAttStopTriggerMethod) Then .StopTriggerMethod = mStopTriggerMethod
    End If
    If isAttributeSupported(OrderAttributeIds.OrderAttSweepToFill) Then .SweepToFill = mSweepToFill
    .Ticker = mTicker
    If mTimeInForce <> OrderTifs.TIFNone Then
        If isAttributeSupported(OrderAttributeIds.OrderAttTimeInForce) Then .timeInForce = mTimeInForce
    End If
    If isAttributeSupported(OrderAttributeIds.OrderAttTriggerPrice) Then .triggerPrice = mTriggerPrice
End With

mPositionManager.Orders.add lOrder

Set newOrder = lOrder
End Function

Friend Function orderPlexClosed( _
                ByVal orderPlex As orderPlex)
mClosedOrderPlexes.add orderPlex
mActiveOrderPlexes.remove orderPlex
mCompletedSize = mCompletedSize + orderPlex.Size
' now check if there are any active order plexes left -
' if not, need to tell the Position Manager
If mActiveOrderPlexes.Count = 0 Then
    mPositionManager.OrderContextInactive Me
End If
End Function

Public Function Sell(ByRef number As Long, _
                    ByVal EntryOrderType As EntryOrderTypes, _
                    ByVal sellPrice As Double, _
                    ByVal sellOffset As Long, _
                    ByVal triggerPrice As Double, _
                    ByVal StopOrderType As StopOrderTypes, _
                    ByVal stopPrice As Double, _
                    ByVal stopOffset As Long, _
                    ByVal stopLimitPrice As Long, _
                    ByVal TargetOrderType As TargetOrderTypes, _
                    ByVal targetPrice As Double, _
                    ByVal targetOffset As Long, _
                    ByVal targetTriggerPrice As Double, _
                    Optional ByVal cancelPrice As Double, _
                    Optional ByVal entryTIF As OrderTifs = TIFDay, _
                    Optional ByVal stopTIF As OrderTifs = TIFGoodTillCancelled, _
                    Optional ByVal targetTIF As OrderTifs = TIFGoodTillCancelled, _
                    Optional ByVal stopStrategy As StopManagementStrategy, _
                    Optional ByVal exitStrategy As PositionManagementStrategy, _
                    Optional ByVal scalingStrategy As PositionManagementStrategy) As orderPlex

Select Case entryTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "TradeBuild.OrderContext::sell", _
    "Invalid entry TIF"
End Select

Select Case stopTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "TradeBuild.OrderContext::sell", _
    "Invalid stop TIF"
End Select

Select Case targetTIF
Case TIFDay
Case TIFGoodTillCancelled
Case TIFImmediateOrCancel
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "TradeBuild.OrderContext::sell", _
    "Invalid target TIF"
End Select

Set Sell = CreateOrderPlex(OrderActions.ActionSell, _
                                number, _
                                EntryOrderType, _
                                sellPrice, _
                                sellOffset, _
                                triggerPrice, _
                                StopOrderType, _
                                stopPrice, _
                                stopOffset, _
                                stopLimitPrice, _
                                TargetOrderType, _
                                targetPrice, _
                                targetOffset, _
                                targetTriggerPrice)
executeOrderPlex Sell, True, cancelPrice, stopStrategy, exitStrategy, scalingStrategy
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub attributeNotSupported()
err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuild.OrderContext", _
            "This attribute not supported by service provider"
End Sub

Private Sub attributeValueNotSupported()
err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "TradeBuild.OrderContext", _
            "This attribute value not supported by service provider"
End Sub

Private Function nextPlexKey() As String
Static currID As Long
currID = currID + 1
nextPlexKey = mTicker.key & "/" & mName & "/" & currID
End Function

