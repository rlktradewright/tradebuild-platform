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

Event ready()
Event NotReady()

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

Private mName As String

Private mTicker As Ticker
Private mPositionManager As PositionManager

Private mTickSize As Double

Private mIsSimulated As Boolean

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

Private mEnabled As Boolean
Private mIsReady As Boolean

Private mOrigin As OrderOrigins
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
Private mOrderType As OrderTypes
Private mQuantity As Long
Private mTimeInForce As OrderTifs
Private mTriggerPrice As Double
Private mStopTriggerMethod As StopTriggerMethods

Private mPreventUnprotectedPositions As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mActiveOrderPlexes = New orderPlexes
Set mClosedOrderPlexes = New orderPlexes
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderContext terminated"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get activeOrderPlexes() As orderPlexes
Set activeOrderPlexes = mActiveOrderPlexes
End Property

Public Property Get ActiveSize() As Long
Dim orderPlex As orderPlex

For Each orderPlex In mActiveOrderPlexes
    ActiveSize = ActiveSize + orderPlex.size
Next
End Property

Public Property Let allOrNone(ByVal value As Boolean)
checkAttributeSupported OrderAttributeIds.OrderAttAllOrNone
mAllOrNone = value
End Property

Public Property Get allOrNone() As Boolean
allOrNone = mAllOrNone
End Property

Public Property Let blockOrder(ByVal value As Boolean)
checkAttributeSupported OrderAttributeIds.OrderAttBlockOrder
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
    If Not op.IsSimulated Then Complete = False: Exit Property
Next
Complete = True
End Property

Public Property Get CompletedSize() As Long
CompletedSize = mCompletedSize
End Property

Public Property Let displaySize(ByVal value As Long)
checkAttributeSupported OrderAttributeIds.OrderAttDisplaySize
mDisplaySize = value
End Property

Public Property Get displaySize() As Long
displaySize = mDisplaySize
End Property

Friend Property Let enabled(ByVal value As Boolean)
mEnabled = value
checkReady
End Property

Public Property Let eTradeOnly(ByVal value As Boolean)
checkAttributeSupported OrderAttributeIds.OrderAttETradeOnly
mETradeOnly = value
End Property

Public Property Get eTradeOnly() As Boolean
eTradeOnly = mETradeOnly
End Property

Public Property Let firmQuoteOnly(ByVal value As Boolean)
checkAttributeSupported OrderAttributeIds.OrderAttFirmQuoteOnly
mFirmQuoteOnly = value
End Property

Public Property Get firmQuoteOnly() As Boolean
firmQuoteOnly = mFirmQuoteOnly
End Property

Public Property Let goodAfterTime(ByVal value As Date)
checkAttributeSupported OrderAttributeIds.OrderAttGoodAfterTime
mGoodAfterTime = value
End Property

Public Property Get goodAfterTime() As Date
goodAfterTime = mGoodAfterTime
End Property

Public Property Let goodAfterTimeTZ(ByVal value As String)
checkAttributeSupported OrderAttributeIds.OrderAttGoodAfterTimeTZ
mGoodAfterTimeTZ = value
End Property

Public Property Get goodAfterTimeTZ() As String
goodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Public Property Let goodTillDate(ByVal value As Date)
checkAttributeSupported OrderAttributeIds.OrderAttGoodTillDate
mGoodTillDate = value
End Property

Public Property Get goodTillDate() As Date
goodTillDate = mGoodTillDate
End Property

Public Property Let goodTillDateTZ(ByVal value As String)
checkAttributeSupported OrderAttributeIds.OrderAttGoodTillDateTZ
mGoodTillDateTZ = value
End Property

Public Property Get goodTillDateTZ() As String
goodTillDateTZ = mGoodTillDateTZ
End Property

Public Property Let hidden(ByVal value As Boolean)
checkAttributeSupported OrderAttributeIds.OrderAttHidden
mHidden = value
End Property

Public Property Get hidden() As Boolean
hidden = mHidden
End Property

Public Property Let ignoreRegularTradingHours(ByVal value As Boolean)
checkAttributeSupported OrderAttributeIds.OrderAttIgnoreRTH
mIgnoreRegularTradingHours = value
End Property

Public Property Get ignoreRegularTradingHours() As Boolean
ignoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Friend Property Let isReady(ByVal value As Boolean)
mIsReady = value
checkReady
End Property

Public Property Get isReady() As Boolean
isReady = (mIsReady And mEnabled)
End Property

Friend Property Let IsSimulated(ByVal value As Boolean)
mIsSimulated = value
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Public Property Let limitPrice(ByVal value As Double)
checkAttributeSupported OrderAttributeIds.OrderAttLimitPrice
mLimitPrice = value
End Property

Public Property Get limitPrice() As Double
limitPrice = mLimitPrice
End Property

Public Property Let minimumQuantity(ByVal value As Long)
checkAttributeSupported OrderAttributeIds.OrderAttMinimumQuantity
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
checkAttributeSupported OrderAttributeIds.OrderAttNBBOPriceCap
mNBBOPriceCap = value
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mNBBOPriceCap
End Property

Public Property Let orderType(ByVal value As OrderTypes)
checkAttributeSupported OrderAttributeIds.OrderAttOrderType
If Not isOrderTypeValueSupported(value) Then attributeValueNotSupported OrderAttributeIds.OrderAttOrderType, value
mOrderType = value
End Property

Public Property Get orderType() As OrderTypes
orderType = mOrderType
End Property

Public Property Let origin(ByVal value As OrderOrigins)
checkAttributeSupported OrderAttributeIds.OrderAttOrigin
mOrigin = value
End Property

Public Property Get origin() As OrderOrigins
origin = mOrigin
End Property

Friend Property Let orderSubmitter( _
                ByVal value As TradeBuildSP.IOrderSubmitter)
Set mOrderSubmitter = value
End Property

Public Property Let overrideConstraints(ByVal value As Boolean)
checkAttributeSupported OrderAttributeIds.OrderAttOverrideConstraints
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
checkAttributeSupported OrderAttributeIds.OrderAttPercentOffset
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

Public Property Let PreventUnprotectedPositions( _
                ByVal value As Boolean)
mPreventUnprotectedPositions = value
End Property

Public Property Get PreventUnprotectedPositions() As Boolean
PreventUnprotectedPositions = mPreventUnprotectedPositions
End Property

Public Property Let quantity(ByVal value As Long)
checkAttributeSupported OrderAttributeIds.OrderAttQuantity
mQuantity = value
End Property

Public Property Get quantity() As Long
quantity = mQuantity
End Property

Public Property Get Risk() As Currency
Dim orderPlex As orderPlex
For Each orderPlex In mActiveOrderPlexes
    Risk = Risk + orderPlex.Risk
Next
End Property

Public Property Let settlingFirm(ByVal value As String)
checkAttributeSupported OrderAttributeIds.OrderAttSettlingFirm
mSettlingFirm = value
End Property

Public Property Get settlingFirm() As String
settlingFirm = mSettlingFirm
End Property

Public Property Get size() As Long
Dim orderPlex As orderPlex

size = mCompletedSize

For Each orderPlex In mActiveOrderPlexes
    size = size + orderPlex.size
Next
End Property

Public Property Let StopTriggerMethod(ByVal value As StopTriggerMethods)
checkAttributeSupported OrderAttributeIds.OrderAttStopTriggerMethod
If Not isStopTriggerMethodValueSupported(value) Then attributeValueNotSupported OrderAttributeIds.OrderAttStopTriggerMethod, value
mStopTriggerMethod = value
End Property

Public Property Get StopTriggerMethod() As StopTriggerMethods
StopTriggerMethod = mStopTriggerMethod
End Property

Public Property Let SweepToFill(ByVal value As Boolean)
checkAttributeSupported OrderAttributeIds.OrderAttSweepToFill
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

Public Property Let timeInForce(ByVal value As OrderTifs)
If Not gIsValidTIF(value) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "timeInForce", _
            "Value must be a member of the OrderTifs enum"
End If

checkAttributeSupported OrderAttributeIds.OrderAttTimeInForce
If Not isOrderTifValueSupported(value) Then attributeValueNotSupported OrderAttributeIds.OrderAttTimeInForce, value
mTimeInForce = value
End Property

Public Property Get timeInForce() As OrderTifs
timeInForce = mTimeInForce
End Property

Public Property Let triggerPrice(ByVal value As Double)
checkAttributeSupported OrderAttributeIds.OrderAttTriggerPrice
mTriggerPrice = value
End Property

Public Property Get triggerPrice() As Double
triggerPrice = mTriggerPrice
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Creates a buy<code>OrderPlex</code>
'
' @return the <code>OrderPlex</code> object created
'
' @param number         quantity to buy
'       EntryOrderType  the type of order to enter the position:
'           EntryOrderTypeMarket            a market order
'           EntryOrderTypeMarketOnOpen      a market on open order
'           EntryOrderTypeMarketOnClose     a market on close order
'           EntryOrderTypeMarketIfTouched   a market if touched order
'           EntryOrderTypeMarketToLimit     a market to limit order
'           EntryOrderTypeBid               a limit order at the current Bid
'           EntryOrderTypeAsk               a limit order at the current Ask
'           EntryOrderTypeLast              a limit order at the last Trade price
'           EntryOrderTypeLimit             a limit order
'           EntryOrderTypeLimitOnOpen       a limit on open order
'           EntryOrderTypeLimitOnClose      a limit on close order
'           EntryOrderTypeLimitIfTouched    a limit if touched order
'           EntryOrderTypeStop              a stop order
'           EntryOrderTypeStopLimit         a stop limit order
' @param BuyPrice       the limit price for a limit order (ignored for EntryOrderTypeBid,
'                       EntryOrderTypeAsk and EntryOrderTypeLast)
' @param BuyOffset      number of ticks to modify the BuyPrice. For example, to place an order
'                       one tick below the current Ask, use EntryOrderTypeAsk and set
'                       BuyOffset to -1
' @param TriggerPrice   used with EntryOrderTypeMarketIfTouched, EntryOrderTypeLimitIfTouched,
'                       EntryOrderTypeStop and EntryOrderTypeStopLimit to specify the price
'                       at which the order is triggered
' @param StopOrderType  the type of stop loss order to place:
'           StopOrderTypeNone               no stop loss order
'           StopOrderTypeStop               use a stop order
'           StopOrderTypeStopLimit          use a stop limit order
'           StopOrderTypeBid                use a stop order at the current Bid (see StopOffset)
'           StopOrderTypeAsk                use a stop order at the current Ask (see StopOffset)
'           StopOrderTypeLast               use a stop order at the last Trade price (see StopOffset)
'           StopOrderTypeAuto               use a stop order at the fill price for the
'                                           entry order (see StopOffset)
' @param StopPrice      the trigger price for StopOrderTypeStop and StopOrderTypeStopLimit
' @param StopOffset     number of ticks to modify the StopPrice. For example, for a stop loss
'                       order 4 ticks below the entry order fill price, use StopOrderTypeAuto
'                       and set StopOffset to -4
' @param StopLimitPrice the limit price for StopOrderTypeStopLimit
' @param TargetOrderType the type of target order to place:
'           TargetOrderTypeNone             no target order
'           TargetOrderTypeLimit            use a limit order
'           TargetOrderTypeLimitIfTouched   use a limit if touched order
'           TargetOrderTypeMarketIfTouched  use a market if touched order
'           TargetOrderTypeBid              use a limit order at the current Bid (see TargetOffset)
'           TargetOrderTypeAsk              use a limit order at the current Ask (see TargetOffset)
'           TargetOrderTypeLast             use a limit order at the last Trade price (see TargetOffset)
'           TargetOrderTypeAuto             use a limit order at the fill price for the
'                                           entry order (see TargetOffset)
' @param TargetPrice    the limit price for the target order
' @param TargetOffset   number of ticks to modify the TargetPrice. For example, for a target
'                       order 10 ticks above the entry order fill price, use TargetOrderTypeAuto
'                       and set TargetOffset to 10
' @param TargetTriggerPrice  the trigger price for TargetOrderTypeLimitIfTouched and TargetOrderTypeMarketIfTouched
' @param CancelPrice    if the entry order is a limit order and market price rises to this
'                       price before the entry order is filled, then the order is cancelled.
'                       If the entry order is a stop or stop limit order and market price
'                       falls to this price before the entry order is filled, then the order
'                       is cancelled
' @param EntryTIF       specifies the time in force for the entry order
' @param StopTIF        specifies the time in force for the stop loss order
' @param TargetTIF      specifies the time in force for the target order
' @param StopStrategy   an object that manages the stop loss order
' @param ExitStrategy   an object that manages the exit strategy
' @param ScalingStrategy an object that manages scaling into the position (ie placing further
'                       orders to increase the position)
'
' @see
'
'@/
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

Dim failpoint As Long
On Error GoTo Err

gLogLogger.Log LogLevelNormal, _
                "Buy " & mTicker.Contract.specifier.localSymbol & _
                " number=" & number & _
                "; EntryOrderType=" & gEntryOrderTypeToShortString(EntryOrderType) & _
                "; buyPrice=" & buyPrice & _
                "; buyOffset=" & buyOffset & _
                "; triggerPrice=" & triggerPrice & _
                "; StopOrderType=" & gStopOrderTypeToShortString(StopOrderType) & _
                "; stopPrice=" & stopPrice & _
                "; stopOffset=" & stopOffset & _
                "; stopLimitPrice=" & stopLimitPrice & _
                "; TargetOrderType=" & gTargetOrderTypeToShortString(TargetOrderType) & _
                "; targetPrice=" & targetPrice & _
                "; targetOffset=" & targetOffset & _
                "; targetTriggerPrice=" & targetTriggerPrice & _
                "; cancelPrice=" & cancelPrice & _
                "; entryTIF=" & gOrderTIFToString(entryTIF) & _
                "; stopTIF=" & gOrderTIFToString(stopTIF) & _
                "; targetTIF=" & gOrderTIFToString(targetTIF) & _
                "; stopStrategy=" & TypeName(stopStrategy) & _
                "; exitStrategy=" & TypeName(exitStrategy) & _
                "; scalingStrategy=" & TypeName(scalingStrategy)
                
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
                                
Buy.entryOrder.timeInForce = entryTIF
If Not Buy.stoporder Is Nothing Then Buy.stoporder.timeInForce = stopTIF
If Not Buy.targetOrder Is Nothing Then Buy.targetOrder.timeInForce = targetTIF

executeOrderPlex Buy, True, cancelPrice, stopStrategy, exitStrategy, scalingStrategy

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Buy" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Function Cancel() As Boolean
Dim orderPlex As orderPlex

Dim failpoint As Long
On Error GoTo Err

For Each orderPlex In mActiveOrderPlexes
    If orderPlex.Cancel(evenIfFilled:=True) Then Cancel = True
Next

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Cancel" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Public Function CancelPending() As Boolean
Dim orderPlex As orderPlex

Dim failpoint As Long
On Error GoTo Err

For Each orderPlex In mActiveOrderPlexes
    If orderPlex.Cancel(evenIfFilled:=False) Then CancelPending = True
Next

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "CancelPending" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

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
Dim failpoint As Long
On Error GoTo Err

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

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "CreateBuyOrderPlex" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
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

If Not isReady Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                            "OrderContext::CreateOrderPlex", _
                            "The service provider is not Ready"

If number <= 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                            "OrderContext::CreateOrderPlex", _
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
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "OrderContext::CreateOderPlex", _
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
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "OrderContext::CreateOderPlex", _
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
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
    "OrderContext::CreateOderPlex", _
    "Invalid target type"
End Select

Set orderPlex = New orderPlex
orderPlex.OrderContext = Me
orderPlex.Key = nextPlexKey
orderPlex.PositionManager = mPositionManager
orderPlex.Ticker = mTicker
orderPlex.orderSubmitter = mOrderSubmitter
orderPlex.IsSimulated = mIsSimulated
orderPlex.PreventUnprotectedPosition = mPreventUnprotectedPositions

Set entryOrder = orderPlex.createEntryOrder
entryOrder.action = action
entryOrder.quantity = number

entryOrder.orderType = gEntryOrderTypeToOrderType(EntryOrderType)
Select Case EntryOrderType
Case EntryOrderTypeMarket
Case EntryOrderTypeMarketOnOpen
Case EntryOrderTypeMarketOnClose
Case EntryOrderTypeMarketIfTouched
    If triggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
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
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    
    entryOrder.limitPrice = price + offset * mTickSize
Case EntryOrderTypeLimitOnOpen
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    
    entryOrder.limitPrice = price + offset * mTickSize
Case EntryOrderTypeLimitOnClose
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    
    entryOrder.limitPrice = price + offset * mTickSize
Case EntryOrderTypeLimitIfTouched
    If price = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The entry price must not be zero for this order type"
    If triggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The trigger price must not be zero for this order type"

    entryOrder.limitPrice = price + offset * mTickSize
    entryOrder.triggerPrice = triggerPrice
Case EntryOrderTypeStop
    If triggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The trigger price must not be zero for this order type"
    
    entryOrder.triggerPrice = triggerPrice
Case EntryOrderTypeStopLimit
    If triggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "OrderContext::CreateOrderPlex", _
                                "The trigger price must not be zero for this order type"
    
    entryOrder.limitPrice = price + offset * mTickSize
    entryOrder.triggerPrice = triggerPrice
End Select

entryOrder.offset = offset
lEntryPrice = entryOrder.limitPrice
lTriggerPrice = entryOrder.triggerPrice

If StopOrderType <> StopOrderTypeNone Then
    Set stoporder = orderPlex.createStopOrder
    stoporder.action = IIf(action = ActionBuy, OrderActions.ActionSell, OrderActions.ActionBuy)
    stoporder.quantity = number
    
    stoporder.orderType = gStopOrderTypeToOrderType(StopOrderType)
    Select Case StopOrderType
    Case StopOrderTypeStop
        If stopTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The stopTrigger price must not be zero for this order type"
        
        stoporder.triggerPrice = stopTriggerPrice + stopOffset * mTickSize
    Case StopOrderTypeStopLimit
        If stopTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The stopTrigger price must not be zero for this order type"
        If stopLimitPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
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
        stoporder.offset = stopOffset * mTickSize
    End Select
    lStopPrice = stoporder.triggerPrice
End If

If TargetOrderType <> TargetOrderTypeNone Then
    Set targetOrder = orderPlex.createTargetOrder
    targetOrder.action = IIf(action = ActionBuy, OrderActions.ActionSell, OrderActions.ActionBuy)
    targetOrder.quantity = number
    
    targetOrder.orderType = gTargetOrderTypeToOrderType(TargetOrderType)
    Select Case TargetOrderType
    Case TargetOrderTypeLimit
        If targetPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The target price must not be zero for this order type"
        
        targetOrder.limitPrice = targetPrice + targetOffset * mTickSize
    Case TargetOrderTypeLimitIfTouched
        If targetPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The target price must not be zero for this order type"
        If targetTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The target trigger price must not be zero for this order type"
    
        targetOrder.limitPrice = targetPrice + offset * mTickSize
        targetOrder.triggerPrice = targetTriggerPrice
    Case TargetOrderTypeMarketIfTouched
        If targetTriggerPrice = 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                    "OrderContext::CreateOrderPlex", _
                                    "The target trigger price must not be zero for this order type"
        
        targetOrder.triggerPrice = targetTriggerPrice
    Case TargetOrderTypeBid
        targetOrder.limitPrice = mTicker.BidPrice + targetOffset * mTickSize
    Case TargetOrderTypeAsk
        targetOrder.limitPrice = mTicker.AskPrice + targetOffset * mTickSize
    Case TargetOrderTypeLast
        targetOrder.limitPrice = mTicker.TradePrice + targetOffset * mTickSize
    Case TargetOrderTypeAuto
        targetOrder.offset = targetOffset * mTickSize
    End Select

    lTargetPrice = targetOrder.limitPrice
End If

orderPlex.checkpoint

Set CreateOrderPlex = orderPlex

gTB.notifyListeners _
            "Create order plex: (" & mTicker.TradePrice & ") " & _
            orderPlex.Key & ": " & _
            gOrderActionToString(action) & " " & number & " at " & _
            gOrderTypeToString(entryOrder.orderType) & " " & lEntryPrice & _
            IIf(lTriggerPrice <> 0, "(" & lTriggerPrice & ")", "") & _
            ", stop " & IIf(StopOrderType = StopOrderTypeAuto, "AUTO(" & stopOffset * mTickSize & ")", lStopPrice) & _
            ", target " & IIf(TargetOrderType = TargetOrderTypeAuto, "AUTO(" & targetOffset * mTickSize & ")", lTargetPrice), _
            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
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
Dim failpoint As Long
On Error GoTo Err

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

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "CreateSellOrderPlex" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
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

Dim failpoint As Long
On Error GoTo Err

If orderPlex Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                            "OrderContext::executeOrderPlex", _
                            "No order plex supplied"

If checkMoneyManagement And Not mIsSimulated Then
    origQuantity = orderPlex.entryOrder.quantity
    newQuantity = origQuantity
    If Not mPositionManager.AllowOrder( _
                            IIf(orderPlex.entryOrder.action = ActionBuy, True, False), _
                            newQuantity, _
                            orderPlex.Risk, _
                            msg) _
    Then
        gTB.notifyListeners "Rejected: " & msg, _
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

gTB.notifyListeners _
            "Execute order plex: (" & mTicker.TradePrice & ") " & _
            orderPlex.Key, _
            IIf(mIsSimulated, TradeBuildListenValueTypes.VTSimulatedOrder, TradeBuildListenValueTypes.VTOrder), _
            Me

mActiveOrderPlexes.Add orderPlex
orderPlex.execute cancelPrice

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "executeOrderPlex" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Friend Function Finish()
Dim op As orderPlex

Set mTicker = Nothing
Set mPositionManager = Nothing
For Each op In mActiveOrderPlexes
    op.Finish
Next
Set mActiveOrderPlexes = Nothing

For Each op In mClosedOrderPlexes
    op.Finish
Next
Set mClosedOrderPlexes = Nothing

Set mOrderSubmitter = Nothing
End Function

Public Function getAccountBalance() As Double
getAccountBalance = 10000                       '!!!!!!!!!!!!!!!!!!!!!!!!! need to fix this up
End Function

Public Function isAttributeSupported( _
                ByVal attributeId As OrderAttributeIds) As Boolean
isAttributeSupported = (mPermittedOrderAttributes And attributeId)
End Function

Public Function isOrderTifValueSupported( _
                ByVal value As OrderTifs) As Boolean
isOrderTifValueSupported = (mPermittedOrderTifs And value)
End Function

Public Function isOrderTypeValueSupported( _
                ByVal value As OrderTypes) As Boolean
If value = OrderTypeAutoLimit Or _
    value = OrderTypeAutoStop _
Then
    isOrderTypeValueSupported = True
Else
    isOrderTypeValueSupported = (mPermittedOrderTypes And value)
End If
End Function

Public Function isStopTriggerMethodValueSupported( _
                ByVal value As StopTriggerMethods) As Boolean
isStopTriggerMethodValueSupported = (mPermittedStopTriggerMethods And value)
End Function

Friend Function newOrder() As order
Dim lOrder As order
Set lOrder = New order
With lOrder
    .id = GenerateGUIDString
    .IsSimulated = mIsSimulated
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
    If isAttributeSupported(OrderAttributeIds.OrderAttSettlingFirm) Then .settlingFirm = mSettlingFirm
    
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

mPositionManager.Orders.Add lOrder

Set newOrder = lOrder
End Function

Friend Function orderPlexClosed( _
                ByVal orderPlex As orderPlex)
mClosedOrderPlexes.Add orderPlex
mActiveOrderPlexes.Remove orderPlex
mCompletedSize = mCompletedSize + orderPlex.size
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

Dim failpoint As Long
On Error GoTo Err

gLogLogger.Log LogLevelNormal, _
                "Sell " & mTicker.Contract.specifier.localSymbol & _
                " number=" & number & _
                "; EntryOrderType=" & gEntryOrderTypeToShortString(EntryOrderType) & _
                "; sellPrice=" & sellPrice & _
                "; sellOffset=" & sellOffset & _
                "; triggerPrice=" & triggerPrice & _
                "; StopOrderType=" & gStopOrderTypeToShortString(StopOrderType) & _
                "; stopPrice=" & stopPrice & _
                "; stopOffset=" & stopOffset & _
                "; stopLimitPrice=" & stopLimitPrice & _
                "; TargetOrderType=" & gTargetOrderTypeToShortString(TargetOrderType) & _
                "; targetPrice=" & targetPrice & _
                "; targetOffset=" & targetOffset & _
                "; targetTriggerPrice=" & targetTriggerPrice & _
                "; cancelPrice=" & cancelPrice & _
                "; entryTIF=" & gOrderTIFToString(entryTIF) & _
                "; stopTIF=" & gOrderTIFToString(stopTIF) & _
                "; targetTIF=" & gOrderTIFToString(targetTIF) & _
                "; stopStrategy=" & TypeName(stopStrategy) & _
                "; exitStrategy=" & TypeName(exitStrategy) & _
                "; scalingStrategy=" & TypeName(scalingStrategy)
                
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

Sell.entryOrder.timeInForce = entryTIF
If Not Sell.stoporder Is Nothing Then Sell.stoporder.timeInForce = stopTIF
If Not Sell.targetOrder Is Nothing Then Sell.targetOrder.timeInForce = targetTIF

executeOrderPlex Sell, True, cancelPrice, stopStrategy, exitStrategy, scalingStrategy

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Sell" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub attributeValueNotSupported( _
                ByVal pAttribute As OrderAttributeIds, _
                ByVal value As Variant)
Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            , _
            "Value " & CStr(value) & " for attribute " & gOrderAttributeToString(pAttribute) & " is not supported by the service provider"
End Sub

Private Sub checkAttributeSupported( _
                ByVal pAttribute As OrderAttributeIds)
If Not isAttributeSupported(pAttribute) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                , _
                "The " & gOrderAttributeToString(pAttribute) & " attribute is not supported by the service provider"
End If
End Sub

Private Sub checkReady()
If isReady Then
    mPermittedOrderAttributes = mOrderSubmitter.permittedOrderAttributes
    mPermittedOrderTifs = mOrderSubmitter.permittedOrderTifs
    mPermittedOrderTypes = mOrderSubmitter.permittedOrderTypes
    mPermittedStopTriggerMethods = mOrderSubmitter.permittedStopTriggerMethods
    RaiseEvent ready
Else
    RaiseEvent NotReady
End If
End Sub

Private Function nextPlexKey() As String
Static currID As Long
currID = currID + 1
nextPlexKey = mTicker.Key & "/" & mName & "/" & currID
End Function

