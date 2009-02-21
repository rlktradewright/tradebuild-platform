VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Order"
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

Implements TradeBuildSP.IOrder

'@================================================================================
' Events
'@================================================================================

Event Clean()
Event Dirty()
Event Fill( _
                ByVal quantityFilled As Long, _
                ByVal fillPrice As Double)
Event PropertyChanged()
Event StatusChanged()

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTicker                     As Ticker
Private mContract                   As Contract
Private mOrderContext               As OrderContext

Private mIsDirty                    As Boolean

Private mIsSimulated                As Boolean

Private mStatus                     As OrderStatuses
Private mIdentifier                 As String

Private mFillTime                   As Date

Private mBrokerId                   As String
    
Private mAction                     As OrderActions
Private mQuantity                   As Long
Private mQuantityFilled             As Long
Private mQuantityRemaining          As Long
Private mAveragePrice               As Double
Private mLastFillPrice              As Double
Private mOrderType                  As OrderTypes
Private mLimitPrice                 As Double
Private mTriggerPrice               As Double
Private mOffset                     As Long

Private mTimeInForce                As OrderTifs
Private mOrigin                     As OrderOrigins
Private mOriginatorRef              As String
Private mProviderProperties         As parameters
Private mBlockOrder                 As Boolean
Private mSweepToFill                As Boolean
Private mDisplaySize                As Long
Private mStopTriggerMethod          As StopTriggerMethods
Private mIgnoreRegularTradingHours  As Boolean
Private mHidden                     As Boolean
Private mDiscretionaryAmount        As Double
Private mGoodAfterTime              As Date
Private mGoodAfterTimeTZ            As String
Private mGoodTillDate               As Date
Private mGoodTillDateTZ             As String
    
Private mSettlingFirm               As String
Private mAllOrNone                  As Boolean
Private mMinimumQuantity            As Long
Private mPercentOffset              As Double
Private mETradeOnly                 As Boolean
Private mFirmQuoteOnly              As Boolean
Private mNBBOPriceCap               As Double
Private mOverrideConstraints        As Boolean
    
'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mStatus = OrderStatuses.OrderStatusCreated
mAction = ActionBuy
mOrderType = OrderTypeNone
mTimeInForce = TIFDay

Set mProviderProperties = New parameters
End Sub

'@================================================================================
' IOrder Interface Members
'@================================================================================

Private Property Get IOrder_Action() As TradeBuildSP.OrderActions
IOrder_Action = mAction
End Property

Private Property Get IOrder_allOrNone() As Boolean
IOrder_allOrNone = mAllOrNone
End Property

Private Property Get IOrder_blockOrder() As Boolean
IOrder_blockOrder = mBlockOrder
End Property

Private Property Let IOrder_brokerId(ByVal RHS As String)
mBrokerId = RHS
End Property

Private Property Get IOrder_brokerId() As String
IOrder_brokerId = mBrokerId
End Property

Private Property Get IOrder_Contract() As Contract
Set IOrder_Contract = mContract
End Property

Private Property Get IOrder_discretionaryAmount() As Double
IOrder_discretionaryAmount = mDiscretionaryAmount
End Property

Private Property Get IOrder_displaySize() As Long
IOrder_displaySize = mDisplaySize
End Property

Private Property Get IOrder_eTradeOnly() As Boolean
IOrder_eTradeOnly = mETradeOnly
End Property

Private Property Get IOrder_firmQuoteOnly() As Boolean
IOrder_firmQuoteOnly = mFirmQuoteOnly
End Property

Private Property Get IOrder_goodAfterTime() As Date
IOrder_goodAfterTime = mGoodAfterTime
End Property

Private Property Get IOrder_goodAfterTimeTZ() As String
IOrder_goodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Private Property Get IOrder_goodTillDate() As Date
IOrder_goodTillDate = mGoodTillDate
End Property

Private Property Get IOrder_goodTillDateTZ() As String
IOrder_goodTillDateTZ = mGoodTillDateTZ
End Property

Private Property Get IOrder_hidden() As Boolean
IOrder_hidden = mHidden
End Property

Private Property Get IOrder_ignoreRegularTradingHours() As Boolean
IOrder_ignoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Private Property Get IOrder_limitPrice() As Double
IOrder_limitPrice = mLimitPrice
End Property

Private Property Get IOrder_minimumQuantity() As Long
IOrder_minimumQuantity = mMinimumQuantity
End Property

Private Property Get IOrder_NbboPriceCap() As Double
IOrder_NbboPriceCap = mNBBOPriceCap
End Property

Private Property Get IOrder_orderType() As TradeBuildSP.OrderTypes
IOrder_orderType = mOrderType
End Property

Private Property Get IOrder_origin() As TradeBuildSP.OrderOrigins
IOrder_origin = mOrigin
End Property

Private Property Get IOrder_originatorRef() As String
IOrder_originatorRef = mOriginatorRef
End Property

Private Property Get IOrder_overrideConstraints() As Boolean
IOrder_overrideConstraints = mOverrideConstraints
End Property

Private Property Get IOrder_providerProperties() As TWUtilities30.parameters
Set IOrder_providerProperties = mProviderProperties
End Property

Private Property Get IOrder_quantity() As Long
IOrder_quantity = mQuantity
End Property

Private Property Get IOrder_quantityFilled() As Long
IOrder_quantityFilled = mQuantityFilled
End Property

Private Property Get IOrder_quantityRemaining() As Long
IOrder_quantityRemaining = mQuantityRemaining
End Property

Private Property Get IOrder_settlingFirm() As String
IOrder_settlingFirm = mSettlingFirm
End Property

Private Property Get IOrder_status() As TradeBuildSP.OrderStatuses
IOrder_status = mStatus
End Property

Private Property Get IOrder_sweepToFill() As Boolean
IOrder_sweepToFill = mSweepToFill
End Property

Private Property Get IOrder_timeInForce() As TradeBuildSP.OrderTifs
IOrder_timeInForce = mTimeInForce
End Property

Private Property Get IOrder_tradeBuildId() As String
IOrder_tradeBuildId = mIdentifier
End Property

Private Property Get IOrder_stopTriggerMethod() As TradeBuildSP.StopTriggerMethods
IOrder_stopTriggerMethod = mStopTriggerMethod
End Property

Private Property Get IOrder_triggerPrice() As Double
IOrder_triggerPrice = mTriggerPrice
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let action(ByVal value As OrderActions)
mAction = UCase$(value)
End Property

Public Property Get action() As OrderActions
action = mAction
End Property

Public Property Let allOrNone(ByVal value As Boolean)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttAllOrNone
If value = mAllOrNone Then Exit Property
setDirty
mAllOrNone = value
End Property

Public Property Get allOrNone() As Boolean
allOrNone = mAllOrNone
End Property

Friend Property Let averagePrice(ByVal value As Double)
mAveragePrice = value
End Property

Public Property Get averagePrice() As Double
averagePrice = mAveragePrice
End Property

Public Property Let blockOrder(value As Boolean)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttBlockOrder
If value = mBlockOrder Then Exit Property
setDirty
mBlockOrder = value
End Property

Public Property Get blockOrder() As Boolean
blockOrder = mBlockOrder
End Property

Friend Property Let brokerId(ByVal value As String)
mBrokerId = value
End Property

Public Property Get brokerId() As String
brokerId = mBrokerId
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Public Property Let displaySize(ByVal value As Long)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttDisplaySize
If value = mDisplaySize Then Exit Property
setDirty
mDisplaySize = value
End Property

Public Property Get displaySize() As Long
displaySize = mDisplaySize
End Property

Public Property Let discretionaryAmount(ByVal value As Double)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttDiscretionaryAmount
If value = mDiscretionaryAmount Then Exit Property
setDirty
mDiscretionaryAmount = value
End Property

Public Property Get discretionaryAmount() As Double
discretionaryAmount = mDiscretionaryAmount
End Property

Public Property Let eTradeOnly(ByVal value As Boolean)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttETradeOnly
If value = mETradeOnly Then Exit Property
setDirty
mETradeOnly = value
End Property

Public Property Get eTradeOnly() As Boolean
eTradeOnly = mETradeOnly
End Property

Friend Property Let fillTime(ByVal value As Date)
mFillTime = value
End Property

Public Property Get fillTime() As Date
fillTime = mFillTime
End Property

Public Property Let firmQuoteOnly(ByVal value As Boolean)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttFirmQuoteOnly
If value = mFirmQuoteOnly Then Exit Property
setDirty
mFirmQuoteOnly = value
End Property

Public Property Get firmQuoteOnly() As Boolean
firmQuoteOnly = mFirmQuoteOnly
End Property

Public Property Let goodAfterTime(ByVal value As Date)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttGoodAfterTime
If value = mGoodAfterTime Then Exit Property
setDirty
mGoodAfterTime = value
End Property

Public Property Get goodAfterTime() As Date
goodAfterTime = mGoodAfterTime
End Property

Public Property Let goodAfterTimeTZ(ByVal value As String)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttGoodAfterTimeTZ
If value = mGoodAfterTimeTZ Then Exit Property
setDirty
mGoodAfterTimeTZ = value
End Property

Public Property Get goodAfterTimeTZ() As String
goodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Public Property Let goodTillDate(ByVal value As Date)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttGoodTillDate
If value = mGoodTillDate Then Exit Property
setDirty
mGoodTillDate = value
End Property

Public Property Get goodTillDate() As Date
goodTillDate = mGoodTillDate
End Property

Public Property Let goodTillDateTZ(ByVal value As String)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttGoodTillDateTZ
If value = mGoodTillDateTZ Then Exit Property
setDirty
mGoodTillDateTZ = value
End Property

Public Property Get goodTillDateTZ() As String
goodTillDateTZ = mGoodTillDateTZ
End Property

Public Property Let hidden(value As Boolean)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttHidden
If value = mHidden Then Exit Property
setDirty
mHidden = value
End Property

Public Property Get hidden() As Boolean
hidden = mHidden
End Property

Friend Property Let id(ByVal value As String)
mIdentifier = value
End Property

Public Property Get id() As String
id = mIdentifier
End Property

Public Property Let ignoreRegularTradingHours(value As Boolean)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttIgnoreRTH
If value = mIgnoreRegularTradingHours Then Exit Property
setDirty
mIgnoreRegularTradingHours = value
End Property

Public Property Get ignoreRegularTradingHours() As Boolean
ignoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Public Property Get isDirty() As Boolean
isDirty = mIsDirty
End Property

Friend Property Let IsSimulated(ByVal value As Boolean)
mIsSimulated = value
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Let lastFillPrice(ByVal value As Double)
mLastFillPrice = value
End Property

Public Property Get lastFillPrice() As Double
lastFillPrice = mLastFillPrice
End Property

Friend Property Let limitPrice(ByVal value As Double)
mLimitPrice = value
End Property

Public Property Get limitPrice() As Double
limitPrice = mLimitPrice
End Property

Public Property Let minimumQuantity(ByVal value As Long)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttMinimumQuantity
If value = mMinimumQuantity Then Exit Property
setDirty
mMinimumQuantity = value
End Property

Public Property Get minimumQuantity() As Long
minimumQuantity = mMinimumQuantity
End Property

Public Property Let NbboPriceCap(ByVal value As Double)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttNBBOPriceCap
If value = mNBBOPriceCap Then Exit Property
setDirty
mNBBOPriceCap = value
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mNBBOPriceCap
End Property

Friend Property Let offset(ByVal value As Long)
mOffset = value
End Property

Friend Property Get offset() As Long
offset = mOffset
End Property

Friend Property Let OrderContext(ByVal value As OrderContext)
Set mOrderContext = value
End Property

Public Property Get OrderContext() As OrderContext
Set OrderContext = mOrderContext
End Property

Friend Property Let orderType(ByVal value As OrderTypes)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttOrderType
If Not mOrderContext.isOrderTypeValueSupported(value) Then attributeValueNotSupported TradeBuild26.OrderAttributeIds.OrderAttOrderType, value
If value = mOrderType Then Exit Property
setDirty
mOrderType = value
End Property

Public Property Get orderType() As OrderTypes
orderType = mOrderType
End Property

Public Property Let origin(ByVal value As OrderOrigins)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttOrigin
If value = mOrigin Then Exit Property
setDirty
mOrigin = value
End Property

Public Property Get origin() As OrderOrigins
origin = mOrigin
End Property

Public Property Let originatorRef(ByVal value As String)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttOriginatorRef
If value = mOriginatorRef Then Exit Property
setDirty
mOriginatorRef = value
End Property

Public Property Get originatorRef() As String
originatorRef = mOriginatorRef
End Property

Public Property Let overrideConstraints(ByVal value As Boolean)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttOverrideConstraints
If value = mOverrideConstraints Then Exit Property
setDirty
mOverrideConstraints = value
End Property

Public Property Get overrideConstraints() As Boolean
overrideConstraints = mOverrideConstraints
End Property

Public Property Let percentOffset(ByVal value As Double)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttPercentOffset
If value = mPercentOffset Then Exit Property
setDirty
mPercentOffset = value
End Property

Public Property Get percentOffset() As Double
percentOffset = mPercentOffset
End Property

Friend Property Let providerProperties( _
                ByVal value As TWUtilities30.parameters)
Set mProviderProperties = value
End Property

Public Property Get providerProperties() As TWUtilities30.parameters
Set providerProperties = mProviderProperties
End Property

Public Property Let quantity(ByVal value As Long)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttQuantity
If value = mQuantity Then Exit Property
mQuantity = value
setDirty
mQuantityRemaining = mQuantity - mQuantityFilled
End Property

Public Property Get quantity() As Long
quantity = mQuantity
End Property

Friend Property Let quantityFilled(ByVal value As Long)
mQuantityFilled = value
End Property

Public Property Get quantityFilled() As Long
quantityFilled = mQuantityFilled
End Property

Friend Property Let quantityRemaining(ByVal value As Long)
mQuantityRemaining = value
End Property

Public Property Get quantityRemaining() As Long
quantityRemaining = mQuantityRemaining
End Property

Public Property Let settlingFirm(value As String)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttSettlingFirm
If value = mSettlingFirm Then Exit Property
setDirty
mSettlingFirm = value
End Property

Public Property Get settlingFirm() As String
settlingFirm = mSettlingFirm
End Property

Friend Property Let status(value As OrderStatuses)
mStatus = value
End Property

Public Property Get status() As OrderStatuses
status = mStatus
End Property

Public Property Let StopTriggerMethod(ByVal value As StopTriggerMethods)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttStopTriggerMethod
If Not mOrderContext.isStopTriggerMethodValueSupported(value) Then attributeValueNotSupported TradeBuild26.OrderAttributeIds.OrderAttStopTriggerMethod, value
If value = mStopTriggerMethod Then Exit Property
setDirty
mStopTriggerMethod = value
End Property

Public Property Get StopTriggerMethod() As StopTriggerMethods
StopTriggerMethod = mStopTriggerMethod
End Property

Public Property Let SweepToFill(value As Boolean)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttSweepToFill
If value = mSweepToFill Then Exit Property
setDirty
mSweepToFill = value
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSweepToFill
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Set mTicker = value
Set mContract = mTicker.Contract
End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
End Property

Public Property Let timeInForce(ByVal value As OrderTifs)
checkAttributeModificationPermitted TradeBuild26.OrderAttributeIds.OrderAttTimeInForce
If Not mOrderContext.isOrderTifValueSupported(value) Then attributeValueNotSupported TradeBuild26.OrderAttributeIds.OrderAttTimeInForce, value
If value = mTimeInForce Then Exit Property
setDirty
mTimeInForce = value
End Property

Public Property Get timeInForce() As OrderTifs
timeInForce = mTimeInForce
End Property

Friend Property Let triggerPrice(ByVal value As Double)
mTriggerPrice = value
End Property

Public Property Get triggerPrice() As Double
triggerPrice = mTriggerPrice
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As order
Set clone = New order
syncToMe clone
End Function

''
' Notifies a fill for this <code>order</code>.
'
' @param exec  an <code>Execution</code> object containing details of the fill
'
'@/
Friend Sub notifyFill( _
                ByVal exec As Execution)

If Not mIsSimulated Then gTB.notifyListeners _
                        mContract.specifier.exchange & "," & _
                        mContract.specifier.symbol & "," & _
                        mContract.specifier.expiry & "," & _
                        exec.quantity & "," & _
                        exec.price & "," & _
                        gOrderActionToString(exec.action) & "," & _
                        exec.time & "," & _
                        exec.orderTradeBuildId & "," & _
                        exec.orderBrokerId, _
                        IIf(mIsSimulated, TradeBuildListenValueTypes.VTOrderDetailSimulated, TradeBuildListenValueTypes.VTOrderDetail), _
                        Me

mAveragePrice = (mQuantityFilled * mAveragePrice + exec.quantity * exec.price) / (mQuantityFilled + exec.quantity)
mQuantityFilled = mQuantityFilled + exec.quantity
mQuantityRemaining = mQuantityRemaining - exec.quantity
mLastFillPrice = exec.price
mFillTime = exec.time
RaiseEvent Fill(exec.quantity, exec.price)

End Sub

Friend Sub finish()
Set mTicker = Nothing
Set mContract = Nothing
Set mOrderContext = Nothing
End Sub

Public Function isAttributeModifiable( _
                ByVal attributeId As OrderAttributeIds) As Boolean

If Not mOrderContext.isAttributeSupported(attributeId) Then Exit Function

Select Case mStatus
Case OrderStatusCreated
    isAttributeModifiable = True
Case OrderStatusRejected
    isAttributeModifiable = True
Case OrderStatusPendingSubmit
    isAttributeModifiable = attributeId And _
                            (OrderAttributeIds.OrderAttAllOrNone Or _
                            OrderAttributeIds.OrderAttBlockOrder Or _
                            OrderAttributeIds.OrderAttDiscretionaryAmount Or _
                            OrderAttributeIds.OrderAttDisplaySize Or _
                            OrderAttributeIds.OrderAttETradeOnly Or _
                            OrderAttributeIds.OrderAttFirmQuoteOnly Or _
                            OrderAttributeIds.OrderAttGoodTillDate Or _
                            OrderAttributeIds.OrderAttGoodTillDateTZ Or _
                            OrderAttributeIds.OrderAttHidden Or _
                            OrderAttributeIds.OrderAttLimitPrice Or _
                            OrderAttributeIds.OrderAttMinimumQuantity Or _
                            OrderAttributeIds.OrderAttOriginatorRef Or _
                            OrderAttributeIds.OrderAttQuantity Or _
                            OrderAttributeIds.OrderAttSweepToFill Or _
                            OrderAttributeIds.OrderAttTimeInForce Or _
                            OrderAttributeIds.OrderAttTriggerPrice)
Case OrderStatusPreSubmitted
    isAttributeModifiable = attributeId And _
                            (OrderAttributeIds.OrderAttAllOrNone Or _
                            OrderAttributeIds.OrderAttBlockOrder Or _
                            OrderAttributeIds.OrderAttDiscretionaryAmount Or _
                            OrderAttributeIds.OrderAttDisplaySize Or _
                            OrderAttributeIds.OrderAttETradeOnly Or _
                            OrderAttributeIds.OrderAttFirmQuoteOnly Or _
                            OrderAttributeIds.OrderAttGoodTillDate Or _
                            OrderAttributeIds.OrderAttGoodTillDateTZ Or _
                            OrderAttributeIds.OrderAttHidden Or _
                            OrderAttributeIds.OrderAttLimitPrice Or _
                            OrderAttributeIds.OrderAttMinimumQuantity Or _
                            OrderAttributeIds.OrderAttOriginatorRef Or _
                            OrderAttributeIds.OrderAttQuantity Or _
                            OrderAttributeIds.OrderAttSweepToFill Or _
                            OrderAttributeIds.OrderAttTimeInForce Or _
                            OrderAttributeIds.OrderAttTriggerPrice)
Case OrderStatusSubmitted
    isAttributeModifiable = attributeId And _
                            (OrderAttributeIds.OrderAttDiscretionaryAmount Or _
                            OrderAttributeIds.OrderAttDisplaySize Or _
                            OrderAttributeIds.OrderAttGoodTillDate Or _
                            OrderAttributeIds.OrderAttGoodTillDateTZ Or _
                            OrderAttributeIds.OrderAttHidden Or _
                            OrderAttributeIds.OrderAttLimitPrice Or _
                            OrderAttributeIds.OrderAttMinimumQuantity Or _
                            OrderAttributeIds.OrderAttQuantity Or _
                            OrderAttributeIds.OrderAttTimeInForce Or _
                            OrderAttributeIds.OrderAttTriggerPrice)
Case OrderStatusFilled
    isAttributeModifiable = False
Case OrderStatusCancelling
    isAttributeModifiable = False
Case OrderStatusCancelled
    isAttributeModifiable = False
End Select

Select Case attributeId
Case OrderAttributeIds.OrderAttLimitPrice
    Select Case mOrderType
    Case OrderTypeLimit, _
        OrderTypeLimitOnClose, _
        OrderTypeStopLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeLimitOnOpen
    Case Else
        isAttributeModifiable = False
    End Select
Case OrderAttributeIds.OrderAttTriggerPrice
    Select Case mOrderType
    Case OrderTypeStop, _
        OrderTypeStopLimit, _
        OrderTypeLimitIfTouched, _
        OrderTypeMarketIfTouched
    Case Else
        isAttributeModifiable = False
    End Select
End Select
End Function

Public Function isModifiable() As Boolean
If mStatus <> OrderStatuses.OrderStatusCancelled And _
    mStatus <> OrderStatuses.OrderStatusCancelling And _
    mStatus <> OrderStatuses.OrderStatusFilled _
Then isModifiable = True
End Function

' commented out 22/5/6
'Friend Sub openOrder(ByVal action As OrderActions, _
'                        ByVal quantity As Long, _
'                        ByVal orderType As OrderTypes, _
'                        ByVal lmtPrice As Double, _
'                        ByVal auxPrice As Double, _
'                        ByVal tif As OrderTifs, _
'                        ByVal ocaGroup As String)
'
'mTWSAction = action
'If action <> mAction Then
'    mDifferent = True
'End If
'
'mTWSQuantity = quantity
'If quantity <> mQuantity Then
'    mDifferent = True
'End If
'
'mTWSOrderType = orderType
'If orderType <> mOrderType Then
'    mDifferent = True
'End If
'If orderType = OrderTypes.OrderTypeLimit Or _
'    orderType = OrderTypes.OrderTypeStopLimit _
'Then
'    mTWSLimitPrice = lmtPrice
'    If lmtPrice <> mLimitPrice Then
'        mDifferent = True
'    End If
'Else
'    mTWSLimitPrice = 0
'End If
'If orderType = OrderTypes.OrderTypeStop Or _
'    orderType = OrderTypes.OrderTypeStopLimit _
'Then
'    mTWSTriggerPrice = auxPrice
'    If auxPrice <> mTriggerPrice Then
'        mDifferent = True
'    End If
'Else
'    mTWSTriggerPrice = 0
'End If
'' NB: mDifferent will be cleared when updateStatus is called (this always
'' happens shortly after openOrder)
''If mDifferent Then Stop
'End Sub

''
' Marks this order as clean, ie all changes have been either
' actioned or cancelled.
'@/
Friend Sub setClean()
mIsDirty = False
RaiseEvent Clean
End Sub

''
' Synchronises the specified order to this order so that both are
' identical.
'
' @param  targetOrder the <code>order</code> that must be made identical
'@/
Friend Sub syncToMe(ByVal targetOrder As order)
With targetOrder
    .OrderContext = mOrderContext   ' set this first because others depend on it
    .status = mStatus               ' ditto
    
    .averagePrice = mAveragePrice
    .action = mAction
    If isAttributeModifiable(OrderAttributeIds.OrderAttAllOrNone) Then .allOrNone = mAllOrNone
    If isAttributeModifiable(OrderAttributeIds.OrderAttBlockOrder) Then .blockOrder = mBlockOrder
    .brokerId = mBrokerId
    If isAttributeModifiable(OrderAttributeIds.OrderAttDiscretionaryAmount) Then .discretionaryAmount = mDiscretionaryAmount
    If isAttributeModifiable(OrderAttributeIds.OrderAttDisplaySize) Then .displaySize = mDisplaySize
    .IsSimulated = mIsSimulated
    If isAttributeModifiable(OrderAttributeIds.OrderAttETradeOnly) Then .eTradeOnly = mETradeOnly
    .fillTime = mFillTime
    If isAttributeModifiable(OrderAttributeIds.OrderAttFirmQuoteOnly) Then .firmQuoteOnly = mFirmQuoteOnly
    If isAttributeModifiable(OrderAttributeIds.OrderAttGoodAfterTime) Then .goodAfterTime = mGoodAfterTime
    If isAttributeModifiable(OrderAttributeIds.OrderAttGoodAfterTimeTZ) Then .goodAfterTimeTZ = mGoodAfterTimeTZ
    If isAttributeModifiable(OrderAttributeIds.OrderAttGoodTillDate) Then .goodTillDate = mGoodTillDate
    If isAttributeModifiable(OrderAttributeIds.OrderAttGoodTillDateTZ) Then .goodTillDateTZ = mGoodTillDateTZ
    If isAttributeModifiable(OrderAttributeIds.OrderAttHidden) Then .hidden = mHidden
    .id = mIdentifier
    If isAttributeModifiable(OrderAttributeIds.OrderAttIgnoreRTH) Then .ignoreRegularTradingHours = mIgnoreRegularTradingHours
    .lastFillPrice = mLastFillPrice
    If isAttributeModifiable(OrderAttributeIds.OrderAttLimitPrice) Then .limitPrice = mLimitPrice
    If isAttributeModifiable(OrderAttributeIds.OrderAttMinimumQuantity) Then .minimumQuantity = mMinimumQuantity
    If isAttributeModifiable(OrderAttributeIds.OrderAttNBBOPriceCap) Then .NbboPriceCap = mNBBOPriceCap
    .offset = mOffset
    If isAttributeModifiable(OrderAttributeIds.OrderAttOrderType) Then .orderType = mOrderType
    If isAttributeModifiable(OrderAttributeIds.OrderAttOrigin) Then .origin = mOrigin
    If isAttributeModifiable(OrderAttributeIds.OrderAttOriginatorRef) Then .originatorRef = mOriginatorRef
    If isAttributeModifiable(OrderAttributeIds.OrderAttOverrideConstraints) Then .overrideConstraints = mOverrideConstraints
    If isAttributeModifiable(OrderAttributeIds.OrderAttPercentOffset) Then .percentOffset = mPercentOffset
    If isAttributeModifiable(OrderAttributeIds.OrderAttQuantity) Then .quantity = mQuantity
    .quantityFilled = mQuantityFilled
    If isAttributeModifiable(OrderAttributeIds.OrderAttSettlingFirm) Then .settlingFirm = mSettlingFirm
    If isAttributeModifiable(OrderAttributeIds.OrderAttStopTriggerMethod) And _
        mStopTriggerMethod <> StopTriggerMethods.StopTriggerNone Then .StopTriggerMethod = mStopTriggerMethod
    If isAttributeModifiable(OrderAttributeIds.OrderAttSweepToFill) Then .SweepToFill = mSweepToFill
    .Ticker = mTicker
    If isAttributeModifiable(OrderAttributeIds.OrderAttTimeInForce) Then .timeInForce = mTimeInForce
    If isAttributeModifiable(OrderAttributeIds.OrderAttTriggerPrice) Then .triggerPrice = mTriggerPrice
End With
End Sub

Friend Sub SyncToOrderReport( _
                ByVal OrderReport As IOrderReport)
Dim changed As Boolean

With OrderReport
    If .action <> mAction Then mAction = .action: changed = True
    If .allOrNone <> mAllOrNone Then mAllOrNone = .allOrNone: changed = True
    If .blockOrder <> mBlockOrder Then mBlockOrder = .blockOrder: changed = True
    If .discretionaryAmount <> mDiscretionaryAmount Then mDiscretionaryAmount = .discretionaryAmount: changed = True
    If .displaySize <> mDisplaySize Then mDisplaySize = .displaySize: changed = True
    If .eTradeOnly <> mETradeOnly Then mETradeOnly = .eTradeOnly: changed = True
    If .firmQuoteOnly <> mFirmQuoteOnly Then mFirmQuoteOnly = .firmQuoteOnly: changed = True
    If .goodAfterTime <> mGoodAfterTime Then mGoodAfterTime = .goodAfterTime: changed = True
    If .goodAfterTimeTZ <> mGoodAfterTimeTZ Then mGoodAfterTimeTZ = .goodAfterTimeTZ: changed = True
    If .goodTillDate <> mGoodTillDate Then mGoodTillDate = .goodTillDate: changed = True
    If .goodTillDateTZ <> mGoodTillDateTZ Then mGoodTillDateTZ = .goodTillDateTZ: changed = True
    If .hidden <> mHidden Then mHidden = .hidden: changed = True
    If .ignoreRegularTradingHours <> mIgnoreRegularTradingHours Then mIgnoreRegularTradingHours = .ignoreRegularTradingHours: changed = True
    If .limitPrice <> mLimitPrice Then mLimitPrice = .limitPrice: changed = True
    If .minimumQuantity <> mMinimumQuantity Then mMinimumQuantity = .minimumQuantity: changed = True
    If .NbboPriceCap <> mNBBOPriceCap Then mNBBOPriceCap = .NbboPriceCap: changed = True
    If .orderType <> mOrderType Then mOrderType = .orderType: changed = True
    If .origin <> mOrigin Then mOrigin = .origin: changed = True
    If .originatorRef <> mOriginatorRef Then mOriginatorRef = .originatorRef: changed = True
    If .overrideConstraints <> mOverrideConstraints Then mOverrideConstraints = .overrideConstraints: changed = True
    Set mProviderProperties = .providerProperties
    If .quantity <> mQuantity Then mQuantity = .quantity: changed = True
    If .settlingFirm <> mSettlingFirm Then mSettlingFirm = .settlingFirm: changed = True
    If .StopTriggerMethod <> mStopTriggerMethod Then mStopTriggerMethod = .StopTriggerMethod: changed = True
    If .SweepToFill <> mSweepToFill Then mSweepToFill = .SweepToFill: changed = True
    If .timeInForce <> mTimeInForce Then mTimeInForce = .timeInForce: changed = True
    If .tradeBuildId <> mIdentifier Then mIdentifier = .tradeBuildId: changed = True
    If .triggerPrice <> mTriggerPrice Then mTriggerPrice = .triggerPrice: changed = True
End With

If changed Then RaiseEvent PropertyChanged
End Sub

''
' Notifies the specified <code>order</code> of a change of status
'
' @param pStatus    a member of the <code>OrderStatuses</code> enumeration
'                   specifying the <code>order</code>'s new status
'
'@/
Friend Sub updateStatus(ByVal pStatus As OrderStatuses)

If pStatus <> mStatus Then
    mStatus = pStatus
    RaiseEvent StatusChanged
End If
End Sub

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

Private Sub checkAttributeModificationPermitted( _
                ByVal attributeId As OrderAttributeIds)
If Not isAttributeModifiable(attributeId) Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                "TradeBuild26.Order", _
                "The " & gOrderAttributeToString(attributeId) & " attribute is not currently modifiable"
End If
End Sub

''
' Set the 'dirty' flag, and raise a Dirty event if one has not already
' been raised.
'@/
Private Sub setDirty()
If Not mIsDirty Then
    mIsDirty = True
    RaiseEvent Dirty
End If
End Sub





