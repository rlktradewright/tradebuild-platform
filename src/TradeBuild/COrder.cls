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

Private mContract       As TradeBuild.Contract

Private mMarketIndex    As Long
Private mAtExchange     As Boolean
Private mSimulated      As Boolean
Private mStatus         As OrderStatuses
Private mIdentifier     As Long
Private mClientID       As Long
Private mPermId         As Long
    
Private mAction         As OrderActions
Private mQuantity       As Long
Private mQuantityFilled As Long
Private mAveragePrice   As Double
Private mLastFillPrice  As Double
Private mOrderType      As TradeBuild.OrderTypes
Private mLimitPrice     As Double
Private mAuxPrice       As Double
    
    ' extended order fields
Private mTif            As OrderTifs
Private mOcaGroup       As String
Private mAccount        As String
Private mOpenClose      As String
Private mOrigin         As OrderOrigins
Private mOrderRef       As String
Private mTransmit       As Boolean
Private mParentId       As Long
Private mBlockOrder     As Boolean
Private mSweepToFill    As Boolean
Private mDisplaySize    As Long
Private mTriggerMethod  As TriggerMethods
Private mIgnoreRth      As Boolean
Private mHidden         As Boolean
Private mDiscretionaryAmt As Double
Private mGoodAfterTime  As String
Private mGoodTillDate   As String
    
    ' share allocation
Private mSharesAllocation As String
    
    'fa
Private mFAGroup        As String
Private mFAProfile      As String
Private mFAMethod       As String
Private mFAPercentage   As String
    
Private Sub Class_Initialize()
mStatus = OrderStatuses.OrderStatusCreated
mAction = ActionBuy
mOrderType = OrderTypeNone
mTif = TIFDay
End Sub

Public Property Let account(ByVal value As String)
mAccount = value
End Property

Public Property Get account() As String
account = mAccount
End Property

Public Property Let action(ByVal value As OrderActions)
mAction = UCase$(value)
End Property

Public Property Get action() As OrderActions
action = mAction
End Property

Friend Property Let atExchange(ByVal value As Boolean)
mAtExchange = value
End Property

Friend Property Get atExchange() As Boolean
atExchange = mAtExchange
End Property

Public Property Let auxPrice(ByVal value As Double)
mAuxPrice = value
End Property

Public Property Get auxPrice() As Double
auxPrice = mAuxPrice
End Property

Public Property Let averagePrice(ByVal value As Double)
mAveragePrice = value
End Property

Public Property Get averagePrice() As Double
averagePrice = mAveragePrice
End Property

Public Property Let blockOrder(value As Boolean)
mBlockOrder = value
End Property

Public Property Get blockOrder() As Boolean
blockOrder = mBlockOrder
End Property

Friend Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Public Property Get clientID() As Long
clientID = mClientID
End Property

Friend Property Let Contract(ByVal value As TradeBuild.Contract)
Set mContract = value
End Property

Public Property Get Contract() As TradeBuild.Contract
Set Contract = mContract
End Property

Public Property Let displaySize(ByVal value As Long)
mDisplaySize = value
End Property

Public Property Get displaySize() As Long
displaySize = mDisplaySize
End Property

Public Property Let discretionaryAmt(ByVal value As Double)
mDiscretionaryAmt = value
End Property

Public Property Get discretionaryAmt() As Double
discretionaryAmt = mDiscretionaryAmt
End Property

Public Property Let FAGroup(ByVal value As String)
mFAGroup = value
End Property

Public Property Get FAGroup() As String
FAGroup = mFAGroup
End Property

Public Property Let FAMethod(ByVal value As String)
mFAMethod = value
End Property

Public Property Get FAMethod() As String
FAMethod = mFAMethod
End Property

Public Property Let FAPercentage(ByVal value As String)
mFAPercentage = value
End Property

Public Property Get FAPercentage() As String
FAPercentage = mFAPercentage
End Property

Public Property Let FAProfile(ByVal value As String)
mFAProfile = value
End Property

Public Property Get FAProfile() As String
FAProfile = mFAProfile
End Property

Public Property Let goodAfterTime(ByVal value As String)
mGoodAfterTime = value
End Property

Public Property Get goodAfterTime() As String
goodAfterTime = mGoodAfterTime
End Property

Public Property Let goodTillDate(ByVal value As String)
mGoodTillDate = value
End Property

Public Property Get goodTillDate() As String
goodTillDate = mGoodTillDate
End Property

Public Property Let hidden(value As Boolean)
mHidden = value
End Property

Public Property Get hidden() As Boolean
hidden = mHidden
End Property

Friend Property Let id(ByVal value As Long)
mIdentifier = value
End Property

Public Property Get id() As Long
id = mIdentifier
End Property

Public Property Let ignoreRTH(value As Boolean)
mIgnoreRth = value
End Property

Public Property Get ignoreRTH() As Boolean
ignoreRTH = mIgnoreRth
End Property

Public Property Let lastFillPrice(ByVal value As Double)
mLastFillPrice = value
End Property

Public Property Get lastFillPrice() As Double
lastFillPrice = mLastFillPrice
End Property

Public Property Let limitPrice(ByVal value As Double)
mLimitPrice = value
End Property

Public Property Get limitPrice() As Double
limitPrice = mLimitPrice
End Property

Friend Property Let marketIndex(ByVal value As Long)
mMarketIndex = value
End Property

Friend Property Get marketIndex() As Long
marketIndex = mMarketIndex
End Property

Public Property Let ocaGroup(ByVal value As String)
mOcaGroup = value
End Property

Public Property Get ocaGroup() As String
ocaGroup = mOcaGroup
End Property

Public Property Let openClose(ByVal value As String)
mOpenClose = value
End Property

Public Property Get openClose() As String
openClose = mOpenClose
End Property

Public Property Let orderRef(ByVal value As String)
mOrderRef = value
End Property

Public Property Get orderRef() As String
orderRef = mOrderRef
End Property

Public Property Let orderType(ByVal value As TradeBuild.OrderTypes)
Select Case value
Case OrderTypeNone
Case OrderTypeMarket
Case OrderTypeMarketOnClose
Case OrderTypeLimit
Case OrderTypeLimitOnClose
Case OrderTypePeggedToMarket
Case OrderTypeStop
Case OrderTypeStopLimit
Case OrderTypeTrail
Case OrderTypeRelative
Case OrderTypeVWAP
Case OrderTypeMarketToLimit
Case OrderTypeQuote
Case OrderTypeAutoStop
Case OrderTypeAutoLimit
Case OrderTypeAdjust
Case OrderTypeAlert
Case OrderTypeLimitIfTouched
Case OrderTypeMarketIfTouched
Case OrderTypeTrailLimit
Case OrderTypeMarketWithProtection
Case OrderTypeMarketOnOpen
Case OrderTypeLimitOnOpen
Case OrderTypePeggedToPrimary
Case Else
    err.Raise ErrorCodes.InvalidOrderType, _
            "TradeBuild.Order::(let)orderType", _
            "Invalid order type " & value
End Select
mOrderType = value
End Property

Public Property Get orderType() As TradeBuild.OrderTypes
orderType = mOrderType
End Property

Public Property Let origin(ByVal value As OrderOrigins)
mOrigin = value
End Property

Public Property Get origin() As OrderOrigins
origin = mOrigin
End Property

Public Property Let permId(value As Long)
mPermId = value
End Property

Public Property Get permId() As Long
permId = mPermId
End Property

Public Property Let parentId(ByVal value As Long)
mParentId = value
End Property

Public Property Get parentId() As Long
parentId = mParentId
End Property

Public Property Let quantity(ByVal value As Long)
mQuantity = value
End Property

Public Property Get quantity() As Long
quantity = mQuantity
End Property

Public Property Let quantityFilled(ByVal value As Long)
mQuantityFilled = value
End Property

Public Property Get quantityFilled() As Long
quantityFilled = mQuantityFilled
End Property

Public Property Let sharesAllocation(value As String)
mSharesAllocation = value
End Property

Public Property Get sharesAllocation() As String
sharesAllocation = mSharesAllocation
End Property

Public Property Let simulated(ByVal value As Boolean)
mSimulated = value
End Property

Public Property Get simulated() As Boolean
simulated = mSimulated
End Property

Public Property Let status(value As OrderStatuses)
mStatus = value
End Property

Public Property Get status() As OrderStatuses
status = mStatus
End Property

Public Property Let sweepToFill(value As Boolean)
mSweepToFill = value
End Property

Public Property Get sweepToFill() As Boolean
sweepToFill = mSweepToFill
End Property

Public Property Let timeInForce(ByVal value As OrderTifs)
mTif = value
End Property

Public Property Get timeInForce() As OrderTifs
timeInForce = mTif
End Property

Public Property Let transmit(value As Boolean)
mTransmit = value
End Property

Public Property Get transmit() As Boolean
transmit = mTransmit
End Property

Public Property Let triggerMethod(value As TriggerMethods)
mTriggerMethod = value
End Property

Public Property Get triggerMethod() As TriggerMethods
triggerMethod = mTriggerMethod
End Property

Public Function clone() As Order
Set clone = New Order
With clone
    .Contract = mContract
    .marketIndex = mMarketIndex
    .simulated = mSimulated
    .status = mStatus
    .id = mIdentifier
    .clientID = mClientID
    .permId = mPermId
    .action = mAction
    .quantity = mQuantity
    .quantityFilled = mQuantityFilled
    .averagePrice = mAveragePrice
    .lastFillPrice = mLastFillPrice
    .orderType = mOrderType
    .limitPrice = mLimitPrice
    .auxPrice = mAuxPrice
    .timeInForce = mTif
    .ocaGroup = mOcaGroup
    .account = mAccount
    .openClose = mOpenClose
    .origin = mOrigin
    .orderRef = mOrderRef
    .transmit = mTransmit
    .parentId = mParentId
    .blockOrder = mBlockOrder
    .sweepToFill = mSweepToFill
    .displaySize = mDisplaySize
    .triggerMethod = mTriggerMethod
    .ignoreRTH = mIgnoreRth
    .hidden = mHidden
    .discretionaryAmt = mDiscretionaryAmt
    .goodAfterTime = mGoodAfterTime
    .sharesAllocation = mSharesAllocation
    .FAGroup = mFAGroup
    .FAProfile = mFAProfile
    .FAMethod = mFAMethod
    .FAPercentage = mFAPercentage
End With
End Function

Friend Sub fill(ByVal quantity As Long, ByVal price As Double)
mLastFillPrice = price
mAveragePrice = (mQuantityFilled * mAveragePrice + quantity * price) / (mQuantityFilled + quantity)
mQuantityFilled = mQuantityFilled + quantity
mQuantity = mQuantity - quantity
End Sub








