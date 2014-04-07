VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsOrder"
Attribute VB_GlobalNameSpace = False
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

Private Const ModuleName                            As String = "TwsOrder"

'@================================================================================
' Member variables
'@================================================================================

Private Const CUSTOMER = 0
Private Const FIRM = 1

' main Order fields
Private mOrderId                                    As Long
Private mClientID                                   As Long
Private mPermId                                     As Long
Private mAction                                     As TwsOrderActions
Private mTotalQuantity                              As Long
Private mOrderType                                  As TwsOrderTypes
Private mLmtPrice                                   As Double
Private mAuxPrice                                   As Double

' extended Order fields
Private mTif                                        As TwsOrderTIFs
Private mOcaGroup                                   As String ' one cancels all group name
Private mOcaType                                    As TwsOcaTypes
Private mOrderRef                                   As String
Private mTransmit                                   As Boolean  ' if false, Order will be created but not transmited
Private mParentId                                   As Long ' Parent Order Id, to associate Auto STP or TRAIL orders with the original Order.
Private mBlockOrder                                 As Boolean
Private mSweepToFill                                As Boolean
Private mDisplaySize                                As Long
Private mTriggerMethod                              As TwsStopTriggerMethods
Private mOutsideRth                                 As Boolean
Private mHidden                                     As Boolean
Private mGoodAfterTime                              As String ' FORMAT: 20060505 08:00:00 {Time zone}
Private mGoodTillDate                               As String ' FORMAT: 20060505 08:00:00 {Time zone}
Private mOverridePercentageConstraints              As Boolean
Private mRule80A                                    As String ' Individual = 'I', Agency = 'A', AgentOtherMember = 'W', IndividualPTIA = 'J', AgencyPTIA = 'U', AgentOtherMemberPTIA = 'M', IndividualPT = 'K', AgencyPT = 'Y', AgentOtherMemberPT = 'N'
Private mAllOrNone                                  As Boolean
Private mMinQty                                     As Long
Private mPercentOffset                              As Double ' REL orders only
Private mTrailStopPrice                             As Double ' for TRAILLIMIT orders only

' Financial advisors only
Private mFaGroup                                    As String
Private mFaProfile                                  As String
Private mFaMethod                                   As String
Private mFaPercentage                               As String

' Institutional orders only
Private mOpenClose                                  As String ' O=Open, C=Close
Private mOrigin                                     As Long ' 0=Customer, 1=Firm
Private mShortSaleSlot                              As Long ' 1 if you hold the shares, 2 if they will be delivered from elsewhere.  Only for Action="SSHORT
Private mDesignatedLocation                         As String ' set when slot=2 only.

' SMART routing only
Private mDiscretionaryAmt                           As Double
Private mETradeOnly                                 As Boolean
Private mFirmQuoteOnly                              As Boolean
Private mNbboPriceCap                               As Double

' BOX or VOL ORDERS ONLY
Private mAuctionStrategy                            As TwsAuctionStrategies

' BOX ORDERS ONLY
Private mStartingPrice                              As Double
Private mStockRefPrice                              As Double
Private mDelta                                      As Double
' pegged to stock or VOL orders
Private mStockRangeLower                            As Double
Private mStockRangeUpper                            As Double

' VOLATILITY ORDERS ONLY
Private mVolatility                                 As Double
Private mVolatilityType                             As Long ' 1=daily, 2=annual
Private mContinuousUpdate                           As Long
Private mReferencePriceType                         As Long ' 1=Average, 2 = BidOrAsk
Private mDeltaNeutralOrderType                      As TwsOrderTypes
Private mDeltaNeutralAuxPrice                       As Double

' COMBO ORDERS ONLY
Private mBasisPoints                                As Double ' EFP orders only
Private mBasisPointsType                            As Long ' EFP orders only

' SCALE ORDERS ONLY
Private mScaleInitLevelSize                         As Long
Private mScaleSubsLevelSize                         As Long
Private mScalePriceIncrement                        As Double

' Clearing info
Private mAccount                                    As String ' IB account
Private mSettlingFirm                               As String
Private mClearingAccount                            As String ' True beneficiary of the Order
Private mClearingIntent                             As String ' "" (Default), "IB", "Away", "PTA" (PostTrade)

' ALGO ORDERS ONLY
Private mAlgoStrategy                               As String
Private mAlgoParams()                               As TwsTagValue

' What-if
Private mWhatIf                                     As Boolean
        
' Not Held
Private mNotHeld                                    As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mOutsideRth = False
mOpenClose = "O"
mOrigin = CUSTOMER
mTransmit = True
mDesignatedLocation = ""
mMinQty = MaxLong
mPercentOffset = MaxDouble
mNbboPriceCap = MaxDouble
mStartingPrice = MaxDouble
mStockRefPrice = MaxDouble
mDelta = MaxDouble
mStockRangeLower = MaxDouble
mStockRangeUpper = MaxDouble
mVolatility = MaxDouble
mVolatilityType = MaxLong
mDeltaNeutralOrderType = TwsOrderTypeNone
mDeltaNeutralAuxPrice = MaxDouble
mReferencePriceType = MaxLong
mTrailStopPrice = MaxDouble
mBasisPoints = MaxDouble
mBasisPointsType = MaxLong
mScaleInitLevelSize = MaxLong
mScaleSubsLevelSize = MaxLong
mScalePriceIncrement = MaxDouble
mWhatIf = False
mNotHeld = False
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

Public Property Let Action(ByVal Value As TwsOrderActions)
mAction = Value
End Property

Public Property Get Action() As TwsOrderActions
Action = mAction
End Property

Public Property Let AuxPrice(ByVal Value As Double)
mAuxPrice = Value
End Property

Public Property Get AuxPrice() As Double
AuxPrice = mAuxPrice
End Property

Public Property Let ClientID(ByVal Value As Long)
mClientID = Value
End Property

Public Property Get ClientID() As Long
ClientID = mClientID
End Property

Public Property Let LmtPrice(ByVal Value As Double)
mLmtPrice = Value
End Property

Public Property Get LmtPrice() As Double
LmtPrice = mLmtPrice
End Property

Public Property Let OrderId(ByVal Value As Long)
mOrderId = Value
End Property

Public Property Get OrderId() As Long
OrderId = mOrderId
End Property

Public Property Let OrderType(ByVal Value As TwsOrderTypes)
mOrderType = Value
End Property

Public Property Get OrderType() As TwsOrderTypes
OrderType = mOrderType
End Property

Public Property Let PermId(ByVal Value As Long)
mPermId = Value
End Property

Public Property Get PermId() As Long
PermId = mPermId
End Property

Public Property Let TotalQuantity(ByVal Value As Long)
mTotalQuantity = Value
End Property

Public Property Get TotalQuantity() As Long
TotalQuantity = mTotalQuantity
End Property

Public Property Let Tif(ByVal Value As TwsOrderTIFs)
mTif = Value
End Property

Public Property Get Tif() As TwsOrderTIFs
Tif = mTif
End Property

Public Property Let OcaGroup(ByVal Value As String)
mOcaGroup = Value
End Property

Public Property Get OcaGroup() As String
OcaGroup = mOcaGroup
End Property

Public Property Let OcaType(ByVal Value As TwsOcaTypes)
mOcaType = Value
End Property

Public Property Get OcaType() As TwsOcaTypes
OcaType = mOcaType
End Property

Public Property Let OrderRef(ByVal Value As String)
mOrderRef = Value
End Property

Public Property Get OrderRef() As String
OrderRef = mOrderRef
End Property

Public Property Let Transmit(ByVal Value As Boolean)
mTransmit = Value
End Property

Public Property Get Transmit() As Boolean
Transmit = mTransmit
End Property

Public Property Let ParentId(ByVal Value As Long)
mParentId = Value
End Property

Public Property Get ParentId() As Long
ParentId = mParentId
End Property

Public Property Let BlockOrder(ByVal Value As Boolean)
mBlockOrder = Value
End Property

Public Property Get BlockOrder() As Boolean
BlockOrder = mBlockOrder
End Property

Public Property Let SweepToFill(ByVal Value As Boolean)
mSweepToFill = Value
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSweepToFill
End Property

Public Property Let DisplaySize(ByVal Value As Long)
mDisplaySize = Value
End Property

Public Property Get DisplaySize() As Long
DisplaySize = mDisplaySize
End Property

Public Property Let TriggerMethod(ByVal Value As Long)
mTriggerMethod = Value
End Property

Public Property Get TriggerMethod() As Long
TriggerMethod = mTriggerMethod
End Property

Public Property Let OutsideRth(ByVal Value As Boolean)
mOutsideRth = Value
End Property

Public Property Get OutsideRth() As Boolean
OutsideRth = mOutsideRth
End Property

Public Property Let Hidden(ByVal Value As Boolean)
mHidden = Value
End Property

Public Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Public Property Let GoodAfterTime(ByVal Value As String)
mGoodAfterTime = Value
End Property

Public Property Get GoodAfterTime() As String
GoodAfterTime = mGoodAfterTime
End Property

Public Property Let GoodTillDate(ByVal Value As String)
mGoodTillDate = Value
End Property

Public Property Get GoodTillDate() As String
GoodTillDate = mGoodTillDate
End Property

Public Property Let OverridePercentageConstraints(ByVal Value As Boolean)
mOverridePercentageConstraints = Value
End Property

Public Property Get OverridePercentageConstraints() As Boolean
OverridePercentageConstraints = mOverridePercentageConstraints
End Property

Public Property Let Rule80A(ByVal Value As String)
mRule80A = Value
End Property

Public Property Get Rule80A() As String
Rule80A = mRule80A
End Property

Public Property Let AllOrNone(ByVal Value As Boolean)
mAllOrNone = Value
End Property

Public Property Get AllOrNone() As Boolean
AllOrNone = mAllOrNone
End Property

Public Property Let MinQty(ByVal Value As Long)
mMinQty = Value
End Property

Public Property Get MinQty() As Long
MinQty = mMinQty
End Property

Public Property Let PercentOffset(ByVal Value As Double)
mPercentOffset = Value
End Property

Public Property Get PercentOffset() As Double
PercentOffset = mPercentOffset
End Property

Public Property Let TrailStopPrice(ByVal Value As Double)
mTrailStopPrice = Value
End Property

Public Property Get TrailStopPrice() As Double
TrailStopPrice = mTrailStopPrice
End Property

Public Property Let FaGroup(ByVal Value As String)
mFaGroup = Value
End Property

Public Property Get FaGroup() As String
FaGroup = mFaGroup
End Property

Public Property Let FaProfile(ByVal Value As String)
mFaProfile = Value
End Property

Public Property Get FaProfile() As String
FaProfile = mFaProfile
End Property

Public Property Let FaMethod(ByVal Value As String)
mFaMethod = Value
End Property

Public Property Get FaMethod() As String
FaMethod = mFaMethod
End Property

Public Property Let FaPercentage(ByVal Value As String)
mFaPercentage = Value
End Property

Public Property Get FaPercentage() As String
FaPercentage = mFaPercentage
End Property

Public Property Let OpenClose(ByVal Value As String)
mOpenClose = Value
End Property

Public Property Get OpenClose() As String
OpenClose = mOpenClose
End Property

Public Property Let Origin(ByVal Value As Long)
mOrigin = Value
End Property

Public Property Get Origin() As Long
Origin = mOrigin
End Property

Public Property Let ShortSaleSlot(ByVal Value As Long)
mShortSaleSlot = Value
End Property

Public Property Get ShortSaleSlot() As Long
ShortSaleSlot = mShortSaleSlot
End Property

Public Property Let DesignatedLocation(ByVal Value As String)
mDesignatedLocation = Value
End Property

Public Property Get DesignatedLocation() As String
DesignatedLocation = mDesignatedLocation
End Property

Public Property Let DiscretionaryAmt(ByVal Value As Double)
mDiscretionaryAmt = Value
End Property

Public Property Get DiscretionaryAmt() As Double
DiscretionaryAmt = mDiscretionaryAmt
End Property

Public Property Let ETradeOnly(ByVal Value As Boolean)
mETradeOnly = Value
End Property

Public Property Get ETradeOnly() As Boolean
ETradeOnly = mETradeOnly
End Property

Public Property Let FirmQuoteOnly(ByVal Value As Boolean)
mFirmQuoteOnly = Value
End Property

Public Property Get FirmQuoteOnly() As Boolean
FirmQuoteOnly = mFirmQuoteOnly
End Property

Public Property Let NbboPriceCap(ByVal Value As Double)
mNbboPriceCap = Value
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mNbboPriceCap
End Property

Public Property Let AuctionStrategy(ByVal Value As TwsAuctionStrategies)
mAuctionStrategy = Value
End Property

Public Property Get AuctionStrategy() As TwsAuctionStrategies
AuctionStrategy = mAuctionStrategy
End Property

Public Property Let StartingPrice(ByVal Value As Double)
mStartingPrice = Value
End Property

Public Property Get StartingPrice() As Double
StartingPrice = mStartingPrice
End Property

Public Property Let StockRefPrice(ByVal Value As Double)
mStockRefPrice = Value
End Property

Public Property Get StockRefPrice() As Double
StockRefPrice = mStockRefPrice
End Property

Public Property Let Delta(ByVal Value As Double)
mDelta = Value
End Property

Public Property Get Delta() As Double
Delta = mDelta
End Property

Public Property Let StockRangeLower(ByVal Value As Double)
mStockRangeLower = Value
End Property

Public Property Get StockRangeLower() As Double
StockRangeLower = mStockRangeLower
End Property

Public Property Let StockRangeUpper(ByVal Value As Double)
mStockRangeUpper = Value
End Property

Public Property Get StockRangeUpper() As Double
StockRangeUpper = mStockRangeUpper
End Property

Public Property Let Volatility(ByVal Value As Double)
mVolatility = Value
End Property

Public Property Get Volatility() As Double
Volatility = mVolatility
End Property

Public Property Let VolatilityType(ByVal Value As Long)
mVolatilityType = Value
End Property

Public Property Get VolatilityType() As Long
VolatilityType = mVolatilityType
End Property

Public Property Let ContinuousUpdate(ByVal Value As Long)
mContinuousUpdate = Value
End Property

Public Property Get ContinuousUpdate() As Long
ContinuousUpdate = mContinuousUpdate
End Property

Public Property Let ReferencePriceType(ByVal Value As Long)
mReferencePriceType = Value
End Property

Public Property Get ReferencePriceType() As Long
ReferencePriceType = mReferencePriceType
End Property

Public Property Let DeltaNeutralOrderType(ByVal Value As TwsOrderTypes)
mDeltaNeutralOrderType = Value
End Property

Public Property Get DeltaNeutralOrderType() As TwsOrderTypes
DeltaNeutralOrderType = mDeltaNeutralOrderType
End Property

Public Property Let DeltaNeutralAuxPrice(ByVal Value As Double)
mDeltaNeutralAuxPrice = Value
End Property

Public Property Get DeltaNeutralAuxPrice() As Double
DeltaNeutralAuxPrice = mDeltaNeutralAuxPrice
End Property

Public Property Let BasisPoints(ByVal Value As Double)
mBasisPoints = Value
End Property

Public Property Get BasisPoints() As Double
BasisPoints = mBasisPoints
End Property

Public Property Let BasisPointsType(ByVal Value As Long)
mBasisPointsType = Value
End Property

Public Property Get BasisPointsType() As Long
BasisPointsType = mBasisPointsType
End Property

Public Property Let ScaleInitLevelSize(ByVal Value As Long)
mScaleInitLevelSize = Value
End Property

Public Property Get ScaleInitLevelSize() As Long
ScaleInitLevelSize = mScaleInitLevelSize
End Property

Public Property Let ScaleSubsLevelSize(ByVal Value As Long)
mScaleSubsLevelSize = Value
End Property

Public Property Get ScaleSubsLevelSize() As Long
ScaleSubsLevelSize = mScaleSubsLevelSize
End Property

Public Property Let ScalePriceIncrement(ByVal Value As Double)
mScalePriceIncrement = Value
End Property

Public Property Get ScalePriceIncrement() As Double
ScalePriceIncrement = mScalePriceIncrement
End Property

Public Property Let Account(ByVal Value As String)
mAccount = Value
End Property

Public Property Get Account() As String
Account = mAccount
End Property

Public Property Let SettlingFirm(ByVal Value As String)
mSettlingFirm = Value
End Property

Public Property Get SettlingFirm() As String
SettlingFirm = mSettlingFirm
End Property

Public Property Let ClearingAccount(ByVal Value As String)
mClearingAccount = Value
End Property

Public Property Get ClearingAccount() As String
ClearingAccount = mClearingAccount
End Property

Public Property Let ClearingIntent(ByVal Value As String)
mClearingIntent = Value
End Property

Public Property Get ClearingIntent() As String
ClearingIntent = mClearingIntent
End Property

Public Property Let AlgoStrategy(ByVal Value As String)
mAlgoStrategy = Value
End Property

Public Property Get AlgoStrategy() As String
AlgoStrategy = mAlgoStrategy
End Property

Public Property Let AlgoParams(ByRef Value() As TwsTagValue)
mAlgoParams = Value
End Property

Public Property Get AlgoParams() As TwsTagValue()
AlgoParams = mAlgoParams
End Property

Public Property Let WhatIf(ByVal Value As Boolean)
mWhatIf = Value
End Property

Public Property Get WhatIf() As Boolean
WhatIf = mWhatIf
End Property

Public Property Let NotHeld(ByVal Value As Boolean)
mNotHeld = Value
End Property

Public Property Get NotHeld() As Boolean
NotHeld = mNotHeld
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

