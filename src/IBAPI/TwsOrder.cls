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

Private Const CUSTOMER = 0
Private Const FIRM = 1
Private Const OPT_UNKNOWN = "?"
Private Const OPT_BROKER_DEALER = "b"
Private Const OPT_CUSTOMER = "c"
Private Const OPT_FIRM = "f"
Private Const OPT_ISEMM = "m"
Private Const OPT_FARMM = "n"
Private Const OPT_SPECIALIST = "y"
Private Const AUCTION_MATCH = 1
Private Const AUCTION_IMPROVEMENT = 2
Private Const AUCTION_TRANSPARENT = 3

'@================================================================================
' Member variables
'@================================================================================

' main Order fields
Public OrderId                                      As Long
Public ClientID                                     As Long
Public PermId                                       As Long
Public Action                                       As TwsOrderActions
Public TotalQuantity                                As BoxedDecimal
Public OrderType                                    As TwsOrderTypes
Public LmtPrice                                     As Double
Public AuxPrice                                     As Double

' extended Order fields
Public Tif                                          As TwsOrderTIFs
Public ActiveStartTime                              As String   ' GTC orders
Public ActiveStopTime                               As String   ' GTC orders
Public OcaGroup                                     As String ' one cancels all group name
Public OcaType                                      As TwsOcaTypes
Public OrderRef                                     As String
Public Transmit                                     As Boolean  ' if false, Order will be created but not transmited
Public ParentId                                     As Long ' Parent Order Id, to associate Auto STP or TRAIL orders with the original Order.
Public BlockOrder                                   As Boolean
Public SweepToFill                                  As Boolean
Public DisplaySize                                  As Long
Public TriggerMethod                                As TwsStopTriggerMethods
Public OutsideRth                                   As Boolean
Public Hidden                                       As Boolean
Public GoodAfterTime                                As String ' FORMAT: 20060505 08:00:00 {Time zone}
Public GoodTillDate                                 As String ' FORMAT: 20060505 08:00:00 {Time zone}
Public OverridePercentageConstraints                As Boolean
Public Rule80A                                      As String ' Individual = 'I', Agency = 'A', AgentOtherMember = 'W', IndividualPTIA = 'J', AgencyPTIA = 'U', AgentOtherMemberPTIA = 'M', IndividualPT = 'K', AgencyPT = 'Y', AgentOtherMemberPT = 'N'
Public AllOrNone                                    As Boolean
Public MinQty                                       As Long
Public PercentOffset                                As Double ' REL orders only
Public TrailStopPrice                               As Double ' for TRAILLIMIT orders only
Public TrailingPercent                              As Double ' specify the percentage, e.g. 3, not .03

' Financial advisors only
Public FaGroup                                      As String
Public FaMethod                                     As String
Public FaPercentage                                 As String

' Institutional orders only
Public OpenClose                                    As String ' O=Open, C=Close
Public Origin                                       As Long ' 0=Customer, 1=Firm
Public ShortSaleSlot                                As Long ' 1 if you hold the shares, 2 if they will be delivered from elsewhere.  Only for Action="SSHORT
Public DesignatedLocation                           As String ' set when slot=2 only.
Public ExemptCode                                   As Long

' SMART routing only
Public DiscretionaryAmt                             As Double

Public OptOutSmartRouting                           As Boolean

' BOX or VOL ORDERS ONLY
Public AuctionStrategy                              As TwsAuctionStrategies

' BOX ORDERS ONLY
Public StartingPrice                                As Double
Public StockRefPrice                                As Double
Public Delta                                        As Double

' pegged to stock or VOL orders
Public StockRangeLower                              As Double
Public StockRangeUpper                              As Double

' VOLATILITY ORDERS ONLY
Public Volatility                                   As Double   ' enter percentage not decimal, e.g. 2 not .02
Public VolatilityType                               As Long ' 1=daily, 2=annual
Public ContinuousUpdate                             As Long
Public ReferencePriceType                           As Long ' 1=Average, 2 = BidOrAsk
Public DeltaNeutralOrderType                        As TwsOrderTypes
Public DeltaNeutralAuxPrice                         As Double
Public DeltaNeutralConId                            As Long
Public DeltaNeutralSettlingFirm                     As String
Public DeltaNeutralClearingAccount                  As String
Public DeltaNeutralClearingIntent                   As String
Public DeltaNeutralOpenClose                        As String
Public DeltaNeutralShortSale                        As Boolean
Public DeltaNeutralShortSaleSlot                    As Long
Public DeltaNeutralDesignatedLocation               As String

' COMBO ORDERS ONLY
Public BasisPoints                                  As Double ' EFP orders only
Public BasisPointsType                              As Long ' EFP orders only

' SCALE ORDERS ONLY
Public ScaleInitLevelSize                           As Long
Public ScaleSubsLevelSize                           As Long
Public ScalePriceIncrement                          As Double
Public ScalePriceAdjustValue                        As Double
Public ScalePriceAdjustInterval                     As Long
Public ScaleProfitOffset                            As Double
Public ScaleAutoReset                               As Boolean
Public ScaleInitPosition                            As Long
Public ScaleInitFillQty                             As Long
Public ScaleRandomPercent                           As Boolean
Public ScaleTable                                   As String

' HEDGE ORDERS ONLY
Public HedgeType                                    As TwsHedgeTypes    ' 'D' - delta, 'B' - beta, 'F' - FX, 'P' - pair
Public HedgeParam                                   As String   ' beta value for beta hedge (in range 0-1), ratio for pair hedge

' Clearing info
Public Account                                      As String ' IB account
Public SettlingFirm                                 As String
Public ClearingAccount                              As String ' True beneficiary of the Order
Public ClearingIntent                               As String ' "" (Default), "IB", "Away", "PTA" (PostTrade)

' ALGO ORDERS ONLY
Public AlgoStrategy                                 As String
Private mAlgoParams                                 As Parameters

' What-if
Public WhatIf                                       As Boolean

Public AlgoId                                       As String

' Not Held
Public NotHeld                                      As Boolean

' Smart combo routing params
Private mSmartComboRoutingParams                    As Parameters

' order combo legs
Private mComboLegs()                                As TwsOrderComboLeg

Private mOptions                                    As Parameters

Public Solicited As Boolean
Public ModelCode As String

''' <summary>
''' A regulatory attribute that applies to all US Commodity (Futures) Exchanges,
''' provided to allow client to comply with CFTC Tag 50 Rules
''' </summary>
''' <returns></returns>
Public ExtOperator As String
Public CashQty As Double
Public Mifid2DecisionMaker As String
Public Mifid2DecisionAlgo As String
Public Mifid2ExecutionTrader As String
Public Mifid2ExecutionAlgo As String

Public DontUseAutoPriceForHedge As Boolean
Public RandomizeSize As Boolean
Public RandomizePrice As Boolean
    
Public AutoCancelDate As String
Public FilledQuantity As BoxedDecimal

Public ReferenceFutureContractId As Long
Public AutoCancelParent As Boolean
Public Shareholder As String
Public ImbalanceOnly As Boolean
Public RouteMarketableToBbo As Boolean
Public ParentPermanentId As BoxedDecimal
Public AdvancedErrorOverride As String
Public ManualOrderTime As String
Public MinimumTradeQuantity As Long
Public MinimumCompeteSize As Long
Public CompeteAgainstBestOffset As Double
Public MidOffsetAtWhole As Double
Public MidOffsetAtHalf As Double

''' <summary>
''' For Pegged-to-benchmark orders, contains the Contract Id Of the contract against which the order will be pegged.
''' </summary>
''' <returns></returns>
Public ReferenceContractId As Long

''' <summary>
''' For Pegged-to-benchmark orders, indicates whether the order's pegged price should increase or decreases.
''' </summary>
''' <returns></returns>
Public IsPeggedChangeAmountDecrease As Boolean

''' <summary>
''' For Pegged-to-benchmark orders, specifies the amount by which the order's pegged price should move.
''' </summary>
''' <returns></returns>
Public PeggedChangeAmount As Double

''' <summary>
''' For Pegged-to-benchmark orders, specifies the amount the reference contract needs to move to adjust the pegged order.
''' </summary>
''' <returns></returns>
Public ReferenceChangeAmount As Double

''' <summary>
''' for Pegged-to-benchmark orders, specifies the exchange for the reference contract.
''' </summary>
''' <returns></returns>
Public ReferenceExchange As String

''' <summary>
''' For Adjusted Stop orders, the parent order will be adjusted to the specified type When the adjusted trigger price is penetrated.
''' </summary>
''' <returns></returns>
Public AdjustedOrderType As TwsOrderTypes

''' <summary>
''' - DOC_TODO
''' </summary>
''' <returns></returns>
Public TriggerPrice As Double

''' <summary>
''' - DOC_TODO
''' </summary>
''' <returns></returns>
Public LmtPriceOffset As Double

''' <summary>
''' For Adjusted Stop orders, specifies the Stop price Of the adjusted (STP) parent.
''' </summary>
''' <returns></returns>
Public AdjustedStopPrice As Double

''' <summary>
''' For Adjusted Stop orders, specifies the Stop limit price Of the adjusted (STPL LMT) parent.
''' </summary>
''' <returns></returns>
Public AdjustedStopLimitPrice As Double

''' <summary>
''' For Adjusted Stop orders, specifies the trailing amount Of the adjusted (TRAIL) parent.
''' </summary>
''' <returns></returns>
Public AdjustedTrailingAmount As Double

''' <summary>
''' For Adjusted Stop orders, specifies whether the trailing unit is an amount (set To 0) or a percentage (set To 1).
''' </summary>
''' <returns></returns>
Public AdjustableTrailingUnit As Long

' Conditions not implemented
'''' <summary>
'''' Conditions that determine when the order will be activated or canceled.
'''' </summary>
'''' <returns></returns>
'Private mConditions() As Object ' TwsOrderCondition
'
'''' <summary>
'''' Indicates whether conditions will also be valid outside Regular Trading Hours.
'''' </summary>
'''' <returns></returns>
'Public ConditionsIgnoreRth As Boolean
'
'''' <summary>
'''' Specifies whether conditions determine if the order should become active or be canceled.
'''' </summary>
'''' <returns></returns>
'Public ConditionsCancelOrder As Boolean

''' <summary>
''' Define the Soft Dollar Tier used for the order. Only provided for registered professional advisors And hedge And mutual funds.
''' </summary>
''' <returns></returns>
Public Tier As New TwsSoftDollarTier

Public IsOmsContainer As Boolean

Public DiscretionaryUpToLimitPrice As Boolean

Public UsePriceMgmtAlgo As Boolean

Public Duration As Long

Public PostToAts As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
LmtPrice = MaxDouble
AuxPrice = MaxDouble
OutsideRth = False
Origin = CUSTOMER
Transmit = True
ExemptCode = -1
MinQty = GIB.MaxLong
PercentOffset = MaxDouble
StartingPrice = MaxDouble
StockRefPrice = MaxDouble
Delta = MaxDouble
StockRangeLower = MaxDouble
StockRangeUpper = MaxDouble
Volatility = MaxDouble
VolatilityType = GIB.MaxLong
DeltaNeutralOrderType = TwsOrderTypeNone
DeltaNeutralAuxPrice = MaxDouble
ReferencePriceType = GIB.MaxLong
TrailStopPrice = MaxDouble
TrailingPercent = MaxDouble
BasisPoints = MaxDouble
BasisPointsType = GIB.MaxLong
ScaleInitLevelSize = GIB.MaxLong
ScaleSubsLevelSize = GIB.MaxLong
ScalePriceIncrement = MaxDouble
ScalePriceAdjustValue = MaxDouble
ScalePriceAdjustInterval = GIB.MaxLong
ScaleProfitOffset = MaxDouble
ScaleInitPosition = GIB.MaxLong
ScaleInitFillQty = GIB.MaxLong
TriggerPrice = MaxDouble
LmtPriceOffset = MaxDouble
AdjustedStopPrice = MaxDouble
AdjustedStopLimitPrice = MaxDouble
AdjustedTrailingAmount = MaxDouble
CashQty = MaxDouble
ReferenceFutureContractId = GIB.MaxLong
Duration = GIB.MaxLong
PostToAts = GIB.MaxLong
MinimumTradeQuantity = GIB.MaxLong
MinimumCompeteSize = GIB.MaxLong
CompeteAgainstBestOffset = MaxDouble
MidOffsetAtWhole = MaxDouble
MidOffsetAtHalf = MaxDouble
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

'Public Property Let Action(ByVal Value As TwsOrderActions)
'mAction = Value
'End Property
'
'Public Property Get Action() As TwsOrderActions
'Action = mAction
'End Property
'
'Public Property Let AuxPrice(ByVal Value As Double)
'mAuxPrice = Value
'End Property
'
'Public Property Get AuxPrice() As Double
'AuxPrice = mAuxPrice
'End Property
'
'Public Property Let ClientID(ByVal Value As Long)
'mClientID = Value
'End Property
'
'Public Property Get ClientID() As Long
'ClientID = mClientID
'End Property
'
'Public Property Let LmtPrice(ByVal Value As Double)
'mLmtPrice = Value
'End Property
'
'Public Property Get LmtPrice() As Double
'LmtPrice = mLmtPrice
'End Property
'
'Public Property Let OrderId(ByVal Value As Long)
'mOrderId = Value
'End Property
'
'Public Property Get OrderId() As Long
'OrderId = mOrderId
'End Property
'
'Public Property Let OrderType(ByVal Value As TwsOrderTypes)
'mOrderType = Value
'End Property
'
'Public Property Get OrderType() As TwsOrderTypes
'OrderType = mOrderType
'End Property
'
'Public Property Let PermId(ByVal Value As Long)
'mPermId = Value
'End Property
'
'Public Property Get PermId() As Long
'PermId = mPermId
'End Property
'
'Public Property Let TotalQuantity(ByVal Value As Double)
'mTotalQuantity = Value
'End Property
'
'Public Property Get TotalQuantity() As Double
'TotalQuantity = mTotalQuantity
'End Property
'
'Public Property Let Tif(ByVal Value As TwsOrderTIFs)
'mTif = Value
'End Property
'
'Public Property Get Tif() As TwsOrderTIFs
'Tif = mTif
'End Property
'
'Public Property Let OcaGroup(ByVal Value As String)
'mOcaGroup = Value
'End Property
'
'Public Property Get OcaGroup() As String
'OcaGroup = mOcaGroup
'End Property
'
'Public Property Let OcaType(ByVal Value As TwsOcaTypes)
'mOcaType = Value
'End Property
'
'Public Property Get OcaType() As TwsOcaTypes
'OcaType = mOcaType
'End Property
'
'Public Property Let OrderRef(ByVal Value As String)
'mOrderRef = Value
'End Property
'
'Public Property Get OrderRef() As String
'OrderRef = mOrderRef
'End Property
'
'Public Property Let Transmit(ByVal Value As Boolean)
'mTransmit = Value
'End Property
'
'Public Property Get Transmit() As Boolean
'Transmit = mTransmit
'End Property
'
'Public Property Let ParentId(ByVal Value As Long)
'mParentId = Value
'End Property
'
'Public Property Get ParentId() As Long
'ParentId = mParentId
'End Property
'
'Public Property Let BlockOrder(ByVal Value As Boolean)
'mBlockOrder = Value
'End Property
'
'Public Property Get BlockOrder() As Boolean
'BlockOrder = mBlockOrder
'End Property
'
'Public Property Let SweepToFill(ByVal Value As Boolean)
'mSweepToFill = Value
'End Property
'
'Public Property Get SweepToFill() As Boolean
'SweepToFill = mSweepToFill
'End Property
'
'Public Property Let DisplaySize(ByVal Value As Long)
'mDisplaySize = Value
'End Property
'
'Public Property Get DisplaySize() As Long
'DisplaySize = mDisplaySize
'End Property
'
'Public Property Let TriggerMethod(ByVal Value As Long)
'mTriggerMethod = Value
'End Property
'
'Public Property Get TriggerMethod() As Long
'TriggerMethod = mTriggerMethod
'End Property
'
'Public Property Let OutsideRth(ByVal Value As Boolean)
'mOutsideRth = Value
'End Property
'
'Public Property Get OutsideRth() As Boolean
'OutsideRth = mOutsideRth
'End Property
'
'Public Property Let Hidden(ByVal Value As Boolean)
'mHidden = Value
'End Property
'
'Public Property Get Hidden() As Boolean
'Hidden = mHidden
'End Property
'
'Public Property Let GoodAfterTime(ByVal Value As String)
'mGoodAfterTime = Value
'End Property
'
'Public Property Get GoodAfterTime() As String
'GoodAfterTime = mGoodAfterTime
'End Property
'
'Public Property Let GoodTillDate(ByVal Value As String)
'mGoodTillDate = Value
'End Property
'
'Public Property Get GoodTillDate() As String
'GoodTillDate = mGoodTillDate
'End Property
'
'Public Property Let OverridePercentageConstraints(ByVal Value As Boolean)
'mOverridePercentageConstraints = Value
'End Property
'
'Public Property Get OverridePercentageConstraints() As Boolean
'OverridePercentageConstraints = mOverridePercentageConstraints
'End Property
'
'Public Property Let Rule80A(ByVal Value As String)
'mRule80A = Value
'End Property
'
'Public Property Get Rule80A() As String
'Rule80A = mRule80A
'End Property
'
'Public Property Let AllOrNone(ByVal Value As Boolean)
'mAllOrNone = Value
'End Property
'
'Public Property Get AllOrNone() As Boolean
'AllOrNone = mAllOrNone
'End Property
'
'Public Property Let MinQty(ByVal Value As Long)
'mMinQty = Value
'End Property
'
'Public Property Get MinQty() As Long
'MinQty = mMinQty
'End Property
'
'Public Property Let PercentOffset(ByVal Value As Double)
'mPercentOffset = Value
'End Property
'
'Public Property Get PercentOffset() As Double
'PercentOffset = mPercentOffset
'End Property
'
'Public Property Let TrailStopPrice(ByVal Value As Double)
'mTrailStopPrice = Value
'End Property
'
'Public Property Get TrailStopPrice() As Double
'TrailStopPrice = mTrailStopPrice
'End Property
'
'Public Property Let FaGroup(ByVal Value As String)
'mFaGroup = Value
'End Property
'
'Public Property Get FaGroup() As String
'FaGroup = mFaGroup
'End Property
'
'Public Property Let FaProfile(ByVal Value As String)
'mFaProfile = Value
'End Property
'
'Public Property Get FaProfile() As String
'FaProfile = mFaProfile
'End Property
'
'Public Property Let FaMethod(ByVal Value As String)
'mFaMethod = Value
'End Property
'
'Public Property Get FaMethod() As String
'FaMethod = mFaMethod
'End Property
'
'Public Property Let FaPercentage(ByVal Value As String)
'mFaPercentage = Value
'End Property
'
'Public Property Get FaPercentage() As String
'FaPercentage = mFaPercentage
'End Property
'
'Public Property Let OpenClose(ByVal Value As String)
'mOpenClose = Value
'End Property
'
'Public Property Get OpenClose() As String
'OpenClose = mOpenClose
'End Property
'
'Public Property Let Origin(ByVal Value As Long)
'mOrigin = Value
'End Property
'
'Public Property Get Origin() As Long
'Origin = mOrigin
'End Property
'
'Public Property Let ShortSaleSlot(ByVal Value As Long)
'mShortSaleSlot = Value
'End Property
'
'Public Property Get ShortSaleSlot() As Long
'ShortSaleSlot = mShortSaleSlot
'End Property
'
'Public Property Let DesignatedLocation(ByVal Value As String)
'mDesignatedLocation = Value
'End Property
'
'Public Property Get DesignatedLocation() As String
'DesignatedLocation = mDesignatedLocation
'End Property
'
'Public Property Let DiscretionaryAmt(ByVal Value As Double)
'mDiscretionaryAmt = Value
'End Property
'
'Public Property Get DiscretionaryAmt() As Double
'DiscretionaryAmt = mDiscretionaryAmt
'End Property
'
'Public Property Let ETradeOnly(ByVal Value As Boolean)
'mETradeOnly = Value
'End Property
'
'Public Property Get ETradeOnly() As Boolean
'ETradeOnly = mETradeOnly
'End Property
'
'Public Property Let FirmQuoteOnly(ByVal Value As Boolean)
'mFirmQuoteOnly = Value
'End Property
'
'Public Property Get FirmQuoteOnly() As Boolean
'FirmQuoteOnly = mFirmQuoteOnly
'End Property
'
'Public Property Let NbboPriceCap(ByVal Value As Double)
'mNbboPriceCap = Value
'End Property
'
'Public Property Get NbboPriceCap() As Double
'NbboPriceCap = mNbboPriceCap
'End Property
'
'Public Property Let AuctionStrategy(ByVal Value As TwsAuctionStrategies)
'mAuctionStrategy = Value
'End Property
'
'Public Property Get AuctionStrategy() As TwsAuctionStrategies
'AuctionStrategy = mAuctionStrategy
'End Property
'
'Public Property Let StartingPrice(ByVal Value As Double)
'mStartingPrice = Value
'End Property
'
'Public Property Get StartingPrice() As Double
'StartingPrice = mStartingPrice
'End Property
'
'Public Property Let StockRefPrice(ByVal Value As Double)
'mStockRefPrice = Value
'End Property
'
'Public Property Get StockRefPrice() As Double
'StockRefPrice = mStockRefPrice
'End Property
'
'Public Property Let Delta(ByVal Value As Double)
'mDelta = Value
'End Property
'
'Public Property Get Delta() As Double
'Delta = mDelta
'End Property
'
'Public Property Let StockRangeLower(ByVal Value As Double)
'mStockRangeLower = Value
'End Property
'
'Public Property Get StockRangeLower() As Double
'StockRangeLower = mStockRangeLower
'End Property
'
'Public Property Let StockRangeUpper(ByVal Value As Double)
'mStockRangeUpper = Value
'End Property
'
'Public Property Get StockRangeUpper() As Double
'StockRangeUpper = mStockRangeUpper
'End Property
'
'Public Property Let Volatility(ByVal Value As Double)
'mVolatility = Value
'End Property
'
'Public Property Get Volatility() As Double
'Volatility = mVolatility
'End Property
'
'Public Property Let VolatilityType(ByVal Value As Long)
'mVolatilityType = Value
'End Property
'
'Public Property Get VolatilityType() As Long
'VolatilityType = mVolatilityType
'End Property
'
'Public Property Let ContinuousUpdate(ByVal Value As Long)
'mContinuousUpdate = Value
'End Property
'
'Public Property Get ContinuousUpdate() As Long
'ContinuousUpdate = mContinuousUpdate
'End Property
'
'Public Property Let ReferencePriceType(ByVal Value As Long)
'mReferencePriceType = Value
'End Property
'
'Public Property Get ReferencePriceType() As Long
'ReferencePriceType = mReferencePriceType
'End Property
'
'Public Property Let DeltaNeutralOrderType(ByVal Value As TwsOrderTypes)
'mDeltaNeutralOrderType = Value
'End Property
'
'Public Property Get DeltaNeutralOrderType() As TwsOrderTypes
'DeltaNeutralOrderType = mDeltaNeutralOrderType
'End Property
'
'Public Property Let DeltaNeutralAuxPrice(ByVal Value As Double)
'mDeltaNeutralAuxPrice = Value
'End Property
'
'Public Property Get DeltaNeutralAuxPrice() As Double
'DeltaNeutralAuxPrice = mDeltaNeutralAuxPrice
'End Property
'
'Public Property Let BasisPoints(ByVal Value As Double)
'mBasisPoints = Value
'End Property
'
'Public Property Get BasisPoints() As Double
'BasisPoints = mBasisPoints
'End Property
'
'Public Property Let BasisPointsType(ByVal Value As Long)
'mBasisPointsType = Value
'End Property
'
'Public Property Get BasisPointsType() As Long
'BasisPointsType = mBasisPointsType
'End Property
'
'Public Property Let ScaleInitLevelSize(ByVal Value As Long)
'mScaleInitLevelSize = Value
'End Property
'
'Public Property Get ScaleInitLevelSize() As Long
'ScaleInitLevelSize = mScaleInitLevelSize
'End Property
'
'Public Property Let ScaleSubsLevelSize(ByVal Value As Long)
'mScaleSubsLevelSize = Value
'End Property
'
'Public Property Get ScaleSubsLevelSize() As Long
'ScaleSubsLevelSize = mScaleSubsLevelSize
'End Property
'
'Public Property Let ScalePriceIncrement(ByVal Value As Double)
'mScalePriceIncrement = Value
'End Property
'
'Public Property Get ScalePriceIncrement() As Double
'ScalePriceIncrement = mScalePriceIncrement
'End Property
'
'Public Property Let Account(ByVal Value As String)
'mAccount = Value
'End Property
'
'Public Property Get Account() As String
'Account = mAccount
'End Property
'
'Public Property Let SettlingFirm(ByVal Value As String)
'mSettlingFirm = Value
'End Property
'
'Public Property Get SettlingFirm() As String
'SettlingFirm = mSettlingFirm
'End Property
'
'Public Property Let ClearingAccount(ByVal Value As String)
'mClearingAccount = Value
'End Property
'
'Public Property Get ClearingAccount() As String
'ClearingAccount = mClearingAccount
'End Property
'
'Public Property Let ClearingIntent(ByVal Value As String)
'mClearingIntent = Value
'End Property
'
'Public Property Get ClearingIntent() As String
'ClearingIntent = mClearingIntent
'End Property
'
'Public Property Let AlgoStrategy(ByVal Value As String)
'mAlgoStrategy = Value
'End Property
'
'Public Property Get AlgoStrategy() As String
'AlgoStrategy = mAlgoStrategy
'End Property
'
Public Property Let AlgoParams(ByVal Value As Parameters)
Set mAlgoParams = Value
End Property

Public Property Get AlgoParams() As Parameters
Set AlgoParams = mAlgoParams
End Property

'Public Property Let WhatIf(ByVal Value As Boolean)
'mWhatIf = Value
'End Property
'
'Public Property Get WhatIf() As Boolean
'WhatIf = mWhatIf
'End Property
'
'Public Property Let NotHeld(ByVal Value As Boolean)
'mNotHeld = Value
'End Property
'
'Public Property Get NotHeld() As Boolean
'NotHeld = mNotHeld
'End Property

Public Property Let ComboLegs(ByRef Value() As TwsOrderComboLeg)
mComboLegs = Value
End Property

Public Property Get ComboLegs() As TwsOrderComboLeg()
ComboLegs = mComboLegs
End Property

Public Property Let Options(ByVal Value As Parameters)
Set mOptions = Value
End Property

Public Property Get Options() As Parameters
Set Options = mOptions
End Property

Public Property Let SmartComboRoutingParams(ByVal Value As Parameters)
Set mSmartComboRoutingParams = Value
End Property

Public Property Get SmartComboRoutingParams() As Parameters
Set SmartComboRoutingParams = mSmartComboRoutingParams
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

