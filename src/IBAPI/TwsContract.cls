VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsContract"
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

Private Const ModuleName                            As String = "TwsContract"

'@================================================================================
' Member variables
'@================================================================================

Private mSpecifier                                  As TwsContractSpecifier
Private mMarketName                                 As String
Private mMinTick                                    As Double
Private mPriceMagnifier                             As Long
Private mOrderTypes                                 As String
Private mValidExchanges                             As String
Private mUnderConId                                 As Long
Private mLongName                                   As String
Private mContractMonth                              As String
Private mIndustry                                   As String
Private mCategory                                   As String
Private mSubcategory                                As String
Private mTimeZoneId                                 As String
Private mTradingHours                               As String
Private mLiquidHours                                As String

Private mEvRule                                     As String
Private mEvMultiplier                               As Double
Public AggGroup                                     As Long
Public UnderSymbol                                  As String
Public UnderSecType                                 As TwsSecTypes
Public MarketRuleIds                                As String
Public RealExpirationDate                           As String
Public LastTradeTime                                As Date

Private mStockType                                  As String
Private mMinimumSize                                As BoxedDecimal
Private mSizeIncrement                              As BoxedDecimal
Private mSuggestedSizeIncrement                     As BoxedDecimal

Private mSecIdList                                  As Parameters  ' CUSIP/ISIN/etc.

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
PriceMagnifier = 1
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

Public Property Let Category(ByVal Value As String)
mCategory = Value
End Property

Public Property Get Category() As String
Category = mCategory
End Property

Public Property Let ContractMonth(ByVal Value As String)
mContractMonth = Value
End Property

Public Property Get ContractMonth() As String
ContractMonth = mContractMonth
End Property

Public Property Let EvRule(ByVal Value As String)
mEvRule = Value
End Property

Public Property Get EvRule() As String
EvRule = mEvRule
End Property

Public Property Let EvMultiplier(ByVal Value As Double)
mEvMultiplier = Value
End Property

Public Property Get EvMultiplier() As Double
EvMultiplier = mEvMultiplier
End Property

Public Property Let Industry(ByVal Value As String)
mIndustry = Value
End Property

Public Property Get Industry() As String
Industry = mIndustry
End Property

Public Property Get Key() As String
Key = mSpecifier.Key
End Property

Public Property Let LiquidHours(ByVal Value As String)
mLiquidHours = Value
End Property

Public Property Get LiquidHours() As String
LiquidHours = mLiquidHours
End Property

Public Property Let LongName(ByVal Value As String)
mLongName = Value
End Property

Public Property Get LongName() As String
LongName = mLongName
End Property

Public Property Let MarketName(ByVal Value As String)
mMarketName = Value
End Property

Public Property Get MarketName() As String
MarketName = mMarketName
End Property

Public Property Let MinTick(ByVal Value As Double)
mMinTick = Value
End Property

Public Property Get MinTick() As Double
MinTick = mMinTick
End Property

Public Property Let OrderTypes(ByVal Value As String)
mOrderTypes = Value
End Property

Public Property Get OrderTypes() As String
OrderTypes = mOrderTypes
End Property

Public Property Let PriceMagnifier(ByVal Value As Long)
mPriceMagnifier = Value
End Property

Public Property Get PriceMagnifier() As Long
PriceMagnifier = mPriceMagnifier
End Property

Public Property Let SecIdList(ByVal Value As Parameters)
Set mSecIdList = Value
End Property

Public Property Get SecIdList() As Parameters
Set SecIdList = mSecIdList
End Property

Public Property Let Specifier(ByVal Value As TwsContractSpecifier)
Set mSpecifier = Value
End Property

Public Property Get Specifier() As TwsContractSpecifier
Set Specifier = mSpecifier
End Property

Public Property Let Subcategory(ByVal Value As String)
mSubcategory = Value
End Property

Public Property Get Subcategory() As String
Subcategory = mSubcategory
End Property

Public Property Let TimeZoneId(ByVal Value As String)
mTimeZoneId = Value
End Property

Public Property Get TimeZoneId() As String
TimeZoneId = mTimeZoneId
End Property

Public Property Let TradingHours(ByVal Value As String)
mTradingHours = Value
End Property

Public Property Get TradingHours() As String
TradingHours = mTradingHours
End Property

Public Property Let UnderConId(ByVal Value As Long)
mUnderConId = Value
End Property

Public Property Get UnderConId() As Long
UnderConId = mUnderConId
End Property

Public Property Let ValidExchanges(ByVal Value As String)
mValidExchanges = Value
End Property

Public Property Get ValidExchanges() As String
ValidExchanges = mValidExchanges
End Property

Public Property Let StockType(ByVal Value As String)
mStockType = Value
End Property

Public Property Get StockType() As String
StockType = mStockType
End Property

Public Property Let MinimumSize(ByVal Value As BoxedDecimal)
Set mMinimumSize = Value
End Property

Public Property Get MinimumSize() As BoxedDecimal
Set MinimumSize = mMinimumSize
End Property

Public Property Let SizeIncrement(ByVal Value As BoxedDecimal)
Set mSizeIncrement = Value
End Property

Public Property Get SizeIncrement() As BoxedDecimal
Set SizeIncrement = mSizeIncrement
End Property

Public Property Let SuggestedSizeIncrement(ByVal Value As BoxedDecimal)
Set mSuggestedSizeIncrement = Value
End Property

Public Property Get SuggestedSizeIncrement() As BoxedDecimal
Set SuggestedSizeIncrement = mSuggestedSizeIncrement
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




