VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsContractDetails"
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

Private Const ModuleName                            As String = "TwsContractDetails"

'@================================================================================
' Member variables
'@================================================================================

Private mSummary                                    As TwsContract
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
Public MDSizeMultiplier                             As Long
Public AggGroup                                     As Long
Public UnderSymbol                                  As String
Public UnderSecType                                 As TwsSecTypes
Public MarketRuleIds                                As String
Public RealExpirationDate                           As String
Public LastTradeTime                                As Date

Private mSecIdList()                                As TwsTagValue  ' CUSIP/ISIN/etc.

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

Public Property Let SecIdList(ByRef Value() As TwsTagValue)
mSecIdList = Value
End Property

Public Property Get SecIdList() As TwsTagValue()
SecIdList = mSecIdList
End Property

Public Property Let Subcategory(ByVal Value As String)
mSubcategory = Value
End Property

Public Property Get Subcategory() As String
Subcategory = mSubcategory
End Property

Public Property Let Summary(ByVal Value As TwsContract)
Set mSummary = Value
End Property

Public Property Get Summary() As TwsContract
Set Summary = mSummary
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

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




