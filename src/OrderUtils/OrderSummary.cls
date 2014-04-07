VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderSummary"
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

Implements JSONable

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

Private Const ModuleName                    As String = "OrderSummary"

'@================================================================================
' Member variables
'@================================================================================

Private mAction                             As OrderActions
Private mAllOrNone                          As Boolean
Private mBlockOrder                         As Boolean
Private mBrokerId                           As String
Private mContractSpecifier                  As IContractSpecifier
Private mDiscretionaryAmount                As Double
Private mDisplaySize                        As Long
Private mETradeOnly                         As Boolean
Private mFirmQuoteOnly                      As Boolean
Private mGoodAfterTime                      As Date
Private mGoodAfterTimeTZ                    As String
Private mGoodTillDate                       As Date
Private mGoodTillDateTZ                     As String
Private mHidden                             As Boolean
Private mIgnoreRegularTradingHours          As Boolean
Private mLimitPrice                         As Double
Private mMinimumQuantity                    As Long
Private mNBBOPriceCap                       As Double
Private mOrderType                          As OrderTypes
Private mOrigin                             As OrderOrigins
Private mOriginatorRef                      As String
Private mOverrideConstraints                As Boolean
Private mContextsName                       As String
Private mProviderProperties                 As Parameters
Private mQuantity                           As Long
Private mSettlingFirm                       As String
Private mStatus                             As OrderStatuses
Private mStopTriggerMethod                  As OrderStopTriggerMethods
Private mSweepToFill                        As Boolean
Private mTimeInForce                        As OrderTIFs
Private mIdentifier                         As String
Private mTriggerPrice                       As Double
Private mAveragePrice                       As Double
Private mFillTime                           As Date
Private mQuantityFilled                     As Long
Private mQuantityRemaining                  As Long
Private mLastFillPrice                      As Double
Private mOffset                             As Long
Private mPercentOffset                      As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mProviderProperties = New Parameters
End Sub

'@================================================================================
' JSONable Interface Members
'@================================================================================

Private Function JSONable_ToJSON() As String
JSONable_ToJSON = ToJSON
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Action(ByVal Value As OrderActions)
mAction = Value
End Property

Public Property Get Action() As OrderActions
Action = mAction
End Property

Public Property Let AllOrNone(ByVal Value As Boolean)
mAllOrNone = Value
End Property

Public Property Get AllOrNone() As Boolean
AllOrNone = mAllOrNone
End Property

Public Property Let AveragePrice(ByVal Value As Double)
mAveragePrice = Value
End Property

Public Property Get AveragePrice() As Double
AveragePrice = mAveragePrice
End Property

Public Property Let BlockOrder(ByVal Value As Boolean)
mBlockOrder = Value
End Property

Public Property Get BlockOrder() As Boolean
BlockOrder = mBlockOrder
End Property

Public Property Let BrokerId(ByVal Value As String)
mBrokerId = Value
End Property

Public Property Get BrokerId() As String
BrokerId = mBrokerId
End Property

Public Property Let ContractSpecifier(ByVal Value As IContractSpecifier)
Set mContractSpecifier = Value
End Property

Public Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

Public Property Let DiscretionaryAmount(ByVal Value As Double)
mDiscretionaryAmount = Value
End Property

Public Property Get DiscretionaryAmount() As Double
DiscretionaryAmount = mDiscretionaryAmount
End Property

Public Property Let DisplaySize(ByVal Value As Long)
mDisplaySize = Value
End Property

Public Property Get DisplaySize() As Long
DisplaySize = mDisplaySize
End Property

Public Property Let ETradeOnly(ByVal Value As Boolean)
mETradeOnly = Value
End Property

Public Property Get ETradeOnly() As Boolean
ETradeOnly = mETradeOnly
End Property

Public Property Let FillTime(ByVal Value As Date)
mFillTime = Value
End Property

Public Property Get FillTime() As Date
FillTime = mFillTime
End Property

Public Property Let FirmQuoteOnly(ByVal Value As Boolean)
mFirmQuoteOnly = Value
End Property

Public Property Get FirmQuoteOnly() As Boolean
FirmQuoteOnly = mFirmQuoteOnly
End Property

Public Property Let GoodAfterTime(ByVal Value As Date)
mGoodAfterTime = Value
End Property

Public Property Get GoodAfterTime() As Date
GoodAfterTime = mGoodAfterTime
End Property

Public Property Let GoodAfterTimeTZ(ByVal Value As String)
mGoodAfterTimeTZ = Value
End Property

Public Property Get GoodAfterTimeTZ() As String
GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Public Property Let GoodTillDate(ByVal Value As Date)
mGoodTillDate = Value
End Property

Public Property Get GoodTillDate() As Date
GoodTillDate = mGoodTillDate
End Property

Public Property Let GoodTillDateTZ(ByVal Value As String)
mGoodTillDateTZ = Value
End Property

Public Property Get GoodTillDateTZ() As String
GoodTillDateTZ = mGoodTillDateTZ
End Property

Public Property Let Hidden(ByVal Value As Boolean)
mHidden = Value
End Property

Public Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Public Property Let Id(ByVal Value As String)
mIdentifier = Value
End Property

Public Property Get Id() As String
Id = mIdentifier
End Property

Public Property Let IgnoreRegularTradingHours(ByVal Value As Boolean)
mIgnoreRegularTradingHours = Value
End Property

Public Property Get IgnoreRegularTradingHours() As Boolean
IgnoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Public Property Let LastFillPrice(ByVal Value As Double)
mLastFillPrice = Value
End Property

Public Property Get LastFillPrice() As Double
LastFillPrice = mLastFillPrice
End Property

Public Property Let LimitPrice(ByVal Value As Double)
mLimitPrice = Value
End Property

Public Property Get LimitPrice() As Double
LimitPrice = mLimitPrice
End Property

Public Property Let MinimumQuantity(ByVal Value As Long)
mMinimumQuantity = Value
End Property

Public Property Get MinimumQuantity() As Long
MinimumQuantity = mMinimumQuantity
End Property

Public Property Let NbboPriceCap(ByVal Value As Double)
mNBBOPriceCap = Value
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mNBBOPriceCap
End Property

Public Property Let Offset(ByVal Value As Long)
mOffset = Value
End Property

Public Property Get Offset() As Long
Offset = mOffset
End Property

Public Property Let OrderType(ByVal Value As OrderTypes)
mOrderType = Value
End Property

Public Property Get OrderType() As OrderTypes
OrderType = mOrderType
End Property

Public Property Let Origin(ByVal Value As OrderOrigins)
mOrigin = Value
End Property

Public Property Get Origin() As OrderOrigins
Origin = mOrigin
End Property

Public Property Let OriginatorRef(ByVal Value As String)
mOriginatorRef = Value
End Property

Public Property Get OriginatorRef() As String
OriginatorRef = mOriginatorRef
End Property

Public Property Let OverrideConstraints(ByVal Value As Boolean)
mOverrideConstraints = Value
End Property

Public Property Get OverrideConstraints() As Boolean
OverrideConstraints = mOverrideConstraints
End Property

Public Property Let PercentOffset(ByVal Value As Double)
mPercentOffset = Value
End Property

Public Property Get PercentOffset() As Double
PercentOffset = mPercentOffset
End Property

Public Property Let ContextsName(ByVal Value As String)
mContextsName = Value
End Property

Public Property Get ContextsName() As String
ContextsName = mContextsName
End Property

Public Property Let ProviderProperties(ByVal Value As Parameters)
Set mProviderProperties = Value
End Property

Public Property Get ProviderProperties() As Parameters
Set ProviderProperties = mProviderProperties
End Property

Public Property Let Quantity(ByVal Value As Long)
mQuantity = Value
End Property

Public Property Get Quantity() As Long
Quantity = mQuantity
End Property

Public Property Let QuantityFilled(ByVal Value As Long)
mQuantityFilled = Value
End Property

Public Property Get QuantityFilled() As Long
QuantityFilled = mQuantityFilled
End Property

Public Property Let QuantityRemaining(ByVal Value As Long)
mQuantityRemaining = Value
End Property

Public Property Get QuantityRemaining() As Long
QuantityRemaining = mQuantityRemaining
End Property

Public Property Let SettlingFirm(ByVal Value As String)
mSettlingFirm = Value
End Property

Public Property Get SettlingFirm() As String
SettlingFirm = mSettlingFirm
End Property

Public Property Let Status(ByVal Value As OrderStatuses)
mStatus = Value
End Property

Public Property Get Status() As OrderStatuses
Status = mStatus
End Property

Public Property Let StopTriggerMethod(ByVal Value As OrderStopTriggerMethods)
mStopTriggerMethod = Value
End Property

Public Property Get StopTriggerMethod() As OrderStopTriggerMethods
StopTriggerMethod = mStopTriggerMethod
End Property

Public Property Let SweepToFill(ByVal Value As Boolean)
mSweepToFill = Value
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSweepToFill
End Property

Public Property Let TimeInForce(ByVal Value As OrderTIFs)
mTimeInForce = Value
End Property

Public Property Get TimeInForce() As OrderTIFs
TimeInForce = mTimeInForce
End Property

Public Property Let TriggerPrice(ByVal Value As Double)
mTriggerPrice = Value
End Property

Public Property Get TriggerPrice() As Double
TriggerPrice = mTriggerPrice
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function ToJSON() As String
Const ProcName As String = "ToJSON"
On Error GoTo Err

Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "Action"
jb.EncodeValue mAction

jb.EncodeName "AllOrNone"
jb.EncodeValue mAllOrNone

jb.EncodeName "AveragePrice"
jb.EncodeValue mAveragePrice

jb.EncodeName "BlockOrder"
jb.EncodeValue mBlockOrder

jb.EncodeName "BrokerId"
jb.EncodeValue mBrokerId

jb.EncodeName "ContractSpecifier"
jb.EncodeValue mContractSpecifier

jb.EncodeName "DiscretionaryAmount"
jb.EncodeValue mDiscretionaryAmount

jb.EncodeName "DisplaySize"
jb.EncodeValue mDisplaySize

jb.EncodeName "ETradeOnly"
jb.EncodeValue mETradeOnly

jb.EncodeName "FillTime"
jb.EncodeValue mFillTime

jb.EncodeName "FirmQuoteOnly"
jb.EncodeValue mFirmQuoteOnly

jb.EncodeName "GoodAfterTime"
jb.EncodeValue mGoodAfterTime

jb.EncodeName "GoodAfterTimeTZ"
jb.EncodeValue mGoodAfterTimeTZ

jb.EncodeName "GoodTillDate"
jb.EncodeValue mGoodTillDate

jb.EncodeName "GoodTillDateTZ"
jb.EncodeValue mGoodTillDateTZ

jb.EncodeName "Hidden"
jb.EncodeValue mHidden

jb.EncodeName "Id"
jb.EncodeValue mIdentifier

jb.EncodeName "IgnoreRegularTradingHours"
jb.EncodeValue mIgnoreRegularTradingHours

jb.EncodeName "LastFillPrice"
jb.EncodeValue mLastFillPrice

jb.EncodeName "LimitPrice"
jb.EncodeValue mLimitPrice

jb.EncodeName "MinimumQuantity"
jb.EncodeValue mMinimumQuantity

jb.EncodeName "NbboPriceCap"
jb.EncodeValue mNBBOPriceCap

jb.EncodeName "Offset"
jb.EncodeValue mOffset

jb.EncodeName "OrderType"
jb.EncodeValue mOrderType

jb.EncodeName "Origin"
jb.EncodeValue mOrigin

jb.EncodeName "OriginatorRef"
jb.EncodeValue mOriginatorRef

jb.EncodeName "OverrideConstraints"
jb.EncodeValue mOverrideConstraints

jb.EncodeName "PercentOffset"
jb.EncodeValue mPercentOffset

jb.EncodeName "ContextsName"
jb.EncodeValue mContextsName

jb.EncodeName "ProviderProperties"
jb.EncodeValue mProviderProperties

jb.EncodeName "Quantity"
jb.EncodeValue mQuantity

jb.EncodeName "QuantityFilled"
jb.EncodeValue mQuantityFilled

jb.EncodeName "QuantityRemaining"
jb.EncodeValue mQuantityRemaining

jb.EncodeName "SettlingFirm"
jb.EncodeValue mSettlingFirm

jb.EncodeName "Status"
jb.EncodeValue mStatus

jb.EncodeName "StopTriggerMethod"
jb.EncodeValue mStopTriggerMethod

jb.EncodeName "SweepToFill"
jb.EncodeValue mSweepToFill

jb.EncodeName "TimeInForce"
jb.EncodeValue mTimeInForce

jb.EncodeName "TriggerPrice"
jb.EncodeValue mTriggerPrice

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================



