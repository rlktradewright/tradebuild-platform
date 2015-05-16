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
'Private mContractSpecifier                  As IContractSpecifier
Private mDiscretionaryAmount                As Double
Private mDisplaySize                        As Long
Private mErrorCode                          As Long
Private mErrorMessage                       As String
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

Private mIsDirty                            As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mStatus = OrderStatuses.OrderStatusCreated
mAction = OrderActionBuy
mOrderType = OrderTypeNone
mStopTriggerMethod = OrderStopTriggerDefault
mTimeInForce = OrderTIFDay

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
mIsDirty = True
End Property

Public Property Get Action() As OrderActions
Action = mAction
End Property

Public Property Let AllOrNone(ByVal Value As Boolean)
mAllOrNone = Value
mIsDirty = True
End Property

Public Property Get AllOrNone() As Boolean
AllOrNone = mAllOrNone
End Property

Public Property Let AveragePrice(ByVal Value As Double)
mAveragePrice = Value
mIsDirty = True
End Property

Public Property Get AveragePrice() As Double
AveragePrice = mAveragePrice
End Property

Public Property Let BlockOrder(ByVal Value As Boolean)
mBlockOrder = Value
mIsDirty = True
End Property

Public Property Get BlockOrder() As Boolean
BlockOrder = mBlockOrder
End Property

Public Property Let BrokerId(ByVal Value As String)
mBrokerId = Value
mIsDirty = True
End Property

Public Property Get BrokerId() As String
BrokerId = mBrokerId
End Property

'Public Property Let ContractSpecifier(ByVal Value As IContractSpecifier)
'Set mContractSpecifier = Value
'mIsDirty = True
'End Property
'
'Public Property Get ContractSpecifier() As IContractSpecifier
'Set ContractSpecifier = mContractSpecifier
'End Property

Public Property Let DiscretionaryAmount(ByVal Value As Double)
mDiscretionaryAmount = Value
mIsDirty = True
End Property

Public Property Get DiscretionaryAmount() As Double
DiscretionaryAmount = mDiscretionaryAmount
End Property

Public Property Let DisplaySize(ByVal Value As Long)
mDisplaySize = Value
mIsDirty = True
End Property

Public Property Get DisplaySize() As Long
DisplaySize = mDisplaySize
End Property

Public Property Let ErrorCode(ByVal Value As Long)
mErrorCode = Value
mIsDirty = True
End Property

Public Property Get ErrorCode() As Long
ErrorCode = mErrorCode
End Property

Public Property Let ErrorMessage(ByVal Value As String)
mErrorMessage = Value
mIsDirty = True
End Property

Public Property Get ErrorMessage() As String
ErrorMessage = mErrorMessage
End Property

Public Property Let ETradeOnly(ByVal Value As Boolean)
mETradeOnly = Value
mIsDirty = True
End Property

Public Property Get ETradeOnly() As Boolean
ETradeOnly = mETradeOnly
End Property

Public Property Let FillTime(ByVal Value As Date)
mFillTime = Value
mIsDirty = True
End Property

Public Property Get FillTime() As Date
FillTime = mFillTime
End Property

Public Property Let FirmQuoteOnly(ByVal Value As Boolean)
mFirmQuoteOnly = Value
mIsDirty = True
End Property

Public Property Get FirmQuoteOnly() As Boolean
FirmQuoteOnly = mFirmQuoteOnly
End Property

Public Property Let GoodAfterTime(ByVal Value As Date)
mGoodAfterTime = Value
mIsDirty = True
End Property

Public Property Get GoodAfterTime() As Date
GoodAfterTime = mGoodAfterTime
End Property

Public Property Let GoodAfterTimeTZ(ByVal Value As String)
mGoodAfterTimeTZ = Value
mIsDirty = True
End Property

Public Property Get GoodAfterTimeTZ() As String
GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Public Property Let GoodTillDate(ByVal Value As Date)
mGoodTillDate = Value
mIsDirty = True
End Property

Public Property Get GoodTillDate() As Date
GoodTillDate = mGoodTillDate
End Property

Public Property Let GoodTillDateTZ(ByVal Value As String)
mGoodTillDateTZ = Value
mIsDirty = True
End Property

Public Property Get GoodTillDateTZ() As String
GoodTillDateTZ = mGoodTillDateTZ
End Property

Public Property Let Hidden(ByVal Value As Boolean)
mHidden = Value
mIsDirty = True
End Property

Public Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Public Property Let Id(ByVal Value As String)
mIdentifier = Value
mIsDirty = True
End Property

Public Property Get Id() As String
Id = mIdentifier
End Property

Public Property Let IgnoreRegularTradingHours(ByVal Value As Boolean)
mIgnoreRegularTradingHours = Value
mIsDirty = True
End Property

Public Property Get IgnoreRegularTradingHours() As Boolean
IgnoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Public Property Get IsDirty() As Boolean
IsDirty = mIsDirty
End Property

Public Property Let LastFillPrice(ByVal Value As Double)
mLastFillPrice = Value
mIsDirty = True
End Property

Public Property Get LastFillPrice() As Double
LastFillPrice = mLastFillPrice
End Property

Public Property Let LimitPrice(ByVal Value As Double)
mLimitPrice = Value
mIsDirty = True
End Property

Public Property Get LimitPrice() As Double
LimitPrice = mLimitPrice
End Property

Public Property Let MinimumQuantity(ByVal Value As Long)
mMinimumQuantity = Value
mIsDirty = True
End Property

Public Property Get MinimumQuantity() As Long
MinimumQuantity = mMinimumQuantity
End Property

Public Property Let NbboPriceCap(ByVal Value As Double)
mNBBOPriceCap = Value
mIsDirty = True
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mNBBOPriceCap
End Property

Public Property Let Offset(ByVal Value As Long)
mOffset = Value
mIsDirty = True
End Property

Public Property Get Offset() As Long
Offset = mOffset
End Property

Public Property Let OrderType(ByVal Value As OrderTypes)
mOrderType = Value
mIsDirty = True
End Property

Public Property Get OrderType() As OrderTypes
OrderType = mOrderType
End Property

Public Property Let Origin(ByVal Value As OrderOrigins)
mOrigin = Value
mIsDirty = True
End Property

Public Property Get Origin() As OrderOrigins
Origin = mOrigin
End Property

Public Property Let OriginatorRef(ByVal Value As String)
mOriginatorRef = Value
mIsDirty = True
End Property

Public Property Get OriginatorRef() As String
OriginatorRef = mOriginatorRef
End Property

Public Property Let OverrideConstraints(ByVal Value As Boolean)
mOverrideConstraints = Value
mIsDirty = True
End Property

Public Property Get OverrideConstraints() As Boolean
OverrideConstraints = mOverrideConstraints
End Property

Public Property Let PercentOffset(ByVal Value As Double)
mPercentOffset = Value
mIsDirty = True
End Property

Public Property Get PercentOffset() As Double
PercentOffset = mPercentOffset
End Property

Public Property Let ContextsName(ByVal Value As String)
mContextsName = Value
mIsDirty = True
End Property

Public Property Get ContextsName() As String
ContextsName = mContextsName
End Property

Public Property Let ProviderProperties(ByVal Value As Parameters)
Set mProviderProperties = Value
mIsDirty = True
End Property

Public Property Get ProviderProperties() As Parameters
Set ProviderProperties = mProviderProperties
End Property

Public Property Let Quantity(ByVal Value As Long)
mQuantity = Value
mIsDirty = True
End Property

Public Property Get Quantity() As Long
Quantity = mQuantity
End Property

Public Property Let QuantityFilled(ByVal Value As Long)
mQuantityFilled = Value
mIsDirty = True
End Property

Public Property Get QuantityFilled() As Long
QuantityFilled = mQuantityFilled
End Property

Public Property Let QuantityRemaining(ByVal Value As Long)
mQuantityRemaining = Value
mIsDirty = True
End Property

Public Property Get QuantityRemaining() As Long
QuantityRemaining = mQuantityRemaining
End Property

Public Property Let SettlingFirm(ByVal Value As String)
mSettlingFirm = Value
mIsDirty = True
End Property

Public Property Get SettlingFirm() As String
SettlingFirm = mSettlingFirm
End Property

Public Property Let Status(ByVal Value As OrderStatuses)
mStatus = Value
mIsDirty = True
End Property

Public Property Get Status() As OrderStatuses
Status = mStatus
End Property

Public Property Let StopTriggerMethod(ByVal Value As OrderStopTriggerMethods)
mStopTriggerMethod = Value
mIsDirty = True
End Property

Public Property Get StopTriggerMethod() As OrderStopTriggerMethods
StopTriggerMethod = mStopTriggerMethod
End Property

Public Property Let SweepToFill(ByVal Value As Boolean)
mSweepToFill = Value
mIsDirty = True
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSweepToFill
End Property

Public Property Let TimeInForce(ByVal Value As OrderTIFs)
mTimeInForce = Value
mIsDirty = True
End Property

Public Property Get TimeInForce() As OrderTIFs
TimeInForce = mTimeInForce
End Property

Public Property Let TriggerPrice(ByVal Value As Double)
mTriggerPrice = Value
mIsDirty = True
End Property

Public Property Get TriggerPrice() As Double
TriggerPrice = mTriggerPrice
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub setClean()
mIsDirty = True
End Sub

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

jb.EncodeNameAndValue "AllOrNone", mAllOrNone, False

jb.EncodeNameAndValue "AveragePrice", mAveragePrice, 0#

jb.EncodeNameAndValue "BlockOrder", mBlockOrder, False

jb.EncodeName "BrokerId"
jb.EncodeValue mBrokerId

'jb.EncodeName "ContractSpecifier"
'jb.EncodeValue mContractSpecifier

jb.EncodeNameAndValue "DiscretionaryAmount", mDiscretionaryAmount, 0

jb.EncodeNameAndValue "DisplaySize", mDisplaySize, 0

jb.EncodeNameAndValue "ErrorCode", mErrorCode, 0

jb.EncodeNameAndValue "ErrorMessage", mErrorMessage, ""

jb.EncodeNameAndValue "ETradeOnly", mETradeOnly, False

jb.EncodeNameAndValue "FillTime", mFillTime, CDate(0#)

jb.EncodeNameAndValue "FirmQuoteOnly", mFirmQuoteOnly, False

jb.EncodeNameAndValue "GoodAfterTime", mGoodAfterTime, CDate(0#)

jb.EncodeNameAndValue "GoodAfterTimeTZ", mGoodAfterTimeTZ, ""

jb.EncodeNameAndValue "GoodTillDate", mGoodTillDate, CDate(0#)

jb.EncodeNameAndValue "GoodTillDateTZ", mGoodTillDateTZ, ""

jb.EncodeNameAndValue "Hidden", mHidden, False

jb.EncodeName "Id"
jb.EncodeValue mIdentifier

jb.EncodeNameAndValue "IgnoreRegularTradingHours", mIgnoreRegularTradingHours, False

jb.EncodeNameAndValue "LastFillPrice", mLastFillPrice, 0#

jb.EncodeNameAndValue "LimitPrice", mLimitPrice, 0#

jb.EncodeNameAndValue "MinimumQuantity", mMinimumQuantity, 0

jb.EncodeNameAndValue "NbboPriceCap", mNBBOPriceCap, 0#

jb.EncodeNameAndValue "Offset", mOffset, 0

jb.EncodeName "OrderType"
jb.EncodeValue mOrderType

jb.EncodeNameAndValue "Origin", mOrigin, OriginCustomer

jb.EncodeNameAndValue "OriginatorRef", mOriginatorRef, ""

jb.EncodeNameAndValue "OverrideConstraints", mOverrideConstraints, False

jb.EncodeNameAndValue "PercentOffset", mPercentOffset, 0#

jb.EncodeName "ContextsName"
jb.EncodeValue mContextsName

If mProviderProperties.Count <> 0 Then
    jb.EncodeName "ProviderProperties"
    jb.EncodeValue mProviderProperties
End If

jb.EncodeName "Quantity"
jb.EncodeValue mQuantity

jb.EncodeName "QuantityFilled"
jb.EncodeValue mQuantityFilled

jb.EncodeName "QuantityRemaining"
jb.EncodeValue mQuantityRemaining

jb.EncodeNameAndValue "SettlingFirm", mSettlingFirm, ""

jb.EncodeName "Status"
jb.EncodeValue mStatus

jb.EncodeNameAndValue "StopTriggerMethod", mStopTriggerMethod, OrderStopTriggerDefault

jb.EncodeNameAndValue "SweepToFill", mSweepToFill, False

jb.EncodeNameAndValue "TimeInForce", mTimeInForce, OrderTIFDay

jb.EncodeNameAndValue "TriggerPrice", mTriggerPrice, 0#

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================



