VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OS"
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

Implements IJSONable

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

Private Const ModuleName                            As String = "OS"

'@================================================================================
' Member variables
'@================================================================================

Private mAction                                     As OrderActions
Private mAllOrNone                                  As Boolean
Private mBlockOrder                                 As Boolean
Private mBrokerId                                   As String
Private mDiscretionaryAmount                        As Double
Private mDisplaySize                                As Long
Private mErrorCode                                  As Long
Private mErrorMessage                               As String
Private mGoodAfterTime                              As Date
Private mGoodAfterTimeTZ                            As String
Private mGoodTillDate                               As Date
Private mGoodTillDateTZ                             As String
Private mGroupName                                  As String
Private mHidden                                     As Boolean
Private mIgnoreRegularTradingHours                  As Boolean
Private mLimitPrice                                 As Double
Private mLimitPriceSpec                             As PriceSpecifier
Private mMessage                                    As String
Private mMinimumQuantity                            As Long
Private mNBBOPriceCap                               As Double
Private mOrderType                                  As OrderTypes
Private mOrigin                                     As OrderOrigins
Private mOriginatorRef                              As String
Private mOverrideConstraints                        As Boolean
Private mQuantity                                   As BoxedDecimal
Private mSettlingFirm                               As String
Private mStatus                                     As OrderStatuses
Private mStopTriggerMethod                          As OrderStopTriggerMethods
Private mSweepToFill                                As Boolean
Private mTimeInForce                                As OrderTIFs
Private mIdentifier                                 As String
Private mTriggerPrice                               As Double
Private mTriggerPriceSpec                           As PriceSpecifier
Private mAveragePrice                               As Double
Private mFillTime                                   As Date
Private mQuantityFilled                             As BoxedDecimal
Private mQuantityRemaining                          As BoxedDecimal
Private mLastFillPrice                              As Double
Private mPercentOffset                              As Double

Private mIsDirty                                    As Boolean

Private mDiff                                       As New ROD

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mStatus = OrderStatuses.OrderStatusCreated
mAction = OrderActionBuy
mOrderType = OrderTypeMarket
mLimitPrice = MaxDouble
Set mLimitPriceSpec = gNewPriceSpecifier
mTriggerPrice = MaxDouble
Set mTriggerPriceSpec = gNewPriceSpecifier
mStopTriggerMethod = OrderStopTriggerDefault
mTimeInForce = OrderTIFDay
Set mQuantity = DecimalZero
Set mQuantityFilled = DecimalZero
Set mQuantityRemaining = DecimalZero
End Sub

'@================================================================================
' IJSONAble Interface Members
'@================================================================================

Private Function IJSONAble_ToJSON() As String
IJSONAble_ToJSON = ToJSON
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Action(ByVal Value As OrderActions)
If mAction = Value Then Exit Property
mAction = Value
mIsDirty = True
mDiff.SetProperty "Action", Value
End Property

Public Property Get Action() As OrderActions
Action = mAction
End Property

Public Property Let AllOrNone(ByVal Value As Boolean)
If mAllOrNone = Value Then Exit Property
mAllOrNone = Value
mIsDirty = True
mDiff.SetProperty "AllOrNone", Value
End Property

Public Property Get AllOrNone() As Boolean
AllOrNone = mAllOrNone
End Property

Public Property Let AveragePrice(ByVal Value As Double)
If mAveragePrice = Value Then Exit Property
mAveragePrice = Value
mIsDirty = True
mDiff.SetProperty "AveragePrice", Value
End Property

Public Property Get AveragePrice() As Double
AveragePrice = mAveragePrice
End Property

Public Property Let BlockOrder(ByVal Value As Boolean)
If mBlockOrder = Value Then Exit Property
mBlockOrder = Value
mIsDirty = True
mDiff.SetProperty "BlockOrder", Value
End Property

Public Property Get BlockOrder() As Boolean
BlockOrder = mBlockOrder
End Property

Public Property Let BrokerId(ByVal Value As String)
If mBrokerId = Value Then Exit Property
mBrokerId = Value
mIsDirty = True
End Property

Public Property Get BrokerId() As String
BrokerId = mBrokerId
End Property

Public Property Let Diff(ByVal Value As ROD)
Set mDiff = Value
End Property

Public Property Get Diff() As ROD
Set Diff = mDiff
End Property

Public Property Let DiscretionaryAmount(ByVal Value As Double)
If mDiscretionaryAmount = Value Then Exit Property
mDiscretionaryAmount = Value
mIsDirty = True
mDiff.SetProperty "DiscretionaryAmount", Value
End Property

Public Property Get DiscretionaryAmount() As Double
DiscretionaryAmount = mDiscretionaryAmount
End Property

Public Property Let DisplaySize(ByVal Value As Long)
If mDisplaySize = Value Then Exit Property
mDisplaySize = Value
mIsDirty = True
mDiff.SetProperty "DisplaySize", Value
End Property

Public Property Get DisplaySize() As Long
DisplaySize = mDisplaySize
End Property

Public Property Let ErrorCode(ByVal Value As Long)
If mErrorCode = Value Then Exit Property
mErrorCode = Value
mIsDirty = True
mDiff.SetProperty "ErrorCode", Value
End Property

Public Property Get ErrorCode() As Long
ErrorCode = mErrorCode
End Property

Public Property Let ErrorMessage(ByVal Value As String)
If mErrorMessage = Value Then Exit Property
mErrorMessage = Value
mIsDirty = True
mDiff.SetProperty "ErrorMessage", Value
End Property

Public Property Get ErrorMessage() As String
ErrorMessage = mErrorMessage
End Property

Public Property Let FillTime(ByVal Value As Date)
If mFillTime = Value Then Exit Property
mFillTime = Value
mIsDirty = True
mDiff.SetProperty "FillTime", Value
End Property

Public Property Get FillTime() As Date
FillTime = mFillTime
End Property

Public Property Let GoodAfterTime(ByVal Value As Date)
If mGoodAfterTime = Value Then Exit Property
mGoodAfterTime = Value
mIsDirty = True
mDiff.SetProperty "GoodAfterTime", Value
End Property

Public Property Get GoodAfterTime() As Date
GoodAfterTime = mGoodAfterTime
End Property

Public Property Let GoodAfterTimeTZ(ByVal Value As String)
If mGoodAfterTimeTZ = Value Then Exit Property
mGoodAfterTimeTZ = Value
mIsDirty = True
mDiff.SetProperty "GoodAfterTimeTZ", Value
End Property

Public Property Get GoodAfterTimeTZ() As String
GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Public Property Let GoodTillDate(ByVal Value As Date)
If mGoodTillDate = Value Then Exit Property
mGoodTillDate = Value
mIsDirty = True
mDiff.SetProperty "GoodTillDate", Value
End Property

Public Property Get GoodTillDate() As Date
GoodTillDate = mGoodTillDate
End Property

Public Property Let GoodTillDateTZ(ByVal Value As String)
If mGoodTillDateTZ = Value Then Exit Property
mGoodTillDateTZ = Value
mIsDirty = True
mDiff.SetProperty "GoodTillDateTZ", Value
End Property

Public Property Get GoodTillDateTZ() As String
GoodTillDateTZ = mGoodTillDateTZ
End Property

Public Property Let GroupName(ByVal Value As String)
If mGroupName = Value Then Exit Property
mGroupName = Value
mIsDirty = True
mDiff.SetProperty "GroupName", Value
End Property

Public Property Get GroupName() As String
GroupName = mGroupName
End Property

Public Property Let Hidden(ByVal Value As Boolean)
If mHidden = Value Then Exit Property
mHidden = Value
mIsDirty = True
mDiff.SetProperty "Hidden", Value
End Property

Public Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Public Property Let Id(ByVal Value As String)
If mIdentifier = Value Then Exit Property
mIdentifier = Value
mIsDirty = True
End Property

Public Property Get Id() As String
Id = mIdentifier
End Property

Public Property Let IgnoreRegularTradingHours(ByVal Value As Boolean)
If mIgnoreRegularTradingHours = Value Then Exit Property
mIgnoreRegularTradingHours = Value
mIsDirty = True
mDiff.SetProperty "IgnoreRegularTradingHours", Value
End Property

Public Property Get IgnoreRegularTradingHours() As Boolean
IgnoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Public Property Get IsDirty() As Boolean
IsDirty = mIsDirty
End Property

Public Property Let LastFillPrice(ByVal Value As Double)
If mLastFillPrice = Value Then Exit Property
mLastFillPrice = Value
mIsDirty = True
mDiff.SetProperty "LastFillPrice", Value
End Property

Public Property Get LastFillPrice() As Double
LastFillPrice = mLastFillPrice
End Property

Public Property Let LimitPrice(ByVal Value As Double)
If mLimitPrice = Value Then Exit Property
mLimitPrice = Value
mIsDirty = True
mDiff.SetProperty "LimitPrice", Value
End Property

Public Property Get LimitPrice() As Double
LimitPrice = mLimitPrice
End Property

Public Property Let LimitPriceSpec(ByVal Value As PriceSpecifier)
Set mLimitPriceSpec = Value
mIsDirty = True
mDiff.SetProperty "LimitPriceSpec", Value
End Property

Public Property Get LimitPriceSpec() As PriceSpecifier
Set LimitPriceSpec = mLimitPriceSpec
End Property

Public Property Let Message(ByVal Value As String)
If mMessage = Value Then Exit Property
mMessage = Value
mIsDirty = True
mDiff.SetProperty "Message", Value
End Property

Public Property Get Message() As String
Message = mMessage
End Property

Public Property Let MinimumQuantity(ByVal Value As Long)
If mMinimumQuantity = Value Then Exit Property
mMinimumQuantity = Value
mIsDirty = True
mDiff.SetProperty "MinimumQuantity", Value
End Property

Public Property Get MinimumQuantity() As Long
MinimumQuantity = mMinimumQuantity
End Property

Public Property Let OrderType(ByVal Value As OrderTypes)
If mOrderType = Value Then Exit Property
mOrderType = Value
mIsDirty = True
mDiff.SetProperty "OrderType", Value
End Property

Public Property Get OrderType() As OrderTypes
OrderType = mOrderType
End Property

Public Property Let Origin(ByVal Value As OrderOrigins)
If mOrigin = Value Then Exit Property
mOrigin = Value
mIsDirty = True
mDiff.SetProperty "Origin", Value
End Property

Public Property Get Origin() As OrderOrigins
Origin = mOrigin
End Property

Public Property Let OriginatorRef(ByVal Value As String)
If mOriginatorRef = Value Then Exit Property
mOriginatorRef = Value
mIsDirty = True
mDiff.SetProperty "mOriginatorRef", Value
End Property

Public Property Get OriginatorRef() As String
OriginatorRef = mOriginatorRef
End Property

Public Property Let OverrideConstraints(ByVal Value As Boolean)
If mOverrideConstraints = Value Then Exit Property
mOverrideConstraints = Value
mIsDirty = True
mDiff.SetProperty "OverrideConstraints", Value
End Property

Public Property Get OverrideConstraints() As Boolean
OverrideConstraints = mOverrideConstraints
End Property

Public Property Let PercentOffset(ByVal Value As Double)
If mPercentOffset = Value Then Exit Property
mPercentOffset = Value
mIsDirty = True
mDiff.SetProperty "PercentOffset", Value
End Property

Public Property Get PercentOffset() As Double
PercentOffset = mPercentOffset
End Property

Public Property Let Quantity(ByVal Value As BoxedDecimal)
If mQuantity.EQ(Value) Then Exit Property
Set mQuantity = Value
mIsDirty = True
mDiff.SetProperty "Quantity", Value
End Property

Public Property Get Quantity() As BoxedDecimal
Set Quantity = mQuantity
End Property

Public Property Let QuantityFilled(ByVal Value As BoxedDecimal)
If mQuantityFilled.EQ(Value) Then Exit Property
Set mQuantityFilled = Value
mIsDirty = True
mDiff.SetProperty "QuantityFilled", Value
End Property

Public Property Get QuantityFilled() As BoxedDecimal
Set QuantityFilled = mQuantityFilled
End Property

Public Property Let QuantityRemaining(ByVal Value As BoxedDecimal)
If mQuantityRemaining.EQ(Value) Then Exit Property
Set mQuantityRemaining = Value
mIsDirty = True
mDiff.SetProperty "QuantityRemaining", Value
End Property

Public Property Get QuantityRemaining() As BoxedDecimal
Set QuantityRemaining = mQuantityRemaining
End Property

Public Property Let SettlingFirm(ByVal Value As String)
If mSettlingFirm = Value Then Exit Property
mSettlingFirm = Value
mIsDirty = True
mDiff.SetProperty "SettlingFirm", Value
End Property

Public Property Get SettlingFirm() As String
SettlingFirm = mSettlingFirm
End Property

Public Property Let Status(ByVal Value As OrderStatuses)
If mStatus = Value Then Exit Property
mStatus = Value
mIsDirty = True
mDiff.SetProperty "Status", Value
End Property

Public Property Get Status() As OrderStatuses
Status = mStatus
End Property

Public Property Let StopTriggerMethod(ByVal Value As OrderStopTriggerMethods)
If mStopTriggerMethod = Value Then Exit Property
mStopTriggerMethod = Value
mIsDirty = True
mDiff.SetProperty "StopTriggerMethod", Value
End Property

Public Property Get StopTriggerMethod() As OrderStopTriggerMethods
StopTriggerMethod = mStopTriggerMethod
End Property

Public Property Let SweepToFill(ByVal Value As Boolean)
If mSweepToFill = Value Then Exit Property
mSweepToFill = Value
mIsDirty = True
mDiff.SetProperty "SweepToFill", Value
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSweepToFill
End Property

Public Property Let TimeInForce(ByVal Value As OrderTIFs)
If mTimeInForce = Value Then Exit Property
mTimeInForce = Value
mIsDirty = True
mDiff.SetProperty "TimeInForce", Value
End Property

Public Property Get TimeInForce() As OrderTIFs
TimeInForce = mTimeInForce
End Property

Public Property Let TriggerPrice(ByVal Value As Double)
If mTriggerPrice = Value Then Exit Property
mTriggerPrice = Value
mIsDirty = True
mDiff.SetProperty "TriggerPrice", Value
End Property

Public Property Get TriggerPrice() As Double
TriggerPrice = mTriggerPrice
End Property

Public Property Let TriggerPriceSpec(ByVal Value As PriceSpecifier)
Set mTriggerPriceSpec = Value
mIsDirty = True
mDiff.SetProperty "TriggerPriceSpec", Value
End Property

Public Property Get TriggerPriceSpec() As PriceSpecifier
Set TriggerPriceSpec = mTriggerPriceSpec
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub SetClean()
mIsDirty = True
mDiff.Clear
End Sub

Friend Function ToJSON() As String
Const ProcName As String = "ToJSON"
On Error GoTo Err

Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "BrokerId"
jb.EncodeValue mBrokerId

jb.EncodeName "Id"
jb.EncodeValue mIdentifier

jb.EncodeNameAndValue "Diff", mDiff, Nothing

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================



