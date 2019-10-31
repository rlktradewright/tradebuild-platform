VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderReport"
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

Implements IOrderReport
Implements IJSONable
Implements IStringable

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

Private Const ModuleName                    As String = "OrderReport"

'@================================================================================
' Member variables
'@================================================================================

Private mAction As OrderActions
Private mAllOrNone As Boolean
Private mBlockOrder As Boolean
Private mBrokerId As String
Private mDiscretionaryAmount As Double
Private mDisplaySize As Long
Private mETradeOnly As Boolean
Private mFirmQuoteOnly As Boolean
Private mGoodAfterTime As Date
Private mGoodAfterTimeTZ As String
Private mGoodTillDate As Date
Private mGoodTillDateTZ As String
Private mHidden As Boolean
Private mIgnoreRegularTradingHours As Boolean
Private mLimitPrice As Double
Private mMessage As String
Private mMinimumQuantity As Long
Private mNBBOPriceCap As Double
Private mOrderType As OrderTypes
Private mOrigin As OrderOrigins
Private mOriginatorRef As String
Private mOverrideConstraints As Boolean
Private mProviderProperties As Parameters
Private mQuantity As Long
Private mSettlingFirm As String
Private mStatus As OrderStatuses
Private mStopTriggerMethod As OrderStopTriggerMethods
Private mSweepToFill As Boolean
Private mTimeInForce As OrderTIFs
Private mIdentifier As String
Private mTriggerPrice As Double
Private mIsSimulated As Boolean
 
'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mProviderProperties = New Parameters
End Sub

'@================================================================================
' IOrderReportIOrderReport Interface Members
'@================================================================================

Private Property Let IOrderReport_Action(ByVal Value As OrderActions)
mAction = Value
End Property

Private Property Get IOrderReport_Action() As OrderActions
IOrderReport_Action = mAction
End Property

Private Property Let IOrderReport_AllOrNone(ByVal Value As Boolean)
mAllOrNone = Value
End Property

Private Property Get IOrderReport_AllOrNone() As Boolean
IOrderReport_AllOrNone = mAllOrNone
End Property

Private Property Let IOrderReport_BlockOrder(ByVal Value As Boolean)
mBlockOrder = Value
End Property

Private Property Get IOrderReport_BlockOrder() As Boolean
IOrderReport_BlockOrder = mBlockOrder
End Property

Private Property Let IOrderReport_BrokerId(ByVal Value As String)
mBrokerId = Value
End Property

Private Property Get IOrderReport_BrokerId() As String
IOrderReport_BrokerId = mBrokerId
End Property

Private Property Let IOrderReport_DiscretionaryAmount(ByVal Value As Double)
mDiscretionaryAmount = Value
End Property

Private Property Get IOrderReport_DiscretionaryAmount() As Double
IOrderReport_DiscretionaryAmount = mDiscretionaryAmount
End Property

Private Property Let IOrderReport_DisplaySize(ByVal Value As Long)
mDisplaySize = Value
End Property

Private Property Get IOrderReport_DisplaySize() As Long
IOrderReport_DisplaySize = mDisplaySize
End Property

Private Property Let IOrderReport_ETradeOnly(ByVal Value As Boolean)
mETradeOnly = Value
End Property

Private Property Get IOrderReport_ETradeOnly() As Boolean
IOrderReport_ETradeOnly = mETradeOnly
End Property

Private Property Let IOrderReport_FirmQuoteOnly(ByVal Value As Boolean)
mFirmQuoteOnly = Value
End Property

Private Property Get IOrderReport_FirmQuoteOnly() As Boolean
IOrderReport_FirmQuoteOnly = mFirmQuoteOnly
End Property

Private Property Let IOrderReport_GoodAfterTime(ByVal Value As Date)
mGoodAfterTime = Value
End Property

Private Property Get IOrderReport_GoodAfterTime() As Date
IOrderReport_GoodAfterTime = mGoodAfterTime
End Property

Private Property Let IOrderReport_GoodAfterTimeTZ(ByVal Value As String)
mGoodAfterTimeTZ = Value
End Property

Private Property Get IOrderReport_GoodAfterTimeTZ() As String
IOrderReport_GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Private Property Let IOrderReport_GoodTillDate(ByVal Value As Date)
mGoodTillDate = Value
End Property

Private Property Get IOrderReport_GoodTillDate() As Date
IOrderReport_GoodTillDate = mGoodTillDate
End Property

Private Property Let IOrderReport_GoodTillDateTZ(ByVal Value As String)
mGoodTillDateTZ = Value
End Property

Private Property Get IOrderReport_GoodTillDateTZ() As String
IOrderReport_GoodTillDateTZ = mGoodTillDateTZ
End Property

Private Property Let IOrderReport_Hidden(ByVal Value As Boolean)
mHidden = Value
End Property

Private Property Get IOrderReport_Hidden() As Boolean
IOrderReport_Hidden = mHidden
End Property

Private Property Let IOrderReport_Id(ByVal Value As String)
mIdentifier = Value
End Property

Private Property Get IOrderReport_Id() As String
IOrderReport_Id = mIdentifier
End Property

Private Property Let IOrderReport_IgnoreRegularTradingHours(ByVal Value As Boolean)
mIgnoreRegularTradingHours = Value
End Property

Private Property Get IOrderReport_IgnoreRegularTradingHours() As Boolean
IOrderReport_IgnoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Private Property Let IOrderReport_IsSimulated(ByVal Value As Boolean)
mIsSimulated = Value
End Property

Private Property Get IOrderReport_IsSimulated() As Boolean
IOrderReport_IsSimulated = mIsSimulated
End Property

Private Property Let IOrderReport_LimitPrice(ByVal Value As Double)
mLimitPrice = Value
End Property

Private Property Get IOrderReport_LimitPrice() As Double
IOrderReport_LimitPrice = mLimitPrice
End Property

Private Property Let IOrderReport_Message(ByVal Value As String)
Message = Value
End Property

Private Property Get IOrderReport_Message() As String
IOrderReport_Message = Message
End Property

Private Property Let IOrderReport_MinimumQuantity(ByVal Value As Long)
mMinimumQuantity = Value
End Property

Private Property Get IOrderReport_MinimumQuantity() As Long
IOrderReport_MinimumQuantity = mMinimumQuantity
End Property

Private Property Let IOrderReport_NbboPriceCap(ByVal Value As Double)
mNBBOPriceCap = Value
End Property

Private Property Get IOrderReport_NbboPriceCap() As Double
IOrderReport_NbboPriceCap = mNBBOPriceCap
End Property

Private Property Let IOrderReport_OrderType(ByVal Value As OrderTypes)
mOrderType = Value
End Property

Private Property Get IOrderReport_OrderType() As OrderTypes
IOrderReport_OrderType = mOrderType
End Property

Private Property Let IOrderReport_Origin(ByVal Value As OrderOrigins)
mOrigin = Value
End Property

Private Property Get IOrderReport_Origin() As OrderOrigins
IOrderReport_Origin = mOrigin
End Property

Private Property Let IOrderReport_OriginatorRef(ByVal Value As String)
mOriginatorRef = Value
End Property

Private Property Get IOrderReport_OriginatorRef() As String
IOrderReport_OriginatorRef = mOriginatorRef
End Property

Private Property Let IOrderReport_OverrideConstraints(ByVal Value As Boolean)
mOverrideConstraints = Value
End Property

Private Property Get IOrderReport_OverrideConstraints() As Boolean
IOrderReport_OverrideConstraints = mOverrideConstraints
End Property

Private Property Get IOrderReport_ProviderProperties() As Parameters
Set IOrderReport_ProviderProperties = mProviderProperties
End Property

Private Property Let IOrderReport_Quantity(ByVal Value As Long)
mQuantity = Value
End Property

Private Property Get IOrderReport_Quantity() As Long
IOrderReport_Quantity = mQuantity
End Property

Private Property Let IOrderReport_SettlingFirm(ByVal Value As String)
mSettlingFirm = Value
End Property

Private Property Get IOrderReport_SettlingFirm() As String
IOrderReport_SettlingFirm = mSettlingFirm
End Property

Private Property Let IOrderReport_Status(ByVal Value As OrderStatuses)
mStatus = Value
End Property

Private Property Get IOrderReport_Status() As OrderStatuses
IOrderReport_Status = mStatus
End Property

Private Property Let IOrderReport_StopTriggerMethod(ByVal Value As OrderStopTriggerMethods)
mStopTriggerMethod = Value
End Property

Private Property Get IOrderReport_StopTriggerMethod() As OrderStopTriggerMethods
IOrderReport_StopTriggerMethod = mStopTriggerMethod
End Property

Private Property Let IOrderReport_SweepToFill(ByVal Value As Boolean)
mSweepToFill = Value
End Property

Private Property Get IOrderReport_SweepToFill() As Boolean
IOrderReport_SweepToFill = mSweepToFill
End Property

Private Property Let IOrderReport_TimeInForce(ByVal Value As OrderTIFs)
mTimeInForce = Value
End Property

Private Property Get IOrderReport_TimeInForce() As OrderTIFs
IOrderReport_TimeInForce = mTimeInForce
End Property

Private Property Let IOrderReport_TriggerPrice(ByVal Value As Double)
mTriggerPrice = Value
End Property

Private Property Get IOrderReport_TriggerPrice() As Double
IOrderReport_TriggerPrice = mTriggerPrice
End Property

'@================================================================================
' IJSONAble Interface Members
'@================================================================================

Private Function IJSONAble_ToJSON() As String
IJSONAble_ToJSON = ToJSON
End Function

'@================================================================================
' IStringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
IStringable_ToString = ToJSON
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

Public Property Let IsSimulated(ByVal Value As Boolean)
mIsSimulated = Value
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Public Property Let LimitPrice(ByVal Value As Double)
mLimitPrice = Value
End Property

Public Property Get LimitPrice() As Double
LimitPrice = mLimitPrice
End Property

Public Property Let Message(ByVal Value As String)
mMessage = Value
End Property

Public Property Get Message() As String
Message = mMessage
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

Public Property Get ProviderProperties() As Parameters
Set ProviderProperties = mProviderProperties
End Property

Public Property Let Quantity(ByVal Value As Long)
mQuantity = Value
End Property

Public Property Get Quantity() As Long
Quantity = mQuantity
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

Public Function ToJSON() As String
Dim jb As JSONBuilder
Const ProcName As String = "ToJSON"
On Error GoTo Err

Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "Action"
jb.EncodeValue mAction

jb.EncodeName "AllOrNone"
jb.EncodeValue mAllOrNone

jb.EncodeName "BlockOrder"
jb.EncodeValue mBlockOrder

jb.EncodeName "BrokerId"
jb.EncodeValue mBrokerId

jb.EncodeName "DiscretionaryAmount"
jb.EncodeValue mDiscretionaryAmount

jb.EncodeName "DisplaySize"
jb.EncodeValue mDisplaySize

jb.EncodeName "ETradeOnly"
jb.EncodeValue mETradeOnly

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

jb.EncodeName "IsSimulated"
jb.EncodeValue mIsSimulated

jb.EncodeName "LimitPrice"
jb.EncodeValue mLimitPrice

jb.EncodeName "MinimumQuantity"
jb.EncodeValue mMinimumQuantity

jb.EncodeName "NbboPriceCap"
jb.EncodeValue mNBBOPriceCap

jb.EncodeName "OrderType"
jb.EncodeValue mOrderType

jb.EncodeName "Origin"
jb.EncodeValue mOrigin

jb.EncodeName "OriginatorRef"
jb.EncodeValue mOriginatorRef

jb.EncodeName "OverrideConstraints"
jb.EncodeValue mOverrideConstraints

jb.EncodeName "ProviderProperties"
jb.EncodeValue mProviderProperties

jb.EncodeName "Quantity"
jb.EncodeValue mQuantity

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


