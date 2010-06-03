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
Private mMinimumQuantity As Long
Private mNBBOPriceCap As Double
Private mOrderType As OrderTypes
Private mOrigin As OrderOrigins
Private moriginatorRef As String
Private mOverrideConstraints As Boolean
Private mProviderProperties As Parameters
Private mQuantity As Long
Private mSettlingFirm As String
Private mStatus As OrderStatuses
Private mStopTriggerMethod As StopTriggerMethods
Private mSweepToFill As Boolean
Private mTimeInForce As OrderTifs
Private mIdentifier As String
Private mTriggerPrice As Double
Private mAveragePrice               As Double
Private mFillTime                   As Date
Private mQuantityFilled             As Long
Private mQuantityRemaining          As Long
Private mLastFillPrice              As Double
Private mOffset                     As Long
Private mPercentOffset              As Double
 
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

Public Property Let action(ByVal value As OrderActions)
mAction = value
End Property

Public Property Get action() As OrderActions
action = mAction
End Property

Public Property Let AllOrNone(ByVal value As Boolean)
mAllOrNone = value
End Property

Public Property Get AllOrNone() As Boolean
AllOrNone = mAllOrNone
End Property

Public Property Let AveragePrice(ByVal value As Double)
mAveragePrice = value
End Property

Public Property Get AveragePrice() As Double
AveragePrice = mAveragePrice
End Property

Public Property Let BlockOrder(ByVal value As Boolean)
mBlockOrder = value
End Property

Public Property Get BlockOrder() As Boolean
BlockOrder = mBlockOrder
End Property

Public Property Let BrokerId(ByVal value As String)
mBrokerId = value
End Property

Public Property Get BrokerId() As String
BrokerId = mBrokerId
End Property

Public Property Let DiscretionaryAmount(ByVal value As Double)
mDiscretionaryAmount = value
End Property

Public Property Get DiscretionaryAmount() As Double
DiscretionaryAmount = mDiscretionaryAmount
End Property

Public Property Let DisplaySize(ByVal value As Long)
mDisplaySize = value
End Property

Public Property Get DisplaySize() As Long
DisplaySize = mDisplaySize
End Property

Public Property Let ETradeOnly(ByVal value As Boolean)
mETradeOnly = value
End Property

Public Property Get ETradeOnly() As Boolean
ETradeOnly = mETradeOnly
End Property

Public Property Let FillTime(ByVal value As Date)
mFillTime = value
End Property

Public Property Get FillTime() As Date
FillTime = mFillTime
End Property

Public Property Let FirmQuoteOnly(ByVal value As Boolean)
mFirmQuoteOnly = value
End Property

Public Property Get FirmQuoteOnly() As Boolean
FirmQuoteOnly = mFirmQuoteOnly
End Property

Public Property Let GoodAfterTime(ByVal value As Date)
mGoodAfterTime = value
End Property

Public Property Get GoodAfterTime() As Date
GoodAfterTime = mGoodAfterTime
End Property

Public Property Let GoodAfterTimeTZ(ByVal value As String)
mGoodAfterTimeTZ = value
End Property

Public Property Get GoodAfterTimeTZ() As String
GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Public Property Let GoodTillDate(ByVal value As Date)
mGoodTillDate = value
End Property

Public Property Get GoodTillDate() As Date
GoodTillDate = mGoodTillDate
End Property

Public Property Let GoodTillDateTZ(ByVal value As String)
mGoodTillDateTZ = value
End Property

Public Property Get GoodTillDateTZ() As String
GoodTillDateTZ = mGoodTillDateTZ
End Property

Public Property Let Hidden(ByVal value As Boolean)
mHidden = value
End Property

Public Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Public Property Let Id(ByVal value As String)
mIdentifier = value
End Property

Public Property Get Id() As String
Id = mIdentifier
End Property

Public Property Let IgnoreRegularTradingHours(ByVal value As Boolean)
mIgnoreRegularTradingHours = value
End Property

Public Property Get IgnoreRegularTradingHours() As Boolean
IgnoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Public Property Let LastFillPrice(ByVal value As Double)
mLastFillPrice = value
End Property

Public Property Get LastFillPrice() As Double
LastFillPrice = mLastFillPrice
End Property

Public Property Let LimitPrice(ByVal value As Double)
mLimitPrice = value
End Property

Public Property Get LimitPrice() As Double
LimitPrice = mLimitPrice
End Property

Public Property Let MinimumQuantity(ByVal value As Long)
mMinimumQuantity = value
End Property

Public Property Get MinimumQuantity() As Long
MinimumQuantity = mMinimumQuantity
End Property

Public Property Let NbboPriceCap(ByVal value As Double)
mNBBOPriceCap = value
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mNBBOPriceCap
End Property

Public Property Let Offset(ByVal value As Long)
mOffset = value
End Property

Public Property Get Offset() As Long
Offset = mOffset
End Property

Public Property Let OrderType(ByVal value As OrderTypes)
mOrderType = value
End Property

Public Property Get OrderType() As OrderTypes
OrderType = mOrderType
End Property

Public Property Let Origin(ByVal value As OrderOrigins)
mOrigin = value
End Property

Public Property Get Origin() As OrderOrigins
Origin = mOrigin
End Property

Public Property Let OriginatorRef(ByVal value As String)
moriginatorRef = value
End Property

Public Property Get OriginatorRef() As String
OriginatorRef = moriginatorRef
End Property

Public Property Let OverrideConstraints(ByVal value As Boolean)
mOverrideConstraints = value
End Property

Public Property Get OverrideConstraints() As Boolean
OverrideConstraints = mOverrideConstraints
End Property

Public Property Let PercentOffset(ByVal value As Double)
mPercentOffset = value
End Property

Public Property Get PercentOffset() As Double
PercentOffset = mPercentOffset
End Property

Public Property Let ProviderProperties(ByVal value As TWUtilities30.Parameters)
Set mProviderProperties = value
End Property

Public Property Get ProviderProperties() As TWUtilities30.Parameters
Set ProviderProperties = mProviderProperties
End Property

Public Property Let Quantity(ByVal value As Long)
mQuantity = value
End Property

Public Property Get Quantity() As Long
Quantity = mQuantity
End Property

Public Property Let QuantityFilled(ByVal value As Long)
mQuantityFilled = value
End Property

Public Property Get QuantityFilled() As Long
QuantityFilled = mQuantityFilled
End Property

Public Property Let QuantityRemaining(ByVal value As Long)
mQuantityRemaining = value
End Property

Public Property Get QuantityRemaining() As Long
QuantityRemaining = mQuantityRemaining
End Property

Public Property Let SettlingFirm(ByVal value As String)
mSettlingFirm = value
End Property

Public Property Get SettlingFirm() As String
SettlingFirm = mSettlingFirm
End Property

Public Property Let Status(ByVal value As OrderStatuses)
mStatus = value
End Property

Public Property Get Status() As OrderStatuses
Status = mStatus
End Property

Public Property Let StopTriggerMethod(ByVal value As StopTriggerMethods)
mStopTriggerMethod = value
End Property

Public Property Get StopTriggerMethod() As StopTriggerMethods
StopTriggerMethod = mStopTriggerMethod
End Property

Public Property Let SweepToFill(ByVal value As Boolean)
mSweepToFill = value
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSweepToFill
End Property

Public Property Let TimeInForce(ByVal value As OrderTifs)
mTimeInForce = value
End Property

Public Property Get TimeInForce() As OrderTifs
TimeInForce = mTimeInForce
End Property

Public Property Let TriggerPrice(ByVal value As Double)
mTriggerPrice = value
End Property

Public Property Get TriggerPrice() As Double
TriggerPrice = mTriggerPrice
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function ToJSON() As String
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

jb.EncodeName "AveragePrice"
jb.EncodeValue mAveragePrice

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
jb.EncodeValue moriginatorRef

jb.EncodeName "OverrideConstraints"
jb.EncodeValue mOverrideConstraints

jb.EncodeName "PercentOffset"
jb.EncodeValue mPercentOffset

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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' Helper Functions
'@================================================================================



