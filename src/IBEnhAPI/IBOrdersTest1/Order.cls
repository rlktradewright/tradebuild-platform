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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements IOrder

'@================================================================================
' Events
'@================================================================================

Event Clean()
Event Dirty()
Event Fill(ByVal pExec As Execution)
Event PropertyChanged()
Event StatusChanged()

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "Order"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpec               As IContractSpecifier

Private mIsDirty                    As Boolean

Private mStatus                     As OrderStatuses
Private mIdentifier                 As String

Private mIsSimulated                As Boolean

Private mFillTime                   As Date

Private mBrokerId                   As String
Private mBrokerData                 As Variant
    
Private mAction                     As OrderActions
Private mQuantity                   As Long
Private mQuantityFilled             As Long
Private mQuantityRemaining          As Long
Private mAveragePrice               As Double
Private mLastFillPrice              As Double
Private mOrderType                  As OrderTypes
Private mLimitPrice                 As Double
Private mLimitPriceSpec             As PriceSpecifier
Private mTriggerPrice               As Double
Private mTriggerPriceSpec           As PriceSpecifier
Private mOffset                     As Long

Private mTimeInForce                As OrderTIFs
Private mOrigin                     As OrderOrigins
Private moriginatorRef              As String
Private mProviderProperties         As Parameters
Private mBlockOrder                 As Boolean
Private mSweepToFill                As Boolean
Private mDisplaySize                As Long
Private mStopTriggerMethod     As OrderStopTriggerMethods
Private mIgnoreRegularTradingHours  As Boolean
Private mHidden                     As Boolean
Private mDiscretionaryAmount        As Double
Private mGoodAfterTime              As Date
Private mGoodAfterTimeTZ            As String
Private mGoodTillDate               As Date
Private mGoodTillDateTZ             As String
    
Private mSettlingFirm               As String
Private mAllOrNone                  As Boolean
Private mMinimumQuantity            As Long
Private mPercentOffset              As Double
Private mETradeOnly                 As Boolean
Private mFirmQuoteOnly              As Boolean
Private mNBBOPriceCap               As Double
Private mOverrideConstraints        As Boolean
    
Private mGroupName                  As String
Private mOrderContext               As OrderContext

Private mErrorCode                  As Long
Private mErrorMessage               As String
Private mMessage               As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mStatus = OrderStatuses.OrderStatusCreated
mIdentifier = GenerateGUIDString
mAction = OrderActionBuy
mOrderType = OrderTypeNone
mTimeInForce = OrderTIFDay
mStopTriggerMethod = OrderStopTriggerDefault

Set mProviderProperties = New Parameters

End Sub

'@================================================================================
' IOrder Interface Members
'@================================================================================

Private Property Let IOrder_Action(ByVal Value As OrderActions)
Action = Value
End Property

Private Property Get IOrder_Action() As OrderActions
IOrder_Action = mAction
End Property

Private Property Let IOrder_AllOrNone(ByVal Value As Boolean)
AllOrNone = Value
End Property

Private Property Get IOrder_AllOrNone() As Boolean
IOrder_AllOrNone = mAllOrNone
End Property

Private Property Let IOrder_AveragePrice(ByVal Value As Double)
AveragePrice = Value
End Property

Private Property Get IOrder_AveragePrice() As Double
IOrder_AveragePrice = AveragePrice
End Property

Private Property Let IOrder_BlockOrder(ByVal Value As Boolean)
BlockOrder = Value
End Property

Private Property Get IOrder_BlockOrder() As Boolean
IOrder_BlockOrder = mBlockOrder
End Property

Private Property Let IOrder_BrokerData(ByVal Value As Variant)
gSetVariant mBrokerData, Value
End Property

Private Property Get IOrder_BrokerData() As Variant
gSetVariant IOrder_BrokerData, mBrokerData
End Property

Private Property Let IOrder_BrokerId(ByVal Value As String)
BrokerId = Value
End Property

Private Property Get IOrder_BrokerId() As String
IOrder_BrokerId = mBrokerId
End Property

Private Property Get IOrder_ContractSpecifier() As IContractSpecifier
Set IOrder_ContractSpecifier = ContractSpecifier
End Property

Private Property Get IOrder_ContractSpec() As IContractSpecifier
Set IOrder_ContractSpec = mContractSpec
End Property

Private Property Let IOrder_DiscretionaryAmount(ByVal Value As Double)
DiscretionaryAmount = Value
End Property

Private Property Get IOrder_DiscretionaryAmount() As Double
IOrder_DiscretionaryAmount = mDiscretionaryAmount
End Property

Private Property Let IOrder_DisplaySize(ByVal Value As Long)
DisplaySize = Value
End Property

Private Property Get IOrder_DisplaySize() As Long
IOrder_DisplaySize = mDisplaySize
End Property

Private Property Let IOrder_ErrorCode(ByVal Value As Long)
ErrorCode = Value
End Property

Private Property Get IOrder_ErrorCode() As Long
IOrder_ErrorCode = ErrorCode
End Property

Private Property Let IOrder_ErrorMessage(ByVal Value As String)
ErrorMessage = Value
End Property

Private Property Get IOrder_ErrorMessage() As String
IOrder_ErrorMessage = ErrorMessage
End Property

Private Property Let IOrder_ETradeOnly(ByVal Value As Boolean)
ETradeOnly = Value
End Property

Private Property Get IOrder_ETradeOnly() As Boolean
IOrder_ETradeOnly = mETradeOnly
End Property

Private Property Let IOrder_FillTime(ByVal Value As Date)
FillTime = Value
End Property

Private Property Get IOrder_FillTime() As Date
IOrder_FillTime = FillTime
End Property

Private Property Let IOrder_FirmQuoteOnly(ByVal Value As Boolean)
FirmQuoteOnly = Value
End Property

Private Property Get IOrder_FirmQuoteOnly() As Boolean
IOrder_FirmQuoteOnly = mFirmQuoteOnly
End Property

Private Property Let IOrder_GoodAfterTime(ByVal Value As Date)
GoodAfterTime = Value
End Property

Private Property Get IOrder_GoodAfterTime() As Date
IOrder_GoodAfterTime = mGoodAfterTime
End Property

Private Property Let IOrder_GoodAfterTimeTZ(ByVal Value As String)
GoodAfterTimeTZ = Value
End Property

Private Property Get IOrder_GoodAfterTimeTZ() As String
IOrder_GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Private Property Let IOrder_GoodTillDate(ByVal Value As Date)
GoodTillDate = Value
End Property

Private Property Get IOrder_GoodTillDate() As Date
IOrder_GoodTillDate = mGoodTillDate
End Property

Private Property Let IOrder_GoodTillDateTZ(ByVal Value As String)
GoodTillDateTZ = Value
End Property

Private Property Get IOrder_GoodTillDateTZ() As String
IOrder_GoodTillDateTZ = mGoodTillDateTZ
End Property

Private Property Get IOrder_GroupName() As String
IOrder_GroupName = GroupName
End Property

Private Property Let IOrder_Hidden(ByVal Value As Boolean)
Hidden = Value
End Property

Private Property Get IOrder_Hidden() As Boolean
IOrder_Hidden = mHidden
End Property

Private Property Let IOrder_Id(ByVal Value As String)
Id = Value
End Property

Private Property Get IOrder_Id() As String
IOrder_Id = mIdentifier
End Property

Private Property Let IOrder_IgnoreRegularTradingHours(ByVal Value As Boolean)
IgnoreRegularTradingHours = Value
End Property

Private Property Get IOrder_IgnoreRegularTradingHours() As Boolean
IOrder_IgnoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Private Sub IOrder_Initialise(ByVal pGroupName As String, ByVal pContractSpec As ContractUtils27.IContractSpecifier, ByVal pOrderContext As OrderUtils27.OrderContext)
Initialise pGroupName, pContractSpec, pOrderContext
End Sub

Private Property Get IOrder_IsActive() As Boolean
IOrder_IsActive = IsActive
End Property

Private Function IOrder_IsAttributeModifiable(ByVal pAttribute As OrderAttributes) As Boolean
IOrder_IsAttributeModifiable = True
End Function

Private Property Get IOrder_IsDirty() As Boolean
IOrder_IsDirty = IsDirty
End Property

Private Function IOrder_IsModifiable() As Boolean
IOrder_IsModifiable = True
End Function

Private Property Let IOrder_IsSimulated(ByVal Value As Boolean)
IOrder_IsSimulated = IsSimulated
End Property

Private Property Get IOrder_IsSimulated() As Boolean
IOrder_IsSimulated = False
End Property

Private Property Let IOrder_LastFillPrice(ByVal Value As Double)
LastFillPrice = Value
End Property

Private Property Get IOrder_LastFillPrice() As Double
IOrder_LastFillPrice = LastFillPrice
End Property

Private Property Let IOrder_LimitPrice(ByVal Value As Double)
LimitPrice = Value
End Property

Private Property Get IOrder_LimitPrice() As Double
IOrder_LimitPrice = mLimitPrice
End Property

Private Property Let IOrder_LimitPriceSpec(ByVal Value As PriceSpecifier)
LimitPriceSpec = Value
End Property

Private Property Get IOrder_LimitPriceSpec() As PriceSpecifier
Set IOrder_LimitPriceSpec = LimitPriceSpec
End Property

Private Property Let IOrder_Message(ByVal Value As String)
Message = Value
End Property

Private Property Get IOrder_Message() As String
IOrder_Message = mMessage
End Property

Private Property Let IOrder_MinimumQuantity(ByVal Value As Long)
MinimumQuantity = Value
End Property

Private Property Get IOrder_MinimumQuantity() As Long
IOrder_MinimumQuantity = mMinimumQuantity
End Property

Private Property Let IOrder_NbboPriceCap(ByVal Value As Double)
NbboPriceCap = Value
End Property

Private Property Get IOrder_NbboPriceCap() As Double
IOrder_NbboPriceCap = mNBBOPriceCap
End Property

Private Sub IOrder_NotifyError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
NotifyError pErrorCode, pErrorMsg
End Sub

Private Sub IOrder_NotifyFill(ByVal pExec As IExecutionReport)
NotifyFill pExec
End Sub

Private Sub IOrder_NotifyMessage(ByVal pMessage As String)

End Sub

Private Property Let IOrder_Offset(ByVal Value As Long)
Offset = Value
End Property

Private Property Get IOrder_Offset() As Long
IOrder_Offset = Offset
End Property

Private Property Get IOrder_OrderContext() As OrderContext
Set IOrder_OrderContext = OrderContext
End Property

Private Property Get IOrder_OrderStopTriggerMethod() As OrderStopTriggerMethods
IOrder_OrderStopTriggerMethod = mStopTriggerMethod
End Property

Private Property Let IOrder_OrderType(ByVal Value As OrderTypes)
OrderType = Value
End Property

Private Property Get IOrder_OrderType() As OrderTypes
IOrder_OrderType = mOrderType
End Property

Private Property Let IOrder_Origin(ByVal Value As OrderOrigins)
Origin = Value
End Property

Private Property Get IOrder_Origin() As OrderOrigins
IOrder_Origin = mOrigin
End Property

Private Property Let IOrder_OriginatorRef(ByVal Value As String)
OriginatorRef = Value
End Property

Private Property Get IOrder_OriginatorRef() As String
IOrder_OriginatorRef = moriginatorRef
End Property

Private Property Let IOrder_OverrideConstraints(ByVal Value As Boolean)
OverrideConstraints = Value
End Property

Private Property Get IOrder_OverrideConstraints() As Boolean
IOrder_OverrideConstraints = mOverrideConstraints
End Property

Private Property Let IOrder_PercentOffset(ByVal Value As Double)
PercentOffset = Value
End Property

Private Property Get IOrder_PercentOffset() As Double
IOrder_PercentOffset = PercentOffset
End Property

Private Property Let IOrder_ProviderProperties(ByVal Value As TWUtilities40.Parameters)
ProviderProperties = Value
End Property

Private Property Get IOrder_ProviderProperties() As Parameters
Set IOrder_ProviderProperties = mProviderProperties
End Property

Private Property Let IOrder_Quantity(ByVal Value As Long)
Quantity = Value
End Property

Private Property Get IOrder_Quantity() As Long
IOrder_Quantity = mQuantity
End Property

Private Property Let IOrder_QuantityFilled(ByVal Value As Long)
QuantityFilled = Value
End Property

Private Property Get IOrder_QuantityFilled() As Long
IOrder_QuantityFilled = mQuantityFilled
End Property

Private Property Let IOrder_QuantityRemaining(ByVal Value As Long)
QuantityRemaining = Value
End Property

Private Property Get IOrder_QuantityRemaining() As Long
IOrder_QuantityRemaining = mQuantityRemaining
End Property

Private Property Let IOrder_SettlingFirm(ByVal Value As String)
SettlingFirm = Value
End Property

Private Property Get IOrder_SettlingFirm() As String
IOrder_SettlingFirm = mSettlingFirm
End Property

Private Property Let IOrder_Status(ByVal Value As OrderStatuses)
Status = Value
End Property

Private Property Get IOrder_Status() As OrderStatuses
IOrder_Status = mStatus
End Property

Private Property Let IOrder_StopTriggerMethod(ByVal Value As OrderStopTriggerMethods)
StopTriggerMethod = Value
End Property

Private Property Get IOrder_StopTriggerMethod() As OrderStopTriggerMethods
IOrder_StopTriggerMethod = StopTriggerMethod
End Property

Private Property Let IOrder_SweepToFill(ByVal Value As Boolean)
SweepToFill = Value
End Property

Private Property Get IOrder_SweepToFill() As Boolean
IOrder_SweepToFill = mSweepToFill
End Property

Private Sub IOrder_SyncToOrderReport(ByVal pOrderReport As IOrderReport)
Const ProcName As String = "IOrder_SyncToOrderReport"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Not implemented"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let IOrder_TimeInForce(ByVal Value As OrderTIFs)
TimeInForce = Value
End Property

Private Property Get IOrder_TimeInForce() As OrderTIFs
IOrder_TimeInForce = mTimeInForce
End Property

Private Property Let IOrder_TriggerPrice(ByVal Value As Double)
TriggerPrice = Value
End Property

Private Property Get IOrder_TriggerPrice() As Double
IOrder_TriggerPrice = mTriggerPrice
End Property

Private Property Let IOrder_TriggerPriceSpec(ByVal Value As PriceSpecifier)
TriggerPriceSpec = Value
End Property

Private Property Get IOrder_TriggerPriceSpec() As PriceSpecifier
Set IOrder_TriggerPriceSpec = TriggerPriceSpec
End Property
'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Action(ByVal Value As OrderActions)
mAction = Value
End Property

Public Property Get Action() As OrderActions
Action = mAction
End Property

Public Property Let AllOrNone(ByVal Value As Boolean)
Const ProcName As String = "allOrNone"

On Error GoTo Err

If Value = mAllOrNone Then Exit Property
setDirty
mAllOrNone = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get AllOrNone() As Boolean
AllOrNone = mAllOrNone
End Property

Friend Property Let AveragePrice(ByVal Value As Double)
mAveragePrice = Value
End Property

Public Property Get AveragePrice() As Double
AveragePrice = mAveragePrice
End Property

Public Property Let BlockOrder(Value As Boolean)
Const ProcName As String = "blockOrder"

On Error GoTo Err

If Value = mBlockOrder Then Exit Property
setDirty
mBlockOrder = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BlockOrder() As Boolean
BlockOrder = mBlockOrder
End Property

Friend Property Let BrokerId(ByVal Value As String)
mBrokerId = Value
End Property

Public Property Get BrokerId() As String
BrokerId = mBrokerId
End Property

Public Property Let ContractSpecifier(ByVal Value As IContractSpecifier)
Set mContractSpec = Value
End Property

Public Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpec
End Property

Public Property Let DiscretionaryAmount(ByVal Value As Double)
Const ProcName As String = "discretionaryAmount"

On Error GoTo Err

If Value = mDiscretionaryAmount Then Exit Property
setDirty
mDiscretionaryAmount = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DiscretionaryAmount() As Double
DiscretionaryAmount = mDiscretionaryAmount
End Property

Public Property Let DisplaySize(ByVal Value As Long)
Const ProcName As String = "displaySize"

On Error GoTo Err

If Value = mDisplaySize Then Exit Property
setDirty
mDisplaySize = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DisplaySize() As Long
DisplaySize = mDisplaySize
End Property

Public Property Let ErrorCode(ByVal Value As Long)
mErrorCode = Value
End Property

Public Property Get ErrorCode() As Long
ErrorCode = mErrorCode
End Property

Public Property Let ErrorMessage(ByVal Value As String)
mErrorMessage = Value
End Property

Public Property Get ErrorMessage() As String
ErrorMessage = mErrorMessage
End Property

Public Property Let ETradeOnly(ByVal Value As Boolean)
Const ProcName As String = "eTradeOnly"

On Error GoTo Err

If Value = mETradeOnly Then Exit Property
setDirty
mETradeOnly = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ETradeOnly() As Boolean
ETradeOnly = mETradeOnly
End Property

Friend Property Let FillTime(ByVal Value As Date)
mFillTime = Value
End Property

Public Property Get FillTime() As Date
FillTime = mFillTime
End Property

Public Property Let FirmQuoteOnly(ByVal Value As Boolean)
Const ProcName As String = "firmQuoteOnly"
On Error GoTo Err

If Value = mFirmQuoteOnly Then Exit Property
setDirty
mFirmQuoteOnly = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FirmQuoteOnly() As Boolean
FirmQuoteOnly = mFirmQuoteOnly
End Property

Public Property Let GoodAfterTime(ByVal Value As Date)
Const ProcName As String = "goodAfterTime"
On Error GoTo Err

If Value = mGoodAfterTime Then Exit Property
setDirty
mGoodAfterTime = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get GoodAfterTime() As Date
GoodAfterTime = mGoodAfterTime
End Property

Public Property Let GoodAfterTimeTZ(ByVal Value As String)
Const ProcName As String = "goodAfterTimeTZ"

On Error GoTo Err

If Value = mGoodAfterTimeTZ Then Exit Property
setDirty
mGoodAfterTimeTZ = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get GoodAfterTimeTZ() As String
GoodAfterTimeTZ = mGoodAfterTimeTZ
End Property

Public Property Let GoodTillDate(ByVal Value As Date)
Const ProcName As String = "goodTillDate"

On Error GoTo Err

If Value = mGoodTillDate Then Exit Property
setDirty
mGoodTillDate = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get GoodTillDate() As Date
GoodTillDate = mGoodTillDate
End Property

Public Property Let GoodTillDateTZ(ByVal Value As String)
Const ProcName As String = "goodTillDateTZ"

On Error GoTo Err

If Value = mGoodTillDateTZ Then Exit Property
setDirty
mGoodTillDateTZ = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get GoodTillDateTZ() As String
GoodTillDateTZ = mGoodTillDateTZ
End Property

Public Property Get GroupName() As String
GroupName = mGroupName
End Property

Public Property Let Hidden(Value As Boolean)
Const ProcName As String = "hidden"

On Error GoTo Err

If Value = mHidden Then Exit Property
setDirty
mHidden = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Friend Property Let Id(ByVal Value As String)
mIdentifier = Value
End Property

Public Property Get Id() As String
Id = mIdentifier
End Property

Public Property Let IgnoreRegularTradingHours(Value As Boolean)
Const ProcName As String = "ignoreRegularTradingHours"

On Error GoTo Err

If Value = mIgnoreRegularTradingHours Then Exit Property
setDirty
mIgnoreRegularTradingHours = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IgnoreRegularTradingHours() As Boolean
IgnoreRegularTradingHours = mIgnoreRegularTradingHours
End Property

Friend Property Get IsActive() As Boolean
IsActive = (mStatus = OrderStatusPendingSubmit) Or _
        (mStatus = OrderStatusPreSubmitted) Or _
        (mStatus = OrderStatusSubmitted) Or _
        (mStatus = OrderStatusCancelling)
End Property

Public Property Get IsDirty() As Boolean
IsDirty = mIsDirty
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Let LastFillPrice(ByVal Value As Double)
mLastFillPrice = Value
End Property

Public Property Get LastFillPrice() As Double
LastFillPrice = mLastFillPrice
End Property

Friend Property Let LimitPrice(ByVal Value As Double)
mLimitPrice = Value
End Property

Public Property Get LimitPrice() As Double
LimitPrice = mLimitPrice
End Property

Private Property Let LimitPriceSpec(ByVal Value As PriceSpecifier)
Set mLimitPriceSpec = Value
End Property

Private Property Get LimitPriceSpec() As PriceSpecifier
Set LimitPriceSpec = mLimitPriceSpec
End Property

Public Property Let Message(ByVal Value As String)
mMessage = Value
End Property

Public Property Get Message() As String
Message = mMessage
End Property

Public Property Let MinimumQuantity(ByVal Value As Long)
Const ProcName As String = "minimumQuantity"

On Error GoTo Err

If Value = mMinimumQuantity Then Exit Property
setDirty
mMinimumQuantity = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MinimumQuantity() As Long
MinimumQuantity = mMinimumQuantity
End Property

Public Property Let NbboPriceCap(ByVal Value As Double)
Const ProcName As String = "NbboPriceCap"

On Error GoTo Err

If Value = mNBBOPriceCap Then Exit Property
setDirty
mNBBOPriceCap = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get NbboPriceCap() As Double
NbboPriceCap = mNBBOPriceCap
End Property

Friend Property Let Offset(ByVal Value As Long)
mOffset = Value
End Property

Friend Property Get Offset() As Long
Offset = mOffset
End Property

Public Property Get OrderContext() As OrderContext
Set OrderContext = mOrderContext
End Property

Friend Property Let OrderType(ByVal Value As OrderTypes)
Const ProcName As String = "orderType"

On Error GoTo Err

If Value = mOrderType Then Exit Property
setDirty
mOrderType = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OrderType() As OrderTypes
OrderType = mOrderType
End Property

Public Property Let Origin(ByVal Value As OrderOrigins)
Const ProcName As String = "origin"

On Error GoTo Err

If Value = mOrigin Then Exit Property
setDirty
mOrigin = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Origin() As OrderOrigins
Origin = mOrigin
End Property

Public Property Let OriginatorRef(ByVal Value As String)
Const ProcName As String = "originatorRef"

On Error GoTo Err

If Value = moriginatorRef Then Exit Property
setDirty
moriginatorRef = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OriginatorRef() As String
OriginatorRef = moriginatorRef
End Property

Public Property Let OverrideConstraints(ByVal Value As Boolean)
Const ProcName As String = "overrideConstraints"

On Error GoTo Err

If Value = mOverrideConstraints Then Exit Property
setDirty
mOverrideConstraints = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OverrideConstraints() As Boolean
OverrideConstraints = mOverrideConstraints
End Property

Public Property Let PercentOffset(ByVal Value As Double)
Const ProcName As String = "percentOffset"

On Error GoTo Err

If Value = mPercentOffset Then Exit Property
setDirty
mPercentOffset = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PercentOffset() As Double
PercentOffset = mPercentOffset
End Property

Friend Property Let ProviderProperties( _
                ByVal Value As Parameters)
Set mProviderProperties = Value
End Property

Public Property Get ProviderProperties() As Parameters
Set ProviderProperties = mProviderProperties
End Property

Public Property Let Quantity(ByVal Value As Long)
Const ProcName As String = "quantity"

On Error GoTo Err

If Value = mQuantity Then Exit Property
mQuantity = Value
setDirty
mQuantityRemaining = mQuantity - mQuantityFilled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Quantity() As Long
Quantity = mQuantity
End Property

Friend Property Let QuantityFilled(ByVal Value As Long)
mQuantityFilled = Value
End Property

Public Property Get QuantityFilled() As Long
QuantityFilled = mQuantityFilled
End Property

Friend Property Let QuantityRemaining(ByVal Value As Long)
mQuantityRemaining = Value
End Property

Public Property Get QuantityRemaining() As Long
QuantityRemaining = mQuantityRemaining
End Property

Public Property Let SettlingFirm(Value As String)
Const ProcName As String = "settlingFirm"

On Error GoTo Err

If Value = mSettlingFirm Then Exit Property
setDirty
mSettlingFirm = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SettlingFirm() As String
SettlingFirm = mSettlingFirm
End Property

Friend Property Let Status(Value As OrderStatuses)
mStatus = Value
End Property

Public Property Get Status() As OrderStatuses
Status = mStatus
End Property

Public Property Let StopTriggerMethod(ByVal Value As OrderStopTriggerMethods)
Const ProcName As String = "StopTriggerMethod"
On Error GoTo Err

If Value = mStopTriggerMethod Then Exit Property
setDirty
mStopTriggerMethod = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StopTriggerMethod() As OrderStopTriggerMethods
StopTriggerMethod = mStopTriggerMethod
End Property

Public Property Let SweepToFill(Value As Boolean)
Const ProcName As String = "SweepToFill"

On Error GoTo Err

If Value = mSweepToFill Then Exit Property
setDirty
mSweepToFill = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SweepToFill() As Boolean
SweepToFill = mSweepToFill
End Property

Public Property Let TimeInForce(ByVal Value As OrderTIFs)
Const ProcName As String = "timeInForce"

On Error GoTo Err

If Value = mTimeInForce Then Exit Property
setDirty
mTimeInForce = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimeInForce() As OrderTIFs
TimeInForce = mTimeInForce
End Property

Friend Property Let TriggerPrice(ByVal Value As Double)
mTriggerPrice = Value
End Property

Public Property Get TriggerPrice() As Double
TriggerPrice = mTriggerPrice
End Property

Friend Property Let TriggerPriceSpec(ByVal Value As PriceSpecifier)
Set mTriggerPriceSpec = Value
End Property

Public Property Get TriggerPriceSpec() As PriceSpecifier
Set TriggerPriceSpec = mTriggerPriceSpec
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub NotifyError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
mErrorCode = pErrorCode
mErrorMessage = pErrorMsg
End Sub

''
' Notifies a fill for this <code>order</code>.
'
' @param exec  an <code>Execution</code> object containing details of the fill
'
'@/
Friend Sub NotifyFill( _
                ByVal exec As IExecutionReport)
Const ProcName As String = "NotifyFill"
On Error GoTo Err

mAveragePrice = (mQuantityFilled * mAveragePrice + exec.Quantity * exec.Price) / (mQuantityFilled + exec.Quantity)
mQuantityFilled = mQuantityFilled + exec.Quantity
mQuantityRemaining = mQuantityRemaining - exec.Quantity
mLastFillPrice = exec.Price
mFillTime = exec.FillTime
If mQuantityRemaining = 0 Then
    mStatus = OrderStatusFilled
End If
RaiseEvent Fill(exec)
If mStatus = OrderStatusFilled Then RaiseEvent StatusChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"

On Error GoTo Err

Set mContractSpec = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise(ByVal pGroupName As String, ByVal pContractSpec As IContractSpecifier, ByVal pOrderContext As OrderContext)
mGroupName = pGroupName
Set mContractSpec = pContractSpec
Set mOrderContext = pOrderContext
End Sub

Public Function IsModifiable() As Boolean
Const ProcName As String = "isModifiable"

On Error GoTo Err

If mStatus <> OrderStatuses.OrderStatusCancelled And _
    mStatus <> OrderStatuses.OrderStatusCancelling And _
    mStatus <> OrderStatuses.OrderStatusFilled _
Then IsModifiable = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Marks this order as clean, ie all changes have been either
' actioned or cancelled.
'@/
Friend Sub setClean()
Const ProcName As String = "setClean"

On Error GoTo Err

mIsDirty = False
RaiseEvent Clean

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Notifies the specified <code>order</code> of a change of status
'
' @param pStatus    a member of the <code>OrderStatuses</code> enumeration
'                   specifying the <code>order</code>'s new status
'
'@/
Friend Sub UpdateStatus( _
                ByVal pStatus As OrderStatuses, _
                ByRef persistenceKey As String)

Const ProcName As String = "UpdateStatus"

On Error GoTo Err

If pStatus <> mStatus Then
    mStatus = pStatus
    RaiseEvent StatusChanged
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function propertiesEqual( _
                ByVal pParams1 As Parameters, _
                ByVal pParams2 As Parameters) As Boolean
If pParams1 Is Nothing And pParams2 Is Nothing Then
    propertiesEqual = True
ElseIf pParams1 Is Nothing Or pParams2 Is Nothing Then
    propertiesEqual = False
Else
    propertiesEqual = pParams1.Equals(pParams2)
End If
End Function

''
' Set the 'dirty' flag, and raise a Dirty event if one has not already
' been raised.
'@/
Private Sub setDirty()
Const ProcName As String = "setDirty"

On Error GoTo Err

If Not mIsDirty Then
    mIsDirty = True
    RaiseEvent Dirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub







